VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Begin VB.Form frmOutAdviceEdit 
   AutoRedraw      =   -1  'True
   Caption         =   "����ҽ���༭"
   ClientHeight    =   8100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10875
   Icon            =   "frmOutAdviceEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8100
   ScaleWidth      =   10875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   40
      Top             =   7740
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   10
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmOutAdviceEdit.frx":058A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11060
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   2
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   2
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   2
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmOutAdviceEdit.frx":0E1E
            Key             =   "PY"
            Object.ToolTipText     =   "ƴ��(F7)"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmOutAdviceEdit.frx":1458
            Key             =   "WB"
            Object.ToolTipText     =   "���(F7)"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   953
            MinWidth        =   25
            Text            =   "�Ƽ�"
            TextSave        =   "�Ƽ�"
            Key             =   "Price"
            Object.ToolTipText     =   "��ʾ���ƼƼ����(F8)"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   953
            MinWidth        =   26
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   26
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
   Begin VB.Frame fraPati 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   0
      TabIndex        =   25
      Top             =   510
      Width           =   10875
      Begin VB.CommandButton cmdAlley 
         Caption         =   "����ʷ/����״̬"
         Height          =   350
         Left            =   9120
         TabIndex        =   22
         Top             =   50
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.ComboBox cboӤ�� 
         Height          =   300
         ItemData        =   "frmOutAdviceEdit.frx":1A92
         Left            =   9435
         List            =   "frmOutAdviceEdit.frx":1AA8
         Style           =   2  'Dropdown List
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   75
         Width           =   1395
      End
      Begin VB.Label lblӤ�� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ӥ��(&B)"
         Height          =   180
         Left            =   8745
         TabIndex        =   20
         Top             =   135
         Width           =   630
      End
      Begin VB.Label lblPati 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����: �Ա�: ����: �����: �ѱ�: ҽ�Ƹ��ʽ:"
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   210
         TabIndex        =   26
         Top             =   135
         Width           =   4050
      End
   End
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   510
      Left            =   0
      TabIndex        =   23
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
         TabIndex        =   24
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
            NumButtons      =   17
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Description     =   "����"
               Object.ToolTipText     =   "����һ��ҽ��(Ctrl+A)"
               Object.Tag             =   "����"
               ImageKey        =   "����"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Description     =   "����"
               Object.ToolTipText     =   "����һ��ҽ��(Ctrl+I)"
               Object.Tag             =   "����"
               ImageKey        =   "����"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ��"
               Key             =   "ɾ��"
               Description     =   "ɾ��"
               Object.ToolTipText     =   "ɾ����ǰҽ��(Del)"
               Object.Tag             =   "ɾ��"
               ImageKey        =   "ɾ��"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "һ��"
               Key             =   "һ��"
               Description     =   "һ��"
               Object.ToolTipText     =   "һ����ҩ(Ctrl+K)"
               Object.Tag             =   "һ��"
               ImageKey        =   "һ��"
               Style           =   1
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�ο�"
               Key             =   "�ο�"
               Description     =   "�ο�"
               Object.ToolTipText     =   "�鿴������Ŀ�ο�(F6)"
               Object.Tag             =   "�ο�"
               ImageKey        =   "�ο�"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "�ο�_"
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Description     =   "����"
               Object.ToolTipText     =   "���Ʋ����µ�ҽ��(Ctrl+Y)"
               Object.Tag             =   "����"
               ImageKey        =   "����"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Description     =   "����"
               Object.ToolTipText     =   "����Ϊ����ҽ��(Ctrl+T)"
               Object.Tag             =   "����"
               ImageKey        =   "����"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Description     =   "����"
               Object.ToolTipText     =   "����ҽ��(F2��Ctrl+S)"
               Object.Tag             =   "����"
               ImageKey        =   "����"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ǩ��"
               Key             =   "ǩ��"
               Description     =   "ǩ��"
               Object.ToolTipText     =   "ҽ��ǩ��"
               Object.Tag             =   "ǩ��"
               ImageKey        =   "ǩ��"
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "����_"
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Description     =   "����"
               Object.ToolTipText     =   "����ҽ��(F3)"
               Object.Tag             =   "����"
               ImageKey        =   "����"
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "����_"
               Style           =   3
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Description     =   "����"
               Object.ToolTipText     =   "����(F1)"
               Object.Tag             =   "����"
               ImageKey        =   "����"
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
   Begin MSComCtl2.MonthView dtpDate 
      Height          =   2220
      Left            =   1725
      TabIndex        =   1
      Top             =   2505
      Visible         =   0   'False
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   3916
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   66781185
      TitleBackColor  =   -2147483636
      TitleForeColor  =   -2147483634
      TrailingForeColor=   -2147483637
      CurrentDate     =   37904
   End
   Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
      Height          =   4800
      Left            =   60
      TabIndex        =   0
      Top             =   960
      Width           =   10770
      _cx             =   18997
      _cy             =   8467
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
      ForeColorSel    =   0
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
      Rows            =   18
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmOutAdviceEdit.frx":1AF7
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
      Begin MSComctlLib.ImageList imgFlag 
         Left            =   405
         Top             =   450
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   8
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOutAdviceEdit.frx":1BDF
               Key             =   "����"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imgPass 
         Left            =   1155
         Top             =   450
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
               Picture         =   "frmOutAdviceEdit.frx":1DF9
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOutAdviceEdit.frx":20F3
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOutAdviceEdit.frx":23ED
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOutAdviceEdit.frx":26E7
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOutAdviceEdit.frx":29E1
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imgSign 
         Left            =   1965
         Top             =   450
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmOutAdviceEdit.frx":2CDB
               Key             =   "ǩ��"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fraAdvice 
      Height          =   2040
      Left            =   45
      TabIndex        =   27
      Top             =   5700
      Width           =   10800
      Begin VB.ComboBox cbo����ִ�� 
         Height          =   300
         Left            =   6255
         TabIndex        =   19
         Text            =   "cbo����ִ��"
         Top             =   1635
         Width           =   1860
      End
      Begin VB.TextBox txt���� 
         Alignment       =   2  'Center
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   2385
         MaxLength       =   3
         TabIndex        =   13
         Top             =   1635
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.CommandButton cmdƵ�� 
         Height          =   240
         Left            =   4860
         Picture         =   "frmOutAdviceEdit.frx":302D
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "ѡ����Ŀ(F4)"
         Top             =   1305
         Width           =   270
      End
      Begin VB.TextBox txtƵ�� 
         Height          =   300
         Left            =   3495
         TabIndex        =   10
         Top             =   1275
         Width           =   1665
      End
      Begin VB.TextBox txt���� 
         Alignment       =   1  'Right Justify
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3495
         MaxLength       =   10
         TabIndex        =   14
         Top             =   1635
         Width           =   1380
      End
      Begin VB.TextBox txt���� 
         Alignment       =   1  'Right Justify
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   930
         MaxLength       =   10
         TabIndex        =   12
         Top             =   1635
         Width           =   1530
      End
      Begin VB.CommandButton cmd�÷� 
         Height          =   240
         Left            =   2445
         Picture         =   "frmOutAdviceEdit.frx":3123
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "ѡ����Ŀ(F4)"
         Top             =   1305
         Width           =   270
      End
      Begin VB.TextBox txt�÷� 
         Height          =   300
         Left            =   930
         TabIndex        =   8
         Top             =   1275
         Width           =   1815
      End
      Begin VB.CommandButton cmd��ʼʱ�� 
         Height          =   240
         Left            =   2460
         Picture         =   "frmOutAdviceEdit.frx":3219
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "ѡ������(F4)"
         Top             =   225
         Width           =   255
      End
      Begin VB.CheckBox chk���� 
         Alignment       =   1  'Right Justify
         Caption         =   "����ҽ��(&E)"
         Height          =   225
         Left            =   3570
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   233
         Width           =   1290
      End
      Begin VB.CommandButton cmdExt 
         Height          =   285
         Left            =   4890
         Picture         =   "frmOutAdviceEdit.frx":330F
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "�༭(F4)"
         Top             =   600
         Width           =   285
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   "��"
         Height          =   285
         Left            =   4890
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "ѡ����Ŀ(*)"
         Top             =   900
         Width           =   285
      End
      Begin VB.ComboBox cboִ������ 
         Height          =   300
         ItemData        =   "frmOutAdviceEdit.frx":3405
         Left            =   9015
         List            =   "frmOutAdviceEdit.frx":3412
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1275
         Width           =   1590
      End
      Begin VB.ComboBox cboִ�п��� 
         Height          =   300
         Left            =   6255
         TabIndex        =   17
         Text            =   "cboִ�п���"
         Top             =   1275
         Width           =   1860
      End
      Begin VB.TextBox txtҽ������ 
         Height          =   660
         Left            =   930
         MaxLength       =   1000
         MultiLine       =   -1  'True
         TabIndex        =   5
         ToolTipText     =   "�� ~ ���л���ݸ������"
         Top             =   552
         Width           =   3945
      End
      Begin VB.TextBox txt��ʼʱ�� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   930
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   195
         Width           =   1815
      End
      Begin VB.ComboBox cboִ��ʱ�� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   6255
         TabIndex        =   16
         Top             =   915
         Width           =   4350
      End
      Begin VB.ComboBox cboҽ������ 
         Height          =   300
         Left            =   6255
         TabIndex        =   15
         Top             =   555
         Width           =   4350
      End
      Begin VB.Label lbl����ִ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ִ��"
         Height          =   180
         Left            =   5490
         TabIndex        =   42
         Top             =   1695
         Width           =   720
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��    ��"
         Height          =   180
         Left            =   2205
         TabIndex        =   41
         Top             =   1695
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lblƵ�� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ƶ��"
         Height          =   180
         Left            =   3105
         TabIndex        =   35
         Top             =   1335
         Width           =   360
      End
      Begin VB.Label lbl������λ 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "��λ"
         Height          =   180
         Left            =   4905
         TabIndex        =   31
         Top             =   1695
         Width           =   570
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   3105
         TabIndex        =   30
         Top             =   1695
         Width           =   360
      End
      Begin VB.Label lbl������λ 
         BackStyle       =   0  'Transparent
         Caption         =   "��λ"
         Height          =   180
         Left            =   2505
         TabIndex        =   33
         Top             =   1695
         Width           =   570
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   540
         TabIndex        =   32
         Top             =   1695
         Width           =   360
      End
      Begin VB.Label lblҽ������ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ������"
         Height          =   180
         Left            =   5490
         TabIndex        =   39
         Top             =   615
         Width           =   720
      End
      Begin VB.Label lblִ������ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ִ������"
         Height          =   180
         Left            =   8250
         TabIndex        =   38
         Top             =   1335
         Width           =   720
      End
      Begin VB.Label lblִ�п��� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ִ�п���"
         Height          =   180
         Left            =   5490
         TabIndex        =   37
         Top             =   1335
         Width           =   720
      End
      Begin VB.Label lbl�÷� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�÷�"
         Height          =   180
         Left            =   540
         TabIndex        =   34
         Top             =   1335
         Width           =   360
      End
      Begin VB.Label lblҽ������ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ������"
         Height          =   180
         Left            =   180
         TabIndex        =   29
         Top             =   600
         Width           =   720
      End
      Begin VB.Label lbl��ʼʱ�� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ʼʱ��"
         Height          =   180
         Left            =   180
         TabIndex        =   28
         Top             =   255
         Width           =   720
      End
      Begin VB.Label lblִ��ʱ�� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ִ��ʱ��"
         Height          =   180
         Left            =   5490
         TabIndex        =   36
         Top             =   975
         Width           =   720
      End
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
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutAdviceEdit.frx":3434
            Key             =   "����"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutAdviceEdit.frx":364E
            Key             =   "����"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutAdviceEdit.frx":3868
            Key             =   "ɾ��"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutAdviceEdit.frx":3A82
            Key             =   "һ��"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutAdviceEdit.frx":3C9C
            Key             =   "�ο�"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutAdviceEdit.frx":3EB6
            Key             =   "����"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutAdviceEdit.frx":40D0
            Key             =   "����"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutAdviceEdit.frx":42EA
            Key             =   "����"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutAdviceEdit.frx":49E4
            Key             =   "����"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutAdviceEdit.frx":4BFE
            Key             =   "�˳�"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutAdviceEdit.frx":4E18
            Key             =   "����"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutAdviceEdit.frx":5512
            Key             =   "ǩ��"
         EndProperty
      EndProperty
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
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutAdviceEdit.frx":5C0C
            Key             =   "����"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutAdviceEdit.frx":5E26
            Key             =   "����"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutAdviceEdit.frx":6040
            Key             =   "ɾ��"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutAdviceEdit.frx":625A
            Key             =   "һ��"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutAdviceEdit.frx":6474
            Key             =   "�ο�"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutAdviceEdit.frx":668E
            Key             =   "����"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutAdviceEdit.frx":68A8
            Key             =   "����"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutAdviceEdit.frx":6AC2
            Key             =   "����"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutAdviceEdit.frx":71BC
            Key             =   "����"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutAdviceEdit.frx":73D6
            Key             =   "�˳�"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutAdviceEdit.frx":75F0
            Key             =   "����"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOutAdviceEdit.frx":7CEA
            Key             =   "ǩ��"
         EndProperty
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
      Begin VB.Menu mnuPassItem 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "ϵͳ����(&U)"
         Index           =   14
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "-"
         Index           =   15
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "��ҩ�о�(&M)"
         Index           =   16
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "-"
         Index           =   17
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "����(&W)"
         Index           =   18
      End
      Begin VB.Menu mnuPassItem 
         Caption         =   "���(&V)"
         Index           =   19
      End
   End
End
Attribute VB_Name = "frmOutAdviceEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mblnOK As Boolean
'��ڲ���
Private mblnModal As Boolean
Private mfrmParent As Object
Private mstrPrivs As String
Private mlng����ID As Long
Private mstr�Һŵ� As String '���˹Һŵ��ݺ�
Private mlngǰ��ID As Long 'ҽ������վ��ҽ��ʱ��
Private mintӤ�� As Integer '�޸�ʱ��
Private mlngҽ��ID As Long '�޸�ʱ��

'�������
Private mobjVBA As Object
Private mobjScript As clsScript
Private mrsDefine As ADODB.Recordset

Private WithEvents mfrmShortCut As frmClinicShortCut
Attribute mfrmShortCut.VB_VarHelpID = -1
Private WithEvents mfrmPrice As frmAdvicePrice
Attribute mfrmPrice.VB_VarHelpID = -1
Private mcolStock As Collection '��Ÿ���ҩƷ�ⷿ�ĳ����鷽ʽ
Private mstrDelIDs As String '��¼��Ҫ��ɾ����ҽ��ID
Private mstr�Ա� As String '������Ŀ���������ж�
Private mint���� As Integer '���˵���������
Private mdat�Һ�ʱ�� As Date '��������ж�
Private mlng���˿���id As Long '����(�Һ�)����ID
Private mint���� As Integer '��ǰ��������
Private mstr������ As String '��ǰ����ҽ�Ƹ��ʽ����
Private mlngPassPati As Long 'Pass:�ϴ��Ѵ���PASS�Ĳ���ID

'���ز���
Private mint���� As Integer
Private mstrLike As String
Private mbln�Զ�Ƥ�� As Boolean
Private mbln���� As Boolean
Private msng���� As Single

'�¼�״̬���Ʊ���
Private mblnRunFirst As Boolean
Private mblnRowChange As Boolean
Private mblnDoCheck As Boolean

Private Const TIME_LIMIT = 30 'ҽ���������ڵ�ʱ��
'ִ��ʱ��ʾ��
Private Const COL_����ִ�� = _
    "ÿ������ 1/8-3/8-5/8 �� 1/8:00-3/8:00-5/8:00" & vbCrLf & _
        vbTab & "��ʾ��ÿ������һ��8:00,��������8:00,�������8:00�⼸��ʱ��ִ��"
Private Const COL_����ִ�� = _
    "ÿ������ 8-12-16 �� 8:00-12:00-16:00" & vbCrLf & _
        vbTab & "��ʾ��ÿ��8:00,12:00,16:00�⼸��ʱ��ִ��" & vbCrLf & _
    "����һ�� 1/8 �� 1/8:00" & vbCrLf & _
        vbTab & "��ʾ��ÿ�����еĵ�1��8:00���ʱ��ִ��"
Private Const COL_��ʱִ�� = _
    "ÿСʱ���� 1:20-1:40" & vbCrLf & _
        vbTab & "��ʾ��ÿСʱ�ڵ�20��40����������ʱ��ִ��" & vbCrLf & _
    "��Сʱһ�� 2:30 �� 1:30 �� 1:00" & vbCrLf & _
        vbTab & "��ʾ��ÿ��Сʱ�ڵĵ�2�ĸ�Сʱ��30�������ʱ��ִ��" & vbCrLf & _
        vbTab & "������ÿ��Сʱ�ڵĵ�1�ĸ�Сʱ��30�������ʱ��ִ��" & vbCrLf & _
        vbTab & "������ÿ��Сʱ�ڵĵ�1�ĸ�Сʱ���ʱ��ִ��"

'�̶���
Private Const COL_F��־ = 0
'�ɼ�������
Private Const COL_��ʾ = 1 'Pass:���ַ������ʹ���,�ձ�ʾû�������
Private Const COL_��ʼʱ�� = 2
Private Const COL_ҽ������ = 3
Private Const COL_���� = 4
Private Const COL_������λ = 5
Private Const COL_���� = 6
Private Const COL_������λ = 7
Private Const COL_Ƶ�� = 8
Private Const COL_�÷� = 9
Private Const COL_ҽ������ = 10
Private Const COL_ִ��ʱ�� = 11
Private Const COL_����ҽ�� = 12

'����������
Private Const COL_EDIT = 13 '�༭��־��0-ԭʼ��,1-������,2-�޸�������,3-�޸������,����Dataֵ=���µĳ��׷���ID
Private Const COL_���ID = 14
Private Const COL_Ӥ�� = 15
Private Const COL_��� = 16 'Pass:Dataֵ���ڼ�¼�Ƿ��������˽��
Private Const COL_״̬ = 17
Private Const COL_��� = 18
Private Const COL_������ĿID = 19
Private Const COL_���� = 20
Private Const COL_�걾��λ = 21
Private Const COL_�շ�ϸĿID = 22
Private Const COL_���� = 23
Private Const COL_Ƶ�ʴ��� = 24
Private Const COL_Ƶ�ʼ�� = 25
Private Const COL_�����λ = 26
Private Const COL_�Ƽ����� = 27
Private Const COL_ִ�п���ID = 28
Private Const COL_ִ������ = 29 '����ҽ����¼.ִ������=������ĿĿ¼.ִ�п���
Private Const COL_��������ID = 30
Private Const COL_����ʱ�� = 31
Private Const COL_��־ = 32

Private Const COL_���㷽ʽ = 33 '������ĿĿ¼.���㷽ʽ
Private Const COL_Ƶ������ = 34 '������ĿĿ¼.ִ��Ƶ��
Private Const COL_�������� = 35 '������ĿĿ¼.��������
Private Const COL_��� = 36 '�������װ��ŵĿ��ÿ��
Private Const COL_�ɷ���� = 37
Private Const COL_����ϵ�� = 38
Private Const COL_���ﵥλ = 39
Private Const COL_�����װ = 40
Private Const COL_�������� = 41 '��ҩ������ĿΪ¼������
Private Const COL_����ְ�� = 42
Private Const COL_������� = 43
Private Const COL_ҩƷ���� = 44
Private Const COL_���� = 45
Private Const COL_ǩ���� = 46

Public Function ShowMe(ByVal frmParent As Object, ByVal strPrivs As String, ByVal lng����ID As Long, ByVal str�Һŵ� As String, _
    Optional ByVal lngǰ��ID As Long, Optional ByVal intӤ�� As Integer, Optional ByVal lngҽ��ID As Long, Optional ByVal blnModal As Boolean) As Boolean
    
    Set mfrmParent = frmParent
    mblnModal = blnModal
    mstrPrivs = strPrivs
    mlng����ID = lng����ID
    mstr�Һŵ� = str�Һŵ�
    mlngǰ��ID = lngǰ��ID
    mintӤ�� = intӤ��
    mlngҽ��ID = lngҽ��ID
    
    Me.Show IIF(blnModal, 1, 0), frmParent
    ShowMe = mblnOK
End Function

Private Property Let mblnNoSave(ByVal vData As Boolean)
    tbr.Buttons("����").Enabled = vData
End Property

Private Property Get mblnNoSave() As Boolean
    mblnNoSave = tbr.Buttons("����").Enabled
End Property

Private Sub InitAdviceTable()
'���ܣ���ʼ��������ݣ����ڴ�����Ի����ûָ�֮ǰ
    Dim strHead As String, i As Integer
    Dim arrHead As Variant, arrCol As Variant

    strHead = _
        ",240,4;��ʼʱ��,1080,1;ҽ������,3500,1;����,600,7;��λ,450,1;����,600,7;��λ,450,1;" & _
        "Ƶ��,1200,1;�÷�,1200,1;ҽ������,1000,1;ִ��ʱ��;����ҽ��,850,1;" & _
        "EDIT;���ID;Ӥ��;���;ҽ��״̬;�������;������ĿID;����;�걾��λ;�շ�ϸĿID;" & _
        "����;Ƶ�ʴ���;Ƶ�ʼ��;�����λ;�Ƽ�����;ִ�п���ID;ִ������;��������ID;����ʱ��;��־;" & _
        "���㷽ʽ;Ƶ������;��������;���;�ɷ����;����ϵ��;���ﵥλ;�����װ;��������;����ְ��;�������;ҩƷ����;����;ǩ����"
        
    arrHead = Split(strHead, ";")
    With vsAdvice
        .Clear
        .FixedRows = 1: .FixedCols = 1
        .Rows = 2: .Cols = .FixedCols + UBound(arrHead) + 1
        
        For i = 0 To UBound(arrHead)
            .FixedAlignment(.FixedCols + i) = 4
            arrCol = Split(arrHead(i), ",")
            .TextMatrix(0, .FixedCols + i) = arrCol(0)
            If UBound(arrCol) > 0 Then
                .ColWidth(.FixedCols + i) = Val(arrCol(1))
                .ColAlignment(.FixedCols + i) = Val(arrCol(2))
                .ColHidden(.FixedCols + i) = False
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
        .ColHidden(COL_��ʾ) = True 'Pass
        '.FrozenCols = COL_ҽ������ + 1 - .FixedCols
        .ColWidth(0) = 14 * Screen.TwipsPerPixelX
    End With
End Sub

Private Sub Set�÷�Input(rsInput As ADODB.Recordset, ByVal int���� As Integer)
'���ܣ������ҩ;������ҩ�÷������
'������rsInput=�����ѡ��ķ��ؼ�¼
'      int����=2-��ҩ;��,4-��ҩ�÷�
'˵���������ѡƵ��,����ϸ�ҩ;���������ִ��ʱ�䷽���ı仯
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim blnValid As Boolean, sng���� As Single
    Dim strƵ�� As String, intƵ�ʴ��� As Integer, intƵ�ʼ�� As Integer, str�����λ As String
    
    cmd�÷�.Tag = rsInput!ID
    txt�÷�.Text = rsInput!����
    txt�÷�.Tag = "1"
    
    With vsAdvice
        '���»�ȡ���õ�ȱʡʱ�䷽��
        If cboִ��ʱ��.Enabled Then '"��ѡƵ��"��ҩƷʱ
            Call Getʱ�䷽��(cboִ��ʱ��, GetƵ�ʷ�Χ(.Row), .TextMatrix(.Row, COL_Ƶ��), rsInput!ID)
            If cboִ��ʱ��.ListCount > 0 Then
                cboִ��ʱ��.ListIndex = 0
                cboִ��ʱ��.Tag = "1"
            Else
                '�жϵ�ǰִ��ʱ���Ƿ�Ϸ�
                If cboִ��ʱ��.Text <> "" Then
                    blnValid = ExeTimeValid(cboִ��ʱ��.Text, Val(.TextMatrix(.Row, COL_Ƶ�ʴ���)), Val(.TextMatrix(.Row, COL_Ƶ�ʼ��)), .TextMatrix(.Row, COL_�����λ))
                    If Not blnValid Then '������Ϸ�,����ȡ,���򱣳�
                        cboִ��ʱ��.Text = ""
                        cboִ��ʱ��.Tag = "1"
                    End If
                End If
            End If
        End If
        
        '���������÷�������ȱʡ����
        If InStr(",5,6,", .TextMatrix(.Row, COL_���)) > 0 Then
            strSQL = "Select Ƶ��,С������,���˼���,ҽ������,�Ƴ�" & _
                " From �����÷����� Where ����>0 And ��ĿID=[1] And �÷�ID=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(.Row, COL_������ĿID)), Val(rsInput!ID))
            If Not rsTmp.EOF Then
                If Not IsNull(rsTmp!Ƶ��) Then
                    Call GetƵ����Ϣ_����(rsTmp!Ƶ��, strƵ��, intƵ�ʴ���, intƵ�ʼ��, str�����λ)
                    txtƵ��.Text = strƵ��
                    cmdƵ��.Tag = strƵ��
                    txtƵ��.Tag = "1"
                End If
                
                '�����µ�Ƶ����������ִ��ʱ��
                If cboִ��ʱ��.Enabled Then
                    Call Getʱ�䷽��(cboִ��ʱ��, GetƵ�ʷ�Χ(.Row), strƵ��, rsInput!ID)
                    If cboִ��ʱ��.ListCount > 0 Then
                        cboִ��ʱ��.ListIndex = 0
                        cboִ��ʱ��.Tag = "1"
                    Else
                        '�жϵ�ǰִ��ʱ���Ƿ�Ϸ�
                        If cboִ��ʱ��.Text <> "" Then
                            blnValid = ExeTimeValid(cboִ��ʱ��.Text, intƵ�ʴ���, intƵ�ʼ��, str�����λ)
                            If Not blnValid Then '������Ϸ�,����ȡ,���򱣳�
                                cboִ��ʱ��.Text = ""
                                cboִ��ʱ��.Tag = "1"
                            End If
                        End If
                    End If
                End If

                'ҩƷ����
                If mint���� > 12 Then
                    If Nvl(rsTmp!���˼���, 0) <> 0 Then
                        txt����.Text = FormatEx(rsTmp!���˼���, 5)
                        txt����.Tag = "1"
                    End If
                Else
                    If Nvl(rsTmp!С������, 0) <> 0 Then
                        txt����.Text = FormatEx(rsTmp!С������, 5)
                        txt����.Tag = "1"
                    ElseIf Nvl(rsTmp!���˼���, 0) <> 0 Then
                        txt����.Text = FormatEx(rsTmp!���˼��� * (mint���� + 2) * 5 / 100, 5)
                        txt����.Tag = "1"
                    End If
                End If
                
                'ȡȱʡ������
                sng���� = msng����
                If mbln���� Then
                    If str�����λ = "��" Then
                        sng���� = IIF(7 > sng����, 7, sng����)
                    ElseIf str�����λ = "��" Then
                        sng���� = IIF(intƵ�ʼ�� > sng����, intƵ�ʼ��, sng����)
                    ElseIf str�����λ = "Сʱ" Then
                        sng���� = IIF(intƵ�ʼ�� \ 24 > sng����, intƵ�ʼ�� \ 24, sng����)
                    End If
                    If sng���� = 0 Then sng���� = 1
                End If
                If Nvl(rsTmp!�Ƴ�, 1) > sng���� Then
                    sng���� = Nvl(rsTmp!�Ƴ�, 1)
                End If
                If Val(.TextMatrix(.Row, COL_����)) > sng���� Then
                    sng���� = Val(.TextMatrix(.Row, COL_����))
                End If
                If Val(.TextMatrix(.Row, COL_����)) <> sng���� Then
                    txt����.Text = sng����
                    txt����.Tag = "1"
                End If
                
                'ҩƷ��������:�����װ
                If strƵ�� <> "" And Val(txt����.Text) <> 0 _
                    And Val(.TextMatrix(.Row, COL_����ϵ��)) <> 0 _
                    And Val(.TextMatrix(.Row, COL_�����װ)) <> 0 Then
                    
                    txt����.Text = FormatEx(CalcȱʡҩƷ����( _
                        Val(txt����.Text), sng����, _
                        intƵ�ʴ���, intƵ�ʼ��, str�����λ, _
                        .TextMatrix(.Row, COL_ִ��ʱ��), _
                        Val(.TextMatrix(.Row, COL_����ϵ��)), _
                        Val(.TextMatrix(.Row, COL_�����װ)), _
                        Val(.TextMatrix(.Row, COL_�ɷ����))), 5)
                    txt����.Tag = "1"
                End If
                
                'ҽ������
                If Not IsNull(rsTmp!ҽ������) Then
                    cboҽ������.Text = rsTmp!ҽ������
                    cboҽ������.Tag = "1"
                End If
            End If
        End If
    End With
    
    '����ǰҽ����ҩ;��/�巨�ı仯
    Call AdviceChange
End Sub

Private Sub SetƵ��Input(rsInput As ADODB.Recordset, ByVal int��Χ As Integer)
'���ܣ�����ִ��Ƶ�ʺ����
'������rsInput=�����ѡ��ķ��ؼ�¼
'      int��Χ=1-��ҽ;2-��ҽ;-1-һ����;-2-������
'˵��������÷��������ִ��ʱ�䷽���ı仯
    Dim lng�÷�ID As Long, blnValid As Boolean
    Dim sng���� As Single, dbl���� As Double
    
    cmdƵ��.Tag = rsInput!����
    txtƵ��.Text = rsInput!����
    txtƵ��.Tag = "1"
            
    If cboִ��ʱ��.Enabled Then '"��ѡƵ��"��ҩƷʱ
        With vsAdvice
            '�������ִ��ʱ�䷽���ı仯
            If InStr(",5,6,", .TextMatrix(.Row, COL_���)) > 0 Then
                '���Ҹ�ҩ;����Ӧ����
                lng�÷�ID = .FindRow(CLng(.TextMatrix(.Row, COL_���ID)), .Row + 1)
                If lng�÷�ID <> -1 Then 'δ�ҵ���ҩ;�������,Ӧ�ò�����
                    lng�÷�ID = Val(.TextMatrix(lng�÷�ID, COL_������ĿID))
                Else
                    lng�÷�ID = 0
                End If
            ElseIf RowIn�䷽��(.Row) Then
                '�õ���Ӧ����ҩ�÷�ID
                lng�÷�ID = Val(.TextMatrix(.Row, COL_������ĿID))
            End If
            
            Call Getʱ�䷽��(cboִ��ʱ��, int��Χ, txtƵ��.Text, lng�÷�ID)
            'ȡ�µ�Ƶ�ʵ�Ĭ��ִ��ʱ��
            If cboִ��ʱ��.ListCount > 0 Then
                cboִ��ʱ��.ListIndex = 0
                cboִ��ʱ��.Tag = "1"
            Else
                '�жϵ�ǰִ��ʱ���Ƿ�Ϸ�
                If cboִ��ʱ��.Text <> "" Then
                    blnValid = ExeTimeValid(cboִ��ʱ��.Text, rsInput!Ƶ�ʴ���, rsInput!Ƶ�ʼ��, rsInput!�����λ)
                    If Not blnValid Then '������Ϸ�,����ȡ,���򱣳�
                        cboִ��ʱ��.Text = ""
                        cboִ��ʱ��.Tag = "1"
                    End If
                End If
            End If
            
            '���¼�������
            If mbln���� And InStr(",5,6,", .TextMatrix(.Row, COL_���)) > 0 Then
                sng���� = Val(txt����.Text)
                If sng���� = 0 Then sng���� = 1
                
                If txtƵ��.Text <> "" And Val(txt����.Text) <> 0 _
                    And Val(.TextMatrix(.Row, COL_����ϵ��)) <> 0 _
                    And Val(.TextMatrix(.Row, COL_�����װ)) <> 0 Then
                    
                    txt����.Text = FormatEx(CalcȱʡҩƷ����( _
                        Val(txt����.Text), sng����, rsInput!Ƶ�ʴ���, _
                        rsInput!Ƶ�ʼ��, rsInput!�����λ, cboִ��ʱ��.Text, _
                        Val(.TextMatrix(.Row, COL_����ϵ��)), _
                        Val(.TextMatrix(.Row, COL_�����װ)), _
                        Val(.TextMatrix(.Row, COL_�ɷ����))), 5)
                    txt����.Tag = "1"
                End If
            End If
        End With
    End If
        
    '����ǰҽ��ִ��Ƶ�ʵı仯
    Call AdviceChange
End Sub

Private Sub cbo����ִ��_Click()
    Dim rsTmp As ADODB.Recordset
    Dim lngRow As Long, strSQL As String
    Dim intIdx As Integer, i As Long
    Dim vRect As RECT, blnCancel As Boolean
        
    If cbo����ִ��.ListIndex = -1 Then Exit Sub
    
    If cbo����ִ��.ItemData(cbo����ִ��.ListIndex) = -1 Then
        strSQL = "Select Distinct A.ID,A.����,A.����,A.����" & _
            " From ���ű� A,��������˵�� B" & _
            " Where A.ID=B.����ID And B.������� IN(1,3)" & _
            " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
            " Order by A.����"
        vRect = GetControlRect(cbo����ִ��.Hwnd)
        Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, lbl����ִ��.Caption, , , , , , True, vRect.Left, vRect.Top, txt�÷�.Height, blnCancel, , True)
        If Not rsTmp Is Nothing Then
            intIdx = SeekCboIndex(cbo����ִ��, rsTmp!ID)
            If intIdx <> -1 Then
                cbo����ִ��.ListIndex = intIdx
            Else
                cbo����ִ��.AddItem rsTmp!���� & "-" & rsTmp!����, cbo����ִ��.ListCount - 1
                cbo����ִ��.ItemData(cbo����ִ��.NewIndex) = rsTmp!ID
                cbo����ִ��.ListIndex = cbo����ִ��.NewIndex
            End If
        Else
            If Not blnCancel Then
                MsgBox "û�п������ݣ����ȵ����Ź��������á�", vbInformation, gstrSysName
            End If
            '�ָ������еĿ���(������Click)
            intIdx = SeekCboIndex(cbo����ִ��, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ִ�п���ID)))
            Call zlControl.CboSetIndex(cbo����ִ��.Hwnd, intIdx)
        End If
    Else
        cbo����ִ��.Tag = "1"
        lngRow = vsAdvice.Row
        
        '���¸����˵�ִ�п���ҽ������
       Call AdviceChange
    End If
End Sub

Private Sub cbo����ִ��_GotFocus()
    Call zlControl.TxtSelAll(cbo����ִ��)
End Sub

Private Sub cbo����ִ��_KeyPress(KeyAscii As Integer)
    Dim blnCancel As Boolean
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cbo����ִ��.ListIndex = -1 Then
            Call cbo����ִ��_Validate(blnCancel)
        End If
        If Not blnCancel Then
            If SeekNextControl Then Call cbo����ִ��_Validate(False)
        End If
    End If
End Sub

Private Sub cbo����ִ��_Validate(Cancel As Boolean)
'���ܣ��������������,�Զ�ƥ��ִ�п���
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intIdx As Long, i As Long
    Dim blnLimit As Boolean, StrInput As String
    Dim vRect As RECT, blnCancel As Boolean
    
    If MousePressButton(tbr.Hwnd, tbr.Buttons("�˳�")) Then Exit Sub
    
    If cbo����ִ��.ListIndex <> -1 Then Exit Sub '��ѡ��
    If cbo����ִ��.Text = "" Then Cancel = True: Exit Sub '������
    
    On Error GoTo errH
    
    '�Ƿ���������ѡ�����
    blnLimit = True
    If cbo����ִ��.ListCount > 0 Then
        If cbo����ִ��.ItemData(cbo����ִ��.ListCount - 1) = -1 Then
            blnLimit = False
        End If
    End If
    StrInput = UCase(NeedName(cbo����ִ��.Text))
    strSQL = "Select Distinct A.ID,A.����,A.����,A.����" & _
        " From ���ű� A,��������˵�� B" & _
        " Where A.ID=B.����ID And B.������� IN(1,3)" & _
        " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        " And (Upper(A.����) Like [1] Or Upper(A.����) Like [2] Or Upper(A.����) Like [2])" & _
        " Order by A.����"
    If blnLimit Then
        'Set rsTmp = New ADODB.Recordset
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, StrInput & "%", mstrLike & StrInput & "%")
        For i = 1 To rsTmp.RecordCount
            intIdx = SeekCboIndex(cbo����ִ��, rsTmp!ID)
            If intIdx <> -1 Then cbo����ִ��.ListIndex = intIdx: Exit For
            rsTmp.MoveNext
        Next
        If cbo����ִ��.ListIndex = -1 Then
            MsgBox "δ����Ӧ�Ŀ��ҡ�", vbInformation, gstrSysName
            Cancel = True: Exit Sub
        End If
    Else
        vRect = GetControlRect(cbo����ִ��.Hwnd)
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, lbl����ִ��.Caption, False, "", "", False, False, True, _
            vRect.Left, vRect.Top, txt�÷�.Height, blnCancel, False, True, StrInput & "%", mstrLike & StrInput & "%")
        If Not rsTmp Is Nothing Then
            intIdx = SeekCboIndex(cbo����ִ��, rsTmp!ID)
            If intIdx <> -1 Then
                cbo����ִ��.ListIndex = intIdx
            Else
                cbo����ִ��.AddItem rsTmp!���� & "-" & rsTmp!����, cbo����ִ��.ListCount - 1
                cbo����ִ��.ItemData(cbo����ִ��.NewIndex) = rsTmp!ID
                cbo����ִ��.ListIndex = cbo����ִ��.NewIndex
            End If
        Else
            If Not blnCancel Then
                MsgBox "δ����Ӧ�Ŀ��ҡ�", vbInformation, gstrSysName
            End If
            Cancel = True: Exit Sub
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdAlley_Click()
'���ܣ��Բ��˹���ʷ/����״̬���й���
    'Pass
    Call AdviceCheckWarn(22)
End Sub

Private Sub cmdƵ��_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim int��Χ As Integer, vRect As RECT
        
    With vsAdvice
        int��Χ = GetƵ�ʷ�Χ(.Row)
        strSQL = "Select Rownum as ID,A.����,A.����,A.����," & _
            " A.Ӣ������,A.Ƶ�ʴ���,A.Ƶ�ʼ��,A.�����λ" & _
            " From ����Ƶ����Ŀ A Where A.���÷�Χ=[1]" & _
            " Order by A.����"
        vRect = GetControlRect(txtƵ��.Hwnd)
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "����Ƶ��", False, "", "", False, False, True, _
            vRect.Left, vRect.Top, txtƵ��.Height, blnCancel, False, True, int��Χ)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "û�п��õ�����Ƶ����Ŀ�����ȵ�ҽ��Ƶ�ʹ��������á�", vbInformation, gstrSysName
            End If
            txtƵ��.Text = .TextMatrix(.Row, COL_Ƶ��)
            Call zlControl.TxtSelAll(txtƵ��)
            txtƵ��.SetFocus: Exit Sub
        End If
        Call SetƵ��Input(rsTmp, int��Χ)
        txtƵ��.SetFocus
        Call SeekNextControl
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mblnNoSave Then
        If MsgBox("��ǰҽ�����ݱ༭����δ���棬ȷʵҪ�˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = True
            Exit Sub
        End If
    End If
    mfrmShortCut.SaveShowState 'ϵͳ�Զ�ж�ظ��Ӵ���
End Sub

Private Sub mfrmPrice_PanelHide()
    Call stbThis_PanelClick(stbThis.Panels("Price"))
End Sub

Private Sub mfrmShortCut_ItemClick(ByVal ���� As Integer, ByVal ����ID As Long)
    If cmdSel.Enabled And cmdSel.Visible Then
        Call ClinicSelecter(����ID)
    End If
End Sub

Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Key = "Price" Then
        If Panel.Bevel <> sbrNoBevel Then
            Panel.Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
            Panel.Tag = IIF(Panel.Bevel = sbrInset, "Show", "")
            Call ShowPrice(vsAdvice.Row)
        End If
    ElseIf Panel.Bevel = sbrRaised And (Panel.Key = "PY" Or Panel.Key = "WB") Then
        '�л����������ƥ�䷽ʽ
        Panel.Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        If Panel.Key = "PY" Then
            stbThis.Panels("WB").Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        Else
            stbThis.Panels("PY").Bevel = IIF(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        End If
        SaveSetting "ZLSOFT", "˽��ȫ��\" & gstrDBUser, "��������", _
            IIF(stbThis.Panels("PY").Bevel = sbrInset And stbThis.Panels("WB").Bevel = sbrInset, 2, IIF(stbThis.Panels("WB").Bevel = sbrInset, 1, 0))
        mint���� = IIF(stbThis.Panels("PY").Bevel = sbrInset And stbThis.Panels("WB").Bevel = sbrInset, 2, IIF(stbThis.Panels("WB").Bevel = sbrInset, 1, 0))
    End If
End Sub

Private Sub txtƵ��_GotFocus()
    Call zlControl.TxtSelAll(txtƵ��)
End Sub

Private Sub txtƵ��_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim int��Χ As Integer, vRect As RECT
    
    With vsAdvice
        If KeyAscii = 13 Then
            KeyAscii = 0
            If cmdƵ��.Tag <> "" And txtƵ��.Text = .TextMatrix(.Row, COL_Ƶ��) And txtƵ��.Text <> "" Then
                Call SeekNextControl
            ElseIf txtƵ��.Text = "" Then
                If cmdƵ��.Enabled And cmdƵ��.Visible Then cmdƵ��_Click
            Else
                int��Χ = GetƵ�ʷ�Χ(.Row)
                strSQL = "Select Rownum as ID,A.����,A.����,A.����," & _
                    " A.Ӣ������,A.Ƶ�ʴ���,A.Ƶ�ʼ��,A.�����λ" & _
                    " From ����Ƶ����Ŀ A Where A.���÷�Χ=[3]" & _
                    " And (A.���� Like [1] Or Upper(A.����) Like [2]" & _
                    " Or Upper(A.����) Like [2] Or Upper(A.Ӣ������) Like [2])" & _
                    " Order by A.����"
                vRect = GetControlRect(txtƵ��.Hwnd)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "����Ƶ��", False, "", "", False, False, True, _
                    vRect.Left, vRect.Top, txtƵ��.Height, blnCancel, False, True, UCase(txtƵ��.Text) & "%", mstrLike & UCase(txtƵ��.Text) & "%", int��Χ)
                If rsTmp Is Nothing Then
                    If Not blnCancel Then
                        MsgBox "δ�ҵ�ƥ�������Ƶ����Ŀ��", vbInformation, gstrSysName
                    End If
                    txtƵ��.Text = .TextMatrix(.Row, COL_Ƶ��)
                    Call zlControl.TxtSelAll(txtƵ��)
                    txtƵ��.SetFocus: Exit Sub
                End If
                Call SetƵ��Input(rsTmp, int��Χ)
                Call SeekNextControl
            End If
        End If
    End With
End Sub

Private Sub txt����_Change()
    txt����.Tag = "1"
End Sub

Private Sub txt����_GotFocus()
    Call zlControl.TxtSelAll(txt����)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If (IsNumeric(txt����.Text) Or txt����.Text = "") _
            And (IsNumeric(txt����.Text) Or txt����.Text = "") Then
            If SeekNextControl Then Call txt����_Validate(False)
        End If
    Else
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt����_Validate(Cancel As Boolean)
    Dim sng���� As Single, i As Long
    Dim strSame As String, strMsg As String
        
    If MousePressButton(tbr.Hwnd, tbr.Buttons("�˳�")) Then Exit Sub
    
    With vsAdvice
        If Val(txt����.Text) = 0 Then
            txt����.Text = 1: txt����.Tag = "1"
        End If
        
        '����������Ҫһ��Ƶ��ͬ�ڵ�����
        If Val(.TextMatrix(.Row, COL_Ƶ�ʼ��)) <> 0 Then
            If .TextMatrix(.Row, COL_�����λ) = "��" Then
                sng���� = 7
            ElseIf .TextMatrix(.Row, COL_�����λ) = "��" Then
                sng���� = Val(.TextMatrix(.Row, COL_Ƶ�ʼ��))
            ElseIf .TextMatrix(.Row, COL_�����λ) = "Сʱ" Then
                sng���� = Val(.TextMatrix(.Row, COL_Ƶ�ʼ��)) \ 24
            End If
            If Val(txt����.Text) < sng���� Then
                If MsgBox("��""" & .TextMatrix(.Row, COL_Ƶ��) & """ִ��ʱ��������Ҫ " & sng���� & " �����ҩ��Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Cancel = True: txt����_GotFocus: Exit Sub
                End If
            End If
        End If

        '���¼�������
        If .TextMatrix(.Row, COL_Ƶ��) <> "" _
            And Val(.TextMatrix(.Row, COL_����)) <> 0 _
            And Val(.TextMatrix(.Row, COL_����ϵ��)) <> 0 _
            And Val(.TextMatrix(.Row, COL_�����װ)) <> 0 Then
            
            txt����.Text = FormatEx(CalcȱʡҩƷ����( _
                Val(.TextMatrix(.Row, COL_����)), Val(txt����.Text), _
                Val(.TextMatrix(.Row, COL_Ƶ�ʴ���)), Val(.TextMatrix(.Row, COL_Ƶ�ʼ��)), _
                .TextMatrix(.Row, COL_�����λ), .TextMatrix(.Row, COL_ִ��ʱ��), _
                Val(.TextMatrix(.Row, COL_����ϵ��)), Val(.TextMatrix(.Row, COL_�����װ)), _
                Val(.TextMatrix(.Row, COL_�ɷ����))), 5)
            txt����.Tag = "1"
        End If
    End With
    
    'ÿ��������������Ϊ�´ε�ȱʡ
    If txt����.Tag = "1" Then
        msng���� = Val(txt����.Text)
    End If
    
    Call AdviceChange
    
    '���׷�����������
    With vsAdvice
        If CStr(.Cell(flexcpData, .Row, COL_EDIT)) <> "" Then
            strSame = CStr(.Cell(flexcpData, .Row, COL_EDIT))
            If InStr(strSame, ",") > 0 Then
                strMsg = "�ôθ��Ƶ����е�ҩƷ�����������ִ����"
            Else
                strMsg = "�ó��׷���������ҩƷ�����������ִ����"
            End If
            If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                For i = .FixedRows To .Rows - 1
                    If InStr(",5,6,", .TextMatrix(i, COL_���)) > 0 Then
                        If Not (Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(.Row, COL_���ID)) _
                            Or .RowData(i) = Val(.TextMatrix(.Row, COL_���ID)) Or i = .Row) _
                                And CStr(.Cell(flexcpData, i, COL_EDIT)) = strSame Then
                            If .TextMatrix(i, COL_Ƶ��) <> "" _
                                And Val(.TextMatrix(i, COL_����)) <> 0 _
                                And Val(.TextMatrix(i, COL_����ϵ��)) <> 0 _
                                And Val(.TextMatrix(i, COL_�����װ)) <> 0 Then
                                .TextMatrix(i, COL_����) = txt����.Text
                                .TextMatrix(i, COL_����) = FormatEx(CalcȱʡҩƷ����( _
                                    Val(.TextMatrix(i, COL_����)), Val(txt����.Text), _
                                    Val(.TextMatrix(i, COL_Ƶ�ʴ���)), Val(.TextMatrix(i, COL_Ƶ�ʼ��)), _
                                    .TextMatrix(i, COL_�����λ), .TextMatrix(i, COL_ִ��ʱ��), _
                                    Val(.TextMatrix(i, COL_����ϵ��)), Val(.TextMatrix(i, COL_�����װ)), _
                                    Val(.TextMatrix(i, COL_�ɷ����))), 5)
                            End If
                        End If
                    End If
                Next
            End If
        End If
    End With
End Sub

Private Sub txt�÷�_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim int���� As Integer, vRect As RECT
    Dim lngBegin As Long, lngEnd As Long
    Dim strLike As String, i As Long
    
    With vsAdvice
        If KeyAscii = 13 Then
            KeyAscii = 0
            If Val(cmd�÷�.Tag) <> 0 And txt�÷�.Text = .TextMatrix(.Row, COL_�÷�) And txt�÷�.Text <> "" Then
                Call SeekNextControl
            ElseIf txt�÷�.Text = "" Then
                If cmd�÷�.Enabled And cmd�÷�.Visible Then cmd�÷�_Click
            Else
                If InStr(",5,6,", .TextMatrix(.Row, COL_���)) > 0 Then
                    int���� = 2 '��ҩ;��
                ElseIf RowIn������(vsAdvice.Row) Then
                    int���� = 6 '�ɼ�����
                Else
                    int���� = 4 '��ҩ�÷�
                End If
                If int���� = 2 Then 'ֻȡ��Ч��Χ�ĸ�ҩ;��(�����û��һ��ʱ����ѡ)
                    strSQL = " And (A.ID IN(Select �÷�ID From �����÷����� Where ��ĿID=[4] And ����>0)" & _
                        " Or (Select Count(A.�÷�ID) From �����÷����� A,������ĿĿ¼ B" & _
                            " Where A.�÷�ID=B.ID And B.������� IN(1,3) And A.��ĿID=[4] And A.����>0)<=1)"
                End If
                
                '�Ż�
                strLike = mstrLike
                If Len(txt�÷�.Text) < 2 Then strLike = ""
                
                strSQL = "Select Distinct A.ID,A.����,A.����" & _
                    " From ������ĿĿ¼ A,������Ŀ���� B" & _
                    " Where A.ID=B.������ĿID" & _
                    " And A.���='E' And A.��������=[3] And A.������� IN(1,3)" & strSQL & _
                    " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
                    " And (A.���� Like [1] Or B.���� Like [2] Or B.���� Like [2])" & _
                    Decode(mint����, 0, " And B.���� IN([5],3)", 1, " And B.���� IN([5],3)", "") & _
                    " Order by A.����"
                vRect = GetControlRect(txt�÷�.Hwnd)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, lbl�÷�.Caption, False, "", "", False, False, True, _
                    vRect.Left, vRect.Top, txt�÷�.Height, blnCancel, False, True, UCase(txt�÷�.Text) & "%", _
                    strLike & UCase(txt�÷�.Text) & "%", CStr(int����), Val(.TextMatrix(.Row, COL_������ĿID)), mint���� + 1)
                If rsTmp Is Nothing Then
                    If Not blnCancel Then
                        MsgBox "δ�ҵ�ƥ���" & lbl�÷�.Caption & "��", vbInformation, gstrSysName
                    End If
                    txt�÷�.Text = .TextMatrix(.Row, COL_�÷�)
                    Call zlControl.TxtSelAll(txt�÷�)
                    txt�÷�.SetFocus: Exit Sub
                End If
                
                '��һ����ҩ������ҩƷ�Ŀ��ø�ҩ;�����м��
                If int���� = 2 Then
                    Call Getһ����ҩ��Χ(Val(.TextMatrix(.Row, COL_���ID)), lngBegin, lngEnd)
                    For i = lngBegin To lngEnd
                        If i <> .Row And .RowData(i) <> 0 Then
                            If Not Check�����÷�(rsTmp!ID, Val(.TextMatrix(i, COL_������ĿID)), 1) Then
                                .Refresh
                                MsgBox """" & rsTmp!���� & """���������뵱ǰҩƷһ����ҩ��""" & .TextMatrix(i, COL_ҽ������) & """��", vbInformation, gstrSysName
                                .Refresh
                                txt�÷�.Text = .TextMatrix(.Row, COL_�÷�)
                                Call zlControl.TxtSelAll(txt�÷�)
                                txt�÷�.SetFocus: Exit Sub
                            End If
                        End If
                    Next
                End If
                
                Call Set�÷�Input(rsTmp, int����)
                Call SeekNextControl
            End If
        End If
    End With
End Sub

Private Sub cmd�÷�_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim int���� As Integer, vRect As RECT
    Dim lngBegin As Long, lngEnd As Long, i As Long
        
    With vsAdvice
        If InStr(",5,6,", .TextMatrix(.Row, COL_���)) > 0 Then
            int���� = 2 '��ҩ;��
        ElseIf RowIn������(vsAdvice.Row) Then
            int���� = 6 '�ɼ�����
        Else
            int���� = 4 '��ҩ�÷�
        End If
        If int���� = 2 Then 'ֻȡ��Ч��Χ�ĸ�ҩ;��(�����û��һ��ʱ����ѡ)
            strSQL = " And (A.ID IN(Select �÷�ID From �����÷����� Where ��ĿID=[2] And ����>0)" & _
                " Or (Select Count(A.�÷�ID) From �����÷����� A,������ĿĿ¼ B" & _
                    " Where A.�÷�ID=B.ID And B.������� IN(1,3) And A.��ĿID=[2] And A.����>0)<=1)"
        End If
        strSQL = "Select Distinct A.ID,A.����,A.����,C.���� as ����" & _
            " From ������ĿĿ¼ A,������Ŀ���� B,���Ʒ���Ŀ¼ C" & _
            " Where A.ID=B.������ĿID And A.����ID=C.ID(+)" & _
            " And A.���='E' And A.��������=[1] And A.������� IN(1,3)" & strSQL & _
            " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
            " Order by A.����"
        vRect = GetControlRect(txt�÷�.Hwnd)
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, lbl�÷�.Caption, False, "", "", False, False, True, _
            vRect.Left, vRect.Top, txt�÷�.Height, blnCancel, False, True, CStr(int����), Val(.TextMatrix(.Row, COL_������ĿID)))
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "û�п��õ�" & lbl�÷�.Caption & "�����ȵ�������Ŀ���������á�", vbInformation, gstrSysName
            End If
            txt�÷�.Text = .TextMatrix(.Row, COL_�÷�)
            Call zlControl.TxtSelAll(txt�÷�)
            txt�÷�.SetFocus: Exit Sub
        End If
        
        '��һ����ҩ������ҩƷ�Ŀ��ø�ҩ;�����м��
        If int���� = 2 Then
            Call Getһ����ҩ��Χ(Val(.TextMatrix(.Row, COL_���ID)), lngBegin, lngEnd)
            For i = lngBegin To lngEnd
                If i <> .Row And .RowData(i) <> 0 Then
                    If Not Check�����÷�(rsTmp!ID, Val(.TextMatrix(i, COL_������ĿID)), 1) Then
                        .Refresh
                        MsgBox """" & rsTmp!���� & """���������뵱ǰҩƷһ����ҩ��""" & .TextMatrix(i, COL_ҽ������) & """��", vbInformation, gstrSysName
                        .Refresh
                        txt�÷�.Text = .TextMatrix(.Row, COL_�÷�)
                        Call zlControl.TxtSelAll(txt�÷�)
                        txt�÷�.SetFocus: Exit Sub
                    End If
                End If
            Next
        End If
        
        Call Set�÷�Input(rsTmp, int����)
        txt�÷�.SetFocus
        Call SeekNextControl
    End With
End Sub

Private Sub txt�÷�_GotFocus()
    Call zlControl.TxtSelAll(txt�÷�)
End Sub

Private Sub txt�÷�_Validate(Cancel As Boolean)
    With vsAdvice
        '�ָ���Ϊ�����
        If Val(cmd�÷�.Tag) <> 0 And txt�÷�.Text <> .TextMatrix(.Row, COL_�÷�) Then
            txt�÷�.Text = .TextMatrix(.Row, COL_�÷�)
        End If
    End With
End Sub

Private Sub txtƵ��_Validate(Cancel As Boolean)
    With vsAdvice
        '�ָ���Ϊ�����
        If cmdƵ��.Tag <> "" And txtƵ��.Text <> .TextMatrix(.Row, COL_Ƶ��) Then
            txtƵ��.Text = .TextMatrix(.Row, COL_Ƶ��)
        End If
    End With
End Sub

Private Sub cboӤ��_Click()
    If Not Visible Then Exit Sub
    If cboӤ��.ListIndex = Val(cboӤ��.Tag) Then Exit Sub
    cboӤ��.Tag = cboӤ��.ListIndex
    
    Call ShowAdvice
    
    vsAdvice.SetFocus
End Sub

Private Sub cboִ�п���_Click()
    Dim rsTmp As ADODB.Recordset
    Dim lngRow As Long, strSQL As String
    Dim intIdx As Integer, i As Long
    Dim vRect As RECT, blnCancel As Boolean
        
    If cboִ�п���.ListIndex = -1 Then Exit Sub
    
    If cboִ�п���.ItemData(cboִ�п���.ListIndex) = -1 Then
        strSQL = "Select Distinct A.ID,A.����,A.����,A.����" & _
            " From ���ű� A,��������˵�� B" & _
            " Where A.ID=B.����ID And B.������� IN(1,3)" & _
            " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
            " Order by A.����"
        vRect = GetControlRect(cboִ�п���.Hwnd)
        Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, lblִ�п���.Caption, , , , , , True, vRect.Left, vRect.Top, cboִ�п���.Height, blnCancel, , True)
        If Not rsTmp Is Nothing Then
            intIdx = SeekCboIndex(cboִ�п���, rsTmp!ID)
            If intIdx <> -1 Then
                cboִ�п���.ListIndex = intIdx
            Else
                cboִ�п���.AddItem rsTmp!���� & "-" & rsTmp!����, cboִ�п���.ListCount - 1
                cboִ�п���.ItemData(cboִ�п���.NewIndex) = rsTmp!ID
                cboִ�п���.ListIndex = cboִ�п���.NewIndex
            End If
        Else
            If Not blnCancel Then
                MsgBox "û�п������ݣ����ȵ����Ź��������á�", vbInformation, gstrSysName
            End If
            '�ָ������еĿ���(������Click)
            intIdx = SeekCboIndex(cboִ�п���, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ִ�п���ID)))
            Call zlControl.CboSetIndex(cboִ�п���.Hwnd, intIdx)
        End If
    Else
        cboִ�п���.Tag = "1"
        lngRow = vsAdvice.Row
        
        '���¸����˵�ִ�п���ҽ������
        Call AdviceChange
        
        '���»�ȡ��沢��ʾ�������ﵥλ����ҩ�䷽����ʾ
        With vsAdvice
            If InStr(",5,6,", .TextMatrix(lngRow, COL_���)) > 0 And Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)) <> 0 Then
                Call GetDrugStock(lngRow)
                stbThis.Panels(3).Text = "���: " & FormatEx(Val(.TextMatrix(lngRow, COL_���)), 5) & .TextMatrix(lngRow, COL_���ﵥλ)
            ElseIf RowIn�䷽��(lngRow) Then
                Call GetDrugStock(lngRow)
            End If
        End With
    End If
End Sub

Private Sub cboִ�п���_GotFocus()
    Call zlControl.TxtSelAll(cboִ�п���)
End Sub

Private Sub cboִ�п���_KeyPress(KeyAscii As Integer)
    Dim blnCancel As Boolean
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cboִ�п���.ListIndex = -1 Then
            Call cboִ�п���_Validate(blnCancel)
            cboִ�п���.SetFocus
        Else
            If SeekNextControl Then Call cboִ�п���_Validate(False)
        End If
    End If
End Sub

Private Sub cboִ�п���_Validate(Cancel As Boolean)
'���ܣ��������������,�Զ�ƥ��ִ�п���
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intIdx As Long, i As Long
    Dim blnLimit As Boolean, StrInput As String
    Dim vRect As RECT, blnCancel As Boolean
    
    If MousePressButton(tbr.Hwnd, tbr.Buttons("�˳�")) Then Exit Sub
    
    If cboִ�п���.ListIndex <> -1 Then Exit Sub '��ѡ��
    If cboִ�п���.Text = "" Then Cancel = True: Exit Sub '������
    
    On Error GoTo errH
    
    '�Ƿ���������ѡ�����
    blnLimit = True
    If cboִ�п���.ListCount > 0 Then
        If cboִ�п���.ItemData(cboִ�п���.ListCount - 1) = -1 Then
            blnLimit = False
        End If
    End If
    StrInput = UCase(NeedName(cboִ�п���.Text))
    strSQL = "Select Distinct A.ID,A.����,A.����,A.����" & _
        " From ���ű� A,��������˵�� B" & _
        " Where A.ID=B.����ID And B.������� IN(1,3)" & _
        " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        " And (Upper(A.����) Like [1] Or Upper(A.����) Like [2] Or Upper(A.����) Like [2])" & _
        " Order by A.����"
    If blnLimit Then
        'Set rsTmp = New ADODB.Recordset
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, StrInput & "%", mstrLike & StrInput & "%")
        For i = 1 To rsTmp.RecordCount
            intIdx = SeekCboIndex(cboִ�п���, rsTmp!ID)
            If intIdx <> -1 Then cboִ�п���.ListIndex = intIdx: Exit For
            rsTmp.MoveNext
        Next
        If cboִ�п���.ListIndex = -1 Then
            MsgBox "δ����Ӧ�Ŀ��ҡ�", vbInformation, gstrSysName
            Cancel = True: Exit Sub
        End If
    Else
        vRect = GetControlRect(cboִ�п���.Hwnd)
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, lblִ�п���.Caption, False, "", "", False, False, True, _
            vRect.Left, vRect.Top, txt�÷�.Height, blnCancel, False, True, StrInput & "%", mstrLike & StrInput & "%")
        If Not rsTmp Is Nothing Then
            intIdx = SeekCboIndex(cboִ�п���, rsTmp!ID)
            If intIdx <> -1 Then
                cboִ�п���.ListIndex = intIdx
            Else
                cboִ�п���.AddItem rsTmp!���� & "-" & rsTmp!����, cboִ�п���.ListCount - 1
                cboִ�п���.ItemData(cboִ�п���.NewIndex) = rsTmp!ID
                cboִ�п���.ListIndex = cboִ�п���.NewIndex
            End If
        Else
            If Not blnCancel Then
                MsgBox "δ�ҵ���Ӧ�Ŀ��ҡ�", vbInformation, gstrSysName
            End If
            Cancel = True: Exit Sub
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cboִ��ʱ��_Change()
    cboִ��ʱ��.Tag = "1"
End Sub

Private Sub cboִ��ʱ��_Click()
    'cboִ��ʱ��_Change
    '��������
    cboִ��ʱ��.Tag = "1"
    Call AdviceChange
End Sub

Private Sub cboִ��ʱ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If SeekNextControl Then Call cboִ��ʱ��_Validate(False)
    Else
        If InStr("0123456789:-/" & Chr(8) & Chr(3) & Chr(22), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub cboִ��ʱ��_Validate(Cancel As Boolean)
    Dim blnValid As Boolean, lngRow As Long, strTmp As String
    
    If MousePressButton(tbr.Hwnd, tbr.Buttons("�˳�")) Then Exit Sub
    
    lngRow = vsAdvice.Row
        
    With vsAdvice
        If cboִ��ʱ��.Text <> "" Then
            '��鳤��
            If Len(cboִ��ʱ��.Text) > 50 Then
                MsgBox "�������ݲ��ܳ��� 50 ���ַ���", vbInformation, gstrSysName
                Call cboִ��ʱ��_GotFocus
                Cancel = True: Exit Sub
            End If
            '���Ϸ���
            If .RowData(lngRow) <> 0 Then
                blnValid = ExeTimeValid(cboִ��ʱ��.Text, Val(.TextMatrix(lngRow, COL_Ƶ�ʴ���)), Val(.TextMatrix(lngRow, COL_Ƶ�ʼ��)), .TextMatrix(lngRow, COL_�����λ))
                If Not blnValid Then
                    If .TextMatrix(lngRow, COL_�����λ) = "��" Then
                        strTmp = COL_����ִ��
                    ElseIf .TextMatrix(lngRow, COL_�����λ) = "��" Then
                        strTmp = COL_����ִ��
                    ElseIf .TextMatrix(lngRow, COL_�����λ) = "Сʱ" Then
                        strTmp = COL_��ʱִ��
                    End If
                    MsgBox "�����ִ��ʱ�䷽����ʽ����ȷ�����顣" & vbCrLf & vbCrLf & "����" & vbCrLf & strTmp, vbInformation, gstrSysName
                    Call cboִ��ʱ��_GotFocus
                    Cancel = True: Exit Sub
                End If
            End If
        End If
    End With
    
    '��������
    Call AdviceChange
End Sub

Private Sub cboִ������_Click()
    cboִ������.Tag = "1"
    '��������
    Call AdviceChange
End Sub

Private Sub cboִ������_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cboִ������.ListIndex <> -1 Then
            Call SeekNextControl
        End If
    ElseIf KeyAscii >= 32 Then
        lngIdx = zlControl.CboMatchIndex(cboִ������.Hwnd, KeyAscii)
        If lngIdx = -1 And cboִ������.ListCount > 0 Then lngIdx = 0
        cboִ������.ListIndex = lngIdx
    End If
End Sub

Private Sub chk����_Click()
    If Not mblnDoCheck Then Exit Sub
    
    chk����.Tag = "1"
    '��������
    Call AdviceChange
End Sub

Private Sub chk����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call SeekNextControl
    End If
End Sub

Private Sub cmdExt_Click()
'���ܣ��޸�����ҽ������������
    Dim rsCurr As New ADODB.Recordset
    Dim strExtData As String, strTmp As String
    Dim lngRow As Long, lngDrugRow As Long
    Dim lng������ĿID As Long, lng�÷�ID As Long
    
    lngRow = vsAdvice.Row
    
    If vsAdvice.TextMatrix(lngRow, COL_���) = "D" Then
        strExtData = Get��鲿λIDs(lngRow)
        frmAdviceEditEx.mintType = 0
    ElseIf vsAdvice.TextMatrix(lngRow, COL_���) = "F" Then
        strExtData = Get��������IDs(lngRow)
        frmAdviceEditEx.mintType = 1
    ElseIf RowIn�䷽��(lngRow) Then
        strExtData = Get��ҩ�䷽IDs(lngRow)
        frmAdviceEditEx.mintType = 2
    ElseIf RowIn������(lngRow) Then
        strExtData = Get�������IDs(lngRow)
        frmAdviceEditEx.mintType = 4
    Else
        Exit Sub '������ǰ�ļ�����Ŀ
    End If
    
    frmAdviceEditEx.mstrPrivs = mstrPrivs
    frmAdviceEditEx.mlngHwnd = txtҽ������.Hwnd
    frmAdviceEditEx.mint��Ч = 1
    frmAdviceEditEx.mstr�Ա� = mstr�Ա�
    frmAdviceEditEx.mint������� = 1 '����
    If frmAdviceEditEx.mintType = 4 Then
        frmAdviceEditEx.mlng��ĿID = 0
    Else
        frmAdviceEditEx.mlng��ĿID = Val(vsAdvice.TextMatrix(lngRow, COL_������ĿID))
    End If
    frmAdviceEditEx.mstrExtData = strExtData
    
    frmAdviceEditEx.mblnҽ�� = InStr(",1,2,", mstr������) > 0 And mstr������ <> ""
    
    On Error Resume Next
    frmAdviceEditEx.Show 1, Me
    On Error GoTo 0
    
    '���������������
    If frmAdviceEditEx.mblnOK Then
        strExtData = frmAdviceEditEx.mstrExtData
        
        '���¿���ʱ��
        vsAdvice.TextMatrix(lngRow, COL_����ʱ��) = Format(zlDatabase.Currentdate, "MM-dd HH:mm")
        vsAdvice.Cell(flexcpData, lngRow, COL_����ʱ��) = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
        vsAdvice.TextMatrix(lngRow, COL_����) = "" '������¼���
        
        If vsAdvice.TextMatrix(lngRow, COL_���) = "D" Then
            '������
            Call AdviceSet�������(1, lngRow, strExtData)
            vsAdvice.TextMatrix(lngRow, COL_ҽ������) = AdviceTextMake(lngRow)
            txtҽ������.Text = vsAdvice.TextMatrix(lngRow, COL_ҽ������)
        ElseIf vsAdvice.TextMatrix(lngRow, COL_���) = "F" Then
            'һ������
            Call AdviceSet�������(2, lngRow, strExtData)
            vsAdvice.TextMatrix(lngRow, COL_ҽ������) = AdviceTextMake(lngRow)
            txtҽ������.Text = vsAdvice.TextMatrix(lngRow, COL_ҽ������)
            
            'ˢ�´������������ִ�п���
            Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
        ElseIf RowIn������(lngRow) Then
            '�������
            lngDrugRow = vsAdvice.FindRow(CStr(vsAdvice.RowData(lngRow)), , COL_���ID)
            lng�÷�ID = Val(vsAdvice.TextMatrix(lngRow, COL_������ĿID))
            
            '�Ȼ�ȡ��ǰ�Ѿ����ú�ֵ
            rsCurr.Fields.Append "Edit", adInteger, , adFldIsNullable
            rsCurr.Fields.Append "ҽ��ID", adBigInt, , adFldIsNullable
            rsCurr.Fields.Append "ִ�п���ID", adBigInt, , adFldIsNullable
            rsCurr.Fields.Append "Ƶ��", adVarChar, 20, adFldIsNullable
            rsCurr.Fields.Append "Ƶ�ʴ���", adInteger, , adFldIsNullable
            rsCurr.Fields.Append "Ƶ�ʼ��", adInteger, , adFldIsNullable
            rsCurr.Fields.Append "�����λ", adVarChar, 4, adFldIsNullable
            rsCurr.Fields.Append "����", adDouble, , adFldIsNullable
            rsCurr.Fields.Append "ִ��ʱ��", adVarChar, 50, adFldIsNullable
            rsCurr.Fields.Append "��ʼʱ��", adVarChar, 20, adFldIsNullable
            rsCurr.Fields.Append "����ҽ��", adVarChar, 20, adFldIsNullable
            rsCurr.Fields.Append "��������ID", adBigInt, , adFldIsNullable
            rsCurr.Fields.Append "����ʱ��", adVarChar, 20, adFldIsNullable
            rsCurr.Fields.Append "ҽ������", adVarChar, 100, adFldIsNullable
            rsCurr.Fields.Append "��־", adVarChar, 4, adFldIsNullable
            
            rsCurr.CursorLocation = adUseClient
            rsCurr.LockType = adLockOptimistic
            rsCurr.CursorType = adOpenStatic
            rsCurr.Open
            rsCurr.AddNew
                        
            '�ɼ�������ִ�п��ҿ����������Ŀ��ͬ
            If Val(vsAdvice.TextMatrix(lngDrugRow, COL_ִ�п���ID)) <> 0 Then
                rsCurr!ִ�п���ID = Val(vsAdvice.TextMatrix(lngDrugRow, COL_ִ�п���ID))
            End If
            If Val(vsAdvice.TextMatrix(lngRow, COL_����)) <> 0 Then
                rsCurr!���� = Val(vsAdvice.TextMatrix(lngRow, COL_����))
            End If
            rsCurr!ִ��ʱ�� = vsAdvice.TextMatrix(lngRow, COL_ִ��ʱ��)
            rsCurr!Ƶ�� = vsAdvice.TextMatrix(lngRow, COL_Ƶ��)
            rsCurr!Ƶ�ʴ��� = Val(vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʴ���))
            rsCurr!Ƶ�ʼ�� = Val(vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʼ��))
            rsCurr!�����λ = vsAdvice.TextMatrix(lngRow, COL_�����λ)
            rsCurr!��ʼʱ�� = vsAdvice.Cell(flexcpData, lngRow, COL_��ʼʱ��)
            rsCurr!����ҽ�� = vsAdvice.TextMatrix(lngRow, COL_����ҽ��)
            rsCurr!��������ID = Val(vsAdvice.TextMatrix(lngRow, COL_��������ID))
            rsCurr!����ʱ�� = vsAdvice.Cell(flexcpData, lngRow, COL_����ʱ��)
            rsCurr!ҽ������ = vsAdvice.TextMatrix(lngRow, COL_ҽ������)
            rsCurr!��־ = vsAdvice.TextMatrix(lngRow, COL_��־)
            '�޸��˼����������,�ɼ�������Ӧ���Ϊ�޸�
            rsCurr!Edit = Val(vsAdvice.TextMatrix(lngRow, COL_EDIT))
            rsCurr!ҽ��ID = vsAdvice.RowData(lngRow)
            rsCurr.Update
            
            '��ȫ�������øü������
            '------------------------
            'ɾ��������Ŀ��:ɾ��֮�����¶�λ�ĵ�ǰ��
            lngRow = Delete�������(lngRow)
            '�����ǰ��(�ɼ�������)
            Call DeleteRow(lngRow, True, False)
            '���²���:����֮�����¶�λ�ĵ�ǰ��
            lngRow = AdviceSet�������(lngRow, lng�÷�ID, strExtData, rsCurr)
        ElseIf RowIn�䷽��(lngRow) Then
            '��ҩ�䷽
            lngDrugRow = vsAdvice.FindRow(CStr(vsAdvice.RowData(lngRow)), , COL_���ID)
            lng������ĿID = Val(vsAdvice.TextMatrix(lngDrugRow, COL_������ĿID))
            lng�÷�ID = Val(vsAdvice.TextMatrix(lngRow, COL_������ĿID))
            
            '�Ȼ�ȡ��ǰ�Ѿ����ú�ֵ
            rsCurr.Fields.Append "Edit", adInteger, , adFldIsNullable
            rsCurr.Fields.Append "ҽ��ID", adBigInt, , adFldIsNullable
            rsCurr.Fields.Append "ִ������", adVarChar, 10, adFldIsNullable
            rsCurr.Fields.Append "ִ�п���ID", adBigInt, , adFldIsNullable
            rsCurr.Fields.Append "Ƶ��", adVarChar, 20, adFldIsNullable
            rsCurr.Fields.Append "Ƶ�ʴ���", adInteger, , adFldIsNullable
            rsCurr.Fields.Append "Ƶ�ʼ��", adInteger, , adFldIsNullable
            rsCurr.Fields.Append "�����λ", adVarChar, 4, adFldIsNullable
            rsCurr.Fields.Append "����", adDouble, , adFldIsNullable
            rsCurr.Fields.Append "ִ��ʱ��", adVarChar, 50, adFldIsNullable
            rsCurr.Fields.Append "��ʼʱ��", adVarChar, 20, adFldIsNullable
            rsCurr.Fields.Append "����ҽ��", adVarChar, 20, adFldIsNullable
            rsCurr.Fields.Append "��������ID", adBigInt, , adFldIsNullable
            rsCurr.Fields.Append "����ʱ��", adVarChar, 20, adFldIsNullable
            rsCurr.Fields.Append "ҽ������", adVarChar, 100, adFldIsNullable
            rsCurr.Fields.Append "��־", adVarChar, 4, adFldIsNullable
            
            rsCurr.CursorLocation = adUseClient
            rsCurr.LockType = adLockOptimistic
            rsCurr.CursorType = adOpenStatic
            rsCurr.Open
            rsCurr.AddNew
            
            rsCurr!ִ������ = NeedName(cboִ������.Text) '����,�Ա�ҩ,��Ժ��ҩ
            If Val(vsAdvice.TextMatrix(lngDrugRow, COL_ִ�п���ID)) <> 0 Then
                rsCurr!ִ�п���ID = Val(vsAdvice.TextMatrix(lngDrugRow, COL_ִ�п���ID))
            End If
            rsCurr!Ƶ�� = vsAdvice.TextMatrix(lngDrugRow, COL_Ƶ��)
            rsCurr!Ƶ�ʴ��� = Val(vsAdvice.TextMatrix(lngDrugRow, COL_Ƶ�ʴ���))
            rsCurr!Ƶ�ʼ�� = Val(vsAdvice.TextMatrix(lngDrugRow, COL_Ƶ�ʼ��))
            rsCurr!�����λ = vsAdvice.TextMatrix(lngDrugRow, COL_�����λ)
            If Val(vsAdvice.TextMatrix(lngDrugRow, COL_����)) <> 0 Then
                rsCurr!���� = Val(vsAdvice.TextMatrix(lngDrugRow, COL_����))
            End If
            rsCurr!ִ��ʱ�� = vsAdvice.TextMatrix(lngDrugRow, COL_ִ��ʱ��)
            rsCurr!��ʼʱ�� = vsAdvice.Cell(flexcpData, lngDrugRow, COL_��ʼʱ��)
            rsCurr!����ҽ�� = vsAdvice.TextMatrix(lngDrugRow, COL_����ҽ��)
            rsCurr!��������ID = Val(vsAdvice.TextMatrix(lngDrugRow, COL_��������ID))
            rsCurr!����ʱ�� = vsAdvice.Cell(flexcpData, lngDrugRow, COL_����ʱ��)
            rsCurr!ҽ������ = vsAdvice.TextMatrix(lngRow, COL_ҽ������)
            rsCurr!��־ = vsAdvice.TextMatrix(lngRow, COL_��־)
            '�޸����䷽����,�÷���Ӧ���Ϊ�޸�
            rsCurr!Edit = Val(vsAdvice.TextMatrix(lngRow, COL_EDIT))
            rsCurr!ҽ��ID = vsAdvice.RowData(lngRow)
            
            rsCurr.Update
            
            '��ȫ�������ø���ҩ�䷽��
            '------------------------
            'ɾ�����ζҩ���巨��:ɾ��֮�����¶�λ�ĵ�ǰ��
            lngRow = Delete��ҩ�䷽(lngRow)
            '�����ǰ��(��ҩ�÷���)
            Call DeleteRow(lngRow, True, False)
            '�����䷽:����֮�����¶�λ�ĵ�ǰ��
            lngRow = AdviceSet��ҩ�䷽(lng������ĿID, lngRow, lng�÷�ID, strExtData, rsCurr)
        End If
        
        'ǿ����ʾ��ǰҽ����Ƭ
        Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
                
        Call CalcAdviceMoney '��ʾ�¿�ҽ�����
        
        If InStr(",0,3,", vsAdvice.TextMatrix(lngRow, COL_EDIT)) > 0 Then
            vsAdvice.TextMatrix(lngRow, COL_EDIT) = 2 '���Ϊ���޸�
            vsAdvice.TextMatrix(lngRow, COL_״̬) = 1 '�޸ĺ��Ϊ�¿�
            Call ReSetColor(lngRow)
        End If
        
        mblnNoSave = True '���Ϊδ����
    End If
    
    Call vsAdvice.AutoSize(COL_ҽ������)
    
    txtҽ������.SetFocus
End Sub

Private Sub ClinicSelecter(Optional ByVal lng����ID As Long)
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = frmClinicSelect.ShowSelect(Me, mstrPrivs, 1, mstr�Ա�, , , 1, lng����ID)
    If rsTmp Is Nothing Then 'ȡ����������
        zlControl.TxtSelAll txtҽ������
        txtҽ������.SetFocus: Exit Sub
    End If
        
    '����ѡ����Ŀ����ȱʡҽ����Ϣ
    If AdviceInput(rsTmp, vsAdvice.Row) Then
        '��ʾ��ȱʡ���õ�ֵ
        Call vsAdvice_AfterRowColChange(-1, vsAdvice.Col, vsAdvice.Row, vsAdvice.Col)
        Call CalcAdviceMoney '��ʾ�¿�ҽ�����
        
        txtҽ������.SetFocus '�����ȶ�λ
        Call SeekNextControl
    Else
        '�ָ�ԭֵ(AdviceInput�����п��ܴ�����һ��)
        txtҽ������.Text = vsAdvice.TextMatrix(vsAdvice.Row, COL_ҽ������)
        txtҽ������.SetFocus
    End If
End Sub

Private Sub cmdSel_Click()
    Call ClinicSelecter
End Sub

Private Sub cmd��ʼʱ��_Click()
    If IsDate(txt��ʼʱ��.Text) Then
        dtpDate.Value = CDate(txt��ʼʱ��.Text)
    Else
        dtpDate.Value = zlDatabase.Currentdate
    End If
    dtpDate.Tag = "��ʼʱ��"
    dtpDate.Left = txt��ʼʱ��.Left + fraAdvice.Left
    dtpDate.Top = txt��ʼʱ��.Top + fraAdvice.Top - dtpDate.Height
    dtpDate.Visible = True
    dtpDate.SetFocus
End Sub

Private Sub dtpDate_DateClick(ByVal DateClicked As Date)
    Dim strDate As String
    
    If dtpDate.Tag = "��ʼʱ��" Then
        'ȡֵ
        If IsDate(txt��ʼʱ��.Text) Then
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(txt��ʼʱ��.Text, "yyyy-MM-dd HH:mm"), 12, 5)
        Else
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm"), 12, 5)
        End If
        
        '�ж�ʱ��Ϸ���
        If Not Check��ʼʱ��(strDate) Then
            dtpDate.SetFocus: Exit Sub
        End If
        
        txt��ʼʱ��.Text = strDate
        dtpDate.Tag = ""
        dtpDate.Visible = False
        Call txt��ʼʱ��_Validate(False) '��������
        txt��ʼʱ��.SetFocus
    End If
End Sub

Private Sub dtpDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call dtpDate_DateClick(dtpDate.Value)
    End If
End Sub

Private Sub dtpDate_Validate(Cancel As Boolean)
    dtpDate.Visible = False
    dtpDate.Tag = ""
End Sub

Private Sub Form_Activate()
    If mblnRunFirst Then
        mblnRunFirst = False
        If txtҽ������.Enabled Then txtҽ������.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbAltMask Then
        If KeyCode = vbKeyX Then
            If tbr.Buttons("�˳�").Enabled And tbr.Buttons("�˳�").Visible Then
                Call tbr_ButtonClick(tbr.Buttons("�˳�"))
            End If
        ElseIf Between(Chr(KeyCode), "1", "6") Then
            Call mfrmShortCut.ShowShortCut(Val(Chr(KeyCode)))
        End If
    ElseIf Shift = vbCtrlMask Then
        Select Case KeyCode
            Case vbKeyA
                If tbr.Buttons("����").Enabled And tbr.Buttons("����").Visible Then
                    Call tbr_ButtonClick(tbr.Buttons("����"))
                End If
            Case vbKeyI
                If tbr.Buttons("����").Enabled And tbr.Buttons("����").Visible Then
                    Call tbr_ButtonClick(tbr.Buttons("����"))
                End If
            Case vbKeyK
                If tbr.Buttons("һ��").Enabled And tbr.Buttons("һ��").Visible Then
                    tbr.Buttons("һ��").Value = IIF(tbr.Buttons("һ��").Value = tbrPressed, tbrUnpressed, tbrPressed)
                    Call tbr_ButtonClick(tbr.Buttons("һ��"))
                End If
            Case vbKeyR
                If tbr.Buttons("����").Enabled And tbr.Buttons("����").Visible Then
                    Call tbr_ButtonClick(tbr.Buttons("����"))
                End If
            Case vbKeyY
                If tbr.Buttons("����").Enabled And tbr.Buttons("����").Visible Then
                    Call tbr_ButtonClick(tbr.Buttons("����"))
                End If
            Case vbKeyT
                If tbr.Buttons("����").Visible And tbr.Buttons("����").Enabled Then
                    Call tbr_ButtonClick(tbr.Buttons("����"))
                End If
            Case vbKeyS
                If tbr.Buttons("����").Enabled And tbr.Buttons("����").Visible Then
                    Call tbr_ButtonClick(tbr.Buttons("����"))
                End If
        End Select
    Else
        Select Case KeyCode
            Case vbKeyEscape
                If dtpDate.Visible Then
                    dtpDate.Visible = False
                    dtpDate.Tag = ""
                End If
            Case vbKeyF4
                If Me.ActiveControl Is txt��ʼʱ�� Then
                    If cmd��ʼʱ��.Visible And cmd��ʼʱ��.Enabled Then cmd��ʼʱ��_Click
                ElseIf Me.ActiveControl Is txtҽ������ Then
                    If cmdExt.Visible And cmdExt.Enabled Then cmdExt_Click
                ElseIf Me.ActiveControl Is txt�÷� Then
                    If cmd�÷�.Visible And cmd�÷�.Enabled Then cmd�÷�_Click
                ElseIf Me.ActiveControl Is txtƵ�� Then
                    If cmdƵ��.Visible And cmdƵ��.Enabled Then cmdƵ��_Click
                End If
            Case vbKeyF1
                Call tbr_ButtonClick(tbr.Buttons("����"))
            Case vbKeyF2
                If tbr.Buttons("����").Enabled And tbr.Buttons("����").Visible Then
                    Call tbr_ButtonClick(tbr.Buttons("����"))
                End If
            Case vbKeyF3
                If tbr.Buttons("����").Visible And tbr.Buttons("����").Enabled Then
                    Call tbr_ButtonClick(tbr.Buttons("����"))
                End If
            Case vbKeyF6
                If tbr.Buttons("�ο�").Visible And tbr.Buttons("�ο�").Enabled Then
                    Call tbr_ButtonClick(tbr.Buttons("�ο�"))
                End If
            Case vbKeyF7 '�л����뷨
                If stbThis.Panels("WB").Visible And stbThis.Panels("PY").Visible Then
                    If stbThis.Panels("WB").Bevel = sbrRaised Then
                        Call stbThis_PanelClick(stbThis.Panels("WB"))
                    Else
                        Call stbThis_PanelClick(stbThis.Panels("PY"))
                    End If
                End If
            Case vbKeyF8 '�л���ʾ�Ƽ���Ŀ
                If stbThis.Panels("Price").Visible Then
                    Call stbThis_PanelClick(stbThis.Panels("Price"))
                End If
        End Select
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    ElseIf KeyAscii = Asc("`") Then
        KeyAscii = 0
        Call mfrmShortCut.ShowMe(Me)
    End If
End Sub

Private Sub Form_Load()
    Call InitAdviceTable
    Call RestoreWinState(Me, App.ProductName)
    Call zlControl.CboSetHeight(cboִ�п���, Me.Height)
    Call zlControl.CboSetWidth(cboִ�п���.Hwnd, cboִ�п���.Width * 1.3)
    
    mblnOK = False
    mblnNoSave = False
    mblnRunFirst = True
    mblnRowChange = True
    mblnDoCheck = True
    mstrDelIDs = ""
    
    '���˹���ʷ/����״̬���ü��
    mlngPassPati = 0
    If gblnPass And InStr(mstrPrivs, "������ҩ���") > 0 Then  'Pass
        cmdAlley.Visible = True
        vsAdvice.ColHidden(COL_��ʾ) = False
        cmdAlley.Enabled = PassGetState("AlleyEnable") = 1
    End If
    
    'Ȩ������
    If InStr(mstrPrivs, "���Ʋο�") = 0 And mlngǰ��ID = 0 Then
        tbr.Buttons("�ο�").Visible = False
        tbr.Buttons("�ο�_").Visible = False
    End If
'    If InStr(mstrPrivs, "������׷���") = 0 Then
'        tbr.Buttons("����").Visible = False
'    End If
    'ҽ��վ��ҽ��ʱ���ɷ���
    If mlngǰ��ID <> 0 Then
        tbr.Buttons("����").Visible = False
        tbr.Buttons("����_").Visible = False
    End If
    
    '����ǩ������
    If gobjESign Is Nothing Then
        tbr.Buttons("ǩ��").Visible = False
    End If
    
    '����ƥ��
    mstrLike = IIF(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = 0, "%", "")
    '����ƥ�䷽ʽ��0-ƴ��,1-���
    mint���� = Val(GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDBUser, "��������", 0))
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
    
    '�Ƽ����״̬
    If mblnModal Then
        stbThis.Panels("Price").Visible = False
    Else
        Set mfrmPrice = New frmAdvicePrice
        stbThis.Panels("Price").Tag = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & mfrmParent.Name, "PricePaneVisible", "")
    End If
    
    'ִ������
    mbln���� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "ҽ��ִ������", 0)) <> 0
    '�Զ�����Ƥ��
    mbln�Զ�Ƥ�� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "�Զ�����Ƥ��", 0)) <> 0 And mlngǰ��ID = 0
    
    'ҩƷ�����鷽ʽ:������ʱû��
    Set mcolStock = InitStockCheck(1)
    
    '��������
    Call ReadEnjoin
    'ҽ�����ݶ���
    Call InitAdviceDefine
    '--------------------------------------------------
    '��ȡ������Ϣ
    Call GetPatiInfo
    Call SetBabyVisible(mlng���˿���id)
        
    '�޸�ʱǿ�ж�λӤ��
    If mlngҽ��ID = 0 Then '����
        cboӤ��.ListIndex = 0 'ȱʡ�������˵�ҽ��
    Else '�޸�
        cboӤ��.ListIndex = mintӤ��
    End If
    cboӤ��.Tag = cboӤ��.ListIndex
    
    '��ȡ����ʾ����ҽ��
    Call ReLoadAdvice(mlngҽ��ID)
    
    '���������봰��
    Set mfrmShortCut = New frmClinicShortCut
    mfrmShortCut.ShowMe Me, True '�����ϴ��Ϸ���ʾ
End Sub

Private Function GetStockCheck(ByVal lng�ⷿID As Long) As Integer
'���ܣ���ȡָ���ⷿ�ĳ������鷽ʽ
    Dim intStyle As Integer
    On Error Resume Next
    intStyle = mcolStock("_" & lng�ⷿID)
    Err.Clear: On Error GoTo 0
    GetStockCheck = intStyle
End Function

Private Sub InitAdviceDefine()
'���ܣ���ʼ��ҽ�����ݶ����������
'˵������mrsDefine��ΪNothingʱ����������ʹ��
    Dim strSQL As String
    
    On Error Resume Next
    Set mobjVBA = CreateObject("ScriptControl")
    Err.Clear: On Error GoTo 0
    
    If Not mobjVBA Is Nothing Then
        mobjVBA.Language = "VBScript"
        Set mobjScript = New clsScript
        mobjVBA.AddObject "clsScript", mobjScript, True
        
        On Error GoTo errH
        strSQL = "Select �������,ҽ������ From ҽ�����ݶ��� Order by �������"
        Set mrsDefine = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(mrsDefine, strSQL, Me.Caption)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Set mrsDefine = Nothing
End Sub

Private Sub ReLoadAdvice(Optional ByVal lngҽ��ID As Long)
'���ܣ����¶�ȡ����ʾ���˵ĵ�ǰҽ���嵥
'������lngҽ��ID=���ڶ�λ
    Dim lngRow As Long
    
    If LoadAdvice Then
        '��ʾҽ��
        Call ShowAdvice
        
        If lngҽ��ID = 0 Then
            If vsAdvice.RowData(vsAdvice.Row) <> 0 Then
                Call tbr_ButtonClick(tbr.Buttons("����"))
            End If
        Else
            '�޸ĵ�ҽ��IDӦ������ʾ��
            lngRow = vsAdvice.FindRow(lngҽ��ID)
            If lngRow <> -1 Then
                If Not vsAdvice.RowHidden(lngRow) Then
                    mblnRowChange = False
                    vsAdvice.Col = COL_ҽ������: vsAdvice.Row = lngRow
                    Call vsAdvice.ShowCell(vsAdvice.Row, vsAdvice.Col)
                    mblnRowChange = True
                End If
            End If
        End If
        '����ʱ������ShowAdvice�еĵ���,ǿ�н���
        Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
    End If
End Sub

Private Function ReadEnjoin() As Boolean
'���ܣ���ȡ�����볣������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strPre As String
        
    On Error GoTo errH
    
    strPre = cboҽ������.Text '����󱣳�ԭ��ֵ
    cboҽ������.Clear
    
    strSQL = "Select Upper(����) as ����,����,Upper(����) as ��д��,Upper(����) as ���� From �������� Where ���� is Not Null Order by ����"
    rsTmp.CursorLocation = adUseClient
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    Do While Not rsTmp.EOF
        AddComboItem cboҽ������.Hwnd, CB_ADDSTRING, 0, rsTmp!����
        rsTmp.MoveNext
    Loop
    cboҽ������.Text = strPre
    ReadEnjoin = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Resize()
    If dtpDate.Visible Then
        dtpDate.Visible = False
        dtpDate.Tag = ""
    End If
    
    On Error Resume Next
    
    fraPati.Left = 0
    fraPati.Top = cbr.Height
    fraPati.Width = Me.ScaleWidth
    
    vsAdvice.Left = 0
    vsAdvice.Top = cbr.Height + fraPati.Height
    vsAdvice.Height = Me.ScaleHeight - fraPati.Height - cbr.Height - stbThis.Height - (fraAdvice.Height - 80)
    vsAdvice.Width = Me.ScaleWidth
    
    fraAdvice.Left = 0
    fraAdvice.Top = vsAdvice.Top + vsAdvice.Height - 80
    fraAdvice.Width = Me.ScaleWidth
    
    'Pass
    cmdAlley.Left = Me.ScaleWidth - cmdAlley.Width - 30
    cboӤ��.Left = Me.ScaleWidth - IIF(cmdAlley.Visible, cmdAlley.Width + 30, 0) - cboӤ��.Width - 30
    lblӤ��.Left = cboӤ��.Left - lblӤ��.Width - 30
    
    If cmdAlley.Visible Or lblӤ��.Visible Then
        lblPati.Width = IIF(lblӤ��.Visible, lblӤ��.Left, cmdAlley.Left) - lblPati.Left - 90
    End If
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mdat�Һ�ʱ�� = Empty
    msng���� = 0
    Set mobjVBA = Nothing
    Set mobjScript = Nothing
    Set mrsDefine = Nothing
    
    '�Ƽ����״̬
    If Not mfrmPrice Is Nothing Then
        Unload mfrmPrice
        Set mfrmPrice = Nothing
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & mfrmParent.Name, "PricePaneVisible", stbThis.Panels("Price").Tag
    End If
    
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Function RowCanMerge(ByVal lngRow1 As Long, ByVal lngRow2 As Long, Optional strMsg As String) As Boolean
'���ܣ��ж������Ƿ����һ����ҩ
'������lngRow1=ǰ��һ���Ѿ������ҩƷ��
'      lngRow2=��ǰ��(�������δ����)
'���أ���������ԣ���strMsg������ʾ��Ϣ
    Dim lngFind As Long, lngRxCount As Long
    
    With vsAdvice
        strMsg = ""
        If Not Between(lngRow1, .FixedRows, .Rows - 1) Then Exit Function
        If Not Between(lngRow2, .FixedRows, .Rows - 1) Then Exit Function
        If .RowHidden(lngRow1) Or .RowHidden(lngRow2) Then Exit Function
        If .RowData(lngRow1) = 0 Then Exit Function
        
        If .RowData(lngRow2) = 0 Then
            '����ȫ��Ϊ��ҩ�������ͬ
            If InStr(",5,6,", .TextMatrix(lngRow1, COL_���)) = 0 Then
                strMsg = "һ����ҩ��ҩƷ���붼Ϊ����ҩ��Ϊ�г�ҩ��"
                Exit Function
            End If
            
            '���ܰ����ѷ��͵�ҽ��
            If Val(.TextMatrix(lngRow1, COL_״̬)) <> 1 Then
                strMsg = "Ҫ����Ϊһ����ҩ��ҩƷ�����Ѿ����͵�ҽ����"
                Exit Function
            End If
            '���ܰ�����ǩ����ҽ��
            If Val(.TextMatrix(lngRow1, COL_ǩ����)) = 1 Then
                strMsg = "Ҫ����Ϊһ����ҩ��ҩƷ�����Ѿ�ǩ����ҽ����"
                Exit Function
            End If
        ElseIf .RowData(lngRow2) <> 0 Then
            '����ȫ��Ϊ��ҩ�������ͬ
'            If Not (.TextMatrix(lngRow1, COL_���) = .TextMatrix(lngRow2, COL_���) _
'                And InStr(",5,6,", .TextMatrix(lngRow1, COL_���)) > 0) Then
'                strMsg = "һ����ҩ��ҩƷ���붼Ϊ����ҩ��Ϊ�г�ҩ��"
'                Exit Function
'            End If
            If InStr(",5,6,", .TextMatrix(lngRow1, COL_���)) = 0 _
                Or InStr(",5,6,", .TextMatrix(lngRow2, COL_���)) = 0 Then
                strMsg = "һ����ҩ��ҩƷ���붼Ϊ����ҩ��Ϊ�г�ҩ��"
                Exit Function
            End If
            
            '���ܰ����ѷ��͵�ҽ��
            If Val(.TextMatrix(lngRow1, COL_״̬)) <> 1 Or Val(.TextMatrix(lngRow2, COL_״̬)) <> 1 Then
                strMsg = "Ҫ����Ϊһ����ҩ��ҩƷ�����Ѿ����͵�ҽ����"
                Exit Function
            End If
            '���ܰ�����ǩ����ҽ��
            If Val(.TextMatrix(lngRow1, COL_ǩ����)) = 1 Or Val(.TextMatrix(lngRow2, COL_ǩ����)) = 1 Then
                strMsg = "Ҫ����Ϊһ����ҩ��ҩƷ�����Ѿ�ǩ����ҽ����"
                Exit Function
            End If
            
            'һ����ҩ(ǰ��ҩƷ)�ĸ�ҩ;���Ƿ������ڵ�ǰҩƷ
            lngFind = .FindRow(CLng(.TextMatrix(lngRow1, COL_���ID)), lngRow1 + 1)
            If lngFind <> -1 Then
                If Not Check�����÷�(Val(.TextMatrix(lngFind, COL_������ĿID)), Val(.TextMatrix(lngRow2, COL_������ĿID)), 1) Then
                    strMsg = """" & .TextMatrix(lngRow2, COL_ҽ������) & """����ʹ��""" & .TextMatrix(lngFind, COL_ҽ������) & """��ҩ;����" & _
                    vbCrLf & "������""" & .TextMatrix(lngRow1, COL_ҽ������) & """����Ϊһ����ҩ��"
                    Exit Function
                End If
            End If
        End If
        
        '��鴦����������
        If gintRXCount > 0 Then
            lngFind = .FindRow(.TextMatrix(lngRow1, COL_���ID), , COL_���ID)
            lngRxCount = GetMergeCount(lngFind)
            If lngRxCount >= gintRXCount Then
                strMsg = "һ����ҩ���� " & lngRxCount & " ���Ѵﵽ�򳬹�ҩƷ���������������� " & gintRXCount & " ����"
                Exit Function
            End If
        End If
    End With
    RowCanMerge = True
End Function

Public Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim lngҽ��ID As Long, lng���ID As Long
    Dim str��� As String, lngTmp As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim lngPreRow As Long, strMsg As String
    Dim lng������ĿID As Long, i As Long, j As Long
    
    Dim lng����ID As Long, str�Һŵ� As String, blnMoved As Boolean
    
    Call AdviceChange 'ǿ�Ƹ���ҽ������
    
    With vsAdvice
        Select Case Button.Key
            Case "����"
                If .RowData(.Row) = 0 Then
'                    If .Row <> .Rows - 1 Then
'                        MsgBox "��ǰ�������ݣ������ڵ�ǰ��¼����Чҽ����ɾ����ǰ�С�", vbInformation, gstrSysName
'                    Else
'                        MsgBox "��ǰ�������ݣ������ڵ�ǰ��¼����Чҽ����", vbInformation, gstrSysName
'                    End If
'                    Exit Sub
                ElseIf .RowData(.Rows - 1) = 0 Then
                    .Row = .Rows - 1
                Else
                    '��ɾ���м����Ŀ���
                    mblnRowChange = False
                    For i = .Rows - 1 To .FixedRows Step -1
                        If .RowData(i) = 0 Then .RemoveItem i
                    Next
                    mblnRowChange = True
                    
                    .AddItem "", .Rows
                    .Row = .Rows - 1
                    .Col = .FixedCols
                End If
                Call .ShowCell(.Row, .Col)
                If Visible And txtҽ������.Enabled Then txtҽ������.SetFocus
            Case "����"
                If .RowData(.Row) = 0 Then
                    MsgBox "��ǰ�������ݣ������ڵ�ǰ��¼����Чҽ����", vbInformation, gstrSysName
                    Exit Sub
                End If
                            
                lngPreRow = GetPreRow(.Row)
                            
                '�������Զ���Ϊһ����ҩ:������һ����ҩ���м����
                If lngPreRow <> -1 Then
                    If Val(.TextMatrix(lngPreRow, COL_���ID)) = Val(.TextMatrix(.Row, COL_���ID)) _
                        And Val(.TextMatrix(lngPreRow, COL_���ID)) <> 0 And InStr(",5,6,", .TextMatrix(.Row, COL_���)) > 0 Then
                        
                        '�������ѷ��͵�һ����ҩ�в���
                        If Val(.TextMatrix(.Row, COL_״̬)) <> 1 Then
                            MsgBox "����һ����ҩ��ҽ���Ѿ����ͣ������ٲ��롣", vbInformation, gstrSysName
                            Exit Sub
                        End If
                        '��������ǩ����һ����ҩ�в���
                        If Val(.TextMatrix(.Row, COL_ǩ����)) = 1 Then
                            MsgBox "����һ����ҩ��ҽ���Ѿ�ǩ���������ٲ��롣", vbInformation, gstrSysName
                            Exit Sub
                        End If
                        
                        lng���ID = Val(.TextMatrix(lngPreRow, COL_���ID))
                    End If
                End If
                
                '��ɾ���м����Ŀ���
                mblnRowChange = False
                lngҽ��ID = .RowData(.Row)
                For i = .Rows - 1 To .FixedRows Step -1
                    If .RowData(i) = 0 Then .RemoveItem i
                Next
                .Row = .FindRow(lngҽ��ID)
                mblnRowChange = True
                            
                '��ǰ��֮ǰ��������
                '--------------------------------------------------------------
                If RowIn�䷽��(.Row) Or RowIn������(.Row) Then
                    '��ҩ�䷽�������������ǰ���������
                    lngBegin = .FindRow(CStr(.RowData(.Row)), , COL_���ID)
                Else
                    lngBegin = .Row
                End If
                
                mblnRowChange = False
                .AddItem "", lngBegin
                .Row = lngBegin
                .Col = .FixedCols
                mblnRowChange = True
                Call vsAdvice_AfterRowColChange(-1, .Col, .Row, .Col)
                Call .ShowCell(.Row, .Col)
                
                txtҽ������.SetFocus '�ȶ�λ�������
            Case "һ��" 'һ����ҩ
                If Button.Value = tbrPressed Then
                    lngBegin = GetPreRow(.Row)
                    'ǰ��û����
                    If lngBegin = -1 Then
                        MsgBox "ǰ��û�п���һ����ҩ��ҽ���С�", vbInformation, gstrSysName
                        Button.Value = tbrUnpressed: Exit Sub
                    End If
                    '���в���������
                    If Not RowCanMerge(lngBegin, .Row, strMsg) Then
                        MsgBox strMsg, vbInformation, gstrSysName
                        Button.Value = tbrUnpressed: Exit Sub
                    End If
                    If .RowData(.Row) = 0 Then
                        '��ǰ����δ�������ݵ����
                        If DateDiff("n", CDate(.Cell(flexcpData, lngBegin, COL_��ʼʱ��)), zlDatabase.Currentdate) <= TIME_LIMIT Then
                            txt��ʼʱ��.Text = .Cell(flexcpData, lngBegin, COL_��ʼʱ��)
                        End If
                        txtҽ������.SetFocus: Exit Sub
                    Else
                        'Ҫ�ѵ�ǰ����ǰ����һ��һ����ҩ
                        Call MergeRow(lngBegin, .Row, False)
                        Call ReSetColor(.Row) 'һ��֮����һ������
                    End If
                Else
                    If .RowData(.Row) = 0 Then
                        '��ǰ����δ�������ݵ����
                        If RowInһ����ҩ(.Row) Then Button.Value = tbrPressed
                        Exit Sub
                    Else
                        '��ǰ����һ����ҩ�е���
                        Call Getһ����ҩ��Χ(Val(.TextMatrix(.Row, COL_���ID)), lngBegin, lngEnd)
                                                
                        '���жϿɷ�ȡ��һ����ҩ
                        '���ܰ����ѷ��͵�ҽ��
                        If Val(.TextMatrix(.Row, COL_״̬)) <> 1 Then
                            MsgBox "��ǰҽ���Ѿ����͡�", vbInformation, gstrSysName
                            Button.Value = tbrPressed: Exit Sub
                        End If
                        '���ܰ�����ǩ����ҽ��
                        If Val(.TextMatrix(.Row, COL_ǩ����)) = 1 Then
                            MsgBox "��ǰҽ���Ѿ�ǩ����", vbInformation, gstrSysName
                            Button.Value = tbrPressed: Exit Sub
                        End If
                                                
                        '����ʾ
                        If Not (.Row = lngEnd And lngEnd - lngBegin > 1) Then
                            '����һ����ҩȡ��Ϊ������ҩ
                            If MsgBox("Ҫ������һ����ҩ��ҩƷȫ��ȡ��Ϊ������ҩ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                Button.Value = tbrPressed: Exit Sub
                            End If
                        End If
                        
                        'ɾ���м�Ŀ���
                        lngTmp = .RowData(.Row)
                        For i = lngEnd To lngBegin Step -1
                            If .RowData(i) = 0 Then
                                .RemoveItem i
                                lngEnd = lngEnd - 1
                            End If
                        Next
                        .Row = .FindRow(lngTmp, lngBegin)
                        
                        If .Row = lngEnd And lngEnd - lngBegin > 1 Then
                            '��һ����ҩ�з������
                            Call ReSetColor(.Row) '��ȡ��֮ǰһ������
                            Call SplitRow(.Row)
                        Else
                            'ȡ��һ����ҩ
                            Call ReSetColor(.Row) '��ȡ��֮ǰһ������
                            lngTmp = .RowData(.Row) '��¼���ڻָ��ж�λ
                            Call AdviceSet������ҩ(lngBegin, lngEnd)
                            .Row = .FindRow(lngTmp)
                        End If
                    End If
                End If
                Call vsAdvice_AfterRowColChange(-1, .Col, .Row, .Col)
            Case "ɾ��"
                If .RowSel <> .Row Then
                    MsgBox "һ��ֻ��ɾ��һ��ҽ������ѡ��Ҫɾ����ҽ���С�", vbInformation, gstrSysName
                    Exit Sub
                End If
                If .RowData(.Row) <> 0 Then
                    '�ѷ��͵�ҽ������ɾ��
                    If Val(.TextMatrix(.Row, COL_״̬)) <> 1 Then
                        MsgBox "����ҽ���Ѿ����ͣ�����ɾ����", vbInformation, gstrSysName
                        Exit Sub
                    End If
                    '��ǩ����ҽ������ɾ��
                    If Val(.TextMatrix(.Row, COL_ǩ����)) = 1 Then
                        MsgBox "����ҽ���Ѿ�ǩ��������ɾ����", vbInformation, gstrSysName
                        Exit Sub
                    End If
                    
                    If MsgBox("ȷʵҪɾ��ҽ��""" & .TextMatrix(.Row, COL_ҽ������) & """��", _
                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                End If
                
                Call AdviceDelete(.Row) 'ɾ����ǰ��
                Call CalcAdviceMoney '��ʾ�¿�ҽ�����
                
                vsAdvice.SetFocus
            Case "�ο�"
                If Val(.TextMatrix(.Row, COL_������ĿID)) <> 0 Then
                    If RowIn�䷽��(.Row) Or RowIn������(.Row) Then
                        i = .FindRow(CStr(.RowData(.Row)), , COL_���ID)
                        If i <> -1 Then
                            lng������ĿID = Val(.TextMatrix(i, COL_������ĿID))
                        End If
                    Else
                        lng������ĿID = Val(.TextMatrix(.Row, COL_������ĿID))
                    End If
                End If
                Call ShowClinicHelp(IIF(mblnModal, 1, 0), Me, lng������ĿID)
            Case "����"
                lng����ID = mlng����ID: str�Һŵ� = mstr�Һŵ�: blnMoved = False
                strMsg = frmAdviceCopy.ShowMe(Me, mstrPrivs, lng����ID, str�Һŵ�, blnMoved, False, mlngǰ��ID)
                If strMsg <> "" Then
                    Call tbr_ButtonClick(tbr.Buttons("����"))
                    Call AdviceSet����ҽ��(lng����ID, str�Һŵ�, strMsg, blnMoved)
                End If
            Case "����"
                Call frmAdviceScheme.ShowMe(mstrPrivs, 1, mlng����ID, 0, mstr�Һŵ�, cboӤ��.ListIndex, Me)
            Case "����"
                If Not CheckAdvice Then Exit Sub '����д����˹�궨λ
                If Not SaveAdvice Then .SetFocus: Exit Sub
            Case "����"
                '����֮ǰ�Զ�����
                If mblnNoSave Then
                    If Not CheckAdvice Then Exit Sub
                    If Not SaveAdvice Then .SetFocus: Exit Sub
                End If
                If frmOutAdviceSend.ShowMe(Me, mstrPrivs, mlng����ID, mstr�Һŵ�, mlngǰ��ID, True) Then
                    '���¶�ȡ��ʾҽ��
                    Call ReLoadAdvice
                    mblnOK = True 'ǿ��
                    If txtҽ������.Enabled Then
                        txtҽ������.SetFocus
                    Else
                        .SetFocus
                    End If
                End If
            Case "ǩ��"
                Call AdviceSign
            Case "����"
                ShowHelp App.ProductName, Me.Hwnd, Me.Name
            Case "�˳�"
                Unload Me
        End Select
    End With
End Sub

Private Sub Getһ����ҩ��Χ(ByVal lng���ID As Long, lngBegin As Long, lngEnd As Long)
'���ܣ�������صĸ�ҩ;��ҽ��ID,ȷ��һ����ҩ��һ��ҩƷ����ֹ�к�
'˵�����м���ܰ����п���
    Dim i As Long
    lngBegin = vsAdvice.FindRow(CStr(lng���ID), , COL_���ID)
    For i = lngBegin To vsAdvice.Rows - 1
        If Not vsAdvice.RowHidden(i) And vsAdvice.RowData(i) <> 0 Then
            If Val(vsAdvice.TextMatrix(i, COL_���ID)) = lng���ID Then
                lngEnd = i
            Else
                Exit For
            End If
        End If
    Next
End Sub

Private Sub txt����_Change()
    txt����.Tag = "1"
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If IsNumeric(txt����.Text) Or txt����.Text = "" Then
            If SeekNextControl Then Call txt����_Validate(False)
        End If
    Else
        If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt����_Validate(Cancel As Boolean)
    Dim strMsg As String, dbl���� As Double, sng���� As Single
    
    If MousePressButton(tbr.Hwnd, tbr.Buttons("�˳�")) Then Exit Sub
    
    With vsAdvice
        If Val(txt����.Text) = 0 Then txt����.Text = ""
        If Not IsNumeric(txt����.Text) Then
            If txt����.Text <> "" Then
                Cancel = True: txt����_GotFocus: Exit Sub
            End If
        ElseIf CDbl(txt����.Text) <= 0 Then
            Cancel = True: txt����_GotFocus: Exit Sub
        ElseIf CDbl(txt����.Text) > LONG_MAX Then
            Cancel = True: txt����_GotFocus: Exit Sub
        Else
            '�����Ϸ��Լ��
            If InStr(",5,6,", .TextMatrix(.Row, COL_���)) > 0 And Val(.TextMatrix(.Row, COL_�շ�ϸĿID)) <> 0 Then
                dbl���� = IIF(Val(.TextMatrix(.Row, COL_����)) = 0, 1, Val(.TextMatrix(.Row, COL_����))) * _
                    Val(.TextMatrix(.Row, COL_�����װ)) * Val(.TextMatrix(.Row, COL_����ϵ��)) / Val(txt����.Text)
                If dbl���� > 200 Then
                    If MsgBox("��ҩƷ��ÿ�� " & FormatEx(txt����.Text, 5) & .TextMatrix(.Row, COL_������λ) & " ʹ�ã�" & _
                        IIF(Val(.TextMatrix(.Row, COL_����)) = 0, "ÿ", Val(.TextMatrix(.Row, COL_����))) & _
                        .TextMatrix(.Row, COL_���ﵥλ) & "����ʹ�� " & FormatEx(dbl����, 5) & " �Ρ�" & _
                        vbCrLf & vbCrLf & "��ȷ�ϵ���������ȷ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Cancel = True: txt����_GotFocus: Exit Sub
                    End If
                End If
            End If
            
            txt����.Text = FormatEx(txt����.Text, 5)
            
            '���¼���ҩƷ����(�����뵥��ʱ)
            If mbln���� And InStr(",5,6,", .TextMatrix(.Row, COL_���)) > 0 Then
                If .TextMatrix(.Row, COL_Ƶ��) <> "" _
                    And Val(.TextMatrix(.Row, COL_����ϵ��)) <> 0 _
                    And Val(.TextMatrix(.Row, COL_�����װ)) <> 0 Then
                    
                    sng���� = Val(.TextMatrix(.Row, COL_����))
                    If sng���� = 0 Then sng���� = 1
                    
                    txt����.Text = FormatEx(CalcȱʡҩƷ����( _
                        Val(txt����.Text), sng����, _
                        Val(.TextMatrix(.Row, COL_Ƶ�ʴ���)), Val(.TextMatrix(.Row, COL_Ƶ�ʼ��)), _
                        .TextMatrix(.Row, COL_�����λ), .TextMatrix(.Row, COL_ִ��ʱ��), _
                        Val(.TextMatrix(.Row, COL_����ϵ��)), Val(.TextMatrix(.Row, COL_�����װ)), _
                        Val(.TextMatrix(.Row, COL_�ɷ����))), 5)
                    txt����.Tag = "1"
                End If
            End If
        End If
        
        '��������
        Call AdviceChange
    End With
End Sub

Private Sub txt��ʼʱ��_Change()
    txt��ʼʱ��.Tag = "1"
End Sub

Private Sub txt��ʼʱ��_GotFocus()
    If txt��ʼʱ��.Text = "" Then txt��ʼʱ��.Text = GetDefaultTime(vsAdvice.Row)
    zlControl.TxtSelAll txt��ʼʱ��
End Sub

Private Sub txt��ʼʱ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt��ʼʱ��.Text <> "" Then
            txt��ʼʱ��.Text = GetFullDate(txt��ʼʱ��.Text)
            If SeekNextControl Then Call txt��ʼʱ��_Validate(False)
        End If
    Else
        If InStr("0123456789 /-:" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt��ʼʱ��_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And txt��ʼʱ��.Locked Then
        glngTXTProc = GetWindowLong(txt��ʼʱ��.Hwnd, GWL_WNDPROC)
        Call SetWindowLong(txt��ʼʱ��.Hwnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt��ʼʱ��_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And txt��ʼʱ��.Locked Then
        Call SetWindowLong(txt��ʼʱ��.Hwnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt��ʼʱ��_Validate(Cancel As Boolean)
    If MousePressButton(tbr.Hwnd, tbr.Buttons("�˳�")) Then Exit Sub
    
    If txt��ʼʱ��.Locked Then Exit Sub
        
    If Not IsDate(txt��ʼʱ��.Text) Then
        If txt��ʼʱ��.Text <> "" Then
            Cancel = True
            txt��ʼʱ��_GotFocus
            Exit Sub
        ElseIf vsAdvice.RowData(vsAdvice.Row) <> 0 Then
            If IsDate(vsAdvice.Cell(flexcpData, vsAdvice.Row, COL_��ʼʱ��)) Then
                '�ָ���Ϊ�����
                txt��ʼʱ��.Text = vsAdvice.Cell(flexcpData, vsAdvice.Row, COL_��ʼʱ��)
            End If
        End If
    Else
        '���ʱ��Ϸ���
        If Not Check��ʼʱ��(txt��ʼʱ��.Text) Then
            Cancel = True
            txt��ʼʱ��_GotFocus
            Exit Sub
        End If
    End If
    
    '��������
    Call AdviceChange
End Sub

Private Sub cboҽ������_Change()
    cboҽ������.Tag = "1"
End Sub

Private Sub cboҽ������_Click()
    cboҽ������.Tag = "1"
    Call AdviceChange
End Sub

Private Sub cboҽ������_GotFocus()
    zlControl.TxtSelAll cboҽ������
End Sub

Private Sub cboҽ������_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If SeekNextControl Then Call cboҽ������_Validate(False)
    Else
        Call zlControl.CboAppendText(cboҽ������, KeyAscii)
    End If
End Sub

Private Sub cboҽ������_Validate(Cancel As Boolean)
    If MousePressButton(tbr.Hwnd, tbr.Buttons("�˳�")) Then Exit Sub
    
    If zlCommFun.ActualLen(cboҽ������.Text) > 100 Then
        MsgBox "�������ݲ������� 50 �����ֻ� 100 ���ַ���", vbInformation, gstrSysName
        cboҽ������_GotFocus
        Cancel = True: Exit Sub
    End If
    
    '��������
    Call AdviceChange
End Sub

Private Sub txtҽ������_DblClick()
    If cmdExt.Visible And cmdExt.Enabled Then cmdExt_Click
End Sub

Private Sub txtҽ������_GotFocus()
    If txt��ʼʱ��.Text = "" Then txt��ʼʱ��_GotFocus
    Call zlControl.TxtSelAll(txtҽ������)
End Sub

Private Sub txtҽ������_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask And KeyCode = vbKeyA Then
        Call zlControl.TxtSelAll(txtҽ������)
    End If
End Sub

Private Sub txtҽ������_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txtҽ������.Text = "" Then Exit Sub
        If txtҽ������.Text = vsAdvice.TextMatrix(vsAdvice.Row, COL_ҽ������) Then
            Call SeekNextControl
            Exit Sub
        End If
        
        Set rsTmp = frmClinicSelect.ShowSelect(Me, mstrPrivs, 1, mstr�Ա�, txtҽ������.Text, txtҽ������, 1)
        If rsTmp Is Nothing Then 'ȡ����������
            '�ָ�ԭֵ
            txtҽ������.Text = vsAdvice.TextMatrix(vsAdvice.Row, COL_ҽ������)
            zlControl.TxtSelAll txtҽ������
            txtҽ������.SetFocus: Exit Sub
        End If
        '����Ŀ��¼��
        '������Ŀ�����������ҩ,���ܰ������ҽ��
        
        '����ѡ����Ŀ����ȱʡҽ����Ϣ
        Me.Refresh
        If AdviceInput(rsTmp, vsAdvice.Row) Then
            '��ʾ��ȱʡ���õ�ֵ
            Call vsAdvice_AfterRowColChange(-1, vsAdvice.Col, vsAdvice.Row, vsAdvice.Col)
            Call CalcAdviceMoney '��ʾ�¿�ҽ�����
            
            Call SeekNextControl
        Else
            '�ָ�ԭֵ
            txtҽ������.Text = vsAdvice.TextMatrix(vsAdvice.Row, COL_ҽ������)
            zlControl.TxtSelAll txtҽ������
            txtҽ������.SetFocus: Exit Sub
        End If
    ElseIf KeyAscii = Asc("*") Then
        KeyAscii = 0
        If cmdSel.Visible And cmdSel.Enabled Then Call cmdSel_Click
    End If
End Sub

Private Sub cboִ��ʱ��_GotFocus()
    zlControl.TxtSelAll cboִ��ʱ��
End Sub

Private Sub txtҽ������_Validate(Cancel As Boolean)
    '�ָ���Ϊ�ĸı�
    If txtҽ������.Text <> vsAdvice.TextMatrix(vsAdvice.Row, COL_ҽ������) Then
        txtҽ������.Text = vsAdvice.TextMatrix(vsAdvice.Row, COL_ҽ������)
    End If
End Sub

Private Sub txt����_Change()
    txt����.Tag = "1"
End Sub

Private Sub txt����_GotFocus()
    zlControl.TxtSelAll txt����
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Dim strMask As String
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If IsNumeric(txt����.Text) Then
            If SeekNextControl Then Call txt����_Validate(False)
        End If
    Else
        If RowIn�䷽��(vsAdvice.Row) Then
            strMask = "0123456789" '��ҩ�䷽ֻ����������
        ElseIf InStr(",5,6,", vsAdvice.TextMatrix(vsAdvice.Row, COL_���)) > 0 Then
            If InStr(mstrPrivs, "ҩƷС������") > 0 Then
                strMask = "0123456789."
            Else
                strMask = "0123456789"
            End If
        Else
            strMask = "0123456789."
        End If
        If InStr(strMask & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt����_Validate(Cancel As Boolean)
    Dim blnTag As Boolean, strMsg As String, bln�䷽�� As Boolean
    Dim dbl���� As Double, sng���� As Single
    
    If MousePressButton(tbr.Hwnd, tbr.Buttons("�˳�")) Then Exit Sub
    
    With vsAdvice
        If Val(txt����.Text) = 0 Then txt����.Text = ""
        If Not IsNumeric(txt����.Text) Then
            If txt����.Text <> "" Then
                Cancel = True: txt����_GotFocus: Exit Sub
            ElseIf .RowData(.Row) <> 0 Then
                '�ָ���Ϊ�����
                If IsNumeric(.TextMatrix(.Row, COL_����)) Then
                    txt����.Text = .TextMatrix(.Row, COL_����)
                End If
            End If
        ElseIf CDbl(txt����.Text) <= 0 Then
            Cancel = True: txt����_GotFocus: Exit Sub
        ElseIf CDbl(txt����.Text) > LONG_MAX Then
            Cancel = True: txt����_GotFocus: Exit Sub
        Else
            txt����.Text = FormatEx(txt����.Text, 5)
        End If
        
        bln�䷽�� = RowIn�䷽��(.Row)
        If IsNumeric(txt����.Text) Then
            If bln�䷽�� Then
                txt����.Text = CInt(txt����.Text)
            ElseIf InStr(",5,6,", .TextMatrix(.Row, COL_���)) > 0 Then
                If InStr(mstrPrivs, "ҩƷС������") = 0 Then
                    txt����.Text = Int(txt����.Text)
                End If
            ElseIf Val(.TextMatrix(.Row, COL_���㷽ʽ)) = 3 Then
                '�ƴ���Ŀ��������Ϊ�������ƴ���Ŀ�����뵥��,��˵�������
                'txt����.Text = Int(txt����.Text)
            End If
        End If
        
        '�����������
        If InStr(",5,6,", .TextMatrix(.Row, COL_���)) > 0 Then
            If .TextMatrix(.Row, COL_Ƶ��) <> "" _
                And Val(.TextMatrix(.Row, COL_����)) <> 0 _
                And Val(.TextMatrix(.Row, COL_����ϵ��)) <> 0 _
                And Val(.TextMatrix(.Row, COL_�����װ)) <> 0 Then
                
                sng���� = Val(.TextMatrix(.Row, COL_����))
                If sng���� = 0 Then sng���� = 1
                
                dbl���� = FormatEx(CalcȱʡҩƷ����( _
                    Val(.TextMatrix(.Row, COL_����)), sng����, _
                    Val(.TextMatrix(.Row, COL_Ƶ�ʴ���)), Val(.TextMatrix(.Row, COL_Ƶ�ʼ��)), _
                    .TextMatrix(.Row, COL_�����λ), .TextMatrix(.Row, COL_ִ��ʱ��), _
                    Val(.TextMatrix(.Row, COL_����ϵ��)), Val(.TextMatrix(.Row, COL_�����װ)), _
                    Val(.TextMatrix(.Row, COL_�ɷ����))), 5)
                    
                If Val(txt����.Text) < dbl���� Then
                    If MsgBox(.TextMatrix(.Row, COL_����) & "��ÿ�� " & _
                        .TextMatrix(.Row, COL_����) & .TextMatrix(.Row, COL_������λ) & "," & _
                        .TextMatrix(.Row, COL_Ƶ��) & IIF(mbln����, ",��ҩ " & sng���� & " ��", "") & _
                        "ִ��ʱ,������Ҫ " & FormatEx(dbl����, 5) & .TextMatrix(.Row, COL_������λ) & ",Ҫ������", _
                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Cancel = True: txt����_GotFocus: Exit Sub
                    End If
                End If
            End If
        End If
        
        '��鴦������
        If InStr(",5,6,", .TextMatrix(.Row, COL_���)) > 0 Then
            If Val(.TextMatrix(.Row, COL_��������)) <> 0 Then
                dbl���� = Val(txt����.Text) * Val(.TextMatrix(.Row, COL_�����װ)) * Val(.TextMatrix(.Row, COL_����ϵ��))
                If dbl���� > Val(.TextMatrix(.Row, COL_��������)) Then
                    If MsgBox(.TextMatrix(.Row, COL_����) & " ��������:" & txt����.Text & lbl������λ.Caption & "(" & dbl���� & lbl������λ.Caption & ")������������:" & _
                        FormatEx(Val(.TextMatrix(.Row, COL_��������)), 5) & lbl������λ.Caption & "��Ҫ������", _
                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Cancel = True: txt����_GotFocus: Exit Sub
                    End If
                End If
            End If
        ElseIf RowIn�䷽��(.Row) Then
            If Not CheckCHLimited(.Row, Val(txt����.Text)) Then
                Cancel = True: txt����_GotFocus: Exit Sub
            End If
        ElseIf InStr(",5,6,7,", .TextMatrix(.Row, COL_���)) = 0 And Val(.TextMatrix(.Row, COL_��������)) > 0 Then
            If Val(txt����.Text) > Val(.TextMatrix(.Row, COL_��������)) Then
                If MsgBox(.TextMatrix(.Row, COL_����) & " ������:" & txt����.Text & lbl������λ.Caption & " ��������¼����������:" & _
                    .TextMatrix(.Row, COL_��������) & lbl������λ.Caption & "��Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Cancel = True: txt����_GotFocus: Exit Sub
                End If
            End If
        End If
        
        '��������
        blnTag = txt����.Tag <> ""
        Call AdviceChange
        
        Call CalcAdviceMoney '��ʾ�¿�ҽ�����
        
        'ҩƷ�����:ֻ����,�޸��˲�����
        If blnTag Then
            If InStr(",5,6,", .TextMatrix(.Row, COL_���)) > 0 Or bln�䷽�� Then
                strMsg = CheckStock(.Row)
                If strMsg <> "" Then MsgBox strMsg, vbInformation, gstrSysName
            End If
        End If
    End With
End Sub

Private Function CheckCHLimited(ByVal lngRow As Long, ByVal int���� As Integer) As Boolean
'���ܣ������ҩ�䷽ÿζҩ�Ĵ�������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    CheckCHLimited = True
    
    With vsAdvice
        For i = lngRow - 1 To .FixedRows Step -1
            If Val(.TextMatrix(i, COL_���ID)) = .RowData(lngRow) Then
                If .TextMatrix(i, COL_���) = "7" Then
                    strSQL = strSQL & " Union ALL " & _
                        "Select ID,����,���㵥λ," & FormatEx(Val(.TextMatrix(i, COL_����)), 5) & " as ���� From ������ĿĿ¼ Where ID=" & Val(.TextMatrix(i, COL_������ĿID))
                End If
            Else
                Exit For
            End If
        Next
    End With
    If strSQL = "" Then Exit Function
    strSQL = "Select A.ID,A.����,A.���㵥λ,A.����,B.�������� From (" & Mid(strSQL, 11) & ") A,ҩƷ���� B Where A.ID=B.ҩ��ID And Nvl(B.��������,0)<>0"
    
    On Error GoTo errH
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption) 'û��
            
    strSQL = ""
    For i = 1 To rsTmp.RecordCount
        If int���� * rsTmp!���� > rsTmp!�������� Then
            strSQL = strSQL & vbCrLf & rsTmp!���� & "������:" & FormatEx(rsTmp!����, 5) & Nvl(rsTmp!���㵥λ) & "," & int���� & "��;��������:" & FormatEx(rsTmp!��������, 5) & Nvl(rsTmp!���㵥λ) & vbTab
        End If
        rsTmp.MoveNext
    Next
    If strSQL <> "" Then
        If MsgBox("���䷽������ҩƷ��������������" & vbCrLf & strSQL & vbCrLf & vbCrLf & "Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            CheckCHLimited = False
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ClearAdviceCard()
'���ܣ����ҽ����ʾ��Ƭ��ص�����
'������bln��ʼʱ��=�Ƿ������ʼʱ��
    Call SetCardEditable(True)
    
    txt��ʼʱ��.Text = ""
    txtҽ������.Text = ""
    cboҽ������.Text = ""
    cboִ�п���.Clear
    cbo����ִ��.Clear
    chk����.Visible = True
    
    mblnDoCheck = False
    chk����.Value = 0
    mblnDoCheck = True
    
    cmdExt.Enabled = False
    Call SetDayState(-1, -1)
    Call SetItemEditable(-1, -1, -1, -1, -1, -1, -1, -1)
    Call SetStartTime(True)
    
    stbThis.Panels(3).Text = ""
    stbThis.Panels(4).Text = ""
End Sub

Private Sub SetCardEditable(ByVal Editable As Boolean)
'���ܣ�����ɫ��ʶ��ǰҽ���Ƿ���Ա༭
    Dim obj As Object
    
    For Each obj In Controls
        If InStr("Label;TextBox;ComboBox;CheckBox", TypeName(obj)) > 0 Then
            If Not obj.Container Is Nothing Then
                If obj.Container Is fraAdvice Then
                    If Editable Then
                        obj.ForeColor = Me.ForeColor
                    Else
                        obj.ForeColor = &H808080
                    End If
                End If
            End If
        End If
    Next
    fraAdvice.Enabled = Editable
End Sub

Private Function GetƵ�ʷ�Χ(ByVal lngRow As Long) As Integer
    Dim lngFind As Long
    
    With vsAdvice
        If RowIn�䷽��(lngRow) Then
            GetƵ�ʷ�Χ = 2 '��ҽ
        Else
            If RowIn������(lngRow) Then '�Լ�����Ŀ��Ϊ׼
                lngFind = .FindRow(CStr(.RowData(lngRow)), , COL_���ID)
                If lngFind <> -1 Then lngRow = lngFind
            End If
            If InStr(",5,6,", .TextMatrix(lngRow, COL_���)) > 0 Or Val(.TextMatrix(lngRow, COL_Ƶ������)) = 0 Then
                GetƵ�ʷ�Χ = 1 '��ҩ���ѡƵ�ʵ���Ŀʹ����ҽƵ����Ŀ
            ElseIf Val(.TextMatrix(lngRow, COL_Ƶ������)) = 1 Then
                GetƵ�ʷ�Χ = -1 'һ����
            ElseIf Val(.TextMatrix(lngRow, COL_Ƶ������)) = 2 Then
                GetƵ�ʷ�Χ = -2 '������
            End If
        End If
    End With
End Function

Private Function SeekVisibleRow() As Boolean
'���ܣ���ǰ��Ϊ������ʱ����λ���������Ŀɼ���
    Dim lngRow As Long
    
    With vsAdvice
        If Not .RowHidden(.Row) Then Exit Function
        If InStr(",F,G,C,D,E,", .TextMatrix(.Row, COL_���)) > 0 And Val(.TextMatrix(.Row, COL_���ID)) <> 0 Then
            lngRow = .FindRow(CLng(Val(.TextMatrix(.Row, COL_���ID))))
        ElseIf .TextMatrix(.Row, COL_���) = "7" Then
            lngRow = .FindRow(CLng(Val(.TextMatrix(.Row, COL_���ID))))
        ElseIf .TextMatrix(.Row, COL_���) = "E" And Val(.TextMatrix(.Row, COL_���ID)) = 0 Then
            lngRow = .Row - 1
        End If
        If lngRow <> -1 Then
            If .RowData(lngRow) <> 0 Then
                .Row = lngRow: SeekVisibleRow = True
            End If
        End If
    End With
End Function

Private Sub vsAdvice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
'���ܣ����иı�ʱ�����¿�Ƭ����
    Dim rsItem As New ADODB.Recordset
    Dim strSQL As String, lngRow As Long
    Dim lng�÷�ID As Long, blnEditable As Boolean
    Dim lngBaseRow As Long, blnGroup As Boolean '��ҩ�䷽�ĵ�һζ���ҩ��
    Dim dblPrice As Double, strTmp As String, i As Long
    Dim lngҩƷID As Long
    
    If vsAdvice.Col >= vsAdvice.FixedCols Then
        vsAdvice.ForeColorSel = vsAdvice.Cell(flexcpForeColor, NewRow, COL_��ʼʱ��)
    End If
    
    If NewRow = OldRow Then Exit Sub
    If Not mblnRowChange Then Exit Sub
    If SeekVisibleRow Then Exit Sub
    
    Me.Refresh
    LockWindowUpdate Me.Hwnd
    
    lngRow = NewRow
    blnGroup = RowInһ����ҩ(lngRow) '����Ҳ������һ����ҩ�ķ�Χ��
    tbr.Buttons("һ��").Value = IIF(blnGroup, tbrPressed, tbrUnpressed)
        
    On Error GoTo errH
    
    With vsAdvice
        If .RowData(lngRow) = 0 Then
            '��Ч�������Ƭ����
            Call ClearAdviceCard
            
            'ȱʡ��ʼʱ��
            Call txt��ʼʱ��_GotFocus
        Else
            '��Ƭ�༭
            blnEditable = True
            '�ѷ��͵�ҽ�������޸�
            If Val(.TextMatrix(lngRow, COL_״̬)) <> 1 Then blnEditable = False
            '��ǩ����ҽ�������޸�
            If Val(.TextMatrix(lngRow, COL_ǩ����)) = 1 Then blnEditable = False
            Call SetCardEditable(blnEditable)
            
            '��ȡ������Ŀ������Ϣ
            '---------------------
            If InStr(",5,6,7,", Val(.TextMatrix(lngRow, COL_���))) > 0 Then
                lngҩƷID = Val(.TextMatrix(lngRow, COL_�շ�ϸĿID))
            End If
            
            If RowIn�䷽��(lngRow) Then
                txt����.MaxLength = 3
                '��ȡ��ҩ�䷽��һζ��ҩ��
                lngBaseRow = .FindRow(CStr(.RowData(lngRow)), , COL_���ID)
                lngҩƷID = Val(.TextMatrix(lngBaseRow, COL_�շ�ϸĿID))
            ElseIf RowIn������(lngRow) Then
                '��ȡһ�������ĵ�һ����Ŀ��
                lngBaseRow = .FindRow(CStr(.RowData(lngRow)), , COL_���ID)
                txt����.MaxLength = txt����.MaxLength
            Else
                lngBaseRow = lngRow
                txt����.MaxLength = txt����.MaxLength
            End If
            strSQL = "Select * From ������ĿĿ¼ Where ID=[1]"
            Set rsItem = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngBaseRow, COL_������ĿID)))
            
            '��չ��ť����״̬(������,�������,����,��ҩ�䷽)
            cmdExt.Enabled = InStr(",7,C,F,", rsItem!���) > 0 Or (rsItem!��� = "D" And Nvl(rsItem!�����Ŀ, 0) = 1)
            
            '��ʾ��ǰҽ����Ƭ����
            '--------------------------------------------------------------------------------------------
            '��ʼʱ�䣺ֻ������ҽ��ʱ�����޸Ŀ�ʼʱ��
            txt��ʼʱ��.Text = .Cell(flexcpData, lngRow, COL_��ʼʱ��)
            Call SetStartTime(.TextMatrix(lngRow, COL_EDIT) = "1")
            
            'ҽ������
            txtҽ������.Text = .TextMatrix(lngRow, COL_ҽ������)
            
            '����������,��ҩ���ѡ��Ƶ�ʵļ�ʱ,������Ŀ����¼��
            '----------------------
            If rsItem!��� = "7" Then '��ҩ�䷽(�в�ҩ)��Ȼ�е���,������������д
                SetItemEditable -1
            ElseIf (Nvl(rsItem!ִ��Ƶ��, 0) = 0 And InStr(",1,2,", Nvl(rsItem!���㷽ʽ, 0)) > 0) _
                    Or InStr(",5,6,", rsItem!���) > 0 Then
                SetItemEditable 1
                txt����.Text = .TextMatrix(lngRow, COL_����)
                lbl������λ.Caption = .TextMatrix(lngRow, COL_������λ)
            Else
                SetItemEditable -1
            End If
            
            '��������ҩ���г�ҩ������ʹ�ã����ڼ�������
            'һ�㣺������ҩƷ(����ҩ)���ѡ��Ƶ�ʵļ�ʱ,������Ŀ����ʹ���������Զ���������
            blnEditable = False
            If InStr(",5,6,", rsItem!���) > 0 Then
                If mbln���� Then blnEditable = True
            End If
            If blnEditable Then
                SetDayState 1, 1
            Else
                SetDayState -1, -1
            End If
            txt����.Text = Val(.TextMatrix(lngRow, COL_����))
            If Val(txt����.Text) = 0 Then txt����.Text = ""
            
            '����
            '--------------------
            If rsItem!��� = "7" Then
                '��ҩ�䷽(�в�ҩ)��дΪ����
                SetItemEditable , 1
                lbl������λ.Caption = "��"
                txt����.Text = .TextMatrix(lngRow, COL_����) '����
            Else
                '��������Ҫ��д����:��������������Ϊ׼
                If rsItem!��� = "Z" And Nvl(rsItem!��������) <> "0" Then
                    SetItemEditable , -1 '����ҽ���������޸�����(�̶�Ϊ1��)
                ElseIf Nvl(rsItem!ִ��Ƶ��, 0) = 1 And Nvl(rsItem!���㷽ʽ, 0) = 3 Then
                    SetItemEditable , -1 'һ���Լƴ���Ŀ����������
                Else
                    SetItemEditable , 1
                End If
                lbl������λ.Caption = .TextMatrix(lngRow, COL_������λ)
                txt����.Text = .TextMatrix(lngRow, COL_����)
            End If
            
            '��ҩ;������ҩ�÷�
            '--------------
            If InStr(",5,6,", rsItem!���) > 0 Then
                SetItemEditable , , 1
                lbl�÷�.Caption = "��ҩ;��"
                '���Ҹ�ҩ;����Ӧ����:���ҵ�Rowdata(Variant)����ҪתΪLong��,���ܾ�ȷƥ��
                lng�÷�ID = .FindRow(CLng(.TextMatrix(lngRow, COL_���ID)), lngRow + 1)
                lng�÷�ID = Val(.TextMatrix(lng�÷�ID, COL_������ĿID))
                cmd�÷�.Tag = lng�÷�ID
                txt�÷�.Text = Get��Ŀ����(lng�÷�ID)
            ElseIf rsItem!��� = "7" Then
                SetItemEditable , , 1
                lbl�÷�.Caption = "��ҩ�÷�"
                
                '��ҩ�䷽��ʾ�о�����ҩ�÷���
                lng�÷�ID = Val(.TextMatrix(lngRow, COL_������ĿID))
                cmd�÷�.Tag = lng�÷�ID
                txt�÷�.Text = Get��Ŀ����(lng�÷�ID)
            ElseIf RowIn������(lngRow) Then '��������ж�,������ǰ�ļ���
                '�������
                SetItemEditable , , 1
                lbl�÷�.Caption = "�ɼ�����"
                
                '���������ʾ�о��ǲɼ�������
                lng�÷�ID = Val(.TextMatrix(lngRow, COL_������ĿID))
                cmd�÷�.Tag = lng�÷�ID
                txt�÷�.Text = Get��Ŀ����(lng�÷�ID)
            Else
                SetItemEditable , , -1
            End If
            
            'Ƶ�ʣ�������ѡ��(������������ָ��ʹ��)
            If True Then
                SetItemEditable , , , 1
                cmdƵ��.Tag = .TextMatrix(lngRow, COL_Ƶ��)
                txtƵ��.Text = .TextMatrix(lngRow, COL_Ƶ��)
            Else
                SetItemEditable , , , -1
            End If
                    
            'ִ��ʱ�䣺"��ѡƵ��"��ҩƷ��
            If Nvl(rsItem!ִ��Ƶ��, 0) = 0 Or InStr(",5,6,7,", rsItem!���) > 0 Then
                SetItemEditable , , , , 1
                Call Getʱ�䷽��(cboִ��ʱ��, GetƵ�ʷ�Χ(lngRow), .TextMatrix(lngRow, COL_Ƶ��), lng�÷�ID)
                cboִ��ʱ��.Text = .TextMatrix(lngRow, COL_ִ��ʱ��)
            Else
                SetItemEditable , , , , -1
            End If
                    
            'ҽ������
            cboҽ������.Text = .TextMatrix(lngRow, COL_ҽ������)
                    
            'ִ������
            If InStr(",5,6,7,", rsItem!���) > 0 Then
                If rsItem!��� = "7" Then
                    '������ҩ�䷽,����������Ŀ���������Ƽ���������,�������÷��ͼ巨һ��ΪԺ��ִ��,һ����Ϊ
                    If Val(.TextMatrix(lngBaseRow, COL_ִ������)) = 5 And Val(.TextMatrix(lngRow, COL_ִ������)) <> 5 Then
                        strTmp = "�Ա�ҩ"
                    ElseIf Val(.TextMatrix(lngBaseRow, COL_ִ������)) <> 5 And Val(.TextMatrix(lngRow, COL_ִ������)) = 5 Then
                        strTmp = "��Ժ��ҩ"
                    Else
                        strTmp = "����"
                    End If
                Else
                    i = .FindRow(CLng(.TextMatrix(lngRow, COL_���ID)), lngRow + 1)
                    If Val(.TextMatrix(lngRow, COL_ִ������)) = 5 And Val(.TextMatrix(i, COL_ִ������)) <> 5 Then
                        strTmp = "�Ա�ҩ"
                    ElseIf Val(.TextMatrix(lngRow, COL_ִ������)) <> 5 And Val(.TextMatrix(i, COL_ִ������)) = 5 Then
                        strTmp = "��Ժ��ҩ"
                    Else
                        strTmp = "����"
                    End If
                End If
                SetItemEditable , , , , , , 1
                Call SeekIndex(cboִ������, strTmp)
            Else
                SetItemEditable , , , , , , -1
            End If
                    
            'ִ�п���:���ۻ�סԺҽ�����ٴ�����
            If rsItem!��� = "Z" And InStr(",1,2,", Nvl(rsItem!��������, 0)) > 0 Then
                SetItemEditable , , , , , 1
                If Nvl(rsItem!��������, 0) = 1 Then
                    lblִ�п���.Caption = "���ۿ���"
                    '����:���������סԺ�ٴ�����,�ɷ������������������ۻ�סԺ����
                    Call Get�ٴ�����(3, , Val(.TextMatrix(lngRow, COL_ִ�п���ID)), cboִ�п���, Not gbln�������Ҷ���)
                ElseIf Nvl(rsItem!��������, 0) = 2 Then
                    lblִ�п���.Caption = "סԺ����"
                    'סԺ:����סԺ�ٴ�����
                    Call Get�ٴ�����(2, , Val(.TextMatrix(lngRow, COL_ִ�п���ID)), cboִ�п���, Not gbln�������Ҷ���)
                End If
            Else
                '��ҩƷ����ҩƷ��Ϊ׼��ʾ,��������Լ�����ĿΪ׼��ʾ
                i = lngRow
                If rsItem!��� = "7" Then
                    i = lngBaseRow
                ElseIf RowIn������(lngRow) Then '��������ж�,������ǰ�ļ���
                    i = lngBaseRow
                End If
                
                If InStr(",0,5,", Val(.TextMatrix(i, COL_ִ������))) = 0 Then
                    '�Ƕ�����Ժ��ִ��ʱ����ʾ�Ϳ���ѡ��(����ҩƷ)
                    SetItemEditable , , , , , 1
                    Call Get����ִ�п���(mlng����ID, 0, cboִ�п���, rsItem!���, rsItem!ID, lngҩƷID, Nvl(rsItem!ִ�п���, 0), _
                        mlng���˿���id, Val(.TextMatrix(i, COL_��������ID)), Val(.TextMatrix(i, COL_ִ�п���ID)), 1, 1)
                ElseIf InStr(",0,5,", Val(.TextMatrix(i, COL_ִ������))) > 0 Then
                    SetItemEditable , , , , , -1
                    If Val(.TextMatrix(i, COL_ִ������)) = 0 Then
                        cboִ�п���.AddItem "<��ִ�ж���>"
                    Else
                        cboִ�п���.AddItem "<Ժ��ִ��>"
                    End If
                    Call zlControl.CboSetIndex(cboִ�п���.Hwnd, 0)
                End If
            End If
            
            '����ִ��:ָ��ҩ;��,��ҩ�÷�,��������,�ɼ���ʽ��ִ�п���
            If Should����ִ��(lngRow, i, strTmp) Then
                SetItemEditable , , , , , , , 1
                Call Get����ִ�п���(mlng����ID, 0, cbo����ִ��, .TextMatrix(i, COL_���), Val(.TextMatrix(i, COL_������ĿID)), lngҩƷID, _
                    Val(.TextMatrix(i, COL_ִ������)), mlng���˿���id, Val(.TextMatrix(i, COL_��������ID)), Val(.TextMatrix(i, COL_ִ�п���ID)), 1, 1)
            Else
                SetItemEditable , , , , , , , -1
                If i <> -1 Then
                    If InStr(",0,5,", Val(.TextMatrix(i, COL_ִ������))) > 0 Then
                        If Val(.TextMatrix(i, COL_ִ������)) = 0 Then
                            cbo����ִ��.AddItem "<��ִ�ж���>"
                        ElseIf Val(.TextMatrix(i, COL_ִ������)) = 5 Then
                            cbo����ִ��.AddItem "<Ժ��ִ��>"
                        End If
                        Call zlControl.CboSetIndex(cbo����ִ��.Hwnd, 0)
                    End If
                End If
            End If
            lbl����ִ��.Caption = strTmp
            
            '������־
            chk����.Visible = True
            mblnDoCheck = False
            chk����.Value = Val(.TextMatrix(lngRow, COL_��־))
            mblnDoCheck = True
            
            '��ʾҩƷ��棺�����ﵥλ����ҩ�䷽����ʾ
            '----------------------------------------
            If InStr(",5,6,", rsItem!���) > 0 And Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)) <> 0 Then
                If .TextMatrix(lngRow, COL_���) = "" Then Call GetDrugStock(lngRow)
                If .TextMatrix(lngRow, COL_���) <> "" Then
                    stbThis.Panels(3).Text = "���:" & FormatEx(Val(.TextMatrix(lngRow, COL_���)), 5) & .TextMatrix(lngRow, COL_���ﵥλ)
                Else
                    stbThis.Panels(3).Text = ""
                End If
            Else
                If rsItem!��� = "7" And Val(.TextMatrix(lngRow, COL_״̬)) = 1 Then
                    Call GetDrugStock(lngRow)
                End If
                stbThis.Panels(3).Text = ""
            End If
            
            '��ʾҽ�����ۺͷ�������
            If .TextMatrix(lngRow, COL_����) = "" Then '��""������
                .TextMatrix(lngRow, COL_����) = GetItemPrice(lngRow)
            End If
            dblPrice = Val(.TextMatrix(lngRow, COL_����))
            If dblPrice <> 0 Then
                If InStr(",5,6,", rsItem!���) > 0 Then
                    stbThis.Panels(4).Text = "ÿ" & .TextMatrix(lngRow, COL_���ﵥλ) & ":" & FormatEx(dblPrice, 5) & "Ԫ"
                ElseIf rsItem!��� = "7" Then
                    stbThis.Panels(4).Text = "ÿ��:" & FormatEx(dblPrice, 5) & "Ԫ"
                Else
                    stbThis.Panels(4).Text = IIF(IsNull(rsItem!���㵥λ), "�۸�:", "ÿ" & Nvl(rsItem!���㵥λ) & ":") & FormatEx(dblPrice, 5) & "Ԫ"
                End If
            Else
                stbThis.Panels(4).Text = ""
            End If
            
            '��ʾ��������
            strTmp = Get��������(lngRow)
            If strTmp <> "" Then
                stbThis.Panels(4).Text = stbThis.Panels(4).Text & IIF(stbThis.Panels(4).Text = "", "����:", ",����:") & strTmp
            End If
        End If
    End With
    
    '����༭��־
    Call ClearItemTag
    
    '����ҽ�����ܿ�����
    Call SetFuncEnabled
    
    '��ʾ�Ƽ۴���
    Call ShowPrice(lngRow)
    
    LockWindowUpdate 0
    Exit Sub
errH:
    LockWindowUpdate 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ShowPrice(ByVal lngRow As Long)
'���ݵ�ǰ�е������ʾ�Ƽ۴���
    If mblnModal Then Exit Sub
    
    If vsAdvice.RowData(lngRow) = 0 Or Val(vsAdvice.TextMatrix(lngRow, COL_������ĿID)) = 0 Then
        stbThis.Panels("Price").Bevel = sbrNoBevel
        stbThis.Panels("Price").Visible = False
    ElseIf InStr(",1,2,", Val(vsAdvice.TextMatrix(lngRow, COL_״̬))) = 0 Then
        stbThis.Panels("Price").Bevel = sbrNoBevel
        stbThis.Panels("Price").Visible = False
    ElseIf InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_���)) > 0 Then
        stbThis.Panels("Price").Bevel = sbrNoBevel
        stbThis.Panels("Price").Visible = False
    ElseIf RowIn�䷽��(lngRow) Then
        stbThis.Panels("Price").Bevel = sbrNoBevel
        stbThis.Panels("Price").Visible = False
    ElseIf stbThis.Panels("Price").Bevel = sbrNoBevel Then
        stbThis.Panels("Price").Visible = True
        If stbThis.Panels("Price").Tag <> "" Then
            stbThis.Panels("Price").Bevel = sbrInset
        Else
            stbThis.Panels("Price").Bevel = sbrRaised
        End If
    End If
    
    If stbThis.Panels("Price").Bevel <> sbrInset Then
        '�رռƼ۴���
        mfrmPrice.HideMe
    Else
        Call mfrmPrice.ShowMe(Me, vsAdvice, mlng����ID, 0, mlng���˿���id, _
            COL_��� & "," & COL_���ID & "," & COL_״̬ & "," & COL_��� & "," & COL_������ĿID & "," & _
            COL_�շ�ϸĿID & "," & COL_�걾��λ & "," & COL_�Ƽ����� & "," & COL_ִ������ & "," & COL_ִ�п���ID)
    End If
End Sub

Private Sub SetFuncEnabled()
'���ܣ�����ҽ�����ܿ�����
    Dim blnEnabled As Boolean
    With vsAdvice
        'ɾ������
        blnEnabled = True
        If .RowData(.Row) <> 0 Then
            If Val(.TextMatrix(.Row, COL_״̬)) <> 1 Then blnEnabled = False
            '��ǩ��ҽ������ɾ��
            If Val(.TextMatrix(.Row, COL_ǩ����)) = 1 Then blnEnabled = False
        End If
        tbr.Buttons("ɾ��").Enabled = blnEnabled
    End With
End Sub

Private Function Get��������(ByVal lngRow As Long) As String
'���ܣ���ȡָ���еķ�������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, str���� As String
    
    With vsAdvice
        If InStr(",5,6,", .TextMatrix(lngRow, COL_���)) > 0 And Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)) <> 0 Then
            'ȡҽ���ķ�������
            If mint���� <> 0 Then
                str���� = gclsInsure.GetItemInsure(mlng����ID, Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)), 0, True, mint����)
                If str���� <> "" Then
                    If UBound(Split(str����, ";")) >= 5 Then
                        str���� = Split(str����, ";")(5)
                    Else
                        str���� = ""
                    End If
                End If
            End If
            'û����ȡHIS�ķ�������
            If str���� = "" Then
                strSQL = "Select �������� From �շ���ĿĿ¼ Where ID=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)))
                If Not rsTmp.EOF Then str���� = Nvl(rsTmp!��������)
            End If
        End If
    End With
    Get�������� = str����
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Should����ִ��(ByVal lngRow As Long, lngRow2 As Long, strִ�п��� As String) As Boolean
'���ܣ��ж�ָ����ҽ����(�ɼ���)�Ƿ�������ø��ӵ�ִ�п���
'������lngRow2=���ظ����е�ҽ���к�
'      strִ�п���=����ִ�п�������
    Dim i As Long
    
    lngRow2 = -1
    strִ�п��� = "����ִ��"
    With vsAdvice
        If lngRow = 0 Or .RowData(lngRow) = 0 Then Exit Function
        
        If RowIn�䷽��(lngRow) Then
            '��ҩ�÷�
            lngRow2 = lngRow
            strִ�п��� = "�÷�ִ��"
            Should����ִ�� = True
        ElseIf InStr(",5,6,", .TextMatrix(lngRow, COL_���)) > 0 Then
            '��ҩ;��
            lngRow2 = .FindRow(CLng(.TextMatrix(lngRow, COL_���ID)), lngRow + 1)
            strִ�п��� = "��ҩִ��"
            Should����ִ�� = True
        ElseIf .TextMatrix(lngRow, COL_���) = "F" Then
            '��������
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, COL_���ID)) = .RowData(lngRow) Then
                    If .TextMatrix(i, COL_���) = "G" Then
                        lngRow2 = i: Exit For
                    End If
                Else
                    Exit For
                End If
            Next
            strִ�п��� = "����ִ��"
            If lngRow2 <> -1 Then Should����ִ�� = True
        ElseIf .TextMatrix(lngRow, COL_���) = "E" _
            And .TextMatrix(lngRow - 1, COL_���) = "C" _
            And Val(.TextMatrix(lngRow - 1, COL_���ID)) = .RowData(lngRow) Then
            '�ɼ���ʽ
            lngRow2 = lngRow
            strִ�п��� = "�ɼ�ִ��"
            Should����ִ�� = True
        End If
        
        '������Ժ��ִ��
        If Should����ִ�� Then
            If InStr(",0,5,", Val(.TextMatrix(lngRow2, COL_ִ������))) > 0 Then
                Should����ִ�� = False
            End If
        End If
    End With
End Function

Private Function GetItemPrice(ByVal lngRow As Long) As Double
'���ܣ���ȡ��ǰҽ���еļ۸�(ҩƷΪһ��ҩ����װ�ĵ���,���������շѶ���)
'˵����ҩƷ��������ҩ;������ҩ�÷��巨
    Dim rsTmp As New ADODB.Recordset
    Dim strҽ��IDs As String, str��ĿIDs As String, str����s As String
    Dim strAdviceIDs As String, lngִ�п���ID As Long
    Dim dblPrice As Double, dbl���� As Double
    Dim blnҩƷ As Boolean, strSQL As String, i As Long
    
    With vsAdvice
        blnҩƷ = True
        If InStr(",5,6,", .TextMatrix(lngRow, COL_���)) > 0 And Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)) <> 0 Then
            '��ҩ���г�ҩ������²��ܼ���۸�
            If Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)) <> 0 Then
                str��ĿIDs = str��ĿIDs & "," & Val(.TextMatrix(lngRow, COL_�շ�ϸĿID))
            End If
            lngִ�п���ID = Val(.TextMatrix(lngRow, COL_ִ�п���ID))
        ElseIf RowIn�䷽��(lngRow) Then
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_���ID)) = .RowData(lngRow) Then
                    If .TextMatrix(i, COL_���) = "7" And Val(.TextMatrix(i, COL_�շ�ϸĿID)) <> 0 Then
                        If lngִ�п���ID = 0 Then
                            lngִ�п���ID = Val(.TextMatrix(i, COL_ִ�п���ID))
                        End If
                        str��ĿIDs = str��ĿIDs & "," & Val(.TextMatrix(i, COL_�շ�ϸĿID))
                        str����s = str����s & ";" & Val(.TextMatrix(i, COL_����))
                    End If
                Else
                    Exit For
                End If
            Next
        Else
            blnҩƷ = False
            '����ҽ��,δУ��(�Ƽ�)�İ��շѶ��ռ���,����ֱ��ȡҽ���Ƽ�
            '���������Ƽۺ��ֹ��Ƽ۵���Ŀ
            If Val(.TextMatrix(lngRow, COL_�Ƽ�����)) = 0 Then
                If InStr(",1,2,", .TextMatrix(lngRow, COL_״̬)) > 0 Then
                    str��ĿIDs = str��ĿIDs & "," & Val(.TextMatrix(lngRow, COL_������ĿID))
                Else
                    strҽ��IDs = strҽ��IDs & "," & .RowData(lngRow)
                End If
            End If
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, COL_���ID)) = .RowData(lngRow) Then
                    If Val(.TextMatrix(i, COL_�Ƽ�����)) = 0 Then
                        If InStr(",1,2,", .TextMatrix(i, COL_״̬)) > 0 Then
                            str��ĿIDs = str��ĿIDs & "," & Val(.TextMatrix(i, COL_������ĿID))
                        Else
                            strҽ��IDs = strҽ��IDs & "," & .RowData(i)
                        End If
                    End If
                Else
                    Exit For
                End If
            Next
            For i = lngRow - 1 To .FixedRows Step -1 '�������
                If Val(.TextMatrix(i, COL_���ID)) = .RowData(lngRow) Then
                    If Val(.TextMatrix(i, COL_�Ƽ�����)) = 0 Then
                        If InStr(",1,2,", .TextMatrix(i, COL_״̬)) > 0 Then
                            str��ĿIDs = str��ĿIDs & "," & Val(.TextMatrix(i, COL_������ĿID))
                        Else
                            strҽ��IDs = strҽ��IDs & "," & .RowData(i)
                        End If
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End With
    strҽ��IDs = Mid(strҽ��IDs, 2)
    str��ĿIDs = Mid(str��ĿIDs, 2)
    str����s = Mid(str����s, 2)
    
    On Error GoTo errH
    
    If blnҩƷ Then
        If str��ĿIDs = "" Then Exit Function
    
        '������ʱ,ID˳��Ϊ��������
        strSQL = "Select A.ID,A.���,A.�Ƿ���,B.�����װ,B.����ϵ��,B.�ɷ����" & _
            " From �շ���ĿĿ¼ A,ҩƷ��� B Where A.ID=B.ҩƷID And A.ID IN(" & str��ĿIDs & ")"
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption) 'In
        For i = 1 To rsTmp.RecordCount
            '����:�����װ
            If str����s <> "" Then '��ҩ�䷽�Ź�ÿζ����
                dbl���� = Val(Split(str����s, ";")(rsTmp.RecordCount - i))
                
                '��ҩҩ����λ�����ɷ��㴦��:ÿ��
                If Nvl(rsTmp!�ɷ����, 0) = 0 Then
                    dbl���� = Format(dbl���� / Nvl(rsTmp!����ϵ��, 1) / Nvl(rsTmp!�����װ, 1), "0.00000")
                Else
                    dbl���� = IntEx(dbl���� / Nvl(rsTmp!����ϵ��, 1) / Nvl(rsTmp!�����װ, 1))
                End If
            Else
                dbl���� = 1
            End If
            If Nvl(rsTmp!�Ƿ���, 0) = 0 Then
                dblPrice = dblPrice + CalcPrice(rsTmp!ID) * Nvl(rsTmp!�����װ, 1) * dbl����
            Else
                dblPrice = dblPrice + CalcDrugPrice(rsTmp!ID, lngִ�п���ID, dbl���� * Nvl(rsTmp!�����װ, 1)) * Nvl(rsTmp!�����װ, 1) * dbl����
            End If
            
            rsTmp.MoveNext
        Next
    Else
        If str��ĿIDs = "" And strҽ��IDs = "" Then Exit Function
    
        If strҽ��IDs <> "" Then
            strSQL = _
                " Select B.����,Decode(C.�Ƿ���,1,B.����,Sum(D.�ּ�)) as ����" & _
                " From ����ҽ���Ƽ� B,�շ���ĿĿ¼ C,�շѼ�Ŀ D" & _
                " Where B.�շ�ϸĿID=C.ID And B.�շ�ϸĿID=D.�շ�ϸĿID" & _
                " And ((Sysdate Between D.ִ������ And D.��ֹ����) Or (Sysdate>=D.ִ������ And D.��ֹ���� is NULL))" & _
                " And B.ҽ��ID IN(" & strҽ��IDs & ")" & _
                " Group by B.����,C.�Ƿ���,B.����"
        End If
        If str��ĿIDs <> "" Then
            strSQL = strSQL & IIF(strSQL = "", "", " Union ALL") & _
                " Select B.�շ����� as ����,Decode(C.�Ƿ���,1,0,Sum(D.�ּ�)) as ����" & _
                " From �����շѹ�ϵ B,�շ���ĿĿ¼ C,�շѼ�Ŀ D" & _
                " Where B.�շ���ĿID=C.ID And B.�շ���ĿID=D.�շ�ϸĿID" & _
                " And ((Sysdate Between D.ִ������ And D.��ֹ����) Or (Sysdate>=D.ִ������ And D.��ֹ���� is NULL))" & _
                " And B.������ĿID IN(" & str��ĿIDs & ")" & _
                " Group by B.�շ�����,C.�Ƿ���"
        End If
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Name) 'In
        For i = 1 To rsTmp.RecordCount
            dblPrice = dblPrice + Format(Nvl(rsTmp!����, 0) * Nvl(rsTmp!����, 0), "0.00000")
            rsTmp.MoveNext
        Next
    End If
    
    GetItemPrice = Format(dblPrice, "0.00000")
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub GetDrugStock(ByVal lngRow As Long)
'���ܣ����»�ȡָ��ҩƷ�е�ҩƷ���
'������lngRow=��ҩ�л���ҩ�÷���
'˵�����������ҩ�䷽��,һ���Ի�ȡ�����䷽�е�������ҩ�Ŀ��
    Dim i As Long
    
    With vsAdvice
        If InStr(",5,6,", .TextMatrix(lngRow, COL_���)) > 0 Then
            If Val(.TextMatrix(lngRow, COL_ִ�п���ID)) = 0 Or Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)) = 0 Then
                .TextMatrix(lngRow, COL_���) = ""
            Else
                .TextMatrix(lngRow, COL_���) = GetStock(Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)), Val(.TextMatrix(lngRow, COL_ִ�п���ID)), 1)
            End If
        ElseIf RowIn�䷽��(lngRow) Then
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_���ID)) = .RowData(lngRow) Then
                    If .TextMatrix(i, COL_���) = "7" Then
                        If Val(.TextMatrix(i, COL_ִ�п���ID)) = 0 Or Val(.TextMatrix(i, COL_�շ�ϸĿID)) = 0 Then
                            .TextMatrix(i, COL_���) = ""
                        Else
                            .TextMatrix(i, COL_���) = GetStock(Val(.TextMatrix(i, COL_�շ�ϸĿID)), Val(.TextMatrix(i, COL_ִ�п���ID)), 1)
                        End If
                    End If
                Else
                    Exit For
                End If
            Next
        End If
    End With
End Sub

Private Sub vsAdvice_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngW As Long
    
    If Row = -1 Then
        lngW = Me.TextWidth(vsAdvice.TextMatrix(0, Col) & "A")
        If vsAdvice.ColWidth(Col) < lngW Then
            vsAdvice.ColWidth(Col) = lngW
        ElseIf vsAdvice.ColWidth(Col) > vsAdvice.Width * 0.5 Then
            vsAdvice.ColWidth(Col) = vsAdvice.Width * 0.5
        End If
        
        If Col = COL_ҽ������ Then Call vsAdvice.AutoSize(COL_ҽ������)
    End If
End Sub

Private Sub vsAdvice_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)
    If dtpDate.Visible Then
        Call Form_KeyDown(vbKeyEscape, 0)
        Cancel = True
    End If
End Sub

Private Sub vsAdvice_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Row = -1 Then
        If Col <= vsAdvice.FixedCols - 1 Then
            Cancel = True
        ElseIf Col = COL_��ʾ Then 'Pass
            Cancel = True
        End If
    End If
End Sub

Private Sub vsAdvice_DblClick()
    With vsAdvice
        If .MouseRow >= .FixedRows And .MouseRow <= .Rows - 1 Then
            If .MouseCol >= .FixedCols And .MouseCol <= .Cols - 1 Then
                Call vsAdvice_KeyPress(13) '��λ����Ӧ�ı༭�ؼ�
            ElseIf .MouseCol = 0 Then
                '��д����
                '##
            End If
        End If
    End With
End Sub

Private Function RowIsLastVisible(ByVal lngRow As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ����һ�ɼ���
    Dim i As Long
    
    With vsAdvice
        For i = .Rows - 1 To .FixedRows Step -1
            If Not .RowHidden(i) Then Exit For
        Next
        If i >= .FixedRows Then
            RowIsLastVisible = lngRow = i
        End If
    End With
End Function

Private Sub vsAdvice_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
'˵����1.OwnerDrawҪ����ΪOver(������Ԫ��������)
'      2.Cell��GridLine�������������ڶ��Ǵӵ�1���߿�ʼ
'      3.Cell��Border�������Ǵӵ�2���߿�ʼ,�����Ǵӵ�1���߿�ʼ
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsAdvice
        If Col <= .FixedCols - 1 Then
            '�����̶����еı����
            SetBkColor hDC, SysColor2RGB(.BackColorFixed)

            '����߱����
            vRect.Left = Left
            vRect.Top = Top
            vRect.Right = Left + 1
            vRect.Bottom = Bottom
            If Row = .Rows - 1 Then vRect.Bottom = vRect.Bottom - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

            '���ϱ߱����
            vRect.Left = Left
            vRect.Top = Top
            vRect.Right = Right
            vRect.Bottom = Top + 1
            If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

            '���±߱����
            vRect.Left = Left
            vRect.Top = Bottom - 1
            vRect.Right = Right
            vRect.Bottom = Bottom
            If RowIsLastVisible(Row) Then vRect.Bottom = vRect.Bottom - 1
            If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

            '���ұ߱����
            vRect.Left = Right - 1
            vRect.Top = Top
            vRect.Right = Right
            vRect.Bottom = Bottom
            If Row = .Rows - 1 Then vRect.Bottom = vRect.Bottom - 1
            If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Else
            lngLeft = COL_��ʼʱ��: lngRight = COL_��ʼʱ��
            If Not Between(Col, lngLeft, lngRight) Then
                lngLeft = COL_Ƶ��: lngRight = COL_�÷�
                If Not Between(Col, lngLeft, lngRight) Then Exit Sub
            End If
            
            If Not RowInһ����ҩ(Row) Then Exit Sub
            If .RowData(Row) = 0 Then
                Call Getһ����ҩ��Χ(Val(.TextMatrix(Row - 1, COL_���ID)), lngBegin, lngEnd)
            Else
                Call Getһ����ҩ��Χ(Val(.TextMatrix(Row, COL_���ID)), lngBegin, lngEnd)
            End If
            
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
            
            If Between(Row, .Row, .RowSel) Then
                SetBkColor hDC, SysColor2RGB(.BackColorSel)
            Else
                SetBkColor hDC, SysColor2RGB(.BackColor)
            End If
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        End If
        Done = True
    End With
End Sub

Private Sub vsAdvice_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If tbr.Buttons("ɾ��").Enabled And tbr.Buttons("ɾ��").Visible Then
            Call tbr_ButtonClick(tbr.Buttons("ɾ��"))
        End If
    End If
End Sub

Private Sub vsAdvice_KeyPress(KeyAscii As Integer)
    Dim objEdit As Object
    
    If KeyAscii = 13 Then
        '��λ����Ӧ�ı༭�ؼ�
        KeyAscii = 0
        Select Case vsAdvice.Col
            Case COL_��ʼʱ��
                If txt��ʼʱ��.TabStop Then
                    Set objEdit = txt��ʼʱ�� 'ȱʡ����λ����ʼʱ��
                Else
                    Set objEdit = txtҽ������
                End If
            Case COL_ҽ������
                Set objEdit = txtҽ������
            Case COL_����
                Set objEdit = txt����
            Case COL_����
                Set objEdit = txt����
            Case COL_�÷�
                Set objEdit = txt�÷�
            Case COL_Ƶ��
                Set objEdit = txtƵ��
            Case COL_ִ��ʱ��
                Set objEdit = cboִ��ʱ��
            Case COL_ִ�п���ID
                Set objEdit = cboִ�п���
            Case COL_ҽ������
                Set objEdit = cboҽ������
            Case COL_��־
                Set objEdit = chk����
        End Select
        If Not objEdit Is Nothing Then
            If objEdit.Enabled And objEdit.Visible Then objEdit.SetFocus
        End If
    End If
End Sub

Private Sub ClearItemTag()
'���ܣ�����ؼ��༭��־
    txt��ʼʱ��.Tag = ""
    txt����.Tag = ""
    txt����.Tag = ""
    txt����.Tag = ""
    txt�÷�.Tag = ""
    txtƵ��.Tag = ""
    cboִ��ʱ��.Tag = ""
    cboҽ������.Tag = ""
    cboִ�п���.Tag = ""
    cboִ������.Tag = ""
    cbo����ִ��.Tag = ""
    chk����.Tag = ""
End Sub

Private Sub SetStartTime(ByVal Editable As Boolean)
'���ܣ����ÿ�ʼʱ���Ƿ�����༭
    'txt��ʼʱ��.TabStop = Editable 'ȱʡ����λ����ʼʱ��
    txt��ʼʱ��.Locked = Not Editable
    cmd��ʼʱ��.Enabled = Editable
    If Editable Then
        txt��ʼʱ��.BackColor = vsAdvice.BackColor
    Else
        txt��ʼʱ��.BackColor = &HE0E0E0
    End If
End Sub

Private Sub SetDayState(Optional ByVal intVisible As Integer, Optional ByVal intEnabled As Integer)
'���ܣ�����ִ���������úͻ��״̬
'������0-���ֲ���,-1-��ֹ,1-����
    If intEnabled = -1 Then
        txt����.Enabled = False
        txt����.BackColor = Me.BackColor
        txt����.Text = ""
    ElseIf intEnabled = 1 Then
        txt����.TabStop = True
        txt����.Enabled = True
        txt����.BackColor = vsAdvice.BackColor
    End If
    
    If intVisible = -1 Then
        lbl����.Visible = False
        txt����.Visible = False
        txt����.Text = ""
        
        lbl����.Left = lbl�÷�.Left + lbl�÷�.Width - lbl����.Width
        txt����.Left = txt�÷�.Left
        txt����.Width = txt�÷�.Width - cmd�÷�.Width - 15
        lbl������λ.Left = txt����.Left + txt����.Width + 30
        
        lbl����.Left = lblƵ��.Left + lblƵ��.Width - lbl����.Width
        txt����.Left = txtƵ��.Left
        txt����.Width = txtƵ��.Width - cmdƵ��.Width - 15
        lbl������λ.Left = txt����.Left + txt����.Width + 30
        
        txt����.TabIndex = cmdƵ��.TabIndex + 1
        txt����.TabIndex = txt����.TabIndex + 1
        txt����.TabIndex = txt����.TabIndex + 1
    ElseIf intVisible = 1 Then
        lbl����.Visible = True
        txt����.Visible = True
        
        lbl����.Left = lbl�÷�.Left + lbl�÷�.Width - lbl����.Width
        txt����.Left = txt�÷�.Left
        txt����.Width = txt�÷�.Width - txt����.Width - Me.TextWidth("������!") - 15
        lbl������λ.Left = txt����.Left + txt����.Width + 30
        
        lbl����.Left = lblƵ��.Left + lblƵ��.Width - lbl����.Width
        txt����.Left = txtƵ��.Left
        txt����.Width = txtƵ��.Width - cmdƵ��.Width - 15
        lbl������λ.Left = txt����.Left + txt����.Width + 30
        
        txt����.TabIndex = cmdƵ��.TabIndex + 1
        txt����.TabIndex = txt����.TabIndex + 1
        txt����.TabIndex = txt����.TabIndex + 1
    End If
End Sub

Private Sub SetItemEditable(Optional int���� As Integer, Optional int���� As Integer, _
    Optional int�÷� As Integer, Optional intƵ�� As Integer, _
    Optional intִ��ʱ�� As Integer, Optional intִ�п��� As Integer, _
    Optional intִ������ As Integer, Optional int����ִ�� As Integer)
'���ܣ�����ָ���༭��Ŀ���״̬
'������0-���ֲ���,-1-��ֹ,1-����,2-����
'˵������ֹʱ,ͬʱ�������Ŀ����(����ȫ��)

    '��������Ϊ��ֹʱ,����������ı�,�Ӷ���������Validate�¼�,�����Ƚ�ֹ����˳��
    If int���� = -1 Then txt����.TabStop = False
    If int���� = -1 Then txt����.TabStop = False
    If int�÷� = -1 Then txt�÷�.TabStop = False
    If intƵ�� = -1 Then txtƵ��.TabStop = False
    If intִ��ʱ�� = -1 Then cboִ��ʱ��.TabStop = False
    If intִ�п��� = -1 Then cboִ�п���.TabStop = False
    If intִ������ = -1 Then cboִ������.TabStop = False
    If int����ִ�� = -1 Then cbo����ִ��.TabStop = False
    
    If int���� = -1 Then
        txt����.Enabled = False
        txt����.BackColor = Me.BackColor
        txt����.Text = ""
        lbl������λ.Caption = "" '"��λ"
    ElseIf int���� = 1 Then
        txt����.TabStop = True
        txt����.Enabled = True
        txt����.BackColor = vsAdvice.BackColor
    End If

    If int���� = -1 Then
        txt����.Enabled = False
        txt����.BackColor = Me.BackColor
        txt����.Text = ""
        lbl������λ.Caption = "" '"��λ"
    ElseIf int���� = 1 Then
        txt����.TabStop = True
        txt����.Enabled = True
        txt����.BackColor = vsAdvice.BackColor
    End If
    
    If int�÷� = -1 Then
        txt�÷�.Enabled = False
        txt�÷�.BackColor = Me.BackColor
        txt�÷�.Text = ""
        cmd�÷�.Enabled = False
        lbl�÷�.Caption = "�÷�"
    ElseIf int�÷� = 1 Then
        txt�÷�.TabStop = True
        txt�÷�.Enabled = True
        cmd�÷�.Enabled = True
        txt�÷�.BackColor = vsAdvice.BackColor
    End If

    If intƵ�� = -1 Then
        txtƵ��.Enabled = False
        cmdƵ��.Enabled = False
        txtƵ��.BackColor = Me.BackColor
        txtƵ��.Text = ""
    ElseIf intƵ�� = 1 Then
        txtƵ��.TabStop = True
        txtƵ��.Enabled = True
        cmdƵ��.Enabled = True
        txtƵ��.BackColor = vsAdvice.BackColor
    End If

    If intִ��ʱ�� = -1 Then
        cboִ��ʱ��.Enabled = False
        cboִ��ʱ��.BackColor = Me.BackColor
        cboִ��ʱ��.Clear
    ElseIf intִ��ʱ�� = 1 Then
        cboִ��ʱ��.TabStop = True
        cboִ��ʱ��.Enabled = True
        cboִ��ʱ��.BackColor = vsAdvice.BackColor
    End If

    If intִ�п��� = -1 Then
        lblִ�п���.Caption = "ִ�п���"
        cboִ�п���.Enabled = False
        cboִ�п���.BackColor = Me.BackColor
        cboִ�п���.Clear
    ElseIf intִ�п��� = 1 Then
        lblִ�п���.Caption = "ִ�п���"
        cboִ�п���.TabStop = True
        cboִ�п���.Enabled = True
        cboִ�п���.BackColor = vsAdvice.BackColor
    End If

    If intִ������ = -1 Then
        cboִ������.Enabled = False
        cboִ������.BackColor = Me.BackColor
        Call zlControl.CboSetIndex(cboִ������.Hwnd, -1) '�����
    ElseIf intִ������ = 1 Then
        cboִ������.TabStop = True
        cboִ������.Enabled = True
        cboִ������.BackColor = vsAdvice.BackColor
    End If
    
    If int����ִ�� = -1 Then
        lbl����ִ��.Caption = "����ִ��"
        cbo����ִ��.Enabled = False
        cbo����ִ��.BackColor = Me.BackColor
        cbo����ִ��.Clear
    ElseIf int����ִ�� = 1 Then
        lbl����ִ��.Caption = "����ִ��"
        cbo����ִ��.TabStop = True
        cbo����ִ��.Enabled = True
        cbo����ִ��.BackColor = vsAdvice.BackColor
    End If
End Sub

Private Function GetPatiInfo() As Boolean
'���ܣ���ȡ������Ϣ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    'ִ�в���(�ű����)�����˿���
    strSQL = "Select B.����,B.�Ա�,B.����,B.�����,B.�ѱ�,B.ҽ�Ƹ��ʽ," & _
        " Nvl(D.Ԥ�����,0)-Nvl(D.�������,0) as Ԥ����,B.����,B.��������," & _
        " C.���� as ִ�в���,A.ִ�в���ID,A.�Ǽ�ʱ��" & _
        " From ���˹Һż�¼ A,������Ϣ B,���ű� C,������� D" & _
        " Where A.NO=[1] And A.����ID+0=[2]" & _
        " And A.����ID=B.����ID And A.ִ�в���ID=C.ID" & _
        " And B.����ID=D.����ID(+) And D.����(+)=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr�Һŵ�, mlng����ID)
    
    lblPati.Caption = _
        "������" & rsTmp!���� & "���Ա�" & Nvl(rsTmp!�Ա�) & "�����䣺" & Nvl(rsTmp!����) & _
        "���ѱ�" & Nvl(rsTmp!�ѱ�) & "�� ҽ�Ƹ��ʽ��" & Nvl(rsTmp!ҽ�Ƹ��ʽ) & "��Ԥ���" & Format(Nvl(rsTmp!Ԥ����, 0), "0.00")
    lblPati.Tag = rsTmp!���� '����ҽ��������ʾ
    
    '���˵�׼ȷ����:�����ж�
    mint���� = GetPatiYear(mlng����ID)
    mstr�Ա� = Nvl(rsTmp!�Ա�)
    mdat�Һ�ʱ�� = rsTmp!�Ǽ�ʱ��
    mlng���˿���id = rsTmp!ִ�в���ID
    mstr������ = Getҽ�Ƹ�����(Nvl(rsTmp!ҽ�Ƹ��ʽ))
    
    '���ղ����ú�ɫ��ʾ
    mint���� = 0
    If Not IsNull(rsTmp!����) Then
        mint���� = rsTmp!����
        lblPati.ForeColor = vbRed
    End If
    
    GetPatiInfo = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetPreRow(ByVal lngRow As Long) As Long
'���ܣ�ȡ��һ�����Ч�ɼ���
'���أ�����Ч��ʱ,����-1
    Dim lngTmp As Long, i As Long
    
    lngTmp = -1
    For i = lngRow - 1 To vsAdvice.FixedRows Step -1
        If vsAdvice.RowData(i) <> 0 And Not vsAdvice.RowHidden(i) Then
            lngTmp = i: Exit For
        End If
    Next
    GetPreRow = lngTmp
End Function

Private Function GetNextRow(ByVal lngRow As Long) As Long
'���ܣ�ȡ��һ�����Ч�ɼ���
'���أ�����Ч��ʱ,����-1
    Dim lngTmp As Long, i As Long
    
    lngTmp = -1
    For i = lngRow + 1 To vsAdvice.Rows - 1
        If vsAdvice.RowData(i) <> 0 And Not vsAdvice.RowHidden(i) Then
            lngTmp = i: Exit For
        End If
    Next
    GetNextRow = lngTmp
End Function

Private Function GetDefaultTime(lngRow As Long) As String
'���ܣ���ȡ�¿�ҽ����ȱʡ��ʼʱ��
'˵����
'      ���һ����Чʱ��Ϊ���죬�Ҽ�������ڰ�Сʱ���ڣ����������ͬ
'      ���û��,��ȡ����¿�һ����ʱ��
'      ���û��,��ȡ��ǰʱ��
    Dim curDate As Date, strDate As String, i As Long
    
    curDate = zlDatabase.Currentdate
    
    With vsAdvice
        '�ȴӵ�ǰ�������:����ȱʡΪ������Ч��ʱ��
        For i = lngRow - 1 To .FixedRows Step -1
            If .RowData(i) <> 0 And Not .RowHidden(i) And IsDate(.Cell(flexcpData, i, COL_��ʼʱ��)) Then
                If Format(.Cell(flexcpData, i, COL_��ʼʱ��), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd") Then
                    If DateAdd("n", 30, CDate(.Cell(flexcpData, i, COL_��ʼʱ��))) >= curDate Then
                        strDate = Format(.Cell(flexcpData, i, COL_��ʼʱ��), "yyyy-MM-dd HH:mm")
                        Exit For
                    End If
                End If
            End If
        Next
            
        '�ٴ�����������:����ȱʡΪ������Ч��ʱ��
        If strDate = "" Then
            For i = .Rows - 1 To lngRow + 1 Step -1
                If .RowData(i) <> 0 And Not .RowHidden(i) And IsDate(.Cell(flexcpData, i, COL_��ʼʱ��)) Then
                    If Format(.Cell(flexcpData, i, COL_��ʼʱ��), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd") Then
                        If DateAdd("n", 30, CDate(.Cell(flexcpData, i, COL_��ʼʱ��))) >= curDate Then
                            strDate = Format(.Cell(flexcpData, i, COL_��ʼʱ��), "yyyy-MM-dd HH:mm")
                            Exit For
                        End If
                    End If
                End If
            Next
        End If
        
        If strDate = "" Then
            '�ȴӵ�ǰ�������
            For i = lngRow - 1 To .FixedRows Step -1
                If .RowData(i) <> 0 And Not .RowHidden(i) And IsDate(.Cell(flexcpData, i, COL_��ʼʱ��)) _
                    And Val(.TextMatrix(i, COL_EDIT)) = 1 Then
                    strDate = Format(.Cell(flexcpData, i, COL_��ʼʱ��), "yyyy-MM-dd HH:mm")
                    Exit For
                End If
            Next
            '�ٴ�����������
            If strDate = "" Then
                For i = .Rows - 1 To lngRow + 1 Step -1
                    If .RowData(i) <> 0 And Not .RowHidden(i) And IsDate(.Cell(flexcpData, i, COL_��ʼʱ��)) _
                        And Val(.TextMatrix(i, COL_EDIT)) = 1 Then
                        strDate = Format(.Cell(flexcpData, i, COL_��ʼʱ��), "yyyy-MM-dd HH:mm")
                        Exit For
                    End If
                Next
            End If
        End If
    End With
    If strDate = "" Then strDate = Format(curDate, "yyyy-MM-dd HH:mm")
    GetDefaultTime = strDate
End Function

Private Function GetCurRow���(lngRow As Long) As Long
'���ܣ���ȡָ���п��õĵ����
'������lngRow=Ҫȡ��ŵ���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lng��� As Long, i As Long
    Dim lng���1 As Long, lng���2 As Long
            
    'ȡ֮�����һ����Ч���,ֱ��ʹ��
    For i = lngRow + 1 To vsAdvice.Rows - 1
        If vsAdvice.RowData(i) <> 0 Then
            If Val(vsAdvice.TextMatrix(i, COL_Ӥ��)) = cboӤ��.ListIndex _
                And IsNumeric(vsAdvice.TextMatrix(i, COL_���)) Then
                lng��� = Val(vsAdvice.TextMatrix(i, COL_���))
                Exit For
            End If
        End If
    Next
    If lng��� = 0 Then
        '����û��,��ȡ���ݿ�֮�е���������֮ǰ�������űȽ�
        On Error GoTo errH
        strSQL = "Select Max(���) as ��� From ����ҽ����¼ Where ����ID+0=[1] And �Һŵ�=[2] And Nvl(Ӥ��,0)=[3]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mstr�Һŵ�, cboӤ��.ListIndex)
        If Not rsTmp.EOF Then lng���1 = Nvl(rsTmp!���, 0)
        On Error GoTo 0
        
        For i = lngRow - 1 To vsAdvice.FixedRows Step -1
            If vsAdvice.RowData(i) <> 0 Then
                If Val(vsAdvice.TextMatrix(i, COL_Ӥ��)) = cboӤ��.ListIndex _
                    And IsNumeric(vsAdvice.TextMatrix(i, COL_���)) Then
                    lng���2 = Val(vsAdvice.TextMatrix(i, COL_���))
                    Exit For
                End If
            End If
        Next
        
        If lng���1 > lng���2 Then
            lng��� = lng���1
        Else
            lng��� = lng���2
        End If

        If lng��� <> 0 Then lng��� = lng��� + 1 '������+1
    End If
    If lng��� = 0 Then lng��� = 1
    GetCurRow��� = lng���
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub AdviceSetҽ�����(lngRow As Long, intStep As Integer)
'���ܣ�����ǰ����ҽ����¼�����ǰ�ƻ����
'������lngRow=��ʼ������,intStep=��������,��1��-1
    Dim i As Long
    
    For i = lngRow To vsAdvice.Rows - 1
        If vsAdvice.RowData(i) <> 0 Then
            If Val(vsAdvice.TextMatrix(i, COL_Ӥ��)) = cboӤ��.ListIndex _
                And IsNumeric(vsAdvice.TextMatrix(i, COL_���)) Then
                vsAdvice.TextMatrix(i, COL_���) = Val(vsAdvice.TextMatrix(i, COL_���)) + intStep
                If Val(vsAdvice.TextMatrix(i, COL_EDIT)) = 0 Then
                    vsAdvice.TextMatrix(i, COL_EDIT) = 3 '��־�޸������
                End If
            End If
        End If
    Next
End Sub

Private Sub AdviceDelete(ByVal lngRow As Long)
'���ܣ�ָ����ҽ��ɾ������
    Dim lngBegin As Long, lngEnd As Long
    Dim lng���ID As Long, blnGroup As Boolean
    Dim lngҽ��ID As Long, i As Integer
    
    mblnRowChange = False
    vsAdvice.Redraw = flexRDNone
    
    If vsAdvice.RowData(lngRow) <> 0 Then
        If InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_���)) > 0 Then
            lngҽ��ID = vsAdvice.RowData(lngRow)
            lng���ID = Val(vsAdvice.TextMatrix(lngRow, COL_���ID))
            blnGroup = RowInһ����ҩ(lngRow)
            If blnGroup Then
                '��ɾ��һ����ҩ�еĿ���(һ��Ҫɾ)
                Call Getһ����ҩ��Χ(lng���ID, lngBegin, lngEnd)
                For i = lngEnd To lngBegin Step -1 '���뷴��
                    If vsAdvice.RowData(i) = 0 Then Call DeleteRow(i)
                Next
                
                'ɾ��֮��ǰ�кſ��ܱ���
                lngRow = vsAdvice.FindRow(lngҽ��ID, lngBegin)
                
                'һ����ҩֻɾ����ǰ��
                Call DeleteRow(lngRow)
            Else
                '�����ĳ�ҩ��ɾ����ҩ;���м���ǰ��
                i = vsAdvice.FindRow(CLng(vsAdvice.TextMatrix(lngRow, COL_���ID)), lngRow + 1)
                Call DeleteRow(i)
                Call DeleteRow(lngRow)
            End If
        ElseIf InStr(",D,F,", vsAdvice.TextMatrix(lngRow, COL_���)) > 0 Then
            Call Delete�������(lngRow)
            Call DeleteRow(lngRow)
        ElseIf RowIn�䷽��(lngRow) Then
            'ɾ�����ζҩ���巨��:ɾ��֮�����¶�λ�ĵ�ǰ��
            lngRow = Delete��ҩ�䷽(lngRow)
            'ɾ����ǰ��(��ҩ�÷���)
            Call DeleteRow(lngRow)
        ElseIf RowIn������(lngRow) Then
            lngRow = Delete�������(lngRow)
            Call DeleteRow(lngRow)
        Else
            Call DeleteRow(lngRow)
        End If
        
        mblnNoSave = True '���Ϊδ����
    Else
        '����ֱ��ɾ��
        Call DeleteRow(lngRow)
    End If
    
    '���¶�λ��
    If vsAdvice.RowHidden(vsAdvice.Row) Then
        i = GetPreRow(vsAdvice.Row)
        If i = -1 Then i = GetNextRow(vsAdvice.Row)
        If i <> -1 Then vsAdvice.Row = i
    End If
    
    Call vsAdvice.ShowCell(vsAdvice.Row, vsAdvice.Col)
    
    mblnRowChange = True
    vsAdvice.Redraw = flexRDDirect
    Call vsAdvice_AfterRowColChange(-1, vsAdvice.Col, vsAdvice.Row, vsAdvice.Col)
End Sub

Private Sub DeleteRow(ByVal lngRow As Long, Optional ByVal blnClear As Boolean, Optional blnDelID As Boolean = True)
'���ܣ�ɾ������е�һ��,�����ı䵱ǰ��
'������blnClear=�Ƿ�������������,��ɾ��
'      blnDelID=�Ƿ��¼Ҫɾ����ҽ��ID
    Dim lngCol As Long, blnDraw As Boolean, blnChange As Boolean
    
    With vsAdvice
        lngCol = .Col
        blnDraw = .Redraw
        blnChange = mblnRowChange
        
        mblnRowChange = False
        .Redraw = flexRDNone
        
        If .RowData(lngRow) <> 0 Then
            '�������
            Call AdviceSetҽ�����(lngRow + 1, -1)
            
            '��¼Ҫɾ����ID(���˲�������)
            If Val(.TextMatrix(lngRow, COL_EDIT)) <> 1 And blnDelID Then
                mstrDelIDs = mstrDelIDs & "," & .RowData(lngRow)
            End If
        End If
            
        '���Ϊ��1�ҽ�ʣ��1������,����
        If Not (lngRow = .FixedRows And .Rows = .FixedRows + 1) And Not blnClear Then
            .RemoveItem lngRow
        Else
            '�����������
            .RowData(lngRow) = Empty
            .Cell(flexcpText, lngRow, 0, lngRow, .Cols - 1) = "" '����
            .Cell(flexcpData, lngRow, 0, lngRow, .Cols - 1) = Empty '����
            .Cell(flexcpFontBold, lngRow, .FixedCols, lngRow, .Cols - 1) = False '����
            .Cell(flexcpForeColor, lngRow, .FixedCols, lngRow, .Cols - 1) = .ForeColor '����ɫ
            .Cell(flexcpForeColor, lngRow, 0, lngRow, .FixedCols - 1) = .ForeColorFixed '�̶�������ɫ
            .Cell(flexcpBackColor, lngRow, 0, lngRow, .FixedCols - 1) = .BackColorFixed '�̶��б���ɫ
            Set .Cell(flexcpPicture, lngRow, 0, lngRow, .Cols - 1) = Nothing '��ԪͼƬ
            Set .Cell(flexcpPicture, lngRow, COL_��ʾ) = Nothing 'Pass��ʾ��
            
            '��Ԫ��߿�
            .Select lngRow, .FixedCols, lngRow, COL_��־
            .CellBorder vbRed, 0, 0, 0, 0, 0, 0
        End If
        
        .Col = lngCol '��Ϊ��ɾ����,���Ե��ó���϶����ж�λ,���Բ��ػָ���
        .Redraw = blnDraw
        mblnRowChange = blnChange
    End With
End Sub

Private Sub Delete�������(ByVal lngRow As Long)
'���ܣ�1.ɾ����������Ŀ�Ĳ�λ��
'      2.ɾ��������Ŀ�ĸ��������м�������Ŀ��
    Dim lngBegin As Long, lngEnd As Long, i As Long
    
    i = vsAdvice.FindRow(CStr(vsAdvice.RowData(lngRow)), lngRow + 1, COL_���ID) '��һ����,�����ò���
    If i <> -1 Then
        lngBegin = i
        For i = lngBegin To vsAdvice.Rows - 1
            If Val(vsAdvice.TextMatrix(i, COL_���ID)) = vsAdvice.RowData(lngRow) Then
                lngEnd = i
            Else
                Exit For
            End If
        Next
        For i = lngEnd To lngBegin Step -1
            Call DeleteRow(i)
        Next
    End If
End Sub

Private Function Delete��ҩ�䷽(ByVal lngRow As Long) As Long
'���ܣ�ɾ����ҩ�䷽�����ζҩ���巨��
'������lngRow=��ҩ�䷽�÷���(�ɼ�)
'���أ�ɾ��֮�����¶�λ�ĵ�ǰ��(��ҩ�÷���)
    Dim lngBegin As Long, lngEnd As Long
    Dim lngҽ��ID As Long, i As Long
    
    lngҽ��ID = vsAdvice.RowData(lngRow)
    
    lngEnd = lngRow - 1
    For i = lngEnd To vsAdvice.FixedRows Step -1
        If Val(vsAdvice.TextMatrix(i, COL_���ID)) = lngҽ��ID Then
            lngBegin = i
        Else
            Exit For
        End If
    Next
    
    mblnRowChange = False
    For i = lngEnd To lngBegin Step -1
        Call DeleteRow(i)
    Next
    
    '��Ϊ����ǰ��ɾ��,��Ҫ���¶�λ����ҩ�÷���
    i = vsAdvice.FindRow(lngҽ��ID)
    vsAdvice.Row = i '�������Ҳ���
    
    mblnRowChange = True
    
    Delete��ҩ�䷽ = vsAdvice.Row
End Function

Private Function Delete�������(ByVal lngRow As Long) As Long
'���ܣ�ɾ��һ���ɼ��Ķ��������Ŀ��
'������lngRow=�ɼ�������(�ɼ�)
'���أ�ɾ��֮�����¶�λ�ĵ�ǰ��(�ɼ�������)
    Dim lngBegin As Long, lngEnd As Long
    Dim lngҽ��ID As Long, i As Long
    
    lngҽ��ID = vsAdvice.RowData(lngRow)
    
    lngEnd = lngRow - 1
    For i = lngEnd To vsAdvice.FixedRows Step -1
        If Val(vsAdvice.TextMatrix(i, COL_���ID)) = lngҽ��ID Then
            lngBegin = i
        Else
            Exit For
        End If
    Next
    
    mblnRowChange = False
    For i = lngEnd To lngBegin Step -1
        Call DeleteRow(i)
    Next
    
    '��Ϊ����ǰ��ɾ��,��Ҫ���¶�λ���ɼ�������
    i = vsAdvice.FindRow(lngҽ��ID)
    vsAdvice.Row = i '�������Ҳ���
    
    mblnRowChange = True
    
    Delete������� = vsAdvice.Row
End Function

Private Function Get��鲿λIDs(ByVal lngRow As Long) As String
'���ܣ���ȡָ���еļ�鲿λID��
'���أ�"��λID1,��λID2,..."
    Dim strTmp As String, i As Long
    
    i = vsAdvice.FindRow(CStr(vsAdvice.RowData(lngRow)), lngRow + 1, COL_���ID)
    If i <> -1 Then
        For i = i To vsAdvice.Rows - 1
            If Val(vsAdvice.TextMatrix(i, COL_���ID)) = vsAdvice.RowData(lngRow) Then
                strTmp = strTmp & "," & Val(vsAdvice.TextMatrix(i, COL_������ĿID))
            Else
                Exit For
            End If
        Next
    End If
    Get��鲿λIDs = Mid(strTmp, 2)
End Function

Private Function Get��������IDs(ByVal lngRow As Long) As String
'���ܣ���ȡָ�������еĸ���������������ĿID��
'���أ�"����ID1,����ID2,...;����ID",���п���û�и�������������
    Dim strTmp As String, lng����ID As Long, i As Long
    
    i = vsAdvice.FindRow(CStr(vsAdvice.RowData(lngRow)), lngRow + 1, COL_���ID)
    If i <> -1 Then
        For i = i To vsAdvice.Rows - 1
            If Val(vsAdvice.TextMatrix(i, COL_���ID)) = vsAdvice.RowData(lngRow) Then
                If vsAdvice.TextMatrix(i, COL_���) = "G" Then
                    lng����ID = Val(vsAdvice.TextMatrix(i, COL_������ĿID))
                Else
                    strTmp = strTmp & "," & Val(vsAdvice.TextMatrix(i, COL_������ĿID))
                End If
            Else
                Exit For
            End If
        Next
    End If
    Get��������IDs = Mid(strTmp, 2) & ";" & IIF(lng����ID = 0, "", lng����ID)
End Function

Private Function Get��ҩ�䷽IDs(ByVal lngRow As Long) As String
'���ܣ���ȡ��ҩ�䷽�����ζҩ���巨ID��
'���أ�"��ҩID1,����1,��ע1;��ҩID2,����2,��ע2;...|�巨ID"
    Dim lng�巨ID As Long, str��ҩIDs As String, i As Long
    
    With vsAdvice
        For i = lngRow - 1 To .FixedRows Step -1
            If Val(.TextMatrix(i, COL_���ID)) = .RowData(lngRow) Then
                If .TextMatrix(i, COL_���) = "E" Then
                    lng�巨ID = Val(.TextMatrix(i, COL_������ĿID))
                ElseIf .TextMatrix(i, COL_���) = "7" Then
                    str��ҩIDs = Val(.TextMatrix(i, COL_������ĿID)) & "," & _
                        .TextMatrix(i, COL_����) & "," & .TextMatrix(i, COL_ҽ������) & _
                        ";" & str��ҩIDs
                End If
            Else
                Exit For
            End If
        Next
    End With
    Get��ҩ�䷽IDs = Mid(str��ҩIDs, 1, Len(str��ҩIDs) - 1) & "|" & lng�巨ID
End Function

Private Function Get�������IDs(ByVal lngRow As Long) As String
'���ܣ���ȡһ���ɼ��ļ��������ĿID���걾
'���أ�"��ĿID1,��ĿID2,...;����걾"
    Dim str��ĿIDs As String, str�걾 As String, i As Long
    
    With vsAdvice
        For i = lngRow - 1 To .FixedRows Step -1
            If Val(.TextMatrix(i, COL_���ID)) = .RowData(lngRow) Then
                str��ĿIDs = Val(.TextMatrix(i, COL_������ĿID)) & "," & str��ĿIDs
                str�걾 = .TextMatrix(i, COL_�걾��λ)
            Else
                Exit For
            End If
        Next
    End With
    Get�������IDs = Left(str��ĿIDs, Len(str��ĿIDs) - 1) & ";" & str�걾
End Function

Private Function RowIn������(ByVal lngRow As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ����ڼ�������е�һ��
'˵���������е�ǰ�Ƿ�����
    If lngRow = -1 Then Exit Function
    If vsAdvice.RowData(lngRow) = 0 Then Exit Function
    
    With vsAdvice
        If .TextMatrix(lngRow, COL_���) = "E" And Val(.TextMatrix(lngRow, COL_���ID)) = 0 Then
            '�ɼ�������
            If .TextMatrix(lngRow - 1, COL_���) = "C" _
                And Val(.TextMatrix(lngRow - 1, COL_���ID)) = .RowData(lngRow) Then
                RowIn������ = True: Exit Function
            End If
        ElseIf .TextMatrix(lngRow, COL_���) = "C" And Val(.TextMatrix(lngRow, COL_���ID)) <> 0 Then
            '������Ŀ��
            RowIn������ = True: Exit Function
        End If
    End With
End Function

Private Function RowIn�䷽��(ByVal lngRow As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ�������ҩ�䷽�е�һ��
'˵���������е�ǰ�Ƿ�����
    If lngRow = -1 Then Exit Function
    If vsAdvice.RowData(lngRow) = 0 Then Exit Function
    
    With vsAdvice
        If .TextMatrix(lngRow, COL_���) = "E" Then
            If Val(.TextMatrix(lngRow, COL_���ID)) = 0 Then
                '�÷���
                If Val(.TextMatrix(lngRow - 1, COL_���ID)) = .RowData(lngRow) _
                    And .TextMatrix(lngRow - 1, COL_���) = "E" Then
                    RowIn�䷽�� = True: Exit Function
                End If
            Else
                '�巨��
                If .TextMatrix(lngRow - 1, COL_���) = "7" _
                    And Val(.TextMatrix(lngRow - 1, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                    RowIn�䷽�� = True: Exit Function
                End If
            End If
        ElseIf .TextMatrix(lngRow, COL_���) = "7" And Val(.TextMatrix(lngRow, COL_���ID)) <> 0 Then
            '��ҩ��
            RowIn�䷽�� = True: Exit Function
        End If
    End With
End Function

Private Function RowInһ����ҩ(ByVal lngRow As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ���һ����ҩ�ķ�Χ��
'������lngRow=�ɼ�����,�����ǿ���
'˵����һ����ҩ�ķ�Χ�п��ܴ��ڿ���
    Dim lngPreRow As Long, lngNextRow As Long
    Dim lng���ID As Long, blnGroup As Boolean, i As Long
    
    lngPreRow = GetPreRow(lngRow)
    lngNextRow = GetNextRow(lngRow)
    
    With vsAdvice
        If .RowData(lngRow) = 0 Then
            If lngPreRow <> -1 And lngNextRow <> -1 Then
                If Val(.TextMatrix(lngPreRow, COL_���ID)) = Val(.TextMatrix(lngNextRow, COL_���ID)) _
                    And Val(.TextMatrix(lngPreRow, COL_���ID)) <> 0 _
                    And InStr(",5,6,", .TextMatrix(lngPreRow, COL_���)) > 0 _
                    And InStr(",5,6,", .TextMatrix(lngNextRow, COL_���)) > 0 Then
                    blnGroup = True
                End If
            End If
        ElseIf InStr(",5,6,", .TextMatrix(lngRow, COL_���)) > 0 _
            And Val(.TextMatrix(lngRow, COL_���ID)) <> 0 Then
            
            lng���ID = Val(.TextMatrix(lngRow, COL_���ID))
            If lngPreRow <> -1 Then
                If InStr(",5,6,", .TextMatrix(lngPreRow, COL_���)) > 0 _
                    And Val(.TextMatrix(lngPreRow, COL_���ID)) = lng���ID Then blnGroup = True
            End If
            If Not blnGroup And lngNextRow <> -1 Then
                If InStr(",5,6,", .TextMatrix(lngNextRow, COL_���)) > 0 _
                    And Val(.TextMatrix(lngNextRow, COL_���ID)) = lng���ID Then blnGroup = True
            End If
        End If
    End With
    RowInһ����ҩ = blnGroup
End Function

Private Function Check��������(ByVal lngҩ��ID As Long, ByVal str���� As String, Optional lngƤ��ID As Long) As String
'���ܣ��������ҩ���г�ҩ�Ĺ�������
'������lngҩ��ID=ҩƷ������ĿID
'      str����=ҩƷ����,������ʾ
'���أ�Ϊ�ձ�ʾͨ��
'      lngƤ��ID=Ҫ�Զ���ӵ�Ƥ����ĿID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strMsg As String
    
    lngƤ��ID = 0
    
    On Error GoTo errH
    
    'ȡ��Чʱ���ڵ����һ�ι�������Ǽ�
    strSQL = "Select ҩ����,���,��¼ʱ�� From ���˹�����¼" & _
        " Where ����ID=[1] And ҩ��ID=[2] And Trunc(��¼ʱ��)>=Trunc(Sysdate-[3])" & _
        " Order by ��¼ʱ�� Desc"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, lngҩ��ID, gint�����Ǽ���Ч����)
    If Not rsTmp.EOF Then
        '�й�������ǼǼ�¼,�����Ƿ����Ծ����Ƿ���ʾ
        If Nvl(rsTmp!���, 0) = 1 Then
            strMsg = "�ò�����" & Format(rsTmp!��¼ʱ��, "M��d��") & "�Ĺ���ʵ���ж�""" & Nvl(rsTmp!ҩ����, str����) & """����(+)��" & _
                vbCrLf & vbCrLf & "�Ƿ���Ȼʹ�ø�ҩƷ��"
        Else
            strMsg = "" 'Ϊ����,ͨ��
        End If
    Else
        '�޹�������ǼǼ�¼,���ȿ���ҩƷ�Ƿ���ҪƤ��
        strSQL = "Select A.�÷�ID,B.����" & _
            " From �����÷����� A,������ĿĿ¼ B" & _
            " Where A.�÷�ID=B.ID And A.����=0 And A.��ĿID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҩ��ID)
        If Not rsTmp.EOF Then
            If mbln�Զ�Ƥ�� Then
                lngƤ��ID = rsTmp!�÷�ID
                strMsg = "" '�Զ����,����ʾ
            Else
                'Ҫ��Ƥ��,����ʾƤ��
                strMsg = "�ڶԲ���ʹ��""" & str���� & """ǰ��Ҫ���Ƚ���""" & rsTmp!���� & """��" & vbCrLf & _
                    "��û�з�����Ч�Ĺ������������Ƿ���Ȼʹ�ø�ҩƷ��"
            End If
        Else
            strMsg = "" 'û��Ƥ��Ҫ��,ͨ��
        End If
    End If
    Check�������� = strMsg
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function AdviceSet��������(ByVal lngRow As Long, ByVal lngƤ��ID As Long) As Boolean
'���ܣ��Զ�����Ƥ����
'������lngRow=��ǰ������,�Ѿ�������ҩ���г�ҩ
'      lngƤ��ID=Ҫ���ӵ�Ƥ����ĿID
'˵�����Զ�����֮��,��ǰ�м�����Զ�λ���Ѹ������ҩƷ��λ��
    Dim rsInput As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    On Error GoTo errH
        
    '���ID,����,������ĿID,�շ�ϸĿID,���,����
    strSQL = "Select ��� as ���ID,����,ID as ������ĿID,NULL as �շ�ϸĿID,NULL as ���,NULL as ���� From ������ĿĿ¼ Where ID=[1]"
    Set rsInput = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngƤ��ID)
        
    'Ѱ��ʵ��Ҫ����Ƥ�Ե���:һ����ҩ�����
    With vsAdvice
        For i = lngRow - 1 To .FixedRows - 1 Step -1 '��������������ǰ��
            If Val(.TextMatrix(i, COL_���ID)) <> Val(.TextMatrix(lngRow, COL_���ID)) Then
                lngRow = i + 1: Exit For '�����е��к�
            End If
        Next
    End With
    
    '�������
    vsAdvice.AddItem "", lngRow
    
    '����Ƥ��
    Call AdviceInput(rsInput, lngRow, True)
    
    '���¶�λ�������ҩƷ��
    mblnRowChange = False
    vsAdvice.Row = vsAdvice.Row + 1
    mblnRowChange = True
    
    AdviceSet�������� = True
    Exit Function
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function AdviceInput(rsInput As ADODB.Recordset, ByVal lngRow As Long, Optional ByVal blnByƤ�� As Boolean) As Boolean
'���ܣ����������������Ŀ(���������)����ȱʡ��ҽ������
'������rsInput=�����ѡ�񷵻صļ�¼��,lngRow=��ǰ������,blnByƤ��=�Ƿ��Զ�����Ƥ������
'���أ�����¼���Ƿ���Ч
    Dim str���� As String, blnGroup As Boolean, i As Long
    Dim lng�÷�ID As Long, lngGroupRow As Long
    Dim lngƤ��ID As Long, blnƤ���� As Boolean
    Dim lngPreRow As Long, lngNextRow As Long
    Dim strExtData As String, intType As Integer
    Dim strMsg As String
    
    On Error GoTo errH
        
    lngPreRow = GetPreRow(lngRow) 'ȡ��һ��Ч��,ĳЩ����ȱʡ����һ����ͬ
    lngNextRow = GetNextRow(lngRow) 'ȡ��һ��Ч��
    
    '��Ŀ�����������뼰����Ϸ��Լ��
    '---------------------------------------------------------------------------------------------------------------
    txtҽ������.Text = rsInput!���� '��ʱ��ʾ
    
    'ҩƷ����ְ����(��ʿվ�ڱ���ʱ���)
    If InStr(",5,6,7,", rsInput!���ID) > 0 Then
        strMsg = CheckOneDuty(rsInput!����, Nvl(rsInput!����ְ��ID), UserInfo.����, InStr(",1,2,", mstr������) > 0 And mstr������ <> "")
        If strMsg <> "" Then
            vsAdvice.Refresh
            MsgBox strMsg, vbInformation, gstrSysName
            vsAdvice.Refresh: Exit Function
        End If
    End If
    
    If Not IsNull(rsInput!�շ�ϸĿID) And InStr(",5,6,7,", rsInput!���ID) > 0 Then 'mint���� <> 0
        Call gclsInsure.GetItemInfo(mint����, mlng����ID, rsInput!�շ�ϸĿID) '��ҽ������ҲҪ��
    End If
    
    With vsAdvice
        '������Ŀ���ɼ������ж�
        If rsInput!���ID = "C" Then
            '����������ȡһ��ȱʡ�Ĳɼ�����,ͬʱ�ж��Ƿ��вɼ���������
            lng�÷�ID = Getȱʡ�÷�ID(6, 1)
            If lng�÷�ID = 0 Then
                .Refresh
                MsgBox "û�п��õı걾�ɼ�����,���ȵ�������Ŀ���������ã�", vbInformation, gstrSysName
                .Refresh: Exit Function
            End If
            'ȱʡ����һ����ͬ
            If lngPreRow <> -1 Then
                If RowIn������(lngPreRow) Then
                    lng�÷�ID = Val(.TextMatrix(lngPreRow, COL_������ĿID))
                End If
            End If
        End If
        
        '��ҩ�䷽����������ҩ�÷��ж�
        If InStr(",7,8,", rsInput!���ID) > 0 Then
            If rsInput!���ID = "8" Then
                If GetGroupCount(rsInput!������ĿID, 1, False) = 0 Then
                    .Refresh
                    MsgBox """" & rsInput!���� & """��һ����ҩ�䷽����û��������Ч�������ҩ��" & vbCrLf & "���ȵ�������Ŀ���������á�", vbInformation, gstrSysName
                    .Refresh: Exit Function
                End If
            
                '����ҩ��Ч����ʾ
                strMsg = GetGroupNone(rsInput!������ĿID, 1)
                If strMsg <> "" Then
                    .Refresh
                    MsgBox "�䷽""" & rsInput!���� & """������ҩƷ�ѳ�����������ƥ�䣺" & _
                        vbCrLf & vbCrLf & vbTab & strMsg & vbCrLf & vbCrLf & "��ЩҩƷ������������䷽�С�", vbInformation, gstrSysName
                    .Refresh
                End If
            End If
        
            '����������ȡһ��ȱʡ����ҩ�÷�,ͬʱ�ж��Ƿ�����ҩ�÷�����
            lng�÷�ID = Getȱʡ�÷�ID(4, 1)
            If lng�÷�ID = 0 Then
                .Refresh
                MsgBox "û�п��õ���ҩ��(��)��,���ȵ�������Ŀ���������ã�", vbInformation, gstrSysName
                .Refresh: Exit Function
            End If
            '��ҩ�÷�ȱʡ����һ����ͬ
            If RowIn�䷽��(lngPreRow) Then
                lng�÷�ID = Val(.TextMatrix(lngPreRow, COL_������ĿID))
            End If
        End If
        
        '������ҩ����ҩ;���ж�
        If InStr(",5,6,", rsInput!���ID) > 0 Then
'            '����������ȡһ��ȱʡ�ĸ�ҩ;��,ͬʱ�ж��Ƿ��и�ҩ;������
'            lng�÷�ID = Getȱʡ�÷�ID(2, 1)
'            If lng�÷�ID = 0 Then
'                .Refresh
'                MsgBox "û�п��õĸ�ҩ;��,���ȵ�������Ŀ���������ã�", vbInformation, gstrSysName
'                .Refresh: Exit Function
'            End If
            '��ҩ;��ȱʡ����һ������ͬ���͵���ͬ
            If lngPreRow <> -1 And Not IsNull(rsInput!ҩƷ����) Then
                If InStr(",5,6,", .TextMatrix(lngPreRow, COL_���)) > 0 And .TextMatrix(lngPreRow, COL_ҩƷ����) = Nvl(rsInput!ҩƷ����) Then
                    i = .FindRow(CLng(.TextMatrix(lngPreRow, COL_���ID)), lngPreRow + 1)
                    lng�÷�ID = Val(.TextMatrix(i, COL_������ĿID))
                End If
            End If
        End If
        
        '������ҩ������������
        If InStr(",5,6,", rsInput!���ID) > 0 And gint�����Ǽ���Ч���� <> 0 Then
            str���� = Check��������(rsInput!������ĿID, rsInput!����, lngƤ��ID)
            If str���� <> "" Then
                .Refresh
                If MsgBox(str����, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    .Refresh: Exit Function
                End If
            End If
            
            '�Զ����Ƥ��
            If lngƤ��ID <> 0 Then
                '�ȼ���Ƿ����и�Ƥ��(����Чʱ�����ֹ����Զ���ӵ�,��������Ƥ��)
                i = .FixedRows - 1
                Do
                    i = .FindRow(CStr(lngƤ��ID), i + 1, COL_������ĿID)
                    If i <> -1 Then
                        If Not .RowHidden(i) Then
                            If Int(CDate(.Cell(flexcpData, i, COL_��ʼʱ��))) >= Int(zlDatabase.Currentdate - gint�����Ǽ���Ч����) Then
                                blnƤ���� = True: Exit Do '��¼������־,��ǰ��������ɺ�������
                            End If
                        End If
                    End If
                Loop Until i = -1
            End If
        End If
        
        '������ҩ��һ����ҩ���ж�
        blnGroup = (RowInһ����ҩ(lngRow) Or tbr.Buttons("һ��").Value = tbrPressed) And Not blnByƤ��
        If blnGroup Then
            If rsInput!���ID = "9" Then
                .Refresh
                MsgBox "������һ����ҩ��ҩƷ��ֱ��������׷�����", vbInformation, gstrSysName
                .Refresh: Exit Function
            End If
        
            If .RowData(lngRow) = 0 Then
                'һ����ҩ�еĴ�������У�ֻ�в�����һ����ҩ���м�,�����Զ���Ϊһ����ҩ
                lngGroupRow = lngPreRow
            Else
                'һ����ҩ�е�ҩƷ�У������ǵ�һ�л����һ��
                If InStr(",5,6,", .TextMatrix(lngPreRow, COL_���)) > 0 _
                    And Val(.TextMatrix(lngPreRow, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                    lngGroupRow = lngPreRow
                Else
                    lngGroupRow = lngNextRow
                End If
            End If
            
            'һ����ҩ��,��𣬱�����ͬ
            If Decode(rsInput!���ID, "5", "Y", "6", "Y", "N") <> Decode(.TextMatrix(lngGroupRow, COL_���), "5", "Y", "6", "Y", "N") Then
                .Refresh
                MsgBox "����һ����ҩ��ҩƷ���붼Ϊ����ҩ���г�ҩ��", vbInformation, gstrSysName
                .Refresh: Exit Function
            End If
            
            i = .FindRow(CLng(.TextMatrix(lngGroupRow, COL_���ID)), lngGroupRow + 1)
            lng�÷�ID = Val(.TextMatrix(i, COL_������ĿID)) 'һ����ҩ�ĸ�ҩ;����ͬ
            
            '���һ����ҩ�ĵĸ�ҩ;���Ƿ��ʺ��ڵ�ǰ����ҩƷ(��һ����ҩ��ȱʡ�÷������뺯���������жϴ���)
            If Not Check�����÷�(lng�÷�ID, rsInput!������ĿID, 1) Then
                .Refresh
                MsgBox "һ���ĸ�ҩ;��Ϊ""" & .TextMatrix(i, COL_ҽ������) & """���������ڵ�ǰ����ҩƷ��", vbInformation, gstrSysName
                .Refresh: Exit Function
            End If
        End If
    
        '������Ŀ
        If rsInput!���ID = "9" Then
            If GetGroupCount(rsInput!������ĿID, 1) = 0 Then
                .Refresh
                MsgBox """" & rsInput!���� & """��һ�����׷�������û��������Ч�������Ŀ��" & vbCrLf & "���ȵ�������Ŀ���������á�", vbInformation, gstrSysName
                .Refresh: Exit Function
            End If
            strExtData = frmSchemeSelect.ShowMe(Me, rsInput!������ĿID, 1)
            If strExtData = "" Then .Refresh: Exit Function
        End If
    
        '��Ҫ����������ݵ�һЩ��Ŀ
        '---------------------------------------------------------------------------------------------------------------
        intType = -1
        If rsInput!���ID = "D" And Nvl(GetItemField("������ĿĿ¼", rsInput!������ĿID, "�����Ŀ"), 0) = 1 Then
            '��������Ŀ
            intType = 0
        ElseIf rsInput!���ID = "F" Then
            '��������Ҫ����������Ŀ������ѡ�񸽼�����
            intType = 1
        ElseIf InStr(",7,8,", rsInput!���ID) > 0 Then
            '��ҩ�䷽(��ζ��ҩ���䷽����)
            intType = 2
        ElseIf rsInput!���ID = "C" Then
            '����һ���ɼ��Ķ��������Ŀ������걾
            intType = 4
            strExtData = rsInput!������ĿID & ";" & Nvl(rsInput!���)
        End If
        If intType <> -1 Then
            frmAdviceEditEx.mstrPrivs = mstrPrivs
            frmAdviceEditEx.mlngHwnd = txtҽ������.Hwnd
            frmAdviceEditEx.mintType = intType
            frmAdviceEditEx.mint��Ч = 1
            frmAdviceEditEx.mstr�Ա� = mstr�Ա�
            frmAdviceEditEx.mint������� = 1 '����
            frmAdviceEditEx.mlng��ĿID = IIF(rsInput!���ID = "C", 0, rsInput!������ĿID)
            frmAdviceEditEx.mstrExtData = IIF(rsInput!���ID = "C", strExtData, "") '��������Ŀ
            
            frmAdviceEditEx.mblnҽ�� = InStr(",1,2,", mstr������) > 0 And mstr������ <> ""
            
            On Error Resume Next
            frmAdviceEditEx.Show 1, Me
            On Error GoTo errH
            
            If Not frmAdviceEditEx.mblnOK Then Exit Function
            strExtData = frmAdviceEditEx.mstrExtData
        End If
    
        '�޸�������Ŀʱ,��ɾ����ǰҽ��������
        '---------------------------------------------------------------------------------------------------------------
        If .RowData(lngRow) <> 0 Then
            If InStr(",5,6,", .TextMatrix(lngRow, COL_���)) > 0 Then
                '����ҩ���г�ҩ
                If Not blnGroup Then
                    '������ҩɾ����ҩ;����,�������ǰ��
                    i = .FindRow(CLng(.TextMatrix(lngRow, COL_���ID)), lngRow + 1)
                    Call DeleteRow(i)
                    Call DeleteRow(lngRow, True)
                Else
                    'һ���ҩʱ,ֻ�����ǰ��
                    Call DeleteRow(lngRow, True)
                End If
            ElseIf InStr(",D,F,", .TextMatrix(lngRow, COL_���)) > 0 Then
                '��������Ŀ��������Ŀ
                'ɾ����λ�л�����������(��������,������Ŀ)
                Call Delete�������(lngRow)
                '�����ǰ��
                Call DeleteRow(lngRow, True)
            ElseIf RowIn�䷽��(lngRow) Then
                '��ҩ�䷽��˳��(���)Ҫ������ϸ����
                'ɾ�����ζҩ���巨��:ɾ��֮�����¶�λ�ĵ�ǰ��
                lngRow = Delete��ҩ�䷽(lngRow)
                '�����ǰ��(��ҩ�÷���)
                Call DeleteRow(lngRow, True)
            ElseIf RowIn������(lngRow) Then
                'ɾ��������Ŀ��:ɾ��֮�����¶�λ�ĵ�ǰ��
                lngRow = Delete�������(lngRow)
                '�����ǰ��(�ɼ�������)
                Call DeleteRow(lngRow, True)
            Else
                '������Ŀֱ�������ǰ������
                Call DeleteRow(lngRow, True)
            End If
        End If
        
        '��ǰ������ҽ��
        '---------------------------------------------------------------------------------------------------------------
        If InStr(",7,8,", rsInput!���ID) > 0 Then
            '��ҩ�䷽(��ζ��ҩ���䷽����):����֮�����¶�λ�ĵ�ǰ��
            lngRow = AdviceSet��ҩ�䷽(rsInput!������ĿID, lngRow, lng�÷�ID, strExtData)
        ElseIf rsInput!���ID = "9" Then
            '����ҽ����Ҫ�ֽ�Ϊ�����Ŀ����
            Call AdviceSet������Ŀ(rsInput!������ĿID, lngRow, strExtData)
        ElseIf rsInput!���ID = "C" Then
            '�������
            lngRow = AdviceSet�������(lngRow, lng�÷�ID, strExtData)
        Else
            '�С�����ҩ�����(���)������(���)��������������Ŀ
            Call AdviceSet������Ŀ(rsInput, lngRow, lng�÷�ID, lngGroupRow, strExtData)
            
            '�Զ�����һ����ҩ
            If InStr(",5,6,", rsInput!���ID) > 0 Then
                If Not RowInһ����ҩ(lngRow) Then
                    If tbr.Buttons("һ��").Value = tbrPressed Then
                        '�ֹ�ʹһ����ҩ
                        Call MergeRow(lngPreRow, lngRow) '����������ʾ��ǰ�е�����,������ǿ��RowChange
                    ElseIf lngPreRow <> -1 Then
                        '�Զ�ʹһ����ҩ
                        If .TextMatrix(lngPreRow, COL_���) = rsInput!���ID Then
                            If RowInһ����ҩ(lngPreRow) And RowCanMerge(lngPreRow, lngRow) And GetNextRow(lngRow) = -1 Then
                                tbr.Buttons("һ��").Value = tbrPressed
                                Call MergeRow(lngPreRow, lngRow, False)
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
        '������ҩ�ɳ�ҩʱ�Զ�����Ƥ����:����֮���Զ�λ�ڵ�ǰҩƷ
        If lngƤ��ID <> 0 And Not blnƤ���� Then
            Call AdviceSet��������(.Row, lngƤ��ID) 'ע���õ�ǰ��,��Ϊһ��֮��λ�ı�
        End If
        
        '�����Զ������и�
        Call .AutoSize(COL_ҽ������)
    End With
    mblnNoSave = True '���Ϊδ����
    AdviceInput = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub MergeRow(ByVal lngRow1 As Long, ByVal lngRow2 As Long, Optional ByVal blnCheck As Boolean = True)
'���ܣ�����������Ϊһ����ҩ
'������lngRow1=ǰ����,���ܱ����Ѿ�����һ����ҩ
'      lngRow2=��ǰ��
'˵����������ɺ�,����Զ�λ��ԭlngRow2�ĵ�ǰ��
    Dim lngBegin As Long, lngEnd As Long
    Dim blnDo As Boolean, lngTmp As Long
    
    With vsAdvice
        If blnCheck Then
            blnDo = RowCanMerge(lngRow1, lngRow2)
        Else
            blnDo = True
        End If
        If blnDo Then
            mblnRowChange = False: .Redraw = flexRDNone
            lngTmp = .RowData(lngRow2) '��¼���ٶ�λ����ǰ��
            '��ȡ��֮ǰ��һ����ҩ
            If RowInһ����ҩ(lngRow1) Then
                Call Getһ����ҩ��Χ(Val(.TextMatrix(lngRow1, COL_���ID)), lngBegin, lngEnd)
                Call AdviceSet������ҩ(lngBegin, lngEnd)
                lngRow1 = lngBegin
                lngRow2 = .FindRow(lngTmp, lngBegin + 1)
            End If
            Call AdviceSetһ����ҩ(lngRow1, lngRow2)
            lngRow2 = .FindRow(lngTmp, lngBegin + 1)
            .Row = lngRow2
            mblnRowChange = True: .Redraw = flexRDDirect
        End If
    End With
End Sub

Private Sub SplitRow(ByVal lngRow As Long)
'���ܣ���ָ���д�һ����ҩ�ж�������(����һ����ҩ�������ٰ�������)
'������lngRow=��ǰ��,��Ϊһ����ҩ�е����һҩƷ��
'˵����������ɺ�,����Զ�λ��ԭlngRow�ĵ�ǰ��
    Dim lngBegin As Long, lngEnd As Long, lngTmp As Long
    
    With vsAdvice
        mblnRowChange = False: .Redraw = flexRDNone
        lngTmp = .RowData(lngRow) '��¼���ڻָ���λ��ǰ��
        Call Getһ����ҩ��Χ(Val(.TextMatrix(lngRow, COL_���ID)), lngBegin, lngEnd)
        
        '��ȡ��������һ����ҩ
        Call AdviceSet������ҩ(lngBegin, lngEnd)
        
        '�����ó�����������Ϊһ����ҩ
        lngRow = .FindRow(lngTmp, lngBegin + 1)
        lngEnd = GetPreRow(lngRow)
        Call AdviceSetһ����ҩ(lngBegin, lngEnd)
        
        '�ָ���ǰ��
        lngRow = .FindRow(lngTmp, lngBegin + 1)
        .Row = lngRow
        mblnRowChange = True: .Redraw = flexRDDirect
    End With
End Sub

Public Sub AdviceSet����ҽ��(ByVal lng����ID As Long, ByVal str�Һŵ� As String, ByVal strIDs As String, Optional ByVal blnHistory As Boolean)
'���ܣ�����ָ�����˵�ָ��ҽ��������Ϊ��ҽ��
'˵�����ɹ��ⲿ����,����֮ǰ��������ҽ����
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, bln�䷽ As Boolean
    Dim lngBegin As Long, lngEnd As Long
    Dim curDate As Date, blnHide As Boolean
    Dim lng��������ID As Long, lng���ID As Long
    Dim lng��� As Long, intCount As Integer
    Dim lngRow As Long, i As Long, j As Long
    
    Dim lng��ҩ��ID As Long, lng��ҩ��ID As Long, lng��ҩ��ID As Long
    Dim strҩ��IDs As String
    
    Screen.MousePointer = 11
    
    On Error GoTo errH
    
    strSQL = _
        " Select A.ID,A.���ID,Nvl(A.Ӥ��,0) as Ӥ��,A.���,A.ҽ����Ч," & _
        " A.ҽ��״̬,A.�������,A.������ĿID,B.����,A.�걾��λ,A.�շ�ϸĿID," & _
        " A.��ʼִ��ʱ��,A.ҽ������,A.ҽ������,A.��������,A.����,A.�ܸ�����,B.���㵥λ," & _
        " A.ִ��Ƶ��,A.Ƶ�ʴ���,A.Ƶ�ʼ��,A.�����λ,B.���㷽ʽ,B.ִ��Ƶ��,B.��������," & _
        " B.�Ƽ�����,A.ִ��ʱ�䷽��,A.ִ������,A.ִ�п���ID,A.��������ID,A.����ҽ��,A.����ʱ��," & _
        " A.������־,C.�������,C.ҩƷ����,B.¼������,C.��������,C.����ְ��," & _
        " D.����ϵ��,D.�����װ,D.���ﵥλ,D.�ɷ����,A.����ID" & _
        " From ����ҽ����¼ A,������ĿĿ¼ B,ҩƷ���� C,ҩƷ��� D" & _
        " Where Nvl(A.ҽ����Ч,0)=1 And A.������ĿID=B.ID" & _
        " And A.������ĿID=C.ҩ��ID(+) And A.�շ�ϸĿID=D.ҩƷID(+)" & _
        " And A.����ID+0=[1] And A.�Һŵ�=[2]" & _
        " And Instr([3],','||Nvl(A.���ID,A.ID)||',')>0" & _
        " Order by Ӥ��,���"
    If blnHistory Then
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, str�Һŵ�, "," & strIDs & ",")
    On Error GoTo 0
    
    If Not rsTmp.EOF Then
        lngBegin = vsAdvice.Row '��ʼ������
        lng��� = GetCurRow���(lngBegin) '��ʼ���
        intCount = 0 '�Ѿ����õ�����
        curDate = zlDatabase.Currentdate
        lng��������ID = Get��������ID(UserInfo.ID, mlng���˿���id, 1)
        
        mblnRowChange = False
        With vsAdvice
            .Redraw = flexRDNone
            For i = lngBegin To rsTmp.RecordCount + lngBegin - 1
                If i > lngBegin Then .AddItem "", i

                bln�䷽ = False
                
                .RowData(i) = -1 * rsTmp!ID
                If Not IsNull(rsTmp!���ID) Then
                    .TextMatrix(i, COL_���ID) = -1 * rsTmp!���ID
                End If
                .TextMatrix(i, COL_���) = lng��� + intCount
                
                .TextMatrix(i, COL_EDIT) = 1 '����
                .Cell(flexcpData, i, COL_EDIT) = CStr(lng����ID & "," & str�Һŵ�) '��¼��صĸ�����Ŀ
                .TextMatrix(i, COL_״̬) = 1 '�¿�
                .TextMatrix(i, COL_Ӥ��) = cboӤ��.ListIndex
                .TextMatrix(i, COL_���) = rsTmp!�������
                .TextMatrix(i, COL_������ĿID) = rsTmp!������ĿID
                .TextMatrix(i, COL_����) = rsTmp!����
                .TextMatrix(i, COL_�걾��λ) = Nvl(rsTmp!�걾��λ)
                .TextMatrix(i, COL_�շ�ϸĿID) = Nvl(rsTmp!�շ�ϸĿID)
                .TextMatrix(i, COL_ҽ������) = Nvl(rsTmp!ҽ������)
                .TextMatrix(i, COL_ҽ������) = Nvl(rsTmp!ҽ������)
                
                .TextMatrix(i, COL_�Ƽ�����) = Nvl(rsTmp!�Ƽ�����, 0)
                .TextMatrix(i, COL_���㷽ʽ) = Nvl(rsTmp!���㷽ʽ, 0)
                .TextMatrix(i, COL_Ƶ������) = Nvl(rsTmp!ִ��Ƶ��, 0)
                .TextMatrix(i, COL_��������) = Nvl(rsTmp!��������)
                .TextMatrix(i, COL_�������) = Nvl(rsTmp!�������)
                .TextMatrix(i, COL_ҩƷ����) = Nvl(rsTmp!ҩƷ����)
                If InStr(",5,6,7,", rsTmp!�������) > 0 Then
                    .TextMatrix(i, COL_��������) = Nvl(rsTmp!��������)
                Else
                    .TextMatrix(i, COL_��������) = Nvl(rsTmp!¼������)
                End If
                .TextMatrix(i, COL_����ְ��) = Nvl(rsTmp!����ְ��)
                .TextMatrix(i, COL_����ϵ��) = Nvl(rsTmp!����ϵ��)
                .TextMatrix(i, COL_�����װ) = Nvl(rsTmp!�����װ)
                .TextMatrix(i, COL_���ﵥλ) = Nvl(rsTmp!���ﵥλ)
                If Not IsNull(rsTmp!����ϵ��) Then
                    .TextMatrix(i, COL_�ɷ����) = Nvl(rsTmp!�ɷ����, 0)
                End If
                
                If IsDate(txt��ʼʱ��.Text) Then
                    .TextMatrix(i, COL_��ʼʱ��) = Format(txt��ʼʱ��.Text, "MM-dd HH:mm")
                    .Cell(flexcpData, i, COL_��ʼʱ��) = txt��ʼʱ��.Text
                End If
                
                .TextMatrix(i, COL_Ƶ��) = Nvl(rsTmp!ִ��Ƶ��)
                .TextMatrix(i, COL_Ƶ�ʴ���) = Nvl(rsTmp!Ƶ�ʴ���)
                .TextMatrix(i, COL_Ƶ�ʼ��) = Nvl(rsTmp!Ƶ�ʼ��)
                .TextMatrix(i, COL_�����λ) = Nvl(rsTmp!�����λ)
                .TextMatrix(i, COL_ִ��ʱ��) = Nvl(rsTmp!ִ��ʱ�䷽��)
                
                .TextMatrix(i, COL_ִ������) = Nvl(rsTmp!ִ������, 0)
                
                '����ִ�п���
                If rsTmp!������� = "Z" Then
                    .TextMatrix(i, COL_ִ�п���ID) = Nvl(rsTmp!ִ�п���ID)
                ElseIf InStr(",0,5,", Nvl(rsTmp!ִ������, 0)) = 0 Then
                    If Nvl(rsTmp!ִ�п���ID, 0) <> 0 Then
                        If InStr(",5,6,7,", rsTmp!�������) > 0 Then
                            strҩ��IDs = Get����ҩ��IDs(rsTmp!�������, rsTmp!������ĿID, Nvl(rsTmp!�շ�ϸĿID, 0), mlng���˿���id, 1)
                            If InStr("," & strҩ��IDs & ",", "," & rsTmp!ִ�п���ID & ",") > 0 Then
                                .TextMatrix(i, COL_ִ�п���ID) = Nvl(rsTmp!ִ�п���ID, 0)
                            End If
                        ElseIf Val(.TextMatrix(i, COL_ִ������)) = 4 Then
                            '4-ָ������ʱ��ȡ,�����Ĺ̶�����
                            .TextMatrix(i, COL_ִ�п���ID) = Nvl(rsTmp!ִ�п���ID, 0)
                        End If
                    End If
                    If Val(.TextMatrix(i, COL_ִ�п���ID)) = 0 Then
                        'ҩƷ�������������ͬ
                        If rsTmp!������� = "5" Then
                            If lng��ҩ��ID = 0 Then
                                lng��ҩ��ID = Get����ִ�п���ID(mlng����ID, 0, rsTmp!�������, rsTmp!������ĿID, Nvl(rsTmp!�շ�ϸĿID, 0), 4, mlng���˿���id, 0, 1, 1, True)
                            End If
                            .TextMatrix(i, COL_ִ�п���ID) = lng��ҩ��ID
                        ElseIf rsTmp!������� = "6" Then
                            If lng��ҩ��ID = 0 Then
                                lng��ҩ��ID = Get����ִ�п���ID(mlng����ID, 0, rsTmp!�������, rsTmp!������ĿID, Nvl(rsTmp!�շ�ϸĿID, 0), 4, mlng���˿���id, 0, 1, 1, True)
                            End If
                            .TextMatrix(i, COL_ִ�п���ID) = lng��ҩ��ID
                        ElseIf rsTmp!������� = "7" Then
                            If lng��ҩ��ID = 0 Then
                                lng��ҩ��ID = Get����ִ�п���ID(mlng����ID, 0, rsTmp!�������, rsTmp!������ĿID, Nvl(rsTmp!�շ�ϸĿID, 0), 4, mlng���˿���id, 0, 1, 1, True)
                            End If
                            .TextMatrix(i, COL_ִ�п���ID) = lng��ҩ��ID
                        Else
                            .TextMatrix(i, COL_ִ�п���ID) = Get����ִ�п���ID(mlng����ID, 0, rsTmp!�������, rsTmp!������ĿID, 0, Nvl(rsTmp!ִ������, 0), mlng���˿���id, lng��������ID, 1, 1)
                        End If
                    End If
                End If
                
                If rsTmp!������� = "E" Then
                    If Nvl(rsTmp!���ID, 0) = 0 And Val(.TextMatrix(i - 1, COL_���ID)) = -1 * rsTmp!ID Then
                        If InStr(",5,6,", .TextMatrix(i - 1, COL_���)) > 0 Then
                            '��ǰ��¼�ǳ�ҩ�ĸ�ҩ;��,������һ����ҩ��
                            For j = i - 1 To .FixedRows Step -1
                                If Val(.TextMatrix(j, COL_���ID)) = -1 * rsTmp!ID Then
                                    '��ʾ��ҩ;��
                                    .TextMatrix(j, COL_�÷�) = rsTmp!����
                                Else
                                    Exit For
                                End If
                            Next
                        ElseIf InStr(",E,7,", .TextMatrix(i - 1, COL_���)) > 0 Then
                            '��ǰ��¼����ҩ�䷽���÷�,���䷽��ʾ��
                            .TextMatrix(i, COL_�÷�) = rsTmp!����
                            bln�䷽ = True
                        ElseIf .TextMatrix(i - 1, COL_���) = "C" Then
                            .TextMatrix(i, COL_�÷�) = rsTmp!����
                        End If
                    ElseIf Not IsNull(rsTmp!���ID) Then
                        '��ǰ��¼����ҩ�䷽�巨��
                        bln�䷽ = True
                    End If
                ElseIf rsTmp!������� = "7" Then
                    bln�䷽ = True
                End If
                
                '����
                .TextMatrix(i, COL_����) = FormatEx(Nvl(rsTmp!��������), 5)
                If InStr(",5,6,7,", rsTmp!�������) > 0 Or Nvl(rsTmp!���㷽ʽ, 0) <> 3 Then
                    .TextMatrix(i, COL_������λ) = Nvl(rsTmp!���㵥λ)
                End If
                
                '����
                .TextMatrix(i, COL_����) = Nvl(rsTmp!����, 0)
                
                '����
                If InStr(",5,6,", rsTmp!�������) > 0 Then
                    '��ҩ����������,�����۵�λ���,���ﵥλ��ʾ
                    If Not IsNull(rsTmp!�ܸ�����) And Not IsNull(rsTmp!�����װ) Then
                        .TextMatrix(i, COL_����) = FormatEx(rsTmp!�ܸ����� / rsTmp!�����װ, 5)
                    End If
                    .TextMatrix(i, COL_������λ) = Nvl(rsTmp!���ﵥλ)
                Else
                    '�����������ҩ����������
                    If Not IsNull(rsTmp!�ܸ�����) Then
                        .TextMatrix(i, COL_����) = rsTmp!�ܸ�����
                    End If
                    If bln�䷽ Then
                        .TextMatrix(i, COL_������λ) = "��" '��ҩ�䷽������λΪ"��"
                    Else
                        .TextMatrix(i, COL_������λ) = Nvl(rsTmp!���㵥λ)
                    End If
                End If
                
                .TextMatrix(i, COL_��־) = 0
                .TextMatrix(i, COL_����ҽ��) = UserInfo.����
                .TextMatrix(i, COL_��������ID) = lng��������ID
                .TextMatrix(i, COL_����ʱ��) = Format(curDate, "MM-dd HH:mm")
                .Cell(flexcpData, i, COL_����ʱ��) = Format(curDate, "yyyy-MM-dd HH:mm")
                
                '���龫ҩƷ��ʶ:��ҩ�䷽�����ζ��ҩ������
                If InStr(",5,6,", rsTmp!�������) > 0 And Not IsNull(rsTmp!�������) Then
                    If InStr(",����ҩ,����ҩ,����ҩ,", rsTmp!�������) > 0 Then
                        .Cell(flexcpFontBold, i, COL_ҽ������) = True
                    End If
                End If
                
                lngEnd = i
                intCount = intCount + 1
                
                rsTmp.MoveNext
            Next
            
            '�����µ�ҽ��ID
            For i = lngBegin To lngEnd
                lng���ID = .RowData(i)
                .RowData(i) = zlDatabase.GetNextId("����ҽ����¼")
                For j = i - 1 To lngBegin Step -1
                    If Val(.TextMatrix(j, COL_���ID)) = lng���ID Then
                        .TextMatrix(j, COL_���ID) = .RowData(i)
                    Else
                        Exit For
                    End If
                Next
                For j = i + 1 To lngEnd
                    If Val(.TextMatrix(j, COL_���ID)) = lng���ID Then
                        .TextMatrix(j, COL_���ID) = .RowData(i)
                    Else
                        Exit For
                    End If
                Next
            Next
            
            '������Ӱ���е����
            Call AdviceSetҽ�����(lngEnd + 1, intCount)
            
            '��ʾ/������
            lngRow = 0
            For i = lngBegin To lngEnd
                blnHide = False
                If .TextMatrix(i, COL_���) = "E" And Val(.TextMatrix(i, COL_���ID)) = 0 Then
                    If Val(.TextMatrix(i - 1, COL_���ID)) = .RowData(i) _
                        And InStr(",5,6,", .TextMatrix(i - 1, COL_���)) > 0 Then
                        blnHide = True
                    End If
                End If
                If InStr(",F,G,D,7,E,C,", .TextMatrix(i, COL_���)) > 0 _
                    And Val(.TextMatrix(i, COL_���ID)) <> 0 Then
                    blnHide = True
                End If
                                
                .RowHidden(i) = blnHide
                If Not blnHide And lngRow = 0 Then lngRow = i
                
                '����ҽ�����ݵı仯
                If Not .RowHidden(i) Then
                    '����ʱ��ʼʱ��仯
                    txt��ʼʱ��.Tag = "1"
                    If AdviceTextChange(i) Then
                        .TextMatrix(i, COL_ҽ������) = AdviceTextMake(i)
                    End If
                    txt��ʼʱ��.Tag = ""
                End If
                
                'Ԥ�ȼ������Ƶ���
                If Not .RowHidden(i) And .TextMatrix(i, COL_����) = "" Then
                    .TextMatrix(i, COL_����) = GetItemPrice(i)
                End If
            Next
            
            'ͼ�����:����Ϊ�ж���,��Ȼ���߿�ʱ����������
            .Cell(flexcpPictureAlignment, .FixedRows, 0, .Rows - 1, .FixedCols - 1) = 4
            
            .Row = lngRow: .Col = COL_ҽ������
            
            Call .AutoSize(COL_ҽ������)
            .Redraw = flexRDDirect
        End With
        mblnRowChange = True
        mblnNoSave = True '���Ϊδ����
    End If

    Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
    Call CalcAdviceMoney '��ʾ�¿�ҽ�����

    Screen.MousePointer = 0
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub AdviceSet������Ŀ(ByVal lng����ID As Long, ByVal lngRow As Long, Optional ByVal str��� As String)
'���ܣ����������Ŀ(����һ����ҩ,������,��������,��ҩ�䷽)
'������lngRow=�յ�������(�����ǲ��������,����λ��һ����ҩ�м�)
    Dim rsItems As New ADODB.Recordset
    Dim rs��� As New ADODB.Recordset
    Dim rs�Ƴ� As New ADODB.Recordset
    Dim strSQL As String, i As Long, j As Long
    
    Dim lngCurRow As Long, intCount As Integer, lng��� As Long
    Dim lngPreRow As Long, vCurDate As Date, lngTmp As Long
    Dim intƵ�ʴ��� As Integer, intƵ�ʼ�� As Integer, str�����λ As String
    Dim bln��ҩ;�� As Boolean, bln�ɼ����� As Boolean, intƵ������ As Integer
    Dim bln��ҩ�÷� As Boolean, bln��ҩ�巨 As Boolean, bln�䷽ As Boolean
    Dim lng��ҩ��ID As Long, lng��ҩ��ID As Long, lng��ҩ��ID As Long
    Dim lng���ID As Long, int���÷�Χ As Integer, strƵ�� As String
    Dim strҽ�� As String, lngҽ��ID As Long, blnFirst As Boolean
    Dim lng���� As Long, vBookMark As Variant, strҩ��IDs As String
    Dim sng���� As Single, strSQL��� As String
    
    On Error GoTo errH
    Screen.MousePointer = 11
    Me.Refresh
    
    '������Ź��˴�
    If str��� <> "" Then
        If Left(str���, 1) = "+" Then
            strSQL��� = " And Instr([2],','||A.���||',')>0"
        ElseIf Left(str���, 1) = "-" Then
            strSQL��� = " And Instr([2],','||A.���||',')=0"
        End If
    End If
    
    'ҩƷ�����Ϣ:��Ȼ�����շ�ϸĿID,����ǰ������û��
    strSQL = "Select A.���,B.ҩ��ID,B.ҩƷID,B.����ϵ��,B.�����װ,B.���ﵥλ," & _
        " B.�ɷ����,C.����,Nvl(D.����,C.����) as ����,C.���,C.����" & _
        " From ������Ŀ��� A,ҩƷ��� B,�շ���ĿĿ¼ C,�շ���Ŀ���� D" & _
        " Where A.������ĿID=B.ҩ��ID And B.ҩƷID=C.ID" & _
        " And C.ID=D.�շ�ϸĿID(+) And D.����(+)=1 And D.����(+)=[3]" & _
        " And A.�������ID=[1]" & strSQL��� & _
        " Order by A.���,C.����"
    Set rs��� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, "," & Mid(str���, 2) & ",", IIF(gbln��Ʒ��, 3, 1))
    
    '��ҩ�Ƴ���Ϣ(���������ֱ�Ӷ�Ӧ�䷽,��ҩȡ�����Ƴ�)
    strSQL = "Select Distinct A.������ĿID,C.�Ƴ�" & _
        " From ������Ŀ��� A,������ĿĿ¼ B,�����÷����� C" & _
        " Where A.������ĿID=B.ID And B.��� IN('5','6')" & _
        " And A.������ĿID=C.��ĿID And A.�������ID=[1]" & strSQL���
    Set rs�Ƴ� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, "," & Mid(str���, 2) & ",")
    
    '��������к�Ӧ����ҽ���༭ʱ�Ĵ���һ��
    '�ſ�ִ��Ƶ��Ϊ�����Եĳ���(���ַ�������,ֻȡ����),���з���תΪ��������
    strSQL = "Select 1 as ��Ч,A.���,A.������,A.������ĿID,A.�շ�ϸĿID,A.�ܸ�����,A.��������," & _
        " A.ҽ������,A.ִ��Ƶ��,A.Ƶ�ʴ���,A.Ƶ�ʼ��,A.�����λ,A.ִ�п���ID,B.���,B.����," & _
        " B.���㵥λ,Nvl(A.�걾��λ,B.�걾��λ) as �걾��λ,A.ʱ�䷽��,Nvl(A.ִ������,B.ִ�п���) as ִ������," & _
        " B.�Ƽ�����,B.��������,B.���㷽ʽ,B.ִ��Ƶ��,B.¼������,C.��������,C.����ְ��,C.�������,C.ҩƷ����" & _
        " From ������Ŀ��� A,������ĿĿ¼ B,ҩƷ���� C" & _
        " Where A.������ĿID=B.ID And A.������ĿID=C.ҩ��ID(+)" & _
        " And A.��Ч=1 And A.�������ID=[1]" & strSQL��� & _
        " Order by A.���"
    Set rsItems = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, "," & Mid(str���, 2) & ",")
    With vsAdvice
        mblnRowChange = False
        .Redraw = flexRDNone
        
        lngPreRow = GetPreRow(lngRow) 'ǰһ������
        intCount = 0 '�Ѿ����õ�����
        lng��� = GetCurRow���(lngRow) '��ʼ���
        vCurDate = zlDatabase.Currentdate
        
        For i = 1 To rsItems.RecordCount
            lngCurRow = lngRow + intCount
            If lngCurRow > lngRow Then .AddItem "", lngCurRow
             
            '��¼���ID
            .RowData(lngCurRow) = -1 * rsItems!���
            If Not IsNull(rsItems!������) Then
                .TextMatrix(lngCurRow, COL_���ID) = -1 * rsItems!������
            End If
            
            .TextMatrix(lngCurRow, COL_EDIT) = 1 '������
            .Cell(flexcpData, lngCurRow, COL_EDIT) = lng����ID '��¼��صĳ�����Ŀ
            
            .TextMatrix(lngCurRow, COL_Ӥ��) = cboӤ��.ListIndex
            .TextMatrix(lngCurRow, COL_���) = lng��� + intCount
            .TextMatrix(lngCurRow, COL_״̬) = 1 '�¿�
            .TextMatrix(lngCurRow, COL_���) = rsItems!���
            
            If IsDate(txt��ʼʱ��.Text) Then
                .TextMatrix(lngCurRow, COL_��ʼʱ��) = Format(txt��ʼʱ��.Text, "MM-dd HH:mm")
                .Cell(flexcpData, lngCurRow, COL_��ʼʱ��) = Format(txt��ʼʱ��.Text, "yyyy-MM-dd HH:mm")
            End If
            
            .TextMatrix(lngCurRow, COL_������ĿID) = rsItems!������ĿID
            .TextMatrix(lngCurRow, COL_����) = rsItems!����
            .TextMatrix(lngCurRow, COL_�걾��λ) = Nvl(rsItems!�걾��λ)
            
            '����
            .TextMatrix(lngCurRow, COL_�Ƽ�����) = Nvl(rsItems!�Ƽ�����, 0)
            .TextMatrix(lngCurRow, COL_���㷽ʽ) = Nvl(rsItems!���㷽ʽ, 0)
            .TextMatrix(lngCurRow, COL_��������) = Nvl(rsItems!��������)
            .TextMatrix(lngCurRow, COL_�������) = Nvl(rsItems!�������)
            .TextMatrix(lngCurRow, COL_ҩƷ����) = Nvl(rsItems!ҩƷ����)
            If InStr(",5,6,7,", rsItems!���) > 0 Then
                .TextMatrix(lngCurRow, COL_��������) = Nvl(rsItems!��������)
            Else
                .TextMatrix(lngCurRow, COL_��������) = Nvl(rsItems!¼������)
            End If
            .TextMatrix(lngCurRow, COL_����ְ��) = Nvl(rsItems!����ְ��)
            
            'ҩƷ�����Ϣ:�в�ҩ�϶���,��ҩ�������������λ�Զ�ƥ��
            lng���� = 0: vBookMark = 0
            If rsItems!��� = "7" Or (InStr(",5,6,", rsItems!���) > 0) Then
                If Not IsNull(rsItems!�շ�ϸĿID) Then '������ǰδ�����շ�ϸĿID
                    rs���.Filter = "ҩƷID=" & rsItems!�շ�ϸĿID
                Else
                    rs���.Filter = "ҩ��ID=" & rsItems!������ĿID
                End If
                If Not rs���.EOF Then
                    If IsNull(rsItems!�շ�ϸĿID) Then
                        'ȡ����ϵ��Ϊ��������С����������һ�����
                        If CInt(Nvl(rsItems!��������, 0)) <> 0 Then
                            Do While Not rs���.EOF
                                If rs���!����ϵ�� / rsItems!�������� = Int(rs���!����ϵ�� / rsItems!��������) Then
                                    If rs���!����ϵ�� / rsItems!�������� < lng���� Or lng���� = 0 Then
                                        vBookMark = rs���.Bookmark
                                        lng���� = rs���!����ϵ�� / rsItems!��������
                                    End If
                                End If
                                rs���.MoveNext
                            Loop
                            If vBookMark <> 0 Then rs���.Bookmark = vBookMark
                        End If
                        If rs���.EOF Then rs���.MoveFirst
                    End If
                    .TextMatrix(lngCurRow, COL_����) = Nvl(rs���!����)
                    .TextMatrix(lngCurRow, COL_�շ�ϸĿID) = rs���!ҩƷID
                    .TextMatrix(lngCurRow, COL_����ϵ��) = Nvl(rs���!����ϵ��)
                    .TextMatrix(lngCurRow, COL_�����װ) = Nvl(rs���!�����װ)
                    .TextMatrix(lngCurRow, COL_���ﵥλ) = Nvl(rs���!���ﵥλ)
                    .TextMatrix(lngCurRow, COL_�ɷ����) = Nvl(rs���!�ɷ����, 0)
                End If
            End If
                                
            '�ж��Ƿ��ض���
            bln��ҩ;�� = False: bln�ɼ����� = False
            bln��ҩ�÷� = False: bln��ҩ�巨 = False: bln�䷽ = False
            If rsItems!��� = "E" Then
                If IsNull(rsItems!������) Then
                    If Val(.TextMatrix(lngCurRow - 1, COL_���ID)) = .RowData(lngCurRow) Then
                        If InStr(",5,6,", .TextMatrix(lngCurRow - 1, COL_���)) > 0 Then
                            bln��ҩ;�� = True
                        ElseIf .TextMatrix(lngCurRow - 1, COL_���) = "C" Then
                            bln�ɼ����� = True
                        Else
                            bln��ҩ�÷� = True
                        End If
                    End If
                Else
                    bln��ҩ�巨 = True
                End If
            End If
            If rsItems!��� = "7" Or bln��ҩ�巨 Or bln��ҩ�÷� Then bln�䷽ = True
                    
            '��ȡ��ǰ��Ŀ�����÷�Χ
            If bln�ɼ����� Then
                '�ɼ������Լ�����Ŀ��Ϊ׼
                lngTmp = .FindRow(CStr(.RowData(lngCurRow)), , COL_���ID)
                intƵ������ = .TextMatrix(lngTmp, COL_Ƶ������)
            Else
                intƵ������ = Nvl(rsItems!ִ��Ƶ��, 0)
            End If
            If bln�䷽ Then
                int���÷�Χ = 2 '��ҩ�䷽(�����巨,�÷�)����ҽ
'            ElseIf bln�ɼ����� Then
'                int���÷�Χ = -1 '�����������Ŀ��ͬ:һ����
            ElseIf intƵ������ = 0 Or bln��ҩ;�� _
                Or InStr(",5,6,", .TextMatrix(lngCurRow, COL_���)) > 0 Then
                int���÷�Χ = 1 '"��ѡƵ��"���ҩ(������ҩ;��)����ҽ
            ElseIf intƵ������ = 1 Then
                int���÷�Χ = -1 'һ����
            ElseIf intƵ������ = 2 Then
                int���÷�Χ = -2 '������
            End If
                    
            'Ƶ��,Ƶ�ʴ���,Ƶ�ʼ��,�����λ
            .TextMatrix(lngCurRow, COL_Ƶ������) = intƵ������
            If Not IsNull(rsItems!ִ��Ƶ��) Then
                .TextMatrix(lngCurRow, COL_Ƶ��) = rsItems!ִ��Ƶ��
                .TextMatrix(lngCurRow, COL_Ƶ�ʴ���) = Nvl(rsItems!Ƶ�ʴ���, 0)
                .TextMatrix(lngCurRow, COL_Ƶ�ʼ��) = Nvl(rsItems!Ƶ�ʼ��, 0)
                .TextMatrix(lngCurRow, COL_�����λ) = Nvl(rsItems!�����λ)
                
'                Call GetƵ����Ϣ_����(rsItems!ִ��Ƶ��, intƵ�ʴ���, intƵ�ʼ��, str�����λ, CStr(int���÷�Χ))
'                .TextMatrix(lngCurRow, COL_Ƶ��) = rsItems!ִ��Ƶ��
'                .TextMatrix(lngCurRow, COL_Ƶ�ʴ���) = intƵ�ʴ���
'                .TextMatrix(lngCurRow, COL_Ƶ�ʼ��) = intƵ�ʼ��
'                .TextMatrix(lngCurRow, COL_�����λ) = str�����λ
            Else 'ȡȱʡ��
                Call GetȱʡƵ��(int���÷�Χ, strƵ��, intƵ�ʴ���, intƵ�ʼ��, str�����λ)
                .TextMatrix(lngCurRow, COL_Ƶ��) = strƵ��
                .TextMatrix(lngCurRow, COL_Ƶ�ʴ���) = intƵ�ʴ���
                .TextMatrix(lngCurRow, COL_Ƶ�ʼ��) = intƵ�ʼ��
                .TextMatrix(lngCurRow, COL_�����λ) = str�����λ
            End If
            
            '����
            .TextMatrix(lngCurRow, COL_����) = FormatEx(Nvl(rsItems!��������), 5)
            If InStr(",5,6,7,", rsItems!���) > 0 Or (intƵ������ = 0 And InStr(",1,2,", Nvl(rsItems!���㷽ʽ, 0)) > 0) Then
                .TextMatrix(lngCurRow, COL_������λ) = Nvl(rsItems!���㵥λ)
            End If
            
            '����
            If InStr(",5,6,", rsItems!���) > 0 Or rsItems!��� = "7" Then
                '��ҩ����(�ж�Ӧ���)����ҩ�䷽
                If InStr(",5,6,", rsItems!���) > 0 Then
                    .TextMatrix(lngCurRow, COL_������λ) = .TextMatrix(lngCurRow, COL_���ﵥλ)
                    
                    sng���� = msng����
                    If mbln���� Then
                        If .TextMatrix(lngCurRow, COL_�����λ) = "��" Then
                            If 7 > sng���� Then sng���� = 7
                        ElseIf .TextMatrix(lngCurRow, COL_�����λ) = "��" Then
                            If Val(.TextMatrix(lngCurRow, COL_Ƶ�ʼ��)) > sng���� Then
                                sng���� = Val(.TextMatrix(lngCurRow, COL_Ƶ�ʼ��))
                            End If
                        ElseIf .TextMatrix(lngCurRow, COL_�����λ) = "Сʱ" Then
                            If Val(.TextMatrix(lngCurRow, COL_Ƶ�ʼ��)) \ 24 > sng���� Then
                                sng���� = Val(.TextMatrix(lngCurRow, COL_Ƶ�ʼ��)) \ 24
                            End If
                        End If
                        If sng���� = 0 Then sng���� = 1
                    End If
                Else
                    .TextMatrix(lngCurRow, COL_������λ) = "��"
                    sng���� = 1
                End If
                
                If Not IsNull(rsItems!�ܸ�����) Then
                    If InStr(",5,6,", rsItems!���) > 0 Then
                        'ת��Ϊ���ﵥλ
                        .TextMatrix(lngCurRow, COL_����) = FormatEx(rsItems!�ܸ����� / Val(.TextMatrix(lngCurRow, COL_�����װ)), 5)
                    Else
                        .TextMatrix(lngCurRow, COL_����) = rsItems!�ܸ�����
                    End If
                Else
                    '����ȱʡ����
                    If .TextMatrix(lngCurRow, COL_Ƶ��) <> "" Then
                        If InStr(",5,6,", rsItems!���) > 0 Then
                            rs�Ƴ�.Filter = "������ĿID=" & rsItems!������ĿID
                            If Not rs�Ƴ�.EOF Then
                                If Nvl(rs�Ƴ�!�Ƴ�, 1) > sng���� Then
                                    sng���� = Nvl(rs�Ƴ�!�Ƴ�, 1)
                                End If
                            End If
                        End If
                        
                        If InStr(",5,6,", rsItems!���) > 0 Then
                            If (Val(.TextMatrix(lngCurRow, COL_����)) <> 0 _
                                And Val(.TextMatrix(lngCurRow, COL_�����װ)) <> 0 _
                                And Val(.TextMatrix(lngCurRow, COL_����ϵ��)) <> 0) Then
                                .TextMatrix(lngCurRow, COL_����) = FormatEx(CalcȱʡҩƷ����( _
                                        Val(.TextMatrix(lngCurRow, COL_����)), sng����, _
                                        Val(.TextMatrix(lngCurRow, COL_Ƶ�ʴ���)), _
                                        Val(.TextMatrix(lngCurRow, COL_Ƶ�ʼ��)), _
                                        .TextMatrix(lngCurRow, COL_�����λ), _
                                        .TextMatrix(lngCurRow, COL_ִ��ʱ��), _
                                        Val(.TextMatrix(lngCurRow, COL_����ϵ��)), _
                                        Val(.TextMatrix(lngCurRow, COL_�����װ)), _
                                        Val(.TextMatrix(lngCurRow, COL_�ɷ����))), 5)
                            End If
                        Else
                            .TextMatrix(lngCurRow, COL_����) = CalcȱʡҩƷ����(1, sng����, _
                                    Val(.TextMatrix(lngCurRow, COL_Ƶ�ʴ���)), _
                                    Val(.TextMatrix(lngCurRow, COL_Ƶ�ʼ��)), _
                                    .TextMatrix(lngCurRow, COL_�����λ))
                        End If
                    End If
                End If

                If mbln���� And InStr(",5,6,", rsItems!���) > 0 Then
                    .TextMatrix(lngCurRow, COL_����) = sng����
                End If
            ElseIf bln�䷽ Then
                '��ҩ�巨,�÷������������ҩ��ͬ(Ϊ����ʾ)
                .TextMatrix(lngCurRow, COL_����) = .TextMatrix(lngCurRow - 1, COL_����)
                .TextMatrix(lngCurRow, COL_������λ) = .TextMatrix(lngCurRow - 1, COL_������λ)
            Else
                '������������Ҫ����
                '���Ϊһ���Ի�ƴ�����ȱʡ����Ϊ1
                If Not IsNull(rsItems!�ܸ�����) Then
                    vsAdvice.TextMatrix(lngCurRow, COL_����) = rsItems!�ܸ�����
                ElseIf intƵ������ = 1 Or Nvl(rsItems!���㷽ʽ, 0) = 3 Then
                    vsAdvice.TextMatrix(lngCurRow, COL_����) = 1
                End If
                .TextMatrix(lngCurRow, COL_������λ) = Nvl(rsItems!���㵥λ)
            End If
                    
            'ִ��ʱ��(����,Ƶ��,ִ��ʱ��֮��)
            If .TextMatrix(lngCurRow, COL_Ƶ��) <> "" Then
                '�������ȱʡִ��ʱ�䷽��
                If bln��ҩ;�� Or bln��ҩ�÷� Then
                    If Not IsNull(rsItems!ʱ�䷽��) Then
                        If ExeTimeValid(rsItems!ʱ�䷽��, Val(.TextMatrix(lngCurRow, COL_Ƶ�ʴ���)), _
                            Val(.TextMatrix(lngCurRow, COL_Ƶ�ʼ��)), .TextMatrix(lngCurRow, COL_�����λ)) Then
                            .TextMatrix(lngCurRow, COL_ִ��ʱ��) = rsItems!ʱ�䷽��
                        End If
                    End If
                    If .TextMatrix(lngCurRow, COL_ִ��ʱ��) = "" Then
                        .TextMatrix(lngCurRow, COL_ִ��ʱ��) = Getȱʡʱ��(int���÷�Χ, .TextMatrix(lngCurRow, COL_Ƶ��), rsItems!������ĿID)
                    End If
                ElseIf intƵ������ = 0 Then
                    If Not IsNull(rsItems!ʱ�䷽��) Then
                        If ExeTimeValid(rsItems!ʱ�䷽��, Val(.TextMatrix(lngCurRow, COL_Ƶ�ʴ���)), _
                            Val(.TextMatrix(lngCurRow, COL_Ƶ�ʼ��)), .TextMatrix(lngCurRow, COL_�����λ)) Then
                            .TextMatrix(lngCurRow, COL_ִ��ʱ��) = rsItems!ʱ�䷽��
                        End If
                    End If
                    If .TextMatrix(lngCurRow, COL_ִ��ʱ��) = "" Then
                        .TextMatrix(lngCurRow, COL_ִ��ʱ��) = Getȱʡʱ��(int���÷�Χ, .TextMatrix(lngCurRow, COL_Ƶ��))
                    End If
                End If
                If bln�ɼ����� Then
                    .TextMatrix(lngCurRow, COL_�÷�) = rsItems!����
                ElseIf bln��ҩ;�� Or bln��ҩ�÷� Then
                    '��ҩ����ҩ�䷽���÷�,ִ��ʱ��
                    If bln��ҩ�÷� Then
                        .TextMatrix(lngCurRow, COL_�÷�) = rsItems!����
                    End If
                    For j = lngCurRow - 1 To lngRow Step -1
                        If Val(.TextMatrix(j, COL_���ID)) = .RowData(lngCurRow) Then
                            If bln��ҩ;�� Then .TextMatrix(j, COL_�÷�) = rsItems!����
                            .TextMatrix(j, COL_ִ��ʱ��) = .TextMatrix(lngCurRow, COL_ִ��ʱ��)
                        Else
                            Exit For
                        End If
                    Next
                End If
            End If
            
            '����ҽ���Ϳ�������
            .TextMatrix(lngCurRow, COL_����ҽ��) = UserInfo.����
            .TextMatrix(lngCurRow, COL_��������ID) = Get��������ID(UserInfo.ID, mlng���˿���id, 1)
                                
            'ִ������
            If InStr(",5,6,7,", rsItems!���) > 0 Then
                If Nvl(rsItems!ִ������, 0) = 5 Then
                    .TextMatrix(lngCurRow, COL_ִ������) = 5
                Else
                    .TextMatrix(lngCurRow, COL_ִ������) = 4
                End If
            ElseIf bln��ҩ;�� Or bln��ҩ�巨 Or bln��ҩ�÷� Or bln�ɼ����� Then
                .TextMatrix(lngCurRow, COL_ִ������) = Nvl(rsItems!ִ������, 0)
            Else
                .TextMatrix(lngCurRow, COL_ִ������) = Nvl(rsItems!ִ������, 0)
            End If
            
            'ִ�п���ID:Ϊ0-����,5-Ժ��ִ��ʱȡ��Ϊ0
            If rsItems!��� = "Z" And InStr(",1,2,", Nvl(rsItems!��������, 0)) > 0 Then
                If Nvl(rsItems!ִ�п���ID, 0) <> 0 Then
                    .TextMatrix(lngCurRow, COL_ִ�п���ID) = Nvl(rsItems!ִ�п���ID, 0)
                Else
                    '���ۻ�סԺҽ��ȡ�ٴ�����(����ִ������)
                    If Nvl(rsItems!��������, 0) = 1 Then
                        '����:���������סԺ�ٴ�����
                        Call Get�ٴ�����(3, , lngTmp, , Not gbln�������Ҷ���)
                    ElseIf Nvl(rsItems!��������, 0) = 2 Then
                        'סԺ:����סԺ�ٴ�����
                        Call Get�ٴ�����(2, , lngTmp, , Not gbln�������Ҷ���)
                    End If
                    .TextMatrix(lngCurRow, COL_ִ�п���ID) = lngTmp
                End If
            ElseIf InStr(",0,5,", Val(.TextMatrix(lngCurRow, COL_ִ������))) = 0 Then
                If Nvl(rsItems!ִ�п���ID, 0) <> 0 Then
                    If InStr(",5,6,7,", rsItems!���) > 0 Then
                        strҩ��IDs = Get����ҩ��IDs(rsItems!���, rsItems!������ĿID, Val(.TextMatrix(lngCurRow, COL_�շ�ϸĿID)), mlng���˿���id, 1)
                        If InStr("," & strҩ��IDs & ",", "," & rsItems!ִ�п���ID & ",") > 0 Then
                            .TextMatrix(lngCurRow, COL_ִ�п���ID) = Nvl(rsItems!ִ�п���ID, 0)
                        End If
                    ElseIf Val(.TextMatrix(lngCurRow, COL_ִ������)) = 4 Then
                        '4-ָ������ʱ��ȡ,�����Ĺ̶�����
                        .TextMatrix(lngCurRow, COL_ִ�п���ID) = Nvl(rsItems!ִ�п���ID, 0)
                    End If
                End If
                If Val(.TextMatrix(lngCurRow, COL_ִ�п���ID)) = 0 Then
                    'ҩƷ�������������ͬ
                    If rsItems!��� = "5" Then
                        If lng��ҩ��ID = 0 Then
                            lng��ҩ��ID = Get����ִ�п���ID(mlng����ID, 0, rsItems!���, rsItems!������ĿID, Val(.TextMatrix(lngCurRow, COL_�շ�ϸĿID)), 4, mlng���˿���id, 0, 1, 1, True)
                        End If
                        .TextMatrix(lngCurRow, COL_ִ�п���ID) = lng��ҩ��ID
                    ElseIf rsItems!��� = "6" Then
                        If lng��ҩ��ID = 0 Then
                            lng��ҩ��ID = Get����ִ�п���ID(mlng����ID, 0, rsItems!���, rsItems!������ĿID, Val(.TextMatrix(lngCurRow, COL_�շ�ϸĿID)), 4, mlng���˿���id, 0, 1, 1, True)
                        End If
                        .TextMatrix(lngCurRow, COL_ִ�п���ID) = lng��ҩ��ID
                    ElseIf rsItems!��� = "7" Then
                        If lng��ҩ��ID = 0 Then
                            lng��ҩ��ID = Get����ִ�п���ID(mlng����ID, 0, rsItems!���, rsItems!������ĿID, Val(.TextMatrix(lngCurRow, COL_�շ�ϸĿID)), 4, mlng���˿���id, 0, 1, 1, True)
                        End If
                        .TextMatrix(lngCurRow, COL_ִ�п���ID) = lng��ҩ��ID
                    Else
                        '֮ǰ����������ID
                        .TextMatrix(lngCurRow, COL_ִ�п���ID) = Get����ִ�п���ID(mlng����ID, 0, rsItems!���, rsItems!������ĿID, 0, _
                            Val(.TextMatrix(lngCurRow, COL_ִ������)), mlng���˿���id, Val(.TextMatrix(lngCurRow, COL_��������ID)), 1, 1)
                    End If
                End If
            End If
            
            'ҽ������
            .TextMatrix(lngCurRow, COL_ҽ������) = Nvl(rsItems!ҽ������)
            
            '����ʱ��
            .TextMatrix(lngCurRow, COL_����ʱ��) = Format(vCurDate, "MM-dd HH:mm")
            .Cell(flexcpData, lngCurRow, COL_����ʱ��) = Format(vCurDate, "yyyy-MM-dd HH:mm")
            
            '������־
            .TextMatrix(lngCurRow, COL_��־) = chk����.Value '�����ڽ�����ͳһ����Ϊ����
            blnFirst = True
            If InStr(",5,6,", .TextMatrix(lngCurRow, COL_���)) > 0 Then
                If Val(.TextMatrix(lngCurRow, COL_���ID)) = Val(.TextMatrix(lngCurRow - 1, COL_���ID)) Then
                    blnFirst = False
                End If
            End If
            If blnFirst Then
                If Val(.TextMatrix(lngCurRow, COL_��־)) = 1 Then
                    Set .Cell(flexcpPicture, lngCurRow, COL_F��־) = imgFlag.ListImages("����").Picture
                    .Cell(flexcpPictureAlignment, lngCurRow, COL_F��־) = 4
                End If
            End If
            
            '��ȡҩƷ���
            If InStr(",5,6,7,", .TextMatrix(lngCurRow, COL_���)) > 0 Then
                If Val(.TextMatrix(lngCurRow, COL_�շ�ϸĿID)) <> 0 And Val(.TextMatrix(lngCurRow, COL_ִ�п���ID)) <> 0 Then
                    .TextMatrix(lngCurRow, COL_���) = GetStock(Val(.TextMatrix(lngCurRow, COL_�շ�ϸĿID)), Val(.TextMatrix(lngCurRow, COL_ִ�п���ID)))
                End If
            End If
            
            '----------------------
            '���龫ҩƷ��ʶ:��ҩ�䷽�����ζ��ҩ������
            If InStr(",5,6,", .TextMatrix(lngCurRow, COL_���)) > 0 And .TextMatrix(lngCurRow, COL_�������) <> "" Then
                If InStr(",����ҩ,����ҩ,����ҩ,", .TextMatrix(lngCurRow, COL_�������)) > 0 Then
                    .Cell(flexcpFontBold, lngCurRow, COL_ҽ������) = True
                End If
            End If
            
            '����һЩ������
            If (InStr(",F,G,D,7,E,C,", rsItems!���) > 0 And Not IsNull(rsItems!������)) Or bln��ҩ;�� Then
                .RowHidden(lngCurRow) = True
            End If
            
            'ҽ������
            If Not .RowHidden(lngCurRow) Then
                If InStr(",F,D,", rsItems!���) > 0 And IsNull(rsItems!������) Then
                    .TextMatrix(lngCurRow, COL_ҽ������) = rsItems!���� '��ʱ
                Else
                    .TextMatrix(lngCurRow, COL_ҽ������) = AdviceTextMake(lngCurRow)
                End If
            Else
                .TextMatrix(lngCurRow, COL_ҽ������) = rsItems!����
            End If
            
            
            If lngPreRow = -1 And Not .RowHidden(lngCurRow) Then lngPreRow = lngCurRow
                        
            '----------------------
            intCount = intCount + 1
            rsItems.MoveNext
        Next
        
        '--------------------------------------------------
        '�������Ӵ���
        For i = lngRow To lngCurRow
            'ȡ����������ҽ������
            If InStr(",F,D,", .TextMatrix(i, COL_���)) > 0 And Val(.TextMatrix(i, COL_���ID)) = 0 Then
                .TextMatrix(i, COL_ҽ������) = AdviceTextMake(i)
            End If
            
            '�������Ƶ���
            If Not .RowHidden(i) And .TextMatrix(i, COL_����) = "" Then
                .TextMatrix(i, COL_����) = GetItemPrice(i)
            End If
        Next
        
        '������Ӱ���е����
        Call AdviceSetҽ�����(lngCurRow + 1, intCount)
        
        '������ʵ��ҽ��ID
        For i = lngRow To lngCurRow
            lng���ID = .RowData(i)
            .RowData(i) = zlDatabase.GetNextId("����ҽ����¼")
            For j = i - 1 To lngRow Step -1
                If Val(.TextMatrix(j, COL_���ID)) = lng���ID Then
                    .TextMatrix(j, COL_���ID) = .RowData(i)
                Else
                    Exit For
                End If
            Next
            For j = i + 1 To lngCurRow
                If Val(.TextMatrix(j, COL_���ID)) = lng���ID Then
                    .TextMatrix(j, COL_���ID) = .RowData(i)
                Else
                    Exit For
                End If
            Next
        Next
        
        '--------------------------------------------------
        If .RowHidden(lngRow) Then 'Ѱ�ҿɼ���(���䷽�ͼ���֮��)
            For i = lngRow + 1 To .Rows - 1
                If Not .RowHidden(i) And .RowData(i) <> 0 Then
                    lngRow = i: Exit For
                End If
            Next
        End If
        
        .Row = lngRow: .Col = COL_ҽ������
        Call .ShowCell(.Row, .Col)
        .Redraw = flexRDDirect
        mblnRowChange = True
    End With
    Screen.MousePointer = 0
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function AdviceSet��ҩ�䷽(lng������ĿID As Long, ByVal lngRow As Long, ByVal lng�÷�ID As Long, ByVal strExtData As String, Optional rsCurr As ADODB.Recordset) As Long
'���ܣ�(����)������ҩ�䷽��ȱʡҽ������
'������lng������ĿID=�������ҩ�䷽ID��ζ��ҩID
'      lngRow=��ǰ������
'      lng�÷�ID=ȱʡ��ҩ�÷�ID
'      strExtData=�����䷽���ζҩ���巨����
'      rsCurr=������޸����䷽���ݺ����,�����Ҫ���ֵ�һЩ��ǰֵ
'���أ���������ҩ�䷽�ĵ�ǰ��ʾ�к�
    Dim rsItems As New ADODB.Recordset '��ҩ��ϸ��Ϣ
    Dim rsUse As New ADODB.Recordset '��ҩ�÷���Ϣ
    Dim rs�巨 As New ADODB.Recordset '��ҩ�巨��Ŀ��Ϣ
    Dim rs�÷� As New ADODB.Recordset '��ҩ�÷���Ŀ��Ϣ
    Dim arr��ҩs As Variant, str��ҩIDs As String, lng���ID As Long
    Dim lngCopyRow As Long 'ȱʡ������
    Dim lngDrugRow As Long '���ȱʡ����������ҩ�䷽,��Ϊ���䷽�ĵ�һ����ҩ��
    Dim lngFirstRow As Long '��ǰ�䷽�ĵ�һ����ҩ��
    Dim strSQL As String, i As Long
    
    Dim strƵ�� As String, intƵ�ʴ��� As Integer, intƵ�ʼ�� As Integer, str�����λ As String
    Dim lng�巨ID As Long, int�Ƴ� As Integer
    Dim strҽ�� As String, lngҽ��ID As Long
        
    On Error GoTo errH
    
    'ȡ��һ����һ��Ч��,ĳЩ����ȱʡ�������ͬ
    lngDrugRow = -1
    lngCopyRow = GetPreRow(lngRow)
    If lngCopyRow = -1 Then lngCopyRow = GetNextRow(lngRow)
    If lngCopyRow <> -1 Then
        If RowIn�䷽��(lngCopyRow) Then
            '�����һ��Ч������ҩ�䷽��,��ȡ���ĵ�һ��ҩ��
            lngDrugRow = vsAdvice.FindRow(CStr(vsAdvice.RowData(lngCopyRow)), , COL_���ID)
        End If
    End If
    
    '��ȡ������ݿ���Ϣ
    '------------------
    arr��ҩs = Split(Split(strExtData, "|")(0), ";")
    For i = 0 To UBound(arr��ҩs)
        str��ҩIDs = str��ҩIDs & "," & CStr(Split(arr��ҩs(i), ",")(0))
    Next
    str��ҩIDs = Mid(str��ҩIDs, 2)
    lng�巨ID = Val(Split(strExtData, "|")(1))
    
    '�䷽�÷���Ϣ:ֱ�������䷽ʱ���п�����,���뵥ζ��ҩ��
    strSQL = "Select A.�÷�ID,A.Ƶ��,A.�Ƴ�,A.ҽ������" & _
        " From �����÷����� A,������ĿĿ¼ B" & _
        " Where A.�÷�ID=B.ID And B.������� IN(1,3)" & _
        " And Nvl(A.����,0)=0 And A.��ĿID=[1]"
    Set rsUse = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng������ĿID)
    If Not rsUse.EOF Then lng�÷�ID = rsUse!�÷�ID 'ȱʡ���õ���ҩ�䷽�÷�����
    
    '�䷽���ζ��ҩ��Ϣ:��ҩ�޹�����,��Ӧ�ĵĹ���¼һ������ֻ��һ��
    strSQL = "Select A.*,B.ҩƷID,B.����ϵ��,B.�����װ,B.���ﵥλ,B.�ɷ����,C.����ְ��" & _
        " From ������ĿĿ¼ A,ҩƷ��� B,ҩƷ���� C" & _
        " Where A.ID=B.ҩ��ID And A.ID=C.ҩ��ID And A.ID IN(" & str��ҩIDs & ")"
    zlDatabase.OpenRecordset rsItems, strSQL, Me.Caption 'In
    
    '�䷽�巨��Ŀ��Ϣ
    strSQL = "Select * From ������ĿĿ¼ Where ID=[1]"
    Set rs�巨 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng�巨ID)
    
    '�䷽�÷���Ŀ��Ϣ
    strSQL = "Select * From ������ĿĿ¼ Where ID=[1]"
    Set rs�÷� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng�÷�ID)
    
    '�����䷽���ζ��ҩ��:�����û�����˳��
    '--------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------
    mblnRowChange = False
    
    '��ҩ�÷���ҽ��ID,ID˳������Ų�һ��һ��
    If Not rsCurr Is Nothing Then
        '�޸����䷽�е�����,�÷��б��Ϊ�޸�,ҽ��ID����
        lng���ID = rsCurr!ҽ��ID
    Else
        '���������ҩ�䷽
        lng���ID = zlDatabase.GetNextId("����ҽ����¼")
    End If
    
    For i = 0 To UBound(arr��ҩs)
        rsItems.Filter = "ID=" & CStr(Split(arr��ҩs(i), ",")(0)) 'Ӧ�ÿ϶���
        
        vsAdvice.AddItem "", lngRow
        
        vsAdvice.RowHidden(lngRow) = True
        vsAdvice.RowData(lngRow) = zlDatabase.GetNextId("����ҽ����¼")
        vsAdvice.TextMatrix(lngRow, COL_���ID) = lng���ID '��Ӧ���������ҩ�÷���
        vsAdvice.TextMatrix(lngRow, COL_EDIT) = 1 '����
        vsAdvice.TextMatrix(lngRow, COL_Ӥ��) = cboӤ��.ListIndex
        vsAdvice.TextMatrix(lngRow, COL_״̬) = 1 '�¿�
        vsAdvice.TextMatrix(lngRow, COL_���) = GetCurRow���(lngRow)
        Call AdviceSetҽ�����(lngRow + 1, 1) '�������
        
        vsAdvice.TextMatrix(lngRow, COL_���) = rsItems!���
        vsAdvice.TextMatrix(lngRow, COL_ҽ������) = rsItems!����
        vsAdvice.TextMatrix(lngRow, COL_������ĿID) = rsItems!ID
        vsAdvice.TextMatrix(lngRow, COL_���㷽ʽ) = Nvl(rsItems!���㷽ʽ, 0)
        vsAdvice.TextMatrix(lngRow, COL_Ƶ������) = Nvl(rsItems!ִ��Ƶ��, 0)
        vsAdvice.TextMatrix(lngRow, COL_��������) = Nvl(rsItems!��������)
        
        vsAdvice.TextMatrix(lngRow, COL_����) = FormatEx(Val(Split(arr��ҩs(i), ",")(1)), 5) '��ζҩ�ĵ�������
        vsAdvice.TextMatrix(lngRow, COL_������λ) = Nvl(rsItems!���㵥λ)
        vsAdvice.TextMatrix(lngRow, COL_ҽ������) = CStr(Split(arr��ҩs(i), ",")(2)) '��ζҩ�Ľ�ע
        
        '�����Ϣ:��ҩ�����ڹ�����,һ����
        vsAdvice.TextMatrix(lngRow, COL_�շ�ϸĿID) = rsItems!ҩƷID
        vsAdvice.TextMatrix(lngRow, COL_����ϵ��) = rsItems!����ϵ��
        vsAdvice.TextMatrix(lngRow, COL_���ﵥλ) = rsItems!���ﵥλ
        vsAdvice.TextMatrix(lngRow, COL_�����װ) = rsItems!�����װ
        vsAdvice.TextMatrix(lngRow, COL_�ɷ����) = Nvl(rsItems!�ɷ����, 0) '����ҩʵ��������
        vsAdvice.TextMatrix(lngRow, COL_����ְ��) = Nvl(rsItems!����ְ��)
        
        '�Ƽ�����:���Զ���
        vsAdvice.TextMatrix(lngRow, COL_�Ƽ�����) = Nvl(rsItems!�Ƽ�����, 0)
        
        If lngFirstRow <> 0 Then
            '����һ�������õ������ҩ��ͬ
            vsAdvice.TextMatrix(lngRow, COL_ִ������) = vsAdvice.TextMatrix(lngFirstRow, COL_ִ������)
            vsAdvice.TextMatrix(lngRow, COL_ִ�п���ID) = vsAdvice.TextMatrix(lngFirstRow, COL_ִ�п���ID)
            vsAdvice.TextMatrix(lngRow, COL_Ƶ��) = vsAdvice.TextMatrix(lngFirstRow, COL_Ƶ��)
            vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʴ���) = vsAdvice.TextMatrix(lngFirstRow, COL_Ƶ�ʴ���)
            vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʼ��) = vsAdvice.TextMatrix(lngFirstRow, COL_Ƶ�ʼ��)
            vsAdvice.TextMatrix(lngRow, COL_�����λ) = vsAdvice.TextMatrix(lngFirstRow, COL_�����λ)
            vsAdvice.TextMatrix(lngRow, COL_����) = vsAdvice.TextMatrix(lngFirstRow, COL_����)
            vsAdvice.TextMatrix(lngRow, COL_ִ��ʱ��) = vsAdvice.TextMatrix(lngFirstRow, COL_ִ��ʱ��)
            
            vsAdvice.TextMatrix(lngRow, COL_��ʼʱ��) = vsAdvice.TextMatrix(lngFirstRow, COL_��ʼʱ��)
            vsAdvice.Cell(flexcpData, lngRow, COL_��ʼʱ��) = vsAdvice.Cell(flexcpData, lngFirstRow, COL_��ʼʱ��)
            
            vsAdvice.TextMatrix(lngRow, COL_����ҽ��) = vsAdvice.TextMatrix(lngFirstRow, COL_����ҽ��)
            vsAdvice.TextMatrix(lngRow, COL_��������ID) = vsAdvice.TextMatrix(lngFirstRow, COL_��������ID)
            
            vsAdvice.TextMatrix(lngRow, COL_����ʱ��) = vsAdvice.TextMatrix(lngFirstRow, COL_����ʱ��)
            vsAdvice.Cell(flexcpData, lngRow, COL_����ʱ��) = vsAdvice.Cell(flexcpData, lngFirstRow, COL_����ʱ��)
            
            vsAdvice.TextMatrix(lngRow, COL_��־) = vsAdvice.TextMatrix(lngFirstRow, COL_��־)
        ElseIf Not rsCurr Is Nothing Then
            '�޸����䷽���ݺ���������,�����뵱ǰ��ֵ
            
            'ִ������:�޸�ʱ���ݵ�ǰ�������þ���
            vsAdvice.TextMatrix(lngRow, COL_ִ������) = Decode(Nvl(rsCurr!ִ������), "�Ա�ҩ", 5, 4)
            'ִ�п���
            vsAdvice.TextMatrix(lngRow, COL_ִ�п���ID) = Nvl(rsCurr!ִ�п���ID)
            
            vsAdvice.TextMatrix(lngRow, COL_Ƶ��) = Nvl(rsCurr!Ƶ��)
            vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʴ���) = Nvl(rsCurr!Ƶ�ʴ���)
            vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʼ��) = Nvl(rsCurr!Ƶ�ʼ��)
            vsAdvice.TextMatrix(lngRow, COL_�����λ) = Nvl(rsCurr!�����λ)
            vsAdvice.TextMatrix(lngRow, COL_����) = Nvl(rsCurr!����)
            vsAdvice.TextMatrix(lngRow, COL_ִ��ʱ��) = Nvl(rsCurr!ִ��ʱ��)
            
            vsAdvice.TextMatrix(lngRow, COL_��ʼʱ��) = Format(Nvl(rsCurr!��ʼʱ��), "MM-dd HH:mm")
            vsAdvice.Cell(flexcpData, lngRow, COL_��ʼʱ��) = CStr(Nvl(rsCurr!��ʼʱ��))
            
            vsAdvice.TextMatrix(lngRow, COL_����ҽ��) = Nvl(rsCurr!����ҽ��)
            vsAdvice.TextMatrix(lngRow, COL_��������ID) = Nvl(rsCurr!��������ID)
            
            vsAdvice.TextMatrix(lngRow, COL_����ʱ��) = Format(Nvl(rsCurr!����ʱ��), "MM-dd HH:mm")
            vsAdvice.Cell(flexcpData, lngRow, COL_����ʱ��) = CStr(Nvl(rsCurr!����ʱ��))
            
            vsAdvice.TextMatrix(lngRow, COL_��־) = Nvl(rsCurr!��־)
        Else
            'ִ������:��ҩ�䷽�����ҩ��ͬ,ȱʡ=4-ָ������
            vsAdvice.TextMatrix(lngRow, COL_ִ������) = 4
        
            'ִ�п���
            If lngDrugRow <> -1 Then 'ȱʡ����һ�䷽����ͬ
                vsAdvice.TextMatrix(lngRow, COL_ִ�п���ID) = vsAdvice.TextMatrix(lngDrugRow, COL_ִ�п���ID)
            End If
            If Val(vsAdvice.TextMatrix(lngRow, COL_ִ�п���ID)) = 0 Then
                vsAdvice.TextMatrix(lngRow, COL_ִ�п���ID) = Get����ִ�п���ID(mlng����ID, 0, rsItems!���, rsItems!ID, rsItems!ҩƷID, Nvl(rsItems!ִ�п���, 0), mlng���˿���id, 0, 1, 1, True)
            End If
            
            'ִ��Ƶ��
            '�����÷��������õ�����
            If Not rsUse.EOF Then
                If Not IsNull(rsUse!Ƶ��) Then
                    Call GetƵ����Ϣ_����(rsUse!Ƶ��, strƵ��, intƵ�ʴ���, intƵ�ʼ��, str�����λ)
                    vsAdvice.TextMatrix(lngRow, COL_Ƶ��) = strƵ��
                    vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʴ���) = intƵ�ʴ���
                    vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʼ��) = intƵ�ʼ��
                    vsAdvice.TextMatrix(lngRow, COL_�����λ) = str�����λ
                End If
            End If
            '��ȱʡ����һ����ͬ
            If vsAdvice.TextMatrix(lngRow, COL_Ƶ��) = "" And lngDrugRow <> -1 Then
                If Val(vsAdvice.TextMatrix(lngDrugRow, COL_EDIT)) = 1 And vsAdvice.TextMatrix(lngDrugRow, COL_Ƶ��) <> "" Then
                    vsAdvice.TextMatrix(lngRow, COL_Ƶ��) = vsAdvice.TextMatrix(lngDrugRow, COL_Ƶ��)
                    vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʴ���) = vsAdvice.TextMatrix(lngDrugRow, COL_Ƶ�ʴ���)
                    vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʼ��) = vsAdvice.TextMatrix(lngDrugRow, COL_Ƶ�ʼ��)
                    vsAdvice.TextMatrix(lngRow, COL_�����λ) = vsAdvice.TextMatrix(lngDrugRow, COL_�����λ)
                End If
            End If
            '��ȡȱʡֵ
            If vsAdvice.TextMatrix(lngRow, COL_Ƶ��) = "" Then
                Call GetȱʡƵ��(2, strƵ��, intƵ�ʴ���, intƵ�ʼ��, str�����λ)
                vsAdvice.TextMatrix(lngRow, COL_Ƶ��) = strƵ��
                vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʴ���) = intƵ�ʴ���
                vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʼ��) = intƵ�ʼ��
                vsAdvice.TextMatrix(lngRow, COL_�����λ) = str�����λ
            End If
            
            '����(����)
            If vsAdvice.TextMatrix(lngRow, COL_Ƶ��) <> "" Then
                int�Ƴ� = 1
                If Not rsUse.EOF Then int�Ƴ� = Nvl(rsUse!�Ƴ�, 1)
                '�䷽����
                vsAdvice.TextMatrix(lngRow, COL_����) = CalcȱʡҩƷ����(1, int�Ƴ�, _
                        Val(vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʴ���)), _
                        Val(vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʼ��)), _
                        vsAdvice.TextMatrix(lngRow, COL_�����λ))
            End If
            
            'ִ��ʱ��
            If lngDrugRow <> -1 Then 'ȱʡ����һ����ͬ
                If vsAdvice.TextMatrix(lngRow, COL_Ƶ��) = vsAdvice.TextMatrix(lngDrugRow, COL_Ƶ��) Then
                    vsAdvice.TextMatrix(lngRow, COL_ִ��ʱ��) = vsAdvice.TextMatrix(lngDrugRow, COL_ִ��ʱ��)
                End If
            End If
            If vsAdvice.TextMatrix(lngRow, COL_ִ��ʱ��) = "" Then 'ȱʡʱ�䷽��
                vsAdvice.TextMatrix(lngRow, COL_ִ��ʱ��) = Getȱʡʱ��(2, vsAdvice.TextMatrix(lngRow, COL_Ƶ��), lng�÷�ID)
            End If
            
            '��ʼʱ��
            If IsDate(txt��ʼʱ��.Text) Then
                vsAdvice.TextMatrix(lngRow, COL_��ʼʱ��) = Format(txt��ʼʱ��.Text, "MM-dd HH:mm")
                vsAdvice.Cell(flexcpData, lngRow, COL_��ʼʱ��) = txt��ʼʱ��.Text
            End If
            
            '����ҽ��
            vsAdvice.TextMatrix(lngRow, COL_����ҽ��) = UserInfo.����
            vsAdvice.TextMatrix(lngRow, COL_��������ID) = Get��������ID(UserInfo.ID, mlng���˿���id, 1)
            
            vsAdvice.TextMatrix(lngRow, COL_����ʱ��) = Format(zlDatabase.Currentdate, "MM-dd HH:mm")
            vsAdvice.Cell(flexcpData, lngRow, COL_����ʱ��) = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
            vsAdvice.TextMatrix(lngRow, COL_��־) = chk����.Value
        End If
        
        '---------------------------------------
        If lngFirstRow = 0 Then lngFirstRow = lngRow '����ҩ�䷽�ĵ�һ�������ҩ��
        lngRow = lngRow + 1 '���ֵ�ǰ������λ��
    Next
    
    '������ҩ�䷽�巨��
    '--------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------
    vsAdvice.AddItem "", lngRow
    vsAdvice.RowHidden(lngRow) = True
    vsAdvice.RowData(lngRow) = zlDatabase.GetNextId("����ҽ����¼")
    vsAdvice.TextMatrix(lngRow, COL_���ID) = lng���ID
    vsAdvice.TextMatrix(lngRow, COL_EDIT) = 1 '����
    vsAdvice.TextMatrix(lngRow, COL_Ӥ��) = cboӤ��.ListIndex
    vsAdvice.TextMatrix(lngRow, COL_״̬) = 1 '�¿�
    vsAdvice.TextMatrix(lngRow, COL_���) = GetCurRow���(lngRow)
    Call AdviceSetҽ�����(lngRow + 1, 1) '�������
    vsAdvice.TextMatrix(lngRow, COL_���) = rs�巨!���
    vsAdvice.TextMatrix(lngRow, COL_������ĿID) = lng�巨ID
    vsAdvice.TextMatrix(lngRow, COL_���㷽ʽ) = Nvl(rs�巨!���㷽ʽ, 0)
    vsAdvice.TextMatrix(lngRow, COL_Ƶ������) = Nvl(rs�巨!ִ��Ƶ��, 0)
    vsAdvice.TextMatrix(lngRow, COL_��������) = Nvl(rs�巨!��������)
    
    '!��ҩ�巨��Ҳ�����ҩ�ĸ���
    vsAdvice.TextMatrix(lngRow, COL_����) = vsAdvice.TextMatrix(lngFirstRow, COL_����)
    
    vsAdvice.TextMatrix(lngRow, COL_ҽ������) = rs�巨!����
    
    vsAdvice.TextMatrix(lngRow, COL_��ʼʱ��) = vsAdvice.TextMatrix(lngFirstRow, COL_��ʼʱ��)
    vsAdvice.Cell(flexcpData, lngRow, COL_��ʼʱ��) = vsAdvice.Cell(flexcpData, lngFirstRow, COL_��ʼʱ��)
    
    vsAdvice.TextMatrix(lngRow, COL_Ƶ��) = vsAdvice.TextMatrix(lngFirstRow, COL_Ƶ��)
    vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʴ���) = vsAdvice.TextMatrix(lngFirstRow, COL_Ƶ�ʴ���)
    vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʼ��) = vsAdvice.TextMatrix(lngFirstRow, COL_Ƶ�ʼ��)
    vsAdvice.TextMatrix(lngRow, COL_�����λ) = vsAdvice.TextMatrix(lngFirstRow, COL_�����λ)
    vsAdvice.TextMatrix(lngRow, COL_ִ��ʱ��) = vsAdvice.TextMatrix(lngFirstRow, COL_ִ��ʱ��)
    
    'ִ������:ȱʡ������Ŀ����(������ΪԺ��ִ��),�޸�ʱ���ݵ�ǰ��������
    If rsCurr Is Nothing Then
        vsAdvice.TextMatrix(lngRow, COL_ִ������) = Nvl(rs�巨!ִ�п���, 0)
    Else
        vsAdvice.TextMatrix(lngRow, COL_ִ������) = Decode(Nvl(rsCurr!ִ������), "��Ժ��ҩ", 5, Nvl(rs�巨!ִ�п���, 0))
    End If
    
    '��ҩ�巨���δ����ִ�п���,��ȱʡΪ�������ڲ���(����Ҫ��Ϊ�������ڿ���!!)
    If InStr(",0,5,", Val(vsAdvice.TextMatrix(lngRow, COL_ִ������))) = 0 Then
        vsAdvice.TextMatrix(lngRow, COL_ִ�п���ID) = Get����ִ�п���ID(mlng����ID, 0, rs�巨!���, lng�巨ID, 0, _
            Nvl(rs�巨!ִ�п���, 0), mlng���˿���id, Val(vsAdvice.TextMatrix(lngFirstRow, COL_��������ID)), 1, 1)
    End If
    
    vsAdvice.TextMatrix(lngRow, COL_�Ƽ�����) = Nvl(rs�巨!�Ƽ�����, 0)
    vsAdvice.TextMatrix(lngRow, COL_��������ID) = vsAdvice.TextMatrix(lngFirstRow, COL_��������ID)
    vsAdvice.TextMatrix(lngRow, COL_����ҽ��) = vsAdvice.TextMatrix(lngFirstRow, COL_����ҽ��)
    
    vsAdvice.TextMatrix(lngRow, COL_����ʱ��) = vsAdvice.TextMatrix(lngFirstRow, COL_����ʱ��)
    vsAdvice.Cell(flexcpData, lngRow, COL_����ʱ��) = vsAdvice.Cell(flexcpData, lngFirstRow, COL_����ʱ��)
    
    vsAdvice.TextMatrix(lngRow, COL_��־) = vsAdvice.TextMatrix(lngFirstRow, COL_��־)
    
    '���ֵ�ǰ������λ��
    lngRow = lngRow + 1
    
    '������ҩ�䷽�÷���:��ҩ�䷽����ʾ��
    '--------------------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------
    vsAdvice.RowData(lngRow) = lng���ID
    
    If Not rsCurr Is Nothing Then
        '�޸����䷽����,���Ϊ�޸�
        If InStr(",0,3,", rsCurr!Edit) > 0 Then
            vsAdvice.TextMatrix(lngRow, COL_EDIT) = 2 '���Ϊ���޸�
        Else
            vsAdvice.TextMatrix(lngRow, COL_EDIT) = rsCurr!Edit '���������������޸�
        End If
    Else
        '���������ҩ�䷽,Ϊ����
        vsAdvice.TextMatrix(lngRow, COL_EDIT) = 1
    End If
    
    vsAdvice.TextMatrix(lngRow, COL_Ӥ��) = cboӤ��.ListIndex
    vsAdvice.TextMatrix(lngRow, COL_״̬) = 1 '�¿�
    vsAdvice.TextMatrix(lngRow, COL_���) = GetCurRow���(lngRow)
    Call AdviceSetҽ�����(lngRow + 1, 1) '�������
    vsAdvice.TextMatrix(lngRow, COL_���) = rs�÷�!���
    vsAdvice.TextMatrix(lngRow, COL_������ĿID) = lng�÷�ID
    vsAdvice.TextMatrix(lngRow, COL_���㷽ʽ) = Nvl(rs�÷�!���㷽ʽ, 0)
    vsAdvice.TextMatrix(lngRow, COL_Ƶ������) = Nvl(rs�÷�!ִ��Ƶ��, 0)
    vsAdvice.TextMatrix(lngRow, COL_��������) = Nvl(rs�÷�!��������)
    
    '!��ҩ�÷���Ҳ�����ҩ�ĸ���
    vsAdvice.TextMatrix(lngRow, COL_����) = vsAdvice.TextMatrix(lngFirstRow, COL_����)
    vsAdvice.TextMatrix(lngRow, COL_������λ) = "��"
    
    vsAdvice.TextMatrix(lngRow, COL_��ʼʱ��) = vsAdvice.TextMatrix(lngFirstRow, COL_��ʼʱ��)
    vsAdvice.Cell(flexcpData, lngRow, COL_��ʼʱ��) = vsAdvice.Cell(flexcpData, lngFirstRow, COL_��ʼʱ��)
    
    vsAdvice.TextMatrix(lngRow, COL_����) = rs�÷�!����
    vsAdvice.TextMatrix(lngRow, COL_�÷�) = rs�÷�!����
    vsAdvice.TextMatrix(lngRow, COL_Ƶ��) = vsAdvice.TextMatrix(lngFirstRow, COL_Ƶ��)
    vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʴ���) = vsAdvice.TextMatrix(lngFirstRow, COL_Ƶ�ʴ���)
    vsAdvice.TextMatrix(lngRow, COL_Ƶ�ʼ��) = vsAdvice.TextMatrix(lngFirstRow, COL_Ƶ�ʼ��)
    vsAdvice.TextMatrix(lngRow, COL_�����λ) = vsAdvice.TextMatrix(lngFirstRow, COL_�����λ)
    vsAdvice.TextMatrix(lngRow, COL_ִ��ʱ��) = vsAdvice.TextMatrix(lngFirstRow, COL_ִ��ʱ��)
    
    'ִ������:ȱʡ������Ŀ����(������ΪԺ��ִ��),�޸�ʱ���ݵ�ǰ��������
    If rsCurr Is Nothing Then
        vsAdvice.TextMatrix(lngRow, COL_ִ������) = Nvl(rs�÷�!ִ�п���, 0)
    Else
        vsAdvice.TextMatrix(lngRow, COL_ִ������) = Decode(Nvl(rsCurr!ִ������), "��Ժ��ҩ", 5, Nvl(rs�÷�!ִ�п���, 0))
    End If
    
    '��ҩ�÷����δ����ִ�п���,��ȱʡΪ�������ڲ���(����Ҫ��Ϊ�������ڿ���!!)
    If InStr(",0,5,", Val(vsAdvice.TextMatrix(lngRow, COL_ִ������))) = 0 Then
        vsAdvice.TextMatrix(lngRow, COL_ִ�п���ID) = Get����ִ�п���ID(mlng����ID, 0, rs�÷�!���, lng�÷�ID, 0, _
            Nvl(rs�÷�!ִ�п���, 0), mlng���˿���id, Val(vsAdvice.TextMatrix(lngFirstRow, COL_��������ID)), 1, 1)
    End If
    
    vsAdvice.TextMatrix(lngRow, COL_�Ƽ�����) = Nvl(rs�÷�!�Ƽ�����, 0)
    vsAdvice.TextMatrix(lngRow, COL_��������ID) = vsAdvice.TextMatrix(lngFirstRow, COL_��������ID)
    vsAdvice.TextMatrix(lngRow, COL_����ҽ��) = vsAdvice.TextMatrix(lngFirstRow, COL_����ҽ��)
    
    vsAdvice.TextMatrix(lngRow, COL_����ʱ��) = vsAdvice.TextMatrix(lngFirstRow, COL_����ʱ��)
    vsAdvice.Cell(flexcpData, lngRow, COL_����ʱ��) = vsAdvice.Cell(flexcpData, lngFirstRow, COL_����ʱ��)
    
    vsAdvice.TextMatrix(lngRow, COL_��־) = vsAdvice.TextMatrix(lngFirstRow, COL_��־)
    If Val(vsAdvice.TextMatrix(lngRow, COL_��־)) = 1 Then
        Set vsAdvice.Cell(flexcpPicture, lngRow, COL_F��־) = imgFlag.ListImages("����").Picture
        vsAdvice.Cell(flexcpPictureAlignment, lngRow, COL_F��־) = 4
    End If
    
    If Not rsCurr Is Nothing Then
        vsAdvice.TextMatrix(lngRow, COL_ҽ������) = Nvl(rsCurr!ҽ������)
    ElseIf Not rsUse.EOF Then
        vsAdvice.TextMatrix(lngRow, COL_ҽ������) = Nvl(rsUse!ҽ������)
    End If
    
    '��ҩ�䷽����ҩ���
    Call GetDrugStock(lngRow)
    
    '��ҩ�䷽ҽ������
    vsAdvice.TextMatrix(lngRow, COL_ҽ������) = AdviceTextMake(lngRow)
    
    '-------------------
    vsAdvice.Row = lngRow
    mblnRowChange = True
        
    AdviceSet��ҩ�䷽ = lngRow
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function AdviceSet�������(ByVal lngRow As Long, ByVal lng�ɼ�����ID As Long, ByVal strExtData As String, Optional rsCurr As ADODB.Recordset) As Long
'���ܣ����������ļ���(���)
'������rsItems=�����ѡ�񷵻صļ�¼��
'      lngRow=��ǰ������
'      lng�ɼ�����ID=ȱʡ�Ĳɼ�����
'      strExtData=���:"��ĿID1,��ĿID2,...;����걾"
'      rsCurr=�޸ļ�����Ŀʱ��
'���أ�����֮��ĵ�ǰ��ʾ�к�
    Dim rsMore As New ADODB.Recordset '�ɼ�������Ϣ
    Dim rsItems As New ADODB.Recordset '������Ŀ��Ϣ
    Dim arrItems As Variant, strItems As String
    Dim strSQL As String, curDate As Date
    Dim strҽ�� As String, lngҽ��ID As Long
    Dim strƵ�� As String, intƵ�ʴ��� As Integer
    Dim intƵ�ʼ�� As Integer, str�����λ As String
    Dim lng���ID As Long, strҽ������ As String
    Dim lngCopyRow As Long, lngFirstRow As Long, i As Long
    
    On Error GoTo errH
    
    'ȡ��һ����һ��Ч��,ĳЩ����ȱʡ�������ͬ
    lngCopyRow = GetPreRow(lngRow)
    If lngCopyRow = -1 Then lngCopyRow = GetNextRow(lngRow)
    '��ǰʱ��
    curDate = zlDatabase.Currentdate
    
    '������Ŀ��Ϣ
    '----------------------------------------------------------------------------
    '����������Ŀ��Ϣ:������˳��
    arrItems = Split(Split(strExtData, ";")(0), ",")
    For i = UBound(arrItems) To 0 Step -1
        strItems = strItems & "," & Val(arrItems(i))
    Next
    strSQL = "Select * From ������ĿĿ¼ Where ID IN(" & Mid(strItems, 2) & ")"
    Call zlDatabase.OpenRecordset(rsItems, strSQL, Me.Caption) 'In
    
    'ȡĳ��������Ŀ�Ĳɼ�����
    strSQL = "Select A.��ĿID,Nvl(A.����,0) as ���,A.�÷�ID" & _
        " From �����÷����� A,������ĿĿ¼ B" & _
        " Where A.�÷�ID=B.ID And B.������� IN(1,3)" & _
        " And A.��ĿID IN(" & Mid(strItems, 2) & ")" & _
        " Order by A.��ĿID,Nvl(A.����,0)"
    Call zlDatabase.OpenRecordset(rsMore, strSQL, Me.Caption) 'In
    If Not rsMore.EOF Then
        If rsCurr Is Nothing Or lng�ɼ�����ID = 0 Then
            lng�ɼ�����ID = rsMore!�÷�ID '�޸�ʱ����
        End If
    End If
    
    strSQL = "Select * From ������ĿĿ¼ Where ID=[1]"
    Set rsMore = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng�ɼ�����ID)
    
    mblnRowChange = False
    
    '���ø��м�����Ŀ
    '----------------------------------------------------------------------------
    '�ɼ�����ҽ��ID,ID˳������Ų�һ��һ��
    If Not rsCurr Is Nothing Then
        '�޸��˼�������е�����,�ɼ������б��Ϊ�޸�,ҽ��ID����
        lng���ID = rsCurr!ҽ��ID
    Else
        '���������ҩ�䷽
        lng���ID = zlDatabase.GetNextId("����ҽ����¼")
    End If
    With vsAdvice
        For i = 1 To rsItems.RecordCount
            .AddItem "", lngRow
            
            .RowHidden(lngRow) = True
            .RowData(lngRow) = zlDatabase.GetNextId("����ҽ����¼")
            .TextMatrix(lngRow, COL_���ID) = lng���ID '��Ӧ���ɼ�������
            .TextMatrix(lngRow, COL_EDIT) = 1 '����
            .TextMatrix(lngRow, COL_Ӥ��) = cboӤ��.ListIndex
            .TextMatrix(lngRow, COL_״̬) = 1 '�¿�
            
            .TextMatrix(lngRow, COL_���) = GetCurRow���(lngRow)
            Call AdviceSetҽ�����(lngRow + 1, 1) '�������
            
            .TextMatrix(lngRow, COL_���) = rsItems!���
            .TextMatrix(lngRow, COL_ҽ������) = rsItems!����
            .TextMatrix(lngRow, COL_������ĿID) = rsItems!ID
            .TextMatrix(lngRow, COL_���㷽ʽ) = Nvl(rsItems!���㷽ʽ, 0)
            .TextMatrix(lngRow, COL_Ƶ������) = Nvl(rsItems!ִ��Ƶ��, 0)
            .TextMatrix(lngRow, COL_��������) = Nvl(rsItems!��������)
            .TextMatrix(lngRow, COL_��������) = Nvl(rsItems!¼������)
            .TextMatrix(lngRow, COL_�Ƽ�����) = Nvl(rsItems!�Ƽ�����, 0)
            .TextMatrix(lngRow, COL_ִ������) = Nvl(rsItems!ִ�п���, 0)
            '����걾
            .TextMatrix(lngRow, COL_�걾��λ) = Split(strExtData, ";")(1)
            
            '��������һ���ɼ��ļ�����Ŀ��ͬ
            If lngFirstRow <> 0 Then
                .TextMatrix(lngRow, COL_����) = .TextMatrix(lngFirstRow, COL_����)
                
                'һ���ɼ��ļ�����ĿӦ����ͬ
                If InStr(",0,5,", Val(.TextMatrix(lngRow, COL_ִ������))) = 0 Then
                    .TextMatrix(lngRow, COL_ִ�п���ID) = .TextMatrix(lngFirstRow, COL_ִ�п���ID)
                End If
                .TextMatrix(lngRow, COL_Ƶ��) = .TextMatrix(lngFirstRow, COL_Ƶ��)
                .TextMatrix(lngRow, COL_Ƶ�ʴ���) = .TextMatrix(lngFirstRow, COL_Ƶ�ʴ���)
                .TextMatrix(lngRow, COL_Ƶ�ʼ��) = .TextMatrix(lngFirstRow, COL_Ƶ�ʼ��)
                .TextMatrix(lngRow, COL_�����λ) = .TextMatrix(lngFirstRow, COL_�����λ)
                .TextMatrix(lngRow, COL_ִ��ʱ��) = .TextMatrix(lngFirstRow, COL_ִ��ʱ��)
                
                .TextMatrix(lngRow, COL_��ʼʱ��) = .TextMatrix(lngFirstRow, COL_��ʼʱ��)
                .Cell(flexcpData, lngRow, COL_��ʼʱ��) = .Cell(flexcpData, lngFirstRow, COL_��ʼʱ��)
                
                .TextMatrix(lngRow, COL_����ҽ��) = .TextMatrix(lngFirstRow, COL_����ҽ��)
                .TextMatrix(lngRow, COL_��������ID) = .TextMatrix(lngFirstRow, COL_��������ID)
                
                .TextMatrix(lngRow, COL_����ʱ��) = .TextMatrix(lngFirstRow, COL_����ʱ��)
                .Cell(flexcpData, lngRow, COL_����ʱ��) = .Cell(flexcpData, lngFirstRow, COL_����ʱ��)
                
                .TextMatrix(lngRow, COL_��־) = .TextMatrix(lngFirstRow, COL_��־)
            ElseIf Not rsCurr Is Nothing Then
                .TextMatrix(lngRow, COL_����) = Nvl(rsCurr!����, 1)
                
                'ִ�п���:ִ������Ϊ(0-����,5-Ժ��ִ��)��ִ�п���
                If InStr(",0,5,", Val(.TextMatrix(lngRow, COL_ִ������))) = 0 Then
                    If Nvl(rsCurr!ִ�п���ID, 0) <> 0 Then
                        .TextMatrix(lngRow, COL_ִ�п���ID) = rsCurr!ִ�п���ID
                    Else
                        .TextMatrix(lngRow, COL_ִ�п���ID) = Get����ִ�п���ID(mlng����ID, 0, rsItems!���, rsItems!ID, 0, _
                            Nvl(rsItems!ִ�п���, 0), mlng���˿���id, Nvl(rsCurr!��������ID, 0), 1, 1)
                    End If
                End If
                
                'ִ��Ƶ��
                .TextMatrix(lngRow, COL_Ƶ��) = Nvl(rsCurr!Ƶ��)
                .TextMatrix(lngRow, COL_Ƶ�ʴ���) = Nvl(rsCurr!Ƶ�ʴ���)
                .TextMatrix(lngRow, COL_Ƶ�ʼ��) = Nvl(rsCurr!Ƶ�ʼ��)
                .TextMatrix(lngRow, COL_�����λ) = Nvl(rsCurr!�����λ)
                .TextMatrix(lngRow, COL_ִ��ʱ��) = Nvl(rsCurr!ִ��ʱ��)
                
                'ʱ��/����/ҽ��
                .TextMatrix(lngRow, COL_��ʼʱ��) = Format(Nvl(rsCurr!��ʼʱ��), "MM-dd HH:mm")
                .Cell(flexcpData, lngRow, COL_��ʼʱ��) = CStr(Nvl(rsCurr!��ʼʱ��))
                
                .TextMatrix(lngRow, COL_����ʱ��) = Format(Nvl(rsCurr!����ʱ��), "MM-dd HH:mm")
                .Cell(flexcpData, lngRow, COL_����ʱ��) = CStr(Nvl(rsCurr!����ʱ��))
                
                .TextMatrix(lngRow, COL_����ҽ��) = Nvl(rsCurr!����ҽ��)
                .TextMatrix(lngRow, COL_��������ID) = Nvl(rsCurr!��������ID)
                
                .TextMatrix(lngRow, COL_��־) = Nvl(rsCurr!��־)
            Else
                .TextMatrix(lngRow, COL_����) = 1 'ȱʡΪ1(��)
                
                '����ҽ��
                .TextMatrix(lngRow, COL_����ҽ��) = UserInfo.����
                .TextMatrix(lngRow, COL_��������ID) = Get��������ID(UserInfo.ID, mlng���˿���id, 1)
                
                'ִ�п���:ִ������Ϊ(0-����,5-Ժ��ִ��)��ִ�п���
                If InStr(",0,5,", Val(.TextMatrix(lngRow, COL_ִ������))) = 0 Then
                    '֮ǰҪ�����������ID
                    .TextMatrix(lngRow, COL_ִ�п���ID) = Get����ִ�п���ID(mlng����ID, 0, rsItems!���, rsItems!ID, 0, _
                        Nvl(rsItems!ִ�п���, 0), mlng���˿���id, Val(.TextMatrix(lngRow, COL_��������ID)), 1, 1)
                End If
                
                'ִ��Ƶ��
                Call GetȱʡƵ��(GetƵ�ʷ�Χ(lngRow), strƵ��, intƵ�ʴ���, intƵ�ʼ��, str�����λ)
                .TextMatrix(lngRow, COL_Ƶ��) = strƵ��
                .TextMatrix(lngRow, COL_Ƶ�ʴ���) = intƵ�ʴ���
                .TextMatrix(lngRow, COL_Ƶ�ʼ��) = intƵ�ʼ��
                .TextMatrix(lngRow, COL_�����λ) = str�����λ
            
                'ִ��ʱ��:"��ѡƵ��"(ҩƷ�ǿ�ѡƵ��,����������Ϊһ����)
                If Val(.TextMatrix(lngRow, COL_Ƶ������)) = 0 Then
                    If lngCopyRow <> -1 Then '����һ����ͬ
                        If .TextMatrix(lngRow, COL_Ƶ��) = .TextMatrix(lngCopyRow, COL_Ƶ��) Then
                            .TextMatrix(lngRow, COL_ִ��ʱ��) = .TextMatrix(lngCopyRow, COL_ִ��ʱ��)
                        End If
                    End If
                    If .TextMatrix(lngRow, COL_ִ��ʱ��) = "" Then  'ȱʡʱ�䷽��
                        .TextMatrix(lngRow, COL_ִ��ʱ��) = Getȱʡʱ��(1, .TextMatrix(lngRow, COL_Ƶ��))
                    End If
                End If
            
                '��ʼʱ��
                If IsDate(txt��ʼʱ��.Text) Then
                    .TextMatrix(lngRow, COL_��ʼʱ��) = Format(txt��ʼʱ��.Text, "MM-dd HH:mm")
                    .Cell(flexcpData, lngRow, COL_��ʼʱ��) = txt��ʼʱ��.Text
                End If
                
                '����ʱ��
                .TextMatrix(lngRow, COL_����ʱ��) = Format(curDate, "MM-dd HH:mm")
                .Cell(flexcpData, lngRow, COL_����ʱ��) = Format(curDate, "yyyy-MM-dd HH:mm")
                
                '������־
                .TextMatrix(lngRow, COL_��־) = chk����.Value
            End If
            
            strҽ������ = strҽ������ & "," & rsItems!���� 'ҽ������
            If lngFirstRow = 0 Then lngFirstRow = lngRow '��һ��Ŀ��
            lngRow = lngRow + 1 '���ֵ�ǰ������λ��
            
            rsItems.MoveNext
        Next
        
        '���ñ걾�Ĳɼ�����
        '----------------------------------------------------------------------------
        rsItems.MoveFirst
        .RowData(lngRow) = lng���ID
        
        If Not rsCurr Is Nothing Then
            '�޸��˼����������,���Ϊ�޸�
            If InStr(",0,3,", rsCurr!Edit) > 0 Then
                vsAdvice.TextMatrix(lngRow, COL_EDIT) = 2 '���Ϊ���޸�
            Else
                vsAdvice.TextMatrix(lngRow, COL_EDIT) = rsCurr!Edit '���������������޸�
            End If
        Else
            '������ļ������,Ϊ����
            vsAdvice.TextMatrix(lngRow, COL_EDIT) = 1
        End If
        
        .TextMatrix(lngRow, COL_Ӥ��) = cboӤ��.ListIndex
        .TextMatrix(lngRow, COL_״̬) = 1 '�¿�
        
        .TextMatrix(lngRow, COL_���) = GetCurRow���(lngRow)
        Call AdviceSetҽ�����(lngRow + 1, 1) '�������
        
        .TextMatrix(lngRow, COL_���) = rsMore!���
        .TextMatrix(lngRow, COL_����) = rsMore!����
        .TextMatrix(lngRow, COL_�÷�) = rsMore!����
        .TextMatrix(lngRow, COL_������ĿID) = rsMore!ID
        .TextMatrix(lngRow, COL_���㷽ʽ) = Nvl(rsMore!���㷽ʽ, 0)
        .TextMatrix(lngRow, COL_Ƶ������) = Nvl(rsMore!ִ��Ƶ��, 0)
        .TextMatrix(lngRow, COL_��������) = Nvl(rsMore!��������)
        .TextMatrix(lngRow, COL_�Ƽ�����) = Nvl(rsMore!�Ƽ�����, 0)
        .TextMatrix(lngRow, COL_�걾��λ) = .TextMatrix(lngFirstRow, COL_�걾��λ)
        
        '����Ϊ������Ŀ��,�������Ŀ��ͬ
        .TextMatrix(lngRow, COL_����) = .TextMatrix(lngFirstRow, COL_����)
        .TextMatrix(lngRow, COL_������λ) = Nvl(rsMore!���㵥λ)
        
        'ִ��Ƶ��
        .TextMatrix(lngRow, COL_Ƶ��) = .TextMatrix(lngFirstRow, COL_Ƶ��)
        .TextMatrix(lngRow, COL_Ƶ�ʴ���) = .TextMatrix(lngFirstRow, COL_Ƶ�ʴ���)
        .TextMatrix(lngRow, COL_Ƶ�ʼ��) = .TextMatrix(lngFirstRow, COL_Ƶ�ʼ��)
        .TextMatrix(lngRow, COL_�����λ) = .TextMatrix(lngFirstRow, COL_�����λ)
        .TextMatrix(lngRow, COL_ִ��ʱ��) = .TextMatrix(lngFirstRow, COL_ִ��ʱ��)
        .TextMatrix(lngRow, COL_ִ������) = Nvl(rsMore!ִ�п���, 0)
        
        'ִ�п���:ִ������Ϊ(0-����,5-Ժ��ִ��)��ִ�п���
        If InStr(",0,5,", Val(.TextMatrix(lngRow, COL_ִ������))) = 0 Then
            .TextMatrix(lngRow, COL_ִ�п���ID) = Get����ִ�п���ID(mlng����ID, 0, rsMore!���, rsMore!ID, 0, _
                Nvl(rsMore!ִ�п���, 0), mlng���˿���id, Val(.TextMatrix(lngFirstRow, COL_��������ID)), 1, 1)
        End If
        
        'ʱ��/����/ҽ��
        .TextMatrix(lngRow, COL_��ʼʱ��) = .TextMatrix(lngFirstRow, COL_��ʼʱ��)
        .Cell(flexcpData, lngRow, COL_��ʼʱ��) = .Cell(flexcpData, lngFirstRow, COL_��ʼʱ��)
        .TextMatrix(lngRow, COL_����ʱ��) = .TextMatrix(lngFirstRow, COL_����ʱ��)
        .Cell(flexcpData, lngRow, COL_����ʱ��) = .Cell(flexcpData, lngFirstRow, COL_����ʱ��)
        .TextMatrix(lngRow, COL_��������ID) = .TextMatrix(lngFirstRow, COL_��������ID)
        .TextMatrix(lngRow, COL_����ҽ��) = .TextMatrix(lngFirstRow, COL_����ҽ��)
        
        '��ʾ������־
        .TextMatrix(lngRow, COL_��־) = .TextMatrix(lngFirstRow, COL_��־)
        If Val(.TextMatrix(lngRow, COL_��־)) = 2 Then
            Set .Cell(flexcpPicture, lngRow, COL_F��־) = imgFlag.ListImages("��¼").Picture
            .Cell(flexcpPictureAlignment, lngRow, COL_F��־) = 4
        ElseIf Val(.TextMatrix(lngRow, COL_��־)) = 1 Then
            Set .Cell(flexcpPicture, lngRow, COL_F��־) = imgFlag.ListImages("����").Picture
            .Cell(flexcpPictureAlignment, lngRow, COL_F��־) = 4
        End If
                
        If Not rsCurr Is Nothing Then
            .TextMatrix(lngRow, COL_ҽ������) = Nvl(rsCurr!ҽ������)
        End If
        
        'ҽ������:����1,����2(�걾 �ɼ�����)
        .TextMatrix(lngRow, COL_ҽ������) = AdviceTextMake(lngRow)
        
        .Row = lngRow
    End With
    mblnRowChange = True
    AdviceSet������� = lngRow
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub AdviceSet������Ŀ(rsInput As ADODB.Recordset, ByVal lngRow As Long, ByVal lng��ҩ;��ID As Long, ByVal lngGroupRow As Long, ByVal strExtData As String)
'���ܣ���������(����)���С�����ҩ�����(���)������(���)��������������Ŀ��ȱʡҽ������
'������rsInput=�����ѡ�񷵻صļ�¼��
'      lngRow=��ǰ������
'      lng��ҩ;��ID=ȱʡ��ҩ;��ID,��һ����ҩʱ�ĸ�ҩ;��ID
'      lngGroupRow=��һ����ҩ��һ���ҩ�в����µĳ�ҩ��ʱ,��Ӧһ����ҩ��һ���к�
'      strExtData=���:������鲿λ��Ϣ,����:���������������������Ϣ,�����޸�������

    Dim rsTmp As New ADODB.Recordset
    Dim rsMore As New ADODB.Recordset '������Ŀ��ϸ��Ϣ
    Dim strSQL As String, lngCopyRow As Long
    Dim blnFirst As Boolean, lngTmp As Long, i As Long
    Dim strҽ�� As String, lngҽ��ID As Long
    Dim strҩ��IDs As String, sng���� As Single
    
    Dim strƵ�� As String, intƵ�ʴ��� As Integer, intƵ�ʼ�� As Integer, str�����λ As String
        
    On Error GoTo errH
    
    'ȡ��һ����һ��Ч��,ĳЩ����ȱʡ�������ͬ
    lngCopyRow = GetPreRow(lngRow)
    If lngCopyRow = -1 Then lngCopyRow = GetNextRow(lngRow)
            
    With vsAdvice
        '��ʼ����ҽ��ȱʡ����
        .RowData(lngRow) = zlDatabase.GetNextId("����ҽ����¼")
        .TextMatrix(lngRow, COL_EDIT) = 1 '����
        .TextMatrix(lngRow, COL_Ӥ��) = cboӤ��.ListIndex
        .TextMatrix(lngRow, COL_״̬) = 1 '�¿�
        
        '���:��������,��ǰ��ռ������ź�,�������������
        .TextMatrix(lngRow, COL_���) = GetCurRow���(lngRow)
        Call AdviceSetҽ�����(lngRow + 1, 1)
        
        .TextMatrix(lngRow, COL_���) = rsInput!���ID
        .TextMatrix(lngRow, COL_����) = rsInput!���� '�����ƿ����Ǳ���
        .TextMatrix(lngRow, COL_������ĿID) = rsInput!������ĿID
        .TextMatrix(lngRow, COL_�շ�ϸĿID) = Nvl(rsInput!�շ�ϸĿID)
        
        'ҩƷ�Ĺ����Ϣ
        If Not IsNull(rsInput!�շ�ϸĿID) Then
            strSQL = "Select Nvl(C.����,A.����) as ����," & _
                " B.����ϵ��,B.���ﵥλ,B.�����װ,B.�ɷ����" & _
                " From �շ���ĿĿ¼ A,ҩƷ��� B,�շ���Ŀ���� C" & _
                " Where A.ID=B.ҩƷID And A.ID=[1]" & _
                " And A.ID=C.�շ�ϸĿID(+) And C.����(+)=1 And C.����(+)=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsInput!�շ�ϸĿID), IIF(gbln��Ʒ��, 3, 1))
            .TextMatrix(lngRow, COL_����) = rsTmp!���� '������������ʽ�������
            .TextMatrix(lngRow, COL_����ϵ��) = rsTmp!����ϵ��
            .TextMatrix(lngRow, COL_���ﵥλ) = rsTmp!���ﵥλ
            .TextMatrix(lngRow, COL_�����װ) = rsTmp!�����װ
            .TextMatrix(lngRow, COL_�ɷ����) = Nvl(rsTmp!�ɷ����, 0)
        End If
        
        'ҩƷ����
        If InStr(",5,6,", rsInput!���ID) > 0 Then
            strSQL = "Select �������,ҩƷ����,��������,����ְ�� From ҩƷ���� Where ҩ��ID=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsInput!������ĿID))
            If Not rsTmp.EOF Then
                .TextMatrix(lngRow, COL_�������) = Nvl(rsTmp!�������)
                .TextMatrix(lngRow, COL_ҩƷ����) = Nvl(rsTmp!ҩƷ����)
                .TextMatrix(lngRow, COL_��������) = Nvl(rsTmp!��������)
                .TextMatrix(lngRow, COL_����ְ��) = Nvl(rsTmp!����ְ��)
            End If
        End If
        
        '��ȡ����������Ŀ��Ϣ
        '----------------------------------------------------------------------------
        strSQL = "Select A.*" & _
            " From �����÷����� A,������ĿĿ¼ B" & _
            " Where A.�÷�ID=B.ID And (Nvl(A.����,0)=0 Or B.������� IN(1,3))" & _
            " And A.��ĿID=[1]"
        strSQL = "Select A.*,Nvl(B.����,0) as ����,B.�÷�ID," & _
            " B.Ƶ��,B.���˼���,B.С������,B.ҽ������,B.�Ƴ�" & _
            " From ������ĿĿ¼ A,(" & strSQL & ") B" & _
            " Where A.ID=B.��ĿID(+) And A.ID=[1]" & _
            " Order by ����"
        Set rsMore = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsInput!������ĿID))
                
        If IsNull(rsInput!�շ�ϸĿID) Then '������������ʽ��������
            .TextMatrix(lngRow, COL_����) = rsMore!����
        End If
                
        If InStr(",5,6,", rsInput!���ID) > 0 Or (Nvl(rsMore!ִ��Ƶ��, 0) = 0 And InStr(",1,2,", Nvl(rsMore!���㷽ʽ, 0)) > 0) Then
            .TextMatrix(lngRow, COL_������λ) = Nvl(rsMore!���㵥λ) 'ҩƷΪ������λ
        End If
        
        If InStr(",5,6,", rsInput!���ID) > 0 Then
            '�С�����ҩ������������λ�������ﵥλ
            .TextMatrix(lngRow, COL_������λ) = .TextMatrix(lngRow, COL_���ﵥλ)
        Else
            '��������Ҫ��������(���㵥λ)
            '���Ϊһ���Ի�ƴ�����ȱʡ����Ϊ1
            If Nvl(rsMore!ִ��Ƶ��, 0) = 1 Or Nvl(rsMore!���㷽ʽ, 0) = 3 Then
                .TextMatrix(lngRow, COL_����) = 1
            End If
            .TextMatrix(lngRow, COL_������λ) = Nvl(rsMore!���㵥λ)
        End If
        
        .TextMatrix(lngRow, COL_���㷽ʽ) = Nvl(rsMore!���㷽ʽ, 0)
        .TextMatrix(lngRow, COL_Ƶ������) = Nvl(rsMore!ִ��Ƶ��, 0)
        .TextMatrix(lngRow, COL_��������) = Nvl(rsMore!��������)
        If InStr(",5,6,7,", rsInput!���ID) = 0 Then
            .TextMatrix(lngRow, COL_��������) = Nvl(rsMore!¼������)
        End If

        '�걾��λ
        If InStr(",5,6,", rsInput!���ID) > 0 Then
            .TextMatrix(lngRow, COL_�걾��λ) = rsInput!���� '��¼ҩƷ����ʱѡ�������
        Else
            .TextMatrix(lngRow, COL_�걾��λ) = Nvl(rsMore!�걾��λ)
        End If
        
        '�Ƽ�����
        .TextMatrix(lngRow, COL_�Ƽ�����) = Nvl(rsMore!�Ƽ�����, 0)
    
        'ִ������:������Ŀʱ������Ŀ����,ҩƷ=4-ָ������,һ����ҩ����ͬ
        If InStr(",5,6,", rsInput!���ID) > 0 Then
            If lngGroupRow <> 0 Then
                .TextMatrix(lngRow, COL_ִ������) = .TextMatrix(lngGroupRow, COL_ִ������)
            Else
                .TextMatrix(lngRow, COL_ִ������) = 4
            End If
        Else
            .TextMatrix(lngRow, COL_ִ������) = Nvl(rsMore!ִ�п���, 0)
        End If
            
        '����ҽ���Ϳ���
        If lngGroupRow = 0 Then
            .TextMatrix(lngRow, COL_����ҽ��) = UserInfo.����
            .TextMatrix(lngRow, COL_��������ID) = Get��������ID(UserInfo.ID, mlng���˿���id, 1)
        Else
            .TextMatrix(lngRow, COL_����ҽ��) = .TextMatrix(lngGroupRow, COL_����ҽ��)
            .TextMatrix(lngRow, COL_��������ID) = .TextMatrix(lngGroupRow, COL_��������ID)
        End If
    
        'ִ�п���:ҩƷȱʡ����һ����ͬ,һ����ҩ����ͬ
        If InStr(",5,6,", rsInput!���ID) > 0 Then
            If lngGroupRow <> 0 Then
                .TextMatrix(lngRow, COL_ִ�п���ID) = .TextMatrix(lngGroupRow, COL_ִ�п���ID)
            ElseIf lngCopyRow <> -1 Then
                If rsInput!���ID = .TextMatrix(lngCopyRow, COL_���) Then
                    strҩ��IDs = Get����ҩ��IDs(rsInput!���ID, rsInput!������ĿID, Nvl(rsInput!�շ�ϸĿID, 0), mlng���˿���id, 1)
                    If InStr("," & strҩ��IDs & ",", "," & .TextMatrix(lngCopyRow, COL_ִ�п���ID) & ",") > 0 Then
                        .TextMatrix(lngRow, COL_ִ�п���ID) = .TextMatrix(lngCopyRow, COL_ִ�п���ID)
                    End If
                End If
            End If
        End If
        If Val(.TextMatrix(lngRow, COL_ִ�п���ID)) = 0 Then
            If rsInput!���ID = "Z" And InStr(",1,2,", Nvl(rsMore!��������, 0)) > 0 Then
                '���ۻ�סԺҽ��ȡ�ٴ�����(����ִ������)
                If Nvl(rsMore!��������, 0) = 1 Then
                    '����:���������סԺ���ٴ�����
                    Call Get�ٴ�����(3, , lngTmp, , Not gbln�������Ҷ���)
                ElseIf Nvl(rsMore!��������, 0) = 2 Then
                    'סԺ:סԺ�ٴ�����
                    Call Get�ٴ�����(2, , lngTmp, , Not gbln�������Ҷ���)
                End If
                .TextMatrix(lngRow, COL_ִ�п���ID) = lngTmp
            ElseIf InStr(",0,5,", Val(.TextMatrix(lngRow, COL_ִ������))) = 0 Then
                'ִ������Ϊ(0-����,5-Ժ��ִ��)��ִ�п���
                '֮ǰ�������������ID
                .TextMatrix(lngRow, COL_ִ�п���ID) = Get����ִ�п���ID(mlng����ID, 0, rsInput!���ID, rsInput!������ĿID, _
                    Nvl(rsInput!�շ�ϸĿID, 0), Nvl(rsMore!ִ�п���, 0), mlng���˿���id, Val(.TextMatrix(lngRow, COL_��������ID)), 1, 1, InStr(",5,6,", rsInput!���ID) > 0)
            End If
        End If
        
        'ҩƷ���
        If InStr(",5,6,", rsInput!���ID) > 0 And Not IsNull(rsInput!�շ�ϸĿID) Then
            Call GetDrugStock(lngRow)
        End If
        
        'ִ��Ƶ��:��ѡƵ��,һ���Ի������
        If True Then 'If Nvl(rsMore!ִ��Ƶ��, 0) = 0 Then
            'ȱʡ����һ��������ͬ
            If lngCopyRow <> -1 Then
                If GetƵ�ʷ�Χ(lngRow) = GetƵ�ʷ�Χ(lngCopyRow) Then
                    If Val(.TextMatrix(lngCopyRow, COL_EDIT)) = 1 And .TextMatrix(lngCopyRow, COL_Ƶ��) <> "" _
                        And Not (.TextMatrix(lngRow, COL_���) = "7" And Not RowIn�䷽��(lngCopyRow)) _
                        And Not (.TextMatrix(lngRow, COL_���) <> "7" And RowIn�䷽��(lngCopyRow)) Then
                        .TextMatrix(lngRow, COL_Ƶ��) = .TextMatrix(lngCopyRow, COL_Ƶ��)
                        .TextMatrix(lngRow, COL_Ƶ�ʴ���) = .TextMatrix(lngCopyRow, COL_Ƶ�ʴ���)
                        .TextMatrix(lngRow, COL_Ƶ�ʼ��) = .TextMatrix(lngCopyRow, COL_Ƶ�ʼ��)
                        .TextMatrix(lngRow, COL_�����λ) = .TextMatrix(lngCopyRow, COL_�����λ)
                    End If
                End If
            End If
            '��ȡȱʡƵ��
            If .TextMatrix(lngRow, COL_Ƶ��) = "" Then
                Call GetȱʡƵ��(GetƵ�ʷ�Χ(lngRow), strƵ��, intƵ�ʴ���, intƵ�ʼ��, str�����λ)
                .TextMatrix(lngRow, COL_Ƶ��) = strƵ��
                .TextMatrix(lngRow, COL_Ƶ�ʴ���) = intƵ�ʴ���
                .TextMatrix(lngRow, COL_Ƶ�ʼ��) = intƵ�ʼ��
                .TextMatrix(lngRow, COL_�����λ) = str�����λ
            End If
        End If
        
        '�У�����ҩ��һЩȱʡ��Ϣ
        If InStr(",5,6,", rsInput!���ID) > 0 Then
            'ִ��Ƶ��
            If lngGroupRow <> 0 Then
                'һ����ҩ����ͬ
                .TextMatrix(lngRow, COL_Ƶ��) = .TextMatrix(lngGroupRow, COL_Ƶ��)
                .TextMatrix(lngRow, COL_Ƶ�ʴ���) = .TextMatrix(lngGroupRow, COL_Ƶ�ʴ���)
                .TextMatrix(lngRow, COL_Ƶ�ʼ��) = .TextMatrix(lngGroupRow, COL_Ƶ�ʼ��)
                .TextMatrix(lngRow, COL_�����λ) = .TextMatrix(lngGroupRow, COL_�����λ)
                .TextMatrix(lngRow, COL_ִ��ʱ��) = .TextMatrix(lngGroupRow, COL_ִ��ʱ��)
            End If
            
            'ȷ��������ҩ������
            '1.����Ϊһ��Ƶ����������
            '2-���Ƴ���Ϊ�Ƴ�����(Ӧ����һ��Ƶ����������)
            sng���� = msng����
            If mbln���� Then
                If .TextMatrix(lngRow, COL_�����λ) = "��" Then
                    If 7 > sng���� Then sng���� = 7
                ElseIf .TextMatrix(lngRow, COL_�����λ) = "��" Then
                    If Val(.TextMatrix(lngRow, COL_Ƶ�ʼ��)) > sng���� Then
                        sng���� = Val(.TextMatrix(lngRow, COL_Ƶ�ʼ��))
                    End If
                ElseIf .TextMatrix(lngRow, COL_�����λ) = "Сʱ" Then
                    If Val(.TextMatrix(lngRow, COL_Ƶ�ʼ��)) \ 24 > sng���� Then
                        sng���� = Val(.TextMatrix(lngRow, COL_Ƶ�ʼ��)) \ 24
                    End If
                End If
                If sng���� = 0 Then sng���� = 1
            End If
            
            rsMore.Filter = "����>0" 'ȡ��һ�ָ�ҩ;����Ϊȱʡ����
            If Not rsMore.EOF Then
                '����һ����ҩʱ,���õ�ȱʡ�÷�Ƶ������
                If lngGroupRow = 0 Then
                    If Not IsNull(rsMore!�÷�ID) Then lng��ҩ;��ID = rsMore!�÷�ID
                    If Not IsNull(rsMore!Ƶ��) Then
                        Call GetƵ����Ϣ_����(rsMore!Ƶ��, strƵ��, intƵ�ʴ���, intƵ�ʼ��, str�����λ)
                        .TextMatrix(lngRow, COL_Ƶ��) = strƵ��
                        .TextMatrix(lngRow, COL_Ƶ�ʴ���) = intƵ�ʴ���
                        .TextMatrix(lngRow, COL_Ƶ�ʼ��) = intƵ�ʼ��
                        .TextMatrix(lngRow, COL_�����λ) = str�����λ
                    End If
                End If
                
                'ҽ������
                .TextMatrix(lngRow, COL_ҽ������) = Nvl(rsMore!ҽ������) 'һ��Ϊ��ҩ;����˵��
                
                'ҩƷ����
                If mint���� > 12 Then
                    If Nvl(rsMore!���˼���, 0) <> 0 Then
                        .TextMatrix(lngRow, COL_����) = FormatEx(rsMore!���˼���, 5)
                    End If
                Else
                    If Nvl(rsMore!С������, 0) <> 0 Then
                        .TextMatrix(lngRow, COL_����) = FormatEx(rsMore!С������, 5)
                    ElseIf Nvl(rsMore!���˼���, 0) <> 0 Then
                        .TextMatrix(lngRow, COL_����) = FormatEx(rsMore!���˼��� * (mint���� + 2) * 5 / 100, 5)
                    End If
                End If
                If Val(.TextMatrix(lngRow, COL_����)) = 0 Then .TextMatrix(lngRow, COL_����) = ""
                
                'ҩƷ��������:�����װ
                If Nvl(rsMore!�Ƴ�, 1) > sng���� Then sng���� = Nvl(rsMore!�Ƴ�, 1)
                If .TextMatrix(lngRow, COL_Ƶ��) <> "" And Val(.TextMatrix(lngRow, COL_����)) <> 0 _
                    And Val(.TextMatrix(lngRow, COL_����ϵ��)) <> 0 And Val(.TextMatrix(lngRow, COL_�����װ)) <> 0 Then
                    '�����Ƴ����Ϊ��������ҩ������
                    .TextMatrix(lngRow, COL_����) = FormatEx(CalcȱʡҩƷ����( _
                            Val(.TextMatrix(lngRow, COL_����)), sng����, _
                            Val(.TextMatrix(lngRow, COL_Ƶ�ʴ���)), _
                            Val(.TextMatrix(lngRow, COL_Ƶ�ʼ��)), _
                            .TextMatrix(lngRow, COL_�����λ), _
                            .TextMatrix(lngRow, COL_ִ��ʱ��), _
                            Val(.TextMatrix(lngRow, COL_����ϵ��)), _
                            Val(.TextMatrix(lngRow, COL_�����װ)), _
                            Val(.TextMatrix(lngRow, COL_�ɷ����))), 5)
                End If
            End If
            
            '��¼ȱʡ����
            If mbln���� Then .TextMatrix(lngRow, COL_����) = sng����
        End If
        
        If rsMore.Filter <> 0 Then rsMore.Filter = 0
        
        'ִ��ʱ��:"��ѡƵ��"��ҩƷ
        If Nvl(rsMore!ִ��Ƶ��, 0) = 0 Or InStr(",5,6,", rsInput!���ID) > 0 Then
            If .TextMatrix(lngRow, COL_ִ��ʱ��) = "" Then
                If lngCopyRow <> -1 Then '����һ����ͬ
                    If .TextMatrix(lngRow, COL_Ƶ��) = .TextMatrix(lngCopyRow, COL_Ƶ��) Then
                        .TextMatrix(lngRow, COL_ִ��ʱ��) = .TextMatrix(lngCopyRow, COL_ִ��ʱ��)
                    End If
                End If
                If .TextMatrix(lngRow, COL_ִ��ʱ��) = "" Then 'ȱʡʱ�䷽��
                    .TextMatrix(lngRow, COL_ִ��ʱ��) = Getȱʡʱ��(1, .TextMatrix(lngRow, COL_Ƶ��), lng��ҩ;��ID)
                End If
            End If
        End If
        
        '����(����Ŀ�޹�)
        '---------------------------------------------------------------------
        If lngGroupRow = 0 Then
            If IsDate(txt��ʼʱ��.Text) Then
                .TextMatrix(lngRow, COL_��ʼʱ��) = Format(txt��ʼʱ��.Text, "MM-dd HH:mm")
                .Cell(flexcpData, lngRow, COL_��ʼʱ��) = txt��ʼʱ��.Text
            End If
            
            .TextMatrix(lngRow, COL_����ʱ��) = Format(zlDatabase.Currentdate, "MM-dd HH:mm")
            .Cell(flexcpData, lngRow, COL_����ʱ��) = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
    
            .TextMatrix(lngRow, COL_��־) = chk����.Value
        Else
            .TextMatrix(lngRow, COL_��ʼʱ��) = .TextMatrix(lngGroupRow, COL_��ʼʱ��)
            .Cell(flexcpData, lngRow, COL_��ʼʱ��) = .Cell(flexcpData, lngGroupRow, COL_��ʼʱ��)
            
            .TextMatrix(lngRow, COL_����ʱ��) = .TextMatrix(lngGroupRow, COL_����ʱ��)
            .Cell(flexcpData, lngRow, COL_����ʱ��) = .Cell(flexcpData, lngGroupRow, COL_����ʱ��)
            
            .TextMatrix(lngRow, COL_��־) = .TextMatrix(lngGroupRow, COL_��־)
        End If
        
        '������־
        blnFirst = True
        If InStr(",5,6,", rsInput!���ID) > 0 Then
            If lngGroupRow <> 0 Then
                lngTmp = .FindRow(CLng(.TextMatrix(lngGroupRow, COL_���ID)), lngGroupRow + 1)
                If lngTmp > lngRow Then
                    blnFirst = False
                End If
            End If
        End If
        If blnFirst Then
            If Val(.TextMatrix(lngRow, COL_��־)) = 1 Then
                Set .Cell(flexcpPicture, lngRow, COL_F��־) = imgFlag.ListImages("����").Picture
                .Cell(flexcpPictureAlignment, lngRow, COL_F��־) = 4
            End If
        End If
        
        
        '�����д������֮��������,�����ҽ������
        '-------------------------------------------------------------------------
        If InStr(",5,6,", rsInput!���ID) > 0 Then
            '����һ����ҩ;����Ŀ,���������
            If lng��ҩ;��ID <> 0 Then
                .TextMatrix(lngRow, COL_�÷�) = Get��Ŀ����(lng��ҩ;��ID)
            End If
            If lngGroupRow <> 0 Then
                'һ����ҩ�Ĺ�����ͬ�ĸ�ҩ;����
                lngTmp = .FindRow(CLng(.TextMatrix(lngGroupRow, COL_���ID)), lngGroupRow + 1)
                If lngTmp > lngRow Then
                    .TextMatrix(lngRow, COL_���ID) = .TextMatrix(lngGroupRow, COL_���ID)
                Else
                    '��������ǽ�Ϊ��ʹ��һ����ҩ����ͬ����
                    .TextMatrix(lngRow, COL_���ID) = AdviceSet��ҩ;��(lngRow, lng��ҩ;��ID)
                End If
            Else '���������ĳ�ҩ���������ĸ�ҩ;����
                .TextMatrix(lngRow, COL_���ID) = AdviceSet��ҩ;��(lngRow, lng��ҩ;��ID)
            End If
            
            '���龫����ɫ��ʶ
            If InStr(",����ҩ,����ҩ,����ҩ,", .TextMatrix(lngRow, COL_�������)) > 0 _
                And .TextMatrix(lngRow, COL_�������) <> "" Then
                .Cell(flexcpFontBold, lngRow, COL_ҽ������) = True
            End If
        ElseIf rsInput!���ID = "D" And strExtData <> "" Then
            '������ϲ�λ��
            Call AdviceSet�������(1, lngRow, strExtData)
        ElseIf rsInput!���ID = "F" And strExtData <> "" Then
            '�����ĸ���������������Ŀ��
            Call AdviceSet�������(2, lngRow, strExtData)
        End If
        
        .TextMatrix(lngRow, COL_ҽ������) = AdviceTextMake(lngRow)
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub AdviceSet�������(ByVal int���� As Integer, ByVal lngRow As Long, ByVal strDataIDs As String)
'���ܣ�1.��������ָ����������Ŀ�Ĳ�λ��,�����������������Ŀ���޸Ĳ�λ
'      2.��������ָ��������Ŀ�ĸ���������������Ŀ��,����������������Ŀ��������Ŀ�ĸ���������������Ŀ
'������int����=1=�����鲿λ��Ŀ,2=������������������Ŀ
'      lngRow=��ǰ������
'      strDataIDs=���:������鲿λ��Ϣ,����:��������������������Ŀ��Ϣ,���п���û�и�������������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim arrIDs As Variant
    
    On Error GoTo errH
            
    'ɾ�����еļ�鲿λ�л����еĸ���������������Ŀ��(�޸���ʱ)
    Call Delete�������(lngRow)
    
    '���¼��벿λ�л򸽼������м�������Ŀ��
    If int���� = 2 Then
        strDataIDs = Trim(Replace(strDataIDs, ";", ","))
        If Left(strDataIDs, 1) = "," Then strDataIDs = Mid(strDataIDs, 2)
        If Right(strDataIDs, 1) = "," Then strDataIDs = Mid(strDataIDs, 1, Len(strDataIDs) - 1)
    End If
    
    If strDataIDs <> "" Then
        strSQL = "Select * From ������ĿĿ¼ Where ID IN(" & strDataIDs & ")"
        zlDatabase.OpenRecordset rsTmp, strSQL, Me.Caption 'In
        If Not rsTmp.EOF Then
            arrIDs = Split(strDataIDs, ",")
            For i = 0 To UBound(arrIDs) '���û�������Ŀ˳��
                rsTmp.Filter = "ID=" & CStr(arrIDs(i)) '������EOF
                
                With vsAdvice
                    .AddItem "", lngRow + i + 1
                    .RowHidden(lngRow + i + 1) = True
                    
                    .RowData(lngRow + i + 1) = zlDatabase.GetNextId("����ҽ����¼")
                    .TextMatrix(lngRow + i + 1, COL_���ID) = .RowData(lngRow)
                    
                    .TextMatrix(lngRow + i + 1, COL_EDIT) = 1 '����
                    
                    .TextMatrix(lngRow + i + 1, COL_Ӥ��) = cboӤ��.ListIndex
                    .TextMatrix(lngRow + i + 1, COL_���) = Val(.TextMatrix(lngRow, COL_���)) + i + 1
                    .TextMatrix(lngRow + i + 1, COL_״̬) = 1 '�¿�
                    
                    .TextMatrix(lngRow + i + 1, COL_���) = rsTmp!���
                    .TextMatrix(lngRow + i + 1, COL_������ĿID) = rsTmp!ID
                    .TextMatrix(lngRow + i + 1, COL_���㷽ʽ) = Nvl(rsTmp!���㷽ʽ, 0)
                    .TextMatrix(lngRow + i + 1, COL_Ƶ������) = Nvl(rsTmp!ִ��Ƶ��, 0)
                    .TextMatrix(lngRow + i + 1, COL_��������) = Nvl(rsTmp!��������)
                    .TextMatrix(lngRow + i + 1, COL_��������) = Nvl(rsTmp!¼������)
                    
                    .TextMatrix(lngRow + i + 1, COL_�걾��λ) = Nvl(rsTmp!�걾��λ)
                    .TextMatrix(lngRow + i + 1, COL_ҽ������) = rsTmp!����
                    
                    .TextMatrix(lngRow + i + 1, COL_�Ƽ�����) = Nvl(rsTmp!�Ƽ�����, 0)
                    
                    .TextMatrix(lngRow + i + 1, COL_����) = .TextMatrix(lngRow, COL_����)
                    .TextMatrix(lngRow + i + 1, COL_����) = .TextMatrix(lngRow, COL_����)
    
                    .TextMatrix(lngRow + i + 1, COL_ִ��ʱ��) = .TextMatrix(lngRow, COL_ִ��ʱ��)
                    .TextMatrix(lngRow + i + 1, COL_Ƶ��) = .TextMatrix(lngRow, COL_Ƶ��)
                    .TextMatrix(lngRow + i + 1, COL_Ƶ�ʴ���) = .TextMatrix(lngRow, COL_Ƶ�ʴ���)
                    .TextMatrix(lngRow + i + 1, COL_Ƶ�ʼ��) = .TextMatrix(lngRow, COL_Ƶ�ʼ��)
                    .TextMatrix(lngRow + i + 1, COL_�����λ) = .TextMatrix(lngRow, COL_�����λ)
                    
                    'ִ������:������Ŀ��������
                    .TextMatrix(lngRow + i + 1, COL_ִ������) = Nvl(rsTmp!ִ�п���, 0)
                    
                    '������Ժ��ִ����ִ�п���,����������ִ�п���
                    '���򲻹���ִ�п�������,һ�������������Ӧ����ͬ
                    If InStr(",0,5,", Nvl(rsTmp!ִ�п���, 0)) > 0 Then
                        .TextMatrix(lngRow + i + 1, COL_ִ�п���ID) = 0
                    Else
                        If rsTmp!��� = "G" Then
                            .TextMatrix(lngRow + i + 1, COL_ִ�п���ID) = Get����ִ�п���ID(mlng����ID, 0, rsTmp!���, rsTmp!ID, 0, _
                                Nvl(rsTmp!ִ�п���, 0), mlng���˿���id, Val(.TextMatrix(lngRow, COL_��������ID)), 1, 1)
                        Else
                            .TextMatrix(lngRow + i + 1, COL_ִ�п���ID) = .TextMatrix(lngRow, COL_ִ�п���ID)
                        End If
                    End If
                    
                    .TextMatrix(lngRow + i + 1, COL_��ʼʱ��) = .TextMatrix(lngRow, COL_��ʼʱ��)
                    .Cell(flexcpData, lngRow + i + 1, COL_��ʼʱ��) = .Cell(flexcpData, lngRow, COL_��ʼʱ��)
                    
                    .TextMatrix(lngRow + i + 1, COL_��������ID) = .TextMatrix(lngRow, COL_��������ID)
                    .TextMatrix(lngRow + i + 1, COL_����ҽ��) = .TextMatrix(lngRow, COL_����ҽ��)
                    
                    .TextMatrix(lngRow + i + 1, COL_����ʱ��) = .TextMatrix(lngRow, COL_����ʱ��)
                    .Cell(flexcpData, lngRow + i + 1, COL_����ʱ��) = .Cell(flexcpData, lngRow, COL_����ʱ��)
                    
                    .TextMatrix(lngRow + i + 1, COL_��־) = .TextMatrix(lngRow, COL_��־)
                End With
            Next
                
            '�������
            Call AdviceSetҽ�����(lngRow + UBound(arrIDs) + 2, UBound(arrIDs) + 1)
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function AdviceSet��ҩ;��(ByVal lngRow As Long, ByVal lng��ҩ;��ID As Long, Optional strִ������ As String) As Long
'���ܣ�Ϊ¼����У�����ҩ���ö�Ӧ�ĸ�ҩ;����(�������޸�)
'������lngRow=Ҫ�����ҩ;����ҩƷ��
'      lng��ҩ;��ID=��ҩ;��ID
'      strִ������=�޸ĸ�ҩ;��ʱ,��ǰ�������õ�ִ������
'���أ������õĸ�ҩ;���е�ҽ��ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lngNewRow As Long
    Dim blnNew As Boolean
    
    On Error GoTo errH
    strSQL = "Select * From ������ĿĿ¼ Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��ҩ;��ID)
    If rsTmp.EOF Then lng��ҩ;��ID = 0 'û�����ݣ��������Ա��ֹ�ϵ
        
    With vsAdvice
        If Val(.TextMatrix(lngRow, COL_���ID)) = 0 Then 'δ����"���ID"ʱ
            blnNew = True
            lngNewRow = lngRow + 1
            .AddItem "", lngNewRow
            .RowHidden(lngNewRow) = True
        Else
            '�޸�ҽ��������ʱ�������ø�ҩ;������(���Ǹ���������Ŀ)
            blnNew = False
            lngNewRow = .FindRow(CLng(.TextMatrix(lngRow, COL_���ID)), lngRow + 1)
        End If
        
        '��Ч���ݣ�����,�շ�ϸĿID,����ϵ��,���ﵥλ,�����װ,�걾��λ,ҽ������,����,����,�÷�
        If blnNew Then
            .RowData(lngNewRow) = zlDatabase.GetNextId("����ҽ����¼")
            .TextMatrix(lngNewRow, COL_EDIT) = 1 '����
            .TextMatrix(lngNewRow, COL_���) = Val(.TextMatrix(lngRow, COL_���)) + 1
        Else
            'ҽ��ID(RowData),���:���ֲ���
            If InStr(",0,3,", .TextMatrix(lngNewRow, COL_EDIT)) > 0 Then
                .TextMatrix(lngNewRow, COL_EDIT) = 2 '��־Ϊ�����޸�
                .TextMatrix(lngNewRow, COL_״̬) = 1 '�޸ĺ��Ϊ�¿�
            End If
        End If
        
        .TextMatrix(lngNewRow, COL_Ӥ��) = cboӤ��.ListIndex
        .TextMatrix(lngNewRow, COL_״̬) = 1 '�¿�
        
        .TextMatrix(lngNewRow, COL_���) = "E" '��ҩ;����������
        .TextMatrix(lngNewRow, COL_������ĿID) = lng��ҩ;��ID
        '���û��ȷ����ҩ;������ʱ�����õ�����
        If Not rsTmp.EOF Then
            .TextMatrix(lngNewRow, COL_���㷽ʽ) = Nvl(rsTmp!���㷽ʽ, 0)
            .TextMatrix(lngNewRow, COL_Ƶ������) = Nvl(rsTmp!ִ��Ƶ��, 0)
            .TextMatrix(lngNewRow, COL_��������) = Nvl(rsTmp!��������)
            .TextMatrix(lngNewRow, COL_ҽ������) = rsTmp!����
            
            .TextMatrix(lngNewRow, COL_�Ƽ�����) = Nvl(rsTmp!�Ƽ�����, 0)
            
            'ִ������:ȱʡ������Ŀ����,�޸�ʱ���ݵ�ǰ��������
            If strִ������ = "" Then
                .TextMatrix(lngNewRow, COL_ִ������) = Nvl(rsTmp!ִ�п���, 0)
            Else
                .TextMatrix(lngNewRow, COL_ִ������) = Decode(strִ������, "��Ժ��ҩ", 5, Nvl(rsTmp!ִ�п���, 0))
            End If
            
            '��ҩ;�����δ����ִ�п���,��ȱʡΪ�������ڲ���(����Ҫ��Ϊ�������ڿ���!!)
            If InStr(",0,5,", Val(.TextMatrix(lngNewRow, COL_ִ������))) = 0 Then
                .TextMatrix(lngNewRow, COL_ִ�п���ID) = Get����ִ�п���ID(mlng����ID, 0, "E", lng��ҩ;��ID, 0, _
                    Nvl(rsTmp!ִ�п���, 0), mlng���˿���id, Val(.TextMatrix(lngRow, COL_��������ID)), 1, 1)
            Else
                .TextMatrix(lngNewRow, COL_ִ�п���ID) = 0
            End If
        End If
        
        '��ҩ;��������ҩƷ��ͬ
        .TextMatrix(lngNewRow, COL_����) = .TextMatrix(lngRow, COL_����)
        
        .TextMatrix(lngNewRow, COL_Ƶ��) = .TextMatrix(lngRow, COL_Ƶ��)
        .TextMatrix(lngNewRow, COL_Ƶ�ʴ���) = .TextMatrix(lngRow, COL_Ƶ�ʴ���)
        .TextMatrix(lngNewRow, COL_Ƶ�ʼ��) = .TextMatrix(lngRow, COL_Ƶ�ʼ��)
        .TextMatrix(lngNewRow, COL_�����λ) = .TextMatrix(lngRow, COL_�����λ)
        .TextMatrix(lngNewRow, COL_ִ��ʱ��) = .TextMatrix(lngRow, COL_ִ��ʱ��)
        
        .TextMatrix(lngNewRow, COL_��ʼʱ��) = .TextMatrix(lngRow, COL_��ʼʱ��)
        .Cell(flexcpData, lngNewRow, COL_��ʼʱ��) = .Cell(flexcpData, lngRow, COL_��ʼʱ��)
        
        .TextMatrix(lngNewRow, COL_��������ID) = .TextMatrix(lngRow, COL_��������ID)
        .TextMatrix(lngNewRow, COL_����ҽ��) = .TextMatrix(lngRow, COL_����ҽ��)
        
        .TextMatrix(lngNewRow, COL_����ʱ��) = .TextMatrix(lngRow, COL_����ʱ��)
        .Cell(flexcpData, lngNewRow, COL_����ʱ��) = .Cell(flexcpData, lngRow, COL_����ʱ��)
        
        .TextMatrix(lngNewRow, COL_��־) = .TextMatrix(lngRow, COL_��־)
            
        '����������
        If blnNew Then Call AdviceSetҽ�����(lngNewRow + 1, 1)
        
        AdviceSet��ҩ;�� = .RowData(lngNewRow)
    End With
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub AdviceChange()
'���ܣ����ݵ�ǰҽ����Ƭ�е����ݣ����µ�ǰҽ������
'˵��������ListIndex=-1����Ӧҽ�����������ݵģ�����ԭ���ݲ�����
    Dim lngRow As Long, lngBeginRow As Long
    Dim intƵ�ʴ��� As Integer, intƵ�ʼ�� As Integer, str�����λ As String
    Dim blnCurDo As Boolean, blnOtherDo As Boolean
    Dim lngTmp As Long, blnTmp As Boolean, strTmp As String
    Dim strCurDate As String, lng��������ID As Long
    Dim blnReInRow As Boolean, i As Long, j As Long
    
    With vsAdvice
        lngRow = .Row
        
        If .RowData(lngRow) = 0 Then Call ClearItemTag: Exit Sub '����༭��־
        
        If RowIn�䷽��(lngRow) Then
            '��ҩ�䷽
            lngBeginRow = .FindRow(CStr(.RowData(lngRow)), , COL_���ID)
            For i = lngBeginRow To lngRow
                '�޸Ĵ����䷽������������(�����巨���÷�)
                If IsDate(txt��ʼʱ��.Text) And txt��ʼʱ��.Tag <> "" Then
                    .TextMatrix(i, COL_��ʼʱ��) = Format(txt��ʼʱ��.Text, "MM-dd HH:mm")
                    .Cell(flexcpData, i, COL_��ʼʱ��) = txt��ʼʱ��.Text
                    blnCurDo = True
                End If
                If chk����.Visible And chk����.Tag <> "" Then
                    .TextMatrix(i, COL_��־) = chk����.Value
                    If i = lngRow Then '�÷�����ʾ������־
                        If Val(.TextMatrix(i, COL_��־)) = 1 Then
                            Set .Cell(flexcpPicture, i, COL_F��־) = imgFlag.ListImages("����").Picture
                        Else
                            Set .Cell(flexcpPicture, i, COL_F��־) = Nothing
                        End If
                        .Cell(flexcpPictureAlignment, i, COL_F��־) = 4
                    End If
                    blnCurDo = True
                End If
                If txt����.Enabled And IsNumeric(txt����.Text) And txt����.Tag <> "" Then
                    .TextMatrix(i, COL_����) = FormatEx(Val(txt����.Text), 5)
                    blnCurDo = True
                End If
                If txtƵ��.Enabled And cmdƵ��.Tag <> "" And txtƵ��.Tag <> "" Then
                    .TextMatrix(i, COL_Ƶ��) = txtƵ��.Text
                    Call GetƵ����Ϣ_����(txtƵ��.Text, intƵ�ʴ���, intƵ�ʼ��, str�����λ, 2) '��ҽ��Χ
                    .TextMatrix(i, COL_Ƶ�ʴ���) = intƵ�ʴ���
                    .TextMatrix(i, COL_Ƶ�ʼ��) = intƵ�ʼ��
                    .TextMatrix(i, COL_�����λ) = str�����λ
                    blnCurDo = True
                End If
                If cboִ��ʱ��.Tag <> "" Then
                    .TextMatrix(i, COL_ִ��ʱ��) = cboִ��ʱ��.Text
                    blnCurDo = True
                End If
                
                If .TextMatrix(i, COL_���) = "7" Then
                    '���ĵ��������ҩ��ִ�п���(�÷��巨�ĸĲ���)
                    If cboִ�п���.ListIndex <> -1 And cboִ�п���.Tag <> "" Then
                        .TextMatrix(i, COL_ִ�п���ID) = cboִ�п���.ItemData(cboִ�п���.ListIndex)
                        blnCurDo = True
                    End If
                    
                    'ִ������:�䷽��������ɵ���ҩ��ͬ
                    If cboִ������.Tag <> "" Then
                        .TextMatrix(i, COL_ִ������) = Decode(NeedName(cboִ������.Text), "�Ա�ҩ", 5, 4)
                        If Val(.TextMatrix(i, COL_ִ������)) = 5 Then
                            .TextMatrix(i, COL_ִ�п���ID) = 0
                        ElseIf Val(.TextMatrix(i, COL_ִ�п���ID)) = 0 Then
                            '�ָ�ȱʡִ�п���,ȱʡ��ǰ����ͬ
                            If i = lngBeginRow Then
                                For j = i - 1 To .FixedRows Step -1
                                    If .TextMatrix(j, COL_���) = "7" And Val(.TextMatrix(j, COL_ִ�п���ID)) <> 0 Then
                                        .TextMatrix(i, COL_ִ�п���ID) = .TextMatrix(j, COL_ִ�п���ID)
                                        Exit For
                                    End If
                                Next
                                If Val(.TextMatrix(i, COL_ִ�п���ID)) = 0 Then
                                    .TextMatrix(i, COL_ִ�п���ID) = Get����ִ�п���ID(mlng����ID, 0, .TextMatrix(i, COL_���), Val(.TextMatrix(i, COL_������ĿID)), Val(.TextMatrix(i, COL_�շ�ϸĿID)), 4, mlng���˿���id, 0, 1, 1, True)
                                End If
                            Else
                                .TextMatrix(i, COL_ִ�п���ID) = .TextMatrix(lngBeginRow, COL_ִ�п���ID)
                            End If
                        End If
                        blnReInRow = True '����ִ�п��ұ༭�Ա仯
                        blnCurDo = True
                    End If
                End If
                
                '�޸�ʱ�Զ����²�������
                blnTmp = False
                If cboҽ������.Tag <> "" Or cboִ������.Tag <> "" _
                    Or (Val(cmd�÷�.Tag) <> 0 And txt�÷�.Tag <> "") Then
                    blnTmp = True
                End If
                If blnCurDo Or blnTmp Then
                    '�޸�����������¿���ʱ��
                    If strCurDate = "" Then
                        strCurDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
                    End If
                    .TextMatrix(i, COL_����ʱ��) = Format(strCurDate, "MM-dd HH:mm")
                    .Cell(flexcpData, i, COL_����ʱ��) = strCurDate
                    
                    '��鿪��ҽ��
                    If .TextMatrix(i, COL_����ҽ��) <> UserInfo.���� Then
                        .TextMatrix(i, COL_����ҽ��) = UserInfo.����
                        If lng��������ID = 0 Then
                            lng��������ID = Get��������ID(UserInfo.ID, mlng���˿���id, 1)
                        End If
                        .TextMatrix(i, COL_��������ID) = lng��������ID
                    End If
                End If
                                                    
                If .TextMatrix(i, COL_���) = "E" And i <> lngRow Then lngTmp = i '�巨�к�
                                                    
                '---------------
                If blnCurDo Then '���Ϊ�޸�:0-ԭʼ��,1-������,2-�޸�������,3-�޸������
                    If InStr(",0,3,", .TextMatrix(i, COL_EDIT)) > 0 Then
                        .TextMatrix(i, COL_EDIT) = 2
                        .TextMatrix(i, COL_״̬) = 1 '�޸ĺ��Ϊ�¿�
                        If Not .RowHidden(i) Then Call ReSetColor(i) '�÷��в�����
                    End If
                    mblnNoSave = True '���Ϊδ����
                End If
            Next
            
            '�漰��ҩ�÷��е�����:ֱ�Ӹ��ĵ�ǰ�е�����(�巨�����䷽�༭�в��ܸ�)
            '-----------------------------------------------------------
            blnCurDo = False
                    
            'ҽ������:�Ƿ�����ҩ�÷���(��ʾ��)�е�
            If cboҽ������.Tag <> "" Then
                .TextMatrix(lngRow, COL_ҽ������) = cboҽ������.Text
                blnCurDo = True
            End If
        
            '��ҩ�÷�
            If Val(cmd�÷�.Tag) <> 0 And txt�÷�.Tag <> "" Then
                .TextMatrix(lngRow, COL_������ĿID) = Val(cmd�÷�.Tag)
                .TextMatrix(lngRow, COL_�÷�) = txt�÷�.Text
                
                'ͬʱ���ļƼ����ʺ�ִ������
                .TextMatrix(lngRow, COL_�Ƽ�����) = Nvl(GetItemField("������ĿĿ¼", Val(cmd�÷�.Tag), "�Ƽ�����"), 0)
                i = Nvl(GetItemField("������ĿĿ¼", Val(cmd�÷�.Tag), "ִ�п���"), 0)
                .TextMatrix(lngRow, COL_ִ������) = Decode(NeedName(cboִ������.Text), "��Ժ��ҩ", 5, i)
                If Val(.TextMatrix(lngRow, COL_ִ������)) = 5 Then
                    .TextMatrix(lngRow, COL_ִ�п���ID) = 0
                Else
                    .TextMatrix(lngRow, COL_ִ�п���ID) = Get����ִ�п���ID(mlng����ID, 0, "E", Val(cmd�÷�.Tag), 0, _
                        Val(.TextMatrix(lngRow, COL_ִ������)), mlng���˿���id, Val(.TextMatrix(lngRow, COL_��������ID)), 1, 1)
                End If
                
                blnReInRow = True '��Ҫˢ����ҩ�÷�ִ�п���
                blnCurDo = True
            End If
            
            '�÷��ͼ巨��ִ������
            If cboִ������.Tag <> "" Then
                '�÷�
                i = Nvl(GetItemField("������ĿĿ¼", Val(.TextMatrix(lngRow, COL_������ĿID)), "ִ�п���"), 0)
                .TextMatrix(lngRow, COL_ִ������) = Decode(NeedName(cboִ������.Text), "��Ժ��ҩ", 5, i)
                If Val(.TextMatrix(lngRow, COL_ִ������)) = 5 Then
                    .TextMatrix(lngRow, COL_ִ�п���ID) = 0
                Else
                    .TextMatrix(lngRow, COL_ִ�п���ID) = Get����ִ�п���ID(mlng����ID, 0, .TextMatrix(lngRow, COL_���), _
                        Val(.TextMatrix(lngRow, COL_������ĿID)), 0, Val(.TextMatrix(lngRow, COL_ִ������)), _
                        mlng���˿���id, Val(Val(.TextMatrix(lngRow, COL_��������ID))), 1, 1)
                End If
                
                '�巨
                i = Nvl(GetItemField("������ĿĿ¼", Val(.TextMatrix(lngTmp, COL_������ĿID)), "ִ�п���"), 0)
                .TextMatrix(lngTmp, COL_ִ������) = Decode(NeedName(cboִ������.Text), "��Ժ��ҩ", 5, i)
                If Val(.TextMatrix(lngTmp, COL_ִ������)) = 5 Then
                    .TextMatrix(lngTmp, COL_ִ�п���ID) = 0
                Else
                    .TextMatrix(lngTmp, COL_ִ�п���ID) = Get����ִ�п���ID(mlng����ID, 0, .TextMatrix(lngTmp, COL_���), _
                        Val(.TextMatrix(lngTmp, COL_������ĿID)), 0, Val(.TextMatrix(lngTmp, COL_ִ������)), _
                        mlng���˿���id, Val(.TextMatrix(lngTmp, COL_��������ID)), 1, 1)
                End If
                
                If InStr(",0,3,", .TextMatrix(lngTmp, COL_EDIT)) > 0 Then
                    .TextMatrix(lngTmp, COL_EDIT) = 2
                    .TextMatrix(lngTmp, COL_״̬) = 1 '�޸ĺ��Ϊ�¿�
                End If
                mblnNoSave = True '���Ϊδ����
                
                blnCurDo = True
            End If
            
            '��ҩ�÷�ִ�п���:���䷽��ǰ��ʾ�е�ִ�п���
            If cbo����ִ��.ListIndex <> -1 And cbo����ִ��.Tag <> "" Then
                .TextMatrix(lngRow, COL_ִ�п���ID) = cbo����ִ��.ItemData(cbo����ִ��.ListIndex)
                blnCurDo = True
            End If
            
            '---------------
            If blnCurDo Then '���Ϊ�޸�:0-ԭʼ��,1-������,2-�޸�������,3-�޸������
                If InStr(",0,3,", .TextMatrix(lngRow, COL_EDIT)) > 0 Then
                    .TextMatrix(lngRow, COL_EDIT) = 2
                    .TextMatrix(lngRow, COL_״̬) = 1 '�޸ĺ��Ϊ�¿�
                    Call ReSetColor(lngRow)
                End If
                mblnNoSave = True '���Ϊδ����
            End If
        Else '����������Ŀ
            If IsDate(txt��ʼʱ��.Text) And txt��ʼʱ��.Tag <> "" Then
                .TextMatrix(lngRow, COL_��ʼʱ��) = Format(txt��ʼʱ��.Text, "MM-dd HH:mm")
                .Cell(flexcpData, lngRow, COL_��ʼʱ��) = txt��ʼʱ��.Text
                blnCurDo = True
            End If
            If chk����.Visible And chk����.Tag <> "" Then
                .TextMatrix(lngRow, COL_��־) = chk����.Value
                
                '��ʾ������־,һ����ҩ��ʾ�ڵ�һ��
                lngBeginRow = lngRow
                If InStr(",5,6,", .TextMatrix(lngRow, COL_���)) > 0 Then
                    lngBeginRow = .FindRow(.TextMatrix(lngRow, COL_���ID), , COL_���ID)
                End If
                If Val(.TextMatrix(lngRow, COL_��־)) = 1 Then
                    Set .Cell(flexcpPicture, lngBeginRow, COL_F��־) = imgFlag.ListImages("����").Picture
                Else
                    Set .Cell(flexcpPicture, lngBeginRow, COL_F��־) = Nothing
                End If
                .Cell(flexcpPictureAlignment, lngBeginRow, COL_F��־) = 4
                
                blnCurDo = True
            End If
            If txt����.Enabled And (IsNumeric(txt����.Text) Or txt����.Text = "") And txt����.Tag <> "" Then
                .TextMatrix(lngRow, COL_����) = FormatEx(txt����.Text, 5)
                blnCurDo = True
            End If
            
            If txt����.Visible And txt����.Enabled And txt����.Tag <> "" Then
                .TextMatrix(lngRow, COL_����) = txt����.Text
                blnCurDo = True
            End If
            
            If txt����.Enabled And IsNumeric(txt����.Text) And txt����.Tag <> "" Then
                .TextMatrix(lngRow, COL_����) = FormatEx(Val(txt����.Text), 5)
                blnCurDo = True
            End If
            
            If txtƵ��.Enabled And cmdƵ��.Tag <> "" And txtƵ��.Tag <> "" Then
                .TextMatrix(lngRow, COL_Ƶ��) = txtƵ��.Text
                Call GetƵ����Ϣ_����(txtƵ��.Text, intƵ�ʴ���, intƵ�ʼ��, str�����λ, GetƵ�ʷ�Χ(lngRow))
                .TextMatrix(lngRow, COL_Ƶ�ʴ���) = intƵ�ʴ���
                .TextMatrix(lngRow, COL_Ƶ�ʼ��) = intƵ�ʼ��
                .TextMatrix(lngRow, COL_�����λ) = str�����λ
                blnCurDo = True
            End If
            
            If cboִ��ʱ��.Tag <> "" Then
                .TextMatrix(lngRow, COL_ִ��ʱ��) = cboִ��ʱ��.Text
                blnCurDo = True
            End If
            If cboҽ������.Tag <> "" Then
                .TextMatrix(lngRow, COL_ҽ������) = cboҽ������.Text
                blnCurDo = True
            End If
            
            If cboִ�п���.ListIndex <> -1 And cboִ�п���.Tag <> "" Then
                If Not RowIn������(lngRow) Then '�ɼ�������ִ�п��Ҳ�ͬ
                    .TextMatrix(lngRow, COL_ִ�п���ID) = cboִ�п���.ItemData(cboִ�п���.ListIndex)
                End If
                blnCurDo = True
            End If
            
            '����ִ�п��ң���ҩ;��,��������,�ɼ�����
            If cbo����ִ��.ListIndex <> -1 And cbo����ִ��.Tag <> "" Then
                lngTmp = -1
                If InStr(",5,6,", .TextMatrix(lngRow, COL_���)) > 0 Then
                    lngTmp = .FindRow(CLng(.TextMatrix(lngRow, COL_���ID)), lngRow + 1)
                ElseIf .TextMatrix(lngRow, COL_���) = "F" Then
                    For i = lngRow + 1 To .Rows - 1
                        If Val(.TextMatrix(i, COL_���ID)) = .RowData(lngRow) Then
                            If .TextMatrix(i, COL_���) = "G" Then
                                lngTmp = i: Exit For
                            End If
                        Else
                            Exit For
                        End If
                    Next
                ElseIf .TextMatrix(lngRow, COL_���) = "E" _
                    And .TextMatrix(lngRow - 1, COL_���) = "C" _
                    And Val(.TextMatrix(lngRow - 1, COL_���ID)) = .RowData(lngRow) Then
                    lngTmp = lngRow
                End If
                
                'ֻ���¶�Ӧ��,��Ӱ��������
                If lngTmp <> -1 Then
                    .TextMatrix(lngTmp, COL_ִ�п���ID) = cbo����ִ��.ItemData(cbo����ִ��.ListIndex)
                    If InStr(",0,3,", .TextMatrix(lngTmp, COL_EDIT)) > 0 Then
                        .TextMatrix(lngTmp, COL_EDIT) = 2
                        .TextMatrix(lngTmp, COL_״̬) = 1 '�޸ĺ��Ϊ�¿�
                    End If
                    mblnNoSave = True '���Ϊδ����
                End If
            End If
            
            'ִ������,��ҩ;��:Ϊ���¿���ʱ��(������ҩ;����ͬ������),���ж��Ƿ�ı�
            If InStr(",5,6,", .TextMatrix(lngRow, COL_���)) > 0 Then
                If cboִ������.Tag <> "" Then blnCurDo = True
                If Val(cmd�÷�.Tag) <> 0 And txt�÷�.Tag <> "" Then blnCurDo = True
            End If
                                    
            '�޸�ʱ�Զ����²�������
            blnTmp = False
            If cboִ������.Tag <> "" Or (Val(cmd�÷�.Tag) <> 0 And txt�÷�.Tag <> "") Then
                blnReInRow = True '��Ҫˢ�¸�ҩ;��,�ɼ���ʽ��ִ�п���
                blnTmp = True
            End If
            If blnCurDo Or blnTmp Then
                '�޸�����������¿���ʱ��
                strCurDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
                .TextMatrix(lngRow, COL_����ʱ��) = Format(strCurDate, "MM-dd HH:mm")
                .Cell(flexcpData, lngRow, COL_����ʱ��) = strCurDate
                
                '��鿪��ҽ��
                If .TextMatrix(lngRow, COL_����ҽ��) <> UserInfo.���� Then
                    .TextMatrix(lngRow, COL_����ҽ��) = UserInfo.����
                    If lng��������ID = 0 Then
                        lng��������ID = Get��������ID(UserInfo.ID, mlng���˿���id, 1)
                    End If
                    .TextMatrix(lngRow, COL_��������ID) = lng��������ID
                End If
            End If
                                    
            '������Ҫͬ������Ĺ�����
            '----------------------------------------------------------------
            If RowIn������(lngRow) Then
                '�ɼ�����
                If Val(cmd�÷�.Tag) <> 0 And txt�÷�.Tag <> "" Then
                    .TextMatrix(lngRow, COL_������ĿID) = Val(cmd�÷�.Tag)
                    .TextMatrix(lngRow, COL_�÷�) = txt�÷�.Text
                    .TextMatrix(lngRow, COL_����) = txt�÷�.Text
                    
                    'ͬʱ���ļƼ����ʺ�ִ������
                    .TextMatrix(lngRow, COL_�Ƽ�����) = Nvl(GetItemField("������ĿĿ¼", Val(cmd�÷�.Tag), "�Ƽ�����"), 0)
                    .TextMatrix(lngRow, COL_ִ������) = Nvl(GetItemField("������ĿĿ¼", Val(cmd�÷�.Tag), "ִ�п���"), 0)
                    If InStr(",0,5,", Val(.TextMatrix(lngRow, COL_ִ������))) = 0 Then
                        .TextMatrix(lngRow, COL_ִ�п���ID) = Get����ִ�п���ID(mlng����ID, 0, "E", Val(cmd�÷�.Tag), 0, _
                            Val(.TextMatrix(lngRow, COL_ִ������)), mlng���˿���id, Val(.TextMatrix(lngRow, COL_��������ID)), 1, 1)
                    Else
                        .TextMatrix(lngRow, COL_ִ�п���ID) = 0
                    End If

                    blnCurDo = True
                End If
                
                '����һ���ɼ��ĸ���������Ŀ
                If blnCurDo Then
                    For i = lngRow - 1 To .FixedRows Step -1
                        If Val(.TextMatrix(i, COL_���ID)) = .RowData(lngRow) Then
                            If txt����.Tag <> "" Then
                                .TextMatrix(i, COL_����) = .TextMatrix(lngRow, COL_����)
                                blnOtherDo = True
                            End If
                            If txtƵ��.Tag <> "" Then
                                .TextMatrix(i, COL_Ƶ��) = .TextMatrix(lngRow, COL_Ƶ��)
                                .TextMatrix(i, COL_Ƶ�ʴ���) = .TextMatrix(lngRow, COL_Ƶ�ʴ���)
                                .TextMatrix(i, COL_Ƶ�ʼ��) = .TextMatrix(lngRow, COL_Ƶ�ʼ��)
                                .TextMatrix(i, COL_�����λ) = .TextMatrix(lngRow, COL_�����λ)
                                blnOtherDo = True
                            End If
                            If cboִ�п���.Tag <> "" And cboִ�п���.ListIndex <> -1 Then
                                If InStr(",0,5,", Val(.TextMatrix(i, COL_ִ������))) > 0 Then
                                    .TextMatrix(i, COL_ִ�п���ID) = 0
                                Else
                                    .TextMatrix(i, COL_ִ�п���ID) = cboִ�п���.ItemData(cboִ�п���.ListIndex)
                                End If
                                blnOtherDo = True
                            End If
                            If cboִ��ʱ��.Tag <> "" Then
                                .TextMatrix(i, COL_ִ��ʱ��) = .TextMatrix(lngRow, COL_ִ��ʱ��)
                                blnOtherDo = True
                            End If
                            If txt��ʼʱ��.Tag <> "" Then
                                .TextMatrix(i, COL_��ʼʱ��) = .TextMatrix(lngRow, COL_��ʼʱ��)
                                .Cell(flexcpData, i, COL_��ʼʱ��) = .Cell(flexcpData, lngRow, COL_��ʼʱ��)
                                blnOtherDo = True
                            End If
                            If chk����.Tag <> "" Then
                                .TextMatrix(i, COL_��־) = .TextMatrix(lngRow, COL_��־)
                                blnOtherDo = True
                            End If
                                            
                            '����ʱ��
                            If .TextMatrix(i, COL_����ʱ��) <> .TextMatrix(lngRow, COL_����ʱ��) Then
                                .TextMatrix(i, COL_����ʱ��) = .TextMatrix(lngRow, COL_����ʱ��)
                                .Cell(flexcpData, i, COL_����ʱ��) = .Cell(flexcpData, lngRow, COL_����ʱ��)
                                blnOtherDo = True
                            End If
                            
                            '����ҽ��
                            If .TextMatrix(i, COL_����ҽ��) <> .TextMatrix(lngRow, COL_����ҽ��) Then
                                .TextMatrix(i, COL_����ҽ��) = .TextMatrix(lngRow, COL_����ҽ��)
                                blnOtherDo = True
                            End If
                                            
                            '��������ID
                            If .TextMatrix(i, COL_��������ID) <> .TextMatrix(lngRow, COL_��������ID) Then
                                .TextMatrix(i, COL_��������ID) = .TextMatrix(lngRow, COL_��������ID)
                                blnOtherDo = True
                            End If
                            
                            '���Ϊ�޸�
                            If blnOtherDo And InStr(",0,3,", .TextMatrix(i, COL_EDIT)) > 0 Then
                                .TextMatrix(i, COL_EDIT) = 2
                                .TextMatrix(i, COL_״̬) = 1 '�޸ĺ��Ϊ�¿�
                            End If
                        Else
                            Exit For
                        End If
                    Next
                End If
            ElseIf InStr(",5,6,", .TextMatrix(lngRow, COL_���)) > 0 Then
                '�С�����ҩ�����ҩ;����һ����ҩ�����
                
                'ִ������
                If cboִ������.Tag <> "" Then
                    .TextMatrix(lngRow, COL_ִ������) = Decode(NeedName(cboִ������.Text), "�Ա�ҩ", 5, 4)
                    If Val(.TextMatrix(lngRow, COL_ִ������)) = 5 Then
                        .TextMatrix(lngRow, COL_ִ�п���ID) = 0
                    ElseIf Val(.TextMatrix(lngRow, COL_ִ�п���ID)) = 0 Then
                        '�ָ�ȱʡҩ��,ȱʡ��ǰ��ĳ�ҩ��ͬ
                        strTmp = Get����ҩ��IDs(.TextMatrix(lngRow, COL_���), Val(.TextMatrix(lngRow, COL_������ĿID)), Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)), mlng���˿���id)
                        For i = lngRow - 1 To .FixedRows Step -1
                            '����ҩ���г�ҩ��ҩ�����ܲ�ͬ,�������Ҫ��ͬ
                            If .TextMatrix(i, COL_���) = .TextMatrix(lngRow, COL_���) And Val(.TextMatrix(i, COL_ִ�п���ID)) <> 0 Then
                                If InStr("," & strTmp & ",", "," & Val(.TextMatrix(i, COL_ִ�п���ID)) & ",") > 0 Then
                                    .TextMatrix(lngRow, COL_ִ�п���ID) = Val(.TextMatrix(i, COL_ִ�п���ID))
                                    Exit For
                                End If
                            End If
                        Next
                        If Val(.TextMatrix(lngRow, COL_ִ�п���ID)) = 0 Then
                            .TextMatrix(lngRow, COL_ִ�п���ID) = Get����ִ�п���ID(mlng����ID, 0, .TextMatrix(lngRow, COL_���), Val(.TextMatrix(lngRow, COL_������ĿID)), Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)), 4, mlng���˿���id, 0, 1, 1, True)
                        End If
                    End If
                    cboִ�п���.Tag = "1" '����ִ�п���һ����ҩ��Ҫͬ����
                    blnReInRow = True '����ִ�п��ұ༭�Ա仯
                End If
                
                '��ҩ;�����������������ͬ������
                If Val(cmd�÷�.Tag) <> 0 And txt�÷�.Tag <> "" Then
                    .TextMatrix(lngRow, COL_�÷�) = txt�÷�.Text
                    Call AdviceSet��ҩ;��(lngRow, Val(cmd�÷�.Tag), NeedName(cboִ������.Text))
                ElseIf blnCurDo Then 'cboִ������.Tag <> "" Then
                    '���ִ�����ʸ�����,��Ҫǿ���޸Ķ�Ӧ�ĸ�ҩ;����ִ�����ʺ�ִ�п���
                    lngTmp = .FindRow(CLng(.TextMatrix(lngRow, COL_���ID)), lngRow + 1)
                    Call AdviceSet��ҩ;��(lngRow, Val(.TextMatrix(lngTmp, COL_������ĿID)), NeedName(cboִ������.Text))
                End If
                
                'һ����ҩ:�������ҩ;��,ǰ���ѵ�������
                If blnCurDo Then
                    lngBeginRow = .FindRow(.TextMatrix(lngRow, COL_���ID), , COL_���ID)
                    For i = lngBeginRow To .Rows - 1
                        If i <> lngRow And .RowData(i) <> 0 Then '���������м��п���
                            If Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                                If txt��ʼʱ��.Tag <> "" Then
                                    .TextMatrix(i, COL_��ʼʱ��) = .TextMatrix(lngRow, COL_��ʼʱ��)
                                    .Cell(flexcpData, i, COL_��ʼʱ��) = .Cell(flexcpData, lngRow, COL_��ʼʱ��)
                                    blnOtherDo = True
                                End If
                                If txt�÷�.Tag <> "" Then
                                    .TextMatrix(i, COL_�÷�) = .TextMatrix(lngRow, COL_�÷�)
                                    blnOtherDo = True
                                End If
                                If txtƵ��.Tag <> "" Then
                                    .TextMatrix(i, COL_Ƶ��) = .TextMatrix(lngRow, COL_Ƶ��)
                                    .TextMatrix(i, COL_Ƶ�ʴ���) = .TextMatrix(lngRow, COL_Ƶ�ʴ���)
                                    .TextMatrix(i, COL_Ƶ�ʼ��) = .TextMatrix(lngRow, COL_Ƶ�ʼ��)
                                    .TextMatrix(i, COL_�����λ) = .TextMatrix(lngRow, COL_�����λ)
                                    blnOtherDo = True
                                End If
                                
                                'һ����ҩ��,������ͬ�仯,�������¼���
                                If txt����.Tag <> "" Then
                                    .TextMatrix(i, COL_����) = .TextMatrix(lngRow, COL_����)
                                    If .TextMatrix(i, COL_Ƶ��) <> "" _
                                        And Val(.TextMatrix(i, COL_����)) <> 0 _
                                        And Val(.TextMatrix(i, COL_����ϵ��)) <> 0 _
                                        And Val(.TextMatrix(i, COL_�����װ)) <> 0 Then
                                        
                                        .TextMatrix(i, COL_����) = FormatEx(CalcȱʡҩƷ����( _
                                            Val(.TextMatrix(i, COL_����)), Val(.TextMatrix(i, COL_����)), _
                                            Val(.TextMatrix(i, COL_Ƶ�ʴ���)), Val(.TextMatrix(i, COL_Ƶ�ʼ��)), _
                                            .TextMatrix(i, COL_�����λ), .TextMatrix(i, COL_ִ��ʱ��), _
                                            Val(.TextMatrix(i, COL_����ϵ��)), Val(.TextMatrix(i, COL_�����װ)), _
                                            Val(.TextMatrix(i, COL_�ɷ����))), 5)
                                    End If
                                    blnOtherDo = True
                                End If
                                
                                If cboִ��ʱ��.Tag <> "" Then
                                    .TextMatrix(i, COL_ִ��ʱ��) = .TextMatrix(lngRow, COL_ִ��ʱ��)
                                    blnOtherDo = True
                                End If
                                
                                'ִ������:��Ժ��ҩ��һ����ҩ����һ�£������ɵ�������
                                If cboִ������.Tag <> "" And NeedName(cboִ������.Text) = "��Ժ��ҩ" Then
                                    .TextMatrix(i, COL_ִ������) = .TextMatrix(lngRow, COL_ִ������)
                                    '���Ա�ҩת����ʱ��Ҫ��������ִ�п���
                                    If Val(.TextMatrix(i, COL_ִ�п���ID)) = 0 Then
                                        .TextMatrix(i, COL_ִ�п���ID) = .TextMatrix(lngRow, COL_ִ�п���ID)
                                    End If
                                    blnOtherDo = True
                                End If
                                
                                'ִ�п���:ִ�п���(ҩ��)���Բ�ͬ
'                                If cboִ�п���.Tag <> "" Then
'                                    .TextMatrix(i, COL_ִ�п���ID) = .TextMatrix(lngRow, COL_ִ�п���ID)
'                                    blnOtherDo = True
'                                End If
                                
                                '������־
                                If chk����.Tag <> "" Then
                                    .TextMatrix(i, COL_��־) = .TextMatrix(lngRow, COL_��־)
                                    blnOtherDo = True
                                End If
                                
                                '����ʱ��
                                If .TextMatrix(i, COL_����ʱ��) <> .TextMatrix(lngRow, COL_����ʱ��) Then
                                    .TextMatrix(i, COL_����ʱ��) = .TextMatrix(lngRow, COL_����ʱ��)
                                    .Cell(flexcpData, i, COL_����ʱ��) = .Cell(flexcpData, lngRow, COL_����ʱ��)
                                    blnOtherDo = True
                                End If
                                
                                '����ҽ��
                                If .TextMatrix(i, COL_����ҽ��) <> .TextMatrix(lngRow, COL_����ҽ��) Then
                                    .TextMatrix(i, COL_����ҽ��) = .TextMatrix(lngRow, COL_����ҽ��)
                                    blnOtherDo = True
                                End If
                                
                                '��������ID
                                If .TextMatrix(i, COL_��������ID) <> .TextMatrix(lngRow, COL_��������ID) Then
                                    .TextMatrix(i, COL_��������ID) = .TextMatrix(lngRow, COL_��������ID)
                                    blnOtherDo = True
                                End If
                                
                                '���Ϊ�޸�
                                If blnOtherDo And InStr(",0,3,", .TextMatrix(i, COL_EDIT)) > 0 Then
                                    .TextMatrix(i, COL_EDIT) = 2
                                    .TextMatrix(i, COL_״̬) = 1 '�޸ĺ��Ϊ�¿�
                                End If
                            Else
                                Exit For
                            End If
                        End If
                    Next
                End If
            ElseIf InStr(",D,F,", .TextMatrix(lngRow, COL_���)) > 0 And blnCurDo Then
                '��������Ŀ�л�����������
                lngBeginRow = .FindRow(CStr(.RowData(lngRow)), lngRow + 1, COL_���ID)
                If lngBeginRow <> -1 Then
                    For i = lngBeginRow To .Rows - 1
                        If Val(.TextMatrix(i, COL_���ID)) = .RowData(lngRow) Then
                            If txt����.Tag <> "" Then
                                .TextMatrix(i, COL_����) = .TextMatrix(lngRow, COL_����)
                                blnOtherDo = True
                            End If
                            If txt����.Tag <> "" Then
                                .TextMatrix(i, COL_����) = .TextMatrix(lngRow, COL_����)
                                blnOtherDo = True
                            End If
                            
                            If cboִ��ʱ��.Tag <> "" Then
                                .TextMatrix(i, COL_ִ��ʱ��) = .TextMatrix(lngRow, COL_ִ��ʱ��)
                                blnOtherDo = True
                            End If
                            If txtƵ��.Tag <> "" Then
                                .TextMatrix(i, COL_Ƶ��) = .TextMatrix(lngRow, COL_Ƶ��)
                                .TextMatrix(i, COL_Ƶ�ʴ���) = .TextMatrix(lngRow, COL_Ƶ�ʴ���)
                                .TextMatrix(i, COL_Ƶ�ʼ��) = .TextMatrix(lngRow, COL_Ƶ�ʼ��)
                                .TextMatrix(i, COL_�����λ) = .TextMatrix(lngRow, COL_�����λ)
                                blnOtherDo = True
                            End If
                            If cboִ�п���.Tag <> "" Then
                                If InStr(",0,5,", Val(.TextMatrix(i, COL_ִ������))) > 0 Then
                                    .TextMatrix(i, COL_ִ�п���ID) = 0
                                ElseIf .TextMatrix(i, COL_���) <> "G" Then '���������ִ�п���Ϊ����
                                    .TextMatrix(i, COL_ִ�п���ID) = .TextMatrix(lngRow, COL_ִ�п���ID)
                                End If
                                blnOtherDo = True
                            End If
                            If txt��ʼʱ��.Tag <> "" Then
                                .TextMatrix(i, COL_��ʼʱ��) = .TextMatrix(lngRow, COL_��ʼʱ��)
                                .Cell(flexcpData, i, COL_��ʼʱ��) = .Cell(flexcpData, lngRow, COL_��ʼʱ��)
                                blnOtherDo = True
                            End If
                            If chk����.Tag <> "" Then
                                .TextMatrix(i, COL_��־) = .TextMatrix(lngRow, COL_��־)
                                blnOtherDo = True
                            End If
                            
                            '����ʱ��
                            If .TextMatrix(i, COL_����ʱ��) <> .TextMatrix(lngRow, COL_����ʱ��) Then
                                .TextMatrix(i, COL_����ʱ��) = .TextMatrix(lngRow, COL_����ʱ��)
                                .Cell(flexcpData, i, COL_����ʱ��) = .Cell(flexcpData, lngRow, COL_����ʱ��)
                                blnOtherDo = True
                            End If
                            
                            '����ҽ��
                            If .TextMatrix(i, COL_����ҽ��) <> .TextMatrix(lngRow, COL_����ҽ��) Then
                                .TextMatrix(i, COL_����ҽ��) = .TextMatrix(lngRow, COL_����ҽ��)
                                blnOtherDo = True
                            End If
                            
                            '��������ID
                            If .TextMatrix(i, COL_��������ID) <> .TextMatrix(lngRow, COL_��������ID) Then
                                .TextMatrix(i, COL_��������ID) = .TextMatrix(lngRow, COL_��������ID)
                                blnOtherDo = True
                            End If
                            
                            '���Ϊ�޸�
                            If blnOtherDo And InStr(",0,3,", .TextMatrix(i, COL_EDIT)) > 0 Then
                                .TextMatrix(i, COL_EDIT) = 2
                                .TextMatrix(i, COL_״̬) = 1 '�޸ĺ��Ϊ�¿�
                            End If
                        Else
                            Exit For
                        End If
                    Next
                End If
            End If
                     
            If blnCurDo Then '���Ϊ�޸�:0-ԭʼ��,1-������,2-�޸�������,3-�޸������
                If InStr(",0,3,", .TextMatrix(lngRow, COL_EDIT)) > 0 Then
                    .TextMatrix(lngRow, COL_EDIT) = 2
                    .TextMatrix(lngRow, COL_״̬) = 1 '�޸ĺ��Ϊ�¿�
                    Call ReSetColor(lngRow)
                End If
                mblnNoSave = True '���Ϊδ����
            End If
        End If
        
        '����ҽ������
        If AdviceTextChange(lngRow) Then
            .TextMatrix(lngRow, COL_ҽ������) = AdviceTextMake(lngRow)
            txtҽ������.Text = .TextMatrix(lngRow, COL_ҽ������)
        End If
    End With
        
    '����༭��־
    Call ClearItemTag
    
    'ĳЩ�������Ҫ�������ÿ�Ƭ����Ŀ�༭��(���޸���ִ������ʱ)
    If blnReInRow Then
        Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
    End If
End Sub

Private Sub ReSetColor(ByVal lngRow As Long)
'���ܣ���������ָ���е���ɫ
'˵������Ϊ���ʵ�ҽ���༭�����¿�
    Dim lngBegin As Long, lngEnd As Long, i As Long
    
    With vsAdvice
        'һ����ҩ��Χ
        lngBegin = lngRow: lngEnd = lngRow
        If InStr(",5,6,", .TextMatrix(lngRow, COL_���)) > 0 Then
            If RowInһ����ҩ(lngRow) Then
                Call Getһ����ҩ��Χ(Val(.TextMatrix(lngRow, COL_���ID)), lngBegin, lngEnd)
            End If
        End If
        '�ָ�������ɫ
        For i = lngBegin To lngEnd
            .Cell(flexcpForeColor, i, .FixedCols, i, COL_����ҽ��) = .ForeColor
            '���龫����ɫ��ʶ
            If InStr(",����ҩ,����ҩ,����ҩ,", .TextMatrix(i, COL_�������)) > 0 _
                And .TextMatrix(i, COL_�������) <> "" Then
                .Cell(flexcpFontBold, i, COL_ҽ������) = True
            End If
        Next
        .ForeColorSel = .Cell(flexcpForeColor, lngRow, COL_��ʼʱ��)
    End With
End Sub

Private Sub AdviceSetһ����ҩ(ByVal lngBegin As Long, ByVal lngEnd As Long)
'���ܣ���ѡ��Χ�ڵ�ҩƷ����Ϊһ����ҩ
'��������ֹ�к�,�м䲻��������,���������һ��ҩƷ�ĸ�ҩ;����
'˵�����Ե�һ��ҩƷ�ĸ�ҩ;��Ϊ׼,��λ�÷������һ��ҩƷ֮��
    Dim varTmp1 As Variant, varTmp2 As Variant
    Dim lngRow1 As Long, lngRow2 As Long
    Dim lng���ID As Long, i As Long
    Dim strStart As String, curDate As Date
    
    lngRow1 = vsAdvice.FindRow(CLng(vsAdvice.TextMatrix(lngBegin, COL_���ID)), lngBegin + 1) '��һ��ҩ;����
    lngRow2 = vsAdvice.FindRow(CLng(vsAdvice.TextMatrix(lngEnd, COL_���ID)), lngEnd + 1) '����ҩ;����
    
    'ɾ����ҩ;����֮ǰ��¼ִ������,�Ա�������ж�
    For i = lngRow2 To lngRow1 Step -1
        If Val(vsAdvice.TextMatrix(i, COL_Ӥ��)) = cboӤ��.ListIndex And vsAdvice.RowHidden(i) Then
            vsAdvice.Cell(flexcpData, i - 1, COL_ִ������) = Val(vsAdvice.TextMatrix(i, COL_ִ������))
        End If
    Next
    
    '���Ƶ�һ�еĸ�ҩ;�������һ�еĸ�ҩ;��
    For i = vsAdvice.FixedCols To vsAdvice.Cols - 1
        If i <> COL_EDIT And i <> COL_���ID And i <> COL_��� And i <> COL_״̬ Then
            vsAdvice.TextMatrix(lngRow2, i) = vsAdvice.TextMatrix(lngRow1, i)
        End If
    Next
    '�༭��־��0-ԭʼ��,1-������,2-�޸�������,3-�޸������
    If InStr(",0,3,", vsAdvice.TextMatrix(lngRow2, COL_EDIT)) > 0 Then
        vsAdvice.TextMatrix(lngRow2, COL_EDIT) = 2 '���Ϊ���޸�
        vsAdvice.TextMatrix(lngRow2, COL_״̬) = 1 '�޸ĺ��Ϊ�¿�
    End If
    lng���ID = vsAdvice.RowData(lngRow2)
    
    varTmp1 = mblnRowChange: varTmp2 = vsAdvice.Redraw
    mblnRowChange = False: vsAdvice.Redraw = flexRDNone
    
    'ɾ�������һ�и�ҩ;�����������ҩ;��
    For i = lngEnd To lngBegin Step -1
        If Val(vsAdvice.TextMatrix(i, COL_Ӥ��)) = cboӤ��.ListIndex Then
            If vsAdvice.RowHidden(i) Then
                Call DeleteRow(i)
            Else
                vsAdvice.TextMatrix(i, COL_���ID) = lng���ID
                If InStr(",0,3,", vsAdvice.TextMatrix(i, COL_EDIT)) > 0 Then
                    vsAdvice.TextMatrix(i, COL_EDIT) = 2 '���Ϊ���޸�
                    vsAdvice.TextMatrix(i, COL_״̬) = 1 '�޸ĺ��Ϊ�¿�
                End If
            End If
        End If
    Next
    
    '�к��ѱ��
    lngRow1 = lngBegin '��ʼһ����ҩ��
    curDate = zlDatabase.Currentdate
    
    '���ҽ���Ƿ���
    If vsAdvice.TextMatrix(lngRow1, COL_����ҽ��) <> UserInfo.���� Then
        '���������Ϣ:ǰ���ѱ��Ϊ�޸�,���ֹ��������ʱ���н������ˢ��
        vsAdvice.TextMatrix(lngRow1, COL_����ҽ��) = UserInfo.����
        vsAdvice.TextMatrix(lngRow1, COL_��������ID) = Get��������ID(UserInfo.ID, mlng���˿���id, 1)
        
        vsAdvice.TextMatrix(lngRow1, COL_����ʱ��) = Format(curDate, "MM-dd HH:mm")
        vsAdvice.Cell(flexcpData, lngRow1, COL_����ʱ��) = Format(curDate, "yyyy-MM-dd HH:mm")
    End If
    
    For i = lngRow1 + 1 To vsAdvice.Rows - 1
        If Val(vsAdvice.TextMatrix(i, COL_���ID)) = lng���ID Then
            lngRow2 = i '��¼�µĽ����к�
            
            'һ����ҩ�Ĳ�����Ϣ��ͬ
            vsAdvice.TextMatrix(i, COL_��ʼʱ��) = vsAdvice.TextMatrix(lngRow1, COL_��ʼʱ��)
            vsAdvice.Cell(flexcpData, i, COL_��ʼʱ��) = vsAdvice.Cell(flexcpData, lngRow1, COL_��ʼʱ��)
            
            vsAdvice.TextMatrix(i, COL_����ҽ��) = vsAdvice.TextMatrix(lngRow1, COL_����ҽ��)
            vsAdvice.TextMatrix(i, COL_��������ID) = vsAdvice.TextMatrix(lngRow1, COL_��������ID)
            
            vsAdvice.TextMatrix(i, COL_����ʱ��) = vsAdvice.TextMatrix(lngRow1, COL_����ʱ��) 'һ����ҩ�Ŀ���ʱ����ͬ
            vsAdvice.Cell(flexcpData, i, COL_����ʱ��) = vsAdvice.Cell(flexcpData, lngRow1, COL_����ʱ��)
            
            vsAdvice.TextMatrix(i, COL_�÷�) = vsAdvice.TextMatrix(lngRow1, COL_�÷�)
            
            vsAdvice.TextMatrix(i, COL_Ƶ��) = vsAdvice.TextMatrix(lngRow1, COL_Ƶ��)
            vsAdvice.TextMatrix(i, COL_Ƶ�ʴ���) = vsAdvice.TextMatrix(lngRow1, COL_Ƶ�ʴ���)
            vsAdvice.TextMatrix(i, COL_Ƶ�ʼ��) = vsAdvice.TextMatrix(lngRow1, COL_Ƶ�ʼ��)
            vsAdvice.TextMatrix(i, COL_�����λ) = vsAdvice.TextMatrix(lngRow1, COL_�����λ)
            vsAdvice.TextMatrix(i, COL_ִ��ʱ��) = vsAdvice.TextMatrix(lngRow1, COL_ִ��ʱ��)
            
            vsAdvice.TextMatrix(i, COL_��־) = vsAdvice.TextMatrix(lngRow1, COL_��־)
            Set vsAdvice.Cell(flexcpPicture, i, COL_F��־) = Nothing '�ڿ�ʼ����ʾ
            
            If Val(vsAdvice.TextMatrix(lngRow1, COL_ִ������)) <> 5 And Val(vsAdvice.Cell(flexcpData, lngRow1, COL_ִ������)) = 5 Then
                '��һ������Ժ��ҩ,ȫ������Ϊ��Ժ��ҩ
                vsAdvice.TextMatrix(i, COL_ִ������) = vsAdvice.TextMatrix(lngRow1, COL_ִ������)
                vsAdvice.TextMatrix(i, COL_ִ�п���ID) = vsAdvice.TextMatrix(lngRow1, COL_ִ�п���ID)
            ElseIf Val(vsAdvice.TextMatrix(i, COL_ִ������)) <> 5 And Val(vsAdvice.Cell(flexcpData, i, COL_ִ������)) = 5 Then
                '��ǰ������Ժ��ҩ,������Ϊ���һ����ͬ
                vsAdvice.TextMatrix(i, COL_ִ������) = vsAdvice.TextMatrix(lngRow1, COL_ִ������)
                vsAdvice.TextMatrix(i, COL_ִ�п���ID) = vsAdvice.TextMatrix(lngRow1, COL_ִ�п���ID)
            Else
                '���򱣳ֲ���
            End If
            
'            'ִ������:һ��������ͬ,��ȱʡ����һ������
'            vsAdvice.TextMatrix(i, COL_ִ������) = vsAdvice.TextMatrix(lngRow1, COL_ִ������)
'            'ִ�п���:ִ�п���(ҩ��)���Բ�ͬ
'            vsAdvice.TextMatrix(i, COL_ִ�п���ID) = vsAdvice.TextMatrix(lngRow1, COL_ִ�п���ID)
            
            '���Ϊ�޸�:0-ԭʼ��,1-������,2-�޸�������,3-�޸������
            If InStr(",0,3,", vsAdvice.TextMatrix(i, COL_EDIT)) > 0 Then
                vsAdvice.TextMatrix(i, COL_EDIT) = 2
                vsAdvice.TextMatrix(i, COL_״̬) = 1 '�޸ĺ��Ϊ�¿�
            End If
        Else
            Exit For
        End If
    Next
    
    '��ʼִ��ʱ�䴦��(�¿��Ĳ���̫��)
    strStart = ""
    For i = lngRow1 To lngRow2
        If Val(vsAdvice.TextMatrix(i, COL_EDIT)) = 1 Then
            If DateDiff("n", CDate(vsAdvice.Cell(flexcpData, i, COL_��ʼʱ��)), curDate) > 30 Then
                strStart = GetDefaultTime(i): Exit For
            End If
        End If
    Next
    If strStart <> "" Then
        For i = lngRow1 To lngRow2 + 1
            vsAdvice.Cell(flexcpData, i, COL_��ʼʱ��) = strStart
            vsAdvice.TextMatrix(i, COL_��ʼʱ��) = Format(strStart, "MM-dd HH:mm")
        Next
    End If

    mblnRowChange = varTmp1: vsAdvice.Redraw = varTmp2
    mblnNoSave = True '���Ϊδ����
End Sub

Private Sub AdviceSet������ҩ(ByVal lngBegin As Long, ByVal lngEnd As Long)
'���ܣ�ȡ��һ��ҩƷ��һ����ҩ
'��������ֹ�к�,�м䲻��������,���������һ��ҩƷ�ĸ�ҩ;����
    Dim varTmp1 As Variant, varTmp2 As Variant
    Dim lng��ҩ;��ID As Long, i As Long
    Dim intִ������ As Integer, strִ������ As String
    Dim lngRow As Long, curDate As Date, blnUpdate As Boolean
    
    With vsAdvice
        varTmp1 = mblnRowChange: varTmp2 = .Redraw
        mblnRowChange = False: .Redraw = flexRDNone
        
        'һ����ҩ;��
        lngRow = .FindRow(CLng(.TextMatrix(lngEnd, COL_���ID)), lngEnd + 1)
        lng��ҩ;��ID = Val(.TextMatrix(lngRow, COL_������ĿID))
        intִ������ = Val(.TextMatrix(lngRow, COL_ִ������))
                
        '���ҽ�����:�Ը�ҩ;����Ϊ׼�仯
        If .TextMatrix(lngRow, COL_����ҽ��) <> UserInfo.���� Then
            '���������Ϣ:�ֹ��������ʱ�н������ˢ��
            .TextMatrix(lngRow, COL_����ҽ��) = UserInfo.����
            .TextMatrix(lngRow, COL_��������ID) = Get��������ID(UserInfo.ID, mlng���˿���id, 1)
            curDate = zlDatabase.Currentdate
            .TextMatrix(lngRow, COL_����ʱ��) = Format(curDate, "MM-dd HH:mm")
            .Cell(flexcpData, lngRow, COL_����ʱ��) = Format(curDate, "yyyy-MM-dd HH:mm")
            
            If InStr(",0,3,", .TextMatrix(lngRow, COL_EDIT)) > 0 Then
                .TextMatrix(lngRow, COL_EDIT) = 2 '���Ϊ���޸�
                .TextMatrix(lngRow, COL_״̬) = 1 '�޸ĺ��Ϊ�¿�
            End If
            blnUpdate = True
        End If
                
        '��ʾ������־:ÿһ��
        For i = lngBegin To lngEnd
            If Val(.TextMatrix(i, COL_��־)) = 1 Then
                Set .Cell(flexcpPicture, i, COL_F��־) = imgFlag.ListImages("����").Picture
            Else
                Set .Cell(flexcpPicture, i, COL_F��־) = Nothing
            End If
            .Cell(flexcpPictureAlignment, i, COL_F��־) = 4
            
            'ҩƷ����Ӧ�仯
            If blnUpdate Then
                .TextMatrix(i, COL_����ҽ��) = .TextMatrix(lngRow, COL_����ҽ��)
                .TextMatrix(i, COL_��������ID) = .TextMatrix(lngRow, COL_��������ID)
                .TextMatrix(i, COL_����ʱ��) = .TextMatrix(lngRow, COL_����ʱ��)
                .Cell(flexcpData, i, COL_����ʱ��) = .Cell(flexcpData, lngRow, COL_����ʱ��)
                If InStr(",0,3,", .TextMatrix(i, COL_EDIT)) > 0 Then
                    .TextMatrix(i, COL_EDIT) = 2 '���Ϊ���޸�
                    .TextMatrix(i, COL_״̬) = 1 '�޸ĺ��Ϊ�¿�
                End If
            End If
        Next
        
        For i = lngEnd - 1 To lngBegin Step -1 '���뷴��
            '���ø�ҩ;����
            If Val(.TextMatrix(i, COL_ִ������)) = 5 And intִ������ <> 5 Then
                strִ������ = "�Ա�ҩ"
            ElseIf Val(.TextMatrix(i, COL_ִ������)) <> 5 And intִ������ = 5 Then
                strִ������ = "��Ժ��ҩ"
            Else
                strִ������ = ""
            End If
            .TextMatrix(i, COL_���ID) = "" '���������Ϊ��־
            .TextMatrix(i, COL_���ID) = AdviceSet��ҩ;��(i, lng��ҩ;��ID, strִ������)
            
            If InStr(",0,3,", .TextMatrix(i, COL_EDIT)) > 0 Then
                .TextMatrix(i, COL_EDIT) = 2 '���Ϊ���޸�
                .TextMatrix(i, COL_״̬) = 1 '�޸ĺ��Ϊ�¿�
            End If
        Next
        
        mblnRowChange = varTmp1: .Redraw = varTmp2
        mblnNoSave = True '���Ϊδ����
    End With
End Sub

Private Sub ShowAdvice()
'���ܣ���ʾ��ǰ���������µ�ҽ����¼
'˵����1.���ݳ���༭��ʽ,��ص��������ǰ�����ϸ�������һ�ڵġ�
'      2.���ﲻ����һ����ҩ�ı߿��䷽�иߣ�״̬��ɫ�ȸ�ʽ����,�������ڶ�ȡ��༭ʱ����
    Dim lngRow As Long, blnHide As Boolean, i As Long
    
    Screen.MousePointer = 11
    mblnRowChange = False
    vsAdvice.Redraw = flexRDNone
        
    '��ɾ����Ч��
    For i = vsAdvice.Rows - 1 To vsAdvice.FixedRows Step -1
        If vsAdvice.RowData(i) = 0 Then vsAdvice.RemoveItem i
    Next
    
    '���ݵ�ǰ��Ч,Ӥ����ʾ
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, COL_Ӥ��)) = cboӤ��.ListIndex Then
                blnHide = False
                '�������������У�
                '1.��ҩ�ĸ�ҩ;����
                '2.�����ĸ���������������Ŀ��
                '3.�����ϵĲ�λ��
                '4.��ҩ�䷽�����ζ��ҩ����ҩ�巨��
                '5.(һ���ɼ���)������Ŀ
                If .TextMatrix(i, COL_���) = "E" And Val(.TextMatrix(i, COL_���ID)) = 0 Then
                    If Val(.TextMatrix(i - 1, COL_���ID)) = .RowData(i) _
                        And InStr(",5,6,", .TextMatrix(i - 1, COL_���)) > 0 Then
                        blnHide = True
                    End If
                End If
                If InStr(",F,G,D,7,E,C,", .TextMatrix(i, COL_���)) > 0 _
                    And Val(.TextMatrix(i, COL_���ID)) <> 0 Then
                    blnHide = True
                End If
                                
                .RowHidden(i) = blnHide
                If Not blnHide And lngRow = 0 Then lngRow = i
                
                '�������Ƶ���:����Ϊ�ӿ��ٶ�,ֻ��ȡ�¿���,�����Ľ����ٶ�
                If Not .RowHidden(i) _
                    And Val(.TextMatrix(i, COL_״̬)) = 1 And .TextMatrix(i, COL_����) = "" Then
                    .TextMatrix(i, COL_����) = GetItemPrice(i)
                End If
            Else
                .RowHidden(i) = True
            End If
        Next
    End With
    
    'û��������,���һ�п�
    If lngRow = 0 Then
        vsAdvice.AddItem ""
        lngRow = vsAdvice.Rows - 1
    End If
    
    vsAdvice.Row = lngRow
    If vsAdvice.RowData(lngRow) = 0 Then
        vsAdvice.Col = vsAdvice.FixedCols
    Else
        vsAdvice.Col = COL_ҽ������
    End If
    vsAdvice.Redraw = flexRDDirect
    mblnRowChange = True
    
    '��ʾ��ǰ��:����ʱ��FormLoad�д���,�Լӿ��ٶ�
    If Me.Visible Then Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
    
    Call CalcAdviceMoney '��ʾ�¿�ҽ�����
    
    Screen.MousePointer = 0
End Sub

Private Function SaveAdvice() As Boolean
'���ܣ����浱ǰ���˵�ҽ����¼
    Dim arrSQL As Variant, arrDelID() As String
    Dim strSQL As String, dbl���� As Double, i As Long
    
    'Pass�Զ���ҩ���
    If gblnPass And InStr(mstrPrivs, "������ҩ���") > 0 Then
        If AdviceCheckWarn(1) = 3 Then
            If MsgBox("������ҩ���ϵͳ�������ںڵ���ҩ��Ҫ�������������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End If
    
    Screen.MousePointer = 11
    
    '����SQL
    arrSQL = Array()
        
    'ɾ���˵ļ�¼
    arrDelID = Split(mstrDelIDs, ",")
    For i = 0 To UBound(arrDelID)
        If Val(arrDelID(i)) <> 0 Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_Delete(" & Val(arrDelID(i)) & ")"
        End If
    Next
                
    '�༭��־��0-ԭʼ��,1-������,2-�޸�������,3-�޸������
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If .RowData(i) <> 0 Then '����ҽ����¼
                '����ת��
                dbl���� = 0
                If InStr(",1,2,", .TextMatrix(i, COL_EDIT)) > 0 Then
                    If Val(.TextMatrix(i, COL_����)) <> 0 Then
                        If InStr(",5,6,", .TextMatrix(i, COL_���)) > 0 Then
                            '��ҩת�������۵�λ
                            dbl���� = Format(Val(.TextMatrix(i, COL_����)) * Val(.TextMatrix(i, COL_�����װ)), "0.00000")
                        Else
                            '��ҩ�䷽�������ҩ��������,��ת��
                            dbl���� = Val(.TextMatrix(i, COL_����))
                        End If
                    End If
                End If
                
                If Val(.TextMatrix(i, COL_EDIT)) = 3 Then '�޸�����ŵļ�¼
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_�������(" & .RowData(i) & "," & Val(.TextMatrix(i, COL_���)) & ")"
                ElseIf Val(.TextMatrix(i, COL_EDIT)) = 2 Then '�޸������ݵļ�¼
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_Update(" & _
                        .RowData(i) & "," & ZVal(.TextMatrix(i, COL_���ID)) & "," & _
                        Val(.TextMatrix(i, COL_���)) & "," & Val(.TextMatrix(i, COL_״̬)) & ",1," & _
                        Val(.TextMatrix(i, COL_������ĿID)) & "," & ZVal(.TextMatrix(i, COL_����)) & "," & _
                        ZVal(.TextMatrix(i, COL_����)) & "," & ZVal(dbl����) & "," & _
                        "'" & Replace(.TextMatrix(i, COL_ҽ������), "'", "''") & "','" & Replace(.TextMatrix(i, COL_ҽ������), "'", "''") & "'," & _
                        "'" & .TextMatrix(i, COL_�걾��λ) & "','" & .TextMatrix(i, COL_Ƶ��) & "'," & _
                        ZVal(.TextMatrix(i, COL_Ƶ�ʴ���)) & "," & ZVal(.TextMatrix(i, COL_Ƶ�ʼ��)) & "," & _
                        "'" & .TextMatrix(i, COL_�����λ) & "','" & .TextMatrix(i, COL_ִ��ʱ��) & "'," & _
                        Val(.TextMatrix(i, COL_�Ƽ�����)) & "," & ZVal(.TextMatrix(i, COL_ִ�п���ID)) & "," & _
                        Val(.TextMatrix(i, COL_ִ������)) & "," & Val(.TextMatrix(i, COL_��־)) & "," & _
                        "To_Date('" & Format(.Cell(flexcpData, i, COL_��ʼʱ��), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI'),NULL," & _
                        mlng���˿���id & "," & Val(.TextMatrix(i, COL_��������ID)) & ",'" & .TextMatrix(i, COL_����ҽ��) & "'," & _
                        "To_Date('" & Format(.Cell(flexcpData, i, COL_����ʱ��), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI'))"
                ElseIf Val(.TextMatrix(i, COL_EDIT)) = 1 Then '�����ļ�¼
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_Insert(" & _
                        .RowData(i) & "," & ZVal(.TextMatrix(i, COL_���ID)) & "," & _
                        Val(.TextMatrix(i, COL_���)) & ",1," & mlng����ID & ",NULL," & _
                        Val(.TextMatrix(i, COL_Ӥ��)) & "," & Val(.TextMatrix(i, COL_״̬)) & ",1," & _
                        "'" & .TextMatrix(i, COL_���) & "'," & Val(.TextMatrix(i, COL_������ĿID)) & "," & _
                        ZVal(.TextMatrix(i, COL_�շ�ϸĿID)) & "," & _
                        ZVal(.TextMatrix(i, COL_����)) & "," & ZVal(.TextMatrix(i, COL_����)) & "," & ZVal(dbl����) & "," & _
                        "'" & Replace(.TextMatrix(i, COL_ҽ������), "'", "''") & "','" & Replace(.TextMatrix(i, COL_ҽ������), "'", "''") & "'," & _
                        "'" & .TextMatrix(i, COL_�걾��λ) & "','" & .TextMatrix(i, COL_Ƶ��) & "'," & _
                        ZVal(.TextMatrix(i, COL_Ƶ�ʴ���)) & "," & ZVal(.TextMatrix(i, COL_Ƶ�ʼ��)) & "," & _
                        "'" & .TextMatrix(i, COL_�����λ) & "','" & .TextMatrix(i, COL_ִ��ʱ��) & "'," & _
                        Val(.TextMatrix(i, COL_�Ƽ�����)) & "," & ZVal(.TextMatrix(i, COL_ִ�п���ID)) & "," & _
                        Val(.TextMatrix(i, COL_ִ������)) & "," & Val(.TextMatrix(i, COL_��־)) & "," & _
                        "To_Date('" & Format(.Cell(flexcpData, i, COL_��ʼʱ��), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI'),NULL," & _
                        mlng���˿���id & "," & Val(.TextMatrix(i, COL_��������ID)) & ",'" & .TextMatrix(i, COL_����ҽ��) & "'," & _
                        "To_Date('" & Format(.Cell(flexcpData, i, COL_����ʱ��), "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')," & _
                        "'" & mstr�Һŵ� & "'," & ZVal(mlngǰ��ID) & ")"
                End If
                
                'Pass:���������
                If Val(.Cell(flexcpData, i, COL_���)) = 1 Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_�������(" & .RowData(i) & "," & _
                        IIF(CStr(.Cell(flexcpData, i, COL_��ʾ)) = "", "NULL", Val(.Cell(flexcpData, i, COL_��ʾ))) & ")"
                End If
            End If
        Next
    End With
    
    '�ύ����
    On Error GoTo errH
    gcnOracle.BeginTrans
    For i = 0 To UBound(arrSQL)
        zlDatabase.ExecuteProcedure CStr(arrSQL(i)), Me.Caption
    Next
    gcnOracle.CommitTrans
    On Error GoTo 0
    
    '����ɹ���,���м�¼���ԭʼ��¼
    With vsAdvice
        For i = vsAdvice.FixedRows To vsAdvice.Rows - 1
            If .RowData(i) <> 0 Then
                .TextMatrix(i, COL_EDIT) = 0
                .Cell(flexcpData, i, COL_���) = Empty 'Pass:����������־
            End If
        Next
    End With
    
    '��������½�����(���翪ʼʱ�䲻׼����)
    Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)

    Screen.MousePointer = 0
    mblnNoSave = False
    mstrDelIDs = ""
    SaveAdvice = True
    mblnOK = True
    Exit Function
errH:
    Screen.MousePointer = 0
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function LoadAdvice() As Boolean
'���ܣ���ȡ��ǰ���˵�ҽ����¼
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, bln�䷽ As Boolean
    Dim blnFirst As Boolean, i As Long, j As Long
    
    Screen.MousePointer = 11
    
    On Error GoTo errH
    
    '��ҽ��ȱʡ������
    If msng���� = 0 Then msng���� = 1
    
    'ҽ���༭ʱֻ��ʾҽ����,ҽ���༭ʱ����ʾҽ����
    strSQL = IIF(mlngǰ��ID <> 0, " And A.ǰ��ID+0=[3]", " And A.ǰ��ID is NULL")
    
    '��ȡ"1-�¿�,8-��ֹͣ(�ѷ���)"��ҽ��
    strSQL = _
        " Select A.ID,A.���ID,Nvl(A.Ӥ��,0) as Ӥ��,A.���,A.ҽ����Ч," & _
        " A.ҽ��״̬,A.�������,A.������ĿID,B.����,A.�걾��λ,A.�շ�ϸĿID," & _
        " A.��ʼִ��ʱ��,A.ҽ������,A.ҽ������,A.��������,A.����,A.�ܸ�����,B.���㵥λ," & _
        " A.ִ��Ƶ��,A.Ƶ�ʴ���,A.Ƶ�ʼ��,A.�����λ,B.���㷽ʽ,B.ִ��Ƶ��,B.��������," & _
        " A.�Ƽ�����,A.ִ��ʱ�䷽��,A.ִ������,A.ִ�п���ID,A.��������ID,A.����ҽ��,A.����ʱ��," & _
        " A.������־,C.��������,C.����ְ��,C.�������,C.ҩƷ����," & _
        " D.����ϵ��,D.�����װ,D.���ﵥλ,D.�ɷ����,A.����ID,A.�����," & _
        " Decode(S.ǩ��ID,NULL,0,1) as ǩ����" & _
        " From ����ҽ����¼ A,������ĿĿ¼ B,ҩƷ���� C,ҩƷ��� D,����ҽ��״̬ S" & _
        " Where Nvl(A.ҽ����Ч,0)=1 And A.������ĿID=B.ID And A.������ĿID=C.ҩ��ID(+)" & _
        " And A.�շ�ϸĿID=D.ҩƷID(+) And A.ҽ��״̬ IN(1,8)" & strSQL & _
        " And A.����ID+0=[1] And A.�Һŵ�=[2] And A.��ʼִ��ʱ�� is Not NULL And A.������Դ<>3" & _
        " And A.ID=S.ҽ��ID And S.��������=1" & _
        " Order by Ӥ��,���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mstr�Һŵ�, mlngǰ��ID)
    On Error GoTo 0
    
    If Not rsTmp.EOF Then
        mblnRowChange = False
        With vsAdvice
            .Redraw = flexRDNone
            .Rows = .FixedRows + rsTmp.RecordCount
            For i = 1 To rsTmp.RecordCount
                bln�䷽ = False
                
                .RowData(i) = CLng(rsTmp!ID)
                .TextMatrix(i, COL_EDIT) = 0 'ԭʼ��¼
                .TextMatrix(i, COL_���ID) = Nvl(rsTmp!���ID)
                .TextMatrix(i, COL_Ӥ��) = Nvl(rsTmp!Ӥ��, 0)
                .TextMatrix(i, COL_���) = rsTmp!���
                .TextMatrix(i, COL_״̬) = Nvl(rsTmp!ҽ��״̬, 0)
                
                .TextMatrix(i, COL_���) = rsTmp!�������
                .TextMatrix(i, COL_������ĿID) = rsTmp!������ĿID
                .TextMatrix(i, COL_����) = rsTmp!����
                .TextMatrix(i, COL_�걾��λ) = Nvl(rsTmp!�걾��λ)
                .TextMatrix(i, COL_�շ�ϸĿID) = Nvl(rsTmp!�շ�ϸĿID)
                .TextMatrix(i, COL_ҽ������) = Nvl(rsTmp!ҽ������)
                .TextMatrix(i, COL_ҽ������) = Nvl(rsTmp!ҽ������)
                
                .TextMatrix(i, COL_�Ƽ�����) = Nvl(rsTmp!�Ƽ�����, 0)
                .TextMatrix(i, COL_���㷽ʽ) = Nvl(rsTmp!���㷽ʽ, 0)
                .TextMatrix(i, COL_Ƶ������) = Nvl(rsTmp!ִ��Ƶ��, 0)
                .TextMatrix(i, COL_��������) = Nvl(rsTmp!��������)
                .TextMatrix(i, COL_�������) = Nvl(rsTmp!�������)
                .TextMatrix(i, COL_ҩƷ����) = Nvl(rsTmp!ҩƷ����)
                .TextMatrix(i, COL_��������) = Nvl(rsTmp!��������)
                .TextMatrix(i, COL_����ְ��) = Nvl(rsTmp!����ְ��)
                .TextMatrix(i, COL_����ϵ��) = Nvl(rsTmp!����ϵ��)
                .TextMatrix(i, COL_�����װ) = Nvl(rsTmp!�����װ)
                .TextMatrix(i, COL_���ﵥλ) = Nvl(rsTmp!���ﵥλ)
                If Not IsNull(rsTmp!����ϵ��) Then
                    .TextMatrix(i, COL_�ɷ����) = Nvl(rsTmp!�ɷ����, 0)
                End If
                
                .TextMatrix(i, COL_��ʼʱ��) = Format(rsTmp!��ʼִ��ʱ��, "MM-dd HH:mm")
                .Cell(flexcpData, i, COL_��ʼʱ��) = Format(rsTmp!��ʼִ��ʱ��, "yyyy-MM-dd HH:mm")
                
                .TextMatrix(i, COL_Ƶ��) = Nvl(rsTmp!ִ��Ƶ��)
                .TextMatrix(i, COL_Ƶ�ʴ���) = Nvl(rsTmp!Ƶ�ʴ���)
                .TextMatrix(i, COL_Ƶ�ʼ��) = Nvl(rsTmp!Ƶ�ʼ��)
                .TextMatrix(i, COL_�����λ) = Nvl(rsTmp!�����λ)
                .TextMatrix(i, COL_ִ��ʱ��) = Nvl(rsTmp!ִ��ʱ�䷽��)
                
                .TextMatrix(i, COL_ִ�п���ID) = Nvl(rsTmp!ִ�п���ID)
                .TextMatrix(i, COL_ִ������) = Nvl(rsTmp!ִ������, 0)
                
                If rsTmp!������� = "E" Then
                    If Nvl(rsTmp!���ID, 0) = 0 And Val(.TextMatrix(i - 1, COL_���ID)) = rsTmp!ID Then
                        If InStr(",5,6,", .TextMatrix(i - 1, COL_���)) > 0 Then
                            '��ǰ��¼�ǳ�ҩ�ĸ�ҩ;��,������һ����ҩ��
                            For j = i - 1 To .FixedRows Step -1
                                If Val(.TextMatrix(j, COL_���ID)) = rsTmp!ID Then
                                    '��ʾ��ҩ;��
                                    .TextMatrix(j, COL_�÷�) = rsTmp!����
                                Else
                                    Exit For
                                End If
                            Next
                        ElseIf InStr(",E,7,", .TextMatrix(i - 1, COL_���)) > 0 Then
                            '��ǰ��¼����ҩ�䷽���÷�,���䷽��ʾ��
                            .TextMatrix(i, COL_�÷�) = rsTmp!����
                            bln�䷽ = True
                        ElseIf .TextMatrix(i - 1, COL_���) = "C" Then
                            .TextMatrix(i, COL_�÷�) = rsTmp!����
                        End If
                    ElseIf Not IsNull(rsTmp!���ID) Then
                        '��ǰ��¼����ҩ�䷽�巨��
                        bln�䷽ = True
                    End If
                ElseIf rsTmp!������� = "7" Then
                    bln�䷽ = True
                End If
                
                '����
                .TextMatrix(i, COL_����) = FormatEx(Nvl(rsTmp!��������), 5)
                If InStr(",5,6,7,", rsTmp!�������) > 0 Or Nvl(rsTmp!���㷽ʽ, 0) <> 3 Then
                    .TextMatrix(i, COL_������λ) = Nvl(rsTmp!���㵥λ)
                End If
                
                '����
                .TextMatrix(i, COL_����) = Nvl(rsTmp!����, 0)
                'ȡ����¿�ҽ���Ŀ�����Ϊȱʡ����
                If InStr(",1,2,", Nvl(rsTmp!ҽ��״̬, 0)) > 0 _
                    And InStr(",5,6,", rsTmp!�������) > 0 And Nvl(rsTmp!����, 0) <> 0 Then
                    msng���� = Nvl(rsTmp!����, 1)
                End If
                
                '����
                If InStr(",5,6,", rsTmp!�������) > 0 Then
                    '��ҩ����������,�����۵�λ���,���ﵥλ��ʾ
                    If Not IsNull(rsTmp!�ܸ�����) And Not IsNull(rsTmp!�����װ) Then
                        .TextMatrix(i, COL_����) = FormatEx(rsTmp!�ܸ����� / rsTmp!�����װ, 5)
                    End If
                    .TextMatrix(i, COL_������λ) = Nvl(rsTmp!���ﵥλ)
                Else
                    '�����������ҩ����������
                    If Not IsNull(rsTmp!�ܸ�����) Then
                        .TextMatrix(i, COL_����) = rsTmp!�ܸ�����
                    End If
                    If bln�䷽ Then
                        .TextMatrix(i, COL_������λ) = "��" '��ҩ�䷽������λΪ"��"
                    Else
                        .TextMatrix(i, COL_������λ) = Nvl(rsTmp!���㵥λ)
                    End If
                End If

                .TextMatrix(i, COL_��������ID) = rsTmp!��������ID
                .TextMatrix(i, COL_����ҽ��) = rsTmp!����ҽ��
                
                .TextMatrix(i, COL_����ʱ��) = Format(rsTmp!����ʱ��, "MM-dd HH:mm")
                .Cell(flexcpData, i, COL_����ʱ��) = Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm")
                
                '��ʾ������־:һ����ҩֻ��ʾ�ڵ�һ��
                .TextMatrix(i, COL_��־) = Nvl(rsTmp!������־, 0)
                blnFirst = True
                If InStr(",5,6,", .TextMatrix(i, COL_���)) > 0 Then
                    If Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(i - 1, COL_���ID)) Then
                        blnFirst = False
                    End If
                End If
                If blnFirst Then
                    If Nvl(rsTmp!������־, 0) = 1 Then
                        Set .Cell(flexcpPicture, i, COL_F��־) = imgFlag.ListImages("����").Picture
                    End If
                End If
                
                '����ҽ��״̬,ҩƷ����������ɫ
                '-------------------------------------------------------------------
                .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = .ForeColor
                If rsTmp!ҽ��״̬ = 8 Then
                    '��ֹͣ(�ѷ���)
                    .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &HC00000 '����
                End If
                
                '���龫ҩƷ��ʶ:��ҩ�䷽�����ζ��ҩ������
                If InStr(",5,6,", rsTmp!�������) > 0 And Not IsNull(rsTmp!�������) Then
                    If InStr(",����ҩ,����ҩ,����ҩ,", rsTmp!�������) > 0 Then
                        .Cell(flexcpFontBold, i, COL_ҽ������) = True
                    End If
                End If
                
                'Pass�����������ʾ��ʾ��
                If Not IsNull(rsTmp!�����) Then
                    .Cell(flexcpData, i, COL_��ʾ) = CStr(Nvl(rsTmp!�����))
                    Set .Cell(flexcpPicture, i, COL_��ʾ) = imgPass.ListImages(rsTmp!����� + 1).Picture
                End If
                
                '����ǩ����ʶ
                .TextMatrix(i, COL_ǩ����) = Nvl(rsTmp!ǩ����)
                If Val(.TextMatrix(i, COL_ǩ����)) = 1 Then
                    Set .Cell(flexcpPicture, i, COL_ҽ������) = imgSign.ListImages(1).Picture
                End If
                
                rsTmp.MoveNext
            Next
            
            '�̶���ͼ�����:����Ϊ�ж���,��Ȼ���߿�ʱ����������
            .Cell(flexcpPictureAlignment, .FixedRows, 0, .Rows - 1, .FixedCols - 1) = 4
            '����ǩ��ͼ�����
            .Cell(flexcpPictureAlignment, .FixedRows, COL_ҽ������, .Rows - 1, COL_ҽ������) = 0
            
            Call .AutoSize(COL_ҽ������)
            .Redraw = flexRDDirect
        End With
        mblnRowChange = True
    Else
        mblnRowChange = False
        vsAdvice.Rows = vsAdvice.FixedRows
        vsAdvice.Rows = vsAdvice.FixedRows + 1
        mblnRowChange = True
    End If

    Screen.MousePointer = 0
    LoadAdvice = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Check��ʼʱ��(ByVal strStart As String, Optional ByVal blnMsg As Boolean = True, Optional strMsg As String) As Boolean
'���ܣ��������Ŀ�ʼʱ���Ƿ�Ϸ�
'˵����
'1.��ʼʱ�䲻��С�ڲ��˵ĹҺ�ʱ��
'2.����¼��ʱ,��ʼʱ�䲻��С�ڵ�ǰʱ��֮ǰ30����(�Ӷ�������ɿ���ʱ����ڿ�ʼʱ��30����)
    If Not IsDate(strStart) Then
        MsgBox "�����ҽ����ʼִ��ʱ����Ч��", vbInformation, gstrSysName
        Exit Function
    End If
        
    If Format(strStart, "yyyy-MM-dd HH:mm") < Format(mdat�Һ�ʱ��, "yyyy-MM-dd HH:mm") Then
        strMsg = "ҽ���Ŀ�ʼִ��ʱ�䲻��С�ڲ��˵ĹҺ�ʱ�� " & Format(mdat�Һ�ʱ��, "MM-dd HH:mm") & " ��"
        If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
        Exit Function
    End If
    
'    If DateDiff("n", CDate(strStart), zlDatabase.Currentdate) > TIME_LIMIT Then
'        strMsg = "��ʼִ��ʱ�䲻��̫���ڵ�ǰʱ�䡣"
'        If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
'        Exit Function
'    End If
    
    Check��ʼʱ�� = True
End Function

Private Function Check����ʱ��(ByVal strDate As String, ByVal strStart As String, _
    Optional ByVal blnMsg As Boolean = True, Optional strMsg As String) As Boolean
'���ܣ���鿪��ʱ���Ƿ���Ч
'˵������ӦС�ڲ��˹Һ�ʱ��
    If Not IsDate(strDate) Then
        strMsg = "����Ŀ���ʱ����Ч��"
        If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
        Exit Function
    End If
        
'    If Format(strDate, "yyyy-MM-dd HH:mm") < Format(mdat�Һ�ʱ��, "yyyy-MM-dd HH:mm") Then
'        strMsg = "����ʱ�䲻��С�ڲ��˵ĹҺ�ʱ�� " & Format(mdat�Һ�ʱ��, "MM-dd HH:mm") & " ��"
'        If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
'        Exit Function
'    End If
    Check����ʱ�� = True
End Function

Private Function Check�������(ByVal strҩƷIDs As String) As Boolean
'���ܣ��������ҩ,�г�ҩ���������;��ҩ�䷽����������
'������strҩƷIDs="1,2,3,..."
    Dim rsTmp As New ADODB.Recordset
    Dim rsMain As New ADODB.Recordset
    Dim strSQL As String, i As Long, j As Long, k As Long
    Dim arr���� As Variant, arr���� As Variant
    Dim arrItems As Variant, strMsg As String, strTmp As String
    Dim lng��ĿID As Long, str���� As String, blnδ�༭ As Boolean
    Dim lng���� As Long, lngRow As Long, lngSeekRow As Long
    
    On Error GoTo errH
        
    arr���� = Array(): arr���� = Array()
    
    strSQL = "Select ���� From ���ƻ�����Ŀ" & _
        " Where ��ĿID IN(" & strҩƷIDs & ") Group by ���� Having Count(*)>1"
    Call zlDatabase.OpenRecordset(rsMain, strSQL, Me.Caption) 'In
    For k = 1 To rsMain.RecordCount
        strSQL = "Select A.����,A.����,A.��ĿID,B.����" & _
            " From ���ƻ�����Ŀ A,������ĿĿ¼ B" & _
            " Where A.��ĿID=B.ID And A.����=" & rsMain!���� & _
            " And A.��ĿID IN(" & strҩƷIDs & ")" & _
            " Order by A.����,B.����"
        zlDatabase.OpenRecordset rsTmp, strSQL, Me.Caption 'In:��ĿID������
        For i = 1 To rsTmp.RecordCount
            If rsTmp!���� <> lng���� Then
                If rsTmp!���� = 1 Then
                    ReDim Preserve arr����(UBound(arr����) + 1)
                Else
                    ReDim Preserve arr����(UBound(arr����) + 1)
                End If
                lng���� = rsTmp!����
            End If
            If rsTmp!���� = 1 Then
                arr����(UBound(arr����)) = arr����(UBound(arr����)) & Chr(234) & rsTmp!��ĿID & Chr(8) & rsTmp!����
            Else
                arr����(UBound(arr����)) = arr����(UBound(arr����)) & Chr(234) & rsTmp!��ĿID & Chr(8) & rsTmp!����
            End If
            rsTmp.MoveNext
        Next
        rsMain.MoveNext
    Next
    
    '�ȼ����ò���(��ֹ����)
    If UBound(arr����) >= 0 Then
        strMsg = "": lngSeekRow = 0
        For i = 0 To UBound(arr����) 'ÿ��
            strTmp = "": blnδ�༭ = True
            arrItems = Split(Mid(arr����(i), 2), Chr(234))
            For j = 0 To UBound(arrItems) 'ÿ��Ŀ
                lng��ĿID = Split(arrItems(j), Chr(8))(0)
                str���� = Split(arrItems(j), Chr(8))(1)
                strTmp = strTmp & "��" & str����
                
                'Ϊ�˶�λ,��ҽ���в��ұ����������޸ĵĸ���Ŀ(�����ж��)������
                lngRow = -1
                Do While True
                    lngRow = vsAdvice.FindRow(CStr(lng��ĿID), lngRow + 1, COL_������ĿID)
                    If lngRow = -1 Then
                        Exit Do
                    ElseIf InStr(",1,2,", vsAdvice.TextMatrix(lngRow, COL_EDIT)) > 0 Then
                        If lngSeekRow = 0 Or lngRow < lngSeekRow Then lngSeekRow = lngRow '�༭������С�����ȶ�λ
                        blnδ�༭ = False: Exit Do
                    End If
                Loop
            Next
            If Not blnδ�༭ Then '���һ���е���Ŀ�ڱ��ζ�δ�༭��,�򲻹�
                strMsg = strMsg & vbCrLf & "�� " & Mid(strTmp, 2)
            End If
        Next
        If strMsg <> "" Then
            If lngSeekRow <> 0 Then
                vsAdvice.Col = COL_ҽ������: vsAdvice.Row = lngSeekRow
                Call vsAdvice.ShowCell(vsAdvice.Row, vsAdvice.Col)
            End If
            MsgBox "�ڲ���ҽ���з�������ҩƷ������ã�" & strMsg, vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '�ټ�����ò���(�����Ƿ����)
    If UBound(arr����) >= 0 Then
        strMsg = "": lngSeekRow = 0
        For i = 0 To UBound(arr����) 'ÿ��
            strTmp = "": blnδ�༭ = True
            arrItems = Split(Mid(arr����(i), 2), Chr(234))
            For j = 0 To UBound(arrItems) 'ÿ��Ŀ
                lng��ĿID = Split(arrItems(j), Chr(8))(0)
                str���� = Split(arrItems(j), Chr(8))(1)
                strTmp = strTmp & "��" & str����
                
                'Ϊ�˶�λ,��ҽ���в��ұ����������޸ĵĸ���Ŀ(�����ж��)������
                lngRow = -1
                Do While True
                    lngRow = vsAdvice.FindRow(CStr(lng��ĿID), lngRow + 1, COL_������ĿID)
                    If lngRow = -1 Then
                        Exit Do
                    ElseIf InStr(",1,2,", vsAdvice.TextMatrix(lngRow, COL_EDIT)) > 0 Then
                        If lngSeekRow = 0 Or lngRow < lngSeekRow Then lngSeekRow = lngRow '�༭������С�����ȶ�λ
                        blnδ�༭ = False: Exit Do
                    End If
                Loop
            Next
            If Not blnδ�༭ Then '���һ���е���Ŀ�ڱ��ζ�δ�༭��,�򲻹�
                strMsg = strMsg & vbCrLf & "�� " & Mid(strTmp, 2)
            End If
        Next
        If strMsg <> "" Then
            If lngSeekRow <> 0 Then
                vsAdvice.Col = COL_ҽ������: vsAdvice.Row = lngSeekRow
                Call vsAdvice.ShowCell(vsAdvice.Row, vsAdvice.Col)
            End If
            If MsgBox("�ڲ���ҽ���з�������ҩƷ�������ã�" & strMsg & vbCrLf & vbCrLf & "Ҫ������", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End If
    
    Check������� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Check���ƻ���(ByVal str����IDs As String) As Boolean
'���ܣ�����ҩƷ(��ҩ,��ҩ)�Ļ���
'������str����IDs="1,2,3,..."
    Dim rsTmp As New ADODB.Recordset
    Dim rsMain As New ADODB.Recordset
    Dim strSQL As String, i As Long, j As Long, k As Long
    Dim arr���� As Variant, arr��ֹ As Variant, arrֹͣ As Variant
    Dim arrItems As Variant, strMsg As String, strTmp As String
    Dim lng��ĿID As Long, str���� As String, blnδ�༭ As Boolean
    Dim lng���� As Long, lngRow As Long, lngSeekRow As Long
    
    On Error GoTo errH
        
    arr���� = Array(): arr��ֹ = Array(): arrֹͣ = Array()
    
    strSQL = "Select ���� From ���ƻ�����Ŀ" & _
        " Where ��ĿID IN(" & str����IDs & ") Group by ���� Having Count(*)>1"
    Call zlDatabase.OpenRecordset(rsMain, strSQL, Me.Caption) 'In
    For k = 1 To rsMain.RecordCount
        strSQL = "Select A.����,A.������,A.����,A.��ĿID,B.����" & _
            " From ���ƻ�����Ŀ A,������ĿĿ¼ B" & _
            " Where A.��ĿID=B.ID And A.����=" & rsMain!���� & _
            " And A.��ĿID IN(" & str����IDs & ")" & _
            " Order by A.����,B.����"
        zlDatabase.OpenRecordset rsTmp, strSQL, Me.Caption 'In:��ĿID������
        For i = 1 To rsTmp.RecordCount
            If rsTmp!���� <> lng���� Then
                If rsTmp!���� = 1 Then
                    ReDim Preserve arr����(UBound(arr����) + 1)
                    arr����(UBound(arr����)) = rsTmp!������
                ElseIf rsTmp!���� = 2 Then
                    ReDim Preserve arr��ֹ(UBound(arr��ֹ) + 1)
                    arr��ֹ(UBound(arr��ֹ)) = rsTmp!������
                Else
                    ReDim Preserve arrֹͣ(UBound(arrֹͣ) + 1)
                    arrֹͣ(UBound(arrֹͣ)) = rsTmp!������
                End If
                lng���� = rsTmp!����
            End If
            If rsTmp!���� = 1 Then
                arr����(UBound(arr����)) = arr����(UBound(arr����)) & Chr(234) & rsTmp!��ĿID & Chr(8) & rsTmp!����
            ElseIf rsTmp!���� = 2 Then
                arr��ֹ(UBound(arr��ֹ)) = arr��ֹ(UBound(arr��ֹ)) & Chr(234) & rsTmp!��ĿID & Chr(8) & rsTmp!����
            Else
                arrֹͣ(UBound(arrֹͣ)) = arrֹͣ(UBound(arrֹͣ)) & Chr(234) & rsTmp!��ĿID & Chr(8) & rsTmp!����
            End If
            rsTmp.MoveNext
        Next
        rsMain.MoveNext
    Next
    '�ȼ���ֹ��������
    If UBound(arr��ֹ) >= 0 Then
        strMsg = "": lngSeekRow = 0
        For i = 0 To UBound(arr��ֹ) 'ÿ��
            strTmp = "": blnδ�༭ = True
            arrItems = Split(arr��ֹ(i), Chr(234))
            For j = 1 To UBound(arrItems) 'ÿ��Ŀ
                lng��ĿID = Split(arrItems(j), Chr(8))(0)
                str���� = Split(arrItems(j), Chr(8))(1)
                strTmp = strTmp & vbCrLf & vbTab & str����
                
                'Ϊ�˶�λ,��ҽ���в��ұ����������޸ĵĸ���Ŀ(�����ж��)������
                lngRow = -1
                Do While True
                    lngRow = vsAdvice.FindRow(CStr(lng��ĿID), lngRow + 1, COL_������ĿID)
                    If lngRow = -1 Then
                        Exit Do
                    ElseIf Val(vsAdvice.TextMatrix(lngRow, COL_EDIT)) = 1 Then
                        If lngSeekRow = 0 Or lngRow < lngSeekRow Then lngSeekRow = lngRow '�༭������С�����ȶ�λ
                        blnδ�༭ = False: Exit Do
                    End If
                Loop
            Next
            If Not blnδ�༭ Then '���һ���е���Ŀ�ڱ��ζ�δ�༭��,�򲻹�
                strMsg = strMsg & vbCrLf & vbCrLf & arrItems(0) & "��" & Mid(strTmp, 2)
            End If
        Next
        If strMsg <> "" Then
            If lngSeekRow <> 0 Then
                vsAdvice.Col = COL_ҽ������: vsAdvice.Row = lngSeekRow
                Call vsAdvice.ShowCell(vsAdvice.Row, vsAdvice.Col)
            End If
            MsgBox "�ڲ���ҽ���з����������ݻ����ų⣺" & strMsg, vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '�ټ���Զ�ֹͣ����,���ﴦ��Ϊ��ֹ(����)
    If UBound(arrֹͣ) >= 0 Then
        strMsg = "": lngSeekRow = 0
        For i = 0 To UBound(arrֹͣ) 'ÿ��
            strTmp = "": blnδ�༭ = True
            arrItems = Split(arrֹͣ(i), Chr(234))
            For j = 1 To UBound(arrItems) 'ÿ��Ŀ
                lng��ĿID = Split(arrItems(j), Chr(8))(0)
                str���� = Split(arrItems(j), Chr(8))(1)
                strTmp = strTmp & vbCrLf & vbTab & str����
                
                'Ϊ�˶�λ,��ҽ���в��ұ����������޸ĵĸ���Ŀ(�����ж��)������
                lngRow = -1
                Do While True
                    lngRow = vsAdvice.FindRow(CStr(lng��ĿID), lngRow + 1, COL_������ĿID)
                    If lngRow = -1 Then
                        Exit Do
                    ElseIf InStr(",1,2,", vsAdvice.TextMatrix(lngRow, COL_EDIT)) > 0 Then
                        If lngSeekRow = 0 Or lngRow < lngSeekRow Then lngSeekRow = lngRow '�༭������С�����ȶ�λ
                        blnδ�༭ = False ': Exit Do
                    End If
                Loop
            Next
            If Not blnδ�༭ Then '���һ���е���Ŀ�ڱ��ζ�δ�༭��,�򲻹�
                strMsg = strMsg & vbCrLf & vbCrLf & arrItems(0) & "��" & Mid(strTmp, 2)
            End If
        Next
        If strMsg <> "" Then
            If lngSeekRow <> 0 Then
                vsAdvice.Col = COL_ҽ������: vsAdvice.Row = lngSeekRow
                Call vsAdvice.ShowCell(vsAdvice.Row, vsAdvice.Col)
            End If
            MsgBox "�ڲ���ҽ���з����������ݻ����ų⣺" & strMsg, vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '�ټ�������Ƿ��������
    If UBound(arr����) >= 0 Then
        strMsg = "": lngSeekRow = 0
        For i = 0 To UBound(arr����) 'ÿ��
            strTmp = "": blnδ�༭ = True
            arrItems = Split(arr����(i), Chr(234))
            For j = 1 To UBound(arrItems) 'ÿ��Ŀ
                lng��ĿID = Split(arrItems(j), Chr(8))(0)
                str���� = Split(arrItems(j), Chr(8))(1)
                strTmp = strTmp & vbCrLf & vbTab & str����
                
                'Ϊ�˶�λ,��ҽ���в��ұ����������޸ĵĸ���Ŀ(�����ж��)������
                lngRow = -1
                Do While True
                    lngRow = vsAdvice.FindRow(CStr(lng��ĿID), lngRow + 1, COL_������ĿID)
                    If lngRow = -1 Then
                        Exit Do
                    ElseIf InStr(",1,2,", vsAdvice.TextMatrix(lngRow, COL_EDIT)) > 0 Then
                        If lngSeekRow = 0 Or lngRow < lngSeekRow Then lngSeekRow = lngRow '�༭������С�����ȶ�λ
                        blnδ�༭ = False: Exit Do
                    End If
                Loop
            Next
            If Not blnδ�༭ Then '���һ���е���Ŀ�ڱ��ζ�δ�༭��,�򲻹�
                strMsg = strMsg & vbCrLf & vbCrLf & arrItems(0) & "��" & Mid(strTmp, 2)
            End If
        Next
        If strMsg <> "" Then
            If lngSeekRow <> 0 Then
                vsAdvice.Col = COL_ҽ������: vsAdvice.Row = lngSeekRow
                Call vsAdvice.ShowCell(vsAdvice.Row, vsAdvice.Col)
            End If
            If MsgBox("�ڲ���ҽ���з����������ݻ����ų⣺" & strMsg & vbCrLf & vbCrLf & "Ҫ������", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End If
    
    Check���ƻ��� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckStock(ByVal lngRow As Long) As String
'���ܣ����ָ��ҩƷ�еĿ�����
'���أ���=��ʾͨ��
    Dim dbl���� As Double, strMsg As String
    Dim lngִ�п���ID As Long, i As Integer
    
    With vsAdvice
        If InStr(",5,6,", .TextMatrix(lngRow, COL_���)) > 0 Then
            If GetStockCheck(Val(.TextMatrix(lngRow, COL_ִ�п���ID))) <> 0 Then
                If .TextMatrix(lngRow, COL_���) <> "" Then
                    '��ҩ����ֱ�Ӽ������
                    dbl���� = Val(.TextMatrix(lngRow, COL_����))
                    If dbl���� > 0 Then
                        If dbl���� > Val(.TextMatrix(lngRow, COL_���)) Then
                            strMsg = """" & .TextMatrix(lngRow, COL_ҽ������) & """������ѣ�" & _
                                vbCrLf & vbCrLf & Get��������(Val(.TextMatrix(lngRow, COL_ִ�п���ID))) & _
                                "��ǰ���ÿ��Ϊ " & FormatEx(Val(.TextMatrix(lngRow, COL_���)), 5) & _
                                .TextMatrix(lngRow, COL_���ﵥλ) & "������ " & _
                                FormatEx(dbl����, 5) & .TextMatrix(lngRow, COL_���ﵥλ) & "��"
                        End If
                    End If
                End If
            End If
        ElseIf RowIn�䷽��(lngRow) And Val(.TextMatrix(lngRow, COL_����)) <> 0 Then
            '���ݸ�����������
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_���ID)) = .RowData(lngRow) Then
                    If .TextMatrix(i, COL_���) = "7" And .TextMatrix(i, COL_���) <> "" Then
                        '����=�����װ(��ζ����*����)
                        '��ҩҩ����λ�����ɷ��㴦��:ÿ��
                        If Val(.TextMatrix(i, COL_�ɷ����)) = 0 Then
                            dbl���� = Val(.TextMatrix(i, COL_����)) * Val(.TextMatrix(i, COL_����)) / Val(.TextMatrix(i, COL_����ϵ��)) / Val(.TextMatrix(i, COL_�����װ))
                        Else
                            dbl���� = Val(.TextMatrix(i, COL_����)) * IntEx(Val(.TextMatrix(i, COL_����)) / Val(.TextMatrix(i, COL_����ϵ��)) / Val(.TextMatrix(i, COL_�����װ)))
                        End If
                        If dbl���� > Val(.TextMatrix(i, COL_���)) Then
                            lngִ�п���ID = Val(.TextMatrix(i, COL_ִ�п���ID))
                            If GetStockCheck(lngִ�п���ID) = 0 Then Exit For
                            
                            strMsg = strMsg & vbCrLf & .TextMatrix(i, COL_ҽ������) & _
                                "���������� " & FormatEx(dbl����, 5) & .TextMatrix(i, COL_���ﵥλ) & _
                                "�����ÿ�� " & FormatEx(Val(.TextMatrix(i, COL_���)), 5) & .TextMatrix(i, COL_���ﵥλ)
                        End If
                    End If
                Else
                    Exit For
                End If
            Next
            If strMsg <> "" Then
                strMsg = "��ҩ�䷽������ѣ�" & Get��������(lngִ�п���ID) & "������ζҩ��治�㣺" & vbCrLf & strMsg
            End If
        End If
    End With
    CheckStock = strMsg
End Function

Private Function CheckMoney() As Boolean
'���ܣ����ñ������
'˵�����������ۼƷ��ñ�����ʽʱ,ֻ���ѡ�
    Dim rsTmp As New ADODB.Recordset
    Dim blnҽ�� As Boolean, strSQL As String
    Dim curԤ�� As Currency, cur��� As Currency
    
    '�������
    strSQL = "Select Ԥ�����,Nvl(Ԥ�����,0)-Nvl(�������,0) as ��� From ������� Where ����=1 And ����ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
    If Not rsTmp.EOF Then
        curԤ�� = Nvl(rsTmp!Ԥ�����, 0)
        cur��� = Nvl(rsTmp!���, 0)
    End If
    
    '��Ԥ����Ĳ��˲��ж�
    If curԤ�� <> 0 Then
        '�Ƿ�ҽ��
        strSQL = "Select B.���� From ������Ϣ A,ҽ�Ƹ��ʽ B Where A.ҽ�Ƹ��ʽ=B.����(+) And A.����ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
        blnҽ�� = Nvl(rsTmp!����) = "1"
            
        '����ֵ:NULL��0������ͬ���崦��
        strSQL = "Select ����ֵ From ���ʱ�����" & _
            " Where ��������=1 And Nvl(����ID,0)=0" & _
            " And ����ֵ is Not NULL And ���ò���=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, IIF(blnҽ��, 2, 1))
        If Not rsTmp.EOF Then
            If cur��� < Nvl(rsTmp!����ֵ, 0) Then
                If MsgBox("���˵�ǰʣ��� " & FormatEx(cur���, 2) & " ���ڱ���ֵ " & FormatEx(Nvl(rsTmp!����ֵ, 0), 2) & "��Ҫ������", _
                    vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            End If
        End If
    End If
    CheckMoney = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetMergeCount(ByVal lngRow As Long) As Long
'���ܣ���ȡָ��һ����ҩ��ҩƷ����(��һ����ҩ����1��)
'������lngRow=һ����ҩ�Ŀ�ʼҩƷ��
    Dim lng���ID As Long, i As Long
    Dim lngCount As Long
    
    With vsAdvice
        lng���ID = Val(.TextMatrix(lngRow, COL_���ID))
        For i = lngRow To .Rows - 1
            If Val(.TextMatrix(i, COL_���ID)) = lng���ID Then
                lngCount = lngCount + 1
            Else
                Exit For
            End If
        Next
    End With
    
    GetMergeCount = lngCount
End Function

Private Function CheckAdvice() As Boolean
'���ܣ���鵱ǰ����(Ӥ��)��ҽ�������Ƿ�Ϸ�
'˵��������в��Ϸ��ĵط����ڱ���������ʾ����λ
    Dim bln�䷽�� As Boolean, bln������ As Boolean
    Dim dbl���� As Double, strMsg As String
    Dim strҩƷIDs As String, str����IDs As String
    Dim lngCount As Long, lngRow As Long, i As Long
    Dim blnSkipStock As Boolean, blnSkipTotal As Boolean
    Dim vMsg As VbMsgBoxResult, sng���� As Single
    Dim blnValid As Boolean, lngRxCount As Long
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            '�����������޸�ҩƷ�еĴ���ְ����
            If .RowData(i) <> 0 _
                And InStr(",5,6,7,", .TextMatrix(i, COL_���)) > 0 _
                And InStr(",1,2,", .TextMatrix(i, COL_EDIT)) > 0 Then
                strMsg = CheckOneDuty(.TextMatrix(i, COL_ҽ������), .TextMatrix(i, COL_����ְ��), .TextMatrix(i, COL_����ҽ��), InStr(",1,2,", mstr������) > 0 And mstr������ <> "")
                If strMsg <> "" Then
                    .Col = COL_ҽ������
                    If .TextMatrix(i, COL_���) = "7" Then
                        lngRow = .FindRow(CLng(.TextMatrix(i, COL_���ID)), i + 1)
                        If lngRow <> -1 Then .Row = lngRow
                    Else
                        .Row = i
                    End If
                    Call .ShowCell(.Row, .Col)
                    MsgBox strMsg, vbInformation, gstrSysName
                    .Refresh
                    Call vsAdvice_KeyPress(13)
                    Exit Function
                End If
            End If
            
            '��������Ϸ��Լ��
            If .RowData(i) <> 0 And Not .RowHidden(i) Then
                bln�䷽�� = RowIn�䷽��(i)
                bln������ = RowIn������(i)
                lngRow = i
                If bln�䷽�� Then '�õ��䷽�ĵ�һҩƷ��
                    lngRow = .FindRow(CStr(.RowData(i)), , COL_���ID)
                ElseIf bln������ Then '�õ�����ҽ����
                    lngRow = .FindRow(CStr(.RowData(i)), , COL_���ID)
                End If
                
                'δ���͵�ҽ����
                '------------------------------------
                If Val(.TextMatrix(i, COL_״̬)) = 1 Then
                    lngCount = lngCount + 1
                    
                    '����¼�뵥��:����:��ҩ���ѡ��Ƶ�ʵļ�ʱ,������Ŀ����¼��(Ҳ�ɲ�¼)
                    If (Val(.TextMatrix(i, COL_Ƶ������)) = 0 And InStr(",1,2,", Val(.TextMatrix(i, COL_���㷽ʽ))) > 0) _
                        Or InStr(",5,6,", .TextMatrix(i, COL_���)) > 0 Then
                        If .TextMatrix(i, COL_����) <> "" Then
                            If Not IsNumeric(.TextMatrix(i, COL_����)) Or Val(.TextMatrix(i, COL_����)) <= 0 Then
                                strMsg = "û��¼����ȷ�ĵ���������"
                                .Col = COL_����: Exit For
                            End If
                        End If
                    End If
                    
                    '����¼������:�䷽,����(ҩƷ������)
                    If Not IsNumeric(.TextMatrix(i, COL_����)) Or Val(.TextMatrix(i, COL_����)) <= 0 Then
                        If bln�䷽�� Then
                            strMsg = "û��¼����ȷ����ҩ�䷽������"
                        ElseIf InStr(",5,6,", .TextMatrix(i, COL_���)) > 0 Then
                            strMsg = "û��¼����ȷ��ҩƷ�ܸ�������"
                        Else
                            strMsg = "û��¼����ȷ��������"
                        End If
                        .Col = COL_����: Exit For
                    End If
                                        
                    '����¼��Ƶ��:����ҲҪ���,����ָ��ʹ��,���Բ�¼��ִ��ʱ��
                    If Val(.TextMatrix(i, COL_Ƶ������)) = 0 Or InStr(",5,6,", .TextMatrix(i, COL_���)) > 0 Or bln�䷽�� Then
                        If .TextMatrix(i, COL_Ƶ��) = "" Then
                            strMsg = "û��ȷ��ִ��Ƶ�ʡ�"
                            .Col = COL_Ƶ��: Exit For
                        End If
                    End If
                    
                    '����¼��ִ�п���:�Ƕ�����Ժ��ִ��ʱ(�䷽��ҩƷ�н����ж�)
                    If Val(.TextMatrix(lngRow, COL_ִ�п���ID)) = 0 Then
                        If .TextMatrix(lngRow, COL_���) = "Z" And InStr(",1,2,", Val(.TextMatrix(lngRow, COL_��������))) > 0 Then
                            If Val(.TextMatrix(lngRow, COL_��������)) = 1 Then
                                strMsg = "û��ȷ������ҽ�������ۿ��ҡ�"
                            ElseIf Val(.TextMatrix(lngRow, COL_��������)) = 2 Then
                                strMsg = "û��ȷ��סԺҽ����סԺ���ҡ�"
                            End If
                            .Col = COL_ִ�п���ID: Exit For
                        ElseIf InStr(",0,5,", .TextMatrix(lngRow, COL_ִ������)) = 0 Then
                            strMsg = "û��ȷ��ִ�п��ҡ�"
                            .Col = COL_ִ�п���ID: Exit For
                        End If
                    End If
                    If lngRow <> i And Val(.TextMatrix(i, COL_ִ�п���ID)) = 0 Then
                        If InStr(",0,5,", .TextMatrix(i, COL_ִ������)) = 0 Then
                            strMsg = "û��ȷ��ִ�п��ҡ�"
                            .Col = COL_ִ�п���ID: Exit For
                        End If
                    End If
                    
                    '����ʱ���ж�
                    If Not Check����ʱ��(.Cell(flexcpData, i, COL_����ʱ��), .Cell(flexcpData, i, COL_��ʼʱ��), False, strMsg) Then
                        .Col = COL_����ʱ��: Exit For
                    End If
                    
                    '�����������Ƽ��
                    If gintRXCount > 0 And InStr(",5,6,", .TextMatrix(i, COL_���)) > 0 _
                        And Val(.TextMatrix(i, COL_���ID)) <> Val(.TextMatrix(i - 1, COL_���ID)) Then
                        lngRxCount = GetMergeCount(i)
                        If lngRxCount > gintRXCount Then
                            strMsg = "һ���һ����ҩ����Ϊ " & lngRxCount & " ������ҩƷ����������������Ϊ " & gintRXCount & " ����"
                            .Col = COL_ҽ������: Exit For
                        End If
                    End If
                End If
                
                '�����������޸ĵ���
                '---------------------------------------------------
                If InStr(",1,2,", .TextMatrix(i, COL_EDIT)) > 0 Then
                    '��ʼʱ���ж�:ֻ��������ҽ�������ж�,��Ϊ�����ǲ�׼�޸Ŀ�ʼʱ���(�����жϱ��޸ĵ�ҽ����ʼʱ��������Ч��)
                    If .TextMatrix(i, COL_EDIT) = "1" Then
                        If Not Check��ʼʱ��(.Cell(flexcpData, i, COL_��ʼʱ��), False, strMsg) Then
                            .Col = COL_��ʼʱ��: Exit For
                        End If
                    End If
                    
                    '��ҩ;������ҩ�÷����ɼ��������ü��
                    If InStr(",5,6,", .TextMatrix(i, COL_���)) > 0 Then
                        If Val(.TextMatrix(i, COL_���ID)) = .RowData(i + 1) And Val(.TextMatrix(i + 1, COL_������ĿID)) = 0 Then
                            strMsg = "û�����ö�Ӧ�ĸ�ҩ;����"
                            .Col = COL_�÷�: Exit For
                        End If
                    End If
                    If .TextMatrix(i, COL_���) = "E" And Val(.TextMatrix(i, COL_������ĿID)) = 0 Then
                        If .RowData(i) = Val(.TextMatrix(i - 1, COL_���ID)) Then
                            If InStr(",7,E,", .TextMatrix(i - 1, COL_���)) > 0 Then
                                strMsg = "��ҩ�䷽û�����ö�Ӧ���÷���"
                            ElseIf .TextMatrix(i - 1, COL_���) = "C" Then
                                strMsg = "û�����ö�Ӧ�ı걾�ɼ�������"
                            End If
                            .Col = COL_�÷�: Exit For
                        End If
                    End If
                    
                    '�����������:����Ҫ����һ��Ƶ�����ڵ�����
                    If InStr(",5,6,", .TextMatrix(i, COL_���)) > 0 Or bln�䷽�� Then
                        If Not blnSkipTotal And .TextMatrix(i, COL_Ƶ��) <> "" Then
                            strMsg = ""
                            If bln�䷽�� Then '�ж�
                                dbl���� = CalcȱʡҩƷ����(1, 1, Val(.TextMatrix(i, COL_Ƶ�ʴ���)), Val(.TextMatrix(i, COL_Ƶ�ʼ��)), .TextMatrix(i, COL_�����λ))
                                If Val(.TextMatrix(i, COL_����)) < dbl���� Then
                                    strMsg = .TextMatrix(i, COL_ҽ������) & vbCrLf & vbCrLf & _
                                        "�ڰ�""" & .TextMatrix(i, COL_Ƶ��) & """ִ��ʱ,������Ҫ " & dbl���� & "����"
                                End If
                            ElseIf Val(.TextMatrix(i, COL_����ϵ��)) <> 0 Then
                                sng���� = Val(.TextMatrix(i, COL_����))
                                If sng���� = 0 Then sng���� = 1
                                dbl���� = CalcȱʡҩƷ����(Val(.TextMatrix(i, COL_����)), sng����, Val(.TextMatrix(i, COL_Ƶ�ʴ���)), Val(.TextMatrix(i, COL_Ƶ�ʼ��)), .TextMatrix(i, COL_�����λ), .TextMatrix(i, COL_ִ��ʱ��), Val(.TextMatrix(i, COL_����ϵ��)), Val(.TextMatrix(i, COL_�����װ)), Val(.TextMatrix(i, COL_�ɷ����)))
                                If Val(.TextMatrix(i, COL_����)) < dbl���� Then
                                    strMsg = .TextMatrix(i, COL_ҽ������) & vbCrLf & vbCrLf & _
                                        "�ڰ�ÿ�� " & .TextMatrix(i, COL_����) & .TextMatrix(i, COL_������λ) & "," & _
                                        .TextMatrix(i, COL_Ƶ��) & IIF(mbln����, ",��ҩ " & sng���� & " ��", "") & _
                                        "ִ��ʱ,������Ҫ " & dbl���� & .TextMatrix(i, COL_������λ) & "��"
                                End If
                            End If
                            If strMsg <> "" And False Then '��ʾ
                                .Row = i: .Col = COL_����: Call .ShowCell(.Row, .Col)
                                vMsg = frmMsgBox.ShowMsgBox(strMsg & "^^Ҫ������", Me)
                                If vMsg = vbNo Or vMsg = vbCancel Then
                                    If txt����.Enabled And txt����.Visible Then txt����.SetFocus
                                    Exit Function
                                ElseIf vMsg = vbIgnore Then
                                    blnSkipTotal = True
                                End If
                            End If
                        End If
                    End If
                    
                    'ҩƷ�����:ֻ����,����Ҳֻ�Ա��α༭�Ĳ��ж�
                    If (InStr(",5,6,", .TextMatrix(i, COL_���)) > 0 Or bln�䷽��) And Not blnSkipStock Then
                        strMsg = CheckStock(i)
                        If strMsg <> "" Then
                            .Row = i: .Col = COL_ҽ������: Call .ShowCell(.Row, .Col)
                            vMsg = frmMsgBox.ShowMsgBox(strMsg & "^^Ҫ������", Me)
                            If vMsg = vbNo Or vMsg = vbCancel Then
                                Exit Function
                            ElseIf vMsg = vbIgnore Then
                                blnSkipStock = True
                            End If
                        End If
                    End If
                    
                    'ִ��ʱ��Ϸ��Լ��
                    If .TextMatrix(i, COL_ִ��ʱ��) <> "" And .TextMatrix(i, COL_Ƶ��) <> "" Then
                        blnValid = ExeTimeValid(.TextMatrix(i, COL_ִ��ʱ��), Val(.TextMatrix(i, COL_Ƶ�ʴ���)), Val(.TextMatrix(i, COL_Ƶ�ʼ��)), .TextMatrix(i, COL_�����λ))
                        If Not blnValid Then
                            If .TextMatrix(i, COL_�����λ) = "��" Then
                                strMsg = COL_����ִ��
                            ElseIf .TextMatrix(i, COL_�����λ) = "��" Then
                                strMsg = COL_����ִ��
                            ElseIf .TextMatrix(i, COL_�����λ) = "Сʱ" Then
                                strMsg = COL_��ʱִ��
                            End If
                            strMsg = "¼���ִ��ʱ�䷽����ʽ����ȷ�����顣" & vbCrLf & vbCrLf & "����" & vbCrLf & strMsg
                            .Col = COL_ִ��ʱ��: Exit For
                        End If
                    End If
                End If
                                
                '���������ռ�:��������Чҽ����,��Ϊ�����ѷ��͵���δ���͵Ļ���
                If InStr(",5,6,", .TextMatrix(i, COL_���)) > 0 Then
                    '����ҩƷ������ɼ��:������Ч
                    strҩƷIDs = strҩƷIDs & "," & Val(.TextMatrix(i, COL_������ĿID))
                ElseIf Not bln�䷽�� Then
                    '���ܼ����������������ڲ�֮�估�ڲ���������Ŀ֮��
                    str����IDs = str����IDs & "," & Val(.TextMatrix(i, COL_������ĿID))
                End If
            End If
        Next
        
        '--------------------------------------------------------------------------
        '�м��˳��Ĵ�����ʾ
        If i <= .Rows - 1 Then
            .Row = i: Call .ShowCell(.Row, .Col)
            If strMsg <> "" Then
                If bln�䷽�� Then
                    strMsg = "����ҩ�䷽" & strMsg
                Else
                    strMsg = """" & .TextMatrix(i, COL_ҽ������) & """" & strMsg
                End If
                MsgBox strMsg, vbInformation, gstrSysName
                .Refresh
            End If
            Call vsAdvice_KeyPress(13)
            Exit Function
        End If
        
        '���ҩƷ�������
        If strҩƷIDs <> "" Then
            If Not Check�������(Mid(strҩƷIDs, 2)) Then Exit Function
        End If
        '���������Ŀ����
        If str����IDs <> "" Then
            If Not Check���ƻ���(Mid(str����IDs, 2)) Then Exit Function
        End If
    End With
    
    '���ñ���:��δ����ҽ��ʱ
    If lngCount > 0 Then
        If Not CheckMoney Then Exit Function
    End If
    
    CheckAdvice = True
End Function

Private Function SeekNextControl() As Boolean
'���ܣ���λ����һ������Ŀؼ���,��������������Ƿ��Զ�����һ��ҽ��
'���أ����ͨ��SetFocusǿ�ƶ�λ��,�򷵻�True
    Dim objActive As Object, objNext As Object
    Dim blnDo As Boolean, i As Long
    Dim strSkip As String
    
    Set objActive = Me.ActiveControl
    
    If Not objActive Is Nothing Then
        If TypeName(objActive) = "TextBox" Or TypeName(objActive) = "ComboBox" Then
            If objActive.Container Is fraAdvice Then
                strSkip = GetInputSkip(vsAdvice.Row)
                Set objNext = GetNextControl(objActive.TabIndex, Me, strSkip)
                If Not objNext Is Nothing Then
                    If objNext Is vsAdvice Then
                        For i = vsAdvice.Row + 1 To vsAdvice.Rows - 1
                            If Not vsAdvice.RowHidden(i) Then
                                Call AdviceChange 'ǿ�Ƹ���ҽ������
                                vsAdvice.Row = i
                                Call zlCommFun.PressKey(vbKeyTab)
                                blnDo = vsAdvice.RowData(i) <> 0 '��������������༭
                                Exit For
                            End If
                        Next
                        If i > vsAdvice.Rows - 1 Then
                            blnDo = True
                            Call tbr_ButtonClick(tbr.Buttons("����"))
                        End If
                    ElseIf strSkip <> "" And InStr(";" & strSkip & ";", objNext.Name) = 0 Then
                        blnDo = True: objNext.SetFocus
                    End If
                End If
            End If
        End If
    End If
    If Not blnDo Then
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        SeekNextControl = True
    End If
End Function

Private Function GetInputSkip(ByVal lngRow As Long) As String
'���ܣ���ȡ����ҽ�������У��س����Ӧ�����Ŀؼ�
    Dim strSkip As String, lngFind As Long
    
    With vsAdvice
        'һ����ҩ�е�ҩƷ����ʱӦ����������
        If InStr(",5,6,", .TextMatrix(lngRow, COL_���)) > 0 And .RowData(lngRow) <> 0 Then
            If Val(.TextMatrix(lngRow, COL_���ID)) = Val(.TextMatrix(lngRow - 1, COL_���ID)) Then
                '��ҩ;��,����ִ��
                If Val(.TextMatrix(lngRow, COL_���ID)) <> 0 Then
                    lngFind = .FindRow(CLng(.TextMatrix(lngRow, COL_���ID)), lngRow + 1)
                    If lngFind <> -1 Then
                        If Val(.TextMatrix(lngFind, COL_������ĿID)) <> 0 Then
                            strSkip = strSkip & ";" & Me.txt�÷�.Name
                        End If
                        If Val(.TextMatrix(lngFind, COL_ִ�п���ID)) <> 0 Then
                            strSkip = strSkip & ";" & Me.cbo����ִ��.Name
                        End If
                    End If
                End If
                'Ƶ��
                If .TextMatrix(lngRow, COL_Ƶ��) <> "" Then strSkip = strSkip & ";" & Me.txtƵ��.Name
                'ִ��ʱ��
                If .TextMatrix(lngRow, COL_ִ��ʱ��) <> "" Then strSkip = strSkip & ";" & Me.cboִ��ʱ��.Name
            End If
        End If
    End With
    GetInputSkip = Mid(strSkip, 2)
End Function

Private Sub SetBabyVisible(ByVal lng����ID As Long)
'���ܣ����ݿ�����������Ӥ��ҽ���Ƿ����ѡ��
'˵�������Ʋ���Ӥ��ҽ��
    If DeptIsWoman(lng����ID) Then
        lblӤ��.Visible = True
        cboӤ��.Visible = True
    Else
        Call zlControl.CboSetIndex(cboӤ��.Hwnd, 0)
        cboӤ��.Tag = 0
        lblӤ��.Visible = False
        cboӤ��.Visible = False
    End If
End Sub

Private Sub CalcAdviceMoney()
'���ܣ������¿�ҽ�����
'˵����ֻ�ܵ�ǰ��ʾ���Ĳ����¿�ҽ��
    Dim dblMoney As Double, i As Long
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If Not .RowHidden(i) And Val(.TextMatrix(i, COL_״̬)) = 1 Then
                dblMoney = dblMoney + Format(Val(.TextMatrix(i, COL_����)) * Val(.TextMatrix(i, COL_����)), gstrDec)
            End If
        Next
        stbThis.Panels(5).Text = "�¿�:" & FormatEx(dblMoney, 5) & "Ԫ"
    End With
End Sub

Private Sub AdviceSign()
'���ܣ���ҽ�����е���ǩ��
    Dim strSQL As String, strIDs As String, i As Long
    Dim strSource As String, strSign As String
    Dim lngǩ��ID As Long, lng֤��ID As Long
    Dim intRule As Integer
    
    If gobjESign Is Nothing Then Exit Sub
    
    '�Զ�����
    If mblnNoSave Then
        If Not CheckAdvice Then Exit Sub
        If Not SaveAdvice Then vsAdvice.SetFocus: Exit Sub
    End If
    
    '��ȡǩ��ҽ��Դ��
    intRule = ReadAdviceSignSource(1, mlng����ID, mstr�Һŵ�, strIDs, 0, False, strSource, mlngǰ��ID)
    If intRule = 0 Then Exit Sub
    If strSource = "" Then
        MsgBox "�ò���Ŀǰû�п���ǩ����ҽ����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    strSign = gobjESign.Signature(strSource, gstrDBUser, lng֤��ID)
    If strSign <> "" Then
        lngǩ��ID = zlDatabase.GetNextId("ҽ��ǩ����¼")
        strSQL = "zl_ҽ��ǩ����¼_Insert(" & lngǩ��ID & ",1," & intRule & ",'" & Replace(strSign, "'", "''") & "'," & lng֤��ID & ",'" & strIDs & "')"
        On Error GoTo errH
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        On Error GoTo 0
        
        '���¶�ȡ��ʾҽ��
        Call ReLoadAdvice(vsAdvice.RowData(vsAdvice.Row))
        mblnOK = True
        If txtҽ������.Enabled Then
            txtҽ������.SetFocus
        Else
            vsAdvice.SetFocus
        End If

        MsgBox "����ɵ���ǩ����", vbInformation, gstrSysName
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function AdviceTextChange(ByVal lngRow As Long) As Boolean
'���ܣ���ҽ����Ƭ�������ݱ仯ʱ���ж�ҽ�������ı��Ƿ�Ӧ��������֯
    Dim str��� As String, strText As String, blnDefine As Boolean
    
    With vsAdvice
        'ȷ��ҽ�����
        str��� = .TextMatrix(lngRow, COL_���)
        If str��� = "E" Then '��ҩ�䷽��һ�����
            lngRow = .FindRow(CStr(.RowData(lngRow)), , COL_���ID)
            If lngRow <> -1 Then str��� = .TextMatrix(lngRow, COL_���)
        End If
        If str��� = "7" Then str��� = "8"
                
        'ȷ���Ƿ���
        blnDefine = Not mrsDefine Is Nothing And Not mobjVBA Is Nothing
        If blnDefine Then
            mrsDefine.Filter = "�������='" & str��� & "'"
            If mrsDefine.EOF Then
                blnDefine = False
            ElseIf Trim(Nvl(mrsDefine!ҽ������)) = "" Then
                blnDefine = False
            End If
        End If
        If blnDefine Then strText = mrsDefine!ҽ������
        
        '������ݱ䶯
        If blnDefine Then '�����ֶβ��ݻ���Թ�������Ĳ���
            If IsDate(txt��ʼʱ��.Text) And txt��ʼʱ��.Tag <> "" And InStr(strText, "[��ʼʱ��]") > 0 Then
                AdviceTextChange = True: Exit Function
            End If
            If cboҽ������.Tag <> "" And InStr(strText, "[ҽ������]") > 0 Then
                AdviceTextChange = True: Exit Function
            End If
            If cmdƵ��.Tag <> "" And txtƵ��.Tag <> "" Then
                If InStr(strText, "[����Ƶ��]") > 0 Or InStr(strText, "[Ӣ��Ƶ��]") > 0 Then
                    AdviceTextChange = True: Exit Function
                End If
            End If
            If cboִ��ʱ��.Tag <> "" And InStr(strText, "[ִ��ʱ��]") > 0 Then
                AdviceTextChange = True: Exit Function
            End If
            If (IsNumeric(txt����.Text) Or txt����.Text = "") And txt����.Tag <> "" And InStr(strText, "[����]") > 0 Then
                AdviceTextChange = True: Exit Function
            End If
            If IsNumeric(txt����.Text) And txt����.Tag <> "" And InStr(strText, "[����]") > 0 Then
                AdviceTextChange = True: Exit Function
            End If
        End If
        
        Select Case str��� '��ͬ�������
        Case "5", "6" '������ҩ
            If Not blnDefine Then
                
            Else
                '[������][ͨ����][��Ʒ��][Ӣ����][���][����]��������޸�����ҩƷʱ�仯
                If Val(cmd�÷�.Tag) <> 0 And txt�÷�.Tag <> "" And InStr(strText, "[��ҩ;��]") > 0 Then
                    AdviceTextChange = True: Exit Function
                End If
            End If
        Case "8" '��ҩ�䷽
            If Not blnDefine Then
                If IsNumeric(txt����.Text) And txt����.Tag <> "" Then AdviceTextChange = True: Exit Function
                If cmdƵ��.Tag <> "" And txtƵ��.Tag <> "" Then AdviceTextChange = True: Exit Function
                If Val(cmd�÷�.Tag) <> 0 And txt�÷�.Tag <> "" Then AdviceTextChange = True: Exit Function
            Else
                '[�䷽���][�巨]��������޸������䷽ʱ�仯
                If IsNumeric(txt����.Text) And txt����.Tag <> "" And InStr(strText, "[����]") > 0 Then
                    AdviceTextChange = True: Exit Function
                End If
                If Val(cmd�÷�.Tag) <> 0 And txt�÷�.Tag <> "" And InStr(strText, "[�÷�]") > 0 Then
                    AdviceTextChange = True: Exit Function
                End If
            End If
        Case "C" '����
            If Not blnDefine Then
                If Val(cmd�÷�.Tag) <> 0 And txt�÷�.Tag <> "" Then AdviceTextChange = True: Exit Function
            Else
                '[������Ŀ][����걾]��������޸�������Ŀʱ�仯
                If Val(cmd�÷�.Tag) <> 0 And txt�÷�.Tag <> "" And InStr(strText, "[�ɼ�����]") > 0 Then
                    AdviceTextChange = True: Exit Function
                End If
            End If
        Case "D" '���
            If Not blnDefine Then
                
            Else
                '[�����Ŀ][��鲿λ]��������޸�������Ŀʱ�仯
            End If
        Case "F" '����
            If Not blnDefine Then
                If IsDate(txt��ʼʱ��.Text) And txt��ʼʱ��.Tag <> "" Then AdviceTextChange = True: Exit Function
            Else
                '[��Ҫ����][��������][������]��������޸�������Ŀʱ�仯
            End If
        Case Else '����
            If Not blnDefine Then
                
            Else
                '[������Ŀ]��������޸�������Ŀʱ�仯
            End If
        End Select
    End With
End Function

Private Function AdviceTextMake(ByVal lngRow As Long) As String
'���ܣ���ȡҽ�������ı�
'������lngRow=����ҽ�����ݵĿɼ���
    Dim rsTmp As New ADODB.Recordset
    Dim blnDefine As Boolean, str��� As String
    Dim strText As String, strSQL As String
    Dim strField As String, intƵ�ʷ�Χ As Integer
    Dim i As Long, k As Long
    
    Dim str��ҩ As String, str�巨 As String
    Dim str���� As String, str���� As String
    Dim str���� As String, str�걾 As String
    Dim str��λ As String
    
    On Error GoTo errH
    
    With vsAdvice
        'ȷ��ҽ�����
        str��� = .TextMatrix(lngRow, COL_���)
        If str��� = "E" Then '��ҩ�䷽��һ�����
            k = .FindRow(CStr(.RowData(lngRow)), , COL_���ID)
            If k <> -1 Then str��� = .TextMatrix(k, COL_���)
        End If
        If str��� = "7" Then str��� = "8"
                
        'ȷ���Ƿ���
        blnDefine = Not mrsDefine Is Nothing And Not mobjVBA Is Nothing
        If blnDefine Then
            mrsDefine.Filter = "�������='" & str��� & "'"
            If mrsDefine.EOF Then
                blnDefine = False
            ElseIf Trim(Nvl(mrsDefine!ҽ������)) = "" Then
                blnDefine = False
            End If
        End If
        
ReDoDefault: '���ڰ����幫ʽ����ʧ�ܣ����°�ȱʡ���������֯
        strText = ""
        If blnDefine Then strText = mrsDefine!ҽ������
        
        '����ҽ������
        Select Case str���
        Case "C" '����-------------------------------------------------------------
            str���� = "": str�걾 = ""
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_���ID)) = .RowData(lngRow) Then
                    str���� = .TextMatrix(i, COL_ҽ������) & "," & str����
                    str�걾 = .TextMatrix(i, COL_�걾��λ)
                Else
                    Exit For
                End If
            Next
            If str���� = "" Then '�ϵķ�ʽ
                str���� = .TextMatrix(lngRow, COL_����)
            Else
                str���� = Left(str����, Len(str����) - 1)
            End If
            
            If Not blnDefine Then
                strText = str���� & IIF(str�걾 <> "", "(" & str�걾 & ")", "")
            Else
                If InStr(strText, "[������Ŀ]") > 0 Then
                    strField = str����
                    strText = Replace(strText, "[������Ŀ]", """" & strField & """")
                End If
                If InStr(strText, "[����걾]") > 0 Then
                    strField = str�걾
                    strText = Replace(strText, "[����걾]", """" & strField & """")
                End If
                If InStr(strText, "[�ɼ�����]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_�÷�)
                    strText = Replace(strText, "[�ɼ�����]", """" & strField & """")
                End If
            End If
        Case "D" '���-------------------------------------------------------------
            str��λ = ""
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, COL_���ID)) = .RowData(lngRow) Then
                    If .TextMatrix(i, COL_�걾��λ) <> "" Then
                        str��λ = str��λ & "," & .TextMatrix(i, COL_�걾��λ)
                    End If
                Else
                    Exit For
                End If
            Next
            str��λ = Mid(str��λ, 2) '��������Ŀ�Ĳ�λ
            If str��λ = "" Then '���������Ŀ�Ĳ�λ
                str��λ = .TextMatrix(lngRow, COL_�걾��λ)
            End If
            
            If Not blnDefine Then
                strText = .TextMatrix(lngRow, COL_����) & IIF(str��λ <> "", "(" & str��λ & ")", "")
            Else
                If InStr(strText, "[�����Ŀ]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_����)
                    strText = Replace(strText, "[�����Ŀ]", """" & strField & """")
                End If
                If InStr(strText, "[��鲿λ]") > 0 Then
                    strField = str��λ
                    strText = Replace(strText, "[��鲿λ]", """" & strField & """")
                End If
            End If
        Case "F" '����-------------------------------------------------------------
            str���� = "": str���� = ""
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, COL_���ID)) = .RowData(lngRow) Then
                    If .TextMatrix(i, COL_���) = "G" Then
                        str���� = .TextMatrix(i, COL_ҽ������)
                    Else
                        str���� = str���� & "," & .TextMatrix(i, COL_ҽ������)
                    End If
                Else
                    Exit For
                End If
            Next
            str���� = Mid(str����, 2)
            
            If Not blnDefine Then
                strText = Format(.Cell(flexcpData, lngRow, COL_��ʼʱ��), "MM��dd��HH:mm")
                If str���� <> "" Then
                    strText = strText & IIF(str���� <> "", " �� " & str���� & " ���� ", " �� ")
                End If
                strText = strText & .TextMatrix(lngRow, COL_����)
                If str���� <> "" Then
                    strText = strText & " �� " & str����
                End If
            Else
                If InStr(strText, "[��Ҫ����]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_����)
                    strText = Replace(strText, "[��Ҫ����]", """" & strField & """")
                End If
                If InStr(strText, "[��������]") > 0 Then
                    strField = str����
                    strText = Replace(strText, "[��������]", """" & strField & """")
                End If
                If InStr(strText, "[������]") > 0 Then
                    strField = str����
                    strText = Replace(strText, "[������]", """" & strField & """")
                End If
            End If
        Case "8" '��ҩ�䷽---------------------------------------------------------
            str��ҩ = "": str�巨 = ""
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_���ID)) = .RowData(lngRow) Then
                    If .TextMatrix(i, COL_���) = "7" Then
                        str��ҩ = RTrim(.TextMatrix(i, COL_ҽ������) & _
                            " " & .TextMatrix(i, COL_����) & .TextMatrix(i, COL_������λ) & _
                            " " & .TextMatrix(i, COL_ҽ������)) & "," & str��ҩ
                    ElseIf .TextMatrix(i, COL_���) = "E" Then
                        str�巨 = .TextMatrix(i, COL_ҽ������)
                    End If
                Else
                    Exit For
                End If
            Next
            If str��ҩ <> "" Then
                str��ҩ = Mid(str��ҩ, 1, Len(str��ҩ) - 1)
            End If
            If Not blnDefine Then
                '���ֺ���˿ո����ı����л��Զ�����
                strText = "��ҩ" & .TextMatrix(lngRow, COL_����) & "��," & _
                    .TextMatrix(lngRow, COL_Ƶ��) & "," & str�巨 & "," & _
                    .TextMatrix(lngRow, COL_�÷�) & ":" & str��ҩ
            Else
                If InStr(strText, "[����]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_����)
                    strText = Replace(strText, "[����]", """" & strField & """")
                End If
                If InStr(strText, "[�䷽���]") > 0 Then
                    strField = str��ҩ
                    strText = Replace(strText, "[�䷽���]", """" & strField & """")
                End If
                If InStr(strText, "[�÷�]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_�÷�)
                    strText = Replace(strText, "[�÷�]", """" & strField & """")
                End If
                If InStr(strText, "[�巨]") > 0 Then
                    strField = str�巨
                    strText = Replace(strText, "[�巨]", """" & strField & """")
                End If
            End If
        Case "5", "6" '����ҩ���г�ҩ---------------------------------------------
            If Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)) <> 0 Then
                '����:0-����,1-Ӣ����,3-��Ʒ��
                strSQL = "Select Nvl(B.����,A.����) as ����,A.���,A.����,B.����" & _
                    " From �շ���ĿĿ¼ A,�շ���Ŀ���� B Where A.ID=B.�շ�ϸĿID(+) And A.ID=[1] Order by B.����,B.����"
                Set rsTmp = New ADODB.Recordset '���Filter
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)))
            ElseIf blnDefine Then
                '����:0-����,1-Ӣ����
                strSQL = "Select Nvl(B.����,A.����) as ����,Null as ���,Null as ����,B.����" & _
                    " From ������ĿĿ¼ A,������Ŀ���� B Where A.ID=B.������ĿID(+) And A.ID=[1] Order by B.����,B.����"
                Set rsTmp = New ADODB.Recordset '���Filter
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngRow, COL_������ĿID)))
            End If
            If Not blnDefine Then
                strText = .TextMatrix(lngRow, COL_�걾��λ)
                If Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)) <> 0 Then
                    If strText = "" Then
                        If gbln��Ʒ�� Then rsTmp.Filter = "����=3"
                        If rsTmp.EOF Then rsTmp.Filter = 0
                        strText = rsTmp!����
                    End If
                    If Not IsNull(rsTmp!����) Then
                        strText = strText & "(" & rsTmp!���� & ")"
                    End If
                    If Not IsNull(rsTmp!���) Then
                        strText = strText & " " & rsTmp!���
                    End If
                Else
                    If strText = "" Then
                        strText = .TextMatrix(lngRow, COL_����)
                    End If
                End If
            Else
                If InStr(strText, "[������]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_�걾��λ)
                    If strField = "" Then
                        If gbln��Ʒ�� Then rsTmp.Filter = "����=3"
                        If rsTmp.EOF Then rsTmp.Filter = 0
                        strField = rsTmp!����
                    End If
                    strText = Replace(strText, "[������]", """" & strField & """")
                End If
                If InStr(strText, "[ͨ����]") > 0 Then
                    rsTmp.Filter = 0
                    strField = rsTmp!����
                    strText = Replace(strText, "[ͨ����]", """" & strField & """")
                End If
                If InStr(strText, "[��Ʒ��]") > 0 Then
                    rsTmp.Filter = "����=3"
                    If rsTmp.EOF Then rsTmp.Filter = 0
                    strField = rsTmp!����
                    strText = Replace(strText, "[��Ʒ��]", """" & strField & """")
                End If
                If InStr(strText, "[Ӣ����]") > 0 Then
                    rsTmp.Filter = "����=2"
                    If rsTmp.EOF Then rsTmp.Filter = 0
                    strField = rsTmp!����
                    strText = Replace(strText, "[Ӣ����]", """" & strField & """")
                End If
                If InStr(strText, "[���]") > 0 Then
                    If rsTmp.EOF Then rsTmp.Filter = 0
                    strField = Nvl(rsTmp!���)
                    strText = Replace(strText, "[���]", """" & strField & """")
                End If
                If InStr(strText, "[����]") > 0 Then
                    If rsTmp.EOF Then rsTmp.Filter = 0
                    strField = Nvl(rsTmp!����)
                    strText = Replace(strText, "[����]", """" & strField & """")
                End If
                If InStr(strText, "[��ҩ;��]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_�÷�)
                    strText = Replace(strText, "[��ҩ;��]", """" & strField & """")
                End If
            End If
        Case Else '�����������-----------------------------------------------------
            If Not blnDefine Then
                strText = .TextMatrix(lngRow, COL_����)
            Else
                If InStr(strText, "[������Ŀ]") > 0 Then
                    strField = .TextMatrix(lngRow, COL_����)
                    strText = Replace(strText, "[������Ŀ]", """" & strField & """")
                End If
            End If
            '����ҽ��������ʾ
            If .TextMatrix(lngRow, COL_���) = "Z" And Val(.TextMatrix(lngRow, COL_��������)) = 4 Then
                strText = "������" & strText & "������"
            End If
        End Select
        
        '�����ֶλ���Թ���������ֶ�-------------------------------------------
        If blnDefine Then
            If InStr(strText, "[��ʼʱ��]") > 0 Then
                strField = .Cell(flexcpData, lngRow, COL_��ʼʱ��)
                strText = Replace(strText, "[��ʼʱ��]", """" & strField & """")
            End If
            If InStr(strText, "[ҽ������]") > 0 Then
                strField = .Cell(flexcpData, lngRow, COL_ҽ������)
                strText = Replace(strText, "[ҽ������]", """" & strField & """")
            End If
            If InStr(strText, "[����Ƶ��]") > 0 Then
                strField = .TextMatrix(lngRow, COL_Ƶ��)
                strText = Replace(strText, "[����Ƶ��]", """" & strField & """")
            End If
            If InStr(strText, "[Ӣ��Ƶ��]") > 0 Then
                strField = ""
                If .TextMatrix(lngRow, COL_Ƶ��) <> "" Then
                    intƵ�ʷ�Χ = GetƵ�ʷ�Χ(lngRow)
                    strSQL = "Select Ӣ������ From ����Ƶ����Ŀ Where ����=[1] And ���÷�Χ=[2]"
                    Set rsTmp = New ADODB.Recordset '���Filter
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .TextMatrix(lngRow, COL_Ƶ��), intƵ�ʷ�Χ)
                    If Not rsTmp.EOF Then strField = Nvl(rsTmp!Ӣ������)
                End If
                strText = Replace(strText, "[Ӣ��Ƶ��]", """" & strField & """")
            End If
            If InStr(strText, "[����]") > 0 Then
                strField = ""
                If .TextMatrix(lngRow, COL_����) <> "" Then
                    strField = .TextMatrix(lngRow, COL_����) & .TextMatrix(lngRow, COL_������λ)
                End If
                strText = Replace(strText, "[����]", """" & strField & """")
            End If
            If InStr(strText, "[����]") > 0 Then
                strField = ""
                If .TextMatrix(lngRow, COL_����) <> "" Then
                    strField = .TextMatrix(lngRow, COL_����) & .TextMatrix(lngRow, COL_������λ)
                End If
                strText = Replace(strText, "[����]", """" & strField & """")
            End If
            If InStr(strText, "[ִ��ʱ��]") > 0 Then
                strField = .TextMatrix(lngRow, COL_ִ��ʱ��)
                strText = Replace(strText, "[ִ��ʱ��]", """" & strField & """")
            End If
        End If
                
        '����ҽ������
        If blnDefine Then
            On Error Resume Next
            strText = mobjVBA.Eval(strText)
            If mobjVBA.Error.Number <> 0 Then
                Err.Clear: On Error GoTo errH
                blnDefine = False: GoTo ReDoDefault
            End If
        End If
    End With
    AdviceTextMake = strText
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function AdviceCheckWarn(ByVal lngCmd As Long, Optional ByVal lngRow As Long) As Long
'���ܣ�����Passϵͳ�ж�ҽ�����к�����ҩ������ع���
'������lngCmd=
'        0-�������PASS�˵�״̬
'        1-�����Զ����,2-�ύ�Զ����,3-�ֹ��������
'        6-��ҩ����,12-��ҩ�о�,22-����״̬/����ʷ����(�༭)
'      lngRow=��ǰҩƷҽ�����кţ�lngCmd=0,6ʱ��Ҫ
'���أ�������˷��ص���߼���ʾֵ,Ϊ-1,-2,-3��ʾû�н������
'      ���PASS�˵�ʱ������>=0��ʾ���Ե����˵�
'˵������ҩ��飺�漰�����µ�����(������ִ��)����δֹͣ�ĳ���
'      ��ҩ�о����漰�������е�ҽ��(���Դ����ݿ��,Ҫ�󱣴�)
'      ��ҩ���棺Ӧ����ҩ����֮����е���(�о���ֵ)
    Dim rsTmp As New ADODB.Recordset
    Dim strҩƷ As String, str�÷� As String, strƵ�� As String
    Dim lngMaxWarn As Long, strOld As String
    Dim strSQL As String, blnDo As Boolean
    Dim lngCount As Long, curDate As Date
    Dim arrLevel(0 To 4) As Long
    Dim i As Long, k As Long
    
    lngMaxWarn = -1
    AdviceCheckWarn = lngMaxWarn
    
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
    If mlng����ID <> mlngPassPati Then
        strSQL = "Select ����ID,Count(Distinct Trunc(�Ǽ�ʱ��)) as ������� From ���˹Һż�¼ Where ����ID=[1] Group by ����ID"
        strSQL = "Select D.�������,A.����,A.�Ա�,A.��������," & _
            " C.���� as ������,C.���� as ������,E.��� as ҽ����,E.���� as ҽ����" & _
            " From ������Ϣ A,���˹Һż�¼ B,���ű� C,(" & strSQL & ") D,��Ա�� E" & _
            " Where A.����ID=B.����ID And B.ִ�в���ID=C.ID And A.����ID=D.����ID" & _
            " And B.ִ����=E.����(+) And A.����ID=[1] And B.NO=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mstr�Һŵ�)
        If rsTmp.EOF Then Screen.MousePointer = 0: Exit Function
    
        Call PassSetPatientInfo(mlng����ID, rsTmp!�������, rsTmp!����, Nvl(rsTmp!�Ա�), Format(rsTmp!��������, "yyyy-MM-dd"), "", "", _
            rsTmp!������ & "/" & rsTmp!������, IIF(Not IsNull(rsTmp!ҽ����), Nvl(rsTmp!ҽ����) & "/" & Nvl(rsTmp!ҽ����), ""), "")
        mlngPassPati = mlng����ID
    End If
    
    'PASS�Զ���˵����
    '-------------------------------------------------------------
    If lngCmd = 0 Then
        With vsAdvice
            If .RowData(lngRow) <> 0 And InStr(",5,6,7,", .TextMatrix(lngRow, COL_���)) > 0 And Val(.TextMatrix(lngRow, COL_�շ�ϸĿID)) <> 0 Then
                'ȡҩƷ����
                strҩƷ = .TextMatrix(lngRow, COL_ҽ������)
                If InStr(strҩƷ, " ") > 0 Then strҩƷ = Left(strҩƷ, InStr(strҩƷ, " ") - 1)
                If InStr(strҩƷ, "(") > 0 Then strҩƷ = Left(strҩƷ, InStr(strҩƷ, "(") - 1)
                'ȡҩƷ��ҩ;��
                str�÷� = ""
                k = .FindRow(CLng(.TextMatrix(lngRow, COL_���ID)), lngRow + 1)
                If k <> -1 Then str�÷� = .TextMatrix(k, COL_ҽ������)
                
                '�����ѯҩƷ��Ϣ
                Call PassSetQueryDrug(.TextMatrix(lngRow, COL_�շ�ϸĿID), strҩƷ, .TextMatrix(lngRow, COL_������λ), str�÷�)
                    
                '���ò˵�����״̬
                Call SetPassMenuState
                
                AdviceCheckWarn = 1 '��ʾ���Ե����˵�
            End If
        End With
        Screen.MousePointer = 0: Exit Function
    End If
    
    '����ʷ/����״̬�༭
    '-------------------------------------------------------------
    If lngCmd = 22 Then
        'lngCmd=21-ֻ��,22-��ǿ�Ʊ༭,23-ǿ�Ʊ༭
        If PassDoCommand(lngCmd) = 2 Then
            '�������ֵΪ2��ʾ"����ʷ/����״̬�༭"�������仯����Ҫ�����Զ����
            lngCmd = 2 'תΪ�Զ��������,����ִ��
        Else
            Screen.MousePointer = 0: Exit Function
        End If
    End If
    
    '���벡��ҽ����Ϣ
    '-------------------------------------------------------------
    With vsAdvice
        If lngCmd = 6 Then
            Call PassSetWarnDrug(.RowData(lngRow)) '��ҩ����(�Ѿ����ҽ��Ψһ��)
        Else
            '��ҩ��˻���ҩ�о�
            lngCount = 0
            curDate = zlDatabase.Currentdate
            strҩƷ = "": str�÷� = "": strƵ�� = ""
            For i = .FixedRows To .Rows - 1
                blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", .TextMatrix(i, COL_���)) > 0 _
                    And Val(.TextMatrix(i, COL_Ӥ��)) = cboӤ��.ListIndex And Val(.TextMatrix(i, COL_�շ�ϸĿID)) <> 0
                blnDo = blnDo And (lngCmd = 12 Or Format(.Cell(flexcpData, i, COL_����ʱ��), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd"))
                If blnDo Then
                    'ȡҩƷ����
                    strҩƷ = .TextMatrix(i, COL_ҽ������)
                    If InStr(strҩƷ, " ") > 0 Then strҩƷ = Left(strҩƷ, InStr(strҩƷ, " ") - 1)
                    If InStr(strҩƷ, "(") > 0 Then strҩƷ = Left(strҩƷ, InStr(strҩƷ, "(") - 1)
                    
                    'ȡҩƷ��ҩ;��
                    If Val(.TextMatrix(i, COL_���ID)) <> Val(.TextMatrix(i - 1, COL_���ID)) Then str�÷� = "" 'һ����ҩ���ظ�ȡ
                    If str�÷� = "" Then
                        k = .FindRow(CLng(.TextMatrix(i, COL_���ID)), i + 1)
                        If k <> -1 Then str�÷� = .TextMatrix(k, COL_ҽ������)
                    End If
                    
                    'ȡ��ҩƵ��(��/��),��Ϊ������������
                    If Val(.TextMatrix(i, COL_���ID)) <> Val(.TextMatrix(i - 1, COL_���ID)) Then strƵ�� = "" 'һ����ҩ���ظ�ȡ
                    If strƵ�� = "" Then
                        If .TextMatrix(i, COL_�����λ) = "��" Then
                            strƵ�� = .TextMatrix(i, COL_Ƶ�ʴ���) & "/" & .TextMatrix(i, COL_Ƶ�ʼ��)
                        ElseIf .TextMatrix(i, COL_�����λ) = "��" Then
                            strƵ�� = .TextMatrix(i, COL_Ƶ�ʴ���) & "/7"
                        ElseIf .TextMatrix(i, COL_�����λ) = "Сʱ" Then
                            If Val(.TextMatrix(i, COL_Ƶ�ʼ��)) <= 24 Then
                                strƵ�� = Format(24 / Val(.TextMatrix(i, COL_Ƶ�ʼ��)) * Val(.TextMatrix(i, COL_Ƶ�ʴ���)), "0") & "/1"
                            Else
                                strƵ�� = Val(.TextMatrix(i, COL_Ƶ�ʴ���)) & "/" & Format(Val(.TextMatrix(i, COL_Ƶ�ʼ��)) / 24, "0")
                            End If
                        End If
                    End If
                    
                    '����ҽ����Ϣ
                    Call PassSetRecipeInfo(.RowData(i), .TextMatrix(i, COL_�շ�ϸĿID), strҩƷ, _
                        .TextMatrix(i, COL_����), .TextMatrix(i, COL_������λ), strƵ��, _
                        Format(.Cell(flexcpData, i, COL_����ʱ��), "yyyy-MM-dd"), "", str�÷�, _
                        .TextMatrix(i, COL_���ID), 1, UserInfo.��� & "/" & UserInfo.����)
                    lngCount = lngCount + 1
                End If
            Next
            '�޿�����ҩƷ
            If (lngCmd = 1 Or lngCmd = 2 Or lngCmd = 3) And lngCount = 0 Then
                Screen.MousePointer = 0: Exit Function
            End If
        End If
    End With
    
    'ִ����Ӧ������
    '-------------------------------------------------------------
    Call PassDoCommand(lngCmd)
    
    '��ȡҽ�������,����д��ʾ��
    '-------------------------------------------------------------
    If lngCmd = 1 Or lngCmd = 2 Or lngCmd = 3 Then
        '����ֵ˳��0-����,1-�Ƶ�,2-���,3-�ڵ�,4-�ȵ�
        '��ʾ��˳��0-����,1-�Ƶ�,4-�ȵ�,2-���,3-�ڵ�(��ΪPASS������ԭ��)
        arrLevel(0) = 0: arrLevel(1) = 1: arrLevel(2) = 3: arrLevel(3) = 4: arrLevel(4) = 2
        With vsAdvice
            For i = .FixedRows To .Rows - 1
                blnDo = .RowData(i) <> 0 And InStr(",5,6,7,", .TextMatrix(i, COL_���)) > 0 _
                    And Val(.TextMatrix(i, COL_Ӥ��)) = cboӤ��.ListIndex And Val(.TextMatrix(i, COL_�շ�ϸĿID)) <> 0
                blnDo = blnDo And Format(.Cell(flexcpData, i, COL_����ʱ��), "yyyy-MM-dd") = Format(curDate, "yyyy-MM-dd")
                If blnDo Then
                    k = PassGetWarn(.RowData(i))
                    strOld = .Cell(flexcpData, i, COL_��ʾ)

                    '���þ�ʾ��
                    If k >= 0 And k <= 4 Then
                        .Cell(flexcpData, i, COL_��ʾ) = CStr(k)
                        Set .Cell(flexcpPicture, i, COL_��ʾ) = imgPass.ListImages(k + 1).Picture
                    Else
                        .Cell(flexcpData, i, COL_��ʾ) = ""
                        Set .Cell(flexcpPicture, i, COL_��ʾ) = Nothing
                    End If
                    
                    '���������仯,�Ա��������ݿ�
                    If CStr(.Cell(flexcpData, i, COL_��ʾ)) <> strOld Then
                        .Cell(flexcpData, i, COL_���) = 1
                        mblnNoSave = True '���Ϊδ����
                    End If
                                        
                    '��¼��߼���ʾֵ
                    If k >= 0 Then
                        If lngMaxWarn >= 0 Then
                            If arrLevel(k) > arrLevel(lngMaxWarn) Then
                                lngMaxWarn = k
                            End If
                        Else
                            lngMaxWarn = k
                        End If
                    End If
                End If
            Next
        End With
    End If
    
    '���������
    AdviceCheckWarn = lngMaxWarn
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
    If Button = 2 And gblnPass And InStr(mstrPrivs, "������ҩ���") > 0 Then
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
    'ϵͳ����
    mnuPassItem(14).Enabled = PassGetState("SYS-SET") = 1
    '��ҩ�о�
    mnuPassItem(16).Enabled = PassGetState("DISQUISITION") = 1
    '����:�о�ʾֵ(��Ϊ��),�Ҵ���0-����
    mnuPassItem(18).Enabled = Val(vsAdvice.Cell(flexcpData, vsAdvice.Row, COL_��ʾ)) > 0
    '���
    'mnuPassItem(19).Enabled = PassGetState("") = 1
    
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
    Case 14 'ϵͳ����
        Call PassDoCommand(11)
    Case 16 '��ҩ�о�
        Call AdviceCheckWarn(12)
    Case 18 '����
        Call AdviceCheckWarn(6, vsAdvice.Row)
    Case 19 '���
        Call AdviceCheckWarn(3)
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
