VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "Mscomm32.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#9.0#0"; "zlIDKind.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCharge 
   AutoRedraw      =   -1  'True
   Caption         =   "�����շѴ���"
   ClientHeight    =   8148
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   11280
   BeginProperty Font 
      Name            =   "����"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCharge.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8148
   ScaleWidth      =   11280
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboTemp 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   0
      TabIndex        =   94
      TabStop         =   0   'False
      Text            =   "cbo������"
      ToolTipText     =   "֧���������ͱ���Զ�ƥ��"
      Top             =   20000
      Width           =   2145
   End
   Begin VSFlex8Ctl.VSFlexGrid vsInvoice 
      Height          =   375
      Left            =   -135
      TabIndex        =   93
      TabStop         =   0   'False
      Top             =   7650
      Width           =   11310
      _cx             =   19950
      _cy             =   661
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483633
      ForeColor       =   -2147483640
      BackColorFixed  =   8421504
      ForeColorFixed  =   16777215
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483633
      GridColor       =   12632256
      GridColorFixed  =   -2147483633
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483633
      FocusRect       =   3
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   8
      FixedRows       =   0
      FixedCols       =   1
      RowHeightMin    =   360
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmCharge.frx":08CA
      ScrollTrack     =   -1  'True
      ScrollBars      =   0
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
      ExplorerBar     =   3
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
   Begin VB.Frame fra�˷�ժҪ 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   15
      TabIndex        =   80
      Top             =   5160
      Visible         =   0   'False
      Width           =   7335
      Begin VB.TextBox txt�˷�ժҪ 
         Height          =   360
         Left            =   1140
         MaxLength       =   100
         TabIndex        =   16
         Top             =   0
         Width           =   5820
      End
      Begin VB.Label lblժҪ 
         Caption         =   "�˷�ժҪ"
         Height          =   225
         Left            =   135
         TabIndex        =   15
         Top             =   60
         Width           =   1035
      End
   End
   Begin VB.Frame fraSubBill 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   15
      TabIndex        =   69
      Top             =   5160
      Visible         =   0   'False
      Width           =   11865
      Begin VB.Label lblAmount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�䷽�ϼ�:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4440
         TabIndex        =   79
         Top             =   45
         Width           =   1155
      End
      Begin VB.Label lblDuty 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������רҵְ��:"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   74
         Top             =   45
         Width           =   1800
      End
      Begin VB.Label lblSubӦ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ӧ��:0.00"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   7245
         TabIndex        =   71
         Top             =   45
         Width           =   1185
      End
      Begin VB.Label lblSubʵ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ʵ��:0.00"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   9345
         TabIndex        =   70
         Top             =   45
         Width           =   1185
      End
   End
   Begin VB.Frame fraBill 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   15
      TabIndex        =   66
      Top             =   1830
      Width           =   11820
      Begin VB.CommandButton cmdDelBill 
         Caption         =   "ɾ��(&D)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   10850
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "ɾ����ǰ����(ALT+D)"
         Top             =   30
         Width           =   960
      End
      Begin VB.CommandButton cmdAddBill 
         Caption         =   "����(&A)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   9870
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "����һ�ŵ���(F12)"
         Top             =   30
         Width           =   960
      End
      Begin MSComctlLib.TabStrip tbsBill 
         Height          =   705
         Left            =   30
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   15
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   1249
         TabWidthStyle   =   2
         TabFixedWidth   =   2117
         TabFixedHeight  =   616
         HotTracking     =   -1  'True
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "����&1"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.ComboBox cbo�������� 
      Height          =   360
      Left            =   1200
      TabIndex        =   7
      Text            =   "cbo��������"
      Top             =   1410
      Width           =   2010
   End
   Begin VB.Frame fraTitle 
      Height          =   1080
      Left            =   0
      TabIndex        =   45
      ToolTipText     =   "���:F6"
      Top             =   -120
      Width           =   11880
      Begin VB.CommandButton cmdSaveWholeSet 
         Caption         =   "����Ϊ�����շ���Ŀ(&W)"
         Height          =   375
         Left            =   6630
         TabIndex        =   85
         Top             =   195
         Width           =   2715
      End
      Begin VB.CommandButton cmdSelWholeSet 
         Caption         =   "����(&T)"
         Height          =   375
         Left            =   5520
         TabIndex        =   84
         TabStop         =   0   'False
         ToolTipText     =   " "
         Top             =   195
         Width           =   1080
      End
      Begin VB.CommandButton cmdYB 
         Caption         =   "ҽ��"
         Height          =   375
         Left            =   1080
         TabIndex        =   78
         TabStop         =   0   'False
         ToolTipText     =   "�ȼ���F6"
         Top             =   660
         Width           =   720
      End
      Begin VB.CommandButton cmdIDCard 
         Caption         =   "ҽ�ƿ�(&K)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9480
         TabIndex        =   72
         ToolTipText     =   "�ȼ���F10"
         Top             =   195
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.CommandButton cmdRegist 
         Caption         =   "�Һ�(&E)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10725
         TabIndex        =   42
         ToolTipText     =   "�ȼ���F3"
         Top             =   195
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.TextBox txtRePrint 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2500
         MaxLength       =   8
         TabIndex        =   35
         Top             =   667
         Width           =   1065
      End
      Begin VB.TextBox txtModi 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   4250
         MaxLength       =   8
         TabIndex        =   37
         Top             =   667
         Width           =   1065
      End
      Begin VB.CommandButton cmd�䷽ 
         Caption         =   "�䷽(&R)"
         Height          =   375
         Left            =   80
         TabIndex        =   33
         TabStop         =   0   'False
         ToolTipText     =   "�ȼ���F11"
         Top             =   660
         Width           =   1000
      End
      Begin VB.TextBox txtInvoice 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   7680
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   660
         Width           =   1545
      End
      Begin VB.ComboBox cboNO 
         ForeColor       =   &H80000007&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   9975
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "��λ:F9,���ݺų��Ȳ���ʱ�Զ����㳤��"
         Top             =   660
         Width           =   1350
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   11370
         Style           =   1  'Graphical
         TabIndex        =   40
         TabStop         =   0   'False
         ToolTipText     =   "�ȼ�:F8"
         Top             =   645
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.CheckBox chkCancel 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   11370
         Style           =   1  'Graphical
         TabIndex        =   41
         TabStop         =   0   'False
         ToolTipText     =   "�ȼ�:F8"
         Top             =   645
         Width           =   435
      End
      Begin VB.TextBox txtIn 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   5985
         MaxLength       =   8
         TabIndex        =   39
         ToolTipText     =   "�����еĵ����и�����Ϣ,��Ӱ�����е���"
         Top             =   667
         Width           =   1065
      End
      Begin VB.TextBox txtMCInvoice 
         ForeColor       =   &H000000FF&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   675
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000014&
         X1              =   7110
         X2              =   7110
         Y1              =   630
         Y2              =   1050
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000010&
         X1              =   7100
         X2              =   7100
         Y1              =   630
         Y2              =   1050
      End
      Begin VB.Label lblRePrint 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��(&P)"
         Height          =   240
         Left            =   1900
         TabIndex        =   34
         Top             =   727
         Width           =   600
      End
      Begin VB.Label lblIn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��(&I)"
         Height          =   240
         Left            =   5350
         TabIndex        =   38
         Top             =   727
         Width           =   600
      End
      Begin VB.Label lblModi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��(&M)"
         Height          =   240
         Left            =   3650
         TabIndex        =   36
         Top             =   727
         Width           =   600
      End
      Begin VB.Label lblFormat 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C000C0&
         Height          =   240
         Left            =   9360
         TabIndex        =   67
         Top             =   255
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label lblFact 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ʊ��"
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   7200
         TabIndex        =   43
         Top             =   720
         Width           =   480
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         X1              =   15
         X2              =   38015
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   0
         X2              =   38000
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label lblFlag 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   390
         Left            =   11370
         TabIndex        =   53
         Top             =   645
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "�����շѵ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   135
         TabIndex        =   48
         ToolTipText     =   "���:F6"
         Top             =   195
         Width           =   1875
      End
      Begin VB.Label lblNO 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "���ݺ�"
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   9250
         TabIndex        =   46
         Top             =   720
         Width           =   720
      End
      Begin VB.Label lblCorp 
         Caption         =   "������λ:"
         Height          =   255
         Left            =   5280
         TabIndex        =   73
         Top             =   960
         Visible         =   0   'False
         Width           =   5895
      End
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   47
      Top             =   7785
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   12
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2625
            MinWidth        =   882
            Picture         =   "frmCharge.frx":0994
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9991
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Visible         =   0   'False
            Object.Width           =   0
            MinWidth        =   88
            Object.Tag             =   "���ڼ��ʻ��շѸ����ʻ���ʾ"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Visible         =   0   'False
            Object.Width           =   370
            MinWidth        =   88
            Object.Tag             =   "�����շ�Ԥ����ʾ"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   360
            MinWidth        =   71
            Key             =   "MedicareType"
            Object.ToolTipText     =   "ҽ������"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   952
            MinWidth        =   952
            Picture         =   "frmCharge.frx":1228
            Key             =   "Drugstore"
            Object.ToolTipText     =   "ҩ������"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   952
            MinWidth        =   952
            Key             =   "PatiSource"
            Object.ToolTipText     =   "������Դ"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   882
            MinWidth        =   882
            Picture         =   "frmCharge.frx":1542
            Key             =   "Calc"
            Object.ToolTipText     =   "������:ALT+?"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmCharge.frx":1C1C
            Key             =   "PY"
            Object.ToolTipText     =   "ƴ��(F7)"
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmCharge.frx":2256
            Key             =   "WB"
            Object.ToolTipText     =   "���(F7)"
         EndProperty
         BeginProperty Panel11 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1101
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel12 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1101
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraInfo 
      Height          =   990
      Left            =   0
      TabIndex        =   44
      Top             =   840
      Width           =   11880
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   390
         Left            =   555
         TabIndex        =   92
         Top             =   180
         Width           =   630
         _ExtentX        =   1101
         _ExtentY        =   699
         Appearance      =   2
         IDKindStr       =   "��|����|0;ҽ|ҽ����|0;��|���֤��|0;IC|IC����|1;��|�����|0;��|���￨|0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   12
         FontName        =   "����"
         IDKind          =   -1
         ShowPropertySet =   -1  'True
         NotContainFastKey=   "F1;CTRL+F1;F2;F3;CTRL+F4;F5;F6;F7;CTRL+F7;F8;F9;F10;F11;F12;CTRL+F12;CTRL+S;CTRL+A;CTRL+R;CTRL+D;CTRL+Q;ESC;ALT+?"
         MustSelectItems =   "����,���￨"
         BackColor       =   -2147483633
      End
      Begin VB.ComboBox cbo���䵥λ 
         Height          =   360
         Left            =   5750
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   180
         Width           =   580
      End
      Begin VB.TextBox txt����� 
         BackColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   9650
         Locked          =   -1  'True
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   180
         Width           =   2145
      End
      Begin VB.CheckBox chk���� 
         Alignment       =   1  'Right Justify
         Caption         =   "����"
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   8040
         TabIndex        =   9
         Top             =   630
         Visible         =   0   'False
         Width           =   790
      End
      Begin VB.ComboBox cboҽ�Ƹ��� 
         Height          =   360
         Left            =   6360
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   180
         Width           =   2505
      End
      Begin VB.TextBox txtPatient 
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   1200
         MaxLength       =   100
         TabIndex        =   2
         ToolTipText     =   "��λ:F6,����:-����ID,*�����,+סԺ��,.�Һŵ���,����:*2536��ʾ������Ų���"
         Top             =   180
         Width           =   1470
      End
      Begin VB.ComboBox cboSex 
         Height          =   360
         Left            =   3200
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   180
         Width           =   1095
      End
      Begin VB.TextBox txt���� 
         Height          =   360
         IMEMode         =   2  'OFF
         Left            =   4920
         MaxLength       =   20
         TabIndex        =   4
         Top             =   180
         Width           =   795
      End
      Begin VB.ComboBox cbo�ѱ� 
         Height          =   360
         Left            =   3765
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   570
         Width           =   1575
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         Caption         =   "����"
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   8910
         TabIndex        =   76
         Top             =   630
         Width           =   2880
      End
      Begin VB.Label lbl����� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����"
         Height          =   240
         Left            =   8910
         TabIndex        =   68
         Top             =   240
         Width           =   720
      End
      Begin VB.Label lbl��̬�ѱ� 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   5520
         TabIndex        =   64
         Top             =   600
         Width           =   2370
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   240
         Left            =   100
         TabIndex        =   10
         Top             =   630
         Width           =   960
      End
      Begin VB.Label lblPatient 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   75
         TabIndex        =   52
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblSex 
         AutoSize        =   -1  'True
         Caption         =   "�Ա�"
         Height          =   240
         Left            =   2680
         TabIndex        =   51
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblOld 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   240
         Left            =   4395
         TabIndex        =   50
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lbl�ѱ� 
         AutoSize        =   -1  'True
         Caption         =   "�ѱ�"
         Height          =   240
         Left            =   3240
         TabIndex        =   49
         Top             =   630
         Width           =   480
      End
   End
   Begin VB.PictureBox picAppend 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2280
      Left            =   0
      ScaleHeight     =   2280
      ScaleWidth      =   11280
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   5505
      Width           =   11280
      Begin VB.CommandButton cmdԤ���� 
         Caption         =   "Ԥ����(&V)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   10305
         TabIndex        =   29
         ToolTipText     =   "�ȼ���F5"
         Top             =   540
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   10305
         TabIndex        =   30
         ToolTipText     =   "�ȼ�F2,�Ҽ���������Ϊ���۵�(��CTRL+S)"
         Top             =   975
         Width           =   1440
      End
      Begin VSFlex8Ctl.VSFlexGrid vsBalance 
         Height          =   1770
         Left            =   5415
         TabIndex        =   91
         Top             =   495
         Width           =   2445
         _cx             =   4313
         _cy             =   3122
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
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
         GridColor       =   -2147483630
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   2
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   2
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   400
         RowHeightMax    =   400
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmCharge.frx":2890
         ScrollTrack     =   0   'False
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
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00C0C0C0&
         Caption         =   "����շ�(&F)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   10320
         TabIndex        =   31
         ToolTipText     =   "�ȼ���Alt+F"
         Top             =   1860
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ȡ��(&C)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   10320
         TabIndex        =   32
         ToolTipText     =   "�ȼ�:Esc"
         Top             =   1410
         Width           =   1440
      End
      Begin VB.TextBox txtTmp 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   240
         IMEMode         =   3  'DISABLE
         Left            =   6510
         MaxLength       =   10
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   570
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Frame fraAppend 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   0
         TabIndex        =   56
         ToolTipText     =   "���:F6"
         Top             =   -90
         Width           =   11880
         Begin VB.ComboBox cboBaby 
            Height          =   360
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   165
            Width           =   1800
         End
         Begin VB.ComboBox cbo���㷽ʽ 
            Height          =   360
            Left            =   3720
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   165
            Width           =   1680
         End
         Begin VB.CheckBox chk�Ӱ� 
            Caption         =   "�Ӱ�(&L)"
            Height          =   270
            Left            =   80
            TabIndex        =   17
            Top             =   210
            Width           =   1170
         End
         Begin VB.ComboBox cbo������ 
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   6615
            TabIndex        =   23
            Text            =   "cbo������"
            ToolTipText     =   "֧���������ͱ���Զ�ƥ��"
            Top             =   165
            Width           =   2145
         End
         Begin MSMask.MaskEdBox txtDate 
            Height          =   360
            Left            =   9390
            TabIndex        =   24
            Top             =   165
            Width           =   2400
            _ExtentX        =   4233
            _ExtentY        =   635
            _Version        =   393216
            AutoTab         =   -1  'True
            HideSelection   =   0   'False
            MaxLength       =   19
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "yyyy-MM-dd hh:mm:ss"
            Mask            =   "####-##-## ##:##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label lblBaby 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Ӥ����(&B)"
            Height          =   240
            Left            =   1320
            TabIndex        =   18
            Top             =   225
            Width           =   1080
         End
         Begin VB.Label lbl������ 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "������(&W)"
            Height          =   240
            Left            =   5505
            TabIndex        =   22
            Top             =   225
            Width           =   1080
         End
         Begin VB.Label lblDate 
            AutoSize        =   -1  'True
            Caption         =   "ʱ��"
            Height          =   240
            Left            =   8880
            TabIndex        =   57
            Top             =   225
            Width           =   480
         End
         Begin VB.Label lbl���㷽ʽ 
            AutoSize        =   -1  'True
            Caption         =   "���㷽ʽ(&X)"
            Height          =   240
            Left            =   2400
            TabIndex        =   20
            Top             =   225
            Width           =   1320
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshMoney 
         Height          =   1770
         Left            =   15
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   510
         Width           =   2775
         _ExtentX        =   4890
         _ExtentY        =   3112
         _Version        =   393216
         Rows            =   6
         Cols            =   4
         FixedCols       =   0
         RowHeightMin    =   280
         BackColorBkg    =   -2147483643
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   2
         MergeCells      =   1
         AllowUserResizing=   1
         FormatString    =   "^���|^��Ŀ     |^    ���|^     �ϼ�"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
      End
      Begin VB.Frame fraStat 
         Height          =   1905
         Left            =   2865
         TabIndex        =   58
         Top             =   375
         Width           =   2490
         Begin VB.TextBox txt�ϼ� 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   450
            Left            =   735
            Locked          =   -1  'True
            TabIndex        =   26
            TabStop         =   0   'False
            Text            =   "0.00"
            ToolTipText     =   "�����շ�ʱδ�ɿ�ݵ�ʵ�ս��ϼ�"
            Top             =   810
            Width           =   1650
         End
         Begin VB.TextBox txtӦ�� 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   450
            Left            =   750
            Locked          =   -1  'True
            TabIndex        =   25
            TabStop         =   0   'False
            Text            =   "0.00"
            ToolTipText     =   "�����շ�ʱδ�ɿ�ݵ�Ӧ�ս��ϼ�"
            Top             =   285
            Width           =   1650
         End
         Begin VB.TextBox txt�ۼ� 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   450
            Left            =   750
            Locked          =   -1  'True
            TabIndex        =   27
            TabStop         =   0   'False
            Text            =   "0.00"
            Top             =   1350
            Width           =   1650
         End
         Begin VB.Label lbl�ϼ� 
            AutoSize        =   -1  'True
            Caption         =   "ʵ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   15.6
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   315
            Left            =   60
            TabIndex        =   61
            Top             =   885
            Width           =   660
         End
         Begin VB.Label lblӦ�� 
            AutoSize        =   -1  'True
            Caption         =   "Ӧ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   15.6
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   60
            TabIndex        =   60
            Top             =   345
            Width           =   690
         End
         Begin VB.Label lbl�ۼ� 
            AutoSize        =   -1  'True
            Caption         =   "�ۼ�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   15.6
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   60
            TabIndex        =   59
            Top             =   1410
            Width           =   690
         End
      End
      Begin MSComctlLib.ImageList imgPati 
         Left            =   4875
         Top             =   1875
         _ExtentX        =   995
         _ExtentY        =   995
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCharge.frx":28DE
               Key             =   "InPati"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCharge.frx":31B8
               Key             =   "OutPati"
            EndProperty
         EndProperty
      End
      Begin VB.Frame fraUpBillShow 
         Caption         =   "���ŵ���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1710
         Left            =   5325
         TabIndex        =   81
         Top             =   540
         Visible         =   0   'False
         Width           =   1920
         Begin VB.TextBox txtPreMoney 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   450
            Left            =   105
            Locked          =   -1  'True
            TabIndex        =   83
            TabStop         =   0   'False
            ToolTipText     =   "���ŵ��ݽ��"
            Top             =   1005
            Width           =   1710
         End
         Begin VB.TextBox txtPreNO 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   450
            Left            =   105
            Locked          =   -1  'True
            TabIndex        =   82
            TabStop         =   0   'False
            ToolTipText     =   "���ŵ��ݺ�"
            Top             =   495
            Width           =   1710
         End
      End
      Begin VB.Frame fra�ɿ� 
         Height          =   1905
         Left            =   7860
         TabIndex        =   86
         Top             =   375
         Width           =   2325
         Begin VB.TextBox txtԤ����� 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   450
            Left            =   795
            TabIndex        =   89
            Text            =   "0.00"
            Top             =   240
            Width           =   1395
         End
         Begin VB.TextBox txtӦ�� 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   450
            Left            =   795
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   87
            TabStop         =   0   'False
            Text            =   "0.00"
            ToolTipText     =   "�����շ�ʱ,ָӦ�ɺϼ�"
            Top             =   765
            Width           =   1395
         End
         Begin VB.Label lblԤ����� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ԥ���"
            ForeColor       =   &H00808080&
            Height          =   240
            Left            =   60
            TabIndex        =   90
            Top             =   360
            Width           =   720
         End
         Begin VB.Label lblӦ�� 
            AutoSize        =   -1  'True
            Caption         =   "��  ��"
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   45
            TabIndex        =   88
            Top             =   870
            Width           =   720
         End
      End
      Begin VB.Label lblSeek 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ڰ�ť��λ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   10200
         TabIndex        =   65
         Top             =   585
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ϼ�:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   225
         TabIndex        =   62
         Top             =   585
         Visible         =   0   'False
         Width           =   945
      End
   End
   Begin MSCommLib.MSComm com 
      Left            =   120
      Top             =   75
      _ExtentX        =   995
      _ExtentY        =   995
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin ZL9BillEdit.BillEdit Bill 
      Height          =   2925
      Left            =   0
      TabIndex        =   14
      Top             =   2220
      Width           =   11865
      _ExtentX        =   20934
      _ExtentY        =   5165
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      TxtCheck        =   -1  'True
      TxtCheck        =   -1  'True
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
      Active          =   -1  'True
      Cols            =   2
      RowHeight0      =   360
      RowHeightMin    =   360
      ColWidth0       =   1008
      BackColor       =   -2147483643
      BackColorBkg    =   -2147483643
      BackColorSel    =   10249818
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      ForeColorSel    =   -2147483634
      GridColor       =   -2147483630
      ColAlignment0   =   9
      ListIndex       =   -1
      CellBackColor   =   -2147483643
      cboStyle        =   0
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ϼ�:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   0
      TabIndex        =   54
      Top             =   0
      Width           =   945
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Visible         =   0   'False
      Begin VB.Menu mnuFileSavePrice 
         Caption         =   "����Ϊ���۵�(&S)"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmCharge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Private Const M_MONEY_ROWS = 6 '���½���Ŀ�б����ʾ����
'����������������������������������������������������������������������������������������������������������������������������������������
'��ڲ�����
Public mbytInFun As Byte '0-�շ�,1-����,2-�������
Public mbytInState As Byte '0-ִ��(���޸�),1-���,2-����,3-�˷�(�շѡ����ʲ����˷�),4-�����շ�;5-�쳣��������
Public mstrInNO As String '�����ĵ��ݺ�(�鿴���������޸ģ��˷ѣ�����,�����շ�ʱ)
Public mblnCopyBill As Boolean '�Ƿ��Զ����Ʋ����ĵ���
Public mblnNOMoved As Boolean '�����ĵ����Ƿ��ں����ݱ���
Public mstrTime As String '�����������ݵĵǼ�ʱ��
Public mblnDelete As Boolean '�Ƿ����˷ѵ���(����)
Public mbytBilling As Byte 'mbytInFun=2ʱ��0-��������,1-���ʻ���,2-�������
Public mstrPrivs As String
Public mlngModul As Long
Public mlngFirstID As Long '��¼���޸ĵ��ݵ�һҩƷ�е�ִ�в���ID,�����շ�
Public mstrFirstWin As String '��¼���޸ĵ��ݵ�һҩƷ�еķ�ҩ����,�����շ�
Public mbln�˷��쳣 As Boolean '�쳣��������

'��������������˲�����ر���
Public mlng����ID As Long
Public mlng��ҳID As Long
Public mlngUnitID As Long '��ǰ���ʲ���,Ϊ0ʱ��ʾ���в���
Public mlngDeptID As Long '��ǰ���ʿ���,Ϊ0ʱ��ʾ���п���
Public mbln���� As Boolean '33744
Public mlng����ҽ�� As Long
Public mstr���ת��ʱ�� As String

'��Ϣ��ض������
Public WithEvents mobjMsgModule As clsMipModule
Attribute mobjMsgModule.VB_VarHelpID = -1

'----------------------------------------------------------------------------------------------------------------------------------------
Private mstrӦ������㷽ʽ As String    '33722
Private mint�˷ѻص���ӡ As Integer '�˷ѻص���ӡ��ʽ 0-����ӡ,1-�Զ���ӡ,2-ѡ���Ƿ��ӡ
Private mblnSaveAsPrice As Boolean '����ҽ�����շ�ʱ�Ƿ񱣴�Ϊ���۵�
Private mintReturnMode As Integer   '�����˷�ʱ,ȫ�˽��ý��㷽ʽʱ�ָ���ʼ�Ľ��㷽ʽ
Private mblnNotValied As Boolean '������Ч��ʧЧ����
Private mblnNotClick As Boolean
Private mstrBalance As String
Private mblnHaveExcuteData As Boolean '�Ƿ�ҽ���Ƽ��д�������:60735
'����������������������������������������������������������������������������������������������������������������������������������������
Private mblnErrBill As Boolean
'���ݶ���
Private mrs���㷽ʽ As ADODB.Recordset
Private mrsWork As ADODB.Recordset      '�����ϰ��ҩ��
Private mrsClass As ADODB.Recordset     '���ݲ�����ȡ�ĵ�ǰ���õ��շ����
Private mrsUnit As ADODB.Recordset      '��ѡ���ִ�п���
Private mrs�������� As ADODB.Recordset  '��ѡ�Ŀ�������
Private mrs������ As ADODB.Recordset    '����ҽ���ͻ�ʿ����
Private mrsInfo As ADODB.Recordset      '������Ϣ
Private mrsWarn As ADODB.Recordset      '���������߼�¼��
Private mrs�ѱ� As ADODB.Recordset      '���зѱ����ÿ���
Private mrs�������� As ADODB.Recordset  '���з�������
Private mrs��ҩ���� As ADODB.Recordset  '��ҩ�����嵥,�����ж�ҩ���Ƿ�ָ���˷�ҩ����
'�������
Private mobjBill As ExpenseBill '���õ��ݶ���
Private mcolMoneys As BillInComes  '���е��ݵ�������Ŀ���ܼ���
Private mobjBillDetail As BillDetail '���ݵ��շ�ϸĿ����
Private mobjBillIncome As BillInCome '�շ�ϸĿ��������Ŀ����
Private mobjDetail As Detail '�������շ�ϸĿ����
Private mcolDetails As Details '�������շ�ϸĿ����
Private mrs�շѶ��� As ADODB.Recordset '�շѶ��� :����:33634

Private mlngShareUseID As Long '������������ID
Private mstrUseType As String 'ʹ�����
Private mintInvoiceFormat As Integer  '��ӡ�ķ�Ʊ��ʽ,��Ʊ��ʽ���
Private mintOldInvoiceFormat As Integer '�ɷ�Ʊ��ʽ��ӡ
Private mblnStartFactUseType As Boolean   '�Ƿ�������ʹ�����
Private mintInvoicePrint As Integer  '0-����ӡ;1-�Զ���ӡ;2-��ʾ��ӡ
Private mblnFirst As Boolean
Private Enum BillColType       '���ݿؼ���������
    CheckBox = -1
    Text_UnModify = 0
    CommandButton = 1
    Date = 2
    ComboBox = 3
    Text = 4
    UnFocus = 5
End Enum
Private Enum BillCol
    �� = 0
    ��� = 1
    ��Ŀ = 2
    ��Ʒ�� = 3
    �������� = 4
    ��� = 5
    ��λ = 6
    ���� = 7
    ���� = 8
    ���� = 9
    Ӧ�ս�� = 10
    ʵ�ս�� = 11
    ִ�п��� = 12
    ��־ = 13
    ҽ����� = 14
    ���� = 15
    ִ�п���ID = 16
End Enum

'�������
Private mintPage As Integer '��ǰ�ǵڼ��ŵ���
Private mstrWarn As String '�Ѿ���������ѡ����������
Private mcolStock1 As Collection '��Ÿ���ҩƷ�ⷿ�ĳ����鷽ʽ
Private mcolStock2 As Collection '��Ÿ������Ŀ�ĳ����鷽ʽ

Private mlngPreRow As Long '��ǰ�к�,�����иı�ʱ�ж�
Private mlngҩƷ���ID As Long '��ǰ���ݲ�����ҩƷ������ID
Private mlng�������ID As Long '��ǰ���ݲ���������������ID

Private mbln����ְ���� As Boolean     '�Ƿ���д���ְ����
Private mbln����������� As Boolean     '�Ƿ���д����������
Private mbln�����޶��� As Boolean     '�Ƿ���д����޶���

Private mcolBalance As Collection '��¼���ŵ��ݱ��ս���ԭʼֵ���޸�ֵ
Private mcolRquareBalance As Collection '���˺�:���������ѿ��Ľ�������

Private mblnHotKey As Boolean '�ֹ�����ʱ,�Ƿ�Ű��˱����ȼ�
Private mbln���ϼ� As Boolean
Private mstrCardNO As String '���￨���۵��ݺ�
Private mstr���ʽ As String 'ȱʡҽ�Ƹ��ʽ
Private mbytBillSource As Byte   '������Դ:1-����;2-סԺ;3-����(���￨�ȶ�����շ�);4-���

Private mstrPrePati As String  '��һ���շѲ���
Private mlngPrePati As Long     '��һ���շѲ���ID
Private mstrPreDoctor As String  '��¼ǰһ������

Private mstr���� As String, mstr�ɴ� As String, mstr�д� As String '��¼���ﲡ�������շѵĴ��ڷ���
Private mlng��ҩ�� As Long, mlng��ҩ�� As Long, mlng��ҩ�� As Long '��¼���ﲡ�������շѵ�ҩ������
Private mblnNewRow As Boolean '��ʾ�Ƿ���Ϊ����
Private mlng����ID As Long '�շ�Ʊ�ݵ���������
Private mbln������۸� As Boolean     '���޸ĺ͵��뵥��ʱ,���÷ѱ�ʱ������۸�,����ʱ����,����Ҳ������

Private mblnF2Save As Boolean   '�Ƿ�F2����
Private mblnValid As Boolean '�Ƿ���Ϊ���㶪ʧ
Private mblnDo As Boolean           '���ƼӰ�_click�¼��Ƿ񼤻�
Private mblnDoing As Boolean        '�����Ƿ����ڶ�������Ϣ
Private mblnEnterCell As Boolean    '�����Ƿ񼤻�EnterCell�¼�
Private mblnDrop As Boolean         '��KeyDown���ж�cbo�����˵�ǰ�Ƿ񵯳�
Private mblnCboClick As Boolean      '�����cbo��keypress�¼������˵����б��API����:sendmessage,�����ͣ��cbo��,����һ���ַ�,�ƿ�����򰴻س���,
'                                    cbo��ֵ�ᱣ������,�����ᴥ��click�¼�,������Ҫ��validate�¼��е���click�¼�
'�շѴ�ͬһ���˲��˵����ۼƽ��
Private mcurBillӦ�� As Currency
Private mcurBillʵ�� As Currency
Private mcurBillӦ�� As Currency
Private mdbl�ɿ� As Double, mdbl�Ҳ� As Double
Private mbln�������� As Boolean     'ȷ����ǰ�����Ƿ���������:44944
Private mintBillNO As Integer '���˵�ǰ�������˼��ŵ���
Private mintMoneyRow As Integer '��ǰ��ʾ���ķ�Ŀ��
Private mblnLoad As Boolean
Private mblnOne As Boolean '�Ƿ�ֻ��һ�������շ����
Private marrColData() As Integer '��ǰ���ݱ༭����ӳ��
Private mblnPrint As Boolean '�շ�ʱ�Ƿ��ӡƱ��,������:���ز��������Ƿ��ӡ,����Ϊ0�Ƿ��ӡ
Private mblnSelect As Boolean '���ڿ����շ�ϸĿ�����Ƿ��������б�ѡ���ѡ����

Private Const STR_HEAD = "��,450,4;���,750,1;��Ŀ,2175,1;��Ʒ��,2000,1;��������,0,0;���,1105,1;��λ,520,4;����,520,1;����,570,1;����,1055,7;" & _
    "Ӧ�ս��,1030,7;ʵ�ս��,1080,7;ִ�п���,1255,1;��־,520,4;ҽ�����,0,0;����,520,1;ִ�п���ID,0,1"

'ҽ�����
Private mintInsure As Integer
Private mstrYBPati As String 'New:�ջ�0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);8����ID
Private mstr�����ʻ� As String '�Ƿ񽫸����ʻ����õ��շѿ���
Private mcur������� As Currency '��ǰ���˸����ʻ����
Private mcur����͸֧ As Currency '�����ʻ�����͸֧���
Private mblnYB�������� As Boolean 'ҽ���Ƿ�֧�ֽ�������,�����˷�ʱ�ж�
Private mstrYBBill As String 'ҽ�����������շѵĵ��ݺ�
Private mlng�������  As Long '�����շ�ʱ��Ч
Private mblnOneCard As Boolean      '�Ƿ�������һ��ͨ�ӿ�
Private mrsOneCard As ADODB.Recordset
Private mrsDelInvoice As ADODB.Recordset

'��ǰ���������ҽ��֧�ֲ���
Private Type TYPE_MedicarePAR
    ��������ҽ����Ŀ As Boolean
    �����շѴ�Ϊ���۵� As Boolean
    �����ѽɿ���� As Boolean    '27536
    ������봫����ϸ As Boolean
    ҽ���ӿڴ�ӡƱ�� As Boolean
    ҽ��ȷ���������� As Boolean
    �൥��һ�ν��� As Boolean
    ���������շ� As Boolean
    ����Ԥ���� As Boolean
    �൥���շ� As Boolean
    �ֱҴ��� As Boolean
    ʵʱ��� As Boolean
    ���Ը� As Boolean
    ȫ�Ը� As Boolean
    blnOnlyBjYb As Boolean '���ؽ�֧�ֱ���ҽ��:���˺�
    �˷Ѻ��ӡ�ص� As Boolean '
    �൥�ݵ�һ�ν��� As Boolean
    ҽ������Ʊ��  As Boolean        'Ԥ����ʱ��Ч
End Type
Private MCPAR As TYPE_MedicarePAR

Private Type TYPE_Original
    ʵ�պϼ� As Currency    '�������,��¼�޸ĵ���ʱ��ԭ����ʵ�ս��ϼ�
    Ӧ�ɽ�� As Currency    '�շ�,��¼�޸ĵ���ʱ��Ӧ�ɽ��
    ��Ԥ���� As Currency    '�շ�,��¼�޸ĵ���ʱ��ԭʼԤ�������
    ����ID As Long          '�˷�,��¼ԭ���ݽ���ID
End Type
Private Original As TYPE_Original
Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private mblnAutoChangePati As Boolean '��ǰ����Ժģʽ���Զ��л�����Ժģʽ��

Private Type Ty_ModulePara
    blnסԺ���������շ� As Boolean    'סԺ������ȫ�����շ�
    '�Ժ���չ
End Type
Private mTy_Para As Ty_ModulePara
Private mobjBaseItem As Object
Private Enum Pan
    C2��ʾ��Ϣ = 2
    C3�����ʻ� = 3
    C4Ԥ����Ϣ = 4
    C5ҽ������ = 5
End Enum
Private mblnSaveData As Boolean  '�Ƿ����ݱ���ɹ�
Private mblnKeyReturn As Boolean '�Ƿ��˻س���
Private mrsErrBlance As ADODB.Recordset  '�쳣���ݵĽ�����Ϣ
Private Type Ty_DelFee  '�˷����
      strNos  As String          '��ǰ�˷ѵ����漰�ĵ��ݺ�,�ö��ŷ���
      rsBlance As ADODB.Recordset
      blnSingleBalance As Boolean '�����㷽ʽ
      dblCurDelMoney As Double '��ǰ�˿�ϼ�
      bln������ȫ�� As Boolean
End Type
Private mTyDelFee As Ty_DelFee
Private mblnNotClearLedDisplay As Boolean   '�������ʾ
'-----------------------------------------------------------------------------------
'���㿨���
Private mstrPassWord As String
Private mlngPreBrushCard As Long  '�ϴ�ˢ���Ŀ����ID
Private Const VK_RETURN = &HD
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private WithEvents mFrmBalanceWin   As frmChargePayMentWin
Attribute mFrmBalanceWin.VB_VarHelpID = -1
'-----------------------------------------------------------------------------------
'���ݱ������
Private mstrModiNOs As String, mstrSaveNos As String
Private mcllPayDrugAndStuff As Collection   '�����Զ����Ϻͷ�ҩ
Private mCllWindows As Collection
Private mobjDrugPacker  As Object ' �Զ���ҩ��(���·�ҩ����)
Private mblnDrugPacker As Boolean
Private mobjDrugMachine As Object '�Զ���ҩ��(�£�
Private mblnDrugMachine As Boolean
Private mblnClearBlance As Boolean '�Ƿ����������Ϣ
Private mlngCardTypeID As Long   '��ǰ��ȡ������Ϣˢ�Ŀ����ID 56615
Private mblnOlnyԤ�� As Boolean '��ʹ��Ԥ��68177
Private mstrҩƷ�۸�ȼ� As String, mstr���ļ۸�ȼ� As String
Public mstr��ͨ�۸�ȼ� As String
Private mblnSetControl As Boolean
Private WithEvents mobjBrushCheck As clsBrushCardInput
Attribute mobjBrushCheck.VB_VarHelpID = -1
Private mobjCard As New Card
Private mbln����ˢ�� As Boolean

Private Sub Bill_cboKeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    '���˺� ����:27378 ����:2010-01-27 13:35:37
    If Bill.cboStyle = DropOlnyDown Then Exit Sub
    
    Select Case Bill.TextMatrix(0, Bill.Col)
        Case "ִ�п���"
            If Bill.ListIndex <> -1 Then Exit Sub
        Case "��ҩҩ��"
            If Bill.ListIndex <> -1 Then Exit Sub
        Case Else
        Exit Sub
    End Select
    
    If mobjBill.Pages(mintPage).Details.Count < Bill.Row Then
        Exit Sub
     End If
    With mobjBill.Pages(mintPage).Details(Bill.Row)
        If InStr(",4,5,6,7,", .�շ����) > 0 Then
            If mrsWork Is Nothing Then Exit Sub
            If mrsWork.State <> 1 Then Exit Sub
            If zlSelectDept(Me, mlngModul, Bill.cboObj, mrsWork, Bill.CboText, True, , False) = False Then Exit Sub
        Else
            If mrsUnit Is Nothing Then Exit Sub
            If mrsUnit.State <> 1 Then Exit Sub
            If zlSelectDept(Me, mlngModul, Bill.cboObj, mrsUnit, Bill.CboText, True, , False) = False Then Exit Sub
        End If
    End With
    Exit Sub
End Sub

Private Sub cbo���䵥λ_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
       
End Sub

Private Sub cboҽ�Ƹ���_Click()
    On Error GoTo errHandler
    If mbytInState <> 0 Then Exit Sub
    If gintPriceGradeStartType < 2 Then Exit Sub
    
    If mrsInfo.State = adStateOpen Then
        If gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, Val(Nvl(mrsInfo!����ID)), Val(Nvl(mrsInfo!��ҳID)), zlStr.NeedName(cboҽ�Ƹ���.Text), mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ�۸�ȼ�) = False Then Exit Sub
    Else
        If gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, 0, 0, zlStr.NeedName(cboҽ�Ƹ���.Text), mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ�۸�ȼ�) = False Then Exit Sub
    End If
    
    If mbln������۸� Then Exit Sub
    If CheckBillsEmpty() Then Exit Sub
    
    '��Ҫ����Ԥ����
    If cmdԤ����.Visible Then
        Call InitBalanceGrid
        cmdԤ����.TabStop = True
        cmdOK.Enabled = False
    End If
    
    'ȫ�����¼���۸�
    Call CalcMoneys
    Call ShowDetails
    Call ShowMoney
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdSaveWholeSet_Click()
    Dim i As Long, strItems As String, lngִ�п���ID As Long
    Dim rsTemp As ADODB.Recordset, dbl�۸� As Double
    Dim strSQL As String
    '����Ϊ�����շ���Ŀ
    '����:27327
    Err = 0: On Error Resume Next
    If mobjBaseItem Is Nothing Then
        Set mobjBaseItem = CreateObject("zl9BaseItem.clsBaseItem")
    End If
    If mobjBaseItem Is Nothing Then Exit Sub
    'OpenEditWholeSetItem(ByVal frmMain As Object, ByVal cnOracle As ADODB.Connection,
    '      ByVal lngSys As Long, ByVal lngModule As Long, ByVal strPrivs As String, ByVal strItems As String) As Boolean
    'strItems:���,����,�շ�ϸĿID,����,����,ִ�п���|���,����,�շ�ϸĿID,����,����,ִ�п���|��
    Err = 0: On Error GoTo Errhand:
   If mbytInState = 1 Then
        '�鿴
         strSQL = _
        " Select Nvl(A.�۸񸸺�,A.���) as ���,A.�շ����,A.��������,A.�շ�ϸĿID,A.ִ�в���ID," & _
        "       ��   Avg(Nvl(A.����,1)) as ����, Avg(A.����) ����, Sum(A.��׼����) as ����,B.ִ�п���, B.�Ƿ���,M.��������" & _
        " From " & IIf(mblnNOMoved, zlGetFullFieldsTable("������ü�¼"), "������ü�¼  A") & ",�շ���ĿĿ¼ B,�������� M" & _
        " Where  A.��¼״̬  IN(0,1,3)  And A.NO=[1]  And A.��¼����=[2]   " & _
        "               And a.�շ�ϸĿID=b.ID And a.�շ�ϸĿID=M.����ID(+) " & _
                        IIf(mstrTime <> "", " And A.�Ǽ�ʱ��=[3]", "") & _
        "  Group by Nvl(A.�۸񸸺�,A.���),A.�շ����,A.�շ�ϸĿID,A.��������,A.ִ�в���id,B.ִ�п���,B.�Ƿ���,M.��������" & _
        " Order by ���"
        If mstrTime <> "" Then
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrInNO, 2, CDate(mstrTime))
        Else
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrInNO, 2)
        End If
        With rsTemp
            Do While Not .EOF
                 '0-����ȷ,1-���˿���,2-���˲���,3-����Ա����,4-ָ������,5-Ժ��ִ��(Ԥ��,������δ��),6-�����˿���
                If InStr(1, ",4,5,6,7,", "," & Nvl(!�շ����)) > 0 Then
                    lngִ�п���ID = 0
                ElseIf InStr(1, ",0,4", Val(Nvl(!ִ�п���))) > 0 Then
                    lngִ�п���ID = Val(Nvl(!ִ�в���ID))
                Else
                    lngִ�п���ID = 0
                End If
                dbl�۸� = 0
                If Val(Nvl(!�Ƿ���)) = 1 Then
                    If InStr(1, "5,6,7", Nvl(!�շ����)) > 0 Or (Nvl(!�շ����) = "4" And Val(Nvl(!��������)) = 1) Then
                        'ҩƷ,����������Ϊ��ȱʡ�۸�,���Բ�����(ͨ��������)
                        dbl�۸� = 0
                    Else
                        dbl�۸� = Val(Nvl(!����))
                    End If
                End If
                strItems = strItems & "|" & Val(Nvl(!���)) & "," & Val(Nvl(!��������)) & "," & Val(Nvl(!�շ�ϸĿID)) & "," & Val(Nvl(!����)) & "," & Val(Nvl(!����)) & "," & dbl�۸� & "," & lngִ�п���ID
                .MoveNext
            Loop
        End With
         If strItems = "" Then
            MsgBox "����δ�����κ���Ϣ,���ܱ���Ϊ�����շ���Ŀ,����!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
            Exit Sub
        End If
        strItems = Mid(strItems, 2)
   Else
        Dim dbl���� As Double, dbl���� As Double
        
        With mobjBill.Pages(mintPage)
            strItems = ""
            For i = 1 To .Details.Count
                 '0-����ȷ,1-���˿���,2-���˲���,3-����Ա����,4-ָ������,5-Ժ��ִ��(Ԥ��,������δ��),6-�����˿���
                If InStr(1, ",4,5,6,7,", "," & .Details(i).Detail.���) > 0 Then
                    lngִ�п���ID = 0
                    
                ElseIf InStr(1, ",0,4", .Details(i).Detail.ִ�п���) > 0 Then
                    lngִ�п���ID = .Details(i).ִ�в���ID
                Else
                    lngִ�п���ID = 0
                End If
                '����:52349
                dbl���� = .Details(i).����: dbl���� = IIf(.Details(i).Detail.���, .Details(i).InComes(1).��׼����, 0)
                If InStr(",5,6,7,", "," & .Details(i).Detail.���) > 0 And gblnҩ����λ Then
                     dbl���� = Format(.Details(i).���� * .Details(i).Detail.ҩ����װ, "0.00000")
                    dbl���� = Format(dbl����, gstrFeePrecisionFmt)
                End If
                
                strItems = strItems & "|" & .Details(i).��� & "," & .Details(i).�������� & "," & .Details(i).�շ�ϸĿID & "," & .Details(i).���� & ","
                strItems = strItems & dbl���� & "," & dbl���� & "," & lngִ�п���ID
             Next
             If strItems = "" Then
                MsgBox "����δ�����κ���Ϣ,���ܱ���Ϊ�����շ���Ŀ,����!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
                Exit Sub
            End If
            strItems = Mid(strItems, 2)
        End With
    End If
    Call mobjBaseItem.OpenEditWholeSetItem(Me, gcnOracle, glngSys, mlngModul, mstrPrivs, strItems)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdSelWholeSet_Click()
    'ѡ������Ŀ
    '����:34465
    Dim rsSel As ADODB.Recordset, lng����ID As Long, lng��������ID As Long
    Dim tmpBill As New ExpenseBill, bytӤ���� As Byte, Curdate As Date
    Dim curTotal  As Currency, rsTmp As ADODB.Recordset, i As Long
    Dim j As Long
    
    Dim bln��ҩ As Boolean
    
    If mobjBill Is Nothing Then
        If mrsInfo Is Nothing Then
            MsgBox "����ѡ����,����!", vbInformation + vbDefaultButton1, gstrSysName
            Exit Sub
        Else
            lng����ID = Val(Nvl(mrsInfo!����ID))
        End If
        
        If cbo��������.ListIndex < 0 Then
            lng��������ID = 0
        Else
            lng��������ID = cbo��������.ItemData(cbo��������.ListIndex)
        End If
        
        If cboBaby.ListIndex < 0 Then
            bytӤ���� = 0
        Else
            bytӤ���� = cboBaby.ItemData(cboBaby.ListIndex)
        End If
    Else
        lng����ID = mobjBill.����ID: lng��������ID = mobjBill.Pages(mintPage).��������ID: bytӤ���� = mobjBill.Ӥ����
    End If
     
    If zlSelectWholeItems(Me, mlngModul, mstrPrivs, rsSel) = False Then Exit Sub
    If rsSel Is Nothing Then Exit Sub
    Err = 0: On Error GoTo Errhand:
    Screen.MousePointer = 11
                         
    Set tmpBill = ImportWholeSet(Me, IIf(mstrYBPati <> "", mintInsure, 0), rsSel, mlng��ҩ��, mlng��ҩ��, mlng��ҩ��, _
        lng����ID, 0, gblnҩ����λ, lng��������ID, bytӤ����, 2, chk�Ӱ�.Value = 1, _
        0, gint������Դ, UserInfo.����, zlStr.NeedName(cbo������.Text), mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ�۸�ȼ�, _
        IIf(mbln���� And mlng��ҳID <> 0, mlng��ҳID, 0), IIf(mbln���� And mstr���ת��ʱ�� <> "", mlngDeptID, 0), _
        IIf(mbln���� And mstr���ת��ʱ�� <> "", mlngUnitID, 0))
    
     'a.���ŵ���ģʽ,�����ǰ���ݶ��󼰲�����Ϣ
    If Not cmdAddBill.Enabled Or Not cmdAddBill.Visible Then
        Dim rsTemp As ADODB.Recordset '95473
        Set rsTemp = mrsInfo
        Call ClearFullBill(False, False, True)
        Set mrsInfo = rsTemp
        
        '����:36764
        '֧��Ԥ����ʱ�Ͳ��̶���ʾ�����ʻ�,������ʾ
        If MCPAR.����Ԥ���� And mintInsure <> 0 Then
            '��ʾԤ���㰴ť
            cmdԤ����.Enabled = True
            Call SetButton(1) 'Ԥ����,ȷ��,ȡ��
            cmdOK.Enabled = False
        ElseIf mstr�����ʻ� <> "" Then 'ֻ��ʹ�ø����ʻ�����
            Call SetButton(2) 'ȷ��,ȡ��
            vsBalance.TextMatrix(0, 0) = mstr�����ʻ�
            vsBalance.TextMatrix(0, 1) = "0.00"
            vsBalance.RowData(0) = 0
        End If
        
        Set mobjBill = tmpBill
        mobjBill.�ѱ� = zlStr.NeedName(cbo�ѱ�.Text)
        If mobjBill.Pages(1).Details.Count > 0 Then
           If mobjBill.Pages(1).Details(1).�շ���� = "7" Then
                    bln��ҩ = True
           End If
        End If
        If InStr(mstrPrivs, "��ʾ������") = 0 Then mobjBill.Pages(mintPage).������ = ""
        '���������Ϣ
       ' Call ClearmobjBill
    Else
        'b.���ŵ���ģ��,��������,������ǰ�������ݼ����������Ϣ,
        If i > 0 Or mobjBill.Pages(mintPage).Details.Count > 0 Then
            Call AddNewBill
        End If
        mintPage = tbsBill.Tabs.Count
        
        '����Ҫ���벡�������Ϣ
        With mobjBill.Pages(mintPage)
            .NO = "" 'Ҫ����Ա��޸�ʱ������ֱ������ķ���
            .Key = tmpBill.Pages(1).Key
            .���ս�� = tmpBill.Pages(1).���ս��
            .��Ԥ���� = tmpBill.Pages(1).��Ԥ����
            .�巨 = tmpBill.Pages(1).�巨
            .����ͳ�� = tmpBill.Pages(1).����ͳ��
            .��������ID = tmpBill.Pages(1).��������ID
            If InStr(mstrPrivs, "��ʾ������") > 0 Then .������ = tmpBill.Pages(1).������
            .ȫ�Ը� = tmpBill.Pages(1).ȫ�Ը�
            .ʵ�ս�� = tmpBill.Pages(1).ʵ�ս��
            .�շѽ��� = tmpBill.Pages(1).�շѽ���
            .����� = tmpBill.Pages(1).�����
            .���Ը� = tmpBill.Pages(1).���Ը�
            .Ӧ�ɽ�� = tmpBill.Pages(1).Ӧ�ɽ��
            .Ӧ�ս�� = tmpBill.Pages(1).Ӧ�ս��

        End With
        bln��ҩ = False
        
        For j = 1 To tmpBill.Pages(1).Details.Count
            With tmpBill.Pages(1).Details(j)
                mobjBill.Pages(mintPage).Details.Add .�ѱ�, .Detail, .�շ�ϸĿID, .���, .��������, .�շ����, .���㵥λ, .��ҩ����, .����, .����, .���ӱ�־, .ִ�в���ID, .InComes, , .������Ŀ��, .���մ���ID, .���ձ���, .ժҪ
                If .�շ���� = "7" Then
                    bln��ҩ = True
                End If
            End With
        Next
         tbsBill.Tabs(mintPage).Selected = True '���ᴥ��Click�¼�
    End If
    Call Set�����˿�������(mobjBill.Pages(mintPage).������, mobjBill.Pages(mintPage).��������ID)
    'Call LoadAndSeek�ѱ�
    'ȡ��һҩƷ��
    For i = 1 To mobjBill.Pages(1).Details.Count
        If InStr(",5,6,7,", mobjBill.Pages(1).Details(i).�շ����) > 0 Then
            mlngFirstID = mobjBill.Pages(1).Details(i).ִ�в���ID
            mstrFirstWin = mobjBill.Pages(1).Details(i).��ҩ����
            Exit For
        End If
    Next
    Bill.Active = False
    Bill.Rows = mobjBill.Pages(mintPage).Details.Count + 1
    Call InitBillColumnColor
    
    If IIf(mlngPrePati = 0, mstrPrePati <> mobjBill.����, mlngPrePati <> mobjBill.����ID) Then
        '�²���
        mcurBillʵ�� = 0:  mcurBillӦ�� = 0: mcurBillӦ�� = 0
        mintBillNO = 0: mintMoneyRow = 0
    End If
    '�޸�ʱӦ���浱ǰ����Ա������
    mobjBill.����Ա��� = UserInfo.���
    mobjBill.����Ա���� = UserInfo.����
    Call CalcMoneys     '��Ϊ�����벡����Ϣ,������Ҫ���ݵ�ǰ�ķѱ�����۸�
    Call ShowDetails
    Call ShowMoney
    txtIn.Text = ""
    If mbytInState = 0 And mstrInNO <> "" Then txtModi.Text = "": mstrInNO = ""
        
    'Ҫ����mstrInNO֮��,��Ϊ�Դ����ж��Ƿ��޸ĵ���,�Լӻ�ԭ���
    Call CalcDrugStock
    Bill.Active = True
    ''�����к�
    Call SetColNum
    Screen.MousePointer = 0
    If bln��ҩ Then
        Call cmd�䷽_Click
    Else
        If mstrYBPati <> "" Then
            If cmdԤ����.Enabled And cmdԤ����.Visible Then
                cmdԤ����.SetFocus
            ElseIf cmdOK.Enabled And cmdOK.Visible Then
                cmdOK.SetFocus
            End If
        Else
            If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
        End If
    End If
    Exit Sub
Errhand:
    Screen.MousePointer = 0
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdYB_Click()
    txtPatient.SetFocus
    Call zlCommFun.PressKey(vbKeyF6)
End Sub



Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng�����ID As Long, strOutCardNO As String, strExpand  As String
    Dim strOutPatiInforXML As String
    If txtPatient.Locked Then Exit Sub
    
    If objCard.���� Like "IC��*" And objCard.ϵͳ Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        If mobjICCard Is Nothing Then Exit Sub
        
        '����:27364 ����:2010-01-13 15:27:50
        If mblnAutoChangePati And gint������Դ = 2 Then
            '��Ҫ���ҵ�������Դ1��
            gint������Դ = 1: zlChangePatiSource (gint������Դ)
        End If
        txtPatient.Text = mobjICCard.Read_Card()
        If txtPatient.Text <> "" Then
            Call txtPatient_KeyPress(vbKeyReturn)
            Call SetOneCardBalance
        End If
        Exit Sub
    End If
    If objCard.�ӿ���� <= 0 Then Exit Sub
    '    zlReadCard(frmMain As Object, _
    '    ByVal lngModule As Long, _
    '    ByVal lngCardTypeID As Long, _
    '    ByVal blnOlnyCardNO As Boolean, _
    '    ByVal strExpand As String, _
    '    ByRef strOutCardNO As String, _
    '    ByRef strOutPatiInforXML As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '����:�����ӿ�
    '    '���:frmMain-���õĸ�����
    '    '       lngModule-���õ�ģ���
    '    '       strExpand-��չ����,������
    '    '       blnOlnyCardNO-������ȡ����
    '    '����:strOutCardNO-���صĿ���
    '    '       strOutPatiInforXML-(������Ϣ����.XML��)
    '    '����:��������    True:���óɹ�,False:����ʧ��\
    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModul, objCard.�ӿ����, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtPatient.Text = strOutCardNO
    If txtPatient.Text <> "" Then
        If mblnAutoChangePati And gint������Դ = 2 Then
            '��Ҫ���ҵ�������Դ1��
            gint������Դ = 1: zlChangePatiSource (gint������Դ)
        End If
        Call txtPatient_KeyPress(vbKeyReturn)
    End If
End Sub

Private Sub SetOneCardBalance()
    Dim CurOneCard As Currency, strName As String
    
    If mblnOneCard And Not mobjICCard Is Nothing Then
        CurOneCard = mobjICCard.GetSpare(strName)
        If CurOneCard <> 0 Then
           mrsOneCard.Filter = "����='" & strName & "'"
           If mrsOneCard.RecordCount > 0 Then
                strName = mrsOneCard!���㷽ʽ
                If zlStr.NeedName(cbo���㷽ʽ) <> strName Then zlControl.CboLocate cbo���㷽ʽ, strName
           End If
        End If
        sta.Panels(Pan.C3�����ʻ�).Text = "�����:" & Format(CurOneCard, "0.00") & "Ԫ"
        sta.Panels(Pan.C3�����ʻ�).Visible = True
    End If
End Sub

Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txtPatient.IMEMode = 0
    Set gobjSquare.objCurCard = objCard
    If txtPatient.Text <> "" And Not mblnNotClick Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    zlControl.TxtSelAll txtPatient
End Sub
 

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, _
    objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    Dim lngPreIDKind As Long, intIndex As Integer
    Dim dtDate As Date
    Dim blnNew As Boolean
    'Or Not Me.ActiveControl Is txtPatient : Or txtPatient.Text <> ""
    '����:60010
    
    If txtPatient.Locked Then Exit Sub
    mblnNotClick = True
    
    intIndex = IDKind.GetKindIndex(objCard.����)
    lngPreIDKind = IDKind.IDKind
    If intIndex > 0 Then IDKind.IDKind = intIndex
    
    txtPatient.Text = objPatiInfor.����
    
    Call txtPatient_KeyPress(vbKeyReturn)
    If mrsInfo Is Nothing Then
        blnNew = True
    ElseIf mrsInfo.State <> 1 Then
        blnNew = True
    End If
    '�����²���
     If (txtPatient.Text = "" Or blnNew) And objPatiInfor.���� <> "" Then
        txtPatient.Text = objPatiInfor.����
        intIndex = IDKind.GetKindIndex("����")
        If intIndex > 0 Then IDKind.IDKind = IDKind.GetKindIndex("����")
        Call txtPatient_KeyPress(vbKeyReturn)
        If txtPatient.Text <> "" Then
        Call zlControl.CboLocate(cboSex, objPatiInfor.����)
            If IsDate(objPatiInfor.��������) = False Then
                 txt����.Text = ReCalcOld(CDate(objPatiInfor.��������), cbo���䵥λ, mobjBill.����ID)
            End If
        End If
    End If
    IDKind.IDKind = lngPreIDKind
    mblnNotClick = False
End Sub

Private Sub mFrmBalanceWin_zlSaveData(lng������� As Long, str����IDs As String, strSaveNos As String, blnNotCommit As Boolean, blnCancel As Boolean)
    '��������
    Dim blnSaveBill As Boolean
    mstrModiNOs = "": mstrSaveNos = "": strSaveNos = ""
    '�ȼ������Ƿ���ȷ:��Ҫ�ǽ�����صķ��ü��(���Ⲣ��ԭ����ɴ���)
    '����:62981
    If CheckChargeDataValied = False Then
        blnCancel = True
        Exit Sub
    End If
    Call SaveChargeBill(lng�������, str����IDs, strSaveNos, mstrModiNOs, blnSaveBill, blnNotCommit)
    '73960,Ƚ����,2014-6-17,һ��ͨд��ʱ,�ڲ����շѹ����������շѺ��zlMzInforWriteToCard�ӿڴ���Ĳ���lngBalanceIDΪ0û�д��벡�˵Ľ���id
    mlng������� = lng�������
    blnCancel = Not blnSaveBill
    mstrSaveNos = strSaveNos
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthday As Date, ByVal strAddress As String)
    Dim lngPreIDKind As Long
    
    If Not txtPatient.Locked And txtPatient.Text = "" And Me.ActiveControl Is txtPatient Then
        mblnNotClick = True
        lngPreIDKind = IDKind.IDKind
        IDKind.IDKind = IDKind.GetKindIndex("���֤��")
        txtPatient.Text = strID
        Call txtPatient_KeyPress(vbKeyReturn)
        
        '�����²���
        If txtPatient.Text = "" Then
            txtPatient.Text = strName
            IDKind.IDKind = IDKind.GetKindIndex("����")
            Call txtPatient_KeyPress(vbKeyReturn)
            If txtPatient.Text <> "" Then
                Call zlControl.CboLocate(cboSex, strSex)
                txt����.Text = ReCalcOld(datBirthday, cbo���䵥λ, mobjBill.����ID)
            End If
        End If
        IDKind.IDKind = lngPreIDKind
        mblnNotClick = False
    End If
End Sub
Private Sub mobjICCard_ShowICCardInfo(ByVal strNo As String)
    Dim lngPreIDKind As Long
    
    If Not txtPatient.Locked And txtPatient.Text = "" And Me.ActiveControl Is txtPatient And strNo <> "" Then
        lngPreIDKind = IDKind.IDKind
        mblnNotClick = True
        Dim intIndex As Integer
        intIndex = IDKind.GetKindIndex("IC����")
        If intIndex <= 0 Then mblnNotClick = False: Exit Sub
        IDKind.IDKind = intIndex
        txtPatient.Text = strNo
        Call txtPatient_KeyPress(vbKeyReturn)
        If txtPatient.Text = "" Then Call mobjICCard.SetEnabled(False)
        IDKind.IDKind = lngPreIDKind
        mblnNotClick = False
        If Not txtPatient.Locked And Me.ActiveControl Is txtPatient Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
    End If
End Sub

Private Sub Bill_BeforeAddRow(Row As Long)
'˵����RowΪ��Ҫ�������к�,��ǰ�к�ΪRow-1
    Dim dbl���� As Double, cur��� As Currency, i As Integer
    
    'LED��̬��ʾ��Ŀ
    If mbytInFun = 0 And gblnLED And mbytInState = 0 And gblnLedDispDetail Then
        If mobjBill.Pages(mintPage).Details.Count >= Row - 1 Then
            With mobjBill.Pages(mintPage).Details(Row - 1)
                dbl���� = 0: cur��� = 0
                For i = 1 To .InComes.Count
                    cur��� = cur��� + .InComes(i).ʵ�ս��
                    dbl���� = dbl���� + .InComes(i).��׼����
                Next
                'LED��ʾ
                If cur��� <> 0 Then
                    If InStr(",5,6,7,", .Detail.���) > 0 And gblnҩ����λ Then
                        '��ҩ����λ��ʾ��λ
                        zl9LedVoice.Display .Detail.����, .Detail.���, .Detail.ҩ����λ, dbl����, IIf(.���� = 0, 1, .����) * .����, cur���
                    Else
                        zl9LedVoice.Display .Detail.����, .Detail.���, .���㵥λ, dbl����, IIf(.���� = 0, 1, .����) * .����, cur���
                    End If
                End If
            End With
        End If
    End If
End Sub

Private Sub ShowGroupLED(ByVal lngMain As Long, ByVal lngBegin As Long, ByVal lngEnd As Long)
'���ܣ�Ϊ�ӿ��ٶȣ�һ���Ե����ײ���Ŀ��LED��ʾ
'�������кŷ�Χ��lngMain=�����к�,lngBegin-lngEnd:�����к�
    Dim dbl���� As Double, dbl���� As Double, cur��� As Currency
    Dim i As Long, j As Long
    
    If mbytInFun = 0 And mbytInState = 0 And gblnLED And gblnLedDispDetail Then
        With mobjBill.Pages(mintPage)
            For j = 1 To .Details(lngMain).InComes.Count
                cur��� = cur��� + .Details(lngMain).InComes(j).ʵ�ս��
            Next
            For i = lngBegin To lngEnd
                For j = 1 To .Details(i).InComes.Count
                    cur��� = cur��� + .Details(i).InComes(j).ʵ�ս��
                Next
            Next
        End With
        With mobjBill.Pages(mintPage).Details(lngMain)
            If cur��� <> 0 Then
                dbl���� = IIf(.���� = 0, 1, .����) * .����
                If dbl���� <> 0 Then
                    dbl���� = cur��� / dbl����
                Else
                    dbl���� = cur���
                End If
                If InStr(",5,6,7,", .Detail.���) > 0 And gblnҩ����λ Then
                    zl9LedVoice.Display .Detail.����, .Detail.���, .Detail.ҩ����λ, dbl����, dbl����, cur���
                Else
                    zl9LedVoice.Display .Detail.����, .Detail.���, .���㵥λ, dbl����, dbl����, cur���
                End If
            End If
        End With
    End If
End Sub


Private Sub Bill_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    Dim i As Long, bytSubs As Byte
    Dim bln��������ۿ� As Boolean
    Dim lngMainRow As Long
    
    If mbytInState <> 0 Or chkCancel.Value = 1 Then Cancel = True: Exit Sub
    
    With mobjBill.Pages(mintPage)
        If .Details.Count >= Row Then
            If .Details(Row).������ Then
                MsgBox "���в����޸ļ�ɾ����", vbInformation, gstrSysName
                Cancel = True: Exit Sub
            End If
        End If
        
        If .Details.Count >= Row Then
            '��������Ŀ����ɾ��ȷ��
            For i = Row + 1 To .Details.Count
                If .Details(i).�������� = Row Then bytSubs = bytSubs + 1
            Next
            If bytSubs > 0 Then
                If MsgBox("����Ŀ���� " & bytSubs & " ��������Ŀ,ɾ������ĿҲ��ɾ�����Ĵ�����Ŀ,������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Cancel = True: Exit Sub
                End If
            ElseIf .Details(Row).�������� <> 0 Then '������Ŀɾ��ȷ��
                If MsgBox("����Ŀ��[" & .Details(.Details(Row).��������).Detail.���� & "]�Ĵ�����Ŀ,ȷ��Ҫɾ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Cancel = True: Exit Sub
                Else
                    bln��������ۿ� = gbln��������ۿ�
                End If
            ElseIf MsgBox("ȷʵҪɾ�����շ���Ŀ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True: Exit Sub
            End If
            
            If bln��������ۿ� Then lngMainRow = .Details(Bill.Row).�������� '����Ǵ���,ɾ��֮ǰ���´���Ĵ�������,���������,����ɾ��,��������
            
            'ɾ���������(��˳��)
            For i = .Details.Count To Row + 1 Step -1
                If .Details(i).�������� = Row Then
                    Call DeleteDetail(i)
                End If
            Next
 
            Call DeleteDetail(Row) 'ɾ������
            
            
            If bln��������ۿ� Then
                If CheckMainItem(lngMainRow) Or lngMainRow > 0 Then
                    Call CalcPItemActualIncome(lngMainRow)
                Else
                    Call CalcMoney(mintPage, lngMainRow, False)  'ֻ��һ��������,����ȫ����ɾ��ʱ,������ͨ���������
                End If
            End If
                        
            '���¼��������в�ˢ��
            Call ShowDetails
            Call ShowMoney(mintPage)
            
            '��Ҫ����Ԥ����
            If cmdԤ����.Visible Then
                Call InitBalanceGrid
                cmdԤ����.TabStop = True
                cmdOK.Enabled = False
            End If
            
'''            Call zlClear���㿨
            
            If CheckBillsEmpty Then ClearMoney
                                   
            Bill.TxtVisible = False
            Bill.CmdVisible = False
            Bill.CboVisible = False
            
            Cancel = True '���ÿؼ�������ɾ��
            
            mlngPreRow = 0  '��ʾ�иı���
            Call Bill_EnterCell(Bill.Row, Bill.Col)
        ElseIf Row = 1 Then
            For i = 1 To Bill.COLS - 1
                Bill.TextMatrix(Row, i) = ""
            Next
            Call SetBillRowForeColor(Row, Bill.ForeColor)
            Cancel = True
        End If
    End With
    
    Call SetColNum(Row)
End Sub

Private Sub Bill_cboClick(ListIndex As Long)
    Dim dblStock As Double, strStock As String
    Dim blnComboxDown As Boolean
    Dim lngִ�п��� As Long, strִ�п��� As String
    'ҩƷ�����
    If ListIndex <> -1 And (Bill.TextMatrix(0, Bill.Col) = "ִ�п���" Or Bill.TextMatrix(0, Bill.Col) = "��ҩҩ��") Then
        If mobjBill.Pages(mintPage).Details.Count >= Bill.Row Then
            With mobjBill.Pages(mintPage).Details(Bill.Row)
                blnComboxDown = SendMessage(Bill.cboHwnd, &H157, 0, 0) = 1
                If .ִ�в���ID <> Bill.ItemData(Bill.ListIndex) Then
                    lngִ�п��� = .ִ�в���ID: strִ�п��� = Bill.TextMatrix(Bill.Row, Bill.Col)
                    .ִ�в���ID = Bill.ItemData(Bill.ListIndex)
                    .��ҩ���� = ""
                    Bill.TextMatrix(Bill.Row, Bill.Col) = Bill.CboText
                     
                    If InStr(",5,6,7,", .�շ����) > 0 Then
                        'ȡ���,������޸Ĺ���,��ʱ��ȡ����,����ϵ�ǰ�����ڸÿⷿ�Ŀ��,��Ƚ��鷳,��ʱ����
                        dblStock = GetStock(.�շ�ϸĿID, .ִ�в���ID)
                        If gblnҩ����λ Then
                            dblStock = dblStock / .Detail.ҩ����װ
                        End If
                        .Detail.��� = dblStock  '��¼��ǰ��ҩƷ���
                        Call ShowStock(.ִ�в���ID, .Detail.����, .Detail.���)
                        Call ShowStatusCargoSpace(.�շ�ϸĿID, .ִ�в���ID)    '��ʾ��λ
                        
                        'ҩ���ı�,ʱ��ҩƷ���¼���۸�
                        'If .Detail.��� Then    '����ѱ�ļ��㷽ʽ�ǳɱ��ۼ��շ�,����Ҫ����۸�,����򻯲����ж�
                            Call CalcMoneys(mintPage, Bill.Row)
                            Call ShowDetails(Bill.Row)
                            Call ShowMoney(mintPage)
                        'End If
                        '�����޶���ʾ:
                        Call SetItemRowColor(mintPage, Bill.Row)
                        If blnComboxDown Then '��ʾ�������˵�:����:25238
                            DoEvents
                             SendMessageLong Bill.cboHwnd, &H14F, True, 0
                        End If
                    
                    ElseIf .�շ���� = "4" And .Detail.�������� Then
                        'ȡ���
                        dblStock = GetStock(.�շ�ϸĿID, .ִ�в���ID, .Detail.����)
                        .Detail.��� = dblStock
                        Call ShowStock(.ִ�в���ID, .Detail.����, .Detail.���)
                        
                        '���ϲ��Ÿı�,ʱ���������¼���۸�
                        If .Detail.��� Then
                            Call CalcMoneys(mintPage, Bill.Row) '�����Ҫ���ܼ���,����������ʵ��
                            Call ShowDetails(Bill.Row)
                            Call ShowMoney(mintPage)
                        End If
                        '�����޶���ʾ:
                        Call SetItemRowColor(mintPage, Bill.Row)
                        If blnComboxDown Then '��ʾ�������˵�:����:25238
                            DoEvents
                             SendMessageLong Bill.cboHwnd, &H14F, True, 0
                        End If
                    '�շ���Ŀ
                    ElseIf InStr(",4,5,6,7,", .�շ����) = 0 Then
                        If CheckMainItem(Bill.Row) Then Call SetSubItemDept(Bill.Row) '������ڴ���,��ı��ҩƷ�е�ִ�п���
                    End If
                    If Bill.TextMatrix(0, Bill.Col) = "ִ�п���" Then
                        If mbytInFun = 0 And mintInsure <> 0 And MCPAR.ʵʱ��� And mobjBill.Pages(mintPage).Details(Bill.Row).���� <> 0 Then
                            If gclsInsure.CheckItem(mintInsure, 0, 0, MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), 1, 0, mintPage, Bill.Row)) = False Then
                                Bill.Text = "": Bill.TxtVisible = False
                                Bill.cboObj.Text = strִ�п���: .ִ�в���ID = lngִ�п���
                                Exit Sub
                            End If
                        End If
                        
                        If mobjBill.Pages(mintPage).Details(Bill.Row).���� <> 0 Then
                            If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModul, 0, 0, _
                                MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), mbytInFun, IIf(mbytInFun = 1, 1, IIf(mbytBilling = 0, 0, 1)), mintPage, Bill.Row)) = False Then
                                Bill.Text = "": Bill.TxtVisible = False
                                Bill.cboObj.Text = strִ�п���: .ִ�в���ID = lngִ�п���
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            End With
        End If
    End If
End Sub

Private Function GetColNum(strHead As String) As Integer
    Dim i As Integer
    For i = 0 To Bill.COLS - 1
        If Bill.TextMatrix(0, i) = strHead Then GetColNum = i: Exit Function
    Next
End Function

Public Function GetOriginalTotal(ByVal objBill As ExpenseBill, ByVal lngҩƷID As Long, ByVal lngҩ��ID As Long, _
    Optional ByVal intPage As Integer) As Double
'���ܣ���ȡ������ָ��ҩƷ��ͬһҩ�����е�ԭʼ������
'������ lngҩ��ID-0��ʾ���뷢ҩʱ,���޶�ҩ�����
    Dim i As Integer, p As Integer, dblCount As Double
    
    For p = 1 To objBill.Pages.Count
        If intPage = 0 Or p = intPage Then
            For i = 1 To objBill.Pages(p).Details.Count
                If objBill.Pages(p).Details(i).�շ�ϸĿID = lngҩƷID Then
                    If IIf(lngҩ��ID <> 0, objBill.Pages(p).Details(i).ԭʼִ�в���ID = lngҩ��ID, 1 = 1) Then
                        dblCount = dblCount + objBill.Pages(p).Details(i).ԭʼ����
                    End If
                End If
            Next
        End If
    Next
    GetOriginalTotal = dblCount
End Function

Private Function GetDelMoney(Optional curError As Currency) As Currency
'���ܣ���ȡ�����˷�ʱ��ǰ��ѡ������˿���
'���أ�curError=�����˷�ʱ�����������
'˵�����˷�ʱ�������ͷֱң������˷�ʱ�Ŵ������
    Dim cur���ݺϼ� As Currency
    Dim curѡ��ϼ� As Currency
    Dim cur�˷Ѻϼ� As Currency
    Dim bln��ȫ�˷� As Boolean, bln�ֽ���� As Boolean
    Dim intCOL_ѡ�� As Integer, intCOL_��� As Integer
    Dim i As Integer, strTempBalance As String, blnԭ���� As Boolean
    Dim bln�������� As Boolean, bln���������˲����� As Boolean
    
    curError = 0
    intCOL_ѡ�� = GetColNum("�˷�")
    intCOL_��� = GetColNum("ʵ�ս��")
    
    '���ݺϼƽ��
    For i = 1 To Bill.Rows - 1
        cur���ݺϼ� = cur���ݺϼ� + Val(Bill.TextMatrix(i, intCOL_���))
        If Bill.TextMatrix(i, intCOL_ѡ��) <> "" Then
            curѡ��ϼ� = curѡ��ϼ� + Val(Bill.TextMatrix(i, intCOL_���))
        End If
    Next
    mTyDelFee.dblCurDelMoney = curѡ��ϼ�
    
    '��ȫ�˷�ʱ�ſ�����������
    bln��ȫ�˷� = Not BillExistDelete(mstrInNO, 1) _
        And BillDeleteAll(mstrInNO, 1, mblnHaveExcuteData) And (cur���ݺϼ� = curѡ��ϼ�)
    blnԭ���� = bln��ȫ�˷�
    bln�������� = False
    vsBalance.Tag = ""
    If bln��ȫ�˷� Then
        For i = 0 To vsBalance.Rows - 1
            '�˷�ʱ������ʾ�˶��ַ�ҽ������,Ҫ�ſ�,-1��ʾ��ҽ������
            If vsBalance.TextMatrix(i, 0) <> "" Then
                If IsNumeric(vsBalance.TextMatrix(i, 1)) And vsBalance.RowData(i) <> -1 Then
                    strTempBalance = vsBalance.TextMatrix(i, 0)
                    '������ֽ��㷽ʽ��֧�ֻ���,Ҫ��Ϊ�ֽ�,���ü�ȥ
                    If mblnYB�������� Then
                        If gclsInsure.GetCapability(support�����������, , mintInsure, strTempBalance) Then
                            curѡ��ϼ� = curѡ��ϼ� - Val(vsBalance.TextMatrix(i, 1))
                            vsBalance.Cell(flexcpBackColor, i, 0, i, 1) = &HE7CFBA
                        Else
                            blnԭ���� = False
                        End If
                    Else     '��֧�������������ʱ,ֻ���������Ϊ�ֽ�,����ԭ����,������ҽ������
                        If strTempBalance <> mstr�����ʻ� Then
                            curѡ��ϼ� = curѡ��ϼ� - Val(vsBalance.TextMatrix(i, 1))
                            vsBalance.Cell(flexcpBackColor, i, 0, i, 1) = &HE7CFBA
                        End If
                        If InStr("3,4,5", vsBalance.Cell(flexcpData, i, 0)) = 0 Then
                            blnԭ���� = False
                        Else
                            curѡ��ϼ� = curѡ��ϼ� - (vsBalance.Cell(flexcpData, i, 1) - Val(vsBalance.TextMatrix(i, 1)))
                            vsBalance.TextMatrix(i, 1) = FormatEx(vsBalance.Cell(flexcpData, i, 1), 2)
                        End If
                    End If
             
                Else
                    If InStr("3,4,5", vsBalance.Cell(flexcpData, i, 0)) > 0 Then
                        'ȱʡ�ų����㿨;ҽ�ƿ�;һ��ͨ������
                        'vsBalance.Cell(flexcpData, lngRow, 0)=����:1-Ԥ���,2-ҽ��,3-ҽ�ƿ�,4-���㿨,5-һ��ͨ,0-������
                        If Val(vsBalance.TextMatrix(i, 1)) <> 0 Then vsBalance.TextMatrix(i, 1) = FormatEx(vsBalance.Cell(flexcpData, i, 1), 2)
                        curѡ��ϼ� = curѡ��ϼ� - Val(vsBalance.TextMatrix(i, 1))
                        vsBalance.Cell(flexcpBackColor, i, 0, i, 1) = &HE7CFBA
                        If Not bln�������� Then
                            bln�������� = Val(vsBalance.TextMatrix(i, 1)) = 0 And Val(vsBalance.Cell(flexcpData, i, 1)) <> 0
                        End If
                    End If
                End If
            End If
        Next
        curѡ��ϼ� = curѡ��ϼ� - Val(txtԤ�����.Text)
    Else
        '������:55064
         vsBalance.Tag = "1"
        With vsBalance
            For i = 0 To .Rows - 1
                If mTyDelFee.blnSingleBalance And Val(.Cell(flexcpData, i, 0)) = 3 Then
                    If Val(.RowData(i)) = -1 And mTyDelFee.bln������ȫ�� = False Then
                        If Val(vsBalance.TextMatrix(i, 1)) <> 0 Then
                            .TextMatrix(i, 1) = FormatEx(curѡ��ϼ�, 2)
                        End If
                        curѡ��ϼ� = curѡ��ϼ� - Val(.TextMatrix(i, 1))
                    ElseIf Val(.RowData(i)) = -1 And mTyDelFee.bln������ȫ�� Then
                        .TextMatrix(i, 1) = ""
                    ElseIf Val(vsBalance.Cell(flexcpData, i, 0)) = 3 Then '3-ҽ�ƿ�
                        .TextMatrix(i, 1) = FormatEx(IIf(vsBalance.Cell(flexcpData, i, 1) > curѡ��ϼ�, _
                                        curѡ��ϼ�, vsBalance.Cell(flexcpData, i, 1)), 2)
                        curѡ��ϼ� = curѡ��ϼ� - Val(.TextMatrix(i, 1))
                        bln���������˲����� = True '������
                    End If
                Else
                    If Val(.RowData(i)) = -1 Then
                        .TextMatrix(i, 1) = ""
                    End If
                End If
            Next
        End With
    End If
    

    If Not bln��ȫ�˷� Then '68177
        '�շ�ʱȫ����Ԥ��,�˷�ʱ,������ָ���˷ѷ�ʽ
        '�����˷�ʱ�����Ԥ����������벿���˵Ľ��һ��ʱ��������ΪԤ����:Val(txtԤ�����.Text) = GetBillSum And Val(txtԤ�����.Text) <> 0
        If mblnOlnyԤ�� Then
             blnԭ���� = True  '�����ǲ�����
            txtԤ�����.Visible = True
            lblԤ�����.Visible = True
            lblӦ��.Caption = "��Ԥ��"
        Else
            txtԤ�����.Visible = False
            lblԤ�����.Visible = False
            lblӦ��.Caption = "�˿�"
        End If
    Else
        '�շ�ʱȫ����Ԥ��,�˷�ʱ,������ָ���˷ѷ�ʽ
        If Val(txtԤ�����.Text) = GetBillSum Then blnԭ���� = True  '�����ǲ�����
        txtԤ�����.Visible = Val(txtԤ�����.Text) <> 0
        lblԤ�����.Visible = Val(txtԤ�����.Text) <> 0
        lblӦ��.Caption = "�˿�"
    End If
    
    If blnԭ���� And Not bln�������� Then
        zlControl.CboSetIndex cbo���㷽ʽ.hWnd, mintReturnMode
    End If
    cbo���㷽ʽ.Enabled = (Not blnԭ���� Or bln��������) And mblnOlnyԤ�� = False
    cbo���㷽ʽ.Locked = (blnԭ���� And Not bln��������) And mblnOlnyԤ�� = False
    fraAppend.Enabled = (Not blnԭ���� Or bln��������) And mblnOlnyԤ�� = False
    If mTyDelFee.blnSingleBalance And bln���������˲����� Then
        cbo���㷽ʽ.Enabled = False
        cbo���㷽ʽ.Locked = True
        fraAppend.Enabled = False
    End If
    If cbo���㷽ʽ.Enabled And cbo���㷽ʽ.Locked = False Then
        If cbo���㷽ʽ.ListCount > 0 And cbo���㷽ʽ.ListIndex = -1 Then
            cbo���㷽ʽ.ListIndex = 0
        End If
    End If
    
    '�ֽ����ʱ����ֱ�
    bln�ֽ���� = False
    If cbo���㷽ʽ.ListIndex <> -1 And mblnOlnyԤ�� = False Then
        If cbo���㷽ʽ.ItemData(cbo���㷽ʽ.ListIndex) = 1 Then
            bln�ֽ���� = True
        End If
    End If
    
    '���ý���λ��,���ֽ����ʱ����ֱ�
    If bln�ֽ���� Then
        If mintInsure > 0 Then
            If gclsInsure.GetCapability(support�ֱҴ���, , mintInsure) Then
                cur�˷Ѻϼ� = CentMoney(curѡ��ϼ�)
            Else
                cur�˷Ѻϼ� = Format(curѡ��ϼ�, "0.00")
            End If
        Else
            cur�˷Ѻϼ� = CentMoney(curѡ��ϼ�)
        End If
    Else
        cur�˷Ѻϼ� = Format(curѡ��ϼ�, "0.00")
    End If
    
    '�����,������,��ҽ��ȫ��ʱ��Ϊ���㷽ʽ��֧�ֻ��˶���Ϊ�ֽ�,���ܲ������
    '���ֽ����ʱ,Ҳ���������,�������Ƿ��ý���λ�������
    '60974
    curError = cur�˷Ѻϼ� - curѡ��ϼ�
    If curError <> 0 Then
        vsBalance.ToolTipText = "�˷Ѳ����������:" & curError
    Else
        vsBalance.ToolTipText = "���β���û�������!"
    End If
'    If Not blnԭ���� Then
'        curError = cur�˷Ѻϼ� - curѡ��ϼ�
'        vsBalance.ToolTipText = "�˷Ѳ����������:" & curError
'    Else
'        vsBalance.ToolTipText = "���β���û�������!"
'    End If
'
    GetDelMoney = cur�˷Ѻϼ�
End Function

Private Sub SelALLRow()
    '���ܣ�ʵ���˷�ʱ��ȫѡ
    Dim i As Long
    If InStr(",����,�˷�,", Bill.TextMatrix(0, Bill.COLS - 1)) > 0 Then
        For i = 1 To Bill.Rows - 1
            If Bill.TextMatrix(i, BillCol.��Ŀ) <> "" Then
                Bill.TextMatrix(i, Bill.COLS - 1) = "��"
            End If
        Next
    End If
    If mbytInFun = 0 And Bill.TextMatrix(0, Bill.COLS - 1) = "�˷�" Then
        Call ReCalce�˿�
    End If
End Sub

Private Sub ClearALLRow()
'���ܣ�ʵ���˷�ʱ��ȫ��
    Dim i As Long
    
    If InStr(",����,�˷�,", Bill.TextMatrix(0, Bill.COLS - 1)) > 0 Then
        '������ʱ��Ȼ���շ�
        If mintInsure <> 0 Then sta.Panels(Pan.C2��ʾ��Ϣ).Text = "ҽ�����˵��շѵ��ݲ��������˷�!": Exit Sub
'        '���˺�:??
'        If mtySquareCard.bln������ Then
'            sta.Panels(Pan.C2��ʾ��Ϣ).Text = "ˢ���㿨���˵��շѵ��ݲ��������˷�!": Exit Sub
'        End If
'
        For i = 1 To Bill.Rows - 1
            Bill.TextMatrix(i, Bill.COLS - 1) = ""
        Next
    End If
    If mbytInFun = 0 And Bill.TextMatrix(0, Bill.COLS - 1) = "�˷�" Then
        Call ReCalce�˿�
    End If
End Sub
Private Sub zlSet���ƹ̶���ϵ(ByVal lngRow As Long, ByVal Col As Long, Optional lngNotCheckRow As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:סԺ���ü�¼
    '����:���˺�
    '����:2010-12-31 15:49:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, varTemp As Variant, bln�̶� As Boolean, i As Long, j As Long
    If Bill.TextMatrix(lngRow, BillCol.ҽ�����) = "" Then Exit Sub
    If mrs�շѶ��� Is Nothing Then Exit Sub
     '����:33634:����ǹ̶�����Ŀ(�����շѹ�ϵ):��ҽ�������Ĳ��ж�
    varData = Split(Bill.TextMatrix(lngRow, BillCol.ҽ�����) & ",", ",")
    If Val(varData(0)) = 0 Then Exit Sub
    
    mrs�շѶ���.Filter = "ҽ�����=" & Val(varData(0)) & " And �շ�ϸĿID=" & Val(varData(1))
    If Not mrs�շѶ���.EOF Then
        bln�̶� = Val(Nvl(mrs�շѶ���!���ж���)) = 1
    Else
        bln�̶� = False
    End If
    mrs�շѶ���.Filter = 0
    If bln�̶� = False Then Exit Sub
    
    For i = 1 To Bill.Rows - 1
        If i <> lngRow And lngNotCheckRow <> i Then
            varTemp = Split(Bill.TextMatrix(i, BillCol.ҽ�����) & ",", ",")
            If varData(0) = varTemp(0) Then    '����ͬ��ҽ�����
                 mrs�շѶ���.Filter = "ҽ�����=" & Val(varTemp(0)) & " And �շ�ϸĿID=" & Val(varTemp(1))
                If Not mrs�շѶ���.EOF Then
                    bln�̶� = Val(Nvl(mrs�շѶ���!���ж���)) = 1
                Else
                    bln�̶� = False
                End If
                If bln�̶� Then
                    Bill.TextMatrix(i, Col) = Bill.TextMatrix(lngRow, Col)
                    '���������,��Ҫ�������
                    If Val(Bill.TextMatrix(i, BillCol.��������)) = 0 Then '�϶�Ϊ����,���,��Ҫ�Ҵ�������
                        For j = i + 1 To Bill.Rows - 1
                             If Bill.RowData(i) = Val(Bill.TextMatrix(j, BillCol.��������)) Then
                                    Bill.TextMatrix(j, Col) = Bill.TextMatrix(i, Col)
                             End If
                        Next
                    End If
                End If
             End If
        End If
    Next
End Sub


Private Sub Bill_CellCheck(Row As Long, Col As Long)
'˵��������ȫ��Ϊ��Ҫ����,������ȫ��Ϊ��������
    Dim i As Long, strCheck As String, bytTime As Byte
    Dim blnReSet As Boolean '��������
    Dim bln�̶� As Boolean, strErrMsg As String, varData As Variant ' (0-ҽ�����;1-�շ�ϸĿID)
    Dim varTemp As Variant
    Dim bln�̶�1 As Boolean
    Dim j As Long
    '�˷�,˫���˷���
    
    If mbytInFun = 0 And Bill.TextMatrix(0, Col) = "�˷�" Then
        If mintInsure <> 0 Then sta.Panels(Pan.C2��ʾ��Ϣ).Text = "ҽ�����˵��շѵ��ݲ��������˷�!": Exit Sub
        '���˺�:??
        sta.Panels(2).Text = ""
        If Bill.TextMatrix(Row, Col) = "" Then
            With mTyDelFee.rsBlance
                 .Filter = 0
                 '����:1-Ԥ���,2-ҽ��,3-ҽ�ƿ�,4-���㿨,5-һ��ͨ,0-������
                .Filter = "(�Ƿ�ȫ��=1 And �Ƿ�����=0) Or (����=4 And �Ƿ�ȫ��=0 And �Ƿ�����=0) " & _
                        "Or (����=5 And �Ƿ�ȫ��=0 And �Ƿ�����=0)"
                If .RecordCount <> 0 Then .MoveFirst
                strErrMsg = ""
                If Not .EOF Then
                    strErrMsg = Nvl(!����) & ":" & Format(Val(Nvl(!������, 0)), "0.00")
                End If
                If strErrMsg <> "" Then
                    sta.Panels(Pan.C2��ʾ��Ϣ).Text = "���ڵ���������(" & strErrMsg & ")�����������˷�": Exit Sub
                End If
                
                If Not mTyDelFee.blnSingleBalance Then
                    .Filter = "����=3 And �Ƿ�ȫ��=0 And �Ƿ�����=0"
                    If .RecordCount <> 0 Then .MoveFirst
                    strErrMsg = ""
                    If Not .EOF Then
                        strErrMsg = Nvl(!����) & ":" & Format(Val(Nvl(!������, 0)), "0.00")
                    End If
                    If strErrMsg <> "" Then
                        sta.Panels(Pan.C2��ʾ��Ϣ).Text = "���ڵ���������(" & strErrMsg & ")�����������˷�": Exit Sub
                    End If
                End If
            End With
        End If
        
        '����:29201
        '��������
        If Val(Bill.TextMatrix(Row, BillCol.��������)) = 0 Then '�϶�Ϊ����,���,��Ҫ�Ҵ�������
            For i = Row + 1 To Bill.Rows - 1
                 If Bill.RowData(Row) = Val(Bill.TextMatrix(i, BillCol.��������)) Then
                        Bill.TextMatrix(i, Col) = Bill.TextMatrix(Row, Col)
                 End If
            Next
            Call zlSet���ƹ̶���ϵ(Row, Col)
        Else
            Call zlSet���ƹ̶���ϵ(Row, Col)
            '��Ҫ��������Ƿ��Ѿ���
                For i = Row - 1 To 1 Step -1
                    If Bill.RowData(i) = Val(Bill.TextMatrix(Row, BillCol.��������)) Then
                        If Bill.TextMatrix(i, Col) <> "" Then
                            Bill.TextMatrix(i, Col) = Bill.TextMatrix(Row, Col)
                        End If
                        Call zlSet���ƹ̶���ϵ(i, Col, Row)
                         Exit For
                    End If
                Next
        End If
        Call ReCalce�˿�
        Call LoadInvoiceData(mTyDelFee.strNos)
        Call ShowInvoiceInfor
    End If
    If mbytInState = 3 Or (chkCancel.Visible And chkCancel.Value = 1) Then
        '����:29201
        '��������
        If mbytInFun <> 0 Then  '�����Ѿ�����
            If Val(Bill.TextMatrix(Row, BillCol.��������)) = 0 Then '�϶�Ϊ����,���,��Ҫ�Ҵ�������
                    For i = Row + 1 To Bill.Rows - 1
                         If Bill.RowData(Row) = Val(Bill.TextMatrix(i, BillCol.��������)) Then
                                Bill.TextMatrix(i, Col) = Bill.TextMatrix(Row, Col)
                         End If
                    Next
            End If
        End If
        Exit Sub
    End If
    
    '������δ��������Ч
    If Bill.TextMatrix(Row, BillCol.��Ŀ) = "" Then Bill.TextMatrix(Row, Col) = "": Exit Sub
    If mobjBill.Pages(mintPage).Details.Count < Row Then
        Bill.TextMatrix(Row, Col) = ""
        Exit Sub
    End If
    
    strCheck = Bill.TextMatrix(Row, Col)
    
    For i = 1 To mobjBill.Pages(mintPage).Details.Count
        If mobjBill.Pages(mintPage).Details(i).�շ���� = "F" _
            And mobjBill.Pages(mintPage).Details(i).���ӱ�־ = 0 And i <> Row Then
            bytTime = bytTime + 1
        End If
    Next
    
    blnReSet = bytTime > 0
    If blnReSet = False Then     '����ֻ���ڸ����������ָĳ���������,��Ҫ���¼ƴ���
        blnReSet = (strCheck = "" And mobjBill.Pages(mintPage).Details(Row).�շ���� = "F" And mobjBill.Pages(mintPage).Details(Row).���ӱ�־ = 1)
    End If
    If blnReSet Then
        With mobjBill.Pages(mintPage).Details(Row)
            
            .���ӱ�־ = IIf(strCheck = "", 0, 1)
            Call CalcMoneys(mintPage, Row)
            Call ShowDetails(Row)
        End With
        
        Call ShowMoney(mintPage)
        
        '��Ҫ����Ԥ����
        If cmdԤ����.Visible Then
            Call InitBalanceGrid
            cmdԤ����.TabStop = True
            cmdOK.Enabled = False
        End If
    ElseIf strCheck <> "" Then
        Bill.TextMatrix(Row, Col) = ""
        MsgBox "�����б�Ȼ��һ���������Ǹ���������", vbInformation, gstrSysName
    End If
    
End Sub

Private Sub Bill_CommandClick()
    Dim lng��Ŀid As Long, blnCancel As Boolean, bln��ʿ As Boolean
    Dim str��� As String, str��׼��Ŀ As String
    Dim str�ų���� As String, lng���� As Long
    
    Call GetOperatorInfo(mobjBill.Pages(mintPage).������, bln��ʿ)
    If gbln�շ���� Then
        If Bill.RowData(Bill.Row) <> 0 Then
            str��� = "'" & Chr(Bill.RowData(Bill.Row)) & "'"
        Else
            str��� = IIf(bln��ʿ, "'E','M','4'", gstr�շ����)
        End If
    Else
        str��� = IIf(bln��ʿ, "'E','M','4'", gstr�շ����)
    End If
    If mbytInFun = 0 And mstrYBPati <> "" Then
        '���˺�:24862
        If zl_Check��׼��Ŀ(gclsInsure, mintInsure, mobjBill.����ID, True) Then str��׼��Ŀ = Get������׼��Ŀ(mobjBill.����ID, "A.ID")
    End If
    If zlCheckBill���ڷ�ɢװ��ҩ(mintPage) = True Then
        mblnSelect = False: Exit Sub
    End If
    lng���� = -1
    lng��Ŀid = frmItemSelect.ShowSelect(Me, mstrPrivs, gint������Դ, mintInsure, gblnҩ����λ, str���, _
        , , str��׼��Ŀ, 0, , , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ�۸�ȼ�, mbln����ˢ��, lng����)
    If lng��Ŀid <> 0 Then
        Bill.Text = lng��Ŀid
        Bill.Tag = lng����
        mblnSelect = True
        Call Bill_KeyDown(13, 0, blnCancel)
        Bill.SetFocus
        If Not blnCancel Then
            Bill.Text = "": Bill.TxtVisible = False
            Call zlCommFun.PressKey(13)
        End If
    Else
        mblnSelect = False
    End If
End Sub

Private Sub ShowStock(ByVal lng�ⷿID As Long, strҩƷ As String, dbl��� As Double)
    '���ܣ���ʾҩƷ�����ĵĿ��
    '31936
    Call zlInitȱʡ����
    If InStr(1, mstrPrivs, "��ʾ���") > 0 Then
        If InStr(1, gstr��������ID & ",", "," & lng�ⷿID & ",") > 0 Or gbyt�����ʾ��ʽ <= 0 Then   '31936
                sta.Panels(Pan.C2��ʾ��Ϣ).Text = "[" & strҩƷ & "]���ÿ��:" & dbl���
        Else
                sta.Panels(Pan.C2��ʾ��Ϣ).Text = "[" & strҩƷ & "]���ÿ��:" & IIf(dbl��� > 0, "��", "��") & "���."
        End If
    Else
        sta.Panels(Pan.C2��ʾ��Ϣ).Text = "[" & strҩƷ & "]" & IIf(dbl��� > 0, "��", "��") & "���."
    End If
End Sub

Private Sub Bill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
'���ܣ�����������
    Dim lng��Ŀid As Long, str��� As String, str��׼��Ŀ As String, bln��ʿ As Boolean
    Dim dblStock As Double, strScope As String
    Dim dblPreTime As Double, dblPreMoney As Double
    Dim blnSkip As Boolean, curTotal As Currency, cur��� As Currency
    Dim blnInput As Boolean, strժҪ As String, lngOld���� As Long
    Dim lngDoUnit As Long, lng���˿���ID As String, strҩ��IDs As String, i As Long, j As Long
    Dim colStock As Collection, str�ų���� As String
    Dim dblNum As Double, strPriceGrade As String, lng���� As Long
    
    If KeyCode = 13 And Not Bill.Active Then
        Cancel = True: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    End If
        
    On Error GoTo errH
    
    
    If KeyCode = 13 And Bill.Active Then
        If mbytInState = 2 Then
            If Bill.Col = Bill.COLS - 1 And Bill.Row = Bill.Rows - 1 Then
                Cancel = True: Exit Sub
            ElseIf Bill.TextMatrix(0, Bill.Col) <> "ִ�п���" And Bill.TextMatrix(0, Bill.Col) <> "��ҩҩ��" Then
                Exit Sub
            End If
        End If
        If Bill.ColData(Bill.Col) = BillColType.Text_UnModify Then Exit Sub
        
        '�շ�ʱ,�����Ѳ����޸�
        If (mbytInFun = 0 Or mbytInFun = 1) And mobjBill.Pages(mintPage).Details.Count >= Bill.Row Then
            If mobjBill.Pages(mintPage).Details(Bill.Row).������ Then Exit Sub
        End If
        
 
        Select Case Bill.TextMatrix(0, Bill.Col)
            Case "���"
                Call Clear�����ۼ�
                If Bill.ListIndex <> -1 Then '���������ʱ���ᶨλ�������
                    If Bill.RowData(Bill.Row) <> Bill.ItemData(Bill.ListIndex) Then
                        'һ���ĸ��շ����,�����(����)ԭ�и���Ŀ����
                        For i = 2 To Bill.COLS - 1
                            Bill.TextMatrix(Bill.Row, i) = ""
                        Next
                        If mobjBill.Pages(mintPage).Details.Count >= Bill.Row Then
                            Set mobjBill.Pages(mintPage).Details(Bill.Row).Detail = New Detail
                            Set mobjBill.Pages(mintPage).Details(Bill.Row).InComes = New BillInComes
                            With mobjBill.Pages(mintPage).Details(Bill.Row)
                                .�շ�ϸĿID = 0: .�շ���� = ""
                            End With
                            Call SetItemRowColor(mintPage, Bill.Row)
                            Call CalcMoneys(mintPage)
                            Call ShowMoney(mintPage)
                        End If
                    End If
                    Bill.TextMatrix(Bill.Row, BillCol.���) = Bill.CboText
                    Bill.RowData(Bill.Row) = Bill.ItemData(Bill.ListIndex) '��ʱ��RowData��¼��ѡ����շ����
                End If
            Case "��Ŀ"
            
                '����Ŀȷ��,���շ�ϸĿ��Ӧ�ĳ�����������,ͬʱ���ﴦ���շѴ�����Ŀ
                If Bill.Text <> "" Then
                    '��������������Ŀ�ϰ��س�,��ѡ����ѡ��
                    If mobjBill.Pages(mintPage).Details.Count >= Bill.Row Then
                        'ͨ����ťѡ���Ƿ��ص�ID,�����������ı�,�����һ����,�򲻸ı�
                        If Bill.TextMatrix(Bill.Row, BillCol.��Ŀ) = Bill.Text Then
                            Bill.TxtVisible = False
                            Bill.CmdVisible = False
                            Exit Sub
                        End If
                    End If
                    Call Clear�����ۼ�
                    sta.Panels(Pan.C2��ʾ��Ϣ) = ""
                    sta.Panels("MedicareType").Text = ""
                    blnInput = True
                    If mblnSelect Then
                        mblnSelect = False '��������ñ�־
                        Set mobjDetail = GetInputDetail(Val(Bill.Text), Val(Bill.Tag))
                    Else
                        If gbln�շ���� Then
                            If Bill.RowData(Bill.Row) = 0 Then
                                sta.Panels(Pan.C2��ʾ��Ϣ) = "û��ȷ���������,�����������"
                                Bill.TxtSetFocus: Cancel = True: Exit Sub
                            End If
                            str��� = "'" & Chr(Bill.RowData(Bill.Row)) & "'"
                        Else
                            Call GetOperatorInfo(mobjBill.Pages(mintPage).������, bln��ʿ)
                            str��� = IIf(bln��ʿ, "'E','M','4'", gstr�շ����)
                        End If
                        
                        If mbytInFun = 0 And mstrYBPati <> "" Then
                            '���˺�:24862
                            If zl_Check��׼��Ŀ(gclsInsure, mintInsure, mobjBill.����ID, True) Then str��׼��Ŀ = Get������׼��Ŀ(mobjBill.����ID, "A.ID")
                        End If
                        If zlCheckBill���ڷ�ɢװ��ҩ(mintPage) Then
                            '���ڷ�ɢװ��,�����оͲ��ܽ���¼��
                            Bill.Text = "": Bill.TxtVisible = False
                            Bill.SetFocus: Cancel = True: Exit Sub
                        End If
                        lng���� = -1
                        lng��Ŀid = frmItemSelect.ShowSelect(Me, mstrPrivs, gint������Դ, mintInsure, gblnҩ����λ, _
                            str���, Bill.Text, Bill.TxtHwnd, str��׼��Ŀ, 0, str�ų����, , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ�۸�ȼ�, mbln����ˢ��, lng����)
                        If lng��Ŀid <> 0 Then
                            Set mobjDetail = GetInputDetail(lng��Ŀid, lng����)
                            If mintInsure <> 0 Then sta.Panels("MedicareType").Text = Getҽ������(lng��Ŀid, mintInsure)
                        Else
                            Bill.Text = "": Bill.TxtVisible = False
                            Bill.SetFocus: Cancel = True: Exit Sub
                        End If
                    End If

                    'ȷ�����շ�ϸĿ
                    Bill.TxtVisible = False '(���Ӳ���)
                                            
                    '���ҩƷ�����Ƿ��ظ�:������ʱ��ͬһҩ���������ظ�(����ֻ����)
                    If InStr(",5,6,7,", mobjDetail.���) > 0 _
                        Or (mobjDetail.��� = "4" And mobjDetail.��������) Then
                        If CheckDrugExist(mobjDetail) Then
                            Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                        End If
                    End If
                    
                
                    '��鴦��ְ��
                    If InStr(",5,6,7,", mobjDetail.���) > 0 And mbln����ְ���� Then
                        mobjDetail.����ְ�� = Get����ְ��(mobjDetail.ID)
                        'ҽ���򹫷Ѳ���
                        If cboҽ�Ƹ���.ListIndex <> -1 Then
                            'ҽ���򹫷Ѳ���
                            '����:45605
                            If zlIsCheckMedicinePayMode(zlStr.NeedName(cboҽ�Ƹ���)) Then
                                If CheckDuty(mobjDetail, False) > 0 Then
                                    Bill.TxtSetFocus: Cancel = True: Exit Sub
                                End If
                            End If
                        End If
                        '���в���
                        If CheckDuty(mobjDetail, True) > 0 Then
                            Bill.TxtSetFocus: Cancel = True: Exit Sub
                        End If
                    End If
                    
                    '��ȡҩƷ�����������Ϣ,����ִ�п���ȱʡΪ����,�������ָ����,��Ϊָ������
                    If mobjDetail.��� = "4" Then
                        lngDoUnit = IIf(glng���ϲ��� > 0, glng���ϲ���, mobjBill.����ID)
                    Else
                        lngDoUnit = mobjBill.����ID      '���˿���ID
                    End If
                    If lngDoUnit = 0 Then lngDoUnit = Get��������ID
                                         
                    '���˿���ID
                    lng���˿���ID = mobjBill.����ID
                    If lng���˿���ID = 0 And cbo��������.ListIndex <> -1 Then lng���˿���ID = cbo��������.ItemData(cbo��������.ListIndex)
                    
                    lngDoUnit = Get�շ�ִ�п���ID(mobjDetail.���, mobjDetail.ID, _
                        mobjDetail.ִ�п���, lng���˿���ID, Get��������ID, gint������Դ, _
                        IIf(mlng��ҩ�� = 0, glng��ҩ��, mlng��ҩ��), _
                        IIf(mlng��ҩ�� = 0, glng��ҩ��, mlng��ҩ��), _
                        IIf(mlng��ҩ�� = 0, glng��ҩ��, mlng��ҩ��), _
                        lngDoUnit, mobjBill.����ID)
                                        
                    '��ȡҩƷ�����Ŀ��
                    Call ReadDrugAndStuffStock(lngDoUnit, mobjDetail)
                    
                   
                    
                    '��������
                    If InStr(",5,6,7,", mobjDetail.���) > 0 And mbln����������� Then
                        mobjDetail.�������� = Get��������(mobjDetail.ID)
                    End If
                                        
                    '����֧����Ŀ��Ӧ���
                    If InStr(",5,6,7,", mobjDetail.���) > 0 Then
                        strPriceGrade = mstrҩƷ�۸�ȼ�
                    ElseIf mobjDetail.��� = "4" Then
                        strPriceGrade = mstr���ļ۸�ȼ�
                    Else
                        strPriceGrade = mstr��ͨ�۸�ȼ�
                    End If
                    If mstrYBPati <> "" And Not MCPAR.��������ҽ����Ŀ Then
                        If Not CheckMediCareItem(mobjDetail.ID, mintInsure, mobjDetail.����, mobjDetail.��� = False, strPriceGrade) Then
                            Bill.Text = "": Bill.TxtSetFocus: Cancel = True: Exit Sub
                        End If
                    End If
                    
                    '����ժҪ(ȡ���е����Ա��޸�)
                    If mobjBill.Pages(mintPage).Details.Count >= Bill.Row Then
                        If mobjBill.Pages(mintPage).Details(Bill.Row).Detail.ID = mobjDetail.ID Then
                            strժҪ = mobjBill.Pages(mintPage).Details(Bill.Row).ժҪ
                        End If
                    End If
                    '������㿨����:
'''                    Call zlClear���㿨
                    
                    '������޸ĸ��շ�ϸĿ��
                    Call SetDetail(mobjDetail, Bill.Row, lngDoUnit)
                    '59051
                    '����ժҪ(������������и���ժҪ)
                    If mobjBill.Pages(mintPage).Details(Bill.Row).Detail.����ժҪ Then
                        If frmInputBox.InputBox(Me, "ժҪ", "������""" & mobjBill.Pages(mintPage).Details(Bill.Row).Detail.���� & """��ժҪ��Ϣ:", 200, 3, True, False, strժҪ) Then
                            mobjBill.Pages(mintPage).Details(Bill.Row).ժҪ = strժҪ
                        End If
                    Else 'If mstrYBPati <> "" Then'90304
                         strժҪ = gclsInsure.GetItemInfo(mintInsure, mobjBill.����ID, mobjBill.Pages(mintPage).Details(Bill.Row).�շ�ϸĿID, strժҪ, 1)
                         mobjBill.Pages(mintPage).Details(Bill.Row).ժҪ = strժҪ
                    End If
                    
                    Call CalcMoney(mintPage, Bill.Row)      '��ʱ��û��ȡ������Ŀ
                    
                    'Calcmoney��ҽ�����ܷ���ժҪ
                    If mobjBill.Pages(mintPage).Details(Bill.Row).ժҪ <> "" Then strժҪ = mobjBill.Pages(mintPage).Details(Bill.Row).ժҪ
                    
                    '���ʷ��౨��(���Ѿ�������з��õ�δ��ʾǰ)
                    If mbytInFun = 2 Then
                        If mrsInfo.State = 1 And Not mrsWarn Is Nothing Then
                            curTotal = GetBillSum
                            If mobjBill.Pages(mintPage).Details.Count = Bill.Row And curTotal > 0 Then
                                cur��� = Val(cmdPrint.Tag)
                                If gbln�����������۷��� Then cur��� = cur��� - GetPriceMoneyTotal(0, mrsInfo!����ID) + IIf(mbytBilling = 1, Original.ʵ�պϼ�, 0)
                                gbytWarn = BillingWarn(mstrPrivs, mrsInfo!����, mrsInfo!���ò���, mrsWarn, cur���, mrsInfo!���ն� - Original.ʵ�պϼ�, curTotal, _
                                            IIf(IsNull(mrsInfo!������), 0, mrsInfo!������), mobjBill.Pages(mintPage).Details(Bill.Row).�շ����, mobjBill.Pages(mintPage).Details(Bill.Row).Detail.�������, mstrWarn)
                                If gbytWarn = 2 Or gbytWarn = 3 Then
                                    mobjBill.Pages(mintPage).Details.Remove Bill.Row 'ɾ���ո���Ҫ����ķ�����
                                    Bill.Text = "": Cancel = True: Exit Sub
                                End If
                            End If
                        End If
                    End If
                                        
                    If mbytInFun = 0 And mintInsure <> 0 And MCPAR.ʵʱ��� And mobjBill.Pages(mintPage).Details(Bill.Row).���� <> 0 Then
                        If gclsInsure.CheckItem(mintInsure, 0, 0, MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), 1, 0, mintPage, Bill.Row)) = False Then
                            mobjBill.Pages(mintPage).Details.Remove Bill.Row 'ɾ���ո���Ҫ����ķ�����
                            Bill.Text = "": Cancel = True: Exit Sub
                        End If
                    End If
                    
                    If mobjBill.Pages(mintPage).Details(Bill.Row).���� <> 0 Then
                        If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModul, 0, 0, _
                            MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), mbytInFun, IIf(mbytInFun = 1, 1, IIf(mbytBilling = 0, 0, 1)), mintPage, Bill.Row)) = False Then
                            mobjBill.Pages(mintPage).Details.Remove Bill.Row  'ɾ���ո���Ҫ����ķ�����
                            Bill.Text = "": Bill.TxtVisible = False
                            Cancel = True: Exit Sub
                        End If
                    End If
                          
                    '�����޶���ʾ,����ҩƷҲҪִ��,���ڻָ���Ԫ����ɫ
                    Call SetItemRowColor(mintPage, Bill.Row)
                          
                    Call ShowDetails(Bill.Row)
                    Call ShowMoney(mintPage)
                    
                    '�������ͼ��
                    Call CheckFeeType(Bill.Row)
                    
                    '�������
                    If gcurMaxMoney > 0 Then
                        If Bill.TextMatrix(Bill.Row, BillCol.����) * Bill.TextMatrix(Bill.Row, BillCol.����) * Bill.TextMatrix(Bill.Row, BillCol.����) > gcurMaxMoney Then
                            If MsgBox("��ǰ������" & gcurMaxMoney & ",��ȷ��Ҫ������?", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
                                Call DeleteDetail(Bill.Row, mintPage): Exit Sub
                            End If
                        End If
                    End If
                    Bill.Text = "": Bill.SetFocus
                End If
                
                If mobjBill.Pages(mintPage).Details.Count >= Bill.Row Then
                    mlngPreRow = 0  '�޸�������ʱ,�ָ���ֵ,�Ա���ʾ���
                    With mobjBill.Pages(mintPage).Details(Bill.Row)
                        '��һ�е�����ȷ��
                        If .�շ���� = "7" And gblnPay Then Bill.ColData(BillCol.����) = BillColType.Text  '����
                        If .�շ���� = "F" Then Bill.ColData(BillCol.��־) = BillColType.CheckBox  '���ӱ�־
                        
                        '���������������
                        If .Detail.��� And InStr(",5,6,7,", .�շ����) = 0 _
                            And Not (.�շ���� = "4" And .Detail.��������) Then
                            Bill.ColData(BillCol.����) = IIf(gblnTime, BillColType.Text, BillColType.UnFocus)   '����
                            Bill.ColData(BillCol.����) = BillColType.Text '����
                        Else
                            Bill.ColData(BillCol.����) = BillColType.Text '����
                            Bill.ColData(BillCol.����) = BillColType.UnFocus '����
                        End If
                        
                        'ִ�п���
                        '��FillBillComboBox������ListIndexʱ����CboClick�¼�
                        mblnEnterCell = False: Bill.Col = BillCol.ִ�п���: mblnEnterCell = True
                        Call FillBillComboBox(Bill.Row, BillCol.ִ�п���, Not blnInput) 'ֱ�ӻس�ʱ����ִ�п���
                        mblnEnterCell = False: Bill.Col = BillCol.��Ŀ: mblnEnterCell = True
                        
                        blnSkip = Bill.ListCount = 1
                        If Not blnSkip And InStr(",4,5,6,7,", .�շ����) > 0 Then
                            Select Case .�շ���� 'ָ���˹̶�ҩ�����ϲ���ʱ,��������ѡ��
                                Case "4"
                                    blnSkip = glng���ϲ��� > 0 And .ִ�в���ID = glng���ϲ���
                                Case "5"
                                    blnSkip = glng��ҩ�� > 0 And .ִ�в���ID = glng��ҩ��
                                Case "6"
                                    blnSkip = glng��ҩ�� > 0 And .ִ�в���ID = glng��ҩ��
                                Case "7"
                                    blnSkip = glng��ҩ�� > 0 And .ִ�в���ID = glng��ҩ��
                            End Select
                        End If
                        If blnSkip Then
                            Bill.ColData(BillCol.ִ�п���) = BillColType.UnFocus: .Key = 1
                        Else
                            Bill.ColData(BillCol.ִ�п���) = BillColType.ComboBox: .Key = Bill.ListCount
                        End If
                        
                        If lngDoUnit <> .ִ�в���ID Then
                            '��ȡҩƷ�����Ŀ��
                            Call ReadDrugAndStuffStock(.ִ�в���ID, .Detail)
                        End If
                        '����������ϵ����Ч��,��ȷ��ִ�п���֮��
                        If .�շ���� = "4" And .Detail.�������� Then
                            Call CheckValidity(.�շ�ϸĿID, .ִ�в���ID, .����, False) '��ȷ������,��������
                        End If
                        
                        '������Ŀ����,�������շ���Ŀ�д�����Ŀ����δȡ��ȡ,ҩƷ�����ж�,ҩƷ��������������
                        If Bill.TextMatrix(0, Bill.Col) = "��Ŀ" And InStr(",5,6,7,", .�շ����) = 0 Then
                            If (gbln��������ۿ� And mobjBill.Pages(mintPage).Details(Bill.Row).�������� = 0) Or Not gbln��������ۿ� Then  '(����м���,ֻȡһ��)
                                If CheckHaveChildren(Bill.Row) Then
                                   Call SetSubItem
                                   mlngPreRow = 0 'ͨ���б仯��־������ȷ��������
                                End If
                            End If
                        End If
                    End With
                End If
'                With mobjBill.Pages(mintPage)
'                    If .Details.Count <> 0 And .Details.Count >= Bill.Row And Bill.Active And Visible Then
'                        If .Details(Bill.Row).�շ���� = "7" Then
'                             Call cmd�䷽_Click
'                             Exit Sub
'                        End If
'                    End If
'                End With
'
                '��ҩ,Ĭ��ֻ����һ�θ���
                If mobjBill.Pages(mintPage).Details.Count >= Bill.Row And Bill.Row >= 2 And Bill.Active And Visible Then
                    If mobjBill.Pages(mintPage).Details(Bill.Row).�շ���� = "7" Then
                        For i = 1 To Bill.Row - 1
                            If mobjBill.Pages(mintPage).Details(i).�շ���� = "7" Then
                                '����ִ�иù��̣�����ᶨλ��һ����Ԫ,�ȶ�λ������,����һ����Ԫ������
                                'ѡ����øù��̣����ú���͸��س������ﲻ���ٻس��������������س���Ч��(�ؼ�ԭ��)��
                                Bill.Col = BillCol.����: Exit For
                            End If
                        Next
                    End If
                End If
            Case "����"
                With mobjBill.Pages(mintPage)
                    If .Details.Count >= Bill.Row And Bill.Text <> "" Then
                        '���ֺϷ���
                        If Not IsNumeric(Bill.Text) Then
                            MsgBox "�Ƿ���ֵ��", vbInformation, gstrSysName
                            Bill.Text = .Details(Bill.Row).����: Cancel = True: Exit Sub
                        End If
                        If Val(Bill.Text) <= 0 Or Val(Bill.Text) <> Int(Val(Bill.Text)) Then
                            MsgBox "����Ӧ��Ϊ����������", vbInformation, gstrSysName
                            Bill.Text = .Details(Bill.Row).����: Cancel = True: Exit Sub
                        End If
    
                        '���в�ҩ�ſɸ��ĸ���(һ����ı�,����Ҳ��,��ҩ�����������ӹ�ϵ)
                        If mobjBill.Pages(mintPage).Details(Bill.Row).�շ���� = "7" Then
                            '������ʱ��ҩƷ�����ֹ����(û�з�����ʱ��ҩƷ�����޸ĸ���������)
                            If .Details(Bill.Row).Detail.���� Or .Details(Bill.Row).Detail.��� Then
                                If CSng(Bill.Text) * .Details(Bill.Row).���� > .Details(Bill.Row).Detail.��� Then
                                    MsgBox """" & .Details(Bill.Row).Detail.���� & """Ϊ������ʱ��ҩƷ,��ǰ���ÿ�治������������", vbInformation, gstrSysName
                                    Bill.Text = .Details(Bill.Row).����: Cancel = True: Exit Sub
                                End If
                            End If
                                  
                            '�������ʱ�ۻ������ҩ���ĸ��������Ƿ��㹻
                            For i = 1 To .Details.Count
                                If i <> Bill.Row And .Details(i).�շ���� = "7" And (.Details(i).Detail.��� Or .Details(i).Detail.����) Then
                                    If Val(Bill.Text) * .Details(i).���� > .Details(i).Detail.��� Then
                                        MsgBox "�� " & i & " ��ҩƷ""" & .Details(i).Detail.���� & """Ϊ������ʱ��ҩƷ,��ǰ���ÿ�治������������", vbInformation, gstrSysName
                                        Bill.Text = .Details(Bill.Row).����: Cancel = True: Exit Sub
                                    End If
                                End If
                            Next
                            '�������
                            If gcurMaxMoney > 0 Then
                                If CSng(Bill.Text) * .Details(Bill.Row).���� * Bill.TextMatrix(Bill.Row, BillCol.����) > gcurMaxMoney Then
                                    If MsgBox("��ǰ������" & gcurMaxMoney & ",��ȷ��Ҫ������?", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
                                        Bill.Text = .Details(Bill.Row).����: Cancel = True: Exit Sub
                                    End If
                                End If
                            End If
                            lngOld���� = .Details(Bill.Row).����
                            '������㿨����:
'''                            Call zlClear���㿨
                            '���㲢ˢ�¸���
                            .Details(Bill.Row).���� = Bill.Text
                            Call CalcMoneys(mintPage, Bill.Row)

                            '���������ٸĸ����ģ����������¼�飬���丶�������������ģ�������������
                            If mbytInFun = 0 And mintInsure <> 0 And MCPAR.ʵʱ��� And .Details(Bill.Row).���� <> 0 Then
                                If gclsInsure.CheckItem(mintInsure, 0, 0, MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), 1, 0, mintPage, Bill.Row)) = False Then
                                    .Details(Bill.Row).���� = lngOld����
                                    Call CalcMoneys(mintPage, Bill.Row)
                                    Bill.Text = "": Bill.TxtVisible = False
                                    Cancel = True: Exit Sub
                                End If
                            End If
                            
                            If .Details(Bill.Row).���� <> 0 Then
                                If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModul, 0, 0, _
                                    MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), mbytInFun, IIf(mbytInFun = 1, 1, IIf(mbytBilling = 0, 0, 1)), mintPage, Bill.Row)) = False Then
                                    .Details(Bill.Row).���� = lngOld����
                                    Call CalcMoneys(mintPage, Bill.Row)
                                    Bill.Text = "": Bill.TxtVisible = False
                                    Cancel = True: Exit Sub
                                End If
                            End If
                            
                            Call ShowDetails(Bill.Row)
                            
                            '����������ҩ����,����Ƕ�����,���޸������Ǵ����,����Ǵ���,���޸�ͬһ����Ĵ����.��Ϊ�޶�Ϊ�в�ҩ,������������
                            For i = 1 To .Details.Count
                                If i <> Bill.Row And .Details(i).�շ���� = "7" And .Details(i).�������� = .Details(Bill.Row).�������� Then
                                    If .Details(i).�������� = 0 Or (.Details(i).�������� <> 0 And .Details(i).Detail.���д��� = 0) Then     '1��2�̶��Ͱ������Ĳ���
                                        .Details(i).���� = Bill.Text
                                        Call CalcMoneys(mintPage, i)
                                        Call ShowDetails(i)
                                    End If
                                End If
                            Next
                                                        
                            Call ShowMoney(mintPage)
                        Else
                            sta.Panels(Pan.C2��ʾ��Ϣ) = "������Ŀ�ĸ������ܸ��ģ�"
                            Bill.Text = .Details(Bill.Row).���� '�ָ�ԭ�и���ֵ
                        End If
                    End If
                End With
            Case "����"
                With mobjBill.Pages(mintPage)
                If .Details.Count >= Bill.Row And Bill.Text <> "" Then
                    '�������ת��
                    If InStr(",7,", .Details(Bill.Row).�շ����) > 0 Then Bill.Text = ConvertABCtoNUM(Bill.Text)
                
                    '���ֺϷ���
                    If Not IsNumeric(Bill.Text) Then
                        MsgBox "�Ƿ���ֵ��", vbInformation, gstrSysName
                        Bill.Text = .Details(Bill.Row).����: Cancel = True: Exit Sub
                    End If
                    If Val(Bill.Text) = 0 Then
                        If MsgBox("��������Ϊ�㣬Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Bill.Text = .Details(Bill.Row).����: Cancel = True: Exit Sub
                        End If
                    End If
                    'ҩƷ����С��
                    If InStr(",5,6,7,", .Details(Bill.Row).�շ����) > 0 Then
                        If Val(Bill.Text) - Int(Val(Bill.Text)) <> 0 And InStr(mstrPrivs, "ҩƷ����С��") = 0 Then
                            MsgBox "��û��Ȩ������С����", vbInformation, gstrSysName
                            Bill.Text = .Details(Bill.Row).����: Cancel = True: Exit Sub
                        End If
                    End If
                    
                    If InStr(",5,6,7,", .Details(Bill.Row).�շ����) > 0 And gblnҩ����λ Then
                        dblNum = Val(Bill.Text) * .Details(Bill.Row).���� * .Details(Bill.Row).Detail.ҩ����װ
                    Else
                        dblNum = Val(Bill.Text) * .Details(Bill.Row).����
                    End If
                                            
                    '�����Ϸ��Լ��
                    If CSng(Bill.Text) * .Details(Bill.Row).���� < 0 Then
                        'Ȩ��
                        If (mbytInFun < 2 And InStr(mstrPrivs, "��������") = 0 Or mbytInFun = 2 And InStr(mstrPrivs, "��������") = 0) Then
                            MsgBox "��û��Ȩ�����븺����", vbInformation, gstrSysName
                            Bill.Text = .Details(Bill.Row).����: Cancel = True: Exit Sub
                        ElseIf .Details(Bill.Row).Detail.���� Then
                            MsgBox "����ҩƷ���������븺����", vbInformation, gstrSysName
                            Bill.Text = .Details(Bill.Row).����: Cancel = True: Exit Sub
                        End If
                        If mbytInFun = 2 Then
                            '���������������:ֻ���������۲��˲Ż���м���Ƿ����
                            '����:36558
                             If Not mrsInfo Is Nothing Then
                                If mrsInfo.State = 1 Then
                                     If Nvl(mrsInfo!����, 0) = 1 Then
                                        If Not CheckNegative(mobjBill.����ID, mobjBill.��ҳID, .Details(Bill.Row).�շ�ϸĿID, .Details(Bill.Row).ִ�в���ID, dblNum, .Details(Bill.Row).Detail.ҩ����װ, mstrPrivs, Format(mrsInfo!��Ժ����, "yyyy-mm-dd HH:MM:SS")) Then
                                            Bill.Text = .Details(Bill.Row).����: Cancel = True: Exit Sub
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                    
                    '�������
                    If gcurMaxMoney > 0 Then
                        If CSng(Bill.Text) * .Details(Bill.Row).���� * Bill.TextMatrix(Bill.Row, BillCol.����) > gcurMaxMoney Then
                            If MsgBox("��ǰ������" & gcurMaxMoney & ",��ȷ��Ҫ������?", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
                                Bill.Text = .Details(Bill.Row).����: Cancel = True: Exit Sub
                            End If
                        End If
                    End If
                    
                          
                    Bill.Text = FormatEx(Bill.Text, 5)
                          
                    'ҩƷ�����
                    With .Details(Bill.Row)
                        If (.�շ���� = "4" And .Detail.��������) Or InStr(",5,6,7,", .�շ����) > 0 Then
                            If .Detail.���� Or .Detail.��� Then
                                If .���� * CSng(Bill.Text) > .Detail.��� Then '������ʱ��ҩƷ�����ֹ����
                                    If .�շ���� = "4" Then
                                        MsgBox """" & .Detail.���� & """Ϊ������ʱ����������,��ǰ���ÿ�治������������", vbInformation, gstrSysName
                                    Else
                                        MsgBox """" & .Detail.���� & """Ϊ������ʱ��ҩƷ,��ǰ���ÿ�治������������", vbInformation, gstrSysName
                                    End If
                                    Bill.Text = .����: Cancel = True: Exit Sub
                                End If
                            Else
                                Set colStock = IIf(.�շ���� = "4", mcolStock2, mcolStock1)
                                
                                If colStock("_" & .ִ�в���ID) <> 0 And InStr(mstrPrivs, "�������") = 0 And Bill.ColData(BillCol.ִ�п���) = BillColType.UnFocus Then
                                    If .���� * CSng(Bill.Text) > .Detail.��� Then '����ҩƷ�������
                                        If colStock("_" & .ִ�в���ID) = 1 Then
                                            If MsgBox("""" & .Detail.���� & """�ĵ�ǰ���ÿ�治����������,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                                Bill.Text = .����: Cancel = True: Exit Sub
                                            End If
                                        ElseIf colStock("_" & .ִ�в���ID) = 2 Then
                                            MsgBox """" & .Detail.���� & """�ĵ�ǰ���ÿ�治������������", vbInformation, gstrSysName
                                            Bill.Text = .����: Cancel = True: Exit Sub
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End With
                    
                    dblPreTime = .Details(Bill.Row).����
                    .Details(Bill.Row).���� = Bill.Text
                    
                    '�����������
                    If mbln����������� And Not gbln�������� Then
                        If Not CheckLimit(mobjBill, mintPage, Bill.Row) Then
                            .Details(Bill.Row).���� = dblPreTime: Bill.Text = dblPreTime
                            Cancel = True: Exit Sub
                        End If
                    End If
                    If .Details(Bill.Row).Detail.¼������ > 0 And .Details(Bill.Row).���� > .Details(Bill.Row).Detail.¼������ Then
                        If MsgBox("��������γ�����¼������" & .Details(Bill.Row).Detail.¼������ & ",�Ƿ����?", vbDefaultButton2 + vbYesNo + vbQuestion, gstrSysName) = vbNo Then
                            .Details(Bill.Row).���� = dblPreTime: Bill.Text = dblPreTime
                            Cancel = True: Exit Sub
                        End If
                    End If
                    
                    
                    '���д������ܸ�������(����Ŀ���θı�,���д���������Ҳ��)
                    If .Details(Bill.Row).�������� <> 0 And .Details(Bill.Row).Detail.���д��� <> 0 Then
                        sta.Panels(Pan.C2��ʾ��Ϣ) = "����Ŀ�ǹ��д�����Ŀ,�����β��ܹ����ġ�"
                        .Details(Bill.Row).���� = dblPreTime: Bill.Text = dblPreTime
                        Exit Sub
                    End If
                    '������㿨����:
'''                    Call zlClear���㿨
                    
                    Call CalcMoneys(mintPage, Bill.Row)
                    
                    '����������(���Ѿ�������з��õ�δ��ʾǰ)
                    If MoneyOverFlow(mobjBill) Then
                        MsgBox "�����������µ��ݽ����������ʵ�������", vbInformation, gstrSysName
                        .Details(Bill.Row).���� = dblPreTime
                        Bill.Text = ""
                        Call CalcMoneys(mintPage, Bill.Row)
                        Cancel = True: Bill.TxtVisible = False: Exit Sub
                    End If
                    
                    If mbytInFun = 0 And mintInsure <> 0 And MCPAR.ʵʱ��� And .Details(Bill.Row).���� <> 0 Then
                        If gclsInsure.CheckItem(mintInsure, 0, 0, MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), 1, 0, mintPage, Bill.Row)) = False Then
                            .Details(Bill.Row).���� = dblPreTime
                            Call CalcMoneys(mintPage, Bill.Row)
                            Bill.Text = "": Bill.TxtVisible = False
                            Cancel = True::  Exit Sub
                        End If
                    End If
                    
                    If .Details(Bill.Row).���� <> 0 Then
                        If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModul, 0, 0, _
                            MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), mbytInFun, IIf(mbytInFun = 1, 1, IIf(mbytBilling = 0, 0, 1)), mintPage, Bill.Row)) = False Then
                            .Details(Bill.Row).���� = dblPreTime
                            Call CalcMoneys(mintPage, Bill.Row)
                            Bill.Text = "": Bill.TxtVisible = False
                            Cancel = True: Exit Sub
                        End If
                    End If
                    
                    '���ʷ��౨��(���Ѿ�������з��õ�δ��ʾǰ)
                    If mbytInFun = 2 Then
                        If mrsInfo.State = 1 And Not mrsWarn Is Nothing Then
                            curTotal = GetBillSum
                            If curTotal > 0 Then
                                cur��� = Val(cmdPrint.Tag)
                                If gbln�����������۷��� Then cur��� = cur��� - GetPriceMoneyTotal(0, mrsInfo!����ID) + IIf(mbytBilling = 1, Original.ʵ�պϼ�, 0)
                                gbytWarn = BillingWarn(mstrPrivs, mrsInfo!����, mrsInfo!���ò���, mrsWarn, cur���, mrsInfo!���ն� - Original.ʵ�պϼ�, _
                                            curTotal, IIf(IsNull(mrsInfo!������), 0, mrsInfo!������), .Details(Bill.Row).�շ����, .Details(Bill.Row).Detail.�������, mstrWarn)
                                If gbytWarn = 2 Or gbytWarn = 3 Then
                                    .Details(Bill.Row).���� = dblPreTime
                                    Bill.Text = ""
                                    Call CalcMoneys(mintPage, Bill.Row)
                                    Cancel = True: Bill.TxtVisible = False: Exit Sub
                                End If
                            End If
                        End If
                    End If
                    
                    Call ShowDetails(Bill.Row)
                    
                    '��������д���������(ҩƷû�д�����Ŀ)
                    If .Details(Bill.Row).�������� = 0 Then
                        For i = Bill.Row + 1 To .Details.Count
                            If .Details(i).�������� = Bill.Row Then
                                '28136
                                '���������ĸ���,��Ҫ���¼��еĸ������и��³ɸ���
                                With .Details(i)
                                    If .Detail.���д��� = 0 Then  '�ǹ��д���
                                        If Abs(.����) <> Abs(.Detail.��������) Then GoTo NotCalc:
                                        .���� = IIf(Val(Bill.Text) < 0, -1, 1) * .Detail.��������
                                    ElseIf .Detail.���д��� = 1 Then '�̶��Ĺ��д���
                                        .���� = IIf(Val(Bill.Text) < 0, -1, 1) * IIf(.Detail.�������� = 0, 1, .Detail.��������)
                                    ElseIf .Detail.���д��� = 2 Then   '�������Ĺ��д���
                                        .���� = Val(Bill.Text) * .Detail.��������
                                    Else
                                         GoTo NotCalc:
                                    End If
                                End With
                                Call CalcMoneys(mintPage, i)
                                Call ShowDetails(i)
NotCalc:
                            End If
                        Next
                    End If
                    
                    Call ShowMoney(mintPage)
                    
                ElseIf .Details.Count >= Bill.Row Then
                    If Val(Bill.TextMatrix(Bill.Row, Bill.Col)) = 0 Then
                        If MsgBox("��������Ϊ�㣬Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Cancel = True: Exit Sub
                        End If
                    End If
                End If
                If Bill.ColData(BillCol.ִ�п���) = BillColType.UnFocus Then
                    If CheckMainItem(Bill.Row) Then
                        KeyCode = 0
                        Call LocateMainItemNextRow(Bill.Row)
                    End If
                End If
                End With
            Case "����"
                With mobjBill.Pages(mintPage)
                If .Details.Count >= Bill.Row And Bill.Text <> "" Then
                    '���ֺϷ���
                    If Not IsNumeric(Bill.Text) Then
                        MsgBox "�Ƿ���ֵ��", vbInformation, gstrSysName
                        Bill.Text = "": Cancel = True: Bill.TxtVisible = False: Exit Sub
                    End If
                    If Val(Bill.Text) < 0 Then
                        MsgBox "��Ŀ�۸�Ӧ��Ϊ������Ҫ�˷ѿ������븺��������ʵ�֣�", vbInformation, gstrSysName
                        Bill.Text = "": Cancel = True: Bill.TxtVisible = False: Exit Sub
                    End If
                    '�������
                    If gcurMaxMoney > 0 Then
                        If CSng(Bill.Text) * .Details(Bill.Row).���� * .Details(Bill.Row).���� > gcurMaxMoney Then
                            If MsgBox("��ǰ������" & gcurMaxMoney & ",��ȷ��Ҫ������?", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
                                Bill.Text = "": Cancel = True: Exit Sub
                            End If
                        End If
                    End If
                    
                    Bill.Text = FormatEx(Bill.Text, 5)
                    
                    '���û�ж�Ӧ��������Ŀ,���޷�����
                    If .Details(Bill.Row).Detail.��� And .Details(Bill.Row).InComes.Count > 0 Then
                        If Not (.Details(Bill.Row).InComes(1).�ּ� = 0 And .Details(Bill.Row).InComes(1).ԭ�� = 0) Then
                            strScope = CheckScope(.Details(Bill.Row).InComes(1).ԭ��, .Details(Bill.Row).InComes(1).�ּ�, CCur(Bill.Text))
                            If strScope <> "" Then
                                sta.Panels(Pan.C2��ʾ��Ϣ) = strScope
                                If Bill.TxtVisible And Len(Bill.Text) > 9 Then Bill.Text = .Details(Bill.Row).InComes(1).��׼����
                                If Bill.TxtVisible Then Bill.SelStart = 0: Bill.SelLength = Len(Bill.Text)
                                Cancel = True: Beep: Exit Sub
                            End If
                        End If
                        '������㿨����:
'''                        Call zlClear���㿨
                        
                        dblPreMoney = .Details(Bill.Row).InComes(1).��׼����
                                                
                        .Details(Bill.Row).InComes(1).��׼���� = Bill.Text '�����շ�ϸĿֻ�ܶ�Ӧһ��������Ŀ
                        Call CalcMoneys(mintPage, Bill.Row)
                        
                        '���ʷ��౨��(���Ѿ�������з��õ�δ��ʾǰ)
                        If mbytInFun = 2 Then
                            If mrsInfo.State = 1 And Not mrsWarn Is Nothing Then
                                curTotal = GetBillSum
                                If curTotal > 0 Then
                                    cur��� = Val(cmdPrint.Tag)
                                    If gbln�����������۷��� Then cur��� = cur��� - GetPriceMoneyTotal(0, mrsInfo!����ID) + IIf(mbytBilling = 1, Original.ʵ�պϼ�, 0)
                                    gbytWarn = BillingWarn(mstrPrivs, mrsInfo!����, mrsInfo!���ò���, mrsWarn, cur���, mrsInfo!���ն� - Original.ʵ�պϼ�, curTotal, _
                                                IIf(IsNull(mrsInfo!������), 0, mrsInfo!������), .Details(Bill.Row).�շ����, .Details(Bill.Row).Detail.�������, mstrWarn)
                                    If gbytWarn = 2 Or gbytWarn = 3 Then
                                        .Details(Bill.Row).InComes(1).��׼���� = dblPreMoney
                                        Bill.Text = ""
                                        Call CalcMoneys(mintPage, Bill.Row)
                                        Cancel = True: Bill.TxtVisible = False: Exit Sub
                                    End If
                                End If
                            End If
                        End If
                        
                        Call ShowDetails(Bill.Row)
                        Call ShowMoney(mintPage)
                    Else
                        Bill.Text = "0"
                        sta.Panels(Pan.C2��ʾ��Ϣ) = "����Ŀ�������ö�Ӧ�ķ�Ŀ�������޷�������ã�"
                    End If
                End If
                End With
            Case "ִ�п���", "��ҩҩ��"
                If mobjBill.Pages(mintPage).Details.Count >= Bill.Row And Bill.ListIndex <> -1 Then
                    With mobjBill.Pages(mintPage).Details(Bill.Row)
                        If .ִ�в���ID <> Bill.ItemData(Bill.ListIndex) Then    'cbo_click���п��ܻ�ִ��һ��
                             .ִ�в���ID = Bill.ItemData(Bill.ListIndex)
                            If CheckMainItem(Bill.Row) Then Call SetSubItemDept(Bill.Row) '������ڴ���,��ı��ҩƷ�е�ִ�п���
                        End If
                
                        'ҩƷ�����:��̬ҩ��,������ʱ��ҩƷҲҪ�����
                        If (.�շ���� = "4" And .Detail.��������) Or InStr(",5,6,7,", .�շ����) > 0 Then
                            If .Detail.���� Or .Detail.��� Then '������ʱ��ҩƷ��治���ֹ����
                                If .���� * .���� > .Detail.��� Then
                                    If .�շ���� = "4" Then
                                        MsgBox "[" & .Detail.���� & "]Ϊ������ʱ����������,��ǰ���ÿ�治������������", vbInformation, gstrSysName
                                    Else
                                        MsgBox "[" & .Detail.���� & "]Ϊ������ʱ��ҩƷ,��ǰ���ÿ�治������������", vbInformation, gstrSysName
                                    End If
                                    Cancel = True
                                End If
                            Else
                                Set colStock = IIf(.�շ���� = "4", mcolStock2, mcolStock1)
                                
                                If colStock("_" & .ִ�в���ID) <> 0 And InStr(mstrPrivs, "�������") = 0 Then
                                    If .���� * .���� > .Detail.��� Then
                                        If colStock("_" & .ִ�в���ID) = 1 Then
                                            If MsgBox("[" & .Detail.���� & "]�ĵ�ǰ���ÿ�治����������,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                                Cancel = True
                                            End If
                                        ElseIf colStock("_" & .ִ�в���ID) = 2 Then
                                            MsgBox "[" & .Detail.���� & "]�ĵ�ǰ���ÿ�治������������", vbInformation, gstrSysName
                                            Cancel = True
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        
                        '����������ϵ����Ч��,��ȷ��ִ�п���֮��
                        If .�շ���� = "4" And .Detail.�������� Then
                            Call CheckValidity(.�շ�ϸĿID, .ִ�в���ID, .����, False) '��ȷ������,��������
                        End If
                        
                        If CheckMainItem(Bill.Row) Then
                            KeyCode = 0
                            Call LocateMainItemNextRow(Bill.Row)
                        End If
                        If Bill.TextMatrix(0, Bill.Col) = "ִ�п���" Then
                            If mbytInFun = 0 And mintInsure <> 0 And MCPAR.ʵʱ��� And mobjBill.Pages(mintPage).Details(Bill.Row).���� <> 0 Then
                                If gclsInsure.CheckItem(mintInsure, 0, 0, MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), 1, 0, mintPage, Bill.Row)) = False Then
                                    Bill.Text = "": Bill.TxtVisible = False
                                    Cancel = True: Exit Sub
                                End If
                            End If
                            
                            If mobjBill.Pages(mintPage).Details(Bill.Row).���� <> 0 Then
                                If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModul, 0, 0, _
                                    MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), mbytInFun, IIf(mbytInFun = 1, 1, IIf(mbytBilling = 0, 0, 1)), mintPage, Bill.Row)) = False Then
                                    Bill.Text = "": Bill.TxtVisible = False
                                    Cancel = True: Exit Sub
                                End If
                            End If
                        End If
                    End With
                End If
        End Select
        
        '��Ҫ����Ԥ����
        If InStr(",���,��Ŀ,����,����,����,", "," & Bill.TextMatrix(0, Bill.Col) & ",") > 0 Then
            If cmdԤ����.Visible Then
                Call InitBalanceGrid
                cmdԤ����.TabStop = True
                cmdOK.Enabled = False
            End If
          '  Call zlClear���㿨
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Cancel = True
End Sub


Private Sub LocateMainItemNextRow(ByVal lngRow As Long)
    Dim i As Long
    
    For i = lngRow + 1 To mobjBill.Pages(mintPage).Details.Count
        If mobjBill.Pages(mintPage).Details(i).�������� = lngRow Then
            If mobjBill.Pages(mintPage).Details(i).Detail.���д��� = 0 Then Exit For
        End If
    Next
    
    If i <= mobjBill.Pages(mintPage).Details.Count Then
        Bill.Col = BillCol.����
        Bill.Row = i: Bill.MsfObj.TopRow = i
    Else
        Call LocateNewRow
    End If
End Sub

Private Sub LocateNewRow()
    If mobjBill.Pages(mintPage).Details.Count >= Bill.Rows - 1 Then
        Bill.Rows = Bill.Rows + 1
        mblnNewRow = True
        Call bill_AfterAddRow(Bill.Rows - 1)
        mblnNewRow = False
        Bill.Row = Bill.Rows - 1
        Bill.MsfObj.TopRow = Bill.Row
        Bill.Col = BillCol.���
    Else
        Bill.Row = Bill.Rows - 1
        Bill.MsfObj.TopRow = Bill.Row
        Bill.Col = BillCol.���
    End If
    '����:27792
    If Not Me.ActiveControl Is Bill Then
        If Bill.Active And Bill.Visible Then Bill.SetFocus
    End If
End Sub

Private Sub SetSubItem()
'����:�����շ���Ŀ��,���ص�ǰ�շ���Ŀ�Ĵ�����Ŀ�����ü�����,����ʾ�ڵ��ݿؼ���
'����:
'������:Bill_KeyDown��������Ŀ��
Dim i As Integer, j As Integer, lngMainRow As Long
Dim lngDoUnit As Long, lng���˿���ID As Long
Dim bln��������ۿ� As Boolean
Dim strժҪ As String, strPriceGrade As String

lngMainRow = Bill.Row               '�������
If gbln��������ۿ� Then            '����������ηѱ�,����ܼ����ۿ۲�����Ч,�����ܼ���
    bln��������ۿ� = Not mobjBill.Pages(mintPage).Details(lngMainRow).Detail.���ηѱ�
End If

lng���˿���ID = mobjBill.����ID
If lng���˿���ID = 0 And cbo��������.ListIndex <> -1 Then lng���˿���ID = cbo��������.ItemData(cbo��������.ListIndex)


With mobjBill.Pages(mintPage).Details(lngMainRow)
    Set mcolDetails = GetSubDetails(.�շ�ϸĿID)
    For i = 1 To mcolDetails.Count
        If mobjBill.Pages(mintPage).Details.Count >= Bill.Rows - 1 Then
            Bill.Rows = Bill.Rows + 1
            mblnNewRow = True
            Call bill_AfterAddRow(Bill.Rows - 1)    '��������
             mblnNewRow = False
        End If
        Bill.TextMatrix(Bill.Rows - 1, BillCol.���) = ""  '�б�Ҫ����
        
        'a.������ĿΪ��ҩƷ��Ŀ��ִ�п���
        lngDoUnit = 0
        If InStr(",4,5,6,7,", mcolDetails(i).���) = 0 Then
             If mcolDetails(i).��� = .�շ���� Or mcolDetails(i).ִ�п��� = 0 Then
                '1.�����շ������������ͬ��,ȱʡ������ִ�п�����ͬ��
                '2.��������Ϊ����ȷ���ҵ�,ȱʡ������ִ�п�����ͬ��
                lngDoUnit = .ִ�в���ID
             Else
                '������ҩ��Ŀ��ִ�п���
                lngDoUnit = Get�շ�ִ�п���ID(mcolDetails(i).���, mcolDetails(i).ID, _
                    mcolDetails(i).ִ�п���, lng���˿���ID, Get��������ID, gint������Դ, , , , , mobjBill.����ID)
             End If
        'b.������ĿΪҩƷ,���ĵ�ִ�п���
        Else
            lngDoUnit = Get�շ�ִ�п���ID(mcolDetails(i).���, mcolDetails(i).ID, mcolDetails(i).ִ�п���, lng���˿���ID, Get��������ID, gint������Դ, _
                IIf(mlng��ҩ�� = 0, glng��ҩ��, mlng��ҩ��), IIf(mlng��ҩ�� = 0, glng��ҩ��, mlng��ҩ��), _
                IIf(mlng��ҩ�� = 0, glng��ҩ��, mlng��ҩ��), .ִ�в���ID, mobjBill.����ID)  '���Ĵ���ȱʡ������ִ�п�����ͬ
        End If
        '����֧����Ŀ��Ӧ���
        If InStr(",5,6,7,", mcolDetails(i).���) > 0 Then
            strPriceGrade = mstrҩƷ�۸�ȼ�
        ElseIf mcolDetails(i).��� = "4" Then
            strPriceGrade = mstr���ļ۸�ȼ�
        Else
            strPriceGrade = mstr��ͨ�۸�ȼ�
        End If
        If mstrYBPati <> "" And Not MCPAR.��������ҽ����Ŀ Then
            If Not CheckMediCareItem(mcolDetails(i).ID, mintInsure, mcolDetails(i).����, mcolDetails(i).��� = False, strPriceGrade) Then
                Exit Sub
            End If
        End If
        
        Call SetDetail(mcolDetails(i), Bill.Rows - 1, lngDoUnit, Bill.Row)
                
        Call CalcMoney(mintPage, Bill.Rows - 1, bln��������ۿ�)
        Call ShowDetails(Bill.Rows - 1, i, mcolDetails.Count)
                
'        If mstrYBPati <> "" Then'90304
            'CalcMoney���ȵ���GetuItemInsure���ܷ���ժҪ
            strժҪ = mobjBill.Pages(mintPage).Details(Bill.Rows - 1).ժҪ
             
            strժҪ = gclsInsure.GetItemInfo(mintInsure, mobjBill.����ID, mcolDetails(i).ID, strժҪ, 1)
            mobjBill.Pages(mintPage).Details(Bill.Rows - 1).ժҪ = strժҪ
'        End If
    Next
        
    If bln��������ۿ� Then
        Call CalcMoney(mintPage, lngMainRow, bln��������ۿ�) '�����������Ӧ����ʵ��,��Ϊ��û�м������ǰ����ȷ���㲻��
        
        Call CalcPItemActualIncome(lngMainRow)
    End If
    
    Call ShowMoney(mintPage)
    
    'һ���Ե����ײ���ĿLED��ʾ
    Call ShowGroupLED(Bill.Row, Bill.Rows - mcolDetails.Count, Bill.Rows - 1)
    
End With

End Sub

Private Sub CalcPItemActualIncome(ByVal lngMainRow As Long, Optional intPage As Integer)
'����:����������ۿ�ʱ,����ָ�����������ID�ĵ�һ��������Ŀ���������ʵ�ս��
'����:  lngMainRow-������ID
'       intpage -ҳ��,Ĭ��Ϊ��ǰҳmintpage

Dim i As Long, j As Long
Dim cur����ǰӦ�պϼ� As Currency     '��¼�����������Ӧ�պϼ�
Dim cur���ۺ�ʵ�� As Currency
Dim str�ѱ� As String               '��¼����Ӧ�յ�ȷ�������Żݵķѱ�
If intPage = 0 Then intPage = mintPage

With mobjBill.Pages(intPage)
    For i = lngMainRow To .Details.Count
        'If i <> lngMainRow And .Details(i).�������� <> lngMainRow Then Exit For    '��ȻĿǰ�����˲������ڴ����м������������,����һ�ŵ�����������,Ϊ�˽������ܵ�����,����ȫ��ɨ��
        
        If i = lngMainRow Or .Details(i).�������� = lngMainRow Then
            For j = 1 To .Details(i).InComes.Count
                cur����ǰӦ�պϼ� = cur����ǰӦ�պϼ� + .Details(i).InComes(j).Ӧ�ս��
            Next
        End If
    Next
    'ҩƷ��֧��������������贫�Ӱ�Ӽ��ʵ�
    '���ۺ��ʵ�ս����㵽����ĵ�һ��������Ŀ��
    str�ѱ� = IIf(glngSys Like "8??" Or mbytInFun = 2, mobjBill.�ѱ�, zlStr.TrimEx(mobjBill.�ѱ� & "," & lbl��̬�ѱ�.Tag, ","))
    
    cur���ۺ�ʵ�� = CCur(Format(ActualMoney(str�ѱ�, .Details(lngMainRow).InComes(1).������ĿID, cur����ǰӦ�պϼ�, 0, 0, 0, 0), gstrDec))
    cur���ۺ�ʵ�� = cur���ۺ�ʵ�� - cur����ǰӦ�պϼ� + .Details(lngMainRow).InComes(1).Ӧ�ս��
    
    .Details(lngMainRow).InComes(1).ʵ�ս�� = Format(cur���ۺ�ʵ��, gstrDec)
    .Details(lngMainRow).�ѱ� = str�ѱ�
    
    Call ShowDetails(lngMainRow)
End With
End Sub

Private Sub SetSubItemDept(ByVal lngRow As Long)
'����:��������ִ�п��ҵı仯,ˢ�·�ҩ�����ִ�п���

    Dim i As Long, j As Long, lng���˿���ID As Long
    
    With mobjBill.Pages(mintPage)
        '��ȡ���д����ִ�п�������,������ȡ(��Ϊ�����ϵĴ�����Ϣ�������޸Ĺ���)
        Set mcolDetails = GetSubDetails(.Details(lngRow).�շ�ϸĿID)
        
        lng���˿���ID = mobjBill.����ID
        If lng���˿���ID = 0 And cbo��������.ListIndex <> -1 Then lng���˿���ID = cbo��������.ItemData(cbo��������.ListIndex)
        
        For i = lngRow + 1 To .Details.Count
            If .Details(i).�������� = lngRow Then
                '������ΪҩƷ�����ĵ���Ŀ��ִ�п��Ҳ�������䶯
                If InStr(",4,5,6,7,", .Details(i).�շ����) = 0 Then
                    If .Details(i).�շ���� = .Details(lngRow).�շ���� Then
                        '1.�����շ������������ͬ��,ȱʡ������ִ�п�����ͬ��
                        .Details(i).ִ�в���ID = .Details(lngRow).ִ�в���ID
                    Else
                        For j = 1 To mcolDetails.Count
                            If mcolDetails.Item(j).ID = .Details(i).Detail.ID Then
                                Exit For
                            End If
                        Next
                        If j <= mcolDetails.Count Then
                            If mcolDetails.Item(j).ִ�п��� = 0 Then
                                '2.��������Ϊ����ȷ���ҵ�,ȱʡ������ִ�п�����ͬ��
                                 .Details(i).ִ�в���ID = .Details(lngRow).ִ�в���ID
                            Else
                                '3.������ҩ��Ŀ��ִ�п���
                                .Details(i).ִ�в���ID = Get�շ�ִ�п���ID(mcolDetails(j).���, mcolDetails(j).ID, _
                                    mcolDetails(j).ִ�п���, lng���˿���ID, Get��������ID, gint������Դ, , , , , mobjBill.����ID)
                            End If
                        End If
                    End If
                    
                    'ˢ����ʾ����ִ�п���
                    If .Details(i).ִ�в���ID <> 0 Then
                        If mbytInState = 0 Then
                            mrsUnit.Filter = "ID=" & .Details(i).ִ�в���ID
                            If mrsUnit.RecordCount <> 0 Then
                                Bill.TextMatrix(i, BillCol.ִ�п���) = IIf(zlIsShowDeptCode, mrsUnit!���� & "-", "") & mrsUnit!����
                            Else
                                Bill.TextMatrix(i, BillCol.ִ�п���) = GET��������(.Details(i).ִ�в���ID, mrsUnit)
                            End If
                        Else
                            '�������ֻ(��)��ʾ����
                            Bill.TextMatrix(i, BillCol.ִ�п���) = GET��������(.Details(i).ִ�в���ID, mrsUnit)
                        End If
                    Else
                        Bill.TextMatrix(i, BillCol.ִ�п���) = ""
                    End If
                End If
            End If
        Next
    End With
End Sub

Private Sub Bill_EnterCell(Row As Long, Col As Long)
    Dim i As Integer, bln������ As Boolean
    
    Dim strStock As String, strTmp As String
    Dim strҩ��IDs As String
    
    If Not mblnEnterCell Then Exit Sub
    
    If Bill.ColData(Col) = BillColType.UnFocus Then Exit Sub
    If mbytInState = 3 Or (chkCancel.Visible And chkCancel.Value = 1) Then
        '����б༭����������ɫ
        Bill.SetColColor BillCol.���, &HE7CFBA  '��ȻҪ�ɰ�ɫ
        Exit Sub
    End If
    
    If Not Bill.Active Then
        '��ʾ���۵�ժҪ:ҽ������
        If mbytInFun = 0 And mbytInState = 0 Then
            If mobjBill.Pages(mintPage).NO <> "" And Bill.RowData(Bill.Row) <> 0 Then
                strTmp = Get����ժҪ(mobjBill.Pages(mintPage).NO, 1, Bill.RowData(Bill.Row))
                If strTmp <> "" Then sta.Panels(Pan.C2��ʾ��Ϣ) = "ժҪ:" & strTmp
            End If
        End If
        Exit Sub
    End If
    If zlCheckBill���ڷ�ɢװ��ҩ(mintPage) = True Then
        '��������д��ڷ�ɢװ��,��������
        Call SetBill�в�ҩEditEnabled
         Exit Sub
    End If
    
     '--------------------------------------------------------------------------
    '1.�иı��������ݴ��������
    If mobjBill.Pages(mintPage).Details.Count >= Bill.Row And mlngPreRow <> Row Then
        '�շ�ʱ,���Ϊ������,�����޸�
        If (mbytInFun = 0 Or mbytInFun = 1) And mbytInState = 0 Then
            If mobjBill.Pages(mintPage).Details(Row).������ Then
                bln������ = True
                For i = 0 To UBound(marrColData)
                    Bill.ColData(i) = IIf(marrColData(i) = BillColType.UnFocus, BillColType.UnFocus, BillColType.Text_UnModify)
                Next
            End If
        End If
        
        If Not bln������ Then
            '��ʾ���
            With mobjBill.Pages(mintPage).Details(Bill.Row)
                If InStr(",5,6,7,", .�շ����) > 0 And .�շ�ϸĿID <> 0 Then
                    If gbln����ҩ�� Or gbln����ҩ�� Then
                        strStock = GetStockInfo(.�շ�ϸĿID, gbln����ҩ��, gbln����ҩ��)
                        If strStock <> "" Then
                            If InStr(1, mstrPrivs, "��ʾ���") > 0 Then
                                sta.Panels(Pan.C2��ʾ��Ϣ) = "��" & Bill.Row & "�п��:" & strStock
                            Else
                                sta.Panels(Pan.C2��ʾ��Ϣ) = "��" & Bill.Row & "���п��."
                            End If
                           ' Call ShowStatusCargoSpace(.�շ�ϸĿID, .ִ�в���ID)     '��ʾ��λ
                        End If
                        
                    End If
                    If strStock = "" Then
                        '���¿����ʾ
                        If Not (mbytInState = 0 And mstrInNO <> "") Then
                            .Detail.��� = GetStock(.�շ�ϸĿID, .ִ�в���ID)
                            If gblnҩ����λ Then .Detail.��� = .Detail.��� / .Detail.ҩ����װ
                        End If
                        Call ShowStock(.ִ�в���ID, .Detail.����, .Detail.���)
                        Call ShowStatusCargoSpace(.�շ�ϸĿID, .ִ�в���ID)     '��ʾ��λ
                    End If
                ElseIf .�շ���� = "4" And .Detail.�������� And .�շ�ϸĿID <> 0 Then
                    If Not (mbytInState = 0 And mstrInNO <> "") Then
                        .Detail.��� = GetStock(.�շ�ϸĿID, .ִ�в���ID, .Detail.����)
                    End If
                    Call ShowStock(.ִ�в���ID, .Detail.����, .Detail.���)
                Else
                    sta.Panels(Pan.C2��ʾ��Ϣ) = ""
                End If
                   
                Bill.ColData(BillCol.���) = IIf(gbln�շ����, BillColType.ComboBox, BillColType.UnFocus)
                Bill.ColData(BillCol.��Ŀ) = BillColType.CommandButton
                 '����Ǵ�����Ŀ������Ŀ�����,���������������Ŀ
                If CheckMainItem(Row) Or mobjBill.Pages(mintPage).Details(Row).�������� > 0 Then
                    Bill.ColData(BillCol.���) = BillColType.Text_UnModify
                    Bill.ColData(BillCol.��Ŀ) = BillColType.Text_UnModify
                End If
            
                '����Ƿǵ���״̬
                If mbytInState <> 2 Then
                    If .�շ���� = "7" And gblnPay Then
                        Bill.ColData(BillCol.����) = BillColType.Text
                    Else
                        Bill.ColData(BillCol.����) = BillColType.UnFocus
                    End If
                    
                    '���������������
                    If .Detail.��� And InStr(",5,6,7,", .�շ����) = 0 _
                        And Not (.�շ���� = "4" And .Detail.��������) Then
                        Bill.ColData(BillCol.����) = IIf(gblnTime, BillColType.Text, BillColType.UnFocus)   '����
                        Bill.ColData(BillCol.����) = BillColType.Text  '���
                    Else
                        Bill.ColData(BillCol.����) = BillColType.Text
                        Bill.ColData(BillCol.����) = BillColType.UnFocus
                    End If
                    
                    If .Key = "1" Then    'ָ���˹̶�ҩ��ʱ,��������ѡ��ִ�п���
                        Bill.ColData(BillCol.ִ�п���) = BillColType.UnFocus
                    Else
                        Bill.ColData(BillCol.ִ�п���) = BillColType.ComboBox
                    End If
                    
                    If .�շ���� = "F" Then
                        Bill.ColData(BillCol.��־) = BillColType.CheckBox
                    Else
                        Bill.ColData(BillCol.��־) = BillColType.UnFocus
                    End If
                    
                    'ֻ����һ�����,������ѡ�����
                    If mblnOne Then Bill.ColData(BillCol.���) = BillColType.UnFocus
                End If
                
                '��ʾ�����ժҪ
                If .ժҪ <> "" Then
                    sta.Panels(Pan.C2��ʾ��Ϣ) = sta.Panels(Pan.C2��ʾ��Ϣ) & "  ժҪ:" & .ժҪ
                End If
            End With
        End If
    End If
    
    '������δ�������,��ָ��е�����
    If mobjBill.Pages(mintPage).Details.Count < Bill.Row Then
        Bill.ColData(BillCol.���) = IIf(gbln�շ����, BillColType.ComboBox, BillColType.UnFocus) '�����,��������ʱ�ᱻ�ı�
        Bill.ColData(BillCol.��Ŀ) = BillColType.CommandButton  '��Ŀ��,��������ʱ�ᱻ�ı�
    End If
    
    
    '-----------------------------------------------------------------
    '2.�иı��������ݴ������ʾ����
    If Bill.ColData(Bill.Col) = BillColType.ComboBox Then   '���ص�ǰ�е�����������
        Call FillBillComboBox(Bill.Row, Bill.Col, True)
    End If
    
    If gbln�շ���� And Bill.TextMatrix(Row, BillCol.���) = "" And mblnOne Then
        mrsClass.Filter = "����=" & gstr�շ����
        Bill.TextMatrix(Row, BillCol.���) = mrsClass!���
        Bill.RowData(Row) = Asc(mrsClass!����)
    End If
    
    Bill.TextLen = 0: Bill.TextMask = ""
    Select Case Bill.TextMatrix(0, Col)
        Case "���" '���������ʱ���ᶨλ�������
            SetWidth Bill.cboHwnd, 70
            '������Ϊ��,���Զ�Ĭ��Ϊ��һ�շ�ϸĿ�����
            If Bill.TextMatrix(Row, Col) = "" Then
                If mblnOne Then
                    mrsClass.Filter = "����=" & gstr�շ����
                    Bill.TextMatrix(Row, Col) = mrsClass!���
                    Bill.RowData(Row) = Asc(mrsClass!����)
                ElseIf Row > 1 Then
                    Bill.ListIndex = -1
                    For i = 0 To Bill.ListCount - 1
                        If InStr(Bill.List(i), Bill.TextMatrix(Row - 1, Col)) > 0 Then Bill.ListIndex = i: Exit For
                    Next
                End If
            ElseIf Row >= 1 And Bill.TextMatrix(Row, Col) <> "" Then
                For i = 0 To Bill.ListCount - 1
                    If InStr(Bill.List(i), Bill.TextMatrix(Row, Col)) > 0 Then
                        Bill.ListIndex = i: Exit For
                    End If
                Next
                If Bill.ListIndex = -1 Then
                    Bill.ListIndex = SendMessage(Bill.cboHwnd, CB_FINDSTRING, -1, ByVal Bill.TextMatrix(Row - 1, Col))
                End If
            End If
        Case "ִ�п���", "��ҩҩ��"
            SetWidth Bill.cboHwnd, 130
        Case "����"
            Bill.TextLen = 3
            Bill.TextMask = "0123456789" & Chr(8)
        Case "����"
            Bill.TextLen = 8
            If (mbytInFun < 2 And InStr(mstrPrivs, "��������") = 0 Or mbytInFun = 2 And InStr(mstrPrivs, "��������") = 0) Then
                Bill.TextMask = "0123456789." & Chr(8)
            Else
                Bill.TextMask = "-0123456789." & Chr(8)
            End If
            
            If mobjBill.Pages(mintPage).Details.Count >= Bill.Row And InStr(Bill.TextMask, "-") > 0 Then
                If mobjBill.Pages(mintPage).Details(Bill.Row).Detail.���� Then
                    Bill.TextMask = Replace(Bill.TextMask, "-", "")
                End If
            End If
            
            If mobjBill.Pages(mintPage).Details.Count >= Bill.Row Then
                If InStr(",5,6,7,", mobjBill.Pages(mintPage).Details(Bill.Row).�շ����) > 0 Then
                    If InStr(mstrPrivs, "ҩƷ����С��") = 0 Then
                        Bill.TextMask = Replace(Bill.TextMask, ".", "")
                    End If
                    '��ҩ�������
                    If mobjBill.Pages(mintPage).Details(Bill.Row).�շ���� = "7" Then
                        Bill.TextMask = Bill.TextMask & gstrABC & LCase(gstrABC)
                    End If
                End If
            End If
        Case "����"
            Bill.TextLen = 10
            Bill.TextMask = "0123456789." & Chr(8)
    End Select
            
    '����,����������е����ʱ,�������л�û�п�ʼ
    If Bill.TextMatrix(Row, BillCol.��Ŀ) = "" Then
        mlngPreRow = 0
    ElseIf mobjBill.Pages(mintPage).Details.Count >= Row Then
        mlngPreRow = Row
    End If
End Sub

Private Sub Bill_LostFocus()
    Bill.TxtVisible = False
    Bill.CmdVisible = False
    Bill.CboVisible = False
End Sub

Private Sub Bill_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Bill.ToolTipText = Bill.TextMatrix(Bill.MouseRow, Bill.MouseCol)
End Sub

Private Sub cboBaby_Click()
    mobjBill.Ӥ���� = cboBaby.ListIndex
End Sub

Private Sub cboBaby_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cboSex_Click()
    If (mbytInFun = 0 Or mbytInFun = 1) And mbytInState = 0 Then
        mobjBill.�Ա� = zlStr.NeedName(cboSex.Text)
    End If
End Sub

Private Sub cbo�ѱ�_Click()
    If mbytInState = 0 Then
        If cbo�ѱ�.ListIndex = -1 Then
            mobjBill.�ѱ� = ""
        Else
            '��ʹ������ͬҲҪ����,��Ϊҽ���鿨���������,Ԥ�������ȷ
            If (mstrYBPati <> "" Or mobjBill.�ѱ� <> zlStr.NeedName(cbo�ѱ�.Text)) And Not mbln������۸� Then
                mobjBill.�ѱ� = zlStr.NeedName(cbo�ѱ�.Text)
                If mbytInState = 0 And Not CheckBillsEmpty Then
                    '��Ҫ����Ԥ����
                    If cmdԤ����.Visible Then
                        Call InitBalanceGrid
                        cmdԤ����.TabStop = True
                        cmdOK.Enabled = False
                    End If
'''                    Call zlClear���㿨
                    
                    'ȫ�����¼���۸�
                    Call CalcMoneys
                    Call ShowDetails
                    Call ShowMoney
                End If
            End If
        End If
    End If
End Sub

Private Sub cbo���㷽ʽ_Click()
'���ܣ����ֽ�����ֽ�֮���л�ʱ����Ҫ������������Ƿ���ֱ�
    If cbo���㷽ʽ.ListIndex = -1 Then Exit Sub
    If mblnNotClick Then Exit Sub
    If Not (Visible And mbytInFun = 0 And gBytMoney <> 0) Then Exit Sub
    
    If Bill.TextMatrix(0, Bill.COLS - 1) = "�˷�" Then
        Call ReCalce�˿�
    Else
        Call ShowMoney(-1) '��������δ��,ȫ���������¼���
    End If
    
'''    Call zlCheck֧Ʊ����
End Sub

Private Sub cbo���㷽ʽ_GotFocus()
    If mbytInState = 3 Or (chkCancel.Visible And chkCancel.Value = 1) Then
        Bill.Row = 1: Bill.Col = Bill.COLS - 1
    End If
End Sub

Private Sub cbo���㷽ʽ_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    
    If KeyAscii >= 32 Then
        If cbo���㷽ʽ.Locked Then Exit Sub
        
        lngIdx = zlControl.CboMatchIndex(cbo���㷽ʽ.hWnd, KeyAscii)
        If lngIdx = -1 And cbo���㷽ʽ.ListCount > 0 Then lngIdx = 0
        cbo���㷽ʽ.ListIndex = lngIdx
        
    ElseIf KeyAscii = 13 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cbo��������_Click()
    Dim i As Long, lng��������ID As Long
    
    mblnCboClick = True
    If Not (mbytInState = 0 And chkCancel.Value = 0) Then Exit Sub
        
    If cbo��������.ListIndex <> -1 Then lng��������ID = cbo��������.ItemData(cbo��������.ListIndex)
    If mobjBill.Pages(mintPage).��������ID = lng��������ID Then Exit Sub
    mobjBill.Pages(mintPage).��������ID = lng��������ID
        
    '��λҽ��
    If gbyt����ҽ�� = 1 Then
        If cbo��������.ListIndex <> -1 Then
            Call FillDoctor(lng��������ID)
            
            If cbo������.ListCount > 0 And (Not (gbln��ȱʡ������ And mbytInFun <> 2)) Then
                Call zlControl.CboSetIndex(cbo������.hWnd, 0)
            End If
        Else
            cbo������.Clear
        End If
        Call cbo������_Click
    End If
    
    
    '���ݿ����������������շ���Ŀ��ִ�п���
    If cbo��������.ListIndex <> -1 And Visible Then
        With mobjBill.Pages(mintPage)
        For i = 1 To .Details.Count
            If InStr(",4,5,6,7,", .Details(i).Detail.���) = 0 And _
             (.Details(i).Detail.ִ�п��� = 6 And gbyt����ҽ�� <> 2 Or InStr(",1,2,", "," & .Details(i).Detail.ִ�п��� & ",") > 0 And gint������Դ = 1) Then '6-�����˿���
                
                .Details(i).ִ�в���ID = lng��������ID
                
                If i <= Bill.Rows - 1 And .Details(i).ִ�в���ID <> 0 Then
                    If mbytInState = 0 Then
                        mrsUnit.Filter = "ID=" & .Details(i).ִ�в���ID
                        If mrsUnit.RecordCount <> 0 Then
                            Bill.TextMatrix(i, BillCol.ִ�п���) = IIf(zlIsShowDeptCode, mrsUnit!���� & "-", "") & mrsUnit!����
                        Else
                            Bill.TextMatrix(i, BillCol.ִ�п���) = GET��������(.Details(i).ִ�в���ID, mrsUnit)
                        End If
                    Else
                        '�������ֻ(��)��ʾ����
                        Bill.TextMatrix(i, BillCol.ִ�п���) = GET��������(.Details(i).ִ�в���ID, mrsUnit)
                    End If
                Else
                    Bill.TextMatrix(i, BillCol.ִ�п���) = ""
                End If
            ElseIf .Details(i).Detail.��� = "4" Then
                Call ReSet����ִ�п���(i) '113644
            End If
        Next
        End With
    End If
    
    
    'Ӥ���ѵĴ���,�������,�������޸�ʱ
    If mbytInFun = 2 And mbytBilling <> 2 Then
        If cboBaby.ListCount > 0 Then cboBaby.ListIndex = 0 '����click�¼�
        cboBaby.Enabled = False
        If cbo��������.ListIndex <> -1 Then cboBaby.Enabled = is����(cbo��������.ItemData(cbo��������.ListIndex), mrs��������)
    End If
    
    '�ѱ���
    Call LoadAndSeek�ѱ�
    
End Sub

Private Sub ReSet����ִ�п���(ByVal lngRow As Long)
    '�������Ҹı�����������������ϵ�ִ�п���
    '����ţ�113644
    '˵��������ִ�п���ȱʡ˳��:
    '    һ�����ﲡ��:
    '    1:ָ�����ϲ���(������ȱʡ���ϲ��š�)
    '    2: ��������
    '    3: ��һ�����õ�ִ�п���
    '    ����סԺ����:
    '    1:ָ�����ϲ���(������ȱʡ���ϲ��š�),�����Ƿ�����ڲ��˿���
    '    2: �����ɷ����ڲ��˿��ҵ�ִ�п���
    '     2.1:���˿���
    '     2.3:���˲���
    '     2.4:�ɷ����ڲ��˿��ҵĵ�һ��ִ�п���
    '    3: ��һ�����õ�ִ�п���
    Dim lngDoUnit As Long, lng���˿���ID As Long
    
    On Error GoTo errHandler
    If Not (mbytInFun = 1 And mbytInState = 0) Then Exit Sub
    With mobjBill.Pages(mintPage)
        If .Details(lngRow).Detail.��� <> "4" Then Exit Sub
        
        '����ִ�п���ȱʡΪ���˿���,�������ָ����,��Ϊָ������
        lngDoUnit = IIf(glng���ϲ��� > 0, glng���ϲ���, mobjBill.����ID)
        If lngDoUnit = 0 Then lngDoUnit = Get��������ID
                             
        '���˿���ID
        lng���˿���ID = mobjBill.����ID
        If lng���˿���ID = 0 And cbo��������.ListIndex <> -1 Then lng���˿���ID = cbo��������.ItemData(cbo��������.ListIndex)
        
        lngDoUnit = Get�շ�ִ�п���ID(.Details(lngRow).Detail.���, .Details(lngRow).Detail.ID, _
            .Details(lngRow).Detail.ִ�п���, lng���˿���ID, Get��������ID, gint������Դ, _
            IIf(mlng��ҩ�� = 0, glng��ҩ��, mlng��ҩ��), _
            IIf(mlng��ҩ�� = 0, glng��ҩ��, mlng��ҩ��), _
            IIf(mlng��ҩ�� = 0, glng��ҩ��, mlng��ҩ��), _
            lngDoUnit, mobjBill.����ID)
        
        .Details(lngRow).ִ�в���ID = lngDoUnit
        
        If lngRow <= Bill.Rows - 1 And .Details(lngRow).ִ�в���ID <> 0 Then
            mrsUnit.Filter = "ID=" & .Details(lngRow).ִ�в���ID
            If mrsUnit.RecordCount <> 0 Then
                Bill.TextMatrix(lngRow, BillCol.ִ�п���) = IIf(zlIsShowDeptCode, mrsUnit!���� & "-", "") & mrsUnit!����
            Else
                Bill.TextMatrix(lngRow, BillCol.ִ�п���) = GET��������(.Details(lngRow).ִ�в���ID, mrsUnit)
            End If
        Else
            Bill.TextMatrix(lngRow, BillCol.ִ�п���) = ""
        End If
    End With
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadAndSeek�ѱ�(Optional blnNew As Boolean)
'����:������ͨ�ѱ��붯̬�ѱ�,��λȱʡ�ѱ���˷ѱ�
'����:blnNew �Ƿ��µ��ݳ�ʼ
'˵��:������ʲ�ʹ�ö�̬�ѱ�
    Dim lngDeptID As Long, blnDo As Boolean, strInfo As String
    
    If glngSys Like "8??" Or mbytInFun = 2 Then Exit Sub
    
    If cbo��������.ListIndex <> -1 Then lngDeptID = cbo��������.ItemData(cbo��������.ListIndex)
    Call Load�ѱ�(cbo�ѱ�, lngDeptID, True, mrs�ѱ�)
                
    '��ʾ���ö�̬�ѱ𣺵�ǰ���ǻ��۵�ʱ,����Ĭ��Ϊ�ɼ�
    If Bill.Active Or blnNew Then
        lbl��̬�ѱ�.Caption = Load��̬�ѱ�(lngDeptID)
        lbl��̬�ѱ�.Tag = lbl��̬�ѱ�.Caption
        lbl��̬�ѱ�.Visible = lbl��̬�ѱ�.Caption <> ""
        If lbl��̬�ѱ�.Caption <> "" Then lbl��̬�ѱ�.Caption = "(" & lbl��̬�ѱ�.Caption & ")"
    End If
    
    cbo�ѱ�.Locked = (Not Bill.Active) Or (mbytInFun = 0 And mrsInfo.State = 1 And InStr(1, mstrPrivs, "�������˷ѱ�") = 0) Or mbytInFun = 2: cbo�ѱ�.TabStop = Not cbo�ѱ�.Locked And gbln�ѱ�
    If mrsInfo.State = 0 Then
         'δ�������Ĳ��˿�������ѡ��
         cbo�ѱ�.Locked = False: cbo�ѱ�.TabStop = Not cbo�ѱ�.Locked And gbln�ѱ�
        If cbo�ѱ�.ListIndex = -1 And cbo�ѱ�.ListCount > 0 Then cbo�ѱ�.ListIndex = 0
    Else
        '��λ�е������˵ķѱ�
        cbo�ѱ�.ListIndex = cbo.FindIndex(cbo�ѱ�, Nvl(mrsInfo!�ѱ�), True)
        If cbo�ѱ�.ListIndex <> -1 Then
            '�ѱ�Ϊ���ﵫ���˲��ǳ���
            If cbo�ѱ�.ItemData(cbo�ѱ�.ListIndex) = 2 And mrsInfo!���� = 0 Then
                blnDo = True
                strInfo = "���˷ѱ�""" & mrsInfo!�ѱ� & """���޳���ʱʹ��,���ò��˲��ǵ�һ�ξ���"
            End If
        Else
            blnDo = True
            strInfo = "���˷ѱ�" & mrsInfo!�ѱ� & "�����ã�������ʧЧ"
        End If
        
        If blnDo Then
            Call Load�ѱ�(cbo�ѱ�, lngDeptID, False, mrs�ѱ�)
            If cbo�ѱ�.ListIndex <> -1 Then
                If Visible And Not mblnDoing Then MsgBox strInfo & ",��ʹ��ȱʡ�ѱ�", vbInformation, gstrSysName
            Else
                cbo�ѱ�.Locked = False: cbo�ѱ�.TabStop = Not cbo�ѱ�.Locked And gbln�ѱ� '�޷�ȷ��,��������ѡ��
                If cbo�ѱ�.Visible And Not mblnDoing Then
                    MsgBox strInfo & ",��ѡ��һ�ַѱ�", vbInformation, gstrSysName
                    If cbo�ѱ�.Enabled Then cbo�ѱ�.SetFocus
                End If
            End If
        End If
    End If
End Sub

Private Sub cbo��������_Validate(Cancel As Boolean)
 '�����cbo��keypress�¼������˵����б��API����:sendmessage,�����ͣ��cbo��,����һ���ַ�,�ƿ�����򰴻س���,
'                                    cbo��ֵ�ᱣ������,�����ᴥ��click�¼�,������Ҫ��validate�¼��е���click�¼�

    If Not mblnCboClick Then cbo��������_Click
    If cbo��������.Text <> "" And cbo��������.ListIndex < 0 Then cbo��������.Text = ""
    mblnCboClick = False
End Sub

Private Function SetDefaultDept(lng������ID As Long) As Boolean
'����:����ȱʡ�Ŀ�������,��������Click�¼�
'˵��:ȱʡ����Ϊ"ֻ����������,������ҽ������"ʱ�����Զ�λȱʡ
'     ���߿����˵����п��Ҷ�Ϊͬһ�������򼶱�ʱ(�綼�Ǽ������������סԺ��)�����Զ�λȱʡ
'     ����,������Ա��ȱʡ���ң���GetDoctorDept�е�ҽ��˳��Ϊ׼,��һ��Ϊȱʡ
'     ��˳��Ϊ: 1.ֻ����������,������ҽ������(���,����,����,����,Ӫ��)
'               2.ֻ����������,����ҽ������(���,����,����,����,Ӫ��)
'               3.��ֻ�����������
    Dim i As Long, lng��������ID As Long, lng���ȼ� As Long, blnDo As Boolean
    
    mrs������.Filter = "ȱʡ=1 And ID=" & lng������ID
    If mrs������.RecordCount > 0 Then lng��������ID = mrs������!����ID
        
    If mrs��������.RecordCount > 1 And lng��������ID > 0 Then
        If gblnȱʡ�������� Then
            blnDo = True
        Else
            mrs��������.MoveFirst
            For i = 1 To mrs��������.RecordCount
                If lng��������ID = mrs��������!ID And mrs��������!���ȼ� = 1 Then blnDo = True: Exit For
                mrs��������.MoveNext
            Next
            
            If Not blnDo Then
                blnDo = True
                mrs��������.MoveFirst
                For i = 1 To mrs��������.RecordCount
                    If lng���ȼ� <> mrs��������!���ȼ� And lng���ȼ� <> 0 Then blnDo = False: Exit For
                    lng���ȼ� = mrs��������!���ȼ�
                    mrs��������.MoveNext
                Next
            End If
        End If
        
        If blnDo Then Call zlControl.CboSetIndex(cbo��������.hWnd, cbo.FindIndex(cbo��������, lng��������ID))
    End If
    
    If cbo��������.ListIndex = -1 Then Call zlControl.CboSetIndex(cbo��������.hWnd, 0)
End Function


Private Sub cbo������_Click()
    Dim i As Long, lng������ID As Long
    
    mblnCboClick = True
    If Not (mbytInState = 0 And chkCancel.Value = 0) Then Exit Sub
    If mobjBill.Pages(mintPage).������ = IIf(cbo������.ListIndex = -1, "", zlStr.NeedName(cbo������.Text)) Then Exit Sub
    
    mobjBill.Pages(mintPage).������ = IIf(cbo������.ListIndex = -1, "", zlStr.NeedName(cbo������.Text))
    If cbo������.ListIndex <> -1 Then
        lng������ID = cbo������.ItemData(cbo������.ListIndex)
        mrs������.Filter = "ID=" & lng������ID
        If mrs������.RecordCount > 0 Then
            lblDuty.Caption = IIf(IsNull(mrs������!רҵ����ְ��), "", mobjBill.Pages(mintPage).������ & "רҵְ��:" & mrs������!רҵ����ְ��)
        Else
            lblDuty.Caption = ""
        End If
    Else
        lblDuty.Caption = ""
    End If
    
    
    '��ҽ��ȷ������
    If gbyt����ҽ�� = 0 Then
        If cbo������.ListIndex <> -1 Then
            Call FillDept(mlngDeptID, lng������ID)
            Call SetDefaultDept(lng������ID)
        Else
            cbo��������.Clear
        End If
        Call cbo��������_Click
    End If
    
    '����ҽ������,��Ϊ�����˱��ˣ�����,ִ�п������ɿ����˿��Ҿ���ʱ����Ҫ����ִ�п���
     '������ʱ��Cbo��������_click�д���
    If cbo������.ListIndex <> -1 And Visible And gbyt����ҽ�� = 2 Then
        With mobjBill.Pages(mintPage)
        For i = 1 To .Details.Count
            If InStr(",4,5,6,7,", .Details(i).Detail.���) = 0 And .Details(i).Detail.ִ�п��� = 6 Then    '6-�����˿���
                
                mrs������.Filter = "ȱʡ=1 And ID=" & lng������ID
                If mrs������.RecordCount = 0 Then mrs������.Filter = "ID=" & lng������ID
                If mrs������.RecordCount > 0 Then
                    .Details(i).ִ�в���ID = mrs������!����ID
                Else
                    .Details(i).ִ�в���ID = 0
                End If
                
                If i <= Bill.Rows - 1 And .Details(i).ִ�в���ID > 0 Then
                    If mbytInState = 0 Then
                        mrsUnit.Filter = "ID=" & .Details(i).ִ�в���ID
                        If mrsUnit.RecordCount <> 0 Then
                            Bill.TextMatrix(i, BillCol.ִ�п���) = IIf(zlIsShowDeptCode, mrsUnit!���� & "-", "") & mrsUnit!����
                        Else
                            Bill.TextMatrix(i, BillCol.ִ�п���) = GET��������(.Details(i).ִ�в���ID, mrsUnit)
                        End If
                    Else
                        '�������ֻ(��)��ʾ����
                        Bill.TextMatrix(i, BillCol.ִ�п���) = GET��������(.Details(i).ִ�в���ID, mrsUnit)
                    End If
                Else
                    Bill.TextMatrix(i, BillCol.ִ�п���) = ""
                End If
            ElseIf .Details(i).Detail.��� = "4" Then
                Call ReSet����ִ�п���(i) '113644
            End If
        Next
        End With
    End If
    
    '��ʿ���
    If Bill.Active And Visible Then
        If mobjBill.Pages(mintPage).Details.Count < Bill.Rows - 1 _
            And Bill.Row = Bill.Rows - 1 And Bill.RowData(Bill.Rows - 1) <> 0 Then
            '�����Ч����
            Bill.TextMatrix(Bill.Rows - 1, BillCol.���) = ""
            Bill.RowData(Bill.Rows - 1) = 0
        ElseIf Bill.Col = BillCol.��� Then
            Call Bill_EnterCell(Bill.Row, Bill.Col) 'ˢ��
        End If
    End If
    
    '��ʿ���:�жϷǷ�����
    If Not mblnDoing Then
        If CheckInhibitiveByNurse(mintPage) Then
            MsgBox "��ʿֻ���������Ƽ�������Ŀ,�������д����������͵���Ŀ��", vbInformation, gstrSysName
        End If
    End If
End Sub


Private Sub cbo������_KeyDown(KeyCode As Integer, Shift As Integer)
    If cbo������.Locked Then Exit Sub
    mblnDrop = False
    If KeyCode = 13 Then mblnDrop = SendMessage(cbo������.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 1
End Sub
Private Sub cbo������_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub cbo������_Validate(Cancel As Boolean)
    If cbo������.Text <> "" Then
        If cbo.FindIndex(cbo������, zlStr.NeedName(cbo������.Text), True) = -1 Then cbo������.ListIndex = -1: cbo������.Text = ""
    End If
    If cbo������.Text = "" Then Call cbo������_KeyPress(vbKeyReturn)
    If gbyt����ҽ�� = 0 And gbln�����俪���� And cbo������.ListIndex = -1 Then Cancel = True
End Sub

Private Sub cbo���䵥λ_Validate(Cancel As Boolean)
    If (mbytInFun = 0 Or mbytInFun = 1) And mbytInState = 0 Then mobjBill.���� = Trim(txt����.Text) & IIf(IsNumeric(txt����.Text), cbo���䵥λ.Text, "")
End Sub

Private Sub cboҽ�Ƹ���_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    
    If KeyAscii >= 32 Then
        If cboҽ�Ƹ���.Locked Then Exit Sub
    
        lngIdx = zlControl.CboMatchIndex(cboҽ�Ƹ���.hWnd, KeyAscii)
        If lngIdx = -1 And cboҽ�Ƹ���.ListCount > 0 Then lngIdx = 0
        cboҽ�Ƹ���.ListIndex = lngIdx
        
    ElseIf KeyAscii = 13 And cboҽ�Ƹ���.ListIndex <> -1 Then
        If Bill.Active Then
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf Not Bill.Active Then
            If gbyt����ҽ�� = 0 Then
                '����:42886
                If txtDate.Enabled And txtDate.Visible Then
                    txtDate.SetFocus
                ElseIf cmdOK.Enabled And cmdOK.Visible Then
                    cmdOK.SetFocus
                End If
            Else
                If cbo��������.Enabled Then cbo��������.SetFocus
            End If
        End If
    End If
End Sub


Private Sub chkCancel_Click()
    Dim i As Integer
    
    mstrInNO = "": txtModi.Text = ""
    mlngFirstID = 0: mstrFirstWin = ""
    Call ClearPayInfo
        
    Call ClearPatientInfo(True)
    Call ClearTotalInfo
        
    Call InitCommVariable
    
    Call ClearBillRows: Call ClearMoney
    
    Bill.AllowAddRow = (chkCancel.Value = 0)
    IDKind.Enabled = (chkCancel.Value = 0)
    
    If chkCancel.Value = 1 Then
        chkCancel.ForeColor = &HFF&
        If cboBaby.Visible Then cboBaby.Enabled = False
        
        Call NewBill(False)
        Set mobjBill = New ExpenseBill
        If fraBill.Visible Then cmdAddBill.Enabled = InStr(1, mstrPrivs, "��ͨ���˶൥���շ�") > 0
        
        cboNO.Text = ""

        Call SetDisible
        If InStr(mstrPrivs, "��ʾ������") = 0 Then
            cbo������.Visible = False
            If gbyt����ҽ�� = 0 Then
                lbl����.Visible = False
            Else
                lbl������.Visible = False
            End If
        End If
        
        fraAppend.Enabled = False
        cbo���㷽ʽ.Enabled = False
        cboNO.Locked = False
        cmd�䷽.Enabled = False
        cmdYB.Enabled = False
        
        txtModi.Enabled = False
        txtIn.Text = ""
        txtIn.Enabled = False
        txtRePrint.Enabled = False
        
        txtInvoice.Text = ""
        txtInvoice.Locked = True
                
        For i = 0 To UBound(marrColData)
            Bill.ColData(i) = BillColType.Text_UnModify
        Next
        Call ShowDeleteCol(True)
        Bill.SetColColor BillCol.���, &HE7CFBA  '��ȻҪ�ɰ�ɫ
        
        If mbytInFun = 0 Then
            lblӦ��.Caption = "�˿�"
            lblӦ��.ForeColor = vbRed
            txtӦ��.ForeColor = vbRed
            txtӦ��.Text = "0.00"
        End If
        
        cboNO.SetFocus
    Else
        
        If InStr(mstrPrivs, "��ʾ������") = 0 Then
            cbo������.Visible = True
            If gbyt����ҽ�� = 0 Then
                lbl����.Visible = True
            Else
                lbl������.Visible = True
            End If
        End If
        
        txtRePrint.Enabled = True
        txtModi.Enabled = True
        txtIn.Text = ""
        txtIn.Enabled = True
        
        chkCancel.ForeColor = 0
        If fraBill.Visible Then cmdAddBill.Enabled = InStr(1, mstrPrivs, "��ͨ���˶൥���շ�") > 0
        txtInvoice.Locked = Not (InStr(1, mstrPrivs, "�޸�Ʊ�ݺ�") > 0) And gblnStrictCtrl
        If mbytBilling <> 2 Then Call SetDisible(True)
        cmd�䷽.Enabled = True
        cmdYB.Enabled = True
        
        Call NewBill(IIf(Not mblnStartFactUseType, False, True), False)
        Call Set�����˿�������(mobjBill.Pages(mintPage).������, mobjBill.Pages(mintPage).��������ID)
        Call LoadAndSeek�ѱ�
        
        For i = 0 To UBound(marrColData)
            Bill.ColData(i) = marrColData(i)
        Next
        Call ShowDeleteCol(False)
        Bill.SetColColor BillCol.���, &HE7CFBA  '��ȻҪ�ɰ�ɫ
        
        If mbytInFun = 0 Then
            lblӦ��.Caption = "Ӧ��"
            lblӦ��.ForeColor = 0
            txtӦ��.ForeColor = &HFF0000
            txtӦ��.Text = "0.00"
        End If
        
        If mbytBilling = 2 Then
            fraInfo.Enabled = False
            Bill.Active = False
            cboNO.Locked = False
            cboNO.SetFocus
        Else
            cbo��������.Enabled = True
            cbo������.Enabled = True
            
            fraAppend.Enabled = True
            If cbo���㷽ʽ.Visible Then cbo���㷽ʽ.Enabled = (mbytInState = 0)
            
            If mlng����ID > 0 Then
                txtPatient.Text = "-" & mlng����ID
                Call txtPatient_KeyPress(13)
                Bill.SetFocus
            Else
                txtPatient.SetFocus
            End If
        End If
    End If
End Sub

Private Sub chk����_Click()
    If chk����.Visible And Visible And mbytInFun = 0 Then
        '��Ҫ����Ԥ����
        If cmdԤ����.Visible Then
            Call InitBalanceGrid
            cmdԤ����.TabStop = True
            cmdOK.Enabled = False
        End If
'''        Call zlClear���㿨
    End If
End Sub

Private Sub chk����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk�Ӱ�_Click()
    Dim blnAdd As Boolean
    
    If Not mblnDo Then Exit Sub
    If mbytInState = 1 Or chkCancel.Value = 1 Or mbytBilling = 2 Then Exit Sub
    If mbytInState = 2 Then Exit Sub
    If Not chk�Ӱ�.Visible Or Not Visible Then Exit Sub
    
    blnAdd = OverTime(zlDatabase.Currentdate)
    If chk�Ӱ�.Value = 0 And blnAdd Then
        If MsgBox("��ǰ���ڼӰ�ʱ�䷶Χ��,Ҫȡ���Ӱ�Ӽ���", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            chk�Ӱ�.Value = 1
        End If
    End If
    If chk�Ӱ�.Value = 1 And Not blnAdd Then
        If MsgBox("��ǰ�����ڼӰ�ʱ�䷶Χ��,Ҫִ�мӰ�Ӽ���", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            chk�Ӱ�.Value = 0
        End If
    End If
    mobjBill.�Ӱ��־ = chk�Ӱ�.Value
    
    '��Ҫ����Ԥ����
    If cmdԤ����.Visible Then
        Call InitBalanceGrid
        cmdԤ����.TabStop = True
        cmdOK.Enabled = False
    End If
'''    Call zlClear���㿨
    
    'ȫ�����¼���۸�
    If Not CheckBillsEmpty Then
        Call CalcMoneys
        Call ShowDetails
        Call ShowMoney
    End If
End Sub

Private Sub chk�Ӱ�_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub AutoSplitBill()
    '����:�Զ������е��ݰ��շ������е��ݷ���
    '     �ݲ�����ҽ��,��ȡ������ģʽ��,����Ĺ����ѱ仯,�ݲ�����

    Dim i As Integer, j As Integer, strFKind As String, strFeeKind As String
    Dim intMinPage As Integer, intMaxPage As Integer, intPage As Integer, intRows As Integer
    Dim intOrder As Integer, intMainItem_New As Integer, intMainItem_Old As Integer, strMainKind As String, curError As Currency
    Dim blnMainItem As Boolean
    
    If cmdAddBill.Enabled = False Then Exit Sub
        
    If mobjBill.Pages.Count = 1 Then
        For i = 1 To mobjBill.Pages(1).Details.Count
            If i = 1 Then
                strFeeKind = IIf(gbytAutoSplitBill = 1, mobjBill.Pages(1).Details(i).�շ����, mobjBill.Pages(1).Details(i).ִ�в���ID)
            ElseIf strFeeKind <> IIf(gbytAutoSplitBill = 1, mobjBill.Pages(1).Details(i).�շ����, mobjBill.Pages(1).Details(i).ִ�в���ID) Then
                Exit For
            End If
        Next
        If i > mobjBill.Pages(1).Details.Count Then Exit Sub
    End If
        
    '�����С�ķǻ��۵���
    For i = 1 To mobjBill.Pages.Count
        If mobjBill.Pages(i).NO = "" Then Exit For
    Next
    If i > mobjBill.Pages.Count Then Exit Sub '����ȫΪ��,��ȫ�ǻ��۵�
    intMinPage = i
    intMaxPage = mobjBill.Pages.Count
    If mobjBill.�շѽ��� <> "" Then curError = mobjBill.Pages(intMaxPage).����� '���ֽ��㷽ʽ������Ǵ������һ�ŵ����ϵ�
    
    For i = intMinPage To intMaxPage
        intMainItem_Old = 0
        intMainItem_New = 0
        strMainKind = ""
        If i <> intMinPage Then
            '1.������ǰ�浥���������ͬ����
            j = 1
            intRows = mobjBill.Pages(i).Details.Count
            Do While j <= intRows
                If mobjBill.Pages(i).Details(j).�������� = 0 Then
                    blnMainItem = CheckMainItem(j, i)
                Else
                    blnMainItem = False
                End If
                If blnMainItem Then
                    intMainItem_Old = mobjBill.Pages(i).Details(j).���
                    intMainItem_New = intMainItem_Old
                    strMainKind = IIf(gbytAutoSplitBill = 1, mobjBill.Pages(i).Details(j).�շ����, mobjBill.Pages(i).Details(j).ִ�в���ID)
                End If
            
                '����ĸ��Ŵ���
                If mobjBill.Pages(i).Details(j).�������� = intMainItem_Old And intMainItem_Old <> 0 Then
                    If IIf(gbytAutoSplitBill = 1, mobjBill.Pages(i).Details(j).�շ����, mobjBill.Pages(i).Details(j).ִ�в���ID) = strMainKind Then
                        mobjBill.Pages(i).Details(j).�������� = intMainItem_New
                    Else
                        mobjBill.Pages(i).Details(j).�������� = 0
                    End If
                End If
                
                intPage = CheckKindInOtherPage(IIf(gbytAutoSplitBill = 1, mobjBill.Pages(i).Details(j).�շ����, mobjBill.Pages(i).Details(j).ִ�в���ID), i, 1) '��ǰ���
                If intPage > 0 Then
                    intOrder = AddRowByOtherPageRow(mobjBill.Pages(i).Details(j), intPage)
                    If blnMainItem Then intMainItem_New = intOrder
                                            
                    Call DeleteDetail(j, i) '��ǰ�������ѱ仯
                    j = j - 1
                    intRows = intRows - 1
                Else
                    If mobjBill.Pages(i).Details(j).�������� = intMainItem_Old And intMainItem_Old <> intMainItem_New Then  '����������,����û�ж�
                        mobjBill.Pages(i).Details(j).�������� = 0
                    End If
                End If
                j = j + 1
            Loop
        End If
        
        '2.�����뱾�����е�һ�����ͬ����.
        If mobjBill.Pages(i).Details.Count > 0 Then '������ǰ����ƶ�,ȫ����������,����Ϊ��
            strFKind = IIf(gbytAutoSplitBill = 1, mobjBill.Pages(i).Details(1).�շ����, mobjBill.Pages(i).Details(1).ִ�в���ID)
            If mobjBill.Pages(i).Details(1).�������� = 0 Then
                blnMainItem = CheckMainItem(1, i)
            Else
                blnMainItem = False
            End If
            If blnMainItem Then
                intMainItem_Old = 1
                intMainItem_New = 1
                strMainKind = strFKind
            End If
        End If
        j = 2
        intRows = mobjBill.Pages(i).Details.Count
        Do While j <= intRows
            If mobjBill.Pages(i).Details(j).�������� = 0 Then
                blnMainItem = CheckMainItem(j, i)
            Else
                blnMainItem = False
            End If
            If blnMainItem Then
                intMainItem_Old = mobjBill.Pages(i).Details(j).���
                intMainItem_New = intMainItem_Old
                strMainKind = IIf(gbytAutoSplitBill = 1, mobjBill.Pages(i).Details(j).�շ����, mobjBill.Pages(i).Details(j).ִ�в���ID)
            End If
            
            If IIf(gbytAutoSplitBill = 1, mobjBill.Pages(i).Details(j).�շ����, mobjBill.Pages(i).Details(j).ִ�в���ID) <> strFKind Then
                
                '����ĸ��Ŵ���
                If mobjBill.Pages(i).Details(j).�������� = intMainItem_Old And intMainItem_Old <> 0 Then
                    If IIf(gbytAutoSplitBill = 1, mobjBill.Pages(i).Details(j).�շ����, mobjBill.Pages(i).Details(j).ִ�в���ID) = strMainKind Then
                        mobjBill.Pages(i).Details(j).�������� = intMainItem_New
                    Else
                        mobjBill.Pages(i).Details(j).�������� = 0
                    End If
                End If
            
                intPage = CheckKindInOtherPage(IIf(gbytAutoSplitBill = 1, mobjBill.Pages(i).Details(j).�շ����, mobjBill.Pages(i).Details(j).ִ�в���ID), i, 0) '�����
                If intPage = 0 Then
                    Call AddNewBill
                    intPage = mobjBill.Pages.Count
                End If
                intOrder = AddRowByOtherPageRow(mobjBill.Pages(i).Details(j), intPage)
                If blnMainItem Then intMainItem_New = intOrder
                
                Call DeleteDetail(j, i)
                j = j - 1
                intRows = intRows - 1
            Else
                If mobjBill.Pages(i).Details(j).�������� = intMainItem_Old And intMainItem_Old <> intMainItem_New Then '����������,����û�ж�
                    mobjBill.Pages(i).Details(j).�������� = 0
                End If
            End If
            j = j + 1
        Loop
    Next
    
    '3.ɾ����Щ�����߶������Ŀյ���
    i = 1
    intMaxPage = mobjBill.Pages.Count
    Do While i <= intMaxPage
        If CheckBillsEmpty(i) Then
            Call DelOneBill(i)
            i = i - 1
            intMaxPage = intMaxPage - 1
        End If
        i = i + 1
    Loop
    
    'ˢ�½�����ʾ
    Call ShowDetails
    Call ShowMoney
    
    If mobjBill.�շѽ��� <> "" Then
        If mobjBill.Pages.Count = 1 Then
            mobjBill.Pages(1).�շѽ��� = mobjBill.�շѽ���
        Else
            Call ApportionMultiBalance(mobjBill.�շѽ���, curError)
        End If
    End If
End Sub

Private Function AddRowByOtherPageRow(tmpBillDetail As BillDetail, intPage As Integer) As Integer
'����:��ĳ�����ж������ӵ�ָ���ĵ���ҳ��
    Dim int��� As Integer
    
    With tmpBillDetail
        int��� = mobjBill.Pages(intPage).Details.Count + 1
        Call mobjBill.Pages(intPage).Details.Add(.�ѱ�, .Detail, .�շ�ϸĿID, int���, .��������, _
            .�շ����, .���㵥λ, .��ҩ����, .����, .����, .���ӱ�־, .ִ�в���ID, _
            .InComes, "", .������Ŀ��, .���մ���ID, .���ձ���, .ժҪ, .ԭʼ����, .ԭʼִ�в���ID)
    End With
    AddRowByOtherPageRow = int���
End Function


Private Function CheckKindInOtherPage(ByVal strKind As String, ByVal intCurrentPage As Integer, bytWay As Byte) As Integer
'����:���ǵ�ǰ����(���Ҳ��ǻ��۵�)���Ƿ����ָ�����շ�����ִ�в���
'����:bytWay-����������ݵķ���,0-�����,1-��ǰ���
'����:����������򷵻�0,�����򷵻ص�һ�����ڵĵ������
    Dim intBegin As Integer, intEnd As Integer, i As Integer, j As Integer

    If mobjBill.Pages.Count < 2 Then Exit Function
    If bytWay = 0 Then
        intBegin = intCurrentPage + 1
        intEnd = mobjBill.Pages.Count
    Else
        intBegin = 1
        intEnd = intCurrentPage - 1
    End If
    
    For i = intBegin To intEnd
        If mobjBill.Pages(i).NO = "" Then
            For j = 1 To mobjBill.Pages(i).Details.Count
                If IIf(gbytAutoSplitBill = 1, mobjBill.Pages(i).Details(j).�շ����, mobjBill.Pages(i).Details(j).ִ�в���ID) = strKind Then
                    CheckKindInOtherPage = i: Exit Function
                End If
            Next
        End If
    Next
End Function

Private Sub AddNewBill()
'���ܣ�����һ�ŵ���
    Dim objPage As New BillPage
    Dim i As Long

    '���뵥��ҳ��ǩ
    If tbsBill.Tabs.Count >= 10 Then
        Call tbsBill.Tabs.Add(, , "����" & tbsBill.Tabs.Count + 1)
    Else
        If tbsBill.Tabs.Count + 1 = 10 Then
            Call tbsBill.Tabs.Add(, , "����1&0")
        Else
            Call tbsBill.Tabs.Add(, , "����&" & tbsBill.Tabs.Count + 1)
        End If
    End If
    cmdDelBill.Enabled = True
    
    '���뵥��ҳ����:��ʹ�ǻ����շ�Ҳ����һ��
    mobjBill.Pages.Add objPage.Details
    
    '����ȱʡ�Ŀ�������,�������뵱ǰ��ͬ
    i = mobjBill.Pages.Count
    If cbo��������.ListIndex <> -1 Then
        mobjBill.Pages(i).��������ID = cbo��������.ItemData(cbo��������.ListIndex)
    Else
        mobjBill.Pages(i).��������ID = 0
    End If
    mobjBill.Pages(i).������ = IIf(cbo������.ListIndex = -1, "", zlStr.NeedName(cbo������.Text))
    
    '������㼯��:�����շ�ҲҪ����һ��
    mcolBalance.Add Array()
        
    '���ŵ���ʱ��ֹ�޸�,����,�˷ѹ���
    txtModi.Text = ""
    txtModi.Enabled = False
    chkCancel.Enabled = False
    cmdDelete.Enabled = False
End Sub

Private Sub cmdAddBill_Click()
    Dim i As Long
    Dim strFirst�ѱ� As String
    
    '��Ӧ�ж���Ŀյ���
    For i = 1 To mobjBill.Pages.Count
        If CheckBillsEmpty(i) Then
            MsgBox "�� " & i & " �ŵ�������Ϊ�գ������ڸõ��������롣", vbInformation, gstrSysName
            tbsBill.Tabs(i).Selected = True
            Bill.SetFocus: Exit Sub 'ȱʡΪֱ���������
        End If
    Next
    
    If tbsBill.Tabs.Count >= 200 Then
        MsgBox "��������̫�࣬��ֳɶ���շѡ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    strFirst�ѱ� = mobjBill.�ѱ�
            
    '���ӵ���
    Call AddNewBill
    
    '����Click,��ʾ�����ӵ��ݵ�����(�հ�)
    tbsBill.Tabs(tbsBill.Tabs.Count).Selected = True
    
    If mobjBill.Pages(1).NO <> "" Then cbo�ѱ�.ListIndex = cbo.FindIndex(cbo�ѱ�, strFirst�ѱ�, True)
    
    Bill.SetFocus 'ȱʡΪֱ���������
End Sub

Private Sub DelOneBill(ByVal intPage As Integer)
'���ܣ�ɾ��ָ���ĵ���
    Dim blnCurEmpty As Boolean, i As Integer
    
    blnCurEmpty = CheckBillsEmpty(intPage)
    
    'ɾ�����ݼ����е�����
    mobjBill.Pages.Remove intPage
    
    'ɾ�����㼯��
    mcolBalance.Remove intPage
    
    'ɾ��ҳ��֮���Զ����¶�λ,���Ҳ��ἤ��Click
    tbsBill.Tabs.Remove intPage
    For i = 1 To tbsBill.Tabs.Count
        If i = 10 Then
            tbsBill.Tabs(i).Caption = "����1&0"
        ElseIf i < 10 Then
            tbsBill.Tabs(i).Caption = "����&" & i
        Else
            tbsBill.Tabs(i).Caption = "����" & i
        End If
    Next
    If tbsBill.Tabs.Count = 1 Then cmdDelBill.Enabled = False
        
    '��Ҫ����Ԥ����
    If Not blnCurEmpty And cmdԤ����.Visible Then
        Call InitBalanceGrid
        cmdԤ����.TabStop = True
        cmdOK.Enabled = False
    End If
            
'''    '���˺�:??
'''    If Not blnCurEmpty Then
'''        Call zlClear���㿨
'''    End If
'''
    '���޸ļ��˷ѹ���
    If tbsBill.Tabs.Count = 1 Then
        txtModi.Text = ""
        txtModi.Enabled = True
        chkCancel.Enabled = True
        cmdDelete.Enabled = True
    End If
    
    '����Click,��ʾ�¶�λ���ݵ�����
    mintPage = 0 'ǿ�м���
    Call tbsBill_Click
End Sub

Private Sub cmdDelBill_Click()
'���ܣ�ɾ����ǰ����
    Dim i As Long
    
    If MsgBox("ȷʵҪɾ���� " & mintPage & " �ŵ�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        If mobjBill.Pages(mintPage).NO = "" Then
            Bill.SetFocus
        Else
            If txtPatient.Text = "" Then
                txtPatient.SetFocus
            Else
                'If txt�ɿ�.Enabled And txt�ɿ�.Visible Then
                '    txt�ɿ�.SetFocus
                If cmdԤ����.Enabled And cmdԤ����.Visible Then
                    cmdԤ����.SetFocus
                ElseIf cmdOK.Enabled And cmdOK.Visible Then
                    cmdOK.SetFocus
                End If
            End If
        End If
        Exit Sub
    End If
            
    'ɾ������
    Call DelOneBill(mintPage)
    
    '���¼���
    Call ShowMoney(-1)  '�������ݷ���δ��
    
    '�������ù�����(���������¼���)
    If gTy_Module_Para.bln������ Then
        If Not CheckBillsEmpty Then Call SetFactMoney
    End If
End Sub

Private Function LoadMultiBills(ByVal lng����ID As Long, ByVal bln������൥�� As Boolean, _
    ByVal lng�Һſ��� As Long, Optional blnCard As Boolean) As Boolean
'���ܣ�һ���Զ�ȡ���˵Ķ��Ż��۵�,�ù����ڲ��˶�ȡ�ɹ�֮�����
'������bln������൥�ݣ�ҽ�������շѻ�֧�ֶ൥���շ�ʱ���������ض��Ż��۵��շ�
'      lng�Һſ���,��ͨ���Һŵ�����ʱ,���벡�˵�ǰ�Һŵ��ĹҺſ���

    Dim objPage As New BillPage
    Dim arrBills As Variant, strBills As String
    Dim blnRead As Boolean, i As Long, k As Long
    
    If Not (gblnMulti And gblnSeekBill) Then Exit Function
    
    If lng����ID = 0 Then Exit Function
    i = SeekPatiBill(lng����ID)
    If i = 0 Then Exit Function
    If gblnUnPopPriceBill Then
        strBills = frmPatiPrice.GetPriceBillString(lng����ID, bln������൥��, lng�Һſ���, mTy_Para.blnסԺ���������շ�, blnCard)
    Else
        strBills = frmPatiPrice.FindBill(Me, mstrPrivs, lng����ID, bln������൥��, lng�Һſ���, mTy_Para.blnסԺ���������շ�, blnCard)
    End If
     
    If strBills = "" Then Exit Function
    
    
    LoadMultiBills = True
    '������е��ݵ�����
    '---------------------------------------------------------------------
    mstrInNO = "": txtModi.Text = ""
    Call ClearTotalInfo
    Call ClearPayInfo
    Call ClearBillRows
        
    'Ԥ����֧��ʱ�����,������Զ���
    If cmdԤ����.Visible Then
        Call InitBalanceGrid
    End If
    

    '��ȡ���۵����¼���ʱ,��Ҫ�ۼ���ʾ�ڱ����
    '���˺�,����:22343;ֻ��������ɿ����,�Ŵ����ۼƵ�����
    '  Not gbln�ɿ���� ȡ��
    '51670: �ֵ������ۼƺͶಡ���ۼ�
    If gTy_Module_Para.byt�ɿ���� <> 1 And gTy_Module_Para.byt�ɿ���� <> 3 Or mstrPrePati = "" Then
        Call ClearMoney
    End If
    
    Set mcolBalance = New Collection
    mcolBalance.Add Array()
        
    '�൥���շ�:ֻ����һҳ����
    For i = mobjBill.Pages.Count To 1 Step -1
        mobjBill.Pages.Remove i
    Next
    mobjBill.Pages.Add objPage.Details
    
    '�൥���շ�:�ָ�ȱʡ����ҳ��
    mintPage = 1
    For i = tbsBill.Tabs.Count To 1 Step -1
        tbsBill.Tabs(i).Tag = ""
        If i <> 1 Then tbsBill.Tabs.Remove i
    Next
        
    '��ȡ��ʾÿ�Ż��۵�
    '---------------------------------------------------------------------
    mblnNOMoved = False '���۵���ȡ���Ӻ󱸱��ж�
    k = 1
    mblnDoing = True '���������Զ���
    arrBills = Split(strBills, ",")
    For i = 0 To UBound(arrBills)
        Me.Refresh
        '���ӵ���ҳ��ǩ(ͬcmdAdd_Click����)
        '-----------------------------------------------------------------------
        If k > 1 And mobjBill.Pages(mobjBill.Pages.Count).NO <> "" Then
            If tbsBill.Tabs.Count >= 10 Then
                Call tbsBill.Tabs.Add(, , "����" & tbsBill.Tabs.Count + 1)
            Else
                If tbsBill.Tabs.Count + 1 = 10 Then
                    Call tbsBill.Tabs.Add(, , "����1&0")
                Else
                    Call tbsBill.Tabs.Add(, , "����&" & tbsBill.Tabs.Count + 1)
                End If
            End If
            
            '���뵥��ҳ����:��ʹ�ǻ����շ�Ҳ����һ��
            mobjBill.Pages.Add objPage.Details
            
            '������㼯��:�����շ�ҲҪ����һ��
            mcolBalance.Add Array()
                
            '���ŵ���ʱ��ֹ�޸ļ��˷ѹ���
            txtModi.Enabled = False
            chkCancel.Enabled = False
            cmdDelete.Enabled = False
                
            '����Click,��ʾ�����ӵ��ݵ�����(�հ�)
            tbsBill.Tabs(tbsBill.Tabs.Count).Selected = True
        End If
                
        '��ȡ���۵�������(ͬcboNO_KeyPress)
        '----------------------------------------------------------------------
        blnRead = ReadBill(arrBills(i), 1, False)
        If blnRead Then k = k + 1: cboNO.Text = arrBills(i)
    Next
    Bill.Active = False
    chk�Ӱ�.Enabled = False
    If mbln���� And mstr���ת��ʱ�� <> "" Then
        txtDate.Text = Format(CDate(mstr���ת��ʱ��) - 1 / 24 / 60, "yyyy-mm-dd HH:MM:SS")
    Else
        txtDate.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    End If
    
    cmdDelBill.Enabled = tbsBill.Tabs.Count > 1
    
    mblnDoing = False '�����Զ���ȡ���
    
    
    '��ʾժҪ
    Call Bill_EnterCell(1, BillCol.��Ŀ)
    '����Ʊ���Ƿ����
    If gTy_Module_Para.bytƱ�ݷ������ <> 0 Then
        Dim str��Ʊ�� As String, int���� As Integer
        If mintInvoicePrint <> 0 Then
            If zlExeCuteBillNoSplit(True, 1, mlng����ID, strBills, 0, txtInvoice.Text, Now, 1, str��Ʊ��, int����) Then
                Call zlCheckFactIsEnough(int����)
            End If
        End If
    End If
    
    If mstrYBPati = "" And gbln���������ɿ� Then
       If cmdOK.Enabled And cmdOK.Visible Then
            cmdOK.SetFocus
        End If
    End If
End Function



Private Sub ReInitPatiInvoice(Optional blnFact As Boolean = True, _
    Optional ByVal intInsure_IN As Integer = 0, Optional ByVal lng����ID_In As Long = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���³�ʼ�����˷�Ʊ��Ϣ
    '���:blnFact-�Ƿ�����ȡ��Ʊ��
    '����:���˺�
    '����:2011-04-29 14:17:33
    '����:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInvoiceFormat As String, lng����ID As Long
    Dim intInsure As Integer
    lng����ID = IIf(lng����ID_In <> 0, lng����ID_In, mobjBill.����ID)
    intInsure = IIf(intInsure_IN <> 0, intInsure_IN, mintInsure)
    
    If lng����ID = 0 Then
        '�ϴβ���ID
        If (mbytInFun = 0 Or mbytInFun = 1) And txtPatient.Text = mstrPrePati And mlngPrePati <> 0 Then
            lng����ID = mlngPrePati
        End If
    End If
    If lng����ID = 0 Then
        If Not mrsInfo Is Nothing Then
            If mrsInfo.State = 1 Then lng����ID = Val(Nvl(mrsInfo!����ID))
        End If
    End If
    
    mstrUseType = "": mlngShareUseID = 0: mintInvoiceFormat = 0
    mstrUseType = zl_GetInvoiceUserType(lng����ID, 0, intInsure)
    mlngShareUseID = zl_GetInvoiceShareID(mlngModul, mstrUseType)
    mintInvoiceFormat = zl_GetInvoicePrintFormat(mlngModul, mstrUseType, mintOldInvoiceFormat)
    mintInvoicePrint = zl_GetInvoicePrintMode(mlngModul, mstrUseType)
    
    Call ZlShowBillFormat(mlngModul, lblFormat, mintInvoiceFormat)
    If blnFact Then Call RefreshFact
End Sub

Private Function zlGetInvoiceGroupUseID(ByRef lng����ID As Long, _
    Optional intNum As Integer = 1, Optional strInvoiceNO As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡƱ�ݵ�����ID
    '���:lng����ID-����id
    '       intNum-ҳ��
    '       strInvoiceNO-����ķ�Ʊ��
    '����:lng����ID-����ID
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-04-29 15:36:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    lng����ID = GetInvoiceGroupID(1, intNum, lng����ID, mlngShareUseID, strInvoiceNO, mstrUseType)
    If lng����ID <= 0 Then
        Select Case lng����ID
            Case 0 '����ʧ��
            Case -1
                If Trim(mstrUseType) = "" Then
                    MsgBox "��û�����ú͹��õ��շ�Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                Else
                    MsgBox "��û�����ú͹��õġ�" & mstrUseType & "���շ�Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                End If
                Exit Function
            Case -2
                If Trim(mstrUseType) = "" Then
                    MsgBox "���صĹ���Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                Else
                    MsgBox "���صĹ���Ʊ�ݵġ�" & mstrUseType & "���շ�Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                End If
                Exit Function
            Case -3
                MsgBox "��ǰƱ�ݺ��벻�ڿ����������ε���ЧƱ�ݺŷ�Χ��,���������룡", vbInformation, gstrSysName
                If txtInvoice.Enabled And txtInvoice.Visible Then txtInvoice.SetFocus
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

Private Sub RefreshFact()
    '���ܣ�ˢ���շ�Ʊ�ݺ�
    If mintInvoicePrint = 0 Then Exit Sub
    If gblnStrictCtrl Then
        'lblFact.tag��Ҫ�Ǽ�鷢Ʊ���Ƿ��ֹ������.�ֹ������,��Ʊ��Ϊ��,�������Զ������ķ�Ʊ��
        If (lblFact.Tag <> "" And txtInvoice.Text <> "") Or Trim(txtInvoice.Text) = "" Then
            If zlGetInvoiceGroupUseID(mlng����ID) = False Then
                txtInvoice.Text = "": txtInvoice.Tag = "": Exit Sub
            End If
            '�ϸ�ȡ��һ������
            txtInvoice.Text = GetNextBill(mlng����ID)
            'Tag�����⣺24363:���˺飺��Ҫ�ǽ���Զ����ɵĺ��Ƿ��û����ģ���Ҫ�����
            '    1.���ĵ�Ʊ�ݺ���Ҫ����Ƿ��ظ����ظ���ֱ�ӷ��ز����ķ�Ʊ��
            '    2.���������������ĵ�����£�����Ƿ��ظ�������ظ����Զ�ȡ��һ�����룡
            txtInvoice.Tag = txtInvoice.Text
            lblFact.Tag = txtInvoice.Tag
            If mblnStartFactUseType Then Call zlCheckFactIsEnough
        End If
    Else
        If (lblFact.Tag <> "" And txtInvoice.Text <> "") Or Trim(txtInvoice.Text) = "" Then
            '��ɢ��ȡ��һ������
            txtInvoice.Text = zlStr.Increase(UCase(zlDatabase.GetPara("��ǰ�շ�Ʊ�ݺ�", glngSys, mlngModul)))
            'Tag�����⣺24363:���˺飺��Ҫ�ǽ���Զ����ɵĺ��Ƿ��û����ģ���Ҫ�����
            '    1.���ĵ�Ʊ�ݺ���Ҫ����Ƿ��ظ����ظ���ֱ�ӷ��ز����ķ�Ʊ��
            '    2.���������������ĵ�����£�����Ƿ��ظ�������ظ����Զ�ȡ��һ�����룡
        End If
        txtInvoice.Tag = txtInvoice.Text
        lblFact.Tag = txtInvoice.Tag
    End If
    txtInvoice.SelStart = Len(txtInvoice.Text)
End Sub

Private Sub cmdDelete_Click()

    If (mbytInFun = 0 And Not gblnMulti) Or mbytInFun = 2 Then
        cmd�䷽.Enabled = Not cmd�䷽.Enabled
        cmdYB.Enabled = Not cmdYB.Enabled
    End If
    If frmMultiBills.ShowMe(Me, 1, mstrPrivs, "", "", , mlng����ID, mblnOneCard) Then
        Call RefreshFact
        If gbln�ۼ� Then txt�ۼ�.Text = Format(GetChargeTotal, "0.00")
    End If
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
End Sub

Private Sub cmdIDCard_Click()
    Dim strCommon As String, intAtom As Integer
            
    On Error Resume Next
    If gobjPatient Is Nothing Then
        Set gobjPatient = CreateObject("zl9Patient.clsPatient")
        If gobjPatient Is Nothing Then Exit Sub
    End If
    
    Err.Clear: On Error GoTo 0
    
    '�������úϷ�������
    strCommon = Format(Now, "yyyyMMddHHmm")
    strCommon = TranPasswd(strCommon) & "||" & OS.ComputerName
    intAtom = GlobalAddAtom(strCommon)
    Call SaveSetting("ZLSOFT", "����ȫ��", "����", intAtom)
    Call gobjPatient.IDCard(Me, gcnOracle, glngSys, gstrDBUser)
    Call GlobalDeleteAtom(intAtom)
    
    If txtPatient.Enabled Then txtPatient.SetFocus
End Sub

Private Sub cmdRegist_Click()
    Dim strCommon As String, intAtom As Integer, blnOK As Boolean
            
    On Error Resume Next
    If gobjRegist Is Nothing Then
        Set gobjRegist = CreateObject("zl9RegEvent.clsRegEvent")
        If gobjRegist Is Nothing Then Exit Sub
    End If
    
    Err.Clear: On Error GoTo 0
    
    '�������úϷ�������
    strCommon = Format(Now, "yyyyMMddHHmm")
    strCommon = TranPasswd(strCommon) & "||" & OS.ComputerName
    intAtom = GlobalAddAtom(strCommon)
    Call SaveSetting("ZLSOFT", "����ȫ��", "����", intAtom)
    
    blnOK = gobjRegist.Register(Me, gcnOracle, glngSys, gstrDBUser, gblnSharedInvoice, IIf(gblnSharedInvoice, mlngShareUseID, 0))
    Call GlobalDeleteAtom(intAtom)
    '��ɹҺ�
    'ˢ��Ʊ�ݺ�
    If gblnSharedInvoice And blnOK Then
        If txtInvoice.Enabled Then Call RefreshFact
    End If
    If txtPatient.Enabled Then txtPatient.SetFocus
End Sub

Private Sub cmd�䷽_Click()
    Call ShowCHRecipe
End Sub

Private Sub zlChangePatiSource(ByVal int������Դ As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ı䲡����Դ״̬
    '����:���˺�
    '����:2010-01-13 11:23:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim Panel As MSComctlLib.Panel
    
    Set Panel = sta.Panels("PatiSource")
    Select Case int������Դ
    Case 1 '����
        Set Panel.Picture = imgPati.ListImages("OutPati").Picture
        Panel.ToolTipText = "���ﲡ��"
        gstrҩ����λ = "���ﵥλ": gstrҩ����װ = "�����װ"
    Case Else    'סԺ
        Set Panel.Picture = imgPati.ListImages("InPati").Picture
        Panel.ToolTipText = "סԺ����"
        gstrҩ����λ = "סԺ��λ": gstrҩ����װ = "סԺ��װ"
    End Select
    sta.Panels(Pan.C2��ʾ��Ϣ).Text = "�ѽ�������Դ����Ϊ" & IIf(int������Դ = 1, "���ﲡ��", "סԺ����")
    Set mrsUnit = GetDepartments("", gint������Դ & ",3")
    Set mrs�������� = Nothing
    Set mrs������ = Nothing
    Call FillDept(mlngDeptID)
    Call FillDoctor
    Call ClearFullBill(False)    '��Ҫ������mobjBill.�����־
End Sub
Private Sub picAppend_Resize()
    Dim sngLeft As Single
    Err = 0: On Error Resume Next
    sngLeft = vsBalance.Left + vsBalance.Width + 100
    If fra�ɿ�.Visible Then sngLeft = fra�ɿ�.Left + fra�ɿ�.Width + 100
    If mbytInFun = 0 Then
        cmdOK.Left = sngLeft + (ScaleWidth - sngLeft - cmdOK.Width) \ 2 '  ScaleWidth - cmdOK.Width - 100
        cmdCancel.Left = cmdOK.Left
        cmdPrint.Left = cmdOK.Left
        cmdԤ����.Left = cmdOK.Left
    End If
    If mbytInFun <> 0 Then Exit Sub
    If mbytInState <> 3 Then
        lblԤ�����.Left = vsBalance.Left
        txtԤ�����.Left = lblԤ�����.Left + lblԤ�����.Width + 10
        txtԤ�����.Top = picAppend.ScaleHeight - txtԤ�����.Height - 10
        lblԤ�����.Top = txtԤ�����.Top + (txtԤ�����.Height - lblԤ�����.Height) \ 2
        txtԤ�����.Width = vsBalance.Left + vsBalance.Width - txtԤ�����.Left
    End If
    If mbytInState = 0 Then
        vsBalance.Height = picAppend.ScaleHeight - vsBalance.Top - 20
    ElseIf mbytInState = 3 Then
        '�˿�
        vsBalance.Height = picAppend.ScaleHeight - vsBalance.Top - 20
        fra�ɿ�.Left = vsBalance.Left + vsBalance.Width + 20
    Else
        vsBalance.Height = IIf(lblԤ�����.Visible = False, picAppend.ScaleHeight, txtԤ�����.Top) - vsBalance.Top - 20
    End If
End Sub

Private Sub sta_PanelClick(ByVal Panel As MSComctlLib.Panel)
    Dim lngR As Long
    If Panel.Key = "Calc" Then
        lngR = FindWindow("SciCalc", "������")
        If lngR <> 0 Then
            BringWindowToTop lngR
        Else
            On Error Resume Next
            Shell "calc.exe", vbNormalFocus
        End If
    ElseIf Panel.Key = "Drugstore" Then
        With frmSetExpence
            .mlngModul = mlngModul
            .mstrPrivs = mstrPrivs
            .mbytInFun = mbytInFun
            .mblnSetDrugStore = True
            .Show 1, Me
        End With
    ElseIf Panel.Key = "PatiSource" Then
        If gbln������Դ��Ȩ�޿��� And InStr(1, mstrPrivs, ";��������;") = 0 Or mbln���� Then
            '��Ȩ�޿���,���ܸ���
            Exit Sub
        End If
        If Not CheckBillsEmpty Or txtPatient.Text <> "" Then
            If MsgBox("����л�������Դ,����յ�ǰ���ݺͲ�����Ϣ" & vbCrLf & "��ȷ��Ҫ������?", vbInformation + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                Exit Sub
            End If
        End If
        
        If gint������Դ = 1 Then    '����
           ' Set Panel.Picture = imgPati.ListImages("InPati").Picture
            'Panel.ToolTipText = "סԺ����"
            gint������Դ = 2
            'gstrҩ����λ = "סԺ��λ": gstrҩ����װ = "סԺ��װ"
        Else
            'Set Panel.Picture = imgPati.ListImages("OutPati").Picture
            'Panel.ToolTipText = "���ﲡ��"
            gint������Դ = 1
            'gstrҩ����λ = "���ﵥλ": gstrҩ����װ = "�����װ"
        End If
        
        zlDatabase.SetPara "������Դ", gint������Դ, glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0
        Call zlChangePatiSource(gint������Դ)
        mblnAutoChangePati = False
    ElseIf Panel.Bevel = sbrRaised And (Panel.Key = "PY" Or Panel.Key = "WB") Then
        If Not gbln�����л� Then Exit Sub     '35242
        '�л����������ƥ�䷽ʽ
        Panel.Bevel = IIf(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        If Panel.Key = "PY" Then
            sta.Panels("WB").Bevel = IIf(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        Else
            sta.Panels("PY").Bevel = IIf(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        End If
        zlDatabase.SetPara "���뷽ʽ", IIf(sta.Panels("PY").Bevel = sbrInset And sta.Panels("WB").Bevel = sbrInset, 2, IIf(sta.Panels("WB").Bevel = sbrInset, 1, 0))
        gbytCode = Val(zlDatabase.GetPara("���뷽ʽ", , , True))
    End If
End Sub

Private Sub ShowDeposit(ByVal lngPatientID As Long)
    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select Nvl(Sum(���), 0) Ԥ�����ܶ�, Nvl(Sum(��Ԥ��), 0) ��Ԥ���ܶ� From ����Ԥ����¼ Where ����id = [1] And ��¼���� In(1,11) and nvl(Ԥ�����,2)=1"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngPatientID)
    
    If rsTmp.RecordCount > 0 Then
        MsgBox "Ԥ�����ܶ�:" & Format(rsTmp!Ԥ�����ܶ�, "0.00") & vbCrLf & "��Ԥ���ܶ�:" & Format((rsTmp!��Ԥ���ܶ� - Original.��Ԥ����), "0.00") & vbCrLf & _
               "δ �� ����:" & Format(Val(cmdCancel.Tag), "0.00") & vbCrLf & _
               "����Ԥ����:" & Format((rsTmp!Ԥ�����ܶ� - (rsTmp!��Ԥ���ܶ� - Original.��Ԥ���� + Val(cmdCancel.Tag))), "0.00"), vbInformation, gstrSysName
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub sta_PanelDblClick(ByVal Panel As MSComctlLib.Panel)
    If Panel Is sta.Panels(Pan.C4Ԥ����Ϣ) And mrsInfo.State = 1 Then
        Call ShowDeposit(mrsInfo!����ID)
    End If
End Sub

Private Sub tbsBill_Click()
'���ܣ���ʾѡ��ҳ����ҳ��������
'˵����Ŀǰֻ���շ�ʱ�ſ��ܻ����
    Dim i As Integer, str�ѱ� As String, blnLock As Boolean
    
    '��ͬ���ʱ�˳�(ֻ��һ��ʱ�൱�ڲ�����)
    If tbsBill.SelectedItem.Index = mintPage Then Exit Sub
    mintPage = tbsBill.SelectedItem.Index
    
    '��������ʾ������
    Call ClearBillRows
    If mobjBill.Pages(mintPage).Details.Count > 0 Then
        Bill.Rows = mobjBill.Pages(mintPage).Details.Count + 1
    Else
        Bill.Rows = 2
    End If
    
    Call InitBillColumnColor
    
    '�����к�
    Call SetColNum
    
    If Not mblnDoing Then
        '������ʾ����ˢ��,��ȷ�����ݿɷ�༭
        mblnDoing = True
        cboNO.Text = mobjBill.Pages(tbsBill.SelectedItem.Index).NO
        If mobjBill.Pages(tbsBill.SelectedItem.Index).NO = "" Then
            Bill.Active = True
            mbln������۸� = True
            
            '��ʾ��������,������
            If mbytInFun = 0 Then '���շѴ���ҽ�����۵�����շ�ʱ��ס,���Խ���
                cbo��������.Locked = False
                cbo������.Locked = False
            End If
            Call Set�����˿�������(mobjBill.Pages(mintPage).������, mobjBill.Pages(mintPage).��������ID)
                        
            '��̬�ѱ����ʾ,Ҫ�ڿ�����ʾ֮��
            If cbo�ѱ�.Visible Then
                str�ѱ� = zlStr.NeedName(cbo�ѱ�.Text)
                blnLock = cbo�ѱ�.Locked
            End If
            
            cbo�ѱ�.Visible = True
            lbl��̬�ѱ�.Visible = True
            lbl��̬�ѱ�.BorderStyle = 0
            lbl��̬�ѱ�.Left = cbo�ѱ�.Left + cbo�ѱ�.Width + 60
            Call LoadAndSeek�ѱ�
            
            If str�ѱ� <> "" Then Call zlControl.CboLocate(cbo�ѱ�, str�ѱ�)
            If cbo�ѱ�.ListIndex <> -1 Then cbo�ѱ�.Locked = blnLock
            cbo�ѱ�.TabStop = Not cbo�ѱ�.Locked And gbln�ѱ�
            
            mbln������۸� = False
            Call ShowDetails
        Else
            Bill.Active = False
            Call ReadBill(mobjBill.Pages(mintPage).NO, 1, False, , True)
        End If
        mblnDoing = False
        
        'ȱʡ��λ��Ԫ
        If mobjBill.Pages(tbsBill.SelectedItem.Index).NO = "" Then
            If mobjBill.Pages(mintPage).Details.Count = 0 Then
                Bill.Col = Bill.MsfObj.FixedCols
            Else
                Bill.Col = Bill.PrimaryCol
                mlngPreRow = 0
            End If
            Bill.Row = 1
        ElseIf Visible Then
            sta.Panels(Pan.C2��ʾ��Ϣ).Text = ""
        End If
        If Visible Then Bill.SetFocus
    End If
End Sub

Private Function CheckBillsEmpty(Optional ByVal intPage As Integer) As Boolean
'���ܣ��ж��Ƿ�൥�ݵ����ݶ�Ϊ��
'������intPage=�Ƿ���ָ��ҳ,����������ҳ
    Dim i As Integer
    
    If intPage = 0 Then
        For i = 1 To mobjBill.Pages.Count
            If mobjBill.Pages(i).NO <> "" Then
                Exit Function
            ElseIf mobjBill.Pages(i).Details.Count > 0 Then
                Exit Function
            End If
        Next
    Else
        If mobjBill.Pages(intPage).NO <> "" Then
            Exit Function
        ElseIf mobjBill.Pages(intPage).Details.Count > 0 Then
            Exit Function
        End If
    End If
    CheckBillsEmpty = True
End Function

Private Function ClearFullBill(ByVal bln��ʾ As Boolean, _
    Optional blnClearPatiInfor As Boolean = True, _
    Optional blnNotClearYb As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������Ϣ
    '���:blnNotClearYb-�����ҽ������
    '����:����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-26 11:55:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strYBPati As String, intInsure As Integer
    Dim blnAdd As Boolean, strYBBill As String
    Dim cur������� As Currency, cur����͸֧ As Currency, blnYB�������� As Boolean
    
    strYBPati = mstrYBPati: intInsure = mintInsure: strYBBill = mstrYBBill
    cur������� = mcur�������: cur����͸֧ = mcur����͸֧: blnYB�������� = mblnYB��������
    blnAdd = cmdAddBill.Enabled
    '�����ҽ����Ϣ
    If bln��ʾ Then
        If MsgBox("ȷʵҪ�����ǰ�����е�������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Function
        End If
    End If
    
    If Not blnNotClearYb Then
        If YBIdentifyCancel = False Then 'ȡ��ҽ�����������֤
            Exit Function                '���ؼ�ʱ�������
        End If
    End If
    
    '˫����ʾ��������ڵ�ǰ������ʾ֮�������ʾ�����ƶ�����
    Call ClearDisplaySHow
    Call ClearPayInfo
    
    If chkCancel.Value = 1 Then '�˾ݵ�״̬
        chkCancel.Value = 0
    Else
        mstrInNO = "": txtModi.Text = ""
        mlngFirstID = 0: mstrFirstWin = ""
        
        If blnClearPatiInfor Then Call ClearPatientInfo(blnClearPatiInfor)
        Call ClearTotalInfo(True)
        
        Call InitCommVariable
        
        If mbytInFun = 0 And gbln�ۼ� Then
            txt�ۼ�.Text = Format(GetChargeTotal, "0.00")
        End If
    End If
    
    Call ClearBillRows
    Call ClearMoney
    Call SetDisible(True)
    Call NewBill(IIf(mblnStartFactUseType, False, True), blnClearPatiInfor, Not mbln����)
    If mbln���� Then
        With mobjBill
            .����ID = IIf(IsNull(mrsInfo!����ID), 0, mrsInfo!����ID)
            .��ҳID = IIf(mbln���� And mlng��ҳID <> 0, mlng��ҳID, Nvl(mrsInfo!��ҳID, 0))
            .��ʶ�� = IIf(gint������Դ = 2, Nvl(mrsInfo!סԺ��, 0), Nvl(mrsInfo!�����, 0))
            .���� = "" & mrsInfo!����
            .�Ա� = "" & mrsInfo!�Ա�
            .���� = Trim(txt����.Text) & IIf(IsNumeric(txt����.Text), cbo���䵥λ.Text, "")
            .���� = "" & mrsInfo!��ǰ����
            .����ID = IIf(mbln���� And mlngUnitID <> 0, mlngUnitID, Val(Nvl(mrsInfo!��ǰ����ID)))
            .����ID = IIf(mbln���� And mlngDeptID <> 0, mlngDeptID, Val(Nvl(mrsInfo!��ǰ����id)))
            .�ѱ� = zlStr.NeedName(cbo�ѱ�.Text) '�Ե�ǰ��ЧΪ׼
        End With
        Bill.SetFocus
    End If
    If blnNotClearYb And intInsure <> 0 Then
        mintInsure = intInsure: mstrYBBill = strYBBill: mstrYBPati = strYBPati
        mcur������� = cur�������: mcur����͸֧ = cur����͸֧: mblnYB�������� = blnYB��������
        sta.Panels(Pan.C3�����ʻ�).Text = "�����ʻ����:" & Format(mcur�������, "0.00")
        sta.Panels(Pan.C3�����ʻ�).Visible = True
        Call SetPatientEnableModi(False)
        '75259�����ϴ�,2014-7-10������������ʾ��ɫ����
        If Not mrsInfo Is Nothing Then
            If mrsInfo.State = 1 Then
                Call SetPatiColor(txtPatient, Nvl(mrsInfo!��������), vbRed)
            Else
                txtPatient.ForeColor = vbRed
            End If
        Else
            txtPatient.ForeColor = vbRed
        End If
        cmdAddBill.Enabled = blnAdd
    End If
    sta.Panels(Pan.C2��ʾ��Ϣ).Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    ClearFullBill = True
End Function


Private Function GetOneCardMoney(Optional ByVal str�շѽ��� As String) As Currency
'���ܣ���ȡ���е��ݻ�ĳ�ŵ���ʹ��һ��ͨ����Ľ��
'������str�շѽ���-��ǰ���ݵĽ��㴮
    Dim arrTmp As Variant, strTmp As String, i As Long
    
    If mblnOneCard = False Then Exit Function
    
    If str�շѽ��� <> "" Then
        arrTmp = Split(str�շѽ���, "||")
        For i = 0 To UBound(arrTmp)
            mrsOneCard.Filter = "���㷽ʽ='" & Split(arrTmp(i), "|")(0) & "'"
            If mrsOneCard.RecordCount > 0 Then
                GetOneCardMoney = Val(Split(arrTmp(i), "|")(1))
                Exit Function
            End If
        Next
    Else
        If mobjBill.Pages(1).�շѽ��� = "" Then
            mrsOneCard.Filter = "���㷽ʽ='" & zlStr.NeedName(cbo���㷽ʽ) & "'"
            If mrsOneCard.RecordCount > 0 Then GetOneCardMoney = GetMustPaySum
        Else
            arrTmp = Split(mobjBill.�շѽ���, "||")
            For i = 0 To UBound(arrTmp)
                mrsOneCard.Filter = "���㷽ʽ='" & Split(arrTmp(i), "|")(0) & "'"
                If mrsOneCard.RecordCount > 0 Then
                    GetOneCardMoney = Val(Split(arrTmp(i), "|")(1))
                    Exit Function
                End If
            Next
        End If
    End If
End Function
Private Function CheckMainOperation() As Boolean
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������������(�����������Ҫ����,�����ڸ�������,���ֹ
    '���:
    '����:lngRow-���ظ�����������
    '����:������������û�����븽������,����true,���򷵻�False
    '����:
    '�޸�:���˺�(�˺�ʱ,���Ӷ�λ����),���Ӳ���;strBackNo
    '����:2009/7/10
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngCount As Long, lngRow As Long   'ָ����
    Dim i As Long, p As Long
    lngCount = 0
    For p = 1 To mobjBill.Pages.Count
         For i = 1 To mobjBill.Pages(p).Details.Count
            lngCount = 0
            If mobjBill.Pages(p).Details(i).�շ���� = "F" Then
               If mobjBill.Pages(p).Details(i).���ӱ�־ = 0 Then lngCount = 0: Exit For  '������Ҫ����,�򲻼��,ֱ�ӷ���true
               lngCount = lngCount + 1  '��ʾ��������
               If lngRow <= 0 Then lngRow = i
            End If
        Next
        If lngCount > 0 Then Exit For
    Next
    If lngCount <> 0 Then
          MsgBox "�����в�����Ҫ����,�����ڸ�������,���飡", vbInformation, gstrSysName
          Err = 0: On Error GoTo Errhand:
          If p <= tbsBill.Tabs.Count Then tbsBill.Tabs(p).Selected = True
          '��λ��:
          Bill.Row = lngRow
          If Bill.Visible Then Bill.SetFocus
          Exit Function
    End If
    CheckMainOperation = True
Errhand:
End Function
Private Function WriteOff() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ʴ���
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-16 10:04:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strNos As String, i As Long
    Dim cllPro As Collection, cur���ѽ�� As Currency
    
    If Not (mbytInState = 3 Or (mbytInState = 0 And chkCancel.Value = 1 And chkCancel.Visible)) Then Exit Function
    If mbytInState = 0 And mstrInNO = "" Then
        MsgBox "û����ȷ��ȡ��������,����ִ�иò�����", vbInformation, gstrSysName
        cboNO.SetFocus: Exit Function
    End If
    For i = 1 To Bill.Rows - 1
        If Bill.TextMatrix(i, Bill.COLS - 1) = "��" And Bill.RowData(i) > 0 Then
            strSQL = strSQL & "," & Bill.RowData(i)
        End If
    Next
    If strSQL = "" Then
        MsgBox "������ѡ��һ��Ҫ�˷ѵ���Ŀ��", vbInformation, gstrSysName
        Bill.SetFocus: Exit Function
    End If
    Set cllPro = New Collection
    '������ѡ����
    strSQL = Mid(strSQL, 2)
    i = GetBillRows(mstrInNO, IIf(mbytInFun = 2, 2, 1))
    
    If UBound(Split(strSQL, ",")) + 1 = i Then strSQL = ""
    
    On Error GoTo errHandle
      If zlCheckIsMzToZY(mstrInNO, 2) Then
        MsgBox "ע��:" & vbCrLf & _
                      "    �õ����Ѿ����������תסԺ���� " & vbCrLf & _
                      "    ���Ѿ�������������תסԺ����,����������", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    '����:37307
    If (mbytBilling = 0 And Val(txt�ϼ�.Text) <> 0) And gbytԤ����˷��鿨 <> 0 Then
        cur���ѽ�� = Val(txt�ϼ�.Text)
        If Not zlDatabase.PatiIdentify(Me, glngSys, mobjBill.����ID, cur���ѽ��, mlngModul, 1, , , , , , (gbytԤ����˷��鿨 = 2)) Then Exit Function
    End If
    strSQL = "zl_������ʼ�¼_DELETE('" & mstrInNO & "','" & strSQL & "','" & UserInfo.��� & "','" & UserInfo.���� & "')"
    
    On Error GoTo errH
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    WriteOff = True
    
    '110319
    If mblnDrugMachine Then
        Dim rsTemp As ADODB.Recordset
        Dim strReturn As String, strData As String '���ﴦ����ҩ��ʽ������ID1,��ҩ����1;����ID2,��ҩ����2;...
        strSQL = "Select Id As ����id, -1 * Nvl(����, 1) * ���� As ��ҩ����" & vbNewLine & _
                " From ������ü�¼" & vbNewLine & _
                " Where ��¼���� = 2 And ��¼״̬ = 2 And NO = [1] And �շ���� In ('5', '6', '7')" & vbNewLine & _
                "       And �Ǽ�ʱ�� + 0 = (Select Max(�Ǽ�ʱ��)" & vbNewLine & _
                "                       From ������ü�¼" & vbNewLine & _
                "                       Where ��¼���� = 2 And ��¼״̬ = 2 And NO = [1])"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ�����˷���Ŀ", mstrInNO)
        Do While Not rsTemp.EOF
            strData = strData & ";" & Nvl(rsTemp!����id) & "," & Nvl(rsTemp!��ҩ����)
            rsTemp.MoveNext
        Loop
        If strData <> "" Then
            strData = Mid(strData, 2)
            Call mobjDrugMachine.Operation(gstrDBUser, Val("24-������ҩ(����/����)"), strData, strReturn)
        End If
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Exit Function
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub SetAllDelSelAll()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ѡ�������˷Ѽ�¼
    '����:���˺�
    '����:2011-08-30 14:55:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    For i = 1 To Bill.Rows - 1
        Bill.TextMatrix(i, Bill.COLS - 1) = "��"
    Next
    Call ReCalce�˿�
End Sub



Private Function DelChargeFee() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�˷Ѵ���
    '����:�˷ѳɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-16 10:07:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strNos As String, i As Long, cur���ѽ�� As Currency, lng����ID As Long
    Dim rsOneCard As ADODB.Recordset, rsThree As ADODB.Recordset
    Dim cllPro As Collection, bln���� As Boolean
    Dim cur��� As Currency, cur����� As Currency, cur��� As Currency
    Dim blnAll�����˷� As Boolean, blnCur�����˷� As Boolean, strInvoices As String
    Dim strBalance As String, strTmp As String, objICCard As Object, strCardNo As String
    Dim strInvoice As String, dtDateDel As Date, lng����ID As Long
    Dim blnTrans As Boolean, blnTransMedicare As Boolean, strAdvance As String
    Dim strErrMsg As String '������Ϣ
    Dim blnCommited As Boolean '�ѵ��ӿ�
    Dim dblErrMoney As Double   'δ�ɹ����׵Ľ��
    Dim cllBalance As Collection, cllUpdate As Collection, cllThreeSwap As Collection
    Dim blnExistThreeSwap As Boolean '���ڵ���������
    Dim blnExistOneCard As Boolean '����һ��ͨ����
    Dim str�˽��㷽ʽ As String, dbl�˽�� As Double
    Dim str���ս��� As String, strOtherBalance As String
    Dim strThreeBalance As String, blnҩƷ As Boolean, intCol As Integer
    Dim lng����ID As Long, lng����ID As Long, intCol��� As Integer
    Dim strReclaimInvoice As String, intInvoiceFormat As Integer '����Ʊ��:Ʊ�ݷ������Ϊ1��2ʱ��Ч25187
    Dim strReturn As String, strReturnRecipt As String '�˷Ѵ�����Ϣ����ʽ��NO,ҩ��ID|NO,ҩ��ID|��
    Dim bln��ȫ�˷� As Boolean
    
    If Not (mbytInState = 3 Or (mbytInState = 0 And chkCancel.Value = 1 And chkCancel.Visible)) Then Exit Function
    
    If mbytInState = 0 And mstrInNO = "" Then
        MsgBox "û����ȷ��ȡ��������,����ִ�иò�����", vbInformation, gstrSysName
        cboNO.SetFocus: Exit Function
    End If
    If CheckBillExistReplenishData(1, , mstrInNO) = True Then
        MsgBox "ѡ��ļ�¼������ҽ��������㣬����������ش�򲹴�Ʊ�ݲ�����", vbInformation, gstrSysName
        Exit Function
    End If
    
    '47400
    intCol��� = -1
    For intCol = 0 To Bill.COLS - 1
        Select Case Bill.TextMatrix(0, intCol)
        Case "���"
            intCol��� = intCol: Exit For
        Case Else
        End Select
    Next
    blnҩƷ = False
    For i = 1 To Bill.Rows - 1
        If Bill.TextMatrix(i, Bill.COLS - 1) = "��" And Bill.RowData(i) > 0 Then
            strSQL = strSQL & "," & Bill.RowData(i)
            If intCol��� <> -1 Then     '47400
                If Bill.TextMatrix(i, intCol���) Like "*��*ҩ*" _
                    Or Bill.TextMatrix(i, intCol���) Like "*��*ҩ*" _
                    Or Bill.TextMatrix(i, intCol���) Like "*����*" Then
                    blnҩƷ = True
                    '81190,Ƚ����,�˷�ҵ����ҩ���ϴ��˷���Ϣ
                    If Not Bill.TextMatrix(i, intCol���) Like "*����*" Then
                        If InStr(strReturnRecipt & "|", "|" & mstrInNO & "," & Bill.TextMatrix(i, BillCol.ִ�п���ID) & "|") = 0 Then
                            strReturnRecipt = strReturnRecipt & "|" & mstrInNO & "," & Bill.TextMatrix(i, BillCol.ִ�п���ID)
                        End If
                    End If
                End If
            End If
        End If
    Next
    
    If strSQL = "" Then
        MsgBox "������ѡ��һ��Ҫ�˷ѵ���Ŀ��", vbInformation, gstrSysName
        Bill.SetFocus: Exit Function
    End If
    Set cllPro = New Collection
    '������ѡ����
    strSQL = Mid(strSQL, 2)
    '47400
    If blnҩƷ Then
        If zlCheckDrugIsPutDrug(mstrInNO) = False Then Exit Function
    End If
    '��ȡ���λ���Ʊ��
    '���ݺ�1:���1(1..n);���ݺ�2:���2(1..n
    strReclaimInvoice = zlGetReclaimInvoice(mstrInNO & ":" & strSQL)
    
    i = GetBillRows(mstrInNO, IIf(mbytInFun = 2, 2, 1))
    
    If UBound(Split(strSQL, ",")) + 1 = i Then strSQL = ""
    
    On Error GoTo errHandle
    '���˺�:28947
    If mintInsure <> 0 Then
        If gclsInsure.CheckInsureValid(mintInsure) = False Then Exit Function
    End If
    
    '�����Ƕ൥���շ��е�һ��,��ȡ���NO,�շѲ����˷�ʱ�ջ���ǰ�ĵ���,ȫ���ش�ӡƱ��
    '����:51080,53145
    If gTy_Module_Para.bytƱ�ݷ������ <> 0 Then
        strNos = GetMultiNOs(mstrInNO, , , True)
    Else
        strNos = GetMultiNOs(mstrInNO, , , False)
    End If
    
    If zlCheckIsMzToZY(mstrInNO, 1) Then
        MsgBox "ע��:" & vbCrLf & _
                      "    �õ����Ѿ����������תסԺ���� " & vbCrLf & _
                      "    ���Ѿ�������������תסԺ����,�������˷�", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    '��ǰ���ݱ����Ƿ񲿷��˷�
    blnCur�����˷� = Not (BillDeleteAll(mstrInNO, 1, mblnHaveExcuteData) And strSQL = "")
    If blnCur�����˷� Then blnAll�����˷� = True '���ŵ���Ϊ�����˷�,�����е���Ϊ�����˷�
    If Not blnCur�����˷� Then bln��ȫ�˷� = True
    bln��ȫ�˷� = bln��ȫ�˷� And Not BillExistDelete(Replace(strNos, "'", ""), 1)

    If Not blnCur�����˷� Then
        strTmp = ""
        If UBound(Split(strNos, ",")) + 1 > 1 Then
            strTmp = Replace("," & strNos, ",'" & mstrInNO & "'", "")
            If Left(strTmp, 1) = "," Then strTmp = Mid(strTmp, 2)
            'ֻ�е����ŵ����е�����������ȫ��,�ҵ�ǰ����Ҳ��ȫ��ʱ,�Ų����ǲ�����
            If BillExistMoney(strTmp, 1) Then blnAll�����˷� = True
        End If
    End If
    If blnAll�����˷� Then
        If InStr(mstrPrivs, "�����˷�") = 0 Then
            MsgBox "��û��Ȩ��ִ�в����˷Ѳ�����", vbInformation, gstrSysName
            Exit Function
        End If
        If mintInsure > 0 And blnCur�����˷� Then
            If strSQL = "" Then
                MsgBox "���ŵ��ݰ������ս�����ã�������һЩ��Ŀ�����Ѿ�ִ�У����������˷ѡ�", vbInformation, gstrSysName
            Else
                MsgBox "���ŵ��ݰ������ս�����ã����������˷ѡ�", vbInformation, gstrSysName
            End If
            Call SetAllDelSelAll: Exit Function
        End If
        '���ŵ��ݲ����˷�ʱ�Ĺ����Ѽ��
        If gTy_Module_Para.bln������ Then
            MsgBox "�Զ���ȡ������ʱ���������˷ѡ�", vbInformation, gstrSysName: Exit Function
        End If
    End If
    strBalance = ""      '��¼ҽ���������˻صĽ��㷽ʽ
    strThreeBalance = ""
    Dim dblDelMoney As Double '��ǰ���㷽ʽ���˿���
    With vsBalance
        For i = 0 To .Rows - 1
            strTmp = vsBalance.TextMatrix(i, 0)
            dblDelMoney = Val(vsBalance.TextMatrix(i, 1))
            If strTmp <> "" Then
                '����:1-Ԥ���,2-ҽ��,3-ҽ�ƿ�,4-���㿨,5-һ��ͨ,0-������
                Select Case Val(.Cell(flexcpData, i, 0))
                Case 1
                Case 2  'ҽ����
                    str���ս��� = str���ս��� & "," & strTmp
                    If mintInsure <> 0 Then
                           If mblnYB�������� Then
                                If Not gclsInsure.GetCapability(support�����������, , mintInsure, strTmp) Then
                                    strBalance = strBalance & "," & strTmp
                                End If
                            Else    '��֧�������������ʱ,ֻ���������Ϊ�ֽ�,����ԭ����,������ҽ������
                                If strTmp = mstr�����ʻ� Then strBalance = "," & strTmp
                            End If
                    End If
                Case 3   '3-ҽ�ƿ�
                    '���֧������,��ֱ�����֣�����ֻ�˵�ǰӦ�˽��
                    If Val(.RowData(i)) = -1 And Val(.TextMatrix(i, 1)) = 0 Then
                        strThreeBalance = strThreeBalance & "," & strTmp
                    Else
                        blnExistThreeSwap = True
                    End If
                    If blnExistThreeSwap Then
                        If blnCur�����˷� Or Not bln��ȫ�˷� Then
                            If blnCur�����˷� And Not mTyDelFee.blnSingleBalance Or mTyDelFee.bln������ȫ�� Then
                                MsgBox "��ǰ����ʹ���˵��������㽻��,���ܽ��в����˷ѣ�", vbInformation, gstrSysName
                                Call SetAllDelSelAll: Exit Function
                            End If
                            strThreeBalance = strThreeBalance & "," & strTmp & "|" & dblDelMoney
                        End If
                        mTyDelFee.rsBlance.Filter = 0
                        mTyDelFee.rsBlance.Filter = "����=" & Val(.Cell(flexcpData, i, 0)) & " And ���㷽ʽ='" & strTmp & "'"
                        If mTyDelFee.rsBlance.EOF Then
                            MsgBox "�����ڵ�������������,����!", vbOKOnly + vbInformation, gstrSysName
                            Exit Function
                        End If
                        With mTyDelFee.rsBlance
                            If blnAll�����˷� And Val(Nvl(!�Ƿ�ȫ��)) = 1 And InStr(1, strNos, ",") > 0 Then
                                If Not mTyDelFee.blnSingleBalance Or Val(Nvl(!�Ƿ�ȫ��)) = 1 Then
                                    MsgBox "��ǰ����ʹ���˵��������㽻��,���е��ݱ���ȫ�ˣ�", vbInformation, gstrSysName
                                    Exit Function
                                End If
                            End If
                            '�����ֲŴ���
                            If zlCheckDelValied(Val(Nvl(!�����ID)), Nvl(!����), Val(Nvl(!����)) = 4, Nvl(!����), Nvl(!������ˮ��), Nvl(!����˵��), Original.����ID, dblDelMoney) = False Then Exit Function
                        End With
                    End If
                Case 4   '4-���㿨
                    '���֧������,��ֱ������
                    If Val(.RowData(i)) = -1 And Val(.TextMatrix(i, 1)) = 0 Then
                        strThreeBalance = strThreeBalance & "," & strTmp
                    Else
                        blnExistThreeSwap = True:
                    End If
                    If blnExistThreeSwap Then
                        If blnCur�����˷� Then
                            MsgBox "��ǰ����ʹ���˵��������㽻��,���ܽ��в����˷ѣ�", vbInformation, gstrSysName
                            Call SetAllDelSelAll: Exit Function
                        End If
                        mTyDelFee.rsBlance.Filter = 0
                        mTyDelFee.rsBlance.Filter = "����=" & Val(.Cell(flexcpData, i, 0)) & " And ���㷽ʽ='" & strTmp & "'"
                        If mTyDelFee.rsBlance.EOF Then
                            MsgBox "�����ڵ�������������,����!", vbOKOnly + vbInformation, gstrSysName
                            Exit Function
                        End If
                        With mTyDelFee.rsBlance
                            If blnAll�����˷� And Val(Nvl(!�Ƿ�ȫ��)) = 1 And InStr(1, strNos, ",") > 0 Then
                                MsgBox "��ǰ����ʹ���˵��������㽻��,���е��ݱ���ȫ�ˣ�", vbInformation, gstrSysName
                                Exit Function
                            End If
                            '�����ֲŴ���
                            If zlCheckDelValied(Val(Nvl(!�����ID)), Nvl(!����), Val(Nvl(!����)) = 4, Nvl(!����), Nvl(!������ˮ��), Nvl(!����˵��), Original.����ID, Val(Nvl(!������))) = False Then Exit Function
                        End With
                    End If
                Case 5 'һ��ͨ
                    If Val(.RowData(i)) <> -1 And Val(.TextMatrix(i, 1)) = 0 Then
                        strThreeBalance = strThreeBalance & "," & strTmp
                    Else
                        blnExistOneCard = True
                    End If
                    
                    If blnExistOneCard Then
                        If blnCur�����˷� Then
                            MsgBox "��ǰ����ʹ����һ��ͨ����,���ܽ��в����˷ѣ�", vbInformation, gstrSysName
                            Call SetAllDelSelAll: Exit Function
                        End If
                        mTyDelFee.rsBlance.Filter = 0
                        mTyDelFee.rsBlance.Filter = "����=5"
                        If mTyDelFee.rsBlance.RecordCount = 0 Then
                            MsgBox "δ�ҵ�һ��ͨ�Ľ�������,����!", vbInformation + vbOKOnly, gstrSysName
                            Exit Function
                        End If
                        '���һ��ͨ�Ƿ�Ϸ�
                        On Error Resume Next
                        Set objICCard = CreateObject("zlICCard.clsICCard")
                        On Error GoTo 0
                        If objICCard Is Nothing Then
                            MsgBox "һ��ͨ�ӿڴ���ʧ��,���ܽ����˷�!����ӿ��ļ�.", vbInformation, gstrSysName
                            Exit Function
                        End If
                        'gobjSquare.objSquareCard
                        'strCardNo = objICCard.Read_Card(Me)
                        '����ˢ������
                        'zlBrushCard(frmMain As Object, _
                        'ByVal lngModule As Long, _
                        'ByVal rsClassMoney As ADODB.Recordset, _
                        'ByVal lngCardTypeID As Long, _
                        'ByVal bln���ѿ� As Boolean, _
                        'ByVal strPatiName As String, ByVal strSex As String, _
                        'ByVal strOld As String, ByVal dbl��� As Double, _
                        'Optional ByRef strCardNo As String, _
                        'Optional ByRef strPassWord As String) As Boolean
                        If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, Nothing, 0, False, _
                          mobjBill.����, mobjBill.�Ա�, mobjBill.����, 0, strCardNo, "") = False Then Exit Function
                        If strCardNo = "" Then Exit Function
                        If strCardNo <> Nvl(mTyDelFee.rsBlance!����) Then
                            MsgBox "��ǰ������ۿ�Ų�һ��,���ܽ����˷�.", vbInformation, gstrSysName
                            Exit Function
                        End If
                    End If
                Case Else
                        strOtherBalance = strOtherBalance & "," & strTmp
                End Select
            End If
        Next
    End With
    If strBalance <> "" Then strBalance = Mid(strBalance, 2)
    If strThreeBalance <> "" Then strThreeBalance = Mid(strThreeBalance, 2)
    If strOtherBalance <> "" Then strOtherBalance = Mid(strOtherBalance, 2)
    
    '����:37307
    If Val(txtԤ�����.Text) <> 0 And gbytԤ����˷��鿨 <> 0 Then
            cur���ѽ�� = Val(txtԤ�����.Text)
        If Not zlDatabase.PatiIdentify(Me, glngSys, mobjBill.����ID, cur���ѽ��, mlngModul, 1, , , , , , (gbytԤ����˷��鿨 = 2)) Then Exit Function
    End If
    If strReclaimInvoice <> "" Then
        '��Ҫ��ʾ��������Ҫ���յķ�Ʊ
        If InStr(1, mstrPrivs, "�˷Ѻ��շ�Ʊ") > 0 Then
            If MsgBox("ע��:" & vbCrLf & " ��ǰ�˷���Ҫ�������·�Ʊ:" & vbCrLf & strReclaimInvoice, vbQuestion + vbDefaultButton1 + vbYesNo, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
        strInvoices = ""
    Else
        If blnAll�����˷� Then
            If InStr(1, mstrPrivs, "�˷Ѻ��շ�Ʊ") > 0 Then
                If frmReInvoice.ShowMe(Me, mstrInNO, Val(txt�ϼ�.Text), Val(txtӦ��.Text), strInvoices) = False Then Exit Function
            End If
        End If
        If Not blnAll�����˷� And strBalance = "" And strThreeBalance = "" Then strOtherBalance = ""
    End If
    
    '�����ҽ��,��ǰ���ݱ���ȫ��,����в�֧���˻ص�ҽ������,��Ҫ�������
    '�����˷Ѳ����������:����ǵ�һ���˷���ȫ���˷�,�򲻴������
    '60974
'    If mintInsure <> 0 Then
'        If strBalance <> "" Then
'            Call GetDelMoney(cur�����)
'        End If
'    ElseIf BillExistDelete(mstrInNO, 1) Or blnCur�����˷� Then
'        Call GetDelMoney(cur�����)
'    End If
    Call GetDelMoney(cur�����)
    If mintInsure <> 0 And MCPAR.ҽ���ӿڴ�ӡƱ�� Then
        If zlGetInvoiceGroupUseID(lng����ID) = False Then Exit Function
        strInvoice = GetNextBill(lng����ID)
    End If
    dtDateDel = zlDatabase.Currentdate
    lng����ID = zlDatabase.GetNextId("���˽��ʼ�¼")
    
    
    '����:43403
    '����SQL���,������ǲ����˷ѣ���Ҫ�ջ�Ʊ�ݣ��ǲ����˷����ڵ��ô�ӡʱ�ջز��ش�
    strSQL = "zl_�����շѼ�¼_DELETE('" & mstrInNO & "','" & UserInfo.��� & "','" & UserInfo.���� & "'," & _
        "'" & strBalance & "','" & strSQL & "','" & zlStr.NeedName(cbo���㷽ʽ.Text) & "'," & cur����� & _
        ",To_Date('" & Format(dtDateDel, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & IIf(blnAll�����˷�, "1", "0") & _
        IIf(Trim(txt�˷�ժҪ.Text) <> "", ",'" & Trim(txt�˷�ժҪ.Text) & "'", ",NULL") & ","
    '     У�Ա�־_In: 0-����Ҫ�϶�;1-��϶�(��������Ա�ɿ����,������Ʊ��)
    strSQL = strSQL & "1," & lng����ID & "," & lng����ID & ",'" & strThreeBalance & IIf(strThreeBalance <> "", ",", "") & strOtherBalance & "')"
    zlAddArray cllPro, strSQL
    
    '�Ȳ���Ʊ�ݣ�ҽ���ӿڲ���ȡ��
    If mintInsure <> 0 And MCPAR.ҽ���ӿڴ�ӡƱ�� And (gTy_Module_Para.bytƱ�ݷ������ = 0 Or _
        gTy_Module_Para.bytƱ�ݷ������ <> 0 And strReclaimInvoice = "") Then
        '25187:���˷���ɺ�,�ٽ����ջ��ش�(�Ѿ���frmPrint�д�����),���Բ��ٴ���:strReclaimInvoice=""
        strSQL = "zl_�����շѼ�¼_RePrint('" & mstrInNO & "','" & strInvoice & "'," & ZVal(lng����ID) & ",'" & UserInfo.���� & "'," & _
            "To_Date('" & Format(dtDateDel, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),1,1)"
       zlAddArray cllPro, strSQL
    End If
    
'    '�������
'    '   �����˷ѵ�����
'60974
'    If cur����� <> 0 Then
'        strSql = "zl_�����շ����_Insert('" & mstrInNO & "'," & cur����� & ",1)"
'       zlAddArray cllPro, strSql
'        vsBalance.ToolTipText = "ҽ�����㷽ʽ" '�ָ���Ϣ,֮ǰ��¼�������
'    End If
    If gblnBillPrint Then
        If gobjBillPrint.zlEraseBill(strNos, 0) = False Then Exit Function
    End If
    cmdOK.Enabled = False   '��ֹ���ô�ӡ�������ķ�ģ̬����,�Լ�ҽ����ʱ
    On Error GoTo errH
    bln���� = False
    If cbo���㷽ʽ.ListIndex >= 0 Then
        bln���� = cbo���㷽ʽ.ItemData(cbo���㷽ʽ.ListIndex) = 1
        str�˽��㷽ʽ = zlStr.NeedName(cbo���㷽ʽ.Text)
    Else
        bln���� = True
        If mrs���㷽ʽ Is Nothing Then
            str�˽��㷽ʽ = "�ֽ�"
        Else
            mrs���㷽ʽ.Filter = "����=1"
            If mrs���㷽ʽ.EOF Then
                str�˽��㷽ʽ = "�ֽ�"
            Else
                str�˽��㷽ʽ = Nvl(mrs���㷽ʽ!����, "�ֽ�")
            End If
        End If
    End If
    strErrMsg = "": dblErrMoney = 0
    '1.ִ���˷�
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    'ҽ������
    If mintInsure <> 0 And mblnYB�������� Then
        'ҽ���˷Ѵ���
        If Not DelInsure(blnExistThreeSwap, mintInsure, str���ս���, Original.����ID, mstrInNO, _
            mobjBill.Pages(1).ʵ�ս��, Val(txtԤ�����.Text), str�˽��㷽ʽ, bln����, dbl�˽��) Then gcnOracle.RollbackTrans: Exit Function
        gcnOracle.CommitTrans: blnTrans = False: blnCommited = True
        
        '�����ڵ��������׵ģ�ֱ���˳�
        If Not blnExistThreeSwap Then
            If dbl�˽�� <> 0 Then
                MsgBox "Ӧ�˽��" & vbCrLf & str�˽��㷽ʽ & "��" & Format(dbl�˽��, "0.00") & "Ԫ", vbInformation, gstrSysName
            End If
            GoTo PrintBill:
        End If
        gcnOracle.BeginTrans: blnTrans = True
    End If
    
    '��һ��ͨ
    blnCommited = False  '54949
    If DelOneCardSwap(objICCard, blnCommited) = False Then Exit Function
    
    '�˵������ӿڽ���
    blnCommited = False '54949
    If DelTreeSwap(blnCommited, lng����ID) = False Then Exit Function
    gcnOracle.CommitTrans: blnTrans = False
    If strErrMsg <> "" Then
        MsgBox "����������ʧ���ܶ�Ϊ:" & Format(dblErrMoney, "0.00") & " �벹�����㽻�׽ӿ�." & vbCrLf & strErrMsg, vbInformation, gstrSysName
        cmdOK.Enabled = True: Exit Function
    End If
    
PrintBill:
    If OverFeeDel(lng����ID, mobjBill.����ID) = False Then Exit Function
    '81190,Ƚ����,�˷�ҵ����ҩ���ϴ��˷���Ϣ
    On Error Resume Next
    If mblnDrugPacker Then
        If strReturnRecipt <> "" Then
            strReturnRecipt = Mid(strReturnRecipt, 2)
            Call mobjDrugPacker.DYEY_MZ_TransRecipeReturn(1, UserInfo.���, UserInfo.����, strReturnRecipt, strReturn)
        End If
    End If
    Err.Clear: On Error GoTo errHandle
    
    lng����ID = mobjBill.����ID
    Call PrintDelBill(strNos, dtDateDel, lng����ID, blnAll�����˷�, strInvoices, strReclaimInvoice)
    '56615
    Call WriteMzInforToCard(lng����ID, lng����ID, True)
    cmdOK.Enabled = True   '��ֹ���ô�ӡ�������ķ�ģ̬����,�Լ�ҽ����ʱ
    DelChargeFee = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Exit Function
errH:
    If blnTrans Then
        gcnOracle.RollbackTrans
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
    cmdOK.Enabled = True
    Call SaveErrLog
End Function
Private Sub PrintDelBill(ByVal strNos As String, ByVal dtDateDel As Date, ByVal lng����ID As Long, ByVal blnAll�����˷� As Boolean, _
    ByVal strInvoices As String, ByVal strReclaimInvoice As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ӡ�˷ѵ���
    '���:strNos-�˷ѵ��ݺ�
    '����:���˺�
    '����:2013-05-27 11:47:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intInvoiceFormat As Integer, strSQL As String
    Dim strPriceGrade As String
    
    On Error GoTo errHandle
    '�����˷�ʱ�ջز��ش�
    If Not blnAll�����˷� Then
         '˰�ز���ȫ��ʱ�ջش���(ȫ��ʱ��zl_�����շѼ�¼_DELETE�����ջ�Ʊ��)
        If Not gobjTax Is Nothing And gblnTax Then
            gstrTax = gobjTax.zlTaxOutErase(gcnOracle, strNos)
            If gstrTax <> "" Then MsgBox gstrTax, vbExclamation, gstrSysName
        End If
        '��ӡ�ص�
        GoTo PrintList:
        Exit Sub
    End If
    
    If mblnPrint Then
        If gintPriceGradeStartType >= 2 Then
            strPriceGrade = GetPriceGradeFromNos(strNos)
        Else
            strPriceGrade = mstr��ͨ�۸�ȼ�
        End If
    End If
    
    If gTy_Module_Para.bytƱ�ݷ������ <> 0 And strReclaimInvoice <> "" Then
        '����Ʊ�ݷ�������ӡ
        '��Ԥ��,��Ʊ���Ƿ����
        Dim str��Ʊ�� As String, intƱ������ As Integer
        str��Ʊ�� = strReclaimInvoice
        If zlExeCuteBillNoSplit(True, 4, mlng����ID, strNos, lng����ID, "", dtDateDel, 1, str��Ʊ��, intƱ������) = False Then GoTo PrintList:
        If intƱ������ = 0 Then
            'ֻ����Ʊ��,������ӡ
            str��Ʊ�� = strReclaimInvoice
            Call zlExeCuteBillNoSplit(False, 4, mlng����ID, strNos, lng����ID, "", dtDateDel, 1, str��Ʊ��, intƱ������)
            GoTo PrintList:
        End If
        
        mblnPrint = True
        ''0-����ӡ;1-�Զ���ӡ;2-��ʾ��ӡ
        If mintInvoicePrint = 0 Then mblnPrint = False   '�Զ���ӡ
        If mintInvoicePrint = 2 Then
            If MsgBox("�Ƿ��ӡƱ�ݣ�", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then mblnPrint = False
        End If
        
        '�ش��ջط�Ʊ
        If mblnPrint Then
            intInvoiceFormat = IIf(strReclaimInvoice = "" And gTy_Module_Para.bytƱ�ݷ������ <> 0, mintOldInvoiceFormat, mintInvoiceFormat)
            Call RePrintCharge(1, strNos, Me, mlng����ID, strReclaimInvoice, True, dtDateDel, _
                intInvoiceFormat, , , mlngShareUseID, mstrUseType, , strPriceGrade)
        End If
        GoTo PrintList:
        Exit Sub
    End If
    
    If strInvoices = "" Then  '�ջز����´�ӡ�����վ�
       '0-����ӡ;1-�Զ���ӡ;2-��ʾ��ӡ
        mblnPrint = True
        ''0-����ӡ;1-�Զ���ӡ;2-��ʾ��ӡ
        If mintInvoicePrint = 0 Then mblnPrint = False   '�Զ���ӡ
        If mintInvoicePrint = 2 Then
            If MsgBox("�Ƿ��ӡƱ�ݣ�", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then mblnPrint = False
        End If
        If mblnPrint Then
            intInvoiceFormat = IIf(strReclaimInvoice = "" And gTy_Module_Para.bytƱ�ݷ������ <> 0, mintOldInvoiceFormat, mintInvoiceFormat)
            Call RePrintCharge(1, strNos, Me, mlng����ID, strReclaimInvoice, True, dtDateDel, _
                intInvoiceFormat, , , mlngShareUseID, mstrUseType, , strPriceGrade)
        End If
        GoTo PrintList:
        Exit Sub
    End If
    
    'b.�շѻ���һ����ʱû�д�ӡƱ��
    If strInvoices <> "�޿���Ʊ��" Then
        'c.ֻ�ջ�Ʊ��
        strSQL = "zl_�����շѼ�¼_RePrint('" & mstrInNO & "',Null,0,'" & UserInfo.���� & "'," & _
            "To_Date('" & Format(dtDateDel, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),1,0,'" & strInvoices & "')"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    End If
 
PrintList:
    '��ӡ�����嵥
    If blnAll�����˷� Then
        If InStr(mstrPrivs, "��ӡ�嵥") > 0 Then
            If gint�շ��嵥 = 1 Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_3", Me, "NO=" & strNos, "ҩƷ��λ=" & IIf(gblnҩ����λ, 1, 0), 2)
            ElseIf gint�շ��嵥 = 2 Then
                If MsgBox("Ҫ��ӡ�շ��嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_3", Me, "NO=" & strNos, "ҩƷ��λ=" & IIf(gblnҩ����λ, 1, 0), 2)
                End If
            End If
        End If
    End If
    If mintInsure <> 0 And MCPAR.�˷Ѻ��ӡ�ص� And InStr(1, mstrPrivs, "ҽ���˷ѻص�") > 0 Then
        '����:35248
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_4", Me, "NO=" & strNos, 2)
    End If
    If mint�˷ѻص���ӡ = 1 Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_5", Me, "NO=" & strNos, 2)
    ElseIf mint�˷ѻص���ӡ = 2 Then
        If MsgBox("�Ƿ��ӡ�˷ѻص���", vbYesNo + vbQuestion, gstrSysName) = vbYes Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_5", Me, "NO=" & strNos, 2)
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function DelOneCardSwap(ByRef objICCard As Object, blnCommited As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:һ��ͨ���óɹ�
    '���:
    '����:blnCommited-�ύ��һ�γɹ���,����true
    '����:���óɹ�, ����true,���򷵻�False(Falseʱ,�������˵�)
    '����:���˺�
    '����:2011-08-30 12:22:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strErrMsg As String, dblErrMoney As Double
    Dim strSuccess As String
    
    On Error GoTo errHandle
    With mTyDelFee.rsBlance
        .Filter = "����=5 "
        '����:1-Ԥ���,2-ҽ��,3-ҽ�ƿ�,4-���㿨,5-һ��ͨ,0-������
        If .RecordCount = 0 Then DelOneCardSwap = True: Exit Function
        Do While Not .EOF
            If Not DelOneCardMoney(objICCard, Nvl(!NO), Nvl(!����), Nvl(!������ˮ��), Nvl(!ҽԺ����), Val(Nvl(!������))) Then
                gcnOracle.RollbackTrans
                If blnCommited Then
                    dblErrMoney = dblErrMoney + Val(Nvl(!������))
                    strErrMsg = strErrMsg & vbCrLf & "    " & Nvl(!����, "һ��ͨ") & ":" & Val(Nvl(!������))
                Else
                    MsgBox "һ��ͨ�˷ѽ��׵���ʧ��,�˷Ѳ���ʧ�ܣ�", vbExclamation, gstrSysName
                    cmdOK.Enabled = True: Exit Function
                End If
             Else
                gcnOracle.CommitTrans: gcnOracle.BeginTrans: blnCommited = True
                strSuccess = strSuccess & vbCrLf & "    " & Nvl(!����, "һ��ͨ") & ":" & Val(Nvl(!������))
             End If
        Loop
     End With
    If strErrMsg <> "" Then '54949
       gcnOracle.RollbackTrans
        MsgBox "    ��һ��ͨ�˷�ʱʧ��(�ܶ�Ϊ:" & Format(dblErrMoney, "0.00") & "), " & vbCrLf & _
                      "�����¶��쳣�����˷�,ʧ�ܽӿ�����:" & vbCrLf & _
                      strErrMsg & vbCrLf & _
                      "   ��һ��ͨ�˷�ʱ,�ɹ��Ľ�������:" & vbCrLf & _
                      strSuccess, vbExclamation, gstrSysName
        cmdOK.Enabled = True: Exit Function
    End If
    DelOneCardSwap = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    Call ErrCenter
    SaveErrLog
End Function
Private Function DelTreeSwap(ByRef blnCommited As Boolean, ByVal lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�˵���������
    '���;blnCommited-�Ƿ��Ѿ����ύ��
    '����:�˳ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-30 11:36:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllBalance  As New Collection, cllUpdate As New Collection, cllThreeSwap As New Collection
    Dim lngRow As Long, str���㷽ʽ As String, dblMoney As Double, strSQL As String
    Dim blnYes As Boolean   '�Ƿ���ýӿ�
    Dim strErrMsg As String, dblErrMoney As Double
    Dim strSucces As String
    
    On Error GoTo errHandle
    With mTyDelFee.rsBlance
        .Filter = "����=3 or ����=4"
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
           '�����ID,����,�Ƿ����ѿ�(1-��;0-��),������ˮ��,����˵��,strNO
           blnYes = True
           str���㷽ʽ = Nvl(!���㷽ʽ)
'            If Val(Nvl(!�Ƿ�����)) = 1 Then
                blnYes = False
                With vsBalance
                    For lngRow = 0 To .Rows - 1
                        If str���㷽ʽ = .TextMatrix(lngRow, 0) And Val(.TextMatrix(lngRow, 1)) <> 0 Then
                            '����:1-Ԥ���,2-ҽ��,3-ҽ�ƿ�,4-���㿨,5-һ��ͨ,0-������
                            Select Case Val(.Cell(flexcpData, lngRow, 0))
                            Case 3, 4
'                                If .RowData(lngRow) = -1 And .RowHidden(lngRow) = False Then
                                     blnYes = True
                                     dblMoney = Val(.TextMatrix(lngRow, 1))
                                     Exit For
'                                End If
                            End Select
                        End If
                    Next
                End With
'            End If
            If blnYes Then
                Set cllBalance = New Collection
                Set cllUpdate = New Collection: Set cllThreeSwap = New Collection
                'cllBalance.Add Array(Val(Nvl(rsTmp!�����ID)), Trim(Nvl(rsTmp!����)), IIf(Val(Nvl(rsTmp!���㿨���)) <> 0, 1, 0), Trim(Nvl(rsTmp!������ˮ��)), Trim(Nvl(rsTmp!����˵��))), strNO
                '�����ID,����,�Ƿ����ѿ�(1-��;0-��),������ˮ��,����˵��,strNO
                cllBalance.Add Array(Val(Nvl(!�����ID)), Nvl(!����), IIf(Val(Nvl(!����)) = 4, 1, 0), Nvl(!������ˮ��), Nvl(!����˵��)), mstrInNO
                
                If CallBackBalanceInterface(cllBalance, Original.����ID, lng����ID, mstrInNO, dblMoney, cllUpdate, cllThreeSwap, strErrMsg) = False Then
                    gcnOracle.RollbackTrans
                    If blnCommited Then
                        dblErrMoney = dblErrMoney + dblMoney
                        strErrMsg = strErrMsg & vbCrLf & "    " & Nvl(!����) & ":" & dblMoney
                    Else
                        MsgBox "�����������׽ӿ�ʱ,�˷ѽ��׵���ʧ��!" & vbCrLf & "    " & Nvl(!����) & ":" & Format(dblMoney, "0.00"), vbExclamation, gstrSysName
                        cmdOK.Enabled = True: Exit Function
                    End If
                Else
                        '��������
                        'Zl_�����շ�_���У��
                        strSQL = "Zl_�����շ�_���У��("
                        '  No_In       ������ü�¼.NO%Type,
                        strSQL = strSQL & "'" & mstrInNO & "',"
                        '  ��������_In Number, 0-һ��ͨ;1-���ѿ�;2-ҽ�ƿ�
                        strSQL = strSQL & "" & IIf(Val(Nvl(!����)) = 4, 1, 2) & ","
                        '  �����id_In ����Ԥ����¼.�����id%Type,
                        strSQL = strSQL & "" & Val(Nvl(!�����ID)) & ","
                        '  ����_In     ����Ԥ����¼.����%Type
                        strSQL = strSQL & "'" & Nvl(!����) & "')"
                        zlDatabase.ExecuteProcedure strSQL, Me.Caption
                        'cllUpdate:�Ѿ���Delete��ִ��,�����ٸ���
                         gcnOracle.CommitTrans
                        Call SaveThreeData(cllThreeSwap)
                         gcnOracle.BeginTrans: blnCommited = True
                         strSucces = strSucces & vbCrLf & "    " & Nvl(!����) & ":" & dblMoney
                         '������ص��˷Ѽ��
                End If
            End If
            .MoveNext
        Loop
    End With
    If strErrMsg <> "" Then '54949
       gcnOracle.RollbackTrans
        MsgBox "    �����������˷�ʱʧ��(�ܶ�Ϊ:" & Format(dblErrMoney, "0.00") & "), " & vbCrLf & _
                      "�����¶��쳣�����˷�,ʧ�ܽӿ�����:" & vbCrLf & _
                      strErrMsg & vbCrLf & _
                      "   �����������˷�ʱ,�ɹ��Ľ�������:" & vbCrLf & _
                      strSucces, vbExclamation, gstrSysName
        cmdOK.Enabled = True: Exit Function
    End If
    DelTreeSwap = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
End Function
Private Function OverFeeDel(ByVal str����IDs As String, ByVal lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����˷��շ�
    '���:strNos-����շѵĵ���(����Ϊ����,��Ŀǰֻ��һ�ŵ���)
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-29 14:50:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo errHandle
    If Left(str����IDs, 1) = "," Then str����IDs = Mid(str����IDs, 2)
    ' Zl_�����շѽ���_����˷�
    strSQL = "Zl_�����շѽ���_����˷�("
    '  ����id_In       ������ü�¼.����id%Type,
    strSQL = strSQL & "" & lng����ID & ","
    '  �˷ѽ������_In ����Ԥ����¼.�������%Type,
    strSQL = strSQL & "NULL,"
    '  ����ids_In      Varchar2,
    strSQL = strSQL & "'" & str����IDs & "',"
    '  ����Ա����_In   ����Ԥ����¼.����Ա����%Type := Null
    strSQL = strSQL & "'" & UserInfo.���� & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption

    OverFeeDel = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function DelOneCardMoney(ByVal objICCard As Object, ByVal strNo As String, ByVal strCardNo As String, _
    ByVal strSwapNO As String, ByVal strHsptName As String, _
    ByVal dblMoney As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:һ��ͨ�˷�
    '���:strCardNo-����
    '        strSwapNo-������ˮ��
    '        strHsptName-ҽԺ����
    '        dblMoney-�˷ѽ��
    '����:
    '����:���׳ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-29 00:07:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo errHandle
    
     If Not objICCard.ReturnSwap(strCardNo, strHsptName, strSwapNO, dblMoney) Then Exit Function
    '����һ��ͨ��У�Ա�־
    'Zl_�����շ�_��ɽ϶�
    strSQL = "Zl_�����շ�_���У��("
    '  No_In       ������ü�¼.NO%Type,
    strSQL = strSQL & "'" & strNo & "',"
    '  ��������_In Number, 0-һ��ͨ;1-���ѿ�;2-ҽ�ƿ�
    strSQL = strSQL & "1,"
    '  �����id_In ����Ԥ����¼.�����id%Type,
    strSQL = strSQL & "NULL,"
    '  ����_In     ����Ԥ����¼.����%Type
    strSQL = strSQL & "'" & strCardNo & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    DelOneCardMoney = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function
Private Function DelInsure(ByVal blnExistThreeSwap As Boolean, _
    ByVal intInsure As Integer, ByVal strҽ������ As String, _
    ByVal lng����ID As Long, ByVal strNo As String, _
    ByVal dblʵ�ս�� As Double, ByVal dbl���γ�Ԥ�� As Double, _
    ByVal str�˽��㷽ʽ As String, ByVal bln���� As Boolean, ByRef cur��� As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����ݺŵ��ýӿ�
    '���:str�˽��㷽ʽ-�˻ص�ָ�����㷽ʽ
    '       bln����-�Ƿ��˻��ֽ�
    '       blnExistThreeSwap-�Ƿ���ڵ�����������
    '       strҽ������-ҽ������,�Զ��ŷ���(�����ʻ�,ҽ���ʻ�)
    '����:cur���˷ѽ��
    '����:���óɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-25 12:21:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cur��� As Currency, strSQL As String, strAdvance As String
    Dim lng����ID As Long, i As Long, cur����� As Double, str�շѽ��� As String, blnTransMedicare As Boolean
    Err = 0: On Error GoTo Errhand:
    If blnExistThreeSwap Then
        ' Zl_�������_�϶Ա�־_Update
        strSQL = "Zl_�������_�϶Ա�־_Update("
        '  ����id_In     ������ü�¼.����id%Type,
        strSQL = strSQL & "" & lng����ID & ","
        '  �������id_In ����Ԥ����¼.�������%Type,
        strSQL = strSQL & "NULL,"
        '  �շѽ���_In   Varchar2,
        strSQL = strSQL & "'" & strҽ������ & "',"
        '  �����id_In   ����Ԥ����¼.�����id%Type := Null,
        strSQL = strSQL & "NULL,"
        '  ���ѿ�_In     Integer := 0,
        strSQL = strSQL & "0,"
        '  ����_In       ����Ԥ����¼.����%Type := Null,
        strSQL = strSQL & "NULL,"
        '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
        strSQL = strSQL & "NULL,"
        '  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
        strSQL = strSQL & "NULL,"
        '  У�Ա�־_In   ����Ԥ����¼.У�Ա�־%Type := 0
        strSQL = strSQL & "2)"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    End If
    
    strAdvance = "1|1"
    If Not gclsInsure.ClinicDelSwap(lng����ID, , intInsure, strAdvance) Then Exit Function
    blnTransMedicare = True
    If strAdvance = "1|1" Or strAdvance = "" Then
        Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, True, intInsure)
        DelInsure = True: Exit Function
     End If
    '���ݷ��صĽ�����Ϣ������Ԥ����¼
    '   strAdvance���ظ�ʽ:���㷽ʽ1|���||���㷽ʽ2:���...
    cur��� = 0
    For i = 0 To UBound(Split(strAdvance, "||"))
        cur��� = cur��� + -1 * Split(Split(strAdvance, "||")(i), "|")(1)
    Next
    cur��� = dblʵ�ս�� - cur��� - Val(txtԤ�����.Text)
    cur��� = cur���: cur����� = 0
    '��Ϊָ���Ľ��㷽ʽ��������ֽ𣬿��ܲ����µ������
    If bln���� Then
        cur��� = Format(CentMoney(cur���), "0.00")
        cur����� = cur��� - cur���
    End If
    str�շѽ��� = str�˽��㷽ʽ & "|" & -1 * cur��� & "| "
    lng����ID = GetDelBalanceID(strNo)
    If Not blnExistThreeSwap Then
        strSQL = "zl_�����շѽ���_Update(" & lng����ID & ",'" & str�շѽ��� & "'," & -1 * dbl���γ�Ԥ�� & ",'" & _
            strAdvance & "'," & -1 * cur����� & ",NULL,NULL,NULL,1)"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, True, intInsure)
        DelInsure = True: Exit Function
    End If
    'Zl_ҽ������У��_Update
    strSQL = "Zl_ҽ������У��_Update("
    '  ����id_In   ������ü�¼.����id%Type,
    strSQL = strSQL & "" & lng����ID & ","
    '  ���ս���_In Varchar2
    strSQL = strSQL & "'" & strAdvance & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, True, mintInsure)
    
    DelInsure = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
    'ҽ����HIS����ͬһ������,HIS����ʧ��,��ҽ���������ϴ�,������Ҫ��"ȡ������"�ӿ�
    If blnTransMedicare Then Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, False, mintInsure)
End Function

Private Function DelCharge() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�˷ѻ�����
    '����:�˷ѳɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-16 09:35:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strNos As String, i As Long
    Dim cllPro As Collection
     
    If mbytInFun = 2 Then '����
        If WriteOff = False Then Exit Function
    ElseIf mbytInFun = 0 Then '�˷�
        If DelChargeFee = False Then Exit Function
    End If
    '��ɺ�Ľ��洦��
    DelCharge = True
    If mbytInState <> 0 Then Unload Me: Exit Function
    If mbytInFun = 0 And gbln�ۼ� Then
        txt�ۼ�.Text = Format(GetChargeTotal, "0.00")
    End If
    mstrInNO = "": cboNO.Text = "": txtInvoice.Text = ""
    Call ClearBillRows: Call ClearMoney
    chkCancel.Value = 0
    Call ClearPatientInfo(True)
    Call ClearTotalInfo
    Call SetDisible(True)
    Call NewBill(IIf(mblnStartFactUseType, False, True))
    If mbytBilling = 2 Then
        cboNO.SetFocus
    Else
        txtPatient.SetFocus
    End If
    mblnSaveData = True
End Function
Private Function SaveVerify() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ʻ��۵���˲���
    '����:��˳ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-16 10:17:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, i As Long
    
    On Error GoTo errHandle
    If Not (mbytInFun = 2 And mbytBilling = 2) Then Exit Function
    If mstrInNO = "" Then
        MsgBox "û�м��ʻ��۵���,�������룡", vbInformation, gstrSysName
        cboNO.SetFocus: Exit Function
    End If
    'ȡ������˵������
    strSQL = ""
    For i = 1 To Bill.Rows - 1
        If Bill.RowData(i) > 0 Then
            strSQL = strSQL & "," & Bill.RowData(i)
        End If
    Next
    strSQL = Mid(strSQL, 2)
    i = GetBillRows(mstrInNO, 2)
    If UBound(Split(strSQL, ",")) + 1 = i Then strSQL = ""
    If Val(txt�ϼ�.Text) <> 0 And gdblԤ��������鿨 <> 0 Then
        If Not zlDatabase.PatiIdentify(Me, glngSys, mobjBill.����ID, Val(txt�ϼ�.Text), mlngModul, 1, , IIf(-1 * gdblԤ��������鿨 >= Val(txt�ϼ�.Text), False, True), , , , (gdblԤ��������鿨 = 2)) Then Exit Function
    End If
    '���ñ���
    If Not AuditingWarn(mstrPrivs, mstrInNO, strSQL) Then Exit Function
    strSQL = "zl_������ʼ�¼_Verify('" & mstrInNO & "','" & UserInfo.��� & "','" & UserInfo.���� & "','" & strSQL & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    '������Ϣ
    Call SendMsgModule
    SaveVerify = True
    If gbln��˴�ӡ Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1122", Me, "NO=" & mstrInNO, "ҩƷ��λ=" & IIf(gblnҩ����λ, 1, 0), 2)
    End If
    
    '110319
    If mblnDrugMachine Then
        '�����ʽ��1|����1,������1;����2,������2
        Dim strData As String, strReturn As String
        strData = "1|" & "9," & mstrInNO
        Call mobjDrugMachine.Operation(gstrDBUser, Val("21-��ҩ[�����סԺ������ϸ�ϴ�]"), strData, strReturn)
    End If
    
    mstrInNO = "": cboNO.Text = ""
    Call ClearPatientInfo(True)
    Call ClearTotalInfo
    Call ClearBillRows: Call ClearMoney
    Call NewBill: Call SetMoneyList
    cboNO.Locked = False: cboNO.SetFocus
    mblnSaveData = True
    SaveVerify = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function isValiedCargeFee() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����������ݵĺϷ���
    '����:���ݺϷ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-16 14:05:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, i As Long, j As Long, strTmp As String
    Dim p As Integer, dblNum As Double, strInfo As String
    Dim cur�ϼ� As Currency, cur���ն� As Currency, cur��� As Currency
    Dim blnMerge As Boolean, k As Integer, bln����� As Boolean, colStock As Collection
    Dim dblToTal As Double, lngҩ��ID As Long
    Dim blnExistValidItem As Boolean
    
    On Error GoTo errHandle
    If mbytInFun = 2 Then
        If mrsInfo.State = adStateClosed Then
            MsgBox "û�з���" & gstrCustomerAppellation & "��Ϣ,��ȷ��" & gstrCustomerAppellation & "��Ϣ��", vbInformation, gstrSysName
            txtPatient.SetFocus: Exit Function
        End If
    ElseIf mbytInFun = 0 Then '�շѺͼ��ʱ�����������
        If txtPatient.Text = "" Then
            MsgBox "û�з���" & gstrCustomerAppellation & "��Ϣ,������" & gstrCustomerAppellation & "��Ϣ��", vbInformation, gstrSysName
            txtPatient.SetFocus: Exit Function
        ElseIf mobjBill.���� = "" Then
            mobjBill.���� = txtPatient.Text
        End If
    End If
    If mbytInFun = 1 And gint������Դ = 2 And Trim(txtPatient.Text) = "" Then
         MsgBox "������ԴΪסԺ����ʱ���������벡����Ϣ��", vbInformation, gstrSysName
        txtPatient.SetFocus: Exit Function
    End If
    If CheckTextLength("����", txtPatient) = False Then Exit Function
    If CheckTextLength("����", txt����) = False Then Exit Function
    If Not CheckOldData(txt����, cbo���䵥λ) Then Exit Function
    '���˺� ����:?? ����:2010-01-07 11:26:40
    '����ҽ�����
    If mobjBill.����ID <> 0 And mbytInFun = 1 And mbytInState = 0 Then
        gstrSQL = "Select ���� from ������Ϣ where ����id=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mobjBill.����ID)
        If Not rsTmp.EOF Then
            MCPAR.ҽ��ȷ���������� = gclsInsure.GetCapability(supportҽ��ȷ����������, mobjBill.����ID, Val(Nvl(rsTmp!����)))
            If zlCheck����ҽ��(Val(Nvl(rsTmp!����))) = False Then Exit Function
        End If
    End If
    If mobjBill.�ѱ� = "" Then
        MsgBox "��ѡ��" & gstrCustomerAppellation & "�ѱ�", vbInformation, gstrSysName
        If cbo�ѱ�.Visible And cbo�ѱ�.Enabled Then cbo�ѱ�.SetFocus
        Exit Function
    End If

    If CheckBillsEmpty Then
        MsgBox "������û���κ�����,����ȷ���뵥�����ݣ�", vbInformation, gstrSysName
        Bill.SetFocus: Exit Function
    ElseIf mobjBill.Pages.Count > 1 Then
        For i = 1 To mobjBill.Pages.Count
            If CheckBillsEmpty(i) Then
                MsgBox "�� " & i & " �ŵ���û�������κ����ݣ�", vbInformation, gstrSysName
                tbsBill.Tabs(i).Selected = True
                Bill.SetFocus: Exit Function
            End If
        Next
    End If
    '�Ƿ�ȫ��������ִ�п���
    i = CheckExecuteDept(j)
    If i > 0 And j > 0 Then
        If mobjBill.Pages.Count > 1 Then
            MsgBox "�� " & j & " �ŵ����е� " & i & " ����Ŀû��ָ��ִ�п��ң�", vbInformation, gstrSysName
            tbsBill.Tabs(j).Selected = True
        Else
            MsgBox "�����е� " & i & " ����Ŀû��ָ��ִ�п��ң�", vbInformation, gstrSysName
        End If
        Bill.SetFocus: Exit Function
    End If
    If Not glngSys Like "8??" Then
        For i = 1 To mobjBill.Pages.Count
            If mobjBill.Pages(i).��������ID = 0 Then
                If mobjBill.Pages.Count > 1 Then
                    MsgBox "�� " & i & " �ŵ���û��ָ���������ң�", vbInformation, gstrSysName
                    tbsBill.Tabs(i).Selected = True
                Else
                    MsgBox "û��ָ���������ң�", vbInformation, gstrSysName
                End If
                If gbyt����ҽ�� = 0 Then
                    cbo������.SetFocus
                Else
                    cbo��������.SetFocus
                End If
                Exit Function
            End If
        Next
    End If
    
    '������
    If mbytInFun <> 2 And gbln�����俪���� Then
        For i = 1 To mobjBill.Pages.Count
            If mobjBill.Pages(i).������ = "" Then
                If mobjBill.Pages.Count > 1 Then
                    MsgBox "�� " & i & " �ŵ���û��ָ�������ˣ�", vbInformation, gstrSysName
                    tbsBill.Tabs(i).Selected = True
                Else
                    MsgBox "û��ָ�������ˣ�", vbInformation, gstrSysName
                End If
                cbo������.SetFocus: Exit Function
            End If
        Next
    End If
    '��鿪�����뿪�����Ҷ�Ӧ��ϵ
    If mbytInFun <> 2 And mbytInState = 0 And (gbyt����ҽ�� = 0 Or gbyt����ҽ�� = 1) Then
        If Not (cbo������.Locked And cbo��������.Locked) Then
            For i = 1 To mobjBill.Pages.Count
                If mobjBill.Pages(i).������ <> "" And mobjBill.Pages(i).NO = "" Then        '25618:And mobjBill.Pages(i).NO = "":���˺����,��Ҫ�ǹҺŲ����Ļ��۵�ʱ,��һ����������ʱ�ٴ���,��˲��ܼ��
                    mrs������.Filter = "����='" & mobjBill.Pages(i).������ & "' And ����ID=" & mobjBill.Pages(i).��������ID
                    If mrs������.RecordCount = 0 Then
                        MsgBox "������""" & mobjBill.Pages(i).������ & """�����ڿ�������""" & zlStr.NeedName(cbo��������.Text) & """,���飡", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
            Next
        End If
    End If
    '��鲡�˹Һſ���,�Ƿ������û�йҺŵĲ����շ�
    If mbytInFun = 0 And gblnCheckRegeventDept And gint������Դ = 1 _
        And (gTy_System_Para.Sy_Reg.bytNODaysGeneral > 0 Or gTy_System_Para.Sy_Reg.bytNoDayseMergency > 0) And mobjBill.����ID > 0 Then
        Set rsTmp = GetDeptByRegevent(mobjBill.����ID)
        If rsTmp.RecordCount > 0 Then 'δ�Һŵĸ��ݱ��ز����Ѽ��
            For i = 1 To mobjBill.Pages.Count
                If Not CheckDeptIsMedTech(mobjBill.Pages(i).��������ID) Then
                    rsTmp.Filter = "ִ�в���ID=" & mobjBill.Pages(i).��������ID
                    If rsTmp.RecordCount = 0 Then
                        MsgBox "��ǰ����û���ڵ�" & i & "�ŵ��ݵĿ������ҹҹ���,�������շ�!", vbInformation, gstrSysName
                        tbsBill.Tabs(i).Selected = True
                        Exit Function
                    End If
                End If
            Next
        End If
    End If
    '��ʿ���:�жϷǷ�����
    For i = 1 To mobjBill.Pages.Count
        If CheckInhibitiveByNurse(i) Then
            If mobjBill.Pages.Count > 1 Then
                MsgBox "��ʿֻ���������Ƽ�������Ŀ,���� " & i & " �ŵ����д����������͵���Ŀ��", vbInformation, gstrSysName
                If tbsBill.SelectedItem.Index <> i Then
                    tbsBill.Tabs(i).Selected = True
                End If
            Else
                MsgBox "��ʿֻ���������Ƽ�������Ŀ,�������д����������͵���Ŀ��", vbInformation, gstrSysName
            End If
            Bill.SetFocus: Exit Function
        End If
    Next
 

    If Not IsDate(txtDate.Text) Then
        MsgBox "��������ȷ�ķ������ڣ�", vbInformation, gstrSysName
        txtDate.SetFocus: Exit Function
    End If
    
    If mbln���� Then
        If txtDate.Text > mstr���ת��ʱ�� And mstr���ת��ʱ�� <> "" Then
            MsgBox "�ò��˲�¼�ķ���ʱ�䳬�������ת����ʱ��(" & mstr���ת��ʱ�� & ")�����ܽ��в��Ѳ�����", vbInformation, gstrSysName
            txtDate.SetFocus: Exit Function
        End If
        If cbo��������.ItemData(cbo��������.ListIndex) <> mlngDeptID And mlngDeptID <> 0 Then
            MsgBox "�������Ҳ��ǲ���ת�ƵĿ��ң����ܽ��в��Ѳ�����", vbInformation, gstrSysName
            Exit Function
        End If
    End If
 
    '���뻮�۵��շ�ʱ,�����ҽ�����ɵ�,����������
    For i = 1 To mobjBill.Pages.Count
        '���ÿ�ŵ����ж�(��Ϊ���ܻ��ۺ��շѻ���),�Ƿ��ǵ���ҽ�����ɵĻ��۵��շ�
        If mobjBill.Pages(i).NO <> "" And mobjBill.Pages(i).ҽ����� <> 0 Then
            If mobjBill.Pages(i).ʵ�ս�� <> GetBillSumByDB(mobjBill.Pages(i).NO) Then
                MsgBox "����[" & mobjBill.Pages(i).NO & "]�Ĳ����շѼ�¼�ѱ������޸Ļ�����,�����¶�ȡ���ݺ����շѣ�", vbInformation, gstrSysName
                tbsBill.Tabs(i).Selected = True
                Exit Function
            End If
        End If
    Next
    
    For p = 1 To mobjBill.Pages.Count
        For i = 1 To mobjBill.Pages(p).Details.Count
            With mobjBill.Pages(p).Details(i)
                If CheckServeRange(0, .�շ�ϸĿID) = False Then Exit Function
            End With
        Next i
    Next p

    '��������Ч��Ŀ����
    strTmp = ""
    For p = 1 To mobjBill.Pages.Count
        blnExistValidItem = False
        For i = 1 To mobjBill.Pages(p).Details.Count
            '27467,106490
            If mobjBill.Pages(p).Details(i).���� <> 0 Then blnExistValidItem = True
            If mobjBill.Pages(p).Details(i).�շ�ϸĿID = 0 Then
                If mobjBill.Pages.Count > 1 Then
                    MsgBox "�� " & p & " �ŵ����е� " & i & " ��û����ȷ��������,��������ɾ�����У�", vbInformation, gstrSysName
                    tbsBill.Tabs(p).Selected = True
                Else
                    MsgBox "�����е� " & i & " ��û����ȷ��������,��������ɾ�����У�", vbInformation, gstrSysName
                End If
                Bill.SetFocus: Exit Function
            ElseIf InStr(1, ",5,6,7,", mobjBill.Pages(p).Details(i).�շ����) > 0 Then
                '�ռ�ҩƷ�ķ�ҩҩ����Ӧ�ķ������
                strTmp = strTmp & "," & mobjBill.Pages(p).Details(i).�շ�ϸĿID
            End If
            
            If mbytInFun = 2 Then
                '���������������:ֻ���������۲��˲Ż���м���Ƿ����
                '����:36558
                 If Not mrsInfo Is Nothing Then
                    If mrsInfo.State = 1 Then
                        If Nvl(mrsInfo!����, 0) = 1 Then
                            If InStr(",5,6,7,", mobjBill.Pages(p).Details(i).�շ����) > 0 And gblnҩ����λ Then
                                 dblNum = mobjBill.Pages(p).Details(i).���� * mobjBill.Pages(p).Details(i).���� * mobjBill.Pages(p).Details(i).Detail.ҩ����װ
                             Else
                                 dblNum = mobjBill.Pages(p).Details(i).���� * mobjBill.Pages(p).Details(i).����
                             End If
                             If dblNum < 0 Then
                                If Not CheckNegative(mobjBill.����ID, mobjBill.��ҳID, mobjBill.Pages(p).Details(i).�շ�ϸĿID, mobjBill.Pages(p).Details(i).ִ�в���ID, dblNum, mobjBill.Pages(p).Details(i).Detail.ҩ����װ, mstrPrivs, Format(mrsInfo!��Ժ����, "yyyy-mm-dd HH:MM:SS")) Then
                                    tbsBill.Tabs(p).Selected = True
                                    Bill.SetFocus: Exit Function
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next
            
        '����:41668,106490
        If mobjBill.Pages(p).NO = "" And mbytInState = 0 And blnExistValidItem = False Then
            If mobjBill.Pages.Count > 1 Then
                MsgBox "�� " & p & " �ŵ���������Ҫ��һ�����β�Ϊ�����Ŀ�����飡", vbInformation, gstrSysName
                tbsBill.Tabs(p).Selected = True
            Else
                MsgBox "����������Ҫ��һ�����β�Ϊ�����Ŀ�����飡", vbInformation, gstrSysName
            End If
            Bill.SetFocus: Exit Function
        End If
    Next
            
    '���ҩƷ�ķ�ҩҩ����Ӧ�ķ������
    If strTmp <> "" Then
        strTmp = Mid(strTmp, 2)
        Set rsTmp = GetServiceDept(strTmp)
        If Not rsTmp Is Nothing Then
            For p = 1 To mobjBill.Pages.Count
                strTmp = ""
                For i = 1 To mobjBill.Pages(p).Details.Count
                    If InStr(1, ",5,6,7,", mobjBill.Pages(p).Details(i).�շ����) > 0 Then
                        strInfo = mobjBill.Pages(p).Details(i).�շ�ϸĿID
                        '�ȼ���Ƿ�������Ĵ洢�ⷿ
                        rsTmp.Filter = "�շ�ϸĿID=" & strInfo & " And ִ�п���id=" & mobjBill.Pages(p).Details(i).ִ�в���ID
                        If rsTmp.RecordCount = 0 Then
                            strTmp = strTmp & "," & i
                        Else
                            '�ټ���Ƿ�������ķ������(û�����÷�����ҵ�,��������IDΪ��)
                            rsTmp.Filter = "(" & rsTmp.Filter & " And ��������ID=" & _
                                IIf(mobjBill.����ID = 0, mobjBill.Pages(p).��������ID, mobjBill.����ID) & ") Or (" & rsTmp.Filter & " And ��������ID=0)"
                            If rsTmp.RecordCount = 0 Then
                                strTmp = strTmp & "," & i
                            End If
                        End If
                    End If
                Next
                If strTmp <> "" Then
                    strTmp = Mid(strTmp, 2)
                    MsgBox "����,��" & p & "�ŵ���,��" & strTmp & "��ҩƷ�Ƿ�Υ�����¹���:" & vbCrLf & vbCrLf & _
                        "A.ѡ���ִ�п��Ҳ���ҩƷ�Ĵ洢�ⷿ" & vbCrLf & _
                        "B.���˿���[" & GET��������(IIf(mobjBill.����ID = 0, mobjBill.Pages(p).��������ID, mobjBill.����ID), mrs��������) & "]������ҩƷ�ڴ˴洢�ⷿ�ķ������.", _
                        vbInformation, gstrSysName
                    Exit Function
                End If
            Next
        End If
    End If

    '����ְ����
    '1.���ѻ�ҽ������
    If cboҽ�Ƹ���.ListIndex <> -1 And mbln����ְ���� Then
        'ҽ���򹫷Ѳ���
        '����:45605
        If zlIsCheckMedicinePayMode(zlStr.NeedName(cboҽ�Ƹ���)) Then
            i = CheckDuty(, False, j)
            If i > 0 And j > 0 Then
                If mobjBill.Pages.Count > 1 Then tbsBill.Tabs(j).Selected = True
                Bill.Row = i: Bill.MsfObj.TopRow = i
                Bill.Col = BillCol.��Ŀ: Bill.SetFocus: Exit Function
            End If
        End If
    End If
    '2.���в�����Ŀ
    If mbln����ְ���� Then
        i = CheckDuty(, True, j)
        If i > 0 And j > 0 Then
            If mobjBill.Pages.Count > 1 Then tbsBill.Tabs(j).Selected = True
            Bill.Row = i: Bill.MsfObj.TopRow = i
            Bill.Col = BillCol.��Ŀ: Bill.SetFocus: Exit Function
        End If
    End If
    
    '�������ͼ��
    If Not CheckFeeType Then Exit Function
    
    '���ʷ��౨��
    If mbytInFun = 2 And Not mrsWarn Is Nothing Then
        '���ݷ���
        cur�ϼ� = GetBillSum
        If cur�ϼ� > 0 Then
           Call LoadFeeInfor(mrsInfo!����ID)
           
            '���¶�ȡ���ն�
            cur���ն� = GetPatiDayMoney(mrsInfo!����ID)
                           
            cur��� = Val(cmdPrint.Tag)
            If gbln�����������۷��� Then cur��� = cur��� - GetPriceMoneyTotal(0, mrsInfo!����ID) + IIf(mbytBilling = 1, Original.ʵ�պϼ�, 0)
                    
            For i = 1 To mobjBill.Pages.Count
                For j = 1 To mobjBill.Pages(i).Details.Count
                    gbytWarn = BillingWarn(mstrPrivs, mrsInfo!����, mrsInfo!���ò���, mrsWarn, cur���, cur���ն� - Original.ʵ�պϼ�, cur�ϼ�, _
                        IIf(IsNull(mrsInfo!������), 0, mrsInfo!������), mobjBill.Pages(i).Details(j).�շ����, _
                        mobjBill.Pages(i).Details(j).Detail.�������, mstrWarn)
                    If gbytWarn = 2 Or gbytWarn = 3 Then Exit Function
                Next
            Next
        End If
    End If
    'ҩƷ���ɼ��
    strInfo = CheckDisable(mobjBill)
    If strInfo <> "" Then
        If strInfo Like "*(�������)*" Then
            MsgBox strInfo, vbInformation, gstrSysName
            Exit Function
        Else
            If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        End If
    End If
                
    '�����������
    If mbln����������� Then
        If Not gbln�������� Then
            If Not CheckLimit(mobjBill) Then Exit Function
        End If
    End If
    
    '���ŵ�����߶�
    If gcurMax <> 0 And (mbytInFun = 0 Or mbytInFun = 1) Then
        For i = 1 To mobjBill.Pages.Count
            If GetBillSum(, i) > gcurMax Then
                If mobjBill.Pages.Count > 1 Then
                    MsgBox "�� " & i & " �ŵ��ݽ���������ƽ��:" & Format(gcurMax, "0.00") & " ,�������棡", vbInformation, gstrSysName: Exit Function
                Else
                    MsgBox "���ݽ���������ƽ��:" & Format(gcurMax, "0.00") & " ,�������棡", vbInformation, gstrSysName: Exit Function
                End If
            End If
        Next
    End If
    '��������ʱ��ҩƷͬһҩ���Ƿ����ظ�����
    blnMerge = False
    For p = 1 To mobjBill.Pages.Count
        For i = 1 To mobjBill.Pages(p).Details.Count
            With mobjBill.Pages(p).Details(i)
                If (.Detail.���� Or .Detail.���) And (InStr(",5,6,7,", .�շ����) > 0 Or .�շ���� = "4" And .Detail.��������) Then
                    For k = 1 To mobjBill.Pages.Count
                        For j = 1 To mobjBill.Pages(k).Details.Count
                            If Not (p = k And i = j) And .�շ�ϸĿID = mobjBill.Pages(k).Details(j).�շ�ϸĿID And .ִ�в���ID = mobjBill.Pages(k).Details(j).ִ�в���ID Then
                                '���ŵ��ݵ����
                                If mobjBill.Pages.Count > 1 Then
                                    '��ʱ�۵ķ���ҩƷ���ڲ�ͬ�ĵ���������ͬ�ģ������ϲ���������
                                    If .Detail.��� Or (Not .Detail.��� And .Detail.���� And p = k) Then
                                        If .�շ���� = "4" Then
                                            If Not blnMerge Then
                                                If .Detail.���� = mobjBill.Pages(k).Details(j).Detail.���� Then
                                                    If MsgBox("�� " & p & " �ŵ��ݵ� " & i & " ��,���� " & k & " �ŵ��ݵ� " & j & " �е�" & _
                                                        vbCrLf & "������ʱ����������""" & .Detail.���� & """��ͬһ�����ϲ��ű��ظ����롣" & _
                                                        vbCrLf & vbCrLf & "Ҫ�Զ��ϲ������������ظ�����ķ�����ʱ����Ŀ��", _
                                                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                                                        blnMerge = True     '��Ӧ�˳�ѭ������Ϊ��Ҫ����Ƿ��в�ͬ�������в�ҩ,����еĻ��������Զ��ϲ�
                                                    Else
                                                        tbsBill.Tabs(k).Selected = True: Exit Function
                                                    End If
                                                End If
                                            End If
                                        Else
                                            '������ͬ�����е���ҩ������ͬʱ��Ӧ�ǲ�ͬ���䷽���޷��Զ��ϲ�
                                            If .�շ���� = "7" And .���� <> mobjBill.Pages(k).Details(j).���� Then
                                                MsgBox "�� " & p & " �ŵ��ݵ� " & i & " ��,���� " & k & " �ŵ��ݵ� " & j & " �е�" & _
                                                    vbCrLf & "������ʱ���в�ҩ""" & .Detail.���� & """(��ͬ����)��ͬһ��ҩ�����ظ����롣", vbInformation, gstrSysName
                                                tbsBill.Tabs(k).Selected = True: Exit Function
                                            ElseIf Not blnMerge Then
                                                If MsgBox("�� " & p & " �ŵ��ݵ� " & i & " ��,���� " & k & " �ŵ��ݵ� " & j & " �е�" & _
                                                    vbCrLf & "������ʱ��ҩƷ""" & .Detail.���� & """��ͬһ��ҩ�����ظ����롣" & _
                                                    vbCrLf & vbCrLf & "Ҫ�Զ��ϲ������������ظ�����ķ�����ʱ����Ŀ��", _
                                                    vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                                                    blnMerge = True
                                                Else
                                                    tbsBill.Tabs(k).Selected = True: Exit Function
                                                End If
                                            End If
                                        End If
                                    End If
                                '���ŵ��ݵ����
                                ElseIf Not blnMerge Then
                                    If .�շ���� = "4" Then
                                        If .Detail.���� = mobjBill.Pages(k).Details(j).Detail.���� Then
                                            strInfo = "�� " & j & " �еķ�����ʱ����������""" & .Detail.���� & """��ͬһ�����ϲ��ű��ظ����롣" & _
                                                        vbCrLf & vbCrLf & "Ҫ�Զ��ϲ������������ظ�����ķ�����ʱ����Ŀ��"
                                        End If
                                    Else
                                        strInfo = "�� " & j & " �еķ�����ʱ��ҩƷ""" & .Detail.���� & """��ͬһ��ҩ�����ظ����롣" & _
                                                    vbCrLf & vbCrLf & "Ҫ�Զ��ϲ������������ظ�����ķ�����ʱ����Ŀ��"
                                    End If
                                    
                                    If strInfo <> "" Then
                                        If MsgBox(strInfo, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                                            blnMerge = True     '�����˳�ѭ��
                                        Else
                                            Exit Function
                                        End If
                                    End If
                                End If
                            End If
                        Next
                    Next
                End If
            End With
        Next
    Next
    '�Զ��ϲ�,ֻҪ�кϲ�,��Ӧ���±���,��Ϊ������շѲ���ӡ,�����ı仯,����Ӱ�칤���ѵ�����
    If blnMerge Then
        Call MergeRepeatItem
        MsgBox "�Զ��ϲ�����ɣ��ϲ�����ý��������ѷ����仯������󱣴档", vbInformation, gstrSysName
        Exit Function
    End If
   'ҩƷ�����(�������ֹʱ�����ʱ��ҩƷ)
    bln����� = (InStr(mstrPrivs, "�������") = 0)    '�Ƿ���Ȩ�޲������(������ʱ�۱�����)
    For p = 1 To mobjBill.Pages.Count
        For i = 1 To mobjBill.Pages(p).Details.Count
            With mobjBill.Pages(p).Details(i)
                Set colStock = IIf(.�շ���� = "4", mcolStock2, mcolStock1)
            
                If InStr(",5,6,7,", .�շ����) > 0 Then
                    If .Detail.���� Or .Detail.��� Then
                        dblToTal = GetDrugTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                        .Detail.��� = GetStock(.�շ�ϸĿID, .ִ�в���ID)
                        If gblnҩ����λ Then .Detail.��� = .Detail.��� / .Detail.ҩ����װ
                        
                        If mbytInState = 0 And mstrInNO <> "" Then .Detail.��� = .Detail.��� + GetOriginalTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                        If dblToTal > .Detail.��� Then
                            If mobjBill.Pages.Count > 1 Then strTmp = "�� " & p & " �ŵ���"
                            MsgBox strTmp & "�� " & i & " ��ʱ�ۻ����ҩƷ""" & .Detail.���� & _
                                """�ĵ�ǰ���" & IIf(InStr(1, mstrPrivs, "��ʾ���") > 0, .Detail.���, "") & "������������""" & _
                                dblToTal & """��", vbInformation, gstrSysName
                            tbsBill.Tabs(p).Selected = True
                            Bill.SetFocus: Exit Function
                        End If
                    Else
                        If colStock("_" & .ִ�в���ID) = 2 And bln����� Then
                            dblToTal = GetDrugTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                            .Detail.��� = GetStock(.�շ�ϸĿID, .ִ�в���ID)
                            If gblnҩ����λ Then .Detail.��� = .Detail.��� / .Detail.ҩ����װ
                            
                            If mbytInState = 0 And mstrInNO <> "" Then .Detail.��� = .Detail.��� + GetOriginalTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                            If dblToTal > .Detail.��� Then
                                If mobjBill.Pages.Count > 1 Then strTmp = "�� " & p & " �ŵ���"
                                MsgBox strTmp & "�� " & i & " ��ҩƷ""" & .Detail.���� & _
                                    """�ĵ�ǰ���" & IIf(InStr(1, mstrPrivs, "��ʾ���") > 0, .Detail.���, "") & "������������""" & _
                                    dblToTal & """,���޸Ļ����Ƿ��ж������롣", vbInformation, gstrSysName
                                tbsBill.Tabs(p).Selected = True
                                Bill.SetFocus: Exit Function
                            End If
                        End If
                    End If
                ElseIf .�շ���� = "4" And .Detail.�������� Then
                    If .Detail.���� Or .Detail.��� Then
                        dblToTal = GetDrugTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                        .Detail.��� = GetStock(.�շ�ϸĿID, .ִ�в���ID, .Detail.����)
                        
                        If mbytInState = 0 And mstrInNO <> "" Then .Detail.��� = .Detail.��� + GetOriginalTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                        If dblToTal > .Detail.��� Then
                            If mobjBill.Pages.Count > 1 Then strTmp = "�� " & p & " �ŵ���"
                            MsgBox strTmp & "�� " & i & " ��ʱ�ۻ������������""" & .Detail.���� & _
                                """�ĵ�ǰ���" & IIf(InStr(1, mstrPrivs, "��ʾ���") > 0, .Detail.���, "") & "������������""" & dblToTal & """��", vbInformation, gstrSysName
                            tbsBill.Tabs(p).Selected = True
                            Bill.SetFocus: Exit Function
                        End If
                    Else
                        If colStock("_" & .ִ�в���ID) = 2 And bln����� Then
                            dblToTal = GetDrugTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                            .Detail.��� = GetStock(.�շ�ϸĿID, .ִ�в���ID, .Detail.����)
                            
                            If mbytInState = 0 And mstrInNO <> "" Then .Detail.��� = .Detail.��� + GetOriginalTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                            If dblToTal > .Detail.��� Then
                                If mobjBill.Pages.Count > 1 Then strTmp = "�� " & p & " �ŵ���"
                                MsgBox strTmp & "�� " & i & " ����������""" & .Detail.���� & _
                                    """�ĵ�ǰ���" & IIf(InStr(1, mstrPrivs, "��ʾ���") > 0, .Detail.���, "") & "������������""" & dblToTal & """,���޸Ļ����Ƿ��ж������롣", vbInformation, gstrSysName
                                tbsBill.Tabs(p).Selected = True
                                Bill.SetFocus: Exit Function
                            End If
                        End If
                    End If
                End If
            End With
        Next
    Next
    '����������ϵ����Ч��
    For i = 1 To mobjBill.Pages.Count
        For j = 1 To mobjBill.Pages(i).Details.Count
            With mobjBill.Pages(i).Details(j)
                If .�շ���� = "4" And .Detail.�������� Then
                    dblToTal = GetDrugTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID, i)
                    If Not CheckValidity(.�շ�ϸĿID, .ִ�в���ID, dblToTal) Then Exit Function
                End If
            End With
        Next
    Next
    '��ҩ���ڼ��(�����۵�)
    For i = 1 To mobjBill.Pages.Count
        If mobjBill.Pages(i).NO <> "" And tbsBill.Tabs(i).Tag = "" Then
            lngҩ��ID = BillExistDrug(mobjBill.Pages(i).NO, 1)
            If lngҩ��ID <> 0 Then
                If ExistWindow(lngҩ��ID, mrs��ҩ����) Then
                    MsgBox "�޷�����" & GET��������(lngҩ��ID, mrsUnit) & "�ķ�ҩ���ڣ���ȷ���Ƿ��������Ŵ����ϰࡣ", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
    Next
    
    '���ۼ��,105872
    If Not gobjPublicDrug Is Nothing Then
        'Private Function zlCheckPriceAdjustBySell(ByVal lngҩƷid As Long, ByVal lngҩ��id As Long) As Boolean
        '���۹���ģʽʱ���жϼ۸��Ƿ��������۹���Ҫ���ɱ��ۺ��ۼ�һ�£�
        '����ҩƷ���ۼ��ǹ̶��ģ��Ƚ�����ҩ���ĳɱ��ۣ�������ڲ�һ�µľͲ������۳���
        'ʱ��ҩƷ���Ƚ�ҩ������¼�����ۼۺͳɱ��ۣ�������ڲ�һ�µľͲ������۳���
        '���۳���ʱֻ�ж�ҩ��
        '���أ�True-�����������۳��⣻false-���ܽ������۳���
        For p = 1 To mobjBill.Pages.Count
            For i = 1 To mobjBill.Pages(p).Details.Count
                With mobjBill.Pages(p).Details(i)
                    If InStr(",5,6,7,", .�շ����) > 0 Then
                        If gobjPublicDrug.zlCheckPriceAdjustBySell(.�շ�ϸĿID, .ִ�в���ID) = False Then
                            tbsBill.Tabs(p).Selected = True
                            Bill.SetFocus: Exit Function
                        End If
                    End If
                End With
            Next
        Next
    End If
    
    If mstrInNO <> "" Then
        If HaveExecute(1, mstrInNO, IIf(mbytInFun = 2, 2, 1)) Then
            MsgBox "�õ��ݰ�����ȫִ�л򲿷�ִ�е���Ŀ,�������޸ġ�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    '���˺�:����Ƿ�ֻ�и�������,���ֻ�и�������,ֱ���˳�:
    '22441
    If CheckMainOperation = False Then Exit Function
    
    If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModul, 0, 1, _
        MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), mbytInFun, IIf(mbytInFun = 1, 1, IIf(mbytBilling = 0, 0, 1)))) = False Then
        Exit Function
    End If
    
    isValiedCargeFee = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function CheckServeRange(intType As Integer, lng�շ�ϸĿID As Long, Optional intRow As Integer = 0) As Boolean
'����:����շ���Ŀ�ķ������,intType:0-�������;1-סԺ����
    Dim strSQL As String, rsTmp As ADODB.Recordset
    strSQL = "Select ����,Nvl(�������,0) As ������� From �շ���ĿĿ¼ Where ID = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "CheckServeRange", lng�շ�ϸĿID)
    If rsTmp.EOF Then
        MsgBox "����ȷ��" & IIf(intRow = 0, "", "��" & intRow & "��") & "�շ���Ŀ�ķ������,������Ŀ�Ƿ���ȷ¼��!"
        Exit Function
    Else
        Select Case intType
        Case 0
            If Val(rsTmp!�������) = 2 Or Val(rsTmp!�������) = 0 Then
                MsgBox IIf(intRow = 0, "", "��" & intRow & "��") & "�շ���Ŀ[" & rsTmp!���� & "]������������,����!"
                Exit Function
            End If
        Case 1
            If Val(rsTmp!�������) = 1 Or Val(rsTmp!�������) = 0 Then
                MsgBox IIf(intRow = 0, "", "��" & intRow & "��") & "�շ���Ŀ[" & rsTmp!���� & "]��������סԺ,����!"
                Exit Function
            End If
        Case Else
            If Val(rsTmp!�������) = 0 Then
                MsgBox IIf(intRow = 0, "", "��" & intRow & "��") & "�շ���Ŀ[" & rsTmp!���� & "]�������ڲ���,����!"
                Exit Function
            End If
        End Select
    End If
    CheckServeRange = True
End Function

Private Function CheckBillNOAndBookeFee(Optional blnReCharge As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:Ʊ�ݺ�����,�����Ѵ�ӡ���
    '���:blnReCharge-�Ƿ������շѵļ��
    '����:���ݺϷ�,����tru,���򷵻�false
    '����:���˺�
    '����:2011-08-16 14:25:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl��� As Double, j As Long, p As Long, i As Long
    
    On Error GoTo errHandle
    'Ʊ�ݺ�����,�����Ѵ�ӡ���
    If Not blnReCharge Then
        If Not (mbytInFun = 0 And Not mblnSaveAsPrice) Then CheckBillNOAndBookeFee = True: Exit Function
    End If
    mblnPrint = True
    '����Ƿ��ӡƱ��
    If mintInvoicePrint = 0 Then
        mblnPrint = False
    Else
        If mintInvoicePrint = 2 Then
            If MsgBox("�Ƿ��ӡƱ��?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                mblnPrint = False
            End If
        End If
    End If
    If Not blnReCharge Then
        '��������(ֻ�й�����)�Ƿ��ӡ,���۲�����������,�����е�ĳһ��ֻ�й�����ʱ���ڴ�ӡ����ʱ�жϲ���ӡ
        If mblnPrint And gTy_Module_Para.bln������ Then
            If GetBillSum = Calc������ Then
                If MsgBox("��ǰ����ʵ��û����ȡ����,Ҫ��ӡ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    mblnPrint = False
                End If
            End If
        End If
    End If
    If Not mblnPrint Then
        If blnReCharge Then
                CheckBillNOAndBookeFee = True: Exit Function
        End If
        If gTy_Module_Para.bln������ Then
            j = 0
            For p = 1 To mobjBill.Pages.Count
                For i = 1 To mobjBill.Pages(p).Details.Count
                    If mobjBill.Pages(p).Details(i).������ Then
                        If j = 0 Then MsgBox "��Ϊ����ӡƱ��,ϵͳ���Զ�ɾ�������ѣ�", vbInformation, gstrSysName
                        j = j + 1
                        Call DeleteDetail(i, p)
                        Call ShowDetails
                        Call ShowMoney(p)
                        Bill.TxtVisible = False: Bill.CmdVisible = False: Bill.CboVisible = False
                        Exit For
                    End If
                Next
            Next
        End If
    Else
        If gblnStrictCtrl Then
            If Trim(txtInvoice.Text) = "" Then
                MsgBox "��������һ����Ч��Ʊ�ݺ��룡", vbInformation, gstrSysName
                txtInvoice.SetFocus: Exit Function
            End If

InvoiceHandle:
            If zlGetInvoiceGroupUseID(mlng����ID, IIf(gTy_Module_Para.bln�ֱ��ӡ And mbytBillSource <> 4, mobjBill.Pages.Count, 1), txtInvoice.Text) = False Then
                Exit Function
            End If
            '�����������,Ʊ���Ƿ�����
            If CheckBillRepeat(mlng����ID, 1, txtInvoice.Text) Then
                'Tag�����⣺24363:���˺飺��Ҫ�ǽ���Զ����ɵĺ��Ƿ��û����ģ���Ҫ�����
                If txtInvoice.Locked = False And txtInvoice.Tag <> Trim(txtInvoice.Text) Then
                    MsgBox "Ʊ�ݺ�""" & txtInvoice.Text & """�Ѿ���ʹ�ã����������롣", vbInformation, gstrSysName
                    txtInvoice.SetFocus: Exit Function
                Else
                    Call RefreshFact
                    If txtInvoice.Text = "" Then
                        txtInvoice.SetFocus: Exit Function
                    Else
                        MsgBox "��ǰƱ�ݺ��Ѿ���ʹ�ã������»�ȡƱ�ݺ�:" & txtInvoice.Text, vbInformation, gstrSysName
                        GoTo InvoiceHandle
                    End If
                End If
            End If
        Else
            If Len(txtInvoice.Text) <> gbytFactLength And txtInvoice.Text <> "" Then
                MsgBox "Ʊ�ݺ��볤��Ӧ��Ϊ " & gbytFactLength & " λ��", vbInformation, gstrSysName
                txtInvoice.SetFocus: Exit Function
            End If
        End If
    End If
    If blnReCharge Then
        CheckBillNOAndBookeFee = True: Exit Function
    End If
    '��ϸ������ڻ���
    dbl��� = GetBillSum - GetMedicareSum
    For j = 1 To mobjBill.Pages.Count
        dbl��� = RoundEx(dbl��� + Val(mobjBill.Pages(j).�����) - Val(mobjBill.Pages(j).Ӧ�ɽ��) - Val(mobjBill.Pages(j).��Ԥ����), 7)
    Next
    If dbl��� <> 0 Then
        MsgBox "ʵ�ս��ϼ���֧�����ϼƲ���,��������!" & vbCrLf & vbCrLf & _
            "������ϸʵ�ս��ϼ�+�����-(����֧���ϼ�+Ӧ�ɺϼ�+��Ԥ�����)=" & dbl���, vbInformation, gstrSysName
        Exit Function
    End If
    CheckBillNOAndBookeFee = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function CheckInsure() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ҽ����ؼ��
    '����:�ɹ�,����true,���򷵻�false
    '����:���˺�
    '����:2011-08-16 16:48:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNone As String
    On Error GoTo errHandle
    If mstrYBPati = "" Then CheckInsure = True: Exit Function
    If mbytInFun <> 0 Then CheckInsure = True: Exit Function
    If mintInsure = 61 Then '����ҽ��
        If Not ����Ԥ����(strNone) Then
            If strNone <> "" Then
                MsgBox "��ǰ���ս���ʹ�õĽ��㷽ʽ" & vbCrLf & vbCrLf & vbTab & strNone & vbCrLf & vbCrLf & _
                    "������δ���ã����ȵ����㷽ʽ������������Щ���㷽ʽ��", vbInformation, gstrSysName
            End If
            If cmdԤ����.Visible Then
                cmdԤ����.TabStop = True
                cmdOK.Enabled = False
                cmdԤ����.SetFocus
            End If
            Exit Function
        End If
    End If
    CheckInsure = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function
Private Function SaveChargeFee() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�շѡ����ۡ�����
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-16 10:22:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cur���ѽ�� As Currency, strNos As String, cur�ѽɺϼ� As Currency
    Dim strModiNos As String, dbl�ɿ� As Double, dbl�Ҳ� As Double
    Dim bln����  As Boolean, blnGetFact As Boolean, i As Long, j As Long
    Dim blnTrans As Boolean, blnTransMedicare As Boolean
    Dim bytReturnMode As ExitMode
    Dim str����Nos As String, rsItems As ADODB.Recordset
    
    On Error GoTo errHandle
    If Not (mstrYBPati <> "" And MCPAR.���������շ�) And Not mblnSaveAsPrice Then
        'If txt�ɿ�.Enabled And txt�ɿ�.Visible Then
            Call AutoBultBookFee '�շ�ʱ�Զ�������������Ŀ
        'End If
    End If
    
    If isValiedCargeFee = False Then Exit Function
    If IsDate(txtDate.Text) Then mobjBill.����ʱ�� = CDate(txtDate.Text)
    mobjBill.�Ǽ�ʱ�� = zlDatabase.Currentdate
    If zlGetSaveDataItems_Plugin(mobjBill, str����Nos, rsItems) = False Then Exit Function
    If zlChargeSaveValied_Plugin(glngModul, IIf(mbytInFun = 2, 2, 1), True, _
                                 mbytInFun = 1 Or (mbytInFun = 2 And mbytBilling = 1), str����Nos, rsItems) = False Then Exit Function
    'ֻ�м���ˢ����֤(��Ԥ���ڽ��㴰�ڴ���)
    If (mbytInFun = 2 And mbytBilling = 0 And Val(txt�ϼ�.Text) <> 0) And gdblԤ��������鿨 <> 0 Then
        cur���ѽ�� = Val(txt�ϼ�.Text)
        If Not zlDatabase.PatiIdentify(Me, glngSys, mobjBill.����ID, cur���ѽ��, mlngModul, 1, , IIf(-1 * gdblԤ��������鿨 >= cur���ѽ��, False, True), , , , (gdblԤ��������鿨 = 2)) Then Exit Function
    End If
    
    'Ʊ�ݺż������Ѽ����ܽ����ؼ��
    If CheckBillNOAndBookeFee = False Then Exit Function
    If CheckInsure = False Then Exit Function
        
    On Error GoTo errH
    cmdOK.Enabled = False   '��ֹ���ô�ӡ�������ķ�ģ̬����,�Լ�ҽ��������ʱ
    cmdCancel.Enabled = False: cmdAddBill.Enabled = False: cmdDelBill.Enabled = False
    If cmdԤ����.Visible And cmdԤ����.Enabled Then cmdԤ����.Enabled = False
    Dim blnSaveBill As Boolean '�����Ƿ񱣴�ɹ�
    '���浥��
    '---------------------------------------------------------------------------------------------
    strNos = "": bytReturnMode = 0
    If Not SaveBill(strNos, strModiNos, blnSaveBill, False, bytReturnMode, bln����) Then
        '�շ�,���浥��ʧ�ܺ�Ĵ���
         If blnSaveBill And bytReturnMode <> EM_�������� Then
            If bytReturnMode <> EM_�˳��շ� Then Call ShowBillChargeFee(mlng�������)
         End If
        mlng������� = 0
        cmdOK.Enabled = True: cmdCancel.Enabled = True
        If mintInsure <> 0 Then
            cmdAddBill.Enabled = Not MCPAR.���������շ� And MCPAR.�൥���շ� And InStr(1, mstrPrivs, "ҽ�����˶൥���շ�") > 0
        Else
            cmdAddBill.Enabled = InStr(1, mstrPrivs, "��ͨ���˶൥���շ�") > 0
        End If
        
        If cmdDelBill.Visible And tbsBill.Tabs.Count > 1 Then cmdDelBill.Enabled = True
        If cmdOK.Visible And cmdOK.Enabled Then cmdOK.SetFocus
        If bytReturnMode = EM_�������� Then
                If mblnAutoChangePati And gint������Դ = 2 Then
                    '��Ҫ���ҵ�������Դ1��
                    gint������Դ = 1: zlChangePatiSource (gint������Դ)
                End If
                Call ClearFullBill(False)
        End If
        Exit Function
    End If
    Call zlChargeSaveAfter_Plugin(glngModul, mobjBill.����ID, mobjBill.��ҳID, True, IIf(mbytInFun = 2, 2, 1), strNos)
    mlng������� = 0
    '��ʾLed�����Ϣ
     If mbytInFun = 0 And Not mblnSaveAsPrice Then
        'LED��ʾ:(�ϼ�,)��ҩ����
        If gblnLED And CCur(txt�ϼ�.Text) <> 0 And (mstr���� <> "" Or mstr�д� <> "" Or mstr�ɴ� <> "") Then
            zl9LedVoice.DisplayBank "���úϼ�:" & txt�ϼ�.Text, _
                "ȡҩ����:" & IIf(mstr���� <> "", " " & mstr����, "") & _
                IIf(mstr�ɴ� <> "", " " & mstr�ɴ�, "") & IIf(mstr�д� <> "", " " & mstr�д�, "")
        End If
     End If
     
    Call SendMsgModule
     '��ӡƱ��
    Call PrintBill(strNos, strModiNos)
    
    If mbytInFun = 2 And mbytBilling = 0 Then
        '110319
        If mblnDrugMachine Then
            Dim strData As String, strReturn As String
            If mstrInNO <> "" Then
                '�޸ĵ��ݣ�ɾ����ԭ���ĵ��ݵ�
                Dim rsTemp As ADODB.Recordset, strSQL As String
                '���ﴦ����ҩ��ʽ������ID1,��ҩ����1;����ID2,��ҩ����2;...
                strSQL = "Select Id As ����id, -1 * Nvl(����, 1) * ���� As ��ҩ����" & vbNewLine & _
                        " From ������ü�¼" & vbNewLine & _
                        " Where ��¼���� = 2 And ��¼״̬ = 2 And NO = [1] And �շ���� In ('5', '6', '7')" & vbNewLine & _
                        "       And �Ǽ�ʱ�� + 0 = (Select Max(�Ǽ�ʱ��)" & vbNewLine & _
                        "                       From ������ü�¼" & vbNewLine & _
                        "                       Where ��¼���� = 2 And ��¼״̬ = 2 And NO = [1])"
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ�����˷���Ŀ", mstrInNO)
                Do While Not rsTemp.EOF
                    strData = strData & ";" & Nvl(rsTemp!����id) & "," & Nvl(rsTemp!��ҩ����)
                    rsTemp.MoveNext
                Loop
                If strData <> "" Then
                    strData = Mid(strData, 2)
                    Call mobjDrugMachine.Operation(gstrDBUser, Val("24-������ҩ(����/����)"), strData, strReturn)
                End If
            End If
        
            '�����ʽ��1|����1,������1;����2,������2
            strData = "1|" & "9," & Replace(Replace(strNos, "'", ""), ",", ";9,")
            Call mobjDrugMachine.Operation(gstrDBUser, Val("21-��ҩ[�����סԺ������ϸ�ϴ�]"), strData, strReturn)
        End If
    End If
    
    cmdOK.Enabled = True   '��ֹ���ô�ӡ�������ķ�ģ̬����,�Լ�ҽ����ʱ
    cmdCancel.Enabled = True
    If cmdԤ����.Visible Then cmdԤ����.Enabled = True
    If mbytInFun = 0 And mbytInState = 0 And gbln�ۼ� Then
        txt�ۼ�.Text = Format(GetChargeTotal, "0.00")
    End If
        
    If mstrInNO = "" And Not mblnCopyBill Or (txtModi.Visible And mbytInState = 0 And mstrInNO <> "") Then   '����,����������ͨ�����뵥�ݺ��޸ĵ���
         If fraUpBillShow.Visible Then
            txtPreNO.Text = mobjBill.NO
            '27505
             txtPreMoney.Text = Format(GetBillSum(False, 1), "0.00")
        Else
            sta.Panels(Pan.C2��ʾ��Ϣ) = "��һ�ŵ���:" & mobjBill.NO '�൥��ʱΪ��һ��
        End If
        
        '������޸ģ����˳��޸�״̬
        If txtModi.Visible And mbytInState = 0 And mstrInNO <> "" Then
            txtModi.Text = "": txtModi.Enabled = True
            If fraBill.Visible Then cmdAddBill.Enabled = InStr(1, mstrPrivs, "��ͨ���˶൥���շ�") > 0
        End If
        
        mstrInNO = "":  mlngFirstID = 0: mstrFirstWin = ""
        
        If mbytInFun = 0 Or mbytInFun = 1 Then
            '���������ֹ�����շѣ�
            '1.ҽ������ÿ��ˢ��,�����շѽ���(�������ý��ɿ��������)
            '2.ʹ��Ԥ�������,�����շѽ���(�������ý��ɿ��������)
            '3.һ�ζ��ŵ���,�����շѽ���(�������ý��ɿ��������)
            '3.���ѽɿ�,��ǿ����Ϊ�����շѽ���
            '4.����ʱû�����벡������
            '5.ʹ�ö��ֽ��㷽ʽ����
            '6.�շ�ʱ����Ϊ���۵�
            
            '���˺�:22343:gTy_Module_Para.byt�ɿ����:0-�������нɿ�������ۼƿ���,1-��������ɿ��Ž��������ۼ�(�ı䲡�˳���)��2-�շ�ʱ����Ҫ����ɿ���
            'bln���� = Not ((mstrYBPati <> "" And Not gbln�ɿ����) _
                        Or (Val(txtԤ�����.Text) <> 0 And Not gbln�ɿ����) _
                        Or mobjBill.Pages.Count > 1 And Not gbln�ɿ���� _
                        Or Val(txt�ɿ�.Text) <> 0 _
                        Or mobjBill.���� = "" And mbytInFun = 1 _
                        Or mobjBill.Pages(mintPage).�շѽ��� <> "") '���ֽ��㷽ʽ
            
            '�ɿ����:0-�������нɿ�������ۼƿ���,1-��������ɿ��Ž��������ۼ�
            '       2-�շ�ʱ����Ҫ����ɿ���
    
               '         Or Val(txt�ɿ�.Text) <> 0
            If mbytInFun = 0 Then
            '    bln���� = bln����
            Else
                bln���� = Not ((mstrYBPati <> "" And gTy_Module_Para.byt�ɿ���� <> 1 And gTy_Module_Para.byt�ɿ���� <> 1 And gTy_Module_Para.byt�ɿ���� <> 3) _
                        Or (Val(txtԤ�����.Text) <> 0 And gTy_Module_Para.byt�ɿ���� <> 1 And gTy_Module_Para.byt�ɿ���� <> 3) _
                        Or (mobjBill.Pages.Count > 1 And gTy_Module_Para.byt�ɿ���� <> 1 And gTy_Module_Para.byt�ɿ���� <> 3) _
                        Or (mobjBill.���� = "" And mbytInFun = 1) _
                        Or mobjBill.Pages(mintPage).�շѽ��� <> "") '���ֽ��㷽ʽ
                bln���� = bln���� Or (mstrYBPati <> "" And MCPAR.���������շ�)
             End If
            If Not bln���� Or mblnSaveAsPrice Then
                If gint������Դ = 2 And mblnAutoChangePati Then
                
                    '�Զ��л���,Ҫ������
                    gint������Դ = 1
                    Call zlChangePatiSource(gint������Դ)
                End If
                Call ClearPatientInfo(True)
                Call ClearTotalInfo(True)
                Call InitCommVariable
                blnGetFact = IIf(mblnStartFactUseType, False, True)
            Else
                '��Ȼ����,��ҽ��������������Ա��ٴ���֤
                If mstrYBPati <> "" Then Call ClearPatientInfo(True)
                blnGetFact = True
                mstrPrePati = mobjBill.���� '��¼��ǰ����
                mlngPrePati = mobjBill.����ID
                mstrPreDoctor = zlStr.NeedName(cbo������.Text)
                
                '���˵��ݽ���ۼ�
                mcurBillӦ�� = mcurBillӦ�� + GetBillSum(True)
                mcurBillʵ�� = mcurBillʵ�� + GetBillSum
                mcurBillӦ�� = GetMustPaySum
                
                mintBillNO = mintBillNO + 1
                For i = 1 To mshMoney.Rows - 1
                    If mshMoney.TextMatrix(i, 0) = "" Then Exit For
                Next
                mintMoneyRow = i - 1
                
                Call SaveDrugID(mobjBill.Pages.Count)
            End If
        End If
        
         '���ﻮ��,������ʻ��۱�����������
        If mbytInFun = 1 Or mbytInFun = 2 And mbytBilling = 1 Then
            Bill.Active = False
        Else
            Call ClearBillRows
        End If
        If Not (mbytInFun = 0 Or mbytInFun = 1 Or mbytInFun = 2 And mbytBilling = 1) Then
            Call ClearMoney
        End If
        
        If mbytInFun = 0 And (mstrYBPati <> "" And MCPAR.���������շ�) Then
            Call NewYBBill
            mobjBill.����ID = CLng(Split(mstrYBPati, ";")(8))
            
            '���¶�ȡԤ�����
            If txtԤ�����.Enabled Then Call LoadFeeInfor(mobjBill.����ID)
            
            '���¶�ȡ�������
            mcur������� = gclsInsure.SelfBalance(mobjBill.����ID, CStr(Split(mstrYBPati, ";")(1)), 10, mcur����͸֧, mintInsure)
            sta.Panels(Pan.C3�����ʻ�).Text = "�����ʻ����:" & Format(mcur�������, "0.00")
            sta.Panels(Pan.C3�����ʻ�).Visible = True

            mstrYBPati = ""
        Else
            Call NewBill(blnGetFact, Not Bill.Active And mbytInFun <> 1, Not mbln����)      '���۵�ʱ�����ķѱ�
            If mbln���� Then
                With mobjBill
                    .����ID = IIf(IsNull(mrsInfo!����ID), 0, mrsInfo!����ID)
                    .��ҳID = IIf(mbln���� And mlng��ҳID <> 0, mlng��ҳID, Nvl(mrsInfo!��ҳID, 0))
                    .��ʶ�� = IIf(gint������Դ = 2, Nvl(mrsInfo!סԺ��, 0), Nvl(mrsInfo!�����, 0))
                    .���� = "" & mrsInfo!����
                    .�Ա� = "" & mrsInfo!�Ա�
                    .���� = Trim(txt����.Text) & IIf(IsNumeric(txt����.Text), cbo���䵥λ.Text, "")
                    .���� = "" & mrsInfo!��ǰ����
                    .����ID = IIf(mbln���� And mlngUnitID <> 0, mlngUnitID, Val(Nvl(mrsInfo!��ǰ����ID)))
                    .����ID = IIf(mbln���� And mlngDeptID <> 0, mlngDeptID, Val(Nvl(mrsInfo!��ǰ����id)))
                    .�ѱ� = zlStr.NeedName(cbo�ѱ�.Text) '�Ե�ǰ��ЧΪ׼
                End With
                Bill.SetFocus
            End If
            If Not (mbytInFun = 1 Or mbytInFun = 2 And mbytBilling = 1) Then Call SetDisible(True) 'Active=False
        End If
        
        '����Ʊ���Ƿ��ss��
        If Not mblnStartFactUseType Then Call zlCheckFactIsEnough
        
        If Not txtPatient.Locked And txtPatient.Enabled Then
            txtPatient.SetFocus
        Else
            Bill.SetFocus
        End If
        mblnSaveData = True
    Else '��������ѡ���޸ĵ���
        '����:44196
        mlng������� = 0: SaveChargeFee = True
        Unload Me: Exit Function
    End If
    If mbytInFun = 0 Then
        Call LoadCurBalance
    End If
    mlng������� = 0
    SaveChargeFee = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Exit Function
errH:
    If blnTrans Then
        gcnOracle.RollbackTrans
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
    If blnTrans Then
        'ҽ����HIS����ͬһ������,HIS����ʧ��,��ҽ���������ϴ�,������Ҫ��"ȡ������"�ӿ�
        If blnTransMedicare Then Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, False, mintInsure)
    End If
    cmdOK.Enabled = True
    Call SaveErrLog
End Function
Private Function zlAutoPayDrugAndStuff(ByRef cllDrugAndStuff As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����Զ�����
    '����:���ϳɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2012-02-06 14:55:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If cllDrugAndStuff Is Nothing Then zlAutoPayDrugAndStuff = True: Exit Function
    
    zlExecuteProcedureArrAy cllDrugAndStuff, Me.Caption
    zlAutoPayDrugAndStuff = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub SetӦ���ۼ�(ByVal bln���� As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ӧ���ۼ�
    '����:���˺�
    '����:2012-02-06 14:59:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, dbl����Ӧ�� As Double
    
    mblnNotClearLedDisplay = True
    mbln�������� = False
    If Not (mstrYBPati <> "" And bln���� Or mstrYBPati = "" And bln����) Then Exit Sub
    mbln�������� = True
    For i = 1 To mobjBill.Pages.Count
        mobjBill.Pages(i).Ӧ�ɽ�� = 0
    Next
    If grsTotal.RecordCount <> 0 Then grsTotal.MoveFirst
    dbl����Ӧ�� = 0
    Do While Not grsTotal.EOF
        '����:0-�ɿ�;1-�Ҳ�,2-��Ԥ��;����(mod 10:0-��ͨ����;1-ҽ������;2-������Ʒ;3-һ��ͨ)
        If Val(Nvl(grsTotal!����)) <> 11 Then
            '��ҽ�����ۼ�
            dbl����Ӧ�� = dbl����Ӧ�� + Val(Nvl(grsTotal!������))
        End If
        grsTotal.MoveNext
    Loop
    mobjBill.Pages(1).Ӧ�ɽ�� = dbl����Ӧ��
End Sub
Public Sub zlGetClassMoney(ByRef rsClass As ADODB.Recordset)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������ܽ��
    '����:���˺�
    '����:2011-12-26 13:19:04
    '����:44944
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim p As Integer, strNos As String, dblʵ�ս�� As Double
    Dim i As Integer, j As Integer, rsTemp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
    strNos = ""
    Set rsClass = New ADODB.Recordset
    rsClass.Fields.Append "�շ����", adVarChar, 10, adFldIsNullable
    rsClass.Fields.Append "���", adDouble, , adFldIsNullable
    rsClass.CursorLocation = adUseClient
    rsClass.LockType = adLockOptimistic
    rsClass.CursorType = adOpenStatic
    rsClass.Open
    With mobjBill
        For p = 1 To .Pages.Count
             If .Pages(p).NO <> "" Then        '��ȡ���ǻ��۵�
                  strNos = strNos & "," & .Pages(p).NO & ""
             Else
                For i = 1 To .Pages(p).Details.Count
                    dblʵ�ս�� = 0
                    With .Pages(p).Details(i)
                        For j = 1 To .InComes.Count
                            dblʵ�ս�� = dblʵ�ս�� + .InComes(j).ʵ�ս��
                        Next
                        rsClass.Find "�շ����='" & .�շ���� & "'", , adSearchForward, 1
                        If rsClass.EOF Then rsClass.AddNew
                        rsClass!�շ���� = .�շ����
                        rsClass!��� = Val(Nvl(rsClass!���)) + dblʵ�ս��
                        rsClass.Update
                    End With
                Next
            End If
        Next
    End With
    If strNos = "" Then Exit Sub
    strNos = Mid(strNos, 2)
    strSQL = _
    "  Select /*+ RULE */ A.�շ����,  Sum(ʵ�ս��) As ʵ�ս�� " & _
    "  From ������ü�¼ A, Table( f_Str2list([1])) J" & _
    "  Where A.NO=J.Column_Value and A.��¼����=1 And A.��¼״̬=0  " & _
    " Group By  �շ���� "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�շ����Ļ�����Ϣ", strNos)
    If rsTemp.RecordCount = 0 Then Exit Sub
    Do While Not rsTemp.EOF
        rsClass.Find "�շ����='" & Nvl(rsTemp!�շ����) & "'", , adSearchForward, 1
        If rsClass.EOF Then rsClass.AddNew
        rsClass!�շ���� = Nvl(rsTemp!�շ����)
        rsClass!��� = Val(Nvl(rsClass!���)) + Val(Nvl(rsTemp!ʵ�ս��))
        rsClass.Update
        rsTemp.MoveNext
    Loop
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function zlChargeFeeWin() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����շѽ���
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2012-02-05 16:20:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bytReturnMode As ExitMode, bln���� As Boolean, dbl����Ӧ�� As Double
    Dim blnGetFact As Boolean, i As Integer, p As Integer
    Dim strReturn As String, lng������� As Long, lng����ID As Long
    
    On Error GoTo Errhand
    If Not (mstrYBPati <> "" And MCPAR.���������շ�) And Not mblnSaveAsPrice Then
            Call AutoBultBookFee '�շ�ʱ�Զ�������������Ŀ
    End If
    If isValiedCargeFee = False Then Exit Function
    
    'Ʊ�ݺż������Ѽ����ܽ����ؼ��
    If CheckBillNOAndBookeFee = False Then Exit Function
    If CheckInsure = False Then Exit Function
          
    
    Set mcllPayDrugAndStuff = New Collection
    Set mFrmBalanceWin = New frmChargePayMentWin
    
    lng����ID = mobjBill.����ID
    If mFrmBalanceWin.zlChargeWin(Me, EM_�����շ�, mlngModul, mstrPrivs, _
        mlngShareUseID, mstrUseType, 0, "", "", mobjBill.����ID, mintInsure, _
        mobjBill.����, mobjBill.�Ա�, mobjBill.����, mobjBill.�ѱ�, mdbl�ɿ�, mdbl�Ҳ�, _
        bytReturnMode, CDbl(mcurBillӦ��), bln����, mlngPreBrushCard, dbl����Ӧ��, mstrBalance) = False Then
        If Not mFrmBalanceWin Is Nothing Then Unload mFrmBalanceWin
        '�շ�,���浥��ʧ�ܺ�Ĵ���
         If bytReturnMode <> EM_�������� And bytReturnMode <> EM_�˳��շ� Then
             Call ShowBillChargeFee(mlng�������)
         End If
         mlng������� = 0
        cmdOK.Enabled = True: cmdCancel.Enabled = True
        If mintInsure <> 0 Then
            cmdAddBill.Enabled = Not MCPAR.���������շ� And MCPAR.�൥���շ� And InStr(1, mstrPrivs, "ҽ�����˶൥���շ�") > 0
        Else
            cmdAddBill.Enabled = InStr(1, mstrPrivs, "��ͨ���˶൥���շ�") > 0
        End If
        If cmdDelBill.Visible And tbsBill.Tabs.Count > 1 Then cmdDelBill.Enabled = True
        If cmdOK.Visible And cmdOK.Enabled Then cmdOK.SetFocus
        If bytReturnMode = EM_�������� Then
                If mblnAutoChangePati And gint������Դ = 2 Then
                    '��Ҫ���ҵ�������Դ1��
                    gint������Դ = 1: zlChangePatiSource (gint������Դ)
                End If
                Call ClearFullBill(False)
        End If
        Exit Function
    End If
    If Not mFrmBalanceWin Is Nothing Then Unload mFrmBalanceWin
    lng������� = mlng�������
    '����Ӧ���ۼ�
    Call SetӦ���ۼ�(bln����)
    If mblnDrugPacker Then
        '51510
        Call mobjDrugPacker.DYEY_MZ_TransRecipeDetail(1, UserInfo.���, UserInfo.����, 0, "8," & Replace(Replace(mstrSaveNos, "'", ""), ",", "|8,"), strReturn)
    End If
    '�Զ���ҩ�ͷ��ϴ���
    Call zlAutoPayDrugAndStuff(mcllPayDrugAndStuff)
    
    '��Ϣ����
    Call SendMsgModule
    
    mlng������� = 0
    '��ʾLed:��ҩ���ڼ����úϼƽ��
    Call ShowLedWinAndSum
    'Ʊ�ݴ�ӡ,��ӡƱ��
    Call PrintBill(mstrSaveNos, mstrModiNOs)
    '���������������
    '��ֹ���ô�ӡ�������ķ�ģ̬����,�Լ�ҽ����ʱ
    'д��:56615
    Call WriteMzInforToCard(lng����ID, lng�������)
    
    cmdOK.Enabled = True: cmdCancel.Enabled = True
    If cmdԤ����.Visible Then cmdԤ����.Enabled = True
    If mbytInFun = 0 And mbytInState = 0 And gbln�ۼ� Then
        txt�ۼ�.Text = Format(GetChargeTotal, "0.00")
    End If
    If Not (mstrInNO = "" Or (txtModi.Visible And mbytInState = 0 And mstrInNO <> "")) Then
        '��������ѡ���޸ĵ���
        '����:44196
        zlChargeFeeWin = True
        Unload Me: Exit Function
    End If
    
    '����,����������ͨ�����뵥�ݺ��޸ĵ���
     If fraUpBillShow.Visible Then
        txtPreNO.Text = mobjBill.NO
         txtPreMoney.Text = Format(GetBillSum(False, 1), "0.00")
    Else
        sta.Panels(Pan.C2��ʾ��Ϣ) = "��һ�ŵ���:" & mobjBill.NO '�൥��ʱΪ��һ��
    End If
    i = UBound(Split(mstrSaveNos, ",")) + 1
    If i <> mobjBill.Pages.Count Then
        If MsgBox("Ŀǰ����ֻ�շ���" & i & "�ŵ���,�Ƿ��δ�շѵ��ݽ������շ�!", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
            'ɾ���Ѿ��շѵĵ���
            'ɾ������
            For p = 1 To i
                Call DelOneBill(1)
            Next
            '���¼���
            Call ShowMoney(0)   '�������ݷ���δ��
            '�������ù�����(���������¼���)
            If gTy_Module_Para.bln������ Then
                If Not CheckBillsEmpty Then Call SetFactMoney
            End If
            Exit Function
        End If
    End If
    '������޸ģ����˳��޸�״̬
    If txtModi.Visible And mbytInState = 0 And mstrInNO <> "" Then
        txtModi.Text = "": txtModi.Enabled = True
        If fraBill.Visible Then cmdAddBill.Enabled = InStr(1, mstrPrivs, "��ͨ���˶൥���շ�") > 0
    End If
    mstrInNO = "":  mlngFirstID = 0: mstrFirstWin = ""
    
    If mbytInFun = 0 Or mbytInFun = 1 Then
        '���������ֹ�����շѣ�
        '1.ҽ������ÿ��ˢ��,�����շѽ���(�������ý��ɿ��������)
        '2.ʹ��Ԥ�������,�����շѽ���(�������ý��ɿ��������)
        '3.һ�ζ��ŵ���,�����շѽ���(�������ý��ɿ��������)
        '3.���ѽɿ�,��ǿ����Ϊ�����շѽ���
        '4.����ʱû�����벡������
        '5.ʹ�ö��ֽ��㷽ʽ����
        '6.�շ�ʱ����Ϊ���۵�
        
        '���˺�:22343:gTy_Module_Para.byt�ɿ����:0-�������нɿ�������ۼƿ���,1-��������ɿ��Ž��������ۼ�(�ı䲡�˳���)��2-�շ�ʱ����Ҫ����ɿ���
        'bln���� = Not ((mstrYBPati <> "" And Not gbln�ɿ����) _
                    Or (Val(txtԤ�����.Text) <> 0 And Not gbln�ɿ����) _
                    Or mobjBill.Pages.Count > 1 And Not gbln�ɿ���� _
                    Or Val(txt�ɿ�.Text) <> 0 _
                    Or mobjBill.���� = "" And mbytInFun = 1 _
                    Or mobjBill.Pages(mintPage).�շѽ��� <> "") '���ֽ��㷽ʽ
        
        '�ɿ����:0-�������нɿ�������ۼƿ���,1-��������ɿ��Ž��������ۼ�
        '       2-�շ�ʱ����Ҫ����ɿ���

           '         Or Val(txt�ɿ�.Text) <> 0
        If Not bln���� Then
            If gint������Դ = 2 And mblnAutoChangePati Then
                '�Զ��л���,Ҫ������
                gint������Դ = 1
                Call zlChangePatiSource(gint������Դ)
            End If
            Call ClearPatientInfo(True)
            Call ClearTotalInfo(True)
            Call InitCommVariable
            blnGetFact = IIf(mblnStartFactUseType, False, True)
        Else
            '��Ȼ����,��ҽ��������������Ա��ٴ���֤
            If mstrYBPati <> "" Then Call ClearPatientInfo(True)
            blnGetFact = True
            mstrPrePati = mobjBill.���� '��¼��ǰ����
            mlngPrePati = mobjBill.����ID
            mstrPreDoctor = zlStr.NeedName(cbo������.Text)
            
            '���˵��ݽ���ۼ�
            mcurBillӦ�� = mcurBillӦ�� + GetBillSum(True)
            mcurBillʵ�� = mcurBillʵ�� + GetBillSum
            mcurBillӦ�� = GetMustPaySum
            mintBillNO = mintBillNO + 1
            For i = 1 To mshMoney.Rows - 1
                If mshMoney.TextMatrix(i, 0) = "" Then Exit For
            Next
            mintMoneyRow = i - 1
            Call SaveDrugID(mobjBill.Pages.Count)
        End If
    End If
    Call ClearBillRows
    If (mstrYBPati <> "" And MCPAR.���������շ�) Then
        Call NewYBBill
        mobjBill.����ID = CLng(Split(mstrYBPati, ";")(8))
        
        '���¶�ȡԤ�����
        If txtԤ�����.Enabled Then Call LoadFeeInfor(mobjBill.����ID)
        '���¶�ȡ�������
        mcur������� = gclsInsure.SelfBalance(mobjBill.����ID, CStr(Split(mstrYBPati, ";")(1)), 10, mcur����͸֧, mintInsure)
        sta.Panels(Pan.C3�����ʻ�).Text = "�����ʻ����:" & Format(mcur�������, "0.00")
        sta.Panels(Pan.C3�����ʻ�).Visible = True

        mstrYBPati = ""
    Else
        Call NewBill(blnGetFact, Not Bill.Active And mbytInFun <> 1)      '���۵�ʱ�����ķѱ�
        If Not (mbytInFun = 1 Or mbytInFun = 2 And mbytBilling = 1) Then Call SetDisible(True) 'Active=False
    End If
    '����Ʊ���Ƿ��ss��
    If Not mblnStartFactUseType Then Call zlCheckFactIsEnough
    If Not txtPatient.Locked Then
       If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    Else
        Bill.SetFocus
    End If
    mblnSaveData = True
    Call LoadCurBalance
    zlChargeFeeWin = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 

Private Sub ShowLedWinAndSum()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ��ҩ���ڼ���غϼ�����
    '����:���˺�
    '����:2012-02-06 14:31:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gblnLED = False Then Exit Sub
    If Not (mbytInFun = 0 And Not mblnSaveAsPrice) Then Exit Sub
    If Not (mstr���� <> "" Or mstr�д� <> "" Or mstr�ɴ� <> "") _
        Or CCur(txt�ϼ�.Text) = 0 Then Exit Sub
    zl9LedVoice.DisplayBank "���úϼ�:" & txt�ϼ�.Text, _
        "ȡҩ����:" & IIf(mstr���� <> "", " " & mstr����, "") & _
        IIf(mstr�ɴ� <> "", " " & mstr�ɴ�, "") & IIf(mstr�д� <> "", " " & mstr�д�, "")
End Sub
 


Private Sub cmdOK_Click()
     mblnSaveData = False
    '�ɶ���������(���ﻮ�۱����������ݣ�����)
    If (mbytInFun = 1 Or mbytInFun = 2 And mbytBilling = 1) And mbytInState = 0 And Not Bill.Active Then
        Call ClearBillRows: Call ClearMoney
        Call ClearTotalInfo(True)
        Bill.Active = True
        MsgBox "�������µĵ������ݣ�", vbInformation, gstrSysName
        txtPatient.SetFocus: Exit Sub
    End If
    
    'If Val(txt�ɿ�.Text) = 0 Then txt�ɿ�.Text = "0.00"
    '���ð�ť������,�����ظ�ִ��
    
    If mbytInState = 3 Or (mbytInState = 0 And chkCancel.Value = 1 And chkCancel.Visible) Then
        '========================================================================================================
        '�˷ѻ�����(��������û���������ʹ���˹���)
        If DelCharge = False Then Exit Sub
    ElseIf mbytInState = 2 Then '��������
        '========================================================================================================
        If Not SaveModi() Then Exit Sub
        mblnSaveData = True
        Unload Me
    ElseIf mbytInFun = 2 And mbytBilling = 2 Then '���ʻ��۵���˲���
        If SaveVerify = False Then Exit Sub
    ElseIf (mbytInState = 0) And chkCancel.Value = 0 Then
        '�շѣ����ʣ����ۣ��������뵥��״̬,�շѿ����ǻ��۵����
        '�����쳣���ݵ������շ�
        Call GetAsyncKeyState(VK_RETURN)
        If mbytInFun = 0 And Not mblnSaveAsPrice Then
            If zlChargeFeeWin = False Then Exit Sub
        Else
            If SaveChargeFee = False Then Exit Sub
            
        End If
    
    ElseIf mbytInState = 4 Then
        cmdOK.Enabled = False   '��ֹ���ô�ӡ�������ķ�ģ̬����,�Լ�ҽ����ʱ
        cmdCancel.Enabled = False: cmdAddBill.Enabled = False:: cmdDelBill.Enabled = False
        If ReChargeFee = False Then
            '61688
            cmdOK.Enabled = True   '��ֹ���ô�ӡ�������ķ�ģ̬����,�Լ�ҽ����ʱ
            cmdCancel.Enabled = True
            Exit Sub
        End If
    ElseIf mbytInState = 5 Then
        '�����쳣����
        cmdOK.Enabled = False   '��ֹ���ô�ӡ�������ķ�ģ̬����,�Լ�ҽ����ʱ
        cmdCancel.Enabled = False: cmdAddBill.Enabled = False:: cmdDelBill.Enabled = False
        If DelErrBillFee = False Then
            cmdOK.Enabled = True   '��ֹ���ô�ӡ�������ķ�ģ̬����,�Լ�ҽ����ʱ
            cmdCancel.Enabled = True
            Exit Sub
        End If
    End If
    mblnSaveData = True
    gblnOK = True
    Exit Sub
End Sub

Private Sub LoadFeeInfor(ByVal lngPatientID As Long, Optional ByVal blnDelete As Boolean)
'����:��ȡ����ʾ����Ԥ��,�����������Ϣ
'����:blnDel-�Ƿ��˷ѹ���
    Dim rsTmp As ADODB.Recordset
    Dim curʵ�պϼ� As Currency
 
    If mbytInFun = 0 Then
        Set rsTmp = GetMoneyInfo(lngPatientID, , , 1)
        If Not rsTmp Is Nothing Then
            cmdOK.Tag = rsTmp!Ԥ�����
            cmdCancel.Tag = rsTmp!�������
            cmdPrint.Tag = Val(cmdOK.Tag) - Val(cmdCancel.Tag)
            If mbytInState = 0 And mstrInNO <> "" Then cmdPrint.Tag = Val(cmdPrint.Tag) + Original.��Ԥ����
        Else
            cmdOK.Tag = 0: cmdCancel.Tag = 0: cmdPrint.Tag = 0
        End If
        sta.Panels(Pan.C4Ԥ����Ϣ).Tag = cmdPrint.Tag
        sta.Panels(Pan.C4Ԥ����Ϣ).Text = "Ԥ��:" & Format(Val(cmdPrint.Tag), "0.00")
        Call ShowPrePayInfo(Val(cmdPrint.Tag) > 0 And Not blnDelete)
        
    ElseIf mbytInFun = 2 Then
        '������ʼ����۲���ʹ��Ԥ��,���Բ��ÿ���Original.��Ԥ����
        
        '��¼������������
        '�޸ļ��ʵ�ʱ,����������ǰ���ݽ��(���ʵ�����ʱδ���벡�����ķ������)
        Set rsTmp = GetMoneyInfo(lngPatientID, IIf(mbytBilling = 0, Original.ʵ�պϼ�, 0), , 1)
        If Not rsTmp Is Nothing Then
            cmdOK.Tag = rsTmp!Ԥ�����
            cmdCancel.Tag = rsTmp!�������
            cmdPrint.Tag = Val(cmdOK.Tag) - Val(cmdCancel.Tag)
            sta.Panels(Pan.C4Ԥ����Ϣ).Visible = True
        Else
            cmdOK.Tag = 0: cmdCancel.Tag = 0: cmdPrint.Tag = 0
            sta.Panels(Pan.C4Ԥ����Ϣ).Visible = False
        End If
                
        
        '������ʾʱ������ǰ���ݷ���(��Ϊδ���),�����۱���Ҫ��
        If mbytBilling = 0 Then
            If mbytInState = 0 And mstrInNO <> "" Then
                curʵ�պϼ� = Original.ʵ�պϼ�
            Else
                curʵ�պϼ� = GetBillSum
            End If
        End If
        sta.Panels(Pan.C4Ԥ����Ϣ).Text = "Ԥ��:" & Format(Val(cmdOK.Tag), "0.00")
        sta.Panels(Pan.C4Ԥ����Ϣ).Text = sta.Panels(Pan.C4Ԥ����Ϣ).Text & "/����:" & Format(Val(cmdCancel.Tag) + curʵ�պϼ�, "0.00")
        sta.Panels(Pan.C4Ԥ����Ϣ).Text = sta.Panels(Pan.C4Ԥ����Ϣ).Text & "/ʣ��:" & Format(Val(cmdPrint.Tag) - curʵ�պϼ�, "0.00")
    End If
End Sub

Private Sub cmdCancel_Click()
    Dim blnBillsEmpty As Boolean
    mbln�������� = False
    blnBillsEmpty = CheckBillsEmpty()
    If mbln���� And blnBillsEmpty Then
        Unload Me: Exit Sub
    End If
    
    If (Not blnBillsEmpty Or txtPatient.Text <> "") And mbytInState = 0 And mstrInNO = "" And Not mblnCopyBill Then
        If ClearFullBill(True, Not mbln����) = False Then Exit Sub
        '����:27364 ����:2010-01-13 15:27:50
        If mblnAutoChangePati And gint������Դ = 2 Then
            '��Ҫ���ҵ�������Դ1��
            gint������Դ = 1: zlChangePatiSource (gint������Դ)
        End If
        Exit Sub
    End If
    Unload Me
End Sub

Private Sub SaveDrugID(intPage As Integer)
'����:���浱ǰָ�����������ǰ�����һ������ҩƷ�ĵ��ݵĵ�һ��ҩƷ�Ĳ���ID
    Dim i As Long, j As Long
    
    '��¼�ò��˱����շѷ���ĸ�ҩ��(��������Ϊ����ʱ)
    For i = 1 To intPage
        If mobjBill.Pages(i).NO = "" Then
            j = GetFirstRow(mobjBill, i)
            If j > 0 Then
                Select Case mobjBill.Pages(i).Details(j).�շ����
                    Case "5"
                        mlng��ҩ�� = mobjBill.Pages(i).Details(j).ִ�в���ID
                    Case "6"
                        mlng��ҩ�� = mobjBill.Pages(i).Details(j).ִ�в���ID
                    Case "7"
                        mlng��ҩ�� = mobjBill.Pages(i).Details(j).ִ�в���ID
                End Select
            End If
        Else
            Call BillDrugDept(mobjBill.Pages(i).NO, mlng��ҩ��, mlng��ҩ��, mlng��ҩ��)
        End If
    Next
End Sub

Private Sub cmdOK_GotFocus()
    If mbytInState = 3 Or (chkCancel.Visible And chkCancel.Value = 1) Then
        Bill.Row = 1: Bill.Col = Bill.COLS - 1
    End If
End Sub

Private Sub cmdOK_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '����Ϊ���۵�
    If Button = 2 Then
        If CheckSaveMultiPrice Then
            PopupMenu mnuFile, 2, cmdOK.Left + picAppend.Left - 800, cmdOK.Top + cmdOK.Height + picAppend.Top
        End If
    End If
End Sub
Private Sub cmdPrint_Click()
    Dim i As Integer, j As Integer
    Dim strPrintNO As String, strInfo As String
    Dim blnPrintList As Boolean
    
    If mstrYBBill = "" Then
        MsgBox "��ҽ�����˱��λ�û����ȡ���ã�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If mblnPrint Then
        If gclsInsure.GetCapability(support�����շ���ɺ���֤, mobjBill.����ID, mintInsure) Then
            If gclsInsure.Identify(id����ȷ��, , mintInsure) = "" Then
                MsgBox "���������֤ʧ�ܣ���������շѴ�ӡ������", vbInformation, gstrSysName
                Exit Sub
            End If
            Me.Refresh
        Else
            If MsgBox("ȷʵҪ����շѲ�������ӡƱ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Sub
            End If
        End If
    End If
    
    
    Screen.MousePointer = 11
    
    blnPrintList = False
    If InStr(mstrPrivs, "��ӡ�嵥") > 0 Then
        If gint�շ��嵥 = 1 Then
            blnPrintList = True
        ElseIf gint�շ��嵥 = 2 Then
            If MsgBox("Ҫ��ӡ�շ��嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                blnPrintList = True
            End If
        End If
    End If
    For i = 0 To UBound(Split(mstrYBBill, ","))
        strPrintNO = CStr(Split(mstrYBBill, ",")(i))
        If strPrintNO <> "" Then
            If mblnPrint Then
                If Not gobjTax Is Nothing And gblnTax Then
                    If Not gobjTax Is Nothing And gblnTax Then
                        gstrTax = gobjTax.zlTaxOutPrint(gcnOracle, "'" & strPrintNO & "'")
                        If gstrTax <> "" Then MsgBox gstrTax, vbExclamation, gstrSysName
                    End If
                Else
                    If gblnBillPrint Then
                        If gobjBillPrint.zlPrintBill("'" & strPrintNO & "'", 0) = False Then Exit Sub
                    End If
                    
                    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_1", Me, _
                        "NO='" & strPrintNO & "'", "�۸�ȼ�=" & IIf(mstr��ͨ�۸�ȼ� = "", "-", mstr��ͨ�۸�ȼ�), _
                        IIf(glngFactMediCare = 0, "", "ReportFormat=" & glngFactMediCare), 2)
                End If
            End If
            
            If blnPrintList Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_3", Me, "NO='" & strPrintNO & "'", "ҩƷ��λ=" & IIf(gblnҩ����λ, 1, 0), 2)
            End If
        End If
    Next
    
    mintInsure = 0: mstrYBPati = ""
    cmdPrint.SetFocus
        
    Call ClearFullBill(False)
    txtPatient.SetFocus
    Set grsTotal = Nothing
    Screen.MousePointer = 0
End Sub


Private Function zlSquareCardFeeList(ByRef rsFeeList As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���㿨��ϸ��Ϣ
    '���:
    '����:rsFreeList-������ϸ����
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-01-05 16:02:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, p As Long, strSQL As String, strDate As String, strInvoice As String
    strInvoice = ""
    If zlCreateFeeListStruc(rsFeeList) = False Then Exit Function
    strDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    Err = 0: On Error GoTo Errhand:
    For p = 1 To tbsBill.Tabs.Count
          If mobjBill.Pages(p).NO = "" Then
              'ֱ������ķ���
              If zlBuldingFeeListdata(mobjBill, chk����.Value = 1, p, strDate, cbo�ѱ�.Text, strInvoice, rsFeeList) = False Then Exit Function
          Else
              '��ȡ�Ļ��۵�(�ۼ۵�λ)
              strSQL = _
                "Select '" & strInvoice & "' as ʵ��Ʊ��,NO,To_Date('" & strDate & "','YYYY-MM-DD HH24:MI:SS') as ����ʱ��," & _
                        mobjBill.����ID & " As ����ID,'" & cbo�ѱ�.Text & "' As �ѱ�,�շ����,�վݷ�Ŀ,���㵥λ,������," & _
                "       �շ�ϸĿID,���մ���ID As ����֧������ID,Nvl(������Ŀ��,0) As �Ƿ�ҽ��,���ձ���," & _
                "       Avg(Nvl(����,0)*����) As ����,Avg(��׼����) As ����," & _
                "       Sum(ʵ�ս��) As ʵ�ս��,Sum(ͳ����) As ͳ����,ժҪ," & _
                        chk����.Value & " as �Ƿ���,��������ID,ִ�в���ID,���մ���ID From ������ü�¼" & _
                " Where ��¼����=1 And ��¼״̬=0 And NO=[1]" & _
                " Group By Nvl(�۸񸸺�,���),�շ����,�վݷ�Ŀ,���㵥λ,������," & _
                "       �շ�ϸĿID,���մ���ID,Nvl(������Ŀ��,0),���ձ���,ժҪ,��������ID,ִ�в���ID,NO"
              
              Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mobjBill.Pages(p).NO)
              Do While Not rsTemp.EOF
                    rsFeeList.AddNew
                    rsFeeList!������� = p
                    rsFeeList!�ѱ� = Nvl(rsTemp!�ѱ�)
                    rsFeeList!NO = Nvl(rsTemp!NO)   '����ȡ���۵�ʱ����ֵ
                    rsFeeList!ʵ��Ʊ�� = strInvoice
                    rsFeeList!����ʱ�� = CDate(strDate)
                    rsFeeList!����ID = IIf(mobjBill.����ID = 0, Null, mobjBill.����ID)
                    rsFeeList!�շ���� = Nvl(rsTemp!�շ����)
                    
                    If Nvl(rsTemp!�վݷ�Ŀ) <> "" Then
                        rsFeeList!�վݷ�Ŀ = Nvl(rsTemp!�վݷ�Ŀ)
                    Else
                        rsFeeList!�վݷ�Ŀ = Null
                    End If
                    rsFeeList!������ = Nvl(rsTemp!������)
                    rsFeeList!�շ�ϸĿID = Val(Nvl(rsTemp!�շ�ϸĿID))
                    rsFeeList!���㵥λ = Nvl(rsTemp!���㵥λ)
                    rsFeeList!���� = Val(Nvl(rsTemp!����))
                    rsFeeList!���� = Format(Val(Nvl(rsTemp!����)), gstrFeePrecisionFmt)
                    rsFeeList!ʵ�ս�� = Format(Val(Nvl(rsTemp!ʵ�ս��)), gstrDec)
                    rsFeeList!ͳ���� = Format(Val(Nvl(rsTemp!ͳ����)), gstrDec)
                    rsFeeList!����֧������ID = IIf(Val(Nvl(rsTemp!���մ���ID)) = 0, Null, Val(Nvl(rsTemp!���մ���ID)))
                    rsFeeList!�Ƿ�ҽ�� = Val(Nvl(rsTemp!�Ƿ�ҽ��))
                    rsFeeList!���ձ��� = Nvl(rsTemp!���ձ���)
                    rsFeeList!ժҪ = Nvl(rsTemp!ժҪ)
                    rsFeeList!�Ƿ��� = Val(Nvl(rsTemp!�Ƿ���))
                    rsFeeList!��������ID = Val(Nvl(rsTemp!��������ID))
                    rsFeeList!ִ�в���ID = Val(Nvl(rsTemp!ִ�в���ID))
                    rsFeeList!���ν��� = 0
                    rsFeeList.Update
                    rsTemp.MoveNext
              Loop
          End If
     Next
    zlSquareCardFeeList = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function ����Ԥ���㲻���ֵ���(ByVal strDate As String, ByRef strNone As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ԥ����ʱ,�����ֵ��ݽ��н���(�൥��ֻ����һ�νӿ�)
    '����:���˺�
    '����:2011-08-15 17:30:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strBalance As String, strAdvance As String
    Dim arrPage As Variant, arrBalance() As String, strInvoice As String
    Dim str���㷽ʽ As String, dbl������ As Double, dbl�ɷ���� As Double
    Dim rsTemp  As ADODB.Recordset, i As Long, k As Long, j As Long, p As Long
    
    On Error GoTo errHandle
    strInvoice = Trim(txtInvoice.Text)
    If MCPAR.�൥�ݵ�һ�ν��� = False Then ����Ԥ���㲻���ֵ��� = True: Exit Function
    Set rsTemp = MakeBillRecord(mobjBill, chk����.Value = 1, 0, strDate, cbo�ѱ�.Text, strInvoice)
    strBalance = ""
    strAdvance = tbsBill.Tabs.Count
    
    If Not gclsInsure.ClinicPreSwap(rsTemp, strBalance, mintInsure, strAdvance) Then
        If tbsBill.Tabs.Count > 1 Then
            sta.Panels(Pan.C2��ʾ��Ϣ).Text = "����Ԥ����ʧ�ܡ�"
        End If
        If mstr�����ʻ� <> "" And Not MCPAR.����Ԥ���� Then  'ֻ��ʹ�ø����ʻ�����
            vsBalance.TextMatrix(0, 0) = mstr�����ʻ�
            vsBalance.TextMatrix(0, 1) = "0.00"
            vsBalance.RowData(0) = 0
            Call ShowMoney(-1, Not (cmdԤ����.Visible And cmdOK.Enabled))
        End If
        Screen.MousePointer = 0
        Exit Function
    End If
    
    If strAdvance <> "" And InStr(1, strAdvance, "|") = 0 Then 'ҽ��Ʊ�ݺ�
            txtMCInvoice.Text = strAdvance
            txtMCInvoice.SelStart = Len(txtMCInvoice.Text)
            txtMCInvoice.Visible = True
    End If
    
     '����Ԥ���������ý��㼯
    arrPage = Array()
    If strBalance <> "" Then
        Set rsTmp = GetBalanceSet
        arrBalance = Split(strBalance, "|")
        For i = 0 To UBound(arrBalance)
            str���㷽ʽ = Split(arrBalance(i), ";")(0)
            dbl������ = Val(Split(arrBalance(i), ";")(1))
            '���������øý��㷽ʽ,��Ϊҽ����Ľ��㷽ʽ
            mrs���㷽ʽ.Filter = "����='" & str���㷽ʽ & "' And ����<>1 And ����<>2"
            If mrs���㷽ʽ.EOF Then
                '��¼ҽ���е�����û�еĽ��㷽ʽ
                If InStr(strNone & ",", "," & str���㷽ʽ & ",") = 0 Then
                    strNone = strNone & "," & str���㷽ʽ
                End If
            End If
            If Not mrs���㷽ʽ.EOF Then
                'ֻ�������ŵ��ݵ��ӿ�ʱ�ŷ��ؽ�����Ϣ�����η�̯����������
                For k = 1 To tbsBill.Tabs.Count
                    dbl�ɷ���� = mobjBill.Pages(k).ʵ�ս��
                    rsTmp.Filter = "�������=" & k
                    For j = 1 To rsTmp.RecordCount
                        dbl�ɷ���� = dbl�ɷ���� - rsTmp!������
                        rsTmp.MoveNext
                    Next
                    If dbl�ɷ���� > 0 Then
                        If dbl�ɷ���� <= dbl������ Then
                            dbl������ = dbl������ - dbl�ɷ����
                        Else
                            dbl�ɷ���� = dbl������
                            dbl������ = 0
                        End If
                        
                        rsTmp.AddNew
                        rsTmp!������� = k
                        rsTmp!���㷽ʽ = str���㷽ʽ
                        rsTmp!������ = dbl�ɷ����
                        rsTmp.Update
                        If dbl������ = 0 Then Exit For
                    End If
                Next
                If dbl������ <> 0 Then
                    '���ܴ���ҽ��������ڵ��ݷ����ܶ�����,ֱ�ӷ������һ�ŵ�����
                    rsTmp.Filter = "�������=" & tbsBill.Tabs.Count & " and ���㷽ʽ='" & str���㷽ʽ & "'"
                    If rsTmp.EOF Then
                        rsTmp.AddNew
                        rsTmp!������� = tbsBill.Tabs.Count
                        rsTmp!���㷽ʽ = str���㷽ʽ
                    End If
                    rsTmp!������ = Val(Nvl(rsTmp!������)) + dbl������
                    rsTmp.Update
                    rsTmp.Filter = 0
                End If
            End If
        Next
        For p = 1 To tbsBill.Tabs.Count
            arrPage = Array()
            rsTmp.Filter = "�������=" & p
            For k = 1 To rsTmp.RecordCount
                ReDim Preserve arrPage(UBound(arrPage) + 1)
                arrPage(UBound(arrPage)) = rsTmp!���㷽ʽ & ";" & rsTmp!������ & ";0;" & rsTmp!������
                rsTmp.MoveNext
            Next
            mcolBalance.Remove p '����Ԫ�ز���ֱ���޸�
            If mcolBalance.Count >= p Then
                mcolBalance.Add arrPage, , p
            Else
                mcolBalance.Add arrPage
            End If
        Next
    End If
    ����Ԥ���㲻���ֵ��� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function ����Ԥ�����ֵ���(ByVal strDate As String, ByRef strNone As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ԥ�����ֵ���
    '����:�ɹ�,����true,���򷵻�false
    '����:���˺�
    '����:2011-08-15 18:20:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strBalance As String, strAdvance As String
    Dim arrPage As Variant, arrBalance() As String
    Dim dbl���ʺϼ� As Double, strInvoice As String, dbl������ As Double, dbl�ɷ���� As Double
    Dim str���㷽ʽ As String, p As Long, strSQL As String, i As Long, k As Long, j As Long
    Dim cur���� As Double
    
    strInvoice = Trim(txtInvoice.Text)
    If MCPAR.�൥�ݵ�һ�ν��� Then ����Ԥ�����ֵ��� = True: Exit Function
    On Error GoTo errHandle
    '�Զ��ŵ���ѭ��Ԥ����
    dbl���ʺϼ� = 0
    For p = 1 To tbsBill.Tabs.Count
        If mobjBill.Pages(p).NO = "" Then
            'ֱ������ķ���
            Set rsTmp = MakeBillRecord(mobjBill, chk����.Value = 1, p, strDate, cbo�ѱ�.Text, strInvoice)
        Else
            '��ȡ�Ļ��۵�(�ۼ۵�λ):�������:42961
            strSQL = _
            "   Select '" & strInvoice & "' as ʵ��Ʊ��,NO,  Nvl(�۸񸸺�, ���) as ���,To_Date('" & strDate & "','YYYY-MM-DD HH24:MI:SS') as ����ʱ��," & _
                    mobjBill.����ID & " As ����ID,'" & cbo�ѱ�.Text & "' As �ѱ�,�շ����,�վݷ�Ŀ,���㵥λ,������," & _
            "       �շ�ϸĿID,���մ���ID As ����֧������ID,Nvl(������Ŀ��,0) As �Ƿ�ҽ��,���ձ���," & _
            "       Avg(Nvl(����,0)*����) As ����,Avg(��׼����) As ����," & _
            "       Sum(ʵ�ս��) As ʵ�ս��,Sum(ͳ����) As ͳ����,ժҪ," & _
                    chk����.Value & " as �Ƿ���,��������ID,ִ�в���ID " & _
            "   From ������ü�¼" & _
            "   Where ��¼����=1 And ��¼״̬=0 And NO=[1]" & _
            "   Group By Nvl(�۸񸸺�,���),�շ����,�վݷ�Ŀ,���㵥λ,������," & _
            "       �շ�ϸĿID,���մ���ID,Nvl(������Ŀ��,0),���ձ���,ժҪ,��������ID,ִ�в���ID,NO" & _
            " Order by  ��� "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mobjBill.Pages(p).NO)
        End If
        
        strBalance = ""
        strAdvance = tbsBill.Tabs.Count & "|" & p
        If Not gclsInsure.ClinicPreSwap(rsTmp, strBalance, mintInsure, strAdvance) Then
             '38821:strAdvance:��Ʊ��;�Ƿ���Ʊ�ݺ�
            If tbsBill.Tabs.Count > 1 Then
                sta.Panels(Pan.C2��ʾ��Ϣ).Text = "�� " & p & " �ŵ���Ԥ����ʧ�ܡ�"
            End If
            
            If mstr�����ʻ� <> "" And Not MCPAR.����Ԥ���� Then  'ֻ��ʹ�ø����ʻ�����
                vsBalance.TextMatrix(0, 0) = mstr�����ʻ�
                vsBalance.TextMatrix(0, 1) = "0.00"
                vsBalance.RowData(0) = 0
                
                Call ShowMoney(-1, Not (cmdԤ����.Visible And cmdOK.Enabled))
            End If
            
            Screen.MousePointer = 0
            Exit Function
        End If
        If MCPAR.�൥��һ�ν��� And InStr(1, strAdvance, ";") > 0 Then
              '38821:strAdvance:��Ʊ��;�Ƿ���Ʊ�ݺ�
              MCPAR.ҽ������Ʊ�� = Val(Split(strAdvance & ";", ";")(1)) = 1
        End If
        If strAdvance <> "" And InStr(1, strAdvance, "|") = 0 Then 'ҽ��Ʊ�ݺ�
             '38821:strAdvance:��Ʊ��;�Ƿ���Ʊ�ݺ�
            txtMCInvoice.Text = Trim(Split(strAdvance & ";", ";")(0))
            txtMCInvoice.SelStart = Len(txtMCInvoice.Text)
            txtMCInvoice.Visible = True
        End If
        '����Ԥ���������ý��㼯
        arrPage = Array()
        If strBalance <> "" Then
            If MCPAR.�൥��һ�ν��� Then Set rsTmp = GetBalanceSet
            
            arrBalance = Split(strBalance, "|")
            For i = 0 To UBound(arrBalance)
                str���㷽ʽ = Split(arrBalance(i), ";")(0)
                dbl������ = Val(Split(arrBalance(i), ";")(1))
                
                '���������øý��㷽ʽ,��Ϊҽ����Ľ��㷽ʽ
                mrs���㷽ʽ.Filter = "����='" & str���㷽ʽ & "' And ����<>1 And ����<>2"
                If Not mrs���㷽ʽ.EOF Then
                    If MCPAR.�൥��һ�ν��� Then
                        'ֻ�������ŵ��ݵ��ӿ�ʱ�ŷ��ؽ�����Ϣ�����η�̯����������
                        For k = 1 To tbsBill.Tabs.Count
                            dbl�ɷ���� = mobjBill.Pages(k).ʵ�ս��
                            rsTmp.Filter = "�������=" & k
                            For j = 1 To rsTmp.RecordCount
                                dbl�ɷ���� = dbl�ɷ���� - rsTmp!������
                                rsTmp.MoveNext
                            Next
                            
                            If dbl�ɷ���� > 0 Then
                                If dbl�ɷ���� <= dbl������ Then
                                    dbl������ = dbl������ - dbl�ɷ����
                                Else
                                    dbl�ɷ���� = dbl������
                                    dbl������ = 0
                                End If
                                rsTmp.AddNew
                                rsTmp!������� = k
                                rsTmp!���㷽ʽ = str���㷽ʽ
                                rsTmp!������ = dbl�ɷ����
                                rsTmp.Update
                                
                                If dbl������ = 0 Then Exit For
                            End If
                        Next
                        If dbl������ <> 0 Then
                            '���ܴ���ҽ��������ڵ��ݷ����ܶ�����,ֱ�ӷ������һ�ŵ�����
                            rsTmp.Filter = "�������=" & tbsBill.Tabs.Count & " and ���㷽ʽ='" & str���㷽ʽ & "'"
                            If rsTmp.EOF Then
                                rsTmp.AddNew
                                rsTmp!������� = tbsBill.Tabs.Count
                                rsTmp!���㷽ʽ = str���㷽ʽ
                            End If
                            rsTmp!������ = Val(Nvl(rsTmp!������)) + dbl������
                            rsTmp.Update
                            rsTmp.Filter = 0
                        End If
                    Else
                        If dbl������ <> 0 Or str���㷽ʽ = mstr�����ʻ� Then
                            If str���㷽ʽ = mstr�����ʻ� Then
                                '����ҽ���޷��������
                                If (mcur������� > -1 * mcur����͸֧ Or mintInsure = 61) And CCur(txt�ϼ�.Text) > 0 Then
                                    cur���� = dbl������
                                    If mintInsure <> 61 Then
                                        '��������ʻ�֧�����
                                        If mcur������� - dbl���ʺϼ� - cur���� >= -1 * mcur����͸֧ Then
                                            cur���� = cur���� '������͸֧��Χ���㹻(����͸֧0Ϊ����)
                                        Else
                                            If mcur����͸֧ = 0 And mcur������� - dbl���ʺϼ� > 0 Then
                                                cur���� = mcur������� - dbl���ʺϼ� '������͸֧�������
                                            Else
                                                '��������͸֧��Χ������͸֧ʱ�����
                                                If mcur����͸֧ <> 0 Then
                                                    cur���� = mcur������� - dbl���ʺϼ� + mcur����͸֧ '������͸֧��Χ��֧��
                                                Else
                                                    cur���� = 0
                                                End If
                                            End If
                                        End If
                                    End If
                                    dbl���ʺϼ� = dbl���ʺϼ� + cur����
                                    cur���� = Format(cur����, "0.00")
                                    
                                    ReDim Preserve arrPage(UBound(arrPage) + 1) '���㷽ʽ;ԭʼ(���)���;�ɷ��޸�;�ĺ���
                                    arrPage(UBound(arrPage)) = mstr�����ʻ� & ";" & cur���� & ";" & Split(arrBalance(i), ";")(2) & ";" & cur����
                                End If
                            Else
                                ReDim Preserve arrPage(UBound(arrPage) + 1)
                                arrPage(UBound(arrPage)) = arrBalance(i) & ";" & Format(dbl������, "0.00")
                            End If
                        End If
                    End If
                Else
                '��¼ҽ���е�����û�еĽ��㷽ʽ
                    If InStr(strNone & ",", "," & str���㷽ʽ & ",") = 0 Then
                        strNone = strNone & "," & str���㷽ʽ
                    End If
                End If
            Next
        End If
        
        If Not MCPAR.�൥��һ�ν��� Then
            'ÿ�����ݶ�Ӧһ������,�������û��Ԫ��
            mcolBalance.Remove p '����Ԫ�ز���ֱ���޸�
            If mcolBalance.Count >= p Then
                mcolBalance.Add arrPage, , p
            Else
                mcolBalance.Add arrPage
            End If
        End If
    Next
    
    If MCPAR.�൥��һ�ν��� And strBalance <> "" Then
        For p = 1 To tbsBill.Tabs.Count
            arrPage = Array()
            rsTmp.Filter = "�������=" & p
            For k = 1 To rsTmp.RecordCount
                ReDim Preserve arrPage(UBound(arrPage) + 1)
                arrPage(UBound(arrPage)) = rsTmp!���㷽ʽ & ";" & rsTmp!������ & ";0;" & rsTmp!������
                rsTmp.MoveNext
            Next
            mcolBalance.Remove p '����Ԫ�ز���ֱ���޸�
            If mcolBalance.Count >= p Then
                mcolBalance.Add arrPage, , p
            Else
                mcolBalance.Add arrPage
            End If
        Next
    End If

    ����Ԥ�����ֵ��� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function ����Ԥ����(ByRef strNone As String) As Boolean
    '���ܣ�����Ԥ����
    Dim arrBalance() As String, dbl���ʺϼ� As Double
    Dim i As Integer, j As Integer, k As Integer, p As Integer
    Dim strDate As String, str���㷽ʽ As String
    Dim dbl�ϼ� As Double
    
    strNone = ""
    Screen.MousePointer = 11
    On Error GoTo errH
    '��ʼ�����������
    Call InitBalanceGrid
    '��ȡ����ʱ��
    strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    If MCPAR.�൥�ݵ�һ�ν��� = True Then
        If ����Ԥ���㲻���ֵ���(strDate, strNone) = False Then Exit Function
    Else
        If ����Ԥ�����ֵ���(strDate, strNone) = False Then Exit Function
    End If
        
    'ȫ��Ԥ�����Ĵ���
    '-----------------------------------------------------------
    '��ʾԤ��ı����
    For p = 1 To mcolBalance.Count
        For i = 0 To UBound(mcolBalance(p))
            '���㷽ʽ;ԭʼ(���)���;�ɷ��޸�;�ĺ���
            arrBalance = Split(mcolBalance(p)(i), ";")
            
            '��λ��ƥ���л����
            k = -1
            For j = 0 To vsBalance.Rows - 1
                If vsBalance.TextMatrix(j, 0) = arrBalance(0) Then
                    k = j: Exit For '��¼����д��ƥ����
                ElseIf vsBalance.TextMatrix(j, 0) = "" Then
                    If k = -1 Then k = j '��¼��һ���ÿ���
                End If
            Next
            If j > vsBalance.Rows - 1 And k = -1 Then
                vsBalance.Rows = vsBalance.Rows + 1
                k = vsBalance.Rows - 1
            End If
            
            '���ܸ��ֽ��㷽ʽ�Ľ��
            vsBalance.TextMatrix(k, 0) = arrBalance(0)
            vsBalance.TextMatrix(k, 1) = Format(Val(vsBalance.TextMatrix(k, 1)) + Val(arrBalance(1)), "0.00")
            dbl�ϼ� = dbl�ϼ� + Val(Format(Val(arrBalance(1)), "0.00"))
            If vsBalance.RowData(k) = 0 Then
                '���ŵ�����,ֻҪ��һ�������޸�,����ܵ������޸�
                vsBalance.RowData(k) = arrBalance(2)
            End If
        Next
    Next
    
    For i = 0 To vsBalance.Rows - 1
        If vsBalance.RowData(i) <> 0 Then
            vsBalance.Row = i: vsBalance.Col = 1
            vsBalance.TabStop = True
            Exit For
        End If
    Next
    
    
    'Ҫ�������Ա������ط�ʶ��
    If cmdԤ����.Visible Then
        cmdԤ����.TabStop = False
        cmdOK.Enabled = True
    End If
    '���¼���Ӧ�ɣ����(�ֱ�)��
    Call ShowMoney(-1, Not (cmdԤ����.Visible And cmdOK.Enabled))
    With vsBalance
        For i = 0 To .Rows - 1
            If Trim(.TextMatrix(i, 0)) = "" Then Exit For
        Next
        If i > .Rows - 1 Then .Rows = .Rows + 1
        .TextMatrix(i, 0) = "�Ը��ϼ�": .TextMatrix(i, 1) = txtӦ��.Text
        
        .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbRed
        .Cell(flexcpFontBold, i, 0, i, .COLS - 1) = vbRed
        .RowPosition(i) = 0
    End With
    Call zl9InsureLedSpeak
'    dbl���ʺϼ� = GetMedicareSum(mstr�����ʻ�)
'    If gblnLED Then zl9LedVoice.DisplayBank "ҽ������:", "�ʻ����" & Format(mcur�������, "0.00"), "�ʻ�֧��" & Format(dbl���ʺϼ�, "0.00"), "ͳ��֧��" & Format(GetMedicareSum - dbl���ʺϼ�, "0.00")
    strNone = Mid(strNone, 2)
    If strNone = "" Then ����Ԥ���� = True
    Screen.MousePointer = 0
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub zl9InsureLedSpeak()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ҽ��Ԥ��Led����
    '����:���˺�
    '����:2011-12-15 13:40:46
    '����:44425
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl���ʺϼ� As Double
    If Not gblnLED Then Exit Sub
    dbl���ʺϼ� = GetMedicareSum(mstr�����ʻ�)
    zl9LedVoice.DisplayBank "ҽ������:", "�ʻ����" & Format(mcur�������, "0.00"), "�ʻ�֧��" & Format(dbl���ʺϼ�, "0.00"), "ͳ��֧��" & Format(GetMedicareSum - dbl���ʺϼ�, "0.00")
    zl9LedVoice.Speak "#21 " & Format(Val(txtӦ��.Text), "0.00")
End Sub

Private Sub cmdԤ����_Click()
    Dim strNone As String
    Call AutoBultBookFee '�շ�ʱ�Զ�����������
    
    If CheckBillsEmpty Then Exit Sub
    If gbytAutoSplitBill > 0 Then Call AutoSplitBill
                  
    If mbytInFun = 0 And mintInsure <> 0 And MCPAR.ʵʱ��� Then
        '�������ڻ��۵��Ŵ�2������ϸ�ͻ��ܵļ�飬���ǣ���������ԭ��������ʵ�ս����������ͨ������ܸı䣬�������ٴμ����ϸ
        '1.���뵥�ݣ�2.�޸ĵ��ݣ�3.������ҩ�䷽��4.�޸���ҩ�����������еĸ���ͬʱ�仯��5.��������Զ���������Լ�������ܼ����ۿ�
        '6.�޸ĵ��ۣ�7.����ִ�п��ң�ҩƷ�۸����㣬8.�����ѱ�ʵ�ս������,9.�����������֤ҽ�����,�����ȵ�
        If gclsInsure.CheckItem(mintInsure, 0, 2, MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), 1, 0)) = False Then
            Exit Sub
        End If
    End If
    
    'Ԥ����
    If Not ����Ԥ����(strNone) Then
        If strNone <> "" Then
            MsgBox "��ǰ���ս���ʹ�õĽ��㷽ʽ" & vbCrLf & vbCrLf & vbTab & strNone & vbCrLf & vbCrLf & _
                "������δ���ã����ȵ����㷽ʽ������������Щ���㷽ʽ��", vbInformation, gstrSysName
        End If
        cmdԤ����.TabStop = True
        cmdOK.Enabled = False
        cmdԤ����.SetFocus
        Exit Sub
    Else
    
    End If
    If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus
End Sub

Private Sub Form_Activate()
    Dim objTemp As Object
    
    If mblnFirst = False Then Exit Sub
    mblnFirst = False: mblnNotClearLedDisplay = False
    If LoadBill = False Then Unload Me: Exit Sub
    If mbytInState = 5 Then cmdOK_Click: Exit Sub
    
    On Error Resume Next
    If mblnCopyBill Then
        cmdOK.SetFocus
    ElseIf mbytBilling = 2 Then
        cboNO.SetFocus
    ElseIf mbytInState = 1 Then
        cmdCancel.SetFocus
    ElseIf mbytInState = 2 Then
        txtDate.SetFocus
    ElseIf mbytInState = 0 And mstrInNO <> "" And Bill.Active Then
        Bill.SetFocus
    ElseIf mbytInState = 3 Then
        cmdOK.SetFocus
    End If
    
    '˫����ʾ��������ڵ�ǰ������ʾ֮�������ʾ�����ƶ�����
    If mbytInFun = 0 And mbytInState = 0 And gblnLED And Trim(txtPatient.Text) = "" Then
        zl9LedVoice.DisplayPatient ""
    End If
    DoEvents
    If mblnSetControl Then
        mblnSetControl = False
        Set objTemp = Me.ActiveControl
        If cboTemp.Visible And cboTemp.Enabled Then cboTemp.SetFocus
        If objTemp.Visible And objTemp.Enabled Then objTemp.SetFocus
    End If
    
    If mbln���� Then
        Call Set���˲��ѱ༭����
        If Bill.Active Then Bill.SetFocus
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If Me.ActiveControl Is Bill Then Exit Sub
    If InStr("',|~:��;��?��" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If InStr("`��", Chr(KeyAscii)) > 0 Then
        '�����ʾ���￨
         KeyAscii = 0
        If gblnLED Then zl9LedVoice.Speak "#30"  '`Ϊ��������:�е����:����Ӧ����192,����֪��ô���229:32663
    End If
End Sub
Private Sub InitCommVariable()
    If Not mbln�������� Then
        mcurBillӦ�� = 0
        mcurBillʵ�� = 0: mcurBillӦ�� = 0:
    End If
    mlng��ҩ�� = 0: mlng��ҩ�� = 0: mlng��ҩ�� = 0
    mstr���� = "": mstr�д� = "": mstr�ɴ� = ""
    mintBillNO = 0: mintMoneyRow = 0
    
    With mTyDelFee
        .strNos = ""
        Set .rsBlance = Nothing
        .blnSingleBalance = False
        .dblCurDelMoney = 0
        .bln������ȫ�� = False
    End With
End Sub

Private Sub InitBillColumnColor()
        Bill.SetColColor BillCol.���, &HE7CFBA
        Bill.SetColColor BillCol.��Ŀ, &HE7CFBA
        Bill.SetColColor BillCol.����, &HE7CFBA
        Bill.SetColColor BillCol.ִ�п���, &HE7CFBA
        Bill.SetColColor BillCol.����, &HE0E0E0
        Bill.SetColColor BillCol.����, &HE0E0E0
        Bill.SetColColor BillCol.��־, &HE0E0E0
End Sub

Private Sub ClearPayInfo()
    txtӦ��.Text = "0.00"
End Sub

Private Sub ClearTotalInfo(Optional ByVal bln����ۼ� As Boolean = False)
'Ĭ��blnΪfalse,������ۼ�,(����ʱ�ۼ�txtbox��ΪӦ����ʾ)
    txt�ϼ�.Text = gstrDec: txtӦ��.Text = gstrDec
    If bln����ۼ� Then
        If mbytInFun = 1 Then txt�ۼ�.Text = "0.00"
    End If
End Sub

Private Sub ClearPatientInfo(Optional ByVal bln������� As Boolean = False)
'Ĭ��blnΪfalse���������txtbox
    If bln������� Then
        mstrPrePati = ""
        mlngPrePati = 0
        mstrPreDoctor = ""
        txtPatient.Text = ""
        txtPatient.Locked = False
        txtPatient.BackColor = &HFFFFFF
        If mbytInFun = 2 Then lblCorp.Visible = False: lblCorp.Caption = ""
    End If
    txt����.Text = "": txt�����.Text = ""
    Call zlControl.CboLocate(cbo���䵥λ, "��")
    Call txt����_Validate(False)
    lbl����.Caption = ""
End Sub

Private Sub ClearmobjBill()
    With mobjBill
        .���� = ""
        .�Ա� = ""
        .���� = ""
        .����ID = 0
        .��ҳID = 0
        .��ʶ�� = 0
        .���� = ""
        
        .����ID = 0
        .����ID = 0
        .Ӥ���� = 0
        .�ѱ� = zlStr.NeedName(cbo�ѱ�.Text)
        .�����־ = gint������Դ
        .�Ӱ��־ = chk�Ӱ�.Value
    End With
End Sub
Private Function CheckDepend() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ƿ���ڹ�������
    '����:���������,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-22 16:49:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    

    CheckDepend = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub Form_Load()
    mblnFirst = True: mbln�������� = False
    mblnHaveExcuteData = False
    mblnSetControl = True
    mblnStartFactUseType = zlStartFactUseType(1)
    '----------------------------�����ʼ-------------------------------------
    If glngSys Like "8??" Then
        lblPatient.Caption = "�ͻ�����"
        lbl�ѱ�.Caption = "��Ա�ȼ�"
        lbl�����.Caption = "�ͻ���"
        lbl����.Visible = False
        cbo��������.Visible = False
        lbl�ѱ�.Left = lblPatient.Left
        cbo�ѱ�.Left = txtPatient.Left
        cbo�ѱ�.Width = txtPatient.Width
        mshMoney.Visible = False
        fraStat.Left = mshMoney.Left
        vsBalance.Left = fraStat.Left + fraStat.Width + 30
        fra�ɿ�.Left = vsBalance.Left + vsBalance.Width + 30
    End If
    
    '��С����ߴ�
    glngFormW = 12000: glngFormH = 7710
    If Not OS.IsDesinMode Then
        glngOld = GetWindowLong(hWnd, GWL_WNDPROC)
        Call SetWindowLong(hWnd, GWL_WNDPROC, AddressOf Custom_WndMessage)
    End If
    
    'Ӧ�÷������Ƴߴ�֮��
    RestoreWinState Me, App.ProductName, mbytInFun & mbytInState
    sta.Visible = True
     
    If mbytInFun = 0 And (mbytInState = 0 Or mbytInState = 3 Or mbytInState = 4 Or mbytInState = 5) Then
        If glng���ϸĿID = 0 Then
            MsgBox "ϵͳ����δ������Ч��������Ŀ��", vbInformation, gstrSysName
            Unload Me: Exit Sub
        End If
    End If
    '----------------------------�����������ʼ��------------------------------
    'LED
    If mbytInFun = 0 And (mbytInState = 0 Or mbytInState = 4 Or mbytInState = 5) And gblnLED Then
        zl9LedVoice.Reset com
        zl9LedVoice.Init UserInfo.��� & " �շ�ԱΪ������", mlngModul, gcnOracle
    End If
    
    '����:51510
    Call CreateDrugPacker '�����Զ���ҩ������
    mblnDrugPacker = False
    If mobjDrugPacker Is Nothing And mbytInFun = 0 And (mbytInState = 0 Or mbytInState = 2 Or mbytInState = 3 Or mbytInState = 4) Then
        Err = 0: On Error Resume Next
        Set mobjDrugPacker = CreateObject("zlDrugPacker.clsDrugPacker")
        If Err <> 0 Then
            mblnDrugPacker = False
        Else
            mblnDrugPacker = mobjDrugPacker.DYEY_MZ_IniSoap
        End If
    End If
    
    Call ClearTotalInfo(True)
    lblSubӦ��.Caption = "Ӧ��:" & gstrDec
    lblSubʵ��.Caption = "ʵ��:" & gstrDec
    lblAmount.Caption = ""
    
    'ģ�����
    Call InitCommVariable
    
    gbln�������� = False
    gblnOK = False:         mblnLoad = False:           mblnDoing = False
    mblnDo = True:          mblnEnterCell = True:       mbln������۸� = False
    mblnCboClick = False
    mstrPrePati = "":       mlngPrePati = 0:            mstr���ʽ = ""
    mstr�����ʻ� = "":      mblnValid = False:          mstrPreDoctor = ""
    mblnF2Save = False:     mblnAutoChangePati = False
    
    '���ݶ���
    mintPage = 1
    Set mobjBill = New ExpenseBill
    Set mcolBalance = New Collection    '�ü�������Ԥ����,�뵥�ݱ�ǩ����һ��
    mcolBalance.Add Array()
    Set mrsInfo = New ADODB.Recordset
    
    '-------------------------���ݳ�ʼ������------------------------------------
    '�鿴����ʱ�������ʼ����
    If mbytInState = 0 Or mbytInState = 2 Or mbytInState = 3 Or mbytInState = 4 Or mbytInState = 5 Then
        If mbytInState = 0 Then
            mstrҩƷ�۸�ȼ� = gstrҩƷ�۸�ȼ�
            mstr���ļ۸�ȼ� = gstr���ļ۸�ȼ�
            mstr��ͨ�۸�ȼ� = gstr��ͨ�۸�ȼ�
        End If
        If Not InitData Then Unload Me: Exit Sub
    Else
        '���䵥λ
        cbo���䵥λ.AddItem "��"
        cbo���䵥λ.AddItem "��"
        cbo���䵥λ.AddItem "��"
        cbo���䵥λ.ListIndex = 0
    End If
    Call InitFace   'InitData��Ҫ�ڴ�֮ǰ
End Sub
Private Function LoadBill() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ص�������
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-22 16:41:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Select Case mbytInState
    Case 0 'b.����,�޸�
        If mbytInFun = 0 And mbytInState = 0 And gbln�ۼ� Then
            txt�ۼ�.Text = Format(GetChargeTotal, "0.00")
            txt�ۼ�.ToolTipText = "��ǰ����Ա�����շ��ۼƶ�"
        End If
        '1.��������
        If Not NewBill(Not mblnStartFactUseType, False) Then Exit Function           '����false��ʾ�����ٶ�ȡ���÷ѱ�,��Ϊǰ��InitData�����˲���
        '2.�޸ĵ���,�൥���շ�ʱ���޸ĵ��ǵ�ǰѡ�е���һ�ŵ���
        If mstrInNO <> "" Then
            Call LoadModifyNO(mstrInNO, IIf(mbytInFun = 2, 2, 1))
        Else
            If mlng����ID <> 0 Then
                txtPatient.Text = "-" & mlng����ID
                Call txtPatient_KeyPress(13)
            End If
        End If
        LoadBill = True: Exit Function
    Case 4, 5     '�쳣���ݵĴ���
        If mstrInNO = "" Then Exit Function
        If LoadErrBillCharge(mstrInNO) = False Then Exit Function
        LoadBill = True: Exit Function
    Case 1, 2, 3    'a.��ʾ����������,�����˷�
        If mbytInState = 3 Then
            If Not ReadBill(mstrInNO, mbytInFun, True) Then Exit Function
        Else
            If Not ReadBill(mstrInNO, mbytInFun) Then Exit Function
        End If
        If InStr(mstrPrivs, "��ʾ������") = 0 Then
            cbo������.Visible = False
            If gbyt����ҽ�� = 0 Then
                lbl����.Visible = False
            Else
                lbl������.Visible = False
            End If
        End If
        cboNO.Text = mstrInNO
        LoadBill = True: Exit Function
    End Select
    LoadBill = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Sub Form_Resize()
    Dim lngCancelW As Long
    Dim lngLeft As Long
    On Error Resume Next
    
    fraTitle.Left = 0
    fraTitle.Width = Me.ScaleWidth
    
    fraInfo.Left = 0
    fraInfo.Width = Me.ScaleWidth
    
    fraBill.Left = 0
    fraBill.Top = fraInfo.Top + fraInfo.Height
    fraBill.Width = Me.ScaleWidth
    cmdDelBill.Left = fraBill.Width - cmdDelBill.Width - 60
    cmdAddBill.Left = cmdDelBill.Left - cmdAddBill.Width
    tbsBill.Width = cmdAddBill.Left - tbsBill.Left - 300
    
    If fraBill.Visible Then
        Bill.Top = fraBill.Top + fraBill.Height
    Else
        Bill.Top = fraInfo.Top + fraInfo.Height
    End If
    Bill.Width = Me.ScaleWidth - Bill.Left
    If vsInvoice.Visible Then
        '25187
        With vsInvoice
            .Left = Bill.Left
            .Width = Bill.Width
        End With
        Call SetInvoceSizeAndShowTittle
    End If
    Bill.Height = Me.ScaleHeight - Bill.Top - sta.Height - picAppend.Height - IIf(fraSubBill.Visible, fraSubBill.Height + 30, 0) _
        - IIf(fra�˷�ժҪ.Visible, fra�˷�ժҪ.Height + 30, 0) _
        - IIf(vsInvoice.Visible, vsInvoice.Height + 30, 0)
    If fraSubBill.Visible Then
        fraSubBill.Left = Bill.Left
        fraSubBill.Width = Bill.Width
        fraSubBill.Top = Bill.Top + Bill.Height + 15
        lblSubʵ��.Left = fraSubBill.Width - 2250
        lblSubӦ��.Left = lblSubʵ��.Left - 2250
        lblAmount.Left = lblSubӦ��.Left - 2250
    End If
    If fra�˷�ժҪ.Visible Then
        With fra�˷�ժҪ
             .Left = Bill.Left
             .Width = Bill.Width
             .Top = Bill.Top + Bill.Height + 15
             txt�˷�ժҪ.Width = .Left + .Width - txt�˷�ժҪ.Left - 50
        End With
    End If
    '25187
    With vsInvoice
         .Top = IIf(fra�˷�ժҪ.Visible, fra�˷�ժҪ.Height + fra�˷�ժҪ.Top + 15, Bill.Top + Bill.Height + 15)
    End With
    
    cbo���㷽ʽ.Left = fraAppend.Left + lbl���㷽ʽ.Left + lbl���㷽ʽ.Width + 30
    cbo���㷽ʽ.Top = fraAppend.Top + lbl���㷽ʽ.Top - (cbo���㷽ʽ.Height - lbl���㷽ʽ) / 2
    
    cmdRegist.Left = fraTitle.Width - cmdRegist.Width - 90
    cmdIDCard.Left = fraTitle.Width - IIf(cmdRegist.Visible, cmdRegist.Width + 90, 0) - cmdIDCard.Width - 90
    
    lngLeft = fraTitle.Width - 90
    lngLeft = IIf(cmdRegist.Visible, cmdRegist.Left - 50, lngLeft)
    lngLeft = IIf(cmdIDCard.Visible, cmdIDCard.Left - 50, lngLeft)
    cmdSaveWholeSet.Left = lngLeft - cmdSaveWholeSet.Width
    lngLeft = IIf(cmdSaveWholeSet.Visible, cmdSaveWholeSet.Left - 50, lngLeft)
    cmdSelWholeSet.Left = lngLeft - cmdSelWholeSet.Width
    
    lngLeft = IIf(cmdSelWholeSet.Visible, cmdSelWholeSet.Left - 50, lngLeft)
    
    lblFormat.Left = lngLeft - lblFormat.Width
    'fraTitle.Width - IIf(cmdRegist.Visible, cmdRegist.Width + 90, 0) - IIf(cmdIDCard.Visible, cmdIDCard.Width + 90, 0) - lblFormat.Width - 90
    If cmdDelete.Visible Or chkCancel.Visible Or lblFlag.Visible Then lngCancelW = chkCancel.Width
    chkCancel.Left = fraTitle.Width - chkCancel.Width - 60
    lblFlag.Left = chkCancel.Left + (chkCancel.Width - lblFlag.Width) / 2
    cmdDelete.Left = chkCancel.Left
    
    cboNO.Left = fraTitle.Width - lngCancelW - 60 - cboNO.Width - 30
    lblNO.Left = cboNO.Left - lblNO.Width - 30
    
    txtInvoice.Left = lblNO.Left - txtInvoice.Width - 40
    lblFact.Left = txtInvoice.Left - lblFact.Width - 40
    txtMCInvoice.Left = txtInvoice.Left
    
    fraAppend.Width = Me.ScaleWidth - fraAppend.Left
    
    txtDate.Left = fraAppend.Width - txtDate.Width - 90
    lblDate.Left = txtDate.Left - lblDate.Width - 45
    If mbytInFun <> 0 Then
        cmdOK.Left = ScaleWidth - cmdOK.Width - 100
        cmdCancel.Left = cmdOK.Left
        cmdPrint.Left = cmdOK.Left
        cmdԤ����.Left = cmdOK.Left
    End If
    If mbytInFun <> 2 Then
        If TypeName(cbo���㷽ʽ.Container) = TypeName(cbo������.Container) Then
            lbl������.Left = IIf(cbo���㷽ʽ.Visible, cbo���㷽ʽ.Left + cbo���㷽ʽ.Width + 100, lbl���㷽ʽ.Left)
            cbo������.Left = lbl������.Left + lbl������.Width + 20
        Else
             lbl������.Left = IIf(cbo���㷽ʽ.Visible, cbo���㷽ʽ.Left + cbo���㷽ʽ.Width + 100, lbl���㷽ʽ.Left)
            cbo��������.Left = lbl������.Left + lbl������.Width + 20
        End If
    End If
    Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mbytInFun = 0 And mbytInState = 0 And mstrYBPati <> "" And mstrInNO = "" Then
        If MsgBox("��ǰ���ڶ�ҽ�������շѣ�ȷʵҪ�˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1: Exit Sub
        End If
        If YBIdentifyCancel = False Then        'ȡ��ҽ�����������֤,���ؼ�ʱ���˳�
            Cancel = 1: Exit Sub
        End If
    End If
    
    SaveWinState Me, App.ProductName, mbytInFun & mbytInState
    If mbytInState = 0 Then
        Call SaveRegisterItem(g˽��ģ��, Me.Name, "idkind", IDKind.IDKind)
    End If
    
    zlCommFun.OpenIme False
        
    mbytInFun = 0
    mbytInState = 0
    mblnCopyBill = False
    mstrInNO = ""
    mstrTime = ""
    mblnDelete = False
    mbytBilling = 0
    mstrCardNO = ""
    mblnNOMoved = False   '�鿴ʱ,���ܴ���true,
    mblnYB�������� = False
    
    mintBillNO = 0: mintMoneyRow = 0
    mlngFirstID = 0: mstrFirstWin = ""
    mlng����ID = 0
    mlngҩƷ���ID = 0
    mlng�������ID = 0
    
    mlng����ID = 0
    mlng��ҳID = 0
    mlngUnitID = 0
    mlngDeptID = 0
    mbln���� = False
    mlng����ҽ�� = 0
    mstr���ת��ʱ�� = ""
    
    '������ݶ���
    Set mrs�������� = Nothing
    Set mrs������ = Nothing
    Set mrs�ѱ� = Nothing
    Set mrs�������� = Nothing
    Set mrs��ҩ���� = Nothing
    Set mrsWarn = Nothing
    Set mobjCard = Nothing
    Set mobjBrushCheck = Nothing
    
    'LED��ʼ��
    If mbytInFun = 0 And mbytInState = 0 And gblnLED Then
        zl9LedVoice.DisplayPatient ""
        zl9LedVoice.Reset com
    End If
    If Not mobjIDCard Is Nothing Then
        Call mobjIDCard.SetEnabled(False)
        Set mobjIDCard = Nothing
    End If
    If Not mobjICCard Is Nothing Then
        Call mobjICCard.SetEnabled(False)
        Set mobjICCard = Nothing
    End If
    mintInvoicePrint = 0
    If Not OS.IsDesinMode Then
        Call SetWindowLong(hWnd, GWL_WNDPROC, glngOld)
    End If
    
    If Not mobjDrugPacker Is Nothing Then
        '51510
        Set mobjDrugPacker = Nothing
    End If
    mblnHaveExcuteData = False
End Sub

Private Sub mobjBrushCheck_ReadCardNoed(ByVal strCardNo As String, ByVal blnBrushCard As Boolean)
    If blnBrushCard Then
        mbln����ˢ�� = True
    Else
        mbln����ˢ�� = False
    End If
End Sub

Private Sub mnuFileSavePrice_Click()
    '����Ϊ���۵�
    mnuFileSavePrice.Checked = True
    mblnSaveAsPrice = True
    
    Call DelFactMoney  'ɾ��������
    Call cmdOK_Click
    If mnuFileSavePrice.Checked Then '������˳�
        mnuFileSavePrice.Checked = False
        mblnSaveAsPrice = False
    End If
End Sub
Private Sub ReCalce�˿�()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���¼����˿�
    '����:���˺�
    '����:2011-11-21 17:27:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    txtӦ��.Text = Format(GetDelMoney, "0.00")
End Sub
Private Sub ModiyVsBalance()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�޸Ľ�������(Ŀǰֻ���޸�(���㿨�����ѿ�)����)
    '����:���˺�
    '����:2011-11-21 17:23:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not (mbytInState = 3 And mbytInFun = 0) Then Exit Sub
    With vsBalance
        '1-Ԥ���,2-ҽ��,3-ҽ�ƿ�,4-���㿨,5-һ��ͨ,0-������
        If InStr("3,4,5", .Cell(flexcpData, .Row, 0)) = 0 Then Exit Sub
        If .RowData(.Row) <> -1 Then Exit Sub
        
        If Val(.TextMatrix(.Row, 1)) <> 0 Then
            .Cell(flexcpForeColor, .Row, 0, .Row, .COLS - 1) = vbRed
        Else
            .Cell(flexcpForeColor, .Row, 0, .Row, .COLS - 1) = Me.ForeColor
        End If
    End With
    Call ReCalce�˿�
End Sub
Private Sub txtPatient_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtPatient.Locked Or txtPatient.Enabled = False Or txtPatient.Text <> "" Then Exit Sub
    If IDKind.ActiveFastKey = True Then Exit Sub
End Sub

Private Sub vsBalance_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        Call ReCalce�˿�
End Sub

Private Sub vsBalance_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
   If Not (mbytInState = 3 And mbytInFun = 0) Then Cancel = True: Exit Sub
    With vsBalance
        '1-Ԥ���,2-ҽ��,3-ҽ�ƿ�,4-���㿨,5-һ��ͨ,0-������
        If InStr("3,4,5", .Cell(flexcpData, .Row, 0)) = 0 Then Cancel = True: Exit Sub
'        If .RowData(.Row) <> -1 Or vsBalance.Tag = "1" Then Cancel = True: Exit Sub
'        .ColComboList(Col) = " ||" & Val(.Cell(flexcpData, Row, Col))
        If .RowData(.Row) <> -1 Then Cancel = True: Exit Sub '������
        If mTyDelFee.blnSingleBalance And mTyDelFee.bln������ȫ�� = False And .Cell(flexcpData, Row, 0) = 3 Then
            .ColComboList(Col) = " ||" & FormatEx(mTyDelFee.dblCurDelMoney, 2): Exit Sub
        End If
        If vsBalance.Tag = "1" Then Cancel = True: Exit Sub
        .ColComboList(Col) = " ||" & FormatEx(Val(.Cell(flexcpData, Row, Col)), 2)
    End With
End Sub

Private Sub vsBalance_DblClick()
    Dim lngRow As Long
    'Call ModiyVsBalance
    
    If Not (mbytInState = 0 And chkCancel.Value = 0) Then Exit Sub
    
    If vsBalance.MouseCol <> 1 Then Exit Sub
    lngRow = vsBalance.MouseRow
    If vsBalance.RowData(lngRow) <> 0 And vsBalance.TextMatrix(lngRow, 0) <> "" Then
        txtTmp.Text = vsBalance.TextMatrix(lngRow, 1)
        txtTmp.SelStart = 0
        txtTmp.SelLength = Len(txtTmp.Text)
        txtTmp.ZOrder
        txtTmp.Left = vsBalance.Left + vsBalance.CellLeft
        txtTmp.Top = vsBalance.Top + vsBalance.CellTop + (vsBalance.CellHeight - txtTmp.Height) / 2 - 15
        txtTmp.Width = vsBalance.CellWidth - 30
        
        txtTmp.Visible = True
        txtTmp.SetFocus
    End If
End Sub

Private Sub vsBalance_EnterCell()
    If vsBalance.Col = 0 Then vsBalance.Col = 1
    
    If Not (mbytInState = 0 And chkCancel.Value = 0) Then Exit Sub
    
    If vsBalance.RowData(vsBalance.Row) = 0 Then
        vsBalance.FocusRect = flexFocusLight
    Else
        vsBalance.FocusRect = flexFocusHeavy
    End If
End Sub

Private Sub vsBalance_GotFocus()
    vsBalance_EnterCell
End Sub

Private Sub vsBalance_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
    ElseIf InStr("0123456789", Chr(KeyAscii)) > 0 _
        And vsBalance.RowData(vsBalance.Row) <> 0 _
        And vsBalance.TextMatrix(vsBalance.Row, 0) <> "" _
        And (mbytInState = 0 And chkCancel.Value = 0) Then
        
        txtTmp.Text = Chr(KeyAscii)
        txtTmp.SelStart = Len(txtTmp.Text)
        txtTmp.SelLength = 0
        txtTmp.ZOrder
                    
        txtTmp.Left = vsBalance.Left + vsBalance.CellLeft
        txtTmp.Top = vsBalance.Top + vsBalance.CellTop + (vsBalance.CellHeight - txtTmp.Height) / 2 - 15
        txtTmp.Width = vsBalance.CellWidth - 30
        
        KeyAscii = 0
        
        txtTmp.Visible = True
        txtTmp.SetFocus
    End If
End Sub

Private Sub tbsBill_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtIn_GotFocus()
    Call zlControl.TxtSelAll(txtIn)
End Sub

Private Sub txtIn_KeyPress(KeyAscii As Integer)
'����:�շѻ򻮼�ʱ���뵥��
    Dim lngPre As Long, strPre As String, strNo As String, strNos As String
    Dim intInsure As Integer, i As Long, j As Long
    Dim lng����ID As Long, lng����ID As Long, bln���� As Boolean
    Dim strTmp As String
    Dim objBill As ExpenseBill
    
    If KeyAscii > 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))

    '��һλ����������ĸ,����λ����
    If KeyAscii <> 13 Then
        Call SetNOInputLimit(txtIn, KeyAscii)
    Else
        KeyAscii = 0
        '���뵥��
        txtIn.Text = GetFullNO(txtIn.Text, 13)
        Call zlControl.TxtSelAll(txtIn)
        strNo = txtIn.Text
               
        'a.���ŵ���ģʽ,�����ǰ���ݶ��󼰲�����Ϣ
        If Not cmdAddBill.Enabled Or Not cmdAddBill.Visible Then
            Call ClearFullBill(False)
            
            Set mobjBill = ImportBill(strNo, False, mbytInFun, , False, , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ�۸�ȼ�)
            If mobjBill.NO = "" Then
                MsgBox "��ȡ����ʧ�ܡ�", vbInformation, gstrSysName
                txtIn.SetFocus: Exit Sub
            End If
            
            If InStr(mstrPrivs, "��ʾ������") = 0 Then mobjBill.Pages(mintPage).������ = ""
            '���������Ϣ
            Call ClearmobjBill
        Else
        'b.���ŵ���ģ��,��������,������ǰ�������ݼ����������Ϣ,
        '���ṩ�Ӻ󱸱��е���Ĺ���
            strNos = GetMultiNOs(strNo, , , True)
            For i = 0 To UBound(Split(strNos, ","))
                strNo = Replace(Split(strNos, ",")(i), "'", "")
                
                Set objBill = ImportBill(strNo, False, mbytInFun, , False, , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ�۸�ȼ�)
                
                If objBill.NO = "" Then
                    MsgBox "��ȡ����ʧ�ܡ�", vbInformation, gstrSysName
                    txtIn.SetFocus: Exit Sub
                End If
                
                If i > 0 Or mobjBill.Pages(mintPage).Details.Count > 0 Then
                    Call AddNewBill
                End If
                mintPage = tbsBill.Tabs.Count
                
                '����Ҫ���벡�������Ϣ
                With mobjBill.Pages(mintPage)
                    .NO = "" 'Ҫ����Ա��޸�ʱ������ֱ������ķ���
                    .Key = objBill.Pages(1).Key
                    .���ս�� = objBill.Pages(1).���ս��
                    .��Ԥ���� = objBill.Pages(1).��Ԥ����
                    .�巨 = objBill.Pages(1).�巨
                    .����ͳ�� = objBill.Pages(1).����ͳ��
                    .��������ID = objBill.Pages(1).��������ID
                    If InStr(mstrPrivs, "��ʾ������") > 0 Then .������ = objBill.Pages(1).������
                    .ȫ�Ը� = objBill.Pages(1).ȫ�Ը�
                    .ʵ�ս�� = objBill.Pages(1).ʵ�ս��
                    .�շѽ��� = objBill.Pages(1).�շѽ���
                    .����� = objBill.Pages(1).�����
                    .���Ը� = objBill.Pages(1).���Ը�
                    .Ӧ�ɽ�� = objBill.Pages(1).Ӧ�ɽ��
                    .Ӧ�ս�� = objBill.Pages(1).Ӧ�ս��
                End With
                
                For j = 1 To objBill.Pages(1).Details.Count
                    With objBill.Pages(1).Details(j)
                        mobjBill.Pages(mintPage).Details.Add .�ѱ�, .Detail, .�շ�ϸĿID, .���, .��������, .�շ����, .���㵥λ, .��ҩ����, .����, .����, .���ӱ�־, .ִ�в���ID, .InComes, , .������Ŀ��, .���մ���ID, .���ձ���, .ժҪ
                    End With
                Next
            Next
            tbsBill.Tabs(mintPage).Selected = True  '��������click�¼�,��Ϊmintpage=index
        End If
        
        Call Set�����˿�������(mobjBill.Pages(mintPage).������, mobjBill.Pages(mintPage).��������ID)
        Call LoadAndSeek�ѱ�
        
        'ȡ��һҩƷ��
        For i = 1 To mobjBill.Pages(1).Details.Count
            If InStr(",5,6,7,", mobjBill.Pages(1).Details(i).�շ����) > 0 Then
                mlngFirstID = mobjBill.Pages(1).Details(i).ִ�в���ID
                mstrFirstWin = mobjBill.Pages(1).Details(i).��ҩ����
                Exit For
            End If
        Next
        
        Bill.Active = False
        Bill.Rows = mobjBill.Pages(mintPage).Details.Count + 1
        Call InitBillColumnColor
        
        If IIf(mlngPrePati = 0, mstrPrePati <> mobjBill.����, mlngPrePati <> mobjBill.����ID) Then
            '�²���
            mcurBillʵ�� = 0:  mcurBillӦ�� = 0: mcurBillӦ�� = 0
            mintBillNO = 0: mintMoneyRow = 0
        End If
        
        '�޸�ʱӦ���浱ǰ����Ա������
        mobjBill.����Ա��� = UserInfo.���
        mobjBill.����Ա���� = UserInfo.����
        
        Call CalcMoneys     '��Ϊ�����벡����Ϣ,������Ҫ���ݵ�ǰ�ķѱ�����۸�
        Call ShowDetails
        Call ShowMoney
                        
        txtIn.Text = ""
        'txt����Ӧ��.Visible = False:
        If mbytInState = 0 And mstrInNO <> "" Then txtModi.Text = "": mstrInNO = "": lblӦ��.Caption = "Ӧ��"
        
        'Ҫ����mstrInNO֮��,��Ϊ�Դ����ж��Ƿ��޸ĵ���,�Լӻ�ԭ���
        Call CalcDrugStock
                    
        Bill.Active = True
        If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
        
    End If
End Sub

Private Sub CalcDrugStock(Optional intPage As Integer)
    Dim i As Integer
    Dim strҩ��IDs As String
    '���¼���ÿ��ҩƷ���

     If intPage = 0 Then intPage = mintPage
     
     For i = 1 To mobjBill.Pages(intPage).Details.Count
        With mobjBill.Pages(intPage).Details(i)
            Bill.RowData(i) = Asc(.�շ����) '���⴦��
            
            If InStr(",5,6,7,", .�շ����) > 0 Then
                .Detail.��� = GetStock(.�շ�ϸĿID, .ִ�в���ID)
                If gblnҩ����λ Then .Detail.��� = .Detail.��� / .Detail.ҩ����װ
                If mbytInState = 0 And mstrInNO <> "" Then .Detail.��� = .Detail.��� + .ԭʼ����
                
                Call SetItemRowColor(1, i)  '�����޶���ʾ
            ElseIf .�շ���� = "4" And .Detail.�������� Then
                .Detail.��� = GetStock(.�շ�ϸĿID, .ִ�в���ID, .Detail.����)
                If mbytInState = 0 And mstrInNO <> "" Then .Detail.��� = .Detail.��� + .ԭʼ����
                
                Call SetItemRowColor(1, i) '�����޶���ʾ
            End If
        End With
    Next
End Sub

Private Sub txtInvoice_Change()
    lblFact.Tag = ""
End Sub

Private Sub txtInvoice_LostFocus()
    If Not (mbytInFun = 0 And mbytInState = 0) Then Exit Sub
    If txtInvoice.Text = "" Then
        Call RefreshFact
    End If
End Sub

Private Sub txt����_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txt����.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt����.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt����_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txt����.hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt����_Validate(Cancel As Boolean)
    If Not IsNumeric(txt����.Text) And Trim(txt����.Text) <> "" Then
        cbo���䵥λ.ListIndex = -1: cbo���䵥λ.Visible = False
    ElseIf cbo���䵥λ.Visible = False Then
        cbo���䵥λ.ListIndex = 0: cbo���䵥λ.Visible = True
    End If
    
    If (mbytInFun = 0 Or mbytInFun = 1) And mbytInState = 0 Then mobjBill.���� = Trim(txt����.Text) & IIf(IsNumeric(txt����.Text), cbo���䵥λ.Text, "")
End Sub

Private Sub txtPatient_Change()
    If Me.ActiveControl Is txtPatient Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
        IDKind.SetAutoReadCard (txtPatient.Text = "")
    End If
End Sub

Private Sub txtTmp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtTmp_Validate(Cancel As Boolean)
    Dim curOrig As Currency, curValue As Currency
    Dim curTotal As Currency, arrValue As Variant
    Dim p As Integer, i As Integer
    
    With vsBalance
        If Not IsNumeric(txtTmp.Text) Then
            Cancel = True
            MsgBox "�����˷Ƿ���""" & .TextMatrix(.Row, 0) & """�����", vbInformation, gstrSysName
            Call zlControl.TxtSelAll(txtTmp): Exit Sub
        Else
            '��������������ص�ԭʼ���(�����ʻ�����͸֧ʱ���ж�)
            curOrig = GetMedicareSum(.TextMatrix(.Row, 0), , True) '�ý��㷽ʽ����ԭʼ���ؽ���
            If (.TextMatrix(.Row, 0) <> mstr�����ʻ� Or mcur����͸֧ = 0) _
                And Val(txtTmp.Text) > curOrig And Val(txtTmp.Text) <> 0 And curOrig <> 0 Then
                Cancel = True
                MsgBox "�����""" & .TextMatrix(.Row, 0) & """������ܳ��� " & Format(curOrig, "0.00") & " ��", vbInformation, gstrSysName
                Call zlControl.TxtSelAll(txtTmp): Exit Sub
            End If
            
            '�����ʻ����
            If .TextMatrix(.Row, 0) = mstr�����ʻ� Then
                '������������͸֧���
                If mcur������� - Val(txtTmp.Text) < -1 * mcur����͸֧ Then
                    Cancel = True
                    MsgBox "�ʻ����:" & Format(mcur�������, "0.00") & _
                        IIf(mcur����͸֧ = 0, "", "(" & "����͸֧:" & Format(mcur����͸֧, "0.00") & ")") & _
                        "����Ҫ����Ľ�", vbInformation, gstrSysName
                    Call zlControl.TxtSelAll(txtTmp): Exit Sub
                End If
            End If
            
            '������������ʣ��ɽ�����
            curTotal = GetBillSum - Val(txtԤ�����.Text)
            For p = 1 To mcolBalance.Count
                For i = 0 To UBound(mcolBalance(p))
                    arrValue = Split(mcolBalance(p)(i), ";")
                    If arrValue(0) <> .TextMatrix(.Row, 0) Then
                        curTotal = curTotal - CCur(arrValue(3))
                    End If
                Next
            Next
            If Val(txtTmp.Text) > curTotal Then
                Cancel = True
                MsgBox "��������󣬳����������������:" & Format(curTotal, "0.00") & "��", vbInformation, gstrSysName
                Call zlControl.TxtSelAll(txtTmp): Exit Sub
            End If
            
            .Text = Format(Val(txtTmp.Text), "0.00")
            txtTmp.Text = "": txtTmp.Visible = False
            
            '���޸ĺ�Ľ����䵽���ŵ�����
            '����ԭ�򣺴ӵ���1��ʼ,�Բ�����ԭʼ���ѭ������
            curValue = CCur(.Text)
            For p = 1 To mcolBalance.Count
                If .TextMatrix(.Row, 0) = mstr�����ʻ� And mcur����͸֧ <> 0 Then
                    '����͸֧�ĸ����ʻ�,�Բ���������ʣ��ɽ�����Ϊ׼(���Ƴ�Ԥ��,��Ϊ�Ǻ����)
                    curOrig = GetBillSum(, CLng(p))
                    For i = 0 To UBound(mcolBalance(p))
                        arrValue = Split(mcolBalance(p)(i), ";")
                        If arrValue(0) <> .TextMatrix(.Row, 0) Then
                            curOrig = curOrig - CCur(arrValue(3))
                        End If
                    Next
                Else
                    curOrig = GetMedicareSum(.TextMatrix(.Row, 0), p, True)
                End If
                If curOrig <= curValue Then
                    Call SetBalanceVal(p, .TextMatrix(.Row, 0), curOrig)
                    curValue = curValue - curOrig
                Else
                    Call SetBalanceVal(p, .TextMatrix(.Row, 0), curValue)
                    curValue = 0
                End If
            Next
            
            '���¼���Ӧ�ɣ����(�ֱ�)��:������ϸδ��,ȫ���������¼���
            Call ShowMoney(-1, Not (cmdԤ����.Visible And cmdOK.Enabled))
        End If
    End With
End Sub

Private Sub txtDate_GotFocus()
    zlControl.TxtSelAll txtDate
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And IsDate(txtDate.Text) Then
        mobjBill.����ʱ�� = CDate(txtDate.Text)
        If cmdԤ����.Visible And cmdԤ����.Enabled Then
            cmdԤ����.SetFocus
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub txtDate_LostFocus()
    txtDate.SelLength = 0
    If IsDate(txtDate.Text) Then mobjBill.����ʱ�� = CDate(txtDate.Text)
End Sub

Private Sub cboNO_GotFocus()
    Call zlControl.TxtSelAll(cboNO)

    If (mbytInFun = 0 And mbytInState = 0 And mobjBill.Pages(mintPage).Details.Count = 0) _
        Or chkCancel.Value = 1 Or mbytBilling = 2 Then
        cboNO.Locked = False '�շ�ʱ���յ��ݿ����Ữ�۵���Ҳ���ظ���ȡ
    Else
        cboNO.Locked = True
    End If
    
    '�շ�ʱ�������֤ҽ���������,���ֹ�ٶ�ȡ���۵�
    If mbytInFun = 0 And mbytInState = 0 And mstrYBPati <> "" Then cboNO.Locked = True
End Sub

Private Sub cboNO_KeyPress(KeyAscii As Integer)
    Dim blnRead As Boolean, blnNull As Boolean, rsTmp As ADODB.Recordset
    Dim strOper As String, strNos As String, vDate As Date, intTmp As Integer
    Dim intInsure As Integer, blnHaveExe As Boolean, blnFlagPrint As Boolean
    Dim i As Integer, strErrMsg As String
    
    
    If KeyAscii > 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))

    '��һλ����������ĸ,����λ����
    If KeyAscii <> 13 Then
        Call SetNOInputLimit(cboNO, KeyAscii)
    End If
    
    If KeyAscii = 13 And cboNO.Text <> "" And Not cboNO.Locked Then
        If chkCancel.Value = 1 Then
            If mbytInFun = 0 Then '�շ��˷�
                cboNO.Text = GetFullNO(cboNO.Text, 13)
            ElseIf mbytInFun = 2 Then '��������
                cboNO.Text = GetFullNO(cboNO.Text, 14)
            End If
        ElseIf mbytInFun = 2 Then
            '��˼���
            cboNO.Text = GetFullNO(cboNO.Text, 14)
        ElseIf mbytInFun = 0 Then
            '�����շ�
            cboNO.Text = GetFullNO(cboNO.Text, 13)
        End If
        
        If chkCancel.Value = 1 Then
            '1.�շ�ʱ�Ữ�۵��������
            '2.ҩ�����۲������
            '3.�����������ʻ����¿����롢��˻��۵���Ҫ�ſ������ؼ��
            If mbytInFun <> 2 Or (mbytInFun = 2 And mbytBilling = 0) Then
                '�Ƿ���ת������ݱ���
                If zlDatabase.NOMoved("������ü�¼", cboNO.Text, , IIf(mbytInFun = 2, "2", "1"), Me.Caption) Then
                    If Not ReturnMovedExes(cboNO.Text, IIf(mbytInFun = 2, "2", "1"), Me.Caption) Then cboNO.Text = "": cboNO.SetFocus: Exit Sub
                    mblnNOMoved = False
                End If
            End If
        
            '�����˻���ȫ��˵Ĳ���������
            If mbytInFun = 2 Then
                If Not BillIdentical(cboNO.Text) Then
                    MsgBox "�����а������ݲ�ȫ����˻�ֶ����˵����ݣ����������������ʡ�" & _
                        vbCrLf & "���˻ع��������˳���Ӧ�ĵ������ݣ�Ȼ�������ʡ�", vbInformation, gstrSysName
                    cboNO.Text = "": cboNO.SetFocus: Exit Sub
                End If
            End If
        
            '�����˷�Ȩ���ж�
            If mbytInFun = 0 Then '�շ�
                If Not ReadBillInfo(1, cboNO.Text, 1, strOper, vDate) Then
                    cboNO.Text = "": cboNO.SetFocus: Exit Sub
                End If
                If InStr(mstrPrivs, "���в���Ա") <= 0 Then
                    If UserInfo.���� <> strOper Then
                        MsgBox "��û��""���в���Ա""Ȩ��,���ܶ�" & strOper & "�ĵ��ݽ����˷ѣ�", vbInformation, gstrSysName
                        cboNO.Text = "": cboNO.SetFocus: Exit Sub
                    End If
                End If
                If Not BillOperCheck(2, strOper, vDate, "�˷�", cboNO.Text, , 1) Then
                    cboNO.Text = "": cboNO.SetFocus: Exit Sub
                End If
            ElseIf mbytInFun = 2 Then '����
                If Not ReadBillInfo(1, cboNO.Text, 2, strOper, vDate) Then
                    cboNO.Text = "": cboNO.SetFocus: Exit Sub
                End If
                If InStr(mstrPrivs, "���в���Ա") <= 0 Then
                    If UserInfo.���� <> strOper Then
                        MsgBox "��û��""���в���Ա""Ȩ��,���ܶ�" & strOper & "�ĵ��ݽ������ʣ�", vbInformation, gstrSysName
                        cboNO.Text = "": cboNO.SetFocus: Exit Sub
                    End If
                End If
                
                If Not BillOperCheck(4, strOper, vDate, "����", cboNO.Text, , 2) Then
                    cboNO.Text = "": cboNO.SetFocus: Exit Sub
                End If
            End If
            
            If mbytInFun = 0 Then '�շ��˷�
                '����Ƿ��쳣����
                If ChargeIsErrBill(cboNO.Text) Then
                    If MsgBox("���ݣ�" & cboNO.Text & "���շѵ���Ϊ�쳣�շѵ���,�õ���ֻ�����ϻ������շѣ��Ƿ����?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        cboNO.Text = "": cboNO.SetFocus: Exit Sub
                    End If
                End If
                
                If gbln�˷�����ģʽ Then
                    Set rsTmp = GetApply(cboNO.Text, 1)
                    rsTmp.Filter = "״̬<>2"
                    If rsTmp.RecordCount = 0 Then
                        MsgBox "���ȶԸõ��ݽ����˷����룡", vbInformation, gstrSysName
                        cboNO.Text = "": cboNO.SetFocus: Exit Sub
                    End If
                    If IsNull(rsTmp!�����) Then
                        MsgBox "�õ���δ�����˷���ˣ����ܽ����˷ѣ�", vbInformation, gstrSysName
                        cboNO.Text = "": cboNO.SetFocus: Exit Sub
                    End If
                End If
                
                strNos = GetMultiNOs(cboNO.Text, , mblnNOMoved, True)
                
                If gblnMultiBalance And InStr(strNos, ",") > 0 Then
                    If CheckSingleBalance(strNos) = False Then
                        MsgBox "���ŵ���ʹ�ö��ֽ��㷽ʽģʽ�²����������һ�ŵ����˷ѣ�", vbInformation, gstrSysName
                        cboNO.Text = "": cboNO.SetFocus: Exit Sub
                    End If
                End If
            
                'ҽ������ƥ���ж�(ȷ��ʱ�����ظ��ж�һ��,��Ϊ��Ҫ��ȡ����ҽ������)
                intInsure = ChargeExistInsure(strNos)
                If intInsure > 0 Then
                    '�����˷�Ȩ�޼��
                    If InStr(mstrPrivs, "�����շ�") = 0 Then
                        MsgBox "��û��Ȩ�޶�ҽ�����˵ĵ����˷ѣ�", vbInformation, gstrSysName
                        cboNO.Text = "": cboNO.SetFocus: Exit Sub
                    End If
                    
                    If InStr(strNos, ",") > 0 Then
                        If gclsInsure.GetCapability(support�൥���շѱ���ȫ��, , intInsure) Then
                            MsgBox "ҽ��Ҫ��һ���շѵĶ��ŵ��ݱ��������˷�,��ʹ�ö൥���˷�ģʽ��", vbInformation, gstrSysName
                            cboNO.Text = "": cboNO.SetFocus: Exit Sub
                        End If
                        If gclsInsure.GetCapability(support�൥��һ�ν���, , intInsure) Then
                            MsgBox " ���ŵ���һ�ν��ױ��������˷�,��ʹ�ö൥���˷�ģʽ��", vbInformation, gstrSysName
                            cboNO.Text = "": cboNO.SetFocus: Exit Sub
                        End If
                    End If
                Else
                    '�Ƿ��з�ҽ�����˵��˷�Ȩ��
                    If InStr(mstrPrivs, "�����ҽ������") = 0 Then
                        MsgBox "��û��Ȩ�޶Է�ҽ�����˽����˷Ѳ�����", vbInformation, gstrSysName
                        cboNO.Text = "": cboNO.SetFocus: Exit Sub
                    End If
                End If
                '�˷�
                With mTyDelFee
                    .strNos = strNos
                    Set .rsBlance = GetChargeBalance(strNos)
                End With
                '�����������
                If InStr(1, strNos, ",") > 0 Then
                    With mTyDelFee.rsBlance
                        .Filter = "�Ƿ�ȫ��=1 And �Ƿ�����=0"
                        If .RecordCount <> 0 Then .MoveFirst
                        strErrMsg = ""
                        Do While Not .EOF
                            strErrMsg = strErrMsg & vbCrLf & Nvl(!����) & ":" & Format(Val(Nvl(!������, 0)), "0.00")
                            .MoveNext
                        Loop
                        '����:43734
                        If strErrMsg <> "" Then
                            MsgBox "���������������ײ��ܽ��в����˷ѣ����飡" & vbCrLf & strErrMsg, vbInformation, gstrSysName
                            cboNO.Text = "": cboNO.SetFocus: Exit Sub
                        End If
                    End With
                End If
            End If
            
            '�Ƿ���ִ��
            intTmp = BillCanDelete(cboNO.Text, IIf(mbytInFun = 2, 2, 1), blnHaveExe, , blnFlagPrint)
            If intTmp <> 0 Then
                Select Case intTmp
                    Case 1 '�õ��ݲ�����
                        MsgBox "ָ���ĵ��ݲ����ڣ�", vbInformation, gstrSysName
                    Case 2 '�Ѿ�ȫ����ȫִ��
                        '�շѲ������˷��Զ���ҩ
                        MsgBox "�õ����е���Ŀ�Ѿ�ȫ����ȫִ�У�", vbInformation, gstrSysName
                    Case 3 'δ��ȫִ�в���ʣ������Ϊ0
                        MsgBox "�õ�����δ��ȫִ�в�����Ŀʣ������Ϊ��,û�п���" & IIf(mbytInFun = 2, "����", "�˷�") & "�ķ��ã�", vbInformation, gstrSysName
                End Select
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            ElseIf mbytInFun = 0 And intInsure > 0 And blnHaveExe Then '�շ�ҽ���˷�
                MsgBox "��ҽ���շѵ����а����Ѿ�ִ�е���Ŀ,�����˷ѣ�", vbInformation, gstrSysName
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If
            If blnHaveExe Then
                MsgBox "ע��:�õ������ڴ�����ִ�е���Ŀ����ǰ��ִ�е��ǲ����˷ѡ�", vbInformation, gstrSysName
            End If
            If blnFlagPrint Then
                If MsgBox("ע��:����ҽ���������Ѵ�ӡ���Ƿ�����˷ѣ�", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                    cboNO.Text = "": cboNO.SetFocus: Exit Sub
                End If
            End If
            '�Ƿ��ѽ���
            If mbytInFun = 2 Then
                If HaveBilling(1, cboNO.Text) <> 0 Then
                    Select Case gbytBillOpt
                        Case 0
                        Case 1
                            If MsgBox("�õ����Ѿ�����,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                cboNO.Text = "": cboNO.SetFocus: Exit Sub
                            End If
                        Case 2
                            MsgBox "�õ����Ѿ�����,�������ʣ�", vbInformation, gstrSysName
                            cboNO.Text = "": cboNO.SetFocus: Exit Sub
                    End Select
                End If
            End If
        ElseIf mbytInFun = 2 And mobjBill.Pages(1).Details.Count = 0 Then
            '���ʻ��۵�(�������)
            If Not BillExistMoney("'" & cboNO.Text & "'", 2) Then
                MsgBox "���ݷ����Ѿ�ȫ�����ʻ򵥾ݲ����ڣ�", vbInformation, gstrSysName
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If
        ElseIf mbytInFun = 0 And chkCancel.Value = 0 Then
            '��ȡ���۵��շ�
            If gblnCheckTest Then
                If Not CheckTest(cboNO.Text) Then
                    cboNO.Text = "": cboNO.SetFocus: Exit Sub
                End If
            End If
            
            '����Ƿ�����ȡ�û��۵�
            For i = 1 To tbsBill.Tabs.Count
                If mobjBill.Pages(i).NO = cboNO.Text And i <> mintPage Then
                    MsgBox "���Ż��۵��Ѿ��ڵ� " & i & " �ŵ��������롣", vbInformation, gstrSysName
                    cboNO.Text = mobjBill.Pages(mintPage).NO: cboNO.SetFocus: Exit Sub
                End If
            Next
        End If
        
        
        Call ClearPayInfo
        txtPatient.PasswordChar = ""
        '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
        txtPatient.IMEMode = 0
        '�����޸�ʱ,mstrInNOΪ��ȡ���˷ѵ�,��˵�,���������۵�
        If Not (mbytInFun = 0 And chkCancel.Value = 0) Then mstrInNO = UCase(cboNO.Text)
        
        If chkCancel.Value = 1 Then '��ȡ�˷ѵ�
            blnRead = ReadBill(cboNO.Text, mbytInFun, True)
        ElseIf mbytInFun = 2 Then '��ȡ���ʻ��۵�
            blnRead = ReadBill(cboNO.Text, 2, False, blnNull)
        ElseIf mbytInFun = 0 Then '��ȡ�շѻ��۵�
            blnRead = ReadBill(cboNO.Text, 1, False, blnNull)
        End If
        If blnRead Then
            If chkCancel.Value = 0 Then '�շѻ���ʵĻ��۵�
                Bill.Active = False
                chk�Ӱ�.Enabled = False
                
                '���û��Ȩ�ޣ���ȡ���۵���,ֻ������ҽ������
                If gint������Դ = 1 And mbytInFun = 0 And InStr(mstrPrivs, "�����ҽ������") = 0 Then
                     ClearPatientInfo (True)
                End If
                
                '����ǹҺŲ�����ʱ��������ģʽ,���ȡ���������Ϣ,�Ա��޸�
                If mbytInFun = 0 And txtPatient.Text = "�²���" Then
                    Call GetPatient("-" & mobjBill.����ID)
                End If
                
                '��ʾժҪ
                Call Bill_EnterCell(1, BillCol.��Ŀ)
                
                If mbytInFun = 0 And txtPatient.Text <> "�²���" Then
                    If Not CheckRegisted(mobjBill.����ID) Then
                        Call ClearFullBill(False)
                        Exit Sub
                    End If
            
                    '�Զ����չҺŷ�
                    Call LoadAddedItem(mobjBill.����ID, mobjBill.����)
                    
                    '���۵��շ�ʱ��LED
                    If tbsBill.Tabs.Count = 1 Then Call ShowWelcomeByLed
                End If
                Call ReInitPatiInvoice '97160
                
                '��궨λ
                If mbytInFun = 2 Then
                    cmdOK.SetFocus
                ElseIf txtPatient.Text = "" Or blnNull Then
                    txtPatient.SetFocus
                Else
'                    If txt�ɿ�.Enabled And txt�ɿ�.Visible Then
'                        txt�ɿ�.SetFocus
'                    Else
                    If cmdԤ����.Enabled And cmdԤ����.Visible Then
                        cmdԤ����.SetFocus
                    ElseIf cmdOK.Enabled And cmdOK.Visible Then
                        cmdOK.SetFocus
                    End If
                End If
            Else '��
                Call SetDisible 'cboNO�ڻ�ȡ�����unLock
                If mbytInFun = 0 Then
                    '�����˷�ֻ֧���˷�ָ�����㷽ʽ
                    cbo���㷽ʽ.Enabled = True
                    cbo���㷽ʽ.Locked = False
                End If
                Bill.Active = True
                cmdOK.SetFocus
            End If
        Else
            If Not (mbytInFun = 0 And chkCancel.Value = 0) Then mstrInNO = ""
            cboNO.Text = ""
            If cboNO.Visible And cboNO.Enabled Then cboNO.SetFocus
        End If
    End If
End Sub

Private Sub txt�����_GotFocus()
    zlControl.TxtSelAll txt�����
End Sub

Private Sub txt�˷�ժҪ_Change()
    txt�˷�ժҪ.Tag = ""
End Sub


Private Sub txt�˷�ժҪ_KeyDown(KeyCode As Integer, Shift As Integer)
    'ѡ���˷�ԭ��
    If KeyCode <> vbKeyReturn Then Exit Sub
        
    If Trim(txt�˷�ժҪ.Tag) <> "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If Trim(txt�˷�ժҪ.Text) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If zl_SelectAndNotAddItem(Me, txt�˷�ժҪ, Trim(txt�˷�ժҪ.Text), "�����˷�ԭ��", "�����˷�ԭ��ѡ��", True, True) = False Then
        If zlCommFun.IsCharChinese(Trim(txt�˷�ժҪ.Text)) = False Then
            Exit Sub
        Else
            zlCommFun.PressKey vbKeyTab
        End If
    End If
End Sub
Private Sub txt�˷�ժҪ_GotFocus()
    zlCommFun.OpenIme True
    zlControl.TxtSelAll txt�˷�ժҪ
End Sub
Private Sub txt�˷�ժҪ_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub txtӦ��_Change()
    If mbytInFun <> 0 Then Exit Sub
'        If Val(txt�ɿ�.Text) = 0 Then txt�ɿ�.Text = "0.00": txt�Ҳ�.Text = "0.00": Exit Sub
'        txt�Ҳ�.Text = Format(Val(txt�ɿ�.Text) - Val(txtӦ��.Text), "0.00")
End Sub

Private Sub txtԤ�����_GotFocus()
    Call AutoBultBookFee '�շ��Զ�����������
    zlControl.TxtSelAll txtԤ�����
    txtԤ�����.Tag = txtԤ�����.Text
End Sub

Private Sub txtԤ�����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
    End If
    If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtԤ�����_Validate(Cancel As Boolean)
    Dim curTotal As Currency
        
    curTotal = GetBillSum
    If txtԤ�����.Text = "" Then
        txtԤ�����.Text = "0.00"
    ElseIf Not IsNumeric(txtԤ�����.Text) And txtԤ�����.Text <> "" Then
        MsgBox "��Ч��ֵ��", vbInformation, gstrSysName
        zlControl.TxtSelAll txtԤ�����: Cancel = True: Exit Sub
    ElseIf Val(txtԤ�����.Text) < 0 Then
        MsgBox "Ԥ���������Ϊ����", vbInformation, gstrSysName
        If curTotal < 0 Then
            txtԤ�����.Text = "0.00"
        Else
            txtԤ�����.Text = Format(IIf(curTotal - GetMedicareSum > Val(sta.Panels(Pan.C4Ԥ����Ϣ).Tag), Val(sta.Panels(Pan.C4Ԥ����Ϣ).Tag), curTotal - GetMedicareSum), "0.00")
        End If
        zlControl.TxtSelAll txtԤ�����: Cancel = True: Exit Sub
    ElseIf Val(txtԤ�����.Text) > 0 And curTotal < 0 Then
        MsgBox "����Ӧ�����Ϊ��ʱ����ʹ��Ԥ���", vbInformation, gstrSysName
        txtԤ�����.Text = "0.00"
        zlControl.TxtSelAll txtԤ�����: Cancel = True: Exit Sub
    ElseIf Val(txtԤ�����.Text) > Val(sta.Panels(Pan.C4Ԥ����Ϣ).Tag) Then
        MsgBox "Ԥ��������ܳ������˵�Ԥ�����:" & Format(Val(sta.Panels(Pan.C4Ԥ����Ϣ).Tag), "0.00") & " ��", vbInformation, gstrSysName
        If curTotal < 0 Then
            txtԤ�����.Text = "0.00"
        Else
            txtԤ�����.Text = Format(IIf(curTotal - GetMedicareSum > Val(sta.Panels(Pan.C4Ԥ����Ϣ).Tag), Val(sta.Panels(Pan.C4Ԥ����Ϣ).Tag), curTotal - GetMedicareSum), "0.00")
        End If
        zlControl.TxtSelAll txtԤ�����: Cancel = True: Exit Sub
    ElseIf Val(txtԤ�����.Text) > Format(curTotal - GetMedicareSum, "0.00") And Val(txtԤ�����.Text) <> 0 Then
        MsgBox "Ԥ��������ܴ���Ӧ�����:" & Format(curTotal - GetMedicareSum, "0.00") & " ��", vbInformation, gstrSysName
        If curTotal < 0 Then
            txtԤ�����.Text = "0.00"
        Else
            txtԤ�����.Text = Format(IIf(curTotal - GetMedicareSum > Val(sta.Panels(Pan.C4Ԥ����Ϣ).Tag), Val(sta.Panels(Pan.C4Ԥ����Ϣ).Tag), curTotal - GetMedicareSum), "0.00")
        End If
        zlControl.TxtSelAll txtԤ�����: Cancel = True: Exit Sub
    Else
        txtԤ�����.Text = Format(txtԤ�����.Text, "0.00")
    End If
    
    If Val(txtԤ�����.Tag) = Val(txtԤ�����.Text) Then Exit Sub
    
    '���¼���Ӧ�ɣ����(�ֱ�)��:������ϸδ��,ȫ���������¼���
    Call ShowMoney(-1, Not (cmdԤ����.Visible And cmdOK.Enabled))
End Sub

Private Sub txtInvoice_GotFocus()
    zlControl.TxtSelAll txtInvoice
End Sub

Private Sub txtInvoice_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
    ElseIf Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or InStr("0123456789" & Chr(8), Chr(KeyAscii)) > 0) Then
        KeyAscii = 0
    ElseIf Len(txtInvoice.Text) = txtInvoice.MaxLength And KeyAscii <> 8 And txtInvoice.SelLength <> Len(txtInvoice) Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtModi_LostFocus()
    If mstrInNO <> "" And txtModi.Text <> mstrInNO Then txtModi.Text = mstrInNO
End Sub

Private Sub txt����_Gotfocus()
    Call zlCommFun.OpenIme
    txt����.SelStart = 0
    txt����.SelLength = Len(txt����.Text)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If cbo���䵥λ.Visible = False And IsNumeric(txt����.Text) Then
            Call txt����_Validate(False)
            Call cbo���䵥λ.SetFocus
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
        If Not IsNumeric(txt����.Text) Then Call zlCommFun.PressKey(vbKeyTab)
    Else
        If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Chr(KeyAscii))) > 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtPatient_GotFocus()
    If mbytInFun = 0 Or mbytInFun = 1 Then
        zlControl.TxtSelAll txtPatient
        zlCommFun.OpenIme True
    Else
        zlControl.TxtSelAll txtPatient
    End If
    
    'LED��������
    If mbytInFun = 0 And mbytInState = 0 And gblnLED And Trim(txtPatient.Text) = "" Then
        zl9LedVoice.Speak "#51" '�����������
    End If
    
    If txtPatient.Text = "" And Not txtPatient.Locked Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(True)
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(True)
        IDKind.SetAutoReadCard (True)
    End If
End Sub

Private Sub bill_AfterAddRow(Row As Long)
    Dim i As Long

    If mbytInState = 3 Or (chkCancel.Visible And chkCancel.Value = 1) Then
        Bill.Row = 1: Call zlCommFun.PressKey(vbKeyTab)
        Exit Sub
    End If
    
    With Bill
        '������ʱ,�������ÿ����Ѿ������ĵĿɱ������е���ֵ
        If mbytInState <> 2 Then
            .ColData(BillCol.���) = IIf(gbln�շ����, BillColType.ComboBox, BillColType.UnFocus) '�����,��������ʱ�ᱻ�ı�
            .ColData(BillCol.��Ŀ) = BillColType.CommandButton  '��Ŀ��,��������ʱ�ᱻ�ı�
            .ColData(BillCol.����) = BillColType.UnFocus  '����ȱʡ����(=1),�����Ϊ��ҩʱ,��Ϊ����(4)(��ֵ,һ��ȫ��)
            .ColData(BillCol.����) = BillColType.UnFocus  '����ȱʡ����,����Ŀ���ʱ,��Ϊ����(4)
            .ColData(BillCol.��־) = BillColType.UnFocus  '��־ȱʡ����,��Ϊ����ʱ,��Ϊ��ѡ(-1)
        End If
        
        '����б༭����������ɫ
        .SetColColor BillCol.���, &HE7CFBA
        .SetColColor BillCol.��Ŀ, &HE7CFBA
        .SetColColor BillCol.����, &HE7CFBA
        .SetColColor BillCol.ִ�п���, &HE7CFBA
        .SetColColor BillCol.����, &HE0E0E0
        .SetColColor BillCol.����, &HE0E0E0
        .SetColColor BillCol.��־, &HE0E0E0
        
        .TextMatrix(Row, BillCol.��) = Row
        
        '����ط��ֶ����ò�ִ��
        If Visible And Bill.Active And Row > 0 And .ColData(BillCol.���) <> BillColType.UnFocus And Not mblnNewRow Then
            Call zlCommFun.PressKey(13)
        End If
    End With
End Sub

Private Sub cboSex_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cboSex.ListIndex <> -1 Then mobjBill.�Ա� = Mid(cboSex.Text, InStr(cboSex.Text, "-") + 1)
        Call zlCommFun.PressKey(vbKeyTab)
    End If
    If cboSex.Locked Then Exit Sub
    If SendMessage(cboSex.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 17 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
End Sub

Private Sub cbo�ѱ�_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If cbo�ѱ�.Locked Then Exit Sub
    
    If KeyAscii >= 32 Then
        If cbo�ѱ�.Locked Then Exit Sub
    
        lngIdx = zlControl.CboMatchIndex(cbo�ѱ�.hWnd, KeyAscii)
        If lngIdx = -1 And cbo�ѱ�.ListCount > 0 Then lngIdx = 0
        cbo�ѱ�.ListIndex = lngIdx
        
    ElseIf KeyAscii = 13 Then
        If cbo�ѱ�.ListIndex = -1 Then
            mobjBill.�ѱ� = ""
        Else
             '��ʹ������ͬҲҪ����,��Ϊҽ���鿨���������,Ԥ�������ȷ
            If (mstrYBPati <> "" Or mobjBill.�ѱ� <> zlStr.NeedName(cbo�ѱ�.Text)) Then
                mobjBill.�ѱ� = zlStr.NeedName(cbo�ѱ�.Text)
                If mbytInState = 0 And Not CheckBillsEmpty Then
                    '��Ҫ����Ԥ����
                    If cmdԤ����.Visible Then
                        Call InitBalanceGrid
                        cmdԤ����.TabStop = True
                        cmdOK.Enabled = False
                    End If
                    
                    'ȫ�����¼���۸�
                    Call CalcMoneys
                    Call ShowDetails
                    Call ShowMoney
                End If
            End If
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub cbo��������_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long, lngҽ��ID As Long
    '���˺� ����:27378 ����:2010-01-27 13:35:37
    If KeyAscii <> 13 Then Exit Sub
    If cbo��������.ListIndex <> -1 Then
        zlCommFun.PressKey vbKeyTab: Exit Sub
    End If
    
    If cbo������.ListIndex >= 0 Then lngҽ��ID = cbo������.ItemData(cbo������.ListIndex)
    If mrs�������� Is Nothing Then Call FillDept(mlngDeptID, lngҽ��ID)
    If zlSelectDept(Me, mlngModul, cbo��������, mrs��������, cbo��������.Text) = False Then KeyAscii = 0: Exit Sub
    Exit Sub
'
'    If KeyAscii = 13 Then
'
'        mblnCboClick = False    '��������������б�ѡ��һ�������,��Ҫ�ƿ�,��ʱֻ����click,��������벢�һس�,������click,������Ҫ�ڴ˸�ֵ,�Ա�validate�¼���ǿ�е���click�¼�
'        Call zlCommfun.PressKey(vbKeyTab)
'    ElseIf KeyAscii >= 32 And Not cbo��������.Locked Then
'        lngIdx = zlControl.CboMatchIndex(cbo��������.hWnd, KeyAscii)
'        If lngIdx = -1 And cbo��������.ListCount > 0 Then lngIdx = 0
'        cbo��������.ListIndex = lngIdx
'    End If
End Sub

Private Function isCheck������Exists(ByVal str���� As String, Optional blnLocateItem As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������Ƿ��ڿ����������б���.
    '���:str����-����
    '     blnLocateItem:�Ƿ�ֱ�Ӷ�λ
    '����:
    '����:
    '����:���˺�
    '����:2009-07-20 17:53:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    For i = 0 To cbo������.ListCount - 1
        If zlStr.NeedName(cbo������.List(i)) = str���� Then
            If blnLocateItem Then cbo������.ListIndex = i
            isCheck������Exists = True
            Exit Function
        End If
    Next
End Function

Private Sub cbo������_KeyPress(KeyAscii As Integer)
    Dim i As Long, intIdx As Integer, iCount As Integer
    Dim strText As String, strResult As String, strFilter As String
    Dim rsTemp As ADODB.Recordset
    Dim strAdded As String
    If KeyAscii = 13 Then
        If cbo������.Locked Then
            If Not mblnF2Save Then Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        
        strText = UCase(cbo������.Text)
        If cbo������.ListIndex <> -1 Then
            '�����б�ʱ,�����ı�������������
            If strText <> UCase(cbo������.List(cbo������.ListIndex)) Then Call zlControl.CboSetIndex(cbo������.hWnd, -1)
        End If
        If strText = "" Then
            cbo������.ListIndex = -1
        ElseIf cbo������.ListIndex = -1 Then
            intIdx = -1
            strFilter = IIf(gbln��ʿ, "��Ա����<>''", "��Ա����<>'��ʿ'")
            
            '���˺�:22383
            '�ȸ��Ƽ�¼��
            Set rsTemp = zlDatabase.zlCopyDataStructure(mrs������)
            Dim intInputType As Integer '0-�������ȫ����,1-�������ȫ��ĸ,2-����
            Dim strCompents As String 'ƥ�䴮
            
            strCompents = Replace(gstrLike, "%", "*") & strText & "*"
            
            If IsNumeric(strText) Then
                intInputType = 0
            ElseIf zlCommFun.IsCharAlpha(strText) Then
                intInputType = 1
            Else
                intInputType = 2
            End If
            
            mrs������.Filter = strFilter: iCount = 0
            With mrs������
                If .RecordCount <> 0 Then .MoveFirst
                Do While Not mrs������.EOF
                    Select Case intInputType
                    Case 0  '�������ȫ����
                        '������������,��Ҫ���:
                        '1.�������ֵ���,��Ҫ������:12 ƥ��000012���ֿ�,������������01����01���,��ֱ�Ӷ�λ��01,�򲻶�λ��1��.
                        '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                        '��Ҫ�Ǽ�����������������ȫ��ͬ,��ֱ�ӾͶ�λ��������
                        If Nvl(!���) = strText Then strResult = Nvl(!����): iCount = 0: Exit Do
                        
                        '1.�������ֵ���,��Ҫ������:12 ƥ��000012�������,��Ϊ��������кܶ�:��0012,012,000012��.���������ڴ������,��Ҫ����ѡ������ѡ��
                        If Val(Nvl(!���)) = Val(strText) Then
                            If iCount = 0 Then strResult = Nvl(!����)
                            iCount = iCount + 1
                        End If
                        '2.���������,����Ϊ�Ǳ���,ֻ����ƥ��,��������12ƥ��00001201��120001��
                         If Val(mrs������!���) Like strText & "*" Then
                            If isCheck������Exists(Nvl(!����)) And InStr(strAdded, "," & Nvl(!���) & ",") = 0 Then
                                Call zlDatabase.zlInsertCurrRowData(mrs������, rsTemp)
                                strAdded = strAdded & "," & Nvl(!���) & ","
                            End If
                         End If
                    Case 1  '�������ȫ��ĸ
                        '����:
                        ' 1.����ļ������,��ֱ�Ӷ�λ
                        ' 2.���ݲ�����ƥ����ͬ����
                        
                        '1.����ļ������,��ֱ�Ӷ�λ
                        If Trim(Nvl(!����)) = strText Then
                            If iCount = 0 Then strResult = Nvl(!����)   '���ܴ��ڶ����ͬ����
                            iCount = iCount + 1
                        End If
                        '2.���ݲ�����ƥ����ͬ����
                        If Trim(Nvl(!����)) Like strCompents Then
                            If isCheck������Exists(Nvl(!����)) And InStr(strAdded, "," & Nvl(!���) & ",") = 0 Then
                                Call zlDatabase.zlInsertCurrRowData(mrs������, rsTemp)
                                strAdded = strAdded & "," & Nvl(!���) & ","
                            End If
                        End If
                    Case Else  ' 2-����
                        '����:���ܴ��ں��ֵ����,����������N001���������ZYK01�������
                        '1.����\�������,ֱ�Ӷ�λ
                        '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                        
                        '1.����\�������,ֱ�Ӷ�λ
                        If Trim(!���) = strText Or Trim(!����) = strText Or Trim(!����) = strText Then
                            If iCount = 0 Then strResult = Nvl(!����)   '���ܴ��ڶ����ͬ�Ķ��
                            iCount = iCount + 1
                        End If
                        '2.������������� ���ݲ�����ƥ����(������ֻ����ƥ��)
                        If Trim(!���) Like strText & "*" Or Trim(Nvl(!����)) Like strCompents Or Trim(Nvl(!����)) Like strCompents Then
                            If isCheck������Exists(Nvl(!����)) And InStr(strAdded, "," & Nvl(!���) & ",") = 0 Then
                                Call zlDatabase.zlInsertCurrRowData(mrs������, rsTemp)
                                strAdded = strAdded & "," & Nvl(!���) & ","
                            End If
                        End If
                    End Select
                    mrs������.MoveNext
                Loop
            End With
            If iCount > 1 Then strResult = ""
            If strResult = "" And rsTemp.RecordCount = 1 Then strResult = Nvl(rsTemp!����)
            '���˺�:ֱ�Ӷ�λ
            If strResult <> "" Then
                rsTemp.Close: Set rsTemp = Nothing
                If isCheck������Exists(strResult, True) Then zlCommFun.PressKey vbKeyTab
                Exit Sub
            End If
            
            
            '��Ҫ����Ƿ��ж������������ļ�¼
            If rsTemp.RecordCount <> 0 Then
                '�Ȱ�ĳ�ַ�ʽ��������
                Select Case intInputType
                Case 0 '����ȫ����
                    rsTemp.Sort = "���"
                Case 1 '����ȫƴ��
                    rsTemp.Sort = "����"
                Case Else
                    '����ѡ������
                    If gbyt��������ʾ = 1 Then '����
                        rsTemp.Sort = "����"
                    Else
                        rsTemp.Sort = "���"
                    End If
                End Select
                '����ѡ����
                Dim rsReturn As ADODB.Recordset
                If zlDatabase.zlShowListSelect(Me, glngSys, mlngModul, cbo������, rsTemp, True, "", "ȱʡ,ְ��,���ȼ���", rsReturn) Then
                    If cbo������.Enabled Then cbo������.SetFocus
                    If Not rsReturn Is Nothing Then
                        If rsReturn.RecordCount <> 0 Then
                            '���ж�λ
                            If isCheck������Exists(Nvl(rsReturn!����), True) Then
                                'zlCommFun.PressKey vbKeyTab
                            End If
                        End If
                    End If
                End If
            Else
                'δ�ҵ�
                rsTemp.Close: Set rsTemp = Nothing
                KeyAscii = 0: zlControl.TxtSelAll cbo������: Exit Sub
            End If
            rsTemp.Close: Set rsTemp = Nothing
            
        ElseIf Not mblnDrop Then
            '�س���꾭��
            Call cbo������_Click
            If Not mblnF2Save Then Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        If cbo������.ListIndex = -1 Then
            cbo������.Text = ""
            mobjBill.Pages(mintPage).������ = ""
            lblDuty.Caption = ""
            If gbyt����ҽ�� = 0 Or gbln�����俪���� Then Exit Sub
        Else
            mobjBill.Pages(mintPage).������ = zlStr.NeedName(cbo������.Text)
            If intIdx <> -1 And mblnDrop Then
                '�����س�-ǿ�м���Click
                Call cbo������_Click
            ElseIf intIdx <> cbo������.ListIndex And intIdx <> -1 Then
                '������ѡ��-�Զ�����Click
                cbo������.SetFocus
                If Not mblnF2Save Then Call zlCommFun.PressKey(vbKeyF4)
                Exit Sub
            ElseIf intIdx <> -1 Then
                'һ��������-ǿ�м���Click
                Call cbo������_Click
            End If
        End If
        If Not mblnF2Save Then Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub ShowCHRecipe()
'���ܣ�������ҩ�䷽���빦��
    Dim objDetails As BillDetails
    Dim str��̬�ѱ� As String, lng���˿���ID As Long
    Dim int��� As Integer, i As Long
    
    If Not (Bill.Active And mbytInState = 0) Then Exit Sub
    
    '����Ƿ��з���ҩ
    For i = 1 To mobjBill.Pages(mintPage).Details.Count
        If mobjBill.Pages(mintPage).Details(i).�շ���� <> "7" _
            And Not mobjBill.Pages(mintPage).Details(i).������ Then
            Call MsgBox("�ڵ�ǰ�����д��ڲ����в�ҩ���շ���Ŀ����ɾ�����в�ҩ�շ���Ŀ��,�ٽ����䷽!", vbInformation + vbDefaultButton1, gstrSysName)
             
            If cmd�䷽.Enabled And cmd�䷽.Visible Then cmd�䷽.SetFocus
            Exit Sub
        End If
    Next
    
    '���˿��һ򿪵�����ID
    lng���˿���ID = mobjBill.����ID
    If lng���˿���ID = 0 Then lng���˿���ID = Get��������ID
        
    '��̬�ѱ�
    If glngSys Like "8??" Or mbytInFun = 2 Then
        str��̬�ѱ� = zlStr.NeedName(cbo�ѱ�.Text)
    Else
        str��̬�ѱ� = zlStr.TrimEx(zlStr.NeedName(cbo�ѱ�.Text) & "," & lbl��̬�ѱ�.Tag, ",")
    End If
    
    '���ô���
    Set objDetails = frmCHRecipe.ShowMe(Me, mstrPrivs, mlngModul, mbytInFun, mbytBilling, Original.ʵ�պϼ�, mobjBill.����ID, lng���˿���ID, Get��������ID, _
        IIf(mlng��ҩ�� = 0, glng��ҩ��, mlng��ҩ��), mobjBill.Pages(mintPage).Details, zlStr.NeedName(cbo�ѱ�.Text), str��̬�ѱ�, _
         IIf(mstrYBPati <> "", mintInsure, 0), chk�Ӱ�.Value = 1, mobjBill.Pages(mintPage).�巨, mrsWarn, mcolStock1, zl��ȡ��ҩ��̬(mintPage, Bill.Row, True))
    If Not objDetails Is Nothing Then
        '���ԭ�����е��в�ҩ
        For i = mobjBill.Pages(mintPage).Details.Count To 1 Step -1
            If mobjBill.Pages(mintPage).Details(i).�շ���� = "7" Then
                mobjBill.Pages(mintPage).Details.Remove i
            End If
        Next
        '��ӱ༭����в�ҩ
        For i = 1 To objDetails.Count
            With objDetails(i)
                int��� = mobjBill.Pages(mintPage).Details.Count + 1
                Call mobjBill.Pages(mintPage).Details.Add(.�ѱ�, .Detail, .�շ�ϸĿID, int���, .��������, _
                    .�շ����, .���㵥λ, .��ҩ����, .����, .����, .���ӱ�־, .ִ�в���ID, _
                    .InComes, "", .������Ŀ��, .���մ���ID, .���ձ���, .ժҪ, .ԭʼ����, .ԭʼִ�в���ID)
            End With
        Next
        
        '������ҩ�巨
        mobjBill.Pages(mintPage).�巨 = frmCHRecipe.mstr�巨
        'ˢ�µ�ǰ�����е���ʾ
        Call ClearBillRows
        Bill.Rows = mobjBill.Pages(mintPage).Details.Count + 1
        
        Call InitBillColumnColor
        
        '�����¼���֮ǰ���
        If cmdԤ����.Visible Then
            Call InitBalanceGrid
            cmdԤ����.TabStop = True
            cmdOK.Enabled = False
        End If

        Call ShowDetails
        Call ShowMoney(mintPage)
        Call SetColNum
                
        Call CalcDrugStock
        Call SetBill�в�ҩEditEnabled
        
        Bill.Col = BillCol.��Ŀ: Bill.CmdVisible = False  '��Ȼ��λ����
        If cmdԤ����.Enabled And cmdԤ����.Visible Then
            cmdԤ����.SetFocus
'        ElseIf txt�ɿ�.Enabled And txt�ɿ�.Visible Then
'            txt�ɿ�.SetFocus
        ElseIf cmdOK.Enabled And cmdOK.Visible Then
            cmdOK.SetFocus
        End If
    Else
        Bill.SetFocus
    End If
End Sub

Private Sub ApportionMultiBalance(ByVal strBalance As String, ByVal curError As Currency)
    Dim i As Long, j As Long
    Dim cur��ǰ���� As Currency, cur��ǰδ�� As Currency, arrPay As Variant
    Dim varData As Variant, str��֧Ʊ(0 To 3) As String, blnӦ���� As Boolean
    Dim bln������֧Ʊ As Boolean
    
    With mobjBill
        For i = 1 To mobjBill.Pages.Count
            .Pages(i).Ӧ�ɽ�� = 0      '����ģʽ�µ�Ӧ��ʵ����û��ʹ��,��������ĳ����������ж�,���,�ۼƵ�Ӧ�ɿ�����С��
            .Pages(i).�շѽ��� = ""
            If i = .Pages.Count Then
                .Pages(i).����� = curError
            Else
                .Pages(i).����� = 0
            End If
        Next
        '������˳���̯(ҽ����Ԥ��������֮ǰ��̯,�������Ԥ����,�������̯��Ϣ)
        arrPay = Split(strBalance, "||")
        ':33722
        bln������֧Ʊ = False
        For j = 0 To UBound(arrPay)
            varData = Split(arrPay(j), "|")
            mrs���㷽ʽ.Filter = "����='" & varData(0) & "'"
            If Not mrs���㷽ʽ.EOF Then
                blnӦ���� = Val(Nvl(mrs���㷽ʽ!Ӧ����)) = 1
                If blnӦ���� Then
                    str��֧Ʊ(0) = varData(0)
                    str��֧Ʊ(1) = varData(1)
                    str��֧Ʊ(2) = varData(2)
                    str��֧Ʊ(3) = varData(3)
                    bln������֧Ʊ = True
                    Exit For
                End If
            End If
        Next
        
        For j = 0 To UBound(arrPay)
            varData = Split(arrPay(j), "|")
            cur��ǰ���� = varData(1) '���㷽ʽ|������|�������|ժҪ||......
            ':33722
            If str��֧Ʊ(0) = varData(0) And bln������֧Ʊ Then
                   blnӦ���� = True
            Else
                blnӦ���� = False
            End If
            
            If blnӦ���� = False Then
                For i = 1 To mobjBill.Pages.Count
                    cur��ǰδ�� = .Pages(i).ʵ�ս�� + .Pages(i).����� - .Pages(i).���ս�� - .Pages(i).��Ԥ���� - .Pages(i).Ӧ�ɽ�� - .Pages(i).���ѿ�ˢ����
                                            
                    If cur��ǰδ�� > 0 Then
                        If cur��ǰδ�� <= cur��ǰ���� Then
                            '֧Ʊ�Ĵ���:33722
                            '���ܴ���Ӧ��֧Ʊ�������,���ܵ������һ�ŵ���,֧Ʊ����δ����������
                            '��ʱ,�����²���ֱ�ӷ�������һ�ŵ���
                            If i = mobjBill.Pages.Count And varData(0) Like "*֧Ʊ*" And bln������֧Ʊ Then
                                  .Pages(i).�շѽ��� = .Pages(i).�շѽ��� & "||" & varData(0) & "|" & cur��ǰ���� & "|" & varData(2) & "|" & varData(3)
                                  .Pages(i).Ӧ�ɽ�� = .Pages(i).Ӧ�ɽ�� + cur��ǰ����
                                  cur��ǰ���� = 0: Exit For
                            Else
                            .Pages(i).�շѽ��� = .Pages(i).�շѽ��� & "||" & _
                                               varData(0) & "|" & cur��ǰδ�� & "|" & varData(2) & "|" & varData(3)
                            .Pages(i).Ӧ�ɽ�� = .Pages(i).Ӧ�ɽ�� + cur��ǰδ��
                            cur��ǰ���� = cur��ǰ���� - cur��ǰδ��
                            End If
                        Else
                            .Pages(i).�շѽ��� = .Pages(i).�շѽ��� & "||" & _
                                               varData(0) & "|" & cur��ǰ���� & "|" & varData(2) & "|" & varData(3)
                            .Pages(i).Ӧ�ɽ�� = .Pages(i).Ӧ�ɽ�� + cur��ǰ����
                            cur��ǰ���� = 0
                        End If
                        If cur��ǰ���� = 0 Then Exit For
                    End If
                Next
            End If
        Next
        
        If str��֧Ʊ(0) <> "" And bln������֧Ʊ Then
            '��֧Ʊ����,ֻ�ܷ������һ��
            .Pages(mobjBill.Pages.Count).�շѽ��� = .Pages(mobjBill.Pages.Count).�շѽ��� & "||" & _
                                               str��֧Ʊ(0) & "|" & Val(str��֧Ʊ(1)) & "|" & str��֧Ʊ(2) & "|" & str��֧Ʊ(3)
             .Pages(mobjBill.Pages.Count).Ӧ�ɽ�� = .Pages(mobjBill.Pages.Count).Ӧ�ɽ�� + RoundEx(Val(str��֧Ʊ(1)), 2)
        End If
        For i = 1 To mobjBill.Pages.Count
            If Mid(.Pages(i).�շѽ���, 1, 2) = "||" Then .Pages(i).�շѽ��� = Mid(.Pages(i).�շѽ���, 3)
        Next
    End With
    mrs���㷽ʽ.Filter = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    '������Shift=-1����ʾ�ǳ���ǿ���ڵ���
    Select Case KeyCode
        Case vbKeyF1  '����
            Select Case mbytInFun
                Case 0
                    ShowHelp App.ProductName, Me.hWnd, Me.Name & "2"
                Case 1
                    ShowHelp App.ProductName, Me.hWnd, Me.Name & "1"
                Case 2
                    ShowHelp App.ProductName, Me.hWnd, Me.Name & "3"
            End Select
        Case vbKeyF2
            If Shift = vbCtrlMask Then
                If mbytInFun = 0 And mbytInState = 0 And mstrInNO = "" And gbytAutoSplitBill > 0 Then
                    Call AutoSplitBill
                End If
            Else
                mblnF2Save = True
                    If ActiveControl Is txtPatient And mbytInFun <> 1 Then
                        Call txtPatient_LostFocus
                        Call txtPatient_Validate(False)
                        Me.Refresh
                    End If
                    If ActiveControl Is cbo������ Then Call cbo������_KeyPress(vbKeyReturn)
                mblnF2Save = False
                If cmdOK.Enabled And cmdOK.Visible Then
                    Call cmdOK.SetFocus
                    Call cmdOK_Click
                End If
            End If
        Case vbKeyF3 '�Һ�
            If cmdRegist.Visible And cmdRegist.Enabled Then
                cmdRegist.SetFocus
                Call cmdRegist_Click
            End If
        Case vbKeyUp
'            '���˺�:27498 20100117
'            If Me.ActiveControl Is txtPatient Then
'                Call IDKind.Locale(-1)
'                'IDKind.IDKind = IIf(IDKind.IDKind = 0, UBound(Split(IDKind.IDKindStr, ";")), IDKind.IDKind - 1)
'            End If
        Case vbKeyDown
'            '���˺�:27498 20100117
'            If Me.ActiveControl Is txtPatient Then
'                Call IDKind.Locale
'                'IDKind.IDKind = IIf(IDKind.IDKind = UBound(Split(IDKind.IDKindStr, ";")), 0, IDKind.IDKind + 1)
'            End If
        Case vbKeyF4 '���ַ�ʽ����
            If Shift = vbCtrlMask Then
                If IDKind.Enabled And txtPatient.Locked = False And txtPatient.Enabled Then
                    Dim intIndex As Integer
                    intIndex = IDKind.GetKindIndex("IC����")
                    If intIndex <= 0 Then Exit Sub
                    IDKind.IDKind = intIndex: Call IDKind_Click(IDKind.GetCurCard)
                End If
            End If
        Case vbKeyF5
            If cmdԤ����.Visible And cmdԤ����.Enabled Then cmdԤ����.SetFocus: cmdԤ����_Click
        Case vbKeyF6 '��λ�����������
            If Me.ActiveControl Is txtPatient And txtPatient.Enabled And mstrYBPati = "" Then   '��ȡ���۵��������������������
                '70143:������,2014-3-3,סԺ����ҽ����֤
                If mbytInFun = 0 And mbytInState = 0 And (gint������Դ = 1 Or gint������Դ = 2) Then
                    If chkCancel.Value = 0 And InStr(mstrPrivs, "�����շ�") > 0 Then
                        Dim lngCur����ID As Long
                        If mrsInfo.State = 1 Then
                            If txtPatient.Text = mrsInfo!���� Then lngCur����ID = mrsInfo!����ID
                        Else
                            If txtPatient.Text = mobjBill.���� Then lngCur����ID = mobjBill.����ID  '����:25486
                        End If
                        Call MCPatientProcess(lngCur����ID)
                    End If
                End If
            Else
                If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
            End If
        Case vbKeyF7 '�л����뷨
            If Shift = vbCtrlMask Then
                If sta.Panels("PatiSource").Visible Then
                    Call sta_PanelClick(sta.Panels("PatiSource"))
                End If
            Else
                If Not gbln�����л� Then Exit Sub
                If sta.Panels("WB").Visible And sta.Panels("PY").Visible Then
                    If sta.Panels("WB").Bevel = sbrRaised Then
                        Call sta_PanelClick(sta.Panels("WB"))
                    Else
                        Call sta_PanelClick(sta.Panels("PY"))
                    End If
                End If
            End If
        Case vbKeyF8 '��(�Զ������¼�)
            If mbytInFun = 1 Then
                cmdCancel.SetFocus
                Call cmdCancel_Click
            Else
                If chkCancel.Visible And chkCancel.Enabled Then
                    chkCancel.Value = IIf(chkCancel.Value = 1, 0, 1)
                ElseIf cmdDelete.Visible And cmdDelete.Enabled Then
                    cmdDelete.SetFocus: Call cmdDelete_Click
                End If
            End If
        Case vbKeyF9 '��λ�����ݺ������
            If cboNO.Enabled And cboNO.Visible Then cboNO.SetFocus
        Case vbKeyF10 '���￨����
            If cmdIDCard.Visible And cmdIDCard.Enabled Then cmdIDCard.SetFocus: cmdIDCard_Click
        Case vbKeyF11
            If cmd�䷽.Enabled And cmd�䷽.Visible Then cmd�䷽.SetFocus: Call cmd�䷽_Click
        Case vbKeyF12
            If Shift = vbCtrlMask Then
''                'ǿ����LED����,(�ϼ�)
''                If mbytInFun = 0 And gblnLED And mbytInState = 0 _
''                    And txt�ɿ�.Enabled And txt�ɿ�.Visible And CCur(txt�ϼ�.Text) <> 0 Then
''                    mblnHotKey = True: txt�ɿ�.SetFocus
''                    If ActiveControl Is txt�ɿ� Then AutoBultBookFee
''                End If
            ElseIf Shift = vbAltMask Then
                Call sta_PanelClick(sta.Panels("Drugstore"))
            Else
                '����:27939
                If Me.ActiveControl Is txtPatient Then
                    Call txtPatient_Validate(False)
                End If
                '���ӵ���
                If cmdAddBill.Enabled And cmdAddBill.Visible Then cmdAddBill.SetFocus: Call cmdAddBill_Click
            End If
        Case vbKeyS
            '����Ϊ���۵�
            If Shift = vbCtrlMask Then
                If CheckSaveMultiPrice Then
                    Call mnuFileSavePrice_Click
                Else
                    MsgBox "�����շ�ʱ������Ϊ���۵�." & vbCrLf & "����Ƕ��ŵ����շ�,Ҫ�󲻺�����ĵ���", vbInformation, gstrSysName
                End If
            End If
        Case vbKeyA, vbKeyR
            'ȫѡ��ȫ��
            If Shift = vbCtrlMask Then
                If KeyCode = vbKeyA Then
                    Call SelALLRow
                ElseIf KeyCode = vbKeyR Then
                    Call ClearALLRow
                End If
            End If
        Case vbKeyD
            If Shift = vbCtrlMask Then
                If sta.Panels(Pan.C4Ԥ����Ϣ).Visible And mrsInfo.State = 1 Then
                    Call ShowDeposit(mrsInfo!����ID)
                End If
            End If
        Case vbKeyF 'Ctrl+F��λ�ɿ������
'            If Shift = vbCtrlMask Then
'                If txt�ɿ�.Enabled And txt�ɿ�.Visible Then txt�ɿ�.SetFocus
'            End If
        Case vbKeyQ
            If Shift = vbCtrlMask Then Call LocateNewRow
        Case vbKeyEscape
            If Bill.TxtVisible Then
                Bill.Text = "": Bill.TxtVisible = False
                Bill.SetFocus
            ElseIf txtTmp.Visible Then
                txtTmp.Visible = False
                If vsBalance.Enabled Then vsBalance.SetFocus
            Else
                cmdCancel.SetFocus: Call cmdCancel_Click
            End If
        Case 191 '"?"������
            If Shift = vbAltMask Then
                Call sta_PanelClick(sta.Panels("Calc"))
            End If
    End Select
End Sub

Private Sub SetMoneyList()
'����:���ݵ�ǰ������Ŀ�����������п�
    Dim lngW As Long
    lngW = mshMoney.Width - 75
    If mshMoney.Rows > mshMoney.Height / mshMoney.RowHeight(0) Then
        lngW = lngW - 250
    End If
    
    mshMoney.ColWidth(0) = 600
    
    lngW = lngW - mshMoney.ColWidth(0)
    
    If mbytInFun = 1 Then
        mshMoney.ColWidth(1) = lngW * 0.4
        mshMoney.ColWidth(2) = lngW * 0.3
        mshMoney.ColWidth(3) = lngW * 0.3
    Else
        mshMoney.ColWidth(1) = lngW * 0.45
        mshMoney.ColWidth(2) = lngW * 0.55
        mshMoney.ColWidth(3) = 0
    End If
    
    mshMoney.ColAlignment(0) = 4
    mshMoney.ColAlignment(1) = 1
    mshMoney.ColAlignment(2) = 7
    mshMoney.ColAlignment(3) = 7
    
    mshMoney.TextMatrix(0, 0) = "���"
    mshMoney.TextMatrix(0, 1) = "��Ŀ"
    mshMoney.TextMatrix(0, 2) = "���"
    mshMoney.TextMatrix(0, 3) = "�ϼ�"
    mshMoney.Row = 0
    mshMoney.Col = 0: mshMoney.CellAlignment = 4
    mshMoney.Col = 1: mshMoney.CellAlignment = 4
    mshMoney.Col = 2: mshMoney.CellAlignment = 4
    mshMoney.Col = 3: mshMoney.CellAlignment = 4
    
    mshMoney.MergeCol(0) = True
End Sub

Private Function InitData() As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim i As Long, strSQL As String
    Dim Curdate As Date     '��������ǰʱ��
    
    On Error GoTo errH
        
    '��ʼ��������Ϣ����
    Set mrsInfo = New ADODB.Recordset
    '�鿴ʱ,��֧�����֤ʶ��,�޸�ʱҪ֧��,��Ϊ�޸ĺ���ܼ����µ��շ�
    If mbytInState = 0 Then
        Set mobjIDCard = New clsIDCard
        Set mobjICCard = New clsICCard
        Call mobjIDCard.SetParent(Me.hWnd)
        Call mobjICCard.SetParent(Me.hWnd)
        Set mobjICCard.gcnOracle = gcnOracle
    End If
    
    
    '���˺�:���㿨��һЩ����
    Call initCardSquareData
    
    '���䵥λ
    cbo���䵥λ.AddItem "��"
    cbo���䵥λ.AddItem "��"
    cbo���䵥λ.AddItem "��"
    cbo���䵥λ.ListIndex = 0
    
    
    '------------------������ȡ------------------
    
    '��ѡ�Ա�,ҽ�Ƹ��ʽ,���㷽ʽ
    strSQL = " Select '�Ա�' as ���,����,����,����,Nvl(ȱʡ��־,0) as ȱʡ From �Ա� Union All " & _
             " Select 'ҽ�Ƹ��ʽ' as ���,����,����,����,Nvl(ȱʡ��־,0) as ȱʡ From ҽ�Ƹ��ʽ "
    
    strSQL = strSQL & " Order by ���,����"
    Set rsTmp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    
    '1.�Ա�
    rsTmp.Filter = "���='�Ա�'"
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cboSex.AddItem rsTmp!���� & "-" & rsTmp!����
            If rsTmp!ȱʡ = 1 Then cboSex.ListIndex = cboSex.NewIndex
            rsTmp.MoveNext
        Next
    End If
    '2.ҽ�Ƹ��ʽ
    rsTmp.Filter = "���='ҽ�Ƹ��ʽ'"
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cboҽ�Ƹ���.AddItem rsTmp!���� & "-" & rsTmp!����
            If rsTmp!ȱʡ = 1 Then
                cboҽ�Ƹ���.ListIndex = cboҽ�Ƹ���.NewIndex
                mstr���ʽ = rsTmp!����
            End If
            rsTmp.MoveNext
        Next
    End If
    
    
    strSQL = " Select '����ְ��' As ���,count(ҩ��ID) As num From ҩƷ���� Where ����ְ��<>'00' Union All " & _
             " Select '��������' As ���,count(ҩ��ID) As num From ҩƷ���� Where ��������>0     Union All " & _
             " Select '�����޶�' As ���,Count(�ⷿID) As num From ҩƷ�����޶� Where ����>0 Or ����>0"
    Set rsTmp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    
    rsTmp.Filter = "���='����ְ��'"
    If Not rsTmp.EOF Then mbln����ְ���� = (rsTmp!Num > 0)
    
    rsTmp.Filter = "���='��������'"
    If Not rsTmp.EOF Then mbln����������� = (rsTmp!Num > 0)
    
    rsTmp.Filter = "���='�����޶�'"
    If Not rsTmp.EOF Then mbln�����޶��� = (rsTmp!Num > 0)
    
    '------------------������ȡ------------------
    
    
    
    '��ȡ��ҩ������
    Call ReadABCNum(mstrPrivs)
    
    '��ͬҩ��ҩƷ�����鷽ʽ(��������ҩ��,��Ϊ����¼��סԺ����)
    Set mcolStock1 = GetStockCheck(0)
    Set mcolStock2 = GetStockCheck(1)
    
    '���ʷ��౨��
    If mbytInFun = 2 And mbytInState = 0 Then Set mrsWarn = GetUnitWarn("", "0")
    
        
    '���㷽ʽ:1-�ֽ���㷽ʽ,2-������ҽ������,3-ҽ�������ʻ�,4-ҽ������ͳ��
    If mbytInFun = 0 Then
        Set mrs���㷽ʽ = Get���㷽ʽ("�շ�")
        If Not mrs���㷽ʽ.EOF Then
            For i = 1 To mrs���㷽ʽ.RecordCount
                If Val(Nvl(mrs���㷽ʽ!Ӧ����)) = 1 Then
                    '����:33722,������Ӧ��������
                    mstrӦ������㷽ʽ = Nvl(mrs���㷽ʽ!����)
                Else
                    '�Ƿ��и����ʻ�
                    If mrs���㷽ʽ!���� = 3 And mstr�����ʻ� = "" Then
                        mstr�����ʻ� = mrs���㷽ʽ!����
                    End If
                    'ֻ�����ҽ���ʹ��տ�Ľ��㷽ʽ��ѡ��
                    If InStr(",1,2,7,", "," & mrs���㷽ʽ!���� & ",") > 0 Then
                        cbo���㷽ʽ.AddItem mrs���㷽ʽ!���� & "-" & mrs���㷽ʽ!����
                        cbo���㷽ʽ.ItemData(cbo���㷽ʽ.NewIndex) = mrs���㷽ʽ!����
                        
                        If mrs���㷽ʽ!���� = gstr���㷽ʽ Then
                            cbo���㷽ʽ.ListIndex = cbo���㷽ʽ.NewIndex
                        End If
                        
                        If mrs���㷽ʽ!ȱʡ = 1 And cbo���㷽ʽ.ListIndex = -1 Then
                            cbo���㷽ʽ.ListIndex = cbo���㷽ʽ.NewIndex
                        End If
                    End If
                End If
                mrs���㷽ʽ.MoveNext
            Next
        End If
        If cbo���㷽ʽ.ListCount = 0 Then   'ȱʡֵ����NewBill���ٴ��趨
            MsgBox "�շѳ���û�п��õĽ��㷽ʽ�����ȵ����㷽ʽ���������á�", vbInformation, gstrSysName
            Exit Function
        End If
        
        If mbytInState = 0 Or mbytInState = 3 Then
            Set mrsOneCard = GetOneCard
            mblnOneCard = mrsOneCard.RecordCount > 0
        End If
    End If
    
    
    '�ѱ�,Ĭ����ʾ���������п��ҵ�
    Call Load�ѱ�(cbo�ѱ�, 0, mbytInFun = 2, mrs�ѱ�)
    mrs�ѱ�.Filter = ""
    If mrs�ѱ�.RecordCount = 0 Then
        MsgBox "û����Ч�ѱ����ã����ȵ��ѱ�����н������ã�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '��ȱʡ�����˺Ϳ�������
    Call FillDept(mlngDeptID)
    If cbo��������.ListCount = 0 Then
        MsgBox "û�п��õĿ�������,���õĿ����������������¹���:" & vbCrLf & _
               "    1.��������Ϊ����" & IIf(mbytInFun = 1, "�����������ơ���顢����", "") & vbCrLf & _
               "    2.���߲�������Ϊ�ٴ�,���Ҳ��ŷ����������סԺ��������������(������ԴΪ���ﲡ��)���������סԺ(������ԴΪסԺ����).", vbInformation, gstrSysName
        Exit Function
    End If
    zlControl.CboSetWidth cbo��������.hWnd, 2500
    Call FillDoctor
    If cbo������.ListCount = 0 Then
        MsgBox "û�п��õĿ�����,���õĿ��������������¹���:" & vbCrLf & _
               "    1.��Ա����Ϊҽ����ʿ," & vbCrLf & _
               "    2.����,��Ա���ڲ�������Ϊ�ٴ�" & vbCrLf & _
               "    3.����,��Ա���ڲ��ŷ����������סԺ��������������(������ԴΪ���ﲡ��)���������סԺ(������ԴΪסԺ����)." & vbCrLf & _
               "    4.��ʿ�Ƿ�������Ϊ���ÿ��������������¹���:" & vbCrLf & _
               "      ���ز�������������Ϊ��ʿ,���ұ��ز����Ŀ����շ�����������,����,����", vbInformation, gstrSysName
        Exit Function
    End If
        
    
    'ִ�в���
    Set mrsUnit = GetDepartments("", gint������Դ & ",3")
    If mrsUnit.EOF Then
        MsgBox "û�г�ʼ��������Ϣ,�����޷�����ִ�в��š����ȵ����Ź��������ã�", vbInformation, gstrSysName
        Exit Function
    End If
    
    
    '�����շ����:���������
    If gstr�շ���� = "" Then
        strSQL = "Select ����,���� as ��� from �շ���Ŀ��� Where ����<>'1' Order by ���"
    Else
        strSQL = "" & _
        "   Select /*+ RULE */   A.����,A.���� as ��� " & _
        "   From �շ���Ŀ��� A, Table( f_Str2list([1])) J " & _
        "   Where A.����=J. Column_Value " & _
        "   Order by ���"
    End If
    Set mrsClass = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Replace(gstr�շ����, "'", ""))
    If mrsClass.EOF Then
        MsgBox "û�����ÿ��õ��շ����,�����ڱ��ز��������ã�", vbInformation, gstrSysName
        Exit Function
    End If
    '��ֻ��һ�ֿ�ѡ�շ����ʱ,�����û�ѡ��
    mblnOne = (mrsClass.RecordCount = 1)
    If InStr(gstr�շ����, "'5'") > 0 Or InStr(gstr�շ����, "'6'") > 0 _
        Or InStr(gstr�շ����, "'7'") > 0 Or gstr�շ���� = "" Then
        mlngҩƷ���ID = ExistIOClass(IIf(mbytInFun = 2, 9, 8))
        If mlngҩƷ���ID = 0 Then
            MsgBox "����ȷ���������ݵ�������,���ȵ����������������ã�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    If InStr(gstr�շ����, "'4'") > 0 Or gstr�շ���� = "" Then
        mlng�������ID = ExistIOClass(IIf(mbytInFun = 2, 41, 40))
        If mlng�������ID = 0 Then
            MsgBox "����ȷ�����ĵ��ݵ�������,���ȵ����������������ã�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
        
    '��������
    strSQL = " Select 'ҽ��' As ���,����,���� From �������� Where ���� In(" & gstrҽ���������� & ") Union All " & _
                 " Select '����' As ���,����,���� From �������� Where ���� In(" & gstr���ѷ������� & ") "
    Set mrs�������� = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(mrs��������, strSQL, Me.Caption)
       
        
    '��������
    If mbln���� And mstr���ת��ʱ�� <> "" Then
        Curdate = CDate(mstr���ת��ʱ��) - 1 / 24 / 60
    Else
        Curdate = zlDatabase.Currentdate
    End If
    txtDate.Text = Format(Curdate, "yyyy-MM-dd HH:mm:ss")
    '�Զ�ʶ��Ӱ�
    If mbytInState <> 2 And mstrInNO = "" Then
        If OverTime(Curdate) Then chk�Ӱ�.Value = 1
    End If
    
    InitData = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetLastDeptID(ByVal str��� As String, ByVal intPage As Integer, ByVal lngRow As Long, ByVal strDeptIDs As String) As Long
'���ܣ���ȡ����������ͬ�����Ŀ��ִ�п���ID
    Dim i As Long, j As Long, k As Long
    
    For i = intPage To 1 Step -1
        If i = intPage Then
            k = lngRow - 1
        Else
            k = mobjBill.Pages(i).Details.Count
        End If
        For j = k To 1 Step -1
            If mobjBill.Pages(i).Details(j).�շ���� = str��� _
                And mobjBill.Pages(i).Details(j).ִ�в���ID <> 0 Then
                If InStr("," & strDeptIDs & ",", "," & mobjBill.Pages(i).Details(j).ִ�в���ID & ",") > 0 Then
                    GetLastDeptID = mobjBill.Pages(i).Details(j).ִ�в���ID
                    Exit Function
                End If
            End If
        Next
    Next
    
    '�������������,��ȡ��������������ƥ���ִ�п���
    If str��� = "4" Then
        For i = intPage To 1 Step -1
            If i = intPage Then
                k = lngRow - 1
            Else
                k = mobjBill.Pages(i).Details.Count
            End If
            For j = k To 1 Step -1
                If mobjBill.Pages(i).Details(j).ִ�в���ID <> 0 Then
                    If InStr("," & strDeptIDs & ",", "," & mobjBill.Pages(i).Details(j).ִ�в���ID & ",") > 0 Then
                        GetLastDeptID = mobjBill.Pages(i).Details(j).ִ�в���ID
                        Exit Function
                    End If
                End If
            Next
        Next
    End If
End Function

Private Sub FillBillComboBox(ByVal lngRow As Long, ByVal lngCol As Long, Optional blnEnter As Boolean)
'���ܣ����ݵ��������������б������
'������blnEnter=�Ƿ񰴹�������д���,��ʱ��ʾ�����ݱ��ֲ���
    Dim rsTmp As ADODB.Recordset
    Dim bln��ʿ As Boolean, strTmp As String
    Dim strSQL As String, strIDs As String, i As Long
    Dim lng����ID As Long, lng����ID As Long, j As Long
    Dim bln��ҩ��� As Boolean '�Ƿ����������ҩ���
    Dim rsUnit As ADODB.Recordset
    Bill.Clear
    Err = 0: On Error GoTo Errhand:
    Select Case Bill.TextMatrix(0, lngCol)
        Case "���"
            Call GetOperatorInfo(mobjBill.Pages(mintPage).������, bln��ʿ)
            
                    
            mrsClass.Filter = 0: j = 1
            For i = 1 To mrsClass.RecordCount
                '��ʿ���:����
                If Not (bln��ʿ And InStr(",E,M,4,", mrsClass!����) = 0) Then
                    Bill.AddItem j & "-" & mrsClass!���
                    Bill.ItemData(Bill.NewIndex) = Asc(mrsClass!����)  '����������ASCII��
                    j = j + 1
                End If
                mrsClass.MoveNext
            Next
            Bill.cboStyle = DropOlnyDown
            
        Case "ִ�п���", "��ҩҩ��"
            Bill.cboStyle = DropDownAndEdit
            'Bill.ToolTipText = "ִ�п��ҵ�ǰ��Ŀ��ִ�п�������,���ұ��������,������Դ�����,�����ҩƷ,��洢�ⷿ,���ʶ�Ӧ�Ĳ��Ź������ʵ����"
            '���ݵ�ǰ��Ŀִ�п�������,��̬���ÿ�ѡ����
            If mobjBill.Pages(mintPage).Details.Count >= lngRow Then
                With mobjBill.Pages(mintPage).Details(lngRow)
                    If InStr(",4,5,6,7,", .�շ����) > 0 Then
                        Call GetWorkUnit(.�շ�ϸĿID, .�շ����)
                        If mrsWork.RecordCount > 0 Then
                            'ȡ��һ��ҩ��ҩ��
                            mrsWork.MoveFirst
                            For i = 1 To mrsWork.RecordCount
                                strIDs = strIDs & "," & mrsWork!ID
                                mrsWork.MoveNext
                            Next
                            If Not blnEnter Then '�������ʱ������ȷ��ֵ����
                                lng����ID = GetLastDeptID(.�շ����, mintPage, lngRow, Mid(strIDs, 2))
                            End If
                            If lng����ID = 0 Then lng����ID = .ִ�в���ID
                            
                            'ȷ����ǰ�е�ҩ��
                            mrsWork.MoveFirst
                            For i = 1 To mrsWork.RecordCount
                                Bill.AddItem IIf(zlIsShowDeptCode, mrsWork!���� & "-", "") & mrsWork!����
                                Bill.ItemData(Bill.NewIndex) = mrsWork!ID
                                If mrsWork!ID = lng����ID Then Bill.ListIndex = Bill.NewIndex
                                mrsWork.MoveNext
                            Next
                            
                        End If
                    Else
                        Bill.TextMatrix(lngRow, lngCol) = ""
                        
                        lng����ID = mobjBill.����ID     '���˿���
                        If lng����ID = 0 Then lng����ID = Get��������ID
                                            
                        If gint������Դ = 2 Then
                            lng����ID = mobjBill.����ID
                            If lng����ID = 0 Then lng����ID = Get����ID(lng����ID)
                        End If
                        If lng����ID = 0 Then lng����ID = lng����ID
                        
                        '0-����ȷ,1-���˿���,2-���˲���,3-����Ա����,4-ָ������,5-Ժ��ִ��(Ԥ��,������δ��),6-�����˿���
                        Select Case .Detail.ִ�п���
                            Case 0 '����ȷ
                                mrsUnit.Filter = 0
                            Case 1 '���˿���
                                mrsUnit.Filter = "ID=" & lng����ID & " Or ID=" & .ִ�в���ID
                            Case 2 '���˲���
                                mrsUnit.Filter = "ID=" & lng����ID & " Or ID=" & .ִ�в���ID
                            Case 3 '����Ա����
                                mrsUnit.Filter = "ID=" & UserInfo.����ID & " Or ID=" & .ִ�в���ID
                            Case 4 'ָ������
                                strSQL = "" & _
                                "   Select Nvl(A.��������ID,0) as ��������ID,A.ִ�п���ID" & _
                                "   From �շ�ִ�п��� A,���ű� C" & _
                                "   Where A.�շ�ϸĿID=[1]��And A.ִ�п���ID+0=C.ID " & _
                                "       And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                                "       And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null) " & vbNewLine & _
                                "       And (A.������Դ is NULL Or A.������Դ=[2])" & _
                                "       And (A.��������ID is NULL Or A.��������ID=[3])" & _
                                " Order by Decode(A.������Դ,Null,2,1)" 'Ĭ�Ͽ�������
                                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .�շ�ϸĿID, gint������Դ, lng����ID)
                                
                                If Not rsTmp.EOF Then
                                    For i = 1 To rsTmp.RecordCount
                                        strTmp = strTmp & "ID=" & rsTmp!ִ�п���ID & " OR "
                                        rsTmp.MoveNext
                                    Next
                                    strTmp = strTmp & "ID=" & .ִ�в���ID & " OR "
                                    strTmp = Left(strTmp, Len(strTmp) - 4)
                                    mrsUnit.Filter = strTmp
                                Else
                                    mrsUnit.Filter = "ID=" & UserInfo.����ID & " Or ID=" & .ִ�в���ID
                                End If
                             Case 5 'Ժ��ִ��(Ԥ��,������δ��)
                             Case 6 '�����˿���
                                mrsUnit.Filter = "ID=" & Get��������ID & " Or ID=" & .ִ�в���ID
                        End Select
                        
                        If mrsUnit.EOF Then mrsUnit.Filter = "ID=" & UserInfo.����ID & " Or ID=" & .ִ�в���ID
                        Set rsUnit = Rec.CopyNew(mrsUnit)
                        If Not rsUnit.EOF Then
                            For i = 1 To rsUnit.RecordCount
                                strTmp = IIf(zlIsShowDeptCode, rsUnit!���� & "-", "") & rsUnit!����
                                If zlCboFindItem(Bill.cboObj, Val(Nvl(rsUnit!ID))) = False Then
                                '���˺�:28947
                                'If Not (SendMessage(Bill.cboHwnd, CB_FINDSTRING, -1, ByVal strTmp) >= 0) Then
                                    Bill.AddItem strTmp
                                    Bill.ItemData(Bill.NewIndex) = rsUnit!ID
                                    
                                   '����ȱʡִ�п���
                                    If Not blnEnter Then '�������ʱ������ȷ��ֵ����
                                        If lngRow = 1 Then
                                            If rsUnit!ID = lng����ID Then Bill.ListIndex = Bill.NewIndex
                                        ElseIf lngRow > 1 Then
                                            '����һ�з�ҩƷ��ͬ
                                            If rsUnit!ID = mobjBill.Pages(mintPage).Details(lngRow - 1).ִ�в���ID And mobjBill.Pages(mintPage).Details(lngRow - 1).Detail.ִ�п��� = .Detail.ִ�п��� _
                                                And InStr(",5,6,7,", mobjBill.Pages(mintPage).Details(lngRow - 1).�շ����) = 0 Then
                                                Bill.ListIndex = Bill.NewIndex
                                            ElseIf rsUnit!ID = lng����ID And Bill.ListIndex = -1 Then
                                                Bill.ListIndex = Bill.NewIndex
                                            End If
                                        End If
                                    End If
                                End If
                                rsUnit.MoveNext
                            Next
                        End If
                            
                        If Not blnEnter And .Detail.ִ�п��� = 4 Then    'ִ�п���Ϊָ�����ҵ�,ȱʡΪ����Ա���ڿ���
                            For i = 0 To Bill.ListCount - 1
                                If Bill.ItemData(i) = UserInfo.����ID Then Bill.ListIndex = i: Exit For
                            Next
                        End If
                        If Bill.ListIndex = -1 Then '���û����ȡ���е�ִ�п���
                            For i = 0 To Bill.ListCount - 1
                                If Bill.ItemData(i) = .ִ�в���ID Then Bill.ListIndex = i: Exit For
                            Next
                        End If
                        If Bill.ListIndex = -1 And Bill.ListCount > 0 Then Bill.ListIndex = 0
                    End If
                    
                    If Bill.ListIndex <> -1 Then
                        .ִ�в���ID = Bill.ItemData(Bill.ListIndex)
                        Bill.TextMatrix(lngRow, lngCol) = Bill.List(Bill.ListIndex)
                    Else
                        .ִ�в���ID = 0
                    End If
                End With
            End If
    End Select
    Exit Sub
Errhand:
    If ErrCenter = 1 Then Resume
End Sub
Private Sub InitPara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������
    '����:���˺�
    '����:2010-01-27 10:17:11
    '����:27663
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With mTy_Para
        .blnסԺ���������շ� = IIf(Val(zlDatabase.GetPara("סԺ���˰������շ�", glngSys, mlngModul, "0")) = 1, True, False)
        If mbytInFun = 2 Then .blnסԺ���������շ� = True 'סԺ���˼���ʱĬ�ϰ������շ�
    End With
    mint�˷ѻص���ӡ = Val(zlDatabase.GetPara("�˷ѻص���ӡ��ʽ", glngSys, mlngModul, "0"))
End Sub


Private Sub InitFace()
'���ܣ����ݱ�Ҫ��ɵĹ������ý��沼��
    Dim arrHead() As String, i As Integer, arrBaby As Variant, strTmp As String
    Dim blnStatusIn As Boolean
    
    '���˺� ����:27331 ����:2010-01-12 09:48:43
    If (mbytInFun = 0 Or mbytInFun = 1) And mbytInState = 0 Then
       'ֻ�л��۲Ż��д��ж�
       MCPAR.blnOnlyBjYb = zlIsOnly����ҽ��
    Else
        MCPAR.blnOnlyBjYb = False
    End If
    
    '���˺� ����:27663 ����:2010-01-27 10:18:39
    Call InitPara
    
    
    '���õ��ݱ��ʽ
    With Bill
        .Font.Size = 10.5
        .CboFont.Size = 11
        .TxtEditFont.Size = 11
        
        arrHead = Split(STR_HEAD, ";")
        .COLS = UBound(arrHead) + 1
        
        .MsfObj.FixedCols = 1
        .MsfObj.ScrollBars = flexScrollBarVertical
        .LocateCol = BillCol.��Ŀ
        .PrimaryCol = BillCol.��Ŀ
        .MsfObj.ColAlignmentFixed(BillCol.��) = 4
        .TextMatrix(1, BillCol.��) = 1
        
        For i = 0 To UBound(arrHead)
            If glngSys Like "8??" And Split(arrHead(i), ",")(0) = "ִ�п���" Then
                .TextMatrix(0, i) = "��ҩҩ��"
            Else
                .TextMatrix(0, i) = Split(arrHead(i), ",")(0)
            End If
            If glngSys Like "8??" And .TextMatrix(0, i) = "��־" Then
                .ColWidth(i) = 0 '��Ҫ������־
            ElseIf glngSys Like "8??" And .TextMatrix(0, i) = "���" Then
                .ColWidth(i) = Split(arrHead(i), ",")(1) + 270
            ElseIf glngSys Like "8??" And .TextMatrix(0, i) = "����" Then
                .ColWidth(i) = Split(arrHead(i), ",")(1) + 250
            Else
                .ColWidth(i) = Split(arrHead(i), ",")(1)
            End If
            .ColAlignment(i) = Split(arrHead(i), ",")(2)
        Next
        
           
        
        If mbytInState = 0 And mbytBilling <> 2 Then
            .ColData(BillCol.��) = BillColType.UnFocus
            .ColData(BillCol.���) = IIf(gbln�շ����, BillColType.ComboBox, BillColType.UnFocus)
            If mblnOne Then .ColData(BillCol.���) = BillColType.UnFocus
            
            .ColData(BillCol.��Ŀ) = BillColType.CommandButton    '��Ŀ����,��Ť��ѡ
            .ColData(BillCol.����) = BillColType.Text             '��/������
            .ColData(BillCol.���) = BillColType.UnFocus          '�������
            .ColData(BillCol.��Ʒ��) = BillColType.UnFocus          '��Ʒ������
            .ColData(BillCol.��λ) = BillColType.UnFocus          '��λ����
            .ColData(BillCol.����) = BillColType.UnFocus          '����ȱʡ����(=1),�����Ϊ��ҩʱ,��Ϊ����(4)(��ֵ,һ��ȫ��)
            .ColData(BillCol.����) = BillColType.UnFocus          '����ȱʡ����,����Ŀ���ʱ,��Ϊ����(4)
            .ColData(BillCol.Ӧ�ս��) = BillColType.UnFocus          'Ӧ�ս������
            .ColData(BillCol.ʵ�ս��) = BillColType.UnFocus          'ʵ�ս������
            .ColData(BillCol.ִ�п���) = BillColType.ComboBox        'Ĭ��ȡ�������һ���һ����
            .ColData(BillCol.��־) = BillColType.UnFocus         '��־ȱʡ����,��Ϊ����ʱ,��Ϊ��ѡ(-1)
            .ColData(BillCol.����) = BillColType.UnFocus         '����ȱʡ����
        End If
        If mbytInState = 0 Or mbytInState = 2 Then '�༭����
            .SetColColor BillCol.���, &HE7CFBA
            .SetColColor BillCol.��Ŀ, &HE7CFBA
            .SetColColor BillCol.����, &HE7CFBA
            .SetColColor BillCol.ִ�п���, &HE7CFBA
            .SetColColor BillCol.����, &HE0E0E0
            .SetColColor BillCol.����, &HE0E0E0
            .SetColColor BillCol.��־, &HE0E0E0
        End If
        
        ReDim marrColData(.COLS - 1)
        For i = 0 To .COLS - 1
            marrColData(i) = .ColData(i)
        Next

        If mbytInState = 3 Then .AllowAddRow = False
    End With
    '�ָ�ע�������
    Call RestoreFlexState(Bill, App.ProductName & "\" & Me.Name & mbytInFun & mbytInState)
    If gTy_System_Para.bytҩƷ������ʾ <> 2 Then
        '0-��ʾͨ������1-��ʾ��Ʒ����2-ͬʱ��ʾͨ��������Ʒ��
        Bill.ColWidth(BillCol.��Ʒ��) = 0
    Else
        If Bill.ColWidth(BillCol.��Ʒ��) = 0 Then
             Bill.ColWidth(BillCol.��Ʒ��) = GetOrigColWidth(BillCol.��Ʒ��)
        End If
    End If
    
    Me.KeyPreview = True
    Set mobjBrushCheck = New clsBrushCardInput
    mobjBrushCheck.OnlyLegalCardNo = False
    Call mobjBrushCheck.InitCompents(Me, Bill, mobjCard)
        
    '��ȡ����ƥ�䷽ʽ
    sta.Panels("MedicareType").Visible = mbytInState = 0
    sta.Panels("PY").Visible = mbytInState = 0 And gbln�����л� '35242
    sta.Panels("WB").Visible = mbytInState = 0 And gbln�����л�
    If mbytInState = 0 Then
        '����ƥ�䷽ʽ��0-ƴ��,1-���,2-����
        If gbytCode = 0 Then
            sta.Panels("PY").Bevel = sbrInset
            sta.Panels("WB").Bevel = sbrRaised
        ElseIf gbytCode = 1 Then
            sta.Panels("PY").Bevel = sbrRaised
            sta.Panels("WB").Bevel = sbrInset
        Else
            sta.Panels("PY").Bevel = sbrInset
            sta.Panels("WB").Bevel = sbrInset
        End If
        
    End If
    
    IDKind.Enabled = mbytInState = 0
    If mbytInState = 0 Then
        Call GetRegisterItem(g˽��ģ��, Me.Name, "idkind", strTmp)
        IDKind.IDKind = Val(strTmp)
    End If
    
    '�൥���շ�:Ŀ¼��֧���շѽ���
    fraBill.Visible = mbytInFun = 0 And mbytInState = 0 And mstrInNO = "" And gblnMulti
    lblDuty.Caption = ""
    fraSubBill.Visible = mbytInFun = 0 And mbytInState = 0    '�����ϻ�Ҫ��ʾ�����˵�רҵ����ְ��
    
    '���˺� ����:26949 ����:2009-12-28 13:52:50
    fra�˷�ժҪ.Visible = (mbytInFun = 0 And mbytInState = 3 Or mblnDelete)
    '25187
    vsInvoice.Visible = (mbytInFun = 0 And mbytInState = 3 Or mblnDelete) And gTy_Module_Para.bytƱ�ݷ������ <> 0
    
    If Not (mbytInFun = 0 And mbytInState = 0 And mstrInNO = "" And InStr(mstrPrivs, "�����շ�") > 0 And gint������Դ = 1) Then
        cmdYB.Visible = False
        lblRePrint.Left = lblRePrint.Left - cmdYB.Width
        txtRePrint.Left = txtRePrint.Left - cmdYB.Width
        lblModi.Left = lblModi.Left - cmdYB.Width
        txtModi.Left = txtModi.Left - cmdYB.Width
        lblIn.Left = lblIn.Left - cmdYB.Width
        txtIn.Left = txtIn.Left - cmdYB.Width
    End If
    cmdSelWholeSet.Visible = mbytInState = 0
    cmdSaveWholeSet.Visible = InStr(mstrPrivs, ";���ӳ�����Ŀ;") > 0
    
    '��ҩ�䷽:�µ����޸�ʱ��Ч
    If Not (mbytInState = 0) Or mbytBilling = 2 Then
        cmd�䷽.Visible = False
        lblRePrint.Left = lblRePrint.Left - cmd�䷽.Width
        txtRePrint.Left = txtRePrint.Left - cmd�䷽.Width
        lblModi.Left = lblModi.Left - cmd�䷽.Width
        txtModi.Left = txtModi.Left - cmd�䷽.Width
        lblIn.Left = lblIn.Left - cmd�䷽.Width
        txtIn.Left = txtIn.Left - cmd�䷽.Width
    End If
                    
    '�ش�(���շ���Ч)
    If Not (mbytInFun = 0 And mbytInState = 0 And mstrInNO = "" And InStr(mstrPrivs, "�ش�Ʊ��") > 0) Then
        lblRePrint.Visible = False
        txtRePrint.Visible = False
        
        lblModi.Left = lblModi.Left - lblRePrint.Width - txtRePrint.Width
        txtModi.Left = txtModi.Left - lblRePrint.Width - txtRePrint.Width
        lblIn.Left = lblIn.Left - lblRePrint.Width - txtRePrint.Width
        txtIn.Left = txtIn.Left - lblRePrint.Width - txtRePrint.Width
    End If
    
    '�޸�(���շ�,������Ч)
    If Not (((mbytInFun = 0 And InStr(";" & mstrPrivs & ";", ";��¼�޸�;") > 0) Or (mbytInFun = 1 And InStr(";" & mstrPrivs & ";", ";�޸�;") > 0)) And mbytInState = 0 And mstrInNO = "") Then
        lblModi.Visible = False
        txtModi.Visible = False
        
        lblIn.Left = lblIn.Left - lblModi.Width - txtModi.Width
        txtIn.Left = txtIn.Left - lblModi.Width - txtModi.Width
    End If

    '����(������ʱ��Ч)
    If Not (mbytInState = 0 And mstrInNO = "") Or mbytBilling = 2 Then
        lblIn.Visible = False
        txtIn.Visible = False
    End If
    Line3.Visible = mbytInFun = 0 And mbytInState = 0 And mstrInNO = ""
    Line4.Visible = mbytInFun = 0 And mbytInState = 0 And mstrInNO = ""
        
    
    '�����ʻ���Ԥ����
    If mbytInFun <> 0 Then
        vsBalance.Visible = False
        lblԤ�����.Visible = False
        txtԤ�����.Visible = False
    End If
    
    'Ӥ����
    cboBaby.Visible = (mbytInFun = 2)
    lblBaby.Visible = (mbytInFun = 2)
    If mbytInFun = 2 Then
        arrBaby = Array("0-���˱���", "1-��1��Ӥ��", "2-��2��Ӥ��", "3-��3��Ӥ��", "4-��4��Ӥ��", "5-��5��Ӥ��")
        For i = 0 To UBound(arrBaby)
            cboBaby.AddItem arrBaby(i)
        Next
        cboBaby.ListIndex = 0
    End If
    '���㷽ʽ
    '������,�޸�,�˷ѿ���
    cbo���㷽ʽ.Enabled = (mbytInFun = 0 And (mbytInState = 0 Or mbytInState = 3))
    
    '����͵���ʱ���ɼ�
    'ֻ���˷Ѳ���ʾ
    lbl���㷽ʽ.Visible = (mbytInFun = 0 And Not (mbytInState = 1 Or mbytInState = 2 Or mbytInState = 0))
    cbo���㷽ʽ.Visible = (mbytInFun = 0 And Not (mbytInState = 1 Or mbytInState = 2 Or mbytInState = 0 Or mbytInState = 4 Or mbytInState = 5))
    fra�ɿ�.Visible = (mbytInFun = 0 And (mbytInState = 3))
    If mbytInFun = 0 And mbytInState = 1 Then
        Set lblԤ�����.Container = picAppend
        Set txtԤ�����.Container = picAppend
         vsBalance.Width = vsBalance.Width + 100
    End If
    'Ʊ�ݺ�
    lblFact.Visible = (mbytInFun = 0)
    txtInvoice.Visible = (mbytInFun = 0)
    txtMCInvoice.Top = txtInvoice.Top   '��Ԥ�����Ż���ʾ
    txtMCInvoice.Left = txtInvoice.Left
    
    '��̬�ѱ�
    If glngSys Like "8??" Or mbytInFun = 2 Then
        lbl��̬�ѱ�.Visible = False
        If mbytInFun = 2 Then cbo�ѱ�.Locked = True: cbo�ѱ�.TabStop = Not cbo�ѱ�.Locked And gbln�ѱ�
    Else
        If mbytInState = 1 Or mbytInState = 2 Or mbytInState = 3 Then
            cbo�ѱ�.Locked = True: cbo�ѱ�.TabStop = Not cbo�ѱ�.Locked And gbln�ѱ�
            lbl��̬�ѱ�.Left = cbo�ѱ�.Left
            lbl��̬�ѱ�.Visible = True
        Else
            lbl��̬�ѱ�.BorderStyle = 0
        End If
    End If
    lbl����.Caption = ""
    
    '�շ�ʱ�Ƿ�����Һ�
    Call ShowRegist
    
    '�շ�ʱ�Ƿ�������￨
    Call ShowIDCard
    
    '�շ�Ʊ�ݴ�ӡ��ʽ:�շ�,�޸�,�˷�ʱ��ʾ
    If mbytInFun = 0 And InStr(",0,3,", mbytInState) > 0 Then
        Call ZlShowBillFormat(mlngModul, lblFormat, mintInvoiceFormat)
    End If
    
    '�˷����ʰ�ť
    If mbytInFun = 0 And mstrInNO = "" And gblnMulti Then
        cmdDelete.Visible = True '�շ�֧�ֶ൥��ʱʹ�ö൥���˷�
        chkCancel.Visible = False
    End If
    If Not (mbytInState = 0 And mstrInNO = "") Or mbytInFun = 1 Then
        chkCancel.Visible = False
    End If
    
    Select Case mbytInFun
        Case 0 '�շ�
            If glngSys Like "8??" Then
                Caption = "ҩ���շѴ���"
                lblTitle.Caption = gstrUnitName & "ҩ���շѵ�"
            Else
                Caption = "�����շѴ���"
                lblTitle.Caption = gstrUnitName & "�����շѵ�"
            End If
            
            Call SetMoneyList
            
            Call InitBalanceGrid
            
            If mbytInState <> 0 Then
                '���շ�״̬
                lbl�ۼ�.Visible = False
                txt�ۼ�.Visible = False
                lblӦ��.Top = lblӦ��.Top + txt�ۼ�.Height / 3
                txtӦ��.Top = txtӦ��.Top + txt�ۼ�.Height / 3
                lbl�ϼ�.Top = lbl�ϼ�.Top + txt�ۼ�.Height / 1.5
                txt�ϼ�.Top = txt�ϼ�.Top + txt�ۼ�.Height / 1.5
            Else
                If Not gbln�ۼ� Then
                    lbl�ۼ�.Visible = False
                    txt�ۼ�.Visible = False
                    lblӦ��.Top = lblӦ��.Top + txt�ۼ�.Height / 3
                    txtӦ��.Top = txtӦ��.Top + txt�ۼ�.Height / 3
                    lbl�ϼ�.Top = lbl�ϼ�.Top + txt�ۼ�.Height / 1.5
                    txt�ϼ�.Top = txt�ϼ�.Top + txt�ۼ�.Height / 1.5
                End If
            End If
            
            '�������
            Call SetInputItem
            
            'Ȩ������
            If InStr(mstrPrivs, "�����˷�") = 0 Then
                chkCancel.Visible = False
                cmdDelete.Visible = False
            End If
            txtInvoice.Locked = Not (InStr(1, mstrPrivs, "�޸�Ʊ�ݺ�") > 0) And gblnStrictCtrl
        Case 1 '����
            Caption = "���ﻮ�۴���"
            
            lblTitle.Caption = gstrUnitName & "���ﻮ�۵�(" & UserInfo.���� & ")"
            mshMoney.Width = mshMoney.Width * 2
            Call SetMoneyList
            Call SetStatPosition
'
'            lbl�ɿ�.Visible = False: txt�ɿ�.Visible = False
'            lbl�Ҳ�.Visible = False: txt�Ҳ�.Visible = False
                        
            '�������
            Call SetInputItem
        Case 2 '�������
            Caption = "������ʴ���"
            
            Select Case mbytBilling
                Case 0
                    lblTitle.Caption = gstrUnitName & "������ʵ�"
                Case 1
                    lblTitle.Caption = gstrUnitName & "������ʵ�(����)"
                Case 2
                    lblTitle.Caption = gstrUnitName & "������ʵ�(���)"
                    
                    cboNO.Locked = False
                    fraInfo.Enabled = False
                    fraAppend.Enabled = False
                    Bill.Active = False
                    
                    Call SetPatientEnableModi(False)
            End Select
            
            lblCorp.Visible = (mbytInState = 0)
            lblCorp.Left = txtIn.Left + txtIn.Width + 100
            lblCorp.Top = lblIn.Top
            
            chkCancel.Caption = "��"
            lblFlag.Caption = "��"
                        
            cboҽ�Ƹ���.Locked = True
            
            Call SetMoneyList
            Call SetStatPosition
            
            
'''            lbl�ɿ�.Visible = False: txt�ɿ�.Visible = False
'''            lbl�Ҳ�.Visible = False: txt�Ҳ�.Visible = False
                       
            
            'Ȩ������
            If InStr(mstrPrivs, "��������") = 0 Then
                chkCancel.Visible = False
            End If
            
            lblTotal.Visible = True
            lblTotal.Top = cmdOK.Top
            
            cboSex.Enabled = False
            txt����.Enabled = False
            cbo���䵥λ.Enabled = False
    End Select
    
    If mbln���� Then
        If mlng��ҳID <> 0 Then
            Dim strSQL As String, rsTemp As ADODB.Recordset
            strSQL = "Select ��ǰ����ID,��Ժ����ID,��Ժ���� From ������ҳ Where ����ID = [1] And ��ҳID = [2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
            If Not rsTemp.EOF Then
                If mlngDeptID = 0 Then
                    mlngDeptID = Val(Nvl(rsTemp!��Ժ����ID))
                End If
                If mlngUnitID = 0 Then
                    mlngUnitID = Val(Nvl(rsTemp!��ǰ����ID))
                End If
                blnStatusIn = IsNull(rsTemp!��Ժ����)
            End If
        End If
        If blnStatusIn Or mlng��ҳID = 0 Or rsTemp.EOF Then
            lblTitle.Caption = lblTitle.Caption & "(" & "����" & ")"
        Else
            lblTitle.Caption = lblTitle.Caption & "(��" & mlng��ҳID & "�β���" & ")"
        End If
    End If
        
    If mbytInState = 0 Or mbytInState = 2 Or mbytInState = 4 Or mbytInState = 5 Then
        'ִ�л����״̬
        
        If mbytInFun = 2 Then
            If mbytInState <> 0 Or mbytBilling = 2 Then Call SetDisible       '�������
        End If
        
        If mbytInState = 0 Then
            If mstrInNO <> "" Then txtPatient.BackColor = &HE0E0E0           '�޸�
            If mbytBilling = 0 Or mbytBilling = 1 Then Call SetShowCol       '���ʡ�����
                        
        ElseIf mbytInState = 2 Then  '���������˺�ʱ��
            Call SetDisible
            
            txtInvoice.Enabled = False
            fraInfo.Enabled = False
                            
            cbo������.Locked = False
            txtDate.Enabled = True
            Call SetShowCol
        End If
        
        Call SetButton(2) 'ȷ��,ȡ��
    Else
        '���� ���˷�,����
        Call SetDisible
        
        fraAppend.Enabled = False
        
        fraTitle.Enabled = False
        fraInfo.Enabled = False
        
        If mbytInState = 3 Then  '�˷�
            If mbytInFun = 0 Then
                '�����˷�ֻ֧��ָ�����㷽ʽ
                fraAppend.Enabled = True
                cbo���㷽ʽ.Locked = False
            End If
            Call SetButton(2) 'ȷ��,ȡ��
            Call ShowDeleteCol(True)
            Bill.Active = True
            fra�˷�ժҪ.Enabled = True
        Else
            Call SetButton(3) 'ȡ��
            fra�˷�ժҪ.Enabled = False
            
        End If
        
        If mblnDelete Then lblFlag.Visible = True
    End If
    
    If gbyt����ҽ�� = 0 Then
        Call ExChangeLocate(cbo��������, cbo������)
        lbl����.Caption = "������(&W)"
        lbl����.Left = lblPatient.Left
        lbl������.Caption = "��������"
        cbo��������.TabStop = False
    End If
    
    If Not (mbytInState = 0 And (mbytInFun = 0 Or mbytInFun = 1)) Then
        sta.Panels("Drugstore").Visible = False
    End If
    
    If mbytInState = 0 And mstrInNO = "" Then
        sta.Panels("PatiSource").Visible = True
        Set sta.Panels("PatiSource").Picture = imgPati.ListImages(IIf(gint������Դ = 1, "OutPati", "InPati")).Picture
    Else
        sta.Panels("PatiSource").Visible = False
    End If
    Bill.ColWidth(BillCol.��������) = 0
    Bill.ColWidth(BillCol.ҽ�����) = 0
    Bill.ColWidth(BillCol.ִ�п���ID) = 0
    
    '82801,Ƚ����,2015-2-26
    txt����.MaxLength = zlGetPatiInforMaxLen.intPatiAge
End Sub

Private Sub SetStatPosition()
'���ܣ����ﻮ�ۺ��������ʱ���úϼ���Ϣ���ڵĿؼ�λ��
    Dim blnVisible As Boolean
    
    If mbytInState = 0 And mstrInNO = "" And (mbytInFun = 1 Or mbytInFun = 2) Then
        fraUpBillShow.Visible = True
        blnVisible = True
    Else
        blnVisible = False: fraUpBillShow.Visible = False
    End If
    
    fraStat.Width = lbl�ϼ�.Width + txt�ϼ�.Width + 600
    fraStat.Left = mshMoney.Left + mshMoney.Width
    lblӦ��.Left = 200: txtӦ��.Left = lblӦ��.Left + lblӦ��.Width + 200
    lbl�ϼ�.Left = lblӦ��.Left
    txt�ϼ�.Left = txtӦ��.Left
          
    
    If mbytInFun = 1 Then
        lbl�ۼ�.Left = lbl�ϼ�.Left: txt�ۼ�.Left = txt�ϼ�.Left
        '�ۼ��������ֱҴ����Ľ��
        lbl�ۼ�.Visible = True: txt�ۼ�.Visible = True: lbl�ۼ�.Caption = "Ӧ��"
        
        If blnVisible Then
            fraUpBillShow.Left = fraStat.Left + fraStat.Width + 50
        End If
    ElseIf mbytInFun = 2 Then
        lbl�ۼ�.Visible = False: txt�ۼ�.Visible = False
         fraUpBillShow.Visible = False
        If blnVisible Then
            fraUpBillShow.Visible = True
            fraUpBillShow.Left = fraStat.Left + fraStat.Width + 50
'            txtPreNO.Width = txt�ϼ�.Width
'            Set lblPreNO.Container = fraStat
'            Set txtPreNO.Container = fraStat
'            lblPreNO.Left = lbl�ϼ�.Left: txtPreNO.Left = txt�ϼ�.Left
'            lblPreNO.Top = lbl�ۼ�.Top: txtPreNO.Top = txt�ۼ�.Top
        Else
            lblӦ��.Top = lblӦ��.Top + txtPreNO.Height / 2
            txtӦ��.Top = txtӦ��.Top + txtPreNO.Height / 2
            lbl�ϼ�.Top = lbl�ϼ�.Top + txtPreNO.Height * 0.75
            txt�ϼ�.Top = txt�ϼ�.Top + txtPreNO.Height * 0.75
        End If
    End If
End Sub

Private Sub SetButton(bytType As Byte)
'���ܣ����ù��ܰ�ť״̬��λ��
'������bytType=1:Ԥ����,ȷ��,ȡ��
'              2:ȷ��,ȡ��
'              3:ȡ��
'              4:Ԥ����,ȷ��,����շ�,ȡ��
'˵�����ú���Ϊ��ʼʱ����,�����ظ�����
    Dim lngTmp As Long
    
    Const H_��� = 45
    
    LockWindowUpdate picAppend
    
    '�ָ�ȱʡ״̬���Ҳ��ɼ�
    cmdԤ����.Visible = False
    cmdOK.Visible = False
    cmdCancel.Visible = False
    cmdPrint.Visible = False
    
    cmdԤ����.Top = lblSeek.Top
    If mbytInFun = 1 Or mbytInFun = 2 Then
        cmdOK.Top = lblSeek.Top
    Else
        cmdOK.Top = cmdԤ����.Top + cmdԤ����.Height + H_���
    End If
    cmdCancel.Top = cmdOK.Top + cmdOK.Height + H_���
    cmdPrint.Top = cmdCancel.Top + cmdCancel.Height + H_���
            
    cmdCancel.Caption = "ȡ��(&C)"
    cmdOK.Enabled = True
    
    Select Case bytType
        Case 1 'Ԥ����,ȷ��,ȡ��
            cmdԤ����.Visible = True
            cmdOK.Visible = True
            cmdCancel.Visible = True
            
            cmdԤ����.Top = cmdԤ����.Top + cmdPrint.Height / 2 + H_���
            cmdOK.Top = cmdOK.Top + cmdPrint.Height / 2 + H_���
            cmdCancel.Top = cmdCancel.Top + cmdPrint.Height / 2 + H_���
            
            cmdԤ����.TabStop = True
        Case 2 'ȷ��,ȡ��
            cmdOK.Visible = True
            cmdCancel.Visible = True
        Case 3 'ȡ��
            cmdCancel.Visible = True
            cmdCancel.Caption = "�˳�(&X)"
            cmdCancel.Top = cmdCancel.Top - cmdPrint.Height / 2 - H_���
        Case 4 'Ԥ����,ȷ��,��ӡ,ȡ��
            cmdԤ����.Visible = True
            cmdOK.Visible = True
            cmdCancel.Visible = True
            cmdPrint.Visible = True
            
            lngTmp = cmdPrint.Top
            cmdPrint.Top = cmdCancel.Top
            cmdCancel.Top = lngTmp
    End Select
    LockWindowUpdate 0
End Sub

Private Sub SetDisible(Optional bln As Boolean = False)
'����:��������Ϊ�����޸�״̬
'����:blnΪTrue��ʾ����Ϊ�����޸ĵ�״̬

    cboNO.Locked = Not bln
    
    cbo�ѱ�.Locked = Not bln Or mbytInFun = 2: cbo�ѱ�.TabStop = Not cbo�ѱ�.Locked And gbln�ѱ�
    cboҽ�Ƹ���.Locked = Not bln Or mbytInFun = 2
    
    cbo��������.Locked = Not bln Or mbytBilling = 2
    cbo������.Locked = Not bln Or mbytBilling = 2
    cbo��������.Enabled = bln And mbytBilling <> 2
    cbo������.Enabled = bln And mbytBilling <> 2
    
    chk�Ӱ�.Enabled = bln And mbytBilling <> 2
    
    If mbytInFun = 2 And mbytInState = 0 Then
        If bln And mbytBilling <> 2 And cbo��������.ListIndex <> -1 Then
            cboBaby.Enabled = is����(cbo��������.ItemData(cbo��������.ListIndex), mrs��������)
        Else
            cboBaby.Enabled = False
        End If
    End If
    
    cbo���㷽ʽ.Locked = Not bln
    txtDate.Enabled = bln
    fraStat.Enabled = bln
    fra�ɿ�.Enabled = bln
    Bill.Active = bln And mbytBilling <> 2
    
'''    If Not bln Then
'''        txt�ɿ�.BackColor = &HE0E0E0
'''    Else
'''        txt�ɿ�.BackColor = &HFFFFFF
'''    End If
    
    SetPatientEnableModi (bln)
End Sub

Private Sub SetDeptDoctorByRegevent(ByVal lng����ID As Long, Optional strRegNO As String)
'���ܣ����ݲ���ID��Һŵ��в��˵ĹҺſ��Һ�ҽ����Ϣ���ÿ������ҺͿ�����
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strTmp As String
    
    On Error GoTo errH
    strTmp = zlGetRegEventsCons("�Ӱ��־")
    If strRegNO <> "" Then
        strTmp = strTmp & " And NO=[2]"
    Else
        strTmp = strTmp & " And ����ID=[1]"
    End If
    
    strSQL = "Select ִ�в���id, ִ����" & vbNewLine & _
            "From (Select ִ�в���id, ִ����, �Ǽ�ʱ��" & vbNewLine & _
            "       From ������ü�¼" & vbNewLine & _
            "       Where ��¼���� = 4 And ��¼״̬ = 1 " & strTmp & vbNewLine & _
            "       Order By �Ǽ�ʱ�� Desc)" & vbNewLine & _
            "Where Rownum < 2"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, strRegNO)
    If Not rsTmp.EOF Then
        Call Set�����˿�������Click("" & rsTmp!ִ����, Val("" & rsTmp!ִ�в���ID))
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Function GetDeptByRegevent(ByVal lng����ID As Long) As ADODB.Recordset
'���ܣ����ݲ���ID������Ч�Һŵ��Ŀ���ID
    Dim strSQL As String, strWhere As String
    strWhere = zlGetRegEventsCons(, , True)
    On Error GoTo errH
    strSQL = "Select ִ�в���ID From ���˹Һż�¼" & _
            " Where ����ID=[1] and ��¼����=1 and ��¼״̬=1  " & strWhere
    Set GetDeptByRegevent = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub LoadAddedItem(ByVal lng����ID As Long, Optional ByVal str�������� As String)
'����:�Զ����չҺŷ�
    Dim i As Long, j As Long, objThis As Control
    
    '������е������Ƿ��Ѽ���
    For i = 1 To mobjBill.Pages.Count
        For j = 1 To mobjBill.Pages(i).Details.Count
            If mobjBill.Pages(i).Details(j).�շ�ϸĿID = glngAddedItem Then
                Exit Sub
            End If
        Next
    Next
    
    If CheckAddedItem(lng����ID, str��������) Then
        Set objThis = Me.ActiveControl
        '�����ǰ�����ǻ��۵���������һ�ŵ���
        If Not Bill.Active Then
            For i = 1 To mobjBill.Pages.Count
                If mobjBill.Pages(i).NO = "" Then Exit For
            Next
            If i <= mobjBill.Pages.Count Then
                tbsBill.Tabs(i).Selected = True
            Else
                If cmdAddBill.Enabled And cmdAddBill.Visible Then
                    Call cmdAddBill_Click
                Else
                    Exit Sub '��������ŵ����շ�ʱ�������м���
                End If
            End If
        End If
        
        Call LocateNewRow
        If gbln�շ���� Then
            Bill.Col = BillCol.��� '�Զ�����call Bill_EnterCell
            For i = 0 To Bill.ListCount - 1
                If Bill.ItemData(i) = Asc("Z") Then Bill.ListIndex = i: Exit For
            Next
            If i > Bill.ListCount - 1 Then Exit Sub '����������ǻ�ʿ�����ܲ������������򲻽��м���
            
            Call Bill_KeyDown(vbKeyReturn, 0, False)
        End If
        
        Bill.Col = BillCol.��Ŀ
        Bill.TxtVisible = True
        Bill.Text = glngAddedItem
        mblnSelect = True
        Call Bill_KeyDown(vbKeyReturn, 0, False)
        
        On Error Resume Next
        If objThis.Visible And objThis.Enabled Then objThis.SetFocus
        On Error GoTo 0
    End If
End Sub

Private Function GetCboIndexByCode(ByRef objCbo As ComboBox, ByVal strCode As String) As Integer
    Dim i As Integer
    
    GetCboIndexByCode = -1
    For i = 0 To objCbo.ListCount - 1
        If strCode = Mid(objCbo.List(i), 1, InStr(1, objCbo.List(i), "-") - 1) Then
            GetCboIndexByCode = i
            Exit For
        End If
    Next
End Function
Private Sub initInsurePara(ByVal lng����ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ҽ������
    '����:���˺�
    '����:2011-08-27 12:25:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    MCPAR.��������ҽ����Ŀ = gclsInsure.GetCapability(support��������ҽ����Ŀ, lng����ID, mintInsure)
    MCPAR.�����շѴ�Ϊ���۵� = gclsInsure.GetCapability(support�����շѴ�Ϊ���۵�, lng����ID, mintInsure)
    MCPAR.������봫����ϸ = gclsInsure.GetCapability(support������봫����ϸ, lng����ID, mintInsure)
    MCPAR.ҽ��ȷ���������� = gclsInsure.GetCapability(supportҽ��ȷ����������, lng����ID, mintInsure)
    MCPAR.ҽ���ӿڴ�ӡƱ�� = gclsInsure.GetCapability(supportҽ���ӿڴ�ӡƱ��, lng����ID, mintInsure)
    MCPAR.�൥��һ�ν��� = gclsInsure.GetCapability(support�൥��һ�ν���, lng����ID, mintInsure)
    MCPAR.���������շ� = gclsInsure.GetCapability(support���������շ�, lng����ID, mintInsure)
    MCPAR.�൥���շ� = gclsInsure.GetCapability(support�൥���շ�, lng����ID, mintInsure)
    MCPAR.����Ԥ���� = gclsInsure.GetCapability(support����Ԥ��, lng����ID, mintInsure)
    MCPAR.�ֱҴ��� = gclsInsure.GetCapability(support�ֱҴ���, lng����ID, mintInsure)
    MCPAR.���Ը� = gclsInsure.GetCapability(support�շ��ʻ������Ը�, lng����ID, mintInsure)
    MCPAR.ȫ�Ը� = gclsInsure.GetCapability(support�շ��ʻ�ȫ�Է�, lng����ID, mintInsure)
    MCPAR.ʵʱ��� = gclsInsure.GetCapability(supportʵʱ���, lng����ID, mintInsure)
    MCPAR.�˷Ѻ��ӡ�ص� = gclsInsure.GetCapability(support�˷Ѻ��ӡ�ص�, lng����ID, mintInsure)
    MCPAR.�൥�ݵ�һ�ν��� = gclsInsure.GetCapability(support����_���ֵ��ݽ���, lng����ID, mintInsure)
    MCPAR.ҽ������Ʊ�� = False
    '���˺�:27536 20100119
    MCPAR.�����ѽɿ���� = gclsInsure.GetCapability(support�����ѽɿ����, lng����ID, mintInsure)
End Sub


Private Sub MCPatientProcess(Optional ByVal lngCur����ID As Long, Optional blnErrBill As Boolean)
    Dim i As Long, blnTran As Boolean
    Dim lng����ID As Long, lng����IDOut As Long
    Dim lng�Һſ��� As Long, str�������� As String, strSQL As String
    Dim rsTmp As ADODB.Recordset, strTemp As String, intInsure As Integer
    Dim blnPriceBill As Boolean
    
    On Error GoTo errH
    If gblnLED Then zl9LedVoice.Speak "#50"
    lng����IDOut = lngCur����ID '����Identify�ӿ����޸ĸñ����󷵻���ֵ
    
    '���أ�0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);8����ID,24��������(1=��������),25������������
    
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(False)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(False)
    IDKind.SetAutoReadCard (False)
    mstrYBPati = gclsInsure.Identify(id�����շ�, lng����IDOut, mintInsure)
    If Me.ActiveControl Is txtPatient Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
        IDKind.SetAutoReadCard (txtPatient.Text = "")
    End If
    
    blnPriceBill = False
    If mstrYBPati <> "" Then
        '��ȡ������Ϣ
        If UBound(Split(mstrYBPati, ";")) >= 8 Then
            If IsNumeric(Split(mstrYBPati, ";")(8)) And Val(Split(mstrYBPati, ";")(8)) <> 0 Then
                lng����ID = Val(CLng(Split(mstrYBPati, ";")(8)))
                If lng����ID <> lngCur����ID And lngCur����ID <> 0 And lng����ID <> 0 Then
                    MsgBox "ҽ����֤�Ĳ�����֮ǰ��ȡ�Ĳ��˲���ͬһ������!", vbInformation, gstrSysName
                    Call YBIdentifyCancel
                    mintInsure = 0: mstrYBPati = ""
                    Exit Sub
                End If
            End If
        End If
        '����:29283
        '  -- ����:���ó���-1-�Һ�;2-�շ�
        '  --        ����id_In-����ID(δ������,������)
        '  --        ����_In: ˢ������;δˢ��ʱ,Ϊ��
        '  --         ˢ����ʽ_In:  1-����ˢ��;2-ҽ��ˢ��
        If zlPatiCardCheck(2, lng����ID, CStr(Split(mstrYBPati, ";")(0)), 2) = False Then
            Call YBIdentifyCancel
            mintInsure = 0: mstrYBPati = ""
            Exit Sub
        End If
        Call initInsurePara(lng����ID)   '��ʼ��ҽ������
        If (MCPAR.���������շ� Or Not MCPAR.�൥���շ�) And tbsBill.Tabs.Count > 1 Then
            If MCPAR.���������շ� Then
                MsgBox "��ҽ�������շ�ģʽ�²�֧�ֶ��ŵ����շѡ�", vbInformation, gstrSysName
            ElseIf Not MCPAR.�൥���շ� Then
                MsgBox "��ǰ���಻֧�ֶ��ŵ����շѡ�", vbInformation, gstrSysName
            End If
            Call YBIdentifyCancel
            If Visible Then txtPatient.SetFocus
            mintInsure = 0: mstrYBPati = ""
            Exit Sub
        End If
        If MCPAR.�൥��һ�ν��� Then
            If gTy_Module_Para.bln�ֱ��ӡ Then
                MsgBox "�൥��ģʽ�£�ҽ��һ�ν���ʱ��������ֱ��ӡ��", vbInformation, gstrSysName
                Call YBIdentifyCancel
                If Visible Then txtPatient.SetFocus
                mintInsure = 0: mstrYBPati = ""
                Exit Sub
            End If
        End If
            
        '����:28240
        strTemp = mstrYBPati: intInsure = mintInsure
            
        If GetPatient("-" & lng����ID, , , True) Then
            mstrYBPati = strTemp: mintInsure = intInsure
            If Not CheckRegisted(lng����ID) Then
                Call YBIdentifyCancel
                Set mrsInfo = New ADODB.Recordset
                mintInsure = 0: mstrYBPati = ""
                Exit Sub
            End If
            With mobjBill
                .����ID = Nvl(mrsInfo!����ID, 0)
                .��ҳID = Nvl(mrsInfo!��ҳID, 0)
                .��ʶ�� = Nvl(mrsInfo!�����, 0)
                .����ID = Nvl(mrsInfo!��ǰ����ID, 0)
                .����ID = Nvl(mrsInfo!��ǰ����id, 0)
                .���� = "" & mrsInfo!��ǰ����
                .���� = "" & mrsInfo!����
                .�Ա� = "" & mrsInfo!�Ա�
                .���� = "" & mrsInfo!����
                '�ѱ��ں������LoadAndSeek�ѱ�ʱ��ֵ
            End With
            txt�����.Text = Nvl(mrsInfo!�����)
            Call InitBalanceGrid(True)
        Else
            Call YBIdentifyCancel
            mintInsure = 0: mstrYBPati = ""
            Exit Sub
        End If
        
        
        If fraBill.Visible Then
            cmdAddBill.Enabled = Not MCPAR.���������շ� And MCPAR.�൥���շ� And InStr(1, mstrPrivs, "ҽ�����˶൥���շ�") > 0
        End If
        '75259:���ϴ�,2014-7-10������������ʾ��ɫ����
        If Not mrsInfo Is Nothing Then
            If mrsInfo.State = 1 Then
                Call SetPatiColor(txtPatient, Nvl(mrsInfo!��������), vbRed)
            Else
                txtPatient.ForeColor = vbRed
            End If
        Else
            txtPatient.ForeColor = vbRed
        End If
        txtPatient.Text = Split(mstrYBPati, ";")(3)
        txtPatient.PasswordChar = ""
        '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
        txtPatient.IMEMode = 0
        cboSex.ListIndex = cbo.FindIndex(cboSex, CStr(Split(mstrYBPati, ";")(4)), True)
        If IsDate(Split(mstrYBPati, ";")(5)) Then
            txt����.Text = ReCalcOld(CDate(Split(mstrYBPati, ";")(5)), cbo���䵥λ, lng����ID)
        Else
            Call LoadOldData("" & mrsInfo!����, txt����, cbo���䵥λ)
            If Not IsNull(mrsInfo!��������) Then txt����.Text = ReCalcOld(mrsInfo!��������, cbo���䵥λ, lng����ID)
            
        End If
        lbl����.Caption = "" & mrsInfo!��������
        
        mobjBill.����ID = lng����ID
        mobjBill.���� = Split(mstrYBPati, ";")(3)
        mobjBill.�Ա� = Split(mstrYBPati, ";")(4)
        mobjBill.���� = Trim(txt����.Text) & IIf(IsNumeric(txt����.Text), cbo���䵥λ.Text, "")
        
        
        '������������
        If UBound(Split(mstrYBPati, ";")) >= 25 And mobjBill.Pages(mintPage).NO = "" Then   '���۵��Ŀ����˿�����������
            str�������� = CStr(Split(mstrYBPati, ";")(25))
            If str�������� <> "" Then
                Call zlControl.CboSetIndex(cbo��������.hWnd, cbo.FindIndex(cbo��������, str��������, True)) '������click�¼�
                Call cbo��������_Click
            End If
        End If
        '���ݲ��˹Һ���Ϣ���ÿ������Һ�ҽ��
        If mobjBill.Pages(mintPage).NO = "" Then    '���۵��Ŀ����˿�����������
            Call SetDeptDoctorByRegevent(lng����ID)
        End If
        
        '��ʾ������
        If UBound(Split(mstrYBPati, ";")) >= 24 Then
            chk����.Visible = Val(Split(mstrYBPati, ";")(24)) = 1
        End If
        
        '�����ʻ�
        mcur������� = gclsInsure.SelfBalance(lng����ID, CStr(Split(mstrYBPati, ";")(1)), 10, mcur����͸֧, mintInsure)
        sta.Panels(Pan.C3�����ʻ�).Text = "�����ʻ����:" & Format(mcur�������, "0.00")
        sta.Panels(Pan.C3�����ʻ�).Visible = True
        
        '֧��Ԥ����ʱ�Ͳ��̶���ʾ�����ʻ�,������ʾ
        If MCPAR.����Ԥ���� Then
            '��ʾԤ���㰴ť
            cmdԤ����.Enabled = True
            Call SetButton(1) 'Ԥ����,ȷ��,ȡ��
            cmdOK.Enabled = False
        ElseIf mstr�����ʻ� <> "" Then 'ֻ��ʹ�ø����ʻ�����
            Call SetButton(2) 'ȷ��,ȡ��
            vsBalance.TextMatrix(0, 0) = mstr�����ʻ�
            vsBalance.TextMatrix(0, 1) = "0.00"
            vsBalance.RowData(0) = 0
        End If
        
        sta.Panels(Pan.C2��ʾ��Ϣ) = ""
        SetPatientEnableModi (False)
        
        txtRePrint.Enabled = False
        txtModi.Enabled = False '�������
        txtIn.Enabled = False
        cboNO.Enabled = False
        chkCancel.Enabled = False
        cmdDelete.Enabled = False
        
        'һ������δ���,��������һ������(�Һ�)
        If cmdIDCard.Visible Then cmdIDCard.Enabled = False
        If cmdRegist.Visible Then cmdRegist.Enabled = False
        
        If MCPAR.���������շ� Then Call SetButton(4)  'Ԥ����,ȷ��,�����շ�,ȡ��
        
        'ҽ�Ƹ��ʽ
        If mrsInfo.State = 1 Then
            If Not IsNull(mrsInfo!ҽ�Ƹ��ʽ) Then
                cboҽ�Ƹ���.ListIndex = cbo.FindIndex(cboҽ�Ƹ���, mrsInfo!ҽ�Ƹ��ʽ, True)
            End If
        End If
        If cboҽ�Ƹ���.ListIndex = -1 Then cboҽ�Ƹ���.ListIndex = GetCboIndexByCode(cboҽ�Ƹ���, "1")
        
        cboҽ�Ƹ���.Locked = True
        
        '��ȡ���˵Ķ��Ż��۵�,֮ǰ��ȡ������֧�ֶ൥���շ�ʱ������ȡ
        If mbytInFun = 0 And mbytInState = 0 And Visible And mstrInNO = "" And txtIn.Text = "" And mrsInfo.State = 1 And _
            Not (lngCur����ID > 0 And Not MCPAR.���������շ� And MCPAR.�൥���շ� And InStr(1, mstrPrivs, "ҽ�����˶൥���շ�") > 0) Then
            If gblnCheckRegeventDept And gint������Դ = 1 And IsRegisterDept Then lng�Һſ��� = Val("" & mrsInfo!ִ�в���ID)
            blnPriceBill = LoadMultiBills(lng����ID, MCPAR.���������շ� Or Not MCPAR.�൥���շ� Or InStr(1, mstrPrivs, "ҽ�����˶൥���շ�") = 0, lng�Һſ���)
        End If
        
        '�Զ����չҺŷ�
        Call LoadAddedItem(lng����ID)
                    
        '���е����������ݵĴ���
        '--------------------------------------------------------------------
        
        'ҽ�����˵��µ��Ӵ���,���ܽɿ�����Լ��Ƿ�����ͬ����
        '���˺�:22343
        If (gTy_Module_Para.byt�ɿ���� <> 1 And gTy_Module_Para.byt�ɿ���� <> 3) Or mstrPrePati = "" Then    '���Խɿ���Ϊ����ʱ,��ʹ��ͬ�Ĳ���Ҳ�����շ�,'���ǽɿ����
            Call ClearPayInfo
            mstrPrePati = "": mlngPrePati = 0: mstrPreDoctor = ""
            Call InitCommVariable
            Call ClearTotalInfo(True)
            Call ClearMoney
        End If
        
        '��������ȡ�Ļ��۵�����ر�������
        gcnOracle.BeginTrans: blnTran = True
        For i = 1 To tbsBill.Tabs.Count
            If mobjBill.Pages(i).NO <> "" Then
                strSQL = "zl_���ﻮ�ۼ�¼_Update(" & mintInsure & "," & lng����ID & ",'" & mobjBill.Pages(i).NO & "',0)"
                Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            End If
        Next
        gcnOracle.CommitTrans: blnTran = False
        
        'ȫ�����¼��㲢��ʾ
        Call ShowMoney
        
        '�������⴦��
        '---------------------------------------------------------------------------------------
        '����ҽ��
        mblnSaveAsPrice = MCPAR.�����շѴ�Ϊ���۵�
        If mblnSaveAsPrice Then
            Call SetButton(2) 'ȷ��,ȡ��
            sta.Panels(Pan.C3�����ʻ�).Text = ""
            sta.Panels(Pan.C3�����ʻ�).Visible = False
            Call ShowPayInfo(False)
        End If

        '����Ԥ������
        '����ҽ����ʹ��Ԥ�����(����ģʽ)
        '����ҽ����ʹ��Ԥ�����
        If Not mblnSaveAsPrice And mintInsure <> 61 Then Call LoadFeeInfor(lng����ID)
        
        '����ҽ�����ɿ�
        If mintInsure = 61 Then Call ShowPayInfo(False)
                
        If mstrInNO = "" Then
            Call LoadAndSeek�ѱ�
            '49573
            If cmdOK.Enabled And cmdOK.Visible Then
                cmdOK.SetFocus
            ElseIf cbo��������.Enabled And cbo��������.Visible And gbyt����ҽ�� <> 0 Then
                cbo��������.SetFocus
            ElseIf cbo������.Enabled And cbo������.Visible Then
                cbo������.SetFocus
            ElseIf cboSex.Enabled And cboSex.Visible Then
                cboSex.SetFocus
            ElseIf Bill.Enabled Then
                Bill.SetFocus
            End If
            
            '����:39253
            If gbln���������ɿ� = False And blnPriceBill Or mstrYBPati <> "" Then
                If cbo���㷽ʽ.Enabled And cbo���㷽ʽ.Visible Then cbo���㷽ʽ.SetFocus
            End If
           
            
            If gbln���������ɿ� And blnPriceBill And mstrYBPati <> "" Then
                If cmdԤ����.Visible And cmdԤ����.Enabled Then
                    cmdԤ����.SetFocus
                End If
            End If
            
            If gbyt����ҽ�� <> 0 Then
                If blnPriceBill Then
                    If cbo��������.Enabled And cbo��������.Visible And cbo��������.ListIndex < 0 Then cbo��������.SetFocus
                Else
                    If cbo��������.Enabled And cbo��������.Visible Then cbo��������.SetFocus
                End If
            Else
                If blnPriceBill Then
                    If cbo������.Enabled And cbo������.Visible And cbo������.ListIndex < 0 Then cbo������.SetFocus
                Else
                    If cbo������.Enabled And cbo������.Visible Then cbo������.SetFocus
                End If
            End If
            
            Call ShowWelcomeByLed
            Call ReInitPatiInvoice
        End If
    Else
        mintInsure = 0: mcur������� = 0: mcur����͸֧ = 0
        Call InitBalanceGrid
        sta.Panels(Pan.C3�����ʻ�).Text = ""
        sta.Panels(Pan.C3�����ʻ�).Visible = False
        
        sta.Panels(Pan.C2��ʾ��Ϣ) = "�����֤���ɹ���"
        If Visible Then
            txtPatient.SetFocus
            Call txtPatient_GotFocus
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    If blnTran Then gcnOracle.RollbackTrans
    Call SaveErrLog
End Sub

Private Sub Set���˲��ѱ༭����()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ò��˲���ʱ�ı༭����
    '����:���˺�
    '����:2010-12-10 14:54:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mbln���� = False Then Exit Sub
    txtPatient.Enabled = False
    cbo��������.Enabled = False
    cboSex.Enabled = False
    IDKind.Enabled = False
    chkCancel.Visible = False
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim i As Long, lng����ID As Long, lng�Һſ��� As Long
    Dim strPati As String, blnIDCard As Boolean
    Dim blnCard As Boolean, blnICCard As Boolean, blnCancel As Boolean
    
    Dim int�ϴβ�����Դ As Integer
    Dim blnHavePriceBill As Boolean '��ǰ�Ƿ���ȡ�Ļ��۵�(���۵�ʱ,ֱ�ӽɿ�)
    Dim blnCheckReg As Boolean
    
    On Error GoTo errH
    blnHavePriceBill = False
    If KeyAscii = 13 And mblnValid = False Then
        mblnKeyReturn = True
    Else
        mblnKeyReturn = False
    End If
    
    '1.ҽ�������֤����:�����ﲡ���շ�ʱʹ��
    '--------------------------------------------------------------------------------------------------------------------
    If KeyAscii = 13 And mbytInFun = 0 And mbytInState = 0 And gint������Դ = 1 And Not mblnValid Then
        If txtPatient.Text = "" And chkCancel.Value = 0 And InStr(mstrPrivs, "�����շ�") > 0 Then
            Call MCPatientProcess
            Exit Sub
        End If
    End If
    If txtPatient.Locked Then Exit Sub '����״ֻ̬����ҽ���鿨
   
   '����:51488
    If (IDKind.Cards.������� = "�ո��" Or IDKind.Cards.������� = " ") And Chr(KeyAscii) = " " Then KeyAscii = 0: Exit Sub
   
    blnCheckReg = False
    '����:27364 ����:2010-01-13 15:27:50
    If mblnAutoChangePati And gint������Դ = 2 And (KeyAscii <> 13) Then
        '��Ҫ���ҵ�������Դ1��
        gint������Դ = 1: zlChangePatiSource (gint������Դ)
    End If
    
    
    '2.���ۺͼ��ʲ����벡��ֱ�ӻس�(���������������Ա�����):סԺ���˻����ݲ��ṩѡ����
    '--------------------------------------------------------------------------------------------------------------------
    If KeyAscii = 13 And Trim(txtPatient.Text) = "" And (mbytInFun = 1 Or mbytInFun = 2) Then
        '���ﻮ�۱�������˵�����Ϣ,���������Ϣ(�ɶ���������)
        If CheckBillsEmpty Then Bill.Active = True:  Call ClearBillRows
               
        Call ClearmobjBill '��������еĲ�����Ϣ
        Call ClearPatientInfo
        Call InitCommVariable
        Call ClearMoney
        If CheckBillsEmpty Then
            mstrPrePati = "": mlngPrePati = 0: mstrPreDoctor = ""
            Call ClearTotalInfo(True)
        Else
            Call ShowMoney
        End If
                    
        If mbytInFun = 1 Then
            If Not mblnValid And Visible Then
                If gint������Դ = 2 Then Exit Sub
                If gblnҽ�Ƹ��� Then
                    If cboҽ�Ƹ���.Enabled And cboҽ�Ƹ���.Visible Then cboҽ�Ƹ���.SetFocus
                Else
                    If gbyt����ҽ�� = 1 Then
                        If cbo��������.Enabled And cbo��������.Visible Then cbo��������.SetFocus
                    Else
                        If cbo������.Enabled And cbo������.Visible Then cbo������.SetFocus
                    End If
                End If
            End If
            
            KeyAscii = 0
            Exit Sub
        End If
    End If
        
       
    '3.�������벡��(�������ֱ�ʶ)����:סԺ�����շ�ʱ�ɵ���ѡ����
    '--------------------------------------------------------------------------------------------------------------------
    If KeyAscii = 13 And mbytInState = 0 And Trim(txtPatient.Text) = "" And Not mblnValid Then
        If mbytInFun = 2 Then
            lng����ID = SelectPatient
            If lng����ID = 0 Then Exit Sub
            txtPatient.Text = "-" & lng����ID
        ElseIf mbytInFun = 0 And gint������Դ = 2 Then
            frmPatiSelect.Show 1, Me
            If frmPatiSelect.mlngPatient = 0 Then Exit Sub
            txtPatient.Text = "-" & frmPatiSelect.mlngPatient
        End If
    End If
    
     
    If IDKind.GetCurCard.���� Like "����*" And Not mblnValid Then
        '103563,ֻҪ����ĵ�һ���ַ��ǡ�-+*����������ȫ���֣�����Ϊ����ˢ��
        If Not (InStr("-+*", Left(txtPatient.Text, 1)) > 0 And IsNumeric(Mid(txtPatient.Text, 2))) Then
            blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.ShowPassText)
        End If
    ElseIf IDKind.GetCurCard.���� = "�����" Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If InStr("0123456789-*+", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
        End If
    Else
        txtPatient.PasswordChar = IIf(IDKind.ShowPassText And Not mblnValid, "*", "")
        '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
        txtPatient.IMEMode = 0
    End If
    
    If blnCard And Len(txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 _
        Or KeyAscii = 13 And Trim(txtPatient.Text) <> "" Then
        If gint������Դ = 1 And mbytInFun = 0 And InStr(mstrPrivs, "�����ҽ������") = 0 Then
            txtPatient.Text = "":  Exit Sub
        End If
        
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii): txtPatient.SelStart = Len(txtPatient.Text)
        End If
        KeyAscii = 0
        
        '����δ�ı��˳�(ָδ����ǰ,����ָ�����շѣ���Ϊ�����շ�ʱmrsInfo����newbill�г�ʼ�˵�)
        If mrsInfo.State = 1 Then
            
            If txtPatient.Text = mrsInfo!���� Then
                If mblnValid Then Exit Sub
                mblnNotValied = True
                Call zlCommFun.PressKey(vbKeyTab): mblnNotValied = False: Exit Sub
            
            End If
            If mrsInfo!���� = "�²���" Then
                mobjBill.���� = txtPatient.Text
                mblnNotValied = True
                Call zlCommFun.PressKey(vbKeyTab): mblnNotValied = False: Exit Sub
            End If
        End If
        
        '���ﻮ�����������Ϣ(��������˵��ݺͲ�����Ϣ),�˴������������Ϣ,�ۼ���Ϣ�ں�������ȷ�Ϻ��ٴ���
        If mbytInFun = 1 Or mbytInFun = 2 And mbytBilling = 1 Then
            If CheckBillsEmpty Then Bill.Active = True:  Call ClearBillRows
        End If
 
        sta.Panels(Pan.C2��ʾ��Ϣ) = ""
        lblTotal.Caption = "�ϼ�:"
        
        '�շѱ��ֲ���ID
        If (mbytInFun = 0 Or mbytInFun = 1) And txtPatient.Text = mstrPrePati And mlngPrePati <> 0 Then
            strPati = "-" & mlngPrePati
        Else
            strPati = txtPatient.Text
        End If
        
        If IDKind.GetCurCard.���� Like "IC��*" And IDKind.GetCurCard.ϵͳ Then blnICCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
        If IDKind.GetCurCard.���� Like "*���֤*" And IDKind.GetCurCard.ϵͳ Then blnIDCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
        
        int�ϴβ�����Դ = gint������Դ
        
        '50200(��ֹ�����ҿ�����,����ʱ����Ǽ�ʱ�����ù���)
        If mbln���� And mstr���ת��ʱ�� <> "" Then
            txtDate.Text = Format(CDate(mstr���ת��ʱ��) - 1 / 24 / 60, "yyyy-mm-dd HH:MM:SS")
        Else
            txtDate.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
        End If
        If Not mobjBill Is Nothing Then mobjBill.����ʱ�� = CDate(txtDate.Text)
                
        'a.���������ȡ������Ϣʧ��
        If Not GetPatient(strPati, blnCancel, blnCard) Then
        
            Call InitBalanceGrid(True)
            If blnCancel Then 'ȡ������
                If Visible Then txtPatient.SetFocus
                txtPatient.Text = ""
                Exit Sub
            End If
            
            If blnCard Then
                MsgBox "����ȷ��" & gstrCustomerAppellation & "��Ϣ�������Ƿ���ȷˢ����", vbInformation, gstrSysName
                Call ClearPatientInfo(True)
                Exit Sub
            Else
                '�����շѡ����ۿ����ֶ����벡����Ϣ(����ʱ)��
                If gint������Դ = 1 And gblnInputName And (mbytInFun = 0 Or mbytInFun = 1) And IDKind.IDKind = IDKind.GetKindIndex("����") And txtPatient.Text <> "" Then
                    If mbytInFun = 0 And mbytInState = 0 And mstrInNO = "" Then
                        If Not CheckRegisted(0) Then
                           Call ClearPatientInfo(True): Exit Sub
                        End If
                    End If
                    If mbytInFun = 0 And mbytInState = 0 Then
                        '����:29283
                         '  -- ����:���ó���-1-�Һ�;2-�շ�
                         '  --        ����id_In-����ID(δ������,������)
                         '  --        ����_In: ˢ������;δˢ��ʱ,Ϊ��
                         '  --         ˢ����ʽ_In:  1-����ˢ��;2-ҽ��ˢ��
                         If zlPatiCardCheck(2, 0, IIf(blnCard Or blnICCard, txtPatient.Text, ""), 1) = False Then
                               Call ClearPatientInfo(True): Exit Sub
                         End If
                    End If
                    sta.Panels(Pan.C2��ʾ��Ϣ) = "����ı�ʶ���ܶ�ȡ" & gstrCustomerAppellation & "��Ϣ����Ĭ��Ϊ��" & gstrCustomerAppellation & "������"
                    Call ClearmobjBill
                    
                    If mbytInFun = 0 And mbytInState = 0 And Not mblnValid And Visible And mstrInNO = "" And txtIn.Text = "" Then
                        Call LoadAddedItem(0, txtPatient.Text)
                    End If
                    
                    If (mbytInFun = 0 Or mbytInFun = 1) And mobjBill.Pages(mintPage).NO = "" Then
                        cbo�ѱ�.Locked = False: cbo�ѱ�.TabStop = Not cbo�ѱ�.Locked And gbln�ѱ�
                        If Not mblnValid And Not (Bill.Active And txtPatient.Text = mstrPrePati And txtPatient.Text <> "") Then 'ͬһ�����˲�����
                            Call LoadAndSeek�ѱ�
                        End If
                    End If
                    cboҽ�Ƹ���.Locked = False
                    Call ShowPrePayInfo(False) 'Ԥ����Ϣ��ʼ
                    mobjBill.���� = txtPatient.Text
                    Call Set�����շѲ���(True)
                    
                    If txtPatient.Text = mstrPrePati And txtPatient.Text <> "" Then 'ͬһ���շѲ���,��ʱû�в���ID
                        mobjBill.���� = Trim(txt����.Text) & IIf(IsNumeric(txt����.Text), cbo���䵥λ.Text, "")
                        mobjBill.�Ա� = zlStr.NeedName(cboSex.Text)
                        mobjBill.�ѱ� = zlStr.NeedName(cbo�ѱ�.Text)
                                                
                        If Bill.Active Then
                            Call zlControl.CboSetIndex(cbo������.hWnd, cbo.FindIndex(cbo������, mstrPreDoctor, True)) '������click�¼�
                            Call cbo������_Click
                        End If
                        If Not mblnValid And Visible Then Bill.SetFocus
                        
                        Exit Sub
                    Else
                        '���ҽ��
                        If gbyt����ҽ�� = 0 And CheckBillsEmpty Then
                            For i = 1 To mobjBill.Pages.Count
                                mobjBill.Pages(i).��������ID = 0: mobjBill.Pages(i).������ = ""
                            Next
                            cbo������.ListIndex = -1: cbo��������.ListIndex = -1: lblDuty.Caption = ""
                        End If
                        
                        'ȡ����ҽ����Ϣ��ʼ,��ΪNewBill���ѳ�ʼ����
                                                           
                        Call ClearPatientInfo   '�������,�����,��ʼ���䵥λ
                        '���˺�:22343 gbln�ɿ������ΪgTy_Module_Para.byt�ɿ���� = 1
                        If Not (mbytInFun = 0 And gTy_Module_Para.byt�ɿ���� = 1) _
                            Or mstrPrePati = "" Then
                            Call ClearPayInfo
                            Call InitCommVariable
                            Call ClearMoney
                            If CheckBillsEmpty Then
                                mstrPrePati = "": mlngPrePati = 0: mstrPreDoctor = ""
                                Call ClearTotalInfo(True)
                            Else
                                Call ShowMoney
                            End If
                        End If
                        Call ReInitPatiInvoice
                        mblnNotValied = True
                        If Not mblnValid Then Call zlCommFun.PressKey(vbKeyTab)
                    mblnNotValied = False
                        If Not mblnValid Then Call ShowWelcomeByLed
                        Exit Sub
                    End If   'ͬһ���շѲ���
                    
                Else '���ʱ����в�����Ϣ
                    MsgBox "������������,���ܶ�ȡ" & gstrCustomerAppellation & "��Ϣ��", vbInformation, gstrSysName
                    Call ClearPatientInfo(True)
                    Exit Sub
                End If
            End If
            
        Else 'b.���������ȡ������Ϣ�ɹ�
            lng����ID = Val("" & mrsInfo!����ID)
            Call InitBalanceGrid(True)
            Call Set�����շѲ���
            
            If mbytInFun = 0 And mbytInState = 0 And mstrInNO = "" And gint������Դ = 1 Then
                If Not CheckRegisted(lng����ID) Then
                    Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "": txtPatient.SetFocus: Exit Sub
                End If
            End If
            If mbytInFun = 0 And mbytInState = 0 Then
                '����:29283
                 '  -- ����:���ó���-1-�Һ�;2-�շ�
                 '  --        ����id_In-����ID(δ������,������)
                 '  --        ����_In: ˢ������;δˢ��ʱ,Ϊ��
                 '  --         ˢ����ʽ_In:  1-����ˢ��;2-ҽ��ˢ��
                 If zlPatiCardCheck(2, lng����ID, IIf(blnCard Or blnICCard, txtPatient.Text, ""), 1) = False Then
                    '�ָ��ϴβ�����Դ
                    If int�ϴβ�����Դ <> gint������Դ And mTy_Para.blnסԺ���������շ� = False Then
                        '����:27364 ����:2010-01-13 15:27:50
                        Call zlChangePatiSource(int�ϴβ�����Դ)
                    End If
                    Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "": txtPatient.SetFocus: Exit Sub
                     Exit Sub
                 End If
            End If
            
            '���￨������
            If mbytInState = 0 And ((blnCard Or blnICCard Or blnIDCard Or IDKind.GetCurCard.�ӿ���� <> 0) And mstrPassWord <> "") Then
                i = Nvl(Choose(mbytInFun + 1, 3, 2, 4), 99)
                If Mid(gstrCardPass, i, 1) = "1" Then
                    If Not zlCommFun.VerifyPassWord(Me, mstrPassWord, mrsInfo!����, mrsInfo!�Ա�, "" & mrsInfo!����) Then
                        '�ָ��ϴβ�����Դ
                        If int�ϴβ�����Դ <> gint������Դ And mTy_Para.blnסԺ���������շ� = False Then
                            '����:27364 ����:2010-01-13 15:27:50
                            Call zlChangePatiSource(gint������Դ)
                        End If
                        Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "": txtPatient.SetFocus: Exit Sub
                    End If
                End If
            
            End If
                
            '�������ۻ��շ�ʱ,����ͬһ������ʱ������û�б���������Ϣ
            If Not IIf(mlngPrePati = 0, mstrPrePati = "" & mrsInfo!����, mlngPrePati = lng����ID) Then
                '���ҽ��
                If mbytInState = 0 And mstrInNO = "" Then
                    If gbyt����ҽ�� = 0 And CheckBillsEmpty Then
                        For i = 1 To mobjBill.Pages.Count
                            mobjBill.Pages(i).��������ID = 0: mobjBill.Pages(i).������ = ""
                        Next
                        cbo������.ListIndex = -1: cbo��������.ListIndex = -1: lblDuty.Caption = ""
                    End If
                End If
                
                Call ClearPatientInfo
                
                '���˺�:22343
                If Not (mbytInFun = 0 And gTy_Module_Para.byt�ɿ���� = 1) _
                    Or mstrPrePati = "" Then
                    Call ClearPayInfo
                    Call InitCommVariable
                    Call ClearMoney
                    If CheckBillsEmpty Then
                        mstrPrePati = "": mlngPrePati = 0: mstrPreDoctor = ""
                        Call ClearTotalInfo(True)
                    Else
                        Call ShowMoney
                    End If
                End If
            End If
                
            '�������뿪������
            '    �������ݲ��Ҹ��ݹҺŵ�����ʱ����ִ�в���ID
            If IsRegisterDept Then
                If IsNull(mrsInfo!����) Then 'û�н���,�����˺�,���ݹҺŵ��������˺Ϳ�������
                    Call SetDeptDoctorByRegevent(0, txtPatient.Text)
                    sta.Panels(Pan.C2��ʾ��Ϣ) = "�ò��˹Һ�ʱû�еǼǵ���,�����벡��������"
                    Call ClearPatientInfo(True)
                    
                    Set mrsInfo = New ADODB.Recordset
                    If Not mblnValid And Visible Then txtPatient.SetFocus
                    Exit Sub
                Else
                    Call Set�����˿�������Click(mrsInfo!ִ���� & "", Val("" & mrsInfo!ִ�в���ID))
                End If
            ElseIf gint������Դ = 2 Then
                If gbyt����ҽ�� <> 0 And mbytInState = 0 And mstrInNO = "" Then
                    '����������ʱ,ȡסԺ���˵Ŀ�������:����ȷ��ҽ������Զ�������
                    If mlngDeptID = 0 Then
                        Call zlControl.CboSetIndex(cbo��������.hWnd, cbo.FindIndex(cbo��������, Val(Nvl(mrsInfo!��ǰ����id))))
                    Else
                        Call zlControl.CboSetIndex(cbo��������.hWnd, cbo.FindIndex(cbo��������, mlngDeptID))
                    End If
                    Call cbo��������_Click
                End If
            ElseIf gint������Դ = 1 Then
                If mbytInState = 0 And mstrInNO = "" Then
                    If gbyt����ҽ�� <> 0 And mlng��ҳID <> 0 Then
                        If mlngDeptID = 0 Then
                            Call zlControl.CboSetIndex(cbo��������.hWnd, cbo.FindIndex(cbo��������, Val(Nvl(mrsInfo!��ǰ����id))))
                        Else
                            Call zlControl.CboSetIndex(cbo��������.hWnd, cbo.FindIndex(cbo��������, mlngDeptID))
                        End If
                        Call cbo��������_Click
                    Else
                        Call SetDeptDoctorByRegevent(lng����ID) '���ݲ��˹Һ���Ϣ���ÿ������Һ�ҽ��
                    End If
                End If
            End If
            
            If mbytInFun = 2 Then
                If Not IsNull(mrsInfo!������λ) Then
                    lblCorp.Visible = True: lblCorp.Caption = "������λ:" & mrsInfo!������λ
                Else
                    lblCorp.Visible = False: lblCorp.Caption = ""
                End If
            End If
             
            '����Ԥ������Ϣ
            If lng����ID <> 0 Then Call LoadFeeInfor(lng����ID)
            
            lbl����.Caption = "" & mrsInfo!��������
            txtPatient.Text = "" & mrsInfo!����
            txtPatient.PasswordChar = ""
            '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
            txtPatient.IMEMode = 0
            cboSex.ListIndex = cbo.FindIndex(cboSex, Nvl(mrsInfo!�Ա�), True)
            txt�����.Text = "" & mrsInfo!�����
            
            Call LoadOldData("" & mrsInfo!����, txt����, cbo���䵥λ)
            If Not IsNull(mrsInfo!��������) Then
                 txt����.Text = ReCalcOld(mrsInfo!��������, cbo���䵥λ, lng����ID)
            End If
            
            If glngSys Like "8??" Or mbytInFun = 2 Then
                cbo�ѱ�.ListIndex = cbo.FindIndex(cbo�ѱ�, Nvl(mrsInfo!�ѱ�), True)
                cbo�ѱ�.Locked = mbytInFun = 2: cbo�ѱ�.TabStop = Not cbo�ѱ�.Locked And gbln�ѱ�
            ElseIf Not mblnValid Then
                If IsRegisterDept And cbo��������.ListIndex <> -1 Then
                    cbo�ѱ�.ListIndex = cbo.FindIndex(cbo�ѱ�, Nvl(mrsInfo!�ѱ�), True) '�Һ�ʱȷ���ķѱ�
                Else
                    If mstrInNO = "" Then Call LoadAndSeek�ѱ�
                End If
            End If
            If gstr�ѱ� <> "" And cbo�ѱ�.ListIndex = -1 Then cbo�ѱ�.ListIndex = cbo.FindIndex(cbo�ѱ�, gstr�ѱ�, True)
            
            cboҽ�Ƹ���.ListIndex = cbo.FindIndex(cboҽ�Ƹ���, Nvl(mrsInfo!ҽ�Ƹ��ʽ), True)
            If mstr���ʽ <> "" And cboҽ�Ƹ���.ListIndex = -1 Then cboҽ�Ƹ���.ListIndex = cbo.FindIndex(cboҽ�Ƹ���, mstr���ʽ, True)
            cboҽ�Ƹ���.Locked = mbytInFun = 2 Or gint������Դ = 2
            

            '���ö����еĲ�����Ϣ
            With mobjBill
                .����ID = lng����ID
                .��ҳID = IIf(mbln���� And mlng��ҳID <> 0, mlng��ҳID, Nvl(mrsInfo!��ҳID, 0))
                .��ʶ�� = IIf(gint������Դ = 2, Nvl(mrsInfo!סԺ��, 0), Nvl(mrsInfo!�����, 0))
                .���� = "" & mrsInfo!����
                .�Ա� = "" & mrsInfo!�Ա�
                .���� = Trim(txt����.Text) & IIf(IsNumeric(txt����.Text), cbo���䵥λ.Text, "")
                .���� = "" & mrsInfo!��ǰ����
                .����ID = IIf(mbln���� And mlngUnitID <> 0, mlngUnitID, Val(Nvl(mrsInfo!��ǰ����ID)))
                .����ID = IIf(mbln���� And mlngDeptID <> 0, mlngDeptID, Val(Nvl(mrsInfo!��ǰ����id)))
                .�ѱ� = zlStr.NeedName(cbo�ѱ�.Text) '�Ե�ǰ��ЧΪ׼
            End With
            Call ReInitPatiInvoice
            
            '������������
            If Not mblnValid And Visible Then
                If mbytInFun = 0 Or mbytInFun = 1 Then
                    '����ͬһ������ʱ
                    If Not (IIf(mlngPrePati = 0, mstrPrePati = mobjBill.����, mlngPrePati = mobjBill.����ID) And txtPatient.Text <> "") Then
                         Call AddCardFee '�������￨������
                    End If
                    
                    '��ȡ���˵Ķ��Ż��۵�
                    If mbytInFun = 0 And mbytInState = 0 And mstrInNO = "" And txtIn.Text = "" Then
                        If mobjBill.����ID <> 0 Then
                            If gblnCheckRegeventDept And gint������Դ = 1 And IsRegisterDept Then lng�Һſ��� = Val("" & mrsInfo!ִ�в���ID)
                           blnHavePriceBill = LoadMultiBills(mobjBill.����ID, InStr(1, mstrPrivs, "��ͨ���˶൥���շ�") = 0, lng�Һſ���, blnCard)
                        End If
                        Call LoadAddedItem(mobjBill.����ID, mobjBill.����)
                    End If
                End If
                '��궨λ
                If mstrInNO = "" Then
                    If mbytInFun = 0 And mbytInState = 0 And txtPatient.Text = "�²���" Then
                        txtPatient.SetFocus
                        Call txtPatient_GotFocus
                    Else
                        If cboҽ�Ƹ���.ListIndex = -1 And gblnҽ�Ƹ��� And mbytInFun <> 2 Then '���ʲ�������ķѱ𡢸��ʽ
                            If cboҽ�Ƹ���.Enabled And cboҽ�Ƹ���.Visible Then cboҽ�Ƹ���.SetFocus
                        Else
                            '����:39253
                            If gbln���������ɿ� = False And blnHavePriceBill Then
                                If cbo���㷽ʽ.Enabled And cbo���㷽ʽ.Visible Then cbo���㷽ʽ.SetFocus
                            End If
                            
                            If gbyt����ҽ�� = 0 Then
                                If blnHavePriceBill Then
                                    If cbo������.Enabled And cbo������.Visible And cbo������.ListIndex < 0 Then cbo������.SetFocus
                                Else
                                    If cbo������.Enabled And cbo������.Visible Then cbo������.SetFocus
                                End If
                            ElseIf glngSys Like "8??" Then
                                Bill.SetFocus
                            Else
                                If blnHavePriceBill Then
                                    If cbo��������.Enabled And cbo��������.Visible And cbo��������.ListIndex < 0 Then cbo��������.SetFocus
                                Else
                                    If cbo��������.Enabled And cbo��������.Visible Then cbo��������.SetFocus
                                End If
                            End If
                        End If
                    End If
                    
                    Call ShowWelcomeByLed
                End If
            End If
        End If
    End If

    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub AddCardFee()
'����:�������￨������
    Dim objDetail As Detail, lngDoUnit As Long
        
    If mbytInFun = 0 And mstrCardNO = "" And Bill.Active Then
        Set objDetail = ReadPatiCardObj(mobjBill.����ID, mstrCardNO)
        
        If mstrCardNO <> "" And Not objDetail Is Nothing Then
            If Not ItemExist(objDetail.ID) Then
                If mobjBill.Pages(mintPage).Details.Count >= Bill.Rows - 1 Then
                    Bill.Rows = Bill.Rows + 1
                    mblnNewRow = True: Call bill_AfterAddRow(Bill.Rows - 1): mblnNewRow = False
                End If
                Bill.TextMatrix(Bill.Rows - 1, BillCol.���) = "" '�б�Ҫ����
                
                lngDoUnit = mobjBill.����ID
                If lngDoUnit = 0 Then lngDoUnit = Get��������ID
                
                lngDoUnit = Get�շ�ִ�п���ID(objDetail.���, objDetail.ID, objDetail.ִ�п���, lngDoUnit, Get��������ID, _
                            gint������Դ, , , , , mobjBill.����ID)
                
                Call SetDetail(objDetail, Bill.Rows - 1, lngDoUnit)
                Call CalcMoneys(mintPage, Bill.Rows - 1)
                Call ShowDetails(Bill.Rows - 1)
                Call ShowMoney
            End If
        End If
    End If
End Sub


Private Sub ShowWelcomeByLed()
'����:��ʾ��ӭ��Ϣ�Ͳ�����Ϣ
    Dim strInfo As String, lngPatient As Long

    If mbytInFun = 0 And mbytInState = 0 And gblnLED Then
        If gblnLedWelcome Then
            zl9LedVoice.Reset com
            zl9LedVoice.Speak "#1"
            zl9LedVoice.Init UserInfo.��� & " �շ�ԱΪ������", mlngModul, gcnOracle
        End If
        
        strInfo = Trim(txtPatient.Text)
        If mrsInfo.State = 1 Then strInfo = strInfo & " " & mrsInfo!�Ա� & " " & mrsInfo!����: lngPatient = Val("" & mrsInfo!����ID)
        zl9LedVoice.DisplayPatient strInfo, lngPatient
    End If
End Sub
Private Function GetPatient(ByVal strInput As String, Optional blnCancel As Boolean, Optional ByVal blnCard As Boolean, Optional blnYbCheckCard As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ
    '���:blnCancel=���ڱ�ʾ����ȡ��
    '       blnCard=��ʾ�Ƿ���￨ˢ��
    '       blnYbCheckCard-ҽ�������鿨(24689)
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-03 16:43:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, lng�����ID As Long, strPassWord As String, strErrMsg As String
    Dim lng����ID As Long, blnHavePassWord As Boolean
    Dim strMoney As String, strWhere As String, strPati As String
    Dim rsTmp As ADODB.Recordset, strTemp As String, strTempYb As String
    Dim bln�Һ� As Boolean
    Dim vRect As RECT
    
    bln�Һ� = False: mblnNotClearLedDisplay = False
    mlngPreBrushCard = 0: mlngCardTypeID = 0
    
ReDO:
    blnCancel = False
    mstrWarn = "" '���ʷ��౨��
    If mbytInFun = 2 Then
        strMoney = "zl_PatiDayCharge(A.����ID) as ���ն�, Zl_Patiwarnscheme(A.����id) As ���ò���,"
    End If
    
    '���벡�˵�Ȩ��
    If mbytInState = 0 And mstrInNO = "" Then  '����
        If gint������Դ = 2 Then
            strWhere = " And Nvl(A.��ǰ����ID,0)<>0"
        End If
    End If
    
    '��ȡ������Ϣ
    If mbln���� And mlng��ҳID <> 0 Then
        strSQL = _
        " Select " & strMoney & "Decode(Sign(A.����ʱ��-A.�Ǽ�ʱ��),0,1,0) as ����,A.����ID," & _
        "        nvl(b1.��������,A.��������) as ��������,A.����," & _
        "        Nvl(b1.��ҳID,A.��ҳID) as ��ҳID," & _
        "        A.IC����,A.���￨��,A.����֤��,A.�����,Nvl(B1.סԺ��,A.סԺ��) As סԺ��,nvl(B1.����,A.����) as ����," & _
        "        nvl(b1.�Ա�,A.�Ա�) as �Ա�,nvl(b1.����,A.����) as ����,C.���� As ��������, A.��������," & _
        "        nvl(b1.�ѱ�,A.�ѱ�) as �ѱ�,A.������,nvl(b1.ҽ�Ƹ��ʽ,A.ҽ�Ƹ��ʽ) as ҽ�Ƹ��ʽ," & _
        "        A.������λ,nvl(b1.��ǰ����ID,A.��ǰ����ID) as ��ǰ����ID,nvl(b1.��Ժ����ID,A.��ǰ����ID) as ��ǰ����ID," & _
        "        nvl(b1.��Ժ����,A.��ǰ����) as ��ǰ����,A.��Ժ," & _
        "        Decode(B1.��������,NULL,0,1,1,0) as ����,B1.��Ժ����" & _
        " From ������Ϣ A,������ҳ B1,������� C  " & _
        " Where A.���� = C.���(+) And A.����ID=B1.����ID(+) And B1.��ҳID = [4] And A.ͣ��ʱ�� is NULL"
    Else
        strSQL = _
        " Select " & strMoney & "Decode(Sign(A.����ʱ��-A.�Ǽ�ʱ��),0,1,0) as ����,A.����ID,A.��������,A.����," & _
                 IIf(gint������Դ = 1, "NULL", "Decode(A.��ǰ����ID,NULL,NULL,A.��ҳID)") & " as ��ҳID,A.IC����,A.���￨��,A.����֤��,A.�����,A.סԺ��,A.����," & _
        "        A.�Ա�,A.����,C.���� ��������, A.��������,A.�ѱ�,A.������,A.ҽ�Ƹ��ʽ,A.������λ,A.��ǰ����ID,A.��ǰ����ID,A.��ǰ����,A.��Ժ," & _
        "        decode(B1.��������,NULL,0,1,1,0) as ����,B1.��Ժ����" & _
        " From ������Ϣ A,������ҳ B1,������� C  " & _
        " Where A.���� = C.���(+) And A.����ID=B1.����ID(+) And A.��ҳID=B1.��ҳID(+) And A.ͣ��ʱ�� is NULL"
    End If

    If blnYbCheckCard = False And blnCard And IDKind.GetCurCard.���� Like "����*" And InStr("-+*", Left(strInput, 1)) = 0 Then   '103563
        If gint������Դ = 1 And Not gblnInputCard Then
            Set mrsInfo = New ADODB.Recordset
            Exit Function
        End If
      
        If IDKind.Cards.��ȱʡ������ And Not IDKind.GetfaultCard Is Nothing Then
            lng�����ID = IDKind.GetfaultCard.�ӿ����
        Else
            lng�����ID = "-1"
        End If
        
        '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�);��
        If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg, lng�����ID) = False Then GoTo NotFoundPati:
        If lng����ID <= 0 Then GoTo NotFoundPati:
        mlngCardTypeID = lng�����ID
        strInput = "-" & lng����ID
        blnHavePassWord = True
        strSQL = strSQL & strWhere & " And A.����ID=[1] "
        mlngPreBrushCard = lng�����ID
        
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Or blnYbCheckCard Then '����ID
        If gint������Դ = 1 And (Not gblnInputID And mstrYBPati = "") _
            And Not (mstrInNO <> "" And mbytInState = 0) Then
            Set mrsInfo = New ADODB.Recordset
            Exit Function
        End If
        strSQL = strSQL & strWhere & " And A.����ID=[1]"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '�����
        If gint������Դ = 1 And Not gblnInputID Then
            Set mrsInfo = New ADODB.Recordset
            Exit Function
        End If
        strSQL = strSQL & strWhere & " And A.�����=[1]"
        '75087,Ƚ����,2014-7-15,���ﲡ���շ�ʱ,����Ҫ���������������,ֻ��Ҫ��������ŵ����˳��ż����ҵ��������Ĳ�����Ϣ������
        strInput = "*" & zlCommFun.GetFullNO(Mid(strInput, 2), 3)
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then 'סԺ��
        If gint������Դ = 1 And Not gblnInputID Then
            Set mrsInfo = New ADODB.Recordset
            Exit Function
        End If
        strSQL = strSQL & strWhere & " And A.����ID = (Select Max(����id) From ������ҳ Where סԺ�� = [1])"
    ElseIf Left(strInput, 1) = "." Then '�Һŵ���(���Ϊִ�в���ID������)
        If gint������Դ = 1 And (mbytInFun = 0 Or mbytInFun = 1) And Not gblnInputNO Then
            Set mrsInfo = New ADODB.Recordset
            Exit Function
        End If
        bln�Һ� = True
        '���ջ���˳���Ź���
        strInput = UCase(GetFullNO(Mid(strInput, 2), 12))
        txtPatient.Text = strInput
        
        '�������ʱ����Ҫ�ҺŽ���
        '����ǳ�Ժ����,��ͨ��������ҳIDΪ0��Ϊ����������ͨ������ʶ�����������,ע��:���һ���ֶ�ִ�в���ID��patient_keypress�л��õ�
        '76451,Ƚ����,2014-8-19
        strSQL = "" & _
            "   Select " & strMoney & "Decode(Sign(A.����ʱ��-A.�Ǽ�ʱ��),0,1,0) as ����,A.����ID,A.��������,A.����," & _
                                IIf(gint������Դ = 1, "NULL", "Decode(A.��ǰ����ID,NULL,NULL,A.��ҳID)") & " as ��ҳID,A.���￨��,A.����֤��,Nvl(B.��ʶ��,A.�����) as �����," & _
            "               A.סԺ��,B.����,B.�Ա�,B.����,C.���� ��������, A.��������,B.�ѱ�,A.������,A.ҽ�Ƹ��ʽ,A.������λ,A.��ǰ����ID,A.��ǰ����ID,A.��ǰ����,B.ִ����,B.ִ�в���ID,A.��Ժ," & _
            "               decode(B1.��������,NULL,0,1,1,0) as ����,B1.��Ժ����" & _
            " From ������Ϣ A,������ҳ B1,������ü�¼ B,������� C " & _
            " Where B.����ID=A.����ID" & IIf(mbytInFun = 2, "", "(+)") & _
            "            And A.����ID=B1.����ID(+) And A.��ҳID=B1.��ҳID(+)  " & _
            "           And A.���� = C.���(+) And B.��¼����=4 And B.��¼״̬=1 " & _
            zlGetRegEventsCons("�Ӱ��־", "B") & _
            strWhere & " And B.NO=[2] And Rownum<2"
    Else
        If mrsInfo.State = 1 Then
            If mrsInfo!���� = strInput Then GetPatient = True: Exit Function
        End If
        mlngCardTypeID = IDKind.GetCurCard.�ӿ����
        Select Case IDKind.GetCurCard.����
            Case "����", "��������￨"
                'ͨ������ģ�����Ҳ���(�������벡�˱�ʶʱ)
                If Not mblnValid And gblnSeekName And (mbytInFun <> 2 And gblnInputID Or mbytInFun = 2) Then
                    strPati = _
                        " Select /*+Rule */1 as ����ID,A.����ID as ID,A.����ID,A.����,A.�Ա�,A.����," & _
                                    IIf(gint������Դ = 2 And mbytInFun <> 2, "A.סԺ��,B.���� as ����,A.��ǰ���� as ����,", "A.�����,") & _
                        "           A.��������,A.���֤��,A.��ͥ��ַ,A.������λ" & _
                        " From ������Ϣ A,���ű� B" & _
                        " Where A.ͣ��ʱ�� is NULL And A.��ǰ����ID=B.ID(+) And Rownum <101 " & strWhere & " And A.���� Like [1]" & _
                        IIf(gintNameDays = 0, "", " And Nvl(A.����ʱ��,A.�Ǽ�ʱ��)>Trunc(Sysdate-[2])")
                    
                    If mbytInFun = 2 And gblnOnlyUnitPatient Then
                        strPati = strPati & " And A.��ͬ��λID Is Not Null"
                    End If
                    
                    '���ﲡ���շ�ʱ���Բ���Ӧ���˵���
                    If gint������Դ = 1 And mbytInFun <> 2 Then
                        strPati = strPati & " Union ALL " & _
                            "Select 0,0 as ID,-NULL,'[�²���]',NULL,NULL,-NULL,To_Date(NULL),NULL,NULL,NULL From Dual"
                    End If
                    strPati = strPati & " Order by ����ID,����"
                        
                    vRect = zlControl.GetControlRect(txtPatient.hWnd)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strPati, 0, "����" & mbytInFun & gint������Դ, 1, "", "��ѡ����", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, False, True, strInput & "%", gintNameDays, "bytSize=1")
                    If Not rsTmp Is Nothing Then
                        If rsTmp!ID = 0 Then '�����²���
                            strSQL = ""
                        Else '�Բ���ID��ȡ
                            '85187,Ƚ����,2015-05-27,��Ժ���������շ�ʱ����ģ�������Ҳ���������Ϣ��������Դ���õ���"���ﲡ��"��
                            strInput = "-" & rsTmp!����ID
                            strSQL = strSQL & strWhere & " And A.����ID=[1]"
                        End If
                    Else 'ȡ��ѡ��
                        strSQL = ""
                    End If
                Else
                    strSQL = ""
                End If
            Case "ҽ����"
                strInput = UCase(strInput)

                If MCPAR.blnOnlyBjYb And zlCommFun.ActualLen(strInput) >= 9 Then
                    '������ҽ������Ч:������:����:27331
                    strSQL = strSQL & strWhere & "  And A.ҽ���� like [3] "
                    strTemp = Left(strInput, 9) & "%"
                Else
                     strSQL = strSQL & strWhere & "  And A.ҽ����=[2]"
                End If
                
                'strSQL = strSQL & strWhere & " And A.ҽ����=[2]"
            Case "���֤��", "���֤", "�������֤"
                strInput = UCase(strInput)
                 If gobjSquare.objSquareCard.zlGetPatiID("���֤", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                 strInput = "-" & lng����ID
                 blnHavePassWord = True
                strSQL = strSQL & strWhere & " And   A.����ID=[1]"
            Case "IC����", "IC��"
                strInput = UCase(strInput)
                If gobjSquare.objSquareCard.zlGetPatiID("IC��", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                strInput = "-" & lng����ID
                strSQL = strSQL & strWhere & " And A.����ID=[1]"
               blnHavePassWord = True
            Case "�����"
                If gint������Դ = 1 And Not gblnInputID Then
                    Set mrsInfo = New ADODB.Recordset
                    Exit Function
                End If
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = strSQL & strWhere & " And A.�����=[2]"
                '75087,Ƚ����,2014-7-15,���ﲡ���շ�ʱ,����Ҫ���������������,ֻ��Ҫ��������ŵ����˳��ż����ҵ��������Ĳ�����Ϣ������
                strInput = zlCommFun.GetFullNO(strInput, 3)
            Case "סԺ��"
                If gint������Դ = 1 And Not gblnInputID Then
                    Set mrsInfo = New ADODB.Recordset
                    Exit Function
                End If
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = strSQL & strWhere & " And A.����ID = (Select Max(����id) From ������ҳ Where סԺ�� = [2])"
            Case Else
                '��������,��ȡ��صĲ���ID
                If IDKind.GetCurCard.�ӿ���� > 0 Then
                    lng�����ID = IDKind.GetCurCard.�ӿ����
                    If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                    If lng����ID = 0 Then GoTo NotFoundPati:
                    mlngPreBrushCard = lng�����ID
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(IDKind.GetCurCard.����, strInput, False, lng����ID, _
                        strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                End If
                
                If lng����ID <= 0 Then GoTo NotFoundPati:
                strSQL = strSQL & strWhere & " And A.����ID=[1]"
                strInput = "-" & lng����ID
                blnHavePassWord = True
        End Select
    End If
        
    On Error GoTo errH
    If strSQL <> "" Then
        Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput, strTemp, mlng��ҳID)
        If Not mrsInfo.EOF Then
            '75259�����ϴ�,2014-7-10��������������ʾ��ɫ����
            Call SetPatiColor(txtPatient, Nvl(mrsInfo!��������), IIf(IsNull(mrsInfo!����), Me.ForeColor, vbRed))
            If gint������Դ = 1 And mTy_Para.blnסԺ���������շ� = False Then
                '��Ҫ����Ƿ�Ϊ��Ժ����
                '����:27364 ����:2010-01-13 15:27:50
                If Val(Nvl(mrsInfo!��Ժ)) = 1 Then
                        If gbln������Դ��Ȩ�޿��� And InStr(1, mstrPrivs, ";��������;") = 0 Then
                            '29720
                            '����ת������
                            Call MsgBox("�ò�������Ժ����,���ܽ����շ�(���ۻ����)����!)", vbOKCancel + vbInformation + vbDefaultButton1, gstrSysName)
                            Set mrsInfo = New ADODB.Recordset
                            Exit Function
                        End If
                    '��Ϊ��Ժ����,�Զ�����Ժ״̬
                    mblnAutoChangePati = True
                    gint������Դ = 2: Call zlChangePatiSource(gint������Դ)
                    Set mrsInfo = New ADODB.Recordset
                     GoTo ReDO:
                End If
                strWhere = ""
            End If
            '���쳣���ݽ����շ�
            If PatiErrBillPay(Val(Nvl(mrsInfo!����ID))) Then
                Call ClearBillRows: Call ClearMoney
                Call ClearTotalInfo(True)
                NewBill True
                blnCancel = True
                Exit Function
            End If
            If mlng����ID <> mrsInfo!����ID Then mlng����ҽ�� = 0
            
            GetPatient = True
            mstrPassWord = strPassWord
            If Not blnHavePassWord Then mstrPassWord = Nvl(mrsInfo!����֤��)
        Else
            Set mrsInfo = New ADODB.Recordset
            If bln�Һ� Then
                 txtPatient.Text = "": GetPatient = False
            End If
        End If
    Else
        Set mrsInfo = New ADODB.Recordset
    End If
    Exit Function
NotFoundPati:
    Set mrsInfo = New ADODB.Recordset
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Set mrsInfo = New ADODB.Recordset
End Function

Private Sub txtPatient_LostFocus()
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(False)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(False)
    '����:60010
    IDKind.SetAutoReadCard (False)
    zlCommFun.OpenIme False
    If (mbytInFun = 0 Or mbytInFun = 1) And mbytInState = 0 And Trim(txtPatient.Text) <> "" Then
        mobjBill.���� = txtPatient.Text
        mobjBill.���� = Trim(txt����.Text) & IIf(IsNumeric(txt����.Text), cbo���䵥λ.Text, "")
        mobjBill.�Ա� = zlStr.NeedName(cboSex.Text)
    End If
End Sub

Private Sub txtPatient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txtPatient.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtPatient_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txtPatient_Validate(Cancel As Boolean)
    If mblnKeyReturn = False Then
        mblnValid = True: Call txtPatient_KeyPress(13): mblnValid = False
    Else
        mblnKeyReturn = False
    End If
End Sub

Private Sub txtRePrint_GotFocus()
    Call zlControl.TxtSelAll(txtRePrint)
End Sub

Private Sub txtRePrint_KeyPress(KeyAscii As Integer)
    Dim strNos As String, strOper As String, vDate As Date, intInsure As Integer, blnVirtualPrint As Boolean
    Dim lng����ID As Long, lng����ID As Long
    Dim strReclaimInvoice As String, intInvoiceFormat As Integer '���յ�Ʊ��
    
    If KeyAscii > 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))

    '��һλ����������ĸ,����λ����
    If KeyAscii <> 13 Then
        Call SetNOInputLimit(txtRePrint, KeyAscii)
    Else
        '�ش�
        txtRePrint.Text = GetFullNO(txtRePrint.Text, 13)
        zlControl.TxtSelAll txtRePrint
        
        '�Ƿ���ת������ݱ���
        If zlDatabase.NOMoved("������ü�¼", txtRePrint.Text, , "1", Me.Caption) Then
            If Not ReturnMovedExes(txtRePrint.Text, 1, Me.Caption) Then Exit Sub
            mblnNOMoved = False
        End If
        
        If Not ReadBillInfo(1, txtRePrint.Text, 1, strOper, vDate, lng����ID) Then
            txtRePrint.SetFocus: Exit Sub
        End If
        If InStr(mstrPrivs, "���в���Ա") <= 0 Then
            If UserInfo.���� <> strOper Then
                MsgBox "��û��""���в���Ա""Ȩ��,�����ش�" & strOper & "�ĵ��ݣ�", vbInformation, gstrSysName
                txtRePrint.Text = "": Exit Sub
            End If
        End If
        If Not BillOperCheck(2, strOper, vDate, "�ش�", txtRePrint.Text, , 1) Then
            txtRePrint.SetFocus: Exit Sub
        End If
        
        '�����Ƕ൥���շ��е�һ��
        If gTy_Module_Para.bytƱ�ݷ������ <> 0 Then
            strNos = GetMultiNOs(txtRePrint.Text, , , True)
        Else
            strNos = GetMultiNOs(txtRePrint.Text)
        End If
        
        '������ʣ�������Ĳſ����ش�
        If Not BillExistMoney(strNos, 1, True) Then
            MsgBox "���ݲ����ڻ��Ѿ�ȫ���˷�,�����ش�", vbInformation, gstrSysName
            txtRePrint.Text = "": Exit Sub
        End If
        '������ҽ��������㣬�������ش�Ͳ���
        If CheckBillExistReplenishData(1, , Replace(strNos, "'", "")) = True Then
            MsgBox "��ǰ���ݽ�����ҽ��������㣬�������ش�Ʊ�ݣ�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '�����ش�ĵ�����ʾ
        If frmMultiBills.ShowMe(Me, 0, mstrPrivs, txtRePrint.Text, "", True) = False Then Exit Sub
        intInsure = ChargeExistInsure(txtRePrint.Text, , lng����ID)
        If intInsure <> 0 Then
            blnVirtualPrint = gclsInsure.GetCapability(supportҽ���ӿڴ�ӡƱ��, lng����ID, intInsure, CStr(lng����ID))
            '�˴�ֻ�ṩ���շ�Ʊ�ݵ��ش�
        End If
        
        Call ReInitPatiInvoice(True, intInsure, lng����ID)
        strReclaimInvoice = zlGetReclaimInvoice(txtRePrint.Text)
        If strReclaimInvoice <> "" Then
            '��Ҫ��ʾ��������Ҫ���յķ�Ʊ
            If MsgBox("ע��:" & vbCrLf & " ��ע��������·�Ʊ:" & vbCrLf & strReclaimInvoice, vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                Call RefreshFact 'ˢ��Ʊ�ݺ�
                txtRePrint.Text = ""
                txtPatient.SetFocus
                Exit Sub
            End If
        End If
        intInvoiceFormat = IIf(strReclaimInvoice = "" And gTy_Module_Para.bytƱ�ݷ������ <> 0, mintOldInvoiceFormat, mintInvoiceFormat)
        Dim strPriceGrade As String
        If gintPriceGradeStartType >= 2 Then
            strPriceGrade = GetPriceGradeFromNos(strNos)
        Else
            strPriceGrade = mstr��ͨ�۸�ȼ�
        End If
        If Not RePrintCharge(1, strNos, Me, mlng����ID, strReclaimInvoice, , , _
            intInvoiceFormat, blnVirtualPrint, , mlngShareUseID, mstrUseType, , strPriceGrade) Then
            txtRePrint.SetFocus
        Else
        
            '��ҽһ��ͨд����85950
            Call WriteInforToCard(Me, mlngModul, mstrPrivs, gobjSquare.objSquareCard, 0, strNos)
            
            Call RefreshFact 'ˢ��Ʊ�ݺ�
            txtRePrint.Text = ""
            txtPatient.SetFocus
        End If
    End If
End Sub

Private Sub txtRePrint_LostFocus()
    txtRePrint.BackColor = vbWhite
End Sub
Public Function GetMustPaySum() As Currency
'���ܣ��󱾴��շѵ�Ӧ�ɺϼƣ���Ҫ���ڶ൥���շ�ģʽ
    Dim curMoney As Currency, i As Integer
    For i = 1 To mobjBill.Pages.Count
        curMoney = curMoney + mobjBill.Pages(i).Ӧ�ɽ��
    Next
    GetMustPaySum = curMoney
End Function

Private Function Get��ҩ����(ByRef str���㵥λ As String) As Long
'���ܣ�ȡ��ǰ��������ҩ��������������ڲ�ͬ��λ��ҩƷ���򷵻�Ϊ0
    Dim i As Integer, str��λ As String
    
    Get��ҩ���� = 0
    With mobjBill.Pages(mintPage)
        For i = 1 To .Details.Count
            If .Details(i).�շ���� = "7" Then
                If gblnҩ����λ Then
                    If str��λ <> "" And str��λ <> .Details(i).Detail.ҩ����λ Then
                        str��λ = "��ͬ��λ"
                        Exit For
                    Else
                        If str��λ = "" Then str��λ = .Details(i).Detail.ҩ����λ
                    End If
                Else
                    If str��λ <> "" And str��λ <> .Details(i).���㵥λ Then
                        str��λ = "��ͬ��λ"
                        Exit For
                    Else
                        If str��λ = "" Then str��λ = .Details(i).���㵥λ
                    End If
                End If
                
                Get��ҩ���� = Get��ҩ���� + .Details(i).���� * .Details(i).����
            End If
        Next
    End With
    If str��λ = "��ͬ��λ" Then
        Get��ҩ���� = 0
    Else
        str���㵥λ = str��λ
    End If
End Function

Private Sub AutoBultBookFee()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Զ����ɹ����ѻ��Զ��ֵ���
    '����:���˺�
    '����:2011-08-16 10:34:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim CurOneCard As Currency
   ' If txt�ɿ�.Tag = "�˳�" Then txt�ɿ�.Tag = "": Exit Sub
    If mbytInFun = 0 And mbytInState = 0 And mstrInNO = "" And gbytAutoSplitBill > 0 And Not (mstrYBPati <> "" And MCPAR.����Ԥ����) Then
        Call AutoSplitBill
    End If
    '�շ�ʱ�Զ�������������Ŀ:�޸�ʱ���ܹ�����
    If mbytInFun = 0 And mbytInState = 0 And gTy_Module_Para.bln������ Then
        If Not CheckBillsEmpty Then Call SetFactMoney
    End If
End Sub
 

Private Sub CalcMoneys(Optional intPage As Integer, Optional lngRow As Long)
'���ܣ���������¼���ָ���л������еĽ��
'������intPage,lngRow=ָ������ҳָ����,Ϊ0��ʾ����������
'˵����ExpenseBill���ϵ�������Ӧ���ݵ��к�
    Dim i As Long, p As Integer
    Dim strMainRows As String
    Dim bln��������ۿ� As Boolean
        
    'If CheckBillsEmpty Then Exit Sub   '�˴��������ж�,�ڸı�ѱ�(�������ʲ��˱䶯)�͸ı�Ӱ�״̬,����ʱ�ж�
    
    Screen.MousePointer = 11
    For p = IIf(intPage = 0, 1, intPage) To IIf(intPage = 0, mobjBill.Pages.Count, intPage)
        strMainRows = ""
        If mobjBill.Pages.Count >= p Then
            For i = IIf(lngRow = 0, 1, lngRow) To IIf(lngRow = 0, mobjBill.Pages(p).Details.Count, lngRow)
                If mobjBill.Pages(p).Details.Count >= i Then
                    
                    bln��������ۿ� = False
                    If gbln��������ۿ� Then                    '����������ηѱ�,����ܼ����ۿ۲�����Ч,�����ܼ���
                        If mobjBill.Pages(p).Details(i).�������� > 0 Then    '����
                            bln��������ۿ� = Not mobjBill.Pages(p).Details(mobjBill.Pages(p).Details(i).��������).Detail.���ηѱ�
                            If bln��������ۿ� And lngRow <> 0 Then strMainRows = strMainRows & "," & mobjBill.Pages(p).Details(i).��������      '��������һ�е�ʱ��
                        Else
                            If CheckMainItem(i, p) Then                          '����������
                                 bln��������ۿ� = Not mobjBill.Pages(p).Details(i).Detail.���ηѱ�
                                 If bln��������ۿ� Then strMainRows = strMainRows & "," & i  'һҳ�����ж��������,�ȼ�¼�����к�,���������������ۿ�
                            End If
                        End If
                    End If
                            
                    Call CalcMoney(p, i, bln��������ۿ�)
                End If
            Next
        
            '������������
            If gbln��������ۿ� Then
                For i = 1 To UBound(Split(strMainRows, ","))
                    Call CalcPItemActualIncome(Split(strMainRows, ",")(i), p)
                Next
            End If
        End If
    Next
    
    Screen.MousePointer = 0
End Sub

Private Sub CalcMoney(intPage As Integer, lngRow As Long, Optional bln��������ۿ� As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������¼���ָ���еĽ��
    '���:intPage=ָ��ҳ����,lngRow=ָ����
    '����:���˺�
    '����:2014-06-06 18:02:30
    '˵����1.ExpenseBill���ϵ�������Ӧ���ݵ��к�
    '      2.���ֻ�ܶ�Ӧһ��������Ŀ:mobjBill.Pages(intPage).Details(lngRow).InComes(1)
    '      3.������ϸĿδ�����������Ŀ(��һ�μ���),��ʹ��Ĭ���ּ�
    '      4.������ϸĿ�Ѿ������������Ŀ(����2��),���ֶ�����(Ҳ����δ��)�˵���,�򰴸õ��ۼ��㡣
    
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strInfo As String, strAdvance As String
    Dim rsTmp As ADODB.Recordset
    Dim dblMoney As Double '�û�����ı�۽��
    Dim str�ѱ� As String
    Dim dblAllTime As Double, dbl�Ӱ�Ӽ��� As Double
    Dim rsPrice As ADODB.Recordset, strPrice As String, varPrice As Variant, dblʣ������ As Double
    Dim strPriceGrade As String, strWherePriceGrade As String
    
    On Error GoTo errH
    
    If mobjBill.Pages.Count < intPage Then Exit Sub
    If mobjBill.Pages(intPage).Details.Count < lngRow Then Exit Sub
    
    If InStr(",5,6,7,", mobjBill.Pages(intPage).Details(lngRow).�շ����) > 0 Then
        strPriceGrade = mstrҩƷ�۸�ȼ�
    ElseIf mobjBill.Pages(intPage).Details(lngRow).�շ���� = "4" Then
        strPriceGrade = mstr���ļ۸�ȼ�
    Else
        strPriceGrade = mstr��ͨ�۸�ȼ�
    End If
    
    If InStr(",4,5,6,7,", mobjBill.Pages(intPage).Details(lngRow).�շ����) > 0 Then
        Call AdjustCpt(mobjBill.Pages(intPage).Details(lngRow).�շ�ϸĿID)
    End If
    
    If strPriceGrade <> "" Then
        strWherePriceGrade = _
            "       And (b.�۸�ȼ� = [2]" & vbNewLine & _
            "            Or (b.�۸�ȼ� Is Null" & vbNewLine & _
            "                And Not Exists(Select 1" & vbNewLine & _
            "                               From �շѼ�Ŀ" & vbNewLine & _
            "                               Where b.�շ�ϸĿId = �շ�ϸĿid And �۸�ȼ� = [2]" & vbNewLine & _
            "                                     And Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')))))"
    Else
        strWherePriceGrade = " And b.�۸�ȼ� Is Null"
    End If
    gstrSQL = _
        " Select B.������ĿID,C.����,C.�վݷ�Ŀ,B.�ּ�,B.ԭ��,B.�Ӱ�Ӽ���,B.�����շ���,B.ȱʡ�۸� " & _
        " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,������Ŀ C " & _
        " Where B.�շ�ϸĿID=A.ID And C.ID=B.������ĿID " & _
        " And Sysdate Between B.ִ������ And Nvl(B.��ֹ����,To_Date('3000-1-1', 'YYYY-MM-DD')) " & _
        "       And A.ID=[1]" & vbNewLine & _
        strWherePriceGrade
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mobjBill.Pages(intPage).Details(lngRow).�շ�ϸĿID, strPriceGrade)
    If Not rsTmp.EOF Then
        With mobjBill.Pages(intPage).Details(lngRow)
            If InStr(",5,6,7,", .�շ����) > 0 Or (.�շ���� = "4" And .Detail.��������) Then
                '����ҩƷʱ��(�����򲻷���),��Ȼ�м�¼(�������Ŀʱ���ж�)
                dblAllTime = .���� * .����
                If gblnҩ����λ And InStr(",5,6,7,", .�շ����) > 0 Then
                    dblAllTime = dblAllTime * .Detail.ҩ����װ '���ʱ�۰��ۼ��������м���
                End If
                
                If dblAllTime <> 0 Or Not .Detail.��� Then
                    If .Detail.���� <= 0 Then
                        gstrSQL = "Select Zl_Fun_Getprice([1],[2],[3]) As Price From Dual"
                    Else
                        gstrSQL = "Select Zl_Fun_Getprice([1],[2],[3],[4],[5]) As Price From Dual"
                    End If
                    Set rsPrice = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, .�շ�ϸĿID, .ִ�в���ID, dblAllTime, 0, .Detail.����)
                    If rsPrice.EOF Then
                        '��ȡ�۸�ʧ��
                        If InStr(",5,6,7,", .�շ����) > 0 Then
                            MsgBox "�� " & lngRow & " ��ҩƷ""" & .Detail.���� & """��ȡ�۸�ʧ�ܣ�", vbInformation, gstrSysName
                        Else
                            MsgBox "�� " & lngRow & " ����������""" & .Detail.���� & """��ȡ�۸�ʧ�ܣ�", vbInformation, gstrSysName
                        End If
                    Else
                        strPrice = Nvl(rsPrice!Price) & "|||"
                        varPrice = Split(strPrice, "|")
                        dblMoney = Val(varPrice(0))
                        dblʣ������ = Val(varPrice(2))
                        
                        If dblʣ������ <> 0 And .Detail.��� Then
                            '����δ�ֽ����
                            If InStr(",5,6,7,", .�շ����) > 0 Then
                                MsgBox "�� " & lngRow & " ��ʱ��ҩƷ""" & .Detail.���� & """��治��,�޷�����۸�", vbInformation, gstrSysName
                            Else
                                MsgBox "�� " & lngRow & " ��ʱ����������""" & .Detail.���� & """��治��,�޷�����۸�", vbInformation, gstrSysName
                            End If
                            dblMoney = 0
                        End If
                    End If
                Else
                    dblMoney = 0
                End If
            Else
                If .Detail.��� Then
                    If .InComes.Count = 0 Then '��һ�μ�����ȡȱʡֵ
                        dblMoney = Val(Nvl(rsTmp!ȱʡ�۸�))
                    Else                        '��ȡ����Ա��ǰ����ı�۽��
                        dblMoney = .InComes(1).��׼����
                        '����û�����ı�۲������۷�Χ����ȡȱʡֵ
                        If CheckScope(Val(Nvl(rsTmp!ԭ��)), Val(Nvl(rsTmp!�ּ�)), dblMoney) <> "" Then
                            dblMoney = Val(Nvl(rsTmp!ȱʡ�۸�))
                        End If
                    End If
                End If
            End If
        End With
        
        '�����ԭ�м�¼
        Set mobjBill.Pages(intPage).Details(lngRow).InComes = New BillInComes
        
        '��д���з��ü�¼
        For i = 1 To rsTmp.RecordCount
            Set mobjBillIncome = New BillInCome
            With mobjBillIncome
                .������ĿID = rsTmp!������ĿID
                .������Ŀ = rsTmp!����
                .�վݷ�Ŀ = Nvl(rsTmp!�վݷ�Ŀ)
                .ԭ�� = Val(Nvl(rsTmp!ԭ��))
                .�ּ� = Val(Nvl(rsTmp!�ּ�))
                
                If InStr(",5,6,7,", mobjBill.Pages(intPage).Details(lngRow).�շ����) > 0 Then
                    If gblnҩ����λ Then
                        .��׼���� = Format(dblMoney * mobjBill.Pages(intPage).Details(lngRow).Detail.ҩ����װ, gstrFeePrecisionFmt)
                    Else
                        .��׼���� = Format(dblMoney, gstrFeePrecisionFmt)
                    End If
                Else
                    If mobjBill.Pages(intPage).Details(lngRow).Detail.��� Then
                        .��׼���� = Format(dblMoney, gstrFeePrecisionFmt)
                    Else
                        .��׼���� = Format(Nvl(rsTmp!�ּ�, 0), gstrFeePrecisionFmt)
                    End If
                End If
                
                'Ӧ�ս��=���� * ���� * ����
                .Ӧ�ս�� = .��׼���� * mobjBill.Pages(intPage).Details(lngRow).���� * mobjBill.Pages(intPage).Details(lngRow).����
                
                '�������������ü���(����������Ŀ)
                If mobjBill.Pages(intPage).Details(lngRow).���ӱ�־ = 1 And mobjBill.Pages(intPage).Details(lngRow).�շ���� = "F" Then
                    .Ӧ�ս�� = .Ӧ�ս�� * IIf(IsNull(rsTmp!�����շ���), 1, rsTmp!�����շ��� / 100)
                End If
                
                '�Ӱ�����ʼ���
                dbl�Ӱ�Ӽ��� = 0
                If mobjBill.�Ӱ��־ = 1 And mobjBill.Pages(intPage).Details(lngRow).Detail.�Ӱ�Ӽ� Then
                    dbl�Ӱ�Ӽ��� = IIf(IsNull(rsTmp!�Ӱ�Ӽ���), 0, rsTmp!�Ӱ�Ӽ��� / 100)             '������ݷѱ����ʵ�ս���
                    .Ӧ�ս�� = .Ӧ�ս�� + .Ӧ�ս�� * dbl�Ӱ�Ӽ���
                End If
                
                .Ӧ�ս�� = CCur(Format(.Ӧ�ս��, gstrDec))
                
                dblAllTime = mobjBill.Pages(intPage).Details(lngRow).���� * mobjBill.Pages(intPage).Details(lngRow).����
                If InStr(",5,6,7,", mobjBill.Pages(intPage).Details(lngRow).�շ����) > 0 Then
                    If gblnҩ����λ Then dblAllTime = dblAllTime * mobjBill.Pages(intPage).Details(lngRow).Detail.ҩ����װ
                End If
                
                If mobjBill.Pages(intPage).Details(lngRow).Detail.���ηѱ� Or bln��������ۿ� Then
                    .ʵ�ս�� = .Ӧ�ս��
                    mobjBill.Pages(intPage).Details(lngRow).�ѱ� = mobjBill.�ѱ�
                Else
                    If .Ӧ�ս�� = 0 Then
                        .ʵ�ս�� = 0
                        mobjBill.Pages(intPage).Details(lngRow).�ѱ� = mobjBill.�ѱ�
                    Else
                        'ҩƷ���ɱ��ۼ���,��������
                        str�ѱ� = IIf(glngSys Like "8??", mobjBill.�ѱ�, zlStr.TrimEx(mobjBill.�ѱ� & "," & lbl��̬�ѱ�.Tag, ","))
                        
                        .ʵ�ս�� = CCur(Format(ActualMoney(str�ѱ�, .������ĿID, .Ӧ�ս��, _
                            mobjBill.Pages(intPage).Details(lngRow).�շ�ϸĿID, mobjBill.Pages(intPage).Details(lngRow).ִ�в���ID, dblAllTime, dbl�Ӱ�Ӽ���), gstrDec))
                        mobjBill.Pages(intPage).Details(lngRow).�ѱ� = str�ѱ�
                    End If
                End If
                
                '��ȡ��Ŀ������Ϣ,����ֻ��ҽ�����˲���
                If mstrYBPati <> "" Then
                    strInfo = gclsInsure.GetItemInsure(mobjBill.����ID, mobjBill.Pages(intPage).Details(lngRow).�շ�ϸĿID, .ʵ�ս��, True, mintInsure, _
                        mobjBill.Pages(intPage).Details(lngRow).ժҪ & "||" & dblAllTime)
                    If strInfo <> "" Then
                        mobjBill.Pages(intPage).Details(lngRow).������Ŀ�� = Val(Split(strInfo, ";")(0)) <> 0
                        mobjBill.Pages(intPage).Details(lngRow).���մ���ID = Val(Split(strInfo, ";")(1))
                        .ͳ���� = Format(Val(Split(strInfo, ";")(2)), gstrDec)
                        mobjBill.Pages(intPage).Details(lngRow).���ձ��� = CStr(Split(strInfo, ";")(3))
                        
                        If UBound(Split(strInfo, ";")) >= 4 Then
                            If CStr(Split(strInfo, ";")(4)) <> "" Then mobjBill.Pages(intPage).Details(lngRow).ժҪ = CStr(Split(strInfo, ";")(4))
                            If UBound(Split(strInfo, ";")) >= 5 Then
                                If Split(strInfo, ";")(5) <> "" Then mobjBill.Pages(intPage).Details(lngRow).Detail.���� = Split(strInfo, ";")(5)
                            End If
                        End If
                    End If
                End If
                'ʵ�ս�����Key��,�Դ���ֱ�����(��Key�д��ԭʼʵ�ս��,����)
                mobjBill.Pages(intPage).Details(lngRow).InComes.Add .������ĿID, .������Ŀ, .�վݷ�Ŀ, .��׼����, .Ӧ�ս��, .ʵ�ս��, .ԭ��, .�ּ�, "_" & .ʵ�ս��, .ͳ����
            End With
            rsTmp.MoveNext
        Next
    Else
        '���û��������Ŀ,�������Ӧ�ĳ������
        Set mobjBill.Pages(intPage).Details(lngRow).InComes = New BillInComes
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ShowDetails(Optional lngRow As Long, Optional intCurSubItem As Integer = 0, Optional intSubItemCount As Integer = 0)
'���ܣ�ˢ����ʾ��ǰ����ָ���л������е�����
'������lngRow=ָ����,Ϊ0��ʾ��ʾ������
'˵����ExpenseBill���ϵ�������Ӧ���ݵ��к�
    Dim curTotal As Currency, i As Long, str���㵥λ As String
    Dim intCount As Integer
    
    Bill.Redraw = False
    If lngRow = 0 Then
        For i = 1 To mobjBill.Pages(mintPage).Details.Count
            Call ShowDetail(i)
        Next
    ElseIf mobjBill.Pages(mintPage).Details.Count > 0 Then
        Call ShowDetail(lngRow, intCurSubItem, intSubItemCount)
    End If
    Bill.Redraw = True
    
    '��ʾ����С��
    lblSubӦ��.Caption = "Ӧ��:" & Format(GetBillSum(True, CLng(mintPage)), gstrDec)
    lblSubʵ��.Caption = "ʵ��:" & Format(GetBillSum(False, CLng(mintPage)), gstrDec)
    
    i = Get��ҩ����(str���㵥λ)
    If i = 0 Then
        lblAmount.Caption = ""
    Else
        lblAmount.Caption = "��ҩ��:" & i & str���㵥λ
    End If
    
    
    If mbytInFun = 2 Or mbytInState = 2 Then
        curTotal = GetBillSum
        lblTotal.Caption = "�ϼ�:" & Format(curTotal, gstrDec)
        If mbytInFun = 2 And IsNumeric(cmdOK.Tag) Then
            '����ʱ��ʾ���㵱ǰ���ݷ���,�����۱���Ҫ��
            sta.Panels(Pan.C4Ԥ����Ϣ).Text = "Ԥ��:" & Format(Val(cmdOK.Tag), "0.00")
            sta.Panels(Pan.C4Ԥ����Ϣ).Text = sta.Panels(Pan.C4Ԥ����Ϣ).Text & "/����:" & Format(Val(cmdCancel.Tag) + IIf(mbytBilling = 0, curTotal, 0), "0.00")
            sta.Panels(Pan.C4Ԥ����Ϣ).Text = sta.Panels(Pan.C4Ԥ����Ϣ).Text & "/ʣ��:" & Format(Val(cmdPrint.Tag) - IIf(mbytBilling = 0, curTotal, 0), "0.00")
        End If
    End If
End Sub

Private Sub ShowDetail(lngRow As Long, Optional intCurSubItem As Integer = 0, Optional intSubItemCount As Integer = 0)
'���ܣ�ˢ����ʾָ���е�����
'������lngRow=ָ����
'         intCurSubItem-���صĵ�ǰ�ײ�
'         intSubItemCount- ��Ҫ������ײ���˵��,�ܹ��ײ���Ŀ��(�Ƿ�Ϊ���һ��)
'˵����ExpenseBill���ϵ�������Ӧ���ݵ��к�
    Dim i As Long, j As Long, strTemp As String
    Dim cur��� As Currency, dbl���� As Double
    
    If lngRow > Bill.Rows - 1 Then Exit Sub
    If lngRow > mobjBill.Pages(mintPage).Details.Count Then Exit Sub
    
    '���������
    For i = 1 To Bill.COLS - 1
        '����ʱ�շ�������
        If Not (i = 1 And Bill.TextMatrix(lngRow, i) <> "") Then Bill.TextMatrix(lngRow, i) = ""
    Next
    
    If mobjBill.Pages(mintPage).Details(lngRow).�շ���� <> "" Then
        Bill.RowData(lngRow) = Asc(mobjBill.Pages(mintPage).Details(lngRow).�շ����)
    End If
    
    'ˢ�µ�����
    '����:29201
    strTemp = ""
    If mobjBill.Pages(mintPage).Details(lngRow).�������� <> 0 Then
         strTemp = "��"
         If intSubItemCount > 0 Then
            If intCurSubItem = intSubItemCount Then
                    strTemp = "��"
            End If
         Else
                If lngRow < mobjBill.Pages(mintPage).Details.Count Then
                    If mobjBill.Pages(mintPage).Details(lngRow).�������� <> mobjBill.Pages(mintPage).Details(lngRow + 1).�������� Then
                         strTemp = "��"
                    End If
                ElseIf lngRow = mobjBill.Pages(mintPage).Details.Count Then
                         strTemp = "��"
                End If
          End If
        strTemp = "  " & strTemp & " "
    End If
    
    For i = 1 To Bill.COLS - 1
        Select Case Bill.TextMatrix(0, i)
            Case "���"
                '������ݻ������Ŀֻ(��)��ʾ����
                Bill.TextMatrix(lngRow, i) = mobjBill.Pages(mintPage).Details(lngRow).Detail.�������
            Case "��������"
                Bill.TextMatrix(lngRow, i) = mobjBill.Pages(mintPage).Details(lngRow).��������
            Case "��Ŀ"
                Bill.TextMatrix(lngRow, i) = strTemp & mobjBill.Pages(mintPage).Details(lngRow).Detail.����
            Case "���"
                Bill.TextMatrix(lngRow, i) = mobjBill.Pages(mintPage).Details(lngRow).Detail.���
            Case "��Ʒ��"
                Bill.TextMatrix(lngRow, i) = strTemp & mobjBill.Pages(mintPage).Details(lngRow).Detail.��Ʒ��
            Case "��λ"
                If InStr(",5,6,7,", mobjBill.Pages(mintPage).Details(lngRow).�շ����) > 0 And gblnҩ����λ Then
                    Bill.TextMatrix(lngRow, i) = mobjBill.Pages(mintPage).Details(lngRow).Detail.ҩ����λ
                Else
                    Bill.TextMatrix(lngRow, i) = mobjBill.Pages(mintPage).Details(lngRow).Detail.���㵥λ
                End If
            Case "����"
                Bill.TextMatrix(lngRow, i) = IIf(mobjBill.Pages(mintPage).Details(lngRow).���� = 0, 1, mobjBill.Pages(mintPage).Details(lngRow).����)
            Case "����"
                '�����ڵ�һ����ʾʱ��Ĭ������Ϊ1
                Bill.TextMatrix(lngRow, i) = FormatEx(mobjBill.Pages(mintPage).Details(lngRow).����, 5)
            Case "����"
                '�����Ǹ��շ�ϸĿ����������Ŀ�ĺϼ�
                '��һ�μ���ʱ����Ĭ������Ϊ1�Ļ����ϼ��������
                dbl���� = 0
                If mobjBill.Pages(mintPage).Details(lngRow).InComes.Count > 0 Then
                    For j = 1 To mobjBill.Pages(mintPage).Details(lngRow).InComes.Count
                        dbl���� = dbl���� + mobjBill.Pages(mintPage).Details(lngRow).InComes(j).��׼����
                    Next
                End If
                Bill.TextMatrix(lngRow, i) = Format(dbl����, gstrFeePrecisionFmt)
            Case "Ӧ�ս��"
                'Ӧ�ս���Ǹ��շ�ϸĿ����������Ŀ�ĺϼ�
                cur��� = 0
                If mobjBill.Pages(mintPage).Details(lngRow).InComes.Count > 0 Then
                    For j = 1 To mobjBill.Pages(mintPage).Details(lngRow).InComes.Count
                        cur��� = cur��� + mobjBill.Pages(mintPage).Details(lngRow).InComes(j).Ӧ�ս��
                    Next
                End If
                Bill.TextMatrix(lngRow, i) = Format(cur���, gstrDec)
            Case "ʵ�ս��"
                'ʵ�ս���Ǹ��շ�ϸĿ����������Ŀ�ĺϼ�
                cur��� = 0
                If mobjBill.Pages(mintPage).Details(lngRow).InComes.Count > 0 Then
                    For j = 1 To mobjBill.Pages(mintPage).Details(lngRow).InComes.Count
                        cur��� = cur��� + mobjBill.Pages(mintPage).Details(lngRow).InComes(j).ʵ�ս��
                    Next
                End If
                Bill.TextMatrix(lngRow, i) = Format(cur���, gstrDec)
            Case "ִ�п���", "��ҩҩ��"
                '������ִ�п���'200402
                If mobjBill.Pages(mintPage).Details(lngRow).ִ�в���ID <> 0 Then
                    If mbytInState = 0 Then
                        mrsUnit.Filter = "ID=" & mobjBill.Pages(mintPage).Details(lngRow).ִ�в���ID
                        If mrsUnit.RecordCount <> 0 Then
                            Bill.TextMatrix(lngRow, i) = IIf(zlIsShowDeptCode, mrsUnit!���� & "-", "") & mrsUnit!����
                        Else
                            Bill.TextMatrix(lngRow, i) = GET��������(mobjBill.Pages(mintPage).Details(lngRow).ִ�в���ID, mrsUnit)
                        End If
                    Else
                        '�������ֻ(��)��ʾ����
                        Bill.TextMatrix(lngRow, i) = GET��������(mobjBill.Pages(mintPage).Details(lngRow).ִ�в���ID, mrsUnit)
                    End If
                Else
                    Bill.TextMatrix(lngRow, i) = ""
                End If
            Case "��־"
                If mobjBill.Pages(mintPage).Details(lngRow).�շ���� = "F" And mobjBill.Pages(mintPage).Details(lngRow).���ӱ�־ = 1 Then
                    Bill.TextMatrix(lngRow, i) = "��"
                End If
            Case "����"
                Bill.TextMatrix(lngRow, i) = mobjBill.Pages(mintPage).Details(lngRow).Detail.����
        End Select
    Next
    Bill.Text = Bill.MsfObj.Text
End Sub

Public Sub ShowMoney(Optional ByVal intPage As Integer, Optional bln���� As Boolean = True)
'���ܣ�ˢ����ʾ������Ŀ����������֧��Ԥ����ʱ�ı��ս����������ݺϼƵ�
'������bln����=�Ƿ�������ʻ���ʾ
'      intPage=�Ƿ�ֻ���¼���ָ������(�ӿ��ٶ�)��0-ȫ������,-1,ȫ������,x-����ָ������
    Dim rsTmp As New ADODB.Recordset, arrDetail As Variant
    Dim cur���ϼ� As Currency, curʵ�ս�� As Currency, cur���ø��� As Currency
    Dim cur���� As Currency, curTotal As Currency
    Dim curȫ�Ը� As Currency, cur���Ը� As Currency, cur����ͳ�� As Currency
    Dim curʵ�պϼ� As Currency, curӦ�պϼ� As Currency, strTmp As String
    Dim i As Integer, j As Integer, k As Integer, p As Integer
    Dim blnExist As Boolean, blnDo As Boolean, strSQL As String

    '�������ܷ�Ŀ,��ͳ�Ʊ�����ؽ��
    '-------------------------------------------------------------------------
    
''    '����ʹ��Ԥ����ɷ�
''    If mbytInFun = 0 And gblnPrePayPriority And Val(sta.Panels(Pan.C4Ԥ����Ϣ).Tag) > 0 And txtԤ�����.Enabled Then
''        If Not Me.ActiveControl Is txtԤ����� Then
''            curTotal = GetBillSum - GetMedicareSum
''            If curTotal > 0 Then
''                txtԤ�����.Text = Format(IIf(curTotal > Val(sta.Panels(Pan.C4Ԥ����Ϣ).Tag), Val(sta.Panels(Pan.C4Ԥ����Ϣ).Tag), curTotal), "0.00")
''            End If
''        End If
''    End If
    
    cur���ϼ� = Format(Val(txtԤ�����.Text), "0.00")
        
    Set mcolMoneys = New BillInComes
    
    For p = 1 To mobjBill.Pages.Count
        arrDetail = Array()
        curӦ�պϼ� = 0: curʵ�պϼ� = 0
        cur����ͳ�� = 0: curȫ�Ը� = 0: cur���Ը� = 0
        If intPage = 0 Or p = intPage Then
            If mobjBill.Pages(p).NO = "" Then
                '���ŵ���������ֱ�������
                For i = 1 To mobjBill.Pages(p).Details.Count
                    For j = 1 To mobjBill.Pages(p).Details(i).InComes.Count
                        With mobjBill.Pages(p).Details(i).InComes(j)
                            '�ϲ������е��ݵ���Ŀ����
                            blnExist = False
                            For k = 1 To mcolMoneys.Count
                                strTmp = IIf(gint����ϼ� = 0, .�վݷ�Ŀ, IIf(gint����ϼ� = 2, "��" & p & "��", .������Ŀ)) '31479
                                If mcolMoneys(k).�վݷ�Ŀ = strTmp Then
                                    blnExist = True: Exit For
                                End If
                            Next
                            If blnExist Then
                                mcolMoneys(k).Ӧ�ս�� = mcolMoneys(k).Ӧ�ս�� + .Ӧ�ս��
                                mcolMoneys(k).ʵ�ս�� = mcolMoneys(k).ʵ�ս�� + .ʵ�ս��
                            Else
                                strTmp = IIf(gint����ϼ� = 0, .�վݷ�Ŀ, IIf(gint����ϼ� = 2, "��" & p & "��", .������Ŀ)) '31479
                                mcolMoneys.Add 0, strTmp, strTmp, 0, .Ӧ�ս��, .ʵ�ս��
                            End If
                            
                            '�ϲ�����ǰ���ݵ���Ŀ����
                            blnExist = False
                            For k = 0 To UBound(arrDetail)
                                strTmp = IIf(gint����ϼ� = 0, .�վݷ�Ŀ, IIf(gint����ϼ� = 2, "��" & p & "��", .������Ŀ)) '31479
                                If CStr(Split(arrDetail(k), ",")(0)) = strTmp Then
                                    blnExist = True: Exit For
                                End If
                            Next
                            If blnExist Then
                                arrDetail(k) = Split(arrDetail(k), ",")(0) & "," & _
                                    Val(Split(arrDetail(k), ",")(1)) + .Ӧ�ս�� & "," & _
                                    Val(Split(arrDetail(k), ",")(2)) + .ʵ�ս��
                            Else
                                strTmp = IIf(gint����ϼ� = 0, .�վݷ�Ŀ, IIf(gint����ϼ� = 2, "��" & p & "��", .������Ŀ)) '31479
                                ReDim Preserve arrDetail(UBound(arrDetail) + 1)
                                arrDetail(UBound(arrDetail)) = strTmp & "," & .Ӧ�ս�� & "," & .ʵ�ս��
                            End If
                                 
                            '--
                            curӦ�պϼ� = curӦ�պϼ� + .Ӧ�ս��
                            curʵ�պϼ� = curʵ�պϼ� + .ʵ�ս��
                            
                            'ͳ�Ʊ��ս��
                            curʵ�ս�� = .ʵ�ս��
                            If .ͳ���� = 0 Or Not mobjBill.Pages(p).Details(i).������Ŀ�� Then
                                '��ԭʼ���Ϊ׼,���ֱܷҴ���
                                curȫ�Ը� = curȫ�Ը� + curʵ�ս��
                            Else
                                cur����ͳ�� = cur����ͳ�� + .ͳ����
                                '��ԭʼ���Ϊ׼,���ֱܷҴ���
                                cur���Ը� = cur���Ը� + curʵ�ս�� - .ͳ����
                            End If
                        End With
                    Next
                Next
            Else
                '�õ�������ȡ�Ļ��۵�����
                strSQL = "Select A.�վݷ�Ŀ,B.���� as ������Ŀ," & _
                    " A.Ӧ�ս��,A.ʵ�ս��,A.ͳ����,A.������Ŀ��" & _
                    " From ������ü�¼ A,������Ŀ B" & _
                    " Where A.��¼����=1 And A.��¼״̬ IN(0,1,3) And A.������ĿID=B.ID And A.NO=[1]" & _
                    " Order by ���"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mobjBill.Pages(p).NO)
                For i = 1 To rsTmp.RecordCount
                    '�ϲ������е��ݵ���Ŀ����
                    blnExist = False
                    For k = 1 To mcolMoneys.Count
                        strTmp = IIf(gint����ϼ� = 0, rsTmp!�վݷ�Ŀ, IIf(gint����ϼ� = 2, "��" & p & "��", rsTmp!������Ŀ)) '31479
                        If mcolMoneys(k).�վݷ�Ŀ = strTmp Then
                            blnExist = True: Exit For
                        End If
                    Next
                    If blnExist Then
                        mcolMoneys(k).Ӧ�ս�� = mcolMoneys(k).Ӧ�ս�� + Nvl(rsTmp!Ӧ�ս��, 0)
                        mcolMoneys(k).ʵ�ս�� = mcolMoneys(k).ʵ�ս�� + Nvl(rsTmp!ʵ�ս��, 0)
                    Else
                        strTmp = IIf(gint����ϼ� = 0, rsTmp!�վݷ�Ŀ, IIf(gint����ϼ� = 2, "��" & p & "��", rsTmp!������Ŀ))  '31479
                        mcolMoneys.Add 0, strTmp, strTmp, 0, Nvl(rsTmp!Ӧ�ս��, 0), Nvl(rsTmp!ʵ�ս��, 0)
                    End If
                    
                    '�ϲ�����ǰ���ݵ���Ŀ����
                    blnExist = False
                    For k = 0 To UBound(arrDetail)
                        strTmp = IIf(gint����ϼ� = 0, rsTmp!�վݷ�Ŀ, IIf(gint����ϼ� = 2, "��" & p & "��", rsTmp!������Ŀ)) '31479
                        If CStr(Split(arrDetail(k), ",")(0)) = strTmp Then
                            blnExist = True: Exit For
                        End If
                    Next
                    If blnExist Then
                        arrDetail(k) = Split(arrDetail(k), ",")(0) & "," & _
                            Val(Split(arrDetail(k), ",")(1)) + Nvl(rsTmp!Ӧ�ս��, 0) & "," & _
                            Val(Split(arrDetail(k), ",")(2)) + Nvl(rsTmp!ʵ�ս��, 0)
                    Else
                        strTmp = IIf(gint����ϼ� = 0, rsTmp!�վݷ�Ŀ, IIf(gint����ϼ� = 2, "��" & p & "��", rsTmp!������Ŀ)) '31479
                        ReDim Preserve arrDetail(UBound(arrDetail) + 1)
                        arrDetail(UBound(arrDetail)) = strTmp & "," & Nvl(rsTmp!Ӧ�ս��, 0) & "," & Nvl(rsTmp!ʵ�ս��, 0)
                    End If
                                        
                    '--
                    curӦ�պϼ� = curӦ�պϼ� + Nvl(rsTmp!Ӧ�ս��, 0)
                    curʵ�պϼ� = curʵ�պϼ� + Nvl(rsTmp!ʵ�ս��, 0)
                    
                    'ͳ�Ʊ��ս��
                    curʵ�ս�� = Nvl(rsTmp!ʵ�ս��, 0)
                    If Nvl(rsTmp!ͳ����, 0) = 0 Or Nvl(rsTmp!������Ŀ��, 0) = 0 Then
                        '��ԭʼ���Ϊ׼,���ֱܷҴ���
                        curȫ�Ը� = curȫ�Ը� + curʵ�ս��
                    Else
                        cur����ͳ�� = cur����ͳ�� + Nvl(rsTmp!ͳ����, 0)
                        '��ԭʼ���Ϊ׼,���ֱܷҴ���
                        cur���Ը� = cur���Ը� + curʵ�ս�� - Nvl(rsTmp!ͳ����, 0)
                    End If
                    
                    rsTmp.MoveNext
                Next
            End If
        Else
            With mobjBill.Pages(p)
                curӦ�պϼ� = mobjBill.Pages(p).Ӧ�ս��
                curʵ�պϼ� = mobjBill.Pages(p).ʵ�ս��
                cur����ͳ�� = mobjBill.Pages(p).����ͳ��
                curȫ�Ը� = mobjBill.Pages(p).ȫ�Ը�
                cur���Ը� = mobjBill.Pages(p).���Ը�
                
                'ֱ��ȡKeyֵ����Ŀ����,Ӧ�ս��,ʵ�ս��;
                arrDetail = Split(.Key, ";")
                For i = 0 To UBound(arrDetail)
                    '�ϲ������е��ݵ���Ŀ����
                    blnExist = False
                    For k = 1 To mcolMoneys.Count
                        If mcolMoneys(k).�վݷ�Ŀ = CStr(Split(arrDetail(i), ",")(0)) Then
                            blnExist = True: Exit For
                        End If
                    Next
                    If blnExist Then
                        mcolMoneys(k).Ӧ�ս�� = mcolMoneys(k).Ӧ�ս�� + Val(Split(arrDetail(i), ",")(1))
                        mcolMoneys(k).ʵ�ս�� = mcolMoneys(k).ʵ�ս�� + Val(Split(arrDetail(i), ",")(2))
                    Else
                        strTmp = CStr(Split(arrDetail(i), ",")(0))
                        mcolMoneys.Add 0, strTmp, strTmp, 0, Val(Split(arrDetail(i), ",")(1)), Val(Split(arrDetail(i), ",")(2))
                    End If
                Next
            End With
        End If
        
        '���µ�ǰ���ݸ����ʻ�֧�����:��֧��Ԥ����ʱ
        'ҽ��������������Ӧ�����Ŵ���,�ϼ�Ϊ�������˵������ʻ�
        If Not MCPAR.����Ԥ���� Then
            If mstrYBPati <> "" And bln���� And mstr�����ʻ� <> "" And mcur������� > -1 * mcur����͸֧ Then
                If curʵ�պϼ� >= 0 Then
                    cur���� = cur����ͳ�� + IIf(MCPAR.���Ը�, cur���Ը�, 0) + IIf(MCPAR.ȫ�Ը�, curȫ�Ը�, 0)
                    
                    'ͳ�Ƴ���֮ǰ���ݸ���֧����ĸ������
                    cur���ø��� = 0
                    For i = 1 To p - 1
                        cur���ø��� = cur���ø��� + GetMedicareSum(mstr�����ʻ�, i)
                    Next
                    cur���ø��� = mcur������� - cur���ø���
                                        
                    '��������ʻ�֧�����
                    If cur���ø��� - cur���� >= -1 * mcur����͸֧ Then
                        Call SetBalanceVal(p, mstr�����ʻ�, Format(cur����, "0.00")) '������͸֧��Χ���㹻(����͸֧0Ϊ����)
                    Else
                        If mcur����͸֧ = 0 And cur���ø��� > 0 Then
                            Call SetBalanceVal(p, mstr�����ʻ�, Format(cur���ø���, "0.00")) '������͸֧�������
                        Else
                            '��������͸֧��Χ������͸֧ʱ�����
                            If mcur����͸֧ <> 0 Then
                                Call SetBalanceVal(p, mstr�����ʻ�, cur���ø��� + mcur����͸֧) '������͸֧��Χ��֧��
                            Else
                                Call SetBalanceVal(p, mstr�����ʻ�, 0)
                            End If
                        End If
                    End If
                Else
                    Call SetBalanceVal(p, mstr�����ʻ�, 0)
                End If
            End If
        End If
        
        '��ǰ���ݵ���ػ��ܽ�����
        '----------------------------------------
        With mobjBill.Pages(p)
            .Ӧ�ս�� = curӦ�պϼ�
            .ʵ�ս�� = curʵ�պϼ�
            
            If mbytInFun = 0 Then
                .����ͳ�� = cur����ͳ��
                .ȫ�Ը� = curȫ�Ը�
                .���Ը� = cur���Ը�
            
                'ҽ��֧�������н��,����ΪԤ���㷵�ص�,Ҳ�����Ǹù��̼����
                .���ս�� = GetMedicareSum(, p)
                                
                '���㵱ǰ����Ӧ�ֽ���Ľ��,Ϊ�˼���Ӧ��(�൥��ʱ�ȳ�Ԥ��)
                If cur���ϼ� <> 0 Then
                    If cur���ϼ� <= Format(.ʵ�ս�� - .���ս�� - .���ѿ�ˢ����, "0.00") Then
                        .��Ԥ���� = cur���ϼ�
                    Else
                        .��Ԥ���� = Format(.ʵ�ս�� - .���ս�� - .���ѿ�ˢ����, "0.00")
                    End If
                    cur���ϼ� = cur���ϼ� - .��Ԥ����
                Else
                    .��Ԥ���� = cur���ϼ�
                End If
                
                
                
                '���㵱ǰ����Ӧ�ɽ��ֱҴ�������
                '�ָ����㷽ʽ
                If cbo���㷽ʽ.ListIndex = -1 And .�շѽ��� <> "" Then
                    Call zlControl.CboSetIndex(cbo���㷽ʽ.hWnd, cbo.FindIndex(cbo���㷽ʽ, gstr���㷽ʽ, True))
                    If cbo���㷽ʽ.ListIndex = -1 And cbo���㷽ʽ.ListCount <> 0 Then
                        Call zlControl.CboSetIndex(cbo���㷽ʽ.hWnd, 0)
                    End If
                End If
                
                blnDo = False '�ֽ�ʽʱ�Ŵ���ֱ�,ҽ��Ҫ��ʱ�Ŵ���
                If cbo���㷽ʽ.ListIndex <> -1 And cbo���㷽ʽ.Visible Then
                    If cbo���㷽ʽ.ItemData(cbo���㷽ʽ.ListIndex) = 1 Then
                        blnDo = True
                    End If
                End If
                
                If blnDo And mstrYBPati <> "" Then
                    If MCPAR.����Ԥ���� Then
                        If Not MCPAR.�ֱҴ��� Then
                            blnDo = False
                        End If
                    End If
                End If
                
                If blnDo Then
                    .Ӧ�ɽ�� = CentMoney(.ʵ�ս�� - .���ս�� - .��Ԥ���� - .���ѿ�ˢ����)
                Else
                     .Ӧ�ɽ�� = Format(.ʵ�ս�� - .���ս�� - .��Ԥ���� - .���ѿ�ˢ����, "0.00")
                End If
            
                .����� = Format(.Ӧ�ɽ�� - (.ʵ�ս�� - .���ս�� - .��Ԥ���� - .���ѿ�ˢ����), gstrDec)
                
                .�շѽ��� = "" '���ַ�ʽ���ͻ
            End If
            
            'Keyֵ�ı���,���ڿ��ټ���
            strTmp = ""
            For i = 0 To UBound(arrDetail)
                strTmp = strTmp & ";" & Split(arrDetail(i), ",")(0) & "," & _
                    Split(arrDetail(i), ",")(1) & "," & Split(arrDetail(i), ",")(2)
            Next
            .Key = Mid(strTmp, 2)
        End With
    Next
    
    'ˢ����ʾ���е��ݵĸ����ʻ�֧�����
    '-------------------------------------------------------------------------
    If mstrYBPati <> "" And bln���� And mstr�����ʻ� <> "" And mcur������� > 0 Then
        If Not MCPAR.����Ԥ���� Then
            With vsBalance
                For i = 0 To .Rows - 1
                    If .TextMatrix(i, 0) = mstr�����ʻ� Then Exit For
                Next
                If i <= .Rows - 1 Then
                    .TextMatrix(i, 1) = Format(GetMedicareSum(mstr�����ʻ�), "0.00")
                End If
            End With
        End If
    End If
    
    'ˢ����ʾ���е��ݵķ�����(�շ�Ҫ��������������)
    '-------------------------------------------------------------------------
    mshMoney.Redraw = False
    If mcolMoneys.Count > 0 Then
        mshMoney.Rows = mcolMoneys.Count + 1 + mintMoneyRow
    End If
    If mshMoney.Rows < M_MONEY_ROWS Then mshMoney.Rows = M_MONEY_ROWS

    Call SetMoneyList
    
    curӦ�պϼ� = 0: curʵ�պϼ� = 0
    For i = mintMoneyRow + 1 To mcolMoneys.Count + mintMoneyRow
        mshMoney.TextMatrix(i, 0) = mintBillNO + 1
        mshMoney.TextMatrix(i, 1) = mcolMoneys(i - mintMoneyRow).�վݷ�Ŀ
        mshMoney.TextMatrix(i, 2) = Format(mcolMoneys(i - mintMoneyRow).ʵ�ս��, gstrDec)
        curӦ�պϼ� = curӦ�պϼ� + mcolMoneys(i - mintMoneyRow).Ӧ�ս��
        curʵ�պϼ� = curʵ�պϼ� + mcolMoneys(i - mintMoneyRow).ʵ�ս��
        
        '����С��
        If i = mcolMoneys.Count + mintMoneyRow Then
            mshMoney.TextMatrix(i, 3) = Format(curʵ�պϼ�, gstrDec)
        Else
            mshMoney.TextMatrix(i, 3) = ""
        End If
    Next
    On Error Resume Next
    For i = 1 To mshMoney.Rows - 1
        If mshMoney.TextMatrix(i, 0) = mintBillNO + 1 Then
            mshMoney.TopRow = i
        End If
    Next
    On Error GoTo 0
    mshMoney.Redraw = True
        
    '���ºϼƽ����ʾ
    '----------------------------------------------------------
    txtӦ��.Text = Format(mcurBillӦ�� + curӦ�պϼ�, gstrDec)
    txt�ϼ�.Text = Format(mcurBillʵ�� + curʵ�պϼ�, gstrDec)
    txtӦ��.Text = Format(GetMustPaySum + mcurBillӦ��, "0.00")
   
    
    '����ʱ,txt�ۼ�������ʾӦ��,���ֱҴ����Ľ��
    If mbytInFun = 1 Then txt�ۼ�.Text = Format(CentMoney(txt�ϼ�.Text), "0.00")
End Sub

Private Function GetInputDetail(ByVal lng��Ŀid As Long, Optional ByVal lng���� As Long) As Detail
'���ܣ���ȡ�շ���Ŀ��Ϣ
    Dim objDetail As New Detail
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    
    '�����������ϲ���
    If mintInsure = 0 Then
        strSQL = _
            " Select A.ID,A.���,B.���� as �������,A.����,Nvl(E.����,A.����) as ����,E1.���� as ��Ʒ��," & _
            " A.���,A.���㵥λ,A.���ηѱ�,A.�Ƿ���,A.�Ӱ�Ӽ�,A.ִ�п���,A.��������,A.����ժҪ," & _
            " Decode(A.���,'4',D.����ID,C.ҩ��ID) as ҩ��ID," & _
            " Decode(A.���,'4',D.���÷���,C.ҩ������) as ����," & _
            " Decode(A.���,'4',1,C." & gstrҩ����װ & ") as ҩ����װ," & _
            " Decode(A.���,'4',A.���㵥λ,C." & gstrҩ����λ & ") as ҩ����λ,D.��������,A.¼������,C.��ҩ��̬,M1.���� as ��������,M1.���㵥λ as ������λ,C.����ϵ��" & _
            " From �շ���ĿĿ¼ A,�շ���Ŀ��� B,ҩƷ��� C,�������� D,�շ���Ŀ���� E,�շ���Ŀ���� E1,������ĿĿ¼ M1" & _
            " Where A.���=B.���� And A.ID=C.ҩƷID(+) And C.ҩ��ID=M1.ID(+) And A.ID=D.����ID(+)" & _
            " And A.ID=E.�շ�ϸĿID(+) And E.����(+)=1 And E.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
            " And A.ID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3" & _
            " And A.ID=[1]"
    Else
        strSQL = _
            " Select A.ID,A.���,B.���� as �������,A.����,Nvl(E.����,A.����) as ����,E1.���� as ��Ʒ��," & _
            " A.���,A.���㵥λ,A.���ηѱ�,A.�Ƿ���,A.�Ӱ�Ӽ�,A.ִ�п���,A.��������,A.����ժҪ," & _
            " Decode(A.���,'4',D.����ID,C.ҩ��ID) as ҩ��ID," & _
            " Decode(A.���,'4',D.���÷���,C.ҩ������) as ����," & _
            " Decode(A.���,'4',1,C." & gstrҩ����װ & ") as ҩ����װ," & _
            " Decode(A.���,'4',A.���㵥λ,C." & gstrҩ����λ & ") as ҩ����λ,D.��������,A.¼������,C.��ҩ��̬,M1.���� as ��������,M1.���㵥λ as ������λ,C.����ϵ��" & _
            " From �շ���ĿĿ¼ A,�շ���Ŀ��� B,ҩƷ��� C,�������� D,�շ���Ŀ���� E,�շ���Ŀ���� E1,����֧����Ŀ M,������ĿĿ¼ M1" & _
            " Where A.���=B.���� And A.ID=C.ҩƷID(+) And C.ҩ��ID=M1.ID(+)  And A.ID=D.����ID(+)" & _
            " And A.ID=E.�շ�ϸĿID(+) And E.����(+)=1 And E.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
            " And A.ID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3" & _
            " And A.ID=M.�շ�ϸĿID(+) And M.����(+)=[2]" & vbNewLine & _
            " And A.ID=[1]"
    End If
        
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��Ŀid, mintInsure)
    With objDetail
        .ID = rsTmp!ID
        .ҩ��ID = Nvl(rsTmp!ҩ��ID, 0) '�����ж������ظ�
        .��� = rsTmp!���
        .������� = rsTmp!�������
        .���� = rsTmp!����
        .���� = rsTmp!����
        .��Ʒ�� = Nvl(rsTmp!��Ʒ��)
        .��� = Nvl(rsTmp!���)
        .���㵥λ = Nvl(rsTmp!���㵥λ)
        .ҩ����λ = Nvl(rsTmp!ҩ����λ)
        .ҩ����װ = Nvl(rsTmp!ҩ����װ, 1)
        .���� = Nvl(rsTmp!����, 0) = 1 '�Ƿ�ҩ������
        .��� = Nvl(rsTmp!�Ƿ���, 0) = 1 '��ҩƷ�����Ƿ�ʱ��
        .���� = Nvl(rsTmp!��������)
        .�Ӱ�Ӽ� = Nvl(rsTmp!�Ӱ�Ӽ�, 0) = 1
        .���ηѱ� = Nvl(rsTmp!���ηѱ�, 0) = 1
        .ִ�п��� = Nvl(rsTmp!ִ�п���, 0)
        .����ժҪ = Nvl(rsTmp!����ժҪ, 0) = 1
        .�������� = Nvl(rsTmp!��������, 0) = 1
        .¼������ = Val("" & rsTmp!¼������)
        .��ҩ��̬ = Val(Nvl(rsTmp!��ҩ��̬))
        .�������� = Nvl(rsTmp!��������)
        .������λ = Nvl(rsTmp!������λ)
        .����ϵ�� = Val(Nvl(rsTmp!����ϵ��))
        .���� = lng����
    End With
    Set GetInputDetail = objDetail
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SetDetail(Detail As Detail, lngRow As Long, lngDoUnit As Long, Optional bytParent As Byte = 0)
'���ܣ�����ָ�����շ�ϸĿ�����趨����ָ�㶨�е��շ�ϸĿ(�����Ļ��޸�)
'˵����
'      1.���������������շ�ϸĿ�У�����
'      2.��bytParent<>0ʱ,��Ϊ���ô�����Ŀ,������Ŀһ����������,������Ŀһ������

    Dim tmpIncomes As New BillInComes
    Dim intPay As Integer, i As Long, dblTime As Double
    
    'ȡ������ҩ�ĸ���
    intPay = GetOtherCTMGroups(lngRow)
    If Detail.��� <> "7" Then intPay = 1
    
    If mobjBill.Pages(mintPage).Details.Count < lngRow Then
        '������ж�Ӧ�ĳ��������δ��ʼ,�����
        With Detail
            '���=�к�,����=0
            '����=1,������Ŀ�Ĵ������������ȷ��
            'ִ�в���ID:����ϸĿִ�п��ұ�־ȡ
            '���ӱ�־:�Ե�һ��Ϊ��,����Ϊ������Ȩ
            '���뼯=��
            If bytParent <> 0 Then
                '���ø���RowData
                Bill.RowData(lngRow) = Asc(Detail.���)
                '��ʼ����
                If Detail.���д��� = 0 Then '�ǹ��д���
                    dblTime = Detail.��������
                ElseIf Detail.���д��� = 1 Then '�̶��Ĺ��д���
                    dblTime = IIf(Detail.�������� = 0, 1, Detail.��������)
                ElseIf Detail.���д��� = 2 Then '�������Ĺ��д���
                    dblTime = Detail.�������� * mobjBill.Pages(mintPage).Details(bytParent).����
                End If
            Else
                
                If InStr(",5,6,7,", Detail.���) > 0 Then
                    dblTime = 0
                Else
                    dblTime = 1
                End If
            End If
            mobjBill.Pages(mintPage).Details.Add mobjBill.�ѱ�, Detail, .ID, CInt(lngRow), CInt(bytParent), .���, .���㵥λ, "", intPay, dblTime, 0, lngDoUnit, tmpIncomes
        End With
    Else '��������Ѿ�����,���޸�
        
        If InStr(",5,6,7,", Detail.���) > 0 Then
            dblTime = 0
        Else
            dblTime = 1
        End If
        
        With mobjBill.Pages(mintPage).Details(lngRow)
            Set .Detail = Detail
            Set .InComes = tmpIncomes
            .�ѱ� = mobjBill.�ѱ�
            .���� = intPay
            .���ӱ�־ = 0
            .���㵥λ = Detail.���㵥λ
            .�շ���� = Detail.���
            .�շ�ϸĿID = Detail.ID
            .���� = dblTime
            .��� = lngRow
            .�������� = 0
            .ִ�в���ID = lngDoUnit
        End With
    End If
End Sub

Private Function CheckHaveChildren(lngRow As Long) As Boolean
'���ܣ��жϸ����Ƿ�Ӧ��ȡ������Ŀ
'˵�����������շ���Ŀ�д�����Ŀ����δȡ��ȡ��
    Dim rsTmp As ADODB.Recordset
    Dim i As Long, blnExist As Boolean
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    strSQL = "Select Count(����ID) as NUM From �շѴ�����Ŀ Where ����ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mobjBill.Pages(mintPage).Details(lngRow).�շ�ϸĿID)
    If rsTmp.RecordCount <> 0 Then
        If IsNull(rsTmp!Num) Then
            CheckHaveChildren = False
        ElseIf rsTmp!Num = 0 Then
            CheckHaveChildren = False
        Else
            blnExist = False
            For i = lngRow + 1 To mobjBill.Pages(mintPage).Details.Count
                If mobjBill.Pages(mintPage).Details(i).�������� = lngRow Then
                    blnExist = True: Exit For
                End If
            Next
            If Not blnExist Then
                CheckHaveChildren = True
            Else
                CheckHaveChildren = False
            End If
        End If
    Else
        CheckHaveChildren = False
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function CheckMainItem(ByVal lngRow As Long, Optional ByVal intPage As Long) As Boolean
'���ܣ��жϵ�ǰ�е���Ŀ�Ƿ���д�����Ŀ
    Dim i As Long
    
    If intPage = 0 Then intPage = mintPage
    
    If mobjBill.Pages(intPage).Details.Count >= lngRow Then
        For i = lngRow + 1 To mobjBill.Pages(intPage).Details.Count
            If mobjBill.Pages(intPage).Details(i).�������� = lngRow Then
                CheckMainItem = True: Exit Function
            End If
        Next
    End If
End Function

Private Function GetSubDetails(ByVal lng��Ŀid As Long) As Details
'���ܣ�����һ���շ�ϸĿ�Ĵ�����Ŀ��
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim objDetail As New Detail
        
    Set GetSubDetails = New Details
    
    '�������Ĳ���
    strSQL = _
    "Select A.ID,Decode(A.���,'4',E.����ID,D.ҩ��ID) as ҩ��ID,A.���,B.���� as �������," & _
    "       A.����,Nvl(F.����,A.����) as ����,E1.���� as ��Ʒ��,A.���㵥λ,A.���,A.���ηѱ�," & _
    "       Decode(A.���,'4',E.���÷���,D.ҩ������) as ����,A.��������," & _
    "       Decode(A.���,'4',1,D." & gstrҩ����װ & ") as ҩ����װ," & _
    "       Decode(A.���,'4',A.���㵥λ,D." & gstrҩ����λ & ") as ҩ����λ," & _
    "       A.�Ƿ���,A.�Ӱ�Ӽ�,A.ִ�п���,C.���д���,C.��������,E.��������,D.��ҩ��̬,M1.���� as ��������,M1.���㵥λ as ������λ,D.����ϵ��" & _
    " From �շ���ĿĿ¼ A,�շ���Ŀ��� B,�շѴ�����Ŀ C,ҩƷ��� D,�������� E,�շ���Ŀ���� F,�շ���Ŀ���� E1,������ĿĿ¼ M1" & _
    " Where A.���=B.���� And C.����ID=A.ID And A.ID=D.ҩƷID(+) and D.ҩ��ID=M1.ID(+) And A.ID=E.����ID(+)" & _
    "   And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
    "   And A.ID=F.�շ�ϸĿID(+) And F.����(+)=1 And F.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
    "   And A.ID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3" & _
    "   And C.����ID=[1] Order by ����"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��Ŀid)
    For i = 1 To rsTmp.RecordCount
        Set objDetail = New Detail
        With objDetail
            .ID = rsTmp!ID
            .ҩ��ID = Nvl(rsTmp!ҩ��ID, 0)
            .���� = rsTmp!����
            .��� = rsTmp!�Ƿ��� = 1
            .��� = Nvl(rsTmp!���)
            .ҩ����װ = Nvl(rsTmp!ҩ����װ, 1)
            .ҩ����λ = Nvl(rsTmp!ҩ����λ)
            .���㵥λ = Nvl(rsTmp!���㵥λ)
            .���� = Nvl(rsTmp!����, 0) = 1
            .�Ӱ�Ӽ� = Nvl(rsTmp!�Ӱ�Ӽ�, 0) = 1
            .��� = rsTmp!���
            .������� = rsTmp!�������
            .���� = rsTmp!����
            .��Ʒ�� = Nvl(rsTmp!��Ʒ��)
            .���ηѱ� = Nvl(rsTmp!���ηѱ�, 0) = 1
            .ִ�п��� = Nvl(rsTmp!ִ�п���, 0) 'ȱʡΪ����ȷ����(�û�ѡ)
            .���д��� = Nvl(rsTmp!���д���, 0) 'ȱʡΪ�ǹ̶�,�û����������������
            .�������� = Nvl(rsTmp!��������, 1)
            .���� = Nvl(rsTmp!��������)
            .�������� = Nvl(rsTmp!��������, 0) = 1
            .��ҩ��̬ = Val(Nvl(rsTmp!��ҩ��̬))
            .�������� = Nvl(rsTmp!��������)
            .������λ = Nvl(rsTmp!������λ)
            .����ϵ�� = Val(Nvl(rsTmp!����ϵ��))
            GetSubDetails.Add .ID, .ҩ��ID, .���, .�������, .����, .����, .����, .���, .���㵥λ, .˵��, .���ηѱ�, _
                .ҩ����װ, .ҩ����λ, .����, .���, .�Ӱ�Ӽ�, .ִ�п���, .����, .����ժҪ, .���д���, .��������, .��������, , , , , , , .��ҩ��̬, .��Ʒ��, .��������, .������λ, .����ϵ��
        End With
        rsTmp.MoveNext
    Next
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub DeleteDetail(ByVal lngRow As Long, Optional ByVal intPage As Integer)
'���ܣ�ɾ��ָ���շ���Ŀ��
'˵������ʱ����������е�ɾ��,��Ҫ�����������д�����ϵ����Ӧ�ĵ���
    Dim i As Long
    
    '���δָ��ҳ,���õ�ǰҳ
    If intPage = 0 Then intPage = mintPage
    
    For i = lngRow + 1 To mobjBill.Pages(intPage).Details.Count
        If mobjBill.Pages(intPage).Details(i).�������� <> 0 And _
            mobjBill.Pages(intPage).Details(i).�������� > lngRow Then
            mobjBill.Pages(intPage).Details(i).�������� = mobjBill.Pages(intPage).Details(i).�������� - 1
        End If
        mobjBill.Pages(intPage).Details(i).��� = mobjBill.Pages(intPage).Details(i).��� - 1 '������кŶ�Ӧ
    Next
    mobjBill.Pages(intPage).Details.Remove lngRow
    
    'ɾ����ǰ��ʾ����ҳ��ָ����
    If tbsBill.SelectedItem.Index = intPage Then
        If lngRow = 1 And mobjBill.Pages(intPage).Details.Count = 0 And Bill.Rows = 2 Then
            For i = 1 To Bill.COLS - 1
                Bill.TextMatrix(lngRow, i) = ""
                Bill.RowData(lngRow) = 0
            Next
            Call SetBillRowForeColor(lngRow, Bill.ForeColor)
        Else
            Bill.RemoveMSFItem lngRow
        End If
    End If
End Sub

Private Sub NewYBBill()
'���ܣ�����ҽ�������շ�ʱ����,�����շ�ģʽ�²���ʹ�ö൥���շ�
    Dim i As Integer
    
    Set mobjBill = New ExpenseBill
    Set mcolBalance = New Collection
    mcolBalance.Add Array()
    '�൥���շ�:�ָ�ȱʡ����ҳ��
    mintPage = 1
    If fraBill.Visible Then
        cmdAddBill.Enabled = InStr(1, mstrPrivs, "��ͨ���˶൥���շ�") > 0
        cmdDelBill.Enabled = False
        tbsBill.TabStop = False
        For i = tbsBill.Tabs.Count To 1 Step -1
            tbsBill.Tabs(i).Tag = ""
            If i <> 1 Then tbsBill.Tabs.Remove i
        Next
    End If
    
    mlngPreRow = 0
    mblnHotKey = False
    mstrCardNO = ""
    If mbln���� And mstr���ת��ʱ�� <> "" Then
        txtDate.Text = Format(CDate(mstr���ת��ʱ��) - 1 / 24 / 60, "yyyy-mm-dd HH:MM:SS")
    Else
        txtDate.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    End If
    
    Call InitBalanceGrid
    txtԤ�����.Text = "0.00"
    sta.Panels(Pan.C4Ԥ����Ϣ).Tag = ""
    Original.��Ԥ���� = 0
    Original.ʵ�պϼ� = 0
    Original.Ӧ�ɽ�� = 0
    ''txt����Ӧ��.Visible = False: lblӦ��.Caption = "Ӧ��"
      
    cboNO.Text = ""
    
    'ˢ��Ʊ�ݺ�,ֻ�����õ�ʱ���ڴ�ӡ����ˢ��
    If mbytInFun = 0 Then Call RefreshFact
        
    With mobjBill
        .����ʱ�� = CDate(txtDate.Text)
        .�ѱ� = IIf(cbo�ѱ�.ListIndex = -1, "", Mid(cbo�ѱ�.Text, InStr(cbo�ѱ�.Text, "-") + 1))
        .�Ӱ��־ = chk�Ӱ�.Value
        If cbo��������.ListIndex = -1 Then
            .Pages(mintPage).��������ID = 0
        Else
            .Pages(mintPage).��������ID = cbo��������.ItemData(cbo��������.ListIndex)
        End If
        .Pages(mintPage).������ = IIf(cbo������.ListIndex = -1, "", zlStr.NeedName(cbo������.Text))
        .�����־ = gint������Դ
        .������ = UserInfo.����
        .����Ա��� = UserInfo.���
        .����Ա���� = UserInfo.����
    End With
End Sub

Private Function NewBill(Optional blnFact As Boolean = True, Optional bln�ѱ� As Boolean = True, _
    Optional ByVal blnClearPatiInfor As Boolean) As Boolean
'���ܣ���ʼ��һ���µĵ���(�������)
'������blnFact=�Ƿ�ȡƱ��
'      bln�ѱ�=�Ƿ����³�ʼ���ѱ�
    Dim i As Long
    Dim Curdate As Date '��������ǰʱ��
    
    If blnClearPatiInfor Then Set mrsInfo = New ADODB.Recordset
    Set mobjBill = New ExpenseBill
    Set mcolBalance = New Collection
    mcolBalance.Add Array()
    '�൥���շ�:�ָ�ȱʡ����ҳ��
    mintPage = 1
    
    Bill.ColData(BillCol.���) = IIf(gbln�շ����, BillColType.ComboBox, BillColType.UnFocus)
    If cmdIDCard.Visible Then cmdIDCard.Enabled = True
    If cmdRegist.Visible Then cmdRegist.Enabled = True
    
    cmdAddBill.Enabled = InStr(1, mstrPrivs, "��ͨ���˶൥���շ�") > 0
    cmdDelBill.Enabled = False
    tbsBill.TabStop = False
    If fraBill.Visible Then
        For i = tbsBill.Tabs.Count To 1 Step -1
            tbsBill.Tabs(i).Tag = ""
            If i <> 1 Then tbsBill.Tabs.Remove i
        Next
    End If
    mdbl�ɿ� = 0: mdbl�Ҳ� = 0
    
    mstrYBBill = "": mstrYBPati = "": mintInsure = 0
    mcur������� = 0: mcur����͸֧ = 0
    mblnYB�������� = False  '��ͬ�Ĳ��˿������಻ͬ��ҽ������֧�ֲ�ͬ,����Ҫ���
    mbytBillSource = 1
    If txtMCInvoice.Visible Then
        txtMCInvoice.Visible = False
        txtMCInvoice.Text = ""
    End If
    
    mblnSaveAsPrice = False
    mblnHotKey = False
    mbln���ϼ� = False
    Original.ʵ�պϼ� = 0: Original.��Ԥ���� = 0: mlngPreRow = 0
    Original.Ӧ�ɽ�� = 0
'''    txt����Ӧ��.Visible = False: lblӦ��.Caption = "Ӧ��"
    
    mstrCardNO = ""
    txtPatient.ForeColor = Me.ForeColor
    mnuFileSavePrice.Checked = False
    chk����.Value = 0: chk����.Visible = False
    
    If mstr���ʽ <> "" Then
        cboҽ�Ƹ���.ListIndex = cbo.FindIndex(cboҽ�Ƹ���, mstr���ʽ, True)
        If cboҽ�Ƹ���.ListIndex = -1 And cboҽ�Ƹ���.ListCount > 0 Then cboҽ�Ƹ���.ListIndex = 0
    ElseIf cboҽ�Ƹ���.ListCount > 0 Then
        cboҽ�Ƹ���.ListIndex = 0
    End If
    cboҽ�Ƹ���.Locked = False Or mbytInFun = 2
    If mbytInFun = 2 And mbytInState = 0 Then cboBaby.ListIndex = 0
    sta.Panels(Pan.C3�����ʻ�).Tag = "": sta.Panels(Pan.C3�����ʻ�).Text = "": sta.Panels(Pan.C3�����ʻ�).Visible = False
            
    Call InitBalanceGrid
    Call SetButton(2) 'ȷ��,ȡ��
    Call ShowPrePayInfo(False) 'Ԥ����Ϣ��ʼ
    Call ShowPayInfo(True) '����ҽ��
    
    If mbytInFun = 0 And blnClearPatiInfor Then
        Call SetPatientEnableModi(True)
        txtRePrint.Enabled = True: txtModi.Enabled = True: txtIn.Enabled = True
        cboNO.Enabled = True: chkCancel.Enabled = True: cmdDelete.Enabled = True
    End If
        
    If gbyt����ҽ�� = 0 And mstrPrePati <> txtPatient.Text Then
        cbo������.ListIndex = -1: cbo��������.ListIndex = -1: lblDuty.Caption = ""
    End If
    
    If mbln���� And mstr���ת��ʱ�� <> "" Then
        Curdate = CDate(mstr���ת��ʱ��) - 1 / 24 / 60
    Else
        Curdate = zlDatabase.Currentdate
    End If
    txtDate.Text = Format(Curdate, "yyyy-MM-dd HH:mm:ss")
    
    If mbytInState = 0 Then
        cboNO.Text = ""
        mstrWarn = ""
        cmdOK.Tag = "": cmdCancel.Tag = "": cmdPrint.Tag = ""
        If blnFact Then
            txtInvoice.Text = ""
            Call ReInitPatiInvoice(blnFact)
        End If
        
'        If mbytInFun = 2 Then Call ClearPatientInfo(True)   '�������ϼ���Ϣ
        chk�Ӱ�.Value = IIf(OverTime(Curdate), 1, 0)
        
        '���㷽ʽ
        If mbytInFun = 0 Then
            i = cbo.FindIndex(cbo���㷽ʽ, gstr���㷽ʽ, True)
            If i = -1 And cbo���㷽ʽ.ListCount > 0 Then i = 0
            Call zlControl.CboSetIndex(cbo���㷽ʽ.hWnd, i)
        End If
        
        '�ѱ����շѻ򻮼�
        If Not (glngSys Like "8??" Or mbytInFun = 2) Then
            cbo�ѱ�.Locked = False: cbo�ѱ�.TabStop = Not cbo�ѱ�.Locked And gbln�ѱ�
            cbo�ѱ�.Visible = True
            lbl��̬�ѱ�.BorderStyle = 0
            lbl��̬�ѱ�.Left = cbo�ѱ�.Left + cbo�ѱ�.Width + 60
            
            If bln�ѱ� Then Call LoadAndSeek�ѱ�(True)
        End If
        
        '����
        With mobjBill
            .����ʱ�� = CDate(txtDate.Text)
            .�ѱ� = IIf(cbo�ѱ�.ListIndex = -1, "", Mid(cbo�ѱ�.Text, InStr(cbo�ѱ�.Text, "-") + 1))
            .�Ӱ��־ = chk�Ӱ�.Value
            If cbo��������.ListIndex = -1 Then
                .Pages(mintPage).��������ID = 0
            Else
                .Pages(mintPage).��������ID = cbo��������.ItemData(cbo��������.ListIndex)
            End If
            .Pages(mintPage).������ = IIf(cbo������.ListIndex = -1, "", zlStr.NeedName(cbo������.Text))
            .�����־ = gint������Դ
            .������ = UserInfo.����
            .����Ա��� = UserInfo.���
            .����Ա���� = UserInfo.����
        End With
        
    End If
    
    NewBill = True
End Function

Private Sub ClearMoney()
'���ܣ����������ʾ��
    Dim i As Integer, j As Integer
    mshMoney.Redraw = False
    mintMoneyRow = 0
    For i = 1 To mshMoney.Rows - 1
        For j = 0 To mshMoney.COLS - 1
            mshMoney.TextMatrix(i, j) = ""
        Next
    Next
    mshMoney.Rows = M_MONEY_ROWS
    mshMoney.Redraw = True
End Sub

Private Function GetDrugWindow(ByVal lngҩ��ID As Long, ByVal str��� As String, ByVal intPage As Integer) As String
'���ܣ���ȡȱʡ�ķ�ҩ����,�������ָ����ȱʡ,����ָ��Ϊ׼,����,����ǻ��۵�,���Ե�һҩƷ�еĴ���Ϊ׼,��������������ͬҩƷ�Ĵ���Ϊ׼
'������intPage=��¼���ĵ��ݱ��
'˵������Ҫ���ڶ൥���շ�ʱ����ͬ����ҩƷ���ܶ�̬���䵽ͬһҩ�����������ǵĴ���ҲӦ��ͬ����ǿ��ָ���ĳ���
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim p As Integer, i As Integer, varData As Variant, varTemp As Variant
    Dim strPayWin As String
    
    Err = 0: On Error GoTo errH:
    strPayWin = ""
    For p = 1 To intPage
         If mobjBill.Pages(p).NO <> "" Then
             If tbsBill.Tabs(p).Tag <> "" Then
                 '����:47489
                 'ȡ���۵��ĵ�һҩƷ�е�ҩ�����бȽ�
                 ''ִ�в���ID|��ҩ����;...
                 varData = Split(tbsBill.Tabs(p).Tag, ";")
                 For i = 0 To UBound(varData)
                     varTemp = Split(varData(i) & "|", "|")
                     If varTemp(0) = lngҩ��ID Then
                          strPayWin = varTemp(1)
                          GoTo GoFind:
                     End If
                 Next
             End If
         Else
             For i = 1 To mobjBill.Pages(p).Details.Count
                 If mobjBill.Pages(p).Details(i).ִ�в���ID = lngҩ��ID _
                     And InStr(",5,6,7,", mobjBill.Pages(p).Details(i).�շ����) > 0 _
                     And mobjBill.Pages(p).Details(i).��ҩ���� <> "" Then
                     strPayWin = mobjBill.Pages(p).Details(i).��ҩ����
                     GoTo GoFind:
                 End If
             Next
         End If
     Next
GoFind:
    If strPayWin <> "" Then GetDrugWindow = strPayWin: Exit Function
    
    strPayWin = GetDefaultWindow(str���, lngҩ��ID)
    '����Ƿ��ϰ�
    strSQL = "Select ���� From ��ҩ���� Where �ϰ��=1 And ҩ��ID=[1] And ����=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҩ��ID, GetDrugWindow)
    If rsTmp.EOF Then strPayWin = ""
    GetDrugWindow = strPayWin
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
 
Private Function ReChargeFee() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ȡ����
    '����:������ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-22 18:18:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPageInfor As Collection, strInvoice As String
    Dim p As Integer, strBalanceIDs As String, lng����ID As Long
    Dim rsErrBlance As ADODB.Recordset
    Dim strSaveNos As String, strSaveSuessNos As String, blnAffair As Boolean
    Dim blnNotCallInsure As Boolean '������ҽ���ӿ�
    Dim blnPatiIndentify As Boolean '�Ƿ��Ѿ�ҽ�������֤
    If mbytInFun <> 0 Then ReChargeFee = True: Exit Function
   Dim strSQL As String, strDate As String
   Dim blnCommit As Boolean
   
    '�������
    If zlIsCheckExistErrBill(mlng�������) = False Then
        MsgBox "��ǰ�쳣�����ѱ������㲻�ܼ�����", vbInformation, gstrSysName
        Exit Function
    End If
    If zlCheckOtherSessionDoing(mlng�������) Then
        MsgBox "��ǰ�������������շѴ����н��д����㲻�ܼ�����", vbInformation, gstrSysName
        Unload Me
        Exit Function
    End If
    
    strInvoice = Trim(txtInvoice.Text)
    If Not CheckBillNOAndBookeFee Then Exit Function

    Set cllPageInfor = New Collection
    On Error GoTo ErrRoll:
    strDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    mobjBill.�Ǽ�ʱ�� = CDate(strDate)
    gcnOracle.BeginTrans: blnCommit = True
    For p = 1 To mobjBill.Pages.Count
        lng����ID = mobjBill.Pages(p).����ID
        strBalanceIDs = strBalanceIDs & "," & lng����ID
        cllPageInfor.Add Array(lng����ID, mobjBill.Pages(p).NO), "K" & p
        strSaveNos = strSaveNos & "," & "'" & mobjBill.Pages(p).NO & "'"
    
        '44507
        strSQL = "Zl_�����շ��쳣_Update('" & mobjBill.Pages(p).NO & "',to_date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'))"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        strSQL = "Zl_Ʊ����ʼ��_Update('" & mobjBill.Pages(p).NO & "','" & strInvoice & "',1)"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Next
    gcnOracle.CommitTrans: blnCommit = False
    On Error GoTo errHandle
    If strBalanceIDs <> "" Then strBalanceIDs = Mid(strBalanceIDs, 2)
    If strSaveNos <> "" Then strSaveNos = Mid(strSaveNos, 2)
    'ҽ������
    If zlInsure��������(strSaveNos) = False Then Exit Function
    If mbytInFun <> 0 Then Exit Function
    
    '�������ʽ
    If frmChargePayMentWin.zlChargeWin(Me, EM_�����շ�, mlngModul, mstrPrivs, mlngShareUseID, mstrUseType, mlng�������, strBalanceIDs, strSaveNos, mobjBill.����ID, mintInsure, mobjBill.����, mobjBill.�Ա�, mobjBill.����, mobjBill.�ѱ�, mdbl�ɿ�, mdbl�Ҳ�) = False Then
        If Not mblnErrBill Then
            Unload Me
        End If
        Exit Function
    End If
    
    '��ʾLed�����Ϣ
    'LED��ʾ:(�ϼ�,)��ҩ����
    If gblnLED And CCur(txt�ϼ�.Text) <> 0 And (mstr���� <> "" Or mstr�д� <> "" Or mstr�ɴ� <> "") Then
        zl9LedVoice.DisplayBank "���úϼ�:" & txt�ϼ�.Text, _
            "ȡҩ����:" & IIf(mstr���� <> "", " " & mstr����, "") & _
            IIf(mstr�ɴ� <> "", " " & mstr�ɴ�, "") & IIf(mstr�д� <> "", " " & mstr�д�, "")
    End If
    
    Call CheckBillNOAndBookeFee(True)
     '��ӡƱ��
     Call PrintBill(strSaveNos, "")
     If Not mblnErrBill Then
        gblnOK = True: Unload Me
     End If
    ReChargeFee = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Exit Function
ErrRoll:
    If blnCommit Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    
End Function

Private Function DelErrBillFee() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�쳣��������
    '����:�쳣�������ϳɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-22 18:18:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPageInfor As Collection, strInvoice As String
    Dim p As Integer, strBalanceIDs As String, lng����ID As Long
    Dim rsErrBlance As ADODB.Recordset
    Dim strSaveNos As String, strSaveSuessNos As String, blnAffair As Boolean
    Dim blnNotCallInsure As Boolean '������ҽ���ӿ�
    Dim blnPatiIndentify As Boolean '�Ƿ��Ѿ�ҽ�������֤
    If mbytInFun <> 0 Then DelErrBillFee = True: Exit Function
   
    strInvoice = Trim(txtInvoice.Text)
    Set cllPageInfor = New Collection
    On Error GoTo errHandle
    '�������
    If zlIsCheckExistErrBill(mlng�������) = False Then
        MsgBox "��ǰ�쳣�����ѱ������㲻�ܼ�����", vbInformation, gstrSysName
        Exit Function
    End If
    If zlCheckOtherSessionDoing(mlng�������) Then
        MsgBox "��ǰ�������������շѴ����н��д����㲻�ܼ�����", vbInformation, gstrSysName
        Unload Me
        Exit Function
    End If
    
    For p = 1 To mobjBill.Pages.Count
        lng����ID = mobjBill.Pages(p).����ID
        strBalanceIDs = strBalanceIDs & "," & lng����ID
        cllPageInfor.Add Array(lng����ID, mobjBill.Pages(p).NO), "K" & p
        strSaveNos = strSaveNos & "," & "'" & mobjBill.Pages(p).NO & "'"
    Next
    If strBalanceIDs <> "" Then strBalanceIDs = Mid(strBalanceIDs, 2)
    If strSaveNos <> "" Then strSaveNos = Mid(strSaveNos, 2)
    mbln�������� = False
    'ҽ������
    '�������ʽ
    If frmChargePayMentWin.zlChargeWin(Me, EM_�쳣����, mlngModul, mstrPrivs, mlngShareUseID, mstrUseType, _
        mlng�������, strBalanceIDs, strSaveNos, mobjBill.����ID, mintInsure, _
        mobjBill.����, mobjBill.�Ա�, mobjBill.����, mobjBill.�ѱ�, mdbl�ɿ�, mdbl�Ҳ�, _
        , , , , , , mbln�˷��쳣) = False Then
        mlng������� = 0
        Unload Me
        Exit Function
    End If
    Dim lng������� As Long
    lng������� = Get���Ͻ������(strSaveNos)
    Call WriteMzInforToCard(mobjBill.����ID, lng�������, True)
    mlng������� = 0:
    gblnOK = True: Unload Me
    DelErrBillFee = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function Get���Ͻ������(ByVal strNos As String) As Long
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���ϵĽ������
    '����:�������ϵĽ������
    '����:���˺�
    '����:2012-12-14 18:52:31
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    strSQL = "Select a.�������" & _
            " From ����Ԥ����¼ A, ������ü�¼ B" & _
            " Where a.����id = b.����id And b.No In (Select Column_Value From Table(f_Str2list([1]))) " & _
            "       And a.��¼״̬ = 2 And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Replace(strNos, "'", ""))
    If rsTemp.EOF Then Exit Function
    Get���Ͻ������ = Val(Nvl(rsTemp!�������))
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function zlInsure��������(ByVal strSaveNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ҽ������
    '����:�����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-20 17:15:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strBillNO As String, blnTrans As Boolean, blnTransMedicare As Boolean
    Dim p As Integer, str���ս��� As String, strAdvance As String, blnMedicareCheck As Boolean
    Dim strTmp As String, i As Long, blnNotCallInsure As Boolean, blnPatiIndentify As Boolean
    Dim strSQL As String, strsuccesNOs As String, strNotSucces As String
    Dim lng����ID  As Long, lng����ID As Long
    Dim dbl�����ʻ� As Double, dblҽ������ As Double
   
    lng����ID = mobjBill.����ID
    strSQL = "" & _
    "   Select A.NO,A.����ID,decode(A.��¼����,1,'Ԥ���',11,'Ԥ���',A.���㷽ʽ) as ���㷽ʽ,A.��Ԥ��, " & _
    "           Decode(B.����,NULL,0,1) as ҽ��, " & _
    "           Decode(C.���㷽ʽ,NULL,0,1) as һ��ͨ, " & _
    "           Decode(nvl(A.�����ID,0),0,0,1) as ҽ�ƿ�, " & _
    "           Decode(nvl(A.���㿨���,0),0,0,1) as ���ѿ�, " & _
    "           nvl(A.У�Ա�־,0) as У�Ա�־  " & _
    "   From ����Ԥ����¼ A, " & _
    "           (Select ���� From ���㷽ʽ where ���� in (3,4)) B," & _
    "           (Select ���㷽ʽ From һ��ͨĿ¼ Where ����=1 ) C " & _
    "   where A.�������=[1]  and a.��¼����=3 and A.���㷽ʽ is not null  " & _
    "               And A.���㷽ʽ=B.����(+)  And A.���㷽ʽ=C.���㷽ʽ(+)"
    
    strSQL = "" & _
    "   Select NO,���㷽ʽ,nvl(sum(��Ԥ��),0) as ������, " & _
    "               nvl(Max(ҽ��),0) as ҽ��, nvl(Max(һ��ͨ),0) as һ��ͨ, " & _
    "               nvl(max(ҽ�ƿ�),0) as ҽ�ƿ�, " & _
    "               nvl(Max(���ѿ�),0) as ���ѿ�,nvl(Max(У�Ա�־),0) as У�Ա�־ " & _
    "   From (" & strSQL & ")" & _
    "   Group by  NO,����ID,���㷽ʽ"
    Err = 0: On Error GoTo errH:
    '�쳣���ݵĽ��㷽ʽ(����Ԥ����)
    Set mrsErrBlance = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng�������)
    
    Err = 0: On Error GoTo errHandle:
    mrsErrBlance.Filter = "ҽ��=0 "
    '������������Ľ��㷽ʽ,��϶�������ҽ��,ֱ�ӷ���
    If mrsErrBlance.RecordCount > 0 Then zlInsure�������� = True: Exit Function
    '����Ƿ����ҽ��
    mrsErrBlance.Filter = 0
    mrsErrBlance.Filter = "ҽ��=1 and У�Ա�־=2  "
    mintInsure = 0
    blnNotCallInsure = False
    If Not mrsErrBlance.EOF Then
        mintInsure = ChargeExistInsure(Nvl(mrsErrBlance!NO), , lng����ID)
        Call initInsurePara(lng����ID)
        '�������������ҽ���ӿ�:
        '1.�൥��һ�ν���(��Ϊ��һ������,���Կ϶�ҽ���ɹ�,����Ҳ����ɹ�)
        '2.�൥�ݵ���һ�νӿ�(��Ϊ��һ������,���Կ϶�ҽ���ɹ�,����Ҳ����ɹ�)
        '3.���е���,������ҽ��,�������ڽ϶Ե����
        '4.����Ԥ����ʱ,ÿ�ŵ��ݶ�Ӧ����ҽ��,ֻ�ǵ��ò��ɹ������
        If MCPAR.�൥��һ�ν��� Then zlInsure�������� = True: Exit Function
        If MCPAR.�൥�ݵ�һ�ν��� Then zlInsure�������� = True:  Exit Function
    End If
    '����ҽ������
    blnPatiIndentify = False: strsuccesNOs = ""
    For p = 1 To mobjBill.Pages.Count
         
        blnNotCallInsure = True
        '�϶��ڵ����д�������
        mrsErrBlance.Filter = "  NO='" & mobjBill.Pages(p).NO & "' and ҽ��=1 and  У�Ա�־=1  "
        If Not mrsErrBlance.EOF Then blnNotCallInsure = False
        If Not MCPAR.����Ԥ���� And blnNotCallInsure Then    '�������������,��ô���ܴ��ڵ���ҽ�������
             mrsErrBlance.Filter = "  NO='" & mobjBill.Pages(p).NO & "' and ҽ��=1"
             If mrsErrBlance.EOF Then blnNotCallInsure = True
        End If
        '��Ҫ��������,��Ҫ��֤���
        If blnPatiIndentify = False And Not blnNotCallInsure Then
            '����ҽ��ˢ����֤
            If MsgBox("ע��:" & vbCrLf & _
                "    ���ݺ�Ϊ" & mobjBill.Pages(p).NO & " ����ҽ�����㵥��, " & _
                "     ĿǰֻԤ������,����δ����ҽ����ʽ����,�Ƿ�����ҽ������?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                If VerifyPati(mobjBill.����ID) = False Then Exit Function
                blnPatiIndentify = True
            End If
        End If
        '��������
        If Not blnNotCallInsure Then
            mrsErrBlance.Filter = "  NO='" & mobjBill.Pages(p).NO & "' and ҽ��=1 and  У�Ա�־=1  "
            With mrsErrBlance
                str���ս��� = ""    '���㷽ʽ|���||
                dbl�����ʻ� = 0: dblҽ������ = 0
                Do While Not .EOF
                    str���ս��� = str���ս��� & "||" & Nvl(!���㷽ʽ) & "|" & Val(Nvl(!������))
                    If Nvl(!���㷽ʽ) = mstr�����ʻ� Then
                        dbl�����ʻ� = dbl�����ʻ� + Val(Nvl(!������))
                    End If
                    If Nvl(!���㷽ʽ) = "ҽ������ " Then
                        dblҽ������ = dblҽ������ + Val(Nvl(!������))
                    End If
                    .MoveNext
                Loop
                If str���ս��� <> "" Then str���ս��� = Mid(str���ս���, 3)
            End With
            gcnOracle.BeginTrans: blnTrans = True: blnTransMedicare = False
            strAdvance = mobjBill.Pages.Count & "|" & p
            If Not gclsInsure.ClinicSwap(mobjBill.Pages(p).����ID, CCur(dbl�����ʻ�), _
                CCur(dblҽ������), mobjBill.Pages(p).ȫ�Ը�, mobjBill.Pages(p).���Ը�, mintInsure, strAdvance) Then
                gcnOracle.RollbackTrans:
                If strsuccesNOs <> "" Then
                    strsuccesNOs = Mid(strsuccesNOs, 2)
                    strSaveNos = "'" & Replace(strsuccesNOs, ",", "','") & "'"
                Else
                    strSaveNos = ""
                End If
                strNotSucces = ""
                For i = p To mobjBill.Pages.Count
                    strNotSucces = strNotSucces & "," & mobjBill.Pages(i).NO
                Next
                If strNotSucces <> "" Then strNotSucces = Mid(strNotSucces, 2)
                If ModifyNotInsureNOs(strNotSucces, strsuccesNOs) = False Then
                    Exit Function
                End If
                zlInsure�������� = True
                Exit Function
            Else
                blnTransMedicare = True
            End If
            blnMedicareCheck = zlInsureCheck(str���ս���, strAdvance)
            '����:
            ' Zl_���������շ�_ҽ������
            gstrSQL = "Zl_���������շ�_ҽ������("
            '  ����id_In   ������ü�¼.����id%Type,
            gstrSQL = gstrSQL & "" & "NULL" & ","
            '  �������_In ����Ԥ����¼.�������%Type,
            gstrSQL = gstrSQL & "" & mlng������� & ","
            '  ���ս���_In Varchar2
            gstrSQL = gstrSQL & IIf(blnMedicareCheck, "'" & strAdvance & "'", "NULL") & ")"
            zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
            strsuccesNOs = strsuccesNOs & "," & mobjBill.Pages(p).NO
            If blnTransMedicare Then Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicSwap, True, mintInsure)
            gcnOracle.CommitTrans: blnTrans = False: blnTransMedicare = False
        End If
    Next
    zlInsure�������� = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    If blnTrans Then
        'ҽ����HIS����ͬһ������,HIS����ʧ��,��ҽ���������ϴ�,������Ҫ��"ȡ������"�ӿ�
        If blnTransMedicare Then Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicSwap, False, mintInsure)
    End If
    Call SaveErrLog
    Exit Function
errH:
    If ErrCenter = 1 Then Resume
End Function


Private Function VerifyPati(ByVal lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ҽ�����֤����
    '����:���˺�
    '����:2011-08-27 12:58:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID1 As Long
    lng����ID1 = lng����ID '����Identify�ӿ����޸ĸñ����󷵻���ֵ
    '���أ�0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);8����ID,24��������(1=��������),25������������
    mstrYBPati = gclsInsure.Identify(id�����շ�, lng����ID, mintInsure)
    If mstrYBPati = "" Then Exit Function
    If UBound(Split(mstrYBPati, ";")) < 8 Then GoTo ErrCancelYb:
    If Val(Split(mstrYBPati, ";")(8)) = 0 Then GoTo ErrCancelYb:
    
    '��ȡ������Ϣ
    lng����ID1 = Val(CLng(Split(mstrYBPati, ";")(8)))
    If lng����ID <> lng����ID1 And lng����ID1 <> 0 And lng����ID <> 0 Then
        MsgBox "ҽ����֤�Ĳ�����֮ǰ��ȡ�Ĳ��˲���ͬһ������!", vbInformation, gstrSysName
        GoTo ErrCancelYb: Exit Function
    End If
    '��ʼҽ������
    Call initInsurePara(lng����ID)
    VerifyPati = True
    Exit Function
ErrCancelYb:
    Call YBIdentifyCancel: mintInsure = 0: mstrYBPati = ""
End Function
Private Function SaveBill(ByRef strSaveNos As String, _
    Optional ByRef strModiNos As String, _
    Optional ByRef blnSaveBill As Boolean, _
    Optional ByRef blnNotPayWin As Boolean, _
    Optional bytReturnMode As ExitMode = EM_�շ����, _
    Optional bln���� As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���浱ǰ����ĵ���(�������շѡ����ۡ��������)
    '����:strSaveNos-�����ѳɹ�����ĵ��ݺţ���ʽΪ"'AAA','BBB',..."
    '       cur�ѽɺϼ�-���strSaveNOs�������ѱ���ɹ��ĵ���ʵ���ѽɵ��ֽ�
    '       strModiNOs -�޸ĵ��Ƕ൥���շ��е�һ��ʱ�����ظö��ŵ��ݵ�����NO����ʽ��"'AAA','BBB',..."
    '       blnSaveBill-�Ƿ񵥾��Ѿ�����ɹ�
    '       blnNotPayWin-�������շѽ���
    '       bytReturnMode As Byte ' '0-�����շ����,1-��ͣ�շ�;2-���������շ�;3-��������
    '����:�շѳɹ��򵥾ݱ���湦,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-26 17:28:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '     *** ҽ���շ�ʱ,����ʱ����Ϊ���۵�,�ڽ���ǰ��תΪ�շѵ�,�Ա������ҩƷ���ʱ��ȴ�ͬһ�����ҽ��������������� ***
    Dim lng��ӡID As Long, lngҽ��ID As Long, lng���ͺ� As Long
    Dim int��� As Integer, int�۸񸸺� As Integer, int�к� As Integer
    Dim lng����ID As Long, intҩƷ�д� As Integer, strҽ�Ƹ��� As String
    Dim dbl���� As Double, dbl���� As Double, cur�ɿ� As Currency
    Dim strDeptIDs As String, strTmp As String, strDelBill As String, strBillNO As String
    Dim str�շѽ��� As String, str���ս��� As String, str�շѽ���У�� As String
    Dim arrSQL As Variant, arrPut As Variant, arrOTMSQL As Variant
    Dim blnֱ���շ� As Boolean, blnTrans As Boolean, blnTransMedicare As Boolean
    Dim i As Integer, j As Integer, p As Integer, strSQL As String
    Dim CurOneCard As Currency, dblOneCardBalance As Double
    Dim strCardNo  As String, intCardType As Integer, strTransFlow As String
    Dim str��ҩ��̬ As String
    Dim strStuffDept As String          '�Զ����ϵĲ���
    Dim strAdvance As String            'ҽ�����㷵�ص���Ϣ:"���㷽ʽ|������||....."
    Dim blnPriceSaved As Boolean        'ҽ���շ�ʱ�Ƿ��Ѵ�Ϊ���۵�,������תΪ�շѵ���ҽ����������ʧ�ܻ��˺�ɾ�����۵�
    Dim blnMedicareCheck As Boolean     '�Ƿ�ִ��ҽ������У��
    Dim strBalanceIDs As String         '���е��ݵĽ���ID����������ҽ���ӿ�
    Dim strInvoice As String            '��ǰ����ʹ�õ�Ʊ�ݺţ�����ҽ��һ�ŵ���ֻ��һ��Ʊ�����
    Dim cllRqure As Collection
    Dim rsSqure As ADODB.Recordset
    Dim str����IDs As String
    Dim blnӦ���� As Boolean
    Dim dblӦ�ɶ� As Double, lng������� As Long
    Dim cllPutout As Collection '�Զ�����
    Dim cllPro As Collection, cllDelete As Collection, cllPageInfor As Collection
    Dim cur�ѽɺϼ� As Currency
    
    Set mCllWindows = New Collection
    
    strSaveNos = "": cur�ѽɺϼ� = 0: strModiNos = ""
    Err = 0: On Error GoTo Errhand:
    If cboҽ�Ƹ���.ListIndex <> -1 Then
        strҽ�Ƹ��� = Mid(cboҽ�Ƹ���.Text, 1, InStr(1, cboҽ�Ƹ���, "-") - 1)
    End If
    strInvoice = Trim(txtInvoice.Text)
    
    arrOTMSQL = Array()
    '�޸Ĺ���ʱ,�Ƿ��޸�ҽ������
    If mstrInNO <> "" Then
        Call BillisAdviceMoney(mstrInNO, IIf(mbytInFun = 2, 2, 1), lngҽ��ID, lng���ͺ�)
    End If
    If mlng����ҽ�� <> 0 And lngҽ��ID = 0 Then lngҽ��ID = mlng����ҽ��
    
    blnSaveBill = False
    dblӦ�ɶ� = 0: lng������� = 0
    Set cllPutout = New Collection: Set cllPro = New Collection
    Set cllPageInfor = New Collection
    '��ÿ�ŵ��ݶ���ִ�б���
    For p = 1 To mobjBill.Pages.Count
        int��� = 0: int�к� = 0: blnPriceSaved = False
        intҩƷ�д� = 0: strDeptIDs = "": strStuffDept = ""
        '��ǰ�շѵ��ݵĸ������
        If mbytInFun = 0 And Not mblnSaveAsPrice Then
            str���ս��� = GetMedicareStr(p)
            lng����ID = zlDatabase.GetNextId("���˽��ʼ�¼")
            If lng������� = 0 Then lng������� = lng����ID
            'If mlng������� <> 0 Then lng������� = mlng�������
            str����IDs = str����IDs & "," & lng����ID
        End If
        '����ÿ���շѵ��ݵĵ��ݺ�
        If mobjBill.Pages(p).NO = "" Then
            'Ϊ����ʧ�ܺ�����ʶ��,���Ķ���NO
            Select Case mbytInFun
                Case 0, 1 '�շѵ������۵�
                    strBillNO = zlDatabase.GetNextNo(13)
                Case 2  '������ʵ�
                    strBillNO = zlDatabase.GetNextNo(14)
            End Select
            blnֱ���շ� = True
        Else
            blnֱ���շ� = False
            strBillNO = mobjBill.Pages(p).NO
        End If
        '��ҪΪ��Ϣ������,Ϊÿҳ����ĵ��ݺ�
        mobjBill.Pages(p).�շѵ��� = strBillNO
        If p = 1 Then
            mobjBill.NO = strBillNO
            gstrModiNO = strBillNO
        End If
        
        arrSQL = Array() '�൥��ʱ,���ŵ����ύ
        If Not blnֱ���շ� Then
            '1.�շ��µ��ݹ���ʱ,��ȡ�Ļ��۵��շ�
            '��ȻZl_���˻����շ�_Insertû�и���ҽ����Ϣ,���ڸ��ݲ�����ȡ�Ļ��۵�ʱִ����zl_���ﻮ�ۼ�¼_Update,�Ѹ���
            '---------------------------------------------------------------
            If Not mblnSaveAsPrice Then
                'Zl_���˻����շ�_Insert
                strSQL = "Zl_���˻����շ�_Insert("
                '  No_In         ������ü�¼.NO%Type,
                strSQL = strSQL & "'" & strBillNO & "',"
                '  ����id_In     ������ü�¼.����id%Type,
                strSQL = strSQL & "" & ZVal(mobjBill.����ID) & ","
                '  ������Դ_In   Number,
                strSQL = strSQL & "" & gint������Դ & ","
                '  ���ʽ_In   ������ü�¼.���ʽ%Type,
                strSQL = strSQL & "'" & strҽ�Ƹ��� & "',"
                '  ����_In       ������ü�¼.����%Type,
                strSQL = strSQL & "'" & mobjBill.���� & "',"
                '  �Ա�_In       ������ü�¼.�Ա�%Type,
                strSQL = strSQL & "'" & mobjBill.�Ա� & "',"
                '  ����_In       ������ü�¼.����%Type,
                strSQL = strSQL & "'" & mobjBill.���� & "',"
                '  ���˿���id_In ������ü�¼.���˿���id%Type,
                strSQL = strSQL & "" & IIf(mobjBill.Pages(p).ҽ����� > 0, "Null", ZVal(mobjBill.����ID, , mobjBill.Pages(p).��������ID)) & ","
                '  ��������id_In ������ü�¼.��������id%Type,
                strSQL = strSQL & "" & ZVal(mobjBill.Pages(p).��������ID) & ","
                '  ������_In     ������ü�¼.������%Type,
                strSQL = strSQL & "'" & mobjBill.Pages(p).������ & "',"
                '  ���ս���_In   Varchar2,
                If mstrYBPati <> "" And str���ս��� <> "" Then
                    strSQL = strSQL & "'" & str���ս��� & "',"
                Else
                    strSQL = strSQL & "NULL,"
                End If
                '  ����id_In     ������ü�¼.����id%Type,
                strSQL = strSQL & "" & lng����ID & ","
                '  ����ʱ��_In   ������ü�¼.����ʱ��%Type,
                strSQL = strSQL & "To_Date('" & Format(mobjBill.����ʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
                '  ����Ա���_In ������ü�¼.����Ա���%Type,
                strSQL = strSQL & "'" & UserInfo.��� & "',"
                '  ����Ա����_In ������ü�¼.����Ա����%Type,
                strSQL = strSQL & "'" & UserInfo.���� & "',"
                '  ��ҩ����_In   ������ü�¼.��ҩ����%Type := Null,
                strSQL = strSQL & "'" & tbsBill.Tabs(p).Tag & "',"
                '  �Ƿ���_In   ������ü�¼.�Ƿ���%Type := 0,
                strSQL = strSQL & "" & chk����.Value & ","
                '  �Ǽ�ʱ��_In   ������ü�¼.�Ǽ�ʱ��%Type := Null,
                strSQL = strSQL & "" & "NULL" & ","
                '  �������_In   ����Ԥ����¼.�������%Type := Null
                strSQL = strSQL & "" & lng������� & ")"
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "0;" & strSQL
                            
                '��ȡ�Զ���ҩ�Ķ��ҩ��
                If gbln�շѺ��Զ���ҩ Then
                    strDeptIDs = strDeptIDs & "," & Get��ҩ����IDs(strBillNO)
                End If
                '���ÿ�ŵ����ռ����Ϸ��ϲ���,�Ա��Զ�����,�Ƿ��Ǹ������ò�����SQL���ж�
                If gbln�����Զ����� Then
                    strStuffDept = strStuffDept & "," & Get��ҩ����IDs(strBillNO, "'4'")
                End If
                
                'ͨ�����۵��շѵķ�ʽ��ȡ�˹Һŷ����ķ���,����ɾ���÷���
                If strBillNO = mstrCardNO Then mstrCardNO = ""
            '��ȡ���۵��շ�,���Ա���Ϊ���۵�,������ҽ���ı���
            ElseIf mstrYBPati <> "" And mobjBill.����ID <> 0 Then
                '���»��۵�������Ϣ
                gstrSQL = "zl_���ﻮ�ۼ�¼_Update(" & mintInsure & "," & mobjBill.����ID & ",'" & strBillNO & "',1)"
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "0;" & gstrSQL
            End If
            
        ElseIf blnֱ���շ� Then
            '2.ֱ������ĵ�������,�����������޸�,�������շ�(���շѽ��汣��Ϊ���۵�),����,����
            '---------------------------------------------------------------
            For Each mobjBillDetail In mobjBill.Pages(p).Details
                If mobjBillDetail.���� <> 0 Then
                    For Each mobjBillIncome In mobjBillDetail.InComes
                        int��� = int��� + 1 '��ǰ��¼���
                        '1.��������---------------------------------------------------------------
                        With mobjBill                              'ҽ���շ�ʱ,����ʱ����Ϊ���۵�,�ڽ���ǰ��תΪ�շѵ�
                            If mbytInFun = 0 And Not mblnSaveAsPrice And mstrYBPati = "" Then
                                gstrSQL = "zl_���������շ�_INSERT('" & strBillNO & "'," & int��� & "," & ZVal(.����ID) & "," & _
                                    IIf(gint������Դ = 2, 2, 1) & "," & ZVal(.��ʶ��) & ",'" & strҽ�Ƹ��� & "'," & _
                                    "'" & .���� & "','" & .�Ա� & "','" & .���� & "','" & IIf(mobjBillDetail.�ѱ� = "", .�ѱ�, mobjBillDetail.�ѱ�) & "'," & _
                                    .�Ӱ��־ & "," & ZVal(.����ID, , .Pages(p).��������ID) & "," & _
                                    ZVal(.Pages(p).��������ID) & ",'" & .Pages(p).������ & "',"
                            ElseIf mbytInFun = 1 Or (mbytInFun = 0 And (mblnSaveAsPrice Or mstrYBPati <> "")) Then
                                gstrSQL = "zl_���ﻮ�ۼ�¼_INSERT('" & strBillNO & "'," & int��� & "," & ZVal(.����ID) & "," & _
                                    ZVal(.��ҳID) & "," & ZVal(.��ʶ��) & ",'" & strҽ�Ƹ��� & "'," & _
                                    "'" & .���� & "','" & .�Ա� & "','" & .���� & "','" & IIf(mobjBillDetail.�ѱ� = "", .�ѱ�, mobjBillDetail.�ѱ�) & "'," & _
                                    .�Ӱ��־ & "," & ZVal(.����ID, , .Pages(p).��������ID) & "," & _
                                    ZVal(.Pages(p).��������ID) & ",'" & .Pages(p).������ & "',"
                            ElseIf mbytInFun = 2 Then
                                gstrSQL = "zl_������ʼ�¼_INSERT('" & strBillNO & "'," & int��� & "," & _
                                    .����ID & "," & .��ʶ�� & ",'" & .���� & "','" & .�Ա� & "','" & .���� & "'," & _
                                    "'" & .�ѱ� & "'," & .�Ӱ��־ & "," & .Ӥ���� & "," & _
                                      ZVal(.����ID, , .Pages(p).��������ID) & "," & _
                                    ZVal(.Pages(p).��������ID) & ",'" & .Pages(p).������ & "',"
                            End If
                        End With
        
                        '2.�շ�ϸĿ����---------------------------------------------------------------
                        With mobjBillDetail
                            If .��� <> int�к� Then     '�����������
                                int�к� = .���
                                int�۸񸸺� = int���
                                '���´����������
                                If mobjBill.Pages(p).Details(.���).�������� = 0 Then
                                    For i = .��� + 1 To mobjBill.Pages(p).Details.Count
                                        If mobjBill.Pages(p).Details(i).�������� = .��� Then
                                            '������Ŀ�ж��������Ŀ(������)ʱ,ȡ��һ�����
                                            mobjBill.Pages(p).Details(i).�������� = int���
                                        End If
                                    Next
                                End If
                            End If
        
                            '�շѡ����۵�ҩƷ��,����ҩ����
                            If (mbytInFun = 0 Or mbytInFun = 1) And InStr(",5,6,7,", .�շ����) > 0 Then
                                If Set��ҩ����(p, mobjBillDetail) = False Then Exit Function
                            End If
                            
                            'ҽ��ֱ���շ�ʱ,��Ϊ���ݴ�Ϊ���۵�,�շ�ʱ��Ҫȡ��ҩ����
                            If mbytInFun = 0 And Not mblnSaveAsPrice And mstrYBPati <> "" Then tbsBill.Tabs(p).Tag = .��ҩ����
                            
                            dbl���� = .����
                            If InStr(",5,6,7,", .�շ����) > 0 And gblnҩ����λ Then
                                dbl���� = Format(.���� * .Detail.ҩ����װ, "0.00000")
                            End If
                            
                            gstrSQL = gstrSQL & .�������� & "," & .�շ�ϸĿID & ",'" & .�շ���� & "','" & .���㵥λ & "',"
                            If mbytInFun = 0 And Not mblnSaveAsPrice And mstrYBPati = "" Then
                                gstrSQL = gstrSQL & IIf(.������Ŀ��, 1, 0) & "," & ZVal(.���մ���ID) & ",'" & .��ҩ���� & "'," & IIf(.���� = 0, 1, .����) & "," & dbl���� & "," & IIf(.������, 8, .���ӱ�־) & ","
                            ElseIf mbytInFun = 1 Or (mbytInFun = 0 And (mblnSaveAsPrice Or mstrYBPati <> "")) Then
                                gstrSQL = gstrSQL & "'" & .��ҩ���� & "'," & IIf(.���� = 0, 1, .����) & "," & dbl���� & "," & .���ӱ�־ & ","
                            ElseIf mbytInFun = 2 Then
                                gstrSQL = gstrSQL & IIf(.���� = 0, 1, .����) & "," & dbl���� & "," & .���ӱ�־ & ","
                            End If
                            gstrSQL = gstrSQL & IIf(.ִ�в���ID = 0, "NULL", .ִ�в���ID) & ","
                           
                        End With
        
                        '3.������Ŀ����---------------------------------------------------------------
                        With mobjBillIncome
                            dbl���� = .��׼����
                            If InStr(",5,6,7,", mobjBillDetail.�շ����) > 0 And gblnҩ����λ Then
                                dbl���� = Format(.��׼���� / mobjBillDetail.Detail.ҩ����װ, gstrFeePrecisionFmt)
                            End If
                            gstrSQL = gstrSQL & IIf(int�۸񸸺� = int���, "NULL", int�۸񸸺�) & "," & .������ĿID & "," & _
                                    "'" & .�վݷ�Ŀ & "'," & dbl���� & "," & .Ӧ�ս�� & "," & .ʵ�ս�� & ","
                            If mbytInFun = 0 And Not mblnSaveAsPrice And mstrYBPati = "" Then
                                gstrSQL = gstrSQL & "NULL,"
                            End If
                        End With
        
                        '4.��������
                        '---------------------------------------------------------------
                        gstrSQL = gstrSQL & _
                                "To_Date('" & Format(mobjBill.����ʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                                "To_Date('" & Format(mobjBill.�Ǽ�ʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),'" & mstrInNO & "',"
                        If mobjBillDetail.�շ���� = "7" Then
                            str��ҩ��̬ = "'" & mobjBillDetail.Detail.��ҩ��̬ & "'"
                        Else
                            str��ҩ��̬ = "NULL"
                        End If
                        '��ҩ��̬_In       סԺ���ü�¼.����%Type := Null
                        
                        If mbytInFun = 0 And Not mblnSaveAsPrice And mstrYBPati = "" Then
                            '��ҽ���շ�,���Ҳ��ǻ���
                            gstrSQL = gstrSQL & lng����ID & "," & lng������� & ","
                            '�������ID
                            gstrSQL = gstrSQL & "'" & UserInfo.��� & "','" & UserInfo.���� & "'," & _
                                "'" & mobjBillDetail.ժҪ & "'," & chk����.Value & ",'|" & mobjBill.Pages(mintPage).�巨 & "'" & _
                                "," & str��ҩ��̬ & ")"
                                'ֻ�ڵ�һ�ŵ��ݵĵ�һ����¼ʱ����
                        ElseIf mbytInFun = 1 Or (mbytInFun = 0 And (mblnSaveAsPrice Or mstrYBPati <> "")) Then
                            '���ﻮ��,�շѹ��ܻ���,ҽ���շ�
                            gstrSQL = gstrSQL & "'" & UserInfo.���� & "'," & _
                                "'" & mobjBillDetail.ժҪ & "'," & ZVal(lngҽ��ID) & ",NULL,NULL,'|" & mobjBill.Pages(mintPage).�巨 & _
                                "',NULL,NULL," & gint������Դ & ",'" & mobjBillDetail.���ձ��� & "'," & _
                                "'" & mobjBillDetail.Detail.���� & "'," & IIf(mobjBillDetail.������Ŀ��, 1, 0) & "," & ZVal(mobjBillDetail.���մ���ID) & "," & _
                                str��ҩ��̬ & ",0," & IIf(mobjBillDetail.Detail.���� = -1 Or mobjBillDetail.Detail.���� = 0, "Null", mobjBillDetail.Detail.����) & "," & _
                                "NULL," & ZVal(mobjBill.����ID) & ")"
                        ElseIf mbytInFun = 2 Then
                            '�������
                            gstrSQL = gstrSQL & IIf(mbytBilling = 1, 1, 0) & ",'" & UserInfo.��� & "','" & UserInfo.���� & "'," & _
                                "NULL,'" & mobjBillDetail.ժҪ & "'," & ZVal(lngҽ��ID) & ",Null,Null,'|" & mobjBill.Pages(mintPage).�巨 & "'," & _
                                "NULL,NULL,1," & str��ҩ��̬ & ",0," & IIf(mobjBillDetail.Detail.���� = -1 Or mobjBillDetail.Detail.���� = 0, "Null", mobjBillDetail.Detail.����) & "," & _
                                ZVal(mobjBill.��ҳID) & "," & ZVal(mobjBill.����ID) & ")"
                        End If
        
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = mobjBillDetail.�շ�ϸĿID & ";" & gstrSQL
                    Next    'ÿһ��������Ŀ
                    
                    '��ÿһ���շѼ�¼�ռ�ҩƷִ�в���,������ʻ��۵�����˲���,��Oracle�����д���:zl_������ʼ�¼_Verify
                    '----------------------------------------------------------------------------------------------------------------
                    '�Զ���ҩ,���շ�ʱ�Ҳ��Ƿ��뷢ҩʱ                    '
                    With mobjBillDetail
                        If gbln�շѺ��Զ���ҩ And mbytInFun = 0 And Not mblnSaveAsPrice Then
                            If .ִ�в���ID <> 0 And InStr("5,6,7", .�շ����) > 0 Then
                                If InStr(strDeptIDs & ",", "," & .ִ�в���ID & ",") = 0 Then
                                    strDeptIDs = strDeptIDs & "," & .ִ�в���ID
                                End If
                            End If
                        End If
                        '�Զ�����,�շ��Ҳ��Ǳ���Ϊ���۵������������,���뷢ҩ������Ӱ������
                        If gbln�����Զ����� And ((mbytInFun = 0 And Not mblnSaveAsPrice) Or (mbytInFun = 2 And mbytBilling = 0)) Then
                                If .ִ�в���ID <> 0 And .�շ���� = "4" And .Detail.�������� Then
                                    If InStr(strStuffDept & ",", "," & .ִ�в���ID & ",") = 0 Then
                                        strStuffDept = strStuffDept & "," & .ִ�в���ID
                                    End If
                                End If
                        End If
                    End With
                End If
            Next            'ÿһ���շ���Ŀ
            
            '����ǰһ�ŵ��ݵ�ҩ��ID,�Ա���ŵ���ʱȷ����ҩ����
            If mobjBill.Pages.Count > 1 Then Call SaveDrugID(p)
                
        
            '�޸ĺ��˳�ԭ����(�޸Ķ��շѵ��е�һ��ʱ��Ҫ���˷���ͳһ�ش�)
            '--------------------------------------------------------------------------------------------------------
            If mstrInNO <> "" Then
                strDelBill = ""
                If mbytInFun = 0 And Not mblnSaveAsPrice Then
                    '�޸�ҽ���շѵ�,��ȻΪ������ȫ��,��Ϊ�޸ĵ���ʱ���ж����������ȫ��,�������޸�
                    strDelBill = "zl_�����շѼ�¼_DELETE('" & mstrInNO & "','" & UserInfo.��� & "','" & UserInfo.���� & "'," & _
                        "NULL,NULL,'" & zlStr.NeedName(cbo���㷽ʽ.Text) & "',0,To_Date('" & Format(mobjBill.�Ǽ�ʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
                        
                    '����Ƕ൥���շ��е�һ�����µ���������ԭ���ݵĴ�ӡID��,�Ա�һ���ش�
                    strTmp = GetMultiNOs(mstrInNO, lng��ӡID)
                    If UBound(Split(strTmp, ",")) = 0 Then
                        lng��ӡID = 0: strModiNos = ""
                    ElseIf lng��ӡID <> 0 Then
                        strModiNos = strTmp
                    End If
                ElseIf mbytInFun = 1 Or (mbytInFun = 0 And mblnSaveAsPrice) Then
                    strDelBill = "zl_���ﻮ�ۼ�¼_DELETE('" & mstrInNO & "')"
                ElseIf mbytInFun = 2 Then
                    strDelBill = "zl_������ʼ�¼_DELETE('" & mstrInNO & "',NULL,'" & UserInfo.��� & "','" & UserInfo.���� & "')"
                End If
            
                '������޸�ҽ���ĸ���,���µ�NO���ڸ�����
                If lngҽ��ID <> 0 And lng���ͺ� <> 0 Then
                    gstrSQL = "ZL_����ҽ������_Insert(" & lngҽ��ID & "," & lng���ͺ� & "," & IIf(mbytInFun = 2, 2, 1) & ",'" & strBillNO & "')"
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "0;" & gstrSQL
                End If
            End If
        End If
        '�շѺ��Զ���ҩ,���ʲ��Զ���ҩ,�շ��Ҳ��Ǳ���Ϊ���۵�,�����������
        '-----------------------------------------------------------------------
        If strDeptIDs <> "" Then
            arrPut = Array()
            strDeptIDs = Mid(strDeptIDs, 2)
            For i = 0 To UBound(Split(strDeptIDs, ","))
                ReDim Preserve arrPut(UBound(arrPut) + 1)
                arrPut(UBound(arrPut)) = "ZL_ҩƷ�շ���¼_������ҩ(" & Val(Split(strDeptIDs, ",")(i)) & ",8,'" & strBillNO & "','" & UserInfo.���� & "','" & UserInfo.���� & "','" & mobjBill.Pages(p).������ & "')"
            Next
        End If
        '�շѺ��Զ�����,���շ�(ֱ���շ�,���۵������շ�),�������ʱִ��
        If strStuffDept <> "" Then
            If strDeptIDs = "" Then arrPut = Array()
            strStuffDept = Mid(strStuffDept, 2)
            For i = 0 To UBound(Split(strStuffDept, ","))          '24-�շѴ������ϣ�25-���ʵ���������
                ReDim Preserve arrPut(UBound(arrPut) + 1)
                arrPut(UBound(arrPut)) = "zl_�����շ���¼_��������(" & Split(strStuffDept, ",")(i) & "," & IIf(mbytInFun = 0, 24, 25) & ",'" & strBillNO & "','" & UserInfo.���� & "','" & UserInfo.���� & "','" & UserInfo.���� & "',1,Sysdate)"
            Next
        End If
        
    
        'ִ�����SQL��估�ύҽ������,���ŵ���ʱ,ÿ�ŵ����ڶ����������ύ
        '--------------------------------------------------------------------------------------------------------------------------------
        If UBound(arrSQL) >= 0 Then
            '��SQL���а��շ�ϸĿID����
            For i = 0 To UBound(arrSQL) - 1
                For j = i + 1 To UBound(arrSQL)
                    If CLng(Split(arrSQL(j), ";")(0)) < CLng(Split(arrSQL(i), ";")(0)) Then
                        strTmp = CStr(arrSQL(j)): arrSQL(j) = arrSQL(i): arrSQL(i) = strTmp
                    End If
                Next
            Next
            
            'ҽ��ֱ���շ�ʱ,�ȱ���Ϊ���۵�,��תΪ�շѵ�
            '-------------------------------------------------------------------
            If blnֱ���շ� And mbytInFun = 0 And Not mblnSaveAsPrice And mstrYBPati <> "" Then
                '1.�ȱ��滮�۵�,���ύ�������Ա㲻����
                On Error GoTo errH
                For i = 0 To UBound(arrSQL)
                    zlAddArray cllPro, Mid(arrSQL(i), InStr(arrSQL(i), ";") + 1)
                Next
                blnPriceSaved = True
                '���»��۵��ı�����Ϣ(������Ŀ��,ҽ������ID,ͳ����)
                gstrSQL = "zl_���ﻮ�ۼ�¼_Update(" & mintInsure & "," & mobjBill.����ID & ",'" & strBillNO & "',0)"
                zlAddArray cllPro, gstrSQL
                    
                '���۵�תΪ�շѵ�
                 'Zl_���˻����շ�_Insert
                gstrSQL = "Zl_���˻����շ�_Insert("
                '  No_In         ������ü�¼.NO%Type,
                gstrSQL = gstrSQL & "'" & strBillNO & "',"
                '  ����id_In     ������ü�¼.����id%Type,
                gstrSQL = gstrSQL & "" & mobjBill.����ID & ","
                '  ������Դ_In   Number,
                gstrSQL = gstrSQL & "" & gint������Դ & ","
                '  ���ʽ_In   ������ü�¼.���ʽ%Type,
                gstrSQL = gstrSQL & "'" & strҽ�Ƹ��� & "',"
                '  ����_In       ������ü�¼.����%Type,
                gstrSQL = gstrSQL & "'" & mobjBill.���� & "',"
                '  �Ա�_In       ������ü�¼.�Ա�%Type,
                gstrSQL = gstrSQL & "'" & mobjBill.�Ա� & "',"
                '  ����_In       ������ü�¼.����%Type,
                gstrSQL = gstrSQL & "'" & mobjBill.���� & "',"
                '  ���˿���id_In ������ü�¼.���˿���id%Type,
                gstrSQL = gstrSQL & "" & ZVal(mobjBill.����ID, , mobjBill.Pages(p).��������ID) & ","
                '  ��������id_In ������ü�¼.��������id%Type,
                gstrSQL = gstrSQL & "" & ZVal(mobjBill.Pages(p).��������ID) & ","
                '  ������_In     ������ü�¼.������%Type,
                gstrSQL = gstrSQL & "'" & mobjBill.Pages(p).������ & "',"
                '  ���ս���_In   Varchar2,
                gstrSQL = gstrSQL & "" & IIf(str���ս��� <> "", "'" & str���ս��� & "'", "NULL") & ","
                '  ����id_In     ������ü�¼.����id%Type,
                gstrSQL = gstrSQL & "" & lng����ID & ","
                '  ����ʱ��_In   ������ü�¼.����ʱ��%Type,
                gstrSQL = gstrSQL & "To_Date('" & Format(mobjBill.����ʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
                '  ����Ա���_In ������ü�¼.����Ա���%Type,
                gstrSQL = gstrSQL & "'" & UserInfo.��� & "',"
                '  ����Ա����_In ������ü�¼.����Ա����%Type,
                gstrSQL = gstrSQL & "'" & UserInfo.���� & "',"
                '  ��ҩ����_In   ������ü�¼.��ҩ����%Type := Null,
                gstrSQL = gstrSQL & "NULL,"
                '  �Ƿ���_In   ������ü�¼.�Ƿ���%Type := 0,
                gstrSQL = gstrSQL & "" & chk����.Value & ","
                '  �Ǽ�ʱ��_In   ������ü�¼.�Ǽ�ʱ��%Type := Null,
                gstrSQL = gstrSQL & "" & "NULL" & ","
                '  �������_In   ����Ԥ����¼.�������%Type := Null
                gstrSQL = gstrSQL & "" & lng������� & ")"
            End If
            'ҽ���൥��һ�ν���ʱ�����е�����Ϊһ�������ύ
            If MCPAR.�൥��һ�ν��� And mstrYBPati <> "" And strDelBill = "" And Not mblnSaveAsPrice Then
                '1.���۵�ת�շ�
                zlAddArray cllPro, gstrSQL
                '2.������
                If mobjBill.Pages(p).����� <> 0 Then '44657
                    gstrSQL = "zl_�����շ����_Insert('" & strBillNO & "'," & mobjBill.Pages(p).����� & ",0,1)"
                    zlAddArray cllPro, gstrSQL
                End If
                '3.�շѺ��Զ���ҩ,�Զ�����
                If strDeptIDs <> "" Or strStuffDept <> "" Then
                    For i = 0 To UBound(arrPut)
                        zlAddArray cllPutout, arrPut(i)
                    Next
                End If
                'strBalanceIDs = IIf(strBalanceIDs = "", "", strBalanceIDs & ",") & lng����ID
            Else
                On Error GoTo errH
                    '�޸Ĺ�����ش���
                    '��ɾ��ԭ����,��Ϊ����Ԥ������Ҫ�Ȼ�ԭ
                    If strDelBill <> "" Then zlAddArray cllPro, strDelBill
                    
                    'a.��ҽ��ֱ���շ�
                    If Not (blnֱ���շ� And mbytInFun = 0 And Not mblnSaveAsPrice And mstrYBPati <> "") Then
                        'ɾ�����￨���۵�:���ŵ���ʱֻɾ��һ��(��Ϊͨ�����￨�Ŷ�����ʱ,���￨���۵��������շ�ϸĿ��,����Ҫɾ��)
                        If mbytInFun = 0 And mstrCardNO <> "" And strSaveNos = "" Then
                            gstrSQL = "zl_���ﻮ�ۼ�¼_Delete('" & mstrCardNO & "')"
                            zlAddArray cllPro, gstrSQL
                        End If
                        'ִ�������SQL���
                        For i = 0 To UBound(arrSQL)
                            'Call zlDatabase.ExecuteProcedure(Mid(arrSQL(i), InStr(arrSQL(i), ";") + 1), Me.Caption)
                            zlAddArray cllPro, Mid(arrSQL(i), InStr(arrSQL(i), ";") + 1)
                        Next
                        'b.ҽ��ֱ���շ�
                    Else
                       ' Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                        zlAddArray cllPro, gstrSQL
                    End If
                    '�շ���ɺ�Ĵ���
                    '-----------------------------------------------------
                    If mbytInFun = 0 And Not mblnSaveAsPrice Then
                        '����д��ʼƱ�ݺ��Ա�ҽ������ʱ�ϴ�,���ŷֱ��ӡʱ,��д��ͬ��,��ӡ����ʱ����д,ȡ����ӡ���ӡʧ�ܽ����
                        '�޸�ʱ,ֻ��д�µ��ݵĿ�ʼƱ�ݺ�,��Ϊҽ��ֻ���µ����ϴ�
                        If strInvoice <> "" And mblnPrint Then
                            gstrSQL = "Zl_Ʊ����ʼ��_Update('" & strBillNO & "','" & strInvoice & "',1)"
                            zlAddArray cllPro, gstrSQL
                        End If
                    
                        'ÿ�ŵ��ݴ������,�ý���ID������ɵ��շѼ�¼��ͬ
                        If mobjBill.Pages(p).����� <> 0 Then '44657
                            gstrSQL = "zl_�����շ����_Insert('" & strBillNO & "'," & mobjBill.Pages(p).����� & ",0,1)"
                            zlAddArray cllPro, gstrSQL
                        End If
                    End If
                    
                    '�շѺ��Զ���ҩ,�Զ�����
                    If strDeptIDs <> "" Or strStuffDept <> "" Then
                        For i = 0 To UBound(arrPut)
                            'Call zlDatabase.ExecuteProcedure(CStr(arrPut(i)), Me.Caption)
                            zlAddArray cllPutout, CStr(arrPut(i))
                        Next
                    End If
                    
                    '�޸Ĺ�����ش���
                    If strDelBill <> "" Then
                        '�շѣ��µ��ݹ�����ԭ���ݵĴ�ӡID��,�Ա�һ���ش�,��ʱ��δ����Ʊ��
                        If lng��ӡID <> 0 And mblnPrint Then
                            gstrSQL = "zl_�����շ�Ʊ��_Insert('" & strBillNO & "','',Null,'',Null," & lng��ӡID & ",0)"
                            zlAddArray cllPro, gstrSQL
                            'Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                        End If
                    End If
                    '��Ҫ�������п�ˢ���������ˢ������(�ݲ�֧��,�ķѷ�ʽ,�ڳ�������ж�)
                On Error GoTo 0
            End If
            strBalanceIDs = IIf(strBalanceIDs = "", "", strBalanceIDs & ",") & lng����ID
            cllPageInfor.Add Array(lng����ID, strBillNO), "K" & p
            
            '�ύ�ɹ������ۼ�
            If mbytInFun = 0 And Not mblnSaveAsPrice Then
                cur�ѽɺϼ� = cur�ѽɺϼ� + mobjBill.Pages(p).Ӧ�ɽ��
            End If
            strSaveNos = strSaveNos & ",'" & strBillNO & "'"
            If Left(strSaveNos, 1) = "," Then strSaveNos = Mid(strSaveNos, 2)
            '���뵥����ʷ��¼(�������͵���)
            cboNO.AddItem strBillNO, 0
            For i = cboNO.ListCount - 1 To 10 Step -1
                cboNO.RemoveItem i 'ֻ��ʾ10��
            Next
        End If
    Next  '��һ�ŵ���
    On Error GoTo errH:
    '�ȱ��浥��
    Dim blnAffair As Boolean, strSaveCuessNos As String
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
   
    If zlInsureClinicSwap(cllPageInfor, lng�������, strInvoice, strDelBill <> "", _
        strBalanceIDs, strSaveNos, strSaveCuessNos, blnAffair) = False Then
        If Not blnAffair Then gcnOracle.RollbackTrans
        If strSaveCuessNos <> "" Then blnSaveBill = True:
        If strSaveNos <> "" Then bytReturnMode = 2
        Exit Function
    End If
    If blnAffair = False Then gcnOracle.CommitTrans
    
    blnSaveBill = True
    
    '���˵��Զ����ϴ���,112720
    If mbytInFun = 2 And mbytBilling = 0 And cllPutout.Count > 0 Then
        Err = 0: On Error GoTo ErrPutOut:
        zlExecuteProcedureArrAy cllPutout, Me.Caption
        SaveBill = True: Exit Function
    End If
    
    If mblnSaveAsPrice Then SaveBill = True: Exit Function
    blnTrans = False
    If mbytInFun <> 0 Then SaveBill = True: Exit Function
    '�������ʽ
    If blnNotPayWin Then SaveBill = True: Exit Function
    Dim dbl����Ӧ�� As Double
    mlng������� = lng�������
    Dim frmNew As frmChargePayMentWin
    Set frmNew = New frmChargePayMentWin
    If frmNew.zlChargeWin(Me, 0, mlngModul, mstrPrivs, mlngShareUseID, mstrUseType, lng�������, strBalanceIDs, strSaveNos, mobjBill.����ID, mintInsure, mobjBill.����, mobjBill.�Ա�, mobjBill.����, mobjBill.�ѱ�, mdbl�ɿ�, mdbl�Ҳ�, bytReturnMode, CDbl(mcurBillӦ��), bln����, mlngPreBrushCard, dbl����Ӧ��, mstrBalance) = False Then
        If Not frmNew Is Nothing Then Unload frmNew
        Exit Function
    End If
    If Not frmNew Is Nothing Then Unload frmNew
    
    mblnNotClearLedDisplay = True
    mbln�������� = False
    If mstrYBPati <> "" And bln���� Or mstrYBPati = "" And bln���� Then
        mbln�������� = True
        For i = 1 To mobjBill.Pages.Count
            mobjBill.Pages(i).Ӧ�ɽ�� = 0
        Next
        If grsTotal.RecordCount <> 0 Then grsTotal.MoveFirst
        dbl����Ӧ�� = 0
        Do While Not grsTotal.EOF
            '����:0-�ɿ�;1-�Ҳ�,2-��Ԥ��;����(mod 10:0-��ͨ����;1-ҽ������;2-������Ʒ;3-һ��ͨ)
            If Val(Nvl(grsTotal!����)) <> 11 Then
                '��ҽ�����ۼ�
                dbl����Ӧ�� = dbl����Ӧ�� + Val(Nvl(grsTotal!������))
            End If
            grsTotal.MoveNext
        Loop
        mobjBill.Pages(1).Ӧ�ɽ�� = dbl����Ӧ��
    End If
    
    '�Զ���ҩ�ͷ��ϴ���
    Err = 0: On Error GoTo ErrPutOut:
    zlExecuteProcedureArrAy cllPutout, Me.Caption
    SaveBill = True
    Exit Function
errH:
    If Err.Description Like "*��ǰ���㵥�۲�һ��*" Then
        If blnTrans Then gcnOracle.RollbackTrans
        If MsgBox("ĳЩ����ҩƷ�۸��ѷ����仯��Ҫ�Զ�����۸���", vbYesNo + vbQuestion + vbDefaultButton1, gstrSysName) = vbYes Then
            Call CalcMoneys
            Call ShowDetails
            Call ShowMoney
            Exit Function
        End If
     Else
        If blnTrans Then gcnOracle.RollbackTrans
        If ErrCenter() = 1 Then
            Resume
        End If
        If blnTransMedicare = False Then    '���ҽ���ɹ��ˣ���ɾ�����۵�������ʧ�ܿ�������
            Call DelMedicareTempNO(blnPriceSaved, strBillNO)
        End If
        Call SaveErrLog
    End If
    
    Exit Function
ErrPutOut:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function

Private Function zlInsureClinicSwap(ByVal cllPageInfor As Collection, _
    ByVal lng������� As Long, _
    ByVal strInvoice As String, _
    ByVal blnModifyBill As Boolean, _
    ByVal strBalanceIDs As String, _
    ByRef strSaveNos As String, _
    ByRef strSaveSucessNos As String, _
    Optional ByRef blnAffair As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ҽ������
    '���:blnModifyBill-�Ƿ��޸ĵ���
    '       strBalanceIDs:���ν��ʵ�ID,�ֱ��ö��ŷ���
    '       strSaveNos-����ĵ��ݺ�
    '����:strSaveNos-�����Ѿ�����ɹ��ĵ��ݺ�
    '       blnAffair-�Ƿ�������
    '       strSaveSucessNos-����ɹ���Ʊ��(�Ի�����Ч)
    '����:ҽ�����óɹ����ҽ��,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-20 17:15:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strBillNO As String, blnTrans As Boolean, blnTransMedicare As Boolean
    Dim p As Integer, str���ս��� As String, strAdvance As String, blnMedicareCheck As Boolean
    Dim strTmp As String, i As Long
    Dim strSuccesInsureNOs As String   'ҽ���ɹ��ĵ���
    Dim strNotSuccesInsureNOs As String   'ҽ���ɹ��ĵ���
    On Error GoTo errHandle
    '��ҽ��������true,���򷵼�
    blnAffair = False
    If mstrYBPati = "" Or mbytInFun <> 0 Then zlInsureClinicSwap = True: Exit Function
    
    '1. ����Ϊ���۵�
    If mblnSaveAsPrice Then
        For p = 1 To mobjBill.Pages.Count
            strBillNO = cllPageInfor("K" & p)(1)
            If blnAffair Then gcnOracle.BeginTrans
            '����Ϊ���۵�
            '���������ҽ��,�շ�ȷ��ʱʵ��ȴ����Ϊ���۵�:�����۵���ϸ,����Oracle������ִ��
            If mbytInFun = 0 And Not mnuFileSavePrice.Checked Then
                If Not gclsInsure.TranChargeDetail(1, strBillNO, 1, 0, "", , mintInsure) Then
                    'ɾ�����۵�(��������)
                    Call DelMedicareTempNO(True, strBillNO)
                Else
                    strSaveSucessNos = strSaveSucessNos & "," & strBillNO
                End If
            End If
            gcnOracle.CommitTrans
            blnAffair = True
        Next
        zlInsureClinicSwap = True
        Exit Function
    End If
    
    '2.ҽ���൥��һ�ν���ʱ�����е�����Ϊһ�������ύ
    If MCPAR.�൥��һ�ν��� And Not blnModifyBill And Not mblnSaveAsPrice Then
        If blnAffair Then gcnOracle.BeginTrans
        
        If MCPAR.ҽ���ӿڴ�ӡƱ�� And MCPAR.ҽ������Ʊ�� = False Then
            '38821
            'Ʊ����������(��Ϊ����HIS�Ĵ�ӡ��ҽ���ӿڴ�ӡ����������Ʊ������)
            gstrSQL = "zl_�����շ�Ʊ��_Insert('" & Replace(strSaveNos, "'", "") & "','" & strInvoice & "'," & ZVal(mlng����ID) & ",'" & UserInfo.���� & "'," & _
                      "To_Date('" & Format(mobjBill.�Ǽ�ʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),0,1)"
            zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
        End If
        
        '���ϸ����Ʊ��ʱ���浱ǰƱ��
        If Not gblnStrictCtrl Then
            zlDatabase.SetPara "��ǰ�շ�Ʊ�ݺ�", strInvoice, glngSys, 1121, InStr(1, mstrPrivs, ";��������;") > 0
        End If
        strAdvance = strBalanceIDs
        If Not gclsInsure.ClinicSwap(Val(Split(strBalanceIDs, ",")(0)), 0, 0, 0, 0, mintInsure, strAdvance) Then
            gcnOracle.RollbackTrans
            For i = 0 To UBound(Split(strSaveNos, ","))
                strBillNO = Replace(Split(strSaveNos, ",")(i), "'", "")
                For p = 1 To mobjBill.Pages.Count
                    If mobjBill.Pages(p).NO = strBillNO Then strBillNO = "": Exit For
                Next
                If strBillNO <> "" Then Call DelMedicareTempNO(True, strBillNO)
            Next
            blnAffair = True
            strSaveNos = ""
            Exit Function
        Else
            blnTransMedicare = True
        End If
        If strAdvance = strBalanceIDs Then strBalanceIDs = ""
         
        '���ݷ��صĽ��㷽ʽ���з�̯
        ' Zl_���������շ�_ҽ������
        gstrSQL = "Zl_���������շ�_ҽ������("
        '  ����id_In   ������ü�¼.����id%Type,
        gstrSQL = gstrSQL & "" & "NULL" & ","
        '  �������_In ����Ԥ����¼.�������%Type,
        gstrSQL = gstrSQL & "" & lng������� & ","
        '  ���ս���_In Varchar2
        '����:47409
        gstrSQL = gstrSQL & "" & IIf(strAdvance = "", "NULL", "'" & strAdvance & "'") & ")"
        zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
        gcnOracle.CommitTrans: blnTrans = False
        '����:47123
        Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicSwap, True, mintInsure)
        blnAffair = True
        zlInsureClinicSwap = True: Exit Function
   End If
   
    '---------------------------------------------------
    '3.�޸�ʱ,���˳�ԭ�շѵ���(�ķѷ�ʽ)
    blnTransMedicare = False
    strAdvance = ""
    If Not mblnSaveAsPrice And blnModifyBill Then
        strAdvance = mobjBill.Pages.Count & "|" & p
        If Not gclsInsure.ClinicDelSwap(Original.����ID, False, mintInsure, strAdvance) Then
            blnAffair = True
            gcnOracle.RollbackTrans:   Exit Function
        End If
        blnTransMedicare = True
        gcnOracle.CommitTrans: blnTrans = False: blnAffair = True
        Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, True, mintInsure)
    End If
    
    '4.�൥�ݵ���һ�ν���
    If MCPAR.�൥�ݵ�һ�ν��� Then
        strAdvance = strBalanceIDs & "|" & lng�������
        If Not gclsInsure.ClinicSwap(lng�������, 0, 0, 0, 0, mintInsure, strAdvance) Then
            Exit Function
        End If
        If strAdvance = strBalanceIDs & "|" & lng������� Then strAdvance = ""
        '���ݷ��صĽ��㷽ʽ���з�̯
        ' Zl_���������շ�_ҽ������
        gstrSQL = "Zl_���������շ�_ҽ������("
        '  ����id_In   ������ü�¼.����id%Type,
        gstrSQL = gstrSQL & "" & "NULL" & ","
        '  �������_In ����Ԥ����¼.�������%Type,
        gstrSQL = gstrSQL & "" & lng������� & ","
        '  ���ս���_In Varchar2
        gstrSQL = gstrSQL & IIf(strAdvance <> "", "'" & strAdvance & "'", "NULL") & ")"
        zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
        gcnOracle.CommitTrans: blnAffair = True
        '47123
         Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicSwap, True, mintInsure)
         zlInsureClinicSwap = True: Exit Function
    End If
    
    '5.�ֵ��ݵ�ҽ������
    '       ��Ϊ�����̶�Ϊҽ������,�������ƹ̶�Ϊҽ������(����ͳ�ﲻ��ȷ��),�Ժ�Ӧȥ���ò���
    strSuccesInsureNOs = ""
    For p = 1 To mobjBill.Pages.Count
           If blnAffair Then gcnOracle.BeginTrans
           blnTrans = True: blnTransMedicare = False
            If (GetMedicareSum(, p) <> 0 Or MCPAR.������봫����ϸ) Then
                strAdvance = mobjBill.Pages.Count & "|" & p
                If Not gclsInsure.ClinicSwap(Val(cllPageInfor("K" & p)(0)), GetMedicareSum(mstr�����ʻ�, p), _
                    GetMedicareSum("ҽ������", p), mobjBill.Pages(p).ȫ�Ը�, mobjBill.Pages(p).���Ը�, mintInsure, strAdvance) Then
                    blnAffair = True: gcnOracle.RollbackTrans:
                    strSaveNos = ""
                    If strSuccesInsureNOs <> "" Then
                        strSuccesInsureNOs = Mid(strSuccesInsureNOs, 2)
                        strSaveNos = "'" & Replace(strSuccesInsureNOs, ",", "','") & "'"
                        strNotSuccesInsureNOs = ""
                        For i = p To mobjBill.Pages.Count
                            strNotSuccesInsureNOs = strNotSuccesInsureNOs & "," & cllPageInfor("K" & i)(1)
                        Next
                        If strNotSuccesInsureNOs <> "" Then strNotSuccesInsureNOs = Mid(strNotSuccesInsureNOs, 2)
                        If ModifyNotInsureNOs(strNotSuccesInsureNOs, strSuccesInsureNOs) = False Then
                            Exit Function
                        End If
                        zlInsureClinicSwap = True
                    End If
                    Exit Function
                Else
                    blnTransMedicare = True
                End If
            End If
            str���ս��� = GetMedicareStr(p)
            blnMedicareCheck = zlInsureCheck(str���ս���, strAdvance)
            ' Zl_���������շ�_ҽ������
            gstrSQL = "Zl_���������շ�_ҽ������("
            '  ����id_In   ������ü�¼.����id%Type,
            gstrSQL = gstrSQL & "" & Val(cllPageInfor("K" & p)(0)) & ","
            '  �������_In ����Ԥ����¼.�������%Type,
            gstrSQL = gstrSQL & "NULL,"
            '  ���ս���_In Varchar2
            gstrSQL = gstrSQL & IIf(blnMedicareCheck, "'" & strAdvance & "'", "NULL") & ")"
            zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
            strSuccesInsureNOs = strSuccesInsureNOs & "," & cllPageInfor("K" & p)(1)
            gcnOracle.CommitTrans: blnTrans = False
            If blnTransMedicare Then Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicSwap, True, mintInsure)
            blnAffair = True
    Next
    zlInsureClinicSwap = True
    Exit Function
Errhand:
    
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If blnTrans Then
        Call ErrCenter
    Else
        If ErrCenter() = 1 Then
            Resume
        End If
    End If
    If blnTrans Then
        'ҽ����HIS����ͬһ������,HIS����ʧ��,��ҽ���������ϴ�,������Ҫ��"ȡ������"�ӿ�
        If blnTransMedicare Then Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicSwap, False, mintInsure)
    End If
    If blnTransMedicare = False Then    '���ҽ���ɹ��ˣ���ɾ�����۵�������ʧ�ܿ�������
        Call DelMedicareTempNO(False, strBillNO)
    End If
    Call SaveErrLog
End Function

Private Sub DelMedicareTempNO(ByVal blnPriceSaved As Boolean, ByVal strBillNO As String)
'ҽ��ֱ���շ�ʱ,ɾ��ǰһ�������ύ�Ļ��۵�
    If blnPriceSaved Then
        gstrSQL = "zl_���ﻮ�ۼ�¼_DELETE('" & strBillNO & "')"
        On Error GoTo errH
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
 

Private Sub ShowBillChargeFee(ByVal lng������� As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ�շѳɹ����쳣����
    '����:���˺�
    '����:2011-08-26 18:59:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl������ As Double, dblδ���� As Double
    Dim strInfor As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errH:

    gstrSQL = "" & _
    "   Select decode(a.��¼����,1,'Ԥ���',11,'Ԥ���',���㷽ʽ) as ���㷽ʽ,  " & _
    "             nvl(sum(decode(nvl(У�Ա�־,0),1, 1,0)* ��Ԥ��),0) as δ����," & _
    "             nvl(sum(decode(nvl(У�Ա�־,0),0,1,2,1,0)* ��Ԥ��),0) as ������" & _
    "   From ����Ԥ����¼ A " & _
    "   Where �������=[1]" & _
    "   Group by  decode(a.��¼����,1,'Ԥ���',11,'Ԥ���',���㷽ʽ) "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng�������)
    strInfor = ""
    With rsTemp
        dbl������ = 0
        Do While Not .EOF
            If Val(Nvl(rsTemp!������)) <> 0 Then
                strInfor = strInfor & vbCrLf & "    " & Nvl(rsTemp!���㷽ʽ) & ":" & Format(rsTemp!������, "0.00")
            End If
            dblδ���� = dblδ���� + Val(Nvl(rsTemp!δ����))
            dbl������ = dbl������ + Val(Nvl(rsTemp!������))
            .MoveNext
        Loop
    End With
    strInfor = "" & _
    "�쳣�շ�(��ע��������ȡ):" & vbCrLf & _
    "    ��ǰ����ȡ����:" & Format(dbl������, "0.00") & "Ԫ" & vbCrLf & _
    "    ��ǰ��δȡ����:" & Format(dblδ����, "0.00") & "Ԫ" & vbCrLf & _
    "    ��ȡ�ɹ��ĸ�����������:" & strInfor
    MsgBox strInfor, vbExclamation, gstrSysName
    '�������������ʾ
    Call ClearPayInfo
    mstrInNO = "": txtModi.Text = ""
    mlngFirstID = 0: mstrFirstWin = ""
    
    mstrPrePati = "": mlngPrePati = 0: mstrPreDoctor = ""
            
    Call ClearPatientInfo(True)
    Call InitCommVariable
    Call ClearTotalInfo
    
    Call ClearBillRows: Call ClearMoney
    Call SetDisible(True): Call NewBill
    If txtPatient.Enabled Then txtPatient.SetFocus

    If gbln�ۼ� Then
        txt�ۼ�.Text = Format(GetChargeTotal, "0.00")
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    gcnOracle.RollbackTrans
    Call SaveErrLog
End Sub



Private Sub Set�����˿�������(ByVal str������ As String, ByVal lng��������ID As Long)
'����:���ݿ����˻򿪵�����ID���ÿ������Ҽ�������,������������¼�
       '���ù�������CboSetIndex������ʽ����cbo_click�¼�
    
    Dim str�������� As String, lng��ԱID As Long
    
    'a.ҽ��ȷ������
    If gbyt����ҽ�� = 0 Then
        Call zlControl.CboSetIndex(cbo������.hWnd, cbo.FindIndex(cbo������, str������, True))  '������click�¼�
        
        If cbo������.ListIndex = -1 And str������ <> "" Then
            lng��ԱID = GetPersonnelID(str������, mrs������)
            cbo������.AddItem str������, 0
            cbo������.ItemData(cbo������.NewIndex) = lng��ԱID
            Call zlControl.CboSetIndex(cbo������.hWnd, cbo������.NewIndex)
        End If
                
        If cbo������.ListIndex <> -1 Then
            cbo��������.Clear
            Call FillDept(mlngDeptID, cbo������.ItemData(cbo������.ListIndex))
        End If
        
        Call zlControl.CboSetIndex(cbo��������.hWnd, cbo.FindIndex(cbo��������, lng��������ID))
        If cbo��������.ListIndex = -1 And lng��������ID > 0 Then
            str�������� = GET��������(lng��������ID, mrs��������)
            If str�������� <> "" Then
                cbo��������.AddItem str��������, 0
                cbo��������.ItemData(cbo��������.NewIndex) = lng��������ID
                Call zlControl.CboSetIndex(cbo��������.hWnd, cbo��������.NewIndex)
            End If
        End If
        
    'b.����ȷ��ҽ�����������
    Else
        Call zlControl.CboSetIndex(cbo��������.hWnd, cbo.FindIndex(cbo��������, lng��������ID))
        
        If cbo��������.ListIndex = -1 And lng��������ID > 0 Then
            str�������� = GET��������(lng��������ID, mrs��������)
            If str�������� <> "" Then
                cbo��������.AddItem str��������, 0
                cbo��������.ItemData(cbo��������.NewIndex) = lng��������ID
                Call zlControl.CboSetIndex(cbo��������.hWnd, cbo��������.NewIndex)
            End If
        End If
        
        If gbyt����ҽ�� = 1 And cbo��������.ListIndex <> -1 Then
            cbo������.Clear
            Call FillDoctor(lng��������ID)
        End If
        
        Call zlControl.CboSetIndex(cbo������.hWnd, cbo.FindIndex(cbo������, str������, True))
        If cbo������.ListIndex = -1 And str������ <> "" Then
            lng��ԱID = GetPersonnelID(str������, mrs������)
            cbo������.AddItem str������, 0
            cbo������.ItemData(cbo������.NewIndex) = lng��ԱID
            Call zlControl.CboSetIndex(cbo������.hWnd, cbo������.NewIndex)
        End If
    End If
    
    '�����˵�רҵ����ְ��
    If cbo������.ListIndex <> -1 And mobjBill.Pages(mintPage).������ <> "" And Not mrs������ Is Nothing Then
        mrs������.Filter = "ID=" & cbo������.ItemData(cbo������.ListIndex)
        If mrs������.RecordCount > 0 Then
            lblDuty.Caption = IIf(IsNull(mrs������!רҵ����ְ��), "", mobjBill.Pages(mintPage).������ & "רҵְ��:" & mrs������!רҵ����ְ��)
        Else
            lblDuty.Caption = ""
        End If
    Else
        lblDuty.Caption = ""
    End If
End Sub


Private Sub Set�����˿�������Click(ByVal str������ As String, ByVal lng��������ID As Long)
'����:���ݿ����˻򿪵�����ID���ÿ������Ҽ�������,����������¼�
'     ��Listindex=xʱ,���Listindex��ֵ�������x,�Ͳ��ᴥ������¼�,����Ҫ��API+Clickǿ�Ƶ���
    Dim i As Long
    
    If gbyt����ҽ�� = 0 Then
        Call zlControl.CboSetIndex(cbo������.hWnd, cbo.FindIndex(cbo������, str������, True)) '������click�¼�
        Call cbo������_Click
        
        'û�д��� ��������ID ��ʱ�������� cbo������_Click ȱʡ��Ϊ׼
        If lng��������ID <> 0 Then
            Call zlControl.CboSetIndex(cbo��������.hWnd, cbo.FindIndex(cbo��������, lng��������ID))
            Call cbo��������_Click
        End If
        
    Else
        '����ȷ��ҽ������Զ�������
        Call zlControl.CboSetIndex(cbo��������.hWnd, cbo.FindIndex(cbo��������, lng��������ID))
        Call cbo��������_Click
        
        'û�д��� ������ ��ʱ�������� cbo��������_Click ȱʡ��Ϊ׼
        If str������ <> "" Then
            Call zlControl.CboSetIndex(cbo������.hWnd, cbo.FindIndex(cbo������, str������, True)) '������click�¼�
            Call cbo������_Click
        End If
    End If
End Sub

Private Function ReadBill(ByVal strNo As String, ByVal bytFun As Byte, _
    Optional ByVal blnDelete As Boolean, Optional blnNoName As Boolean, _
    Optional blnShow As Boolean, Optional blnErrBill As Boolean) As Boolean
'���ܣ�1.��ȡ������ԭʼ���ݻ����˵���,2.��ȡ�����շ�,������˵���,3.��ȡҪ�����˵ĵ���
'���ã�Ŀǰ�����²�������
'      1.�Ữ�۵��շѻ���ʣ������䵥�ݺ��Ữ�۵��շѣ�ȷ��������ݺ��Զ���ȡ���۵��շѣ������շ�ʱ�л�������ҳʱ���¶����۵�
'      2.�鿴���������˷ѣ����ʵ���ʱ�����ݣ��������շѵ������۵������ʵ������ʻ��۵�
'������strNo=���ݺ�
'      bytFun=0:�շѵ�,1:���۵�,2:������ʵ�
'      blnDelete=�Ƿ�����˷ѻ�����(��������Ϊ׼����,�����㴦��)
'      blnShow=�Ƿ�����Ϊ�л����ݶ�ȡ(����ʾ����)
'      blnErrBill-��ʾ�쳣����
'���أ�blnNoName=���������Ƿ�Ϊ��
'˵������ȡҪ�˷ѵĵ���ʱ(�շ�),�ſ��������,������ݲ��������Ƿ���ʾ
'      ��Ϊ��β����˷�ʱ,ÿ�ζ����ܲ������,ԭʼ�����ʼ���˲��ꡣ
    Dim rsTmp As ADODB.Recordset
    Dim rs���� As ADODB.Recordset
    Dim i As Long, j As Long, k  As Long, intSign As Integer
    Dim strSQL As String, strSQL1 As String, strSQL2 As String
    Dim curBillʵ�� As Currency, curBillӦ�� As Currency
    Dim str�ѱ� As String, str��ҩ���� As String, lng����ID As Long
    Dim lng����ID As Long
    Dim strPayDrugWins As String 'ִ�в���ID|��ҩ����;ִ�в���IDn|��ҩ����n
    Dim strTemp As String, strҽ����� As String '�˷�ʱ��Ч:�ֺŷָ�
    On Error GoTo errH
    strPayDrugWins = ""
    '�շ�ʱ,Ҫô�ں󱸱���,Ҫô�����߱���
    '����ʱ,����һ�ŵ��ݼ��ں󱸱����������߱���,����Ϊ��;��������ֻ��һ����
    '��һ�ŵ��ݵ�������,�����Ϊ������,���ű����Ӳ�ѯ
    
    '��ȡ��������
    '----------------------------------------------------------------------------------------------------
    strҽ����� = ""
    If Not blnShow Then
        '�շѵ��ݶ���Ʊ��ֻ��ȡһ��Ʊ��
        strSQL = _
        " Select A.����ID,A.ʵ��Ʊ�� as Ʊ�ݺ�,A.����ID,0 as ��ҳID,A.��ʶ��,B.��������,B.����," & _
        "       A.����,A.�Ա�,A.����,A.�ѱ�,A.���ʽ ,0 as ���˲���ID,A.���˿���ID," & _
        "       A.��������ID,Nvl(A.�Ӱ��־,0) as �Ӱ��־,a.��������ID," & _
        "       Nvl(A.Ӥ����,0) as Ӥ����,A.������,A.������,A.����Ա����,A.����ʱ��,A.�Ǽ�ʱ��," & _
        "       B.ҽ�Ƹ��ʽ,Nvl(A.�Ƿ���,0) as �Ƿ���,A.�����־,Nvl(A.ҽ�����,0) as ҽ�����,A.ժҪ,A.��¼״̬" & _
        " From " & IIf(mblnNOMoved, zlGetFullFieldsTable("������ü�¼"), "������ü�¼ A") & " ,������Ϣ B,��Ա�� C" & _
        " Where Rownum=1 And Nvl(A.����Ա����,A.������)=C.����" & _
        "       And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & vbNewLine & _
        "       And A.��¼����=" & IIf(bytFun = 2, 2, 1) & _
        "       And A.��¼״̬" & IIf(mblnDelete, "=2", " IN(0,1,3)") & _
                IIf(mstrTime <> "", " And A.�Ǽ�ʱ��=[2]", "") & _
        "       And NO=[1] And A.����ID=B.����ID(+)" & _
                IIf(bytFun = 0, " And A.����Ա���� is Not Null", "") & _
                IIf(bytFun = 1, " And A.����Ա���� is Null And A.������ is Not NULL", "") & _
                IIf(bytFun = 2 And mbytInState = 0 And mbytBilling = 0, " And A.����Ա���� is Not Null", "") & _
                IIf(bytFun = 2 And mbytInState = 0 And mbytBilling = 1, " And A.����Ա���� is Null And A.������ is Not NULL", "") & _
                IIf(bytFun = 2 And mbytInState = 0 And mbytBilling = 2, " And A.����Ա���� is Null And A.������ is Not NULL", "")
        '��סԺ������ȡ���۵�ʱ��������ȡ���﷢���ĵ���
        If bytFun = 1 And blnDelete = False And blnShow = False Then strSQL = strSQL & IIf(gint������Դ = 2, " ", " And �����־ In(1,3,4)")
        
        If mstrTime <> "" Then
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo, CDate(mstrTime))
        Else
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo)
        End If
        If rsTmp.EOF Then
            MsgBox "û�з���ָ���ĵ��ݣ�", vbInformation, gstrSysName
            Exit Function
        ElseIf bytFun = 1 And Not mblnDoing And Not IsNull(rsTmp!����) And txtPatient.Text <> "" Then
            '�ж��Ƿ���ͬ���ˣ���Ҫʹ�õĲ�����Ϣ
            If txtPatient.Text <> rsTmp!���� Then
                If MsgBox("�����в���Ϊ""" & rsTmp!���� & """���뵱ǰ���˲������Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            End If
        End If
        
        If zlStr.IsHavePrivs(mstrPrivs, "���п���") = False And mlngDeptID > 0 Then
            If Val(Nvl(rsTmp!��������ID)) <> mlngDeptID Then
                MsgBox "��û��Ȩ�޶�ȡ�������ҿ����ĵ��ݣ�", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        Original.����ID = Nvl(rsTmp!����ID, 0) '����ҽ�������˷�,һ��ͨ�����޸�
        If mbytBillSource <> 4 Then mbytBillSource = Val("" & rsTmp!�����־)   'ֻҪ��һ�������,����Ϊȫ������쵥��
        
    
        '���������Ϣ��ȡ:�������ڻ��۵��շ�,�Զ���ȡ���ŵ���ʱ����
        '����:30717
        If Not IsNull(rsTmp!�Ǽ�ʱ��) Then
            mobjBill.�Ǽ�ʱ�� = CDate(Format(rsTmp!�Ǽ�ʱ��, "yyyy-mm-dd HH:MM:SS"))
        End If
        If Val(Nvl(rsTmp!��¼״̬)) = 0 And gbytUnRegevent <> 0 Then
            mobjBill.����ID = Val(Nvl(rsTmp!����ID, mobjBill.����ID))
            mobjBill.��ҳID = Val(Nvl(rsTmp!��ҳID, mobjBill.��ҳID))
            mobjBill.��ʶ�� = Nvl(rsTmp!��ʶ��, mobjBill.��ʶ��)
        Else
            mobjBill.����ID = Val("" & rsTmp!����ID)
            mobjBill.��ҳID = Val("" & rsTmp!��ҳID)
            mobjBill.��ʶ�� = Nvl(rsTmp!��ʶ��, 0)
        End If
        lng����ID = mobjBill.����ID
        mobjBill.���� = ""            'IIf(gint������Դ = 2, "" & rsTmp!����, "")
        mobjBill.����ID = Val("" & rsTmp!���˲���ID)
        mobjBill.����ID = Val("" & rsTmp!���˿���ID)
        If mobjBill.�ѱ� = "" Then
            mobjBill.�ѱ� = Nvl(rsTmp!�ѱ�)
        End If
        mobjBill.Pages(mintPage).��������ID = Val("" & rsTmp!��������ID)
        mobjBill.Pages(mintPage).������ = "" & rsTmp!������
        mobjBill.Pages(mintPage).ҽ����� = Val("" & rsTmp!ҽ�����)
        txtPatient.Locked = (mobjBill.����ID <> 0 And "" & rsTmp!���� <> "�²���")    'Ϊ����ҽ���鿨,�ı��򲻱�Ϊ����״̬��ɫ
        cboSex.Locked = txtPatient.Locked
        txt����.Locked = txtPatient.Locked
        cbo���䵥λ.Locked = txtPatient.Locked
        txt�˷�ժҪ.Text = Nvl(rsTmp!ժҪ)
        
        If Not mblnDoing Then
            If Not IsNull(rsTmp!Ʊ�ݺ�) Then txtInvoice.Text = rsTmp!Ʊ�ݺ�: txtInvoice.SelStart = Len(txtInvoice.Text) '�в���ʾ,���۵���û�е�
            
            
            mobjBill.���� = Nvl(rsTmp!����)
            '75259�����ϴ�,2014-7-10������������ɫ����
            Call SetPatiColor(txtPatient, Nvl(rsTmp!��������), IIf(IsNull(rsTmp!����), txtPatient.ForeColor, vbRed))
            mobjBill.�Ա� = Nvl(rsTmp!�Ա�)
            'mobjBill.���� = Nvl(rsTmp!����)
            
            '��������
            If mbytInFun = 0 And chkCancel.Value = 0 And (IsNull(rsTmp!����) Or IIf(mlngPrePati = 0, mstrPrePati = mobjBill.����, mlngPrePati = mobjBill.����ID)) Then
                'ͬһ������:����������ͬ����
                
                If IsNull(rsTmp!����) Then
                    blnNoName = True
                    If Val(Nvl(rsTmp!��¼״̬)) = 0 And mstrPrePati = "" Then
                            
                    Else
                        txtPatient.Text = mstrPrePati 'ȱʡΪ��һ����������
                    End If
                Else
                    txtPatient.Text = Nvl(rsTmp!����)
                End If
            Else
                '��ͬ�Ĳ���
                txtPatient.Text = Nvl(rsTmp!����)
                '���˺�:22343,51670
                If Not (mbytInFun = 0 And gTy_Module_Para.byt�ɿ���� = 1) _
                    Or mstrPrePati = "" Then
                    mstrPrePati = "": mlngPrePati = 0: mstrPreDoctor = ""
                    Call ClearPatientInfo
                    Call ClearTotalInfo
                    Call InitCommVariable
                    Call ClearMoney
                End If
            End If
            
            Call zlControl.CboSetText(cboSex, "" & rsTmp!�Ա�)
            Call LoadOldData("" & rsTmp!����, txt����, cbo���䵥λ)
            '���˺�:24348,������ִ��ClearPatientInfo���������,���Ӧ�ý�����mobjBill.���� = Nvl(rsTmp!����),����������Ŷ�.
            mobjBill.���� = Nvl(rsTmp!����)
            
            txt�����.Text = Nvl(rsTmp!��ʶ��)
            
            If Nvl(rsTmp!�����־, 0) = 2 Or bytFun = 2 Or Not IsNull(rsTmp!ҽ�Ƹ��ʽ) Then
                cboҽ�Ƹ���.ListIndex = cbo.FindIndex(cboҽ�Ƹ���, Nvl(rsTmp!ҽ�Ƹ��ʽ), True)
                If cboҽ�Ƹ���.ListIndex = -1 And Not IsNull(rsTmp!ҽ�Ƹ��ʽ) Then
                    cboҽ�Ƹ���.AddItem "0-" & rsTmp!ҽ�Ƹ��ʽ, 0
                    cboҽ�Ƹ���.ListIndex = cboҽ�Ƹ���.NewIndex
                End If
            Else
                cboҽ�Ƹ���.ListIndex = GetCboIndexByCode(cboҽ�Ƹ���, "" & rsTmp!���ʽ)
                If cboҽ�Ƹ���.ListIndex = -1 And Not IsNull(rsTmp!���ʽ) Then
                    cboҽ�Ƹ���.AddItem rsTmp!���ʽ & "-" & GetMedPayModeName(rsTmp!���ʽ), 0
                    cboҽ�Ƹ���.ListIndex = cboҽ�Ƹ���.NewIndex
                ElseIf cboҽ�Ƹ���.ListIndex = -1 Then
                    cboҽ�Ƹ���.ListIndex = cbo.FindIndex(cboҽ�Ƹ���, mstr���ʽ, True)
                End If
            End If
            
            If bytFun = 2 Then
                cboBaby.ListIndex = IIf(Val("" & rsTmp!Ӥ����) > cboBaby.ListCount - 1, 0, Val("" & rsTmp!Ӥ����))
                cboBaby.Enabled = mbytInState = 0 And mbytBilling <> 2
            End If
            
            txtDate.Text = Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm:ss")
                        
            If Not rsTmp!����ID Is Nothing Then Call LoadFeeInfor(Val("" & rsTmp!����ID), blnDelete)
            
            If Nvl(rsTmp!�Ƿ���, 0) = 1 Then chk����.Value = 1: chk����.Visible = True
            mblnDo = False: chk�Ӱ�.Value = Nvl(rsTmp!�Ӱ��־, 0): mblnDo = True
        End If
    End If
    
    '��������,������
    Call Set�����˿�������(mobjBill.Pages(mintPage).������, mobjBill.Pages(mintPage).��������ID)
    
    '�շѶ����۵�ʱ��Ŀǰ�����޸Ŀ����˺Ϳ�������,������ҽ�����͹����ġ�
    If mbytInFun = 0 And mbytInState = 0 And chkCancel.Value = 0 Then
        cbo������.Locked = False
        cbo��������.Locked = False
        
        If mobjBill.Pages(mintPage).ҽ����� <> 0 Then
            If cbo������.ListIndex <> -1 Then cbo������.Locked = True
            If cbo��������.ListIndex <> -1 Then cbo��������.Locked = True
        End If
    End If
    '��ȡ���㷽ʽ
    '----------------------------------------------------------------------------------------------------
    If bytFun = 0 And Not blnShow Then
        '��ȡ���㷽ʽ
        Call ReadBalance(CLng(rsTmp!����ID), blnDelete)
    End If
    
    '��ȡ�����շ�ϸĿ����:���뷢ҩʱû��ҩ��
    '---------------------------------------------------------------------------------------------
    If blnDelete Then
        '�˷�ʱ�����Ǻ󱸱�,ǰ��Ĳ����ѽ���
        '��ȡ׼����,������Ӧ�ս��,ʵ�ս��(���=ʣ����*(׼����/ʣ����))
        '��ȡ������ԭʼ��¼�ķ���ID
        Dim strTableNo As String
        mblnHaveExcuteData = zlCheckIsExcuteData(strNo, IIf(bytFun = 2, 2, 1))     '60735
        
        '���˺�45685,58077
        strTableNo = " With ׼����  as ( " & _
        "            Select  A.����ID,Sum(Nvl(A.����,1)*A.ʵ������" & IIf(gblnҩ����λ, "/Nvl(B." & gstrҩ����װ & ",1)", "") & ") as ׼������" & _
        "            From ҩƷ�շ���¼ A,ҩƷ��� B " & _
        "           Where A.NO=[1]    And Mod(A.��¼״̬,3)=1  " & _
        "                       And (A.����=[4] or A.����=[5]) And A.����� is NULL  " & _
        "                       And A.ҩƷID=B.ҩƷID(+)  " & _
        "           Group by A.����ID" & _
        "           Union ALL    "
        'ȡ���ƵĲ����˷�
       If mblnHaveExcuteData Then
            '60735:��ҽ��ִ�мƼ��д�������ʱ,��ҽ��ִ�мƼ���ȡ��
            '77686,���ϴ�,2014/9/18,�����������
            strTableNo = strTableNo & _
            " Select Max(ID) As ����id, Decode(Sign(Sum(����)), -1, 0, Sum(����)) As ׼����" & vbNewLine & _
            " From ( Select Decode(a.��¼״̬, 2, 0, a.Id) As ID, a.ҽ����� As ҽ��id, a.�շ�ϸĿid, Nvl(a.����, 1) * Nvl(a.����, 1) As ����," & vbNewLine & _
            "              Decode(a.��¼״̬, 2, 0, Nvl(a.����, 1) * Nvl(a.����, 1)) As ԭʼ����" & vbNewLine & _
            "       From ������ü�¼ A, ����ҽ����¼ M" & vbNewLine & _
            "       Where a.ҽ����� = m.Id And Instr(',C,D,F,G,K,', ',' || m.������� || ',') = 0 And Instr('5,6,7', a.�շ����) = 0 And" & vbNewLine & _
            "             a.No = [1] And a.��¼���� = [3] And a.��¼״̬ In (1, 2, 3)��and A.�۸񸸺� is null " & vbNewLine & _
            "          And Not Exists" & _
            "                (Select 1 From ����ҽ������ Where a.ҽ����� = ҽ��id and a.No = NO and Mod(a.��¼����, 10) = ��¼����)" & _
            "       Union All" & vbNewLine & _
            "       Select a.Id, a.ҽ����� As ҽ��id, a.�շ�ϸĿid, -1 * b.���� As ��ִ��, 0 ԭʼ����" & vbNewLine & _
            "       From ������ü�¼ A, ҽ��ִ�мƼ� B, ����ҽ����¼ M" & vbNewLine & _
            "       Where a.ҽ����� = b.ҽ��id And a.�շ�ϸĿid = b.�շ�ϸĿid + 0 And a.ҽ����� = m.Id And Instr(',C,D,F,G,K,', ',' || m.������� || ',') = 0" & vbNewLine & _
            "           And Instr('5,6,7', a.�շ����) = 0" & vbNewLine & _
            "           And (Exists (Select 1  From ����ҽ��ִ��  Where b.ҽ��id = ҽ��id And b.���ͺ� = ���ͺ� And b.Ҫ��ʱ�� = Ҫ��ʱ�� And Nvl(ִ�н��, 0) = 1)" & vbNewLine & _
            "                Or Exists (Select 1 From ����ҽ������ Where b.ҽ��id = ҽ��id And b.���ͺ� = ���ͺ� And Nvl(ִ��״̬, 0) = 1))" & vbNewLine & _
            "          And a.No = [1] And a.��¼���� = [3] And a.��¼״̬ In (1, 3)��and A.�۸񸸺� is null " & vbNewLine & _
            "          And Not Exists" & _
            "                (Select 1 From ����ҽ������ Where a.ҽ����� = ҽ��id and a.No = NO and Mod(a.��¼����, 10) = ��¼����)" & _
            "       ) Q1" & vbNewLine & _
            " Where Not Exists (Select 1 From ҩƷ�շ���¼ Where ����id = Q1.Id And instr( ',8,9,10,21,24,25,26,',','||����||',')>0) " & vbNewLine & _
            " Group by ҽ��ID,�շ�ϸĿID  Having Max(ID)<>0 )"
       Else
            'And A.��������=0 :61879,����������ȷ��,��������������ֻ��0-��������
            
            strTableNo = strTableNo & "" & _
             " Select Max(ID) as ����ID,decode(sign(Sum(����)),-1,0,Sum(����))  as ׼���� " & _
             " From (  Select decode(J.��¼״̬,2,0,J.ID) as ID,J.ҽ����� as ҽ��ID,J.�շ�ϸĿID, " & _
             "                       nvl(J.����,1)*nvl(J.����,1) as ����,  decode(J.��¼״̬,2,0,nvl(J.����,1)*nvl(J.����,1)) as  ԭʼ����" & _
             "              From  ������ü�¼ J,����ҽ����¼ M" & _
             "              where   J.ҽ�����=M.ID  " & _
             "                       And J.No=[1] and J.��¼����=[3] And J.��¼״̬ in (1,2,3) and J.�۸񸸺� is null   " & _
             "                       And Exists(Select 1 From   ����ҽ������ A Where   A.ҽ��ID=J.ҽ����� and  Nvl( A.ִ��״̬, 0) <> 1 And A.No||''=[1]  ) " & _
             "                       And Exists(Select 1 From   ����ҽ���Ƽ� A Where   A.ҽ��ID=J.ҽ����� and A.�շ�ϸĿID=J.�շ�ϸĿID And A.��������=0 And  Nvl( A.�շѷ�ʽ, 0) =0 ) " & _
             "                       And Instr('5,6,7', j.�շ����) = 0 And  Not Exists  (Select 1  From ��������  Where ����id = j.�շ�ϸĿid And Nvl(��������, 0) = 1)  " & _
             "                       And  instr(',C,D,F,G,K,',','||M.�������||',')=0  " & _
             "           Union all  " & _
             "           Select j.id, A.ҽ��ID,a.�շ�ϸĿID,-1*nvl(a.����,1)*nvl(C.��������,1) as ����,0 as ԭʼ���� " & _
             "           From ����ҽ���Ƽ� A,����ҽ������ B,����ҽ��ִ�� C,������ü�¼ J,����ҽ����¼ M " & _
             "           where  A.ҽ��ID=b.ҽ��id  and b.ҽ��id=c.ҽ��id and b.���ͺ�=c.���ͺ� And a.ҽ��id=M.ID " & _
             "                       And Nvl(C.ִ�н��, 1) =1 And A.��������=0 And  Nvl( A.�շѷ�ʽ, 0) =0  And Nvl(b.ִ��״̬, 0) <> 1 and  Nvl( B.ִ��״̬, 0) <> 1 And B.No||''=[1]  " & _
             "                       And a.ҽ��id=J.ҽ����� and a.�շ�ϸĿid=j.�շ�ϸĿid  " & _
             "                       And Instr('5,6,7', j.�շ����) = 0 And  Not Exists  (Select 1  From ��������  Where ����id = j.�շ�ϸĿid And Nvl(��������, 0) = 1)  " & _
             "                       And J.No=[1] and J.��¼����=[3] And J.��¼״̬ in (1,3) and J.�۸񸸺� is null   " & _
             "                       And  instr(',C,D,F,G,K,',','||M.�������||',')=0)  " & _
             "   Group by ҽ��ID,�շ�ϸĿID  Having Max(ID)<>0 )"
        End If
        strSQL1 = _
            " Select A.ID,A.���,A.�շ�ϸĿID," & _
            "       Nvl(A.����,1)*A.����" & IIf(gblnҩ����λ, "/Nvl(B." & gstrҩ����װ & ",1)", "") & " as ԭʼ����" & _
            " From ������ü�¼ A,ҩƷ��� B" & _
            " Where A.NO=[1] And A.��¼״̬ IN(0,1,3) And A.�۸񸸺� is null" & _
            "           And A.�շ�ϸĿID=B.ҩƷID(+) And A.��¼����=" & IIf(bytFun = 2, 2, 1) & _
                        IIf(mstrTime <> "", " And A.�Ǽ�ʱ��=[2]", "") & _
            "           And Nvl(A.���ӱ�־,0)<>9"
         
        
        '���ŵ��ݻ��ܽ��(��ϸ���շ�ϸĿ)
        'ִ��״̬Ӧ����ԭʼ��¼���ж�(������ҩ�Ҳ����˷ѵļ�¼)
        '������������ʱ"��¼״̬,���"�ظ�,AVG������,����Ҫ��"ִ��״̬"
        
        '��Ҫ�ſ�ҽ���ƻ��в�Ϊ������ȡ�ķ���:
        '   0-������ȡ��1-�����Թܷ��ã�2-һ�η���ֻ��ȡһ�Σ�3-����ֻ��ȡһ�Σ�4-����δִ����ȡһ�Σ�5-����ֻ��ȡһ�Σ��ų�������Ŀ��6-����δִ����ȡһ�Σ��ų�������Ŀ��7-ÿ���״β���ȡ
        
        strSQL = "" & _
            "  Select Nvl(�۸񸸺�,���) as ��� From ������ü�¼  " & _
            "  Where ��¼����=[3] And ��¼״̬ IN(0,1,3) And NO=[1] And Nvl(ִ��״̬,0)<>1     " & _
                        IIf(mstrTime <> "", " And �Ǽ�ʱ��=[2]", "") & _
            "            And Nvl(���ӱ�־,0)<>9 "
         If mblnHaveExcuteData = False Then
             '60735
           strSQL = strSQL & _
            "   Minus " & _
            "  Select Nvl(�۸񸸺�,���) as ��� " & _
            "  From ������ü�¼ A1,����ҽ���Ƽ� B1 " & _
            "  Where A1.ҽ�����=B1.ҽ��id And A1.�շ�ϸĿID=B1.�շ�ϸĿID And B1.��������=0 And Nvl( B1.�շѷ�ʽ, 0) <>0  " & _
            "           And A1.��¼����=[3] And A1.��¼״̬ IN(0,1,3) And A1.NO=[1] And Nvl(A1.ִ��״̬,0)=2 " & _
            "           And Instr('5,6,7', a1.�շ����) = 0 And  Not Exists  (Select 1  From ��������  Where ����id = a1.�շ�ϸĿid And Nvl(��������, 0) = 1)  " & _
            "           And Not Exists (Select 1 From ҩƷ�շ���¼ Where ����id =a1.Id) " & _
                        IIf(mstrTime <> "", " And A1.�Ǽ�ʱ��=[2]", "") & _
            "           And Nvl(A1.���ӱ�־,0)<>9 "
        End If
        '��Ϊ�ǽ�Ҫ��������ʣ�������ģ����Բ�����ֱ����ʱ�����ƣ����������
        strSQL = _
            " Select A.��¼״̬,A.ִ��״̬,Nvl(A.�۸񸸺�,A.���) as ���,A.�������� ," & _
            "       A.�ѱ�,C.����,C.���� as ���,A.�շ�ϸĿID,B.����,B.���,Nvl(A.��������,B.��������) ��������," & _
                    IIf(gblnҩ����λ, "Decode(X.ҩƷID,NULL,A.���㵥λ,X." & gstrҩ����λ & ")", "A.���㵥λ") & " as ���㵥λ,max(A.ҽ�����) as ҽ�����," & _
            "       Avg(Nvl(A.����,1)) as ����," & _
            "       Avg(A.����" & IIf(gblnҩ����λ, "/Nvl(X." & gstrҩ����װ & ",1)", "") & ") as ����," & _
            "       Sum(A.��׼����" & IIf(gblnҩ����λ, "*Nvl(X." & gstrҩ����װ & ",1)", "") & ") as ����," & _
            "       Sum(A.Ӧ�ս��) as Ӧ�ս��,Sum(A.ʵ�ս��) as ʵ�ս��, " & _
            "       A.ִ�в���ID,D.���� as ִ�в���,A.���ӱ�־,A.��ҩ����" & _
            " From ������ü�¼ A,�շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D,ҩƷ��� X" & _
            " Where A.�շ�ϸĿID=B.ID And C.����=A.�շ���� And A.ִ�в���ID=D.ID(+) " & _
            "       And A.�շ�ϸĿID=X.ҩƷID(+) And A.��¼����=[3]" & _
            "       And A.NO=[1] And Nvl(A.�۸񸸺�,A.���) IN(" & strSQL & ")" & _
            "       And Nvl(A.���ӱ�־,0)<>9" & _
            " Group by A.��¼״̬,A.ִ��״̬,Nvl(A.�۸񸸺�,A.���),A.��������,A.�ѱ�,C.����,C.����,A.�շ�ϸĿID,B.����," & _
            "       B.���,Nvl(A.��������,B.��������),A.���㵥λ,A.ִ�в���ID,D.����,A.���ӱ�־,A.��ҩ����,X.ҩƷID,X." & gstrҩ����λ
            
        '����������
        '��"׼������=ԭʼ����"ʱ,�����ű���
        '�ſ��Ѿ�ȫ���˷ѵ���,��ʣ������=0.(ִ��״̬=0��һ�ֿ���)
        '��ʣ��������׼�������������������
            '1.�޶�Ӧ���շ���¼(����ͨ���û򲻸������õ�����),��ʱӦ��ʣ������
            '2.�շ���¼����ȫ������,����ȫ��ִ��,SQL���ų����ּ�¼
        strSQL = strTableNo & vbCrLf & _
            " Select A.���,A.��������,A.�ѱ�,A.����,A.���,A.�շ�ϸĿID,A.����,A.���,A.��������,A.���㵥λ, " & _
            "       max(A.ҽ�����) as ҽ�����," & _
            "       Decode(Sign(Nvl(C.׼������,Sum(A.����*A.����))-B.ԭʼ����),0,Avg(A.����),1) as ׼�˸���," & _
            "       Decode(Sign(Nvl(C.׼������,Sum(A.����*A.����))-B.ԭʼ����),0,Sum(A.����),Nvl(C.׼������,Sum(A.����*A.����))) as ׼������," & _
            "        Nvl(C.׼������,Sum(A.����*A.����)) as ׼������,Sum(A.����*A.����) as ʣ������," & _
            "        A.����,Sum(A.Ӧ�ս��) as ʣ��Ӧ��,Sum(A.ʵ�ս��) as ʣ��ʵ��," & _
            "        A.ִ�в���ID,A.ִ�в���,A.���ӱ�־,A.��ҩ����" & _
            " From (" & strSQL & ") A,(" & strSQL1 & ") B,׼���� C" & _
            " Where A.���=B.��� And A.�շ�ϸĿID=B.�շ�ϸĿID+0 And B.ID=C.����ID(+)" & _
            " Group by A.���,A.��������,A.�ѱ�,A.����,A.���,A.�շ�ϸĿID,A.����,A.���,A.��������," & _
            "       A.���㵥λ,A.����,B.ԭʼ����,C.׼������,A.ִ�в���ID,A.ִ�в���,A.���ӱ�־,A.��ҩ����" & _
            " Having Sum(A.����*A.����)<>0"
            
        strSQL = _
            " Select A.���,A.��������,A.�ѱ�,A.����,A.���,A.�շ�ϸĿID,Nvl(B.����,A.����) as ����,E1.���� as ��Ʒ��," & _
            "           A.���,A.��������,A.���㵥λ,A.ҽ�����,A.׼�˸��� as ����,A.׼������ as ����,A.����," & _
            "           A.ʣ��Ӧ��*(A.׼������/A.ʣ������) as Ӧ�ս��," & _
            "           A.ʣ��ʵ��*(A.׼������/A.ʣ������) as ʵ�ս��," & _
            "           A.ִ�в���ID,A.ִ�в���,A.���ӱ�־,A.��ҩ����" & _
            " From (" & strSQL & ") A,�շ���Ŀ���� B,�շ���Ŀ���� E1 " & _
            " Where     A.�շ�ϸĿID=B.�շ�ϸĿID(+) And B.����(+)=1 And B.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
            "       And A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3 " & _
            " Order by A.���"
    ElseIf bytFun = 2 And mbytInState = 0 And mbytBilling = 2 Then
        '����ʱ,�������߱������
        '��ȡ���ʻ��۵�����(�������),ֻ��ȡδ��˲���
        strSQL = _
            " Select Nvl(A.�۸񸸺�,A.���) as ���,A.��������," & _
            "       A.�ѱ�,C.����,C.���� as ���,A.�շ�ϸĿID,B.����,B.���,Nvl(A.��������,B.��������) ��������," & _
                    IIf(gblnҩ����λ, "Decode(X.ҩƷID,NULL,A.���㵥λ,X." & gstrҩ����λ & ")", "A.���㵥λ") & " as ���㵥λ,max(A.ҽ�����) as ҽ�����," & _
            "       Avg(Nvl(A.����,1)) as ����," & _
            "       Avg(A.����" & IIf(gblnҩ����λ, "/Nvl(X." & gstrҩ����װ & ",1)", "") & ") as ����," & _
            "       Sum(A.��׼����" & IIf(gblnҩ����λ, "*Nvl(X." & gstrҩ����װ & ",1)", "") & ") as ����," & _
            "       Sum(A.Ӧ�ս��) as Ӧ�ս��,Sum(A.ʵ�ս��) as ʵ�ս��," & _
            "       A.ִ�в���ID,D.���� as ִ�в���,A.���ӱ�־,A.��ҩ����" & _
            " From ������ü�¼ A,�շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D,ҩƷ��� X" & _
            " Where A.��¼״̬=0 And A.�շ�ϸĿID=B.ID And C.����=A.�շ���� And A.ִ�в���ID=D.ID(+) " & _
                " And A.�շ�ϸĿID=X.ҩƷID(+) And A.NO=[1] And A.��¼����=2" & _
            " Group by Nvl(A.�۸񸸺�,A.���),A.��������,A.��¼״̬,A.�ѱ�,C.����,C.����,A.�շ�ϸĿID,B.����," & _
                " B.���,Nvl(A.��������,B.��������),A.���㵥λ,A.ִ�в���ID,D.����,A.���ӱ�־,A.��ҩ����,X.ҩƷID,X." & gstrҩ����λ
            
        strSQL = "Select" & _
            " A.���,A.��������,A.�ѱ�,A.����,A.���,A.�շ�ϸĿID,Nvl(B.����,A.����) as ����,E1.���� as ��Ʒ��,A.���,A.��������," & _
            " A.���㵥λ,A.ҽ�����,A.����,A.����,A.����,A.Ӧ�ս��,A.ʵ�ս��,A.ִ�в���ID,A.ִ�в���,A.���ӱ�־,A.��ҩ����" & _
            " From (" & strSQL & ") A,�շ���Ŀ���� B,�շ���Ŀ���� E1" & _
            " Where     A.�շ�ϸĿID=B.�շ�ϸĿID(+) And B.����(+)=1 And B.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
            "       And A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3 " & _
            " Order by A.���"
    Else
        '��ȡ����ԭʼ����
        intSign = IIf(mblnDelete, -1, 1) '����,�����������
        strSQL = _
            " Select Nvl(A.�۸񸸺�,A.���) as ���,A.��������," & _
            "       A.�ѱ�,C.����,C.���� as ���,A.�շ�ϸĿID,B.����,B.���,Nvl(A.��������,B.��������) ��������," & _
                    IIf(gblnҩ����λ, "Decode(X.ҩƷID,NULL,A.���㵥λ,X." & gstrҩ����λ & ")", "A.���㵥λ") & " as ���㵥λ,max(A.ҽ�����) as ҽ�����," & _
            "       Avg(Nvl(A.����,1)) as ����," & _
            "       Avg(" & intSign & "*A.����" & IIf(gblnҩ����λ, "/Nvl(X." & gstrҩ����װ & ",1)", "") & ") as ����," & _
            "       Sum(A.��׼����" & IIf(gblnҩ����λ, "*Nvl(X." & gstrҩ����װ & ",1)", "") & ") as ����," & _
            "       Sum(" & intSign & "*A.Ӧ�ս��) as Ӧ�ս��,Sum(" & intSign & "*A.ʵ�ս��) as ʵ�ս��," & _
            "       A.ִ�в���ID,D.���� as ִ�в���,A.���ӱ�־,A.��ҩ����" & _
            " From " & IIf(mblnNOMoved, zlGetFullFieldsTable("������ü�¼"), "������ü�¼  A") & ",�շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D,ҩƷ��� X" & _
            " Where A.�շ���� IN('5','6','7') And A.�շ�ϸĿID=B.ID And C.����=A.�շ���� And A.ִ�в���ID=D.ID(+) " & _
            "       And A.�շ�ϸĿID=X.ҩƷID And A.��¼����=" & IIf(bytFun = 2, 2, 1) & _
            "       And A.��¼״̬" & IIf(mblnDelete, "=2", " IN(0,1,3)") & " And A.NO=[1]" & _
                    IIf(mstrTime <> "", " And A.�Ǽ�ʱ��=[2]", "") & _
                    IIf(Not gblnShowErr, " And Nvl(A.���ӱ�־,0)<>9", "") & _
            " Group by Nvl(A.�۸񸸺�,A.���),A.��������,A.�ѱ�,C.����,C.����,A.�շ�ϸĿID,B.����," & _
            "   B.���,Nvl(A.��������,B.��������),A.���㵥λ,A.ִ�в���ID,D.����,A.���ӱ�־,A.��ҩ����,X.ҩƷID,X." & gstrҩ����λ
        
        strSQL = strSQL & " Union ALL " & _
            " Select Nvl(A.�۸񸸺�,A.���) as ���,A.��������," & _
            " A.�ѱ�,C.����,C.���� as ���,A.�շ�ϸĿID,B.����,B.���,Nvl(A.��������,B.��������) ��������," & _
            " A.���㵥λ,max(A.ҽ�����) as ҽ�����,Avg(Nvl(A.����,1)) as ����," & _
            " Avg(" & intSign & "*A.����) as ����,Sum(A.��׼����) as ����," & _
            " Sum(" & intSign & "*A.Ӧ�ս��) as Ӧ�ս��,Sum(" & intSign & "*A.ʵ�ս��) as ʵ�ս��," & _
            " A.ִ�в���ID,D.���� as ִ�в���,A.���ӱ�־,A.��ҩ����" & _
            " From " & IIf(mblnNOMoved, zlGetFullFieldsTable("������ü�¼"), "������ü�¼  A") & ",�շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D" & _
            " Where A.�շ���� Not IN('5','6','7') And A.�շ�ϸĿID=B.ID And C.����=A.�շ���� And A.ִ�в���ID=D.ID(+) " & _
            " And A.��¼����=" & IIf(bytFun = 2, 2, 1) & _
            " And A.��¼״̬" & IIf(mblnDelete, "=2", " IN(0,1,3)") & " And A.NO=[1]" & _
            IIf(mstrTime <> "", " And A.�Ǽ�ʱ��=[2]", "") & _
            IIf(Not gblnShowErr, " And Nvl(A.���ӱ�־,0)<>9", "") & _
            " Group by Nvl(A.�۸񸸺�,A.���),A.��������,A.�ѱ�,C.����,C.����,A.�շ�ϸĿID,B.����," & _
            " B.���,Nvl(A.��������,B.��������),A.���㵥λ,A.ִ�в���ID,D.����,A.���ӱ�־,A.��ҩ����"
            
        strSQL = "Select" & _
            " A.���,A.��������,A.�ѱ�,A.����,A.���,A.�շ�ϸĿID,Nvl(B.����,A.����) as ����,E1.���� as ��Ʒ��,A.���,A.��������," & _
            " A.���㵥λ,A.ҽ�����,A.����,A.����,A.����,A.Ӧ�ս��,A.ʵ�ս��,A.ִ�в���ID,A.ִ�в���,A.���ӱ�־,A.��ҩ����" & _
            " From (" & strSQL & ") A,�շ���Ŀ���� B,�շ���Ŀ���� E1" & _
            " Where A.�շ�ϸĿID=B.�շ�ϸĿID(+) And B.����(+)=1 And B.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
            "       And A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3 " & _
            " Order by A.���"
    End If
    If mstrTime <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo, CDate(mstrTime), IIf(bytFun = 2, 2, 1), IIf(bytFun = 2, 9, 8), IIf(bytFun = 2, 25, 24))
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo, "", IIf(bytFun = 2, 2, 1), IIf(bytFun = 2, 9, 8), IIf(bytFun = 2, 25, 24))
    End If
    If rsTmp.EOF Then Exit Function
    
 
    j = 0
    Bill.Redraw = False
    Call ClearBillRows
    Bill.Rows = rsTmp.RecordCount + 1
    curBillӦ�� = 0: curBillʵ�� = 0
    For i = 1 To rsTmp.RecordCount
        '�ѱ�
        If Not IsNull(rsTmp!�ѱ�) Then
            If InStr(str�ѱ� & ",", "," & rsTmp!�ѱ� & ",") = 0 Then
                str�ѱ� = str�ѱ� & "," & rsTmp!�ѱ�
            End If
        End If
        
        '�����շ�ʱ���´���ҩ����
        If mbytInFun = 0 And bytFun = 1 And InStr(",5,6,7,", rsTmp!����) > 0 Then
            j = j + 1
            'If j = 1 Then 'ֻ��δ���䷢ҩ����ʱ�����·���,�Ե�һҩƷ��Ϊ׼
                If IsNull(rsTmp!��ҩ����) Then
                    '���䴰��ʱ���������ҩ�������ŵ��ݲ�ͬ�������ȱʡ����,����ҩ����ͬ���䵽��ͬ����
                    If rsTmp!���� = "5" Then
                        If rsTmp!ִ�в���ID <> mlng��ҩ�� And mlng��ҩ�� <> 0 Then mstr���� = ""
                        mlng��ҩ�� = rsTmp!ִ�в���ID '��¼�ò���ʹ�õ�ҩ��(�����Ѷ�)
                    ElseIf rsTmp!���� = "6" Then
                        If rsTmp!ִ�в���ID <> mlng��ҩ�� And mlng��ҩ�� <> 0 Then mstr�ɴ� = ""
                        mlng��ҩ�� = rsTmp!ִ�в���ID
                    ElseIf rsTmp!���� = "7" Then
                        If rsTmp!ִ�в���ID <> mlng��ҩ�� And mlng��ҩ�� <> 0 Then mstr�д� = ""
                        mlng��ҩ�� = rsTmp!ִ�в���ID
                    End If
                    
                    '71902,Ƚ����,2014-04-09,ͬһ���˲��˲�ͬʱ��ζ��ŵ����շѣ�����ͬһ����ҩ���ڣ����㲡��ȡҩ
                    '�жϵ�ǰ�����Ƿ������ִͬ�в��ŵ�δ��ҩƷ���������򷵻�δ��ҩƷ�ķ�ҩ����
                    str��ҩ���� = Getδ��ҩƷ��ҩ����(lng����ID, rsTmp!ִ�в���ID)
        
                    '��ͬ����ҩƷ����ʹ����ͬ��ҩ��,���Ѱ���Է�����ͬ����
                    If str��ҩ���� = "" Then
                        str��ҩ���� = GetDrugWindow(rsTmp!ִ�в���ID, rsTmp!����, tbsBill.SelectedItem.Index)
                    End If
                    If str��ҩ���� = "" Then
                        str��ҩ���� = Get��ҩ����(zlDatabase.Currentdate, rsTmp!ִ�в���ID, rsTmp!����, mstr����, mstr�ɴ�, mstr�д�)
                    End If
                Else
                    str��ҩ���� = rsTmp!��ҩ����
                End If
                '����:47489
                If InStr(1, strPayDrugWins & ";", ";" & rsTmp!ִ�в���ID & "|") = 0 Then
                    strPayDrugWins = strPayDrugWins & ";" & rsTmp!ִ�в���ID & "|" & str��ҩ����
                End If
            'End If
        End If
        
        Bill.RowData(i) = rsTmp!��� '�۸񸸺�(���ڲ����˷ѻ�����)
        Bill.TextMatrix(i, BillCol.���) = rsTmp!���
        Bill.TextMatrix(i, BillCol.��������) = Nvl(rsTmp!��������)
        Bill.TextMatrix(i, BillCol.ҽ�����) = Nvl(rsTmp!ҽ�����) & "," & Nvl(rsTmp!�շ�ϸĿID)
        If Val(Nvl(rsTmp!ҽ�����)) <> 0 And InStr(strҽ����� & ",", "," & Val(Nvl(rsTmp!ҽ�����)) & ",") = 0 Then
            strҽ����� = strҽ����� & "," & Val(Nvl(rsTmp!ҽ�����))
        End If
        '����:29201
        strTemp = ""
        If Val(Nvl(rsTmp!��������)) <> 0 Then
            rsTmp.MoveNext
            strTemp = "��"
            If rsTmp.EOF Then
                strTemp = "��"
            ElseIf Bill.TextMatrix(i, BillCol.��������) <> Nvl(rsTmp!��������) Then
                strTemp = "��"
            End If
            rsTmp.MovePrevious
            strTemp = "  " & strTemp & " "
        End If
        Bill.TextMatrix(i, BillCol.��Ŀ) = strTemp & rsTmp!����
        Bill.TextMatrix(i, BillCol.��Ʒ��) = strTemp & Nvl(rsTmp!��Ʒ��)
        Bill.TextMatrix(i, BillCol.���) = Nvl(rsTmp!���)
        Bill.TextMatrix(i, BillCol.��λ) = Nvl(rsTmp!���㵥λ)
        Bill.TextMatrix(i, BillCol.����) = Nvl(rsTmp!����)
        Bill.TextMatrix(i, BillCol.����) = FormatEx(rsTmp!����, 5)
        Bill.TextMatrix(i, BillCol.����) = Format(rsTmp!����, gstrFeePrecisionFmt)
        Bill.TextMatrix(i, BillCol.Ӧ�ս��) = Format(rsTmp!Ӧ�ս��, gstrDec)
        Bill.TextMatrix(i, BillCol.ʵ�ս��) = Format(rsTmp!ʵ�ս��, gstrDec)
        Bill.TextMatrix(i, BillCol.ִ�п���) = Nvl(rsTmp!ִ�в���)
        Bill.TextMatrix(i, BillCol.��־) = IIf(rsTmp!���ӱ�־ = 1, "��", "")
        Bill.TextMatrix(i, BillCol.����) = Nvl(rsTmp!��������)
        Bill.TextMatrix(i, BillCol.ִ�п���ID) = Nvl(rsTmp!ִ�в���ID)
        
        curBillӦ�� = curBillӦ�� + rsTmp!Ӧ�ս��
        curBillʵ�� = curBillʵ�� + rsTmp!ʵ�ս��
        
        '�������ʱ�־
        If InStr("����,�˷�", Bill.TextMatrix(0, Bill.COLS - 1)) > 0 Then
            Bill.TextMatrix(i, Bill.COLS - 1) = "��"
        End If
        
        rsTmp.MoveNext
    Next
 
    If strҽ����� <> "" And Bill.TextMatrix(0, Bill.COLS - 1) = "�˷�" Then
        strҽ����� = Mid(strҽ�����, 2)
        Set mrs�շѶ��� = zlGet�����շѶ���(strҽ�����)
    Else
        Set mrs�շѶ��� = Nothing
    End If
    
    Set mrsDelInvoice = Nothing
    '25187
    Call LoadInvoiceData(strNo)
    Call ShowInvoiceInfor
    '��ʾ����С��
    lblSubӦ��.Caption = "Ӧ��:" & Format(curBillӦ��, gstrDec)
    lblSubʵ��.Caption = "ʵ��:" & Format(curBillʵ��, gstrDec)
    lblAmount.Caption = ""
    
    '��ʾ�ѱ�(����һ�ŵ����ж�̬�ѱ�����Ķ��ַѱ�)
    str�ѱ� = Mid(str�ѱ�, 2)
    i = UBound(Split(str�ѱ�, ","))
    lbl��̬�ѱ�.Visible = (i <> 0 And mbytInFun <> 2)
    cbo�ѱ�.Visible = Not (i <> 0 And mbytInFun <> 2)
    If i <> 0 And mbytInFun <> 2 Then
        lbl��̬�ѱ�.Caption = str�ѱ�
        lbl��̬�ѱ�.BorderStyle = 1
        lbl��̬�ѱ�.Left = cbo�ѱ�.Left
    Else
        cbo�ѱ�.ListIndex = cbo.FindIndex(cbo�ѱ�, str�ѱ�, True)
        If cbo�ѱ�.ListIndex = -1 Then
            cbo�ѱ�.AddItem str�ѱ�, 0
            cbo�ѱ�.ListIndex = cbo�ѱ�.NewIndex
        End If
        cbo�ѱ�.Locked = bytFun <> 0    '�շ��Ữ�۵�ʱ�������޸ķѱ�,��Ϊ���ò��ܱ�
        cboҽ�Ƹ���.Locked = bytFun <> 0 And gintPriceGradeStartType >= 2 '�շ��Ữ�۵�ʱ��ҽ�Ƹ��ʽ�����˼۸�ȼ��������޸ķѱ�,��Ϊ���ò��ܱ�
    End If
    cbo�ѱ�.TabStop = Not cbo�ѱ�.Locked And gbln�ѱ�
    
    '�շ���ʾ�˿�ϼ�
    If bytFun = 0 And blnDelete Then
        lblӦ��.Caption = "�˿�"
        lblӦ��.ForeColor = vbRed
        txtӦ��.ForeColor = vbRed
        
        mblnYB�������� = False
        MCPAR.ҽ���ӿڴ�ӡƱ�� = False
        mintInsure = ChargeExistInsure(strNo, , lng����ID)
        If mintInsure <> 0 Then
            mblnYB�������� = gclsInsure.GetCapability(support�����������, , mintInsure)
            MCPAR.ҽ���ӿڴ�ӡƱ�� = gclsInsure.GetCapability(supportҽ���ӿڴ�ӡƱ��, lng����ID, mintInsure, CStr(lng����ID))
            MCPAR.�˷Ѻ��ӡ�ص� = gclsInsure.GetCapability(support�˷Ѻ��ӡ�ص�, , mintInsure)
        End If
        Call ReCalce�˿�
        ReInitPatiInvoice (False)
    ElseIf bytFun = 0 And blnErrBill Then
        '�쳣���ݵĴ���
        If mintInsure = 0 Then
            mintInsure = ChargeExistInsure(strNo, , lng����ID)
        End If
    End If
    
    Call InitBillColumnColor
    Call SetColNum
    Bill.Redraw = True
    '��ȡ�����վݷ�Ŀ����
    If Not blnShow Then
        If blnDelete Then
            '�˷�ʱ�����Ǻ󱸱�,ǰ��Ĳ����ѽ�ֹ
            '��ȡ׼����,������Ӧ�ս��,ʵ�ս��(���=ʣ����*(׼����/ʣ����))
    
            '��ȡҩƷ�շ���¼�е�׼����
            strSQL1 = _
                " Select A.����ID,Sum(Nvl(A.����,1)*A.ʵ������" & IIf(gblnҩ����λ, "/Nvl(B." & gstrҩ����װ & ",1)", "") & ") as ׼������" & _
                " From ҩƷ�շ���¼ A,ҩƷ��� B" & _
                " Where A.NO=[1] And Mod(A.��¼״̬,3)=1 And A.����� is NULL" & _
                " And A.ҩƷID=B.ҩƷID(+) And A.���� IN(" & IIf(bytFun = 2, "9,25", "8,24") & ")" & _
                " Group by A.����ID"
            
            '���ŷ��õ���(��ϸ��������Ŀ)
            'ִ��״̬Ӧ����ԭʼ��¼���ж�(������ҩ�Ҳ����˷ѵļ�¼)
            strSQL = "Select Nvl(�۸񸸺�,���) From ������ü�¼" & _
                " Where ��¼����=" & IIf(bytFun = 2, 2, 1) & _
                " And ��¼״̬ IN(0,1,3) And NO=[1] And Nvl(ִ��״̬,0)<>1" & _
                IIf(mstrTime <> "", " And �Ǽ�ʱ��=[2]", "") & _
                " And Nvl(���ӱ�־,0)<>9"
            strSQL = _
                " Select Sum(A.ID) as ID,A.���,A.����,A.�շ����," & _
                    " Sum(A.����) as ʣ������,Sum(A.Ӧ�ս��) as ʣ��Ӧ��," & _
                    " Sum(A.ʵ�ս��) as ʣ��ʵ��" & _
                " From (" & _
                    " Select Decode(A.��¼״̬,2,0,A.ID) as ID,A.���," & _
                        IIf(gint����ϼ� = 0, "A.�վݷ�Ŀ", IIf(gint����ϼ� = 2, "'���ݺϼ�'", "B.����")) & " as ����,A.�շ����," & _
                        " Nvl(A.����,1)*A.����" & IIf(gblnҩ����λ, "/Nvl(X." & gstrҩ����װ & ",1)", "") & " as ����," & _
                        " A.Ӧ�ս��,A.ʵ�ս��" & _
                    " From ������ü�¼ A,������Ŀ B,ҩƷ��� X" & _
                    " Where A.��¼����=" & IIf(bytFun = 2, 2, 1) & _
                        " And A.�շ�ϸĿID=X.ҩƷID(+) And A.NO=[1] And A.������ĿID=B.ID" & _
                        " And Nvl(A.�۸񸸺�,A.���) IN(" & strSQL & ") And Nvl(A.���ӱ�־,0)<>9" & _
                        IIf(mstrTime <> "", " And A.�Ǽ�ʱ��=[2]", "") & _
                    " ) A" & _
                " Group by A.���,A.����,A.�շ����" & _
                " Having Sum(A.����)<>0"
                        
            '��������,����Ҳ��׼������:Nvl(B.׼������,A.ʣ������)
            '��ʣ��������׼�������������������
                '1.�޶�Ӧ���շ���¼(����ͨ���û򲻸������õ�����),��ʱӦ��ʣ������
                '2.�շ���¼����ȫ������,����ȫ��ִ��,SQL���ų����ּ�¼
            strSQL = _
                " Select A.����,Sum(A.ʣ��Ӧ��*(A.׼������/A.ʣ������)) as Ӧ�ս��," & _
                " Sum(ʣ��ʵ��*(A.׼������/A.ʣ������)) as ʵ�ս�� From (" & _
                " Select A.����,A.ʣ������,A.ʣ��Ӧ��,A.ʣ��ʵ��," & _
                " Decode(Instr(',4,5,6,7,',A.�շ����),0,A.ʣ������,Nvl(B.׼������,A.ʣ������)) as ׼������" & _
                " From (" & strSQL & ") A,(" & strSQL1 & ") B" & _
                " Where A.ID=B.����ID(+)" & _
                " ) A Group by A.����"
        ElseIf bytFun = 2 And mbytInState = 0 And mbytBilling = 2 Then
            '��ȡ���ʻ��۵�����(�������),ֻ��ȡδ��˲���
            strSQL = _
                "Select " & IIf(gint����ϼ� = 0, "A.�վݷ�Ŀ", IIf(gint����ϼ� = 2, "'���ݺϼ�'", "B.����")) & " as ����," & _
                " Sum(A.Ӧ�ս��) as Ӧ�ս��," & _
                " Sum(A.ʵ�ս��) as ʵ�ս�� " & _
                " From ������ü�¼ A,������Ŀ B" & _
                " Where A.��¼״̬=0 And A.��¼����=2" & _
                " And A.������ĿID=B.ID And A.NO=[1]" & _
                IIf(gint����ϼ� = 2, "", " Group By " & IIf(gint����ϼ� = 0, "A.�վݷ�Ŀ", "B.����"))
        Else
            '��ȡ����ԭʼ����
            intSign = IIf(mblnDelete, -1, 1) '����,�����������
            strSQL = _
                "Select " & IIf(gint����ϼ� = 0, "A.�վݷ�Ŀ", IIf(gint����ϼ� = 2, "'���ݺϼ�'", "B.����")) & " as ����," & _
                " Sum(" & intSign & "*A.Ӧ�ս��) as Ӧ�ս��," & _
                " Sum(" & intSign & "*A.ʵ�ս��) as ʵ�ս�� " & _
                " From " & IIf(mblnNOMoved, zlGetFullFieldsTable("������ü�¼"), "������ü�¼  A") & " ,������Ŀ B" & _
                " Where A.��¼״̬" & IIf(mblnDelete, "=2", " IN(0,1,3)") & _
                " And A.��¼����=" & IIf(bytFun = 2, 2, 1) & _
                IIf(mstrTime <> "", " And A.�Ǽ�ʱ��=[2]", "") & _
                " And A.NO=[1] And A.������ĿID=B.ID" & _
                IIf(Not gblnShowErr, " And Nvl(A.���ӱ�־,0)<>9 ", "") & _
                IIf(gint����ϼ� = 2, "", " Group By " & IIf(gint����ϼ� = 0, "A.�վݷ�Ŀ", "B.����"))
        End If
        If mstrTime <> "" Then
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo, CDate(mstrTime))
        Else
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo)
        End If
        If rsTmp.EOF Then Exit Function
        
        'ˢ����ʾ(�շ�Ҫ����)
        mshMoney.Rows = rsTmp.RecordCount + 1 + mintMoneyRow
        If mshMoney.Rows < M_MONEY_ROWS Then mshMoney.Rows = M_MONEY_ROWS
        Call SetMoneyList
        For i = mintMoneyRow + 1 To mshMoney.Rows - 1
            mshMoney.TextMatrix(i, 0) = ""
            mshMoney.TextMatrix(i, 1) = ""
            mshMoney.TextMatrix(i, 2) = ""
        Next
        curBillӦ�� = 0: curBillʵ�� = 0
        For i = mintMoneyRow + 1 To rsTmp.RecordCount + mintMoneyRow
            mshMoney.TextMatrix(i, 0) = mintBillNO + 1
            mshMoney.TextMatrix(i, 1) = rsTmp!����
            mshMoney.TextMatrix(i, 2) = Format(rsTmp!ʵ�ս��, gstrDec)
            curBillӦ�� = curBillӦ�� + rsTmp!Ӧ�ս��
            curBillʵ�� = curBillʵ�� + rsTmp!ʵ�ս��
            rsTmp.MoveNext
        Next
        On Error Resume Next
        For i = 1 To mshMoney.Rows - 1
            If mshMoney.TextMatrix(i, 0) = mintBillNO + 1 Then
                mshMoney.TopRow = i
            End If
        Next
        On Error GoTo errH
        
        '���൥����ʾ�ϼ�
        txtӦ��.Text = Format(mcurBillӦ�� + curBillӦ��, gstrDec)
        txt�ϼ�.Text = Format(mcurBillʵ�� + curBillʵ��, gstrDec)
        
        '����ʱ������ʾӦ��,���ֱҴ����Ľ��
        If mbytInFun = 1 Then txt�ۼ�.Text = Format(CentMoney(txt�ϼ�.Text), "0.00")
        
        '���ڼ��ʵ�����ʾ�ϼ�
        lblTotal.Caption = "�ϼ�:" & Format(curBillʵ��, gstrDec)
        
        'ˢ���շ��ۼ�
        If chkCancel.Value = 0 And mbytInFun = 0 And gbln�ۼ� And Not mblnDoing Then
            txt�ۼ�.Text = Format(GetChargeTotal, "0.00")
            txt�ۼ�.ToolTipText = "��ǰ����Ա�����շ��ۼƶ�"
        End If
        
        '�൥���շ�֧��:�����ڸ��ֵ���
        With mobjBill.Pages(tbsBill.SelectedItem.Index)
            .NO = strNo
            .Ӧ�ս�� = curBillӦ��
            .ʵ�ս�� = curBillʵ��
            
            '���շ�ʱ��ȡ���۵���
            If mbytInFun = 0 And bytFun = 1 Then
                '47489
                If strPayDrugWins <> "" Then strPayDrugWins = Mid(strPayDrugWins, 2)
                tbsBill.SelectedItem.Tag = strPayDrugWins ' str��ҩ����
                Call ShowMoney(mintPage) 'ֻ��Ҫ���㵱ǰ����
            End If
        End With
    End If
    If mbytInFun = 0 And (mbytInState = 0 Or mbytInState = 3 Or mbytInState = 2) Then
        '�˷�
        With mTyDelFee
            .strNos = strNo
        End With
    End If
    ReadBill = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function ReadBalance(ByVal lng����ID As Long, _
    ByVal blnDel As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ؽ����㷽ʽ
    '���:lng����ID- ����ID
    '       blnDel-�Ƿ��˷�
    '����:
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-30 08:16:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim k As Long, i As Long, j As Long, dbl��Ԥ�� As Double, str���㷽ʽ As String
    Dim blnFind As Boolean, lngRow As Long, intSign As Integer
    On Error GoTo errHandle
     '�˷�
    With mTyDelFee
        Set .rsBlance = GetChargeBalance("", 0, lng����ID, mblnNOMoved)
    End With
    '��ȡ����ԭʼ����ʱ,��ʾ���ֽ��
    '(����)�˷�ʱ,��ʾԭʼ���ݵĽ�����
    intSign = IIf(mblnDelete, -1, 1) '����,�����������
    
    '����:�շ���صĽ��㷽ʽ(����:1-Ԥ���,2-ҽ��,3-ҽ�ƿ�,4-���㿨,5-һ��ͨ,0-������)
    '       �ֶ�:A.����ID,A.NO,A.����,A.��������,A.���㷽ʽ,A.������,
    '               A.�����ID,A.����,A.�Ƿ�ȫ��,A.�Ƿ�����,A.�������,A.����,A.������ˮ��,
    '               A.����˵��,A.�������,A.У�Ա�־
    With mTyDelFee.rsBlance
        .Filter = 0: k = 0: j = 0
        Do While Not .EOF
            '1-�ֽ���㷽ʽ,2-������ҽ������,3-ҽ�������ʻ�,4-ҽ������ͳ��,5-���տ���,6-�����ۿ�,7-һ��ͨ����,8-���㿨����
            ',7,8:59673
            If Val(Nvl(!����)) <> 1 And InStr(",1,2,", Val(Nvl(!��������))) > 0 Then
                k = k + 1
            End If
            If InStr(str���㷽ʽ & ",", "," & Nvl(!���㷽ʽ) & ",") = 0 Then
                str���㷽ʽ = str���㷽ʽ & "," & Nvl(!���㷽ʽ)
            End If
            .MoveNext
        Loop
        If str���㷽ʽ <> "" Then str���㷽ʽ = Mid(str���㷽ʽ, 2)
        mTyDelFee.blnSingleBalance = InStr(str���㷽ʽ, ",") = 0
        str���㷽ʽ = ""
        
        If .RecordCount <> 0 Then .MoveFirst
        dbl��Ԥ�� = 0
        vsBalance.Rows = 1
         mblnOlnyԤ�� = True
        Do While Not .EOF
            If Val(Nvl(!����)) = 1 Then
                dbl��Ԥ�� = dbl��Ԥ�� + intSign * Val(Nvl(!������))
            Else    '�鿴�շѵ�ʱȫ����ʾ���б���
                mblnOlnyԤ�� = False
                str���㷽ʽ = Nvl(!���㷽ʽ, " ")
                ''1-�ֽ���㷽ʽ,2-������ҽ������,3-ҽ�������ʻ�,4-ҽ������ͳ��,5-���տ���,6-�����ۿ�,7-һ��ͨ����,8-���㿨����
                
                If InStr("1,2,3,4,8,7", !��������) > 0 Or k > 1 Or Not blnDel Then
                    'ҽ�����㷽ʽ,���ŵ��ݶ��ֽ��㷽ʽ
                    With vsBalance
                        blnFind = False: lngRow = 0
                        For i = 0 To .Rows - 1
                            If .TextMatrix(i, 0) = str���㷽ʽ Then
                                lngRow = i
                                blnFind = True: Exit For
                            End If
                        Next
                        If .TextMatrix(lngRow, 0) = "" And lngRow = 0 Then blnFind = True
                        If Not blnFind Then
                              .Rows = .Rows + 1: lngRow = .Rows - 1
                        End If
                    End With
                    vsBalance.TextMatrix(lngRow, 0) = str���㷽ʽ
                    vsBalance.Cell(flexcpData, lngRow, 0) = Val(Nvl(!����))
                    vsBalance.TextMatrix(lngRow, 1) = Val(vsBalance.TextMatrix(lngRow, 1)) + intSign * Val(Nvl(!������))
                    vsBalance.Cell(flexcpData, lngRow, 1) = vsBalance.TextMatrix(lngRow, 1)
                    
                    If blnDel Then
                     '   vsBalance.Cell(flexcpForeColor, lngRow, 0, lngRow, vsBalance.COLS - 1) = vbRed
                    End If
                    If blnDel Then
                        '�˷�ʱ����ҽ���ķ�ʽ�������ע,�Լ����˷ѽ��
                        '1-Ԥ���,2-ҽ��,3-ҽ�ƿ�,4-���㿨,5-һ��ͨ,0-������
                        Select Case Val(Nvl(!����))
                        Case 3, 4 '3-ҽ�ƿ�,4-���㿨,
                            If Val(Nvl(!�Ƿ�����)) = 1 Then vsBalance.RowData(lngRow) = -1
                            mTyDelFee.bln������ȫ�� = Val(Nvl(!�Ƿ�ȫ��)) = 1
                        Case Else
                            If InStr(",1,2,", !��������) > 0 Then vsBalance.RowData(lngRow) = -1
                        End Select
                    End If
                End If
                If InStr(",1,2,", !��������) > 0 And k = 1 Then
                    '��ҽ�����㷽ʽ
                    mblnNotClick = True
                    zlControl.CboSetText cbo���㷽ʽ, str���㷽ʽ
                    mblnNotClick = False
                    If cbo���㷽ʽ.ListIndex = -1 Then
                        cbo���㷽ʽ.AddItem str���㷽ʽ
                        zlControl.CboSetIndex cbo���㷽ʽ.hWnd, cbo���㷽ʽ.NewIndex
                    End If
                End If
            End If
            .MoveNext
        Loop
    End With
    If mblnOlnyԤ�� = True And dbl��Ԥ�� = 0 Then mblnOlnyԤ�� = False
    txtԤ�����.Text = Format(intSign * dbl��Ԥ��, "0.00")
    If dbl��Ԥ�� <> 0 Then
        If vsBalance.Rows < 4 Then vsBalance.Rows = 4
        txtԤ�����.Visible = True: lblԤ�����.Visible = True
    Else
        txtԤ�����.Visible = False: lblԤ�����.Visible = False
        If vsBalance.Rows < 6 Then vsBalance.Rows = 6
    End If
    
    With vsBalance
        For i = 0 To .Rows - 1
            If Val(.TextMatrix(i, 1)) <> 0 Then
                .TextMatrix(i, 1) = Trim(Format(Val(.TextMatrix(i, 1)), "##0.00###"))
            End If
        Next
    End With
    vsBalance.ToolTipText = "���㷽ʽ�б�"
    mintReturnMode = cbo���㷽ʽ.ListIndex  '�����˷�ʱ,ȫ�˽��ý��㷽ʽʱ�ָ���ʼ�Ľ��㷽ʽ
    Call picAppend_Resize
    ReadBalance = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Sub SetShowCol()
'���ܣ��Զ�ȷ���Ƿ����ظ�����
    mrsClass.Filter = "����='7'"
    If mrsClass.RecordCount = 0 Then
        Bill.ColWidth(BillCol.����) = 0
    ElseIf Bill.ColWidth(BillCol.����) = 0 Then
        Bill.ColWidth(BillCol.����) = 520 'ǿ����ʾ
    End If
End Sub

Private Sub DelFactMoney()
'���ܣ�ɾ�������еĹ�������(������Ҫ������ʱ)
    Dim i As Long, p As Integer
    
    '���ж��Ƿ��Ѿ������˹�����
    For p = 1 To mobjBill.Pages.Count
        For i = 1 To mobjBill.Pages(p).Details.Count
            If mobjBill.Pages(p).Details(i).������ Then
                Call DeleteDetail(i, p)
                
                '�����к�(��ǰ����)
                If mintPage = p Then
                    Call SetColNum(i)
                End If
                
                'ֻ�й�����ʱ��ͬʱɾ������
                If mobjBill.Pages(p).Details.Count = 0 And fraBill.Visible Then
                    If tbsBill.Tabs.Count > 1 Then Call DelOneBill(p)
                End If
                
                Call ShowMoney(p)
                
                If CheckBillsEmpty Then ClearMoney
                Exit Sub
            End If
        Next
    Next
End Sub

Private Sub SetFactMoney()
'���ܣ��շ�ʱ���á���ʾ�����㹤����
'˵�����������Զ����ڵ�ǰ��ʾ�ĵ�����
    Dim objDetail As Detail
    Dim colIncomes As New BillInComes
    Dim lngDoUnit As Long, blnExist As Boolean
    Dim intPage As Integer, lngRow As Long
    Dim i As Integer, p As Integer
    Dim int���� As Integer, blnReCalc As Boolean
    
    int���� = GetInvoiceCount '��ӡ����(������������)
    If int���� = 0 Then Call DelFactMoney: Exit Sub 'ɾ��������
    
    '���ж��Ƿ��Ѿ������˹�����
    For p = 1 To mobjBill.Pages.Count
        For i = 1 To mobjBill.Pages(p).Details.Count
            If mobjBill.Pages(p).Details(i).������ Then
                intPage = p: lngRow = i '���ڵ���
                blnExist = True: Exit For
            End If
        Next
        If blnExist Then Exit For
    Next
    
    '����������ӹ�����
    If Not blnExist Then
        blnReCalc = True
        Set objDetail = Get������
        If objDetail Is Nothing Then Exit Sub '�Ҳ���������,������
        
        'Ѱ�ҿ�����ӹ����ѵĵ���
        For p = 1 To mobjBill.Pages.Count
            If mobjBill.Pages(p).NO = "" Then
                intPage = p: lngRow = mobjBill.Pages(p).Details.Count + 1
                Exit For
            End If
        Next
        If intPage = 0 Then
            '�޿��Ա༭����,����һ�ŵ���
            If Not cmdAddBill.Enabled Or Not cmdAddBill.Visible Then Exit Sub '��֧�ֶ൥��
            Call AddNewBill
            intPage = mobjBill.Pages.Count: lngRow = 1
        ElseIf intPage = mintPage Then
            '�ǵ�ǰ����,�������
            If mobjBill.Pages(intPage).Details.Count >= Bill.Rows - 1 Then
                Bill.Rows = Bill.Rows + 1
            Else
                For i = 1 To Bill.COLS - 1
                    Bill.TextMatrix(Bill.Rows - 1, i) = ""
                Next
            End If
        End If
        
        With objDetail
            lngDoUnit = mobjBill.����ID
            If lngDoUnit = 0 Then lngDoUnit = mobjBill.Pages(intPage).��������ID
            lngDoUnit = Get�շ�ִ�п���ID(.���, .ID, .ִ�п���, lngDoUnit, Get��������ID, gint������Դ, , , , , mobjBill.����ID)
            mobjBill.Pages(intPage).Details.Add "", objDetail, .ID, CInt(lngRow), 0, .���, .���㵥λ, "", 1, 1, 0, lngDoUnit, colIncomes
        End With
        mobjBill.Pages(intPage).Details(lngRow).������ = True
    Else
        '�������������δ��,��������
        If mobjBill.Pages(intPage).Details(lngRow).���� <> int���� Then blnReCalc = True
    End If
    
    If blnReCalc Then
        '���¸��ݵ�ǰ�����������ù���������
        mobjBill.Pages(intPage).Details(lngRow).���� = int����
        Call CalcMoney(intPage, lngRow)
        
        If mintPage = intPage Then
            Call ShowDetails(lngRow)
        End If
        Call ShowMoney(intPage)
    End If
End Sub

Private Sub ClearBillRows()
'���ܣ�������ݱ����ʾ����
    Dim i As Integer
    For i = 1 To Bill.Rows - 1
        Bill.RowData(i) = 0
        Call SetBillRowForeColor(i, Bill.ForeColor)
    Next
    Bill.ClearBill
    Call SetColNum
    
    lblSubӦ��.Caption = "Ӧ��:" & gstrDec
    lblSubʵ��.Caption = "ʵ��:" & gstrDec
    lblAmount.Caption = ""
End Sub

Private Function GetOtherCTMGroups(lngRow As Long) As Integer
'���ܣ�ȡ��ǰ������������ҩ�ĸ���
    Dim i As Integer
    
    GetOtherCTMGroups = 1
    For i = 1 To mobjBill.Pages(mintPage).Details.Count
        If mobjBill.Pages(mintPage).Details(i).�շ���� = "7" And i <> lngRow Then
            GetOtherCTMGroups = mobjBill.Pages(mintPage).Details(i).����
            Exit For
        End If
    Next
End Function
Private Function Get��������ID() As Long
    Dim lng������ID As Long
    Dim rs������ As ADODB.Recordset
    
    If gbyt����ҽ�� = 2 Then
        If cbo������.ListIndex <> -1 Then
            lng������ID = cbo������.ItemData(cbo������.ListIndex)
            Set rs������ = mrs������ '����Ӱ���ⲿ���õļ�¼��
            
            rs������.Filter = "ȱʡ=1 And ID=" & lng������ID
            If rs������.RecordCount = 0 Then rs������.Filter = "ID=" & lng������ID
            If rs������.RecordCount > 0 Then Get��������ID = rs������!����ID
        End If
    End If
    
    If Get��������ID = 0 Then
        If cbo��������.ListIndex <> -1 Then
            Get��������ID = cbo��������.ItemData(cbo��������.ListIndex)
        Else
            Get��������ID = UserInfo.����ID
        End If
    End If
End Function
Private Function GetWorkUnit(ByVal lngҩƷID As Long, ByVal str��� As String) As Boolean
'���ܣ�ȡ���пɹ�ѡ���ҩ��
    Dim strSQL As String, bytDay As Byte
    Dim strҩ�� As String, lng��������ID As Long
    
    lng��������ID = mobjBill.����ID     '������������
    If lng��������ID = 0 And cbo��������.ListIndex <> -1 Then lng��������ID = cbo��������.ItemData(cbo��������.ListIndex)
    
    If str��� = "4" Then
        strSQL = _
        "Select Distinct c.Id, c.����, c.����, c.����, b.��������, b.�������" & vbNewLine & _
        "From �շ�ִ�п��� A, ��������˵�� B, ���ű� C" & vbNewLine & _
        "Where a.ִ�п���id + 0 = b.����id And b.�������� = '���ϲ���' And b.������� IN(" & gint������Դ & ",3) And b.����id = c.Id And" & vbNewLine & _
        "      (c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.����ʱ�� Is Null) And (c.վ�� = '" & gstrNodeNo & "' Or c.վ�� Is Null) And" & vbNewLine & _
        "      (a.������Դ Is Null Or A.������Դ=" & gint������Դ & ") And" & vbNewLine & _
        "      (a.��������id Is Null Or a.��������id = [1] Or Exists (Select 1 From �������Ҷ�Ӧ Where ����id = [1] And a.��������id = ����id)) And a.�շ�ϸĿid = [2]" & vbNewLine & _
        "Order By b.�������, c.����"
    Else
        '��ҩƷ����ȷ��ҩ������
        Select Case str���
            Case "5"
                strҩ�� = "��ҩ��"
            Case "6"
                strҩ�� = "��ҩ��"
            Case "7"
                strҩ�� = "��ҩ��"
        End Select
        
        'ҩƷ��ϵͳָ���Ĵ���ҩ������
        If Not gblnҩ���ϰల�� Then
            strSQL = _
            " Select Distinct C.ID,C.����,C.����,C.����,B.��������,B.������� " & _
            " From �շ�ִ�п��� A,��������˵�� B,���ű� C" & _
            " Where A.ִ�п���ID+0=B.����ID And B.��������='" & strҩ�� & "'" & _
            "       And B.������� IN(" & gint������Դ & ",3) And B.����ID=C.ID" & _
            "       And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
            "       And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null) " & vbNewLine & _
            "       And (A.������Դ is NULL Or A.������Դ=" & gint������Դ & ")" & _
            "       And (A.��������ID is NULL Or A.��������ID=[1])" & _
            "       And A.�շ�ϸĿID=[2]" & _
            " Order by B.�������,C.����"
        Else
            bytDay = Weekday(zlDatabase.Currentdate, vbMonday) Mod 7 '0=����,1=��һ
            strSQL = _
            " Select Distinct C.ID,C.����,C.����,C.����,B.��������,B.������� " & _
            " From �շ�ִ�п��� A,��������˵�� B,���ű� C,���Ű��� D" & _
            " Where A.ִ�п���ID+0=B.����ID And B.��������='" & strҩ�� & "'" & _
            "       And B.������� IN(" & gint������Դ & ",3) And B.����ID=C.ID" & _
            "       And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
            "       And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null) " & vbNewLine & _
            "       And D.����ID=C.ID And D.����=" & bytDay & _
            "       And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.��ʼʱ��,'HH24:MI:SS') and To_Char(D.��ֹʱ��,'HH24:MI:SS') " & _
            "       And (A.������Դ is NULL Or A.������Դ=" & gint������Դ & ")" & _
            "       And (A.��������ID is NULL Or A.��������ID=[1])" & _
            "       And A.�շ�ϸĿID=[2]" & _
            " Order by B.�������,C.����"
        End If
    End If
    
    On Error GoTo errH
    Set mrsWork = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��������ID, lngҩƷID)
    GetWorkUnit = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetPrePaySum(Optional intPage As Integer) As Currency
    Dim curTotal As Currency, i As Long
    
    For i = 1 To mobjBill.Pages.Count
        If intPage = 0 Or i = intPage Then
            curTotal = curTotal + mobjBill.Pages(i).��Ԥ����
        End If
    Next
    GetPrePaySum = curTotal
End Function
Public Function GetBillSum(Optional blnӦ�� As Boolean, Optional ByVal intPage As Integer) As Currency
'���ܣ���ȡ���ݺϼƽ��
'������intPage=ָ������,����Ϊ���е���
    Dim objTmpDetail As New BillDetail
    Dim objTmpIncome As New BillInCome
    Dim curTotal As Currency, intCol As Integer
    Dim i As Integer, j As Integer, k As Integer
    
    For i = 1 To mobjBill.Pages.Count
        If intPage = 0 Or i = intPage Then
            If mobjBill.Pages(i).Details.Count > 0 Then
                For j = 1 To mobjBill.Pages(i).Details.Count
                    For k = 1 To mobjBill.Pages(i).Details(j).InComes.Count
                        If blnӦ�� Then
                            curTotal = curTotal + mobjBill.Pages(i).Details(j).InComes(k).Ӧ�ս��
                        Else
                            curTotal = curTotal + mobjBill.Pages(i).Details(j).InComes(k).ʵ�ս��
                        End If
                    Next
                Next
            Else    '��ȡ���۵��շ�ʱû����ϸ����
                If blnӦ�� Then
                    curTotal = curTotal + mobjBill.Pages(i).Ӧ�ս��
                Else
                    curTotal = curTotal + mobjBill.Pages(i).ʵ�ս��
                End If
            End If
        End If
    Next
    
    '���û��,�ٳ��Դӱ����ȡ(��һ�ŵ���ʱ)
    If curTotal = 0 And tbsBill.Tabs.Count = 1 And Bill.Rows > 1 Then
        If Not (Bill.Rows = 2 And Bill.TextMatrix(1, BillCol.��Ŀ) = "") Then
            intCol = IIf(blnӦ��, BillCol.Ӧ�ս��, BillCol.ʵ�ս��)
            For i = 1 To Bill.Rows - 1
                If IsNumeric(Bill.TextMatrix(i, intCol)) Then
                    curTotal = curTotal + Format(Val(Bill.TextMatrix(i, intCol)), gstrDec)
                End If
            Next
        End If
    End If
    GetBillSum = Format(curTotal, gstrDec)
End Function

Private Function Calc������(Optional ByVal intPage As Integer) As Currency
Dim i As Integer, j As Integer, k As Integer

    For i = 1 To mobjBill.Pages.Count
        If intPage = 0 Or i = intPage Then
            For j = 1 To mobjBill.Pages(i).Details.Count
                If mobjBill.Pages(i).Details(j).������ Then
                    For k = 1 To mobjBill.Pages(i).Details(j).InComes.Count
                        Calc������ = Calc������ + mobjBill.Pages(i).Details(j).InComes(k).ʵ�ս��
                    Next
                End If
            Next
        End If
    Next
End Function

Private Sub txtModi_GotFocus()
    Call zlControl.TxtSelAll(txtModi)
End Sub

Private Sub txtModi_KeyPress(KeyAscii As Integer)
    '�շѺͻ��۹��ܿ����ڴ��ڵ��뵥���޸�
    Dim strNo As String
    If KeyAscii > 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))

    '��һλ����������ĸ,����λ����
    If KeyAscii <> 13 Then
        Call SetNOInputLimit(txtModi, KeyAscii)
    Else
        KeyAscii = 0
        
        '��ȡ�޸ĵ���
        txtModi.Text = GetFullNO(txtModi.Text, 13)
        Call zlControl.TxtSelAll(txtModi)
        strNo = txtModi.Text
        
        Call ClearFullBill(False)
        
        mstrInNO = strNo
        Call LoadModifyNO(strNo, IIf(mbytInFun = 2, 2, 1))
    End If
End Sub

Private Function ExistOneCardSwap(strNo As String) As Boolean
'���ܣ����ָ���ĵ����Ƿ����һ��ͨ����
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select 1" & vbNewLine & _
            " From ������ü�¼ a, ����Ԥ����¼ b, ���㷽ʽ c" & vbNewLine & _
            " Where a.����id = b.����id And b.���㷽ʽ = c.���� And a.��¼���� = 1 And a.No = [1] And c.����=7" & vbNewLine & _
            "       And Rownum < 2"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo)
    ExistOneCardSwap = rsTmp.RecordCount > 0
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Function


Private Sub LoadModifyNO(strNo As String, bytFlag As Byte)
'����:����,�շ�,�����޸ĵ��ݼ������ݼ���ؼ��
'     ���ۺ��շѳ������嵥��������޸���,�������ڴ��ڽ������뵥�ݺ����޸�
'����:strNO-��ǰ�޸ĵĵ��ݺ�
'     bytFlag-��¼���ʣ�1-�շѻ򻮼ۣ�2-����
    Dim lng����ID As Long, lng����ID As Long, bln���� As Boolean
    Dim strMessage As String, strNos As String
    Dim strTmp As String, i As Long
                
    'a.������
    '-------------------------------------------------------------------------------------
    '�Ƿ���ת������ݱ���
    If mbytInFun = 0 Or mbytInFun = 2 And mblnCopyBill = False Then
        If zlDatabase.NOMoved("������ü�¼", strNo, , bytFlag, Me.Caption) Then
            If Not ReturnMovedExes(strNo, bytFlag, Me.Caption) Then GoTo ExitHandle
        End If
        
        '�Ѿ��˹���(����)�����ʵĵ��ݲ������޸�
        If BillExistDelete(strNo, bytFlag) Then
            strMessage = "�õ��ݰ�����" & IIf(mbytInFun = 2, "����", "�˷�") & "����,�������޸ġ�": GoTo ExitHandle
        End If
    End If
    
    If mblnCopyBill = False Then
        '����������ʱ��ҩƷ�ĵ��ݲ������޸�
        If Not BillCanModi(strNo, bytFlag) Then
            strMessage = "���ŵ��ݰ���������ʱ��ҩƷ,���ܿ���ѷ����仯,�������޸ġ�": GoTo ExitHandle
        End If
        
        'Ҫ��黮�۵�,��Ϊϵͳ������������δ�շѵĴ�����ҩ
        '�����������ִ�л�ȫ��ִ�е���Ŀ,���˷Ѻ������Ҫ��ӡƱ��,�������޸�
        '                                   ����Ǽ���,��һ������ȫ������,�������޸�
        If HaveExecute(1, strNo, bytFlag) Then
            strMessage = "�õ��ݰ�����ȫִ�л򲿷�ִ�е���Ŀ,�������޸ġ�": GoTo ExitHandle
        End If
    End If
    
    '�ڶ�ȡ������Ϣǰ�ȶ�
    If mbytInFun = 2 And mblnCopyBill = False Then
        Original.ʵ�պϼ� = GetBillMoney(1, strNo, , IIf(mbytInFun = 2, 2, 1))
    ElseIf mbytInFun = 0 Then
        Call GetBillPay(strNo, Original.��Ԥ����, Original.Ӧ�ɽ��)
    End If
    
    '�շѹ�����ؼ��
    If mbytInFun = 0 Then
        'δ�շѵĻ��۵��������޸�
        If Billδ�շ�(strNo, 1) Then
            strMessage = "�û��۵�����δ�շ�,�������޸ġ�": GoTo ExitHandle
        End If
        If gblnMultiBalance Or gTy_Module_Para.bln������ Then strNos = GetMultiNOs(strNo, , , True)
        If gblnMultiBalance And InStr(1, strNos, ",") > 0 Then
            If CheckSingleBalance(strNos) = False Then
                strMessage = "���ŵ���ʹ�ö��ֽ��㷽ʽģʽ�²������޸Ķ��ŵ��ݡ�": GoTo ExitHandle
            End If
        End If
        If gTy_Module_Para.bln������ Then
            If InStr(1, strNos, ",") > 0 Then
                If BillExistFact(strNo) Then
                    strMessage = "��õ���һ���շѵĶ��ŵ����а��������ѣ����ܽ����޸ġ�": GoTo ExitHandle
                End If
            End If
        End If
        If mbytInFun = 0 Then
            '������ҽ��������㣬�������޸�
            If CheckBillExistReplenishData(1, , Replace(strNos, "'", "")) = True Then
                MsgBox "��ǰ���ݽ�����ҽ��������㣬�������޸ģ�", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        mintInsure = ChargeExistInsure(strNo, lng����ID, lng����ID, bln����)
        If mintInsure > 0 Then
            If InStr(mstrPrivs, "�����շ�") = 0 Then
                strMessage = "��û��Ȩ�޶�ҽ�����˵ĵ��ݽ����޸Ĳ�����": GoTo ExitHandle
            End If
        
            '�޸�ҽ����������֤���
            Call txtPatient_KeyPress(13)
            If mstrYBPati = "" Then GoTo ExitHandle
            
            '��֤�Ĳ�����ݱ������
            If lng����ID <> mobjBill.����ID Then
                strMessage = "��֤�Ĳ�������뵥���еĲ�����ݲ����������޸ĵ��ݡ�"
                mstrYBPati = "": mintInsure = 0: GoTo ExitHandle
            End If
            
            '�ж�ÿһ�ֽ��㷽ʽ�Ƿ�֧���˷�,ֻҪ����һ�ֲ�֧��,�������޸�
            If Not Check�����������(lng����ID, mintInsure) Then
                mstrYBPati = "": mintInsure = 0: GoTo ExitHandle
            End If
            
            Original.����ID = lng����ID '��¼�޸�ʱҪ�˷ѵĵ��ݵĽ���ID
            If bln���� Then chk����.Value = 1
            
            '��Ϊ���޸�,�������ʻ�������
            mcur������� = mcur������� + Read�����ʻ�����(lng����ID)
            sta.Panels(Pan.C3�����ʻ�).Text = "�����ʻ����:" & Format(mcur�������, "0.00")
            
        Else
            '�Ƿ��з�ҽ�����˵��˷�Ȩ��
            If InStr(mstrPrivs, "�����ҽ������") = 0 Then
                strMessage = "��û��Ȩ�޶Է�ҽ�����˵ĵ��ݽ����޸Ĳ�����": GoTo ExitHandle
            End If
           
            If mblnOneCard Then
                If ExistOneCardSwap(strNo) Then
                    strMessage = "�õ��ݲ�����һ��ͨ����,�������޸ġ�": GoTo ExitHandle
                End If
            End If
           
        End If
        
    ElseIf mbytInFun = 2 And mblnCopyBill = False Then
        'δȫ����˻�����˵Ĳ������޸�
        If Not BillIdentical(strNo) Then
            strMessage = "�����а������ݲ�ȫ����˻�ֶ����˵����ݣ��������޸ġ�": GoTo ExitHandle
        End If
                    
        '����Ѿ�����,���ݲ��������Ƿ������޸�
        If HaveBilling(1, strNo) Then
            Select Case gbytBillOpt
                Case 0
                Case 1
                    If MsgBox("�ü��ʵ��Ѿ�����,Ҫ�޸���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then GoTo ExitHandle
                Case 2
                    strMessage = "�ü��ʵ��Ѿ�����,�������޸ġ�": GoTo ExitHandle
            End Select
        End If
    End If
    
    
    'b.��ȡ������ϸ�����ݶ���
    '---------------------------------------------------------------------------------------------------
    Set mobjBill = ImportBill(strNo, IIf(mblnCopyBill, False, True), mbytInFun, mintInsure, , , _
        mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ�۸�ȼ�)
    If mobjBill.NO = "" Then
        strMessage = "��ȡ����ʧ�ܡ�": GoTo ExitHandle
    End If
    
    If mblnCopyBill = False Then
        If mbytInFun = 0 And InStr(mstrPrivs, "���в���Ա") <= 0 Then
            If UserInfo.���� <> mobjBill.����Ա���� Then
                strMessage = "��û��""���в���Ա""Ȩ��,�����޸�" & mobjBill.����Ա���� & "�ĵ��ݡ�": GoTo ExitHandle
            End If
        End If
        
        If Not BillOperCheck(Choose(mbytInFun + 1, 2, 3, 4), mobjBill.����Ա����, mobjBill.�Ǽ�ʱ��, "�޸�", strNo, , IIf(mbytInFun = 2, 2, 1)) Then
            GoTo ExitHandle
        End If
        
        'ҽ�����ɵĻ��۵�,�����շѵ�ҽ�����ɵĻ��۵�,�������޸�
        If mobjBill.Pages(1).ҽ����� <> 0 Then strMessage = "��ҽ�������ĵ��ݲ������޸ģ�": GoTo ExitHandle
    End If
    
    'c.��ʾ��Ϣ
    '------------------------------------------------------------------------------------------------------
    mbln������۸� = True
        Call Set�����˿�������(mobjBill.Pages(mintPage).������, mobjBill.Pages(mintPage).��������ID)
        Call LoadAndSeek�ѱ�        '���طѱ�Ͷ�̬�ѱ�
        
        'a.�ѽ����Ĳ���
        If mobjBill.����ID <> 0 Then
            If mstrYBPati = "" Then 'ҽ��������ǰ������֤���
                txtPatient.Text = "-" & mobjBill.����ID
                Call txtPatient_KeyPress(13)
            End If
        Else
        'b.���ۻ��շ�ʱ��δ�����Ĳ���
            txtPatient.Text = mobjBill.����
            cboSex.ListIndex = cbo.FindIndex(cboSex, mobjBill.�Ա�, True)
            Call LoadOldData(mobjBill.����, txt����, cbo���䵥λ)
            txt�����.Text = IIf(mobjBill.��ʶ�� = 0, "", mobjBill.��ʶ��)
            cboҽ�Ƹ���.ListIndex = GetCboIndexByCode(cboҽ�Ƹ���, mobjBill.����)
            
            strTmp = GetBill�ѱ�(mobjBill)
            If strTmp <> "" Then cbo�ѱ�.ListIndex = cbo.FindIndex(cbo�ѱ�, strTmp, True)
        End If
        If cbo�ѱ�.ListIndex = -1 And cbo�ѱ�.ListCount > 0 Then cbo�ѱ�.ListIndex = 0
    mbln������۸� = False
    
    
    'ȡ��һҩƷ��
    For i = 1 To mobjBill.Pages(1).Details.Count
        If InStr(",5,6,7,", mobjBill.Pages(1).Details(i).�շ����) > 0 Then
            mlngFirstID = mobjBill.Pages(1).Details(i).ִ�в���ID
            mstrFirstWin = mobjBill.Pages(1).Details(i).��ҩ����
            Exit For
        End If
    Next
    
        
    If mblnCopyBill Then
        txtDate.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
        mstrInNO = "" '�������ݺ����,�������޸�
    Else
        '��ʾ����ԭ���ݺ�,��������µ��ݺ�
        cboNO.Text = strNo
        txtDate.Text = Format(mobjBill.����ʱ��, "yyyy-MM-dd HH:mm:ss")
    End If
    mblnDo = False: chk�Ӱ�.Value = mobjBill.�Ӱ��־: mblnDo = True
    Bill.Rows = mobjBill.Pages(1).Details.Count + 1
    
    If mbytInFun = 2 Then
        Call zlControl.CboSetIndex(cboBaby.hWnd, mobjBill.Ӥ����)
        If cbo��������.ListIndex <> -1 Then cboBaby.Enabled = is����(cbo��������.ItemData(cbo��������.ListIndex), mrs��������)
    End If
    
    '�޸�ʱӦ���浱ǰ����Ա������
    mobjBill.����Ա��� = UserInfo.���: mobjBill.����Ա���� = UserInfo.����
    
    'ȱʡΪԭ���ݵĽ��㷽ʽ
    If mbytInFun = 0 Then
        strTmp = GetBalanceName(strNo)
        If strTmp <> "" Then
            i = cbo.FindIndex(cbo���㷽ʽ, strTmp, True)
            If i <> -1 Then cbo���㷽ʽ.ListIndex = i
        End If
    End If
    
    '�²���
    If IIf(mlngPrePati = 0, mstrPrePati = mobjBill.����, mlngPrePati = mobjBill.����ID) Then
        mcurBillʵ�� = 0:  mcurBillӦ�� = 0: mcurBillӦ�� = 0
        mintBillNO = 0: mintMoneyRow = 0
    End If
    
''    '�����������ʹ��Ԥ����,�򱣳�ԭ�ȵĳ���
''    If Not gblnPrePayPriority And Original.��Ԥ���� > 0 And txtԤ�����.Enabled Then
''       txtԤ�����.Text = Format(IIf(Original.��Ԥ���� > Val(sta.Panels(Pan.C4Ԥ����Ϣ).Tag), Val(sta.Panels(Pan.C4Ԥ����Ϣ).Tag), Original.��Ԥ����), "0.00")
''    End If
''
    If gintPriceGradeStartType < 2 Then
        If gbln��������ۿ� Then Call CalcMoneys
    Else
        Call CalcMoneys
    End If
    Call ShowDetails
    Call ShowMoney
              
                
    'd.�������
    '------------------------------------------------------------------------------------------
'    txt����Ӧ��.Visible = True: lblӦ��.Caption = "�ٽ�"
    
    Call InitBillColumnColor
    Call SetColNum
    Call SetPatientEnableModi(mobjBill.����ID = 0)
           
    chkCancel.Enabled = False: cmdDelete.Enabled = False
    cmdAddBill.Enabled = False
    
    If Me.Visible Then
        Bill.Active = True: Bill.SetFocus
        If mintInsure > 0 Then txtModi.Enabled = False  'ҽ�����ݱ��밴"ȡ��"��ť������ȡ�������֤
    End If
    
    Exit Sub
ExitHandle:
    If strMessage <> "" Then
        
        MsgBox strMessage, vbInformation, gstrSysName
    End If
    If Me.Visible Then
        Set mobjBill = New ExpenseBill
        txtModi.Text = ""
        If txtModi.Visible And txtModi.Enabled Then txtModi.SetFocus
    Else
        Unload Me
    End If
End Sub
'''
'''
'''Private Sub txt�ɿ�_LostFocus()
'''    mblnHotKey = False
'''    txt�ɿ�.Text = Format(Val(txt�ɿ�.Text), "0.00")
'''End Sub

Private Sub SetPatientEnableModi(blnModi As Boolean)
    
    txtPatient.Locked = Not blnModi
    
    If blnModi Then
        txtPatient.BackColor = &HFFFFFF
    Else
        txtPatient.BackColor = &HE0E0E0
    End If

    cboSex.Locked = txtPatient.Locked
    txt����.Locked = txtPatient.Locked
    txt����.BackColor = txtPatient.BackColor
    cbo���䵥λ.Locked = txtPatient.Locked
End Sub

Private Sub SetInputItem()
    '������Ŀ
    If Not gbln�Ա� Then cboSex.TabStop = False
    If Not gbln���� Then txt����.TabStop = False: cbo���䵥λ.TabStop = False
    If Not gbln�ѱ� Then cbo�ѱ�.TabStop = False
    If Not gblnҽ�Ƹ��� Then cboҽ�Ƹ���.TabStop = False
    If Not gbln�Ӱ� Then chk�Ӱ�.TabStop = False
    If Not gbln�������� Then txtDate.TabStop = False
    If Not gbln������ Then cbo������.TabStop = False
End Sub

Private Function SaveModi() As Boolean
    '���ܣ����浱ǰ�޸ĵķ��õ���
    Dim strSQL As String
    If Not IsDate(txtDate.Text) Then
        MsgBox "������Ϸ��ķ���ʱ�䣡", vbInformation, gstrSysName
        If txtDate.Enabled And txtDate.Visible Then txtDate.SetFocus
        Exit Function
    End If
    strSQL = "zl_���˷��ü�¼_Update('" & cboNO.Text & "'," & IIf(mbytInFun = 2, 2, 1) & "," & _
        "'" & zlStr.NeedName(cbo������.Text) & "',To_Date('" & txtDate.Text & "','YYYY-MM-DD HH24:MI:SS'))"
    On Error GoTo errH
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    SaveModi = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ShowDeleteCol(blnShow As Boolean)
'���ܣ���ʾ\�������ʱ�־��
    Dim i As Integer, blnACT As Boolean
    If blnShow Then
        If InStr("����,�˷�", Bill.TextMatrix(0, Bill.COLS - 1)) = 0 Then
            Bill.Redraw = False
            Bill.COLS = Bill.COLS + 1
            If mbytInFun = 2 Then
                Bill.TextMatrix(0, Bill.COLS - 1) = "����"
            Else
                Bill.TextMatrix(0, Bill.COLS - 1) = "�˷�"
            End If
            Bill.ColAlignment(Bill.COLS - 1) = 4
            Bill.ColWidth(Bill.COLS - 1) = 550
            Bill.ColData(Bill.COLS - 1) = BillColType.CheckBox
            
            blnACT = Bill.Active: Bill.Active = False
            Bill.Row = 0: Bill.Col = Bill.COLS - 1: Bill.MsfObj.CellForeColor = vbRed
            Bill.Row = 1: Bill.Col = Bill.COLS - 1
            Bill.Active = blnACT
            
            Bill.ColWidth(BillCol.���) = GetOrigColWidth(BillCol.���) - 100
            Bill.ColWidth(BillCol.��Ŀ) = GetOrigColWidth(BillCol.��Ŀ) - 100
            Bill.ColWidth(BillCol.ִ�п���) = GetOrigColWidth(BillCol.ִ�п���) - 200
            
            Bill.ColWidth(BillCol.����) = GetOrigColWidth(BillCol.����) - 50
            Bill.ColWidth(BillCol.Ӧ�ս��) = GetOrigColWidth(BillCol.Ӧ�ս��) - 50
            Bill.ColWidth(BillCol.ʵ�ս��) = GetOrigColWidth(BillCol.ʵ�ս��) - 50
            Bill.Redraw = True
        End If
    Else
        If InStr("����,�˷�", Bill.TextMatrix(0, Bill.COLS - 1)) > 0 Then
            Bill.Redraw = False
            Bill.COLS = Bill.COLS - 1
            Bill.ColWidth(BillCol.���) = GetOrigColWidth(BillCol.���)
            Bill.ColWidth(BillCol.��Ŀ) = GetOrigColWidth(BillCol.��Ŀ)
            Bill.ColWidth(BillCol.ִ�п���) = GetOrigColWidth(BillCol.ִ�п���)
            
            Bill.ColWidth(BillCol.����) = GetOrigColWidth(BillCol.����)
            Bill.ColWidth(BillCol.Ӧ�ս��) = GetOrigColWidth(BillCol.Ӧ�ս��)
            Bill.ColWidth(BillCol.ʵ�ս��) = GetOrigColWidth(BillCol.ʵ�ս��)
            Bill.Redraw = True
        End If
    End If
End Sub

Private Function GetOrigColWidth(ByVal intIdx As Integer) As Long
'���ܣ���ȡָ���е�ԭʼ�п�
    GetOrigColWidth = Val(Split(Split(STR_HEAD, ";")(intIdx), ",")(1))
End Function

Private Sub SetColNum(Optional intRow As Long = 1)
'���ܣ�������ʾ���е��к�
'������intRow=�Ӹ��п�ʼ
    Dim bln As Boolean, i As Integer
    
    Bill.Redraw = False
    For i = intRow To Bill.Rows - 1
        Bill.TextMatrix(i, BillCol.��) = i
    Next
    Bill.Redraw = True
End Sub

Private Function CheckDuty(Optional tmpDetail As Detail, Optional blnCommon As Boolean = True, Optional intPage As Long) As Integer
'���ܣ����ָ��ҩƷ�е�ְ���Ƿ��뵱ǰҽ����ְ����ƥ��
'������tmpDetail=�����������Ŀ,����Ϊ���е���������,blnCommon=�Ƿ��������ж�,����Ϊҽ���򹫷Ѳ��˵��ж�
'���أ���ƥ�����,0Ϊ��ȷ,intPage=����ҳ��
'˵����ְ��1=����,2=����,3=�м�,4=����/ʦ��,5=Ա/ʿ,9=��Ƹ
    Dim i As Integer, p As Integer, strTmp As String
    Dim intְ��A As Integer, intְ��B As Integer
    Dim strMsg As String
    
    strTmp = "����,����,�м�,����/ʦ��,Ա/ʿ,,,,��Ƹ"
    intPage = 0
    
    If tmpDetail Is Nothing Then
        For p = 1 To mobjBill.Pages.Count
            If mobjBill.Pages(p).������ <> "" Then
                'ÿ�ŵ��ݿ����˲�ͬ,��ǰ���ݵĿ�������ְ��
                Call GetOperatorInfo(mobjBill.Pages(p).������, , intְ��A)
                
                For i = 1 To mobjBill.Pages(p).Details.Count
                    If InStr(",5,6,7,", mobjBill.Pages(p).Details(i).�շ����) > 0 Then
                        If mobjBill.Pages.Count > 1 Then strMsg = "�ڵ��� " & p & "��"
                        If Not blnCommon Then
                            intְ��B = Val(Right(mobjBill.Pages(p).Details(i).Detail.����ְ��, 1))
                            If intְ��B > 0 Then
                                If intְ��A = 0 Then
                                    strMsg = "��ҽ���򹫷�" & gstrCustomerAppellation & "," & strMsg & _
                                        "�� " & p & " ҳ " & i & " ��ҩƷ""" & mobjBill.Pages(p).Details(i).Detail.���� & _
                                        """Ҫ�󿪵���ְ������Ϊ""" & Split(strTmp, ",")(intְ��B - 1) & """," & _
                                        "��""" & mobjBill.Pages(p).������ & """δ����ְ��"
                                    CheckDuty = 1: intPage = p
                                ElseIf intְ��B < intְ��A Then
                                    strMsg = "��ҽ���򹫷�" & gstrCustomerAppellation & "," & strMsg & _
                                        "�� " & p & " ҳ " & i & " ��ҩƷ""" & mobjBill.Pages(p).Details(i).Detail.���� & _
                                        """Ҫ�󿪵���ְ��Ϊ""" & Split(strTmp, ",")(intְ��B - 1) & """����," & _
                                        "��""" & mobjBill.Pages(p).������ & """ְ��Ϊ""" & Split(strTmp, ",")(intְ��A - 1) & """��"
                                    CheckDuty = i: intPage = p: Exit For
                                End If
                            End If
                        Else
                            intְ��B = Val(Left(mobjBill.Pages(p).Details(i).Detail.����ְ��, 1))
                            If intְ��B > 0 Then
                                If intְ��A = 0 Then
                                    strMsg = strMsg & "�� " & p & " ҳ " & i & " ��ҩƷ""" & mobjBill.Pages(p).Details(i).Detail.���� & _
                                        """Ҫ�󿪵���ְ������Ϊ""" & Split(strTmp, ",")(intְ��B - 1) & """," & _
                                        "��""" & mobjBill.Pages(p).������ & """δ����ְ��"
                                    CheckDuty = 1: intPage = p
                                ElseIf intְ��B < intְ��A Then
                                    strMsg = strMsg & "�� " & p & " ҳ " & i & " ��ҩƷ""" & mobjBill.Pages(p).Details(i).Detail.���� & _
                                        """Ҫ�󿪵���ְ��Ϊ""" & Split(strTmp, ",")(intְ��B - 1) & """����," & _
                                        "��""" & mobjBill.Pages(p).������ & """ְ��Ϊ""" & Split(strTmp, ",")(intְ��A - 1) & """��"
                                    CheckDuty = i: intPage = p: Exit For
                                End If
                            End If
                        End If
                    End If
                Next
            End If
        Next
    ElseIf mobjBill.Pages(mintPage).������ <> "" Then
        If InStr(",5,6,7,", tmpDetail.���) = 0 Then Exit Function
        Call GetOperatorInfo(mobjBill.Pages(mintPage).������, , intְ��A)
        
        If Not blnCommon Then
            intְ��B = Val(Right(tmpDetail.����ְ��, 1))
            If intְ��B > 0 Then
                If intְ��A = 0 Then
                    strMsg = "��ҽ���򹫷�" & gstrCustomerAppellation & ",ҩƷ""" & tmpDetail.���� & _
                        """Ҫ�󿪵���ְ������Ϊ""" & Split(strTmp, ",")(intְ��B - 1) & """," & _
                        "����ǰ������δ����ְ��"
                    CheckDuty = 1
                ElseIf intְ��B < intְ��A Then
                    strMsg = "��ҽ���򹫷�" & gstrCustomerAppellation & ",ҩƷ""" & tmpDetail.���� & _
                        """Ҫ�󿪵���ְ��Ϊ""" & Split(strTmp, ",")(intְ��B - 1) & """����," & _
                        "����ǰ������ְ��Ϊ""" & Split(strTmp, ",")(intְ��A - 1) & """��"
                    CheckDuty = 1
                End If
            End If
        Else
            intְ��B = Val(Left(tmpDetail.����ְ��, 1))
            If intְ��B > 0 Then
                If intְ��A = 0 Then
                    strMsg = "ҩƷ""" & tmpDetail.���� & """Ҫ�󿪵���ְ������Ϊ""" & Split(strTmp, ",")(intְ��B - 1) & """," & _
                        "����ǰ������δ����ְ��"
                    CheckDuty = 1
                ElseIf intְ��B < intְ��A Then
                    strMsg = "ҩƷ""" & tmpDetail.���� & """Ҫ�󿪵���ְ��Ϊ""" & Split(strTmp, ",")(intְ��B - 1) & """����," & _
                        "����ǰ������ְ��Ϊ""" & Split(strTmp, ",")(intְ��A - 1) & """��"
                    CheckDuty = 1
                End If
            End If
        End If
    End If
    
    If CheckDuty > 0 Then MsgBox strMsg, vbInformation, gstrSysName
End Function

Private Function CheckInhibitiveByNurse(ByVal intPage As Integer) As Boolean
'���ܣ��ж�ָ���������Ƿ��л�ʿ��ֹ���������
    Dim rsTmp As New ADODB.Recordset
    Dim bln��ʿ As Boolean, strSQL As String
    Dim i As Integer
    
    CheckInhibitiveByNurse = False
    If mobjBill.Pages(intPage).������ <> "" Then
        Call GetOperatorInfo(mobjBill.Pages(intPage).������, bln��ʿ)
        If Not bln��ʿ Then Exit Function
        
        If mobjBill.Pages(intPage).NO = "" Then
            For i = 1 To mobjBill.Pages(intPage).Details.Count
                If InStr(",E,M,4,", mobjBill.Pages(intPage).Details(i).�շ����) = 0 Then
                    CheckInhibitiveByNurse = True: Exit Function
                End If
            Next
'            '���۵����ټ��
        End If
    End If
End Function

Private Sub FillDoctor(Optional lng����ID As Long)
'���ܣ�����ָ���Ŀ�������ID��ȡ����дҽ���б�,����ȱʡҽ��
    Dim strOldID As String
    Dim bln������Ա���� As Boolean, str�������� As String
    
    cbo������.Clear
    If mbytInFun = 1 And mbytInState = 0 Then '113577
        bln������Ա���� = zlStr.IsHavePrivs(mstrPrivs, "���п���") = False And gblnUserIsClinic
    End If
    If mbytInFun = 1 Then
        str�������� = "'�ٴ�','����','����','���','����','����'"
        Call GetDoctor(lng����ID, mrs������, bln������Ա����, str��������)
    Else
        Call GetDoctor(lng����ID, mrs������, bln������Ա����)
    End If
    
    Do While Not mrs������.EOF
    '70857:������,2014-03-07,�����˼���һ��ʱ���ڼ����ظ�������
        If InStr("," & strOldID & ",", "," & mrs������!ID & ",") = 0 Then
            If gbyt��������ʾ = 1 Then
                cbo������.AddItem mrs������!���� & "-" & mrs������!����
            Else
                cbo������.AddItem mrs������!��� & "-" & mrs������!����
            End If
            cbo������.ItemData(cbo������.NewIndex) = mrs������!ID
            strOldID = strOldID & mrs������!ID & ","
        End If
        mrs������.MoveNext
    Loop
End Sub



Private Sub FillDept(ByVal lngDeptID As Long, Optional lng��ԱID As Long)
'���ܣ���ȡ�����ؿ����б�,����ȱʡ����
'������
'   lngDeptID-��ǰ�����Ĳ���
'   lng��ԱID=ֻ��ȡָ����Ա���ڿ���(������ȱʡ��)
'���أ����Ҹ���
    
    Dim strSQL As String, i As Long, lngOldDepID As Long
    Dim strDepts As String  'ָ����Ա�����Ķ������
    Dim bln������Ա���� As Boolean, str���� As String
        
    cbo��������.Clear
    If mbytInFun = 1 And mbytInState = 0 Then '113577
        bln������Ա���� = zlStr.IsHavePrivs(mstrPrivs, "���п���") = False And gblnUserIsClinic
    End If
    If mrs�������� Is Nothing Then
        If mbytInFun = 1 Then
            str���� = "'�ٴ�','����','����','���','����','����'"
        Else
            str���� = "'�ٴ�','����'"
        End If
        Call GetDoctorDept(mrs��������, bln������Ա����, str����, IIf(zlStr.IsHavePrivs(mstrPrivs, "���п���"), 0, lngDeptID))
    End If
   
    If lng��ԱID <> 0 Then
        If Not mrs������ Is Nothing Then
            mrs������.Filter = "ID=" & lng��ԱID
            For i = 1 To mrs������.RecordCount
                strDepts = strDepts & " OR ID=" & mrs������!����ID      'filter��֧��in
                mrs������.MoveNext
            Next
        End If
        If strDepts <> "" Then
            mrs��������.Filter = Mid(strDepts, 4)
        Else
            mrs��������.Filter = "ID=0" '��Աû�����ò���,����ʾ��������
        End If
    Else
        mrs��������.Filter = ""
    End If
    
    If mrs��������.RecordCount > 0 Then
        For i = 1 To mrs��������.RecordCount
            If lngOldDepID <> mrs��������!ID Then   'һ�����ſ���ͬʱ���ڲ��ƺ��ٴ�,��������ͬ��
                cbo��������.AddItem IIf(zlIsShowDeptCode, mrs��������!���� & "-", "") & mrs��������!����     '������:27658
                cbo��������.ItemData(cbo��������.NewIndex) = mrs��������!ID
                lngOldDepID = mrs��������!ID
            End If
            mrs��������.MoveNext
        Next
    End If
End Sub

Private Function CheckDrugExist(objDetail As Detail) As Boolean
'���ܣ��ж�ָ��ҩƷ������(��������)�ڵ������Ƿ��Ѿ�����
'������objDetail=��Ŀ,intRow=Ҫ�жϵ���
'˵����ʱ�ۻ������ͬһִ�п��ҽ�ֹ�ظ�����(�������ʾ,����ʱ��ֹ)
'      ��ʱ�۵ķ���ҩƷ���ڲ�ͬ�ĵ���������ͬ�ģ������ϲ���������
    Dim i As Integer, p As Integer
    Dim strTmp As String
    
    For p = 1 To mobjBill.Pages.Count
        For i = 1 To mobjBill.Pages(p).Details.Count
            If Not (p = mintPage And i = Bill.Row) And InStr(",4,5,6,7,", mobjBill.Pages(p).Details(i).�շ����) > 0 Then
                If mobjBill.Pages(p).Details(i).Detail.ID = objDetail.ID Then
                    If mobjBill.Pages.Count > 1 Then strTmp = "�� " & p & " �ŵ���"
                    If (mobjBill.Pages(p).Details(i).Detail.���� Or mobjBill.Pages(p).Details(i).Detail.���) _
                        And (objDetail.���� Or objDetail.���) Then
                        
                        '��ʱ�۵ķ���ҩƷ���ڲ�ͬ�ĵ���������ͬ�ģ������ϲ���������
                        If objDetail.��� Or (Not objDetail.��� And objDetail.���� And mintPage = p) Then
                            If objDetail.��� = "4" Then
                                If MsgBox("��������""" & objDetail.���� & """��" & strTmp & "�� " & i & " ���Ѿ�����,Ҫ������" & _
                                    vbCrLf & vbCrLf & "ע�⣺����������Ϊ������ʱ�۲���,�ظ�����ʱ���뱣֤���ǵķ��ϲ��Ų�ͬ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                    CheckDrugExist = True
                                End If
                            Else
                                If MsgBox("ҩƷ""" & objDetail.���� & """��" & strTmp & "�� " & i & " ���Ѿ�����,Ҫ������" & _
                                    vbCrLf & vbCrLf & "ע�⣺��ҩƷΪ������ʱ��ҩƷ,�ظ�����ʱ���뱣֤���ǵ�ִ��ҩ����ͬ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                    CheckDrugExist = True
                                End If
                            End If
                            Exit Function
                        End If
                    Else
                        If objDetail.��� = "4" Then
                            If MsgBox("��������""" & objDetail.���� & """��" & strTmp & "�� " & i & " ���Ѿ�����,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                CheckDrugExist = True
                            End If
                        Else
                            If MsgBox("ҩƷ""" & objDetail.���� & """��" & strTmp & "�� " & i & " ���Ѿ�����,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                CheckDrugExist = True
                            End If
                        End If
                        Exit Function
                    End If
                End If
            End If
        Next
    Next
End Function

Private Function CheckFeeType(Optional intRow As Integer) As Boolean
'���ܣ����ݵ�ǰ���˵������ж�ָ���е���Ŀ�Ƿ��������,����������������Ŀ
    Dim strSQL As String, strType As String
    Dim i As Integer, p As Integer
    Dim strTmp As String, blnҽ�� As Boolean, bln���� As Boolean
    
    On Error GoTo errHandle
    
    CheckFeeType = True
    
    '�޷����
    If cboҽ�Ƹ���.ListIndex = -1 Then Exit Function
    'ҽ���򹫷Ѳ���
    '����:45605
    If zlIsCheckMedicinePayMode(zlStr.NeedName(cboҽ�Ƹ���), blnҽ��, bln����) = False Then Exit Function
    'ֻ���ҽ�����˺͹��Ѳ���
    strType = IIf(blnҽ��, 1, 2)
    
    '��ȡ�������
    If mrs�������� Is Nothing Then
        strSQL = " Select 'ҽ��' As ���,����,���� From �������� Where ���� In(" & gstrҽ���������� & ") Union All " & _
                 " Select '����' As ���,����,���� From �������� Where ���� In(" & gstr���ѷ������� & ") "
        Set mrs�������� = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(mrs��������, strSQL, Me.Caption)
    End If
    mrs��������.Filter = ""
    If mrs��������.RecordCount = 0 Then Exit Function
        
    If strType = "1" Then
        strSQL = " And ���='ҽ��'"
    Else
        strSQL = " And ���='����'"
    End If
    
    
    If intRow > 0 Then
        If mobjBill.Pages(mintPage).Details(intRow).Detail.���� = "" Then
            MsgBox """" & mobjBill.Pages(mintPage).Details(intRow).Detail.���� & """�ķ�������δ���ã�", vbInformation, gstrSysName
            CheckFeeType = False
        Else
            mrs��������.Filter = "����='" & mobjBill.Pages(mintPage).Details(intRow).Detail.���� & "'" & strSQL
            If mrs��������.EOF Then
                MsgBox """" & mobjBill.Pages(mintPage).Details(intRow).Detail.���� & """�ķ�������Ϊ""" & _
                    mobjBill.Pages(mintPage).Details(intRow).Detail.���� & """,����" & _
                    IIf(strType = "1", "ҽ��", "����") & "�������ͣ�", vbInformation, gstrSysName
                CheckFeeType = False
            End If
        End If
    Else
        For p = 1 To mobjBill.Pages.Count
            For i = 1 To mobjBill.Pages(p).Details.Count
                If mobjBill.Pages.Count > 1 Then strTmp = "�� " & p & " ��"
                If mobjBill.Pages(p).Details(i).Detail.���� = "" Then
                    If MsgBox(strTmp & "�����е� " & i & " ����Ŀ""" & mobjBill.Pages(p).Details(i).Detail.���� & """�ķ�������δ���ã�" & vbCrLf & "ȷʵҪ���浥����", _
                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        CheckFeeType = False: Exit For
                    End If
                Else
                    mrs��������.Filter = "����='" & mobjBill.Pages(p).Details(i).Detail.���� & "'" & strSQL
                    If mrs��������.EOF Then
                        If MsgBox(strTmp & "�����е� " & i & " ����Ŀ""" & mobjBill.Pages(p).Details(i).Detail.���� & """�ķ�������Ϊ""" & _
                            mobjBill.Pages(p).Details(i).Detail.���� & """,����" & _
                            IIf(strType = "1", "ҽ��", "����") & "�������ͣ�" & vbCrLf & "ȷʵҪ���浥����", _
                            vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            CheckFeeType = False: Exit For
                        End If
                    End If
                End If
            Next
        Next
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 

Private Function ItemExist(lng�շ�ϸĿID As Long) As Boolean
    Dim i As Integer, p As Integer
    
    If CheckBillsEmpty Then Exit Function
    
    For p = 1 To mobjBill.Pages.Count
        For i = 1 To mobjBill.Pages(p).Details.Count
            If mobjBill.Pages(p).Details(i).�շ�ϸĿID = lng�շ�ϸĿID Then
                ItemExist = True: Exit Function
            End If
        Next
    Next
End Function

Private Function CheckExecuteDept(intPage As Long) As Integer
'���ܣ���鵥�����Ƿ�����δ����ִ�п���
    Dim i As Integer, p As Integer
    
    For p = 1 To mobjBill.Pages.Count
        For i = 1 To mobjBill.Pages(p).Details.Count
            If mobjBill.Pages(p).Details(i).ִ�в���ID = 0 Then
                intPage = p: CheckExecuteDept = i: Exit Function
            End If
        Next
    Next
End Function
Private Sub InitBalanceGrid(Optional blnOnlyClearBalace As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ�����ս�����
    '���:blnOnlyBalace-�������������Ϣ
    '����:���˺�
    '����:2011-11-02 13:53:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    
    vsBalance.Clear
    vsBalance.Rows = 4
    If mbytInFun = 0 And InStr("01245", mbytInState) > 0 Then
        ''����:44425
        'mbytInState As Byte '0-ִ��(���޸�),1-���,2-����,3-�˷�(�շѡ����ʲ����˷�),4-�����շ�;5-�쳣��������
        vsBalance.Width = 2415 * 1.4
        Call picAppend_Resize
    Else
        vsBalance.Width = 2415 * 1.2
        Call picAppend_Resize
    End If
    vsBalance.ColWidth(0) = (vsBalance.Width - 300) * 0.6
    vsBalance.ColWidth(1) = (vsBalance.Width - 300) * 0.4
    vsBalance.ColAlignment(0) = 1
    vsBalance.ColAlignment(1) = 7
    If mbytInState = 3 And mbytInFun = 0 Then vsBalance.Editable = flexEDKbdMouse
    vsBalance.Row = 0
    vsBalance.Col = 1
    vsBalance.TabStop = False
    With vsBalance
        .Cell(flexcpFontBold, 0, 0, .Rows - 1, .COLS - 1) = False
        .Cell(flexcpForeColor, 0, 0, .Rows - 1, .COLS - 1) = Me.ForeColor
    End With
    For i = 0 To vsBalance.Rows - 1
        vsBalance.RowData(i) = 0
    Next
    If blnOnlyClearBalace Then Exit Sub
    '������㼯����
    Set mcolBalance = New Collection
    For i = 1 To tbsBill.Tabs.Count
        mcolBalance.Add Array()
    Next
End Sub
Private Sub ShowPrePayInfo(Optional blnShow As Boolean)
      '65748:ȡ��Ԥ������Ȩ��
    txtԤ�����.Enabled = blnShow And mbytInState = 0  'And InStr(1, mstrPrivs, "Ԥ������") > 0
    sta.Panels(Pan.C4Ԥ����Ϣ).Visible = blnShow
    
    If blnShow Then
        lblԤ�����.ForeColor = 0
        txtԤ�����.ForeColor = 0
        txtԤ�����.Text = "0.00"
    Else
        lblԤ�����.ForeColor = &H808080
        txtԤ�����.ForeColor = &H808080
        txtԤ�����.Text = "0.00"
        sta.Panels(Pan.C4Ԥ����Ϣ).Tag = ""
        sta.Panels(Pan.C4Ԥ����Ϣ).Text = ""
    End If
End Sub

Private Sub ShowPayInfo(Optional blnShow As Boolean)
    txtӦ��.Enabled = blnShow
    If blnShow Then
        lblӦ��.ForeColor = 0
        txtӦ��.ForeColor = &HFF0000
    Else
        lblӦ��.ForeColor = &H808080
        txtӦ��.ForeColor = &H808080
    End If
End Sub

Public Function GetMedicareSum(Optional strItem As String, Optional intPage As Integer, Optional blnOrig As Boolean) As Currency
'���ܣ���ȡ���ս���Ľ��
'������strItem=�Ƿ�ָ�����㷽ʽ,����Ϊ���н��㷽ʽ
'      intPage=�Ƿ�ָ������,����Ϊ���е���
'      blnOrig=�Ƿ�ȡԭʼ(���)������,����ȡ����(�޸ĺ�)��Ч���
'˵�����ú�����mcolBalanceΪ׼����,����ҽ�������շ�Ҳ��
    Dim arrValue As Variant, curMoney As Currency
    Dim i As Integer, p As Integer
    
    For p = IIf(intPage = 0, 1, intPage) To IIf(intPage = 0, mcolBalance.Count, intPage)
        For i = 0 To UBound(mcolBalance(p))
            '���㷽ʽ;ԭʼ(���)���;�ɷ��޸�;��Ч���
            arrValue = Split(mcolBalance(p)(i), ";")
            If strItem = "" Or (strItem <> "" And arrValue(0) = strItem) Then
                If blnOrig Then
                    curMoney = curMoney + CCur(arrValue(1))
                Else
                    curMoney = curMoney + CCur(arrValue(3))
                End If
            End If
        Next
    Next
    GetMedicareSum = Format(curMoney, "0.00")
End Function

Private Function GetMedicareStr(intPage As Integer) As String
'���ܣ����ر��ս��㷽ʽ��,"���㷽ʽ|���||...."
'������intPage=ָ�����ݱ��
'˵�����ú�����mcolBalanceΪ׼����,����ҽ�������շ�Ҳ��
    Dim i As Integer, strTmp As String
    Dim arrValue As Variant
    
    For i = 0 To UBound(mcolBalance(intPage))
        '���㷽ʽ;ԭʼ(���)���;�ɷ��޸�;��Ч���
        arrValue = Split(mcolBalance(intPage)(i), ";")
        strTmp = strTmp & "||" & arrValue(0) & "|" & Format(arrValue(3), "0.00")
    Next
    GetMedicareStr = Mid(strTmp, 3)
End Function

Private Sub SetBalanceVal(intPage As Integer, strItem As String, curVal As Currency)
'���ܣ�����ָ����ŵ���ָ�����ս��㷽ʽ����Чֵ
'˵�����ú�����mcolBalanceΪ׼����,����ҽ�������շ�Ҳ��
'˵������������ҽ���շ��޸ı��ս���������۵�ҽ���շ����ø����ʻ��Ƚ�����
    Dim arrValue As Variant, arrPage As Variant
    Dim blnDo As Boolean, strTmp As String, i As Long
        
    arrPage = Array()
    
    If UBound(mcolBalance(intPage)) >= 0 Then
        For i = 0 To UBound(mcolBalance(intPage))
            '���㷽ʽ;ԭʼ(���)���;�ɷ��޸�;��Ч���
            arrValue = Split(mcolBalance(intPage)(i), ";")
            If arrValue(0) = strItem And arrValue(3) <> curVal Then
                blnDo = True
                strTmp = arrValue(0) & ";" & arrValue(1) & ";" & arrValue(2) & ";" & Format(curVal, "0.00")
            Else
                strTmp = arrValue(0) & ";" & arrValue(1) & ";" & arrValue(2) & ";" & arrValue(3)
            End If
            ReDim Preserve arrPage(UBound(arrPage) + 1)
            arrPage(UBound(arrPage)) = strTmp
        Next
    Else
        '������ʱǿ������:��֧��Ԥ�����ҽ�������շ�ʱ��
        ReDim Preserve arrPage(UBound(arrPage) + 1)
        arrPage(UBound(arrPage)) = strItem & ";" & Format(curVal, "0.00") & ";0;" & Format(curVal, "0.00")
        blnDo = True
    End If
    
    '���µ��ݽ��㼯(���ϲ���ֱ�Ӹ���)
    If blnDo Then
        mcolBalance.Remove intPage
        If mcolBalance.Count >= intPage Then
            mcolBalance.Add arrPage, , intPage
        Else
            mcolBalance.Add arrPage
        End If
    End If
End Sub

Private Function GetExecDepts(Optional ByVal i As Integer) As String
'����:��ȡĳ���ŵ������е�ִ�в���,�������۵��շ�
'����:i-�������,���i=0,���ȡ���е���
    Dim j As Integer, p As Integer, strTmp As String
    
    For p = IIf(i = 0, 1, i) To IIf(i = 0, mobjBill.Pages.Count, i)
        For j = 1 To mobjBill.Pages(p).Details.Count
            If mobjBill.Pages(p).NO = "" Then
                If InStr(1, "," & strTmp & ",", "," & mobjBill.Pages(p).Details(j).ִ�в���ID & ",") <= 0 Then
                    strTmp = strTmp & "," & mobjBill.Pages(p).Details(j).ִ�в���ID
                End If
            End If
        Next
        If gTy_Module_Para.bln���ռ��Ʊ�� Then
            If mobjBill.Pages(p).����� <> 0 Then '������ִ�в��Ź̶�Ϊ����Աȱʡ����,��zl_�����շ����_Insert
                If InStr(1, "," & strTmp & ",", "," & UserInfo.����ID & ",") <= 0 Then
                    strTmp = strTmp & "," & UserInfo.����ID
                End If
            End If
        End If
    Next
    GetExecDepts = Mid(strTmp, 2)
End Function
Private Function GetInvoiceCount() As Integer
    '���ܣ����㵱ǰ�շ���Ҫ��ӡ������Ʊ��
    '˵�������������ṹ
    '   ���ŵ��ݷֱ��ӡ--��ִ�п��ҷֱ��ӡ--���շ�ϸĿ���վݷ�Ŀ��ӡ
    '   ��������ռ��Ʊ���У�����Ҫ�����￼��,��Ϊ����漰������,�����Ӱ�����ѵ�����
                    
    Dim rsTmp As ADODB.Recordset
    Dim strItems As String, strSQL As String, strNos As String, strTmp As String
    Dim i As Integer, j As Integer, k As Integer, X As Integer, intid As Integer, cur�����н�� As Currency
    Dim strִ�в���IDs As String, lngִ�в���ID As Long
    Dim str��Ʊ�� As String, int���� As Integer
    On Error GoTo errH
    
    '����Ʊ���Ƿ����
    '25187
    If gTy_Module_Para.bytƱ�ݷ������ <> 0 Then
        strNos = ""
        For i = 1 To mobjBill.Pages.Count
                If mobjBill.Pages(i).NO <> "" Then
                    strNos = strNos & "," & mobjBill.Pages(i).NO
                End If
        Next
        If strNos <> "" Then strNos = Mid(strNos, 2)
        If strNos = "" Then GetInvoiceCount = 1: Exit Function
        Call zlExeCuteBillNoSplit(True, 1, mlng����ID, strNos, 0, txtInvoice.Text, Now, 1, str��Ʊ��, int����)
        If mintInvoicePrint <> 0 Then
            '����ӡ,�����
            Call zlCheckFactIsEnough(int����)
        End If
        GetInvoiceCount = int����
        Exit Function
    End If
    
    
    If gTy_Module_Para.blnһ��Ʊ�� Then
        If mobjBill.Pages.Count > 1 And gTy_Module_Para.bln�ֱ��ӡ And mbytBillSource <> 4 Then
            GetInvoiceCount = mobjBill.Pages.Count
        Else
            GetInvoiceCount = 1
        End If
        Exit Function
    End If
    
    
    If mobjBill.Pages.Count > 1 And gTy_Module_Para.bln�ֱ��ӡ And mbytBillSource <> 4 Then
        'a.���ŷֱ��ӡ(ÿ�Ŷ���)
        For i = 1 To mobjBill.Pages.Count
            'a.aÿ�Ű�ִ�п��ҷֱ��ӡ
            '------------------------------------------------
            If gTy_Module_Para.bytƱ�����ɷ�ʽ >= 10 Then
                'a.a.aֱ���շѵ�
                If mobjBill.Pages(i).NO = "" Then
                    strִ�в���IDs = GetExecDepts(i)
                    For intid = 0 To UBound(Split(strִ�в���IDs, ","))
                        lngִ�в���ID = Val(Split(strִ�в���IDs, ",")(intid))
                        For j = 1 To mobjBill.Pages(i).Details.Count
                            If Not mobjBill.Pages(i).Details(j).������ And mobjBill.Pages(i).Details(j).ִ�в���ID = lngִ�в���ID Then '�ſ�������
                                If gTy_Module_Para.bytƱ�����ɷ�ʽ = 10 Then
                                    For k = 1 To mobjBill.Pages(i).Details(j).InComes.Count
                                        If mobjBill.Pages(i).Details(j).InComes(k).ʵ�ս�� <> 0 Then '��Ϊ��
                                            strTmp = mobjBill.Pages(i).Details(j).InComes(k).�վݷ�Ŀ
                                            If InStr("," & strItems & ",", "," & strTmp & ",") = 0 Then strItems = strItems & "," & strTmp
                                        End If
                                    Next
                                Else
                                    k = k + 1
                                End If
                            End If
                        Next
                        If gTy_Module_Para.bln���ռ��Ʊ�� And mobjBill.Pages(i).����� <> 0 And lngִ�в���ID = UserInfo.����ID Then
                            If gTy_Module_Para.bytƱ�����ɷ�ʽ = 10 Then
                                If InStr("," & strItems & ",", "," & gstr����վݷ�Ŀ & ",") = 0 Then strItems = strItems & "," & gstr����վݷ�Ŀ
                            Else
                                k = k + 1
                            End If
                        End If
                        
                        If gTy_Module_Para.bytƱ�����ɷ�ʽ = 10 Then
                            If strItems <> "" Then X = X + IntEx((UBound(Split(Mid(strItems, 2), ",")) + 1) / gTy_Module_Para.byt�����վ��д�)
                            strItems = ""
                        Else
                            X = X + IntEx(k / gTy_Module_Para.byt�����վ��д�)
                            k = 0
                        End If
                    Next
                'a.a.b���۵��շ�
                Else
                    '��������
                    If gTy_Module_Para.bln���ռ��Ʊ�� And mobjBill.Pages(i).����� <> 0 Then
                        strSQL = "Select count(��Ŀ) AS num" & vbNewLine & _
                                "From (Select " & IIf(gTy_Module_Para.bytƱ�����ɷ�ʽ = 10, "�վݷ�Ŀ", "�շ�ϸĿid") & " as ��Ŀ, ִ�в���id" & vbNewLine & _
                                "            From ������ü�¼" & vbNewLine & _
                                "            Where ��¼���� = 1 And ��¼״̬ = 0 And Nvl(ʵ�ս��, 0) <> 0 And No = [1]" & vbNewLine & _
                                "            Union" & vbNewLine & _
                                "            Select " & IIf(gTy_Module_Para.bytƱ�����ɷ�ʽ = 10, "'" & gstr����վݷ�Ŀ & "'", glng���ϸĿID) & " as ��Ŀ," & UserInfo.����ID & vbNewLine & _
                                "            From Dual)" & vbNewLine & _
                                "Group By ִ�в���id"
                    Else
                        strSQL = "Select Count(" & IIf(gTy_Module_Para.bytƱ�����ɷ�ʽ = 10, "Distinct �վݷ�Ŀ", "ID") & ") AS num From ������ü�¼" & _
                            " Where ��¼����=1 And ��¼״̬=0 And Nvl(ʵ�ս��,0)<>0 And NO=[1]" & _
                            " Group by ִ�в���id"
                    End If
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mobjBill.Pages(i).NO)
                    Do While Not rsTmp.EOF
                        X = X + IntEx(rsTmp!Num / gTy_Module_Para.byt�����վ��д�)
                        rsTmp.MoveNext
                    Loop
                End If
                
            'a.b����ִ�п��ҷֱ��ӡ
            '---------------------------------------------
            Else
                If mobjBill.Pages(i).NO = "" Then
                    For j = 1 To mobjBill.Pages(i).Details.Count
                        If Not mobjBill.Pages(i).Details(j).������ Then '�ſ�������
                            If gTy_Module_Para.bytƱ�����ɷ�ʽ = 0 Then
                                For k = 1 To mobjBill.Pages(i).Details(j).InComes.Count
                                    If mobjBill.Pages(i).Details(j).InComes(k).ʵ�ս�� <> 0 Then '��Ϊ��
                                        strTmp = mobjBill.Pages(i).Details(j).InComes(k).�վݷ�Ŀ
                                        If InStr("," & strItems & ",", "," & strTmp & ",") = 0 Then strItems = strItems & "," & strTmp
                                    End If
                                Next
                            Else
                                k = k + 1
                            End If
                        End If
                        If gTy_Module_Para.bln���ռ��Ʊ�� And mobjBill.Pages(i).����� <> 0 Then
                            If gTy_Module_Para.bytƱ�����ɷ�ʽ = 0 Then
                                If InStr("," & strItems & ",", "," & gstr����վݷ�Ŀ & ",") = 0 Then strItems = strItems & "," & gstr����վݷ�Ŀ
                            Else
                                k = k + 1
                            End If
                        End If
                    Next
                    If gTy_Module_Para.bytƱ�����ɷ�ʽ = 0 Then
                        If strItems <> "" Then X = X + IntEx((UBound(Split(Mid(strItems, 2), ",")) + 1) / gTy_Module_Para.byt�����վ��д�)
                        strItems = ""
                    Else
                        X = X + IntEx(k / gTy_Module_Para.byt�����վ��д�)
                        k = 0
                    End If
                Else
                    If gTy_Module_Para.bln���ռ��Ʊ�� And mobjBill.Pages(i).����� <> 0 Then
                        strSQL = "Select count(��Ŀ) AS num" & vbNewLine & _
                                "From (Select " & IIf(gTy_Module_Para.bytƱ�����ɷ�ʽ = 0, "�վݷ�Ŀ", "�շ�ϸĿid") & " as ��Ŀ" & vbNewLine & _
                                "            From ������ü�¼" & vbNewLine & _
                                "            Where ��¼���� = 1 And ��¼״̬ = 0 And Nvl(ʵ�ս��, 0) <> 0 And No = [1]" & vbNewLine & _
                                "            Union" & vbNewLine & _
                                "            Select " & IIf(gTy_Module_Para.bytƱ�����ɷ�ʽ = 0, "'" & gstr����վݷ�Ŀ & "'", glng���ϸĿID) & " as ��Ŀ" & vbNewLine & _
                                "            From Dual)"
                    Else
                        strSQL = "Select Count(" & IIf(gTy_Module_Para.bytƱ�����ɷ�ʽ = 0, "Distinct �վݷ�Ŀ", "ID") & ") AS num From ������ü�¼" & _
                            " Where ��¼����=1 And ��¼״̬=0 And Nvl(ʵ�ս��,0)<>0 And NO=[1]"
                    End If
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mobjBill.Pages(i).NO)
                    X = X + IntEx(rsTmp!Num / gTy_Module_Para.byt�����վ��д�)
                End If
            End If
        Next
        
    'b.ֻ��һ��,���ж��ŵ���һ���ӡ
    '---------------------------------------------------------------------------
    Else
        'b.a��ִ�п��ҷֱ��ӡ
        '----------------------------------------------
        If gTy_Module_Para.bytƱ�����ɷ�ʽ >= 10 Then
            strִ�в���IDs = GetExecDepts()   '���е��ݵ�ִ�в���,�������۵�������
            
            '���ռ����еĻ��۵�,����һ���
            For i = 1 To mobjBill.Pages.Count
                If mobjBill.Pages(i).NO <> "" Then strNos = strNos & ",'" & mobjBill.Pages(i).NO & "'"
            Next
            If strNos <> "" Then
                strNos = Mid(strNos, 2)
                strSQL = "Select Distinct " & IIf(gTy_Module_Para.bytƱ�����ɷ�ʽ = 10, "�վݷ�Ŀ", "�շ�ϸĿID") & " AS ��Ŀ,ִ�в���id From ������ü�¼" & _
                    " Where ��¼����=1 And ��¼״̬=0 And Nvl(ʵ�ս��,0)<>0 And " & IIf(InStr(1, strNos, ",") > 0, "NO IN(" & strNos & ")", " NO = [1]")
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Replace(strNos, "'", ""))
                
                Do While Not rsTmp.EOF
                    If InStr(1, "," & strִ�в���IDs & ",", "," & rsTmp!ִ�в���ID & ",") = 0 Then strִ�в���IDs = strִ�в���IDs & "," & rsTmp!ִ�в���ID
                    rsTmp.MoveNext
                Loop
                If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst   '���滹Ҫ�õ�
            End If
            
            If InStr(1, strִ�в���IDs, ",") = 1 Then strִ�в���IDs = Mid(strִ�в���IDs, 2)
            
            '�ٺ�ֱ���շѵ�һ����
            For intid = 0 To UBound(Split(strִ�в���IDs, ","))
                lngִ�в���ID = Val(Split(strִ�в���IDs, ",")(intid))
                For i = 1 To mobjBill.Pages.Count
                    If mobjBill.Pages(i).NO = "" Then
                        For j = 1 To mobjBill.Pages(i).Details.Count
                            If Not mobjBill.Pages(i).Details(j).������ And mobjBill.Pages(i).Details(j).ִ�в���ID = lngִ�в���ID Then '�ſ�������
                                If gTy_Module_Para.bytƱ�����ɷ�ʽ = 10 Then
                                    For k = 1 To mobjBill.Pages(i).Details(j).InComes.Count
                                        If mobjBill.Pages(i).Details(j).InComes(k).ʵ�ս�� <> 0 Then '��Ϊ��
                                            strTmp = mobjBill.Pages(i).Details(j).InComes(k).�վݷ�Ŀ
                                            If strTmp = gstr����վݷ�Ŀ Then strTmp = i & "-" & strTmp '��i��Ϊ�˱��ں�������ѵ�ִ�п��Ҵ���
                                            If InStr("," & strItems & ",", "," & strTmp & ",") = 0 Then strItems = strItems & "," & strTmp
                                        End If
                                    Next
                                Else    '����Ϊ������ڱ���ǰ�Ѽ���ֹ����
                                    strTmp = mobjBill.Pages(i).Details(j).�շ�ϸĿID
                                    If strTmp = glng���ϸĿID Then strTmp = i & "-" & strTmp
                                    If InStr("," & strItems & ",", "," & strTmp & ",") = 0 Then strItems = strItems & "," & strTmp
                                End If
                            End If
                        Next
                    End If
                    
                    '���۵���ֱ���շѵ����һ����
                    '����һ���,��������,�������ÿ�ŵ���һ��,����Ҫ��i������
                    If gTy_Module_Para.bln���ռ��Ʊ�� And mobjBill.Pages(i).����� <> 0 And lngִ�в���ID = UserInfo.����ID Then
                        If gTy_Module_Para.bytƱ�����ɷ�ʽ = 10 Then
                            If InStr("," & strItems & ",", "," & i & "-" & gstr����վݷ�Ŀ & ",") = 0 Then strItems = strItems & "," & i & "-" & gstr����վݷ�Ŀ
                        Else
                            If InStr("," & strItems & ",", "," & i & "-" & glng���ϸĿID & ",") = 0 Then strItems = strItems & "," & i & "-" & glng���ϸĿID
                        End If
                    End If
                Next
                
                '�ٴ������е��շѻ��۵�
                If strNos <> "" And Not rsTmp Is Nothing Then
                    rsTmp.Filter = "ִ�в���id=" & lngִ�в���ID
                    For k = 1 To rsTmp.RecordCount
                        If InStr("," & strItems & ",", "," & rsTmp!��Ŀ & ",") = 0 Then strItems = strItems & "," & rsTmp!��Ŀ
                        rsTmp.MoveNext
                    Next
                End If
                
                '�����շѵ���ֱ���շѵ����ܻ���,������Ҫ,�����
                If strItems <> "" Then X = X + IntEx((UBound(Split(Mid(strItems, 2), ",")) + 1) / gTy_Module_Para.byt�����վ��д�)
                strItems = ""
            Next
            
            
        'b.b����ִ�п��ҷֱ��ӡ
        '-----------------------------------------------------
        Else
            For i = 1 To mobjBill.Pages.Count
                If mobjBill.Pages(i).NO = "" Then
                    For j = 1 To mobjBill.Pages(i).Details.Count
                        If Not mobjBill.Pages(i).Details(j).������ Then '�ſ�������
                            If gTy_Module_Para.bytƱ�����ɷ�ʽ = 0 Then
                                For k = 1 To mobjBill.Pages(i).Details(j).InComes.Count
                                    If mobjBill.Pages(i).Details(j).InComes(k).ʵ�ս�� <> 0 Then '��Ϊ��
                                        strTmp = mobjBill.Pages(i).Details(j).InComes(k).�վݷ�Ŀ
                                        If strTmp = gstr����վݷ�Ŀ Then strTmp = i & "-" & strTmp
                                        If InStr("," & strItems & ",", "," & strTmp & ",") = 0 Then strItems = strItems & "," & strTmp
                                    End If
                                Next
                            Else    '����Ϊ������ڱ���ǰ�Ѽ���ֹ����
                                strTmp = mobjBill.Pages(i).Details(j).�շ�ϸĿID
                                If strTmp = glng���ϸĿID Then strTmp = i & "-" & strTmp
                                If InStr("," & strItems & ",", "," & strTmp & ",") = 0 Then strItems = strItems & "," & strTmp
                            End If
                        End If
                    Next
                Else
                    strNos = strNos & ",'" & mobjBill.Pages(i).NO & "'"
                End If
                '���۵���ֱ���շѵ����һ����
                '����һ���,��������,�������ÿ�ŵ���һ��,����Ҫ��i������
                If gTy_Module_Para.bln���ռ��Ʊ�� And mobjBill.Pages(i).����� <> 0 Then
                    If gTy_Module_Para.bytƱ�����ɷ�ʽ = 0 Then
                        If InStr("," & strItems & ",", "," & i & "-" & gstr����վݷ�Ŀ & ",") = 0 Then strItems = strItems & "," & i & "-" & gstr����վݷ�Ŀ
                    Else
                        If InStr("," & strItems & ",", "," & i & "-" & glng���ϸĿID & ",") = 0 Then strItems = strItems & "," & i & "-" & glng���ϸĿID
                    End If
                End If
            Next
            If strNos <> "" Then
                strNos = Mid(strNos, 2)
                strSQL = "Select Distinct " & IIf(gTy_Module_Para.bytƱ�����ɷ�ʽ = 0, "�վݷ�Ŀ", "�շ�ϸĿID") & " AS ��Ŀ,ִ�в���id From ������ü�¼" & _
                    " Where ��¼����=1 And ��¼״̬=0 And Nvl(ʵ�ս��,0)<>0 And " & IIf(InStr(1, strNos, ",") > 0, "NO IN(" & strNos & ")", " NO = [1]")
                
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Replace(strNos, "'", ""))
                For k = 1 To rsTmp.RecordCount
                    If InStr("," & strItems & ",", "," & rsTmp!��Ŀ & ",") = 0 Then strItems = strItems & "," & rsTmp!��Ŀ
                    rsTmp.MoveNext
                Next
            End If
            ''�����շѵ���ֱ���շѵ����ܻ���,������Ҫ,�����
            X = IntEx((UBound(Split(Mid(strItems, 2), ",")) + 1) / gTy_Module_Para.byt�����վ��д�)
        End If
    End If
    GetInvoiceCount = X
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function GetBillSumByDB(strNo As String) As Currency
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
        strSQL = "Select SUM(ʵ�ս��) AS ʵ�ս�� From ������ü�¼ " & _
                " Where ��¼����=1 And ��¼״̬=0 And NO=[1] And ����Ա���� is Null"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo)
        If Not rsTmp.EOF Then
            GetBillSumByDB = Nvl(rsTmp!ʵ�ս��, 0)
        Else
            GetBillSumByDB = 0
        End If
        Exit Function
errH:
        If ErrCenter() = 1 Then Resume
        Call SaveErrLog
End Function

Private Function Check�����������(ByVal lng����ID As Long, ByVal intInsure As Long) As Boolean
'���ܣ�����ָ������ID��ҽ�����㷽ʽ�����ж��Ƿ�ȫ��֧�������������
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Integer
    
    strSQL = "Select B.���㷽ʽ From ����Ԥ����¼ B,���㷽ʽ C" & _
        " Where B.��¼����=3 And B.���㷽ʽ=C.���� And Nvl(C.����,1) IN(3,4)  And B.����ID=[1]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
        
    If rsTmp.RecordCount > 0 Then
       Check����������� = True
       For i = 1 To rsTmp.RecordCount
           If Not gclsInsure.GetCapability(support�����������, , intInsure, rsTmp!���㷽ʽ) Then
                MsgBox "ҽ�����㷽ʽ:" & rsTmp!���㷽ʽ & ",��֧�������������!" & vbCrLf & _
                "�޸ĵ���Ҫ�����н��㷽ʽȫ�����Ϻ����²����µ���,���Բ����޸Ĵ˵���!", vbInformation, gstrSysName
                Check����������� = False
                Exit For
           End If
       Next
    Else
        MsgBox "���ݷ����쳣,��ȡ������ʹ�õ�ҽ�����㷽ʽʧ��,�����޸Ĵ˵��ݣ�", vbInformation, gstrSysName
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ShowRegist()
'���ܣ�����Ƿ������ʾ�ҺŰ�ť
    Dim strPrivs As String
    On Error GoTo errH
    
    If mbytInFun = 0 And mbytInState = 0 Then
        strPrivs = GetPrivFunc(glngSys, 1111)
        If InStr(";" & strPrivs & ";", ";�Һ�;") > 0 Then '�����Ƿ���Ȩ
            cmdRegist.Visible = True
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub ShowIDCard()
'���ܣ�����Ƿ������ʾ���￨��ť
    Dim strPrivs As String
    On Error GoTo errH
    
    If mbytInFun = 0 And mbytInState = 0 Then
        strPrivs = GetPrivFunc(glngSys, 1107)
        If InStr(";" & strPrivs & ";", ";����;") > 0 Then '�����Ƿ���Ȩ
            cmdIDCard.Visible = True
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function GetOperatorInfo(ByVal str���� As String, Optional bln��ʿ As Boolean, Optional intְ�� As Integer) As Boolean
'���ܣ���ȡָ������������(ҽ����ʿ)�����ʻ�ְ��
'���أ�intְ��:0-δ���ã�bln��ʿ:�Ƿ�ֻ�ǻ�ʿ
'˵������ǰ��ֱ�Ӷ�ȡmarrDr�е�����,��Ϊ�൥�ݶ࿪���˺�һЩ�ط�����
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim i As Long
    
    bln��ʿ = False: intְ�� = 0
    If Not mrs������ Is Nothing Then
        mrs������.Filter = "����='" & str���� & "' " & IIf(gbln��ʿ, "", " And ��Ա����<>'��ʿ'")
        If mrs������.RecordCount > 0 Then
            intְ�� = mrs������!ְ��
            strSQL = mrs������!��Ա����
            If strSQL = "��ʿ" Then bln��ʿ = True
            If strSQL = "ҽ��" Then bln��ʿ = False
        End If
    Else
        strSQL = _
            " Select Nvl(A.Ƹ�μ���ְ��,0) as ְ��,B.��Ա���� From ��Ա�� A,��Ա����˵�� B" & _
            " Where A.ID=B.��ԱID And B.��Ա���� IN('ҽ��','��ʿ') And A.����=[1] And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str����)
        If Not rsTmp.EOF Then
            intְ�� = rsTmp!ְ��
            Do While Not rsTmp.EOF
                If rsTmp!��Ա���� = "��ʿ" Then bln��ʿ = True
                If rsTmp!��Ա���� = "ҽ��" Then bln��ʿ = False: Exit Do
                rsTmp.MoveNext
            Loop
        End If
    End If
    GetOperatorInfo = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function YBIdentifyCancel() As Boolean
'���ܣ�ȡ��ҽ�����������֤
'���أ����ؼ�ʱ���˳�������������
    Dim lng����ID As Long
    YBIdentifyCancel = True
    If mbytInFun = 0 And mbytInState = 0 Then
        If mstrYBPati <> "" And txtPatient.Text <> "" Then
            If UBound(Split(mstrYBPati, ";")) >= 8 Then
                If IsNumeric(Split(mstrYBPati, ";")(8)) And Val(Split(mstrYBPati, ";")(8)) <> 0 Then
                    lng����ID = Val(CLng(Split(mstrYBPati, ";")(8)))
                End If
            End If
            If lng����ID <> 0 Then
                YBIdentifyCancel = gclsInsure.IdentifyCancel(0, lng����ID, mintInsure)
            End If
        End If
    End If
End Function

Private Function SelectPatient() As Long
'���ܣ���ʾ��Լ��λ���˲�ѡ��
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    
    strSQL = _
        "Select 0 As ĩ��, ID, �ϼ�id, -null As �����, '[' || ���� || ']' || ���� As ����, Null As �Ա�, Null As ����," & vbNewLine & _
        "       Null As �ѱ�, Null As ���ʽ, Null As ��������, Null As ��ͥ��ַ" & vbNewLine & _
        "From ��Լ��λ" & vbNewLine & _
        "Where ����ʱ�� Is Null Or ����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')" & vbNewLine & _
        "Start With �ϼ�id Is Null" & vbNewLine & _
        "Connect By Prior ID = �ϼ�id" & vbNewLine & _
        "Union All" & vbNewLine & _
        "Select 1 As ĩ��, A.����id As ID, A.��ͬ��λid As �ϼ�id, A.�����, A.����, A.�Ա�, A.����, A.�ѱ�," & vbNewLine & _
        "       A.ҽ�Ƹ��ʽ As ���ʽ, To_Char(A.��������, 'YYYY-MM-DD') As ��������, A.��ͥ��ַ" & vbNewLine & _
        "From ������Ϣ A, ��Լ��λ B" & vbNewLine & _
        "Where B.ID = A.��ͬ��λid And A.ͣ��ʱ�� Is Null And A.��ǰ����id Is Null And A.��ͬ��λid Is Not Null And B.ĩ�� = 1"
            
    Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 2, "��Լ��λ����", , , "����ѡ���Լ��λ����ѡ����", , , , , , , blnCancel, , , 1)
    If rsTmp Is Nothing Then Exit Function
    SelectPatient = rsTmp!ID
End Function

Private Sub SetBillRowForeColor(ByVal lngRow As Long, ByVal lngColor As Long)
    Dim lngPreRow As Long, lngPreCol As Long
    Dim blnActive As Boolean, blnRedraw As Boolean
    Dim i As Long
    
    '��������
    mblnEnterCell = False
    lngPreRow = Bill.Row: lngPreCol = Bill.Col
    blnActive = Bill.Active: blnRedraw = Bill.MsfObj.Redraw
        
    '��ʼ����
    Bill.Active = False: Bill.Redraw = False
    Bill.Row = lngRow
    For i = Bill.MsfObj.FixedCols To Bill.COLS - 1
        Bill.Col = i: Bill.MsfObj.CellForeColor = lngColor
    Next
    
    '�ָ�����
    Bill.Row = lngPreRow: Bill.Col = lngPreCol
    Bill.Active = blnActive: Bill.Redraw = blnRedraw
    mblnEnterCell = True
End Sub

Private Sub SetItemRowColor(ByVal intPage As Integer, ByVal lngRow As Long)
'���ܣ�����ҩƷ/���ϵĴ����޶���������ɫ��ʾ
    If mobjBill.Pages(intPage).Details.Count >= lngRow And InStr(",0,1,", mbytInFun) > 0 And mbytInState = 0 Then
        With mobjBill.Pages(intPage).Details(lngRow)
            If mbln�����޶��� And (InStr(",5,6,7,", .�շ����) > 0 Or (.�շ���� = "4" And .Detail.��������)) Then
                If ItemUnderSet(.�շ����, .�շ�ϸĿID, .ִ�в���ID, IIf(gblnҩ����λ, .Detail.ҩ����װ, 1) * .Detail.���) Then
                    Call SetBillRowForeColor(lngRow, &HC000C0)
                Else
                    Call SetBillRowForeColor(lngRow, Bill.ForeColor)
                End If
            Else
                Call SetBillRowForeColor(lngRow, Bill.ForeColor)
            End If
        End With
    End If
End Sub

Private Function CheckSaveMultiPrice() As Boolean
    Dim p As Integer
    
    If mbytInFun = 0 And mbytInState = 0 And mstrInNO = "" And chkCancel.Value = 0 Then
        For p = 1 To mobjBill.Pages.Count
            If mobjBill.Pages(p).NO <> "" Then
                Exit Function
            End If
        Next
        CheckSaveMultiPrice = True  '������Ϊ���۵�
    Else
        Exit Function
    End If
End Function

Private Sub MergeRepeatItem()
'���ܣ��ϲ������������ظ�����ķ���/ʱ��ҩƷ/��������(��Ŀ��ִ�п�����ͬ)
'˵��������֮ǰӦ��ȷ����������ҩ��Ҫ�ϲ���������ͬ�����
    Dim i As Integer, j As Integer
    Dim m As Integer, n As Integer
    Dim objDetail As BillDetail
    Dim rsItem As New ADODB.Recordset
    Dim blnRefresh As Boolean
    
    rsItem.Fields.Append "Type", adBigInt
    rsItem.Fields.Append "Page", adBigInt
    rsItem.Fields.Append "Row", adBigInt
    rsItem.CursorLocation = adUseClient
    rsItem.LockType = adLockOptimistic
    rsItem.CursorType = adOpenStatic
    rsItem.Open
        
    For i = 1 To mobjBill.Pages.Count
        For j = 1 To mobjBill.Pages(i).Details.Count
            With mobjBill.Pages(i).Details(j)
                If (.Detail.���� Or .Detail.���) And .���� * .���� <> 0 _
                    And (InStr(",5,6,7,", .�շ����) > 0 Or .�շ���� = "4" And .Detail.��������) Then
                    For m = i To mobjBill.Pages.Count
                        For n = IIf(m = i, j + 1, 1) To mobjBill.Pages(m).Details.Count
                            Set objDetail = mobjBill.Pages(m).Details(n)
                            If objDetail.�շ�ϸĿID = .�շ�ϸĿID _
                                And objDetail.ִ�в���ID = .ִ�в���ID And objDetail.���� * objDetail.���� <> 0 Then
                                .���� = .���� + objDetail.����
                                objDetail.���� = 0
                                                                
                                rsItem.AddNew
                                rsItem!Type = 1 '�ϲ�������
                                rsItem!Page = i
                                rsItem!Row = j
                                rsItem.Update
                                                                
                                rsItem.AddNew
                                rsItem!Type = 2 '���ϲ�����
                                rsItem!Page = m
                                rsItem!Row = n
                                rsItem.Update
                            End If
                        Next
                    Next
                End If
            End With
        Next
    Next
    
    If rsItem.RecordCount > 0 Then
        'ɾ�����ϲ�����(����)
        rsItem.Sort = "Page,Row Desc"
        rsItem.Filter = "Type=2"
        Do While Not rsItem.EOF
            Call DeleteDetail(rsItem!Row, rsItem!Page)
            If rsItem!Page = mintPage Then blnRefresh = True
            rsItem.MoveNext
        Loop
        
        '����ϲ�������
        For i = 1 To mobjBill.Pages.Count
            rsItem.Filter = "Type=1 And Page=" & i
            If rsItem.RecordCount > 1 Then          'һ�ŵ����м���ϲ�ʱ,ɾ���кź�,֮ǰ��¼�ĺϲ������кſ��ܱ���
                Call CalcMoneys(i)
            ElseIf rsItem.RecordCount = 1 Then
                Call CalcMoneys(rsItem!Page, rsItem!Row)
            End If
            If i = mintPage Then blnRefresh = True
        Next
    End If
    
    If blnRefresh Then
        Call ShowDetails
    End If
    Call ShowMoney
    
    '��Ҫ����Ԥ����
    If cmdԤ����.Visible Then
        Call InitBalanceGrid
        cmdԤ����.TabStop = True
        cmdOK.Enabled = False
    End If
End Sub
Public Function zlCheck����ҽ��(ByVal intInsure As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Ա���ҽ����һЩ���
    '���:intInsuer-����
    '����:
    '����:���ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-01-07 10:25:04
    '����:27278
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String, p As Long
    If intInsure = 0 Then zlCheck����ҽ�� = True: Exit Function
    
    Err = 0: On Error GoTo Errhand:
    '���˺�:???
    '��ڲ�����
    'mbytInFun As Byte '0-�շ�,1-����,2-�������
    'mbytInState As Byte '0-ִ��(���޸�),1-���,2-����,3-�˷�(�շѡ����ʲ����˷�)
    'mstrInNO As String '�����ĵ��ݺ�(�鿴���������޸ģ��˷ѣ�����)
    'mbytBilling As Byte 'mbytInFun=2ʱ��0-��������,1-���ʻ���,2-�������

    
    'ֻ�л��۲�֧�ּ��
    If mbytInFun <> 1 And mbytInState <> 0 Or MCPAR.ҽ��ȷ���������� = False Then
        zlCheck����ҽ�� = True: Exit Function
    End If
    'showmsgbox
    '������strCaption=��Ϣ�������
    '      strInfo=������ʾ����,����"^"��ʾ����,">"��ʾ������
    '      strCmds=��ť����,��"����(&R),!����(&A),?ȡ��(&C)"
    '              ����Ҫ��������ť,"!"��ʾȱʡ��ť,"?"��ʾȡ����ť
    '              ÿ����ť�������֧��4������
    '      vStyle=vbInformation,vbQuestion,vbExclamation,vbCritical
    '���أ���ť����,��"��ť2"(������()��&),������رջ�ȡ���򷵻�""
    strTemp = zlCommFun.ShowMsgbox("��������", "��ȷ����ǰҽ�����˱���Ҫ���͵�ҩƷ���������͡�", "!ҽ����(&A),ҽ����(&B),?ȡ��(&C)", Me)
    If strTemp = "" Or strTemp = "ȡ��" Then Exit Function
    '����ǲ������շѻ��۵�������ҽ�����ˣ���ҽ��������supportҽ��ȷ���������͡���Чʱ������ʱ��ʾ�õ����ǡ�ҽ���ڣ�ҽ���⡱�������ҽ���ڷ��ü�¼ժҪ�д��1��ҽ������2��
    strTemp = IIf(strTemp = "ҽ����", 1, 2)
    
    '--����ժҪ
    '��ÿ�ŵ��ݶ���ִ�б���
    For p = 1 To mobjBill.Pages.Count
        '����ÿ���շѵ��ݵĵ��ݺ�
        If mobjBill.Pages(p).NO = "" Then
            For Each mobjBillDetail In mobjBill.Pages(p).Details
                mobjBillDetail.ժҪ = strTemp
            Next
        End If
    Next
    zlCheck����ҽ�� = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function

Private Function Get��ˢ���() As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���㿨�Ŀ�ˢ���
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2010-01-08 14:53:45
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim i As Long, j As Long, intCol As Integer
    Dim dbl��ˢ��� As Double, dbl�ܶ� As Double
    
    With mobjBill
        For i = 1 To mobjBill.Pages.Count
            dbl��ˢ��� = dbl��ˢ��� + .Pages(i).ʵ�ս�� - .Pages(i).���ս�� - .Pages(i).��Ԥ����            ' + .Pages(i).�����:
            dbl�ܶ� = dbl�ܶ� + 0 + .Pages(i).ʵ�ս��
        Next
    End With
  
    '���û��,�ٳ��Դӱ����ȡ(��һ�ŵ���ʱ)
    If dbl�ܶ� = 0 And tbsBill.Tabs.Count = 1 _
        And Not (Bill.Rows = 2 And Bill.TextMatrix(1, BillCol.��Ŀ) = "") Then
        
        intCol = BillCol.ʵ�ս��
        For i = 1 To Bill.Rows - 1
            If IsNumeric(Bill.TextMatrix(i, intCol)) Then
                dbl�ܶ� = dbl�ܶ� + Format(Val(Bill.TextMatrix(i, intCol)), gstrDec)
            End If
        Next
        dbl��ˢ��� = dbl�ܶ� - Val(txtԤ�����.Text)
    End If
    Get��ˢ��� = Format(dbl��ˢ���, gstrDec)
End Function
  
Private Sub ShowStatusCargoSpace(ByVal lng�շ�ϸĿID As Long, lngִ�пⷿID As Long, _
    Optional bln���� As Boolean = False)
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ʾ�ⷿ��λ
    '���ƣ����˺�
    '���ڣ�2010-04-13 14:30:20
    '˵����27505(�������ɹ��ú���)
    '         Ŀǰֻ��Ի��۵�
    '------------------------------------------------------------------------------------------------------------------------
    Static lngPre�շ�ϸĿID As Long
    Static lngPreִ�пⷿID As Long
    Static strCargo_Space As String  '�ϴλ�λ
    Dim strTemp As String
    Err = 0: On Error GoTo Errhand:
    '����ʱҪ��ʾ�ⷿ��λ
    If mbytInFun <> 1 Then Exit Sub
    If Not (lngPre�շ�ϸĿID = lng�շ�ϸĿID And lngִ�пⷿID = lngPreִ�пⷿID) Then
        lngPre�շ�ϸĿID = lng�շ�ϸĿID: lngPreִ�пⷿID = lngִ�пⷿID
        strCargo_Space = GetPlace(lng�շ�ϸĿID, lngִ�пⷿID, bln����)     '���»�ȡ�ⷿ��λ
    End If
    If strCargo_Space <> "" And InStr(1, strCargo_Space, "��λ:") = 0 Then strCargo_Space = "��λ:" & strCargo_Space
    strTemp = Split(sta.Panels(Pan.C2��ʾ��Ϣ), ",��λ:")(0)
    strTemp = Split(strTemp, "��λ:")(0)
    If strTemp <> "" And strCargo_Space <> "" Then strTemp = strTemp & ","
    strTemp = strTemp & strCargo_Space
    sta.Panels(Pan.C2��ʾ��Ϣ) = strTemp    '��ʾ����λ
Errhand:
End Sub

Public Function zl��ȡ��ҩ��̬(Optional ByVal intPage As Integer = 0, Optional ByVal lngRow As Long = -1, Optional blnOnly�г�ҩ As Boolean = False) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�����Ƿ�¼�����в�ҩ��
    '���:intPage-��ǰ�ڼ�ҳ
    '     blnOnly�г�ҩ-���ж��Ƿ����г�ҩ(���䷽ʱ�ж���Ч):ԭ�����л�ҩ���䷽���Ѿ�����,�Ͳ���Ҫ���
    '     lngRow-��ǰ��������
    '����:
    '����:¼�����в�ҩ��,�򷵻��������(1-���,0-��Ҫ��),���򷵻�-1 ��ʾ��û��¼�������Ŀ
    '����:���˺�
    '����:2010-02-02 11:44:17
    '����:27816
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strTemp As String
    
    zl��ȡ��ҩ��̬ = -1
    '���δָ��ҳ,���õ�ǰҳ
    If intPage = 0 Then intPage = mintPage
    If mobjBill Is Nothing Then Exit Function
    strTemp = IIf(blnOnly�г�ҩ, ",6,", ",6,7,")
    With mobjBill.Pages(intPage).Details
        For i = 1 To .Count
            If InStr(1, strTemp, "," & .Item(i).�շ���� & ",") > 0 And .Item(i).�շ�ϸĿID <> 0 And i <> lngRow Then
                zl��ȡ��ҩ��̬ = .Item(i).Detail.��ҩ��̬
                Exit Function
            End If
        Next
    End With
End Function
Private Sub SetBill�в�ҩEditEnabled()
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ������в�ҩ�ı༭״̬
    '���ƣ����˺�
    '���ڣ�2010-08-06 10:58:45
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    With Bill
        For i = 0 To .COLS - 1
            If .TextMatrix(0, i) = "��Ŀ" Then
                .ColData(i) = 0
            Else
                .ColData(i) = 5
            End If
        Next
    End With
End Sub
'''Private Sub txt�Ҳ�_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'''    If lbl�Ҳ�.Caption <> "�Ҳ�" Then
'''        zlCommFun.ShowTipInfo txt�Ҳ�.hWnd, mstrӦ������㷽ʽ, False
'''    Else
'''        zlCommFun.ShowTipInfo txt�Ҳ�.hWnd, "", False
'''    End If
'''End Sub
Private Function IsRegisterDept() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Ƿ�ͨ���Һŵ���ȡ�Ĳ���
    '����:�Ƿ���true,���򷵻�False
    '����:���˺�
    '����:2010-11-19 15:31:01
    '����:34032
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    If mrsInfo Is Nothing Then Exit Function
    If mrsInfo.State <> 1 Then Exit Function
    For i = mrsInfo.Fields.Count - 1 To 0 Step -1
        If UCase(mrsInfo.Fields(i).Name) = "ִ�в���ID" Then
            IsRegisterDept = True: Exit Function
        End If
    Next
End Function
Private Sub zlCheckFactIsEnough(Optional ByVal intInvoicePages As Integer = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鵱ǰƱ���Ƿ�����
    ' ���:intInvoicePages-��Ҫ�ķ�Ʊ����,���Ϊ0,��ϵͳ��������
    '����:���˺�
    '����:2011-05-10 17:54:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngʣ������ As Long, lngNums As Long
    
    If Not (mbytInFun = 0 And mbytInState = 0) Then Exit Sub
    '���˺� ����:26948 ����:2009-12-28 17:43:00
    '��Ҫ���ʣ�������Ƿ����:
    If intInvoicePages <> 0 Then
        If zlCheckInvoiceOverplusEnough(1, intInvoicePages, lngʣ������, mlng����ID, mstrUseType) = False Then
            MsgBox "ע��:" & vbCrLf & _
                   "    ��ǰʣ��Ʊ�ݲ���(" & lngʣ������ & ") ,��ǰ��Ҫ" & intInvoicePages & "��Ʊ��,��ע�������Ʊ!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
        End If
    Else
        If zlCheckInvoiceOverplusEnough(1, gTy_Module_Para.int����ʣ��Ʊ������, lngʣ������, mlng����ID, mstrUseType) = False Then
            MsgBox "ע��:" & vbCrLf & _
                   "    ��ǰʣ��Ʊ��(" & lngʣ������ & ") С���˱���������(" & gTy_Module_Para.int����ʣ��Ʊ������ & "),��ע�������Ʊ!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
        End If
    End If
End Sub
Private Function zlCheckBill���ڷ�ɢװ��ҩ(intPage As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鵥���д��ڷ�ɢװ��ҩ��̬
    '���:intPage-ָ����ҳ
    '����:����,����true,���򷵻�False
    '����:���˺�
    '����:2011-05-26 10:19:46
    '����:38328
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    If mobjBill Is Nothing Then Exit Function
    With mobjBill.Pages(intPage)
        If .Details.Count = 0 Then Exit Function
        For i = 1 To .Details.Count
            If .Details(i).�շ���� = "7" Then
                If .Details(i).Detail.��ҩ��̬ <> 0 Then    '0-ɢװ;1-��ҩ��Ƭ;2-����
                    zlCheckBill���ڷ�ɢװ��ҩ = True: Exit Function
                End If
            End If
        Next
    End With
End Function

Private Sub initCardSquareData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���㿨����������Ϣ
    '���:blnClosed:�رն���
    '����:���˺�
    '����:2010-01-05 14:51:23
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mbytInState = 1 Then Exit Sub
    Dim objCard As Card
    If gobjSquare.objSquareCard Is Nothing Then Exit Sub
    
    Call IDKind.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, _
        gobjSquare.objSquareCard, "", txtPatient)
        
    Set objCard = IDKind.GetfaultCard
    Set gobjSquare.objDefaultCard = objCard
    If IDKind.Cards.��ȱʡ������ And Not objCard Is Nothing Then
        gobjSquare.blnȱʡ�������� = objCard.�������Ĺ��� <> ""
        gobjSquare.intȱʡ���ų��� = objCard.���ų���
    Else
        gobjSquare.blnȱʡ�������� = IDKind.Cards.������ʾ
        gobjSquare.intȱʡ���ų��� = 100
    End If
    gobjSquare.bln��ȱʡ������ = IDKind.Cards.��ȱʡ������
End Sub
Private Function zlCheckDelValied(ByVal lng�����ID As Long, _
     ByVal strName As String, _
     ByVal bln���ѿ� As Boolean, ByVal strCardNo As String, _
     ByVal strSwapNO As String, strSwapMemo As String, _
     ByRef lng����ID As Long, _
    ByVal dbl�˿��� As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����˷ѽ��׽ӿ�
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-02-08 16:40:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strXMLExend As String
    If lng�����ID = 0 Then zlCheckDelValied = True: Exit Function
    If Not gobjSquare.objSquareCard Is Nothing Then
        Call CreateSquareCardObject(gfrmMain, mlngModul)
        Call initCardSquareData
    End If
    
    If gobjSquare.objSquareCard Is Nothing Then
        MsgBox "ע��:" & vbCrLf & _
                     "      ��ǰ�շ��ǰ�" & strName & " �շѵ�,�������ڲ�������ز���,�����˿�,����ϵͳ����Ա��ϵ!", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    
'zlReturnCheck(frmMain As Object, ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, bln���ѿ� As Boolean, ByVal strCardNo As String, _
    ByVal strBalanceIDs As String, _
    ByVal dblMoney As Double, ByVal strSwapNo As String, _
    ByVal strSwapMemo As String, ByRef strXMLExpend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ʻ����˽���ǰ�ļ��
    '���:frmMain-���õ�������
    '       lngModule-���õ�ģ���
    '       lngCardTypeID-�����ID
    '       strCardNo-����
    '       strBalanceIDs   String  In  ����֧�����漰�Ľ���ID ��ʽ:�շ�����|ID1,ID2��IDn||�շ�����n|ID1,ID2��IDn
    '                                   �շ�����: 1-Ԥ����,2-����,3-�շ�,4-�Һ�,5-ҽ�ƿ��տ�
    '       dblMoney-�˿���
    '       strSwapNo-������ˮ��(�˿�ʱ���)
    '       strSwapMemo-����˵��(�˿�ʱ����)
    '       strXMLExpend    XML IN  ��ѡ����(��չ��).��δ����
    '����:�˿�Ϸ�,����true,���򷵻�Flase
      If gobjSquare.objSquareCard.zlReturnCheck(Me, mlngModul, lng�����ID, bln���ѿ�, strCardNo, _
        "3|" & lng����ID, dbl�˿���, strSwapNO, strSwapMemo, strXMLExend) = False Then
          zlCheckDelValied = False
          Exit Function
     End If
goEnd:
    zlCheckDelValied = True
    Exit Function
End Function
Private Function CheckBrushCard(ByVal lng�����ID As Long, ByVal bln���ѿ� As Boolean, _
    ByVal dbl�˷Ѷ� As Double, ByRef strBrushCardNo As String, ByRef strbrPassWord As String, Optional ByRef bln���� As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ˢ��
    '����:�Ϸ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-18 14:52:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsMoney As ADODB.Recordset
    On Error GoTo errHandle
    Dim dblMoney As Double
     '����ˢ������
    'zlBrushCard(frmMain As Object, _
    'ByVal lngModule As Long, _
    'ByVal rsClassMoney As ADODB.Recordset, _
    'ByVal lngCardTypeID As Long, _
    'ByVal bln���ѿ� As Boolean, _
    'ByVal strPatiName As String, ByVal strSex As String, _
    'ByVal strOld As String, ByVal dbl��� As Double, _
    'Optional ByRef strCardNo As String, _
    'Optional ByRef strPassWord As String, _
    Optional ByRef bln�˷� As Boolean = False, _
    Optional ByRef blnShowPatiInfor As Boolean = False, _
    Optional ByRef bln���� As Boolean) As Boolean
    If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, Nothing, lng�����ID, bln���ѿ�, Trim(txtPatient.Text), cboSex.Text, txt����.Text & cbo���䵥λ.Text, dbl�˷Ѷ�, strBrushCardNo, strbrPassWord, _
        True, True, bln����) = False Then Exit Function
    CheckBrushCard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 

Private Function CallBackBalanceInterface(ByVal cllBalance As Collection, _
    ByVal lng����ID As Long, ByVal lng����ID As Long, _
    ByVal strNo As String, _
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
    Dim strTemp As String
    
    Err = 0: On Error GoTo Errhand:
    '�����ID,����,�Ƿ����ѿ�(1-��;0-��),������ˮ��,����˵��,strNO
    'cllBalance.Add Array(Val(Nvl(rsTmp!�����ID)), Trim(Nvl(rsTmp!����)), IIf(Val(Nvl(rsTmp!���㿨���)) <> 0, 1, 0), Trim(Nvl(rsTmp!������ˮ��)), Trim(Nvl(rsTmp!����˵��))), strNO
    If cllBalance Is Nothing Then CallBackBalanceInterface = True: Exit Function
    
    bln���ѿ� = Val(cllBalance(1)(2)) = 1
    lng�����ID = cllBalance(1)(0)
    
    '�����ID,����,�Ƿ����ѿ�(1-��;0-��),������ˮ��,����˵��,strNO
    If lng�����ID = 0 Then CallBackBalanceInterface = True: Exit Function
    
    str���� = cllBalance(1)(1)
    strSwapGlideNO = cllBalance(1)(3)
    strSwapMemo = cllBalance(1)(4)
    If lng����ID <> 0 Then str������Ϣ = str������Ϣ & "||3|" & lng����ID
    If str������Ϣ <> "" Then str������Ϣ = Mid(str������Ϣ, 3)
'    strSQL = "Select ����ID From ������ü�¼  Where ��¼����=1 And NO =[1] and ��¼״̬=2"
'    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo)
'    If rsTemp.EOF Then
'        strErrMsg = "δ�ҵ��������㽻����Ϣ���˷�ʧ��": Exit Function
'    End If
'    lng����ID = Val(Nvl(rsTemp!����ID))
    
    '81489,Ƚ����,2015-1-22,�˷Ѵ������ID
    strSwapExtendInfor = "3|" & lng����ID: strTemp = strSwapExtendInfor
    
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
    Call zlAddUpdateSwapSQL(False, lng����ID, lng�����ID, bln���ѿ�, str����, strSwapGlideNO, strSwapMemo, cllUpdate, 2)
    If strTemp <> strSwapExtendInfor Then
        Call zlAddThreeSwapSQLToCollection(False, lng����ID, lng�����ID, bln���ѿ�, str����, strSwapExtendInfor, cllThreeSwap)
    End If
    CallBackBalanceInterface = True
Errhand:
    
End Function

Private Sub SaveThreeData(ByVal cllThree As Collection)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������������
    '����:���˺�
    '����:2011-08-18 17:53:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    zlExecuteProcedureArrAy cllThree, Me.Caption
Errhand:
    Err = 0: On Error GoTo 0
End Sub

Private Function LoadErrBillCharge(ByVal strNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ش�����շ�Ʊ��,���������շ�
    '���:strNo-������շѵ��ݺ�
    '����:
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-22 16:14:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsNos As ADODB.Recordset, strSQL As String
    Dim objPage As New BillPage
    Dim arrBills As Variant, strBills As String
    Dim blnRead As Boolean, i As Long, k As Long
    Dim lng����ID As Long, blnMulitNos As Boolean '�൥��
    Dim lng����ID As Long, lngRow As Long
    
    If Not (mbytInFun = 0 And (mbytInState = 4 Or mbytInState = 5 Or mblnErrBill)) Then LoadErrBillCharge = True: Exit Function
     
    Err = 0: On Error GoTo Errhand:
    
    strSQL = "" & _
    "   Select    A.NO, A.����ID,A.����ID,max(B.�������) as �������   " & _
    "   From ������ü�¼ A,����Ԥ����¼ B , " & _
    "           (Select max(j.�������) as ������� From ������ü�¼ m,����Ԥ����¼ j  Where m.��¼����=1 and m.��¼״̬=[2] and m.����ID=j.����ID And   m.NO=[1]) I" & _
    "   Where  A.����ID=B.����ID  " & _
    "           And B.�������=I.������� And A.��¼״̬=[2]  " & _
    "   Group by A.NO, A.����ID,A.����ID" & _
    "   Order by A.����ID " & IIf(mbln�˷��쳣, " Desc ", "") & ",A.NO"
    
    Set rsNos = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo, IIf(mbln�˷��쳣, 2, 1))
    If rsNos.RecordCount = 0 Then Exit Function
    With rsNos
        '����Ƿ����δ���ҽ������
        blnMulitNos = .RecordCount > 1
        mlng������� = Val(Nvl(!�������))
    End With
    mblnDelete = mbln�˷��쳣
    '57682
    strSQL = "" & _
    "   Select  decode(A.��¼����,1,'Ԥ���',11,'Ԥ���',A.���㷽ʽ) as ���㷽ʽ, " & _
    "            sum(nvl(A.��Ԥ��,0)) as ������ " & _
    "   From ����Ԥ����¼ A" & _
    "   where A.�������=[1]  And  A.��¼״̬=[2]  " & _
    "   Group by decode(A.��¼����,1,'Ԥ���',11,'Ԥ���',A.���㷽ʽ)" & _
    "   Order by ���㷽ʽ"
    
    '�쳣���ݵĽ��㷽ʽ(����Ԥ����)
    Set mrsErrBlance = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng�������, IIf(mbln�˷��쳣, 2, 1))
    LoadErrBillCharge = True
    
    '������е��ݵ�����
    '---------------------------------------------------------------------
     txtModi.Text = ""
    Call ClearTotalInfo
    Call ClearPayInfo
    Call ClearBillRows
        
    'Ԥ����֧��ʱ�����,������Զ���
    If cmdԤ����.Visible Then Call InitBalanceGrid
    Call ClearMoney
    Set mcolBalance = New Collection
    mcolBalance.Add Array()
    '�൥���շ�:ֻ����һҳ����
    For i = mobjBill.Pages.Count To 1 Step -1
        mobjBill.Pages.Remove i
    Next
    mobjBill.Pages.Add objPage.Details
    
    '�൥���շ�:�ָ�ȱʡ����ҳ��
    mintPage = 1
    For i = tbsBill.Tabs.Count To 1 Step -1
        tbsBill.Tabs(i).Tag = ""
        If i <> 1 Then tbsBill.Tabs.Remove i
    Next
        
    '��ȡ��ʾÿ�Ż��۵�
    '---------------------------------------------------------------------
    mblnNOMoved = False  '���Ӻ󱸱��ж�ȡ
    k = 1: i = 0
    mblnDoing = False '���������Զ���
    tbsBill.Visible = blnMulitNos
    cmdAddBill.Visible = blnMulitNos
    cmdDelBill.Visible = blnMulitNos
    cmdAddBill.Enabled = True
    fraBill.Visible = blnMulitNos
    Form_Resize
    mintInsure = 0
    Do While Not rsNos.EOF
        If mintInsure = 0 Then
                mintInsure = ChargeExistInsure(Nvl(rsNos!NO), lng����ID, lng����ID)
                If mintInsure <> 0 Then Call initInsurePara(lng����ID)
        End If
        Me.Refresh
        '���ӵ���ҳ��ǩ(ͬcmdAdd_Click����)
        '-----------------------------------------------------------------------
        If k > 1 And mobjBill.Pages(mobjBill.Pages.Count).NO <> "" Then
            If tbsBill.Tabs.Count >= 10 Then
                Call tbsBill.Tabs.Add(, , "����" & tbsBill.Tabs.Count + 1)
            Else
                If tbsBill.Tabs.Count + 1 = 10 Then
                    Call tbsBill.Tabs.Add(, , "����1&0")
                Else
                    Call tbsBill.Tabs.Add(, , "����&" & tbsBill.Tabs.Count + 1)
                End If
            End If
            
            '���뵥��ҳ����:��ʹ�ǻ����շ�Ҳ����һ��
            mobjBill.Pages.Add objPage.Details
            '������㼯��:�����շ�ҲҪ����һ��
            mcolBalance.Add Array()
            '���ŵ���ʱ��ֹ�޸ļ��˷ѹ���
            txtModi.Enabled = False
            chkCancel.Enabled = False
            cmdDelete.Enabled = False
            '����Click,��ʾ�����ӵ��ݵ�����(�հ�)
            tbsBill.Tabs(tbsBill.Tabs.Count).Selected = True
        End If
                
        '��ȡ���۵�������(ͬcboNO_KeyPress)
        '----------------------------------------------------------------------
        strNo = Nvl(rsNos!NO)
        blnRead = ReadBill(strNo, 0, False, , , True)
        mobjBill.Pages(k).����ID = Val(Nvl(rsNos!����ID))
        If blnRead Then k = k + 1: cboNO.Text = strNo
        i = i + 1
        rsNos.MoveNext
    Loop
    Dim blnFind As Boolean
    '���ؽ��㷽ʽ
    mrsErrBlance.Filter = 0
    With mrsErrBlance
        If mrsErrBlance.RecordCount <> 0 Then mrsErrBlance.MoveFirst
        vsBalance.Clear
        vsBalance.Rows = 1
        i = 1
        Do While Not .EOF
            lngRow = 0
            blnFind = False
            For i = 0 To vsBalance.Rows - 1
                If vsBalance.TextMatrix(i, 0) = Nvl(!���㷽ʽ, " ") Then
                    blnFind = True
                    lngRow = i: Exit For
                End If
            Next
            If Not blnFind And vsBalance.TextMatrix(lngRow, 0) <> "" Then
                vsBalance.Rows = vsBalance.Rows + 1
                lngRow = vsBalance.Rows - 1
            End If
             vsBalance.TextMatrix(lngRow, 0) = Nvl(!���㷽ʽ, " ")
             vsBalance.TextMatrix(lngRow, 1) = Format(Val(Nvl(!������)) + Val(vsBalance.TextMatrix(lngRow, 1)), "0.00")
            .MoveNext
        Loop
    End With
    txtInvoice.Text = ""
    Call ReInitPatiInvoice(True, mintInsure, lng����ID)
    Bill.Active = False
    chk�Ӱ�.Enabled = False
    'txtDate.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    cmdDelBill.Enabled = False
    cmdAddBill.Enabled = False
    mblnDoing = False '�����Զ���ȡ���
    Call ShowMoney
    '��ʾժҪ
    Call Bill_EnterCell(1, BillCol.��Ŀ)
    cmdOK.Enabled = True: cmdOK.Visible = True
    If cmdOK.Enabled Then cmdOK.SetFocus
    LoadErrBillCharge = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub PrintBill(ByVal strNos, strModiNos As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:Ʊ�ݴ�ӡ
    '����:���˺�
    '����:2011-08-26 18:38:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNotValiedNos As String
    Dim strReclaimInvoice As String '���յķ�Ʊ��
    
    If mbytInFun = 1 Or mblnSaveAsPrice Then  '��ӡ����֪ͨ��
        If gint����֪ͨ�� = 1 Then
           Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1120", Me, "NO=" & mobjBill.NO, 2)
        ElseIf gint����֪ͨ�� = 2 Then
            If MsgBox("Ҫ��ӡ����֪ͨ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1120", Me, "NO=" & mobjBill.NO, 2)
            End If
        End If
        Exit Sub
    End If
    
    If mbytInFun = 2 Then   '���ʵ���ӡ
        If mbytBilling = 0 And gbln���ʴ�ӡ Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1122", Me, "NO=" & mobjBill.NO, "ҩƷ��λ=" & IIf(gblnҩ����λ, 1, 0), 2)
        ElseIf mbytBilling = 1 And gbln���۴�ӡ Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1122", Me, "NO=" & mobjBill.NO, "ҩƷ��λ=" & IIf(gblnҩ����λ, 1, 0), 2)
        End If
        Exit Sub
    End If
    If mbytInFun <> 0 Then Exit Sub
    If mstrYBPati <> "" And MCPAR.���������շ� Then
        'ҽ�������շ�ģʽʱ��ȷ��ʱ����ӡ����ͬһ���˵ļ��ŵ���ȷ����󣬰�[����շ�]��ťһ���ӡ��
        'ҽ�������շ�ʱ��֧�ֶ൥��,ȡһ��������
        mstrYBBill = mstrYBBill & "," & mobjBill.NO
        Exit Sub
    End If
    
   '��ӡ�����վ�
     '����:34941:And Not (MCPAR.�൥��һ�ν��� And mstrYBPati <> "")
     Dim blnPrintBillEmpty As Boolean   '55052
    If mblnPrint And Not (MCPAR.ҽ���ӿڴ�ӡƱ�� And mstrYBPati <> "") Then
        '����:42708
        If Format(mobjBill.�Ǽ�ʱ��, "yyyy") < 2000 Then mobjBill.�Ǽ�ʱ�� = zlDatabase.Currentdate
        '����:44322
RePrint:
        strReclaimInvoice = ""
        Call frmPrint.ReportPrint(1, strNos, strModiNos, strReclaimInvoice, mlng����ID, mlngShareUseID, txtInvoice.Text, mobjBill.�Ǽ�ʱ��, CStr(mdbl�ɿ�), CStr(mdbl�Ҳ�), _
            gTy_Module_Para.bln�ֱ��ӡ And mbytBillSource <> 4, mintInvoiceFormat, , , mstrUseType, blnPrintBillEmpty, , , mstr��ͨ�۸�ȼ�)
        If gblnStrictCtrl And blnPrintBillEmpty = False Then
            If zlIsNotSucceedPrintBill(1, strNos, strNotValiedNos) = True Then
                    If MsgBox("����[" & strNotValiedNos & "]Ʊ�ݴ�ӡδ�ɹ�,�Ƿ����½���Ʊ�ݴ�ӡ!", vbYesNo + vbDefaultButton1 + vbQuestion, gstrSysName) = vbYes Then GoTo RePrint:
            End If
        End If
    End If
    '��ӡ�����嵥:�̶����ֱ��ӡ
    If InStr(mstrPrivs, "��ӡ�嵥") > 0 Then
        If gint�շ��嵥 = 1 Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_3", Me, "NO=" & strNos & IIf(strModiNos <> "", "," & strModiNos, ""), "ҩƷ��λ=" & IIf(gblnҩ����λ, 1, 0), 2)
        ElseIf gint�շ��嵥 = 2 Then
            If MsgBox("Ҫ��ӡ�շ��嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_3", Me, "NO=" & strNos & IIf(strModiNos <> "", "," & strModiNos, ""), "ҩƷ��λ=" & IIf(gblnҩ����λ, 1, 0), 2)
            End If
        End If
    End If
End Sub

Public Function ChargeIsErrBill(ByVal strNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鵥���Ƿ��쳣����
    '����:���쳣����,������True,���򷵻�False
    '����:���˺�
    '����:2011-08-28 11:32:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errH
    strSQL = "Select /*+cardinality(j,10)*/ 1" & vbNewLine & _
            " From ������ü�¼ A, Table(f_Str2list([1])) J" & vbNewLine & _
            " Where a.��¼���� = 1 And Nvl(a.����״̬, 0) = 1 And a.No = j.Column_Value And Rownum < 2"
            
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "����Ƿ��쳣����", Replace(strNos, "'", ""))
    ChargeIsErrBill = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function PatiErrBillPay(ByVal lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݲ���,���쳣���ݽ����շ�
    '����:�����쳣����,�������շ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-10-23 21:04:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim strNo As String, lng����ID As Long, lng������� As Long
    Dim str����Ա���� As String
    '����:44271
    If (mbytInState = 1 Or mbytInState = 2 Or mbytInState = 3) Or mbytInFun <> 0 Then Exit Function
    If mbytInFun = 0 And mbytInState = 0 And mstrInNO <> "" Then PatiErrBillPay = False: Exit Function
   
    On Error GoTo errHandle
    
    mblnErrBill = False
    strSQL = "" & _
    "   Select  A.NO,A.����ID,A.����Ա����" & _
    "   From ������ü�¼ A" & _
    "   Where A.����ID=[1] and nvl(A.����״̬,0)=1  " & _
    "               And A.��¼״̬=1 and A.��¼����=1 " & _
    "               And Rownum=1 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    If rsTemp.EOF Then Exit Function
    strNo = Nvl(rsTemp!NO): lng����ID = Val(Nvl(rsTemp!����ID))
    
    If isCheckExiseSingularity(strNo) Then         '�����ϵ�
        Exit Function
    End If
    
    str����Ա���� = Nvl(rsTemp!����Ա����)
    If str����Ա���� <> UserInfo.���� Then
        strSQL = "" & _
        "   Select  ������� From ����Ԥ����¼  A" & _
        "   Where ����ID=[1]  " & _
        "               And exists(Select 1 From ���㷽ʽ B Where nvl(A.���㷽ʽ,'-')=b.���� and ���� not in ('3','4') )"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
        If Not rsTemp.EOF Then

            '�����������㷽ʽ
           Call MsgBox("ע��:" & vbCrLf & _
            "       �ò��˴����쳣���շѵ���,������Ա " & vbCrLf & _
            "       [" & str����Ա���� & "]��ȡ��һ����, " & _
            "       �뵽�ò���Ա�������շ�!", vbOKOnly + vbInformation, gstrSysName)
            PatiErrBillPay = True
            Exit Function
        End If


    End If
    If MsgBox("ע��:" & vbCrLf & _
                        "       �ò��˴����쳣���շѵ���" & IIf(str����Ա���� <> UserInfo.����, ",�õ����ǲ���Ա[" & str����Ա���� & "]��ȡ��," & vbCrLf, "") & " ,�Ƿ����¶Ըõ��ݽ����շ�?" & vbCrLf & vbCrLf & _
                        "���ǡ��������¶��쳣�����շ� " & vbCrLf & _
                        "���񡻴������쳣���ݽ��д���,�������շ�.", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
        Exit Function
    End If
    PatiErrBillPay = True
    mblnErrBill = True
    If LoadErrBillCharge(strNo) = False Then Exit Function
    
    '�������
    If zlIsCheckExistErrBill(mlng�������) = False Then
        MsgBox "��ǰ�쳣�����ѱ������㲻�ܼ�����", vbInformation, gstrSysName
        Exit Function
    End If
    If zlCheckOtherSessionDoing(mlng�������) Then
        MsgBox "��ǰ�������������շѴ����н��д����㲻�ܼ�����", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���ⲿ�ֵ��ݸ���Ϊ��ǰ����Ա
    'Zl_�����쳣�շ�_���²���Ա
    gstrSQL = "Zl_�����쳣�շ�_���²���Ա("
    '����id_In     ������ü�¼.����id%Type,
    gstrSQL = gstrSQL & "" & lng����ID & ","
    '����Ա���_In ������ü�¼.����Ա���%Type,
    gstrSQL = gstrSQL & "'" & UserInfo.��� & "',"
    '����Ա����_In ������ü�¼.����Ա����%Type,
    gstrSQL = gstrSQL & "'" & UserInfo.���� & "',"
    '�������_In   ����Ԥ����¼.�������%Type
    gstrSQL = gstrSQL & IIf(lng������� = 0, mlng�������, lng�������) & ")"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    If ReChargeFee = False Then Exit Function
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function LoadCurBalance()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ص�ǰ������Ϣ
    '���:strBalance:�����շѵĽ��㷽ʽ,��ʽ����:
    '        ���:�ɿ��־(1-�ɿ�;2-�Ҳ�)|���㷽ʽ1:���1:�ɿ��־(1-�ɿ�;2-�Ҳ�)|...
    '���أ����������շѵ��ܶ�
    '����:���˺�
    '����:2011-11-02 13:27:04
    '����:42791
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long
    Dim int���� As Integer
    
    Call InitBalanceGrid
    If grsTotal Is Nothing Then Exit Function
    If grsTotal.State <> 1 Then Exit Function
    
    With vsBalance
        '����:0-�ɿ�;1-�Ҳ�,2-��Ԥ��;����(mod 10:0-��ͨ����;1-ҽ������;2-������Ʒ;3-һ��ͨ)
        grsTotal.Sort = "����"
        .Rows = IIf(.Rows >= grsTotal.RecordCount, .Rows, grsTotal.RecordCount)
        lngRow = 0
        Do While Not grsTotal.EOF
            '���� ,���㷽ʽ  ������
            '��frmChargePayMentWin-����,��Ҫ��һЩ�ۼ���
            .TextMatrix(lngRow, 0) = Nvl(grsTotal!���㷽ʽ)
            .TextMatrix(lngRow, 1) = Format(Val(Nvl(grsTotal!������)), "###0.00;-###0.00;0.00;0.00")
             int���� = Val(Nvl(grsTotal!����))
            .Cell(flexcpForeColor, lngRow, 0, lngRow, .COLS - 1) = Me.ForeColor
            If int���� = 0 Then
                .Cell(flexcpFontBold, lngRow, 0, lngRow, .COLS - 1) = True
            ElseIf int���� = 1 Then
                .Cell(flexcpFontBold, lngRow, 0, lngRow, .COLS - 1) = True
                'If Val(.TextMatrix(lngRow, 1)) < 0 Then '45416
                    .Cell(flexcpForeColor, lngRow, 0, lngRow, .COLS - 1) = vbRed
               ' End If
            End If
            lngRow = lngRow + 1
            grsTotal.MoveNext
        Loop
    End With
End Function

Private Function ModifyNotInsureNOs(ByVal strNotSucceedNo As String, _
    ByVal strSucceedNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�޸�δ���óɹ���ҽ������
    '���:strNotSucceedNo-ҽ�����㲻�ɹ��ĵ���
    '        strSucceedNos-ҽ������ɹ��ĵ���
    '        blnErrReChager-�쳣���������շ�
    '����:�޸ĳɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-12-17 22:37:04
    '����:44535
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strInfor As String
    Dim varNos As Variant, varNotNOs As Variant
    Dim intNum As Integer, intNotNum As Integer
    Dim intType As Integer
    If strNotSucceedNo = "" Then Exit Function
    varNos = Split(strSucceedNos, ","): varNotNOs = Split(strNotSucceedNo, ",")
    If strSucceedNos <> "" Then intNum = UBound(varNos) + 1
    If strNotSucceedNo <> "" Then intNotNum = UBound(varNotNOs) + 1
    intType = 0
    If intNum <> 0 Then
        strInfor = "ҽ���ɹ�����" & intNum & "��" & vbCrLf & _
        "    " & strSucceedNos & vbCrLf
    End If
    strInfor = strInfor & "" & _
    "ҽ���ǳɹ�����" & intNotNum & "��" & vbCrLf & _
    "    " & strNotSucceedNo & vbCrLf

    If intNum = 0 Then
        strInfor = strInfor & "" & _
        "���ܽ���ҽ������!"
        Call MsgBox(strInfor, vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName)
        Exit Function
        intType = 1
    Else
       strInfor = strInfor & "" & _
        "Ŀǰֻ�ܶԳɹ����ײ��ֽ����շ�!"
    End If
    Call MsgBox(strInfor, vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName)
       
    On Error GoTo errHandle
    'Zl_ҽ���շ��쳣_Update
    strSQL = "Zl_ҽ���շ��쳣_Update("
    '  Nos_In          Varchar2,
    strSQL = strSQL & "'" & strNotSucceedNo & "',"
    '  ���½��㷽ʽ_In Integer:=0
    strSQL = strSQL & "" & intType & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption

    ModifyNotInsureNOs = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub ClearDisplaySHow()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���˫����ʾ
    '����:���˺�
    '����:2011-12-29 09:54:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '˫����ʾ��������ڵ�ǰ������ʾ֮�������ʾ�����ƶ�����
    If Not gblnLED Then Exit Sub
    If Not (mbytInFun = 0 And mbytInState = 0) Then Exit Sub
    If mblnNotClearLedDisplay Then Exit Sub
    zl9LedVoice.DisplayPatient ""
End Sub
Private Function SaveChargeBill(ByRef lng������� As Long, ByRef strBalanceIDs As String, ByRef strSaveNos As String, _
    Optional ByRef strModiNos As String, _
    Optional ByRef blnSaveBill As Boolean, _
    Optional ByRef blnNotCommit As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���浱ǰ����ĵ���(�������շ� )
    '����:lng�������-���ر��α��浥�ݵĽ������
    '       strBalanceIDs-���سɹ����ʵĽ���IDs( ��������ҽ���ӿڻ���������ݵ�����)
    '       strSaveNos-�����ѳɹ�����ĵ��ݺţ���ʽΪ"'AAA','BBB',..."
    '       strModiNOs -�޸ĵ��Ƕ൥���շ��е�һ��ʱ�����ظö��ŵ��ݵ�����NO����ʽ��"'AAA','BBB',..."
    '       blnSaveBill-�Ƿ񵥾��Ѿ�����ɹ�
    '       blnNotCommit-�����������ύ����Ҫ�Ǵ�����ͨ�����շ�(����һ��ͨ���㣬����ҽ���������������ٳ����쳣�ĳ���)
    '����:�շѳɹ��򵥾ݱ���湦,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-26 17:28:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '     *** ҽ���շ�ʱ,����ʱ����Ϊ���۵�,�ڽ���ǰ��תΪ�շѵ�,�Ա������ҩƷ���ʱ��ȴ�ͬһ�����ҽ��������������� ***
    Dim lng��ӡID As Long, lngҽ��ID As Long, lng���ͺ� As Long
    Dim int��� As Integer, int�۸񸸺� As Integer, int�к� As Integer
    Dim lng����ID As Long, intҩƷ�д� As Integer, strҽ�Ƹ��� As String
    Dim dbl���� As Double, dbl���� As Double, cur�ɿ� As Currency
    Dim strDeptIDs As String, strTmp As String, strDelBill As String, strBillNO As String
    Dim str�շѽ��� As String, str���ս��� As String, str�շѽ���У�� As String
    Dim arrSQL As Variant, arrPut As Variant, arrOTMSQL As Variant
    Dim blnֱ���շ� As Boolean, blnTrans As Boolean, blnTransMedicare As Boolean
    Dim i As Integer, j As Integer, p As Integer, strSQL As String
    Dim CurOneCard As Currency, dblOneCardBalance As Double
    Dim strCardNo  As String, intCardType As Integer, strTransFlow As String
    Dim str��ҩ��̬ As String
    Dim strStuffDept As String          '�Զ����ϵĲ���
    Dim strAdvance As String            'ҽ�����㷵�ص���Ϣ:"���㷽ʽ|������||....."
    Dim blnPriceSaved As Boolean        'ҽ���շ�ʱ�Ƿ��Ѵ�Ϊ���۵�,������תΪ�շѵ���ҽ����������ʧ�ܻ��˺�ɾ�����۵�
    Dim blnMedicareCheck As Boolean     '�Ƿ�ִ��ҽ������У��
    Dim strInvoice As String            '��ǰ����ʹ�õ�Ʊ�ݺţ�����ҽ��һ�ŵ���ֻ��һ��Ʊ�����
    Dim cllRqure As Collection
    Dim rsSqure As ADODB.Recordset
    Dim str����IDs As String
    Dim blnӦ���� As Boolean
    Dim dblӦ�ɶ� As Double, lng������� As Long
    Dim cllPutout As Collection '�Զ�����
    Dim cllYBCommit As Collection   'SQL,key(���ݺ�)
    Dim cllPro As Collection, cllDelete As Collection, cllPageInfor As Collection
    Dim cur�ѽɺϼ� As Currency
    Dim strSaveCuessNos As String
    
    'ֻ�����շѵ�
    If mblnSaveAsPrice Then Exit Function
    If mbytInFun <> 0 Then Exit Function
    '�µķ�ҩ��Ʒ��(Ŀǰֻ���ֹ�¼����Ч)
    Set mCllWindows = New Collection
    
    strBalanceIDs = ""
    strSaveNos = "": cur�ѽɺϼ� = 0: strModiNos = ""
    Err = 0: On Error GoTo Errhand:
    If cboҽ�Ƹ���.ListIndex <> -1 Then
        strҽ�Ƹ��� = Mid(cboҽ�Ƹ���.Text, 1, InStr(1, cboҽ�Ƹ���, "-") - 1)
    End If
    mobjBill.����ʱ�� = CDate(txtDate.Text)
    mobjBill.�Ǽ�ʱ�� = zlDatabase.Currentdate
    strInvoice = Trim(txtInvoice.Text)
    arrOTMSQL = Array()
    
    '�޸Ĺ���ʱ,�Ƿ��޸�ҽ������
    If mstrInNO <> "" Then
        Call BillisAdviceMoney(mstrInNO, 1, lngҽ��ID, lng���ͺ�)
    End If
    
    blnSaveBill = False
    dblӦ�ɶ� = 0: lng������� = 0
     Set cllPro = New Collection
    Set cllPageInfor = New Collection
    Set mcllPayDrugAndStuff = New Collection
     lng������� = 0
    
    '��ÿ�ŵ��ݶ���ִ�б���
    For p = 1 To mobjBill.Pages.Count
        int��� = 0: int�к� = 0: blnPriceSaved = False
        intҩƷ�д� = 0: strDeptIDs = "": strStuffDept = ""
        '��ǰ�շѵ��ݵĸ������
        If mbytInFun = 0 And Not mblnSaveAsPrice Then
            str���ս��� = GetMedicareStr(p)
            lng����ID = zlDatabase.GetNextId("���˽��ʼ�¼")
            If lng������� = 0 Then lng������� = lng����ID
            str����IDs = str����IDs & "," & lng����ID
        End If
        
        '����ÿ���շѵ��ݵĵ��ݺ�
        blnֱ���շ� = False
        strBillNO = mobjBill.Pages(p).NO
        If mobjBill.Pages(p).NO = "" Then
            'Ϊ����ʧ�ܺ�����ʶ��,���Ķ���NO
            strBillNO = zlDatabase.GetNextNo(13)    '�շѵ�
            blnֱ���շ� = True
        End If
        If p = 1 Then
            mobjBill.NO = strBillNO: gstrModiNO = strBillNO
        End If
        
        '��ҪΪ��Ϣ������,Ϊÿҳ����ĵ��ݺ�
        mobjBill.Pages(p).�շѵ��� = strBillNO
        
        
        arrSQL = Array() '�൥��ʱ,���ŵ����ύ
        If Not blnֱ���շ� Then
            '1.�շ��µ��ݹ���ʱ,��ȡ�Ļ��۵��շ�
            '��ȻZl_���˻����շ�_Insertû�и���ҽ����Ϣ,���ڸ��ݲ�����ȡ�Ļ��۵�ʱִ����zl_���ﻮ�ۼ�¼_Update,�Ѹ���
            '---------------------------------------------------------------
            'Zl_���˻����շ�_Insert
           gstrSQL = "Zl_���˻����շ�_Insert("
            '  No_In         ������ü�¼.NO%Type,
            gstrSQL = gstrSQL & "'" & strBillNO & "',"
            '  ����id_In     ������ü�¼.����id%Type,
            gstrSQL = gstrSQL & "" & ZVal(mobjBill.����ID) & ","
            '  ������Դ_In   Number,
            gstrSQL = gstrSQL & "" & gint������Դ & ","
            '  ���ʽ_In   ������ü�¼.���ʽ%Type,
            gstrSQL = gstrSQL & "'" & strҽ�Ƹ��� & "',"
            '  ����_In       ������ü�¼.����%Type,
            gstrSQL = gstrSQL & "'" & mobjBill.���� & "',"
            '  �Ա�_In       ������ü�¼.�Ա�%Type,
            gstrSQL = gstrSQL & "'" & mobjBill.�Ա� & "',"
            '  ����_In       ������ü�¼.����%Type,
            gstrSQL = gstrSQL & "'" & mobjBill.���� & "',"
            '  ���˿���id_In ������ü�¼.���˿���id%Type,
            gstrSQL = gstrSQL & "" & IIf(mobjBill.Pages(p).ҽ����� > 0, "Null", ZVal(mobjBill.����ID, , mobjBill.Pages(p).��������ID)) & ","
            '  ��������id_In ������ü�¼.��������id%Type,
            gstrSQL = gstrSQL & "" & ZVal(mobjBill.Pages(p).��������ID) & ","
            '  ������_In     ������ü�¼.������%Type,
            gstrSQL = gstrSQL & "'" & mobjBill.Pages(p).������ & "',"
            '  ���ս���_In   Varchar2,
            If mstrYBPati <> "" And str���ս��� <> "" Then
                gstrSQL = gstrSQL & "'" & str���ս��� & "',"
            Else
                gstrSQL = gstrSQL & "NULL,"
            End If
            '  ����id_In     ������ü�¼.����id%Type,
            gstrSQL = gstrSQL & "" & lng����ID & ","
            '  ����ʱ��_In   ������ü�¼.����ʱ��%Type,
            gstrSQL = gstrSQL & "To_Date('" & Format(mobjBill.����ʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
            '  ����Ա���_In ������ü�¼.����Ա���%Type,
            gstrSQL = gstrSQL & "'" & UserInfo.��� & "',"
            '  ����Ա����_In ������ü�¼.����Ա����%Type,
            gstrSQL = gstrSQL & "'" & UserInfo.���� & "',"
            '  ��ҩ����_In   ������ü�¼.��ҩ����%Type := Null,
            gstrSQL = gstrSQL & "'" & tbsBill.Tabs(p).Tag & "',"
            '  �Ƿ���_In   ������ü�¼.�Ƿ���%Type := 0,
            gstrSQL = gstrSQL & "" & chk����.Value & ","
            '  �Ǽ�ʱ��_In   ������ü�¼.�Ǽ�ʱ��%Type := Null,
            gstrSQL = gstrSQL & "" & "NULL" & ","
            '  �������_In   ����Ԥ����¼.�������%Type := Null
            gstrSQL = gstrSQL & "" & lng������� & ")"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "0;" & gstrSQL
                        
            '��ȡ�Զ���ҩ�Ķ��ҩ��
            If gbln�շѺ��Զ���ҩ Then
                strDeptIDs = strDeptIDs & "," & Get��ҩ����IDs(strBillNO)
            End If
            '���ÿ�ŵ����ռ����Ϸ��ϲ���,�Ա��Զ�����,�Ƿ��Ǹ������ò�����SQL���ж�
            If gbln�����Զ����� Then
                strStuffDept = strStuffDept & "," & Get��ҩ����IDs(strBillNO, "'4'")
            End If
            'ͨ�����۵��շѵķ�ʽ��ȡ�˹Һŷ����ķ���,����ɾ���÷���
            If strBillNO = mstrCardNO Then mstrCardNO = ""
        ElseIf blnֱ���շ� Then
            '2.ֱ������ĵ�������,�����������޸�,�������շ�(���շѽ��汣��Ϊ���۵�),����,����
            '---------------------------------------------------------------
            For Each mobjBillDetail In mobjBill.Pages(p).Details
                If mobjBillDetail.���� <> 0 Then
                    For Each mobjBillIncome In mobjBillDetail.InComes
                        int��� = int��� + 1 '��ǰ��¼���
                        '1.��������---------------------------------------------------------------
                        With mobjBill                              'ҽ���շ�ʱ,����ʱ����Ϊ���۵�,�ڽ���ǰ��תΪ�շѵ�
                            If mstrYBPati = "" Then
                                gstrSQL = "zl_���������շ�_INSERT('" & strBillNO & "'," & int��� & "," & ZVal(.����ID) & "," & _
                                    IIf(gint������Դ = 2, 2, 1) & "," & ZVal(.��ʶ��) & ",'" & strҽ�Ƹ��� & "'," & _
                                    "'" & .���� & "','" & .�Ա� & "','" & .���� & "','" & IIf(mobjBillDetail.�ѱ� = "", .�ѱ�, mobjBillDetail.�ѱ�) & "'," & _
                                    .�Ӱ��־ & "," & ZVal(.����ID, , .Pages(p).��������ID) & "," & _
                                    ZVal(.Pages(p).��������ID) & ",'" & .Pages(p).������ & "',"
                            Else
                                gstrSQL = "zl_���ﻮ�ۼ�¼_INSERT('" & strBillNO & "'," & int��� & "," & ZVal(.����ID) & "," & _
                                    ZVal(.��ҳID) & "," & ZVal(.��ʶ��) & ",'" & strҽ�Ƹ��� & "'," & _
                                    "'" & .���� & "','" & .�Ա� & "','" & .���� & "','" & IIf(mobjBillDetail.�ѱ� = "", .�ѱ�, mobjBillDetail.�ѱ�) & "'," & _
                                    .�Ӱ��־ & "," & ZVal(.����ID, , .Pages(p).��������ID) & "," & _
                                    ZVal(.Pages(p).��������ID) & ",'" & .Pages(p).������ & "',"
                            End If
                            
                        End With
        
                        '2.�շ�ϸĿ����---------------------------------------------------------------
                        With mobjBillDetail
                            If .��� <> int�к� Then     '�����������
                                int�к� = .���
                                int�۸񸸺� = int���
                                '���´����������
                                If mobjBill.Pages(p).Details(.���).�������� = 0 Then
                                    For i = .��� + 1 To mobjBill.Pages(p).Details.Count
                                        If mobjBill.Pages(p).Details(i).�������� = .��� Then
                                            '������Ŀ�ж��������Ŀ(������)ʱ,ȡ��һ�����
                                            mobjBill.Pages(p).Details(i).�������� = int���
                                        End If
                                    Next
                                End If
                            End If
        
                            If Not Set��ҩ����(p, mobjBillDetail) Then
                                Exit Function
                            End If
                            
                            'ҽ��ֱ���շ�ʱ,��Ϊ���ݴ�Ϊ���۵�,�շ�ʱ��Ҫȡ��ҩ����
                            If mstrYBPati <> "" Then tbsBill.Tabs(p).Tag = .��ҩ����
                            dbl���� = .����
                            If InStr(",5,6,7,", .�շ����) > 0 And gblnҩ����λ Then
                                dbl���� = Format(.���� * .Detail.ҩ����װ, "0.00000")
                            End If
                            
                            gstrSQL = gstrSQL & .�������� & "," & .�շ�ϸĿID & ",'" & .�շ���� & "','" & .���㵥λ & "',"
                            If mstrYBPati = "" Then
                                gstrSQL = gstrSQL & IIf(.������Ŀ��, 1, 0) & "," & ZVal(.���մ���ID) & ",'" & .��ҩ���� & "'," & IIf(.���� = 0, 1, .����) & "," & dbl���� & "," & IIf(.������, 8, .���ӱ�־) & ","
                            Else
                                gstrSQL = gstrSQL & "'" & .��ҩ���� & "'," & IIf(.���� = 0, 1, .����) & "," & dbl���� & "," & .���ӱ�־ & ","
                            End If
                            gstrSQL = gstrSQL & IIf(.ִ�в���ID = 0, "NULL", .ִ�в���ID) & ","
                        End With
        
                        '3.������Ŀ����---------------------------------------------------------------
                        With mobjBillIncome
                            dbl���� = .��׼����
                            If InStr(",5,6,7,", mobjBillDetail.�շ����) > 0 And gblnҩ����λ Then
                                dbl���� = Format(.��׼���� / mobjBillDetail.Detail.ҩ����װ, gstrFeePrecisionFmt)
                            End If
                            gstrSQL = gstrSQL & IIf(int�۸񸸺� = int���, "NULL", int�۸񸸺�) & "," & .������ĿID & "," & _
                                    "'" & .�վݷ�Ŀ & "'," & dbl���� & "," & .Ӧ�ս�� & "," & .ʵ�ս�� & ","
                            If mbytInFun = 0 And Not mblnSaveAsPrice And mstrYBPati = "" Then
                                gstrSQL = gstrSQL & "NULL,"
                            End If
                        End With
        
                        '4.��������
                        '---------------------------------------------------------------
                        gstrSQL = gstrSQL & _
                                "To_Date('" & Format(mobjBill.����ʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                                "To_Date('" & Format(mobjBill.�Ǽ�ʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),'" & mstrInNO & "',"
                        If mobjBillDetail.�շ���� = "7" Then
                            str��ҩ��̬ = "'" & mobjBillDetail.Detail.��ҩ��̬ & "'"
                        Else
                            str��ҩ��̬ = "NULL"
                        End If
                        '��ҩ��̬_In       סԺ���ü�¼.����%Type := Null
                        
                        If mstrYBPati = "" Then
                            '��ҽ���շ�,���Ҳ��ǻ���
                            gstrSQL = gstrSQL & lng����ID & "," & lng������� & ","
                            '�������ID
                            gstrSQL = gstrSQL & "'" & UserInfo.��� & "','" & UserInfo.���� & "'," & _
                                "'" & mobjBillDetail.ժҪ & "'," & chk����.Value & ",'|" & mobjBill.Pages(mintPage).�巨 & "'" & _
                                "," & str��ҩ��̬ & ")"
                                'ֻ�ڵ�һ�ŵ��ݵĵ�һ����¼ʱ����
                        Else
                            '���ﻮ��,�շѹ��ܻ���,ҽ���շ�
                            gstrSQL = gstrSQL & "'" & UserInfo.���� & "'," & _
                                "'" & mobjBillDetail.ժҪ & "'," & ZVal(lngҽ��ID) & ",NULL,NULL,'|" & mobjBill.Pages(mintPage).�巨 & _
                                "',NULL,NULL," & gint������Դ & ",'" & mobjBillDetail.���ձ��� & "'," & _
                                "'" & mobjBillDetail.Detail.���� & "'," & IIf(mobjBillDetail.������Ŀ��, 1, 0) & "," & ZVal(mobjBillDetail.���մ���ID) & "," & _
                                str��ҩ��̬ & ")"
                        End If
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = mobjBillDetail.�շ�ϸĿID & ";" & gstrSQL
                    Next    'ÿһ��������Ŀ
                    
                    '��ÿһ���շѼ�¼�ռ�ҩƷִ�в���,������ʻ��۵�����˲���,��Oracle�����д���:zl_������ʼ�¼_Verify
                    '----------------------------------------------------------------------------------------------------------------
                    '�Զ���ҩ,���շ�ʱ�Ҳ��Ƿ��뷢ҩʱ                    '
                    With mobjBillDetail
                        If gbln�շѺ��Զ���ҩ And mbytInFun = 0 And Not mblnSaveAsPrice Then
                            If .ִ�в���ID <> 0 And InStr("5,6,7", .�շ����) > 0 Then
                                If InStr(strDeptIDs & ",", "," & .ִ�в���ID & ",") = 0 Then
                                    strDeptIDs = strDeptIDs & "," & .ִ�в���ID
                                End If
                            End If
                        End If
                        '�Զ�����,�շ��Ҳ��Ǳ���Ϊ���۵������������,���뷢ҩ������Ӱ������
                        If gbln�����Զ����� And ((mbytInFun = 0 And Not mblnSaveAsPrice) Or (mbytInFun = 2 And mbytBilling = 0)) Then
                                If .ִ�в���ID <> 0 And .�շ���� = "4" And .Detail.�������� Then
                                    If InStr(strStuffDept & ",", "," & .ִ�в���ID & ",") = 0 Then
                                        strStuffDept = strStuffDept & "," & .ִ�в���ID
                                    End If
                                End If
                        End If
                    End With
                End If
            Next            'ÿһ���շ���Ŀ
            '����ǰһ�ŵ��ݵ�ҩ��ID,�Ա���ŵ���ʱȷ����ҩ����
            If mobjBill.Pages.Count > 1 Then Call SaveDrugID(p)
            '�޸ĺ��˳�ԭ����(�޸Ķ��շѵ��е�һ��ʱ��Ҫ���˷���ͳһ�ش�)
            '--------------------------------------------------------------------------------------------------------
            If mstrInNO <> "" Then
                strDelBill = ""
                '�޸�ҽ���շѵ�,��ȻΪ������ȫ��,��Ϊ�޸ĵ���ʱ���ж����������ȫ��,�������޸�
                strDelBill = "zl_�����շѼ�¼_DELETE('" & mstrInNO & "','" & UserInfo.��� & "','" & UserInfo.���� & "'," & _
                    "NULL,NULL,'" & zlStr.NeedName(cbo���㷽ʽ.Text) & "',0,To_Date('" & Format(mobjBill.�Ǽ�ʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
                '����Ƕ൥���շ��е�һ�����µ���������ԭ���ݵĴ�ӡID��,�Ա�һ���ش�
                strTmp = GetMultiNOs(mstrInNO, lng��ӡID)
                If UBound(Split(strTmp, ",")) = 0 Then
                    lng��ӡID = 0: strModiNos = ""
                ElseIf lng��ӡID <> 0 Then
                    strModiNos = strTmp
                End If
                '������޸�ҽ���ĸ���,���µ�NO���ڸ�����
                If lngҽ��ID <> 0 And lng���ͺ� <> 0 Then
                    gstrSQL = "ZL_����ҽ������_Insert(" & lngҽ��ID & "," & lng���ͺ� & "," & IIf(mbytInFun = 2, 2, 1) & ",'" & strBillNO & "')"
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "0;" & gstrSQL
                End If
            End If
        End If
        
        '�շѺ��Զ���ҩ,���ʲ��Զ���ҩ,�շ��Ҳ��Ǳ���Ϊ���۵�,�����������
        '-----------------------------------------------------------------------
        If strDeptIDs <> "" Then
            arrPut = Array()
            strDeptIDs = Mid(strDeptIDs, 2)
            For i = 0 To UBound(Split(strDeptIDs, ","))
                ReDim Preserve arrPut(UBound(arrPut) + 1)
                arrPut(UBound(arrPut)) = "ZL_ҩƷ�շ���¼_������ҩ(" & Val(Split(strDeptIDs, ",")(i)) & ",8,'" & strBillNO & "','" & UserInfo.���� & "','" & UserInfo.���� & "','" & mobjBill.Pages(p).������ & "')"
            Next
        End If
        '�շѺ��Զ�����,���շ�(ֱ���շ�,���۵������շ�),�������ʱִ��
        If strStuffDept <> "" Then
            If strDeptIDs = "" Then arrPut = Array()
            strStuffDept = Mid(strStuffDept, 2)
            For i = 0 To UBound(Split(strStuffDept, ","))          '24-�շѴ������ϣ�25-���ʵ���������
                ReDim Preserve arrPut(UBound(arrPut) + 1)
                arrPut(UBound(arrPut)) = "zl_�����շ���¼_��������(" & Split(strStuffDept, ",")(i) & "," & IIf(mbytInFun = 0, 24, 25) & ",'" & strBillNO & "','" & UserInfo.���� & "','" & UserInfo.���� & "','" & UserInfo.���� & "',1,Sysdate)"
            Next
        End If
        'ִ�����SQL��估�ύҽ������,���ŵ���ʱ,ÿ�ŵ����ڶ����������ύ
        '--------------------------------------------------------------------------------------------------------------------------------
        If UBound(arrSQL) >= 0 Then
            '��SQL���а��շ�ϸĿID����
            For i = 0 To UBound(arrSQL) - 1
                For j = i + 1 To UBound(arrSQL)
                    If CLng(Split(arrSQL(j), ";")(0)) < CLng(Split(arrSQL(i), ";")(0)) Then
                        strTmp = CStr(arrSQL(j)): arrSQL(j) = arrSQL(i): arrSQL(i) = strTmp
                    End If
                Next
            Next
            
            'ҽ��ֱ���շ�ʱ,�ȱ���Ϊ���۵�,��תΪ�շѵ�
            '-------------------------------------------------------------------
            If blnֱ���շ� And mstrYBPati <> "" Then
                '1.�ȱ��滮�۵�,���ύ�������Ա㲻����
                If MCPAR.�൥�ݵ�һ�ν��� Or MCPAR.�൥��һ�ν��� Then
                    For i = 0 To UBound(arrSQL)
                        zlAddArray cllPro, Mid(arrSQL(i), InStr(arrSQL(i), ";") + 1)
                    Next
                Else
                    On Error GoTo errH
                    gcnOracle.BeginTrans
                        blnTrans = True
                        For i = 0 To UBound(arrSQL)
                            Call zlDatabase.ExecuteProcedure(Mid(arrSQL(i), InStr(arrSQL(i), ";") + 1), Me.Caption)
                        Next
                    gcnOracle.CommitTrans
                    blnTrans = False: blnPriceSaved = True
                End If
                
                '���»��۵��ı�����Ϣ(������Ŀ��,ҽ������ID,ͳ����)
                gstrSQL = "zl_���ﻮ�ۼ�¼_Update(" & mintInsure & "," & mobjBill.����ID & ",'" & strBillNO & "',0)"
                If MCPAR.�൥�ݵ�һ�ν��� Or MCPAR.�൥��һ�ν��� Then
                    zlAddArray cllPro, gstrSQL
                Else
                   Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                End If
                
                '���۵�תΪ�շѵ�
                 'Zl_���˻����շ�_Insert
                gstrSQL = "Zl_���˻����շ�_Insert("
                '  No_In         ������ü�¼.NO%Type,
                gstrSQL = gstrSQL & "'" & strBillNO & "',"
                '  ����id_In     ������ü�¼.����id%Type,
                gstrSQL = gstrSQL & "" & mobjBill.����ID & ","
                '  ������Դ_In   Number,
                gstrSQL = gstrSQL & "" & gint������Դ & ","
                '  ���ʽ_In   ������ü�¼.���ʽ%Type,
                gstrSQL = gstrSQL & "'" & strҽ�Ƹ��� & "',"
                '  ����_In       ������ü�¼.����%Type,
                gstrSQL = gstrSQL & "'" & mobjBill.���� & "',"
                '  �Ա�_In       ������ü�¼.�Ա�%Type,
                gstrSQL = gstrSQL & "'" & mobjBill.�Ա� & "',"
                '  ����_In       ������ü�¼.����%Type,
                gstrSQL = gstrSQL & "'" & mobjBill.���� & "',"
                '  ���˿���id_In ������ü�¼.���˿���id%Type,
                gstrSQL = gstrSQL & "" & ZVal(mobjBill.����ID, , mobjBill.Pages(p).��������ID) & ","
                '  ��������id_In ������ü�¼.��������id%Type,
                gstrSQL = gstrSQL & "" & ZVal(mobjBill.Pages(p).��������ID) & ","
                '  ������_In     ������ü�¼.������%Type,
                gstrSQL = gstrSQL & "'" & mobjBill.Pages(p).������ & "',"
                '  ���ս���_In   Varchar2,
                gstrSQL = gstrSQL & "" & IIf(str���ս��� <> "", "'" & str���ս��� & "'", "NULL") & ","
                '  ����id_In     ������ü�¼.����id%Type,
                gstrSQL = gstrSQL & "" & lng����ID & ","
                '  ����ʱ��_In   ������ü�¼.����ʱ��%Type,
                gstrSQL = gstrSQL & "To_Date('" & Format(mobjBill.����ʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
                '  ����Ա���_In ������ü�¼.����Ա���%Type,
                gstrSQL = gstrSQL & "'" & UserInfo.��� & "',"
                '  ����Ա����_In ������ü�¼.����Ա����%Type,
                gstrSQL = gstrSQL & "'" & UserInfo.���� & "',"
                '  ��ҩ����_In   ������ü�¼.��ҩ����%Type := Null,
                gstrSQL = gstrSQL & "'" & tbsBill.Tabs(p).Tag & "',"
                '  �Ƿ���_In   ������ü�¼.�Ƿ���%Type := 0,
                gstrSQL = gstrSQL & "" & chk����.Value & ","
                '  �Ǽ�ʱ��_In   ������ü�¼.�Ǽ�ʱ��%Type := Null,
                gstrSQL = gstrSQL & "" & "NULL" & ","
                '  �������_In   ����Ԥ����¼.�������%Type := Null
                gstrSQL = gstrSQL & "" & lng������� & ")"
            End If
            
            'ҽ���൥��һ�ν���ʱ�����е�����Ϊһ�������ύ
            If (MCPAR.�൥��һ�ν��� Or MCPAR.�൥�ݵ�һ�ν���) And mstrYBPati <> "" And strDelBill = "" Then
                '1.���۵�ת�շ�
                zlAddArray cllPro, gstrSQL
                '2.������
                If mobjBill.Pages(p).����� <> 0 Then '44657
                    gstrSQL = "zl_�����շ����_Insert('" & strBillNO & "'," & mobjBill.Pages(p).����� & ",0,1)"
                    zlAddArray cllPro, gstrSQL
                End If
                '3.�շѺ��Զ���ҩ,�Զ�����
                If strDeptIDs <> "" Or strStuffDept <> "" Then
                    For i = 0 To UBound(arrPut)
                        zlAddArray mcllPayDrugAndStuff, arrPut(i)
                    Next
                End If
            Else
                On Error GoTo errH
                    '�޸Ĺ�����ش���
                    If mstrYBPati <> "" Then
                        gcnOracle.BeginTrans: blnTrans = True
                    End If
                    '��ɾ��ԭ����,��Ϊ����Ԥ������Ҫ�Ȼ�ԭ
                    If strDelBill <> "" Then
                        If mstrYBPati <> "" Then
                            Call zlDatabase.ExecuteProcedure(strDelBill, Me.Caption)
                        Else
                            zlAddArray cllPro, strDelBill
                        End If
                    End If
                    'a.��ҽ��ֱ���շ�
                    If Not (blnֱ���շ� And mstrYBPati <> "") Then
                        'ɾ�����￨���۵�:���ŵ���ʱֻɾ��һ��(��Ϊͨ�����￨�Ŷ�����ʱ,���￨���۵��������շ�ϸĿ��,����Ҫɾ��)
                        If mstrCardNO <> "" And strSaveNos = "" Then
                            gstrSQL = "zl_���ﻮ�ۼ�¼_Delete('" & mstrCardNO & "')"
                            If mstrYBPati <> "" Then
                                 Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                            Else
                                    zlAddArray cllPro, gstrSQL
                            End If
                        End If
                        'ִ�������SQL���
                        For i = 0 To UBound(arrSQL)
                            If mstrYBPati <> "" Then
                                Call zlDatabase.ExecuteProcedure(Mid(arrSQL(i), InStr(arrSQL(i), ";") + 1), Me.Caption)
                            Else
                                zlAddArray cllPro, Mid(arrSQL(i), InStr(arrSQL(i), ";") + 1)
                            End If
                        Next
                        'b.ҽ��ֱ���շ�
                    Else
                         If mstrYBPati <> "" Then
                            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                         Else
                            zlAddArray cllPro, gstrSQL
                         End If
                    End If
                    
                    '�շ���ɺ�Ĵ���
                    '-----------------------------------------------------
                    '����д��ʼƱ�ݺ��Ա�ҽ������ʱ�ϴ�,���ŷֱ��ӡʱ,��д��ͬ��,��ӡ����ʱ����д,ȡ����ӡ���ӡʧ�ܽ����
                    '�޸�ʱ,ֻ��д�µ��ݵĿ�ʼƱ�ݺ�,��Ϊҽ��ֻ���µ����ϴ�
                    If strInvoice <> "" And mblnPrint Then
                        gstrSQL = "Zl_Ʊ����ʼ��_Update('" & strBillNO & "','" & strInvoice & "',1)"
                        If mstrYBPati <> "" Then
                            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                        Else
                            zlAddArray cllPro, gstrSQL
                        End If
                    End If
                
                    'ÿ�ŵ��ݴ������,�ý���ID������ɵ��շѼ�¼��ͬ
                    If mobjBill.Pages(p).����� <> 0 Then '44657
                        gstrSQL = "zl_�����շ����_Insert('" & strBillNO & "'," & mobjBill.Pages(p).����� & ",0,1)"
                        If mstrYBPati <> "" Then
                            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                        Else
                            zlAddArray cllPro, gstrSQL
                        End If
                    End If
                    
                    '�շѺ��Զ���ҩ,�Զ�����
                    If strDeptIDs <> "" Or strStuffDept <> "" Then
                        For i = 0 To UBound(arrPut)
                            zlAddArray mcllPayDrugAndStuff, CStr(arrPut(i))
                        Next
                    End If
                    
                    '�޸Ĺ�����ش���
                    If strDelBill <> "" Then
                        '�շѣ��µ��ݹ�����ԭ���ݵĴ�ӡID��,�Ա�һ���ش�,��ʱ��δ����Ʊ��
                        If lng��ӡID <> 0 And mblnPrint Then
                            gstrSQL = "zl_�����շ�Ʊ��_Insert('" & strBillNO & "','',Null,'',Null," & lng��ӡID & ",0)"
                            If mstrYBPati <> "" Then
                                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                            Else
                                zlAddArray cllPro, gstrSQL
                            End If
                        End If
                    End If
                    'ҽ������
                    If mstrYBPati <> "" Then
                        If zlInsureOneBillClinicSwap(lng����ID, strBillNO, strInvoice, strDelBill <> "", p, Original.����ID, blnPriceSaved) = False Then
                           'If strSaveNos <> "" Then strSaveNos = Mid(strSaveNos, 2)
                           If strSaveCuessNos <> "" Then
                                 strSaveCuessNos = Mid(strSaveCuessNos, 2)
                                 i = UBound(Split(strSaveCuessNos, ",")) + 1
                                Call MsgBox("ע��:" & vbCrLf & _
                                                      "      ҽ���Ѿ��ɹ��շ���" & i & "�ŵ���,��������" & mobjBill.Pages.Count - i & "��" & vbCrLf & _
                                                      "����δ�շѳɹ�,��ֻ�Գɹ����ݽ����շ�!" & vbCrLf & _
                                                      "ҽ���ɹ��շѵ�������: " & vbCrLf & _
                                                       strSaveCuessNos, vbDefaultButton1 + vbInformation + vbOKOnly, gstrSysName)
                               blnSaveBill = True: SaveChargeBill = True
                            End If
                            
                            Exit Function
                        End If
                        strSaveCuessNos = strSaveCuessNos & "," & strBillNO
                    End If
                    
                On Error GoTo 0
            End If
            strBalanceIDs = IIf(strBalanceIDs = "", "", strBalanceIDs & ",") & lng����ID
            cllPageInfor.Add Array(lng����ID, strBillNO), "K" & p
            
            '�ύ�ɹ������ۼ�
            If mbytInFun = 0 And Not mblnSaveAsPrice Then
                cur�ѽɺϼ� = cur�ѽɺϼ� + mobjBill.Pages(p).Ӧ�ɽ��
            End If
            
            strSaveNos = strSaveNos & ",'" & strBillNO & "'"
            If Left(strSaveNos, 1) = "," Then strSaveNos = Mid(strSaveNos, 2)
            '���뵥����ʷ��¼(�������͵���)
            cboNO.AddItem strBillNO, 0
            For i = cboNO.ListCount - 1 To 10 Step -1
                cboNO.RemoveItem i 'ֻ��ʾ10��
            Next
        End If
    Next  '��һ�ŵ���
    
    On Error GoTo errH:
    '�ȱ��浥��
    Dim blnAffair As Boolean
    blnTrans = True
    If mstrYBPati = "" Or (mstrYBPati <> "" And (MCPAR.�൥�ݵ�һ�ν��� Or MCPAR.�൥��һ�ν���)) Then
        If mcllPayDrugAndStuff.Count <> 0 And blnNotCommit Then
            '�Զ�����(ֻһ���ύ����ʱ,ͬʱ������,����ͬһ������)
            For i = 1 To mcllPayDrugAndStuff.Count
                zlAddArray cllPro, mcllPayDrugAndStuff(i)
            Next
            Set mcllPayDrugAndStuff = New Collection
        End If
        zlExecuteProcedureArrAy cllPro, Me.Caption, True
        If mstrYBPati <> "" Then
            If zlInsureClinicSwap(cllPageInfor, lng�������, strInvoice, strDelBill <> "", _
                strBalanceIDs, strSaveNos, strSaveCuessNos, blnAffair) = False Then
                If Not blnAffair Then gcnOracle.RollbackTrans
                If strSaveCuessNos <> "" Then blnSaveBill = True:
                Exit Function
            End If
        End If
        If blnAffair = False And blnNotCommit = False Then
            gcnOracle.CommitTrans
        Else
            '�����Զ�����(ͬһ����)
            If blnNotCommit Then
                zlExecuteProcedureArrAy mcllPayDrugAndStuff, Me.Caption, True, True
                Set mcllPayDrugAndStuff = Nothing
            End If
        End If
    End If
    blnSaveBill = True: blnTrans = False: mblnNotClearLedDisplay = True
    SaveChargeBill = True
    Exit Function
errH:
    If Err.Description Like "*��ǰ���㵥�۲�һ��*" Then
        If blnTrans Then gcnOracle.RollbackTrans
        If MsgBox("ĳЩ����ҩƷ�۸��ѷ����仯��Ҫ�Զ�����۸���", vbYesNo + vbQuestion + vbDefaultButton1, gstrSysName) = vbYes Then
            Call CalcMoneys
            Call ShowDetails
            Call ShowMoney
            Exit Function
        End If
     Else
        If blnTrans Then gcnOracle.RollbackTrans
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    End If
    Exit Function
ErrPutOut:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
Private Function zlInsureOneBillClinicSwap(ByVal lng����ID As Long, _
    ByVal strBillNO As String, _
    ByVal strInvoice As String, _
    ByVal blnModifyBill As Boolean, _
    ByVal intPage As Integer, _
    ByVal lngԭ����ID As Long, _
    ByVal blnPriceSaved As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ҽ������(���ŵ���)
    '���:blnModifyBill-�Ƿ��޸ĵ���
    '       strBalanceIDs:���ν��ʵ�ID,�ֱ��ö��ŷ���
    '       strSaveNos-����ĵ��ݺ�
    '       lngԭ����ID-ԭ���޸ĵĵ��ݵĽ���ID
    '       blnPriceSaved-�Ƿ񱣴��˻��۵���
    '����:ҽ�����óɹ����ҽ��,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-20 17:15:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnTransMedicare As Boolean, str���ս��� As String, strAdvance As String, blnMedicareCheck As Boolean
    Dim strTmp As String, i As Long
    Dim intExeCount As Integer
    
    On Error GoTo errHandle
    '���±�־
    ' Zl_���������շ�_ҽ������
    gstrSQL = "Zl_���������շ�_ҽ������("
    '  ����id_In   ������ü�¼.����id%Type,
    gstrSQL = gstrSQL & "" & lng����ID & ","
    '  �������_In ����Ԥ����¼.�������%Type,
    gstrSQL = gstrSQL & "NULL,"
    '  ���ս���_In Varchar2
    gstrSQL = gstrSQL & "NULL)"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    '---------------------------------------------------
    '1.�޸�ʱ,���˳�ԭ�շѵ���(�ķѷ�ʽ)
    blnTransMedicare = False
    strAdvance = ""
    If blnModifyBill Then
        strAdvance = mobjBill.Pages.Count & "|" & intPage
        If Not gclsInsure.ClinicDelSwap(lngԭ����ID, False, mintInsure, strAdvance) Then
            gcnOracle.RollbackTrans: Call DelMedicareTempNO(blnPriceSaved, strBillNO): Exit Function
        Else
            blnTransMedicare = True
        End If
    End If
    
    '2.��ҽ������
    '��Ϊ�����̶�Ϊҽ������,�������ƹ̶�Ϊҽ������(����ͳ�ﲻ��ȷ��),�Ժ�Ӧȥ���ò���
    If (GetMedicareSum(, intPage) <> 0 Or MCPAR.������봫����ϸ) Then
        strAdvance = mobjBill.Pages.Count & "|" & intPage
        If Not gclsInsure.ClinicSwap(lng����ID, GetMedicareSum(mstr�����ʻ�, intPage), _
            GetMedicareSum("ҽ������", intPage), mobjBill.Pages(intPage).ȫ�Ը�, mobjBill.Pages(intPage).���Ը�, mintInsure, strAdvance) Then
            gcnOracle.RollbackTrans:  Call DelMedicareTempNO(blnPriceSaved, strBillNO)
            If blnModifyBill Then Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, False, mintInsure)
            Exit Function
        Else
            blnTransMedicare = True
        End If
    End If
    gcnOracle.CommitTrans
    If blnModifyBill Then Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, True, mintInsure)
    str���ս��� = GetMedicareStr(intPage)
    blnMedicareCheck = zlInsureCheck(str���ս���, strAdvance)
    If blnMedicareCheck Then
        ' Zl_���������շ�_ҽ������
        gstrSQL = "Zl_���������շ�_ҽ������("
        '  ����id_In   ������ü�¼.����id%Type,
        gstrSQL = gstrSQL & "" & lng����ID & ","
        '  �������_In ����Ԥ����¼.�������%Type,
        gstrSQL = gstrSQL & "NULL,"
        '  ���ս���_In Varchar2
        gstrSQL = gstrSQL & IIf(blnMedicareCheck, "'" & strAdvance & "'", "NULL") & ")"
        Err = 0: On Error GoTo ErrModifyTag:
        zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    End If
    Err = 0: On Error GoTo errHandle:
    If blnTransMedicare Then Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicSwap, True, mintInsure)
    zlInsureOneBillClinicSwap = True
    Exit Function
ErrModifyTag:
    If intExeCount > 3 Then
        MsgBox "����[" & strBillNO & "]����ҽ������У��3������ʧ��,����ϵͳ����Ա��ϵ,�϶���������:" & vbCrLf & strAdvance & vbCrLf & "����ԭ������:" & vbCrLf & Err.Description, vbInformation + vbOKOnly, gstrSysName
        intExeCount = 3
    Else
        MsgBox "����[" & strBillNO & "]����ҽ������У��ʧ��,�������ȷ,���ȷ�������½϶�, �϶���������:" & vbCrLf & _
         strAdvance & vbCrLf & "����ԭ������:" & vbCrLf & Err.Description, vbInformation, gstrSysName
    End If
    intExeCount = intExeCount + 1
    Resume
    Exit Function
errHandle:
     gcnOracle.RollbackTrans
     Call ErrCenter
    'ҽ����HIS����ͬһ������,HIS����ʧ��,��ҽ���������ϴ�,������Ҫ��"ȡ������"�ӿ�
    If blnTransMedicare Then
        Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicSwap, False, mintInsure)
    Else     '���ҽ���ɹ��ˣ���ɾ�����۵�������ʧ�ܿ�������
        Call DelMedicareTempNO(False, strBillNO)
    End If
    Call SaveErrLog
End Function
Public Function zlGetToTatal() As Currency
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���νɿ��ܶ�
    '����:���˺�
    '����:2012-02-17 15:25:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objTmpDetail As New BillDetail
    Dim objTmpIncome As New BillInCome
    Dim curTotal As Currency, intCol As Integer
    Dim i As Integer, j As Integer, k As Integer
    
    For i = 1 To mobjBill.Pages.Count
        If mobjBill.Pages(i).Details.Count > 0 Then
            curTotal = curTotal + mobjBill.Pages(i).�����
            For j = 1 To mobjBill.Pages(i).Details.Count
                For k = 1 To mobjBill.Pages(i).Details(j).InComes.Count
                    curTotal = curTotal + mobjBill.Pages(i).Details(j).InComes(k).ʵ�ս��
                Next
            Next
        Else    '��ȡ���۵��շ�ʱû����ϸ����
            curTotal = curTotal + mobjBill.Pages(i).�����
            curTotal = curTotal + mobjBill.Pages(i).ʵ�ս��
        End If
    Next
    
    '���û��,�ٳ��Դӱ����ȡ(��һ�ŵ���ʱ)
    If curTotal = 0 And tbsBill.Tabs.Count = 1 _
        And Not (Bill.Rows = 2 And Bill.TextMatrix(1, BillCol.��Ŀ) = "") Then
        intCol = BillCol.ʵ�ս��
        For i = 1 To Bill.Rows - 1
            If IsNumeric(Bill.TextMatrix(i, intCol)) Then
                curTotal = curTotal + Format(Val(Bill.TextMatrix(i, intCol)), gstrDec)
            End If
        Next
    End If
    zlGetToTatal = Format(curTotal, gstrDec)
End Function


Private Function Getδ��ҩƷ��ҩ����(ByVal lng����ID As Long, ByVal lngִ�в���ID As Long) As String
    '-------------------------------------------------------------------------
    '���ܣ��жϵ�ǰ�����Ƿ������ִͬ�в��ŵ�δ��ҩƷ���������򷵻�δ��ҩƷ�ķ�ҩ����
    '���أ���������ִͬ�в��ŵ�δ��ҩƷ���򷵻�δ��ҩƷ�ķ�ҩ���ڣ����򷵻ؿ�
    '���ƣ�Ƚ����
    '���ڣ�2014-04-09
    '���⣺71902
    '˵����
    '   ͬһ���˲��˲�ͬʱ��ζ��ŵ����շѣ�����ͬһ����ҩ���ڣ����㲡��ȡҩ
    '-------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0: On Error GoTo Errhand
    strSQL = "Select ��ҩ����" & vbNewLine & _
            "From δ��ҩƷ��¼" & vbNewLine & _
            "Where ���� = 8 And ��ҩ���� Is Not Null And ����id = [1] And �ⷿid = [2]" & vbNewLine & _
            "Order By ���շ� Desc, �������� Desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����δ��ҩƷ��ҩ����", lng����ID, lngִ�в���ID)
    
    If Not rsTemp.EOF Then
        Getδ��ҩƷ��ҩ���� = Nvl(rsTemp!��ҩ����)
    End If
    rsTemp.Close: Set rsTemp = Nothing
    
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function Set��ҩ����(ByVal p As Integer, ByRef objBillDetail As BillDetail) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���÷�ҩ����
    '����:  ���óɹ�������true,���򷵻�False
    '����:���˺�
    '����:2012-07-03 09:53:33
    '����:45172
    '˵��:
    '   ����ҩ��ID��ȷ��,��ͬ��ҩ��ID������ͬ�ķ�ҩ����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, i As Long, strSendWindows As String
    Dim blnFind As Boolean
    Dim strTemp As String
    
    Err = 0:     On Error GoTo errHandle:
    
    With objBillDetail
        '�շѡ����۵�ҩƷ��,����ҩ����
        If Not InStr(",5,6,7,", .�շ����) > 0 Then Set��ҩ���� = True: Exit Function
        
        '���޸ĵ���
        'array(ҩ��ID,����),����Ƿ���ڸô��ڣ���֤��ͬҩ����ͬһ������
        strSendWindows = ""
        blnFind = False
        For i = 1 To mCllWindows.Count
            If mCllWindows(i)(0) = .ִ�в���ID Then
                strSendWindows = mCllWindows(i)(1): blnFind = True
            End If
        Next
        
        If mstrInNO <> "" Then
            '�޸ĵ���
            .��ҩ���� = IIf(strSendWindows <> "", strSendWindows, .��ҩ����) '�޸�ʱ����ԭ�з�ҩ����
            Set��ҩ���� = True
            Exit Function
        End If
        
        '71902,Ƚ����,2014-04-09,ͬһ���˲��˲�ͬʱ��ζ��ŵ����շѣ�����ͬһ����ҩ���ڣ����㲡��ȡҩ
        '�жϵ�ǰ�����Ƿ������ִͬ�в��ŵ�δ��ҩƷ���������򷵻�δ��ҩƷ�ķ�ҩ����
        strTemp = Getδ��ҩƷ��ҩ����(mobjBill.����ID, .ִ�в���ID)
        If strTemp <> "" Then
            .��ҩ���� = strTemp
            Set��ҩ���� = True: Exit Function
        End If
        
        If strSendWindows <> "" Then    '���ڷ�ҩ���ڣ��Ե�һ��Ϊ׼
            .��ҩ���� = strSendWindows: Set��ҩ���� = True: Exit Function
        End If
        
        .��ҩ���� = GetDrugWindow(.ִ�в���ID, .�շ����, p)
        If .��ҩ���� = "" Then
           .��ҩ���� = Get��ҩ����(mobjBill.�Ǽ�ʱ��, .ִ�в���ID, .�շ����, _
                       IIf(.ִ�в���ID <> mlng��ҩ��, "", mstr����), IIf(.ִ�в���ID <> mlng��ҩ��, "", mstr�ɴ�), IIf(.ִ�в���ID <> mlng��ҩ��, "", mstr�д�))
        End If
        If .��ҩ���� <> "" Then
            Select Case .�շ����
                Case "5"
                    mstr���� = .��ҩ����
                Case "6"
                    mstr�ɴ� = .��ҩ����
                Case "7"
                    mstr�д� = .��ҩ����
            End Select
        ElseIf ExistWindow(.ִ�в���ID, mrs��ҩ����) Then
            MsgBox "�޷�����" & GET��������(.ִ�в���ID, mrsUnit) & "�ķ�ҩ���ڣ������Ƿ��������Ŵ����ϰࡣ", vbInformation, gstrSysName
            Exit Function
        End If
        If Not blnFind Then
            mCllWindows.Add Array(.ִ�в���ID, .��ҩ����), "K" & .ִ�в���ID
        End If
    End With
    Set��ҩ���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub Clear�����ۼ�()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��Ϣ�ı�ʱ,��������ۼƵ���ʾ
    '����:���˺�
    '����:2012-08-01 10:28:35
    '����:51670
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not mblnClearBlance Then Exit Sub
    '���������Ϣ
    Call InitBalanceGrid(True)
    mblnClearBlance = False
End Sub
Private Sub Set�����շѲ���(Optional blnδ���� As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������շѲ����������Ϣ
    '���:blnδ����-����δ����ʱ
    '����:���˺�
    '����:2012-08-01 10:37:22
    '˵��:
    '����:51670
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not mbln�������� Or mbytInFun <> 0 Or mbytInState <> 0 Then Exit Sub
    
    'ֻ���շѲ��������շ�
    With gTy_Module_Para
        If Not (.byt�ɿ���� = 1 Or .byt�ɿ���� = 3) Then Exit Sub
    End With
    
    '��ʾ�����շѵ���������
    Call LoadCurBalance: sta.Panels(2).Text = IIf(mstrPrePati = "", "", "��һ����:" & mstrPrePati)
    If gTy_Module_Para.byt�ɿ���� <> 3 Then Exit Sub
    If blnδ���� Then
        If mstrPrePati = Trim(txtPatient.Text) Or Trim(txtPatient.Text) = "" Then Exit Sub
    Else
        If mrsInfo Is Nothing Then Exit Sub
        If mrsInfo.State <> 1 Then Exit Sub
        'ͬһ����,������������
        If mstrPrePati = mrsInfo!���� Or mlngPrePati = Val(mrsInfo!����ID) Then Exit Sub
    End If
    '��ͬ����ʱ,������������
    mblnClearBlance = True
    mbln�������� = False: Set grsTotal = Nothing
End Sub

Private Sub WriteMzInforToCard(ByVal lng����ID As Long, ByVal lng������� As Long, Optional blnDelete As Boolean = False)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������Ϣд�뿨��
    '���:blnDelete-�Ƿ��˷�
    '����:���˺�
    '����:2012-12-14 17:06:27
    '˵��:
    '����:56615
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngCardTypeID As Long, strExpend As String
    'δȷ��ˢ�����,ֱ���˳�
    If InStr(1, mstrPrivs, ";������Ϣд��;") = 0 Then Exit Sub
    If lng����ID = 0 Then Exit Sub
    If mlngCardTypeID = 0 Then
        If blnDelete Then GoTo goDelete:
        Exit Sub
    End If
    Dim objCard As Card
    If IDKind.GetCurCard.�ӿ���� = mlngCardTypeID Then
        Set objCard = IDKind.GetCurCard
    Else
        Set objCard = IDKind.GetIDKindCard(mlngCardTypeID, CardTypeID)
    End If
    If objCard Is Nothing Then Exit Sub
    If objCard.�Ƿ�д�� = False Or objCard.�ӿ���� <= 0 Then Exit Sub '��׼д����,�����ýӿ�
    lngCardTypeID = objCard.�ӿ����
goDelete:
   Call gobjSquare.objSquareCard.zlMzInforWriteToCard(Me, mlngModul, lngCardTypeID, _
    lng����ID, lng�������, strExpend)
End Sub
Private Sub SetInvoceSizeAndShowTittle()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������Ʊ��ʾ�ؼ��Ĵ�С����ʾ
    '����:���˺�
    '����:2013-05-07 16:14:02
    '����:25187
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllInvoice As New Collection
    Dim r As Long, c As Long
    Dim bytSel As Byte '1-ѡ��;2-��ѡ��,3-����ȡ����ѡ��(������Ʊ)
    Dim strInvoice As String '��Ʊ��
    Dim sngColWidth As Single
    Dim i As Long
    Err = 0: On Error GoTo Errhand:
    Set cllInvoice = New Collection
    With vsInvoice
        If .Rows = 1 And .Cell(flexcpLeft, 0, .COLS - 1) + .ColWidth(.COLS - 1) <= .Width Then Exit Sub
        For r = 0 To .Rows - 1
            For c = 1 To .COLS - 1
                bytSel = .Cell(flexcpChecked, r, c)
                strInvoice = Trim(.Cell(flexcpData, r, c))
                sngColWidth = .ColWidth(c)
                If strInvoice <> "" Then
                    cllInvoice.Add Array(bytSel, strInvoice, sngColWidth)
                End If
            Next
        Next
        .Redraw = flexRDNone
        .Rows = 1
        .COLS = 1
        .Clear
        .TextMatrix(0, 0) = "���շ�Ʊ"
        sngColWidth = .ColWidth(0)
        For i = 1 To cllInvoice.Count
            If sngColWidth + cllInvoice(i)(2) * 0.5 > .Width Then
                If .COLS <= 1 Then
                    .COLS = .COLS + 1
                    .ColWidth(.COLS - 1) = cllInvoice(i)(2)
                End If
                Exit For
            End If
            .COLS = .COLS + 1
            .ColWidth(.COLS - 1) = cllInvoice(i)(2)
            sngColWidth = sngColWidth + .ColWidth(.COLS - 1)
        Next
        .Cell(flexcpChecked, 0, .COLS - 1, .Rows - 1, .COLS - 1) = 0
        c = 0: r = 0
        For i = 1 To cllInvoice.Count
            If c >= .COLS - 1 Then
                .Rows = .Rows + 1
                r = r + 1
                c = 1
            Else
                c = c + 1
            End If
            .TextMatrix(r, c) = cllInvoice(i)(1)
            .Cell(flexcpData, r, c) = cllInvoice(i)(1)
            .Cell(flexcpChecked, r, c) = cllInvoice(i)(0)
            .ColWidth(c) = cllInvoice(i)(2)
        Next
        .Height = (.RowHeight(0) + 90) * (.Rows)
        .Redraw = flexRDBuffered
    End With
    Exit Sub
Errhand:
    vsInvoice.Redraw = flexRDBuffered
End Sub
Private Sub ShowInvoiceInfor()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ��Ʊ��Ϣ
    '����:���˺�
    '����:2013-05-27 11:36:16
    '25187
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If gTy_Module_Para.bytƱ�ݷ������ = 0 Then Exit Sub
    If Not (mbytInFun = 0 And mbytInState = 3) Then Exit Sub
    
    If mrsDelInvoice Is Nothing Then
        vsInvoice.Visible = False:
        Call Form_Resize
    End If
    If mrsDelInvoice.RecordCount = 0 Then
        vsInvoice.Visible = False:
        Call Form_Resize
        Exit Sub
    End If
    vsInvoice.Visible = True
    Call Form_Resize
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub LoadInvoiceData(ByVal strNo As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ط�Ʊ��Ϣ
    '����:���˺�
    '����:2013-05-07 17:07:38
    '����:25187
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str��� As String, varTemp As Variant
    Dim i As Long, str��Ʊ�� As String
    If gTy_Module_Para.bytƱ�ݷ������ = 0 Then Exit Sub
    If Not (mbytInFun = 0 And mbytInState = 3) Then Exit Sub
    
    If mrsDelInvoice Is Nothing Then
        Set mrsDelInvoice = zlGetFromNoTOInvoice(strNo)
    End If
    If mrsDelInvoice Is Nothing Then Exit Sub
    If mrsDelInvoice.RecordCount = 0 Then Exit Sub
    With Bill
        For i = 1 To .Rows - 1
            If InStr("����,�˷�", Bill.TextMatrix(0, Bill.COLS - 1)) > 0 Then
                If Bill.TextMatrix(i, Bill.COLS - 1) = "��" Then
                    str��� = str��� & "," & Bill.RowData(i)
                End If
            End If
        Next
    End With
    If str��� <> "" Then str��� = Mid(str���, 2)
    str��Ʊ�� = GetFromNumToInvoiceNo(strNo, str���)
    '���ط�Ʊ��
    varTemp = Split(str��Ʊ��, ",")
    With vsInvoice
        .Clear
        .Rows = 1: .COLS = 1
        .FixedCols = 1
        .TextMatrix(0, 0) = "����Ʊ��"
        .Redraw = flexRDNone
        .COLS = .COLS + UBound(varTemp) + 1
        For i = 0 To UBound(varTemp)
            If i + 1 > .COLS - 1 Then
                .COLS = .COLS + 1
            End If
            .TextMatrix(0, i + 1) = CStr(varTemp(i))
            .Cell(flexcpData, 0, i + 1) = CStr(varTemp(i))
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .COLS - 1)
        Call Form_Resize
        .Redraw = flexRDBuffered
    End With
End Sub
Private Function GetFromNumToInvoiceNo(ByVal strNo As String, ByVal str��� As String) As String
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������Ż�ȡ��Ӧ�ķ�Ʊ��
    '���:strNO-���ݺ�
    '       str���-���,����Ϊ���,����ö��ŷ���
    '       strNotInvoice-�������ķ�Ʊ��
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-05-07 17:38:24
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str��Ʊ�� As String, str������� As String
    Dim varTemp As Variant, i As Long, strTemp As String
    On Error GoTo errHandle
    If mrsDelInvoice Is Nothing Then Exit Function
    If mrsDelInvoice.State <> 1 Then Exit Function
    If mrsDelInvoice.RecordCount = 0 Then Exit Function
    With mrsDelInvoice
        str������� = "": str��Ʊ�� = ""
        varTemp = Split(str���, ",")
        .Filter = "NO='" & strNo & "'"
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
                strTemp = "," & Nvl(!���) & ","
                For i = 0 To UBound(varTemp)
                    If InStr(1, strTemp, "," & varTemp(i) & ",") > 0 _
                        And InStr(str��Ʊ�� & ",", "," & Nvl(!Ʊ��) & ",") = 0 Then
                        str��Ʊ�� = str��Ʊ�� & "," & Nvl(!Ʊ��)
                        If Val(Nvl(!����Ʊ�����)) <> 0 Then
                            str������� = str������� & "," & Val(Nvl(!����Ʊ�����))
                        End If
                    End If
                Next
            .MoveNext
        Loop
        .Filter = 0: .MoveFirst
        If str������� = "" Then GoTo GoSort:
        '��Ҫ���ҹ���Ʊ��
       varTemp = Split(Mid(str�������, 2), ",")
        Do While Not .EOF
                For i = 0 To UBound(varTemp)
                    If Val(varTemp(i)) = Val(Nvl(!����Ʊ�����)) _
                        And InStr(str��Ʊ�� & ",", "," & Nvl(!Ʊ��) & ",") = 0 Then
                        str��Ʊ�� = str��Ʊ�� & "," & Nvl(!Ʊ��)
                    End If
                Next
            .MoveNext
        Loop
    End With
    '����������
GoSort:
    If str��Ʊ�� = "" Then Exit Function
    str��Ʊ�� = Mid(str��Ʊ��, 2)
    GetFromNumToInvoiceNo = zlStringSort(str��Ʊ��)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function CheckChargeDataValied() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����շ������Ƿ�Ϸ�
    '����:���ݺϷ�,����true,���򷵻�False
    '����:���˺�
    '����:2013-06-25 16:34:58
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, bln����� As Boolean, p As Integer
    Dim dblToTal As Double, strTmp As String, strInfo As String
    Dim lngҩ��ID As Long
    Dim colStock As Collection
    
    On Error GoTo errHandle

    '���뻮�۵��շ�ʱ,�����ҽ�����ɵ�,����������
    For i = 1 To mobjBill.Pages.Count
        '���ÿ�ŵ����ж�(��Ϊ���ܻ��ۺ��շѻ���),�Ƿ��ǵ���ҽ�����ɵĻ��۵��շ�
        If mobjBill.Pages(i).NO <> "" And mobjBill.Pages(i).ҽ����� <> 0 Then
            If mobjBill.Pages(i).ʵ�ս�� <> GetBillSumByDB(mobjBill.Pages(i).NO) Then
                MsgBox "����[" & mobjBill.Pages(i).NO & "]�Ĳ����շѼ�¼�ѱ������޸Ļ�����,�����¶�ȡ���ݺ����շѣ�", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    Next
    
    For p = 1 To mobjBill.Pages.Count
        For i = 1 To mobjBill.Pages(p).Details.Count
            With mobjBill.Pages(p).Details(i)
                If CheckServeRange(0, .�շ�ϸĿID) = False Then Exit Function
            End With
        Next i
    Next p
    
   'ҩƷ�����(�������ֹʱ�����ʱ��ҩƷ)
    bln����� = (InStr(mstrPrivs, "�������") = 0)    '�Ƿ���Ȩ�޲������(������ʱ�۱�����)
    For p = 1 To mobjBill.Pages.Count
        For i = 1 To mobjBill.Pages(p).Details.Count
            With mobjBill.Pages(p).Details(i)
                Set colStock = IIf(.�շ���� = "4", mcolStock2, mcolStock1)
            
                If InStr(",5,6,7,", .�շ����) > 0 Then
                    If .Detail.���� Or .Detail.��� Then
                        dblToTal = GetDrugTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                        .Detail.��� = GetStock(.�շ�ϸĿID, .ִ�в���ID)
                        If gblnҩ����λ Then .Detail.��� = .Detail.��� / .Detail.ҩ����װ
                        
                        If mbytInState = 0 And mstrInNO <> "" Then .Detail.��� = .Detail.��� + GetOriginalTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                        If dblToTal > .Detail.��� Then
                            If mobjBill.Pages.Count > 1 Then strTmp = "�� " & p & " �ŵ���"
                            MsgBox strTmp & "�� " & i & " ��ʱ�ۻ����ҩƷ""" & .Detail.���� & _
                                """�ĵ�ǰ���" & IIf(InStr(1, mstrPrivs, "��ʾ���") > 0, .Detail.���, "") & "������������""" & _
                                dblToTal & """��", vbInformation, gstrSysName
                            'tbsBill.Tabs(p).Selected = True
                            Exit Function
                        End If
                    Else
                        If colStock("_" & .ִ�в���ID) = 2 And bln����� Then
                            dblToTal = GetDrugTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                            .Detail.��� = GetStock(.�շ�ϸĿID, .ִ�в���ID)
                            If gblnҩ����λ Then .Detail.��� = .Detail.��� / .Detail.ҩ����װ
                            
                            If mbytInState = 0 And mstrInNO <> "" Then .Detail.��� = .Detail.��� + GetOriginalTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                            If dblToTal > .Detail.��� Then
                                If mobjBill.Pages.Count > 1 Then strTmp = "�� " & p & " �ŵ���"
                                MsgBox strTmp & "�� " & i & " ��ҩƷ""" & .Detail.���� & _
                                    """�ĵ�ǰ���" & IIf(InStr(1, mstrPrivs, "��ʾ���") > 0, .Detail.���, "") & "������������""" & _
                                    dblToTal & """,���޸Ļ����Ƿ��ж������롣", vbInformation, gstrSysName
                                'tbsBill.Tabs(p).Selected = True
                                Exit Function
                            End If
                        End If
                    End If
                ElseIf .�շ���� = "4" And .Detail.�������� Then
                    If .Detail.���� Or .Detail.��� Then
                        dblToTal = GetDrugTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                        .Detail.��� = GetStock(.�շ�ϸĿID, .ִ�в���ID, .Detail.����)
                        
                        If mbytInState = 0 And mstrInNO <> "" Then .Detail.��� = .Detail.��� + GetOriginalTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                        If dblToTal > .Detail.��� Then
                            If mobjBill.Pages.Count > 1 Then strTmp = "�� " & p & " �ŵ���"
                            MsgBox strTmp & "�� " & i & " ��ʱ�ۻ������������""" & .Detail.���� & _
                                """�ĵ�ǰ���" & IIf(InStr(1, mstrPrivs, "��ʾ���") > 0, .Detail.���, "") & "������������""" & dblToTal & """��", vbInformation, gstrSysName
                            'tbsBill.Tabs(p).Selected = True:
                            Exit Function
                        End If
                    Else
                        If colStock("_" & .ִ�в���ID) = 2 And bln����� Then
                            dblToTal = GetDrugTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                            .Detail.��� = GetStock(.�շ�ϸĿID, .ִ�в���ID, .Detail.����)
                            
                            If mbytInState = 0 And mstrInNO <> "" Then .Detail.��� = .Detail.��� + GetOriginalTotal(mobjBill, .�շ�ϸĿID, .ִ�в���ID)
                            If dblToTal > .Detail.��� Then
                                If mobjBill.Pages.Count > 1 Then strTmp = "�� " & p & " �ŵ���"
                                MsgBox strTmp & "�� " & i & " ����������""" & .Detail.���� & _
                                    """�ĵ�ǰ���" & IIf(InStr(1, mstrPrivs, "��ʾ���") > 0, .Detail.���, "") & "������������""" & dblToTal & """,���޸Ļ����Ƿ��ж������롣", vbInformation, gstrSysName
                                'tbsBill.Tabs(p).Selected = True:
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End With
        Next
    Next
    
    '��ҩ���ڼ��(�����۵�)
    For i = 1 To mobjBill.Pages.Count
        If mobjBill.Pages(i).NO <> "" And tbsBill.Tabs(i).Tag = "" Then
            lngҩ��ID = BillExistDrug(mobjBill.Pages(i).NO, 1)
            If lngҩ��ID <> 0 Then
                If ExistWindow(lngҩ��ID, mrs��ҩ����) Then
                    MsgBox "�޷�����" & GET��������(lngҩ��ID, mrsUnit) & "�ķ�ҩ���ڣ���ȷ���Ƿ��������Ŵ����ϰࡣ", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
    Next
    
    If mstrInNO <> "" Then
        If HaveExecute(1, mstrInNO, IIf(mbytInFun = 2, 2, 1)) Then
            MsgBox "�õ��ݰ�����ȫִ�л򲿷�ִ�е���Ŀ,�������޸ġ�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    CheckChargeDataValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 Public Sub SendMsgModule()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��Ϣ���ʹ���
    '���: 0-�շѻ��۵�;1-�����շѵ�;2-���ʻ��۵�;3-���ʵ�
    '     strNO-���ݺ�
    '����:���˺�
    '����:2014-03-11 11:59:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, objDrugXML As New clsXML, objCheckXML As New clsXML
    Dim objTemp As clsXML, str�շ�ʱ�� As String
    Dim rsTemp As ADODB.Recordset, int���� As Integer
    Dim blnֱ���շ� As Boolean, p As Long
    Dim lngDrug As Long, lngCheck As Long, blnAddBill As Boolean, blnHaveCheck As Boolean, blnHaveDrug As Boolean
    
    'mbytInFun:0-�շ�,1-����,2-�������
    '  mbytInState  :0-ִ��(���޸�),1-���,2-����,3-�˷�(�շѡ����ʲ����˷�),4-�����շ�;5-�쳣��������
    On Error GoTo errHandle
    
    
    If Not (mbytInState = 0 Or mbytInState = 5) Then Exit Sub
    If mobjMsgModule Is Nothing Then Exit Sub
    If mobjMsgModule.IsConnect = False Then Exit Sub
    
    If Format(mobjBill.�Ǽ�ʱ��, "yyyy") < 2000 Then mobjBill.�Ǽ�ʱ�� = zlDatabase.Currentdate
    str�շ�ʱ�� = mobjBill.�Ǽ�ʱ��
    
    
    'ZLHIS_CHARGE_003 ������õ���
    '�ڵ�����    ����    ����    �ظ�    ����    ȱʡֵ  ֵ������
    'patient_info        ������Ϣ    1
    '   patient_id      ����id  1   N
    '   patient_name        ����    1   S
    '   patient_sex     �Ա�    1   S
    '   patient_age     ����    1   S
    '   identity_card       ���֤��    0..1    S
    '   in_number       סԺ��  0..1    S
    '   out_number      �����  0..1    S
    'charge_bill         1..*
    '   bill_no     ���ݺ���    1   S
    '   bill_kind       ��������    1   N       1-�շѵ�;2-���ʵ�
    '   charge_state        �շ�״̬    1   N       1-δ�շ�;2-���շ�
    '   charge_time     �շ�ʱ��    1   S
    '   charge_person       �շ���Ա    1   S
    '   bill_item           1..*
    '       charge_item_id      �շ���Ŀid  1   N
    '       charge_item_kind        �շ����    1   S
    '       execute_dept_id     ִ�в���id  1   N
    '       drug_window     ��ҩ����    0..1    S
    objDrugXML.ClearXmlText
    objCheckXML.ClearXmlText
    blnHaveCheck = False: blnHaveDrug = False
    For p = 1 To mobjBill.Pages.Count
    
        If mobjBill.Pages(p).NO = "" Then
            blnֱ���շ� = True
        Else
            blnֱ���շ� = False
        End If
        
        If p = 1 Then
            'ҩƷ
            Call objDrugXML.AppendNode("patient_info")
                Call objDrugXML.appendData("patient_id", mobjBill.����ID)
                Call objDrugXML.appendData("patient_name", mobjBill.����)
                Call objDrugXML.appendData("patient_sex", mobjBill.�Ա�)
                Call objDrugXML.appendData("patient_age", mobjBill.����)
                '���֤�ź�סԺ���ݲ���(���岻��)
                Call objDrugXML.appendData("out_number", mobjBill.��ʶ��)
            Call objDrugXML.AppendNode("patient_info", True)
            '���
            Call objCheckXML.AppendNode("patient_info")
                Call objCheckXML.appendData("patient_id", mobjBill.����ID)
                Call objCheckXML.appendData("patient_name", mobjBill.����)
                Call objCheckXML.appendData("patient_sex", mobjBill.�Ա�)
                Call objCheckXML.appendData("patient_age", mobjBill.����)
                '���֤�ź�סԺ���ݲ���(���岻��)
                Call objCheckXML.appendData("out_number", mobjBill.��ʶ��)
            Call objCheckXML.AppendNode("patient_info", True)
        End If
        
        If blnֱ���շ� Then
          '��Ի��۵������շѵ�
          lngDrug = 1: lngCheck = 1
          
          For Each mobjBillDetail In mobjBill.Pages(p).Details
            
              blnAddBill = False
              If InStr(1, ",5,6,7,", "," & mobjBillDetail.�շ���� & ",") > 0 _
                And Not gbln�շѺ��Զ���ҩ Then
                '�����Զ���ҩ
                  'ҩƷ
                  Set objTemp = objDrugXML
                  If lngDrug = 1 Then blnAddBill = True
                  blnHaveDrug = True
                  lngDrug = lngDrug + 1
                  
              ElseIf InStr(1, ",D,", "," & mobjBillDetail.�շ���� & ",") > 0 Then
                  '���
                  Set objTemp = objCheckXML
                  If lngCheck = 1 Then blnAddBill = True
                  lngCheck = lngCheck + 1
                  blnHaveCheck = True
              Else
                  Set objTemp = Nothing
              End If
              
              If Not objTemp Is Nothing Then
                If blnAddBill Then
                    Call objTemp.AppendNode("charge_bill")
                    Call objTemp.appendData("bill_no", mobjBill.Pages(p).�շѵ���)
                    If mbytInFun = 1 Or (mbytInFun = 0 And (mblnSaveAsPrice Or mstrYBPati <> "")) Then
                        '���ﻮ��(�շ�)
                        Call objTemp.appendData("bill_kind", 1)  '1-�շѵ�;2-���ʵ�
                        Call objTemp.appendData("charge_state", 1)   '1-δ�շ�;2-���շ�
                    ElseIf mbytInFun = 2 Then
                        '�������
                        Call objTemp.appendData("bill_kind", 2)  '1-�շѵ�;2-���ʵ�
                        Call objTemp.appendData("charge_state", IIf(mbytBilling = 1, 1, 2))  '1-δ�շ�;2-���շ�
                    Else
                        Call objTemp.appendData("bill_kind", 1)  '1-�շѵ�;2-���ʵ�
                        Call objTemp.appendData("charge_state", 2)   '1-δ�շ�;2-���շ�
                    End If
                    Call objTemp.appendData("charge_time", str�շ�ʱ��)
                    Call objTemp.appendData("charge_person", UserInfo.����)
                End If
                '----------------------------------------------------------------------------
                '��ϸ��
                objTemp.AppendNode ("bill_item")
                '       charge_item_id      �շ���Ŀid  1   N
                    Call objTemp.appendData("charge_item_id", mobjBillDetail.�շ�ϸĿID)
                '       charge_item_kind        �շ����    1   S
                    Call objTemp.appendData("charge_item_kind", mobjBillDetail.�շ����)
                '       execute_dept_id     ִ�в���id  1   N
                    Call objTemp.appendData("execute_dept_id", mobjBillDetail.ִ�в���ID)
                '       drug_window     ��ҩ����    0..1    S
                    Call objTemp.appendData("drug_window", mobjBillDetail.��ҩ����)
                Call objTemp.AppendNode("bill_item", True)
              End If
          Next
        End If
        If Not blnֱ���շ� Then
            '���۵�,��˵�
            strSQL = "" & _
            "   Select NO,�շ����,�շ�ϸĿID,ִ�в���ID,��ҩ����,�Ǽ�ʱ��,����Ա����" & _
            "   From ������ü�¼ " & _
            "   Where NO=[1] And ��¼����=[2] And  ��¼״̬=1 " & _
            "   Order by �շ����"
            If mbytInFun = 2 Then
                int���� = 2
            Else
                int���� = 1
            End If
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mobjBill.Pages(p).NO, int����)
            If rsTemp.EOF Then Exit Sub
            
            lngDrug = 1: lngCheck = 1
            Do While Not rsTemp.EOF
                 blnAddBill = False
                If InStr(1, ",5,6,7,", "," & rsTemp!�շ���� & ",") > 0 Then
                    'ҩƷ
                    Set objTemp = objDrugXML
                    If lngDrug = 1 Then blnAddBill = True
                    blnHaveDrug = True
                    lngDrug = lngDrug + 1
                ElseIf InStr(1, ",D,", "," & rsTemp!�շ���� & ",") > 0 Then
                    '���
                    Set objTemp = objCheckXML
                    If lngCheck = 1 Then blnAddBill = True
                    lngCheck = lngCheck + 1
                    blnHaveCheck = True
                Else
                    Set objTemp = Nothing
                End If
                
                If Not objTemp Is Nothing Then
                  If blnAddBill Then
                        Call objTemp.AppendNode("charge_bill")
                        Call objTemp.appendData("bill_no", Nvl(rsTemp!NO))
                        If mbytInFun = 2 Then
                            '�������
                            mbytBilling = 1
                            Call objTemp.appendData("bill_kind", 2)  '1-�շѵ�;2-���ʵ�
                            Call objTemp.appendData("charge_state", 2)   '1-δ�շ�;2-���շ�
                        Else
                            Call objTemp.appendData("bill_kind", 1)  '1-�շѵ�;2-���ʵ�
                            Call objTemp.appendData("charge_state", 2)   '1-δ�շ�;2-���շ�
                        End If
                      
                        Call objTemp.appendData("charge_time", Format(rsTemp!�Ǽ�ʱ��, "yyyy-mm-dd HH:MM:SS"))
                        Call objTemp.appendData("charge_person", Nvl(rsTemp!����Ա����))
                  End If
                  '----------------------------------------------------------------------------
                  '��ϸ��
                  Call objTemp.AppendNode("bill_item")
                  '       charge_item_id      �շ���Ŀid  1   N
                      Call objTemp.appendData("charge_item_id", Val(Nvl(rsTemp!�շ�ϸĿID)))
                  '       charge_item_kind        �շ����    1   S
                      Call objTemp.appendData("charge_item_kind", Nvl(rsTemp!�շ����))
                  '       execute_dept_id     ִ�в���id  1   N
                      Call objTemp.appendData("execute_dept_id", Nvl(rsTemp!ִ�в���ID))
                  '       drug_window     ��ҩ����    0..1    S
                      Call objTemp.appendData("drug_window", Nvl(rsTemp!��ҩ����))
                  Call objTemp.AppendNode("bill_item", True)
               End If
            rsTemp.MoveNext
          Loop
        End If
        If lngDrug > 1 Then Call objDrugXML.AppendNode("charge_bill", True)
        If lngCheck > 1 Then Call objCheckXML.AppendNode("charge_bill", True)
    
    Next
     
    If blnHaveDrug = True _
        And Not gbln�շѺ��Զ���ҩ Then
        '�����Զ���ҩ
        '��ҩƷ��Ϣ
        Call zlDebugWriteFile(objDrugXML.XmlText)
        Call mobjMsgModule.CommitMessage("ZLHIS_CHARGE_003", objDrugXML.XmlText)
    End If
    If blnHaveCheck Then
        '�������Ϣ
        Call zlDebugWriteFile(objCheckXML.XmlText)
        Call mobjMsgModule.CommitMessage("ZLHIS_CHARGE_003", objCheckXML.XmlText)
    End If
    objDrugXML.ClearXmlText: objCheckXML.ClearXmlText
    Set objDrugXML = Nothing: Set objCheckXML = Nothing
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Sub

Private Sub CreateDrugPacker()
    '����:����������ҩ��(�Զ���ҩ��)
    Dim objComLib As New zl9ComLib.clsComLib
    Dim strPrivs As String
    Dim strMessage As String
    
    mblnDrugPacker = False: mblnDrugMachine = False

    If Not (mbytInFun = 0 And (mbytInState = 0 Or mbytInState = 2 Or mbytInState = 3 _
        Or mbytInState = 4) Or mbytInFun = 2) Then Exit Sub

    Err = 0: On Error Resume Next
    If Val(zlDatabase.GetPara("����ҩƷ�Զ����豸�ӿ�", glngSys, Val("9010-ҩƷ�Զ����豸�ӿ�"))) = 1 Then
        '�����½ӿ�
        Set mobjDrugMachine = CreateObject("zlDrugMachine.clsDrugMachine")
        If Err = 0 Then mblnDrugMachine = True
    End If
    
    If mblnDrugMachine = False Then
        '�ɲ���
        Err = 0
        Set mobjDrugPacker = CreateObject("zlDrugPacker.clsDrugPacker")
        If Err = 0 Then mblnDrugPacker = True
    End If
    
    Err = 0: On Error GoTo 0
    If mblnDrugMachine Then
        'Ȩ�޼��
        strPrivs = GetPrivFunc(glngSys, Val("9010-ҩƷ�Զ����豸�ӿ�"))
        If InStr(";" & strPrivs & ";", ";����;") > 0 Then
            mblnDrugMachine = mobjDrugMachine.Init(1, objComLib, strMessage)
        Else
            mblnDrugMachine = False
        End If
    ElseIf mblnDrugPacker Then

        mblnDrugPacker = mobjDrugPacker.DYEY_MZ_IniSoap
    End If
End Sub
Private Function ReadDrugAndStuffStock(ByVal lng�ⷿID As Long, ByRef objDetail As Detail) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡҩƷ�������ϵĿ����Ϣ
    '���:lng�ⷿID-�ⷿID
    '����:objDetail-Detail����
    '����:�ɹ�����true,���򷵻�Fale
    '����:���˺�
    '����:2018-01-10 09:34:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblStock As Double, strҩ��IDs As String
    
    On Error GoTo errHandle
    If objDetail Is Nothing Then Exit Function
    If InStr(",5,6,7,4,", objDetail.���) = 0 Then ReadDrugAndStuffStock = True: Exit Function
    If objDetail.��� = "4" And objDetail.�������� = False Then ReadDrugAndStuffStock = True: Exit Function
   
    If objDetail.��� = "4" And objDetail.�������� Then
        dblStock = GetStock(objDetail.ID, lng�ⷿID, objDetail.����)
        objDetail.��� = dblStock
        Call ShowStock(lng�ⷿID, objDetail.����, objDetail.���)
        ReadDrugAndStuffStock = True: Exit Function
    End If
    
    '��ǰ��ҩƷ���
    If InStr(",5,6,7,", objDetail.���) > 0 Then
        dblStock = GetStock(objDetail.ID, lng�ⷿID)
        If gblnҩ����λ Then dblStock = dblStock / objDetail.ҩ����װ
        objDetail.��� = dblStock
        Call ShowStock(lng�ⷿID, objDetail.����, objDetail.���)
    End If
    ReadDrugAndStuffStock = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
