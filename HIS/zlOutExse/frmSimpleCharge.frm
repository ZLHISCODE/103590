VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "Mscomm32.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#9.0#0"; "zlIDKind.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSimpleCharge 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���շѴ���"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9915
   BeginProperty Font 
      Name            =   "����"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSimpleCharge.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   9915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox pic��� 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   7620
      ScaleHeight     =   600
      ScaleWidth      =   2310
      TabIndex        =   57
      Top             =   5340
      Visible         =   0   'False
      Width           =   2310
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0111"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1020
         TabIndex        =   58
         Top             =   330
         Width           =   1140
      End
      Begin VB.Label lbl��� 
         Caption         =   "�������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   0
         TabIndex        =   59
         Top             =   120
         Width           =   1155
      End
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   8880
      Top             =   4110
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   1
      ImageHeight     =   18
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSimpleCharge.frx":08CA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   35
      Top             =   7035
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2619
            MinWidth        =   882
            Picture         =   "frmSimpleCharge.frx":09BC
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11245
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Visible         =   0   'False
            Object.Width           =   0
            MinWidth        =   88
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmSimpleCharge.frx":1250
            Key             =   "PY"
            Object.ToolTipText     =   "ƴ��(F7)"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmSimpleCharge.frx":188A
            Key             =   "WB"
            Object.ToolTipText     =   "���(F7)"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1111
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1111
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ZL9BillEdit.BillEdit Bill 
      Height          =   3045
      Left            =   30
      TabIndex        =   9
      Top             =   1830
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   5371
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
      ColWidth0       =   1005
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
   End
   Begin VB.Frame fraTitle 
      Height          =   1035
      Left            =   30
      TabIndex        =   33
      ToolTipText     =   "���:F6"
      Top             =   -120
      Width           =   9885
      Begin VB.TextBox txtRePrint 
         Height          =   360
         Left            =   1140
         MaxLength       =   8
         TabIndex        =   25
         Top             =   615
         Width           =   1320
      End
      Begin VB.TextBox txtInvoice 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   4875
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   615
         Width           =   1560
      End
      Begin VB.ComboBox cboNO 
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   7860
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "���ݺų��Ȳ���ʱ�Զ����㳤��"
         Top             =   615
         Width           =   1500
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
         Height          =   375
         Left            =   9390
         Style           =   1  'Graphical
         TabIndex        =   44
         TabStop         =   0   'False
         ToolTipText     =   "�ȼ�:F8"
         Top             =   615
         Width           =   435
      End
      Begin VB.Label lblRePrint 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ش�(&P)"
         Height          =   240
         Left            =   225
         TabIndex        =   24
         Top             =   675
         Width           =   840
      End
      Begin VB.Label lblFact 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ʊ�ݺ�"
         Height          =   240
         Left            =   4080
         TabIndex        =   10
         Top             =   675
         Width           =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   30
         X2              =   11490
         Y1              =   555
         Y2              =   555
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         X1              =   0
         X2              =   11460
         Y1              =   570
         Y2              =   570
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
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   9405
         TabIndex        =   45
         Top             =   630
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "�����շѵ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   90
         TabIndex        =   38
         ToolTipText     =   "���:F6"
         Top             =   210
         Width           =   1725
      End
      Begin VB.Label lbl���ݺ� 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         Caption         =   "���ݺ�"
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   7080
         TabIndex        =   34
         Top             =   675
         Width           =   720
      End
   End
   Begin VB.Frame fraAppend 
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   15
      TabIndex        =   28
      ToolTipText     =   "���:F6"
      Top             =   4785
      Width           =   9900
      Begin VB.ComboBox cbo���㷽ʽ 
         Height          =   360
         Left            =   1905
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   150
         Width           =   1500
      End
      Begin VB.ComboBox cbo������ 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   4695
         TabIndex        =   13
         Text            =   "cbo������"
         Top             =   150
         Width           =   1710
      End
      Begin MSMask.MaskEdBox txtDate 
         Height          =   360
         Left            =   7380
         TabIndex        =   14
         Top             =   150
         Width           =   2415
         _ExtentX        =   4260
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
      Begin VB.CheckBox chk�Ӱ� 
         Alignment       =   1  'Right Justify
         Caption         =   "�Ӱ�(&W)"
         Height          =   270
         Left            =   90
         TabIndex        =   11
         Top             =   210
         Width           =   1155
      End
      Begin VB.Label lbl���㷽ʽ 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   240
         Left            =   1395
         TabIndex        =   47
         Top             =   210
         Width           =   480
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         Caption         =   "ʱ��"
         Height          =   240
         Left            =   6810
         TabIndex        =   36
         Top             =   210
         Width           =   480
      End
      Begin VB.Label lbl������ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "������"
         Height          =   240
         Left            =   3930
         TabIndex        =   29
         Top             =   210
         Width           =   720
      End
   End
   Begin VB.Frame fraMoney 
      Height          =   1815
      Left            =   15
      TabIndex        =   30
      Top             =   5220
      Width           =   2925
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshMoney 
         Height          =   1635
         Left            =   30
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   150
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   2884
         _Version        =   393216
         Rows            =   5
         Cols            =   3
         FixedCols       =   0
         RowHeightMin    =   320
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   2
         MergeCells      =   1
         AllowUserResizing=   1
         FormatString    =   "^���|^��Ŀ      |^      ���"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   3
      End
   End
   Begin VB.Frame fraStat 
      Height          =   1815
      Left            =   2940
      TabIndex        =   26
      Top             =   5220
      Width           =   4620
      Begin VB.TextBox txtӦ�� 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   3300
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   600
         Width           =   1260
      End
      Begin VB.TextBox txtԤ����� 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   360
         Left            =   3300
         TabIndex        =   18
         Text            =   "0.00"
         Top             =   210
         Width           =   1260
      End
      Begin VB.TextBox txt�ۼ� 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   735
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   1290
         Width           =   1395
      End
      Begin VB.TextBox txtӦ�� 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   735
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   315
         Width           =   1395
      End
      Begin VB.TextBox txt�ɿ� 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   3300
         MaxLength       =   12
         TabIndex        =   20
         Text            =   "0.00"
         Top             =   990
         Width           =   1260
      End
      Begin VB.TextBox txt�ϼ� 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   735
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   810
         Width           =   1395
      End
      Begin VB.TextBox txt�Ҳ� 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3300
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   1380
         Width           =   1260
      End
      Begin VB.Label lblӦ�� 
         AutoSize        =   -1  'True
         Caption         =   "Ӧ�ɽ��"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   2280
         TabIndex        =   53
         Top             =   660
         Width           =   960
      End
      Begin VB.Label lblDeposit 
         AutoSize        =   -1  'True
         Caption         =   "Ԥ�����"
         ForeColor       =   &H00808080&
         Height          =   240
         Left            =   2280
         TabIndex        =   51
         Top             =   270
         Width           =   960
      End
      Begin VB.Label lbl�ۼ� 
         AutoSize        =   -1  'True
         Caption         =   "�ۼ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   90
         TabIndex        =   49
         Top             =   1335
         Width           =   630
      End
      Begin VB.Label lblӦ�� 
         AutoSize        =   -1  'True
         Caption         =   "Ӧ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   90
         TabIndex        =   48
         Top             =   360
         Width           =   630
      End
      Begin VB.Label lbl�ɿ� 
         AutoSize        =   -1  'True
         Caption         =   "�ɿ���"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   2280
         TabIndex        =   43
         Top             =   1050
         Width           =   960
      End
      Begin VB.Label lbl�ϼ� 
         AutoSize        =   -1  'True
         Caption         =   "�ϼ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   300
         Left            =   90
         TabIndex        =   37
         Top             =   855
         Width           =   630
      End
      Begin VB.Label lbl�Ҳ� 
         AutoSize        =   -1  'True
         Caption         =   "�Ҳ����"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   2280
         TabIndex        =   27
         Top             =   1440
         Width           =   960
      End
   End
   Begin VB.Frame fraInfo 
      Height          =   1020
      Left            =   30
      TabIndex        =   32
      Top             =   795
      Width           =   9885
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   360
         Left            =   645
         TabIndex        =   54
         Top             =   195
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   635
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
         NotContainFastKey=   "F1;F12;CTRL+F12;F6;F7;F8;F9;F12;CTRL+F12;ESC"
         AllowAutoICCard =   -1  'True
         AllowAutoIDCard =   -1  'True
         MustSelectItems =   "����"
         BackColor       =   -2147483633
      End
      Begin VB.ComboBox cbo���䵥λ 
         Height          =   360
         Left            =   6105
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   195
         Width           =   580
      End
      Begin VB.ComboBox cboҽ�Ƹ��� 
         Height          =   360
         Left            =   7845
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   195
         Width           =   1980
      End
      Begin VB.ComboBox cbo�������� 
         Height          =   360
         Left            =   1125
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   600
         Width           =   1890
      End
      Begin VB.TextBox txtPatient 
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   1320
         MaxLength       =   100
         TabIndex        =   2
         Top             =   195
         Width           =   1680
      End
      Begin VB.ComboBox cboSex 
         Height          =   360
         Left            =   3675
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   195
         Width           =   1035
      End
      Begin VB.TextBox txt���� 
         Height          =   360
         IMEMode         =   2  'OFF
         Left            =   5310
         MaxLength       =   20
         TabIndex        =   4
         Top             =   195
         Width           =   765
      End
      Begin VB.ComboBox cbo�ѱ� 
         Height          =   360
         Left            =   3675
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lbl��̬�ѱ� 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   5955
         TabIndex        =   52
         Top             =   630
         Width           =   3855
      End
      Begin VB.Label lblҽ�Ƹ��� 
         AutoSize        =   -1  'True
         Caption         =   "���ʽ"
         Height          =   240
         Left            =   6825
         TabIndex        =   50
         Top             =   255
         Width           =   960
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   240
         Left            =   135
         TabIndex        =   46
         Top             =   660
         Width           =   960
      End
      Begin VB.Label lblPatient 
         AutoSize        =   -1  'True
         Caption         =   "����"
         ForeColor       =   &H80000007&
         Height          =   240
         Index           =   7
         Left            =   150
         TabIndex        =   42
         Top             =   255
         Width           =   480
      End
      Begin VB.Label lblSex 
         AutoSize        =   -1  'True
         Caption         =   "�Ա�"
         Height          =   240
         Left            =   3135
         TabIndex        =   41
         Top             =   255
         Width           =   480
      End
      Begin VB.Label lblOld 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   240
         Left            =   4800
         TabIndex        =   40
         Top             =   255
         Width           =   480
      End
      Begin VB.Label lbl�ѱ� 
         AutoSize        =   -1  'True
         Caption         =   "�ѱ�"
         Height          =   240
         Left            =   3165
         TabIndex        =   39
         Top             =   660
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   390
      Left            =   8010
      TabIndex        =   22
      ToolTipText     =   "�ȼ���F2"
      Top             =   6000
      Width           =   1500
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   390
      Left            =   8010
      TabIndex        =   23
      ToolTipText     =   "�ȼ�:Esc"
      Top             =   6450
      Width           =   1500
   End
   Begin MSCommLib.MSComm com 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label Label2 
      Caption         =   "�������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      TabIndex        =   56
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "0.0111"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   100
      TabIndex        =   55
      Top             =   310
      Width           =   1890
   End
End
Attribute VB_Name = "frmSimpleCharge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
'����������������������������������������������������������������������������������������������������������������������������������������
'��ڲ�����
Public mbytInState As Byte '0ִ��(�޸�)�շ�,1-����շѵ�,2-��������
Public mstrInNO As String '��mbytInState=0ʱ��Ч,���ڵ��ݺ�
Public mblnNOMoved As Boolean '�����ĵ����Ƿ��ں����ݱ���
Public mstrDelete As String '�鿴���˷ѵ��ݵĵǼ�ʱ��,Ϊ""��Ч
Public mstrPrivs As String
Public mlngModul As Long
'����������������������������������������������������������������������������������������������������������������������������������������
'���ݶ���
Private mrsUnit As ADODB.Recordset '��ѡ���ִ�п���
Private mrsInfo As New ADODB.Recordset '������Ϣ
Private mrs�ѱ� As ADODB.Recordset      '���÷ѱ����ÿ���
Private mrs������ As ADODB.Recordset    '����ҽ���ͻ�ʿ����
Private mrs�������� As ADODB.Recordset  '��ѡ�Ŀ�������

'�������
Private mobjBill As ExpenseBill '���õ��ݶ���
Private mobjBillDetail As BillDetail '���ݵ��շ�ϸĿ����
Private mobjBillIncome As BillInCome '�շ�ϸĿ��������Ŀ����
Private mcolDetails As Details '�������շ�ϸĿ����
Private mcolMoneys As BillInComes  '������Ŀ���ܼ���(��ʾ����ӡʱʹ��)

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
    ��Ŀ = 0
    Ӧ�ս�� = 1
    ʵ�ս�� = 2
    ִ�п��� = 3
    ���� = 4
End Enum

'�������
Private mblnHotKey As Boolean '�ֹ�����ʱ,�Ƿ�Ű��˱����ȼ�
Private mbln���ϼ� As Boolean
Private mstrCardNO As String
Private mblnKeyPress As Boolean
Private mblnDo As Boolean
Private mblnDrop As Boolean         '��KeyDown���ж�cbo�����˵�ǰ�Ƿ񵯳�
Private mblnCboClick As Boolean      '�����cbo��keypress�¼������˵����б��API����:sendmessage,�����ͣ��cbo��,����һ���ַ�,�ƿ�����򰴻س���,
'                                    cbo��ֵ�ᱣ������,�����ᴥ��click�¼�,������Ҫ��validate�¼��е���click�¼�
Private mobjICCard As Object
Private mblnNotClick As Boolean

Private mstrPreUnit As String
Private mblnValid As Boolean
Private mstr���ʽ As String
Private mlng����ID As Long
Private mintBillNO As Integer
Private mintMoneyRow As Integer '��ǰ��ʾ���ķ�Ŀ��
Private mbln������۸� As Boolean

Private mlngShareUseID As Long '������������ID
Private mstrUseType As String 'ʹ�����
Private mintInvoiceFormat As Integer  '��ӡ�ķ�Ʊ��ʽ,��Ʊ��ʽ���
Private mintOldInvoiceFormat As Integer '��Ʊ�ݴ�ӡ��ʽ
Private mintInvoicePrint As Integer '0-����ӡ;1-�Զ���ӡ;2-��ʾ��ӡ
Private mblnStartFactUseType As Boolean

'�շѴ�ͬһ���˲��˵����ۼƽ��
Private mstrPrePati As String  '��һ���շѲ���
Private mcurBillӦ�� As Currency
Private mcurBillʵ�� As Currency
Private mcurBillӦ�� As Currency

Private marrColData() As Integer '��ǰ���ݱ༭����ӳ��
Private mblnPrint As Boolean
Private mblnSelect As Boolean '���ڿ����շ�ϸĿ�����Ƿ��������б�ѡ���ѡ����
Private Const STR_HEAD = "��Ŀ,3000,1;Ӧ�ս��,1500,7;ʵ�ս��,1500,7;ִ�п���,1950,1;����,1000,1"
'-----------------------------------------------------------------------------------
'���㿨���
Private mstrPassWord As String
'-----------------------------------------------------------------------------------
Private mstrҩƷ�۸�ȼ� As String, mstr���ļ۸�ȼ� As String, mstr��ͨ�۸�ȼ� As String

Private Sub Bill_BeforeAddRow(Row As Long)
    Dim dbl����  As Double, curMoney As Currency, i As Integer
    'LED��̬��ʾ��Ŀ
    If gblnLED And mbytInState = 0 And mobjBill.Pages(1).Details.Count >= Row - 1 Then
        With mobjBill.Pages(1).Details(Row - 1)
            dbl���� = 0: curMoney = 0
            For i = 1 To .InComes.Count
                curMoney = curMoney + .InComes(i).ʵ�ս��
                dbl���� = dbl���� + .InComes(i).��׼����
            Next
            'LED��ʾ
            If curMoney <> 0 Then
                zl9LedVoice.Display .Detail.����, .Detail.���, .���㵥λ, dbl����, .����, curMoney
            End If
        End With
    End If
End Sub

Private Sub ShowGroupLED(ByVal lngMain As Long, ByVal lngBegin As Long, ByVal lngEnd As Long)
'���ܣ�Ϊ�ӿ��ٶȣ�һ���Ե����ײ���Ŀ��LED��ʾ
'�������кŷ�Χ��lngMain=�����к�,lngBegin-lngEnd:�����к�
    Dim dbl���� As Double, dbl���� As Double, cur��� As Currency
    Dim i As Long, j As Long
    
    If gblnLED Then
        With mobjBill.Pages(1)
            For j = 1 To .Details(lngMain).InComes.Count
                cur��� = cur��� + .Details(lngMain).InComes(j).ʵ�ս��
            Next
            For i = lngBegin To lngEnd
                For j = 1 To .Details(i).InComes.Count
                    cur��� = cur��� + .Details(i).InComes(j).ʵ�ս��
                Next
            Next
        End With
        With mobjBill.Pages(1).Details(lngMain)
            If cur��� <> 0 Then
                dbl���� = .����
                If dbl���� <> 0 Then
                    dbl���� = cur��� / dbl����
                Else
                    dbl���� = cur���
                End If
                zl9LedVoice.Display .Detail.����, .Detail.���, .���㵥λ, dbl����, dbl����, cur���
            End If
        End With
    End If
End Sub

Private Sub Bill_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    Dim i As Long, bytSubs As Byte
    
    If mobjBill.Pages(1).Details.Count >= Row Then
        If mobjBill.Pages(1).Details(Row).������ Then
            MsgBox "���в����޸ļ�ɾ����", vbInformation, gstrSysName
            Cancel = True: Exit Sub
        End If
    End If
    
    If mobjBill.Pages(1).Details.Count >= Row Then
        '��������Ŀ����ɾ��ȷ��
        For i = Row + 1 To mobjBill.Pages(1).Details.Count
            If mobjBill.Pages(1).Details(i).�������� = Row Then bytSubs = bytSubs + 1
        Next
        If bytSubs > 0 Then
            If MsgBox("����Ŀ���� " & bytSubs & " ��������Ŀ,ɾ������ĿҲ��ɾ�����Ĵ�����Ŀ,������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True: Exit Sub
            End If
        ElseIf mobjBill.Pages(1).Details(Row).�������� <> 0 Then '������Ŀɾ��ȷ��
            If MsgBox("����Ŀ��[" & mobjBill.Pages(1).Details(mobjBill.Pages(1).Details(Row).��������).Detail.���� & "]�Ĵ�����Ŀ,ȷ��Ҫɾ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True: Exit Sub
            End If
        ElseIf MsgBox("ȷʵҪɾ�����շ���Ŀ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = True: Exit Sub
        End If
        
        'ɾ������
        For i = mobjBill.Pages(1).Details.Count To Row + 1 Step -1
            If mobjBill.Pages(1).Details(i).�������� = Row Then
                Call DeleteDetail(i) '��˳��ɾ���������
            End If
        Next
        Call DeleteDetail(Row) 'ɾ������
        
        '���¼��㲢ˢ��
        'Call CalcMoneys
        Call ShowDetails
        Call ShowMoney
        
        If mobjBill.Pages(1).Details.Count = 0 Then ClearMoney
        
        Bill.TxtVisible = False
        Bill.CmdVisible = False
        Bill.CboVisible = False
        
        Cancel = True '���ÿؼ���������
    End If
End Sub

Private Sub Bill_cboClick(ListIndex As Long)
    Dim lngִ�п��� As Long, strִ�п��� As String
    If mobjBill.Pages(1).Details.Count >= Bill.Row Then
        If Bill.ListIndex <> -1 Then
            If mobjBill.Pages(1).Details(Bill.Row).ִ�в���ID <> Bill.ItemData(Bill.ListIndex) Then
                lngִ�п��� = mobjBill.Pages(1).Details(Bill.Row).ִ�в���ID: strִ�п��� = Bill.TextMatrix(Bill.Row, Bill.Col)
                mobjBill.Pages(1).Details(Bill.Row).ִ�в���ID = Bill.ItemData(Bill.ListIndex)
                If ItemHaveSub(Bill.Row) Then Call SetSubItemDept(Bill.Row)
                If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModul, 0, 0, _
                    MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), 1, 0, 1, Bill.Row)) = False Then
                    Bill.Text = "": Bill.TxtVisible = False
                    Bill.cboObj.Text = strִ�п���: mobjBill.Pages(1).Details(Bill.Row).ִ�в���ID = lngִ�п���: Exit Sub
                End If
            End If
        End If
    End If
End Sub

Private Sub Bill_CommandClick()
    Dim lng��Ŀid As Long, blnCancel As Boolean
        
    lng��Ŀid = frmItemSelect.ShowSelect(Me, mstrPrivs, gint������Դ, 0, False, "'Z'", , , _
        , , , , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ�۸�ȼ�)
    If lng��Ŀid <> 0 Then
        Bill.Text = lng��Ŀid
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

Private Sub Bill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
'���ܣ�����������
    Dim strScope As String, i As Long
    Dim objDetail As Detail, lng��Ŀid As Long, lngDoUnit As Long
    
    If KeyCode = 13 And Not Bill.Active Then
        Cancel = True: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    End If
        
    On Error GoTo errH
    
    If KeyCode = 13 Then
        '�շ�ʱ,�����Ѳ����޸�
        If mobjBill.Pages(1).Details.Count >= Bill.Row Then
            If mobjBill.Pages(1).Details(Bill.Row).������ Then Exit Sub
        End If
        If Bill.ColData(Bill.Col) = 0 Then Exit Sub
        
        Select Case Bill.TextMatrix(0, Bill.Col)
            Case "��Ŀ"
                '����Ŀȷ��,���շ�ϸĿ��Ӧ�ĳ�����������,ͬʱ���ﴦ���շѴ�����Ŀ
                If Bill.Text <> "" Then
                    If mblnSelect Then
                        mblnSelect = False '��������ñ�־
                        Set objDetail = GetInputDetail(Val(Bill.Text))
                    Else
                        lng��Ŀid = frmItemSelect.ShowSelect(Me, mstrPrivs, gint������Դ, 0, False, "'Z'", Bill.Text, Bill.TxtHwnd, _
                                        , , , , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ�۸�ȼ�)
                        If lng��Ŀid <> 0 Then
                            Set objDetail = GetInputDetail(lng��Ŀid)
                        Else
                            Bill.Text = "": Bill.TxtVisible = False
                            Bill.SetFocus: Cancel = True: Exit Sub
                        End If
                    End If
                    
                    sta.Panels(2) = ""
                    Bill.TxtVisible = False '(���Ӳ���)
                    '������޸ĸ��շ�ϸĿ��
                    Call SetDetail(objDetail, Bill.Row)
                    
                    '����ժҪ(������������и���ժҪ)
                    Dim strժҪ As String '90304
                    If mobjBill.Pages(1).Details(Bill.Row).Detail.����ժҪ Then
                        If frmInputBox.InputBox(Me, "ժҪ", "������""" & mobjBill.Pages(1).Details(Bill.Row).Detail.���� & """��ժҪ��Ϣ:", 200, 3, True, False, strժҪ) Then
                            mobjBill.Pages(1).Details(Bill.Row).ժҪ = strժҪ
                        End If
                    Else
                         strժҪ = gclsInsure.GetItemInfo(0, mobjBill.����ID, mobjBill.Pages(1).Details(Bill.Row).�շ�ϸĿID, strժҪ, 1)
                         mobjBill.Pages(1).Details(Bill.Row).ժҪ = strժҪ
                    End If
                    
                    Call CalcMoneys(Bill.Row)
                    
                    If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModul, 0, 0, _
                        MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), 1, 0, 1, Bill.Row)) = False Then
                        mobjBill.Pages(1).Details.Remove Bill.Row 'ɾ���ո���Ҫ����ķ�����
                        Bill.Text = "": Bill.TxtVisible = False
                        Cancel = True: Exit Sub
                    End If
                    
                    Call ShowDetails(Bill.Row)
                    Call ShowMoney
                    
                    '�������ͼ��
                    Call Check��������(Bill.Row)
                    
                    Bill.Text = "": Bill.SetFocus
                End If
                
                If mobjBill.Pages(1).Details.Count >= Bill.Row Then
                    '��һ�е�����ȷ��
                    If mobjBill.Pages(1).Details(Bill.Row).Detail.��� Then Bill.ColData(1) = 4 'Ӧ�ս��
                    
                    'ִ�п���!!!
                    Call FillBillComboBox(Bill.Row, 3)
                    If Bill.ListCount = 1 Then
                        Bill.ColData(3) = 5
                        mobjBill.Pages(1).Details(Bill.Row).Key = 1
                    Else
                        Bill.ColData(3) = 3
                        mobjBill.Pages(1).Details(Bill.Row).Key = Bill.ListCount
                    End If
                    
                    '������Ŀ����(��������Դ���༶����-�����Ĵ���...)
                    If Bill.TextMatrix(0, Bill.Col) = "��Ŀ" Then
                        If ShouldDO(Bill.Row) Then
                            Set mcolDetails = New Details
                            Set mcolDetails = GetSubDetails(mobjBill.Pages(1).Details(Bill.Row).�շ�ϸĿID)
                            For i = 1 To mcolDetails.Count
                                If mobjBill.Pages(1).Details.Count >= Bill.Rows - 1 Then
                                    Bill.Rows = Bill.Rows + 1
                                    Call bill_AfterAddRow(Bill.Rows - 1)
                                End If
                                Bill.TextMatrix(Bill.Rows - 1, 0) = "" '�б�Ҫ����
                                
                                If mcolDetails(i).��� = mobjBill.Pages(1).Details(Bill.Row).�շ���� Then
                                    '1.�����շ������������ͬ��,ȱʡ������ִ�п�����ͬ��
                                    lngDoUnit = mobjBill.Pages(1).Details(Bill.Row).ִ�в���ID
                                Else
                                    If mcolDetails(i).ִ�п��� = 0 Then
                                        '2.��������Ϊ����ȷ���ҵ�,ȱʡ������ִ�п�����ͬ��
                                        lngDoUnit = mobjBill.Pages(1).Details(Bill.Row).ִ�в���ID
                                    End If
                                        '�������,ȡ�������õ�ִ�п���
                                End If
                                            
                                Call SetDetail(mcolDetails(i), Bill.Rows - 1, Bill.Row, lngDoUnit)
                                Call CalcMoneys(Bill.Rows - 1)
                                Call ShowDetails(Bill.Rows - 1)
                                Call ShowMoney
                            Next
                            'һ���Ե����ײ���ĿLED��ʾ
                            Call ShowGroupLED(Bill.Row, Bill.Rows - mcolDetails.Count, Bill.Rows - 1)
                        End If
                    End If
                End If
            Case "Ӧ�ս��" 'ʵ�����ǵ���(��Ϊ���ݴ�ȱʡΪ1,�Ҳ��ܸ���)
                If mobjBill.Pages(1).Details.Count >= Bill.Row And Bill.Text <> "" Then
                    '���ֺϷ���
                    If Not IsNumeric(Bill.Text) Then
                        MsgBox "�Ƿ���ֵ��", vbInformation, gstrSysName
                        Bill.Text = "": Cancel = True: Bill.TxtVisible = False: Exit Sub
                    End If
                    '����Ȩ��
                    If InStr(mstrPrivs, "��������") = 0 And CDbl(Bill.Text) < 0 Then
                        MsgBox "��û��Ȩ�����븺����", vbInformation, gstrSysName
                        Bill.Text = "": Cancel = True: Bill.TxtVisible = False: Exit Sub
                    End If
                    
                    Bill.Text = Format(Bill.Text, gstrDec)
                    
                    If mobjBill.Pages(1).Details.Count >= Bill.Row And Bill.Text <> "" Then
                        '���û�ж�Ӧ��������Ŀ,���޷�����
                        If mobjBill.Pages(1).Details(Bill.Row).Detail.��� And mobjBill.Pages(1).Details(Bill.Row).InComes.Count > 0 Then
                            If Not (mobjBill.Pages(1).Details(Bill.Row).InComes(1).�ּ� = 0 And mobjBill.Pages(1).Details(Bill.Row).InComes(1).ԭ�� = 0) Then
                                strScope = CheckScope(mobjBill.Pages(1).Details(Bill.Row).InComes(1).ԭ��, mobjBill.Pages(1).Details(Bill.Row).InComes(1).�ּ�, CCur(Bill.Text))
                                If strScope <> "" Then
                                    sta.Panels(2) = strScope
                                    If Bill.TxtVisible And Len(Bill.Text) > 9 Then Bill.Text = mobjBill.Pages(1).Details(Bill.Row).InComes(1).��׼����
                                    If Bill.TxtVisible Then Bill.SelStart = 0: Bill.SelLength = Len(Bill.Text)
                                    Cancel = True: Beep: Exit Sub
                                End If
                            End If
                            
                            '�����շ�ϸĿֻ�ܶ�Ӧһ��������Ŀ
                            mobjBill.Pages(1).Details(Bill.Row).���� = Sgn(Val(Bill.Text))
                            mobjBill.Pages(1).Details(Bill.Row).InComes(1).��׼���� = Abs(Val(Bill.Text))
                            
                            Call CalcMoneys(Bill.Row)
                            Call ShowDetails(Bill.Row)
                            Call ShowMoney
                        Else
                            Bill.Text = "0"
                            sta.Panels(2) = "����Ŀ�������ö�Ӧ�ķ�Ŀ�������޷�������ã�"
                            Beep
                        End If
                    End If
                End If
            Case "ִ�п���"
                If mobjBill.Pages(1).Details.Count >= Bill.Row Then
                   If Bill.ListIndex <> -1 Then
                        'If mobjBill.Pages(1).Details(Bill.Row).ִ�в���ID <> Bill.ItemData(Bill.ListIndex) Then
                            mobjBill.Pages(1).Details(Bill.Row).ִ�в���ID = Bill.ItemData(Bill.ListIndex)
                            If ItemHaveSub(Bill.Row) Then Call SetSubItemDept(Bill.Row)
                        'End If
                        If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModul, 0, 0, _
                            MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), 1, 0, 1, Bill.Row)) = False Then
                            Bill.Text = "": Bill.TxtVisible = False
                            Cancel = True: Exit Sub
                        End If
                    End If
                End If
        End Select
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Cancel = True
End Sub

Private Sub SetSubItemDept(ByVal lngRow As Long)
'����:��������ִ�п��ҵı仯,ˢ�·�ҩ�����ִ�п���
    Dim i As Long, j As Long, lng���˿���ID As Long
    
    lng���˿���ID = mobjBill.����ID
    If lng���˿���ID = 0 And cbo��������.ListIndex <> -1 Then lng���˿���ID = cbo��������.ItemData(cbo��������.ListIndex)

    With mobjBill.Pages(1)
        For i = lngRow + 1 To .Details.Count
            If .Details(i).�������� = lngRow Then
                '������ΪҩƷ�����ĵ���Ŀ��ִ�п��Ҳ�������䶯
                If InStr(",4,5,6,7,", .Details(i).�շ����) = 0 Then
                    If .Details(i).�շ���� = .Details(lngRow).�շ���� Then
                        '1.�����շ������������ͬ��,ȱʡ������ִ�п�����ͬ��
                        .Details(i).ִ�в���ID = .Details(lngRow).ִ�в���ID
                    Else
                        Set mcolDetails = GetSubDetails(.Details(lngRow).�շ�ϸĿID) '������ȡ
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
                    
                    If .Details(i).ִ�в���ID > 0 Then
                        If mbytInState = 0 Then
                            mrsUnit.Filter = "ID=" & .Details(i).ִ�в���ID
                            If mrsUnit.RecordCount <> 0 Then
                                Bill.TextMatrix(i, 3) = mrsUnit!���� & "-" & mrsUnit!����
                            Else
                                Bill.TextMatrix(i, 3) = GET��������(.Details(i).ִ�в���ID, mrsUnit)
                            End If
                        Else
                            '�������ֻ(��)��ʾ����
                            Bill.TextMatrix(i, 3) = GET��������(.Details(i).ִ�в���ID, mrsUnit)
                        End If
                    End If
                End If
            End If
        Next
    End With

End Sub

Private Sub Set�����˿�������(ByVal str������ As String, ByVal lng��������ID As Long)
'����:���ݿ����˻򿪵�����ID���ÿ������Ҽ�������,������������¼�
       '���ù�������CboSetIndex������ʽ����cbo_click�¼�
    
    Dim str�������� As String, lng��ԱID As Long
    
    'a.ҽ��ȷ������
    If gbyt����ҽ�� = 0 Then
        Call zlControl.CboSetIndex(cbo������.hWnd, cbo.FindIndex(cbo������, str������, True)) '������click�¼�
        
        If cbo������.ListIndex = -1 And str������ <> "" Then
            lng��ԱID = GetPersonnelID(str������, mrs������)
            cbo������.AddItem str������, 0
            cbo������.ItemData(cbo������.NewIndex) = lng��ԱID
            Call zlControl.CboSetIndex(cbo������.hWnd, cbo������.NewIndex)
        End If
                
        If cbo������.ListIndex <> -1 Then
            cbo��������.Clear
            Call FillDept(cbo������.ItemData(cbo������.ListIndex))
        End If
        
        Call zlControl.CboSetIndex(cbo��������.hWnd, cbo.FindIndex(cbo��������, lng��������ID))
        If cbo��������.ListIndex = -1 And lng��������ID > 0 Then
            str�������� = GET��������(lng��������ID)
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
            str�������� = GET��������(lng��������ID)
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
End Sub

Private Sub Set�����˿�������Click(ByVal str������ As String, ByVal lng��������ID As Long)
'����:���ݿ����˻򿪵�����ID���ÿ������Ҽ�������,����������¼�
'     ��Listindex=xʱ,���Listindex��ֵ�������x,�Ͳ��ᴥ������¼�,����Ҫ��API+Clickǿ�Ƶ���
    Dim i As Long
    
    If gbyt����ҽ�� = 0 Then
        Call zlControl.CboSetIndex(cbo������.hWnd, cbo.FindIndex(cbo������, str������, True)) '������click�¼�
        Call cbo������_Click
        
        Call zlControl.CboSetIndex(cbo��������.hWnd, cbo.FindIndex(cbo��������, lng��������ID))
        Call cbo��������_Click
        
    Else
        '����ȷ��ҽ������Զ�������
        Call zlControl.CboSetIndex(cbo��������.hWnd, cbo.FindIndex(cbo��������, lng��������ID))
        Call cbo��������_Click
        
        Call zlControl.CboSetIndex(cbo������.hWnd, cbo.FindIndex(cbo������, str������, True)) '������click�¼�
        Call cbo������_Click
    End If
End Sub

Private Function ItemHaveSub(ByVal lngRow As Long) As Boolean
'���ܣ��жϵ�ǰ�е���Ŀ�Ƿ���д�����Ŀ
    Dim i As Long
    
    If mobjBill.Pages(1).Details.Count >= lngRow Then
        For i = lngRow + 1 To mobjBill.Pages(1).Details.Count
            If mobjBill.Pages(1).Details(i).�������� = lngRow Then
                ItemHaveSub = True: Exit Function
            End If
        Next
    End If
End Function

Private Sub Bill_EnterCell(Row As Long, Col As Long)
    Dim i As Integer, bln������ As Boolean
    
    If Not Bill.Active Then Exit Sub
    
    '�ָ��б༭����
    If mbytInState = 0 Then
        For i = 0 To UBound(marrColData)
            Bill.ColData(i) = marrColData(i)
        Next
    End If
    
    '�շ�ʱ,���Ϊ������,�����޸�
    If mobjBill.Pages(1).Details.Count >= Row And mbytInState = 0 Then
        If mobjBill.Pages(1).Details(Row).������ Then
            bln������ = True
            For i = 0 To UBound(marrColData)
                Bill.ColData(i) = IIf(marrColData(i) = 5, 5, 0)
            Next
        End If
    End If
    
    '����Ǵ�����Ŀ������Ŀ�����,���������������Ŀ
    If mobjBill.Pages(1).Details.Count >= Row Then
        If ItemHaveSub(Row) Or mobjBill.Pages(1).Details(Row).�������� > 0 Then
            Bill.ColData(0) = BillColType.Text_UnModify
        End If
    End If
    
    'ִ�п�����
    If mobjBill.Pages(1).Details.Count >= Bill.Row And mbytInState <> 2 And Not bln������ Then
        If mobjBill.Pages(1).Details(Bill.Row).Key = "1" Then
            Bill.ColData(3) = 5
        Else
            Bill.ColData(3) = 3
        End If
    End If
    If Bill.ColData(Bill.Col) = 3 Then Call FillBillComboBox(Bill.Row, Bill.Col)
    
    Bill.TextLen = 0: Bill.TextMask = ""
    Select Case Bill.TextMatrix(0, Col)
        Case "ִ�п���"
            SetWidth Bill.cboHwnd, 130
        Case "Ӧ�ս��"
            Bill.TextLen = 10
            If InStr(mstrPrivs, "��������") = 0 Then
                Bill.TextMask = "0123456789." & Chr(8)
            Else
                Bill.TextMask = "-0123456789." & Chr(8)
            End If
    End Select

    '������ʱ,�������ø��еı༭����
    If mobjBill.Pages(1).Details.Count >= Bill.Row And Not bln������ Then
        If mobjBill.Pages(1).Details(Bill.Row).Detail.��� Then
            Bill.ColData(1) = 4
        Else
            Bill.ColData(1) = 5
        End If
    End If
End Sub

Private Sub Bill_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Bill.ToolTipText = Bill.TextMatrix(Bill.MouseRow, Bill.MouseCol)
End Sub

Private Sub cboSex_Click()
    If mbytInState = 0 Then
        mobjBill.�Ա� = zlStr.NeedName(cboSex.Text)
    End If
End Sub

Private Sub cbo�ѱ�_Click()
    If cbo�ѱ�.ListIndex <> -1 Then
        If mobjBill.�ѱ� <> zlStr.NeedName(cbo�ѱ�.Text) And Not mbln������۸� Then
            mobjBill.�ѱ� = zlStr.NeedName(cbo�ѱ�.Text)
            
            If mbytInState = 0 And mobjBill.Pages(1).Details.Count > 0 Then
                '���¼���۸�
                Call CalcMoneys
                Call ShowDetails
                Call ShowMoney
            End If
        End If
    End If
End Sub

Private Sub cbo���㷽ʽ_Click()
'���ܣ����ֽ�����ֽ�֮���л�ʱ����Ҫ������������Ƿ���ֱ�
    Dim dblTemp As Double
    
    If Not (Visible And gBytMoney <> 0) Then Exit Sub
    If Bill.Active Then
        Call ShowMoney
    ElseIf chkCancel.Value = 1 Then
        txtӦ��.Text = Format(GetDelMoney, "0.00")
        
        '�����ʾ
        dblTemp = -1 * Format((Val(txt�ϼ�.Text) - Val(txtԤ�����.Text) - Val(txtӦ��.Text)), gstrDec)
        If dblTemp <> 0 Then
            pic���.Visible = True
            lbl����.Caption = Format(dblTemp, "0.00")
        Else
            pic���.Visible = False
        End If
    Else
        Call ShowPrice
    End If
End Sub

Private Function GetDelMoney() As Currency
    Dim cur�˷Ѻϼ� As Currency
    Dim bln�ֽ���� As Boolean
    Dim i As Integer
    
    cur�˷Ѻϼ� = Format(Val(txt�ϼ�.Text) - Val(txtԤ�����.Text), "0.00")
    
    '�ֽ����ʱ����ֱ�(��Ϊ�����˷�ʱ������ҽ���ӿ�,��˲���ҽ���Ƿ�֧�ֱַ�)
    bln�ֽ���� = False
    If cbo���㷽ʽ.ListIndex <> -1 Then
        If cbo���㷽ʽ.ItemData(cbo���㷽ʽ.ListIndex) = 1 Then
            bln�ֽ���� = True
        End If
    End If
    If bln�ֽ���� Then
        cur�˷Ѻϼ� = CentMoney(Val(txt�ϼ�.Text) - Val(txtԤ�����.Text))
    End If
    GetDelMoney = cur�˷Ѻϼ�
End Function

Private Sub cbo���㷽ʽ_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    
    If KeyAscii >= 32 Then
        If cbo���㷽ʽ.Locked Then Exit Sub
        
        lngIdx = zlControl.CboMatchIndex(cbo���㷽ʽ.hWnd, KeyAscii)
        If lngIdx = -1 And cbo���㷽ʽ.ListCount > 0 Then lngIdx = 0
        cbo���㷽ʽ.ListIndex = lngIdx
    ElseIf KeyAscii = 13 Then
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    End If
End Sub

Private Sub cbo��������_Click()
    Dim i As Long, lng��������ID As Long
    
    mblnCboClick = True
    If Not (mbytInState = 0 And chkCancel.Value = 0) Then Exit Sub
        
    If cbo��������.ListIndex <> -1 Then lng��������ID = cbo��������.ItemData(cbo��������.ListIndex)
    If mobjBill.Pages(1).��������ID = lng��������ID Then Exit Sub
    mobjBill.Pages(1).��������ID = lng��������ID
    
    '��λҽ��
    If gbyt����ҽ�� = 1 Then
        If cbo��������.ListIndex <> -1 Then
            Call FillDoctor(lng��������ID)
            
            If cbo������.ListCount > 0 And Not gbln��ȱʡ������ Then
                Call zlControl.CboSetIndex(cbo������.hWnd, 0)
            End If
        Else
            cbo������.Clear
        End If
        Call cbo������_Click
    End If
    
    
    '���ݿ����������������շ���Ŀ��ִ�п���
    If cbo��������.ListIndex <> -1 And Visible Then
        With mobjBill.Pages(1)
        For i = 1 To .Details.Count
            If InStr(",4,5,6,7,", .Details(i).Detail.���) = 0 And _
             (.Details(i).Detail.ִ�п��� = 6 And gbyt����ҽ�� <> 2 Or InStr(",1,2,", "," & .Details(i).Detail.ִ�п��� & ",") > 0 And gint������Դ = 1) Then '6-�����˿���
                
                .Details(i).ִ�в���ID = lng��������ID
                
                If i <= Bill.Rows - 1 And .Details(i).ִ�в���ID <> 0 Then
                    If mbytInState = 0 Then
                        mrsUnit.Filter = "ID=" & .Details(i).ִ�в���ID
                        If mrsUnit.RecordCount <> 0 Then
                            Bill.TextMatrix(i, BillCol.ִ�п���) = mrsUnit!���� & "-" & mrsUnit!����
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
        Next
        End With
    End If
        
    '�ѱ���
    Call LoadAndSeek�ѱ�
End Sub


Private Sub cbo��������_Validate(Cancel As Boolean)
 '�����cbo��keypress�¼������˵����б��API����:sendmessage,�����ͣ��cbo��,����һ���ַ�,�ƿ�����򰴻س���,
'                                    cbo��ֵ�ᱣ������,�����ᴥ��click�¼�,������Ҫ��validate�¼��е���click�¼�

    If Not mblnCboClick Then cbo��������_Click
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
    If mobjBill.Pages(1).������ = IIf(cbo������.ListIndex = -1, "", zlStr.NeedName(cbo������.Text)) Then Exit Sub
    
    mobjBill.Pages(1).������ = IIf(cbo������.ListIndex = -1, "", zlStr.NeedName(cbo������.Text))
    '��ҽ��ȷ������
    If gbyt����ҽ�� = 0 Then
        If cbo������.ListIndex <> -1 Then
            lng������ID = cbo������.ItemData(cbo������.ListIndex)
            
            Call FillDept(lng������ID)
            Call SetDefaultDept(lng������ID)
        Else
            cbo��������.Clear
        End If
        Call cbo��������_Click
    End If
    
    '����ҽ������,��Ϊ�����˱��ˣ�����,ִ�п������ɿ����˿��Ҿ���ʱ����Ҫ����ִ�п���
     '������ʱ��Cbo��������_click�д���
    If cbo������.ListIndex <> -1 And Visible And gbyt����ҽ�� = 2 Then
        With mobjBill.Pages(1)
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
                            Bill.TextMatrix(i, BillCol.ִ�п���) = mrsUnit!���� & "-" & mrsUnit!����
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
        Next
        End With
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

Private Sub cbo���䵥λ_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo���䵥λ_Validate(Cancel As Boolean)
    If mbytInState = 0 Then mobjBill.���� = Trim(txt����.Text) & IIf(IsNumeric(txt����.Text), cbo���䵥λ.Text, "")
End Sub

Private Sub cboҽ�Ƹ���_Click()
    On Error GoTo errHandler
    If mbytInState <> 0 Then Exit Sub
    If gintPriceGradeStartType < 2 Then Exit Sub
    
    If mrsInfo.State = adStateOpen Then
        Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, Val(Nvl(mrsInfo!����ID)), Val(Nvl(mrsInfo!��ҳID)), zlStr.NeedName(cboҽ�Ƹ���.Text), mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ�۸�ȼ�)
    Else
        Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, 0, 0, zlStr.NeedName(cboҽ�Ƹ���.Text), mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ�۸�ȼ�)
    End If
    
    If mbln������۸� Then Exit Sub
    If mobjBill.Pages(1).Details.Count = 0 Then Exit Sub
    
    '���¼���۸�
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
                If cbo���㷽ʽ.Enabled Then cbo���㷽ʽ.SetFocus
            Else
                If cbo��������.Enabled Then cbo��������.SetFocus
            End If
        End If
    End If
End Sub

Private Sub chkCancel_Click()
    mstrInNO = ""
    txt�Ҳ�.Text = "0.00": txt�ɿ�.Text = "0.00": txtӦ��.Text = "0.00"
    mcurBillʵ�� = 0: mcurBillӦ�� = 0: mcurBillӦ�� = 0
    mstrPrePati = "": mintBillNO = 0: mintMoneyRow = 0
    Call ClearPatientInfo
    txt�ϼ�.Text = gstrDec: txtӦ��.Text = gstrDec
    
    If chkCancel.Value = Checked Then
        chkCancel.ForeColor = &HFF&
        Call ClearRows: Call Bill.ClearBill
        Call ClearMoney
        Call NewBill(False)
        Call SetDisible
        
        cboNO.Text = "": cboNO.Locked = False
        txtInvoice.Text = "": txtInvoice.Locked = True
        txtRePrint.Locked = True
        
        lblӦ��.Caption = "Ӧ�˽��"
        lblӦ��.ForeColor = vbRed
        txtӦ��.ForeColor = vbRed
        txtӦ��.Text = "0.00"
        
        cboNO.SetFocus
    Else
        chkCancel.ForeColor = 0
        txtInvoice.Locked = Not (InStr(1, mstrPrivs, "�޸�Ʊ�ݺ�") > 0) And gblnStrictCtrl
        txtRePrint.Locked = False
        
        lblӦ��.Caption = "Ӧ�ɽ��"
        lblӦ��.ForeColor = 0
        txtӦ��.ForeColor = &HFF0000
        txtӦ��.Text = "0.00"
        
        Call ClearRows: Call Bill.ClearBill
        Call ClearMoney
        Call NewBill
        Call SetDisible(True)
        txtPatient.SetFocus
    End If
End Sub

Private Sub chk�Ӱ�_Click()
    If Not mblnDo Then Exit Sub
    If mbytInState = 1 Or chkCancel.Value = 1 Then Exit Sub
    If Not chk�Ӱ�.Visible Or Not Visible Then Exit Sub
    
    Dim blnAdd As Boolean
    
    
    blnAdd = OverTime(zlDatabase.Currentdate)
    If chk�Ӱ�.Value = Unchecked And blnAdd Then
        If MsgBox("��ǰ���ڼӰ�ʱ�䷶Χ��,Ҫȡ���Ӱ�Ӽ���", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            chk�Ӱ�.Value = Checked
        End If
    End If
    If chk�Ӱ�.Value = Checked And Not blnAdd Then
        If MsgBox("��ǰ�����ڼӰ�ʱ�䷶Χ��,Ҫִ�мӰ�Ӽ���", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            chk�Ӱ�.Value = Unchecked
        End If
    End If
    mobjBill.�Ӱ��־ = IIf(chk�Ӱ�.Value = Checked, 1, 0)
    
    '���¼���۸�
    If Not mobjBill.Pages(1).Details.Count = 0 Then
        Call CalcMoneys
        Call ShowDetails
        Call ShowMoney
    End If
End Sub

Private Sub chk�Ӱ�_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdCancel_Click()
    If (mobjBill.Pages(1).Details.Count > 0 Or txtPatient.Text <> "") And Bill.Active And mbytInState = 0 And mstrInNO = "" Then
        If MsgBox("ȷʵҪ�����ǰ�����е�������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
    
        txt�Ҳ�.Text = "0.00"
        txt�ɿ�.Text = "0.00"
        txtӦ��.Text = "0.00"
        If chkCancel.Value = Checked Then '�˾ݵ�״̬
            Call ClearRows: Call Bill.ClearBill
            Call ClearMoney
            chkCancel.Value = Unchecked
            Call NewBill
            Call SetDisible(True)
            If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        ElseIf Bill.Active Then '�������뵥��״̬'(����������²��˵���)
            mcurBillʵ�� = 0:  mcurBillӦ�� = 0: mcurBillӦ�� = 0
            mstrPrePati = "": mintBillNO = 0: mintMoneyRow = 0
            Call ClearPatientInfo
            txt�ϼ�.Text = gstrDec: txtӦ��.Text = gstrDec
            If gbln�ۼ� Then txt�ۼ�.Text = Format(GetChargeTotal, "0.00")
            Call ClearRows: Call Bill.ClearBill
            Call ClearMoney
            Call NewBill '����ԭ���ݺ�
            If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        ElseIf Not Bill.Active Then '��ȡ���۵�����״̬
            Call ClearPatientInfo
            txt�ϼ�.Text = gstrDec: txtӦ��.Text = gstrDec
            Call ClearRows: Call Bill.ClearBill
            Call ClearMoney
            Call NewBill
            Call SetDisible(True)
            If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        End If
        Exit Sub
    End If
    
    Unload Me
End Sub

Private Sub ClearPatientInfo()
    txtPatient.Text = "": txtPatient.Tag = ""
    txt����.Text = ""
    Call zlControl.CboLocate(cbo���䵥λ, "��")
    Call txt����_Validate(False)
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

Private Sub cmdOK_Click()
    Dim strInfo As String, strSQL As String
    Dim i As Long, j As Long, lng����ID As Long
    Dim curMoney As Currency, cur������ As Currency
    Dim strҽ�Ƹ��� As String
    Dim str����Nos As String, rsItems As ADODB.Recordset
    
    If Val(txt�ɿ�.Text) = 0 Then txt�ɿ�.Text = "0.00"
    
    If mbytInState = 2 Then
        If Not IsDate(txtDate.Text) Then
            MsgBox "������Ϸ��ķ���ʱ�䣡", vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
        If Not SaveModi() Then Exit Sub
        Unload Me
    ElseIf Bill.Active Then '�������뵥��״̬
        Call txt�ɿ�_GotFocus
        
        If txtPatient.Text = "" And mobjBill.���� = "" Then
            MsgBox "û�з��ֲ�����Ϣ,�����벡����Ϣ��", vbInformation, gstrSysName
            txtPatient.SetFocus: Exit Sub
        Else
            If mobjBill.���� = "" Then
                mobjBill.���� = txtPatient.Text
            Else
                txtPatient.Text = mobjBill.����
            End If
        End If
        
        
        If CheckTextLength("����", txtPatient) = False Then Exit Sub
        If CheckTextLength("����", txt����) = False Then Exit Sub
        If Not CheckOldData(txt����, cbo���䵥λ) Then Exit Sub
        
        If cbo�ѱ�.ListIndex = -1 Or mobjBill.�ѱ� = "" Then
            MsgBox "��ѡ���˷ѱ�", vbInformation, gstrSysName
            If cbo�ѱ�.Enabled And cbo�ѱ�.Visible Then cbo�ѱ�.SetFocus: Exit Sub
        End If
        If mobjBill.Pages(1).Details.Count = 0 Then
            MsgBox "������û���κ�����,����ȷ���뵥�����ݣ�", vbInformation, gstrSysName
            Bill.SetFocus: Exit Sub
        End If
        
        i = Checkִ�п���
        If i <> 0 Then
            MsgBox "�����е� " & i & " ����Ŀû��ָ��ִ�п��ң�", vbInformation, gstrSysName
            Bill.SetFocus: Exit Sub
        End If
        
        If cbo��������.ListIndex = -1 Then
            MsgBox "��ȷ���������ң�", vbInformation, gstrSysName
            If gbyt����ҽ�� = 0 Then
                cbo������.SetFocus
            Else
                cbo��������.SetFocus
            End If
            Exit Sub
        End If
        
        '������
        If gbln�����俪���� And cbo������.ListIndex = -1 Then
            MsgBox "�����뿪���ˣ�", vbInformation, gstrSysName
            cbo������.SetFocus: Exit Sub
        End If
        
        If cbo���㷽ʽ.ListIndex = -1 Then
            MsgBox "��ȷ���շѵĽ��㷽ʽ��", vbInformation, gstrSysName
            cbo���㷽ʽ.SetFocus: Exit Sub
        End If
    
        If Not IsDate(txtDate.Text) Then
            MsgBox "��������ȷ�ķ������ڣ�", vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
        If Val(txt�ɿ�.Text) <> 0 And txt�ɿ�.Enabled Then
            If Val(txt�ɿ�.Text) < Val(txtӦ��.Text) Then
                MsgBox "���˽ɿ���㣬�벹��Ӧ�ɽ�", vbInformation, gstrSysName
                Call zlControl.TxtSelAll(txt�ɿ�): txt�ɿ�.SetFocus: Exit Sub
            End If
        End If
        '���˺�:22343,�ɿ������
        Select Case gTy_Module_Para.byt�ɿ����
        Case 1  '1-��������ɿ��Ž��������ۼ�
        Case 2  '2-�շ�ʱ����Ҫ����ɿ���
            If Val(txtӦ��.Text) > 0 And Val(txt�ɿ�.Text) = 0 Then
                MsgBox "ע��:" & vbCrLf & _
                "    �ò���δ����ɿ���,���ܽ����շ�!", vbInformation + vbDefaultButton1, gstrSysName
                If txt�ɿ�.Enabled And txt�ɿ�.Visible Then txt�ɿ�.SetFocus
                Exit Sub
            End If
        Case Else   ',0-�������нɿ�������ۼƿ���
        End Select
                
                
        '�Ƿ���
        For i = 1 To mobjBill.Pages(1).Details.Count
            If mobjBill.Pages(1).Details(i).�շ�ϸĿID = 0 Then
                MsgBox "�����е� " & i & " ��û����ȷ��������,��������ɾ�����У�", vbInformation, gstrSysName
                Bill.SetFocus: Exit Sub
            End If
        Next
        
        '�������ͼ��
        If Not Check�������� Then Exit Sub
                
        '���ŵ�����߶�
        If gcurMax <> 0 And CalcBillToTal > gcurMax Then
            MsgBox "���ݽ���������ƽ��:" & Format(gcurMax, "0.00") & " ,�������棡", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If CheckChargeItemByPlugIn(gobjPlugIn, glngSys, mlngModul, 0, 1, _
            MakeDetailRecord(mobjBill, zlStr.NeedName(cbo������.Text), zlStr.NeedName(cbo��������.Text), 1, 0)) = False Then
            Exit Sub
        End If
        
        'Ʊ�ݺ�����,�����Ѵ�ӡ���
        mblnPrint = True
        '����Ƿ��ӡƱ��
        If mintInvoicePrint = 0 Then
            mblnPrint = False
        Else
            If mintInvoicePrint = 2 Then
                If MsgBox("�Ƿ��ӡƱ�ݣ�" & vbCrLf & "Ҫȡ������ʾ,���ڱ��ز���������Ʊ�ݴ�ӡ���Ʋ���!", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                    mblnPrint = False
                End If
            End If
        End If
        
        '��������(ֻ�й�����)�Ƿ��ӡ,���۲�����������
        If mblnPrint Then
            If CalcBillToTal = Calc������ Then
                If MsgBox("��ǰ����ʵ��û����ȡ����,Ҫ��ӡ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    mblnPrint = False
                End If
            End If
        End If
           
        If Not mblnPrint Then
            j = 0
            For i = 1 To mobjBill.Pages(1).Details.Count
                If mobjBill.Pages(1).Details(i).������ Then
                    If j = 0 Then MsgBox "��Ϊ����ӡƱ��,ϵͳ���Զ�ɾ�������ѣ�", vbInformation, gstrSysName
                    j = j + 1
                    Call DeleteDetail(i)
                    Call ShowDetails
                    Call ShowMoney
                    Bill.TxtVisible = False:  Bill.CmdVisible = False: Bill.CboVisible = False
                    Exit For
                End If
            Next
        Else
            If gblnStrictCtrl Then
                If Trim(txtInvoice.Text) = "" Then
                    MsgBox "��������һ����Ч��Ʊ�ݺ��룡", vbInformation, gstrSysName
                    txtInvoice.SetFocus: Exit Sub
                End If
                If zlGetInvoiceGroupUseID(mlng����ID, 1, txtInvoice.Text) = False Then
                    Exit Sub
                End If
                 
                '�����������,Ʊ���Ƿ�����
                If CheckBillRepeat(mlng����ID, 1, txtInvoice.Text) Then
                    MsgBox "Ʊ�ݺ�""" & txtInvoice.Text & """�Ѿ���ʹ�ã����������롣", vbInformation, gstrSysName
                    txtInvoice.SetFocus: Exit Sub
                End If
            Else
                If Len(txtInvoice.Text) <> gbytFactLength And txtInvoice.Text <> "" Then
                    MsgBox "Ʊ�ݺ��볤��Ӧ��Ϊ " & gbytFactLength & " λ��", vbInformation, gstrSysName
                    txtInvoice.SetFocus: Exit Sub
                End If
            End If
        End If
        
                
        If gblnLED And Val(txtӦ��.Text) <> 0 And Not mbln���ϼ� And Not gbln�ֹ����� Then
            zl9LedVoice.Speak "#21 " & txtӦ��.Text
        End If
        
        If IsDate(txtDate.Text) Then mobjBill.����ʱ�� = CDate(txtDate.Text)
        mobjBill.�Ǽ�ʱ�� = zlDatabase.Currentdate
        If zlGetSaveDataItems_Plugin(mobjBill, str����Nos, rsItems) = False Then Exit Sub
        If zlChargeSaveValied_Plugin(glngModul, 1, True, False, str����Nos, rsItems) = False Then Exit Sub
        
        '���浥��
        If Not SaveBill Then Exit Sub
        
        Call zlChargeSaveAfter_Plugin(glngModul, mobjBill.����ID, mobjBill.��ҳID, True, 1, mobjBill.NO)
        
        '�����Ĵ���
        If mblnPrint Then '��ӡ�����վ�
            Call frmPrint.ReportPrint(1, "'" & mobjBill.NO & "'", "", "", mlng����ID, mlngShareUseID, txtInvoice.Text, _
                mobjBill.�Ǽ�ʱ��, , , , mintInvoiceFormat, , , mstrUseType, , , , mstr��ͨ�۸�ȼ�)
        End If
        
        '�����嵥�Ĵ�ӡ
        If InStr(mstrPrivs, "��ӡ�嵥") > 0 Then
            If gint�շ��嵥 = 1 Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_3", Me, "NO='" & mobjBill.NO & "'", "ҩƷ��λ=0", 2)
            ElseIf gint�շ��嵥 = 2 Then
                If MsgBox("Ҫ��ӡ�շ��嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_3", Me, "NO='" & mobjBill.NO & "'", "ҩƷ��λ=0", 2)
                End If
            End If
        End If
        
        If mbytInState = 0 And gbln�ۼ� Then
            txt�ۼ�.Text = Format(GetChargeTotal, "0.00")
        End If
        
        If mstrInNO = "" Then
            mstrInNO = ""
            sta.Panels(2) = "��һ�ŵ���:" & mobjBill.NO
            '�Ƿ���������շѣ�
            'ʹ��Ԥ�������,�����շѽ���(�������ý��ɿ��������)
            '���ѽɿ�,��ǿ����Ϊ�����շѽ���
            
            '���˺�:22343
            If CCur(txt�ɿ�.Text) <> 0 Or (Val(txtԤ�����.Text) <> 0 And gTy_Module_Para.byt�ɿ���� <> 1) Then
                mcurBillʵ�� = 0:  mcurBillӦ�� = 0: mcurBillӦ�� = 0
                mstrPrePati = "": mintBillNO = 0: mintMoneyRow = 0
                Call ClearPatientInfo
                txt�ϼ�.Text = gstrDec: txtӦ��.Text = gstrDec
            Else
                mstrPrePati = mobjBill.���� '��¼��ǰ����
                
                '���˵��ݽ���ۼ�
                mcurBillӦ�� = mcurBillӦ�� + CalcBillToTal(True)
                mcurBillʵ�� = mcurBillʵ�� + CalcBillToTal
                mcurBillӦ�� = mcurBillӦ�� + mobjBill.Pages(1).Ӧ�ɽ��
                
                mintBillNO = mintBillNO + 1
                For i = 1 To mshMoney.Rows - 1
                    If mshMoney.TextMatrix(i, 0) = "" Then Exit For
                Next
                mintMoneyRow = i - 1
            End If
            
            Call ClearRows: Call Bill.ClearBill
            Call NewBill(, False) '�����÷ѱ�
            txtPatient.SetFocus
        Else '��ȡ�޸�
            Unload Me
        End If
    ElseIf chkCancel.Value = 1 Then '�˵���״̬
        If mstrInNO = "" Then
            MsgBox "û����ȷ��ȡ��������,����ִ�иò�����", vbInformation, gstrSysName
            cboNO.SetFocus: Exit Sub
        End If
        
        If gblnBillPrint Then
            If gobjBillPrint.zlEraseBill("'" & mstrInNO & "'", 0) = False Then Exit Sub
        End If
        
        strSQL = "Zl_������շ�_Delete('" & mstrInNO & "','" & UserInfo.��� & "','" & UserInfo.���� & "')"
        
        On Error GoTo errH
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        On Error GoTo 0
        
        If Not gobjTax Is Nothing And gblnTax Then
            gstrTax = gobjTax.zlTaxOutErase(gcnOracle, "'" & mstrInNO & "'")
            If gstrTax <> "" Then MsgBox gstrTax, vbExclamation, gstrSysName
        End If
        
        If mbytInState = 0 And gbln�ۼ� Then txt�ۼ�.Text = Format(GetChargeTotal, "0.00")
        
        mstrInNO = "": cboNO.Text = "": txtInvoice.Text = ""
        Call ClearRows: Call Bill.ClearBill
        Call ClearMoney
        chkCancel.Value = Unchecked
        Call ClearPatientInfo
        txt�ϼ�.Text = gstrDec: txtӦ��.Text = gstrDec
        Call NewBill
        Call SetDisible(True)
        txtPatient.SetFocus
    ElseIf Not Bill.Active Then '��ȡ���۵�����״̬
        If mstrInNO = "" Then
            MsgBox "û����ȷ��ȡ���ݣ�", vbInformation, gstrSysName
            cboNO.SetFocus: Exit Sub
        End If
        If txtPatient.Text = "" Then
            MsgBox "û�з��ֲ�����Ϣ,�����벡����Ϣ��", vbInformation, gstrSysName
            txtPatient.SetFocus: Exit Sub
        End If
        If Not IsDate(txtDate.Text) Then
            MsgBox "��������ȷ�ķ���ʱ�䣡", vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
    
        If cbo��������.ListIndex = -1 Or cbo��������.Text = "" Then
            MsgBox "��ȷ���������ң�", vbInformation, gstrSysName
            cbo��������.SetFocus: Exit Sub
        End If
        If cbo���㷽ʽ.ListIndex = -1 Then
            MsgBox "��ȷ���շѵĽ��㷽ʽ��", vbInformation, gstrSysName
            cbo���㷽ʽ.SetFocus: Exit Sub
        End If
        If Val(txt�ɿ�.Text) <> 0 And txt�ɿ�.Enabled Then
            If Val(txt�ɿ�.Text) < Val(txtӦ��.Text) Then
                MsgBox "���˽ɿ���㣬�벹��Ӧ�ɽ�", vbInformation, gstrSysName
                Call zlControl.TxtSelAll(txt�ɿ�): txt�ɿ�.SetFocus: Exit Sub
            End If
        End If
        '���˺�:22343,�ɿ������
        Select Case gTy_Module_Para.byt�ɿ����
        Case 1  '1-��������ɿ��Ž��������ۼ�
        Case 2  '2-�շ�ʱ����Ҫ����ɿ���
            If Val(txtӦ��.Text) > 0 And Val(txt�ɿ�.Text) = 0 Then
                MsgBox "ע��:" & vbCrLf & _
                "    �ò���δ����ɿ���,���ܽ����շ�!", vbInformation + vbDefaultButton1, gstrSysName
                If txt�ɿ�.Enabled And txt�ɿ�.Visible Then txt�ɿ�.SetFocus
                Exit Sub
            End If
        Case Else   ',0-�������нɿ�������ۼƿ���
        End Select
        
                    
        'Ʊ�ݺ�����,�����Ѵ�ӡ���
        mblnPrint = True
        '����Ƿ��ӡƱ��
        If mintInvoicePrint = 0 Then
            mblnPrint = False
        Else
            If mintInvoicePrint = 2 Then
                If MsgBox("�Ƿ��ӡƱ�ݣ�" & vbCrLf & "Ҫȡ������ʾ,���ڱ��ز���������Ʊ�ݴ�ӡ���Ʋ���!", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                    mblnPrint = False
                End If
            End If
        End If
        
        If mblnPrint Then
            If gblnStrictCtrl Then
                If Trim(txtInvoice.Text) = "" Then
                    MsgBox "��������һ����Ч��Ʊ�ݺ��룡", vbInformation, gstrSysName
                    txtInvoice.SetFocus: Exit Sub
                End If
                If zlGetInvoiceGroupUseID(mlng����ID, 1, txtInvoice.Text) = False Then
                    Exit Sub
                End If
                '�����������,Ʊ���Ƿ�����
                If CheckBillRepeat(mlng����ID, 1, txtInvoice.Text) Then
                    MsgBox "Ʊ�ݺ�""" & txtInvoice.Text & """�Ѿ���ʹ�ã����������롣", vbInformation, gstrSysName
                    txtInvoice.SetFocus: Exit Sub
                End If
            Else
                If Len(txtInvoice.Text) <> gbytFactLength And txtInvoice.Text <> "" Then
                    MsgBox "Ʊ�ݺ��볤��Ӧ��Ϊ " & gbytFactLength & " λ��", vbInformation, gstrSysName
                    txtInvoice.SetFocus: Exit Sub
                End If
            End If
        End If
        
        If cboҽ�Ƹ���.ListIndex <> -1 Then
            strҽ�Ƹ��� = Mid(cboҽ�Ƹ���.Text, 1, InStr(1, cboҽ�Ƹ���, "-") - 1)
        End If
        
        lng����ID = zlDatabase.GetNextId("���˽��ʼ�¼")
        '�ɿ����
        strSQL = "zl_�����շѼ�¼_INSERT('" & mstrInNO & "'," & Val(txtPatient.Tag) & "," & gint������Դ & ",'" & _
            strҽ�Ƹ��� & "','" & txtPatient.Text & "'," & _
            "'" & zlStr.NeedName(cboSex.Text) & "','" & mobjBill.���� & "'," & _
             ZVal(mobjBill.����ID, , cbo��������.ItemData(cbo��������.ListIndex)) & "," & _
            cbo��������.ItemData(cbo��������.ListIndex) & ",'" & zlStr.NeedName(cbo������.Text) & "'," & _
            "'" & zlStr.NeedName(cbo���㷽ʽ.Text) & "|" & mobjBill.Pages(1).Ӧ�ɽ�� & "| ',"
        'Ԥ������
        If Val(txtԤ�����.Text) <> 0 Then
            strSQL = strSQL & mobjBill.Pages(1).��Ԥ���� & ","
        Else
            strSQL = strSQL & "NULL,"
        End If
        strSQL = strSQL & "NULL," & lng����ID & ",To_Date('" & txtDate.Text & "','YYYY-MM-DD HH24:MI:SS')," & _
            "'" & UserInfo.��� & "','" & UserInfo.���� & "','Z',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,1)"
        
        On Error GoTo errH
        gcnOracle.BeginTrans
        '��ȡ���۵�����
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        
        '������
        If mobjBill.Pages(1).����� <> 0 Then
            '  Zl_���շ����_Insert(
            '  No_In         ������ü�¼.No%Type,
            '  ����id_In     ������ü�¼.����id%Type,
            '  ����id_In     ������ü�¼.����id%Type,
            '  �����_In   ������ü�¼.ʵ�ս��%Type,
            '  �Ǽ�ʱ��_In   ������ü�¼.�Ǽ�ʱ��%Type,
            '  ����Ա���_In ������ü�¼.����Ա���%Type,
            '  ����Ա����_In ������ü�¼.����Ա����%Type
            strSQL = "Zl_���շ����_Insert('" & mstrInNO & "'," & Val(txtPatient.Tag) & "," & lng����ID & "," & _
                mobjBill.Pages(1).����� & ",To_Date('" & txtDate.Text & "','YYYY-MM-DD HH24:MI:SS')," & _
                "'" & UserInfo.��� & "','" & UserInfo.���� & "')"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        End If
        gcnOracle.CommitTrans
        On Error GoTo 0
        
        If mblnPrint Then '��ӡ�����վ�
            Call frmPrint.ReportPrint(1, "'" & mstrInNO & "'", "", "", mlng����ID, mlngShareUseID, txtInvoice.Text, _
                zlDatabase.Currentdate, , , , mintInvoiceFormat, , , mstrUseType, , , , mstr��ͨ�۸�ȼ�)
        End If
        '�����嵥�Ĵ�ӡ
        If InStr(mstrPrivs, "��ӡ�嵥") > 0 Then
            If gint�շ��嵥 = 1 Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_3", Me, "NO='" & mstrInNO & "'", "ҩƷ��λ=0", 2)
            ElseIf gint�շ��嵥 = 2 Then
                If MsgBox("Ҫ��ӡ�շ��嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_3", Me, "NO='" & mstrInNO & "'", "ҩƷ��λ=0", 2)
                End If
            End If
        End If
        
        '�Ƿ���������շѣ�
        'ʹ��Ԥ�������,�����շѽ���(�������ý��ɿ��������)
        '���ѽɿ�,��ǿ����Ϊ�����շѽ���
        
        '���˺�:22343
        If CCur(txt�ɿ�.Text) <> 0 Or (Val(txtԤ�����.Text) <> 0 And gTy_Module_Para.byt�ɿ���� <> 1) Then
            mcurBillʵ�� = 0: mcurBillӦ�� = 0: mcurBillӦ�� = 0
            mstrPrePati = "": mintBillNO = 0: mintMoneyRow = 0
            Call ClearPatientInfo
            txt�ϼ�.Text = gstrDec: txtӦ��.Text = gstrDec
        Else
            mstrPrePati = txtPatient.Text
            mcurBillӦ�� = mcurBillӦ�� + CalcBillToTal(True)
            mcurBillʵ�� = mcurBillʵ�� + CalcBillToTal
            mcurBillӦ�� = mcurBillӦ�� + mobjBill.Pages(1).Ӧ�ɽ��
            
            mintBillNO = mintBillNO + 1
            For i = 1 To mshMoney.Rows - 1
                If mshMoney.TextMatrix(i, 0) = "" Then Exit For
            Next
            mintMoneyRow = i - 1
        End If
        
        mstrInNO = ""
        
        Call SetDisible(True)
        Call ClearRows: Call Bill.ClearBill
        Call NewBill
        
        mstrPreUnit = ""
        Call cbo��������_Click 'ɾ����Ч������
        
        txtPatient.SetFocus
    End If
    gblnOK = True
    Exit Sub
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    If mbytInState = 1 Then
        cmdCancel.SetFocus
    ElseIf mbytInState = 2 Then
        txtDate.SetFocus
    ElseIf mstrInNO <> "" Then
        Bill.SetFocus
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("',|~:��;��?��" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim strPre As String, strTmp As String
    Dim lngPre As Long, i As Long
    mblnStartFactUseType = zlStartFactUseType(1)
    Call RestoreWinState(Me, App.ProductName)
    Me.Width = 10000: Me.Height = 7770
    Call initCardSquareData
    If IsCheck����() = False Then Unload Me: Exit Sub
    
    'LED��ʼ��
    If mbytInState = 0 And gblnLED Then
        zl9LedVoice.Reset com
        zl9LedVoice.Init UserInfo.��� & " �շ�ԱΪ������", mlngModul, gcnOracle
    End If
    
    gblnOK = False
    mstrPrePati = "": mcurBillʵ�� = 0: mcurBillӦ�� = 0
    mstr���ʽ = ""
    mstrPreUnit = ""
    mblnDo = True
    mbln������۸� = False
    txtӦ��.Text = gstrDec: txt�ϼ�.Text = gstrDec
    
    Set mobjBill = New ExpenseBill
    Set mrsInfo = New ADODB.Recordset
    
    '�鿴����ʱ�������ʼ����
    If mbytInState = 0 Or mbytInState = 2 Then
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
    
    Call InitFace
    
    If mbytInState = 1 Or mbytInState = 2 Then '���������
        If Not ReadBill(mstrInNO) Then Unload Me: Exit Sub
        cboNO.Text = mstrInNO
    Else '��������
        If mbytInState = 0 And gbln�ۼ� Then txt�ۼ�.Text = Format(GetChargeTotal, "0.00")
        
        If Not NewBill(IIf(mblnStartFactUseType, False, True)) Then Unload Me: Exit Sub
        
        '��ȡ�õ��ݵ�����
        If mstrInNO <> "" Then '�޸�ԭ����
            Set mobjBill = ImportBill(mstrInNO, True, 0, , , True, mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ�۸�ȼ�)
            If mobjBill.NO = "" Then
                MsgBox "��ȡ����ʧ�ܡ�", vbInformation, gstrSysName
                Unload Me: Exit Sub
            Else
                txtPatient.BackColor = &HE0E0E0
                cboNO.Text = mobjBill.NO
                
                mbln������۸� = True               '�ڷѱ�_click�¼��в�����۸�
                Call Set�����˿�������(mobjBill.Pages(1).������, mobjBill.Pages(1).��������ID)
                                    
                If mobjBill.����ID <> 0 Then
                    txtPatient.Text = "-" & mobjBill.����ID
                    Call txtPatient_KeyPress(13)
                Else
                    txtPatient.Text = mobjBill.����
                    cboSex.ListIndex = cbo.FindIndex(cboSex, mobjBill.�Ա�, True)
                    Call LoadOldData(mobjBill.����, txt����, cbo���䵥λ)
                    
                    '������ŵ��ݷѱ���ͬ,��λ
                    strTmp = GetBill�ѱ�(mobjBill)
                    If strTmp <> "" Then
                        cbo�ѱ�.ListIndex = cbo.FindIndex(cbo�ѱ�, strTmp, True)
                        If cbo�ѱ�.ListIndex = -1 Then
                           cbo�ѱ�.AddItem strTmp
                           cbo�ѱ�.ListIndex = cbo�ѱ�.NewIndex
                        End If
                    End If
                    If cbo�ѱ�.ListIndex = -1 And cbo�ѱ�.ListCount > 0 Then cbo�ѱ�.ListIndex = 0
                    
                    If gint������Դ <> 2 Then cboҽ�Ƹ���.ListIndex = GetCboIndexByCode(cboҽ�Ƹ���, "" & mobjBill.����)
                End If
                mbln������۸� = False
                
                Bill.ClearBill
                Bill.Rows = mobjBill.Pages(1).Details.Count + 1
                '����б༭����������ɫ
                Bill.SetColColor 0, &HE7CFBA
                Bill.SetColColor 1, &HE7CFBA
                Bill.SetColColor 3, &HE7CFBA
                txtDate.Text = Format(mobjBill.����ʱ��, "yyyy-MM-dd HH:mm:ss")
                chk�Ӱ�.Value = mobjBill.�Ӱ��־
                chkCancel.Enabled = False
                
                mobjBill.����Ա��� = UserInfo.���
                mobjBill.����Ա���� = UserInfo.����
                
                'ȱʡΪԭ���ݵĽ��㷽ʽ
                strTmp = GetBalanceName(mstrInNO)
                If strTmp <> "" Then
                    i = cbo.FindIndex(cbo���㷽ʽ, strTmp, True)
                    If i <> -1 Then cbo���㷽ʽ.ListIndex = i
                End If
                
                Call ShowDetails
                Call ShowMoney
                
                '�޸ĵ���ʱ�������޸Ĳ�����Ϣ
                txtPatient.Locked = True
                Call ReInitPatiInvoice
            End If
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
    
    mbytInState = Empty
    mstrInNO = Empty
    mintBillNO = 0: mintMoneyRow = 0
    mlng����ID = 0
    mstrCardNO = ""
    mstrDelete = ""
    zlCommFun.OpenIme False
    mblnNOMoved = False
    mintInvoicePrint = 0
    Set mrs�������� = Nothing
    Set mrs������ = Nothing
    Set mrs�ѱ� = Nothing
    Call initCardSquareData
    'LED��ʼ��
    If mbytInState = 0 And gblnLED Then
        zl9LedVoice.DisplayPatient ""
        zl9LedVoice.Reset com
    End If
End Sub

Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng�����ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
    If txtPatient.Locked Then Exit Sub
    
   If objCard.���� Like "IC��*" And objCard.ϵͳ Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        If mobjICCard Is Nothing Then Exit Sub
        txtPatient.Text = mobjICCard.Read_Card()
        If txtPatient.Text <> "" Then
            Call txtPatient_KeyPress(vbKeyReturn)
        End If
        Exit Sub
    End If
    
    lng�����ID = objCard.�ӿ����
    If lng�����ID <= 0 Then Exit Sub
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
    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModul, lng�����ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtPatient.Text = strOutCardNO
    If txtPatient.Text <> "" Then
        Call txtPatient_KeyPress(vbKeyReturn)
    End If
End Sub

Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txtPatient.IMEMode = 0
    Set gobjSquare.objCurCard = objCard
    If txtPatient.Text <> "" Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    zlControl.TxtSelAll txtPatient
End Sub

 
Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    Dim lngPreIDKind As Long
    If txtPatient.Locked Then Exit Sub
    mblnNotClick = True
    lngPreIDKind = IDKind.IDKind
    IDKind.IDKind = IDKind.GetKindIndex(objCard.����)
    txtPatient.Text = objPatiInfor.����
    Call txtPatient_KeyPress(vbKeyReturn)
    IDKind.IDKind = lngPreIDKind
    mblnNotClick = False
End Sub

Private Sub sta_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Not gbln�����л� Then Exit Sub
    If Panel.Bevel = sbrRaised And (Panel.Key = "PY" Or Panel.Key = "WB") Then
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

Private Sub txtDate_GotFocus()
    zlControl.TxtSelAll txtDate
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And IsDate(txtDate.Text) Then
        mobjBill.����ʱ�� = CDate(txtDate.Text)
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtDate_LostFocus()
    txtDate.SelLength = 0
    If IsDate(txtDate.Text) Then mobjBill.����ʱ�� = CDate(txtDate.Text)
End Sub

Private Sub cboNO_GotFocus()
    cboNO.SelStart = 0
    cboNO.SelLength = Len(cboNO.Text)
    If (mobjBill.Pages(1).Details.Count = 0 And mbytInState = 0) Or chkCancel.Value = Checked Then
        cboNO.Locked = False
    Else
        cboNO.Locked = True
    End If
End Sub

Private Sub cboNO_KeyPress(KeyAscii As Integer)
    Dim blnRead As Boolean, blnNull As Boolean
    Dim strOper As String, vDate As Date, i As Integer
    
    If KeyAscii > 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))

    '��һλ����������ĸ,����λ����
    If KeyAscii <> 13 Then
        Call SetNOInputLimit(cboNO, KeyAscii)
    End If
    
    If KeyAscii = 13 And cboNO.Text <> "" And Not cboNO.Locked Then
        txt�Ҳ�.Text = "0.00"
        txt�ɿ�.Text = "0.00"
        txtӦ��.Text = "0.00"
    
        cboNO.Text = GetFullNO(cboNO.Text, 13)
        If chkCancel.Value = 1 Then
            '�Ƿ���ת������ݱ���
            If zlDatabase.NOMoved("������ü�¼", cboNO.Text, , "1") Then
                If Not ReturnMovedExes(cboNO.Text, 1, Me.Caption) Then Exit Sub
                mblnNOMoved = False
            End If
        
            '�����˷�Ȩ���ж�
            If Not ReadBillInfo(1, cboNO.Text, 1, strOper, vDate) Then
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If
            If InStr(mstrPrivs, "���в���Ա") <= 0 Then
                If UserInfo.���� <> strOper Then
                    MsgBox "��û��""���в���Ա""Ȩ��,���ܶ�" & strOper & "�ĵ��ݽ����˷�!", vbInformation, gstrSysName
                    cboNO.Text = "": cboNO.SetFocus: Exit Sub
                End If
            End If
            If Not BillOperCheck(2, strOper, vDate, "�˷�", cboNO.Text, , 1) Then
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If
            
            '�˷�ʱ,���ݱ���Ϊ���շѵĵ���
            If Not isSimple(cboNO.Text) Then
                MsgBox "�õ��ݲ����ڻ��Ǽ��շѵ��ݣ�", vbInformation, gstrSysName
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If
                        
            '�Ƿ���ִ��
            i = BillCanDelete(cboNO.Text, 1)
            If i <> 0 Then
                Select Case i
                    Case 1 '�õ��ݲ�����
                        MsgBox "ָ���ĵ��ݲ����ڣ�", vbInformation, gstrSysName
                    Case 2 '�Ѿ�ȫ����ȫִ��
                        MsgBox "�õ����е���Ŀ�Ѿ�ȫ����ȫִ�У�", vbInformation, gstrSysName
                    Case 3 'δ��ȫִ�в���ʣ������Ϊ0
                        MsgBox "�õ�����δ��ȫִ�в�����Ŀʣ������Ϊ��,û�п����˷ѵķ��ã�", vbInformation, gstrSysName
                End Select
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If
        End If
                
        txtPatient.PasswordChar = ""
        '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
        txtPatient.IMEMode = 0
        If chkCancel.Value = 1 Then '��ȡ�˷ѵ�
            blnRead = ReadBill(cboNO.Text, False, True)
        ElseIf mobjBill.Pages(1).Details.Count = 0 Then '��ȡ���۵�
            blnRead = ReadBill(cboNO.Text, True, False, blnNull)
        End If
        
        lbl��̬�ѱ�.Visible = blnRead
        If blnRead Then
            '��ʾ��̬�ѱ�:��ʾ�˷ѵ�����ʾҪ��ȡ�Ļ��۵�ʱ
            cbo�ѱ�.Locked = True
            cbo�ѱ�.Visible = False
            lbl��̬�ѱ�.BorderStyle = 1
            lbl��̬�ѱ�.Left = cbo�ѱ�.Left
            lbl��̬�ѱ�.Width = cboҽ�Ƹ���.Left - cboSex.Left
        
            mstrInNO = cboNO.Text 'ȷ��ʱ��mstrInNOΪ׼
            If chkCancel.Value = 0 Then '���۵�
                chk�Ӱ�.Enabled = False
                Bill.Active = False
                
                If gint������Դ = 1 And InStr(mstrPrivs, "�����ҽ������") = 0 Then
                     Call ClearPatientInfo
                End If
                                
                If txtPatient.Text = "" Or blnNull Then
                    txtPatient.SetFocus
                Else
                    If txt�ɿ�.Visible Then
                        txt�ɿ�.SetFocus
                    Else
                        cmdOK.SetFocus
                    End If
                End If
            Else '��
                Call SetDisible 'cboNO�ڻ�ȡ�����unLock
                '�����˷�ֻ֧���˷�ָ�����㷽ʽ
                cbo���㷽ʽ.Locked = False
                cmdOK.SetFocus
            End If
        Else
            Call ClearPatientInfo: Call ClearMoney: Call ClearRows
            mstrInNO = "": cboNO.Text = "": cboNO.SetFocus
        End If
    End If
End Sub

Private Sub txtPatient_Change()
    If txtPatient.Enabled = False Or txtPatient.Locked Then Exit Sub
    IDKind.SetAutoReadCard (txtPatient.Text = "")
End Sub

Private Sub txtPatient_KeyDown(KeyCode As Integer, Shift As Integer)

    If txtPatient.Locked Or txtPatient.Enabled = False Or txtPatient.Text <> "" Then Exit Sub
    If IDKind.GetCurCard Is Nothing Then Exit Sub
    If IDKind.ActiveFastKey = True Then Exit Sub
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
    
    If mbytInState = 0 Then mobjBill.���� = Trim(txt����.Text) & IIf(IsNumeric(txt����.Text), cbo���䵥λ.Text, "")
End Sub

Private Sub txtӦ��_Change()
    If Val(txt�ɿ�.Text) = 0 Then txt�ɿ�.Text = "0.00": txt�Ҳ�.Text = "0.00": Exit Sub
    txt�Ҳ�.Text = Format(Val(txt�ɿ�.Text) - Val(txtӦ��.Text), "0.00")
End Sub

Private Sub txtԤ�����_GotFocus()
    Call txt�ɿ�_GotFocus '�շ��Զ�����������
    zlControl.TxtSelAll txtԤ�����
End Sub

Private Sub txtԤ�����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
    End If
    If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtԤ�����_Validate(Cancel As Boolean)
    Dim curTotal As Currency
    
    curTotal = CalcBillToTal
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
            txtԤ�����.Text = Format(IIf(curTotal > Val(sta.Panels(3).Tag), Val(sta.Panels(3).Tag), curTotal), "0.00")
        End If
        zlControl.TxtSelAll txtԤ�����: Cancel = True: Exit Sub
    ElseIf Val(txtԤ�����.Text) > 0 And curTotal < 0 Then
        MsgBox "����Ӧ�����Ϊ��ʱ����ʹ��Ԥ���", vbInformation, gstrSysName
        txtԤ�����.Text = "0.00"
        zlControl.TxtSelAll txtԤ�����: Cancel = True: Exit Sub
    ElseIf Val(txtԤ�����.Text) > Val(sta.Panels(3).Tag) Then
        MsgBox "Ԥ��������ܳ������˵�Ԥ�����:" & Format(Val(sta.Panels(3).Tag), "0.00") & " ��", vbInformation, gstrSysName
        If curTotal < 0 Then
            txtԤ�����.Text = "0.00"
        Else
            txtԤ�����.Text = Format(IIf(curTotal > Val(sta.Panels(3).Tag), Val(sta.Panels(3).Tag), curTotal), "0.00")
        End If
        zlControl.TxtSelAll txtԤ�����: Cancel = True: Exit Sub
    ElseIf Val(txtԤ�����.Text) > Format(curTotal, "0.00") And Val(txtԤ�����.Text) <> 0 Then
        MsgBox "Ԥ��������ܴ���Ӧ�����:" & Format(curTotal, "0.00") & " ��", vbInformation, gstrSysName
        If curTotal < 0 Then
            txtԤ�����.Text = "0.00"
        Else
            txtԤ�����.Text = Format(IIf(curTotal > Val(sta.Panels(3).Tag), Val(sta.Panels(3).Tag), curTotal), "0.00")
        End If
        zlControl.TxtSelAll txtԤ�����: Cancel = True: Exit Sub
    Else
        txtԤ�����.Text = Format(txtԤ�����.Text, "0.00")
        
        '���¼���Ӧ�ɣ����(�ֱ�)��
        If Bill.Active Then
            Call ShowMoney
        Else
            Call ShowPrice
        End If
    End If
End Sub

Private Sub txtInvoice_GotFocus()
    zlControl.TxtSelAll txtInvoice
End Sub

Private Sub txtInvoice_LostFocus()
'    If Not txtInvoice.Locked And txtInvoice.Text <> "" Then
'        txtInvoice.Text = Format(Left(txtInvoice.Text, gbytFactLength), String(gbytFactLength, "0"))
'    End If
End Sub

Private Sub txt����_Gotfocus()
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
    End If
End Sub

Private Sub txtPatient_GotFocus()
    zlControl.TxtSelAll txtPatient
    zlCommFun.OpenIme True
    If txtPatient.Enabled = False Or txtPatient.Locked Then Exit Sub
    IDKind.SetAutoReadCard (txtPatient.Text = "")
End Sub

Private Sub bill_AfterAddRow(Row As Long)
    With Bill
        '������ʱ,�������ÿ����Ѿ������ĵĿɱ������е���ֵ
        .ColData(1) = 5 'Ӧ��ȱʡ����,����Ŀ���ʱ,��Ϊ����(4)
        '����б༭����������ɫ
        .SetColColor 0, &HE7CFBA
        .SetColColor 1, &HE7CFBA
        .SetColColor 3, &HE7CFBA
    End With
End Sub

Private Sub cboSex_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cboSex.ListIndex <> -1 Then mobjBill.�Ա� = Mid(cboSex.Text, InStr(cboSex.Text, "-") + 1)
        Call zlCommFun.PressKey(vbKeyTab)
    End If
    If cboSex.Locked Then Exit Sub
    If SendMessage(cboSex.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
End Sub

Private Sub cbo�ѱ�_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
   
    If cbo�ѱ�.Locked Then Exit Sub
    
    If KeyAscii >= 32 Then
        lngIdx = zlControl.CboMatchIndex(cbo�ѱ�.hWnd, KeyAscii)
        If lngIdx = -1 And cbo�ѱ�.ListCount > 0 Then lngIdx = 0
        cbo�ѱ�.ListIndex = lngIdx
        
    ElseIf KeyAscii = 13 And cbo�ѱ�.ListIndex <> -1 Then
        mobjBill.�ѱ� = Mid(cbo�ѱ�.Text, InStr(cbo�ѱ�.Text, "-") + 1)
        If mbytInState = 0 And mstrInNO <> "" And mobjBill.Pages(1).Details.Count > 0 Then
            '���¼���۸�
            Call CalcMoneys
            Call ShowDetails
            Call ShowMoney
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cbo��������_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    
    If KeyAscii = 13 Then
        mblnCboClick = False    '��������������б�ѡ��һ�������,��Ҫ�ƿ�,��ʱֻ����click,��������벢�һس�,������click,������Ҫ�ڴ˸�ֵ,�Ա�validate�¼���ǿ�е���click�¼�
        Call zlCommFun.PressKey(vbKeyTab)
        
    ElseIf KeyAscii >= 32 And Not cbo��������.Locked Then
        lngIdx = zlControl.CboMatchIndex(cbo��������.hWnd, KeyAscii)
        If lngIdx = -1 And cbo��������.ListCount > 0 Then lngIdx = 0
        cbo��������.ListIndex = lngIdx
    End If
End Sub

Private Sub cbo������_KeyPress(KeyAscii As Integer)
    Dim i As Long, intIdx As Integer
    Dim strText As String
    
    If KeyAscii = 13 Then
        If cbo������.Locked Then Exit Sub
        
        strText = cbo������.Text
        If cbo������.ListIndex <> -1 Then
            '�����б�ʱ,�����ı�������������
            If strText <> cbo������.List(cbo������.ListIndex) Then Call zlControl.CboSetIndex(cbo������.hWnd, -1)
        End If
        If strText = "" Then
            cbo������.ListIndex = -1
        ElseIf cbo������.ListIndex = -1 Then
            intIdx = -1
            For i = 0 To cbo������.ListCount - 1
                If UCase(cbo������.List(i)) Like UCase(strText) & "*" Then
                    If intIdx = -1 Then cbo������.ListIndex = i
                    intIdx = i
                End If
            Next
        ElseIf Not mblnDrop Then
            '�س���꾭��
            Call cbo������_Click
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        If cbo������.ListIndex = -1 Then
            cbo������.Text = ""
            mobjBill.Pages(1).������ = ""
            If gbyt����ҽ�� = 0 Or gbln�����俪���� Then Exit Sub
        Else
            mobjBill.Pages(1).������ = zlStr.NeedName(cbo������.Text)
            If intIdx <> -1 And mblnDrop Then
                '�����س�-ǿ�м���Click
                Call cbo������_Click
            ElseIf intIdx <> cbo������.ListIndex And intIdx <> -1 Then
                '������ѡ��-�Զ�����Click
                cbo������.SetFocus
                Call zlCommFun.PressKey(vbKeyF4)
                Exit Sub
            ElseIf intIdx <> -1 Then
                'һ��������-ǿ�м���Click
                Call cbo������_Click
            End If
        End If
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1  '����
            ShowHelp App.ProductName, Me.hWnd, Me.Name
        Case vbKeyF2
            If ActiveControl Is txtPatient Then
                Call txtPatient_Validate(False)
                Me.Refresh
            End If
            If ActiveControl Is cbo������ Then Call cbo������_KeyPress(vbKeyReturn)
            If cmdOK.Enabled And cmdOK.Visible Then
                Call cmdOK.SetFocus
                Call cmdOK_Click
            End If
        Case vbKeyF6 '��λ�����������
            txtPatient.SetFocus
            Call zlControl.TxtSelAll(txtPatient)
        Case vbKeyF7 '�л����뷨
            If Not gbln�����л� Then Exit Sub
            If sta.Panels("WB").Visible And sta.Panels("PY").Visible Then
                If sta.Panels("WB").Bevel = sbrRaised Then
                    Call sta_PanelClick(sta.Panels("WB"))
                Else
                    Call sta_PanelClick(sta.Panels("PY"))
                End If
            End If
        Case vbKeyF8 '��(�Զ������¼�)
            If chkCancel.Visible And chkCancel.Enabled Then chkCancel.Value = IIf(chkCancel.Value = Checked, Unchecked, Checked)
        Case vbKeyF9 '��λ�����ݺ������
            cboNO.SetFocus
            Call zlControl.TxtSelAll(cboNO)
        Case vbKeyF12
            If Shift = 2 Then
                'ǿ����LED����,(�ϼ�)
                If gblnLED And (Bill.Active Or (Not Bill.Active And chkCancel.Value = 0)) _
                    And txt�ɿ�.Enabled And txt�ɿ�.Visible And CCur(txt�ϼ�.Text) <> 0 Then
                    mblnHotKey = True: txt�ɿ�.SetFocus
                    If ActiveControl Is txt�ɿ� Then txt�ɿ�_GotFocus
                End If
            End If
        Case vbKeyEscape
            If Bill.TxtVisible Then
                Bill.Text = "": Bill.TxtVisible = False: Bill.SetFocus
            Else
                Call cmdCancel_Click
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
    
    mshMoney.ColWidth(1) = lngW * 0.45
    mshMoney.ColWidth(2) = lngW * 0.55
    
    mshMoney.ColAlignment(0) = 4
    mshMoney.ColAlignment(1) = 1
    mshMoney.ColAlignment(2) = 7
    
    mshMoney.TextMatrix(0, 0) = "���"
    mshMoney.TextMatrix(0, 1) = "��Ŀ"
    mshMoney.TextMatrix(0, 2) = "���"
    mshMoney.Row = 0
    mshMoney.Col = 0: mshMoney.CellAlignment = 4
    mshMoney.Col = 1: mshMoney.CellAlignment = 4
    mshMoney.Col = 2: mshMoney.CellAlignment = 4
    
    mshMoney.MergeCol(0) = True
End Sub

Private Function InitData() As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim i As Long, strSQL As String
    
    On Error GoTo errH
        
    '�Զ�ʶ��Ӱ�
    If OverTime(zlDatabase.Currentdate) Then chk�Ӱ�.Value = Checked
    
    '���䵥λ
    cbo���䵥λ.AddItem "��"
    cbo���䵥λ.AddItem "��"
    cbo���䵥λ.AddItem "��"
    cbo���䵥λ.ListIndex = 0
    
    '��ѡ�Ա�
    strSQL = "Select ����,����,����,Nvl(ȱʡ��־,0) as ȱʡ From �Ա� Order by ����"
    Set rsTmp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cboSex.AddItem rsTmp!���� & "-" & rsTmp!����
            If rsTmp!ȱʡ = 1 Then cboSex.ListIndex = cboSex.NewIndex
            rsTmp.MoveNext
        Next
    End If
    
    '��ѡҽ�Ƹ��ʽ
    strSQL = "Select ����,����,����,Nvl(ȱʡ��־,0) as ȱʡ From ҽ�Ƹ��ʽ Order by ����"
    Set rsTmp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
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
    
    '���㷽ʽ
    Set rsTmp = Get���㷽ʽ("�շ�")
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            'ֻ�����ҽ���Ľ��㷽ʽ��ѡ��
            If InStr(",1,2,", rsTmp!����) > 0 Then
                cbo���㷽ʽ.AddItem rsTmp!���� & "-" & rsTmp!����
                cbo���㷽ʽ.ItemData(cbo���㷽ʽ.NewIndex) = rsTmp!����
                
                If rsTmp!���� = gstr���㷽ʽ Then
                    cbo���㷽ʽ.ListIndex = cbo���㷽ʽ.NewIndex
                End If
                
                If rsTmp!ȱʡ = 1 And cbo���㷽ʽ.ListIndex = -1 Then
                    cbo���㷽ʽ.ListIndex = cbo���㷽ʽ.NewIndex
                End If
            End If
            rsTmp.MoveNext
        Next
    End If
    If cbo���㷽ʽ.ListCount = 0 Then
        MsgBox "�շѳ���û�п��õĽ��㷽ʽ�����ȵ����㷽ʽ���������á�", vbInformation, gstrSysName
        Exit Function
    End If
   
    
    '��ȱʡ�����˺Ϳ�������
    Call FillDept
    If cbo��������.ListCount = 0 Then
        MsgBox "û�п��õĿ�������,���õĿ����������������¹���:" & vbCrLf & _
               "    1.��������Ϊ����" & vbCrLf & _
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
    
    '��������
    txtDate.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    Set mrsInfo = New ADODB.Recordset
    
    InitData = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub InitFace()
'���ܣ����ݱ�Ҫ��ɵĹ������ý��沼��
    Dim arrHead() As String, i As Integer
    
    lblTitle.Caption = gstrUnitName & "�����շѵ�"
    
    '���õ��ݱ��ʽ
    With Bill
        .LocateCol = 0
        .PrimaryCol = 0
        .Font.Size = 11
        .CboFont.Size = 11
        .TxtEditFont.Size = 11
        arrHead = Split(STR_HEAD, ";")
        .COLS = UBound(arrHead) + 1
        For i = 0 To UBound(arrHead)
            .TextMatrix(0, i) = Split(arrHead(i), ",")(0)
            .ColWidth(i) = Split(arrHead(i), ",")(1)
            .ColAlignment(i) = Split(arrHead(i), ",")(2)
        Next
        If mbytInState = 0 Then
            .ColData(0) = 1 '��Ŀ����,��Ť��ѡ
            .ColData(1) = 5 'Ӧ�ս��ȱʡ����,����Ŀ���ʱ,��Ϊ����(4)
            .ColData(2) = 5 'ʵ�ս������
            .ColData(3) = 3 'Ĭ��ȡ�������һ���һ����
            .ColData(4) = 5
            
            .SetColColor 0, &HE7CFBA
            .SetColColor 1, &HE7CFBA
            .SetColColor 3, &HE7CFBA
            
            ReDim marrColData(.COLS - 1)
            For i = 0 To .COLS - 1
                marrColData(i) = .ColData(i)
            Next
        End If
    End With
    Call RestoreFlexState(Bill, App.ProductName & "\" & Me.Name)
    
    '��ȡ����ƥ�䷽ʽ
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
    
    If mbytInState = 1 Or mbytInState = 2 Then
        cbo�ѱ�.Visible = False
        lbl��̬�ѱ�.Left = cbo�ѱ�.Left
        lbl��̬�ѱ�.Visible = True
    Else
        lbl��̬�ѱ�.BorderStyle = 0
        lbl��̬�ѱ�.AutoSize = True
    End If
    
    Call SetMoneyList
    
    'Ȩ������
    If mbytInState = 0 Then
        If InStr(mstrPrivs, "�����˷�") = 0 Then
            chkCancel.Visible = False
            lblFact.Left = lblFact.Left + chkCancel.Width
            txtInvoice.Left = txtInvoice.Left + chkCancel.Width
            lbl���ݺ�.Left = lbl���ݺ�.Left + chkCancel.Width
            cboNO.Left = cboNO.Left + chkCancel.Width
        End If
        
        If InStr(mstrPrivs, "�ش�Ʊ��") = 0 Then
            lblRePrint.Visible = False
            txtRePrint.Visible = False
        End If
        txtInvoice.Locked = Not (InStr(1, mstrPrivs, "�޸�Ʊ�ݺ�") > 0) And gblnStrictCtrl
            
        If Not gbln�ۼ� Then
            lbl�ۼ�.Visible = False
            txt�ۼ�.Visible = False
            lblӦ��.Top = lblӦ��.Top + txt�ۼ�.Height / 3
            txtӦ��.Top = txtӦ��.Top + txt�ۼ�.Height / 3
            lbl�ϼ�.Top = lbl�ϼ�.Top + txt�ۼ�.Height / 1.5
            txt�ϼ�.Top = txt�ϼ�.Top + txt�ۼ�.Height / 1.5
        End If
    Else
        lblӦ��.Visible = False
        txtӦ��.Visible = False
        lbl�ɿ�.Visible = False
        lbl�Ҳ�.Visible = False
        txt�ɿ�.Visible = False
        txt�Ҳ�.Visible = False
        
        lbl�ۼ�.Visible = False
        txt�ۼ�.Visible = False
        
        lblӦ��.Top = lblӦ��.Top + txt�ۼ�.Height / 3
        txtӦ��.Top = txtӦ��.Top + txt�ۼ�.Height / 3
        lbl�ϼ�.Top = lbl�ϼ�.Top + txt�ۼ�.Height / 1.5
        txt�ϼ�.Top = txt�ϼ�.Top + txt�ۼ�.Height / 1.5
        
        txtԤ�����.Top = txtӦ��.Top
        lblDeposit.Top = txtԤ�����.Top + (txtԤ�����.Height - lblDeposit.Height) / 2
        
        fraTitle.Enabled = False
        
        lblRePrint.Visible = False
        txtRePrint.Visible = False
        
        chkCancel.Visible = False
        If mstrDelete <> "" Then
            lblFlag.Visible = True
        Else
            lblFact.Left = lblFact.Left + chkCancel.Width
            txtInvoice.Left = txtInvoice.Left + chkCancel.Width
            lbl���ݺ�.Left = lbl���ݺ�.Left + chkCancel.Width
            cboNO.Left = cboNO.Left + chkCancel.Width
        End If
        
        Call SetDisible
        
        If mbytInState = 2 Then
            txtDate.Enabled = True
            cbo������.Locked = False
        Else
            cmdOK.Visible = False
            cmdCancel.Caption = "�˳�(&X)"
            cmdCancel.Top = cmdCancel.Top - cmdCancel.Height / 2
        End If
    End If
    
    '�������
    If Not gbln�Ա� Then cboSex.TabStop = False
    If Not gbln���� Then txt����.TabStop = False: cbo���䵥λ.TabStop = False
    If Not gbln�ѱ� Then cbo�ѱ�.TabStop = False
    If Not gbln�Ӱ� Then chk�Ӱ�.TabStop = False
    If Not gbln�������� Then txtDate.TabStop = False
    If Not gbln������ Then cbo������.TabStop = False
    If Not gblnҽ�Ƹ��� Then cboҽ�Ƹ���.TabStop = False
       
    If gbyt����ҽ�� = 0 Then
        Call ExChangeLocate(cbo��������, cbo������)
        lbl����.Caption = "������"
        lbl������.Caption = "��������"
        cbo��������.TabStop = False
    End If
    
    '82801,Ƚ����,2015-2-26
    txt����.MaxLength = zlGetPatiInforMaxLen.intPatiAge
End Sub

Private Sub SetDisible(Optional bln As Boolean = False)
'��������Ϊ�����޸�״̬
    cboNO.Locked = Not bln
    txtPatient.Locked = Not bln
    cboSex.Locked = Not bln
    txt����.Locked = Not bln
    cbo���䵥λ.Locked = Not bln
    
    cbo�ѱ�.Locked = Not bln
    cboҽ�Ƹ���.Locked = Not bln
    
    cbo��������.Locked = Not bln
    cbo������.Locked = Not bln
    cbo��������.Enabled = bln
    cbo������.Enabled = bln
    
    chk�Ӱ�.Enabled = bln
    cbo���㷽ʽ.Locked = Not bln
    txtDate.Enabled = bln
    fraStat.Enabled = bln
    Bill.Active = bln
    
    If Not bln Then
        txt�ɿ�.BackColor = &HE0E0E0
        txtPatient.BackColor = &HE0E0E0
        txt����.BackColor = &HE0E0E0
    Else
        txt�ɿ�.BackColor = &HFFFFFF
        txtPatient.BackColor = &HFFFFFF
        txt����.BackColor = &HFFFFFF
    End If
End Sub

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

Private Sub SetDeptDoctorByRegevent(ByVal lng����ID As Long, _
    Optional strRegNO As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݲ���ID��Һŵ��в��˵ĹҺſ��Һ�ҽ����Ϣ���ÿ������ҺͿ�����
    '���:lng����ID-����ID
    '     strRegNO-�Һŵ���
    '����:���˺�
    '����:2014-06-06 17:38:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
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

Private Sub ShowWelcomeByLed()
'����:��ʾ��ӭ��Ϣ
    If mbytInState = 0 And gblnLED Then
        zl9LedVoice.Reset com
        zl9LedVoice.Speak "#1"
        zl9LedVoice.Init UserInfo.��� & " �շ�ԱΪ������", mlngModul, gcnOracle
    End If

End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim i As Integer, intNum As Long
    Dim rsTemp As ADODB.Recordset
    Dim lng����ID As Long, strPati As String
    Dim objDetail As Detail, blnCancel As Boolean
    Dim blnCard As Boolean, blnICCard As Boolean, blnIDCard As Boolean
    
    On Error Resume Next
    
    If txtPatient.Locked Then Exit Sub
    
    '����:51488
    If (IDKind.Cards.������� = "�ո��" Or IDKind.Cards.������� = " ") And Chr(KeyAscii) = " " Then KeyAscii = 0: Exit Sub

    '�����ַ�������Form_KeyPress�н���
    If IDKind.GetCurCard.���� Like "����*" Then
        '103563,ֻҪ����ĵ�һ���ַ��ǡ�-+*����������ȫ���֣�����Ϊ����ˢ��
        If Not (InStr("-+*", Left(txtPatient.Text, 1)) > 0 And IsNumeric(Mid(txtPatient.Text, 2))) Then
            blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.ShowPassText)
        End If
    ElseIf IDKind.GetCurCard.���� = "�����" Or IDKind.GetCurCard.���� = "סԺ��" Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If InStr("0123456789-*+", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
        End If
     Else
        txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
        '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
        txtPatient.IMEMode = 0
    End If
    
    '���˻س���ˢ��ִ�б����̺�Ͳ���ִ��Validate�¼�
    If blnCard And Len(txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 Then
        mblnKeyPress = True
    Else
        mblnKeyPress = False   'ˢ���￨ʱ�������validate�¼������ô˱���,������Ҫ��������
    End If
    
    
    '�������벡��(�������ֱ�ʶ)����:סԺ�����շ�ʱ�ɵ���ѡ����'@
    '--------------------------------------------------------------------------------------------------------------------
    If KeyAscii = 13 And gint������Դ = 2 And mbytInState = 0 And txtPatient.Text = "" And Not mblnValid Then
        frmPatiSelect.Show 1, Me
        If frmPatiSelect.mlngPatient = 0 Then Exit Sub
        txtPatient.Text = "-" & frmPatiSelect.mlngPatient
    End If
    
    If blnCard And Len(txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Trim(txtPatient.Text) <> "" Then
        If gint������Դ = 1 And InStr(mstrPrivs, "�����ҽ������") = 0 Then
            Call ClearPatientInfo: Exit Sub
        End If
        
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        End If
        KeyAscii = 0
        
        '����δ�ı��˳�
        If mrsInfo.State = 1 Then
            If txtPatient.Text = mrsInfo!���� Then
                If Not mblnValid Then Call zlCommFun.PressKey(vbKeyTab)
                 Exit Sub
            End If
        End If
        
        '��ȡ������Ϣ
        txt�Ҳ�.Text = "0.00": txt�ɿ�.Text = "0.00": sta.Panels(2) = ""
        
        '�շѱ��ֲ���ID
        If Val(txtPatient.Tag) <> 0 And txtPatient.Text = mstrPrePati Then
            strPati = "-" & Val(txtPatient.Tag)
        Else
            strPati = txtPatient.Text
        End If
        
        If IDKind.GetCurCard.���� Like "IC��*" And IDKind.GetCurCard.ϵͳ Then blnICCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
        If IDKind.GetCurCard.���� Like "*���֤*" And IDKind.GetCurCard.ϵͳ Then blnIDCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
        
        If Not GetPatient(strPati, blnCancel, blnCard) Then
            Call ReInitPatiInvoice
            If blnCancel Then 'ȡ������
                If Visible Then txtPatient.SetFocus
                txtPatient.Text = ""
                Exit Sub
            End If
            If blnCard Then
                MsgBox "����ȷ��������Ϣ�������Ƿ���ȷˢ����", vbInformation, gstrSysName
                Call ClearPatientInfo
                 Exit Sub
            ElseIf gint������Դ = 1 And gblnInputName Then
                If mstrInNO = "" Then
                    If Not CheckRegisted(0) Then
                       txtPatient.Text = "": Exit Sub
                    End If
                End If
                
                sta.Panels(2) = "����ı�ʶ���ܶ�ȡ������Ϣ����Ĭ��Ϊ�²���������"
                mobjBill.����ID = 0: mobjBill.��ʶ�� = 0: mobjBill.��ҳID = 0
                txtPatient.PasswordChar = ""
                '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
                txtPatient.IMEMode = 0
                If mstrInNO <> "" And Not Bill.Active Then
                    '����ʱ������Ҳ������ѱ�
                    cbo�ѱ�.Locked = True 'ʵ�ʴ�ʱ�Ѳ��ɼ�
                Else
                    cbo�ѱ�.Locked = False
                    If Not mblnValid Then 'ͬһ�����˲����÷ѱ�
                        If Not (Bill.Active And txtPatient.Text = mstrPrePati) Then Call LoadAndSeek�ѱ�
                    End If
                End If
                
                cboҽ�Ƹ���.Locked = False
                
                'Ԥ����Ϣ��ʼ
                lblDeposit.ForeColor = &H808080
                txtԤ�����.Enabled = False: txtԤ�����.ForeColor = &H808080: txtԤ�����.Text = "0.00"
                sta.Panels(3).Tag = "": sta.Panels(3).Text = "": sta.Panels(3).Visible = False
                txtPatient.Tag = ""
                If Bill.Active Then
                    If txtPatient.Text = mstrPrePati Then
                        mobjBill.���� = txtPatient.Text
                        mobjBill.�Ա� = Mid(cboSex.Text, InStr(cboSex.Text, "-") + 1)
                        mobjBill.���� = Trim(txt����.Text) & IIf(IsNumeric(txt����.Text), cbo���䵥λ.Text, "")
                        mobjBill.�ѱ� = Mid(cbo�ѱ�.Text, InStr(cbo�ѱ�.Text, "-") + 1)
                        If Not mblnValid Then Bill.SetFocus
                         Exit Sub
                    Else
                        '��ͬ���շѲ���
                        '���ҽ������
                        If gbyt����ҽ�� = 0 And mstrInNO = "" Then
                            cbo������.ListIndex = -1: cbo��������.ListIndex = -1
                            mobjBill.Pages(1).������ = "":  mobjBill.Pages(1).��������ID = 0
                        End If
                        
                        sta.Panels(3).Tag = "": sta.Panels(3).Text = "": sta.Panels(3).Visible = False
                        mobjBill.���� = txtPatient.Text
                        txt����.Text = ""
                        
                        '���Խɿ���Ϊ����ʱ,��ʹ��ͬ�Ĳ���Ҳ�����շ�
                        '���Ǹպýɿ����(mstrPrePati = "")
                        '���˺�:22343
                        If gTy_Module_Para.byt�ɿ���� <> 1 Or mstrPrePati = "" Then
                            Call ClearMoney
                            mintBillNO = 0: mintMoneyRow = 0
                            mcurBillʵ�� = 0: mcurBillӦ�� = 0: mcurBillӦ�� = 0
                            txt�Ҳ�.Text = "0.00": txt�ɿ�.Text = "0.00": txtӦ��.Text = "0.00"
                            If mobjBill.Pages(1).Details.Count = 0 Then
                                mstrPrePati = ""
                                If mstrInNO = "" Then
                                    txt�ϼ�.Text = gstrDec: txtӦ��.Text = gstrDec
                                End If
                            Else
                                Call ShowMoney
                            End If
                        End If
                        If Not mblnValid Then Call zlCommFun.PressKey(vbKeyTab)
                        
                        'LED��ʼ��
                        If Not mblnValid Then ShowWelcomeByLed
                        
                        Exit Sub
                    End If
                Else
                    '���������۵����벻ͬ����
                    '���˺�:22343
                    If txtPatient.Text <> mstrPrePati And mstrPrePati <> "" And gTy_Module_Para.byt�ɿ���� <> 1 Then
                        mcurBillʵ�� = 0: mcurBillӦ�� = 0: mcurBillӦ�� = 0
                        
                        txtӦ��.Text = Format(CalcBillToTal(True), gstrDec)
                        txt�ϼ�.Text = Format(CalcBillToTal, gstrDec)
                        txtӦ��.Text = Format(mobjBill.Pages(1).Ӧ�ɽ��, "0.00")
                        
                        '���������б�
                        For i = mshMoney.Rows - 1 To 1 Step -1
                            If mshMoney.TextMatrix(i, 0) <> "" Then
                                intNum = Val(mshMoney.TextMatrix(i, 0))
                                Exit For
                            End If
                        Next
                        If intNum > 1 Then
                            mintBillNO = 0
                            
                            mshMoney.Redraw = False
                            For i = mshMoney.Rows - 1 To 1 Step -1
                                If Val(mshMoney.TextMatrix(i, 0)) <> intNum Then
                                    mshMoney.RemoveItem i
                                Else
                                    mshMoney.TextMatrix(i, 0) = 1
                                End If
                            Next
                            If mshMoney.Rows < 5 Then mshMoney.Rows = 5
                            mshMoney.Redraw = True
                        End If
                    End If
                    If Not mblnValid Then Call zlCommFun.PressKey(vbKeyTab)
                     Exit Sub
                End If
            Else
                MsgBox "����ȷ��������Ϣ��", vbInformation, gstrSysName
                Call ClearPatientInfo
                If Not mblnValid Then txtPatient.SetFocus
                 Exit Sub
            End If
        '��ȷ�����˲�����Ϣ
        Else
            lng����ID = Val("" & mrsInfo!����ID)
            If mbytInState = 0 And mstrInNO = "" And gint������Դ = 1 Then
                If Not CheckRegisted(lng����ID) Then
                    Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "": txtPatient.SetFocus: Exit Sub
                End If
            End If
            '���￨������
            If mbytInState = 0 And (blnCard Or blnICCard Or blnIDCard Or IDKind.GetCurCard.�ӿ���� <> 0) And mstrPassWord <> "" Then
                If Mid(gstrCardPass, 3, 1) = "1" Then
                    If Not zlCommFun.VerifyPassWord(Me, mstrPassWord, mrsInfo!����, mrsInfo!�Ա�, "" & mrsInfo!����) Then
                        Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "": txtPatient.SetFocus: Exit Sub
                    End If
                End If
            End If
            
            '102234,������Ҳ����ӿ�
            If PatiValiedCheckByPlugIn(mlngModul, lng����ID) = False Then
                Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "": txtPatient.SetFocus: Exit Sub
            End If
        
            '����������ʱ,�����˺Ϳ������ҵĴ���
            '-----------------------------------------------------------------
            If mbytInState = 0 And mstrInNO = "" Then
                '����ͬһ������ʱ���ҽ��
                If Not (Nvl(mrsInfo!����) = mstrPrePati And Nvl(mrsInfo!����) <> "") Then
                    If gbyt����ҽ�� = 0 And mstrInNO = "" Then
                        cbo������.ListIndex = -1
                        cbo��������.ListIndex = -1
                        mobjBill.Pages(1).������ = ""
                        mobjBill.Pages(1).��������ID = 0
                    End If
                End If
            
                '�ɹҺŵ�����ʱ��ִ�в���
                If IsRegisterDept Then
                    If IsNull(mrsInfo!����) Then 'û�н���,�����˺�,���ݹҺŵ��������˺Ϳ�������
                        Call SetDeptDoctorByRegevent(0, txtPatient.Text)
                        sta.Panels(2) = "�ò��˹Һ�ʱû�еǼǵ���,�����벡��������"
                        Call ClearPatientInfo
                        
                        Set mrsInfo = New ADODB.Recordset
                        If Not mblnValid And Visible Then txtPatient.SetFocus
                        Exit Sub
                    Else
                        Call Set�����˿�������Click(mrsInfo!ִ���� & "", Val("" & mrsInfo!ִ�в���ID))
                    End If
                ElseIf gint������Դ = 2 Then
                    If gbyt����ҽ�� <> 0 Then
                        'ȡסԺ���˵Ŀ�������:����ȷ��ҽ������Զ�������
                        Call zlControl.CboSetIndex(cbo��������.hWnd, cbo.FindIndex(cbo��������, Val("" & mrsInfo!��ǰ����id)))
                        Call cbo��������_Click
                        
                    End If
                ElseIf gint������Դ = 1 Then
                    Call SetDeptDoctorByRegevent(lng����ID) '��ͼ���Ҳ��˵ĹҺſ��Һ�ҽ��
                End If
            End If
            
            'Ԥ����Ϣ
            Set rsTemp = GetMoneyInfo(lng����ID, 0, False, 1, False, 0, True)
            Dim dbl������� As Double, dbl������� As Double, dbl����Ԥ�� As Double
            Do While Not rsTemp.EOF
                If Nvl(rsTemp!����, 0) = 0 Then
                    dbl������� = Val(Nvl(rsTemp!Ԥ�����)) - Val(Nvl(rsTemp!�������))
                Else
                    dbl������� = Val(Nvl(rsTemp!Ԥ�����)) - Val(Nvl(rsTemp!�������))
                End If
                dbl����Ԥ�� = dbl����Ԥ�� + (Val(Nvl(rsTemp!Ԥ�����)) - Val(Nvl(rsTemp!�������)))
                rsTemp.MoveNext
            Loop
            sta.Panels(3).Tag = dbl����Ԥ��
            sta.Panels(3).Text = "Ԥ��:" & Format(dbl������� + dbl�������, "0.00") & _
                IIf(dbl������� > 0, "(������:" & Format(dbl�������, "0.00") & ")", "")
                
            'Ԥ����Ϣ��ʼ
            If Val(sta.Panels(3).Tag) > 0 Then
                lblDeposit.ForeColor = 0
                txtԤ�����.Enabled = True
                txtԤ�����.ForeColor = 0
                txtԤ�����.Text = "0.00"
                sta.Panels(3).Visible = True
            Else
                lblDeposit.ForeColor = &H808080
                txtԤ�����.Enabled = False
                txtԤ�����.ForeColor = &H808080
                txtԤ�����.Text = "0.00"
                sta.Panels(3).Tag = ""
                sta.Panels(3).Text = ""
                sta.Panels(3).Visible = False
            End If
            
            txtPatient.Text = IIf(IsNull(mrsInfo!����), "", mrsInfo!����)
            cboSex.ListIndex = cbo.FindIndex(cboSex, IIf(IsNull(mrsInfo!�Ա�), "", mrsInfo!�Ա�), True)
            Call LoadOldData("" & mrsInfo!����, txt����, cbo���䵥λ)
            If Not IsNull(mrsInfo!��������) Then
                 txt����.Text = ReCalcOld(mrsInfo!��������, cbo���䵥λ, lng����ID)
            End If
            
            If Not mblnValid Then
                If Not (mrsInfo.Fields(mrsInfo.Fields.Count - 1).Name = "ִ�в���ID" _
                    And cbo��������.ListIndex <> -1) Then
                    If mstrInNO <> "" And Not Bill.Active Then
                        '����ʱ������Ҳ������ѱ�
                        cbo�ѱ�.Locked = True 'ʵ�ʴ�ʱ�Ѳ��ɼ�
                    Else
                        Call LoadAndSeek�ѱ�
                    End If
                Else
                    '�Һ�ʱȷ���ķѱ�
                    cbo�ѱ�.ListIndex = cbo.FindIndex(cbo�ѱ�, IIf(IsNull(mrsInfo!�ѱ�), "", mrsInfo!�ѱ�), True)
                End If
            End If
            
            cboҽ�Ƹ���.ListIndex = cbo.FindIndex(cboҽ�Ƹ���, Nvl(mrsInfo!ҽ�Ƹ��ʽ), True)
            cboҽ�Ƹ���.Locked = gint������Դ = 2 'Or (cboҽ�Ƹ���.ListIndex <> -1)

            If gstr�ѱ� <> "" And cbo�ѱ�.ListIndex = -1 Then cbo�ѱ�.ListIndex = cbo.FindIndex(cbo�ѱ�, gstr�ѱ�, True)
            If mstr���ʽ <> "" And cboҽ�Ƹ���.ListIndex = -1 Then cboҽ�Ƹ���.ListIndex = cbo.FindIndex(cboҽ�Ƹ���, mstr���ʽ, True)

            txtPatient.PasswordChar = ""
            '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
            txtPatient.IMEMode = 0
            txtPatient.Tag = lng����ID
            
            '��д�����еĲ�����Ϣ
            With mobjBill
                .����ID = lng����ID
                .��ҳID = Nvl(mrsInfo!��ҳID, 0)
                .��ʶ�� = IIf(gint������Դ = 2, Nvl(mrsInfo!סԺ��, 0), Nvl(mrsInfo!�����, 0))
                .����ID = Nvl(mrsInfo!��ǰ����ID, 0)
                .����ID = Nvl(mrsInfo!��ǰ����id, 0)
                .���� = "" & mrsInfo!��ǰ����
                .���� = txtPatient.Text
                .�Ա� = Nvl(mrsInfo!�Ա�)
                .���� = Trim(txt����.Text) & IIf(IsNumeric(txt����.Text), cbo���䵥λ.Text, "")
                .�ѱ� = zlStr.NeedName(cbo�ѱ�.Text) '�Ե�ǰ��ЧΪ׼
            End With
            Call ReInitPatiInvoice
            If Bill.Active Then
                If txtPatient.Text = mstrPrePati And txtPatient.Text <> "" Then
                    'ͬһ������
                    If Not mblnValid Then Bill.SetFocus
                     Exit Sub
                Else
                    '��ͬ�Ĳ���
                    '��ͬ���շѲ���:������Խɿ���Ϊ����,��ͬ����Ҳ�����շ�
                    '���Ǹպýɿ����(mstrPrePati = "")
                    '���˺�:22343
                    If gTy_Module_Para.byt�ɿ���� <> 1 Or mstrPrePati = "" Then
                        Call ClearMoney
                        mintBillNO = 0: mintMoneyRow = 0
                        mcurBillʵ�� = 0: mcurBillӦ�� = 0: mcurBillӦ�� = 0
                        If mobjBill.Pages(1).Details.Count = 0 Then
                            mstrPrePati = ""
                            If mstrInNO = "" Then
                                txt�ϼ�.Text = gstrDec: txtӦ��.Text = gstrDec
                            End If
                        Else
                            Call ShowMoney
                        End If
                    End If
                    
                    '�������￨������
                    If mstrCardNO = "" Then
                        Set objDetail = ReadPatiCardObj(mobjBill.����ID, mstrCardNO)
                        If mstrCardNO <> "" And Not objDetail Is Nothing Then
                            If Not ItemExist(objDetail.ID) Then
                                If mobjBill.Pages(1).Details.Count >= Bill.Rows - 1 Then
                                    Bill.Rows = Bill.Rows + 1
                                    Call bill_AfterAddRow(Bill.Rows - 1)
                                End If
                                Bill.TextMatrix(Bill.Rows - 1, 1) = "" '�б�Ҫ����
                                Call SetDetail(objDetail, Bill.Rows - 1)
                                Call CalcMoneys(Bill.Rows - 1)
                                Call ShowDetails(Bill.Rows - 1)
                                Call ShowMoney
                            End If
                        End If
                    End If
                End If
            Else
                '���������۵����벻ͬ����
                '���˺�:22343
                If txtPatient.Text <> mstrPrePati And mstrPrePati <> "" And gTy_Module_Para.byt�ɿ���� <> 1 Then
                    mcurBillʵ�� = 0: mcurBillӦ�� = 0: mcurBillӦ�� = 0
                    
                    txtӦ��.Text = Format(CalcBillToTal(True), gstrDec)
                    txt�ϼ�.Text = Format(CalcBillToTal, gstrDec)
                    txtӦ��.Text = Format(mobjBill.Pages(1).Ӧ�ɽ��, "0.00")
                    
                    '���������б�
                    For i = mshMoney.Rows - 1 To 1 Step -1
                        If mshMoney.TextMatrix(i, 0) <> "" Then
                            intNum = Val(mshMoney.TextMatrix(i, 0))
                            Exit For
                        End If
                    Next
                    If intNum > 1 Then
                        mintBillNO = 0
                        
                        mshMoney.Redraw = False
                        For i = mshMoney.Rows - 1 To 1 Step -1
                            If Val(mshMoney.TextMatrix(i, 0)) <> intNum Then
                                mshMoney.RemoveItem i
                            Else
                                mshMoney.TextMatrix(i, 0) = 1
                            End If
                        Next
                        If mshMoney.Rows < 5 Then mshMoney.Rows = 5
                        mshMoney.Redraw = True
                    End If
                End If
            End If
            
            If Not mblnValid Then
                If cboҽ�Ƹ���.ListIndex = -1 And gblnҽ�Ƹ��� Then
                    cboҽ�Ƹ���.SetFocus
                Else
                    If gbyt����ҽ�� = 0 Then
                        cbo������.SetFocus
                    Else
                        cbo��������.SetFocus
                    End If
                End If
            End If
            
            'LED��ʼ��
            If Not mblnValid Then ShowWelcomeByLed
        End If
    End If
    mblnValid = False
End Sub

Private Function GetPatient(ByVal strInput As String, Optional blnCancel As Boolean, Optional ByVal blnCard As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ
    '���:blnCard=�Ƿ���￨ˢ��
    '����:
    '����:���˺�
    '����:2011-08-03 16:50:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, lng�����ID As Long, strPassWord As String, strErrMsg As String
    Dim lng����ID As Long, blnHavePassWord As Boolean
    Dim strWhere As String
    Dim rsTmp As ADODB.Recordset, strPati As String
    Dim vRect As RECT
    
    blnCancel = False
    
    '���������Ȩ��
    If gint������Դ = 1 Then
        'strWhere = " And Nvl(A.��ǰ����ID,0)=0"
        strWhere = " And Not Exists(Select 1 From ������ҳ Where ����ID=A.����ID And ��ҳID<>0 And ��ҳID=A.��ҳID And Nvl(��������,0)=0 And ��Ժ���� is Null)"
    ElseIf gint������Դ = 2 Then
        strWhere = " And Nvl(A.��ǰ����ID,0)<>0"
    End If
    
    '��ȡ������Ϣ
    '76451,Ƚ����,2014-8-19
    strSQL = "Select Decode(Sign(A.����ʱ��-A.�Ǽ�ʱ��),0,1,0) as ����,A.����ID,A.��������,A.����," & _
        IIf(gint������Դ = 1, "NULL", "Decode(A.��ǰ����ID,NULL,NULL,A.��ҳID)") & " as ��ҳID,A.���￨��,A.����֤��,A.�����," & _
        " A.סԺ��,A.����,A.�Ա�,A.����,A.��������,A.�ѱ�,A.������," & _
        " A.ҽ�Ƹ��ʽ,A.��ǰ����ID,A.��ǰ����ID,A.��ǰ����" & _
        " From ������Ϣ A Where A.ͣ��ʱ�� is NULL"
    
    '��������ˢ����, �˳�
    If blnCard And gint������Դ = 1 And Not gblnInputCard Then Set mrsInfo = New ADODB.Recordset: Exit Function
    
    If blnCard = True And IDKind.GetCurCard.���� Like "����*" And InStr("-+*", Left(strInput, 1)) = 0 Then '103563
        If IDKind.Cards.��ȱʡ������ And Not IDKind.GetfaultCard Is Nothing Then
            lng�����ID = IDKind.GetfaultCard.�ӿ����
        Else
            lng�����ID = "-1"
        End If
        '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�);��
        If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg, lng�����ID) = False Then GoTo NotFoundPati:
        If lng����ID <= 0 Then GoTo NotFoundPati:
        strInput = "-" & lng����ID
        blnHavePassWord = True
        strSQL = strSQL & strWhere & " And A.����ID=[1] "
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then '����ID
        If gint������Դ = 1 And Not gblnInputID And Not (mstrInNO <> "" And mbytInState = 0) Then
            Set mrsInfo = New ADODB.Recordset: Exit Function
        End If
        strSQL = strSQL & strWhere & " And A.����ID=[1]"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '�����
        If gint������Դ = 1 And Not gblnInputID Then Set mrsInfo = New ADODB.Recordset: Exit Function
        
        strSQL = strSQL & strWhere & " And A.�����=[1]"
        '75087,Ƚ����,2014-7-29,���ﲡ���շ�ʱ,����Ҫ���������������,ֻ��Ҫ��������ŵ����˳��ż����ҵ��������Ĳ�����Ϣ������
        strInput = "*" & zlCommFun.GetFullNO(Mid(strInput, 2), 3)
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then 'סԺ��
        If gint������Դ = 1 And Not gblnInputID Then Set mrsInfo = New ADODB.Recordset: Exit Function
        
        strSQL = strSQL & strWhere & " And A.����ID = (Select Max(����id) From ������ҳ Where סԺ�� = [1])"
    ElseIf Left(strInput, 1) = "." Then '�Һŵ���(���Ϊִ�в���ID������)
        If gint������Դ = 1 And Not gblnInputNO Then Set mrsInfo = New ADODB.Recordset: Exit Function
        '���ջ���˳���Ź���
        strInput = GetFullNO(Mid(strInput, 2), 12)
        txtPatient.Text = strInput
        
        '76451,Ƚ����,2014-8-19
        strSQL = "" & _
        "Select Decode(Sign(A.����ʱ��-A.�Ǽ�ʱ��),0,1,0) as ����,A.����ID,A.��������,A.����," & _
                IIf(gint������Դ = 1, "NULL", "Decode(A.��ǰ����ID,NULL,NULL,A.��ҳID)") & " as ��ҳID,A.���￨��,A.����֤��,Nvl(B.��ʶ��,A.�����) as �����," & _
        "       A.סԺ��,B.����,B.�Ա�,B.����,A.��������,B.�ѱ�,A.������,A.ҽ�Ƹ��ʽ,A.��ǰ����ID,A.��ǰ����ID,A.��ǰ����,B.ִ����,B.ִ�в���ID" & _
        " From ������Ϣ A,������ü�¼ B" & _
        " Where B.����ID=A.����ID(+) And B.��¼����=4 And B.��¼״̬=1" & _
            zlGetRegEventsCons("�Ӱ��־", "B") & _
                strWhere & " And B.NO=[2]"
    Else
    
        Select Case IDKind.GetCurCard.����
        Case "����", "��������￨"
                If mrsInfo.State = 1 Then
                    If mrsInfo!���� = strInput Then GetPatient = True: Exit Function
                End If
                'ͨ������ģ�����Ҳ���(�������벡�˱�ʶʱ)
                If Not mblnValid And gblnSeekName And gblnInputID Then
                    strWhere = " And A.���� Like '" & strInput & "%' " & strWhere
                    strPati = _
                    " Select /*+Rule */1 as ����ID,A.����ID as ID,A.����ID,A.����,A.�Ա�,A.����," & _
                    IIf(gint������Դ = 2, "A.סԺ��,B.���� as ����,A.��ǰ���� as ����,", "A.�����,") & _
                    " A.��������,A.���֤��,A.��ͥ��ַ,A.������λ" & _
                    " From ������Ϣ A,���ű� B" & _
                    " Where A.ͣ��ʱ�� is NULL And A.��ǰ����ID=B.ID(+) And Rownum<101 " & strWhere & _
                    IIf(gintNameDays = 0, "", " And (A.����ʱ��>Trunc(Sysdate-" & gintNameDays & ") Or A.�Ǽ�ʱ��>Trunc(Sysdate-" & gintNameDays & "))")
                    
                    '���ﲡ���շ�ʱ���Բ���Ӧ���˵���
                    If gint������Դ = 1 Then
                        strPati = strPati & " Union ALL " & _
                            "Select 0,0 as ID,-NULL,'[�²���]',NULL,NULL,-NULL,To_Date(NULL),NULL,NULL,NULL From Dual"
                    End If
                    strPati = strPati & " Order by ����ID,����"
                        
                    vRect = zlControl.GetControlRect(txtPatient.hWnd)
                    Set rsTmp = zlDatabase.ShowSelect(Me, strPati, 0, "����0" & gint������Դ, , , , , , True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, , True, 1)
                    If Not rsTmp Is Nothing Then
                        If rsTmp!ID = 0 Then '�����²���
                            strSQL = ""
                        Else '�Բ���ID��ȡ
                            strInput = rsTmp!����ID
                            strSQL = strSQL & strWhere & " And A.����ID=[2]"
                        End If
                    Else 'ȡ��ѡ��
                        strSQL = ""
                    End If
                Else
                    strSQL = ""
                End If
        Case "ҽ����"
            strInput = UCase(strInput)
             strSQL = strSQL & strWhere & "  And A.ҽ����=[2]"
        Case "���֤��", "�������֤", "���֤"
            strInput = UCase(strInput)
            If gobjSquare.objSquareCard.zlGetPatiID("���֤", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
            strInput = "-" & lng����ID
            strSQL = strSQL & strWhere & " And A.����ID=[2]"
        Case "IC����", "IC��"
            strInput = UCase(strInput)
            If gobjSquare.objSquareCard.zlGetPatiID("IC��", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
            strInput = "-" & lng����ID
            strSQL = strSQL & strWhere & " And A.����ID=[2]"
        Case "�����"
            If gint������Դ = 1 And Not gblnInputID Then
                Set mrsInfo = New ADODB.Recordset
                Exit Function
            End If
            If Not IsNumeric(strInput) Then strInput = "0"
            If gint������Դ = 1 Then strWhere = ""
            strSQL = strSQL & strWhere & " And A.�����=[2]"
            '75087,Ƚ����,2014-7-29,���ﲡ���շ�ʱ,����Ҫ���������������,ֻ��Ҫ��������ŵ����˳��ż����ҵ��������Ĳ�����Ϣ������
            strInput = zlCommFun.GetFullNO(strInput, 3)
        Case "סԺ��"
            If gint������Դ = 1 And Not gblnInputID Then
                Set mrsInfo = New ADODB.Recordset
                Exit Function
            End If
            If Not IsNumeric(strInput) Then strInput = "0"
            If gint������Դ = 1 Then strWhere = ""
            strSQL = strSQL & strWhere & " And A.����ID = (Select Max(����id) From ������ҳ Where סԺ�� = [2])"
        Case Else
            '��������,��ȡ��صĲ���ID
            If IDKind.GetCurCard.�ӿ���� > 0 Then
                lng�����ID = IDKind.GetCurCard.�ӿ����
                If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                If lng����ID = 0 Then GoTo NotFoundPati:
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
    '75259:���ϴ�,2014-7-10������������ɫ����
    If strSQL = "" Then GoTo NotFoundPati:
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput)
    If mrsInfo.EOF Then GoTo NotFoundPati:
    Call SetPatiColor(txtPatient, Nvl(mrsInfo!��������), IIf(IsNull(mrsInfo!����), &HC00000, vbRed))
    mstrPassWord = strPassWord
    If Not blnHavePassWord Then mstrPassWord = Nvl(mrsInfo!����֤��)
    GetPatient = True
    Exit Function
NotFoundPati:
    txtPatient.ForeColor = &HC00000
    Set mrsInfo = New ADODB.Recordset
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Set mrsInfo = New ADODB.Recordset
End Function

Private Sub txtPatient_LostFocus()
    If mbytInState = 0 And Trim(txtPatient.Text) <> "" Then
        mobjBill.���� = txtPatient.Text
        mobjBill.���� = Trim(txt����.Text) & IIf(IsNumeric(txt����.Text), cbo���䵥λ.Text, "")
        mobjBill.�Ա� = zlStr.NeedName(cboSex.Text)
    End If
    zlCommFun.OpenIme False
    mblnKeyPress = False
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
    If gint������Դ = 1 And InStr(mstrPrivs, "�����ҽ������") = 0 And txtPatient.Text <> "" And Not txtPatient.Locked Then
        Call ClearPatientInfo:  Exit Sub
    End If
    
    If Not mblnKeyPress Then
        mblnValid = True: Call txtPatient_KeyPress(13): mblnValid = False
    End If
End Sub

Private Sub txt�ɿ�_Change()
    If Val(txt�ɿ�.Text) = 0 Then txt�Ҳ�.Text = "0.00": Exit Sub
    txt�Ҳ�.Text = Format(Val(txt�ɿ�.Text) - Val(txtӦ��.Text), "0.00")
End Sub

Private Sub txt�ɿ�_GotFocus()
    '�޸�ʱ���ܹ�����
    If mbytInState = 0 And mobjBill.Pages(1).Details.Count <> 0 And gTy_Module_Para.bln������ Then 'And mstrInNO = "" Then
        Call SetFactMoney
    End If
    
    'ֻ�Խɿ���Ϊ�շѽ�������ʱ,��������ɿ��0
    '���˺�:22343
    If mbytInState = 0 And (gTy_Module_Para.byt�ɿ���� = 1 Or gTy_Module_Para.byt�ɿ���� = 2) Then
        If Val(txt�ɿ�.Text) = 0 And Me.ActiveControl Is txt�ɿ� Then
            txt�ɿ�.Text = ""
        End If
    End If
    
    Call zlControl.TxtSelAll(txt�ɿ�)
    
    'LED��ʾ
    If mbytInState = 0 And gblnLED Then
        '�Զ����ۻ��ֹ�����ʱ���ȼ�����
        If (Not gbln�ֹ����� And ActiveControl Is txt�ɿ�) Or (gbln�ֹ����� And mblnHotKey) Then
            mblnHotKey = False
            zl9LedVoice.Speak "#21 " & txtӦ��.Text
            mbln���ϼ� = True
        End If
    End If
End Sub

Private Sub txt�ɿ�_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'ֻ�Խɿ���Ϊ�շѽ�������ʱ,��������ɿ��0
        '���˺�:22343
        If mbytInState = 0 And (gTy_Module_Para.byt�ɿ���� = 1 Or gTy_Module_Para.byt�ɿ���� = 2) Then
            If txt�ɿ�.Text = "" Then Exit Sub
        End If
        
        If Val(txt�ɿ�.Text) = 0 Then txt�ɿ�.Text = "0.00"
        If Val(txt�ɿ�.Text) <> 0 Then
            If CSng(txt�Ҳ�.Text) >= 0 Then
                Call zlCommFun.PressKey(vbKeyTab)
                If gblnLED And CCur(txt�ϼ�.Text) <> 0 And mbytInState = 0 Then 'LED��ʾ
                    mblnHotKey = False
                    If Val(txtԤ�����.Text) = 0 Then
                        zl9LedVoice.DispCharge txtӦ��.Text, txt�ɿ�.Text, txt�Ҳ�.Text
                    Else '����֧���ֽ�ʱ�Ĵ���
                        Call zl9LedVoice.DisplayBank( _
                            "�ϼ�:" & txt�ϼ�.Text & "Ԫ,Ӧ��" & txtӦ��.Text & "Ԫ", _
                            "����:" & txt�ɿ�.Text & "Ԫ" & IIf(Val(txt�Ҳ�.Text) = 0, "", ",����:" & txt�Ҳ�.Text & "Ԫ"))
                    End If
                    
                    zl9LedVoice.Speak "#22 " & txt�ɿ�.Text
                    zl9LedVoice.Speak "#23 " & txt�Ҳ�.Text
                    zl9LedVoice.Speak "#3"
                End If
            Else
                MsgBox "�ɿ����,�벹��Ӧ�ɽ�", vbInformation, gstrSysName
                Call zlControl.TxtSelAll(txt�ɿ�): txt�ɿ�.SetFocus
            End If
        Else
            Call zlCommFun.PressKey(vbKeyTab) '�����ۼӽɿ�
        End If
    End If
    If KeyAscii = Asc(".") And InStr(txt�ɿ�.Text, ".") > 0 Then KeyAscii = 0: Beep: Exit Sub
    If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
End Sub

Private Sub CalcMoneys(Optional lngRow As Long = 0)
'���ܣ���������¼���ָ���л������еĽ��
'������lngRow=ָ����,Ϊ0��ʾ����������
'˵����ExpenseBill���ϵ�������Ӧ���ݵ��к�
    Dim i As Long
    If mobjBill.Pages(1).Details.Count = 0 Then Exit Sub
    If lngRow = 0 Then
        For i = 1 To mobjBill.Pages(1).Details.Count
            CalcMoney i
        Next
    Else
        CalcMoney lngRow
    End If
End Sub

Private Sub CalcMoney(lngRow As Long)
'���ܣ���������¼���ָ���еĽ��
'������lngRow=ָ����
'˵����1.ExpenseBill���ϵ�������Ӧ���ݵ��к�
'      2.���ֻ�ܶ�Ӧһ��������Ŀ:mobjBill.Details(lngRow).InComes(1)
'      3.������ϸĿδ�����������Ŀ(��һ�μ���),��ʹ��Ĭ���ּ�
'      4.������ϸĿ�Ѿ������������Ŀ(����2��),���ֶ�����(Ҳ����δ��)�˵���,�򰴸õ��ۼ��㡣
    Dim i As Long, strInfo As String
    Dim rsTmp As ADODB.Recordset
    Dim dblMoney As Double '�û�����ı�۽��
    Dim str�ѱ� As String
    Dim dbl�Ӱ�Ӽ��� As Double
    Dim strWherePriceGrade As String
    
    On Error GoTo errH
    If mstr��ͨ�۸�ȼ� <> "" Then
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
        " Select B.������ĿID,C.����,C.�վݷ�Ŀ,B.�ּ�,B.ԭ��,B.�Ӱ�Ӽ���,B.�����շ���,b.ȱʡ�۸� " & _
        " From �շ���ĿĿ¼ A,�շѼ�Ŀ B,������Ŀ C " & _
        " Where B.�շ�ϸĿID=A.ID And C.ID = B.������ĿID " & _
        "       And Sysdate Between B.ִ������ And Nvl(B.��ֹ����,To_Date('3000-1-1', 'YYYY-MM-DD')) " & _
        "       And A.ID=[1]" & vbNewLine & _
        strWherePriceGrade
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mobjBill.Pages(1).Details(lngRow).�շ�ϸĿID, mstr��ͨ�۸�ȼ�)
    
    If rsTmp.RecordCount > 0 Then
        With mobjBill.Pages(1).Details(lngRow)
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
        End With
        
        '�����ԭ�м�¼
        Set mobjBill.Pages(1).Details(lngRow).InComes = New BillInComes
        
        '��д���з��ü�¼
        For i = 1 To rsTmp.RecordCount
            Set mobjBillIncome = New BillInCome
            With mobjBillIncome
                .������ĿID = rsTmp!������ĿID
                .������Ŀ = rsTmp!����
                .�վݷ�Ŀ = Nvl(rsTmp!�վݷ�Ŀ)
                .ԭ�� = Val(Nvl(rsTmp!ԭ��))
                .�ּ� = Val(Nvl(rsTmp!�ּ�))
                If mobjBill.Pages(1).Details(lngRow).Detail.��� Then
                    .��׼���� = Format(dblMoney, gstrFeePrecisionFmt)
                Else
                    .��׼���� = Format(Val(Nvl(rsTmp!�ּ�)), gstrFeePrecisionFmt)
                End If
                
                'Ӧ�ս��=���� * ���� * ����
                .Ӧ�ս�� = .��׼���� * IIf(mobjBill.Pages(1).Details(lngRow).���� = 0, 1, mobjBill.Pages(1).Details(lngRow).����) * mobjBill.Pages(1).Details(lngRow).����
                
                '�������������ü���(����������Ŀ)
                If mobjBill.Pages(1).Details(lngRow).���ӱ�־ = 1 And mobjBill.Pages(1).Details(lngRow).�շ���� = "F" Then
                    .Ӧ�ս�� = .Ӧ�ս�� * IIf(IsNull(rsTmp!�����շ���), 1, rsTmp!�����շ��� / 100)
                End If
                
                '�Ӱ�����ʼ���
                dbl�Ӱ�Ӽ��� = 0
                If mobjBill.�Ӱ��־ = 1 And mobjBill.Pages(1).Details(lngRow).Detail.�Ӱ�Ӽ� Then
                    dbl�Ӱ�Ӽ��� = IIf(IsNull(rsTmp!�Ӱ�Ӽ���), 0, rsTmp!�Ӱ�Ӽ��� / 100)             '������ݷѱ����ʵ�ս���
                    .Ӧ�ս�� = .Ӧ�ս�� + .Ӧ�ս�� * dbl�Ӱ�Ӽ���
                End If
                
                .Ӧ�ս�� = CCur(Format(.Ӧ�ս��, gstrDec))
                
                If mobjBill.Pages(1).Details(lngRow).Detail.���ηѱ� Then
                    .ʵ�ս�� = .Ӧ�ս��
                    mobjBill.Pages(1).Details(lngRow).�ѱ� = mobjBill.�ѱ�
                Else
                    If .Ӧ�ս�� = 0 Then
                        .ʵ�ս�� = 0
                        mobjBill.Pages(1).Details(lngRow).�ѱ� = mobjBill.�ѱ�
                    Else
                        str�ѱ� = IIf(glngSys Like "8??", mobjBill.�ѱ�, zlStr.TrimEx(mobjBill.�ѱ� & "," & lbl��̬�ѱ�.Tag, ","))
                        
                        .ʵ�ս�� = CCur(Format(ActualMoney(str�ѱ�, .������ĿID, .Ӧ�ս��, _
                            mobjBill.Pages(1).Details(lngRow).�շ�ϸĿID, 0, 0, dbl�Ӱ�Ӽ���), gstrDec))
                        mobjBill.Pages(1).Details(lngRow).�ѱ� = str�ѱ�
                    End If
                End If
                
                'ʵ�ս�����Key��,�Դ���ֱ�����(��Key�д��ԭʼʵ�ս��,����)
                mobjBill.Pages(1).Details(lngRow).InComes.Add .������ĿID, .������Ŀ, .�վݷ�Ŀ, .��׼����, .Ӧ�ս��, .ʵ�ս��, .ԭ��, .�ּ�, "_" & .ʵ�ս��
            End With
            rsTmp.MoveNext
        Next
    Else
        '���û��������Ŀ,�������Ӧ�ĳ������
        Set mobjBill.Pages(1).Details(lngRow).InComes = New BillInComes
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ShowDetails(Optional lngRow As Long = 0)
'���ܣ�ˢ����ʾָ���л������е�����
'������lngRow=ָ����,Ϊ0��ʾ��ʾ������
'˵����ExpenseBill���ϵ�������Ӧ���ݵ��к�
    Dim i As Long
    
    Bill.Redraw = False
    If lngRow = 0 Then
        For i = 1 To mobjBill.Pages(1).Details.Count
            ShowDetail i
        Next
    Else
        ShowDetail lngRow
    End If
    Bill.Redraw = True
End Sub

Private Sub ShowDetail(lngRow As Long)
'���ܣ�ˢ����ʾָ���е�����
'������lngRow=ָ����
'˵����ExpenseBill���ϵ�������Ӧ���ݵ��к�
    Dim i As Long, j As Long, curMoney As Currency
    
    If lngRow > Bill.Rows - 1 Then Exit Sub
    
    '���������
    For i = 0 To Bill.COLS - 1
        '����ʱ�շ�������
        If Not (i = 0 And Bill.TextMatrix(lngRow, i) <> "") Then Bill.TextMatrix(lngRow, i) = ""
    Next
    'ˢ�µ�����
    For i = 0 To Bill.COLS - 1
        Select Case Bill.TextMatrix(0, i)
            Case "��Ŀ"
                Bill.TextMatrix(lngRow, i) = mobjBill.Pages(1).Details(lngRow).Detail.����
            Case "Ӧ�ս��" 'ʵ�����ǵ���
                '�����Ǹ��շ�ϸĿ����������Ŀ�ĺϼ�
                '��һ�μ���ʱ����Ĭ������Ϊ1�Ļ����ϼ��������
                curMoney = 0
                If mobjBill.Pages(1).Details(lngRow).InComes.Count > 0 Then
                    For j = 1 To mobjBill.Pages(1).Details(lngRow).InComes.Count
                        curMoney = curMoney + mobjBill.Pages(1).Details(lngRow).InComes(j).Ӧ�ս��
                    Next
                End If
                Bill.TextMatrix(lngRow, i) = Format(curMoney, gstrDec)
            Case "ʵ�ս��"
                'ʵ�ս���Ǹ��շ�ϸĿ����������Ŀ�ĺϼ�
                curMoney = 0
                If mobjBill.Pages(1).Details(lngRow).InComes.Count > 0 Then
                    For j = 1 To mobjBill.Pages(1).Details(lngRow).InComes.Count
                        curMoney = curMoney + mobjBill.Pages(1).Details(lngRow).InComes(j).ʵ�ս��
                    Next
                End If
                Bill.TextMatrix(lngRow, i) = Format(curMoney, gstrDec)
            Case "ִ�п���"
                If mbytInState = 0 Then
                    mrsUnit.Filter = "ID=" & mobjBill.Pages(1).Details(lngRow).ִ�в���ID
                    If mrsUnit.RecordCount <> 0 Then
                        Bill.TextMatrix(lngRow, i) = mrsUnit!���� & "-" & mrsUnit!����
                    Else
                        Bill.TextMatrix(lngRow, i) = GET��������(mobjBill.Pages(1).Details(lngRow).ִ�в���ID, mrsUnit)
                    End If
                Else
                    '�������ֻ(��)��ʾ����
                    Bill.TextMatrix(lngRow, i) = GET��������(mobjBill.Pages(1).Details(lngRow).ִ�в���ID, mrsUnit)
                End If
            Case "����"
                Bill.TextMatrix(lngRow, i) = mobjBill.Pages(1).Details(lngRow).Detail.����
        End Select
    Next
    Bill.Text = Bill.MsfObj.Text
End Sub

Public Sub ShowMoney()
'���ܣ�ˢ����ʾ������Ŀ������
    Dim curʵ�պϼ� As Currency, curӦ�պϼ� As Currency, cur���ϼ� As Currency
    Dim blnExist As Boolean, i As Integer, j As Integer, k As Integer
    
    '�������ܷ�Ŀ
    Set mcolMoneys = New BillInComes
    For i = 1 To mobjBill.Pages(1).Details.Count
        For j = 1 To mobjBill.Pages(1).Details(i).InComes.Count
            '�����Ƿ��Ѿ���������վݷ�Ŀ,������ϼ�,��������
            blnExist = False
            For k = 1 To mcolMoneys.Count
                If gint����ϼ� = 0 Then
                    If mcolMoneys(k).�վݷ�Ŀ = mobjBill.Pages(1).Details(i).InComes(j).�վݷ�Ŀ Then
                        blnExist = True: Exit For
                    End If
                Else
                    If mcolMoneys(k).������Ŀ = mobjBill.Pages(1).Details(i).InComes(j).������Ŀ Then
                        blnExist = True: Exit For
                    End If
                End If
            Next
            
            If blnExist Then
                mcolMoneys(k).ʵ�ս�� = mcolMoneys(k).ʵ�ս�� + mobjBill.Pages(1).Details(i).InComes(j).ʵ�ս��
                mcolMoneys(k).Ӧ�ս�� = mcolMoneys(k).Ӧ�ս�� + mobjBill.Pages(1).Details(i).InComes(j).Ӧ�ս��
            Else
                With mobjBill.Pages(1).Details(i).InComes(j)
                    mcolMoneys.Add .������ĿID, .������Ŀ, .�վݷ�Ŀ, .��׼����, .Ӧ�ս��, .ʵ�ս��
                End With
            End If
        Next
    Next
    
    'ˢ����ʾ(�շ�Ҫ����)
    mshMoney.Redraw = False
    If mcolMoneys.Count > 0 Then
        mshMoney.Rows = mcolMoneys.Count + 1 + mintMoneyRow
    End If
    If mshMoney.Rows < 5 Then mshMoney.Rows = 5
    
    Call SetMoneyList
    
    For i = mintMoneyRow + 1 To mcolMoneys.Count + mintMoneyRow
        mshMoney.TextMatrix(i, 0) = mintBillNO + 1
        If gint����ϼ� = 0 Then
            mshMoney.TextMatrix(i, 1) = mcolMoneys(i - mintMoneyRow).�վݷ�Ŀ
        Else
            mshMoney.TextMatrix(i, 1) = mcolMoneys(i - mintMoneyRow).������Ŀ
        End If
        mshMoney.TextMatrix(i, 2) = Format(mcolMoneys(i - mintMoneyRow).ʵ�ս��, gstrDec)
        curʵ�պϼ� = curʵ�պϼ� + mcolMoneys(i - mintMoneyRow).ʵ�ս��
        curӦ�պϼ� = curӦ�պϼ� + mcolMoneys(i - mintMoneyRow).Ӧ�ս��
    Next
    For i = 1 To mshMoney.Rows - 1
        If Val(mshMoney.TextMatrix(i, 0)) = mintBillNO + 1 Then
            mshMoney.TopRow = i
        End If
    Next
    mshMoney.Redraw = True
    
    '��ǰ���ݵ���ػ��ܽ�����
    '----------------------------------------
    With mobjBill.Pages(1)
        cur���ϼ� = Format(Val(txtԤ�����.Text), "0.00")
        
        .Ӧ�ս�� = curӦ�պϼ�
        .ʵ�ս�� = curʵ�պϼ�
        
        '���㵱ǰ����Ӧ�ֽ���Ľ��,Ϊ�˼���Ӧ��(�൥��ʱ�ȳ�Ԥ��)
        If cur���ϼ� <> 0 Then
            If cur���ϼ� <= Format(.ʵ�ս��, "0.00") Then
                .��Ԥ���� = cur���ϼ�
            Else
                .��Ԥ���� = Format(.ʵ�ս��, "0.00")
            End If
            cur���ϼ� = cur���ϼ� - .��Ԥ����
        End If
        
        '���㵱ǰ����Ӧ�ɽ��ֱҴ�������
        .Ӧ�ɽ�� = Format(.ʵ�ս�� - .��Ԥ����, "0.00")
        
        '�ֽ�ʽʱ�Ŵ���ֱ�
        If cbo���㷽ʽ.ListIndex <> -1 Then
            If cbo���㷽ʽ.ItemData(cbo���㷽ʽ.ListIndex) = 1 Then
                .Ӧ�ɽ�� = CentMoney(.ʵ�ս�� - .��Ԥ����)
            End If
        End If
    
        .����� = Format((.ʵ�ս�� - .��Ԥ����) - .Ӧ�ɽ��, gstrDec)
    End With
    
    txtӦ��.Text = Format(mcurBillӦ�� + curӦ�պϼ�, gstrDec)
    txt�ϼ�.Text = Format(mcurBillʵ�� + curʵ�պϼ�, gstrDec)
    txtӦ��.Text = Format(mcurBillӦ�� + mobjBill.Pages(1).Ӧ�ɽ��, "0.00")
    
    '�����ʾ
    If mobjBill.Pages(1).����� <> 0 Then
        pic���.Visible = True
        lbl����.Caption = Format(mobjBill.Pages(1).�����, "0.00")
    Else
        pic���.Visible = False
    End If
End Sub

Private Sub ShowPrice()
'���ܣ����շ���ȡ���۵�����ʱ�����㲢��ʾ���۵��ݸ��������Ϣ
    Dim cur���ϼ� As Currency
    
    With mobjBill.Pages(1)
        '���㵱ǰ����Ӧ�ֽ���Ľ��,Ϊ�˼���Ӧ��(�൥��ʱ�ȳ�Ԥ��)
        cur���ϼ� = Val(txtԤ�����.Text)
        If cur���ϼ� <> 0 Then
            If cur���ϼ� <= Format(.ʵ�ս��, "0.00") Then
                .��Ԥ���� = cur���ϼ�
            Else
                .��Ԥ���� = Format(.ʵ�ս��, "0.00")
            End If
            cur���ϼ� = cur���ϼ� - .��Ԥ����
        End If
        
        '���㵱ǰ����Ӧ�ɽ��ֱҴ�������
        .Ӧ�ɽ�� = Format(.ʵ�ս�� - .��Ԥ����, "0.00")
        
        '�ֽ�ʽʱ�Ŵ���ֱ�
        If cbo���㷽ʽ.ListIndex <> -1 Then
            If cbo���㷽ʽ.ItemData(cbo���㷽ʽ.ListIndex) = 1 Then
                .Ӧ�ɽ�� = CentMoney(.ʵ�ս�� - .��Ԥ����)
            End If
        End If
    
        .����� = Format((.ʵ�ս�� - .��Ԥ����) - .Ӧ�ɽ��, gstrDec)
        
        '��ʾ�ϼ�
        txtӦ��.Text = Format(.Ӧ�ս�� + mcurBillӦ��, gstrDec)
        txt�ϼ�.Text = Format(.ʵ�ս�� + mcurBillʵ��, gstrDec)
        txtӦ��.Text = Format(.Ӧ�ɽ�� + mcurBillӦ��, "0.00")
        
        '�����ʾ
        If .����� <> 0 Then
            pic���.Visible = True
            lbl����.Caption = Format(.�����, "0.00")
        Else
            pic���.Visible = False
        End If
    End With
End Sub

Private Function GetInputDetail(ByVal lng��Ŀid As Long) As Detail
'���ܣ���ȡ�շ���Ŀ��Ϣ
    Dim objDetail As New Detail
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    
    strSQL = _
        " Select A.ID,A.���,B.���� as �������,A.����,A.����,A.���,A.���㵥λ," & _
        " A.���ηѱ�,A.�Ƿ���,A.�Ӱ�Ӽ�,A.ִ�п���,A.��������,A.����ժҪ" & _
        " From �շ���ĿĿ¼ A,�շ���Ŀ��� B" & _
        " Where A.���=B.���� And A.ID=[1]"
        
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��Ŀid)
    With objDetail
        .ID = rsTmp!ID
        .��� = rsTmp!���
        .������� = rsTmp!�������
        .���� = rsTmp!����
        .���� = rsTmp!����
        .��� = Nvl(rsTmp!���)
        .���㵥λ = Nvl(rsTmp!���㵥λ)
        .��� = Nvl(rsTmp!�Ƿ���, 0) = 1 '��ҩƷ�����Ƿ�ʱ��
        .���� = Nvl(rsTmp!��������)
        .�Ӱ�Ӽ� = Nvl(rsTmp!�Ӱ�Ӽ�, 0) = 1
        .���ηѱ� = Nvl(rsTmp!���ηѱ�, 0) = 1
        .ִ�п��� = Nvl(rsTmp!ִ�п���, 0)
        .����ժҪ = Nvl(rsTmp!����ժҪ, 0) = 1
    End With
    Set GetInputDetail = objDetail
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SetDetail(Detail As Detail, lngRow As Long, Optional bytParent As Byte = 0, Optional ByVal lngDoUnit As Long)
'���ܣ�����ָ�����շ�ϸĿ�����趨����ָ�㶨�е��շ�ϸĿ(�����Ļ��޸�)
'˵����
'      1.���������������շ�ϸĿ�У�����
'      2.��bytParent<>0ʱ,��Ϊ���ô�����Ŀ,������Ŀһ����������,������Ŀһ������
    Dim tmpIncomes As New BillInComes
    Dim dblTime As Double, i As Long
    
        
    'ִ�п���
    If bytParent <> 0 Then
        '������Ŀ��ִ�п���,��������������ͬ,����Ϊ����ȷִ�п���,��ȡ����ִ�п���,����ȡ�����
        If lngDoUnit <> 0 Then
            lngDoUnit = mobjBill.Pages(1).Details(bytParent).ִ�в���ID
        Else
            If cbo��������.ListIndex <> -1 Then
                lngDoUnit = cbo��������.ItemData(cbo��������.ListIndex)
            End If
            lngDoUnit = Get�շ�ִ�п���ID("Z", Detail.ID, Detail.ִ�п���, lngDoUnit, Get��������ID, gint������Դ, , , , , mobjBill.����ID)
        End If
    Else
        lngDoUnit = mobjBill.����ID
        If lngDoUnit = 0 And cbo��������.ListIndex <> -1 Then
            lngDoUnit = cbo��������.ItemData(cbo��������.ListIndex)
        End If
        lngDoUnit = Get�շ�ִ�п���ID("Z", Detail.ID, Detail.ִ�п���, lngDoUnit, Get��������ID, gint������Դ, , , , , mobjBill.����ID)
    End If
    
    If mobjBill.Pages(1).Details.Count < lngRow Then
        '������ж�Ӧ�ĳ��������δ��ʼ,�����
        With Detail
            '���=�к�,����=0
            '����=1
            '����=1,������Ŀ�Ĵ������������ȷ��
            'ִ�в���ID:����ϸĿִ�п��ұ�־ȡ
            '���ӱ�־:�Ե�һ��Ϊ��,����Ϊ������Ȩ
            '���뼯=��
            If bytParent <> 0 Then
                '��ʼ����
                If Detail.���д��� = 0 Then '�ǹ��д���
                    dblTime = mobjBill.Pages(1).Details(bytParent).����
                ElseIf Detail.���д��� = 1 Then '�̶��Ĺ��д���
                    dblTime = Detail.��������
                ElseIf Detail.���д��� = 2 Then '�������Ĺ��д���
                    dblTime = Detail.�������� * mobjBill.Pages(1).Details(bytParent).����
                End If
            Else
                dblTime = 1
            End If
            
            mobjBill.Pages(1).Details.Add mobjBill.�ѱ�, Detail, .ID, CByte(lngRow), CInt(bytParent), .���, .���㵥λ, "", 1, dblTime, 0, lngDoUnit, tmpIncomes
        End With
    Else
        '��������Ѿ�����,���޸�
        With mobjBill.Pages(1).Details(lngRow)
            Set .Detail = Detail
            Set .InComes = tmpIncomes
            .�ѱ� = mobjBill.�ѱ�
            .���� = 1
            .���ӱ�־ = 0
            .���㵥λ = Detail.���㵥λ
            .�շ���� = Detail.���
            .�շ�ϸĿID = Detail.ID
            .���� = 1
            .��� = lngRow
            .�������� = 0
            .ִ�в���ID = lngDoUnit
        End With
    End If
End Sub

Private Function ShouldDO(lngRow As Long) As Boolean
'���ܣ��жϸ����Ƿ�Ӧ��ȡ������Ŀ
'˵�����������շ���Ŀ�д�����Ŀ����δȡ��ȡ��
    Dim rsTmp As ADODB.Recordset
    Dim i As Long, blnExist As Boolean
    Dim strSQL As String
    
    On Error GoTo errHandle

    strSQL = "Select count(����ID) as NUM From �շѴ�����Ŀ Where ����ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mobjBill.Pages(1).Details(lngRow).�շ�ϸĿID)
    If rsTmp.RecordCount <> 0 Then
        If IsNull(rsTmp!Num) Then
            ShouldDO = False
        ElseIf rsTmp!Num = 0 Then
            ShouldDO = False
        Else
            blnExist = False
            For i = lngRow + 1 To mobjBill.Pages(1).Details.Count
                If mobjBill.Pages(1).Details(i).�������� = lngRow Then
                    blnExist = True: Exit For
                End If
            Next
            If Not blnExist Then
                ShouldDO = True
            Else
                ShouldDO = False
            End If
        End If
    Else
        ShouldDO = False
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetSubDetails(lng��Ŀid As Long) As Details
'���ܣ�����һ���շ�ϸĿ�Ĵ�����Ŀ��
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim objDetail As New Detail
    
    Set GetSubDetails = New Details
    
    strSQL = _
        "Select A.ID,A.���,B.���� as �������," & _
        " A.��������,A.����,A.����,A.���,A.���㵥λ,A.���ηѱ�,A.�Ƿ���," & _
        " A.�Ӱ�Ӽ�,A.ִ�п���,C.���д���,C.�������� " & _
        " From �շ���ĿĿ¼ A,�շ���Ŀ��� B,�շѴ�����Ŀ C" & _
        " Where B.����=A.��� And C.����ID=A.ID And A.���='Z' And C.����ID=[1]" & _
        " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)"
        
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��Ŀid)
    For i = 1 To rsTmp.RecordCount
        Set objDetail = New Detail
        With objDetail
            .ID = rsTmp!ID
            .���� = rsTmp!����
            .��� = Nvl(rsTmp!�Ƿ���, 0) = 1
            .��� = Nvl(rsTmp!���)
            .���㵥λ = Nvl(rsTmp!���㵥λ)
            .�Ӱ�Ӽ� = Nvl(rsTmp!�Ӱ�Ӽ�, 0) = 1
            .��� = rsTmp!���
            .������� = rsTmp!�������
            .���� = rsTmp!����
            .���ηѱ� = Nvl(rsTmp!���ηѱ�, 0) = 1
            .ִ�п��� = Nvl(rsTmp!ִ�п���, 0) 'ȱʡΪ����ȷ����(�û�ѡ)
            .���д��� = Nvl(rsTmp!���д���, 0) 'ȱʡΪ�ǹ̶�,�û����������������
            .�������� = Nvl(rsTmp!��������, 1)
            .���� = Nvl(rsTmp!��������)
            
            GetSubDetails.Add .ID, .ҩ��ID, .���, .�������, .����, .����, .����, .���, .���㵥λ, .˵��, .���ηѱ�, _
                1, .���㵥λ, .����, .���, .�Ӱ�Ӽ�, .ִ�п���, .����, .����ժҪ, .���д���, .��������
        End With
        rsTmp.MoveNext
    Next
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub DeleteDetail(lngRow As Long)
'���ܣ�ɾ��ָ���շ���Ŀ��
'˵������ʱ����������е�ɾ��,��Ҫ�����������д�����ϵ����Ӧ�ĵ���
    Dim i As Long
    For i = lngRow + 1 To mobjBill.Pages(1).Details.Count
        If mobjBill.Pages(1).Details(i).�������� <> 0 And _
            mobjBill.Pages(1).Details(i).�������� > lngRow Then
            mobjBill.Pages(1).Details(i).�������� = mobjBill.Pages(1).Details(i).�������� - 1
        End If
        mobjBill.Pages(1).Details(i).��� = mobjBill.Pages(1).Details(i).��� - 1 '������кŶ�Ӧ
    Next
    mobjBill.Pages(1).Details.Remove lngRow
    If lngRow = 1 And mobjBill.Pages(1).Details.Count = 0 And Bill.Rows = 2 Then
        For i = 0 To Bill.COLS - 1
            Bill.TextMatrix(lngRow, i) = ""
        Next
    Else
        Bill.RemoveMSFItem lngRow
    End If
End Sub

Private Function NewBill(Optional blnFact As Boolean = True, Optional bln�ѱ� As Boolean = True) As Boolean
'���ܣ���ʼ��һ���µĵ���(�������)
'������blnFact=�Ƿ�ȡƱ��
'      bln�ѱ�=�Ƿ����³�ʼ���ѱ�
    Dim i As Long
    
    mblnKeyPress = False
    mblnHotKey = False
    mbln���ϼ� = False
    
    cbo�ѱ�.Locked = False
    cboҽ�Ƹ���.Locked = False
    
    Set mobjBill = New ExpenseBill
    
    mstrCardNO = ""
    
    'Ԥ����Ϣ��ʼ
    txtԤ�����.Text = "0.00"
    lblDeposit.ForeColor = &H808080
    txtԤ�����.Enabled = False
    txtԤ�����.ForeColor = &H808080
    sta.Panels(3).Tag = ""
    sta.Panels(3).Text = ""
    sta.Panels(3).Visible = False
    '�������
    pic���.Visible = False
    
    txtPatient.Locked = False
    cboSex.Locked = False
    txt����.Locked = False
    cbo���䵥λ.Locked = False
    
    txtDate.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    Set mrsInfo = New ADODB.Recordset
    
    If mbytInState = 0 Then
        cboNO.Text = ""
        
        'ʵ�ʺ���
        Call ReInitPatiInvoice(blnFact)
        
        '��Ϊִ��״̬�ų�ʼ
        chk�Ӱ�.Value = IIf(OverTime(zlDatabase.Currentdate), Checked, Unchecked)
        
        '���㷽ʽ
        i = cbo.FindIndex(cbo���㷽ʽ, gstr���㷽ʽ, True)
        If i = -1 And cbo���㷽ʽ.ListCount > 0 Then i = 0
        Call zlControl.CboSetIndex(cbo���㷽ʽ.hWnd, i)
        
        '����
        With mobjBill
            .NO = cboNO.Text
            .����Ա��� = UserInfo.���
            .����Ա���� = UserInfo.����
            .����ʱ�� = CDate(txtDate.Text)
            .�ѱ� = IIf(cbo�ѱ�.ListIndex = -1, "", Mid(cbo�ѱ�.Text, InStr(cbo�ѱ�.Text, "-") + 1))
            .�Ӱ��־ = IIf(chk�Ӱ�.Value = Checked, 1, 0)
            If cbo��������.ListIndex = -1 Then
                .Pages(1).��������ID = 0
            Else
                .Pages(1).��������ID = cbo��������.ItemData(cbo��������.ListIndex)
            End If
            .�����־ = gint������Դ
            .Pages(1).������ = IIf(cbo������.ListIndex = -1, "", zlStr.NeedName(cbo������.Text))
        End With
        
        '�ѱ���:�շ�
        cbo�ѱ�.Locked = False
        cbo�ѱ�.Visible = True
        lbl��̬�ѱ�.AutoSize = True
        lbl��̬�ѱ�.BorderStyle = 0
        lbl��̬�ѱ�.Left = cbo�ѱ�.Left + cbo�ѱ�.Width + 60
        
        If bln�ѱ� Then
            Call LoadAndSeek�ѱ�
        End If
    End If
    NewBill = True
End Function

Private Sub ClearMoney()
'���ܣ����������ʾ��
    Dim i As Integer, j As Integer
    mshMoney.Redraw = False
    For i = 1 To mshMoney.Rows - 1
        For j = 0 To mshMoney.COLS - 1
            mshMoney.TextMatrix(i, j) = ""
        Next
    Next
    mintMoneyRow = 0
    mshMoney.Rows = 5
    mshMoney.Redraw = True
End Sub

Private Function SaveBill() As Boolean
'����:���浱ǰ����ĵ���
'���:mobjBill=���ݶ���
    Dim i As Integer, j As Integer, strҽ�Ƹ��� As String
    Dim int��� As Integer, int�к� As Integer, strNo As String, strTmp As String
    Dim intParent As Integer, intParentNO As Integer
    Dim arrSQL As Variant, strDelBill As String, lng����ID As Long
    
    If cboҽ�Ƹ���.ListIndex <> -1 Then
        strҽ�Ƹ��� = Mid(cboҽ�Ƹ���.Text, 1, InStr(1, cboҽ�Ƹ���, "-") - 1)
    End If
    Err = 0: On Error GoTo Errhand:
    mobjBill.NO = zlDatabase.GetNextNo(13)
    lng����ID = zlDatabase.GetNextId("���˽��ʼ�¼")
    
    gstrModiNO = mobjBill.NO
    arrSQL = Array()
    
    For Each mobjBillDetail In mobjBill.Pages(1).Details
        intParent = 0: intParentNO = int���
        For Each mobjBillIncome In mobjBillDetail.InComes
            int��� = int��� + 1 '��ǰ��¼���
            '��������
            '76451,Ƚ����,2014-8-19
            With mobjBill
                gstrSQL = "zl_�����շѼ�¼_INSERT('" & .NO & "'," & int��� & "," & ZVal(.����ID) & "," & IIf(.��ҳID = 0, 1, ZVal(.��ҳID)) & "," & _
                    ZVal(.��ʶ��) & ",'" & IIf(gint������Դ = 2, .����, strҽ�Ƹ���) & "','" & .���� & "'," & _
                    "'" & .�Ա� & "','" & .���� & "','" & IIf(mobjBillDetail.�ѱ� = "", .�ѱ�, mobjBillDetail.�ѱ�) & "'," & _
                    .�Ӱ��־ & "," & ZVal(.����ID, , .Pages(1).��������ID) & "," & ZVal(.Pages(1).��������ID) & ",'" & .Pages(1).������ & "',"
            End With
            
            '�շ�ϸĿ����
            With mobjBillDetail
                '�����������
                If .��� <> int�к� Then
                    int�к� = .���
                    '���´����������
                    If mobjBill.Pages(1).Details(.���).�������� = 0 Then
                        For i = .��� + 1 To mobjBill.Pages(1).Details.Count
                            If mobjBill.Pages(1).Details(i).�������� = .��� Then
                                mobjBill.Pages(1).Details(i).�������� = int��� '������Ŀ�ж��������Ŀ(������)ʱ,ȡ��һ�����
                            End If
                        Next
                    End If
                End If
                
                gstrSQL = gstrSQL & .�������� & "," & .�շ�ϸĿID & ",'" & .�շ���� & "','" & .���㵥λ & "',"
                gstrSQL = gstrSQL & "NULL,NULL,'" & .�շ���� & "',"
                gstrSQL = gstrSQL & IIf(.���� = 0, 1, .����) & "," & .���� & "," & _
                    IIf(.������, 8, .���ӱ�־) & "," & .ִ�в���ID & ","
            End With
            
            '������Ŀ����
            With mobjBillIncome
                intParent = intParent + 1
                gstrSQL = gstrSQL & IIf(intParent = 1, "Null", intParentNO + 1) & "," & .������ĿID & "," & _
                    "'" & .�վݷ�Ŀ & "'," & .��׼���� & "," & .Ӧ�ս�� & "," & .ʵ�ս�� & ","
                gstrSQL = gstrSQL & "NULL,"
            End With
                                            
            '��������
            '�ɿ����
            gstrSQL = gstrSQL & _
                "To_Date('" & Format(mobjBill.����ʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                "To_Date('" & Format(mobjBill.�Ǽ�ʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                "'" & mstrInNO & "'," & lng����ID & ",'" & zlStr.NeedName(cbo���㷽ʽ.Text) & "|" & mobjBill.Pages(1).Ӧ�ɽ�� & "| ',"
            'Ԥ������
            If Val(txtԤ�����.Text) <> 0 Then
                gstrSQL = gstrSQL & mobjBill.Pages(1).��Ԥ���� & ","
            Else
                gstrSQL = gstrSQL & "NULL,"
            End If
            gstrSQL = gstrSQL & "NULL,"
            gstrSQL = gstrSQL & "'" & UserInfo.��� & "','" & UserInfo.���� & "')"
            
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = gstrSQL
        Next
    Next
    
    '�޸�ǰ�˳�ԭ����
    If mstrInNO <> "" Then
        strDelBill = "zl_������շ�_DELETE('" & mstrInNO & "','" & UserInfo.��� & "','" & UserInfo.���� & "')"
    End If
    If UBound(arrSQL) >= 0 Then
        'ִ��SQL���
        On Error GoTo errH
        gcnOracle.BeginTrans
            'ɾ�����￨���۵�
            If mstrCardNO <> "" Then
                gstrSQL = "zl_���ﻮ�ۼ�¼_Delete('" & mstrCardNO & "')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            End If
    
            '�޸�ǰ���ϱ��޸ĵ���
            If strDelBill <> "" Then
                Call zlDatabase.ExecuteProcedure(strDelBill, Me.Caption)
            End If
            '�����·���
            For i = 0 To UBound(arrSQL)
                Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
            Next
            '���������
            If mobjBill.Pages(1).����� <> 0 Then
                '  Zl_���շ����_Insert(
                '  No_In         ������ü�¼.No%Type,
                '  ����id_In     ������ü�¼.����id%Type,
                '  ����id_In     ������ü�¼.����id%Type,
                '  �����_In   ������ü�¼.ʵ�ս��%Type,
                '  �Ǽ�ʱ��_In   ������ü�¼.�Ǽ�ʱ��%Type,
                '  ����Ա���_In ������ü�¼.����Ա���%Type,
                '  ����Ա����_In ������ü�¼.����Ա����%Type
                gstrSQL = "Zl_���շ����_Insert('" & mobjBill.NO & "'," & Val(mobjBill.����ID) & "," & lng����ID & "," & _
                    mobjBill.Pages(1).����� & ",To_Date('" & Format(mobjBill.����ʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                    "'" & UserInfo.��� & "','" & UserInfo.���� & "')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            End If
        gcnOracle.CommitTrans
        
        '���뵥����ʷ��¼(�������͵���)
        cboNO.AddItem mobjBill.NO, 0
        For i = cboNO.ListCount - 1 To 10 Step -1
            cboNO.RemoveItem i 'ֻ��ʾ10��
        Next
    End If
    SaveBill = True
    Exit Function
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
Private Function ReadBill(ByVal strNo As String, Optional ByVal bln���� As Boolean, _
    Optional blnDelete As Boolean, Optional blnNull As Boolean) As Boolean
'���ܣ����ݵ��ݺŶ�ȡһ�ŵ��ݲ�����������
'������blnDelete=�Ƿ��ȡҪ�˷ѵĵ���,���˷ѷ�ʽ��
'˵����Ϊͳһ���˷�ʱ������ʾ������(��Ȼ���ɲ����˷�)
    Dim rsTmp As ADODB.Recordset, rs���� As ADODB.Recordset
    Dim rsPatiMoney As ADODB.Recordset, strSQL As String
    Dim i As Long, curBillʵ�� As Currency, curBillӦ�� As Currency
    Dim blnSame As Boolean, str�ѱ� As String, intSign As Integer
    Dim str���÷ѱ� As String, blnHaveNoOne As Boolean
    
    On Error GoTo errH
    
    strNo = GetFullNO(strNo, 13)
    Call ClearRows: Call Bill.ClearBill
    
    '��ȡ��������
    strSQL = _
    " Select A.����ID,A.ʵ��Ʊ�� as Ʊ�ݺ�,A.����ID,0 as ��ҳID,A.��ʶ��,A.����,A.�Ա�,A.����,A.�ѱ�,A.���ʽ ,B.��������,B.����," & _
    "        0 as ���˲���ID,A.���˿���ID,A.��������ID,Nvl(A.�Ӱ��־,0) as �Ӱ��־,A.������,A.������,A.����ʱ��,B.ҽ�Ƹ��ʽ,A.�����־" & _
    " From " & IIf(mblnNOMoved, "H", "") & "������ü�¼ A,������Ϣ B,��Ա�� C" & _
    " Where Rownum=1 And A.����ID=B.����ID(+)" & _
    "       And A.��¼״̬" & IIf(mstrDelete <> "", "=2", IIf(bln����, "=0", " IN(1,3)")) & _
    "       And A.��¼����=1 And A.NO=[1] And Nvl(A.����Ա����,A.������)=C.����" & _
    "       And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & vbNewLine & _
            IIf(mstrDelete <> "", " And A.�Ǽ�ʱ��=[2]", "") & _
            IIf(bln����, " And A.����Ա���� is Null And ������ is Not NULL", "")
    If mstrDelete <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo, CDate(mstrDelete))
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo)
    End If
    
    If rsTmp.EOF Then
        MsgBox "û�з��ָõ���,�õ��ݿ����Ѿ����ϣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���ݺ�
    cboNO.Text = strNo
    If Trim(Nvl(rsTmp!Ʊ�ݺ�)) <> "" Then txtInvoice.Text = Nvl(rsTmp!Ʊ�ݺ�)
    
    cboNO.Tag = IIf(IsNull(rsTmp!����ID), "", rsTmp!����ID) '����ҽ�������˷�
    
    '����ID
    txtPatient.Tag = Nvl(rsTmp!����ID)
    
    '���������Ϣ��ȡ:�������ڻ��۵��շ�
    mobjBill.���� = Nvl(rsTmp!����)
    mobjBill.�Ա� = Nvl(rsTmp!�Ա�)
    mobjBill.���� = Nvl(rsTmp!����)
    mobjBill.����ID = Nvl(rsTmp!����ID, 0)
    mobjBill.��ҳID = Nvl(rsTmp!��ҳID, 0)
    mobjBill.��ʶ�� = Nvl(rsTmp!��ʶ��, 0)
    mobjBill.���� = IIf(gint������Դ = 2, "" & rsTmp!���ʽ, "") '�����ݴ渶�ʽ
    mobjBill.����ID = Nvl(rsTmp!���˲���ID, 0)
    mobjBill.����ID = Nvl(rsTmp!���˿���ID, 0)
    mobjBill.Pages(1).������ = Nvl(rsTmp!������)
    mobjBill.Pages(1).��������ID = Nvl(rsTmp!��������ID, 0)
    
    Call ReInitPatiInvoice
    '����
    If (IsNull(rsTmp!����) Or Nvl(rsTmp!����) = mstrPrePati) And chkCancel.Value = 0 Then
        If IsNull(rsTmp!����) Then
            blnNull = True
            txtPatient.Text = mstrPrePati 'ȱʡΪ��һ����������
        Else
            txtPatient.Text = rsTmp!����
            '75259:���ϴ�,2014-7-10������������ʾ��ɫ����
            Call SetPatiColor(txtPatient, Nvl(rsTmp!��������), IIf(IsNull(rsTmp!����), &HC00000, vbRed))
        End If
        blnSame = True
    Else
        '��ͬ�Ĳ���
        txtPatient.Text = Nvl(rsTmp!����)
        '75259:���ϴ�,2014-7-10������������ʾ��ɫ����
        Call SetPatiColor(txtPatient, Nvl(rsTmp!��������), IIf(IsNull(rsTmp!����), &HC00000, vbRed))
        '������Խɿ���Ϊ����,��ʹ��ͬ�Ĳ���Ҳ�����շ�
        '���Ǹպýɿ����(mstrPrePati = "")
        '���˺�:22343
        If gTy_Module_Para.byt�ɿ���� <> 1 Or mstrPrePati = "" Then
            mcurBillʵ�� = 0: mcurBillӦ�� = 0: mcurBillӦ�� = 0
            mstrPrePati = "": mintBillNO = 0: mintMoneyRow = 0
            txt�ϼ�.Text = gstrDec: txtӦ��.Text = gstrDec
            Call ClearMoney
        End If
    End If
    
    '�Ա�
    cboSex.ListIndex = cbo.FindIndex(cboSex, Nvl(rsTmp!�Ա�), True)
    If cboSex.ListIndex = -1 Then
        If Not IsNull(rsTmp!�Ա�) Then
            cboSex.AddItem rsTmp!�Ա�, 0
            cboSex.ListIndex = 0
        ElseIf cboSex.ListCount > 0 Then
            cboSex.ListIndex = 0
        End If
    End If
    
    '����
    Call LoadOldData("" & rsTmp!����, txt����, cbo���䵥λ)
    '�ѱ�
    cbo�ѱ�.ListIndex = cbo.FindIndex(cbo�ѱ�, Nvl(rsTmp!�ѱ�), True)
    If cbo�ѱ�.ListIndex = -1 And Not IsNull(rsTmp!�ѱ�) Then
        cbo�ѱ�.AddItem rsTmp!�ѱ�, 0
        cbo�ѱ�.ListIndex = 0
    End If
    
    'ҽ�Ƹ��ʽ
    If Nvl(rsTmp!�����־, 0) = 2 Or Not IsNull(rsTmp!ҽ�Ƹ��ʽ) Then
        cboҽ�Ƹ���.ListIndex = cbo.FindIndex(cboҽ�Ƹ���, rsTmp!ҽ�Ƹ��ʽ, True)
        If cboҽ�Ƹ���.ListIndex = -1 Then
            cboҽ�Ƹ���.AddItem "0-" & rsTmp!ҽ�Ƹ��ʽ, 0
            cboҽ�Ƹ���.ListIndex = 0
        End If
    Else
        cboҽ�Ƹ���.ListIndex = GetCboIndexByCode(cboҽ�Ƹ���, "" & rsTmp!���ʽ)
        If cboҽ�Ƹ���.ListIndex = -1 And Not IsNull(rsTmp!���ʽ) Then
            cboҽ�Ƹ���.AddItem rsTmp!���ʽ & "-" & GetMedPayModeName(rsTmp!���ʽ), 0
            cboҽ�Ƹ���.ListIndex = 0
        ElseIf cboҽ�Ƹ���.ListIndex = -1 Then
            cboҽ�Ƹ���.ListIndex = cbo.FindIndex(cboҽ�Ƹ���, mstr���ʽ, True)
        End If
    End If
    
    Call Set�����˿�������(mobjBill.Pages(1).������, mobjBill.Pages(1).��������ID)
    
    '���㷽ʽ:�������ҽ�����㷽ʽ
    If Not bln���� Then
        '��ȡ����ԭʼ����ʱ,��ʾ���ֽ��
        '(����)�˷�ʱ,��ʾԭʼ���ݵĽ�����
        intSign = IIf(mstrDelete <> "", -1, 1) '����,�����������
        strSQL = _
            " Select 1 As ��ʽ, ���㷽ʽ, Sum(1 * ��Ԥ��) As ���" & _
            " From ����Ԥ����¼ A, ���㷽ʽ B" & _
            " Where a.���㷽ʽ = b.���� And b.���� <> 9 And ��¼���� = 3 And ����id = [1]" & _
            " Group By ���㷽ʽ" & _
            " Having Nvl(Sum(1 * ��Ԥ��), 0) <> 0"
        'Ԥ����
        strSQL = strSQL & _
            " Union All" & _
            " Select 2 As ��ʽ, Null, Sum(1 * ��Ԥ��) As ���" & _
            " From ����Ԥ����¼" & _
            " Where ��¼���� In (1, 11) And ����id = [1] Having Nvl(Sum(1 * ��Ԥ��), 0) <> 0"
        '����
        strSQL = strSQL & _
            " Union All" & _
            " Select 3 As ��ʽ, ���㷽ʽ, Sum(1 * ��Ԥ��) As ���" & _
            " From ����Ԥ����¼ A, ���㷽ʽ B" & _
            " Where a.���㷽ʽ = b.���� And b.���� = 9 And ��¼���� = 3 And ����id = [1]" & _
            " Group By ���㷽ʽ" & _
            " Having Nvl(Sum(1 * ��Ԥ��), 0) <> 0"

            
        Set rs���� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CLng(rsTmp!����ID))
        
        For i = 1 To rs����.RecordCount
            If rs����!��ʽ = 2 Then
                txtԤ�����.Text = Format(rs����!���, "0.00")
            ElseIf rs����!��ʽ = 3 Then
                pic���.Visible = True: lbl����.Caption = Format(rs����!���, "0.00")
            Else
                cbo���㷽ʽ.ListIndex = cbo.FindIndex(cbo���㷽ʽ, rs����!���㷽ʽ, True)
                If cbo���㷽ʽ.ListIndex = -1 Then
                    cbo���㷽ʽ.AddItem rs����!���㷽ʽ, 0
                    cbo���㷽ʽ.ListIndex = 0
                End If
            End If
            rs����.MoveNext
        Next
    End If
    
    '�Ӱ�״̬
    chk�Ӱ�.Value = IIf(IsNull(rsTmp!�Ӱ��־), 0, rsTmp!�Ӱ��־)
    
    '����ʱ��
    txtDate.Text = Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm:ss")
    
    '��ȡ�����շ�ϸĿ����
    '---------------------------------------------------------------------------------------------
    If blnDelete Then
        '��ȡ׼����,������Ӧ�ս��,ʵ�ս��(���=ʣ����*(׼����/ʣ����))
        '���ŵ��ݻ��ܽ��(��ϸ���շ�ϸĿ)
        'ִ��״̬Ӧ����ԭʼ��¼���ж�(������ҩ�Ҳ����˷ѵļ�¼)
        strSQL = "" & _
        " Select Nvl(�۸񸸺�,���) " & _
        " From " & IIf(mblnNOMoved, "H", "") & "������ü�¼" & _
        " Where ��¼����=1 And ��¼״̬ IN(0,1,3) And NO=[1] And Nvl(ִ��״̬,0)<>1 And Nvl(���ӱ�־,0)<>9"
        
        strSQL = _
        " Select A.��¼״̬,Nvl(A.�۸񸸺�,A.���) as ���," & _
        "        A.�ѱ�,C.����,C.���� as ���,A.�շ�ϸĿID,B.����,B.���,Nvl(A.��������,B.��������) ��������,A.���㵥λ," & _
        "        Avg(Nvl(A.����,1)*A.����) as ����,Sum(A.��׼����) as ����," & _
        "        Sum(A.Ӧ�ս��) as Ӧ�ս��,Sum(A.ʵ�ս��) as ʵ�ս��, " & _
        "        A.ִ�в���ID,D.���� as ִ�в���,A.���ӱ�־" & _
        " From " & IIf(mblnNOMoved, "H", "") & "������ü�¼ A,�շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D " & _
        " Where A.�շ�ϸĿID=B.ID And C.����=A.�շ���� And A.ִ�в���ID=D.ID " & _
        "       And A.��¼����=1 And A.NO=[1] And Nvl(A.�۸񸸺�,A.���) IN(" & strSQL & ")" & _
        "       And Nvl(A.���ӱ�־,0)<>9" & _
        " Group by A.��¼״̬,Nvl(A.�۸񸸺�,A.���),A.�ѱ�,C.����,C.����,A.�շ�ϸĿID,B.����," & _
        "          B.���,Nvl(A.��������,B.��������),A.���㵥λ,A.ִ�в���ID,D.����,A.���ӱ�־"
            
        '��������(ʣ��������Ϊ׼������,���ؼ���)
        '�ſ��Ѿ�ȫ���˷ѵ���(ִ��״̬=0��һ�ֿ���)
        strSQL = _
        " Select A.���,A.�ѱ�,A.����,A.���,A.�շ�ϸĿID,A.����,A.���," & _
        "        A.��������,A.���㵥λ,A.ִ�в���ID,A.ִ�в���,A.���ӱ�־," & _
        "        Sum(A.����) as ����,A.����,Sum(A.Ӧ�ս��) as Ӧ�ս��,Sum(A.ʵ�ս��) as ʵ�ս��" & _
        " From (" & strSQL & ") A" & _
        " Group by A.���,A.�ѱ�,A.����,A.���,A.�շ�ϸĿID,A.����,A.���,A.��������," & _
        "          A.���㵥λ,A.����,A.ִ�в���ID,A.ִ�в���,A.���ӱ�־" & _
        " Having Sum(A.����)<>0" & _
        " Order by A.���"
    Else
        '��ȡ����ԭʼ����
        intSign = IIf(mstrDelete <> "", -1, 1) '����,�����������
        strSQL = _
        " Select Nvl(A.�۸񸸺�,A.���) as ���," & _
        "        A.�ѱ�,C.����,C.���� as ���,A.�շ�ϸĿID,B.����,B.���,Nvl(A.��������,B.��������) ��������,A.���㵥λ," & _
        "        Avg(" & intSign & "*Nvl(A.����,1)*A.����) as ����," & _
        "        Sum(A.��׼����) as ����,Sum(" & intSign & "*A.Ӧ�ս��) as Ӧ�ս��, " & _
        "        Sum(" & intSign & "*A.ʵ�ս��) as ʵ�ս��, " & _
        "        A.ִ�в���ID,D.���� as ִ�в���,A.���ӱ�־" & _
        " From " & IIf(mblnNOMoved, "H", "") & "������ü�¼ A,�շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D " & _
        " Where A.�շ�ϸĿID=B.ID And C.����=A.�շ���� And A.ִ�в���ID=D.ID " & _
        "       And A.��¼����=1 And A.NO=[1]" & _
        "       And A.��¼״̬" & IIf(mstrDelete <> "", "=2", IIf(bln����, "=0", " IN(1,3)")) & _
                IIf(mstrDelete <> "", " And A.�Ǽ�ʱ��=[2]", "") & _
                IIf(Not gblnShowErr, " And Nvl(A.���ӱ�־,0)<>9", "") & _
        " Group by Nvl(A.�۸񸸺�,A.���),A.�ѱ�,C.����,C.����,A.�շ�ϸĿID,B.����," & _
        "           B.���,Nvl(A.��������,B.��������),A.���㵥λ,A.ִ�в���ID,D.����,A.���ӱ�־" & _
        " Order by ���"
    End If
    
    If mstrDelete <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo, CDate(mstrDelete))
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo)
    End If
    
    If rsTmp.EOF Then Exit Function
    
    Bill.Redraw = False
    Bill.Rows = rsTmp.RecordCount + 1
    str���÷ѱ� = "": blnHaveNoOne = False
    For i = 1 To rsTmp.RecordCount
        '�ѱ�
        If Not IsNull(rsTmp!�ѱ�) Then
            If InStr(str�ѱ� & ",", "," & rsTmp!�ѱ� & ",") = 0 Then
                str�ѱ� = str�ѱ� & "," & rsTmp!�ѱ�
            End If
        End If
    
        Bill.TextMatrix(i, 0) = rsTmp!����
        Bill.TextMatrix(i, 1) = Format(rsTmp!Ӧ�ս��, gstrDec)
        Bill.TextMatrix(i, 2) = Format(rsTmp!ʵ�ս��, gstrDec)
        Bill.TextMatrix(i, 3) = rsTmp!ִ�в���
        Bill.TextMatrix(i, 4) = IIf(IsNull(rsTmp!��������), "", rsTmp!��������)
        
        If Nvl(rsTmp!���) <> "����" Then
            If InStr(1, "," & str���÷ѱ� & ",", "," & Nvl(rsTmp!���) & ",") = 0 Then
                str���÷ѱ� = str���÷ѱ� & "," & Nvl(rsTmp!���)
            End If
        End If
        If Val(Nvl(rsTmp!����)) <> 1 Then blnHaveNoOne = True
        rsTmp.MoveNext
    Next
    
    If str���÷ѱ� <> "" Then
        str���÷ѱ� = Mid(str���÷ѱ�, 2)
        str���÷ѱ� = Replace(str���÷ѱ�, ",", "��")
        MsgBox "���� [" & strNo & "] �д������������շ���Ŀ�����ܽ��м��շѣ�" & vbCrLf & vbCrLf & _
            "        " & str���÷ѱ�, vbInformation, gstrSysName
        Exit Function
    ElseIf blnHaveNoOne Then
        MsgBox "���� [" & strNo & "] �д���������Ϊ1���շ���Ŀ�����ܽ��м��շѣ�", vbInformation, gstrSysName
        Exit Function
    End If

    
    '�ѱ�
    lbl��̬�ѱ�.Caption = Mid(str�ѱ�, 2)
    
    '����б༭����������ɫ
    Bill.SetColColor 0, &HE7CFBA
    Bill.SetColColor 1, &HE7CFBA
    Bill.SetColColor 3, &HE7CFBA
    Bill.Redraw = True
    
    '��ȡ�����վݷ�Ŀ����
    If blnDelete Then
        '��ȡ׼����,������Ӧ�ս��,ʵ�ս��(���=ʣ����*(׼����/ʣ����))
        '���ŷ��õ���(��ϸ��������Ŀ)
        'ִ��״̬Ӧ����ԭʼ��¼���ж�(������ҩ�Ҳ����˷ѵļ�¼)
        strSQL = "" & _
        " Select Nvl(�۸񸸺�,���)  " & _
        " From " & IIf(mblnNOMoved, "H", "") & "������ü�¼" & _
        " Where ��¼����=1 And ��¼״̬ IN(0,1,3) And NO=[1] And Nvl(ִ��״̬,0)<>1 And Nvl(���ӱ�־,0)<>9"
        
        strSQL = _
        " Select A.���,A.����," & _
        "       Sum(A.����) as ʣ������,Sum(A.Ӧ�ս��) as ʣ��Ӧ��," & _
        "       Sum(A.ʵ�ս��) as ʣ��ʵ�� " & _
        " From ( Select A.��¼״̬,A.���," & IIf(gint����ϼ� = 0, "A.�վݷ�Ŀ", "B.����") & " as ����," & _
        "               Nvl(A.����,1)*A.���� as ����,A.Ӧ�ս��,A.ʵ�ս��" & _
        "        From " & IIf(mblnNOMoved, "H", "") & "������ü�¼ A,������Ŀ B" & _
        "       Where A.��¼����=1 And A.NO=[1] And Nvl(A.���ӱ�־,0)<>9" & _
        "            And A.������ĿID=B.ID And Nvl(A.�۸񸸺�,A.���) IN(" & strSQL & ")" & _
        "        ) A" & _
        " Group by A.���,A.����" & _
        " Having Sum(����)<>0"
                    
        '��������(׼��������ʣ������,������������)
        strSQL = _
        " Select A.����,Sum(A.ʣ��Ӧ��) as Ӧ�ս��," & _
        "       Sum(A.ʣ��ʵ��) as ʵ�ս��" & _
        " From (" & strSQL & ") A" & _
        " Group by A.����"
    Else
        '��ȡ����ԭʼ����
        intSign = IIf(mstrDelete <> "", -1, 1) '����,�����������
        strSQL = _
        " Select " & IIf(gint����ϼ� = 0, "A.�վݷ�Ŀ", "B.����") & " as ����," & _
        "        Sum(" & intSign & "*A.Ӧ�ս��) as Ӧ�ս��," & _
        "        Sum(" & intSign & "*A.ʵ�ս��) as ʵ�ս�� " & _
        " From " & IIf(mblnNOMoved, "H", "") & "������ü�¼ A,������Ŀ B" & _
        " Where A.������ĿID=B.ID And A.��¼״̬" & IIf(mstrDelete <> "", "=2", IIf(bln����, "=0", " IN(1,3)")) & _
        "       AND A.��¼����=1 And A.NO=[1]" & _
                IIf(mstrDelete <> "", " And A.�Ǽ�ʱ��=[2]", "") & _
                IIf(Not gblnShowErr, " And Nvl(A.���ӱ�־,0)<>9", "") & _
        " Group By " & IIf(gint����ϼ� = 0, "A.�վݷ�Ŀ", "B.����")
    End If
    
    If mstrDelete <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo, CDate(mstrDelete))
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo)
    End If
    If rsTmp.EOF Then Exit Function
    
    'ˢ����ʾ(�շ�Ҫ����)
    mshMoney.Rows = rsTmp.RecordCount + 1 + mintMoneyRow
    If mshMoney.Rows < 5 Then mshMoney.Rows = 5

    Call SetMoneyList
    
    For i = mintMoneyRow + 1 To mshMoney.Rows - 1
        mshMoney.TextMatrix(i, 0) = ""
        mshMoney.TextMatrix(i, 1) = ""
        mshMoney.TextMatrix(i, 2) = ""
    Next
    For i = mintMoneyRow + 1 To rsTmp.RecordCount + mintMoneyRow
        mshMoney.TextMatrix(i, 0) = mintBillNO + 1
        mshMoney.TextMatrix(i, 1) = rsTmp!����
        mshMoney.TextMatrix(i, 2) = Format(rsTmp!ʵ�ս��, gstrDec)
        curBillӦ�� = curBillӦ�� + rsTmp!Ӧ�ս��
        curBillʵ�� = curBillʵ�� + rsTmp!ʵ�ս��
        rsTmp.MoveNext
    Next
    
    '���ܴ���
    With mobjBill.Pages(1)
        .NO = strNo
        .Ӧ�ս�� = curBillӦ��
        .ʵ�ս�� = curBillʵ��
        '�շ�ʱ��ȡ���۵�ʱ
        If bln���� Then Call ShowPrice
    End With
    
    txtӦ��.Text = Format(mcurBillӦ�� + curBillӦ��, gstrDec)
    txt�ϼ�.Text = Format(mcurBillʵ�� + curBillʵ��, gstrDec)
    
    '�շ���ʾ�˿�ϼ�
    If blnDelete Then
        lblӦ��.Caption = "Ӧ�˽��"
        lblӦ��.ForeColor = vbRed
        txtӦ��.ForeColor = vbRed
        txtӦ��.Text = Format(GetDelMoney, "0.00")
    End If
    
    'ˢ���շ��ۼ�
    If chkCancel.Value = 0 And gbln�ۼ� Then
        txt�ۼ�.Text = Format(GetChargeTotal, "0.00")
    End If
    
    On Error Resume Next
    For i = 1 To mshMoney.Rows - 1
        If mshMoney.TextMatrix(i, 0) = mintBillNO + 1 Then
            mshMoney.TopRow = i
        End If
    Next
    ReadBill = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
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

Private Function GetBillCount() As Integer
'���ܣ����㵱ǰ�շ���Ҫ��ӡ������Ʊ��
    Dim strItems As String
    Dim i As Integer, j As Integer
    
    If gTy_Module_Para.blnһ��Ʊ�� Then GetBillCount = 1: Exit Function
    
    '���ŵ��ݰ���Ŀ���ܼ���
    For i = 1 To mobjBill.Pages(1).Details.Count
        If Not mobjBill.Pages(1).Details(i).������ Then '�ſ�������
            For j = 1 To mobjBill.Pages(1).Details(i).InComes.Count
                If mobjBill.Pages(1).Details(i).InComes(j).ʵ�ս�� <> 0 Then '��Ϊ��
                    If InStr(strItems & ",", "," & mobjBill.Pages(1).Details(i).InComes(j).�վݷ�Ŀ & ",") = 0 Then
                        strItems = strItems & "," & mobjBill.Pages(1).Details(i).InComes(j).�վݷ�Ŀ
                    End If
                End If
            Next
        End If
    Next
    GetBillCount = IntEx((UBound(Split(Mid(strItems, 2), ",")) + 1) / gTy_Module_Para.byt�����վ��д�)
End Function

Private Sub DelFactMoney()
'���ܣ�ɾ�������еĹ�������(������Ҫ������ʱ)
    Dim i As Long
    
    '���ж��Ƿ��Ѿ������˹�����
    For i = 1 To mobjBill.Pages(1).Details.Count
        If mobjBill.Pages(1).Details(i).������ Then
            Call DeleteDetail(i)
            Call ShowMoney
            If mobjBill.Pages(1).Details.Count = 0 Then ClearMoney
            Exit Sub
        End If
    Next
End Sub

Private Sub SetFactMoney()
'���ܣ��շ�ʱ���á���ʾ�����㹤����
'˵�����������Զ����ڵ�ǰ��ʾ�ĵ�����
    Dim objDetail As Detail
    Dim colIncomes As New BillInComes
    Dim blnExist As Boolean, i As Integer
    Dim lngRow As Long, lngDoUnit As Long
    Dim int���� As Integer
    
    int���� = GetBillCount
    If int���� = 0 Then Call DelFactMoney: Exit Sub 'ɾ��������
    
    '���ж��Ƿ��Ѿ������˹�����
    For i = 1 To mobjBill.Pages(1).Details.Count
        If mobjBill.Pages(1).Details(i).������ Then
            lngRow = i: blnExist = True: Exit For
        End If
    Next

    If Not blnExist Then
        Set objDetail = Get������
        If objDetail Is Nothing Then Exit Sub '�Ҳ���������,������
        
        If mobjBill.Pages(1).Details.Count >= Bill.Rows - 1 Then
            Bill.Rows = Bill.Rows + 1
        Else
            For i = 1 To Bill.COLS - 1
                Bill.TextMatrix(Bill.Rows - 1, i) = ""
            Next
        End If
        lngRow = mobjBill.Pages(1).Details.Count + 1
        
        lngDoUnit = mobjBill.����ID '���˿���
        If lngDoUnit = 0 Then lngDoUnit = Get��������ID
        lngDoUnit = Get�շ�ִ�п���ID(objDetail.���, objDetail.ID, objDetail.ִ�п���, lngDoUnit, Get��������ID, gint������Դ, , , , , mobjBill.����ID)
        With objDetail
            mobjBill.Pages(1).Details.Add "", objDetail, .ID, CInt(lngRow), 0, .���, .���㵥λ, .���, 1, 1, 0, lngDoUnit, colIncomes
        End With
        mobjBill.Pages(1).Details(lngRow).������ = True
    End If
    
    '���¸��ݵ�ǰ�����������ù���������(�̶�Ϊ1)
    mobjBill.Pages(1).Details(lngRow).���� = int����
    Call CalcMoney(lngRow)
    
    Call ShowDetails(lngRow)
    Call ShowMoney
End Sub

Private Sub ClearRows()
    Dim i As Integer
    For i = 1 To Bill.Rows - 1
        Bill.RowData(i) = 0
    Next
End Sub

Private Sub FillBillComboBox(lngRow As Long, lngCol As Long)
'���ܣ����ݵ��������������б������
    Dim rsTmp As ADODB.Recordset
    Dim str��Ա���� As String, strTmp As String
    Dim strSQL As String, i As Long
    Dim lng����ID As Long, lng����ID As Long
    Dim rsUnit As ADODB.Recordset
    On Error GoTo errHandle
    

    Bill.Clear
    
    Select Case Bill.TextMatrix(0, lngCol)
        Case "ִ�п���"
            '���ݵ�ǰ��Ŀִ�п�������,��̬���ÿ�ѡ����
            If mobjBill.Pages(1).Details.Count >= lngRow Then
                With mobjBill.Pages(1).Details(lngRow)
                    Bill.TextMatrix(lngRow, lngCol) = ""
                    
                    lng����ID = mobjBill.����ID
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
                            strTmp = rsUnit!���� & "-" & rsUnit!����
                            If Not (SendMessage(Bill.cboHwnd, CB_FINDSTRING, -1, ByVal strTmp) >= 0) Then
                                Bill.AddItem strTmp
                                Bill.ItemData(Bill.ListCount - 1) = rsUnit!ID
                                                                
                                '����ȱʡִ�п���
                                If lngRow = 1 Then
                                    If rsUnit!ID = lng����ID Then Bill.ListIndex = Bill.NewIndex
                                ElseIf lngRow > 1 Then
                                    '����һ�з�ҩƷ��ͬ
                                    If rsUnit!ID = mobjBill.Pages(1).Details(lngRow - 1).ִ�в���ID And _
                                        mobjBill.Pages(1).Details(lngRow - 1).Detail.ִ�п��� = .Detail.ִ�п��� Then
                                        Bill.ListIndex = Bill.NewIndex
                                    ElseIf rsUnit!ID = lng����ID And Bill.ListIndex = -1 Then
                                        Bill.ListIndex = Bill.NewIndex
                                    End If
                                End If
                            End If
                            rsUnit.MoveNext
                        Next
                    End If
                        
                    If .Detail.ִ�п��� = 4 Then     'ִ�п���Ϊָ�����ҵ�,ȱʡΪ����Ա���ڿ���
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
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub FillDoctor(Optional lng����ID As Long)
'���ܣ�����ָ���Ŀ�������ID��ȡ����дҽ���б�,����ȱʡҽ��
    Dim lngOldID As Long
    
    cbo������.Clear
    Call GetDoctor(lng����ID, mrs������)
    
    Do While Not mrs������.EOF
        If lngOldID <> mrs������!ID Then
            If gbyt��������ʾ = 1 Then
                cbo������.AddItem mrs������!���� & "-" & mrs������!����
            Else
                cbo������.AddItem mrs������!��� & "-" & mrs������!����
            End If
            cbo������.ItemData(cbo������.NewIndex) = mrs������!ID
            lngOldID = mrs������!ID
        End If
        mrs������.MoveNext
    Loop
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

Private Sub txtRePrint_GotFocus()
    zlControl.TxtSelAll txtRePrint
End Sub

Private Sub txtRePrint_KeyPress(KeyAscii As Integer)
    Dim strNo As String, strOper As String, vDate As Date
    Dim strReclaimIvoice As String  '����Ʊ��
    
    If KeyAscii > 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))

    '��һλ����������ĸ,����λ����
    If KeyAscii <> 13 Then
        Call SetNOInputLimit(txtRePrint, KeyAscii)
    Else
        '�ش�
        txtRePrint.Text = GetFullNO(txtRePrint.Text, 13)
        zlControl.TxtSelAll txtRePrint
       
        '�Ƿ���ת������ݱ���
        If zlDatabase.NOMoved("������ü�¼", txtRePrint.Text, , "1") Then
            If Not ReturnMovedExes(txtRePrint.Text, 1, Me.Caption) Then Exit Sub
            mblnNOMoved = False
        End If
        
        If Not ReadBillInfo(1, txtRePrint.Text, 1, strOper, vDate) Then
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
        
        
        strNo = "'" & txtRePrint.Text & "'"
        '������ʣ�������Ĳſ����ش�
        If Not BillExistMoney(strNo, 1, True) Then
            MsgBox "���ݲ����ڻ��Ѿ�ȫ���˷�,�����ش�", vbInformation, gstrSysName
            txtRePrint.Text = "": Exit Sub
        End If
        
        '56963
        strReclaimIvoice = zlGetReclaimInvoice(strNo)
        If strReclaimIvoice <> "" Then
            Call MsgBox("ע��:" & vbCrLf & " ��ע��������·�Ʊ:" & vbCrLf & strReclaimIvoice, vbOKOnly + vbInformation, gstrSysName)
        End If
        Dim intInvoiceFormat As Integer
        intInvoiceFormat = IIf(strReclaimIvoice = "" And gTy_Module_Para.bytƱ�ݷ������ <> 0, mintOldInvoiceFormat, mintInvoiceFormat)
        
        Dim strPriceGrade As String
        If gintPriceGradeStartType >= 2 Then
            strPriceGrade = GetPriceGradeFromNos(strNo)
        Else
            strPriceGrade = mstr��ͨ�۸�ȼ�
        End If
        If Not RePrintCharge(1, strNo, Me, mlng����ID, strReclaimIvoice, , , intInvoiceFormat, , , mlngShareUseID, _
            mstrUseType, , strPriceGrade) Then
            txtRePrint.SetFocus
        Else
            Call RefreshFact
            txtRePrint.Text = ""
            txtPatient.SetFocus
        End If
    End If
End Sub

Private Sub RefreshFact()
'���ܣ�ˢ���շ�Ʊ�ݺ�
    If gblnStrictCtrl Then
        If zlGetInvoiceGroupUseID(mlng����ID) = False Then
            txtInvoice.Text = "": Exit Sub
        End If
        '�ϸ�ȡ��һ������
        txtInvoice.Text = GetNextBill(mlng����ID)
    Else
        '��ɢ��ȡ��һ������
        txtInvoice.Text = zlStr.Increase(UCase(zlDatabase.GetPara("��ǰ�շ�Ʊ�ݺ�", glngSys, mlngModul)))
    End If
End Sub

Private Function CalcBillToTal(Optional blnӦ�� As Boolean) As Currency
    Dim objTmpDetail As New BillDetail
    Dim objTmpIncome As New BillInCome
    Dim i As Integer, intCol As Integer

    If mobjBill.Pages(1).Details.Count > 0 Then
        For Each objTmpDetail In mobjBill.Pages(1).Details
            For Each objTmpIncome In objTmpDetail.InComes
                If blnӦ�� Then
                    CalcBillToTal = CalcBillToTal + objTmpIncome.Ӧ�ս��
                Else
                    CalcBillToTal = CalcBillToTal + objTmpIncome.ʵ�ս��
                End If
            Next
        Next
    Else
        For i = 0 To Bill.COLS - 1
            If blnӦ�� Then
                If Bill.TextMatrix(0, i) = "Ӧ�ս��" Then intCol = i: Exit For
            Else
                If Bill.TextMatrix(0, i) = "ʵ�ս��" Then intCol = i: Exit For
            End If
        Next
    
        For i = 1 To Bill.Rows - 1
            CalcBillToTal = CalcBillToTal + Val(Bill.TextMatrix(i, intCol))
        Next
    End If
    CalcBillToTal = Format(CalcBillToTal, gstrDec)
End Function

Private Function Calc������() As Currency
    Dim objTmpDetail As New BillDetail
    Dim objTmpIncome As New BillInCome

    For Each objTmpDetail In mobjBill.Pages(1).Details
        If objTmpDetail.������ Then
            For Each objTmpIncome In objTmpDetail.InComes
                Calc������ = Calc������ + objTmpIncome.ʵ�ս��
            Next
        End If
    Next
End Function

Private Sub txt�ɿ�_LostFocus()
    mblnHotKey = False
    If Val(txt�ɿ�.Text) = 0 Then txt�ɿ�.Text = "0.00"
End Sub

Private Function SaveModi() As Boolean
'���ܣ����浱ǰ�޸ĵķ��õ���
    Dim strSQL As String
    
    strSQL = "zl_���˷��ü�¼_Update('" & cboNO.Text & "',1," & _
        "'" & zlStr.NeedName(cbo������.Text) & "',To_Date('" & txtDate.Text & "','YYYY-MM-DD HH24:MI:SS'))"
    On Error GoTo errH
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    SaveModi = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub FillDept(Optional lng��ԱID As Long)
'���ܣ���ȡ�����ؿ����б�,����ȱʡ����
'������lng��ԱID=ֻ��ȡָ����Ա���ڿ���(������ȱʡ��)
'���أ����Ҹ���
    
    Dim strSQL As String, i As Long, lngOldDepID As Long
    Dim strDepts As String  'ָ����Ա�����Ķ������
        
    cbo��������.Clear
    If mrs�������� Is Nothing Then Call GetDoctorDept(mrs��������)
   
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
            If lngOldDepID <> mrs��������!ID Then   'һ�����ſ���ͬʱ�����������ٴ�,��������ͬ��
                cbo��������.AddItem mrs��������!���� & "-" & mrs��������!����
                cbo��������.ItemData(cbo��������.NewIndex) = mrs��������!ID
                lngOldDepID = mrs��������!ID
            End If
            mrs��������.MoveNext
        Next
    End If
End Sub

Private Function Check��������(Optional intRow As Integer) As Boolean
'���ܣ����ݵ�ǰ���˵������ж�ָ���е���Ŀ�Ƿ��������,����������������Ŀ
    Dim strSQL As String
    Dim i As Integer, strType As String
    Dim rsTmp As New ADODB.Recordset
    
    Check�������� = True
    
    On Error GoTo errHandle
    
    '�޷����
    If cboҽ�Ƹ���.ListIndex = -1 Then Exit Function
    
    'ȷ����������
    strType = cboҽ�Ƹ���.Text
    
    'ֻ���ҽ�����˺͹��Ѳ���
    If strType <> "1" And strType <> "2" Then Exit Function
    
    '��ȡ�������
    If strType = "1" Then
        strSQL = "Select ����,����,����,����,ȱʡ��־ From �������� Where ���� In(" & gstrҽ���������� & ") Order by ����"
    Else
        strSQL = "Select ����,����,����,����,ȱʡ��־ From �������� Where ���� In(" & gstr���ѷ������� & ") Order by ����"
    End If
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    If rsTmp.EOF Then Exit Function
    
    If intRow > 0 Then
        If mobjBill.Pages(1).Details(intRow).Detail.���� = "" Then
            MsgBox """" & mobjBill.Pages(1).Details(intRow).Detail.���� & """�ķ�������δ���ã�", vbInformation, gstrSysName
            Check�������� = False
        Else
            rsTmp.Filter = "����='" & mobjBill.Pages(1).Details(intRow).Detail.���� & "'"
            If rsTmp.EOF Then
                MsgBox """" & mobjBill.Pages(1).Details(intRow).Detail.���� & """������Ϊ""" & _
                    mobjBill.Pages(1).Details(intRow).Detail.���� & """,����" & _
                    IIf(strType = "1", "ҽ��", "����") & "�������ͣ�", vbInformation, gstrSysName
                Check�������� = False
            End If
        End If
    Else
        For i = 1 To mobjBill.Pages(1).Details.Count
            If mobjBill.Pages(1).Details(i).Detail.���� = "" Then
                If MsgBox("�����е� " & i & " ����Ŀ""" & mobjBill.Pages(1).Details(i).Detail.���� & """�ķ�������δ���ã�" & vbCrLf & "ȷʵҪ���浥����", _
                    vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Check�������� = False: Exit For
                End If
            Else
                rsTmp.Filter = "����='" & mobjBill.Pages(1).Details(i).Detail.���� & "'"
                If rsTmp.EOF Then
                    If MsgBox("�����е� " & i & " ����Ŀ""" & mobjBill.Pages(1).Details(i).Detail.���� & """�ķ�������Ϊ""" & _
                        mobjBill.Pages(1).Details(i).Detail.���� & """,����" & _
                        IIf(strType = "1", "ҽ��", "����") & "�������ͣ�" & vbCrLf & "ȷʵҪ���浥����", _
                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Check�������� = False: Exit For
                    End If
                End If
            End If
        Next
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub txt�ɿ�_Validate(Cancel As Boolean)
    If Val(txt�ɿ�.Text) = 0 Then txt�ɿ�.Text = "0.00"
End Sub

Private Sub LoadAndSeek�ѱ�()
    Dim lngDeptID As Long
     
    '�ѱ���
    If cbo��������.ListIndex <> -1 Then lngDeptID = cbo��������.ItemData(cbo��������.ListIndex)
    
    '��ȡΨһ�Էѱ�(�������޳���,�Ա㶨λ)����̬�ѱ�
    Call Load�ѱ�(cbo�ѱ�, lngDeptID, True, mrs�ѱ�)
    
    '��ȡ��̬�ѱ�,Ĭ��Ϊ�ɼ�
    If lbl��̬�ѱ�.Visible Then     '����Ĭ��ΪTrue
        lbl��̬�ѱ�.Caption = Load��̬�ѱ�(lngDeptID)
        lbl��̬�ѱ�.Tag = lbl��̬�ѱ�.Caption
        lbl��̬�ѱ�.Visible = lbl��̬�ѱ�.Caption <> ""
        If lbl��̬�ѱ�.Caption <> "" Then lbl��̬�ѱ�.Caption = "(" & lbl��̬�ѱ�.Caption & ")"
    End If
    
    If mrsInfo.State = 0 Then
        '�����������˿�������ѡ��
        cbo�ѱ�.Locked = False
        If cbo�ѱ�.ListIndex = -1 And cbo�ѱ�.ListCount > 0 Then cbo�ѱ�.ListIndex = 0
    ElseIf mrsInfo.State = 1 Then
        '��λ�е������˵ķѱ�
        cbo�ѱ�.ListIndex = cbo.FindIndex(cbo�ѱ�, Nvl(mrsInfo!�ѱ�), True)
        If cbo�ѱ�.ListIndex <> -1 Then
            '���жϳ����Ƿ����
            If cbo�ѱ�.ItemData(cbo�ѱ�.ListIndex) = 2 And mrsInfo!���� = 0 Then
                'ʹ��ȱʡ�ѱ�(�������޳���ѱ�)
                Call Load�ѱ�(cbo�ѱ�, lngDeptID, False, mrs�ѱ�)
                If cbo�ѱ�.ListIndex <> -1 Then
                    If Visible Then MsgBox "����ʹ�ý��޳���ķѱ�:" & mrsInfo!�ѱ� & ",�����˲��ǵ�һ�ξ���,��ʹ��ȱʡ�ѱ�", vbInformation, gstrSysName
                Else
                    cbo�ѱ�.Locked = False '�޷�ȷ��,��������ѡ��
                    If Visible Then MsgBox "����ʹ�ý��޳���ķѱ�:" & mrsInfo!�ѱ� & ",�����˲��ǵ�һ�ξ���,��ѡ��һ�ַѱ�", vbInformation, gstrSysName
                    If cbo�ѱ�.Enabled And cbo�ѱ�.Visible Then cbo�ѱ�.SetFocus
                End If
            Else
                'cbo�ѱ�.Locked = True '��λ���˶�Ӧ�ѱ�,�����޸�
            End If
        Else
            'ʹ��ȱʡ�ѱ�(�������޳���ѱ�)
            Call Load�ѱ�(cbo�ѱ�, lngDeptID, False, mrs�ѱ�)
            If cbo�ѱ�.ListIndex <> -1 Then
            
                If Visible Then MsgBox "û�з��ֲ��˷ѱ�:" & mrsInfo!�ѱ� & ",��ʹ��ȱʡ�ѱ�", vbInformation, gstrSysName
            Else
                cbo�ѱ�.Locked = False '�޷�ȷ��,��������ѡ��
                If Visible Then MsgBox "û�з��ֲ��˷ѱ�:" & mrsInfo!�ѱ� & "��ȱʡ�ѱ�,��ѡ��һ�ַѱ�", vbInformation, gstrSysName
                If cbo�ѱ�.Enabled And cbo�ѱ�.Visible Then cbo�ѱ�.SetFocus
            End If
        End If
    End If
End Sub

Private Function ItemExist(lng�շ�ϸĿID As Long) As Boolean
    Dim i As Long
    
    If mobjBill Is Nothing Then Exit Function
    
    For i = 1 To mobjBill.Pages(1).Details.Count
        If mobjBill.Pages(1).Details(i).�շ�ϸĿID = lng�շ�ϸĿID Then
            ItemExist = True: Exit Function
        End If
    Next
End Function

Private Function Checkִ�п���() As Integer
    Dim i As Integer
    For i = 1 To mobjBill.Pages(1).Details.Count
        If mobjBill.Pages(1).Details(i).ִ�в���ID = 0 Or Bill.TextMatrix(i, 3) = "" Then
            Checkִ�п��� = i: Exit Function
        End If
    Next
End Function
Private Sub ReInitPatiInvoice(Optional blnFact As Boolean = True)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���³�ʼ�����˷�Ʊ��Ϣ
    '����:���˺�
    '����:2011-04-29 14:17:33
    '����:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInvoiceFormat As String
    mstrUseType = "": mlngShareUseID = 0: mintInvoiceFormat = 0
    mstrUseType = zl_GetInvoiceUserType(mobjBill.����ID, 0, 0)
    mlngShareUseID = zl_GetInvoiceShareID(mlngModul, mstrUseType)
    mintInvoiceFormat = zl_GetInvoicePrintFormat(mlngModul, mstrUseType, mintOldInvoiceFormat)
    mintInvoicePrint = zl_GetInvoicePrintMode(mlngModul, mstrUseType)
    If blnFact Then RefreshFact
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

 
 
Private Sub initCardSquareData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������رս��㿨����
    '���:blnClosed:�رն���
    '����:���˺�
    '����:2010-01-05 14:51:23
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    If mbytInState = 1 Then Exit Sub
    If gobjSquare.objSquareCard Is Nothing Then Exit Sub
    
    Dim objCard As Card
    Call IDKind.zlInit(Me, glngSys, mlngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", txtPatient)
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

Private Function IsCheck����() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������Ƿ���������
    '����:��������,����true,���򷵻�False
    '����:���˺�
    '����:2014-06-05 15:17:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If gstr�������� <> "" Then IsCheck���� = True: Exit Function
    Select Case mbytInState
        Case 0
            MsgBox "ϵͳ����δ������Ч������,����[���㷽ʽ����]�����á�", vbInformation, gstrSysName
            Exit Function
        Case Else
            IsCheck���� = True
    End Select
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function
