VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#9.0#0"; "zlIDKind.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReplenishTheBalanceDel 
   AutoRedraw      =   -1  'True
   Caption         =   "���ղ������"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11265
   Icon            =   "frmReplenishTheBalanceDel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   11265
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pic�˷�ժҪ 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   11265
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   4590
      Width           =   11265
      Begin VB.TextBox txt�˷�ժҪ 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1035
         MaxLength       =   100
         TabIndex        =   4
         Top             =   45
         Width           =   5820
      End
      Begin VB.Label lblժҪ 
         AutoSize        =   -1  'True
         Caption         =   "�˷�ժҪ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   45
         TabIndex        =   3
         Top             =   120
         Width           =   960
      End
   End
   Begin VB.PictureBox picCmd 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   11265
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   7404
      Width           =   11265
      Begin VB.CommandButton cmdBillSel 
         Caption         =   "ȫѡ��ǰ����(&B)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3240
         TabIndex        =   23
         ToolTipText     =   "�ȼ���Ctrl+B"
         Top             =   135
         Width           =   2040
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ȡ��(&C)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   9375
         TabIndex        =   11
         Top             =   135
         Width           =   1440
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "ȫ��(&R)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1695
         TabIndex        =   18
         ToolTipText     =   "�ȼ���Ctrl+R"
         Top             =   135
         Width           =   1440
      End
      Begin VB.CommandButton cmdSelAll 
         Caption         =   "ȫѡ(&A)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   165
         TabIndex        =   17
         ToolTipText     =   "�ȼ���Ctrl+A"
         Top             =   135
         Width           =   1440
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   7845
         TabIndex        =   10
         Top             =   144
         Width           =   1440
      End
      Begin VB.Line LineCmd_1 
         X1              =   0
         X2              =   12000
         Y1              =   0
         Y2              =   0
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   12
      Top             =   8064
      Width           =   11268
      _ExtentX        =   19870
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmReplenishTheBalanceDel.frx":058A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12224
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "���"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin VSFlex8Ctl.VSFlexGrid vsBill 
      Align           =   1  'Align Top
      Height          =   3630
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   11265
      _cx             =   19870
      _cy             =   6403
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
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
      BackColorSel    =   12632256
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   280
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmReplenishTheBalanceDel.frx":0E1E
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
      AllowUserFreezing=   1
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.PictureBox picMoney 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   11265
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5445
      Width           =   11265
      Begin VB.TextBox txt�˿�ϼ� 
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
         ForeColor       =   &H00C00000&
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   9984
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   75
         Width           =   1275
      End
      Begin VB.TextBox txtAllTotal 
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
         IMEMode         =   3  'DISABLE
         Left            =   3465
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   75
         Width           =   1275
      End
      Begin VB.TextBox txtCurTotal 
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
         IMEMode         =   3  'DISABLE
         Left            =   1065
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   75
         Width           =   1275
      End
      Begin VB.Label lbl�˿�ϼ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�˿�ϼ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   9000
         TabIndex        =   25
         Top             =   132
         Width           =   960
      End
      Begin VB.Label lblAllTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�շѺϼ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2460
         TabIndex        =   8
         Top             =   132
         Width           =   960
      End
      Begin VB.Label lblCurTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ǰ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   75
         TabIndex        =   6
         Top             =   135
         Width           =   960
      End
   End
   Begin VB.PictureBox picPati 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   0
      ScaleHeight     =   960
      ScaleWidth      =   11265
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   0
      Width           =   11265
      Begin zlIDKind.IDKindNew IDKindNO 
         Height          =   300
         Left            =   7725
         TabIndex        =   29
         Top             =   120
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   529
         ShowSortName    =   0   'False
         IDKindStr       =   "��|�շѵ���|0|0|0|0|0|0;��|��Ʊ��|0|0|0|0|0|0;��|�Һŵ���|0|0|0|0|0|0;��|���㵥��|0|0|0|0|0|0"
         CaptionAlignment=   1
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
         DefaultCardType =   "0"
         NotAutoAppendKind=   -1  'True
         AllowAutoCommCard=   0   'False
         BackColor       =   -2147483633
      End
      Begin VB.PictureBox picPatiBack 
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   540
         ScaleHeight     =   360
         ScaleWidth      =   2640
         TabIndex        =   26
         Top             =   525
         Width           =   2640
         Begin VB.TextBox txtPatient 
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Left            =   645
            MaxLength       =   100
            TabIndex        =   1
            ToolTipText     =   "��λ:F6,����:-����ID,*�����,+סԺ��,.�Һŵ���,����:*2536��ʾ������Ų���"
            Top             =   -15
            Width           =   1980
         End
         Begin zlIDKind.IDKindNew IDKind 
            Height          =   360
            Left            =   0
            TabIndex        =   27
            Top             =   0
            Width           =   660
            _ExtentX        =   1164
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
            MustSelectItems =   "����"
            BackColor       =   -2147483633
         End
      End
      Begin VB.PictureBox pic�� 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   10470
         ScaleHeight     =   360
         ScaleWidth      =   615
         TabIndex        =   20
         Top             =   45
         Visible         =   0   'False
         Width           =   645
         Begin VB.Label lbl�� 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
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
            Height          =   360
            Left            =   90
            TabIndex        =   21
            Top             =   0
            Width           =   405
         End
      End
      Begin VB.Frame fraInfo_1 
         Height          =   120
         Left            =   -120
         TabIndex        =   19
         Top             =   390
         Width           =   12000
      End
      Begin VB.TextBox txtNO 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   9315
         TabIndex        =   0
         Top             =   120
         Width           =   1800
      End
      Begin VB.Label lblFormat 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C000C0&
         Height          =   240
         Left            =   6735
         TabIndex        =   28
         Top             =   165
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label lblPati 
         AutoSize        =   -1  'True
         Caption         =   "����: "
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   45
         TabIndex        =   14
         Top             =   585
         Width           =   720
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsBalance 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5070
      Width           =   11265
      _cx             =   19870
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
      FormatString    =   $"frmReplenishTheBalanceDel.frx":0E98
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
End
Attribute VB_Name = "frmReplenishTheBalanceDel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Enum gEM_ReplenishBalanceDelType
    EM_RBDTY_�鿴 = 0
    EM_RBDTY_�˷� = 1
    EM_RBDTY_�쳣���� = 2
End Enum
'----------------------------------------------------------------
'�ӿڱ���
Private mstrPrivs As String
Private mbytMode As gEM_ReplenishBalanceDelType
Private mstr������� As String    'Ҫ�鿴���˷ѵĶ��ŵ����н������
Private mblnNOMoved As Boolean '�����ĵ����Ƿ��ں����ݱ���
Private mstrDelTime As String '�鿴�˷ѵ��ݵĵǼ�ʱ��(yyyy-MM-dd HH:mm:ss) 'ֻ�в鿴�˷ѵ���ʱ�Ŵ���ʱ��,��������������
Private mstr���㵥�� As String
'-----------------------------------------------------------
'ҽ���������
Private mstr�����ʻ� As String   'ҽ�������ʻ�������
Private Type TY_Insure
    dbl����͸֧ As Double
    dbl�ʻ���� As Double
End Type
Private mTy_Insure As TY_Insure
Private mlngModule  As Long
Private mlng����ID As Long
Private mblnOK As Boolean
Private mblnPrintView As Boolean    '��ӡǰ�鿴����
Private mblnFirst As Boolean
Private mblnNotClick As Boolean
Private mstrTittle As String
Private mstrNo As String 'Ҫ�鿴���˷ѵĶ��ŵ����е�ĳ��NO,�˷�ʱ����û��

Private mrs���㷽ʽ As ADODB.Recordset
Private mrs�շѶ��� As ADODB.Recordset '�շѶ��� :����:33634
Private mrsBalance As ADODB.Recordset '��¼ÿ�ŵ��ݵĽ������
Private mrsInfo As ADODB.Recordset

Private mobjPayCards As Cards

Private Type tyBillType
    str���㵥 As String
    bln�Һ�   As Boolean    '�Ƿ�ǰΪ�ҺŽ���
    strNos As String 'ʵ�ʶ��������˷ѵĵ��ݺ�
    strAllNOs As String '���е��ݺ�(һ���շѵ����е���)
    strDelNOs As String '��ǰѡ��Ҫ�˵ĵ���
    strNosOverFlow As String '����������޵ĵ��ݺ�
    strNosPatiDel As String '��¼�����˷ѵĵ���
    strNotCanDelNOs As String  '(�����˵ĵ���)�Ѿ�����ĵ��ݻ�ִ�в����˵ĵ���
    
    str���㷽ʽ As String '��ǰ���㷽ʽ:����ʱ,�ö��ŷָ�
    bln���ڿ����� As Boolean
    intInsure  As Integer   'ҽ�����ݵ�����
    bln���Ų����˷� As Boolean
    blnExistOnCard As Boolean '�Ƿ����һ��ͨ����
    blnExistThreeAllDel As Boolean '�Ƿ����һ��ͨȫ�˵�
    strInvoice As String '��ǰ��Ʊ��
    lngԭ����ID As Long
    lng����ID As Long '���½���ID
    lng����ID As Long '����ID
    lng������� As Long
    lng���ó���ID As Long '����ID
    lng����ID As Long
    str���� As String
    str�Ա� As String
    str���� As String
    str�ѱ� As String
    str�������� As String
End Type
Private mCurBillType As tyBillType  '��ǰ��������

Private mobjSquare As Object
Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
'-------------------------------------------------------------------------------
'��ͷ����
Private Type TY_ColHead
     strRegColHead As String
     strFeeColHead As String
End Type
Private mTyColHead As TY_ColHead
'-------------------------------------------------------------------------------
'ҽ����ض���:����
Private Type TYPE_MedicarePAR
    ҽ���ӿڴ�ӡƱ�� As Boolean
    �ֱҴ��� As Boolean
    �˷Ѻ��ӡ�ص� As Boolean
    ҽ������Ʊ��  As Boolean        'Ԥ����ʱ��Ч
    ����������� As Boolean             'ҽ���Ƿ�֧�������������
    ����Ԥ���� As Boolean
    ���Ը� As Boolean
    ȫ�Ը� As Boolean
End Type
Private MCPAR As TYPE_MedicarePAR
'-------------------------------------------------------------------------------
'Api����
Private Const VK_RETURN = &HD
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer


'-------------------------------------------------------------------------------
'��Ʊ����
Private mobjInvoice As zlPublicExpense.clsInvoice, mobjFact As zlPublicExpense.clsFactProperty
Private Type Ty_Module_Para
     int����ʣ��Ʊ������ As Integer
     blnҩ����λ As Boolean
     int�嵥��ӡ��ʽ As Integer
End Type
Private mtyMoudlePara As Ty_Module_Para
Private mobjDrugPacker  As Object ' �Զ���ҩ��(���·�ҩ����)
Private mblnDrugPacker As Boolean
Private mobjDrugMachine As Object
Private mblnDrugMachine As Boolean
Private mcllForceDelToCash As Collection 'ǿ��������Ϣ��Array(����Ա,���������,���㷽ʽ)
Private mstr�ų����㷽ʽ As String '����ʹ�õĽ��㷽ʽ,����ö��ŷָ�

Private Sub InitModulePara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ģ�����
    '����:���˺�
    '����:2014-09-16 16:28:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String, varTemp As Variant
    With mtyMoudlePara
        .blnҩ����λ = zlDatabase.GetPara("ҩƷ��λ��ʾ", glngSys, mlngModule) = "1"
        .int�嵥��ӡ��ʽ = Val(zlDatabase.GetPara("�����嵥��ӡ��ʽ", glngSys, mlngModule))
        strTemp = Trim(zlDatabase.GetPara("Ʊ��ʣ��X��ʱ��ʼ�����շ�Ա", glngSys, mlngModule, "0|10"))
        varTemp = Split(strTemp & "|", "|")
        If Val(varTemp(0)) = 0 Then
            .int����ʣ��Ʊ������ = -1
        Else
            .int����ʣ��Ʊ������ = Val(varTemp(1))
        End If
    End With
End Sub
Private Function CheckDepend() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݹ����Լ��
    '����:���ݹ������Ϸ�����true,���򷵻�False
    '����:���˺�
    '����:2014-07-07 11:41:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mbytMode = EM_RBDTY_�鿴 Then CheckDepend = True: Exit Function
    
    Set mrs���㷽ʽ = Get���㷽ʽ("�շ�")
    mrs���㷽ʽ.Filter = "����=3"
    If Not mrs���㷽ʽ.EOF Then
       mstr�����ʻ� = mrs���㷽ʽ!����
    End If
    mrs���㷽ʽ.Filter = 0
    If mrs���㷽ʽ.RecordCount = 0 Then
        MsgBox "�շѳ���û�п��õĽ��㷽ʽ�����ȵ����㷽ʽ���������á�", vbInformation, gstrSysName
        Exit Function
    End If
    mrs���㷽ʽ.MoveFirst
    
    Set mobjPayCards = GetPayCardsObject
    If mobjPayCards Is Nothing Then Exit Function
    If mobjPayCards.Count = 0 Then Exit Function
    
    CheckDepend = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetPayCardsObject() As Cards
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������֧�ֵĽ���������
    '����:����Cards����
    '����:���˺�
    '����:2015-03-18 09:56:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card, objCards As Cards, objPayCards As Cards
    Dim rsTemp As ADODB.Recordset
    Dim lngKey As Long, i As Long, blnFind As Boolean
    
    On Error GoTo errHandle
    
    Set objCards = New Cards: Set objPayCards = New Cards
    Set rsTemp = Get���㷽ʽ("������")
    '83533:���ϴ�,2015/3/25,û����Ч�Ĳ�����
    If rsTemp.RecordCount = 0 Then
        MsgBox "������û�п��õĽ��㷽ʽ�����ȵ������㷽ʽ���������ò������Ӧ�ó��ϡ�", vbInformation, gstrSysName
        Exit Function
    End If
    If Not gobjSquare Is Nothing Then
        ' zlGetCards(ByVal BytType As Byte)
        '���:bytType-0-����ҽ�ƿ�;
        '             1-���õ�ҽ�ƿ�,
        '             2-���д��������˻���������
        '             3-���õ������˻���ҽ�ƿ�
       Set objCards = gobjSquare.objSquareCard.zlGetCards(3)
    End If
    
    With rsTemp
        .Filter = 0
        If .RecordCount <> 0 Then .MoveFirst
        lngKey = 1
        Do While Not .EOF
            For i = 1 To objCards.Count
                If objCards(i).���㷽ʽ = Nvl(rsTemp!����) Then
                    blnFind = True
                    Exit For
                End If
            Next
            If Not blnFind Then
                '83266:���ϴ�,2015/3/18,ҽ�ƿ������ж��Ƿ�����
                If InStr(",1,2,", "," & Val(Nvl(rsTemp!����)) & ",") > 0 _
                    And Val(Nvl(rsTemp!Ӧ����)) <> 1 Then
                    '������ҽ���Ľ��㷽ʽ����֧Ʊ��
                     Set objCard = New Card
                     objCard.���� = Mid(Nvl(!����), 1, 1)
                     objCard.�ӿڱ��� = Nvl(!����)
                     objCard.�ӿڳ����� = ""
                     objCard.�ӿ���� = -1 * lngKey
                     objCard.���㷽ʽ = Nvl(!����)
                     objCard.���� = Nvl(!����)
                     objCard.���� = True
                     objCard.ȱʡ��־ = Val(Nvl(rsTemp!ȱʡ)) = 1
                     objCard.֧������ = True
                     objCard.�������� = Val(!����)
                    objPayCards.Add objCard, "K" & lngKey
                    lngKey = lngKey + 1
                End If
            End If
            .MoveNext
        Loop
    End With
    '��������
    For Each objCard In objCards
        If objCard.���ѿ� = False Then 'And objCard.�Ƿ�ת�ʼ����� Then
            rsTemp.Filter = "����='" & objCard.���㷽ʽ & "'"
            If Not rsTemp.EOF Then
                objPayCards.Add objCard, "K" & lngKey
                lngKey = lngKey + 1
            End If
        End If
    Next
    If objPayCards.Count = 0 Then
        MsgBox "���㿨��������,ԭ���������:" & vbCrLf & _
            "δ�������ý��㿨,�뵽��ҽ�ƿ���𡻺͡��豸���á�������", vbInformation, gstrSysName
    End If
    Set GetPayCardsObject = objPayCards
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function zlShowMe(frmMain As Object, ByVal lngModule As Long, _
    ByVal strPrivs As String, ByVal bytMode As gEM_ReplenishBalanceDelType, _
    Optional ByVal str������� As String, _
    Optional blnPrintView As Boolean, _
    Optional lng����ID As Long = 0, _
    Optional blnNOMoved As Boolean = False, Optional strDelTime As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݲ鿴,�˷�
    '���:bytMode-0-���ŵ��ݲ鿴,1-���ŵ����˷�,2-���쳣���˷ѵ����������˷�
    '     strPrivs-Ȩ�޴�
    '     str�������-�˷�ѡ�еĽ��㵥��
    '     blnPrintView-��ӡǰ�鿴����
    '     blnNOMoved-�Ƿ�ת�������ݱ�
    '     strDelTime-�鿴�˷ѵ��ݵĵǼ�ʱ��(yyyy-MM-dd HH:mm:ss),ֻ�в鿴�˷ѵ���ʱ�Ŵ���ʱ��,��������������
    '����:�˷ѳɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-09-30 17:10:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mblnNOMoved = blnNOMoved: mstrPrivs = strPrivs
    mlng����ID = lng����ID: mstr������� = str�������
    mlngModule = lngModule: mblnPrintView = blnPrintView
    mbytMode = bytMode: mstrDelTime = strDelTime              'ֻ�в鿴�˷ѵ���ʱ�Ŵ���ʱ��,��������������
    mblnOK = False
    If CheckDepend = False Then Exit Function
    On Error Resume Next
    Me.Show IIf(gfrmMain Is Nothing, 0, 1), frmMain
    On Error GoTo 0
    zlShowMe = mblnOK
End Function

Private Sub Form_Load()
    mblnFirst = True
    Call InitFace
    Call RestoreWinState(Me, App.ProductName, mstrTittle)

    If mstr������� <> "" Then    'ָ���˽������ݵ�
        If mbytMode = EM_RBDTY_�˷� Then
        'intFindType -0 - ��������Ų���
        '             1-���շѵ��ݺŲ���
        '             2.�����㵥�Ų���
        '             3.������ķ�Ʊ�Ų���
        '             4.���Һŵ��Ų���
            If ReadBills(0, mstr�������) = False Then Unload Me: Exit Sub
        Else
            If LoadViewBills(mstr�������) = False Then Unload Me: Exit Sub
        End If
    End If
    
    '81190,Ƚ����,�˷�ҵ����ҩ���ϴ��˷���Ϣ
    Call CreateDrugPacker
End Sub

Private Sub CreateDrugPacker()
    '����:����������ҩ��(�Զ���ҩ��)
    Dim objComLib As New zl9ComLib.clsComLib
    Dim strPrivs As String
    Dim strMessage As String
    
    '81190,Ƚ����,�˷�ҵ����ҩ���ϴ��˷���Ϣ
    mblnDrugPacker = False: mblnDrugMachine = False
    If Not (mbytMode = EM_RBDTY_�˷� Or mbytMode = EM_RBDTY_�쳣����) Then Exit Sub

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

Private Sub InitFace()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������
    '����:���˺�
    '����:2014-06-24 14:36:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim TY_Temp As tyBillType, bytTemp As Byte
    
    mCurBillType = TY_Temp
    
    Set mobjInvoice = New zlPublicExpense.clsInvoice
    Call mobjInvoice.zlInitCommon(glngSys, gcnOracle, gstrDBUser)
    Set mobjFact = New zlPublicExpense.clsFactProperty
    
    
    Call InitModulePara '��ʼ��ģ�����
    
    Call InitBillHead(True)    '���ùҺŵ�����ͷ
    mTyColHead.strRegColHead = zl_vsGrid_GetCols_Property(vsBill)
    
    Call InitBillHead(False)       '���÷�����ͷ
    mTyColHead.strFeeColHead = zl_vsGrid_GetCols_Property(vsBill)
    
    bytTemp = Val(zlDatabase.GetPara("�˷Ѻ�������ģʽ", glngSys, mlngModule, 0))
    IDKindNO.IDKindStr = "��|�շѵ���;��|��Ʊ��;��|�Һŵ���;��|���㵥��"
    IDKindNO.IDKind = bytTemp
    
    Call NewCardObject
    Call ClearFace
    Call SetFunCtrlVisible
    
    Select Case mbytMode
    Case EM_RBDTY_�鿴
        mstrTittle = "���ղ������-����"
        Caption = mstrTittle
        vsBill.ColHidden(0) = True
        cmdOK.Visible = False
        cmdCancel.Caption = "�˳�(&X)"
        If mblnPrintView Then cmdCancel.Caption = "ȷ��(&X)"
        pic��.Visible = mstrDelTime <> ""
        lbl�˿�ϼ�.Visible = mstrDelTime <> ""
        txt�˿�ϼ�.Visible = mstrDelTime <> ""
    Case EM_RBDTY_�쳣����
        mstrTittle = "���ղ������-�쳣�˷ѵ������˷�"
        Caption = mstrTittle
        vsBill.ColHidden(0) = True
        cmdOK.Visible = True
        pic��.Visible = mstrDelTime <> ""
        vsBill.Editable = flexEDNone
        Call initCardSquareData
    Case Else 'EM_RBDTY_�˷�
        mstrTittle = "���ղ������-�˷�"
        Caption = mstrTittle
        Call initCardSquareData
    End Select
    If mstr������� <> "" Then
        picPatiBack.Top = txtNO.Top
        lblPati.Top = picPatiBack.Top + (picPatiBack.Height - lblPati.Height) \ 2
        picPati.Height = 480
    End If
End Sub

Private Sub InitBillHead(ByVal bln�Һ� As Boolean, Optional blnInit As Boolean = True)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���˷ѻ��˺ŵı�ͷ����Ϣ
    '���:bln�Һ�-�Ƿ��˺�:true�˺�,False-�˷�
    '����:���˺�
    '����:2014-06-24 14:39:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim arrHead As Variant, strHead As String, i As Long
    Dim varTemp As Variant, intCol As Integer
    
    If bln�Һ� Then
        strHead = "" & _
        "ѡ��,300,4;���ݺ�,1000,1;���,720,1;��Ŀ,2800,1;����,750,7;��λ,550,1;����,1100,7;" & _
        "Ӧ�ս��,1100,7;ʵ�ս��,1100,7;��������,1000,1;ִ�п���,1000,1;ҽ��,850,1;����Ա,850,1;" & _
        "�Ǽ�ʱ��,1400,1;����ʱ��,1400,1;ԤԼʱ��,1400;����ʱ��,1400,1;����ʱ��,1400,1;����,1000,1;����,720,1;����,720,1;����ID;" & _
        "ԭʼ����,0,4;׼������,0,4;ҽ�����,0,4;ִ�п���ID,0,1"
    Else
        strHead = "" & _
        "ѡ��,300,4;���ݺ�,1000,1;���,720,1;��Ŀ,2800,1;��Ʒ��,2000,1;����,750,7;��λ,550,1;����,1100,7;" & _
        "Ӧ�ս��,1100,7;ʵ�ս��,1100,7;��������,1000,1;ִ�п���,1000,1;����Ա,850,1;ʱ��,1260,1;����ID;ҽ��,1560,1;" & _
        "ԭʼ����,0,4;׼������,0,4;ҽ�����,0,4;ִ�п���ID,0,1"
    End If
    
    arrHead = Split(strHead, ";")
    With vsBill
        .Clear
        .FixedRows = 1: .FixedCols = 0
        .COLS = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        For i = 0 To UBound(arrHead)
            varTemp = Split(arrHead(i) & ",,,", ",")
            intCol = .FixedCols + i
            .ColKey(intCol) = Trim(varTemp(0))
            .TextMatrix(.FixedRows - 1, intCol) = varTemp(0)
            If UBound(varTemp) > 0 Then
                .ColHidden(intCol) = False
                .ColWidth(intCol) = Val(varTemp(1))
                If .ColWidth(intCol) = 0 Then .ColHidden(intCol) = True
                .ColAlignment(intCol) = Val(varTemp(2))
            Else
                .ColHidden(intCol) = True
            End If
        Next
         .TextMatrix(.FixedRows - 1, .ColIndex("ѡ��")) = ""
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .COLS - 1) = 4
        If Not bln�Һ� Then .ColHidden(.ColIndex("��Ʒ��")) = gTy_System_Para.bytҩƷ������ʾ <> 2
        zl_vsGrid_Para_Restore mlngModule, vsBill, mstrTittle, IIf(bln�Һ�, "�Һ���ͷ��Ϣ", "������ͷ��Ϣ")
        
        If Not blnInit Then
            zl_vsGrid_RestoreCols_Property vsBill, IIf(bln�Һ�, mTyColHead.strRegColHead, mTyColHead.strFeeColHead)
        End If
        .FrozenCols = 2
        .ColHidden(.ColIndex("ѡ��")) = True
        .Editable = flexEDNone
        If mbytMode = EM_RBDTY_�˷� Then
            .ColDataType(.ColIndex("ѡ��")) = flexDTBoolean
            .ColHidden(.ColIndex("ѡ��")) = False
            .Editable = flexEDKbdMouse
        End If
    End With
    
End Sub

Private Sub ClearFace(Optional ByVal blnNO As Boolean = True)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������Ϣ
    '���:blnNo=������ݺ�
    '����:���˺�
    '����:2014-06-24 15:19:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim tmpBillType As tyBillType
    
    mCurBillType = tmpBillType
    Set mrsBalance = Nothing
    With vsBill
        .Rows = .FixedRows '�Էǹ̶��еĵ�һ�б�����ʱ�ָ��ɼ�
        .Rows = .FixedRows + 1
        .Row = .FixedRows: .Col = .ColIndex("��Ŀ")
        .Clear 1
    End With
    lblPati.Caption = "����:"
    If blnNO Then txtNO.Text = ""
    
    Call ClearBalance
    With vsBalance
         .COLS = 1
         .TextMatrix(0, 0) = IIf(mstrDelTime = "", "�տ����", "�˿����")
    End With
    txtCurTotal.Text = ""
    txtAllTotal.Text = ""
    txt�˿�ϼ�.Text = ""
    stbThis.Panels(2).Text = ""
    Call SetFunCtrlVisible
End Sub

Private Sub NewCardObject()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���µĿ�����
    '����:���˺�
    '����:2014-06-24 14:43:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mbytMode <> EM_RBDTY_�鿴 Then Exit Sub
   
    If mobjIDCard Is Nothing Then
        Set mobjIDCard = New clsIDCard
        Call mobjIDCard.SetParent(Me.hWnd)
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(False)
    End If
    If mobjICCard Is Nothing Then
        Set mobjICCard = New clsICCard
        Call mobjICCard.SetParent(Me.hWnd)
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(False)
    End If
    IDKind.SetAutoReadCard (False)
End Sub
Private Sub CloseIDCard()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ر�������������
    '����:���˺�
    '����:2014-06-24 14:43:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not mobjIDCard Is Nothing Then
        mobjIDCard.SetEnabled (False)
        Set mobjIDCard = Nothing
    End If
    If Not mobjICCard Is Nothing Then
        mobjICCard.SetEnabled (False)
        Set mobjICCard = Nothing
    End If
End Sub

Private Function LoadViewBills(ByVal str������� As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݽ����������������(ֻ��Բ鿴���쳣�˷�)
    '���:str�������-�������
    '����:���ػ��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-06-24 16:17:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllInvoiceNoInfor As Collection
    Dim rsTemp As ADODB.Recordset, rsAdvice As ADODB.Recordset
    Dim str����IDs As String, strNos As String, strAllNOs As String, strFeeNos As String, strRegNos As String
    Dim strTemp As String, strWhere As String, strSQL As String, strҽ����� As String
    Dim lng����ID As Long, lngԭ������� As Long, j As Long, lngҽ����� As Long, i As Long, lng����ID As Long
    Dim intInsure As Integer, intSign As Integer
    Dim dbl�ϼ� As Double
    Dim varData As Variant
 
    Screen.MousePointer = 11
    intSign = IIf(mstrDelTime <> "", -1, 1) '����,�����������
    On Error GoTo errHandle
    
    str����IDs = zlGet����ID(Val(str�������), strNos, intInsure, mblnNOMoved, lng����ID, True)
    
    mCurBillType.str���㵥 = strNos
    mCurBillType.lng����ID = lng����ID
    mCurBillType.intInsure = intInsure
    
    varData = Split(str����IDs & ",,", ",")
    If Val(varData(0)) = lng����ID Then
         mCurBillType.lng����ID = Val(varData(1))
    ElseIf Val(varData(0)) = lng����ID Then
         mCurBillType.lng����ID = Val(varData(0))
    End If
    
    
    strSQL = "" & _
    " Select A.����ID,B.����,B.�Ա�,B.����,B.�����,B.�ѱ�,b.ҽ�Ƹ��ʽ as ���ʽ,B.��������,B.����,nvl(A.���ӱ�־,0) as ���ӱ�־" & _
    " From ���ò����¼ A, ��Ա�� D,������Ϣ B" & _
    " Where  A.�������=[1] And A.����Ա����=D.���� And A.����ID=B.����ID(+) " & _
    "       And (D.վ��='" & gstrNodeNo & "' Or D.վ�� is Null)" & vbNewLine & _
    "       And mod(A.��¼����,10)=1 And Rownum <2 "
    
    If mblnNOMoved Then
        strSQL = Replace(strSQL, "���ò����¼", "H���ò����¼")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(str�������))
    If rsTemp.EOF Then
        Screen.MousePointer = 0
        MsgBox "û���ҵ��������صļ�¼��", vbInformation, gstrSysName
        mCurBillType.lng����ID = 0
        Exit Function
    End If
    mCurBillType.bln�Һ� = Val(Nvl(rsTemp!���ӱ�־)) = 1
    txtPatient.Text = Nvl(rsTemp!����)
    lblPati.Caption = "����:" & IIf(txtPatient.Visible, "       ", rsTemp!����) & _
        "���Ա�:" & Nvl(rsTemp!�Ա�) & _
        "������:" & Nvl(rsTemp!����) & _
        "�������:" & Nvl(rsTemp!�����) & _
        "���ѱ�:" & Nvl(rsTemp!�ѱ�) & _
        "�����ʽ:" & rsTemp!���ʽ
    
    With mCurBillType
        .lng����ID = Val(Nvl(rsTemp!����ID))
        .str�Ա� = Nvl(rsTemp!�Ա�)
        .str���� = Nvl(rsTemp!����)
        .str���� = Nvl(rsTemp!����)
        .str�������� = Nvl(rsTemp!��������)
        .lngԭ����ID = zlGetFromNOToLastBalanceID(mCurBillType.str���㵥, False, , , True)
    End With
    
    
    If mbytMode <> EM_RBDTY_�鿴 Then
        Call initInsurePara(mCurBillType.intInsure, mCurBillType.lng����ID, lng����ID)
    
        'bytType-0-����NO������;1-���ݽ���ID������,2-���ݽ������������
        If GetBalanceFeeNos(0, mCurBillType.str���㵥, strFeeNos, strRegNos, mblnNOMoved) = False Then Exit Function
        If mCurBillType.bln�Һ� Then
            mCurBillType.strAllNOs = strRegNos
            strNos = strRegNos
        Else
            mCurBillType.strAllNOs = strFeeNos
            strNos = strFeeNos
        End If
    End If
    
    If CheckPrivsIsValied = False Then Exit Function    '����Ȩ�޼��
    lblPati.ForeColor = vbRed
    txtPatient.ForeColor = vbRed
    Call SetPatiColor(txtPatient, Nvl(rsTemp!��������), txtPatient.ForeColor)
    lblPati.ForeColor = txtPatient.ForeColor
    
    '���ؽ��㷽ʽ
    Set mrsBalance = GetChargeBalance(1, str�������, mblnNOMoved)
    Call LoadBalanceInfor
 
    'InStr(str����ID, ",") > 0:��ʾ���ܴ������յ���������Կ϶��ǲ���˷Ѽ�¼������ժҪӦ�����˷ѵ�ժҪΪ׼
    strSQL = "" & _
    "   Select A.NO,Nvl(A.�۸񸸺�,A.���) as ���,A.��������,A.��������ID,A.ִ�в���ID,A.�շ����,A.�ѱ�,A.�շ�ϸĿID," & _
    "          A.��������,A.���㵥λ,max(A.ҽ�����) as ҽ�����," & _
    "          Avg(Nvl(A.����,1)*A.����) as ����," & _
    "          Sum(A.��׼����) as ����, Sum(A.Ӧ�ս��) as Ӧ�ս��,Sum(A.ʵ�ս��) as ʵ�ս��," & _
    "          Max(A.����Ա����) as ����Ա����,max(A.�Ǽ�ʱ��) as �Ǽ�ʱ��," & _
    "           " & IIf(InStr(str����IDs, ",") > 0, "Max(Decode(A.��¼״̬,2,A.ժҪ,NULL))", "Max(A.ժҪ)") & " as ժҪ,A.����ID" & _
    "   From ������ü�¼ A," & _
    "       (Select ����id" & _
    "         From (Select T1.�շѽ���id As ����id From ���ò����¼ T1 Where T1.������� = [1]" & _
    "                Union All" & _
    "                Select T1.�շѽ���id As ����id From ���ò����¼ T1 Where T1.������� = [1]" & _
    "                      And Not Exists (Select 1 From ���ò����¼ Where T1.������� = ������� And ��¼״̬ In (1, 3))" & _
    "                Union All" & _
    "                Select Distinct ����id From ����Ԥ����¼ T1 Where ������� = [1] And ����id = Abs(�������)" & _
    "                       And Not Exists (Select 1 From ���ò����¼ Where ������� = [1] And ��¼״̬ In (1, 3)))" & _
    "         Group By ����id" & _
    "  Having Count(*) <= 1) B" & _
    "   Where Mod(A.��¼����,10)= [2] and A.����ID=B.����ID  " & _
    "   Group by A.����ID,A.NO,Nvl(A.�۸񸸺�,A.���),A.��������,A.��������ID,A.ִ�в���ID,A.�շ����,A.�ѱ�,A.�շ�ϸĿID,A.��������,A.���㵥λ"
    
    If mblnNOMoved Then
        strSQL = Replace(strSQL, "������ü�¼", "H������ü�¼")
        strSQL = Replace(strSQL, "���ò����¼", "H���ò����¼")
    End If
    If mCurBillType.bln�Һ� Then
        strSQL = _
            " Select A.NO,A.���,A.��������,A.�ѱ�,A.�շ�ϸĿID,C.���� as �����,C.���� as �����,B.����, " & _
            "       Nvl(M1.����,B.����) as ����,Max(Nvl(A.��������,B.��������)) ��������," & _
            "       A.���㵥λ  as ���㵥λ,Max(A.ҽ�����) as  ҽ�����," & _
            "       sum(A.����) as ����,Max(A.����) as ����,Sum(A.Ӧ�ս��) as Ӧ�ս��,Sum(A.ʵ�ս��) as ʵ�ս��," & _
            "       1 as ��¼��־,0 as ԭʼ����,0 as ׼������," & _
            "       D.���� as ִ�п���,A.ִ�в���ID,E.���� as ��������,Max(A.����Ա����) As ����Ա����,Max(B1.ִ����) as ҽ��, " & _
            "       Max(A.�Ǽ�ʱ��) As �Ǽ�ʱ��,Max(B1.����ʱ��) as ����ʱ��,max(B1.ԤԼʱ��) as ԤԼʱ��,max(B1.����ʱ��) as ����ʱ��,max( B1.����ʱ��) as ����ʱ��, " & _
            "       Max(B1.����) as ����,max(B1.����) as ����,max( B1.�ű�) as ����,  " & _
            "       Max(A.ժҪ) as ժҪ" & _
            " From (" & strSQL & ") A,���˹Һż�¼ B1,�շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D,���ű� E,ҩƷ��� X," & _
            "       �շ���Ŀ���� M1,�շ���Ŀ���� E1" & _
            " Where A.NO=B1.NO  And B1.��¼״̬ in (1,3) And " & _
            "       A.�շ�ϸĿID=B.ID And C.����=A.�շ���� And A.�շ�ϸĿID=X.ҩƷID(+)" & _
            "       And A.ִ�в���ID=D.ID(+) And A.��������ID=E.ID(+) " & _
            "       And A.�շ�ϸĿID=M1.�շ�ϸĿID(+) And M1.����(+)=1 And M1.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
            "       And A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3 " & _
            " Group by A.NO,A.���,A.��������,A.�ѱ�,A.�շ�ϸĿID,C.����,C.����,B.����,Nvl(M1.����,B.����)," & _
            "       E1.����,B.���,A.���㵥λ,D.����,A.ִ�в���ID,E.����,X.ҩƷID,X." & gstrҩ����λ & _
            " Having Sum(A.����)<>0 " & _
            " Order by NO,���"
        If mblnNOMoved Then
            strSQL = Replace(strSQL, "���˹Һż�¼", "H���˹Һż�¼")
        End If
    Else
        strSQL = _
        " Select A.NO,A.���,A.��������,A.�ѱ�,A.�շ�ϸĿID,C.���� as �����,C.���� as �����,B.����, " & _
        "       Nvl(M1.����,B.����) as ����,E1.���� as ��Ʒ�� ,B.���,Max(Nvl(A.��������,B.��������)) ��������," & _
                IIf(mtyMoudlePara.blnҩ����λ, "Decode(X.ҩƷID,NULL,A.���㵥λ,X." & gstrҩ����λ & ")", "A.���㵥λ") & " as ���㵥λ," & _
        "       Max(A.ҽ�����) as ҽ�����," & _
        "       sum(A.����" & IIf(mtyMoudlePara.blnҩ����λ, "/Nvl(X." & gstrҩ����װ & ",1)", "") & ") as ����," & _
        "       Max(A.����" & IIf(mtyMoudlePara.blnҩ����λ, "*Nvl(X." & gstrҩ����װ & ",1)", "") & ") as ����," & _
        "       Sum(A.Ӧ�ս��) as Ӧ�ս��,Sum(A.ʵ�ս��) as ʵ�ս��, 1 as ��¼��־,0 as ԭʼ����,0 as ׼������," & _
        "       D.���� as ִ�п���,A.ִ�в���ID,E.���� as ��������,Max(a.����Ա����) As ����Ա����, Max(a.�Ǽ�ʱ��) As �Ǽ�ʱ��, " & _
        "       Max(A.ժҪ) as ժҪ" & _
        " From (" & strSQL & ") A,�շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D,���ű� E,ҩƷ��� X," & _
        "       �շ���Ŀ���� M1,�շ���Ŀ���� E1" & _
        " Where A.�շ�ϸĿID=B.ID And C.����=A.�շ���� And A.�շ�ϸĿID=X.ҩƷID(+)" & _
        "       And A.ִ�в���ID=D.ID(+) And A.��������ID=E.ID(+) " & _
        "       And A.�շ�ϸĿID=M1.�շ�ϸĿID(+) And M1.����(+)=1 And M1.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
        "       And A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3 " & _
        " Group by A.NO,A.���,A.��������,A.�ѱ�,A.�շ�ϸĿID,C.����,C.����,B.����,Nvl(M1.����,B.����)," & _
        "       E1.����,B.���,A.���㵥λ,D.����,A.ִ�в���ID,E.����,X.ҩƷID,X." & gstrҩ����λ & _
        " Having Sum(A.����)<>0 " & _
        " Order by NO,���"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str�������, IIf(mCurBillType.bln�Һ�, 4, 1))
    
    If rsTemp.EOF Then
        Screen.MousePointer = 0
        If mbytMode = EM_RBDTY_�鿴 Then
            MsgBox "û���ҵ�ָ��������Ϣ�ķ��ü�¼,�����򲢷�ԭ�����˲���������˴���Ľ��㵥�ݡ�", vbInformation, gstrSysName
        Else
            MsgBox "û���ҵ��������Ϣ��صĿ����˷ѵļ�¼��" & _
                vbCrLf & "��Щ�շѼ�¼�����Ѿ��˷ѻ��Ѿ���ȫִ�С�", vbInformation, gstrSysName
        End If
        Call ClearFace(False)
        Exit Function
    End If
    
    If mbytMode <> EM_RBDTY_�˷� Then
        pic�˷�ժҪ.Enabled = mbytMode = EM_RBDTY_�쳣����
        txt�˷�ժҪ.Text = Nvl(rsTemp!ժҪ)
    End If
    strҽ����� = ""
    If Not mCurBillType.bln�Һ� Then
        With rsTemp
            Do While Not .EOF
                lngҽ����� = Val(Nvl(!ҽ�����))
                If InStr(strҽ����� & ",", "," & lngҽ����� & ",") = 0 And lngҽ����� <> 0 Then
                    strҽ����� = strҽ����� & "," & Val(Nvl(!ҽ�����))
                End If
                .MoveNext
            Loop
            .MoveFirst
        End With
    End If
    
    Set rsAdvice = Nothing
    If strҽ����� <> "" Then
        strҽ����� = Mid(strҽ�����, 2)
        Set rsAdvice = zlGetAdviceFromID(strҽ�����)
    End If
    
    Call InitBillHead(mCurBillType.bln�Һ�, False)
    stbThis.Panels(2).Text = "��ǰ���㵥��:" & mCurBillType.str���㵥
    With vsBill
        .Redraw = flexRDNone
        .Rows = .FixedRows + rsTemp.RecordCount
        mCurBillType.strNos = ""
        For i = 1 To rsTemp.RecordCount
            .RowData(i) = Val(Nvl(rsTemp!���))
            .TextMatrix(i, .ColIndex("ѡ��")) = 0
            .Cell(flexcpData, i, .ColIndex("��Ŀ")) = Val(Nvl(rsTemp!��������))
            .Cell(flexcpData, i, .ColIndex("����ID")) = Nvl(rsTemp!ҽ�����) & "," & Nvl(rsTemp!�շ�ϸĿID)
            strTemp = ""
            If Val(Nvl(rsTemp!��������)) <> 0 Then
                rsTemp.MoveNext
                strTemp = "��"
                If rsTemp.EOF Then
                    strTemp = "��"
                ElseIf Val(.Cell(flexcpData, i, .ColIndex("��Ŀ"))) <> Nvl(rsTemp!��������) Then
                    strTemp = "��"
                End If
                rsTemp.MovePrevious
                strTemp = "  " & strTemp & " "
            End If
            .TextMatrix(i, .ColIndex("���ݺ�")) = Nvl(rsTemp!NO)
            .TextMatrix(i, .ColIndex("���")) = Nvl(rsTemp!�����)
            If mCurBillType.bln�Һ� Then
                .TextMatrix(i, .ColIndex("��Ŀ")) = strTemp & rsTemp!����
                .TextMatrix(i, .ColIndex("����")) = FormatEx(intSign * rsTemp!����, 5)
                .Cell(flexcpData, i, .ColIndex("����")) = intSign * Val(Nvl(rsTemp!����))
            Else
                .TextMatrix(i, .ColIndex("��Ŀ")) = strTemp & rsTemp!���� & IIf(IsNull(rsTemp!���), "", " " & rsTemp!���)
                .TextMatrix(i, .ColIndex("��Ʒ��")) = strTemp & Nvl(rsTemp!��Ʒ��)
                .TextMatrix(i, .ColIndex("����")) = FormatEx(intSign * Val(Nvl(rsTemp!����)), 5)
                .Cell(flexcpData, i, .ColIndex("����")) = intSign * Val(Nvl(rsTemp!����))
            End If
            
            .TextMatrix(i, .ColIndex("��λ")) = Nvl(rsTemp!���㵥λ)
            .TextMatrix(i, .ColIndex("����")) = Format(rsTemp!����, gstrFeePrecisionFmt)
            .TextMatrix(i, .ColIndex("Ӧ�ս��")) = Format(intSign * Val(Nvl(rsTemp!Ӧ�ս��)), gstrDec)
            .TextMatrix(i, .ColIndex("ʵ�ս��")) = Format(intSign * Val(Nvl(rsTemp!ʵ�ս��)), gstrDec)
            .TextMatrix(i, .ColIndex("��������")) = Nvl(rsTemp!��������)
            .TextMatrix(i, .ColIndex("ִ�п���")) = Nvl(rsTemp!ִ�п���)
            .TextMatrix(i, .ColIndex("����Ա")) = rsTemp!����Ա����
            If mCurBillType.bln�Һ� Then
                .TextMatrix(i, .ColIndex("ҽ��")) = Nvl(rsTemp!ҽ��)
                .TextMatrix(i, .ColIndex("�Ǽ�ʱ��")) = Format(rsTemp!�Ǽ�ʱ��, "MM-dd HH:mm")
                .TextMatrix(i, .ColIndex("����ʱ��")) = Format(rsTemp!����ʱ��, "MM-dd HH:mm")
                .TextMatrix(i, .ColIndex("ԤԼʱ��")) = Format(rsTemp!ԤԼʱ��, "MM-dd HH:mm")
                .TextMatrix(i, .ColIndex("����ʱ��")) = Format(rsTemp!����ʱ��, "MM-dd HH:mm")
                .TextMatrix(i, .ColIndex("����ʱ��")) = Format(rsTemp!����ʱ��, "MM-dd HH:mm")
                .TextMatrix(i, .ColIndex("����")) = Nvl(rsTemp!����)
                .TextMatrix(i, .ColIndex("����")) = Nvl(rsTemp!����)
                .TextMatrix(i, .ColIndex("����")) = Nvl(rsTemp!����)
            Else
                .TextMatrix(i, .ColIndex("ʱ��")) = Format(rsTemp!�Ǽ�ʱ��, "MM-dd HH:mm")
                .TextMatrix(i, .ColIndex("ҽ�����")) = Nvl(rsTemp!ҽ�����)
            End If
            .TextMatrix(i, .ColIndex("����ID")) = str����IDs
            str������� = Val(Nvl(rsTemp!ҽ�����))
            If Not rsAdvice Is Nothing And strҽ����� <> "" And Val(str�������) <> 0 Then
                rsAdvice.Filter = "ҽ��ID=" & Val(str�������)
                If rsAdvice.EOF = False Then
                    .TextMatrix(i, .ColIndex("ҽ��")) = Nvl(rsAdvice!ҽ������)
                End If
            End If
            .TextMatrix(i, .ColIndex("ԭʼ����")) = Nvl(rsTemp!ԭʼ����)
            .TextMatrix(i, .ColIndex("׼������")) = Nvl(rsTemp!׼������)
            .TextMatrix(i, .ColIndex("ҽ�����")) = Nvl(rsTemp!ҽ�����)
            .TextMatrix(i, .ColIndex("ִ�п���ID")) = Nvl(rsTemp!ִ�в���ID)
            .Cell(flexcpData, i, .ColIndex("ѡ��")) = Val(Nvl(rsTemp!��¼��־))    '�����ж��Ƿ����ʹ�,>1��ʾ������
            If Val(Nvl(rsTemp!��¼��־)) > 1 And InStr(1, mCurBillType.strNosPatiDel & ",", "," & rsTemp!NO & "") = 0 Then mCurBillType.strNosPatiDel = mCurBillType.strNosPatiDel & "," & rsTemp!NO
            If InStr(mCurBillType.strNos & ",", "," & rsTemp!NO & ",") = 0 Then
                '�����ָ���
                If mCurBillType.strNos <> "" Then
                    .Select i, .FixedCols, i, .COLS - 1
                    .CellBorder vbBlue, 0, 1, 0, 0, 0, 0
                End If
                mCurBillType.strNos = mCurBillType.strNos & "," & rsTemp!NO
            End If
            dbl�ϼ� = dbl�ϼ� + Val(Nvl(rsTemp!ʵ�ս��))
            rsTemp.MoveNext
        Next
        If mCurBillType.strNos <> "" Then mCurBillType.strNos = Mid(mCurBillType.strNos, 2)
        
        .Row = .FixedRows: .Col = .ColIndex("��Ŀ")
        Call vsBill_AfterRowColChange(-1, -1, .Row, .Col)
        .Redraw = flexRDDirect
    End With
    txtAllTotal.Text = Format(intSign * dbl�ϼ�, gstrDec)
    Call ReInitPatiInvoice
    txt�˿�ϼ�.Text = Format(GetDelMoney, "0.00")
    
    Screen.MousePointer = 0
    LoadViewBills = True
    Exit Function
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function PrivsValied(ByVal strNo As String) As Boolean
    '�����㵥�ݲ���Ȩ�޼��
    '����:81022
    '����:Ƚ����
    'ʱ��:2014-12-22
    Dim strOper As String, vDate As Date
    
    On Error GoTo errHandle
    If Not ReadBillInfo(1, strNo, -3, strOper, vDate) Then
        MsgBox "����[" & strNo & "]�����ڣ�", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If InStr(mstrPrivs, "���в���Ա") <= 0 And UserInfo.���� <> strOper Then
        MsgBox "��û��""���в���Ա""Ȩ�ޣ����ܶ�" & strOper & "�ĵ��ݽ��в�����", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If Not BillOperCheck(2, strOper, vDate, , strNo, , 1) Then
        Exit Function
    End If
    PrivsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Private Function ReadBills(ByVal intFindType As Integer, ByVal strFindValue As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݵ�ǰ����Ľ��㵥�ݺŻ�Ʊ�ݺ�
    '���:intFindType-0-��������Ų���
    '             1-���շѵ��ݺŲ���
    '             2.�����㵥�Ų���
    '             3.������ķ�Ʊ�Ų���
    '             4.���Һŵ��Ų���
    '     strFindValue-���ҵ�ֵ(0-�������;1-�շѵ��ݺ�;2-���㵥�ݺ�)
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-06-24 15:41:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String, strFeeNos As String, strRegNos As String
    Dim strNos As String, strAllNOs As String
    Dim strSQLIn As String, blnNOMoved As Boolean
    Dim strTmp As String, strCanDelNos As String
    Dim i As Long, j As Integer
    Dim dbl�ϼ� As Currency, arrNo As Variant
    Dim strTemp As String, strҽ����� As String
    Dim blnNotFind As Boolean
    Dim lng����ID As Long, cllInvoiceNoInfor As Collection
    Dim str������� As String
    Dim strInvoiceNO As String
    Dim str���㵥�� As String, bln�ҺŲ��� As Boolean
    Dim strTittle As String
    
     
    On Error GoTo errH
    
    If mbytMode <> EM_RBDTY_�˷� Then Exit Function
    
    Screen.MousePointer = 11
    
    Call ClearFace(False)
    Set cllInvoiceNoInfor = New Collection
    Select Case intFindType
    Case 0  '��������Ų���
        If Not GetBalanceNO(0, strFindValue, str���㵥��, bln�ҺŲ���) Then Exit Function
        strTittle = "�����"
    Case 1  '���շѵ��ݺŲ���
        If Not GetBalanceNO(1, strFindValue, str���㵥��, bln�ҺŲ���) Then Exit Function
        strTittle = "�շѵ���"
    Case 2  '�����㵥�Ų���
        If Not GetBalanceNO(4, strFindValue, str���㵥��, bln�ҺŲ���) Then Exit Function
        strTittle = "���㵥"
    Case 3  '������ķ�Ʊ�Ų���
        If Not GetBalanceNO(2, strFindValue, str���㵥��, bln�ҺŲ���) Then Exit Function
        strTittle = "��Ʊ��"
    Case 4 '���Һŵ��Ų���
        If Not GetBalanceNO(3, strFindValue, str���㵥��, bln�ҺŲ���) Then Exit Function
        strTittle = "�Һŵ���"
    End Select
    If str���㵥�� = "" Then
        Screen.MousePointer = 0
        MsgBox "û���ҵ�" & strTittle & "Ϊ" & strFindValue & "��صĽ����¼��", vbInformation, gstrSysName
        Exit Function
    End If
    
    mblnNOMoved = zlDatabase.NOMoved("���ò����¼", str���㵥��, , 1)
    mCurBillType.str���㵥 = str���㵥��
    mCurBillType.bln�Һ� = bln�ҺŲ���
    
    'bytType-0-����NO������;1-���ݽ���ID������,2-���ݽ������������
    If GetBalanceFeeNos(0, str���㵥��, strFeeNos, strRegNos, mblnNOMoved) = False Then Exit Function
    
    '���ݲ���Ȩ�޼��
    If Not PrivsValied(str���㵥��) Then Screen.MousePointer = 0:  Exit Function
    If bln�ҺŲ��� Then
        mCurBillType.strAllNOs = strRegNos
        strNos = strRegNos
        If CheckDelRegisChargeFeeValied(mCurBillType.strAllNOs, mCurBillType.strNotCanDelNOs, strCanDelNos) = False Then
            Screen.MousePointer = 0:  Exit Function
        End If
    Else
        mCurBillType.strAllNOs = strFeeNos
        strNos = strFeeNos
        
        '����ҽ��ִ�мƼ�.ִ��״̬
        If Upgradeҽ��ִ�мƼ�ִ��״̬(strNos) = False Then
            Screen.MousePointer = 0
            MsgBox "ҽ��ִ�мƼ���������ʧ�ܣ����ܼ�����", vbInformation, gstrSysName
            Exit Function
        End If
        
        If CheckDelChargeIsValied(mCurBillType.strAllNOs, mCurBillType.strNotCanDelNOs, strCanDelNos) = False Then
            Screen.MousePointer = 0:  Exit Function
        End If
    End If
    
    '�˷���ؼ��
    If strCanDelNos <> "" Then strNos = strCanDelNos

    '��ȡ������Ϣ
    '----------------------------------------------------------------------------------
    strSQL = "" & _
    " Select A.����ID,E.����,E.�Ա�,E.����,E.����� as ��ʶ��,E.�ѱ�,E.ҽ�Ƹ��ʽ as ���ʽ,B.����,E.��������" & _
    " From " & IIf(mblnNOMoved, "H", "") & "���ò����¼ A,������Ϣ E,���ս����¼ B, ��Ա�� D" & _
    " Where A.����ID=E.����ID(+) And A.����ID=B.��¼ID(+) And B.����(+)=1 And A.����Ա����=D.����" & _
    "       And (D.վ��='" & gstrNodeNo & "' Or D.վ�� is Null)" & vbNewLine & _
    "       And mod(A.��¼����,10)=1 And A.��¼״̬ IN(1,3) And A.NO=[1] and rownum <2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str���㵥��)
    If rsTemp.EOF Then
        Screen.MousePointer = 0
        MsgBox "û���ҵ������""" & str���㵥�� & """��صĽ����¼��", vbInformation, gstrSysName
        mCurBillType.lng����ID = 0
        Exit Function
    End If
    
    mCurBillType.lngԭ����ID = zlGetFromNOToLastBalanceID(str���㵥��, blnNOMoved, , , True)
    mCurBillType.intInsure = Val(Nvl(rsTemp!����))
    
    Call initInsurePara(mCurBillType.intInsure, lng����ID, mCurBillType.lngԭ����ID)
    If CheckPrivsIsValied = False Then Exit Function    '����Ȩ�޼��
    
    
    txtPatient.Text = Nvl(rsTemp!����)

    lblPati.Caption = "����:" & IIf(txtPatient.Visible, "                       ", rsTemp!����) & _
        "���Ա�:" & Nvl(rsTemp!�Ա�) & _
        "������:" & Nvl(rsTemp!����) & _
        "�������:" & Nvl(rsTemp!��ʶ��) & _
        "���ѱ�:" & Nvl(rsTemp!�ѱ�) & _
        "�����ʽ:" & rsTemp!���ʽ

    With mCurBillType
        .lng����ID = Val(Nvl(rsTemp!����ID))
        .str�Ա� = Nvl(rsTemp!�Ա�)
        .str���� = Nvl(rsTemp!����)
        .str���� = Nvl(rsTemp!����)
    End With

    If Not IsNull(rsTemp!����) Then
        lblPati.ForeColor = vbRed
        txtPatient.ForeColor = vbRed
    Else
        lblPati.ForeColor = &HC00000
        txtPatient.ForeColor = &HC00000
    End If
    
    Call SetPatiColor(txtPatient, Nvl(rsTemp!��������), txtPatient.ForeColor)
    lblPati.ForeColor = txtPatient.ForeColor

    '----------------------------------------------------------------------------------
    '��ȡ���㷽ʽ
    '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�)
    '��������:0-���ݽ���ID����;1-���ݽ�����Ų���,2-���ݵ��ݺ�����ȡ���㷽ʽ
    Set mrsBalance = GetChargeBalance(2, str���㵥��, mblnNOMoved)
    
    
    '��ʼ�����㷽ʽ��ر���
    Call InitBalanceVar: Call LoadBalanceInfor
    
    strҽ����� = ""
    If mCurBillType.bln�Һ� Then
        If GetRegListData(strNos, rsTemp) = False Then Exit Function
        If rsTemp.EOF Then
            Screen.MousePointer = 0
            MsgBox "û���ҵ���Һŵ���Ϊ""" & Split(strNos, ",")(0) & """��صĿ����˺ŵļ�¼��" & _
                vbCrLf & "��Щ�շѼ�¼�����Ѿ��˷ѻ��Ѿ���ȫִ�С�", vbInformation, gstrSysName
            Call ClearFace(False)
            Exit Function
        End If
    Else
        If GetFeeListData(strNos, rsTemp) = False Then Exit Function
        If rsTemp.EOF Then
            Screen.MousePointer = 0
            MsgBox "û���ҵ������""" & Split(strNos, ",")(0) & """��صĿ����˷ѵļ�¼��" & _
                vbCrLf & "��Щ�շѼ�¼�����Ѿ��˷ѻ��Ѿ���ȫִ�С�", vbInformation, gstrSysName
            Call ClearFace(False)
            Exit Function
        End If
    End If
    mCurBillType.strNosOverFlow = ""
    strTmp = ""
    For i = 0 To UBound(Split(strNos, ","))
        strTmp = Replace(Split(strNos, ",")(i), "'", "")
        '����Ƿ��������
        If Not BillOperCheck(IIf(mCurBillType.bln�Һ�, 1, 2), rsTemp!����Ա����, rsTemp!�Ǽ�ʱ��, IIf(mCurBillType.bln�Һ�, "�˺�", "�˷�"), strTmp, , 1, True) Then
            mCurBillType.strNosOverFlow = mCurBillType.strNosOverFlow & " ," & strTmp
        End If
    Next
    If mCurBillType.strNosOverFlow <> "" Then mCurBillType.strNosOverFlow = Mid(mCurBillType.strNosOverFlow, 2)
    
    Call InitBillHead(mCurBillType.bln�Һ�, False)
    stbThis.Panels(2).Text = "��ǰ���㵥��:" & mCurBillType.str���㵥
    With vsBill
        .Redraw = flexRDNone
        .Rows = .FixedRows + rsTemp.RecordCount
        mCurBillType.strNos = ""
        For i = 1 To rsTemp.RecordCount
            .Cell(flexcpData, i, .ColIndex("��Ŀ")) = Nvl(rsTemp!��������)
            .Cell(flexcpData, i, .ColIndex("����ID")) = Nvl(rsTemp!ҽ�����) & "," & Nvl(rsTemp!�շ�ϸĿID)
            If Not mCurBillType.bln�Һ� Then
                If Val(Nvl(rsTemp!ҽ�����)) <> 0 And InStr(strҽ����� & ",", "," & Nvl(rsTemp!ҽ�����) & ",") = 0 Then
                    strҽ����� = strҽ����� & "," & Nvl(rsTemp!ҽ�����)
                End If
            End If
            strTemp = ""
            If Val(Nvl(rsTemp!��������)) <> 0 Then
                rsTemp.MoveNext
                strTemp = "��"
                If rsTemp.EOF Then
                    strTemp = "��"
                ElseIf Val(.Cell(flexcpData, i, .ColIndex("��Ŀ"))) <> Nvl(rsTemp!��������) Then
                    strTemp = "��"
                End If
                rsTemp.MovePrevious
                strTemp = "  " & strTemp & " "
            End If

            .RowData(i) = CLng(rsTemp!���)
            .TextMatrix(i, .ColIndex("ѡ��")) = 0
            .TextMatrix(i, .ColIndex("���ݺ�")) = rsTemp!NO
            .TextMatrix(i, .ColIndex("���")) = rsTemp!�����
            .Cell(flexcpData, i, .ColIndex("���")) = Nvl(rsTemp!�����)
            If mCurBillType.bln�Һ� Then
                .TextMatrix(i, .ColIndex("��Ŀ")) = strTemp & rsTemp!����
                .TextMatrix(i, .ColIndex("����")) = FormatEx(rsTemp!����, 5)
                .Cell(flexcpData, i, .ColIndex("����")) = Val(Nvl(rsTemp!����))
            Else
                .TextMatrix(i, .ColIndex("��Ŀ")) = strTemp & rsTemp!���� & IIf(IsNull(rsTemp!���), "", " " & rsTemp!���)
                .TextMatrix(i, .ColIndex("��Ʒ��")) = strTemp & Nvl(rsTemp!��Ʒ��)
                .TextMatrix(i, .ColIndex("����")) = FormatEx(Nvl(rsTemp!����, 1) * rsTemp!����, 5)
                .Cell(flexcpData, i, .ColIndex("����")) = Nvl(rsTemp!����, 1) * Val(Nvl(rsTemp!����))
            End If
            .TextMatrix(i, .ColIndex("��λ")) = Nvl(rsTemp!���㵥λ)
            .TextMatrix(i, .ColIndex("����")) = Format(Val(Nvl(rsTemp!����)), gstrFeePrecisionFmt)
            .TextMatrix(i, .ColIndex("Ӧ�ս��")) = Format(Val(Nvl(rsTemp!Ӧ�ս��)), gstrDec)
            .TextMatrix(i, .ColIndex("ʵ�ս��")) = Format(Val(Nvl(rsTemp!ʵ�ս��)), gstrDec)
            .TextMatrix(i, .ColIndex("��������")) = Nvl(rsTemp!��������)
            .TextMatrix(i, .ColIndex("ִ�п���")) = Nvl(rsTemp!ִ�п���)
            .TextMatrix(i, .ColIndex("����Ա")) = rsTemp!����Ա����
            If mCurBillType.bln�Һ� Then
                .TextMatrix(i, .ColIndex("ҽ��")) = Nvl(rsTemp!ҽ��)
                .TextMatrix(i, .ColIndex("�Ǽ�ʱ��")) = Format(rsTemp!�Ǽ�ʱ��, "MM-dd HH:mm")
                .TextMatrix(i, .ColIndex("����ʱ��")) = Format(rsTemp!����ʱ��, "MM-dd HH:mm")
                .TextMatrix(i, .ColIndex("ԤԼʱ��")) = Format(rsTemp!ԤԼʱ��, "MM-dd HH:mm")
                .TextMatrix(i, .ColIndex("����ʱ��")) = Format(rsTemp!����ʱ��, "MM-dd HH:mm")
                .TextMatrix(i, .ColIndex("����ʱ��")) = Format(rsTemp!����ʱ��, "MM-dd HH:mm")
                .TextMatrix(i, .ColIndex("����")) = Nvl(rsTemp!����)
                .TextMatrix(i, .ColIndex("����")) = Nvl(rsTemp!����)
                .TextMatrix(i, .ColIndex("����")) = Nvl(rsTemp!����)
            Else
                .TextMatrix(i, .ColIndex("ʱ��")) = Format(rsTemp!�Ǽ�ʱ��, "MM-dd HH:mm")
                .TextMatrix(i, .ColIndex("ҽ��")) = Nvl(rsTemp!ҽ������)
                .TextMatrix(i, .ColIndex("ҽ�����")) = Nvl(rsTemp!ҽ�����)
            End If
            .TextMatrix(i, .ColIndex("ִ�п���ID")) = Nvl(rsTemp!ִ�в���ID)
            
            .TextMatrix(i, .ColIndex("����ID")) = rsTemp!����ID
            .TextMatrix(i, .ColIndex("ԭʼ����")) = Val(Nvl(rsTemp!ԭʼ����))
            .TextMatrix(i, .ColIndex("׼������")) = Val(Nvl(rsTemp!׼������))
            If mCurBillType.bln�Һ� Then
                .Cell(flexcpChecked, i, .ColIndex("ѡ��")) = -1  'ȱʡȫѡ
            ElseIf intFindType = 1 And Nvl(rsTemp!NO) = strFindValue Then
                .Cell(flexcpChecked, i, .ColIndex("ѡ��")) = -1 'ȱʡȫѡ
            End If
            .Cell(flexcpData, i, .ColIndex("ѡ��")) = Val(Nvl(rsTemp!��¼��־))    '�����ж��Ƿ����ʹ�,>1��ʾ������
            If Val(Nvl(rsTemp!��¼��־)) > 1 And InStr(1, mCurBillType.strNosPatiDel & ",", "," & rsTemp!NO & "") = 0 Then mCurBillType.strNosPatiDel = mCurBillType.strNosPatiDel & "," & rsTemp!NO
            
            If InStr(mCurBillType.strNos & ",", "," & rsTemp!NO & ",") = 0 Then
                '�����ָ���
                If mCurBillType.strNos <> "" Then
                    .Select i, .FixedCols, i, .COLS - 1
                    .CellBorder vbBlue, 0, 1, 0, 0, 0, 0
                End If
                mCurBillType.strNos = mCurBillType.strNos & "," & rsTemp!NO
            End If
            dbl�ϼ� = dbl�ϼ� + Val(Nvl(rsTemp!ʵ�ս��))
            rsTemp.MoveNext
        Next
        .Row = .FixedRows: .Col = .ColIndex("��Ŀ")
        Call vsBill_AfterRowColChange(-1, -1, .Row, .Col)
        .Redraw = flexRDDirect
    End With
    
    If strҽ����� <> "" Then
        Set mrs�շѶ��� = zlGet�����շѶ���(Mid(strҽ�����, 2))
    Else
        Set mrs�շѶ��� = Nothing
    End If
    
    If mCurBillType.strNos <> "" Then mCurBillType.strNos = Mid(mCurBillType.strNos, 2)
    
    txtAllTotal.Text = Format(dbl�ϼ�, gstrDec)
    Call LoadSelDelTotal
    Call SetFunCtrlVisible
    
    Screen.MousePointer = 0
    Call ReInitPatiInvoice
    ReadBills = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function CheckPrivsIsValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������Ա�Ƿ�߱������˷ѵ�
    '����:�߱�����true,���򷵻�False
    '����:���˺�
    '����:2014-06-26 16:31:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If Not (mbytMode = EM_RBDTY_�˷� Or mbytMode = EM_RBDTY_�쳣����) Then CheckPrivsIsValied = True: Exit Function
    
    If mCurBillType.intInsure = 0 Then
        Screen.MousePointer = 0
        MsgBox "��ǰ���˷�ҽ�����˽��㵥��,����������˷Ѳ�����", vbInformation, gstrSysName
        Exit Function
    End If
    '�����˷�Ȩ�޼��
    If zlStr.IsHavePrivs(mstrPrivs, "�����˷�") = False Then
        Screen.MousePointer = 0
        MsgBox "��û��Ȩ�޶Խ��н���ԷѲ�����", vbInformation, gstrSysName
        Exit Function
    End If
    CheckPrivsIsValied = True: Exit Function
    Exit Function
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Private Sub cmdBillSel_Click()
    Dim i As Long
    With vsBill
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("���ݺ�")) <> "" And _
               .TextMatrix(i, .ColIndex("���ݺ�")) = .TextMatrix(.Row, .ColIndex("���ݺ�")) And InStr(1, mCurBillType.strNosOverFlow, vsBill.TextMatrix(i, .ColIndex("���ݺ�"))) <= 0 Then
                .TextMatrix(i, .ColIndex("ѡ��")) = -1
            End If
        Next
    End With
    Call LoadSelDelTotal
End Sub

Private Sub cmdCancel_Click()
    If mCurBillType.strNos <> "" And txtNO.Visible Then
        Call ClearFace
        txtNO.SetFocus
    Else
        Unload Me
    End If
End Sub

Private Function FromNOSelect(ByVal strNo As String, ByVal blnSel As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ȫѡ��ȫ�嵥��
    '���:strNO-ָ����NO
    '     blnSel:true��ʾȫѡ,����ȫ��
    '����:���óɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-08-05 11:06:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    On Error GoTo errHandle
    With vsBill
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("���ݺ�")) <> "" _
                And .TextMatrix(i, .ColIndex("���ݺ�")) = strNo Then
                .TextMatrix(i, .ColIndex("ѡ��")) = IIf(blnSel, -1, 0)
            End If
        Next
    End With
    FromNOSelect = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Private Sub cmdClear_Click()
    Dim i As Long, j As Long
    With vsBill
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, .ColIndex("ѡ��")) = 0
        Next
    End With
    Call LoadSelDelTotal
End Sub
 


Private Sub cmdOK_Click()
    If mbytMode = EM_RBDTY_�鿴 Then Unload Me: Exit Sub
    
    If mbytMode = EM_RBDTY_�쳣���� Then
        '�쳣���������˷�
        If ExecuteReDelFee = False Then
            '���¼����쳣����,�Ա��ȡ��ȷ�Ľ�������
            Call LoadViewBills(mstr�������)
            Exit Sub
        End If
        mblnOK = True
        Unload Me: Exit Sub
    End If
    
    '�˺�
    If mCurBillType.bln�Һ� Then
        Call ExecuteDelRegister: Exit Sub
    End If
    '���շѷ���
    Call ExecuteDelChargeFee
End Sub

Private Sub cmdSelAll_Click()
    Dim i As Long
    With vsBill
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("���ݺ�")) <> "" And InStr(1, mCurBillType.strNosOverFlow, vsBill.TextMatrix(i, .ColIndex("���ݺ�"))) <= 0 Then
                .TextMatrix(i, .ColIndex("ѡ��")) = -1
            End If
        Next
    End With
    Call LoadSelDelTotal
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
    If txtNO.Visible And txtNO.Text = "" Then
        txtNO.SetFocus
    Else
        If vsBill.Enabled And vsBill.Visible Then vsBill.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        '###
    ElseIf KeyCode = vbKeyE And Shift = vbCtrlMask Then
        If cmdOK.Visible Then Call cmdOK_Click
    ElseIf KeyCode = vbKeyA And Shift = vbCtrlMask Then
        If cmdSelAll.Visible Then Call cmdSelAll_Click
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        If cmdClear.Visible Then Call cmdClear_Click
    ElseIf KeyCode = vbKeyEscape Or KeyCode = vbKeyX And Shift = vbAltMask Then
        If cmdCancel.Visible Then Call cmdCancel_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("',|~:��;��?��" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub
 

Private Sub Form_Resize()
    Dim staH As Long

    On Error Resume Next

    staH = IIf(stbThis.Visible, stbThis.Height, 0)

    vsBill.Height = Me.ScaleHeight - picCmd.Height - staH - picPati.Height - _
            picMoney.Height - IIf(pic�˷�ժҪ.Visible, pic�˷�ժҪ.Height, 0) - vsBalance.Height

   
    txtNO.Left = Me.ScaleWidth - txtNO.Width - 45
    IDKindNO.Left = txtNO.Left - IDKindNO.Width - 30
    pic��.Left = Me.ScaleWidth - pic��.Width - 45
    lblFormat.Left = IIf(IDKindNO.Visible, IDKindNO.Left, Me.ScaleWidth) _
            - IIf(pic��.Visible, pic��.Width + 45, 0) - lblFormat.Width - 30
    If Me.ScaleWidth - cmdSelAll.Left - cmdCancel.Width > 5500 Then
        cmdCancel.Left = Me.ScaleWidth - cmdSelAll.Left - cmdCancel.Width
    Else
        cmdCancel.Left = 5500
    End If
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 90

    fraInfo_1.Width = Me.ScaleWidth + 300
    LineCmd_1.x2 = Me.ScaleWidth + 300
    With txt�˿�ϼ�
        .Left = Me.ScaleWidth - .Width - 100
        lbl�˿�ϼ�.Left = .Left - lbl�˿�ϼ�.Width - 20
    End With
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim tyTempBillType As tyBillType
    
    If mbytMode <> EM_RBDTY_�鿴 Then zlDatabase.SetPara "�˷Ѻ�������ģʽ", IDKindNO.IDKind, glngSys, mlngModule, InStr(1, mstrPrivs, ";��������;") > 0
    mCurBillType = tyTempBillType
    mbytMode = EM_RBDTY_�鿴
    mstrNo = "": mstrDelTime = "": mblnNOMoved = False  '�鿴ʱ,���ܴ���true
    Call initCardSquareData: Call CloseIDCard
    
    zl_vsGrid_Para_Save mlngModule, vsBill, mstrTittle, IIf(mCurBillType.bln�Һ�, "�Һ���ͷ��Ϣ", "������ͷ��Ϣ")
    Call SaveWinState(Me, App.ProductName, mstrTittle)
    
    If Not mobjFact Is Nothing Then Set mobjFact = Nothing
    If Not mobjInvoice Is Nothing Then Set mobjFact = Nothing
    If Not mrs���㷽ʽ Is Nothing Then Set mrs���㷽ʽ = Nothing
    If Not mrs�շѶ��� Is Nothing Then Set mrs�շѶ��� = Nothing
    If Not mrsBalance Is Nothing Then Set mrsBalance = Nothing
    If Not mrsInfo Is Nothing Then Set mrsInfo = Nothing
    If Not mobjDrugPacker Is Nothing Then Set mobjDrugPacker = Nothing
    If Not mobjDrugMachine Is Nothing Then Set mobjDrugMachine = Nothing
End Sub

Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng�����ID As Long, strOutCardNO As String, strExpand As String
    Dim strOutPatiInforXML As String
    If txtPatient.Locked Or txtPatient.Enabled = False Or txtPatient.Visible = False Then Exit Sub   '

    If objCard.���� Like "IC��*" And objCard.ϵͳ Then
        If mobjICCard Is Nothing Then
            On Error Resume Next
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            If Err <> 0 Then
                Err = 0: On Error GoTo 0
                Exit Sub
            End If
        End If
        txtPatient.Text = mobjICCard.Read_Card()
        If txtPatient.Text <> "" Then Call FindPati(objCard, False, txtPatient.Text)
        Exit Sub
    End If
    lng�����ID = objCard.�ӿ����

    If lng�����ID = 0 Then Exit Sub
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
    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModule, lng�����ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtPatient.Text = strOutCardNO
    If txtPatient.Text <> "" Then Call FindPati(objCard, False, txtPatient.Text)

End Sub

Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    txtPatient.IMEMode = 0
    Set gobjSquare.objCurCard = objCard
    If txtPatient.Text <> "" And Not mblnNotClick Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    zlControl.TxtSelAll txtPatient
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    If txtPatient.Locked Then Exit Sub
    txtPatient.Text = objPatiInfor.����
    If txtPatient.Text <> "" Then Call FindPati(objCard, False, txtPatient.Text)
End Sub
Private Sub IDKindNO_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    zlControl.TxtSelAll txtNO
    If txtNO.Enabled And txtNO.Visible Then txtNO.SetFocus
End Sub

Private Sub pic�˷�ժҪ_Resize()
    Err = 0: On Error Resume Next
    With pic�˷�ժҪ
        txt�˷�ժҪ.Width = .ScaleWidth - txt�˷�ժҪ.Left - 50
    End With
End Sub

Private Sub txtAllTotal_GotFocus()
    zlControl.TxtSelAll txtAllTotal
End Sub

Private Sub txtCurTotal_GotFocus()
    zlControl.TxtSelAll txtCurTotal
End Sub

Private Sub txtNO_GotFocus()
    zlControl.TxtSelAll txtNO
End Sub
Private Function FromNOFind() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����շѵ���Ʊ�Ż�Һŵ����������˷ѵĵ���
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-10-08 16:01:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card
    Dim blnSucces As Boolean
    
    On Error GoTo errHandle
    If Trim(txtNO.Text) = "" Then Exit Function
    
    Set objCard = IDKindNO.GetCurCard
    If objCard Is Nothing Then Exit Function
     
    Select Case objCard.����
    Case "�շѵ���"
        txtNO.Text = GetFullNO(txtNO.Text, 13)
        Call zlControl.TxtSelAll(txtNO)
        '���:intFindType-0-��������Ų���
         '             1-���շѵ��ݺŲ���
         '             2.�����㵥�Ų���
         '             3.������ķ�Ʊ�Ų���
         '             4.���Һŵ��Ų���
        blnSucces = ReadBills(1, txtNO.Text)
    Case "��Ʊ��"
        Call zlControl.TxtSelAll(txtNO)
        blnSucces = ReadBills(3, txtNO.Text)
    Case "�Һŵ���"
        txtNO.Text = GetFullNO(txtNO.Text, 12)
        Call zlControl.TxtSelAll(txtNO)
        blnSucces = ReadBills(4, txtNO.Text)
    Case "���㵥��"
        txtNO.Text = GetFullNO(txtNO.Text, 13)
        Call zlControl.TxtSelAll(txtNO)
        '���:intFindType-0-��������Ų���
         '             1-���շѵ��ݺŲ���
         '             2.�����㵥�Ų���
         '             3.������ķ�Ʊ�Ų���
         blnSucces = ReadBills(2, txtNO.Text)
    End Select
    Screen.MousePointer = 0
    If blnSucces Then vsBill.SetFocus
    FromNOFind = blnSucces
    Exit Function
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 99
        Resume
    End If
End Function

Private Sub txtNO_KeyPress(KeyAscii As Integer)
    Dim strAbc As String, str1 As String, str2 As String
    Dim objCard As Card
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txtNO.Text <> "" Then
            Call FromNOFind
            Exit Sub
        End If
        If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
        Exit Sub
    End If
    Set objCard = IDKindNO.GetCurCard
    If objCard Is Nothing Then Exit Sub
    Call SetNOInputLimit(txtNO, KeyAscii, IIf(objCard.���� = "��Ʊ��", 1, 0))
End Sub

Private Sub txtPatient_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtPatient.Locked Or txtPatient.Enabled = False Or txtPatient.Text <> "" Then Exit Sub
    If IDKind.ActiveFastKey = True Then Exit Sub
End Sub

Private Sub txtPatient_LostFocus()
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(False)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(False)
    IDKind.SetAutoReadCard (False)
End Sub
 
Private Sub txt�˷�ժҪ_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txt�˷�ժҪ, KeyAscii, m�ı�ʽ
End Sub
Private Sub vsBalance_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Row = 1 Then
        With vsBalance
            If Val(.TextMatrix(Row, Col)) <> 0 Then
                .Cell(flexcpForeColor, Row, Col - 1, Row, Col) = vbRed
                .Cell(flexcpFontBold, Row, .Col - 1, Row, .Col) = True
            Else
                .Cell(flexcpForeColor, Row, Col - 1, Row, Col) = Me.ForeColor
                .Cell(flexcpFontBold, Row, .Col - 1, Row, .Col) = False
            End If
        End With
    End If
    Call LoadSelDelTotal
End Sub

Private Sub vsBalance_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If mbytMode = 2 Or mbytMode = 0 Then Cancel = True: Exit Sub
    With vsBalance
        If Col Mod 2 <> 0 Then Cancel = True: Exit Sub
        If Row <> 1 Then Cancel = True: Exit Sub
        If Val(.ColData(Col)) = 0 Then Cancel = True: Exit Sub
        .ColComboList(Col) = " ||" & Val(.Cell(flexcpData, Row, Col))
    End With
End Sub

Private Sub vsBalance_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If vsBalance.MouseCol > 0 Then vsBalance.ToolTipText = vsBalance.ColData(vsBalance.MouseCol)  '��ʾ����ժҪ
End Sub

Private Sub zlSet���ƹ̶���ϵ(ByVal lngRow As Long, ByVal Col As Long, Optional lngNotCheckRow As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������շѹ�ϵ,�Զ����й�ѡ
    '����:���˺�
    '����:2014-10-11 11:26:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, varTemp As Variant, bln�̶� As Boolean, i As Long, j As Long
    If vsBill.Cell(flexcpData, lngRow, vsBill.ColIndex("����ID")) = "" Then Exit Sub
    If mrs�շѶ��� Is Nothing Then Exit Sub
     
     '����:33634:����ǹ̶�����Ŀ(�����շѹ�ϵ):��ҽ�������Ĳ��ж�
    varData = Split(vsBill.Cell(flexcpData, lngRow, vsBill.ColIndex("����ID")) & ",", ",")
    If Val(varData(0)) = 0 Then Exit Sub

    mrs�շѶ���.Filter = "ҽ�����=" & Val(varData(0)) & " And �շ�ϸĿID=" & Val(varData(1))
    If Not mrs�շѶ���.EOF Then
        bln�̶� = Val(Nvl(mrs�շѶ���!���ж���)) = 1
    Else
        bln�̶� = False
    End If
    mrs�շѶ���.Filter = 0
    If bln�̶� = False Then Exit Sub
    With vsBill
        For i = 1 To .Rows - 1
            If i <> lngRow And lngNotCheckRow <> i Then
                varTemp = Split(vsBill.Cell(flexcpData, i, .ColIndex("����ID")) & ",", ",")
                If varData(0) = varTemp(0) Then    '����ͬ��ҽ�����
                     mrs�շѶ���.Filter = "ҽ�����=" & Val(varTemp(0)) & " And �շ�ϸĿID=" & Val(varTemp(1))
                    If Not mrs�շѶ���.EOF Then
                        bln�̶� = Val(Nvl(mrs�շѶ���!���ж���)) = 1
                    Else
                        bln�̶� = False
                    End If
                    If bln�̶� Then
                        .Cell(flexcpChecked, i, .ColIndex("ѡ��")) = .Cell(flexcpChecked, lngRow, .ColIndex("ѡ��"))
                        .TextMatrix(i, .ColIndex("ѡ��")) = .TextMatrix(lngRow, .ColIndex("ѡ��"))
                        '���������,��Ҫ�������
                        If Val(.Cell(flexcpData, i, .ColIndex("��Ŀ"))) = 0 Then  '�϶�Ϊ����,���,��Ҫ�Ҵ�������
                            For j = i + 1 To vsBill.Rows - 1
                                If .RowData(i) <> Val(.Cell(flexcpData, j, .ColIndex("��Ŀ"))) Then Exit For
                                .Cell(flexcpChecked, j, .ColIndex("ѡ��")) = .Cell(flexcpChecked, i, .ColIndex("ѡ��"))
                                .TextMatrix(j, .ColIndex("ѡ��")) = .TextMatrix(i, .ColIndex("ѡ��"))
                            Next
                        End If
                    End If
                 End If
            End If
        Next
    End With
End Sub

Private Sub vsBill_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long, varData As Variant, bln�̶� As Boolean
    Dim varTemp As Variant, j As Long
    Dim strNo As String
    With vsBill
        If Col <> .ColIndex("ѡ��") Then Exit Sub
        stbThis.Panels(2).Text = ""
        If mCurBillType.bln�Һ� Then
            '�����ݺ�ѡ��
            For i = .FixedRows To .Rows - 1
                If .TextMatrix(i, .ColIndex("���ݺ�")) <> "" And _
                   .TextMatrix(i, .ColIndex("���ݺ�")) = .TextMatrix(.Row, .ColIndex("���ݺ�")) And InStr(1, mCurBillType.strNosOverFlow, vsBill.TextMatrix(i, .ColIndex("���ݺ�"))) <= 0 Then
                    .TextMatrix(i, .ColIndex("ѡ��")) = .TextMatrix(Row, .ColIndex("ѡ��"))
                End If
            Next
        Else
            If Val(.Cell(flexcpData, Row, .ColIndex("��Ŀ"))) = 0 Then
                For i = Row + 1 To .Rows - 1
                     If Val(.RowData(Row)) <> Val(.Cell(flexcpData, i, .ColIndex("��Ŀ"))) Then Exit For
                    .TextMatrix(i, .ColIndex("ѡ��")) = .TextMatrix(Row, .ColIndex("ѡ��"))
                Next
                Call zlSet���ƹ̶���ϵ(Row, Col)
            Else
                Call zlSet���ƹ̶���ϵ(Row, Col)
                '��Ҫ��������Ƿ��Ѿ���
                For i = Row - 1 To 1 Step -1
                    If Val(.RowData(i)) = Val(.Cell(flexcpData, Row, .ColIndex("��Ŀ"))) Then
                        If .TextMatrix(i, .ColIndex("ѡ��")) <> 0 Then
                             .TextMatrix(i, .ColIndex("ѡ��")) = .TextMatrix(Row, .ColIndex("ѡ��"))
                        End If
                        Call zlSet���ƹ̶���ϵ(i, Col, Row)
                         Exit For
                    End If
                Next
            End If
        End If
        Call LoadSelDelTotal
    End With
End Sub

Private Sub vsBill_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim dbl�ϼ� As Currency, i As Long
    If NewRow = OldRow Then Exit Sub
    With vsBill
        If Trim(.TextMatrix(NewRow, .ColIndex("���ݺ�"))) = "" Then
            txtCurTotal.Text = Format(dbl�ϼ�, gstrDec)
            Exit Sub
        End If
        For i = NewRow - 1 To .FixedRows Step -1
            If .TextMatrix(i, .ColIndex("���ݺ�")) <> .TextMatrix(NewRow, .ColIndex("���ݺ�")) Then Exit For
            dbl�ϼ� = dbl�ϼ� + Val(.TextMatrix(i, .ColIndex("ʵ�ս��")))
        Next
        For i = NewRow To .Rows - 1
            If .TextMatrix(i, .ColIndex("���ݺ�")) <> .TextMatrix(NewRow, .ColIndex("���ݺ�")) Then Exit For
            dbl�ϼ� = dbl�ϼ� + Val(.TextMatrix(i, .ColIndex("ʵ�ս��")))
        Next
        txtCurTotal.Text = Format(dbl�ϼ�, gstrDec)
    End With
End Sub

Private Sub vsBill_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    If mCurBillType.bln�Һ� Then
        mTyColHead.strRegColHead = zl_vsGrid_GetCols_Property(vsBill)
    Else
        mTyColHead.strFeeColHead = zl_vsGrid_GetCols_Property(vsBill)
    End If
End Sub

Private Sub vsBill_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsBill
        If .Col <> .ColIndex("ѡ��") Then Cancel = True: Exit Sub
        If .ColIndex("���ݺ�") < 0 Then Cancel = True: Exit Sub
        If Trim(.TextMatrix(Row, .ColIndex("���ݺ�"))) = "" Then Cancel = True: Exit Sub
    End With
End Sub

Private Sub vsBill_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = vsBill.ColIndex("ѡ��") Then Cancel = True
End Sub

Private Sub GetBillRow(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݵ��ݵ�ָ���У���ȡ���ݵĿ�ʼ�кͽ�����
    '���:lngRow-��ǰ��
    '����:lngBegin-���ݵĿ�ʼ��
    '     lngEnd-���ݵĽ�����
    '����:���˺�
    '����:2014-07-03 17:39:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    lngBegin = lngRow: lngEnd = lngRow
    With vsBill
        For i = lngRow - 1 To .FixedRows Step -1
            If .TextMatrix(i, .ColIndex("���ݺ�")) <> .TextMatrix(lngRow, .ColIndex("���ݺ�")) Then Exit For
            lngBegin = i
        Next
        For i = lngRow To .Rows - 1
            If .TextMatrix(i, .ColIndex("���ݺ�")) <> .TextMatrix(lngRow, .ColIndex("���ݺ�")) Then Exit For
            lngEnd = i
        Next
    End With
End Sub

Private Sub vsBill_CellChanged(ByVal Row As Long, ByVal Col As Long)
    With vsBill
        If .ColIndex("���ݺ�") < 0 Then Exit Sub
        '�����޶������
        If .TextMatrix(Row, .ColIndex("���ݺ�")) <> "" _
            And InStr(1, mCurBillType.strNosOverFlow, .TextMatrix(Row, .ColIndex("���ݺ�"))) > 0 Then
             .TextMatrix(Row, .ColIndex("ѡ��")) = 0
        End If
    End With
End Sub

Private Sub vsBill_DrawCell(ByVal hDC As Long, ByVal Row As Long, _
    ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, _
    ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����
    '����:���˺�
    '����:2014-07-03 17:41:52
    '˵����1.OwnerDrawҪ����ΪOver(������Ԫ��������)
    '      2.Cell��GridLine�������������ڶ��Ǵӵ�1���߿�ʼ
    '      3.Cell��Border�������Ǵӵ�2���߿�ʼ,�����Ǵӵ�1���߿�ʼ
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT

    With vsBill
        '����һ����ҩ������еı��߼�����
        lngLeft = .ColIndex("���ݺ�"): lngRight = .ColIndex("���ݺ�")
        If Not Between(Col, lngLeft, lngRight) Then Exit Sub
        Call GetBillRow(Row, lngBegin, lngEnd)
        If lngBegin = lngEnd Then Exit Sub

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
            SetBkColor hDC, OS.SysColor2RGB(.BackColorSel)
        Else
            SetBkColor hDC, OS.SysColor2RGB(.BackColor)
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Done = True
    End With
End Sub

Private Sub vsBill_KeyPress(KeyAscii As Integer)
    With vsBill
        Select Case KeyAscii
        Case 32 '�ո�
            If .ColHidden(.ColIndex("ѡ��")) Then Exit Sub
            KeyAscii = 0
            If Trim(.TextMatrix(.Row, .ColIndex("���ݺ�"))) = "" Then Exit Sub
            
            If .TextMatrix(.Row, .ColIndex("ѡ��")) = 0 _
                And InStr(1, mCurBillType.strNosOverFlow, .TextMatrix(.Row, .ColIndex("���ݺ�"))) <= 0 Then
                 .TextMatrix(.Row, .ColIndex("ѡ��")) = -1
            Else
                 .TextMatrix(.Row, .ColIndex("ѡ��")) = 0
            End If
            Call LoadSelDelTotal
            
            '87675,��Ҫ�ֶ�����AfterEdit�¼�
            Call vsBill_AfterEdit(.Row, .ColIndex("ѡ��"))
        Case 13 '�س�
            KeyAscii = 0
            If .Row + 1 <= .Rows - 1 Then
               .Row = .Row + 1: .ShowCell .Row, .Col
            End If
        End Select
    End With
End Sub
Private Sub vsBill_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsBill
        Select Case Col
        Case .ColIndex("ѡ��")
        Case Else
             Cancel = True
        End Select
    End With
End Sub

Private Function CheckDelChargeIsValied(ByVal strNos As String, _
    ByRef strNotCanDelNOs As String, _
    ByRef strCanDelNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����˷ѵ����Ƿ�Ϸ�
    '���:strNOs-��Ҫ���ĵ��ݺ�(����ö��ŷ���)
    '����:strNotCanDelNOs-�����˵ĵ���(�Ѿ�ִ�м������˵ĵ���)
    '     strCanDelNos-���˵ĵ���
    '����:�Ϸ�����true,���򷵻�False
    '����:���˺�
    '����:2014-10-11 11:28:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim arrNo As Variant, i As Long, intTmp As Integer
    Dim strInfo As String, strFlagPrintInfor As String
    Dim blnFlagPrint As Boolean, strNo As String, strCurNO As String
    Dim blnHaveExe As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle

    '����:54728
    If Not mbytMode = EM_RBDTY_�˷� Then CheckDelChargeIsValied = True: Exit Function   '�˷�ʱ�ж�

    arrNo = Split(strNos, ","): strNotCanDelNOs = ""
    strCanDelNos = ""   '��¼�����˵ĵ��ݺ�
    strInfo = ""        '�������ʾ��Ϣ
    strFlagPrintInfor = ""
    For i = 0 To UBound(arrNo)
        strCurNO = Replace(arrNo(i), "'", "")
        If strNo = "" Then strNo = strCurNO

        blnHaveExe = False: blnFlagPrint = False
        intTmp = BillCanDelete(strCurNO, 1, blnHaveExe, , blnFlagPrint)
        If intTmp <> 0 Then
            strNotCanDelNOs = strNotCanDelNOs & "," & strCurNO  '�����˵ĵ���
            Select Case intTmp
                Case 1 '�õ��ݲ�����
                    strInfo = strInfo & "ָ���ĵ��ݲ����ڣ�" & vbCrLf
                    Exit For
                Case 2 '�Ѿ�ȫ����ȫִ��(�շѲ������˷��Զ���ҩ)
                    strInfo = strInfo & "[" & strCurNO & "]�е���Ŀ�Ѿ�ȫ����ȫִ��,�����˷�!" & vbCrLf
                Case 3 'δ��ȫִ�в���ʣ������Ϊ0
                    strInfo = strInfo & "[" & strCurNO & "]��δ��ȫִ�е���Ŀʣ������Ϊ��,û�п��˷��ã�" & vbCrLf
            End Select

        ElseIf blnHaveExe Then
            If gbln�˷�����ģʽ Then
                'δ�����δ��˵ĵ��ݲ����˷�
                Set rsTemp = GetApply(strCurNO, 1)
                rsTemp.Filter = "״̬<>2"
                If rsTemp.RecordCount = 0 Then
                    strInfo = strInfo & "[" & strCurNO & "]δ�����˷����뼰��ˣ����ܽ����˷ѣ�" & vbCrLf
                    strNotCanDelNOs = strNotCanDelNOs & "," & strCurNO  '�����˵ĵ���
                ElseIf IsNull(rsTemp!�����) Then
                    strInfo = strInfo & "[" & strCurNO & "]δ�����˷���ˣ����ܽ����˷ѣ�" & vbCrLf
                    strNotCanDelNOs = strNotCanDelNOs & "," & strCurNO  '�����˵ĵ���
                Else
                    strInfo = strInfo & "[" & strCurNO & "]�д�����ִ�е���Ŀ���˵��ݽ�ִ�е��ǲ����˷ѡ�" & vbCrLf
                    strCanDelNos = strCanDelNos & "," & strCurNO
                End If
            Else
                strInfo = strInfo & "[" & strCurNO & "]�д�����ִ�е���Ŀ���˵��ݽ�ִ�е��ǲ����˷ѡ�" & vbCrLf
                strCanDelNos = strCanDelNos & "," & strCurNO
            End If
        ElseIf gbln�˷�����ģʽ Then
            'δ�����δ��˵ĵ��ݲ����˷�
            Set rsTemp = GetApply(strCurNO, 1)
            rsTemp.Filter = "״̬<>2"
            If rsTemp.RecordCount = 0 Then
                strInfo = strInfo & "[" & strCurNO & "]δ�����˷����뼰��ˣ����ܽ����˷ѣ�" & vbCrLf
                strNotCanDelNOs = strNotCanDelNOs & "," & strCurNO  '�����˵ĵ���
            ElseIf IsNull(rsTemp!�����) Then
                strInfo = strInfo & "[" & strCurNO & "]δ�����˷���ˣ����ܽ����˷ѣ�" & vbCrLf
                strNotCanDelNOs = strNotCanDelNOs & "," & strCurNO  '�����˵ĵ���
            Else
                strCanDelNos = strCanDelNos & "," & strCurNO
            End If
        Else
            strCanDelNos = strCanDelNos & "," & strCurNO
        End If

        If blnFlagPrint Then
            '����Ӧ�������Ƿ��Ѵ�ӡ(����ҽ���еĲɼ���ʽ��ִ��)
            strFlagPrintInfor = strFlagPrintInfor & "[" & strCurNO & "]����ҽ���������Ѵ�ӡ��" & vbCrLf
        End If
    Next

    If strNotCanDelNOs <> "" Then strNotCanDelNOs = Mid(strNotCanDelNOs, 2)
    strCanDelNos = Mid(strCanDelNos, 2)

    If strFlagPrintInfor <> "" Then
        If MsgBox("ע��:" & vbCrLf & strFlagPrintInfor & vbCrLf & " �Ƿ�����˷ѣ�", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Screen.MousePointer = 0
            Exit Function
        End If
    End If

    If strCanDelNos = "" Then
        '���ŵ�����Ϊ�Ǽ�����һ��,��Ȼ��һ��ת����û��ת��
        '�Ƿ���ת������ݱ���
        If zlDatabase.NOMoved("������ü�¼", strNo, , "1") Then
            If Not ReturnMovedExes(strNo, 1, Me.Caption) Then
                Screen.MousePointer = 0
                Exit Function
            End If
        End If
        Screen.MousePointer = 0
        MsgBox strInfo, vbInformation, gstrSysName
        Exit Function
    End If

    If strInfo <> "" Then
        MsgBox strInfo, vbInformation, gstrSysName
    End If
    strNos = strCanDelNos

    '���ŵ�����Ϊ�Ǽ�����һ��,��Ȼ��һ��ת����û��ת��
    '�Ƿ���ת������ݱ���
    If zlDatabase.NOMoved("������ü�¼", strNo, , "1") Then
        If Not ReturnMovedExes(strNo, 1, Me.Caption) Then
            Screen.MousePointer = 0
            Exit Function
        End If
    End If
    CheckDelChargeIsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Private Function CheckDelRegisChargeFeeValied(ByVal strNos As String, _
    ByRef strNotCanDelNOs As String, _
    ByRef strCanDelNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���Һŵ����˺��Ƿ�Ϸ�
    '���:strNOs-��Ҫ���ĵ��ݺ�(����ö��ŷ���)
    '����:strNotCanDelNOs-�����˵ĵ���(�Ѿ�ִ�м������˵ĵ���)
    '     strCanDelNos-���˵ĵ���
    '����:�Ϸ�����true,���򷵻�False
    '����:���˺�
    '����:2014-10-11 11:28:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim arrNo As Variant, i As Long, intTmp As Integer
    Dim strInfo As String, strFlagPrintInfor As String
    Dim blnFlagPrint As Boolean, strNo As String, strCurNO As String
    Dim blnHaveExe As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle

    '����:54728
    If Not mbytMode = EM_RBDTY_�˷� Then CheckDelRegisChargeFeeValied = True: Exit Function   '�˷�ʱ�ж�

    arrNo = Split(strNos, ","): strNotCanDelNOs = ""
    strCanDelNos = ""   '��¼�����˵ĵ��ݺ�
    strInfo = ""        '�������ʾ��Ϣ
    strFlagPrintInfor = ""
    For i = 0 To UBound(arrNo)
        strCurNO = Replace(arrNo(i), "'", "")
        If strNo = "" Then strNo = strCurNO

        blnHaveExe = False: blnFlagPrint = False
        intTmp = BillCanDelete(strCurNO, 4, blnHaveExe, , blnFlagPrint)
        If intTmp <> 0 Then
            strNotCanDelNOs = strNotCanDelNOs & "," & strCurNO  '�����˵ĵ���
            Select Case intTmp
                Case 1 '�õ��ݲ�����
                    strInfo = strInfo & "ָ���ĵ��ݲ����ڣ�" & vbCrLf
                    Exit For
                Case 2 '�Ѿ�ȫ����ȫִ��(�շѲ������˷��Զ���ҩ)
                    strInfo = strInfo & "[" & strCurNO & "]�е���Ŀ�Ѿ�ȫ����ȫִ��,�����˺�!" & vbCrLf
                Case 3 'δ��ȫִ�в���ʣ������Ϊ0
                    strInfo = strInfo & "[" & strCurNO & "]��δ��ȫִ�е���Ŀʣ������Ϊ��,�����˺ţ�" & vbCrLf
            End Select

        ElseIf blnHaveExe Then
            '������ִ����Ŀ
            strInfo = strInfo & "[" & strCurNO & "]�д�����ִ�е���Ŀ���˵��ݽ�ִ�е��ǲ����˷ѡ�" & vbCrLf
            strCanDelNos = strCanDelNos & "," & strCurNO
        Else
            strCanDelNos = strCanDelNos & "," & strCurNO
        End If
        
        If blnFlagPrint Then
            '����Ӧ�������Ƿ��Ѵ�ӡ(����ҽ���еĲɼ���ʽ��ִ��)
            strFlagPrintInfor = strFlagPrintInfor & "[" & strCurNO & "]����ҽ���������Ѵ�ӡ��" & vbCrLf
        End If
    Next

    If strNotCanDelNOs <> "" Then strNotCanDelNOs = Mid(strNotCanDelNOs, 2)
    strCanDelNos = Mid(strCanDelNos, 2)

    If strFlagPrintInfor <> "" Then
        If MsgBox("ע��:" & vbCrLf & strFlagPrintInfor & vbCrLf & " �Ƿ�����˺ţ�", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Screen.MousePointer = 0
            Exit Function
        End If
    End If

    If strCanDelNos = "" Then
        '���ŵ�����Ϊ�Ǽ�����һ��,��Ȼ��һ��ת����û��ת��
        '�Ƿ���ת������ݱ���
        If zlDatabase.NOMoved("������ü�¼", strNo, , "4") Then
            If Not ReturnMovedExes(strNo, 4, Me.Caption) Then
                Screen.MousePointer = 0
                Exit Function
            End If
        End If
        Screen.MousePointer = 0
        MsgBox strInfo, vbInformation, gstrSysName
        Exit Function
    End If

    If strInfo <> "" Then
        MsgBox strInfo, vbInformation, gstrSysName
    End If
    strNos = strCanDelNos

    '���ŵ�����Ϊ�Ǽ�����һ��,��Ȼ��һ��ת����û��ת��
    '�Ƿ���ת������ݱ���
    If zlDatabase.NOMoved("������ü�¼", strNo, , "4") Then
        If Not ReturnMovedExes(strNo, 4, Me.Caption) Then
            Screen.MousePointer = 0
            Exit Function
        End If
    End If
    CheckDelRegisChargeFeeValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Private Sub InitBalanceVar()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������
    '����:���˺�
    '����:2014-07-04 10:02:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str���㷽ʽ As String
    
    If mrsBalance Is Nothing Then Exit Sub
    If mrsBalance.State <> 1 Then Exit Sub
    
    mrsBalance.Filter = "����<>2 And ����<>1"
    '       �ֶ�:����,��¼����,���㷽ʽ,ժҪ,�����ID,���������,���ƿ�,���㿨���,�������,����,������ˮ��, ����˵��,�������,У�Ա�־,ҽ��,���ѿ�id
    '            �Ƿ�����,�Ƿ�ȫ��,�Ƿ�����,��Ԥ��
    '       ����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�)
    str���㷽ʽ = ""
    mrsBalance.Sort = "����,��������"
    With mrsBalance
        Do While Not .EOF
            If InStr(str���㷽ʽ & ",", "," & Nvl(!���㷽ʽ) & ",") = 0 Then
                str���㷽ʽ = str���㷽ʽ & "," & Nvl(!���㷽ʽ)
            End If
            If Val(Nvl(!����)) = 3 Or Val(Nvl(!����)) = 4 Then mCurBillType.bln���ڿ����� = True
            .MoveNext
        Loop
    End With
    If str���㷽ʽ <> "" Then str���㷽ʽ = Mid(str���㷽ʽ, 2)
    mCurBillType.str���㷽ʽ = str���㷽ʽ
    
    '4-һ��ͨ(��)
    mrsBalance.Filter = "����=4"
    mCurBillType.blnExistOnCard = mrsBalance.EOF = False
    
    '3.һ��ͨ
    mrsBalance.Filter = "����=3 And  �Ƿ�ȫ��=1 and �Ƿ�����=0"
    mCurBillType.blnExistThreeAllDel = mrsBalance.EOF = False
    mrsBalance.Filter = 0
End Sub

Private Function ExecuteClinicDelSwap(ByVal lng����ID As Long, ByVal intInsure As Integer, _
    ByVal lng����ID As Long, ByVal lngԭ����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ��ҽ���˷ѽ���
    '���:lng����ID-����ID
    '     intInsure-����
    '     lng����ID-����ID
    '     lngԭ����ID-ԭʼ����ID
    '����:
    '����:ҽ���˷ѽ��׳ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-31 23:38:11
    '˵��:
    '   ���ýӿ�ǰ,�����ȴ�����,��ɺ�,���Զ��ύ����;ʧ��ʱ,���������
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strAdvance As String, strAllBalance As String, strSQL As String
    Dim varData As Variant, varTemp As Variant, i As Long
    
    On Error GoTo errHandle
    
    If intInsure = 0 Then ExecuteClinicDelSwap = True: gcnOracle.CommitTrans: Exit Function
    strAllBalance = GetYBOldBalance(lng����ID, intInsure, lngԭ����ID)
    
    strAdvance = ""
    If MCPAR.����������� Then
        strAdvance = lng����ID
        'ClinicDelSwap (ҽ���˷ѽ���)
        '������  ��������    ��/��   ԭ����˵��  �ֵ���˵��
        'lngStlID    long    IN  ��Ҫ�˷ѵķ��ü�¼�Ľ���ID(ԭ����ID)
        'bln�˷� Boolean IN  �������˷ѽ��׻��Ǹķѽ����ڵ��ñ��ӿ�
        'intInsure   Intger  In  ����
        'strAdvance  String  In  NULL    ����ID:���Ӵ������ID
        'ҽ�����Ը��ݳ���ID������ȡ��
        '        Out �˷ѽ��㣺���㷽ʽ1|���||���㷽ʽ2|���...
        '    Boolean ��������    True:���óɹ�,False:����ʧ��
        '����ID|��������־|
        strAdvance = lng����ID & "|1"
        If Not gclsInsure.ClinicDelSwap(lngԭ����ID, , intInsure, strAdvance) Then
            gcnOracle.RollbackTrans
            Exit Function
        End If
        If strAdvance = CStr(lng����ID) & "|1" Then strAdvance = ""
    Else
        strAdvance = strAllBalance
        varData = Split(strAdvance, "||")
        strAdvance = ""
        For i = 0 To UBound(varData)
            varTemp = Split(varData(i) & "|||", "|")
            strAdvance = strAdvance & "||" & varTemp(0) & "|" & -1 * Val(varTemp(1))
        Next
        If strAdvance <> "" Then strAdvance = Mid(strAdvance, 3)
    End If
    
    If MCPAR.����������� Then
        If Not zlInsureCheck(strAllBalance, strAdvance) Or strAdvance = "" Then
            gcnOracle.CommitTrans
            If MCPAR.����������� Then Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, True, intInsure)
            ExecuteClinicDelSwap = True: Exit Function
        End If
        gcnOracle.CommitTrans: gcnOracle.BeginTrans
    End If
    
    '�˷Ѻ��շѲ�һ��ʱ,��ҪЧ��
    'Zl_���ò������_Modify
    strSQL = "Zl_���ò������_Modify("
    '  ��������_In   Number,
    '  --   0-��ͨ���㷽ʽ:
    '  --     ���㷽ʽ_IN:��������,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ||.." ;Ҳ�������.
    '  --   1.����������:
    '  --     �ٽ��㷽ʽ_IN:ֻ�ܴ���һ�����㷽ʽ,���������һЩ������Ϣ,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ"
    '  --     �ڿ����ID_IN,����_IN,������ˮ��_IN,����˵��_In:��Ҫ����
    '  --   2-ҽ������(�������ҽ���Ľ���,��Ҫ��ɾ��ԭҽ������,���´���ĸ���)
    '  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:���㷽ʽ|������||.."
    strSQL = strSQL & "" & 2 & ","
    '  ����id_In     In ���ò����¼.����id%Type,
    strSQL = strSQL & "" & lng����ID & ","
    '  ���㷽ʽ_In   Varchar2,
    strSQL = strSQL & "'" & strAdvance & "')"
    '  �����id_In   ����Ԥ����¼.�����id%Type := Null,
    '  ����_In       ����Ԥ����¼.����%Type := Null,
    '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
    '  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
    '  �����_In   ������ü�¼.ʵ�ս��%Type := Null,
    '  ��ɽ���_In Number:=0
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    gcnOracle.CommitTrans
    Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, True, intInsure)
    ExecuteClinicDelSwap = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    Call ErrCenter
    Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, False, mCurBillType.intInsure)
End Function
Private Function isChargeFeeValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�˷��Ƿ�Ϸ�
    '����:�Ϸ�����true,���򷵻�False
    '����:���˺�
    '����:2014-10-11 11:34:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim arrNo As Variant, strYPNos As String, blnҩƷ As Boolean, blnSel As Boolean
    Dim i As Long, strDelNOs As String, strNo As String, strOperatorName As String
    Dim varTemp As Variant, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    If Not CheckTextLength("�˷�ժҪ", txt�˷�ժҪ) Then Exit Function
    '��������Ƿ���ȷ
    If mCurBillType.strNos = "" Then
        MsgBox "����ȷ����Ҫ�˷ѵ������շѵ��ݡ�", vbInformation, gstrSysName
        If txtNO.Visible Then txtNO.SetFocus: Exit Function
    End If
    
    '��鱾�ν��㵥�����Ƿ�����˷��쳣���ݣ������ڣ�����������˷�
    If CheckIsExistDelErrBill(mCurBillType.str���㵥, strOperatorName) Then
        MsgBox "ע�⣺" & vbCrLf & _
            "    ���ν����д����쳣�Ľ����¼�����ȶ�����������˷ѣ�" & _
            IIf(strOperatorName <> UserInfo.����, vbCrLf & "    ��ʾ���쳣�����ǲ���Ա��" & strOperatorName & "����ȡ�ġ�", ""), vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    arrNo = Split(mCurBillType.strNos, ",")
    strYPNos = "": strDelNOs = ""
    blnҩƷ = False: blnSel = False
    With vsBill
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("ѡ��"))) <> 0 Then
                blnSel = True
                strNo = .TextMatrix(i, .ColIndex("���ݺ�"))
                If InStr(strDelNOs & ",", "," & strNo & ",") = 0 Then
                    strDelNOs = strDelNOs & "," & strNo
                End If
                
                If .ColIndex("���") <> -1 And blnҩƷ = False Then     '47400
                    If .TextMatrix(i, .ColIndex("���")) Like "*��*ҩ*" _
                        Or .TextMatrix(i, .ColIndex("���")) Like "*��*ҩ*" _
                        Or .TextMatrix(i, .ColIndex("���")) Like "*����*" Then
                        If InStr(strYPNos & ",", "," & strNo & ",") = 0 Then
                            strYPNos = strYPNos & "," & strNo
                        End If
                        blnҩƷ = True
                    End If
                End If
            End If
        Next
    End With
    If strDelNOs <> "" Then strDelNOs = Mid(strDelNOs, 2)
    
    If strDelNOs <> "" And gbln�˷�����ģʽ And Not mCurBillType.bln�Һ� Then
        Set rsTemp = GetApply(strDelNOs, 1)
        varTemp = Split(strDelNOs, ",")
        For i = 0 To UBound(varTemp)
            strNo = varTemp(i)
            rsTemp.Filter = "NO='" & strNo & "' And ״̬<>2"
            If rsTemp.RecordCount = 0 Then
                Screen.MousePointer = 0
                MsgBox "���ȶ��շѵ���:" & strNo & " �����˷����룡", vbInformation, gstrSysName
                Exit Function
            End If
            If IsNull(rsTemp!�����) Then
                Screen.MousePointer = 0
                MsgBox "����:" & strNo & " δ�����˷���ˣ����ܽ����˷ѣ�", vbInformation, gstrSysName
                Exit Function
            End If
        Next
    End If
    If blnSel = False Then
        MsgBox "���ڵ���������ѡ��һ��Ҫ�˷ѵ���Ŀ��", vbInformation, gstrSysName
        vsBill.SetFocus: Exit Function
    End If
    
    If blnҩƷ And Not mCurBillType.bln�Һ� Then
        If strYPNos <> "" Then strYPNos = Mid(strYPNos, 2)
        If zlCheckDrugIsPutDrug(strYPNos) = False Then Exit Function
    End If
    
    'ҽ�����
    If mCurBillType.intInsure = 0 Then
        MsgBox "��ǰ���㲻��ҽ�����˽���,���������" & IIf(Not mCurBillType.bln�Һ�, "�˷�", "�˺�") & "����!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If gclsInsure.CheckInsureValid(mCurBillType.intInsure) = False Then Exit Function
    If Not mCurBillType.bln�Һ� Then
        If zlCheckIsMzToZY(strDelNOs, 1) Then
              MsgBox "ע��:" & vbCrLf & _
                "    �õ����Ѿ����������תסԺ���� " & vbCrLf & _
                "    ���Ѿ�������������תסԺ����,�������˷�", vbInformation + vbOKOnly, gstrSysName
              Exit Function
        End If
        
        If MCPAR.����������� = False Then '112843
            MsgBox "��ǰҽ����֧������������ϣ����ܽ����˷ѣ�", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
    
    '���������㷽ʽ��Ч�Լ��
    If ThreeBalanceCheck(mobjPayCards, mrsBalance, mcllForceDelToCash, mstr�ų����㷽ʽ) = False Then Exit Function
    
    isChargeFeeValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function ThreeBalanceCheck(objCards As Cards, ByVal rsBalance As ADODB.Recordset, _
    ByRef cllForceDelToCash As Collection, ByRef str�ų����㷽ʽ As String) As Boolean
    '���������㷽ʽ��Ч�Լ��
    '��Σ�
    '   objCards ������������Ч��֧����ʽ
    '   rsBalance ������Ϣ
    '���Σ�
    '   cllForceDelToCash ǿ��������Ϣ��Array(����Ա,���������,���㷽ʽ)
    '   str�ų����㷽ʽ �ų����㷽ʽ,����ö��ŷָ�
    '���أ����ͨ��������True�����򣬷���False
    '105432
    Dim objCard As Card
    Dim cllFeeBalance As New Collection, i As Integer
    Dim blnFind As Boolean, blnQuestion As Boolean
    Dim str����Ա As String, strKey As String
    Dim dblMoney  As Double
    Dim j As Integer, lngCount As Long
    Dim varData As Variant
    
    On Error GoTo errHandler
    Set cllForceDelToCash = New Collection
    str�ų����㷽ʽ = ""
    If rsBalance Is Nothing Then ThreeBalanceCheck = True: Exit Function
    
    '���ͣ�0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
    rsBalance.Filter = "����= 3"
    'ȥ��
    With rsBalance
        Do While Not .EOF
            strKey = "_" & Val(Nvl(!�����ID))
            If CollectionExitsValue(cllFeeBalance, strKey) Then
                dblMoney = cllFeeBalance(strKey)(4) + Val(Nvl(!��Ԥ��))
                cllFeeBalance.Remove strKey
            Else
                dblMoney = Val(Nvl(!��Ԥ��))
            End If
            If RoundEx(dblMoney, 6) > 0 Then 'ȫ������ľͲ��ټ���
                'Array(���㷽ʽ,�����ID,�Ƿ�����,���������,��Ԥ��,�Ƿ�ȫ��,�Ƿ�ת�ʼ�����)
                cllFeeBalance.Add Array(Nvl(!���㷽ʽ), Val(Nvl(!�����ID)), Val(Nvl(!�Ƿ�����)), _
                    Nvl(!���������), dblMoney, Val(Nvl(!�Ƿ�ȫ��)), Nvl(!�Ƿ�ת�ʼ�����)), strKey
            End If
            .MoveNext
        Loop
    End With
    If cllFeeBalance.Count = 0 Then ThreeBalanceCheck = True: Exit Function
    
    For i = 1 To cllFeeBalance.Count
        blnQuestion = False
        'ҽ�ƿ����
        If objCards Is Nothing Then
            If MsgBox("��" & cllFeeBalance(i)(3) & "��δ���ã���ҽ�ƿ�֧���Ľ�����Ϊ�������㷽ʽ���Ƿ������", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            blnQuestion = True
        Else
            blnFind = False
            For Each objCard In objCards
                If objCard.�ӿ���� = cllFeeBalance(i)(1) Then blnFind = True: Exit For
            Next
            If blnFind = False Then
                If MsgBox("��" & cllFeeBalance(i)(3) & "��δ���ã���ҽ�ƿ�֧���Ľ�����Ϊ�������㷽ʽ���Ƿ������", _
                    vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                blnQuestion = True
            End If
        End If
        
        If blnQuestion Then
            If cllFeeBalance(i)(2) = 0 Then 'ǿ������
                If str����Ա = "" Then '���ֿ����ʱֻ��֤һ��
                    If zlStr.IsHavePrivs(GetPrivFunc(glngSys, 1151), "�����˿�ǿ������") Then
                        str����Ա = UserInfo.����
                    Else
                        str����Ա = zlDatabase.UserIdentifyByUser(Me, "ҽ�ƿ���" & cllFeeBalance(i)(3) & "��ǿ�����֣�Ȩ����֤��", _
                            glngSys, mlngModule, "�����˿�ǿ������", , True)
                        If str����Ա = "" Then Exit Function
                    End If
                End If
                'Array(����Ա,���������,���㷽ʽ)
                cllForceDelToCash.Add Array(str����Ա, cllFeeBalance(i)(3), cllFeeBalance(i)(0))
            End If
        ElseIf cllFeeBalance(i)(5) = 1 Then '����ȫ��
            If cllFeeBalance(i)(2) = 1 Then '�������֣�����ȫ��
                If cllFeeBalance(i)(6) = 0 Then '��֧��ת�ʼ�����
                    If MsgBox("��" & cllFeeBalance(i)(3) & "������ȫ�ˣ���˲����˻�ԭ����" & _
                        "���������������ô��ҽ�ƿ�֧���Ľ�����Ϊ�������㷽ʽ���Ƿ������", _
                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                    str�ų����㷽ʽ = str�ų����㷽ʽ & "," & cllFeeBalance(i)(0)
                End If
            ElseIf cllFeeBalance(i)(6) = 0 Then '���������֣�����ȫ�ˣ��Ҳ�֧��ת�ʼ�����
                If MsgBox("��" & cllFeeBalance(i)(3) & "������ȫ���Ҳ������֣�ͬʱҲ��֧��ת�ʼ����ۣ�����޷��˻�ԭ����" & _
                    "���������������ô��ҽ�ƿ�֧���Ľ�����Ϊ�������㷽ʽ���Ƿ������", _
                    vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                If str����Ա = "" Then '���ֿ����ʱֻ��֤һ��
                    If zlStr.IsHavePrivs(GetPrivFunc(glngSys, 1151), "�����˿�ǿ������") Then
                        str����Ա = UserInfo.����
                    Else
                        str����Ա = zlDatabase.UserIdentifyByUser(Me, "��" & cllFeeBalance(i)(3) & "��ǿ�����֣�Ȩ����֤��", _
                            glngSys, mlngModule, "�����˿�ǿ������", , True)
                        If str����Ա = "" Then Exit Function
                    End If
                End If
                'Array(����Ա,���������,���㷽ʽ)
                cllForceDelToCash.Add Array(str����Ա, cllFeeBalance(i)(3), cllFeeBalance(i)(0))
                str�ų����㷽ʽ = str�ų����㷽ʽ & "," & cllFeeBalance(i)(0)
            End If
        End If
    Next
    If str�ų����㷽ʽ <> "" Then str�ų����㷽ʽ = Mid(str�ų����㷽ʽ, 2)
    

    If str�ų����㷽ʽ = "" Then ThreeBalanceCheck = True: Exit Function
    '�ж��Ƿ�����Ч�Ľ��㷽ʽ
    varData = Split(str�ų����㷽ʽ, ",")
    lngCount = mobjPayCards.Count
    For i = 1 To mobjPayCards.Count
        If mobjPayCards(i).�ӿ���� <= 0 Or mobjPayCards(i).�ӿ���� > 0 And mobjPayCards(i).���ѿ� Then
            Exit For
        End If
        
        blnFind = False
        For j = 0 To UBound(varData)
            If mobjPayCards(i).���㷽ʽ = varData(j) Then
                lngCount = lngCount - 1: blnFind = True
            End If
        Next
        If blnFind = False Then Exit For
    Next
    If lngCount <= 0 Then
        MsgBox "�ų�ǿ�����ֵĽ��㷽ʽ����û�п��õĽ��㷽ʽ�����ܼ�����", vbInformation, gstrSysName
        Exit Function
    End If
    ThreeBalanceCheck = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetCashMoney(ByVal strNo As String) As Currency
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '���ܣ�ҽ����֧���˸����ʻ�ʱ,�����ʻ����ֽ�,��ȡ�ֽ��˿���
    '������
    '   strNO-�Һŵ���
    '���أ�������
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errH
    strSQL = "" & _
            "   Select -1 * a.��Ԥ�� As �ֽ�" & _
            "   From ����Ԥ����¼ A, ������ü�¼ B, ���㷽ʽ C" & _
            "   Where a.����id = b.����id And a.���㷽ʽ Is Null" & _
            "         And b.No = [1] And a.��¼���� = 4 And a.��¼״̬ = 2 And Rownum = 1"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡҽ��������", strNo)
    
    If Not rsTmp.BOF Then GetCashMoney = CCur(Nvl(rsTmp!�ֽ�))
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ExecuteDelRegister() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ���˺Ų���
    '����:�˺ųɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-10-09 16:57:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, cllPro As Collection, strNo As String
    Dim str����ID As String, str������� As String, strDelDate As String
    Dim str�������ID As String, strAdvance As String, str�����ʻ� As String
    
    On Error GoTo errHandle
    If isRegisterValied(strNo) = False Then Exit Function

    '����Ƿ�����ҽ������
    str�����ʻ� = IIf(mstr�����ʻ� <> "", mstr�����ʻ�, "�����ʻ�")
    If mCurBillType.intInsure <> 0 Then
        If gclsInsure.GetCapability(support�����������, , mCurBillType.intInsure, str�����ʻ�) Then
            str�����ʻ� = ""     '����̴��벻�����˵Ľ��㷽ʽ,�ձ�ʾȫ������
        End If
    End If
    
    Set cllPro = New Collection
    str����ID = zlDatabase.GetNextId("���˽��ʼ�¼")
    str������� = "-" & str����ID
    strDelDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    
    '1.��������
    'Zl_���˹ҺŲ�����_Delete
    strSQL = "Zl_���˹ҺŲ�����_Delete("
    '  ���ݺ�_In     ������ü�¼.No%Type,
    strSQL = strSQL & "'" & strNo & "',"
    '  ����Ա���_In ������ü�¼.����Ա���%Type,
    strSQL = strSQL & "'" & UserInfo.��� & "',"
    '  ����Ա����_In ������ü�¼.����Ա����%Type,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '  ����id_In     ������ü�¼.����id%Type := Null,
    strSQL = strSQL & "" & str����ID & ","
    '  �������_In   ����Ԥ����¼.�������%Type := Null,
    strSQL = strSQL & "" & str������� & ","
    '  �˺�ʱ��_In   ������ü�¼.�Ǽ�ʱ��%Type := Null,
    strSQL = strSQL & "to_date('" & strDelDate & "','yyyy-mm-dd hh24:mi:ss'))"
    '  ɾ�������_In Number:=0
    zlAddArray cllPro, strSQL
    
    str�������ID = zlDatabase.GetNextId("���˽��ʼ�¼")
    'Zl_���ò����¼_Delete
    strSQL = "Zl_���ò����¼_Delete("
    '  No_In         In ���ò����¼.No%Type,
    strSQL = strSQL & "'" & mCurBillType.str���㵥 & "',"
    '  ����id_In     In ���ò����¼.����id%Type,
    strSQL = strSQL & "" & str�������ID & ","
    '  �ؽ�id_In     In ���ò����¼.����id%Type,
    strSQL = strSQL & "NULL,"
    '  �������_In   In ���ò����¼.�������%Type,
    strSQL = strSQL & "" & str������� & ","
    '  �˷ѽ���id_In varchar2(����ö��������),
    strSQL = strSQL & "'" & str����ID & "',"
    '  ����Ա���_In In ���ò����¼.����Ա���%Type,
    strSQL = strSQL & "'" & UserInfo.��� & "',"
    '  ����Ա����_In In ���ò����¼.����Ա����%Type,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '  �Ǽ�ʱ��_In   In ���ò����¼.�Ǽ�ʱ��%Type := Null
    strSQL = strSQL & "to_date('" & strDelDate & "','yyyy-mm-dd hh24:mi:ss'),"
    '  ��ԭ���˽���_In In Varchar2 := Null
    strSQL = strSQL & "'" & str�����ʻ� & "')"
    zlAddArray cllPro, strSQL
    Err = 0: On Error GoTo ErrRoll:
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    
    '����ҽ���ӿ�
    '�Һŷ���ȡ��ʽ|�Һŵ���|�������־,��|�ָ�
    strAdvance = "0|" & strNo & "|1"
    If Not gclsInsure.RegistDelSwap(mCurBillType.lngԭ����ID, mCurBillType.intInsure, strAdvance) Then
        gcnOracle.RollbackTrans: Exit Function
    End If
    gcnOracle.CommitTrans
    Call gclsInsure.BusinessAffirm(����Enum.Busi_RegistDelSwap, True, mCurBillType.intInsure)

    If str�����ʻ� <> "" Then
        MsgBox "ҽ����֧��[" & str�����ʻ� & "]���ˣ�����Ϊ�������㷽ʽ��" & vbCrLf & vbCrLf & "�˿��:" & Format(GetCashMoney(strNo), "0.00") & " Ԫ��", vbInformation, gstrSysName
    End If

    '2.��ʾ�������
    mCurBillType.lng������� = Val(str�������) '��¼���ڴ�ӡ��Ʊ
    On Error GoTo errHandle
    Dim frmBalance As New frmReplenishTheBalanceDelWin, objDelBalance As New clsCliniDelBalance
    Set objDelBalance.rsBalance = mrsBalance
    Set objDelBalance.rs���㷽ʽ = mrs���㷽ʽ
    With objDelBalance
        .intInsure = mCurBillType.intInsure
        .CurDelNos = mCurBillType.strNos
        .AllNos = mCurBillType.strAllNOs
        .PrintNOs = "'" & mCurBillType.str���㵥 & "'"
        .PatiUseType = mobjFact.ʹ�����
        .SaveBilled = True
        .ShareUserID = mobjFact.��������ID
        .����ID = mCurBillType.lng����ID
        .����ID = Val(str�������ID)
        .��ǰ��Ʊ�� = ""
        .���շ�Ʊ = ""
        .������� = Val(str�������)
        .����ID = 0
        .ȱʡ���㷽ʽ = mCurBillType.str���㷽ʽ
        .�˷Ѻϼ� = -1 * GetDelMoney
        .�ѱ� = mCurBillType.str�ѱ�
        .���� = mCurBillType.str����
        .�Ա� = mCurBillType.str�Ա�
        .���� = mCurBillType.str����
        .�������� = mCurBillType.str��������
        .ҽ������Ʊ�� = MCPAR.ҽ������Ʊ��
        .ԭ����ID = mCurBillType.lngԭ����ID
        .�˷�ʱ�� = strDelDate
        .�����˷� = False
        .ԭ���� = False
    End With
    
    Call GetAsyncKeyState(VK_RETURN)
    If frmBalance.zlChargeWin(Me, mlngModule, mstrPrivs, EM_BalanceDel, mobjPayCards, objDelBalance, MCPAR.�ֱҴ���, _
        mcllForceDelToCash, mstr�ų����㷽ʽ, True) = False Then Exit Function
    If Not gfrmMain Is Nothing Then
        Call zlExeBalanceWinRefrshData(True, objDelBalance)
    End If
    ExecuteDelRegister = True
    Exit Function
ErrRoll:
    gcnOracle.RollbackTrans
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function isRegisterValied(ByRef strNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���Һ��˺��Ƿ�Ϸ�
    '����:strNO-���عҺŵ���
    '����:�Ϸ�����true,���򷵻�False
    '����:���˺�
    '����:2014-10-09 16:58:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnSel  As Boolean, strNos As String, strTemp As String, blnTemp As Boolean
    Dim strOperatorName As String, i As Long
    
    On Error GoTo errHandle
    If Not CheckTextLength("�˷�ժҪ", txt�˷�ժҪ) Then Exit Function
    '��������Ƿ���ȷ
    If mCurBillType.strNos = "" Then
        MsgBox "����ȷ����Ҫ�˺ŵ�����Һŵ��ݡ�", vbInformation, gstrSysName
        If txtNO.Visible Then txtNO.SetFocus: Exit Function
    End If
    If mCurBillType.str���㵥 = "" Then
        MsgBox "δ�ҵ���Ӧ�Ĳ�������¼,�������˺�!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
        
    '��鱾�ν��㵥�����Ƿ�����˷��쳣���ݣ������ڣ�����������˷�
    If CheckIsExistDelErrBill(mCurBillType.str���㵥, strOperatorName) Then
        MsgBox "ע�⣺" & vbCrLf & _
            "    ���ν����д����쳣�Ľ����¼�����ȶ�����������˷ѣ�" & _
            IIf(strOperatorName <> UserInfo.����, vbCrLf & "    ��ʾ���쳣�����ǲ���Ա��" & strOperatorName & "����ȡ�ġ�", ""), vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If

    blnSel = False
    With vsBill
        strNos = ""
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("ѡ��"))) <> 0 Then
                strTemp = Trim(.TextMatrix(i, .ColIndex("���ݺ�")))
                If strTemp <> "" Then
                    If InStr(1, strNos & ",", "," & strTemp & ",") = 0 Then
                        strNos = strNos & "," & strTemp
                        blnTemp = False
                        If Not zlCheckRegBillIsExecuted(strTemp, True, blnTemp) Then vsBill.SetFocus: Exit Function
                        If blnTemp Then
                            MsgBox "�Һŵ�" & strTemp & "�Ѿ���ҽ��������¹�ҽ��,�����˺ţ�", vbInformation, gstrSysName
                             vsBill.SetFocus: Exit Function
                        End If
                    End If
                    blnSel = True
                End If
            End If
        Next
    End With
    
    If blnSel = False Then
        MsgBox "���ڵ���������ѡ��һ��Ҫ�˺ŵĹҺŵ���", vbInformation, gstrSysName
        vsBill.SetFocus: Exit Function
    End If
    
    If strNos <> "" Then strNos = Mid(strNos, 2)
    If InStr(1, strNos, ",") > 0 Then
        MsgBox "����һ���˶���Һŵ���,���顣", vbInformation, gstrSysName
        vsBill.SetFocus: Exit Function
    End If
    strNo = strNos
    'ҽ�����
    If mCurBillType.intInsure = 0 Then
        MsgBox "��ǰ���㲻��ҽ�����˽���,����������˺Ų���!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If gclsInsure.CheckInsureValid(mCurBillType.intInsure) = False Then Exit Function
    
    '���������㷽ʽ��Ч�Լ��
    If ThreeBalanceCheck(mobjPayCards, mrsBalance, mcllForceDelToCash, mstr�ų����㷽ʽ) = False Then Exit Function
    
    isRegisterValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Private Function ExecuteDelChargeFee() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ�����շѷ��ò���
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-10-11 11:07:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim frmBalance As frmReplenishTheBalanceDelWin, objDelBalance As New clsCliniDelBalance
    Dim arrNo As Variant, k As Long, i As Long, j As Long, lngCount As Long
    Dim lngCheck����ID As Long, intCheckInsure As Integer
    Dim strBalanceInfor As String, strCurSelNos As String, strNo As String, str��� As String
    Dim strTemp As String, strReclaimInvoice As String, strInvoice As String, strYBPati As String
    Dim str�������ID As String, str�ؽ�ID As String, strSQL As String
    Dim str����ID As Long, str����ID As Long, str������� As Long, lng����ID As Long
    Dim blnAll�����˷� As Boolean, blnCur�����˷� As Boolean, blnȫ�� As Boolean, blnTrans As Boolean
    Dim cllPro As Collection, colOrder As New Collection
    Dim cur����͸֧ As Currency
    Dim varTemp As Variant
    Dim dtDelDate As Date
    Dim strReturn As String, strReturnRecipt As String '�˷Ѵ�����Ϣ����ʽ��NO,ҩ��ID|NO,ҩ��ID|��
    
    If isChargeFeeValied = False Then Exit Function
    
    On Error GoTo Errhand:
    '���ж����е����Ƿ񲿷��˷�,�Ծ���Ʊ�ݵĴ���ʽ
    arrNo = Split(mCurBillType.strNos, ",")
    
    blnAll�����˷� = False
    strCurSelNos = ""
    Set cllPro = New Collection
    For i = 0 To UBound(arrNo)
        strNo = arrNo(i)
        str��� = "":   lngCount = 0
        '�ռ���ǰ����Ҫ�˷ѵ��к�
        With vsBill
            k = vsBill.FindRow(strNo, , .ColIndex("���ݺ�"))
            For j = k To vsBill.Rows - 1
                If vsBill.TextMatrix(j, .ColIndex("���ݺ�")) <> strNo Then Exit For
                If Val(vsBill.TextMatrix(j, .ColIndex("ѡ��"))) <> 0 Then
                    str��� = str��� & "," & CLng(vsBill.RowData(j))
                    If InStr(1, strCurSelNos & ",", "," & strNo & ",") = 0 Then
                        strCurSelNos = strCurSelNos & "," & strNo
                    End If
                    '81190,Ƚ����,�˷�ҵ����ҩ���ϴ��˷���Ϣ
                    '��ʽ��NO,ҩ��ID|NO,ҩ��ID|��
                    If vsBill.TextMatrix(j, vsBill.ColIndex("���")) Like "*��*ҩ*" _
                        Or vsBill.TextMatrix(j, vsBill.ColIndex("���")) Like "*��*ҩ*" Then
                        If InStr(strReturnRecipt & "|", _
                            "|" & vsBill.TextMatrix(j, vsBill.ColIndex("���ݺ�")) & "," & vsBill.TextMatrix(j, vsBill.ColIndex("ִ�п���ID")) & "|") = 0 Then
                            strReturnRecipt = strReturnRecipt & "|" & vsBill.TextMatrix(j, vsBill.ColIndex("���ݺ�")) & "," & vsBill.TextMatrix(j, vsBill.ColIndex("ִ�п���ID"))
                        End If
                    End If
                End If
                lngCount = lngCount + 1
            Next
        End With
        str��� = Mid(str���, 2)
        If str��� <> "" Then
            blnCur�����˷� = Not BillDeleteAllNew(strNo, 1)
            If UBound(Split(str���, ",")) + 1 = lngCount And blnCur�����˷� = False Then str��� = ""
            blnCur�����˷� = Not (Not blnCur�����˷� And str��� = "")
            If blnCur�����˷� Then blnAll�����˷� = True '���ŵ���Ϊ�����˷�,�����е���Ϊ�����˷�
            colOrder.Add str���, "_" & strNo
        Else
            blnAll�����˷� = True                       '���ŵ��ݲ��˷�,�����е���Ϊ�����˷�
            colOrder.Add "δѡ��", "_" & strNo
        End If
    Next
    
    '�������������Ƿ�δ����,����жϳ����е����Ƿ񲿷��˷�
    If Not blnAll�����˷� Then
        varTemp = Split(mCurBillType.strAllNOs, ",")
        strTemp = ""
        For i = 0 To UBound(varTemp)
            If InStr(1, "," & mCurBillType.strNos & ",", "," & varTemp(i) & ",") = 0 Then
                strTemp = strTemp & "," & varTemp(i)
                 blnAll�����˷� = True: Exit For
            End If
        Next
    End If
    
    If CheckSelectItemCanDel(strCurSelNos) = False Then Exit Function
    
    '��ʾ����Ʊ��
    If ShowReclaimInvoice(mCurBillType.str���㵥, strReclaimInvoice) = False Then Exit Function
    
    If mCurBillType.intInsure <> 0 And MCPAR.ҽ���ӿڴ�ӡƱ�� Then
        If zlGetInvoiceGroupUseID(lng����ID) = False Then Exit Function
        strInvoice = GetNextBill(lng����ID)
    End If
    
    dtDelDate = zlDatabase.Currentdate
    blnȫ�� = True
'    If blnAll�����˷� Then blnȫ�� = False
    If blnȫ�� Then blnȫ�� = CheckIsAllDel(mCurBillType.strAllNOs)
     '����ҽ��
    If Not blnȫ�� Then
        '���ܴ��������շ�,���,��Ҫ���������֤�ӿ�(Identifiy)
        'strAdvace:ҽ��������ʱ:����1,��ʾҽ�������˺��������շѵ������֤;��������: ��
        lngCheck����ID = mCurBillType.lng����ID
        intCheckInsure = mCurBillType.intInsure
        strYBPati = gclsInsure.Identify(0, lngCheck����ID, intCheckInsure, 2)
        
        If strYBPati = "" Then
            MsgBox "ҽ�������֤ʧ��,����������˷�!", vbOKOnly + vbDefaultButton1 + vbExclamation, gstrSysName
            Exit Function
        End If
        
        If Val(CLng(Split(strYBPati, ";")(8))) <> mCurBillType.lng����ID Then
            MsgBox "ҽ����֤�Ĳ������˷ѵĲ��˲���ͬһ������!", vbInformation, gstrSysName
            Call ExecuteYBIdentifyCancel(mCurBillType.lng����ID, mCurBillType.intInsure)
            Exit Function
        End If
    End If
    
       
    '��������:����Ҫִ�е�SQL
    str����ID = zlDatabase.GetNextId("���˽��ʼ�¼")
    str������� = -1 * str����ID
    mCurBillType.strDelNOs = ""
    For i = UBound(arrNo) To 0 Step -1
        strNo = arrNo(i)
        If colOrder("_" & strNo) <> "δѡ��" Then
            ' Zl_�����շѼ�¼_����
            strSQL = "Zl_�����շѼ�¼_����("
            '  No_In         ������ü�¼.No%Type,
            strSQL = strSQL & "'" & strNo & "',"
            '  ����Ա���_In ������ü�¼.����Ա���%Type,
            strSQL = strSQL & "'" & UserInfo.��� & "',"
            '  ����Ա����_In ������ü�¼.����Ա����%Type,
            strSQL = strSQL & "'" & UserInfo.���� & "',"
            '  ���_In       Varchar2 := Null,
            strSQL = strSQL & "'" & colOrder("_" & strNo) & "',"
            '  �˷�ʱ��_In   ������ü�¼.�Ǽ�ʱ��%Type := Null,
            strSQL = strSQL & "To_Date('" & Format(dtDelDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
            '  �˷�ժҪ_In   ������ü�¼.ժҪ%Type := Null,
            strSQL = strSQL & "" & IIf(Trim(txt�˷�ժҪ.Text) = "", "NULL", "'" & Trim(txt�˷�ժҪ.Text) & "'") & ","
            '  ����id_In     ����Ԥ����¼.����id%Type := Null,
            strSQL = strSQL & str����ID & ","
            '  ����Ʊ��_In Number:=0
            strSQL = strSQL & "0)"  '�����¼�н��л���
            zlAddArray cllPro, strSQL
            mCurBillType.strDelNOs = mCurBillType.strDelNOs & "," & strNo
        End If
    Next
     
    ' Zl_�����˷ѽ���_Modify
    strSQL = "Zl_�����˷ѽ���_Modify("
    '  ��������_In   Number,
    '  --   0-ԭ����
    '  --      ԭ������һ��ȫ��,����У�Ա�־��Ϊ1,ҽ�����óɹ���,����Ϊ2,��ɺ���0
    '  --   1-��ͨ�˷ѷ�ʽ:
    '  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ||.." ;Ҳ�������.
    '  --     �ڳ�Ԥ��_In:����漰Ԥ����,���뱾�ε���Ԥ��,�������շ�ʱ,������(<0 ��ʾ��Ԥ����;>0 ��ʾ��ʣ�������Ԥ����¼
    '  --   2.�������˷ѽ���:
    '  --     �ٽ��㷽ʽ_IN:ֻ�ܴ���һ�����㷽ʽ,���������һЩ������Ϣ,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ"
    '  --     ����Ԥ��_In: ������
    '  --     �ۿ����ID_IN,����_IN,������ˮ��_IN,����˵��_In:��Ҫ����
    '  --   3-ҽ������(�������ҽ���Ľ���,��Ҫ��ɾ��ԭҽ������,���´���ĸ���)
    '  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:���㷽ʽ|������||.."
    '  --     ����Ԥ��_In: ������
    '  --     ����֧Ʊ��_In:������
    '  --   4-���ѿ�����:
    '  --     �ٽ��㷽ʽ_IN:����һ��ˢ���ſ�,��ʽΪ:�����ID|����|���ѿ�ID|���ѽ��||."
    '  --     ����Ԥ��_In: ������
    '  --     ����֧Ʊ��_In:������
    strSQL = strSQL & "" & 1 & ","
    '  ����id_In     ������ü�¼.����id%Type,
    strSQL = strSQL & "" & mCurBillType.lng����ID & ","
    '  ����id_In     ����Ԥ����¼.����id%Type,
    strSQL = strSQL & "" & str����ID & ","
    '  ���㷽ʽ_In   Varchar2,
    strSQL = strSQL & "" & "NULL" & ")"
    '  ��Ԥ��_In     ����Ԥ����¼.��Ԥ��%Type := Null,
    '  �����id_In   ����Ԥ����¼.�����id%Type := Null,
    '  ����_In       ����Ԥ����¼.����%Type := Null,
    '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
    '  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
    '  �ɿ�_In       ����Ԥ����¼.�ɿ�%Type := Null,
    '  �Ҳ�_In       ����Ԥ����¼.�Ҳ�%Type := Null,
    '  �����_In   ������ü�¼.ʵ�ս��%Type := Null,
    '  ����˷�_In   Number := 0,
    '  ԭ����id_In   ����Ԥ����¼.����id%Type := Null
    zlAddArray cllPro, strSQL
    
    str�������ID = zlDatabase.GetNextId("���˽��ʼ�¼")
    str�ؽ�ID = zlDatabase.GetNextId("���˽��ʼ�¼")
    
    '�ȳ���ԭʼ�Ľ����¼
    'Zl_���ò����¼_Delete
    strSQL = "Zl_���ò����¼_Delete("
    '  No_In         In ���ò����¼.No%Type,
    strSQL = strSQL & "'" & mCurBillType.str���㵥 & "',"
    '  ����id_In     In ���ò����¼.����id%Type,
    strSQL = strSQL & "" & str�������ID & ","
    '  �ؽ�id_In     In ���ò����¼.����id%Type,
    strSQL = strSQL & "" & str�ؽ�ID & ","
    '  �������_In   In ���ò����¼.�������%Type,
    strSQL = strSQL & "" & str������� & ","
    '  �˷ѽ���id_In In ���ò����¼.����id%Type,(�����˷Ѽ�¼�Ľ���ID)
    strSQL = strSQL & "" & str����ID & ","
    '  ����Ա���_In In ���ò����¼.����Ա���%Type,
    strSQL = strSQL & "'" & UserInfo.��� & "',"
    '  ����Ա����_In In ���ò����¼.����Ա����%Type,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '  �Ǽ�ʱ��_In   In ���ò����¼.�Ǽ�ʱ��%Type := Null
    strSQL = strSQL & "To_Date('" & Format(dtDelDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
    zlAddArray cllPro, strSQL
    
    
    '����ҽ���ӿ�
    '�Ȼ���Ʊ�ݣ�Ԥ����֮���ٲ���Ʊ��
    If MCPAR.ҽ���ӿڴ�ӡƱ�� Then
        If Not blnȫ�� Then 'Ԥ����֮���ٷ���Ʊ��
            '56963,77058
            strSQL = "zl_�����շѼ�¼_RePrint('" & mCurBillType.str���㵥 & "',NULL," & ZVal(lng����ID) & ",'" & UserInfo.���� & "'," & _
                "To_Date('" & Format(dtDelDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),1,1)"
            zlAddArray cllPro, strSQL
        Else 'ȫ�˷�ҲҪ����Ʊ�ݺţ�����ҽ��
            strSQL = "zl_�����շѼ�¼_RePrint('" & mCurBillType.str���㵥 & "','" & strInvoice & "'," & ZVal(lng����ID) & ",'" & UserInfo.���� & "'," & _
                "To_Date('" & Format(dtDelDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),1,1)"
            zlAddArray cllPro, strSQL
        End If
    End If
    
    blnTrans = True
    '1.���ݱ���:��������,�ؽ�����
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    If ExecuteClinicDelSwap(mCurBillType.lng����ID, mCurBillType.intInsure, Val(str�������ID), mCurBillType.lngԭ����ID) = False Then Exit Function
    
     blnTrans = False
    '2.���ݱ���:��������,�ؽ�����
    Set cllPro = New Collection
    
    '���½����շѴ���
    If Not blnȫ�� Then
        '��ȡ�������
        cur����͸֧ = mTy_Insure.dbl����͸֧
        mTy_Insure.dbl�ʻ���� = gclsInsure.SelfBalance(mCurBillType.lng����ID, CStr(Split(strYBPati, ";")(1)), 10, cur����͸֧, mCurBillType.intInsure)
        mTy_Insure.dbl����͸֧ = cur����͸֧
        '�������ռ�¼�ı�����Ϣ
        If GetExcutInsureInforUpdateSQL(str�������, strBalanceInfor, cllPro) = False Then Exit Function
        blnTrans = True: zlExecuteProcedureArrAy cllPro, Me.Caption, True
        '77058
        If ExcuteInsureReCharge(mCurBillType.lng����ID, mCurBillType.intInsure, str�ؽ�ID, str�������, strBalanceInfor, _
                mCurBillType.str���㵥, lng����ID, strInvoice, dtDelDate) = False Then Exit Function
    End If
    
    blnTrans = False
    '4.��ʾ�������
    
    mCurBillType.lng������� = Val(str�������) '��¼���ڴ�ӡ��Ʊ
    Set objDelBalance.rsBalance = mrsBalance
    Set objDelBalance.rs���㷽ʽ = mrs���㷽ʽ
    
    With objDelBalance
        .intInsure = mCurBillType.intInsure
        .CurDelNos = mCurBillType.strDelNOs
        .AllNos = mCurBillType.strAllNOs
        .PrintNOs = "'" & mCurBillType.str���㵥 & "'"
        .PatiUseType = mobjFact.ʹ�����
        .SaveBilled = True
        .ShareUserID = mobjFact.��������ID
        .����ID = mCurBillType.lng����ID
        .����ID = str�������ID
        .��ǰ��Ʊ�� = strInvoice
        .���շ�Ʊ = strReclaimInvoice
        .������� = str�������
        .����ID = str����ID
        .ȱʡ���㷽ʽ = mCurBillType.str���㷽ʽ
        .�˷Ѻϼ� = -1 * GetDelMoney
        .�ѱ� = mCurBillType.str�ѱ�
        .���� = mCurBillType.str����
        .�Ա� = mCurBillType.str�Ա�
        .���� = mCurBillType.str����
        .�������� = mCurBillType.str��������
        .ҽ������Ʊ�� = MCPAR.ҽ������Ʊ��
        .ԭ����ID = mCurBillType.lngԭ����ID
        .�˷�ʱ�� = dtDelDate
        .�����˷� = Not blnȫ��
        .ԭ���� = False
    End With
    Call GetAsyncKeyState(VK_RETURN)
    Set frmBalance = New frmReplenishTheBalanceDelWin
    If frmBalance.zlChargeWin(Me, mlngModule, mstrPrivs, EM_BalanceDel, mobjPayCards, objDelBalance, MCPAR.�ֱҴ���, _
        mcllForceDelToCash, mstr�ų����㷽ʽ, False) = False Then Exit Function

    '81190,Ƚ����,�˷�ҵ����ҩ���ϴ��˷���Ϣ
    On Error Resume Next
    If mblnDrugMachine Then
        Dim rsTemp As ADODB.Recordset, strData As String '���ﴦ����ҩ��ʽ������ID1,��ҩ����1;����ID2,��ҩ����2;...
        '�����˵ļ�ȥ���յľ���ʵ���˵�
        strSQL = "Select Max(Decode(a.��¼״̬, 2, a.Id, 0)) As ����id, -1 * Nvl(Sum(a.���� * a.����), 0) As ��ҩ����" & vbNewLine & _
                " From ������ü�¼ A,(Select Distinct ����ID From ����Ԥ����¼ Where ������� = [1]) B" & vbNewLine & _
                " Where a.����id = b.����ID And Mod(a.��¼����, 10) = 1 And a.�շ���� In ('5', '6', '7')" & vbNewLine & _
                " Group By NO, Nvl(�۸񸸺�, ���)" & vbNewLine & _
                " Having Nvl(Sum(a.���� * a.����), 0) <> 0"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ�����˷���Ŀ", objDelBalance.�������)
        Do While Not rsTemp.EOF
            strData = strData & ";" & Nvl(rsTemp!����id) & "," & Nvl(rsTemp!��ҩ����)
            rsTemp.MoveNext
        Loop
        If strData <> "" Then
            strData = Mid(strData, 2)
            Call mobjDrugMachine.Operation(gstrDBUser, Val("24-������ҩ(����/����)"), strData, strReturn)
        End If
    ElseIf mblnDrugPacker Then
        If strReturnRecipt <> "" Then
            strReturnRecipt = Mid(strReturnRecipt, 2)
            Call mobjDrugPacker.DYEY_MZ_TransRecipeReturn(1, UserInfo.���, UserInfo.����, strReturnRecipt, strReturn)
        End If
    End If
    Err.Clear: On Error GoTo Errhand
    
    If Not gfrmMain Is Nothing Then
        Call zlExeBalanceWinRefrshData(True, objDelBalance)
    End If
    ExecuteDelChargeFee = True
    Exit Function
Errhand:
    If blnTrans Then
        gcnOracle.RollbackTrans
        Call ErrCenter
    Else
        If ErrCenter = 1 Then Resume
    End If
    Call SaveErrLog
End Function

Private Sub PrintDelBill(ByVal strNo As String, ByVal lng����ID As Long, _
    ByVal dtDateDel As Date, ByVal bln������ As Boolean, _
    ByVal strInvoices As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ӡ���Ʊ��
    '���:  strNO-��ǰ���㵥��
    '       dtDateDel-�˷�����
    '����:���˺�
    '����:2014-10-11 11:36:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intInvoiceFormat As Integer, blnPrint As Integer
    Dim str��Ʊ�� As String, intƱ������ As Integer
    Dim strSQL As String, strTempNO As String, i As Integer

    On Error GoTo errHandle
    If Not bln������ Then
         '˰�ز���ȫ��ʱ�ջش���(ȫ��ʱ��Zl_���ò����¼_Delete�����ջ�Ʊ��)
        If Not gobjTax Is Nothing And gblnTax Then
            gstrTax = gobjTax.zlTaxOutErase(gcnOracle, strNo)
            If gstrTax <> "" Then MsgBox gstrTax, vbExclamation, gstrSysName
        End If
        GoTo PrintList:
        Exit Sub
    End If
    '77058
    If bln������ And mCurBillType.intInsure <> 0 And MCPAR.ҽ���ӿڴ�ӡƱ�� Then GoTo PrintList
    If (strInvoices = "�޿���Ʊ��" Or strInvoices = "") And bln������ Then  'a.�ջز����´�ӡ�����վ�
        blnPrint = True
        ''0-����ӡ;1-�Զ���ӡ;2-��ʾ��ӡ
        blnPrint = True
        If mobjFact.��ӡ��ʽ = 0 Then blnPrint = False
        If mobjFact.��ӡ��ʽ = 2 Then
            If MsgBox("�Ƿ��ӡƱ�ݣ�", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then blnPrint = False
        End If
        
        If blnPrint Then
            intInvoiceFormat = mobjFact.��ӡ��ʽ
            Call zlRePrintReplenishTheBalanceBill(Me, mlngModule, 1, mCurBillType.str���㵥, mCurBillType.intInsure, mobjInvoice, mobjFact, True, dtDateDel)
        End If
        GoTo PrintList:
        Exit Sub
    End If


    'b.�շѻ���һ����ʱû�д�ӡƱ��
    If strInvoices <> "�޿���Ʊ��" And strInvoices <> "" Then
        'c.ֻ�ջ�Ʊ��
        strSQL = "Zl_�������Ʊ��_Reprint('" & strNo & "',Null,0,'" & UserInfo.���� & "'," & _
            "To_Date('" & Format(dtDateDel, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),1,0,'" & strInvoices & "')"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    End If
PrintList:
    '�˷ѷ�Ʊ(��Ʊ)��ӡ��91998
    Call mobjInvoice.zlGetInvoicePreperty(mlngModule, EM_�˷��վ�, mCurBillType.lng����ID, 0, mCurBillType.intInsure, mobjFact)
    '0-����ӡ;1-�Զ���ӡ;2-��ʾ��ӡ
    If mobjFact.��ӡ��ʽ = 1 Then
        Call zlPrintReplenishTheDelBalanceBill(Me, mlngModule, mCurBillType.lng�������, mCurBillType.intInsure, mobjInvoice, mobjFact, True, dtDateDel)
    ElseIf mobjFact.��ӡ��ʽ = 2 Then
        If MsgBox("�Ƿ��ӡ�˷�Ʊ��(��Ʊ)��", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
            Call zlPrintReplenishTheDelBalanceBill(Me, mlngModule, mCurBillType.lng�������, mCurBillType.intInsure, mobjInvoice, mobjFact, True, dtDateDel)
        End If
    End If
    
    If bln������ Then
        '��ӡ�����嵥
        If zlStr.IsHavePrivs(mstrPrivs, "��������嵥") Then
            If mtyMoudlePara.int�嵥��ӡ��ʽ = 1 Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1124_1", Me, "NO=" & strNo, "ҩƷ��λ=" & IIf(mtyMoudlePara.blnҩ����λ, 1, 0), 2)
            ElseIf mtyMoudlePara.int�嵥��ӡ��ʽ = 2 Then
                If MsgBox("Ҫ��ӡ�շѽ����嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1124_1", Me, "NO=" & strNo, "ҩƷ��λ=" & IIf(mtyMoudlePara.blnҩ����λ, 1, 0), 2)
                End If
            End If
        End If
    End If
    If mCurBillType.intInsure <> 0 And MCPAR.�˷Ѻ��ӡ�ص� And InStr(1, mstrPrivs, "ҽ���˷ѻص�") > 0 Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1124_2", Me, "NO=" & strNo, 2)
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
     
Public Function Getʵ�ս��(ByVal strNo As String) As Currency
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ�����ݵ�ʵ�ս��
    '����:����ʵ�ս��
    '����:���˺�
    '����:2014-09-30 14:03:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    With vsBill
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("���ݺ�")) = strNo Then
                Getʵ�ս�� = Getʵ�ս�� + Val(.TextMatrix(i, .ColIndex("ʵ�ս��")))
            End If
        Next
    End With
End Function


Private Sub txt�˷�ժҪ_KeyDown(KeyCode As Integer, Shift As Integer)
    'ѡ���˷�ԭ��
    If KeyCode <> vbKeyReturn Then Exit Sub

    If Trim(txt�˷�ժҪ.Tag) <> "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If Trim(txt�˷�ժҪ.Text) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If zl_SelectAndNotAddItem(Me, txt�˷�ժҪ, Trim(txt�˷�ժҪ.Text), "�����˷�ԭ��", "�����˷�ԭ��ѡ��", True, True) = False Then
        If zlCommFun.IsCharChinese(Trim(txt�˷�ժҪ.Text)) = False Then Exit Sub
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txt�˷�ժҪ_GotFocus()
    zlCommFun.OpenIme True
    zlControl.TxtSelAll txt�˷�ժҪ
End Sub
Private Sub txt�˷�ժҪ_LostFocus()
    zlCommFun.OpenIme False
    If zlCommFun.ActualLen(txt�˷�ժҪ.Text) > 100 Then
        MsgBox "�˷�ժҪ�����������100���ַ���50�����֣�", vbInformation, gstrSysName
        If txt�˷�ժҪ.Visible And txt�˷�ժҪ.Enabled Then txt�˷�ժҪ.SetFocus
    End If
End Sub

Private Sub txt�˷�ժҪ_Change()
    txt�˷�ժҪ.Tag = ""
End Sub

Private Sub initCardSquareData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������رս��㿨����
    '����:���˺�
    '����:2014-09-30 14:04:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    Set mobjSquare = gobjSquare.objSquareCard
    If mbytMode = 0 Then Exit Sub
    If gobjSquare.objSquareCard Is Nothing Then
        '��������
        Call CreateSquareCardObject(gfrmMain, mlngModule)
    End If
    If gobjSquare.objSquareCard Is Nothing Then Exit Sub
    Call IDKind.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", txtPatient)
    Dim objCard As Card
    Set objCard = IDKind.GetfaultCard
    Set gobjSquare.objDefaultCard = objCard
    If IDKind.Cards.��ȱʡ������ And Not objCard Is Nothing Then
        gobjSquare.blnȱʡ�������� = objCard.�������Ĺ��� <> ""
        gobjSquare.intȱʡ���ų��� = objCard.���ų���
    Else
        gobjSquare.blnȱʡ�������� = IDKind.Cards.������ʾ
        gobjSquare.intȱʡ���ų��� = 100
    End If
    Set mobjSquare = gobjSquare.objSquareCard
End Sub


Private Function CheckBillIsAllDels(ByVal strNo As String, _
    Optional ByRef strSel��� As String) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ָ���ĵ����Ƿ�ȫ��ѡ���˷�
    '���:strNO-���ݺ�
    '����:strSel���-����ѡ�е����
    '����:0-ȫ��δѡ��;1-ȫ��ѡ��;2-ѡ����һ����
    '����:���˺�
    '����:2014-09-30 14:06:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim k As Long, j As Long, lngCount As Long, str��� As String
    With vsBill
        k = vsBill.FindRow(strNo, , vsBill.ColIndex("���ݺ�"))
         For j = k To vsBill.Rows - 1
             If vsBill.TextMatrix(j, .ColIndex("���ݺ�")) <> strNo Then Exit For
             If Val(vsBill.TextMatrix(j, .ColIndex("ѡ��"))) <> 0 Then
                 str��� = str��� & "," & CLng(vsBill.RowData(j))
             End If
             lngCount = lngCount + 1
         Next
     End With

     If str��� <> "" Then str��� = Mid(str���, 2)
     strSel��� = str���
     If str��� = "" Then CheckBillIsAllDels = 0: Exit Function
     
     If lngCount = UBound(Split(str���, ",")) + 1 Then
        If InStr(1, mCurBillType.strNosPatiDel & ",", "," & strNo & ",") > 0 Then
            CheckBillIsAllDels = 2: Exit Function
        End If
        CheckBillIsAllDels = 1: Exit Function
     End If
    CheckBillIsAllDels = 2
End Function
Private Sub ReInitPatiInvoice()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���³�ʼ�����˷�Ʊ��Ϣ
    '����:���˺�
    '����:2014-10-11 11:39:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mbytMode = EM_RBDTY_�鿴 Then Exit Sub
    Call mobjInvoice.zlGetInvoicePreperty(mlngModule, EM_�շ��վ�, mCurBillType.lng����ID, 0, mCurBillType.intInsure, mobjFact)
    Call ZlShowBillFormat(mlngModule, lblFormat, mobjFact.��ӡ��ʽ)
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
    '����:2014-09-30 14:15:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjInvoice.zlGetInvoiceGroupID(mlngModule, UserInfo.����, EM_�շ��վ�, mobjFact.ʹ�����, lng����ID, mobjFact.��������ID, lng����ID, intNum, strInvoiceNO) = False Then Exit Function
    If lng����ID <= 0 Then
        Select Case lng����ID
            Case 0 '����ʧ��
            Case -1
                If Trim(mobjFact.ʹ�����) = "" Then
                    MsgBox "��û�����ú͹��õ��շ�Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                Else
                    MsgBox "��û�����ú͹��õġ�" & mobjFact.ʹ����� & "���շ�Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                End If
                Exit Function
            Case -2
                If Trim(mobjFact.ʹ�����) = "" Then
                    MsgBox "���صĹ���Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                Else
                    MsgBox "���صĹ���Ʊ�ݵġ�" & mobjFact.ʹ����� & "���շ�Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                End If
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
 
Private Sub ClearBalance()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����������
    '����:���˺�
    '����:2014-09-30 14:16:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsBalance
        .Clear 1: .COLS = 1
        .Editable = flexEDKbdMouse
    End With
End Sub

Private Sub LoadBalanceInfor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����տ���㷽ʽ
    '����:���˺�
    '����:2014-09-30 14:17:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str���㷽ʽ As String, lngRow As Long
    Dim lngCol As Long, i As Long, intSign As Integer
    Dim lngNullCol As Long
    
    
    If mrsBalance Is Nothing Then Exit Sub
    If mrsBalance.State <> 1 Then Exit Sub
    intSign = IIf(mstrDelTime <> "", -1, 1) '����,�����������
    '�ֶ�:���� ,����ID, ��¼����, ���㷽ʽ, ժҪ, �����ID, ���������, ���ƿ�, ���㿨���, �������, ����, ������ˮ��, ����˵��, �������, У�Ա�־, ҽ��, ���ѿ�id
    '            �Ƿ�����,�Ƿ�ȫ��,�Ƿ�����,��Ԥ��
    '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
    lngRow = 0
    mrsBalance.Filter = "����=2"
    mrsBalance.Sort = "����,���㷽ʽ"
    
    '1.����ҽ������
    With vsBalance
        .Redraw = flexRDNone
        Call ClearBalance
        
        .TextMatrix(lngRow, 0) = "���ս���"
        
        If mrsBalance.RecordCount <> 0 Then mrsBalance.MoveFirst
        Do While Not mrsBalance.EOF
            '--����:52530
            str���㷽ʽ = Nvl(mrsBalance!���㷽ʽ)
            
            If str���㷽ʽ <> "" Then
                '�Ȳ����Ƿ������ͬ�Ľ��㷽ʽ,����ֱ�ӻ���
                lngCol = -1
                For i = 1 To .COLS - 1 Step 2
                    If str���㷽ʽ = .Cell(flexcpData, lngRow, i) Then
                        lngCol = i: Exit For
                    End If
                Next
                If lngCol = -1 Then
                    .COLS = .COLS + 2
                    .ColAlignment(.COLS - 2) = 7: .ColAlignment(.COLS - 1) = 1
                    lngCol = .COLS - 2
                End If
                .TextMatrix(lngRow, lngCol) = str���㷽ʽ & ":"
                .Cell(flexcpData, lngRow, lngCol) = str���㷽ʽ
                .TextMatrix(lngRow, lngCol + 1) = zlFormatNum(Val(.TextMatrix(lngRow, .COLS - 1)) + intSign * Val(Nvl(mrsBalance!��Ԥ��, 0)))
                
                .Cell(flexcpData, lngRow, lngCol + 1, lngRow, lngCol + 1) = Val(Nvl(mrsBalance!�Ƿ�����))
                If mbytMode = EM_RBDTY_�˷� Then
                    .ColData(lngCol) = "ժҪ:" & IIf(Val(Nvl(mrsBalance!�˷�)) = 1, "", Nvl(mrsBalance!ժҪ))
                    .ColData(lngCol + 1) = "�������:" & IIf(Val(Nvl(mrsBalance!�˷�)) = 1, "", Nvl(mrsBalance!�������))
                ElseIf mbytMode = EM_RBDTY_�쳣���� Then
                    .ColData(lngCol) = "ժҪ:" & IIf(Val(Nvl(mrsBalance!�˷�)) = 1, "", Nvl(mrsBalance!ժҪ))
                    .ColData(lngCol + 1) = "�������:" & IIf(Val(Nvl(mrsBalance!�˷�)) = 1, "", Nvl(mrsBalance!�������))
                Else
                    If mstrDelTime <> "" Then
                        .ColData(lngCol) = "ժҪ:" & IIf(Val(Nvl(mrsBalance!�˷�)) = 1, "", Nvl(mrsBalance!ժҪ))
                        .ColData(lngCol + 1) = "�������:" & IIf(Val(Nvl(mrsBalance!�˷�)) = 1, "", Nvl(mrsBalance!�������))
                    Else
                        .ColData(lngCol) = "ժҪ:" & IIf(Val(Nvl(mrsBalance!�˷�)) = 1, Nvl(mrsBalance!ժҪ), "")
                        .ColData(lngCol + 1) = "�������:" & IIf(Val(Nvl(mrsBalance!�˷�)) = 1, Nvl(mrsBalance!�������), "")
                    End If
                End If
            End If
             mrsBalance.MoveNext
            .Redraw = flexRDBuffered
         Loop
         
         '�ϲ�Ϊ��ļ�¼
         i = 1
         Do While i < .COLS - 1
            If Trim(.TextMatrix(lngRow, i + 1)) = "" Then
                For lngCol = i To .COLS - 3
                    .TextMatrix(lngRow, lngCol) = .TextMatrix(lngRow, lngCol + 2)
                    .Cell(flexcpData, lngRow, lngCol) = .Cell(flexcpData, lngRow, lngCol + 2)
                    .ColData(lngCol) = .ColData(lngCol + 2)
                    .Cell(flexcpForeColor, lngRow, lngCol) = .Cell(flexcpForeColor, lngRow, lngCol + 2)
                    .Cell(flexcpForeColor, 1, lngCol) = .Cell(flexcpForeColor, 1, lngCol + 2)
                    .Cell(flexcpFontBold, 1, lngCol) = .Cell(flexcpFontBold, 1, lngCol + 2)
                Next
                .COLS = .COLS - 2
            Else
                i = i + 2
            End If
         Loop
         
         '���ط�ҽ���˷Ѳ���(�˿��),��֧��Ԥ�����˿�,���Բ�����������
        mrsBalance.Filter = "����<>2"
        mrsBalance.Sort = "����,���㷽ʽ"
        If mrsBalance.RecordCount <> 0 Then mrsBalance.MoveFirst
        lngRow = 1
        .TextMatrix(lngRow, 0) = "������Ϣ"
        Do While Not mrsBalance.EOF
            If Val(mrsBalance!����) = 1 Then 'Ԥ����
                str���㷽ʽ = "��Ԥ���"
            Else
                str���㷽ʽ = Nvl(mrsBalance!���㷽ʽ)
            End If
            If str���㷽ʽ <> "" Then
                '�Ȳ����Ƿ������ͬ�Ľ��㷽ʽ,����ֱ�ӻ���
                lngCol = -1: lngNullCol = -1
                For i = 1 To .COLS - 1 Step 2
                    If str���㷽ʽ = .Cell(flexcpData, lngRow, i) Then
                        lngCol = i: Exit For
                    End If
                    If .Cell(flexcpData, lngRow, i) = "" And lngNullCol = -1 Then
                        lngNullCol = i
                    End If
                Next
                If lngCol = -1 And lngNullCol <> -1 Then lngCol = lngNullCol
                If lngCol = -1 Then
                    .COLS = .COLS + 2
                    .ColAlignment(.COLS - 2) = 7: .ColAlignment(.COLS - 1) = 1
                    lngCol = .COLS - 2
                End If
                
                .TextMatrix(lngRow, lngCol) = str���㷽ʽ & ":"
                .Cell(flexcpData, lngRow, lngCol) = str���㷽ʽ
                .TextMatrix(lngRow, lngCol + 1) = FormatEx(Val(.TextMatrix(lngRow, lngCol + 1)) + intSign * Val(Nvl(mrsBalance!��Ԥ��, 0)), 5)
                
                .Cell(flexcpData, lngRow, lngCol + 1, lngRow, lngCol + 1) = Val(Nvl(mrsBalance!�Ƿ�����))
                If mbytMode = EM_RBDTY_�˷� Then
                    .ColData(lngCol) = "ժҪ:" & IIf(Val(Nvl(mrsBalance!�˷�)) = 1, "", Nvl(mrsBalance!ժҪ))
                    .ColData(lngCol + 1) = "�������:" & IIf(Val(Nvl(mrsBalance!�˷�)) = 1, "", Nvl(mrsBalance!�������))
                ElseIf mbytMode = EM_RBDTY_�쳣���� Then
                    .ColData(lngCol) = "ժҪ:" & IIf(Val(Nvl(mrsBalance!�˷�)) = 1, "", Nvl(mrsBalance!ժҪ))
                    .ColData(lngCol + 1) = "�������:" & IIf(Val(Nvl(mrsBalance!�˷�)) = 1, "", Nvl(mrsBalance!�������))
                Else
                    If mstrDelTime <> "" Then
                        .ColData(lngCol) = "ժҪ:" & IIf(Val(Nvl(mrsBalance!�˷�)) = 1, "", Nvl(mrsBalance!ժҪ))
                        .ColData(lngCol + 1) = "�������:" & IIf(Val(Nvl(mrsBalance!�˷�)) = 1, "", Nvl(mrsBalance!�������))
                    Else
                        .ColData(lngCol) = "ժҪ:" & IIf(Val(Nvl(mrsBalance!�˷�)) = 1, Nvl(mrsBalance!ժҪ), "")
                        .ColData(lngCol + 1) = "�������:" & IIf(Val(Nvl(mrsBalance!�˷�)) = 1, Nvl(mrsBalance!�������), "")
                    End If
                End If
            End If
             mrsBalance.MoveNext
            .Redraw = flexRDBuffered
         Loop
         
         '�ϲ�Ϊ��ļ�¼
         i = 1
         Do While i < .COLS - 1
            If .TextMatrix(lngRow, i + 1) = "" Then
                For lngCol = i To .COLS - 3
                    .TextMatrix(lngRow, lngCol) = .TextMatrix(lngRow, lngCol + 2)
                    .Cell(flexcpData, lngRow, lngCol) = .Cell(flexcpData, lngRow, lngCol + 2)
                    .ColData(lngCol) = .ColData(lngCol + 2)
                    .Cell(flexcpForeColor, lngRow, lngCol) = .Cell(flexcpForeColor, lngRow, lngCol + 2)
                    .Cell(flexcpForeColor, 1, lngCol) = .Cell(flexcpForeColor, 1, lngCol + 2)
                    .Cell(flexcpFontBold, 1, lngCol) = .Cell(flexcpFontBold, 1, lngCol + 2)
                Next
                If Trim(.TextMatrix(0, .COLS - 2)) = "" Then
                    .COLS = .COLS - 2
                Else
                  i = i + 2
                End If
            Else
                i = i + 2
            End If
         Loop
         .RowHidden(1) = False
         vsBalance.AutoSizeMode = flexAutoSizeColWidth
         Call vsBalance.AutoSize(0, .COLS - 1)
          ControlResize
    End With
End Sub
 
Private Sub LoadSelDelTotal()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ѡ����˿�ϼ�
    '����:���˺�
    '����:2014-10-11 11:41:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    txt�˿�ϼ� = Format(GetDelMoney, gstrDec)
End Sub

Private Function GetDelMoney() As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�˿�ϼ�
    '����:��ȡ�˿�ϼ�
    '����:���˺�
    '����:2014-07-03 17:24:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl�˿�ϼ� As Double, i As Long
    With vsBill
        For i = 1 To .Rows - 1
            If Val(vsBill.TextMatrix(i, .ColIndex("ѡ��"))) <> 0 Or _
                mbytMode = EM_RBDTY_�쳣���� Or mbytMode = EM_RBDTY_�鿴 Then
                dbl�˿�ϼ� = dbl�˿�ϼ� + Val(vsBill.TextMatrix(i, .ColIndex("ʵ�ս��")))
            End If
        Next
    End With
    GetDelMoney = RoundEx(dbl�˿�ϼ�, 6)
End Function

Private Sub ControlResize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����ؼ�λ��
    '����:���˺�
    '����:2014-10-11 11:42:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, blnFind As Boolean
    With vsBalance
        For i = 1 To .COLS - 1 Step 2
            If .Cell(flexcpData, 1, i) <> "" Then
                blnFind = True: Exit For
            End If
        Next
        If blnFind = False Then .RowHidden(1) = True
        .Height = IIf(.RowHidden(1), 375, 735)
    End With
    '85153,�ҺŲ�������˷�ʱ����"�˷�ժҪ"
    pic�˷�ժҪ.Visible = Not mCurBillType.bln�Һ�
    
    Form_Resize
End Sub

Private Sub txtPatient_Change()
    Dim blnAutoFind As Boolean
    blnAutoFind = False
    If Me.ActiveControl Is txtPatient And txtPatient.Visible Then
        blnAutoFind = txtPatient.Text = ""
    End If
    
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(blnAutoFind)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(blnAutoFind)
    IDKind.SetAutoReadCard (blnAutoFind)

End Sub

Private Sub txtPatient_GotFocus()
    If txtPatient.Locked Or Not txtPatient.Visible Then Exit Sub
    
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (txtPatient.Text = "")
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (txtPatient.Text = "")
    IDKind.SetAutoReadCard (txtPatient.Text = "" And txtPatient.Visible)
    zlControl.TxtSelAll txtPatient
    
End Sub
Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim blnCard As Boolean, blnICCard As Boolean, blnCancel As Boolean
    
    On Error GoTo errH
    
    If txtPatient.Locked Then Exit Sub

    If IDKind.GetCurCard.���� Like "����*" Then
        '103563,ֻҪ����ĵ�һ���ַ��ǡ�-+*����������ȫ���֣�����Ϊ����ˢ��
        If Not (InStr("-+*", Left(txtPatient.Text, 1)) > 0 And IsNumeric(Mid(txtPatient.Text, 2))) Then
            blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.ShowPassText)
        End If
    ElseIf IDKind.GetCurCard.���� = "�����" Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If InStr("0123456789-*+", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
        End If
    Else
        txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
        txtPatient.IMEMode = 0
    End If

    If Not (blnCard And Len(txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 _
        Or KeyAscii = 13 And Trim(txtPatient.Text) <> "") Then Exit Sub

    If KeyAscii <> 13 Then
        txtPatient.Text = txtPatient.Text & Chr(KeyAscii): txtPatient.SelStart = Len(txtPatient.Text)
    End If
    KeyAscii = 0
    Call FindPati(IDKind.GetCurCard, blnCard, Trim(txtPatient.Text))
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FindPati(ByVal objCard As Card, ByVal blnCard As Boolean, ByVal strInput As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���Ҳ���
    '����:���˺�
    '����:2014-09-30 14:29:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnICCard As Boolean, blnCancel As Boolean
    If objCard.���� Like "IC��*" And objCard.ϵͳ Then blnICCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
    
    'a.���������ȡ������Ϣʧ��
    If Not GetPatient(objCard, Trim(txtPatient.Text), blnCancel, blnCard) Then
        If blnCancel Then 'ȡ������
            If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
            txtPatient.Text = ""
            Exit Sub
        End If
        stbThis.Panels(2) = "δ�ҵ��ò��ˣ�������������!"
        If blnCard = True Then
            txtPatient.PasswordChar = "": txtPatient.Text = ""
            '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
            txtPatient.IMEMode = 0
        Else
            txtPatient.SelStart = 0: txtPatient.SelLength = Len(txtPatient.Text)
        End If
        Set mrsInfo = New ADODB.Recordset
        If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
        Call ClearFace
        Exit Sub
    End If
    mCurBillType.lng����ID = Val("" & mrsInfo!����ID)
    txtPatient = Nvl(mrsInfo!����)

    lblPati.Caption = "����:" & "                 " & _
        "���Ա�:" & Nvl(mrsInfo!�Ա�) & _
        "������:" & Nvl(mrsInfo!����) & _
        "�������:" & Nvl(mrsInfo!�����) & _
        "���ѱ�:" & Nvl(mrsInfo!�ѱ�) & _
        "�����ʽ:" & mrsInfo!ҽ�Ƹ��ʽ
        
    With mCurBillType
        .str�Ա� = Nvl(mrsInfo!�Ա�)
        .str���� = Nvl(mrsInfo!����)
        .str���� = Nvl(mrsInfo!����)
        .str�ѱ� = Nvl(mrsInfo!�ѱ�)
    End With
    If SelectNO(mCurBillType.lng����ID) = False Then
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Call ClearFace
        Exit Sub
    End If
    If vsBill.Enabled And vsBill.Visible Then vsBill.SetFocus
End Sub

Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, _
    blnCancel As Boolean, Optional blnCard As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ
    '���:objCard-ָ���Ŀ����
    '     strInput-�����ֵ
    '     blnCancel-
    '     blnCard-�Ƿ�ˢ��
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-09-30 14:30:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strPati As String, strSQL As String
    Dim vRect As RECT, i As Integer, lng�����ID As Long, bln�����ʻ� As Boolean, lng����ID As Long, strPassWord As String, strErrMsg As String
    Dim strWhere As String, blnICCard As Boolean
    Dim blnHavePassWord As Boolean
    
    blnCancel = False
    strWhere = ""
    If blnCard And objCard.���� Like "����*" And InStr("-+*", Left(strInput, 1)) = 0 Then  '103563
        lng�����ID = IDKind.GetDefaultCardTypeID
        '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�);��
        If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg, lng�����ID) = False Then GoTo NotFoundPati:
        If lng����ID <= 0 Then GoTo NotFoundPati:
        strWhere = strWhere & " And A.����ID=[1]"
        strInput = "-" & lng����ID
        blnHavePassWord = True
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then  '����ID
        strWhere = strWhere & " And A.����ID=[1]"
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then  'סԺ��(��ס(��)Ժ�Ĳ���)
        strWhere = strWhere & " And A.����ID = (Select Max(����id) From ������ҳ Where סԺ�� = [1])"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '�����(�������ﲡ��)
        strWhere = strWhere & " And A.�����=[1]"
        '75087,Ƚ����,2014-7-29,���ﲡ���շ�ʱ,����Ҫ���������������,ֻ��Ҫ��������ŵ����˳��ż����ҵ��������Ĳ�����Ϣ������
        strInput = "*" & zlCommFun.GetFullNO(Mid(strInput, 2), 3)
    Else '��������
        Select Case objCard.����
            Case "����", "��������￨"
                strPati = _
                " Select /*+Rule */A.����ID as ID,A.����ID,A.����,A.�Ա�,A.����," & _
                "           A.סԺ��,B.���� as ����,A.��ǰ���� as ����," & _
                "           A.��������,A.���֤��,A.��ͥ��ַ,A.����֤�� " & _
                " From ������Ϣ A,���ű� B" & _
                " Where A.ͣ��ʱ�� is NULL And A.��ǰ����ID=B.ID(+) And A.���� Like [1]" & _
                "   Order by A.����"
                vRect = zlControl.GetControlRect(txtPatient.hWnd)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strPati, 0, "���˲���", 1, "", "��ѡ����", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, False, True, strInput & "%", "bytSize=1")
                If Not rsTmp Is Nothing Then
                    strInput = rsTmp!����ID
                    strWhere = strWhere & " And A.����ID=[2]"
                Else
                    Set mrsInfo = New ADODB.Recordset: Exit Function
                End If
            Case "ҽ����"
                strInput = UCase(strInput)
                strWhere = strWhere & " And A.ҽ����=[2]"
            Case "���֤��", "�������֤", "���֤"
                strInput = UCase(strInput)
                If gobjSquare.objSquareCard.zlGetPatiID("���֤", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                strInput = "-" & lng����ID
                strWhere = strWhere & " And A.����ID=[1]"
            Case "IC����", "IC��"
                strInput = UCase(strInput)
                If gobjSquare.objSquareCard.zlGetPatiID("IC��", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                strInput = "-" & lng����ID
                strWhere = strWhere & " And A.����ID=[1]"
                blnICCard = (InStr(1, "-+*.", Left(strInput, 1)) = 0)
            Case "�����"
                If Not IsNumeric(strInput) Then strInput = "0"
                strWhere = strWhere & " And A.�����=[2]"
                strInput = zlCommFun.GetFullNO(strInput, 3)
            Case "סԺ��"
                If Not IsNumeric(strInput) Then strInput = "0"
                strWhere = strWhere & " And A.����ID = (Select Max(����id) From ������ҳ Where סԺ�� = [2])"
            Case Else
                '��������,��ȡ��صĲ���ID
                If objCard.�ӿ���� > 0 Then
                    lng�����ID = objCard.�ӿ����
                    bln�����ʻ� = objCard.�Ƿ�����ʻ� = 1
                    If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                    If lng����ID = 0 Then GoTo NotFoundPati:
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(objCard.����, strInput, False, lng����ID, _
                        strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                End If
                If lng����ID <= 0 Then GoTo NotFoundPati:
                strWhere = strWhere & " And A.����ID=[1]"
                strInput = "-" & lng����ID
                blnHavePassWord = True
        End Select
    End If
    strSQL = _
    " Select A.����ID,Nvl(C.��ҳID,0) as ��ҳID,A.�����,Nvl(C.��ǰ����ID,0) as ����ID,Nvl(c.��Ժ����ID,0) as ����ID,Nvl(A.��ǰ����ID,0) as ��ǰ����ID, Nvl(a.��Ժ,0) as ��Ժ," & _
    "           Decode(Nvl(A.��ҳID,0),0,A.ҽ�Ƹ��ʽ,C.ҽ�Ƹ��ʽ) ҽ�Ƹ��ʽ,Nvl(A.��������,C.��������) as ��������," & _
    "           A.����,A.�Ա�,A.����,Nvl(A.סԺ��,0) as סԺ��,Nvl(C.��Ժ����,0) as ����,A.��ͥ��ַ,A.����֤��," & _
    "           B.����,B.����,Nvl(B.ҽ����,A.ҽ����) ҽ����,B.����,Nvl(C.�ѱ�,A.�ѱ�) �ѱ�,A.������,A.������,Nvl(A.��������,0) as ��������, C.��ע " & _
    " From ������Ϣ A,ҽ�����˵��� B,������ҳ C,ҽ�����˹����� E " & _
    " Where A.ͣ��ʱ�� is NULL" & _
    "       And A.����ID=C.����ID(+) And Nvl(A.��ҳID,0)=C.��ҳID(+)" & _
    "       And C.����ID=E.����ID(+) And E.��־(+)=1  " & _
    "       And E.ҽ����=B.ҽ����(+) And E.����=B.����(+) And E.���� = B.����(+) " & strWhere

    On Error GoTo errH
    txtPatient.ForeColor = &HC00000: lblPati.ForeColor = txtPatient.ForeColor
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput)
    If mrsInfo.EOF Then
        Set mrsInfo = New ADODB.Recordset: Exit Function
    End If
    Call SetPatiColor(txtPatient, Nvl(mrsInfo!��������), IIf(IsNull(mrsInfo!����), &HC00000, vbRed))
    lblPati.ForeColor = txtPatient.ForeColor
    GetPatient = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
NotFoundPati:
    Set mrsInfo = New ADODB.Recordset
    txtPatient.ForeColor = &HC00000
    lblPati.ForeColor = txtPatient.ForeColor
End Function

Private Function SelectNO(ByVal lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݲ���IDѡ����ʵ��˷ѵ���
    '���:lng����ID-��ȡ����ID
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-09-30 14:47:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, blnCancel As Boolean
    Dim strNo As String, intFindType As Integer
    
    On Error GoTo errHandle
    '80602,Ƚ����,2014-12-8,����ȡ�������ϵĲ�����㵥�ݣ�����״̬=2��
    strSQL = "" & _
        "  With �շѵ� as ( " & _
        "           Select b.���㵥��,Max(a.ID) as ID,max(b.�������) as ����ID ,max(A.����ID) as ����ID, " & _
        "                  max(decode(a.��¼����,4,'�Һ�','�շ�')) as ����,  max(mod(a.��¼����,10)) as ��¼����ID,a.No as ���ݺ�,  c.���� as ��������, a.������, a.����Ա���, a.����Ա����, a.ʵ��Ʊ��, To_Char(a.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, " & vbCrLf & _
        "                  To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ�� " & vbCrLf & _
        "           From  ������ü�¼ A,( Select distinct  A.NO as ���㵥��, A.�������,A.�շѽ���ID,b.��¼����,b.NO,A.���ӱ�־" & _
        "                   From ���ò����¼ A,������ü�¼ B " & _
        "                   Where a.�շѽ���ID=b.����ID And A.����ID=[1] And nvl(a.����״̬,0)=0) B, " & _
        "                   ���ű� C " & vbCrLf & _
        "           Where  A.����ID=b.�շѽ���ID And nvl(A.���ӱ�־,0)<>9 and A.��������ID=C.ID(+)  And a.��¼״̬ in (1,3) " & vbCrLf & _
        "                And Nvl(a.ִ��״̬, 0) <> 1 And Nvl(a.����״̬, 0) <> 1 " & vbCrLf & _
        "              " & vbCrLf & _
        "          Group by b.���㵥��,mod(a.��¼����,10),a.No,a.������,c.����,a.����Ա���, a.����Ա����, a.ʵ��Ʊ��, To_Char(a.����ʱ��, 'yyyy-mm-dd hh24:mi:ss'),To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') " & vbCrLf & _
        "           )"

     strSQL = strSQL & vbCrLf & _
     "  Select J.*  " & vbCrLf & _
     "  From �շѵ� J," & vbCrLf & _
     "           ( Select mod(A.��¼����,10) as ��¼����,A.NO,sum(nvl(A.����,1)*nvl(A.����,1)) ����" & vbCrLf & _
     "             From ������ü�¼ A,�շѵ� B  " & vbCrLf & _
     "             Where A.NO=B.���ݺ� And mod(A.��¼����,10)= b.��¼����ID  And a.�۸񸸺� is null  " & vbCrLf & _
     "             Group by A.��¼����,A.NO " & vbCrLf & _
     "             Having sum(nvl(A.����,1)*nvl(A.����,1))>0 ) M" & vbCrLf & _
     "  Where J.���ݺ�=M.NO and J.��¼����ID=M.��¼���� " & vbCrLf
     strSQL = "Select * From (" & strSQL & ") Order by ��¼����ID,�Ǽ�ʱ�� desc,���ݺ�"
     
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "�˷ѵ���", 1, "", "��ѡ����Ҫ�˷ѵĵ���", False, False, False, 0, 0, 0, blnCancel, False, False, lng����ID, "bytSize=1")
    If blnCancel = True Then Exit Function
    If rsTemp Is Nothing Then
        MsgBox "�ò��˲����ڲ��������,���ڲ����շѹ����н����˷�", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If rsTemp.State <> 1 Then
        MsgBox "�ò��˲����ڲ��������,���ڲ����շѹ����н����˷�", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    Dim int��¼���� As Integer
    
    strNo = Nvl(rsTemp!���ݺ�): int��¼���� = Nvl(rsTemp!��¼����ID)
    mCurBillType.str���㵥 = Nvl(rsTemp!���㵥��)
    intFindType = IIf(int��¼���� = 4, 4, 1)
    
    If Not ReadBills(intFindType, strNo) Then
        Call ClearFace: Exit Function
    End If
    SelectNO = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub mobjICCard_ShowICCardInfo(ByVal strNo As String)
    Dim objCard As Card
    
    If strNo = "" Then Exit Sub
    If Not Me.ActiveControl Is txtPatient _
        Or txtPatient.Locked Or txtPatient.Text <> "" Then Exit Sub
        
    Set objCard = IDKind.GetIDKindCard("IC����", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    Call FindPati(objCard, False, strNo)
    If Not txtPatient.Locked And Me.ActiveControl Is txtPatient Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
End Sub
Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthday As Date, ByVal strAddress As String)
    Dim objCard As Card
    
    If strID = "" Then Exit Sub
    If Not Me.ActiveControl Is txtPatient _
        Or txtPatient.Locked Or txtPatient.Text <> "" Then Exit Sub
        
    Set objCard = IDKind.GetIDKindCard("���֤��", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    Call FindPati(objCard, False, strID)
    If Not txtPatient.Locked And Me.ActiveControl Is txtPatient Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
End Sub

  
Private Sub SynchronizationSelect(ByVal lngRow As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݵ�ǰ��,ͬ��ѡ�������
    '����:���˺�
    '����:2014-10-11 11:51:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    
    On Error GoTo errHandle
    
    With vsBill
        If Val(.Cell(flexcpData, lngRow, .ColIndex("��Ŀ"))) = 0 Then
            For i = lngRow + 1 To vsBill.Rows - 1
                 If Val(vsBill.RowData(lngRow)) = Val(vsBill.Cell(flexcpData, i, .ColIndex("��Ŀ"))) Then
                       vsBill.TextMatrix(i, .ColIndex("ѡ��")) = vsBill.TextMatrix(lngRow, .ColIndex("ѡ��"))
                 Else
                    Exit For
                 End If
            Next
            Call zlSet���ƹ̶���ϵ(lngRow, .ColIndex("ѡ��"))
            Exit Sub
        End If
        
        Call zlSet���ƹ̶���ϵ(lngRow, .ColIndex("ѡ��"))
        '��Ҫ��������Ƿ��Ѿ���
        For i = lngRow - 1 To 1 Step -1
            If Val(.RowData(i)) = Val(.Cell(flexcpData, lngRow, .ColIndex("��Ŀ"))) Then
                If .TextMatrix(i, .ColIndex("ѡ��")) <> 0 Then
                     .TextMatrix(i, .ColIndex("ѡ��")) = .TextMatrix(lngRow, .ColIndex("ѡ��"))
                End If
                Call zlSet���ƹ̶���ϵ(i, .ColIndex("ѡ��"), lngRow)
                 Exit For
            End If
        Next
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub
 
Public Function CheckDiff(strNos As String, strDiffNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Ƚ��������ݺ��Ƿ�һ��
    '����:ȫ��һ��,����true,���򷵻�False
    '����:���˺�
    '����:2014-03-21 17:19:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varTemp As Variant, varData As Variant, i As Long

    On Error GoTo errHandle
    varTemp = Split(Replace(strDiffNos, "'", ""), ",")
    varData = Split(Replace(strNos, "'", ""), ",")
    If UBound(varTemp) <> UBound(varData) Then Exit Function
    For i = 0 To UBound(varData)
        If InStr(1, "," & strDiffNos & ",", "," & varData(i) & ",") = 0 Then Exit Function
    Next
    For i = 0 To UBound(varTemp)
        If InStr(1, "," & strNos & ",", "," & varTemp(i) & ",") = 0 Then Exit Function
    Next
    CheckDiff = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub initInsurePara(ByVal intInsure As Integer, ByVal lng����ID As Long, ByVal lng����ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ҽ������
    '����:���˺�
    '����:2014-06-26 16:25:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If intInsure = 0 Then Exit Sub
    MCPAR.����������� = gclsInsure.GetCapability(support�����������, lng����ID, intInsure)
    MCPAR.ҽ���ӿڴ�ӡƱ�� = gclsInsure.GetCapability(supportҽ���ӿڴ�ӡƱ��, lng����ID, intInsure, CStr(lng����ID))
    MCPAR.�˷Ѻ��ӡ�ص� = gclsInsure.GetCapability(support�˷Ѻ��ӡ�ص�, lng����ID, intInsure)
    MCPAR.����Ԥ���� = gclsInsure.GetCapability(support����Ԥ��, lng����ID, intInsure)
    MCPAR.���Ը� = gclsInsure.GetCapability(support�շ��ʻ������Ը�, lng����ID, intInsure)
    MCPAR.ȫ�Ը� = gclsInsure.GetCapability(support�շ��ʻ�ȫ�Է�, lng����ID, intInsure)
    MCPAR.�ֱҴ��� = gclsInsure.GetCapability(support�ֱҴ���, lng����ID, intInsure)
End Sub

Private Sub SetFunCtrlVisible()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ù��ܿؼ���visible����
    '����:���˺�
    '����:2014-10-11 11:54:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    cmdSelAll.Visible = mbytMode = EM_RBDTY_�˷�
    cmdClear.Visible = mbytMode = EM_RBDTY_�˷�
    cmdBillSel.Visible = mbytMode = EM_RBDTY_�˷�
    If mstr������� <> "" Then   '���洫��ʱ,�����ֹ�����
        txtNO.Visible = False
        IDKindNO.Visible = False
        picPatiBack.Visible = False
        fraInfo_1.Visible = False
    End If
End Sub

Private Function GetYBOldBalance(ByVal lng����ID As Long, ByVal intInsure As Integer, ByVal lngԭ����ID As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡҽ��ԭ���㷽ʽ�ͽ�����
    '����:���ؽ�����Ϣ,��ʽ:���㷽ʽ|������||...
    '����:���˺�
    '����:2014-07-07 09:57:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str���㷽ʽ As String
    
    On Error GoTo errHandle
    If intInsure = 0 Then Exit Function
    
    '�ֶ�:����,��¼����,���㷽ʽ,ժҪ,�����ID,���������,���ƿ�,���㿨���,�������,����,������ˮ��, ����˵��,�������,У�Ա�־,ҽ��,���ѿ�id
    '     �Ƿ�����,�Ƿ�ȫ��,�Ƿ�����,��Ԥ��
    '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�)
    mrsBalance.Filter = "����=2 and ����ID=" & lngԭ����ID
    If mrsBalance.RecordCount = 0 Then Exit Function
    With mrsBalance
        Do While Not .EOF
            '������ֽ��㷽ʽ��֧�ֻ���,Ҫ��Ϊ�ֽ�,���ü�ȥ
            If MCPAR.����������� Then
                If gclsInsure.GetCapability(support�����������, lng����ID, intInsure, !���㷽ʽ) Then
                    str���㷽ʽ = str���㷽ʽ & "||" & !���㷽ʽ & "|" & Val(Nvl(!��Ԥ��))
                End If
            Else     '��֧�������������ʱ,ֻ���������Ϊ�ֽ�,����ԭ����,������ҽ������
                If !���㷽ʽ <> mstr�����ʻ� Then
                    str���㷽ʽ = str���㷽ʽ & "||" & !���㷽ʽ & "|" & Val(Nvl(!��Ԥ��))
                End If
            End If
            .MoveNext
        Loop
    End With
    If str���㷽ʽ <> "" Then str���㷽ʽ = Mid(str���㷽ʽ, 3)
    GetYBOldBalance = str���㷽ʽ
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function ExcuteInsureReCharge(ByVal lng����ID As Long, ByVal intInsure As Integer, _
    ByVal lng����ID As Long, ByVal lng������� As Long, ByVal strBalnaceInfor As String, _
    ByVal strNo As String, ByVal lng����ID As Long, ByVal strInvoice As String, ByVal dtDelDate As Date) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ��ҽ�������շ�
    '���:strBalnaceInfor:������Ϣ,��ʽΪ:ʵ�պϼ�;����ͳ��;ȫ�Ը�;����
    '����:ִ�гɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-10-11 11:55:07
    '˵��:����strNO,lng����ID,strInvoice,dtDelDate����ҽ���ӿڴ�ӡƱ��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strAdvance As String, arrBalance As Variant, str���㷽ʽ As String
    Dim dbl������ As Double, dbl�ɷ���� As Double, dbl��� As Double
    Dim strBalance As String, dbl�˿�ϼ� As Double, str�˻ؽ��� As String
    Dim strSQL As String, rsTmp As ADODB.Recordset, strYbInvoice As String
    Dim i As Long, k As Long, j As Long, cur����� As Double
    Dim strNone As String, strNos As String, varTemp As Variant, cur���� As Currency
    
    On Error GoTo errHandle
    If mCurBillType.intInsure = 0 Then
        ExcuteInsureReCharge = False
        gcnOracle.RollbackTrans: Exit Function
    End If
    
    
    strBalance = ""
    If Not MCPAR.����Ԥ���� Then '��������ʻ�֧�����
        varTemp = Split(strBalnaceInfor, ";") 'curʵ�պϼ�;cur����ͳ��;curȫ�Ը�;cur���Ը�
        If mstr�����ʻ� <> "" And mTy_Insure.dbl�ʻ���� > -1 * mTy_Insure.dbl����͸֧ Then
            If RoundEx(Val(varTemp(0)), 6) >= 0 Then
                cur���� = RoundEx(Val(varTemp(1)), 6) + IIf(MCPAR.���Ը�, RoundEx(Val(varTemp(3)), 6), 0) + IIf(MCPAR.ȫ�Ը�, RoundEx(Val(varTemp(2)), 6), 0)
                If mTy_Insure.dbl�ʻ���� - cur���� >= -1 * mTy_Insure.dbl����͸֧ Then
                    strBalance = mstr�����ʻ� & "|" & cur����   '������͸֧��Χ���㹻(����͸֧0Ϊ����)
                Else
                    If mTy_Insure.dbl����͸֧ = 0 And mTy_Insure.dbl�ʻ���� > 0 Then
                        strBalance = mstr�����ʻ� & "|" & mTy_Insure.dbl�ʻ����  '������͸֧�������
                    Else
                        '��������͸֧��Χ������͸֧ʱ�����
                        If mTy_Insure.dbl����͸֧ <> 0 Then
                            strBalance = mstr�����ʻ� & "|" & mTy_Insure.dbl�ʻ���� + mTy_Insure.dbl����͸֧ '������͸֧��Χ��֧��
                        Else
                            strBalance = mstr�����ʻ� & "|0"
                        End If
                    End If
                End If
            Else
                strBalance = mstr�����ʻ� & "|0"
            End If
        End If
    Else
        If ExecuteClinicPreSwap(intInsure, lng����ID, lng����ID, strBalance, strNone, strYbInvoice, strNos) = False Then
            gcnOracle.RollbackTrans
            If strNone <> "" Then
                MsgBox "��ǰ���ս���ʹ�õĽ��㷽ʽ" & vbCrLf & vbCrLf & vbTab & strNone & vbCrLf & vbCrLf & _
                    "������δ���ã����ȵ����㷽ʽ������������Щ���㷽ʽ��", vbInformation, gstrSysName
            End If
            Exit Function
        End If
    End If
    
    'Zl_���ò������_Modify
    strSQL = "Zl_���ò������_Modify("
    '  ��������_In   Number,
    '  --   0-��ͨ���㷽ʽ:
    '  --     ���㷽ʽ_IN:��������,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ||.." ;Ҳ�������.
    '  --   1.����������:
    '  --     �ٽ��㷽ʽ_IN:ֻ�ܴ���һ�����㷽ʽ,���������һЩ������Ϣ,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ"
    '  --     �ڿ����ID_IN,����_IN,������ˮ��_IN,����˵��_In:��Ҫ����
    '  --   2-ҽ������(�������ҽ���Ľ���,��Ҫ��ɾ��ԭҽ������,���´���ĸ���)
    '  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:���㷽ʽ|������||.."
    strSQL = strSQL & "" & 2 & ","
    '  ����id_In     In ���ò����¼.����id%Type,
    strSQL = strSQL & "" & lng����ID & ","
    '  ���㷽ʽ_In   Varchar2,
    strSQL = strSQL & "'" & strBalance & "')"
    '  �����id_In   ����Ԥ����¼.�����id%Type := Null,
    '  ����_In       ����Ԥ����¼.����%Type := Null,
    '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
    '  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
    '  �����_In   ������ü�¼.ʵ�ս��%Type := Null,
    '  ��ɽ���_In Number:=0
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    If MCPAR.ҽ���ӿڴ�ӡƱ�� And MCPAR.ҽ������Ʊ�� = False Then
        '38821,77058
        'Ʊ����������(��Ϊ����HIS�Ĵ�ӡ��ҽ���ӿڴ�ӡ����������Ʊ������)
        strSQL = "zl_�����շѼ�¼_RePrint('" & strNo & "','" & strInvoice & "'," & ZVal(lng����ID) & ",'" & UserInfo.���� & "'," & _
                  "To_Date('" & Format(dtDelDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),1,1)"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    End If
    
    '����ҽ������ӿ�
    If ExecuteClinicSwap(lng����ID, intInsure, lng����ID, lng�������, strBalance, strNos, strBalnaceInfor) = False Then Exit Function
    ExcuteInsureReCharge = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    Call ErrCenter
    Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, False, mCurBillType.intInsure)
End Function

Private Function ExecuteClinicPreSwap(ByVal intInsure As Integer, _
    ByVal lng����ID As Long, ByVal lng����ID As Long, ByRef strBalance As String, _
    ByRef strNone As String, ByRef strYbInvoice As String, ByRef strNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ������Ԥ����
    '���:intInsure-����
    '     lng����ID-�����շѵĽ���ID
    '����:strNone-�����ڵĽ��㷽ʽ
    '     strBalance-���ؽ��㷽ʽ(���㷽ʽ|���||...)
    '     strYbInvoice-ҽ�����صķ�Ʊ��
    '     strNOs-���ر��ν����NOs
    '����:Ԥ����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-07-07 11:18:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInvoice As String, varData As Variant
    Dim rsTemp As ADODB.Recordset, strAdvance As String
    Dim i As Long, str���㷽ʽ As String
    Dim varTemp As Variant
    
    
    On Error GoTo errHandle
    
    strInvoice = mCurBillType.strInvoice
    Set rsTemp = zlMakeClinicPreSwapData(strInvoice, lng����ID, strNos, True)
    
RePreSwap:
    strAdvance = "3": strBalance = ""
    If Not gclsInsure.ClinicPreSwap(rsTemp, strBalance, intInsure, strAdvance) Then
        Screen.MousePointer = 0
        If MsgBox("���½���ҽ���շ�ʱ,����Ԥ����ʧ��,�Ƿ����½���Ԥ����?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then GoTo RePreSwap:
        Exit Function
    End If
    If strAdvance <> "" And InStr(1, strAdvance, "|") = 0 Then 'ҽ��Ʊ�ݺ�
        strYbInvoice = strAdvance
    End If
    
    MCPAR.ҽ������Ʊ�� = False
    If InStr(1, strAdvance, ";") > 0 Then
        varData = Split(strAdvance & ";", ";")
        strYbInvoice = Trim(varData(0))
        '38821:strAdvance:��Ʊ��;�Ƿ���Ʊ�ݺ�
        MCPAR.ҽ������Ʊ�� = Val(varData(1)) = 1
    End If
    '���㷽ʽ;ԭʼ(���)���;�ɷ��޸�;�ĺ���
    varData = Split(strBalance, "|")
    
    '���㷽ʽ|������||..
    strBalance = "": strNone = ""
    For i = 0 To UBound(varData)
        varTemp = Split(varData(i), ";")
        str���㷽ʽ = varTemp(0)
        mrs���㷽ʽ.Filter = "����='" & str���㷽ʽ & "' And  ����>=3 and ����<= 4"
        If mrs���㷽ʽ.EOF Then
            strNone = strNone & "," & str���㷽ʽ
        End If
        strBalance = strBalance & "||" & varTemp(0) & "|" & Val(varTemp(1))
    Next
    If strBalance <> "" Then strBalance = Mid(strBalance, 3)
    If strNone <> "" Then
        strNone = Mid(strNone, 2): Exit Function
    End If
    
    ExecuteClinicPreSwap = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function ExecuteClinicSwap(ByVal lng����ID As Long, _
    ByVal intInsure As Integer, ByVal lng����ID As Long, _
    ByVal lng������� As Long, ByVal strԤ���� As String, _
    ByVal strNos As String, Optional ByVal strBalnaceInfor As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ҽ������ӿ�
    '���:  lng����ID:���ν��ʵ�ID
    '       strBalnaceInfor:������Ϣ,��ʽΪ:ʵ�պϼ�;����ͳ��;ȫ�Ը�;����
    '����:ҽ�����óɹ����ҽ��,����true,���򷵻�False
    '����:���˺�
    '����:2014-10-11 11:55:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnTrans As Boolean, blnTransMedicare As Boolean
    Dim cur����֧�� As Currency, curҽ������ As Currency
    Dim strSQL As String, strAdvance As String
    Dim varTemp As Variant
    Dim i As Long
    
    
    On Error GoTo errHandle
     
    blnTrans = True
    If MCPAR.ҽ���ӿڴ�ӡƱ�� And MCPAR.ҽ������Ʊ�� = False Then
        '���ϸ����Ʊ��ʱ���浱ǰƱ��
        If Not gblnStrictCtrl Then
            zlDatabase.SetPara "��ǰ�շ�Ʊ�ݺ�", mCurBillType.strInvoice, glngSys, mlngModule, zlStr.IsHavePrivs(mstrPrivs, "��������")
        End If
    End If
    
    
    cur����֧�� = 0: curҽ������ = 0
    If strԤ���� <> "" Then
        varTemp = Split(strԤ����, "||")
        For i = 0 To UBound(varTemp)
            If Split(varTemp(i), "|")(0) = mstr�����ʻ� Then
                cur����֧�� = cur����֧�� + CCur(Val(Split(varTemp(i), "|")(1)))
            ElseIf Split(varTemp(i), "|")(0) = "ҽ������" Then
                curҽ������ = curҽ������ + CCur(Val(Split(varTemp(i), "|")(1)))
            End If
        Next
    End If
    varTemp = Split(strBalnaceInfor, ";")  'curʵ�պϼ�;cur����ͳ��;curȫ�Ը�;cur���Ը�
    strAdvance = CStr(lng�������)
    If Not gclsInsure.ClinicSwap(lng����ID, cur����֧��, curҽ������, _
                        CCur(Val(varTemp(2))), CCur(Val(varTemp(3))), intInsure, strAdvance) Then
        gcnOracle.RollbackTrans:  Exit Function
    End If
  
    
    blnTransMedicare = True
    
    If strAdvance = CStr(lng�������) Then strAdvance = ""
     
    If strAdvance = "" Then
       gcnOracle.CommitTrans
       Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicSwap, True, intInsure)
       ExecuteClinicSwap = True: Exit Function
    End If
    
    If Not zlInsureCheck(strԤ����, strAdvance) Then
       gcnOracle.CommitTrans
       Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicSwap, True, intInsure)
       ExecuteClinicSwap = True: Exit Function
    End If
     'Zl_���ò������_Modify
        strSQL = "Zl_���ò������_Modify("
        '  ��������_In   Number,
        '  --   0-��ͨ���㷽ʽ:
        '  --     ���㷽ʽ_IN:��������,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ||.." ;Ҳ�������.
        '  --   1.����������:
        '  --     �ٽ��㷽ʽ_IN:ֻ�ܴ���һ�����㷽ʽ,���������һЩ������Ϣ,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ"
        '  --     �ڿ����ID_IN,����_IN,������ˮ��_IN,����˵��_In:��Ҫ����
        '  --   2-ҽ������(�������ҽ���Ľ���,��Ҫ��ɾ��ԭҽ������,���´���ĸ���)
        '  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:���㷽ʽ|������||.."
        strSQL = strSQL & "" & 2 & ","
        '  ����id_In     In ���ò����¼.����id%Type,
        strSQL = strSQL & "" & lng����ID & ","
        '  ���㷽ʽ_In   Varchar2,
        strSQL = strSQL & "'" & strAdvance & "')"
        '  �����id_In   ����Ԥ����¼.�����id%Type := Null,
        '  ����_In       ����Ԥ����¼.����%Type := Null,
        '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
        '  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
        '  �����_In   ������ü�¼.ʵ�ս��%Type := Null,
        '  ��ɽ���_In Number:=0
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
         
    gcnOracle.CommitTrans: blnTrans = False
    Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicSwap, True, intInsure)
    ExecuteClinicSwap = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    'ҽ����HIS����ͬһ������,HIS����ʧ��,��ҽ���������ϴ�,������Ҫ��"ȡ������"�ӿ�
    If blnTransMedicare Then Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicSwap, False, intInsure)
    Call SaveErrLog
End Function

Private Function ExecuteYBIdentifyCancel(ByVal lng����ID As Long, ByVal intInsure As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ȡ��ҽ�����������֤
    '����:���ؼ�ʱ���˳�������������
    '����:���˺�
    '����:2014-06-09 14:37:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    ExecuteYBIdentifyCancel = True
    If mbytMode = EM_RBDTY_�鿴 Or lng����ID = 0 Then Exit Function
    
    ExecuteYBIdentifyCancel = gclsInsure.IdentifyCancel(0, lng����ID, intInsure)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Public Sub zlExeBalanceWinRefrshData(ByVal blnSaveOK As Boolean, _
    ByRef objDelBalance As clsCliniDelBalance)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ���˷ѽ���������ˢ�²���
    '���:blnSaveOK-�Ƿ񱣴�ɹ�
    '     objChargeInfor-������Ϣ
    '����:���˺�
    '����:2014-06-17 10:50:41
    '˵��:֮��Ҫ��������,��Ҫԭ���ǽ��ҽ�����Ե�����(ģ̬���岻�õ���)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strPrintNos As String, strReclaimInvoice As String
    Dim strNo As String
    
    On Error GoTo errHandle
    
    If blnSaveOK = False Then Exit Sub
    
    strPrintNos = objDelBalance.PrintNOs
    
    If gblnBillPrint Then
        If gobjBillPrint.zlEraseBill(strPrintNos, 0) = False Then Exit Sub
    End If
   '��ӡ�˷ѵ���
    Call PrintDelBill(strPrintNos, objDelBalance.����ID, objDelBalance.�˷�ʱ��, objDelBalance.�����˷�, "")

Completed:
    mblnOK = True: Call ClearFace
    If txtNO.Visible Then txtNO.SetFocus: Exit Sub
    Unload Me
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function IsFeeAllDel() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ж��Ƿ�ȫ�˷�
    '����:���˷ѷ��سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-07-14 16:36:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnDel As Boolean, blnAllDel As Boolean
    Dim j As Long
    On Error GoTo errHandle
    '1.���Ƿ�Ϊȫѡ��ȫѡ��ԭ����
    If mCurBillType.bln���Ų����˷� Then Exit Function
    With vsBill
        For j = 1 To .Rows - 1
            If .TextMatrix(j, .ColIndex("���ݺ�")) <> "" And Abs(Val(.TextMatrix(j, .ColIndex("ѡ��")))) <> 1 Then
                IsFeeAllDel = False: Exit Function
            End If
        Next
    End With
    
    '2.��ǰ�˷��뱾���շѵ�����ȫһ��
    If CheckDiff(Replace(mCurBillType.strAllNOs, "'", ""), Replace(mCurBillType.strNos, "'", "")) = False Then Exit Function
    
    
    IsFeeAllDel = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetFeeDelNumRecord(ByVal strAllNOs As String) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���õ�ʣ��������
    '���:strAllNos-���е���
    '����:
    '����:��¼��(NO,���,ԭʼ����,ʣ������)
    '����:���˺�
    '����:2014-07-15 11:35:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    strSQL = "" & _
    "   Select A.NO,nvl(A.�۸񸸺�,A.���) as ���,a.�շ�ϸĿID,A.��¼����,A.����ID, " & _
    "         Decode(A.��¼����,1, 1,0)*decode(A.��¼״̬,1,1,3,1,0)*Avg(nvl(A.����,1) *����) as ԭʼ����," & _
    "         Avg(nvl(A.����,1) *����) as ����" & _
    "   From ������ü�¼ A" & _
    "   Where A.NO in (select J.Column_value From  Table(f_str2List([1])) J )  " & _
    "       And mod(a.��¼����,10)=1 And nvl(A.����״̬,0)<>1" & _
    "   Group by A.NO,nvl(A.�۸񸸺�,A.���),A.��¼����,A.��¼״̬,A.����ID,a.�շ�ϸĿID"
    
    strSQL = "" & _
    "   Select /*+ Rule */ A.NO,A.���,A.�շ�ϸĿID," & _
    "      sum(A.ԭʼ����/" & IIf(mtyMoudlePara.blnҩ����λ, "nvl(B." & gstrҩ����װ & ",1)", "1") & ") as ԭʼ����, " & _
    "      sum(A.����/" & IIf(mtyMoudlePara.blnҩ����λ, "nvl(B." & gstrҩ����װ & ",1)", "1") & ")  as ʣ������ " & _
    "   From (" & strSQL & ") A,ҩƷ��� B" & _
    "   Where A.�շ�ϸĿID=B.ҩƷID(+) " & _
    "   Group by A.NO,A.���,a.�շ�ϸĿID" & _
    "   Order by NO,���"

    On Error GoTo errHandle
    Set GetFeeDelNumRecord = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strAllNOs)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckIsAllDel(ByVal strAllNOs As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������з����Ƿ�ȫ��
    '���:strAllNos-���е���,����ö��ŷָ�
    '����:
    '����:����ȫ��ʱ,����true,���򷵻�False
    '����:���˺�
    '����:2014-07-15 11:28:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strNo As String, int��� As Integer
    Dim blnFind As Boolean, dblʣ������ As Double
    Dim j As Long, k As Long
    
    On Error GoTo errHandle
    If mbytMode = EM_RBDTY_�˷� Then
        With vsBill
            For j = 1 To vsBill.Rows - 1
                If Abs(Val(.TextMatrix(j, .ColIndex("ѡ��")))) <> 1 And InStr(strAllNOs, .TextMatrix(j, .ColIndex("���ݺ�"))) > 0 Then
                   CheckIsAllDel = False: Exit Function
                End If
            Next
        End With
    End If
    Set rsTemp = GetFeeDelNumRecord(strAllNOs)
    With rsTemp
        Do While Not .EOF
            strNo = Nvl(!NO): int��� = Val(Nvl(!���))
            dblʣ������ = Val(Nvl(!ʣ������))
            If dblʣ������ <> 0 Then
                With vsBill
                    k = vsBill.FindRow(strNo, , .ColIndex("���ݺ�"))
                    If k <= 0 Then Exit Function
                    blnFind = False
                    For j = k To vsBill.Rows - 1
                        If .TextMatrix(j, .ColIndex("���ݺ�")) <> strNo Then Exit For
                        If Abs(Val(.TextMatrix(j, .ColIndex("ѡ��")))) <> 1 _
                            And mbytMode <> EM_RBDTY_�쳣���� Then
                            CheckIsAllDel = False: Exit Function
                        End If
                        If Val(.RowData(j)) = int��� Then
                            If Val(dblʣ������) <> Val(.Cell(flexcpData, j, .ColIndex("����"))) Then
                               CheckIsAllDel = False: Exit Function
                            End If
                            blnFind = True: Exit For
                        End If
                    Next
                End With
                If blnFind = False Then Exit Function
            End If
            .MoveNext
        Loop
    End With
    CheckIsAllDel = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function ExecuteReDelFee() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���쳣���������˷�
    '����:�˷ѳɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-07-17 15:43:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim frmBalance As frmReplenishTheBalanceDelWin, objDelBalance As clsCliniDelBalance
    
    Dim blnȫ�� As Boolean, str����ID As String, lng����ID As Long, lng����ID As Long
    Dim strNos As String, varData As Variant, strCmdCaptions As String
    Dim cllPro  As New Collection, strReclaimInvoice As String, strInvoice As String
    Dim lngCheck����ID As Long, intCheckInsure   As Integer, strYBPati As String
    Dim dtDelDate As Date, blnTrans As Boolean, strNo As String
    Dim str��� As String, j As Long, strPrintNOInfor As String
    Dim strSQL As String, strBalanceInfor As String, cur����͸֧ As Currency
    Dim strReturn As String, strReturnRecipt As String '�˷Ѵ�����Ϣ����ʽ��NO,ҩ��ID|NO,ҩ��ID|��
    Dim rsҩƷ��¼ As ADODB.Recordset, lng����ID As Long
    Dim rsBalance As ADODB.Recordset
    
    On Error GoTo errHandle
    '�������
    If zlIsCheckExistErrBill(Val(mstr�������), True) = False Then
        MsgBox "��ǰ�쳣�����ѱ������㲻�ܼ�����", vbInformation, gstrSysName
        Exit Function
    End If
    If zlCheckOtherSessionDoing(Val(mstr�������)) Then
        MsgBox "��ǰ�����������������㴰���н��д����㲻�ܼ�����", vbInformation, gstrSysName
        Exit Function
    End If
    '���������㷽ʽ��Ч�Լ��
    Set rsBalance = zlFromIDGetChargeBalance(2, mCurBillType.strAllNOs, , , True, IIf(mCurBillType.bln�Һ�, 4, 1))
    If ThreeBalanceCheck(mobjPayCards, rsBalance, mcllForceDelToCash, mstr�ų����㷽ʽ) = False Then Exit Function
    
    If ShowReclaimInvoice(mCurBillType.str���㵥, strReclaimInvoice) = False Then Exit Function
    blnȫ�� = CheckIsAllDel(mCurBillType.strAllNOs)
    If Not blnȫ�� Then
        If MCPAR.ҽ���ӿڴ�ӡƱ�� Then
            If zlGetInvoiceGroupUseID(lng����ID) = False Then Exit Function
            strInvoice = GetNextBill(lng����ID)
        End If
    End If
    With vsBill
        str��� = "": strNo = ""
        For j = 1 To vsBill.Rows - 1
            If strNo <> Trim(.TextMatrix(j, .ColIndex("���ݺ�"))) Then
                If str��� <> "" Then
                    strPrintNOInfor = strPrintNOInfor & ";" & strNo & ":" & Mid(str���, 2)
                End If
                strNo = .TextMatrix(j, .ColIndex("���ݺ�"))
                str��� = ""
            End If
            str��� = str��� & "," & CLng(vsBill.RowData(j))
        Next
    End With
    
    Set objDelBalance = New clsCliniDelBalance
    'bytType-��������:0-���ݽ���ID����;1-���ݽ�����Ų���,2-���ݵ��ݺ�����ȡ���㷽ʽ
    Set objDelBalance.rsBalance = zlFromIDGetChargeBalance(1, Val(mstr�������), False)
    Set objDelBalance.rs���㷽ʽ = mrs���㷽ʽ
    
    lng����ID = mCurBillType.lng����ID
    lng����ID = mCurBillType.lng����ID
    dtDelDate = zlDatabase.Currentdate
    
    '����ҽ��
    If mCurBillType.intInsure <> 0 And lng����ID <> 0 Then
        '�����ҽ��,�����쳣,�϶���ֻ�����ղ��ֲų����쳣
        '�ֶ�:���� ,����ID, ��¼����, ���㷽ʽ, ժҪ, �����ID, ���������, ���ƿ�, ���㿨���, �������, ����, ������ˮ��, ����˵��, �������, У�Ա�־, ҽ��, ���ѿ�id
        '            �Ƿ�����,�Ƿ�ȫ��,�Ƿ�����,��Ԥ��
        '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
        mrsBalance.Filter = "����ID=" & lng����ID & " And ����=2 "
        If mrsBalance.EOF Then
            'δ����ҽ��Ԥ����,���,��Ҫ����Ԥ��,Ȼ�����
            '���ܴ��������շ�,���,��Ҫ���������֤�ӿ�(Identifiy)
            'strAdvace:ҽ��������ʱ:����1,��ʾҽ�������˺��������շѵ������֤;��������: ��
            lngCheck����ID = mCurBillType.lng����ID
            intCheckInsure = mCurBillType.intInsure
            strYBPati = gclsInsure.Identify(0, lngCheck����ID, intCheckInsure, 1)
            If strYBPati = "" Then
                 MsgBox "ҽ�������֤ʧ��,����������˷�!", vbOKOnly + vbDefaultButton1 + vbExclamation, gstrSysName
                 Exit Function
            End If
             
            If Val(CLng(Split(strYBPati, ";")(8))) <> mCurBillType.lng����ID Then
                MsgBox "ҽ����֤�Ĳ������˷ѵĲ��˲���ͬһ������!", vbInformation, gstrSysName
                Call ExecuteYBIdentifyCancel(mCurBillType.lng����ID, mCurBillType.intInsure)
                Exit Function
            End If
            
            If GetExcutInsureInforUpdateSQL(Val(mstr�������), strBalanceInfor, cllPro) = False Then Exit Function
            '��ȡ�������
            cur����͸֧ = mTy_Insure.dbl����͸֧
            mTy_Insure.dbl�ʻ���� = gclsInsure.SelfBalance(mCurBillType.lng����ID, CStr(Split(strYBPati, ";")(1)), 10, cur����͸֧, mCurBillType.intInsure)
            mTy_Insure.dbl����͸֧ = cur����͸֧
            
            
            blnTrans = True
            zlExecuteProcedureArrAy cllPro, Me.Caption, True
            
            '�ؽ����ҽ���ӿ�
            '77058
            If ExcuteInsureReCharge(mCurBillType.lng����ID, mCurBillType.intInsure, lng����ID, Val(mstr�������), strBalanceInfor, _
                        mCurBillType.str���㵥, lng����ID, strInvoice, dtDelDate) = False Then Exit Function
            blnTrans = False
        End If
    End If
    
    '4.��ʾ�������
    mCurBillType.lng������� = Val(mstr�������) '��¼���ڴ�ӡ��Ʊ
    With objDelBalance
        .intInsure = mCurBillType.intInsure
        .CurDelNos = mCurBillType.strNos
        .AllNos = mCurBillType.strAllNOs
        .PrintNOs = "'" & mCurBillType.str���㵥 & "'"
        .PatiUseType = mobjFact.ʹ�����
        .SaveBilled = True
        .ShareUserID = mobjFact.��������ID
        .����ID = mCurBillType.lng����ID
        .����ID = lng����ID
        .��ǰ��Ʊ�� = strInvoice
        .���շ�Ʊ = strReclaimInvoice
        .������� = Val(mstr�������)
        .����ID = lng����ID
        .ȱʡ���㷽ʽ = mCurBillType.str���㷽ʽ
        .�˷Ѻϼ� = -1 * GetDelMoney
        .�ѱ� = mCurBillType.str�ѱ�
        .���� = mCurBillType.str����
        .�Ա� = mCurBillType.str�Ա�
        .���� = mCurBillType.str����
        .�������� = mCurBillType.str��������
        .ҽ������Ʊ�� = MCPAR.ҽ������Ʊ��
        .ԭ����ID = mCurBillType.lngԭ����ID
        .�˷�ʱ�� = dtDelDate
        .�����˷� = Not blnȫ��
    End With
    
    Set frmBalance = New frmReplenishTheBalanceDelWin
    If frmBalance.zlChargeWin(Me, mlngModule, mstrPrivs, EM_BalanceReDel, mobjPayCards, objDelBalance, MCPAR.�ֱҴ���, _
        mcllForceDelToCash, mstr�ų����㷽ʽ, mCurBillType.bln�Һ�) = False Then Exit Function

    '81190,Ƚ����,�˷�ҵ����ҩ���ϴ��˷���Ϣ
    On Error Resume Next
    If Not mCurBillType.bln�Һ� Then
        If mblnDrugMachine Then
            Dim rsTemp As ADODB.Recordset, strData As String '���ﴦ����ҩ��ʽ������ID1,��ҩ����1;����ID2,��ҩ����2;...
            '�����˵ļ�ȥ���յľ���ʵ���˵�
            strSQL = "Select Max(Decode(a.��¼״̬, 2, a.Id, 0)) As ����id, -1 * Nvl(Sum(a.���� * a.����), 0) As ��ҩ����" & vbNewLine & _
                    " From ������ü�¼ A,(Select Distinct ����ID From ����Ԥ����¼ Where ������� = [1]) B" & vbNewLine & _
                    " Where a.����id = b.����ID And Mod(a.��¼����, 10) = 1 And a.�շ���� In ('5', '6', '7')" & vbNewLine & _
                    " Group By NO, Nvl(�۸񸸺�, ���)" & vbNewLine & _
                    " Having Nvl(Sum(a.���� * a.����), 0) <> 0"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ�����˷���Ŀ", objDelBalance.�������)
            Do While Not rsTemp.EOF
                strData = strData & ";" & Nvl(rsTemp!����id) & "," & Nvl(rsTemp!��ҩ����)
                rsTemp.MoveNext
            Loop
            If strData <> "" Then
                strData = Mid(strData, 2)
                Call mobjDrugMachine.Operation(gstrDBUser, Val("24-������ҩ(����/����)"), strData, strReturn)
            End If
        ElseIf mblnDrugPacker Then
            strSQL = "Select a.No, a.ִ�в���id" & _
                "   From ������ü�¼ A, ����Ԥ����¼ B" & _
                "   Where a.����id = b.����id And a.��¼״̬=2 And a.�շ���� In ('5', '6', '7') And b.������� = [1]"
            Set rsҩƷ��¼ = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mstr�������))
            
            If rsҩƷ��¼.RecordCount <> 0 Then
                Do While Not rsҩƷ��¼.EOF
                    If InStr(strReturnRecipt & "|", "|" & Nvl(rsҩƷ��¼!NO) & "," & Nvl(rsҩƷ��¼!ִ�в���ID) & "|") = 0 Then
                        strReturnRecipt = strReturnRecipt & "|" & Nvl(rsҩƷ��¼!NO) & "," & Nvl(rsҩƷ��¼!ִ�в���ID)
                    End If
                    rsҩƷ��¼.MoveNext
                Loop
            End If
    
            If strReturnRecipt <> "" Then
                strReturnRecipt = Mid(strReturnRecipt, 2)
                Call mobjDrugPacker.DYEY_MZ_TransRecipeReturn(1, UserInfo.���, UserInfo.����, strReturnRecipt, strReturn)
            End If
        End If
    End If
    Err.Clear: On Error GoTo errHandle
    

    
    If Not gfrmMain Is Nothing Then
        Call zlExeBalanceWinRefrshData(True, objDelBalance)
    End If
    ExecuteReDelFee = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckIsExistDelErrBill(ByVal strNos As String, Optional ByRef strOperatorName As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݵ��ݺ�,����Ƿ�����˷��쳣��¼
    '���:strNOs=���ݺ�,��ʽ NO1,NO2,NO3,...
    '����:
    '     strOperatorName=�����˷��쳣���ݵĲ���Ա����
    '����:�����˷��쳣����,����true,���򷵻�False
    '����:���˺�
    '����:2014-10-11 12:06:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    strOperatorName = ""
    If strNos = "" Then Exit Function
    
    On Error GoTo Errhand
    strSQL = "" & _
            " Select ����Ա����" & _
            " From ���ò����¼ A" & _
            " Where Nvl(����״̬, 0) = 1 And ��¼���� = 1 And ��¼״̬ = 2 " & _
            "       And a.No In (Select Column_Value From Table(f_Str2list([1])))" & _
            "       And Not Exists (Select 1 From ����Ԥ����¼ B Where a.����id = b.����id And Nvl(b.У�Ա�־, 0) = 0)" & _
            "       And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����Ƿ�����˷��쳣��¼", strNos)
    
    If Not rsTemp.EOF Then
        strOperatorName = Nvl(rsTemp!����Ա����)
        CheckIsExistDelErrBill = True
    End If
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
End Function
Private Function GetExcutInsureInforUpdateSQL(ByVal lng������� As Long, _
    ByRef strBalanceInfor As String, ByRef cllPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��Ҫ���µ����SQL
    '����:strBalanceInfor:��Ŀ������Ϣ(ͨ��GetItemInsure����),��ʽ: ʵ�պϼ�;����ͳ��;ȫ�Ը�;���Ը�
    '     cllPro-������Ҫִ�е�SQL
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:Ƚ����
    '����:2014-9-16
    '����:77951
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, strBXInfo As String
    Dim blnTrans As Boolean, curʵ�պϼ� As Currency, cur����ͳ�� As Currency, curȫ�Ը� As Currency, cur���Ը� As Currency
    Dim curʵ�ս�� As Currency, curͳ���� As Currency, bln������Ŀ As Boolean

    
    On Error GoTo Errhand
    strSQL = "" & _
    "   Select max(decode(a.��¼״̬,2,0,a.Id)) as ID," & _
    "          A.NO,a.����id, a.�շ�ϸĿid,a.���,a.������ĿID,sum(nvl(a.����,1)*a.����) as ����," & _
    "          Nvl(sum(a.ʵ�ս��), 0) As ʵ�ս��,max(decode(a.��¼״̬,2,'',a.ժҪ)) as ժҪ " & _
    "   From ������ü�¼ A,(Select distinct �շѽ���ID From ���ò����¼ Where �������=[1] ) B" & _
    "   Where a.����id =B.�շѽ���id " & _
    "   Group by  A.NO,a.����id,a.�շ�ϸĿid,a.���,a.������ĿID"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���շ��ü�¼", lng�������)
    
    With rsTemp
        If .RecordCount > 0 Then
            Set cllPro = New Collection
            Do While Not .EOF
                '������Ŀ��(0/1);���մ���ID;����ͳ����;������Ŀ����;ժҪ;��������
                strBXInfo = gclsInsure.GetItemInsure(Nvl(!����ID), Nvl(!�շ�ϸĿID), Val(Nvl(!ʵ�ս��)), True, mCurBillType.intInsure, _
                        Nvl(!ժҪ) & "||" & Val(Nvl(!����)))
                If strBXInfo <> "" Then
                    '  Zl_�����շѼ�¼_Update
                    strSQL = "Zl_�����շѼ�¼_Update("
                    '  Id_In         In ������ü�¼.Id%Type,
                    strSQL = strSQL & Nvl(!ID) & ","
                    '  ���մ���id_In In ������ü�¼.���մ���id%Type,
                    strSQL = strSQL & ZVal(Split(strBXInfo, ";")(1)) & ","
                    '  ������Ŀ��_In In ������ü�¼.������Ŀ��%Type,
                    strSQL = strSQL & Val(Split(strBXInfo, ";")(0)) & ","
                    '  ���ձ���_In   In ������ü�¼.���ձ���%Type,
                    strSQL = strSQL & "'" & CStr(Split(strBXInfo, ";")(3)) & "',"
                    '  ��������_In   In ������ü�¼.��������%Type,
                    strSQL = strSQL & "'" & CStr(Split(strBXInfo, ";")(5)) & "',"
                    '  ͳ����_In   In ������ü�¼.ͳ����%Type,
                    strSQL = strSQL & Format(Val(Split(strBXInfo, ";")(2)), gstrDec) & ","
                    '  ժҪ_In       In ������ü�¼.ժҪ%Type
                    strSQL = strSQL & "'" & CStr(Split(strBXInfo, ";")(4)) & "')"
                    zlAddArray cllPro, strSQL
                    
                    curͳ���� = CCur(Val(Split(strBXInfo, ";")(2)))
                    bln������Ŀ = Val(Split(strBXInfo, ";")(0)) = 1
                Else
                    curͳ���� = Val(Nvl(!ͳ����))
                    bln������Ŀ = Val(Nvl(!������Ŀ��)) = 1
                End If
                
                'ͳ�Ʊ��ս��
                curʵ�ս�� = Val(Nvl(!ʵ�ս��))
                If curͳ���� = 0 Or Not bln������Ŀ Then
                    '��ԭʼ���Ϊ׼,���ֱܷҴ���
                    curȫ�Ը� = curȫ�Ը� + curʵ�ս��
                Else
                    cur����ͳ�� = cur����ͳ�� + curͳ����
                    '��ԭʼ���Ϊ׼,���ֱܷҴ���
                    cur���Ը� = cur���Ը� + curʵ�ս�� - curͳ����
                End If
                curʵ�պϼ� = curʵ�պϼ� + CCur(Val(Nvl(!ʵ�ս��)))
                rsTemp.MoveNext
            Loop
        End If
    End With
    '���ս����Ϣ
    strBalanceInfor = curʵ�պϼ� & ";" & cur����ͳ�� & ";" & curȫ�Ը� & ";" & cur���Ը�
    GetExcutInsureInforUpdateSQL = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetChargeBalance(ByVal bytType As Byte, _
    ByVal strValue As String, Optional blnHistory As Boolean, _
    Optional ByRef blnDel As Boolean, Optional ByVal bln���쳣 As Boolean) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݽ���ID��ȡ�շѽ�����Ϣ
    '���:bytType-��������:0-���ݽ���ID����;1-���ݽ�����Ų���,2-���ݵ��ݺ�����ȡ���㷽ʽ
    '     strValue-Ҫ���ҵ�ֵ(Ϊ0ʱ,����ID,Ϊ1ʱ,�������,2ʱΪһ���շ����漰�����е���)
    '     blnDel-�˷ѽ���:true-���˷ѽ���;false-���˷ѽ���
    '     bln���쳣-�Ƿ�����쳣���㣬���ݵ��ݺ�����ȡ��������ʱ��Ч
    '����:�շѽ���������Ϣ��
    '       �ֶ�:����,��¼����,���㷽ʽ,ժҪ,�����ID,���������,���ƿ�,���㿨���,�������,����,������ˮ��, ����˵��,�������,У�Ա�־,ҽ��,���ѿ�id
    '            �Ƿ�����,�Ƿ�ȫ��,�Ƿ�����,��Ԥ��
    '       ����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
    '����:���˺�
    '����:2014-06-24 16:37:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strTable As String, strWhere As String
    Dim strTable1 As String
    On Error GoTo errHandle
    
    strTable = IIf(blnHistory, "H", "") & "����Ԥ����¼"
    Select Case bytType
    Case 0  '0-���ݽ���ID����
        strWhere = " And  A.����ID= [1]"
    Case 1  ';1-���ݽ�����Ų���
        strWhere = "  And A.�������= [1]"
    Case 2 '���ݵ��ݺ�����ȡ��������
        strTable1 = "" & _
        "   Select distinct �շѽ���ID as ����ID " & _
        "   From ���ò����¼ M " & _
        "   Where M.NO= [2] And Mod(M.��¼����,10)=1" & IIf(bln���쳣, "", " And Nvl(M.����״̬,0)<>1")
        strTable1 = strTable1 & " union ALL" & Replace(strTable1, "�շѽ���ID", "����ID")
        strTable1 = ",(" & strTable1 & ") Q1"
        If blnHistory Then strTable1 = Replace(strTable1, "���ò����¼", "H���ò����¼")
        strWhere = " And A.����ID=Q1.����ID"
    End Select
    
    If blnDel Then
        '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��))(��);5-���ѿ�(��)
        strSQL = "" & _
        "   Select  A.ID,decode(A.��¼״̬,2,A.����ID,NULL) as ����ID," & _
        "        Case when Mod(A.��¼����,10)=1 then 1  " & _
        "             when B.���� is not null then  2 " & _
        "             when nvl(A.�����ID,0)<>0  then  3 " & _
        "             when J.���㷽ʽ is not null   then  4 " & _
        "             else 0 end as ����, " & _
        "        Mod(A.��¼����,10) as ��¼����,A.���㷽ʽ,A.��Ԥ��," & _
        "        decode(A.��¼״̬,2,A.ժҪ,NULL) as ժҪ,decode(A.��¼״̬,2,1,0) as �˷�," & _
        "        A.�����ID,A.���㿨���, " & _
        "        decode(A.��¼״̬,2,A.�������,NULL) as �������,decode(A.��¼״̬,2,A.����,NULL) as ����, " & _
        "        decode(A.��¼״̬,2,A.������ˮ��,NULL) as ������ˮ��,nvl(C.�Ƿ�����,0) as ���ƿ�, " & _
        "        nvl(C.�Ƿ�����,0) as �Ƿ�����,nvl(C.�Ƿ�ȫ��,0) as �Ƿ�ȫ��, " & _
        "        Decode(C.��������,NULL,0,1) as  �Ƿ�����,Nvl(C.�Ƿ�ת�ʼ�����,0) as �Ƿ�ת�ʼ�����," & _
        "        C.���� as ���������,decode(A.��¼״̬,2,A.����˵��,NULL) as ����˵��,A.�������,decode(A.��¼״̬,2,A.У�Ա�־,0) as У�Ա�־, " & _
        "        decode(B.����,Null,0,1) as ҽ��,0 as ���ѿ�id,nvl(q.����,1) as ��������" & _
        "   From " & strTable & " A ,ҽ�ƿ���� C,һ��ͨĿ¼ J,���㷽ʽ q," & _
        "        (Select ���� From ���㷽ʽ where ���� in (3,4)) B " & strTable1 & _
        "   Where A.���㷽ʽ=J.���㷽ʽ(+) And A.�����ID=C.ID(+) " & _
        "         And A.���㷽ʽ=B.����(+) and A.���㷽ʽ=q.����(+) " & _
        "         And (a.��¼���� In (1, 11) Or Nvl(a.���㿨���, 0) = 0) " & strWhere
     
        strSQL = "" & _
        "   Select  max(����id) as ����id,����,max(�˷�) as �˷�,��¼����,���㷽ʽ,Max(ժҪ) as ժҪ,�����ID,���������,max(���ƿ�) as ���ƿ�,���㿨���, " & _
        "         max(�������) as �������,max(����) as ����,max(������ˮ��) as ������ˮ��, max(����˵��) as ����˵��, " & _
        "         �������,max(У�Ա�־) as У�Ա�־,ҽ��,���ѿ�id,��������,max(�Ƿ�ת�ʼ�����) as �Ƿ�ת�ʼ�����," & _
        "         max(�Ƿ�����) as �Ƿ�����,max(�Ƿ�ȫ��) as �Ƿ�ȫ��,max(�Ƿ�����) as �Ƿ����� , nvl(sum(��Ԥ��),0) as ��Ԥ��" & _
        "   From (" & strSQL & ") " & _
        "   Group by ����, ��¼����,���㷽ʽ,�����ID,���������,���㿨���,�������,ҽ��,���ѿ�id,�������� having  sum(��Ԥ��) <>0"
        Set GetChargeBalance = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�շѽ��㷽ʽ", Val(strValue), strValue)
        Exit Function
    End If
    
    '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��)(��);5-���ѿ�(��)
    strSQL = "" & _
    "   Select  A.ID,A.����ID," & _
    "        Case when Mod(A.��¼����,10)=1 then 1  " & _
    "             when B.���� is not null then  2 " & _
    "             when nvl(A.�����ID,0)<>0  then  3 " & _
    "             when J.���㷽ʽ is not null   then  4 " & _
    "             else 0 end as ����, " & _
    "        Mod(A.��¼����,10) as ��¼����,A.���㷽ʽ,A.��Ԥ��," & _
    "        A.ժҪ,decode(A.��¼״̬,2,1,0) as �˷�," & _
    "        A.�����ID,A.���㿨���, " & _
    "        A.�������,A.����,A.������ˮ��,nvl(C.�Ƿ�����,0) as ���ƿ�, " & _
    "        nvl(C.�Ƿ�����,0) as �Ƿ�����,nvl(C.�Ƿ�ȫ��,0) as �Ƿ�ȫ��, " & _
    "        Decode(C.��������,NULL,0,1) as  �Ƿ�����,Nvl(C.�Ƿ�ת�ʼ�����,0) as �Ƿ�ת�ʼ�����," & _
    "        C.���� as ���������,A.����˵��,A.�������,A.У�Ա�־, " & _
    "        decode(B.����,Null,0,1) as ҽ��,0 as ���ѿ�id,nvl(q.����,1) as ��������" & _
    "   From " & strTable & " A ,ҽ�ƿ���� C,һ��ͨĿ¼ J,���㷽ʽ q," & _
    "        (Select ���� From ���㷽ʽ where ���� in (3,4)) B " & strTable1 & _
    "   Where A.���㷽ʽ=J.���㷽ʽ(+) And A.�����ID=C.ID(+) " & _
    "         And A.���㷽ʽ=B.����(+) and A.���㷽ʽ=q.����(+) " & _
    "         And (a.��¼���� In (1, 11) Or Nvl(a.���㿨���, 0) = 0) " & strWhere
    
    gstrSQL = "" & _
    "   Select  ����ID,����,�˷�,��¼����,���㷽ʽ,ժҪ,�����ID,���������,���ƿ�,���㿨���,�������,����,������ˮ��, ����˵��,�������,У�Ա�־,ҽ��,���ѿ�id,��������," & _
    "         max(�Ƿ�ת�ʼ�����) as �Ƿ�ת�ʼ�����,max(�Ƿ�����) as �Ƿ�����,max(�Ƿ�ȫ��) as �Ƿ�ȫ��,max(�Ƿ�����) as �Ƿ����� , nvl(sum(��Ԥ��),0) as ��Ԥ��" & _
    "   From (" & gstrSQL & ") " & _
    "   Group by ����ID,����,�˷�,��¼����,���㷽ʽ,ժҪ,�����ID,���������,���ƿ�,���㿨���,�������,����,������ˮ��, ����˵��,�������,У�Ա�־,ҽ��,���ѿ�id,��������"
    Set GetChargeBalance = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�շѽ��㷽ʽ", Val(strValue), strValue)
    Exit Function
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function GetBalanceNO(ByVal intFindType As Integer, _
    ByVal strFindValue As String, _
    ByRef strNo As String, Optional bln�ҺŲ��� As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݽ�����Ż�ȡ���㵥�ݺ�
    '���:intFindType:��������(0-���ݽ�����Ų���;1-�����շѵ��Ų���;2-���ݷ�Ʊ��������;3-���ݹҺŵ���������;4-���ݽ��㵥�Ų���)
    '      strFindValue-intFindType=0:�������;intFindType=1:�շѵ���;intFindType=2:��Ʊ��;intFindType=3:�Һŵ�
    '����:strNo-���ؽ��㵥��
    '     bln�ҺŲ���-�Ƿ�ҺŲ������
    '����:��ȡ�ɹ�,����true,��ȡʧ�ܻ�δ�ҵ���������,����False
    '����:���˺�
    '����:2014-09-29 10:06:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim intMouse As Integer
    
    On Error GoTo errHandle
    
    If strFindValue = "" Then strFindValue = "0"
    
    Select Case intFindType
    Case 0 '���ݽ�����Ų���
        strSQL = "Select NO,���ӱ�־ From ���ò����¼ Where �������=[1] and rownum <2"
    Case 1 '�����շѵ��Ų���
        strSQL = "" & _
        "   Select    A.NO,A.���ӱ�־ From ���ò����¼ A,������ü�¼ B " & _
        "   Where A.�շѽ���ID=B.����ID And B.NO=[1] and mod(B.��¼����,10)=1 and rownum <2"
    Case 2 '���ݷ�Ʊ��������
        strSQL = "" & _
        "   Select  C.NO,C.���ӱ�־ From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ���� B,���ò����¼ C " & _
        "   Where ���� = [1] and B.NO=C.NO and mod(C.��¼����,10)=1 And b.�������� = 1 And A.��ӡid = b.Id and rownum<2"
    Case 3  '���ݹҺŵ���������
        strSQL = "" & _
        "   Select    A.NO,A.���ӱ�־ From ���ò����¼ A,������ü�¼ B " & _
        "   Where A.�շѽ���ID=B.����ID And B.NO=[1] and mod(B.��¼����,10)=4 and rownum <2"
    Case 4  '4-���ݽ��㵥�Ų���
        strSQL = "Select A.NO,A.���ӱ�־ From ���ò����¼ A Where A.NO=[1] and rownum <2"
    Case Else
        Exit Function
    End Select
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strFindValue)
    If Not rsTemp.EOF Then strNo = Nvl(rsTemp!NO): bln�ҺŲ��� = Val(Nvl(rsTemp!���ӱ�־)) = 1
    GetBalanceNO = True
    Exit Function
errHandle:
    
    intMouse = Me.MousePointer: Me.MousePointer = 0
    If ErrCenter() = 1 Then
        Me.MousePointer = intMouse
        Resume
    End If
    Me.MousePointer = intMouse
End Function

Private Function GetChargeInsure(ByVal str����ID As String, ByVal strNo As String, _
    ByRef lng����ID As Long, Optional ByVal blnNOMoved As Boolean) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�շѵ�ҽ����
    '���:lng����ID-����ID
    '     blnNOMoved-�Ƿ�����ת��
    '����:lng����ID-����ID
    '����:����
    '����:���˺�
    '����:2014-07-02 14:30:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim strWhere  As String
    
    On Error GoTo errHandle
    
    lng����ID = 0
    strWhere = " And A.����ID=[1]"
    If str����ID = "" Or str����ID = "0" Then strWhere = " And A.NO=[2]"
    If str����ID = "" Then str����ID = "0"
    
    strSQL = "" & _
    "    Select B.��¼ID,B.����,B.����ID " & _
    "    From ���ò����¼ A,���ս����¼ B " & _
    "    Where A.����ID=[1] And  mod(A.��¼����,10)=1 " & _
    "         And B.����=1 And A.����ID=B.��¼ID and Rownum<2 "
    If blnNOMoved Then
        strSQL = Replace(strSQL, "���ò����¼", "H���ò����¼")
        strSQL = Replace(strSQL, "���ս����¼", "H���ս����¼")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "������ID��ȡָ����ҽ������", str����ID, strNo)
    If rsTemp.EOF Then Exit Function
    lng����ID = Nvl(rsTemp!����ID, 0)
    GetChargeInsure = Nvl(rsTemp!����, 0)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function IsRegisterBalance(ByVal strNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�жϵ�ǰ���㵥���Ƿ�Һŵ��ݵĲ������
    '���:strNO-���㵥��
    '����:
    '����:�Һŵ��������,����true,���򷵻�False
    '����:���˺�
    '����:2014-10-08 16:19:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    strSQL = "" & _
    "   Select A.���ӱ�־ From ���ò����¼ A " & _
    "   where A.NO=[1] And mod(A.��¼����,10)=1   and rownum <2"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo)
    If rsTemp.EOF Then Exit Function
    IsRegisterBalance = Val(Nvl(rsTemp!���ӱ�־)) = 1
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetBalanceFeeNos(ByVal bytType As Byte, _
    ByVal strFindValue As String, _
    Optional ByRef strFeeNos As String, Optional ByRef strRegNos As String, _
    Optional ByVal blnNOMoved As Boolean) As Boolean
    
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����һ�Ž��㵥�ݵ�NO�����ID�������ţ�����ͬһ�ν�����շѵ���NOs
    '���:bytType-0-����NO������;1-���ݽ���ID������,2-���ݽ������������
    '    strFindValue-���ҵ�ֵ
    '    blnNOMoved-�Ƿ��ں󱸱��У���ѯ����֮ǰ���ж���Ҫ���������
    '����:strFeeNos-���ص�ǰ����ķ��õ���,��ʽ��"AAA,BBB,CCC',..."
    '     strRegNos-���ص�ǰ����ĹҺŵ���,��ʽ��"AAA,BBB,CCC',..."
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-06-17 17:06:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String, strNo As String
    On Error GoTo errHandle:
    Select Case bytType
    Case 0 '0-���ݽ���NO������
        strSQL = "" & _
        "   Select distinct A.NO,mod(A.��¼����,10) as ��¼����" & _
        "   From ������ü�¼ A," & _
        "        (Select distinct �շѽ���ID as ����ID From ���ò����¼ Where NO=[1] and ��¼����=1 ) B" & _
        "   Where A.����ID=B.����ID" & _
        "   Order by ��¼����,NO"
        
    Case 1  '1-���ݽ���ID������
        strSQL = "" & _
        "    Select Distinct A.No,mod(A.��¼����,10) as ��¼���� " & _
        "    From ������ü�¼ A," & _
        "        (Select distinct C1.�շѽ���ID as ����ID " & _
        "         From ���ò����¼ A1,���ò����¼ B1,���ò����¼ C1  " & _
        "         Where A1.����ID=[2] and A1.��¼����=1  " & _
        "               And A1.NO=B1.NO and A1.��¼����=B1.��¼���� " & _
        "               And B1.�������=C1.������� and C1.��¼״̬ in (1,3) ) B " & _
        "    Where A.����ID=B.����ID" & _
        "    Order By ��¼����,NO"
    Case 2  '2-���ݽ������������
        strSQL = "" & _
        "    Select Distinct A.No,mod(A.��¼����,10) as ��¼����" & _
        "    From ������ü�¼ A," & _
        "        (Select distinct C1.�շѽ���ID as ����ID " & _
        "         From ���ò����¼ A1,���ò����¼ B1,���ò����¼ C1  " & _
        "         Where A1.�������=[2] and A1.��¼����=1  " & _
        "               And A1.NO=B1.NO and A1.��¼����=B1.��¼���� " & _
        "               And B1.�������=C1.������� and C1.��¼״̬ in (1,3) ) B " & _
        "    Where A.����ID=B.����ID" & _
        "    Order By ��¼����,NO"
    End Select
    
    If blnNOMoved Then
        strSQL = Replace(strSQL, "������ü�¼", "H������ü�¼")
        strSQL = Replace(strSQL, "���ò����¼", "H���ò����¼")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡһ�ν������漰�ķ��õ���", strFindValue, Val(strFindValue))
    
    With rsTemp
        strFeeNos = "": strRegNos = ""
        Do While Not .EOF
            strNo = Nvl(!NO)
            If Val(Nvl(!��¼����)) = 1 Then
                If InStr(1, strFeeNos & ",", "," & strNo & ",") = 0 Then
                    strFeeNos = strFeeNos & "," & strNo
                End If
            Else
                If InStr(1, strRegNos & ",", "," & strNo & ",") = 0 Then
                    strRegNos = strRegNos & "," & strNo
                End If
            End If
            .MoveNext
        Loop
    End With
    If strFeeNos <> "" Then strFeeNos = Mid(strFeeNos, 2)
    If strRegNos <> "" Then strRegNos = Mid(strRegNos, 2)
    If strFeeNos = "" And strRegNos = "" Then Exit Function
    GetBalanceFeeNos = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetFeeListData(ByVal strNos As String, ByRef rsFeeList As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ǰ׼�˵ķ�����Ŀ
    '���:strNos-׼�˵���
    '����:rsFeeList-����׼�˷Ѽ�
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-10-08 17:49:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTableNo As String, strSQLIn As String
    Dim strSQL As String, strSqlSub As String
    
    strSqlSub = "" & _
        " Select /*+cardinality(j,10)*/ A.ID,A.��¼����,A.NO,A.��¼״̬,A.���,A.��������,A.�۸񸸺�,A.�շ�ϸĿID, " & _
        "        nvl(A.����,1) as ����, nvl(A.����,0) as ����, " & _
        "        nvl(A.Ӧ�ս��,0) as Ӧ�ս�� ,nvl(A.ʵ�ս��,0) as ʵ�ս��,nvl(A.���ʽ��,0) as ���ʽ��," & _
        "        Nvl(A.����,1)*A.���� as ����, nvl(��׼����,0)  as ��׼����," & _
                 IIf(mtyMoudlePara.blnҩ����λ, "nvl(B." & gstrҩ����װ & ",1)", "1") & " as ����ϵ��, " & _
                 IIf(mtyMoudlePara.blnҩ����λ, " decode(B.ҩƷID,NULL,A.���㵥λ,B." & gstrҩ����λ & ")", "A.���㵥λ ") & " as ���㵥λ," & _
        "        A.��������ID,A.ִ�в���ID,A.ҽ�����, " & _
        "        A.ִ��״̬,A.��������,A.����״̬ ,A.���ӱ�־,A.�ѱ�,A.�շ����,A.����Ա����,A.�Ǽ�ʱ��,A.����ID," & _
        "        B.ҩƷID" & _
        " From ������ü�¼ A,ҩƷ��� B,Table(f_Str2list([1])) J  " & _
        " Where mod(A.��¼����,10)=1 And A.NO=J.Column_Value and A.��¼״̬<>0" & _
        "       And A.�շ�ϸĿID=B.ҩƷID(+)"
    '��׼�˷�(����,ҩƷ,����������)
    strTableNo = _
        " With ������� as (" & strSqlSub & ")," & vbNewLine & _
        "      ׼���� as (Select /*+cardinality(j,10)*/ A.����ID," & _
        "                        Sum(Nvl(A.����,1)*A.ʵ������" & IIf(mtyMoudlePara.blnҩ����λ, "/Nvl(B." & gstrҩ����װ & ",1)", "") & ") as ׼������" & _
        "                 From ҩƷ�շ���¼ A,ҩƷ��� B, Table(f_Str2list([1])) J" & _
        "                 Where A.ҩƷID=B.ҩƷID(+) And Mod(A.��¼״̬,3)=1  " & _
        "                       And (A.���� =8 or a.����=24) And A.����� is NULL And A.NO =J.Column_Value" & _
        "                 Group by A.����ID"

    '��������ص�׼����
    '*��ҽ��ִ�мƼ��д�������ʱ,��ҽ��ִ�мƼ���ȡ��
    '*����ҽ������.ִ��״̬=1�����ִ�У�ʱ��׼����Ϊ0�����ٸ���ҽ��ִ�мƼ���ͳ��׼����,112447
    strTableNo = strTableNo & vbNewLine & _
        "   Union ALL " & vbNewLine & _
        "   Select Max(ID) As ����ID, Nvl(Sum(����), 0) As ׼����" & vbNewLine & _
        "   From(Select a.Id, a.ҽ����� As ҽ��id, a.�շ�ϸĿid, Decode(b.ִ��״̬, 1, 0, Decode(c.ִ��״̬, 0, 1, 0)) * c.���� As ����" & vbNewLine & _
        "        From (" & strSqlSub & ") A, ����ҽ������ B, ҽ��ִ�мƼ� C, ����ҽ����¼ M" & vbNewLine & _
        "        Where a.ҽ����� = b.ҽ��id And a.No = b.No And b.ҽ��id = c.ҽ��id And b.ҽ��ID = m.id" & vbNewLine & _
        "              And b.���ͺ� = c.���ͺ� And a.�շ�ϸĿid = c.�շ�ϸĿid + 0 And a.�۸񸸺� Is Null" & vbNewLine & _
        "              And a.��¼���� = 1 And a.��¼״̬ in (1, 3) And Instr(',5,6,7,', ',' || a.�շ���� || ',') = 0" & vbNewLine & _
        "              And Not Exists(Select 1 From �������� C Where a.�շ�ϸĿid = c.����id And c.�������� = 1)" & vbNewLine & _
        "              And Instr(',C,D,F,G,K,',','||m.�������||',')=0 And b.��¼���� = 1" & vbNewLine & _
        "        )" & vbNewLine & _
        "   Group By ҽ��ID, �շ�ϸĿID" & vbNewLine & _
        "   Having Max(ID) <> 0" & vbNewLine & _
        "  )"
    
    '���ŵ��ݻ��ܽ��(��ϸ���շ�ϸĿ)
    'ִ��״̬Ӧ����ԭʼ��¼���ж�(������ҩ�Ҳ����˷ѵļ�¼)
    '������������ʱ"��¼״̬,���"�ظ�,AVG������,����Ҫ��"ִ��״̬"
    '   *��ҽ��ִ�мƼ۵Ĳ����˷��޷��ж�׼���������������˷�
    strSQLIn = "" & _
        " Select NO, Nvl(�۸񸸺�, ���) As ���" & vbNewLine & _
        " From �������" & vbNewLine & _
        " Where ��¼���� = 1 And ��¼״̬ In (1, 3) And Nvl(ִ��״̬, 0) <> 1" & vbNewLine & _
        " Minus" & vbNewLine & _
        " Select NO, Nvl(�۸񸸺�, ���) As ���" & vbNewLine & _
        " From ������� A1" & vbNewLine & _
        " Where A1.��¼���� = 1 And A1.��¼״̬ In (1, 3) And Nvl(A1.ִ��״̬, 0) = 2" & vbNewLine & _
        "       And Not Exists(Select 1" & vbNewLine & _
        "                      From ����ҽ������ B, ҽ��ִ�мƼ� C" & vbNewLine & _
        "                      Where b.ҽ��id = A1.ҽ����� And b.No = A1.No" & vbNewLine & _
        "                            And b.ҽ��id = c.ҽ��id And b.���ͺ� = c.���ͺ�" & vbNewLine & _
        "                            And c.�շ�ϸĿid + 0 = A1.�շ�ϸĿid And b.��¼���� = 1)" & vbNewLine & _
        "       And Instr('5,6,7', A1.�շ����) = 0" & vbNewLine & _
        "       And Not Exists(Select 1 From �������� Where ����id = A1.�շ�ϸĿid And Nvl(��������, 0) = 1)"
    
    strSQL = _
    " Select A.NO,A.��¼״̬,A.��¼����,A.ִ��״̬,Nvl(A.�۸񸸺�,A.���) as ���,A.��������," & _
    "       A.�ѱ�,C.���� as �����,C.���� as �����,A.�շ�ϸĿID,B.����,B.����,B.���,Max(Nvl(A.��������,B.��������)) ��������," & _
    "       A.���㵥λ,Max(A.ҽ�����) as ҽ�����, " & _
    "       Avg(Nvl(A.����,1)) as ����,Avg(A.����/A.����ϵ��) as ����," & _
    "       Sum(A.��׼����*A.����ϵ��) as ����," & _
    "       Sum(A.Ӧ�ս��) as Ӧ�ս��,Sum(A.ʵ�ս��) as ʵ�ս��, " & _
    "       D.���� as ִ�п���,A.ִ�в���ID,E.���� as ��������" & _
    " From  ������� A,�շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D,���ű� E" & _
    " Where A.�շ�ϸĿID=B.ID And C.����=A.�շ����" & _
    "       And A.ִ�в���ID=D.ID(+) And A.��������ID=E.ID(+)" & _
    "       And (A.NO,Nvl(A.�۸񸸺�,A.���)) IN( " & strSQLIn & ")  " & _
    "       And A.NO IN( Select NO From ������� where  ��¼����=1 and ��¼״̬ in (1,3) )" & _
    " Group by A.NO,A.��¼����,A.��¼״̬,A.ִ��״̬,Nvl(A.�۸񸸺�,A.���),A.�ѱ�,A.��������," & _
    "       C.����,C.����,A.�շ�ϸĿID,B.����,B.����,B.���,A.���㵥λ," & _
    "       D.����,A.ִ�в���ID,E.����,A.ҩƷID,a.����ID "
     
    '��������
    '��"׼������=ԭʼ����"ʱ,�����ű���
    '�ſ��Ѿ�ȫ���˷ѵ���(ִ��״̬=0��һ�ֿ���)
    '��ʣ��������׼�������������������
        '1.�޶�Ӧ���շ���¼(����ͨ���û򲻸������õ�����),��ʱӦ��ʣ������
        '2.�շ���¼����ȫ������,����ȫ��ִ��,SQL���ų����ּ�¼
    strSQL = strTableNo & vbCrLf & _
    " Select A.NO,A.���,A.��������,A.�ѱ�,A.�����,A.�����,A.�շ�ϸĿID,A.����,A.����,A.���,Max(A.��������) As ��������,A.���㵥λ, Max(A.ҽ�����) as ҽ�����," & _
    "       Decode(Sign(Nvl(C.׼������,Sum(A.����*A.����))-B.ԭʼ����),0,avg(A.����),1) as ׼�˸���," & _
    "       Decode(Sign(Nvl(C.׼������,Sum(A.����*A.����))-B.ԭʼ����),0,Sum(A.����),Nvl(C.׼������,Sum(A.����*A.����))) as ׼������," & _
    "       Nvl(C.׼������,Sum(A.����*A.����)) as ׼������,Sum(A.����*A.����) as ʣ������," & _
    "       A.����,Sum(A.Ӧ�ս��) as ʣ��Ӧ��,Sum(A.ʵ�ս��) as ʣ��ʵ��,max(q1.��¼��־) as ��¼��־," & _
    "       A.ִ�п���,A.ִ�в���ID,A.��������,B.����Ա����,B.�Ǽ�ʱ��,B.����ID,Max(M.ҽ������) as ҽ������,b.ԭʼ����" & _
    " From (" & strSQL & ") A, ׼���� C,����ҽ����¼ M," & _
    "          ( Select  ID, NO,���, �շ�ϸĿID,Nvl( ����,0)/NVL(����ϵ��,1) as ԭʼ����,����Ա����,�Ǽ�ʱ��,����ID" & _
    "            From �������   " & _
    "            Where  ��¼״̬ IN(1,3) and ��¼����=1 And Nvl( ���ӱ�־,0)<>9 And  �۸񸸺� is NULL )B, " & _
    "            ( Select NO,Max(��¼״̬) as ��¼��־ From �������  Where ��¼״̬ in (1,3) Group by NO) Q1" & _
    " Where A.NO=B.NO And A.���=B.��� And A.�շ�ϸĿID=B.�շ�ϸĿID+0  And B.ID=C.����ID(+)" & _
    "            and A.ҽ�����=M.ID(+) and A.NO=q1.NO(+) " & _
    " Group by A.NO,A.���,A.��������,A.�ѱ�,A.�����,A.�����,A.�շ�ϸĿID,A.����,A.����,A.���," & _
    "       A.���㵥λ,A.����,B.ԭʼ����,C.׼������,A.ִ�п���,A.ִ�в���ID,A.��������,B.����Ա����,B.�Ǽ�ʱ��,B.����ID" & _
    " Having Sum(A.����*A.����)<>0"

    strSQL = _
    " Select A.NO,A.���,A.��������,A.�ѱ�,A.�����,A.�����,A.����,Nvl(B.����,A.����) as ����,E1.���� as ��Ʒ��," & _
    "       A.���,A.��������,A.���㵥λ,A.�շ�ϸĿID,A.׼�˸��� as ����,A.׼������ as ����,A.����, A.ҽ����� ," & _
    "       A.ʣ��Ӧ��*(A.׼������/A.ʣ������) as Ӧ�ս��," & _
    "       A.ʣ��ʵ��*(A.׼������/A.ʣ������) as ʵ�ս��," & _
    "       A.ִ�п���,A.ִ�в���ID,A.��������,A.����Ա����,A.�Ǽ�ʱ��,A.����ID,A.ҽ������,A.��¼��־, " & _
    "       A.ԭʼ����,A.׼������,A.ʣ������" & _
    " From (" & strSQL & ") A,�շ���Ŀ���� B,�շ���Ŀ���� E1" & _
    " Where A.�շ�ϸĿID=B.�շ�ϸĿID(+) And B.����(+)=1 And B.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
    "       And A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3 " & _
    " Order by A.NO,A.���"
    
    On Error GoTo errHandle
    Set rsFeeList = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNos)
    GetFeeListData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetRegListData(ByVal strNos As String, ByRef rsFeeList As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ǰ׼�˵ĹҺ���Ŀ
    '���:strNos-׼�˵���
    '����:rsFeeList-����׼�˷Ѽ�
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-10-08 17:49:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTableNo As String, strSQLIn As String
    Dim strSQL As String
    
    strTableNo = "" & _
    "   With  �������  as (" & _
    "           Select A.ID,A.��¼����,A.NO,A.��¼״̬,A.���,A.��������,A.�۸񸸺�,A.�շ�ϸĿID, " & _
    "                  nvl(A.����,1) as ����, nvl(A.����,0) as ����, " & _
    "                  nvl(A.Ӧ�ս��,0) as Ӧ�ս�� ,nvl(A.ʵ�ս��,0) as ʵ�ս��,nvl(A.���ʽ��,0) as ���ʽ��," & _
    "                  Nvl(A.����,1)*A.���� as ����, nvl(��׼����,0)  as ��׼����,1 as ����ϵ��, A.���㵥λ as ���㵥λ," & _
    "                  A.��������ID,A.ִ�в���ID,A.ҽ�����, " & _
    "                  A.ִ��״̬,A.��������,A.����״̬ ,A.���ӱ�־,A.�ѱ�,A.�շ����,A.����Ա����,A.����ID, " & _
    "                  A.�Ǽ�ʱ��,A.����ʱ��,E.ԤԼʱ��,E.����ʱ��,E.����ʱ�� as ����ʱ��,E.����,E.ִ���� as ҽ��," & _
    "                  Decode(E.����, Null, A.��ҩ����, To_Char(E.����)) as  ����,To_Char(E.�ű�)  as  ����  " & _
    "           From ������ü�¼ A, ���˹Һż�¼ E" & _
    "           Where mod(A.��¼����,10)=4 And A.NO IN (Select  Column_Value as No From Table(f_Str2list([1]))) " & _
    "                  And A.��¼״̬<>0  And A.NO=E.NO and E.��¼״̬ in (1,3)" & _
    "              )"
    
    strSQL = _
    " Select A.NO,A.��¼״̬,A.��¼����,A.ִ��״̬,Nvl(A.�۸񸸺�,A.���) as ���,A.��������," & _
    "       A.�ѱ�,A.�շ����,A.�շ�ϸĿID,A.��������," & _
    "       A.���㵥λ,Max(A.ҽ�����) as ҽ�����, " & _
    "       Avg(Nvl(A.����,1)) as ����,Avg(A.����/A.����ϵ��) as ����," & _
    "       Sum(A.��׼����*A.����ϵ��) as ����," & _
    "       Sum(A.Ӧ�ս��) as Ӧ�ս��,Sum(A.ʵ�ս��) as ʵ�ս��, " & _
    "       A.��������ID,A.ִ�в���ID" & _
    " From  ������� A" & _
    " Group by A.NO,A.��¼����,A.��¼״̬,A.ִ��״̬,Nvl(A.�۸񸸺�,A.���),A.�ѱ�,A.��������," & _
    "          A.�շ����,A.�շ�ϸĿID,A.��������,A.���㵥λ,A.��������ID,A.ִ�в���ID,a.����ID "
     
 
    strSQL = strTableNo & vbCrLf & _
    " Select A.NO,A.���,A.��������,A.�ѱ�,A.�շ����,A.�շ�ϸĿID,Max(A.��������) As ��������,A.���㵥λ, Max(A.ҽ�����) as ҽ�����," & _
    "       sum(a.����*A.����) as ׼������,sum(A.����*A.����) as ׼������,Sum(A.����*A.����) as ʣ������," & _
    "       max(A.����) as ����,Sum(A.Ӧ�ս��) as ʣ��Ӧ��,Sum(A.ʵ�ս��) as ʣ��ʵ��," & _
    "       max(decode(A.��¼״̬,2,0,A.��¼״̬))  as ��¼��־," & _
    "       max(A.��������ID) as ��������ID,max(A.ִ�в���ID) as ִ�в���ID, " & _
    "       max(B.����Ա����) as ����Ա����,max(B.ҽ��) as ҽ��,max(B.�Ǽ�ʱ��) as �Ǽ�ʱ��,max(B.����ʱ��) as ����ʱ��, " & _
    "       max(B.ԤԼʱ��) as ԤԼʱ��,max(B.����ʱ��) as ����ʱ��,max(B.����ʱ��) as ����ʱ��, " & _
    "       max(B.����) as ����,max(B.����) as ����,max(B.����) as ����, " & _
    "       max(B.����ID) as ����ID,max(b.ԭʼ����) as ԭʼ����" & _
    " From (" & strSQL & ") A," & _
    "      ( Select  ID, NO,���, �շ�ϸĿID,Nvl( ����,0)/NVL(����ϵ��,1) as ԭʼ����," & _
    "               ����Ա����,ҽ��, �Ǽ�ʱ��, ����ʱ��, ԤԼʱ��, ����ʱ��, ����ʱ�� as ����ʱ��, ����," & _
    "               ����,����,����ID" & _
    "        From �������   " & _
    "        Where  ��¼״̬ IN(1,3) and ��¼����=4 And Nvl( ���ӱ�־,0)<>9 And  �۸񸸺� is NULL ) B " & _
    " Where A.NO=B.NO And A.���=B.��� And A.�շ�ϸĿID=B.�շ�ϸĿID+0 " & _
    " Group by A.NO,A.���,A.��������,A.�ѱ�,A.�շ����, A.�շ�ϸĿID," & _
    "       A.���㵥λ" & _
    " Having Sum(A.����*A.����)<>0"
 
    strSQL = _
    " Select /*+ Rule */ A.NO,A.���,A.��������,A.�ѱ�,Q.���� as �����,Q.���� as �����,B1.����,Nvl(B.����,B1.����) as ����," & _
    "       Nvl(A.��������,B1.��������) ��������,A.���㵥λ,A.�շ�ϸĿID,A.׼������ as ����,A.����,A.ҽ�����," & _
    "       A.ʣ��Ӧ�� as Ӧ�ս��,A.ʣ��ʵ�� as ʵ�ս��," & _
    "       C1.���� as ִ�п���,A.ִ�в���ID,M.���� as ��������,A.����ID,A.��¼��־,  " & _
    "       A.����Ա����,A.ҽ��,A.�Ǽ�ʱ��,A.����ʱ��,A.ԤԼʱ��,A.����ʱ��,A.����ʱ��,A.����,A.����,A.����, " & _
    "       A.ԭʼ����,A.׼������,A.ʣ������" & _
    " From (" & strSQL & ") A,�շ���ĿĿ¼ B1,���ű� C1,���ű� M,�շ���Ŀ���� B, �շ���Ŀ��� Q" & _
    " Where A.�շ�ϸĿID=B1.ID And A.ִ�в���ID=C1.ID And A.��������ID=M.ID And A.�շ����=Q.���� And   " & _
    "       A.�շ�ϸĿID=B.�շ�ϸĿID(+) And B.����(+)=1 And B.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
    " Order by A.NO,A.���"
    On Error GoTo errHandle
    Set rsFeeList = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNos)
    GetRegListData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function ShowReclaimInvoice(ByVal strNos As String, ByRef strReclaimInvoice As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ�ͷ�����Ҫ���յķ�Ʊ
    '���:strNos-��ǰ�ĵ��ݺ�,����ö��ŷ���(����Ǳ������,��Ϊ������㵥��)
    '����:strReclaimInvoice-���ػ��յķ�Ʊ��(����ö��ŷָ�),��ʽ:AAAA,BBB,....)
    '����:��ʾ���ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-10-10 17:53:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim frmReInvoiceTemp As frmReInvoice
    
    On Error GoTo errHandle
    
    Set frmReInvoiceTemp = New frmReInvoice
    If frmReInvoiceTemp.ShowMe(Me, strNos, 0, 0, strReclaimInvoice, True) = False Then Exit Function
    If Not frmReInvoiceTemp Is Nothing Then Unload frmReInvoiceTemp
    Set frmReInvoiceTemp = Nothing
    ShowReclaimInvoice = True
    ShowReclaimInvoice = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function CheckSelectItemCanDel(ByVal strNos As String) As Boolean
    '���ܣ��ж�ѡ����˷���Ŀ�Ƿ���������˷ѣ���Ҫ��鲢���������е���Ŀ��������ݳ������ֱ�ִ����
    '������
    '   strNos - ����ѡ����˷ѵ��ݺ�
    '���أ�
    '   ���ͨ��������True�����򣬷���False
    '����ţ�105429
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim i As Long, j As Long, k As Long
    Dim arrNo As Variant
    Dim dblʣ������ As Double, dbl�������� As Double
    
    On Error GoTo errHandler
    If Left(strNos, 1) = "," Then strNos = Mid(strNos, 2)
    strNos = Replace(strNos, "'", "")
    If GetFeeListData(strNos, rsTemp) = False Then Exit Function
    If rsTemp.EOF Then
        MsgBox "����:" & strNos & " ��û�п��˷ѵ���Ŀ�������˷ѣ�", vbExclamation, gstrSysName
        Exit Function
    End If
    
    arrNo = Split(strNos, ",")
    For i = 0 To UBound(arrNo)
        With vsBill
            k = .FindRow(arrNo(i), , .ColIndex("���ݺ�"))
            For j = k To vsBill.Rows - 1
                If .TextMatrix(j, .ColIndex("���ݺ�")) <> arrNo(i) Then Exit For
                If Val(.TextMatrix(j, .ColIndex("ѡ��"))) <> 0 Then
                    rsTemp.Filter = "NO='" & arrNo(i) & "' And ���=" & .RowData(j)
                    If rsTemp.EOF Then
                        MsgBox "����:" & arrNo(i) & " �е� " & (j - k + 1) & " ����Ŀ��ʣ��δ������Ϊ�㣬�����˷ѣ�" & _
                            "�����»�ȡ�������ݣ�", vbExclamation, gstrSysName
                        .Row = j: .SetFocus
                        Exit Function
                    ElseIf Val(Nvl(rsTemp!ԭʼ����)) > 0 Then
                        '�����շѵĲ����
                        dblʣ������ = Val(Nvl(rsTemp!����, 1)) * Val(Nvl(rsTemp!����))
                        dbl�������� = Val(.TextMatrix(j, .ColIndex("����")))
                        If RoundEx(dbl��������, 6) > RoundEx(dblʣ������, 6) Then
                            MsgBox "����:" & arrNo(i) & " �е� " & (j - k + 1) & " ����Ŀ�ı����˷�����(" & _
                                FormatEx(dbl��������, 5) & ")������ʣ��δ������(" & FormatEx(dblʣ������, 5) & ")��" & _
                                "�����˷ѣ������»�ȡ�������ݣ�", vbExclamation, gstrSysName
                            .Row = j: .SetFocus
                            Exit Function
                        End If
                    End If
                End If
            Next
        End With
    Next
    CheckSelectItemCanDel = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetDelXMLExpend() As String
    '��ȡ�����������˷ѽӿ�zlRetuenCheck��strXMLExpend����ֵ
    If mbytMode = EM_RBDTY_�˷� Then
        GetDelXMLExpend = ZlGetDelXMLExpendByGrid(Me.vsBill)
    ElseIf mbytMode = EM_MULTI_�쳣���� Then
        GetDelXMLExpend = ZlGetDelXMLExpend(mstr�������, True)
    End If
End Function

