VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "*\A..\zlIDKind\zlIDKind.vbp"
Begin VB.Form frmChargeTurn 
   AutoRedraw      =   -1  'True
   Caption         =   "��(��)�����תסԺ"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   11715
   Icon            =   "frmChargeTurn.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8460
   ScaleWidth      =   11715
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picBill 
      Height          =   2100
      Left            =   90
      ScaleHeight     =   2040
      ScaleWidth      =   10485
      TabIndex        =   21
      Top             =   645
      Width           =   10545
      Begin VSFlex8Ctl.VSFlexGrid mshList 
         Height          =   1470
         Left            =   75
         TabIndex        =   22
         Top             =   90
         Width           =   5490
         _cx             =   9684
         _cy             =   2593
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
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483633
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
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
   End
   Begin VB.PictureBox picBalance 
      Height          =   1950
      Left            =   6285
      ScaleHeight     =   1890
      ScaleWidth      =   2985
      TabIndex        =   19
      Top             =   4035
      Width           =   3045
      Begin VSFlex8Ctl.VSFlexGrid vsBalance 
         Height          =   1335
         Left            =   0
         TabIndex        =   20
         Top             =   135
         Width           =   2565
         _cx             =   4524
         _cy             =   2355
         Appearance      =   3
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
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483633
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483633
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
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
      Begin VB.Label lblSum 
         AutoSize        =   -1  'True
         Caption         =   "ת���ϼ�:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   105
         TabIndex        =   23
         Top             =   1605
         Width           =   1155
      End
   End
   Begin VB.PictureBox picList 
      Height          =   1935
      Left            =   105
      ScaleHeight     =   1875
      ScaleWidth      =   5415
      TabIndex        =   17
      Top             =   3945
      Width           =   5475
      Begin VSFlex8Ctl.VSFlexGrid mshDetail 
         Height          =   1185
         Left            =   30
         TabIndex        =   18
         Top             =   165
         Width           =   5130
         _cx             =   9049
         _cy             =   2090
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
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483633
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
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
   End
   Begin VB.PictureBox picTop 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   11715
      TabIndex        =   7
      Top             =   0
      Width           =   11715
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "ˢ��(&R)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   8670
         TabIndex        =   15
         Top             =   95
         Width           =   1100
      End
      Begin VB.Frame fraPati 
         BorderStyle     =   0  'None
         Height          =   570
         Left            =   150
         TabIndex        =   10
         Top             =   -45
         Width           =   2910
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
            ForeColor       =   &H00C00000&
            Height          =   360
            Left            =   1200
            MaxLength       =   64
            TabIndex        =   11
            ToolTipText     =   "�ȼ���F11"
            Top             =   135
            Width           =   1650
         End
         Begin zlIDKind.IDKindNew IDKind 
            Height          =   345
            Left            =   570
            TabIndex        =   25
            Top             =   135
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   609
            Appearance      =   2
            IDKindStr       =   $"frmChargeTurn.frx":058A
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
         Begin VB.Label lblPatient 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   240
            Left            =   60
            TabIndex        =   12
            Top             =   180
            Width           =   480
         End
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   348
         Left            =   1152
         TabIndex        =   0
         Top             =   96
         Width           =   2664
         _ExtentX        =   4710
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   93323267
         CurrentDate     =   36588
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   348
         Left            =   4116
         TabIndex        =   16
         Top             =   96
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   93323267
         CurrentDate     =   36588
      End
      Begin VB.CheckBox chkShow 
         Caption         =   "����ʾ��ת������"
         Height          =   345
         Left            =   6870
         TabIndex        =   24
         Top             =   120
         Value           =   1  'Checked
         Width           =   1890
      End
      Begin VB.Label lbl�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Left            =   3840
         TabIndex        =   9
         Top             =   180
         Width           =   120
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ʱ��"
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
         Left            =   192
         TabIndex        =   8
         Top             =   168
         Width           =   960
      End
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H00808080&
      Height          =   435
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   11715
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   7665
      Width           =   11715
      Begin VB.CommandButton cmdParaSet 
         Caption         =   "��������(&R)"
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
         Left            =   3804
         TabIndex        =   14
         Top             =   0
         Width           =   1500
      End
      Begin VB.CommandButton cmdSave 
         Cancel          =   -1  'True
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
         Left            =   7545
         TabIndex        =   13
         Top             =   0
         Width           =   1100
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "����(&H)"
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
         Left            =   60
         TabIndex        =   4
         Top             =   0
         Width           =   1100
      End
      Begin VB.CommandButton cmdAll 
         Caption         =   "ȫ��(&C)"
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
         Index           =   1
         Left            =   2640
         TabIndex        =   2
         Top             =   0
         Width           =   1100
      End
      Begin VB.CommandButton cmdAll 
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
         Index           =   0
         Left            =   1455
         TabIndex        =   1
         Top             =   0
         Width           =   1100
      End
      Begin VB.CommandButton cmdExit 
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
         Left            =   8670
         TabIndex        =   3
         Top             =   0
         Width           =   1100
      End
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   8100
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmChargeTurn.frx":0620
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15584
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
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmChargeTurn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Private mstrNOS As String 'Ҫ���з���ת��ĵ�����Ϣ,��ʽ������,Ʊ��,����ID,����(��ҽ��Ϊ��),��������,������㵥�ݺ�:H0000001,F000023,81235,901,�շѵ�(���ʵ�),S0000001;...
Private mlngPatient As Long, mlng��ҳID As Long
Private msngOldY As Single, msngOldX As Single
Private Enum ����Enum
    Busi_Identify
    Busi_Identify2
    Busi_SelfBalance
    Busi_ClinicPreSwap
    Busi_ClinicSwap
    Busi_ClinicDelSwap
End Enum
Private Enum ҽԺҵ��
    support����������� = 33        'ҽ���Ƿ�֧������������ϣ���֧��ֻ�и������ʻ�ԭ����,�����ҽ�����㷽ʽ��Ϊ�ֽ�,֧�ֵ����ж�ÿһ�ֽ��㷽ʽ�Ƿ������˻�
    support�൥���շѱ���ȫ�� = 39  '�൥���շѱ���ȫ��
End Enum
Private Enum IDKinds
    C0���� = 0
    C1ҽ���� = 1
    C2���֤�� = 2
    C3IC���� = 3
    C4����� = 4
    C5סԺ�� = 5
    C6���￨ = 6
End Enum
Private WithEvents mobjIDCard As zlIDCard.clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1

Private mblnSelPati As Boolean '�Ƿ�ѡ����
Private mintPatientRange As Integer
Private mrsInfo As ADODB.Recordset
Private mlngTXTProc As Long
Private mstrPrivs As String, mlngModule As Long
Private mbln����תסԺ����� As Boolean
Private mbln�������� As Boolean
Private Enum mObjPancel
    Pan_Search = 1
    Pan_Bill = 2
    Pan_List = 3
    Pan_Balance = 4
    Pan_Bottom = 5
End Enum
Private mrsOneCard  As ADODB.Recordset

'�������ѿ��Ĵ������
Private Type Ty_SquareCard
    blnExistsObjects As Boolean     '��װ�����ѿ���
    rsSquare As ADODB.Recordset
    dblˢ���ܶ� As Double
    bln������ As Boolean '��ǰ��ȡ�ĵ����ǿ�����
    strˢ������ As String   'ˢ�����㷽ʽ;���;�Ƿ������޸�|..."
End Type
Private mtySquareCard As Ty_SquareCard
Private mstrThreeSwapBalance As String
Private mstrThreeSwapCardType As String
Private mstrThreeSwapMoney As String
Private mintIDKind As Integer
Private mobjSquare As Object
Private mblnPassInputCardNo As Boolean  '�Ƿ��������뿨��
Private mblnDefaultPassInputCardNo As Boolean 'ȱʡˢ���Ƿ��������뿨��
Private mlngҽ�ƿ�����  As Long
Private mblnNotClick As Boolean
Private mstrTitle As String

Private Sub InitPancel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������
    '����:���˺�
    '����:2011-03-25 17:37:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim sngWidth As Single, strReg As String, panThis As Pane
    Dim panTop As Pane, panBottom As Pane, panRight As Pane
    
    Set panTop = dkpMan.CreatePane(mObjPancel.Pan_Search, 200, 580, DockTopOf, Nothing)
    panTop.Title = "��������"
    panTop.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panTop.Tag = mObjPancel.Pan_Search
    panTop.Handle = picTop.hWnd
    panTop.MaxTrackSize.Height = 495 / Screen.TwipsPerPixelY
    panTop.MinTrackSize.Height = 495 / Screen.TwipsPerPixelY
    
    Set panThis = dkpMan.CreatePane(mObjPancel.Pan_Bill, 250, 580, DockBottomOf, panTop)
    panThis.Title = "����תסԺ�б�"
    panThis.Tag = mObjPancel.Pan_Bill
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    panThis.Handle = picBill.hWnd
    

    Set panRight = dkpMan.CreatePane(mObjPancel.Pan_Balance, 1500 / Screen.TwipsPerPixelX, 580, DockRightOf, panThis)
    panRight.Title = "����תסԺ������Ϣ"
    panRight.Tag = mObjPancel.Pan_Balance
    panRight.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    panRight.Handle = picBalance.hWnd
    
    Set panThis = dkpMan.CreatePane(mObjPancel.Pan_List, 250, 580, DockBottomOf, panThis)
    panThis.Title = "������ϸ�б�"
    panThis.Tag = mObjPancel.Pan_List
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    panThis.Handle = picList.hWnd
    
    dkpMan.Options.ThemedFloatingFrames = True
    dkpMan.Options.HideClient = True
End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case Pan_Search
        Item.Handle = picTop.hWnd
    Case Pan_Bill
        Item.Handle = picBill.hWnd
    Case Pan_List
        Item.Handle = picList.hWnd
    Case Pan_Balance
        Item.Handle = picBalance.hWnd
    End Select
End Sub

Public Sub ShowMe(objParent As Object, ByVal lngPatient As Long, ByRef strNos As String, _
    Optional blnSelPati As Boolean = False, Optional intPatientRange As Integer = 0, _
    Optional strPrivs As String, Optional lngModule As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������תסԺ����
    '���:lngPatient-����ID
    '      blnSelPati-�Ƿ���Ҫѡ����
    '      intPatientRange:(0-���в���,1-�κη���δ���岡��;2-���δ����Ĳ���;3-סԺδ����Ĳ���;4-����δ����Ĳ���)
    '����:
    '   strNOS:Ҫ���з���ת��ĵ�����Ϣ,��ʽ��
    '       ����,Ʊ��,����ID,����(��ҽ��Ϊ��),��������,������㵥�ݺ�:H0000001,F000023,81235,901,�շѵ�(���ʵ�),S0000001;...
    '����:
    '����:���˺�
    '����:2010-11-09 17:09:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mblnSelPati = blnSelPati: mintPatientRange = intPatientRange
    mlngPatient = lngPatient: mstrPrivs = strPrivs: mlngModule = lngModule
    mstrNOS = strNos
    
    If mblnSelPati = False Then
        '��ʱ������ʽ�����¼�Form_Load
        Call ShowBills(mlngPatient, dtpBegin.Value, dtpEnd.Value)
        Call SetBillSelected(strNos)
    Else
            If lngPatient <> 0 Then
                If GetPatient(IDKind.GetCurCard, "-" & lngPatient, 0) Then
                    Call ShowBills(mlngPatient, dtpBegin.Value, dtpEnd.Value)
                End If
            Else
                Call ClearData
            End If
    End If
    If mblnSelPati = False Then
        fraPati.Visible = False: cmdSave.Visible = True
    Else
        fraPati.Visible = True: cmdSave.Visible = True
    End If
    Call picTop_Resize
    Call Me.Show(vbModal, objParent)
    strNos = mstrNOS
End Sub

Private Sub ClearData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鵥��������ĸ����������˻ؿ����Ƿ���ȷ
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2010-11-09 17:30:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mshList.Redraw = flexRDNone
    mshList.Clear 1: mshList.Rows = 2
    sta.Panels(2).Text = ""
    Call setHeader: Call SetBillColor
    mshList.Redraw = flexRDBuffered
End Sub

Private Sub SetBillSelected(ByVal strNos As String)
'˵��:���ת�뼸���ʧ��,�ٽ���ѡ����,��ǰѡ������ѱ�ת��ĵ���������"����ת��",���Բ�Ӧ��ѡ��
    Dim i As Long
    With mshList
        For i = 1 To .Rows - 1
            If InStr(";" & strNos, ";" & .TextMatrix(i, .ColIndex("���ݺ�"))) > 0 And .TextMatrix(i, .ColIndex("���")) = "��ת��" Then
                .TextMatrix(i, .ColIndex("ѡ��")) = "��"
            Else
                .TextMatrix(i, .ColIndex("ѡ��")) = ""
            End If
        Next
    End With
End Sub

Public Function CheckExistTurn(ByVal lngPatient As Long, ByRef dat��Ժʱ�� As Date) As Boolean
'����:�����Ժʱ��֮���Ƿ����ת������
'����:ת�����ݵĵǼ�ʱ��
    Dim rsTmp As ADODB.Recordset, strSQL As String
        
    On Error GoTo errH
    strSQL = "" & _
    " Select Max(����ʱ��) ����ʱ�� " & _
    " From סԺ���ü�¼" & vbNewLine & _
    " Where ��¼���� = 2 And ��¼״̬ In(1,3) And ����id = [1] And ��ҳid Is Null And ��ʶ�� Is Null And �����־=2" & vbNewLine & _
    "       And ժҪ='�������ת��'"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "����Ƿ������ת����", lngPatient, dat��Ժʱ��)
    
    If Not IsNull(rsTmp!����ʱ��) Then
        dat��Ժʱ�� = rsTmp!����ʱ��
        CheckExistTurn = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub ExecteUpdate(ByVal lngPatient As Long, ByVal strסԺ�� As String, ByVal lngPageID As Long, ByVal dat��Ժʱ�� As Date)
'����:���¼��ʵ�����ҳID
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Zl_�������תסԺ_Update(" & lngPatient & "," & strסԺ�� & "," & lngPageID & _
            ",To_Date('" & Format(dat��Ժʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
    Call zlDatabase.ExecuteProcedure(strSQL, "���¼��ʵ�")
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function IsYBSingle(ByVal strno As String, ByVal intInsure As Integer) As Boolean
    Dim strSQL As String, rsTmp As ADODB.Recordset, blnInsureSingle As Boolean
    
    blnInsureSingle = gclsInsure.GetCapability(83, , intInsure)
    If blnInsureSingle = False Then
        IsYBSingle = False
        Exit Function
    Else
        strSQL = "Select 1 From ҽ��������ϸ Where NO = [1] And Rownum < 2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strno)
        If rsTmp.EOF Then
            IsYBSingle = False
        Else
            If CheckAllTurn(strno) Then
                IsYBSingle = False
            Else
                IsYBSingle = True
            End If
        End If
    End If
End Function

Public Function ExecuteTurn(ByVal frmMain As Object, ByVal lngModule As Long, ByVal strPrivs As String, _
    ByVal strNos As String, ByVal strסԺ�� As String, ByVal lng��ҳID As Long, _
    ByVal dat��Ժʱ�� As Date, ByVal lng��Ժ����ID As Long, ByVal lng��Ժ����ID As Long, _
    Optional ByRef strOutDelDate As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ָ���ĵ��ݺ�����,ִ���������תסԺ����,��ҽ���˷ѽ������
    '���:
    '   strNOS:Ҫ���з���ת��ĵ�����Ϣ,��ʽ��
    '       ����,Ʊ��,����ID,����(��ҽ��Ϊ��),��������,������㵥�ݺ�:H0000001,F000023,81235,901,�շѵ�(���ʵ�),S0000001;...
    '   lngסԺ��-סԺ��,lng��ҳID-��ҳID,��������������ҽ����Ժ����Ǽ�ʱ�Ŵ���
    '����:strDelDate-����ת������(Ŀǰ��Ҫ�����»�ȡԤ��������)
    '����:
    '����:���˺�
    '����:2011-02-16 10:26:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim DateDel As Date, arrNO As Variant, arrInfo As Variant
    Dim i As Long, j As Long, lngcnt As Long, blnҽ�������� As Boolean
    Dim strSQL As String, strInvoice As String, strInDate As String, strDelDate As String
    Dim cllJzPro As Collection, rsTemp As ADODB.Recordset, str��ת����ID As String
    Dim blnTrans As Boolean, blnTransMedicare As Boolean, blnExecuteThreeSwap As Boolean
    Dim intInsure As Integer, strAdvance As String, strJzNOs As String
    Dim rsDeposit As ADODB.Recordset, lng����ID As Long, blnTransMC As Boolean
    Dim str����˵�� As String, str������ˮ�� As String, blnTurnAll As Boolean
    
    '�������ĵ��ݴ���˼·���Ƚ����õ���תΪסԺ���ü�¼���ٵ������������˷�
    Dim strReplenishNo As String, strReplenishNos As String '��ʽ�����ݺ�,����
    Dim cllReplenishPro As Collection
    
    mstrPrivs = strPrivs: mlngModule = lngModule
    If strNos = "" Then Exit Function
    
    strInDate = "To_Date('" & Format(dat��Ժʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    strOutDelDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    strDelDate = "To_Date('" & strOutDelDate & "','YYYY-MM-DD HH24:MI:SS')"
    arrNO = Split(strNos, ";")
    Set cllJzPro = New Collection
    Set cllReplenishPro = New Collection
    
    On Error GoTo errH
    strJzNOs = ""
    i = LBound(arrNO)
    Do While i <= UBound(arrNO)
        lngcnt = 1
        strInvoice = Trim(Split(arrNO(i), ",")(1))
        If strInvoice <> "" Then
            For j = i + 1 To UBound(arrNO)
                If strInvoice = Split(arrNO(j), ",")(1) Then
                    lngcnt = lngcnt + 1
                Else
                    Exit For
                End If
            Next
        End If
        
        'ҽ��Ҫ������һ�ſ�ʼ��,�����������ǰ����ڵ������еģ����Դ˴����򼴿�
        For j = i To i + lngcnt - 1
            arrInfo = Split(arrNO(j), ",")
            blnҽ�������� = False: blnTurnAll = False
            
            strReplenishNo = arrInfo(5)
            If strReplenishNo = "" Then
                If Val(arrInfo(3)) <> 0 Then
                    blnҽ�������� = IsYBSingle(arrInfo(0), Val(arrInfo(3)))
                Else
                    blnTurnAll = CheckAllTurn(arrInfo(0))
                    If InStr("," & str��ת����ID & ",", "," & arrInfo(2) & ",") > 0 Then blnTurnAll = True
                End If
            
                If mbln�������� Then lng����ID = zlDatabase.GetNextId("���˽��ʼ�¼")
            End If
            
            If blnҽ�������� Or (Val(arrInfo(3)) = 0 And Not blnTurnAll) Or strReplenishNo <> "" Then
                'Zl_�������תסԺ_Insert
                strSQL = "Zl_�������תסԺ_insert("
                '  No_In         סԺ���ü�¼.NO%Type,
                strSQL = strSQL & "'" & arrInfo(0) & "',"
                '  סԺ��_In     סԺ���ü�¼.��ʶ��%Type, --ҽ����Ժ����Ǽ�ʱ�Ŵ���
                strSQL = strSQL & "" & IIf(strסԺ�� = "", "Null", strסԺ��) & ","
                '  ��ҳid_In     סԺ���ü�¼.��ҳid%Type, --ҽ����Ժ����Ǽ�ʱ�Ŵ���
                strSQL = strSQL & "" & IIf(lng��ҳID = 0, "Null", lng��ҳID) & ","
                '  ��Ժʱ��_In   סԺ���ü�¼.����ʱ��%Type,
                strSQL = strSQL & "" & strInDate & ","
                '  ��Ժ����id_In ����Ԥ����¼.����id%Type,
                strSQL = strSQL & "" & IIf(lng��Ժ����ID = 0, "NULL", lng��Ժ����ID) & ","
                '  �˷�ʱ��_In   סԺ���ü�¼.�Ǽ�ʱ��%Type, --���ŵ����˷�ʱ,ÿ�ŵ��ݵ��˷�ʱ����ͬ,����ϵͳ��ǰʱ��
                strSQL = strSQL & "" & strDelDate & ","
                '  ����Ա���_In סԺ���ü�¼.����Ա���%Type,
                strSQL = strSQL & "'" & UserInfo.��� & "',"
                '  ����Ա����_In סԺ���ü�¼.����Ա����%Type,
                strSQL = strSQL & "'" & UserInfo.���� & "',"
                '  ��Ժ����id_In סԺ���ü�¼.���˲���id%Type := Null,
                strSQL = strSQL & "" & IIf(lng��Ժ����ID = 0, "NULL", lng��Ժ����ID) & ","
                '  ����_In Number:=1(1-�����շѵ�;2-���ʵ�)
                strSQL = strSQL & "" & IIf(arrInfo(4) = "���ʵ�", 2, 1) & ","
                '  ����ID_In     סԺ���ü�¼.����id%Type,
                strSQL = strSQL & "" & IIf(mbln�������� And strReplenishNo = "", lng����ID, "NULL") & ","
                '  ԭ����id_In   סԺ���ü�¼.����id%Type := Null,
                strSQL = strSQL & "" & "NULL" & ","
                '  ��������_In Number:=1
                strSQL = strSQL & "" & IIf(mbln�������� And strReplenishNo = "", "1", "0") & ")"
                
                blnExecuteThreeSwap = False
                mstrThreeSwapBalance = ""
                mstrThreeSwapCardType = ""
                mstrThreeSwapMoney = ""
                
                If strReplenishNo <> "" Then
                    If InStr(strReplenishNos & ";", ";" & strReplenishNo & "," & arrInfo(3) & ";") = 0 Then
                        strReplenishNos = strReplenishNos & ";" & strReplenishNo & "," & arrInfo(3)
                    End If
                    cllReplenishPro.Add Array(strReplenishNo, strSQL)
                ElseIf arrInfo(4) = "���ʵ�" And mbln�������� Then
                    If InStr(strJzNOs & ",", "," & arrInfo(0) & ",") = 0 Then
                        strJzNOs = strJzNOs & "," & arrInfo(0)
                        cllJzPro.Add strSQL, "K" & arrInfo(0)
                    End If
                Else
                    gcnOracle.BeginTrans: blnTrans = True
                
                    Call zlDatabase.ExecuteProcedure(strSQL, "�������תסԺ")
                    '����ҽ��
                    blnTransMedicare = False
                    intInsure = IIf(arrInfo(4) = "���ʵ�", 0, Val(arrInfo(3)))
                    If mbln�������� = False Then intInsure = 0  'ֻ���������ʵ�,�Ż�ȥ����ҽ���ӿ�
                    If intInsure <> 0 Then
                        strAdvance = lng����ID & "|0|" & arrInfo(0)
                        If Not gclsInsure.ClinicDelSwap(Abs(Val(arrInfo(2))), , intInsure, strAdvance) Then
                            gcnOracle.RollbackTrans
                            MsgBox "ҽ������ʧ�ܣ��޷������������תסԺ������", vbInformation, gstrSysName
                            Exit Function
                        Else
                            blnTransMedicare = True
                        End If
                    End If
                    gcnOracle.CommitTrans: blnTrans = False
                    If blnTransMedicare And mbln�������� Then Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, True, intInsure)
                    
                    If mbln�������� And arrInfo(4) <> "���ʵ�" Then
                        If ExecuteThreeSwap(Val(arrInfo(2)), lng����ID, str������ˮ��, str����˵��) = True Then
                            blnExecuteThreeSwap = True
                        End If
                        'Zl_����תסԺ_����������
                        strSQL = "Zl_����תסԺ_����������("
                        '  No_In         סԺ���ü�¼.NO%Type,
                        strSQL = strSQL & "'" & arrInfo(0) & "',"
                        '  ����Ա���_In סԺ���ü�¼.����Ա���%Type,
                        strSQL = strSQL & "'" & UserInfo.��� & "',"
                        '  ����Ա����_In סԺ���ü�¼.����Ա����%Type,
                        strSQL = strSQL & "'" & UserInfo.���� & "',"
                        '  �˷�ʱ��_In   סԺ���ü�¼.�Ǽ�ʱ��%Type, --���ŵ����˷�ʱ,ÿ�ŵ��ݵ��˷�ʱ����ͬ,����ϵͳ��ǰʱ��
                        strSQL = strSQL & "" & strDelDate & ","
                        '  �����˷�_In   Number := 0,
                        strSQL = strSQL & "" & 0 & ","
                        '  ��Ժ����id_In ����Ԥ����¼.����id%Type,
                        strSQL = strSQL & "" & IIf(lng��Ժ����ID = 0, "NULL", lng��Ժ����ID) & ","
                        '  ��ҳid_In     סԺ���ü�¼.��ҳid%Type, --ҽ����Ժ����Ǽ�ʱ�Ŵ���
                        strSQL = strSQL & "" & IIf(lng��ҳID = 0, "Null", lng��ҳID) & ","
                        '  �����˷�_In   Number := 0,
                        strSQL = strSQL & "" & IIf(blnExecuteThreeSwap = True, 1, 0) & ","
                        '  ����ID_In     סԺ���ü�¼.����id%Type)
                        strSQL = strSQL & "" & lng����ID & ")"
                        Call zlDatabase.ExecuteProcedure(strSQL, "����������")
                    End If
                End If
            Else
                If InStr("," & str��ת����ID & ",", "," & arrInfo(2) & ",") = 0 Then
                    'Zl_�������תסԺ_Insert
                    strSQL = "Zl_�������תסԺ_insert("
                    '  No_In         סԺ���ü�¼.NO%Type,
                    strSQL = strSQL & "'" & arrInfo(0) & "',"
                    '  סԺ��_In     סԺ���ü�¼.��ʶ��%Type, --ҽ����Ժ����Ǽ�ʱ�Ŵ���
                    strSQL = strSQL & "" & IIf(strסԺ�� = "", "Null", strסԺ��) & ","
                    '  ��ҳid_In     סԺ���ü�¼.��ҳid%Type, --ҽ����Ժ����Ǽ�ʱ�Ŵ���
                    strSQL = strSQL & "" & IIf(lng��ҳID = 0, "Null", lng��ҳID) & ","
                    '  ��Ժʱ��_In   סԺ���ü�¼.����ʱ��%Type,
                    strSQL = strSQL & "" & strInDate & ","
                    '  ��Ժ����id_In ����Ԥ����¼.����id%Type,
                    strSQL = strSQL & "" & IIf(lng��Ժ����ID = 0, "NULL", lng��Ժ����ID) & ","
                    '  �˷�ʱ��_In   סԺ���ü�¼.�Ǽ�ʱ��%Type, --���ŵ����˷�ʱ,ÿ�ŵ��ݵ��˷�ʱ����ͬ,����ϵͳ��ǰʱ��
                    strSQL = strSQL & "" & strDelDate & ","
                    '  ����Ա���_In סԺ���ü�¼.����Ա���%Type,
                    strSQL = strSQL & "'" & UserInfo.��� & "',"
                    '  ����Ա����_In סԺ���ü�¼.����Ա����%Type,
                    strSQL = strSQL & "'" & UserInfo.���� & "',"
                    '  ��Ժ����id_In סԺ���ü�¼.���˲���id%Type := Null,
                    strSQL = strSQL & "" & IIf(lng��Ժ����ID = 0, "NULL", lng��Ժ����ID) & ","
                    '  ����_In Number:=1(1-�����շѵ�;2-���ʵ�)
                    strSQL = strSQL & "" & IIf(arrInfo(4) = "���ʵ�", 2, 1) & ","
                    '  ����ID_In     סԺ���ü�¼.����id%Type)
                    strSQL = strSQL & "" & IIf(mbln��������, lng����ID, "NULL") & ","
                    '  ԭ����ID_In     סԺ���ü�¼.����id%Type,
                    strSQL = strSQL & "" & arrInfo(2) & ","
                    '  ��������_In Number:=1
                    strSQL = strSQL & "" & IIf(mbln��������, "1", "0") & ")"
                    
                    blnExecuteThreeSwap = False
                    mstrThreeSwapBalance = ""
                    mstrThreeSwapCardType = ""
                    mstrThreeSwapMoney = ""
                    
                    If arrInfo(4) = "���ʵ�" And mbln�������� Then
                        If InStr(strJzNOs & ",", "," & arrInfo(0) & ",") = 0 Then
                            strJzNOs = strJzNOs & "," & arrInfo(0)
                            cllJzPro.Add strSQL, "K" & arrInfo(0)
                        End If
                    Else
                        gcnOracle.BeginTrans: blnTrans = True
                        Call zlDatabase.ExecuteProcedure(strSQL, "�������תסԺ")
                        '����ҽ��
                        blnTransMedicare = False
                        intInsure = IIf(arrInfo(4) = "���ʵ�", 0, Val(arrInfo(3)))
                        If mbln�������� = False Then intInsure = 0  'ֻ���������ʵ�,�Ż�ȥ����ҽ���ӿ�
                        If intInsure <> 0 Then
                            strAdvance = lng����ID & "|0"
                            If Not gclsInsure.ClinicDelSwap(Abs(Val(arrInfo(2))), , intInsure, strAdvance) Then
                                gcnOracle.RollbackTrans
                                MsgBox "ҽ������ʧ�ܣ��޷������������תסԺ������", vbInformation, gstrSysName
                                Exit Function
                            Else
                                blnTransMedicare = True
                            End If
                        End If
                        gcnOracle.CommitTrans: blnTrans = False
                        If blnTransMedicare And mbln�������� Then Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, True, intInsure)
                        
                        If mbln�������� And arrInfo(4) <> "���ʵ�" Then
                            If ExecuteThreeSwap(Val(arrInfo(2)), lng����ID, str������ˮ��, str����˵��) = True Then
                                blnExecuteThreeSwap = True
                            End If
                            'Zl_����תסԺ_����������
                            strSQL = "Zl_����תסԺ_����������("
                            '  No_In         סԺ���ü�¼.NO%Type,
                            strSQL = strSQL & "'" & arrInfo(0) & "',"
                            '  ����Ա���_In סԺ���ü�¼.����Ա���%Type,
                            strSQL = strSQL & "'" & UserInfo.��� & "',"
                            '  ����Ա����_In סԺ���ü�¼.����Ա����%Type,
                            strSQL = strSQL & "'" & UserInfo.���� & "',"
                            '  �˷�ʱ��_In   סԺ���ü�¼.�Ǽ�ʱ��%Type, --���ŵ����˷�ʱ,ÿ�ŵ��ݵ��˷�ʱ����ͬ,����ϵͳ��ǰʱ��
                            strSQL = strSQL & "" & strDelDate & ","
                            '  �����˷�_In   Number := 0,
                            strSQL = strSQL & "" & 0 & ","
                            '  ��Ժ����id_In ����Ԥ����¼.����id%Type,
                            strSQL = strSQL & "" & IIf(lng��Ժ����ID = 0, "NULL", lng��Ժ����ID) & ","
                            '  ��ҳid_In     סԺ���ü�¼.��ҳid%Type, --ҽ����Ժ����Ǽ�ʱ�Ŵ���
                            strSQL = strSQL & "" & IIf(lng��ҳID = 0, "Null", lng��ҳID) & ","
                            '  �����˷�_In   Number := 0,
                            strSQL = strSQL & "" & IIf(blnExecuteThreeSwap = True, 1, 0) & ","
                            '  ����ID_In     סԺ���ü�¼.����id%Type)
                            strSQL = strSQL & "" & lng����ID & ")"
                            Call zlDatabase.ExecuteProcedure(strSQL, "����������")
                        End If
                    End If
                    str��ת����ID = str��ת����ID & "," & arrInfo(2)
                End If
            End If
        Next
        i = i + lngcnt
    Loop
    
    '�Բ�����㵥�ݽ����˷Ѵ���
    If strReplenishNos <> "" Then
        strReplenishNos = Mid(strReplenishNos, 2)
        If ExecuteReplenishDel(strReplenishNos, cllReplenishPro, lng��ҳID, lng��Ժ����ID, strOutDelDate) = False Then
            Exit Function
        End If
    End If
    
    '��סԺ���ʽ������ʴ���
    If strJzNOs <> "" Then
        strJzNOs = Mid(strJzNOs, 2)
        If DelBalaceMz(strJzNOs, cllJzPro, strOutDelDate) = False Then
            Exit Function
        End If
    End If
    
     '��ӡԤ�����
     Call PrintPrePayPrint(frmMain, strOutDelDate)
     If strJzNOs <> "" And mbln�������� = True Then
        '��ʾ���ʴ���
        Call SHowBalanceWindows(strOutDelDate)
     End If
    ExecuteTurn = True
    Exit Function
errH:
    If blnTrans Then
        gcnOracle.RollbackTrans
        'ҽ����HIS����ͬһ������,HIS����ʧ��,��ҽ���������ϴ�,������Ҫ��"ȡ������"�ӿ�
        If blnTransMedicare And mbln�������� Then Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, False, intInsure)
    End If
    Call ErrCenter
    Call SaveErrLog
End Function

Private Function ExecuteReplenishDel(ByVal strNos As String, ByVal cllPro As Collection, _
    ByVal lng��ҳID As Long, ByVal lng��Ժ����ID As Long, ByVal strDelDate As String) As Boolean
    '����:�Բ������ĵ��ݽ���ת���ü��˷Ѵ���
    '���:
    '   strNos �����㵥��,��ʽ�����ݺ�,����;...
    '   cllPro ������˷ѹ��̵ļ��ϣ�Array(�����㵥�ݺ�,ת����SQL)
    '   strDelDate �˷�ʱ��
    Dim strSQL As String, strNoTemp As String
    Dim varNos As Variant, i As Long, p As Long, blnTrans As Boolean
    Dim strno As String, intInsure As Integer
    Dim lng�������ID  As Long, lng���ó���ID As Long, lng������� As Long
    Dim lngԭ����ID As Long, strAdvance As String
    
    Err = 0: On Error GoTo errH
    If strNos = "" Then ExecuteReplenishDel = True: Exit Function
    
    varNos = Split(strNos, ";")
    For i = 0 To UBound(varNos)
        '���ݺ�,����;...
        strno = Split(varNos(i), ",")(0): intInsure = Split(varNos(i), ",")(1)
        
        lng���ó���ID = zlDatabase.GetNextId("���˽��ʼ�¼")
        lng�������ID = zlDatabase.GetNextId("���˽��ʼ�¼")
        lng������� = -1 * lng���ó���ID
        
        gcnOracle.BeginTrans: blnTrans = True
        For p = 1 To cllPro.Count
            'Array(�����㵥�ݺ�,ת����SQL)
            strNoTemp = cllPro(p)(0): strSQL = cllPro(p)(1)
            If strNoTemp = strno Then
                zlDatabase.ExecuteProcedure strSQL, Me.Caption
            End If
        Next
        
        'Zl_����תסԺ_������ת��(
        strSQL = "Zl_����תסԺ_������ת��("
        '  No_In         ���ò����¼.No%Type,
        strSQL = strSQL & "'" & strno & "',"
        '  ���ó���id_In     ����Ԥ����¼.����id%Type,
        strSQL = strSQL & "" & lng���ó���ID & ","
        '  �������id_In     ����Ԥ����¼.����id%Type,
        strSQL = strSQL & "" & lng�������ID & ","
        '  �������_In     ����Ԥ����¼.�������%Type,
        strSQL = strSQL & "" & lng������� & ","
        '  �˷�ʱ��_In   סԺ���ü�¼.����ʱ��%Type,
        strSQL = strSQL & "To_Date('" & strDelDate & "','yyyy-mm-dd hh24:mi:ss'),"
        '  ����Ա���_In סԺ���ü�¼.����Ա���%Type,
        strSQL = strSQL & "'" & UserInfo.��� & "',"
        '  ����Ա����_In סԺ���ü�¼.����Ա����%Type,
        strSQL = strSQL & "'" & UserInfo.���� & "',"
        '  ��ҳid_In     ����Ԥ����¼.��ҳid%Type,
        strSQL = strSQL & "" & lng��ҳID & ","
        '  ��Ժ����id_In ����Ԥ����¼.����id%Type,
        strSQL = strSQL & "" & lng��Ժ����ID & ")"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        'Public Function ClinicDelSwap(lngStlID As Long, Optional ByVal bln�˷� As Boolean = True, _
            Optional ByVal intinsure As Integer = 0, Optional ByRef strAdvance As String = "") As Boolean
            '---------------------------------------------------------------------------------------------------------------------------------------------
            '����:�������˷ѵ���ϸ�ͽ�������ת����ҽ��ǰ�÷�����ȷ��
            '���:lngStlID-��Ҫ�˵ķѼ�¼�Ľ���ID������Ԥ����¼�п��Լ���ҽ���ź�����
            '     bln�˷� -�������˷ѽ��׻��Ǹķѽ����ڵ��ñ��ӿ�
            '     strAdvance:��ʽ:����ID|��������־|��,ÿλ|�ָ�
            '           ��һλ:�������ID,ҽ�����Ը��ݳ���ID������ȡ��
            '           �ڶ�λ:��������־,1-����������;0�ǲ���������
            '           ����λ:NO:��ǰ�����NO
            '           ����λ��: ���Ժ���չ
            '     ע�⣺
            '           strAdvance��10.34.0��ǰ(�������ʽ���)
            '               �൥��һ�ν���ʱ,�������ԭ����IDs:����ID1,����ID2,...
            '               �����������ʽΪ:�˷ѵ���������|��ǰ�˵ڼ��ŵ���
            '����:strAdvance:1.ԭ���˻�ʱ�����ؿ�
            '                2.�˷ѽ��㷽ʽ���շѽ��㷽ʽ��һ��ʱ�����ظ�ʽΪ�����㷽ʽ|���||���㷽ʽ|���||�������У����Ϊ����
            '���أ����׳ɹ�����true�����򣬷���false
        strAdvance = lng�������ID & "|1"
        lngԭ����ID = zlGetFromNOToLastBalanceID(strno, , , , True)
        If Not gclsInsure.ClinicDelSwap(lngԭ����ID, True, intInsure, strAdvance) Then
            gcnOracle.RollbackTrans
            MsgBox "ҽ������ʧ�ܣ��޷����������������תסԺ������", vbInformation, gstrSysName
            Exit Function
        End If
        gcnOracle.CommitTrans: blnTrans = False
        Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, True, intInsure)
    Next
    ExecuteReplenishDel = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        'Resume
    End If
    Call SaveErrLog
End Function

Private Function zlGetFromNOToLastBalanceID(ByVal strNos As String, _
    Optional ByVal blnNOMoved As Boolean, _
    Optional ByVal bln��ʷ��ͬ���� As Boolean = False, _
    Optional lng������� As Long, Optional bln������ As Boolean = False) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����һ���շѵ��ݵ�NO���������һ����Ч�Ľ��ʵ�ID
    '���:blnNoMoved�Ƿ��ں󱸱��У���ѯ����֮ǰ���ж���Ҫ���������
    '     bln��ʷ��ͬ����-�Ƿ�������ʷ��һ���ѯ
    '     bln������-�Ƿ񲹳����
    '����:lng�������-�������һ����Ч�Ľ������
    '����:����ID
    '����:���˺�
    '����:2014-06-17 17:06:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String, strSQL1 As String
    Dim i As Long
    
    On Error GoTo errHandle:
    '87975
    strSQL = "With c_���� As (Select Column_Value As NO From Table(f_Str2list([1])))" & vbNewLine & _
            " Select Max(a.����id) As ����id" & vbNewLine & _
            " From ������ü�¼ A, c_���� M" & vbNewLine & _
            " Where a.No = m.No" & vbNewLine & _
            "       And a.�Ǽ�ʱ�� + 0 =" & vbNewLine & _
            "           (Select Max(m.�Ǽ�ʱ��)" & vbNewLine & _
            "            From ������ü�¼ M, c_���� J" & vbNewLine & _
            "            Where m.No = j.No And Mod(m.��¼����, 10) = 1 And m.��¼״̬ In (1, 3) And Nvl(m.����״̬, 0) <> 1)" & vbNewLine & _
            "            And Mod(a.��¼����, 10) = 1 And a.��¼״̬ In (1, 3) And Nvl(a.����״̬, 0) <> 1"

    If bln������ Then
        strSQL = Replace(strSQL, "������ü�¼", "���ò����¼")
        strSQL = Replace(strSQL, "Max(a.����id)", "Max(a.����id)")
    End If

    strSQL = "" & _
            "   Select /*+ Rule */ A.����ID,B.������� " & _
            "   From (" & strSQL & ") A,����Ԥ����¼ B " & _
            "   Where A.����ID=B.����ID(+) And Rownum<2"

    If Not blnNOMoved And bln��ʷ��ͬ���� Then
        strSQL1 = Replace(strSQL, "������ü�¼", "H������ü�¼")
        strSQL1 = Replace(strSQL, "���ò����¼", "H���ò����¼")
        strSQL1 = Replace(strSQL, "����Ԥ����¼", "H����Ԥ����¼")
        strSQL = strSQL & " Union ALL " & strSQL1
    ElseIf blnNOMoved Then
        strSQL = Replace(strSQL, "������ü�¼", "H������ü�¼")
        strSQL1 = Replace(strSQL, "���ò����¼", "H���ò����¼")
        strSQL = Replace(strSQL, "����Ԥ����¼", "H����Ԥ����¼")
    End If

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���ݵ��ݻ�ȡ���һ���������ʵĽ���ID", strNos)

    If rsTemp.EOF Then Exit Function

    lng������� = Val(Nvl(rsTemp!�������))
    zlGetFromNOToLastBalanceID = Val(Nvl(rsTemp!����ID))
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function DelBalaceMz(ByVal strNos As String, ByVal cllPro As Collection, ByVal strDelDate As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鵥��������ĸ����������˻ؿ����Ƿ���ȷ
    '���:strNos-���ʵ���(�ö��ŷ���)
    '        cllPro-����ļ������ʹ��̵ļ���(����,"K"+NO)
    '        strDelDate-����ʱ��
    '����:
    '����:
    '����:���˺�
    '����:2011-03-29 14:01:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strBalance As String, strBalanceNo As String, strBalanceNos As String
    Dim strBalanceIDs As String, i As Long, j As Long, lng����ID As Long, intInsure As Integer
    Dim varBalance As Variant, varJz As Variant, varTemp As Variant, blnTrans As Boolean, blnTransMedicare As Boolean
    Dim strδ��NOs As String
    
    Err = 0: On Error GoTo errH
    '1.�������
    strSQL = "" & _
    "   Select  /*+ rule */  distinct B.ID as ����ID,B.NO as ���ʵ�,A.NO as ���ʵ�,C.���� as ҽ��" & _
    "   From ������ü�¼ A, ���˽��ʼ�¼ B,���ս����¼ C, Table(f_Str2list([1])) J" & _
    "   Where A.NO = J.Column_Value  " & _
    "           And A.����id = B.ID And B.��¼״̬=1  " & _
    "           And A.����ID=C.��¼ID(+)  " & _
    "           And C.����(+)=1 And A.��¼���� In (2, 12) " & _
    "   Order by ���ʵ�,���ʵ�"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNos)
    strδ��NOs = "," & strNos & ","
    With rsTemp
        strBalance = "": strBalanceNos = "": strBalanceIDs = ""
        strBalanceNo = ""
        Do While Not .EOF
               If strBalanceNo <> Nvl(rsTemp!���ʵ�) Then
                    intInsure = Val(Nvl(!ҽ��)): lng����ID = Val(Nvl(rsTemp!����ID))
                    strBalanceNo = Nvl(rsTemp!���ʵ�)
                    strBalanceIDs = strBalanceIDs & "," & lng����ID
                    strBalanceNos = strBalanceNos & "," & strBalanceNo
                    strBalance = strBalance & "||" & strBalanceNo & "," & lng����ID & "," & intInsure & "|"
               End If
               strBalance = strBalance & "," & Nvl(rsTemp!���ʵ�)
               strδ��NOs = "," & Replace(strδ��NOs, "," & Nvl(rsTemp!���ʵ�) & ",", "") & ","
               .MoveNext
        Loop
        '����δ�����ʳ������ֵĵ���
        varTemp = Split(strδ��NOs, ",")
        strBalance = strBalance & "||,0,0|"
        For i = 0 To UBound(varTemp)
            If Trim(varTemp(i)) <> "" Then
                strBalance = strBalance & "," & varTemp(i)
            End If
        Next
    End With
    '����Ƿ�������ѿ�����
    If strBalanceNos <> "" Then strBalanceNos = Mid(strBalanceNos, 2)
    If zlIsExistsSquareCard(strBalanceNos, 2) Then
        '���ѿ����
        MsgBox "�ڽ��ʵ���" & strBalanceNos & "�д������ѿ����ݲ�֧�ֶ����ѿ�������תסԺ����,����!", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    '����Ƿ����һ��ͨ����
    If strBalanceIDs <> "" Then strBalanceIDs = Mid(strBalanceIDs, 2)
    Set mrsOneCard = zlGetOneCard(strBalanceIDs)
    If mrsOneCard.RecordCount > 0 Then
        MsgBox "�ڽ��ʵ���" & strBalanceNos & "�д���һ��ͨ���㣬�ݲ�֧������תסԺ����,����!", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    Set mrsOneCard = zlGetThreeCard(strBalanceIDs)
    If mrsOneCard.RecordCount > 0 Then
        MsgBox "�ڽ��ʵ���" & strBalanceNos & "�д������������㣬�ݲ�֧������תסԺ����,����!", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    
    Err = 0: On Error GoTo errH:
    '��ʽ�������
    '            Dim varBalance As Variant, varJz As Variant, varTemp As Variant
    If strBalance <> "" Then strBalance = Mid(strBalance, 3)
    '��ʽ:����NO,����ID,����|���ʵ�1,���ʵ�2,....||����NO1...
    varBalance = Split(strBalance, "||")
    For i = 0 To UBound(varBalance)
        varTemp = Split(Split(varBalance(i), "|")(0), ",")
        varJz = Split(Split(varBalance(i), "|")(1), ",")
        intInsure = Val(varTemp(2)): lng����ID = Val(varTemp(1)): strBalanceNo = varTemp(0)
        gcnOracle.BeginTrans: blnTrans = True: blnTransMedicare = False
        '���ʵ����ʴ���
        For j = 0 To UBound(varJz)
            If varJz(j) <> "" Then
                Call zlDatabase.ExecuteProcedure(cllPro("K" & varJz(j)), "�������תסԺ-��������")
            End If
        Next
        '���ʵ�����
        If lng����ID <> 0 Then
            If DelBalance(strDelDate, strBalanceNo, lng����ID, intInsure, blnTransMedicare) = False Then
                    If blnTrans Then gcnOracle.RollbackTrans
                    'ҽ����HIS����ͬһ������,HIS����ʧ��,��ҽ���������ϴ�,������Ҫ��"ȡ������"�ӿ�
                    If blnTransMedicare And mbln�������� Then Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, False, intInsure)
                    Exit Function
            End If
        End If
        gcnOracle.CommitTrans: blnTrans = False
        If blnTransMedicare Then Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, True, intInsure)
    Next
    DelBalaceMz = True
    Exit Function
errH:
    If blnTrans Then
        gcnOracle.RollbackTrans
        'ҽ����HIS����ͬһ������,HIS����ʧ��,��ҽ���������ϴ�,������Ҫ��"ȡ������"�ӿ�
        If blnTransMedicare And mbln�������� Then Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, False, intInsure)
    End If
End Function

Private Function SHowBalanceWindows(ByVal strDelDate As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ���ʴ���
    ' ���:strDelDate-��������(��ҪӦ�����ٴν���ʱ��Ԥ����)
    '����:���˺�
    '����:2011-03-29 17:38:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objInExse As Object
    Dim lng����ID As Long
   '4.�������ʲ���
    If objInExse Is Nothing Then
        Err = 0: On Error Resume Next
        Set objInExse = CreateObject("zl9InExse.clsFeeQuery")
        If Err <> 0 Then
            MsgBox "ע��:" & "�ڴ���סԺ���ò���ʱ����,���ܸò���δ����ע��,����ʧ��,��ע�����½���!", vbInformation, gstrSysName
            SHowBalanceWindows = True
            Exit Function
        End If
    End If
    On Error GoTo errHandle
    'zlPatiBalance(ByVal frmMain As Object, _
    '    ByVal cnOracle As ADODB.Connection, ByVal lngSys As Long, strDBUser As String, _
    '    ByVal lng����ID As Long, ByVal lng��ҳID As   long ) as boolean
    lng����ID = 0
    If mlngPatient <> 0 Then
        lng����ID = mlngPatient
    ElseIf Not mrsInfo Is Nothing Then
            If mrsInfo.State = 1 Then lng����ID = Val(Nvl(mrsInfo!����ID))
    End If
    If objInExse.zlPatiBalance(Me, gcnOracle, glngSys, gstrDBUser, lng����ID, mlng��ҳID, strDelDate) = False Then
        '���ý���
    End If
    SHowBalanceWindows = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub ShowBills(ByVal lngPatient As Long, ByVal DatBegin As Date, ByVal DatEnd As Date)
'����:��ȡ����ʾ����ָ�������ڵ�������õ���
    Dim i As Long, DatTmp As Date, strSQL As String
    Dim rsList As ADODB.Recordset
    Dim strWhere As String, strInsure As String
    If DatBegin > DatEnd Then
        DatTmp = DatEnd
        DatEnd = DatBegin
        DatBegin = DatTmp
    End If
    If mbln����תסԺ����� Then
       strWhere = " And A.����id = [1] "
       strInsure = " And ����id = [1] "
    Else
        If DatEnd - DatBegin < 4 Then   '36170
            strWhere = " And A.����id+0 = [1] And A.����ʱ�� Between [2] And [3]  "
        Else
            strWhere = " And A.����id = [1] And A.����ʱ��+0 Between [2] And [3]  "
        End If
    End If
    strInsure = " And ����id = [1]  "
    sta.Panels(2).Text = "���ڶ�ȡ�շѵ���,���Ժ� ..."
    Screen.MousePointer = 11
    DoEvents
    Me.Refresh
    On Error GoTo errH
        
   strSQL = strSQL & _
            " Select x.ѡ��, x.���, x.����, Decode(Nvl(z.����, 0), 0, '', '��') As ҽ��, x.No As ���ݺ�, x.Ʊ�ݺ�, x.������, x.Ӧ�ս��, x.ʵ�ս��, x.����ʱ��, Max(y.����id) As ����id," & vbNewLine & _
            "       Nvl(z.����, 0) As ����" & vbNewLine & _
            " From (Select '��' As ѡ��, '��ת��' As ���, '�շѵ�' As ����, a.No," & vbNewLine & _
            "        a.ʵ��Ʊ�� As Ʊ�ݺ�, a.������, LTrim(To_Char(Sum(a.Ӧ�ս��), '9999999990.0000')) As Ӧ�ս��," & vbNewLine & _
            "        LTrim(To_Char(Sum(a.ʵ�ս��), '9999999990.0000')) As ʵ�ս��, To_Char(a.����ʱ��, 'YYYY-MM-DD HH24:MI:SS') As ����ʱ��" & vbNewLine & _
            "       From ������ü�¼ A" & vbNewLine & _
            "       Where Mod(a.��¼����, 10) = 1 And a.��¼״̬ <> 0 " & strWhere & " " & _
                        IIf(mbln����תסԺ�����, "And Exists (Select 1 From ������ü�¼ M,������˼�¼ J " & _
            "                                                  Where M.ID=J.����ID And M.����ID = [1] and M.NO=A.NO And Mod(M.��¼����,10)=Mod(A.��¼����,10) And " & _
            "                                                       J.������� is Not NULL and  nvl(J.��¼״̬,0)=0 and J.����=1) " & vbNewLine, " And Not Exists (Select 1 From ������ü�¼ M,������˼�¼ J Where M.ID=J.����ID And M.����ID = [1] and M.NO=A.NO And Mod(M.��¼����,10)=Mod(A.��¼����,10) And J.������� is Not NULL and  nvl(J.��¼״̬,0) > 0 and J.����=1)") & _
            " And Exists" & vbNewLine & _
            "        (Select 1" & vbNewLine & _
            "              From ������ü�¼ K" & vbNewLine & _
            "              Where k.No = a.No And k.����id = [1] And Mod(k.��¼����, 10) = Mod(a.��¼����, 10) And Nvl(k.���ӱ�־, 0) <> 9" & vbNewLine & _
            "              Group By k.���" & vbNewLine & _
            "              Having Sum(k.ʵ�ս��) <> 0)" & vbNewLine & _
            "       Group By a.No, a.ʵ��Ʊ��, a.������, a.����ʱ��) X, ������ü�¼ Y," & vbNewLine & _
            "     (Select Distinct ��¼id, ����" & vbNewLine & _
            "       From ���ս����¼" & vbNewLine & _
            "       Where ���� = 1 " & strInsure & ") Z" & vbNewLine & _
            " Where x.No = y.No And Mod(y.��¼����, 10) = 1 And y.��¼״̬ In (1, 3) And y.����ID = [1]" & _
            " And y.�Ǽ�ʱ�� = (Select Max(�Ǽ�ʱ��) From ������ü�¼ Where NO = x.No And Mod(��¼����, 10) = 1 And ����ID = [1] And ��¼״̬ In (1, 3)) And y.����id = z.��¼id(+)" & _
            " Group By x.ѡ��, x.���, x.����, Decode(Nvl(z.����, 0), 0, '', '��'), x.No, x.Ʊ�ݺ�, x.������, x.Ӧ�ս��, x.ʵ�ս��, x.����ʱ��, Nvl(z.����, 0)  "
 
    If chkShow.Value = 0 Then
        strSQL = strSQL & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select x.ѡ��, x.���, x.����, Decode(Nvl(z.����, 0), 0, '', '��') As ҽ��, x.No As ���ݺ�, x.Ʊ�ݺ�, x.������, x.Ӧ�ս��, x.ʵ�ս��, x.����ʱ��, Max(y.����id) As ����id," & vbNewLine & _
            "       Nvl(z.����, 0) As ����" & vbNewLine & _
            "From (Select " & vbNewLine & _
            "        '' As ѡ��, '����ת��' As ���, '�շѵ�' As ����, a.No," & vbNewLine & _
            "        a.ʵ��Ʊ�� As Ʊ�ݺ�, a.������, LTrim(To_Char(Sum(a.Ӧ�ս��), '9999999990.0000')) As Ӧ�ս��," & vbNewLine & _
            "        LTrim(To_Char(Sum(a.ʵ�ս��), '9999999990.0000')) As ʵ�ս��, To_Char(a.����ʱ��, 'YYYY-MM-DD HH24:MI:SS') As ����ʱ��" & vbNewLine & _
            "       From ������ü�¼ A" & vbNewLine & _
            "       Where Mod(a.��¼����, 10) = 1 And a.��¼״̬ = 3 " & strWhere & " And Nvl(a.���ӱ�־, 0) <> 9 And Not Exists" & vbNewLine & _
            "        (Select 1" & vbNewLine & _
            "              From ������ü�¼ K" & vbNewLine & _
            "              Where k.No = a.No And k.����id = [1] And Mod(k.��¼����, 10) = Mod(a.��¼����, 10) And Nvl(k.���ӱ�־, 0) <> 9" & vbNewLine & _
            "              Group By k.���" & vbNewLine & _
            "              Having Sum(k.ʵ�ս��) <> 0)" & vbNewLine & _
            "       Group By a.No, a.ʵ��Ʊ��, a.������, a.����ʱ��) X, ������ü�¼ Y," & vbNewLine & _
            "     (Select Distinct ��¼id, ����" & vbNewLine & _
            "       From ���ս����¼" & vbNewLine & _
            "       Where ���� = 1" & strInsure & ") Z" & vbNewLine & _
            " Where x.No = y.No And Mod(y.��¼����, 10) = 1 And y.��¼״̬ In (1, 3) And y.����ID = [1]" & _
            " And y.�Ǽ�ʱ�� = (Select Max(�Ǽ�ʱ��) From ������ü�¼ Where NO = x.No And Mod(��¼����, 10) = 1 And ����ID = [1] And ��¼״̬ In (1, 3)) And y.����id = z.��¼id(+)" & _
            " Group By x.ѡ��, x.���, x.����, Decode(Nvl(z.����, 0), 0, '', '��'), x.No, x.Ʊ�ݺ�, x.������, x.Ӧ�ս��, x.ʵ�ս��, x.����ʱ��, Nvl(z.����, 0)  "

            
        strSQL = strSQL & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select x.ѡ��, x.���, x.����, Decode(Nvl(z.����, 0), 0, '', '��') As ҽ��, x.No As ���ݺ�, x.Ʊ�ݺ�, x.������, x.Ӧ�ս��, x.ʵ�ս��, x.����ʱ��, Max(y.����id) As ����id," & vbNewLine & _
            "       Nvl(z.����, 0) As ����" & vbNewLine & _
            "From (Select " & vbNewLine & _
            "        '' As ѡ��, '����ת��' As ���, '�շѵ�' As ����, a.No," & vbNewLine & _
            "        a.ʵ��Ʊ�� As Ʊ�ݺ�, a.������, LTrim(To_Char(Sum(a.Ӧ�ս��), '9999999990.0000')) As Ӧ�ս��," & vbNewLine & _
            "        LTrim(To_Char(Sum(a.ʵ�ս��), '9999999990.0000')) As ʵ�ս��, To_Char(a.����ʱ��, 'YYYY-MM-DD HH24:MI:SS') As ����ʱ��" & vbNewLine & _
            "       From ������ü�¼ A" & vbNewLine & _
            "       Where Mod(a.��¼����, 10) = 1 And a.��¼״̬ <> 0 " & strWhere & " " & _
            " And Exists (Select 1 From ������ü�¼ M,������˼�¼ J Where M.ID=J.����ID And M.����ID = [1] and M.NO=A.NO And Mod(M.��¼����,10)=Mod(A.��¼����,10) And J.������� is Not NULL and  nvl(J.��¼״̬,0) = 1 and J.����=1)" & _
            " And Exists" & vbNewLine & _
            "        (Select 1" & vbNewLine & _
            "              From ������ü�¼ K" & vbNewLine & _
            "              Where k.No = a.No And k.����id = [1] And Mod(k.��¼����, 10) = Mod(a.��¼����, 10) And Nvl(k.���ӱ�־, 0) <> 9" & vbNewLine & _
            "              Group By k.���" & vbNewLine & _
            "              Having Sum(k.ʵ�ս��) <> 0)" & vbNewLine & _
            "       Group By a.No, a.ʵ��Ʊ��, a.������, a.����ʱ��) X, ������ü�¼ Y," & vbNewLine & _
            "     (Select Distinct ��¼id, ����" & vbNewLine & _
            "       From ���ս����¼" & vbNewLine & _
            "       Where ���� = 1 " & strInsure & ") Z" & vbNewLine & _
            " Where x.No = y.No And Mod(y.��¼����, 10) = 1 And y.��¼״̬ In (1, 3) And y.����ID = [1]" & _
            " And y.�Ǽ�ʱ�� = (Select Max(�Ǽ�ʱ��) From ������ü�¼ Where NO = x.No And Mod(��¼����, 10) = 1 And ����ID = [1] And ��¼״̬ In (1, 3)) And y.����id = z.��¼id(+)" & _
            " Group By x.ѡ��, x.���, x.����, Decode(Nvl(z.����, 0), 0, '', '��'), x.No, x.Ʊ�ݺ�, x.������, x.Ӧ�ս��, x.ʵ�ս��, x.����ʱ��, Nvl(z.����, 0)  "

    End If
     
    strSQL = strSQL & " UNION ALL " & _
            " Select    '��' as ѡ��,'��ת��' as ���,'���ʵ�' as ����,Decode(NULL,Null,'','��') as ҽ��, A.NO As ���ݺ�, A.ʵ��Ʊ�� As Ʊ�ݺ�, A.������," & vbNewLine & _
            "       LTrim(To_Char(Sum(A.Ӧ�ս��), '999999999" & gstrDec & "')) As Ӧ�ս��," & vbNewLine & _
            "       LTrim(To_Char(Sum(A.ʵ�ս��), '999999999" & gstrDec & "')) As ʵ�ս��," & vbNewLine & _
            "       To_Char(A.����ʱ��, 'YYYY-MM-DD HH24:MI:SS') As ����ʱ��, 0 as ����ID,0 as ����" & vbNewLine & _
            " From ������ü�¼ A" & vbNewLine & _
            " Where A.��¼���� =2 And A.��¼״̬ <> 0 " & strWhere & vbNewLine & _
            "           And Exists (Select 1 From ������ü�¼ K Where K.NO=A.NO And K.��¼����=A.��¼���� And K.���ӱ�־ <> 9 Group By K.��� Having Sum(K.����) <> 0) " & vbNewLine & _
                        IIf(mbln����תסԺ�����, "           And Exists(Select 1 From ������ü�¼ M,������˼�¼ J where M.ID=J.����ID and M.NO=A.NO And M.��¼����=A.��¼���� And J.������� is Not NULL and  nvl(J.��¼״̬,0)=0 and J.����=1) " & vbNewLine, " And Not Exists(Select 1 From ������ü�¼ M,������˼�¼ J where M.ID=J.����ID and M.NO=A.NO And M.��¼����=A.��¼���� And J.������� is Not NULL and  nvl(J.��¼״̬,0) > 0 and J.����=1) ") & _
            "Group By A.NO, A.ʵ��Ʊ��, A.������, A.����ʱ�� "
         
    If chkShow.Value = 0 Then
        strSQL = strSQL & " UNION ALL " & _
            " Select C.ѡ��,C.���,C.����,C.ҽ��,C.���ݺ�, C.Ʊ�ݺ�, C.������," & vbNewLine & _
            "       LTrim(To_Char(Sum(D.Ӧ�ս��), '999999999" & gstrDec & "')) As Ӧ�ս��," & vbNewLine & _
            "       LTrim(To_Char(Sum(D.ʵ�ս��), '999999999" & gstrDec & "')) As ʵ�ս��," & vbNewLine & _
            "       C.����ʱ��, C.����ID, C.����" & vbNewLine & _
            " From " & _
            " (Select    '' as ѡ��,'����ת��' as ���,'���ʵ�' as ����,Decode(NULL,Null,'','��') as ҽ��, A.NO As ���ݺ�, A.ʵ��Ʊ�� As Ʊ�ݺ�, A.������," & vbNewLine & _
            "       LTrim(To_Char(Sum(A.Ӧ�ս��), '999999999" & gstrDec & "')) As Ӧ�ս��," & vbNewLine & _
            "       LTrim(To_Char(Sum(A.ʵ�ս��), '999999999" & gstrDec & "')) As ʵ�ս��," & vbNewLine & _
            "       To_Char(A.����ʱ��, 'YYYY-MM-DD HH24:MI:SS') As ����ʱ��,0 as ����ID,0 as ����" & vbNewLine & _
            " From ������ü�¼  A" & vbNewLine & _
            " Where A.��¼���� = 2 And A.��¼״̬ In (2,3)  And Not Exists (Select 1 From ������ü�¼ Where NO=A.NO And ��¼״̬=1 And ��¼����=2) " & strWhere & vbNewLine & _
            "           And Not Exists (Select 1 From ������ü�¼ K Where K.NO=A.NO And K.��¼����=A.��¼���� And K.���ӱ�־ <> 9 Group By K.��� Having Sum(K.ʵ�ս��) <> 0) " & vbNewLine & _
            " Group By A.NO, A.ʵ��Ʊ��, A.������, A.����ʱ�� Having Sum(A.ʵ�ս��)=0) C,������ü�¼ D Where C.���ݺ�=D.NO And D.��¼����=2 And D.��¼״̬=3" & vbNewLine & _
            " Group By C.ѡ��,C.���,C.����,C.ҽ��,C.���ݺ�, C.Ʊ�ݺ�, C.������,C.����ʱ��, C.����ID, C.���� "
            
        strSQL = strSQL & " UNION ALL " & _
            " Select    '' as ѡ��,'����ת��' as ���,'���ʵ�' as ����,Decode(NULL,Null,'','��') as ҽ��, A.NO As ���ݺ�, A.ʵ��Ʊ�� As Ʊ�ݺ�, A.������," & vbNewLine & _
            "       LTrim(To_Char(Sum(A.Ӧ�ս��), '999999999" & gstrDec & "')) As Ӧ�ս��," & vbNewLine & _
            "       LTrim(To_Char(Sum(A.ʵ�ս��), '999999999" & gstrDec & "')) As ʵ�ս��," & vbNewLine & _
            "       To_Char(A.����ʱ��, 'YYYY-MM-DD HH24:MI:SS') As ����ʱ��, 0 as ����ID,0 as ����" & vbNewLine & _
            " From ������ü�¼ A" & vbNewLine & _
            " Where A.��¼���� = 2 And A.��¼״̬ <> 0 " & strWhere & vbNewLine & _
            "           And Exists (Select 1 From ������ü�¼ K Where K.NO=A.NO And K.��¼����=A.��¼���� And K.���ӱ�־ <> 9 Group By K.��� Having Sum(K.����) <> 0) " & vbNewLine & _
            " And  Exists (Select 1 From ������ü�¼ M,������˼�¼ J where M.ID=J.����ID and M.NO=A.NO And M.��¼����=A.��¼���� And J.������� is Not NULL and  nvl(J.��¼״̬,0) = 1 and J.����=1) " & _
            "Group By A.NO, A.ʵ��Ʊ��, A.������, A.����ʱ�� "
        
    End If
    strSQL = strSQL & "Order By ����,���, Ʊ�ݺ� Desc, ���ݺ� Desc"
   'ע��:����ҽ��Ҫ������һ�ſ�ʼ��,��������ܹؼ�
    Set rsList = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngPatient, DatBegin, DatEnd)
    mshList.Redraw = flexRDNone: mshList.Clear
    mshList.Rows = 2
    Set mshList.DataSource = rsList
    If rsList.EOF Then
        sta.Panels(2).Text = "û���ҵ�ָ��ʱ�䷶Χ���շѻ���ʵ���!"
        mshList.Rows = 2
    Else
        sta.Panels(2).Text = "�� " & rsList.RecordCount & " ���շѵ���"
    End If
    Call setHeader
    Call SetInsure
    Call SetBillColor
    mshList.Redraw = flexRDBuffered
    Call mshList_AfterRowColChange(0, 0, 1, 0)
    If mshList.Rows >= 2 Then mshList.Select 1, 0
    Screen.MousePointer = 0
    Call SetSumMoney
    Me.Refresh
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetInsure()
    Dim intInsure As Integer, lngRow As Long
    Dim str���� As String, strno As String
    
    With mshList
        For lngRow = 1 To .Rows - 1
            If .TextMatrix(lngRow, .ColIndex("���")) = "��ת��" And .TextMatrix(lngRow, .ColIndex("ѡ��")) = "��" Then
                intInsure = Val(.TextMatrix(lngRow, .ColIndex("����")))
                str���� = Trim(.TextMatrix(lngRow, .ColIndex("����")))
                If intInsure > 0 And str���� = "�շѵ�" Then
                    If Not gclsInsure.GetCapability(support�����������, mlngPatient, intInsure) Then
                        .TextMatrix(lngRow, .ColIndex("ѡ��")) = ""
                    End If
                End If
            End If
        Next lngRow
    End With
End Sub

Private Sub dtpEnd_Change()
    dtpBegin.MaxDate = dtpEnd.Value
End Sub

Private Sub CheckInsure()
    Dim i As Integer, intInsure As Integer, blnSelect As Boolean
    With mshList
        For i = 1 To .Rows - 1
            intInsure = Val(.TextMatrix(i, .ColIndex("����")))
            blnSelect = .TextMatrix(i, .ColIndex("ѡ��")) <> ""
            If intInsure > 0 And blnSelect Then
                If gclsInsure.GetCapability(support�����������, mlngPatient, intInsure) = False Then
                    .TextMatrix(i, .ColIndex("ѡ��")) = ""
                End If
            End If
        Next i
    End With
End Sub

Private Function ExecuteThreeSwap(lngBalance As Long, lng����ID As Long, Optional ByRef str������ˮ�� As String, Optional ByRef str����˵�� As String) As Boolean
    Dim strSQL As String, rsTemp As ADODB.Recordset, strBalanceIDs As String, rsTotal As ADODB.Recordset
    Dim dblMoney As Double, strAll As String, strDetail() As String, strItem() As String, strCardNO As String
    Dim i As Integer, lngCardID As Long
    
    If mobjSquare Is Nothing Then Set mobjSquare = gobjSquare.objSquareCard
    If mobjSquare Is Nothing Then Exit Function
    strSQL = _
        "Select ժҪ" & vbNewLine & _
        "    From ����Ԥ����¼" & vbNewLine & _
        "    Where ���㷽ʽ Is Null And ��¼���� = 3 And ��¼״̬ = 2 And ����id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    If rsTemp.RecordCount = 0 Then Exit Function
    strAll = Nvl(rsTemp!ժҪ)
    If strAll = "" Then Exit Function
    
    strDetail = Split(strAll, "|")
    For i = 0 To UBound(strDetail)
        If strDetail(i) <> "" Then
            strItem = Split(strDetail(i), ",")
            If Val(strItem(0)) = 1 Then
                lngCardID = Val(strItem(1))
                dblMoney = -1 * Val(strItem(2))
                strSQL = "Select Distinct a.����id" & vbNewLine & _
                            "From ������ü�¼ A" & vbNewLine & _
                            "Where a.No In (Select Distinct a.No From ������ü�¼ A Where Mod(a.��¼����, 10) = 1 And a.����id = [1]) And Mod(a.��¼����, 10) = 1 And" & vbNewLine & _
                            "      a.��¼״̬ <> 0"
                strSQL = "Select Min(����ID) As ����ID,Min(����) As ���� From ����Ԥ����¼ Where ����ID IN (" & strSQL & ") And �����ID = [2]"
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lngCardID)
                strBalanceIDs = "3|" & Val(Nvl(rsTemp!����ID))
                If mobjSquare.zlReturnCheck(Me, mlngModule, lngCardID, False, Nvl(rsTemp!����), _
                    strBalanceIDs, dblMoney, str������ˮ��, str����˵��, "3|" & lng����ID) = False Then Exit Function
                If mobjSquare.zlReturnMoney(Me, mlngModule, lngCardID, False, Nvl(rsTemp!����), _
                    strBalanceIDs, dblMoney, str������ˮ��, str����˵��, "3|" & lng����ID) = False Then Exit Function
            End If
        End If
    Next i
    
    ExecuteThreeSwap = True
End Function

Private Sub setHeader()
    Dim strHead As String
    Dim i As Long
    With mshList
        If .DataSource Is Nothing Then
            strHead = "ѡ��,4,500|���,4,850|����,4,800|ҽ��,4,500|���ݺ�,4,850|Ʊ�ݺ�,4,1100|������,4,800|Ӧ�ս��,7,850|ʵ�ս��,7,850|����ʱ��,4,1850|����ID,4,0|����,4,0"
            .Cols = UBound(Split(strHead, "|")) + 1
            For i = 0 To UBound(Split(strHead, "|"))
                .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
                .ColKey(i) = Trim(.TextMatrix(0, i))
            Next
            .Rows = 2
        End If
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .Cols - 1)
        'ѡ��,4,500|���,4,850|ҽ��,4,500|���ݺ�,4,850|Ʊ�ݺ�,4,1100|������,4,800|Ӧ�ս��,7,850|ʵ�ս��,7,850|����ʱ��,4,1850|����ID,4,0|����,4,0
        For i = 0 To .Cols - 1
             .FixedAlignment(i) = flexAlignCenterCenter
             .colAlignment(i) = flexAlignLeftCenter
             .ColKey(i) = Trim(.TextMatrix(0, i))
             Select Case .ColKey(i)
             Case "ѡ��", "���", "����", "ҽ��", "���ݺ�", "Ʊ�ݺ�"
                .colAlignment(i) = flexAlignCenterCenter
             Case "Ӧ�ս��", "ʵ�ս��"
                .colAlignment(i) = flexAlignRightCenter
             End Select
             If .ColKey(i) Like "*ID" Or .ColKey(i) = "����" Then
                .ColHidden(i) = True: .ColWidth(i) = 0
             End If
        Next
        zl_vsGrid_Para_Restore 1131, mshList, Me.Caption, "����תסԺ�б�", True
        .RowHeight(0) = 320
        .Row = 1
        .Col = 0: .ColSel = .Cols - 1
    End With
End Sub

Private Sub SetBillColor()
    Dim i As Long, j As Long
    With mshList
        For i = 1 To .Rows - 1
            .Row = i
            If .TextMatrix(i, .ColIndex("���")) = "����ת��" Then
                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = &H8000000C
            Else
                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = .ForeColor
            End If
        Next
    End With
End Sub

Private Sub cmdParaSet_Click()
    frmChargeTurnParSet.ShowSet Me, 1131, mstrPrivs
    mbln�������� = Val(zlDatabase.GetPara("����ת�������˷�", glngSys, 1131)) = 1
End Sub

Private Sub cmdSave_Click()
    Dim i As Long, strno As String, strNos As String
    Dim strBalanceID As String, strTemp As String
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim lng����ID As Long, str���ݺ� As String, lngInsure As Long
    Dim strReplenishNo As String, strNotSelectNos As String
    Dim varData As Variant, blnErrBill As Boolean
    
    mstrNOS = ""
    If mlngPatient = 0 Then
        MsgBox "δ���ֲ�����Ϣ�����飡", vbInformation, gstrSysName
        Exit Sub
    End If
    
    With mshList
        strno = ""
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("���")) = "��ת��" And .TextMatrix(i, .ColIndex("ѡ��")) = "��" Then
                lng����ID = Val(.TextMatrix(i, .ColIndex("����ID")))
                str���ݺ� = .TextMatrix(i, .ColIndex("���ݺ�"))
                lngInsure = Val(.TextMatrix(i, .ColIndex("����")))
                strReplenishNo = "": strNotSelectNos = ""
                
                If InStr(1, "," & strno, "," & str���ݺ� & ",") = 0 Then
                    strno = strno & "," & str���ݺ�
                End If
                
                If .TextMatrix(i, .ColIndex("����")) = "�շѵ�" Then
                    If CheckBillExistReplenishData(1, , str���ݺ�, strReplenishNo, blnErrBill) Then
                        If mbln�������� Then
                            If blnErrBill Then
                                MsgBox "���ݺ�Ϊ[" & str���ݺ� & "]�ļ�¼�ѽ���ҽ��������㣬���������쳣����״̬������ת�������ȵ������ղ�����㡿���д���", vbInformation, gstrSysName
                                Exit Sub
                            End If
                            If CheckReplenishAllNosIsSelected(strReplenishNo, .TextMatrix(i, .ColIndex("����")), strNotSelectNos) = False Then
                                MsgBox "���ݺ�Ϊ[" & str���ݺ� & "]�ļ�¼�ѽ��в�����㣬���µ���Ҳ����һ��ת����" & vbCrLf & strNotSelectNos, vbInformation, gstrSysName
                                Exit Sub
                            End If
                            '��ȡҽ������
                            lngInsure = GetReplenishInsure(strReplenishNo)
                            If lngInsure = 0 Then
                                MsgBox "���ݺ�Ϊ[" & str���ݺ� & "]�ļ�¼�ѽ��в�����㣬��δ��ȡ��ҽ������,����ת����", vbInformation, gstrSysName
                                Exit Sub
                            End If
                            '���ҽ���Ƿ��ܹ�ԭ������
                            strTemp = CheckInsureCancel(mlngPatient, lngInsure, strReplenishNo, True)
                            If strTemp <> "" Then
                                MsgBox strTemp, vbInformation, gstrSysName
                                Exit Sub
                            End If
                        Else
                            MsgBox "���ݺ�Ϊ[" & str���ݺ� & "]�ļ�¼�ѽ��в�����㣬����ת����", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                End If
                
                If lngInsure <> 0 Then
                    '���ҽ�������Ƿ�ȫת��
                    If IsYBSingle(str���ݺ�, lngInsure) = False Then
                        If CheckBalanceAllNosIsSelected(lng����ID, .TextMatrix(i, .ColIndex("����"))) = False Then
                            MsgBox "ҽ�����ݺ�Ϊ[" & str���ݺ� & "]�ļ�¼����δת��ȫ����ؽ��㵥�ݣ����ܼ�����", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                Else
                    If CheckAllTurn(str���ݺ�) Then
                        If CheckBalanceAllNosIsSelected(lng����ID, .TextMatrix(i, .ColIndex("����"))) = False Then
                            MsgBox "���ݺ�Ϊ[" & str���ݺ� & "]�ļ�¼����δת��ȫ����ؽ��㵥�ݣ����ܼ�����", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                End If
                
                If .TextMatrix(i, .ColIndex("����")) = "���ʵ�" Then
                    If zlIsExistsSquareCard(str���ݺ�, 2) Then
                        '���ѿ����
                        MsgBox "�ڽ��ʵ���[" & str���ݺ� & "]�д������ѿ����ݲ�֧�ֶ����ѿ�������תסԺ����,����!", vbOKOnly + vbInformation, gstrSysName
                        Exit Sub
                    End If
                    strBalanceID = ""
                    strSQL = "Select Distinct A.����ID From ������ü�¼ A,���˽��ʼ�¼ B" & _
                            " Where A.����ID=B.ID And (b.��¼״̬=1 or nvl(b.����״̬,0)=1)" & _
                            "       and  Mod(A.��¼����,10)=2 And A.No=[1] "
                    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str���ݺ�)
                    Do While Not rsTemp.EOF
                        strBalanceID = strBalanceID & "," & Nvl(rsTemp!����ID)
                        rsTemp.MoveNext
                    Loop
                    '����Ƿ����һ��ͨ����
                    If strBalanceID <> "" Then strBalanceID = Mid(strBalanceID, 2)
                    If strBalanceID <> "" Then
                        Set mrsOneCard = zlGetOneCard(strBalanceID)
                        If mrsOneCard.RecordCount > 0 Then
                            MsgBox "�ڽ��ʵ���[" & str���ݺ� & "]�д���һ��ͨ���㣬�ݲ�֧������תסԺ����,����!", vbOKOnly + vbInformation, gstrSysName
                            Exit Sub
                        End If
                        Set mrsOneCard = zlGetThreeCard(strBalanceID)
                        If mrsOneCard.RecordCount > 0 Then
                            MsgBox "�ڽ��ʵ���[" & str���ݺ� & "]�д������������㣬�ݲ�֧������תסԺ����,����!", vbOKOnly + vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                End If
                
                mstrNOS = mstrNOS & ";" & str���ݺ� & "," & .TextMatrix(i, .ColIndex("Ʊ�ݺ�")) & "," & _
                    lng����ID & "," & lngInsure & "," & .TextMatrix(i, .ColIndex("����")) & "," & strReplenishNo
            End If
        Next
    End With
    If strno <> "" Then strno = Mid(strno, 2)
    If mstrNOS <> "" Then mstrNOS = Mid(mstrNOS, 2)
        
    If mstrNOS = "" Then
        MsgBox "�㻹δѡ��Ҫת��סԺ���õĵ��ݣ��������̣�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '����Ҫѡ����
    If mblnSelPati = False Then Unload Me: Exit Sub
    
    varData = Split(strno, ","): strno = ""
    For i = 0 To UBound(varData)
        If i > 60 Then strno = strno & ",...": Exit For
        strno = strno & IIf(strno = "", "", ",")
        strno = strno & IIf(i > 0 And i Mod 6 = 0, vbCrLf, "")
        strno = strno & varData(i)
    Next
    If MsgBox("���Ƿ���Ҫ�������������ת��סԺ������" & vbCrLf & _
        strno, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        mstrNOS = ""
        Exit Sub
    End If
    
    Err = 0: On Error GoTo ErrHand:
    If Val(Nvl(mrsInfo!��ҳID)) = 0 Then
        MsgBox "�ò��˻�δ��Ժ�������������תסԺ���ã�", vbInformation, gstrSysName
        Exit Sub
    End If
    If ExecuteTurn(Me, mlngModule, mstrPrivs, mstrNOS, Val(Nvl(mrsInfo!סԺ��)), _
        Val(Nvl(mrsInfo!��ҳID)), CDate(Format(mrsInfo!��Ժ����, "yyyy-mm-dd HH:MM:SS")), _
        Val(Nvl(mrsInfo!��Ժ����ID)), Val(Nvl(mrsInfo!��Ժ����ID))) = False Then
        'ת��δ�ɹ�
        Call cmdRefresh_Click
        Exit Sub
    End If
    Unload Me
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetReplenishAllNos(ByVal strno As String) As String
    '��ȡ�����������з��õ���
    '���أ�
    '   �����������з��õ���:A001,A002,...
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strNos As String
    
    On Error GoTo ErrHandler
    strSQL = _
        " Select Distinct a.No" & vbNewLine & _
        " From ������ü�¼ A, ������ü�¼ B, ���ò����¼ C" & vbNewLine & _
        " Where a.No = b.No And a.��� = b.��� And a.��¼���� In (1, 11)" & vbNewLine & _
        "       And b.����id = c.�շѽ���id" & vbNewLine & _
        "       And c.��¼���� = 1 And c.���ӱ�־ = 0 And c.No = [1]" & vbNewLine & _
        " Group By a.No, a.���" & vbNewLine & _
        " Having Nvl(Sum(Nvl(a.����, 1) * a.����), 0) <> 0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strno)
    Do While Not rsTmp.EOF
        strNos = strNos & "," & Nvl(rsTmp!NO)
        rsTmp.MoveNext
    Loop
    If strNos <> "" Then strNos = Mid(strNos, 2)
    
    GetReplenishAllNos = strNos
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckReplenishAllNosIsSelected(ByVal strno As String, ByVal str���� As String, _
    Optional ByRef strNotSelectNos As String) As Boolean
    '��鲹����������ʣ��δ�˷��ñ����Ƿ�ѡ����ת��
    '��Σ�
    '   str���� �շѵ�/���ʵ�
    '���Σ�
    '   strNotSelectNos û�б�ѡ�����Ҫһ��ת���ĵ���
    Dim i As Integer, k As Long, blnFind As Boolean
    Dim strNos As String, varNos As Variant
    
    On Error GoTo ErrHandler
    strNotSelectNos = ""
    strNos = GetReplenishAllNos(strno)
    
    varNos = Split(strNos, ",")
    With mshList
        For i = 0 To UBound(varNos)
            blnFind = False
            For k = 1 To .Rows - 1
                If .TextMatrix(k, .ColIndex("����")) = str���� And .TextMatrix(k, .ColIndex("���ݺ�")) = varNos(i) Then
                    If .TextMatrix(k, .ColIndex("���")) = "��ת��" And .TextMatrix(k, .ColIndex("ѡ��")) = "��" Then
                        blnFind = True: Exit For
                    End If
                End If
            Next
            
            If blnFind = False Then
                strNotSelectNos = strNotSelectNos & "," & varNos(i)
            End If
        Next
    End With
    
    If strNotSelectNos <> "" Then
        strNotSelectNos = Mid(strNotSelectNos, 2)
        Exit Function
    End If
    CheckReplenishAllNosIsSelected = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetReplenishInsure(ByVal strno As String) As Long
    '��ȡ��������ҽ������
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    strSQL = _
        " Select Max(b.����) As ����" & vbNewLine & _
        " From ����Ԥ����¼ A, ���ս����¼ B, ���ò����¼ C" & vbNewLine & _
        " Where a.����id = b.��¼id And a.��¼���� = 6" & vbNewLine & _
        "       And a.����id = c.����id And c.��¼���� = 1" & vbNewLine & _
        "       And c.��¼״̬ In(1,3) And c.���ӱ�־ = 0 And c.No = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strno)
    If Not rsTmp.EOF Then GetReplenishInsure = Nvl(rsTmp!����)
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckBalanceAllNosIsSelected(ByVal lng����ID As Long, ByVal str���� As String) As Boolean
    '���һ�ν��������ʣ��δ�˷��ñ����Ƿ�ѡ����ת��
    '��Σ�
    '   str���� �շѵ�/���ʵ�
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo ErrHandler
    strSQL = _
        " Select Distinct a.No" & vbNewLine & _
        " From ������ü�¼ A, ������ü�¼ B" & vbNewLine & _
        " Where a.No = b.No And Mod(a.��¼����,10) = Mod(b.��¼����,10)" & vbNewLine & _
        "       And a.���=b.��� And b.����id = [1]" & vbNewLine & _
        " Group By a.No,a.���" & vbNewLine & _
        " Having Nvl(Sum(Nvl(a.����,1)*a.����),0) <> 0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    Do While Not rsTmp.EOF
        With mshList
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("����")) = str���� And .TextMatrix(i, .ColIndex("���ݺ�")) = Nvl(rsTmp!NO) Then
                    If Not (.TextMatrix(i, .ColIndex("���")) = "��ת��" And .TextMatrix(i, .ColIndex("ѡ��")) = "��") Then
                        Exit Function
                    End If
                End If
            Next
        End With
        rsTmp.MoveNext
    Loop
    CheckBalanceAllNosIsSelected = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub dtpBegin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub dtpEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_Activate()
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    Call picTop_Resize
End Sub

Private Sub Form_Load()
    Dim strTmp As String, Datsys As Date
    
    If Not gobjSquare Is Nothing Then
        Call IDKind.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "")
         '��ʼ����صı������ݼ�
        Set mtySquareCard.rsSquare = New ADODB.Recordset
        mtySquareCard.blnExistsObjects = Not gobjSquare.objSquareCard Is Nothing
        If Not gobjSquare.objSquareCard Is Nothing Then IDKind.IDKindStr = gobjSquare.objSquareCard.zlGetIDKindStr(IDKind.IDKindStr)
        Set mobjSquare = gobjSquare.objSquareCard
    End If
    Call GetRegInFor(g˽��ģ��, Me.Name, "idkind", strTmp)
    mintIDKind = Val(strTmp)
    If mintIDKind > 0 And mintIDKind <= IDKind.ListCount Then IDKind.IDKind = mintIDKind
    mstrTitle = Me.Caption
    
    Call RestoreWinState(Me, App.ProductName)
    
    mbln����תסԺ����� = IIf(Val(zlDatabase.GetPara("����תסԺ�����", glngSys, 1143, 0)) = 1, True, False)
    mbln�������� = Val(zlDatabase.GetPara("����ת�������˷�", glngSys, 1131)) = 1
    chkShow.Value = IIf(Val(zlDatabase.GetPara("����ʾ��ת������", glngSys, 1131, 1, Array(chkShow))) = 1, 1, 0)
    picBalance.BorderStyle = 0: picList.BorderStyle = 0:    picBill.BorderStyle = 0
    Call InitPancel
    Datsys = zlDatabase.Currentdate
    strTmp = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��ʼʱ��")
    If IsDate(strTmp) Then
        dtpBegin.Value = CDate(strTmp)
    Else
        dtpBegin.Value = Format(DateAdd("d", -3, Datsys), "yyyy-mm-dd 00:00:00")
    End If
    dtpBegin.MaxDate = Format(Datsys, "yyyy-mm-dd 23:59:59")
    If mstrNOS <> "" Then
        strTmp = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "����ʱ��")
    Else
        strTmp = ""
    End If
    If IsDate(strTmp) Then
        dtpEnd.Value = CDate(strTmp)
    Else
        dtpEnd.Value = Format(Datsys, "yyyy-mm-dd 23:59:59")
    End If
    Call SetVisibleCtl
    Call setHeader: Call SetDetail: Call SetBalanceHead
    Call zlCreateObject
End Sub

Private Sub SetVisibleCtl()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ÿؼ���visible����
    '����:���˺�
    '����:2011-03-29 21:49:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    dtpBegin.Visible = Not mbln����תסԺ�����
    dtpEnd.Visible = Not mbln����תסԺ�����
    lbl��.Visible = Not mbln����תסԺ�����
    lblDate.Visible = Not mbln����תסԺ�����
End Sub

Private Sub cmdExit_Click()
    mstrNOS = ""
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub cmdRefresh_Click()
    Call ShowBills(mlngPatient, dtpBegin.Value, dtpEnd.Value)
    If cmdAll(0).Visible And cmdAll(0).Enabled Then Call cmdAll(0).SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    mintIDKind = IDKind.IDKind
    Call SaveRegInFor(g˽��ģ��, Me.Name, "idkind", mintIDKind)
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "��ʼʱ��", Format(dtpBegin.Value, "yyyy-MM-dd HH:mm:ss")
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName & "\" & Me.Name, "����ʱ��", Format(dtpEnd.Value, "yyyy-MM-dd HH:mm:ss")
    Call SaveWinState(Me, App.ProductName)
    Set mtySquareCard.rsSquare = Nothing
    Call zlDatabase.SetPara("����ʾ��ת������", chkShow.Value, glngSys, 1131)
    zlSaveDockPanceToReg Me, dkpMan, "����"
    zl_vsGrid_Para_Save 1131, mshDetail, Me.Caption, "����תסԺ��ϸ�б�", True
    zl_vsGrid_Para_Save 1131, vsBalance, Me.Caption, "����תסԺ�����б�", True
    zl_vsGrid_Para_Save 1131, mshList, Me.Caption, "����תסԺ�б�", True
    Call zlCloseObject
End Sub
 
Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng�����ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
    
    If objCard.���� Like "*IC��*" And objCard.ϵͳ Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        If mobjICCard Is Nothing Then Exit Sub
        txtPatient.Text = mobjICCard.Read_Card()
        If txtPatient.Text <> "" Then Call FindPati(objCard, True, Trim(txtPatient.Text))
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
   If gobjSquare.objSquareCard.zlReadCard(Me, mlngModule, lng�����ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
   txtPatient.Text = strOutCardNO
   If txtPatient.Text <> "" Then Call FindPati(objCard, True, Trim(txtPatient.Text))
End Sub

Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txtPatient.IMEMode = 0
    If txtPatient.Text <> "" And Not mblnNotClick Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    zlControl.TxtSelAll txtPatient
End Sub

Private Sub mobjICCard_ShowICCardInfo(ByVal strCardNO As String)
    If txtPatient.Text <> "" Or txtPatient.Locked Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    Dim objCard As Card
    Set objCard = IDKind.GetIDKindCard("IC��", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    txtPatient.Text = strCardNO
    If txtPatient.Text <> "" Then Call FindPati(objCard, True, txtPatient.Text)
End Sub

Private Sub mshDetail_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save 1131, mshDetail, Me.Caption, "����תסԺ��ϸ�б�", True
End Sub

Private Sub mshDetail_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save 1131, mshDetail, Me.Caption, "����תסԺ��ϸ�б�", True
End Sub


Private Sub mshList_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save 1131, mshList, Me.Caption, "����תסԺ�б�", True
End Sub

Private Sub mshList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strno As String, str���� As String
    
    If NewRow = OldRow Then Exit Sub
    With mshList
        strno = Trim(.TextMatrix(NewRow, .ColIndex("���ݺ�")))
        str���� = Trim(.TextMatrix(NewRow, .ColIndex("����")))
        If NewRow = 0 Or strno = "" Then
            mshDetail.Clear 1: mshDetail.Rows = 2
            Call SetDetail
        Else
            Call ShowDetail(str����, strno)
        End If
        .ForeColorSel = mshList.CellForeColor
    End With
End Sub

Private Sub mshList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save 1131, mshList, Me.Caption, "����תסԺ�б�", True
End Sub

Private Sub mshList_DblClick()
    With mshList
        If .MouseRow = 0 Then Exit Sub
        If .TextMatrix(.Row, .ColIndex("���ݺ�")) = "" Then Exit Sub
        Call SetRowSelected(.Row, Trim(.TextMatrix(.Row, .ColIndex("ѡ��"))) = "")
    End With
    Call SetSumMoney
    
End Sub
Private Sub mshList_KeyPress(KeyAscii As Integer)
     If KeyAscii <> 32 Then Exit Sub
    With mshList
        If .TextMatrix(.Row, .ColIndex("���ݺ�")) = "" Then Exit Sub
       Call SetRowSelected(.Row, Trim(.TextMatrix(.Row, .ColIndex("ѡ��"))) = "")
    End With
    Call SetSumMoney
End Sub

Private Sub cmdAll_Click(Index As Integer)
    Dim i As Long
    
    With mshList
        .Redraw = False
        For i = 1 To .Rows - 1
            If Not SetRowSelected(i, Index = 0) Then
                .Row = i: .Col = 0: .ColSel = .Cols - 1
                Call mshList_AfterRowColChange(0, 0, .Row, .Col)
                Exit For
            End If
        Next
        .Redraw = True
    End With
    Call SetSumMoney(Index = 1)
End Sub

Private Function CheckInsureCancel(ByVal lng����ID As Long, ByVal lngInsure As Long, _
    ByVal strno As String, Optional ByVal bln������ As Long) As String
    '���ҽ���Ƿ��ܹ�ԭ������
    '���أ�����ԭ�����ϣ��򷵻ؿգ����򣬷�����ʾ��Ϣ
    Dim strTmp As String, i As Integer
    Dim arrBalanceType As Variant, strBalanceType As String
    
    On Error GoTo ErrHandler
    If Not gclsInsure.GetCapability(support�����������, lng����ID, lngInsure) Then
        CheckInsureCancel = IIf(bln������, "ҽ���������", "") & "����[" & strno & "]�Ĳ������಻֧������������ϣ�������ת����"
        Exit Function
    Else
        '���жϸõ��ݵ�ÿ�ֽ��㷽ʽ�Ƿ�֧��,�����˷�ʱ,������Ϊָ�����㷽ʽ,�˴��򻯹���Ϊ�������˷�
        strTmp = GetBalanceType(strno, bln������)
        arrBalanceType = Split(strTmp, ",")
        For i = 0 To UBound(arrBalanceType)
            strBalanceType = arrBalanceType(i)
            If Not gclsInsure.GetCapability(support�����������, lng����ID, lngInsure, strBalanceType) Then
                CheckInsureCancel = IIf(bln������, "ҽ���������", "") & "����[" & strno & "]�Ĳ������಻֧��" & strBalanceType & "�������ϣ�������ת����"
                Exit Function
            End If
        Next
    End If
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SetRowSelected(ByVal lngRow As Long, blnSelect As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����һ�е�ѡ��״̬
    '       ����Ƕ��ŵ����е�һ��,����ͬʱ���ö����е���������
    '����:���˺�
    '����:2011-02-21 16:10:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intInsure As Integer, strno As String, strTmp As String
    Dim str���� As String
    
    With mshList
        If .TextMatrix(lngRow, .ColIndex("���")) = "��ת��" And .TextMatrix(lngRow, .ColIndex("ѡ��")) <> IIf(blnSelect, "��", "") Then
            intInsure = Val(.TextMatrix(lngRow, .ColIndex("����")))
            str���� = Trim(.TextMatrix(lngRow, .ColIndex("����")))
            strno = .TextMatrix(lngRow, .ColIndex("���ݺ�"))
            
            If intInsure > 0 And blnSelect And str���� = "�շѵ�" Then
                strTmp = CheckInsureCancel(mlngPatient, intInsure, strno)
                If strTmp <> "" Then
                    sta.Panels(2).Text = strTmp
                    .TextMatrix(lngRow, .ColIndex("ѡ��")) = ""
                    Exit Function
                End If
            End If
            
            .TextMatrix(lngRow, .ColIndex("ѡ��")) = IIf(blnSelect, "��", "")
            If str���� = "�շѵ�" Then
                If intInsure > 0 Then      'ȫ��ѡ���ȡ��
                    If gclsInsure.GetCapability(support�൥���շѱ���ȫ��, mlngPatient, intInsure) _
                        Or Not IsYBSingle(strno, intInsure) Then
                        If Not SetMultiOther(lngRow, blnSelect, intInsure) Then Exit Function
                    End If
                Else '�ֽ�����Ҫ����൥���շ����
                    If Not SetMultiOther(lngRow, blnSelect, intInsure) Then Exit Function
                End If
            End If
        End If
        If .TextMatrix(lngRow, .ColIndex("���")) = "����ת��" Then .TextMatrix(lngRow, .ColIndex("ѡ��")) = ""
    End With
    SetRowSelected = True
End Function

Private Function CheckAllTurn(ByVal strno As String) As Boolean
    Dim strSQL As String, rsData As ADODB.Recordset
    strSQL = "Select 1" & vbNewLine & _
            " From ����Ԥ����¼ A," & vbNewLine & _
            "     (Select Distinct ����id" & vbNewLine & _
            "       From ������ü�¼" & vbNewLine & _
            "       Where NO In (Select Distinct NO" & vbNewLine & _
            "                    From ������ü�¼" & vbNewLine & _
            "                    Where ����id In" & vbNewLine & _
            "                          (Select ����id" & vbNewLine & _
            "                           From ����Ԥ����¼" & vbNewLine & _
            "                           Where ������� In (Select b.�������" & vbNewLine & _
            "                                          From ������ü�¼ A, ����Ԥ����¼ B" & vbNewLine & _
            "                                          Where a.No = [1] And a.��¼���� = 1 And a.��¼״̬ <> 0 And a.����id = b.����id))) And" & vbNewLine & _
            "             ��¼���� = 1 And ��¼״̬ <> 0) B" & vbNewLine & _
            " Where a.����id = b.����id And a.��¼���� = 3 And (Exists (Select 1 From ҽ�ƿ���� Where ID = a.�����id And �Ƿ�ȫ�� = 1) Or Exists" & vbNewLine & _
            "       (Select 1 From ���ѿ����Ŀ¼ Where ��� = a.���㿨��� And �Ƿ�ȫ�� = 1))" & vbNewLine & _
            " Group By ���㷽ʽ" & vbNewLine & _
            " Having Sum(��Ԥ��) <> 0"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strno)
    If rsData.EOF Then
        CheckAllTurn = False
    Else
        CheckAllTurn = True
    End If
End Function

Private Function SetMultiOther(ByVal lngRow As Long, blnSelect As Boolean, intInsure As Integer) As Boolean
'����:���ŵ�������ѡ���ȡ��
'     ���ҽ�����ŵ���Ҫ�������˷�,ѡ������һ��ʱ,ȫѡ����,ȡ��ʱȫȡ��
    Dim i As Long, j As Long, k As Long, strno As String, strTmp As String
    Dim strBalanceType As String, arrBalanceType As Variant, blnAllTurn As Boolean
    Dim str���� As String, strReplenishNo As String, strNotSelectNos As String
    Dim strNos As String, varNos As Variant
    
    With mshList
        str���� = .TextMatrix(lngRow, .ColIndex("����"))
        If str���� = "���ʵ�" Then SetMultiOther = True: Exit Function
        If intInsure = 0 Then
            '����Ƿ�Ϊ�����㵥��
            If CheckBillExistReplenishData(1, , .TextMatrix(lngRow, .ColIndex("���ݺ�")), strReplenishNo) Then
                If mbln�������� Then
                    strNos = GetReplenishAllNos(strReplenishNo)
                    varNos = Split(strNos, ",")
                    For i = 0 To UBound(varNos)
                        For k = 1 To .Rows - 1
                            If .TextMatrix(k, .ColIndex("����")) = str���� And .TextMatrix(k, .ColIndex("���ݺ�")) = varNos(i) Then
                                .TextMatrix(k, .ColIndex("ѡ��")) = IIf(blnSelect, "��", "")
                                Exit For
                            End If
                        Next
                    Next
                    SetMultiOther = True
                    Exit Function
                End If
            End If
            
            If CheckAllTurn(.TextMatrix(lngRow, .ColIndex("���ݺ�"))) = True Then
                blnAllTurn = True
            Else
                blnAllTurn = False
            End If
            If gblnMultiBalance Or blnAllTurn Then     '   �൥��,���ֽ��㷽ʽ
                '33635:ԭ���Ƕ൥���Ҷ��ֽ��㷽ʽ,���ܲ�����
                strno = ""
                For k = 1 To .Rows - 1
                      If .TextMatrix(k, .ColIndex("����ID")) = .TextMatrix(lngRow, .ColIndex("����ID")) _
                            And .TextMatrix(k, .ColIndex("����")) = str���� _
                            And Trim(.TextMatrix(lngRow, .ColIndex("����ID"))) <> "" Then
                            strno = strno & "," & .TextMatrix(k, .ColIndex("���ݺ�"))
                      End If
                Next
                If strno <> "" Then strno = Mid(strno, 2)
                If InStr(1, strno, ",") > 0 Then    '֤��Ϊ�൥��
                    '����������,�����˵Ļ�,Ʊ���ջش�������
                    'If CheckSingleBalance(strNO) = False Then    '�Ƕ��ֽ��㷽ʽ,�������˷�,'ȫѡ
                        For k = 1 To .Rows - 1
                              If .TextMatrix(k, .ColIndex("����ID")) = .TextMatrix(lngRow, .ColIndex("����ID")) _
                                And .TextMatrix(k, .ColIndex("����")) = str���� _
                                And Trim(.TextMatrix(lngRow, .ColIndex("����ID"))) <> "" Then
                                    .TextMatrix(k, .ColIndex("ѡ��")) = IIf(blnSelect, "��", "")
                              End If
                        Next
                    'End If
                End If
            End If
            '����Ƿ�������ѿ��Ľ���,�������,�ֲ�֧���ⲿ�����ݵĴ���
            If strno = "" Then strno = .TextMatrix(lngRow, .ColIndex("���ݺ�"))
'            If zlIsExistsSquareCard(strNO) Then
'                sta.Panels(2).Text = "�ݲ�֧�ֶ����ѿ����ݵ�ת��!"
'                For k = 1 To .Rows - 1
'                      If .TextMatrix(k, .ColIndex("����ID")) = .TextMatrix(lngRow, .ColIndex("����ID")) _
'                        And .TextMatrix(k, .ColIndex("����")) = str���� _
'                        And Trim(.TextMatrix(lngRow, .ColIndex("����ID"))) <> "" Then
'                            .TextMatrix(k, .ColIndex("ѡ��")) = ""
'                      End If
'                Next
'            End If
            '����Ƿ�������ѿ�,����൥���д������ѿ�,Ҳ����ȫѡ
            SetMultiOther = True
            Exit Function
        End If
        
        If IsYBSingle(.TextMatrix(lngRow, .ColIndex("���ݺ�")), intInsure) Then SetMultiOther = True: Exit Function
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("���")) = "��ת��" _
                And .TextMatrix(i, .ColIndex("����ID")) = .TextMatrix(lngRow, .ColIndex("����ID")) _
                And i <> lngRow Then
                If .TextMatrix(i, .ColIndex("ѡ��")) <> .TextMatrix(lngRow, .ColIndex("ѡ��")) Then
                   If intInsure <> 0 And blnSelect Then
                        strno = .TextMatrix(i, .ColIndex("���ݺ�"))
                        '�жϸõ��ݵ�ÿ�ֽ��㷽ʽ�Ƿ�֧��,�����˷�ʱ,������Ϊָ�����㷽ʽ,�˴��򻯹���Ϊ�������˷�
                         strTmp = GetBalanceType(strno)
                         If strTmp <> "" Then
                             arrBalanceType = Split(strTmp, ",")
                             For j = 0 To UBound(arrBalanceType)
                                 strBalanceType = arrBalanceType(j)
                                 If Not gclsInsure.GetCapability(support�����������, mlngPatient, intInsure, strBalanceType) Then
                                     sta.Panels(2).Text = "����[" & strno & "]�Ĳ������಻֧��" & strBalanceType & "����,���в�����ѡ��ת��!"
                                     For k = 1 To .Rows - 1
                                        If .TextMatrix(k, .ColIndex("����ID")) = .TextMatrix(i, .ColIndex("����ID")) _
                                            And .TextMatrix(k, .ColIndex("����")) = str���� Then
                                            .TextMatrix(k, .ColIndex("ѡ��")) = ""
                                        End If
                                     Next
                                     Exit Function
                                 End If
                             Next
                         End If
                    End If
                    .TextMatrix(i, .ColIndex("ѡ��")) = IIf(blnSelect, "��", "")
                End If
            End If
        Next
    End With
    SetMultiOther = True
End Function

Private Function zlIsExistsSquareCard(ByVal strNos As String, Optional int��¼���� As Integer = 3) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���õ����Ƿ�Ϊ�����㵥��
    '���:strNos-���ݺ�(����Ϊ����,�ö��ŷ���)
    '       int��¼����:3-�����շ�;2-����
    '����:
    '����:����,�򷵻�true,���򷵻�False
    '����:���˺�
    '����:2010-01-11 12:04:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String, strNoIns As String
    
    On Error GoTo errHandle
    
    strNoIns = Replace(strNos, "'", "")
    strSQL = "" & _
    "   Select /*+ rule */ A.ID As ������id " & _
    "   From ���˿������¼ A, ����Ԥ����¼ B, ������ü�¼ C,Table( f_Str2list([1])) J " & _
    "   Where A.����id = B.ID and B.��¼����=[2] And C.NO = J.Column_Value And C.����ID = B.����ID And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����շѵ��Ƿ����ˢ����¼", strNoIns, int��¼����)
    zlIsExistsSquareCard = rsTemp.EOF = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function zlGetOneCard(ByVal strIDs As String) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���õ����Ƿ�һ��ͨ���㵥��
    '���:strIDs-����ID(����Ϊ����,�ö��ŷ���)
    '����:
    '����:һ��ͨ��������,�򷵻�true,���򷵻�False
    '����:���˺�
    '����:2010-01-11 12:04:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    
    On Error GoTo errHandle
    
     strSQL = "" & _
    "   Select /*+ rule */  A.����ID,A.��λ�ʺ�, A.�������, B.ҽԺ����, A.��Ԥ�� as ���" & vbNewLine & _
    "   From ����Ԥ����¼ A, һ��ͨĿ¼ B,Table( f_Num2list([1])) J " & vbNewLine & _
    "   Where A.����id = J.Column_Value  And A.���㷽ʽ = B.���㷽ʽ" & _
    "   Order by ����ID"
    Set zlGetOneCard = zlDatabase.OpenSQLRecord(strSQL, "��ȡһ��ͨ��������", strIDs)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function zlGetThreeCard(ByVal strIDs As String) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���õ����Ƿ����������㵥��
    '���:strIDs-����ID(����Ϊ����,�ö��ŷ���)
    '����:
    '����:��������������,�򷵻�true,���򷵻�False
    '����:������
    '����:2015-12-29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    
    On Error GoTo errHandle
    
     strSQL = "" & _
    "   Select /*+ rule */  A.����ID, A.��Ԥ�� as ���, B.���� " & vbNewLine & _
    "   From ����Ԥ����¼ A, ҽ�ƿ���� B,Table( f_Num2list([1])) J " & vbNewLine & _
    "   Where A.����id = J.Column_Value  And A.���㷽ʽ = B.���㷽ʽ" & _
    "   Order by ����ID"
    Set zlGetThreeCard = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��������������", strIDs)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckSingleBalance(ByVal strno As String) As Boolean
'���ܣ��ж�ָ���������Ƿ�ֻ��һ�ַ�ҽ�����㷽ʽ(��Ԥ������)
'       :strNO(��ʽΪ"E01,E02"):����:34035
'������
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strno = Replace(strno, "'", "")
    CheckSingleBalance = True
    
    strSQL = "" & _
    " Select /*+ rule */ Count(Distinct A.���㷽ʽ) num" & vbNewLine & _
    " From ����Ԥ����¼ A, ���㷽ʽ B,Table( f_Str2list([1])) J" & vbNewLine & _
    " Where   A.��¼���� = 3 And A.��¼״̬ In (1, 3) " & _
    "           And A.���㷽ʽ = B.���� And B.���� In (1, 2)  And A.NO = J.Column_Value"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strno)
    If rsTmp!num > 1 Then CheckSingleBalance = False
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetBalanceType(ByVal strno As String, _
    Optional ByVal bln������ As Boolean) As String
    '����:��ȡһ�ŵ����е�ҽ�����㷽ʽ��
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim i As Long
        
    On Error GoTo errH
    If bln������ Then
        strSQL = _
            " Select Distinct a.���㷽ʽ" & vbNewLine & _
            " From ����Ԥ����¼ A, ���㷽ʽ B, ���ò����¼ C" & vbNewLine & _
            " Where a.���㷽ʽ = b.���� And a.��¼���� = 6 And b.���� In(3,4)" & vbNewLine & _
            "       And a.����id = c.����id And c.��¼���� = 1" & vbNewLine & _
            "       And c.���ӱ�־ = 0 And Nvl(c.����״̬, 0) <> 2 And c.No = [1]"
    Else
        strSQL = _
            " Select Distinct a.���㷽ʽ" & vbNewLine & _
            " From ����Ԥ����¼ A, ���㷽ʽ B, ������ü�¼ C" & vbNewLine & _
            " Where a.���㷽ʽ = b.���� And b.���� In(3,4)" & vbNewLine & _
            "       And a.����id = c.����ID And c.��¼���� = 1 And c.No = [1]"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strno)
    Do While Not rsTmp.EOF
        GetBalanceType = GetBalanceType & "," & rsTmp!���㷽ʽ
        rsTmp.MoveNext
    Loop
    GetBalanceType = Mid(GetBalanceType, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ShowDetail(ByVal str���� As String, ByVal strno As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ��ϸ����
    '���:str����:�շѵ�(���ʵ�)
    '        strNO-���ݺ�
    '����:���˺�
    '����:2011-02-22 11:14:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, i As Long, strSQL As String
    Err = 0: On Error GoTo errH
    If mshList.Row < 0 Then Exit Sub
    
    If mshList.TextMatrix(mshList.Row, mshList.ColIndex("���")) = "��ת��" Then
        strSQL = "Select C.���� As ���, Nvl(E.����, B.����) As ����, B.���, A.���㵥λ As ��λ, Sum(Nvl(A.����, 1) * A.����) As ����," & vbNewLine & _
                "       LTrim(To_Char(A.��׼����, '999990.00000')) As ����, LTrim(To_Char(Sum(A.Ӧ�ս��), '99999" & gstrDec & "')) As Ӧ�ս��," & vbNewLine & _
                "       LTrim(To_Char(Sum(A.ʵ�ս��), '99999" & gstrDec & "')) As ʵ�ս��, D.���� As ִ�п���, 3 As ��¼״̬" & vbNewLine & _
                "From ������ü�¼ A, �շ���ĿĿ¼ B, �շ���Ŀ��� C, ���ű� D, �շ���Ŀ���� E" & vbNewLine & _
                "Where A.�շ�ϸĿid = B.ID And A.�շ���� = C.���� And A.ִ�в���id = D.ID(+) And A.NO = [1] And Mod(A.��¼����,10) = [2] And" & vbNewLine & _
                "      A.��¼״̬ In (2,3) And A.�շ�ϸĿid = E.�շ�ϸĿid(+) And E.����(+) = 1 And E.����(+) = 3 And A.���ӱ�־ <> 9 " & vbNewLine & _
                "Group By A.��׼����,A.���, C.����, Nvl(E.����, B.����), B.���, A.���㵥λ, D.���� Having Sum(A.����) <> 0 " & vbNewLine & _
                " Union " & vbNewLine & _
                "Select C.���� As ���, Nvl(E.����, B.����) As ����, B.���, A.���㵥λ As ��λ, Sum(Nvl(A.����, 1) * A.����) As ����," & vbNewLine & _
                "       LTrim(To_Char(A.��׼����, '999990.00000')) As ����, LTrim(To_Char(Sum(A.Ӧ�ս��), '99999" & gstrDec & "')) As Ӧ�ս��," & vbNewLine & _
                "       LTrim(To_Char(Sum(A.ʵ�ս��), '99999" & gstrDec & "')) As ʵ�ս��, D.���� As ִ�п���, 1 As ��¼״̬" & vbNewLine & _
                "From ������ü�¼ A, �շ���ĿĿ¼ B, �շ���Ŀ��� C, ���ű� D, �շ���Ŀ���� E" & vbNewLine & _
                "Where A.�շ�ϸĿid = B.ID And A.�շ���� = C.���� And A.ִ�в���id = D.ID(+) And A.NO = [1] And Mod(A.��¼����,10) = [2] And" & vbNewLine & _
                "      A.��¼״̬=1 And A.�շ�ϸĿid = E.�շ�ϸĿid(+) And E.����(+) = 1 And E.����(+) = 3 And A.���ӱ�־ <> 9 " & vbNewLine & _
                "Group By A.��׼����,A.���, C.����, Nvl(E.����, B.����), B.���, A.���㵥λ, D.���� Having Sum(A.����) <> 0 " & vbNewLine
    
    ElseIf mshList.TextMatrix(mshList.Row, mshList.ColIndex("���")) = "����ת��" Then
        strSQL = "Select C.���� As ���, Nvl(E.����, B.����) As ����, B.���, A.���㵥λ As ��λ, Sum(Nvl(A.����, 1) * A.����) As ����," & vbNewLine & _
                "       LTrim(To_Char(A.��׼����, '999990.00000')) As ����, LTrim(To_Char(Sum(A.Ӧ�ս��), '99999" & gstrDec & "')) As Ӧ�ս��," & vbNewLine & _
                "       LTrim(To_Char(Sum(A.ʵ�ս��), '99999" & gstrDec & "')) As ʵ�ս��, D.���� As ִ�п���, 2 As ��¼״̬" & vbNewLine & _
                "From ������ü�¼ A, �շ���ĿĿ¼ B, �շ���Ŀ��� C, ���ű� D, �շ���Ŀ���� E" & vbNewLine & _
                "Where A.�շ�ϸĿid = B.ID And A.�շ���� = C.���� And A.ִ�в���id = D.ID(+) And A.NO = [1] And Mod(A.��¼����,10) = [2] And" & vbNewLine & _
                "      A.��¼״̬ In (1,3) And A.�շ�ϸĿid = E.�շ�ϸĿid(+) And E.����(+) = 1 And E.����(+) = 3 And A.���ӱ�־ <> 9 " & vbNewLine & _
                "Group By A.��׼����,A.���, C.����, Nvl(E.����, B.����), B.���, A.���㵥λ, D.���� Having Sum(A.����) <> 0 " & vbNewLine
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strno, IIf(str���� = "���ʵ�", 2, 1))
    
    mshDetail.Redraw = flexRDNone
    mshDetail.Clear
    Set mshDetail.DataSource = rsTmp
    If rsTmp.EOF Then mshDetail.Rows = 2
    Call SetDetail
    mshDetail.Redraw = flexRDBuffered
    Exit Sub
errH:
    mshDetail.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetDetail()
    Dim strHead As String
    Dim i As Long
    strHead = "���,1,650|����,1,1500|���,1,1450|��λ,4,500|����,7,500|����,7,850|Ӧ�ս��,7,850|ʵ�ս��,7,850|ִ�п���,4,1000|��¼״̬,4,0"
    With mshDetail
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColKey(i) = .TextMatrix(0, i)
            .colAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .FixedAlignment(i) = 4
        Next
        .ColHidden(9) = True
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, 9)) = 1 Then .Cell(flexcpForeColor, i, 0, i, 9) = vbBlack
            'If Val(.TextMatrix(i, 9)) = 2 Then .Cell(flexcpForeColor, i, 0, i, 9) = vbRed
            If Val(.TextMatrix(i, 9)) = 3 Then .Cell(flexcpForeColor, i, 0, i, 9) = vbBlue
        Next i
        zl_vsGrid_Para_Restore 1131, mshDetail, Me.Caption, "����תסԺ��ϸ�б�", True
        .RowHeight(0) = 320
        .Col = 0: .ColSel = .Cols - 1
    End With
End Sub

Private Sub SetBalanceHead()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����б�
    '����:���˺�
    '����:2011-03-28 11:27:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strHead As String
    Dim i As Long
    strHead = "���,4,650|��־,1,600|���㵥��,1,1500|������,7,1000|���㷢Ʊ,1, 2600"
    With vsBalance
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColKey(i) = .TextMatrix(0, i)
            .colAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .FixedAlignment(i) = 4
        Next
        zl_vsGrid_Para_Restore 1131, vsBalance, Me.Caption, "����תסԺ�����б�", True
        .RowHeight(0) = 320
        .Col = 0: .ColSel = .Cols - 1
    End With
End Sub

Private Sub picBill_Resize()
    Err = 0: On Error Resume Next
    With picBill
        mshList.Left = .ScaleLeft
        mshList.Top = .ScaleTop
        mshList.width = .ScaleWidth
        mshList.Height = .ScaleHeight
    End With
End Sub

Private Sub picBalance_Resize()
    Err = 0: On Error Resume Next
    With picBalance
        vsBalance.Left = .ScaleLeft
        vsBalance.Top = .ScaleTop
        vsBalance.width = .ScaleWidth
        lblSum.Top = .ScaleHeight - lblSum.Height
        vsBalance.Height = lblSum.Top - mshDetail.Top
    End With
End Sub

Private Sub picBottom_Resize()
    Err = 0: On Error Resume Next
    With picBottom
            cmdExit.Left = .ScaleLeft + .ScaleWidth - cmdExit.width - 100
            cmdSave.Left = cmdExit.Left - cmdSave.width - 20
            cmdSave.Top = cmdExit.Top
    End With
End Sub

Private Sub picList_Resize()
    Err = 0: On Error Resume Next
    With picList
        mshDetail.Left = .ScaleLeft
        mshDetail.Top = .ScaleTop
        mshDetail.width = .ScaleWidth
        mshDetail.Height = .ScaleHeight
    End With
End Sub

Private Sub picTop_Resize()
    Err = 0: On Error Resume Next
    If mblnSelPati Then
        fraPati.Left = picTop.ScaleLeft
        lblDate.Left = fraPati.Left + fraPati.width + 20
        dtpBegin.Left = lblDate.Left + lblDate.width + 10
        lbl��.Left = dtpBegin.Left + dtpBegin.width + 20
        dtpEnd.Left = lbl��.Left + lbl��.width + 20
    End If
    chkShow.Left = IIf(dtpEnd.Visible, dtpEnd.Left + dtpEnd.width, (fraPati.Left + fraPati.width) * IIf(fraPati.Visible = False, 0, 1) + 50)
    cmdRefresh.Left = chkShow.Left + chkShow.width + 50
End Sub

Private Sub txtPatient_Change()
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
    Call IDKind.SetAutoReadCard(txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
End Sub

Private Sub txtPatient_GotFocus()
    zlControl.TxtSelAll txtPatient
    If Not mobjIDCard Is Nothing And Not txtPatient.Locked Then Call mobjIDCard.SetEnabled(txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
    If Not mobjICCard Is Nothing And Not txtPatient.Locked Then Call mobjICCard.SetEnabled(txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
    If Not txtPatient.Locked Then Call IDKind.SetAutoReadCard(txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim blnCard As Boolean, strInput As String
    If txtPatient.Locked Then Exit Sub
    '����ѡ����
    If Not (Len(Trim(Me.txtPatient.Text)) = 0 And KeyAscii = 13) Then
       If IDKind.GetCurCard.���� Like "����*" Then
            blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.ShowPassText)
        ElseIf IDKind.GetCurCard.���� = "�����" Or IDKind.GetCurCard.���� = "סԺ��" Then
            If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
                If InStr("0123456789-*+", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
            End If
        Else
            txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
            '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
            txtPatient.IMEMode = 0
        End If
    End If
    
    Me.Refresh
    'ˢ����ϻ���������س�
    If blnCard And Len(txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Trim(txtPatient.Text) <> "" Then
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        End If
        KeyAscii = 0
        Call FindPati(IDKind.GetCurCard, blnCard, Trim(txtPatient.Text))
    End If
End Sub

Private Sub FindPati(ByVal objCard As Card, ByVal blnCard As Boolean, ByVal strInput As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���Ҳ���
    '����:���˺�
    '����:2012-10-18 16:35:27
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If GetPatient(objCard, strInput, blnCard) Then
        '69526:������,2014-02-13,��Ժ�����޷���������תסԺ����
        If Val(zlDatabase.GetPara("��Ժ������������תסԺ", glngSys, 1137, "0")) = 0 Then
            If HaveOut(mlngPatient) = True Then
                MsgBox "����" & mrsInfo!���� & "�Ѿ���Ժ��δ����סԺ������������������תסԺ������", vbInformation, gstrSysName
                txtPatient.Text = "": mlngPatient = 0
                Call ClearData
                Set mrsInfo = Nothing
                If txtPatient.Visible And txtPatient.Enabled Then Call txtPatient.SetFocus: zlControl.TxtSelAll txtPatient
                Exit Sub
            End If
        End If
        '��ʱ������ʽ�����¼�Form_Load
        Call ShowBills(mlngPatient, dtpBegin.Value, dtpEnd.Value)
        If mshList.TextMatrix(1, mshList.ColIndex("���ݺ�")) <> "" Then
            If mshList.TextMatrix(1, mshList.ColIndex("ѡ��")) <> "" Then
                If cmdSave.Visible And cmdSave.Enabled Then Call cmdSave.SetFocus
            Else
                If cmdAll(0).Visible And cmdAll(0).Enabled Then Call cmdAll(0).SetFocus
            End If
        Else
            If txtPatient.Visible And txtPatient.Enabled Then Call txtPatient.SetFocus: zlControl.TxtSelAll txtPatient
        End If
    Else
        txtPatient.Text = "": mlngPatient = 0
        Call ClearData
        If txtPatient.Visible And txtPatient.Enabled Then Call txtPatient.SetFocus: zlControl.TxtSelAll txtPatient
    End If
    txtPatient.PasswordChar = ""
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txtPatient.IMEMode = 0
End Sub

Private Sub txtPatient_LostFocus()
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(False)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(False)
    Call IDKind.SetAutoReadCard(False)
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
    If mrsInfo Is Nothing Then Exit Sub
    If mrsInfo.State <> 1 Then Exit Sub
    If txtPatient.Text <> mrsInfo!���� Then txtPatient.Text = mrsInfo!����
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    If txtPatient.Text <> "" Or txtPatient.Locked Or Not Me.ActiveControl Is txtPatient Then Exit Sub
    Dim objCard As Card '54894
    Set objCard = IDKind.GetIDKindCard("���֤", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    txtPatient.Text = strID
    If txtPatient.Text <> "" Then Call FindPati(objCard, True, txtPatient.Text)
End Sub

Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, ByVal blnCard As Boolean, Optional ByVal lng��ҳID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ
    '���:blnCard=�Ƿ���￨ˢ��,lng��ҳID=��ȡָ��סԺ�����Ĳ�����Ϣ
    '����:
    '����:�Ƿ��ȡ�ɹ�,�ɹ�ʱmrsInfo�а���������Ϣ,ʧ��ʱmrsInfo=Close,strInput�����������ж��Ƿ�����ʾ��,�����ٴ���ʾû���ҵ�����
    '����:���˺�
    '����:2010-11-09 17:17:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, lng�����ID As Long, strPassWord As String, strErrMsg As String
    Dim lng����ID As Long, blnHavePassWord As Boolean
    Dim strPati As String, strRange As String
    Dim vRect As RECT
    Dim blnCancel As Boolean
    
    On Error GoTo errH
    strSQL = _
    " Select A.����ID,Nvl(B.��ҳID,0) as ��ҳID,A.סԺ��,A.��ǰ����,B.��Ժ����ID,B.��Ժ����ID,B.��Ժ����," & _
    "        Nvl(b.����, a.����) As ����, Nvl(b.�Ա�, a.�Ա�) As �Ա�,Nvl(B.����,A.����) as ����,A.IC����,A.���￨��,A.����֤��," & _
    "       Nvl(B.�ѱ�,A.�ѱ�) as �ѱ�,C.���� as ��ǰ����,A.��ǰ����ID,D.���� as ��Ժ����,B.��Ժ����ID,A.���� as ����,E.����,E.ҽ����,E.����," & _
    "       A.�Ǽ�ʱ��,Nvl(B.״̬,0) as ״̬,Nvl(B.ҽ�Ƹ��ʽ,A.ҽ�Ƹ��ʽ) as ҽ�Ƹ��ʽ,Nvl(B.��˱�־,0) as ��˱�־,B.��Ժ����,B.��Ժ����,B.��������,B.��������" & _
    " From ������Ϣ A,������ҳ B,���ű� C,���ű� D,ҽ�����˵��� E,ҽ�����˹����� F" & _
    " Where A.ͣ��ʱ�� is NULL And A.����ID=B.����ID(+) And " & IIf(lng��ҳID = 0, "A.��ҳID=B.��ҳID(+)", "B.��ҳID=[3]") & _
    "           And A.����ID=F.����ID(+) And F.��־(+)=1 And F.ҽ����=E.ҽ����(+) And F.����=E.����(+) And F.���� = E.����(+)" & _
    "           And A.��ǰ����ID=C.ID(+) And B.��Ժ����ID=D.ID(+) "
        
    If blnCard = True And objCard.���� Like "����*" Then  'ˢ��
        lng�����ID = IDKind.GetDefaultCardTypeID
        If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
        If lng����ID <= 0 Then GoTo NotFoundPati:
        strInput = "-" & lng����ID
        blnHavePassWord = True
        strSQL = strSQL & " And A.����ID=[1]"
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then '����ID
        strSQL = strSQL & " And A.����ID=[1]"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '�����
        strSQL = strSQL & " And A.�����=[1]"
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then 'סԺ��
        strSQL = strSQL & " And A.����ID = (Select Max(����id) From ������ҳ Where סԺ�� = [1])"
    Else '��������
        Select Case objCard.����
            Case "����", "��������￨"
                If Not mrsInfo Is Nothing Then
                    If mrsInfo.State = 1 Then
                        If mrsInfo!���� = Trim(txtPatient.Text) Then
                            mlngPatient = Val(Nvl(mrsInfo!����ID))
                            GetPatient = True
                            Exit Function
                        End If
                    End If
                End If
                If mintPatientRange > 0 Then
                    Select Case mintPatientRange
                        Case 1  '�κη���δ���岡��
                            strRange = ""
                        Case 2  '���δ����Ĳ���
                            strRange = " And C.��Դ;�� = 4"
                        Case 3  'סԺδ����Ĳ���
                            strRange = " And C.��Դ;�� = 2"
                        Case 4  '����δ����Ĳ���
                            strRange = " And C.��Դ;�� = 1"
                    End Select
                    strPati = " And Exists(Select 1 From ����δ����� C Where C.����id=A.����ID And Nvl(C.��ҳID,0)=A.��ҳID" & strRange & ")"
                End If
                 'ͨ����������
                strPati = "Select A.����ID as ID,A.����ID,A.סԺ��, A.�����, Nvl(b.�Ա�, a.�Ա�) As �Ա�, Nvl(b.����, a.����) as ����, A.סԺ����, A.��ͥ��ַ, A.������λ," & vbNewLine & _
                        "To_Char(A.��������,'YYYY-MM-DD') as ��������,  To_Char(B.��Ժ����,'YYYY-MM-DD') as ��Ժ����, To_Char(B.��Ժ����,'YYYY-MM-DD') as ��Ժ����" & vbNewLine & _
                        "From ������Ϣ A, ������ҳ B,��Ժ���� C" & vbNewLine & _
                        "Where A.����id = B.����id(+) And A.��ҳID = B.��ҳid(+) And A.ͣ��ʱ�� Is Null And A.����ID=C.����ID And A.���� = [1] " & vbNewLine & strPati & vbNewLine & _
                        "Order By Decode(סԺ��, Null, 1, 0), ��Ժ���� Desc"
                vRect = zlControl.GetControlRect(txtPatient.hWnd)
                Set mrsInfo = zlDatabase.ShowSQLSelect(Me, strPati, 0, "���˲���", 1, "", "��ѡ����", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, False, True, strInput)
                            
                If Not mrsInfo Is Nothing Then
                    strInput = Val(mrsInfo!����ID)
                    strSQL = strSQL & " And A.����ID=[2]"
                Else
                    Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "": Exit Function
                End If
            Case "ҽ����"
                strInput = UCase(strInput)
                strSQL = strSQL & " And A.ҽ����=[2]"
            Case "�����"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = strSQL & " And A.�����=[2]"
            Case "סԺ��"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = strSQL & " And A.����ID = (Select Max(����id) From ������ҳ Where סԺ�� = [2])"
            Case Else
                '��������,��ȡ��صĲ���ID
                If objCard.�ӿ���� > 0 Then
                    lng�����ID = objCard.�ӿ����
                    If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                    If lng����ID = 0 Then GoTo NotFoundPati:
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(objCard.����, strInput, False, lng����ID, _
                        strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                End If
                If lng����ID <= 0 Then GoTo NotFoundPati:
                strSQL = strSQL & " And A.����ID=[1]"
                strInput = "-" & lng����ID
                blnHavePassWord = True
        End Select
    End If
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput, lng��ҳID)
    If mrsInfo.EOF Then Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "": Exit Function
    
    txtPatient.Text = Nvl(mrsInfo!����): mlngPatient = Val(Nvl(mrsInfo!����ID))
    If IsDate(Format(mrsInfo!��Ժ����, "yyyy-mm-dd HH:MM:SS")) Then
        '�������Ϊ��Ժ����,����ת��סԺ�����е��������
        dtpEnd.MaxDate = CDate(Format(mrsInfo!��Ժ����, "yyyy-mm-dd 23:59:59"))
        dtpEnd.Value = dtpEnd.MaxDate
        dtpEnd.MaxDate = dtpEnd.MaxDate + 1
        dtpBegin.MaxDate = dtpEnd.Value
        '   ����: 36609����Ժʱ��Ҫ��һ��,��Ϊ���ܴ��ڲ�����û���������ʱ,����Ժ,��ȥ�������,�Ӷ�����������ת���˵����.
    End If
    GetPatient = True
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
  
Private Function PrintPrePayPrint(ByVal frmMain As Object, ByVal strDelDate As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ӡԤ����
    '���:strDelDate-����ת������
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-02-16 10:30:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, bytPrepayPrint As Byte
    Dim strNos As String
    
    On Error GoTo errHandle
    If InStr(1, mstrPrivs, ";Ԥ�����վݴ�ӡ;") = 0 Then
       PrintPrePayPrint = True: Exit Function '����ӡ
    End If
    bytPrepayPrint = Val(zlDatabase.GetPara("����תסԺԤ����ӡ", glngSys, 1131))
    If bytPrepayPrint = 0 Then PrintPrePayPrint = True: Exit Function '����ӡ
    
    strSQL = "Select distinct NO From ����Ԥ����¼ Where ��¼����=1 and �տ�ʱ��= [1] and ժҪ='����תסԺԤ��'"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡתԤ����", CDate(strDelDate))
    If rsTemp.EOF Then
        'û��תΪԤ�����ݣ���Ҳ����ӡ
        PrintPrePayPrint = True: Exit Function
    End If
    If bytPrepayPrint = 2 Then   '��ʾ��ӡ
        If MsgBox("�����������תסԺ����ʱ�������ֽ�Ƚ��㷽ʽתΪ��Ԥ����,���Ƿ�Ҫ��ӡԤ����Ʊ�ݣ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
              PrintPrePayPrint = True: Exit Function
        End If
    End If
    
    If Val(zlDatabase.GetPara(283, glngSys, , "0")) = 1 Then '112862
        Do While Not rsTemp.EOF
            strNos = strNos & "," & Nvl(rsTemp!NO)
            rsTemp.MoveNext
        Loop
        If strNos <> "" Then strNos = Mid(strNos, 2)
        If zlPrintInvoice(strNos, strDelDate) = False Then Exit Function
    Else
        With rsTemp
            Do While Not .EOF
                If zlPrintInvoice(Nvl(rsTemp!NO), strDelDate) = False Then Exit Function
                .MoveNext
            Loop
        End With
    End If
    PrintPrePayPrint = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub SetSumMoney(Optional blnCls As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ú���ʾ�ϼ�
    '����:���˺�
    '����:2011-03-04 14:17:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, dblSumMoney As Double
    Dim strJzNOs As String, strSFNos As String
    With mshList
        If blnCls = False Then
            For i = .FixedRows To .Rows - 1
                If Trim(.TextMatrix(i, .ColIndex("ѡ��"))) <> "" Then
                    dblSumMoney = dblSumMoney + Val(.TextMatrix(i, .ColIndex("ʵ�ս��")))
                End If
                If .TextMatrix(i, .ColIndex("���")) = "��ת��" And .TextMatrix(i, .ColIndex("ѡ��")) = "��" Then
                    If .TextMatrix(i, .ColIndex("����")) = "���ʵ�" Then
                        strJzNOs = strJzNOs & "," & .TextMatrix(i, .ColIndex("���ݺ�"))
                    Else
                        strSFNos = strSFNos & "," & .TextMatrix(i, .ColIndex("���ݺ�"))
                    End If
                End If
            Next
        Else
            dblSumMoney = 0
        End If
    End With
    lblSum.Caption = "����ת���ϼ�:" & Format(dblSumMoney, "###0.00;-###0.00;0.00;0.00")
    '����ѡ�������ͨ��
    Call LoadBalance(strJzNOs, strSFNos)
End Sub

Private Sub dkpMan_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Bottom = sta.Height + picBottom.Height + 100
End Sub

Private Sub LoadBalance(ByVal strJzNOs As String, ByVal strSFNos As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ؽ�����Ϣ
    '����:���˺�
    '����:2011-03-28 11:33:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String, i As Long
    If strJzNOs = "" And strSFNos = "" Then
        With mshList
            If .TextMatrix(i, .ColIndex("���")) = "��ת��" And .TextMatrix(i, .ColIndex("ѡ��")) = "��" Then
                If .TextMatrix(i, .ColIndex("����")) = "���ʵ�" Then
                    strJzNOs = strJzNOs & "," & .TextMatrix(i, .ColIndex("���ݺ�"))
                Else
                    strSFNos = strSFNos & "," & .TextMatrix(i, .ColIndex("���ݺ�"))
                End If
            End If
        End With
    End If
    If strJzNOs = "" Then strJzNOs = ",lxh"
    strJzNOs = Mid(strJzNOs, 2)
    If strSFNos = "" Then strSFNos = ",lxh"
    strSFNos = Mid(strSFNos, 2)
    
    On Error GoTo errHandle
    '��:Wmsys.Wm_Concat��Ϊ��f_List2Str(Cast(collect ()))�ķ�ʽ.ԭ����oracle10gĿǰֻ�ǲ��԰�
    '����:38528
    
    strSQL = "" & _
    "     Select /*+ rule */  Rownum As ���, ��־, NO As ���㵥��, ������, ��Ʊ�� " & _
    "     From (Select A.��־, A.NO, A.������, f_List2str(Cast(COLLECT(distinct C.����) as t_Strlist))  As ��Ʊ�� " & _
    "            From (Select '�շ�' As ��־, A.NO, To_Char(Sum(a.���ʽ��),'9999990.00') As ������ " & _
    "                   From ������ü�¼ A, Table(f_Str2list([1])) J " & _
    "                   Where A.NO = J.Column_Value And Mod(A.��¼����,10) = 1 " & _
    "                   Group By A.NO) A, Ʊ�ݴ�ӡ���� B, Ʊ��ʹ����ϸ C " & _
    "            Where A.NO = B.NO(+) and B.��������(+)=1 And B.ID = C.��ӡid(+) " & _
    "            And C.����(+)=1 " & _
    "            Group By A.��־, A.NO, A.������ " & _
    "            Union All " & _
    "            Select A.��־, A.NO, A.������, f_List2str(Cast(COLLECT(distinct C.����) as t_Strlist)) As ��Ʊ�� " & _
    "            From (Select '����' As ��־, B.NO, To_Char(Sum(a.���ʽ��),'9999990.00') As ������ " & _
    "                   From ������ü�¼ A, ���˽��ʼ�¼ B, Table(f_Str2list([2])) J " & _
    "                   Where A.NO = J.Column_Value  And A.����id = B.ID  And B.��¼״̬=1 And A.��¼���� In (2, 12) " & _
    "                   Group By B.NO) A, Ʊ�ݴ�ӡ���� B, Ʊ��ʹ����ϸ C " & _
    "            Where A.NO = B.NO(+) and B.��������(+)=3 And B.ID = C.��ӡid(+) " & _
    "            Group By A.��־, A.NO, A.������)"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strSFNos, strJzNOs)
    Set vsBalance.DataSource = rsTemp
    If rsTemp.RecordCount = 0 Then
        vsBalance.Rows = 2
    End If
    Call SetBalanceHead
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function DelBalance(ByVal strDelDate As String, ByVal strno As String, ByVal lng����ID As Long, _
    ByVal intInsure As Integer, Optional ByRef blnTransMC As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������
    '���:strNO-���ʵ��ݺ�
    '       strDelDate:����ʱ��
    '����:blnTransMC-ҽ������
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-03-29 11:58:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strAdvance As String
    
    '��������
     '  Zl_����תסԺ����_��������
     strSQL = "Zl_����תסԺ����_��������("
     '  No_In         ���˽��ʼ�¼.NO%Type,
     strSQL = strSQL & "'" & strno & "',"
     '  ����Ա���_In ���˽��ʼ�¼.����Ա���%Type,
     strSQL = strSQL & "'" & UserInfo.��� & "',"
     '  ����Ա����_In ���˽��ʼ�¼.����Ա����%Type
     strSQL = strSQL & "'" & UserInfo.���� & "',"
     '    ��������_In   ���˽��ʼ�¼.�շ�ʱ��%Type
     strSQL = strSQL & "to_date('" & strDelDate & "','yyyy-mm-dd hh24:mi;ss'))"
     Call zlDatabase.ExecuteProcedure(strSQL, "�������תסԺ-��������")
    '���սӿ�
    blnTransMC = False
    If intInsure <> 0 Then
        If gclsInsure.CheckInsureValid(intInsure) = False Then
             Exit Function
        End If
        If gclsInsure.GetCapability(support�����������, , intInsure) Then
            strAdvance = "1|1"
            If Not gclsInsure.ClinicDelSwap(lng����ID, , intInsure, strAdvance) Then
                Exit Function
            Else
                blnTransMC = True
            End If
        Else
            MsgBox "����(" & strno & ")������֧�ֽ������ϵ�ҽ�����㣬�޷������������תסԺ������", vbInformation, gstrSysName
            Exit Function
        End If
  End If

'һ��ͨ���ݲ�����
'    ElseIf Not rsOneCard Is Nothing Then
'        If rsOneCard.RecordCount > 0 Then
'            If Not objICCard.ReturnSwap(rsOneCard!��λ�ʺ�, rsOneCard!ҽԺ����, "" & rsOneCard!�������, rsOneCard!���) Then
'                MsgBox "һ��ͨ�˷ѽ��׵���ʧ�ܣ��˷Ѳ���ʧ�ܣ�", vbExclamation, gstrSysName
'                exit function
'            End If
'        End If
'    '4.�����㴦��,�ݲ�����
'    If zlCallSquare_DelFree(lng����ID) = False Then
'        '�����������,�ڹ����оͻ�����
'                exit function
'    End If
    DelBalance = True
End Function

Private Sub vsBalance_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save 1131, vsBalance, Me.Caption, "����תסԺ�����б�", True
End Sub

Private Sub vsBalance_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save 1131, vsBalance, Me.Caption, "����תסԺ�����б�", True
End Sub

Private Function zlPrintInvoice(ByVal strNos As String, ByVal strDelDate As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��Ʊ����
    '��Σ�
    '   strNos ���δ�ӡԤ�����ݺţ���ʽ��A001,A002,A003,...
    '����:���˺�
    '����:2011-04-02 09:48:13
    '����:36984
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngShareUseID As Long, lng����ID As Long, strInvoice As String
    Dim blnInput As Boolean, blnValid As Boolean
    Dim strSQL As String
    Dim intInvoiceFormat As Integer
    
    '����ϸ����Ʊ��ʹ��
    On Error GoTo errHandle
    If gblnPrepayStrict Then
        lngShareUseID = zlDatabase.GetPara("����Ԥ��Ʊ������", glngSys, 1131, 0)
        '1.�ϸ����Ʊ��ʱ������ʵ�ʵ�Ʊ������,���¼������ID��Ʊ�ݺ�
        lng����ID = GetInvoiceGroupID(2, 1, lng����ID, lngShareUseID, strInvoice, "2")
        If lng����ID <= 0 Then
            Select Case lng����ID
                Case -1
                    MsgBox "Ԥ������[" & strNos & "]����Ҫ1��Ʊ��!" & vbCrLf & _
                        "��û���㹻�����ú͹��õ�Ʊ��,������һ�������ñ��ع���Ʊ�ݺ��ش�õ��ݣ�", vbInformation, gstrSysName
                Case -2
                    MsgBox "����[" & strNos & "]����Ҫ1��Ʊ��!" & vbCrLf & _
                        "��û���㹻�ĵĹ���Ʊ��,������һ�������ñ��ع���Ʊ�ݺ��ش�õ��ݣ�", vbInformation, gstrSysName
                Case -3
                    MsgBox "����[" & strNos & "]����Ҫ1��Ʊ��!" & vbCrLf & _
                        "Ʊ�ݺ�[" & strInvoice & "]���ڿ����������ε���ЧƱ�ݺŷ�Χ�ڣ�" & _
                        "������������Ч��Ʊ�ݺź��ش�õ��ݣ�", vbInformation, gstrSysName
                Case -4
                    MsgBox "����[" & strNos & "]����Ҫ1��Ʊ��!" & vbCrLf & _
                        "Ʊ�ݺ�[" & strInvoice & "]���ڵ���������û���㹻��Ʊ�ݣ�" & _
                        "���ȴ�ӡ����Ʊ��,���굱ǰ�������κ�,�ش�õ��ݣ�", vbInformation, gstrSysName
                Case Else
                    MsgBox "Ʊ��������Ϣ����ʧ�ܣ�������������ش򵥾�[" & strNos & "]", vbInformation, gstrSysName
            End Select
            Exit Function
        End If
        Do
            '����Ʊ�����ö�ȡ
            blnInput = False
            strInvoice = GetNextBill(lng����ID)
            If strInvoice = "" Then
                '�����;���ÿ���ĺ���,�������δ����,����һ�����ѳ�����Χ
                strInvoice = UCase(InputBox("�޷�����Ʊ�����������ȡ��Ҫʹ�õĿ�ʼƱ�ݺţ�" & _
                                vbCrLf & "�������뽫Ҫʹ�õĿ�ʼƱ�ݺ��룺", gstrSysName, _
                                "", Me.Left + 1500, Me.Top + 1500))
                blnInput = True
            Else
                strInvoice = UCase(InputBox("��ȷ���ش�ʹ�õĿ�ʼƱ�ݺ��룺", gstrSysName, _
                                strInvoice, Me.Left + 1500, Me.Top + 1500))
                blnInput = True
            End If
            
            '�û�ȡ������,����ӡ
            If strInvoice = "" Then Exit Function
            '���������Ч��
            If blnInput Then
                If GetInvoiceGroupID(2, 1, lng����ID, lngShareUseID, strInvoice, "2") = -3 Then
                    MsgBox "�������Ʊ�ݺ��벻�ڵ�ǰ�������ε���Ч���÷�Χ��,���������룡", vbInformation, gstrSysName
                Else
                    blnValid = True
                End If
            Else
                blnValid = True
            End If
        Loop While Not blnValid
    Else
        '�п����ǵ�һ��ʹ��
         Do
             blnInput = False
             '���ϸ����ʱֱ�Ӵӱ��ض�ȡ
             strInvoice = UCase(zlDatabase.GetPara("��ǰԤ��Ʊ�ݺ�", glngSys, 1131, ""))
             If strInvoice = "" Then
                 strInvoice = UCase(InputBox("û���ҵ����õ����Ʊ�ݺ��룬�޷�ȷ����Ҫʹ�õĿ�ʼƱ�ݺš�" & _
                                 vbCrLf & "�����뽫Ҫʹ�õĿ�ʼƱ�ݺ��룺", gstrSysName, _
                                 "", Me.Left + 1500, Me.Top + 1500))
                 blnInput = True
             Else
                 strInvoice = zlCommFun.IncStr(strInvoice)
                 strInvoice = UCase(InputBox("��ȷ���ش�ʹ�õĿ�ʼƱ�ݺ��룺", gstrSysName, _
                                 strInvoice, Me.Left + 1500, Me.Top + 1500))
                 blnInput = True
             End If
                 
             '�û�ȡ������,�����ӡ
             If strInvoice = "" Then
                 If MsgBox("��ȷ��������Ʊ�ݺż�����ӡ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                 blnValid = True
             Else
                 '���������Ч��
                 If blnInput Then
                     If zlCommFun.ActualLen(strInvoice) <> gbytPrepayLen Then
                         MsgBox "�����Ʊ�ݺ��볤��Ӧ��Ϊ " & gbytPrepayLen & " λ��", vbInformation, gstrSysName
                     Else
                         blnValid = True
                     End If
                 Else
                     blnValid = True
                 End If
             End If
         Loop While Not blnValid
    End If
    
    'ִ�����ݴ���
    'Zl_����Ԥ����¼_Reprint
    strSQL = "Zl_����Ԥ����¼_Reprint("
    '  ���ݺ�_In Varchar2,
    strSQL = strSQL & "'" & strNos & "',"
    '  Ʊ�ݺ�_In Ʊ��ʹ����ϸ.����%Type,
    strSQL = strSQL & "'" & strInvoice & "',"
    '  ����id_In Ʊ��ʹ����ϸ.����id%Type,
    strSQL = strSQL & "" & IIf(lng����ID = 0, "NULL", lng����ID) & ","
    '  ʹ����_In Ʊ��ʹ����ϸ.ʹ����%Type
    strSQL = strSQL & "'" & UserInfo.���� & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    '���Ʊ��
    intInvoiceFormat = Val(zlDatabase.GetPara(284, glngSys, , "0"))
    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103", Me, _
        "NO=" & strNos, "�տ�ʱ��=" & Format(strDelDate, "yyyy-mm-dd HH:MM:SS"), _
        "����ID=" & mlngPatient, IIf(intInvoiceFormat = 0, "", "ReportFormat=" & intInvoiceFormat), 2)
    
    '���±���Ʊ��
    If Not gblnPrepayStrict Then
        zlDatabase.SetPara "��ǰԤ��Ʊ�ݺ�", strInvoice, glngSys, 1131
    End If
    zlPrintInvoice = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
 
Private Sub zlCreateObject()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������¼�����
    '����: �����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2012-08-28 16:16:00
    '˵��:
    '����:54894
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '������������
    Err = 0: On Error Resume Next
    If mobjICCard Is Nothing Then
        Set mobjICCard = New clsICCard
        Call mobjICCard.SetParent(Me.hWnd)
         Set mobjICCard.gcnOracle = gcnOracle
    End If
    
    If mobjIDCard Is Nothing Then
        Set mobjIDCard = New clsIDCard
        Call mobjIDCard.SetParent(Me.hWnd)
    End If
    
End Sub
Private Sub zlCloseObject()
    '�ر���ض���
    Err = 0: On Error Resume Next
    If Not mobjICCard Is Nothing Then
        Call mobjICCard.SetEnabled(False)
    End If
    If Not mobjIDCard Is Nothing Then
        Call mobjIDCard.SetEnabled(False)
    End If
    Set mobjIDCard = Nothing
    Set mobjICCard = Nothing
End Sub
