VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#9.0#0"; "zlIDKind.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMultiBills 
   AutoRedraw      =   -1  'True
   Caption         =   "�൥���˷�"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11265
   Icon            =   "frmMultiBills.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   11265
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picInvoice 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   11265
      TabIndex        =   39
      Top             =   960
      Width           =   11265
      Begin VB.Frame fraSelectDownSplit 
         Height          =   30
         Left            =   -15
         TabIndex        =   41
         Top             =   900
         Width           =   11535
      End
      Begin VB.Frame fraSelectTopSplit 
         Height          =   45
         Left            =   -30
         TabIndex        =   40
         Top             =   0
         Width           =   11385
      End
      Begin VSFlex8Ctl.VSFlexGrid vsInvoice 
         Height          =   375
         Left            =   300
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   180
         Width           =   9960
         _cx             =   17568
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
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483640
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483633
         GridColor       =   12632256
         GridColorFixed  =   -2147483641
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   3
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   14
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
         FormatString    =   $"frmMultiBills.frx":058A
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
   Begin VB.PictureBox pic�˷�ժҪ 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   11265
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   5565
      Width           =   11265
      Begin VB.TextBox txt�˷�ժҪ 
         Height          =   360
         Left            =   1035
         MaxLength       =   100
         TabIndex        =   6
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
         TabIndex        =   5
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
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   7404
      Width           =   11265
      Begin VB.TextBox txtYB 
         Height          =   300
         Left            =   945
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   180
         Visible         =   0   'False
         Width           =   450
      End
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
         TabIndex        =   33
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   25
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
         TabIndex        =   24
         ToolTipText     =   "�ȼ���Ctrl+A"
         Top             =   135
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
      TabIndex        =   18
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
            Picture         =   "frmMultiBills.frx":0648
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
      TabIndex        =   4
      Top             =   1935
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
      FormatString    =   $"frmMultiBills.frx":0EDC
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
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   6420
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
         Left            =   5850
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   75
         Width           =   1275
      End
      Begin VB.Frame fra�˿� 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   6705
         TabIndex        =   22
         Top             =   75
         Width           =   4515
         Begin VB.ComboBox cbo�˿ʽ 
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   1005
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   0
            Width           =   1620
         End
         Begin VB.TextBox txt�˿��� 
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
            ForeColor       =   &H000000C0&
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   3195
            Locked          =   -1  'True
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   0
            Width           =   1275
         End
         Begin VB.Label lbl�˿ʽ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�����˿�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   240
            Left            =   0
            TabIndex        =   12
            Top             =   75
            Width           =   960
         End
         Begin VB.Label lbl�˿��� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   240
            Left            =   2685
            TabIndex        =   14
            Top             =   60
            Width           =   480
         End
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
         TabIndex        =   11
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
         TabIndex        =   9
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
         Left            =   4875
         TabIndex        =   36
         Top             =   135
         Width           =   960
      End
      Begin VB.Label lblAllTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
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
         TabIndex        =   10
         Top             =   135
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
         TabIndex        =   8
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
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   0
      Width           =   11265
      Begin VB.PictureBox picPatiBack 
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   540
         ScaleHeight     =   360
         ScaleWidth      =   2115
         TabIndex        =   37
         Top             =   525
         Width           =   2115
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
            TabIndex        =   3
            ToolTipText     =   "��λ:F6,����:-����ID,*�����,+סԺ��,.�Һŵ���,����:*2536��ʾ������Ų���"
            Top             =   0
            Width           =   1450
         End
         Begin zlIDKind.IDKindNew IDKind 
            Height          =   360
            Left            =   0
            TabIndex        =   38
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
      Begin VB.TextBox txtPatientPrint 
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
         Left            =   9315
         MaxLength       =   64
         TabIndex        =   30
         ToolTipText     =   "�ȼ�:F11"
         Top             =   540
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.OptionButton optNO 
         Caption         =   "Ʊ�ݺ�"
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
         Index           =   1
         Left            =   8280
         TabIndex        =   1
         Top             =   165
         Width           =   1035
      End
      Begin VB.OptionButton optNO 
         Caption         =   "���ݺ�"
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
         Index           =   0
         Left            =   7245
         TabIndex        =   0
         Top             =   165
         Value           =   -1  'True
         Width           =   1035
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
         TabIndex        =   28
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
            TabIndex        =   29
            Top             =   0
            Width           =   405
         End
      End
      Begin VB.Frame fraInfo_1 
         Height          =   120
         Left            =   -120
         TabIndex        =   27
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
         TabIndex        =   2
         Top             =   120
         Width           =   1800
      End
      Begin VB.Label lblPatiName 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   8760
         TabIndex        =   31
         Top             =   600
         Visible         =   0   'False
         Width           =   480
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
         Left            =   75
         TabIndex        =   26
         ToolTipText     =   "���:F6"
         Top             =   45
         Width           =   1875
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
         TabIndex        =   20
         Top             =   585
         Width           =   720
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsBalance 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   6045
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
      FormatString    =   $"frmMultiBills.frx":0F56
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
Attribute VB_Name = "frmMultiBills"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytMode As Integer '0-���ŵ��ݲ鿴,1-���ŵ����˷�,2-���쳣���˷ѵ����������˷�
Private mstrNo As String 'Ҫ�鿴���˷ѵĶ��ŵ����е�ĳ��NO,�˷�ʱ����û��
Private mblnNOMoved As Boolean '�����ĵ����Ƿ��ں����ݱ���
Private mstrPrivs As String
Private mlng����ID As Long
Private mblnOneCard As Boolean
Private mblnSingleBlance As Boolean '���ֽ��㷽ʽ
Private mrsALL As ADODB.Recordset  '���е��ݵ���ϸ��¼
Private mstrDelTime As String '�鿴�˷ѵ��ݵĵǼ�ʱ��(yyyy-MM-dd HH:mm:ss) 'ֻ�в鿴�˷ѵ���ʱ�Ŵ���ʱ��,��������������
Private mstrNOs As String 'ʵ�ʶ��������˷ѵĵ��ݺ�
Private mstrNOsOverFlow As String '����������޵ĵ��ݺ�
Private mrsBalance As ADODB.Recordset '��¼ÿ�ŵ��ݵĽ������
Private mcolError As Collection '��¼ÿ�ŵ��ݵ������
Private mstrDelNOs As String '�Ѿ�����ĵ��ݻ�ִ�в����˵ĵ���
Private mstr�����ʻ� As String   'ҽ�������ʻ�������
Private mintInsure As Integer   'ҽ�����ݵ�����
Private mblnYB�������� As Boolean 'ҽ���Ƿ�֧�������������
Private mint�˷ѻص���ӡ As Integer '�˷ѻص���ӡ��ʽ 0-����ӡ,1-�Զ���ӡ,2-ѡ���Ƿ��ӡ
Private mblnOK As Boolean
Private mblnPrintView As Boolean    '��ӡǰ�鿴����
Private mintReturnMode As Integer   '�����˷�ʱ,ȫ�˽��ý��㷽ʽʱ�ָ���ʼ�Ľ��㷽ʽ
Private mrs�շѶ��� As ADODB.Recordset '�շѶ��� :����:33634
Private Const mlngModule = 1121
Private mstrNOsPatiDel As String    '��¼�����˷ѵĵ���
Private Type TYPE_MedicarePAR
    ��������ҽ����Ŀ As Boolean
    ������봫����ϸ As Boolean
    ҽ���ӿڴ�ӡƱ�� As Boolean
    �൥��һ�ν��� As Boolean
    �ֱҴ��� As Boolean
    ʵʱ��� As Boolean
    �˷Ѻ��ӡ�ص� As Boolean
    �൥�ݵ�һ�ν��� As Boolean
    �൥���շѱ���ȫ�� As Boolean
    
End Type
Private mstr�ֽ���㷽ʽ As String
Private MCPAR As TYPE_MedicarePAR
Private mobjSquare As Object
Private mlngShareUseID As Long '������������ID
Private mstrUseType As String 'ʹ�����
Private mintInvoiceFormat As Integer  '��ӡ�ķ�Ʊ��ʽ,��Ʊ��ʽ���
Private mintOldInvoiceFormat As Integer '�ɷ�Ʊ��ʽ
Private mintInvoicePrint As Integer '0-����ӡ;1-�Զ���ӡ;2-��ʾ��ӡ
Private mblnNotClick As Boolean
'�������ѿ��Ĵ������
Private Type Ty_SquareCard
    blnExistsObjects As Boolean '��װ�����ѿ���
    rsSquare As ADODB.Recordset
    dblˢ���ܶ� As Double
    bln������ As Boolean '��ǰ��ȡ�ĵ����ǿ�����
End Type
Private mtySquareCard As Ty_SquareCard
Private mlng����ID As Long
Private Type Ty_Pati
    ����ID As Long
    ���� As String
    �Ա� As String
    ���� As String
End Type
Private mtyPati As Ty_Pati
Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private mrsInfo As ADODB.Recordset
Private Enum EM_BillDelType
        EM_����ȫ�� = 0
        EM_����ȫ�� = 1
        EM_���Ų����� = 2
End Enum
Private mBillDelType As EM_BillDelType
Private mblnHaveExcuteData As Boolean '�Ƿ�ҽ���Ƽ��д�������:60735
Private Type tyBillType
    bln�൥�� As Boolean
    bln���ֽ��㷽ʽ As Boolean
    bln���Ų����˷� As Boolean  '���ڵ��Ų�����
    bln���Ų����˷� As Boolean  '���ڶ��Ų�����
    bln���ڿ����� As Boolean '���ڿ�����
    bln����������� As Boolean  '����������˷�
    strNos As String '�����շѵ���
    str���㷽ʽ As String '��ǰ���㷽ʽ:����ʱ,�ö��ŷָ�
    blnSingleBalance As Boolean
    bln����ҽ�ƿ����� As Boolean
    bln������ȫ�� As Boolean
End Type
Private mCurBillType As tyBillType  '��ǰ��������
Private mrsDelInvoice As ADODB.Recordset
Private mobjDrugPacker  As Object ' �Զ���ҩ��(���·�ҩ����)
Private mblnDrugPacker As Boolean
Private mblnFromInNewDel As Boolean ' �Ƿ��Ǵ����˷Ѵ��ڽ�����

Public Function ShowMe(frmParent As Object, _
    ByVal bytMode As Byte, ByVal strPrivs As String, _
    ByVal strNo As String, ByVal strTime As String, _
    Optional blnPrintView As Boolean, _
    Optional lng����ID As Long = 0, _
    Optional blnOneCard As Boolean = False, _
    Optional blnNOMoved As Boolean = False, _
    Optional blnFromInNewDel As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�൥���˷����
    '���:bytMode-0-���ŵ��ݲ鿴,1-���ŵ����˷�,2-���쳣���˷ѵ����������˷�
    '       strPrivs-Ȩ�޴�
    '       blnNOMoved-�Ƿ�ת�������ݱ�
    '       blnFromInNewDel-�Ƿ��Ǵ����˷Ѵ��ڽ�����
    '����:
    '����:�˷ѳɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-09-04 10:16:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mblnNOMoved = blnNOMoved: mstrPrivs = strPrivs
    mlng����ID = lng����ID: mblnOneCard = blnOneCard
    mbytMode = bytMode: mstrNo = strNo: mblnFromInNewDel = blnFromInNewDel
    mstrDelTime = strTime              'ֻ�в鿴�˷ѵ���ʱ�Ŵ���ʱ��,��������������
    mblnPrintView = blnPrintView
    mblnOK = False
    On Error Resume Next
    If frmParent Is Nothing Then
        'ҽ�����Ե���
        Me.Show 0
    Else
        Me.Show 1, frmParent
    End If
    On Error GoTo 0
    ShowMe = mblnOK
End Function
Private Sub cbo�˿ʽ_Click()
    If mblnNotClick Then Exit Sub
    Call ReCalcDelMoney
End Sub

Private Sub cmdBillSel_Click()
    Dim i As Long
    With vsBill
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("���ݺ�")) <> "" And _
               .TextMatrix(i, .ColIndex("���ݺ�")) = .TextMatrix(.Row, .ColIndex("���ݺ�")) And InStr(1, mstrNOsOverFlow, vsBill.TextMatrix(i, .ColIndex("���ݺ�"))) <= 0 Then
                .TextMatrix(i, .ColIndex("ѡ��")) = -1
            End If
        Next
    End With
    Call LoadDelBalanceInfor
    Call ReCalcDelMoney
    Call FromNoSelectInvoice
    Call ShowAndHideDelBillRow
End Sub

Private Sub cmdCancel_Click()
    If mblnPrintView Then
        If txtPatientPrint.Visible Then
            If txtPatientPrint.Text = "" Then
                MsgBox "����Ϊ��,������������", vbInformation, gstrSysName
                If txtPatientPrint.Enabled Then txtPatientPrint.SetFocus
                Exit Sub
            End If
            
            If zlCommFun.ActualLen(txtPatientPrint.Text) > txtPatientPrint.MaxLength Then
                MsgBox "�����������������ֻ�������� " & txtPatientPrint.MaxLength & " ���ַ��� " & txtPatientPrint.MaxLength \ 2 & " �����֡�", vbInformation, gstrSysName
                If txtPatientPrint.Enabled Then txtPatientPrint.SetFocus
                Exit Sub
            End If
            
            If txtPatientPrint.Text <> txtPatientPrint.Tag Then
                
                Call ExecuteModifyPatiName
            End If
        End If
        mblnOK = True
    End If
    If mstrNOs <> "" And txtNO.Visible Then
        Call ClearFace
        txtNO.SetFocus
    Else
        Unload Me
    End If
End Sub

Private Function SetNOBill(ByVal strNo As String, ByVal blnSel As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ȫѡ��ȫ�嵥��
    '���:strNO-ָ����NO
    '        blnSel:true��ʾȫѡ,����ȫ��
    '����:
    '����:
    '����:���˺�
    '����:2011-01-24 10:47:58
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
    SetNOBill = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Sub ExecuteModifyPatiName()
    Dim arrSQL As Variant, arrNo As Variant
    Dim i As Long, blnTrans As Boolean
    
    On Error GoTo errH
    arrNo = Split(mstrNOs, ",")
    arrSQL = Array()
    
    For i = 0 To UBound(arrNo)
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_���˷��ü�¼_Update('" & arrNo(i) & "',1,null,null,'" & txtPatientPrint.Text & "')"
    Next
    
    gcnOracle.BeginTrans: blnTrans = True
        For i = 0 To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
        Next
    gcnOracle.CommitTrans: blnTrans = False
    
    Exit Sub
errH:
    If Err.Number <> 0 Then
        If blnTrans Then gcnOracle.RollbackTrans
        If ErrCenter() = 1 Then Resume
    End If
     
    If Err.Number <> 0 Then Call SaveErrLog
End Sub

Private Sub cmdClear_Click()
    Dim i As Long
    With vsBill
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, .ColIndex("ѡ��")) = 0
        Next
    End With
    Call LoadDelBalanceInfor
    Call ReCalcDelMoney
  '62492
    If vsInvoice.Visible Then
        If vsInvoice.Rows - 1 >= 1 And vsInvoice.COLS - 1 >= 1 Then
            vsInvoice.Cell(flexcpChecked, 0, 1, vsInvoice.Rows - 1, vsInvoice.COLS - 1) = 2
        End If
    End If
    Call ShowAndHideDelBillRow
End Sub
Private Function GetErrBillPartDelFee() As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ȷ���쳣�����Ƿ񲿷��˷�
    '����:bln���Ų�����
    '����:�����˷�����(0-ȫ��;1-���Ų�����;2-�൥�ݰ����ݲ�����)
    '����:���˺�
    '����:2011-09-04 14:31:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim varData As Variant, i As Long, strNo As String
    On Error GoTo errHandle
    strSQL = "" & _
    "   Select /*+ rule */ A.NO,Max(decode(A.��¼״̬,2,1,0)) as ����, " & _
    "           nvl(sum(A.ʵ�ս��),0)-nvl(sum(A.���ʽ��),0) as δ���� " & _
    "   From ������ü�¼ A,Table(f_str2List([1])) J" & _
    "   Where A.NO=J.Column_Value and A.��¼����=1  and A.ִ��״̬<>9 " & _
    "   Having  nvl(sum(A.ʵ�ս��),0)-nvl(sum(A.���ʽ��),0) <>0 " & _
    "   Group by A.NO"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrNOs)
    
    If rsTemp.RecordCount = 0 Then GetErrBillPartDelFee = 0: Exit Function
    '����˷ѵ����Ƿ����е�������
    With rsTemp
        varData = Split(mstrNo, ",")
        Do While Not .EOF
            strNo = Nvl(rsTemp!NO)
            If Val(Nvl(!����)) = 1 Then GetErrBillPartDelFee = 1: Exit Function
            For i = 0 To UBound(varData)
                If varData(i) = strNo Then strNo = "HAVE": Exit For
            Next
            If strNo <> "HAVE" Then GetErrBillPartDelFee = 2
            .MoveNext
        Loop
    End With
    GetErrBillPartDelFee = 0
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function ErrBillReDelFee() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���쳣���������˷�
    '����:�˷ѳɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-09-04 10:42:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln���� As Boolean, str�˽��㷽ʽ As Boolean
    Dim blnYbComit As Boolean, blnCommited As Boolean, blnOneCardComit As Boolean
    Dim arrNo As Variant, i As Long, intҽ�� As Integer
    Dim lngPages As Long, lngPage, cllYB As Collection
    Dim rsԭ���� As ADODB.Recordset, strҽ������ As String
    Dim rsҽ�� As ADODB.Recordset, strInvoices As String
    Dim str����ID As String, strNo As String, lng����ID As Long
    Dim blnPrint As Boolean, strAllNOs As String
    Dim varTemp As Variant, blnTrans As Boolean
    Dim strReclaimInvoice  As String, intInvoiceFormat As Integer
    Dim strReturn As String, strReturnRecipt As String '�˷Ѵ�����Ϣ����ʽ��NO,ҩ��ID|NO,ҩ��ID|��
    Dim rsҩƷ��¼ As ADODB.Recordset
    
    On Error GoTo errHandle
    Dim strSQL As String, blnAll�����˷� As Boolean, bln���Ų����� As Boolean
    '�������
    If zlIsCheckExistErrBill(0, False, mstrNOs) = False Then
        MsgBox "��ǰ�쳣�����ѱ������㲻�ܼ�����", vbInformation, gstrSysName
        Exit Function
    End If
    If zlCheckOtherSessionDoing(0, mstrNOs) Then
        MsgBox "��ǰ�������������շѴ����н��д����㲻�ܼ�����", vbInformation, gstrSysName
        Exit Function
    End If
    '0-ȫ��;1-���Ų�����;2-�൥�ݰ����ݲ�����
    bln���Ų����� = False: blnAll�����˷� = False
    Select Case GetErrBillPartDelFee
    Case 1
        blnAll�����˷� = True: bln���Ų����� = True
    Case 2
        bln���Ų����� = False: blnAll�����˷� = True
    Case Else
    End Select
    
    If blnAll�����˷� Then
        If InStr(mstrPrivs, ";�����˷�;") = 0 Then
            MsgBox "��û��Ȩ��ִ�в����˷Ѳ�����", vbInformation, gstrSysName
            vsBill.SetFocus: Exit Function
        End If
        '���˺� ����:27352 ����:2010-01-13 10:26:08
        If InStr(1, mstrPrivs, ";�˷Ѻ��շ�Ʊ;") > 0 Then
            If frmReInvoice.ShowMe(Me, mstrNo, Val(txtAllTotal.Text), Val(txt�˿���.Text), strInvoices) = False Then
                vsBill.SetFocus: Exit Function
            End If
        End If
    End If
      
    With mrsBalance
        .Filter = 0
        str����ID = ""
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If InStr(1, str����ID & ",", "," & Val(Nvl(!����ID)) & ",") = 0 Then
                    str����ID = str����ID & "," & Val(Nvl(!����ID))
            End If
            .MoveNext
        Loop
        If str����ID <> "" Then str����ID = Mid(str����ID, 2)
        If str����ID = "" Then
            MsgBox "δ�ҵ���������,����!", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
    End With
    varTemp = Split(str����ID, ",")
    '����:43347
    For i = 0 To UBound(varTemp)
        'һ��ͨ���
        If Not CheckOnCardValied(bln���Ų�����, Val(varTemp(i))) Then Exit Function
        '�������׼��
        If Not CheckThreeSwapValied(bln���Ų�����, Val(varTemp(i)), InStr(1, mstrNOs, ",") > 0, True) Then Exit Function
    Next
    
    strSQL = "" & _
    "   Select /*+ rule */ distinct  A.����ID,A.No  " & _
    "   From  ������ü�¼ A,Table(f_str2List([1])) J" & _
    "   Where A.No=J.Column_Value and A.��¼����=1 And  A.��¼״̬ in (1,3)"
    Set rsԭ���� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrNOs)
 
    
    strSQL = "Select /*+ Rule*/ Distinct a.No, a.����id, Decode(Nvl(b.����, 0), 0, 0, 1) As ҽ������" & vbNewLine & _
            " From ������ü�¼ A," & vbNewLine & _
            "      (Select Distinct j.Column_Value As ����id, m.����" & vbNewLine & _
            "        From Table(f_Num2list([1])) J, ���ս����¼ M" & vbNewLine & _
            "        Where j.Column_Value = m.��¼id(+) And m.����(+) = 1) B" & vbNewLine & _
            " Where a.����id = b.����id"
    Set rsҽ�� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str����ID)
    arrNo = Split(mstrNOs, ",")
'    bln���� = False
'    If cbo�˿ʽ.ListIndex >= 0 Then
'        bln���� = cbo�˿ʽ.ItemData(cbo�˿ʽ.ListIndex) = 1
'        str�˽��㷽ʽ = zlStr.NeedName(cbo�˿ʽ.Text)
'    End If

    
    If gblnBillPrint Then
        If gobjBillPrint.zlEraseBill(strAllNOs, 0) = False Then Exit Function
    End If
    blnTrans = False
    blnCommited = False: blnYbComit = False
    '--------------------------------------------------------------------------
    'ҽ���˷�:
    If mintInsure <> 0 Then
        '�൥���˷��Ƿ���һ�������е�,�϶��˷ѳɹ�
        ' ������ҽ������ʱ,��ԭ���˵�,δ���ӿ�.���Ҳ�����ܳ���
        If Not (MCPAR.�൥��һ�ν��� Or MCPAR.�൥�ݵ�һ�ν��� Or Not mblnYB��������) Then
            '-------------------------------------------------------------------------------------------------------
            '���˺�:ҽ����strAdvancey����:�����˷�������|��ǰ�˷ѵڼ���:27231
            Set cllYB = New Collection
            lngPage = 0: lngPages = 0
            For i = 0 To UBound(arrNo)
                strNo = arrNo(i)
                lngPage = UBound(arrNo) + 1 - i
                lngPages = lngPages + 1
                rsԭ����.Filter = "NO='" & arrNo(i) & "'"
                If rsԭ����.EOF Then
                    MsgBox "δ�ҵ����ݺ�Ϊ" & arrNo(i) & "��ԭʼ��������,����!", vbOKOnly + vbInformation, gstrSysName
                    Exit Function
                End If
                mrsBalance.Filter = "NO='" & arrNo(i) & "' And ����=2"
                If mrsBalance.RecordCount <> 0 Then mrsBalance.MoveFirst
                With mrsBalance
                    strҽ������ = ""
                    Do While Not .EOF
                        strҽ������ = strҽ������ & "," & Nvl(!���㷽ʽ)
                        .MoveNext
                    Loop
                    If strҽ������ <> "" Then strҽ������ = Mid(strҽ������, 2)
                End With
                rsҽ��.Filter = "NO='" & arrNo(i) & "' and ҽ������=1"
                intҽ�� = IIf(rsҽ��.RecordCount <> 0, 1, 0)
                lng����ID = Val(Nvl(rsԭ����!����ID))
                cllYB.Add Array(lngPage, lng����ID, strҽ������, intҽ��), "_" & strNo
            Next
            'ҽ��
            gcnOracle.BeginTrans: blnTrans = True
             For i = 0 To UBound(arrNo)
                strNo = arrNo(i)
                If Val(cllYB("_" & strNo)(3)) = 0 Then
                   'strAdace
                    lngPage = Val(cllYB("_" & strNo)(0)):  lng����ID = Val(cllYB("_" & strNo)(1))
                    If Not DelInsureOneBill(strҽ������, True, lng����ID, i + 1, lngPages, blnCommited) Then
                        If blnCommited = False Then gcnOracle.RollbackTrans: blnTrans = False
                        Exit Function
                    End If
                    If blnCommited = False Then gcnOracle.CommitTrans
                    gcnOracle.BeginTrans: blnTrans = True
                End If
            Next
        End If
    End If
    
    If blnTrans = False Then gcnOracle.BeginTrans: blnTrans = True
    '------------------------------------------------------------------------------------------
    '��һ��ͨ
    blnCommited = False
    If Not DelOneCardPay(arrNo, blnCommited) Then
        If blnCommited = False Then gcnOracle.RollbackTrans: blnTrans = False
        Exit Function
    End If
    If blnCommited Then gcnOracle.BeginTrans: blnTrans = True
    '------------------------------------------------------------------------------------------
    '��һ��ͨ�ȵ���������
    blnCommited = False
    If Not DelThreeSwapFee(arrNo, blnCommited) Then
        If blnCommited = False Then gcnOracle.RollbackTrans: blnTrans = False
        Exit Function
    End If
    If blnCommited Then gcnOracle.BeginTrans: blnTrans = True
 
    '------------------------------------------------------------------------------------------
    '����˷�
    blnCommited = False
    If OverFeeDel(str����ID, mtyPati.����ID, blnCommited) = False Then
        If blnCommited = False Then gcnOracle.RollbackTrans
        Exit Function
    End If
    If blnCommited = False Then gcnOracle.CommitTrans
    
    '81190,Ƚ����,�˷�ҵ����ҩ���ϴ��˷���Ϣ
    On Error Resume Next
    If mblnDrugPacker Then
        strSQL = "Select NO, ִ�в���id" & _
            "   From ������ü�¼" & _
            "   Where ����id In (Select Column_Value From Table(f_Str2list([1]))) And �շ���� In ('5', '6', '7')"
        Set rsҩƷ��¼ = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str����ID)
        
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
    Err.Clear: On Error GoTo errHandle
    
   '�����˷�ʱ�ջز��ش�,�������Ų����˺��˶����е�ĳ����
    If blnAll�����˷� Then
        'If InStr(1, mstrPrivs, "�˷Ѻ��շ�Ʊ") > 0 Then strInvoices = frmReInvoice.ShowMe(Me, strNO, Val(txtAllTotal.Text), Val(txt�˿���.Text))
        If strInvoices = "" Then 'a.�ջز����´�ӡ�����վ�
            blnPrint = True
            If mintInvoicePrint = 0 Then
                blnPrint = False
            Else
                If mintInvoicePrint = 2 Then
                    If MsgBox("�Ƿ��ӡƱ�ݣ�", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                        blnPrint = False
                    End If
                End If
            End If
            If blnPrint Then
                intInvoiceFormat = IIf(strReclaimInvoice = "" And gTy_Module_Para.bytƱ�ݷ������ <> 0, mintOldInvoiceFormat, mintInvoiceFormat)
                Dim strPriceGrade As String
                If gintPriceGradeStartType >= 2 Then
                    strPriceGrade = GetPriceGradeFromNos(strAllNOs)
                Else
                    strPriceGrade = gstr��ͨ�۸�ȼ�
                End If
                Call RePrintCharge(1, strAllNOs, Me, mlng����ID, strReclaimInvoice, True, CDate(mstrDelTime), _
                     intInvoiceFormat, , , mlngShareUseID, mstrUseType, , strPriceGrade)
            End If
        Else
            If strInvoices = "�޿���Ʊ��" Then
                'b.�շѻ���һ����ʱû�д�ӡƱ��
            Else
                'c.ֻ�ջ�Ʊ��
                strSQL = "zl_�����շѼ�¼_RePrint('" & strNo & "',Null,0,'" & UserInfo.���� & "'," & _
                        "To_Date('" & mstrDelTime & "','YYYY-MM-DD HH24:MI:SS'),1,0,'" & strInvoices & "')"
                Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            End If
        End If
        
        '��ӡ�����嵥
        If InStr(mstrPrivs, ";��ӡ�嵥;") > 0 Then
            If gint�շ��嵥 = 1 Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_3", Me, "NO=" & strAllNOs, "ҩƷ��λ=" & IIf(gblnҩ����λ, 1, 0), 2)
            ElseIf gint�շ��嵥 = 2 Then
                If MsgBox("Ҫ��ӡ�շ��嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_3", Me, "NO=" & strAllNOs, "ҩƷ��λ=" & IIf(gblnҩ����λ, 1, 0), 2)
                End If
            End If
        End If
    Else
         '˰�ز���ȫ��ʱ�ջش���(ȫ��ʱ��zl_�����շѼ�¼_DELETE�����ջ�Ʊ��)
        If Not gobjTax Is Nothing And gblnTax Then
            gstrTax = gobjTax.zlTaxOutErase(gcnOracle, strAllNOs)
            If gstrTax <> "" Then MsgBox gstrTax, vbExclamation, gstrSysName
        End If
    End If
    If mintInsure <> 0 And MCPAR.�˷Ѻ��ӡ�ص� And InStr(1, mstrPrivs, ";ҽ���˷ѻص�;") > 0 Then
        '����:35248
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_4", Me, "NO=" & mstrNOs, 2)
    End If
    If mint�˷ѻص���ӡ = 1 Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_5", Me, "NO=" & mstrNOs, 2)
    ElseIf mint�˷ѻص���ӡ = 2 Then
        If MsgBox("�Ƿ��ӡ�˷ѻص���", vbYesNo + vbQuestion, gstrSysName) = vbYes Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_5", Me, "NO=" & mstrNOs, 2)
        End If
    End If
    ErrBillReDelFee = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub cmdOK_Click()
    If mbytMode = 2 Then
        '�쳣���������˷�
        If ErrBillReDelFee = False Then Exit Sub
        mblnOK = True
        Unload Me: Exit Sub
    End If
    If ExecDelete Then
        mblnOK = True
        Call ClearFace(True, False)
        If txtNO.Visible Then
            txtNO.SetFocus
        Else
            Unload Me
            Exit Sub
        End If
    End If
 
End Sub

Private Sub cmdSelAll_Click()
    Dim i As Long
    With vsBill
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("���ݺ�")) <> "" And InStr(1, mstrNOsOverFlow, vsBill.TextMatrix(i, .ColIndex("���ݺ�"))) <= 0 Then
                .TextMatrix(i, .ColIndex("ѡ��")) = -1
            End If
        Next
    End With
    If mbytMode <> 0 Then
        Call LoadDelBalanceInfor
    End If
    Call ReCalcDelMoney
    Call FromNoSelectInvoice
    Call ShowAndHideDelBillRow
End Sub

Private Sub Form_Activate()
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
    ElseIf KeyCode = vbKeyEscape Then
        If mblnPrintView Then
            Unload Me
        Else
            If cmdCancel.Visible Then Call cmdCancel_Click
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("',|~:��;��?��" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub
Private Sub ClearVar()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����ر���
    '����:���˺�
    '����:2012-09-17 13:23:35
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With mCurBillType
        .bln���Ų����˷� = False
        .bln���ֽ��㷽ʽ = False
        .bln�൥�� = False
        .bln���Ų����˷� = False
        .strNos = ""
        .str���㷽ʽ = ""
        .blnSingleBalance = False
        .bln����ҽ�ƿ����� = False
        .bln������ȫ�� = False
    End With
End Sub

Private Sub Form_Load()

    Call SetpicInvoiceVisible
    Call InitBillHead
    Call RestoreWinState(Me, App.ProductName)
    If Val(zlDatabase.GetPara("�˷Ѻ�������ģʽ", glngSys, 1121, 0)) = 0 Then
        optNO(0).Value = True
    Else
        optNO(1).Value = True
    End If
    
    Call NewCardObject
    Call ClearFace
    
     lblTitle.Caption = gstrUnitName & "�����շѵ�"
    mint�˷ѻص���ӡ = Val(zlDatabase.GetPara("�˷ѻص���ӡ��ʽ", glngSys, mlngModule, "0"))
    If mbytMode = 0 Then '�鿴����
        Caption = "�鿴���ŵ���"
        fra�˿�.Visible = False
        vsBill.ColHidden(0) = True
        cmdSelAll.Visible = False
        cmdClear.Visible = False
        cmdOK.Visible = False
        cmdBillSel.Visible = False
        If mblnPrintView Then cmdCancel.Caption = "ȷ��(&X)"
        pic��.Visible = mstrDelTime <> ""
        lbl�˿�ϼ�.Visible = False: txt�˿�ϼ�.Visible = False
    ElseIf mbytMode = 2 Then
        '�쳣�����˷�
        Caption = "�쳣�˷ѵ������˷�"
        Call initCardSquareData
        vsBill.ColHidden(0) = True
        cmdSelAll.Visible = False
        cmdClear.Visible = False
        cmdBillSel.Visible = False
        fra�˿�.Visible = False
        cmdOK.Visible = True
        pic��.Visible = mstrDelTime <> ""
        vsBill.Editable = flexEDNone
    Else
        Caption = "���ŵ����˷�"
        Call initCardSquareData
    End If
    
    If mstrNo <> "" Then 'ָ���˵���
        txtNO.Visible = False
        optNO(0).Visible = False
        optNO(1).Visible = False
        picPatiBack.Visible = False
        If Not ReadBills(mstrNo) Then Unload Me: Exit Sub
    End If

    '81190,Ƚ����,�˷�ҵ����ҩ���ϴ��˷���Ϣ
    mblnDrugPacker = False
    If mobjDrugPacker Is Nothing And (mbytMode = 1 Or mbytMode = 2) Then
        Err = 0: On Error Resume Next
        Set mobjDrugPacker = CreateObject("zlDrugPacker.clsDrugPacker")
        If Err <> 0 Then
            mblnDrugPacker = False
        Else
            mblnDrugPacker = mobjDrugPacker.DYEY_MZ_IniSoap
        End If
    End If
End Sub

Private Function Load���㷽ʽ() As Boolean
'˵��:1-�ֽ���㷽ʽ,2-������ҽ������,3-ҽ�������ʻ�,4-ҽ������ͳ��
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    
    cbo�˿ʽ.Clear
    
    On Error GoTo errH
    Set rsTmp = Get���㷽ʽ("�շ�")
    For i = 1 To rsTmp.RecordCount
        If rsTmp!���� = 3 Then
            mstr�����ʻ� = rsTmp!����
        ElseIf InStr(",1,2,7,", "," & rsTmp!���� & ",") > 0 And Val(Nvl(rsTmp!Ӧ����)) = 0 Then
            cbo�˿ʽ.AddItem rsTmp!���� & "-" & rsTmp!����
            cbo�˿ʽ.ItemData(cbo�˿ʽ.NewIndex) = rsTmp!����
            If rsTmp!���� = 1 Then
                mstr�ֽ���㷽ʽ = Nvl(rsTmp!����)
            End If
            If rsTmp!���� = gstr���㷽ʽ Then
                Call zlControl.CboSetIndex(cbo�˿ʽ.hWnd, cbo�˿ʽ.NewIndex)
            End If
            If rsTmp!ȱʡ = 1 And cbo�˿ʽ.ListIndex = -1 Then
                Call zlControl.CboSetIndex(cbo�˿ʽ.hWnd, cbo�˿ʽ.NewIndex)
            End If
        End If
        
        rsTmp.MoveNext
    Next
    If mstr�ֽ���㷽ʽ = "" Then mstr�ֽ���㷽ʽ = "�ֽ�"
    If cbo�˿ʽ.ListIndex = -1 And cbo�˿ʽ.ListCount > 0 Then Call zlControl.CboSetIndex(cbo�˿ʽ.hWnd, 0)
    Load���㷽ʽ = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Resize()
    Dim staH As Long

    On Error Resume Next
    
    staH = IIf(stbThis.Visible, stbThis.Height, 0)
    
    vsBill.Height = Me.ScaleHeight - picCmd.Height - staH - picPati.Height - picMoney.Height - pic�˷�ժҪ.Height - vsBalance.Height - IIf(picInvoice.Visible, picInvoice.Height, 0)
    
    If picMoney.ScaleWidth - fra�˿�.Width - 45 < txtAllTotal.Left + txtAllTotal.Width + 90 Then
        fra�˿�.Left = txtAllTotal.Left + txtAllTotal.Width + 90
    Else
        fra�˿�.Left = picMoney.ScaleWidth - fra�˿�.Width - 45
    End If
    
    If Me.ScaleWidth - cmdSelAll.Left - cmdCancel.Width > 5500 Then
        cmdCancel.Left = Me.ScaleWidth - cmdSelAll.Left - cmdCancel.Width
    Else
        cmdCancel.Left = 5500
    End If
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 90
    
    fraInfo_1.Width = Me.ScaleWidth + 300
    LineCmd_1.x2 = Me.ScaleWidth + 300
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mbytMode = 0
    mstrNo = ""
    mstrDelTime = ""
    mstrNOsOverFlow = ""
    mblnNOMoved = False   '�鿴ʱ,���ܴ���true
    Call initCardSquareData
    Call CloseIDCard
    zlDatabase.SetPara "�˷Ѻ�������ģʽ", IIf(optNO(0).Value, "0", "1"), glngSys, 1121, InStr(1, mstrPrivs, ";��������;") > 0
    Call SaveWinState(Me, App.ProductName)
    
    If Not mobjDrugPacker Is Nothing Then
        '81190
        Set mobjDrugPacker = Nothing
    End If
End Sub

Private Sub IDKind_Click(objCard As zlIDKind.Card)
    '����:50885
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
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
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

Private Sub optNO_Click(Index As Integer)
    If Visible Then txtNO.SetFocus
    If Index = 0 Then
        If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", Mid(txtNO.Text, 2, 1)) > 0 Then
            txtNO.Text = ""
        End If
    End If
End Sub
Private Sub picInvoice_Resize()
    Err = 0: On Error Resume Next
    With fraSelectTopSplit
        .Top = picInvoice.ScaleTop
        .Left = picInvoice.ScaleLeft
        .Width = picInvoice.ScaleWidth
    End With
    With vsInvoice
        .Top = fraSelectTopSplit.Top + fraSelectTopSplit.Height + 50
        .Left = picInvoice.ScaleLeft + 50
        .Width = picInvoice.ScaleWidth - .Left * 2
    End With
    Call SetInvoceSizeAndShowTittle
    With fraSelectDownSplit
        .Top = vsInvoice.Top + vsInvoice.Height + 50
        .Left = picInvoice.ScaleLeft
        .Width = picInvoice.ScaleWidth
    End With
    picInvoice.Height = fraSelectDownSplit.Top + fraSelectDownSplit.Height + 50
End Sub

Private Sub picPati_Resize()
    txtNO.Left = picPati.ScaleWidth - txtNO.Width - 45
    optNO(1).Left = txtNO.Left - optNO(1).Width - 30
    optNO(0).Left = optNO(1).Left - optNO(0).Width - 15
    pic��.Left = picPati.ScaleWidth - pic��.Width - 45
    
    If txtPatientPrint.Visible Then
        txtPatientPrint.Left = picPati.Left + picPati.Width - txtPatientPrint.Width - 50
        lblPatiName.Left = txtPatientPrint.Left - lblPatiName.Width - 50
        txtPatientPrint.Top = txtNO.Top - 50
        lblPatiName.Top = txtNO.Top
    End If
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

Private Sub txtNO_KeyPress(KeyAscii As Integer)
    Dim strAbc As String, str1 As String, str2 As String
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txtNO.Text <> "" Then
            If optNO(0).Value Then
                txtNO.Text = GetFullNO(txtNO.Text, 13)
            End If
            Call zlControl.TxtSelAll(txtNO)
            If ReadBills(txtNO.Text) Then vsBill.SetFocus
        ElseIf txtPatient.Visible And txtPatient.Enabled Then
            txtPatient.SetFocus
        End If
    Else
        Call SetNOInputLimit(txtNO, KeyAscii, IIf(optNO(0).Value, 0, 1))
    End If
End Sub
Private Sub InitBillHead()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���˷ѵı�ͷ����Ϣ
    '����: �ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2012-09-11 09:47:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim arrHead As Variant, strHead As String, i As Long
    Dim varTemp As Variant, intCol As Integer
    
    strHead = "ѡ��,300,4;���ݺ�,1000,1;���,720,1;��Ŀ,2800,1;��Ʒ��,2000,1;����,750,7;��λ,550,1;����,1100,7;" & _
        "Ӧ�ս��,1100,7;ʵ�ս��,1100,7;��������,1000,1;ִ�п���,1000,1;����Ա,850,1;ʱ��,1260,1;����ID;ҽ��,1560,1;" & _
        "ԭʼ����,0,4;׼������,0,4;ҽ�����,0,4;ִ�п���ID,0,1"
    
    arrHead = Split(strHead, ";")
    With vsBill
        .Clear
        .FixedRows = 1: .FixedCols = 0
        .COLS = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        For i = 0 To UBound(arrHead)
            varTemp = Split(arrHead(i) & ",,,", ",")
            intCol = .FixedCols + i
            .ColKey(intCol) = varTemp(0)
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
        .ColHidden(.ColIndex("��Ʒ��")) = gTy_System_Para.bytҩƷ������ʾ <> 2
        .FrozenCols = 2
        .Editable = flexEDKbdMouse
        .ColDataType(0) = flexDTBoolean
    End With
End Sub

Private Sub txtPatient_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtPatient.Locked Or txtPatient.Enabled = False Or txtPatient.Text <> "" Then Exit Sub
    If IDKind.ActiveFastKey = True Then Exit Sub
End Sub

Private Sub txtPatient_LostFocus()
    '����:60010
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(False)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(False)
    IDKind.SetAutoReadCard (False)
End Sub

Private Sub txtPatientPrint_Validate(Cancel As Boolean)
    txtPatientPrint.Text = Trim(txtPatientPrint.Text)
End Sub
Private Sub txt�˷�ժҪ_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txt�˷�ժҪ, KeyAscii, m�ı�ʽ
End Sub

Private Sub vsBalance_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Row = 1 Then
        '����:43403
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
    Call ReCalcDelMoney
End Sub

Private Sub vsBalance_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If mbytMode = 2 Or mbytMode = 0 Then Cancel = True: Exit Sub
    With vsBalance
        '        '����:43403
        If Col Mod 2 <> 0 Then Cancel = True: Exit Sub
        If Row <> 1 Then Cancel = True: Exit Sub
        If Val(.ColData(Col)) = 0 Or (mCurBillType.bln������ȫ�� And .RowHidden(1)) Then Cancel = True: Exit Sub
        '1.����������֧�ֲ����ˣ���ѡ���˻��������Ľ��
        '2.�������ֲ�֧�ֲ����ˣ�����ȫ��ʱ��ѡ���˻��������Ľ��
        .ColComboList(Col) = " ||" & FormatEx(Val(.Cell(flexcpData, Row, Col)), 2)
    End With
End Sub

Private Sub vsBalance_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If vsBalance.MouseCol > 0 Then vsBalance.ToolTipText = vsBalance.ColData(vsBalance.MouseCol)  '��ʾ����ժҪ
End Sub
Private Sub zlSet���ƹ̶���ϵ(ByVal lngRow As Long, ByVal Col As Long, Optional lngNotCheckRow As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:סԺ���ü�¼
    '����:���˺�
    '����:2010-12-31 15:49:15
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
                                 If .RowData(i) = Val(.Cell(flexcpData, j, .ColIndex("��Ŀ"))) Then
                                        .Cell(flexcpChecked, j, .ColIndex("ѡ��")) = .Cell(flexcpChecked, i, .ColIndex("ѡ��"))
                                         .TextMatrix(j, .ColIndex("ѡ��")) = .TextMatrix(i, .ColIndex("ѡ��"))
                                 End If
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
    
    With vsBill
        If Col = .ColIndex("ѡ��") Then
            If mBillDelType = EM_����ȫ�� Then
                vsBill.TextMatrix(Row, .ColIndex("ѡ��")) = 1
                    '���ݵ���ѡ��Ʊ
                    Call FromNoSelectInvoice
                Exit Sub
            ElseIf mBillDelType = EM_����ȫ�� Then
                    Call SetNOBill(vsBill.TextMatrix(Row, .ColIndex("���ݺ�")), Val(vsBill.TextMatrix(Row, .ColIndex("ѡ��"))) <> 0)
                    Call LoadBalanceInfor
                    Call LoadDelBalanceInfor
                    Call ReCalcDelMoney
                    '���ݵ���ѡ��Ʊ
                    Call FromNoSelectInvoice
            
                    Exit Sub
            End If
             stbThis.Panels(2).Text = ""
            '29201
            If Val(vsBill.Cell(flexcpData, Row, .ColIndex("��Ŀ"))) = 0 Then
                For i = Row + 1 To vsBill.Rows - 1
                     If Val(vsBill.RowData(Row)) = Val(vsBill.Cell(flexcpData, i, .ColIndex("��Ŀ"))) Then
                           vsBill.TextMatrix(i, .ColIndex("ѡ��")) = vsBill.TextMatrix(Row, .ColIndex("ѡ��"))
                     Else
                        Exit For
                     End If
                Next
                Call zlSet���ƹ̶���ϵ(Row, Col)
            Else
                Call zlSet���ƹ̶���ϵ(Row, Col)
                '��Ҫ��������Ƿ��Ѿ���
                    For i = Row - 1 To 1 Step -1
                        If Val(vsBill.RowData(i)) = Val(vsBill.Cell(flexcpData, Row, .ColIndex("��Ŀ"))) Then
                            If vsBill.TextMatrix(i, .ColIndex("ѡ��")) <> 0 Then
                                vsBill.TextMatrix(i, .ColIndex("ѡ��")) = vsBill.TextMatrix(Row, .ColIndex("ѡ��"))
                            End If
                            Call zlSet���ƹ̶���ϵ(i, Col, Row)
                             Exit For
                        End If
                    Next
            End If
            Call LoadBalanceInfor
            Call LoadDelBalanceInfor
            Call ReCalcDelMoney
            '���ݵ���ѡ��Ʊ
            Call FromNoSelectInvoice
        End If
    End With
End Sub

Private Sub vsBill_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim cur�ϼ� As Currency, i As Long
        
    If NewRow <> OldRow Then
        With vsBill
            If .TextMatrix(NewRow, .ColIndex("���ݺ�")) <> "" Then
                For i = NewRow - 1 To .FixedRows Step -1
                    If .TextMatrix(i, .ColIndex("���ݺ�")) = .TextMatrix(NewRow, .ColIndex("���ݺ�")) Then
                        cur�ϼ� = cur�ϼ� + Val(.TextMatrix(i, .ColIndex("ʵ�ս��")))
                    Else
                        Exit For
                    End If
                Next
                For i = NewRow To .Rows - 1
                    If .TextMatrix(i, .ColIndex("���ݺ�")) = .TextMatrix(NewRow, .ColIndex("���ݺ�")) Then
                        cur�ϼ� = cur�ϼ� + Val(.TextMatrix(i, .ColIndex("ʵ�ս��")))
                    Else
                        Exit For
                    End If
                Next
            End If
            txtCurTotal.Text = Format(cur�ϼ�, gstrDec)
        End With
    End If
End Sub

Private Sub vsBill_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
       With vsBill
            Select Case Col
            Case .ColIndex("ѡ��")
                If mBillDelType = EM_����ȫ�� Then Cancel = True: Exit Sub
            End Select
        End With
End Sub

Private Sub vsBill_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = vsBill.ColIndex("ѡ��") Then Cancel = True
End Sub

Private Sub GetBillRow(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long)
    Dim i As Long
    
    lngBegin = lngRow: lngEnd = lngRow
    With vsBill
        For i = lngRow - 1 To .FixedRows Step -1
            If .TextMatrix(i, .ColIndex("���ݺ�")) = .TextMatrix(lngRow, .ColIndex("���ݺ�")) Then
                lngBegin = i
            Else
                Exit For
            End If
        Next
        For i = lngRow To .Rows - 1
            If .TextMatrix(i, .ColIndex("���ݺ�")) = .TextMatrix(lngRow, .ColIndex("���ݺ�")) Then
                lngEnd = i
            Else
                Exit For
            End If
        Next
    End With
End Sub

Private Sub vsBill_CellChanged(ByVal Row As Long, ByVal Col As Long)
    
    With vsBill
        If .ColIndex("���ݺ�") < 0 Then Exit Sub
        If .TextMatrix(Row, .ColIndex("���ݺ�")) <> "" And InStr(1, mstrNOsOverFlow, .TextMatrix(Row, .ColIndex("���ݺ�"))) > 0 Then
             .TextMatrix(Row, .ColIndex("ѡ��")) = 0
        End If
    End With
End Sub

Private Sub vsBill_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
'˵����1.OwnerDrawҪ����ΪOver(������Ԫ��������)
'      2.Cell��GridLine�������������ڶ��Ǵӵ�1���߿�ʼ
'      3.Cell��Border�������Ǵӵ�2���߿�ʼ,�����Ǵӵ�1���߿�ʼ
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
            If .TextMatrix(.Row, .ColIndex("���ݺ�")) <> "" Then
                If mBillDelType = EM_����ȫ�� Then Exit Sub
                If .TextMatrix(.Row, .ColIndex("ѡ��")) = 0 _
                    And InStr(1, mstrNOsOverFlow, .TextMatrix(.Row, .ColIndex("���ݺ�"))) <= 0 Then
                     .TextMatrix(.Row, .ColIndex("ѡ��")) = -1
                Else
                     .TextMatrix(.Row, .ColIndex("ѡ��")) = 0
                End If
                If mBillDelType = EM_����ȫ�� Then
                    Call SetNOBill(.TextMatrix(.Row, .ColIndex("���ݺ�")), Val(.TextMatrix(.Row, .ColIndex("ѡ��"))) <> 0)
                    Call LoadDelBalanceInfor
                    Call ReCalcDelMoney
                    Exit Sub
                End If
                 Call ReCalcDelMoney
            End If
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
        If Col <> .ColIndex("ѡ��") Then Cancel = True
    End With
End Sub

Private Sub ClearFace(Optional ByVal blnNO As Boolean = True, Optional blnCllAllData As Boolean = True)
    Set mrsBalance = Nothing
    Set mcolError = Nothing
    Call ClearBalance
    vsBill.Rows = vsBill.FixedRows
    vsBill.Rows = vsBill.FixedRows + 1
    vsBill.Row = vsBill.FixedRows: vsBill.Col = vsBill.ColIndex("��Ŀ")
    If blnCllAllData = True Then
        vsBalance.COLS = 1
        vsBalance.TextMatrix(0, 0) = IIf(mstrDelTime = "", "�տ����", "�˿����")
    End If
    mstrNOs = ""
    mintInsure = 0: mstrDelNOs = ""
    mblnYB�������� = False  '��ͬ��ҽ������֧�ֲ�һ��,����Ҫ���
    txt�˿���.ToolTipText = ""    '��¼�������
    
    lblPati.Caption = "����:"
    If blnNO Then txtNO.Text = ""
    If blnCllAllData Then
        txtCurTotal.Text = ""
        txtAllTotal.Text = ""
        txt�˿���.Text = "":
        stbThis.Panels(2).Text = ""
    End If
    If (mbytMode = 1 Or mbytMode = 2) And blnCllAllData Then
        Call Load���㷽ʽ
        If mbytMode = 1 Then
            cmdSelAll.Visible = True
            cmdClear.Visible = True
            cmdBillSel.Visible = True
        End If
    End If
End Sub
Private Sub initInsurePara(ByVal strNo As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ҽ������
    '���:strNo-���ݺ�
    '����:���˺�
    '����:2012-09-12 09:35:26
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long, lng����ID As Long
    mintInsure = ChargeExistInsure(strNo, lng����ID, lng����ID)
    If mintInsure = 0 Then Exit Sub
    MCPAR.�൥���շѱ���ȫ�� = gclsInsure.GetCapability(support�൥���շѱ���ȫ��, lng����ID, mintInsure)
    mblnYB�������� = gclsInsure.GetCapability(support�����������, lng����ID, mintInsure)
    MCPAR.�൥��һ�ν��� = gclsInsure.GetCapability(support�൥��һ�ν���, lng����ID, mintInsure)
    MCPAR.ҽ���ӿڴ�ӡƱ�� = gclsInsure.GetCapability(supportҽ���ӿڴ�ӡƱ��, lng����ID, mintInsure, CStr(lng����ID))
    MCPAR.�˷Ѻ��ӡ�ص� = gclsInsure.GetCapability(support�˷Ѻ��ӡ�ص�, lng����ID, mintInsure)
    MCPAR.�൥�ݵ�һ�ν��� = gclsInsure.GetCapability(support����_���ֵ��ݽ���, lng����ID, mintInsure)
End Sub
Private Function CheckPrivsIsValied() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������Ա�Ƿ�߱������˷ѵ�
    '����:�߱�����true,���򷵻�False
    '����:���˺�
    '����:2012-09-12 09:46:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If Not (mbytMode = 1 Or mbytMode = 2) Then CheckPrivsIsValied = True: Exit Function
    
    '���Ȩ���Ƿ�����
    If mintInsure > 0 Then
        '�����˷�Ȩ�޼��
        If InStr(mstrPrivs, ";�����շ�;") = 0 Then
            Screen.MousePointer = 0
            MsgBox "��û��Ȩ�޶�ҽ�����˵ĵ����˷ѣ�", vbInformation, gstrSysName
            Exit Function
        End If
        CheckPrivsIsValied = True: Exit Function
    End If
    '��ͨ���˵Ĵ���
    '�Ƿ��з�ҽ�����˵��˷�Ȩ��
    If InStr(mstrPrivs, ";�����ҽ������;") = 0 Then
        Screen.MousePointer = 0
        MsgBox "��û��Ȩ�޶Է�ҽ�����˽����˷Ѳ�����", vbInformation, gstrSysName
        Exit Function
    End If
    CheckPrivsIsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetBillDelType(ByVal strNos As String) As EM_BillDelType
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�����ݵ��˷�����
    '���:strNos-���β����ĵ���
    '����:��������˷ѹ���,���ر����˷ѵ�����(����ȫ��;����ȫ��;���Ų�����)
    '����:���˺�
    '����:2012-09-12 09:59:23
    '�˷ѹ���˵������:
    '    ��ͨ����:
    '        1.����Ƕ൥�ݶ��ֽ��㷽ʽ
    '            a)���ֻ��һ�ֽ��㷽ʽ,�����������е�һ��,��Ҫѡ���˵Ľ��㷽ʽ
    '            b)����ж��ֽ��㷽ʽ,��ֻ�������е�һ�ŵ���,������һ��.
    '        2. ������ڽ��㿨:
    '            a)����Ƿ�����=0�Ļ�,������Ƿ�ȫ����ȷ���Ƿ�ȫ���˻��ߵ�����
    '            b)����Ƿ�����=1�Ļ�,����������,������ʱ,�˳�ָ���Ľ��㷽ʽ
    '        3. �������ҽ�ƿ�:
    '           a.���ֻ��һ�ֽ��㷽ʽ
    '            a)���������Ϊ"ȫ��"�Ҳ�֧������(�Ƿ�����=0),�����е��ݱ���ȫ��Ϊԭ���㷽ʽ
    '            b)���������Ϊ"ȫ��"��֧������(�Ƿ�����=1),������ѡ�񵥱���,�˳�ָ���Ľ��㷽ʽ
    '            c)���������Ϊ"������"�Ҳ�֧������(�Ƿ�����=0),������ѡ�񵥱���,��ֻ��Ϊԭ���㷽ʽ
    '            d)���������Ϊ"������"��֧������(�Ƿ�����=1),������ѡ�񵥱���,�˳�ָ���Ľ��㷽ʽ
    '           b.������ڶ��ֽ��㷽ʽ
    '            a)����Ƿ�����=0�Ļ�,������Ƿ�ȫ����ȷ���Ƿ���ȫ�˻�����ȫ��
    '            b)����Ƿ�����=1�Ļ�,����������,������ʱ,�˳�ָ���Ľ��㷽ʽ
    '        4. �������һ��ͨ:Ӧ�ð����ŵ���ȫ��
    '   ҽ������:
    '        a.support�൥���շѱ���ȫ��:Ϊtrueʱ,���ܲ�����
    '        b.����������ͨ���˵Ĵ���ʽһ��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnDelSingleNO As Boolean
    On Error GoTo errHandle
    mblnSingleBlance = False
   '35461
    If mintInsure > 0 Then        'ҽ������
        If InStr(strNos, ",") > 0 And MCPAR.�൥���շѱ���ȫ�� Then
            mBillDelType = EM_����ȫ��: Exit Function
        End If
        blnDelSingleNO = True
    End If
    
    If mintInsure = 0 Then
        '��ͨ���˵Ĵ���
        mBillDelType = EM_���Ų�����
        '������е��ݶ�ֻʹ����һ�ֽ���,���������˷�(�����е�һ��,��һ���еļ���,�����еļ���)
        mblnSingleBlance = CheckSingleBalance(Replace(strNos, "'", ""))
        '���ֽ��㷽ʽ,���뵥��ȫ��
        'If Not mblnSingleBlance Then blnDelSingleNO = True
    End If
    
    '���������������
    '       1.һ��ͨ���ڱ���ȫ��,ֱ�ӷ���
    '       2.��������,ֻҪ��ҽ�ƿ����.�Ƿ�ȫ��Ϊtrue�Ҳ�������,�򷵻�true
    '����:1-Ԥ���,2-ҽ��,3-ҽ�ƿ�(һ��ͨ),4-���㿨,5-һ��ͨ,0-������
    mrsBalance.Filter = "����>2"
    With mrsBalance
        Do While Not .EOF
            '�Ƿ�ȫ��
            Select Case mrsBalance!����
            Case 5:  '1.һ��ͨ����,����ȫ��,ֱ�ӷ���
                If InStr(strNos, ",") > 0 Then
                    GetBillDelType = EM_����ȫ��: Exit Function
                End If
                GetBillDelType = EM_����ȫ��: Exit Function
            Case 3:  'ҽ�ƿ�
                If Val(Nvl(!�Ƿ�ȫ��)) = 1 And Val(Nvl(!�Ƿ�����)) = 0 Then 'ȫ�˲�����
                    If InStr(strNos, ",") > 0 Then
                        GetBillDelType = EM_����ȫ��: Exit Function
                    End If
                    GetBillDelType = EM_����ȫ��: Exit Function
                End If
                If Not mCurBillType.blnSingleBalance Then
                    If Val(Nvl(!�Ƿ�����)) = 0 Then
                        '���ű���ȫ��,��֧��ȫ�˵�,����뵥��ȫ��
                        blnDelSingleNO = True
                    End If
                End If
            Case 4: '���㿨
                If Val(Nvl(!�Ƿ�ȫ��)) = 1 And Val(Nvl(!�Ƿ�����)) = 0 Then
                    If InStr(strNos, ",") > 0 Then
                        GetBillDelType = EM_����ȫ��: Exit Function
                    End If
                    GetBillDelType = EM_����ȫ��: Exit Function
                End If
                If Val(Nvl(!�Ƿ�����)) = 0 Then
                    '���ű���ȫ��,��֧��ȫ�˵�,����뵥��ȫ��
                    blnDelSingleNO = True
                End If
            End Select
            .MoveNext
        Loop
    End With
    GetBillDelType = IIf(blnDelSingleNO, EM_����ȫ��, EM_���Ų�����)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function CheckDelIsValied(ByVal strNos As String, ByRef strNotCanDelNOs As String, ByRef strCanDelNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����˷ѵ����Ƿ�Ϸ�
    '����:strNotCanDelNOs-�����˵ĵ���(�Ѿ�ִ�м����ܼ޵ĵ���)
    '        strCanDelNos-���˵ĵ���
    '����:�Ϸ�����true,���򷵻�False
    '����:���˺�
    '����:2012-09-12 15:12:40
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim arrNo As Variant, i As Long, intTmp As Integer
    Dim strInfo As String, strFlagPrintInfor As String
    Dim blnFlagPrint As Boolean, strNo As String, strCurNO As String
    Dim strOper As String, vDate As Date, blnHaveExe As Boolean
    
    On Error GoTo errHandle
    
    '����:54728
    If Not mbytMode = 1 Then CheckDelIsValied = True: Exit Function   '�˷�ʱ�ж�
    
    arrNo = Split(strNos, ",")
    strNotCanDelNOs = ""
     '�Ƿ���ִ��
    strCanDelNos = ""   '��¼�����˵ĵ��ݺ�
    strInfo = ""        '�������ʾ��Ϣ
    strFlagPrintInfor = ""
    For i = 0 To UBound(arrNo)
        strCurNO = Replace(arrNo(i), "'", "")
        If strNo = "" Then strNo = strCurNO
        If i = 0 Then
            If Not ReadBillInfo(1, strCurNO, 1, strOper, vDate) Then
                strInfo = "����[" & strCurNO & "]������!"
                Exit For
            End If
            If InStr(mstrPrivs, "���в���Ա") <= 0 And UserInfo.���� <> strOper Then
                strInfo = "��û��""���в���Ա""Ȩ��,���ܶ�" & strOper & "�ĵ��ݽ����˷�!"
                Exit For
            End If
            If Not BillOperCheck(2, strOper, vDate, "�˷�", strCurNO, , 1) Then
                Screen.MousePointer = 0:  Exit Function
            End If
        End If
        
        blnHaveExe = False: blnFlagPrint = False
        intTmp = BillCanDelete(strCurNO, 1, blnHaveExe, , blnFlagPrint)
        If intTmp <> 0 Then
            strNotCanDelNOs = strNotCanDelNOs & "," & strCurNO  '�����˵ĵ���
            If Not mCurBillType.bln���Ų����˷� Then mCurBillType.bln���Ų����˷� = True
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
            '������ִ����Ŀ
            If mintInsure > 0 Then '�շ�ҽ���˷�
                strInfo = strInfo & "[" & strCurNO & "]������ҽ�����˵��շѵ�,�����Ѿ�ִ�е���Ŀ,�����˷ѣ�" & vbCrLf
                strNotCanDelNOs = strNotCanDelNOs & "," & strCurNO  '�����˵ĵ���
            ElseIf mBillDelType <> EM_���Ų����� Then
                strInfo = strInfo & "[" & strCurNO & "]�д�����ִ�е���Ŀ,�����˷ѡ�" & vbCrLf
                strNotCanDelNOs = strNotCanDelNOs & "," & strCurNO  '�����˵ĵ���
            Else
                strInfo = strInfo & "[" & strCurNO & "]�д�����ִ�е���Ŀ���˵��ݽ�ִ�е��ǲ����˷ѡ�" & vbCrLf
                strCanDelNos = strCanDelNos & "," & strCurNO
            End If
            If Not mCurBillType.bln���Ų����˷� Then mCurBillType.bln���Ų����˷� = True
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
    
    If strNotCanDelNOs <> "" And mBillDelType = EM_����ȫ�� Then
        Screen.MousePointer = 0
        MsgBox "�����շѱ���ȫ��:" & vbCrLf & strInfo, vbInformation, gstrSysName
        Exit Function
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
    CheckDelIsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub InitBillVar(ByVal strNos As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������
    '����:���˺�
    '����:2012-09-17 13:29:53
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str���㷽ʽ As String
    With mCurBillType
        .bln�൥�� = InStr(strNos, ",") > 0
        .strNos = strNos
        .bln���ڿ����� = False
        .bln����ҽ�ƿ����� = False
        .bln������ȫ�� = False
    End With
    '���������������
    '       1.һ��ͨ���ڱ���ȫ��,ֱ�ӷ���
    '       2.��������,ֻҪ��ҽ�ƿ����.�Ƿ�ȫ��Ϊtrue�Ҳ�������,�򷵻�true
    '����:1-Ԥ���,2-ҽ��,3-ҽ�ƿ�(һ��ͨ),4-���㿨,5-һ��ͨ,0-������
    mrsBalance.Filter = "����<>2 And ����<>1"
    str���㷽ʽ = ""
    mrsBalance.Sort = "NO,����"
    'W.NO,A.����ID
    With mrsBalance
        Do While Not .EOF
            If InStr(str���㷽ʽ & ",", "," & Nvl(!���㷽ʽ) & ",") = 0 Then
                str���㷽ʽ = str���㷽ʽ & "," & Nvl(!���㷽ʽ)
            End If
            If Val(Nvl(!����)) = 4 Or Val(Nvl(!����)) = 3 Then mCurBillType.bln���ڿ����� = True
            If Val(Nvl(!����)) = 3 Then mCurBillType.bln����ҽ�ƿ����� = True
             
            .MoveNext
        Loop
    End With
    If str���㷽ʽ <> "" Then str���㷽ʽ = Mid(str���㷽ʽ, 2)
    mCurBillType.bln���ֽ��㷽ʽ = InStr(str���㷽ʽ, ",") = 0
    mCurBillType.str���㷽ʽ = str���㷽ʽ
    
    str���㷽ʽ = ""
    mrsBalance.Filter = 0
    With mrsBalance
        Do While Not .EOF
            If InStr(str���㷽ʽ & ",", "," & Nvl(!���㷽ʽ) & ",") = 0 Then
                str���㷽ʽ = str���㷽ʽ & "," & Nvl(!���㷽ʽ)
            End If
            .MoveNext
        Loop
    End With
    If str���㷽ʽ <> "" Then str���㷽ʽ = Mid(str���㷽ʽ, 2)
    mCurBillType.blnSingleBalance = InStr(str���㷽ʽ, ",") = 0
End Sub



Private Function ReadBills(ByVal strNo As String) As Boolean
'���ܣ����ݵ�ǰ����ĵ��ݺŻ�Ʊ�ݺ�,��ȡ����ʾ���ŵ���
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strSQL1 As String, strSQL2 As String, strSQL3 As String
    Dim strNos As String, strBalance As String, str���� As String
    Dim strSub As String, strTmp As String, strCurNO As String, strCanDelNos As String
    Dim blnDelSingleNO As Boolean
    Dim i As Long, intTmp As Integer, intSign As Integer, j As Integer
    Dim arrNOs As Variant, cur�ϼ� As Currency, arrNo As Variant
    Dim strTemp As String, strҽ����� As String
    Dim blnNotFind As Boolean
    Dim lng����ID As Long, cllInvoiceNoInfor As Collection
    Dim str������� As String
    Dim strInvoiceNO As String, strOldNO As String
    
    On Error GoTo errH
    Set mrsDelInvoice = Nothing
    mblnSingleBlance = False
    Call ClearFace(False)
    strOldNO = strNo
    
    Call ClearVar   '�����ǰ���ݵ���ر���
    Screen.MousePointer = 11
    Set cllInvoiceNoInfor = New Collection
    'ȷ��һ���շѵĶ��ŵ���
    '����Ƕ����˷�״̬mbytMode = 1,�Ƿ������߱���ȷ������Ҫ���Ӳ�ѯ
    '----------------------------------------------------------------------------------
    '56963
    strInvoiceNO = ""
    If Not (mstrNo <> "" Or optNO(0).Value) Then
         '��Ʊ�ݺ�:���ܲ�ͬ����Ʊ���ظ�
        strInvoiceNO = strNo
        strNos = zlInvoiceFromNOs(strInvoiceNO, True, str�������, cllInvoiceNoInfor)
        If InStr(str�������, ",") > 0 Then
            '֤���ж���������,��Ҫ����Աȷ���Ĵη���
            strNo = ""
            Screen.MousePointer = 0
            If frmMulitChargeSelect.zlShowSelect(Me, mlngModule, strNos, strInvoiceNO, strNo) = False Then
                Screen.MousePointer = 0
                Exit Function
            End If
            If strNo = "" Then
                Screen.MousePointer = 0
                Exit Function
            End If
            Screen.MousePointer = 99
        Else
            strNo = Replace(Split(strNos & ",", ",")(0), "'", "")
        End If
        strOldNO = ""
        strNos = GetMultiNOs(strNo, , , True, True)
    End If
    
    Dim strTempNos As String, intInsure As Integer
    strTempNos = GetMultiNOs(strNo, , , True, True)
    strNos = GetMultiNOs(strNo, , , False, True)
    mCurBillType.bln����������� = False
    If mbytMode = 0 Then
        strNos = strTempNos
    Else
        If InStr(strTempNos, ",") > 0 And InStr(strNos, ",") = 0 Then
            '�϶��ǰ����ݷֱ��ӡ��
            'Ҫ�൥���˷�,�ͱ���������������
            '1.ҽ���൥�ݱ���ȫ��ʱ,���밴������Ž����˷�
            '2.�����˻�ȫ��ʱ,���밴������Ž����˷�
            intInsure = ChargeExistInsure(strNo)
            If intInsure <> 0 Then
                If gclsInsure.GetCapability(support�൥���շѱ���ȫ��, , intInsure) Then
                    strNos = strTempNos: mCurBillType.bln����������� = True
                End If
            ElseIf zlIsExistsSquareCard(strTempNos, True) Then
                '���һ��ͨ���㲿���Ƿ����ȫ�˵�
                strNos = strTempNos: mCurBillType.bln����������� = True
            Else
                If mbytMode = 1 And mstrNo <> "" And Not mblnFromInNewDel Then
                    If frmMulitChargeSelect.zlShowSelect(Me, mlngModule, strTempNos, "", strNos, True) = False Then
                        Screen.MousePointer = 0
                        Exit Function
                    End If
                End If
            End If
        End If
    End If
    
    If strNos = "" Then
        If optNO(1).Value Then
            Screen.MousePointer = 0
            MsgBox "û���ҵ������""" & strNo & """��ص��շѼ�¼��", vbInformation, gstrSysName
            Exit Function
        End If
        '������Ϊδ��Ʊ�ݶ���������
        strNos = strNo
    End If
    '��Ҫ������
    If InStr(1, strNos, "'") = 0 Then
        strNos = "'" & Replace(strNos, ",", "','") & "'"
    End If
    'Ƚ����:ѡ��ĵ��ݽ�����ҽ��������㣬�������˷�
    If mbytMode <> 0 Then
        If CheckBillExistReplenishData(1, , Replace(strNos, "'", "")) = True Then
            Screen.MousePointer = 0
            MsgBox "��ǰ���ݽ�����ҽ��������㣬����������˷Ѳ�����", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    arrNo = Split(strNos, ",")
    
    If gbln�˷�����ģʽ And mbytMode = 1 Then
        Set rsTmp = GetApply(strNo, 1)
        rsTmp.Filter = "״̬<>2"
        If rsTmp.RecordCount = 0 Then
            Screen.MousePointer = 0
            MsgBox "���ȶԸõ��ݽ����˷����룡", vbInformation, gstrSysName
            Exit Function
        End If
        If IsNull(rsTmp!�����) Then
            Screen.MousePointer = 0
            MsgBox "�õ���δ�����˷���ˣ����ܽ����˷ѣ�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '����:1-Ԥ���,2-ҽ��,3-ҽ�ƿ�,4-���㿨,5-һ��ͨ,0-������
    intSign = IIf(mstrDelTime <> "", -1, 1) '����,�����������
    Set mrsBalance = GetChargeBalance(strNos, , , , mstrDelTime, intSign)
    Call initInsurePara(strNo)
    If CheckPrivsIsValied = False Then Exit Function    '����Ȩ�޼��
    
    Call InitBillVar(strNos)    '��ʼ����ǰ�˷ѵĵ�����Ϣ����
    
    'ȷ���˷�����
    mBillDelType = GetBillDelType(strNos)
    
    '�˷���ؼ��
    If CheckDelIsValied(strNos, mstrDelNOs, strCanDelNos) = False Then
        Screen.MousePointer = 0:  Exit Function
        Exit Function
    End If
    If strCanDelNos <> "" Then strNos = strCanDelNos
        
    '��ȡ������Ϣ
    '----------------------------------------------------------------------------------
    strSQL = "" & _
    " Select A.����ID,A.����,A.�Ա�,A.����,A.��ʶ��,A.�ѱ�,C.���� as ���ʽ,B.����,E.��������" & _
    " From " & IIf(mblnNOMoved, "H", "") & "������ü�¼ A,������Ϣ E,���ս����¼ B,ҽ�Ƹ��ʽ C,��Ա�� D" & _
    " Where A.����ID=E.����ID(+) And A.���ʽ=C.����(+) And A.����ID=B.��¼ID(+) And B.����(+)=1 And A.����Ա����=D.����" & _
    "       And (D.վ��='" & gstrNodeNo & "' Or D.վ�� is Null)" & vbNewLine & _
    "       And A.��¼����=1 And A.��¼״̬ IN(1,3) And A.NO=[1] And Rownum < 2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo)
    If rsTmp.EOF Then
        Screen.MousePointer = 0
        MsgBox "û���ҵ������""" & strNo & """��ص��շѼ�¼��", vbInformation, gstrSysName
        mlng����ID = 0
        Exit Function
    End If
    txtPatient.Text = Nvl(rsTmp!����)
    
    lblPati.Caption = "����:" & IIf(txtPatient.Visible, "       ", rsTmp!����) & _
        "���Ա�:" & Nvl(rsTmp!�Ա�) & _
        "������:" & Nvl(rsTmp!����) & _
        "�������:" & Nvl(rsTmp!��ʶ��) & _
        "���ѱ�:" & Nvl(rsTmp!�ѱ�) & _
        "�����ʽ:" & rsTmp!���ʽ
        
    mlng����ID = Val(Nvl(rsTmp!����ID))
    With mtyPati
        .����ID = mlng����ID
        .�Ա� = Nvl(rsTmp!�Ա�)
        .���� = Nvl(rsTmp!����)
        .���� = Nvl(rsTmp!����)
    End With
    
    If Not IsNull(rsTmp!����) Then
        lblPati.ForeColor = vbRed
        txtYB.Text = Val(Nvl(rsTmp!����))   '����:41760
        txtPatient.ForeColor = vbRed
    Else
        lblPati.ForeColor = &HC00000
        txtYB.Text = ""
        txtPatient.ForeColor = &HC00000
    End If
    '75259�����ϴ�,2014-7-10������������ʾ��ɫ����
    Call SetPatiColor(txtPatient, Nvl(rsTmp!��������), txtPatient.ForeColor)
    lblPati.ForeColor = txtPatient.ForeColor
     If mblnPrintView And InStr(1, mstrPrivs, "�޸������ش�") > 0 And IsNull(rsTmp!����ID) Then
        txtPatientPrint.Text = "" & rsTmp!����
        txtPatientPrint.Tag = txtPatientPrint.Text
        txtPatientPrint.Visible = True
        lblPatiName.Visible = True
    End If
    
    '��ȡ��������:ԭʼ���˷ѵ�,���㷽ʽΪ��ָ��Ԥ���ļ�¼
    '----------------------------------------------------------------------------------
    Call LoadBalanceInfor
      mintReturnMode = cbo�˿ʽ.ListIndex  '�����˷�ʱ,ȫ�˽��ý��㷽ʽʱ�ָ���ʼ�Ľ��㷽ʽ
    '��ȡ��������
    '----------------------------------------------------------------------------------
    
    If mbytMode = 1 Then
        '0-���ŵ��ݲ鿴,1-���ŵ����˷�
        '�˷�ʱ���ÿ��Ǻ󱸱�,ǰ��Ĳ����ѽ���
        '��ȡ׼����,������Ӧ�ս��,ʵ�ս��(���=ʣ����*(׼����/ʣ����))
        '��ȡ������ԭʼ��¼�ķ���ID
        Dim strTableNo As String
        mblnHaveExcuteData = zlCheckIsExcuteData(Replace(strNos, "'", ""), 1)     '60735
        
        '���˺�45685,58077
        strTableNo = "" & _
        "   With ���ﵥ�� as (Select  Column_Value as No From Table(f_Str2list([2])))," & _
        "            �������  as (" & _
        "           Select A.ID,A.��¼����,A.NO,A.��¼״̬,A.���,A.��������,A.�۸񸸺�,A.�շ�ϸĿID, " & _
        "                      nvl(A.����,1) as ����, nvl(A.����,0) as ����, " & _
        "                      nvl(A.Ӧ�ս��,0) as Ӧ�ս�� ,nvl(A.ʵ�ս��,0) as ʵ�ս��,nvl(A.���ʽ��,0) as ���ʽ��," & _
        "                      Nvl(A.����,1)*A.���� as ����, nvl(��׼����,0)  as ��׼����," & _
                               IIf(gblnҩ����λ, "nvl(B." & gstrҩ����װ & ",1)", "1") & " as ����ϵ��, " & _
                               IIf(gblnҩ����λ, " decode(B.ҩƷID,NULL,A.���㵥λ,B." & gstrҩ����λ & ")", "A.���㵥λ ") & " as ���㵥λ," & _
        "                      A.��������ID,A.ִ�в���ID,A.ҽ�����, " & _
        "                      A.ִ��״̬,A.��������,A.����״̬ ,A.���ӱ�־,A.�ѱ�,A.�շ����,A.����Ա����,A.�Ǽ�ʱ��,A.����ID," & _
        "                      B.ҩƷID" & _
        "           From ������ü�¼ A,ҩƷ��� B,���ﵥ�� J  " & _
        "           Where A.��¼����=1 And A.NO=J.NO and A.��¼״̬<>0" & _
        "                       And A.�շ�ϸĿID=B.ҩƷID(+)" & _
        "              )," & _
        ""
        '��׼�˷�(����,ҩƷ,����������)
        strTableNo = strTableNo & vbCrLf & _
        "            ׼����  as ( " & _
        "            Select  A.����ID,Sum(Nvl(A.����,1)*A.ʵ������" & IIf(gblnҩ����λ, "/Nvl(B." & gstrҩ����װ & ",1)", "") & ") as ׼������" & _
        "            From ҩƷ�շ���¼ A,ҩƷ��� B, ���ﵥ�� J" & _
        "           Where A.ҩƷID=B.ҩƷID(+) And Mod(A.��¼״̬,3)=1  " & _
        "                       And (A.���� =8 or a.����=24) And A.����� is NULL And A.NO =J.NO" & _
        "           Group by A.����ID"
        
        '��������ص�׼����
        If mblnHaveExcuteData Then
            '60735:��ҽ��ִ�мƼ��д�������ʱ,��ҽ��ִ�мƼ���ȡ��
            strTableNo = strTableNo & " Union ALL  " & _
            " Select Max(ID) As ����id, Decode(Sign(Sum(����)), -1, 0, Sum(����)) As ׼����" & vbNewLine & _
            " From ( Select Decode(a.��¼״̬, 2, 0, a.Id) As ID, a.ҽ����� As ҽ��id, a.�շ�ϸĿid, Nvl(a.����, 1) * Nvl(a.����, 1) As ����," & vbNewLine & _
            "              Decode(a.��¼״̬, 2, 0, Nvl(a.����, 1) * Nvl(a.����, 1)) As ԭʼ����" & vbNewLine & _
            "       From ������� A, ����ҽ����¼ M" & vbNewLine & _
            "       Where a.ҽ����� = m.Id And Instr(',C,D,F,G,K,', ',' || m.������� || ',') = 0 And Instr('5,6,7', a.�շ����) = 0 And" & vbNewLine & _
            "            ��a.��¼״̬ In (1, 2, 3)��and  a.�۸񸸺� is null " & vbNewLine & _
            "          And Not Exists" & _
            "                (Select 1 From ����ҽ������ Where a.ҽ����� = ҽ��id and a.No = NO and Mod(a.��¼����, 10) = ��¼����)" & _
            "       Union All" & vbNewLine & _
            "       Select a.Id, a.ҽ����� As ҽ��id, a.�շ�ϸĿid, -1 * b.���� As ��ִ��, 0 ԭʼ����" & vbNewLine & _
            "       From ������� A, ҽ��ִ�мƼ� B, ����ҽ����¼ M" & vbNewLine & _
            "       Where a.ҽ����� = b.ҽ��id And a.�շ�ϸĿid = b.�շ�ϸĿid + 0 And a.ҽ����� = m.Id And Instr(',C,D,F,G,K,', ',' || m.������� || ',') = 0" & vbNewLine & _
            "           And Instr('5,6,7', a.�շ����) = 0" & vbNewLine & _
            "           And (Exists (Select 1  From ����ҽ��ִ��  Where b.ҽ��id = ҽ��id And b.���ͺ� = ���ͺ� And b.Ҫ��ʱ�� = Ҫ��ʱ�� And Nvl(ִ�н��, 0) = 1)" & vbNewLine & _
            "                Or Exists (Select 1 From ����ҽ������ Where b.ҽ��id = ҽ��id And b.���ͺ� = ���ͺ� And Nvl(ִ��״̬, 0) = 1))" & vbNewLine & _
            "          And a.��¼״̬ In (1, 3)��and a.�۸񸸺� Is Null " & vbNewLine & _
            "          And Not Exists" & _
            "                (Select 1 From ����ҽ������ Where a.ҽ����� = ҽ��id and a.No = NO and Mod(a.��¼����, 10) = ��¼����)" & _
            "       ) Q1" & vbNewLine & _
            " Where Not Exists (Select 1 From ҩƷ�շ���¼ Where ����id = Q1.Id)" & vbNewLine & _
            " Group by ҽ��ID,�շ�ϸĿID  Having Max(ID)<>0 )"
        Else
            '     And A.��������=0 :61879,����������ȷ��,��������������ֻ��0-��������
            strTableNo = strTableNo & " Union ALL  " & _
             " Select Max(ID) as ����ID,decode(sign(Sum(����)),-1,0,Sum(����)) as ׼���� " & _
             " From (  Select decode(J.��¼״̬,2,0,J.ID) as ID,J.ҽ����� as ҽ��ID,J.�շ�ϸĿID,nvl(J.����,1)*nvl(J.����,1) as ���� " & _
             "              From  ������� J,����ҽ����¼ M " & _
             "              Where  J.ҽ�����=M.ID  " & _
             "                      And Exists(Select 1 From   ����ҽ������ where ҽ��ID=J.ҽ����� and  Nvl( ִ��״̬, 0) <> 1 And No =J.NO  ) " & _
            "                       And Exists(Select 1 From   ����ҽ���Ƽ� A Where   A.ҽ��ID=J.ҽ����� and A.�շ�ϸĿID=J.�շ�ϸĿID And A.��������=0  And  Nvl( A.�շѷ�ʽ, 0) =0 ) " & _
             "                      And J.��¼״̬ in (1,2,3) and J.�۸񸸺� is null   " & _
             "                      And Instr('5,6,7', j.�շ����) = 0 And  Not Exists  (Select 1  From ��������  Where ����id = j.�շ�ϸĿid And Nvl(��������, 0) = 1)  " & _
             "                      And  instr(',C,D,F,G,K,',','||M.�������||',')=0 " & _
             "              Union all  " & _
             "              Select j.id, A.ҽ��ID,a.�շ�ϸĿID,-1*nvl(a.����,1)*nvl(C.��������,1) as ���� " & _
             "              From ����ҽ���Ƽ� A,����ҽ������ B,����ҽ��ִ�� C,������� J,����ҽ����¼ M " & _
             "              Where  A.ҽ��ID=b.ҽ��id And A.��������=0  and  Nvl( A.�շѷ�ʽ, 0) =0  and b.ҽ��id=c.ҽ��id and b.���ͺ�=c.���ͺ� And a.ҽ��id=M.ID " & _
             "                      And Nvl(C.ִ�н��, 1) =1 And Nvl(b.ִ��״̬, 0) <> 1 And B.NO=J.No and B.��¼����=1 " & _
             "                      And a.ҽ��id=J.ҽ����� and a.�շ�ϸĿid=j.�շ�ϸĿid  " & _
             "                      And  J.��¼״̬ in (1,3) and J.�۸񸸺� is null   " & _
             "                      And Instr('5,6,7', j.�շ����) = 0 And  Not Exists  (Select 1  From ��������  Where ����id = j.�շ�ϸĿid And Nvl(��������, 0) = 1)  " & _
             "                      And  instr(',C,D,F,G,K,',','||M.�������||',')=0  " & _
              "       ) " & _
             " group by ҽ��ID,�շ�ϸĿID  Having Max(ID) <>0)"
        End If
        '���ŵ��ݻ��ܽ��(��ϸ���շ�ϸĿ)
        'ִ��״̬Ӧ����ԭʼ��¼���ж�(������ҩ�Ҳ����˷ѵļ�¼)
        '������������ʱ"��¼״̬,���"�ظ�,AVG������,����Ҫ��"ִ��״̬"

        
   '58077:��Ҫ�ſ�ҽ���ƻ��в�Ϊ������ȡ�ķ���:
        '   0-������ȡ��1-�����Թܷ��ã�2-һ�η���ֻ��ȡһ�Σ�3-����ֻ��ȡһ�Σ�4-����δִ����ȡһ�Σ�5-����ֻ��ȡһ�Σ��ų�������Ŀ��6-����δִ����ȡһ�Σ��ų�������Ŀ��7-ÿ���״β���ȡ
        '
        Dim strSQLIn As String
        
        strSQLIn = "" & _
            "  Select NO,Nvl(�۸񸸺�,���) as ��� From �������  " & _
            "  Where ��¼����=1 And ��¼״̬ IN( 1,3)  And Nvl(ִ��״̬,0)<>1     " & _
            "   Minus " & _
            "  Select NO,Nvl(�۸񸸺�,���) as ��� " & _
            "  From ������� A1,����ҽ���Ƽ� B1 " & _
            "  Where A1.ҽ�����=B1.ҽ��id And A1.�շ�ϸĿID=B1.�շ�ϸĿID And B1.��������=0  And Nvl( B1.�շѷ�ʽ, 0) <>0  " & _
            "           And A1.��¼����=1 And A1.��¼״̬ IN(1,3)  And Nvl(A1.ִ��״̬,0)=2 " & _
            "           And Instr('5,6,7', a1.�շ����) = 0 And  Not Exists  (Select 1  From ��������  Where ����id = a1.�շ�ϸĿid And Nvl(��������, 0) = 1)  " & _
            "           And Not Exists (Select 1 From ҩƷ�շ���¼ Where ����id =a1.Id) "
        
        
        strSQL = _
        " Select A.NO,A.��¼״̬,A.ִ��״̬,Nvl(A.�۸񸸺�,A.���) as ���,A.��������," & _
        "       A.�ѱ�,C.���� as �����,C.���� as �����,A.�շ�ϸĿID,B.����,B.����,B.���,Nvl(A.��������,B.��������) ��������," & _
        "       A.���㵥λ,Max(A.ҽ�����) as ҽ�����, " & _
        "       Avg(Nvl(A.����,1)) as ����,Avg(A.����/A.����ϵ��) as ����," & _
        "       Sum(A.��׼����*A.����ϵ��) as ����," & _
        "       Sum(A.Ӧ�ս��) as Ӧ�ս��,Sum(A.ʵ�ս��) as ʵ�ս��, " & _
        "       D.���� as ִ�п���,A.ִ�в���ID,E.���� as ��������" & _
        " From ������� A,�շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D,���ű� E" & _
        " Where A.�շ�ϸĿID=B.ID And C.����=A.�շ����" & _
        "       And A.ִ�в���ID=D.ID(+) And A.��������ID=E.ID(+)" & _
        "       And (A.NO,Nvl(A.�۸񸸺�,A.���)) IN( " & strSQLIn & ")  " & _
        "       And A.NO IN( Select NO From ������� where  ��¼����=1 and ��¼״̬ in (1,3) )" & _
        " Group by A.NO,A.��¼״̬,A.ִ��״̬,Nvl(A.�۸񸸺�,A.���),A.�ѱ�,A.��������," & _
        "       C.����,C.����,A.�շ�ϸĿID,B.����,B.����,B.���,Nvl(A.��������,B.��������),A.���㵥λ," & _
        "       D.����,A.ִ�в���ID,E.����,A.ҩƷID "
        
        '��������
        '��"׼������=ԭʼ����"ʱ,�����ű���
        '�ſ��Ѿ�ȫ���˷ѵ���(ִ��״̬=0��һ�ֿ���)
        '��ʣ��������׼�������������������
            '1.�޶�Ӧ���շ���¼(����ͨ���û򲻸������õ�����),��ʱӦ��ʣ������
            '2.�շ���¼����ȫ������,����ȫ��ִ��,SQL���ų����ּ�¼
        strSQL = strTableNo & vbCrLf & _
        " Select A.NO,A.���,A.��������,A.�ѱ�,A.�����,A.�����,A.�շ�ϸĿID,A.����,A.����,A.���,A.��������,A.���㵥λ, Max(A.ҽ�����) as ҽ�����," & _
        "       Decode(Sign(Nvl(C.׼������,Sum(A.����*A.����))-B.ԭʼ����),0,Avg(A.����),1) as ׼�˸���," & _
        "       Decode(Sign(Nvl(C.׼������,Sum(A.����*A.����))-B.ԭʼ����),0,Sum(A.����),Nvl(C.׼������,Sum(A.����*A.����))) as ׼������," & _
        "       Nvl(C.׼������,Sum(A.����*A.����)) as ׼������,Sum(A.����*A.����) as ʣ������," & _
        "       A.����,Sum(A.Ӧ�ս��) as ʣ��Ӧ��,Sum(A.ʵ�ս��) as ʣ��ʵ��,max(q1.��¼��־) as ��¼��־," & _
        "       A.ִ�п���,A.ִ�в���ID,A.��������,B.����Ա����,B.�Ǽ�ʱ��,B.����ID,Max(M.ҽ������) as ҽ������,b.ԭʼ����" & _
        " From (" & strSQL & ") A, ׼���� C,����ҽ����¼ M," & _
        "          ( Select  ID, NO,���, �շ�ϸĿID,Nvl( ����,0)/NVL(����ϵ��,1) as ԭʼ����,����Ա����,�Ǽ�ʱ��,����ID" & _
        "            From �������   " & _
        "            Where  ��¼״̬ IN(1,3) And Nvl( ���ӱ�־,0)<>9 And  �۸񸸺� is NULL )B, " & _
        "            ( Select NO,Max(��¼״̬) as ��¼��־ From �������  Where ��¼״̬ in (1,3) Group by NO) Q1" & _
        " Where A.NO=B.NO And A.���=B.��� And A.�շ�ϸĿID=B.�շ�ϸĿID+0  And B.ID=C.����ID(+)" & _
        "            and A.ҽ�����=M.ID(+) and A.NO=q1.NO(+) " & _
        " Group by A.NO,A.���,A.��������,A.�ѱ�,A.�����,A.�����,A.�շ�ϸĿID,A.����,A.����,A.���,A.��������," & _
        "       A.���㵥λ,A.����,B.ԭʼ����,C.׼������,A.ִ�п���,A.ִ�в���ID,A.��������,B.����Ա����,B.�Ǽ�ʱ��,B.����ID" & _
        " Having Sum(A.����*A.����)<>0"
            
        strSQL = _
        " Select /*+ rule */  A.NO,A.���,A.��������,A.�ѱ�,A.�����,A.�����,A.����,Nvl(B.����,A.����) as ����,E1.���� as ��Ʒ��," & _
        "       A.���,A.��������,A.���㵥λ,A.�շ�ϸĿID,A.׼�˸��� as ����,A.׼������ as ����,A.����, A.ҽ����� ," & _
        "       A.ʣ��Ӧ��*(A.׼������/A.ʣ������) as Ӧ�ս��," & _
        "       A.ʣ��ʵ��*(A.׼������/A.ʣ������) as ʵ�ս��," & _
        "       A.ִ�п���,A.ִ�в���ID,A.��������,A.����Ա����,A.�Ǽ�ʱ��,A.����ID,A.ҽ������,A.��¼��־, " & _
        "       A.ԭʼ����,A.׼������,A.ʣ������" & _
        " From (" & strSQL & ") A,�շ���Ŀ���� B,�շ���Ŀ���� E1" & _
        " Where     A.�շ�ϸĿID=B.�շ�ϸĿID(+) And B.����(+)=1 And B.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
        "       And A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3 " & _
        " Order by A.NO,A.���"
    Else
        '��ȡ����ԭʼ����
        intSign = IIf(mstrDelTime <> "", -1, 1) '����,�����������
        
        strSQL = "" & _
        " Select A.NO " & _
        " From " & IIf(mblnNOMoved, "H", "") & "������ü�¼ A,Table(f_Str2list([2])) J" & _
        " Where A.��¼����=1 And A.��¼״̬ IN(1,3) And A.NO=J.Column_Value"
        
        strSQL = _
        " Select A.����ID,A.NO,Nvl(A.�۸񸸺�,A.���) as ���,A.��������,A.�ѱ�," & _
        "        A.�շ�ϸĿID,C.���� as �����,C.���� as �����,B.����,B.����,B.���,Nvl(A.��������,B.��������) ��������," & _
                IIf(gblnҩ����λ, "Decode(X.ҩƷID,NULL,A.���㵥λ,X." & gstrҩ����λ & ")", "A.���㵥λ") & " as ���㵥λ," & _
        "       Max(A.ҽ�����) as ҽ�����,Avg(Nvl(A.����,1)) as ����," & _
        "       Avg(" & intSign & "*A.����" & IIf(gblnҩ����λ, "/Nvl(X." & gstrҩ����װ & ",1)", "") & ") as ����," & _
        "       Sum(A.��׼����" & IIf(gblnҩ����λ, "*Nvl(X." & gstrҩ����װ & ",1)", "") & ") as ����," & _
        "       Sum(" & intSign & "*A.Ӧ�ս��) as Ӧ�ս��,Sum(" & intSign & "*A.ʵ�ս��) as ʵ�ս��," & _
        "       D.���� as ִ�п���,A.ִ�в���ID,E.���� as ��������,A.����Ա����,A.�Ǽ�ʱ��,Max(A.ժҪ) as ժҪ" & _
        " From " & IIf(mblnNOMoved, "H", "") & "������ü�¼ A,�շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D,���ű� E,ҩƷ��� X" & _
        " Where A.�շ�ϸĿID=B.ID And C.����=A.�շ���� And A.�շ�ϸĿID=X.ҩƷID(+)" & _
        "       And A.ִ�в���ID=D.ID(+) And A.��������ID=E.ID(+) And A.��¼����=1" & _
        "       And A.��¼״̬" & IIf(mstrDelTime <> "", "=2", " IN(1,3)") & " And A.NO IN(" & strSQL & ")" & _
                IIf(mstrDelTime <> "", " And A.�Ǽ�ʱ��=[1]", "") & _
                IIf(Not gblnShowErr, " And Nvl(A.���ӱ�־,0)<>9", "") & _
        " Group by A.����ID,A.NO,Nvl(A.�۸񸸺�,A.���),A.��������,A.�ѱ�,A.�շ�ϸĿID,C.����,C.����,B.����,B.����," & _
        "       B.���,Nvl(A.��������,B.��������),A.���㵥λ,D.����,A.ִ�в���ID,E.����,X.ҩƷID,X." & gstrҩ����λ & ",A.����Ա����,A.�Ǽ�ʱ��"
            
        strSQL = "Select /*+ rule */ " & _
            "       A.����ID,A.NO,A.���,A.��������,A.�ѱ�,A.�����,A.�����,A.����,Nvl(B.����,A.����) as ����,E1.���� as ��Ʒ��,A.���,A.��������," & _
            "       A.���㵥λ,A.ҽ����� ,A.�շ�ϸĿID,A.����,A.����,A.����,A.Ӧ�ս��,A.ʵ�ս��,A.ִ�п���,A.ִ�в���ID,A.��������,A.����Ա����,A.�Ǽ�ʱ��,A.ժҪ,M.ҽ������, " & _
            "       1 as ��¼��־,0 as ԭʼ����,0 as ׼������,0 as ʣ������" & _
            " From (" & strSQL & ") A,�շ���Ŀ���� B,�շ���Ŀ���� E1,����ҽ����¼ M" & _
            " Where     A.�շ�ϸĿID=B.�շ�ϸĿID(+) And B.����(+)=1 And B.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
            "       And A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3 " & _
            "       And A.ҽ�����=M.ID(+) " & _
            " Order by A.NO,A.���"
    End If
    
    If mstrDelTime <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(mstrDelTime), Replace(strNos, "'", ""))
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate("1991-01-01"), Replace(strNos, "'", ""))
    End If
    
    Call LoadInvoiceData(Replace(strNos, "'", ""))
    
    strҽ����� = ""
    If rsTmp.EOF Then
        Screen.MousePointer = 0
        MsgBox "û���ҵ������""" & strNo & """��صĿ����˷ѵļ�¼��" & _
            vbCrLf & "��Щ�շѼ�¼�����Ѿ��˷ѻ��Ѿ���ȫִ�С�", vbInformation, gstrSysName
        Call ClearFace(False)
        Exit Function
    End If
    
    mstrNOsOverFlow = ""
    If mbytMode = 1 Then
        strTmp = ""
        For i = 0 To UBound(Split(strNos, ","))
            strTmp = Replace(Split(strNos, ",")(i), "'", "")
            '����Ƿ��������
            If Not BillOperCheck(2, rsTmp!����Ա����, rsTmp!�Ǽ�ʱ��, "�˷�", strTmp, , 1, True) Then
                mstrNOsOverFlow = mstrNOsOverFlow & " ," & strTmp
            End If
        Next
        If mstrNOsOverFlow <> "" Then mstrNOsOverFlow = Mid(mstrNOsOverFlow, 2)
        If mBillDelType = EM_����ȫ�� And mstrNOsOverFlow <> "" Then
            Screen.MousePointer = 0
            MsgBox "���ŵ���ʹ��һ��ͨģʽ��ҽ���˷�Ҫ�������ˣ����������˷ѣ�", vbInformation, gstrSysName
            Call ClearFace(False)
            Exit Function
        End If
    End If
    
    If mbytMode = 0 Or mbytMode = 2 Then
        pic�˷�ժҪ.Enabled = False
        txt�˷�ժҪ.Text = Nvl(rsTmp!ժҪ)
    End If
    
    mCurBillType.bln���Ų����˷� = False
    With vsBill
        .Redraw = flexRDNone
        .Rows = .FixedRows + rsTmp.RecordCount
        mstrNOs = ""
        For i = 1 To rsTmp.RecordCount
            '����:29201
            .Cell(flexcpData, i, .ColIndex("��Ŀ")) = Nvl(rsTmp!��������)
            '����:33634
            .Cell(flexcpData, i, .ColIndex("����ID")) = Nvl(rsTmp!ҽ�����) & "," & Nvl(rsTmp!�շ�ϸĿID)
            If mbytMode = 1 Then
                If Val(Nvl(rsTmp!ҽ�����)) <> 0 And InStr(strҽ����� & ",", "," & Nvl(rsTmp!ҽ�����) & ",") = 0 Then
                    strҽ����� = strҽ����� & "," & Nvl(rsTmp!ҽ�����)
                End If
            End If
            strTemp = ""
            If Val(Nvl(rsTmp!��������)) <> 0 Then
                rsTmp.MoveNext
                strTemp = "��"
                If rsTmp.EOF Then
                    strTemp = "��"
                ElseIf Val(.Cell(flexcpData, i, .ColIndex("��Ŀ"))) <> Nvl(rsTmp!��������) Then
                    strTemp = "��"
                End If
                rsTmp.MovePrevious
                strTemp = "  " & strTemp & " "
            End If
    
            .RowData(i) = CLng(rsTmp!���)
            .TextMatrix(i, .ColIndex("ѡ��")) = 0
            
            .TextMatrix(i, .ColIndex("���ݺ�")) = rsTmp!NO
            .TextMatrix(i, .ColIndex("���")) = rsTmp!�����
            .TextMatrix(i, .ColIndex("��Ŀ")) = strTemp & rsTmp!���� & IIf(IsNull(rsTmp!���), "", " " & rsTmp!���)
            .TextMatrix(i, .ColIndex("��Ʒ��")) = strTemp & Nvl(rsTmp!��Ʒ��)
            .TextMatrix(i, .ColIndex("����")) = FormatEx(Nvl(rsTmp!����, 1) * rsTmp!����, 5)
            .TextMatrix(i, .ColIndex("��λ")) = Nvl(rsTmp!���㵥λ)
            .TextMatrix(i, .ColIndex("����")) = Format(rsTmp!����, gstrFeePrecisionFmt)
            .TextMatrix(i, .ColIndex("Ӧ�ս��")) = Format(rsTmp!Ӧ�ս��, gstrDec)
            .TextMatrix(i, .ColIndex("ʵ�ս��")) = Format(rsTmp!ʵ�ս��, gstrDec)
            .TextMatrix(i, .ColIndex("��������")) = Nvl(rsTmp!��������)
            .TextMatrix(i, .ColIndex("ִ�п���")) = Nvl(rsTmp!ִ�п���)
            .TextMatrix(i, .ColIndex("����Ա")) = rsTmp!����Ա����
            .TextMatrix(i, .ColIndex("ʱ��")) = Format(rsTmp!�Ǽ�ʱ��, "MM-dd HH:mm")
            .TextMatrix(i, .ColIndex("����ID")) = rsTmp!����ID
            .TextMatrix(i, .ColIndex("ҽ��")) = Nvl(rsTmp!ҽ������)
            .TextMatrix(i, .ColIndex("ԭʼ����")) = Nvl(rsTmp!ԭʼ����)
            .TextMatrix(i, .ColIndex("׼������")) = Nvl(rsTmp!׼������)
            .TextMatrix(i, .ColIndex("ҽ�����")) = Nvl(rsTmp!ҽ�����)
            .TextMatrix(i, .ColIndex("ִ�п���ID")) = Nvl(rsTmp!ִ�в���ID)
            If Not mCurBillType.bln���Ų����˷� Then mCurBillType.bln���Ų����˷� = RoundEx(Val(Nvl(rsTmp!ԭʼ����)), 7) <> RoundEx(Val(Nvl(rsTmp!׼������)), 7)
            .Cell(flexcpData, i, .ColIndex("ѡ��")) = Val(Nvl(rsTmp!��¼��־))    '�����ж��Ƿ����ʹ�,>1��ʾ������
            If Val(Nvl(rsTmp!��¼��־)) > 1 And InStr(1, mstrNOsPatiDel & ",", "," & rsTmp!NO & "") = 0 Then mstrNOsPatiDel = mstrNOsPatiDel & "," & rsTmp!NO
            If InStr(mstrNOs & ",", "," & rsTmp!NO & ",") = 0 Then
                '�����ָ���
                If mstrNOs <> "" Then
                    .Select i, .FixedCols, i, .COLS - 1
                    .CellBorder vbBlue, 0, 1, 0, 0, 0, 0
                End If
                mstrNOs = mstrNOs & "," & rsTmp!NO
            End If
            cur�ϼ� = cur�ϼ� + rsTmp!ʵ�ս��
            rsTmp.MoveNext
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
    
    Call SetpicInvoiceVisible   '���÷�Ʊ�ؼ�����ʾ
    If mbytMode = 1 Or mbytMode = 2 Then
        '--����:31179:��Ҫ���ҽ���������շѵĺ��˷ѵĴ���.Ϊ�˱�����ǰ��Ǹ�ݣ����û��ֱ����SQL������(�����ϻ������ֵİ����ݺŽ���������)
        rsTmp.Sort = "����ID,NO,���"
        If rsTmp.RecordCount <> 0 Then rsTmp.MoveFirst
        With rsTmp
                mstrNOs = ""
                Do While Not .EOF
                        If InStr(1, mstrNOs & ",", "," & Nvl(rsTmp!NO) & ",") = 0 Then
                            mstrNOs = mstrNOs & "," & Nvl(rsTmp!NO)
                            If Not mCurBillType.bln���Ų����˷� Then mCurBillType.bln���Ų����˷� = Not BillDeleteAll(Nvl(rsTmp!NO), 1, mblnHaveExcuteData)
                        End If
                        .MoveNext
                Loop
        End With
    End If
     If mstrNOs <> "" Then mstrNOs = Mid(mstrNOs, 2)
    txtAllTotal.Text = Format(cur�ϼ�, gstrDec)
    If mbytMode = 1 Then
        If strInvoiceNO <> "" And gTy_Module_Para.bytƱ�ݷ������ <> 0 Then
            If mBillDelType = EM_���Ų����� Or mBillDelType = EM_����ȫ�� Then
                'ֻ�е��Ų�����,�Ż���ڲ���ѡ������
                vsBill.Cell(flexcpChecked, 1, vsBill.ColIndex("ѡ��"), vsBill.Rows - 1, vsBill.ColIndex("ѡ��")) = 0
                Call FromInvoiceSelectNO(strInvoiceNO)
            End If
            If mBillDelType <> EM_����ȫ�� Then
                Call SelectRelatingInvoice(strInvoiceNO, True)
                '����ʾ����ѡ�ķ�Ʊ
                'Call OlnyShowSelectedInvoice
                Call ShowAndHideDelBillRow
            End If
        Else
            '78569,Ƚ����,2014-10-14,Ĭ�Ϲ�ѡ����
            If mBillDelType = EM_���Ų����� Or mBillDelType = EM_����ȫ�� Then
                With vsBill
                    For i = .FixedRows To .Rows - 1
                        If .TextMatrix(i, .ColIndex("���ݺ�")) = strOldNO Then
                            .Row = i: Exit For
                        End If
                    Next
                End With
                Call cmdBillSel_Click
            End If
        End If
        '40391
        If mBillDelType = EM_���Ų����� Or mBillDelType = EM_����ȫ�� Then
            Call LoadBalanceInfor
            Call LoadDelBalanceInfor
            Call ReCalcDelMoney
            Call FromNoSelectInvoice
        End If
        If mBillDelType = EM_����ȫ�� Then
            Call cmdSelAll_Click
        End If
        '78569,Ƚ����,2014-10-14,Ĭ�Ϲ�ѡ����
        If InStr(";" & mstrPrivs & ";", ";�����˷�;") = 0 Then Call cmdSelAll_Click
    Else
         Call cmdSelAll_Click
    End If
 
    If mbytMode = 1 Then
        cmdSelAll.Visible = mBillDelType <> EM_����ȫ��
        cmdClear.Visible = mBillDelType <> EM_����ȫ��
        cmdBillSel.Visible = mBillDelType <> EM_����ȫ��
    End If
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


Private Sub CalcDelMoney()
'���ܣ����ݵ�ǰ�����˷�ѡ������������˿���������
    Dim cur���ݺϼ� As Currency, curѡ��ϼ� As Currency
    Dim cur�˷Ѻϼ� As Currency, cur����� As Currency, cur���ϼ� As Currency
    Dim bln��ȫ�˷� As Boolean, bln�ֽ���� As Boolean
    Dim curTotal As Currency, strNo As String
    Dim i As Long, j As Long, k As Long, blnԭ���� As Boolean
    Dim colAllReturn As Collection
    
    If mbytMode = 0 Then Exit Sub
    If mrsBalance Is Nothing Then Exit Sub
        
    Set mcolError = New Collection
    Set colAllReturn = New Collection
    
    If mBillDelType = EM_����ȫ�� Then
        '���ŵ���һ����,�����
        For i = 0 To UBound(Split(mstrNOs, ","))
            strNo = CStr(Split(mstrNOs, ",")(i))
            mcolError.Add 0, "_" & strNo
        Next
        If cbo�˿ʽ.ListIndex = -1 And cbo�˿ʽ.ListCount > 0 Then cbo�˿ʽ.ListIndex = 0
        cbo�˿ʽ.Enabled = False
        cbo�˿ʽ.Locked = True
        
        curTotal = 0
        ''������ҽ����Ԥ������
         '����:1-Ԥ���,2-ҽ��,3-ҽ�ƿ�,4-���㿨,5-һ��ͨ,0-������
         mrsBalance.Filter = "����<>1 And ����<>2"
        With mrsBalance
             If .RecordCount <> 0 Then .MoveFirst
             Do While Not .EOF
                Select Case Val(Nvl(!����))
                Case 3, 4, 5    '3-ҽ�ƿ�,4-���㿨,5-һ��ͨ
                Case Else
                    curTotal = curTotal + !������
                End Select
                 .MoveNext
            Loop
            .Filter = 0
        End With
        txt�˿���.Text = curTotal
        Exit Sub
    End If
    
    '1.���ж������Ƿ���ԭ����,�Ծ����Ƿ���ý��㷽ʽѡ��,�Լ��ֱ���������
    blnԭ���� = True
    For i = 0 To UBound(Split(mstrNOs, ","))
        strNo = CStr(Split(mstrNOs, ",")(i))
        cur���ݺϼ� = 0: curѡ��ϼ� = 0
        With vsBill
            k = vsBill.FindRow(strNo, , .ColIndex("���ݺ�"))
            For j = k To vsBill.Rows - 1
                If vsBill.TextMatrix(j, .ColIndex("���ݺ�")) <> strNo Then Exit For
                cur���ݺϼ� = cur���ݺϼ� + Val(vsBill.TextMatrix(j, .ColIndex("ʵ�ս��")))
                If Val(vsBill.TextMatrix(j, .ColIndex("ѡ��"))) <> 0 Then
                    curѡ��ϼ� = curѡ��ϼ� + Val(vsBill.TextMatrix(j, .ColIndex("ʵ�ս��")))
                End If
            Next
        End With
        bln��ȫ�˷� = Not BillExistDelete(strNo, 1) And BillDeleteAll(strNo, 1, mblnHaveExcuteData) And (cur���ݺϼ� = curѡ��ϼ�)
        colAllReturn.Add Array(IIf(bln��ȫ�˷�, 1, 0), strNo, cur���ݺϼ�, curѡ��ϼ�), "_" & strNo   '�������ں�����ж�
        If Not bln��ȫ�˷� Then blnԭ���� = False
        
        If bln��ȫ�˷� Then
            mrsBalance.Filter = "NO='" & strNo & "'"
            With mrsBalance
                If .RecordCount <> 0 Then .MoveFirst
                Do While Not .EOF
                    If Val(Nvl(!����)) = 2 Then 'ҽ��
                        If Not mblnYB�������� Then blnԭ���� = False
                        If mblnYB�������� Then
                            If Not gclsInsure.GetCapability(support�����������, mlng����ID, mintInsure, !���㷽ʽ) Then
                               blnԭ���� = False
                            End If
                        End If
                    ElseIf InStr("3,4,5", Val(Nvl(!����))) > 0 Then
                        'һ��ͨ���
                        'If Nvl(!�Ƿ�����) = 1 Then blnԭ���� = False
                    End If
                    .MoveNext
                Loop
            End With
        End If
    Next
    
    '�շ�ʱȫ����Ԥ��(���㷽ʽΪ��),�˷�ʱ,������ָ���˷ѷ�ʽ
    '����:1-Ԥ���,2-ҽ��,3-ҽ�ƿ�(һ��ͨ),4-���㿨,5-һ��ͨ,0-������
    mrsBalance.Filter = "����<>1"
    If mrsBalance.RecordCount = 0 Then blnԭ���� = True
    mrsBalance.Filter = 0
    If mBillDelType = EM_����ȫ�� Then blnԭ���� = True
    
    If blnԭ���� Then
        zlControl.CboSetIndex cbo�˿ʽ.hWnd, mintReturnMode
    End If
    
    cbo�˿ʽ.Enabled = Not blnԭ����
    cbo�˿ʽ.Locked = blnԭ����
    
    '2.�����˿�����
    If cbo�˿ʽ.ListIndex <> -1 Then
        If cbo�˿ʽ.ItemData(cbo�˿ʽ.ListIndex) = 1 Then
            bln�ֽ���� = True
        End If
    End If
    Dim varTemp As Variant
    
    For i = 0 To colAllReturn.Count     ' UBound(Split(mstrNOs, ","))
        '0-�Ƿ���ȫ�˷�;1-NO,2-���ݺϼ�,3-ѡ��ϼ�
        varTemp = colAllReturn(i)
        strNo = varTemp(1)
        cur���ݺϼ� = Val(varTemp(2)): curѡ��ϼ� = Val(varTemp(3))
        cur�˷Ѻϼ� = 0: cur����� = 0
        '��ȫ�˷�ʱ�ſ�ҽ�����㼰��Ԥ�����
        bln��ȫ�˷� = IIf(Val(varTemp(0)) = 1, True, False)
        If bln��ȫ�˷� Then
            mrsBalance.Filter = "NO='" & strNo & "'"
            '����:1-Ԥ���,2-ҽ��,3-ҽ�ƿ�,4-���㿨,5-һ��ͨ,0-������
            With mrsBalance
                Do While Not .EOF
                    Select Case Val(Nvl(!����))
                    Case 1 'Ԥ����
                         curѡ��ϼ� = curѡ��ϼ� - Nvl(!������, 0)
                    Case 2 'ҽ��
                        '������ֽ��㷽ʽ��֧�ֻ���,Ҫ��Ϊ�ֽ�,���ü�ȥ
                        If mblnYB�������� Then
                            If gclsInsure.GetCapability(support�����������, mlng����ID, mintInsure, !���㷽ʽ) Then
                                curѡ��ϼ� = curѡ��ϼ� - Nvl(!������, 0)
                            End If
                        Else     '��֧�������������ʱ,ֻ���������Ϊ�ֽ�,����ԭ����,������ҽ������
                            If !���㷽ʽ <> mstr�����ʻ� Then
                                curѡ��ϼ� = curѡ��ϼ� - Nvl(!������, 0)
                            End If
                        End If
                    Case 3, 4 'ҽ�ƿ��ͽ��㿨
                        If Val(Nvl(!�Ƿ�����)) = 0 Then
                            curѡ��ϼ� = curѡ��ϼ� - Nvl(!������, 0)
                        End If
                    Case 5 'һ��ͨ
                            curѡ��ϼ� = curѡ��ϼ� - Nvl(!������, 0)
                    Case Else
                    End Select
                    .MoveNext
                Loop
            End With
        End If
        
        '���ý���λ��,���ֽ����ʱ����ֱ�
        If bln�ֽ���� Then
            If mintInsure > 0 Then
                If gclsInsure.GetCapability(support�ֱҴ���, mlng����ID, mintInsure) Then
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
        If Not blnԭ���� Then
            cur����� = cur�˷Ѻϼ� - curѡ��ϼ�
        End If
        
        curTotal = curTotal + cur�˷Ѻϼ�
        mcolError.Add cur�����, "_" & strNo
        cur���ϼ� = cur���ϼ� + cur�����
    Next
    
    txt�˿���.ToolTipText = "�˷������:" & Format(cur���ϼ�, gstrDec)
    txt�˿���.Text = Format(curTotal, "0.00")
    Call Show�˿ʽ(cbo�˿ʽ.Enabled)
    
End Sub

Private Sub Show�˿ʽ(ByVal blnVisible As Boolean)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ�˿ʽ
    '���:blnVisible-true,��ʾ,��������
    '����:���˺�
    '����:2012-09-12 11:27:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    cbo�˿ʽ.Visible = blnVisible
    lbl�˿ʽ.Visible = blnVisible
    lbl�˿���.Visible = blnVisible
    txt�˿���.Visible = blnVisible
End Sub


Private Function CheckOnCardValied(ByVal blnCur�����˷� As Boolean, ByVal lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���һ��ͨ�Ƿ�Ϸ�
    '����:�Ϸ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-31 12:00:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCardNo As String
    On Error GoTo errHandle
    If Not mblnOneCard Then CheckOnCardValied = True: Exit Function
    mrsBalance.Filter = "����ID=" & lng����ID & " And ����=5"
    If mrsBalance.RecordCount = 0 Then CheckOnCardValied = True: Exit Function
    If blnCur�����˷� Then
         MsgBox "��ǰ����ʹ����һ��ͨ����,���ܽ��в����˷ѣ�", vbInformation, gstrSysName
        Exit Function
    End If
    If mobjICCard Is Nothing Then
        On Error Resume Next
        Set mobjICCard = CreateObject("zlICCard.clsICCard")
        On Error GoTo 0
        If mobjICCard Is Nothing Then
            MsgBox "һ��ͨ�ӿڴ���ʧ��,���ܽ����˷�!����ӿ��ļ�.", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    With mrsBalance
        'mObjSquare.objSquareCard
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
        If mobjSquare.zlBrushCard(Me, mlngModule, Nothing, 0, False, _
          mtyPati.����, mtyPati.�Ա�, mtyPati.����, 0, strCardNo, "") = False Then Exit Function
        If strCardNo = "" Then Exit Function
        If strCardNo <> Nvl(!����) Then
            MsgBox "��ǰ������ۿ�Ų�һ��,���ܽ����˷�.", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    CheckOnCardValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckThreeSwapValied(ByVal blnCur�����˷� As Boolean, _
    ByVal lng����ID As Long, ByVal blnMulitNo As Boolean, Optional bln�쳣���� As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������׼��
    '���:blnCur�����˷�
    '        lng����ID
    '       blnMulitNo-�Ƿ�൥��
    '       bln�쳣����-�쳣���������շ�
    '����:
    '����:���ݺϷ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-31 15:35:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln���� As Boolean, intCol As Integer
    Dim str�������㷽ʽ As String
    On Error GoTo errHandle
    If cbo�˿ʽ.Visible Then
        mrsBalance.Filter = "����ID=" & lng����ID & " And ����>=3 and ����<=4  and ���㷽ʽ='" & zlStr.NeedName(cbo�˿ʽ.Text) & "'"
        With mrsBalance
            If .RecordCount <> 0 Then .MoveFirst
            Do While Not .EOF
                 str�������㷽ʽ = str�������㷽ʽ & "," & Nvl(!���㷽ʽ)
                .MoveNext
            Loop
        End With
        mrsBalance.Filter = 0
    End If
    mrsBalance.Filter = "����ID=" & lng����ID & " And ����>=3 and ����<=4 "
    If mrsBalance.RecordCount = 0 Then CheckThreeSwapValied = True: Exit Function
    With mrsBalance
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            bln���� = False
            If Val(Nvl(!�Ƿ�����)) = 1 Then
                '�Ŷӽ��ֽ�
                For intCol = 1 To vsBalance.COLS - 1 Step 2
                    If vsBalance.Cell(flexcpData, 1, intCol) = !���㷽ʽ And _
                        (Val(vsBalance.TextMatrix(1, intCol + 1)) = 0) Or vsBalance.RowHidden(1) Then
                         bln���� = True
                    End If
                Next
            End If
        
            If blnCur�����˷� And Val(Nvl(!�Ƿ�ȫ��)) = 1 And blnMulitNo Then
                If Val(Nvl(!�Ƿ�����)) <> 1 Then
                    MsgBox "��ǰ����ʹ���˵��������㽻��,���е��ݱ���ȫ�ˣ�", vbInformation, gstrSysName
                    Exit Function
                ElseIf bln���� = False Then
                    MsgBox "��ǰ����ʹ���˵��������㽻��,���е��ݱ���ȫ�ˣ�", vbInformation, gstrSysName
                    Exit Function
                ElseIf InStr(str�������㷽ʽ & ",", "," & zlStr.NeedName(cbo�˿ʽ.Text) & ",") > 0 Then
                    MsgBox "��ǰ����ʹ���˵��������㽻��,���е��ݱ���ȫ�ˣ�", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
            
            If cbo�˿ʽ.Visible And cbo�˿ʽ.Enabled Then
                If bln���� And blnCur�����˷� And InStr(str�������㷽ʽ & ",", "," & zlStr.NeedName(cbo�˿ʽ.Text) & ",") > 0 Then
                    MsgBox "��ǰ����ʹ���˵��������㽻��,���ŵ��ݱ���ȫ�ˣ�", vbInformation, gstrSysName
                    Exit Function
                End If
                If InStr(str�������㷽ʽ & ",", "," & zlStr.NeedName(cbo�˿ʽ.Text) & ",") > 0 Then
                    MsgBox "�����˷�ʱ,������ѡ����������", vbInformation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            End If
            
            If Not bln���� And Not blnCur�����˷� Then
                '1.δ���ֽ�
                '2.���ǲ����˵�(����)
                If zlCheckDelValied(Val(Nvl(!�����ID)), Nvl(!����), Val(Nvl(!����)) = 4, Nvl(!����), Nvl(!������ˮ��), Nvl(!����˵��), lng����ID, Val(Nvl(!������)), bln�쳣����) = False Then Exit Function
                If Val(Nvl(!�Ƿ��˿��鿨)) = 1 And Val(Nvl(!����)) = 4 Then
                    '��Ҫ�鿨
                    If CheckBrushCard(Val(Nvl(!�����ID)), Val(Nvl(!����)) = 4, Val(Nvl(!������)), Nvl(!����), "", bln����) = False Then Exit Function
                End If
            End If
            .MoveNext
        Loop
    End With
    CheckThreeSwapValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function DelInsureMulitOneBalance(ByVal blnExistThreeSwap As Boolean, _
     ByVal arrNo As Variant, ByVal lng����ID As Long, ByVal strAllBalance As String, _
     ByVal strҽ������ As String, ByVal str�˽��㷽ʽ As String, ByVal bln���� As Boolean, _
    ByRef blnCommited As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�൥��һ�ν����˷�
    '���:arrNO-�����˵ĵ��ݺ�
    '       str�˽��㷽ʽ-�˵Ľ��㷽ʽ
    '       bln����-�˵Ľ��㷽ʽ�Ƿ��ֽ�
    '����:
    '����:�ɹ����ҽ����Ƕ൥��һ�ν���,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-31 23:38:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strAdvance As String, arrBalance As Variant, str���㷽ʽ As String
    Dim dbl������ As Double, dbl�ɷ���� As Double, dbl��� As Double
    Dim strBalance As String, dbl�˿�ϼ� As Double, str�˻ؽ��� As String
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim i As Long, k As Long, j As Long, cur����� As Double
    Dim lng����ID As Long
    
    On Error GoTo errHandle
    If mintInsure = 0 Then DelInsureMulitOneBalance = True: Exit Function
    If Not (MCPAR.�൥��һ�ν��� Or MCPAR.�൥�ݵ�һ�ν���) Then DelInsureMulitOneBalance = True: Exit Function
    
    strAdvance = strAllBalance
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
    If mintInsure <> 0 And mblnYB�������� Then
        If Not gclsInsure.ClinicDelSwap(lng����ID, , mintInsure, strAdvance) Then Exit Function
    Else
        strAdvance = ""
    End If
    
    If strAdvance = strAllBalance Or strAdvance = "" Then
        gcnOracle.CommitTrans: blnCommited = True
        If mblnYB�������� Then Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, True, mintInsure)
        DelInsureMulitOneBalance = True: Exit Function
    End If
    
    '���ݷ��صĽ�����Ϣ������Ԥ����¼��strAdvance���ظ�ʽ:���㷽ʽ1|���||���㷽ʽ2:���...
    '�ȷ�̯��ÿ�ŵ�����
    Set rsTmp = GetBalanceSet
    arrBalance = Split(strAdvance, "||")
    For i = 0 To UBound(arrBalance)
        str���㷽ʽ = Split(arrBalance(i), "|")(0)
        dbl������ = -1 * Val(Split(arrBalance(i), "|")(1))
        For k = 0 To UBound(arrNo)
            dbl�ɷ���� = Getʵ�ս��(arrNo(k))
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
    Next
    For k = 0 To UBound(arrNo)
        strBalance = "": cur����� = 0
        dbl��� = Getʵ�ս��(arrNo(k))
        rsTmp.Filter = "�������=" & k
        For i = 1 To rsTmp.RecordCount
            strBalance = IIf(strBalance = "", "", strBalance & "||") & rsTmp!���㷽ʽ & "|" & -1 * rsTmp!������
            dbl��� = dbl��� - rsTmp!������
            rsTmp.MoveNext
        Next
        '��Ϊָ���Ľ��㷽ʽ��������ֽ𣬿��ܲ����µ������
        dbl������ = dbl���
        If bln���� Then
            dbl������ = Format(CentMoney(dbl���), "0.00")
            cur����� = dbl������ - dbl���
        End If
        dbl�˿�ϼ� = dbl�˿�ϼ� + dbl������
        str�˻ؽ��� = str�˽��㷽ʽ & "|" & -1 * dbl������ & "| "
        lng����ID = GetDelBalanceID(arrNo(k))
        If Not blnExistThreeSwap Then
            strSQL = "zl_�����շѽ���_Update(" & lng����ID & ",'" & str�˻ؽ��� & "',0,'" & _
                strBalance & "'," & -1 * cur����� & ",NULL,NULL,NULL,1)"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        Else
            'Zl_ҽ������У��_Update
             strSQL = "Zl_ҽ������У��_Update("
             '  ����id_In   ������ü�¼.����id%Type,
             strSQL = strSQL & "" & lng����ID & ","
             '  ���ս���_In Varchar2
             strSQL = strSQL & "'" & strBalance & "')"
             Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        End If
    Next
    gcnOracle.CommitTrans: blnCommited = True
    Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, True, mintInsure)
    If Not blnExistThreeSwap Then
        If Not (strAdvance = strAllBalance Or strAdvance = "") Then
            MsgBox "Ӧ�˽��" & vbCrLf & str�˽��㷽ʽ & "��" & Format(dbl�˿�ϼ�, "0.00") & "Ԫ", vbInformation + vbOKOnly, gstrSysName
        End If
    End If
    DelInsureMulitOneBalance = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans: blnCommited = True
    Call ErrCenter
    Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, False, mintInsure)
End Function
Private Function DelInsureOneBill(ByVal strҽ������ As String, ByVal blnExistThreeSwap As Boolean, _
     ByVal lng����ID As Long, _
     ByVal lngPage As Long, ByVal lngPages As Long, _
     ByRef blnCommited As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ҽ���ӿ��˷�
    '���:blnExistThreeSwap-���ڵ������ӿ�
    '        lng����ID-����ID
    '       lngPage(lngPages)-��ǰҳ(��ǰҳ��)
    '����:
    '����:�˷ѳɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-09-01 01:08:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strAdvance As String
    Dim blnTransMedicare As Boolean
    On Error GoTo errHandle
    blnTransMedicare = False
    If mintInsure = 0 Then DelInsureOneBill = True: Exit Function
    If Not mblnYB�������� Or lng����ID = 0 Then DelInsureOneBill = True: Exit Function
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
    strAdvance = lngPages & "|" & lngPage
    'strAdvance = CStr(UBound(arrNO) + 1) & "|" & CStr(UBound(arrNO) + 1 - i)
    If Not gclsInsure.ClinicDelSwap(lng����ID, , mintInsure, strAdvance) Then GoTo errHandle:
    blnTransMedicare = True
    gcnOracle.CommitTrans: blnCommited = True
    Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, True, mintInsure)
    DelInsureOneBill = True
    Exit Function
errHandle:
    '50134
    gcnOracle.RollbackTrans: blnCommited = True
    If blnTransMedicare Then Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, False, mintInsure)
    If Err.Number <> 0 Then
        Call ErrCenter
    End If
 End Function
Private Function DelOneCardPay(ByVal varNO As Variant, _
     ByRef blnCommited As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:һ��ͨ�˷�
    '����:
    '����:���˺�
    '����:2011-09-01 01:47:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCardNo As String, strSwap As String, strHsptCode As String 'ҽԺ����
    Dim i As Long, dblMoney As Double, strNos As String, strSQL As String

    If mblnOneCard = False Then DelOneCardPay = True: Exit Function
    mrsBalance.Filter = "����=5"
    If mrsBalance.RecordCount = 0 Then
        mrsBalance.Filter = 0
        DelOneCardPay = True: Exit Function
    End If
    For i = 0 To UBound(varNO)
        strNos = strNos & "," & varNO(i)
    Next
    If strNos <> "" Then strNos = Mid(strNos, 2)
    With mrsBalance
        .MoveFirst
        dblMoney = 0
        Do While Not .EOF
            If InStr(1, "," & strNos & ",", "," & Nvl(!NO) & ",") > 0 Then
                strCardNo = Nvl(!����): strSwap = Nvl(!������ˮ��): strHsptCode = Nvl(!ҽԺ����)
                dblMoney = dblMoney + Nvl(!������)
            End If
            .MoveNext
        Loop
    End With
    mrsBalance.Filter = 0
    If dblMoney = 0 Then DelOneCardPay = True: Exit Function
    On Error GoTo errHandle
    If Not mobjICCard.ReturnSwap(strCardNo, strHsptCode, strSwap, dblMoney) Then
        gcnOracle.RollbackTrans
        MsgBox "һ��ͨ�˷ѽ��׵���ʧ�ܣ��˷Ѳ���ʧ�ܣ�", vbExclamation, gstrSysName
        Exit Function
    End If
    ' Zl_�����շ�_���У��
    strSQL = "Zl_�����շ�_���У��("
    '  No_In       Varchar2,
    strSQL = strSQL & "'" & strNos & "',"
    '  ��������_In Number,
    '  --��������_In:0-һ��ͨ;1-���ѿ�;2-ҽ�ƿ�
    strSQL = strSQL & "0,"
    '  �����id_In ����Ԥ����¼.�����id%Type,
    strSQL = strSQL & "NULL,"
    '  ����_In     ����Ԥ����¼.����%Type
    strSQL = strSQL & "'" & strCardNo & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    blnCommited = True: gcnOracle.CommitTrans
    DelOneCardPay = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    Call ErrCenter
    MsgBox "һ��ͨ�˷ѽ��׵���ʧ�ܣ��˷Ѳ���ʧ�ܣ�", vbExclamation, gstrSysName
End Function

Private Function DelThreeSwapFeeSingle(ByVal varNO As Variant, colThreeBalance As Collection, _
    colOrder As Collection, ByVal str����IDs As String, ByRef blnCommited As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�˵������ӿڽ���,�˷ѳɹ�,����true,���򷵻�False
    '����:
    '       varNO - ���ν��㵥�ݺ�
    '       colThreeBalance - �����˷���Ϣ��NO|�˷ѽ��
    '       colOrder - �˷�����Ϣ�����
    '����:���������˷ѳɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-09-01 02:45:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, dblMoney As Double
    Dim i As Long, strNos As String, strSQL As String
    Dim strSelNos As String, str����IDs As String, strErrMsg As String
    Dim cllUpdate As Collection, cllThreeSwap As Collection
     
    On Error GoTo errHandle
    If Left(str����IDs, 1) = "," Then str����IDs = Mid(str����IDs, 2)
    
    For i = 0 To UBound(varNO)
        strNos = strNos & "," & varNO(i)
        If colOrder("_" & varNO(i)) <> "δѡ��" Then
            strSelNos = strSelNos & "," & varNO(i)
            mrsBalance.Filter = "NO='" & varNO(i) & "'"
            If mrsBalance.RecordCount <> 0 Then
                mrsBalance.MoveFirst
                str����IDs = str����IDs & "," & Nvl(mrsBalance!����ID)
            End If
        End If
        If colThreeBalance("_" & varNO(i)) <> "" Then
            '�����ܽ��
            dblMoney = dblMoney + Val(Split(colThreeBalance("_" & varNO(i)) & "|", "|")(1))
        End If
    Next
    If strNos <> "" Then strNos = Mid(strNos, 2)
    
    mrsBalance.Filter = "����=3"
    With mrsBalance
        If .RecordCount <> 0 Then
            .MoveFirst
            varData = Array(0, 0, "", "", "", 0, "")
            varData(0) = Val(Nvl(!����))
            varData(1) = Val(Nvl(!�����ID))
            varData(2) = Nvl(!����)
            varData(3) = Nvl(!������ˮ��)
            varData(4) = Nvl(!����˵��)
            varData(5) = dblMoney
        End If
    End With
    mrsBalance.Filter = 0
    
    If strSelNos = "" Or RoundEx(dblMoney, 5) = 0 Then DelThreeSwapFeeSingle = True: Exit Function
    strSelNos = Mid(strSelNos, 2)
    If str����IDs = "" Then str����IDs = ",0"
    str����IDs = Mid(str����IDs, 2)
    'varData = Array(Val(Nvl(!����)), Val(Nvl(!�����ID)), _
                    CStr(Nvl(!����)), CStr(Nvl(!������ˮ��)), CStr(Nvl(!����˵��)), dblMoney)
    ' Zl_�����շ�_���У��
    strSQL = "Zl_�����շ�_���У��("
    '  No_In       Varchar2,
    strSQL = strSQL & "'" & strSelNos & "',"
    '  ��������_In Number,
    '  --��������_In:0-һ��ͨ;1-���ѿ�;2-ҽ�ƿ�
    strSQL = strSQL & "2,"
    '  �����id_In ����Ԥ����¼.�����id%Type,
    strSQL = strSQL & "" & varData(1) & ","
    '  ����_In     ����Ԥ����¼.����%Type
    strSQL = strSQL & "'" & varData(2) & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    Set cllUpdate = New Collection: Set cllThreeSwap = New Collection
    If CallBackBalanceInterface(varData(1), False, varData(2), _
        varData(3), varData(4), str����IDs, str����IDs, varData(5), cllUpdate, cllThreeSwap, strErrMsg) = False Then
        gcnOracle.RollbackTrans: blnCommited = True
        If strErrMsg <> "" Then
            MsgBox strErrMsg, vbExclamation, gstrSysName
        Else
            MsgBox "�����˷ѽ��׵���ʧ�ܣ��˷Ѳ���ʧ�ܣ�", vbExclamation, gstrSysName
        End If
        Exit Function
    End If
    ' zlExecuteProcedureArrAy cllUpdate, Me.Caption
    gcnOracle.CommitTrans: blnCommited = True
    On Error GoTo Errhand:
    zlExecuteProcedureArrAy cllThreeSwap, Me.Caption
    
    DelThreeSwapFeeSingle = True
    Exit Function
errHandle:
    If Not blnCommited Then gcnOracle.RollbackTrans
    Call ErrCenter
    MsgBox "�����˷ѽ��׵���ʧ�ܣ�", vbExclamation, gstrSysName
    blnCommited = True
    Exit Function
Errhand:
    If Not blnCommited Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        gcnOracle.BeginTrans: blnCommited = False
    End If
End Function

Private Function DelThreeSwapFee(ByVal varNO As Variant, _
     ByRef blnCommited As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�˵������ӿڽ���,�˷ѳɹ�,����true,���򷵻�False
    '����:���������˷ѳɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-09-01 02:45:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCardNo As String, strSwap As String, strHsptCode As String 'ҽԺ����
    Dim varData As Variant, cllBlance As Collection, dblMoney As Double
    Dim i As Long, strNos As String, strSQL As String
    Dim strSelNos As String, str����ID As String, strErrMsg As String
    Dim cllUpdate As Collection, cllThreeSwap As Collection
    Dim dbl�˿�ϼ� As Double
    Dim bln��ָ�����  As Boolean   '��
     
    On Error GoTo errHandle
    Dim blnHaveData As Boolean
    
    DelThreeSwapFee = False
    mrsBalance.Filter = "����>=3 and ����<=4 and У�Ա�־<>2"
    If mrsBalance.RecordCount = 0 Then
        mrsBalance.Filter = 0
        DelThreeSwapFee = True: Exit Function
    End If
    '����
    bln��ָ����� = False
    Set cllBlance = New Collection
     If cbo�˿ʽ.Enabled And cbo�˿ʽ.Visible Then
        mrsBalance.Filter = "����>=3 and  ����<=4 and ���㷽ʽ='" & zlStr.NeedName(cbo�˿ʽ.Text) & "'"
        If mrsBalance.RecordCount <> 0 Then
            With mrsBalance
                bln��ָ����� = True
                varData = Array(0, 0, "", "", "", 0, "")
                varData(0) = Val(Nvl(!����)): varData(1) = Val(Nvl(!�����ID))
                varData(2) = Nvl(!����): varData(3) = Nvl(!������ˮ��)
                varData(4) = Nvl(!����˵��): varData(6) = Nvl(!���㷽ʽ)
                varData(5) = Val(txt�˿���.Text)
                cllBlance.Add varData
            End With
        Else
            '�˸�ָ���Ľ��㷽ʽ
            mrsBalance.Filter = "����>=3 and  ����<=4  "
            If mrsBalance.RecordCount = 0 Then
                mrsBalance.Filter = 0: DelThreeSwapFee = True: Exit Function
            End If
        End If
    End If
     
    mrsBalance.Filter = "����=3 or ����=4"
    dbl�˿�ϼ� = 0
    For i = 0 To UBound(varNO)
        strNos = strNos & "," & varNO(i)
    Next
    If strNos <> "" Then strNos = Mid(strNos, 2)
    
    With mrsBalance
        .Sort = "����,�����ID,����,������ˮ��,����˵��"
        .MoveFirst
        varData = Array(0, 0, "", "", "", 0, "")
        dblMoney = 0
        If .RecordCount <> 0 Then .MoveFirst
        
        Do While Not .EOF
            If InStr(1, "," & strNos & ",", "," & Nvl(!NO) & ",") > 0 Then
                If Not bln��ָ����� Then
                        If Not (varData(0) = Val(Nvl(!����)) And varData(1) = Val(Nvl(!�����ID)) _
                            And varData(2) = Nvl(!����) _
                            And varData(3) = Nvl(!������ˮ��) _
                            And varData(4) = Nvl(!����˵��)) Then
                            
                            If varData(0) <> 0 Then
                               ' varData = Array(Val(Nvl(!����)), Val(Nvl(!�����ID)), _
                                CStr(Nvl(!����)), CStr(Nvl(!������ˮ��)), CStr(Nvl(!����˵��)), dblMoney)
                                blnHaveData = False
                                For i = 1 To vsBalance.COLS - 1 Step 2
                                    If vsBalance.Cell(flexcpData, 1, i) = Nvl(!���㷽ʽ) _
                                        And vsBalance.RowHidden(1) = False _
                                        And Val(vsBalance.TextMatrix(1, i + 1)) <> 0 Then
                                        blnHaveData = True: Exit For
                                    End If
                                Next
                                If blnHaveData Then cllBlance.Add varData
                            End If
                            varData(5) = 0
                            varData(0) = Val(Nvl(!����)): varData(1) = Val(Nvl(!�����ID))
                            varData(2) = Nvl(!����): varData(3) = Nvl(!������ˮ��)
                            varData(4) = Nvl(!����˵��): varData(6) = Nvl(!���㷽ʽ)
                        End If
                        dblMoney = dblMoney + Nvl(!������)
                        varData(5) = varData(5) + Nvl(!������)
                End If
                If InStr(1, "," & strSelNos & ",", "," & Nvl(!NO) & ",") = 0 Then
                        strSelNos = strSelNos & "," & Nvl(!NO)
                        str����ID = str����ID & "," & Nvl(!����ID)
                End If
            End If
            .MoveNext
        Loop
    End With
    If varData(0) <> 0 And Not bln��ָ����� Then
        blnHaveData = False
        For i = 1 To vsBalance.COLS - 1 Step 2
            If vsBalance.Cell(flexcpData, 1, i) = varData(6) _
               And vsBalance.RowHidden(1) = False And Val(vsBalance.TextMatrix(1, i + 1)) <> 0 Then
                blnHaveData = True: Exit For
            End If
        Next
        If blnHaveData Then cllBlance.Add varData
    End If
    
    mrsBalance.Filter = 0
    If strSelNos = "" Or cllBlance.Count = 0 Then DelThreeSwapFee = True: Exit Function
    strSelNos = Mid(strSelNos, 2)
    If str����ID = "" Then str����ID = ",0"
    str����ID = Mid(str����ID, 2)
    For i = 1 To cllBlance.Count
      'varData = Array(Val(Nvl(!����)), Val(Nvl(!�����ID)), _
                        CStr(Nvl(!����)), CStr(Nvl(!������ˮ��)), CStr(Nvl(!����˵��)), dblMoney)
        ' Zl_�����շ�_���У��
        strSQL = "Zl_�����շ�_���У��("
        '  No_In       Varchar2,
        strSQL = strSQL & "'" & strSelNos & "',"
        '  ��������_In Number,
        '  --��������_In:0-һ��ͨ;1-���ѿ�;2-ҽ�ƿ�
        strSQL = strSQL & "" & IIf(cllBlance(i)(0) = 3, 2, 1) & ","
        '  �����id_In ����Ԥ����¼.�����id%Type,
        strSQL = strSQL & "" & cllBlance(i)(1) & ","
        '  ����_In     ����Ԥ����¼.����%Type
        strSQL = strSQL & "'" & cllBlance(i)(2) & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        Set cllUpdate = New Collection: Set cllThreeSwap = New Collection
        If CallBackBalanceInterface(cllBlance(i)(1), cllBlance(i)(0) = 4, cllBlance(i)(2), _
            cllBlance(i)(3), cllBlance(i)(4), str����ID, "", cllBlance(i)(5), cllUpdate, cllThreeSwap, strErrMsg) = False Then
            gcnOracle.RollbackTrans: blnCommited = True
            If strErrMsg <> "" Then
                    MsgBox strErrMsg, vbExclamation, gstrSysName
            Else
                   MsgBox "�����˷ѽ��׵���ʧ�ܣ��˷Ѳ���ʧ�ܣ�", vbExclamation, gstrSysName
            End If
            Exit Function
        End If
       ' zlExecuteProcedureArrAy cllUpdate, Me.Caption
        gcnOracle.CommitTrans: blnCommited = True
        On Error GoTo Errhand:
        zlExecuteProcedureArrAy cllThreeSwap, Me.Caption
        gcnOracle.BeginTrans: blnCommited = False
    Next
    gcnOracle.CommitTrans: blnCommited = True
    DelThreeSwapFee = True
    Exit Function
errHandle:
    If Not blnCommited Then gcnOracle.RollbackTrans
    Call ErrCenter
    MsgBox "�����˷ѽ��׵���ʧ�ܣ�", vbExclamation, gstrSysName
      blnCommited = True
    Exit Function
Errhand:
     If Not blnCommited Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        gcnOracle.BeginTrans: blnCommited = False
        Resume
    End If
    
End Function
Private Function ExecDelete() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ�ж൥���˷�
    '����:�˷ѳɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-30 18:52:56
    '˵��:��Ϊҽ����ԭ�򣬶��ŵ����˷�ʱ���ִ��ύ
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim colOrder As New Collection, colBalanceID As New Collection
    Dim colBalance As New Collection    '���ս��㷽ʽ,���
    Dim colThreeBalance As New Collection    '��������,���
    Dim colOtherBalance As New Collection    '��������,���
    Dim colSQL As New Collection
    Dim arrSQL As Variant, strSQL As String, strInvoices As String, strInvoice As String
    Dim blnCur�����˷� As Boolean, blnAll�����˷� As Boolean, blnTrans As Boolean, blnTransMedicare As Boolean, blnPrint As Boolean
    Dim strBalance As String, strAllBalance As String, strTmp As String, strAdvance As String
    Dim strNo As String, str��� As String, strDelNOs As String, strOtherNOs As String, strAllNOs As String
    Dim cur����� As Currency, DateDel As Date
    Dim i As Long, j As Long, k As Long, lngCount As Long, arrNo As Variant, lng����ID As Long
    Dim strThreeSwapBanace As String '��������
    Dim objICCard As Object, strCardNo As String, rsOneCard As ADODB.Recordset
    Dim colOneCard As New Collection, blnTransOneCard As Boolean
    Dim strҽ������ As String, rsTmp As ADODB.Recordset
    Dim arrBalance() As String, str���㷽ʽ As String, lng����ID As Long
    Dim cur�ɷ���� As Currency, cur������ As Currency, cur��� As Currency, cur�˿�ϼ� As Currency
    Dim strCurDelNOs As String '�ö��ŷ���,��:'J0002','J00023'
    Dim blnRllTrans As Boolean  '�Ƿ����
    Dim strCurSelNos As String '��ǰѡ�еĵ���
    Dim str�˽��㷽ʽ As String, bln���� As Boolean
    Dim blnThreeSwapComit As Boolean
    Dim lng����ID As Long, lng������� As Long  '43395
    Dim strThreeBalance As String, intCol As Integer
    Dim strOtherBalance As String '�����˷ѷ�ʽ
    Dim blnNotFind As Boolean, str����IDs As String
    Dim blnExistThreeSwap As Boolean, blnExistOneCardSwap As Boolean, blnȫ�� As Boolean
    Dim blnYbComit As Boolean, blnCommited As Boolean, blnOneCardComit As Boolean
    Dim varTemp As Variant, strReclaimInvoice As String, intInvoiceFormat As Integer '����Ʊ��:56963
    Dim cll�˷ѽ���ID As Collection, str�ɹ��˷�ID As String
    Dim strCmdCaptions As String, blnҩƷ As Boolean, blnSel As Boolean
    Dim strYPNos As String, strPrintNOInfor As String  '��ǰ��ӡ�ĵ�����Ϣ:NO:���;
    Dim strReturn As String, strReturnRecipt As String '�˷Ѵ�����Ϣ����ʽ��NO,ҩ��ID|NO,ҩ��ID|��
    Dim dblDelMoney As Double, bln��ȫ�˷� As Boolean
    
    str����IDs = ""
    '��������Ƿ���ȷ
    If mstrNOs = "" Then
        MsgBox "��������Ҫ�˷ѵĵ��ݡ�", vbInformation, gstrSysName
        If txtNO.Visible Then txtNO.SetFocus: Exit Function
    End If
    If CheckBillExistReplenishData(1, , mstrNOs) Then
        MsgBox "ѡ����˷Ѽ�¼������ҽ��������㣬����������˷Ѳ�����", vbInformation, gstrSysName
        Exit Function
    End If
    arrNo = Split(mstrNOs, ",")
    strYPNos = ""
    blnҩƷ = False: blnSel = False
    For i = 1 To vsBill.Rows - 1
        If Val(vsBill.TextMatrix(i, vsBill.ColIndex("ѡ��"))) <> 0 Then
            blnSel = True
            If vsBill.ColIndex("���") <> -1 Then     '47400
                If vsBill.TextMatrix(i, vsBill.ColIndex("���")) Like "*��*ҩ*" _
                    Or vsBill.TextMatrix(i, vsBill.ColIndex("���")) Like "*��*ҩ*" _
                    Or vsBill.TextMatrix(i, vsBill.ColIndex("���")) Like "*����*" Then
                    strYPNos = strYPNos & "," & vsBill.TextMatrix(i, vsBill.ColIndex("���ݺ�"))
                    blnҩƷ = True
                    '81190,Ƚ����,�˷�ҵ����ҩ���ϴ��˷���Ϣ
                    '��ʽ��NO,ҩ��ID|NO,ҩ��ID|��
                    If Not vsBill.TextMatrix(i, vsBill.ColIndex("���")) Like "*����*" Then
                        If InStr(strReturnRecipt & "|", _
                            "|" & vsBill.TextMatrix(i, vsBill.ColIndex("���ݺ�")) & "," & vsBill.TextMatrix(i, vsBill.ColIndex("ִ�п���ID")) & "|") = 0 Then
                            strReturnRecipt = strReturnRecipt & "|" & vsBill.TextMatrix(i, vsBill.ColIndex("���ݺ�")) & "," & vsBill.TextMatrix(i, vsBill.ColIndex("ִ�п���ID"))
                        End If
                    End If
                End If
            End If
        End If
    Next
    
    If blnSel = False Then
        MsgBox "���ڵ���������ѡ��һ��Ҫ�˷ѵ���Ŀ��", vbInformation, gstrSysName
        vsBill.SetFocus: Exit Function
    End If
    
    '47400
    If blnҩƷ Then
        If strYPNos <> "" Then strYPNos = Mid(strYPNos, 2)
        If zlCheckDrugIsPutDrug(strYPNos) = False Then Exit Function
    End If
    
    '���˺�:28947
    If mintInsure <> 0 Then
        If gclsInsure.CheckInsureValid(mintInsure) = False Then
            Exit Function
        End If
    End If
    '���ж����е����Ƿ񲿷��˷�,�Ծ���Ʊ�ݵĴ���ʽ
    blnAll�����˷� = False
    
    '�����жϽ�ʹ��ҽ�ƿ�����ʱ���Ƿ�Ϊ������
    Dim strCurNO As String
    For j = 1 To vsBill.Rows - 1
        If Val(vsBill.TextMatrix(j, vsBill.ColIndex("ѡ��"))) = 0 Then bln��ȫ�˷� = False: Exit For
        If strCurNO = "" Or strCurNO <> vsBill.TextMatrix(j, vsBill.ColIndex("���ݺ�")) Then
            strCurNO = vsBill.TextMatrix(j, vsBill.ColIndex("���ݺ�"))
            bln��ȫ�˷� = BillDeleteAll(strCurNO, 1, mblnHaveExcuteData)
            If bln��ȫ�˷� Then bln��ȫ�˷� = Not BillExistDelete(strCurNO, 1)
            If bln��ȫ�˷� = False Then Exit For
        End If
    Next
    
    'һ���շѵ���������:mstrNOsֻ�ǿ����˵�,�������е�
    strAllNOs = GetMultiNOs(CStr(arrNo(0)), , , mCurBillType.bln�����������)
    
    strOtherNOs = strAllNOs
    If zlCheckIsMzToZY(mstrNOs, 1) Then
          MsgBox "ע��:" & vbCrLf & _
            "    �õ����Ѿ����������תסԺ���� " & vbCrLf & _
            "    ���Ѿ�������������תסԺ����,�������˷�", vbInformation + vbOKOnly, gstrSysName
          Exit Function
    End If
      
    strCurSelNos = ""
    
    For i = 0 To UBound(arrNo)
        strNo = arrNo(i)
        str��� = "": strBalance = "": lngCount = 0: strThreeBalance = ""
        dblDelMoney = 0
                       
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
                    dblDelMoney = dblDelMoney + Val(vsBill.TextMatrix(j, .ColIndex("ʵ�ս��")))
                End If
                lngCount = lngCount + 1
            Next
        End With
        str��� = Mid(str���, 2)
        
        If str��� <> "" Then
            blnCur�����˷� = Not BillDeleteAll(strNo, 1, mblnHaveExcuteData)
            strPrintNOInfor = strPrintNOInfor & ";" & strNo & ":" & str���
            
            If UBound(Split(str���, ",")) + 1 = lngCount And blnCur�����˷� = False Then str��� = ""
            If mintInsure <> 0 Then
                strAllBalance = Getҽ�����㷽ʽ(strNo)
                For j = 0 To UBound(Split(strAllBalance, ","))
                    strTmp = Split(strAllBalance, ",")(j)
                    If Not mblnYB�������� Then
                          '��֧�������������ʱ,ֻ���������Ϊ�ֽ�,����ԭ����,������ҽ������
                        If strTmp = mstr�����ʻ� Then strBalance = "," & strTmp
                    End If
                    If mblnYB�������� Then
                        If Not gclsInsure.GetCapability(support�����������, mlng����ID, mintInsure, strTmp) Then
                            strBalance = strBalance & "," & strTmp
                        End If
                    End If
                Next
            End If
            'ҽ�������˷Ѽ��
            'Or BillExistDelete(strNO, 1):�������˷�,���һ���˷ѵ�û�з�Ʊ�ջ�,����ʾ�ش�,Ҳû�ش�Ʊ��,���Բ�Ӧ�ü�:Or BillExistDelete(strNO, 1)
            blnCur�����˷� = Not (Not blnCur�����˷� And str��� = "")
            If blnCur�����˷� Then blnAll�����˷� = True '���ŵ���Ϊ�����˷�,�����е���Ϊ�����˷�
           '��������
            If mCurBillType.blnSingleBalance And mCurBillType.bln����ҽ�ƿ����� And Not bln��ȫ�˷� Then
                If vsBalance.RowHidden(1) = False And Val(vsBalance.TextMatrix(1, 2)) <> 0 Then '������
                    strThreeBalance = strThreeBalance & "," & vsBalance.Cell(flexcpData, 1, 1) & "|" & dblDelMoney
                End If
            Else
                mrsBalance.Filter = "NO='" & strNo & "' and ����>=3"
                With mrsBalance
                     Do While Not .EOF
                         If blnCur�����˷� Then
                            strThreeBalance = strThreeBalance & "," & Nvl(!���㷽ʽ)
                        End If
                        If Val(Nvl(!�Ƿ�����)) = 1 Then
                            '�Ŷӽ��ֽ�
                            For intCol = 1 To vsBalance.COLS - 1 Step 2
                                If vsBalance.Cell(flexcpData, 1, intCol) = !���㷽ʽ And _
                                    Val(vsBalance.TextMatrix(1, intCol + 1)) = 0 Then
                                    strThreeBalance = strThreeBalance & "," & Nvl(!���㷽ʽ)
                                End If
                            Next
                        End If
                        .MoveNext
                    Loop
               End With
               
                strOtherBalance = ""
               '�������㷽ʽ:�����˷�
                mrsBalance.Filter = "NO='" & strNo & "' "
                With mrsBalance
                     Do While Not .EOF
                         If InStr(",1,2,", "," & Val(Nvl(!��������)) & ",") > 0 Then
                             blnNotFind = True
                             For intCol = 1 To vsBalance.COLS - 1 Step 2
                                 If vsBalance.Cell(flexcpData, 1, intCol) = !���㷽ʽ And vsBalance.TextMatrix(1, intCol) <> "" Then
                                     blnNotFind = False: Exit For
                                 End If
                             Next
                             If blnNotFind Then
                                 strOtherBalance = strOtherBalance & "," & Nvl(!���㷽ʽ)
                             End If
                         End If
                         .MoveNext
                     Loop
                End With
            End If
            colOrder.Add str���, "_" & strNo
            lng����ID = Val(vsBill.TextMatrix(k, vsBill.ColIndex("����ID")))
        
        'һ��ͨ���
        If Not CheckOnCardValied(blnCur�����˷�, lng����ID) Then Exit Function
        '�������׼��
            If mCurBillType.blnSingleBalance And mCurBillType.bln����ҽ�ƿ����� And Not bln��ȫ�˷� Then
                If mCurBillType.bln������ȫ�� Then
                    If Val(vsBalance.ColData(2)) = 0 Then '������
                        MsgBox "��ǰ����ʹ���˵��������㽻�ף����е��ݱ���ȫ�ˣ�", vbInformation, gstrSysName
                        Exit Function
                    ElseIf cbo�˿ʽ.Visible = False Then
                        MsgBox "��ǰ����ʹ���˵��������㽻�ף����е��ݱ���ȫ�ˣ����������ѡ����Ϊ�������㷽ʽ��", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
                
                If strThreeBalance <> "" Then
                    mrsBalance.Filter = "����ID=" & lng����ID & " And ����=3"
                    If mrsBalance.RecordCount = 0 Then
                        MsgBox "��ǰ���� " & strNo & " ʹ���˵��������㽻�ף���δ����ԭʼ�������ݣ���˲飡", vbInformation, gstrSysName
                        Exit Function
                    End If
                    With mrsBalance
                        If .RecordCount <> 0 Then
                            .MoveFirst
                            If zlCheckDelValied(Val(Nvl(!�����ID)), Nvl(!����), False, Nvl(!����), Nvl(!������ˮ��), _
                                Nvl(!����˵��), lng����ID, dblDelMoney) = False Then Exit Function
                        End If
                    End With
                End If
            Else
                If Not CheckThreeSwapValied(blnCur�����˷�, lng����ID, InStr(1, mstrNOs, ",") > 0) Then Exit Function
            End If
            If mintInsure <> 0 And blnCur�����˷� Then      'ҽ��֧���˷�ʱ,ÿһ��Ҫ��ȫ��
                If str��� <> "" Then
                    MsgBox "����""" & strNo & """�������ս�����ã�������һЩ��Ŀ�����Ѿ�ִ�У����������˷ѡ�", vbInformation, gstrSysName
                Else
                    MsgBox "����""" & strNo & """�������ս�����ã����������˷ѡ�", vbInformation, gstrSysName
                End If
                vsBill.SetFocus: Exit Function
            End If
            
            '�жϱ����Ƿ�����ʱ���ſ����ŵ���
            strOtherNOs = Mid(Replace("," & strOtherNOs, ",'" & strNo & "'", ""), 2)
        Else
            blnAll�����˷� = True                       '���ŵ��ݲ��˷�,�����е���Ϊ�����˷�
            colOrder.Add "δѡ��", "_" & strNo
        End If
        
        'ҽ���������˵Ľ��㷽ʽ,��ҽ��ʱΪ��
        If strBalance <> "" Then strBalance = Mid(strBalance, 2)
        If strThreeBalance <> "" Then strThreeBalance = Mid(strThreeBalance, 2)
        If strOtherBalance <> "" Then strOtherBalance = Mid(strOtherBalance, 2)
        colBalance.Add strBalance, "_" & strNo
        colThreeBalance.Add strThreeBalance, "_" & strNo
        colOtherBalance.Add strOtherBalance, "_" & strNo
        
        'ҽ���˷ѽ���Ҫ�õĽ���ID,��ҽ��ʱΪ0,��֧�����ϵ�����ҽ��,������ҽ������
        If mblnYB�������� And mintInsure <> 0 Then
            colBalanceID.Add Val(vsBill.TextMatrix(k, vsBill.ColIndex("����ID"))), "_" & strNo
        Else
            colBalanceID.Add 0, "_" & strNo
        End If
    Next
    
    '�������������Ƿ�δ����,����жϳ����е����Ƿ񲿷��˷�
    If (Not blnAll�����˷�) And strOtherNOs <> "" Then
        If BillExistMoney(strOtherNOs, 1) Then blnAll�����˷� = True
    End If
    
    'Ԥ����ؼ�����֤
    If strCurSelNos <> "" Then
        strCurSelNos = Mid(strCurSelNos, 2)
        If Not zlCheckPrepayBack(mlng����ID, strCurSelNos) Then
            Exit Function
        End If
    End If
    
    If blnAll�����˷� Then
        '56963
        If gTy_Module_Para.bytƱ�ݷ������ <> 0 And strReclaimInvoice = "" Then
            strReclaimInvoice = zlGetReclaimInvoice(Mid(strPrintNOInfor, 2))
        End If
        If Not (gTy_Module_Para.bytƱ�ݷ������ <> 0 And strReclaimInvoice <> "") Then
                If InStr(mstrPrivs, "�����˷�") = 0 Then
                    MsgBox "��û��Ȩ��ִ�в����˷Ѳ�����", vbInformation, gstrSysName
                    vsBill.SetFocus: Exit Function
                End If
                If gTy_Module_Para.bln������ Then
                    MsgBox "�Զ���ȡ������ʱ���������˷ѡ�", vbInformation, gstrSysName: vsBill.SetFocus: Exit Function
                End If
                
                '���˺� ����:27352 ����:2010-01-13 10:26:08
                If InStr(1, mstrPrivs, "�˷Ѻ��շ�Ʊ") > 0 Then
                    If frmReInvoice.ShowMe(Me, strNo, Val(txtAllTotal.Text), Val(txt�˿���.Text), strInvoices) = False Then
                        vsBill.SetFocus: Exit Function
                    End If
                End If
        End If
    End If
    
    If mBillDelType = EM_����ȫ�� And blnAll�����˷� Then
        MsgBox "���ŵ���ʹ��һ��ͨ����ģʽ��ҽ���˷�Ҫ�������ˣ����������˷ѣ�", vbInformation, gstrSysName
        vsBill.SetFocus: Exit Function
    End If
                      
    If mintInsure <> 0 And MCPAR.ҽ���ӿڴ�ӡƱ�� Then
        If zlGetInvoiceGroupUseID(lng����ID) = False Then Exit Function
        strInvoice = GetNextBill(lng����ID)
    End If
    
    
    DateDel = zlDatabase.Currentdate
    Set cll�˷ѽ���ID = New Collection
    
    '����Ҫִ�е�SQL
    lng������� = 0
    'ҽ��Ҫ����Ϊ����,���,�����һ���ȳ���
    For i = UBound(arrNo) To 0 Step -1
        arrSQL = Array(): strNo = arrNo(i)
        If colOrder("_" & strNo) <> "δѡ��" Then
            cur����� = Val(mcolError("_" & strNo))
           '60974
            If mintInsure <> 0 And colBalance("_" & strNo) = "" And MCPAR.�൥���շѱ���ȫ�� Then cur����� = 0    'ҽ������ȫ���ҽ���ȫ֧������ʱ�����
            lng����ID = zlDatabase.GetNextId("���˽��ʼ�¼")
            str����IDs = str����IDs & "," & lng����ID
            If lng������� = 0 Then lng������� = lng����ID
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            
            'Zl_�����շѼ�¼_Delete
            strSQL = "zl_�����շѼ�¼_DELETE("
            '  No_In           ������ü�¼.NO%Type,
            strSQL = strSQL & "'" & strNo & "',"
            '  ����Ա���_In   ������ü�¼.����Ա���%Type,
            strSQL = strSQL & "'" & UserInfo.��� & "',"
            '  ����Ա����_In   ������ü�¼.����Ա����%Type,
            strSQL = strSQL & "'" & UserInfo.���� & "',"
            '  ҽ�����㷽ʽ_In Varchar2 := Null,
            strSQL = strSQL & "'" & colBalance("_" & strNo) & "',"
            '  ���_In         Varchar2 := Null,
            strSQL = strSQL & "'" & colOrder("_" & strNo) & "',"
            '  ���㷽ʽ_In     ����Ԥ����¼.���㷽ʽ%Type := Null,
            strSQL = strSQL & "'" & zlStr.NeedName(cbo�˿ʽ.Text) & "',"
            '  ���_In         ������ü�¼.ʵ�ս��%Type := 0,
            strSQL = strSQL & "" & cur����� & ","
            '  �˷�ʱ��_In     ������ü�¼.�Ǽ�ʱ��%Type := Null,
            strSQL = strSQL & "To_Date('" & Format(DateDel, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
            '  ����Ʊ��_In     Number := 0,
            strSQL = strSQL & "" & IIf(blnAll�����˷�, "1", "0") & ","
            '  �˷�ժҪ_In     ������ü�¼.ժҪ%Type := Null
            strSQL = strSQL & "" & IIf(Trim(txt�˷�ժҪ.Text) = "", "NULL", "'" & Trim(txt�˷�ժҪ.Text) & "'") & ","
            '     У�Ա�־_In: 0-����Ҫ�϶�;1-��϶�(��������Ա�ɿ����,������Ʊ��)
            strSQL = strSQL & "1,"
            '  ����id_In       ����Ԥ����¼.����id%Type := Null,
            strSQL = strSQL & lng����ID & ","
            '  �������_In     ����Ԥ����¼.�������%Type := Null
            strSQL = strSQL & lng������� & ","
              'һ��ͨ����_In   Varchar2 := Null
             strOtherBalance = colOtherBalance("_" & strNo)
            'If Not blnAll�����˷� Then strOtherBalance = ""
             strSQL = strSQL & "'" & colThreeBalance("_" & strNo) & _
                IIf(colThreeBalance("_" & strNo) <> "" And strOtherBalance <> "", ",", "") & strOtherBalance & "',"
             '�˿����_In:1-���в�����(���˿ʽ�˵�ָ���Ľ��㷽ʽ<���㷽ʽ_In>��,0-��ָ���˿ʽ)
             If (blnAll�����˷� Or mCurBillType.bln���Ų����˷�) And mintInsure = 0 Then
                '��ͨ����
                '����Ƿ��˵�ָ���Ľ��㷽ʽ<���㷽ʽ_In>��
                blnNotFind = True
                For intCol = 1 To vsBalance.COLS - 1 Step 2
                    If Val(vsBalance.TextMatrix(1, intCol + 1)) <> 0 Then blnNotFind = False: Exit For
                Next
                strSQL = strSQL & IIf(cbo�˿ʽ.Visible And (vsBalance.RowHidden(1) Or blnNotFind), "1", "0") & ","
             Else
                strSQL = strSQL & "0,"
             End If
             '�൥��ȫ��_IN=1-�൥��ȫ��(���ŵ���ȫ��,ԭ����);0-��ԭ����:60974
              'strSQL = strSQL & IIf(Not vsBalance.RowHidden(1), "1", "0") & ")"
              If mintInsure <> 0 And colBalance("_" & strNo) = "" And MCPAR.�൥���շѱ���ȫ�� Then
                 strSQL = strSQL & "1)"
              Else
                 strSQL = strSQL & IIf(cbo�˿ʽ.Visible Or blnAll�����˷� Or cur����� <> 0, "0", "1") & ")"
              End If
            arrSQL(UBound(arrSQL)) = strSQL
            '60974
'            If cur����� <> 0 Then
'                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
'                arrSQL(UBound(arrSQL)) = "zl_�����շ����_Insert('" & strNO & "'," & cur����� & ",1)"
'            End If
            cll�˷ѽ���ID.Add lng����ID, "_" & strNo
            strCurDelNOs = strCurDelNOs & ",'" & strNo & "'"
        End If
        colSQL.Add arrSQL, "_" & strNo '��ǰ���ݵ�SQL��
    Next
    
    bln���� = False
    If cbo�˿ʽ.ListIndex >= 0 Then
        bln���� = cbo�˿ʽ.ItemData(cbo�˿ʽ.ListIndex) = 1
        str�˽��㷽ʽ = zlStr.NeedName(cbo�˿ʽ.Text)
    Else
        bln���� = True
        str�˽��㷽ʽ = IIf(mstr�ֽ���㷽ʽ = "", "�ֽ�", mstr�ֽ���㷽ʽ)
    End If
    '56963
    If strPrintNOInfor <> "" Then strPrintNOInfor = Mid(strPrintNOInfor, 2)
    strReclaimInvoice = zlGetReclaimInvoice(strPrintNOInfor)
    
    If gTy_Module_Para.bytƱ�ݷ������ <> 0 And strReclaimInvoice <> "" Then
        If InStr(1, mstrPrivs, "�˷Ѻ��շ�Ʊ") > 0 Then
            If MsgBox("ע��:" & vbCrLf & " ��ǰ�˷ѵĵ����а��������շ�Ʊ�ݣ��Ƿ������ЩƱ��?" & vbCrLf & strReclaimInvoice, vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Function
        End If
    End If
    If gblnBillPrint Then
        If gobjBillPrint.zlEraseBill(strAllNOs, 0) = False Then Exit Function
    End If
    
    '1-Ԥ���,2-ҽ��,3-ҽ�ƿ�,4-���㿨,5-һ��ͨ,0-������
    mrsBalance.Filter = "����=3 or ����=4"  '���ѿ������п�
    blnExistThreeSwap = mrsBalance.RecordCount <> 0
    mrsBalance.Filter = "�Ƿ�ȫ��=1"
    blnȫ�� = mrsBalance.RecordCount <> 0
    mrsBalance.Filter = "����=5"
    blnExistOneCardSwap = mrsBalance.RecordCount <> 0
    mrsBalance.Filter = 0
    
    'ִ���˷ѵ�SQL
    On Error GoTo errH
    strDelNOs = ""
    blnCommited = False: blnYbComit = False
    If mintInsure <> 0 And (MCPAR.�൥��һ�ν��� Or MCPAR.�൥�ݵ�һ�ν���) Then
        '���ŵ���ҽ��һ�ν���
        gcnOracle.BeginTrans: blnTrans = True
        strAllBalance = "": strBalance = ""
        For i = 0 To UBound(arrNo)
            strNo = arrNo(i)          '�����һ�ſ�ʼ��
            For j = 0 To UBound(colSQL("_" & strNo))
                Call zlDatabase.ExecuteProcedure(CStr(colSQL("_" & strNo)(j)), Me.Caption)
            Next
            strAllBalance = IIf(strAllBalance = "", "", strAllBalance & ",") & colBalanceID("_" & strNo)
            If i = 0 Then strBalance = colBalanceID("_" & strNo)
        Next
        '�Ȳ���Ʊ�ݣ�ҽ���ӿڲ���ȡ��
        If MCPAR.ҽ���ӿڴ�ӡƱ�� _
            And Not (gTy_Module_Para.bytƱ�ݷ������ <> 0 And strReclaimInvoice <> "") Then
            '56963
            strSQL = "zl_�����շѼ�¼_RePrint('" & strNo & "','" & strInvoice & "'," & ZVal(lng����ID) & ",'" & UserInfo.���� & "'," & _
                "To_Date('" & Format(DateDel, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),1,1)"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        End If
        
        If DelInsureMulitOneBalance(blnExistThreeSwap, arrNo, Val(strBalance), strAllBalance, strҽ������, str�˽��㷽ʽ, bln����, blnCommited) = False Then
            If Not blnCommited Then gcnOracle.RollbackTrans
            Exit Function
        End If
        If blnCommited = True Then
            blnYbComit = True: gcnOracle.BeginTrans: blnTrans = True
        End If
    Else
        '-------------------------------------------------------------------------------------------------------
        '���˺�:ҽ����strAdvancey����:�����˷�������|��ǰ�˷ѵڼ���:27231
        Dim lngPages As Long, lngPage, cllYB As Collection
        Set cllYB = New Collection
        lngPage = 0: lngPages = 0
        For i = UBound(arrNo) To 0 Step -1
            strNo = arrNo(i)
            If UBound(colSQL("_" & strNo)) >= 0 And mintInsure <> 0 Then
                'ҽ����
                 If mblnYB�������� And colBalanceID("_" & strNo) <> 0 Then
                    lngPage = lngPage + 1
                    lngPages = lngPages + 1
                    cllYB.Add lngPage, "_" & strNo
                 End If
            End If
        Next
        
        '-------------------------------------------------------------------------------------------------------
        '�ȴ���ҽ��
        If blnExistThreeSwap And blnȫ�� Then
               '�Ƚ����е����˷�,Ȼ��ҽ��������
               gcnOracle.BeginTrans '��������
               blnTrans = True
                For i = 0 To UBound(arrNo)
                    strNo = arrNo(UBound(arrNo) - i)        '�����һ�ſ�ʼ��
                    If UBound(colSQL("_" & strNo)) >= 0 Then
                        For j = 0 To UBound(colSQL("_" & strNo))
                            Call zlDatabase.ExecuteProcedure(CStr(colSQL("_" & strNo)(j)), Me.Caption)
                        Next
                    End If
                Next
                '�ֵ��ݴ���ҽ��
                For i = 0 To UBound(arrNo)
                
                    strNo = arrNo(UBound(arrNo) - i)        '�����һ�ſ�ʼ��
                    If UBound(colSQL("_" & strNo)) >= 0 Then
                        '��ҽ��
                        blnCommited = False
                        If mintInsure <> 0 And mblnYB�������� Then
                            lngPage = Val(cllYB("_" & strNo)): lng����ID = Val(colBalanceID("_" & strNo))
                            If Not DelInsureOneBill(strҽ������, blnExistThreeSwap, lng����ID, lngPage, lngPages, blnCommited) Then
                                If blnCommited = False Then gcnOracle.RollbackTrans
                                blnTrans = False
                                '��ʾ�˷ѳɹ�������ʾ
                                Call ShowErrBill(strDelNOs, strNo, 3): Exit Function
                            End If
                        End If
                        If blnCommited Then
                            blnTrans = False
                            gcnOracle.BeginTrans    'ֻ���ύ�󣬲�������
                            blnYbComit = True: blnTrans = True 'ֻҪ��һ�֣��ͻ��ύ
                        End If
                        strDelNOs = strDelNOs & IIf(strDelNOs = "", "", ",") & strNo
                    End If
                Next
        Else
               str�ɹ��˷�ID = ""
                gcnOracle.BeginTrans '��������
                blnTrans = True
                For i = 0 To UBound(arrNo)
                    strNo = arrNo(UBound(arrNo) - i)       'ҽ��Ҫ������һ�ſ�ʼ��
                    If UBound(colSQL("_" & strNo)) >= 0 Then
                        For j = 0 To UBound(colSQL("_" & strNo))
                            Call zlDatabase.ExecuteProcedure(CStr(colSQL("_" & strNo)(j)), Me.Caption)
                        Next
                        '��ҽ��
                        blnCommited = False
                        If mintInsure <> 0 And mblnYB�������� Then
                            lngPage = Val(cllYB("_" & strNo)): lng����ID = Val(colBalanceID("_" & strNo))
                            If Not DelInsureOneBill(strҽ������, blnExistThreeSwap, lng����ID, lngPage, lngPages, blnCommited) Then
                                If Not blnCommited Then gcnOracle.RollbackTrans: blnTrans = False
                                gcnOracle.BeginTrans:  blnTrans = True
                                If strDelNOs <> "" And blnExistThreeSwap And blnExistOneCardSwap Then
                                    varTemp = Split(strDelNOs, ",")
                                    If Not DelSawpSpecifyNOs(varTemp, blnExistThreeSwap, blnExistOneCardSwap, strNo, blnCommited) Then
                                        If Not blnCommited Then gcnOracle.RollbackTrans
                                        Exit Function
                                    End If
                                    If Not blnCommited Then gcnOracle.CommitTrans
                                    gcnOracle.BeginTrans:   blnTrans = True
                                End If
                                If strDelNOs <> "" Then
                                    If OverFeeDel(str�ɹ��˷�ID, mtyPati.����ID, blnCommited) = False Then
                                        If blnOneCardComit = False And blnYbComit = False And blnThreeSwapComit = False Then
                                            gcnOracle.RollbackTrans: Exit Function
                                        End If
                                        Exit Function
                                    End If
                                    If blnCommited Then blnTrans = False
                                End If
                                If blnTrans Then gcnOracle.RollbackTrans: blnTrans = True
                                '�Գɹ����ֽ�������շ�
                                '��ʾ�˷ѳɹ�������ʾ
                                Call ShowErrBill(strDelNOs, strNo): Exit Function
                            End If
                        End If
                        If blnCommited Then
                            gcnOracle.BeginTrans    'ֻ���ύ�󣬲�������
                            blnYbComit = True: blnTrans = True 'ֻҪ��һ�֣��ͻ��ύ
                        End If
                        strDelNOs = strDelNOs & IIf(strDelNOs = "", "", ",") & strNo
                        str�ɹ��˷�ID = str�ɹ��˷�ID & IIf(str�ɹ��˷�ID = "", "", ",") & cll�˷ѽ���ID("_" & strNo)
                    End If
                Next
            End If
    End If
    
    If Not blnTrans Then gcnOracle.BeginTrans: blnTrans = True
    If strDelNOs <> "" Then
        varTemp = Split(strDelNOs, ",")
    Else
        varTemp = arrNo
    End If
    '------------------------------------------------------------------------------------------
 
    '��һ��ͨ
ReDOOneCard:
    blnCommited = False
    If Not DelOneCardPay(varTemp, blnCommited) Then
        If blnCommited = False Then gcnOracle.RollbackTrans: blnTrans = False
        If blnYbComit Then
            strCmdCaptions = "�쳣����(&C)|��ʾ������һ��ͨ����,���ݽ����쳣��ʽ����,�������ڽ����д���"
            strCmdCaptions = strCmdCaptions & ";����(&R)|��ʾ���µ���һ��ͨ���㽻��"
            If frmVerfyCodeInput.ShowMsg(Me, "����[" & strDelNOs & "]�Ѿ��˷ѳɹ�,��һ��ͨ����ʧ��,[�쳣����]����������֤��,���鲻�����쳣���ݱ���", strCmdCaptions) = False Then
                 gcnOracle.BeginTrans: blnTrans = True
                GoTo ReDOOneCard:
            End If
        End If
        Call ClearFace(True, False)
        Exit Function
    End If
    
    If blnCommited Then
        blnOneCardComit = True: blnTrans = False
        gcnOracle.BeginTrans: blnTrans = True
    End If
    '------------------------------------------------------------------------------------------
    '��һ��ͨ�ȵ���������
ReDOThreeSwap:
    blnCommited = False
    If mCurBillType.blnSingleBalance And mCurBillType.bln����ҽ�ƿ����� And Not bln��ȫ�˷� Then
        If Not DelThreeSwapFeeSingle(varTemp, colThreeBalance, colOrder, str����IDs, blnCommited) Then
            If blnCommited = False Then gcnOracle.RollbackTrans: blnTrans = True
            
            If blnOneCardComit Or blnYbComit Then
                strCmdCaptions = "�쳣����(&C)|��ʾ��������������,���ݽ����쳣��ʽ����,�������ڽ����д���"
                strCmdCaptions = strCmdCaptions & ";����(&R)|��ʾ���µ����������㽻�׽����˷�"
                If frmVerfyCodeInput.ShowMsg(Me, "����[" & "4243;432432" & "]�Ѿ�" & IIf(blnYbComit, "ҽ��", "") & IIf(blnOneCardComit, IIf(blnYbComit, "��", "") & "һ��ͨ", "") & "�˷ѳɹ�,�����������˷�ʧ��,[�쳣����]����������֤��,���鲻�����쳣���ݱ���", strCmdCaptions) = False Then
                  If blnCommited Then gcnOracle.BeginTrans: blnTrans = True
                  GoTo ReDOThreeSwap:
                End If
            End If
            Call ClearFace(True, False)
            Exit Function
        End If
    Else
        If Not DelThreeSwapFee(varTemp, blnCommited) Then
            If blnCommited = False Then gcnOracle.RollbackTrans: blnTrans = True
            
            If blnOneCardComit Or blnYbComit Then
                strCmdCaptions = "�쳣����(&C)|��ʾ��������������,���ݽ����쳣��ʽ����,�������ڽ����д���"
                strCmdCaptions = strCmdCaptions & ";����(&R)|��ʾ���µ����������㽻�׽����˷�"
                If frmVerfyCodeInput.ShowMsg(Me, "����[" & "4243;432432" & "]�Ѿ�" & IIf(blnYbComit, "ҽ��", "") & IIf(blnOneCardComit, IIf(blnYbComit, "��", "") & "һ��ͨ", "") & "�˷ѳɹ�,�����������˷�ʧ��,[�쳣����]����������֤��,���鲻�����쳣���ݱ���", strCmdCaptions) = False Then
                  If blnCommited Then gcnOracle.BeginTrans: blnTrans = True
                  GoTo ReDOThreeSwap:
                End If
            End If
            Call ClearFace(True, False)
            Exit Function
        End If
    End If
    If blnCommited Then
        blnThreeSwapComit = True: blnTrans = False
        gcnOracle.BeginTrans: blnTrans = True
    End If
    '------------------------------------------------------------------------------------------
    '����շ�
    blnCommited = False
    If OverFeeDel(str����IDs, mtyPati.����ID, blnCommited) = False Then
        If blnCommited = False Then
            gcnOracle.RollbackTrans: Exit Function
        End If
        Call ClearFace(True, False)
        Exit Function
    End If
    
    If Not blnCommited Then         '��ͨ����,����,ֱ���ύ
        gcnOracle.CommitTrans: Exit Function
    End If
    
    '81190,Ƚ����,�˷�ҵ����ҩ���ϴ��˷���Ϣ
    On Error Resume Next
    If mblnDrugPacker Then
        If strReturnRecipt <> "" Then
            strReturnRecipt = Mid(strReturnRecipt, 2)
            Call mobjDrugPacker.DYEY_MZ_TransRecipeReturn(1, UserInfo.���, UserInfo.����, strReturnRecipt, strReturn)
        End If
    End If
    Err.Clear: On Error GoTo errH
    
    '��ӡ�˷ѵ���
    Call PrintDelBill(strAllNOs, strCurDelNOs, strNo, mtyPati.����ID, DateDel, blnAll�����˷�, strInvoices, strReclaimInvoice)
    ExecDelete = True
    Exit Function
errH:
    blnRllTrans = False
    If Err.Number <> 0 Then
        If blnTrans Then gcnOracle.RollbackTrans: blnRllTrans = True
        If ErrCenter() = 1 Then
            Resume
        End If
    End If
        
    If blnTrans Then
        If Not blnRllTrans Then gcnOracle.RollbackTrans: blnRllTrans = True
        'ҽ����HIS����ͬһ������,HIS����ʧ��,��ҽ���������ϴ�,������Ҫ��"ȡ������"�ӿ�
        If blnTransMedicare Then Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, False, mintInsure)
        If blnTransOneCard Then MsgBox "һ��ͨ�˷ѽ��׵���ʧ�ܣ��˷Ѳ���ʧ�ܣ�", vbExclamation, gstrSysName
    End If
    
    If Err.Number <> 0 Then Call SaveErrLog
    '�ж���ʾ,����ӡ�������˷Ѻ��ٴ�ӡ���Լ�ѡ���ش�
    Call ShowErrBill(strDelNOs, strNo)
    Exit Function
ErrRquare:
    blnRllTrans = False
    If Err.Number <> 0 Then
        If blnTrans Then gcnOracle.RollbackTrans: blnRllTrans = True
        If ErrCenter() = 1 Then
            Resume
        End If
    End If
    If blnTrans Then
        If Not blnRllTrans Then gcnOracle.RollbackTrans: blnRllTrans = True
         MsgBox "���㿨�˷ѽ��׵���ʧ�ܣ�", vbExclamation, gstrSysName
    End If
    If Err.Number <> 0 Then Call SaveErrLog
    If txtNO.Visible Then txtNO.SetFocus
End Function

Private Sub PrintDelBill(ByVal strAllNOs As String, ByVal strCurDelNOs As String, _
    ByVal strNo As String, _
    ByVal lng����ID As Long, _
    ByVal dtDateDel As Date, ByVal blnAll�����˷� As Boolean, _
    ByVal strInvoices As String, ByVal strReclaimInvoice As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ӡ���Ʊ��
    '���: strAllNOs-��ǰ�漰�����е���
    '       strCurDelNOs-��ǰ�˷ѵĵ���
    '       dtDateDel-�˷�����
    '       strInvoices-ѡ��ķ�Ʊ��(��ģʽ)
    '       strReclaimInvoice-���յķ�Ʊ��
    '����:
    '����:���˺�
    '����:2013-05-27 16:41:06
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intInvoiceFormat As Integer, blnPrint As Integer
    Dim str��Ʊ�� As String, intƱ������ As Integer
    Dim strSQL As String
    Dim strPriceGrade As String
    
    On Error GoTo errHandle
    If Not blnAll�����˷� Then
         '˰�ز���ȫ��ʱ�ջش���(ȫ��ʱ��zl_�����շѼ�¼_DELETE�����ջ�Ʊ��)
        If Not gobjTax Is Nothing And gblnTax Then
            gstrTax = gobjTax.zlTaxOutErase(gcnOracle, strAllNOs)
            If gstrTax <> "" Then MsgBox gstrTax, vbExclamation, gstrSysName
        End If
        GoTo PrintList:
        Exit Sub
    End If
    
    '�����˷�ʱ�ջز��ش�,�������Ų����˺��˶����е�ĳ����
    If gTy_Module_Para.bytƱ�ݷ������ <> 0 And strReclaimInvoice <> "" Then
        '����Ʊ�ݷ�������ӡ
        '��Ԥ��,��Ʊ���Ƿ����
        str��Ʊ�� = strReclaimInvoice
        If zlExeCuteBillNoSplit(True, 4, mlng����ID, strAllNOs, lng����ID, "", dtDateDel, 1, str��Ʊ��, intƱ������) = False Then GoTo PrintList:
        If intƱ������ = 0 Then
            'ֻ����Ʊ��,������ӡ
            str��Ʊ�� = strReclaimInvoice
            Call zlExeCuteBillNoSplit(False, 4, mlng����ID, strAllNOs, lng����ID, "", dtDateDel, 1, str��Ʊ��, intƱ������)
            GoTo PrintList:
        End If
        blnPrint = True
        ''0-����ӡ;1-�Զ���ӡ;2-��ʾ��ӡ
        If mintInvoicePrint = 0 Then blnPrint = False   '�Զ���ӡ
        If mintInvoicePrint = 2 Then
            If MsgBox("�Ƿ��ӡƱ�ݣ�", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then blnPrint = False
        End If
        '�ش��ջط�Ʊ
        If blnPrint Then
            intInvoiceFormat = IIf(strReclaimInvoice = "" And gTy_Module_Para.bytƱ�ݷ������ <> 0, mintOldInvoiceFormat, mintInvoiceFormat)
            If gintPriceGradeStartType >= 2 Then
                strPriceGrade = GetPriceGradeFromNos(strAllNOs)
            Else
                strPriceGrade = gstr��ͨ�۸�ȼ�
            End If
            Call RePrintCharge(1, strAllNOs, Me, mlng����ID, strReclaimInvoice, True, dtDateDel, _
                intInvoiceFormat, , , mlngShareUseID, mstrUseType, , strPriceGrade)
        End If
        GoTo PrintList:
        Exit Sub
    End If
    If strInvoices = "" Then 'a.�ջز����´�ӡ�����վ�
        blnPrint = True
        ''0-����ӡ;1-�Զ���ӡ;2-��ʾ��ӡ
        If mintInvoicePrint = 0 Then blnPrint = False   '�Զ���ӡ
        If mintInvoicePrint = 2 Then
            If MsgBox("�Ƿ��ӡƱ�ݣ�", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then blnPrint = False
        End If
        
        If blnPrint Then
            intInvoiceFormat = IIf(strReclaimInvoice = "" And gTy_Module_Para.bytƱ�ݷ������ <> 0, mintOldInvoiceFormat, mintInvoiceFormat)
            If gintPriceGradeStartType >= 2 Then
                strPriceGrade = GetPriceGradeFromNos(strAllNOs)
            Else
                strPriceGrade = gstr��ͨ�۸�ȼ�
            End If
            Call RePrintCharge(1, strAllNOs, Me, mlng����ID, strReclaimInvoice, True, dtDateDel, _
            intInvoiceFormat, , , mlngShareUseID, mstrUseType, , strPriceGrade)
        End If
        GoTo PrintList:
        Exit Sub
    End If
    
    'b.�շѻ���һ����ʱû�д�ӡƱ��
    If strInvoices <> "�޿���Ʊ��" Then
        'c.ֻ�ջ�Ʊ��
        strSQL = "zl_�����շѼ�¼_RePrint('" & strNo & "',Null,0,'" & UserInfo.���� & "'," & _
            "To_Date('" & Format(dtDateDel, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),1,0,'" & strInvoices & "')"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    End If
PrintList:
    If blnAll�����˷� Then
        '��ӡ�����嵥
        If InStr(mstrPrivs, "��ӡ�嵥") > 0 Then
            If gint�շ��嵥 = 1 Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_3", Me, "NO=" & strAllNOs, "ҩƷ��λ=" & IIf(gblnҩ����λ, 1, 0), 2)
            ElseIf gint�շ��嵥 = 2 Then
                If MsgBox("Ҫ��ӡ�շ��嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_3", Me, "NO=" & strAllNOs, "ҩƷ��λ=" & IIf(gblnҩ����λ, 1, 0), 2)
                End If
            End If
        End If
    End If
    '79448,Ƚ����,2014-11-10,��ӡ�ص�ʱ�����������,��Ϊ��",'O0000678','O0000679'"��Ӧ��ȥ����һ������","
    If strCurDelNOs <> "" Then strCurDelNOs = Mid(strCurDelNOs, 2)
    If mintInsure <> 0 And MCPAR.�˷Ѻ��ӡ�ص� And InStr(1, mstrPrivs, "ҽ���˷ѻص�") > 0 Then
        '����:35248
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_4", Me, "NO=" & strCurDelNOs, 2)
    End If
    If mint�˷ѻص���ӡ = 1 Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_5", Me, "NO=" & strCurDelNOs, 2)
    ElseIf mint�˷ѻص���ӡ = 2 Then
        If MsgBox("�Ƿ��ӡ�˷ѻص���", vbYesNo + vbQuestion, gstrSysName) = vbYes Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_5", Me, "NO=" & strCurDelNOs, 2)
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function DelSawpSpecifyNOs(ByVal varNO As Variant, _
    ByVal blnExistThreeSwap As Boolean, _
    ByVal blnExistOneCardSwap As Boolean, _
     Optional strNOLost As String, Optional ByRef blnCommited As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������׺�һ��ͨ��ָ������
    '���:varTemp():���ݼ�
    '����:blnCommited-�Ƿ��Ѿ�����������
    '����:�˷ѳɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2012-01-11 15:47:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strDelNOs As String, i As Long
    
    blnCommited = False
    If UBound(varNO) < 0 Then DelSawpSpecifyNOs = True: Exit Function
    If blnExistOneCardSwap = False And blnExistThreeSwap = False Then DelSawpSpecifyNOs = True:  Exit Function
    '����һ��ͨ
    If Not DelOneCardPay(varNO, blnCommited) Then
        '��ʾ������Ϣ
        If Not blnCommited = False Then gcnOracle.RollbackTrans: blnCommited = True
        For i = 0 To UBound(varNO)
            strDelNOs = strDelNOs & "," & varNO(i)
        Next
        If strDelNOs <> "" Then strDelNOs = Mid(strDelNOs, 2)
        Call ShowErrBill(strDelNOs, strNOLost, 1)
        Exit Function
    End If
    blnCommited = False
    '�������ӿڽ���
    If Not DelThreeSwapFee(varNO, blnCommited) Then
        If blnCommited = False Then gcnOracle.RollbackTrans: blnCommited = True
        For i = 0 To UBound(varNO)
            strDelNOs = strDelNOs & "," & varNO(i)
        Next
        If strDelNOs <> "" Then strDelNOs = Mid(strDelNOs, 2)
        '��ʾ������Ϣ
        Call ShowErrBill(strDelNOs, strNOLost, 1)
        Exit Function
    End If
    DelSawpSpecifyNOs = True
End Function

Private Function ShowErrBill(ByVal strDelSucceedNos As String, _
    ByVal strDelLost As String, Optional bytType As Byte = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ֳɹ�������ʾ��Ϣ
    '���:strDelSucceedNos-�˷ѳɹ��ĵ���;
    '       strDelLost-�˷�ʧ�ܵĵ���
    '       bytType-0-ҽ��ʧ��;1-һ��ͨʧ��;2-����������ʧ��;3-ҽ���˷ѳɹ�,������������δ����
    '����:���˺�
    '����:���Է���true,���򷵻�False
    '����:2012-01-11 13:58:53
    '˵��:�ж���ʾ,����ӡ�������˷Ѻ��ٴ�ӡ���Լ�ѡ���ش�
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If strDelSucceedNos = "" Then Exit Function
    
    If bytType = 1 Then
        MsgBox "����[" & strDelLost & "]�˷�ʧ�ܡ����ǵ���[" & strDelSucceedNos & "]�ѳɹ�����ҽ���˷�, " & vbCrLf & _
            "   ��һ��ͨ�˷�ʧ��, �����ִ��ʧ�ܵĵ��������˷�", vbExclamation, gstrSysName
        GoTo GoClear:
    ElseIf bytType = 2 Then
        MsgBox "����[" & strDelLost & "]�˷�ʧ�ܡ����ǵ���[" & strDelSucceedNos & "]�ѳɹ�����ҽ���˷�, " & vbCrLf & _
            "   �������ӿڽ����˷�ʧ��, �����ִ��ʧ�ܵĵ��������˷�", vbExclamation, gstrSysName
        GoTo GoClear:
    ElseIf bytType = 3 Then
        MsgBox "����[" & strDelLost & "]�˷�ʧ�ܡ����ǵ���[" & strDelSucceedNos & "]�ѳɹ�����ҽ���˷ѡ�" & vbCrLf & _
            "���������׻�δ�����˷ѣ��������˷ѣ�", vbExclamation, gstrSysName
    Else
        MsgBox "����[" & strDelLost & "]�˷�ʧ�ܡ����ǵ���[" & strDelSucceedNos & "]�ѳɹ��˷ѡ�" & vbCrLf & _
            "����δ��ӡ�����ִ��ʧ�ܵĵ��������˷ѣ�", vbInformation, gstrSysName
    End If
GoClear:
    Call ClearFace
    If txtNO.Visible Then txtNO.SetFocus
End Function

Public Function Getҽ�����㷽ʽ(ByVal strNo As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ�����ݵ�ҽ�����㷽ʽ
    '����:���ؽ��㷽ʽ,�ö��ŷָ�:�����ʻ�,ҽ������...
    '����:���˺�
    '����:2011-08-30 18:54:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strBalance As String
    On Error GoTo errH
    With mrsBalance
         .Filter = "NO='" & strNo & "' and ����=2"
         Do While Not .EOF
            strBalance = strBalance & "," & !���㷽ʽ
            .MoveNext
         Loop
    End With
    If strBalance <> "" Then strBalance = Mid(strBalance, 2)
    Getҽ�����㷽ʽ = strBalance
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get�������׽��㷽ʽ(ByVal strNo As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ�����ݵ��������׵���ؽ��㷽ʽ
    '����:���ؽ��㷽ʽ,�ö��ŷָ�:����,һ��ͨ...
    '����:���˺�
    '����:2011-08-30 18:54:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strBalance As String
    On Error GoTo errH
    With mrsBalance
         .Filter = "NO='" & strNo & "' and ����>=5"
         Do While Not .EOF
            strBalance = strBalance & "," & !���㷽ʽ
            .MoveNext
         Loop
    End With
    If strBalance <> "" Then strBalance = Mid(strBalance, 2)
    Get�������׽��㷽ʽ = strBalance
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function Getʵ�ս��(ByVal strNo As String) As Currency
    Dim i As Long
    With vsBill
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("���ݺ�")) = strNo Then Getʵ�ս�� = Getʵ�ս�� + Val(.TextMatrix(i, .ColIndex("ʵ�ս��")))
        Next
    End With
End Function
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

Private Sub txt�˷�ժҪ_Change()
    txt�˷�ժҪ.Tag = ""
End Sub
Private Sub initCardSquareData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������رս��㿨����
    '����:���˺�
    '����:2010-01-05 14:51:23
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


Private Function CheckBillIsAllDels(ByVal strNo As String, Optional ByRef strSel��� As String) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ָ���ĵ����Ƿ�ȫ��ѡ���˷�
    '���:strNO-���ݺ�
    '����:strSel���-����ѡ�е����
    '����:0-ȫ��δѡ��;1-ȫ��ѡ��;2-ѡ����һ����
    '����:���˺�
    '����:2011-01-24 16:43:19
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
        If InStr(1, mstrNOsPatiDel & ",", "," & strNo & ",") > 0 Then
            CheckBillIsAllDels = 2: Exit Function
        End If
        CheckBillIsAllDels = 1: Exit Function
     End If
    CheckBillIsAllDels = 2
End Function
Private Function zlCheckPrepayBack(ByVal lng����ID As Long, ByVal strSelNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ƿ����Ԥ��,�������Ԥ��,���������ȷ��ԭ�����û�ˢ������
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-04-15 14:46:41
    '����:37307
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varTmp As Variant, dblMoney As Double
    Dim strFilter As String, i As Long
    'û��ѡ��ĵ���,����true
    If strSelNos = "" Then zlCheckPrepayBack = True: Exit Function
    If lng����ID = 0 Then zlCheckPrepayBack = True: Exit Function
    If gbytԤ����˷��鿨 = 0 Then zlCheckPrepayBack = True: Exit Function
    varTmp = Split(strSelNos, ","): strFilter = ""
    For i = 0 To UBound(varTmp)
        strFilter = strFilter & " or NO='" & varTmp(i) & "'"
    Next
    strFilter = Mid(strFilter, 4)
    On Error GoTo errHandle
    mrsBalance.Filter = strFilter
    If mrsBalance.RecordCount <> 0 Then mrsBalance.MoveFirst
    Do While Not mrsBalance.EOF
        If Nvl(mrsBalance!����) = 1 Then
            dblMoney = dblMoney + Val(Nvl(mrsBalance!������))
        End If
        mrsBalance.MoveNext
    Loop
    mrsBalance.Filter = 0
    '����:37307
    If dblMoney = 0 Then zlCheckPrepayBack = True: Exit Function
    If Not zlDatabase.PatiIdentify(Me, glngSys, lng����ID, dblMoney, , , , , , , , (gbytԤ����˷��鿨 = 2)) Then Exit Function
    zlCheckPrepayBack = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub ReInitPatiInvoice()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���³�ʼ�����˷�Ʊ��Ϣ
    '����:���˺�
    '����:2011-04-29 14:17:33
    '����:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInvoiceFormat As String
    If mbytMode = 0 Then Exit Sub
    mstrUseType = "": mlngShareUseID = 0: mintInvoiceFormat = 0
    mstrUseType = zl_GetInvoiceUserType(mlng����ID, 0, mintInsure)
    mlngShareUseID = zl_GetInvoiceShareID(mlngModule, mstrUseType)
    mintInvoiceFormat = zl_GetInvoicePrintFormat(mlngModule, mstrUseType, mintOldInvoiceFormat)
    mintInvoicePrint = zl_GetInvoicePrintMode(mlngModule, mstrUseType)
    
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
        End Select
    End If
    zlGetInvoiceGroupUseID = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function



Private Function zlCheckDelValied(ByVal lng�����ID As Long, _
     ByVal strName As String, _
     ByVal bln���ѿ� As Boolean, ByVal strCardNo As String, _
     ByVal strSwapNO As String, strSwapMemo As String, _
     ByRef lng����ID As Long, _
    ByVal dbl�˿��� As Double, Optional bln�쳣���� As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����˷ѽ��׽ӿ�
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-02-08 16:40:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strXMLExend As String
    If lng�����ID = 0 Then zlCheckDelValied = True: Exit Function
    If Not mobjSquare Is Nothing Then
        Call initCardSquareData
    End If
    If mobjSquare Is Nothing Then
    
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
    '       strXMLExpend    XML IN  ��ѡ����:�쳣���������˷�(1)
    '����:�˿�Ϸ�,����true,���򷵻�Flase
    strXMLExend = IIf(bln�쳣����, 1, "")
      If mobjSquare.zlReturnCheck(Me, mlngModule, lng�����ID, bln���ѿ�, strCardNo, _
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
    If mobjSquare.zlBrushCard(Me, mlngModule, Nothing, lng�����ID, bln���ѿ�, mtyPati.����, mtyPati.�Ա�, mtyPati.����, dbl�˷Ѷ�, strBrushCardNo, strbrPassWord, _
        True, True, bln����) = False Then Exit Function
    CheckBrushCard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function CallBackBalanceInterface(ByVal lng�����ID As Long, ByVal bln���ѿ� As Boolean, _
    ByVal strCardNo As String, ByVal strSwapGlideNO As String, ByVal strSwapMemo As String, _
    ByVal str����IDs As String, str����IDs As String, _
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
    Dim str���� As String, str������Ϣ As String, strSwapExtendInfor As String
    Dim varData As Variant, varTemp As Variant, i As Long, cllPro As Collection
    Dim strSQL As String, rsTemp As ADODB.Recordset, strTemp As String
    
    Err = 0: On Error GoTo Errhand:
    If lng�����ID = 0 Then CallBackBalanceInterface = True: Exit Function
    
    If str����IDs <> "" Then str������Ϣ = str������Ϣ & "||3|" & str����IDs
    If str������Ϣ <> "" Then str������Ϣ = Mid(str������Ϣ, 3)
    
    If str����IDs = "" Then
    strSQL = "" & _
    "   Select /*+ RULE */ distinct   A.����ID  " & _
    "   From  ������ü�¼ A,������ü�¼ B,table(f_num2list([1])) P " & _
    "   Where A.NO=B.NO and A.��¼����=1 And A.��¼״̬=2  " & _
    "           And B.����ID=P.Column_Value " & _
    "             "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str����IDs)
    With rsTemp
            str����IDs = ""
            Do While Not .EOF
                str����IDs = str����IDs & "," & Val(Nvl(!����ID))
                .MoveNext
            Loop
        End With
        If str����IDs <> "" Then str����IDs = Mid(str����IDs, 2)
    End If
    If str����IDs = "" Then str����IDs = "0"
    '81489,Ƚ����,2015-1-22,�˷Ѵ������ID
    strSwapExtendInfor = "3|" & str����IDs: strTemp = strSwapExtendInfor
    
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
    If mobjSquare.zlReturnMoney(Me, mlngModule, lng�����ID, bln���ѿ�, strCardNo, str������Ϣ, dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor) = False Then Exit Function
    Set cllUpdate = New Collection: Set cllThreeSwap = New Collection
    Call zlAddUpdateSwapSQL(False, str����IDs, lng�����ID, bln���ѿ�, str����, strSwapGlideNO, strSwapMemo, cllUpdate, 2)
    If strTemp <> strSwapExtendInfor Then
        Call zlAddThreeSwapSQLToCollection(False, str����IDs, lng�����ID, bln���ѿ�, strCardNo, strSwapExtendInfor, cllThreeSwap)
    End If
    CallBackBalanceInterface = True
Errhand:
End Function

Private Function OverFeeDel(ByVal str����IDs As String, ByVal lng����ID As Long, ByRef blnCommited As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����˷��շ�
    '���:str����IDs-����շѵĵ���(����Ϊ���ŵĽ���ID,��Ŀǰֻ��һ�ŵ���)
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-29 14:50:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    If Left(str����IDs, 1) = "," Then str����IDs = Mid(str����IDs, 2)

    On Error GoTo errHandle
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
    gcnOracle.CommitTrans: blnCommited = True
    OverFeeDel = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
    blnCommited = True
End Function
Private Sub ClearBalance()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����������
    '����:���˺�
    '����:2011-11-22 15:59:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsBalance
        .Clear 1: .COLS = 1
'        .Cell(flexcpData, 0, 1, .Rows - 1, .COLS - 1) = ""
'        .Cell(flexcpText, 0, 1, .Rows - 1, .COLS - 1) = ""
        .Editable = flexEDKbdMouse
    End With
End Sub
Private Sub LoadBalanceInfor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ӧ�տ����
    '����:���˺�
    '����:2011-11-22 15:45:46
    '����:43403
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNotNos  As String, lngRow As Long, strFilter As String, str���㷽ʽ As String
    Dim strBalance As String, strȱʡ���㷽ʽ As String
    
    strNotNos = Replace(mstrDelNOs, "'", "")
    lngRow = 0
    mrsBalance.Filter = 0
    If strNotNos <> "" Then
         strFilter = Replace(strNotNos, ",", "' and  NO<>'")
         strFilter = " NO<>'" & strFilter & "'"
         mrsBalance.Filter = strFilter
    End If
    mrsBalance.Sort = "��������,Ӧ����,���㷽ʽ"
    With vsBalance
        .Redraw = flexRDNone
        Call ClearBalance
         If mrsBalance.RecordCount <> 0 Then mrsBalance.MoveFirst
         Do While Not mrsBalance.EOF
            '--����:52530
            If InStr(1, ",1,2,3,4,5,", "," & Val(Nvl(mrsBalance!����)) & ",") = 0 Then
                 '--1Ԥ�� ,2  'ҽ�� ,3, 4 'ҽ�ƿ��ͽ��㿨,5 'һ��ͨ
                 strȱʡ���㷽ʽ = Nvl(mrsBalance!���㷽ʽ, "  ")
            End If
             str���㷽ʽ = Nvl(mrsBalance!���㷽ʽ, "  ")
             If str���㷽ʽ <> strBalance Then
                 strBalance = str���㷽ʽ: .COLS = .COLS + 2
                  .ColAlignment(.COLS - 2) = 7: .ColAlignment(.COLS - 1) = 1
             End If
             
             .TextMatrix(lngRow, .COLS - 2) = strBalance & ":"
             .Cell(flexcpData, lngRow, .COLS - 2) = strBalance
             .TextMatrix(lngRow, .COLS - 1) = Val(.TextMatrix(lngRow, .COLS - 1)) + Nvl(mrsBalance!������, 0)
             .Cell(flexcpData, lngRow, .COLS - 1, lngRow, .COLS - 1) = Val(Nvl(mrsBalance!�Ƿ�����))
             mCurBillType.bln������ȫ�� = Val(Nvl(mrsBalance!�Ƿ�ȫ��)) = 1
             
             '�൥��ʹ�ö��ֽ���ʱ,���ʽ����û�н��зֱҴ���,���Բ�����formatȡ��λ��
             .ColData(.COLS - 2) = "ժҪ:" & mrsBalance!ժҪ
             .ColData(.COLS - 1) = "�������:" & mrsBalance!�������
             
             If mrsBalance!�������� <> 1 Then
                .Cell(flexcpForeColor, lngRow, .COLS - 1, lngRow, .COLS - 2) = vbBlue
                .Cell(flexcpForeColor, 1, .COLS - 1, 1, .COLS - 2) = vbRed
                .Cell(flexcpFontBold, 1, .COLS - 1, 1, .COLS - 2) = True    '����
            End If
             mrsBalance.MoveNext
            .Redraw = flexRDBuffered
         Loop
         vsBalance.AutoSizeMode = flexAutoSizeColWidth
         Call vsBalance.AutoSize(0, .COLS - 1)
         
         If mblnSingleBlance And strȱʡ���㷽ʽ <> "" Then
            mblnNotClick = True
            zlControl.CboSetText cbo�˿ʽ, strȱʡ���㷽ʽ
            mblnNotClick = False
         End If
         If mbytMode = 0 Then
            .RowHidden(1) = True: ControlResize
         End If
    End With
End Sub
Private Sub LoadPartԤ���(ByVal strNos As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ز�����δ�ѵ�Ԥ���
    '����:���˺�
    '����:2011-12-01 11:26:48
    '����:43403
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varNos  As Variant, i As Long, strNo As String, j As Long, k As Long
    Dim dblѡ��ϼ� As Double
    If strNos = "" Then Exit Sub
    If Left(strNos, 1) = "," Then strNos = Mid(strNos, 2)
    'δѡ��,�����ܴ��ڲ���ѡ��,��ֻ����һ��Ԥ����,��ֻ����Ԥ���
    varNos = Split(strNos, ",")
    With vsBill
        For i = 0 To UBound(varNos)
            strNo = varNos(i)
            mrsBalance.Filter = " NO='" & strNo & "' and ����<>1 "
            If mrsBalance.RecordCount = 0 Then
                'ֻ��һ�ֽ��㷽ʽ
                k = vsBill.FindRow(strNo, , .ColIndex("���ݺ�"))
                For j = k To vsBill.Rows - 1
                    If vsBill.TextMatrix(j, .ColIndex("���ݺ�")) <> strNo Then Exit For
                    If Val(vsBill.TextMatrix(j, .ColIndex("ѡ��"))) <> 0 Then
                        dblѡ��ϼ� = RoundEx(dblѡ��ϼ� + Val(vsBill.TextMatrix(j, .ColIndex("ʵ�ս��"))), 6)
                    End If
                Next
            End If
        Next
    End With
    With vsBalance
        For i = 1 To .COLS - 1 Step 2
            If .Cell(flexcpData, 0, i) = "Ԥ���" Then
                If Trim(.TextMatrix(1, i)) = "" Then
                    .Cell(flexcpData, 1, i) = "Ԥ���"
                    .TextMatrix(1, i) = "Ԥ���"
                End If
                .TextMatrix(1, i + 1) = Val(.TextMatrix(1, i + 1)) + dblѡ��ϼ�
                .Cell(flexcpData, 1, i + 1) = Val(.TextMatrix(1, i + 1))
                .Cell(flexcpFontBold, 1, i + 1) = True
                .Cell(flexcpForeColor, 1, i + 1) = vbRed
                .RowHidden(1) = False
                Exit For
            End If
        Next
        txt�˿���.Tag = dblѡ��ϼ�
    End With
End Sub
Private Sub LoadDelBalanceInfor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����˿���㷽ʽ
    '����:���˺�
    '����:2011-11-22 15:45:46
    '����:43403
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strBalance As String, varNO As Variant, i As Long, str���� As String
    Dim strSelNo As String, blnȫѡ As Boolean, blnδѡ As Boolean
    Dim strFilter As String, bln�˿� As Boolean, intCol As Integer
    Dim str���㷽ʽ As String, strNotNos As String
    Dim bln����ѡ�� As Boolean, lngRow As Long
    Dim blnYb As Boolean, strPartNO As String, varPatiNo As Variant
    Dim strTemp As String
    Dim blnԤ�� As Boolean
    Dim blnThreeSwap As Boolean
    Dim strTempBalance As String
    Dim str��ͨ���� As String
    
    Err = 0: On Error GoTo Errhand:
    strNotNos = Replace(mstrDelNOs, "'", "")
    lngRow = 1
    If mstrNOs = "" Then Exit Sub
    varNO = Split(mstrNOs, ",")
    blnThreeSwap = False
    blnȫѡ = True: blnδѡ = True: strSelNo = ""
    strPartNO = ""
    For i = 0 To UBound(varNO)
       Select Case CheckBillIsAllDels(varNO(i))
       Case 0   'δѡ��
            blnȫѡ = False
       Case 1   'ȫѡ��
            strSelNo = strSelNo & "," & varNO(i): blnδѡ = False
       Case Else    '����ѡ��
            blnȫѡ = False: blnδѡ = False: bln����ѡ�� = True
            strPartNO = strPartNO & "," & varNO(i)
       End Select
    Next
    
    'δѡ��
    vsBalance.RowHidden(lngRow) = bln����ѡ��
    If bln����ѡ�� And mCurBillType.bln����ҽ�ƿ����� And mCurBillType.blnSingleBalance Then
        With vsBalance
            .RowHidden(lngRow) = mCurBillType.bln������ȫ��
            .Redraw = flexRDNone
            .TextMatrix(1, 1) = .TextMatrix(0, 1)
            .Cell(flexcpData, 1, 1) = .Cell(flexcpData, 0, 1)
            .TextMatrix(1, 2) = .Cell(flexcpData, 1, 2)
            .ColData(2) = .Cell(flexcpData, 0, 2) '�Ƿ�����
            Call ControlResize
            .Redraw = flexRDDirect
        End With
        Exit Sub
    End If
    If strSelNo = "" Then
        vsBalance.Redraw = flexRDNone
        Call LoadPartԤ���(strPartNO)
        If blnδѡ And vsBalance.COLS > 1 Then
            vsBalance.Cell(flexcpText, 1, 1, 1, vsBalance.COLS - 1) = ""
            vsBalance.Cell(flexcpData, 1, 1, 1, vsBalance.COLS - 1) = ""
        End If
        Call ControlResize
        vsBalance.Redraw = flexRDDirect
        Exit Sub
    End If
       
    strSelNo = Mid(strSelNo, 2): strTempBalance = ""
    strFilter = Replace(strSelNo, ",", "' Or NO='")
    strFilter = " NO='" & strFilter & "'"
    mrsBalance.Filter = strFilter
    mrsBalance.Sort = "��������,Ӧ����,���㷽ʽ"
    With vsBalance
        .Redraw = flexRDNone
        .Cell(flexcpData, 1, 0, 1, .COLS - 1) = ""
        If .COLS - 1 > 0 Then
            .Cell(flexcpText, 1, 1, 1, .COLS - 1) = ""
        End If
         If mrsBalance.RecordCount <> 0 Then mrsBalance.MoveFirst
         intCol = 1: strBalance = ""
         Do While Not mrsBalance.EOF
             str���㷽ʽ = Nvl(mrsBalance!���㷽ʽ, "  ")
             If str���㷽ʽ <> strBalance Then
                 For intCol = 1 To .COLS - 1 Step 2
                    If .Cell(flexcpData, 0, intCol) = str���㷽ʽ Then
                        Select Case Val(Nvl(mrsBalance!����))
                        Case 1  'Ԥ��
                            blnԤ�� = True
                            Exit For
                        Case 2 'ҽ��
                             '������ֽ��㷽ʽ��֧�ֻ���,Ҫ��Ϊ�ֽ�,���ü�ȥ
                            If mblnYB�������� Then
                                If gclsInsure.GetCapability(support�����������, mlng����ID, mintInsure, str���㷽ʽ) Then Exit For
                                 blnYb = True
                            Else     '��֧�������������ʱ,ֻ���������Ϊ�ֽ�,����ԭ����,������ҽ������
                                If str���㷽ʽ <> mstr�����ʻ� Then Exit For
                                blnYb = True
                            End If
                        Case 3, 4 'ҽ�ƿ��ͽ��㿨
                            blnThreeSwap = True
                            Exit For
                        Case 5 'һ��ͨ
                            blnThreeSwap = True
                            Exit For
                        Case Else
                            str��ͨ���� = str��ͨ���� & "," & str���㷽ʽ
                            If mBillDelType = EM_����ȫ�� Then
                                '49155:���� 'ȫ��,����������
'                                If InStr(",1,", "," & mrsBalance!�������� & ",") = 0 Then Exit For
                                 Exit For
                            End If
                            '������ȫ��ʱ
                            If mBillDelType = EM_����ȫ�� Then Exit For
                            strTempBalance = strTempBalance & "," & str���㷽ʽ
                            If InStr(",1,2,", "," & mrsBalance!�������� & ",") = 0 Then Exit For
                            If blnȫѡ And mstrDelNOs = "" Then Exit For
                        End Select
                    End If
                 Next
                  strBalance = str���㷽ʽ
             End If
             If intCol < .COLS - 1 Then
                .TextMatrix(lngRow, intCol) = strBalance & ":"
                .TextMatrix(lngRow, intCol + 1) = Val(.TextMatrix(lngRow, intCol + 1)) + Nvl(mrsBalance!������, 0)
                .Cell(flexcpData, lngRow, intCol, lngRow, intCol) = strBalance
                .Cell(flexcpData, lngRow, intCol + 1, lngRow, intCol + 1) = .TextMatrix(lngRow, intCol + 1)
                .ColData(intCol + 1) = Val(Nvl(mrsBalance!�Ƿ�����))
            End If
             mrsBalance.MoveNext
         Loop
         
         If blnYb Then
            For i = 1 To .COLS - 1 Step 2
               If InStr(strTempBalance & ",", "," & .Cell(flexcpData, lngRow, i)) > 0 Then
                   .TextMatrix(lngRow, i) = "": .Cell(flexcpData, lngRow, i) = ""
                   .TextMatrix(lngRow, i + 1) = "": .Cell(flexcpData, lngRow, i + 1) = ""
               End If
            Next
         End If
         If blnThreeSwap And (mCurBillType.bln���Ų����˷� Or Not blnȫѡ) And bln����ѡ�� = False Then
            For i = 1 To .COLS - 1 Step 2
               If InStr(str��ͨ���� & ",", "," & .Cell(flexcpData, lngRow, i)) > 0 Then
                   .TextMatrix(lngRow, i) = "": .Cell(flexcpData, lngRow, i) = ""
                   .TextMatrix(lngRow, i + 1) = "": .Cell(flexcpData, lngRow, i + 1) = ""
               End If
            Next
         End If
        Call LoadPartԤ���(strPartNO)
        If .COLS - 1 > 0 Then
            .Cell(flexcpForeColor, lngRow, 1, lngRow, .COLS - 1) = vbRed
        End If
        .Cell(flexcpFontBold, lngRow, 0, lngRow, .COLS - 1) = True  '����
         Call vsBalance.AutoSize(0, .COLS - 1)
         If .COLS - 1 > 0 Then
            .Row = .FixedRows: .Col = .FixedCols
        End If
        
        .RowHidden(lngRow) = (bln����ѡ�� Or mCurBillType.bln���Ų����˷� _
                            And Not (mCurBillType.bln����ҽ�ƿ����� And mCurBillType.blnSingleBalance)) And blnԤ�� = False
         
        If Not mblnSingleBlance Then
            '���ǵ��ֽ��㷽ʽ
            If mintInsure = 0 And Not mCurBillType.bln���ڿ����� Then
                .RowHidden(lngRow) = .RowHidden(lngRow) Or Not blnȫѡ
            End If
              
        End If
         ControlResize
        .Redraw = flexRDBuffered
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub InitRecErrCurStruct(ByRef rsErr As ADODB.Recordset)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ�쳣���ݵ����ݽṹ
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2012-01-16 15:09:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set rsErr = New ADODB.Recordset
    rsErr.Fields.Append "NO", adVarChar, 20, adFldIsNullable
    rsErr.Fields.Append "������", adDouble, , adFldIsNullable
    rsErr.CursorLocation = adUseClient
    rsErr.LockType = adLockOptimistic
    rsErr.CursorType = adOpenStatic
    rsErr.Open
End Sub

Private Sub zlALLNosBack()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���е���һ����
    '����:���˺�
    '����:2011-11-24 09:53:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varNO As Variant, i As Long
    Dim dblToTal As Double, strNo As String
    Dim dblMoney As Double, intCol As Integer
    Dim bln���� As Boolean, bln�ֽ���� As Boolean
    Dim cur�˷Ѻϼ� As Double, cur����� As Double
    Dim dblTemp As Double, dbl���ϼ� As Double
    Dim rsErr As ADODB.Recordset
    Dim blnԭ���� As Boolean
    bln�ֽ���� = False
    If cbo�˿ʽ.ListIndex <> -1 Then
        If cbo�˿ʽ.ItemData(cbo�˿ʽ.ListIndex) = 1 Then
            bln�ֽ���� = True
        End If
    End If
    
    Call InitRecErrCurStruct(rsErr)
    varNO = Split(mstrNOs, ",")

    For i = 0 To UBound(varNO)
        strNo = CStr(varNO(i))
        mcolError.Add 0, "_" & strNo
    Next
    
    If cbo�˿ʽ.ListIndex = -1 And cbo�˿ʽ.ListCount > 0 Then cbo�˿ʽ.ListIndex = 0
    cbo�˿ʽ.Enabled = False
    cbo�˿ʽ.Locked = True
    
    dblToTal = 0
    ''������ҽ����Ԥ������
    '����:1-Ԥ���,2-ҽ��,3-ҽ�ƿ�,4-���㿨,5-һ��ͨ,0-������

    mrsBalance.Filter = 0 '  "����<>1"
    mrsBalance.Sort = "���� DESC"
    blnԭ���� = True
    With mrsBalance
         If .RecordCount <> 0 Then .MoveFirst
         Do While Not .EOF
            bln���� = False
            Select Case Val(Nvl(!����))
            Case 1 'Ԥ��
            Case 2
                '49155:����
                If mblnYB�������� Then
                    If Not gclsInsure.GetCapability(support�����������, mlng����ID, mintInsure, Nvl(!���㷽ʽ)) Then
                        bln���� = True: blnԭ���� = False
                    End If
                Else
                    '��֧�������������ʱ,ֻ���������Ϊ�ֽ�,����ԭ����,������ҽ������
                    If !���㷽ʽ = mstr�����ʻ� Then
                        bln���� = True: blnԭ���� = False
                    End If
                End If
            Case 3, 4, 5    '3-ҽ�ƿ�,4-���㿨,5-һ��ͨ
                If Val(Nvl(!�Ƿ�����)) = 1 Then
                        For intCol = 1 To vsBalance.COLS - 1 Step 2
                            If vsBalance.Cell(flexcpData, 1, intCol) = !���㷽ʽ Then
                                If Val(vsBalance.TextMatrix(1, intCol + 1)) = 0 And vsBalance.RowHidden(1) = False Then
                                    bln���� = True: blnԭ���� = False: Exit For
                                End If
                            End If
                        Next
                End If
            Case Else
                '49155:����
               ' If Val(Nvl(mrsBalance!��������)) = 1 Then
               If !���㷽ʽ = zlStr.NeedName(cbo�˿ʽ) And blnԭ���� = False Then
                    bln���� = True
               End If
        
                For intCol = 1 To vsBalance.COLS - 1 Step 2
                    If 1 = 1 _
                        And vsBalance.Cell(flexcpData, 1, intCol) = !���㷽ʽ Then
                        If bln���� And Not blnԭ���� Then
                            vsBalance.TextMatrix(1, intCol) = ""
                            vsBalance.TextMatrix(1, intCol + 1) = ""
                        ElseIf vsBalance.TextMatrix(1, intCol) = "" Then
                            vsBalance.TextMatrix(1, intCol) = vsBalance.Cell(flexcpData, 1, intCol) & ":"
                            vsBalance.TextMatrix(1, intCol + 1) = vsBalance.Cell(flexcpData, 1, intCol + 1)
                        End If
                    End If
                Next
            End Select
            
            If bln���� Then
                rsErr.Find "NO='" & Nvl(!NO) & "'"
                If rsErr.EOF Then rsErr.AddNew
                rsErr!NO = CStr(Nvl(!NO))
                rsErr!������ = RoundEx(Val(Nvl(rsErr!������)) + Val(Nvl(!������)), 6)
                rsErr.Update
                dblMoney = RoundEx(dblMoney + !������, 6)
            End If
            dblToTal = RoundEx(dblToTal + !������, 6)
             .MoveNext
        Loop
    End With
    
    dbl���ϼ� = 0: dblMoney = 0:   dblTemp = 0
    rsErr.Filter = 0
    With rsErr
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            dblTemp = dblTemp + Val(Nvl(!������))
            cur�˷Ѻϼ� = cur�˷Ѻϼ� + Format(Val(Nvl(!������)), "0.00")
            .MoveNext
        Loop
    End With
    
    If bln�ֽ���� And dblTemp <> 0 Then
          dblTemp = dblTemp - Get����(mstrNOs)
          If mintInsure > 0 Then
              If gclsInsure.GetCapability(support�ֱҴ���, mlng����ID, mintInsure) Then
                  cur�˷Ѻϼ� = CentMoney(dblTemp)
              End If
          Else
              cur�˷Ѻϼ� = CentMoney(dblTemp)
          End If
    End If
    cur����� = RoundEx(cur�˷Ѻϼ� - dblTemp, 6)
    dbl���ϼ� = dbl���ϼ� + cur�����
    dblMoney = dblMoney + cur�˷Ѻϼ�
    mcolError.Remove "_" & strNo
    mcolError.Add dbl���ϼ�, "_" & strNo

    txt�˿���.ToolTipText = ""
    If dbl���ϼ� <> 0 Then
        txt�˿���.ToolTipText = "�˷������:" & Format(dbl���ϼ�, gstrDec)
    End If

    txt�˿�ϼ�.ToolTipText = txt�˿���.ToolTipText
    txt�˿���.Text = Format(dblMoney, "0.00")
    txt�˿�ϼ�.Text = Format(dblToTal, "0.00")
    
    If mBillDelType = EM_����ȫ�� And blnԭ���� Then dblMoney = 0      'ԭ����
         
    '���ý���λ��,���ֽ����ʱ����ֱ�
    cbo�˿ʽ.Locked = dblMoney = 0
    cbo�˿ʽ.Enabled = dblMoney <> 0
    cbo�˿ʽ.Visible = dblMoney <> 0
End Sub

Private Sub ReCalcDelMoney()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���¼����˷ѽ��
    '����:���˺�
    '����:2011-11-22 16:50:38
    '˵��:���ݵ�ǰ�����˷�ѡ������������˿���������
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cur���ݺϼ� As Currency, curѡ��ϼ� As Currency
    Dim cur�˷Ѻϼ� As Currency, cur����� As Currency, cur���ϼ� As Currency
    Dim bln��ȫ�˷� As Boolean, bln�ֽ���� As Boolean
    Dim curTotal As Currency, strNo As String
    Dim i As Long, j As Long, k As Long, blnԭ���� As Boolean
    Dim colAllReturn As Collection, bln���� As Boolean
    Dim intCol As Long, bln���� As Boolean
    Dim blnδѡ As Boolean, varNO As Variant
    Dim blnȫ�� As Boolean, blnFind As Boolean
    Dim dbl�˿�ϼ� As Double, dblBalanceSum As Double
    Dim dblCashSum As Double '�ֽ�ϼ�
    Dim blnHaveSelected As Boolean, blnHaveNotSelected As Boolean
    
    If mbytMode = 0 Then Exit Sub
    If mrsBalance Is Nothing Then Exit Sub
        
    Set mcolError = New Collection
    Set colAllReturn = New Collection
    
       
    If mBillDelType = EM_����ȫ�� Then
        '���ŵ���һ����,�����
        Call zlALLNosBack: curTotal = Val(txt�˿�ϼ�): GoTo GoSetVisible: Exit Sub
    End If
    
    '��һ���˷�
   varNO = Split(mstrNOs, ",")
    
    '1.���ж������Ƿ���ԭ����,�Ծ����Ƿ���ý��㷽ʽѡ��,�Լ��ֱ���������
    blnԭ���� = True: bln���� = False: bln���� = False
    blnȫ�� = True
    dblBalanceSum = 0: dblCashSum = 0
    For i = 0 To UBound(varNO)
        strNo = CStr(varNO(i))
        cur���ݺϼ� = 0: curѡ��ϼ� = 0
        blnHaveNotSelected = False
        With vsBill
            k = vsBill.FindRow(strNo, , .ColIndex("���ݺ�"))
            For j = k To vsBill.Rows - 1
                If vsBill.TextMatrix(j, .ColIndex("���ݺ�")) <> strNo Then Exit For
                cur���ݺϼ� = cur���ݺϼ� + Val(vsBill.TextMatrix(j, .ColIndex("ʵ�ս��")))
                If Val(vsBill.TextMatrix(j, .ColIndex("ѡ��"))) <> 0 Then
                    curѡ��ϼ� = curѡ��ϼ� + Val(vsBill.TextMatrix(j, .ColIndex("ʵ�ս��")))
                    blnHaveSelected = True
                Else
                    blnHaveNotSelected = True
                End If
            Next
        End With
        
        bln��ȫ�˷� = BillDeleteAll(strNo, 1, mblnHaveExcuteData)
        bln��ȫ�˷� = bln��ȫ�˷� And Not BillExistDelete(strNo, 1) And (cur���ݺϼ� = curѡ��ϼ� And blnHaveNotSelected = False) '����ã�blnHaveNotSelected

        If mCurBillType.bln���ڿ����� And mCurBillType.bln���ֽ��㷽ʽ And bln��ȫ�˷� Then
        
        ElseIf mCurBillType.bln���ڿ����� = False And mintInsure = 0 And bln��ȫ�˷� Then
            If InStr(mstrNOs, ",") > 0 Then
                bln��ȫ�˷� = Not mCurBillType.bln���Ų����˷�
            End If
        End If
                
        If blnȫ�� And Not bln��ȫ�˷� Then blnȫ�� = False
        
        colAllReturn.Add Array(IIf(bln��ȫ�˷�, 1, 0), strNo, cur���ݺϼ�, curѡ��ϼ�), "_" & strNo    '�������ں�����ж�
        If Not bln��ȫ�˷� Then blnԭ���� = False
        
        dbl�˿�ϼ� = RoundEx(dbl�˿�ϼ� + curѡ��ϼ�, 5)
        If bln��ȫ�˷� Then
            mrsBalance.Filter = "NO='" & strNo & "'"
            With mrsBalance
                If .RecordCount <> 0 Then .MoveFirst
                Do While Not .EOF
                    If Val(Nvl(!����)) = 2 Then 'ҽ��
                        If mblnYB�������� Then
                            If Not gclsInsure.GetCapability(support�����������, mlng����ID, mintInsure, !���㷽ʽ) Then
                               blnԭ���� = False
                            End If
                        ElseIf !���㷽ʽ = mstr�����ʻ� Then
                             blnԭ���� = False
                        End If
                    ElseIf InStr("3,4,5", Val(Nvl(!����))) > 0 Then
                        'һ��ͨ���
                        If Val(Nvl(!�Ƿ�����)) = 1 Then
                            For intCol = 1 To vsBalance.COLS - 1 Step 2
                                If vsBalance.Cell(flexcpData, 1, intCol) = !���㷽ʽ And Val(vsBalance.TextMatrix(1, intCol + 1)) = 0 Then
                                    bln���� = True: blnԭ���� = False
                                End If
                            Next
                        End If
                    Else
                        If mCurBillType.bln���Ų����˷� Then blnԭ���� = False  '55675
                        bln���� = True
                    End If
                    dblBalanceSum = dblBalanceSum + Nvl(!������, 0)
                    If !���㷽ʽ = mstr�ֽ���㷽ʽ And Val(Nvl(!����)) <> 1 Then dblCashSum = dblCashSum + Nvl(!������, 0)
                    .MoveNext
            Loop
            End With
        End If
    Next
    
    If blnȫ�� And mstrDelNOs <> "" Then blnȫ�� = False
    
    '�շ�ʱȫ����Ԥ��(���㷽ʽΪ��),�˷�ʱ,������ָ���˷ѷ�ʽ
    '����:1-Ԥ���,2-ҽ��,3-ҽ�ƿ�(һ��ͨ),4-���㿨,5-һ��ͨ,0-������
    mrsBalance.Filter = "����<>1"
    If mrsBalance.RecordCount = 0 Then blnԭ���� = True
    mrsBalance.Filter = 0
   ' If mBillDelType = EM_����ȫ�� Then blnԭ���� = True
         
 
    '��Ҫȷ�����ŵ����еĵ�����
    If blnԭ���� Then
        '���ܴ��ڲ����˵����
        If dblBalanceSum <> dbl�˿�ϼ� Then blnԭ���� = False
    End If
    
    txt�˿�ϼ�.Text = Format(dbl�˿�ϼ�, "0.00")
    If blnԭ���� Then
        '���ܴ��ڶ൥���н����������һ�ŵ���,��ɵ����˷�ʱ,�ֽ���������
         If mintInsure > 0 Then
            If gclsInsure.GetCapability(support�ֱҴ���, mlng����ID, mintInsure) Then
                cur���ݺϼ� = CentMoney(dblCashSum)
            Else
                cur���ݺϼ� = Format(dblCashSum, "0.00")
            End If
        Else
            cur���ݺϼ� = CentMoney(dblCashSum)
        End If
        If cur���ݺϼ� <> dblCashSum Then blnԭ���� = False
    End If
    
    If blnԭ���� Then
        zlControl.CboSetIndex cbo�˿ʽ.hWnd, mintReturnMode
    End If
    cbo�˿ʽ.Enabled = Not blnԭ����
    cbo�˿ʽ.Locked = blnԭ����
    '2.�����˿�����
    If cbo�˿ʽ.ListIndex <> -1 Then
        If cbo�˿ʽ.ItemData(cbo�˿ʽ.ListIndex) = 1 Then
            bln�ֽ���� = True
        End If
    End If
    Dim varTemp As Variant
    
    For i = 1 To colAllReturn.Count
        '0-�Ƿ���ȫ�˷�;1-NO,2-���ݺϼ�,3-ѡ��ϼ�
        varTemp = colAllReturn(i)
        strNo = varTemp(1)
        cur���ݺϼ� = Val(varTemp(2)): curѡ��ϼ� = Val(varTemp(3))
        cur�˷Ѻϼ� = 0: cur����� = 0
        '��ȫ�˷�ʱ�ſ�ҽ�����㼰��Ԥ�����
        bln��ȫ�˷� = IIf(Val(varTemp(0)) = 1, True, False)
        
        If bln��ȫ�˷� Or blnȫ�� Then
            mrsBalance.Filter = "NO='" & strNo & "'"
            '����:1-Ԥ���,2-ҽ��,3-ҽ�ƿ�,4-���㿨,5-һ��ͨ,0-������
            With mrsBalance
                Do While Not .EOF
                    Select Case Val(Nvl(!����))
                    Case 1 'Ԥ����
                         curѡ��ϼ� = curѡ��ϼ� - Nvl(!������, 0)
                    Case 2 'ҽ��
                        '������ֽ��㷽ʽ��֧�ֻ���,Ҫ��Ϊ�ֽ�,���ü�ȥ
                        If mblnYB�������� Then
                            If gclsInsure.GetCapability(support�����������, mlng����ID, mintInsure, !���㷽ʽ) Then
                                curѡ��ϼ� = curѡ��ϼ� - Nvl(!������, 0)
                            End If
                        Else     '��֧�������������ʱ,ֻ���������Ϊ�ֽ�,����ԭ����,������ҽ������
                            If !���㷽ʽ <> mstr�����ʻ� Then
                                curѡ��ϼ� = curѡ��ϼ� - Nvl(!������, 0)
                            End If
                        End If
                    Case 3, 4 'ҽ�ƿ��ͽ��㿨
                            If Val(Nvl(!�Ƿ�����)) = 0 Then
                                curѡ��ϼ� = curѡ��ϼ� - Nvl(!������, 0)
                            End If
                            If Val(Nvl(!�Ƿ�����)) = 1 Then
                                For intCol = 1 To vsBalance.COLS - 1 Step 2
                                    If vsBalance.Cell(flexcpData, 1, intCol) = !���㷽ʽ Then
                                        If Val(vsBalance.TextMatrix(1, intCol + 1)) <> 0 And vsBalance.RowHidden(1) = False Then
                                            curѡ��ϼ� = curѡ��ϼ� - Nvl(!������, 0)
                                        End If
                                    End If
                                Next
                            End If
                    Case 5 'һ��ͨ
                            curѡ��ϼ� = curѡ��ϼ� - Nvl(!������, 0)
                    Case Else
                        blnFind = False
                        For intCol = 1 To vsBalance.COLS - 1 Step 2
                            If vsBalance.Cell(flexcpData, 1, intCol) = zlStr.NeedName(cbo�˿ʽ.Text) _
                                And vsBalance.Cell(flexcpData, 1, intCol) = !���㷽ʽ Then
                                If bln���� Or Not blnԭ���� Then    'Or Not blnԭ����:55187
                                    vsBalance.TextMatrix(1, intCol) = ""
                                    vsBalance.TextMatrix(1, intCol + 1) = ""
                                ElseIf vsBalance.TextMatrix(1, intCol) = "" Then
                                    vsBalance.TextMatrix(1, intCol) = vsBalance.Cell(flexcpData, 1, intCol) & ":"
                                    vsBalance.TextMatrix(1, intCol + 1) = vsBalance.Cell(flexcpData, 1, intCol + 1)
                                End If
                                blnFind = True: Exit For
                            End If
                        Next
                        
                        blnFind = False
                        For intCol = 1 To vsBalance.COLS - 1 Step 2
                            If zlStr.NeedName(cbo�˿ʽ.Text) <> vsBalance.Cell(flexcpData, 1, intCol) _
                                And vsBalance.TextMatrix(1, intCol) = "" _
                                And vsBalance.Cell(flexcpData, 1, intCol) <> "" Then
                                    vsBalance.TextMatrix(1, intCol) = vsBalance.Cell(flexcpData, 1, intCol) & ":"
                                    vsBalance.TextMatrix(1, intCol + 1) = vsBalance.Cell(flexcpData, 1, intCol + 1)
                            End If
                            If vsBalance.Cell(flexcpData, 1, intCol) = !���㷽ʽ And vsBalance.TextMatrix(1, intCol) <> "" Then
                                blnFind = True: Exit For
                            End If
                        Next
                        If blnȫ�� And blnFind Then
                            curѡ��ϼ� = curѡ��ϼ� - Nvl(!������, 0)
                        End If

                    End Select
                    .MoveNext
                Loop
            End With
        Else
            '�����˷ѣ�����Ƿ񲿷��˷�
            mrsBalance.Filter = "NO='" & strNo & "'  and ����<>1 "
            If mrsBalance.RecordCount = 0 Then
                curѡ��ϼ� = 0
            Else
              
'                '���ܴ���Ԥ�������������˷�,���,��Ҫ�ų�������
'                If vsBalance.RowHidden(1) = False Then
'                    For intCol = 1 To vsBalance.COLS - 1 Step 2
'                        If vsBalance.Cell(flexcpData, 1, intCol) <> "" Then
'                            curѡ��ϼ� = curѡ��ϼ� - Val(vsBalance.TextMatrix(1, intCol + 1))
'                        End If
'                    Next
'                End If
                If mCurBillType.bln����ҽ�ƿ����� And mCurBillType.blnSingleBalance Then
                    vsBalance.Cell(flexcpData, 1, 2) = IIf(blnHaveSelected, dbl�˿�ϼ�, "")
                    If vsBalance.TextMatrix(1, 2) = "" And mCurBillType.bln������ȫ�� = False Then
                        vsBalance.TextMatrix(1, 2) = IIf(blnHaveSelected, FormatEx(dbl�˿�ϼ�, 2), "")
                    End If
                    If Val(vsBalance.TextMatrix(1, 2)) = 0 Or vsBalance.RowHidden(1) Then
                        cbo�˿ʽ.Enabled = True
                        cbo�˿ʽ.Locked = False
                    Else
                        vsBalance.TextMatrix(1, 2) = FormatEx(dbl�˿�ϼ�, 2)
                        cbo�˿ʽ.Enabled = False
                        cbo�˿ʽ.Locked = True
                        curѡ��ϼ� = 0 ' curѡ��ϼ� - dbl�˿�ϼ�
                    End If
                    If cbo�˿ʽ.Visible And cbo�˿ʽ.ListIndex <> -1 Then
                        If cbo�˿ʽ.ItemData(cbo�˿ʽ.ListIndex) = 1 Then
                            bln�ֽ���� = True
                        End If
                    End If
                End If
            End If
            
        End If
        
        '���ý���λ��,���ֽ����ʱ����ֱ�
        If bln�ֽ���� Then
            If mintInsure > 0 Then
                If gclsInsure.GetCapability(support�ֱҴ���, mlng����ID, mintInsure) Then
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
        If Not blnԭ���� Then
            cur����� = cur�˷Ѻϼ� - curѡ��ϼ�
        End If
        
        curTotal = curTotal + cur�˷Ѻϼ�
        mcolError.Add cur�����, "_" & strNo
        cur���ϼ� = cur���ϼ� + cur�����
    Next
    txt�˿���.ToolTipText = "�˷������:" & Format(cur���ϼ�, gstrDec)
    
    txt�˿���.Text = Format(curTotal, "0.00")
    vsBalance.AutoSizeMode = flexAutoSizeColWidth
    Call vsBalance.AutoSize(0, vsBalance.COLS - 1)
    
GoSetVisible:
        Call Show�˿ʽ(cbo�˿ʽ.Enabled And curTotal <> 0)
End Sub

Private Sub ControlResize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����ؼ�λ��
    '����:���˺�
    '����:2011-11-23 14:21:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim blnFind As Boolean
    With vsBalance
        For i = 1 To .COLS - 1 Step 2
            If .Cell(flexcpData, 1, i) <> "" Then
                blnFind = True: Exit For
            End If
        Next
        If blnFind = False Then .RowHidden(1) = True
        .Height = IIf(.RowHidden(1), 375, 735)
    End With
    Form_Resize
End Sub

 
 
Private Sub txtPatient_Change()
    '����:50885
    If Me.ActiveControl Is txtPatient Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
        IDKind.SetAutoReadCard (txtPatient.Text = "")
    End If
End Sub
Private Sub txtPatient_GotFocus()
    '����:50885
    If txtPatient.Locked Or Not txtPatient.Visible Then Exit Sub
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (txtPatient.Text = "")
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (txtPatient.Text = "")
    IDKind.SetAutoReadCard (txtPatient.Text = "")
End Sub
Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim blnCard As Boolean, blnICCard As Boolean, blnCancel As Boolean
    '����:50885
    
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
        '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
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
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FindPati(ByVal objCard As Card, ByVal blnCard As Boolean, ByVal strInput As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���Ҳ���
    '����:���˺�
    '����:2012-09-03 09:46:17
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
        Exit Sub
    End If
    mlng����ID = Val("" & mrsInfo!����ID)
    txtPatient = Nvl(mrsInfo!����)

    lblPati.Caption = "����:" & "                 " & _
        "���Ա�:" & Nvl(mrsInfo!�Ա�) & _
        "������:" & Nvl(mrsInfo!����) & _
        "�������:" & Nvl(mrsInfo!�����) & _
        "���ѱ�:" & Nvl(mrsInfo!�ѱ�) & _
        "�����ʽ:" & mrsInfo!ҽ�Ƹ��ʽ
    With mtyPati
        .����ID = mlng����ID
        .�Ա� = Nvl(mrsInfo!�Ա�)
        .���� = Nvl(mrsInfo!����)
        .���� = Nvl(mrsInfo!����)
    End With
    If SelectNO(mlng����ID) = False Then Exit Sub
    If vsBill.Enabled And vsBill.Visible Then vsBill.SetFocus
End Sub

Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, blnCancel As Boolean, Optional blnCard As Boolean = False) As Boolean
    '���ܣ���ȡ������Ϣ
    '������strInput=[ˢ��]|[A����ID]|[BסԺ��]
    '˵����
    '     1.�����ڲ���Ԥ����
    '     2.�Զ�ʶ������Ժ״̬,����(����ID,��ҳID,����,�Ա�,����,סԺ��,����,��Ժ��־)
    '����:�Ƿ��ȡ�ɹ�,�ɹ�ʱmrsInfo�а���������Ϣ,ʧ��ʱmrsInfo=Close
    '����:50885
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
        strWhere = strWhere & " And A.סԺ��=[1]"
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
                '75087,Ƚ����,2014-7-29,���ﲡ���շ�ʱ,����Ҫ���������������,ֻ��Ҫ��������ŵ����˳��ż����ҵ��������Ĳ�����Ϣ������
                strInput = zlCommFun.GetFullNO(strInput, 3)
            Case "סԺ��"
                If Not IsNumeric(strInput) Then strInput = "0"
                strWhere = strWhere & " And A.סԺ��=[2]"
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
    '76451,Ƚ����,2014-8-19
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
    '75259�����ϴ�,2014-7-10������������ɫ����
    txtPatient.ForeColor = &HC00000
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput)
    If mrsInfo.EOF Then
        Set mrsInfo = New ADODB.Recordset: Exit Function
    End If
    Call SetPatiColor(txtPatient, Nvl(mrsInfo!��������), IIf(IsNull(mrsInfo!����), &HC00000, vbRed))
    GetPatient = True
    Exit Function
errH:
     If ErrCenter() = 1 Then Resume
    Call SaveErrLog
NotFoundPati:
    Set mrsInfo = New ADODB.Recordset
End Function
Private Function SelectNO(ByVal lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݲ���IDѡ����ʵ��˷ѵ���
    '���:
    '����:
    '����: �ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2012-07-04 10:32:40
    '����:50885
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, blnCancel As Boolean
    On Error GoTo errHandle
    '�����:50885
    strSQL = "" & _
        "  With �շѵ� as ( " & _
        "           Select Max(a.ID) as ID,a.No as ���ݺ�,  B.���� as ��������, a.������, a.����Ա���, a.����Ա����, a.ʵ��Ʊ��, To_Char(a.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, " & vbCrLf & _
        "                   To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ�� " & vbCrLf & _
        "           From ������ü�¼ A,���ű� B " & vbCrLf & _
        "           Where a.��¼���� = 1 And nvl(A.���ӱ�־,0)<>9 and A.��������ID=B.ID(+) And a.��¼״̬ =1 " & vbCrLf & _
        "                       And Nvl(a.ִ��״̬, 0) <> 1 And Nvl(a.����״̬, 0) <> 1 And a.����id = [1] " & vbCrLf & _
        "          Group by   a.No,  a.������, B.����,a.����Ա���, a.����Ա����, a.ʵ��Ʊ��, To_Char(a.����ʱ��, 'yyyy-mm-dd hh24:mi:ss'),To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') " & vbCrLf & _
        "           )"
        
     strSQL = strSQL & vbCrLf & _
     "  Select J.*  " & vbCrLf & _
     "  From �շѵ� J," & vbCrLf & _
     "           (Select A.NO,sum(nvl(A.����,1)*nvl(A.����,1)) ����" & vbCrLf & _
     "             From ������ü�¼ A,�շѵ� B  " & vbCrLf & _
     "             Where A.NO=B.���ݺ� And A.��¼����=1 And a.�۸񸸺� is null  " & vbCrLf & _
     "             Group by A.NO " & vbCrLf & _
     "              Having sum(nvl(A.����,1)*nvl(A.����,1))>0 ) M" & vbCrLf & _
     "  Where J.���ݺ�=M.NO " & vbCrLf
     
     strSQL = "Select * From (" & strSQL & ") Order by �Ǽ�ʱ�� desc,���ݺ�"
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "�˷ѵ���", 1, "", "��ѡ����Ҫ�˷ѵĵ���", False, False, False, 0, 0, 0, blnCancel, False, False, lng����ID, "bytSize=1")
    If blnCancel = True Then Exit Function
    If rsTemp Is Nothing Then Exit Function
    If rsTemp.State <> 1 Then Exit Function
    
    Dim strNo As String
    mstrNo = Nvl(rsTemp!���ݺ�)
    mblnOneCard = GetOneCard.RecordCount > 0
    If Not ReadBills(mstrNo) Then
        ClearFace True, True
        Exit Function
    End If
    SelectNO = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub mobjICCard_ShowICCardInfo(ByVal strNo As String)
    '����:50885
    Dim lngPreIDKind As Long
    If Not txtPatient.Locked And txtPatient.Text = "" And Me.ActiveControl Is txtPatient And strNo <> "" Then
        lngPreIDKind = IDKind.IDKind
        mblnNotClick = True
        Dim intIndex As Integer
        intIndex = IDKind.GetKindIndex("IC����")
        If intIndex <= 0 Then mblnNotClick = False: Exit Sub
        IDKind.IDKind = intIndex: Call IDKind_Click(IDKind.GetCurCard)
        txtPatient.Text = strNo
        Call txtPatient_KeyPress(vbKeyReturn)
        If txtPatient.Text = "" Then Call mobjICCard.SetEnabled(False)
        IDKind.IDKind = lngPreIDKind
        mblnNotClick = False
        If Not txtPatient.Locked And Me.ActiveControl Is txtPatient Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
    End If
End Sub
Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthday As Date, ByVal strAddress As String)
    Dim lngPreIDKind As Long
    '����:50885
    If Not txtPatient.Locked And txtPatient.Text = "" And Me.ActiveControl Is txtPatient Then
        mblnNotClick = True
        lngPreIDKind = IDKind.IDKind
        IDKind.IDKind = IDKind.GetKindIndex("���֤��")
        txtPatient.Text = strID
        Call txtPatient_KeyPress(vbKeyReturn)
        IDKind.IDKind = lngPreIDKind
        mblnNotClick = False
    End If
End Sub

Private Sub CloseIDCard()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ر�������������
    '����:���˺�
    '����:2012-03-09 16:26:40
    '����:50885
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
Private Sub NewCardObject()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���µĿ�����
    '����:���˺�
    '����:2012-03-09 16:28:23
    '����:50885
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mbytMode = 0 Then Exit Sub
    
    If mobjIDCard Is Nothing Then
        Set mobjIDCard = New clsIDCard
        Call mobjIDCard.SetParent(Me.hWnd)
    End If
    If mobjICCard Is Nothing Then
        Set mobjICCard = New clsICCard
        Call mobjICCard.SetParent(Me.hWnd)
    End If
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
        .TextMatrix(0, 0) = "��Ʊ��"
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
        Call MergeFixedCol
        .Redraw = flexRDBuffered
    End With
    Exit Sub
Errhand:
    vsInvoice.Redraw = flexRDBuffered
End Sub
Private Sub SetpicInvoiceVisible()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���÷�Ʊ�ؼ�����ʾ
    '����:���˺�
    '����:2013-05-09 11:30:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    picInvoice.Visible = False
    If gTy_Module_Para.bytƱ�ݷ������ = 0 Then GoTo ReSizing:
    If mbytMode <> 1 Then GoTo ReSizing:
    If mrsDelInvoice Is Nothing Then GoTo ReSizing:
    mrsDelInvoice.Filter = 0
    If mrsDelInvoice.RecordCount = 0 Then GoTo ReSizing:
    picInvoice.Visible = True
ReSizing:
    '���µ�����С
    Form_Resize
    picInvoice_Resize
    picPati_Resize
End Sub
Private Sub LoadInvoiceData(ByVal strNos As String, Optional ByVal strInvoiceNO As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ط�Ʊ��Ϣ
    '���:strNos-���ݺ�,����ö��ŷָ�
    '       strInvoiceNo-��Ʊ��(��ָ���ķ�Ʊ�ŷ�Ʊ�Ų���)
    '����:���˺�
    '����:2013-05-07 17:07:38
    '����:25187
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str��� As String, varTemp As Variant
    Dim i As Long, str��Ʊ�� As String
    If gTy_Module_Para.bytƱ�ݷ������ = 0 Then Exit Sub
    If mbytMode <> 1 Then Exit Sub
    If mrsDelInvoice Is Nothing Then
        Set mrsDelInvoice = zlGetFromNoTOInvoice(strNos)
    End If
    If mrsDelInvoice Is Nothing Then Exit Sub
    If mrsDelInvoice.RecordCount = 0 Then Exit Sub
    mrsDelInvoice.Sort = "Ʊ��"
    str��Ʊ�� = ""
    With mrsDelInvoice
        Do While Not .EOF
            If InStr(str��Ʊ�� & ",", "," & Nvl(!Ʊ��) & ",") = 0 Then
                str��Ʊ�� = str��Ʊ�� & "," & Nvl(!Ʊ��)
            End If
            .MoveNext
        Loop
    End With
    If str��Ʊ�� <> "" Then str��Ʊ�� = Mid(str��Ʊ��, 2)
      '���ط�Ʊ��
    varTemp = Split(str��Ʊ��, ",")
    With vsInvoice
        .Clear
        .Rows = 1: .COLS = 1
        .FixedCols = 1
        .TextMatrix(0, 0) = "��Ʊ��"
        .Redraw = flexRDNone
        .COLS = .COLS + UBound(varTemp) + 1
        For i = 0 To UBound(varTemp)
            If i + 1 > .COLS - 1 Then
                .COLS = .COLS + 1
            End If
            .TextMatrix(0, i + 1) = CStr(varTemp(i))
            .Cell(flexcpData, 0, i + 1) = CStr(varTemp(i))
            .Cell(flexcpChecked, 0, i + 1) = 2
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .COLS - 1)
        Call picInvoice_Resize
        Call Form_Resize
        
        .Editable = flexEDKbdMouse
        .Redraw = flexRDBuffered
    End With
End Sub
Private Sub FromNoSelectInvoice()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݵ���ѡ��Ʊ
    '����:���˺�
    '����:2013-05-08 15:52:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str��Ʊ�� As String, strNo As String
    Dim strNos As String, i As Long, j As Long
    If mbytMode <> 1 Or gTy_Module_Para.bytƱ�ݷ������ = 0 Then Exit Sub
    
    On Error GoTo errHandle
    With vsBill
        str��Ʊ�� = ""
        For i = 1 To .Rows - 1
              If Abs(Val(.TextMatrix(i, .ColIndex("ѡ��")))) = 1 Then
                    strNo = .TextMatrix(i, .ColIndex("���ݺ�"))
                    If strNo <> "" Then
                        str��Ʊ�� = str��Ʊ�� & "," & GetFromNumToInvoiceNo(strNo, CStr(.RowData(i)))
                    End If
              End If
        Next
    End With
    With vsInvoice
        For i = 0 To .Rows - 1
            For j = 1 To .COLS - 1
                If InStr(1, str��Ʊ�� & ",", "," & .Cell(flexcpData, i, j) & ",") > 0 Then
                    .Cell(flexcpChecked, i, j) = 1
                ElseIf Trim(.Cell(flexcpData, i, j)) <> "" Then
                    .Cell(flexcpChecked, i, j) = 2
                Else
                End If
            Next
        Next
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
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
Private Sub MergeFixedCol()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ϲ��̶���
    '����:���˺�
    '����:2013-05-08 15:38:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, c As Long
    On Error GoTo errHandle
    If mbytMode <> 1 Or gTy_Module_Para.bytƱ�ݷ������ = 0 Then Exit Sub
    With vsInvoice
        If .FixedCols = 0 Then Exit Sub
        For i = 0 To .Rows - 1
            .MergeRow(c) = True
            For c = 0 To .FixedCols - 1
                .TextMatrix(i, c) = "��Ʊ��"
            Next
        Next
        .MergeCellsFixed = flexMergeRestrictRows
        For c = 0 To .FixedCols - 1
            .MergeCol(c) = True
        Next
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub vsInvoice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strInvoice As String
    With vsInvoice
        strInvoice = Trim(.Cell(flexcpData, Row, Col))
        If strInvoice <> "" Then
            'ͬʱѡ�������Ʊ
            Call SelectRelatingInvoice(strInvoice, Abs(Val(.Cell(flexcpChecked, Row, Col))) = 1)
        End If
    End With
    Call FromAllInvoiceSelectNO
    '��Ҫѡ���й����ջصķ�Ʊ
    Call FromNoSelectInvoice
    Call ShowAndHideDelBillRow
End Sub

Private Sub vsInvoice_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then Cancel = True: Exit Sub
    With vsInvoice
'        Select Case Val(.Cell(flexcpChecked, Row, Col))
'        Case 3
'            Cancel = True
'        End Select
    End With
End Sub

Private Sub SelectRelatingInvoice(ByVal strInvoiceNO As String, ByVal blnSel As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ѡ��ָ����Ʊ�Ĺ�����Ʊ
    '���:strInvoiceNo-��Ʊ��
    '����:���˺�
    '����:2013-05-09 10:41:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim r As Long, c As Long, lng������� As Long
    Dim str��Ʊ�� As String
    On Error GoTo errHandle
    If mbytMode <> 1 Or gTy_Module_Para.bytƱ�ݷ������ = 0 Then Exit Sub
    'ѡ����Ʊ
    With vsInvoice
        For r = 0 To .Rows - 1
            For c = 1 To .COLS - 1
                str��Ʊ�� = .Cell(flexcpData, r, c)
                If str��Ʊ�� = strInvoiceNO Then
                        .Cell(flexcpChecked, r, c) = IIf(blnSel, 1, 2)
                End If
            Next
        Next
    End With
    If mrsDelInvoice Is Nothing Then Exit Sub
    If mrsDelInvoice.RecordCount = 0 Then Exit Sub
    lng������� = 0
    mrsDelInvoice.Filter = "Ʊ��='" & strInvoiceNO & "'"
    If mrsDelInvoice.RecordCount <> 0 Then
        lng������� = Val(Nvl(mrsDelInvoice!����Ʊ�����))
    End If
    If lng������� = 0 Then
        mrsDelInvoice.Filter = 0: Exit Sub
    End If
    mrsDelInvoice.Filter = "����Ʊ�����=" & lng�������
    If mrsDelInvoice.RecordCount = 0 Then
        mrsDelInvoice.Filter = 0: Exit Sub
    End If
    
    With mrsDelInvoice
        .MoveFirst
        Do While Not .EOF
            With vsInvoice
                For r = 0 To .Rows - 1
                    For c = 1 To .COLS - 1
                        str��Ʊ�� = .Cell(flexcpData, r, c)
                        If str��Ʊ�� = Nvl(mrsDelInvoice!Ʊ��) Or str��Ʊ�� = strInvoiceNO Then
                            .Cell(flexcpChecked, r, c) = IIf(blnSel, 1, 2)
                        End If
                    Next
                Next
            End With
            .MoveNext
        Loop
    End With
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub FromInvoiceSelectNO(ByVal strInvoiceNO As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ָ���ķ�Ʊ��,ѡ�����ĵ���
    '����:���˺�
    '����:2013-05-08 16:23:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNo As String, str��� As String
    Dim k As Long, j As Long
    If mbytMode <> 1 Or gTy_Module_Para.bytƱ�ݷ������ = 0 Then Exit Sub
    If mrsDelInvoice Is Nothing Then Exit Sub
    If mrsDelInvoice.RecordCount = 0 Then Exit Sub
    On Error GoTo errHandle
    mrsDelInvoice.Filter = "Ʊ��='" & strInvoiceNO & "'"
    
    With mrsDelInvoice
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            strNo = Nvl(!NO): str��� = "," & Nvl(!���) & ","
              With vsBill
                  k = vsBill.FindRow(strNo, , .ColIndex("���ݺ�"))
                  For j = k To .Rows - 1
                      If .TextMatrix(j, .ColIndex("���ݺ�")) <> strNo Then Exit For
                      If InStr(1, str���, "," & .RowData(j) & ",") > 0 Then
                            .Cell(flexcpChecked, j, .ColIndex("ѡ��")) = 1
                      End If
                      'ͬ��ѡ����������Ŀ
                      Call SynchronizationSelect(j)
                  Next
              End With
             .MoveNext
        Loop
    End With
    mrsDelInvoice.Filter = 0
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
   mrsDelInvoice.Filter = 0
End Sub
Private Sub FromAllInvoiceSelectNO()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������ѡ��ķ�Ʊ��,ѡ�����ĵ���
    '����:���˺�
    '����:2013-05-08 16:23:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str��Ʊ�� As String, i As Long, c As Long, k As Long, j As Long
    Dim strNo As String, str��� As String
    On Error GoTo errHandle
    If mrsDelInvoice Is Nothing Then Exit Sub
    If mrsDelInvoice.RecordCount = 0 Then Exit Sub
    If mbytMode <> 1 Or gTy_Module_Para.bytƱ�ݷ������ = 0 Then Exit Sub
    
    With vsBill
        .Cell(flexcpChecked, .FixedRows, .ColIndex("ѡ��"), .Rows - 1, .ColIndex("ѡ��")) = 0
    End With
    With vsInvoice
        For i = 0 To .Rows - 1
            For c = 1 To .COLS - 1
                str��Ʊ�� = Trim(.Cell(flexcpData, i, c))
                If Abs(Val(.Cell(flexcpChecked, i, c))) = 1 Then
                    Call FromInvoiceSelectNO(str��Ʊ��)
                End If
            Next
        Next
    End With
    '��ʾ��صĽ�����Ϣ
    Call LoadBalanceInfor
    Call LoadDelBalanceInfor
    Call ReCalcDelMoney
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    mrsDelInvoice.Filter = 0
End Sub
Private Sub SynchronizationSelect(ByVal lngRow As Long)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ͬ��ѡ��
    '���:lngRow-��ǰѡ�����
    '����:���˺�
    '����:2013-05-08 16:54:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    On Error GoTo errHandle
    With vsBill
        If mBillDelType = EM_����ȫ�� Then
           .TextMatrix(lngRow, .ColIndex("ѡ��")) = 1
          Exit Sub
        End If
        If mBillDelType = EM_����ȫ�� Then
          Call SetNOBill(.TextMatrix(lngRow, .ColIndex("���ݺ�")), Val(.TextMatrix(lngRow, .ColIndex("ѡ��"))) <> 0)
          Exit Sub
        End If
        '29201
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
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub ShowAndHideDelBillRow()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ�������˷���
    '����:���˺�
    '����:2013-05-09 10:14:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnSeled As Boolean, r As Long, c As Long
    On Error GoTo errHandle
    If mbytMode <> 1 Or gTy_Module_Para.bytƱ�ݷ������ = 0 Then Exit Sub
    If mBillDelType = EM_����ȫ�� Then Exit Sub
    blnSeled = False
    With vsInvoice
        For r = 0 To .Rows - 1
            For c = 1 To .COLS - 1
                    If Trim(.Cell(flexcpData, r, c)) <> "" Then
                        If Abs(Val(.Cell(flexcpChecked, r, c))) = 1 Then
                            blnSeled = True: Exit For
                        End If
                    End If
            Next
            If blnSeled Then Exit For
        Next
    End With
    With vsBill
        '����δѡ�����
        For r = 1 To .Rows - 1
            .RowHidden(r) = False
            If Abs(Val(.Cell(flexcpChecked, r, .ColIndex("ѡ��")))) <> 1 And blnSeled Then
                .RowHidden(r) = True
            End If
        Next
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub OlnyShowSelectedInvoice()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ʾ����ѡ�ķ�Ʊ,δ��ѡ��,ɾ��
    '����:���˺�
    '����:2013-05-09 10:23:34
    '˵��:ֻ��ͨ����Ʊ��ȡ����ʱ��ʾ
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str��Ʊ�� As String, r As Long, c As Long, i As Long
    Dim varTemp As Variant
    If mbytMode <> 1 Or gTy_Module_Para.bytƱ�ݷ������ = 0 Then Exit Sub
    On Error GoTo errHandle
    With vsInvoice
        For r = 0 To .Rows - 1
            For c = 1 To .COLS - 1
                If .Cell(flexcpData, r, c) <> "" Then
                    If Abs(Val(.Cell(flexcpChecked, r, c))) = 1 Then
                        str��Ʊ�� = str��Ʊ�� & "," & .Cell(flexcpData, r, c)
                    End If
                End If
            Next
        Next
        '���ط�Ʊ��
        If str��Ʊ�� = "" Then Exit Sub
        str��Ʊ�� = Mid(str��Ʊ��, 2)
        varTemp = Split(str��Ʊ��, ",")
        .Clear
        .Rows = 1: .COLS = 1
        .FixedCols = 1
        .TextMatrix(0, 0) = "��Ʊ��"
        .Redraw = flexRDNone
        .COLS = .COLS + UBound(varTemp) + 1
        For i = 0 To UBound(varTemp)
            If i + 1 > .COLS - 1 Then
                .COLS = .COLS + 1
            End If
            .TextMatrix(0, i + 1) = CStr(varTemp(i))
            .Cell(flexcpData, 0, i + 1) = CStr(varTemp(i))
            .Cell(flexcpChecked, 0, i + 1) = 1
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .COLS - 1)
        Call picInvoice_Resize
        Call Form_Resize
        .Redraw = flexRDBuffered
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsInvoice_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If vsBill.Enabled Then vsBill.SetFocus
    End If
End Sub

Private Function Get����(ByVal strNos As String) As Double
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡԭʼ���ݵ�����,�Ա��˷�ʱ����
    '���:strNos-���ݺ�(����ö��ŷ���)
    '����:�ɹ����������
    '����:���˺�
    '����:2013-11-29 15:06:11
    '˵��:��Զ൥��ȫ��ʱ,��Ҫ��ȥ����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    strSQL = "" & _
    "   Select /*+ rule */ nvl(Sum(ʵ�ս��),0) as ���� " & _
    "   From ������ü�¼ A,table(f_str2List([1])) J " & _
    "   where A.NO=J.Column_value and A.��¼����=1 And A.��¼״̬ in (1,3) And nvl(A.���ӱ�־,0)=9 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNos)
    Get���� = RoundEx(Val(Nvl(rsTemp!����)), 6)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
