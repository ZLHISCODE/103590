VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#9.0#0"; "zlIDKind.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClinicDelAndView 
   AutoRedraw      =   -1  'True
   Caption         =   "�����˷ѹ���"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11265
   Icon            =   "frmClinicDelAndView.frx":0000
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
      TabIndex        =   33
      Top             =   960
      Width           =   11265
      Begin VB.Frame fraSelectDownSplit 
         Height          =   30
         Left            =   -15
         TabIndex        =   35
         Top             =   900
         Width           =   11535
      End
      Begin VB.Frame fraSelectTopSplit 
         Height          =   45
         Left            =   -30
         TabIndex        =   34
         Top             =   0
         Width           =   11385
      End
      Begin VSFlex8Ctl.VSFlexGrid vsInvoice 
         Height          =   375
         Left            =   300
         TabIndex        =   32
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
         FormatString    =   $"frmClinicDelAndView.frx":058A
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
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   5565
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
         TabIndex        =   6
         Top             =   60
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
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   7404
      Width           =   11265
      Begin VB.CommandButton cmdRefuseApply 
         Caption         =   "�ܾ�(&N)"
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
         Left            =   6300
         TabIndex        =   37
         Top             =   150
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.TextBox txtYB 
         Height          =   300
         Left            =   945
         TabIndex        =   28
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
         TabIndex        =   27
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
         TabIndex        =   13
         Top             =   150
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   12
         Top             =   150
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
      TabIndex        =   14
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
            Picture         =   "frmClinicDelAndView.frx":0648
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
      FormatString    =   $"frmClinicDelAndView.frx":0EDC
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
      TabIndex        =   17
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
         Left            =   10020
         Locked          =   -1  'True
         TabIndex        =   29
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
         Left            =   9000
         TabIndex        =   30
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
         TabIndex        =   10
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
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   0
      Width           =   11265
      Begin VB.PictureBox picPatiBack 
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   540
         ScaleHeight     =   360
         ScaleWidth      =   2640
         TabIndex        =   31
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
            TabIndex        =   3
            ToolTipText     =   "��λ:F6,����:-����ID,*�����,+סԺ��,.�Һŵ���,����:*2536��ʾ������Ų���"
            Top             =   -15
            Width           =   1980
         End
         Begin zlIDKind.IDKindNew IDKind 
            Height          =   360
            Left            =   0
            TabIndex        =   36
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
         TabIndex        =   24
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
         TabIndex        =   22
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
            TabIndex        =   23
            Top             =   0
            Width           =   405
         End
      End
      Begin VB.Frame fraInfo_1 
         Height          =   120
         Left            =   -120
         TabIndex        =   21
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
         TabIndex        =   25
         Top             =   600
         Visible         =   0   'False
         Width           =   480
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
         TabIndex        =   16
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
      FormatString    =   $"frmClinicDelAndView.frx":0F56
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
Attribute VB_Name = "frmClinicDelAndView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public Enum gEM_ChargeDelType
    EM_MULTI_�鿴 = 0
    EM_MULTI_�˷� = 1
    EM_MULTI_�쳣���� = 2
    EM_MULTI_�˷����� = 3
    EM_MULTI_ȡ������ = 4
    EM_MULTI_�˷���� = 5
    EM_MULTI_�ܾ����� = 6
    EM_MULTI_ȡ����� = 7
End Enum
'----------------------------------------------------------------
'�ӿڱ���
Private mstrPrivs As String
Private mbytMode As gEM_ChargeDelType  '0-���ŵ��ݲ鿴,1-���ŵ����˷�,2-���쳣���˷ѵ����������˷�;3-�˷�����
Private mlng������� As Long  'Ҫ�鿴���˷ѵĶ��ŵ����н������
Private mblnNOMoved As Boolean '�����ĵ����Ƿ��ں����ݱ���
Private mstrDelTime As String '�鿴�˷ѵ��ݵĵǼ�ʱ��(yyyy-MM-dd HH:mm:ss) 'ֻ�в鿴�˷ѵ���ʱ�Ŵ���ʱ��,��������������
Private mstrApplyTime As String
'----------------------------------------------------------------
Private mlngModule  As Long
Private mlng����ID As Long
Private mstr�����ʻ� As String   'ҽ�������ʻ�������
Private mdbl������� As Double   '��ǰ���˸����ʻ����,���շ�����
Private mdbl����͸֧ As Double   '�����ʻ�����͸֧���,���շ�����

Private mblnOK As Boolean
Private mlngShareUseID As Long '������������ID
Private mstrUseType As String 'ʹ�����
Private mintInvoiceFormat As Integer  '��ӡ�ķ�Ʊ��ʽ,��Ʊ��ʽ���
Private mintOldInvoiceFormat As Integer '�ɷ�Ʊ��ʽ
Private mintInvoicePrint As Integer '0-����ӡ;1-�Զ���ӡ;2-��ʾ��ӡ
Private mint�˷ѻص���ӡ As Integer '�˷ѻص���ӡ��ʽ 0-����ӡ,1-�Զ���ӡ,2-ѡ���Ƿ��ӡ
Private mintInvoiceFormatDel As Integer  '�˷Ѵ�ӡ�ķ�Ʊ��ʽ,��Ʊ��ʽ���(91998)
Private mintInvoicePrintDel As Integer '0-����ӡ;1-�Զ���ӡ;2-��ʾ��ӡ
Private mblnPrintView As Boolean    '��ӡǰ�鿴����
Private mblnOneCard As Boolean
Private mblnFirst As Boolean
Private mblnNotClick As Boolean
Private mstrTittle As String
Private mstrNo As String 'Ҫ�鿴���˷ѵĶ��ŵ����е�ĳ��NO,�˷�ʱ����û��

Private mrs���㷽ʽ As ADODB.Recordset
Private mrs�շѶ��� As ADODB.Recordset '�շѶ��� :����:33634
Private mrsDelInvoice As ADODB.Recordset
Private mrsBalance As ADODB.Recordset '��¼ÿ�ŵ��ݵĽ������
Private mrsInsureBalance As ADODB.Recordset '��¼ÿ�ŵ��ݵ�ҽ��������ϸ
Private mrsInfo As ADODB.Recordset

Private mstrOnePatiPrintNos As String, mblnOnePatiPrint As Boolean

Private Type tyBillType
    bln���ֽ��㷽ʽ As Boolean
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
    
    lng����ID As Long
    str���� As String
    str�Ա� As String
    str���� As String
    str�ѱ� As String
End Type
Private mCurBillType As tyBillType  '��ǰ��������

Private mobjSquare As Object
Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1

Private mobjDrugPacker  As Object ' �Զ���ҩ��(���·�ҩ����)
Private mblnDrugPacker As Boolean
Private mobjDrugMachine  As Object
Private mblnDrugMachine As Boolean

Private mblnHistoryData As Boolean '�Ƿ�Ϊ���൥�ݷֵ��ݽ��㡱��һ�ν���ֵ����˷ѡ�ʱ����ʷ����
Private mblnDelByNo As Boolean '�Ƿ�ֵ����˷� = (�൥�ݷֵ��ݽ���=True Or һ�ν���ֵ����˷�=True) And Not mblnHistoryData
Private mcllForceDelToCash As Collection 'ǿ��������Ϣ��Array(����Ա,���������)
'-------------------------------------------------------------------------------
'ҽ����ض���:����
Private Type TYPE_MedicarePAR
    ҽ���ӿڴ�ӡƱ�� As Boolean
    �˷Ѻ��ӡ�ص� As Boolean
    ҽ������Ʊ��  As Boolean        'Ԥ����ʱ��Ч
    ����������� As Boolean             'ҽ���Ƿ�֧�������������
    ����Ԥ���� As Boolean
    ���Ը� As Boolean
    ȫ�Ը� As Boolean
    ������ȫ�� As Boolean
    �൥�ݷֵ��ݽ��� As Boolean '86321
    һ�ν���ֵ����˷� As Boolean '91602
End Type
Private MCPAR As TYPE_MedicarePAR
'-------------------------------------------------------------------------------
'Api����
Private Const VK_RETURN = &HD
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Function CheckDepend() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݹ����Լ��
    '����:���ݹ������Ϸ�����true,���򷵻�False
    '����:���˺�
    '����:2014-07-07 11:41:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mbytMode = EM_MULTI_�鿴 Then CheckDepend = True: Exit Function
    
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
    
    CheckDepend = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ShowMe(frmMain As Object, ByVal bytMode As gEM_ChargeDelType, _
    ByVal strPrivs As String, lng������� As Long, _
    Optional blnPrintView As Boolean, _
    Optional lng����ID As Long = 0, _
    Optional blnNOMoved As Boolean = False, _
    Optional strDelTime As String = "", _
    Optional strApplyTime As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݲ鿴,�˷�
    '���:bytMode-0-���ŵ��ݲ鿴,1-���ŵ����˷�,2-���쳣���˷ѵ����������˷�
    '     strPrivs-Ȩ�޴�
    '     mblnPrintView-��ӡǰ�鿴����
    '     blnNOMoved-�Ƿ�ת�������ݱ�
    '     strDelTime-�鿴�˷ѵ��ݵĵǼ�ʱ��(yyyy-MM-dd HH:mm:ss) 'ֻ�в鿴�˷ѵ���ʱ�Ŵ���ʱ��,��������������
    '     strApplyTime-����ʱ��(yyyy-MM-dd HH:mm:ss)���˷�����ģʽʱ���봫��
    '����:
    '����:�˷ѳɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-06-24 14:34:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mblnNOMoved = blnNOMoved: mstrPrivs = strPrivs
    mlng����ID = lng����ID: mlng������� = lng�������
    mlngModule = 1121: mblnPrintView = blnPrintView
    mbytMode = bytMode:
    mstrDelTime = strDelTime              'ֻ�в鿴�˷ѵ���ʱ�Ŵ���ʱ��,��������������
    mstrApplyTime = strApplyTime
    mblnOK = False
    If CheckDepend = False Then Exit Function
    On Error Resume Next
    
    Me.Show IIf(gfrmMain Is Nothing, 0, 1), frmMain
    On Error GoTo 0
    ShowMe = mblnOK
End Function

Private Sub cmdRefuseApply_Click()
    If SaveDelApplied(EM_MULTI_�ܾ�����) = False Then Exit Sub
    mblnOK = True
    Unload Me: Exit Sub
End Sub

Private Sub Form_Load()
    mblnFirst = True
    mstrTittle = "�����˷ѹ���"
    Call InitFace
    Call RestoreWinState(Me, App.ProductName, mstrTittle)
    
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
    Dim TY_Temp As tyBillType
    
    Call SetpicInvoiceVisible '���÷�Ʊ�ؼ���ʾ
    Call InitBillHead   '���õ�����ͷ
    
    If Val(zlDatabase.GetPara("�˷Ѻ�������ģʽ", glngSys, 1121, 0)) = 0 Then
        optNO(0).Value = True
    Else
        optNO(1).Value = True
    End If
    mCurBillType = TY_Temp
    mint�˷ѻص���ӡ = Val(zlDatabase.GetPara("�˷ѻص���ӡ��ʽ", glngSys, mlngModule, "0"))
    Call NewCardObject
    Call ClearFace
    Call SetFunCtrlVisible
    
    Select Case mbytMode
    Case EM_MULTI_�鿴
        mstrTittle = "�����շѵ��ݲ���"
        Caption = mstrTittle
        vsBill.ColHidden(0) = True
        cmdCancel.Caption = "�˳�(&X)"
        If mblnPrintView Then cmdCancel.Caption = "ȷ��(&X)"
        pic��.Visible = mstrDelTime <> ""
        
        mblnOneCard = False
    Case EM_MULTI_�쳣����
        mstrTittle = "�����˷ѹ���-�쳣�˷ѵ������˷�"
        Caption = mstrTittle
        vsBill.ColHidden(0) = True
        pic��.Visible = mstrDelTime <> ""
        vsBill.Editable = flexEDNone
        mblnOneCard = GetOneCard.RecordCount <> 0
        Call initCardSquareData
    Case EM_MULTI_�˷�����, EM_MULTI_ȡ������, EM_MULTI_�˷����, EM_MULTI_ȡ�����
        mstrTittle = "�����˷ѹ���-" & Switch(mbytMode = EM_MULTI_�˷�����, "�˷�����", mbytMode = EM_MULTI_ȡ������, "ȡ������", _
                                            mbytMode = EM_MULTI_�˷����, "�˷����", mbytMode = EM_MULTI_ȡ�����, "ȡ�����")
        Caption = mstrTittle
        Call initCardSquareData
    Case Else 'EM_MULTI_�˷�
        mstrTittle = "�����˷ѹ���"
        Caption = mstrTittle
        Call initCardSquareData
        mblnOneCard = GetOneCard.RecordCount <> 0
    End Select
    
    If mlng������� <> 0 Then
        picPatiBack.Top = txtNO.Top
        lblPati.Top = picPatiBack.Top + (picPatiBack.Height - lblPati.Height) \ 2
        txtPatientPrint.Top = txtNO.Top
        lblPatiName.Top = txtPatientPrint.Top + (txtPatientPrint.Height - lblPatiName.Height) \ 2
        picPati.Height = 480
    End If
    
    
End Sub

Private Sub SetpicInvoiceVisible()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���÷�Ʊ�ؼ�����ʾ
    '����:���˺�
    '����:2014-06-24 14:36:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    picInvoice.Visible = False
    If gTy_Module_Para.bytƱ�ݷ������ = 0 And gTy_Module_Para.bln�ֱ��ӡ = False Then GoTo ReSizing:
    If mbytMode <> EM_MULTI_�˷� Then GoTo ReSizing:
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

Private Sub InitBillHead()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���˷ѵı�ͷ����Ϣ
    '����:���˺�
    '����:2014-06-24 14:39:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim arrHead As Variant, strHead As String, i As Long
    Dim varTemp As Variant, intCol As Integer

    strHead = "" & _
    "ѡ��,300,4;���ݺ�,1000,1;���,720,1;��Ŀ,2800,1;��Ʒ��,2000,1;����,750,7;��λ,550,1;����,1100,7;" & _
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

Private Sub ClearFace(Optional ByVal blnNO As Boolean = True)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������Ϣ
    '���:blnNo=������ݺ�
    '����:���˺�
    '����:2014-06-24 15:19:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mrsBalance = Nothing: Set mrsInsureBalance = Nothing
    Set mrsDelInvoice = Nothing

    With vsBill
        .Rows = .FixedRows '�Էǹ̶��еĵ�һ�б�����ʱ�ָ��ɼ�
        .Rows = .FixedRows + 1
        .Row = .FixedRows: .Col = .ColIndex("��Ŀ")
        .Clear 1
    End With
    mCurBillType.strNos = ""
    mCurBillType.intInsure = 0
    lblPati.Caption = "����:"
    If blnNO Then txtNO.Text = ""
    Call SetpicInvoiceVisible
    
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
    If mbytMode <> EM_MULTI_�鿴 Then Exit Sub
   
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

Private Function LoadViewBills(ByVal lng������� As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݽ����������������(ֻ��Բ鿴���쳣�˷�)
    '���:lng�������-�������
    '����:���ػ��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-06-24 16:17:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllInvoiceNoInfor As Collection
    Dim intSign As Integer, strSQL As String, rsTemp As ADODB.Recordset
    Dim str����ID As String, strNos As String, strAllNOs As String, intInsure As Integer
    Dim strWhere As String, lng����ID As Long, strҽ����� As String, lngԭ������� As Long
    Dim lngҽ����� As Long, rsAdvice As ADODB.Recordset, i As Long
    Dim strTemp As String, j As Long, dbl�ϼ� As Double, lng����ID As Long
    Dim varData As Variant, strInfos As String
    Dim varNos As Variant, blnNotDelAll As Boolean, blnHaveExe As Boolean
    
    If mbytMode = EM_MULTI_�˷� Then
        '�˷���Ҫ������
        LoadViewBills = ReadBills("")
        Exit Function
    End If
    
    Screen.MousePointer = 11
    intSign = IIf(mstrDelTime <> "", -1, 1) '����,�����������
    On Error GoTo errHandle
    
    str����ID = zlGet����ID(lng�������, strNos, intInsure, mblnNOMoved, lng����ID)
    
    mCurBillType.lng����ID = lng����ID
    varData = Split(str����ID & ",,", ",")
     If Val(varData(0)) = lng����ID Then
         mCurBillType.lng����ID = Val(varData(1))
    End If
    
    If InStr(str����ID, ",") > 0 Then
        strWhere = "And A.����ID IN (Select Column_Value From table(f_num2List([1])))"
        lng����ID = Split(str����ID & ",", ",")(0)
    Else
        strWhere = "And A.����ID=[2]"
        lng����ID = Val(str����ID)
    End If

        
    'bytType-0-����NO������;1-���ݽ���ID������,2-���ݽ������������
    strAllNOs = zlGetBalanceNos(1, lng����ID, mblnNOMoved)
    mCurBillType.strAllNOs = strAllNOs
    mCurBillType.intInsure = intInsure
    mCurBillType.strNos = strNos
     
     
    'bytType-��������:0-���ݽ���ID����;1-���ݽ�����Ų���,2-����NOs����ȡ
    If mbytMode = EM_MULTI_�쳣���� Then
        mCurBillType.lngԭ����ID = zlGetFromNOToLastBalanceID(strAllNOs, mblnNOMoved, False, lngԭ�������)
    End If
    Set mrsBalance = zlFromIDGetChargeBalance(1, lng�������, mblnNOMoved)
    Set mrsInsureBalance = zlGetInsureBalanceDetail(1, lng�������, mblnNOMoved)
    
    strSQL = "" & _
    " Select A.����ID,A.����,A.�Ա�,A.����,A.��ʶ��,A.�ѱ�,C.���� as ���ʽ,B.��������,B.���� " & _
    " From ������ü�¼ A,ҽ�Ƹ��ʽ C,��Ա�� D,������Ϣ B" & _
    " Where A.���ʽ=C.����(+)  And  A.����Ա����=D.���� And A.����ID=B.����ID(+) " & _
    "       And (D.վ��='" & gstrNodeNo & "' Or D.վ�� is Null)" & vbNewLine & _
    "       And mod(A.��¼����,10)=1 And Rownum <2 " & strWhere
    If mblnNOMoved Then
        strSQL = Replace(strSQL, "������ü�¼", "H������ü�¼")
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str����ID, lng����ID)
    If rsTemp.EOF Or strAllNOs = "" Then
        Screen.MousePointer = 0
        MsgBox "û���ҵ��������ص��շѼ�¼��", vbInformation, gstrSysName
        mCurBillType.lng����ID = 0
        Exit Function
    End If
    
    txtPatient.Text = Nvl(rsTemp!����)
    lblPati.Caption = "����:" & IIf(txtPatient.Visible, "       ", rsTemp!����) & _
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
    
    If mbytMode <> EM_MULTI_�鿴 Then
        Call initInsurePara(mCurBillType.intInsure, mCurBillType.lng����ID, lng����ID)
        mblnDelByNo = MCPAR.�൥�ݷֵ��ݽ��� Or MCPAR.һ�ν���ֵ����˷�
    End If
    If CheckPrivsIsValied = False Then Screen.MousePointer = 0: Exit Function   '����Ȩ�޼��
    
    If mCurBillType.intInsure <> 0 Then
        lblPati.ForeColor = vbRed
        txtYB.Text = mCurBillType.intInsure
        txtPatient.ForeColor = vbRed
    Else
        lblPati.ForeColor = &HC00000
        txtYB.Text = ""
        txtPatient.ForeColor = &HC00000
    End If
    Call SetPatiColor(txtPatient, Nvl(rsTemp!��������), txtPatient.ForeColor)
    lblPati.ForeColor = txtPatient.ForeColor
    If mblnPrintView And zlStr.IsHavePrivs(mstrPrivs, "�޸������ش�") _
        And mCurBillType.lng����ID = 0 Then
        txtPatientPrint.Text = "" & rsTemp!����
        txtPatientPrint.Tag = txtPatientPrint.Text
        txtPatientPrint.Visible = True
        lblPatiName.Visible = True
    End If
    
    If mblnDelByNo Then
        If mrsBalance.RecordCount > 0 Then
            mblnHistoryData = zlGetInsureBalanceDetail(1, Val(Nvl(mrsBalance!�������))).RecordCount = 0
        End If
        mblnDelByNo = Not mblnHistoryData
    End If
    
    '���ؽ��㷽ʽ
    Call LoadBalanceInfor
    If mbytMode = EM_MULTI_�˷����� Then
        strWhere = strWhere & " And Not exists(select 1 From �����˷����� where NO=A.NO And ��¼����=1 And Nvl(״̬,0) In(0,1) ) "
    ElseIf mbytMode = EM_MULTI_ȡ������ Or mbytMode = EM_MULTI_�˷���� Then
        strWhere = strWhere & " And Exists(select 1 From �����˷����� where NO=A.NO And ��¼����=1 And Nvl(״̬, 0) = 0 " & _
                              " And ����ʱ��=To_Date('" & mstrApplyTime & "','yyyy-mm-dd hh24:mi:ss')) "
    ElseIf mbytMode = EM_MULTI_ȡ����� Then
        strWhere = strWhere & " And Exists(select 1 From �����˷����� where NO=A.NO And ��¼����=1 And Nvl(״̬, 0) = 1 " & _
                              " And ����ʱ��=To_Date('" & mstrApplyTime & "','yyyy-mm-dd hh24:mi:ss')) "
        '���˹��ѵĲ�����ȡ�����
        strWhere = strWhere & " And Not Exists(select 1 From ������ü�¼ where NO=A.NO And ��¼����=A.��¼���� And ��¼״̬=2)"
    End If
    'InStr(str����ID, ",") > 0:��ʾ���ܴ������յ���������Կ϶��ǲ���˷Ѽ�¼������ժҪӦ�����˷ѵ�ժҪΪ׼
    '104788�����������㸶����ֱ�Ӽ������Σ���Ϊҽ�����˶൥��һ�ν���ʱ�����ˣ��鿴�˷ѵ�����ʾ������������
    '" Avg(Nvl(A.����,1)) as ����,Avg(A.����) as ����," ��Ϊ " Avg(Nvl(A.����,1)*A.����) as ����,"
    strSQL = "" & _
    "   Select A.NO,Nvl(A.�۸񸸺�,A.���) as ���,A.��������,A.��������ID,A.ִ�в���ID,A.�շ����,A.�ѱ�,A.�շ�ϸĿID," & vbNewLine & _
    "          A.��������,A.���㵥λ,max(A.ҽ�����) as ҽ�����," & vbNewLine & _
    "          Avg(Nvl(A.����,1)*A.����) as ����," & vbNewLine & _
    "          Sum(A.��׼����) as ����, Sum(A.Ӧ�ս��) as Ӧ�ս��,Sum(A.ʵ�ս��) as ʵ�ս��," & vbNewLine & _
    "          Max(A.����Ա����) as ����Ա����,max(A.�Ǽ�ʱ��) as �Ǽ�ʱ��," & _
    "           " & IIf(InStr(str����ID, ",") > 0, "Max(Decode(A.��¼״̬,2,A.ժҪ,NULL))", "Max(A.ժҪ)") & " as ժҪ,A.����ID" & vbNewLine & _
    "   From ������ü�¼ A" & vbNewLine & _
    "   Where Mod(A.��¼����,10)=1  " & strWhere & vbNewLine & _
    "   Group by A.����ID,A.NO,Nvl(A.�۸񸸺�,A.���),A.��������,A.��������ID,A.ִ�в���ID,A.�շ����,A.�ѱ�,A.�շ�ϸĿID,A.��������,A.���㵥λ,A.����ID"
    
    If mblnNOMoved Then
        strSQL = Replace(strSQL, "������ü�¼", "H������ü�¼")
    End If
    
    strSQL = _
    " Select /*+ Rule*/ A.NO,A.���,A.��������,A.�ѱ�,A.�շ�ϸĿID,C.���� as �����,C.���� as �����,B.����, " & _
    "       Nvl(M1.����,B.����) as ����,E1.���� as ��Ʒ�� ,B.���,Nvl(A.��������,B.��������) ��������," & _
            IIf(gblnҩ����λ, "Decode(X.ҩƷID,NULL,A.���㵥λ,X." & gstrҩ����λ & ")", "A.���㵥λ") & " as ���㵥λ," & _
    "       Max(A.ҽ�����) as ҽ�����," & _
    "       sum(A.����" & IIf(gblnҩ����λ, "/Nvl(X." & gstrҩ����װ & ",1)", "") & ") as ����," & _
    "       Max(A.����" & IIf(gblnҩ����λ, "*Nvl(X." & gstrҩ����װ & ",1)", "") & ") as ����," & _
    "       Sum(A.Ӧ�ս��) as Ӧ�ս��,Sum(A.ʵ�ս��) as ʵ�ս��, 1 as ��¼��־,0 as ԭʼ����,0 as ׼������," & _
    "       D.���� as ִ�п���,A.ִ�в���ID,E.���� as ��������,A.����Ա����,A.�Ǽ�ʱ��, " & _
    "       Max(A.ժҪ) as ժҪ" & _
    " From (" & strSQL & ") A,�շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D,���ű� E,ҩƷ��� X," & _
    "       �շ���Ŀ���� M1,�շ���Ŀ���� E1" & _
    " Where A.�շ�ϸĿID=B.ID And C.����=A.�շ���� And A.�շ�ϸĿID=X.ҩƷID(+)" & _
    "       And A.ִ�в���ID=D.ID(+) And A.��������ID=E.ID(+) " & _
    "       And A.�շ�ϸĿID=M1.�շ�ϸĿID(+) And M1.����(+)=1 And M1.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
    "       And A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3 " & _
    " Group by A.NO,A.���,A.��������,A.�ѱ�,A.�շ�ϸĿID,C.����,C.����,B.����,Nvl(M1.����,B.����)," & _
    "       E1.����,B.���,Nvl(A.��������,B.��������),A.���㵥λ,D.����,A.ִ�в���ID,E.����,X.ҩƷID,X." & gstrҩ����λ & ",A.����Ա����,A.�Ǽ�ʱ��" & _
    " Having Sum(A.����)<>0 " & _
    " Order by NO,���"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str����ID, lng����ID)
    
    If rsTemp.EOF Then
        Screen.MousePointer = 0
        If mbytMode = EM_MULTI_�鿴 Then
            MsgBox "û���ҵ�ָ��������Ϣ�ķ��ü�¼,�����򲢷�ԭ�����˲���������˴���Ľ��㵥�ݡ�", vbInformation, gstrSysName
        ElseIf mbytMode = EM_MULTI_�˷����� Then
            MsgBox "û���ҵ���Ҫ�˷�����ĵ��ݣ������ѱ����������ˡ�", vbInformation, gstrSysName
        ElseIf mbytMode = EM_MULTI_ȡ������ Then
            MsgBox "û���ҵ���Ҫȡ������ĵ��ݣ������ѱ�����ȡ�����������ˡ�", vbInformation, gstrSysName
        ElseIf mbytMode = EM_MULTI_�˷���� Then
            MsgBox "û���ҵ���Ҫ�˷���˵ĵ��ݣ������ѱ���������ˡ�", vbInformation, gstrSysName
        ElseIf mbytMode = EM_MULTI_ȡ����� Then
            MsgBox "û���ҵ���Ҫȡ����˵ĵ��ݣ��������˷ѻ�����ȡ������ˡ�", vbInformation, gstrSysName
        Else
            MsgBox "û���ҵ��������Ϣ��صĿ����˷ѵļ�¼��" & _
                vbCrLf & "��Щ�շѼ�¼�����Ѿ��˷ѻ��Ѿ���ȫִ�С�", vbInformation, gstrSysName
        End If
        Call ClearFace(False)
        Exit Function
    End If
    
    If mbytMode <> EM_MULTI_�˷� Then
        If mbytMode = EM_MULTI_�˷����� Or mbytMode = EM_MULTI_ȡ������ Or mbytMode = EM_MULTI_�˷���� Or mbytMode = EM_MULTI_ȡ����� Then
            pic�˷�ժҪ.Enabled = True
            pic�˷�ժҪ.Visible = True
            lblժҪ.Caption = Switch(mbytMode = EM_MULTI_�˷�����, "����ԭ��", mbytMode = EM_MULTI_ȡ������, "����ԭ��", mbytMode = EM_MULTI_�˷����, "���/�ܾ�ԭ��", _
                                    mbytMode = EM_MULTI_ȡ�����, "���ԭ��")
            txt�˷�ժҪ.Text = ""
            If mbytMode = EM_MULTI_�˷����� Or mbytMode = EM_MULTI_ȡ������ Then
                lbl�˿�ϼ�.Caption = "����ϼ�"
            ElseIf mbytMode = EM_MULTI_�˷���� Or mbytMode = EM_MULTI_ȡ����� Then
                lbl�˿�ϼ�.Caption = "��˺ϼ�"
            End If
        Else
            pic�˷�ժҪ.Enabled = mbytMode = EM_MULTI_�쳣����
            txt�˷�ժҪ.Text = Nvl(rsTemp!ժҪ)
        End If
    End If
    
    With rsTemp
        strҽ����� = ""
        Do While Not .EOF
            lngҽ����� = Val(Nvl(!ҽ�����))
            If InStr(strҽ����� & ",", "," & lngҽ����� & ",") = 0 And lngҽ����� <> 0 Then
                strҽ����� = strҽ����� & "," & Val(Nvl(!ҽ�����))
            End If
            .MoveNext
        Loop
        .MoveFirst
    End With
    
    Set rsAdvice = Nothing
    If strҽ����� <> "" Then
        strҽ����� = Mid(strҽ�����, 2)
        Set rsAdvice = zlGetAdviceFromID(strҽ�����)
    End If
    Call LoadInvoiceData(Replace(strAllNOs, "'", ""))
    With vsBill
        .Redraw = flexRDNone
        .Rows = .FixedRows + rsTemp.RecordCount
        mCurBillType.strDelNOs = ""
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
            .TextMatrix(i, .ColIndex("��Ŀ")) = strTemp & rsTemp!���� & IIf(IsNull(rsTemp!���), "", " " & rsTemp!���)
            .TextMatrix(i, .ColIndex("��Ʒ��")) = strTemp & Nvl(rsTemp!��Ʒ��)
            .TextMatrix(i, .ColIndex("����")) = FormatEx(intSign * rsTemp!����, 5)
            .Cell(flexcpData, i, .ColIndex("����")) = intSign * rsTemp!����
            .TextMatrix(i, .ColIndex("��λ")) = Nvl(rsTemp!���㵥λ)
            .TextMatrix(i, .ColIndex("����")) = Format(rsTemp!����, gstrFeePrecisionFmt)
            .TextMatrix(i, .ColIndex("Ӧ�ս��")) = Format(intSign * Val(Nvl(rsTemp!Ӧ�ս��)), gstrDec)
            .TextMatrix(i, .ColIndex("ʵ�ս��")) = Format(intSign * Val(Nvl(rsTemp!ʵ�ս��)), gstrDec)
            .TextMatrix(i, .ColIndex("��������")) = Nvl(rsTemp!��������)
            .TextMatrix(i, .ColIndex("ִ�п���")) = Nvl(rsTemp!ִ�п���)
            .TextMatrix(i, .ColIndex("����Ա")) = rsTemp!����Ա����
            .TextMatrix(i, .ColIndex("ʱ��")) = Format(rsTemp!�Ǽ�ʱ��, "MM-dd HH:mm")
            .TextMatrix(i, .ColIndex("����ID")) = str����ID
            lng������� = Val(Nvl(rsTemp!ҽ�����))
            If Not rsAdvice Is Nothing And strҽ����� <> "" And lng������� <> 0 Then
                rsAdvice.Filter = "ҽ��ID=" & lng�������
                If rsAdvice.EOF = False Then
                    .TextMatrix(i, .ColIndex("ҽ��")) = Nvl(rsAdvice!ҽ������)
                End If
            End If
            .TextMatrix(i, .ColIndex("ԭʼ����")) = Nvl(rsTemp!ԭʼ����)
            .TextMatrix(i, .ColIndex("׼������")) = Nvl(rsTemp!׼������)
            .TextMatrix(i, .ColIndex("ҽ�����")) = Nvl(rsTemp!ҽ�����)
            .TextMatrix(i, .ColIndex("ִ�п���ID")) = Nvl(rsTemp!ִ�в���ID)
            .Cell(flexcpData, i, .ColIndex("ѡ��")) = Val(Nvl(rsTemp!��¼��־))    '�����ж��Ƿ����ʹ�,>1��ʾ������
            If Val(Nvl(rsTemp!��¼��־)) > 1 And InStr(1, mCurBillType.strNosPatiDel & ",", "," & rsTemp!NO & "") = 0 Then
                mCurBillType.strNosPatiDel = mCurBillType.strNosPatiDel & "," & rsTemp!NO
            End If
            If InStr(mCurBillType.strDelNOs & ",", "," & rsTemp!NO & ",") = 0 Then
                '�����ָ���
                If mCurBillType.strDelNOs <> "" Then
                    .Select i, .FixedCols, i, .COLS - 1
                    .CellBorder vbBlue, 0, 1, 0, 0, 0, 0
                End If
                mCurBillType.strDelNOs = mCurBillType.strDelNOs & "," & rsTemp!NO
            End If
            dbl�ϼ� = dbl�ϼ� + Val(Nvl(rsTemp!ʵ�ս��))
            rsTemp.MoveNext
        Next
        .Row = .FixedRows: .Col = .ColIndex("��Ŀ")
        Call vsBill_AfterRowColChange(-1, -1, .Row, .Col)
        .Redraw = flexRDDirect
    End With
    
    Call SetpicInvoiceVisible   '���÷�Ʊ�ؼ�����ʾ
    txtAllTotal.Text = Format(intSign * dbl�ϼ�, gstrDec)
    Call ReInitPatiInvoice
    txt�˿�ϼ�.Text = Format(GetDelMoney, "0.00")
    
    If mbytMode = EM_MULTI_�˷����� Or mbytMode = EM_MULTI_�˷���� Then '������ִ�е���Ŀʱ������ʾ
        varNos = Split(mCurBillType.strDelNOs, ",")
        For i = 0 To UBound(varNos)
            Call BillCanDelete(varNos(i), 1, blnHaveExe)
            If blnHaveExe Then
                strInfos = strInfos & "," & varNos(i)
            End If
        Next
        If strInfos <> "" Then
            strInfos = Mid(strInfos, 2)
            If MsgBox("����[" & strInfos & "]�д�����ִ�е���Ŀ����ȷ��Ҫ������", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                Call ClearFace(False): Screen.MousePointer = 0: Exit Function
            End If
        End If
    End If
    
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

Private Function ReadBills(ByVal strNo As String, Optional blnCheckMulitBalance As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݵ�ǰ����ĵ��ݺŻ�Ʊ�ݺ�,��ȡ����ʾ���ŵ���
    '���:strNO-ָ���ĵ��ݺŻ�Ʊ��
    '     blnCheckMulitBalance-�Ѿ�����˶൥��һ�ν���,���������ڲ����
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-06-24 15:41:24
    '˵��:
    '   ֻ���˷�ģʽ�²Ž����ģ��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim strNos As String, strAllNOs As String
    Dim blnNOMoved As Boolean
    Dim strTmp As String, strCanDelNos As String
    Dim i As Long, j As Integer
    Dim dbl�ϼ� As Currency, arrNo As Variant
    Dim strTemp As String, strҽ����� As String
    Dim blnNotFind As Boolean
    Dim lng����ID As Long, cllInvoiceNoInfor As Collection
    Dim str������� As String, blnFind As Boolean
    Dim strInvoiceNO As String, strOldNO As String
    On Error GoTo errH
    If mbytMode <> EM_MULTI_�˷� Then Exit Function
    
    Screen.MousePointer = 11
    
    Call ClearFace(False)
    strOldNO = strNo
    Set cllInvoiceNoInfor = New Collection
    mblnOnePatiPrint = False: mstrOnePatiPrintNos = ""
    If mlng������� = 0 And strNo <> "" Then
        strInvoiceNO = ""
        If Not (mstrNo <> "" Or optNO(0).Value) Then
             '��Ʊ�ݺ�:���ܲ�ͬ����Ʊ���ظ�
            strInvoiceNO = strNo
            blnNOMoved = zlDatabase.NOMoved("Ʊ�ݴ�ӡ��ϸ", "Ʊ�� =", "1")
            strNos = zlInvoiceGetNOs(strInvoiceNO, cllInvoiceNoInfor, blnNOMoved)
            strNo = Split(strNos & ",", ",")(0)
            If zlIsOnePatiPrint(strNo, mstrOnePatiPrintNos, mblnOnePatiPrint, blnNOMoved) = False Then Exit Function
            
            If mblnOnePatiPrint Then    '��ν���ʱ����Ҫѡ��ָ���Ľ��㷽ʽ
                 If SelectMulitBalance(mstrOnePatiPrintNos, strNo) = False Then Exit Function
            End If
        Else
            If zlIsOnePatiPrint(strNo, mstrOnePatiPrintNos, mblnOnePatiPrint, blnNOMoved) = False Then Exit Function

        End If
        
        blnNOMoved = zlDatabase.NOMoved("������ü�¼", strNo, , "1")
        strNos = zlGetBalanceNos(0, strNo, blnNOMoved)
        If blnCheckMulitBalance = False Then
            '78663,Ƚ����,2014-10-15,���շѵ���ʱ������ʾ
            If Trim(strNos) = "" Then
                Screen.MousePointer = 0
                MsgBox "û���ҵ������""" & strOldNO & """��ص��շѼ�¼��", vbInformation, gstrSysName
                Exit Function
            End If
            If Not zlIsMulitOneBalance(strNos) Then
                '�Ƕ൥��һ�ν���,�����34��ǰ�汾���������
                'bytMode-0-���ŵ��ݲ鿴,1-���ŵ����˷�,2-���쳣���˷ѵ����������˷�
                frmMultiBills.ShowMe Me, 1, mstrPrivs, strNo, "", False, mlng����ID, mblnOneCard, False, True
                Exit Function
            End If
        End If
    Else
        'bytType-0-����NO������;1-���ݽ���ID������,2-���ݽ������������
        strNos = zlGetBalanceNos(2, mlng�������, mblnNOMoved)
        strNo = Split(strNos & ",", ",")(0)
        If zlIsOnePatiPrint(strNo, mstrOnePatiPrintNos, mblnOnePatiPrint, blnNOMoved) = False Then Exit Function '������������ģ��϶�������ȥѡ�񵥾�
    End If
    
    strAllNOs = strNos
    
    If strNos = "" Then
        If optNO(1).Value Then
            Screen.MousePointer = 0
            MsgBox "û���ҵ������""" & strNo & """��ص��շѼ�¼��", vbInformation, gstrSysName
            Exit Function
        End If
        '������Ϊδ��Ʊ�ݶ���������
        strNos = strNo
    End If
    mCurBillType.lngԭ����ID = zlGetFromNOToLastBalanceID(strNos, blnNOMoved)
    mCurBillType.strAllNOs = strAllNOs
    'ִ�в�����ĵ��ݲ��ܽ����˷�
    If CheckBillExistReplenishData(1, , strNos) = True Then
        Screen.MousePointer = 0
        MsgBox "ѡ����˷Ѽ�¼������ҽ��������㣬����������˷Ѳ�����", vbInformation, gstrSysName
        Exit Function
    End If
    
    '����ҽ��ִ�мƼ�.ִ��״̬
    If Upgradeҽ��ִ�мƼ�ִ��״̬(strNos) = False Then
        Screen.MousePointer = 0
        MsgBox "ҽ��ִ�мƼ���������ʧ�ܣ����ܼ�����", vbInformation, gstrSysName
        Exit Function
    End If
    
    '��Ҫ������
    If InStr(1, strNos, "'") = 0 Then
        strNos = "'" & Replace(strNos, ",", "','") & "'"
    End If
    arrNo = Split(strNos, ",")
    
    '��ȡ���㷽ʽ
    '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�)
    Set mrsBalance = zlFromIDGetChargeBalance(2, strAllNOs, mblnNOMoved)
    Set mrsInsureBalance = zlGetInsureBalanceDetail(2, strAllNOs, mblnNOMoved)
    mCurBillType.intInsure = zlGetChargeInsure(mCurBillType.lngԭ����ID, lng����ID, mblnNOMoved)
    
    '��ʼ�����㷽ʽ��ر���
    Call InitBalanceVar
    
    
    Call initInsurePara(mCurBillType.intInsure, lng����ID, mCurBillType.lngԭ����ID)
    mblnDelByNo = MCPAR.�൥�ݷֵ��ݽ��� Or MCPAR.һ�ν���ֵ����˷�
    
    If CheckPrivsIsValied = False Then Exit Function    '����Ȩ�޼��
    
    '�˷���ؼ��
    If CheckDelIsValied(strNos, mCurBillType.strNotCanDelNOs, strCanDelNos) = False Then
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
    "       And mod(A.��¼����,10)=1 And A.��¼״̬ IN(1,3) And A.NO=[1] And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo)
    If rsTemp.EOF Then
        Screen.MousePointer = 0
        MsgBox "û���ҵ������""" & strNo & """��ص��շѼ�¼��", vbInformation, gstrSysName
        mCurBillType.lng����ID = 0
        Exit Function
    End If
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
        txtYB.Text = Val(Nvl(rsTemp!����))   '����:41760
        txtPatient.ForeColor = vbRed
    Else
        lblPati.ForeColor = &HC00000
        txtYB.Text = ""
        txtPatient.ForeColor = &HC00000
    End If
    Call SetPatiColor(txtPatient, Nvl(rsTemp!��������), txtPatient.ForeColor)
    lblPati.ForeColor = txtPatient.ForeColor
    If mblnPrintView And zlStr.IsHavePrivs(mstrPrivs, "�޸������ش�") _
        And mCurBillType.lng����ID = 0 Then
        txtPatientPrint.Text = "" & rsTemp!����
        txtPatientPrint.Tag = txtPatientPrint.Text
        txtPatientPrint.Visible = True
        lblPatiName.Visible = True
    End If

    If mblnDelByNo Then
        If mrsBalance.RecordCount > 0 Then
            mblnHistoryData = zlGetInsureBalanceDetail(1, Val(Nvl(mrsBalance!�������))).RecordCount = 0
        End If
        mblnDelByNo = Not mblnHistoryData
    End If

    '��ȡ��������:ԭʼ���˷ѵ�,���㷽ʽΪ��ָ��Ԥ���ļ�¼
    '----------------------------------------------------------------------------------
    Call LoadBalanceInfor
    Call LoadInvoiceData(strNos)

    If GetFeeListData(strNos, rsTemp) = False Then
        Call ClearFace(False)
        Exit Function
    End If

    strҽ����� = ""
    If rsTemp.EOF Then
        Screen.MousePointer = 0
        MsgBox "û���ҵ������""" & strNo & """��صĿ����˷ѵļ�¼��" & _
            vbCrLf & "��Щ�շѼ�¼�����Ѿ��˷ѻ��Ѿ���ȫִ�С�", vbInformation, gstrSysName
        Call ClearFace(False)
        Exit Function
    End If

    mCurBillType.strNosOverFlow = ""
    strTmp = ""
    For i = 0 To UBound(Split(strNos, ","))
        strTmp = Replace(Split(strNos, ",")(i), "'", "")
        '����Ƿ��������
        If Not BillOperCheck(2, rsTemp!����Ա����, rsTemp!�Ǽ�ʱ��, "�˷�", strTmp, , 1, True) Then
            mCurBillType.strNosOverFlow = mCurBillType.strNosOverFlow & " ," & strTmp
        End If
    Next
    If mCurBillType.strNosOverFlow <> "" Then mCurBillType.strNosOverFlow = Mid(mCurBillType.strNosOverFlow, 2)
    
    With vsBill
        .Redraw = flexRDNone
        .Rows = .FixedRows + rsTemp.RecordCount
        mCurBillType.strNos = ""
        For i = 1 To rsTemp.RecordCount
            '����:29201
            .Cell(flexcpData, i, .ColIndex("��Ŀ")) = Nvl(rsTemp!��������)
            '����:33634
            .Cell(flexcpData, i, .ColIndex("����ID")) = Nvl(rsTemp!ҽ�����) & "," & Nvl(rsTemp!�շ�ϸĿID)
            If Val(Nvl(rsTemp!ҽ�����)) <> 0 And InStr(strҽ����� & ",", "," & Nvl(rsTemp!ҽ�����) & ",") = 0 Then
                strҽ����� = strҽ����� & "," & Nvl(rsTemp!ҽ�����)
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
            For j = 1 To cllInvoiceNoInfor.Count
                If cllInvoiceNoInfor(j)(0) = Nvl(rsTemp!NO) Then
                    If InStr(1, "," & cllInvoiceNoInfor(j)(1) & ",", "," & Nvl(rsTemp!���) & ",") > 0 Then
                         .TextMatrix(i, .ColIndex("ѡ��")) = 1: Exit For
                    End If
                End If
            Next
            .TextMatrix(i, .ColIndex("���ݺ�")) = rsTemp!NO
            .TextMatrix(i, .ColIndex("���")) = rsTemp!�����
            .Cell(flexcpData, i, .ColIndex("���")) = Nvl(rsTemp!�����)
            .TextMatrix(i, .ColIndex("��Ŀ")) = strTemp & rsTemp!���� & IIf(IsNull(rsTemp!���), "", " " & rsTemp!���)
            .TextMatrix(i, .ColIndex("��Ʒ��")) = strTemp & Nvl(rsTemp!��Ʒ��)
            .TextMatrix(i, .ColIndex("����")) = FormatEx(Nvl(rsTemp!����, 1) * rsTemp!����, 5)
            .Cell(flexcpData, i, .ColIndex("����")) = Nvl(rsTemp!����, 1) * Val(Nvl(rsTemp!����))
            
            .TextMatrix(i, .ColIndex("��λ")) = Nvl(rsTemp!���㵥λ)
            .TextMatrix(i, .ColIndex("����")) = Format(Val(Nvl(rsTemp!����)), gstrFeePrecisionFmt)
            .TextMatrix(i, .ColIndex("Ӧ�ս��")) = Format(Val(Nvl(rsTemp!Ӧ�ս��)), gstrDec)
            .TextMatrix(i, .ColIndex("ʵ�ս��")) = Format(Val(Nvl(rsTemp!ʵ�ս��)), gstrDec)
            .TextMatrix(i, .ColIndex("��������")) = Nvl(rsTemp!��������)
            .TextMatrix(i, .ColIndex("ִ�п���")) = Nvl(rsTemp!ִ�п���)
            .TextMatrix(i, .ColIndex("����Ա")) = rsTemp!����Ա����
            .TextMatrix(i, .ColIndex("ʱ��")) = Format(rsTemp!�Ǽ�ʱ��, "MM-dd HH:mm")
            .TextMatrix(i, .ColIndex("����ID")) = rsTemp!����ID
            .TextMatrix(i, .ColIndex("ҽ��")) = Nvl(rsTemp!ҽ������)
            .TextMatrix(i, .ColIndex("ԭʼ����")) = Val(Nvl(rsTemp!ԭʼ����))
            .TextMatrix(i, .ColIndex("׼������")) = Val(Nvl(rsTemp!׼������))
            .TextMatrix(i, .ColIndex("ҽ�����")) = Nvl(rsTemp!ҽ�����)
            .TextMatrix(i, .ColIndex("ִ�п���ID")) = Nvl(rsTemp!ִ�в���ID)
            
            If Not mCurBillType.bln���Ų����˷� Then mCurBillType.bln���Ų����˷� = RoundEx(Val(Nvl(rsTemp!ԭʼ����)), 7) <> RoundEx(Val(Nvl(rsTemp!׼������)), 7)
            If Not mCurBillType.bln���Ų����˷� Then mCurBillType.bln���Ų����˷� = Val(Nvl(rsTemp!��¼��־)) > 1
            
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
    
    If Not mCurBillType.bln���Ų����˷� Then
        mCurBillType.bln���Ų����˷� = zlExistDelFeeChargeBill(mCurBillType.strAllNOs)
    End If
    
    Call SetpicInvoiceVisible   '���÷�Ʊ�ؼ�����ʾ
    
    If mCurBillType.strNos <> "" Then mCurBillType.strNos = Mid(mCurBillType.strNos, 2)
    
    
    txtAllTotal.Text = Format(dbl�ϼ�, gstrDec)
    If strInvoiceNO <> "" Then
        vsBill.Cell(flexcpChecked, 1, vsBill.ColIndex("ѡ��"), vsBill.Rows - 1, vsBill.ColIndex("ѡ��")) = 0
        Call FromInvoiceSelectNO(strInvoiceNO)
        Call SelectRelatingInvoice(strInvoiceNO, True)
        '����ʾ����ѡ�ķ�Ʊ
        Call ShowAndHideDelBillRow
    End If
    '78596,Ƚ����,2014-10-15,Ĭ�Ϲ�ѡ����
    If mlng������� = 0 Then
        '87489
        If gTy_Module_Para.byt�˷�ȱʡѡ��ʽ = 0 Then
            blnFind = False
            With vsBill
                For i = .FixedRows To .Rows - 1
                    If .TextMatrix(i, .ColIndex("���ݺ�")) = strOldNO Then
                        .Row = i: blnFind = True
                        Exit For
                    End If
                Next
            End With
            If blnFind Then Call cmdBillSel_Click
        End If
    End If
'    '78596,Ƚ����,2014-10-14,Ĭ�Ϲ�ѡ����
'    If InStr(";" & mstrPrivs & ";", ";�����˷�;") = 0 Then Call cmdSelAll_Click
    If gTy_Module_Para.byt�˷�ȱʡѡ��ʽ = 1 Then
        Call cmdSelAll_Click
    End If
    
    Call LoadBalanceInfor
    Call ReCalcDelMoney
    Call FromNoSelectInvoice
    
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
    If Not (mbytMode = EM_MULTI_�˷� Or mbytMode = EM_MULTI_�쳣����) Then CheckPrivsIsValied = True: Exit Function

    '���Ȩ���Ƿ�����
    If mCurBillType.intInsure > 0 Then
        '�����˷�Ȩ�޼��
        If zlStr.IsHavePrivs(mstrPrivs, "�����շ�") = False Then
            Screen.MousePointer = 0
            MsgBox "��û��Ȩ�޶�ҽ�����˵ĵ����˷ѣ�", vbInformation, gstrSysName
            Exit Function
        End If
        CheckPrivsIsValied = True: Exit Function
    End If
    
    '��ͨ���˵Ĵ���
    '�Ƿ��з�ҽ�����˵��˷�Ȩ��
    If zlStr.IsHavePrivs(mstrPrivs, "�����ҽ������") = False Then
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
    Call FromNoSelectInvoice
    Call ShowAndHideDelBillRow
    Call ReCalcDelMoney
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
Private Sub ExecuteModifyPatiName()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ�в�����Ϣ�޸�
    '����:���˺�
    '����:2014-07-03 17:00:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPro As New Collection
    Dim strSQL As String, arrNo As Variant, i As Long
    
    arrNo = Split(mCurBillType.strNos, ",")
    For i = 0 To UBound(arrNo)
        strSQL = "Zl_���˷��ü�¼_Update('" & arrNo(i) & "',1,null,null,'" & txtPatientPrint.Text & "')"
        zlAddArray cllPro, strSQL
    Next

    On Error GoTo errHandle:
    zlExecuteProcedureArrAy cllPro, Me.Caption
    Exit Sub
errHandle:
    gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
End Sub

Private Sub cmdClear_Click()
    Dim i As Long, j As Long
    
    With vsBill
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, .ColIndex("ѡ��")) = 0
        Next
    End With
    
    With vsInvoice
        If .Visible Then
            If .Rows - 1 >= 0 And .COLS - 1 >= 1 Then
                For i = .FixedRows To .Rows - 1
                    For j = .FixedCols To .COLS - 1
                        If Trim(.TextMatrix(i, j)) <> "" Then .Cell(flexcpChecked, i, j) = 2
                    Next
                Next
            End If
        End If
    End With
    
    Call ShowAndHideDelBillRow
    Call ReCalcDelMoney
End Sub
 


Private Sub cmdOK_Click()
    If mbytMode = EM_MULTI_�鿴 Then Unload Me: Exit Sub
    
    If mbytMode = EM_MULTI_�쳣���� Then
        '�쳣���������˷�
        If ExecuteReDelFee = False Then
            '���¼����쳣����,�Ա��ȡ��ȷ�Ľ�������
            Call LoadViewBills(mlng�������)
            Exit Sub
        End If
        mblnOK = True
        Unload Me: Exit Sub
    End If
    If mbytMode = EM_MULTI_�˷����� Or mbytMode = EM_MULTI_ȡ������ _
        Or mbytMode = EM_MULTI_�˷���� Or mbytMode = EM_MULTI_ȡ����� Then
        If SaveDelApplied(mbytMode) = False Then Exit Sub
        mblnOK = True
        Unload Me: Exit Sub
    End If
    Call ExecDelete
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
    Call ReCalcDelMoney
    Call FromNoSelectInvoice
    Call ShowAndHideDelBillRow
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
    If mlng������� <> 0 Then  'ָ���˽������ݵ�
        If LoadViewBills(mlng�������) = False Then Unload Me: Exit Sub
    End If
    
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
 

Private Sub Form_Resize()
    Dim staH As Long

    On Error Resume Next

    staH = IIf(stbThis.Visible, stbThis.Height, 0)
    
    vsBill.Height = Me.ScaleHeight - picCmd.Height - staH - picPati.Height - picMoney.Height - pic�˷�ժҪ.Height - vsBalance.Height - IIf(picInvoice.Visible, picInvoice.Height, 0)
    
    If Me.ScaleWidth - cmdSelAll.Left - cmdCancel.Width > 5500 Then
        cmdCancel.Left = Me.ScaleWidth - cmdSelAll.Left - cmdCancel.Width
    Else
        cmdCancel.Left = 5500
    End If
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 90
    If mbytMode = EM_MULTI_�˷���� Then
        cmdRefuseApply.Left = cmdOK.Left - cmdRefuseApply.Width - 90
    End If


    fraInfo_1.Width = Me.ScaleWidth + 300
    LineCmd_1.x2 = Me.ScaleWidth + 300

    With txt�˿�ϼ�
        .Left = Me.ScaleWidth - .Width - 100
        lbl�˿�ϼ�.Left = .Left - lbl�˿�ϼ�.Width - 20
    End With
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)

    mbytMode = EM_MULTI_�鿴
    mstrNo = "": mstrDelTime = "": mCurBillType.strNosOverFlow = ""
    mstrApplyTime = ""
    mblnHistoryData = False: mblnDelByNo = False
    mblnNOMoved = False   '�鿴ʱ,���ܴ���true
    Call initCardSquareData
    Call CloseIDCard
    zlDatabase.SetPara "�˷Ѻ�������ģʽ", IIf(optNO(0).Value, "0", "1"), glngSys, 1121, InStr(1, mstrPrivs, ";��������;") > 0
    Call SaveWinState(Me, App.ProductName, mstrTittle)
    
    If Not mrs���㷽ʽ Is Nothing Then Set mrs���㷽ʽ = Nothing
    If Not mrs�շѶ��� Is Nothing Then Set mrs�շѶ��� = Nothing
    If Not mrsDelInvoice Is Nothing Then Set mrsDelInvoice = Nothing
    If Not mrsBalance Is Nothing Then Set mrsBalance = Nothing
    If Not mrsInsureBalance Is Nothing Then Set mrsInsureBalance = Nothing
    If Not mrsInfo Is Nothing Then Set mrsInfo = Nothing
    If Not mobjDrugPacker Is Nothing Then Set mobjDrugPacker = Nothing
    If Not mobjDrugMachine Is Nothing Then Set mobjDrugMachine = Nothing
    If Not mcllForceDelToCash Is Nothing Then Set mcllForceDelToCash = Nothing
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
        If mbytMode = EM_MULTI_�˷���� Then txt�˷�ժҪ.Left = 1600
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
        mstrNo = "" '78663,Ƚ����,2014-10-15,���벡��ID��ʽ���ҵ��ݳɹ�����ͨ�����롰Ʊ�ݺš����Ҳ�������
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
        If Val(.ColData(Col)) = 0 Then Cancel = True: Exit Sub
        .ColComboList(Col) = " ||" & Val(.Cell(flexcpData, Row, Col))
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

    With vsBill
        If Col <> .ColIndex("ѡ��") Then Exit Sub
        If mbytMode = EM_MULTI_�˷����� Or mbytMode = EM_MULTI_ȡ������ Or mbytMode = EM_MULTI_�˷���� Or mbytMode = EM_MULTI_ȡ����� Then
            Call FromNOSelect(vsBill.TextMatrix(Row, .ColIndex("���ݺ�")), Val(vsBill.TextMatrix(Row, .ColIndex("ѡ��"))) <> 0)
            Call ReCalcDelMoney
            Exit Sub
        End If
        If mCurBillType.intInsure <> 0 And (MCPAR.������ȫ�� Or mblnDelByNo) Then '86176
            Call FromNOSelect(vsBill.TextMatrix(Row, .ColIndex("���ݺ�")), Val(vsBill.TextMatrix(Row, .ColIndex("ѡ��"))) <> 0)
            Call ReCalcDelMoney
            '���ݵ���ѡ��Ʊ
            Call FromNoSelectInvoice
            Exit Sub
        End If
        
        stbThis.Panels(2).Text = ""
        If Val(.Cell(flexcpData, Row, .ColIndex("��Ŀ"))) = 0 Then
            For i = Row + 1 To .Rows - 1
                 If Val(.RowData(Row)) <> Val(.Cell(flexcpData, i, .ColIndex("��Ŀ"))) Then Exit For
                .TextMatrix(i, .ColIndex("ѡ��")) = vsBill.TextMatrix(Row, .ColIndex("ѡ��"))
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
        Call ReCalcDelMoney
        '���ݵ���ѡ��Ʊ
        Call FromNoSelectInvoice
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

Private Sub vsBill_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsBill
       If .Col = .ColIndex("ѡ��") Then
            If .ColIndex("���ݺ�") < 0 Then Cancel = True: Exit Sub
            If Trim(.TextMatrix(Row, .ColIndex("���ݺ�"))) = "" Then Cancel = True
       Else
            Cancel = True
       End If
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
            Call ReCalcDelMoney
            Call FromNoSelectInvoice
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

Private Function CheckDelIsValied(ByVal strNos As String, _
    ByRef strNotCanDelNOs As String, _
    ByRef strCanDelNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����˷ѵ����Ƿ�Ϸ�
    '����:strNotCanDelNOs-�����˵ĵ���(�Ѿ�ִ�м������˵ĵ���)
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
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle

    '����:54728
    If Not mbytMode = EM_MULTI_�˷� Then CheckDelIsValied = True: Exit Function   '�˷�ʱ�ж�

    arrNo = Split(strNos, ","): strNotCanDelNOs = ""
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
            Select Case intTmp
                Case 1 '�õ��ݲ�����
                    strInfo = strInfo & "ָ���ĵ��ݲ����ڣ�" & vbCrLf
                    Exit For
                Case 2 '�Ѿ�ȫ����ȫִ��(�շѲ������˷��Զ���ҩ)
                    strInfo = strInfo & "[" & strCurNO & "]�е���Ŀ�Ѿ�ȫ����ȫִ�У������˷ѣ�" & vbCrLf
                Case 3 'δ��ȫִ�в���ʣ������Ϊ0
                    strInfo = strInfo & "[" & strCurNO & "]��δ��ȫִ�е���Ŀʣ������Ϊ�㣬û�п��˷��ã�" & vbCrLf
            End Select
        ElseIf blnHaveExe Then
            '������ִ����Ŀ
            If mCurBillType.intInsure > 0 And (MCPAR.������ȫ�� Or mblnDelByNo) Then '�շ�ҽ���˷�
                strInfo = strInfo & "[" & strCurNO & "]����ҽ�����˵��շѵ��������Ѿ�ִ�е���Ŀ�������˷ѣ�" & vbCrLf
                strNotCanDelNOs = strNotCanDelNOs & "," & strCurNO  '�����˵ĵ���
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
    CheckDelIsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
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
    mCurBillType.bln���ֽ��㷽ʽ = InStr(str���㷽ʽ, ",") = 0
    mCurBillType.str���㷽ʽ = str���㷽ʽ
    
    '4-һ��ͨ(��)
    mrsBalance.Filter = "����=4"
    mCurBillType.blnExistOnCard = mrsBalance.EOF = False
    
    '3.һ��ͨ
    mrsBalance.Filter = "����=3 And  �Ƿ�ȫ��=1 and �Ƿ�����=0"
    mCurBillType.blnExistThreeAllDel = mrsBalance.EOF = False
    mrsBalance.Filter = 0
End Sub

Private Function ExecuteClinicDelNo(ByVal lng����ID As Long, ByVal intInsure As Integer, _
    ByVal lng����ID As Long, ByVal lngԭ����ID As Long, ByRef strAdvance As String, _
    Optional ByVal blnReDelete As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ִ��ҽ���˷ѽ���
    '���:lng����ID-����ID
    '     intInsure-����
    '     lng����ID-����ID
    '     lngԭ����ID-ԭʼ����ID
    '     strAdvance - ���㷽ʽ
    '     blnReDelete - �Ƿ������˷�
    '����:
    '����:ҽ���˷ѽ��׳ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-31 23:38:11
    '˵��:
    '   ���ýӿ�ǰ,�����ȴ�����,��ɺ�,���Զ��ύ����;ʧ��ʱ,���������
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strAllBalance As String, strSQL As String
    Dim varData As Variant, varTemp As Variant, i As Long, p As Integer, strAdvanceOld As String
    Dim colBalance As Collection '��¼���ŵ��ݱ��ս���
    Dim strNo As String, strDel���㷽ʽ As String
    Dim rsCharge As ADODB.Recordset, str���㷽ʽ As String
    On Error GoTo errHandle
    
    strAdvance = lng����ID & "|" & "0"
    Set colBalance = New Collection
    strAdvanceOld = strAdvance
    
    '93337,�˷�ʱ�����ݺŵ�����нӿڵ���
    strSQL = "Select Distinct NO From ������ü�¼ Where ����id = [1] Order By No Desc"
    Set rsCharge = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���γ������õ��ݺ�", lng����ID)
    
    p = 1
    Do While Not rsCharge.EOF
        colBalance.Add Array()
        strDel���㷽ʽ = ""
        strNo = Nvl(rsCharge!NO)
        '���õ����Ƿ���ҽ���˷�
        '������óɹ����ӿڣ���û���κ�ҽ�����ϣ������һ�ε���ҽ���ӿڣ���Ϊ�޷�ȷ���Ƿ���óɹ���
        If blnReDelete Then
            strDel���㷽ʽ = zlGetYBBalanceNo(lng����ID, strNo)
            Call SetBalanceVal(colBalance, p, strDel���㷽ʽ)
        End If
        
        str���㷽ʽ = zlGetYBBalanceNo(lngԭ����ID, strNo, lng����ID, intInsure, True)
        'str���㷽ʽ Ϊ�գ���ʾҽ����֧��ҽ������
        If str���㷽ʽ <> "" And strDel���㷽ʽ = "" Then
            '    Zl_ҽ��������ϸ_Insert(
            strSQL = "Zl_ҽ��������ϸ_Insert("
            '      ����id_In   ҽ��������ϸ.����id%Type,
            strSQL = strSQL & "" & lng����ID & ","
            '      No_In       ҽ��������ϸ.No%Type,
            strSQL = strSQL & "'" & strNo & "',"
            '      ���㷽ʽ_In Varchar2,
            strSQL = strSQL & "'" & str���㷽ʽ & "')"
            '      ��ע_In     ҽ��������ϸ.��ע%Type := Null
            zlDatabase.ExecuteProcedure strSQL, Me.Caption
            
            strAdvance = strAdvanceOld & "|" & strNo
            '��Ϊ�����̶�Ϊҽ������,�������ƹ̶�Ϊҽ������(����ͳ�ﲻ��ȷ��),�Ժ�Ӧȥ���ò���
            If Not gclsInsure.ClinicDelSwap(lngԭ����ID, True, intInsure, _
                                            strAdvance) Then Exit Function
            If strAdvance = strAdvanceOld & "|" & strNo Then strAdvance = ""
            
            If zlInsureCheck(str���㷽ʽ, strAdvance) Then
                str���㷽ʽ = strAdvance
                '    Zl_ҽ��������ϸ_Insert(
                strSQL = "Zl_ҽ��������ϸ_Insert("
                '      ����id_In   ҽ��������ϸ.����id%Type,
                strSQL = strSQL & "" & lng����ID & ","
                '      No_In       ҽ��������ϸ.No%Type,
                strSQL = strSQL & "'" & strNo & "',"
                '      ���㷽ʽ_In Varchar2,
                strSQL = strSQL & "'" & strAdvance & "')"
                '      ��ע_In     ҽ��������ϸ.��ע%Type := Null
                zlDatabase.ExecuteProcedure strSQL, Me.Caption
            End If
            gcnOracle.CommitTrans
            
            Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, True, intInsure)
            Call SetBalanceVal(colBalance, p, str���㷽ʽ)
            
            gcnOracle.BeginTrans
        End If
        
        p = p + 1
        rsCharge.MoveNext
    Loop

    'ȫ���ɹ��������ܵĽ��㷽ʽ
    strAdvance = GetMedicareStr(colBalance)
    
    ExecuteClinicDelNo = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

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
    Dim varData As Variant, varTemp As Variant, i As Long, p As Integer, strAdvanceOld As String
    Dim colBalance As Collection '��¼���ŵ��ݱ��ս���
    Dim rsBalance As ADODB.Recordset
    Dim blnDo As Boolean, strTemp As String, strNo As String
    Dim rsCharge As ADODB.Recordset
    On Error GoTo errHandle
    
    If intInsure = 0 Then ExecuteClinicDelSwap = True: gcnOracle.CommitTrans: Exit Function
    strAllBalance = GetYBOldBalance(lng����ID, intInsure, lngԭ����ID)
    
    strAdvance = ""
    If MCPAR.����������� Then
        
        If Not mblnDelByNo Then
            strAdvance = lng����ID & "|" & "0"
            'ClinicDelSwap (ҽ���˷ѽ���)
            '������  ��������    ��/��   ԭ����˵��  �ֵ���˵��
            'lngStlID    long    IN  ��Ҫ�˷ѵķ��ü�¼�Ľ���ID(ԭ����ID)
            'bln�˷� Boolean IN  �������˷ѽ��׻��Ǹķѽ����ڵ��ñ��ӿ�
            'intInsure   Intger  In  ����
            'strAdvance  String  In  NULL    ����ID:���Ӵ������ID
            'ҽ�����Ը��ݳ���ID������ȡ��
            '        Out �˷ѽ��㣺���㷽ʽ1|���||���㷽ʽ2|���...
            '    Boolean ��������    True:���óɹ�,False:����ʧ��
            If Not gclsInsure.ClinicDelSwap(lngԭ����ID, , intInsure, strAdvance) Then
                gcnOracle.RollbackTrans
                Exit Function
            End If
            If strAdvance = CStr(lng����ID) & "|" & "0" Then strAdvance = ""
        Else
            If ExecuteClinicDelNo(lng����ID, intInsure, lng����ID, lngԭ����ID, strAdvance) = False Then
                gcnOracle.RollbackTrans
                Exit Function
            End If
        End If
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
        If Not zlInsureCheck(strAllBalance, strAdvance) Then
            '�޸�У�Ա�־
            ' Zl_���������շ�_ҽ������
            strSQL = "Zl_���������շ�_ҽ������("
            '  ����id_In   ������ü�¼.����id%Type,
            strSQL = strSQL & lng����ID & ","
            '  �������_In ����Ԥ����¼.�������%Type,
            strSQL = strSQL & "Null,"
            '  ���ս���_In Varchar2
            strSQL = strSQL & "Null)"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            gcnOracle.CommitTrans
            If Not mblnDelByNo Then Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, True, intInsure)
            ExecuteClinicDelSwap = True: Exit Function
        End If
        gcnOracle.CommitTrans: gcnOracle.BeginTrans
    End If
    '�˷Ѻ��շѲ�һ��ʱ,��ҪЧ��
        '���ӽ��㷽ʽΪ�յļ�¼
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
        strSQL = strSQL & "" & 3 & ","
        '  ����id_In     ������ü�¼.����id%Type,
        strSQL = strSQL & "" & lng����ID & ","
        '  ����id_In     ����Ԥ����¼.����id%Type,
        strSQL = strSQL & "" & lng����ID & ","
        '  ���㷽ʽ_In   Varchar2,
        strSQL = strSQL & "'" & strAdvance & "')"
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
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    '�޸�У�Ա�־
    ' Zl_���������շ�_ҽ������
    strSQL = "Zl_���������շ�_ҽ������("
    '  ����id_In   ������ü�¼.����id%Type,
    strSQL = strSQL & lng����ID & ","
    '  �������_In ����Ԥ����¼.�������%Type,
    strSQL = strSQL & "Null,"
    '  ���ս���_In Varchar2
    strSQL = strSQL & "Null)"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    gcnOracle.CommitTrans
    If MCPAR.����������� Then
        If Not mblnDelByNo Then Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, True, intInsure)
    End If
    ExecuteClinicDelSwap = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    Call ErrCenter
    Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, False, mCurBillType.intInsure)
End Function

Private Function ExecuteOneCardDelInterface(ByVal rsBalance As ADODB.Recordset, _
        ByVal lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:һ��ͨ�˷�(��)�ӿ�
    '���:lng����ID-����ID
    '     rsBalance-ԭ�����¼��
    '����:ִ�гɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-08-01 15:45:26
    '˵��:���ñ��ӿ�ǰ�����뿪ͨ����,��ɻ��쳣������ֹ����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCardNo As String, strSwap As String, strHsptCode As String 'ҽԺ����
    Dim i As Long, dblMoney As Double, strNos As String, strSQL As String
    Dim str���㷽ʽ As String
    
    '�ֶ�:����,��¼����,���㷽ʽ,ժҪ,�����ID,���������,���ƿ�,���㿨���,�������,����,������ˮ��, ����˵��,�������,У�Ա�־,ҽ��,���ѿ�id
    '     �Ƿ�����,�Ƿ�ȫ��,�Ƿ�����,��Ԥ��
    '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�)
    On Error GoTo errHandle
    rsBalance.Filter = "����=4"
    If rsBalance.RecordCount = 0 Then
        rsBalance.Filter = 0
        gcnOracle.CommitTrans
        ExecuteOneCardDelInterface = True: Exit Function
    End If
    
    'һ��ͨ(��):ֻ��ʹ��һ��
    With rsBalance
        .MoveFirst
        Do While Not .EOF
            dblMoney = dblMoney + Val(Nvl(!��Ԥ��))
            .MoveNext
        Loop
        dblMoney = RoundEx(dblMoney, 6)
        .MoveFirst
        If dblMoney = 0 Then
            rsBalance.Filter = 0: gcnOracle.CommitTrans
            ExecuteOneCardDelInterface = True: Exit Function
        End If
        strCardNo = Nvl(!����)
        str���㷽ʽ = Nvl(!���㷽ʽ)
        '���㷽ʽ|������|�������|����ժҪ||..
        str���㷽ʽ = str���㷽ʽ & "|" & -1 * dblMoney
        str���㷽ʽ = str���㷽ʽ & "|" & IIf(Trim(Nvl(!�������)) = "", " ", Trim(Nvl(!�������)))
        str���㷽ʽ = str���㷽ʽ & "| "
        
        'Zl_�����˷ѽ���_Modify
        '--��������_In:
        '--   0-ԭ����
        '--      ԭ������һ��ȫ��,����У�Ա�־��Ϊ1,ҽ�����óɹ���,����Ϊ2,��ɺ���0
        '--   1-��ͨ�˷ѷ�ʽ:
        '--     �ٽ��㷽ʽ_IN:��������,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ||.." ;Ҳ�������.
        '--   2.�������˷ѽ���:
        '--     �ٽ��㷽ʽ_IN:ֻ�ܴ���һ�����㷽ʽ,���������һЩ������Ϣ,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ"
        '--     �ڿ����ID_IN,����_IN,������ˮ��_IN,����˵��_In:��Ҫ����
        '--   3-ҽ������(�������ҽ���Ľ���,��Ҫ��ɾ��ԭҽ������,���´���ĸ���)
        '--     �ٽ��㷽ʽ_IN:��������,��ʽΪ:���㷽ʽ|������||.."
        '--     ����֧Ʊ��_In:������
        '--   4-���ѿ�����:
        '--     �ٽ��㷽ʽ_IN:����һ��ˢ���ſ�,��ʽΪ:�����ID|����|���ѿ�ID|���ѽ��||."
        '--     ����֧Ʊ��_In:������
        strSQL = "Zl_�����˷ѽ���_Modify("
        '  ��������_In   Number,
        strSQL = strSQL & "" & 2 & ","
        '  ����id_In     ������ü�¼.����id%Type,
        strSQL = strSQL & "" & mCurBillType.lng����ID & ","
        '  ����id_In     ����Ԥ����¼.����id%Type,
        strSQL = strSQL & "" & lng����ID & ","
        '  ���㷽ʽ_In   Varchar2,
        strSQL = strSQL & "'" & str���㷽ʽ & "',"
        '  ��Ԥ��_In     ����Ԥ����¼.��Ԥ��%Type := Null,
        strSQL = strSQL & "NULL,"
        '  �����id_In   ����Ԥ����¼.�����id%Type := Null,
        strSQL = strSQL & "NULL,"
        '  ����_In       ����Ԥ����¼.����%Type := Null,
        strSQL = strSQL & "'" & strCardNo & "',"
        '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
        strSQL = strSQL & "'" & Nvl(!������ˮ��) & "',"
        '  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
        strSQL = strSQL & "'" & Nvl(!����˵��) & "')"
        '  �ɿ�_In       ����Ԥ����¼.�ɿ�%Type := Null,
        '  �Ҳ�_In       ����Ԥ����¼.�Ҳ�%Type := Null,
        '  �����_In   ������ü�¼.ʵ�ս��%Type := Null,
        '  ����˷�_In   Number := 0,
        '  ԭ����id_In   ����Ԥ����¼.����id%Type := Null
    End With
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    If Not mobjICCard.ReturnSwap(strCardNo, strHsptCode, strSwap, dblMoney) Then
        gcnOracle.RollbackTrans
        MsgBox "һ��ͨ�˷ѽ��׵���ʧ��,���ܼ����˷Ѳ�����", vbExclamation, gstrSysName
        Exit Function
    End If
    gcnOracle.CommitTrans
    ExecuteOneCardDelInterface = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
End Function

Private Function isValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����˷��Ƿ�Ϸ�
    '����:�Ϸ�����true,���򷵻�False
    '����:���˺�
    '����:2014-07-04 11:23:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim arrNo As Variant, strYPNos As String, blnҩƷ As Boolean, blnSel As Boolean
    Dim i As Long, strDelNOs As String, strNo As String, str����Ա���� As String
    Dim varTemp As Variant, rsTemp As ADODB.Recordset
    On Error GoTo errHandle
 
    '��������Ƿ���ȷ
    If mCurBillType.strNos = "" Then
        MsgBox "��������Ҫ�˷ѵĵ��ݡ�", vbInformation, gstrSysName
        If txtNO.Visible Then txtNO.SetFocus: Exit Function
    End If
    
    '��鱾�ν��㵥�����Ƿ�����˷��쳣���ݣ������ڣ�����������˷�
    If CheckIsExistDelErrBill(mCurBillType.strNos, str����Ա����) Then
        MsgBox "ע�⣺" & vbCrLf & _
            "    ���ν����д����쳣���˷Ѽ�¼�����ȶ�����������˷ѣ�" & _
            IIf(str����Ա���� <> UserInfo.����, vbCrLf & "    ��ʾ���쳣�����ǲ���Ա��" & str����Ա���� & "����ȡ�ġ�", ""), vbInformation + vbOKOnly, gstrSysName
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
    
    If strDelNOs <> "" And gbln�˷�����ģʽ Then
        Set rsTemp = GetApply(strDelNOs, 1)
        varTemp = Split(strDelNOs, ",")
        For i = 0 To UBound(varTemp)
            strNo = varTemp(i)
            rsTemp.Filter = "NO='" & strNo & "' And ״̬<>2"
            If rsTemp.RecordCount = 0 Then
                Screen.MousePointer = 0
                MsgBox "���ȶԵ���:" & strNo & " �����˷����룡", vbInformation, gstrSysName
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
    
    If blnҩƷ Then
        If strYPNos <> "" Then strYPNos = Mid(strYPNos, 2)
        If zlCheckDrugIsPutDrug(strYPNos) = False Then Exit Function
    End If
    
    'ҽ�����
    If mCurBillType.intInsure <> 0 Then
        If gclsInsure.CheckInsureValid(mCurBillType.intInsure) = False Then Exit Function
    End If
    
    If zlCheckIsMzToZY(strDelNOs, 1) Then
          MsgBox "ע��:" & vbCrLf & _
            "    �õ����Ѿ����������תסԺ���� " & vbCrLf & _
            "    ���Ѿ�������������תסԺ����,�������˷�", vbInformation + vbOKOnly, gstrSysName
          Exit Function
    End If
    
    If CheckBillExistReplenishData(0, mlng�������) = True Then
        MsgBox "ѡ����˷Ѽ�¼������ҽ��������㣬����������˷Ѳ�����", vbInformation, gstrSysName
        Exit Function
    End If
    
    '105432,���������㷽ʽ��Ч�Լ��
    If ThreeBalanceCheck(mrsBalance, mrs���㷽ʽ, mcllForceDelToCash) = False Then Exit Function
    
    isValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function ThreeBalanceCheck(ByVal rsBalance As ADODB.Recordset, ByVal rs���㷽ʽ As ADODB.Recordset, _
    ByRef cllForceDelToCash As Collection) As Boolean
    '���������㷽ʽ��Ч�Լ��
    '��Σ�
    '   rsBalance ��������
    '   rs���㷽ʽ ���շѡ����ϵ����н��㷽ʽ
    '���Σ�
    '   cllForceDelToCash ǿ��������Ϣ��Array(����Ա,���������)
    '���أ����ͨ��������True�����򣬷���False
    '105432
    Dim objCards As Cards, objCard As Card
    Dim cllFeeBalance As New Collection, i As Integer
    Dim blnFind As Boolean, blnQuestion As Boolean
    Dim str����Ա As String, strKey As String
    Dim dblMoney  As Double
    
    On Error GoTo errHandler
    Set cllForceDelToCash = New Collection
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
                'Array(���㷽ʽ,�����ID,�Ƿ�����,���������,��Ԥ��)
                cllFeeBalance.Add Array(Nvl(!���㷽ʽ), Val(Nvl(!�����ID)), Val(Nvl(!�Ƿ�����)), Nvl(!���������), dblMoney), strKey
            End If
            .MoveNext
        Loop
    End With
    If cllFeeBalance.Count = 0 Then ThreeBalanceCheck = True: Exit Function
    
    If Not gobjSquare Is Nothing Then
        ' zlGetCards(ByVal BytType As Byte)
        '   ���:bytType-  0-����ҽ�ƿ�;
        '                    1-���õ�ҽ�ƿ�,
        '                    2-���д��������˻���������
        '                    3-���õ������˻���ҽ�ƿ�
        Set objCards = gobjSquare.objSquareCard.zlGetCards(3)
    End If
    For i = 1 To cllFeeBalance.Count
        blnQuestion = False
        '���㷽ʽ���
        If rs���㷽ʽ Is Nothing Then
            If MsgBox("���㷽ʽ��" & cllFeeBalance(i)(0) & "��δ���ã��ý��㷽ʽ֧���Ľ�����Ϊ�������㷽ʽ���Ƿ������", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            blnQuestion = True
        Else
            rs���㷽ʽ.Filter = "����='" & cllFeeBalance(i)(0) & "'" '���㷽ʽҪ������"����"Ӧ�ó��ϲ���ʹ��
            If rs���㷽ʽ.EOF Then
                If MsgBox("���㷽ʽ��" & cllFeeBalance(i)(0) & "��δ���ã��ý��㷽ʽ֧���Ľ�����Ϊ�������㷽ʽ���Ƿ������", _
                    vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                blnQuestion = True
            End If
        End If
        
        If blnQuestion = False Then
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
        End If
        
        If blnQuestion And cllFeeBalance(i)(2) = 0 Then 'ǿ������
            If str����Ա = "" Then '���ֿ����ʱֻ��֤һ��
                If zlStr.IsHavePrivs(GetPrivFunc(glngSys, 1151), "�����˿�ǿ������") Then
                    str����Ա = UserInfo.����
                Else
                    str����Ա = zlDatabase.UserIdentifyByUser(Me, "��" & cllFeeBalance(i)(3) & "��ǿ�����֣�Ȩ����֤��", _
                        glngSys, mlngModule, "�����˿�ǿ������", , True)
                    If str����Ա = "" Then Exit Function
                End If
                'Array(����Ա,���������)
                cllForceDelToCash.Add Array(str����Ա, cllFeeBalance(i)(3))
            End If
        End If
    Next
    ThreeBalanceCheck = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function InputFactNo(ByRef lng����ID As Long, ByRef strInvoice As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������Ч�ķ�Ʊ��
    '���:
    '     lng����ID-��ǰ������ID
    '����:���صķ�Ʊ��
    '����:����ɹ�������true,���򷵻�False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnValid As Boolean
    
    On Error GoTo errHandle
    Do
        '����Ʊ�����ö�ȡ
        blnValid = False
        
        If gblnStrictCtrl Then
            If zlCheckInvoiceValied(lng����ID, 1, , mlngShareUseID, mstrUseType) = False Then Exit Function
            strInvoice = GetNextBill(lng����ID)
        Else
            strInvoice = zlStr.Increase(UCase(zlDatabase.GetPara("��ǰ�շ�Ʊ�ݺ�", glngSys, mlngModule)))
        End If
        
        If strInvoice = "" Then
            If frmInputBox.InputBox(Me, "��ʼ��Ʊ��", "" & _
                 "�������뽫Ҫʹ�õĿ�ʼƱ�ݺ��룺", 30, 1, False, False, strInvoice, _
                False, Me.Left + 1500, Me.Top + 1500) = False Then Exit Function
        End If
                    
        '�û�ȡ������,��ֹ����
        If strInvoice = "" Then Exit Function
        If gblnStrictCtrl Then
            If zlCheckInvoiceValied(lng����ID, 1, strInvoice, mlngShareUseID, mstrUseType) Then blnValid = True
        Else
            blnValid = True
        End If
    Loop While Not blnValid
    
    InputFactNo = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
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
    Dim strTemp As String, varTemp As Variant, cllPro As Collection
    Dim arrNo As Variant, k As Long
    Dim i As Long, j As Long, dtDelDate As Date, lngCount As Long
    Dim strNo As String, blnTrans As Boolean
    Dim colOrder As New Collection
    Dim lng����ID As Long, lng����ID As Long, lng������� As Long
    Dim lngԭ����ID As Long, blnAll�����˷� As Boolean, blnCur�����˷� As Boolean
    Dim blnȫ�� As Boolean, lngCheck����ID As Long, intCheckInsure As Integer
    Dim strYBPati As String, strPrintNOInfor As String, strInvoice As String
    Dim str���  As String, strCurSelNos As String, strReclaimInvoice As String
    Dim strInvoices As String, lng����ID As Long
    Dim strSQL As String, strCmdCaptions As String
    Dim blnԭ���� As Boolean
    Dim cur����͸֧ As Currency, str���ս�� As String 'curʵ�պϼ�;cur����ͳ��;curȫ�Ը�;cur���Ը�
    Dim strReturn As String, strReturnRecipt As String '�˷Ѵ�����Ϣ����ʽ��NO,ҩ��ID|NO,ҩ��ID|��
    Dim strPartSelectNos As String '����ѡ��ĵ���
    Dim strPartDoNos As String 'ȫѡ�����ڲ���ִ�еĵ���
    Dim bln�ֱ��ӡ As Boolean

    If isValied = False Then Exit Function
 
    lngԭ����ID = mCurBillType.lngԭ����ID
    bln�ֱ��ӡ = gTy_Module_Para.bln�ֱ��ӡ And mblnOnePatiPrint = False
      
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
            strPrintNOInfor = strPrintNOInfor & ";" & strNo & ":" & str���
            blnCur�����˷� = Not BillDeleteAllNew(strNo, 1)
            If blnCur�����˷� Then strPartDoNos = strPartDoNos & "," & strNo '���ڲ���ִ�еĵ���
            
            If UBound(Split(str���, ",")) + 1 = lngCount And blnCur�����˷� = False Then str��� = ""
            blnCur�����˷� = Not (Not blnCur�����˷� And str��� = "")
            If blnCur�����˷� And str��� <> "" Then strPartSelectNos = strPartSelectNos & "," & strNo '����ѡ��ĵ���
            
            If blnCur�����˷� Then blnAll�����˷� = True '���ŵ���Ϊ�����˷�,�����е���Ϊ�����˷�
            colOrder.Add str���, "_" & strNo
        Else
            blnAll�����˷� = True                       '���ŵ��ݲ��˷�,�����е���Ϊ�����˷�
            colOrder.Add "δѡ��", "_" & strNo
        End If
    Next
    If strPartSelectNos <> "" Then strPartSelectNos = Mid(strPartSelectNos, 2)
    If strPartDoNos <> "" Then strPartDoNos = Mid(strPartDoNos, 2)
    
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
    
    If blnAll�����˷� Then
        If mCurBillType.intInsure > 0 And (MCPAR.������ȫ�� Or mblnDelByNo) Then '86176
            If strPartSelectNos <> "" Then
                MsgBox "����[" & strPartSelectNos & "]�������ս�����ã����������˷ѡ�", vbInformation, gstrSysName
                Exit Function
            ElseIf strPartDoNos <> "" Then
                MsgBox "����[" & strPartDoNos & "]�������ս�����ã�������һЩ��Ŀ�Ѿ�ִ�У����������˷ѡ�", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        '56963
        If gTy_Module_Para.bytƱ�ݷ������ <> 0 And strReclaimInvoice = "" Then
            strReclaimInvoice = zlGetReclaimInvoice(Mid(strPrintNOInfor, 2))
        End If
        
        If Not (gTy_Module_Para.bytƱ�ݷ������ <> 0 And strReclaimInvoice <> "") Then
            If zlStr.IsHavePrivs(mstrPrivs, "�����˷�") = False Then
                MsgBox "��û��Ȩ��ִ�в����˷Ѳ�����", vbInformation, gstrSysName
                vsBill.SetFocus: Exit Function
            End If
            If gTy_Module_Para.bln������ Then
                MsgBox "�Զ���ȡ������ʱ���������˷ѡ�", vbInformation, gstrSysName: vsBill.SetFocus: Exit Function
            End If
            '���˺� ����:27352 ����:2010-01-13 10:26:08
            If zlStr.IsHavePrivs(mstrPrivs, "�˷Ѻ��շ�Ʊ") Then
                
                If frmReInvoice.ShowMe(Me, strCurSelNos, Val(txtAllTotal.Text), Val(txt�˿�ϼ�.Text), strInvoices) = False Then
                    vsBill.SetFocus: Exit Function
                End If
            End If
        End If
    End If
    
    If mCurBillType.intInsure <> 0 And MCPAR.ҽ���ӿڴ�ӡƱ�� Then
        If InputFactNo(lng����ID, strInvoice) = False Then Exit Function
    End If
    dtDelDate = zlDatabase.Currentdate
    blnȫ�� = CheckIsAllDel(mCurBillType.strAllNOs)
    blnԭ���� = blnȫ��
    If blnԭ���� Then
        blnԭ���� = Not zlExistDelFeeChargeBill(mCurBillType.strAllNOs)
    End If
    
    '����Ҫִ�е�SQL
    lng����ID = zlDatabase.GetNextId("���˽��ʼ�¼")
    lng������� = -1 * lng����ID
    mCurBillType.strDelNOs = ""
    For i = UBound(arrNo) To 0 Step -1
        strNo = arrNo(i)
        If bln�ֱ��ӡ And gTy_Module_Para.bytƱ�ݷ������ = 0 Then
            blnȫ�� = CheckIsAllDel(strNo)
        End If
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
            strSQL = strSQL & lng����ID & ","
            '  ����Ʊ��_In Number:=0
            If blnȫ�� And mblnOnePatiPrint And gTy_Module_Para.bytƱ�ݷ������ <> 0 Then
                '���ݱ���ʵ�ʴ�ӡ����Ҫ����Ʊ��(�Ͳ�����������,��������ʽ����Ҫ�ں������������
                strSQL = strSQL & "0)" '�����˴�ӡ�����л���Ʊ��,�ں��洦��
            Else
                strSQL = strSQL & "" & IIf(blnȫ��, "1", "0") & ")"
            End If
            zlAddArray cllPro, strSQL
            mCurBillType.strDelNOs = mCurBillType.strDelNOs & "," & strNo
        End If
    Next
    blnȫ�� = CheckIsAllDel(mCurBillType.strAllNOs)
    If mCurBillType.intInsure <> 0 And MCPAR.����������� Then
        If Not mblnDelByNo Then
            If Not blnȫ�� Then lng����ID = zlDatabase.GetNextId("���˽��ʼ�¼")
            'ֻ��ҽ������,�Ż����������ȡ�����
            'Zl_�����շѼ�¼_����
            strSQL = "Zl_�����շѼ�¼_����("
            '  ԭ����id_In ������ü�¼.����id%Type,
            strSQL = strSQL & "" & lngԭ����ID & ","
            '  ����id_In   ������ü�¼.����id%Type,
            strSQL = strSQL & "" & lng����ID & ","
            '  ���ս���id_In ������ü�¼.����id%Type
            strSQL = strSQL & "" & IIf(lng����ID = 0, "NULL", lng����ID) & ","
            '  �ſ�ҽ������_In Varchar2:=Null
            strSQL = strSQL & "'" & GetYBTOCash(mCurBillType.lng����ID, mCurBillType.intInsure) & "')"
            zlAddArray cllPro, strSQL
            '����ҽ���ӿ�
            '�Ȼ���Ʊ�ݣ�Ԥ����֮���ٲ���Ʊ��
            If MCPAR.ҽ���ӿڴ�ӡƱ�� Then '81684
                If Not blnȫ�� Then 'Ԥ����֮���ٷ���Ʊ��
                    '56963,77058
                    strSQL = "zl_�����շѼ�¼_RePrint('" & strNo & "',NULL," & ZVal(lng����ID) & ",'" & UserInfo.���� & "'," & _
                        "To_Date('" & Format(dtDelDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),1,1)"
                    zlAddArray cllPro, strSQL
                ElseIf Not (gTy_Module_Para.bytƱ�ݷ������ <> 0 And strReclaimInvoice <> "") Then  'ȫ�˷�ҲҪ����Ʊ�ݺţ�����ҽ��
                    strSQL = "zl_�����շѼ�¼_RePrint('" & strNo & "','" & strInvoice & "'," & ZVal(lng����ID) & ",'" & UserInfo.���� & "'," & _
                        "To_Date('" & Format(dtDelDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),1,1)"
                    zlAddArray cllPro, strSQL
                End If
            End If
        Else
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
            strSQL = strSQL & "" & 3 & ","
            '  ����id_In     ������ü�¼.����id%Type,
            strSQL = strSQL & "" & mCurBillType.lng����ID & ","
            '  ����id_In     ����Ԥ����¼.����id%Type,
            strSQL = strSQL & "" & lng����ID & ","
            '  ���㷽ʽ_In   Varchar2,
            strSQL = strSQL & "'" & zlGetYBBalanceNo(lngԭ����ID, mCurBillType.strDelNOs, mCurBillType.lng����ID, mCurBillType.intInsure, True) & "')"
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
        End If
    Else
        '���ӽ��㷽ʽΪ�յļ�¼
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
        strSQL = strSQL & "" & lng����ID & ","
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
    End If
    
    '����ҽ��
    If mCurBillType.intInsure <> 0 And MCPAR.����������� Then
        If Not blnȫ�� And Not mblnDelByNo Then
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
        End If
        blnTrans = True
        zlExecuteProcedureArrAy cllPro, Me.Caption, True
        
        If Not blnȫ�� And Not mblnDelByNo Then
            '�������ռ�¼�ı�����Ϣ
            '77951,Ƚ����,2014-9-16��
            If ExecuteInsureInfoUpdate(lng����ID, str���ս��) = False Then Exit Function
            '��ȡ�������
            cur����͸֧ = mdbl����͸֧
            mdbl������� = gclsInsure.SelfBalance(mCurBillType.lng����ID, CStr(Split(strYBPati, ";")(1)), 10, cur����͸֧, mCurBillType.intInsure)
            mdbl����͸֧ = cur����͸֧
        End If
        If ExecuteClinicDelSwap(mCurBillType.lng����ID, mCurBillType.intInsure, lng����ID, lngԭ����ID) = False Then Exit Function
        Set cllPro = New Collection
        
        '���½����շѴ���
        '77058
        If Not blnȫ�� And Not mblnDelByNo Then
            gcnOracle.BeginTrans
            If ExcuteInsureReCharge(mCurBillType.lng����ID, mCurBillType.intInsure, lng����ID, lng�������, str���ս��, _
                        strNo, lng����ID, strInvoice, dtDelDate) = False Then Exit Function
        End If
        blnTrans = False
    End If
    
    '2.����һ��ͨ(�ϰ汾)
    If mCurBillType.blnExistOnCard Then
ReDOOneCard:
        If CheckOnCardValied(mrsBalance) = False Then Exit Function
        blnTrans = True
        zlExecuteProcedureArrAy cllPro, Me.Caption, True
        If Not ExecuteOneCardDelInterface(mrsBalance, lng����ID) Then
            If mCurBillType.intInsure <> 0 Then
                strCmdCaptions = "�쳣����(&C)|��ʾ������һ��ͨ����,���ݽ����쳣��ʽ����,�������ڽ����д���"
                strCmdCaptions = strCmdCaptions & ";����(&R)|��ʾ���µ���һ��ͨ���㽻��"
                If frmVerfyCodeInput.ShowMsg(Me, "����[" & mCurBillType.strDelNOs & "]�Ѿ��˷ѳɹ�,��һ��ͨ����ʧ��,[�쳣����]����������֤��,���鲻�����쳣���ݱ���", strCmdCaptions) = False Then
                     gcnOracle.BeginTrans: blnTrans = True
                    GoTo ReDOOneCard:
                End If
            End If
            Exit Function
        End If
        Set cllPro = New Collection: blnTrans = False
    End If
    
    '4.��ʾ�������
    Dim frmBalance As New frmClinicDelBalance, objDelBalance As New clsCliniDelBalance
    
    Set objDelBalance.rsBalance = mrsBalance
    Set objDelBalance.rs���㷽ʽ = mrs���㷽ʽ
    If strPrintNOInfor <> "" Then strPrintNOInfor = Mid(strPrintNOInfor, 2)
    mCurBillType.lng������� = lng������� '��¼���ڴ�ӡ��Ʊ
    
    With objDelBalance
        .intInsure = mCurBillType.intInsure
        .CurDelNos = mCurBillType.strDelNOs
        .AllNos = mCurBillType.strAllNOs
        .PrintNOs = strPrintNOInfor
        
        .PatiUseType = mstrUseType
        .SaveBilled = cllPro.Count = 0
        .ShareUserID = mlngShareUseID
        .����ID = mCurBillType.lng����ID
        .����ID = lng����ID
        .��ǰ��Ʊ�� = strInvoice
        .���շ�Ʊ = strInvoices
        .������� = lng�������
        .����ID = lng����ID
        .ȱʡ���㷽ʽ = mCurBillType.str���㷽ʽ
        .�˷Ѻϼ� = -1 * GetDelMoney
        .�ѱ� = mCurBillType.str�ѱ�
        .���� = mCurBillType.str����
        .�Ա� = mCurBillType.str�Ա�
        .���� = mCurBillType.str����
        .ҽ������Ʊ�� = MCPAR.ҽ������Ʊ��
        .ԭ����ID = mCurBillType.lngԭ����ID
        .�˷�ʱ�� = dtDelDate
        .�����˷� = Not blnȫ��
        .ԭ���� = blnԭ����
        .blnOnePatiPrint = mblnOnePatiPrint
        .strOnePatiPrintNos = mstrOnePatiPrintNos
    End With
    Call GetAsyncKeyState(VK_RETURN)
    If frmBalance.zlDelCharge(Me, EM_FUN_�˷�, mlngModule, mstrPrivs, objDelBalance, cllPro, , mcllForceDelToCash) = False Then
        Exit Function
    End If
    
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
    ExecDelete = True
    Exit Function
Errhand:
    If blnTrans Then
        gcnOracle.RollbackTrans
        Call ErrCenter
    Else
        If ErrCenter = 1 Then
            Resume
        End If
    End If
    If Err.Number <> 0 Then Call SaveErrLog
    
    '�ж���ʾ,����ӡ�������˷Ѻ��ٴ�ӡ���Լ�ѡ���ش�
    Call ShowErrBill(mCurBillType.strDelNOs, strNo)
End Function



Private Sub PrintDelBill(ByVal strAllNOs As String, ByVal strCurDelNOs As String, _
    ByVal strNo As String, _
    ByVal lng����ID As Long, _
    ByVal dtDateDel As Date, ByVal blnAll�����˷� As Boolean, _
    ByVal strInvoices As String, ByVal strReclaimInvoice As String, _
    Optional blnOnePatiPrint As Boolean, Optional strOnePatiPrintNos As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ӡ���Ʊ��
    '���: strAllNOs-��ǰ�漰�����е���
    '       strCurDelNOs-��ǰ�˷ѵĵ���
    '       dtDateDel-�˷�����
    '       strInvoices-ѡ��ķ�Ʊ��(��ģʽ)
    '       strReclaimInvoice-���յķ�Ʊ��
    '       blnOnePatiPrint-�Ƿ񰴲��˴�ӡƱ��
    '       strOnePatiPrintNos-�����˴�ӡ�ĵ��ݺ�(����ö��ŷ���,��:a,b,c
    '����:
    '����:���˺�
    '����:2013-05-27 16:41:06
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intInvoiceFormat As Integer, blnPrint As Boolean
    Dim str��Ʊ�� As String, intƱ������ As Integer
    Dim strSQL As String, strTempNO As String, i As Integer
    Dim lng��ӡID As Long
    Dim strBillPrintNos As String
    Dim varNos As Variant, strNotAllDelNos As String '��ǰ�����˵��ݣ���ʵ�ʷ���Ʊ�ŷֱ��ӡʱ��¼�����˵���
    Dim strPriceGrade As String

    On Error GoTo errHandle
    If InStr(strAllNOs, "'") = 0 Then
        strAllNOs = "'" & Replace(strAllNOs, ",", "','") & "'"
    End If
    strBillPrintNos = strAllNOs

    If InStr(strCurDelNOs, "'") = 0 Then
        strCurDelNOs = Mid(strCurDelNOs, 2)
        strCurDelNOs = "'" & Replace(strCurDelNOs, ",", "','") & "'"
    End If

    If blnOnePatiPrint Then
        '�����˲���Ʊ�ݣ���Ҫ��������ʱ������
        Dim blnAllDel As Boolean
        If zlSaveTempPrintData(strOnePatiPrintNos, mlng����ID, "", lng��ӡID) = False Then GoTo PrintList
        If zlChargeBillIsAllDel("", lng��ӡID, blnAllDel, strBillPrintNos) = False Then GoTo PrintList
        
        If InStr(strBillPrintNos, "'") = 0 Then strBillPrintNos = "'" & Replace(strBillPrintNos, ",", "','") & "'"

        If blnAllDel Then
            If gTy_Module_Para.bytƱ�ݷ������ <> 0 Then
                'ȫ���ˣ���ֱ�Ӵ�ӡ�嵥(����Ʊ���Ѿ����˷ѵ����д�����)
                str��Ʊ�� = strReclaimInvoice
                Call zlExeCuteBillNoSplit(False, 4, mlng����ID, strAllNOs, lng����ID, "", dtDateDel, 1, str��Ʊ��, intƱ������, lng��ӡID)
'            Else
'                'Zl_�����շѼ�¼_Reprint
'                strSQL = "Zl_�����շѼ�¼_Reprint("
'                '  No_In         ������ü�¼.No%Type,
'                strSQL = strSQL & "'" & Split(strBillPrintNos & ",", ",")(0) & "',"
'                '  Ʊ�ݺ�_In     Ʊ��ʹ����ϸ.����%Type,
'                strSQL = strSQL & "NULL,"   'ȫ���ջأ�û�з���Ʊ��
'                '  ����id_In     Ʊ��ʹ����ϸ.����id%Type,
'                strSQL = strSQL & "" & 0 & ","
'                '  ʹ����_In     Ʊ��ʹ����ϸ.ʹ����%Type,
'                strSQL = strSQL & "'" & UserInfo.���� & "',"
'                '  ʹ��ʱ��_In   Ʊ��ʹ����ϸ.ʹ��ʱ��%Type,
'                strSQL = strSQL & "to_date('" & Format(dtDateDel, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
'                '  �˷�_In       Number := 0,
'                strSQL = strSQL & "1,"
'                '  Ʊ������_In   Number := 0,
'                strSQL = strSQL & "0,"
'                '  �ջ�Ʊ�ݺ�_In Varchar2 := Null,
'                strSQL = strSQL & "NULL,"
'                '  Ʊ��_In Number:=1
'                strSQL = strSQL & "1)"
'                zlDatabase.ExecuteProcedure strSQL, Me.Caption
            End If
            GoTo PrintList:
            Exit Sub
        End If
    End If
    
    If (Not blnAll�����˷� And blnOnePatiPrint = False) Or (blnAllDel And blnOnePatiPrint) Then
         '˰�ز���ȫ��ʱ�ջش���(ȫ��ʱ��zl_�����շѼ�¼_DELETE�����ջ�Ʊ��)
        If Not gobjTax Is Nothing And gblnTax Then
            gstrTax = gobjTax.zlTaxOutErase(gcnOracle, strAllNOs)
            If gstrTax <> "" Then MsgBox gstrTax, vbExclamation, gstrSysName
        End If
        GoTo PrintList:
        Exit Sub
    End If
    '77058
    If blnAll�����˷� And mCurBillType.intInsure <> 0 And MCPAR.ҽ���ӿڴ�ӡƱ�� And blnOnePatiPrint = False Then GoTo PrintList
    
    '�����˷�ʱ�ջز��ش�,�������Ų����˺��˶����е�ĳ����
    If gTy_Module_Para.bytƱ�ݷ������ <> 0 And strReclaimInvoice <> "" Then
        '����Ʊ�ݷ�������ӡ
        '��Ԥ��,��Ʊ���Ƿ����
        str��Ʊ�� = strReclaimInvoice
        If zlExeCuteBillNoSplit(True, 4, mlng����ID, strAllNOs, lng����ID, "", dtDateDel, 1, str��Ʊ��, intƱ������, , , lng��ӡID) = False Then GoTo PrintList:
        If intƱ������ = 0 Then
            'ֻ����Ʊ��,������ӡ
            str��Ʊ�� = strReclaimInvoice
            Call zlExeCuteBillNoSplit(False, 4, mlng����ID, strAllNOs, lng����ID, "", dtDateDel, 1, str��Ʊ��, intƱ������, , , lng��ӡID)
            GoTo PrintList:
        End If
        
        '0-����ӡ;1-�Զ���ӡ;2-��ʾ��ӡ
        Select Case mintInvoicePrint
        Case 0
            blnPrint = False
        Case 1
            blnPrint = True
        Case 2
            blnPrint = MsgBox("�Ƿ��ӡƱ�ݣ�", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes
        End Select
        
        '�ش��ջط�Ʊ
        If blnPrint Then
            intInvoiceFormat = IIf(strReclaimInvoice = "" And gTy_Module_Para.bytƱ�ݷ������ <> 0 And blnOnePatiPrint = False, mintOldInvoiceFormat, mintInvoiceFormat)
            If gintPriceGradeStartType >= 2 Then
                strPriceGrade = GetPriceGradeFromNos(strAllNOs)
            Else
                strPriceGrade = gstr��ͨ�۸�ȼ�
            End If
            Call RePrintCharge(1, strBillPrintNos, Me, mlng����ID, strReclaimInvoice, True, dtDateDel, _
                intInvoiceFormat, , , mlngShareUseID, mstrUseType, blnOnePatiPrint, strPriceGrade)
        End If
        GoTo PrintList:
        Exit Sub
    End If
    
    If strInvoices = "�޿���Ʊ��" Or strInvoices = "" Then 'a.�ջز����´�ӡ�����վ�
        '0-����ӡ;1-�Զ���ӡ;2-��ʾ��ӡ
        If gTy_Module_Para.bln�ֱ��ӡ And blnOnePatiPrint = False Then
            If mintInvoicePrint = 0 Then
                blnPrint = False
            Else
                strNotAllDelNos = ""
                varNos = Split(Replace(strCurDelNOs, "'", ""), ",")
                For i = 0 To UBound(varNos)
                    If CheckIsAllDel(varNos(i), True) = False Then
                        strNotAllDelNos = strNotAllDelNos & ",'" & varNos(i) & "'"
                    End If
                Next
                If strNotAllDelNos <> "" Then strNotAllDelNos = Mid(strNotAllDelNos, 2)
                
                '���ڲ����˵ĵ��ݣ���Ҫ�ش�
                If strNotAllDelNos = "" Then
                    blnPrint = False
                Else
                    If mintInvoicePrint = 1 Then
                        blnPrint = True
                    Else
                        blnPrint = MsgBox("�Ƿ��ӡƱ�ݣ�", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes
                    End If
                End If
            End If
        Else
            '0-����ӡ;1-�Զ���ӡ;2-��ʾ��ӡ
            Select Case mintInvoicePrint
            Case 0
                blnPrint = False
            Case 1
                blnPrint = True
            Case 2
                blnPrint = MsgBox("�Ƿ��ӡƱ�ݣ�", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes
            End Select
        End If

        If blnPrint Then
            intInvoiceFormat = IIf(strReclaimInvoice = "" And gTy_Module_Para.bytƱ�ݷ������ <> 0 And blnOnePatiPrint = False, mintOldInvoiceFormat, mintInvoiceFormat)
            If gintPriceGradeStartType >= 2 Then
                strPriceGrade = GetPriceGradeFromNos(strAllNOs)
            Else
                strPriceGrade = gstr��ͨ�۸�ȼ�
            End If
            If gTy_Module_Para.bln�ֱ��ӡ = True And blnOnePatiPrint = False Then
                Call RePrintCharge(1, strCurDelNOs, Me, mlng����ID, strReclaimInvoice, True, dtDateDel, _
                intInvoiceFormat, , , mlngShareUseID, mstrUseType, , strPriceGrade)
            Else
                Call RePrintCharge(1, strBillPrintNos, Me, mlng����ID, strReclaimInvoice, True, dtDateDel, _
                intInvoiceFormat, , , mlngShareUseID, mstrUseType, blnOnePatiPrint, strPriceGrade)
            End If
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
    '�˷ѷ�Ʊ(��Ʊ)��ӡ��91998
    '0-����ӡ;1-�Զ���ӡ;2-��ʾ��ӡ
    If mintInvoicePrintDel = 1 Then
        Call PrintDelCharge(mCurBillType.lng�������, Me, mlng����ID, True, dtDateDel, mintInvoiceFormatDel, , , mlngShareUseID, mstrUseType)
    ElseIf mintInvoicePrintDel = 2 Then
        If MsgBox("�Ƿ��ӡ�˷�Ʊ��(��Ʊ)��", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
            Call PrintDelCharge(mCurBillType.lng�������, Me, mlng����ID, True, dtDateDel, mintInvoiceFormatDel, , , mlngShareUseID, mstrUseType)
        End If
    End If
    
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
    If mCurBillType.intInsure <> 0 And MCPAR.�˷Ѻ��ӡ�ص� And InStr(1, mstrPrivs, "ҽ���˷ѻص�") > 0 Then
        '����:35248
        'If strCurDelNOs <> "" Then strCurDelNOs = Mid(strCurDelNOs, 2) 'Ƚ����,2014-9-10,�ָ�������ǰ��ȥ�������ﲻ���ٴ���
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_4", Me, "NO=" & strCurDelNOs, 2)
    End If
    '77058
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
    '77635:���ϴ�,2014/9/9,����ԭ�򳤶ȿ���
    If zlCommFun.ActualLen(txt�˷�ժҪ.Text) > 100 Then
        MsgBox "����ԭ������������� " & 100 & " ���ַ��� " & 50 & " �����֣�", vbInformation, gstrSysName
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
    '����:2011-04-29 14:17:33
    '����:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInvoiceFormat As String
    If mbytMode = 0 Then Exit Sub
    mstrUseType = "": mlngShareUseID = 0: mintInvoiceFormat = 0
    mstrUseType = zl_GetInvoiceUserType(mCurBillType.lng����ID, 0, mCurBillType.intInsure)
    mlngShareUseID = zl_GetInvoiceShareID(mlngModule, mstrUseType)
    mintInvoiceFormat = zl_GetInvoicePrintFormat(mlngModule, mstrUseType, mintOldInvoiceFormat, mblnOnePatiPrint)
    mintInvoiceFormatDel = zl_GetInvoicePrintFormat(mlngModule, mstrUseType, , , True)
    mintInvoicePrint = zl_GetInvoicePrintMode(mlngModule, mstrUseType)
    mintInvoicePrintDel = zl_GetInvoicePrintMode(mlngModule, mstrUseType, True)
End Sub

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
        .Editable = flexEDKbdMouse
    End With
End Sub
Private Sub LoadBalanceInfor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����տ���㷽ʽ
    '����:���˺�
    '����:2014-07-02 14:46:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str���㷽ʽ As String, lngRow As Long
    Dim lngCol As Long, i As Long, intSign As Integer

    If mrsBalance Is Nothing Then Exit Sub
    If mrsBalance.State <> 1 Then Exit Sub
    intSign = IIf(mstrDelTime <> "", -1, 1) '����,�����������
    '�ֶ�:���� ,����ID, ��¼����, ���㷽ʽ, ժҪ, �����ID, ���������, ���ƿ�, ���㿨���, �������, ����, ������ˮ��, ����˵��, �������, У�Ա�־, ҽ��, ���ѿ�id
    '            �Ƿ�����,�Ƿ�ȫ��,�Ƿ�����,��Ԥ��
    '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
    lngRow = 0
    mrsBalance.Filter = 0
    mrsBalance.Sort = "����,���㷽ʽ"
    With vsBalance
        .Redraw = flexRDNone
        Call ClearBalance
        If mrsBalance.RecordCount <> 0 Then mrsBalance.MoveFirst
        Do While Not mrsBalance.EOF
            '--����:52530
            If Val(Nvl(mrsBalance!����)) = 1 Then
                str���㷽ʽ = "��Ԥ���"
            Else
                str���㷽ʽ = Nvl(mrsBalance!���㷽ʽ, "δ����")
            End If
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
                If mbytMode = EM_MULTI_�˷� Then
                    .ColData(lngCol) = "ժҪ:" & IIf(Val(Nvl(mrsBalance!�˷�)) = 1, "", Nvl(mrsBalance!ժҪ))
                    .ColData(lngCol + 1) = "�������:" & IIf(Val(Nvl(mrsBalance!�˷�)) = 1, "", Nvl(mrsBalance!�������))
                ElseIf mbytMode = EM_MULTI_�쳣���� Then
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
                If Val(Nvl(mrsBalance!��������)) <> 1 Then
                   .Cell(flexcpForeColor, lngRow, .COLS - 1, lngRow, .COLS - 2) = IIf(mrsBalance!�������� = 9, vbRed, vbBlue)
                   .Cell(flexcpForeColor, 1, .COLS - 1, 1, .COLS - 2) = vbRed
                   .Cell(flexcpFontBold, 1, .COLS - 1, 1, .COLS - 2) = True    '����
                
                End If
            End If
             mrsBalance.MoveNext
            .Redraw = flexRDBuffered
         Loop
         '77210,Ƚ����,2014-8-27,�����˷Ѻ����˷�,����ʾ���Ϊ��Ľ��㷽ʽ��Ϣ
         i = 1
         Do While i < .COLS - 1
            If .TextMatrix(lngRow, i + 1) = "0" Then
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
         vsBalance.AutoSizeMode = flexAutoSizeColWidth
         Call vsBalance.AutoSize(0, .COLS - 1)
         If mbytMode = EM_MULTI_�鿴 Or mbytMode = EM_MULTI_�˷����� Or mbytMode = EM_MULTI_ȡ������ _
            Or mbytMode = EM_MULTI_�˷���� Or mbytMode = EM_MULTI_ȡ����� Then
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
'        txt�˿���.Tag = dblѡ��ϼ�
    End With
End Sub

Private Sub ReCalcDelMoney()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���¼����˿�ϼ�
    '����:���˺�
    '����:2014-07-03 17:24:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnAll As Boolean, blnShowDel As Boolean
    Dim i As Long, strTemp As String
    Dim blnSeled As Boolean
    Dim blnȫ�� As Boolean
    Dim strSelNos As String, varSelNos As Variant
    Dim strFilter As String, dblMoneyNo As Double, strBalances As String
    
    txt�˿�ϼ� = Format(GetDelMoney, gstrDec)
    blnȫ�� = IsFeeAllDel
    
    blnShowDel = blnȫ�� Or mCurBillType.intInsure <> 0 Or mCurBillType.blnExistThreeAllDel
    blnShowDel = blnShowDel And Not (mbytMode = EM_MULTI_�˷����� Or mbytMode = EM_MULTI_ȡ������ _
                Or mbytMode = EM_MULTI_�˷���� Or mbytMode = EM_MULTI_ȡ�����) '�˷�����ģʽ����ʾ�˷���
    
    blnSeled = False
    With vsBill
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("���ݺ�")) <> "" And Abs(Val(.TextMatrix(i, .ColIndex("ѡ��")))) = 1 Then
                blnSeled = True
                If InStr(strSelNos, "," & .TextMatrix(i, .ColIndex("���ݺ�"))) = 0 Then
                    strSelNos = strSelNos & "," & .TextMatrix(i, .ColIndex("���ݺ�"))
                End If
            End If
        Next
        If strSelNos <> "" Then strSelNos = Mid(strSelNos, 2)
    End With
    blnShowDel = blnShowDel And blnSeled
    
    vsBalance.RowHidden(1) = Not blnShowDel
    If vsBalance.RowHidden(1) Then
        If vsBalance.COLS > 1 Then
            vsBalance.Cell(flexcpData, 1, 1, 1, vsBalance.COLS - 1) = ""
            vsBalance.Cell(flexcpText, 1, 1, 1, vsBalance.COLS - 1) = ""
        End If
        Call ControlResize
        Exit Sub
    End If

    With vsBalance
        If blnȫ�� And MCPAR.����������� Then
            For i = 1 To .COLS - 1
                .TextMatrix(1, i) = .TextMatrix(0, i)
                .Cell(flexcpData, 1, i) = .Cell(flexcpData, 0, i)
            Next
            Call ControlResize
            Exit Sub
        End If
        '����ҽ����
        '�ֶ�:���� ,����ID, ��¼����, ���㷽ʽ, ժҪ, �����ID, ���������, ���ƿ�, ���㿨���, �������, ����, ������ˮ��, ����˵��, �������, У�Ա�־, ҽ��, ���ѿ�id
         '            �Ƿ�����,�Ƿ�ȫ��,�Ƿ�����,��Ԥ��
         '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
         mrsBalance.Filter = 0
         mrsBalance.Sort = "����,���㷽ʽ"
         If vsBalance.COLS > 1 Then
            .Cell(flexcpText, 1, 1, 1, .COLS - 1) = ""
            .Cell(flexcpData, 1, 1, 1, .COLS - 1) = ""
         End If
         With vsBalance
             .Redraw = flexRDNone
             
             If mrsBalance.RecordCount <> 0 Then mrsBalance.MoveFirst
             Do While Not mrsBalance.EOF
                strTemp = ""
                Select Case Val(Nvl(mrsBalance!����))
                Case 2 'ҽ��
                    If MCPAR.����������� Then
                        strTemp = mrsBalance!���㷽ʽ
                        
                        If mblnDelByNo Then
                            'ÿһ�ֽ���ֻ����һ��
                            If InStr(strBalances, "," & strTemp) = 0 Then
                                '��ѡ�񵥾ݼ���ҽ�����㷽ʽ�˷ѽ��
                                dblMoneyNo = 0: strFilter = ""
                                varSelNos = Split(strSelNos, ",")
                                For i = 0 To UBound(varSelNos)
                                    If UBound(varSelNos) = 0 Then 'ֻ��һ�ŵ���
                                        strFilter = " or No='" & varSelNos(i) & "' and ���㷽ʽ='" & strTemp & "'"
                                    Else    '���ŵ���
                                        strFilter = strFilter & " or (No='" & varSelNos(i) & "' and ���㷽ʽ='" & strTemp & "')"
                                    End If
                                Next
                                If strFilter <> "" Then strFilter = Mid(strFilter, 4)
                                mrsInsureBalance.Filter = strFilter
                                Do While Not mrsInsureBalance.EOF
                                    dblMoneyNo = dblMoneyNo + Val(Nvl(mrsInsureBalance!���))
                                    mrsInsureBalance.MoveNext
                                Loop
                                For i = 1 To .COLS - 1 Step 2
                                    If .Cell(flexcpData, 0, i) = strTemp Then
                                        .TextMatrix(1, i) = .TextMatrix(0, i)
                                        .Cell(flexcpData, 1, i) = strTemp
                                        .TextMatrix(1, i + 1) = zlFormatNum(Val(.TextMatrix(1, i + 1)) + dblMoneyNo)
                                        Exit For
                                    End If
                                Next
                                strBalances = strBalances & "," & strTemp
                            End If
                            strTemp = "" '��գ�������治�ټ���
                        End If
                    End If
                Case 4 'һ��ͨ(��)
                   strTemp = mrsBalance!���㷽ʽ
                End Select
                If strTemp <> "" Then
                    For i = 1 To .COLS - 1 Step 2
                        If .Cell(flexcpData, 0, i) = strTemp Then
                            .TextMatrix(1, i) = .TextMatrix(0, i)
                            .Cell(flexcpData, 1, i) = strTemp
                            .TextMatrix(1, i + 1) = zlFormatNum(Val(.TextMatrix(1, i + 1)) + Val(Nvl(mrsBalance!��Ԥ��)))
                            Exit For
                        End If
                    Next
                End If
                mrsBalance.MoveNext
            Loop
            If vsBalance.COLS > 1 Then .AutoSize 1, .COLS - 1
            
            .Redraw = flexRDBuffered
         End With
        
    End With
    Call ControlResize
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
            If Val(vsBill.TextMatrix(i, .ColIndex("ѡ��"))) <> 0 Or mbytMode = EM_MULTI_�쳣���� Or mbytMode = EM_MULTI_�鿴 Then
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
    If Me.ActiveControl Is txtPatient And txtPatient.Visible Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
        IDKind.SetAutoReadCard (txtPatient.Text = "" And txtPatient.Visible)
    Else
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(False)
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(False)
        IDKind.SetAutoReadCard (False)
    End If
End Sub
Private Sub txtPatient_GotFocus()
    '����:50885
    If txtPatient.Locked Or Not txtPatient.Visible Then Exit Sub
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (txtPatient.Text = "")
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (txtPatient.Text = "")
    IDKind.SetAutoReadCard (txtPatient.Text = "" And txtPatient.Visible)
    zlControl.TxtSelAll txtPatient
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
    '77198,Ƚ����,2014-8-27,�ڲ����˷�ʱ���������,�������շѵ�ѡ��ֻ����ȡ��Ч�Һ��������շѵ������˷�
    strSQL = "" & _
        "  With �շѵ� as ( " & _
        "           Select Max(a.ID) as ID,max(M.�������) as ����ID ,max(A.����ID) as ����ID,a.No as ���ݺ�,  B.���� as ��������, a.������, a.����Ա���, a.����Ա����, a.ʵ��Ʊ��, To_Char(a.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, " & vbCrLf & _
        "                   To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ�� " & vbCrLf & _
        "           From ������ü�¼ A,���ű� B,����Ԥ����¼ M" & vbCrLf & _
        "           Where a.��¼���� = 1 And nvl(A.���ӱ�־,0)<>9 and A.��������ID=B.ID(+) And a.��¼״̬ in (1,3) " & vbCrLf & _
        "                and A.����ID=M.����ID And Nvl(a.ִ��״̬, 0) <> 1 And Nvl(a.����״̬, 0) <> 1 And a.����id = [1] " & vbCrLf & _
        "                And a.�Ǽ�ʱ�� Between Sysdate - " & gTy_System_Para.Sy_Reg.bytNODaysGeneral & " And Sysdate " & vbCrLf & _
        "          Group by   a.No,  a.������, B.����,a.����Ա���, a.����Ա����, a.ʵ��Ʊ��, To_Char(a.����ʱ��, 'yyyy-mm-dd hh24:mi:ss'),To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') " & vbCrLf & _
        "           )"

     strSQL = strSQL & vbCrLf & _
     "  Select J.*  " & vbCrLf & _
     "  From �շѵ� J," & vbCrLf & _
     "           (Select A.NO,sum(nvl(A.����,1)*nvl(A.����,1)) ����" & vbCrLf & _
     "             From ������ü�¼ A,�շѵ� B  " & vbCrLf & _
     "             Where A.NO=B.���ݺ� And mod(A.��¼����,10)=1 And a.�۸񸸺� is null  " & vbCrLf & _
     "             Group by A.NO " & vbCrLf & _
     "              Having sum(nvl(A.����,1)*nvl(A.����,1))>0 ) M" & vbCrLf & _
     "  Where J.���ݺ�=M.NO " & vbCrLf

     strSQL = "Select * From (" & strSQL & ") Order by �Ǽ�ʱ�� desc,���ݺ�"
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "�˷ѵ���", 1, "", "��ѡ����Ҫ�˷ѵĵ���", False, False, False, 0, 0, 0, blnCancel, False, False, lng����ID, "bytSize=1")
    If blnCancel = True Then Exit Function
    If rsTemp Is Nothing Then Exit Function
    If rsTemp.State <> 1 Then Exit Function

    Dim strNo As String
    strNo = Nvl(rsTemp!���ݺ�)
    mblnOneCard = GetOneCard.RecordCount > 0
    mstrNo = strNo
    
    If Val(Nvl(rsTemp!����ID)) >= 0 Then
        'bytMode-0-���ŵ��ݲ鿴,1-���ŵ����˷�,2-���쳣���˷ѵ����������˷�
        frmMultiBills.ShowMe Me, 1, mstrPrivs, strNo, "", False, mlng����ID, mblnOneCard, False, True
        Call ClearFace: Exit Function
    End If
    
    If Not ReadBills(mstrNo, True) Then
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
        If .Rows = 1 And .Cell(flexcpLeft, 0, .COLS - 1) + .ColWidth(.COLS - 1) <= .Width Then
            .Height = .RowHeight(0) + 90
            Exit Sub
        End If
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

Private Sub LoadInvoiceData(ByVal strNos As String, Optional ByVal strInvoiceNO As String)

   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ط�Ʊ��Ϣ
    '���:strNos-���ݺ�,����ö��ŷָ�
    '       strInvoiceNo-��Ʊ��(��ָ���ķ�Ʊ�ŷ�Ʊ�Ų���)
    '����:���˺�
    '����:2013-05-07 17:07:38
    '����:25187
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str��� As String, varTemp As Variant, strSQL As String
    Dim i As Long, str��Ʊ�� As String
    If mbytMode <> EM_MULTI_�˷� Then Exit Sub
    
    If gTy_Module_Para.bytƱ�ݷ������ = 0 And gTy_Module_Para.bln�ֱ��ӡ = True Then
        strSQL = "Select b.No, a.���� As Ʊ��, Null As ���, Null As ����Ʊ�����" & vbNewLine & _
                " From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ���� B" & vbNewLine & _
                " Where b.�������� = 1 And b.No In (Select Column_Value From Table(f_Str2list([1])))" & vbNewLine & _
                "       And b.Id = a.��ӡid And a.Ʊ�� = 1 And a.���� = 1 And a.ԭ��<>6" & vbNewLine & _
                "       And Not Exists (Select 1 From Ʊ��ʹ����ϸ Where ��ӡid = b.Id And ���� = 2)"

        Set mrsDelInvoice = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNos)
        GoTo LoadIntoVS
    End If
    If gTy_Module_Para.bytƱ�ݷ������ = 0 Then Exit Sub
    If mrsDelInvoice Is Nothing Then
        Set mrsDelInvoice = zlGetFromNoTOInvoice(strNos)
    End If
LoadIntoVS:
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

Private Function GetInvoiceNo(ByVal strNo As String) As String
    Dim str��Ʊ�� As String
    On Error GoTo errHandle
    If mrsDelInvoice Is Nothing Then Exit Function
    If mrsDelInvoice.State <> 1 Then Exit Function
    If mrsDelInvoice.RecordCount = 0 Then Exit Function
    With mrsDelInvoice
        .Filter = "NO='" & strNo & "'"
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            str��Ʊ�� = str��Ʊ�� & "," & Nvl(!Ʊ��)
            .MoveNext
        Loop
        .Filter = 0: .MoveFirst
    End With
    '����������
    If str��Ʊ�� = "" Then Exit Function
    str��Ʊ�� = Mid(str��Ʊ��, 2)
    GetInvoiceNo = zlStringSort(str��Ʊ��)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub FromNoSelectInvoice()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݵ���ѡ��Ʊ
    '����:���˺�
    '����:2013-05-08 15:52:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str��Ʊ�� As String, strNo As String
    Dim strNos As String, i As Long, j As Long
    If mbytMode <> 1 Or (gTy_Module_Para.bytƱ�ݷ������ = 0 And gTy_Module_Para.bln�ֱ��ӡ = False) Then Exit Sub

    On Error GoTo errHandle
    If gTy_Module_Para.bytƱ�ݷ������ = 0 And gTy_Module_Para.bln�ֱ��ӡ = True Then
        With vsBill
            str��Ʊ�� = ""
            For i = 1 To .Rows - 1
                  If Abs(Val(.TextMatrix(i, .ColIndex("ѡ��")))) = 1 Then
                        strNo = .TextMatrix(i, .ColIndex("���ݺ�"))
                        If strNo <> "" Then
                            str��Ʊ�� = str��Ʊ�� & "," & GetInvoiceNo(strNo)
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
    Else
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
    End If
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
    If mbytMode <> 1 Or (gTy_Module_Para.bytƱ�ݷ������ = 0 And gTy_Module_Para.bln�ֱ��ӡ = False) Then Exit Sub
    If mrsDelInvoice Is Nothing Then Exit Sub
    If mrsDelInvoice.RecordCount = 0 Then Exit Sub
    On Error GoTo errHandle
    mrsDelInvoice.Filter = "Ʊ��='" & strInvoiceNO & "'"
    If gTy_Module_Para.bytƱ�ݷ������ = 0 And gTy_Module_Para.bln�ֱ��ӡ = True Then
        With mrsDelInvoice
            If .RecordCount <> 0 Then .MoveFirst
            Do While Not .EOF
                strNo = Nvl(!NO)
                  With vsBill
                      k = vsBill.FindRow(strNo, , .ColIndex("���ݺ�"))
                      For j = k To .Rows - 1
                          If .TextMatrix(j, .ColIndex("���ݺ�")) <> strNo Then Exit For
                          .Cell(flexcpChecked, j, .ColIndex("ѡ��")) = 1
                          'ͬ��ѡ����������Ŀ
                          Call SynchronizationSelect(j)
                      Next
                  End With
                 .MoveNext
            Loop
        End With
    Else
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
    End If
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
    If mbytMode <> 1 Or (gTy_Module_Para.bytƱ�ݷ������ = 0 And gTy_Module_Para.bln�ֱ��ӡ = False) Then Exit Sub

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
Private Function ExistsBalance(ByVal str���㷽ʽ As String, ByRef intCol As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ָ���Ľ��㷽ʽ
    '���:
    '����:intCol-ָ�����㷽ʽ��(-1��ʾδ�ҵ�)
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-03-20 13:40:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer

    On Error GoTo errHandle
    intCol = -1
    With vsBalance
        For i = 1 To .COLS - 1 Step 2
            If .Cell(flexcpData, 1, i) = str���㷽ʽ Then
                intCol = i
                ExistsBalance = True: Exit Function
            End If
        Next
    End With
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function CheckDiff(strNos As String, strDiffNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Ƚϸ�ֵ�ĵ��ݺ��Ƿ�һ��
    '���:
    '����:
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
    MCPAR.ҽ������Ʊ�� = False
    MCPAR.������ȫ�� = gclsInsure.GetCapability(support������ȫ��, lng����ID, intInsure) '86176
    MCPAR.�൥�ݷֵ��ݽ��� = gclsInsure.GetCapability(support�൥�ݷֵ��ݽ���, lng����ID, intInsure)
    MCPAR.һ�ν���ֵ����˷� = gclsInsure.GetCapability(supportһ�ν���ֵ����˷�, lng����ID, intInsure)
End Sub

Private Sub SetFunCtrlVisible()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ù��ܿؼ���visible����
    '����:���˺�
    '����:2014-07-03 16:41:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnVisible As Boolean
    
    blnVisible = mbytMode = EM_MULTI_�˷� Or mbytMode = EM_MULTI_�˷����� Or mbytMode = EM_MULTI_ȡ������ _
                Or mbytMode = EM_MULTI_�˷���� Or mbytMode = EM_MULTI_ȡ�����
    cmdSelAll.Visible = blnVisible
    cmdClear.Visible = blnVisible
    cmdBillSel.Visible = mbytMode = EM_MULTI_�˷�
    cmdRefuseApply.Visible = mbytMode = EM_MULTI_�˷����
    cmdOK.Visible = Not mbytMode = EM_MULTI_�鿴
    If mlng������� <> 0 Then   '���洫��ʱ,�����ֹ�����
        txtNO.Visible = False
        optNO(0).Visible = False
        optNO(1).Visible = False
        picPatiBack.Visible = False
        fraInfo_1.Visible = False
    End If
End Sub
Private Function GetYBTOCash(ByVal lng����ID As Long, ByVal intInsure As Integer) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡҽ��ʹ���ֽ�֧���Ľ��㷽ʽ(����ö��ŷָ�)
    '����:���ؽ��㷽ʽ,����ö��ŷָ�:�����ʻ�,ҽ������...
    '����:���˺�
    '����:2014-07-07 09:57:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str���㷽ʽ As String
    
    On Error GoTo errHandle
    If intInsure = 0 Then Exit Function
    
    '�ֶ�:����,��¼����,���㷽ʽ,ժҪ,�����ID,���������,���ƿ�,���㿨���,�������,����,������ˮ��, ����˵��,�������,У�Ա�־,ҽ��,���ѿ�id
    '     �Ƿ�����,�Ƿ�ȫ��,�Ƿ�����,��Ԥ��
    '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�)
    mrsBalance.Filter = "����=2"
    If mrsBalance.RecordCount = 0 Then Exit Function
    With mrsBalance
        Do While Not .EOF
            '������ֽ��㷽ʽ��֧�ֻ���,Ҫ��Ϊ�ֽ�,���ü�ȥ
            If MCPAR.����������� Then
                If Not gclsInsure.GetCapability(support�����������, lng����ID, intInsure, !���㷽ʽ) Then
                    str���㷽ʽ = str���㷽ʽ & "," & !���㷽ʽ
                End If
            Else     '��֧�������������ʱ,ֻ���������Ϊ�ֽ�,����ԭ����,������ҽ������
                If !���㷽ʽ = mstr�����ʻ� Then
                    str���㷽ʽ = str���㷽ʽ & "," & !���㷽ʽ
                End If
            End If
            .MoveNext
        Loop
    End With
    If str���㷽ʽ <> "" Then str���㷽ʽ = Mid(str���㷽ʽ, 2)
    GetYBTOCash = str���㷽ʽ
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
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
    ByVal lng����ID As Long, ByVal lng������� As Long, ByVal str���ս�� As String, _
    ByVal strNo As String, ByVal lng����ID As Long, ByVal strInvoice As String, ByVal dtDelDate As Date) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ��ҽ�������շ�
    '����:ִ�гɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-31 23:38:11
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
        varTemp = Split(str���ս��, ";") 'curʵ�պϼ�;cur����ͳ��;curȫ�Ը�;cur���Ը�
        If intInsure <> 0 And mstr�����ʻ� <> "" And mdbl������� > -1 * mdbl����͸֧ Then
            If RoundEx(Val(varTemp(0)), 6) >= 0 Then
                cur���� = RoundEx(Val(varTemp(1)), 6) + IIf(MCPAR.���Ը�, RoundEx(Val(varTemp(3)), 6), 0) + IIf(MCPAR.ȫ�Ը�, RoundEx(Val(varTemp(2)), 6), 0)
                If mdbl������� - cur���� >= -1 * mdbl����͸֧ Then
                    strBalance = mstr�����ʻ� & "|" & cur����   '������͸֧��Χ���㹻(����͸֧0Ϊ����)
                Else
                    If mdbl����͸֧ = 0 And mdbl������� > 0 Then
                        strBalance = mstr�����ʻ� & "|" & mdbl�������  '������͸֧�������
                    Else
                        '��������͸֧��Χ������͸֧ʱ�����
                        If mdbl����͸֧ <> 0 Then
                            strBalance = mstr�����ʻ� & "|" & mdbl������� + mdbl����͸֧ '������͸֧��Χ��֧��
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
    ' Zl_�����շѽ���_Modify
    strSQL = "Zl_�����շѽ���_Modify("
    '  ��������_In   Number,
    '  --��������_In:
    '  --   0-��ͨ�շѷ�ʽ:
    '  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ||.." ;Ҳ�������.
    '  --     �ڳ�Ԥ��_In:����漰Ԥ����,���뱾�εĳ�Ԥ��,�������շ�ʱ,������
    '  --     ����֧Ʊ��_In:����漰��֧Ʊ,���뱾�ε���֧Ʊ��,�������շ�ʱ,������
    '  --   1.����������:
    '  --     �ٽ��㷽ʽ_IN:ֻ�ܴ���һ�����㷽ʽ,���������һЩ������Ϣ,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ"
    '  --     �ڳ�Ԥ��_In: ������
    '  --     ����֧Ʊ��_In:������
    '  --     �ܿ����ID_IN,����_IN,������ˮ��_IN,����˵��_In:��Ҫ����
    '  --   2-ҽ������(�������ҽ���Ľ���,��Ҫ��ɾ��ԭҽ������,���´���ĸ���)
    '  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:���㷽ʽ|������||.."
    '  --     �ڳ�Ԥ��_In: ������
    '  --     ����֧Ʊ��_In:������
    '  --   3-���ѿ�����:
    '  --     �ٽ��㷽ʽ_IN:����һ��ˢ���ſ�,��ʽΪ:�����ID|����|���ѿ�ID|���ѽ��||."
    '  --     �ڳ�Ԥ��_In: ������
    '  --     ����֧Ʊ��_In:������
    strSQL = strSQL & "" & 2 & ","
    '  ����id_In     ������ü�¼.����id%Type,
    strSQL = strSQL & "" & lng����ID & ","
    '  ����id_In     ����Ԥ����¼.����id%Type,
    strSQL = strSQL & "" & lng����ID & ","
    '  ���㷽ʽ_In   Varchar2,
    strSQL = strSQL & "'" & strBalance & "')"
    '  ��Ԥ��_In     ����Ԥ����¼.��Ԥ��%Type := Null,
    '  ��֧Ʊ��_In   ����Ԥ����¼.��Ԥ��%Type := Null,
    '  �����id_In   ����Ԥ����¼.�����id%Type := Null,
    '  ����_In       ����Ԥ����¼.����%Type := Null,
    '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
    '  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
    '  �ɿ�_In       ����Ԥ����¼.�ɿ�%Type := Null,
    '  �Ҳ�_In       ����Ԥ����¼.�Ҳ�%Type := Null,
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
    If ExecuteClinicSwap(lng����ID, intInsure, lng����ID, lng�������, strBalance, strNos, str���ս��) = False Then Exit Function
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
    Set rsTemp = zlMakeClinicPreSwapData(strInvoice, lng����ID, strNos)
RePreSwap:
    strAdvance = "1": strBalance = ""
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
    ByVal lng������� As Long, ByVal strԤ���� As String, ByVal strNos As String, Optional ByVal str���ս�� As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ҽ������ӿ�
    '���:  lng����ID:���ν��ʵ�ID
    '����:
    '����:ҽ�����óɹ����ҽ��,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-20 17:15:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varNos As Variant
    Dim strBillNO As String, blnTrans As Boolean, blnTransMedicare As Boolean
    Dim p As Integer, strAdvance As String
    Dim strTmp As String, i As Long, strSQL As String
    Dim cur����֧�� As Currency, curҽ������ As Currency, varTemp As Variant
    
    On Error GoTo errHandle
     
    blnTrans = True
    '��������ҽ��,Ҫ����Ϊ���۵�,������δ����(���˽�ҵ��)
''    '1. ����Ϊ���۵�
''    If mblnSavePrice Then
''        '����Ϊ���۵�
''        '���������ҽ��,�շ�ȷ��ʱʵ��ȴ����Ϊ���۵�:�����۵���ϸ,����Oracle������ִ��
''        varNos = Split(mobjChargeInfor.Nos, ",")
''        For p = 1 To UBound(varNos)
''            strBillNO = mobjChargeInfor(p)
''            If Not gclsInsure.TranChargeDetail(1, strBillNO, 1, 0, "", , mobjChargeInfor.intInsure) Then
''                'ɾ�����۵�(��������)
''                Call DelMedicareTempNO(True, strBillNO)
''                gcnOracle.RollbackTrans: Exit Function
''            End If
''        Next
''        mblnYbBalanced = True   'ҽ���Ѿ�����
''        ExecuteClinicSwap = True
''        Exit Function
''    End If
      
    If MCPAR.ҽ���ӿڴ�ӡƱ�� And MCPAR.ҽ������Ʊ�� = False Then
        '���ϸ����Ʊ��ʱ���浱ǰƱ��
        If Not gblnStrictCtrl Then
            zlDatabase.SetPara "��ǰ�շ�Ʊ�ݺ�", mCurBillType.strInvoice, glngSys, 1121, zlStr.IsHavePrivs(mstrPrivs, "��������")
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
    varTemp = Split(str���ս��, ";") 'curʵ�պϼ�;cur����ͳ��;curȫ�Ը�;cur���Ը�
    
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
        '�޸�У�Ա�־
        ' Zl_���������շ�_ҽ������
        strSQL = "Zl_���������շ�_ҽ������("
        '  ����id_In   ������ü�¼.����id%Type,
        strSQL = strSQL & lng����ID & ","
        '  �������_In ����Ԥ����¼.�������%Type,
        strSQL = strSQL & "Null,"
        '  ���ս���_In Varchar2
        strSQL = strSQL & "Null)"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
       gcnOracle.CommitTrans
       Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicSwap, True, intInsure)
       ExecuteClinicSwap = True: Exit Function
    End If
    
    '����������������
    '��Ҫ������������
    'Zl_�����շѽ���_Modify
    strSQL = "Zl_�����շѽ���_Modify("
    '  ��������_In   Number,
    strSQL = strSQL & "" & 2 & ","
    '  ����id_In     ������ü�¼.����id%Type,
    strSQL = strSQL & "" & lng����ID & ","
    '  ����id_In     ����Ԥ����¼.����id%Type,
    strSQL = strSQL & "" & lng����ID & ","
    '  ���㷽ʽ_In   Varchar2,
    strSQL = strSQL & "'" & strAdvance & "')"
    '  ��Ԥ��_In     ����Ԥ����¼.��Ԥ��%Type,
    '  ��֧Ʊ��_In   ����Ԥ����¼.��Ԥ��%Type,
    '  �����id_In   ����Ԥ����¼.�����id%Type := Null,
    '  ����_In       ����Ԥ����¼.����%Type := Null,
    '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
    '  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
    '  �ɿ�_In       ����Ԥ����¼.�ɿ�%Type := Null,
    '  �Ҳ�_In       ����Ԥ����¼.�Ҳ�%Type := Null,
    '  �����_In   ������ü�¼.ʵ�ս��%Type := Null,
    '  ��ɽ���_In Number:=0
    ') As
    '  ------------------------------------------------------------------------------------------------------------------------------
    '  --����:�շѽ���ʱ,�޸Ľ���������Ϣ
    '  --��������_In:
    '  --   0-��ͨ�շѷ�ʽ:
    '  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ||.." ;Ҳ�������.
    '  --     �ڳ�Ԥ��_In:����漰Ԥ����,���뱾�εĳ�Ԥ��,�������շ�ʱ,������
    '  --     ����֧Ʊ��_In:����漰��֧Ʊ,���뱾�ε���֧Ʊ��,�������շ�ʱ,������
    '  --   1.����������:
    '  --     �ٽ��㷽ʽ_IN:ֻ�ܴ���һ�����㷽ʽ,���������һЩ������Ϣ,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ"
    '  --     �ڳ�Ԥ��_In: ������
    '  --     ����֧Ʊ��_In:������
    '  --     �ܿ����ID_IN,����_IN,������ˮ��_IN,����˵��_In:��Ҫ����
    '  --   2-ҽ������(�������ҽ���Ľ���,��Ҫ��ɾ��ԭҽ������,���´���ĸ���)
    '  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:���㷽ʽ|������||.."
    '  --     �ڳ�Ԥ��_In: ������
    '  --     ����֧Ʊ��_In:������
    '  --   3-���ѿ�����:
    '  --     �ٽ��㷽ʽ_IN:����һ��ˢ���ſ�,��ʽΪ:�����ID|����|���ѿ�ID|���ѽ��||."
    '  --     �ڳ�Ԥ��_In: ������
    '  --     ����֧Ʊ��_In:������
    '  -- �����_In:��������ʱ,����
    '  -- ��ɽ���_In:1-����շ�;0-δ����շ�
    '  ------------------------------------------------------------------------------------------------------------------------------
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    '�޸�У�Ա�־
    ' Zl_���������շ�_ҽ������
    strSQL = "Zl_���������շ�_ҽ������("
    '  ����id_In   ������ü�¼.����id%Type,
    strSQL = strSQL & lng����ID & ","
    '  �������_In ����Ԥ����¼.�������%Type,
    strSQL = strSQL & "Null,"
    '  ���ս���_In Varchar2
    strSQL = strSQL & "Null)"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    gcnOracle.CommitTrans: blnTrans = False
    Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicSwap, True, intInsure)
    ExecuteClinicSwap = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
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
    ExecuteYBIdentifyCancel = True
    If mbytMode = EM_MULTI_�鿴 Then Exit Function
    If lng����ID = 0 Then Exit Function
    On Error GoTo errHandle
    ExecuteYBIdentifyCancel = gclsInsure.IdentifyCancel(0, lng����ID, intInsure)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
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
    Dim strNo As String, lng��ӡID As Long
    
    On Error GoTo errHandle
    
    If blnSaveOK = False Then Exit Sub
    
    If objDelBalance.blnOnePatiPrint Then
        strPrintNos = "'" & Replace(objDelBalance.strOnePatiPrintNos, ",", "';'") & "'"
    Else
        strPrintNos = objDelBalance.PrintNOs
    End If
 
    If Mid(objDelBalance.CurDelNos, 1, 1) = "," Then
        strNo = Split(objDelBalance.CurDelNos, ",")(1)
    Else
        strNo = Split(objDelBalance.CurDelNos, ",")(0)
    End If
    
    strReclaimInvoice = zlGetReclaimInvoice(strPrintNos)
    If gTy_Module_Para.bytƱ�ݷ������ <> 0 And strReclaimInvoice <> "" Then
        If InStr(1, mstrPrivs, "�˷Ѻ��շ�Ʊ") > 0 Then
            If MsgBox("ע��:" & vbCrLf & " ��ǰ�˷ѵĵ����а��������շ�Ʊ�ݣ��Ƿ������ЩƱ��?" & vbCrLf & strReclaimInvoice, _
                vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then GoTo Completed '����ӡ�˷������˳�
        End If
    End If
    
    If gblnBillPrint Then
        If objDelBalance.blnOnePatiPrint Then
            If gobjBillPrint.zlEraseBill(objDelBalance.strOnePatiPrintNos, 0) = False Then Exit Sub
        Else
            If gobjBillPrint.zlEraseBill(mCurBillType.strAllNOs, 0) = False Then Exit Sub
        End If

    End If
    
   '��ӡ�˷ѵ���
    Call PrintDelBill(objDelBalance.AllNos, objDelBalance.CurDelNos, strNo, objDelBalance.����ID, _
        objDelBalance.�˷�ʱ��, objDelBalance.�����˷�, objDelBalance.���շ�Ʊ, strReclaimInvoice, objDelBalance.blnOnePatiPrint, objDelBalance.strOnePatiPrintNos)


Completed:
    mblnOK = True: Call ClearFace
    
    If txtNO.Visible Then
        txtNO.SetFocus: Exit Sub
    End If
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
    
    '78004,Ƚ����,2014-9-16,�ڲ�����ҩƷ��λ��Ϊ�����ﵥλ��ʱ����
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
    "      sum(A.ԭʼ����/" & IIf(gblnҩ����λ, "nvl(B." & gstrҩ����װ & ",1)", "1") & ") as ԭʼ����, " & _
    "      sum(A.����/" & IIf(gblnҩ����λ, "nvl(B." & gstrҩ����װ & ",1)", "1") & ")  as ʣ������ " & _
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

Private Function CheckIsAllDel(ByVal strAllNOs As String, _
    Optional ByVal blnBillSaved As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������з����Ƿ�ȫ��
    '���:strAllNos-���е���,����ö��ŷָ�
    '����:
    '����:����ȫ��ʱ,����true,���򷵻�False
    '����:���˺�
    '����:2014-07-15 11:28:38
    
    '�޸ģ�104573
    '��Σ�
    '   blnBillSaved - ���������Ƿ��ѱ��棬�ѱ����ֻҪ������������ʣ�������ͱ�ʾδ���꣬��Ҫ��Ʊ�ݴ�ӡ���쳣���˵���
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strNo As String, int��� As Integer
    Dim blnFind As Boolean, dblʣ������ As Double
    Dim j As Long, k As Long
    
    On Error GoTo errHandle
    If mbytMode = EM_MULTI_�˷� Then
        With vsBill
            For j = 1 To vsBill.Rows - 1
                If Abs(Val(.TextMatrix(j, .ColIndex("ѡ��")))) <> 1 And InStr(strAllNOs, .TextMatrix(j, .ColIndex("���ݺ�"))) > 0 Then
                   CheckIsAllDel = False: Exit Function
                End If
            Next
        End With
    End If
    
    Set rsTemp = GetFeeDelNumRecord(strAllNOs)
    If blnBillSaved = False Then
        Do While Not rsTemp.EOF
            strNo = Nvl(rsTemp!NO): int��� = Val(Nvl(rsTemp!���))
            dblʣ������ = Val(Nvl(rsTemp!ʣ������))
            If dblʣ������ <> 0 Then
                With vsBill
                    k = vsBill.FindRow(strNo, , .ColIndex("���ݺ�"))
                    If k <= 0 Then Exit Function
                    blnFind = False
                    For j = k To vsBill.Rows - 1
                        If .TextMatrix(j, .ColIndex("���ݺ�")) <> strNo Then Exit For
                        If Abs(Val(.TextMatrix(j, .ColIndex("ѡ��")))) <> 1 _
                            And mbytMode <> EM_MULTI_�쳣���� Then
                            CheckIsAllDel = False: Exit Function
                        End If
                        If Val(.RowData(j)) = int��� Then
                            If dblʣ������ <> Val(.Cell(flexcpData, j, .ColIndex("����"))) Then
                               CheckIsAllDel = False: Exit Function
                            End If
                            blnFind = True: Exit For
                        End If
                    Next
                End With
                If blnFind = False Then Exit Function
            End If
            rsTemp.MoveNext
        Loop
    Else
        If rsTemp.RecordCount = 0 Then
            CheckIsAllDel = True: Exit Function
        End If
        Do While Not rsTemp.EOF
            If RoundEx(Val(Nvl(rsTemp!ʣ������)), 6) <> 0 Then
                'Ʊ�ݴ�ӡʱ���˷������ѱ��棬��ʱֻҪ��ʣ��������������ͱ�ʾû����
                CheckIsAllDel = False: Exit Function
            End If
            rsTemp.MoveNext
         Loop
    End If
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
    Dim frmBalance As frmClinicDelBalance, objDelBalance As clsCliniDelBalance
    Dim blnȫ�� As Boolean, str����ID As String, lng����ID As Long, lng����ID As Long
    Dim strNos As String, varData As Variant, strCmdCaptions As String
    Dim cllPro  As New Collection, strInvoices As String, strInvoice As String
    Dim lngCheck����ID As Long, intCheckInsure   As Integer, strYBPati As String
    Dim dtDelDate As Date, blnTrans As Boolean, strNo As String
    Dim str��� As String, j As Long, strPrintNOInfor As String
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    Dim cur����͸֧ As Currency, str���ս�� As String 'curʵ�պϼ�;cur����ͳ��;curȫ�Ը�;cur���Ը�
    Dim strReturn As String, strReturnRecipt As String '�˷Ѵ�����Ϣ����ʽ��NO,ҩ��ID|NO,ҩ��ID|��
    Dim rsҩƷ��¼ As ADODB.Recordset, lng����ID As Long
    Dim strAllBalance As String, strAdvance As String
    
    On Error GoTo errHandle
    '�������
    If zlIsCheckExistErrBill(mlng�������) = False Then
        MsgBox "��ǰ�쳣�����ѱ������㲻�ܼ�����", vbInformation, gstrSysName
        Exit Function
    End If
    If zlCheckOtherSessionDoing(mlng�������) Then
        MsgBox "��ǰ�������������շѴ����н��д����㲻�ܼ�����", vbInformation, gstrSysName
        Exit Function
    End If
    '105432,���������㷽ʽ��Ч�Լ��
    If ThreeBalanceCheck(zlFromIDGetChargeBalance(2, mCurBillType.strAllNOs, mblnNOMoved, , True), mrs���㷽ʽ, mcllForceDelToCash) = False Then Exit Function
    
    blnȫ�� = CheckIsAllDel(mCurBillType.strAllNOs, True)
    If Not blnȫ�� Then
        If zlStr.IsHavePrivs(mstrPrivs, "�����˷�") = False Then
            MsgBox "��û��Ȩ��ִ�в����˷Ѳ�����", vbInformation, gstrSysName
            vsBill.SetFocus: Exit Function
        End If
        If zlStr.IsHavePrivs(mstrPrivs, "�˷Ѻ��շ�Ʊ") Then
            If frmReInvoice.ShowMe(Me, mstrNo, Val(txtAllTotal.Text), 0, strInvoices) = False Then
                vsBill.SetFocus: Exit Function
            End If
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
        If strNo <> "" And str��� <> "" Then
            strPrintNOInfor = strPrintNOInfor & ";" & strNo & ":" & str���
        End If
    End With
    
    Set objDelBalance = New clsCliniDelBalance
    'bytType-��������:0-���ݽ���ID����;1-���ݽ�����Ų���,2-���ݵ��ݺ�����ȡ���㷽ʽ
    Set objDelBalance.rsBalance = zlFromIDGetChargeBalance(2, mCurBillType.strAllNOs, False)
    Set objDelBalance.rs���㷽ʽ = mrs���㷽ʽ
    
    lng����ID = mCurBillType.lng����ID
    lng����ID = mCurBillType.lng����ID
    
    If mCurBillType.intInsure <> 0 And lng����ID <> 0 And MCPAR.ҽ���ӿڴ�ӡƱ�� Then
        If InputFactNo(lng����ID, strInvoice) = False Then Exit Function
    End If
    
    dtDelDate = zlDatabase.Currentdate
    
    '�����ڸ���Ϊ��ǰ����
    '�����շ�ʱ�����շѵĵǼ�ʱ����ʱ����еǼǴ���
    'Zl_�����շ��쳣_Update
    strSQL = "Zl_�����շ��쳣_Update("
    '  No_In       ������ü�¼.No%Type,
    strSQL = strSQL & "NULL,"
    '  �Ǽ�ʱ��_In ������ü�¼.�Ǽ�ʱ��%Type,
    strSQL = strSQL & "to_date('" & Format(dtDelDate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
    '  ����id_In   ������ü�¼.����id%Type := Null
    strSQL = strSQL & "" & mCurBillType.lng����ID & ")"
    zlAddArray cllPro, strSQL
    If mCurBillType.lng����ID <> 0 Then
        'Zl_�����շ��쳣_Update
        strSQL = "Zl_�����շ��쳣_Update("
        '  No_In       ������ü�¼.No%Type,
        strSQL = strSQL & "NULL,"
        '  �Ǽ�ʱ��_In ������ü�¼.�Ǽ�ʱ��%Type,
        strSQL = strSQL & "to_date('" & Format(dtDelDate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
        '  ����id_In   ������ü�¼.����id%Type := Null
        strSQL = strSQL & "" & mCurBillType.lng����ID & ")"
        zlAddArray cllPro, strSQL
    End If
    
    '�൥�ݷֵ��ݽ���ʱ��û�����ռ�¼��lng����ID�϶�����0
    '����ҽ��
    If mCurBillType.intInsure <> 0 And lng����ID <> 0 And MCPAR.����������� Then
        '�����ҽ��,�����쳣,�϶���ֻ�����ղ��ֲų����쳣
        '�ֶ�:���� ,����ID, ��¼����, ���㷽ʽ, ժҪ, �����ID, ���������, ���ƿ�, ���㿨���, �������, ����, ������ˮ��, ����˵��, �������, У�Ա�־, ҽ��, ���ѿ�id
        '            �Ƿ�����,�Ƿ�ȫ��,�Ƿ�����,��Ԥ��
        '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
        mrsBalance.Filter = "����ID=" & lng����ID & " And ����=2 "
        If mrsBalance.EOF Then
            '79237,Ƚ����,2014-11-5
            '�п����Ѿ��ɹ�������ҽ�����㣬����ҽ���������Ϊ��
            strSQL = "" & _
                "   Select 1" & _
                "   From ����Ԥ����¼ A, ���ս����¼ B" & _
                "   Where a.����id = b.��¼id And a.��¼���� = 3 And a.��¼״̬ = 1 And b.���� = 1 " & _
                "         And a.������� = [1] And Rownum < 2"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����Ƿ��ѵ���ҽ���ӿ�", mlng�������)
            If rsTemp.EOF Then
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
                blnTrans = True
                zlExecuteProcedureArrAy cllPro, Me.Caption, True
                
                '�������ռ�¼�ı�����Ϣ����Ϊ���˷�ʱ����δ���£�Ϊ�˱�������������¸���һ��
                '77951,Ƚ����,2014-9-16
                If ExecuteInsureInfoUpdate(lng����ID, str���ս��) = False Then Exit Function
                '��ȡ�������
                cur����͸֧ = mdbl����͸֧
                mdbl������� = gclsInsure.SelfBalance(mCurBillType.lng����ID, CStr(Split(strYBPati, ";")(1)), 10, cur����͸֧, mCurBillType.intInsure)
                mdbl����͸֧ = cur����͸֧
                '77058
                If ExcuteInsureReCharge(mCurBillType.lng����ID, mCurBillType.intInsure, lng����ID, mlng�������, str���ս��, _
                            strNo, lng����ID, strInvoice, dtDelDate) = False Then Exit Function
                blnTrans = False
                Set cllPro = New Collection
            End If
        End If
    ElseIf mCurBillType.intInsure <> 0 And mblnDelByNo Then
        strAllBalance = GetYBOldBalance(mCurBillType.lng����ID, mCurBillType.intInsure, mCurBillType.lng����ID)
        
        blnTrans = True
        zlExecuteProcedureArrAy cllPro, Me.Caption, True
        If ExecuteClinicDelNo(mCurBillType.lng����ID, mCurBillType.intInsure, lng����ID, mCurBillType.lngԭ����ID, strAdvance, True) = False Then
            gcnOracle.RollbackTrans
            Exit Function
        End If
        gcnOracle.CommitTrans: blnTrans = False
        
        If zlInsureCheck(strAllBalance, strAdvance) And strAdvance <> "" Then
            '�˷Ѻ��շѲ�һ��ʱ,��ҪЧ��
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
            strSQL = strSQL & "" & 3 & ","
            '  ����id_In     ������ü�¼.����id%Type,
            strSQL = strSQL & "" & mCurBillType.lng����ID & ","
            '  ����id_In     ����Ԥ����¼.����id%Type,
            strSQL = strSQL & "" & mCurBillType.lng����ID & ","
            '  ���㷽ʽ_In   Varchar2,
            strSQL = strSQL & "'" & strAdvance & "')"
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
            zlDatabase.ExecuteProcedure strSQL, Me.Caption
        End If
    
        '�޸�У�Ա�־
        ' Zl_���������շ�_ҽ������
        strSQL = "Zl_���������շ�_ҽ������("
        '  ����id_In   ������ü�¼.����id%Type,
        strSQL = strSQL & mCurBillType.lng����ID & ","
        '  �������_In ����Ԥ����¼.�������%Type,
        strSQL = strSQL & "Null,"
        '  ���ս���_In Varchar2
        strSQL = strSQL & "Null)"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    End If
    
    '2.����һ��ͨ(�ϰ汾)
    '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
    mrsBalance.Filter = "����=4 "
    objDelBalance.rsBalance.Filter = "����=4 "
    If mrsBalance.EOF And objDelBalance.rsBalance.EOF = False Then
ReDOOneCard:
        If CheckOnCardValied(objDelBalance.rsBalance) = False Then Exit Function
        
        blnTrans = True
        zlExecuteProcedureArrAy cllPro, Me.Caption, True
        If Not ExecuteOneCardDelInterface(objDelBalance.rsBalance, lng����ID) Then
            mrsBalance.Filter = 0
            If mCurBillType.intInsure <> 0 Then
                If frmVerfyCodeInput.ShowMsg(Me, "����[" & mCurBillType.strDelNOs & "]�Ѿ��˷ѳɹ�,��һ��ͨ����ʧ��,[�쳣����]����������֤��,���鲻�����쳣���ݱ���", strCmdCaptions) = False Then
                    gcnOracle.BeginTrans: blnTrans = True
                    GoTo ReDOOneCard:
                End If
            End If
            Exit Function
        End If
        blnTrans = False
        Set cllPro = New Collection
    End If
    
    '4.��ʾ�������
    mCurBillType.lng������� = mlng������� '��¼���ڴ�ӡ��Ʊ
    If strPrintNOInfor <> "" Then strPrintNOInfor = Mid(strPrintNOInfor, 2)
    With objDelBalance
        .intInsure = mCurBillType.intInsure
        .CurDelNos = mCurBillType.strDelNOs
        .AllNos = mCurBillType.strAllNOs
        .PrintNOs = strPrintNOInfor
        
        .PatiUseType = mstrUseType
        .SaveBilled = True
        
        .ShareUserID = mlngShareUseID
        .����ID = mCurBillType.lng����ID
        .����ID = lng����ID
        .��ǰ��Ʊ�� = strInvoice
        .���շ�Ʊ = strInvoices
        .������� = mlng�������
        .����ID = lng����ID
        .ȱʡ���㷽ʽ = mCurBillType.str���㷽ʽ
        .�˷Ѻϼ� = -1 * GetDelMoney
        .�ѱ� = mCurBillType.str�ѱ�
        .���� = mCurBillType.str����
        .�Ա� = mCurBillType.str�Ա�
        .���� = mCurBillType.str����
        .ҽ������Ʊ�� = MCPAR.ҽ������Ʊ��
        .ԭ����ID = mCurBillType.lngԭ����ID
        .�˷�ʱ�� = dtDelDate
        .�����˷� = Not blnȫ��
    End With
    
    Set frmBalance = New frmClinicDelBalance
    If frmBalance.zlDelCharge(Me, EM_FUN_����, mlngModule, mstrPrivs, objDelBalance, cllPro, , mcllForceDelToCash) = False Then Exit Function
    
    '81190,Ƚ����,�˷�ҵ����ҩ���ϴ��˷���Ϣ
    On Error Resume Next
    If mblnDrugMachine Then
        Dim strData As String '���ﴦ����ҩ��ʽ������ID1,��ҩ����1;����ID2,��ҩ����2;...
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
            "   Where a.����id = b.����id And a.��¼״̬=2  And a.�շ���� In ('5', '6', '7') And b.������� = [1]"
        Set rsҩƷ��¼ = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mlng�������))
        
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


Private Function CheckOnCardValied(ByVal rsBalance As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���һ��ͨ�Ƿ�Ϸ�
    '���:rsBalance-ԭʼ�Ľ�������
    '����:�Ϸ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-31 12:00:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCardNo As String
    On Error GoTo errHandle
    
    '�ֶ�:����,��¼����,���㷽ʽ,ժҪ,�����ID,���������,���ƿ�,���㿨���,�������,����,������ˮ��, ����˵��,�������,У�Ա�־,ҽ��,���ѿ�id
    '     �Ƿ�����,�Ƿ�ȫ��,�Ƿ�����,��Ԥ��
    '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�)
    mrsBalance.Filter = "����=4"
    If rsBalance.RecordCount = 0 Then CheckOnCardValied = True: Exit Function
    If mobjICCard Is Nothing Then
        On Error Resume Next
        Set mobjICCard = CreateObject("zlICCard.clsICCard")
        On Error GoTo 0
        If mobjICCard Is Nothing Then
            MsgBox "һ��ͨ�ӿڴ���ʧ��,���ܽ����˷�!����ӿ��ļ�.", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    strCardNo = mobjICCard.Read_Card(Me)
    If strCardNo <> Nvl(rsBalance!����) Then
        MsgBox "��ǰ������ۿ�Ų�һ��,���ܽ����˷�.", vbInformation, gstrSysName
        Exit Function
    End If
    CheckOnCardValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckDelAppliedValied(ByVal bytMode As gEM_ChargeDelType, ByRef strNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����˷�����ĺϷ���
    '���:
    '����:strNos-�����˷�����ĵ��ݺ�,����ö��ŷ���
    '����:�˷�����Ϸ��ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-08-05 11:20:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNo As String, j As Long, i As Long
    Dim rsTemp As ADODB.Recordset, varTemp As Variant
    
    On Error GoTo errHandle
    strNos = ""
    With vsBill
        strNo = ""
        For j = 1 To vsBill.Rows - 1
            If strNo <> Trim(.TextMatrix(j, .ColIndex("���ݺ�"))) Then
                strNo = .TextMatrix(j, .ColIndex("���ݺ�"))
                If InStr(strNos & ",", "," & strNo & ",") = 0 Then
                    If Abs(Val(.TextMatrix(j, .ColIndex("ѡ��")))) = 1 Then
                        strNos = strNos & "," & strNo
                    End If
                End If
            End If
        Next
    End With
    If strNos <> "" Then strNos = Mid(strNos, 2)
    If strNos = "" Then
        MsgBox "δѡ�񵥾ݣ���ѡ��", vbInformation + vbOKOnly, gstrSysName
        vsBill.SetFocus: Exit Function
    End If
    
    If bytMode = EM_MULTI_�ܾ����� Then
        If Trim(txt�˷�ժҪ.Text) = "" Then
            MsgBox "��������ܾ�ԭ��", vbInformation + vbOKOnly, gstrSysName
            If txt�˷�ժҪ.Visible And txt�˷�ժҪ.Enabled Then txt�˷�ժҪ.SetFocus
            Exit Function
        End If
    End If
    
    Set rsTemp = GetApply(strNos, 1)
    varTemp = Split(strNos, ",")
    For i = 0 To UBound(varTemp)
        strNo = varTemp(i)
        Select Case bytMode
            Case EM_MULTI_�˷�����
                rsTemp.Filter = "NO='" & strNo & "' And ״̬=0" '������
                If rsTemp.RecordCount <> 0 Then
                    MsgBox "����:" & strNo & " �ѱ��˷����룬�����ٽ������룡", vbInformation, gstrSysName
                    Exit Function
                End If
                rsTemp.Filter = "NO='" & strNo & "' And ״̬=1" '�����
                If rsTemp.RecordCount <> 0 Then
                    MsgBox "����:" & strNo & " �ѱ��˷����벢��������ˣ������ٽ������룡", vbInformation, gstrSysName
                    Exit Function
                End If
            Case EM_MULTI_ȡ������
                rsTemp.Filter = "NO='" & strNo & "' And ״̬=0" '������
                If rsTemp.RecordCount = 0 Then
                    MsgBox "����:" & strNo & " �ѱ�ȡ�����룬�����ٽ���ȡ�����룡", vbInformation, gstrSysName
                    Exit Function
                End If
            Case EM_MULTI_�˷����, EM_MULTI_�ܾ�����, EM_MULTI_ȡ�����
                rsTemp.Filter = "NO='" & strNo & "' And ����ʱ��=#" & mstrApplyTime & "#" '������
                If rsTemp.RecordCount = 0 Then
                    MsgBox "����:" & strNo & " �ѱ�ȡ�����룬���ܽ���" & _
                            IIf(bytMode = EM_MULTI_�˷����, "�˷����", IIf(bytMode = EM_MULTI_�ܾ�����, "�ܾ�����", "ȡ�����")) & "��", vbInformation, gstrSysName
                    Exit Function
                End If
                If bytMode = EM_MULTI_�˷���� Then
                    rsTemp.Filter = "(NO='" & strNo & "' And ״̬=1) " & _
                                    "Or (NO='" & strNo & "' And ״̬=2 And ����ʱ��=#" & mstrApplyTime & "#)" '����˻�ܾ�
                    If rsTemp.RecordCount <> 0 Then
                        MsgBox "����:" & strNo & " �ѱ��˷���˻�ܾ����룬�����ٽ����˷���ˣ�", vbInformation, gstrSysName
                        Exit Function
                    End If
                ElseIf bytMode = EM_MULTI_�ܾ����� Then
                    rsTemp.Filter = "NO='" & strNo & "' And ״̬=2 And ����ʱ��=#" & mstrApplyTime & "#"
                    If rsTemp.RecordCount <> 0 Then
                        MsgBox "����:" & strNo & " �ѱ��ܾ����룬�����ٽ��оܾ����룡", vbInformation, gstrSysName
                        Exit Function
                    End If
                Else
                    rsTemp.Filter = "NO='" & strNo & "' And ״̬=0 And ����ʱ��=#" & mstrApplyTime & "#"
                    If rsTemp.RecordCount <> 0 Then
                        MsgBox "����:" & strNo & " �ѱ�ȡ����ˣ������ٽ���ȡ����ˣ�", vbInformation, gstrSysName
                        Exit Function
                    End If
                    If BillExistDelete(strNo, 1) Then
                        MsgBox "����:" & strNo & " ���˷ѣ�����ȡ����ˡ�", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
        End Select
    Next
    
    CheckDelAppliedValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function

Private Function SaveDelApplied(ByVal bytMode As gEM_ChargeDelType) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����˷�����
    '���:strNos-����ĵ��ݺ�,����ö��ŷ���
    '����:�˷�����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-08-05 11:14:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPro As Collection, strSQL As String
    Dim strDate As String, varNO As Variant
    Dim strԭ�� As String, strNos As String
    Dim i As Long
    
    On Error GoTo errHandle
    
    If CheckDelAppliedValied(bytMode, strNos) = False Then Exit Function
      
    strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    strԭ�� = txt�˷�ժҪ.Text
    Set cllPro = New Collection
    varNO = Split(strNos, ",")
    For i = 0 To UBound(varNO)
        Select Case bytMode
            Case EM_MULTI_�˷�����
                'Zl_�����˷�����_Apply
                strSQL = "Zl_�����˷�����_Apply("
                '  ��������_In Number,
                strSQL = strSQL & "" & "0" & ","
                '  No_In       �����˷�����.No%Type,
                strSQL = strSQL & "'" & varNO(i) & "',"
                '  ��¼����_In �����˷�����.��¼����%Type,
                strSQL = strSQL & "" & "1" & ","
                '  ������_In   �����˷�����.������%Type,
                strSQL = strSQL & "'" & UserInfo.���� & "',"
                '  ����ʱ��_In �����˷�����.����ʱ��%Type,
                strSQL = strSQL & "" & "To_Date('" & strDate & "','YYYY-MM-DD HH24:MI:SS')" & ","
                '  ����ԭ��_In �����˷�����.����ԭ��%Type := Null
                strSQL = strSQL & "'" & strԭ�� & "')"
            Case EM_MULTI_ȡ������
                'Zl_�����˷�����_Apply
                strSQL = "Zl_�����˷�����_Apply("
                '  ��������_In Number,
                strSQL = strSQL & "" & "1" & ","
                '  No_In       �����˷�����.No%Type,
                strSQL = strSQL & "'" & varNO(i) & "',"
                '  ��¼����_In �����˷�����.��¼����%Type,
                strSQL = strSQL & "" & "1" & ","
                '  ������_In   �����˷�����.������%Type,
                strSQL = strSQL & "'" & "" & "',"
                '  ����ʱ��_In �����˷�����.����ʱ��%Type,
                strSQL = strSQL & "" & "To_Date('" & mstrApplyTime & "','YYYY-MM-DD HH24:MI:SS')" & ")"
                '  ����ԭ��_In �����˷�����.����ԭ��%Type := Null
            Case EM_MULTI_�˷����
                'Zl_�����˷�����_Audit
                strSQL = "Zl_�����˷�����_Audit("
                '  No_In       �����˷�����.No%Type,
                strSQL = strSQL & "'" & varNO(i) & "',"
                '  ��¼����_In �����˷�����.��¼����%Type,
                strSQL = strSQL & "" & "1" & ","
                '  ����ʱ��_In �����˷�����.����ʱ��%Type,
                strSQL = strSQL & "" & "To_Date('" & mstrApplyTime & "','YYYY-MM-DD HH24:MI:SS')" & ","
                '  �����_In   �����˷�����.�����%Type,
                strSQL = strSQL & "'" & UserInfo.���� & "',"
                '  ���ʱ��_In �����˷�����.���ʱ��%Type,
                strSQL = strSQL & "" & "To_Date('" & strDate & "','YYYY-MM-DD HH24:MI:SS')" & ","
                '  ���ԭ��_In �����˷�����.���ԭ��%Type := Null,
                strSQL = strSQL & "'" & strԭ�� & "',"
                '  ״̬_In     �����˷�����.״̬%Type := 1
                '--       ״̬_In��1-���ͨ����2-�ܾ�����(��˲�ͨ��)��3-ȡ�����
                strSQL = strSQL & "" & "1" & ")"
            Case EM_MULTI_�ܾ�����
                'Zl_�����˷�����_Audit
                strSQL = "Zl_�����˷�����_Audit("
                '  No_In       �����˷�����.No%Type,
                strSQL = strSQL & "'" & varNO(i) & "',"
                '  ��¼����_In �����˷�����.��¼����%Type,
                strSQL = strSQL & "" & "1" & ","
                '  ����ʱ��_In �����˷�����.����ʱ��%Type,
                strSQL = strSQL & "" & "To_Date('" & mstrApplyTime & "','YYYY-MM-DD HH24:MI:SS')" & ","
                '  �����_In   �����˷�����.�����%Type,
                strSQL = strSQL & "'" & UserInfo.���� & "',"
                '  ���ʱ��_In �����˷�����.���ʱ��%Type,
                strSQL = strSQL & "" & "To_Date('" & strDate & "','YYYY-MM-DD HH24:MI:SS')" & ","
                '  ���ԭ��_In �����˷�����.���ԭ��%Type := Null,
                strSQL = strSQL & "'" & strԭ�� & "',"
                '  ״̬_In     �����˷�����.״̬%Type := 1
                '--       ״̬_In��1-���ͨ����2-�ܾ�����(��˲�ͨ��)��3-ȡ�����
                strSQL = strSQL & "" & "2" & ")"
            Case EM_MULTI_ȡ�����
                'Zl_�����˷�����_Audit
                strSQL = "Zl_�����˷�����_Audit("
                '  No_In       �����˷�����.No%Type,
                strSQL = strSQL & "'" & varNO(i) & "',"
                '  ��¼����_In �����˷�����.��¼����%Type,
                strSQL = strSQL & "" & "1" & ","
                '  ����ʱ��_In �����˷�����.����ʱ��%Type,
                strSQL = strSQL & "" & "To_Date('" & mstrApplyTime & "','YYYY-MM-DD HH24:MI:SS')" & ","
                '  �����_In   �����˷�����.�����%Type,
                strSQL = strSQL & "" & "NULL" & ","
                '  ���ʱ��_In �����˷�����.���ʱ��%Type,
                strSQL = strSQL & "" & "NULL" & ","
                '  ���ԭ��_In �����˷�����.���ԭ��%Type := Null,
                strSQL = strSQL & "" & "NULL" & ","
                '  ״̬_In     �����˷�����.״̬%Type := 1
                '--       ״̬_In��1-���ͨ����2-�ܾ�����(��˲�ͨ��)��3-ȡ�����
                strSQL = strSQL & "" & "3" & ")"
        End Select
        zlAddArray cllPro, strSQL
    Next
    On Error GoTo ErrRoll:
    zlExecuteProcedureArrAy cllPro, Me.Caption
    SaveDelApplied = True
    Exit Function
ErrRoll:
    gcnOracle.RollbackTrans
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function CheckIsExistDelErrBill(ByVal strNos As String, Optional ByRef str����Ա���� As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݵ��ݺ�,����Ƿ�����˷��쳣��¼
    '���:
    '     strNOs=���ݺ�,��ʽ NO1,NO2,NO3,...
    '����:
    '     strUser=�����˷��쳣���ݵĲ���Ա����
    '����:�����˷��쳣����,����true,���򷵻�False
    '����:Ƚ����
    '����:2014-08-18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    str����Ա���� = ""
    If strNos = "" Then Exit Function
   
    On Error GoTo Errhand
    strSQL = "" & _
            " Select ����Ա����" & _
            " From ������ü�¼ A" & _
            " Where Nvl(����״̬, 0) = 1 And ��¼���� = 1 And ��¼״̬ = 2" & _
            "       And a.No In (Select Column_Value From Table(f_Str2list([1])))" & _
            "       And Not Exists (Select 1 From ����Ԥ����¼ B Where a.����id = b.����id And Nvl(b.У�Ա�־, 0) = 0)" & _
            "       And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����Ƿ�����˷��쳣��¼", strNos)
    
    If Not rsTemp.EOF Then
        str����Ա���� = Nvl(rsTemp!����Ա����)
        CheckIsExistDelErrBill = True
    End If
    
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
End Function

Public Sub zlGetClassMoney(ByRef rsClass As ADODB.Recordset)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������ܽ��
    '����:���˺�
    '����:2011-12-26 13:19:04
    '����:44944
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    On Error GoTo errHandle
    Set rsClass = New ADODB.Recordset
    rsClass.Fields.Append "�շ����", adVarChar, 10, adFldIsNullable
    rsClass.Fields.Append "���", adDouble, , adFldIsNullable
    rsClass.CursorLocation = adUseClient
    rsClass.LockType = adLockOptimistic
    rsClass.CursorType = adOpenStatic
    rsClass.Open
    With vsBill
        For i = 1 To .Rows - 1
'            If .TextMatrix(i, .ColIndex("ѡ��")) <> 0 Then
                rsClass.Find "�շ����='" & .Cell(flexcpData, i, .ColIndex("���")) & "'", , adSearchForward, 1
                If rsClass.EOF Then rsClass.AddNew
                rsClass!�շ���� = .Cell(flexcpData, i, .ColIndex("���"))
                rsClass!��� = Val(Nvl(rsClass!���)) + .TextMatrix(i, .ColIndex("ʵ�ս��"))
                rsClass.Update
'            End If
        Next
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Function ExecuteInsureInfoUpdate(ByVal lng����ID As Long, ByRef str���ս�� As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ռ�¼�ı�����Ϣ
    '����:
    '   str���ս��-"ʵ�պϼ�;����ͳ��;ȫ�Ը�;���Ը�"
    '����:�������ռ�¼�ı�����Ϣ���³ɹ�����True�����򷵻�False
    '����:Ƚ����
    '����:2014-9-16
    '����:77951
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsReCharge As ADODB.Recordset, strBXInfo As String, cllReChargePro As Collection
    Dim blnTrans As Boolean
    Dim curʵ�պϼ� As Currency, cur����ͳ�� As Currency
    Dim curȫ�Ը� As Currency, cur���Ը� As Currency
    Dim curʵ�ս�� As Currency, curͳ���� As Currency, bln������Ŀ As Boolean
    
    On Error GoTo Errhand
    str���ս�� = ""
    strSQL = " Select a.Id, a.����id, a.�շ�ϸĿid, Nvl(a.����, 1) * Nvl(a.����, 0) As ����, Nvl(a.ʵ�ս��, 0) As ʵ�ս��, a.ժҪ, " & _
            " Nvl(a.������Ŀ��, 0) As ������Ŀ��, a.���մ���id, Nvl(a.ͳ����, 0) As ͳ����, a.���ձ���, a.��������" & _
            " From ������ü�¼ A" & _
            " Where a.��¼���� = 11 And a.����id = [1]"
    Set rsReCharge = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���շ��ü�¼", lng����ID)
    With rsReCharge
        If .RecordCount > 0 Then
            Set cllReChargePro = New Collection
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
                    zlAddArray cllReChargePro, strSQL
                    
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
                rsReCharge.MoveNext
            Loop
            'ִ�й���
            blnTrans = True
            zlExecuteProcedureArrAy cllReChargePro, Me.Caption, True, True
            blnTrans = False
        End If
    End With
    '���ս����Ϣ
    str���ս�� = curʵ�պϼ� & ";" & cur����ͳ�� & ";" & curȫ�Ը� & ";" & cur���Ը�
    ExecuteInsureInfoUpdate = True
    Exit Function
Errhand:
    If blnTrans Then
        gcnOracle.RollbackTrans
        Call ErrCenter
    Else
        If ErrCenter = 1 Then Resume
    End If
End Function

Private Function SelectMulitBalance(ByVal strNos As String, ByRef strNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ѡ���ν����һ�ν��㵥��
    '���:strNos-���ݺ�,����ö���
    '     strNo -��ǰ����ĵ��ݺ�,��ֻһ�ν���ʱ��ֱ�ӷ���
    '����:strNo-���ص�ǰѡ�еĵ��ݺ�
    '����:ѡ��ɹ�������true,���򷵻�False
    '����:���˺�
    '����:2016-05-04 17:16:56
    '˵����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, rsSel As ADODB.Recordset
    Dim strWithTable As String, cllPro As Collection, varData() As Variant
    On Error GoTo errHandle
    
    '�ȼ�鱾�ΰ����˴�ӡ�Ƿ�ֻ��һ�ν���ģ����ֻ��һ�ν��㣬��ֱ���˳�,������ѡ��
   If Len(strNos) <= 4000 Then  '����4000,����м��ν���
       strSQL = "" & _
       " Select /*+cardinality(b,10)*/ Count(Distinct nvl(C.�������,C.����ID)) as ���� " & vbNewLine & _
       " From ������ü�¼ A,����Ԥ����¼ C, Table(f_Str2list([1])) B" & vbNewLine & _
       " where A.����ID=C.����ID And Mod(A.��¼����, 10) = 1  And A.��¼״̬ in (1,3) And A.NO=B.Column_Value "
      Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNos)
      If Nvl(rsTemp!����, 0) <= 1 Then
            SelectMulitBalance = True
            Exit Function
      End If
   End If

    If Len(strNos) <= 4000 Then
        strSQL = "" & _
        " Select A.NO as ���ݺ�,max(C.����) as ��������, " & _
        "       A.���,sum(nvl(A.����,1)*nvl(A.����,0)) as ����, " & _
        "       max(decode(A.��¼״̬,1,A.����Ա���,3,A.����Ա���,NULL)) as ����Ա���,max(decode(A.��¼״̬,1,A.����Ա����,3,A.����Ա����,NULL)) as ����Ա����, " & _
        "       to_char(max(decode(A.��¼״̬,1,a.�Ǽ�ʱ��,3,a.�Ǽ�ʱ��,NULL )),'yyyy-mm-dd hh24:mi:ss') as �տ�ʱ��" & vbNewLine & _
        " From ������ü�¼ A, Table(f_Str2list([1])) B,���ű� C" & vbNewLine & _
        " where Mod(A.��¼����, 10) = 1 And A.�۸񸸺� is null  And A.NO=B.Column_Value " & _
        " AND A.��������id=C.id " & _
        " Group by A.NO,A.���" & _
        " Having sum(nvl(A.����,1)*nvl(A.����,0)) <>0"
        strSQL = "" & _
        "   Select distinct ���ݺ�,��������,����Ա���,����Ա����,�տ�ʱ�� " & _
        "   From (" & strSQL & ")" & _
        "   Order by �տ�ʱ��,���ݺ�"

        
        strSQL = "" & _
        "  Select Rownum as ID,���ݺ�,��������,����Ա���,����Ա����,�տ�ʱ�� " & _
        "  From (" & strSQL & ") "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNos)
    Else
        
        If zlGetSplitString4000(strNos, cllPro) = False Then Exit Function
        If zlFromCollectBulidSQL(cllPro, strSQL, varData) = False Then Exit Function
        
        strSQL = "With ������Ϣ as (" & strSQL & ")" & vbCrLf
        strSQL = strSQL & vbCrLf & _
        " Select A.NO as ���ݺ�,max(C.����) as ��������, " & _
        "       A.���,sum(nvl(A.����,1)*nvl(A.����,0)) as ����, " & _
        "       max(decode(A.��¼״̬,1,A.����Ա���,3,A.����Ա���,NULL)) as ����Ա���,max(decode(A.��¼״̬,1,A.����Ա����,3,A.����Ա����,NULL)) as ����Ա����, " & _
        "       to_char(max(decode(A.��¼״̬,1,a.�Ǽ�ʱ��,3,a.�Ǽ�ʱ��,NULL )),'yyyy-mm-dd hh24:mi:ss') as �տ�ʱ��" & vbNewLine & _
        " From ������ü�¼ A, ������Ϣ B,���ű� C" & vbNewLine & _
        " where Mod(A.��¼����, 10) = 1 And A.�۸񸸺� is null  And A.NO=B.NO " & _
        " AND A.��������id=C.id " & _
        " Group by A.NO,A.���" & _
        " Having sum(nvl(A.����,1)*nvl(A.����,0)) <>0"
       
        strSQL = "" & _
        "   Select distinct ���ݺ�,��������,����Ա���,����Ա����,�տ�ʱ�� " & _
        "   From (" & strSQL & ")" & _
        "   Order by �տ�ʱ��,���ݺ�"
        
        strSQL = "" & _
        "  Select Rownum as ID,���ݺ�,��������,����Ա���,����Ա����,�տ�ʱ�� " & _
        "  From (" & strSQL & ") "
        Set rsTemp = zlDatabase.OpenSQLRecordByArray(strSQL, "�����ν���", varData)
    End If

    If rsTemp Is Nothing Then Exit Function
    If rsTemp.State <> 1 Then Exit Function
    If rsTemp.EOF Then Exit Function
    If rsTemp.RecordCount = 1 Then
        strNo = Nvl(rsSel!���ݺ�)
        SelectMulitBalance = True: Exit Function
    End If
    
    If zlDatabase.zlShowListSelect(Me, glngSys, mlngModule, txtNO, rsTemp, True, "ѡ��ָ������", , rsSel) = False Then Exit Function
    If rsSel Is Nothing Then Exit Function
    If rsSel.State <> 1 Then Exit Function
    If rsSel.EOF Then Exit Function
    strNo = Nvl(rsSel!���ݺ�)
    SelectMulitBalance = True
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
                        If .Visible And .Enabled Then .Row = j: .SetFocus
                        Exit Function
                    ElseIf Val(Nvl(rsTemp!ԭʼ����)) > 0 Then
                        '�����շѵĲ����
                        dblʣ������ = Val(Nvl(rsTemp!����, 1)) * Val(Nvl(rsTemp!����))
                        dbl�������� = Val(.TextMatrix(j, .ColIndex("����")))
                        If RoundEx(dbl��������, 6) > RoundEx(dblʣ������, 6) Then
                            MsgBox "����:" & arrNo(i) & " �е� " & (j - k + 1) & " ����Ŀ�ı����˷�����(" & _
                                FormatEx(dbl��������, 5) & ")������ʣ��δ������(" & FormatEx(dblʣ������, 5) & ")��" & _
                                "�����˷ѣ������»�ȡ�������ݣ�", vbExclamation, gstrSysName
                            If .Visible And .Enabled Then .Row = j: .SetFocus
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

Private Function GetFeeListData(ByVal strNos As String, ByRef rsFeeList As ADODB.Recordset) As Boolean
    '��ȡ���˷ѵ�������
    '����:rsFeeList-����׼�˷Ѽ�
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '----------------------------------------------------------------------------------
    '�˷�ʱ���ÿ��Ǻ󱸱�,ǰ��Ĳ����ѽ���
    '��ȡ׼����,������Ӧ�ս��,ʵ�ս��(���=ʣ����*(׼����/ʣ����))
    '��ȡ������ԭʼ��¼�ķ���ID
    Dim strSQL As String
    Dim strTableNo As String, strSQLIn As String
    Dim strSqlSub As String
    
    On Error GoTo errHandler
    strSqlSub = _
        " Select /*+cardinality(j,10)*/ a.Id, a.��¼����, a.No, a.��¼״̬, a.���, a.��������, a.�۸񸸺�, a.�շ�ϸĿid," & vbNewLine & _
        "        Nvl(a.����, 1) As ����, Nvl(a.����, 0) As ����," & vbNewLine & _
        "        Nvl(a.Ӧ�ս��, 0) As Ӧ�ս��, Nvl(a.ʵ�ս��, 0) As ʵ�ս��, Nvl(a.���ʽ��, 0) As ���ʽ��," & vbNewLine & _
        "        Nvl(a.����, 1) * a.���� As ����, Nvl(��׼����, 0) As ��׼����," & vbNewLine & _
                 IIf(gblnҩ����λ, "Nvl(b." & gstrҩ����װ & ",1)", "1") & " As ����ϵ��, " & vbNewLine & _
                 IIf(gblnҩ����λ, "Decode(B.ҩƷID,NULL,A.���㵥λ,B." & gstrҩ����λ & ")", "A.���㵥λ ") & " As ���㵥λ," & vbNewLine & _
        "        a.��������id, a.ִ�в���id, a.ҽ�����, " & vbNewLine & _
        "        a.ִ��״̬,a.��������, a.����״̬, a.���ӱ�־,a.�ѱ�, a.�շ����, a.����Ա����, a.�Ǽ�ʱ��, a.����id," & vbNewLine & _
        "        b.ҩƷid" & vbNewLine & _
        " From ������ü�¼ A, ҩƷ��� B, Table(f_Str2list([1])) J" & vbNewLine & _
        " Where Mod(a.��¼����, 10) = 1 And a.No = j.Column_Value And a.��¼״̬ <> 0" & _
        "       And a.�շ�ϸĿid = b.ҩƷid(+)"
    '��׼�˷�(����,ҩƷ,����������)
    strTableNo = _
        " With ������� As (" & strSqlSub & ")," & vbNewLine & _
        "      ׼���� As (Select /*+cardinality(j,10)*/ A.����ID," & _
        "                        Sum(Nvl(A.����,1)*A.ʵ������" & IIf(gblnҩ����λ, "/Nvl(B." & gstrҩ����װ & ",1)", "") & ") as ׼������" & vbNewLine & _
        "                 From ҩƷ�շ���¼ A,ҩƷ��� B, Table(f_Str2list([1])) J" & vbNewLine & _
        "                 Where A.ҩƷID=B.ҩƷID(+) And Mod(A.��¼״̬,3)=1  " & vbNewLine & _
        "                       And (A.���� =8 or a.����=24) And A.����� is NULL And A.NO =J.Column_Value" & vbNewLine & _
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
        "       A.�ѱ�,C.���� as �����,C.���� as �����,A.�շ�ϸĿID,B.����,B.����,B.���," & _
        "       Max(Nvl(A.��������,B.��������)) ��������," & _
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
        " Select A.NO,A.���,A.��������,A.�ѱ�,A.�����,A.�����,A.�շ�ϸĿID,A.����,A.����,A.���," & _
        "       Max(A.��������) As ��������,A.���㵥λ, Max(A.ҽ�����) as ҽ�����," & _
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
        " Where     A.�շ�ϸĿID=B.�շ�ϸĿID(+) And B.����(+)=1 And B.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
        "       And A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3 " & _
        " Order by A.NO,A.���"

    Set rsFeeList = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���˷���Ŀ", strNos)
    GetFeeListData = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
       Resume
    End If
    Call SaveErrLog
End Function

Public Function GetDelXMLExpend() As String
    '��ȡ�����������˷ѽӿ�zlRetuenCheck��strXMLExpend����ֵ
    If mbytMode = EM_MULTI_�˷� Then
        GetDelXMLExpend = ZlGetDelXMLExpendByGrid(Me.vsBill)
    ElseIf mbytMode = EM_MULTI_�쳣���� Then
        GetDelXMLExpend = ZlGetDelXMLExpend(mlng�������, True)
    End If
End Function
