VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#6.9#0"; "zlIDKind.ocx"
Begin VB.Form frmPatiBooks 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�˲�����"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8910
   Icon            =   "frmPatiBooks.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   2235
      Left            =   0
      TabIndex        =   25
      Top             =   4065
      Width           =   7575
      Begin VSFlex8Ctl.VSFlexGrid vsPay 
         Height          =   1740
         Left            =   120
         TabIndex        =   26
         ToolTipText     =   "�������ֵ�����£���ѡ���˷ѷ�ʽ����""�ֽ�""��ʽ�����˷�"
         Top             =   465
         Width           =   7425
         _cx             =   13097
         _cy             =   3069
         Appearance      =   3
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
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   0
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPatiBooks.frx":6852
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
         AllowUserFreezing=   2
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.PictureBox picPay 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         ForeColor       =   &H80000008&
         Height          =   1725
         Left            =   5130
         ScaleHeight     =   1695
         ScaleWidth      =   2250
         TabIndex        =   32
         Top             =   480
         Width           =   2280
         Begin VB.TextBox txtӦ��1 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            IMEMode         =   3  'DISABLE
            Left            =   855
            MaxLength       =   10
            TabIndex        =   35
            Top             =   90
            Width           =   1290
         End
         Begin VB.TextBox txt�ɿ� 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            IMEMode         =   3  'DISABLE
            Left            =   855
            MaxLength       =   12
            TabIndex        =   34
            Top             =   615
            Width           =   1290
         End
         Begin VB.TextBox txt�Ҳ� 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            IMEMode         =   3  'DISABLE
            Left            =   855
            Locked          =   -1  'True
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   1155
            Width           =   1290
         End
         Begin VB.Label lbl��Ԥ�� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ӧ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   15.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   38
            Top             =   195
            Width           =   660
         End
         Begin VB.Label lblPayType 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�˿�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   15.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   120
            TabIndex        =   37
            Top             =   705
            Width           =   660
         End
         Begin VB.Label lbl�Ҳ� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�Ҳ�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   15.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   135
            TabIndex        =   36
            Top             =   1230
            Width           =   660
         End
      End
      Begin VB.TextBox txtδ�� 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   450
         IMEMode         =   3  'DISABLE
         Left            =   6255
         MaxLength       =   10
         TabIndex        =   30
         Top             =   -15
         Width           =   1290
      End
      Begin VB.TextBox txtӦ�� 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         IMEMode         =   3  'DISABLE
         Left            =   4185
         MaxLength       =   10
         TabIndex        =   28
         Top             =   0
         Width           =   1290
      End
      Begin VB.Label lblδ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "δ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5565
         TabIndex        =   31
         Top             =   75
         Width           =   660
      End
      Begin VB.Label lblӦ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ӧ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3480
         TabIndex        =   29
         Top             =   75
         Width           =   660
      End
      Begin VB.Label Label4 
         Caption         =   "�����˷����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   27
         Top             =   135
         Width           =   2145
      End
   End
   Begin VB.Frame fraSplit3 
      Height          =   6435
      Left            =   7620
      TabIndex        =   24
      Top             =   -90
      Width           =   45
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   45
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   7620
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   7710
      TabIndex        =   11
      Top             =   705
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   7710
      TabIndex        =   10
      Top             =   210
      Width           =   1095
   End
   Begin VB.Frame fraSplit1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   45
      Left            =   0
      TabIndex        =   15
      Top             =   1410
      Width           =   7620
   End
   Begin VB.Frame fraSplit2 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   45
      Left            =   0
      TabIndex        =   12
      Top             =   3990
      Width           =   7620
   End
   Begin VB.Frame fraGroup 
      BorderStyle     =   0  'None
      Height          =   2505
      Left            =   0
      TabIndex        =   14
      Top             =   1470
      Width           =   7620
      Begin VSFlex8Ctl.VSFlexGrid vsBooks 
         Height          =   2385
         Left            =   120
         TabIndex        =   9
         Top             =   60
         Width           =   7440
         _cx             =   13123
         _cy             =   4207
         Appearance      =   3
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
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   0
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPatiBooks.frx":6936
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
         AllowUserFreezing=   2
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.Frame fraPatiCard 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   7575
      Begin VB.TextBox txt�ֻ� 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         IMEMode         =   3  'DISABLE
         Left            =   4395
         MaxLength       =   18
         TabIndex        =   8
         Tag             =   "�����"
         Top             =   975
         Width           =   1290
      End
      Begin VB.TextBox txtPatient 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1285
         TabIndex        =   0
         Tag             =   "����"
         Top             =   120
         Width           =   2205
      End
      Begin VB.TextBox txtNation 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   6300
         TabIndex        =   6
         Top             =   555
         Width           =   1170
      End
      Begin VB.TextBox txt���֤�� 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         IMEMode         =   3  'DISABLE
         Left            =   1290
         MaxLength       =   18
         TabIndex        =   7
         Tag             =   "���֤��"
         Top             =   975
         Width           =   2205
      End
      Begin VB.TextBox txt����� 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         IMEMode         =   3  'DISABLE
         Left            =   6315
         MaxLength       =   18
         TabIndex        =   2
         Tag             =   "�����"
         Top             =   120
         Width           =   1170
      End
      Begin VB.TextBox txt���� 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   4395
         TabIndex        =   5
         Top             =   555
         Width           =   990
      End
      Begin VB.TextBox txt�Ա� 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   4395
         TabIndex        =   1
         Top             =   120
         Width           =   1005
      End
      Begin MSMask.MaskEdBox txt����ʱ�� 
         Height          =   345
         Left            =   2505
         TabIndex        =   4
         Tag             =   "����ʱ��"
         Top             =   555
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   609
         _Version        =   393216
         BackColor       =   -2147483633
         Enabled         =   0   'False
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "hh:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txt�������� 
         Height          =   345
         Left            =   1305
         TabIndex        =   3
         Tag             =   "��������"
         Top             =   555
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   609
         _Version        =   393216
         BackColor       =   -2147483633
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "YYYY-MM-DD"
         Mask            =   "####-##-##"
         PromptChar      =   "_"
      End
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   360
         Left            =   670
         TabIndex        =   39
         Top             =   120
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   635
         Appearance      =   2
         IDKindStr       =   $"frmPatiBooks.frx":6A14
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
         AllowAutoICCard =   -1  'True
         AllowAutoIDCard =   -1  'True
         BackColor       =   -2147483633
      End
      Begin VB.Label lbl�ֻ��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ֻ���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3735
         TabIndex        =   40
         Top             =   1035
         Width           =   630
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5865
         TabIndex        =   22
         Top             =   600
         Width           =   420
      End
      Begin VB.Label lbl���֤�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���֤��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   420
         TabIndex        =   21
         Top             =   1050
         Width           =   840
      End
      Begin VB.Label lbl����� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5660
         TabIndex        =   20
         Top             =   180
         Width           =   630
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3930
         TabIndex        =   19
         Top             =   615
         Width           =   420
      End
      Begin VB.Label lbl�������� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   420
         TabIndex        =   18
         Top             =   615
         Width           =   840
      End
      Begin VB.Label lbl�Ա� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3915
         TabIndex        =   17
         Top             =   180
         Width           =   420
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   165
         TabIndex        =   16
         Top             =   210
         Width           =   420
      End
   End
End
Attribute VB_Name = "frmPatiBooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbytFunc As Long    '1-�鿴;2-�༭
Private mlngModule As Long
Private mobjKeyboard As Object
Private mrsBooks As ADODB.Recordset
Private mobjPayCards As Cards
Private mobjPayCard As Card '������ԭ����֧����ʽ
Private mobjDelObjects  As clsCardObjects
Private mobjDelObject  As clsCardObject
Private mblnUnLoad As Boolean
Private mblnNotClick As Boolean
Private mstr���㷽ʽ As String
Private mlngҽ�ƿ�����  As Long
Private Type T_Pati
    ����ID As Long
    ���� As String
    �Ա� As String
    ���� As String
    ���� As String
    ���� As String
    ���� As String
    �������� As String
    ����� As Long
    ���֤�� As String
    �ֻ��� As String
    �������� As String
End Type
Private mPati As T_Pati

Private Type M_PayInfo
    blnȫ�� As Boolean
    bln���� As Boolean
    blnҽ�� As Boolean
End Type
Private mPayInfo As M_PayInfo

Private Const C_BookInfoColumHeader = "���ݺ�,905,1;����,605,4;������,905,7;����,605,4;����ʱ��,1205,4;����Ա,905,1;��ע,605,1" '��ʽ:"����","����","�п�"(���ж���ȡֵΪ:1-����� 4-���� 7-�Ҷ���)
Private Const C_PayInfoColumHeader = "ѡ��,605,4;֧����ʽ,1205,1;֧�����,1005,7;����,605,1;������ˮ��,1505,1;����˵��,1205,1;��ע,605,1" '��ʽ:"����","����","�п�"(���ж���ȡֵΪ:1-����� 4-���� 7-�Ҷ���)
Private Const C_COLOR_���� = &H80000005
Private Const C_COLOR_��ɫ = &H80000005
Private Const C_COLOR_��ɫ = &H8000000D

Public Sub ShowMe(ByRef frmMain As Object, ByVal bytFunc As Byte, ByVal lngModul As Long, ByVal lng����ID As Long, ByVal str���� As String)
'����:��ʾ������
'����: FrmMain-������
'      bytFunc=1-�鿴,2-�༭
'      lng����ID-�鿴ʱ���� ��bytFunc=1ʱ����,bytFunc=2ʱˢ����ȡ��
'      lngModul ģ���

    mbytFunc = bytFunc
    mlngModule = lngModul
    mPati.����ID = lng����ID
    mPati.���� = str����
    Me.Show 1, frmMain
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDel_Click()
    '������֧�����ֵ�����£�ɾ���������˿ʽ�������ֽ��˷ѷ�ʽ������������ֽ�ʽ�������˿���
    Dim dblMoney As Double, i As Integer, blnFind As Boolean
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo Errhand
    strSQL = "Select ���� From ���㷽ʽ Where ����=1 Order By ȱʡ��־ Desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rsTemp.RecordCount = 0 Then Exit Sub
    With vsPay
        dblMoney = CDbl(.TextMatrix(.RowSel, .ColIndex("֧�����")))
        .RemoveItem
        
        For i = 1 To .Rows - 1
            If .Cell(flexcpData, i, .ColIndex("֧����ʽ")) = 1 Then
                blnFind = True
                dblMoney = dblMoney + CDbl(.TextMatrix(i, .ColIndex("֧����ʽ")))
                .TextMatrix(i, .ColIndex("֧�����")) = Format(dblMoney, "0.00")
            End If
            If blnFind Then Exit For
        Next
        
        If Not blnFind Then
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("֧����ʽ")) = Nvl(rsTemp!����)
            .Cell(flexcpData, .Rows - 1, .ColIndex("֧����ʽ")) = 1
            .TextMatrix(.Rows - 1, .ColIndex("֧�����")) = Format(dblMoney, "0.00")
            .RowData(.Rows - 1) = 1
        End If
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Sub

Private Sub cmdOK_Click()
    If IsCheckCancelValied = False Then Exit Sub
    If IsNoCanc = False Then Exit Sub
    If SaveData = False Then Exit Sub
    Call ClearPatiInfo
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0: Exit Sub
End Sub

Private Sub Form_Load()
    '��ʼ��
    Call InitFace
    Call InitVsFlex
    Call Init�˷ѷ�ʽ
    Call LoadPatiBooks
    Call IDKind.zlInit(Me, glngSys, mlngModule, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, , txtPatient)

    If mPati.����ID <> 0 Then Call LoadPati(mPati.����ID)
    
End Sub

Private Sub Form_Resize()
    Dim lngW As Long
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjKeyboard = Nothing
    Set mrsBooks = Nothing
    Set mobjPayCards = Nothing
    Set mobjPayCard = Nothing
    Set mobjDelObject = Nothing
End Sub

Private Sub InitFace()
    Dim objCtl As Object
    
    For Each objCtl In Me.Controls
        Select Case UCase(TypeName(objCtl))
            Case "TEXTBOX", "MASKEDBOX"
                objCtl.Enabled = False
        End Select
    Next
    txtPatient.Enabled = True
    txtӦ��.Enabled = True: txtӦ��.Locked = True
    txtδ��.Enabled = True: txtδ��.Locked = True
End Sub

Private Sub txtPatient_Change()
    If Trim(txtPatient.Text) = "" Then
        Call ClearPatiInfo
    End If
End Sub

Private Sub LoadPati(ByVal lng����ID As Long)
    '����:���ز�����Ϣ
    Dim strSQL As String, rsTemp As ADODB.Recordset
    strSQL = "Select A.����,A.�Ա�,A.����,A.��������,A.����,A.���֤��,A.�����,A.�ֻ���," & vbNewLine & _
    "         Nvl(Nvl(A.��������,B.��������),Decode(Nvl(A.����,B.����),Null,'��ͨ����','ҽ������')) �������� " & vbNewLine & _
    " From ������Ϣ A,������ҳ B " & vbNewLine & _
    " Where A.����ID=B.����ID(+) And A.��ҳID=B.��ҳID(+) And A.ͣ��ʱ�� is NULL And A.����ID=[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "������Ϣ", lng����ID)
    With rsTemp
        If .RecordCount = 0 Then Exit Sub
        txtPatient.Text = Nvl(!����)
        txt�Ա�.Text = Nvl(!�Ա�)
        txtNation.Text = Nvl(!����)
        txt����.Text = Nvl(!����)
        txt��������.Text = Format(Nvl(!��������), "YYYY-MM-DD")
        txt����ʱ��.Text = Format(Nvl(!��������), "MM:SS")
        txt�����.Text = Nvl(!�����)
        txt���֤��.Text = Nvl(!���֤��)
        txt�ֻ�.Text = Nvl(!�ֻ���)
        
        mPati.����ID = lng����ID
        mPati.���� = Nvl(!����)
        mPati.���� = ""
        mPati.���� = Nvl(!����)
        mPati.�Ա� = Nvl(!�Ա�)
        mPati.�������� = Format(Nvl(!��������), "YYYY-MM-DD MM:SS")
        mPati.���� = Nvl(!����)
        mPati.����� = Nvl(!�����)
        mPati.���֤�� = Nvl(!���֤��)
        mPati.�ֻ��� = Nvl(!�ֻ���)
        mPati.�������� = Nvl(!��������)
    End With
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Sub

Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng�����ID As Long, strOutCardNo As String, strExpand
    Dim strOutPatiInforXml As String

    lng�����ID = IDKind.GetCurCard.�ӿ����
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
    'Call InitInterFacel(Me, mlngModule, lng�����ID, False, mobjCardObject)
    strExpand = lng�����ID
    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModule, lng�����ID, True, strExpand, strOutCardNo, strOutPatiInforXml) = False Then Exit Sub
    txtPatient.Text = strOutCardNo
    If txtPatient.Text <> "" Then Call txtPatient_KeyPress(vbKeyReturn)
    Exit Sub
 
End Sub

 
Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
     '����|ȫ��|������־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|
    '�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�)|��������(�ڼ�λ���ڼ�λ����,��Ϊ������)
    Set gobjSquare.objCurCard = objCard
    mlngҽ�ƿ����� = objCard.���ų���
    '105667:���ϴ���2017/5/23�����ż��ܵ��µ�һ������ƴ�����ܴ������뷨
    txtPatient.PasswordChar = ""
    '��Ҫ�����Ϣ,����ˢ����,���л�,���������ʾʧȥ����
    If txtPatient.Text <> "" And Not mblnNotClick Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txtPatient.IMEMode = 0
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    Dim lngPreIDKind As Long, intIndex As Integer
    Dim dtDate As Date
    Dim blnNew As Boolean

    If txtPatient.Locked Then Exit Sub  'Or Not Me.ActiveControl Is txtPatient Or txtPatient.Text <> ""
    mblnNotClick = True
    intIndex = IDKind.GetKindIndex(objCard.����)
    lngPreIDKind = IDKind.IDKind
    If intIndex > 0 Then IDKind.IDKind = intIndex

    txtPatient.Text = objPatiInfor.����
    Call txtPatient_KeyPress(vbKeyReturn)
    IDKind.IDKind = lngPreIDKind
    mblnNotClick = False
End Sub

Private Sub InitVsFlex()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��VSGrid�ؼ�
    '����:56599
    '����:2012-12-05 11:39:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strTmp As String
    Dim i As Integer
    On Error GoTo Errhand
    
    With vsBooks
        .Redraw = False
        .Cols = UBound(Split(C_BookInfoColumHeader, ";")) + 1
        For i = 0 To UBound(Split(C_BookInfoColumHeader, ";"))
            .FixedAlignment(i) = flexAlignCenterCenter
            .ColKey(i) = Split(Split(C_BookInfoColumHeader, ";")(i), ",")(0)
            .TextMatrix(0, i) = Split(Split(C_BookInfoColumHeader, ";")(i), ",")(0)
            .ColWidth(i) = Split(Split(C_BookInfoColumHeader, ";")(i), ",")(1)
            .ColAlignment(i) = Split(Split(C_BookInfoColumHeader, ";")(i), ",")(2)
           ' .ColAlignmentFixed(i) = 4
        Next
      '   .ColHidden(getColNum("��¼״̬")) = True
        .RowHeight(0) = 320
        .ExtendLastCol = True
        
        .ForeColorSel = C_COLOR_��ɫ
        .Redraw = True
    End With
    
    With vsPay
        .Redraw = False
        .Cols = UBound(Split(C_PayInfoColumHeader, ";")) + 1
        For i = 0 To UBound(Split(C_PayInfoColumHeader, ";"))
            .FixedAlignment(i) = flexAlignCenterCenter
            .ColKey(i) = Split(Split(C_PayInfoColumHeader, ";")(i), ",")(0)
            .TextMatrix(0, i) = Split(Split(C_PayInfoColumHeader, ";")(i), ",")(0)
            .ColWidth(i) = Split(Split(C_PayInfoColumHeader, ";")(i), ",")(1)
            .ColAlignment(i) = Split(Split(C_PayInfoColumHeader, ";")(i), ",")(2)
           ' .ColAlignmentFixed(i) = 4
        Next
        '����һ�����صĿ����ID
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "�����ID"
        .TextMatrix(0, .Cols - 1) = "�����ID"
        .ColHidden(.Cols - 1) = True
      '   .ColHidden(getColNum("��¼״̬")) = True
        .RowHeight(0) = 320
        .ExtendLastCol = True
        .ColDataType(0) = flexDTBoolean
        .Editable = flexEDKbd
        .ForeColorSel = C_COLOR_��ɫ
        .Redraw = True
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub LoadPatiBooks()
'����:���ز��˼�����Ϣ
    '���˼���
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    
    On Error GoTo errH
    
    If mPati.����ID = 0 Then Exit Sub
    strSQL = strSQL & _
    "   Select A.ID, '����' as ����,A.no as ���ݺ�,A.ʵ�ս��,A.���ʷ���,A.����ʱ��,A.����Ա����,A.ժҪ,A.���,A.��¼״̬, " & vbNewLine & _
    "          A.����ID,Decode(B.��¼���� , 11,'��Ԥ��',B.���㷽ʽ) as ���㷽ʽ,Sum(B.��Ԥ��) as ���, B.�����ID,B.����,B.����˵��,B.���㿨���,B.������ˮ��, " & vbNewLine & _
    "          Decode(B.��¼���� , 11,0,D.����) as ����,Nvl(C.�Ƿ�����,1) as �Ƿ�����,Nvl(C.�Ƿ�ȫ��,0) as �Ƿ�ȫ�� " & vbNewLine & _
    "   From סԺ���ü�¼ A,����Ԥ����¼ B, ҽ�ƿ���� C,���㷽ʽ D " & vbNewLine & _
    "   Where A.NO = B.NO(+) and A.��¼����  = B.��¼����(+) And B.�����ID = C.ID(+) And B.���㷽ʽ = D.����(+) And A.��¼����=5 And A.��¼״̬ = 1 and A.���ӱ�־=8 And A.����ID=[1]" & vbNewLine & _
    "   Group by A.ID,A.no,A.ʵ�ս��,A.���ʷ���,A.����ʱ��,A.����Ա����,A.ժҪ,A.���,A.��¼״̬,A.����ID,Decode(B.��¼���� , 11,'��Ԥ��',B.���㷽ʽ), " & vbNewLine & _
    "            B.�����ID , B.����, B.����˵��, B.���㿨���, B.������ˮ��, Decode(B.��¼����, 11, 0, D.����), Nvl(C.�Ƿ�����, 1), Nvl(C.�Ƿ�ȫ��, 0) " & vbNewLine & _
    "   Order by A.ID Desc,����"
'    "   Union All" & vbNewLine & _
'    "   Select A.ID, '�Һ�' as ����,A.no as ���ݺ�,A.ʵ�ս��,A.���ʷ���,A.����ʱ��,A.����Ա����,A.ժҪ,A.���,A.��¼״̬, " & _
'    "          B.����ID,Decode(B.��¼���� , 11,'��Ԥ��',B.���㷽ʽ) as ���㷽ʽ,B.��Ԥ�� as ���, B.�����ID,B.����,B.����˵��,B.�������,B.���㿨���,B.������ˮ��, " & _
'    "          Decode(B.��¼���� , 11,0,D.����) as ����,Nvl(C.�Ƿ�����,1) as �Ƿ�����,Nvl(C.�Ƿ�ȫ��,0) as �Ƿ�ȫ�� " & vbNewLine & _
'    "   From ������ü�¼ A,����Ԥ����¼ B, ҽ�ƿ���� C,���㷽ʽ D " & vbNewLine & _
'    "   Where A.����ID = B.����ID(+) And B.�����ID = C.ID(+) And B.���㷽ʽ = D.���� And A.��¼����=4 And A.��¼״̬ = 1 And A.���ӱ�־=1 And A.����ID=[1]" & vbNewLine & _
'    ") Order by ID Desc,����"
    Set mrsBooks = zlDatabase.OpenSQLRecord(strSQL, "���˲�����", mPati.����ID)

    With vsBooks
       .Rows = 1 'ȱʡ��ʾһ��
        If mrsBooks Is Nothing Then Exit Sub
        Do While Not mrsBooks.EOF
            '������ݺźͳ�����ͬ���Ƕ���֧����ʽ���㣬���ظ��Ǽ�
            If Nvl(mrsBooks!���ݺ�) <> .TextMatrix(i, .ColIndex("���ݺ�")) Or Nvl(mrsBooks!����) <> .TextMatrix(i, .ColIndex("����")) Then
                i = i + 1: .Rows = i + 1
                .RowData(i) = mrsBooks!id
                .TextMatrix(i, .ColIndex("���ݺ�")) = mrsBooks!���ݺ� & ""
                .TextMatrix(i, .ColIndex("����")) = mrsBooks!���� & ""
                .TextMatrix(i, .ColIndex("������")) = Format(mrsBooks!ʵ�ս�� & "", "0.00")
                .TextMatrix(i, .ColIndex("����ʱ��")) = Format(mrsBooks!����ʱ�� & "", "YYYY-MM-DD")
                .TextMatrix(i, .ColIndex("����Ա")) = mrsBooks!����Ա���� & ""
                .TextMatrix(i, .ColIndex("��ע")) = mrsBooks!ժҪ & ""
                .TextMatrix(i, .ColIndex("����")) = IIf(Val(mrsBooks!���ʷ��� & "") = 1, "��", "")
                '1-ʵ��;2-����;3-����
                .Cell(flexcpData, i, .ColIndex("����")) = IIf(Val(mrsBooks!���ʷ��� & "") = 1, 2, IIf(Val(mrsBooks!��¼״̬ & "") = 1, 1, 3))
            End If
            mrsBooks.MoveNext
        Loop
        If .Rows > 1 Then .Row = 1
    End With
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ClearPatiInfo()
    mPati.����ID = 0
    mPati.���� = ""
    mPati.���� = ""
    mPati.���� = ""
    mPati.�Ա� = ""
    mPati.���� = ""
    mPati.�������� = ""
    mPati.���� = ""
    mPati.����� = 0
    
    txtPatient.Text = ""
    txtNation.Text = ""
    txt�Ա�.Text = ""
    txt��������.Text = "____-__-__"
    txt����ʱ��.Text = "__:__"
    txt����.Text = ""
    txt���֤��.Text = ""
    txt�����.Text = ""
    txtӦ��.Text = "0.00"
    txtδ��.Text = "0.00"
    vsBooks.Rows = 1
    vsPay.Rows = 1
End Sub

Private Sub txtPatient_GotFocus()
    If Not txtPatient.Enabled Or txtPatient.Locked Then Exit Sub
    zlControl.TxtSelAll txtPatient
    If IsCardType(IDKind, "����") Then
        Call zlCommFun.OpenIme(True)
    End If
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim blnCard As Boolean, blnICCard As Boolean
    Dim strCardNo As String
    Dim blnPass As Boolean
    On Error GoTo errH
    If txtPatient.Locked Or txtPatient.Enabled = False Then Exit Sub
    If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub

    If IsCardType(IDKind, "����") Then
        '105567:���ϴ�,2017/5/25,���ż��ܵ��µ�һ������ƴ�����ܴ������뷨
        blnPass = txtPatient.PasswordChar <> ""
        blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, False)
        txtPatient.IMEMode = 0
        blnPass = txtPatient.PasswordChar = "" And blnPass
        If blnPass Then
            If txtPatient.SelLength = Len(txtPatient.Text) Then
                txtPatient.Text = ""
            End If
            SendKeys Chr(KeyAscii): KeyAscii = 0: Exit Sub
        End If
    ElseIf IsCardType(IDKind, "�����") Or IsCardType(IDKind, "סԺ��") Or IsCardType(IDKind, "�ֻ���") Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0: Exit Sub
        End If
    Else
        txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
        '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
        txtPatient.IMEMode = 0
    End If
    
    If Not (blnCard And Len(txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 _
        Or KeyAscii = 13 And Trim(txtPatient.Text) <> "") Then
        '����ˢ���ͻس�,���˳�
        Exit Sub
    End If

    If KeyAscii <> 13 Then
        txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
        txtPatient.SelStart = Len(txtPatient.Text)
    End If

    KeyAscii = 0
    strCardNo = Trim(txtPatient.Text)
    If Not GetPatient(txtPatient.Text, blnCard) Then
        Call ClearPatiInfo
        txtPatient.Text = strCardNo: zlControl.TxtSelAll txtPatient

        If InStr(1, "+*-", Left(txtPatient.Text & " ", 1)) > 0 Then
            KeyAscii = 0
            DoEvents
            If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
            zlControl.TxtSelAll txtPatient
            
            Exit Sub
        End If
        Exit Sub
    End If

    txtPatient.Text = mPati.����
    txtPatient.PasswordChar = ""
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txtPatient.IMEMode = 0

    Call LoadPatiBooks
    zlCommFun.PressKey vbKeyTab
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtPatient_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Function GetPatient(ByVal strInput As String, Optional ByVal blnCard As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ
    '���:blnCard=��ʾ�Ƿ���￨ˢ��
    '����:
    '����:���˶�ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-03 10:46:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim vRect As RECT, rsTemp As ADODB.Recordset
    Dim strSQL As String, strPati As String, strWhere As String, blnHavePass As Boolean
    Dim lng����ID As Long, blnCancel As Boolean, blnICCard As Boolean
    Dim strPassWord As String, bln�����ʻ� As Boolean, strErrMsg As String
    Dim strCardNo As String, lng�����ID As Long, blnIsMobileNO As Boolean
    
    txtPatient.ForeColor = &HFF0000
    strErrMsg = ""
    blnIsMobileNO = IDKind.IsMobileNO(strInput)
    If IsCardType(IDKind, "IC����") Then blnICCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
    If blnCard And Not (Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2))) Then    'ˢ����ȱʡ�Ŀ�
        
        If IDKind.Cards.��ȱʡ������ And Not IDKind.GetfaultCard Is Nothing Then
            lng�����ID = IDKind.GetfaultCard.�ӿ����
        ElseIf IDKind.GetCurCard.�ӿ���� > 0 Then
            lng�����ID = IDKind.GetCurCard.�ӿ����
        Else
            lng�����ID = -1
        End If
        '����|ȫ��|������־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�);��
        If GetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then
            If blnIsMobileNO Then
                '�ֻ��Ų���
                If GetPatiID("�ֻ���", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
            Else
                GoTo NotFoundPati:
            End If
        End If
        
        If lng����ID <= 0 Then GoTo NotFoundPati:
        strWhere = strWhere & " And A.����ID=[1]"
        strCardNo = strInput
        strInput = "-" & lng����ID
        blnHavePass = True
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then   '����ID
        strWhere = strWhere & " And A.����ID=[1]"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '�����
        strWhere = strWhere & " And A.�����=[1]"
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then 'סԺ��
        strWhere = strWhere & " And A.����ID = (Select Nvl(Max(����ID),0) As ����ID From ������ҳ Where סԺ�� = [1])"
    ElseIf IsCardType(IDKind, "����") And blnIsMobileNO Then
        '�ֻ��Ų���
        If GetPatiID("�ֻ���", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
        If lng����ID <= 0 Then GoTo NotFoundPati:
        strWhere = strWhere & " And A.����ID=[1]"
        strInput = "-" & lng����ID
    Else
        If mPati.����ID <> 0 Then
            If mPati.���� = strInput Then
                '74309:���ϴ���2014-7-7������������ʾ��ɫ����
                Call SetPatiColor(txtPatient, mPati.��������, txtPatient.ForeColor)
                GetPatient = True: Exit Function
            End If
        End If
        Select Case IDKind.GetCurCard.����
            Case "����", "��������￨"
                'ͨ������ģ�����Ҳ���(�������벡�˱�ʶʱ)
                strPati = _
                " Select A.����ID as ID,A.����ID,A.����,A.�Ա�,A.����," & _
                "        A.�����,A.סԺ��,A.��������,A.���֤��,A.�ֻ���" & _
                " From ������Ϣ A,���ű� B" & _
                " Where A.ͣ��ʱ�� is NULL And A.��ǰ����ID=B.ID(+) And Rownum <101 And A.���� Like [1]"
                strPati = strPati & " Order by  A.����"
                
                vRect = zlControl.GetControlRect(txtPatient.hWnd)
                Set rsTemp = zlDatabase.ShowSQLSelect(Me, strPati, 0, "����ѡ��", 1, "", "��ѡ����", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, False, True, strInput & "%", 101)
                If blnCancel Then GoTo NotFoundPati:
                If rsTemp Is Nothing Then GoTo NotFoundPati:
                If rsTemp.State <> 1 Then GoTo NotFoundPati:
                If rsTemp.RecordCount = 0 Then GoTo NotFoundPati:
                If Val(Nvl(rsTemp!����ID)) = 0 Then GoTo NotFoundPati:
                
                strInput = "-" & rsTemp!����ID
                strWhere = strWhere & " And A.����ID=[1]"
            Case "ҽ����"
                strInput = UCase(strInput)
                strWhere = strWhere & "  And A.ҽ����=[2]"
             Case "���֤��", "���֤", "�������֤"
                strInput = UCase(strInput)
                '�����:54197
                 If GetPatiID("���֤", strInput, False, lng����ID, strPassWord, strErrMsg, , , , False) = False Then lng����ID = 0
                 strInput = "-" & lng����ID
                strWhere = strWhere & " And A.����ID=[1]"
            Case "��ϵ�����֤��", "��ϵ�����֤" '�����:51071
                strInput = UCase(strInput)
                 If GetPatiID("��ϵ�����֤", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                 strInput = "-" & lng����ID
                strWhere = strWhere & " And A.����ID=[1]"
            Case "IC����", "IC��"
                strInput = UCase(strInput)
                If GetPatiID("IC��", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                strInput = "-" & lng����ID
                strWhere = strWhere & " And A.����ID=[1]"
            Case "�����"
                If Not IsNumeric(strInput) Then strInput = "0"
                strWhere = strWhere & " And A.�����=[2]"
            Case "סԺ��"
                If Not IsNumeric(strInput) Then strInput = "0"
                strWhere = strWhere & " And A.����ID = (Select Max(����id) From ������ҳ Where סԺ�� = [2])"
            Case Else
                '�������ĺ���
                If Val(IDKind.GetCurCard.�ӿ����) > 0 Then
                    lng�����ID = IDKind.GetCurCard.�ӿ����
                    bln�����ʻ� = IDKind.GetCurCard.�Ƿ�����ʻ�
                    If GetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                    If lng����ID = 0 Then GoTo NotFoundPati:
                    strCardNo = strInput
                    blnHavePass = True
                Else
                    If GetPatiID(IDKind.GetCurCard.����, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                End If
                If lng����ID <= 0 Then GoTo NotFoundPati:
                strWhere = strWhere & " And A.����ID=[1]"
                strInput = "-" & lng����ID
        End Select
    End If
    On Error GoTo errH
    strSQL = "Select A.����ID,A.����,A.�Ա�,A.����,A.��������,A.����,A.���֤��,A.�����,A.�ֻ���," & vbNewLine & _
    "         Nvl(Nvl(A.��������,B.��������),Decode(Nvl(A.����,B.����),Null,'��ͨ����','ҽ������')) �������� " & vbNewLine & _
    " From ������Ϣ A,������ҳ B " & vbNewLine & _
    " Where A.����ID=B.����ID(+) And A.��ҳID=B.��ҳID(+) And A.ͣ��ʱ�� is NULL " & strWhere
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "������Ϣ", Val(Mid(strInput, 2)), strInput)
    If rsTemp.EOF Then GoTo NotFoundPati:
    LoadPati (rsTemp!����ID)
    Call SetPatiColor(txtPatient, mPati.��������, txtPatient.ForeColor) '74309:���ϴ���2014-7-7������������ʾ��ɫ����
    GetPatient = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Call ClearPatiInfo
    Exit Function
NotFoundPati:
    If strErrMsg <> "" Then Exit Function
    Call ClearPatiInfo
    If blnCard Then
        MsgBox "����ȷ��������Ϣ�������Ƿ���ȷˢ����", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
    Else
        MsgBox "������Ϣδ�ҵ�,�����Ƿ�������ȷ!", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
    End If
End Function

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

Private Sub vsBooks_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim str���㷽ʽ As String
    Dim dblBackMoney As Double
    On Error GoTo Errhand
    With vsBooks
        If NewRow < 1 Then Exit Sub
        If OldRow = NewRow Then Exit Sub
        .Cell(flexcpBackColor, 1, 0, .Rows - 1, .Cols - 1) = C_COLOR_����
        .Cell(flexcpBackColor, NewRow, 0, NewRow, .Cols - 1) = C_COLOR_��ɫ
        
        If .Cell(flexcpData, .RowSel, .ColIndex("����")) = 1 Then
            dblBackMoney = Val(.TextMatrix(NewRow, .ColIndex("������")))
            txtӦ��.Text = Format(dblBackMoney, "0.00")
            txtδ��.Text = Format(dblBackMoney, "0.00")
        Else
            txtӦ��.Text = "0.00": txtδ��.Text = "0.00"
        End If
        
        '��λ��֧����ʽ
        vsPay.Clear 1: vsPay.Rows = 1
        mPayInfo.blnҽ�� = False: mPayInfo.blnȫ�� = False
        mrsBooks.Filter = " ���ݺ� = '" & .TextMatrix(NewRow, .ColIndex("���ݺ�")) & "' And ���� = '" & .TextMatrix(NewRow, .ColIndex("����")) & "'"
        If mrsBooks.RecordCount = 0 Then
            cmdOK.Enabled = False
        Else
            cmdOK.Enabled = True
            With vsPay
                Do While Not mrsBooks.EOF
                    If Not IsNull(mrsBooks!���㷽ʽ) Then
                        .Rows = .Rows + 1
                        .RowData(.Rows - 1) = IIf(Nvl(mrsBooks!����, 0) = 1, 0, Nvl(mrsBooks!�Ƿ�����, 0))
                        .TextMatrix(.Rows - 1, .ColIndex("֧����ʽ")) = mrsBooks!���㷽ʽ
                        .Cell(flexcpData, .Rows - 1, .ColIndex("֧����ʽ")) = Val(Nvl(mrsBooks!����))
                        .TextMatrix(.Rows - 1, .ColIndex("֧�����")) = Format(Nvl(mrsBooks!���), "0.00")
                        .TextMatrix(.Rows - 1, .ColIndex("����")) = Nvl(mrsBooks!����)
                        .TextMatrix(.Rows - 1, .ColIndex("������ˮ��")) = Nvl(mrsBooks!������ˮ��)
                        .TextMatrix(.Rows - 1, .ColIndex("����˵��")) = Nvl(mrsBooks!����˵��)
                        If mrsBooks!�Ƿ�ȫ�� = 1 Then mPayInfo.blnȫ�� = True
                        If Val(Nvl(mrsBooks!����)) = 3 Or Val(Nvl(mrsBooks!����)) = 4 Then mPayInfo.blnҽ�� = True
                        .TextMatrix(.Rows - 1, .ColIndex("�����ID")) = Nvl(mrsBooks!�����ID, Nvl(mrsBooks!���㿨���, 0))
                        .Cell(flexcpData, .Rows - 1, .ColIndex("�����ID")) = IIf(Val(Nvl(mrsBooks!���㿨���)) > 0, 1, 0)
                        If dblBackMoney > 0 Then
                            .Cell(flexcpChecked, .Rows - 1, .ColIndex("ѡ��")) = 1
                            dblBackMoney = dblBackMoney - Val(Nvl(mrsBooks!���))
                        End If
                    End If
                    mrsBooks.MoveNext
                Loop
            End With
            mrsBooks.MoveFirst
        End If
    End With
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Init�˷ѷ�ʽ()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������Ч��֧����ʽ
    '����:���˺�
    '����:2011-07-08 11:41:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim j As Long, strPayType As String, varData As Variant, varTemp As Variant, i As Long
    Dim rsTemp As ADODB.Recordset, blnFind As Boolean
    Dim strSQL As String, objCard As Card, objCards As Cards
    Dim lngKey As Long
    
    Set mobjPayCards = New Cards
    Set objCards = New Cards
    
    On Error GoTo Errhand
    strSQL = "" & _
    "   Select Nvl(A.ȱʡ��־,0) as ȱʡ,B.����,B.����,B.����,B.Ӧ����" & _
    "   From ���㷽ʽӦ�� A,���㷽ʽ B" & _
    "   Where A.���㷽ʽ=B.���� And A.Ӧ�ó���=[1]" & _
    "           And Nvl(B.����,1) IN(1,2)  " & _
    "   Order by B.����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "���￨")
    '��|ȫ��|ˢ����־|�����ID(���ѿ����)|����|�Ƿ����ѿ�|���㷽ʽ|�Ƿ�����|�Ƿ����ƿ�;��
    If Not gobjSquare Is Nothing Then
    ' zlGetCards(ByVal BytType As Byte)
        '���:bytType-  0-����ҽ�ƿ�;
    '                        1-���õ�ҽ�ƿ�,
    '                        2-���д��������˻���������
    '                        3-���õ������˻���ҽ�ƿ�
       Set objCards = gobjSquare.objSquareCard.zlGetCards(3)
    End If
    With rsTemp
        .Filter = 0
        If .RecordCount <> 0 Then .MoveFirst
        lngKey = 1
        Do While Not .EOF
            For i = 1 To objCards.count
                If objCards(i).���㷽ʽ = Nvl(rsTemp!����) Then
                    blnFind = True
                    Exit For
                End If
            Next
            If Not blnFind Then
                If Not (Val(Nvl(rsTemp!����)) = 3 Or Val(Nvl(rsTemp!����)) = 4 _
                    Or Val(Nvl(rsTemp!����)) = 7 Or Val(Nvl(rsTemp!����)) = 8 _
                    Or Val(Nvl(rsTemp!Ӧ����)) = 1) Then
                    
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
                    objCard.�Ƿ����� = True
                     
                    mobjPayCards.Add objCard, "K" & lngKey
                    lngKey = lngKey + 1
              End If
            End If
            .MoveNext
        Loop
    End With
    
    '��������
    For i = 1 To objCards.count
        rsTemp.Filter = "����='" & objCards(i).���㷽ʽ & "'" '���㷽ʽҪ������"���￨"Ӧ�ó��ϲ���ʹ��
        If Not rsTemp.EOF Then
            mobjPayCards.Add objCards(i), "K" & lngKey
            lngKey = lngKey + 1
        End If
    Next
    
    If mobjPayCards.count = 0 Then
        MsgBox "û�п��õĽ��㷽ʽ,���ȵ����㷽ʽ���������á�", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function IsCheckCancelValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����˷�ʱ��������Ч��
    '����:��Ч,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-12 18:43:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strName As String, bln���ѿ� As Boolean, lng�����ID As Long
    Dim lngRow As Long, objDelObject  As clsCardObject
    Dim dblMoney As Double, strErrMsg As String
   '����:48249
    Dim strSQL As String, rsBill As Recordset, rsTemp As ADODB.Recordset, lngCardBill As Long
    Dim intStyle As Integer, bln�˷� As Boolean
    
    On Error GoTo Errhand
    If mrsBooks Is Nothing Then
        strErrMsg = "û���ҵ���������Ϣ�������˷ѣ�"
    ElseIf mrsBooks.EOF Then
        strErrMsg = "û���ҵ���������Ϣ�������˷ѣ�"
    End If
    If strErrMsg <> "" Then
        MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If mPayInfo.blnҽ�� Then
        MsgBox "���ڵ���" & Nvl(mrsBooks!���ݺ�) & "ʹ����ҽ��֧����ʽ��ֻ��ͨ��������ҺŹ��������˷ѣ�", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    'Ҫ����Ƿ񵥶��յĲ�����
    mrsBooks.MoveFirst
    Do While Not mrsBooks.EOF
        If Val(Nvl(mrsBooks!�����ID)) > 0 Or Val(Nvl(mrsBooks!���㿨���)) > 0 Then
            If Nvl(mrsBooks!���, 0) <> 1 Then
                If mrsBooks!���� = "�Һ�" Then
                    MsgBox "ʹ������֧���Ĳ����Ѳ��ܵ����˷ѣ��뵽������ҺŹ���ͨ���˺Ź����˷ѣ�", vbInformation + vbOKOnly, gstrSysName
                Else
                    MsgBox "ʹ������֧���Ĳ����Ѳ��ܵ����˷ѣ��뵽��ҽ�ƿ����Ź���ͨ���˿������˷ѣ�", vbInformation + vbOKOnly, gstrSysName
                End If
                Exit Function
            End If
        End If
        mrsBooks.MoveNext
    Loop
    mrsBooks.MoveFirst
    
    intStyle = Val(zlDatabase.GetPara("�ѽ��ʵ��ݲ���", 100))
    strSQL = "Select B.NO From סԺ���ü�¼ a,���˽��ʼ�¼ b Where a.����id=b.id And a.��¼���� In (5,15) And a.��¼״̬=1 And b.��¼״̬=1 And a.no=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, vsBooks.TextMatrix(vsBooks.RowSel, vsBooks.ColIndex("���ݺ�")))
    If rsTemp.EOF Then bln�˷� = True
    Select Case intStyle
        Case 0
            bln�˷� = True
        Case 1
            If bln�˷� = False Then
                If MsgBox("����" & vsBooks.TextMatrix(vsBooks.RowSel, vsBooks.ColIndex("���ݺ�")) & "�������˴����Ƿ�����˷�", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    bln�˷� = True
                End If
            End If
        Case 2
            If bln�˷� = False Then
                MsgBox "����" & vsBooks.TextMatrix(vsBooks.RowSel, vsBooks.ColIndex("���ݺ�")) & "�������˴��������Ƚ����������˷�", vbInformation + vbOKOnly, gstrSysName
            End If
    End Select
    If bln�˷� = False Then Exit Function
    
    Set mobjDelObjects = New clsCardObjects
    With vsPay
        For lngRow = 1 To .Rows - 1
            If .Cell(flexcpChecked, lngRow, .ColIndex("ѡ��")) = 1 And .TextMatrix(lngRow, .ColIndex("�����ID")) > 0 Then
                bln���ѿ� = .Cell(flexcpData, lngRow, .ColIndex("�����ID")) = 1
                lng�����ID = .TextMatrix(lngRow, .ColIndex("�����ID"))
                
                If Val(Nvl(mrsBooks!���ʷ���)) = 1 Then IsCheckCancelValied = True: Exit Function
                If lng�����ID <= 0 Then IsCheckCancelValied = True: Exit Function
            
                '��Ϊ��,��Ҫ��ȡ��Ӧ��֧������
                Set objDelObject = zlGetClsCardObject(lng�����ID, bln���ѿ�)
            If objDelObject Is Nothing Then
                
                    MsgBox "��δ����ѡ����˷ѽӿ� ,�����ڴ˹���վ�Ͻ����˷�!", vbInformation + vbOKOnly, gstrSysName
                    Exit Function
                End If
                If Not objDelObject.CardPreporty.���� Then
                    MsgBox "��δ����" & mobjDelObject.CardPreporty.���� & "�ӿ� ,�����ڴ˹���վ�Ͻ����˷�!", vbInformation + vbOKOnly, gstrSysName
                    Exit Function
                End If
                If objDelObject.CardObject Is Nothing Then
                    If zlCreatePatiCardObject(objDelObject.CardPreporty, mobjDelObject.CardObject) = False Then
                        Exit Function
                    End If
                End If
                If Not objDelObject.InitCompents Then
                    If objDelObject.CardObject.zlInitComponents(Me, mlngModule, glngSys, gstrDBUser, gcnOracle, False, "") = False Then
                          Exit Function
                    End If
                    objDelObject.InitCompents = True
                End If
                
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
                Dim strXMLExpend As String, str���� As String
                str���� = .TextMatrix(lngRow, .ColIndex("����"))
                str������ˮ�� = .TextMatrix(lngRow, .ColIndex("������ˮ��"))
                str����˵�� = .TextMatrix(lngRow, .ColIndex("����˵��"))
                str������Ϣ = IIf(mrsBooks!���� = "�Һ�", 4, 5) & "|" & Nvl(mrsBooks!����ID)
                dblMoney = Val(.Cell(flexcpData, lngRow, .ColIndex("֧�����")))
                If objDelObject.CardObject.zlReturncheck(Me, mlngModule, objDelObject.CardPreporty.�ӿ����, str����, str������Ϣ, dblMoney, str������ˮ��, str����˵��, strXMLExpend) = False Then
                    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus: Exit Function
                    Exit Function
                End If
                
                '100610:���ϴ�,2016/10/13��Ԥ���˿������˿��Ƿ���֤ˢ��
                If objDelObject.CardPreporty.�Ƿ��˿��鿨 Then
                '   zlBrushCard(frmMain As Object, _
                    ByVal lngModule As Long, _
                    ByVal lngCardTypeID As Long, _
                    ByVal strPatiName As String, ByVal strSex As String, _
                    ByVal strOld As String, ByRef dbl��� As Double, _
                    Optional ByRef strCardNo As String, _
                    Optional ByRef strPassWord As String, _
                    Optional ByVal strXmlIn As String = "") As Boolean
                    '       strXmlIn-����������XML���,Ŀǰ��ʽ����:
                    '       <IN>
                    '           <CZLX>0</CZLX>    //��������,0-��������ˢ��,1-ת�˵���ˢ��,2-�˿����ˢ��
                    '       </IN>
                    Err = 0: On Error Resume Next
                    If objDelObject.CardObject.zlBrushCard(Me, mlngModule, objDelObject.CardPreporty.�ӿ����, _
                     mPati.����, mPati.�Ա�, mPati.����, dblMoney, _
                     str����, str����, "<IN><CZLX>2</CZLX></IN>") = False Then
                        If Err = 450 Then
                            Err = 0: On Error GoTo Errhand
                            If mobjDelObject.CardObject.zlBrushCard(Me, mlngModule, objDelObject.CardPreporty.�ӿ����, _
                             mPati.����, mPati.�Ա�, mPati.����, dblMoney, str����, str����) = False Then Exit Function
                        Else
                            Exit Function
                        End If
                    End If
                End If
    
                mobjDelObjects.Add objDelObject, False, lng�����ID, Nothing, bln���ѿ�, IIf(bln���ѿ�, "X", "K") & lng�����ID
            End If
        Next
    End With
    IsCheckCancelValied = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function IsNoCanc()
    '��鲡�����Ƿ��Ѿ����˷�
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strTable As String, strWhere As String, blnHave As Boolean
    Dim i As Integer, dblBalance As Double
    On Error GoTo Errhand
    If mrsBooks Is Nothing Then Exit Function
    If mrsBooks.EOF Then Exit Function
    
    With vsPay
        For i = 1 To .Rows - 1
            If .Cell(flexcpChecked, i, .ColIndex("ѡ��")) = 1 Then
                dblBalance = dblBalance + CDbl(.Cell(flexcpData, i, .ColIndex("֧�����")))
                blnHave = True
            End If
        Next
        If dblBalance <> CDbl(txtӦ��.Text) And blnHave Then
            MsgBox "�˷ѽ�һ�£�������ѡ���˷ѽ�", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End With
    If Not blnHave Then
        mrsBooks.MoveFirst
        Do While Not mrsBooks.EOF
            If mrsBooks!�Ƿ����� = 0 Then
                MsgBox "ԭ֧����ʽ��֧�����֣���ѡ���˷ѷ�ʽ��", vbInformation + vbOKOnly, gstrSysName
                Exit Function
            End If
            mrsBooks.MoveNext
        Loop
        mrsBooks.MoveFirst
    End If
    
    If Nvl(mrsBooks!����) = "�Һ�" Then
        strTable = "������ü�¼"
        strWhere = "��¼����=4 And ��¼״̬ = 1 And ���ӱ�־=1"
    Else
        strTable = "סԺ���ü�¼"
        strWhere = "��¼����=5 And ��¼״̬ = 1 and ���ӱ�־=8"
    End If
    
    strSQL = "Select 1 From " & strTable & " Where " & strWhere & " And ����ID=[1] And no=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�����Ѽ��", mPati.����ID, Nvl(mrsBooks!���ݺ�))
    If rsTemp.RecordCount = 0 Then
        MsgBox "��ǰ�������ѱ�������Ա�˷�!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    IsNoCanc = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub vsPay_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Dim dblMoney As Double, i As Long
    
    With vsPay
        If Row < 1 Or Col <> .ColIndex("ѡ��") Then Exit Sub
        If mPayInfo.blnҽ�� Then
            .Cell(flexcpChecked, Row, .ColIndex("ѡ��")) = 2
            MsgBox "���ڵ���" & vsBooks.TextMatrix(vsBooks.RowSel, .ColIndex("���ݺ�")) & "ʹ����ҽ��֧����ʽ��ֻ��ͨ��������ҺŹ��������˷ѣ�", vbInformation + vbOKOnly, gstrSysName
            Exit Sub
        End If
        'ȡ��ѡ��
        If .Cell(flexcpChecked, Row, .ColIndex("ѡ��")) = 2 Then
            .Cell(flexcpChecked, Row, .ColIndex("֧�����")) = 0
            For i = 1 To .Rows - 1
                If .Cell(flexcpChecked, i, .ColIndex("ѡ��")) = 1 Then dblMoney = dblMoney + CDbl(.TextMatrix(i, .ColIndex("֧�����")))
            Next
            txtδ��.Text = Format(CDbl(txtӦ��.Text) - dblMoney, "0.00")
            Exit Sub
        End If
        '����Ѿ��������˿��������ѡ��
        If Val(txtδ��.Text) = 0 Then .Cell(flexcpChecked, Row, .ColIndex("ѡ��")) = 2: Exit Sub
        dblMoney = Val(.TextMatrix(Row, .ColIndex("֧�����")))
        '���ѡ��Ľ���������˿ȡ������ѡ��
        If dblMoney >= CDbl(txtӦ��.Text) Then
            .Cell(flexcpChecked, 1, .ColIndex("ѡ��"), .Rows - 1, .ColIndex("ѡ��")) = 2
            .Cell(flexcpData, Row, .ColIndex("֧�����"), .Rows - 1, .ColIndex("֧�����")) = 0
            .Cell(flexcpChecked, Row, .ColIndex("ѡ��")) = 1
            .Cell(flexcpData, Row, .ColIndex("֧�����")) = CDbl(txtӦ��.Text)
            txtδ��.Text = "0.00"
        Else
            '��δ�˽��
            dblMoney = CDbl(txtδ��.Text) - dblMoney
            If dblMoney <= 0 Then
                .Cell(flexcpData, Row, .ColIndex("֧�����")) = Val(txtδ��.Text)
                txtδ��.Text = "0.00"
            Else
                .Cell(flexcpData, Row, .ColIndex("֧�����")) = Val(txtδ��.Text) - dblMoney
                txtδ��.Text = Format(dblMoney, "0.00")
            End If
        End If
    End With
End Sub

Private Sub vsPay_EnterCell()
    With vsPay
        If .RowSel < 1 Then Exit Sub
        .Cell(flexcpBackColor, 1, 0, .Rows - 1, .Cols - 1) = C_COLOR_����
        .Cell(flexcpBackColor, .RowSel, 0, .RowSel, .Cols - 1) = C_COLOR_��ɫ
    End With
End Sub

Private Sub vsPay_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> vsPay.ColIndex("ѡ��") Then Cancel = True
End Sub

Private Function SaveData() As Boolean
    Dim strSQL As String, rsTemp As ADODB.Recordset, str���㷽ʽ As String
    Dim blnTrans As Boolean, blnHave As Boolean, blnOraclTrans As Boolean
    Dim i As Integer, strBalance As String, dblDeposit As Double
    On Error GoTo Errhand
    If mrsBooks Is Nothing Then Exit Function
    If mrsBooks.EOF Then Exit Function
    
    If Nvl(mrsBooks!����) = "�Һ�" Then
        With vsPay
            For i = 1 To .Rows - 1
                If .Cell(flexcpChecked, i, .ColIndex("ѡ��")) = 1 Then
                    If .Cell(flexcpData, i, .ColIndex("֧����ʽ")) = 7 Or .Cell(flexcpData, i, .ColIndex("֧����ʽ")) = 8 Then
                        strBalance = strBalance & "|" & .TextMatrix(i, .ColIndex("֧����ʽ")) & "," & .Cell(flexcpData, i, .ColIndex("֧�����")) & "," & "1"
                    ElseIf .Cell(flexcpData, i, .ColIndex("֧����ʽ")) = 0 Then
                    '��Ԥ��
                    dblDeposit = CDbl(.Cell(flexcpData, i, .ColIndex("֧�����")))
                    Else
                        strBalance = strBalance & "|" & .TextMatrix(i, .ColIndex("֧����ʽ")) & "," & .Cell(flexcpData, i, .ColIndex("֧�����")) & "," & "0"
                    End If
                    blnHave = True
                End If
            Next
        End With
        'û��ѡ֧����ʽ�������ֽ���
        If blnHave = False Then
            strSQL = "Select ���� from ���㷽ʽ Where ���� = 1"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
            If rsTemp.RecordCount > 0 Then
                strBalance = Nvl(rsTemp!����) & "," & Val(txtӦ��.Text) & "," & "0"
            Else
                strBalance = "�ֽ�," & Val(txtӦ��.Text) & "," & "0"
            End If
        End If
        'zl_���˹Һż�¼_Delete
        strSQL = "zl_���˹Һż�¼_����_DELETE("
        '  ���ݺ�_In       ������ü�¼.No%Type,
        strSQL = strSQL & "'" & Nvl(mrsBooks!���ݺ�) & "',"
        '  ����Ա���_In   ������ü�¼.����Ա���%Type,
        strSQL = strSQL & "'" & UserInfo.��� & "',"
        '  ����Ա����_In   ������ü�¼.����Ա����%Type,
        strSQL = strSQL & "'" & UserInfo.���� & "',"
        '  ժҪ_In         ������ü�¼.ժҪ%Type := Null, --ԤԼȡ��ʱ ��д ���ԤԼȡ��ԭ��
        strSQL = strSQL & " NULL ,"
        '  ɾ�������_In   Number := 0,
        strSQL = strSQL & "" & 0 & ","
        '  ��ԭ���˽���_In Varchar2 := Null,
        strSQL = strSQL & "NULL" & ","
        '  �˷�����_In     In Number := 0, --0-ȫ�� 1-�˹Һŷ� 2-�˲�����
        strSQL = strSQL & "" & 2 & ","
        '  ��ָ������_In   ����Ԥ����¼.���㷽ʽ%Type := Null
        strSQL = strSQL & "NULL" & ","
        '  �˺�����_In   Number := 1
        strSQL = strSQL & 1 & ",'"
        '  ���㷽ʽ_In   Varchar2 := Null
        strSQL = strSQL & strBalance & "',"
        '   ��Ԥ��_In       ����Ԥ����¼.��Ԥ��%Type
        strSQL = strSQL & dblDeposit & ")"
    Else
        With vsPay
            For i = 1 To .Rows - 1
                If .Cell(flexcpChecked, i, .ColIndex("ѡ��")) = 1 Then
                    str���㷽ʽ = .TextMatrix(i, .ColIndex("֧����ʽ"))
                    blnHave = True: Exit For
                End If
            Next
        End With
        'û��ѡ֧����ʽ�������ֽ���
        If blnHave = False Then
            strSQL = "Select ���� from ���㷽ʽ Where ���� = 1"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
            If rsTemp.RecordCount > 0 Then
                str���㷽ʽ = Nvl(rsTemp!����)
            Else
                str���㷽ʽ = "�ֽ�"
            End If
        End If
        strSQL = "zl_ҽ�ƿ���¼_DELETE('" & Nvl(mrsBooks!���ݺ�) & "','" & UserInfo.��� & "','" & UserInfo.���� & "',2,'" & str���㷽ʽ & "')"
    End If
    gcnOracle.BeginTrans: blnTrans = True
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    If CallBackBalanceInterface(Nvl(mrsBooks!���ݺ�), blnOraclTrans) = False Then
        If blnOraclTrans = False Then gcnOracle.RollbackTrans
        Exit Function
    End If
    If blnOraclTrans = False Then gcnOracle.CommitTrans
    blnTrans = False
    SaveData = True
    Exit Function
Errhand:
    If blnTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CallBackBalanceInterface(ByVal strNO As String, ByRef blnTrancs As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���û��˽ӿ�
    '���:blnTrancs-�Ƿ���������
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-13 10:33:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str������Ϣ As String, str���� As String, strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim varData As Variant, varTemp As Variant, i As Long, lng����ID As Long, cllPro As Collection, cllProAfter As Collection
    Dim bln���ѿ� As Boolean, lng�����ID As Long, strSQL As String, rsTemp As ADODB.Recordset
    Dim str������Ϣ As String, strTemp As String, dblMoney As Double, blnThree As Boolean
    Dim objDelObject  As clsCardObject, lngRow As Long
    On Error GoTo errHandle
    blnTrancs = False
    
    If Val(Nvl(mrsBooks!���ʷ���)) = 1 Then CallBackBalanceInterface = True: Exit Function
    Set cllPro = New Collection: Set cllProAfter = New Collection
    With vsPay
        For lngRow = 1 To .Rows - 1
            If .Cell(flexcpChecked, lngRow, .ColIndex("ѡ��")) = 1 And .TextMatrix(lngRow, .ColIndex("�����ID")) > 0 Then
                bln���ѿ� = .Cell(flexcpData, lngRow, .ColIndex("�����ID")) = 1
                lng�����ID = .TextMatrix(lngRow, .ColIndex("�����ID"))
                
                If lng����ID = 0 Then
                    If .TextMatrix(.Row, .Col) = "�Һ�" Then
                        strSQL = "Select ����ID,���ʷ��� From ������ü�¼  Where ��¼����=4 and NO=[1] and ��¼״̬=2"
                    Else
                        strSQL = "Select ����ID,���ʷ��� From סԺ���ü�¼  Where ��¼����=5 and NO=[1] and ��¼״̬=2"
                    End If
                    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
                    If rsTemp.EOF Then
                        gcnOracle.RollbackTrans: blnTrancs = True
                        MsgBox "δ�ҵ������ѵ��˷���Ϣ��Ϣ�����ܼ���", vbInformation + vbOKOnly, gstrSysName
                        Exit Function
                    End If
                    lng����ID = Val(Nvl(rsTemp!����ID))
                End If
                If .TextMatrix(.Row, .Col) = "�Һ�" Then
                    strSwapExtendInfor = "4|" & lng����ID: strTemp = strSwapExtendInfor
                Else
                    strSwapExtendInfor = "5|" & lng����ID: strTemp = strSwapExtendInfor
                End If
                
                '�˷ѽӿ�
                Set objDelObject = mobjDelObjects(IIf(bln���ѿ�, "X", "K") & lng�����ID)
                If Not objDelObject.InitCompents Then
                    If objDelObject.CardObject.zlInitComponents(Me, mlngModule, glngSys, gstrDBUser, gcnOracle, False, "") = False Then
                          Exit Function
                    End If
                    objDelObject.InitCompents = True
                End If
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
                '       strSwapNo-������ˮ��(�տ�ʱ�Ľ�����ˮ��)
                '       strSwapMemo-����˵��(�տ�ʱ�Ľ���˵��)
                '       strSwapExtendInfor-���׵���չ��Ϣ
                '           ��ʽΪ:��Ŀ����1|��Ŀ����2||��||��Ŀ����n|��Ŀ����n ÿ����Ŀ�в��ܰ���|�ַ�
                str���� = .TextMatrix(lngRow, .ColIndex("����"))
                strSwapGlideNO = .TextMatrix(lngRow, .ColIndex("������ˮ��"))
                strSwapMemo = .TextMatrix(lngRow, .ColIndex("����˵��"))
                str������Ϣ = IIf(mrsBooks!���� = "�Һ�", 4, 5) & "|" & Nvl(mrsBooks!����ID)
                dblMoney = Val(.Cell(flexcpData, lngRow, .ColIndex("֧�����")))
                If objDelObject.CardObject.zlReturnMoney(Me, mlngModule, lng�����ID, str����, str������Ϣ, dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor) = False Then
                    Exit Function
                End If
                
                '���½�����Ϣ
                '    Zl_�����ӿڸ���_Update
                strSQL = "Zl_�����ӿڸ���_Update("
                '  �����id_In   ����Ԥ����¼.�����id%Type,
                strSQL = strSQL & "" & lng�����ID & ","
                '  ���ѿ�_In     Number,
                strSQL = strSQL & "" & IIf(bln���ѿ�, 1, 0) & ","
                '  ����_In       ����Ԥ����¼.����%Type,
                strSQL = strSQL & "'" & str���� & "',"
                '  ����ids_In    Varchar2,
                strSQL = strSQL & "'" & lng����ID & "',"
                '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type,
                strSQL = strSQL & "'" & strSwapGlideNO & "',"
                '  ����˵��_In   ����Ԥ����¼.����˵��%Type
                strSQL = strSQL & "'" & strSwapMemo & "')"
                zlAddArray cllPro, strSQL
                
                If strTemp <> strSwapExtendInfor Then
                    'strSwapExtendInfor:������չ��Ϣ,��ʽ:��Ŀ����|��Ŀ����||...
                    varData = Split(strSwapExtendInfor, "||")
                    Set cllPro = New Collection
                    For i = 0 To UBound(varData)
                        If Trim(varData(i)) <> "" Then
                            varTemp = Split(varData(i) & "|", "|")
                            If varTemp(0) <> "" Then
                                strTemp = varTemp(0) & "|" & varTemp(1)
                                If zlCommFun.ActualLen(str������Ϣ & "||" & strTemp) > 2000 Then
                                    str������Ϣ = Mid(str������Ϣ, 3)
                                    'Zl_�������㽻��_Insert
                                    strSQL = "Zl_�������㽻��_Insert("
                                    '�����id_In ����Ԥ����¼.�����id%Type,
                                    strSQL = strSQL & "" & lng�����ID & ","
                                    '���ѿ�_In   Number,
                                    strSQL = strSQL & "" & IIf(bln���ѿ�, 1, 0) & ","
                                    '����_In     ����Ԥ����¼.����%Type,
                                    strSQL = strSQL & "'" & str���� & "',"
                                    '����ids_In  Varchar2,
                                    strSQL = strSQL & "'" & lng����ID & "',"
                                    '������Ϣ_In Varchar2:������Ŀ|��������||...
                                    strSQL = strSQL & "'" & str������Ϣ & "')"
                                    zlAddArray cllProAfter, strSQL
                                    str������Ϣ = ""
                                End If
                                str������Ϣ = str������Ϣ & "||" & strTemp
                            End If
                        End If
                    Next
                    If str������Ϣ <> "" Then
                        str������Ϣ = Mid(str������Ϣ, 3)
                        'Zl_�������㽻��_Insert
                        strSQL = "Zl_�������㽻��_Insert("
                        '�����id_In ����Ԥ����¼.�����id%Type,
                        strSQL = strSQL & "" & lng�����ID & ","
                        '���ѿ�_In   Number,
                        strSQL = strSQL & "" & IIf(bln���ѿ�, 1, 0) & ","
                        '����_In     ����Ԥ����¼.����%Type,
                        strSQL = strSQL & "'" & str���� & "',"
                        '����ids_In  Varchar2,
                        strSQL = strSQL & "'" & lng����ID & "',"
                        '������Ϣ_In Varchar2:������Ŀ|��������||...
                        strSQL = strSQL & "'" & str������Ϣ & "')"
                        zlAddArray cllProAfter, strSQL
                    End If
                End If
            End If
        Next
    End With
    
    '���½�����Ϣ,���ύ,�����������,�ٸ�����صĽ�����Ϣ
    zlExecuteProcedureArrAy cllPro, Me.Caption, , True

    Err = 0: On Error GoTo ErrOthers:
    zlExecuteProcedureArrAy cllProAfter, Me.Caption

    CallBackBalanceInterface = True: blnTrancs = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans: blnTrancs = True
    Call ErrCenter
    Exit Function
ErrOthers:
    '��չ��Ϣ,������һ����,�Ա��֤
    If ErrCenter() = 1 Then
        gcnOracle.RollbackTrans
        Resume
    End If
    CallBackBalanceInterface = True
    gcnOracle.CommitTrans: blnTrancs = True
End Function

'�ؼ������Ƿ�ƥ��
Private Function IsCardType(ByVal IDKindCtl As IDKindNew, ByVal strCardName As String) As Boolean
    If IDKindCtl Is Nothing Then Exit Function
    If UCase(TypeName(IDKindCtl)) <> "IDKINDNEW" Then Exit Function
    Select Case strCardName
     Case "����", "��������￨"
          IsCardType = IDKindCtl.GetCurCard.���� Like "����*"
     Case "���֤", "���֤��", "�������֤"
          IsCardType = IDKindCtl.GetCurCard.���� Like "*���֤*"
     Case "IC����", "IC��"
          IsCardType = IDKindCtl.GetCurCard.���� Like "IC��*"
     Case "ҽ����"
          IsCardType = IDKindCtl.GetCurCard.���� = "ҽ����"
     Case "�����"
          IsCardType = IDKindCtl.GetCurCard.���� = "�����"
     Case "סԺ��"
          IsCardType = IDKindCtl.GetCurCard.���� = "סԺ��"
     Case "�ֻ���"
          IsCardType = IDKindCtl.GetCurCard.���� = "�ֻ���"
     Case Else
            If IDKindCtl.GetCurCard Is Nothing Then Exit Function
            If Not IsNumeric(strCardName) Or Val(strCardName) <= 0 Then
                  IsCardType = strCardName = IDKindCtl.GetCurCard.����
            Else
                If IDKindCtl.GetCurCard.�ӿ���� <= 0 Then Exit Function
                IsCardType = IDKindCtl.GetCurCard.�ӿ���� = Val(strCardName)
            End If
     End Select
End Function
