VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJOCK.DOCKINGPANE.UNICODE.9600.OCX"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#7.0#0"; "zlIDKind.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPatiBalanceSplit 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000D&
   Caption         =   "���˽��ʵ�(�������)"
   ClientHeight    =   10290
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14985
   Icon            =   "frmPatiBalanceSplit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10290
   ScaleWidth      =   14985
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picFormat 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   3825
      ScaleHeight     =   315
      ScaleWidth      =   2505
      TabIndex        =   105
      Top             =   210
      Visible         =   0   'False
      Width           =   2535
      Begin VB.Label lblFormat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ʊ�ݸ�ʽ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   240
         Left            =   15
         TabIndex        =   106
         Top             =   30
         Visible         =   0   'False
         Width           =   945
      End
   End
   Begin VB.PictureBox picBalanceBack 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   6795
      Left            =   5235
      ScaleHeight     =   6795
      ScaleWidth      =   7245
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   2910
      Width           =   7245
      Begin VB.CommandButton cmdDelBalance 
         Caption         =   "��������(&D)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5610
         TabIndex        =   125
         Top             =   4110
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.PictureBox picNotPayment 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   60
         ScaleHeight     =   1305
         ScaleWidth      =   2475
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   1305
         Width           =   2505
         Begin XtremeSuiteControls.ShortcutCaption stcTittile 
            Height          =   450
            Left            =   15
            TabIndex        =   75
            Top             =   15
            Width           =   3330
            _Version        =   589884
            _ExtentX        =   5874
            _ExtentY        =   794
            _StockProps     =   6
            Caption         =   "��ǰδ��"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   15.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
         End
         Begin VB.Label lblʣ���Ը� 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "123791.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   26.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   615
            Left            =   15
            TabIndex        =   76
            Top             =   585
            Width           =   2430
         End
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ������(&C)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5610
         TabIndex        =   64
         Top             =   4125
         Width           =   1515
      End
      Begin VB.PictureBox picBalanceInfor 
         AutoRedraw      =   -1  'True
         Height          =   1050
         Left            =   60
         ScaleHeight     =   990
         ScaleWidth      =   6900
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   90
         Width           =   6960
         Begin zl9InExse.txtEdit txtBalance 
            Height          =   360
            Index           =   1
            Left            =   4515
            TabIndex        =   45
            Tag             =   "���ν���"
            Top             =   120
            Width           =   1590
            _extentx        =   2805
            _extenty        =   635
            font            =   "frmPatiBalanceSplit.frx":058A
            inputmode       =   4
            text            =   ""
         End
         Begin zl9InExse.txtEdit txtBalance 
            Height          =   360
            Index           =   2
            Left            =   1110
            TabIndex        =   47
            Tag             =   "����˵��"
            Top             =   555
            Width           =   4920
            _extentx        =   8678
            _extenty        =   635
            font            =   "frmPatiBalanceSplit.frx":05AE
            text            =   ""
         End
         Begin zl9InExse.txtEdit txtBalance 
            Height          =   360
            Index           =   0
            Left            =   1125
            TabIndex        =   43
            Tag             =   "����δ��"
            Top             =   105
            Width           =   1920
            _extentx        =   3387
            _extenty        =   635
            backcolor       =   -2147483633
            font            =   "frmPatiBalanceSplit.frx":05D2
            appearance      =   0
            text            =   ""
         End
         Begin VB.Label lblBalance 
            AutoSize        =   -1  'True
            Caption         =   "����δ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   60
            TabIndex        =   42
            Top             =   180
            Width           =   1020
         End
         Begin VB.Label lblBalance 
            AutoSize        =   -1  'True
            Caption         =   "���ν���"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   3465
            TabIndex        =   44
            Top             =   180
            Width           =   1020
         End
         Begin VB.Label lblBalance 
            AutoSize        =   -1  'True
            Caption         =   "����˵��"
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
            Index           =   2
            Left            =   90
            TabIndex        =   46
            Top             =   615
            Width           =   960
         End
      End
      Begin VB.PictureBox picCurBalance 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   2055
         Left            =   2640
         ScaleHeight     =   2055
         ScaleWidth      =   4170
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   1890
         Width           =   4170
         Begin zlIDKind.IDKindNew IDKind�Ҳ� 
            Height          =   375
            Left            =   15
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   615
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            ShowSortName    =   0   'False
            IDKindStr       =   "��|�Ҳ�|0|0|0|0|0|;��|����Ԥ��|0|0|0|0|0|;ס|סԺԤ��|0|0|0|0|0|"
            CaptionAlignment=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontBold        =   -1  'True
            FontSize        =   12
            FontName        =   "����"
            IDKind          =   -1
            DefaultCardType =   "0"
            NotAutoAppendKind=   -1  'True
            BackColor       =   -2147483633
         End
         Begin zl9InExse.txtEdit txtBalance 
            Height          =   360
            Index           =   5
            Left            =   1350
            TabIndex        =   54
            Tag             =   "�Ҳ�"
            Top             =   615
            Width           =   2835
            _extentx        =   5001
            _extenty        =   635
            font            =   "frmPatiBalanceSplit.frx":05F6
            inputmode       =   2
            text            =   ""
         End
         Begin zl9InExse.txtEdit txtBalance 
            Height          =   360
            Index           =   7
            Left            =   1350
            TabIndex        =   60
            Tag             =   "ժҪ"
            Top             =   1485
            Width           =   2835
            _extentx        =   5001
            _extenty        =   635
            font            =   "frmPatiBalanceSplit.frx":061A
            text            =   ""
         End
         Begin zl9InExse.txtEdit txtBalance 
            Height          =   360
            Index           =   6
            Left            =   1350
            TabIndex        =   56
            Tag             =   "�������"
            Top             =   1050
            Width           =   2835
            _extentx        =   5001
            _extenty        =   635
            font            =   "frmPatiBalanceSplit.frx":063E
            text            =   ""
         End
         Begin zl9InExse.txtEdit txtBalance 
            Height          =   360
            Index           =   4
            Left            =   2475
            TabIndex        =   52
            Tag             =   "�ɿ�"
            Top             =   120
            Width           =   1695
            _extentx        =   2990
            _extenty        =   635
            font            =   "frmPatiBalanceSplit.frx":0662
            maxlength       =   10
            inputmode       =   2
            text            =   "0.00"
         End
         Begin zlIDKind.IDKindNew IDKindPaymentsType 
            Height          =   345
            Left            =   1350
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   135
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   609
            ShowSortName    =   0   'False
            Appearance      =   2
            IDKindStr       =   "��|�ֽ�|0|0|0|0|0|0;֧|֧Ʊ|0|0|0|0|0|"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontBold        =   -1  'True
            FontSize        =   12
            FontName        =   "����"
            IDKind          =   -1
            DefaultCardType =   "0"
            AllowAutoCommCard=   0   'False
            BackColor       =   -2147483633
         End
         Begin VB.Label lblBalance 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��    ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   4
            Left            =   255
            TabIndex        =   50
            Top             =   180
            Width           =   1050
         End
         Begin VB.Label lblBalance 
            AutoSize        =   -1  'True
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
            Height          =   240
            Index           =   5
            Left            =   255
            TabIndex        =   55
            Top             =   1110
            Width           =   1020
         End
         Begin VB.Label lblBalance 
            AutoSize        =   -1  'True
            Caption         =   "ժ    Ҫ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   6
            Left            =   255
            TabIndex        =   58
            Top             =   1545
            Width           =   1050
         End
      End
      Begin VB.PictureBox picOwerFee 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         ForeColor       =   &H80000008&
         Height          =   1275
         Left            =   60
         ScaleHeight     =   1245
         ScaleWidth      =   2475
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   2760
         Width           =   2505
         Begin XtremeSuiteControls.ShortcutCaption stcTittleTotal 
            Height          =   420
            Left            =   15
            TabIndex        =   73
            Top             =   30
            Width           =   3330
            _Version        =   589884
            _ExtentX        =   5874
            _ExtentY        =   741
            _StockProps     =   6
            Caption         =   "�Ը��ϼ�"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   15.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
         End
         Begin VB.Label lbl�Ը��ϼ� 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   26.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   615
            Left            =   1440
            TabIndex        =   72
            Top             =   555
            Width           =   1005
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsBlance 
         Height          =   1770
         Left            =   75
         TabIndex        =   65
         Top             =   4560
         Width           =   7110
         _cx             =   12541
         _cy             =   3122
         Appearance      =   2
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
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483634
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   7
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmPatiBalanceSplit.frx":0686
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
         Begin VB.Image imgDel 
            Height          =   240
            Left            =   75
            Picture         =   "frmPatiBalanceSplit.frx":079C
            Top             =   45
            Visible         =   0   'False
            Width           =   240
         End
      End
      Begin VB.PictureBox picCurDeposit 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   570
         Left            =   2655
         ScaleHeight     =   570
         ScaleWidth      =   3885
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   1305
         Width           =   3885
         Begin zl9InExse.txtEdit txtBalance 
            Height          =   360
            Index           =   3
            Left            =   1275
            TabIndex        =   49
            Tag             =   "��Ԥ��"
            Top             =   90
            Width           =   1635
            _extentx        =   2884
            _extenty        =   635
            font            =   "frmPatiBalanceSplit.frx":0D26
            inputmode       =   2
            text            =   ""
         End
         Begin VB.CheckBox chkDeposit 
            Caption         =   "��Ԥ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   105
            TabIndex        =   93
            Top             =   150
            Visible         =   0   'False
            Width           =   1110
         End
         Begin VB.Label lblBalance 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "�� Ԥ ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   3
            Left            =   180
            TabIndex        =   48
            Top             =   150
            Width           =   1035
         End
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "��ɽ���(&O)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4035
         TabIndex        =   63
         Top             =   4140
         Width           =   1515
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "��������(&N)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2445
         TabIndex        =   62
         Top             =   4125
         Width           =   1515
      End
      Begin VB.CommandButton cmdYBBalance 
         Caption         =   "ҽ������(&Y)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4020
         TabIndex        =   61
         Top             =   4125
         Width           =   1515
      End
      Begin MSMask.MaskEdBox txtDate 
         Height          =   360
         Left            =   4395
         TabIndex        =   96
         TabStop         =   0   'False
         Top             =   6375
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483633
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
      Begin VB.Shape shpBalance 
         BackColor       =   &H8000000D&
         BorderColor     =   &H8000000D&
         BorderWidth     =   5
         Height          =   6555
         Left            =   30
         Top             =   -90
         Visible         =   0   'False
         Width           =   9255
      End
      Begin VB.Label lblPrevious 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ϴ��Է�9999.99"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   135
         TabIndex        =   107
         Top             =   4110
         Width           =   1740
      End
      Begin VB.Label lblBalance 
         AutoSize        =   -1  'True
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
         Index           =   7
         Left            =   3360
         TabIndex        =   97
         Top             =   6435
         Width           =   960
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���:0.0"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   1845
         TabIndex        =   91
         Top             =   4230
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label lblҽ������ 
         AutoSize        =   -1  'True
         Caption         =   "ͳ��֧��:"
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
         Left            =   75
         TabIndex        =   83
         Top             =   6435
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label lblCurPaymentInfor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ǰ֧�����"
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
         Left            =   150
         TabIndex        =   79
         Top             =   4215
         Width           =   1440
      End
      Begin VB.Shape shpFun 
         BackColor       =   &H80000000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000000&
         Height          =   585
         Left            =   75
         Top             =   4080
         Width           =   8190
      End
   End
   Begin VB.PictureBox picPati 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2205
      Left            =   0
      ScaleHeight     =   2205
      ScaleWidth      =   14985
      TabIndex        =   87
      TabStop         =   0   'False
      Top             =   765
      Width           =   14985
      Begin VB.ComboBox cboDiag 
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
         Left            =   8820
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1140
         Width           =   1755
      End
      Begin VB.Frame Frame2 
         Height          =   30
         Left            =   -30
         TabIndex        =   104
         Top             =   1590
         Width           =   30000
      End
      Begin VB.Frame Frame1 
         Height          =   30
         Left            =   -30
         TabIndex        =   103
         Top             =   1035
         Width           =   30000
      End
      Begin VB.CommandButton cmdYB 
         Caption         =   "��֤(&Y)"
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
         Left            =   2520
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "ҽ�����������֤,�ȼ�F6"
         Top             =   60
         Visible         =   0   'False
         Width           =   1100
      End
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
         Height          =   360
         Left            =   1230
         TabIndex        =   2
         Top             =   60
         Width           =   2205
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "ˢ��δ�����(&R)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   12600
         TabIndex        =   17
         Top             =   1650
         Width           =   2115
      End
      Begin VB.OptionButton opt��; 
         Appearance      =   0  'Flat
         Caption         =   "��;����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   9495
         TabIndex        =   18
         Top             =   1755
         Width           =   1275
      End
      Begin VB.Frame fraSplitPati 
         Height          =   30
         Left            =   -195
         TabIndex        =   88
         Top             =   495
         Width           =   30000
      End
      Begin MSMask.MaskEdBox txtPatiEnd 
         Height          =   360
         Left            =   7080
         TabIndex        =   16
         Top             =   1695
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   635
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "yyyy-MM-dd"
         Mask            =   "####-##-##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPatiBegin 
         Height          =   360
         Left            =   5370
         TabIndex        =   15
         Top             =   1695
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   635
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "yyyy-MM-dd"
         Mask            =   "####-##-##"
         PromptChar      =   "_"
      End
      Begin zl9InExse.txtEdit txt���� 
         Height          =   360
         Left            =   8490
         TabIndex        =   41
         Top             =   1695
         Width           =   495
         _extentx        =   873
         _extenty        =   635
         backcolor       =   -2147483633
         enabled         =   0   'False
         font            =   "frmPatiBalanceSplit.frx":0D4A
         appearance      =   0
         text            =   "123"
      End
      Begin MSMask.MaskEdBox txtEnd 
         Height          =   360
         Left            =   2925
         TabIndex        =   12
         Top             =   1695
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   635
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "yyyy-MM-dd"
         Mask            =   "####-##-##"
         PromptChar      =   "_"
      End
      Begin zl9InExse.ComboxExpend cboӤ�� 
         Height          =   360
         Left            =   13095
         TabIndex        =   5
         Top             =   630
         Width           =   1665
         _extentx        =   2937
         _extenty        =   635
         borderstyle     =   1
         text            =   "���˼�Ӥ��"
         font            =   "frmPatiBalanceSplit.frx":0D6E
         fontname        =   "����"
         fontsize        =   12
      End
      Begin zl9InExse.ComboxExpend cboDept 
         Height          =   360
         Left            =   1095
         TabIndex        =   6
         Top             =   1140
         Width           =   1920
         _extentx        =   3387
         _extenty        =   635
         borderstyle     =   1
         text            =   "�ڿ�,���"
         font            =   "frmPatiBalanceSplit.frx":0D92
         fontname        =   "����"
         fontsize        =   12
      End
      Begin zl9InExse.ComboxExpend cboChargeType 
         Height          =   360
         Left            =   11745
         TabIndex        =   10
         Top             =   1140
         Width           =   3000
         _extentx        =   5292
         _extenty        =   635
         borderstyle     =   1
         text            =   "����,����ҩ,���"
         font            =   "frmPatiBalanceSplit.frx":0DB6
         fontname        =   "����"
         fontsize        =   12
      End
      Begin zl9InExse.ComboxExpend cboFeeType 
         Height          =   360
         Left            =   4110
         TabIndex        =   7
         Top             =   1140
         Width           =   1455
         _extentx        =   2566
         _extenty        =   635
         borderstyle     =   1
         text            =   "����,����"
         font            =   "frmPatiBalanceSplit.frx":0DDA
         fontname        =   "����"
         fontsize        =   12
      End
      Begin zl9InExse.ComboxExpend cboFeeItem 
         Height          =   360
         Left            =   6690
         TabIndex        =   8
         Top             =   1140
         Width           =   1440
         _extentx        =   2540
         _extenty        =   635
         borderstyle     =   1
         text            =   "���Ʒѣ�����"
         font            =   "frmPatiBalanceSplit.frx":0DFE
         fontname        =   "����"
         fontsize        =   12
      End
      Begin zl9InExse.txtEdit txtPatiType 
         Height          =   360
         Left            =   1110
         TabIndex        =   33
         Top             =   585
         Width           =   1815
         _extentx        =   3201
         _extenty        =   635
         backcolor       =   -2147483633
         enabled         =   0   'False
         font            =   "frmPatiBalanceSplit.frx":0E22
         appearance      =   0
         showline        =   1
         text            =   "��ͨ����"
      End
      Begin zl9InExse.txtEdit txtSex 
         Height          =   345
         Left            =   4245
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   60
         Width           =   600
         _extentx        =   1058
         _extenty        =   609
         backcolor       =   -2147483633
         font            =   "frmPatiBalanceSplit.frx":0E46
         locked          =   -1  'True
         appearance      =   0
         showline        =   1
         text            =   "��"
      End
      Begin zl9InExse.txtEdit txtOld 
         Height          =   345
         Left            =   5505
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   60
         Width           =   1005
         _extentx        =   1773
         _extenty        =   609
         backcolor       =   -2147483633
         font            =   "frmPatiBalanceSplit.frx":0E6A
         locked          =   -1  'True
         appearance      =   0
         showline        =   1
         text            =   "23��10��"
      End
      Begin zl9InExse.txtEdit txt�ѱ� 
         Height          =   345
         Left            =   7140
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   53
         Width           =   1560
         _extentx        =   2752
         _extenty        =   609
         backcolor       =   -2147483633
         font            =   "frmPatiBalanceSplit.frx":0E8E
         locked          =   -1  'True
         appearance      =   0
         showline        =   1
         text            =   "��ͨ"
      End
      Begin zl9InExse.txtEdit txt��ʶ�� 
         Height          =   345
         Left            =   9585
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   60
         Width           =   1500
         _extentx        =   2646
         _extenty        =   609
         backcolor       =   -2147483633
         font            =   "frmPatiBalanceSplit.frx":0EB2
         locked          =   -1  'True
         appearance      =   0
         showline        =   1
         text            =   "123"
      End
      Begin zl9InExse.txtEdit txtBed 
         Height          =   345
         Left            =   11715
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   60
         Width           =   780
         _extentx        =   1376
         _extenty        =   609
         backcolor       =   -2147483633
         font            =   "frmPatiBalanceSplit.frx":0ED6
         locked          =   -1  'True
         appearance      =   0
         showline        =   1
         text            =   "123"
      End
      Begin zl9InExse.txtEdit txt���� 
         Height          =   345
         Left            =   13185
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   60
         Width           =   1710
         _extentx        =   3016
         _extenty        =   609
         backcolor       =   -2147483633
         font            =   "frmPatiBalanceSplit.frx":0EFA
         locked          =   -1  'True
         appearance      =   0
         showline        =   1
         text            =   "�����ڿ�"
      End
      Begin zlIDKind.IDKindNew IDKIND 
         Height          =   345
         Left            =   600
         TabIndex        =   1
         Top             =   60
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   609
         Appearance      =   2
         IDKindStr       =   $"frmPatiBalanceSplit.frx":0F1E
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
         NotContainFastKey=   "F1;F2;CTRL+F4;F6;F8;F9;F11;F12;ESC"
         AllowAutoICCard =   -1  'True
         AllowAutoIDCard =   -1  'True
         BackColor       =   -2147483633
      End
      Begin VB.CheckBox chkPatiType 
         Caption         =   "������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   5805
         TabIndex        =   14
         Top             =   1995
         Width           =   1410
      End
      Begin VB.CheckBox chkPatiType 
         Caption         =   "��ͨ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   4380
         TabIndex        =   13
         Top             =   1980
         Width           =   1275
      End
      Begin MSMask.MaskEdBox txtBegin 
         Height          =   360
         Left            =   1110
         TabIndex        =   11
         Top             =   1695
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   635
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "yyyy-MM-dd"
         Mask            =   "####-##-##"
         PromptChar      =   "_"
      End
      Begin VB.OptionButton opt��Ժ 
         Appearance      =   0  'Flat
         Caption         =   "��Ժ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   10965
         TabIndex        =   19
         Top             =   1755
         Value           =   -1  'True
         Width           =   1275
      End
      Begin zl9InExse.txtEdit txtԤ����� 
         Height          =   360
         Left            =   4110
         TabIndex        =   98
         Top             =   585
         Width           =   1275
         _extentx        =   2249
         _extenty        =   635
         backcolor       =   -2147483633
         enabled         =   0   'False
         font            =   "frmPatiBalanceSplit.frx":0FB4
         appearance      =   0
         showline        =   1
         text            =   ""
      End
      Begin zl9InExse.txtEdit txtδ����� 
         Height          =   360
         Left            =   6690
         TabIndex        =   99
         Top             =   585
         Width           =   1275
         _extentx        =   2249
         _extenty        =   635
         backcolor       =   -2147483633
         enabled         =   0   'False
         font            =   "frmPatiBalanceSplit.frx":0FD8
         appearance      =   0
         showline        =   1
         text            =   ""
      End
      Begin zl9InExse.txtEdit txtʣ���� 
         Height          =   360
         Left            =   9270
         TabIndex        =   101
         Top             =   585
         Width           =   1275
         _extentx        =   2249
         _extenty        =   635
         backcolor       =   -2147483633
         enabled         =   0   'False
         font            =   "frmPatiBalanceSplit.frx":0FFC
         appearance      =   0
         showline        =   1
         text            =   ""
      End
      Begin zl9InExse.ComboxExpend cboPatiNums 
         Height          =   360
         Left            =   11745
         TabIndex        =   4
         Top             =   630
         Width           =   2850
         _extentx        =   5027
         _extenty        =   635
         borderstyle     =   1
         text            =   "��1��,��2��"
         font            =   "frmPatiBalanceSplit.frx":1020
         fontname        =   "����"
         fontsize        =   12
      End
      Begin VB.Label lblDiag 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   8205
         TabIndex        =   124
         Top             =   1200
         Width           =   480
      End
      Begin VB.Label lblFeeItem 
         AutoSize        =   -1  'True
         Caption         =   "������Ŀ"
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
         Left            =   5685
         TabIndex        =   123
         Top             =   1200
         Width           =   960
      End
      Begin VB.Label lblFeeType 
         AutoSize        =   -1  'True
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
         Left            =   3105
         TabIndex        =   122
         Top             =   1200
         Width           =   960
      End
      Begin VB.Label lblԤ����� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ԥ�����"
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
         Left            =   3105
         TabIndex        =   121
         Top             =   645
         Width           =   960
      End
      Begin VB.Label lblʣ���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ʣ����"
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
         Left            =   8205
         TabIndex        =   102
         Top             =   645
         Width           =   960
      End
      Begin VB.Label lblδ����� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "δ�����"
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
         Left            =   5685
         TabIndex        =   100
         Top             =   645
         Width           =   960
      End
      Begin VB.Label lblBalanceType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��;����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   13695
         TabIndex        =   92
         Top             =   885
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Line lnPatiSplit 
         BorderColor     =   &H80000003&
         X1              =   -180
         X2              =   30000
         Y1              =   2130
         Y2              =   2130
      End
      Begin VB.Label lblFsTime 
         AutoSize        =   -1  'True
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
         Left            =   105
         TabIndex        =   37
         Top             =   1755
         Width           =   960
      End
      Begin VB.Label lblFsTimeRange 
         AutoSize        =   -1  'True
         Caption         =   "��"
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
         Left            =   2580
         TabIndex        =   38
         Top             =   1755
         Width           =   240
      End
      Begin VB.Label lblPatiTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ�ڼ�"
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
         Left            =   4380
         TabIndex        =   39
         Top             =   1755
         Width           =   960
      End
      Begin VB.Label lblPatiTimeRange 
         AutoSize        =   -1  'True
         Caption         =   "��"
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
         Left            =   6750
         TabIndex        =   40
         Top             =   1755
         Width           =   240
      End
      Begin VB.Label lblDayName 
         AutoSize        =   -1  'True
         Caption         =   "��"
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
         Left            =   9075
         TabIndex        =   89
         Top             =   1755
         Width           =   240
      End
      Begin VB.Label lblPatiType 
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
         Left            =   105
         TabIndex        =   32
         Top             =   645
         Width           =   960
      End
      Begin VB.Label lblPatiNums 
         AutoSize        =   -1  'True
         Caption         =   "סԺ����"
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
         Left            =   10755
         TabIndex        =   34
         Top             =   675
         Width           =   960
      End
      Begin VB.Label lblDept 
         AutoSize        =   -1  'True
         Caption         =   "���ÿ���"
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
         Left            =   105
         TabIndex        =   35
         Top             =   1200
         Width           =   960
      End
      Begin VB.Label lblFeeStyle 
         AutoSize        =   -1  'True
         Caption         =   "�շ����"
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
         Left            =   10755
         TabIndex        =   36
         Top             =   1200
         Width           =   960
      End
      Begin VB.Label lbl�ѱ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ѱ�"
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
         Left            =   6630
         TabIndex        =   24
         Top             =   105
         Width           =   480
      End
      Begin VB.Label lbl���� 
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
         Height          =   240
         Left            =   12585
         TabIndex        =   30
         Top             =   105
         Width           =   480
      End
      Begin VB.Label lblBed 
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
         Height          =   225
         Left            =   11190
         TabIndex        =   28
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lbl��ʶ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ��"
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
         Left            =   8805
         TabIndex        =   26
         Top             =   105
         Width           =   720
      End
      Begin VB.Label lblOld 
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
         Height          =   240
         Left            =   4965
         TabIndex        =   22
         Top             =   112
         Width           =   480
      End
      Begin VB.Label lblSex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
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
         Left            =   3720
         TabIndex        =   20
         Top             =   112
         Width           =   480
      End
      Begin VB.Label lblPatient 
         Alignment       =   1  'Right Justify
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
         Left            =   -60
         TabIndex        =   0
         Top             =   105
         Width           =   690
      End
   End
   Begin VB.PictureBox pic״̬ 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   7170
      ScaleHeight     =   315
      ScaleWidth      =   3225
      TabIndex        =   84
      TabStop         =   0   'False
      Top             =   210
      Visible         =   0   'False
      Width           =   3255
      Begin VB.Label lbl״̬ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         Left            =   90
         TabIndex        =   86
         Top             =   30
         Width           =   960
      End
      Begin VB.Label lbl���ʽ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         Left            =   645
         TabIndex        =   85
         Top             =   30
         Width           =   1920
      End
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   1275
      Top             =   135
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
            Picture         =   "frmPatiBalanceSplit.frx":1044
            Key             =   "Tools"
            Object.Tag             =   "Tools"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiBalanceSplit.frx":15DE
            Key             =   "Down"
            Object.Tag             =   "Down"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiBalanceSplit.frx":1718
            Key             =   "ColImg"
            Object.Tag             =   "ColImg"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraSplitMenu 
      Height          =   45
      Left            =   -30
      TabIndex        =   80
      Top             =   735
      Width           =   30000
   End
   Begin VB.PictureBox picFeeList 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   5925
      Left            =   45
      ScaleHeight     =   5925
      ScaleWidth      =   5895
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   2880
      Width           =   5895
      Begin TabDlg.SSTab tabFeelist 
         Height          =   4770
         Left            =   180
         TabIndex        =   108
         Top             =   225
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   8414
         _Version        =   393216
         Tabs            =   2
         TabHeight       =   582
         TabMaxWidth     =   2646
         WordWrap        =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "������Ϣ"
         TabPicture(0)   =   "frmPatiBalanceSplit.frx":1CB2
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "picFeeContain"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "������ϸ"
         TabPicture(1)   =   "frmPatiBalanceSplit.frx":1CCE
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "picDetailContain"
         Tab(1).ControlCount=   1
         Begin VB.PictureBox picFeeContain 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   4620
            Left            =   0
            ScaleHeight     =   4620
            ScaleWidth      =   5370
            TabIndex        =   111
            Top             =   465
            Width           =   5370
            Begin VB.PictureBox picDeposit 
               AutoRedraw      =   -1  'True
               BorderStyle     =   0  'None
               Height          =   2865
               Left            =   15
               ScaleHeight     =   2865
               ScaleWidth      =   6615
               TabIndex        =   112
               TabStop         =   0   'False
               Top             =   1905
               Width           =   6615
               Begin VB.CommandButton cmdDepositUp 
                  Caption         =   "��"
                  Height          =   525
                  Left            =   4875
                  TabIndex        =   120
                  Top             =   540
                  Width           =   330
               End
               Begin VB.CommandButton cmdDepositDown 
                  Caption         =   "��"
                  Height          =   525
                  Left            =   4875
                  TabIndex        =   119
                  Top             =   1410
                  Width           =   330
               End
               Begin zl9InExse.Command cmdTools 
                  Height          =   330
                  Left            =   4815
                  TabIndex        =   113
                  Top             =   45
                  Width           =   420
                  _extentx        =   741
                  _extenty        =   582
                  font            =   "frmPatiBalanceSplit.frx":1CEA
                  picture         =   "frmPatiBalanceSplit.frx":1D0E
               End
               Begin VSFlex8Ctl.VSFlexGrid vsDeposit 
                  Height          =   1695
                  Left            =   0
                  TabIndex        =   114
                  Top             =   450
                  Width           =   4305
                  _cx             =   7594
                  _cy             =   2990
                  Appearance      =   2
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
                  ForeColorSel    =   -2147483640
                  BackColorBkg    =   -2147483634
                  BackColorAlternate=   -2147483643
                  GridColor       =   -2147483638
                  GridColorFixed  =   -2147483638
                  TreeColor       =   -2147483632
                  FloodColor      =   192
                  SheetBorder     =   -2147483643
                  FocusRect       =   1
                  HighLight       =   1
                  AllowSelection  =   0   'False
                  AllowBigSelection=   0   'False
                  AllowUserResizing=   1
                  SelectionMode   =   1
                  GridLines       =   1
                  GridLinesFixed  =   1
                  GridLineWidth   =   1
                  Rows            =   50
                  Cols            =   4
                  FixedRows       =   1
                  FixedCols       =   0
                  RowHeightMin    =   350
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
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
                  ExplorerBar     =   1
                  PicturesOver    =   0   'False
                  FillStyle       =   0
                  RightToLeft     =   0   'False
                  PictureType     =   0
                  TabBehavior     =   0
                  OwnerDraw       =   0
                  Editable        =   2
                  ShowComboButton =   0
                  WordWrap        =   0   'False
                  TextStyle       =   0
                  TextStyleFixed  =   0
                  OleDragMode     =   0
                  OleDropMode     =   0
                  DataMode        =   0
                  VirtualData     =   -1  'True
                  DataMember      =   ""
                  ComboSearch     =   0
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
                  Begin zl9InExse.Command cmdColSet 
                     Height          =   255
                     Left            =   45
                     TabIndex        =   115
                     Top             =   45
                     Width           =   195
                     _extentx        =   344
                     _extenty        =   450
                     font            =   "frmPatiBalanceSplit.frx":22A8
                  End
               End
               Begin VB.Label lblDeposit 
                  AutoSize        =   -1  'True
                  Caption         =   "Ԥ�����"
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
                  Left            =   105
                  TabIndex        =   117
                  Top             =   105
                  Width           =   960
               End
               Begin VB.Label lblTicketCount 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Ԥ�����վ�:"
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
                  Left            =   1170
                  TabIndex        =   116
                  Top             =   105
                  Width           =   2400
               End
               Begin VB.Line lnDepositSplit 
                  BorderColor     =   &H80000003&
                  X1              =   5415
                  X2              =   5415
                  Y1              =   240
                  Y2              =   2910
               End
            End
            Begin VSFlex8Ctl.VSFlexGrid vsFeeList 
               Height          =   1695
               Left            =   0
               TabIndex        =   118
               Top             =   0
               Width           =   4305
               _cx             =   7594
               _cy             =   2990
               Appearance      =   2
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
               ForeColorSel    =   -2147483640
               BackColorBkg    =   -2147483634
               BackColorAlternate=   -2147483643
               GridColor       =   -2147483638
               GridColorFixed  =   -2147483638
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
               Rows            =   50
               Cols            =   4
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   350
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
               ShowComboButton =   0
               WordWrap        =   0   'False
               TextStyle       =   0
               TextStyleFixed  =   0
               OleDragMode     =   0
               OleDropMode     =   0
               DataMode        =   0
               VirtualData     =   -1  'True
               DataMember      =   ""
               ComboSearch     =   0
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
         Begin VB.PictureBox picDetailContain 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2160
            Left            =   -75000
            ScaleHeight     =   2160
            ScaleWidth      =   2250
            TabIndex        =   109
            Top             =   465
            Width           =   2250
            Begin VSFlex8Ctl.VSFlexGrid vsDetailList 
               Height          =   1695
               Left            =   30
               TabIndex        =   110
               Top             =   -120
               Width           =   4305
               _cx             =   7594
               _cy             =   2990
               Appearance      =   2
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
               ForeColorSel    =   -2147483640
               BackColorBkg    =   -2147483634
               BackColorAlternate=   -2147483643
               GridColor       =   -2147483638
               GridColorFixed  =   -2147483638
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
               Rows            =   50
               Cols            =   4
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   350
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
               Editable        =   2
               ShowComboButton =   0
               WordWrap        =   0   'False
               TextStyle       =   0
               TextStyleFixed  =   0
               OleDragMode     =   0
               OleDropMode     =   0
               DataMode        =   0
               VirtualData     =   -1  'True
               DataMember      =   ""
               ComboSearch     =   0
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
      End
      Begin VB.Line lnFeeSplit 
         BorderColor     =   &H80000003&
         X1              =   5355
         X2              =   5355
         Y1              =   75
         Y2              =   2745
      End
   End
   Begin VB.PictureBox picNO 
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   10875
      ScaleHeight     =   405
      ScaleWidth      =   2085
      TabIndex        =   69
      Top             =   195
      Width           =   2085
      Begin VB.ComboBox cboNO 
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
         Left            =   15
         Locked          =   -1  'True
         TabIndex        =   95
         TabStop         =   0   'False
         Top             =   15
         Width           =   1515
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
         Height          =   375
         Left            =   1620
         Style           =   1  'Graphical
         TabIndex        =   67
         TabStop         =   0   'False
         ToolTipText     =   "�ȼ���F8"
         Top             =   15
         Width           =   450
      End
      Begin VB.Label lblDelCaption 
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
         Left            =   1665
         TabIndex        =   94
         Top             =   15
         Visible         =   0   'False
         Width           =   405
      End
   End
   Begin VB.TextBox txtInvoice 
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
      Left            =   13125
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   225
      Width           =   1425
   End
   Begin VB.PictureBox pic���� 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   6390
      ScaleHeight     =   420
      ScaleWidth      =   720
      TabIndex        =   81
      Top             =   165
      Visible         =   0   'False
      Width           =   750
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   30
         TabIndex        =   82
         Top             =   45
         Width           =   660
      End
   End
   Begin MSCommLib.MSComm com 
      Left            =   2055
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   68
      Top             =   9855
      Width           =   14985
      _ExtentX        =   26432
      _ExtentY        =   767
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2884
            MinWidth        =   882
            Picture         =   "frmPatiBalanceSplit.frx":22CC
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15266
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "�ϴν��ʽ��"
            Object.ToolTipText     =   "�ϴν��ʽ��"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Visible         =   0   'False
            Object.Width           =   0
            Key             =   "����"
            Object.ToolTipText     =   "����"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "�����ʻ����"
            Object.ToolTipText     =   "�����ʻ����"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1270
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
            Object.Width           =   1270
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   270
      Top             =   165
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Image imgCol 
      Height          =   195
      Left            =   300
      ToolTipText     =   "ѡ����Ҫ��ʾ����(ALT+C)"
      Top             =   100
      Width           =   195
   End
   Begin VB.Label lblFact 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   12450
      TabIndex        =   90
      Top             =   285
      Visible         =   0   'False
      Width           =   720
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmPatiBalanceSplit.frx":2B60
      Left            =   810
      Top             =   300
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   0
   End
End
Attribute VB_Name = "frmPatiBalanceSplit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'----------------------------------------------------------------------
'1.������ڲ���
Private mEditType As gBalanceBill
Private mintPreEditType As Integer   '�ϴα༭����
Private mstrPrivs As String, mstrPrivsCard As String, mlngModul As Long
Private mstrInNO As String  '���ʵ���
Private mbln����תסԺ As Boolean 'true:����תסԺ���ýӿ�;FalseΪ����
Private mstrPepositDate As String 'ָ���ص��Ԥ������(��Ҫ��Ӧ��������תסԺ����ʱ,ʹ��ת���Ԥ�����н���)
Private mlngPatientID As Long        '��ǰҪ���ʵĲ���ID
Private mstr��ҳId As String   '��ĳ�η���:0-������;1-��סԺ�ڼ��η���;��Ϊ������
Private mblnNOMoved As Boolean       '�����ĵ����Ƿ��ں����ݱ���
Private mobjInPati As Object
Private mblnViewCancel As Boolean, mblnCancel As Boolean
'----------------------------------------------------------------------
'2.�˵���ر���
Private mcbrControl As CommandBarControl, mcbrToolBar As CommandBar
Private mobjPopup As CommandBarPopup
Private mobjCommandBar As CommandBar
Private mobjControl As CommandBarControl
Private mblnNotChange As Boolean

Private Const M_VIEW_ICO = 102 '��ѯ������ʾ��ͼ��
Private Const conMenu_View_Balance = 9000
Private Const conMenu_View_List = 9001
Private Const conMenu_View_ListItem = 9002
Private Const conMenu_View_SplitType = 9003
Private Const conMenu_View_SplitMonth = 9004
Private Const conMenu_View_DayBill = 9005
Private Const conMenu_View_DayFM = 9006

Private Const conMenu_View_LblFPH = 9010
Private Const conMenu_View_BillFPH = 9011
Private Const conMenu_View_LblNo = 9012
Private Const conMenu_View_BillNo = 9013
Private Const conMenu_View_CHKCancel = 9012
Private Const conMenu_Edit_NotUseDeposit = 9101 '��ʹ��Ԥ��
Private Const conMenu_Edit_UseAllDeposit = 9102 'ʹ�õ�����Ԥ��
Private Const conMenu_Edit_MoneyUseDeposit = 9103  '�����ʽ��ʹ��Ԥ��

'3.����ģ�����
Private mblnFirst As Boolean
Private mblnUnload As Boolean, mblnInterUse As Boolean
Private mrsInfo As ADODB.Recordset
Private mblnDateMoved As Boolean '���˵ĵǼ�ʱ���Ƿ���ת������֮ǰ
Private mblnCurMzBalanceNo As Boolean '��ǰΪ������ʵ�(�ǽ�����ʱ����Ч)
Private mlngCardTypeID As Long '��ǰˢ������56615
Private mstrPassWord As String
Private mblnInvalidLoad As Boolean
Private mstrForceNote As String
Private mstr����סԺ���� As String
Private mblnNotify As Boolean, mstrInvoice As String
Private mblnPrintInvoice As Boolean
Private mstrPatiBegin As String, mstrPatiEnd As String
Private mblnCurPatiInsure As Boolean
Private mblnReadByZYNo As Boolean
Private mstrReadByZYNo As String
Private mstrBalanceLimit As String
Private mstrInputInNo As String, mblnBatchState As Boolean
Private mintSucces As Integer  '�ɹ����ɵ�������
Private mrsFeeList As ADODB.Recordset '����δ�Ს����ϸ
Private mrsDeposit As ADODB.Recordset  '����Ԥ����Ϣ
Private mrsBalance As ADODB.Recordset  '���˽�����Ϣ
Private mrsOldBalance As ADODB.Recordset  '�����˽�����Ϣ
Private mbln�������� As Boolean           '��ǰ�����Ƿ��������ʲ���
Private mrsClassMoney As ADODB.Recordset
Private mblnDepositBillPrint As Boolean '�Ƿ��ӡ�Խ���Ʊ��
Private mrs���㷽ʽ As ADODB.Recordset  '��ǰ��Ч�Ľ��㷽ʽ
Private mstrDec As String       '���ν��ʵķ���С��λ��
Private mblnNotClick As Boolean
Private mblnNotClearBill As Boolean '��������ʽ���
Private mblnLockScreen As Boolean '��ǰ�Ƿ�ˢ��
Private mstr��֧Ʊ As String
Private mstrȱʡ���㷽ʽ As String  'ȱʡ��֧����ʽ
Private mblnConsChange As Boolean '�Ƿ��м������˸ı�
Private mblnSecondLoadPati As Boolean
Private mfrmParent As Object
Private mblnManualEdit As Boolean
Private mlngBalanceRows As Long, mstrCardPara As String
Private mstrNoSort As String, mblnNoTrigger As Boolean
Private mstrPatient As String

'3.1�ӿڶ�����
Private mobjInPatient As Object
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1

'3.2 �ؼ����������������
Private Enum mInput_Idx
    Idx_����δ�� = 0
    Idx_���ν��� = 1
    Idx_����˵�� = 2
    Idx_��Ԥ�� = 3
    Idx_�ɿ� = 4
    Idx_�Ҳ� = 5
    Idx_������� = 6
    Idx_ժҪ = 7
End Enum
Private Enum mCheck_Idx
    CK_Idx_��ͨ = 0
    CK_Idx_��� = 1
End Enum
 
'3.3 ģ���������
Private Type Ty_ModulePara
    int�˿�Ʊ�� As Integer  '0-����ӡ,1-��ʾ��ӡ,2-����ʾ��ӡ;'���˺� ����:27776 ����:2010-02-04 16:49:03
    bln���ʺ�����Ϣ As Boolean    ''���˺� ����:27776 ����:2010-02-04 16:49:03
    bln���ʼ�鲡������ As Boolean '30036
    byt�ɿ�������� As Byte  '
    bytMzDeposit As Byte    '����Ԥ��ȱʡʹ�÷�ʽ:0-ȱʡ��ʹ�ý�;1-�����ʽ��ʹ��Ԥ��;2-ʹ������Ԥ��
    bln�����˿ʽ As Boolean 'True-�����˿�Ĭ�ϰ�Ԥ�����㷽ʽ False-�����˿�Ĭ���ֽ�
    intPatientRange As Integer  '����������ʱ,�Ƿ�ֻ��ʾδ����õĲ���,0-���ѽ���,1-δ����,2-���δ����,3-סԺδ����
    blnZero  As Boolean '����ʱ�Ƿ��������
    strOwnerPayFeeType As String '�Ը��������
    int����ʱ�� As Integer '0-���Ǽ�ʱ��,1-������ʱ��
    byt����ʱ��Ѫ�Ѽ�� As Byte   '34260
    bln����ָ��Ԥ���� As Boolean  '��ʹ��ָ��סԺ������Ԥ����
    bln��;������Ԥ�� As Boolean '��;����ȱʡ��Ԥ����
    bytInvoiceKindZY As Byte     '0-סԺҽ�Ʒ��վ�,1-����ҽ�Ʒ��վ�
    bytInvoiceKindMZ As Byte
    int����ʣ��Ʊ������ As Integer
    blnNotPrintInvioce As Boolean '�Ƚ��Է�ʱ����ӡƱ��
    blnLedWelcome As Boolean
    intOutDay As Integer '���ʿ�ѡ���Ժ��������
    blnAutoOut As Integer   '�Ƿ��Զ���Ժ
    bytFeePrintSet As Byte      '0-����ӡ;1-��ӡ��ʾ;2-��ӡ������ʾ
    byt���ʼ����տ��� As Byte '��Ժ����ʱ��鲡�˵Ĵ��տ���,0-��ֹ,1-����
    bln�Է�ȱʡʹ��Ԥ�� As Boolean '����Էѷ��ý���ʱ,�����Ƿ�����ȱʡʹ��Ԥ������н���: 0-��ʹ��Ԥ����;1-ʹ��Ԥ����,ȱʡΪ��ʹ��Ԥ����
    bytˢ��ȱʡ������ As Byte '86853
    bytԤ��Ʊ�ݴ�ӡ As Byte
    str�ѻ�ҽ�����㷽ʽ As String
    str�Ը��ϼ�ɫ As String
    str��ǰ����ɫ As String
    str�ɿ�ɫ As String
    bln�˿��ֽ�ȱʡ��� As Boolean
    bln�����������˿���� As Boolean
End Type
Private mty_ModulePara As Ty_ModulePara
Private mblnMC_TwoMode As Boolean '�Ƿ�֧�������סԺҽ���������֤������ģʽ

'3.4 ҽ����ض���
Private Type TY_YBInfor
      bln���ʽ��� As Boolean '�����Ƿ񷵻��˸��ʽ���
      cur������� As Currency '�����ʻ����
      cur�����޶� As Currency '�����ʻ�����޶�
      cur����͸֧ As Currency '�����ʻ�����͸֧���
      cur����֧�� As Currency   '��ǰ�����ʻ�֧��
      curͳ��֧�� As Currency   '��ǰҽ��ͳ��֧��
      strYBPati As String    'ҽ�����������Ϣ
      intInsure As Integer   '����ʱ,��ȡ�ĵ����е�����,�����ж��Ƿ����ֽ�,������
      blnҽ������ȫ�� As Boolean     '�Ƿ��в�֧�ֵ����Ͻ��㷽ʽ
      bytMCMode As Byte 'ҽ���������֤��ģʽ,����1-����,2-סԺ����ģʽ,0-��ʾ��ҽ��
      strBalance As String 'ҽ�����صĸ��ֽ�����:"���㷽ʽ;���;�Ƿ������޸�|...
      blnAutoOut As Boolean '��Ժ���˽��ʺ��Ƿ��Զ���Ժ
End Type
Private mYBInFor As TY_YBInfor 'ҽ�������Ϣ
'��ǰ���������ҽ��֧�ֲ���
Private Type TYPE_MedicarePAR
    '1.���סԺ���㹲�õĲ���
    �ֱҴ��� As Boolean
    ҽ���ӿڴ�ӡƱ�� As Boolean
    '2.��������õĲ���
    ���ﲡ�˽������� As Boolean
    ������봫����ϸ As Boolean
    ����Ԥ���� As Boolean
    �������_�������� As Boolean
    '3.סԺ�����õĲ���
    δ�����Ժ As Boolean
    ����ʹ�ø����ʻ� As Boolean
    ��Ժ��������Ժ As Boolean
    ��Ժ���˽������� As Boolean
    ��;������������ϴ����� As Boolean
    �������ú���ýӿ� As Boolean
    �������Ϻ��ӡ�ص� As Boolean
    סԺ�������� As Boolean
    �������סԺ���� As Boolean
End Type
Private MCPAR As TYPE_MedicarePAR

'-----------------------------------------------------------------
'3.4�ϰ�һ��ͨ���
Private Type TY_OneCard
      blnOneCard As Boolean      '�Ƿ�������һ��ͨ�ӿ�
      rsOneCard As ADODB.Recordset
      strOneCard As String       '����ʱ��ѡ���һ��ͨ�ӿڶ�Ӧ�Ľ��㷽ʽ
End Type
Private mOldOneCard As TY_OneCard
'-----------------------------------------------------------------
'3.5 �����������
Private mobjBalanceAll As clsBalanceAllCon
Private mobjBalanceCon As clsBalanceCon

'��ǰ��������
Private Type TY_Balance_Infor
    dblҽ��֧���ϼ� As Currency  'ҽ��֧���ϼ�
    dbl��Ԥ���ϼ� As Double
    dbl����δ�� As Double
    dbl��ǰ���� As Double
    dbl�Ѹ��ϼ� As Double
    dblδ���ϼ� As Double
    dblԤ�����ܶ� As Double
    blnԤ����֤ As Boolean 'Ԥ�����Ƿ��Ѿ���֤
    blnԤ��ˢ�� As Boolean 'Ԥ�����Ƿ��Ѿ�ˢ��
    blnSaveBill As Boolean '��ǰ�Ѿ�������ʵ�
    strNO As String   '��ǰ����Ľ��ʵ�
    lng����ID As Long '��ǰ����Ľ���ID
    dtBalanceDate As Date '��ǰ����ʱ��
    str����ԭ�� As String '����ԭ��
    dbl�ɿ� As Double
    dbl�Ҳ� As Double
    dbl��֧Ʊ As Double
    dbl���� As Double
    dbl�ֽ� As Double
    lngԤ��ID As Long
    strԤ��No As String
    lng����ID As Long
End Type
Private mBalanceInfor As TY_Balance_Infor
'���˵�ǰ��Ϣ
Private Type ty_Pati_Infor
    lng����ID  As Long
    lng��ҳID As Long
    str���� As String
    str�Ա� As String
    str���� As String
    objCard As Card         '�ϴν�����Ϣ
    bln�������� As Boolean  '�Ƿ���������
    bln��Ժ As Boolean      '��ǰ�����Ƿ��Ժ
    dblԤ����� As Double   '����Ԥ�����
    dbl������� As Double   'δ�����
    dblʣ��ϼ� As Double   '����Ԥ�����-δ�����
    dblʵ����� As Double   'Ԥ����ϸ���
    dblδ���ۼ� As Double  '�ϴ�δ���ۼƽ��
    bln�˿��־ As Boolean
    int�������� As Integer
End Type
Private mPatiInfor As ty_Pati_Infor

'��ǰ��Ʊ��Ϣ
Private mobjInvoice As clsInvoice
Private mobjFactProperty As clsFactProperty
Private mobjRedProperty As clsFactProperty
Private mobjDepositFactProperty As clsFactProperty
Private mstrDepositInvioce As String '��ǰԤ����Ʊ��
Private mlng����ID As Long
Private mlngԤ������ID As Long

'���ѿ�ˢ����Ϣ
Private mcllSquareBalance As Collection '���ѿ�������Ϣ
Private mcllCurSquareBalance As Collection '��ǰ���ѿ�ˢ����Ϣ

'��ǰˢ����Ϣ
Private Type TY_BrushCard    'ˢ������
    str���� As String
    str���� As String
    str������ˮ�� As String    '������ˮ��
    str����˵��  As String     '������Ϣ
    str��չ��Ϣ As String    '���׵���չ��Ϣ
    dbl�ʻ���� As Double
    str������� As String
    str����ժҪ As String
    blnת�� As Boolean '�Ƿ�ǰΪת�ʽ���
End Type
Private Enum mConPans
    Pan_PatiCon = 1
    Pan_FeeList = 2
    Pan_Deposit = 3
    Pan_Balance = 4
End Enum
Private mbln�ѱ��� As Boolean

Public Function ShowMe(ByVal frmMain As Object, ByVal EditType As gBalanceBill, _
    ByVal strPrivs As String, Optional lng����ID As Long = 0, Optional str��ҳID As String = "", _
    Optional ByVal strNO As String, Optional blnViewCancel As Boolean, Optional blnNOMoved As Boolean, _
    Optional bln����תסԺ As Boolean, Optional strPepositDate As String, Optional blnCancel As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ʳ������
    '���:EditType-�༭����
    '     strPrivs-Ȩ�޴�
    '     lng����ID-��ǰҪ���ʵĲ���ID
    '     str��ҳId As String   '��ĳ�η���:0-������;1-��סԺ�ڼ��η���;��Ϊ������
    '     strNo-����Ҫ�����Ľ��ʵ���,�½���ʱ,������
    '     blnViewCancel-�Ƿ�鿴�����ϵ���
    '     blnNOMoved-strNo�Ƿ��Ѿ�ת��󱸱���
    '     bln����תסԺ-true:����תסԺ���ýӿ�;FalseΪ����
    '     strPepositDate-ָ���ص��Ԥ������(��Ҫ��Ӧ��������תסԺ����ʱ,ʹ��ת���Ԥ�����н���)
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-12-29 15:24:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error Resume Next
    mEditType = EditType: mstrPrivs = strPrivs
    mstrInNO = strNO: mbln����תסԺ = bln����תסԺ
    mstrPepositDate = strPepositDate: mlngPatientID = lng����ID
    mstr��ҳId = str��ҳID: mintSucces = 0: mblnNOMoved = blnNOMoved
    Set mfrmParent = frmMain
    mblnViewCancel = blnViewCancel
    mblnCancel = blnCancel
    mintPreEditType = -1 '�ϴα༭��������Ϊ����,�����ڱ������ݺ���н���ָ�����
    Set mobjInvoice = New zlPublicExpense.clsInvoice
    If mobjInvoice.zlInitCommon(glngSys, gcnOracle, gstrDBUser) = False Then Exit Function
    Me.Show IIf(gfrmMain Is Nothing, 0, 1), frmMain
    ShowMe = mintSucces > 0
End Function


Private Sub zlInitModulePara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ģ�����
    '����:���˺�
    '����:2010-02-04 16:50:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strErrMsg As String
    With mty_ModulePara
        '0-����ӡ,1-��ʾ��ӡ,2-����ʾ��ӡ;'���˺� ����:27776 ����:2010-02-04 16:49:03
        .int�˿�Ʊ�� = Val(zlDatabase.GetPara("�˿��վݴ�ӡ", glngSys, mlngModul))
        .bln���ʺ�����Ϣ = IIf(Val(zlDatabase.GetPara("���ʺ������Ϣ", glngSys, mlngModul)) = 1, True, False)
        .bln���ʼ�鲡������ = IIf(Val(zlDatabase.GetPara("���ʼ�鲡������", glngSys, mlngModul)) = 1, True, False) '30036
        '����:43153:0-�����п���;1-������ȡ�ֽ�ʱ,��������ɿ�;2-����ʱ���������ۼ�
        .byt�ɿ�������� = Val(zlDatabase.GetPara("���ʽɿ��������", glngSys, mlngModul, 0))
        .bytMzDeposit = Val(zlDatabase.GetPara("����Ԥ��ȱʡʹ�÷�ʽ", glngSys, mlngModul, 2))
        .bln�����˿ʽ = IIf(Val(zlDatabase.GetPara("�����˿�ȱʡ��ʽ", glngSys, mlngModul)) = 1, True, False)
        .intPatientRange = Val(zlDatabase.GetPara("��ʾ���岡��", glngSys, mlngModul, 0))
        .blnZero = zlDatabase.GetPara("���������", glngSys, mlngModul) = "1"
        .strOwnerPayFeeType = zlDatabase.GetPara("����ǰ�Ƚ��Էѷ���", glngSys, mlngModul, "")
        .int����ʱ�� = IIf(zlDatabase.GetPara("���ʷ���ʱ��", glngSys, mlngModul) = "1", 1, 0)
        .byt����ʱ��Ѫ�Ѽ�� = Val(zlDatabase.GetPara("����ʱ��Ѫ�Ѽ��", glngSys, mlngModul, "0"))
        .bln����ָ��Ԥ���� = zlDatabase.GetPara("����ָ��Ԥ����", glngSys, mlngModul) = "1"
        .bln��;������Ԥ�� = zlDatabase.GetPara("��;������Ԥ��", glngSys, mlngModul) = "1"
        .bytInvoiceKindZY = Val(zlDatabase.GetPara("סԺ����Ʊ������", glngSys, mlngModul, "0"))
        .bytInvoiceKindMZ = Val(zlDatabase.GetPara("�������Ʊ������", glngSys, mlngModul, "0"))
        .blnLedWelcome = zlDatabase.GetPara("LED��ʾ��ӭ��Ϣ", glngSys, mlngModul, "1") = "1"
        .blnNotPrintInvioce = Val(zlDatabase.GetPara("�Ƚ��Էѷ��ò���ӡ����Ʊ��", glngSys, mlngModul, "0")) = 1
        .blnAutoOut = zlDatabase.GetPara("��Ժ���˽��ʺ��Զ���Ժ", glngSys, mlngModul) = "1"
        .bytFeePrintSet = Val(zlDatabase.GetPara("������ϸ��ӡ", glngSys, mlngModul, "0"))
        .byt���ʼ����տ��� = zlDatabase.GetPara("���ʼ����տ���", glngSys, mlngModul, , "0")
        .int����ʣ��Ʊ������ = 0 '��ʱδ�з�Ʊ�����Ĳ�������
        .bln�Է�ȱʡʹ��Ԥ�� = Val(zlDatabase.GetPara("�Է�ȱʡʹ��Ԥ��", glngSys, mlngModul, "0")) = 1
        .bytˢ��ȱʡ������ = Val(zlDatabase.GetPara("ˢ��ȱʡ������", glngSys, 1151, "0")) '86853
        .bytԤ��Ʊ�ݴ�ӡ = Val(zlDatabase.GetPara("Ԥ��Ʊ�ݴ�ӡ��ʽ", glngSys, mlngModul, "0"))
        .str�ѻ�ҽ�����㷽ʽ = zlDatabase.GetPara("�ѻ�ҽ�����㷽ʽ", glngSys)
        .str��ǰ����ɫ = zlDatabase.GetPara("��ǰ����������ɫ", glngSys, mlngModul, "255|255")
        .str�ɿ�ɫ = zlDatabase.GetPara("�ɿ�������ɫ", glngSys, mlngModul, "16711680|255")
        .str�Ը��ϼ�ɫ = zlDatabase.GetPara("�Ը��ϼ�������ɫ", glngSys, mlngModul, "16711680")
        .bln�˿��ֽ�ȱʡ��� = zlDatabase.GetPara("�˿��ֽ����ȱʡ���", glngSys, mlngModul) = "1"
        .bln�����������˿���� = zlDatabase.GetPara("�����������˿����", glngSys, mlngModul) = "1"
    End With
    
    lbl�Ը��ϼ�.ForeColor = mty_ModulePara.str�Ը��ϼ�ɫ
    lblʣ���Ը�.ForeColor = Mid(mty_ModulePara.str��ǰ����ɫ, 1, InStr(mty_ModulePara.str��ǰ����ɫ, "|") - 1)
    txtBalance(4).ForeColor = Mid(mty_ModulePara.str�ɿ�ɫ, 1, InStr(mty_ModulePara.str�ɿ�ɫ, "|") - 1)
    lblBalance(4).ForeColor = Mid(mty_ModulePara.str�ɿ�ɫ, 1, InStr(mty_ModulePara.str�ɿ�ɫ, "|") - 1)
    IDKindPaymentsType.ForeColor = Mid(mty_ModulePara.str�ɿ�ɫ, 1, InStr(mty_ModulePara.str�ɿ�ɫ, "|") - 1)
    
    '�ɶ��ϰ�ҽ��֧�������סԺ���������֤ģʽ
    mblnMC_TwoMode = InStr("," & GetSetting("ZLSOFT", "����ȫ��", "����֧�ֵ�ҽ��", "") & ",", ",20,") > 0
    
    mstrPrivsCard = ";" & GetPrivFunc(glngSys, 1151) & ";"
End Sub
Private Sub SetCurBalanceVisible()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ý�����Ϣ�Ƿ���ʾ
    '����:���˺�
    '����:2015-01-19 16:49:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnVisible As Boolean
    
    blnVisible = Not mEditType = g_Ed_���ݲ鿴
    picCurBalance.Visible = blnVisible
    picOwerFee.Visible = blnVisible
    picNotPayment.Visible = blnVisible
    picCurDeposit.Visible = blnVisible
     
    If mEditType = g_Ed_�������� Then
        lblBalance(0).Visible = False
        txtBalance(Idx_����δ��).Visible = False
    Else
        lblBalance(0).Visible = blnVisible
        txtBalance(Idx_����δ��).Visible = blnVisible
        If blnVisible = False Then
            Set lblBalance(7).Container = picBalanceInfor
            lblBalance(7).Left = lblBalance(0).Left + 30
            lblBalance(7).Top = lblBalance(0).Top
            txtDate.Top = txtBalance(Idx_����δ��).Top + 15
            txtDate.Left = txtBalance(Idx_����δ��).Left - 15
            txtDate.Width = txtDate.Width - 15
            Set txtDate.Container = picBalanceInfor
            shpFun.Visible = False
        End If
    End If
    
    lblBalance(3).Visible = True
    If (mEditType = g_Ed_�������� Or mEditType = g_Ed_�������� Or chkCancel.Value = 1) Then
        If InStr(1, mstrPrivs, ";Ԥ�����ֽ�;") > 0 Then
            'blnVisible = InStr(";" & mstrPrivsCard & ";", ";�����˿�ǿ������;") > 0
            'If blnVisible = False Then blnVisible = CheckIsDepositAllowCash
            
            blnVisible = True
            chkDeposit.Visible = blnVisible
            cmdTools.Visible = blnVisible
            lblBalance(3).Visible = Not blnVisible
        Else
            cmdTools.Visible = False
            chkDeposit.Visible = False
        End If
    End If
    
    If mEditType = g_Ed_�������� Or mEditType = g_Ed_�������� _
        Or mEditType = g_Ed_���ݲ鿴 Or chkDeposit.Visible Or chkCancel.Value = 1 Then
        cmdTools.Visible = False
        cmdDepositUp.Visible = False
        cmdDepositDown.Visible = False
    Else
        cmdTools.Visible = blnVisible
        cmdDepositUp.Visible = blnVisible
        cmdDepositDown.Visible = blnVisible
    End If
    Call picDeposit_Resize
    
End Sub

Private Function CheckIsDepositAllowCash() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ж��Ƿ�������������������ֵĽ�����Ϣ
    '���:
    '����:
    '����:�������֣�����True,����Fale
    '����:���˺�
    '����:2018-02-08 15:34:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng�����ID As Long, i As Long
    
    On Error GoTo errHandle
    With vsDeposit
        For i = 1 To .Rows - 1
            If .ColIndex("�����ID") >= 0 Then
                lng�����ID = Val(.TextMatrix(i, .ColIndex("�����ID")))
                'bln���ѿ� = Val(.TextMatrix(i, .ColIndex("�Ƿ����ѿ�"))) = 1
                If lng�����ID <> 0 And Val(.TextMatrix(i, .ColIndex("�Ƿ�����"))) = 0 Then Exit Function
            End If
        Next
    End With
    CheckIsDepositAllowCash = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Sub InitFace()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������
    '����:���˺�
    '����:2014-12-19 11:18:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim blnStartFactUseType As Boolean    '�Ƿ������˶���ʹ������Ʊ��

    If mEditType = g_Ed_������� Then
        blnStartFactUseType = zlStartFactUseType(IIf(mty_ModulePara.bytInvoiceKindMZ = 0, 3, 1))
    ElseIf mEditType = g_Ed_סԺ���� Then
        blnStartFactUseType = zlStartFactUseType(IIf(mty_ModulePara.bytInvoiceKindZY = 0, 3, 1))
    End If
    dkpMain.SetCommandBars Me.cbsThis
    
    Call InitPatiBalanceVariableCon
    
    Call InitVar    '��ʼ���ڲ���ر���
        
    Set cmdColSet.Picture = imgCol.Picture
    Call initCardSquareData '��ʼ��������
    Call Load�Ҳ���(0, "��   ��") '��ʼ���Ҳ���
    Call InitOldOneCardInfor '��ʼ����һ��ͨ��ر���
    Call InitCombox_Cons '��ʼ������������Ϣ
    Call InitGrid
    
    Call SetCurBalanceVisible   '���õ�ǰ������Ϣ����ʾ
    Call InitPancel '��ʼ������
    'all SetPatiConsControlVisible '���ò��������ؼ�����ʾ
    'Call SetOperationCtrl(0) 'bytFun-0-����ǰ;1-ҽ����������;2-�ѱ����˽��ʵ�;
    
    '�����:112545,����,2017/08/25,��Ʊ�������ʱ,������ʽ������ʾ��ǰ����Ա��Ʊ�ݺ�
    Call ReInitPatiInvoice(Not blnStartFactUseType)
    
    Set cmdColSet.Picture = imgList.ListImages("ColImg").Picture
    Call SetOperatonCommandCaption
    cboChargeType.RaisEffect picOwerFee, Dw_SubKen
    cboChargeType.RaisEffect picNotPayment, Dw_SubKen
    txtInvoice.Locked = Not (zlStr.IsHavePrivs(mstrPrivs, "�޸�Ʊ�ݺ�")) And gblnStrictCtrl '89302
    
    If mblnMC_TwoMode Then
        cmdYB.Caption = "ˢ"
        cmdYB.Width = 400
    End If
    Call NewBill
    vsFeeList_LostFocus
    vsDeposit_LostFocus
    vsBlance_LostFocus
    
End Sub
Private Sub InitVar()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������ڲ�����
    '����:���˺�
    '����:2015-01-14 11:27:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    On Error GoTo errHandle
'    mstrDec = gstrDec
    mstrDec = "0.00"
    
    Set mobjFactProperty = New clsFactProperty
    Set mobjRedProperty = New clsFactProperty
    Set mobjDepositFactProperty = New clsFactProperty
    
    If mEditType = g_Ed_������� Then
        Call mobjInvoice.zlGetInvoicePreperty(mlngModul, IIf(mty_ModulePara.bytInvoiceKindMZ = 0, 3, 1), 0, 0, 0, mobjFactProperty, , , 1)
    Else
        Call mobjInvoice.zlGetInvoicePreperty(mlngModul, IIf(mty_ModulePara.bytInvoiceKindZY = 0, 3, 1), 0, 0, 0, mobjFactProperty, , , 2)
    End If
    
   mstr��֧Ʊ = ""
   strSQL = " " & _
    " Select B.���� " & _
    " From ���㷽ʽӦ�� A, ���㷽ʽ B " & _
    " Where A.Ӧ�ó��� = '����' And B.���� = A.���㷽ʽ " & _
    "       And Nvl(B.Ӧ����, 0) = 1 And Rownum <= 1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If Not rsTemp.EOF Then
        mstr��֧Ʊ = NVL(rsTemp!����)
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub InitRedInvoice(Optional blnFact As Boolean = True)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���³�ʼ����Ʊ��Ϣ
    '���:blnFact-�Ƿ�ˢ�·�Ʊ��
    '����:���˺�
    '����:2015-01-07 16:32:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInvoiceFormat As String, bytKind As Byte
    Dim lng����ID As Long, lng��ҳID As Long, intInsure As Integer
    
    intInsure = mYBInFor.intInsure
    
    If Not mrsInfo Is Nothing Then
        If mrsInfo.State = 1 Then
            lng����ID = Val(NVL(mrsInfo!����ID)): lng��ҳID = Val(NVL(mrsInfo!��ҳID))
            intInsure = mYBInFor.intInsure
        End If
    End If
    
    If mPatiInfor.int�������� = 1 Then
        bytKind = IIf(mty_ModulePara.bytInvoiceKindMZ = 0, 13, 11)
    Else
        bytKind = IIf(mty_ModulePara.bytInvoiceKindZY = 0, 13, 11)
    End If
    
    Call mobjInvoice.zlGetInvoicePreperty(mlngModul, bytKind, lng����ID, lng��ҳID, intInsure, mobjRedProperty)
    If mobjRedProperty.����ʹ����� Then mlng����ID = 0
    If blnFact Then Call RefreshRed
End Sub

Private Sub RefreshRed()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ˢ�½��ʵ�Ʊ�ݺ�
    '����:���˺�
    '����:2015-01-07 17:16:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFactNO As String
    If mobjRedProperty Is Nothing Then Exit Sub
    If mobjRedProperty.��ӡ��ʽ = 0 Then Exit Sub
      
    If Not mobjRedProperty.�ϸ���� Then
        '���ϸ������
        '��ɢ��ȡ��һ������
        mstrInvoice = zlCommFun.IncStr(UCase(zlDatabase.GetPara("��ǰ����Ʊ�ݺ�", glngSys, 1137, "")))
        Exit Sub
    End If
    
    If zlGetRedGroupUseID(mlng����ID, 1, "") = False Then
        mstrInvoice = ""
        Exit Sub
    End If
    
    '�ϸ�ȡ��һ������
    If mobjInvoice.zlGetNextBill(mlngModul, mlng����ID, strFactNO) = False Then strFactNO = ""
    mstrInvoice = strFactNO
    
    If mobjRedProperty.����ʹ����� Then Call zlCheckFactIsEnough(0, True)
End Sub

Private Sub InitCombox_Cons()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����������
    '����:���˺�
    '����:2015-01-05 14:31:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    cboDept.Clear: cboDept.AddItem "R", "���п���", True, True, True
    cboFeeItem.Clear: cboFeeItem.AddItem "R", "���з�����Ŀ", True, True, True
    cboChargeType.Clear: cboChargeType.AddItem "R", "�����շ����", True, True, True
    cboFeeType.Clear: cboFeeType.AddItem "R", "���з�������", True, True, True
    
'    cboӤ��.Checkboxes = False
    cboӤ��.Clear: cboӤ��.AddItem "R", "���˼�Ӥ��", True, True, True
    cboӤ��.AddItem "0", "���˷���", False, True
    cboDiag.Clear
    cboDiag.AddItem "�������"
    cboDiag.ListIndex = cboDiag.NewIndex
    
    cboPatiNums.Clear
    If mEditType = g_Ed_������� Then
        cboPatiNums.AddItem "R", "��������", True, True, True
    Else
        cboPatiNums.AddItem "R", "����סԺ", True, True, True
    End If
End Sub


Public Function zlDefCommandBars() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���˵���������
    '����:���óɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-01-25 15:29:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    Dim objCustom As CommandBarControlCustom
    Dim objControl As CommandBarControl
      
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsThis.VisualTheme = xtpThemeOffice2003
    With cbsThis.Options
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        .ShowExpandButtonAlways = False
        Set .Font = cboPatiNums.Font
    End With
    cbsThis.EnableCustomization False
    '-----------------------------------------------------
    '�˵�����
    cbsThis.DeleteAll
    Set mcbrToolBar = cbsThis.Add("������", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = True
    mcbrToolBar.ContextMenuPresent = False
    mcbrToolBar.EnableDocking xtpFlagStretched
    With mcbrToolBar.Controls
        Set objControl = .Add(xtpControlSplitButtonPopup, conMenu_View_Balance, "���ʱ�")
        objControl.IconId = M_VIEW_ICO
        With objControl.CommandBar.Controls
            Set mcbrControl = .Add(xtpControlButton, conMenu_View_List, "��ϸ��")
            mcbrControl.IconId = M_VIEW_ICO
            Set mcbrControl = .Add(xtpControlButton, conMenu_View_ListItem, "��Ŀ��ϸ")
            mcbrControl.IconId = M_VIEW_ICO
            Set mcbrControl = .Add(xtpControlButton, conMenu_View_SplitType, "�����")
            mcbrControl.IconId = M_VIEW_ICO
            Set mcbrControl = .Add(xtpControlButton, conMenu_View_SplitMonth, "���±�")
            mcbrControl.IconId = M_VIEW_ICO
            Set mcbrControl = .Add(xtpControlButton, conMenu_View_DayBill, "���յ���")
            mcbrControl.IconId = M_VIEW_ICO
            Set mcbrControl = .Add(xtpControlButton, conMenu_View_DayFM, "���շ�Ŀ")
            mcbrControl.IconId = M_VIEW_ICO
        End With
        If InStr(1, mstrPrivs, ";�������תסԺ;") > 0 Then
            Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ClinicToHos, "����תסԺ")
            mcbrControl.IconId = 3036
            mcbrControl.BeginGroup = True
        End If
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
        mcbrControl.BeginGroup = True
    End With
        '���˵��Ҳ�Ĳ���
    With mcbrToolBar.Controls
         Set mcbrControl = .Add(xtpControlLabel, conMenu_View_LblFPH, "��Ʊ��")
         mcbrControl.flags = xtpFlagRightAlign
         
        Set objCustom = .Add(xtpControlCustom, conMenu_View_BillFPH, "")
        objCustom.Handle = txtInvoice.hWnd
        objCustom.flags = xtpFlagRightAlign
        

        Set mcbrControl = .Add(xtpControlLabel, conMenu_View_LblNo, " ���ݺ�")
        mcbrControl.flags = xtpFlagRightAlign
        
        Set objCustom = .Add(xtpControlCustom, conMenu_View_BillNo, "")
        objCustom.Handle = picNO.hWnd
        objCustom.flags = xtpFlagRightAlign
  
'        Set objCustom = .Add(xtpControlCustom, conMenu_View_CHKCancel, "")
'        objCustom.Handle = picCancel.hWnd
'        objCustom.Flags = xtpFlagRightAlign
        
        'IDKind.BackColor = picBillNo.BackColor
    End With

    For Each mcbrControl In mcbrToolBar.Controls
        Select Case mcbrControl.ID
        Case conMenu_View_LblFPH, conMenu_View_LblNo
        Case Else
            mcbrControl.Style = xtpButtonIconAndCaption
        End Select
    Next
    zlDefCommandBars = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub cboChargeType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cboChargeType_NodeCheck(ByVal Node As MSComctlLib.Node, strCaption As String)
    Dim blnOwner As Boolean, strKey As String
    Dim objNode As Node, blnAll As Boolean, blnCurOwner As Boolean
    Dim strValue As String, blnExistOwner As Boolean '�����Էѵ�
    
    blnOwner = InStr(1, "," & mty_ModulePara.strOwnerPayFeeType & ",", "," & Replace(Node.Tag, "'", "") & ",") > 0 Or Node.Key = "RootOwner"
    blnAll = True
    blnExistOwner = False
    For Each objNode In cboChargeType.Nodes
        
        blnCurOwner = InStr(1, "," & mty_ModulePara.strOwnerPayFeeType & ",", "," & Replace(objNode.Tag, "'", "") & ",") > 0 Or objNode.Key = "RootOwner"
        If blnCurOwner And Not blnExistOwner Then blnExistOwner = True
        
        If blnOwner Then
            '��ǰΪ�Է�
            If blnCurOwner Then
                If Node.Key Like "Root*" Then
                    objNode.Checked = Node.Checked
                End If
                If objNode.Checked = False And Not objNode.Key Like "Root*" Then blnAll = False
            Else
                '��ǰΪ���Է�
                If Node.Key Like "Root*" Then
                    objNode.Checked = Not Node.Checked
                Else
                    objNode.Checked = False  ' Not Node.Checked
                End If
            End If
        Else
            '��ǰΪ���Է�
            If blnCurOwner Then
                If Node.Key Like "Root*" Then
                    objNode.Checked = Not Node.Checked
                Else
                    objNode.Checked = False  'Not Node.Checked
                End If
            Else
                If Node.Key Like "Root*" Then
                    objNode.Checked = Node.Checked
                End If
                If objNode.Checked = False And Not objNode.Key Like "Root*" Then blnAll = False
            End If
        End If
        If objNode.Checked And Not objNode.Key Like "Root*" Then strValue = strValue & "," & objNode.Text
    Next
    If strValue <> "" Then strValue = Mid(strValue, 2)
    
    strKey = IIf(blnOwner, "RootOwner", "RootBalance")
    If cboChargeType.ListCount = 1 Then Exit Sub
    Set objNode = cboChargeType.Nodes(strKey)
    objNode.Checked = blnAll
    If blnAll Then strCaption = objNode.Text: GoTo GoEnd:
    If strValue <> "" Then strCaption = strValue:  GoTo GoEnd:
    
    For Each objNode In cboChargeType.Nodes
        blnCurOwner = InStr(1, "," & mty_ModulePara.strOwnerPayFeeType & ",", "," & Replace(objNode.Tag, "'", "") & ",") > 0 Or objNode.Key = "RootOwner"
        If blnOwner Then
           If Not blnCurOwner Then
              '���Էѵģ�ȫ��
              objNode.Checked = True
           End If
        Else
           If blnCurOwner Then
              '�Էѵģ�ȫ��
              objNode.Checked = True
           End If
        End If
    Next
    If strValue <> "" Then strValue = Mid(strValue, 2)
    If blnExistOwner Then
        strKey = IIf(Not blnOwner, "RootOwner", "RootBalance")
    Else
        strKey = "RootBalance"
    End If
    Set objNode = cboChargeType.Nodes(strKey)
    strCaption = objNode.Text
    
    
GoEnd:
    lblFeeStyle.ForeColor = &H80000008
    cboChargeType.ForeColor = &H80000008
    Call GetChargeTypeCheckDatas(blnCurOwner)
    If blnCurOwner Then
        lblFeeStyle.ForeColor = vbRed
        cboChargeType.ForeColor = vbRed
    End If
        
    If mblnConsChange Then Exit Sub
    With mobjBalanceCon
         mblnConsChange = .strChargeType <> GetChargeTypeCheckDatas()
    End With
    Call ClearListData
End Sub
Private Function GetChargeTypeCheckDatas(Optional ByRef blnOutCurOwner As Boolean) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�շ����ѡ�е�����
    '����: blnOutCurOwner-��ǰ�����Ƿ��Էѽ���
    '����:��ȡ�ɹ�,����ѡ�е��շ����,��ʽΪ:'C','D',....
    '     �������û���Է����,��ȫѡʱ,�򷵻ؿ�
    '����:���˺�
    '����:2015-02-15 11:14:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objNode As Node, strValue As String
    Dim blnAll As Boolean, blnCurOwner As Boolean, blnExistOwner As Boolean
    On Error GoTo errHandle
    strValue = "":  blnOutCurOwner = False: blnExistOwner = False
    For Each objNode In cboChargeType.Nodes
    
        If objNode.Key Like "Root*" And objNode.Checked Then blnAll = True
        blnCurOwner = InStr(1, "," & mty_ModulePara.strOwnerPayFeeType & ",", "," & Replace(objNode.Tag, "'", "") & ",") > 0 Or objNode.Key = "RootOwner"
        If Not blnExistOwner And blnCurOwner Then blnExistOwner = True
        If objNode.Checked And Not objNode.Key Like "Root*" Then
            strValue = strValue & "," & objNode.Tag
            If Not blnOutCurOwner And blnCurOwner Then blnOutCurOwner = True
        End If
    Next
    If strValue <> "" Then strValue = Mid(strValue, 2)
    If Not blnExistOwner And blnAll Then 'ѡ���������շ����
       strValue = ""
    End If
    GetChargeTypeCheckDatas = strValue
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 

Private Sub cboChargeType_Validate(Cancel As Boolean)
    Dim blnCurOwner As Boolean
    If chkCancel.Value = 1 Then Exit Sub
    If Not (mEditType = g_Ed_������� Or mEditType = g_Ed_סԺ����) Then Exit Sub
    lblFeeStyle.ForeColor = &H80000008
    cboChargeType.ForeColor = &H80000008
    Call GetChargeTypeCheckDatas(blnCurOwner)
    If blnCurOwner Then
        lblFeeStyle.ForeColor = vbRed
        cboChargeType.ForeColor = vbRed
    End If
End Sub

Private Sub cboDept_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
 
Private Sub cboDept_NodeCheck(ByVal Node As MSComctlLib.Node, strCaption As String)
   If mblnConsChange Then Exit Sub
   
   With mobjBalanceCon
        mblnConsChange = .strDeptIDs <> cboDept.GetNodesCheckedDatas()
   End With
   Call ClearListData
End Sub

Private Sub cboDiag_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cboNO_GotFocus()
    If Not cboNO.Locked Then cboNO.SelStart = 0: cboNO.SelLength = Len(cboNO.Text)
End Sub
Private Sub cboNO_KeyPress(KeyAscii As Integer)
    Dim strOper As String, vDate As Date, bytFlag As Byte
    Dim lng����ID  As Long
    
    'ת���ɴ�д(���ֲ��ɴ���)
    If KeyAscii > 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
    '��һλ����������ĸ,����λ����
    If KeyAscii <> 13 Then
        Call SetNOInputLimit(cboNO, KeyAscii)
    ElseIf cboNO.Text <> "" And Not cboNO.Locked Then
        cboNO.Text = GetFullNO(cboNO.Text, 15)
        
        '�Ƿ���ת������ݱ���
        If zlDatabase.NOMoved("���˽��ʼ�¼", cboNO.Text, , , Me.Caption) Then
            If Not ReturnMovedExes(cboNO.Text, 7, Me.Caption) Then Exit Sub
            mblnNOMoved = False
        End If
        
        '����Ȩ��
        If Not ReadBillInfo(2, cboNO.Text, -1, strOper, vDate) Then
            cboNO.Text = "": If cboNO.Visible And cboNO.Enabled Then cboNO.SetFocus
            Exit Sub
        End If
        
        If Not BillOperCheck(7, strOper, vDate, "����") Then
            cboNO.Text = "": If cboNO.Visible And cboNO.Enabled Then cboNO.SetFocus
            Exit Sub
        End If
        
        'lng����ID:49084
        mYBInFor.intInsure = BalanceExistInsure(cboNO.Text, bytFlag, lng����ID)
        mYBInFor.bytMCMode = bytFlag
        If mYBInFor.intInsure <> 0 Then
            '���ս���Ȩ���ж�
            If InStr(mstrPrivs, ";���ս���;") = 0 Then
                MsgBox "��û��Ȩ�����ϱ��ղ��˵Ľ��ʵ��ݡ�", vbInformation, gstrSysName
                Exit Sub
            End If
            MCPAR.�ֱҴ��� = gclsInsure.GetCapability(support�ֱҴ���, lng����ID, mYBInFor.intInsure)
            If mYBInFor.bytMCMode = 1 Then
                MCPAR.���ﲡ�˽������� = gclsInsure.GetCapability(support�����������, lng����ID, mYBInFor.intInsure)
            Else
                MCPAR.��Ժ���˽������� = gclsInsure.GetCapability(support��Ժ���˽�������, lng����ID, mYBInFor.intInsure)
            End If
            MCPAR.�������Ϻ��ӡ�ص� = gclsInsure.GetCapability(support�������Ϻ��ӡ�ص�, lng����ID, mYBInFor.intInsure)
        Else
            If InStr(mstrPrivs, ";��ͨ���˽���;") = 0 Then
                MsgBox "��û��Ȩ��������ͨ���˵Ľ��ʵ��ݡ�", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        If PatiErrBillPay(0, cboNO.Text) = False Then
            If cboNO.Enabled And cboNO.Visible Then cboNO.SetFocus
            Exit Sub
        End If
        If CheckExistsGathering(cboNO.Text) Then
            MsgBox "�ý��ʵ��ݴ����ѽɿ��Ӧ�տ��¼�����˿����ִ�����ϡ�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If CheckBillBeforIN(cboNO.Text) Then
            If MsgBox("�ý��ʵ��Ǳ���סԺ֮ǰ�����ģ���ȷ��Ҫ���ϸõ�����?", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then
                Exit Sub
            End If
        End If
        
        If InStr(mstrPrivs, ";��������;") = 0 Then
             MsgBox "��û��Ȩ�����Ͻ��ʵ��ݡ�", vbInformation, gstrSysName
             Exit Sub
        End If
        
        '��ȡҪ���ϵĽ��ʵ�
        If Not ReadBalance(cboNO.Text, True) Then
            cboNO.Text = "": If cboNO.Visible And cboNO.Enabled Then cboNO.SetFocus
        Else
            Call Set�Ҳ���Ϣ
            If cmdOK.Visible And cmdOK.Enabled Then cmdOK.SetFocus
        End If
    Else
           If InStr(mstrPrivs, ";��ͨ���˽���;") = 0 Then
                MsgBox "��û��Ȩ�����ϷǱ��ղ��˵Ľ��ʵ��ݡ�", vbInformation, gstrSysName
                Exit Sub
           End If
    End If
End Sub

Private Sub cboFeeItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cboFeeItem_NodeCheck(ByVal Node As MSComctlLib.Node, strCaption As String)
   If mblnConsChange Then Exit Sub
   With mobjBalanceCon
        mblnConsChange = .strItem <> cboFeeItem.GetNodesCheckedDatas()
   End With
   Call ClearListData
End Sub

Private Sub cboFeeType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
 
Private Sub cboFeeType_NodeCheck(ByVal Node As MSComctlLib.Node, strCaption As String)
    If mblnConsChange Then Exit Sub
    With mobjBalanceCon
         mblnConsChange = .strClass <> cboFeeType.GetNodesCheckedDatas()
    End With
    Call ClearListData
End Sub

Private Sub cboPatiNums_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
Private Sub cboPatiNums_NodeCheck(ByVal Node As MSComctlLib.Node, strCaption As String)

    Dim i As Integer, intMaxTime As Integer, intNum As Integer, arrNum As Variant
    Dim strAllSelTime As String, strNodesChecked As String
    Dim intInsure As Integer, strInsureName As String
   
    strNodesChecked = zlGetAllTims(cboPatiNums.GetNodesCheckedDatas) '���е�סԺ�������ؿ�
    strAllSelTime = zlGetAllTims(cboPatiNums.GetNodesCheckedDatas(False))
    
    If strAllSelTime <> "" Then
        arrNum = Split(strAllSelTime, ",")
        intNum = Val(arrNum(0))
    End If
    
    If Not mrsInfo Is Nothing Then
        Call SetPatiNums(intNum)
    End If
    Call LoadDefaultOutStatu(mPatiInfor.lng����ID, intNum)
    
    If Not mblnNotChange Then
        mblnNotChange = True
        Call RecalcFeeTotalDate
        mblnNotChange = False
    End If
    
    mblnConsChange = mobjBalanceCon.strTime <> strNodesChecked
    Call ClearListData
End Sub

Private Sub SetPatiNums(ByVal intTime As Integer)
 
    Dim blnInsure As Boolean, i As Integer
    Dim intInsure As Integer, strInsureName As String
    
    If intTime = 0 Then Exit Sub
    If mobjBalanceCon.blnCurBalanceOwnerFee Then Exit Sub
    If mblnConsChange Then Exit Sub
    
    If mEditType <> g_Ed_������� Then
        '���ò���ҽ��״̬
        blnInsure = False
        mobjBalanceAll.rsAllTime.Filter = "��ҳID=" & intTime
        mYBInFor.intInsure = 0
        If Not mobjBalanceAll.rsAllTime.EOF Then
            mYBInFor.intInsure = Val(mobjBalanceAll.rsAllTime!����):  blnInsure = True
            mYBInFor.strBalance = ""
            mobjBalanceAll.rsAllTime.Filter = ""
        End If
        mblnCurPatiInsure = blnInsure
    End If
End Sub

Private Sub ClearListData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����б�����
    '����:���˺�
    '����:2015-02-05 18:13:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim tyBalance As TY_Balance_Infor
    
    If Not mblnConsChange Then Exit Sub
    mblnNotChange = True
    txtBalance(Idx_���ν���).Text = ""
    txtBalance(Idx_����δ��).Text = ""
    mblnNotChange = False
    
    txtBalance(Idx_����˵��).Text = ""
    txtBalance(Idx_�ɿ�).Text = ""
    txtBalance(Idx_�Ҳ�).Text = ""
    txtBalance(Idx_�������).Text = ""
    txtBalance(Idx_ժҪ).Text = ""
    txtBalance(Idx_��Ԥ��).Text = ""
    
    Set mrsFeeList = Nothing
    Set mrsBalance = Nothing
    mBalanceInfor = tyBalance
    Call ClearFeeList   '������б�
    Call ClearAdjustBalance '��������б�
    Call ClearAdjustDeposit  '���Ԥ���б�
    Call InitPatiBalanceVariableCon
    Call SetOperationCtrl(3)
    Call LoadCurOwnerPayInfor
End Sub
 

Private Sub cboPatiNums_NodeCheckValied(ByVal Node As MSComctlLib.Node, blnCancel As Boolean)
    Dim objNode As MSComctlLib.Node
    Dim varTemp As Variant, str��ҳIds As String, lng����ID As Long
    Dim int��ҳID As Integer, intInsure As Integer, strInsureName As String
    Dim int��ҳID1 As Integer, intInsure1 As Integer, strInsureName1 As String
    Dim blnFirst As Boolean
    If mrsInfo Is Nothing Then blnCancel = True: Exit Sub
    If mrsInfo.State <> 1 Then blnCancel = True: Exit Sub
    
    lng����ID = Val(NVL(mrsInfo!����ID))
    'ѡ��鵱ǰ�ڵ����Ч��
    '��ҳID|����|��������
    str��ҳIds = cboPatiNums.GetNodesCheckedDatas(False)
    
    If str��ҳIds = "" Then 'Ϊ��ʱ������ѡ��һ��
        
        blnCancel = True: Exit Sub
    End If
    varTemp = Split(str��ҳIds, ",")
    
    
    If zlGetTimeDataFromTimes(varTemp(0), int��ҳID, intInsure, strInsureName) = False Then
         blnCancel = True: Exit Sub
    End If
    If intInsure <> 0 Then Call InitInsurePara(lng����ID, intInsure)

    If Node.Key = "Root" Then '��ǰ����������סԺ���ڵ�
        If Node.Checked Then     'ѡ������
            blnFirst = True
            'If intInsure = 0 Then Exit Sub '��ҽ�������Էѽ��н���
            For Each objNode In cboPatiNums.Nodes
                If objNode.Key <> "Root" Then
                    If zlGetTimeDataFromTimes(objNode.Tag, int��ҳID1, intInsure1, strInsureName1) Then
                        If blnFirst Then
                            intInsure = int��ҳID1: intInsure = intInsure1: strInsureName = strInsureName1
                            If intInsure <> 0 Then Call InitInsurePara(lng����ID, intInsure)
                            If MCPAR.�������סԺ���� Then Exit Sub
                        Else
                            If intInsure <> 0 Then
                               Node.Checked = False
                               If int��ҳID1 <> int��ҳID Then objNode.Checked = False
                            End If
                        End If
                    End If
                    blnFirst = False
                End If
            Next
            Exit Sub
        End If
        blnCancel = True: Exit Sub '����һ������ѡ
    End If
    
    If zlGetTimeDataFromTimes(Node.Tag, int��ҳID1, intInsure1, strInsureName1) = False Then blnCancel = True: Exit Sub
      
    If Node.Checked Then      '��ǰ��ѡ
        If int��ҳID1 = int��ҳID Then   '��ǰѡ�еģ����ǵ�һ��ѡ���
            If intInsure = 0 Or MCPAR.�������סԺ���� Then Exit Sub '��ҽ���ģ��������Էѻ�������סԺһ�νᣬ��ȫ����ҽ������
            'ֻ��ѡ���һ�ε�סԺ��
            For Each objNode In cboPatiNums.Nodes
                If zlGetTimeDataFromTimes(objNode.Tag, int��ҳID1, intInsure1, strInsureName1) = False Then blnCancel = True: Exit Sub
                If int��ҳID <> int��ҳID1 Or objNode.Key = "Root" Then objNode.Checked = False
            Next
            Exit Sub
        End If
        '�϶�����ѡ��ĵ�һ��
        If intInsure = 0 Or MCPAR.�������סԺ���� Then Exit Sub '��ҽ���ģ��������Էѻ�������סԺһ�νᣬ��ȫ����ҽ������
        
        If zlGetTimeDataFromTimes(Node.Tag, int��ҳID1, intInsure1, strInsureName1) = False Then blnCancel = True: Exit Sub
        If intInsure <> 0 And intInsure1 = 0 Then
           '���ԭ��ѡ��ģ��������ѡ���Ϊ׼
           int��ҳID = int��ҳID1: intInsure = intInsure1: strInsureName = strInsureName1
            For Each objNode In cboPatiNums.Nodes
                If zlGetTimeDataFromTimes(objNode.Tag, int��ҳID1, intInsure1, strInsureName1) = False Then blnCancel = True: Exit Sub
                If int��ҳID <> int��ҳID1 Or objNode.Key = "Root" Then objNode.Checked = False
            Next
            Exit Sub
        
        End If
        
        For Each objNode In cboPatiNums.Nodes
            If zlGetTimeDataFromTimes(objNode.Tag, int��ҳID1, intInsure1, strInsureName1) = False Then blnCancel = True: Exit Sub
            If int��ҳID <> int��ҳID1 Or objNode.Key = "Root" Then objNode.Checked = False
        Next
        Exit Sub
    Else
         If intInsure <> 0 And MCPAR.�������סԺ���� = False Then '��һ����ҽ��������Ҫ���Ų���
            For Each objNode In cboPatiNums.Nodes
                If objNode.Key <> Node.Key Then
                    If zlGetTimeDataFromTimes(objNode.Tag, int��ҳID1, intInsure1, strInsureName1) = False Then blnCancel = True: Exit Sub
                    If int��ҳID <> int��ҳID1 Or objNode.Key = "Root" Then objNode.Checked = False
                End If
            Next
            Exit Sub
         End If
    End If
    
    '��ǰѡ���ֻ��һ�����Ҿ��ǵ�ǰ���
    If UBound(varTemp) = 0 Then
        If int��ҳID = int��ҳID1 Then blnCancel = True: Exit Sub   '������ȡ��������ѡ��һ��
    End If
End Sub

Private Sub cboӤ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
Private Sub cboӤ��_NodeCheck(ByVal Node As MSComctlLib.Node, strCaption As String)
    If mblnConsChange Then Exit Sub
    mblnConsChange = mobjBalanceCon.strBaby <> cboӤ��.GetNodesCheckedDatas
    Call ClearListData
End Sub


Private Sub ExecuteFeeQuery(ByVal lngControlID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ�з��ò�ѯ
    '���:lngControlID-�˵��ؼ���ID
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-02-12 10:33:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long
    Dim objCon As clsBalanceCon
    Dim EditType As gBalanceBill
    
    If (mblnConsChange Or mrsInfo Is Nothing) And (mEditType = g_Ed_������� Or mEditType = g_Ed_סԺ����) And chkCancel.Value = 0 Then
        MsgBox "��ǰ�����ڽ��ʷ���,����!", vbInformation, gstrSysName
        Exit Sub
    End If
    Set objCon = New clsBalanceCon
    With objCon
        .blnCurBalanceOwnerFee = mobjBalanceCon.blnCurBalanceOwnerFee
        .strBaby = mobjBalanceCon.strBaby
        .strChargeType = mobjBalanceCon.strChargeType
        .lng����ID = IIf(mobjBalanceCon.lng����ID = 0, mPatiInfor.lng����ID, mobjBalanceCon.lng����ID)
        .bytKind = mobjBalanceCon.bytKind
        .dtBeginDate = mobjBalanceCon.dtBeginDate
        .dtEndDate = mobjBalanceCon.dtEndDate
        .strClass = mobjBalanceCon.strClass
        .strDeptIDs = mobjBalanceCon.strDeptIDs
        .strItem = mobjBalanceCon.strItem
        .strTime = mobjBalanceCon.strTime
    End With
    lng����ID = IIf(mBalanceInfor.lng����ID <> 0, mBalanceInfor.lng����ID, mBalanceInfor.lng����ID)
    
    If chkCancel.Value = 1 Then
        EditType = g_Ed_��������
    Else
        EditType = mEditType
    End If
    
    Select Case lngControlID
    Case conMenu_View_List ' "��ϸ��"
        Call frmBalanceQuery.ShowMe(Me, EditType, objCon, lng����ID, mlngModul, mstrPrivs, g_Ed_��ϸ��)
    Case conMenu_View_ListItem ' "��Ŀ��ϸ"
        Call frmBalanceQuery.ShowMe(Me, EditType, objCon, lng����ID, mlngModul, mstrPrivs, g_Ed_��Ŀ��ϸ)
    Case conMenu_View_SplitType ' "�����"
        Call frmBalanceQuery.ShowMe(Me, EditType, objCon, lng����ID, mlngModul, mstrPrivs, g_Ed_�����)
    Case conMenu_View_SplitMonth ' "���±�"
        Call frmBalanceQuery.ShowMe(Me, EditType, objCon, lng����ID, mlngModul, mstrPrivs, g_Ed_���±�)
    Case conMenu_View_DayBill ' "���յ���"
        Call frmBalanceQuery.ShowMe(Me, EditType, objCon, lng����ID, mlngModul, mstrPrivs, g_Ed_���յ���)
    Case conMenu_View_DayFM ' "���շ�Ŀ"
        Call frmBalanceQuery.ShowMe(Me, EditType, objCon, lng����ID, mlngModul, mstrPrivs, g_Ed_���շ���)
    Case conMenu_View_Balance '���ʱ�
        Call frmBalanceQuery.ShowMe(Me, EditType, objCon, lng����ID, mlngModul, mstrPrivs, g_Ed_���ʱ�)
    End Select
End Sub
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim dblMoney As Double
    Dim bytSetFocus As Byte '1-Ԥ��;0-�ɿ�
    'ִ�в���
    Select Case Control.ID
    Case conMenu_View_List ' "��ϸ��"
        Call ExecuteFeeQuery(Control.ID)
    Case conMenu_View_ListItem ' "��Ŀ��ϸ"
        Call ExecuteFeeQuery(Control.ID)
    Case conMenu_View_SplitType ' "�����"
        Call ExecuteFeeQuery(Control.ID)
    Case conMenu_View_SplitMonth ' "���±�"
        Call ExecuteFeeQuery(Control.ID)
    Case conMenu_View_DayBill ' "���յ���"
        Call ExecuteFeeQuery(Control.ID)
    Case conMenu_View_DayFM ' "���շ�Ŀ"
        Call ExecuteFeeQuery(Control.ID)
    Case conMenu_View_Balance '��ϸ��
        Call ExecuteFeeQuery(Control.ID)
    Case conMenu_Edit_ClinicToHos
        If InStr(1, mstrPrivs, ";�������תסԺ;") = 0 Then Exit Sub
        If mobjInPati Is Nothing Then
            Err = 0: On Error Resume Next
            Set mobjInPati = CreateObject("zl9InPatient.clsInPatient")
            
            If Err <> 0 Then
                MsgBox "ע��:" & vbCrLf & "   סԺ���˲���(zl9InPatient)����ʧ��,����ϵͳ����Ա��ϵ!"
                Exit Sub
            End If
        End If
        Call mobjInPati.zlOutFeeToInFee(Me, gcnOracle, glngSys, mlngModul, mstrPrivs, gstrDBUser, mobjBalanceCon.lng����ID, 0)
    Case conMenu_Edit_NotUseDeposit   '��ʹ��Ԥ����(C)
        '0-������г�Ԥ��;1-��ȱʡʹ��Ԥ����;2-��ָ���������Ԥ��(��ʱ���Ⱥ�����̯��;3-ȫ��
        Call RecalcDepositMoney(0): mbln�ѱ��� = False: GoTo GoFullDeposit:
        bytSetFocus = 0
    Case conMenu_Edit_MoneyUseDeposit   '�����ʽ��ʹ��Ԥ��(L)
        Call RecalcDepositMoney(0)
        mblnNotChange = True
        txtBalance(Idx_��Ԥ��).Text = Format(mBalanceInfor.dbl��Ԥ���ϼ�, "0.00")
        txtBalance(Idx_��Ԥ��).BackColor = &H80000005
        mBalanceInfor.blnԤ����֤ = False
        mBalanceInfor.blnԤ��ˢ�� = False
        mblnNotChange = False
        Call LoadCurOwnerPayInfor
        If bytSetFocus = 1 Then
            If txtBalance(Idx_��Ԥ��).Enabled And txtBalance(Idx_��Ԥ��).Visible Then txtBalance(Idx_��Ԥ��).SetFocus
        Else
            Call txtBalance_Validate(Idx_��Ԥ��, False)
            If txtBalance(Idx_�ɿ�).Enabled And txtBalance(Idx_�ɿ�).Visible Then txtBalance(Idx_�ɿ�).SetFocus
        End If
        dblMoney = RoundEx(mBalanceInfor.dblδ���ϼ�, 6)
        bytSetFocus = 1
        Call RecalcDepositMoney(2, dblMoney): mbln�ѱ��� = False: GoTo GoFullDeposit:
    Case conMenu_Edit_UseAllDeposit   'ʹ������Ԥ����(A)
        bytSetFocus = 0
        Call RecalcDepositMoney(3): mbln�ѱ��� = False: GoTo GoFullDeposit:
    Case conMenu_File_Exit: Unload Me '�˳�
    Case Else
    End Select
    Exit Sub
GoFullDeposit:
    mblnNotChange = True
    txtBalance(Idx_��Ԥ��).Text = Format(mBalanceInfor.dbl��Ԥ���ϼ�, "0.00")
    txtBalance(Idx_��Ԥ��).BackColor = &H80000005
    mBalanceInfor.blnԤ����֤ = False
    mBalanceInfor.blnԤ��ˢ�� = False
    mblnNotChange = False
    Call LoadIntendBalance
    Call LoadCurOwnerPayInfor
    If bytSetFocus = 1 Then
        If txtBalance(Idx_��Ԥ��).Enabled And txtBalance(Idx_��Ԥ��).Visible Then txtBalance(Idx_��Ԥ��).SetFocus
    Else
        Call txtBalance_Validate(Idx_��Ԥ��, False)
        If txtBalance(Idx_�ɿ�).Enabled And txtBalance(Idx_�ɿ�).Visible Then txtBalance(Idx_�ɿ�).SetFocus
    End If
    
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Err = 0: On Error Resume Next
    Top = txtPatient.Top - 60
    If Me.staThis.Visible Then Bottom = Me.staThis.Height
    staThis.Top = Me.ScaleHeight - Me.staThis.Height
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    'ִ�в���
    Select Case Control.ID
    Case conMenu_View_Balance   '���ʱ�
        Control.Enabled = Not mblnLockScreen
    Case conMenu_View_List      '��ϸ��
    Case conMenu_View_SplitType '�����
    Case conMenu_View_SplitMonth   '���±�
    Case conMenu_View_DayBill   '���յ���
    Case conMenu_View_DayBill   '���յ���
    Case conMenu_View_DayFM '���շ�Ŀ
    Case conMenu_File_Exit '�˳�
        If mEditType <> g_Ed_���ݲ鿴 Then
            Control.Visible = Not mBalanceInfor.blnSaveBill
        End If
    Case conMenu_Edit_ClinicToHos
        Control.Visible = mEditType = g_Ed_סԺ����
    Case Else
    End Select
End Sub

Private Sub chkCancel_Click()
    Dim blnNew As Boolean
    
    If mblnNotChange Then Exit Sub

    If mBalanceInfor.blnSaveBill = True Then
        MsgBox "�Ѿ������˽��ʵ���,������ɵ�ǰ�������л�����ģʽ!", vbInformation, gstrSysName
        mblnNotChange = True
        chkCancel.Value = 0
        mblnNotChange = False
        Exit Sub
    End If
    
    If chkCancel.Value = 1 Then
        chkCancel.ForeColor = &HFF&
        lblCurPaymentInfor.Caption = "��ǰ�˿����"
        txtԤ�����.Visible = False
        txtδ�����.Visible = False
        txtʣ����.Visible = False
        lblԤ�����.Visible = False
        lblδ�����.Visible = False
        lblʣ����.Visible = False
    Else
        chkCancel.ForeColor = &H80000012
        lblCurPaymentInfor.Caption = "��ǰ֧�����"
        txtԤ�����.Visible = True
        txtδ�����.Visible = True
        txtʣ����.Visible = True
        lblԤ�����.Visible = True
        lblδ�����.Visible = True
        lblʣ����.Visible = True
    End If
    
    blnNew = chkCancel.Value = 0
    If blnNew Then cboNO.Text = "": mstrInNO = ""
    mstrForceNote = ""
    
    txtInvoice.Locked = Not blnNew Or Not (zlStr.IsHavePrivs(mstrPrivs, "�޸�Ʊ�ݺ�")) And gblnStrictCtrl '89302
    cboNO.Locked = blnNew
      
    If mblnNotChange Then Exit Sub
      
    If mty_ModulePara.bln���ʺ�����Ϣ And mblnNotClearBill Then
        If mrsInfo Is Nothing Then
            Call LoadBalanceBill
            mblnNotClearBill = False
        ElseIf mrsInfo.State <> 1 Then
            Call LoadBalanceBill     '���е�InitBalanceSet������һЩ�ؼ�״̬
             mblnNotClearBill = False
        End If
    Else
        Call LoadBalanceBill     '���е�InitBalanceSet������һЩ�ؼ�״̬
    End If
    Call SetFeeListColumnShow
    Call SetPatiConsControlVisible
    Call SetOperatonCommandCaption
    Call SetControlEnabled(chkCancel.Value <> 1)
    If Not blnNew Then
        If cboNO.Visible And cboNO.Enabled Then cboNO.SetFocus
        zlControl.TxtSelAll cboNO
    Else
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        zlControl.TxtSelAll txtPatient
    End If
End Sub

Private Sub chkDeposit_Click()
    Dim blnSearch As Boolean, strInfor As String
    Dim strMsg As String
    Dim strCardTypes As String
    Dim strTemp As String, str������ˮ�� As String, str����˵�� As String, str���� As String
    Dim lngCardTypeID As Long, bln���ѿ� As Boolean
    Dim j As Long, i As Long
    Dim dblMoney As Double, strXMLExpend As String, str����Ա���� As String
    
    
    If mblnNotChange Then Exit Sub
    
    If Not (mEditType = g_Ed_�������� Or mEditType = g_Ed_�������� Or chkCancel.Value = 1) Then Exit Sub
    
    If chkDeposit.Value = 0 Then
        strMsg = ""
        With vsDeposit
            strCardTypes = ""
            For i = 1 To .Rows - 1
                lngCardTypeID = Val(.TextMatrix(i, .ColIndex("�����ID"))): bln���ѿ� = Val(.TextMatrix(i, .ColIndex("�Ƿ����ѿ�")))
                strTemp = lngCardTypeID & ":" & Val(.TextMatrix(i, .ColIndex("�Ƿ����ѿ�")))
                
                If InStr(strCardTypes & ",", "," & strTemp & ",") = 0 And Val(.TextMatrix(i, .ColIndex("�Ƿ�����"))) = 0 And lngCardTypeID <> 0 Then
                    
                   dblMoney = Val(.TextMatrix(i, .ColIndex("��Ԥ��")))
                    If Not bln���ѿ� Then
                        str������ˮ�� = .TextMatrix(i, .ColIndex("������ˮ��"))
                        str����˵�� = .TextMatrix(i, .ColIndex("����˵��"))
                        str���� = .TextMatrix(i, .ColIndex("����"))
                        If zlCallReturnCashCheckInterface(Me, mlngModul, lngCardTypeID, str����, "1|" & Val(.TextMatrix(i, .ColIndex("Ԥ��ID"))), dblMoney, str������ˮ��, str����˵��) = False Then
                            mblnNotChange = True: chkDeposit.Value = 1: mblnNotChange = False
                            Exit Sub
                        End If
                    End If
                    For j = i + 1 To .Rows - 1
                        If strTemp = Val(.TextMatrix(j, .ColIndex("�����ID"))) & ":" & Val(.TextMatrix(j, .ColIndex("�Ƿ����ѿ�"))) Then
                            dblMoney = dblMoney + Val(.TextMatrix(j, .ColIndex("��Ԥ��")))
                            If Not bln���ѿ� Then
                                str������ˮ�� = .TextMatrix(j, .ColIndex("������ˮ��"))
                                str����˵�� = .TextMatrix(j, .ColIndex("����˵��"))
                                str���� = .TextMatrix(j, .ColIndex("����"))
                                
                                If zlCallReturnCashCheckInterface(Me, mlngModul, lngCardTypeID, str����, "1|" & Val(.TextMatrix(j, .ColIndex("Ԥ��ID"))), Val(.TextMatrix(j, .ColIndex("��Ԥ��"))), str������ˮ��, str����˵��) = False Then
                                    mblnNotChange = True: chkDeposit.Value = 1: mblnNotChange = False
                                    Exit Sub
                                End If
                            End If
                            
                        End If
                    Next
                    strCardTypes = strCardTypes & "," & strTemp
                    strMsg = strMsg & vbCrLf & .TextMatrix(i, .ColIndex("���������")) & ":" & Format(dblMoney, gstrDec) & "Ԫ"
                End If
            Next
        End With
        
        If strMsg <> "" And chkDeposit.Value = 0 Then
            strMsg = Mid(strMsg, 2)
            If MsgBox("���������ʻ��в��������֣��Ƿ�ǿ�ƽ������֣�" & vbCrLf & strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) <> vbYes Then
                mblnNotChange = True: chkDeposit.Value = 1: mblnNotChange = False
                 Exit Sub
            End If
            '����֧�����ֵ����
            If InStr(";" & mstrPrivsCard & ";", ";�����˿�ǿ������;") = 0 Then
                str����Ա���� = zlDatabase.UserIdentifyByUser(Me, "ǿ��������֤", glngSys, 1151, "�����˿�ǿ������")
                If str����Ա���� = "" Then
                    MsgBox "¼��Ĳ���Ա��֤ʧ�ܻ���¼��Ĳ���Ա���߱�ǿ������Ȩ�ޣ�����ǿ�����֣�", vbInformation, gstrSysName
                    mblnNotChange = True: chkDeposit.Value = 1: mblnNotChange = False
                    Exit Sub
                End If
                mstrForceNote = str����Ա���� & "ǿ������:" & Replace(strMsg, vbCrLf, ";")
            Else
                mstrForceNote = UserInfo.���� & "ǿ������:" & Replace(strMsg, vbCrLf, ";")
            End If
            
        End If
 
    End If
    
    If chkDeposit.Value = 1 Then
        txtBalance(Idx_��Ԥ��).Text = Format(Val(chkDeposit.Tag), "0.00")
        mBalanceInfor.dbl��Ԥ���ϼ� = Val(chkDeposit.Tag)
     Else
        txtBalance(Idx_��Ԥ��).Text = "0.00"
        mBalanceInfor.dbl��Ԥ���ϼ� = 0
    End If
    Call Set�Ҳ���Ϣ
    Call LoadCurOwnerPayInfor
    If txtBalance(Idx_�ɿ�).Enabled And txtBalance(Idx_�ɿ�).Visible Then
        txtBalance(Idx_�ɿ�).SetFocus
        zlControl.TxtSelAll txtBalance(Idx_�ɿ�)
    End If
    
End Sub

Private Sub cmdCancel_Click()
 
    'ȡ������
    If mintPreEditType <> -1 Then
        mEditType = mintPreEditType '�ָ��ϴβ���
        Call NewBill
        If picPati.Enabled And txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        zlControl.TxtSelAll txtPatient
        mintPreEditType = -1
        Exit Sub
    End If
    
    If mEditType = g_Ed_���ݲ鿴 _
        Or mEditType = g_Ed_ȡ������ _
        Or mEditType = g_Ed_�������� _
        Or mEditType = g_Ed_�������� _
        Or mEditType = g_Ed_���½��� Then
        '�˳�
        Unload Me: Exit Sub
    End If
    
    If mEditType = g_Ed_������� Or mEditType = g_Ed_סԺ���� Then
        If mblnNotClearBill And mty_ModulePara.bln���ʺ�����Ϣ Then
            '��ǰΪ���ʲ����Ʊ��,��ȡ��ʱ,���
             If mrsInfo Is Nothing Then
                Call NewBill: mblnNotClearBill = False: Exit Sub
             End If
             If mrsInfo.State <> 1 Then
                Call NewBill: mblnNotClearBill = False: Exit Sub
             End If
        End If
        
        If chkCancel.Value = Checked And txtPatient.Text <> "" Then
           '��ǰ����Ϊ���� ,����ʾ�Ƿ�Ҫ�˳�
            If MsgBox("ȷʵҪ�˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            Unload Me: Exit Sub
        End If
        
        '�Ѿ���֤ҽ���Ĳ���
        If mYBInFor.bytMCMode = 1 Then
            If MsgBox("ȷʵҪȡ����ǰ���������֤��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Sub
            End If
            If YBIdentifyCancel Then Call NewBill
            If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
            Exit Sub
            '���˳�����,�Ա�ѡ���������˽��������֤
        End If
        If Not mrsInfo Is Nothing Then
            If Val(txtBalance(Idx_���ν���).Text) <> 0 And mrsInfo.State = adStateOpen Then
                If MsgBox("�ò�����δȷ������,ȷʵȡ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
                Call NewBill
                If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
                Exit Sub
            End If
        End If
        If txtPatient.Text <> "" Then
            If MsgBox("ȷʵҪ�˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
    End If
    Unload Me
End Sub

Private Sub cmdColSet_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    If Button <> 1 Then Exit Sub
    vRect = zlControl.GetControlRect(cmdColSet.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + cmdColSet.Height
    Call Grid.SetColVisible(Me, Me.Caption, vsDeposit, lngLeft, lngTop, cmdColSet.Height)
    zl_vsGrid_Para_Save mlngModul, vsDeposit, Me.Name, "Ԥ���б�"
End Sub

Private Sub SetBalanceConInfor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���õ�ǰ����������Ϣ
    '����:���˺�
    '����:2015-01-07 18:30:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnCurOwner As Boolean
    On Error GoTo errHandle
    
    Set mobjBalanceCon = New clsBalanceCon
    Call InitPatiBalanceVariableCon
    With mobjBalanceCon
        .strTime = zlGetAllTims(cboPatiNums.GetNodesCheckedDatas())     '���˽��ʴ���(��ʼ="",����Ϊ"1,2,3...")
        .strDeptIDs = cboDept.GetNodesCheckedDatas()    '���˽��ʿ���ID��(��ʼ="",����Ϊ"0,1,2,3...",0��ʾ��������IDΪ��)
        .strClass = cboFeeType.GetNodesCheckedDatas()   '��������=""-���з���(��δ����),"'����','����',..."
        .strChargeType = GetChargeTypeCheckDatas(blnCurOwner)     '�շ���� '34260
        .strItem = cboFeeItem.GetNodesCheckedDatas()   'Ҫ����վݷ�Ŀ
        .strBaby = cboӤ��.GetNodesCheckedDatas()  '�Ƿ������Ӥ������(0-���з���,1-���˷���,2������-��mbytbaby-1��Ӥ������)
        If chkPatiType(CK_Idx_��ͨ).Value = 1 And chkPatiType(CK_Idx_���).Value <> 1 Then
            .bytKind = 0
        ElseIf chkPatiType(CK_Idx_��ͨ).Value <> 1 And chkPatiType(CK_Idx_���).Value = 1 Then
              .bytKind = 1
        Else
            .bytKind = 2  '0-����ͨ����,1-��������,2-��ͨ���ú�������
        End If
        If IsDate(txtBegin.Text) Then   '���˽��ʵĿ�ʼʱ��
            .dtBeginDate = Format(CDate(txtBegin.Text & " 00:00:00"), "yyyy-mm-dd HH:MM:SS")
        End If
        If IsDate(txtEnd.Text) Then   '���˽��ʵĽ���ʱ��
            .dtEndDate = Format(CDate(txtEnd.Text & " 00:00:00"), "yyyy-mm-dd HH:MM:SS")
        End If
        .blnCurBalanceOwnerFee = blnCurOwner
        .lng����ID = mPatiInfor.lng����ID
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Function SaveDeposit(ByRef blnԤ�� As Boolean, Optional ByVal blnNoRecal As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ԥ����
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-27 11:00:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnChargeEnd As Boolean, objSetFocus As Object
    Dim tyBrushCard As TY_BrushCard
       
    On Error GoTo errHandle
    
    If mEditType = g_Ed_�������� Or mEditType = g_Ed_�������� _
    Or chkCancel.Value = 1 And chkCancel.Visible Then Exit Function
    
    If mBalanceInfor.blnԤ����֤ Then
        If txtBalance(Idx_�ɿ�).Enabled And txtBalance(Idx_�ɿ�).Visible Then
            txtBalance(Idx_�ɿ�).SetFocus
            zlControl.TxtSelAll txtBalance(Idx_�ɿ�)
        End If
        Exit Function
    End If
    
    Screen.MousePointer = 99
    mblnNotChange = True
    LockedScreen True
    mblnNotChange = False
 

    '���ж��Ƿ����Ԥ����ˢ���ģ����ȴ���Ԥ����
    If Not CheckDepositValied(blnԤ��) Then
        LockedScreen False
        Set objSetFocus = txtBalance(Idx_��Ԥ��)
        If Not objSetFocus Is Nothing Then
            If objSetFocus.Enabled And objSetFocus.Visible Then objSetFocus.SetFocus
        End If
        zlControl.TxtSelAll objSetFocus
        Screen.MousePointer = 0
        Exit Function
    End If
    If Not blnԤ�� Then
        Screen.MousePointer = 0
        LockedScreen False
        SaveDeposit = True: Exit Function
    End If
    
    If Not SaveBalaceCharge(True, tyBrushCard, blnChargeEnd, objSetFocus) Then
        LockedScreen False
        If Not objSetFocus Is Nothing Then
            If objSetFocus.Enabled And objSetFocus.Visible Then
                objSetFocus.SetFocus
            End If
        End If
        Screen.MousePointer = 0
        Exit Function
    End If

    If blnChargeEnd And mEditType = g_Ed_���½��� Then Unload Me: Exit Function
    
    LockedScreen False
 
    '0-����ǰ;1-ҽ����������;2-�ѱ����˽��ʵ�;
    If mBalanceInfor.blnSaveBill Then
       Call SetOperationCtrl(2)
    Else
       Call SetOperationCtrl(0)
    End If
    
    If blnChargeEnd Then
        Call NewBill
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        SaveDeposit = True
        Exit Function
    End If
    
    If Not blnNoRecal Then Call LoadIntendBalance
    Call LoadCurOwnerPayInfor
    If txtBalance(Idx_�ɿ�).Enabled And txtBalance(Idx_�ɿ�).Visible Then
        mbln�ѱ��� = False
        txtBalance(Idx_�ɿ�).SetFocus
    End If
    
    SaveDeposit = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    LockedScreen False
End Function

Private Function CheckInputValied() As Boolean
    On Error GoTo errHandle
    If zlCommFun.ActualLen(txtBalance(Idx_ժҪ).Text) > 50 Then
        MsgBox "�����ժҪ���ܳ���25�����ֻ�50���ַ�,����!", vbInformation + vbOKOnly, gstrSysName
        If txtBalance(Idx_ժҪ).Enabled And txtBalance(Idx_ժҪ).Visible Then txtBalance(Idx_ժҪ).SetFocus
        Exit Function
    End If
    
    If zlCommFun.ActualLen(txtBalance(Idx_�������).Text) > 30 Then
        MsgBox "����Ľ�����벻�ܳ���15�����ֻ�30���ַ�,����!", vbInformation + vbOKOnly, gstrSysName
        If txtBalance(Idx_�������).Enabled And txtBalance(Idx_�������).Visible Then txtBalance(Idx_�������).SetFocus
        Exit Function
    End If
    
    CheckInputValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function SaveBalanceData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����������
    '����:����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-30 09:44:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnԤ�� As Boolean, tyBrushCard As TY_BrushCard
    Dim blnChargeEnd As Boolean, blnFind As Boolean
    Dim objSetFocus As Object, blnSaved As Boolean
    Dim strErrMsg As String, i As Long
    Dim blnNotClearPati As Boolean
    Dim blnHaveFee As Boolean
    Dim objCard As Card
    Dim blnCancel As Boolean
    Dim bln�������۲��� As Boolean, lng����ID As Long, str���� As String, dbl������� As Double
    
    On Error GoTo errHandle
    If mEditType = g_Ed_ȡ������ Then Exit Function
    
    If CheckInputValied = False Then Exit Function
    
    If mEditType = g_Ed_�������� Or mEditType = g_Ed_�������� Or chkCancel.Value = 1 And chkCancel.Visible Then
        If ExecuteBalaceCancel = False Then Exit Function
        mintSucces = mintSucces + 1
        mbln�ѱ��� = False
        SaveBalanceData = True: Exit Function
    End If
    
    If CheckChargeAudit(mPatiInfor.lng����ID, True, mobjBalanceCon.strTime) = False Then Exit Function
    
    Screen.MousePointer = 99
    
    If mPatiInfor.lng����ID = 0 Or Trim(txtPatient.Text) = "" Then
        MsgBox "δ���뱾�ν��ʵĲ���,��������ʲ���", vbInformation + vbOKOnly, gstrSysName
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Exit Function
    End If
    
    blnHaveFee = False
    If Not mrsFeeList Is Nothing Then
       If mrsFeeList.State = 1 Then
            If mrsFeeList.RecordCount <> 0 Then blnHaveFee = True
       End If
    End If
    
    If blnHaveFee = False And mEditType <> g_Ed_���½��� Then
        MsgBox "���˲�������Ҫ���ʵķ���,������ˢ��δ����á�", vbInformation + vbOKOnly, gstrSysName
        If cmdRefresh.Enabled And cmdRefresh.Visible Then cmdRefresh.SetFocus
        Exit Function
    End If
    
    Set objCard = IDKindPaymentsType.GetCurCard
    If Not objCard Is Nothing Then
        If objCard.�������� = 1 And mty_ModulePara.byt�ɿ�������� = 1 Then
            If (Val(lblʣ���Ը�.Caption) <> 0 And Val(txtBalance(4).Text) = 0) Then
                If Not (InStr(IDKind�Ҳ�.GetCurCard.����, "Ԥ��") > 0 And Val(lblʣ���Ը�.Caption) - Val(txtBalance(5).Text) = 0) Then
                    MsgBox "δ����ɿ������ʹ��" & objCard.���㷽ʽ & "���н��㣡", vbInformation + vbOKOnly, gstrSysName
                    If txtBalance(4).Enabled And txtBalance(4).Visible Then txtBalance(4).SetFocus
                    Exit Function
                End If
            End If
        End If
        If (objCard.���ѿ� Or objCard.�Ƿ�����ʻ�) And (mEditType = g_Ed_������� Or mEditType = g_Ed_סԺ���� Or mEditType = g_Ed_���½���) Then
            If mPatiInfor.bln�˿��־ And objCard.�Ƿ�ת�ʼ����� = False Then
                With vsBlance
                    blnFind = False
                    For i = 1 To .Rows - 1
                        If Not objCard.���ѿ� And objCard.�ӿ���� <> 0 Then '���ѿ�,�Ѿ����,�����ٴ���
                            If .TextMatrix(i, .ColIndex("���㷽ʽ")) = objCard.���㷽ʽ Then blnFind = True
                        End If
                    Next
                    If blnFind Then
                        Screen.MousePointer = 0
                        MsgBox objCard.���㷽ʽ & " �Ѿ�֧����,��������" & objCard.���㷽ʽ & "����֧��", vbOKOnly + vbDefaultButton1, gstrSysName
                        Exit Function
                    End If
                End With
                Call LoadIntendBalance(Val(txtBalance(4).Text), objCard)
                Call LoadCurOwnerPayInfor
                SaveBalanceData = True
                Exit Function
            End If
        End If
    End If
     
    If mBalanceInfor.blnԤ����֤ = False And Val(txtBalance(Idx_��Ԥ��).Text) <> 0 Then
        If DepositMonyVerfy(True) = False Then Exit Function
       ' MsgBox "Ԥ�����δ��֤,����������ɿ���!", vbInformation + vbOKOnly, gstrSysName
'        blnCancel = False
'        Call txtBalance_Validate(Idx_��Ԥ��, blnCancel)
'        If blnCancel Then Exit Function
'
'        If txtBalance(Idx_�ɿ�).Visible And txtBalance(Idx_�ɿ�).Enabled Then txtBalance(Idx_�ɿ�).SetFocus
'        Exit Function
    End If
    
    If mBalanceInfor.blnԤ��ˢ�� = False And Val(txtBalance(Idx_��Ԥ��).Text) <> 0 And (mEditType = g_Ed_������� Or mblnCurMzBalanceNo) Then
        If Not zlDatabase.PatiIdentify(Me, glngSys, mPatiInfor.lng����ID, Val(txtBalance(Idx_��Ԥ��).Text), _
                , , , IIf(-1 * gdblԤ��������鿨 >= Val(txtBalance(Idx_��Ԥ��).Text), False, True), , , (gdblԤ��������鿨 <> 0), (gdblԤ��������鿨 = 2)) Then
            Exit Function
        Else
            mBalanceInfor.blnԤ��ˢ�� = True
        End If
    End If
    
    '���ж��Ƿ����Ԥ����ˢ���ģ����ȴ���Ԥ����
    If Not CheckDepositValied(blnԤ��) Then
        If txtBalance(Idx_��Ԥ��).Enabled And txtBalance(Idx_��Ԥ��).Visible Then txtBalance(Idx_��Ԥ��).SetFocus
        zlControl.TxtSelAll txtBalance(Idx_��Ԥ��):
        Exit Function
    End If
    
    Call LedVoiceSpeak(False)
    
    If Not blnԤ�� Then
        If CheckCurBalanceIsValied(tyBrushCard, , objSetFocus) = False Then
            If Not objSetFocus Is Nothing Then
                If objSetFocus.Enabled And objSetFocus.Visible Then objSetFocus.SetFocus
                If UCase(TypeName(objSetFocus)) = UCase("txtEdit") Then
                    zlControl.TxtSelAll objSetFocus
                End If
            End If
            Exit Function
        End If
        If CheckDepositFactValied = False Then Exit Function
    End If
    
    If mblnNotify = False Then
        If MsgBox("��ȷ��Ҫ�Ըò��˽��н�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Function
        End If
        mblnPrintInvoice = False
        If Not mobjBalanceCon.blnCurBalanceOwnerFee Then   '���Էѷ���ʱ,Ҫ��ӡ��Ʊ
            If Not (mYBInFor.intInsure <> 0 And MCPAR.ҽ���ӿڴ�ӡƱ��) Then
                '���ղ��˸���ʹ�����������ȷ����
                Select Case mobjFactProperty.��ӡ��ʽ
                Case 0  '����ӡ
                Case 1
                    mblnPrintInvoice = True '�Զ���ӡ
                Case 2  '��ʾ��ӡ
                    If MsgBox("�Ƿ��ӡƱ��?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then mblnPrintInvoice = True
                End Select
            End If
        Else
            If Not mty_ModulePara.blnNotPrintInvioce Then
                Select Case mobjFactProperty.��ӡ��ʽ
                Case 0  '����ӡ
                Case 1
                    mblnPrintInvoice = True '�Զ���ӡ
                Case 2  '��ʾ��ӡ
                    If MsgBox("�Ƿ��ӡƱ��?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then mblnPrintInvoice = True
                End Select
            End If
        End If
        mblnNotify = True
    End If
    
    LockedScreen True
    blnSaved = SaveBalaceCharge(blnԤ��, tyBrushCard, blnChargeEnd, objSetFocus)
    LockedScreen False
    
    mbln�ѱ��� = False
    
    If blnChargeEnd Then
        If mEditType = g_Ed_���½��� Then Unload Me: Exit Function
    End If
    
    '0-����ǰ;1-ҽ����������;2-�ѱ����˽��ʵ�;
    If mBalanceInfor.blnSaveBill Then
       Call SetOperationCtrl(2)
    Else
       Call SetOperationCtrl(0)
    End If
    
    If blnChargeEnd Then
        mblnNotify = False
        Call AddNoToCombox  '���ص��ݺ�
        Call SetOperationCtrl(0)
        mintSucces = mintSucces + 1
        If mintPreEditType <> -1 Then mEditType = mintPreEditType
        mlngPatientID = 0
        mBalanceInfor.blnSaveBill = False
        staThis.Panels(3).Text = "�ϴν���:" & Format(mBalanceInfor.dbl��ǰ����, "0.00")
        
        bln�������۲��� = Val(NVL(mrsInfo!��������)) = 1
        lng����ID = Val(NVL(mrsInfo!����ID)): str���� = NVL(mrsInfo!����)
        
        If mbln�������� Or mobjBalanceCon.blnCurBalanceOwnerFee Then
            '�����������Ϣ
            If mobjBalanceCon.blnCurBalanceOwnerFee Then
                lblPrevious.Visible = True
                If (12 - Len(Format(mBalanceInfor.dbl��ǰ����, "0.00"))) / 2 < 0 Then
                    lblPrevious.Caption = "�ϴ��Էѽ���" & vbCrLf & Format(mBalanceInfor.dbl��ǰ����, "0.00")
                Else
                    lblPrevious.Caption = "�ϴ��Էѽ���" & vbCrLf & String((12 - Len(Format(mBalanceInfor.dbl��ǰ����, "0.00"))) / 2, " ") & Format(mBalanceInfor.dbl��ǰ����, "0.00")
                End If
                lblCurPaymentInfor.Visible = False
            End If
           If ShowBalance(True, strErrMsg, blnNotClearPati) = False Then
                cmdOK.Enabled = False
                MsgBox "�ڵ�ǰ������,���˲�������Ҫ���ʵķ��ã�", vbInformation, gstrSysName
                If cmdRefresh.Visible And cmdRefresh.Enabled Then cmdRefresh.SetFocus
                Call SetBatchControl(False)
                SaveBalanceData = blnSaved
                Exit Function
           End If
           Call Load�����Ϣ(Val(NVL(mrsInfo!����ID)), IIf(mEditType = g_Ed_������� Or mblnCurMzBalanceNo, 1, 2))
           Call ReInitPatiInvoice
        Else
            '���˺�:27503
            If mty_ModulePara.bln���ʺ�����Ϣ Then
                Set mrsInfo = New ADODB.Recordset
                If txtInvoice.Tag <> "" And txtInvoice.Text <> txtInvoice.Tag Then txtInvoice.Text = txtInvoice.Tag '��Ҫ��Ҫ������Ϣ,��ȷ������Ҫ�����̶�
                Dim strTemp As String
                strTemp = txtInvoice.Text
                Call ReInitPatiInvoice
                txtInvoice.Text = strTemp   '��Ҫ�ǲ�Ҫ����ϴεķ�Ʊ,�µķ�Ʊ����.tag��,�ڸı䲡��ʱ,ֱ�Ӵ�����ط���ȡ
                mblnNotClearBill = True
                Call SetBatchControl(False)
            Else
                Call LoadBalanceBill
                Call ReInitPatiInvoice(Not mobjFactProperty.����ʹ�����)
            End If
        End If
        
        '139063���������۲����������δ���סԺ���ã��������ʾ
        If bln�������۲��� And mEditType = g_Ed_������� Then
            dbl������� = GetRemainderMoney(lng����ID, 2)
            If dbl������� > 0 Then
                MsgBox "ע�⣺" & vbCrLf & _
                       "    ��ǰ���ˡ�" & str���� & "��������δ�����סԺ���ã�ע�������н��ˣ�", vbInformation, gstrSysName
            End If
        End If
        
        staThis.Panels(2) = "������ϣ��������������˱�ʶ��"
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    ElseIf blnSaved Then
        If blnԤ�� And txtBalance(Idx_��Ԥ��).Enabled And txtBalance(Idx_��Ԥ��).Visible Then
            txtBalance(Idx_��Ԥ��).SetFocus
            zlControl.TxtSelAll txtBalance(Idx_��Ԥ��)
        ElseIf txtBalance(Idx_�ɿ�).Enabled And txtBalance(Idx_�ɿ�).Visible Then
            txtBalance(Idx_�ɿ�).SetFocus
            zlControl.TxtSelAll txtBalance(Idx_�ɿ�)
        End If
    ElseIf Not objSetFocus Is Nothing Then
        If objSetFocus.Enabled And objSetFocus.Visible Then objSetFocus.SetFocus
        zlControl.TxtSelAll objSetFocus
    End If
    
    SaveBalanceData = blnSaved
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    LockedScreen False
    '0-����ǰ;1-ҽ����������;2-�ѱ����˽��ʵ�;
    If mBalanceInfor.blnSaveBill Then
       Call SetOperationCtrl(2)
    Else
       Call SetOperationCtrl(0)
    End If
End Function

Private Sub SetBatchControl(ByVal blnState As Boolean)
    mblnBatchState = Not blnState
    cmdOK.Enabled = blnState
    cmdCancel.Enabled = blnState
    cmdRefresh.Enabled = blnState
    cmdNext.Enabled = blnState
    '���Ժ����������ʹ�ã�����������
'    IDKindPaymentsType.Enabled = blnState
'    txtBalance(Idx_�ɿ�).Enabled = blnState
    txtBalance(Idx_�Ҳ�).Enabled = blnState
    txtBalance(Idx_ժҪ).Enabled = blnState
    txtBalance(Idx_�������).Enabled = blnState
    txtBalance(Idx_����˵��).Enabled = blnState
    txtBalance(Idx_��Ԥ��).Enabled = blnState
    txtBalance(Idx_���ν���).Enabled = blnState And InStr(mstrPrivs, ";��������;") > 0
    txtBalance(Idx_���ν���).Locked = InStr(mstrPrivs, ";��������;") = 0
 
    txtBegin.Enabled = blnState
    txtEnd.Enabled = blnState
    txtPatiBegin.Enabled = blnState
    txtPatiEnd.Enabled = blnState
    cboӤ��.Enabled = blnState And InStr(mstrPrivs, ";��������;") > 0
    
    cboPatiNums.Enabled = blnState And InStr(mstrPrivs, ";��������;") > 0
    cboFeeType.Enabled = blnState And InStr(mstrPrivs, ";��������;") > 0
    cboFeeItem.Enabled = blnState And InStr(mstrPrivs, ";��������;") > 0
    cboDept.Enabled = blnState And InStr(mstrPrivs, ";��������;") > 0
    cboChargeType.Enabled = blnState And InStr(mstrPrivs, ";��������;") > 0
    cboDiag.Enabled = blnState And InStr(mstrPrivs, ";��������;") > 0
    
    
    cboӤ��.BackColor = IIf(cboӤ��.Enabled, &H80000005, &H8000000F)
    cboPatiNums.BackColor = IIf(cboPatiNums.Enabled, &H80000005, &H8000000F)
    cboFeeType.BackColor = IIf(cboFeeType.Enabled, &H80000005, &H8000000F)
    cboFeeItem.BackColor = IIf(cboFeeItem.Enabled, &H80000005, &H8000000F)
    cboDept.BackColor = IIf(cboDept.Enabled, &H80000005, &H8000000F)
    cboChargeType.BackColor = IIf(cboChargeType.Enabled, &H80000005, &H8000000F)
    cboDiag.BackColor = IIf(cboDiag.Enabled, &H80000005, &H8000000F)
    
    txtBalance(Idx_���ν���).BackColor = IIf(txtBalance(Idx_���ν���).Enabled, &H80000005, &H8000000F)
      
    
    opt��;.Enabled = blnState
    opt��Ժ.Enabled = blnState
    If blnState Then
        txtBalance(Idx_�Ҳ�).BackColor = &H80000005
        txtBalance(Idx_ժҪ).BackColor = &H80000005
        txtBalance(Idx_�������).BackColor = &H80000005
        txtBalance(Idx_����˵��).BackColor = &H80000005
        txtBalance(Idx_��Ԥ��).BackColor = &H80000005
    Else
        txtBalance(Idx_�Ҳ�).BackColor = &H8000000F
        txtBalance(Idx_ժҪ).BackColor = &H8000000F
        txtBalance(Idx_�������).BackColor = &H8000000F
        txtBalance(Idx_����˵��).BackColor = &H8000000F
        txtBalance(Idx_��Ԥ��).BackColor = &H8000000F
    End If
End Sub

Private Sub AddNoToCombox()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ص������Combox�ؼ���
    '����:���˺�
    '����:2015-02-11 17:30:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String, i As Long
    
    '���뵥����ʷ��¼(�������͵���)
    On Error GoTo errHandle
    strTmp = mBalanceInfor.strNO
    For i = 0 To cboNO.ListCount - 1
        strTmp = strTmp & "," & cboNO.List(i)
    Next
    cboNO.Clear
    For i = 0 To UBound(Split(strTmp, ","))
        cboNO.AddItem Split(strTmp, ",")(i)
        If i = 9 Then Exit For 'ֻ��ʾ10��
    Next

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Sub cmdDepositUp_Click()
    If mEditType <> g_Ed_������� And mEditType <> g_Ed_סԺ���� And mEditType <> g_Ed_���½��� Then Exit Sub
    With vsDeposit
        If .Row <= 1 Then Exit Sub
        .RowPosition(.Row) = .Row - 1
        .Select .Row - 1, 1
'        Call RecalcDepositMoney(2, Val(mBalanceInfor.dbl��Ԥ���ϼ�))
        Call LoadIntendBalance
        Call LoadCurOwnerPayInfor
    End With
End Sub

Private Sub cmdDepositDown_Click()
    If mEditType <> g_Ed_������� And mEditType <> g_Ed_סԺ���� And mEditType <> g_Ed_���½��� Then Exit Sub
    With vsDeposit
        If .Row >= .Rows - 1 Then Exit Sub
        .RowPosition(.Row) = .Row + 1
        .Select .Row + 1, 1
'        Call RecalcDepositMoney(2, Val(mBalanceInfor.dbl��Ԥ���ϼ�))
        Call LoadIntendBalance
        Call LoadCurOwnerPayInfor
    End With
End Sub

Private Sub cmdNext_Click()
   If chkCancel.Value = 1 Then Exit Sub
   If Not (mEditType = g_Ed_������� Or mEditType = g_Ed_סԺ����) Then Exit Sub
   mbln�������� = True
   If SaveBalanceData = False Then Exit Sub
   mbln�������� = False
End Sub

Private Sub cmdOK_Click()
    mbln�������� = False
    If mEditType = g_Ed_���ݲ鿴 Then Unload Me: Exit Sub
    If mEditType = g_Ed_ȡ������ Then
        If DeleteBalance = False Then Exit Sub
        mintSucces = mintSucces + 1
        Unload Me: Exit Sub
    End If
    If SaveBalanceData = False Then Exit Sub
End Sub
 
Private Sub cmdDelBalance_Click()
    '��������
    If MsgBox("�����Ҫ���ϵ�ǰ�Ľ�����Ϣ��?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    '������ȡ����
    If zlGetFromIDToBalanceData(mBalanceInfor.lng����ID, False, mrsBalance) = False Then Exit Sub
    
    If DeleteBalance(True) = False Then Exit Sub
    mintSucces = mintSucces + 1
    Call NewBill
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
End Sub

Private Sub cmdRefresh_Click()
    Dim blnAll As Boolean, i As Long
    Dim objCard As Card, arrTime  As Variant
    Dim blnNotPati As Boolean, intMaxTime As Integer
    Dim dblδ���ۼ� As Double
    Dim strTime As String, strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim intInsure As Integer, strInsureName As String
    
    
    blnNotPati = False
    If mrsInfo Is Nothing Then blnNotPati = True
    If blnNotPati = False Then
        If mrsInfo.State = 0 Then blnNotPati = True
    End If
    
    If blnNotPati Then
        MsgBox "û��ȷ�����ʲ���,����ˢ��δ�������Ϣ��", vbExclamation, gstrSysName
        If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
        Exit Sub
    End If
    
    If mEditType = g_Ed_������� Then
        If chkPatiType(CK_Idx_��ͨ).Value = 0 And chkPatiType(CK_Idx_���).Value = 0 Then
            MsgBox "��ͨ���ú���������Ҫ����ѡ��һ��,����ˢ��δ�������Ϣ��", vbExclamation, gstrSysName
            Exit Sub
        End If
    End If
    
    Call SetBalanceConInfor
    
    If cboPatiNums.Text = "" Or cboDept.Text = "" Or cboFeeType.Text = "" Or cboFeeItem.Text = "" Or cboChargeType.Text = "" Then
        MsgBox "����δ���õĽ����������޷���ȡ������Ϣ��", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    If mPatiInfor.bln�������� Then
        mbln�������� = True
        dblδ���ۼ� = mPatiInfor.dblδ���ۼ�
    End If
    
    If mEditType <> g_Ed_������� Then
        With mobjBalanceCon
            intMaxTime = 0
            '��������ж��סԺ����δ�ᣬ��ֻѡ���ĳ��סԺ���ã�����ݸô�סԺ��Ϣ�����������Ƿ���ҽ������
            Set objCard = IDKIND.GetIDKindCard("����", CardTypeName)
            txtPatient.Text = "-" & mrsInfo!����ID
            mblnSecondLoadPati = True
            If .strTime = "" Then
                arrTime = Split(mobjBalanceAll.strAllTime, ",")
                For i = 0 To UBound(arrTime)
                    If Val(arrTime(i)) > intMaxTime Then intMaxTime = Val(arrTime(i))
                Next i
            Else
                If InStr(1, .strTime, ",") > 0 Then
                    arrTime = Split(.strTime, ",")
                    For i = 0 To UBound(arrTime)
                        If Val(arrTime(i)) > intMaxTime Then intMaxTime = Val(arrTime(i))
                    Next i
                Else
                    intMaxTime = Val(.strTime)
                End If
            End If
            Call LoadPatientInfo(IDKIND.GetCurCard, False, 0, intMaxTime)
            mblnSecondLoadPati = False
        End With
        
        Call CheckPatiFromZyNumIsYB(mrsInfo!����ID, intMaxTime, intInsure, strInsureName)
        
        mYBInFor.intInsure = intInsure
        If mYBInFor.intInsure = 0 Then
            txtPatient.ForeColor = Me.ForeColor
            staThis.Panels(4).Text = ""
            staThis.Panels(4).Visible = False
            cmdOK.Enabled = True
        End If
        
    End If
    
    If mbln�������� Then
        mPatiInfor.bln�������� = mbln��������
        mPatiInfor.dblδ���ۼ� = dblδ���ۼ�
    End If
 
    
    If Not ShowBalance() Then
        cmdOK.Enabled = False
        mbln�������� = False
        If mrsInfo.State <> 1 Then
            txtPatient.Locked = False: txtPatient.Text = ""
           Call NewBill
           If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
            Screen.MousePointer = 0
           Exit Sub
        End If
        Screen.MousePointer = 0
        MsgBox "�ڵ�ǰ������,���˲�����Ҫ���ʵķ��ã�", vbInformation, gstrSysName
        If cmdRefresh.Visible And cmdRefresh.Enabled Then cmdRefresh.SetFocus
        Exit Sub
    End If
    
    mbln�������� = False
    'ȷ������˳��
    If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus

    If Val(txtBalance(Idx_��Ԥ��).Text) <> 0 And txtBalance(Idx_��Ԥ��).Enabled And txtBalance(Idx_��Ԥ��).Visible Then
        txtBalance(Idx_��Ԥ��).SetFocus
        zlControl.TxtSelAll txtBalance(Idx_��Ԥ��)
    Else
        If txtBalance(Idx_�ɿ�).Enabled And txtBalance(Idx_�ɿ�).Visible Then
            txtBalance(Idx_�ɿ�).SetFocus
            zlControl.TxtSelAll txtBalance(Idx_�ɿ�)
        End If
    End If
    
    If cmdYBBalance.Visible And cmdYBBalance.Enabled Then cmdYBBalance.SetFocus
    mblnConsChange = False
    mbln�ѱ��� = False
    Screen.MousePointer = 0
End Sub

Private Sub cmdTools_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    Call AddPopu
End Sub

Private Sub cmdYB_Click()
    '���ﲡ�˽���ǰ�������֤(�ɶ�ҽ����֧��סԺ����ҽ�������֤)
    Dim lng����ID As Long, bytMode As Byte
    Dim strMessage As String, intInsure As Integer
    Dim strPatiName As String
    lng����ID = 0
    If Not mrsInfo Is Nothing Then
        If mrsInfo.State = 1 Then lng����ID = Val(NVL(mrsInfo!����ID))
    End If
    strPatiName = Trim(txtPatient.Text)
    Call NewBill
    txtPatient.Text = strPatiName
    
    bytMode = 0
    If mblnMC_TwoMode Then
        If InStr(mstrPrivs, ";������ý���;") = 0 Then
            bytMode = 4
        Else
            If zlCommFun.ShowMsgbox("ҽ����֤��֤", "��ѡ���������֤ģʽ��", "!סԺҽ��(&Z),����ҽ��(&M)", Me, vbInformation) = "סԺҽ��" Then
                bytMode = 4
            End If
        End If
    End If
        
    '���˺�:����תסԺ����ʱ����
    mYBInFor.strYBPati = gclsInsure.Identify(bytMode, lng����ID, intInsure)
    mYBInFor.intInsure = intInsure
    
    If mYBInFor.strYBPati = "" Then GoTo ExceptionHand
    cmdOK.Enabled = False   '����:43776
    
    mYBInFor.bytMCMode = IIf(bytMode = 0, 1, 2) '������LoadPatientInfo֮ǰ
    
    If mYBInFor.bytMCMode = 1 Then
        'lng����ID:49084
        If Not gclsInsure.GetCapability(support�������, lng����ID, intInsure) Then
            strMessage = "���˵�ǰ���಻֧������ҽ�����ʡ�": GoTo ExceptionHand
        End If
    End If
    
    'New:�ջ�0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);8����ID
    If UBound(Split(mYBInFor.strYBPati, ";")) >= 8 Then lng����ID = Val(Split(mYBInFor.strYBPati, ";")(8))
    If lng����ID <> 0 Then
        txtPatient.Text = "-" & lng����ID
        Call LoadPatientInfo(IDKIND.GetCurCard, False, intInsure)
        If mrsInfo.State = 0 Then GoTo ExceptionHand
    Else
        strMessage = "���������֤�ɹ�,��δ���ֲ��˵��ʻ���Ϣ!" & vbCrLf & "�����ǲ�����Ժʱû�н�����֤,���ܽ��б��ս��㣡"
        GoTo ExceptionHand
    End If
    Exit Sub
ExceptionHand:
    If strMessage <> "" Then Call MsgBox(strMessage, vbInformation, gstrSysName)
    Set mrsInfo = New ADODB.Recordset
    mYBInFor.strYBPati = "": mYBInFor.bytMCMode = 0
    txtPatient.Text = "": txtPatient.SetFocus
    cmdOK.Enabled = True
End Sub
Private Function ExcuteInsureSwapInteface(ByVal lng����ID As Long, ByVal cllSaveBill As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ��ҽ�����׽ӿ�
    '���:cllSaveBill-���浥�ݵ�sql
    '����:ִ�гɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-13 15:11:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPro As Collection, i As Long, str���㷽ʽ As String
    Dim blnTrans As Boolean, intInsure As Integer, strAdvance As String
    Dim blnTransMC As Boolean, blnMark As Boolean
    Dim cur�����ʻ� As Currency, curҽ������ As Currency
    Dim blnInsureCheck As Boolean
    On Error GoTo errHandle
    
    intInsure = mYBInFor.intInsure
    '��ҽ������Էѷ���,����ִ��
    If intInsure = 0 Or mobjBalanceCon.blnCurBalanceOwnerFee Then Exit Function
    
    Set cllPro = New Collection
    For i = 1 To cllSaveBill.Count
        zlAddArray cllPro, cllSaveBill(i)
    Next
    
    str���㷽ʽ = GetMedicareStr(cur�����ʻ�, curҽ������)
    If ҽ�����ݸ���(Val(NVL(mrsInfo!����ID)), lng����ID, str���㷽ʽ, False, cllPro) = False Then Exit Function
    
    
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    '����ҽ���ӿ�
    blnTransMC = False
    If mYBInFor.bytMCMode = 1 Then
        '����ҽ������
        strAdvance = ""
        If cur�����ʻ� <> 0 Or curҽ������ <> 0 Or MCPAR.������봫����ϸ Then
            Call SetCmdStatus(False)
            If Not gclsInsure.ClinicSwap(lng����ID, cur�����ʻ�, curҽ������, 0, 0, intInsure, strAdvance) Then
                Call SetCmdStatus(True)
                gcnOracle.RollbackTrans: Screen.MousePointer = 0: Exit Function
            End If
            Call SetCmdStatus(True)
            blnTransMC = True
        End If
        GoTo SaveEnd:
    End If
    'סԺҽ������
    Call SetCmdStatus(False)
    If Not gclsInsure.SettleSwap(lng����ID, intInsure, strAdvance) Then
        Call SetCmdStatus(True)
        gcnOracle.RollbackTrans: Screen.MousePointer = 0: Exit Function
    Else
        Call SetCmdStatus(True)
        blnTransMC = True
    End If

SaveEnd:
    If strAdvance <> "" Then
        If zlInsure_Check(str���㷽ʽ, strAdvance) Then
            blnInsureCheck = True
            Call ҽ�����ݸ���(Val(NVL(mrsInfo!����ID)), lng����ID, strAdvance, False, Nothing)
CheckAgain:
            blnMark = False
            For i = 1 To vsBlance.Rows - 1
                If Val(vsBlance.TextMatrix(i, vsBlance.ColIndex("����"))) = 2 Then
                    Call DeletePayInfor(i, True)
                    blnMark = True
                    Exit For
                End If
            Next i
            If blnMark = True Then GoTo CheckAgain
        End If
    End If
    mBalanceInfor.blnSaveBill = True
    gcnOracle.CommitTrans: blnTrans = False
    If blnTransMC Then
        Call gclsInsure.BusinessAffirm(IIf(mYBInFor.bytMCMode = 1, ����Enum.Busi_ClinicSwap, ����Enum.Busi_SettleSwap), True, intInsure)
    End If
    Set cllSaveBill = New Collection
    Screen.MousePointer = 0
    ExcuteInsureSwapInteface = True
    Exit Function
errHandle:
    Screen.MousePointer = 0
    If blnTrans Then gcnOracle.RollbackTrans
    Call SetCmdStatus(True)
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
    Call SaveErrLog
    If blnTrans Then
        'ҽ����HIS����ͬһ������,HIS����ʧ��,��ҽ���������ϴ�,������Ҫ��"ȡ������"�ӿ�
        If blnTransMC Then Call gclsInsure.BusinessAffirm(IIf(mYBInFor.bytMCMode = 1, ����Enum.Busi_ClinicSwap, ����Enum.Busi_SettleSwap), False, intInsure)
    End If
End Function

Private Function ҽ�����ݸ���(ByVal lng����ID As Long, ByVal lng����ID As Long, _
    ByVal strҽ������ As String, ByVal bln���� As Boolean, _
    ByRef cllPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ҽ������У�Ը���
    '����:У�Գɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-06-12 17:45:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    If bln���� Then
        'Zl_�����˷ѽ���_Modify
        strSQL = "Zl_�����˷ѽ���_Modify("
        '  ��������_In   Number,
        strSQL = strSQL & "" & 3 & ","
        '  ����id_In     ������ü�¼.����id%Type,
        strSQL = strSQL & "" & lng����ID & ","
        '  ����id_In     ����Ԥ����¼.����id%Type,
        strSQL = strSQL & "" & lng����ID & ","
        '  ���㷽ʽ_In   Varchar2,
        strSQL = strSQL & "'" & strҽ������ & "')"
        '  ��Ԥ��_In     ����Ԥ����¼.��Ԥ��%Type := Null,
        '  �����id_In   ����Ԥ����¼.�����id%Type := Null,
        '  ����_In       ����Ԥ����¼.����%Type := Null,
        '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
        '  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
        '  �ɿ�_In       ����Ԥ����¼.�ɿ�%Type := Null,
        '  �Ҳ�_In       ����Ԥ����¼.�Ҳ�%Type := Null,
        '  �����_In   ������ü�¼.ʵ�ս��%Type := Null,
        '  ����˷�_In Number:=0
        ') As
        '  ------------------------------------------------------------------------------------------------------------------------------
        '  --����:�շѽ���ʱ,�޸Ľ���������Ϣ
        '  --��������_In:
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
        '  -- �����_In:��������ʱ,����
        '  -- ����˷�_In:0-δ����˷�;1-�쳣����˷�;2-����˷�
        '  ------------------------------------------------------------------------------------------------------------------------------
     Else
  
        'Zl_���˽��ʽ���_Modify
        strSQL = "Zl_���˽��ʽ���_Modify("
        '  ��������_In     Number,
        '  --   2-ҽ������(�������ҽ���Ľ���,��Ҫ��ɾ��ԭҽ������,���´���ĸ���)
        '  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:���㷽ʽ|������||.."
        strSQL = strSQL & "" & 2 & ","
        '  ����id_In       ���˽��ʼ�¼.����id%Type,
        strSQL = strSQL & "" & lng����ID & ","
        '  ����id_In       ����Ԥ����¼.����id%Type,
        strSQL = strSQL & "" & lng����ID & ","
        '  ���㷽ʽ_In     Varchar2,
        strSQL = strSQL & "" & IIf(strҽ������ = "", "NULL", "'" & strҽ������ & "'") & ","
        '  ��Ԥ��_In       ����Ԥ����¼.��Ԥ��%Type := Null,
        strSQL = strSQL & "NULL,"
        '  ��֧Ʊ��_In     ����Ԥ����¼.��Ԥ��%Type := Null,
        strSQL = strSQL & "NULL,"
        '  �����id_In     ����Ԥ����¼.�����id%Type := Null,
        strSQL = strSQL & "NULL,"
        '  ����_In         ����Ԥ����¼.����%Type := Null,
        strSQL = strSQL & "NULL,"
        '  ������ˮ��_In   ����Ԥ����¼.������ˮ��%Type := Null,
        strSQL = strSQL & "NULL,"
        '  ����˵��_In     ����Ԥ����¼.����˵��%Type := Null,
        strSQL = strSQL & "NULL,"
        '  �ɿ�_In         ����Ԥ����¼.�ɿ�%Type := Null,
        strSQL = strSQL & "NULL,"
        '  �Ҳ�_In         ����Ԥ����¼.�Ҳ�%Type := Null,
        strSQL = strSQL & "NULL,"
        '  �����_In     ������ü�¼.ʵ�ս��%Type := Null,
        strSQL = strSQL & "NULL,"
        '  ��������_In     Number := 2,(1-�������;2-סԺ����)
        strSQL = strSQL & "" & IIf(mEditType = g_Ed_������� Or mblnCurMzBalanceNo, 1, 2) & ","
        '  ȱʡ���㷽ʽ_In ���㷽ʽ.����%Type := Null,
        strSQL = strSQL & "NULL,"
        '    ����Ա���_In    ����Ԥ����¼.����Ա���%Type := Null,
        strSQL = strSQL & "'" & UserInfo.��� & "',"
        '    ����Ա����_In    ����Ԥ����¼.����Ա����%Type := Null,
        strSQL = strSQL & "'" & UserInfo.���� & "',"
        '    �տ�ʱ��_In      ����Ԥ����¼.����Ա����%Type := Null,
        strSQL = strSQL & "to_date('" & Format(mBalanceInfor.dtBalanceDate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
        '    ��Ԥ������ids_In Varchar2 := Null,
        strSQL = strSQL & "NULL,"
        '  ��ɽ���_In Number:=0: (1-����շ�;0-δ����շ�)
        strSQL = strSQL & "0)"
     End If
     
    If cllPro Is Nothing Then
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Else
        zlAddArray cllPro, strSQL
    End If
    ҽ�����ݸ��� = True
End Function
Public Function zlInsure_Check(ByVal str���ս��� As String, ByVal strAdvance As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鵱ǰ��ҽ���Ƿ���Ҫ�϶�
    '���:str���ս���-���ս���
    '       strAdvance-ҽ�����صĽ���
    '����:
    '����:��Ҫ�϶�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-20 18:03:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnMedicareCheck As Boolean, strTmp As String, i As Long, j As Long
    Dim varData As Variant, varData1 As Variant
    Dim varTemp As Variant, varTemp1 As Variant
    
    On Error GoTo errHandle
    If Not (strAdvance <> "" And str���ս��� <> strAdvance) Then Exit Function
    '��ʽ����ǰ��,���㷽ʽ�ͽ�����δ�����仯ʱ��У��
    blnMedicareCheck = True
    varData = Split(str���ս���, "||"): varData1 = Split(strAdvance, "||")
    
    If UBound(varData) = UBound(varData1) Then
        For i = 0 To UBound(varData)
            blnMedicareCheck = True
            strTmp = varData(i)
            varTemp = Split(strTmp, "|")
            For j = 0 To UBound(varData1)
                varTemp1 = Split(varData1(j), "|")
                If varTemp(0) = varTemp1(0) Then
                    If Val(varTemp(1)) = Val(varTemp1(1)) Then
                        blnMedicareCheck = False
                    End If
                End If
            Next
            If blnMedicareCheck Then Exit For
        Next
    End If
    zlInsure_Check = blnMedicareCheck
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetMedicareStr(ByRef cur�����ʻ� As Currency, curҽ������ As Currency) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ر��ս��㷽ʽ��,"���㷽ʽ|���||...."
    '����:cur�����ʻ�-�����ʻ�
    '     curҽ������-ҽ������
    '����:���ر��ս��㷽ʽ��,"���㷽ʽ|���||...."
    '����:���˺�
    '����:2015-01-13 15:16:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strTemp As String
    Dim curMoney As Currency, int���� As Integer
    strTemp = ""
    cur�����ʻ� = 0: curҽ������ = 0
    With vsBlance
        For i = 1 To .Rows - 1
            '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
            int���� = Val(.TextMatrix(i, .ColIndex("����")))
            curMoney = Val(.TextMatrix(i, .ColIndex("������")))
            
            If int���� = 2 And .TextMatrix(i, .ColIndex("���㷽ʽ")) <> "" Then
                strTemp = strTemp & "||" & .TextMatrix(i, .ColIndex("���㷽ʽ")) & "|" & Format(curMoney, gstrDec)
                If Val(.TextMatrix(i, .ColIndex("��������"))) = 3 Then cur�����ʻ� = cur�����ʻ� + curMoney
                If Val(.TextMatrix(i, .ColIndex("��������"))) = 4 Then curҽ������ = curҽ������ + curMoney
            End If
        Next
    End With
    If strTemp <> "" Then strTemp = Mid(strTemp, 3)
    GetMedicareStr = strTemp
End Function


Private Sub cmdYBBalance_Click()
    Dim cllPro As Collection
    Dim objFocus As Object
    Dim lng����ID As Long
    
    '������Ч�Լ��
    If CheckInputConsValied(objFocus) = False Then
        If objFocus Is Nothing Then Exit Sub
        If objFocus.Enabled And objFocus.Visible Then objFocus.SetFocus
        If UCase(TypeName(objFocus)) = UCase("txtEdit") Then
            zlControl.TxtSelAll objFocus
        End If
        Exit Sub
    End If
    
    If mblnNotify = False Then
        If MsgBox("��ȷ��Ҫ�Ըò��˽��н�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
        mblnPrintInvoice = False
        If Not mobjBalanceCon.blnCurBalanceOwnerFee Then   '���Էѷ���ʱ,Ҫ��ӡ��Ʊ
            If Not (mYBInFor.intInsure <> 0 And MCPAR.ҽ���ӿڴ�ӡƱ��) Then
                '���ղ��˸���ʹ�����������ȷ����
                Select Case mobjFactProperty.��ӡ��ʽ
                Case 0  '����ӡ
                Case 1
                    mblnPrintInvoice = True '�Զ���ӡ
                Case 2  '��ʾ��ӡ
                    If MsgBox("�Ƿ��ӡƱ��?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then mblnPrintInvoice = True
                End Select
            End If
        Else
            If Not mty_ModulePara.blnNotPrintInvioce Then
                Select Case mobjFactProperty.��ӡ��ʽ
                Case 0  '����ӡ
                Case 1
                    mblnPrintInvoice = True '�Զ���ӡ
                Case 2  '��ʾ��ӡ
                    If MsgBox("�Ƿ��ӡƱ��?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then mblnPrintInvoice = True
                End Select
            End If
        End If
        mblnNotify = True
    End If
    
    Call LockedScreen(True)
    If GetSaveBalanceSQL(cllPro) = False Then
        Call LockedScreen(False)      '����
        Exit Sub
    End If
    
    If ExcuteInsureSwapInteface(mBalanceInfor.lng����ID, cllPro) = False Then
        Call LockedScreen(False)      '����
        Exit Sub
    End If
    
    Call LockedScreen(False)      '����
    '���ؽ�����Ϣ
    lng����ID = Val(NVL(mrsInfo!����ID))
    Call LoadBalancePayData(lng����ID, mBalanceInfor.lng����ID, , , True)
'    Call RecalcDepositMoney(1)  '���°�ȱʡ����Ԥ��
    Call LoadIntendBalance
    mblnNotChange = True
    txtBalance(Idx_��Ԥ��).Text = Format(mBalanceInfor.dbl��Ԥ���ϼ�, "0.00")
    mblnNotChange = False
    '0-ҽ��Ԥ����Ϣ��ʾ;1-��ʾ������Ϣ
    Call ShowLedDisplayBank(1)
    Call ShowYBMoneyOrInvioceFormatInfor
    Call LoadCurOwnerPayInfor '���ص�ǰ֧����Ϣ
    'bytFun-0-����ǰ;1-ҽ����������;2-�ѱ����˽��ʵ�;
    Call SetOperationCtrl(2)
    If mBalanceInfor.dbl��Ԥ���ϼ� <> 0 Then
        '��궨λ���ɿ
        If txtBalance(Idx_��Ԥ��).Enabled And txtBalance(Idx_��Ԥ��).Visible Then txtBalance(Idx_��Ԥ��).SetFocus
    Else
        '��궨λ���ɿ
        If txtBalance(Idx_�ɿ�).Enabled And txtBalance(Idx_�ɿ�).Visible Then txtBalance(Idx_�ɿ�).SetFocus
    End If
    If mBalanceInfor.dbl��Ԥ���ϼ� = 0 And RoundEx(mBalanceInfor.dbl��ǰ���� - mBalanceInfor.dblҽ��֧���ϼ�, 5) = 0 Then cmdOK_Click
End Sub

Private Sub LockedScreen(ByVal blnLocked As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ˢ������������ִ�й����е����ؿؼ�
    '���:blnLocked-true,����,False-����
    '����:���˺�
    '����:2015-01-13 16:41:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnUnLocked As Boolean
    
    Screen.MousePointer = IIf(blnLocked, 99, 0)
    
    mblnLockScreen = blnLocked
    blnUnLocked = Not blnLocked
    mblnInvalidLoad = True
    picPati.Enabled = blnUnLocked
    mblnInvalidLoad = False
    picBalanceInfor.Enabled = blnUnLocked
    picCurDeposit.Enabled = blnUnLocked
    picCurBalance.Enabled = blnUnLocked
    cmdCancel.Enabled = blnUnLocked
    vsBlance.Enabled = blnUnLocked
    cmdOK.Enabled = blnUnLocked
    cmdYB.Enabled = blnUnLocked
    txtInvoice.Enabled = blnUnLocked
    picNO.Enabled = blnUnLocked
    picFeeList.Enabled = blnUnLocked
    picDeposit.Enabled = blnUnLocked
    
    
End Sub

Private Sub Form_Activate()
    '˫����ʾ��������ڵ�ǰ������ʾ֮�������ʾ�����ƶ�����
    If mblnUnload = True Then Unload Me: Exit Sub
    
    
    mblnFirst = False
    Call Led_ClearDisplayPatient
    
    If mstrInNO <> "" And mEditType = g_Ed_���ݲ鿴 Then
        '����ʱ
        If txtPatient.Text = "" Then Unload Me: Exit Sub
        If cboNO.Visible And cboNO.Enabled Then cboNO.SetFocus
    ElseIf mEditType = g_Ed_�������� Then
        If txtBalance(Idx_�ɿ�).Enabled And txtBalance(Idx_�ɿ�).Visible Then txtBalance(Idx_�ɿ�).SetFocus
'    Else
'        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    End If
    
    If mEditType = g_Ed_������� Or mEditType = g_Ed_סԺ���� Then
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    End If
    '111222,����,2017/08/03,�Ǽ����������Ĳ�����ʾcboӤ��
    Call picPati_Resize
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            'ȡ����ť
            If cmdCancel.Enabled And cmdCancel.Visible Then cmdCancel.SetFocus: Call cmdCancel_Click
        Case vbKeyF1
            ShowHelp App.ProductName, Me.hWnd, Me.Name
        Case vbKeyF2
            If cmdYBBalance.Enabled And cmdYBBalance.Visible Then cmdYBBalance.SetFocus: cmdYBBalance_Click: Exit Sub
            If cmdOK.Enabled And cmdOK.Visible Then cmdOK.SetFocus: Call cmdOK_Click
        Case vbKeyF4
            If Shift = vbCtrlMask Then
                If IDKIND.Enabled Then
                    Dim intIndex As Integer
                    intIndex = IDKIND.GetKindIndex("IC����")
                    If intIndex <= 0 Then Exit Sub
                    IDKIND.IDKIND = intIndex: Call IDKind_Click(IDKIND.GetCurCard)
                End If
                Exit Sub
            End If
            If Me.ActiveControl Is txtPatient Then
                If IDKIND.Enabled Then
                    If Shift = vbShiftMask Then
                        IDKIND.IDKIND = IIf(IDKIND.IDKIND = 0, UBound(Split(IDKIND.IDKindStr, ";")), IDKIND.IDKIND - 1)
                    Else
                        IDKIND.IDKIND = IIf(IDKIND.IDKIND = UBound(Split(IDKIND.IDKindStr, ";")), 0, IDKIND.IDKIND + 1)
                    End If
                End If
            End If
        Case vbKeyF6
            If cmdYB.Enabled And cmdYB.Visible Then cmdYB.SetFocus: Call cmdYB_Click
        Case vbKeyF8 '�˺ſ��
            chkCancel.Value = IIf(chkCancel.Value = 1, 0, 1)
        Case vbKeyF9 '��������
            If cmdRefresh.Enabled And cmdRefresh.Visible Then cmdRefresh.SetFocus: Call cmdRefresh_Click
        Case vbKeyF11 '��λ�����������
            If Not txtPatient.Locked And txtPatient.Enabled Then txtPatient.SetFocus
        Case vbKeyF12 '��λ�����ſ��ǿ�Ʊ���
            If Shift = vbCtrlMask Then
                'ǿ����LED����,(�ϼ�)
                mbln�ѱ��� = False
                Call LedVoiceSpeak(True)
            Else
                If Not cboNO.Locked And cboNO.Enabled Then cboNO.SetFocus
            End If
    End Select
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub
Private Sub Form_Load()
    If mblnInvalidLoad = True Then
        mintSucces = 1
        mblnUnload = True: Exit Sub
    End If
    mlngModul = 1137
    mblnFirst = True: mblnUnload = False
    Call RestoreWinState(Me, App.ProductName)
    Call InitGrid_PayList
    Call zlInitModulePara
    If Init���㷽ʽ = False Then Exit Sub
    '��ʼ������
    Call InitFace
    '��ʼ���˵��򹤾���
    Call zlDefCommandBars
    
    Call InitLed '��ʼ��Led

    If LoadBalanceBill = False Then mblnUnload = True: Exit Sub
    If mblnCancel Then
        chkCancel.Value = 1: chkCancel.Enabled = False
        mblnCancel = False
    End If
End Sub
Private Sub SetDefaultPayType()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ȱʡ��֧����ʽ
    '����:���˺�
    '����:2015-01-28 10:01:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnEnabled As Boolean
    Dim emEditType As gBalanceBill
    Dim strDefaultType As String, lng�����ID As Long
    Dim strBalance As String, strCash As String
    Dim dblʣ���� As Double, intKindIdx As Integer
    Dim i As Long, objCard As Card
    Dim blnFind As Boolean, dblMoney As Double
    
    On Error GoTo errHandle
    emEditType = mEditType
    If chkCancel.Value = 1 Then emEditType = g_Ed_��������
    
    Select Case emEditType
    Case g_Ed_�������, g_Ed_סԺ����, g_Ed_���½���
        strBalance = mstrȱʡ���㷽ʽ
        With mBalanceInfor
            dblʣ���� = RoundEx(.dblδ���ϼ� - .dbl��Ԥ���ϼ�, 5)
        End With
        If dblʣ���� >= 0 Then GoTo GoLocal:
        If mPatiInfor.dblδ���ۼ� <> 0 Then Exit Sub
        '�˿��ȱʡ��ʽ
        If mrsDeposit Is Nothing Then GoTo GoLocal:
        If mrsDeposit.State <> 1 Then GoTo GoLocal:
        If mrsDeposit.RecordCount = 0 Then GoTo GoLocal:
        If mty_ModulePara.bln�����˿ʽ Then
            mrsDeposit.Sort = "�����ID Desc,ת�ʼ�����,��������"
            With mrsDeposit
                .MoveFirst
                Do While Not .EOF
                    '1.������ʱ��ֻ�д��۵Ĳ���ȱʡ�˿�(��Ҫ�ǽ�Ԥ�������ܴ��ڶཻ���ף��ּ򵥴���)
                    If Val(NVL(!�����ID)) > 0 And NVL(NVL(!ת�ʼ�����, 0)) = 1 Then
                        '��鵱ǰ�Ƿ�֧�ַ����㿨
                        If Not GetLocalePayCard(Val(NVL(!�����ID)), False, intKindIdx) Is Nothing Then
                            IDKindPaymentsType.IDKIND = intKindIdx
                            Exit Sub
                        End If
                    End If
                    '2.���㷽ʽΪXX����,��ȱʡΪ�÷�ʽ
                    If Val(NVL(!��������)) = 2 And NVL(!���㷽ʽ) Like "*��" Then
                        strBalance = NVL(!���㷽ʽ): GoTo GoLocal:
                    End If
                    If Val(NVL(!��������)) = 1 Then strCash = NVL(!���㷽ʽ)
                    If Val(NVL(!��������)) = 2 And NVL(!���㷽ʽ) Like "*֧Ʊ" Then strBalance = NVL(!���㷽ʽ)
                    If strBalance = "" And Val(NVL(!��������)) = 2 Then
                        strBalance = NVL(!���㷽ʽ)
                    End If
                    .MoveNext
                Loop
                If strCash <> "" Then strBalance = strCash
                GoTo GoLocal:
            End With
        Else
            mrs���㷽ʽ.Filter = "ȱʡ = 1"
            If Not mrs���㷽ʽ.EOF Then
                strBalance = NVL(mrs���㷽ʽ!����)
            End If
            mrs���㷽ʽ.Filter = 0
        End If
    Case g_Ed_��������, g_Ed_��������
        With vsBlance
            
        End With
    Case Else
    End Select
GoLocal:
    '��λ
    blnFind = False
    
    blnEnabled = txtBalance(Idx_�ɿ�).Enabled '���������벡�˺󣬹���ƶ������������
    
    txtBalance(Idx_�ɿ�).Enabled = False
    For i = 1 To IDKindPaymentsType.ListCount
        'ȱʡ��λ���ֽ���
        Set objCard = IDKindPaymentsType.GetIDKindCard(i, CardTypeIndex)
        
        If strBalance = "" And objCard.�������� = 1 Then IDKindPaymentsType.IDKIND = i: blnFind = True: Exit For
        If objCard.���㷽ʽ = strBalance Then IDKindPaymentsType.IDKIND = i: blnFind = True: Exit For
    Next
    txtBalance(Idx_�ɿ�).Enabled = blnEnabled
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mBalanceInfor.blnSaveBill And (mEditType = g_Ed_סԺ���� Or mEditType = g_Ed_������� Or mEditType = g_Ed_��������) Then
        MsgBox "�Ѿ������˽��ʵ���,�����˳�!", vbInformation, gstrSysName
        Cancel = 1: Exit Sub
    End If
    If (mEditType = g_Ed_������� Or mEditType = g_Ed_סԺ����) And mstrInNO = "" And mYBInFor.strYBPati <> "" And Not mobjBalanceCon.blnCurBalanceOwnerFee Then
        If MsgBox("��ǰ���ڶ�ҽ�����˽��ʣ�ȷʵҪ�˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1: Exit Sub
        End If
        If YBIdentifyCancel = False Then        'ȡ��ҽ�����������֤,���ؼ�ʱ���˳�
            Cancel = 1: Exit Sub
        End If
    End If
    '�����ڲ���
    mlngPatientID = 0: mblnViewCancel = False: mstrInNO = ""
    mblnNOMoved = False: mstrPrivs = ""
    mlng����ID = 0: mbln����תסԺ = False
    mstr��ҳId = "": mstrPepositDate = ""
    mblnNotify = False
    mstrPatient = ""
    mstrȱʡ���㷽ʽ = ""
    Call ClearCustomType '����Զ���������ر���
 
    Call InitBalanceCondition
     
    Set mrsBalance = Nothing
    Set mrsFeeList = Nothing
    Set mrsDeposit = Nothing
    Set mrsOldBalance = Nothing
    Set mrsInfo = New ADODB.Recordset
    If mEditType <> g_Ed_���ݲ鿴 And gblnLED Then
        zl9LedVoice.DisplayPatient ""
        zl9LedVoice.Reset com
    End If
    If mEditType <> g_Ed_���ݲ鿴 Then
        Call SaveRegInFor(g˽��ģ��, Me.Name, "IDKIND", IDKIND.IDKIND)
    End If
    mlngBalanceRows = 0
    mblnBatchState = False
    Me.Visible = False
    mstrForceNote = ""
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
    If objCard Is Nothing Then Exit Sub
    txtPatient.Text = objPatiInfor.����
    Call FindPati(objCard, True, objPatiInfor.����)
End Sub
Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng�����ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
    If objCard.���� Like "IC��*" And objCard.ϵͳ Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        If mobjICCard Is Nothing Then Exit Sub
        txtPatient.Text = mobjICCard.Read_Card()
        If txtPatient.Text <> "" Then
            Call FindPati(objCard, True, txtPatient.Text)
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
    If txtPatient.Text <> "" Then Call FindPati(objCard, True, txtPatient.Text)
End Sub

Private Sub IDKindPaymentsType_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    If mblnNotChange = True Then Exit Sub
    Call LoadDefaultMoney
    
   If txtBalance(Idx_�ɿ�).Enabled And txtBalance(Idx_�ɿ�).Visible _
        And picCurBalance.Enabled And picBalanceBack.Enabled Then txtBalance(Idx_�ɿ�).SetFocus
    mblnNotChange = True
    Call LoadCurOwnerPayInfor
    mblnNotChange = False
End Sub
Private Sub LoadDefaultMoney()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ȱʡ�Ľɿ���˿���
    '����:���˺�
    '����:2015-01-30 17:38:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card
    Dim i As Long, blnHave As Boolean
    On Error GoTo errHandle
    txtBalance(Idx_�ɿ�).Text = "": txtBalance(Idx_�ɿ�).Locked = False
    Set objCard = IDKindPaymentsType.GetCurCard
    If objCard Is Nothing Then Exit Sub
        
    If mEditType = g_Ed_�������� Or mEditType = g_Ed_�������� Then
        blnHave = False
        If vsBlance.Rows >= 2 Then
            For i = 2 To vsBlance.Rows
                If objCard.���㷽ʽ = vsBlance.TextMatrix(i - 1, 0) Then
                    blnHave = True
                End If
            Next i
        End If
        If Not blnHave Then
'            If objCard.�ӿ���� > 0 Then
'                 If Not objCard.���ѿ� Then
'                     '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
'                     txtBalance(Idx_�ɿ�).Text = -1 * Format(GetOldBalanceMoney(3, objCard), "0.00")
'                 Else
'                     txtBalance(Idx_�ɿ�).Text = -1 * Format(GetOldBalanceMoney(5, objCard), "0.00")
'                 End If
'            Else
                 If objCard.�������� <> 1 Then
                    txtBalance(Idx_�ɿ�).Text = Format(Val(lblʣ���Ը�.Caption), "0.00")
                 End If
'            End If
       End If
       Exit Sub
    ElseIf mEditType = g_Ed_������� Or mEditType = g_Ed_סԺ���� Or mEditType = g_Ed_���½��� Then
        If objCard.�������� = 2 Then
            txtBalance(6).Enabled = True
            lblBalance(5).Enabled = True
        Else
            txtBalance(6).Enabled = False
            lblBalance(5).Enabled = False
        End If
        If objCard.�ӿ���� > 0 And Not objCard.���ѿ� Then '86853
            If mty_ModulePara.bytˢ��ȱʡ������ <> 0 Then
                txtBalance(Idx_�ɿ�).Text = Format(Val(lblʣ���Ը�.Caption), "0.00")
                If mty_ModulePara.bytˢ��ȱʡ������ = 2 Then txtBalance(Idx_�ɿ�).Locked = True
            End If
            Exit Sub
        End If
        If objCard.���ѿ� Then
            txtBalance(Idx_�ɿ�).Text = Format(Val(lblʣ���Ը�.Caption), "0.00")
            Exit Sub
        End If
        If objCard.�������� = 2 And mPatiInfor.bln�˿��־ Then
            txtBalance(Idx_�ɿ�).Text = Format(Val(lblʣ���Ը�.Caption), "0.00")
            Exit Sub
        End If
        If objCard.�������� <> 1 Then
            txtBalance(Idx_�ɿ�).Text = Format(Val(lblʣ���Ը�.Caption), "0.00")
        End If
        If objCard.�������� = 1 And mPatiInfor.bln�˿��־ And mty_ModulePara.bln�˿��ֽ�ȱʡ��� Then
            txtBalance(Idx_�ɿ�).Text = Format(Val(lblʣ���Ը�.Caption), "0.00")
        End If
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Sub


Private Sub IDKindPaymentsType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub IDKindPaymentsType_KeyPress(KeyAscii As Integer)
    Call MoveIDKindItem(IDKindPaymentsType, KeyAscii)
End Sub

Private Sub IDKind�Ҳ�_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    Dim dblTemp As Double
    If mblnNotChange Then Exit Sub
    mblnNotChange = True
    If objCard.�ӿ���� = 1 Then
        Call Set�Ҳ���Ϣ
        Call LoadCurOwnerPayInfor
    Else
        dblTemp = RoundEx(Abs(mBalanceInfor.dblδ���ϼ� - mBalanceInfor.dbl��Ԥ���ϼ�), 5)
        dblTemp = GetCentMoney(dblTemp)
        If RoundEx(Val(txtBalance(Idx_�ɿ�).Text) - dblTemp - mPatiInfor.dblδ���ۼ�, 6) >= 0 Then
            dblTemp = RoundEx(Val(txtBalance(Idx_�ɿ�).Text) - dblTemp - mPatiInfor.dblδ���ۼ�, 6)
            txtBalance(Idx_�Ҳ�).Text = Format(dblTemp, "0.00")
        Else
            txtBalance(Idx_�ɿ�).Text = "0.00"
            txtBalance(Idx_�Ҳ�).Text = Format(dblTemp, "0.00")
        End If
        
        Call LoadCurOwnerPayInfor
        If txtBalance(Idx_�ɿ�).Enabled And txtBalance(Idx_�ɿ�).Visible Then
            txtBalance(Idx_�ɿ�).SetFocus
            zlControl.TxtSelAll txtBalance(Idx_�ɿ�)
        End If
    End If
    mblnNotChange = False
End Sub

Private Sub IDKind�Ҳ�_KeyPress(KeyAscii As Integer)
    Call MoveIDKindItem(IDKindPaymentsType, KeyAscii)
End Sub

Private Sub opt��Ժ_Click()
    Dim dtBeginDate As Date, dtEndDate As Date
    If GetPatiHospitalzedDateRange(dtBeginDate, dtEndDate) Then
        txtPatiBegin.Text = Format(dtBeginDate, txtPatiBegin.Format)
        txtPatiEnd.Text = Format(dtEndDate, txtPatiEnd.Format)
        txtPatiEnd.Tag = Format(dtEndDate, txtPatiEnd.Format)
        Call zlChangeDefaultTime
    End If
    If IsDate(txtPatiEnd.Text) = False Or IsDate(txtPatiBegin.Text) = False Then Exit Sub
    txt����.Text = CDate(txtPatiEnd.Text) - CDate(txtPatiBegin.Text) + IIf(opt��;.Value = True, 1, 0)
    If Val(txt����.Text) = 0 Then txt����.Text = 1
End Sub

Private Sub opt��Ժ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub opt��;_Click()
    Dim dtBeginDate As Date, dtEndDate As Date
    If GetPatiHospitalzedDateRange(dtBeginDate, dtEndDate) Then
        txtPatiBegin.Text = Format(dtBeginDate, txtPatiBegin.Format)
        txtPatiEnd.Text = Format(dtEndDate, txtPatiEnd.Format)
        txtPatiEnd.Tag = Format(dtEndDate, txtPatiEnd.Format)
        Call zlChangeDefaultTime
    End If
    If IsDate(txtPatiEnd.Text) = False Or IsDate(txtPatiBegin.Text) = False Then Exit Sub
    txt����.Text = CDate(txtPatiEnd.Text) - CDate(txtPatiBegin.Text) + IIf(opt��;.Value = True, 1, 0)
    If Val(txt����.Text) = 0 Then txt����.Text = 1
End Sub

Private Sub opt��;_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub picBalanceBack_Resize()
    Dim lngStep As Long
    Err = 0: On Error Resume Next
    lngStep = 100
    With picBalanceBack
        shpBalance.Left = .ScaleLeft
        shpBalance.Top = .ScaleTop
        shpBalance.Width = .ScaleWidth
        shpBalance.Height = .ScaleHeight
        '��ǰ������Ϣ
        picBalanceInfor.Left = .ScaleLeft + lngStep
        picBalanceInfor.Top = .ScaleTop + lngStep
        picBalanceInfor.Width = .ScaleWidth - picBalanceInfor.Left - lngStep
        
        '��ǰδ����
        picNotPayment.Top = picBalanceInfor.Top + picBalanceInfor.Height + lngStep
        picNotPayment.Left = picBalanceInfor.Left
        
        '��ǰ�Ը�
        picOwerFee.Left = picBalanceInfor.Left
        picOwerFee.Top = picNotPayment.Top + picNotPayment.Height + lngStep
        
        '��ǰ��Ԥ��
        picCurDeposit.Top = picNotPayment.Top
        picCurDeposit.Left = picNotPayment.Left + picNotPayment.Width + lngStep
        picCurDeposit.Width = .ScaleWidth - picCurDeposit.Left - lngStep
        
        '��ǰ������Ϣ
        picCurBalance.Top = picCurDeposit.Top + picCurDeposit.Height + lngStep
        picCurBalance.Left = picCurDeposit.Left
        picCurBalance.Width = .ScaleWidth - picCurBalance.Left - lngStep
        
        cmdCancel.Left = .ScaleWidth - cmdCancel.Width - 60
        cmdCancel.Top = picCurBalance.Top + picCurBalance.Height + lngStep
        cmdDelBalance.Left = cmdCancel.Left
        cmdDelBalance.Top = cmdCancel.Top
        shpFun.Top = cmdCancel.Top - 50
        shpFun.Left = picOwerFee.Left - 20
        shpFun.Width = .ScaleWidth - shpFun.Left - 50
        
        cmdOK.Left = IIf(cmdCancel.Visible Or cmdDelBalance.Visible, cmdCancel.Left, .ScaleWidth) - cmdOK.Width - 60
        cmdOK.Top = cmdCancel.Top
        
        cmdYBBalance.Left = cmdOK.Left '- cmdYBBalance.Width - 60
        cmdYBBalance.Top = cmdCancel.Top
        
        cmdNext.Left = cmdOK.Left - cmdNext.Width - 60
        cmdNext.Top = cmdCancel.Top
        
        If mEditType <> g_Ed_���ݲ鿴 Then
            vsBlance.Top = cmdCancel.Top + cmdCancel.Height + lngStep
            '��֧���б�
            lblCurPaymentInfor.Top = cmdYBBalance.Top + cmdYBBalance.Height - lblCurPaymentInfor.Height
        Else
            lblCurPaymentInfor.Top = .ScaleTop + lngStep
            lblCurPaymentInfor.Left = .ScaleLeft + lngStep
            vsBlance.Top = lblCurPaymentInfor.Top + lblCurPaymentInfor.Height + lngStep
            picBalanceInfor.Top = .ScaleHeight - picBalanceInfor.Height - lngStep
            vsBlance.Height = picBalanceInfor.Top - vsBlance.Top - lngStep
        End If
        vsBlance.Left = picOwerFee.Left
        vsBlance.Width = .ScaleWidth - vsBlance.Left - lngStep
        lbl����.Top = lblCurPaymentInfor.Top
        If mEditType <> g_Ed_���ݲ鿴 Then
            Call ShowYBMoneyOrInvioceFormatInfor
        End If
        lbl����.Left = lblCurPaymentInfor.Left + lblCurPaymentInfor.Width + 100
    End With
    Call picCurBalance_Resize
    Call picBalanceInfor_Resize
    Call picCurDeposit_Resize
    
    cboChargeType.RaisEffect picBalanceBack, Dw_Heave
End Sub
Private Sub ShowYBMoneyOrInvioceFormatInfor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾҽ���������Ʊ�ݸ�ʽ
    '����:���˺�
    '����:2015-01-15 11:12:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngStep As Long
    Err = 0: On Error Resume Next
    lngStep = 100
    With picBalanceBack
        
        lblҽ������.Top = .ScaleHeight - lblҽ������.Height - 100
        lblҽ������.Left = picOwerFee.Left
        lblBalance(7).Top = lblҽ������.Top
        txtDate.Top = lblBalance(7).Top - 60
'        lbl�����ʻ�.Top = lblҽ������.Top
'        lbl�����ʻ�.Left = lblҽ������.Left + lblҽ������.Width + 1000
        lbl����.Top = lblCurPaymentInfor.Top
        
'        lblFormat.Top = lblҽ������.Top
        vsBlance.Height = lblҽ������.Top - vsBlance.Top - lngStep
    End With
End Sub
Private Sub picCurDeposit_Resize()
    Err = 0: On Error Resume Next
    With picCurDeposit
        txtBalance(Idx_��Ԥ��).Width = .ScaleWidth - txtBalance(Idx_��Ԥ��).Left - 100
    End With
    cboChargeType.RaisEffect picCurDeposit, Dw_SubKen
End Sub
Private Sub picBalanceInfor_Resize()
    Err = 0: On Error Resume Next
    With picBalanceInfor
        txtBalance(Idx_����˵��).Width = .ScaleWidth - txtBalance(Idx_����˵��).Left - 100
        txtBalance(Idx_���ν���).Width = .ScaleWidth - txtBalance(Idx_���ν���).Left - 100
        txtBalance(Idx_����δ��).Width = txtBalance(Idx_���ν���).Width
    End With
    cboChargeType.RaisEffect picBalanceInfor, Dw_SubKen
End Sub

Private Sub picCurBalance_Resize()
    Err = 0: On Error Resume Next
    With picCurBalance
        txtBalance(Idx_�ɿ�).Width = .ScaleWidth - txtBalance(Idx_�ɿ�).Left - 100
        txtBalance(Idx_�Ҳ�).Width = .ScaleWidth - txtBalance(Idx_�Ҳ�).Left - 100
        txtBalance(Idx_�������).Width = .ScaleWidth - txtBalance(Idx_�������).Left - 100
        txtBalance(Idx_ժҪ).Width = .ScaleWidth - txtBalance(Idx_ժҪ).Left - 100
    End With
    cboChargeType.RaisEffect picCurBalance, Dw_SubKen
End Sub

Private Sub picDetailContain_Resize()
    On Error Resume Next
    With vsDetailList
        .Top = 0
        .Left = 0
        .Height = picDetailContain.ScaleHeight
        .Width = picDetailContain.ScaleWidth
    End With
End Sub

Private Sub picFeeContain_Resize()
    On Error Resume Next
    With vsFeeList
        .Top = 0
        .Left = 0
        .Height = 3000
        .Width = picFeeContain.ScaleWidth
    End With
    With picDeposit
        .Top = vsFeeList.Top + vsFeeList.Height + 60
        .Left = 0
        .Width = picFeeContain.ScaleWidth
        .Height = picFeeContain.ScaleHeight - .Top - 30
    End With
End Sub

Private Sub picFeeList_Resize()
    Err = 0: On Error Resume Next
    With picFeeList
        tabFeeList.Left = 15
        tabFeeList.Top = 15
        tabFeeList.Height = .ScaleHeight - 30
        tabFeeList.Width = .ScaleWidth - 30
        
        picFeeContain.Left = 15
        picFeeContain.Top = 330
        picFeeContain.Width = .ScaleWidth - 30
        picFeeContain.Height = .ScaleHeight - 340
        
        picDetailContain.Left = 15
        picDetailContain.Top = 330
        picDetailContain.Width = .ScaleWidth - 30
        picDetailContain.Height = .ScaleHeight - 340
        
        If tabFeeList.Tab = 1 Then
            picDetailContain.Visible = True
            picFeeContain.Visible = False
        Else
            picDetailContain.Visible = False
            picFeeContain.Visible = True
        End If
    End With
    
    Call cboChargeType.RaisEffect(picFeeList, Dw_Heave)
End Sub

Private Sub picDeposit_Resize()
    Err = 0: On Error Resume Next
    With picDeposit
        vsDeposit.Left = 15
        vsDeposit.Top = lblDeposit.Top + lblDeposit.Height + 50
        vsDeposit.Height = .ScaleHeight - vsDeposit.Top - 30
        
        cmdDepositUp.Top = vsDeposit.Top + vsDeposit.Height / 4
        cmdDepositDown.Top = cmdDepositUp.Top + cmdDepositUp.Height + 250
        cmdDepositUp.Left = .ScaleWidth - cmdDepositUp.Width - 100
        cmdDepositDown.Left = cmdDepositUp.Left
        
        If cmdDepositUp.Visible Then
            vsDeposit.Width = cmdDepositUp.Left - vsDeposit.Left - 60
        Else
            vsDeposit.Width = .ScaleWidth - vsDeposit.Left - 100
        End If
        
        lnDepositSplit.X1 = .ScaleWidth - 10
        lnDepositSplit.X2 = lnDepositSplit.X1
        lnDepositSplit.Y1 = .ScaleTop
        lnDepositSplit.Y2 = .ScaleHeight
        cmdTools.Left = .ScaleWidth - cmdTools.Width - 100
    End With
End Sub

Private Sub SetUpDown()
    With vsDeposit
        cmdDepositUp.Enabled = True
        cmdDepositDown.Enabled = True
        If .Row = 1 Then cmdDepositUp.Enabled = False
        If .Row = .Rows - 1 Then cmdDepositDown.Enabled = False
    End With
End Sub

Private Sub picNO_Resize()
    Err = 0: On Error Resume Next
    With picNO
        chkCancel.Left = .ScaleWidth - chkCancel.Width
        chkCancel.Top = .ScaleTop
        lblDelCaption.Left = .ScaleWidth - lblDelCaption.Width
        lblDelCaption.Top = .ScaleTop
        
        cboNO.Left = .ScaleLeft
        cboNO.Top = .ScaleTop
        If Not (mEditType = g_Ed_������� Or mEditType = g_Ed_סԺ����) Then
            If mblnViewCancel Or mEditType = g_Ed_ȡ������ Or mEditType = g_Ed_�������� Or mEditType = g_Ed_�������� Then
                cboNO.Width = lblDelCaption.Left - cboNO.Left - 30
            Else
                cboNO.Width = .ScaleWidth
            End If
        Else
            cboNO.Width = chkCancel.Left - cboNO.Left - 30
        End If
        cboNO.Height = .ScaleHeight
    End With
End Sub
 
Private Sub AddPopu()
    Dim vRect As RECT
    vRect = zlControl.GetControlRect(cmdTools.hWnd)
    vRect.Left = vRect.Left + 10
    vRect.Top = vRect.Top + 50
    Call CreatePopuMenu
    If Not mobjCommandBar Is Nothing Then Call mobjCommandBar.ShowPopup(, vRect.Left, vRect.Top + cmdTools.Height)
End Sub

Private Sub CreatePopuMenu()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ʱ�˵�
    '����:���˺�
    '����:2012-11-21 09:49:35
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim j As Long
    Dim objCustom As CommandBarControlCustom
   
    Set mobjCommandBar = cbsThis.Add("PopupPati", xtpBarPopup)
    With mobjCommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_NotUseDeposit, "��ʹ��Ԥ����(&C)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_UseAllDeposit, "ʹ������Ԥ����(&A)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_MoneyUseDeposit, "�����ʽ��ʹ��Ԥ��(&J)")
    End With
End Sub

Private Function InitGrid() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������ؼ�
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-12-29 15:08:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    On Error GoTo errHandle
    With vsDeposit
        .Clear
        .Cols = 20: .Rows = 2
        i = 0
        .TextMatrix(0, i) = "ID": i = i + 1
        .TextMatrix(0, i) = "���ݺ�": i = i + 1
        .TextMatrix(0, i) = "Ʊ�ݺ�": i = i + 1
        .TextMatrix(0, i) = "�տ�����": i = i + 1
        .TextMatrix(0, i) = "���㷽ʽ": i = i + 1
        .TextMatrix(0, i) = "���": i = i + 1
        .TextMatrix(0, i) = "��Ԥ��": i = i + 1
        .TextMatrix(0, i) = "���": i = i + 1
        .TextMatrix(0, i) = "Ԥ��ID": i = i + 1
        .TextMatrix(0, i) = "�༭״̬": i = i + 1
        .TextMatrix(0, i) = "�����ID": i = i + 1
        .TextMatrix(0, i) = "�Ƿ����ѿ�": i = i + 1
        .TextMatrix(0, i) = "���������": i = i + 1
        .TextMatrix(0, i) = "����": i = i + 1
        .TextMatrix(0, i) = "������ˮ��": i = i + 1
        .TextMatrix(0, i) = "����˵��": i = i + 1
        .TextMatrix(0, i) = "�Ƿ�����": i = i + 1
        .TextMatrix(0, i) = "�Ƿ�ȫ��": i = i + 1
        .TextMatrix(0, i) = "�Ƿ�ȱʡ����": i = i + 1
        .TextMatrix(0, i) = "�Ƿ�ת�ʼ�����": i = i + 1
        For i = 0 To .Cols - 1
            .ColKey(i) = .TextMatrix(0, i)
            .FixedCols = 1
            .FixedAlignment(i) = flexAlignCenterCenter
            .ColAlignment(i) = flexAlignLeftCenter
            
            ''ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)|������(0-��������,1-��ֹ����,2-��������,�����س���������)
            Select Case .ColKey(i)
            Case "���ݺ�"
                .ColData(i) = "1|0"
                .FixedAlignment(i) = flexAlignRightCenter
            Case "���"
                 If mEditType = g_Ed_������� Or mEditType = g_Ed_סԺ���� _
                    Or mEditType = g_Ed_���½��� Then
                    .ColData(i) = "0|0"
                    .ColHidden(i) = False
                 Else
                      .ColHidden(i) = True: .ColData(i) = "-1|1"
                 End If
            Case "��Ԥ��"
                    .ColData(i) = "1|0"
                    .ColHidden(i) = False
            Case "���"
                 If mEditType = g_Ed_������� Or mEditType = g_Ed_סԺ���� Or mEditType = g_Ed_���½��� Then
                     .ColHidden(i) = True: .ColData(i) = "0|1"
                 Else
                      .ColHidden(i) = True: .ColData(i) = "-1|0"
                 End If
            Case "���������", "����", "������ˮ��", "����˵��"
                 .ColHidden(i) = True: .ColData(i) = "0|0"
            Case Else
                If Not .ColKey(i) Like "*ID" Then
                    .ColData(i) = "0|0"
                End If
            End Select
            
            If InStr(",�Ƿ����ѿ�,�Ƿ�����,�Ƿ�ȫ��,�Ƿ�ȱʡ����,�Ƿ�ת�ʼ�����,�༭״̬,", "," & .ColKey(i) & ",") > 0 Or .ColKey(i) Like "*ID" Then
                .ColHidden(i) = True: .ColWidth(i) = 0
                .ColData(i) = "-1|1"
            ElseIf .ColKey(i) Like "*��" Or .ColKey(i) Like "*��Ԥ��" Then
                .ColAlignment(i) = flexAlignRightCenter
            End If
          
        Next
        .ColHidden(.ColIndex("Ʊ�ݺ�")) = True
        
        .ColWidth(.ColIndex("Ʊ�ݺ�")) = 1100
        .ColWidth(.ColIndex("�տ�����")) = 1200
        .ColWidth(.ColIndex("���ݺ�")) = 1100
        .ColWidth(.ColIndex("���㷽ʽ")) = 1400
        .ColWidth(.ColIndex("���")) = 1100
        .ColWidth(.ColIndex("��Ԥ��")) = 1100
        .ColWidth(.ColIndex("���������")) = 1800
        .ColWidth(.ColIndex("����")) = 1100
        .ColWidth(.ColIndex("������ˮ��")) = 1100
        .ColWidth(.ColIndex("����˵��")) = 1600
        
        .ExtendLastCol = False
        zl_vsGrid_Para_Restore mlngModul, vsDeposit, Me.Name, "Ԥ���б�"
        If mEditType = g_Ed_���ݲ鿴 Then
             .ColHidden(.ColIndex("���")) = True: .ColData(.ColIndex("���")) = "-1|1"
        End If
    End With
    With vsDetailList
        .FocusRect = flexFocusSolid
        .SelectionMode = flexSelectionFree
        .AllowBigSelection = False
        .HighLight = flexHighlightWithFocus
    End With
    Call InitTride_FeeList
    
    '������Ϣ
'    Call InitGrid_PayList
    InitGrid = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub InitTride_FeeList()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ�������б�
    '����:���˺�
    '����:2015-01-23 17:23:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    On Error GoTo errHandle
    With vsFeeList
        .Clear
        .Cols = 5: .Rows = 2
        i = 0
        .TextMatrix(0, i) = "��Ŀ": i = i + 1
        .TextMatrix(0, i) = "Ӧ�ս��": i = i + 1
        .TextMatrix(0, i) = "ʵ�ս��": i = i + 1
        .TextMatrix(0, i) = "δ����": i = i + 1
        .TextMatrix(0, i) = "���ʽ��": i = i + 1
        For i = 0 To .Cols - 1
            .ColKey(i) = .TextMatrix(0, i)
            .FixedAlignment(i) = flexAlignCenterCenter
            .ColAlignment(i) = flexAlignLeftCenter
            ''ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)|������(0-��������,1-��ֹ����,2-��������,�����س���������)
            If .ColKey(i) Like "*ID" Then
                .ColHidden(i) = True: .ColWidth(i) = 0
                .ColData(i) = "-1|1"
            ElseIf .ColKey(i) Like "*��" Then
                .ColAlignment(i) = flexAlignRightCenter
            End If
        Next
        .ColWidth(.ColIndex("��Ŀ")) = 2000
        .ColWidth(.ColIndex("Ӧ�ս��")) = 1400
        .ColWidth(.ColIndex("ʵ�ս��")) = 1400
        .ColWidth(.ColIndex("δ����")) = 1400
        .ColWidth(.ColIndex("���ʽ��")) = 1400
        
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        zl_vsGrid_Para_Restore mlngModul, vsFeeList, Me.Name, "�����б�" & mEditType
    End With
    Call SetFeeListColumnShow
    With vsDetailList
        .Clear
        .Cols = 10: .Rows = 2
        i = 0
        .TextMatrix(0, i) = "����": i = i + 1
        .TextMatrix(0, i) = "����": i = i + 1
        .TextMatrix(0, i) = "��Ŀ": i = i + 1
        .TextMatrix(0, i) = "δ����": i = i + 1
        .TextMatrix(0, i) = "���ʽ��": i = i + 1
        .TextMatrix(0, i) = "ID": i = i + 1
        .TextMatrix(0, i) = "��¼����": i = i + 1
        .TextMatrix(0, i) = "��¼״̬": i = i + 1
        .TextMatrix(0, i) = "ִ��״̬": i = i + 1
        .TextMatrix(0, i) = "���": i = i + 1
        For i = 0 To .Cols - 1
            .ColKey(i) = .TextMatrix(0, i)
            .FixedAlignment(i) = flexAlignCenterCenter
            .ColAlignment(i) = flexAlignLeftCenter
            ''ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)|������(0-��������,1-��ֹ����,2-��������,�����س���������)
            If .ColKey(i) Like "*ID" Then
                .ColHidden(i) = True: .ColWidth(i) = 0
                .ColData(i) = "-1|1"
            ElseIf .ColKey(i) Like "*��" Then
                .ColAlignment(i) = flexAlignRightCenter
            ElseIf .ColKey(i) = "��¼����" Or .ColKey(i) = "��¼״̬" Or .ColKey(i) = "ִ��״̬" Or .ColKey(i) = "���" Or .ColKey(i) = "��Ŀ" Then
                .ColHidden(i) = True: .ColWidth(i) = 0
                .ColData(i) = "-1|1"
            End If
            If InStr(",����,����,", "," & .ColKey(i) & ",") > 0 Then .ColAlignment(i) = flexAlignCenterCenter
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        .ColWidth(.ColIndex("����")) = 1400
        .ColWidth(.ColIndex("����")) = 1100
        .ColWidth(.ColIndex("��Ŀ")) = 2800
        .ColWidth(.ColIndex("δ����")) = 1400
        .ColWidth(.ColIndex("���ʽ��")) = 1400
    End With
    zl_vsGrid_Para_Restore mlngModul, vsDetailList, Me.Name, "��ϸ�б�" & mEditType
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub SetFeeListColumnShow()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���÷��ñ���ʾ��Ϣ
    '����:���˺�
    '����:2015-01-23 17:29:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    With vsFeeList
        If (mEditType = g_Ed_������� Or mEditType = g_Ed_סԺ����) And chkCancel.Value = 0 Then
            .ColHidden(.ColIndex("���ʽ��")) = True: .ColWidth(.ColIndex("���ʽ��")) = 0
        Else
            .ColHidden(.ColIndex("δ����")) = True: .ColWidth(.ColIndex("δ����")) = 0
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub InitGrid_PayList()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��֧���б�
    '����:���˺�
    '����:2015-01-23 14:14:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    
    On Error GoTo errHandle
    With vsBlance
        .Clear: .Rows = 2: i = 0: .Cols = 20
        .TextMatrix(0, i) = "�����ID": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "���ѿ�ID": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "��������": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "�༭״̬": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "����": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "����״̬": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "�Ƿ�����": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "�Ƿ�ȫ��": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "У�Ա�־": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "�Ƿ�����": .ColWidth(i) = 0: i = i + 1
        
        .TextMatrix(0, i) = "���㷽ʽ": .ColWidth(i) = 2000: i = i + 1
        .TextMatrix(0, i) = "������": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "�������": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "����": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "������ˮ��": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "����˵��": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "��ע": .ColWidth(i) = 2500: i = i + 1
        .TextMatrix(0, i) = "���������": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "�����Ϣ": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "�Ƿ�ת��": .ColWidth(i) = 0: i = i + 1
        
        For i = 0 To .Cols - 1
            .ColKey(i) = .TextMatrix(0, i)
            .ColAlignment(i) = flexAlignLeftCenter
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) Like "*ID" Then .ColHidden(i) = True
            Select Case .ColKey(i)
            Case "�Ƿ�ת��", "�����Ϣ", "��������", "����", "�Ƿ񱣴�", "�Ƿ�����", "У�Ա�־", "�༭״̬", "�Ƿ�����", "�Ƿ�ȫ��", "���������", "����״̬", "�Ƿ���֤"
                .ColHidden(i) = True
            Case "������"
                .ColAlignment(i) = flexAlignRightCenter
            End Select
        Next
        zl_vsGrid_Para_Restore mlngModul, vsBlance, Me.Name, "�����б�"
        If Not mEditType = g_Ed_���ݲ鿴 Then
            .Editable = flexEDKbdMouse
        End If
    End With

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Sub
 
Private Sub picPati_Resize()
    Err = 0: On Error Resume Next
    With picPati
        lnPatiSplit.Y1 = .ScaleHeight - 10
        lnPatiSplit.Y2 = .ScaleHeight - 10
        txtSex.Width = 600 * (.ScaleWidth / 14000)
        lblOld.Left = txtSex.Left + txtSex.Width + 100
        txtOld.Left = lblOld.Left + lblOld.Width + 30
        txtOld.Width = 1000 * (.ScaleWidth / 14000)
        lbl�ѱ�.Left = txtOld.Left + txtOld.Width + 100
        txt�ѱ�.Left = lbl�ѱ�.Left + lbl�ѱ�.Width + 30
        txt�ѱ�.Width = 1560 * (.ScaleWidth / 14000)
        lbl��ʶ��.Left = txt�ѱ�.Left + txt�ѱ�.Width + 100
        txt��ʶ��.Left = lbl��ʶ��.Left + lbl��ʶ��.Width + 30
        txt��ʶ��.Width = 1500 * (.ScaleWidth / 14000)
        lblBed.Left = txt��ʶ��.Left + txt��ʶ��.Width + 100
        txtBed.Left = lblBed.Left + lblBed.Width + 30
        txtBed.Width = 780 * (.ScaleWidth / 14000)
        lbl����.Left = txtBed.Left + txtBed.Width + 100
        txt����.Left = lbl����.Left + lbl����.Width + 30
        txt����.Width = 1440 * (.ScaleWidth / 14000)
        cboPatiNums.Width = .ScaleWidth - cboPatiNums.Left - 50
        cboӤ��.Left = cboPatiNums.Left + cboPatiNums.Width - cboӤ��.Width
        If cboӤ��.Visible Then
            cboPatiNums.Width = cboӤ��.Left - cboPatiNums.Left - 30
        End If
        cboChargeType.Width = .ScaleWidth - cboChargeType.Left - 50
        cmdRefresh.Left = .ScaleWidth - cmdRefresh.Width - 350
    End With
End Sub

'Private Sub SetLine()
'    On Error Resume Next
'    lneSex.Y1 = txtSex.Top + txtSex.Height
'    lneSex.Y2 = txtSex.Top + txtSex.Height
'    lneSex.X1 = txtSex.Left
'    lneSex.X2 = txtSex.Left + txtSex.Width
'
'    lneOld.Y1 = txtOld.Top + txtOld.Height
'    lneOld.Y2 = txtOld.Top + txtOld.Height
'    lneOld.X1 = txtOld.Left
'    lneOld.X2 = txtOld.Left + txtOld.Width
'
'    lne�ѱ�.Y1 = txt�ѱ�.Top + txt�ѱ�.Height
'    lne�ѱ�.Y2 = txt�ѱ�.Top + txt�ѱ�.Height
'    lne�ѱ�.X1 = txt�ѱ�.Left
'    lne�ѱ�.X2 = txt�ѱ�.Left + txt�ѱ�.Width
'
'    lne��ʶ��.Y1 = txt��ʶ��.Top + txt��ʶ��.Height
'    lne��ʶ��.Y2 = txt��ʶ��.Top + txt��ʶ��.Height
'    lne��ʶ��.X1 = txt��ʶ��.Left
'    lne��ʶ��.X2 = txt��ʶ��.Left + txt��ʶ��.Width
'
'    lneBed.Y1 = txtBed.Top + txtBed.Height
'    lneBed.Y2 = txtBed.Top + txtBed.Height
'    lneBed.X1 = txtBed.Left
'    lneBed.X2 = txtBed.Left + txtBed.Width
'
'    lne����.Y1 = txt����.Top + txt����.Height
'    lne����.Y2 = txt����.Top + txt����.Height
'    lne����.X1 = txt����.Left
'    lne����.X2 = txt����.Left + txt����.Width
'
'    lnePatiType.Y1 = txtPatiType.Top + txtPatiType.Height
'    lnePatiType.Y2 = txtPatiType.Top + txtPatiType.Height
'    lnePatiType.X1 = txtPatiType.Left
'    lnePatiType.X2 = txtPatiType.Left + txtPatiType.Width
'
'    lneԤ�����.Y1 = txtԤ�����.Top + txtԤ�����.Height
'    lneԤ�����.Y2 = txtԤ�����.Top + txtԤ�����.Height
'    lneԤ�����.X1 = txtԤ�����.Left
'    lneԤ�����.X2 = txtԤ�����.Left + txtԤ�����.Width
'
'    lneδ�����.Y1 = txtδ�����.Top + txtδ�����.Height
'    lneδ�����.Y2 = txtδ�����.Top + txtδ�����.Height
'    lneδ�����.X1 = txtδ�����.Left
'    lneδ�����.X2 = txtδ�����.Left + txtδ�����.Width
'
'    lneʣ����.Y1 = txtʣ����.Top + txtʣ����.Height
'    lneʣ����.Y2 = txtʣ����.Top + txtʣ����.Height
'    lneʣ����.X1 = txtʣ����.Left
'    lneʣ����.X2 = txtʣ����.Left + txtʣ����.Width
'End Sub

Private Sub Set�Ҳ���Ϣ()
    Dim dblMoney As Double
    Dim dbl��ǰδ�� As Double
    Dim objCard As Card
    Dim objBackCard As Card
    Dim objCards As Cards
    Dim objTemp As Card
    
    
    dblMoney = Val(txtBalance(Idx_�ɿ�).Text)
    Set objCard = IDKindPaymentsType.GetCurCard
    Set objBackCard = IDKind�Ҳ�.GetCurCard
    With mBalanceInfor
        dbl��ǰδ�� = RoundEx(.dbl��ǰ���� - .dbl��Ԥ���ϼ� - .dbl�Ѹ��ϼ�, 6)
    End With
    
    If mEditType = g_Ed_�������� Or mEditType = g_Ed_�������� Or chkCancel.Value = 1 Then
        dbl��ǰδ�� = -1 * dbl��ǰδ��
    End If
    
    dbl��ǰδ�� = RoundEx(dbl��ǰδ�� + mPatiInfor.dblδ���ۼ�, 6)
    
    If dblMoney = 0 Or objCard Is Nothing Then
        '0-ֻ���Ҳ�;1-����Ԥ��
        If dbl��ǰδ�� >= 0 Then
            Call Load�Ҳ���(0, "Ӧ�տ� ")
        Else
            Call Load�Ҳ���(1, "Ӧ�տ� ")
        End If
        txtBalance(Idx_�Ҳ�).ForeColor = txtBalance(Idx_����˵��).ForeColor
        If Not objBackCard Is Nothing Then
            If Not objBackCard.�ӿ���� <> 1 Then
                txtBalance(Idx_�Ҳ�).Text = "0.00"
            End If
        Else
            txtBalance(Idx_�Ҳ�).Text = "0.00"
        End If
        Exit Sub
    End If
        
    If dbl��ǰδ�� < 0 Then
        '��ǰ״̬Ϊ�˿�
        dbl��ǰδ�� = RoundEx(Val(lblʣ���Ը�.Caption), 6)
'        IDKind�Ҳ�.IDKind = 0   '֤���ǰ�ĳ�ַ�ʽ�˿�,��˲��ܴ�ΪԤ����
        If objCard.�������� = 1 Then
            'ֻ���ֽ�,�Ż����˿�ʱ�ึ������,����:��100������,Ҫ�һ�50
            If Abs(dbl��ǰδ��) <= dblMoney Then
                txtBalance(Idx_�Ҳ�).Text = Format(dblMoney - Abs(dbl��ǰδ��), "0.00")
                '0-ֻ���Ҳ�;1-����Ԥ��
                Call Load�Ҳ���(0, "Ӧ�տ� ")
                txtBalance(Idx_�Ҳ�).ForeColor = vbRed
            Else
                txtBalance(Idx_�Ҳ�).Text = Format(Abs(dblMoney - Abs(dbl��ǰδ��)), "0.00")
                '0-ֻ���Ҳ�;1-����Ԥ��
                Call Load�Ҳ���(0, "Ӧ�˿� ")
                txtBalance(Idx_�Ҳ�).ForeColor = txtBalance(Idx_����˵��).ForeColor
            End If
            Exit Sub
        End If
        If Abs(dbl��ǰδ��) < dblMoney Then
            '�������㷽ʽ��,ֻ����ʣ��δ�˿�,����:�˿���ҽԺ��֧Ʊ������,��˲������һ�֧Ʊ�Ŀ���
'            MsgBox "��ǰ�˿���(" & objCard.���� & ")��֧�ֶ���,����������!", vbInformation, gstrSysName
'            mblnNotChange = True
'            txtBalance(Idx_�ɿ�).Text = ""
'            mblnNotChange = False
            txtBalance(Idx_�Ҳ�).Text = Format(Abs(dblMoney - Abs(dbl��ǰδ��)), "0.00")
            Call Load�Ҳ���(0, "Ӧ�տ� ")
            txtBalance(Idx_�Ҳ�).ForeColor = txtBalance(Idx_����˵��).ForeColor
'            txtBalance(Idx_�Ҳ�).Text = "0.00": Exit Sub
        Else
            txtBalance(Idx_�Ҳ�).Text = Format(Abs(dblMoney - Abs(dbl��ǰδ��)), "0.00")
            '0-ֻ���Ҳ�;1-����Ԥ��
            Call Load�Ҳ���(0, "Ӧ�˿� ")
            txtBalance(Idx_�Ҳ�).ForeColor = txtBalance(Idx_����˵��).ForeColor
        End If
        Exit Sub
    End If
    
    '��ǰ״̬Ϊ�տ�
    dbl��ǰδ�� = RoundEx(Val(lblʣ���Ը�.Caption), 6)
'    IDKind�Ҳ�.IDKind = 0
    If objCard.�������� = 1 Then
        'ֻ���ֽ�,�Ż����˿�ʱ�ึ������,����:��100������,Ҫ�һ�50
        If dbl��ǰδ�� >= dblMoney Then
            '��Ҫ��ȡ����Ǯ
            txtBalance(Idx_�Ҳ�).Text = Format(Abs(dblMoney - Abs(dbl��ǰδ��)), "0.00")
            '0-ֻ���Ҳ�;1-����Ԥ��
            Call Load�Ҳ���(0, "Ӧ�տ� ")
            txtBalance(Idx_�Ҳ�).ForeColor = vbRed
        Else
            '�˿�
            txtBalance(Idx_�Ҳ�).Text = Format(Abs(dblMoney - Abs(dbl��ǰδ��)), "0.00")
            '0-ֻ���Ҳ�;1-����Ԥ��
            Call Load�Ҳ���(1, "Ӧ�˿� ")
            txtBalance(Idx_�Ҳ�).ForeColor = txtBalance(Idx_����˵��).ForeColor
        End If
        Exit Sub
    End If
    
    If dbl��ǰδ�� >= dblMoney Then
        'Ҫ�տ�
        Call Load�Ҳ���(0, "Ӧ�տ� ")
        txtBalance(Idx_�Ҳ�).Text = Format(Abs(dblMoney - Abs(dbl��ǰδ��)), "0.00")
        txtBalance(Idx_�Ҳ�).ForeColor = vbRed
        Exit Sub
    Else
        If objCard.�������� = 2 And objCard.���㷽ʽ Like "*֧Ʊ*" Then
            '0-ֻ���Ҳ�;1-����Ԥ��
            Call Load�Ҳ���(1, "�� ֧ Ʊ")
        Else
            '0-ֻ���Ҳ�;1-����Ԥ��
            Call Load�Ҳ���(1, "Ӧ�˿� ")
        End If
        txtBalance(Idx_�Ҳ�).Text = Format(Abs(dblMoney - Abs(dbl��ǰδ��)), "0.00")
        txtBalance(Idx_�Ҳ�).ForeColor = txtBalance(Idx_����˵��).ForeColor
    End If
End Sub

Private Sub tabFeelist_Click(PreviousTab As Integer)
    If tabFeeList.Tab = 1 Then
        picDetailContain.Visible = True
        picFeeContain.Visible = False
    Else
        picDetailContain.Visible = False
        picFeeContain.Visible = True
    End If
End Sub

Private Sub txtBalance_Change(Index As Integer)
    If mblnNotChange Then Exit Sub
    If mEditType = g_Ed_���ݲ鿴 Then Exit Sub
    
    Select Case Index
    Case Idx_�ɿ�
        If Not IDKind�Ҳ�.GetCurCard Is Nothing Then
           If IDKind�Ҳ�.GetCurCard.�ӿ���� <> 1 Then
               txtBalance(Idx_�Ҳ�).Text = "0.00"
               LoadCurOwnerPayInfor
           End If
        End If
        Call Set�Ҳ���Ϣ
        Call SetNextBalanceCmdVisible
        
        If mty_ModulePara.bytˢ��ȱʡ������ = 2 Then
            Dim objCard As Card
            Set objCard = IDKindPaymentsType.GetCurCard
            txtBalance(Idx_�ɿ�).Locked = objCard.�ӿ���� > 0 And Not objCard.���ѿ�
        End If
    Case Idx_��Ԥ��
        If mEditType = g_Ed_�������� Or mEditType = g_Ed_ȡ������ Or chkCancel.Value = 1 Or mblnManualEdit Then Exit Sub
        
        mBalanceInfor.blnԤ����֤ = False
        mBalanceInfor.blnԤ��ˢ�� = False
        If mBalanceInfor.dbl��Ԥ���ϼ� <> 0 Then
            mBalanceInfor.dbl��Ԥ���ϼ� = 0
            If mEditType <> g_Ed_�������� Then Call RecalcDepositMoney(0)
            Call LoadCurOwnerPayInfor
        End If
        
        mbln�ѱ��� = False
        txtBalance(Idx_��Ԥ��).BackColor = IIf(txtBalance(Idx_��Ԥ��).Enabled, &H80000005, &H8000000F)
    Case Idx_���ν���
        mbln�ѱ��� = False
    Case Idx_ժҪ, Idx_����˵��
    Case Else
    End Select
End Sub


Private Sub txtBalance_GotFocus(Index As Integer)
    Select Case Index
    Case Idx_�ɿ�
        Debug.Print "�ɿ�" & Time
        Call LedVoiceSpeak(True)
'        txtBalance(Index).Text = ""
    Case Idx_��Ԥ��
        Debug.Print "Ԥ��" & Time
    Case Idx_ժҪ, Idx_����˵��
        zlCommFun.OpenIme True
    End Select
    zlControl.TxtSelAll txtBalance(Index)
End Sub

Private Sub LedVoiceSpeak(ByVal blnGotFocus As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������
    '���: blnGotFocus-�Ƿ����ɿ�ؼ�,True�ǽ���ʱ,False-�뿪ʱ
    '����:���˺�
    '����:2015-01-28 14:10:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim curTotal As Currency, dblʣ�� As Double
    Dim intSign As Integer
    Dim blnSign As Boolean
    Dim intMark As Integer
    Dim dbl�Ҳ� As Double

    If Not gblnLED Then Exit Sub
    '#21 1234.56   --��������һǧ������ʮ�ĵ�����Ԫ  J
    '#22 1234.56   --Ԥ��һǧ������ʮ�ĵ�����Ԫ Y
    '#23 1234.56   --����һǧ������ʮ�ĵ�����Ԫ Z
    If mrsInfo Is Nothing Then Exit Sub
    If mrsInfo.State <> 1 Then Exit Sub
    
    intSign = IIf(mEditType = g_Ed_�������� Or mEditType = g_Ed_��������, -1, 1)
    curTotal = Val(lblʣ���Ը�.Caption)
    With mBalanceInfor
        dblʣ�� = RoundEx(intSign * (.dbl��ǰ���� + mPatiInfor.dblδ���ۼ� - .dbl��Ԥ���ϼ� - .dblҽ��֧���ϼ�), 5)
    End With
    
    If blnGotFocus Then
        If mbln�ѱ��� Then Exit Sub
        zl9LedVoice.DisplayBank (" ")
        If dblʣ�� >= 0 Then
            zl9LedVoice.Speak "#21 " & curTotal
        Else
            zl9LedVoice.Speak "#23 " & Abs(curTotal)
        End If
        mbln�ѱ��� = True
        Exit Sub
    End If
    intMark = IIf(dblʣ�� >= 0, 1, -1)
    
    dbl�Ҳ� = Val(txtBalance(Idx_�Ҳ�).Text)
    If intMark = 1 Then
        dbl�Ҳ� = Val(IIf(Replace(IDKind�Ҳ�.GetCurCard.����, " ", "") = "Ӧ�˿�", Val(txtBalance(Idx_�Ҳ�).Text), 0))
        zl9LedVoice.DispCharge Format(curTotal, "0.00"), Val(txtBalance(Idx_�ɿ�).Text), dbl�Ҳ�
        zl9LedVoice.Speak "#22 " & Val(txtBalance(Idx_�ɿ�).Text)
        zl9LedVoice.Speak "#23 " & dbl�Ҳ�
        zl9LedVoice.Speak "#3"   '#3  --�뵱�����, лл!
    Else
        If Val(txtBalance(Idx_�ɿ�).Text) > Val(curTotal) Then
            zl9LedVoice.DispCharge Format(intMark * curTotal, "0.00"), Val(txtBalance(Idx_�Ҳ�).Text), Val(txtBalance(Idx_�ɿ�).Text)
            zl9LedVoice.Speak "#22 " & dbl�Ҳ�
            zl9LedVoice.Speak "#23 " & Val(txtBalance(Idx_�ɿ�).Text)
            zl9LedVoice.Speak "#3"   '#3  --�뵱�����, лл!
        Else
        
            zl9LedVoice.DispCharge Format(intMark * curTotal, "0.00"), 0, Val(txtBalance(Idx_�ɿ�).Text) + dbl�Ҳ�
            zl9LedVoice.Speak "#22 " & 0 'intMark * Val(txtBalance(Idx_�ɿ�).Text)
            zl9LedVoice.Speak "#23 " & Val(txtBalance(Idx_�ɿ�).Text) + dbl�Ҳ�
            zl9LedVoice.Speak "#3"   '#3  --�뵱�����, лл!
        End If
    End If
End Sub

Private Sub MoveIDKindItem(ByVal objKind As IDKindNew, ByVal KeyAscii As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ƶ�IDKind��Ŀ
    '���:objKind-�ƶ���IDKind����
    '     Keyascii-��ֵ
    '����:���˺�
    '����:2015-01-29 15:22:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If objKind Is Nothing Then Exit Sub
    If Not (KeyAscii = Asc("+") Or KeyAscii = Asc("-")) Then Exit Sub
    If objKind.ListCount = 1 Then Exit Sub
    
    If KeyAscii = Asc("+") Then
        '����һ��
        If objKind.IDKIND + 1 > objKind.ListCount Then
            objKind.IDKIND = 1
        Else
            objKind.IDKIND = objKind.IDKIND + 1
        End If
        Exit Sub
    End If
    If KeyAscii = Asc("-") Then '����һ��
        If objKind.IDKIND - 1 <= 0 Then
            objKind.IDKIND = objKind.ListCount
        Else
            objKind.IDKIND = objKind.IDKIND - 1
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub txtBalance_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim dblMoney As Double, blnChargeEnd As Boolean
    Dim objCard As Card, objKind As IDKindNew
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim blnNoRecal As Boolean
    
    If KeyAscii <> 13 Then
        If mPatiInfor.dblδ���ۼ� <> 0 Then Exit Sub
        If Index = Idx_�ɿ� Then
             Set objKind = IDKindPaymentsType
        ElseIf Idx_�Ҳ� = Index Then
             Set objKind = IDKind�Ҳ�
        Else
             Exit Sub
        End If
        Call MoveIDKindItem(objKind, KeyAscii)
        Exit Sub
    End If
    KeyAscii = 0
    Select Case Index
    Case Idx_�ɿ�
        dblMoney = RoundEx(Val(txtBalance(Index).Text), 6)
        Set objCard = IDKindPaymentsType.GetCurCard
        If objCard Is Nothing Then zlCommFun.PressKey vbKeyTab: Exit Sub
        Select Case mty_ModulePara.byt�ɿ��������
        Case 2   '�����˽ɿ��ۼ�
            If objCard.�������� = 1 Then '�ֽ�
                If txtBalance(Index).Text = "" And cmdNext.Visible And cmdNext.Enabled Then
                    cmdNext.SetFocus: Exit Sub
                End If
                If SaveBalanceData = False Then Exit Sub
            ElseIf objCard.�������� = 2 Then '��ҽ������
                If txtBalance(Index).Text = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
                If SaveBalanceData = False Then Exit Sub
            ElseIf objCard.�ӿ���� > 0 Then
                If dblMoney = 0 And Val(lblʣ���Ը�.Caption) <> 0 Then
                    MsgBox "δ����ɿ���,����ʹ�á�" & objCard.���㷽ʽ & "�����н���", vbInformation, gstrSysName
                    Exit Sub
                End If
                If SaveBalanceData = False Then Exit Sub
            End If
            Exit Sub
        Case Else '0-�����нɿ����,1-�����ֽ�ʱ,��������ɿ���
            If mty_ModulePara.byt�ɿ�������� = 1 And objCard.�������� = 1 And (Val(lblʣ���Ը�.Caption) <> 0 And Val(txtBalance(4).Text) = 0) Then
                MsgBox "δ����ɿ���,����ʹ�á�" & objCard.���㷽ʽ & "�����н���", vbInformation, gstrSysName
                Exit Sub
            End If
            If objCard.�������� = 1 Then Call cmdOK_Click: Exit Sub
        End Select
        
        If objCard.�ӿ���� > 0 Then
            If dblMoney = 0 And RoundEx(mBalanceInfor.dblδ���ϼ� - mBalanceInfor.dbl��Ԥ���ϼ�, 5) <> 0 Then
                MsgBox "δ����ɿ���,����ʹ�á�" & objCard.���㷽ʽ & "�����н���", vbInformation, gstrSysName
                Exit Sub
            End If
            If SaveBalanceData = False Then Exit Sub
            Exit Sub
        End If
        If dblMoney <> 0 Then
            zlCommFun.PressKey vbKeyTab: Exit Sub
        End If
    Case Idx_��Ԥ��
        If chkDeposit.Visible Then Exit Sub
        If mEditType = g_Ed_�������� Or mEditType = g_Ed_�������� _
        Or chkCancel.Value = 1 And chkCancel.Visible Then Exit Sub
        
        If DepositMonyVerfy = False Then Exit Sub
        
        dblMoney = RoundEx(Val(txtBalance(Idx_��Ԥ��).Text), 6)
        If dblMoney = 0 Then
            zlCommFun.PressKey vbKeyTab: Exit Sub
        End If
        
        Call SaveDeposit(True, blnNoRecal)
        If (mEditType = g_Ed_������� Or mblnCurMzBalanceNo) Then
            If zlDatabase.PatiIdentify(Me, glngSys, mPatiInfor.lng����ID, dblMoney, _
                    , , , IIf(-1 * gdblԤ��������鿨 >= dblMoney, False, True), , , (gdblԤ��������鿨 <> 0), (gdblԤ��������鿨 = 2)) Then
                mBalanceInfor.blnԤ��ˢ�� = True
            End If
        End If
    Case Else
        zlCommFun.PressKey vbKeyTab
    End Select
End Sub

Private Sub txtBalance_LostFocus(Index As Integer)
    Select Case Index
    Case Idx_�ɿ�
        If txtBalance(Index).Text <> "" Then
            txtBalance(Index).Text = Format(Val(txtBalance(Index).Text), "0.00")
        End If
    Case Idx_��Ԥ��
    Case Idx_ժҪ, Idx_����˵��
        zlCommFun.OpenIme False
    End Select
End Sub

Private Sub txtBalance_Validate(Index As Integer, Cancel As Boolean)
    Dim dblMoney As Double, dbl�Ҳ� As Double
    Dim intSign As Integer, strSQL As String, rsTmp As ADODB.Recordset
    Dim blnNoRecal As Boolean
    Select Case Index
    Case Idx_�ɿ�
        '
    Case Idx_��Ԥ��
        If DepositMonyVerfy(False) = False Then Cancel = True: Exit Sub
    Case Idx_���ν���
        If chkCancel.Value = 1 Then Exit Sub
        If Not (mEditType = g_Ed_������� Or mEditType = g_Ed_סԺ����) Then Exit Sub
        
        If RoundEx(Val(txtBalance(Idx_���ν���).Text), 6) = 0 Then
            txtBalance(Idx_���ν���).Text = Format(mBalanceInfor.dbl����δ��, gstrDec)
        Else
            txtBalance(Idx_���ν���).Text = Format(Val(txtBalance(Idx_���ν���).Text), gstrDec)
        End If
         
        If Val(txtBalance(Idx_���ν���).Text) > Val(txtBalance(Idx_����δ��).Text) Then
            MsgBox "��ǰ���ʽ������˱��ν��ʵ��ܶ�,���������!", vbInformation + vbOKOnly, gstrSysName
            zlControl.TxtSelAll txtBalance(Index)
            Cancel = True: Exit Sub
        End If
        
        If RoundEx(Val(txtBalance(Idx_���ν���).Text), 6) < 0 Then
            MsgBox "��ǰ���ʽ��С����,���������!", vbInformation + vbOKOnly, gstrSysName
            zlControl.TxtSelAll txtBalance(Index)
            Cancel = True: Exit Sub
        End If
        
        If mblnNotClick Then Exit Sub
        mblnNotClick = True
        
        Call RelocateMoney
        mBalanceInfor.dbl��ǰ���� = Val(txtBalance(Idx_���ν���).Text)
        mBalanceInfor.dblδ���ϼ� = RoundEx(mBalanceInfor.dbl��ǰ���� - mBalanceInfor.dbl�Ѹ��ϼ�, 5)
        
        Call LoadIntendBalance
        Call LoadCurOwnerPayInfor
        mblnNotClick = False
    Case Else
    End Select
 
End Sub

Private Sub RelocateMoney()
    '������
    Dim dblMoney As Double, i As Long
    Dim blnAll As Boolean
    
    
    dblMoney = Val(txtBalance(Idx_���ν���).Text)
    blnAll = Val(txtBalance(Idx_���ν���).Text) = Val(txtBalance(Idx_����δ��).Text)
    
    With vsDetailList
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("����")) <> "" Then
                If dblMoney >= Val(.Cell(flexcpData, i, .ColIndex("δ����"))) And dblMoney <> 0 Or blnAll Then
                    .Cell(flexcpData, i, .ColIndex("���ʽ��")) = Val(.Cell(flexcpData, i, .ColIndex("δ����")))
                    dblMoney = dblMoney - Val(.Cell(flexcpData, i, .ColIndex("���ʽ��")))
                Else
                    If dblMoney = 0 Then
                        .Cell(flexcpData, i, .ColIndex("���ʽ��")) = ""
                    Else
                        .Cell(flexcpData, i, .ColIndex("���ʽ��")) = dblMoney
                    End If
                    dblMoney = 0
                End If
                .TextMatrix(i, .ColIndex("���ʽ��")) = Format(Val(.Cell(flexcpData, i, .ColIndex("���ʽ��"))), gstrDec)
            End If
        Next i
    End With
End Sub

Private Sub txtBegin_Change()
    If mblnNotChange Then Exit Sub
    If mblnConsChange Then Exit Sub
    
    mblnConsChange = True
    Call ClearListData
End Sub

Private Sub txtBegin_GotFocus()
    zlControl.TxtSelAll txtBegin
End Sub

Private Sub txtBegin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

 
Private Sub txtBegin_Validate(Cancel As Boolean)
    If txtBegin.Text <> "____-__-__" And Not IsDate(txtBegin.Text) Then
        MsgBox "��������ȷ�����ڣ�", vbInformation, gstrSysName
        Cancel = True
    Else
        If txtEnd.Text < txtBegin.Text Then
            MsgBox "��ʼʱ�䲻�ܴ��ڽ���ʱ�䣡", vbInformation, gstrSysName
            Cancel = True
        End If
    End If
End Sub

Private Sub txtEnd_Change()
    If mblnNotChange Then Exit Sub
    If mblnConsChange Then Exit Sub
    mblnConsChange = True
    Call ClearListData
End Sub

Private Sub txtEnd_GotFocus()
    zlControl.TxtSelAll txtEnd
End Sub

Private Sub txtEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtEnd_Validate(Cancel As Boolean)
    If txtEnd.Text <> "____-__-__" And Not IsDate(txtEnd.Text) Then
        MsgBox "��������ȷ�����ڣ�", vbInformation, gstrSysName
        Cancel = True
    Else
        If txtEnd.Text < txtBegin.Text Then
            MsgBox "����ʱ�䲻��С�ڿ�ʼʱ�䣡", vbInformation, gstrSysName
            Cancel = True
        End If
    End If
End Sub

Private Sub txtPatiBegin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
'    txtPatiBegin.Text = mstrPatiBegin
End Sub

Private Sub txtPatiBegin_KeyPress(KeyAscii As Integer)
'    If KeyAscii <> vbKeyReturn Then KeyAscii = 0
End Sub

Private Sub txtPatiEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
'    txtPatiEnd.Text = mstrPatiEnd
End Sub

Private Sub txtPatiEnd_KeyPress(KeyAscii As Integer)
'    If KeyAscii <> vbKeyReturn Then KeyAscii = 0
End Sub

Private Sub vsBlance_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim dblCurMoney As Double, dbl�����ʻ� As Double, dblҽ������ As Double
    With vsBlance
        Select Case Col
        Case .ColIndex("���㷽ʽ")
        Case .ColIndex("������")
            If .RowData(Row) = 110 Then mbln�ѱ��� = False: Exit Sub
            mBalanceInfor.dblҽ��֧���ϼ� = RoundEx(GetYBTotal(0, dbl�����ʻ�, dblҽ������), 5)
            mBalanceInfor.dbl�Ѹ��ϼ� = mBalanceInfor.dblҽ��֧���ϼ�
            mBalanceInfor.dblδ���ϼ� = RoundEx(mBalanceInfor.dbl��ǰ���� - mBalanceInfor.dbl�Ѹ��ϼ�, 5)
            Call LoadCurOwnerPayInfor
            mbln�ѱ��� = False
        Case Else
        End Select
    End With
End Sub
Private Sub vsBlance_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModul, vsBlance, Me.Name, "�����б�"
End Sub

Private Sub vsBlance_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow = 0 Then Exit Sub
    If OldRow = 0 Then Exit Sub
    zl_VsGridRowChange vsBlance, OldRow, NewRow, OldCol, NewCol
End Sub

Private Sub vsBlance_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModul, vsBlance, Me.Name, "�����б�"
End Sub

Private Sub vsBlance_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If mblnBatchState Then Cancel = True: Exit Sub
    If mEditType = g_Ed_���ݲ鿴 Then Cancel = True: Exit Sub
    With vsBlance
        If Val(.TextMatrix(Row, .ColIndex("����״̬"))) = 1 Then
            Cancel = True: Exit Sub
        End If
        .ComboList = ""
        '�༭״̬:0-��ֹɾ��;1-����༭���;2-����ɾ��
        Select Case Val(.TextMatrix(Row, .ColIndex("�༭״̬")))
        Case 0: Cancel = True: Exit Sub
        Case 1
            If Col <> .ColIndex("������") And Col <> .ColIndex("�������") Then Cancel = True: Exit Sub
        Case 2
            If Col = .ColIndex("���㷽ʽ") And .RowData(Row) <> 110 Then
                 .ComboList = "..."
                 .CellButtonPicture = imgDel
            Else
                If .RowData(Row) = 110 Then
                    If Col <> .ColIndex("������") And Col <> .ColIndex("�������") Then Cancel = True: Exit Sub
                Else
                    Cancel = True: Exit Sub
                End If
            End If
        End Select
    End With
End Sub

Private Sub DeletePayInfor(ByVal lngDelRow As Long, Optional ByVal blnForceDel As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ɾ��֧����Ϣ
    '����:���˺�
    '����:2015-01-28 15:18:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, lngRow As Long
    Dim str����Ա���� As String, strDBUser As String
    Dim strPrivs As String, i As Long
    Dim lng�����ID As Long, str���� As String, str����˵�� As String, str������ˮ�� As String
    Dim dblCheckMoney As Double, strBalanceIDs As String
    Dim strArray() As String
    
    On Error GoTo errHandle
    With vsBlance
        If Val(.TextMatrix(lngDelRow, .ColIndex("����"))) = 3 Then
            lng�����ID = Val(.TextMatrix(lngDelRow, .ColIndex("�����ID")))
            str���� = .Cell(flexcpData, lngDelRow, .ColIndex("����"))
            str����˵�� = .TextMatrix(lngDelRow, .ColIndex("����˵��"))
            str������ˮ�� = .TextMatrix(lngDelRow, .ColIndex("������ˮ��"))
            dblCheckMoney = -1 * Val(.TextMatrix(lngDelRow, .ColIndex("������")))
            If .TextMatrix(lngDelRow, .ColIndex("�����Ϣ")) = "" Then
                If mBalanceInfor.lng����ID <> 0 Then
                    strBalanceIDs = "2|" & mBalanceInfor.lng����ID
                End If
            Else
                If Val(.Cell(flexcpData, lngDelRow, .ColIndex("�����Ϣ"))) = 1 Then
                    strBalanceIDs = "1|" & .TextMatrix(lngDelRow, .ColIndex("�����Ϣ"))
                Else
                    strArray = Split(.TextMatrix(lngDelRow, .ColIndex("�����Ϣ")), "|")
                    For i = 0 To UBound(strArray)
                        strBalanceIDs = strBalanceIDs & "," & Split(strArray(i), ",")(4)
                    Next i
                    If strBalanceIDs <> "" Then
                        strBalanceIDs = "1|" & Mid(strBalanceIDs, 2)
                    End If
                End If
            End If
            If zlCallReturnCashCheckInterface(Me, mlngModul, lng�����ID, str����, strBalanceIDs, dblCheckMoney, str������ˮ��, str����˵��) = False Then Exit Sub
        End If
    
        If Val(.TextMatrix(lngDelRow, .ColIndex("�Ƿ�����"))) = 0 And Val(.TextMatrix(lngDelRow, .ColIndex("����"))) = 3 Then
            '����֧�����ֵ����
            If InStr(";" & mstrPrivsCard & ";", ";�����˿�ǿ������;") = 0 Then
                If mstrForceNote = "" Then
                    '�Ѿ���֤���ģ�������֤
                    str����Ա���� = zlDatabase.UserIdentifyByUser(Me, "ǿ��������֤", glngSys, 1151, "�����˿�ǿ������")
                    If str����Ա���� = "" Then
                        MsgBox "¼��Ĳ���Ա��֤ʧ�ܻ���¼��Ĳ���Ա���߱�ǿ������Ȩ�ޣ�����ǿ�����֣�", vbInformation, gstrSysName
                        Exit Sub
                    End If
                    
                    mstrForceNote = str����Ա���� & "ǿ������:" & .TextMatrix(lngDelRow, .ColIndex("���������")) & Format(Abs(Val(.TextMatrix(lngDelRow, .ColIndex("������")))), gstrDec) & "Ԫ" & ";"
                Else
                    mstrForceNote = mstrForceNote & .TextMatrix(lngDelRow, .ColIndex("���������")) & Format(Abs(Val(.TextMatrix(lngDelRow, .ColIndex("������")))), gstrDec) & "Ԫ" & ";"
                End If
            Else
                If MsgBox(.TextMatrix(lngDelRow, .ColIndex("���㷽ʽ")) & "��֧������,�Ƿ�ǿ�ƽ������֣�", _
                            vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) <> vbYes Then Exit Sub
                mstrForceNote = mstrForceNote & IIf(mstrForceNote = "", UserInfo.���� & "ǿ������:", ";") & .TextMatrix(lngDelRow, .ColIndex("���������")) & Format(Abs(Val(.TextMatrix(lngDelRow, .ColIndex("������")))), gstrDec) & "Ԫ"
            End If
        End If
    
        dblMoney = Val(.TextMatrix(lngDelRow, .ColIndex("������")))
        '�༭״̬:0-��ֹɾ��;1-����༭���;2-����ɾ��
        If Val(.TextMatrix(lngDelRow, .ColIndex("�༭״̬"))) <> 2 And blnForceDel = False Then Exit Sub
        
        lngRow = lngDelRow
        mBalanceInfor.dblδ���ϼ� = RoundEx(mBalanceInfor.dblδ���ϼ� + dblMoney, 6)
        mBalanceInfor.dbl�Ѹ��ϼ� = RoundEx(mBalanceInfor.dbl�Ѹ��ϼ� - dblMoney, 6)
        
        Call LoadCurOwnerPayInfor
        If .Rows <= 2 Then
            .Clear 1
            .RowData(1) = ""
            .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
        Else
            vsBlance.RemoveItem lngDelRow
        End If
        
        If lngRow <= 1 Then
            lngRow = 1
        ElseIf lngRow >= .Rows - 1 Then
            lngRow = .Rows - 1
        Else
            lngRow = lngDelRow + 1
        End If
        If lngRow > .Rows - 1 Or lngRow <= 1 Then lngRow = 1
        .Row = lngRow
        If .RowIsVisible(.Row) = False Then .ShowCell .Row, .Col
        Call LoadCurOwnerPayInfor
    End With
    mbln�ѱ��� = False
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Sub
Private Sub vsBlance_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
   
    If mEditType = g_Ed_���ݲ鿴 Then Exit Sub
    Call DeletePayInfor(Row)
End Sub
Private Sub vsBlance_DblClick()
    If mEditType = g_Ed_���ݲ鿴 Then Exit Sub
    With vsBlance
        If .Col <> .ColIndex("������") Then Exit Sub
        '�༭״̬:0-��ֹɾ��;1-����༭���;2-����ɾ��
        If Val(.TextMatrix(.Row, .ColIndex("�༭״̬"))) <> 1 Then Exit Sub
        .EditCell
        .EditSelStart = 0
        .EditSelLength = zlCommFun.ActualLen(.EditText)
    End With
End Sub
 

Private Sub vsBlance_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyDelete Then Exit Sub
    With vsBlance
        '�༭״̬:0-��ֹɾ��;1-����༭���;2-����ɾ��
        If Val(.TextMatrix(.Row, .ColIndex("�༭״̬"))) = 2 Then
            Call DeletePayInfor(.Row)
        End If
    End With

End Sub

Private Sub vsBlance_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
        Exit Sub
    End If
    If Not (mEditType = g_Ed_������� Or mEditType = g_Ed_סԺ����) Then KeyAscii = 0: Exit Sub
    
    With vsBlance
        '�༭״̬:0-��ֹɾ��;1-����༭���;2-����ɾ��
        If Val(.TextMatrix(.Row, .ColIndex("�༭״̬"))) <> 1 Then KeyAscii = 0: Exit Sub
        If .Col <> .ColIndex("������") Then KeyAscii = 0: Exit Sub
    End With
    Call VsFlxGridCheckKeyPress(vsBlance, vsBlance.Row, vsBlance.Col, KeyAscii, m���ʽ)
End Sub

Private Sub vsBlance_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Not (mEditType = g_Ed_������� Or mEditType = g_Ed_סԺ����) Then Exit Sub
    With vsBlance
        If .Col <> .ColIndex("������") Then Exit Sub
    End With
    Call VsFlxGridCheckKeyPress(vsBlance, Row, Col, KeyAscii, m���ʽ)
End Sub

Private Sub vsBlance_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    With vsBlance
        .ToolTipText = ""
        If .MouseRow < 1 Or .MouseRow > .Rows - 1 Then Exit Sub
        If .MouseCol < 0 Or .MouseCol > .Cols - 1 Then Exit Sub
        .ToolTipText = Trim(.TextMatrix(.MouseRow, .MouseCol))
    End With
End Sub

Private Sub vsBlance_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strKey As String
    Dim dblԭʼ��� As Currency
    Dim dblTotal As Double, arrValue As Variant
    Dim i As Long, str���㷽ʽ As String
    Dim varData As Variant
    Dim dblMoney As Double, blnYB As Boolean
      
    With vsBlance
        If .Col <> .ColIndex("������") Then Exit Sub
    End With
    
    
    With vsBlance
        If Row <= 0 Then Exit Sub
        strKey = Trim(.EditText): strKey = Replace(strKey, Chr(vbKeyReturn), ""): strKey = Replace(strKey, Chr(10), "")
        If Not IsNumeric(strKey) And strKey <> "" Then
            MsgBox "����Ľ�����Ϊ���֣�", vbInformation, gstrSysName
            .EditCell
            .EditSelStart = 0
            .EditSelLength = zlCommFun.ActualLen(.EditText)
            Cancel = True
            Exit Sub
        End If
        
        If zlDblIsValid(strKey, 16, False, False, 0, .ColKey(Col)) = False Then
            Cancel = True: Exit Sub
        End If
        
        If vsBlance.RowData(Row) = 110 Then
            dblԭʼ��� = Val(.TextMatrix(Row, Col))
            strKey = Format(Val(strKey), "0.00")
            .EditText = strKey
            dblMoney = RoundEx(Val(strKey) - dblԭʼ���, 6)
            mBalanceInfor.dbl�Ѹ��ϼ� = RoundEx(mBalanceInfor.dbl�Ѹ��ϼ� + dblMoney, 6)
            mBalanceInfor.dblδ���ϼ� = RoundEx(mBalanceInfor.dblδ���ϼ� - dblMoney, 6)
'            Call LoadIntendBalance
            Call LoadDefaultMoney
            Call LoadCurOwnerPayInfor
'            Call Set�Ҳ���Ϣ
            .TextMatrix(Row, Col) = strKey
            Call SetNextBalanceCmdVisible
        Else
            If Val(strKey) < 0 Then
                MsgBox "����Ľ�����Ϊ���ڵ����㣡", vbInformation, gstrSysName
                .EditCell
                .EditSelStart = 0
                .EditSelLength = zlCommFun.ActualLen(.EditText)
                Cancel = True
                Exit Sub
            End If
            
            str���㷽ʽ = Trim(.TextMatrix(.Row, .ColIndex("���㷽ʽ")))
            If str���㷽ʽ = "" Then Exit Sub
            
            '��������������ص�ԭʼ���(�����ʻ�����͸֧ʱ���ж�)
            dblԭʼ��� = Val(.Cell(flexcpData, .Row, .ColIndex("������")))
            Select Case Val(.TextMatrix(.Row, .ColIndex("��������")))
            Case 3 '�����ʻ�
                If Val(strKey) > dblԭʼ��� And Val(strKey) <> 0 And dblԭʼ��� <> 0 Then
                    MsgBox "�����""" & str���㷽ʽ & """������ܳ��� " & Format(dblԭʼ���, "0.00") & " ��", vbInformation, gstrSysName
                    .EditCell
                    .EditSelStart = 0
                    .EditSelLength = zlCommFun.ActualLen(.EditText)
                    Cancel = True
                    Exit Sub
                End If
                 '������������͸֧���
                If mYBInFor.cur������� + mYBInFor.cur����͸֧ - Val(strKey) < 0 Then
                    MsgBox "�ʻ����:" & Format(mYBInFor.cur�������, "0.00") & _
                        IIf(mYBInFor.cur����͸֧ = 0, "", "(" & "����͸֧:" & Format(mYBInFor.cur����͸֧, "0.00") & ")") & _
                        "����Ҫ����Ľ�", vbInformation, gstrSysName
                    .EditCell
                    .EditSelStart = 0
                    .EditSelLength = zlCommFun.ActualLen(.EditText)
                    Cancel = True
                    Exit Sub
                End If
                blnYB = True
            Case 4 'ҽ������
                If Val(strKey) > dblԭʼ��� And Val(strKey) <> 0 And dblԭʼ��� <> 0 Then
                    MsgBox "�����""" & str���㷽ʽ & """������ܳ��� " & Format(dblԭʼ���, "0.00") & " ��", vbInformation, gstrSysName
                    .EditCell
                    .EditSelStart = 0
                    .EditSelLength = zlCommFun.ActualLen(.EditText)
                    Cancel = True
                    Exit Sub
                End If
                blnYB = True
            End Select
            
            '������������ʣ��ɽ�����
            dblTotal = RoundEx(mBalanceInfor.dbl��ǰ���� - GetYBTotal(.Row) - mBalanceInfor.dbl�Ѹ��ϼ�, 5)
            If Val(strKey) > dblTotal And RoundEx(Val(strKey), 6) <> 0 And blnYB = False Then
                MsgBox "��������󣬳����������������:" & Format(dblTotal, "0.00") & "��", vbInformation, gstrSysName
                .EditCell
                .EditSelStart = 0
                .EditSelLength = zlCommFun.ActualLen(.EditText)
                Cancel = True
                Exit Sub
            End If
            strKey = Format(Val(strKey), "0.00")
            .EditText = strKey
            .TextMatrix(Row, Col) = strKey
            '���¼���ҽ��������
            Call ReCalcYBMoney
        End If
    End With
End Sub

Private Sub ReCalcYBMoney()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���¼���ҽ�����
    '����:���˺�
    '����:2015-01-21 15:41:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, i As Long
    Dim dbl�����ʻ� As Double, dblҽ������ As Double, dblMoney As Double
    Dim str���㷽ʽ As String
    
    On Error GoTo errHandle
    With vsBlance
        For i = 1 To .Rows - 1
            str���㷽ʽ = Trim(.TextMatrix(i, .ColIndex("���㷽ʽ")))
            If str���㷽ʽ <> "" Then
                 varData = Split(.RowData(i) & "|||", "|")
                '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
                 dblMoney = Val(.TextMatrix(i, .ColIndex("������")))
                 Select Case Val(.TextMatrix(i, .ColIndex("��������")))
                 Case 3 '�����ʻ�
                    dbl�����ʻ� = dbl�����ʻ� + dblMoney
                 Case 4 'ҽ������
                    dblҽ������ = dblҽ������ + dblMoney
                 End Select
            End If
        Next
    End With
        
    mBalanceInfor.dblҽ��֧���ϼ� = RoundEx(dbl�����ʻ� + dblҽ������, 5)
    mYBInFor.cur����֧�� = dbl�����ʻ�
    mYBInFor.curͳ��֧�� = dblҽ������
    
'    lbl�����ʻ�.Caption = "�ʻ����:" & Format(mYBInFor.cur�������, "0.00")
    staThis.Panels(5).Text = Format(mYBInFor.cur�������, "0.00")
    staThis.Panels(5).Visible = True
    lblҽ������.Caption = "ͳ��֧��:" & Format(dblҽ������, "0.00")
    
    lblҽ������.Visible = True
'    lblCurPaymentInfor.Visible = True
    txtBalance(Idx_���ν���).Enabled = False
    
    Call ShowYBMoneyOrInvioceFormatInfor
    
    'bytFun-0-����ǰ;1-ҽ����������;2-�ѱ����˽��ʵ�;
    Call SetOperationCtrl(1)
    '��ʾҽ��������Ϣ:bytFun-0-ҽ��Ԥ����Ϣ��ʾ
    Call ShowLedDisplayBank(0)
    Call LoadCurOwnerPayInfor  '����֧���ϼ�
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    
End Sub


Private Function GetYBTotal(ByVal lngRow As Long, _
    Optional ByRef dbl�����ʻ� As Double, _
    Optional ByRef dblҽ������ As Double) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡҽ��֧���ܶ�
    '���:lngRow-����������
    '����:ҽ��֧���ܶ�
    '����:���˺�
    '����:2015-01-21 15:41:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, dblMoney As Double, str���㷽ʽ As String
    
    On Error GoTo errHandle
    With vsBlance
        For i = 1 To .Rows - 1
            str���㷽ʽ = Trim(.TextMatrix(.Row, .ColIndex("���㷽ʽ")))
            If str���㷽ʽ <> "" And i <> lngRow Then
                '��������:���㷽ʽ.����
                 dblMoney = Val(.TextMatrix(i, .ColIndex("������")))
                 Select Case Val(.TextMatrix(i, .ColIndex("��������")))
                 Case 3 '�����ʻ�
                    dbl�����ʻ� = dbl�����ʻ� + dblMoney
                 Case 4 'ҽ������
                    dblҽ������ = dblҽ������ + dblMoney
                 End Select
            End If
        Next
    End With
    
    GetYBTotal = dbl�����ʻ� + dblҽ������
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    
End Function

Private Sub vsDeposit_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim dblԤ����� As Double, dbl��Ԥ�� As Double
    Dim i As Long
    Dim dblMoney As Double
    
    If mblnNoTrigger Then
        mblnNoTrigger = False
        Exit Sub
    End If
    
    With vsDeposit
        If IsNumeric(.TextMatrix(Row, .ColIndex("��Ԥ��"))) = False And .TextMatrix(Row, .ColIndex("��Ԥ��")) <> "" Then
            MsgBox "��������ȷ�ĳ�Ԥ�����!", vbInformation, gstrSysName
            .TextMatrix(Row, .ColIndex("��Ԥ��")) = ""
            If Val(.TextMatrix(Row, .ColIndex("��Ԥ��"))) = 0 Then .Cell(flexcpForeColor, Row, 0, Row, .Cols - 1) = vbBlack
            Exit Sub
        End If
        If Val(.TextMatrix(Row, .ColIndex("���"))) < Val(.TextMatrix(Row, .ColIndex("��Ԥ��"))) Then
            MsgBox "����ĳ�Ԥ��������,����������!", vbInformation, gstrSysName
            .TextMatrix(Row, .ColIndex("��Ԥ��")) = ""
            If Val(.TextMatrix(Row, .ColIndex("��Ԥ��"))) = 0 Then .Cell(flexcpForeColor, Row, 0, Row, .Cols - 1) = vbBlack
            Exit Sub
        End If
        If Val(.TextMatrix(Row, .ColIndex("���"))) < 0 And Val(.TextMatrix(Row, .ColIndex("��Ԥ��"))) > 0 Then
            MsgBox "��������ȷ�ĳ�Ԥ�����!", vbInformation, gstrSysName
            .TextMatrix(Row, .ColIndex("��Ԥ��")) = ""
            If Val(.TextMatrix(Row, .ColIndex("��Ԥ��"))) = 0 Then .Cell(flexcpForeColor, Row, 0, Row, .Cols - 1) = vbBlack
            Exit Sub
        End If
        If Val(.TextMatrix(Row, .ColIndex("���"))) > 0 And Val(.TextMatrix(Row, .ColIndex("��Ԥ��"))) < 0 Then
            MsgBox "��������ȷ�ĳ�Ԥ�����!", vbInformation, gstrSysName
            .TextMatrix(Row, .ColIndex("��Ԥ��")) = ""
            If Val(.TextMatrix(Row, .ColIndex("��Ԥ��"))) = 0 Then .Cell(flexcpForeColor, Row, 0, Row, .Cols - 1) = vbBlack
            Exit Sub
        End If
        For i = 1 To .Rows - 1
            dblԤ����� = RoundEx(dblԤ����� + Val(.TextMatrix(i, .ColIndex("���"))), 5)
            dbl��Ԥ�� = RoundEx(dbl��Ԥ�� + Val(.TextMatrix(i, .ColIndex("��Ԥ��"))), 5)
        Next i
        If Val(dblԤ�����) < Val(dbl��Ԥ��) Then
            MsgBox "����ĳ�Ԥ��������,����������!", vbInformation, gstrSysName
            .TextMatrix(Row, .ColIndex("��Ԥ��")) = ""
            If Val(.TextMatrix(Row, .ColIndex("��Ԥ��"))) = 0 Then .Cell(flexcpForeColor, Row, 0, Row, .Cols - 1) = vbBlack
            Exit Sub
        End If
        .TextMatrix(Row, .ColIndex("��Ԥ��")) = Format(Val(.TextMatrix(Row, .ColIndex("��Ԥ��"))), "0.00")
        mblnManualEdit = True
        txtBalance(Idx_��Ԥ��).Text = Format(dbl��Ԥ��, "0.00")
        mBalanceInfor.dbl��Ԥ���ϼ� = dbl��Ԥ��
        
        If chkDeposit.Visible Then Exit Sub
        dblMoney = RoundEx(Val(txtBalance(Idx_��Ԥ��).Text), 6)
        
        If mblnNotChange = False Then
            If Val(dblMoney) > Val(mPatiInfor.dblʵ�����) Then
                MsgBox "��ǰ����ĳ�Ԥ������Ԥ�����,���ܼ���!" & vbCrLf & "ʵ�����:" & Format(mPatiInfor.dblʵ�����, "0.00") & vbCrLf & "��Ԥ��:" & Format(Val(txtBalance(Idx_��Ԥ��).Text), "0.00")
                .TextMatrix(Row, .ColIndex("��Ԥ��")) = ""
                If Val(.TextMatrix(Row, .ColIndex("��Ԥ��"))) = 0 Then .Cell(flexcpForeColor, Row, 0, Row, .Cols - 1) = vbBlack
                mblnManualEdit = False
                Exit Sub
            End If
        End If
        
        If Val(.TextMatrix(Row, .ColIndex("��Ԥ��"))) = 0 Then .Cell(flexcpForeColor, Row, 0, Row, .Cols - 1) = vbBlack
        
        If Not mBalanceInfor.blnԤ����֤ Then
            If CheckDepositValied(True) = False Then mblnManualEdit = False: Exit Sub
        End If
        Call LoadIntendBalance
        Call LoadCurOwnerPayInfor
        Call Set�Ҳ���Ϣ
        mbln�ѱ��� = False
        mblnManualEdit = False
    End With
End Sub

Private Sub vsDeposit_AfterMoveColumn(ByVal Col As Long, Position As Long)
     zl_vsGrid_Para_Save mlngModul, vsDeposit, Me.Name, "Ԥ���б�"
End Sub

Private Sub vsDeposit_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    zl_VsGridRowChange vsDeposit, OldRow, NewRow, OldCol, NewCol
    Call SetUpDown
End Sub

Private Sub vsDeposit_AfterSort(ByVal Col As Long, Order As Integer)
    Dim i As Long
    If mstrNoSort <> "" Then
        With vsDeposit
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("���ݺ�")) = mstrNoSort Then
                    .Select i, Col
                    Exit For
                End If
            Next i
        End With
    End If
    If mEditType = g_Ed_���ݲ鿴 Or mEditType = g_Ed_�������� Or chkCancel.Value = 1 Then Exit Sub
    
    Call RecalcDepositMoney(2, Val(mBalanceInfor.dbl��Ԥ���ϼ�))
    Call LoadIntendBalance
    Call LoadCurOwnerPayInfor
End Sub

Private Sub vsDeposit_BeforeSort(ByVal Col As Long, Order As Integer)
    mstrNoSort = vsDeposit.TextMatrix(vsDeposit.Row, vsDeposit.ColIndex("���ݺ�"))
End Sub

Private Sub vsDeposit_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
     zl_vsGrid_Para_Save mlngModul, vsDeposit, Me.Name, "Ԥ���б�"
End Sub

Private Sub vsDeposit_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If mblnBatchState Then Cancel = True: Exit Sub
    If Not (mEditType = g_Ed_������� Or mEditType = g_Ed_סԺ���� Or mEditType = g_Ed_���½���) Then Cancel = True
    If chkCancel.Value = 1 Then Cancel = True
    If Val(vsDeposit.TextMatrix(Row, vsDeposit.ColIndex("�༭״̬"))) <> 0 Then Cancel = True
    If Col <> vsDeposit.ColIndex("��Ԥ��") Then Cancel = True
End Sub

Private Sub vsDeposit_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsDeposit
        If Col = .ColIndex("���ݺ�") Then Cancel = True: Exit Sub
    End With
End Sub
Private Sub txtPatient_Change()
    If Not Me.ActiveControl Is txtPatient Or txtPatient.Locked Then Exit Sub
    IDKIND.SetAutoReadCard (txtPatient.Text = "")
End Sub

Private Sub txtPatient_GotFocus()
    zlControl.TxtSelAll txtPatient
    If txtPatient.Locked Then Exit Sub
    IDKIND.SetAutoReadCard (txtPatient.Text = "")
End Sub
Private Sub txtPatient_LostFocus()
    IDKIND.SetAutoReadCard (False)
End Sub
Private Sub txtPatient_Validate(Cancel As Boolean)
    If mrsInfo Is Nothing Then Exit Sub
    If mrsInfo.State <> 1 Then Exit Sub
    If mrsInfo.EOF Then Exit Sub
    If txtPatient.Text <> mrsInfo!���� Then txtPatient.Text = mrsInfo!����
End Sub
Private Sub txtPatient_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtPatient.Locked Or txtPatient.Enabled = False Or txtPatient.Text <> "" Then Exit Sub
    If IDKIND.ActiveFastKey = True Then Exit Sub
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim blnCard As Boolean, strInput As String
    
    If txtPatient.Locked Then Exit Sub
    If KeyAscii = 13 Then
        If Not mrsInfo Is Nothing Then
            If mrsInfo.State = 1 Then
                If mrsInfo!���� = txtPatient.Text Then
                    zlCommFun.PressKey vbKeyTab: Exit Sub
                End If
            End If
        End If
    End If
    
    '����ѡ����
    If Len(Trim(Me.txtPatient.Text)) = 0 And KeyAscii = 13 Then
        If mEditType = g_Ed_������� Then
            Call cmdYB_Click
            Exit Sub
        Else
            With frmPatiSelect
                .mstrPrivs = mstrPrivs
                .mbytUseType = 3
                Set .mfrmParent = Me
                .Show 1, Me
                mty_ModulePara.intPatientRange = Val(zlDatabase.GetPara("��ʾ���岡��", glngSys, mlngModul, 0))
            End With
        End If
    Else
        If IDKIND.GetCurCard.���� Like "����*" Then
            blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKIND.ShowPassText)
        ElseIf IDKIND.GetCurCard.���� = "�����" Or IDKIND.GetCurCard.���� = "סԺ��" Then
            If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
                If InStr("0123456789-*+", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
            End If
        Else
            txtPatient.PasswordChar = IIf(IDKIND.ShowPassText, "*", "")
            '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
            txtPatient.IMEMode = 0
        End If
    End If
    'ˢ����ϻ���������س�
    If blnCard And Len(txtPatient.Text) = IDKIND.GetCardNoLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Trim(txtPatient.Text) <> "" Then
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        End If
        KeyAscii = 0
        strInput = txtPatient.Text
        mstrPatient = txtPatient.Text
        Call FindPati(IDKIND.GetCurCard, blnCard, strInput)
    End If
End Sub
Private Sub Led_ClearDisplayPatient()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���Led��ʾ��
    '����:���˺�
    '����:2014-12-31 10:38:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mstrInNO <> "" Or Not gblnLED Then Exit Sub
    If mEditType = g_Ed_���ݲ鿴 Then Exit Sub
    zl9LedVoice.DisplayPatient ""
End Sub


Private Sub HideYBMoneyInfo()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ͳ��֧����Ϣ
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-12-31 11:39:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
'    lblCurPaymentInfor.Visible = True
    lblҽ������.Caption = "ͳ��֧��:"
    lblҽ������.Visible = False
'    lbl�����ʻ�.Caption = "�ʻ����:"
    staThis.Panels(5).Text = ""
    staThis.Panels(5).Visible = False
'    lbl�����ʻ�.Visible = False
End Sub

Private Sub NewBill()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ�����ʽ���
    '����:���˺�
    '����:2014-12-31 10:05:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call ClearCustomType '����Զ���������ر���
    Call SetBatchControl(True)
    Call Led_ClearDisplayPatient '���Led������ʾ
    Set mrsInfo = New ADODB.Recordset '���������Ϣ
    mblnCurMzBalanceNo = False
    mbln�ѱ��� = False
    
    '������ü�Ԥ����Ϣ
    Call InitGrid
    '���������Ϣ
    Call ClearBalance '���������Ϣ
    Call HideYBMoneyInfo    '����ͳ��֧�������
    Call InitBalanceCondition   '��ʼ������������ر���
    Call InitPatiBalanceVariableCon     '������˽�����ر���
     
    Call SetControlEnabled(True) '���ÿؼ������״̬
    
    txtPatient.ForeColor = Me.ForeColor
    
    pic״̬.Visible = False: lbl״̬.Caption = "":  lbl���ʽ.Caption = ""
    txtPatient.Text = "":    txtSex.Text = "":      txtOld.Text = ""
    txt�ѱ�.Text = "":       txt��ʶ��.Text = "":   txtBed.Text = "": txt����.Text = ""
    
    txtBegin.Text = "____-__-__": txtEnd.Text = "____-__-__"
    txtPatiBegin.Text = "____-__-__": txtPatiEnd.Text = "____-__-__":    txtPatiEnd.Tag = "____-__-__"
    txtDate.Text = "____-__-__ __:__:__": txt����.Text = ""
    txtBalance(Idx_����˵��).Text = ""
    lblBed.Visible = False:     txtBed.Visible = False
    lbl��ʶ��.Visible = True:  txt��ʶ��.Visible = True
    lbl����.Visible = False:    txt����.Visible = False
    
    mblnNotify = False
    txtԤ�����.Text = ""
    txtδ�����.Text = ""
    txtʣ����.Text = ""
    mstrBalanceLimit = ""
    mstrForceNote = ""
    mstrCardPara = ""
    
    lblPrevious.Visible = False
    lblPrevious.Caption = ""
    lblCurPaymentInfor.Visible = True
    
    lblTicketCount.Caption = "Ԥ�����վ�:"
    staThis.Panels(2) = ""
    lblBalanceType.Visible = False
    Call SetOperationCtrl(0)
    Call SetFeeListColumnShow
    Call SetPatiConsControlVisible
    Call SetOperatonCommandCaption
    Call SetDefaultPayType
End Sub

Private Sub SetPatiConsControlVisible()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ò��������ؼ�����ʾ
    '����:���˺�
    '����:2014-12-31 14:26:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnMzBalance As Boolean, blnVisible As Boolean
    If mEditType = g_Ed_������� Or mblnCurMzBalanceNo Then
        blnMzBalance = True
    ElseIf mEditType = g_Ed_סԺ���� Then
        blnMzBalance = False
    End If
    lblBed.Visible = Not blnMzBalance
    lbl����.Visible = Not blnMzBalance
    txt����.Visible = Not blnMzBalance
    blnVisible = mEditType = g_Ed_������� And InStr(mstrPrivs, ";���ս���;") > 0
    cmdYB.Visible = blnVisible
    If blnVisible And Not mblnMC_TwoMode And InStr(mstrPrivs, ";������ý���;") = 0 Then
       cmdYB.Visible = False
    End If
    
    lblPatiTime.Visible = Not blnMzBalance
    lblPatiTimeRange.Visible = Not blnMzBalance
    txtPatiBegin.Visible = Not blnMzBalance
    txtPatiEnd.Visible = Not blnMzBalance
    txt����.Visible = Not blnMzBalance
    lblDayName.Visible = Not blnMzBalance
    
    lblPatiNums.Caption = IIf(blnMzBalance, "�������", "סԺ����")
    lblPatiNums.Visible = True
    cboPatiNums.Visible = True
    
    cboӤ��.Visible = Not blnMzBalance And cboӤ��.ListCount > 2
    If cboӤ��.Visible Then
        cboPatiNums.Width = cboӤ��.Left - cboPatiNums.Left - 30
    End If
     
    opt��;.Visible = Not blnMzBalance
    opt��Ժ.Visible = Not blnMzBalance
    
    txtBed.Visible = Not blnMzBalance
    lblBed.Visible = Not blnMzBalance
    chkPatiType(CK_Idx_��ͨ).Visible = blnMzBalance
    chkPatiType(CK_Idx_���).Visible = blnMzBalance
    
    chkCancel.Visible = (mEditType = g_Ed_������� Or mEditType = g_Ed_סԺ����)
    lblDelCaption.Visible = mblnViewCancel Or mEditType = g_Ed_ȡ������ Or mEditType = g_Ed_�������� Or mEditType = g_Ed_��������
    
    Call picNO_Resize
    If (mEditType <> g_Ed_������� And mEditType <> g_Ed_סԺ����) _
        Or chkCancel.Value Or mEditType = g_Ed_���ݲ鿴 Then
        '�ǽ���ʱ����������������
        cboӤ��.Visible = False
        lblDept.Visible = False: cboDept.Visible = False
        lblFeeType.Visible = False: cboFeeType.Visible = False
        lblFeeItem.Visible = False: cboFeeItem.Visible = False
        lblFeeStyle.Visible = False: cboChargeType.Visible = False
        opt��;.Visible = False: opt��Ժ.Visible = False
        chkPatiType(CK_Idx_��ͨ).Visible = False
        chkPatiType(CK_Idx_���).Visible = False
        lblPatiNums.Visible = False
        cboPatiNums.Visible = False
        cmdRefresh.Visible = False
        cmdYB.Visible = False
    Else
        lblDept.Visible = True: cboDept.Visible = True
        lblFeeType.Visible = True: cboFeeType.Visible = True
        lblFeeItem.Visible = True: cboFeeItem.Visible = True
        lblFeeStyle.Visible = True: cboChargeType.Visible = True
        cmdRefresh.Visible = True
    End If
    
    If blnMzBalance Then
        lbl��ʶ��.Caption = "�����"
    End If
    
    Call MovePatiConsControl
End Sub


Private Sub MovePatiConsControl()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����ؼ�λ��
    '����:���˺�
    '����:2014-12-31 15:03:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnMzBalance As Boolean
    Dim lngStep As Long, sngLeft As Single
    Dim objPan As Pane
    
    '1.סԺ����ԭ����
    txtBegin.Width = 1335
    txtEnd.Width = txtBegin.Width
    
    Set objPan = dkpMain.FindPane(Pan_PatiCon)
    If objPan Is Nothing Then Exit Sub
    
    If mEditType = g_Ed_סԺ���� And chkCancel.Value <> 1 Then
        
        objPan.MaxTrackSize.Height = 2150 \ Screen.TwipsPerPixelY
        objPan.MinTrackSize.Height = 2150 \ Screen.TwipsPerPixelY
        dkpMain.RecalcLayout
        
        txtBegin.Top = cboFeeItem.Top + cboFeeItem.Height + 180
        txtEnd.Top = txtBegin.Top
        txtPatiBegin.Top = txtBegin.Top
        txtPatiEnd.Top = txtBegin.Top
        txt����.Top = txtBegin.Top
        opt��Ժ.Top = txtBegin.Top + (txtBegin.Height - opt��Ժ.Height) \ 2
        opt��;.Top = opt��Ժ.Top
        lblFsTime.Top = txtBegin.Top + (txtBegin.Height - lblFsTime.Height) \ 2
        lblFsTimeRange.Top = lblFsTime.Top
        lblPatiTime.Top = lblFsTime.Top
        lblPatiTimeRange.Top = lblFsTime.Top
        lblDayName.Top = lblFsTime.Top
        
        lblFsTime.Left = lblDept.Left
        txtBegin.Left = cboDept.Left
        lblFsTimeRange.Left = txtBegin.Left + txtBegin.Width + 80
        txtEnd.Left = lblFsTimeRange.Left + lblFsTimeRange.Width + 80
        
        lblPatiTime.Left = txtEnd.Left + txtEnd.Width + 210
        txtPatiBegin.Left = lblPatiTime.Left + lblPatiTime.Width + 30
        lblPatiTimeRange.Left = txtPatiBegin.Left + txtPatiBegin.Width + 20
        txtPatiEnd.Left = lblPatiTimeRange.Left + lblPatiTimeRange.Width + 20
        txt����.Left = txtPatiEnd.Left + txtPatiEnd.Width + 20
        lblDayName.Left = txt����.Left + txt����.Width + 20
        Exit Sub
    End If
    
    If mEditType = g_Ed_������� And chkCancel.Value <> 1 Then
        '2.������ʽ���
        objPan.MaxTrackSize.Height = 2150 \ Screen.TwipsPerPixelY
        objPan.MinTrackSize.Height = 2150 \ Screen.TwipsPerPixelY
        dkpMain.RecalcLayout
                
        lngStep = 500
        sngLeft = txtPatient.Left + txtPatient.Width
        cmdYB.Left = sngLeft + 10
        sngLeft = cmdYB.Left + cmdYB.Width
        lblSex.Left = sngLeft + lngStep
        txtSex.Left = lblSex.Left + lblSex.Width + 10
        
        lblOld.Left = txtSex.Left + txtSex.Width + lngStep
        txtOld.Left = lblOld.Left + lblOld.Width + 10
        
        lbl�ѱ�.Left = txtOld.Left + txtOld.Width + lngStep
        txt�ѱ�.Left = lbl�ѱ�.Left + lbl�ѱ�.Width + 10
            
        lbl��ʶ��.Left = txt�ѱ�.Left + txt�ѱ�.Width + lngStep
        txt��ʶ��.Left = lbl��ʶ��.Left + lbl��ʶ��.Width + 10
        
        txtBegin.Top = cboFeeItem.Top + cboFeeItem.Height + 100
        txtEnd.Top = txtBegin.Top
        lblFsTime.Top = txtBegin.Top + (txtBegin.Height - lblFsTime.Height) \ 2
        lblFsTimeRange.Top = lblFsTime.Top
        
        lblFsTime.Left = lblDept.Left
        txtBegin.Left = cboDept.Left
        lblFsTimeRange.Left = txtBegin.Left + txtBegin.Width + 80
        txtEnd.Left = lblFsTimeRange.Left + lblFsTimeRange.Width + 80
        
'        txtBegin.Width = (cboFeeItem.Width - lblFsTimeRange.Width - 20) \ 2
'        txtEnd.Width = txtBegin.Width
'        lblFsTimeRange.Left = txtBegin.Left + txtBegin.Width + 10
'        txtEnd.Left = lblFsTimeRange.Left + lblFsTimeRange.Width + 10

        chkPatiType(CK_Idx_��ͨ).Top = lblFsTime.Top
        chkPatiType(CK_Idx_���).Top = chkPatiType(CK_Idx_��ͨ).Top
        Exit Sub
    End If
    
    '3.��������(����,�ؽ�,���ĵ�)
    
    objPan.MinTrackSize.Height = (txtPatiType.Top + txtPatiType.Height + 100) \ Screen.TwipsPerPixelY
    objPan.MaxTrackSize.Height = objPan.MinTrackSize.Height
    dkpMain.RecalcLayout

    
    lngStep = IIf(mblnCurMzBalanceNo, 500, 300)
    
    sngLeft = txtPatient.Left + txtPatient.Width + lngStep
    lblSex.Left = sngLeft
    txtSex.Left = lblSex.Left + lblSex.Width + 10
    
    lblOld.Left = txtSex.Left + txtSex.Width + lngStep
    txtOld.Left = lblOld.Left + lblOld.Width + 10
    
    lbl�ѱ�.Left = txtOld.Left + txtOld.Width + lngStep
    txt�ѱ�.Left = lbl�ѱ�.Left + lbl�ѱ�.Width + 10
    
    lbl��ʶ��.Left = txt�ѱ�.Left + txt�ѱ�.Width + lngStep
    txt��ʶ��.Left = lbl��ʶ��.Left + lbl��ʶ��.Width + 10
        
    lblBed.Left = txt��ʶ��.Left + txt��ʶ��.Width + lngStep
    txtBed.Left = lblBed.Left + lblBed.Width + 10
        
    lbl����.Left = txtBed.Left + txtBed.Width + lngStep
    txt����.Left = lbl����.Left + lbl����.Width + 10
         
         
    txtBegin.Top = cboӤ��.Top
    txtEnd.Top = txtBegin.Top
    txtPatiBegin.Top = txtBegin.Top
    txtPatiEnd.Top = txtBegin.Top
    txt����.Top = txtBegin.Top
    opt��Ժ.Top = txtBegin.Top + (txtBegin.Height - opt��Ժ.Height) \ 2
    opt��;.Top = opt��Ժ.Top
    lblFsTime.Top = txtBegin.Top + (txtBegin.Height - lblFsTime.Height) \ 2
    lblFsTimeRange.Top = lblFsTime.Top
    lblPatiTime.Top = lblFsTime.Top
    lblPatiTimeRange.Top = lblFsTime.Top
    lblDayName.Top = lblFsTime.Top
    
    
    lblFsTime.Left = txtPatiType.Left + txtPatiType.Width + 100
    txtBegin.Left = lblFsTime.Left + lblFsTime.Width + 10
    
    lblFsTimeRange.Left = txtBegin.Left + txtBegin.Width + 20
    txtEnd.Left = lblFsTimeRange.Left + lblFsTimeRange.Width + 20
    
    lblPatiTime.Left = txtEnd.Left + txtEnd.Width + 400
    txtPatiBegin.Left = lblPatiTime.Left + lblPatiTime.Width + 20
    lblPatiTimeRange.Left = txtPatiBegin.Left + txtPatiBegin.Width + 20
    txtPatiEnd.Left = lblPatiTimeRange.Left + lblPatiTimeRange.Width + 20
    txt����.Left = txtPatiEnd.Left + txtPatiEnd.Width + 20
    lblDayName.Left = txt����.Left + txt����.Width + 20
    lblBalanceType.Top = txt����.Top
    lblBalanceType.Left = picPati.ScaleWidth - lblBalanceType.Width - 50
End Sub

Private Sub SetPatiEnabled(blnEnabled As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ò�����صı༭����
    '����:���˺�
    '����:2015-01-04 16:39:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    chkCancel.Enabled = blnEnabled And Not mPatiInfor.bln��������
    cmdYB.Enabled = blnEnabled
    txtPatient.Locked = Not blnEnabled
    txtPatiBegin.Enabled = blnEnabled
    txtPatiEnd.Enabled = blnEnabled
    txtBalance(Idx_���ν���).Locked = (InStr(mstrPrivs, ";��������;") = 0)
    
    If mEditType = g_Ed_������� Then
        opt��;.Enabled = False
        opt��Ժ.Enabled = False
    Else
        opt��;.Enabled = blnEnabled
        opt��Ժ.Enabled = blnEnabled
    End If
End Sub

Private Sub SetControlEnabled(blnEanbled As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ƽ���״̬
    '���:blnEanbled-�Ƿ���Ч
    '����:���˺�
    '����:2014-12-31 12:01:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim EditType As gBalanceBill
    
    EditType = mEditType
    If chkCancel.Value = 1 And (mEditType = g_Ed_������� Or mEditType = g_Ed_סԺ����) Then
        EditType = g_Ed_��������
    End If
    
    Select Case EditType
    Case g_Ed_�������
        txtPatient.Locked = Not blnEanbled
        chkCancel.Enabled = blnEanbled And Not mPatiInfor.bln��������
        cmdYBBalance.Enabled = blnEanbled
        cmdYB.Enabled = blnEanbled
        txtPatient.Locked = Not blnEanbled
        txtBalance(Idx_���ν���).Locked = (InStr(mstrPrivs, ";��������;") = 0)
        txtBalance(Idx_���ν���).Enabled = Not txtBalance(Idx_���ν���).Locked
        txtBalance(Idx_����˵��).Enabled = blnEanbled
        
        cboDept.Enabled = blnEanbled And InStr(mstrPrivs, ";��������;") > 0
        cboFeeType.Enabled = blnEanbled And InStr(mstrPrivs, ";��������;") > 0
        cboFeeItem.Enabled = blnEanbled And InStr(mstrPrivs, ";��������;") > 0
        cboChargeType.Enabled = blnEanbled And InStr(mstrPrivs, ";��������;") > 0
        txtInvoice.Enabled = blnEanbled
        cboPatiNums.Enabled = blnEanbled And InStr(mstrPrivs, ";��������;") > 0
        cboDiag.Enabled = blnEanbled And InStr(mstrPrivs, ";��������;") > 0
        txtBegin.Enabled = blnEanbled
        txtEnd.Enabled = blnEanbled
        cboӤ��.Enabled = False
        txtPatiBegin.Enabled = False
        txtPatiEnd.Enabled = False
        opt��;.Enabled = False
        opt��Ժ.Enabled = False
    Case g_Ed_סԺ����
        txtPatient.Locked = Not blnEanbled
        chkCancel.Enabled = blnEanbled And Not mPatiInfor.bln��������
        cmdYBBalance.Enabled = blnEanbled
        txtPatient.Locked = Not blnEanbled
        txtPatiBegin.Enabled = blnEanbled
        txtPatiEnd.Enabled = blnEanbled
        
        cboӤ��.Enabled = blnEanbled And InStr(mstrPrivs, ";��������;") > 0
        cboDept.Enabled = blnEanbled And InStr(mstrPrivs, ";��������;") > 0
        cboFeeType.Enabled = blnEanbled And InStr(mstrPrivs, ";��������;") > 0
        cboFeeItem.Enabled = blnEanbled And InStr(mstrPrivs, ";��������;") > 0
        cboChargeType.Enabled = blnEanbled And InStr(mstrPrivs, ";��������;") > 0
        cboDiag.Enabled = blnEanbled And InStr(mstrPrivs, ";��������;") > 0
        
        txtInvoice.Enabled = blnEanbled
        opt��;.Enabled = blnEanbled
        opt��Ժ.Enabled = blnEanbled
        opt��Ժ.Caption = "��Ժ����"
        txtBalance(Idx_���ν���).Locked = (InStr(mstrPrivs, ";��������;") = 0)
        txtBalance(Idx_���ν���).Enabled = Not txtBalance(Idx_���ν���).Locked
        txtBalance(Idx_����˵��).Enabled = blnEanbled
        cboPatiNums.Enabled = blnEanbled And InStr(mstrPrivs, ";��������;") > 0
    Case Else  'g_Ed_ȡ������, g_Ed_���ݲ鿴, g_Ed_��������, g_Ed_���½���, g_Ed_��������
        IDKIND.Enabled = False
        txtPatient.Locked = True
        txtPatient.Locked = True
        chkCancel.Enabled = Not mPatiInfor.bln�������� And IIf(mEditType = g_Ed_������� Or mEditType = g_Ed_סԺ����, True, False)
        cmdYBBalance.Enabled = False
        txtPatiBegin.Enabled = False
        txtPatiEnd.Enabled = False
        opt��;.Enabled = False
        opt��Ժ.Enabled = False
        cboPatiNums.Enabled = False
        
        cboӤ��.Enabled = False
        cboDept.Enabled = False
        cboFeeType.Enabled = False
        cboFeeItem.Enabled = False
        cboChargeType.Enabled = False
        txtBegin.Enabled = False
        txtEnd.Enabled = False
        txtInvoice.Enabled = IIf(mEditType = g_Ed_���½���, True, False)
        
        txtBalance(Idx_���ν���).Enabled = False
        txtBalance(Idx_����˵��).Enabled = blnEanbled And mEditType = g_Ed_���½���
        If mEditType = g_Ed_���ݲ鿴 Or mEditType = g_Ed_ȡ������ Or mEditType = g_Ed_�������� Then
            txtInvoice.Enabled = False
            cboNO.Enabled = False
        End If
    End Select
    
    txtBegin.BackColor = IIf(txtBegin.Enabled, &H80000005, &H8000000F)
    txtEnd.BackColor = IIf(txtEnd.Enabled, &H80000005, &H8000000F)
          
    txtPatiBegin.BackColor = IIf(txtPatiBegin.Enabled, &H80000005, &H8000000F)
    txtPatiEnd.BackColor = IIf(txtPatiEnd.Enabled, &H80000005, &H8000000F)
    txtBalance(Idx_���ν���).BackColor = IIf(txtBalance(Idx_���ν���).Enabled, &H80000005, &H8000000F)
    txtBalance(Idx_����˵��).BackColor = IIf(txtBalance(Idx_����˵��).Enabled, &H80000005, &H8000000F)
    cboNO.BackColor = IIf(cboNO.Enabled, &H80000005, &H8000000F)
    txtInvoice.BackColor = IIf(txtInvoice.Enabled, &H80000005, &H8000000F)
            
    cboӤ��.BackColor = IIf(cboӤ��.Enabled, &H80000005, &H8000000F)
    cboPatiNums.BackColor = IIf(cboPatiNums.Enabled, &H80000005, &H8000000F)
    cboFeeType.BackColor = IIf(cboFeeType.Enabled, &H80000005, &H8000000F)
    cboFeeItem.BackColor = IIf(cboFeeItem.Enabled, &H80000005, &H8000000F)
    cboDept.BackColor = IIf(cboDept.Enabled, &H80000005, &H8000000F)
    cboChargeType.BackColor = IIf(cboChargeType.Enabled, &H80000005, &H8000000F)
    cboDiag.BackColor = IIf(cboDiag.Enabled, &H80000005, &H8000000F)
    
            
            
End Sub


Private Sub InitBalanceCondition()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������������ر���
    '����:���˺�
    '����:2014-12-31 11:46:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mobjBalanceAll = New clsBalanceAllCon
    With mobjBalanceAll
        .strAllTime = ""
        .strAllDeptIDs = ""
        .strAllItem = ""
        .strAllDiag = ""
        .strAllClass = ""
        .strUnAuditTime = ""
        .strAllChargeType = ""  '34260
        .MinDate = #1/1/1900#
        .MaxDate = #1/1/1900#
    End With
End Sub

Private Sub InitPatiBalanceVariableCon()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��̨��������ؽ�����������
    '����:���˺�
    '����:2014-12-31 11:56:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mobjBalanceCon = New clsBalanceCon
    With mobjBalanceCon
        .strTime = ""
        .strDeptIDs = ""
        .strClass = ""
        .strBaby = ""
        .strItem = ""
        .bytKind = 0
        .dtBeginDate = CDate("0:00:00"):
        .dtEndDate = CDate("0:00:00")
        .strChargeType = ""
        .blnCurBalanceOwnerFee = False
    End With
End Sub

Private Sub FindPati(ByVal objCard As Card, ByVal blnCard As Boolean, ByVal strInput As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���Ҳ���
    '����:���˺�
    '����:2012-08-29 17:53:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call NewBill
    txtPatient.Text = strInput
    '���˺�:27503
    If mty_ModulePara.bln���ʺ�����Ϣ Then
        If txtInvoice.Tag <> "" And txtInvoice.Text <> txtInvoice.Tag Then txtInvoice.Text = txtInvoice.Tag '��Ҫ��Ҫ������Ϣ,��ȷ������Ҫ�����̶�
    End If
    
    If mOldOneCard.blnOneCard And Not mobjICCard Is Nothing And objCard.���� Like "IC��*" And objCard.ϵͳ Then
        Call SetOldOneCardBalance  '��ʾ��һ��ͨ���
    End If
    
    Call LoadPatientInfo(objCard, blnCard)
    
'    If Not mrsInfo Is Nothing Then
'        If mrsInfo.State <> 0 Then
'            If InStr("," & mobjBalanceCon.strTime & ",", "," & mrsInfo!��ҳID & ",") = 0 Then Call cmdRefresh_Click
'        End If
'    End If
End Sub
Private Sub SetOldOneCardBalance()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����һ��ͨ���㷽ʽ
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-04 09:55:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim curOneCard As Currency, strName As String
    If mOldOneCard.blnOneCard = False Or mobjICCard Is Nothing Then Exit Sub
    curOneCard = mobjICCard.GetSpare(strName)
    If curOneCard <> 0 Then
       mOldOneCard.rsOneCard.Filter = "����='" & strName & "'"
       If mOldOneCard.rsOneCard.RecordCount > 0 Then mOldOneCard.strOneCard = mOldOneCard.rsOneCard!���㷽ʽ
    End If
    staThis.Panels(2).Text = "�����:" & Format(curOneCard, "0.00") & "Ԫ"
End Sub


Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, _
    ByVal blnCard As Boolean, Optional ByVal lng��ҳID As Long, _
    Optional blnOnlyReadPati As Boolean, Optional ByVal intInsure As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ
    '���:blnCard=�Ƿ���￨ˢ��
    '     lng��ҳID=��ȡָ��סԺ�����Ĳ�����Ϣ
    '     intInsure-����(��Ҫ���ؽ������ʱ����)
    '     blnOnlyReadPati-ֻ��ȡ������Ϣ��������ؼ��(��Ҫ���ؽ������ʱ����)
    '����:
    '����:�Ƿ��ȡ�ɹ�,�ɹ�ʱmrsInfo�а���������Ϣ,ʧ��ʱmrsInfo=Close,strInput�����������ж��Ƿ�����ʾ��,�����ٴ���ʾû���ҵ�����
    '����:���˺�
    '����:2015-01-04 12:15:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnCancel As Boolean, strWhere As String, strField As String, bytMzMode As Byte
    Dim strSQL As String, lng�����ID As Long, strPassWord As String, strErrMsg As String
    Dim lng����ID As Long, blnHavePassWord As Boolean, rsIn As ADODB.Recordset, strInSQL As String
    Dim strPati As String, strRange As String, rsTemp As ADODB.Recordset
    Dim vRect As RECT
    Dim strסԺ�� As String
    

    mstrPassWord = "": mstrInputInNo = "": mblnReadByZYNo = False
    mlngCardTypeID = 0
    On Error GoTo errH
    
    strField = ",A.��ǰ����ID"
    
    bytMzMode = mYBInFor.bytMCMode
    
    
    If mEditType = g_Ed_סԺ���� Then
        If Not (blnCard = True And objCard.���� Like "����*") Then    '��ˢ������
            If Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then 'סԺ��
                strסԺ�� = Val(Mid(strInput, 2))
            ElseIf objCard.���� = "סԺ��" Then
                strסԺ�� = Val(strInput)
            End If
        End If
    End If
    
    
    
    If mlngPatientID <> 0 And mblnFirst Then
        '�����⴫��Ĳ���ID�������ж�
        '��һ��ȡ��ʱ
        lng��ҳID = Val(mstr��ҳId)
        If Val(mstr��ҳId) = 0 Then '����
            strWhere = strWhere & " And B.��ҳID(+)=-100"
            bytMzMode = IIf(bytMzMode = 0, 0, 1): strField = " ,NULL as ��ǰ����ID"
            If mEditType = g_Ed_סԺ���� Then bytMzMode = 2    'סԺ��:44022
        Else    'ָ������
            strWhere = strWhere & "  And B.��ҳID=[3]"
            bytMzMode = 2   'סԺ��
        End If
    Else
        If mEditType = g_Ed_������� Then   '����
            strWhere = strWhere & " And   A.��ҳID=B.��ҳID(+)"
            '����:43730
            bytMzMode = IIf(bytMzMode = 0, 0, 1): strField = " ,NULL as ��ǰ����ID"
        Else
            If lng��ҳID <> 0 Then
                strField = ",Decode(A.��ҳID,[3],A.��ǰ����ID,NULL) as ��ǰ����ID"
                strWhere = " And B.��ҳID=[3]"
            ElseIf strסԺ�� <> "" Then '��סԺ�Ų��Ҳ���
                strWhere = "And (B.����ID,B.��ҳID)=(Select max(����ID)as ����ID, Max(��ҳID) As ��ҳID From ������ҳ Where סԺ��=[2])"
            Else
                strWhere = " And A.��ҳID=B.��ҳID(+)"
            End If
            bytMzMode = 2
        End If
    End If
    

    If intInsure <> 0 Then
        strField = strField & ",[4] as ����"
    ElseIf bytMzMode = 0 Then
        strField = strField & ",NULL as ����"
    ElseIf bytMzMode = 1 Then
        strField = strField & ",A.���� as ����"
    Else
        strField = strField & ",B.���� as ����"
    End If

    
    strSQL = _
    " Select A.����ID,Nvl(B.��ҳID,0) as ��ҳID,A.�����,nvl(B.סԺ��,A.סԺ��) as סԺ��,B.��Ժ����,B.��Ժ����," & _
    "       nvl(B.����,A.����) as ����, nvl(B.�Ա�,Nvl(A.�Ա�,'δ֪')) as  �Ա�,Nvl(B.����,A.����) as ����,A.IC����,A.���￨��,A.����֤��," & _
    "       Nvl(B.�ѱ�,A.�ѱ�) as �ѱ�,C.���� as ��Ժ����" & strField & ",D.���� as ��Ժ����,B.��Ժ����ID," & _
    "       E.����,E.ҽ����,E.����," & _
    "       A.�Ǽ�ʱ��,Nvl(B.״̬,0) as ״̬,Nvl(B.ҽ�Ƹ��ʽ,A.ҽ�Ƹ��ʽ) as ҽ�Ƹ��ʽ,Nvl(B.��˱�־,0) as ��˱�־, " & _
    "       B.��Ժ����,B.��Ժ����,B.��������,B.��������,Decode(B.����ID,Null,A.��Ժ,Decode(B.��Ժ����,Null,1,0)) As ��Ժ" & _
    " From ������Ϣ A,������ҳ B,���ű� C,���ű� D,ҽ�����˵��� E,ҽ�����˹����� F" & _
    " Where A.ͣ��ʱ�� is NULL And A.����ID=B.����ID(+)   " & strWhere & _
    "   And A.����ID=F.����ID(+) And F.��־(+)=1 And F.ҽ����=E.ҽ����(+) And F.����=E.����(+) And F.���� = E.����(+)" & _
    "   And B.��Ժ����ID=C.ID(+) And B.��Ժ����ID=D.ID(+)"
        
    If blnCard = True And objCard.���� Like "����*" Then    'ˢ��
        If IDKIND.Cards.��ȱʡ������ And Not IDKIND.GetfaultCard Is Nothing Then
            lng�����ID = IDKIND.GetfaultCard.�ӿ����
        Else
            lng�����ID = "-1"
        End If
        
        '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�);��
        If gobjSquare.objSquareCard.zlGetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
        If lng����ID <= 0 Then GoTo NotFoundPati:
        strInput = "-" & lng����ID
        blnHavePassWord = True
        mlngCardTypeID = lng�����ID
        strSQL = strSQL & " And A.����ID=[1] "
        
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then '����ID
        strSQL = strSQL & " And A.����ID=[1]"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '�����
        strSQL = strSQL & " And A.�����=[1]"
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then 'סԺ��
        strInput = Mid(strInput, 2)
        
        If mEditType <> g_Ed_סԺ���� Then
            strSQL = strSQL & " And A.����ID=(Select nvl(Max(����ID),0) As ����ID From ������ҳ   Where  סԺ��=[2])"
        Else
           mblnReadByZYNo = True
           mstrInputInNo = mobjBalanceAll.zlGetNumsFromZyNo(Val(strInput))
           If InStr(1, mstrInputInNo, ",") > 0 Then mstrInputInNo = "": mblnReadByZYNo = False
        End If

    Else '��������
        mlngCardTypeID = objCard.�ӿ����
        Select Case objCard.����
            Case "����", "��������￨"
                If mrsInfo.State = 1 Then
                    If mrsInfo!���� = Trim(txtPatient.Text) Then
                        GetPatient = True
                        Exit Function
                    End If
                End If
                
                If mty_ModulePara.intPatientRange > 0 Then
                    Select Case mty_ModulePara.intPatientRange
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
                strPati = "" & _
                " Select A.����ID as ID,A.����ID,A.����,A.סԺ��, A.�����, nvl(B.�Ա�,Nvl(A.�Ա�,'δ֪')) as  �Ա�, A.����, A.סԺ����, A.��ͥ��ַ, A.������λ," & vbNewLine & _
                "   To_Char(A.��������,'YYYY-MM-DD') as ��������,  To_Char(B.��Ժ����,'YYYY-MM-DD') as ��Ժ����, To_Char(B.��Ժ����,'YYYY-MM-DD') as ��Ժ����" & vbNewLine & _
                " From ������Ϣ A, ������ҳ B" & vbNewLine & _
                " Where A.����id = B.����id(+) And A.��ҳID = B.��ҳid(+) And A.ͣ��ʱ�� Is Null And A.���� = [1] " & vbNewLine & strPati & vbNewLine & _
                " Order By Decode(סԺ��, Null, 1, 0), ��Ժ���� Desc"
                        
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
            Case "���֤��", "�������֤", "���֤"
                If gobjSquare.objSquareCard.zlGetPatiID("���֤", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                If lng����ID <= 0 Then GoTo NotFoundPati:
                strInput = "-" & lng����ID
                blnHavePassWord = True
                strSQL = strSQL & " And A.����ID=[1] "
            Case "IC����"
                If gobjSquare.objSquareCard.zlGetPatiID("IC��", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                If lng����ID <= 0 Then GoTo NotFoundPati:
                strInput = "-" & lng����ID
                blnHavePassWord = True
                strSQL = strSQL & " And A.����ID=[1] "
            Case "�����"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = strSQL & " And A.�����=[2]"
            Case "סԺ��"
                If Not IsNumeric(strInput) Then strInput = "0"
                    If mEditType <> g_Ed_סԺ���� Then
                    strSQL = strSQL & " And A.����ID=(Select nvl(Max(����ID),0) As ����ID From ������ҳ   Where  סԺ��=[2])"
                Else
                   mblnReadByZYNo = True
                   mstrInputInNo = mobjBalanceAll.zlGetNumsFromZyNo(Val(strInput))
                   If InStr(1, mstrInputInNo, ",") > 0 Then mstrInputInNo = "": mblnReadByZYNo = False
                   
                End If
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
    
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput, lng��ҳID, intInsure)
    If mrsInfo.EOF Then Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "": Exit Function
    If mstr��ҳId <> "" Then mstrInputInNo = mstr��ҳId: mblnReadByZYNo = True: mstr��ҳId = ""
    mYBInFor.intInsure = Val(NVL(mrsInfo!����))
    
    mstrPassWord = strPassWord
    If Not blnHavePassWord Then
        mstrPassWord = NVL(mrsInfo!����֤��)
    End If
    
    If blnOnlyReadPati Then GetPatient = True: Exit Function
    
    '����������:�����������ʾ
    '34681:35686
    If zlCheckPatiIsDeath(Val(NVL(mrsInfo!����ID))) = True Then
        pic����.Visible = True
        If MsgBox("ע��:" & vbCrLf & "    �ò����Ѿ�����,�Ƿ��������?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            pic����.Visible = False
            Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "": Exit Function
        End If
    Else
        pic����.Visible = False
    End If
    
    '��Ҫ�ٴμ��,�Է������ڼ�����˵Ĳ��˱�ȡ�����
    '36209
    If (InStr(mstrPrivs, ";δ��˲�����;����;") = 0 And opt��;.Value _
        Or InStr(mstrPrivs, ";δ��˲��˳�Ժ����;") = 0 And opt��Ժ.Value) _
        And mEditType = g_Ed_סԺ���� Then
        If Not Chk�������(mrsInfo!����ID, Val(NVL(mrsInfo!��ҳID))) Then
            If MsgBox("�����ʷ����а������˵�" & Val(NVL(mrsInfo!��ҳID)) & "��סԺδ��˵ķ��ü�¼��" & vbCrLf & _
                " �Ƿ��������?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "": Exit Function
            End If
        End If
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

Private Sub LoadPatientInfo(ByVal objCard As Card, ByVal blnCard As Boolean, _
    Optional ByVal intInsure As Integer, _
    Optional ByVal lng��ҳID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ز�����Ϣ
    '���:objCard-��ǰˢ����Ŀ�����
    '     blnCard-�Ƿ�ˢ��
    '     intInsure-��ǰ������
    '     lng��ҳID-��ȡָ��סԺ�����Ĳ�����Ϣ
    '����:���˺�
    '����:2015-01-04 12:12:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String, i As Long, strSQL As String
    Dim tyPatiInfor As ty_Pati_Infor
    Dim blnICCard As Boolean, curDue As Currency, blnIDCard As Boolean
    Dim blnNotClearPati As Boolean
    Dim lngPageID As Long
    Dim strPage() As String
    
    On Error GoTo errHandle
    Dim rsTmp As New ADODB.Recordset

    txtPatient.ForeColor = Me.ForeColor
    
    mPatiInfor = tyPatiInfor '��ղ�����Ϣ
    If objCard.���� Like "IC��*" And objCard.ϵͳ = True Then blnICCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
    If objCard.���� Like "*���֤*" And objCard.ϵͳ = True Then blnIDCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
    
    staThis.Panels(2).Text = ""
    
    If Not GetPatient(objCard, Trim(txtPatient.Text), blnCard, lng��ҳID, , intInsure) Then
        If txtPatient.Text = "" Then MsgBox "û���ҵ��ò���,�������������Ƿ���ȷ��", vbInformation, gstrSysName
        txtPatient.PasswordChar = "": txtPatient.Text = ""
        '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
        txtPatient.IMEMode = 0
        If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
        mstr����סԺ���� = ""
        Call ReInitPatiInvoice
        Exit Sub
    End If
    
    mstr����סԺ���� = ""
    '���￨������
    If (objCard.���� Like "IC��*" Or objCard.���� Like "*���֤*") And objCard.ϵͳ = True And blnCard Then blnCard = False
    
    If gTy_System_Para.TY_Balance.blnˢ���������� _
        And (blnCard Or ((blnICCard Or blnIDCard Or IDKIND.GetCurCard.�ӿ���� <> 0) And mstrPassWord <> "")) Then
        If Not zlCommFun.VerifyPassWord(Me, mstrPassWord, mrsInfo!����, mrsInfo!�Ա�, "" & mrsInfo!����) Then
            GoTo ExitHandle
        End If
    End If
    
    '102236,������Ҳ����ӿ�
    If Not gobjPlugIn Is Nothing Then
        On Error Resume Next
        'PatiValiedCheck(ByVal lngSys As Long, ByVal lngModule As Long, _
        '    ByVal lngType As Long, ByVal lngPatiID As Long, ByVal lngPageID As Long, _
        '    ByVal strPatiInforXML As String, Optional ByRef strReserve As String) As Boolean
        ''���ܣ���鵱ǰ�����Ƿ���ָ�������ⲡ��
        ''���أ�trueʱ�������������Falseʱ���������
        ''������
        ''      lngSys,lngModual=��ǰ���ýӿڵ�������ϵͳ�ż�ģ���
        ''      lngType �������ͣ�1������Һţ�2��סԺ��Ժ��3�������շѣ�4��סԺ���ʣ�5��������ʡ�
        ''      lngPatiID-����ID: �½����ģ�Ϊ0,�����뽨������ID
        ''      lngPageID-��ҳID: �½����ģ�Ϊ0,�����뽨����ҳID(סԺ������ҳID) ����˵������ lngType=4 ʱ�Ŵ��� lngPageID����������0
        ''      strPatiInforXML-������Ϣ:���δ�������˴��룬"�������Ա����䣬�������ڣ�ҽ���ţ����֤��"���������� ��ʽ:2016-11-11 12:12:12
        ''                      �̶���ʽ��<XM></XM><XB></XB><NL></NL><CSRQ></CSRQ><YBH></YBH><SFZH></SFZH>
        ''      strReserve=��������,������չʹ��
        Dim blnChecked As Boolean
        blnChecked = gobjPlugIn.PatiValiedCheck(glngSys, mlngModul, IIf(mEditType = g_Ed_�������, 5, 4), Val(NVL(mrsInfo!����ID)), Val(NVL(mrsInfo!��ҳID)), "")
        If Err <> 0 Then
            Call zlPlugInErrH(Err, "PatiValiedCheck"): Err.Clear
        Else
            If blnChecked = False Then GoTo ExitHandle
        End If
        On Error GoTo errHandle
    End If
        
    '����:27690
    If mYBInFor.intInsure = 0 Then
        If InStr(1, mstrPrivs, ";��ͨ���˽���;") = 0 Then
            MsgBox "��û��Ȩ�޶ԷǱ��ղ��˽��н��㡣", vbInformation, gstrSysName
            GoTo ExitHandle
        End If
    End If
    
    'ҽ������ж�
    If mYBInFor.intInsure <> 0 Then
        If InStr(mstrPrivs, ";���ս���;") = 0 Then
            MsgBox "��û��Ȩ�޶Ա��ղ��˽��н��㡣", vbInformation, gstrSysName
            GoTo ExitHandle
        End If
        
        If mYBInFor.strYBPati <> "" And intInsure <> mYBInFor.intInsure Then
            MsgBox "���˵Ǽǵ�������ҽ�������֤�����಻����", vbInformation, gstrSysName
            GoTo ExitHandle
        End If
        
        If mYBInFor.bytMCMode = 1 And Not IsNull(mrsInfo!��ǰ����id) Then
            MsgBox "��Ժ���˲��ܽ�������ҽ�������֤��", vbInformation, gstrSysName
            GoTo ExitHandle
        End If
        Call InitInsurePara(Val(NVL(mrsInfo!����ID)), mYBInFor.intInsure)
    ElseIf mYBInFor.strYBPati <> "" Then
        MsgBox "���������֤�ɹ�,�����˵Ǽǵ�����Ϊ�գ�", vbInformation, gstrSysName
        GoTo ExitHandle
    End If
    
    If mblnReadByZYNo Then
        strPage = Split(mstrInputInNo, ",")
        For i = 0 To UBound(strPage)
            If Val(strPage(i)) > lngPageID Then lngPageID = Val(strPage(i))
        Next i
        '����:34763 ��鲡���Ƿ���ڱ�ע��Ϣ
        If zlCheckPatiIsMemo(Val(NVL(mrsInfo!����ID)), lngPageID) = True Then
            Call zlCallPatiMemoWriteAndRead(Me, mlngModul, mstrPrivs, Val(NVL(mrsInfo!����ID)), lngPageID, mobjInPatient)
        End If
        
        If lng��ҳID = 0 Then
            '����ȱʡ��Ժ״̬
            If Not LoadDefaultOutStatu(mrsInfo!����ID, lngPageID) Then GoTo ExitHandle
            '����������
            If Not CheckPatiBlacklist(mrsInfo!����ID) Then GoTo ExitHandle
                                                                                        
            '����δ��˼��
            If Not CheckChargeAudit(mrsInfo!����ID) Then GoTo ExitHandle
    
            '�Զ����㲡�˵Ĵ�λ���úͻ�������
            Call AutoCalcChareFee(Val(NVL(mrsInfo!����ID)), lngPageID)
            
            '���ز��������Ϣ
            Call Load�����Ϣ(Val(NVL(mrsInfo!����ID)), IIf(mEditType = g_Ed_������� Or mblnCurMzBalanceNo, 1, 2))
            
            '���غͼ��Ӧ�տ����
            Call LoadӦ�տ���Ϣ(Val(NVL(mrsInfo!����ID)))
            '88786,���ʲ�������ʷ����
            mblnDateMoved = False
        Else
            If Val(NVL(mrsInfo!��Ժ)) = 1 And NVL(mrsInfo!״̬, 0) <> 3 Then '��Ժ����()
                '״̬:0-����סԺ��1-��δ��ƣ�2-����ת�ƻ�����ת������3-��Ԥ��Ժ
                If zlDatabase.GetPara("Ĭ�ϳ�Ժ����", glngSys, mlngModul, "1") <> "0" Then
                    opt��Ժ.Value = True
                    opt��;.Value = False
                Else
                    opt��;.Value = True
                    opt��Ժ.Value = False
                End If
                If gbln��Ժ��׼���� Then opt��;.Value = True: opt��Ժ.Enabled = False
            Else
                '��Ժ����(����Ԥ��Ժ�Ĳ���)
                 opt��Ժ.Value = True
                 opt��;.Value = False
                 opt��Ժ.Enabled = True
            End If
        End If
    Else
        '����:34763 ��鲡���Ƿ���ڱ�ע��Ϣ
        If zlCheckPatiIsMemo(Val(NVL(mrsInfo!����ID)), Val(NVL(mrsInfo!��ҳID))) = True Then
            Call zlCallPatiMemoWriteAndRead(Me, mlngModul, mstrPrivs, Val(NVL(mrsInfo!����ID)), Val(NVL(mrsInfo!��ҳID)), mobjInPatient)
        End If
        
        If lng��ҳID = 0 Then
            '����ȱʡ��Ժ״̬
            If Not LoadDefaultOutStatu(mrsInfo!����ID, Val(NVL(mrsInfo!��ҳID))) Then GoTo ExitHandle
            '����������
            If Not CheckPatiBlacklist(mrsInfo!����ID) Then GoTo ExitHandle
                                                                                        
            '����δ��˼��
            If Not CheckChargeAudit(mrsInfo!����ID) Then GoTo ExitHandle
    
            '�Զ����㲡�˵Ĵ�λ���úͻ�������
            Call AutoCalcChareFee(Val(NVL(mrsInfo!����ID)), Val(NVL(mrsInfo!��ҳID)))
            
            '���ز��������Ϣ
            Call Load�����Ϣ(Val(NVL(mrsInfo!����ID)), IIf(mEditType = g_Ed_������� Or mblnCurMzBalanceNo, 1, 2))
            
            '���غͼ��Ӧ�տ����
            Call LoadӦ�տ���Ϣ(Val(NVL(mrsInfo!����ID)))
            '88786,���ʲ�������ʷ����
            mblnDateMoved = False
        Else
            If Val(NVL(mrsInfo!��Ժ)) = 1 And NVL(mrsInfo!״̬, 0) <> 3 Then '��Ժ����()
                '״̬:0-����סԺ��1-��δ��ƣ�2-����ת�ƻ�����ת������3-��Ԥ��Ժ
                If zlDatabase.GetPara("Ĭ�ϳ�Ժ����", glngSys, mlngModul, "1") <> "0" Then
                    opt��Ժ.Value = True
                    opt��;.Value = False
                Else
                    opt��;.Value = True
                    opt��Ժ.Value = False
                End If
                If gbln��Ժ��׼���� Then opt��;.Value = True: opt��Ժ.Enabled = False
            Else
                '��Ժ����(����Ԥ��Ժ�Ĳ���)
                 opt��Ժ.Value = True
                 opt��;.Value = False
                 opt��Ժ.Enabled = True
            End If
        End If
    End If

    txtPatient.PasswordChar = ""
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txtPatient.IMEMode = 0
    txtPatient.Text = mrsInfo!����: txtSex.Text = NVL(mrsInfo!�Ա�): txtOld.Text = NVL(mrsInfo!����)
    With mPatiInfor
        .lng����ID = Val(NVL(mrsInfo!����ID))
        .lng��ҳID = Val(NVL(mrsInfo!��ҳID))
        .str���� = NVL(mrsInfo!����)
        .str�Ա� = NVL(mrsInfo!�Ա�)
        .str���� = NVL(mrsInfo!����)
        .bln��Ժ = Val(NVL((mrsInfo!��Ժ))) <> 1
    End With
'    mYBInFor.intInsure = Val(NVL(mrsInfo!����))
    '���ز���״̬
    Call LoadסԺ״̬(Val(NVL(mrsInfo!����ID)))
    
    cmdYB.Enabled = IIf(mEditType = g_Ed_�������, True, False)
    If mYBInFor.intInsure <> 0 Then
        staThis.Panels(4).Text = GetInsureName(mYBInFor.intInsure)
        staThis.Panels(4).Visible = True
        If mYBInFor.bytMCMode = 1 Then Call SetPatiEnabled(False)
        cmdOK.Enabled = False
    Else
        staThis.Panels(4).Visible = False
    End If
    
    If NVL(mrsInfo!��������) = "" And mYBInFor.intInsure <> 0 Then
        txtPatient.ForeColor = vbRed
    Else
        txtPatient.ForeColor = zlDatabase.GetPatiColor(NVL(mrsInfo!��������))
    End If
    
    txtPatiType.Text = NVL(mrsInfo!��������)
    
    txt�ѱ�.Text = NVL(mrsInfo!�ѱ�)
    
    If mEditType = g_Ed_סԺ���� Then
        If Not IsNull(mrsInfo!סԺ��) Then
            txt��ʶ��.Text = mrsInfo!סԺ��
            lbl��ʶ��.Visible = True: txt��ʶ��.Visible = True
            lbl��ʶ��.Caption = "סԺ��"
        End If
        If Not IsNull(mrsInfo!��Ժ����) Then
            txtBed.Text = "" & NVL(mrsInfo!��Ժ����, mrsInfo!��Ժ����)
            txt����.Text = NVL(mrsInfo!��Ժ����, mrsInfo!��Ժ����)
            lblBed.Visible = True: txtBed.Visible = True
            lbl����.Visible = True: txt����.Visible = True
        ElseIf Not IsNull(mrsInfo!��Ժ����) Then
            txtBed.Text = NVL(mrsInfo!��Ժ����)
            txt����.Text = mrsInfo!��Ժ����
            lblBed.Visible = True: txtBed.Visible = True
            lbl����.Visible = True: txt����.Visible = True
        End If
    ElseIf mEditType = g_Ed_������� Then
        If Not IsNull(mrsInfo!�����) Then
            txt��ʶ��.Text = mrsInfo!�����
            lbl��ʶ��.Visible = True: txt��ʶ��.Visible = True
            lbl��ʶ��.Caption = "�����"
        End If
    End If
    
    '�쳣���ݴ���
    If PatiErrBillPay(Val(NVL(mrsInfo!����ID))) Then Exit Sub
    
    '��ʾ����Ҫ��������,����ʼ��������
    '-------------------------------------------------------------------------------------------
    If lng��ҳID = 0 Then
        strTmp = ""
        If Not ShowBalance(True, strTmp, blnNotClearPati) Then
            If strTmp <> "" Then MsgBox strTmp, vbInformation, gstrSysName
            If blnNotClearPati = False Then GoTo ExitHandle:
            If cmdRefresh.Enabled And cmdRefresh.Visible Then cmdRefresh.SetFocus
            cmdYB.Enabled = IIf(mEditType = g_Ed_�������, True, False)
            Exit Sub
        End If
        Call Led��ӭ��Ϣ
    End If
     
    Call ReInitPatiInvoice  '����ˢ�·�Ʊ��Ϣ
    
    
    
    mblnNotChange = True
    Call txtBalance_Validate(Idx_��Ԥ��, False)
    mblnNotChange = False
    
  
    If cmdOK.Visible And cmdOK.Enabled Then cmdOK.SetFocus
    
    If cmdYBBalance.Visible And cmdYBBalance.Enabled Then
        cmdYBBalance.SetFocus
    ElseIf txtBalance(Idx_��Ԥ��).Enabled And txtBalance(Idx_��Ԥ��).Visible Then
        txtBalance(Idx_��Ԥ��).SetFocus
    ElseIf txtBalance(Idx_�ɿ�).Enabled And txtBalance(Idx_�ɿ�).Visible Then
       If mrsInfo Is Nothing Then
         If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
       ElseIf mrsInfo.State <> 1 Then
         If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
       Else
         If txtBalance(Idx_�ɿ�).Enabled And txtBalance(Idx_�ɿ�).Visible Then txtBalance(Idx_�ɿ�).SetFocus
       End If
    Else
        zlCommFun.PressKey vbKeyTab
    End If

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
ExitHandle:
    Call NewBill
    If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
    Screen.MousePointer = 0
    Exit Sub
End Sub



Private Sub InitInsurePara(ByVal lng����ID As Long, ByVal intInsure As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ҽ������
    '����:���˺�
    '����:2015-01-04 13:48:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    MCPAR.�ֱҴ��� = gclsInsure.GetCapability(support�ֱҴ���, lng����ID, intInsure)
    MCPAR.ҽ���ӿڴ�ӡƱ�� = gclsInsure.GetCapability(supportҽ���ӿڴ�ӡƱ��, lng����ID, intInsure)
    MCPAR.�������Ϻ��ӡ�ص� = gclsInsure.GetCapability(support�������Ϻ��ӡ�ص�, lng����ID, intInsure)
    If mYBInFor.bytMCMode = 1 Then
        MCPAR.����Ԥ���� = gclsInsure.GetCapability(support����Ԥ��, lng����ID, intInsure)
        MCPAR.������봫����ϸ = gclsInsure.GetCapability(support������봫����ϸ, lng����ID, intInsure)
        MCPAR.�������_�������� = gclsInsure.GetCapability(support�������_�������ú���ýӿ�, lng����ID, intInsure)
        MCPAR.���ﲡ�˽������� = gclsInsure.GetCapability(support�����������, lng����ID, intInsure)
    Else
        MCPAR.δ�����Ժ = gclsInsure.GetCapability(supportδ�����Ժ, lng����ID, intInsure)
        MCPAR.����ʹ�ø����ʻ� = gclsInsure.GetCapability(support����ʹ�ø����ʻ�, lng����ID, intInsure)
        MCPAR.��Ժ��������Ժ = gclsInsure.GetCapability(support��Ժ��������Ժ, lng����ID, intInsure)
        MCPAR.��;������������ϴ����� = gclsInsure.GetCapability(support��;������������ϴ�����, lng����ID, intInsure)
        MCPAR.�������ú���ýӿ� = gclsInsure.GetCapability(support����_�������ú���ýӿ�, lng����ID, intInsure)
        MCPAR.סԺ�������� = gclsInsure.GetCapability(supportסԺ��������, lng����ID, intInsure)
        MCPAR.�������_�������� = False
        MCPAR.��Ժ���˽������� = gclsInsure.GetCapability(support��Ժ���˽�������, lng����ID, intInsure)
        MCPAR.�������סԺ���� = gclsInsure.GetCapability(support����һ�ν���סԺ����, lng����ID, intInsure)
        
    End If
End Sub

Private Function LoadDefaultOutStatu(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ĭ�ϵĳ�Ժ״̬
    '����:���˺�
    '����:2015-01-04 14:00:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
        
    On Error GoTo errHandle
    
    If mYBInFor.bytMCMode = 1 Then LoadDefaultOutStatu = True: Exit Function
    If mEditType = g_Ed_������� Then LoadDefaultOutStatu = True: Exit Function
    
    If lng��ҳID = 0 Then
        opt��Ժ.Value = True: opt��Ժ.Enabled = False
        opt��;.Enabled = False: LoadDefaultOutStatu = True: Exit Function
    Else
        'Ĭ�Ͻ���ǰסԺ������,��Ժ����
        If lng��ҳID < Val(NVL(mrsInfo!��ҳID)) Then
            opt��Ժ.Value = True: LoadDefaultOutStatu = True: Exit Function
        End If
    End If
    
    '����:30027:����ȱʡ����;����
    '       1.��Ժ����,Ĭ��Ϊ��Ժ���� ����:û��"��;����"Ȩ�޵�,ҲĬ��Ϊ��Ժ����
    '       2.��Ժ����(�����ϴγ�Ժ���˵�ѡ���Ϊ׼)
    '              Ĭ�ϳ�Ժ��(���ϴ�ѡ�����;���ʻ�סԺ����)����Ϊtrue,Ĭ��Ϊ��Ժ����,����Ĭ��Ϊ��;����
    If InStr(mstrPrivs, ";��;����;") = 0 Then
        opt��Ժ.Value = True: opt��;.Enabled = False
    ElseIf Val(NVL(mrsInfo!��Ժ)) = 1 And NVL(mrsInfo!״̬, 0) <> 3 Then '��Ժ����()
        '״̬:0-����סԺ��1-��δ��ƣ�2-����ת�ƻ�����ת������3-��Ԥ��Ժ
        If zlDatabase.GetPara("Ĭ�ϳ�Ժ����", glngSys, mlngModul, "1") <> "0" Then
            opt��Ժ.Value = True
        Else
            opt��;.Value = True
        End If
        If gbln��Ժ��׼���� Then opt��;.Value = True: opt��Ժ.Enabled = False
    Else
        '��Ժ����(����Ԥ��Ժ�Ĳ���)
         opt��Ժ.Value = True
    End If
    opt��Ժ.Enabled = True
    
    If CheckOutBalanceIsvalied = False Then Exit Function
    
    If mEditType = g_Ed_������� Then
        If Val(NVL(mrsInfo!��Ժ)) = 1 And NVL(mrsInfo!״̬, 0) <> 3 Then
            If MsgBox("��ǰ������Ժ����Ҫ�����Ըò��˽������������?", vbYesNo + vbInformation + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        End If
    Else
        If Val(NVL(mrsInfo!��Ժ)) = 1 And NVL(mrsInfo!״̬, 0) <> 3 And gbln��Ժ��׼���� Then
            If MsgBox("��ǰ������Ժ���������Ժ���ʡ� ����ǳ�Ժ���ʣ����Ƚ����˳�Ժ��" & _
                vbCrLf & "��Ҫ�Ըò��˽�����;������?", vbYesNo + vbInformation + vbDefaultButton2, gstrSysName) <> vbYes Then Exit Function
        End If
    End If
    
    If mblnFirst And mlngPatientID <> 0 Then
        If Val(NVL(mrsInfo!��Ժ)) = 1 And NVL(mrsInfo!״̬, 0) <> 3 Then '��Ժ����()
            '״̬:0-����סԺ��1-��δ��ƣ�2-����ת�ƻ�����ת������3-��Ԥ��Ժ
            If zlDatabase.GetPara("Ĭ�ϳ�Ժ����", glngSys, mlngModul, "1") <> "0" Then
                opt��Ժ.Value = True
                opt��;.Value = False
            Else
                opt��;.Value = True
                opt��Ժ.Value = False
            End If
            If gbln��Ժ��׼���� Then opt��;.Value = True: opt��Ժ.Enabled = False
        Else
            '��Ժ����(����Ԥ��Ժ�Ĳ���)
             opt��Ժ.Value = True
             opt��;.Value = False
        End If
        
        LoadDefaultOutStatu = True: Exit Function
    End If
    
    If opt��;.Value Then
        opt��Ժ.Value = False
        If gbln��Ժ��׼���� Then opt��Ժ.Enabled = False
        LoadDefaultOutStatu = True: Exit Function
    End If

    LoadDefaultOutStatu = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckOutBalanceIsvalied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��Ժ���ʼ��
    '����:��Ժ������Ч,���سɹ�,���򷵻�False
    '����:���˺�
    '����:2015-01-04 14:15:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If Val(NVL(mrsInfo!��ҳID)) = 0 Or Val(NVL(mrsInfo!��Ժ)) <> 1 Then CheckOutBalanceIsvalied = True: Exit Function
    If Not gTy_System_Para.TY_Balance.bln��Ժ��׼���� Then CheckOutBalanceIsvalied = True: Exit Function
    If Not opt��;.Enabled Then
        MsgBox "��Ժ���˲������Ժ����,������û����;���ʵ�Ȩ��,���Բ��ܶԸò��˽���!", vbInformation, gstrSysName
        Exit Function
    End If
    CheckOutBalanceIsvalied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function CheckPatiBlacklist(ByVal lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鲡�˺�����
    '���:lng����ID-����ID
    '����:�޺��������������,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-04 14:30:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String
    '����������
    On Error GoTo errHandle
    strTmp = inBlackList(mrsInfo!����ID)
    If strTmp = "" Then CheckPatiBlacklist = True: Exit Function
    If MsgBox("����""" & mrsInfo!���� & """�����ⲡ�������С�" & vbCrLf & vbCrLf & "ԭ��" & vbCrLf & vbCrLf & "����" & strTmp & vbCrLf & vbCrLf & "Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    CheckPatiBlacklist = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckChargeAudit(ByVal lng����ID As Long, Optional blnSaveCheck As Boolean = False, Optional ByVal strTimes As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������˼��
    '����:���Ϸ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-04 15:04:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    'bytAuditing:0-�����,1-��鲢��ʾ,2-��鲢��ֹ
    Dim strSQL As String, rsTmp As ADODB.Recordset
    If gTy_System_Para.TY_Balance.bytAuditing = 0 Then CheckChargeAudit = True: Exit Function
    '�����ˣ��˳�
    If mblnNotify = True Then CheckChargeAudit = True: Exit Function
    If strTimes = "" Then
        strSQL = _
            "Select 1 From סԺ���ü�¼ A" & _
                " Where ���ʷ���=1 And ��¼״̬=0 And Nvl(ʵ�ս��,0)<>0 And ����ID=[1] And Not Exists" & _
                " (Select 1 From ҩƷ�շ���¼ C Where A.ID = C.����ID And Mod(C.��¼״̬, 3) = 1 And Nvl(C.ժҪ,'��һ')='�ܷ�' And instr( ',8,9,10,21,24,25,26,',','||C.����||',')>0) And Not Exists" & _
                " (Select 1 From ����ҽ������ B Where A.NO=B.NO And A.��¼����=B.��¼���� And A.ҽ�����=B.ҽ��ID And B.ִ��״̬ = 2) And Rownum=1"
    Else
        strSQL = _
            "Select 1 From סԺ���ü�¼ A" & _
                " Where ���ʷ���=1 And ��¼״̬=0 And Nvl(ʵ�ս��,0)<>0 And ����ID=[1] And Not Exists" & _
                " (Select 1 From ҩƷ�շ���¼ C Where A.ID = C.����ID And Mod(C.��¼״̬, 3) = 1 And Nvl(C.ժҪ,'��һ')='�ܷ�' And instr( ',8,9,10,21,24,25,26,',','||C.����||',')>0) And Not Exists" & _
                " (Select 1 From ����ҽ������ B Where A.NO=B.NO And A.��¼����=B.��¼���� And A.ҽ�����=B.ҽ��ID And B.ִ��״̬ = 2) And Rownum < 2 And a.��ҳID In (Select Column_Value From Table(f_str2list([2]))) "
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng����ID, strTimes)
    If rsTmp.RecordCount = 0 Then CheckChargeAudit = True: Exit Function
    Select Case gTy_System_Para.TY_Balance.bytAuditing
    Case 1
        If MsgBox("�ò��˻�����δ��˵ļ��ʷ��ã�Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    Case 2
        If blnSaveCheck Then
            If opt��Ժ.Value = True Then
                MsgBox "�ò��˻�����δ��˵ļ��ʷ���,���ܳ�Ժ���ʣ�", vbInformation, gstrSysName
                Exit Function
            Else
                If MsgBox("�ò��˻�����δ��˵ļ��ʷ��ã�Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            End If
        Else
            If opt��;.Enabled Then opt��;.Value = True 'ʹ����;����
        End If
    Case Else
    End Select
    CheckChargeAudit = True
End Function

Private Function AutoCalcChareFee(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Զ����ʼ���
    '���:lng����ID-����ID
    '     lng��ҳID-��ҳID
    '����:����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-04 15:13:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    'bytMCMode:1-����,2-סԺ����ģʽ,0-��ʾ��ҽ��
    If mYBInFor.bytMCMode = 1 Then AutoCalcChareFee = True: Exit Function
    If lng��ҳID = 0 Then AutoCalcChareFee = True: Exit Function
    
    '�Զ����㲡�˵Ĵ�λ���úͻ�������
    strSQL = "ZL1_AUTOCPTPATI(" & lng����ID & "," & lng��ҳID & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    AutoCalcChareFee = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function Load�����Ϣ(ByVal lng����ID As Long, ByVal byt���� As Byte) As Boolean
   
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ز�����ص������Ϣ
    '���:lng����ID=����ID
    '     byt����-0-����;1-����;2-סԺ
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-04 15:30:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    '��ȡ���˷������
    On Error GoTo errHandle
    If byt���� = 0 Then
        strSQL = "Select sum(Ԥ�����) As Ԥ�����,sum(�������) As ������� From ������� Where ����ID= [1] And ����=1"
    Else
        strSQL = "Select Ԥ�����,������� From ������� Where ����ID= [1] And ����=1 And ����= [2]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, byt����)
    If rsTemp.RecordCount <> 0 Then
        mPatiInfor.dblԤ����� = Format(Val(NVL(rsTemp!Ԥ�����)), "0.00")
        mPatiInfor.dbl������� = Format(Val(NVL(rsTemp!�������)), "0.00")
        mPatiInfor.dblʣ��ϼ� = Format(Val(NVL(rsTemp!Ԥ�����)) - Val(NVL(rsTemp!�������)), "0.00")
        If mEditType <> g_Ed_������� And mEditType <> g_Ed_סԺ���� Or ((mEditType = g_Ed_������� Or mEditType = g_Ed_סԺ����) And chkCancel.Value = 1) Then
            staThis.Panels(3).Text = "" & _
            "Ԥ��:" & Format(mPatiInfor.dblԤ�����, "0.00") & _
            "/����:" & Format(mPatiInfor.dbl�������, "0.00") & _
            "/ʣ��:" & Format(mPatiInfor.dblʣ��ϼ�, "0.00")
            txtԤ�����.Visible = False
            txtδ�����.Visible = False
            txtʣ����.Visible = False
            lblԤ�����.Visible = False
            lblδ�����.Visible = False
            lblʣ����.Visible = False
        Else
            txtԤ�����.Visible = True
            txtδ�����.Visible = True
            txtʣ����.Visible = True
            lblԤ�����.Visible = True
            lblδ�����.Visible = True
            lblʣ����.Visible = True
            txtԤ�����.Text = Format(mPatiInfor.dblԤ�����, "0.00")
            txtδ�����.Text = Format(mPatiInfor.dbl�������, "0.00")
            txtʣ����.Text = Format(mPatiInfor.dblʣ��ϼ�, "0.00")
        End If
    Else
        If mEditType <> g_Ed_������� And mEditType <> g_Ed_סԺ���� Then
            txtԤ�����.Visible = False
            txtδ�����.Visible = False
            txtʣ����.Visible = False
            lblԤ�����.Visible = False
            lblδ�����.Visible = False
            lblʣ����.Visible = False
        End If
    End If
    
    Load�����Ϣ = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function LoadӦ�տ���Ϣ(ByVal lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ز���Ӧ�տ���Ϣ
    '���:lng����ID-����ID
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-04 15:42:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim curDue As Currency
    
    On Error GoTo errHandle
    If InStr(mstrPrivs, ";Ӧ�տ����;") = 0 Then LoadӦ�տ���Ϣ = True: Exit Function
    curDue = GetPatientDue(lng����ID)
    If curDue = 0 Then LoadӦ�տ���Ϣ = True: Exit Function
    
    MsgBox mrsInfo!���� & ",Ӧ�տ����:" & Format(curDue, "0.00") & "Ԫ", vbInformation, gstrSysName
    staThis.Panels(2).Text = "����Ӧ�տ����:" & Format(curDue, "0.00") & "Ԫ"
    LoadӦ�տ���Ϣ = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function LoadסԺ״̬(ByVal lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����סԺ״̬
    '����:���˺�
    '����:2015-01-04 16:47:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    lbl״̬.Caption = GetPatiState(lng����ID)
    lbl���ʽ.Left = lbl״̬.Left + lbl״̬.Width + 60
    lbl���ʽ.Caption = "" & mrsInfo!ҽ�Ƹ��ʽ
    pic״̬.Width = lbl״̬.Width + lbl���ʽ.Width + 180
    If pic״̬.Width >= 2500 Then
        pic״̬.Width = 2500
    End If
    pic״̬.Visible = True
    
    LoadסԺ״̬ = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub ClearVsGrid()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������Ϣ
    '����:���˺�
    '����:2015-01-04 17:25:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsBlance
        .Rows = 2
        .Clear 1
    End With
    With vsDeposit
        .Rows = 2
        .Clear 1
    End With
    With vsFeeList
        .Rows = 2
        .Clear 1
    End With
    With vsDetailList
        .Rows = 2
        .Clear 1
    End With
End Sub
Private Sub ClearBalance()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������Ϣ
    '����:���˺�
    '����:2014-12-31 11:17:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long
    
    Call InitBalanceMoney  '���������Ϣ
    txtDate.Text = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    txtBalance(Idx_�������).Text = ""
    txtBalance(Idx_����˵��).Text = ""
    txtBalance(Idx_��Ԥ��).Text = ""
    txtBalance(Idx_ժҪ).Text = ""
    txtBalance(Idx_�ɿ�).Text = ""
    txtBalance(Idx_�Ҳ�).Text = "0.00"
    txtBalance(Idx_����δ��).Text = gstrDec
    txtBalance(Idx_����δ��).Tag = gstrDec
    lblʣ���Ը�.Caption = "0.00"
    lbl�Ը��ϼ�.Caption = "0.00"
End Sub

Private Sub ClearFeeList()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������Ϣ
    '����:���˺�
    '����:2015-01-04 17:29:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsFeeList
        .Redraw = False
        .Clear 1
        .Row = 1: .Col = .FixedCols
        .Redraw = True
    End With
    With vsDetailList
        .Redraw = False
        .Clear 1
        .Row = 1: .Col = .FixedCols
        .Redraw = True
    End With
End Sub

Private Sub ClearAdjustBalance()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������Ŀ�б�
    '����:���˺�
    '����:2015-01-04 17:31:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intRedraw As RedrawSettings
    Dim i As Long
    With mYBInFor
        .bln���ʽ��� = False
        .cur������� = 0
        .cur�����޶� = 0
        .cur����͸֧ = 0
    End With
    With vsBlance
        intRedraw = .Redraw
        .Redraw = flexRDNone
        .Rows = IIf(mlngBalanceRows < 2, 2, mlngBalanceRows)
        If mlngBalanceRows = 0 Then
            .Clear 1
        Else
            For i = 1 To .Rows - 1
                If .RowData(i) <> 110 Then .TextMatrix(i, .ColIndex("���㷽ʽ")) = ""
                .TextMatrix(i, .ColIndex("������")) = ""
                .TextMatrix(i, .ColIndex("�������")) = ""
                .TextMatrix(i, .ColIndex("��������")) = 2
                .TextMatrix(i, .ColIndex("����״̬")) = 0
                .TextMatrix(i, .ColIndex("�༭״̬")) = 1
                .RowData(i) = 110
            Next i
        End If
        .Redraw = intRedraw
    End With
End Sub
Private Sub ClearAdjustDeposit()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���Ԥ���б�
    '����:���˺�
    '����:2015-01-04 17:35:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intRedraw As RedrawSettings
    With vsDeposit
        intRedraw = .Redraw
        .Redraw = flexRDNone
        .Rows = 2
        .Clear 1
        .Redraw = intRedraw
    End With
End Sub

Private Function ShowBalance(Optional ByVal blnInputPatiAfterID As Boolean, _
    Optional ByRef strMessage As String, Optional blnNotClearPati As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������,��ʾ����Ҫ��������,����ʼ��������
    '���:blnInputPatiAfterID-�������ȷ��ʱ����
    '����:strMessage-������ʾ��Ϣ
    '     blnNotClearPati-true:��������ˣ�����Ա����ѡ������
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-04 16:54:10
    '˵����
    '   �ù��ܿ�������һ�����˽�����ɺ����,Ҳ�����ǵ�һ�������ڽ���ʱ��һ������;����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnUpload As Boolean, blnZero As Boolean
    Dim dtBeginDate As Date, dtEndDate As Date
    Dim str��ҳIds As String, i As Long
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim lngLast As Long, blnLastYb As Boolean
    Dim varTemp As Variant, blnFind As Boolean, intInsure As Integer, strInsureName As String
     
    On Error GoTo errHandle
        
    blnNotClearPati = False
    Call ClearFeeList   '������б�
    Call ClearAdjustBalance '��������б�
    Call ClearAdjustDeposit  '���Ԥ���б�
    If mrsInfo.State <> 1 Then Exit Function
    
    
    Screen.MousePointer = 11
    If blnInputPatiAfterID Then
        Call InitPatiBalanceVariableCon
    End If
    
    blnZero = mty_ModulePara.blnZero
    If mYBInFor.intInsure <> 0 And mYBInFor.bytMCMode <> 1 Then
        If opt��;.Value And MCPAR.��;������������ϴ����� Then blnUpload = True
    End If
    
    If blnInputPatiAfterID Then mobjBalanceCon.bytKind = 2
    
    If ReadBalanceData(mrsFeeList, blnUpload, blnInputPatiAfterID) = False Then Screen.MousePointer = 0: Exit Function
    If blnInputPatiAfterID Then Call LoadDefaultFilterCons
    
    If mbln�������� Then
        mobjBalanceCon.strTime = zlGetAllTims(cboPatiNums.GetNodesCheckedDatas(False))
        If mEditType = g_Ed_סԺ���� Then
            Call CheckPatiFromZyNumIsYB(Val(NVL(mrsInfo!����ID)), Val(Split(mobjBalanceCon.strTime, ",")(0)), intInsure, strInsureName)
             mYBInFor.intInsure = intInsure
        End If
    End If
    
    If mstrInputInNo <> "" Then
        varTemp = Split(mstrInputInNo, ",")
        blnFind = False
        For i = 0 To UBound(varTemp)
            If InStr("," & mobjBalanceAll.strAllTime & ",", "," & varTemp(i) & ",") > 0 Then
                blnFind = True: Exit For
            End If
        Next
        
        If blnFind = False Then
            mstrInputInNo = ""
            If mobjBalanceAll.strAllTime <> "" Then
                If MsgBox("����:" & mrsInfo!���� & "�ĵ�" & mstrInputInNo & "��סԺ�����Ѿ�����,������������סԺ����δ�ᣬ�Ƿ���������ã�", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Function
                
            Else
                MsgBox "����:" & mrsInfo!���� & "�ĵ�" & mstrInputInNo & "��סԺ�����Ѿ����壬�����¶�ȡ�ò��˵�δ�����!", vbInformation + vbDefaultButton1, gstrSysName
                Exit Function
            End If
            txtPatient.Text = "-" & mrsInfo!����ID
            '��ʾ���һ��δ����õ���Ϣ
            Call LoadPatientInfo(IDKIND.GetCurCard, False, , Val(Split(mobjBalanceAll.strAllTime, ",")(0)))
        End If
    End If
    
    If mstrInputInNo <> "" Then
    
        cboPatiNums.Text = "": blnFind = False

        For i = 1 To cboPatiNums.ListCount
            If InStr("," & mstrInputInNo & ",", "," & Val(Mid(cboPatiNums.Nodes.Item(i).Key, 2)) & ",") > 0 Then
                cboPatiNums.Nodes.Item(i).Checked = True
                cboPatiNums.Text = cboPatiNums.Text & "," & cboPatiNums.Nodes.Item(i).Text
                blnFind = True
            Else
                cboPatiNums.Nodes.Item(i).Checked = False
            End If
        Next i
        If cboPatiNums.Text <> "" Then cboPatiNums.Text = Mid(cboPatiNums.Text, 2)
        
        If blnFind = False Then
            'ȫѡδ�Ჿ��
            MsgBox "����:" & mrsInfo!���� & "�ĵ�" & mstrInputInNo & "��סԺ�����Ѿ����壬�����¶�ȡ�ò��˵�δ�����!", vbInformation + vbDefaultButton1, gstrSysName
            txtPatient.Text = "-" & Val(mrsInfo!����ID)
            mstrInputInNo = "": mblnReadByZYNo = False
            
            Call LoadPatientInfo(IDKIND.GetCurCard, False, , Split(mobjBalanceAll.strAllTime, ",")(0))
            mYBInFor.intInsure = Val(NVL(mrsInfo!����))
            
            For i = 1 To cboPatiNums.ListCount
                If mYBInFor.intInsure > 0 Then
                    If Split(mobjBalanceAll.strAllTime, ",")(0) = Val(Mid(cboPatiNums.Nodes.Item(i).Key, 2)) Then
                        cboPatiNums.Nodes.Item(i).Checked = True
                        cboPatiNums.Text = cboPatiNums.Text & "," & cboPatiNums.Nodes.Item(i).Text
                    Else
                        cboPatiNums.Nodes.Item(i).Checked = False
                    End If
                Else
                    cboPatiNums.Nodes.Item(i).Checked = True
                    cboPatiNums.Text = cboPatiNums.Text & "," & cboPatiNums.Nodes.Item(i).Text
                End If
            Next i
            If cboPatiNums.Text <> "" Then cboPatiNums.Text = Mid(cboPatiNums.Text, 2)
        Else
            mYBInFor.intInsure = Val(NVL(mrsInfo!����))
        End If
        mobjBalanceCon.strTime = zlGetAllTims(cboPatiNums.GetNodesCheckedDatas())
        mblnNotChange = True
        Call RecalcFeeTotalDate
        mblnNotChange = False
        If ReadBalanceData(mrsFeeList, blnUpload, blnInputPatiAfterID) = False Then Screen.MousePointer = 0: Exit Function
    End If
    
    If blnInputPatiAfterID And mrsFeeList.RecordCount = 0 And mstrInputInNo <> "" Then
        mstrInputInNo = ""
        mobjBalanceCon.strTime = mstrInputInNo
        If ReadBalanceData(mrsFeeList, blnUpload, blnInputPatiAfterID) = False Then Screen.MousePointer = 0: Exit Function
    End If
    
    If mrsFeeList Is Nothing Then Screen.MousePointer = 0: Exit Function
    If blnInputPatiAfterID And mrsFeeList.RecordCount = 0 And mEditType = g_Ed_������� Then
        mobjBalanceCon.bytKind = 1 'ȱʡֻȡ��ͨ���ã����û���ټ��ֻ���������������
        If ReadBalanceData(mrsFeeList, blnUpload, blnInputPatiAfterID) = False Then Screen.MousePointer = 0: Exit Function
        If mrsFeeList Is Nothing Then Screen.MousePointer = 0: Exit Function
        
        If mrsFeeList.RecordCount > 0 Then
            If MsgBox("�ò�����ͨ�����ѽ���,Ҫ�������ý��н�����?", vbInformation + vbYesNo, Me.Caption) = vbNo Then
                Set mrsFeeList = Nothing
                Screen.MousePointer = 0: Exit Function
            End If
        End If
    End If
    
    If mrsFeeList.RecordCount = 0 Then
        Set mrsFeeList = Nothing
        If blnInputPatiAfterID Then strMessage = "�ò���û����Ҫ���ʵķ��ã�"
        Screen.MousePointer = 0: Exit Function
    End If
    
    If blnInputPatiAfterID Then
         '����ȱʡ�Ĺ�������
        Call SetBalanceConInfor
        If mobjBalanceAll.strAllOwnerFeeType <> "" Then
            blnNotClearPati = True
            Call SetBalanceConInfor
            '��ȱʡ���Է���Ŀ
            If ReadBalanceData(mrsFeeList, blnUpload, blnInputPatiAfterID) = False Then Screen.MousePointer = 0: Exit Function
            If mrsFeeList Is Nothing Then Screen.MousePointer = 0: Exit Function
            If mrsFeeList.RecordCount = 0 Then Screen.MousePointer = 0: Exit Function
        End If
        
        '�������Ƿ����
        If CheckPatiIsVerfy(strMessage) = False Then Screen.MousePointer = 0: Exit Function
        '�����Ѫ��
        If CheckInputBlood = False Then Screen.MousePointer = 0: Exit Function
        
        If mobjBalanceCon.blnCurBalanceOwnerFee = False _
            And (mYBInFor.intInsure <> 0 And MCPAR.�������ú���ýӿ�) Or MCPAR.�������_�������� Then
            '----------------------------------------------------------------
            '��ȡסԺ���ڷ�Χ��ȱʡ��סԺʱ��
            If GetPatiHospitalzedDateRange(dtBeginDate, dtEndDate) = False Then Exit Function
            txtPatiBegin.Text = Format(dtBeginDate, txtPatiBegin.Format)
            txtPatiEnd.Text = Format(dtEndDate, txtPatiEnd.Format)
            txtPatiEnd.Tag = Format(dtEndDate, txtPatiEnd.Format)
            Call zlChangeDefaultTime
            mblnConsChange = True
            Call ClearListData
            Screen.MousePointer = 0
            mblnConsChange = False
            ShowBalance = True
            mblnInterUse = True
            Call ShowBalance(False)
            mblnInterUse = False
            mstrInputInNo = ""
            Exit Function
        End If
        mblnInterUse = True
        ShowBalance = ShowBalance(False)
        mblnInterUse = False
        Call ResetTime
        mstrInputInNo = ""
        'ShowBalance = True
        Exit Function
    End If
    
    '78317:ҽ������Ĭ��ֻ��ȡ���һ��סԺ������
    If mEditType <> g_Ed_������� And mobjBalanceCon.blnCurBalanceOwnerFee = False _
        And mYBInFor.intInsure <> 0 And (blnInputPatiAfterID Or mblnInterUse) And mstrInputInNo = "" Then
        
        lngLast = Val(Split(mobjBalanceAll.strAllTime & ",", ",")(0))
        Call CheckPatiFromZyNumIsYB(Val(NVL(mrsInfo!����ID)), lngLast, intInsure, strInsureName)
        If intInsure <> 0 Then
            If intInsure <> mYBInFor.intInsure Then Call InitInsurePara(Val(NVL(mrsInfo!����ID)), intInsure)
            mYBInFor.intInsure = intInsure
            
            If NVL(mrsInfo!��������) = "" Then
                txtPatient.ForeColor = vbRed
            Else
                txtPatient.ForeColor = zlDatabase.GetPatiColor(NVL(mrsInfo!��������))
            End If
            staThis.Panels(4).Text = strInsureName
            staThis.Panels(4).Visible = True
        Else
            mYBInFor.intInsure = 0
            txtPatient.ForeColor = Me.ForeColor
            staThis.Panels(4).Text = ""
            staThis.Panels(4).Visible = False
        End If
        
        '���һ�β���ҽ����Ժ,������ͨ���˴���
        mobjBalanceCon.strTime = lngLast
        For i = 1 To cboPatiNums.ListCount
            blnFind = InStr("," & mobjBalanceCon.strTime & ",", "," & Val(Mid(cboPatiNums.Nodes.Item(i).Key, 2)) & ",") > 0
            If Not blnFind And mYBInFor.intInsure <> 0 And MCPAR.�������סԺ���� Then blnFind = True
            cboPatiNums.Nodes.Item(i).Checked = blnFind
        Next i
        mobjBalanceCon.strTime = zlGetAllTims(cboPatiNums.GetNodesCheckedDatas())
        Call cboPatiNums.Refresh
        If ReadBalanceData(mrsFeeList, blnUpload, blnInputPatiAfterID) = False Then Screen.MousePointer = 0: Exit Function
    End If
    
    '���ط����б���Ϣ
    If LoadFeeList = False Then Screen.MousePointer = 0: Exit Function
    
    '���ؽ�����Ϣ
    str��ҳIds = IIf(mty_ModulePara.bln����ָ��Ԥ���� And mbln����תסԺ = False, _
        IIf(mobjBalanceCon.strTime = "", mobjBalanceAll.strAllTime, mobjBalanceCon.strTime), "")
    
    
    If LoadDepositList(Val(NVL(mrsInfo!����ID)), str��ҳIds) = False Then Screen.MousePointer = 0: Exit Function
                                
    '----------------------------------------------------------------
    '��ȡסԺ���ڷ�Χ��ȱʡ��סԺʱ��
    If GetPatiHospitalzedDateRange(dtBeginDate, dtEndDate) = False Then Exit Function
    txtPatiBegin.Text = Format(dtBeginDate, txtPatiBegin.Format)
    txtPatiEnd.Text = Format(dtEndDate, txtPatiEnd.Format)
    txtPatiEnd.Tag = Format(dtEndDate, txtPatiEnd.Format)
    Call zlChangeDefaultTime
    
    '----------------------------------------------------------------
    'ҽ��Ԥ����(��ͨ����Ҳ���ã��ڲ��д���ֱ�ӷ���true
    If InsureBudgeting(blnUpload) = False Then
        Set mrsInfo = New ADODB.Recordset
        Exit Function
    End If
    
    
    '���·���Ԥ����(bytOperationType-��������(0-������г�Ԥ��;1-��ȱʡʹ��Ԥ����;2-�����ʽ������Ԥ��(��ʱ���Ⱥ�����̯��;3-ȫ��)
    If mobjBalanceCon.blnCurBalanceOwnerFee Then
        Call RecalcDepositMoney(IIf(mty_ModulePara.bln�Է�ȱʡʹ��Ԥ��, 2, 0))
    Else
        Call RecalcDepositMoney(1)
    End If
    Call LoadIntendBalance
    Call LoadCurOwnerPayInfor
    Call SetDefaultPayType '����ȱʡ��֧����ʽ
    mblnNotChange = True
    txtBalance(Idx_��Ԥ��).Text = Format(mBalanceInfor.dbl��Ԥ���ϼ�, "0.00")
    mblnNotChange = False
    txtDate.Text = Format(zlDatabase.Currentdate, txtDate.Format)
    Screen.MousePointer = 0
    mblnConsChange = False
    ShowBalance = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub ResetTime()
    Dim dtDate As Date
    With mrsFeeList
        If .RecordCount <> 0 Then
            .MoveFirst
            If mty_ModulePara.int����ʱ�� = 0 Then
                dtDate = mrsFeeList!�Ǽ�ʱ��
            Else
                dtDate = mrsFeeList!ʱ��
            End If
             mobjBalanceAll.MinDate = dtDate: mobjBalanceAll.MaxDate = dtDate
        End If
        
        Do While Not .EOF
            '�Ƚ�ȡ�����Сֵ
            If mty_ModulePara.int����ʱ�� = 0 Then
                dtDate = mrsFeeList!�Ǽ�ʱ��
            Else
                dtDate = mrsFeeList!ʱ��
            End If
            If dtDate < mobjBalanceAll.MinDate Then mobjBalanceAll.MinDate = dtDate
            If dtDate > mobjBalanceAll.MaxDate Then mobjBalanceAll.MaxDate = dtDate
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
        '��ʾ����ʱ��
        mblnNotChange = True
        Call RecalcFeeTotalDate
        If Format(mobjBalanceAll.MinDate, txtBegin.Format) < Format(txtBegin.Text, txtBegin.Format) Then txtBegin.Text = Format(mobjBalanceAll.MinDate, txtBegin.Format)
        If Format(mobjBalanceAll.MaxDate, txtEnd.Format) > Format(txtEnd.Text, txtEnd.Format) Then txtEnd.Text = Format(mobjBalanceAll.MaxDate, txtEnd.Format)
        mblnNotChange = False
    End With
End Sub

Private Sub LoadIntendBalance(Optional ByVal dblSum As Double = 0, Optional objCard As Card)
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim str���㷽ʽ As String, lngCount As Long
    Dim dblBalanceSum As Double, blnSingle As Boolean
    Dim i As Long, j As Long
    Dim lngRow As Long, blnThirdSingle As Boolean, strErrMsg As String
    Dim dblAdd As Double
    Dim blnAdd As Boolean
    Dim dblTotal As Double, blnDo As Boolean
    Dim dblAlr As Double, strArray() As String, intArray As Integer
    On Error GoTo errHandle

    mstrBalanceLimit = ""
    If mstrForceNote <> "" Then
        mstrForceNote = Mid(mstrForceNote, 1, InStr(mstrForceNote, "ǿ������") + 4)
    End If
    If objCard Is Nothing Then
        With vsBlance
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("�����Ϣ")) <> "" Then
                    dblTotal = dblTotal + Val(.TextMatrix(i, .ColIndex("������")))
                    lngCount = lngCount + 1
                Else
                    dblAlr = dblAlr + Val(.TextMatrix(i, .ColIndex("������")))
                End If
            Next i
            For i = 1 To lngCount
                For j = 1 To .Rows - 1
                    If .TextMatrix(j, .ColIndex("�����Ϣ")) <> "" Then
                        .RemoveItem j
                        Exit For
                    End If
                Next j
            Next i
        End With
        
        mBalanceInfor.dbl�Ѹ��ϼ� = dblAlr
        mBalanceInfor.dblδ���ϼ� = RoundEx(mBalanceInfor.dbl��ǰ���� - dblAlr, 5)
    Else
        With vsBlance
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("�����Ϣ")) <> "" And Val(.TextMatrix(i, .ColIndex("�����ID"))) = objCard.�ӿ���� Then
                    dblTotal = RoundEx(dblTotal + Val(.TextMatrix(i, .ColIndex("������"))), 5)
                    lngCount = lngCount + 1
                Else
                    dblAlr = RoundEx(dblAlr + Val(.TextMatrix(i, .ColIndex("������"))), 5)
                End If
            Next i
        End With
    End If
    
    If dblSum = 0 Then
        dblBalanceSum = RoundEx(mBalanceInfor.dbl��Ԥ���ϼ� + dblAlr - mBalanceInfor.dbl��ǰ����, 2)
    Else
        dblBalanceSum = RoundEx(mBalanceInfor.dbl��Ԥ���ϼ� + dblAlr - mBalanceInfor.dbl��ǰ����, 2)
        If RoundEx(dblSum, 5) <= dblBalanceSum Then
            dblBalanceSum = dblSum
        Else
            If MsgBox("������˿������������˿���(" & dblBalanceSum & ")" & vbCrLf & "�Ƿ���������˿�����н���?", vbQuestion + vbYesNo, gstrSysName) <> vbYes Then Exit Sub
        End If
    End If
    
    dblBalanceSum = RoundEx(dblBalanceSum + -1 * Get�ѻ����, 5)
    If dblBalanceSum <= 0 Then Exit Sub
    If mrsDeposit Is Nothing Then Exit Sub
    If mrsDeposit.RecordCount = 0 Then
        If objCard Is Nothing Then
            Exit Sub
        Else
            MsgBox objCard.���� & "��֧�ֽӿ��˿�!", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    With vsDeposit
        For i = .Rows - 1 To 1 Step -1
            If dblBalanceSum > 0 Then
                mrsDeposit.Filter = "Ԥ��ID=" & Val(.TextMatrix(i, .ColIndex("Ԥ��ID")))
                
                If Val(NVL(mrsDeposit!�����ID)) = 0 Then
                    dblAdd = Val(.TextMatrix(i, .ColIndex("��Ԥ��")))
                    If dblBalanceSum > dblAdd Then
                        dblBalanceSum = RoundEx(dblBalanceSum - dblAdd, 2)
                    Else
                        dblBalanceSum = 0
                    End If
                    GoTo GoNext
                End If

                If mrsDeposit.RecordCount <> 0 Then
                    If Val(NVL(mrsDeposit!�����ID)) <> 0 Then
                        If Not objCard Is Nothing Then
                            If objCard.�ӿ���� <> Val(NVL(mrsDeposit!�����ID)) Then GoTo GoNext
                        End If
                        If Val(NVL(mrsDeposit!��������)) = 8 And Val(.TextMatrix(i, .ColIndex("��Ԥ��"))) <> 0 Then
                            strSQL = "Select �Ƿ�����,�Ƿ�ȫ��,��������,����,�Ƿ�ȱʡ���� From ҽ�ƿ���� Where ID= [1]"
                            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(NVL(mrsDeposit!�����ID)))
                            If Val(NVL(rsTmp!�Ƿ�����)) = 1 And mty_ModulePara.bln�����������˿���� And Val(NVL(rsTmp!�Ƿ�ȱʡ����)) = 1 Then
                                dblAdd = Val(.TextMatrix(i, .ColIndex("��Ԥ��")))
                                If dblBalanceSum > dblAdd Then
                                    dblBalanceSum = RoundEx(dblBalanceSum - dblAdd, 2)
                                Else
                                    dblBalanceSum = 0
                                End If
                                GoTo GoNext
                            End If
                            If mstrCardPara <> "" Then
                                strArray = Split(mstrCardPara, "|")
                                blnDo = True
                                For intArray = 0 To UBound(strArray)
                                    If Val(Split(strArray(intArray), ",")(0)) = Val(NVL(mrsDeposit!�����ID)) Then
                                        blnDo = False
                                        blnThirdSingle = Val(Split(strArray(intArray), ",")(1)) = 1
                                        Exit For
                                    End If
                                Next intArray
                            Else
                                blnDo = True
                            End If
                            
                            If blnDo And Val(NVL(mrsDeposit!ת�ʼ�����)) = 0 Then
                                '2-��������ˮ�ŷֱ�����˿�ӿ�"
                                blnThirdSingle = gobjSquare.objSquareCard.ZlGetParaConfig(Me, Val(NVL(mrsDeposit!�����ID)), False, 2, strErrMsg)
                                mstrCardPara = mstrCardPara & IIf(mstrCardPara = "", "", "|") & Val(NVL(mrsDeposit!�����ID)) & "," & IIf(blnThirdSingle, 1, 0)
                            End If
                            
                            dblAdd = Val(.TextMatrix(i, .ColIndex("��Ԥ��")))
                            blnAdd = True
                            With vsBlance
                                If blnThirdSingle And Val(NVL(mrsDeposit!ת�ʼ�����)) = 0 Then
                                    If Val(vsDeposit.TextMatrix(i, vsDeposit.ColIndex("�༭״̬"))) <> 0 Then GoTo GoNext
                                    
                                    If .TextMatrix(.Rows - 1, .ColIndex("���㷽ʽ")) <> "" Then
                                        .Rows = .Rows + 1
                                    End If
                                    
                                    '��������
                                    .RowData(.Rows - 1) = ""
                                    .TextMatrix(.Rows - 1, .ColIndex("����")) = 3
                                    .TextMatrix(.Rows - 1, .ColIndex("�����ID")) = Val(NVL(mrsDeposit!�����ID))
                                    .TextMatrix(.Rows - 1, .ColIndex("���ѿ�ID")) = 0
                                    .TextMatrix(.Rows - 1, .ColIndex("��������")) = Val(NVL(mrsDeposit!��������))
                                    .TextMatrix(.Rows - 1, .ColIndex("�༭״̬")) = 2   '0-��ֹɾ��;1-����༭���;2-����ɾ��
                                    .TextMatrix(.Rows - 1, .ColIndex("����״̬")) = 0  '�Ƿ��ѽ���:1-�ѽ���;0-δ����
                                    .TextMatrix(.Rows - 1, .ColIndex("�Ƿ�����")) = Val(NVL(rsTmp!�Ƿ�����))
                                    .TextMatrix(.Rows - 1, .ColIndex("�Ƿ�ȫ��")) = Val(NVL(rsTmp!�Ƿ�ȫ��))
                                    .TextMatrix(.Rows - 1, .ColIndex("У�Ա�־")) = 0
                                    .TextMatrix(.Rows - 1, .ColIndex("�Ƿ�ת��")) = Val(NVL(mrsDeposit!ת�ʼ�����))
                                    .TextMatrix(.Rows - 1, .ColIndex("�Ƿ�����")) = Val(NVL(rsTmp!��������))
                                    .TextMatrix(.Rows - 1, .ColIndex("���������")) = Trim(NVL(rsTmp!����))
                                    .TextMatrix(.Rows - 1, .ColIndex("���㷽ʽ")) = NVL(mrsDeposit!���㷽ʽ)
                                    
                                    If dblBalanceSum > dblAdd Then
                                        .TextMatrix(.Rows - 1, .ColIndex("������")) = Format(-1 * dblAdd, "0.00")
                                    Else
                                        .TextMatrix(.Rows - 1, .ColIndex("������")) = Format(-1 * dblBalanceSum, "0.00")
                                    End If
                                    .TextMatrix(.Rows - 1, .ColIndex("�������")) = ""
                                    .TextMatrix(.Rows - 1, .ColIndex("��ע")) = ""
                                    .TextMatrix(.Rows - 1, .ColIndex("������ˮ��")) = NVL(mrsDeposit!������ˮ��)
                                    .TextMatrix(.Rows - 1, .ColIndex("����˵��")) = NVL(mrsDeposit!����˵��)
                                    
                                    .TextMatrix(.Rows - 1, .ColIndex("����")) = IIf(Val(NVL(rsTmp!��������)) = 1, String(Len(NVL(mrsDeposit!����)), "*"), NVL(mrsDeposit!����))
                                    .TextMatrix(.Rows - 1, .ColIndex("�����Ϣ")) = NVL(mrsDeposit!Ԥ��ID)
                                    .Cell(flexcpData, .Rows - 1, .ColIndex("�����Ϣ")) = 1
                                    .Cell(flexcpData, .Rows - 1, .ColIndex("����")) = NVL(mrsDeposit!����)
                                Else
                                    For j = 1 To .Rows - 1
                                        If Val(.TextMatrix(j, .ColIndex("�����ID"))) = Val(NVL(mrsDeposit!�����ID)) And Val(.TextMatrix(j, .ColIndex("����״̬"))) = 1 Then
                                            GoTo GoNext
                                        End If
                                    Next j
                                    
                                    lngRow = 0
                                    For j = 1 To .Rows - 1
                                        If .TextMatrix(j, .ColIndex("���㷽ʽ")) = NVL(mrsDeposit!���㷽ʽ) Then lngRow = j
                                    Next j
                                    
                                    If lngRow = 0 Then
                                        If .TextMatrix(.Rows - 1, .ColIndex("���㷽ʽ")) <> "" Then
                                            .Rows = .Rows + 1
                                        End If
                                        
                                        '��������
                                        .RowData(.Rows - 1) = ""
                                        .TextMatrix(.Rows - 1, .ColIndex("����")) = 3
                                        .TextMatrix(.Rows - 1, .ColIndex("�����ID")) = Val(NVL(mrsDeposit!�����ID))
                                        .TextMatrix(.Rows - 1, .ColIndex("���ѿ�ID")) = 0
                                        .TextMatrix(.Rows - 1, .ColIndex("��������")) = Val(NVL(mrsDeposit!��������))
                                        .TextMatrix(.Rows - 1, .ColIndex("�༭״̬")) = 2   '0-��ֹɾ��;1-����༭���;2-����ɾ��
                                        .TextMatrix(.Rows - 1, .ColIndex("����״̬")) = 0  '�Ƿ��ѽ���:1-�ѽ���;0-δ����
                                        .TextMatrix(.Rows - 1, .ColIndex("�Ƿ�����")) = Val(NVL(rsTmp!�Ƿ�����))
                                        .TextMatrix(.Rows - 1, .ColIndex("�Ƿ�ȫ��")) = Val(NVL(rsTmp!�Ƿ�ȫ��))
                                        .TextMatrix(.Rows - 1, .ColIndex("У�Ա�־")) = 0
                                        .TextMatrix(.Rows - 1, .ColIndex("�Ƿ�ת��")) = Val(NVL(mrsDeposit!ת�ʼ�����))
                                        .TextMatrix(.Rows - 1, .ColIndex("�Ƿ�����")) = Val(NVL(rsTmp!��������))
                                        .TextMatrix(.Rows - 1, .ColIndex("���������")) = Trim(NVL(rsTmp!����))
                                        .TextMatrix(.Rows - 1, .ColIndex("���㷽ʽ")) = NVL(mrsDeposit!���㷽ʽ)
                                        If dblBalanceSum > dblAdd Then
                                            .TextMatrix(.Rows - 1, .ColIndex("������")) = Format(-1 * dblAdd, "0.00")
                                        Else
                                            .TextMatrix(.Rows - 1, .ColIndex("������")) = Format(-1 * dblBalanceSum, "0.00")
                                        End If
                                        .TextMatrix(.Rows - 1, .ColIndex("�������")) = ""
                                        .TextMatrix(.Rows - 1, .ColIndex("��ע")) = ""
                                        .TextMatrix(.Rows - 1, .ColIndex("������ˮ��")) = ""
                                        .TextMatrix(.Rows - 1, .ColIndex("����˵��")) = ""
                                        .TextMatrix(.Rows - 1, .ColIndex("����")) = IIf(Val(NVL(rsTmp!��������)) = 1, String(Len(NVL(mrsDeposit!����)), "*"), NVL(mrsDeposit!����))
                                        
                                        If dblBalanceSum > dblAdd Then
                                            .TextMatrix(.Rows - 1, .ColIndex("�����Ϣ")) = NVL(mrsDeposit!����) & "," & TruncStringEx(NVL(mrsDeposit!������ˮ��)) & "," & TruncStringEx(NVL(mrsDeposit!����˵��)) & "," & RoundEx(-1 * dblAdd, 2) & "," & NVL(mrsDeposit!Ԥ��ID)
                                        Else
                                            .TextMatrix(.Rows - 1, .ColIndex("�����Ϣ")) = NVL(mrsDeposit!����) & "," & TruncStringEx(NVL(mrsDeposit!������ˮ��)) & "," & TruncStringEx(NVL(mrsDeposit!����˵��)) & "," & RoundEx(-1 * dblBalanceSum, 2) & "," & NVL(mrsDeposit!Ԥ��ID)
                                        End If
                                        .Cell(flexcpData, .Rows - 1, .ColIndex("����")) = NVL(mrsDeposit!����)
                                        
                                    Else
                                        '���½���
                                        If dblBalanceSum > dblAdd Then
                                            .TextMatrix(lngRow, .ColIndex("������")) = Format(Val(.TextMatrix(lngRow, .ColIndex("������"))) - dblAdd, "0.00")
                                            .TextMatrix(lngRow, .ColIndex("�����Ϣ")) = .TextMatrix(lngRow, .ColIndex("�����Ϣ")) & "|" & NVL(mrsDeposit!����) & "," & TruncStringEx(NVL(mrsDeposit!������ˮ��)) & "," & TruncStringEx(NVL(mrsDeposit!����˵��)) & "," & RoundEx(-1 * dblAdd, 2) & "," & NVL(mrsDeposit!Ԥ��ID)
                                        Else
                                            .TextMatrix(lngRow, .ColIndex("������")) = Format(Val(.TextMatrix(lngRow, .ColIndex("������"))) - dblBalanceSum, "0.00")
                                            .TextMatrix(lngRow, .ColIndex("�����Ϣ")) = .TextMatrix(lngRow, .ColIndex("�����Ϣ")) & "|" & NVL(mrsDeposit!����) & "," & TruncStringEx(NVL(mrsDeposit!������ˮ��)) & "," & TruncStringEx(NVL(mrsDeposit!����˵��)) & "," & RoundEx(-1 * dblBalanceSum, 2) & "," & NVL(mrsDeposit!Ԥ��ID)
                                        End If
                                    End If
                                End If
                            End With
                            
                            If dblBalanceSum > dblAdd Then
                                mBalanceInfor.dbl�Ѹ��ϼ� = RoundEx(mBalanceInfor.dbl�Ѹ��ϼ� - RoundEx(dblAdd, 2), 5)
                                mBalanceInfor.dblδ���ϼ� = RoundEx(mBalanceInfor.dblδ���ϼ� + RoundEx(dblAdd, 2), 5)
                                dblBalanceSum = dblBalanceSum - dblAdd
                            Else
                                mBalanceInfor.dbl�Ѹ��ϼ� = RoundEx(mBalanceInfor.dbl�Ѹ��ϼ� - RoundEx(dblBalanceSum, 2), 5)
                                mBalanceInfor.dblδ���ϼ� = RoundEx(mBalanceInfor.dblδ���ϼ� + RoundEx(dblBalanceSum, 2), 5)
                                dblBalanceSum = 0
                            End If
                        End If
                    End If
                End If
            End If
GoNext:
        Next i
    End With
    
    mrsDeposit.Filter = ""
    With vsBlance
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("��������"))) = 8 And .TextMatrix(i, .ColIndex("�����Ϣ")) <> "" And Val(.Cell(flexcpData, i, .ColIndex("�����Ϣ"))) = 0 Then
                mstrBalanceLimit = mstrBalanceLimit & "|" & .TextMatrix(i, .ColIndex("�����ID")) & "," & .TextMatrix(i, .ColIndex("������"))
            End If
        Next i
        If mstrBalanceLimit <> "" Then mstrBalanceLimit = Mid(mstrBalanceLimit, 2)
    End With
    If blnAdd = False And Not objCard Is Nothing Then
        MsgBox "û�п����˿�Ľ��,����ʹ��" & objCard.���� & "�˿�", vbInformation, gstrSysName
        Exit Sub
    End If
    If Not objCard Is Nothing Then
        IDKindPaymentsType.IDKIND = 1
    End If
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function Get�ѻ����() As Double
    Dim dblMoney As Double
    Dim i As Long
    dblMoney = 0
    With vsBlance
        For i = 1 To .Rows - 1
            If Val(.RowData(i)) = 110 And Val(.TextMatrix(i, .ColIndex("������"))) <> 0 Then
                dblMoney = dblMoney + Val(.TextMatrix(i, .ColIndex("������")))
            End If
        Next i
    End With
    Get�ѻ���� = dblMoney
End Function

Private Function InsureBudgeting(ByVal blnOnlyUpload As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ҽ��Ԥ����
    '���: blnOnlyUpload-�Ƿ�ֻ�������ϴ�����
    '����:Ԥ��ɹ�(����ͨ����δ��ҽ���������),����true,���򷵻�False
    '����:���˺�
    '����:2015-01-06 16:48:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln����ʱ�� As Boolean, intInsure As Integer, lng����ID As Long, strҽ���� As String
    Dim strBalance As String, varData As Variant, varTemp As Variant
    Dim strNotBalance As String '�����ڵĽ��㷽ʽ
    Dim lngRow As Long, blnOk As Boolean
    Dim cur�����ʻ� As Currency, curͳ��֧�� As Currency
    Dim curMoney As Currency
    Dim rsDetail As ADODB.Recordset
    Dim i As Long, byt״̬ As Byte, bytEditSta As Byte
    On Error GoTo errHandle
    With vsBlance
        .Redraw = flexRDNone
        .Rows = IIf(mlngBalanceRows < 2, 2, mlngBalanceRows)
        If mlngBalanceRows = 0 Then
            .Clear 1
        Else
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("������")) = ""
                .TextMatrix(i, .ColIndex("�������")) = ""
                .TextMatrix(i, .ColIndex("��������")) = 2
                .TextMatrix(i, .ColIndex("����״̬")) = 0
                .TextMatrix(i, .ColIndex("�༭״̬")) = 1
                .RowData(i) = 110
            Next i
        End If
        .Redraw = flexRDBuffered
    End With
    
    txtBalance(Idx_���ν���).Enabled = InStr(mstrPrivs, ";��������;") > 0
    txtBalance(Idx_���ν���).Locked = InStr(mstrPrivs, ";��������;") = 0
    
    Call HideYBMoneyInfo '����ҽ��֧����Ϣ
 
    intInsure = mYBInFor.intInsure
    lng����ID = Val(NVL(mrsInfo!����ID))
    strҽ���� = "" & mrsInfo!ҽ����
    
    cmdOK.Enabled = True
    If mobjBalanceCon.blnCurBalanceOwnerFee Or intInsure = 0 Then
        '��ǰ���ڽ��Էѵ�,�򲻴���ҽ��
        Call SetOperationCtrl(0)
        InsureBudgeting = True: Exit Function     '�Ƚ��Էѷ��ã�������Ԥ����
    End If
     
    bln����ʱ�� = mty_ModulePara.int����ʱ�� = 1 '0-���Ǽ�ʱ��,1-������ʱ��
    'ҽ��Ԥ����
    If mEditType = g_Ed_������� Then
        With mobjBalanceCon
            Set rsDetail = GetMzBalance_Insure(intInsure, lng����ID, _
                .dtBeginDate, .dtEndDate, blnOnlyUpload, mblnDateMoved, mYBInFor.bytMCMode = 1, .bytKind, .strItem, .strDeptIDs, .strClass, .strChargeType, bln����ʱ��)
        End With
    Else
        With mobjBalanceCon
            Set rsDetail = GetZYBalance_Insure(intInsure, lng����ID, _
                .strTime, .dtBeginDate, .dtEndDate, blnOnlyUpload, mblnDateMoved, .strBaby, .strItem, .strDeptIDs, .strClass, .strChargeType, bln����ʱ��)
        End With
    End If
    
    mYBInFor.strBalance = ""
    'ҽ���ӿ�:���ظ��ֱ������
    If mYBInFor.bytMCMode = 1 Then
        If MCPAR.����Ԥ���� Then
            If rsDetail.RecordCount = 0 Then
                Screen.MousePointer = 0:
                MsgBox "��ȡҽ��Ԥ��������ʧ��!", vbInformation, gstrSysName
                Exit Function
            End If
        
            'strAdvance:
            '1.�շ�ʱ�����
            '2.�˷�ʱ��������������շ�ʱ������1,��ʾ�����շѵ���
            '3. ҽ�����ν���ʱ������2
            '4. ҽ�����ν��㷢�������˷�ʱ�����¶��ν��㣬����3
            '5��������ʴ���4
            Call SetCmdStatus(False)
            If Not gclsInsure.ClinicPreSwap(rsDetail, strBalance, intInsure, "4") Then
                Call SetCmdStatus(True)
                Screen.MousePointer = 0
                cmdYB.Enabled = True
                MsgBox "����ҽ��Ԥ����ʧ��!", vbInformation, gstrSysName
                Exit Function
            End If
            Call SetCmdStatus(True)
        End If
    Else
        Call SetCmdStatus(False)
        strBalance = gclsInsure.WipeoffMoney(rsDetail, lng����ID, strҽ����, "1", intInsure, "|" & IIf(opt��;.Value, 0, 1))
        Call SetCmdStatus(True)
    End If
    
    '��ʾ�������
    mYBInFor.cur������� = gclsInsure.SelfBalance(lng����ID, strҽ����, IIf(mYBInFor.bytMCMode = 1, 10, 40), _
        mYBInFor.cur����͸֧, intInsure)
    
    
    '���㷽ʽ;���;�Ƿ������޸�|...
    mYBInFor.strBalance = strBalance
    varData = Split(mYBInFor.strBalance, "|")
    
    '��ʾ����ͳ�ﱨ���ܶ�
    curͳ��֧�� = 0: cur�����ʻ� = 0
    strNotBalance = ""
    blnOk = True
    
    With vsBlance
        .Redraw = flexRDNone
        For i = 0 To UBound(varData)
            '���㷽ʽ;���;�Ƿ������޸�|...
            varTemp = Split(varData(i) & ";;;;", ";")
            mrs���㷽ʽ.Filter = "���� ='" & varTemp(0) & "'"
            curMoney = Val(varTemp(1))
            byt״̬ = 0: bytEditSta = IIf(Val(varTemp(2)) = 1, "1", "0")
            
            If .TextMatrix(.Rows - 1, .ColIndex("���㷽ʽ")) = "" Then
                lngRow = .Rows - 1
            Else
                .Rows = .Rows + 1
                lngRow = .Rows - 1
            End If
'            lngRow = i + 1
'            If lngRow > 1 Then .Rows = .Rows + 1
                        
            If mrs���㷽ʽ.EOF = False Then
                '1-�ֽ���㷽ʽ,2-������ҽ������,3-ҽ�������ʻ�,4-ҽ������ͳ��,5-���տ���,6-�����ۿ�,7-һ��ͨ����,8-���㿨����
                Select Case Val(NVL(mrs���㷽ʽ!����))
                Case 3  '3-ҽ�������ʻ�
                    cur�����ʻ� = cur�����ʻ� + curMoney
                    If mYBInFor.cur������� - curMoney < -1 * mYBInFor.cur����͸֧ Then
                        curMoney = 0
                        MsgBox "�����ʻ������δ����,������ҽ������!", vbInformation, Me.Caption
                        blnOk = False
                        Exit Function
                    End If
                    byt״̬ = 2
                Case 4  '4-ҽ������ͳ��
                    curͳ��֧�� = curͳ��֧�� + curMoney
                    byt״̬ = 2
                Case Else  '��ҽ����,��Ҫ����
                    strNotBalance = strNotBalance & "," & varTemp(0)
                End Select
                .TextMatrix(lngRow, .ColIndex("��������")) = Val(NVL(mrs���㷽ʽ!����))
            Else
                strNotBalance = strNotBalance & "," & varTemp(0)
            End If
            
            .TextMatrix(lngRow, .ColIndex("����")) = byt״̬
            .TextMatrix(lngRow, .ColIndex("�༭״̬")) = bytEditSta   '0-��ֹɾ��;1-����༭���;2-����ɾ��
            If bytEditSta <> 0 Then
                .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = vbBlue
            End If
            .TextMatrix(lngRow, .ColIndex("����״̬")) = 0  '�Ƿ��ѽ���:1-�ѽ���;0-δ����
            
            .TextMatrix(lngRow, .ColIndex("���㷽ʽ")) = varTemp(0)
            .TextMatrix(lngRow, .ColIndex("������")) = Format(curMoney, gstrDec)
            .Cell(flexcpData, lngRow, .ColIndex("������")) = Val(varTemp(1))
       
        Next
        If strNotBalance <> "" Then
            .Rows = IIf(mlngBalanceRows < 2, 2, mlngBalanceRows)
            If mlngBalanceRows = 0 Then
                .Clear 1
            Else
                For i = 1 To .Rows - 1
                    
                    .TextMatrix(i, .ColIndex("������")) = ""
                    .TextMatrix(i, .ColIndex("�������")) = ""
                    .TextMatrix(i, .ColIndex("��������")) = 2
                    .TextMatrix(i, .ColIndex("����״̬")) = 0
                    .TextMatrix(i, .ColIndex("�༭״̬")) = 1
                    .RowData(i) = 110
                Next i
            End If
            .Redraw = flexRDBuffered
            Screen.MousePointer = 0
            MsgBox "���ʳ��ϵı��ս��㷽ʽδ������ȫ,�ò��˻������±��ս��㷽ʽ���Ա�����" & _
            vbCrLf & strNotBalance & vbCrLf & vbCrLf & "�����Ե����û�����Ŀ\���㷽ʽ������ȥ������Щ���㷽ʽ��", vbInformation, gstrSysName
            Exit Function
        End If
        .Redraw = flexRDBuffered
    End With
    mYBInFor.cur����֧�� = cur�����ʻ�
    mYBInFor.curͳ��֧�� = curͳ��֧��
    
    mBalanceInfor.dblδ���ϼ� = RoundEx(mBalanceInfor.dbl��ǰ���� - (cur�����ʻ� + curͳ��֧��), 6)
    mBalanceInfor.dblҽ��֧���ϼ� = RoundEx(cur�����ʻ� + curͳ��֧��, 3)
    mBalanceInfor.dblԤ�����ܶ� = mBalanceInfor.dblҽ��֧���ϼ�
'    lbl�����ʻ�.Caption = "�ʻ����:" & Format(mYBInFor.cur�������, "0.00")
    staThis.Panels(5).Text = Format(mYBInFor.cur�������, "0.00")
    staThis.Panels(5).Visible = True
    lblҽ������.Caption = "ͳ��֧��:" & Format(curͳ��֧��, "0.00")
    
    lblҽ������.Visible = True
'    lblCurPaymentInfor.Visible = True
    txtBalance(Idx_���ν���).Enabled = False
    
    Call ShowYBMoneyOrInvioceFormatInfor
    
    'bytFun-0-����ǰ;1-ҽ����������;2-�ѱ����˽��ʵ�;
    Call SetOperationCtrl(1)
    '��ʾҽ��������Ϣ:bytFun-0-ҽ��Ԥ����Ϣ��ʾ
    Call ShowLedDisplayBank(0)
    Call LoadCurOwnerPayInfor  '����֧���ϼ�
    InsureBudgeting = True
    Exit Function
errHandle:
    vsBlance.Redraw = flexRDBuffered
     Screen.MousePointer = 0
    Call SetCmdStatus(True)
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub SetCmdStatus(blnStatus As Boolean)
    cmdRefresh.Enabled = blnStatus
    cmdCancel.Enabled = blnStatus
    cmdOK.Enabled = blnStatus
    cmdNext.Enabled = blnStatus
    cmdYBBalance.Enabled = blnStatus
End Sub

Private Sub ShowLedDisplayBank(ByVal bytFun As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:Led��Ϣ��ʾ
    '���:bytFun-0-ҽ��Ԥ����Ϣ��ʾ;1-��ʾ������Ϣ
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-07 13:55:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dtTmpDate As Date, dblTemp As Double, strDepositName As String
    If Not gblnLED Then Exit Sub
    
    On Error GoTo errHandle
    
    If Not (mEditType = g_Ed_������� Or mEditType = g_Ed_���½��� Or mEditType = g_Ed_סԺ����) Then Exit Sub
    
    strDepositName = "סԺԤ����"
    If mEditType = g_Ed_������� Or mblnCurMzBalanceNo Then strDepositName = "����Ԥ��"
    Select Case bytFun
    Case 0 'ҽ��Ԥ����Ϣ��ʾ
        zl9LedVoice.DisplayBank "ҽ������:", _
            "�ʻ����" & Format(mYBInFor.cur�������, "0.00"), _
            "�ʻ�֧��" & Format(mYBInFor.cur����֧��, "0.00"), _
            "ͳ��֧��" & Format(mYBInFor.curͳ��֧��, "0.00")
    Case 1 '��ʾ������Ϣ
        zl9LedVoice.DisplayBank _
            "�ܷ���" & Format(mBalanceInfor.dbl��ǰ����, "0.00"), _
             strDepositName & Format(mPatiInfor.dblԤ�����, "0.00"), _
            "��Ԥ��" & Format(mBalanceInfor.dbl��Ԥ���ϼ�, "0.00"), _
            IIf(mBalanceInfor.dbl����δ�� < 0, "�Ҳ�", "Ӧ��") & Format(Abs(mBalanceInfor.dbl����δ��), "0.00")
    End Select
    
    '�ӳ�ʱ��
    dtTmpDate = Time
    Do While Time < DateAdd("s", 4, dtTmpDate)
    Loop
    
    Exit Sub
errHandle:
    
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub SetOperationCtrl(ByVal bytFun As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ò����ؼ����������
    '���:bytFun-0-����ǰ;1-ҽ����������;2-�ѱ����˽��ʵ�
    '            3-δ������������
    '����:���˺�
    '����:2015-01-07 11:21:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnEnabled As Boolean
    Dim lngColor As Long, objCard As Card
    Dim blnTemp As Boolean
    
    If mEditType = g_Ed_���ݲ鿴 Then cmdCancel.Visible = False: Exit Sub
    
    If mobjBalanceCon.blnCurBalanceOwnerFee Then
        cmdYBBalance.Visible = False
    Else
        cmdYBBalance.Visible = bytFun = 1 Or mYBInFor.intInsure <> 0 And mBalanceInfor.blnSaveBill = False
    End If
    
    If mYBInFor.bytMCMode = 1 Then
        cmdYBBalance.Enabled = True ' MCPAR.����Ԥ����
    Else
        cmdYBBalance.Enabled = mYBInFor.strBalance <> ""
    End If
    
    If cmdYBBalance.Visible Then
        cmdYBBalance.Left = cmdOK.Left
    End If
    cmdCancel.Visible = bytFun <> 2 And chkCancel.Value = 0
    Call SetNextBalanceCmdVisible   '�����������ʰ�ť
  
    cmdOK.Visible = bytFun <> 1 And cmdYBBalance.Visible = False
    cmdCancel.Visible = bytFun <> 2 And chkCancel.Value = 0
    
    
    
    cmdYB.Enabled = True
    txtBalance(Idx_�Ҳ�).Locked = True
    txtBalance(Idx_����δ��).Enabled = False
    'IDKindPaymentsType.Enabled = Not mPatiInfor.bln�������� '������ѡ���������㷽ʽ
    IDKindPaymentsType.Locked = mPatiInfor.bln��������
    Select Case bytFun
    Case 0  '����ǰ
        txtBalance(Idx_����˵��).Enabled = True
         
        txtBalance(Idx_��Ԥ��).Enabled = mPatiInfor.dblʵ����� <> 0
        txtBalance(Idx_�ɿ�).Enabled = True
        txtBalance(Idx_�Ҳ�).Enabled = True
        txtBalance(Idx_ժҪ).Enabled = True
        txtBalance(Idx_�������).Enabled = True
        txtBalance(Idx_����˵��).BackColor = &H80000005
        txtBalance(Idx_���ν���).Enabled = InStr(mstrPrivs, ";��������;") > 0
        txtBalance(Idx_���ν���).Locked = InStr(mstrPrivs, ";��������;") = 0
        cboPatiNums.Enabled = InStr(mstrPrivs, ";��������;") > 0
        cboӤ��.Enabled = InStr(mstrPrivs, ";��������;") > 0
        cboDept.Enabled = InStr(mstrPrivs, ";��������;") > 0
        cboFeeType.Enabled = InStr(mstrPrivs, ";��������;") > 0
        cboFeeItem.Enabled = InStr(mstrPrivs, ";��������;") > 0
        cboDiag.Enabled = InStr(mstrPrivs, ";��������;") > 0
        cboChargeType.Enabled = InStr(mstrPrivs, ";��������;") > 0
        
        txtBalance(Idx_���ν���).BackColor = &H80000005
        If Not mBalanceInfor.blnԤ����֤ Then
            txtBalance(Idx_��Ԥ��).BackColor = IIf(txtBalance(Idx_��Ԥ��).Enabled, &H80000005, &H8000000F)
        End If
        txtBalance(Idx_�ɿ�).BackColor = &H80000005
        txtBalance(Idx_�Ҳ�).BackColor = &H80000005
        txtBalance(Idx_ժҪ).BackColor = &H80000005
        txtBalance(Idx_�������).BackColor = &H80000005
        txtBalance(Idx_����δ��).Enabled = False
        txtBalance(Idx_����δ��).BackColor = &H8000000F
        IDKindPaymentsType.Enabled = True
        IDKind�Ҳ�.Enabled = True
        
        cmdRefresh.Visible = True
        txtPatient.Locked = False
        IDKIND.Enabled = True
        picPati.Enabled = True
        cmdDelBalance.Visible = False
    Case 1, 3 'ҽ���������� ��δ���ù�������ʱ
        If bytFun = 3 Then txtBalance(Idx_��Ԥ��).Text = "0.00"
        
        txtBalance(Idx_����˵��).Enabled = bytFun <> 3 ' True
        txtBalance(Idx_����δ��).Enabled = False
        txtBalance(Idx_���ν���).Enabled = False
        
        txtBalance(Idx_��Ԥ��).Enabled = False
        txtBalance(Idx_�ɿ�).Enabled = False
        txtBalance(Idx_�Ҳ�).Enabled = False
        txtBalance(Idx_ժҪ).Enabled = False
        txtBalance(Idx_�������).Enabled = False
        
        txtBalance(Idx_����˵��).BackColor = IIf(bytFun <> 3, &H80000005, &H8000000F)
        txtBalance(Idx_���ν���).BackColor = &H8000000F
        txtBalance(Idx_����δ��).BackColor = &H8000000F

        txtBalance(Idx_��Ԥ��).BackColor = &H8000000F
        txtBalance(Idx_�ɿ�).BackColor = &H8000000F
        txtBalance(Idx_�Ҳ�).BackColor = &H8000000F
        txtBalance(Idx_ժҪ).BackColor = &H8000000F
        txtBalance(Idx_�������).BackColor = &H8000000F
        txtBalance(Idx_���ν���).BackColor = IIf(txtBalance(Idx_���ν���).Enabled, &H80000005, &H8000000F)
        txtBalance(Idx_����δ��).BackColor = IIf(txtBalance(Idx_����δ��).Enabled, &H80000005, &H8000000F)
        IDKindPaymentsType.Enabled = False
        IDKind�Ҳ�.Enabled = False
        cmdDelBalance.Visible = False
    Case Else   '�ѱ����˽��ʵ�
        blnEnabled = mEditType <> g_Ed_ȡ������
        lngColor = IIf(blnEnabled, &H80000005, &H8000000F)
        txtBalance(Idx_����˵��).Enabled = False
        txtBalance(Idx_���ν���).Enabled = False
        txtBalance(Idx_����δ��).Enabled = False
        txtBalance(Idx_��Ԥ��).Enabled = mPatiInfor.dblʵ����� <> 0 And blnEnabled
        txtBalance(Idx_�ɿ�).Enabled = blnEnabled
        txtBalance(Idx_�Ҳ�).Enabled = blnEnabled
        txtBalance(Idx_ժҪ).Enabled = blnEnabled
        txtBalance(Idx_�������).Enabled = blnEnabled
        
        txtBalance(Idx_����˵��).BackColor = lngColor
        txtBalance(Idx_���ν���).BackColor = IIf(txtBalance(Idx_���ν���).Enabled, &H80000005, &H8000000F)
        txtBalance(Idx_����δ��).BackColor = IIf(txtBalance(Idx_����δ��).Enabled, &H80000005, &H8000000F)

        If mBalanceInfor.blnԤ����֤ Then
            txtBalance(Idx_��Ԥ��).BackColor = IIf(txtBalance(Idx_��Ԥ��).Enabled, &HE0E0E0, &H8000000F)
        Else
            txtBalance(Idx_��Ԥ��).BackColor = IIf(txtBalance(Idx_��Ԥ��).Enabled, &H80000005, &H8000000F)
        End If
        
        txtBalance(Idx_�ɿ�).BackColor = lngColor
        txtBalance(Idx_�Ҳ�).BackColor = lngColor
        txtBalance(Idx_ժҪ).BackColor = lngColor
        txtBalance(Idx_�������).BackColor = lngColor
        
        IDKindPaymentsType.Enabled = blnEnabled
        IDKind�Ҳ�.Enabled = blnEnabled
        
        cmdRefresh.Visible = False
        txtPatient.Locked = True
        IDKIND.Enabled = False
        picPati.Enabled = False
        cmdYBBalance.Visible = False
        cmdOK.Visible = True
        cmdOK.Enabled = True
        cmdOK.Visible = True
        
        cmdDelBalance.Visible = chkCancel.Value = 0
        
        cmdDelBalance.Left = cmdCancel.Left
        cmdDelBalance.Top = cmdCancel.Top
            
        cmdCancel.Visible = IIf(mEditType = g_Ed_�������� Or mEditType = g_Ed_���½��� Or mEditType = g_Ed_ȡ������, True, False)
        If cmdCancel.Visible Or cmdDelBalance.Visible Then
            cmdOK.Left = cmdCancel.Left - cmdOK.Width - 50
        Else
            cmdOK.Left = picBalanceBack.ScaleWidth - cmdOK.Width - 100
        End If
        
        cmdNext.Left = cmdOK.Left - cmdNext.Width - 50
    End Select
    cboӤ��.BackColor = IIf(cboӤ��.Enabled, &H80000005, &H8000000F)
    cboPatiNums.BackColor = IIf(cboPatiNums.Enabled, &H80000005, &H8000000F)
    cboFeeType.BackColor = IIf(cboFeeType.Enabled, &H80000005, &H8000000F)
    cboFeeItem.BackColor = IIf(cboFeeItem.Enabled, &H80000005, &H8000000F)
    cboDept.BackColor = IIf(cboDept.Enabled, &H80000005, &H8000000F)
    cboChargeType.BackColor = IIf(cboChargeType.Enabled, &H80000005, &H8000000F)
    cboDiag.BackColor = IIf(cboDiag.Enabled, &H80000005, &H8000000F)
    txtBalance(Idx_���ν���).BackColor = IIf(txtBalance(Idx_���ν���).Enabled, &H80000005, &H8000000F)
       
    
End Sub
Private Sub SetNextBalanceCmdVisible()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����������ʰ�ť
    '����:���˺�
    '����:2015-02-26 16:09:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnTemp As Boolean, objCard As Card
    Dim blnHave As Boolean, i As Long
    
    On Error GoTo errHandle
    
    If mty_ModulePara.byt�ɿ�������� <> 2 Then
        cmdNext.Visible = False: Exit Sub
    End If
    blnHave = False
    If Not mrsFeeList Is Nothing Then
        If mrsFeeList.State = 1 Then
            blnHave = mrsFeeList.RecordCount <> 0
        End If
    End If
    '��ͨ�շѻ�ҽ���Ѿ�����
    blnTemp = (mEditType = g_Ed_������� Or mEditType = g_Ed_סԺ����) And chkCancel.Value = 0
    blnTemp = blnTemp And mPatiInfor.lng����ID <> 0
    blnTemp = blnTemp And (mYBInFor.intInsure = 0 Or mobjBalanceCon.blnCurBalanceOwnerFee Or mYBInFor.intInsure <> 0 And Not cmdYBBalance.Visible)
    blnTemp = blnTemp And Val(txtBalance(Idx_�ɿ�).Text) = 0
    blnTemp = blnTemp And blnHave
    
    If blnTemp Then
        '������û��Ҫ�ۼ�
        Set objCard = IDKindPaymentsType.GetCurCard
        If Not objCard Is Nothing Then
            blnTemp = blnTemp And objCard.�ӿ���� <= 0 And Not objCard.���ѿ�
        Else
            blnTemp = False
        End If
    End If
    
    For i = 1 To vsBlance.Rows - 1
        If Val(vsBlance.TextMatrix(i, vsBlance.ColIndex("������"))) <> 0 Then
            blnTemp = False
        End If
    Next i
    
    cmdNext.Visible = blnTemp
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
 
Private Function LoadDefaultFilterCons() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ȱʡ�Ĺ�������
    '����:���˺�
    '����:2015-01-05 14:07:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dtDate As Date, cllOwnerFeeType As Collection, cllBalanceFeeType As Collection
    Dim blnCheck As Boolean, i As Long, bln��� As Boolean, bln��ͨ As Boolean
    Dim int��ҳID As Integer, strSQL As String, rsTmp As ADODB.Recordset
    Dim blnAll As Boolean, rsAllTime As ADODB.Recordset
    Dim objNode As Node, intInsure As Integer, strInsureName As String
    Dim varTemp As Variant
     
     
    
    On Error GoTo errHandle
    
    '������δ���ʵ����з�Χ����
    Set mobjBalanceAll = New clsBalanceAllCon
    
    With mobjBalanceAll
        .MinDate = #1/1/1900#: .MaxDate = #1/1/1900#
    End With
    Call InitCombox_Cons
    Set cllOwnerFeeType = New Collection
    Set cllBalanceFeeType = New Collection
    
    With mrsFeeList
        If .RecordCount <> 0 Then
            .MoveFirst
            If mty_ModulePara.int����ʱ�� = 0 Then
                dtDate = mrsFeeList!�Ǽ�ʱ��
            Else
                dtDate = mrsFeeList!ʱ��
            End If
             mobjBalanceAll.MinDate = dtDate: mobjBalanceAll.MaxDate = dtDate
        End If
        
        Do While Not .EOF
            If mEditType <> g_Ed_������� Then
                If InStr(mobjBalanceAll.strAllTime & ",", "," & Val(NVL(!��ҳID)) & ",") = 0 And Val(NVL(!��ҳID)) <> 0 Then
                    mobjBalanceAll.strAllTime = mobjBalanceAll.strAllTime & "," & Val(NVL(!��ҳID))
                End If
            Else
                If InStr(mobjBalanceAll.strAllTime & ",", "," & Val(NVL(!��ҳID)) & ",") = 0 Then
                    mobjBalanceAll.strAllTime = mobjBalanceAll.strAllTime & "," & Val(NVL(!��ҳID))
                End If
            End If
            
            If Val(NVL(mrsFeeList!��������ID)) <> 0 Then
                If InStr(mobjBalanceAll.strAllDeptIDs & ",", "," & Val(NVL(!��������ID)) & ",") = 0 Then
                    mobjBalanceAll.strAllDeptIDs = mobjBalanceAll.strAllDeptIDs & "," & mrsFeeList!��������ID
                    cboDept.AddItem Val(NVL(!��������ID)), NVL(!����, "δ֪"), False, True
                End If
            End If
            
            If Trim(NVL(!��Ŀ, "")) <> "" Then
                If InStr(mobjBalanceAll.strAllItem & ",", ",'" & !��Ŀ & "',") = 0 Then
                     mobjBalanceAll.strAllItem = mobjBalanceAll.strAllItem & ",'" & !��Ŀ & "'"
                    cboFeeItem.AddItem "'" & mrsFeeList!��Ŀ & "'", !��Ŀ, False, True
                End If
            End If
            
            If Trim(NVL(!���, "")) <> "" Then
                If InStr(mobjBalanceAll.strAllDiag & ",", ",'" & !��� & "',") = 0 Then
                     mobjBalanceAll.strAllDiag = mobjBalanceAll.strAllDiag & ",'" & !��� & "'"
                    cboDiag.AddItem !���
                End If
            End If
            
            If Trim(NVL(!�շ����)) <> "" Then  '34260
                If InStr("," & mobjBalanceAll.strAllChargeType & ",", ",'" & !�շ���� & "',") = 0 Then
                    mobjBalanceAll.strAllChargeType = mobjBalanceAll.strAllChargeType & ",'" & !�շ���� & "'"
                    If InStr(1, "," & mty_ModulePara.strOwnerPayFeeType & ",", "," & !�շ���� & ",") > 0 Then
                        If InStr("," & mobjBalanceAll.strAllOwnerFeeType & ",", ",'" & !�շ���� & "',") = 0 Then
                            mobjBalanceAll.strAllOwnerFeeType = mobjBalanceAll.strAllOwnerFeeType & ",'" & !�շ���� & "'"
                        End If
                        cllOwnerFeeType.Add Array("'" & !�շ���� & "'", NVL(!�շ������, "δ֪"))
                    Else
                        cllBalanceFeeType.Add Array("'" & !�շ���� & "'", NVL(!�շ������, "δ֪"))
                    End If
                End If
             
            End If
            '���Ϊ��,ָû�����÷�������
            If InStr("," & mobjBalanceAll.strAllClass & ",", ",'" & NVL(!����, "��") & "',") = 0 Then
                mobjBalanceAll.strAllClass = mobjBalanceAll.strAllClass & ",'" & NVL(!����, "��") & "'"
                cboFeeType.AddItem "'" & NVL(!����, "δ֪") & "'", NVL(!����, "δ֪"), False, True
            End If
               
            If InStr("," & mobjBalanceAll.strAllBabys & ",", "," & Val(NVL(!Ӥ����)) & ",") = 0 And Val(NVL(!Ӥ����)) <> 0 Then
                mobjBalanceAll.strAllBabys = mobjBalanceAll.strAllBabys & "," & Val(NVL(!Ӥ����)) & ""
                cboӤ��.AddItem Val(NVL(!Ӥ����)), "��" & Val(NVL(!Ӥ����)) & "��Ӥ��", False, True
            End If
            
            
            '�Ƚ�ȡ�����Сֵ
            If mty_ModulePara.int����ʱ�� = 0 Then
                dtDate = mrsFeeList!�Ǽ�ʱ��
            Else
                dtDate = mrsFeeList!ʱ��
            End If
            If dtDate < mobjBalanceAll.MinDate Then mobjBalanceAll.MinDate = dtDate
            If dtDate > mobjBalanceAll.MaxDate Then mobjBalanceAll.MaxDate = dtDate
            If mEditType = g_Ed_������� Then
                If Val(NVL(mrsFeeList!�����־)) = 4 Then
                    bln��� = True
                Else
                    bln��ͨ = True
                End If
            End If
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
    End With
    
    
    
    If mEditType = g_Ed_������� Then
        chkPatiType(0).Enabled = True
        chkPatiType(1).Enabled = True
        If bln��ͨ Then chkPatiType(0).Value = 1
        If bln��� Then chkPatiType(1).Value = 1
        If Not bln��ͨ And Not bln��� Then
            chkPatiType(0).Value = 1
            chkPatiType(1).Value = 1
        End If
        If Not bln��ͨ Then
            chkPatiType(0).Value = 0
            chkPatiType(0).Enabled = False
            chkPatiType(1).Enabled = False
        End If
        If Not bln��� Then
            chkPatiType(1).Value = 0
            chkPatiType(0).Enabled = False
            chkPatiType(1).Enabled = False
        End If
    End If
    
    lblFeeStyle.ForeColor = &H80000008
    cboChargeType.ForeColor = &H80000008
    cboChargeType.Clear
    If cllOwnerFeeType.Count <> 0 Then
        Set objNode = cboChargeType.Nodes.Add(, , "RootOwner", "�����Է����")
        objNode.ForeColor = vbRed
        objNode.Bold = True
        objNode.Expanded = True
        objNode.Tag = "'RootOwner'"
        objNode.Checked = True
        cboChargeType.Text = "�����Է����"
        lblFeeStyle.ForeColor = vbRed
        cboChargeType.ForeColor = vbRed
    End If
    
    For i = 1 To cllOwnerFeeType.Count
        Set objNode = cboChargeType.Nodes.Add("RootOwner", 4, cllOwnerFeeType(i)(0), cllOwnerFeeType(i)(1))
        objNode.ForeColor = vbRed
        objNode.Expanded = True
        objNode.Tag = cllOwnerFeeType(i)(0)
        objNode.Checked = True
    Next
    
    
    blnCheck = cllOwnerFeeType.Count = 0
    Set objNode = cboChargeType.Nodes.Add(, , "RootBalance", "�����շ����")
    
   objNode.Bold = True
   objNode.Expanded = True
   objNode.Tag = ""
   objNode.Checked = blnCheck
   objNode.Tag = "'RootBalance'"
   If blnCheck Then cboChargeType.Text = objNode.Text

    For i = 1 To cllBalanceFeeType.Count
        Set objNode = cboChargeType.Nodes.Add("RootBalance", 4, cllBalanceFeeType(i)(0), cllBalanceFeeType(i)(1))
        objNode.Expanded = True
        objNode.Tag = cllBalanceFeeType(i)(0)
        objNode.Checked = blnCheck
    Next
    
    If LoadDataPatiNumsToComBox(Val(NVL(mrsInfo!����ID)), Mid(mobjBalanceAll.strAllTime, 2), blnAll, rsAllTime, intInsure, strInsureName) = False Then Exit Function
    Set mobjBalanceAll.rsAllTime = rsAllTime

    With mobjBalanceAll
        .strAllTime = Mid(.strAllTime, 2)
        .strAllItem = Mid(.strAllItem, 2)
        .strAllDiag = Mid(.strAllDiag, 2)
        .strAllDeptIDs = Mid(.strAllDeptIDs, 2)
        .strAllChargeType = Mid(.strAllChargeType, 2)
        .strAllOwnerFeeType = Mid(.strAllOwnerFeeType, 2)
        .strAllClass = Mid(.strAllClass, 2)
        '��ʾ����ʱ��
        mblnNotChange = True
        txtBegin.Text = Format(.MinDate, txtBegin.Format)
        txtEnd.Text = Format(.MaxDate, txtEnd.Format)
        mblnNotChange = False
    End With
    Call SetPatiConsControlVisible
    LoadDefaultFilterCons = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function LoadDataPatiNumsToComBox(ByVal lng����ID As Long, ByVal str��ҳIds As String, ByRef blnAllSel As Boolean, _
    ByRef rsTimeAll As ADODB.Recordset, ByRef intInsure As Integer, Optional strInsureName As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����סԺ�������������б��
    '���: str��ҳIDs-����סԺ����,�ö��ŷָ�
    '����:blnAllSel-��ǰ�Ƿ�ѡ��������סԺ����
    '     intInsure-���ص�һ��ѡ���ҽ�����
    '     strInsureName-���ص�һ��ѡ���ҽ������
    '����:���سɹ�����true,���򷵻�False
    '����:���˺�
    '����:2017-11-13 11:23:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varTemp As Variant, int��ҳID As Long, strTag As String
    Dim i As Long, intInsure1 As Integer, strInsureName1 As String
    
    On Error GoTo errHandle
    
    cboPatiNums.Clear
    If mEditType <> g_Ed_סԺ���� Then
        cboPatiNums.AddItem "R", "��������", True, True, True, , "0"
        varTemp = Split(str��ҳIds, ",")
        blnAllSel = True
        For i = 0 To UBound(varTemp)
            If Val(varTemp(i)) = 0 Then
                cboPatiNums.AddItem Val(varTemp(i)), "��ͨ����", False, True
            Else
                cboPatiNums.AddItem Val(varTemp(i)), "��" & Val(varTemp(i)) & "������", False, True
            End If
        Next
        Call cboPatiNums.Refresh
        Set rsTimeAll = Nothing
        LoadDataPatiNumsToComBox = True
        Exit Function
    End If
    
    cboPatiNums.AddItem "R", "����סԺ", True, True, True, , "0"
    '��ȡ��ǰδ��סԺ�����漰��ҽ�����ݼ�
    Call mobjBalanceAll.zlGetTimeRecordFromTimeString(lng����ID, str��ҳIds, rsTimeAll)

    '����סԺ�����ı���
    Dim blnSelect As Boolean
    With rsTimeAll
        intInsure = 0
        If .RecordCount <> 0 Then
            .MoveFirst:  intInsure = Val(NVL(!����)): strInsureName = NVL(!��������)
        End If
        
        i = 1: blnAllSel = True
        Do While Not .EOF
            '�Էѵģ���ȱʡȫѡ,���һ��סԺΪҽ���ģ����Ƚ�ҽ����
            
            blnSelect = mobjBalanceAll.strAllOwnerFeeType <> "" Or (intInsure <> 0 And i = 1) Or intInsure = 0
            If Not blnSelect And intInsure <> 0 And MCPAR.�������סԺ���� Then blnSelect = True
            
            If blnAllSel And Not blnSelect Then blnAllSel = False
            
            int��ҳID = Val(NVL(!��ҳID)): intInsure1 = Val(NVL(!����)): strInsureName1 = NVL(!��������)
            strTag = int��ҳID & "|" & Val(NVL(!����)) & "|" & NVL(!��������)
            
            cboPatiNums.AddItem int��ҳID, "��" & int��ҳID & "��סԺ" & IIf(Val(NVL(!����)) <> 0, "(ҽ��)", ""), False, blnSelect, , , strTag
            i = i + 1
            .MoveNext
        Loop
     End With
     Call cboPatiNums.Refresh
    LoadDataPatiNumsToComBox = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function CheckPatiIsVerfy(Optional ByRef strMessage As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鲡���Ƿ����
    '����:strMessage-������Ϣ
    '����:���˺�
    '����:2015-01-05 14:55:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnAll As Boolean, lng��ҳID As Long, i As Long
    

    On Error GoTo errHandle
    '���ﲻ���м��
    If mEditType = g_Ed_������� Or mrsInfo Is Nothing Then CheckPatiIsVerfy = True: Exit Function
    If InStr(mstrPrivs, ";δ��˲�����;����;") > 0 Or InStr(mstrPrivs, ";δ��˲��˳�Ժ����;") > 0 Then CheckPatiIsVerfy = True: Exit Function
    If Val(NVL(mrsInfo!��ҳID)) = 0 Then CheckPatiIsVerfy = True: Exit Function
    
    If CStr(mrsInfo!��ҳID) = mobjBalanceAll.strAllTime Then  'ֻ�����һ��δ��
        If mrsInfo!��˱�־ = 0 Then
            strMessage = "��ǰ����δ��ˣ��㲻�ܶ�δ��˵Ĳ��˽��н��ʡ�"
            Exit Function
        End If
        CheckPatiIsVerfy = True: Exit Function
    End If
    blnAll = True
    For i = 0 To UBound(Split(mobjBalanceAll.strAllTime, ","))
        lng��ҳID = Val(Split(mobjBalanceAll.strAllTime, ",")(i))
        If lng��ҳID <> 0 Then
            If Not Chk�������(mrsInfo!����ID, lng��ҳID) Then
                 mobjBalanceAll.strUnAuditTime = mobjBalanceAll.strUnAuditTime & "," & lng��ҳID
            Else
                blnAll = False
            End If
        Else
            blnAll = False
        End If
    Next
    If mobjBalanceAll.strUnAuditTime <> "" Then mobjBalanceAll.strUnAuditTime = Mid(mobjBalanceAll.strUnAuditTime, 2)
    If blnAll Then
        strMessage = "�ò�������סԺ���ö�û����ˣ����ܽ��н��ʣ�"
        Exit Function
    End If
    CheckPatiIsVerfy = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function CheckInputBlood() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��Ѫ�Ѽ��
    '����:Ѫ�Ѽ��Ϸ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-05 15:18:37
    '����:34260:��Ѫ�Ѽ��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    '0:�����;1-��鲢��ʾ
    If mty_ModulePara.byt����ʱ��Ѫ�Ѽ�� <> 1 Then CheckInputBlood = True: Exit Function
    If InStr(1, "," & mobjBalanceAll.strAllChargeType & ",", ",'K',") = 0 Then CheckInputBlood = True: Exit Function
    If MsgBox("ע��:" & vbCrLf & "    �ò���δ������а�������Ѫ��,�����Ƿ�ֻ����Ѫ��?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then CheckInputBlood = True: Exit Function
    
    mobjBalanceCon.strChargeType = "'K'"
    If ShowBalance(False) Then CheckInputBlood = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function ReadBalanceData(ByRef rsBalance As ADODB.Recordset, ByVal blnUpload As Boolean, Optional ByVal blnInputAfterPati As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��������
    '���:blnUpload-�Ƿ�ֻ���ϴ�����
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-05 15:35:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long
    
    On Error GoTo errHandle
    If Not mrsInfo Is Nothing Then
        If mrsInfo.State = 1 Then lng����ID = Val(NVL(mrsInfo!����ID))
    End If

    
        '��ȡ������Ϣ
    If mEditType = g_Ed_������� Then
        With mobjBalanceCon
            If .strChargeType = "" And .blnCurBalanceOwnerFee = False And blnInputAfterPati = False Then
                Set rsBalance = GetMzBalanceData(lng����ID, .strDeptIDs, _
                        .strClass, .dtBeginDate, .dtEndDate, .strItem, blnUpload, _
                       mty_ModulePara.blnZero, mblnDateMoved, .bytKind, .strChargeType, mty_ModulePara.int����ʱ�� = 1, .strTime, mty_ModulePara.strOwnerPayFeeType, cboDiag.Text)
            Else
                Set rsBalance = GetMzBalanceData(lng����ID, .strDeptIDs, _
                        .strClass, .dtBeginDate, .dtEndDate, .strItem, blnUpload, _
                       mty_ModulePara.blnZero, mblnDateMoved, .bytKind, .strChargeType, mty_ModulePara.int����ʱ�� = 1, .strTime, , cboDiag.Text)
            End If
        End With
        ReadBalanceData = True
        Exit Function
    End If
    With mobjBalanceCon
        If .strChargeType = "" And .blnCurBalanceOwnerFee = False And blnInputAfterPati = False Then
            Set rsBalance = GetZYBalanceData(lng����ID, .strTime, .strDeptIDs, .strClass, _
                .dtBeginDate, .dtEndDate, .strBaby, .strItem, blnUpload, mty_ModulePara.blnZero, _
                mblnDateMoved, .strChargeType, mty_ModulePara.int����ʱ�� = 1, mty_ModulePara.strOwnerPayFeeType, cboDiag.Text)
        Else
            Set rsBalance = GetZYBalanceData(lng����ID, .strTime, .strDeptIDs, .strClass, _
                .dtBeginDate, .dtEndDate, .strBaby, .strItem, blnUpload, mty_ModulePara.blnZero, _
                mblnDateMoved, .strChargeType, mty_ModulePara.int����ʱ�� = 1, , cboDiag.Text)
        End If
    End With
    ReadBalanceData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub InitBalanceMoney()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ�����ʽ��
    '����:���˺�
    '����:2015-01-12 14:11:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With mBalanceInfor
        .dbl����δ�� = 0
        .dbl��ǰ���� = 0
        .dbl�Ѹ��ϼ� = 0
        .dblδ���ϼ� = 0
        .dblҽ��֧���ϼ� = 0
        .dbl��Ԥ���ϼ� = 0
    End With
End Sub

Private Function LoadFeeListFromBalanceID(ByVal lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݽ���ID�����ط�Ŀ������
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-08 18:18:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, str��ҳIds As String
    Dim lngRow As Long, dblMoney(0 To 2) As Double
    Dim intSign As Integer
    On Error GoTo errHandle
    
    Call LoadDetailListFromBalanceID(lng����ID)
    intSign = IIf(mEditType = g_Ed_�������� Or (mEditType = g_Ed_���ݲ鿴 And mblnViewCancel), -1, 1)
    
    strSQL = _
    " Select Mod(B.��¼����,10) as ��¼����, B.NO,B.���,B.�վݷ�Ŀ, " & _
    "          Sum(B.Ӧ�ս��) As Ӧ�ս��," & _
    "          Sum(B.ʵ�ս��) As ʵ�ս��,0 as ���ʽ��" & _
    " From סԺ���ü�¼ A,סԺ���ü�¼ B " & _
    " Where A.����ID=[1] And  Mod(A.��¼����,10)=Mod(B.��¼����,10)  " & _
    "       And A.NO=B.NO And A.���=B.��� And A.��¼״̬ = B.��¼״̬ " & _
    " Group by Mod(B.��¼����,10), B.NO,B.���,B.�վݷ�Ŀ"
    strSQL = strSQL & " UNION ALL " & _
    "   Select Mod(A.��¼����,10) as ��¼����, A.NO,���,A.�վݷ�Ŀ, " & _
    "           0 as Ӧ�ս��,0 as ʵ�ս��,sum(A.���ʽ��) as ���ʽ�� " & _
    "   From סԺ���ü�¼ A " & _
    "   Where A.����ID= [1]  " & _
    "   Group by Mod(A.��¼����,10),A.NO,A.���,A.�վݷ�Ŀ "
    
   
    strSQL = strSQL & " UNION ALL " & vbCrLf & _
        Replace(strSQL, "סԺ���ü�¼", "������ü�¼")

    If mblnNOMoved Then
        strSQL = Replace(Replace(strSQL, "סԺ���ü�¼", "HסԺ���ü�¼"), "������ü�¼", "H������ü�¼")
    End If
    
    strSQL = "" & _
    "   Select �վݷ�Ŀ, sum(Ӧ�ս��) as Ӧ�ս��,sum(ʵ�ս��) as ʵ�ս��,sum(���ʽ��) as ���ʽ�� " & _
    "   From (" & strSQL & ")" & _
    "   Group by �վݷ�Ŀ" & _
    "   Order by �վݷ�Ŀ"
    

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    
    On Error GoTo errHandle
    With vsFeeList
        .Redraw = flexRDNone
        .Clear 1: .Rows = 2
        If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
        str��ҳIds = "": lngRow = 1
        Do While Not rsTemp.EOF
          .TextMatrix(lngRow, .ColIndex("��Ŀ")) = NVL(rsTemp!�վݷ�Ŀ, "δ֪")
          .Cell(flexcpData, lngRow, .ColIndex("Ӧ�ս��")) = Val(.Cell(flexcpData, lngRow, .ColIndex("Ӧ�ս��"))) + Val(NVL(rsTemp!Ӧ�ս��))
          .TextMatrix(lngRow, .ColIndex("Ӧ�ս��")) = Format(Val(.Cell(flexcpData, lngRow, .ColIndex("Ӧ�ս��"))), gstrDec)
          
          .Cell(flexcpData, lngRow, .ColIndex("ʵ�ս��")) = Val(.Cell(flexcpData, lngRow, .ColIndex("ʵ�ս��"))) + Val(NVL(rsTemp!ʵ�ս��))
          .TextMatrix(lngRow, .ColIndex("ʵ�ս��")) = Format(Val(.Cell(flexcpData, lngRow, .ColIndex("ʵ�ս��"))), gstrDec)
          
          .Cell(flexcpData, lngRow, .ColIndex("δ����")) = Val(.Cell(flexcpData, lngRow, .ColIndex("���ʽ��"))) + Val(NVL(rsTemp!���ʽ��))
          .TextMatrix(lngRow, .ColIndex("δ����")) = Format(Val(.Cell(flexcpData, lngRow, .ColIndex("���ʽ��"))), gstrDec)
          
          .Cell(flexcpData, lngRow, .ColIndex("���ʽ��")) = Val(.Cell(flexcpData, lngRow, .ColIndex("���ʽ��"))) + RoundEx(intSign * Val(NVL(rsTemp!���ʽ��)), 6)
          .TextMatrix(lngRow, .ColIndex("���ʽ��")) = Format(Val(.Cell(flexcpData, lngRow, .ColIndex("���ʽ��"))), gstrDec)
          dblMoney(0) = dblMoney(0) + Val(NVL(rsTemp!Ӧ�ս��))
          dblMoney(1) = dblMoney(1) + Val(NVL(rsTemp!ʵ�ս��))
          dblMoney(2) = dblMoney(2) + RoundEx(intSign * Val(NVL(rsTemp!���ʽ��)), 6)
          .Rows = .Rows + 1: lngRow = .Rows - 1
          rsTemp.MoveNext
        Loop
        If str��ҳIds <> "" Then str��ҳIds = Mid(str��ҳIds, 2)
        
        If .TextMatrix(1, .ColIndex("��Ŀ")) <> "" Then
           lngRow = .Rows - 1
          .TextMatrix(lngRow, .ColIndex("��Ŀ")) = "�ϼ�"
          .Cell(flexcpData, lngRow, .ColIndex("Ӧ�ս��")) = dblMoney(0)
          .TextMatrix(lngRow, .ColIndex("Ӧ�ս��")) = Format(dblMoney(0), gstrDec)
          
          .Cell(flexcpData, lngRow, .ColIndex("ʵ�ս��")) = dblMoney(1)
          .TextMatrix(lngRow, .ColIndex("ʵ�ս��")) = Format(dblMoney(1), gstrDec)
          
          .Cell(flexcpData, lngRow, .ColIndex("δ����")) = Val(dblMoney(2))
          .TextMatrix(lngRow, .ColIndex("δ����")) = Format(dblMoney(2), gstrDec)
         
          .Cell(flexcpData, lngRow, .ColIndex("���ʽ��")) = Val(dblMoney(2))
          .TextMatrix(lngRow, .ColIndex("���ʽ��")) = Format(dblMoney(2), gstrDec)
         
           .Cell(flexcpFontBold, lngRow, 0, lngRow, .Cols - 1) = True
        End If
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        .Redraw = flexRDBuffered
    End With
    
    mBalanceInfor.dbl����δ�� = RoundEx(dblMoney(2), 6)
    mBalanceInfor.dbl��ǰ���� = mBalanceInfor.dbl����δ��
    mBalanceInfor.dblδ���ϼ� = mBalanceInfor.dbl����δ��
    
    mblnNotChange = True
    zl_vsGrid_Para_Restore mlngModul, vsFeeList, Me.Name, "�����б�" & mEditType
    txtBalance(Idx_����δ��).Text = Format(dblMoney(2), gstrDec)
    txtBalance(Idx_����δ��).Enabled = False
    txtBalance(Idx_���ν���).Text = Format(dblMoney(2), gstrDec)
    mblnNotChange = False
    LoadFeeListFromBalanceID = True
    Exit Function
errHandle:
    vsFeeList.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function LoadDetailListFromBalanceID(ByVal lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݽ���ID�����ط�Ŀ������
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-08 18:18:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, str��ҳIds As String
    Dim lngRow As Long, dblMoney(0 To 2) As Double
    Dim intSign As Integer
    On Error GoTo errHandle
    
    intSign = IIf(mEditType = g_Ed_�������� Or (mEditType = g_Ed_���ݲ鿴 And mblnViewCancel), -1, 1)
    
    strSQL = _
    " Select Mod(B.��¼����,10) as ��¼����, B.NO,B.���,C.���� As ��Ŀ,Max(B.�Ǽ�ʱ��) As �Ǽ�ʱ��, " & _
    "          Avg(B.Ӧ�ս��) As Ӧ�ս��," & _
    "          Avg(B.ʵ�ս��) As ʵ�ս��,0 as ���ʽ��,Decode(B.��¼״̬,2,2,1) As ��¼״̬,Max(a.�����־) As �����־" & _
    " From סԺ���ü�¼ A,סԺ���ü�¼ B,�շ���ĿĿ¼ C " & _
    " Where A.����ID=[1] And  Mod(A.��¼����,10)=Mod(B.��¼����,10)  " & _
    "       And B.�շ�ϸĿID=C.ID And A.NO=B.NO And A.���=B.��� And A.��¼״̬ = B.��¼״̬ " & _
    " Group by Mod(B.��¼����,10), B.NO,B.���,C.����,Decode(B.��¼״̬,2,2,1)"
    strSQL = strSQL & " UNION ALL " & _
    "   Select Mod(A.��¼����,10) as ��¼����, A.NO,���,B.���� As ��Ŀ,Max(A.�Ǽ�ʱ��) As �Ǽ�ʱ��, " & _
    "           0 as Ӧ�ս��,0 as ʵ�ս��,Sum(A.���ʽ��) as ���ʽ��,Decode(A.��¼״̬,2,2,1) As ��¼״̬,Max(a.�����־) As �����־ " & _
    "   From סԺ���ü�¼ A,�շ���ĿĿ¼ B " & _
    "   Where A.����ID= [1] And A.�շ�ϸĿID=B.ID  " & _
    "   Group by Mod(A.��¼����,10),A.NO,A.���,B.����,Decode(A.��¼״̬,2,2,1) "
    
   
    strSQL = strSQL & " UNION ALL " & vbCrLf & _
        Replace(strSQL, "סԺ���ü�¼", "������ü�¼")

    If mblnNOMoved Then
        strSQL = Replace(Replace(strSQL, "סԺ���ü�¼", "HסԺ���ü�¼"), "������ü�¼", "H������ü�¼")
    End If
    
    strSQL = "" & _
    "   Select Max(�Ǽ�ʱ��) As �Ǽ�ʱ��,NO,���,��Ŀ, sum(Ӧ�ս��) as Ӧ�ս��,sum(ʵ�ս��) as ʵ�ս��," & _
    "          sum(���ʽ��) as ���ʽ��,��¼״̬,Max(�����־) As �����־ " & _
    "   From (" & strSQL & ")" & _
    "   Group by NO,���,��Ŀ,��¼״̬" & _
    "   Order by NO,���"

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    
    On Error GoTo errHandle
    With vsDetailList
        .Redraw = flexRDNone
        .Clear 1: .Rows = 2
        If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
        str��ҳIds = "": lngRow = 1
        .TextMatrix(0, .ColIndex("δ����")) = "ʵ�ս��"
        If intSign = -1 Then
            .TextMatrix(0, .ColIndex("���ʽ��")) = "���Ͻ��"
        Else
            .TextMatrix(0, .ColIndex("���ʽ��")) = "���ʽ��"
        End If
        Do While Not rsTemp.EOF
            .TextMatrix(.Rows - 1, .ColIndex("����")) = Format(NVL(rsTemp!�Ǽ�ʱ��), "yyyy-mm-dd")
            .TextMatrix(.Rows - 1, .ColIndex("����")) = NVL(rsTemp!NO)
            .TextMatrix(.Rows - 1, .ColIndex("��Ŀ")) = NVL(rsTemp!��Ŀ)
            .TextMatrix(.Rows - 1, .ColIndex("δ����")) = Format(NVL(rsTemp!ʵ�ս��), gstrDec)
            .Cell(flexcpData, .Rows - 1, .ColIndex("δ����")) = Val(NVL(rsTemp!ʵ�ս��))
            .TextMatrix(.Rows - 1, .ColIndex("���ʽ��")) = Format(intSign * Val(NVL(rsTemp!���ʽ��)), gstrDec)
            .Cell(flexcpData, .Rows - 1, .ColIndex("���ʽ��")) = intSign * Val(NVL(rsTemp!���ʽ��))
            .TextMatrix(.Rows - 1, .ColIndex("���")) = Val(NVL(rsTemp!���))
            If mEditType = g_Ed_������� Or mblnCurMzBalanceNo Then
                .Cell(flexcpData, .Rows - 1, .ColIndex("���")) = Val(NVL(rsTemp!�����־))
            End If
            .Rows = .Rows + 1
            rsTemp.MoveNext
        Loop
        If .Rows > 2 Then .Rows = .Rows - 1
        .Cell(flexcpBackColor, 1, .ColIndex("���ʽ��"), .Rows - 1, .ColIndex("���ʽ��")) = .Cell(flexcpBackColor, 1, .ColIndex("����"), 0.1, .ColIndex("����"))
        .Redraw = flexRDBuffered
    End With
    LoadDetailListFromBalanceID = True
    Exit Function
errHandle:
    vsDetailList.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetCentMoney(ByVal dblMoney As Double) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݷֱҴ������,���طֱҴ����Ľ��
    '���:dblMoney-δ�����ԭʼ���
    '����:���طֱҴ����Ľ��
    '����:���˺�
    '����:2015-01-26 10:57:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card
    On Error GoTo errHandle
    Set objCard = IDKindPaymentsType.GetCurCard
    If objCard Is Nothing Then GetCentMoney = Format(dblMoney, "0.00"): Exit Function
    '���ֽ��,������λС��
    If objCard.�������� <> 1 Then GetCentMoney = Format(dblMoney, "0.00"): Exit Function
    
    If mYBInFor.intInsure = 0 Then
        GetCentMoney = CentMoney(CCur(dblMoney))
        Exit Function
    End If
    If MCPAR.�ֱҴ��� Then
        GetCentMoney = CentMoney(CCur(dblMoney))
    Else
        GetCentMoney = Format(dblMoney, "0.00")
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function

Private Sub LoadCurOwnerPayInfor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ص�ǰ������Ϣ
    '����:���˺�
    '����:2015-01-12 14:14:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, lngColor As Long, objCard As Card
    Dim dblTemp As Double
    Dim dblʣ���Ը� As Double
    
    With mBalanceInfor
        dblMoney = RoundEx(.dblδ���ϼ� - .dbl��Ԥ���ϼ�, 5)
        lbl�Ը��ϼ�.Caption = Format(.dbl��ǰ���� - .dblҽ��֧���ϼ�, mstrDec)
        If Val(lbl�Ը��ϼ�.Caption) < 0 Then
            stcTittleTotal.Caption = "�˿�ϼ�"
        Else
            stcTittleTotal.Caption = "�Ը��ϼ�"
        End If
        lbl�Ը��ϼ�.Caption = Format(Abs(.dbl��ǰ���� - .dblҽ��֧���ϼ�), mstrDec)
        dblTemp = GetCentMoney(dblMoney)
            
        lblʣ���Ը�.Caption = Format(RoundEx(dblTemp + mPatiInfor.dblδ���ۼ�, 6), mstrDec)
        lblʣ���Ը�.Tag = Format(dblTemp, mstrDec)
        
        If dblMoney < 0 Then
            Set objCard = IDKind�Ҳ�.GetCurCard
            If Not objCard Is Nothing Then
                If objCard.�ӿ���� <> 1 Then
                    '��Ԥ��
                    dblTemp = Val(lblʣ���Ը�.Caption) - Val(txtBalance(Idx_�Ҳ�).Text)
                    dblTemp = GetCentMoney(Abs(dblTemp))
                    lblʣ���Ը�.Caption = Format(dblTemp, mstrDec)
                    lblʣ���Ը�.Tag = dblTemp
                End If
            End If
        End If
        
        Select Case mEditType
        Case g_Ed_ȡ������, g_Ed_��������
            stcTittleTotal.Caption = "�˿�ϼ�"
            stcTittile.Caption = IIf(dblMoney >= 0, "��ǰδ��", "��ǰδ��")
            mPatiInfor.bln�˿��־ = dblMoney >= 0
            lblBalance(4).Caption = IIf(dblMoney >= 0, "��    ��", "��    ��")
            If dblMoney >= 0 Then
                '�˿�
                lblʣ���Ը�.ForeColor = Mid(mty_ModulePara.str��ǰ����ɫ, InStr(mty_ModulePara.str��ǰ����ɫ, "|") + 1)
                txtBalance(4).ForeColor = Mid(mty_ModulePara.str�ɿ�ɫ, InStr(mty_ModulePara.str�ɿ�ɫ, "|") + 1)
                lblBalance(4).ForeColor = Mid(mty_ModulePara.str�ɿ�ɫ, InStr(mty_ModulePara.str�ɿ�ɫ, "|") + 1)
                IDKindPaymentsType.ForeColor = Mid(mty_ModulePara.str�ɿ�ɫ, InStr(mty_ModulePara.str�ɿ�ɫ, "|") + 1)
            Else
                '�տ�
                lblʣ���Ը�.ForeColor = Mid(mty_ModulePara.str��ǰ����ɫ, 1, InStr(mty_ModulePara.str��ǰ����ɫ, "|") - 1)
                txtBalance(4).ForeColor = Mid(mty_ModulePara.str�ɿ�ɫ, 1, InStr(mty_ModulePara.str�ɿ�ɫ, "|") - 1)
                lblBalance(4).ForeColor = Mid(mty_ModulePara.str�ɿ�ɫ, 1, InStr(mty_ModulePara.str�ɿ�ɫ, "|") - 1)
                IDKindPaymentsType.ForeColor = Mid(mty_ModulePara.str�ɿ�ɫ, 1, InStr(mty_ModulePara.str�ɿ�ɫ, "|") - 1)
            End If
            
            Set objCard = IDKind�Ҳ�.GetCurCard
            If Not objCard Is Nothing Then
                If objCard.�ӿ���� <> 1 And dblMoney > 0 Then
                    '������Ԥ��
                    dblTemp = Val(lblʣ���Ը�.Caption) - Val(txtBalance(Idx_�Ҳ�).Text)
                    dblTemp = GetCentMoney(Abs(dblTemp))
                    lblʣ���Ը�.Caption = Format(dblTemp, mstrDec)
                    lblʣ���Ը�.Tag = dblTemp
                
                End If
            End If
            lngColor = IIf(dblMoney >= 0, vbRed, vbBlue)
            lblCurPaymentInfor.Caption = "��ǰ�˿����"
        Case g_Ed_��������
            stcTittleTotal.Caption = "�˿�ϼ�"
            stcTittile.Caption = IIf(dblMoney >= 0, "��ǰδ��", "��ǰδ��")
            lblBalance(4).Caption = IIf(dblMoney >= 0, "��    ��", "��    ��")
            mPatiInfor.bln�˿��־ = dblMoney >= 0
            If dblMoney >= 0 Then
                '�˿�
                lblʣ���Ը�.ForeColor = Mid(mty_ModulePara.str��ǰ����ɫ, InStr(mty_ModulePara.str��ǰ����ɫ, "|") + 1)
                txtBalance(4).ForeColor = Mid(mty_ModulePara.str�ɿ�ɫ, InStr(mty_ModulePara.str�ɿ�ɫ, "|") + 1)
                lblBalance(4).ForeColor = Mid(mty_ModulePara.str�ɿ�ɫ, InStr(mty_ModulePara.str�ɿ�ɫ, "|") + 1)
                IDKindPaymentsType.ForeColor = Mid(mty_ModulePara.str�ɿ�ɫ, InStr(mty_ModulePara.str�ɿ�ɫ, "|") + 1)
            Else
                '�տ�
                lblʣ���Ը�.ForeColor = Mid(mty_ModulePara.str��ǰ����ɫ, 1, InStr(mty_ModulePara.str��ǰ����ɫ, "|") - 1)
                txtBalance(4).ForeColor = Mid(mty_ModulePara.str�ɿ�ɫ, 1, InStr(mty_ModulePara.str�ɿ�ɫ, "|") - 1)
                lblBalance(4).ForeColor = Mid(mty_ModulePara.str�ɿ�ɫ, 1, InStr(mty_ModulePara.str�ɿ�ɫ, "|") - 1)
                IDKindPaymentsType.ForeColor = Mid(mty_ModulePara.str�ɿ�ɫ, 1, InStr(mty_ModulePara.str�ɿ�ɫ, "|") - 1)
            End If
            
            Set objCard = IDKind�Ҳ�.GetCurCard
            If Not objCard Is Nothing Then
                If objCard.�ӿ���� <> 1 And dblMoney > 0 Then
                    '������Ԥ��
                    dblTemp = Val(lblʣ���Ը�.Caption) - Val(txtBalance(Idx_�Ҳ�).Text)
                    dblTemp = GetCentMoney(Abs(dblTemp))
                    lblʣ���Ը�.Caption = Format(dblTemp, mstrDec)
                    lblʣ���Ը�.Tag = dblTemp
                End If
            End If
            lngColor = IIf(dblMoney >= 0, vbRed, vbBlue)
            lblCurPaymentInfor.Caption = "��ǰ�˿����"
        Case Else
            If chkCancel.Value = 1 Then
                stcTittleTotal.Caption = "�˿�ϼ�"
                stcTittile.Caption = IIf(dblMoney >= 0, "��ǰδ��", "��ǰδ��")
                lblBalance(4).Caption = IIf(dblMoney >= 0, "��    ��", "��    ��")
                mPatiInfor.bln�˿��־ = dblMoney >= 0
                If dblMoney >= 0 Then
                    '�˿�
                    lblʣ���Ը�.ForeColor = Mid(mty_ModulePara.str��ǰ����ɫ, InStr(mty_ModulePara.str��ǰ����ɫ, "|") + 1)
                    txtBalance(4).ForeColor = Mid(mty_ModulePara.str�ɿ�ɫ, InStr(mty_ModulePara.str�ɿ�ɫ, "|") + 1)
                    lblBalance(4).ForeColor = Mid(mty_ModulePara.str�ɿ�ɫ, InStr(mty_ModulePara.str�ɿ�ɫ, "|") + 1)
                    IDKindPaymentsType.ForeColor = Mid(mty_ModulePara.str�ɿ�ɫ, InStr(mty_ModulePara.str�ɿ�ɫ, "|") + 1)
                Else
                    '�տ�
                    lblʣ���Ը�.ForeColor = Mid(mty_ModulePara.str��ǰ����ɫ, 1, InStr(mty_ModulePara.str��ǰ����ɫ, "|") - 1)
                    txtBalance(4).ForeColor = Mid(mty_ModulePara.str�ɿ�ɫ, 1, InStr(mty_ModulePara.str�ɿ�ɫ, "|") - 1)
                    lblBalance(4).ForeColor = Mid(mty_ModulePara.str�ɿ�ɫ, 1, InStr(mty_ModulePara.str�ɿ�ɫ, "|") - 1)
                    IDKindPaymentsType.ForeColor = Mid(mty_ModulePara.str�ɿ�ɫ, 1, InStr(mty_ModulePara.str�ɿ�ɫ, "|") - 1)
                End If
                
                Set objCard = IDKind�Ҳ�.GetCurCard
                If Not objCard Is Nothing Then
                    If objCard.�ӿ���� <> 1 And dblMoney > 0 Then
                        '������Ԥ��
                        dblTemp = Val(lblʣ���Ը�.Caption) - Val(txtBalance(Idx_�Ҳ�).Text)
                        dblTemp = GetCentMoney(Abs(dblTemp))
                        lblʣ���Ը�.Caption = Format(dblTemp, mstrDec)
                        lblʣ���Ը�.Tag = dblTemp
                    
                    End If
                End If
                lngColor = IIf(dblMoney >= 0, vbRed, vbBlue)
                lblCurPaymentInfor.Caption = "��ǰ�˿����"
            Else
                stcTittile.Caption = IIf(Val(lblʣ���Ը�.Caption) < 0, "��ǰδ��", "��ǰδ��")
                lblBalance(4).Caption = IIf(Val(lblʣ���Ը�.Caption) < 0, "��    ��", "��    ��")
                mPatiInfor.bln�˿��־ = dblMoney < 0
                If Val(lblʣ���Ը�.Caption) < 0 Then
                    '�˿�
                    lblʣ���Ը�.ForeColor = Mid(mty_ModulePara.str��ǰ����ɫ, InStr(mty_ModulePara.str��ǰ����ɫ, "|") + 1)
                    txtBalance(4).ForeColor = Mid(mty_ModulePara.str�ɿ�ɫ, InStr(mty_ModulePara.str�ɿ�ɫ, "|") + 1)
                    lblBalance(4).ForeColor = Mid(mty_ModulePara.str�ɿ�ɫ, InStr(mty_ModulePara.str�ɿ�ɫ, "|") + 1)
                    IDKindPaymentsType.ForeColor = Mid(mty_ModulePara.str�ɿ�ɫ, InStr(mty_ModulePara.str�ɿ�ɫ, "|") + 1)
                Else
                    '�տ�
                    lblʣ���Ը�.ForeColor = Mid(mty_ModulePara.str��ǰ����ɫ, 1, InStr(mty_ModulePara.str��ǰ����ɫ, "|") - 1)
                    txtBalance(4).ForeColor = Mid(mty_ModulePara.str�ɿ�ɫ, 1, InStr(mty_ModulePara.str�ɿ�ɫ, "|") - 1)
                    lblBalance(4).ForeColor = Mid(mty_ModulePara.str�ɿ�ɫ, 1, InStr(mty_ModulePara.str�ɿ�ɫ, "|") - 1)
                    IDKindPaymentsType.ForeColor = Mid(mty_ModulePara.str�ɿ�ɫ, 1, InStr(mty_ModulePara.str�ɿ�ɫ, "|") - 1)
                End If
                
                Set objCard = IDKind�Ҳ�.GetCurCard
                If Not objCard Is Nothing Then
                    If objCard.�ӿ���� <> 1 And dblMoney < 0 Then
                        '��Ԥ��
                        dblTemp = Val(lblʣ���Ը�.Caption) - Val(txtBalance(Idx_�Ҳ�).Text)
                        dblTemp = GetCentMoney(Abs(dblTemp))
                        lblʣ���Ը�.Caption = Format(dblTemp, mstrDec)
                        lblʣ���Ը�.Tag = dblTemp
                    End If
                End If
                '����������ʾ
                lngColor = IIf(dblMoney < 0, vbRed, vbBlue)
                If chkCancel.Value = 1 Then
                    lblCurPaymentInfor.Caption = "��ǰ�˿����"
                Else
                    lblCurPaymentInfor.Caption = "��ǰ֧�����"
                End If
                
            End If
        End Select
    End With
    If Val(lblʣ���Ը�.Caption) < 0 Then lblʣ���Ը�.Caption = Format(Abs(RoundEx(lblʣ���Ը�.Caption, 6)), mstrDec)
    Set objCard = IDKind�Ҳ�.GetCurCard
    If Not objCard Is Nothing Then
        If objCard.�ӿ���� <> 2 And objCard.�ӿ���� <> 3 Then
            Call LoadDefaultMoney
        End If
    Else
        Call LoadDefaultMoney
    End If
    Show����� False

End Sub

Private Function LoadFeeList() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ط�Ŀ������
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-05 18:00:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, dblMoney(0 To 2) As Double
 
    On Error GoTo errHandle
    Call LoadDetailList
    With vsFeeList
        .Redraw = flexRDNone
        .Clear 1: .Rows = 2
        If mrsFeeList.RecordCount <> 0 Then mrsFeeList.MoveFirst
        Do While Not mrsFeeList.EOF
           lngRow = .FindRow(NVL(mrsFeeList!��Ŀ, "δ֪"), "1", .ColIndex("��Ŀ"), , True)
           If lngRow < 0 Then
                If .TextMatrix(1, .ColIndex("��Ŀ")) = "" Then
                    lngRow = 1
                Else
                    .Rows = .Rows + 1: lngRow = .Rows - 1
                End If
           End If
           
           If .TextMatrix(1, .ColIndex("��Ŀ")) = "" Then lngRow = 1
          .TextMatrix(lngRow, .ColIndex("��Ŀ")) = NVL(mrsFeeList!��Ŀ, "δ֪")
          
          .Cell(flexcpData, lngRow, .ColIndex("Ӧ�ս��")) = Val(.Cell(flexcpData, lngRow, .ColIndex("Ӧ�ս��"))) + Val(NVL(mrsFeeList!Ӧ�ս��))
          .TextMatrix(lngRow, .ColIndex("Ӧ�ս��")) = Format(Val(.Cell(flexcpData, lngRow, .ColIndex("Ӧ�ս��"))), gstrDec)
          
          .Cell(flexcpData, lngRow, .ColIndex("ʵ�ս��")) = Val(.Cell(flexcpData, lngRow, .ColIndex("ʵ�ս��"))) + Val(NVL(mrsFeeList!ʵ�ս��))
          .TextMatrix(lngRow, .ColIndex("ʵ�ս��")) = Format(Val(.Cell(flexcpData, lngRow, .ColIndex("ʵ�ս��"))), gstrDec)
          
          .Cell(flexcpData, lngRow, .ColIndex("δ����")) = Val(.Cell(flexcpData, lngRow, .ColIndex("δ����"))) + Val(NVL(mrsFeeList!δ����))
          .TextMatrix(lngRow, .ColIndex("δ����")) = Format(Val(.Cell(flexcpData, lngRow, .ColIndex("δ����"))), gstrDec)
            
          dblMoney(0) = RoundEx(dblMoney(0) + Val(NVL(mrsFeeList!Ӧ�ս��)), 5)
          dblMoney(1) = RoundEx(dblMoney(1) + Val(NVL(mrsFeeList!ʵ�ս��)), 5)
          dblMoney(2) = RoundEx(dblMoney(2) + Val(NVL(mrsFeeList!δ����)), 5)
          
            mrsFeeList.MoveNext
        Loop
        .ColSort(.ColIndex("��Ŀ")) = flexSortUseColSort
        If .TextMatrix(1, .ColIndex("��Ŀ")) <> "" Then
          .Rows = .Rows + 1: lngRow = .Rows - 1
          .TextMatrix(lngRow, .ColIndex("��Ŀ")) = "�ϼ�"
          .Cell(flexcpData, lngRow, .ColIndex("Ӧ�ս��")) = dblMoney(0)
          .TextMatrix(lngRow, .ColIndex("Ӧ�ս��")) = Format(dblMoney(0), gstrDec)
          
          .Cell(flexcpData, lngRow, .ColIndex("ʵ�ս��")) = dblMoney(1)
          .TextMatrix(lngRow, .ColIndex("ʵ�ս��")) = Format(dblMoney(1), gstrDec)
          
          .Cell(flexcpData, lngRow, .ColIndex("δ����")) = Val(dblMoney(2))
          .TextMatrix(lngRow, .ColIndex("δ����")) = Format(dblMoney(2), gstrDec)
         
           .Cell(flexcpFontBold, lngRow, 0, lngRow, .Cols - 1) = True
        End If
         .Redraw = flexRDBuffered
    End With
    mBalanceInfor.dbl����δ�� = dblMoney(2)
    mBalanceInfor.dbl��ǰ���� = dblMoney(2)
    mBalanceInfor.dblδ���ϼ� = dblMoney(2)
    
    mblnNotChange = True
    txtBalance(Idx_����δ��).Text = Format(dblMoney(2), gstrDec)
    txtBalance(Idx_���ν���).Text = Format(dblMoney(2), gstrDec)
    mblnNotChange = False
    Call LoadCurOwnerPayInfor '���ص�ǰ�Ը���Ϣ
    LoadFeeList = True
    Exit Function
errHandle:
     vsFeeList.Redraw = flexRDNone
     Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function LoadDetailList() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ط�Ŀ������
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-05 18:00:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, dblMoney(0 To 2) As Double
    Dim i As Long
 
    On Error GoTo errHandle
    With vsDetailList
        .Redraw = flexRDNone
        .Clear 1: .Rows = 2
        If mrsFeeList.RecordCount <> 0 Then mrsFeeList.MoveFirst
        Do While Not mrsFeeList.EOF
            .TextMatrix(.Rows - 1, .ColIndex("����")) = Format(NVL(mrsFeeList!ʱ��), "yyyy-mm-dd")
            .TextMatrix(.Rows - 1, .ColIndex("����")) = NVL(mrsFeeList!���ݺ�)
            .TextMatrix(.Rows - 1, .ColIndex("��Ŀ")) = NVL(mrsFeeList!��Ŀ)
            .TextMatrix(.Rows - 1, .ColIndex("δ����")) = Format(NVL(mrsFeeList!δ����), gstrDec)
            .Cell(flexcpData, .Rows - 1, .ColIndex("δ����")) = Val(NVL(mrsFeeList!δ����))
            .TextMatrix(.Rows - 1, .ColIndex("���ʽ��")) = Format(NVL(mrsFeeList!���ʽ��), gstrDec)
            .Cell(flexcpData, .Rows - 1, .ColIndex("���ʽ��")) = Val(NVL(mrsFeeList!���ʽ��))
            .TextMatrix(.Rows - 1, .ColIndex("ID")) = NVL(mrsFeeList!ID, 0)
            .TextMatrix(.Rows - 1, .ColIndex("��¼����")) = Val(NVL(mrsFeeList!��¼����))
            .TextMatrix(.Rows - 1, .ColIndex("��¼״̬")) = IIf(Val(NVL(mrsFeeList!��¼״̬)) = 3, 1, Val(NVL(mrsFeeList!��¼״̬)))
            .TextMatrix(.Rows - 1, .ColIndex("ִ��״̬")) = Val(NVL(mrsFeeList!ִ��״̬))
            .TextMatrix(.Rows - 1, .ColIndex("���")) = Val(NVL(mrsFeeList!���))
            If mEditType = g_Ed_������� Then
                .Cell(flexcpData, .Rows - 1, .ColIndex("���")) = Val(NVL(mrsFeeList!�����־))
            End If
            .Rows = .Rows + 1
            mrsFeeList.MoveNext
        Loop
        If mYBInFor.intInsure <> 0 Then
            .Cell(flexcpBackColor, 1, .ColIndex("���ʽ��"), .Rows - 1, .ColIndex("���ʽ��")) = .Cell(flexcpBackColor, 1, .ColIndex("����"))
        Else
            .Cell(flexcpBackColor, 1, .ColIndex("���ʽ��"), .Rows - 1, .ColIndex("���ʽ��")) = &HFFFFC0
        End If
        If .TextMatrix(1, .ColIndex("����")) <> "" Then .Rows = .Rows - 1
         .Redraw = flexRDBuffered
    End With
    LoadDetailList = True
    Exit Function
errHandle:
     vsDetailList.Redraw = flexRDNone
     Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function LoadBalanceDepositList(ByVal lng����ID As Long, _
    ByVal lng����ID As Long, ByVal blnDateMoved As Boolean, _
    str��ҳIds As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ָ�����ʵ��ĳ�Ԥ����Ϣ
    '���:lng����ID-ָ���Ľ���ID
    '     blnDateMoved-��ǰ�Ƿ��ƶ����󱸱���
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-08 15:09:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, i As Long
    Dim dblTotal As Double
    Dim lngԭ����ID As Long
    
    On Error GoTo errHandle
    Set rsTemp = GetBalanceDeposit(lng����ID, blnDateMoved)
    With vsDeposit
        .Redraw = flexRDNone
        .Rows = 2: .Clear 1
        i = 1
        .Cell(flexcpData, 0, 0, .Rows - 1, .Cols - 1) = ""
        If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
        'ID,���ݺ�,Ʊ�ݺ�,����,���㷽ʽ, ���
        Do While Not rsTemp.EOF
            .RowData(i) = ""
            .TextMatrix(i, .ColIndex("ID")) = rsTemp!ID
            .TextMatrix(i, .ColIndex("���ݺ�")) = rsTemp!���ݺ�
            .TextMatrix(i, .ColIndex("Ʊ�ݺ�")) = "" & rsTemp!Ʊ�ݺ�
            .TextMatrix(i, .ColIndex("�տ�����")) = Format(rsTemp!����, "yyyy-MM-dd")
            .TextMatrix(i, .ColIndex("���㷽ʽ")) = NVL(rsTemp!���㷽ʽ)
            .TextMatrix(i, .ColIndex("��Ԥ��")) = Format(rsTemp!���, "0.00")
            .TextMatrix(i, .ColIndex("�����ID")) = Val(NVL(rsTemp!�����ID))
            .TextMatrix(i, .ColIndex("�Ƿ����ѿ�")) = Val(NVL(rsTemp!�Ƿ����ѿ�))
            .TextMatrix(i, .ColIndex("���������")) = NVL(rsTemp!���������)
            .TextMatrix(i, .ColIndex("������ˮ��")) = NVL(rsTemp!������ˮ��)
            .TextMatrix(i, .ColIndex("����")) = NVL(rsTemp!����)
            .TextMatrix(i, .ColIndex("����˵��")) = NVL(rsTemp!����˵��)
            .TextMatrix(i, .ColIndex("�Ƿ�����")) = Val(NVL(rsTemp!�Ƿ�����))
            .TextMatrix(i, .ColIndex("�Ƿ�ȫ��")) = Val(NVL(rsTemp!�Ƿ�ȫ��))
            .TextMatrix(i, .ColIndex("�Ƿ�ȱʡ����")) = Val(NVL(rsTemp!�Ƿ�ȱʡ����))
            .TextMatrix(i, .ColIndex("�Ƿ�ת�ʼ�����")) = Val(NVL(rsTemp!�Ƿ�ת�ʼ�����))
            .Rows = .Rows + 1: i = i + 1
            dblTotal = dblTotal + Val(NVL(rsTemp!���))
            rsTemp.MoveNext
        Loop
        .Row = 1: .Col = .Cols - 1
        .Redraw = flexRDBuffered
        If i > 1 Then .Rows = .Rows - 1
    End With
    
    txtBalance(Idx_��Ԥ��).Text = Format(dblTotal, "0.00")
    chkDeposit.Tag = dblTotal
    mBalanceInfor.dbl��Ԥ���ϼ� = dblTotal
    
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    If Not rsTemp.EOF Then
        lblTicketCount.Caption = "Ԥ�����վ�:" & rsTemp.RecordCount & "��  �ϼ�:" & Format(dblTotal, "0.00") & "Ԫ"
    Else
        lblTicketCount.Caption = ""
    End If
    If rsTemp.RecordCount <> 0 Then LoadBalanceDepositList = True: Exit Function
    
    If mEditType = g_Ed_�������� Then
        '����ԭ��������
        If mblnNotChange Then Exit Function
        mblnNotChange = True
        lngԭ����ID = zlGetFormerBalanceID(mBalanceInfor.strNO)
        LoadBalanceDepositList = LoadBalanceDepositList(lng����ID, lngԭ����ID, blnDateMoved, str��ҳIds)
      
        If mBalanceInfor.dbl��Ԥ���ϼ� <> 0 Then chkDeposit.Value = 1
        mblnNotChange = False
        LoadBalanceDepositList = True
        Exit Function
    End If
    If mEditType <> g_Ed_���ݲ鿴 And mEditType <> g_Ed_�������� And chkCancel.Value <> 1 Then
        '���¼���Ԥ��
        If LoadDepositList(lng����ID, str��ҳIds) = False Then Exit Function
    End If
    LoadBalanceDepositList = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function

Private Function LoadDepositList(ByVal lng����ID As Long, _
    ByVal str��ҳIds As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ԥ������
    '���:lng����ID-����ID
    '     str��ҳIDs:����ö��ŷ���
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-05 18:32:22
    '   mbln����תסԺ:36984
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strסԺ���� As String, i As Long, str���㷽ʽ As String
    Dim intTYPE As Integer, dblMoney As Double, dblTotal As Double
    Dim lng�����ID As Long
    On Error GoTo errHandle
    
    '��ʾԤ����ϸ
    strסԺ���� = "": intTYPE = 1
    If mEditType = g_Ed_סԺ���� Or (mEditType <> g_Ed_������� And mblnCurMzBalanceNo = False) Then
        strסԺ���� = str��ҳIds
        intTYPE = 2
    End If
    
    Set mrsDeposit = GetDeposit(lng����ID, mblnDateMoved, strסԺ����, mbln����תסԺ, mstrPepositDate, intTYPE, mrs���㷽ʽ)
    dblMoney = mBalanceInfor.dblδ���ϼ�
    
    With vsDeposit
        .Redraw = flexRDNone
        .Rows = 2
        .Clear 1
        .Cell(flexcpData, 0, 0, .Rows - 1, .Cols - 1) = ""
        str���㷽ʽ = ""
        If mrsDeposit.RecordCount <> 0 Then mrsDeposit.MoveFirst
        i = 1
        Do While Not mrsDeposit.EOF
            .RowData(i) = Val(NVL(mrsDeposit!��¼״̬))
            '.TextMatrix(i, .ColIndex("��־")) = i
            .TextMatrix(i, .ColIndex("ID")) = mrsDeposit!ID
            .TextMatrix(i, .ColIndex("���ݺ�")) = mrsDeposit!NO
            .TextMatrix(i, .ColIndex("Ʊ�ݺ�")) = "" & mrsDeposit!Ʊ�ݺ�
            .TextMatrix(i, .ColIndex("�տ�����")) = Format(mrsDeposit!����, "yyyy-MM-dd")
            .TextMatrix(i, .ColIndex("���㷽ʽ")) = NVL(mrsDeposit!���㷽ʽ)
            .TextMatrix(i, .ColIndex("���")) = Format(mrsDeposit!���, "0.00")
            .TextMatrix(i, .ColIndex("Ԥ��ID")) = NVL(mrsDeposit!Ԥ��ID)
            
             lng�����ID = Val(NVL(mrsDeposit!�����ID))
             If lng�����ID = 0 Then lng�����ID = Val(NVL(mrsDeposit!���㿨���))

            .TextMatrix(i, .ColIndex("�����ID")) = lng�����ID
            .TextMatrix(i, .ColIndex("�Ƿ����ѿ�")) = Val(NVL(mrsDeposit!�Ƿ����ѿ�))
            .TextMatrix(i, .ColIndex("����")) = NVL(mrsDeposit!����)
            .TextMatrix(i, .ColIndex("���������")) = NVL(mrsDeposit!���������)
            .TextMatrix(i, .ColIndex("������ˮ��")) = NVL(mrsDeposit!������ˮ��)
            .TextMatrix(i, .ColIndex("����˵��")) = NVL(mrsDeposit!����˵��)
            .TextMatrix(i, .ColIndex("�Ƿ�����")) = Val(NVL(mrsDeposit!�Ƿ�����))
            .TextMatrix(i, .ColIndex("�Ƿ�ȫ��")) = Val(NVL(mrsDeposit!�Ƿ�ȫ��))
            .TextMatrix(i, .ColIndex("�Ƿ�ȱʡ����")) = Val(NVL(mrsDeposit!�Ƿ�ȱʡ����))
            .TextMatrix(i, .ColIndex("�Ƿ�ת�ʼ�����")) = Val(NVL(mrsDeposit!ת�ʼ�����))
            
            
            If mbln����תסԺ Or _
                (mobjBalanceCon.blnCurBalanceOwnerFee And mty_ModulePara.bln�Է�ȱʡʹ��Ԥ��) Then
                If Val(NVL(mrsDeposit!���)) <= dblMoney Then
                    .TextMatrix(i, .ColIndex("��Ԥ��")) = Format(mrsDeposit!���, "0.00")
                    dblMoney = dblMoney - RoundEx(Val(NVL(mrsDeposit!���)), 2)
                ElseIf dblMoney <> 0 Then
                    .TextMatrix(i, .ColIndex("��Ԥ��")) = Format(dblMoney, "0.00")
                    dblMoney = 0
                End If
            ElseIf Not mobjBalanceCon.blnCurBalanceOwnerFee Then
                .TextMatrix(i, .ColIndex("��Ԥ��")) = Format(mrsDeposit!���, "0.00")
            End If
            
            dblTotal = dblTotal + RoundEx(Val(NVL(mrsDeposit!���)), 2)
            i = i + 1
            .Rows = .Rows + 1
            mrsDeposit.MoveNext
        Loop
        .Row = 1: .Col = .Cols - 1
        If i >= 2 And .Rows >= 2 Then .Rows = .Rows - 1
        .Redraw = flexRDBuffered
    End With
    zl_vsGrid_Para_Restore mlngModul, vsDeposit, Me.Name, "Ԥ���б�"
    
    '�����113702,����,2017/08/30,��ʽ������ʵ�ʽ��
    mPatiInfor.dblʵ����� = RoundEx(dblTotal, 6)
    If mrsDeposit.RecordCount <> 0 Then mrsDeposit.MoveFirst
    If Not mrsDeposit.EOF Then
        lblTicketCount.Caption = "Ԥ�����վ�:" & mrsDeposit.RecordCount & "��  �ϼ�:" & Format(dblTotal, "0.00") & "Ԫ"
    Else
        lblTicketCount.Caption = ""
    End If
    Call SetUpDown
    LoadDepositList = True
    Exit Function
errHandle:
    vsDeposit.Redraw = flexRDBuffered
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function GetDefaultHospitalizedDate(ByVal lng����ID As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡȱʡ��סԺ����
    '���:lng����ID-����ID
    '����:�����ϴ���;���ʵĽ�������,����;����ʱ,���ؿ�
    '����:���˺�
    '����:2015-01-06 15:25:02
    '˵��:ԭ�������30043
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    
    strSQL = "" & _
    "   Select To_Char(Max(��������) + 1,'yyyy-mm-dd') as �������� " & _
    "   From ���˽��ʼ�¼ " & _
    "   Where  ��¼״̬=1  And ����iD=[1] and Nvl(��;����,0)=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    If rsTemp.EOF Then Exit Function
    GetDefaultHospitalizedDate = NVL(rsTemp!��������)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 

Private Function GetPatiHospitalzedDateRange(ByRef dtBeginDate As Date, ByRef dtEndDate As Date) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����: ��ȡ���˵����Ժʱ��,���ﲡ��ȡ������С����ʱ��
    '����:dtBeginDate-��ʼʱ��
    '     dtEndDate-����ʱ��
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-06 15:43:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strDefaultDate As String, lng��ҳID As Long, lng����ID As Long
    Dim strTime As String
    
    On Error GoTo errHandle
    
    If mrsInfo Is Nothing Then Exit Function
    If mrsInfo.State = 0 Then Exit Function
    If mrsInfo.RecordCount = 0 Then Exit Function
    
    lng����ID = Val(NVL(mrsInfo!����ID))
    
    strTime = mobjBalanceCon.strTime
    If strTime = "" Then strTime = mobjBalanceAll.strAllTime
        
    strDefaultDate = ""
    If mEditType <> g_Ed_������� And strTime <> "" Then
        strDefaultDate = GetDefaultHospitalizedDate(lng����ID)
    End If

    Call GetFeeDate(dtBeginDate, dtEndDate)
    If Val(NVL(mrsInfo!��ҳID)) = 0 Then GetPatiHospitalzedDateRange = True: Exit Function
    
    lng��ҳID = GetMinMaxTime(0)     '��СסԺ����
    If lng��ҳID = 0 Then GetPatiHospitalzedDateRange = True: Exit Function
    
    
    If lng��ҳID = Val(NVL(mrsInfo!��ҳID)) Then
        dtBeginDate = mrsInfo!��Ժ����
        
        If Not IsNull(mrsInfo!��Ժ����) Then
            dtEndDate = mrsInfo!��Ժ����
        Else
            dtEndDate = zlDatabase.Currentdate
        End If
        
        '��Ժʱ���ȱʡ�����һ�ν���ʱ�仹С,��ʼʱ�������һ�ν���ʱ��Ϊ׼
        If IsDate(strDefaultDate) Then    '����:30043
            If Format(dtBeginDate, "yyyy-mm-dd") < strDefaultDate And Format(dtEndDate, "yyyy-mm-dd") > strDefaultDate Then dtBeginDate = CDate(strDefaultDate)
        End If

        GetPatiHospitalzedDateRange = True: Exit Function
    End If
    
    If CStr(lng��ҳID) = strTime Then '�����ǽ���ǰĳ��סԺ����
        strSQL = "Select ��Ժ����,Nvl(��Ժ����,Sysdate) as ��Ժ���� From ������ҳ" & _
                " Where ����ID=[1] And ��ҳID=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID)
        dtBeginDate = rsTmp!��Ժ����
        dtEndDate = rsTmp!��Ժ����

        If IsDate(strDefaultDate) Then
            If Format(dtBeginDate, "yyyy-mm-dd") < strDefaultDate And Format(dtEndDate, "yyyy-mm-dd") > strDefaultDate Then dtBeginDate = CDate(strDefaultDate)
        End If

    End If
    GetPatiHospitalzedDateRange = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



Private Function GetFeeDate(ByRef dtBeginDate As Date, ByRef dtEndDate As Date) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���˵���С��������ʱ��
    '����:dtBeginDate-��ʼʱ��
    '     dtEndDate-����ʱ��
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-06 15:54:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, dtDate As Date
    On Error GoTo errHandle
    
    If mrsFeeList Is Nothing Then Exit Function
    If mrsFeeList.State <> 1 Then Exit Function
    If mrsFeeList.RecordCount = 0 Then GoTo GoEnd:
    mrsFeeList.MoveFirst
    If mty_ModulePara.int����ʱ�� = 0 Then
        dtDate = mrsFeeList!�Ǽ�ʱ��
    Else
        dtDate = mrsFeeList!ʱ��
    End If
    
    dtBeginDate = dtDate: dtEndDate = dtDate
    With mrsFeeList
        Do While Not .EOF
            If mty_ModulePara.int����ʱ�� = 0 Then
                dtDate = mrsFeeList!�Ǽ�ʱ��
            Else
                dtDate = mrsFeeList!ʱ��
            End If
            If dtDate < dtBeginDate Then dtBeginDate = dtDate
            If dtDate > dtEndDate Then dtEndDate = dtDate
            .MoveNext
        Loop
    End With
    mrsFeeList.MoveFirst
GoEnd:
    GetFeeDate = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetMinMaxTime(ByVal bytMode As Byte) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ȡδ������е���С������סԺ����,���ܷ���0
    '���:bytMode,0-��С����,1-������
    '����:סԺ����
    '����:���˺�
    '����:2015-01-06 16:02:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTime As String, varData As Variant
    Dim i As Long, intTime As Integer
    
    On Error GoTo errHandle
        
    strTime = mobjBalanceCon.strTime
    If strTime = "" Then strTime = mobjBalanceAll.strAllTime
    
    varData = Split(strTime, ",")
    For i = 0 To UBound(varData)
        If i = 0 Then intTime = Val(varData(i))
        If bytMode = 0 Then
            If intTime > Val(varData(i)) Then intTime = Val(varData(i))
        Else
            If intTime < Val(varData(i)) Then intTime = Val(varData(i))
        End If
    Next
    GetMinMaxTime = intTime
Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub zlChangeDefaultTime()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ı�ȱʡ��סԺʱ�䷶Χ
    '����:���˺�
    '����:2015-01-06 16:42:36
    '˵����30043
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If opt��Ժ.Value Then txtPatiEnd.Text = txtPatiEnd.Tag: Exit Sub

    txtPatiEnd.Text = Format(zlDatabase.Currentdate - 1, "yyyy-mm-dd")
    If txtPatiEnd.Text < txtPatiBegin.Text Then
        txtPatiEnd.Text = txtPatiEnd.Tag
    End If
    If txtPatiEnd.Text > txtPatiEnd.Tag Then
        txtPatiEnd.Text = txtPatiEnd.Tag
    End If
End Sub
Private Sub RecalcDepositMoney(ByVal bytOperationType As Byte, _
    Optional ByVal dblMoney As Double = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���¼����Ԥ�����
    '���:bytOperationType-��������(0-������г�Ԥ��;1-��ȱʡʹ��Ԥ����;2-��ָ���������Ԥ��(��ʱ���Ⱥ�����̯��;3-ȫ��
    '     dblMoneny-��Ԥ�����
    '����:���˺�
    '����:2015-01-07 14:49:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bytCurFun As Byte  '0-ȫ��Ԥ����;1-�����ʽ������Ԥ��;2-ʹ������Ԥ����;
    Dim dblTotal As Double, i As Long
    
    On Error GoTo errHandle
    
    mBalanceInfor.dbl��Ԥ���ϼ� = 0
    
    Select Case bytOperationType
    Case 0  '0-������г�Ԥ��
        bytCurFun = 0
    Case 1  '1-��ȱʡʹ��Ԥ����
        bytCurFun = 1   '������ʻ���;���ʣ�ȱʡ�����ʽ����ʹ��
        If mEditType = g_Ed_סԺ���� And opt��Ժ.Value Then bytCurFun = 2
        If mEditType = g_Ed_סԺ���� And mty_ModulePara.bln��;������Ԥ�� Then bytCurFun = 2
        If mEditType = g_Ed_������� Or mblnCurMzBalanceNo Then
            Select Case mty_ModulePara.bytMzDeposit '����Ԥ��ȱʡʹ�÷�ʽ
            Case 0 ' 0-ȱʡ��ʹ�ý�;1-�����ʽ��ʹ��Ԥ��;2-ʹ������Ԥ��
                bytCurFun = 0
            Case 1 '1-�����ʽ��ʹ��Ԥ��
                bytCurFun = 1
            Case 2 '2-ʹ������Ԥ��
                bytCurFun = 2
            End Select
        End If
        If mEditType = g_Ed_���½��� Then
            If InStr(lblBalanceType.Caption, "��Ժ") > 0 Then
                bytCurFun = 2
            End If
        End If
        dblMoney = RoundEx(mBalanceInfor.dblδ���ϼ�, 2)
    Case 2 '2-��ָ���������Ԥ��(��ʱ���Ⱥ�����̯��
        bytCurFun = 1
        If dblMoney = 0 Then dblMoney = RoundEx(mBalanceInfor.dblδ���ϼ�, 2)
    Case 3 '3-ȫ��
        bytCurFun = 2
    Case Else
         bytCurFun = 0
    End Select
    
    If dblMoney < 0 Then dblMoney = 0
    With vsDeposit
        dblTotal = 0
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 1) <> "" Then
                    If Val(.TextMatrix(i, .ColIndex("�༭״̬"))) = 0 Then
                        .Cell(flexcpText, i, .ColIndex("��Ԥ��"), i, .ColIndex("��Ԥ��")) = "0.00"
                        .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbBlack
                        Select Case bytCurFun
                            Case 1 '�����ʽ��ʹ��
                                If dblMoney = 0 Then GoTo NextDeposit
                                If Val(.TextMatrix(i, .ColIndex("���"))) <= dblMoney Then
                                      .TextMatrix(i, .ColIndex("��Ԥ��")) = Format(Val(.TextMatrix(i, .ColIndex("���"))), "0.00")
                                Else
                                    .TextMatrix(i, .ColIndex("��Ԥ��")) = Format(dblMoney, "0.00")
                                End If
                                dblTotal = dblTotal + RoundEx(Val(.TextMatrix(i, .ColIndex("��Ԥ��"))), 2)
                                dblMoney = dblMoney - Val(.TextMatrix(i, .ColIndex("��Ԥ��")))
                            Case 2 'ȫ��
                                .TextMatrix(i, .ColIndex("��Ԥ��")) = Format(Val(.TextMatrix(i, .ColIndex("���"))), "0.00")
                                dblTotal = dblTotal + RoundEx(Val(.TextMatrix(i, .ColIndex("��Ԥ��"))), 2)
                            Case Else
                        End Select
                    Else
                        dblTotal = dblTotal + RoundEx(Val(.TextMatrix(i, .ColIndex("��Ԥ��"))), 2)
                        dblMoney = dblMoney - Val(.TextMatrix(i, .ColIndex("��Ԥ��")))
                    End If
                End If
NextDeposit:
            Next
    End With
    mBalanceInfor.dbl��Ԥ���ϼ� = dblTotal
    '0-ҽ��Ԥ����Ϣ��ʾ;1-��ʾ������Ϣ
    Call ShowLedDisplayBank(1)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub ReInitPatiInvoice(Optional blnFact As Boolean = True)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���³�ʼ����Ʊ��Ϣ
    '���:blnFact-�Ƿ�ˢ�·�Ʊ��
    '����:���˺�
    '����:2015-01-07 16:32:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInvoiceFormat As String
    Dim lng����ID As Long, lng��ҳID As Long, intInsure As Integer
    
    intInsure = mYBInFor.intInsure
    
    If Not mrsInfo Is Nothing Then
        If mrsInfo.State = 1 Then
            lng����ID = Val(NVL(mrsInfo!����ID)): lng��ҳID = Val(NVL(mrsInfo!��ҳID))
            intInsure = mYBInFor.intInsure
        End If
    End If
    If mEditType = g_Ed_������� Then
        Call mobjInvoice.zlGetInvoicePreperty(mlngModul, IIf(mty_ModulePara.bytInvoiceKindMZ = 0, 3, 1), lng����ID, lng��ҳID, intInsure, mobjFactProperty, , , 1)
    Else
        Call mobjInvoice.zlGetInvoicePreperty(mlngModul, IIf(mty_ModulePara.bytInvoiceKindZY = 0, 3, 1), lng����ID, lng��ҳID, intInsure, mobjFactProperty, , , 2)
    End If
    If mobjFactProperty.����ʹ����� Then mlng����ID = 0
    If blnFact Then Call RefreshFact
    
    If mEditType = g_Ed_������� Then
        Call ZlShowBillFormat(mty_ModulePara.bytInvoiceKindMZ, lblFormat, mobjFactProperty.��ӡ��ʽ)
    Else
        Call ZlShowBillFormat(mty_ModulePara.bytInvoiceKindZY, lblFormat, mobjFactProperty.��ӡ��ʽ)
    End If
    picFormat.Visible = lblFormat.Visible
    Call ShowYBMoneyOrInvioceFormatInfor
End Sub

Private Function CheckDepositFactValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡԤ����Ʊ��
    '����:������ȡ,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-30 11:14:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card, lng����ID As Long, strInvoice As String
    Dim blnInput As Boolean, blnValid As Boolean
    
    On Error GoTo errHandle
    mlngԤ������ID = 0
    
    mstrDepositInvioce = "": mblnDepositBillPrint = False
    Set objCard = IDKind�Ҳ�.GetCurCard
    If objCard Is Nothing Then CheckDepositFactValied = True: Exit Function
    If objCard.�ӿ���� <= 1 Then CheckDepositFactValied = True: Exit Function
    '�������Ҳ�
    If Val(txtBalance(Idx_�Ҳ�).Text) = 0 Then CheckDepositFactValied = True: Exit Function
    
    If mobjInvoice.zlGetInvoicePreperty(mlngModul, EM_Ԥ���վ�, mPatiInfor.lng����ID, mPatiInfor.lng��ҳID, 0, mobjDepositFactProperty, , objCard.�ӿ���� = 2) = False Then Exit Function
    
    Select Case mty_ModulePara.bytԤ��Ʊ�ݴ�ӡ
    Case 0 '����ӡ
        CheckDepositFactValied = True: Exit Function
    Case 1 '�Զ���ӡ
        mblnDepositBillPrint = True
    Case 2 'ѡ���Ƿ��ӡ
        If MsgBox("�Ƿ��ӡԤ��Ʊ��?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) <> vbYes Then CheckDepositFactValied = True: Exit Function
        mblnDepositBillPrint = True
    End Select
    
    If mobjDepositFactProperty.�ϸ���� = False Then
        '�п����ǵ�һ��ʹ��
        Do
            blnInput = False
            '���ϸ����ʱֱ�Ӵӱ��ض�ȡ
            strInvoice = UCase(zlDatabase.GetPara("��ǰԤ��Ʊ�ݺ�", glngSys, mlngModul, ""))
            
            If strInvoice = "" Then
                strInvoice = UCase(InputBox("û���ҵ����õ�Ԥ��Ʊ�ݵ����Ʊ�ݺ��룬�޷�ȷ����Ҫʹ�õĿ�ʼƱ�ݺš�" & _
                                vbCrLf & "�����뽫Ҫʹ�õ�Ԥ��Ʊ�ݵĿ�ʼƱ�ݺ��룺", gstrSysName, _
                                "", Me.Left + 1500, Me.Top + 1500))
                blnInput = True
            Else
                strInvoice = zlCommFun.IncStr(strInvoice)
                strInvoice = UCase(InputBox("��ȷ��ʹ�õ�Ԥ��Ʊ�ݵĿ�ʼƱ�ݺ��룺", gstrSysName, _
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
                    If zlCommFun.ActualLen(strInvoice) <> mobjDepositFactProperty.Ʊ�ų��� Then
                        MsgBox "����Ԥ����Ʊ�ݺ��볤��Ӧ��Ϊ " & mobjDepositFactProperty.Ʊ�ų��� & " λ��", vbInformation, gstrSysName
                    Else
                        blnValid = True
                    End If
                Else
                    blnValid = True
                End If
            End If
        Loop While Not blnValid
        mstrDepositInvioce = strInvoice
        CheckDepositFactValied = True: Exit Function
    End If
    
    Do
        '����Ʊ�����ö�ȡ
        blnInput = False
        If mobjInvoice.zlGetInvoiceGroupID(mlngModul, UserInfo.����, EM_Ԥ���վ�, _
            mobjDepositFactProperty.ʹ�����, lng����ID, mobjDepositFactProperty.��������ID, lng����ID, 1, strInvoice) = False Then Exit Function
        If lng����ID <= 0 Then
            Select Case lng����ID
                Case 0 '����ʧ��
                Case -1
                    If Trim(mobjDepositFactProperty.ʹ�����) = "" Then
                        MsgBox "��û�����ú͹��õ�Ԥ��Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                    Else
                        MsgBox "��û�����ú͹��õġ�" & mobjFactProperty.ʹ����� & "��Ԥ��Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                    End If
                    Exit Function
                Case -2
                    If Trim(mobjFactProperty.ʹ�����) = "" Then
                        MsgBox "���صĹ���Ԥ��Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                    Else
                        MsgBox "���صĹ���Ԥ��Ʊ�ݵġ�" & mobjFactProperty.ʹ����� & "��Ԥ��Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                    End If
                    Exit Function
                    strInvoice = ""
            End Select
        End If
        If Not mobjInvoice.zlGetNextBill(mlngModul, lng����ID, strInvoice) Then Exit Function
        
        If strInvoice = "" Then
            '�����;���ÿ���ĺ���,�������δ����,����һ�����ѳ�����Χ
            strInvoice = UCase(InputBox("�޷�����Ʊ�����������ȡ��Ҫʹ��Ԥ��Ʊ�ݵĿ�ʼƱ�ݺţ�" & _
                            vbCrLf & "�������뽫Ҫʹ�õ�Ʊ�ݺ��룺", gstrSysName, _
                            "", Me.Left + 1500, Me.Top + 1500))
            blnInput = True
        Else
            strInvoice = UCase(InputBox("��ȷ��ʹ��ʹ��Ԥ��Ʊ�ݵ�Ʊ�ݺ��룺", gstrSysName, _
                            strInvoice, Me.Left + 1500, Me.Top + 1500))
            blnInput = True
        End If
        
        '�û�ȡ������,����ӡ
        If strInvoice = "" Then Exit Function
        
        '���������Ч��
        If blnInput Then
            If mobjInvoice.zlGetInvoiceGroupID(mlngModul, UserInfo.����, EM_Ԥ���վ�, _
                     mobjDepositFactProperty.ʹ�����, lng����ID, mobjDepositFactProperty.��������ID, lng����ID, 1, strInvoice) = False Then Exit Function
            If lng����ID < 0 Then
                MsgBox "�������Ʊ�ݺ��벻�ڵ�ǰ�������ε���Ч���÷�Χ��,���������룡", vbInformation, gstrSysName
            Else
                blnValid = True
            End If
        Else
            blnValid = True
        End If
    Loop While Not blnValid
    mstrDepositInvioce = strInvoice
    mlngԤ������ID = lng����ID
    CheckDepositFactValied = True: Exit Function
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub RefreshFact()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ˢ�½��ʵ�Ʊ�ݺ�
    '����:���˺�
    '����:2015-01-07 17:16:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFactNO As String
    If mobjFactProperty Is Nothing Then Exit Sub
    If mobjFactProperty.��ӡ��ʽ = 0 Then Exit Sub
      
    If Not mobjFactProperty.�ϸ���� Then
        '���ϸ������
        '��ɢ��ȡ��һ������
        txtInvoice.Text = zlCommFun.IncStr(UCase(zlDatabase.GetPara("��ǰ����Ʊ�ݺ�", glngSys, 1137, "")))
        txtInvoice.Tag = txtInvoice.Text
        txtInvoice.SelStart = Len(txtInvoice.Text)
        Exit Sub
    End If
    
    If zlGetInvoiceGroupUseID(mlng����ID, 1, "") = False Then
          txtInvoice.Text = "": txtInvoice.Tag = ""
        Exit Sub
    End If
    
    '�ϸ�ȡ��һ������
    If mobjInvoice.zlGetNextBill(mlngModul, mlng����ID, strFactNO) = False Then strFactNO = ""
    txtInvoice.Text = strFactNO
    
    'Tag�����⣺24363:���˺飺��Ҫ�ǽ���Զ����ɵĺ��Ƿ��û����ģ���Ҫ�����
    '    1.���ĵ�Ʊ�ݺ���Ҫ����Ƿ��ظ����ظ���ֱ�ӷ��ز����ķ�Ʊ��
    '    2.���������������ĵ�����£�����Ƿ��ظ�������ظ����Զ�ȡ��һ�����룡
    txtInvoice.Tag = txtInvoice.Text
    lblFact.Tag = txtInvoice.Tag
    
    If mobjFactProperty.����ʹ����� Then Call zlCheckFactIsEnough
    txtInvoice.SelStart = Len(txtInvoice.Text)
End Sub

Private Function zlGetRedGroupUseID(ByRef lng����ID As Long, _
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
    Dim bytKind As Byte
    On Error GoTo errHandle
    
    If mPatiInfor.int�������� = 1 Then
        bytKind = IIf(mty_ModulePara.bytInvoiceKindMZ = 0, 3, 1)
    Else
        bytKind = IIf(mty_ModulePara.bytInvoiceKindZY = 0, 3, 1)
    End If
    
    If mobjInvoice.zlGetInvoiceGroupID(mlngModul, UserInfo.����, bytKind, _
        mobjRedProperty.ʹ�����, lng����ID, mobjFactProperty.��������ID, lng����ID, intNum, strInvoiceNO) = False Then Exit Function
    
    If lng����ID > 0 Then zlGetRedGroupUseID = True: Exit Function
    
    Select Case lng����ID
        Case 0 '����ʧ��
        Case -1
            If Trim(mobjRedProperty.ʹ�����) = "" Then
                MsgBox "��û�����ú͹��õĽ���Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
            Else
                MsgBox "��û�����ú͹��õġ�" & mobjRedProperty.ʹ����� & "������Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
            End If
            Exit Function
        Case -2
            If Trim(mobjRedProperty.ʹ�����) = "" Then
                MsgBox "���صĹ���Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
            Else
                MsgBox "���صĹ���Ʊ�ݵġ�" & mobjRedProperty.ʹ����� & "������Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
            End If
            Exit Function
        Case -3
            MsgBox "��ǰƱ�ݺ��벻�ڿ����������ε���ЧƱ�ݺŷ�Χ��,���������룡", vbInformation, gstrSysName
            Exit Function
    End Select
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

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
    If mEditType = g_Ed_������� Then
        If mobjInvoice.zlGetInvoiceGroupID(mlngModul, UserInfo.����, IIf(mty_ModulePara.bytInvoiceKindMZ = 0, 3, 1), _
            mobjFactProperty.ʹ�����, lng����ID, mobjFactProperty.��������ID, lng����ID, intNum, strInvoiceNO) = False Then Exit Function
    Else
        If mobjInvoice.zlGetInvoiceGroupID(mlngModul, UserInfo.����, IIf(mty_ModulePara.bytInvoiceKindZY = 0, 3, 1), _
            mobjFactProperty.ʹ�����, lng����ID, mobjFactProperty.��������ID, lng����ID, intNum, strInvoiceNO) = False Then Exit Function
    End If
    If lng����ID > 0 Then zlGetInvoiceGroupUseID = True: Exit Function
    
    Select Case lng����ID
        Case 0 '����ʧ��
        Case -1
            If Trim(mobjFactProperty.ʹ�����) = "" Then
                MsgBox "��û�����ú͹��õĽ���Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
            Else
                MsgBox "��û�����ú͹��õġ�" & mobjFactProperty.ʹ����� & "������Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
            End If
            Exit Function
        Case -2
            If Trim(mobjFactProperty.ʹ�����) = "" Then
                MsgBox "���صĹ���Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
            Else
                MsgBox "���صĹ���Ʊ�ݵġ�" & mobjFactProperty.ʹ����� & "������Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
            End If
            Exit Function
        Case -3
            MsgBox "��ǰƱ�ݺ��벻�ڿ����������ε���ЧƱ�ݺŷ�Χ��,���������룡", vbInformation, gstrSysName
            If txtInvoice.Enabled And txtInvoice.Visible Then txtInvoice.SetFocus
            Exit Function
    End Select
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 


Private Sub zlCheckFactIsEnough(Optional ByVal intInvoicePages As Integer = 0, Optional ByVal blnCheckRed As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鵱ǰƱ���Ƿ�����
    '���:intInvoicePages-��Ҫ�ķ�Ʊ����,���Ϊ0,��ϵͳ��������
    '����:���˺�
    '����:2015-01-07 18:21:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngʣ������ As Long, lngNums As Long
    Dim bytKind As Byte
    If blnCheckRed Then
        If mPatiInfor.int�������� = 1 Then
            bytKind = IIf(mty_ModulePara.bytInvoiceKindMZ = 0, 3, 1)
        Else
            bytKind = IIf(mty_ModulePara.bytInvoiceKindZY = 0, 3, 1)
        End If
    Else
        If mEditType = g_Ed_���ݲ鿴 Or mEditType = g_Ed_ȡ������ Or mEditType = g_Ed_�������� Then Exit Sub
        If mEditType = g_Ed_������� Then
            bytKind = IIf(mty_ModulePara.bytInvoiceKindMZ = 0, 3, 1)
        Else
            bytKind = IIf(mty_ModulePara.bytInvoiceKindZY = 0, 3, 1)
        End If
    End If
    If intInvoicePages <> 0 Then
        If mobjInvoice.zlCheckInvoiceOverplusEnough(bytKind, intInvoicePages, lngʣ������, mlng����ID, mobjFactProperty.ʹ�����) = False Then
            MsgBox "ע��:" & vbCrLf & _
                   "    ��ǰʣ��Ʊ�ݲ���(" & lngʣ������ & ") ,��ǰ��Ҫ" & intInvoicePages & "��Ʊ��,��ע�������Ʊ!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
        End If
    Else
        If mobjInvoice.zlCheckInvoiceOverplusEnough(bytKind, mty_ModulePara.int����ʣ��Ʊ������, lngʣ������, mlng����ID, mobjFactProperty.ʹ�����) = False Then
            MsgBox "ע��:" & vbCrLf & _
                   "    ��ǰʣ��Ʊ��(" & lngʣ������ & ") С���˱���������(" & mty_ModulePara.int����ʣ��Ʊ������ & "),��ע�������Ʊ!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
        End If
    End If
End Sub


Public Function Chk�������(lng����ID As Long, lng��ҳID As Long) As Boolean
'���ܣ��жϲ����Ƿ������
'������
    On Error GoTo errH
    
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select Nvl(��˱�־,0) as ��˱�־" & _
        " From ������ҳ Where ����ID=[1] And ��ҳID=[2]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng����ID, lng��ҳID)
    
    '49501
    If gTy_System_Para.byt������˷�ʽ = 0 Then
        Chk������� = (rsTmp!��˱�־ >= 1)
    Else
        Chk������� = (rsTmp!��˱�־ > 1)
    End If

    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub Led��ӭ��Ϣ()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:Led��ʼ��
    '����:���˺�
    '����:2015-01-08 10:54:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    If mEditType = g_Ed_���ݲ鿴 Or Not gblnLED Then Exit Sub
    If mty_ModulePara.blnLedWelcome Then
        zl9LedVoice.Reset com
        zl9LedVoice.Speak "#1"
        zl9LedVoice.Init UserInfo.��� & "�� Ϊ������", mlngModul, gcnOracle
    End If
    If mrsInfo Is Nothing Then Exit Sub
    If mrsInfo.State <> 1 Then Exit Sub
    If mrsInfo.EOF Then Exit Sub
    zl9LedVoice.DisplayPatient txtPatient.Text & " " & txtSex.Text & " " & txtOld.Text, Val("" & mrsInfo!����ID)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub InitLed()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��Led
    '����:���˺�
    '����:2015-01-08 14:28:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mEditType = g_Ed_���ݲ鿴 Or Not gblnLED Then Exit Sub
    zl9LedVoice.Reset com
    zl9LedVoice.Init UserInfo.��� & "��Ϊ������", mlngModul, gcnOracle
End Sub

Private Function GetPatiState(lng����ID As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����״̬˵��
    '���:lng����ID-����ID
    '����:
    '����:���ز���״̬˵��
    '     ��ͨ��Ժ,������Ժ,ҽ����Ժ;��ͨ��Ժ,���۳�Ժ,ҽ����Ժ;������ͨ,��������,����ҽ��
    '����:���˺�
    '����:2015-01-08 10:59:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng��ҳID As Long, str˵�� As String
    
    On Error GoTo errHandle
     
    If mrsInfo Is Nothing Then Exit Function
    If mrsInfo.State <> 1 Then Exit Function
    If mrsInfo.EOF Then Exit Function
    
    If mrsInfo!��ҳID = 0 Or mYBInFor.bytMCMode = 1 Then
        If mYBInFor.intInsure = 0 Then
            GetPatiState = "������ͨ"
        Else
            GetPatiState = "����ҽ��"
        End If
        Exit Function
    End If
    
    If NVL(mrsInfo!��������, 0) = 1 Then
        str˵�� = "��������"
        If NVL(mrsInfo!״̬, 0) = 3 Then
            str˵�� = str˵�� & "(Ԥ��Ժ)"
        End If
        GetPatiState = str˵��
        Exit Function
    End If

    If mYBInFor.intInsure <> 0 Then
        str˵�� = "ҽ��"
    ElseIf NVL(mrsInfo!��������, 0) = 2 Then
        str˵�� = "����"
    Else
        str˵�� = "��ͨ"
    End If
    
    If mEditType = g_Ed_������� Or mblnCurMzBalanceNo Then
        If Is��������(mrsInfo!����ID, lng��ҳID) Then
            str˵�� = "��������"
        Else
            str˵�� = "����" & str˵��
        End If
    Else
        If IsNull(mrsInfo!��Ժ����) Then
            str˵�� = str˵�� & "��Ժ"
        Else
            str˵�� = str˵�� & "��Ժ"
        End If
    End If
    
    If NVL(mrsInfo!״̬, 0) = 3 Then
        str˵�� = str˵�� & "(Ԥ��Ժ)"
    End If
    
    GetPatiState = str˵��
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Function Is��������(ByVal lng����ID As Long, ByRef lng��ҳID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鵱ǰ�����Ƿ����������۲��˷����ڼ�
    '���:lng����ID
    '����:lng��ҳID-���ص�ǰ����ID(�ڼ������۵�)
    '����:����������,�򷵻�true,���򷵻�False
    '����:���˺�
    '����:2015-01-08 11:23:41
    '����:45302
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    Dim strSQL As String, dtStartDate As Date, dtEndDate As Date
    Dim strʱ�� As String, strCond As String, rsTemp As ADODB.Recordset
    strʱ�� = IIf(mty_ModulePara.int����ʱ�� = 0, "A.�Ǽ�ʱ��", "A.����ʱ��")
    strCond = "": dtStartDate = CDate("1901-01-01"): dtEndDate = dtStartDate
    If Not mobjBalanceCon.dtBeginDate = CDate("0:00:00") Then
        strCond = " " & strʱ�� & " Between [3] And [4]"
        dtStartDate = CDate(Format(mobjBalanceCon.dtBeginDate, "yyyy-MM-dd 00:00:00"))
        dtEndDate = CDate(Format(mobjBalanceCon.dtEndDate, "yyyy-MM-dd 23:59:59"))
    End If
    
    gstrSQL = "" & _
    "Select A.��ҳid " & _
    "   From ������ҳ A, " & _
    "        (Select Min(" & strʱ�� & ") As ��С����ʱ��, Max(" & strʱ�� & " ) ������ʱ�� " & _
    "          From ������ü�¼ A " & _
    "          Where  ����id = [1] " & strCond & ") B " & _
    "   Where A.����id = [1] And A.�������� = 1  " & _
    "       And (B.��С����ʱ�� Between A.��Ժ���� And Nvl(A.��Ժ����, Sysdate) Or " & _
    "                B.������ʱ�� Between A.��Ժ���� And Nvl(A.��Ժ����, Sysdate) Or " & _
    "                A.��Ժ���� Between B.��С����ʱ�� And B.������ʱ�� Or " & _
    "                Nvl(A.��Ժ����, Sysdate) Between B.��С����ʱ�� And B.������ʱ��)" & _
    "   Order by ��ҳID Desc"
    
    On Error GoTo errHandle
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng����ID, dtStartDate, dtEndDate)
    If rsTemp.EOF Then rsTemp.Close: Set rsTemp = Nothing: Exit Function
    lng��ҳID = Val(NVL(rsTemp!��ҳID))
    rsTemp.Close: Set rsTemp = Nothing
    Is�������� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub InitOldOneCardInfor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����һ��ͨ��Ϣ
    '����:���˺�
    '����:2015-01-08 12:02:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mEditType = g_Ed_���ݲ鿴 Then Exit Sub
    Set mOldOneCard.rsOneCard = GetOneCard
    With mOldOneCard
        .blnOneCard = .rsOneCard.RecordCount > 0
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Function Init���㷽ʽ() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��������Ϣ
    '����:���˺�
    '����:2015-01-08 12:06:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String
    Dim objCards As Cards, objCard As Card
    Dim objPayCards As Cards, i As Long
    Dim blnOnlyDeposit As Boolean
    
    On Error GoTo errHandle

    mstrȱʡ���㷽ʽ = ""
    If mEditType = g_Ed_���ݲ鿴 Then Init���㷽ʽ = True: Exit Function
    
    Set objCards = New Cards: Set objPayCards = New Cards
    '����:1-�ֽ���㷽ʽ,2-������ҽ������,3-ҽ�������ʻ�,4-ҽ������ͳ��,5-���տ���, _
    '     6-�����ۿ�,7-һ��ͨ����,8-���㿨����
    
    If InStr(1, mstrPrivs, ";���ô��۽���;") = 0 Then
        strTmp = "1,2,3,4,5,9,7,8"
    Else
        strTmp = "1,2,3,4,5,6,9,7,8"
    End If
    
    If InStr(1, mstrPrivs, ";�����ֽ����;") = 0 Then
        blnOnlyDeposit = True
    End If
    
    Set mrs���㷽ʽ = Get���㷽ʽ("����", strTmp)
    If mrs���㷽ʽ.RecordCount = 0 Then
        MsgBox "���ʳ���û�п��õĽ��㷽ʽ�����ȵ����㷽ʽ����������", vbInformation, gstrSysName
        Exit Function
        
    End If
     
    '��|ȫ��|ˢ����־|�����ID(���ѿ����)|����|�Ƿ����ѿ�|���㷽ʽ|�Ƿ�����|�Ƿ����ƿ�;��
    If Not gobjSquare Is Nothing Then
        '0-����ҽ�ƿ�;1-���õ�ҽ�ƿ�,2-���д��������˻��������� 3-���õ������˻���ҽ�ƿ�
       Set objCards = gobjSquare.objSquareCard.zlGetCards(3)   '��ȡ��Ч�������ʻ��Ḷ
    End If
    If blnOnlyDeposit Then
        mrs���㷽ʽ.Filter = "����=3 Or ����=4"
    Else
        mrs���㷽ʽ.Filter = "����<7"
    End If
    With mrs���㷽ʽ
        If .RecordCount <> 0 Then .MoveFirst
        i = 1
        Do While Not .EOF
            If (InStr(",3,4,", "," & Val(NVL(!����)) & ",") = 0) And Val(NVL(!Ӧ����)) <> 1 Then
                If InStr("|" & mty_ModulePara.str�ѻ�ҽ�����㷽ʽ & "|", "|" & !���� & "|") > 0 Then
                    '�ѻ�ҽ�����㷽ʽ
                    With vsBlance
                        If .TextMatrix(.Rows - 1, .ColIndex("���㷽ʽ")) <> "" Then
                            .Rows = .Rows + 1
                        End If
                        .TextMatrix(.Rows - 1, .ColIndex("���㷽ʽ")) = NVL(mrs���㷽ʽ!����)
                        .TextMatrix(.Rows - 1, .ColIndex("��������")) = 2
                        .TextMatrix(.Rows - 1, .ColIndex("����״̬")) = 0
                        .TextMatrix(.Rows - 1, .ColIndex("�༭״̬")) = 1
                        .RowData(.Rows - 1) = 110
                    End With
                    mlngBalanceRows = vsBlance.Rows
                Else
                    Set objCard = New Card
                    objCard.�ӿ���� = -1 * i
                    objCard.�ӿڱ��� = !����
                    objCard.���� = !����
                    objCard.���㷽ʽ = !����
                    objCard.�������� = Val(NVL(!����))
                    objCard.���� = True
                    objCard.�Ƿ�ˢ�� = 1
                    objCard.ȱʡ��־ = Val(NVL(!ȱʡ)) = 1
                    objPayCards.Add objCard
                    If objCard.ȱʡ��־ Then
                        mstrȱʡ���㷽ʽ = objCard.���㷽ʽ
                    End If
                    i = i + 1
                End If
            End If
            .MoveNext
        Loop
        .Filter = 0
    End With
    
    If InStr(";" & mstrPrivsCard & ";", ";�����ӿ�����;") > 0 Then
        mrs���㷽ʽ.Filter = "����>=7 and ����<9" 'һ��ͨ����
        With mrs���㷽ʽ
            If .RecordCount <> 0 Then .MoveFirst
            Do While Not .EOF
                For Each objCard In objCards
                    If objCard.���㷽ʽ = NVL(!����) Then
                        '�ҵ���,����
                        '85565,���ϴ�,2015/7/19:��������
                        objCard.�Ƿ�ˢ�� = True
                        objCard.ȱʡ��־ = Val(NVL(!ȱʡ)) = 1
                        objPayCards.Add objCard
                        If objCard.ȱʡ��־ Then
                            mstrȱʡ���㷽ʽ = objCard.���㷽ʽ
                        End If
                        Exit For
                    End If
                Next
                .MoveNext
            Loop
            .Filter = 0
        End With
    End If
    
    mrs���㷽ʽ.Filter = 0
    mblnNotChange = True
    Set IDKindPaymentsType.Cards = objPayCards
    If objPayCards.Count = 0 And blnOnlyDeposit = False Then
        mblnNotChange = True
        MsgBox "Ԥ������û�п��õĽ��㷽ʽ,���ȵ����㷽ʽ���������á�", vbExclamation, gstrSysName
        mblnUnload = True: Exit Function
    End If
    mblnNotChange = False
    Init���㷽ʽ = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub Load�Ҳ���(ByVal bytFun As Byte, ByVal str�Ҳ����� As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����Ҳ���
    '���:bytFun-0-ֻ���Ҳ�;1-����Ԥ��
    '����:���˺�
    '����:2015-01-09 15:13:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCards As Cards, objCard As Card
     
    On Error GoTo errHandle
    If mEditType = g_Ed_���ݲ鿴 Then Exit Sub
    
    On Error GoTo errHandle
    
    Set objCards = New Cards
    Set objCard = New Card
    objCard.�ӿ���� = 1
    objCard.�ӿڱ��� = 1
    objCard.���� = IIf(str�Ҳ����� = "", "�Ҳ�", str�Ҳ�����)
    objCard.���㷽ʽ = objCard.����
    objCard.�������� = 0
    objCard.���� = True
    '85565,���ϴ�,2015/7/10:��������
    objCard.�Ƿ�ˢ�� = True
    objCards.Add objCard
    If bytFun <> 0 Then
        Set objCard = New Card
        objCard.�ӿ���� = 2
        objCard.�ӿڱ��� = 2
        objCard.���� = "����Ԥ��"
        objCard.���㷽ʽ = objCard.����
        objCard.�������� = 0
        objCard.���� = True
        '85565,���ϴ�,2015/7/10:��������
        objCard.�Ƿ�ˢ�� = True
        objCards.Add objCard
        
        Set objCard = New Card
        objCard.�ӿ���� = 3
        objCard.�ӿڱ��� = 3
        objCard.���� = "סԺԤ��"
        objCard.���㷽ʽ = objCard.����
        objCard.�������� = 0
        objCard.���� = True
        '85565,���ϴ�,2015/7/10:��������
        objCard.�Ƿ�ˢ�� = True
        
        objCards.Add objCard
    End If
    mblnNotChange = True
    Set IDKind�Ҳ�.Cards = objCards
    mblnNotChange = False
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    
End Sub

Private Function LoadBalanceBill() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ؽ��ʵ��������Ϣ
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-08 14:30:45
    '---------------------------------------------------------------------------------------------------------------------------------------------

    On Error GoTo errHandle
    If mEditType = g_Ed_������� Or mEditType = g_Ed_סԺ���� Then
        'ִ�н���
        Call NewBill
        If mlngPatientID <> 0 Then
            txtPatient.Text = "-" & mlngPatientID
            mobjBalanceCon.strTime = mstr��ҳId
            Call txtPatient_KeyPress(vbKeyReturn)
            If Val(mstr��ҳId) = "0" Then cmdYB.Enabled = True
            If mrsInfo Is Nothing Then mblnUnload = True: Exit Function
            If mrsInfo.State = 0 Then mblnUnload = True: Exit Function
        End If
        Me.Caption = IIf(mEditType = g_Ed_�������, "���ﲡ�˽��ʵ�", "סԺ���˽��ʵ�")
        LoadBalanceBill = True: Exit Function
    End If
    
    Select Case mEditType
    Case g_Ed_ȡ������, g_Ed_��������, g_Ed_��������
        mblnNotChange = True
        chkCancel.Value = 1
        mblnNotChange = False
    Case Else
    End Select

    If Not ReadBalance(mstrInNO) Then mblnUnload = True: Exit Function
    LoadBalanceBill = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function LoadBalancePayData(ByVal lng����ID As Long, ByVal lng����ID As Long, _
    Optional ByVal blnNOMoved As Boolean = False, Optional blnԭ���� As Boolean, Optional blnInsure As Boolean, _
    Optional ByVal intCustomSign As Integer = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����Ѿ�֧��������
    '���:lng����ID-����ID
    '     blnNOMoved-�Ƿ��Ѿ�ת��󱸱�
    '     blnԭ����-��ȡ����ԭ��������
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-08 15:24:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, i As Long, strSQL As String
    Dim dblTotal As Double, strTable As String, blnYB As Boolean
    Dim strCardNo As String, cllBillPro As New Collection
    Dim objCard As Card, bytEdit As Byte
    Dim lng�����ID  As Long, dblMoney As Double
    Dim TyBrushCardInor As TY_BrushCard
    Dim blnAdd As Boolean, intYBpara As Integer
    Dim byt����״̬ As Byte
    Dim dblҽ������ As Double
    Dim intSign As Integer
    Dim blnUnload As Boolean
    Dim blnNoPrepay As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim j As Long
    Dim blnCheck As Boolean
    
    On Error GoTo errHandle
     
    '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
    If zlGetFromIDToBalanceData(lng����ID, blnNOMoved, mrsBalance) = False Then Exit Function
    
    With mrsBalance
        i = 1: blnYB = False
'        vsBlance.Clear 1: vsBlance.Rows = 2
        mBalanceInfor.dbl�Ѹ��ϼ� = 0
        mBalanceInfor.dblҽ��֧���ϼ� = 0
        If Not mEditType = g_Ed_�������� And blnInsure = False Then
            mBalanceInfor.dbl��Ԥ���ϼ� = 0
        End If
        If intCustomSign <> 0 Then
            intSign = intCustomSign
        Else
            intSign = IIf(mEditType = g_Ed_��������, -1, 1)
        End If
        
        Do While Not .EOF
            dblMoney = RoundEx(intSign * Val(NVL(!��Ԥ��)), 6)
            blnAdd = True
            Select Case NVL(!����)
            Case 1 'Ԥ����
                If Not mEditType = g_Ed_�������� Then
                    '���������ڼ���Ԥ����ʱ,�Ѿ���ֵ,ԭ����Ҫ��ԭʼ����ʱ�ĳ�Ԥ��
                    mBalanceInfor.dbl��Ԥ���ϼ� = RoundEx(mBalanceInfor.dbl��Ԥ���ϼ� + dblMoney, 6)
                End If
            Case 2, 3, 5 'ҽ��,һ��ͨ,���ѿ�
                blnAdd = True
                If NVL(!����) = 2 Then
                    If mEditType = g_Ed_�������� Or mEditType = g_Ed_ȡ������ Or blnԭ���� Or chkCancel.Value = 1 Then
                        Select Case Val(NVL(mrsBalance!����))
                        Case 3   '�����ʻ�
                            If mYBInFor.bytMCMode = 1 And Not MCPAR.���ﲡ�˽������� Then
                                blnAdd = False
                            Else
                                intYBpara = IIf(mYBInFor.bytMCMode = 1, support�����������, supportסԺ��������)
                                blnAdd = gclsInsure.GetCapability(intYBpara, lng����ID, mYBInFor.intInsure, NVL(mrsBalance!���㷽ʽ))
                            End If
                        Case 4  'ҽ������
                            intYBpara = IIf(mYBInFor.bytMCMode = 1, support�����������, supportסԺ��������)
                            blnAdd = gclsInsure.GetCapability(intYBpara, lng����ID, mYBInFor.intInsure, NVL(mrsBalance!���㷽ʽ))
                        End Select
                    End If
                    
                    If blnAdd Then
                        mBalanceInfor.dblҽ��֧���ϼ� = RoundEx(mBalanceInfor.dblҽ��֧���ϼ� + dblMoney, 6)
                        If Val(NVL(mrsBalance!����)) = 4 Then
                            dblҽ������ = dblҽ������ + dblMoney
                        End If
                        blnYB = True
                    End If
                End If
                
                If Not blnAdd Then GoTo GoAddEnd:
                
                With vsBlance
                    strCardNo = NVL(mrsBalance!����)
                    lng�����ID = IIf(Val(NVL(mrsBalance!����)) = 5, Val(NVL(mrsBalance!���㿨���)), Val(NVL(mrsBalance!�����ID)))
                    TyBrushCardInor.str���� = strCardNo
                    TyBrushCardInor.str������� = NVL(mrsBalance!�������)
                    TyBrushCardInor.str����ժҪ = NVL(mrsBalance!ժҪ)
                    TyBrushCardInor.str������ˮ�� = NVL(mrsBalance!������ˮ��)
                    TyBrushCardInor.str����˵�� = NVL(mrsBalance!����˵��)
                    TyBrushCardInor.str��չ��Ϣ = ""
                    If Val(NVL(mrsBalance!У�Ա�־)) = 1 And mEditType <> g_Ed_���ݲ鿴 Then
                        Select Case Val(NVL(mrsBalance!����))
                        Case 3 '3-һ��ͨ
                            If MsgBox("����:" & vbCrLf & _
                                       "     ��ʹ�á�" & NVL(mrsBalance!���������, NVL(mrsBalance!���㷽ʽ, "")) & "������֧������ʱʧ��,�Ƿ��������֧��?" & vbCrLf & _
                                       "������:" & Format(dblMoney, "0.00"), vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                                '����һ��֧ͨ���ӿ�
                                Set objCard = IDKindPaymentsType.GetIDKindCard(lng�����ID, CardTypeID)
                                If objCard Is Nothing Then
                                    MsgBox "��ǰվ��δ����:" & NVL(mrsBalance!���������, NVL(mrsBalance!���㷽ʽ, "")) & ",���ڡ����㷽ʽ�����򱾵ز������豸��������������!", vbInformation + vbOKOnly, gstrSysName
                                    Exit Function
                                End If
                                '�ȼ���Ƿ�Ϸ�
                                If CheckThreeSwapValied(objCard, dblMoney, TyBrushCardInor) = False Then Exit Function
                                If ExecuteThreeSwapPayInterface(lng����ID, lng����ID, objCard, dblMoney, cllBillPro, TyBrushCardInor) = False Then Exit Function
                                byt����״̬ = 1
                            Else
                                Exit Function
                            End If
                        Case 4 '4-һ��ͨ(��)
                            
                            If MsgBox("����:" & vbCrLf & _
                                       "     ��ʹ�á�" & NVL(mrsBalance!���������, NVL(mrsBalance!���㷽ʽ, "")) & "������֧������ʱʧ��,�Ƿ��������֧��?" & vbCrLf & _
                                       "������:" & Format(dblMoney, "0.00"), vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                                '����һ��֧ͨ��(�ϰ�)�ӿ�
                                Set objCard = GetOldCard(mrsBalance!���㷽ʽ)
                                If objCard Is Nothing Then
                                    MsgBox "��ǰվ��δ����:" & NVL(mrsBalance!���������, NVL(mrsBalance!���㷽ʽ, "")) & ",���ڡ������������á�����������!", vbInformation + vbOKOnly, gstrSysName
                                    Exit Function
                                End If
                                '1.�ȼ���Ƿ�Ϸ�
                                If CheckOldOneCardIsValied(objCard, dblMoney, TyBrushCardInor) = False Then Exit Function
                                '2.����֧��
                                If ExecuteOldOneCardPayInterface(lng����ID, lng����ID, objCard, dblMoney, TyBrushCardInor, cllBillPro) = False Then Exit Function
                                byt����״̬ = 1
                            Else
                                Exit Function
                            End If
                        End Select
                    End If
                    
                    blnNoPrepay = False
                    blnUnload = False
                    If Val(NVL(mrsBalance!����)) = 3 Then
                        If mEditType = g_Ed_�������� Or chkCancel.Value = 1 Then
                            strSQL = "Select 1 From �����˿���Ϣ Where ����ID=[1]"
                            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
                            If rsTmp.EOF Then
                                blnUnload = False
                                blnNoPrepay = False
                            Else
                                blnNoPrepay = True
                                blnUnload = True
                            End If
                        End If
                    End If
                    
                    If blnUnload = False Then
                        If blnYB Then
                            blnYB = False
                            For j = 1 To .Rows - 1
                                If .TextMatrix(j, .ColIndex("���㷽ʽ")) = NVL(mrsBalance!���㷽ʽ) Then
                                    i = j
                                    blnYB = True
                                End If
                            Next j
                        End If
                        If .TextMatrix(i, .ColIndex("���㷽ʽ")) <> "" And blnYB = False Then
                            .Rows = .Rows + 1
                            i = .Rows - 1
                        End If
                        bytEdit = 0
                        If (mEditType = g_Ed_�������� Or chkCancel.Value = 1) And mEditType <> g_Ed_�������� And mEditType <> g_Ed_ȡ������ Then
                            If Val(NVL(mrsBalance!����)) = 3 Then     'һ��ͨ
                                bytEdit = 2
                            End If
                            If Val(NVL(mrsBalance!����)) = 5 Then bytEdit = 2
                        End If
                        If byt����״̬ <> 1 Then
                            If mEditType = g_Ed_�������� Or chkCancel.Value = 1 Then
                                If mEditType = g_Ed_�������� Then
                                    byt����״̬ = IIf(Val(NVL(mrsBalance!У�Ա�־)) = 1, 0, 1)
                                Else
                                    byt����״̬ = 0
                                End If
                            Else
                                byt����״̬ = 1
                            End If
                        End If
                        '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
                        .TextMatrix(i, .ColIndex("����")) = Val(NVL(mrsBalance!����))
                        .TextMatrix(i, .ColIndex("�����ID")) = lng�����ID
                        .TextMatrix(i, .ColIndex("���ѿ�ID")) = Val(NVL(mrsBalance!���ѿ�ID))
                        .TextMatrix(i, .ColIndex("��������")) = Val(NVL(mrsBalance!����))
                        .TextMatrix(i, .ColIndex("�༭״̬")) = bytEdit   '0-��ֹɾ��;1-����༭���;2-����ɾ��
                        .TextMatrix(i, .ColIndex("����״̬")) = byt����״̬  '�Ƿ��ѽ���:1-�ѽ���;0-δ����
                        .TextMatrix(i, .ColIndex("�Ƿ�����")) = Val(NVL(mrsBalance!�Ƿ�����))
                        .TextMatrix(i, .ColIndex("�Ƿ�ȫ��")) = Val(NVL(mrsBalance!�Ƿ�ȫ��))
                        .TextMatrix(i, .ColIndex("У�Ա�־")) = Val(NVL(mrsBalance!У�Ա�־))
                        .TextMatrix(i, .ColIndex("�Ƿ�����")) = Val(NVL(mrsBalance!�Ƿ�����))
                        .TextMatrix(i, .ColIndex("���������")) = Trim(NVL(mrsBalance!���������))
                        .TextMatrix(i, .ColIndex("���㷽ʽ")) = NVL(mrsBalance!���㷽ʽ)
                        .TextMatrix(i, .ColIndex("������")) = FormatEx(dblMoney, 6, , , 2)
                        .TextMatrix(i, .ColIndex("�������")) = NVL(mrsBalance!�������)
                        .TextMatrix(i, .ColIndex("��ע")) = NVL(mrsBalance!ժҪ)
                        .TextMatrix(i, .ColIndex("������ˮ��")) = NVL(mrsBalance!������ˮ��)
                        .TextMatrix(i, .ColIndex("����˵��")) = NVL(mrsBalance!����˵��)
                        .TextMatrix(i, .ColIndex("����")) = IIf(Val(NVL(mrsBalance!�Ƿ�����)) = 1, String(Len(strCardNo), "*"), strCardNo)
                        .Cell(flexcpData, i, .ColIndex("����")) = NVL(mrsBalance!����)
                        
                        If mEditType = g_Ed_���ݲ鿴 Then
                            If Val(NVL(mrsBalance!У�Ա�־)) = 1 Then    'δִ�гɹ���
                                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbRed
                            ElseIf Val(NVL(mrsBalance!У�Ա�־)) = 2 Then 'ִ�гɹ��ҵ�ǰ���ڲ鿴��
                                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbBlue
                            Else
                                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = Me.ForeColor
                            End If
                        End If
                        
                        mBalanceInfor.dbl�Ѹ��ϼ� = RoundEx(mBalanceInfor.dbl�Ѹ��ϼ� + dblMoney, 6)
                    End If
                End With
GoAddEnd:
        Case Else '0-��ͨ����
            
            If mEditType = g_Ed_�������� Or mEditType = g_Ed_���½��� Or chkCancel.Value = 1 Or blnԭ���� Then
                'ֻ��ȱʡΪ�տ�
                If Val(NVL(!����)) = 1 Then blnAdd = False
            End If
            With vsBlance
                If NVL(mrsBalance!���㷽ʽ) <> "" And (NVL(mrsBalance!����) <> 6 Or mEditType = g_Ed_���ݲ鿴) And blnAdd Then
                    blnCheck = False
                    For j = 1 To .Rows - 1
                        If .TextMatrix(j, .ColIndex("���㷽ʽ")) = NVL(mrsBalance!���㷽ʽ) Then
                            i = j
                            blnCheck = True
                        End If
                    Next j
                     If .TextMatrix(i, .ColIndex("���㷽ʽ")) <> "" And NVL(mrsBalance!���㷽ʽ) <> "" And blnCheck = False Then
                         .Rows = .Rows + 1
                         i = .Rows - 1
                     End If
                     bytEdit = 0
                     If mEditType = g_Ed_ȡ������ Or mEditType = g_Ed_�������� Or chkCancel.Value = 1 Then bytEdit = 2
                    
                     If mEditType = g_Ed_�������� Or chkCancel.Value = 1 Then
                         byt����״̬ = 0
                     Else
                         byt����״̬ = 1
                     End If
                     '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
                     .TextMatrix(i, .ColIndex("����")) = Val(NVL(mrsBalance!����))
                     .TextMatrix(i, .ColIndex("�����ID")) = lng�����ID
                     .TextMatrix(i, .ColIndex("���ѿ�ID")) = Val(NVL(mrsBalance!���ѿ�ID))
                     .TextMatrix(i, .ColIndex("��������")) = Val(NVL(mrsBalance!����))
                     .TextMatrix(i, .ColIndex("�༭״̬")) = bytEdit   '0-��ֹɾ��;1-����༭���;2-����ɾ��
                     .TextMatrix(i, .ColIndex("����״̬")) = byt����״̬  '�Ƿ��ѽ���:1-�ѽ���;0-δ����
                     .TextMatrix(i, .ColIndex("�Ƿ�����")) = Val(NVL(mrsBalance!�Ƿ�����))
                     .TextMatrix(i, .ColIndex("�Ƿ�ȫ��")) = Val(NVL(mrsBalance!�Ƿ�ȫ��))
                     .TextMatrix(i, .ColIndex("У�Ա�־")) = Val(NVL(mrsBalance!У�Ա�־))
                     .TextMatrix(i, .ColIndex("�Ƿ�����")) = Val(NVL(mrsBalance!�Ƿ�����))
                     .TextMatrix(i, .ColIndex("���������")) = Trim(NVL(mrsBalance!���������))
                     
                     .TextMatrix(i, .ColIndex("���㷽ʽ")) = NVL(mrsBalance!���㷽ʽ)
                     .TextMatrix(i, .ColIndex("������")) = FormatEx(intSign * Val(NVL(mrsBalance!��Ԥ��)), 6, , , 2)
                     .TextMatrix(i, .ColIndex("�������")) = NVL(mrsBalance!�������)
                     .TextMatrix(i, .ColIndex("��ע")) = NVL(mrsBalance!ժҪ)
                     .TextMatrix(i, .ColIndex("������ˮ��")) = NVL(mrsBalance!������ˮ��)
                     .TextMatrix(i, .ColIndex("����˵��")) = NVL(mrsBalance!����˵��)
                     .TextMatrix(i, .ColIndex("����")) = IIf(Val(NVL(mrsBalance!�Ƿ�����)) = 1, String(Len(strCardNo), "*"), strCardNo)
                     .Cell(flexcpData, i, .ColIndex("����")) = NVL(mrsBalance!����)
                     mBalanceInfor.dbl�Ѹ��ϼ� = RoundEx(mBalanceInfor.dbl�Ѹ��ϼ� + intSign * Val(NVL(mrsBalance!��Ԥ��)), 6)
                 End If
            End With
        End Select
        .MoveNext
        Loop
    End With
    
    If mEditType = g_Ed_�������� Then
        strSQL = "Select 1 From �����˿���Ϣ Where ����ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mBalanceInfor.lng����ID)
        If rsTmp.EOF Then
            blnNoPrepay = False
        Else
            blnNoPrepay = True
        End If
    End If
    
    If blnNoPrepay Then
        mBalanceInfor.dbl��Ԥ���ϼ� = 0
        chkDeposit.Enabled = False
    End If
    
    If vsBlance.TextMatrix(vsBlance.Rows - 1, vsBlance.ColIndex("���㷽ʽ")) = "" Then
        vsBlance.RemoveItem vsBlance.Rows - 1
    End If
    
    mrsBalance.Filter = "���� = 3 Or ���� = 4"
    If mrsBalance.EOF Then
        blnCheck = True
        Do While blnCheck = True
            blnCheck = False
            For i = 1 To vsBlance.Rows - 1
                If Val(vsBlance.TextMatrix(i, vsBlance.ColIndex("��������"))) = 3 Or Val(vsBlance.TextMatrix(i, vsBlance.ColIndex("��������"))) = 4 Then
                    mBalanceInfor.dbl�Ѹ��ϼ� = RoundEx(mBalanceInfor.dbl�Ѹ��ϼ� - Val(vsBlance.TextMatrix(i, vsBlance.ColIndex("������"))), 6)
                    vsBlance.RemoveItem i
                    blnCheck = True
                    Exit For
                End If
            Next i
        Loop
    End If
    mrsBalance.Filter = ""

    mblnNotChange = True
    txtBalance(Idx_��Ԥ��).Text = Format(mBalanceInfor.dbl��Ԥ���ϼ�, "0.00")
    chkDeposit.Tag = mBalanceInfor.dbl��Ԥ���ϼ�
    chkDeposit.Value = 0
    If mBalanceInfor.dbl��Ԥ���ϼ� <> 0 Then chkDeposit.Value = 1
    mBalanceInfor.dblδ���ϼ� = RoundEx(mBalanceInfor.dbl��ǰ���� - mBalanceInfor.dbl�Ѹ��ϼ�, 5)
    lblҽ������.Caption = "ͳ��֧��:" & Format(dblҽ������, "0.00")
    mblnNotChange = False
    LoadBalancePayData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetOldCard(ByVal str���㷽ʽ As String) As Card
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݽ��㷽ʽ,��ȡ��һ��ͨ�Ŀ�����
    '����:���˺�
    '����:2015-01-08 18:05:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card, objCards As Cards
    
    Set objCards = IDKindPaymentsType.Cards
    For Each objCard In objCards
        If objCard.���㷽ʽ = str���㷽ʽ And objCard.�������� = 7 Then
            GetOldCard = objCard: Exit Function
        End If
    Next
    Set GetOldCard = Nothing
End Function
Private Sub ClearCustomType()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Զ���������ر���
    '����:���˺�
    '����:2015-01-26 17:16:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim tyBalanceInfor As TY_Balance_Infor
    Dim tyYBInFor As TY_YBInfor, tyPatiInfor As ty_Pati_Infor
    
    On Error GoTo errHandle
        
    mPatiInfor = tyPatiInfor
    Set mobjBalanceCon = New clsBalanceCon    '��ʼ������
    Set mobjBalanceAll = New clsBalanceAllCon
    
    mBalanceInfor = tyBalanceInfor
    mYBInFor = tyYBInFor
    mPatiInfor = tyPatiInfor '��ղ�����Ϣ
    '�ϰ�һ��ͨ
    With mOldOneCard
        .strOneCard = ""
    End With
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function ReadBalance(strNO As String, Optional blnInputNo As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�鿴������ʱ,��ȡ����ʾ���ʵ�
    '���:strNo-���ʵ��ź�
    '     blnInputNo-���뵥�ݺŽ�������
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-08 14:43:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strFullNO As String, lng����ID As Long
    Dim lngID As Long, i As Long, j As Long, lngDefault As Long
    Dim strSQL As String, dMax As Date, dMin As Date, blnUndo As Boolean
    Dim curTmp As Currency, curMoney As Currency, curDeposit As Currency
    Dim lngMaxLength As Long, lngP As Long, lng����ID As Long
    Dim rsUnit As ADODB.Recordset, rsFee As New ADODB.Recordset
    Dim strTable As String, lng��ҳID As Long
    Dim str��ҳIds As String, rsTmp As ADODB.Recordset
    Dim strOper As String, vDate As Date
    

    
    On Error GoTo errH
    Call ClearCustomType
    
    '��������
    strFullNO = GetFullNO(strNO, 15)
     
    strSQL = "" & _
    "   Select A.ID,A.ʵ��Ʊ��,A.����ID,B.�����,B.סԺ��,b.��ǰ����,B.��ǰ����ID,B.�ѱ�,B.����,B.�Ա�,B.����, " & _
    "          A.�շ�ʱ��,A.��ʼ����,A.��������,A.��ע,A.ԭ��,A.����״̬,A.��������,A.סԺ����,A.���ʽ��, " & _
    "          nvl(A.��ҳID,nvl(B.��ҳID,0)) as ��ҳID,B.��Ժ,nvl(A.��;����,0) as ��;����,A.��¼״̬" & _
    "   From ���˽��ʼ�¼ A,������Ϣ B" & _
    "   Where A.����ID=B.����ID(+) " & _
    "       And A.NO=[1] And A.��¼״̬ " & IIf(mblnViewCancel, "= 2", "In(1,3)")
    
    If mblnNOMoved Then strSQL = Replace(strSQL, "���˽��ʼ�¼", "H���˽��ʼ�¼")
    
    strSQL = _
    "Select A.ID,A.ʵ��Ʊ�� as Ʊ�ݺ�,A.����ID,A.�����, " & _
    "       nvl(D.סԺ��,A.סԺ��) as סԺ��, Nvl(D.��Ժ����,A.��ǰ����)  as ��ǰ����, " & _
    "       Nvl(E.����,C.����) as ��ǰ����,A.��Ժ," & _
    "       Nvl(D.�ѱ�,A.�ѱ�) as �ѱ�,nvl(D.����,A.����) as ����,nvl(D.�Ա�,A.�Ա�) as �Ա�,nvl(D.����,A.����) as ����, " & _
    "       A.�շ�ʱ��,A.��ʼ����,A.��������,A.��ע,A.ԭ��,A.����״̬,A.��������,A.סԺ����,A.���ʽ��,A.��ҳID,A.��;����,A.��¼״̬" & _
    " From (" & strSQL & ") A,���ű� C,������ҳ D,���ű� E" & _
    " Where  A.��ǰ����ID=C.ID(+) And D.��Ժ����ID=E.ID(+)" & _
    "       And A.����ID=D.����ID(+) And A.��ҳID =D.��ҳID(+) "
        
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strFullNO)
    If rsTemp.EOF Then
        MsgBox "û�з��ָý��ʵ���,�����Ѿ����ϣ�", vbInformation, gstrSysName
        Exit Function
    End If
    If blnInputNo = True And Val(NVL(rsTemp!��¼״̬)) <> 1 Then
        MsgBox "�ý��ʵ���Ϊ�Ѿ��������ϣ������ٽ������ϲ�����", vbInformation, gstrSysName
        Exit Function
    End If
    If Not GetMinMaxDate(rsTemp!ID, dMin, dMax, mblnNOMoved) Then
        MsgBox "�ý��ʵ������ݲ���ȷ��û�з��ֽ��ʵķ�����ϸ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    If mEditType = g_Ed_�������� And Val(NVL(rsTemp!����״̬)) <> 1 Then
        MsgBox "�ý��ʵ��ݲ�Ϊ�쳣���ݣ������������ϣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    If mEditType = g_Ed_���½��� And Val(NVL(rsTemp!����״̬)) <> 1 Then
        MsgBox "�ý��ʵ��ݲ�Ϊ�쳣���ݣ��������½��ʣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    If mEditType = g_Ed_ȡ������ And Val(NVL(rsTemp!����״̬)) <> 1 Then
        MsgBox "�ý��ʵ��ݲ�Ϊ�쳣���ݣ�����ȡ�����ʣ�", vbInformation, gstrSysName
        Exit Function
    End If
    If mEditType = g_Ed_�������� And Val(NVL(rsTemp!����״̬)) = 1 Then
        MsgBox "�ý��ʵ���Ϊ�쳣���ݣ����ܽ������ϣ�", vbInformation, gstrSysName
        Exit Function
    End If
    If mEditType = g_Ed_�������� And Val(NVL(rsTemp!��¼״̬)) <> 1 Then
        MsgBox "�ý��ʵ���Ϊ�Ѿ��������ϣ������ٽ������ϲ�����", vbInformation, gstrSysName
        Exit Function
    End If
       
    
    lng����ID = Val(NVL(rsTemp!ID))
    cboNO.Text = strFullNO
    
    If mEditType = g_Ed_�������� Then
        If CheckExistsGathering(cboNO.Text) Then
            MsgBox "�ý��ʵ��ݴ����ѽɿ��Ӧ�տ��¼�����˿����ִ�����ϡ�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    txtInvoice.Text = NVL(rsTemp!Ʊ�ݺ�)

    lng����ID = Val(NVL(rsTemp!����ID))
    lng��ҳID = Val(NVL(rsTemp!��ҳID))
    
    If mEditType = g_Ed_�������� Then
        '����Ȩ��
        If Not ReadBillInfo(2, cboNO.Text, -1, strOper, vDate) Then
            Exit Function
        End If

        If Not BillOperCheck(7, strOper, vDate, "����") Then
            Exit Function
        End If
    End If
    
    mobjBalanceAll.strAllTime = NVL(rsTemp!סԺ����)
    mblnCurMzBalanceNo = False
    If Val(NVL(rsTemp!��������)) = 0 Then
        Me.Caption = gstrUnitName & "���˽��ʵ�"
        If mobjBalanceAll.strAllTime = "" Then mobjBalanceAll.strAllTime = GetFromalanceIDToPatiNum(lng����ID, lng��ҳID)
    ElseIf Val(NVL(rsTemp!��������)) = 1 Then
        Me.Caption = gstrUnitName & "���ﲡ�˽��ʵ�"
        mobjBalanceAll.strAllTime = "": lng��ҳID = 0
        mblnCurMzBalanceNo = True
    Else
        Me.Caption = gstrUnitName & "סԺ���˽��ʵ�"
        If mobjBalanceAll.strAllTime = "" Then mobjBalanceAll.strAllTime = GetFromalanceIDToPatiNum(lng����ID, lng��ҳID)
    End If
    mobjBalanceCon.strTime = mobjBalanceAll.strAllTime
    mBalanceInfor.strNO = strFullNO
    With mPatiInfor
        .lng����ID = lng����ID
        .lng��ҳID = lng��ҳID
        .str���� = NVL(rsTemp!����)
        .str�Ա� = NVL(rsTemp!�Ա�)
        .str���� = NVL(rsTemp!����)
        .bln��Ժ = Val(NVL((rsTemp!��Ժ))) <> 1
        .int�������� = Val(NVL(rsTemp!��������))
    End With
    
    With mBalanceInfor
        .strNO = strFullNO
        .blnSaveBill = IIf(mEditType = g_Ed_�������� Or blnInputNo, False, True)
        If mblnViewCancel And mEditType <> g_Ed_���ݲ鿴 Then
            .lng����ID = lng����ID
            .lng����ID = zlGetFormerBalanceID(mBalanceInfor.strNO)
        Else
            .lng����ID = 0
            .lng����ID = lng����ID
        End If
        .dtBalanceDate = CDate(Format(rsTemp!�շ�ʱ��, "yyyy-mm-dd hh:MM:SS"))
    End With
    
    If mEditType <> g_Ed_���ݲ鿴 Then
        mYBInFor.intInsure = BalanceExistInsure(strNO, mYBInFor.bytMCMode)
        If mYBInFor.intInsure <> 0 Then
            Call InitInsurePara(mPatiInfor.lng����ID, mYBInFor.intInsure)
        End If
    End If
    
    If mEditType = g_Ed_���½��� Or mEditType = g_Ed_ȡ������ Then
        If Val(NVL(rsTemp!��������)) = 0 Then
            If zlStr.IsHavePrivs(mstrPrivs, "������ý���") = False And zlStr.IsHavePrivs(mstrPrivs, "סԺ���ý���") = False Then
                MsgBox "��û�н���Ȩ�ޣ����ܽ��н��ʲ�����", vbInformation, gstrSysName
                Exit Function
            End If
        ElseIf Val(NVL(rsTemp!��������)) = 1 Then
            If zlStr.IsHavePrivs(mstrPrivs, "������ý���") = False Then
                MsgBox "��û��������ý���Ȩ�ޣ����ܽ��н��ʲ�����", vbInformation, gstrSysName
                Exit Function
            End If
        Else
            If zlStr.IsHavePrivs(mstrPrivs, "סԺ���ý���") = False Then
                MsgBox "��û��סԺ���ý���Ȩ�ޣ����ܽ��н��ʲ�����", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        If mYBInFor.intInsure <> 0 Then
            If zlStr.IsHavePrivs(mstrPrivs, "���ս���") = False Then
                MsgBox "��û�б��ս���Ȩ�ޣ����ܽ��н��ʲ�����", vbInformation, gstrSysName
                Exit Function
            End If
        Else
            If zlStr.IsHavePrivs(mstrPrivs, "��ͨ���˽���") = False Then
                MsgBox "��û����ͨ���˽���Ȩ�ޣ����ܽ��н��ʲ�����", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    
    If mEditType = g_Ed_�������� Or mEditType = g_Ed_�������� Then
        If zlStr.IsHavePrivs(mstrPrivs, "��������") = False Then
            MsgBox "��û�н�������Ȩ�ޣ����ܽ��н������ϲ�����", vbInformation, gstrSysName
            Exit Function
        End If
        If mYBInFor.intInsure <> 0 Then
            If zlStr.IsHavePrivs(mstrPrivs, "���ս���") = False Then
                MsgBox "��û�б��ս���Ȩ�ޣ����ܽ��н������ϲ�����", vbInformation, gstrSysName
                Exit Function
            End If
        Else
            If zlStr.IsHavePrivs(mstrPrivs, "��ͨ���˽���") = False Then
                MsgBox "��û����ͨ���˽���Ȩ�ޣ����ܽ��н������ϲ�����", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
 
    '���������Ϣ
    Call Load�����Ϣ(lng����ID, Val(NVL(rsTemp!��������)))

    '����Ƿ��Լ��λ����:����:35090
    If Val(NVL(rsTemp!����ID)) = 0 Then
        If NVL(rsTemp!ԭ��) <> "" Then
            txtPatient.Text = NVL(rsTemp!ԭ��)
        Else
            strSQL = "" & _
            "   Select  D.���� " & _
            "   From " & IIf(mblnNOMoved, "H", "") & "������ü�¼ A, ������Ϣ C, ��Լ��λ D " & _
            "   Where A.����ID=[1]  And A.����ID=C.����ID And C.��ͬ��λid = D.ID(+) and Rownum=1 " & _
            "    Union ALL " & _
            "   Select  D.���� " & _
            "   From " & IIf(mblnNOMoved, "H", "") & "סԺ���ü�¼ A, ������Ϣ C, ��Լ��λ D " & _
            "   Where A.����ID=[1] And C.��ͬ��λid = D.ID(+) and Rownum=1 " & _
            "   "
            Set rsUnit = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(NVL(rsTemp!ID)))
            If Not rsUnit.EOF Then
                txtPatient.Text = NVL(rsUnit!����)
            Else
                txtPatient.Text = "δ�ҵ���Լ��λ"
            End If
        End If
        txtPatient.Tag = "��Լ��λ"
    Else
        txtPatient.Text = NVL(rsTemp!����)
        txtPatient.Tag = Val(NVL(rsTemp!����ID))
    End If
     
    txtSex.Text = NVL(rsTemp!�Ա�)
    txtOld.Text = NVL(rsTemp!����)
    txt�ѱ�.Text = NVL(rsTemp!�ѱ�)
    txtDate.Text = Format(rsTemp!�շ�ʱ��, "yyyy-MM-dd HH:mm:ss")
    txtInvoice.Text = NVL(rsTemp!Ʊ�ݺ�)
    '����65105,������:���˲�������������������ʾ
    mobjBalanceCon.blnCurBalanceOwnerFee = False
    lblBalanceType.Visible = False
    Select Case Val(NVL(rsTemp!��������))
        '10.29��ǰ�����ͣ���������
        Case 0
        Case 1
            txt��ʶ��.Text = NVL(rsTemp!�����)
            txt��ʶ��.Visible = True
            lbl��ʶ��.Visible = True
            lbl��ʶ��.Caption = "�����"
'            mobjBalanceCon.blnCurBalanceOwnerFee = True
            lblPatiTime.Visible = False
            txtPatiBegin.Visible = False
            lblPatiTimeRange.Visible = False
            txtPatiEnd.Visible = False
            txt����.Visible = False
            lblDayName.Visible = False
        Case 2
            txt��ʶ��.Text = NVL(rsTemp!סԺ��)
            txt��ʶ��.Visible = True
            lbl��ʶ��.Visible = True
            lbl��ʶ��.Caption = "סԺ��"

            If Not IsNull(rsTemp!��ǰ����) Then
                txtBed.Text = rsTemp!��ǰ����
                txtBed.Visible = True
                lblBed.Visible = True
            End If

            If Not IsNull(rsTemp!��ǰ����) Then
                txt����.Text = rsTemp!��ǰ����
                txt����.Visible = True
                lbl����.Visible = True
            End If
            opt��Ժ.Value = IIf(Val(NVL(rsTemp!��;����)) = 1, False, True)
            opt��;.Value = IIf(Val(NVL(rsTemp!��;����)) = 1, True, False)
            lblBalanceType.Visible = True
            lblBalanceType.Caption = IIf(Val(NVL(rsTemp!��;����)) = 1, "��;����", "��Ժ����")
    End Select

    txtBegin.Text = Format(dMin, txtBegin.Format)
    txtEnd.Text = Format(dMax, txtEnd.Format)
    txtBalance(Idx_����˵��).Text = NVL(rsTemp!��ע)

    If mobjBalanceCon.blnCurBalanceOwnerFee = False Then
        '���������ʱ
        If Not IsNull(rsTemp!��ʼ����) Then
            txtPatiBegin.Text = Format(rsTemp!��ʼ����, "yyyy-MM-dd")
        End If

        If Not IsNull(rsTemp!��������) Then
            txtPatiEnd.Text = Format(rsTemp!��������, "yyyy-MM-dd")
        End If
    End If

    lngID = rsTemp!ID
    
    
    str��ҳIds = IIf(mty_ModulePara.bln����ָ��Ԥ���� And mbln����תסԺ = False, _
    IIf(mobjBalanceCon.strTime = "", mobjBalanceAll.strAllTime, mobjBalanceCon.strTime), "")
    If Not LoadFeeListFromBalanceID(lngID) Then Exit Function    '���ط�����ϸ
    If Not LoadBalanceDepositList(lng����ID, lngID, mblnNOMoved, str��ҳIds) Then Exit Function   '���س�Ԥ����
    
    If Not LoadBalancePayData(lng����ID, lngID, mblnNOMoved) Then Exit Function  '�����Ѿ�֧������
    If mEditType = g_Ed_�������� Then
        Dim blnReadOldBalan As Boolean
        
        mrsBalance.Filter = 0
        blnReadOldBalan = mrsBalance.RecordCount = 0
        If mrsBalance.RecordCount = 1 Then
            blnReadOldBalan = NVL(mrsBalance!���㷽ʽ) = ""
        End If
        If blnReadOldBalan Then
            If Not LoadBalancePayData(lng����ID, mBalanceInfor.lng����ID, mblnNOMoved, True) Then Exit Function     '�����Ѿ�֧������
        End If
        If zlGetFromIDToBalanceData(mBalanceInfor.lng����ID, mblnNOMoved, mrsOldBalance) = False Then Exit Function
           
    End If
    If mEditType = g_Ed_���½��� Then
        '  dblMoney = RoundEX(mBalanceInfor.dbl��ǰ���� - mBalanceInfor.dblҽ��֧���ϼ�, 2)
        '��������(0-������г�Ԥ��;1-��ȱʡʹ��Ԥ����;2-��ָ���������Ԥ��(��ʱ���Ⱥ�����̯��;3-ȫ��
        strSQL = "Select 1 From �����˿���Ϣ Where ����ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
        If rsTmp.EOF Then
            Call RecalcDepositMoney(1)
            mblnNotChange = True
            txtBalance(Idx_��Ԥ��).Text = Format(mBalanceInfor.dbl��Ԥ���ϼ�, "0.00")
            mblnNotChange = False
        Else
            Call RecalcDepositMoney(3)
            mblnNotChange = True
            txtBalance(Idx_��Ԥ��).Text = Format(mBalanceInfor.dbl��Ԥ���ϼ�, "0.00")
            mblnNotChange = False
            txtBalance(Idx_��Ԥ��).Enabled = False
            chkDeposit.Enabled = False
        End If
    End If

    If mEditType <> g_Ed_���ݲ鿴 Then
        mblnNotChange = True
        Call LoadCurOwnerPayInfor
        '0-ҽ��Ԥ����Ϣ��ʾ;1-��ʾ������Ϣ
        Call ShowLedDisplayBank(1)
        Call SetOperationCtrl(2)     'bytFun-0-����ǰ;1-ҽ����������;2-�ѱ����˽��ʵ�;
        If mEditType = g_Ed_�������� Or mEditType = g_Ed_�������� Or chkCancel.Value = 1 Then
            ReInitPatiInvoice False
            InitRedInvoice True
        Else
            ReInitPatiInvoice True
        End If
        mblnNotChange = False
    End If
    Call ShowYBMoneyOrInvioceFormatInfor
    Call SetCurBalanceVisible
    If mEditType = g_Ed_���½��� Then
        Call txtBalance_Validate(Idx_��Ԥ��, False)
        SetDefaultPayType
        If txtBalance(Idx_�ɿ�).Enabled And txtBalance(Idx_�ɿ�).Visible Then txtBalance(Idx_�ɿ�).SetFocus
    End If
    ReadBalance = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function GetFromalanceIDToPatiNum(ByVal lng����ID As Long, Optional ByVal lngMax As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݽ���ID����ȡ���ν��ʵ�סԺ����
    '����:lngMax-����סԺ����
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-16 11:10:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strTime As String
    On Error GoTo errHandle
    
    strSQL = "" & _
    "   Select distinct ��ҳID " & _
    "   From סԺ���ü�¼  " & _
    "   Where ����ID= [1]  "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    lngMax = 0
    With rsTemp
        Do While Not .EOF
            If lngMax < Val(NVL(!��ҳID)) Then lngMax = Val(NVL(!��ҳID))
            strTime = strTime & "," & Val(NVL(!��ҳID))
            .MoveNext
        Loop
    End With
    If strTime <> "" Then strTime = Mid(strTime, 2)
    GetFromalanceIDToPatiNum = strTime
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function
Private Function ExecuteOldOneCardPayInterface(ByVal lng����ID As Long, ByVal lng����ID As Long, _
    ByVal objCard As Card, ByVal dblMoney As Double, tyBrushCardInfor As TY_BrushCard, _
    ByRef cllBillPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:һ��֧ͨ��(�ϰ汾)
    '���:lng�������-��������Ž��д���
    '     dblMoney-���ν�����
    '     TYBrushCardInfor-��ǰˢ����Ϣ
    '     cllBillPro-���ݹ���(ִ��������,�Ա�����´νӿ�ʱ�ظ�ִ��)
    '����:ִ�гɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-08 16:14:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl��� As Double, strҽԺ���� As String
    Dim i As Long, strSQL As String, str���㷽ʽ As String
    Dim cllPro As Collection, blnTrans As Boolean
    Dim intCardType As Integer, strSwapNO As String
    '��һ��֧ͨ��,ֱ�ӷ���
    If objCard.�������� <> 7 Then ExecuteOldOneCardPayInterface = True: Exit Function

    mOldOneCard.rsOneCard.Filter = "���㷽ʽ='" & objCard.���㷽ʽ & "'"
    If mOldOneCard.rsOneCard.EOF Then
        MsgBox objCard.���㷽ʽ & "δ����,���ڡ������������á�����������!", vbInformation, gstrSysName
        ExecuteOldOneCardPayInterface = False: Exit Function
    End If
    
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    '����֮ǰ,�ȴ�������
    'Zl_���˽��ʽ���_Modify
    strSQL = "Zl_���˽��ʽ���_Modify("
    '  ��������_In     Number,
    '  --��������_In:
    '  --   0-��ͨ�շѷ�ʽ:
    '  --   1.����������:
    '  --     �ٽ��㷽ʽ_IN:ֻ�ܴ���һ�����㷽ʽ,���������һЩ������Ϣ,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ"
    '  --     ����֧Ʊ��_In:������
    '  --     �ۿ����ID_IN,����_IN,������ˮ��_IN,����˵��_In:��Ҫ����
    strSQL = strSQL & "1,"
    '  ����id_In       ���˽��ʼ�¼.����id%Type,
    strSQL = strSQL & "" & lng����ID & ","
    '  ����id_In       ����Ԥ����¼.����id%Type,
    strSQL = strSQL & "" & lng����ID & ","
    '  ���㷽ʽ_In     Varchar2,
    str���㷽ʽ = objCard.���㷽ʽ
    str���㷽ʽ = str���㷽ʽ & "|" & dblMoney
    str���㷽ʽ = str���㷽ʽ & "|" & IIf(tyBrushCardInfor.str������� = "", " ", tyBrushCardInfor.str�������)
    str���㷽ʽ = str���㷽ʽ & "|" & IIf(tyBrushCardInfor.str����ժҪ = "", " ", tyBrushCardInfor.str����ժҪ)
    strSQL = strSQL & "'" & str���㷽ʽ & "',"
    '  ��Ԥ��_In       ����Ԥ����¼.��Ԥ��%Type := Null,
    strSQL = strSQL & "NULL,"
    '  ��֧Ʊ��_In     ����Ԥ����¼.��Ԥ��%Type := Null,
    strSQL = strSQL & "NULL,"
    '  �����id_In     ����Ԥ����¼.�����id%Type := Null,
    strSQL = strSQL & "NULL,"
    '  ����_In         ����Ԥ����¼.����%Type := Null,
    strSQL = strSQL & "NULL,"
    '  ������ˮ��_In   ����Ԥ����¼.������ˮ��%Type := Null,
    strSQL = strSQL & "NULL,"
    '  ����˵��_In     ����Ԥ����¼.����˵��%Type := Null,
    strSQL = strSQL & "NULL,"
    '  �ɿ�_In         ����Ԥ����¼.�ɿ�%Type := Null,
    strSQL = strSQL & "NULL,"
    '  �Ҳ�_In         ����Ԥ����¼.�Ҳ�%Type := Null,
    strSQL = strSQL & "NULL,"
    '  �����_In     ������ü�¼.ʵ�ս��%Type := Null,
    strSQL = strSQL & "NULL,"
    '  ��������_In     Number := 2,
    strSQL = strSQL & "" & IIf(mEditType = g_Ed_������� Or mblnCurMzBalanceNo, 1, 2) & ","
    '  ȱʡ���㷽ʽ_In ���㷽ʽ.����%Type := Null,
    strSQL = strSQL & "NULL,"
    '    ����Ա���_In    ����Ԥ����¼.����Ա���%Type := Null,
    strSQL = strSQL & "'" & UserInfo.��� & "',"
    '    ����Ա����_In    ����Ԥ����¼.����Ա����%Type := Null,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '    �տ�ʱ��_In      ����Ԥ����¼.����Ա����%Type := Null,
    strSQL = strSQL & "to_date('" & Format(mBalanceInfor.dtBalanceDate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
    '    ��Ԥ������ids_In Varchar2 := Null,
    strSQL = strSQL & "NULL,"
    '  ��ɽ���_In Number:=0
    strSQL = strSQL & "0)"
    zlAddArray cllPro, strSQL
    
    'һ��ͨ����
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    
    If Not mobjICCard.PaymentSwap(dblMoney, dbl���, intCardType, Val("" & mOldOneCard.rsOneCard!ҽԺ����), tyBrushCardInfor.str����, tyBrushCardInfor.str������ˮ��, lng����ID, lng����ID) Then
        gcnOracle.RollbackTrans
        MsgBox objCard.���㷽ʽ & "����ʧ��!", vbOKOnly, gstrSysName
        Exit Function
    End If
    gcnOracle.CommitTrans
    blnTrans = False
    strSQL = "Zl_һ��ͨ����_Update(" & 0 & ",'" & objCard.���㷽ʽ & "','" & tyBrushCardInfor.str���� & "','" & intCardType & "','" & strSwapNO & "'," & dbl��� & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Set cllBillPro = New Collection
    blnTrans = False
    ExecuteOldOneCardPayInterface = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
 End Function
 
Private Function CheckOldOneCardIsValied(ByVal objCard As Card, _
    ByVal dblMoney As Double, ByRef tyBrushCard As TY_BrushCard, _
    Optional bln�˿� As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���һ��ͨ�Ƿ���ȷ
    '���:objCard-��ǰ������
    '     bln�˿�-�Ƿ��˿�
    '����:tyBrushCard-����ˢ����Ϣ
    '����:һ��ͨ��֤��ȷ���һ��ͨ,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-08 17:19:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblδ����� As Double, strCardNo As String
    Dim dblTemp As Double, strXmlIn As String
    Dim intMousePointer As Integer
    intMousePointer = Screen.MousePointer
    
    On Error GoTo errHandle
    
    If objCard.�������� <> 7 Then CheckOldOneCardIsValied = True: Exit Function
    
    mOldOneCard.rsOneCard.Filter = "���㷽ʽ='" & objCard.���㷽ʽ & "'"
    If mOldOneCard.rsOneCard.EOF Then
        Screen.MousePointer = 0
        MsgBox objCard.���㷽ʽ & "δ����,���ڡ������������á�����������!", vbInformation, gstrSysName
        CheckOldOneCardIsValied = False: Exit Function
    End If

    If mobjICCard Is Nothing Then
        Set mobjICCard = New clsICCard
        Call mobjICCard.SetParent(Me.hWnd)
        Set mobjICCard.gcnOracle = gcnOracle
    End If
    If mobjICCard Is Nothing Then
        Screen.MousePointer = 0
        MsgBox "һ��ͨ�ӿڴ���ʧ��!", vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If dblMoney = 0 Then dblMoney = Val(txtBalance(Idx_�ɿ�).Text)
     
    If dblMoney = 0 Then
        Screen.MousePointer = 0
        MsgBox IIf(bln�˿�, "�˿�", "�տ�") & "���δ����,����!", vbInformation + vbOKOnly, gstrSysName
         Exit Function
    End If
    
    dblδ����� = RoundEx(mBalanceInfor.dblδ���ϼ� - mBalanceInfor.dbl��Ԥ���ϼ�, 6)
    If Abs(dblMoney) > Format(Abs(dblδ�����), "0.00") And dblMoney <> 0 Then
        Screen.MousePointer = 0
        MsgBox IIf(bln�˿�, "�˿�", "�տ�") & "���ܴ��ڱ���" & IIf(bln�˿�, "δ��", "δ��") & "���:" & Format(Abs(dblδ�����), "0.00") & " ��", vbInformation, gstrSysName
        Exit Function
    End If
        
    If Not bln�˿� Then
       
       '����ˢ������
       'zlBrushCard(frmMain As Object, _
       '    ByVal lngModule As Long, _
       '    ByVal rsClassMoney As ADODB.Recordset, _
       '    ByVal lngCardTypeID As Long, _
       '    ByVal bln���ѿ� As Boolean, _
       '    ByVal strPatiName As String, ByVal strSex As String, _
       '    ByVal strOld As String, ByVal dbl��� As Double, _
       '    Optional ByRef strCardNo As String, _
       '    Optional ByRef strPassWord As String, _
       '    Optional ByRef bln�˷� As Boolean = False, _
       '    Optional ByRef blnShowPatiInfor As Boolean = False, _
       '    Optional ByRef bln���� As Boolean = False, _
       '    Optional ByVal bln�����ֹ As Boolean = True) As Boolean
       '---------------------------------------------------------------------------------------------------------------------------------------------
       '����:����ָ��֧�����,����ˢ������
       '���:rsClassMoney:�շ����,���
       '        lngCardTypeID-Ϊ��ʱ,Ϊ��һ��ͨˢ��
       '       bln�����ֹ-Ŀǰֻ������ѿ�,��ʾ����ʱ,��ֹ��������,������������֧��
        
        strXmlIn = "<IN><CZLX>0</CZLX></IN>"
        If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, Nothing, 0, False, _
        mrsInfo!����, NVL(mrsInfo!�Ա�), NVL(mrsInfo!����), IIf(mPatiInfor.bln�˿��־, -1, 1) * dblMoney, tyBrushCard.str����, tyBrushCard.str����, _
        False, True, False, False, Nothing, False, False, strXmlIn) = False Then Exit Function
        
        tyBrushCard.dbl�ʻ���� = mobjICCard.GetSpare
        If tyBrushCard.dbl�ʻ���� < dblMoney Then
            Screen.MousePointer = 0
            MsgBox "������֧��,����!" & vbCrLf & vbCrLf & _
            "   �� ��  ��" & Format(tyBrushCard.dbl�ʻ����, "0.00") & vbCrLf & _
            "   ����֧��" & Format(dblMoney, "0.00"), vbInformation, gstrSysName
            Exit Function
        End If
        staThis.Panels(2).Text = Format(tyBrushCard.dbl�ʻ����, "0.00")
        staThis.Panels(2).ToolTipText = objCard.���㷽ʽ & "���ʻ����:" & Format(tyBrushCard.dbl�ʻ����, "0.00")
       
        CheckOldOneCardIsValied = True
        Exit Function
    End If
    '�˿���
    If mrsBalance Is Nothing Then
        Screen.MousePointer = 0
        MsgBox "ע��:" & vbCrLf & "  δ�ҵ�ԭʼ�Ľ����¼,����ʹ��" & objCard.���� & "�����˿�!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If mrsBalance.State <> 1 Then
        Screen.MousePointer = 0
        MsgBox "ע��:" & vbCrLf & "  δ�ҵ�ԭʼ�Ľ����¼,����ʹ��" & objCard.���� & "�����˿�!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
    mrsBalance.Filter = "����=4"
    If mrsBalance.EOF Then
        Screen.MousePointer = 0
        MsgBox "ע��:" & vbCrLf & "  δ�ҵ�ԭʼ�Ľ����¼,����ʹ��" & objCard.���� & "�����˿�!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    strCardNo = mobjICCard.Read_Card(Me)
    If strCardNo = "" Then
        Screen.MousePointer = 0
        MsgBox "һ��ͨ����ʧ��,�뽫IC�����ڶ�������", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    If strCardNo <> NVL(mrsBalance!����) Then
        Screen.MousePointer = 0
        MsgBox "��ǰ������ۿ�Ų�һ��,���ܽ����˷�.", vbInformation, gstrSysName
        Exit Function
    End If
    
    dblTemp = Format(Val(NVL(mrsBalance!��Ԥ��)), "0.00")
    If RoundEx(dblMoney, 6) <> Format(dblTemp, "0.00") Then
        Screen.MousePointer = 0
        MsgBox "һ��ͨ�������ȫ��,����!" & vbCrLf & vbCrLf & _
        "   ������" & Format(dblTemp, "0.00") & vbCrLf & _
        "   ����֧��" & Format(dblMoney, "0.00"), vbInformation, gstrSysName
        Exit Function
    End If
    CheckOldOneCardIsValied = True
    Exit Function
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = intMousePointer
        Resume
    End If
End Function

Private Function CheckThreeSwapValied(ByVal objCard As Card, _
    ByVal dblMoney As Double, ByRef tyBrushCard As TY_BrushCard, _
    Optional ByVal bln�˿� As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������֤
    '���:objCard-������
    '     dblMoney-ˢ�����,>=0��ʾ�տ�;С�����ʾ�˿�
    '     bln�˿�-true,��ʾ��ǰΪ�˿���;False��ʾ��ǰΪ�տ���
    '����:tyBrushCard-ˢ����Ϣ
    '����:���׺Ϸ�����true,���򷵻�False
    '����:���˺�
    '����:2014-07-08 18:00:34
    '˵��:ͬ����֤�˽ӿں�ˢ���ӿڵ�
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblTemp As Double, cllSquareBalance As Collection
    Dim strXMLExpend As String, bln���� As Boolean
    Dim dbl�ʻ���� As Double, dblδ����� As Double
    Dim strExpand As String, strXmlIn As String
    Dim strBalanceIDs As String
    Dim intMousePointer As Integer
    Dim blnCurInput As Boolean
    
    intMousePointer = Screen.MousePointer
    
    
    If objCard Is Nothing Then
        If InStr(";" & mstrPrivsCard & ";", ";�����ӿ�����;") = 0 Then
            MsgBox "��û�������ӿ�����Ȩ�ޣ��޷����ýӿڲ�����", vbInformation, gstrSysName
            Exit Function
        Else
            MsgBox "δ�ҵ��˿�ӿ�,����ӿڲ�����", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    If objCard.�ӿ���� <= 0 Or objCard.���ѿ� Then CheckThreeSwapValied = True: Exit Function
    
    
    On Error GoTo errHandle
    tyBrushCard.blnת�� = False
    If dblMoney = 0 Then dblMoney = Val(txtBalance(Idx_�ɿ�).Text): blnCurInput = True
    
    dblδ����� = RoundEx(mBalanceInfor.dblδ���ϼ� - mBalanceInfor.dbl��Ԥ���ϼ�, 6)
     
    If dblMoney = 0 Then
        If dblδ����� = 0 Then
            CheckThreeSwapValied = True: Exit Function
        Else
            Screen.MousePointer = 0
            MsgBox IIf(bln�˿�, "�˿�", "�տ�") & "���δ����,����!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
    
    
    If Abs(dblMoney) > Format(Abs(dblδ�����), "0.00") And dblMoney <> 0 And blnCurInput Then
        Screen.MousePointer = 0
        MsgBox IIf(bln�˿�, "�˿�", "ˢ��") & "���ܴ��ڱ���" & IIf(bln�˿�, "δ��", "δ��") & "���:" & Format(Abs(dblδ�����), "0.00") & " ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    If Abs(dblMoney) <> Format(Abs(dblδ�����), "0.00") And blnCurInput Then
        If mty_ModulePara.bytˢ��ȱʡ������ = 1 Then
            If MsgBox(IIf(bln�˿�, "�˿�", "ˢ��") & "���(" & Format(dblMoney, "0.00") & ")�뱾��" & IIf(bln�˿�, "δ��", "δ��") & "���(" & Format(Abs(dblδ�����), "0.00") & _
                ")��ͬ���Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        ElseIf mty_ModulePara.bytˢ��ȱʡ������ = 2 Then
            MsgBox IIf(bln�˿�, "�˿�", "ˢ��") & "���(" & Format(dblMoney, "0.00") & ")�뱾��" & IIf(bln�˿�, "δ��", "δ��") & "���(" & Format(Abs(dblδ�����), "0.00") & _
                ")��ͬ�����ܼ�����", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    If Not bln�˿� Then
        'zlBrushCard(frmMain As Object, _
           ByVal lngModule As Long, _
           ByVal rsClassMoney As ADODB.Recordset, _
           ByVal lngCardTypeID As Long, _
           ByVal bln���ѿ� As Boolean, _
           ByVal strPatiName As String, ByVal strSex As String, _
           ByVal strOld As String, ByRef dbl��� As Double, _
           Optional ByRef strCardNo As String, _
           Optional ByRef strPassWord As String, _
           Optional ByRef bln�˷� As Boolean = False, _
           Optional ByRef blnShowPatiInfor As Boolean = False, _
           Optional ByRef bln���� As Boolean = False, _
           Optional ByVal bln�����ֹ As Boolean = True, _
           Optional ByRef varSquareBalance As Variant) As Boolean
           '       varSquareBalance- Collection����,���ص�ǰˢ������(array(�����ID,���ѿ�ID,ˢ�����,����,����,�������,�Ƿ�����))
        strXmlIn = "<IN><CZLX>0</CZLX></IN>"
        If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, Nothing, _
            objCard.�ӿ����, objCard.���ѿ�, _
            mPatiInfor.str����, mPatiInfor.str�Ա�, mPatiInfor.str����, IIf(mPatiInfor.bln�˿��־, -1, 1) * dblMoney, _
            tyBrushCard.str����, tyBrushCard.str����, _
            False, True, False, False, cllSquareBalance, False, False, strXmlIn) = False Then Exit Function
            '����ǰ,һЩ���ݼ��
        'zlPaymentCheck(frmMain As Object, ByVal lngModule As Long, _
            ByVal strCardTypeID As Long, ByVal strCardNo As String, _
            ByVal dblMoney As Double, ByVal strNOs As String, _
            Optional ByVal strXMLExpend As String
        If gobjSquare.objSquareCard.zlPaymentCheck(Me, mlngModul, objCard.�ӿ����, _
            objCard.���ѿ�, tyBrushCard.str����, dblMoney, "", strXMLExpend) = False Then Exit Function
        '    zlGetAccountMoney(ByVal frmMain As Object, ByVal lngModule As Long, _
        '    ByVal strCardTypeID As Long, _
        '    ByVal strCardNo As String, strExpand As String, dblMoney As Double
          '���:frmMain-���õ�������
          '        lngModule-ģ���
          '        strCardNo-����
          '        strExpand-Ԥ����Ϊ��,�Ժ���չ
          '����:dblMoney-�����ʻ����
        If gobjSquare.objSquareCard.zlGetAccountMoney(Me, mlngModul, objCard.�ӿ����, _
              tyBrushCard.str����, strExpand, dbl�ʻ����, objCard.���ѿ�) = False Then Exit Function
        
        staThis.Panels(2).Text = Format(dbl�ʻ����, "0.00")
        staThis.Panels(2).ToolTipText = objCard.���㷽ʽ & "���ʻ����:" & Format(dbl�ʻ����, "0.00")
        tyBrushCard.dbl�ʻ���� = RoundEx(dbl�ʻ����, 2)
        If dbl�ʻ���� <> 0 And dbl�ʻ���� < dblMoney Then
            Screen.MousePointer = 0
            MsgBox objCard.���㷽ʽ & "���ʻ�����!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        CheckThreeSwapValied = True
        Exit Function
    End If
    
    '�˿���
    If mrsBalance Is Nothing Then
        If objCard.�Ƿ�ת�ʼ����� Then GoTo GoTransferAccount:
        Screen.MousePointer = 0
        MsgBox "ע��:" & vbCrLf & "  δ�ҵ�ԭʼ�Ľ����¼,����ʹ��" & objCard.���� & "�����˿�!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If mrsBalance.State <> 1 Then
        If objCard.�Ƿ�ת�ʼ����� Then GoTo GoTransferAccount:
        Screen.MousePointer = 0
        MsgBox "ע��:" & vbCrLf & "  δ�ҵ�ԭʼ�Ľ����¼,����ʹ��" & objCard.���� & "�����˿�!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    
    '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
    If mEditType = g_Ed_�������� Then
        mrsOldBalance.Filter = "����=3 And �����ID=" & objCard.�ӿ����
        If mrsOldBalance.EOF Then
            If objCard.�Ƿ�ת�ʼ����� Then GoTo GoTransferAccount:
            Screen.MousePointer = 0
            MsgBox "ע��:" & vbCrLf & "  δ�ҵ�ԭʼ�Ľ����¼,����ʹ��" & objCard.���㷽ʽ & "�����˿�!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    Else
        mrsBalance.Filter = "����=3 And �����ID=" & objCard.�ӿ����
        If mrsBalance.EOF Then
            If objCard.�Ƿ�ת�ʼ����� Then GoTo GoTransferAccount:
            Screen.MousePointer = 0
            MsgBox "ע��:" & vbCrLf & "  δ�ҵ�ԭʼ�Ľ����¼,����ʹ��" & objCard.���㷽ʽ & "�����˿�!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
   
    dblTemp = 0
    If mEditType = g_Ed_�������� Then
        With mrsOldBalance
            Do While Not .EOF
                dblTemp = dblTemp + Val(NVL(!��Ԥ��))
                .MoveNext
            Loop
            mrsOldBalance.MoveFirst
            dblTemp = RoundEx(dblTemp, 5)
        End With
    Else
        With mrsBalance
            Do While Not .EOF
                dblTemp = dblTemp + Val(NVL(!��Ԥ��))
                .MoveNext
            Loop
            mrsBalance.MoveFirst
            dblTemp = RoundEx(dblTemp, 5)
        End With
    End If
    
    If dblTemp = 0 Then
        If objCard.�Ƿ�ת�ʼ����� Then GoTo GoTransferAccount:
        Screen.MousePointer = 0
        MsgBox "ע��:" & vbCrLf & objCard.���㷽ʽ & "�Ѿ����꣬�������ˣ�", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If objCard.�Ƿ�ȫ�� Then
        If dblTemp <> dblMoney Then
            If objCard.�Ƿ�ת�ʼ����� Then GoTo GoTransferAccount:
            Screen.MousePointer = 0
            MsgBox "ע��:" & vbCrLf & objCard.���� & "�����˿�ʱ������ȫ�ˣ�" & vbCrLf & _
            "  ʣ��δ��:" & Format(Abs(dblTemp), "0.00") & vbCrLf & _
            "  ��ǰ���:" & Format(Abs(dblMoney), "0.00"), vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    Else
        If dblMoney > dblTemp Then
            If objCard.�Ƿ�ת�ʼ����� Then GoTo GoTransferAccount:
        End If
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
        
    strXMLExpend = ""
    If mEditType = g_Ed_�������� Then
        tyBrushCard.str���� = NVL(mrsOldBalance!����)
        tyBrushCard.str������ˮ�� = NVL(mrsOldBalance!������ˮ��)
        tyBrushCard.str����˵�� = NVL(mrsOldBalance!����˵��)
    Else
        tyBrushCard.str���� = NVL(mrsBalance!����)
        tyBrushCard.str������ˮ�� = NVL(mrsBalance!������ˮ��)
        tyBrushCard.str����˵�� = NVL(mrsBalance!����˵��)
    End If

    strBalanceIDs = "2|" & mBalanceInfor.lng����ID & IIf(mBalanceInfor.lng����ID = 0, "", "," & mBalanceInfor.lng����ID)
    If gobjSquare.objSquareCard.zlReturnCheck(Me, mlngModul, objCard.�ӿ����, objCard.���ѿ�, tyBrushCard.str����, _
        strBalanceIDs, dblMoney, tyBrushCard.str������ˮ��, tyBrushCard.str����˵��, strXMLExpend) = False Then Exit Function
    
    If objCard.�Ƿ��˿��鿨 Then
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
        strXmlIn = "<IN><CZLX>2</CZLX></IN>"
        If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, Nothing, objCard.�ӿ����, _
            objCard.���ѿ�, mPatiInfor.str����, mPatiInfor.str�Ա�, _
            mPatiInfor.str����, IIf(mPatiInfor.bln�˿��־, -1, 1) * dblMoney, tyBrushCard.str����, tyBrushCard.str����, _
            True, True, bln����, True, Nothing, False, False, strXmlIn) = False Then Exit Function
    End If
    CheckThreeSwapValied = True
    Exit Function
    
GoTransferAccount:
    strXmlIn = "<IN><CZLX>1</CZLX></IN>"
    If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, Nothing, objCard.�ӿ����, _
        objCard.���ѿ�, mPatiInfor.str����, mPatiInfor.str�Ա�, _
        mPatiInfor.str����, IIf(bln�˿�, -1, 1) * dblMoney, tyBrushCard.str����, tyBrushCard.str����, _
        True, True, bln����, True, Nothing, False, False, strXmlIn) = False Then Exit Function
    
    tyBrushCard.blnת�� = True
    '����ת�ʽӿ�
    '    7.1.    zltransferAccountsCheck(ת�ʼ��ӿ�)
    'zlTransferAccountsCheck ת�ʼ��ӿ�
    '������  ��������    ��/��   ��ע
    'frmMain Object  In  ���õ�������
    'lngModule   Long    In  HIS����ģ���
    'lngCardTypeID   Long    In  �����ID
    'strCardNo   String  In  ����
    'dblMoney    Double  In  ת�ʽ��(����ʱΪ����)
    'strBalanceIDs   String  In  ����IDs������ö��ŷ��룬��ʾ���ζ��Ĵ��շ���Ŀ��������ҽ��������
    'strXMLExpend String In   XML��:
    '                            <IN>
    '                                <CZLX >��������</CZLX> //0��NULL:������ҵ��;1-�������˷�ҵ��2-����ҵ��;3-�����˷�ҵ��
    '                            </IN>
    '                    Out  XML��:
    '                            <OUT>
    '                               <ERRMSG>������Ϣ</ERRMSG >
    '                            </OUT>
    '    Boolean ��������    �������ݺϷ�,����True:���򷵻�False
    '˵��:
    '��. ��ҽ���������ʱ���е�����ת��ʱ��һЩ�Ϸ��Լ�飬������ת��ʱ�����Ի���֮��ĵȴ������������������ķ�����
    '��. �����ڼ�����Ҫ����ΪTrue�����������ת�ʹ��ܵĵ��á�
    '����XML��
    strXMLExpend = "<IN><CZLX>3</CZLX></IN>"
    If gobjSquare.objSquareCard.zltransferAccountsCheck(Me, mlngModul, objCard.�ӿ����, _
        tyBrushCard.str����, dblMoney, mBalanceInfor.lng����ID, strXMLExpend) = False Then
        Screen.MousePointer = 0
        Call zlShowThreeSwapErrInfor(0, strXMLExpend)
        Exit Function
    End If
    
'    zlGetAccountMoney(ByVal frmMain As Object, ByVal lngModule As Long, _
'    ByVal strCardTypeID As Long, _
'    ByVal strCardNo As String, strExpand As String, dblMoney As Double
    '���:frmMain-���õ�������
    '        lngModule-ģ���
    '        strCardNo-����
    '        strExpand-Ԥ����Ϊ��,�Ժ���չ
    '����:dblMoney-�����ʻ����
    Call gobjSquare.objSquareCard.zlGetAccountMoney(Me, mlngModul, objCard.�ӿ����, _
          tyBrushCard.str����, strExpand, dbl�ʻ����, objCard.���ѿ�)
    If dbl�ʻ���� <> 0 Then
        staThis.Panels(2).Text = objCard.���㷽ʽ & "�ʻ����:" & Format(dbl�ʻ����, "0.00")
        staThis.Panels(2).ToolTipText = objCard.���㷽ʽ & "���ʻ����:" & Format(dbl�ʻ����, "0.00")
    End If
    tyBrushCard.dbl�ʻ���� = RoundEx(dbl�ʻ����, 2)
    CheckThreeSwapValied = True
    Exit Function
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = intMousePointer
        Resume
    End If
    Call SaveErrLog
End Function
 


Private Function ExecuteThreeSwapPayInterface(ByVal lng����ID As Long, ByVal lng����ID As Long, objCard As Card, ByVal dblMoney As Double, _
    ByRef cllBillPro As Collection, ByRef tyBrushCard As TY_BrushCard) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:һ��֧ͨ��(�����ӿ�)
    '���:lng�������-��������Ž��д���
    '     dblMoney-���ν�����
    '     cllBillPro-���ݹ���(ִ��������,�Ա�����´νӿ�ʱ�ظ�ִ��)
    '     tyBrushCard-��ǰˢ����Ϣ
    '����:ִ�гɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-07-09 18:14:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim cllPro As Collection, blnTrans As Boolean
    Dim str����IDs As String, i As Long, strSQL As String
    Dim cllUpdate As Collection, cllThreeSwap As Collection
    Dim str���㷽ʽ  As String
    
    Set cllUpdate = New Collection
    Set cllThreeSwap = New Collection
    
    '��һ��֧ͨ��,ֱ�ӷ���
    If objCard.�ӿ���� <= 0 Or objCard.���ѿ� Then ExecuteThreeSwapPayInterface = True: Exit Function
    
    If dblMoney = 0 Then ExecuteThreeSwapPayInterface = True: Exit Function
    
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    
   '����֮ǰ,�ȴ�������
    'Zl_���˽��ʽ���_Modify
    strSQL = "Zl_���˽��ʽ���_Modify("
    '  ��������_In     Number,
    '  --��������_In:
    '  --   0-��ͨ�շѷ�ʽ:
    '  --   1.����������:
    '  --     �ٽ��㷽ʽ_IN:ֻ�ܴ���һ�����㷽ʽ,���������һЩ������Ϣ,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ"
    '  --     ����֧Ʊ��_In:������
    '  --     �ۿ����ID_IN,����_IN,������ˮ��_IN,����˵��_In:��Ҫ����
    strSQL = strSQL & "1,"
    '  ����id_In       ���˽��ʼ�¼.����id%Type,
    strSQL = strSQL & "" & lng����ID & ","
    '  ����id_In       ����Ԥ����¼.����id%Type,
    strSQL = strSQL & "" & lng����ID & ","
    '  ���㷽ʽ_In     Varchar2,
    str���㷽ʽ = objCard.���㷽ʽ
    str���㷽ʽ = str���㷽ʽ & "|" & dblMoney
    str���㷽ʽ = str���㷽ʽ & "|" & IIf(tyBrushCard.str������� = "", " ", tyBrushCard.str�������)
    str���㷽ʽ = str���㷽ʽ & "|" & IIf(tyBrushCard.str����ժҪ = "", " ", tyBrushCard.str����ժҪ)
    strSQL = strSQL & "'" & str���㷽ʽ & "',"
    '  ��Ԥ��_In       ����Ԥ����¼.��Ԥ��%Type := Null,
    strSQL = strSQL & "NULL,"
    '  ��֧Ʊ��_In     ����Ԥ����¼.��Ԥ��%Type := Null,
    strSQL = strSQL & "NULL,"
    '  �����id_In     ����Ԥ����¼.�����id%Type := Null,
    strSQL = strSQL & "" & objCard.�ӿ���� & ","
    '  ����_In         ����Ԥ����¼.����%Type := Null,
    strSQL = strSQL & "'" & tyBrushCard.str���� & "',"
    '  ������ˮ��_In   ����Ԥ����¼.������ˮ��%Type := Null,
    strSQL = strSQL & "'" & tyBrushCard.str������ˮ�� & "',"
    '  ����˵��_In     ����Ԥ����¼.����˵��%Type := Null,
    strSQL = strSQL & "'" & tyBrushCard.str����˵�� & "',"
    '  �ɿ�_In         ����Ԥ����¼.�ɿ�%Type := Null,
    strSQL = strSQL & "" & mBalanceInfor.dbl�ɿ� & ","
    '  �Ҳ�_In         ����Ԥ����¼.�Ҳ�%Type := Null,
    strSQL = strSQL & "" & mBalanceInfor.dbl�Ҳ� & ","
    '  �����_In     ������ü�¼.ʵ�ս��%Type := Null,
    strSQL = strSQL & "NULL,"
    '  ��������_In     Number := 2,
    strSQL = strSQL & "" & IIf(mEditType = g_Ed_������� Or mblnCurMzBalanceNo, 1, 2) & ","
    '  ȱʡ���㷽ʽ_In ���㷽ʽ.����%Type := Null,
    strSQL = strSQL & "NULL,"
    '    ����Ա���_In    ����Ԥ����¼.����Ա���%Type := Null,
    strSQL = strSQL & "'" & UserInfo.��� & "',"
    '    ����Ա����_In    ����Ԥ����¼.����Ա����%Type := Null,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '    �տ�ʱ��_In      ����Ԥ����¼.����Ա����%Type := Null,
    strSQL = strSQL & "to_date('" & Format(mBalanceInfor.dtBalanceDate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
    '    ��Ԥ������ids_In Varchar2 := Null,
    strSQL = strSQL & "NULL,"
    '  ��ɽ���_In Number:=0
    strSQL = strSQL & "0)"
    zlAddArray cllPro, strSQL
    
    'zlPaymentMoney(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
    ByVal bln���ѿ� As Boolean, _
    ByVal strCardNo As String, ByVal strBalanceIDs As String, _
    byval  strPrepayNos as string , _
    ByVal dblMoney As Double, _
    ByRef strSwapGlideNO As String, _
    ByRef strSwapMemo As String, _
    Optional ByRef strSwapExtendInfor As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ʻ��ۿ��
    '���:frmMain-���õ�������
    '        lngModule-����ģ���
    '        strBalanceIDs-����ID,����ö��ŷ���
    '        strPrepayNos-��Ԥ��ʱ��Ч. Ԥ�����ݺ�,����ö��ŷ���
    '       strCardNo-����
    '       dblMoney-������
    '����:strSwapGlideNO-������ˮ��
    '       strSwapMemo-����˵��
    '       strSwapExtendInfor-������չ��Ϣ: ��ʽΪ:��Ŀ����1|��Ŀ����2||��||��Ŀ����n|��Ŀ����n
    '����:�ۿ�ɹ�,����true,���򷵻�Flase
    '˵��:
    '   ��������Ҫ�ۿ�ĵط����øýӿ�,Ŀǰ�滮��:�շ��ң��Һ���;������ѯ��;ҽ������վ��ҩ���ȡ�
    '   һ����˵���ɹ��ۿ�󣬶�Ӧ�ô�ӡ��صĽ���Ʊ�ݣ����Է��ڴ˽ӿڽ��д���.
    '   �ڿۿ�ɹ��󣬷��ؽ�����ˮ�ź���ر�ע˵���������������������Ϣ�����Է��ڽ���˵�����Ա��˷�.
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    str����IDs = lng����ID
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    
    If gobjSquare.objSquareCard.zlPaymentMoney(Me, mlngModul, objCard.�ӿ����, objCard.���ѿ�, tyBrushCard.str����, _
         str����IDs, "", dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor) = False Then
        gcnOracle.RollbackTrans
        Exit Function
    End If
    
    tyBrushCard.str������ˮ�� = strSwapGlideNO
    tyBrushCard.str����˵�� = strSwapMemo
    
    If objCard.���ѿ� = False Then
        Call zlAddUpdateSwapSQL(False, str����IDs, objCard.�ӿ����, objCard.���ѿ�, tyBrushCard.str����, strSwapGlideNO, strSwapMemo, cllUpdate, 2)
    End If
    Call zlAddThreeSwapSQLToCollection(False, str����IDs, objCard.�ӿ����, objCard.���ѿ�, tyBrushCard.str����, strSwapExtendInfor, cllThreeSwap)
    zlExecuteProcedureArrAy cllUpdate, Me.Caption, True, True
    gcnOracle.CommitTrans
    Set cllBillPro = New Collection
    '��������������Ϣ
    zlExecuteProcedureArrAy cllThreeSwap, Me.Caption
    blnTrans = False
    mBalanceInfor.blnSaveBill = True
    ExecuteThreeSwapPayInterface = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub initCardSquareData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���㿨����������Ϣ
    '���:blnClosed:�رն���
    '����:���˺�
    '����:2010-01-05 14:51:23
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card, strKinds As String
    Dim intIdkind As Integer
    Dim strIdkind As String
    If mEditType = g_Ed_���ݲ鿴 Then Exit Sub
        
    On Error GoTo errHandle
    
    If gobjSquare.objSquareCard Is Nothing Then Exit Sub
    'strKinds = "��|����|0|0|0|0|0|;ҽ|ҽ����|0|0|0|0|0|;��|���֤��|0|0|0|0|0|;IC|IC����|1|0|0|0|0|;��|�����|0|0|0|0|0|;ס|סԺ��|0|0|0|0|0|;��|���￨|0|0|0|0|0|"
    
    Call IDKIND.zlInit(Me, glngSys, mlngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, strKinds, txtPatient)
    Call GetRegInFor(g˽��ģ��, Me.Name, "IDKIND", strIdkind)
    If Val(strIdkind) > 0 And Val(strIdkind) <= IDKIND.ListCount Then IDKIND.IDKIND = Val(strIdkind)
    Set objCard = IDKIND.GetfaultCard
    Set gobjSquare.objDefaultCard = objCard
    If IDKIND.Cards.��ȱʡ������ And Not objCard Is Nothing Then
        gobjSquare.blnȱʡ�������� = objCard.�������Ĺ��� <> ""
        gobjSquare.intȱʡ���ų��� = objCard.���ų���
    Else
        gobjSquare.blnȱʡ�������� = IDKIND.Cards.������ʾ
        gobjSquare.intȱʡ���ų��� = 100
    End If
    gobjSquare.bln��ȱʡ������ = IDKIND.Cards.��ȱʡ������
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub InitPancel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����ʼ����
    '����:���˺�
    '����:2014-05-26 10:30:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim sngWidth As Single, lngHeight As Long
    Dim strReg As String
    Dim panThis As Pane, panThis1 As Pane
    lngHeight = picPati.Height \ Screen.TwipsPerPixelY
    Set panThis = dkpMain.CreatePane(mConPans.Pan_PatiCon, 200, lngHeight, DockLeftOf, Nothing)
    panThis.Title = "��������"
    panThis.Handle = picPati.hWnd
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panThis.Tag = mConPans.Pan_PatiCon
    panThis.MaxTrackSize.Height = lngHeight
    panThis.MinTrackSize.Height = lngHeight
    
    Set panThis1 = dkpMain.CreatePane(mConPans.Pan_FeeList, 250, 580, DockBottomOf, panThis)
    panThis1.Title = "��Ŀ��"
    panThis1.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panThis1.Handle = picFeeList.hWnd
    panThis1.Tag = mConPans.Pan_FeeList
    
    If mEditType = g_Ed_���ݲ鿴 Then
'        Set panThis = dkpMain.CreatePane(mConPans.Pan_Deposit, 250, 580, DockRightOf, panThis1)
'        panThis.Title = "Ԥ�����"
'        panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
'        panThis.Handle = picDeposit.hWnd
'        panThis.Tag = mConPans.Pan_Deposit
        Set panThis = dkpMain.CreatePane(mConPans.Pan_Balance, 250, 580, DockRightOf, panThis1)
        panThis.Title = "�����б�"
        panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
        panThis.Handle = picBalanceBack.hWnd
        panThis.Tag = mConPans.Pan_Balance
        panThis.MaxTrackSize.Width = 8000 \ Screen.TwipsPerPixelY
        panThis.MinTrackSize.Width = panThis.MaxTrackSize.Width
    Else
        Set panThis = dkpMain.CreatePane(mConPans.Pan_Balance, 250, 580, DockRightOf, panThis1)
        panThis.Title = "�����б�"
        panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
        panThis.Handle = picBalanceBack.hWnd
        panThis.Tag = mConPans.Pan_Balance
        panThis.MaxTrackSize.Width = 7000 \ Screen.TwipsPerPixelY
        panThis.MinTrackSize.Width = panThis.MaxTrackSize.Width
    
'        Set panThis = dkpMain.CreatePane(mConPans.Pan_Deposit, 250, 580, DockBottomOf, panThis1)
'        panThis.Title = "Ԥ�����"
'        panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
'        panThis.Handle = picDeposit.hWnd
'        panThis.Tag = mConPans.Pan_Deposit
    End If
    
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.HideClient = True
    dkpMain.Options.LockSplitters = True
    dkpMain.VisualTheme = ThemeDefault
    dkpMain.RecalcLayout
    'zlRestoreDockPanceToReg  Me, dkpMain, "����"
End Sub

Private Sub txtPatiBegin_Change()
    If IsDate(txtPatiBegin.Text) And IsDate(txtPatiEnd.Text) Then
        txt����.Text = CDate(txtPatiEnd.Text) - CDate(txtPatiBegin.Text) + IIf(opt��;.Value = True, 1, 0)
        If Val(txt����.Text) = 0 Then txt����.Text = 1
    Else
        txt����.Text = ""
    End If
End Sub

Private Sub txtPatiBegin_GotFocus()
    zlControl.TxtSelAll txtPatiBegin
    mstrPatiBegin = txtPatiBegin.Text
End Sub

Private Sub txtPatiBegin_Validate(Cancel As Boolean)
    If txtPatiBegin.Text <> "____-__-__" And Not IsDate(txtPatiBegin.Text) Then
        MsgBox "��������ȷ��סԺ��ʼ����!", vbInformation, gstrSysName
        Cancel = True
    End If
End Sub

Private Sub txtPatiEnd_Change()
    If IsDate(txtPatiBegin.Text) And IsDate(txtPatiEnd.Text) Then
        txt����.Text = CDate(txtPatiEnd.Text) - CDate(txtPatiBegin.Text) + IIf(opt��;.Value = True, 1, 0)
        If Val(txt����.Text) = 0 Then txt����.Text = 1
    Else
        txt����.Text = ""
    End If
End Sub

Private Sub txtPatiEnd_GotFocus()
    zlControl.TxtSelAll txtPatiEnd
    mstrPatiEnd = txtPatiEnd.Text
End Sub

Private Sub txtPatiEnd_Validate(Cancel As Boolean)
    If txtPatiEnd.Text <> "____-__-__" And Not IsDate(txtPatiEnd.Text) Then
        MsgBox "��������ȷ��סԺ��������!", vbInformation, gstrSysName
        Cancel = True
   End If
End Sub
Private Function YBIdentifyCancel() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ȡ��ҽ�����������֤
    '����:���ؼ�ʱ���˳�������������
    '����:���˺�
    '����:2015-01-12 16:08:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long, varData As Variant
    On Error GoTo errHandle
        
    YBIdentifyCancel = True
    If mYBInFor.strYBPati <> "" Then
        varData = Split(mYBInFor.strYBPati, ";")
        If UBound(varData) >= 8 Then lng����ID = Val(varData(8))
        If lng����ID <> 0 Then YBIdentifyCancel = gclsInsure.IdentifyCancel(0, lng����ID, mYBInFor.intInsure)
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function RecalcFeeTotalDate() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������÷��õ�ͳ��ʱ��
    '����:����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-12 16:20:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str��ҳIds As String, strStartDate As String, strEndDate As String
    Dim i As Long, lngMax As Long, lngMin As Long, dtDate As Date
    Dim varData As Variant, lng����ID As Long
    
    If mEditType = g_Ed_������� Then RecalcFeeTotalDate = True: Exit Function
    
    If mrsInfo Is Nothing Then RecalcFeeTotalDate = True: Exit Function
    If mrsInfo.State = 0 Then RecalcFeeTotalDate = True: Exit Function
    
    varData = Split(cboPatiNums.GetNodesCheckedDatas(False), ",")
    For i = 0 To UBound(varData)
        If lngMax = 0 Then lngMax = Val(varData(i))
        If lngMin = 0 Then lngMin = Val(varData(i))
        If lngMax < Val(varData(i)) Then
            lngMax = Val(varData(i))
        End If
        If lngMin > Val(varData(i)) Then
            lngMin = Val(varData(i))
        End If
    Next
    
    If lngMin = 0 And lngMax = 0 Then
        MsgBox "����ѡ��סԺ����!", vbInformation, Me.Caption
        Exit Function
    End If
        
    lng����ID = Val(NVL(mrsInfo!����ID)): str��ҳIds = IIf(lngMin = lngMax, lngMax, lngMin & "," & lngMax)
    If mobjBalanceAll.GetPatiFeeDateRang(lng����ID, str��ҳIds, strStartDate, strEndDate, gint����ʱ�� = 0) = False Then
        strEndDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
        strStartDate = Format(CDate(strEndDate), "yyyy-mm-dd") & " 00:00:00"
    End If
    txtBegin.Text = Format(strStartDate, "yyyy-mm-dd")
    txtEnd.Text = Format(strEndDate, "yyyy-mm-dd")
    
    RecalcFeeTotalDate = True
End Function
 

Private Function CheckFactIsValied(ByRef objSetFocus As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鷢Ʊ�Ƿ�Ϸ�
    '����:objSetFocus -����ʱ,��궨λ���ĸ�����
    '����:���˺�
    '����:2015-01-13 10:21:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    '����ӡ,ֱ�ӷ���true,�����Ʊ�ݺ�
    If mobjFactProperty.��ӡ��ʽ = 0 Then CheckFactIsValied = True: Exit Function

    '�Ƚ��Էѷ���ʱ����ӡ��ƱƱ��
    If mty_ModulePara.blnNotPrintInvioce And mobjBalanceCon.blnCurBalanceOwnerFee Then CheckFactIsValied = True:  Exit Function

    If Not mobjFactProperty.�ϸ���� Then      '���ϸ����
        If Len(txtInvoice.Text) <> mobjFactProperty.Ʊ�ų��� And txtInvoice.Text <> "" Then
            MsgBox "Ʊ�ݺ��볤��Ӧ��Ϊ " & mobjFactProperty.Ʊ�ų��� & " λ��", vbInformation, gstrSysName
            Set objSetFocus = txtInvoice
            Exit Function
        End If
        CheckFactIsValied = True
        Exit Function
    End If

    If Trim(txtInvoice.Text) = "" Then
        MsgBox "��������һ����Ч��Ʊ�ݺ��룡", vbInformation, gstrSysName
        Set objSetFocus = txtInvoice
        Exit Function
    End If
    If zlGetInvoiceGroupUseID(mlng����ID, 1, txtInvoice.Text) = False Then Exit Function
    CheckFactIsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function PrintBill(ByVal lng����ID As Long, ByVal strNO As String, ByVal lng����ID As Long, _
    ByVal dtBalanceDate As Date, ByVal dbl�ɿ� As Double, ByVal dbl�Ҳ� As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Ƿ��ӡƱ��
    '���:strNO-���ʵ���
    '     lng����ID-����ID
    '     dtBalanceDate-��������
    '����:��ӡƱ��,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-13 10:08:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln��ӡ�˿��վ� As Boolean, bytKind As Byte
    Dim bln��ӡ������ϸ As Boolean, bln�Է��嵥 As Boolean, blnPrintBillEmpty As Boolean
        
    On Error GoTo errHandle
    
    bln��ӡ�˿��վ� = False
    If mty_ModulePara.int�˿�Ʊ�� <> 0 And InStr(1, mstrPrivs, ";���˽����˿��վ�;") > 0 Then
        '0-����ӡ,1-��ʾ��ӡ,2-����ʾ��ӡ;'���˺� ����:27776 ����:2010-02-04 16:49:03
        If mty_ModulePara.int�˿�Ʊ�� = 1 Then
           If MsgBox("���Ƿ�Ҫ��ӡ�����˽����˿��վݡ���" & vbCrLf & _
                   "   ���ǡ�����ӡ���˽����˿��վ�" & vbCrLf & _
                   "   ���񡻣�����ӡ���˽����˿��վ�", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                bln��ӡ�˿��վ� = True
            End If
        Else
            bln��ӡ�˿��վ� = True
        End If
    End If
  
    bln��ӡ������ϸ = False
     Select Case mty_ModulePara.bytFeePrintSet
     Case 1  '��ӡ.
         If MsgBox("���Ƿ�Ҫ��ӡ�����˽��ʷ�����ϸ����" & vbCrLf & _
                 "   ���ǡ�����ӡ���˽��ʷ�����ϸ" & vbCrLf & _
                 "   ���񡻣�����ӡ���˽��ʷ�����ϸ", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                 bln��ӡ������ϸ = True
         End If
     Case 0  '����ӡ
     Case 2  '��ӡ.������ʾ
         bln��ӡ������ϸ = True
     End Select
     
    If mobjBalanceCon.blnCurBalanceOwnerFee Then   '�Է��嵥��ӡ����
       bln�Է��嵥 = False
       Select Case Val(zlDatabase.GetPara("�Էѷ��ô�ӡ��ʽ", glngSys, mlngModul, "0"))
           Case 2  '��ӡ.
               If MsgBox("���Ƿ�Ҫ��ӡ�������Էѷ����嵥����" & vbCrLf & _
                       "   ���ǡ�����ӡ�����Էѷ����嵥" & vbCrLf & _
                       "   ���񡻣�����ӡ�����Էѷ����嵥", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                       bln�Է��嵥 = True
               End If
           Case 0  '����ӡ
           Case 1  '��ӡ.������ʾ
               bln�Է��嵥 = True
       End Select
    End If
        
    If bln��ӡ�˿��վ� Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1137_3", Me, "����ID=" & lng����ID, 2)
    End If
    
    'Ʊ�ݴ�ӡ
    If mblnPrintInvoice Or (mYBInFor.intInsure <> 0 And MCPAR.ҽ���ӿڴ�ӡƱ��) Then
        '����:44332
RePrint:
        Dim strNotValiedNos As String
        mobjFactProperty.LastUseID = mlng����ID
        Call UpateStartInvoice(mBalanceInfor.strNO, txtInvoice.Text)
        Call frmPrint.ReportPrint(1, strNO, lng����ID, mobjFactProperty, txtInvoice.Text, _
             dtBalanceDate, CCur(dbl�ɿ�), CCur(dbl�Ҳ�), , mobjFactProperty.��ӡ��ʽ, blnPrintBillEmpty, mYBInFor.intInsure <> 0 And MCPAR.ҽ���ӿڴ�ӡƱ��)
        If mEditType = g_Ed_������� Then
            bytKind = mty_ModulePara.bytInvoiceKindMZ
        Else
            bytKind = mty_ModulePara.bytInvoiceKindZY
        End If
        If mobjFactProperty.�ϸ���� And blnPrintBillEmpty = False And _
            ((bytKind = 0 And InStr(1, mstrPrivs, ";�վݴ�ӡ;") > 0) _
               Or (bytKind <> 0 And InStr(1, mstrPrivs, ";��ӡ�����շ�Ʊ��;") > 0)) Then    'blnPrintBillEmpty:55052
            '60155
             If zlIsNotSucceedPrintBill(3, strNO, strNotValiedNos) = True Then
                     If MsgBox("���ʵ���Ϊ[" & strNotValiedNos & "]�Ľ���Ʊ�ݴ�ӡδ�ɹ�,�Ƿ����´�ӡ����Ʊ��?", vbYesNo + vbDefaultButton1 + vbQuestion, gstrSysName) = vbYes Then GoTo RePrint:
             End If
        End If
    End If
    

    If bln��ӡ������ϸ Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1137_3", Me, "����ID=" & lng����ID, "����ID=" & lng����ID, 2)
    End If
    
    If bln�Է��嵥 Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1137_4", Me, "����ID=" & lng����ID, "����ID=" & lng����ID, 2)
    End If
    
    If mblnDepositBillPrint Then
        '��ӡԤ��Ʊ��
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103", Me, "NO=" & mBalanceInfor.strԤ��No, "����ID=" & mPatiInfor.lng����ID, "�տ�ʱ��=" & Format(mBalanceInfor.dtBalanceDate, "yyyy-mm-dd HH:MM:SS"), 2)
    End If
    
    PrintBill = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function UpateStartInvoice(ByVal strNO As String, ByVal strInvoice As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�޸Ŀ�ʼ��Ʊ��
    '���:strNO-���ʵ���
    '����:���˺�
    '����:2015-01-14 10:21:52
    '˵��:���������������,���Բ���ʹ�ô����������(�ɸ����ڲ���)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    strSQL = "Zl_Ʊ����ʼ��_Update('" & strNO & "','" & Trim(strInvoice) & "',3)"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
End Function
 
 
Private Function CheckInputConsValied(ByRef objSetFocus As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������������Ч�Լ��
    '����:objSetFocus-����ƶ���ָ���Ŀؼ�
    '����:����������Ч������True,���򷵻�False
    '����:���˺�
    '����:2015-01-12 18:03:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnNotFondPati As Boolean
    On Error GoTo errHandle
    If mrsInfo Is Nothing Then blnNotFondPati = True
    If Not blnNotFondPati Then blnNotFondPati = mrsInfo.State = 0
    
    If blnNotFondPati Then
        MsgBox "û��ȷ�����ʲ���,���ܽ��н��ʲ�����", vbExclamation, gstrSysName
        Set objSetFocus = txtPatient
        Exit Function
    End If

    If txtPatiBegin.Text <> "____-__-__" And Not IsDate(txtPatiBegin.Text) Then
        MsgBox "������һ����Ч�Ŀ�ʼʱ�䣡", vbInformation, gstrSysName
        Set objSetFocus = txtPatiBegin
        Exit Function
    End If
    If txtPatiEnd.Text <> "____-__-__" And Not IsDate(txtPatiEnd.Text) Then
        MsgBox "������һ����Ч�Ľ���ʱ�䣡", vbInformation, gstrSysName
        Set objSetFocus = txtPatiEnd
        Exit Function
    End If
    
    If IsDate(txtPatiBegin.Text) And IsDate(txtPatiEnd.Text) Then
        If txtPatiEnd < txtPatiBegin.Text Then
            MsgBox "����ʱ�䲻��С�ڿ�ʼʱ�䣡", vbInformation, gstrSysName
            Set objSetFocus = txtPatiBegin
            Exit Function
        End If
    End If
    If IsDate(txtPatiBegin.Text) And Not IsDate(txtPatiEnd.Text) Then
        MsgBox "��һ��������Ч�Ľ���ʱ�䣡", vbInformation, gstrSysName
        Set objSetFocus = txtPatiBegin
        Exit Function
    End If
    If Not IsDate(txtPatiBegin.Text) And IsDate(txtPatiEnd.Text) Then
        MsgBox "��һ��������Ч�Ŀ�ʼʱ�䣡", vbInformation, gstrSysName
        Set objSetFocus = txtPatiBegin
        Exit Function
    End If
    If mrsFeeList Is Nothing Then
        MsgBox "�����ò���û����Ҫ���ʵķ���������", vbInformation, gstrSysName
        Set objSetFocus = cmdRefresh
        Exit Function
    End If
    If mrsFeeList.State <> 1 Then
        MsgBox "�������²���û����Ҫ���ʵķ��ã�", vbInformation, gstrSysName
        Set objSetFocus = cmdRefresh
        Exit Function
    End If
    If mrsFeeList.RecordCount = 0 Then
        MsgBox "�������²���û����Ҫ���ʵķ��ã�", vbInformation, gstrSysName
        Set objSetFocus = cmdRefresh
        Exit Function
    End If
        
    If zlCommFun.StrIsValid(txtBalance(Idx_����˵��).Text, 50, txtBalance(Idx_����˵��).hWnd, "����˵��") = False Then
        Set objSetFocus = txtBalance(Idx_����˵��)
        Exit Function
    End If
    
    If Val(txtBalance(Idx_����δ��).Text) < Val(txtBalance(Idx_���ν���).Text) Then
        Call MsgBox("��ǰ���ʽ�������δ������ܽ��н��ʲ�����", vbInformation, gstrSysName)
        Set objSetFocus = txtBalance(Idx_���ν���)
        Exit Function
    End If

    If Val(txtBalance(Idx_����δ��).Text) <> 0 And Val(txtBalance(Idx_���ν���).Text) = 0 Then
        Call MsgBox("δ���뱾��Ҫ���ʵĽ����ܽ��н��ʲ�����", vbInformation, gstrSysName)
        Set objSetFocus = txtBalance(Idx_���ν���)
        Exit Function
    End If
    
    If Val(txtBalance(Idx_����δ��).Text) <= 0 Then
        If MsgBox("����ʵ��û�пɽ����,Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Set objSetFocus = txtPatient
            Exit Function
        End If
    End If
    
    '��鷢Ʊ�Ƿ���Ч
    If CheckFactIsValied(objSetFocus) = False Then Exit Function
    If CheckBusinessRuleIsValied(objSetFocus) = False Then Exit Function     'ҵ�������
    
    CheckInputConsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetSaveStrickDepositSQL(ByRef cllDeposit As Collection, ByRef objSetFocus As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����Ԥ���������
    '����:cllDeposit-��ص����ݼ�
    '     objSetFocus-��ȡʧ��ʱ,ȱʡ��궨λ��ָ���Ŀؼ���
    '����:��ȡ�ɹ�������true,���򷵻�False
    '����:���˺�
    '����:2015-01-19 15:12:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, intԤ����� As Integer, lng����ID As Long
    Dim dbl��Ԥ���ϼ� As Double, dblMoney As Double, dblԤ����� As Double
    Dim dbl��Ԥ�� As Double, dblԤ�����ϼ� As Double
    Dim strTime As String
    Dim rsDeposit As ADODB.Recordset, i As Long
    
    
    On Error GoTo errHandle
    lng����ID = mPatiInfor.lng����ID
    strTime = ""
    If mty_ModulePara.bln����ָ��Ԥ���� Then
        strTime = IIf(mobjBalanceCon.strTime = "", mobjBalanceAll.strAllTime, mobjBalanceCon.strTime)
    End If
    Set objSetFocus = txtBalance(Idx_��Ԥ��)
    
    intԤ����� = 2
    If mEditType = g_Ed_������� Or mblnCurMzBalanceNo Then intԤ����� = 1
    If cllDeposit Is Nothing Then Set cllDeposit = New Collection
    dblMoney = RoundEx(Val(txtBalance(Idx_��Ԥ��).Text), 2)
    With vsDeposit
        If .TextMatrix(1, .ColIndex("ID")) <> "" Then
            '�ض�����Ԥ��,���������ж�
            Set rsDeposit = GetDeposit(lng����ID, mblnDateMoved, strTime, , , intԤ�����, mrs���㷽ʽ)
            For i = 1 To .Rows - 1
                dblԤ����� = Val(.TextMatrix(i, .ColIndex("���")))
                dbl��Ԥ�� = Val(.TextMatrix(i, .ColIndex("��Ԥ��")))
                If dbl��Ԥ�� <> 0 Then
                    rsDeposit.Filter = "ID=" & CLng(.TextMatrix(i, .ColIndex("ID"))) & _
                        " And NO='" & .TextMatrix(i, .ColIndex("���ݺ�")) & "' And ��¼״̬=" & .RowData(i) & " And ���=" & dblԤ�����
                    If rsDeposit.RecordCount = 0 Then
                        If MsgBox("���ڲ�������,����Ԥ�����ѷ����仯,��������ȡ���˽���,�Ƿ�������ȡԤ������?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                             Call LoadDepositList(lng����ID, strTime)
                        End If
                        Screen.MousePointer = 0
                        Exit Function
                    End If
                    
                    strSQL = "zl_����Ԥ����¼_Insert(" & CLng(.TextMatrix(i, .ColIndex("ID"))) & "," & _
                        "'" & .TextMatrix(i, .ColIndex("���ݺ�")) & "'," & .RowData(i) & "," & _
                        dbl��Ԥ�� & "," & mBalanceInfor.lng����ID & "," & lng����ID & ")"
                    zlAddArray cllDeposit, strSQL
                   dbl��Ԥ���ϼ� = RoundEx(dbl��Ԥ���ϼ� + dbl��Ԥ��, 6)
                End If
                dblԤ�����ϼ� = RoundEx(dblԤ�����ϼ� + dblԤ�����, 6)
            Next
            '���ʳ����Ԥ��������Ԥ��������б����Ϻ�,����ָ���Ԥ������
            If Val(dbl��Ԥ���ϼ�) > Val(dblԤ�����ϼ�) And dbl��Ԥ���ϼ� <> 0 Then
                Call MsgBox("����Ԥ����������!", vbInformation, gstrSysName)
                Screen.MousePointer = 0
                Exit Function
            End If
        End If
    End With
    
    dbl��Ԥ���ϼ� = RoundEx(dbl��Ԥ���ϼ�, 6)
    If Val(dbl��Ԥ���ϼ�) = Val(dblMoney) Then
        GetSaveStrickDepositSQL = True: Exit Function
    End If
    
    If MsgBox("��ǰ��Ԥ��������Ԥ����ϸ��һ��,�Ƿ����°���ǰδ������Ԥ��?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
        '��������(0-������г�Ԥ��;1-��ȱʡʹ��Ԥ����;2-�����ʽ������Ԥ��(��ʱ���Ⱥ�����̯��;3-ȫ��
        dblMoney = RoundEx(mBalanceInfor.dbl��ǰ���� - mBalanceInfor.dbl�Ѹ��ϼ�, 2)
        If dblMoney < 0 Then
            dblMoney = 0
            Call RecalcDepositMoney(0)
        Else
            Call RecalcDepositMoney(2, dblMoney)
        End If
        mblnNotChange = True
        txtBalance(Idx_��Ԥ��).Text = Format(mBalanceInfor.dbl��Ԥ���ϼ�, "0.00")
        mblnNotChange = False
    End If
    Screen.MousePointer = 0
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function CheckDepositValied(Optional blnCurBrushDeposit As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鵱ǰ��Ԥ�����Ƿ�Ϸ�
    '����:blnCurBrushDeposit-��ǰ��ˢ��Ԥ����
    '����:�Ϸ�����true,���򷵻�False
    '����:���˺�
    '����:2015-01-14 15:15:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney  As Double, rsDeposit As ADODB.Recordset, i As Long, strSQL As String
    Dim lng����ID As Long, strTime As String, intԤ����� As Integer
    Dim dblԤ����� As Double, dbl��Ԥ�� As Double, dblԤ�����ϼ� As Double, dbl��Ԥ���ϼ� As Double
    Dim dblTemp As Double
    
    On Error GoTo errHandle
    
    dblMoney = Val(txtBalance(Idx_��Ԥ��).Text)
    '��Ԥ�������ڱ���ʱ���
    If dblMoney = 0 Then CheckDepositValied = True: Exit Function
    
    'ˢ��Ԥ����ģ���Ҫ����ˢ��
    If mBalanceInfor.blnԤ����֤ Then CheckDepositValied = True: Exit Function
    
    blnCurBrushDeposit = True
    
    If Not IsNumeric(txtBalance(Idx_��Ԥ��).Text) And txtBalance(Idx_��Ԥ��).Text <> "" Then
        Screen.MousePointer = 0:
        MsgBox "��Ч��ֵ��", vbInformation, gstrSysName
        Exit Function
    ElseIf Val(txtBalance(Idx_��Ԥ��).Text) < 0 Then
        dblTemp = 0
        For i = 1 To vsDeposit.Rows - 1
            dblTemp = dblTemp + vsDeposit.TextMatrix(i, vsDeposit.ColIndex("���"))
        Next i
        If dblTemp >= 0 Then
            mblnNotChange = True
            MsgBox "Ԥ��������Ϊ����", vbInformation, gstrSysName
            mblnNotChange = False
            Screen.MousePointer = 0: Exit Function
        End If
    Else

    End If
    
    If Val(dblMoney) > Val(mPatiInfor.dblʵ�����) Then
        Screen.MousePointer = 0
        mblnNotChange = True
        MsgBox "��Ԥ���������˲��˵�Ԥ�����,����!" & vbCrLf & _
               "��ǰ��Ԥ:" & Format(dblMoney, "0.00") & vbCrLf & _
               "��ǰ���:" & Format(mPatiInfor.dblԤ�����, "0.00"), vbInformation + vbOKOnly, gstrSysName
        mblnNotChange = False
        Exit Function
    End If
    
    lng����ID = mPatiInfor.lng����ID
    With vsDeposit
        If .TextMatrix(1, .ColIndex("ID")) <> "" Then
            For i = 1 To .Rows - 1
                dblԤ����� = Val(.TextMatrix(i, .ColIndex("���")))
                dbl��Ԥ�� = Val(.TextMatrix(i, .ColIndex("��Ԥ��")))
                dbl��Ԥ���ϼ� = RoundEx(dbl��Ԥ���ϼ� + dbl��Ԥ��, 5)
                dblԤ�����ϼ� = RoundEx(dblԤ�����ϼ� + dblԤ�����, 5)
            Next
            '���ʳ����Ԥ��������Ԥ��������б����Ϻ�,����ָ���Ԥ������
            If Val(dbl��Ԥ���ϼ�) > Val(dblԤ�����ϼ�) And dbl��Ԥ���ϼ� <> 0 Then
                Screen.MousePointer = 0
                Call MsgBox("����Ԥ����������!", vbInformation, gstrSysName)
                Exit Function
            End If
        End If
    End With
    
    dbl��Ԥ���ϼ� = RoundEx(dbl��Ԥ���ϼ�, 6)
    If Val(dbl��Ԥ���ϼ�) <> Val(dblMoney) Then
        Screen.MousePointer = 0
        If MsgBox("��ǰ��Ԥ��������Ԥ����ϸ��һ��,�Ƿ����°���ǰδ������Ԥ��?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            '��������(0-������г�Ԥ��;1-��ȱʡʹ��Ԥ����;2-�����ʽ������Ԥ��(��ʱ���Ⱥ�����̯��;3-ȫ��
            dblMoney = RoundEx(mBalanceInfor.dbl��ǰ���� - mBalanceInfor.dbl�Ѹ��ϼ�, 2)
            If dblMoney < 0 Then
                dblMoney = 0
                Call RecalcDepositMoney(0)
            Else
                Call RecalcDepositMoney(2, dblMoney)
            End If
            mblnNotChange = True
            txtBalance(Idx_��Ԥ��).Text = Format(mBalanceInfor.dbl��Ԥ���ϼ�, "0.00")
            mblnNotChange = False
        End If
        Exit Function
    End If
    
    '����ˢ����֤
    If gdblԤ��������鿨 = 0 Then
        txtBalance(Idx_��Ԥ��).BackColor = &HE0E0E0
        mBalanceInfor.blnԤ����֤ = True
        CheckDepositValied = True: Exit Function
    End If
    
    'סԺ�Ĳ���ˢ����֤
    If Not (mEditType = g_Ed_������� Or mblnCurMzBalanceNo) Then
        mBalanceInfor.blnԤ����֤ = True
        txtBalance(Idx_��Ԥ��).BackColor = &HE0E0E0
        CheckDepositValied = True: Exit Function
    End If
    
    txtBalance(Idx_��Ԥ��).BackColor = &H8000000F
    mBalanceInfor.blnԤ����֤ = True
    CheckDepositValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckCurBalanceIsValied(ByRef tyBrushCard As TY_BrushCard, _
    Optional ByVal blnԤ�� As Boolean = False, _
    Optional ByRef objSetFocus As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鵱ǰ�����Ƿ���Ч
    '����:tyBrushCard��ǰˢ����Ϣ
    '     objSetFocus-����ƶ�����
    '����:��Ч����true,���򷵻�False
    '����:���˺�
    '����:2015-01-14 14:57:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card, lng����ID As Long, varData As Variant
    Dim dblMoney As Double, i As Long, blnFind As Boolean
    Dim cllDeposit As Collection, int���� As Integer
    Dim dblCheck As Double
    
    Dim intCount As Integer '���ֽ��㷽ʽ(�ſ�ҽ��)
    On Error GoTo errHandle
    
    '������������Ч�Լ��
    If Not mBalanceInfor.blnSaveBill And mblnNotify = False Then
        If CheckInputConsValied(objSetFocus) = False Then Exit Function
    End If
    
    If InStr(txtBalance(Idx_�������).Text, "'") > 0 Then
        MsgBox "������뺬�зǷ��ַ�������,���������", vbInformation + vbOKOnly, gstrSysName
         Set objSetFocus = txtBalance(Idx_�������)
         Exit Function
    End If
    
    If zlCommFun.ActualLen(txtBalance(Idx_�������).Text) > 30 Then
        MsgBox "����������ֻ������30���ַ���15������,���������", vbInformation + vbOKOnly, gstrSysName
         Set objSetFocus = txtBalance(Idx_�������)
         Exit Function
    End If
    
    If InStr(txtBalance(Idx_ժҪ).Text, "'") > 0 Then
        MsgBox "ժҪ���зǷ��ַ�������,���������", vbInformation + vbOKOnly, gstrSysName
         Set objSetFocus = txtBalance(Idx_ժҪ)
         Exit Function
    End If
 
    If zlCommFun.ActualLen(txtBalance(Idx_ժҪ).Text) > 30 Then
        MsgBox "����������ֻ������50���ַ���25������,���������", vbInformation + vbOKOnly, gstrSysName
         Set objSetFocus = txtBalance(Idx_ժҪ)
         Exit Function
    End If
        
        
    Set objCard = IDKindPaymentsType.GetCurCard
    If objCard Is Nothing Then
        Screen.MousePointer = 0
        MsgBox "��ǰ��������Ч�Ľ��㷽ʽ����ѡ����Ч�Ľ��㷽ʽ!", vbInformation + vbOKOnly, gstrSysName
        Set objSetFocus = IDKindPaymentsType
        Exit Function
    End If


    With vsBlance
        blnFind = False
        For i = 1 To .Rows - 1
            '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
            int���� = Val(.TextMatrix(i, .ColIndex("����")))
            If blnԤ�� Then
                If int���� = 1 Then blnFind = True: Exit For
            End If
            If Not (objCard.���ѿ� And objCard.���ƿ�) Then '���ѿ�,�Ѿ����,�����ٴ���
                If .TextMatrix(i, .ColIndex("���㷽ʽ")) = objCard.���㷽ʽ Then blnFind = True
            End If
            If InStr("34", int����) > 0 And mbln�������� Then
                MsgBox "��������ģʽ��,������ʹ��:" & .TextMatrix(i, .ColIndex("���㷽ʽ")) & "���н���!", vbInformation + vbOKOnly, gstrSysName
                Exit Function
            End If
            int���� = Val(.TextMatrix(i, .ColIndex("��������")))
            dblCheck = Val(.TextMatrix(i, .ColIndex("������")))
            If InStr(",1,2,", "," & int���� & ",") > 0 And dblCheck <> 0 Then intCount = intCount + 1
        Next
        
        If blnFind Then
            Screen.MousePointer = 0
            If blnԤ�� Then
                MsgBox "�Ѿ���Ԥ���֧��,ֻ��ɾ��Ԥ�������֧��!", vbOKOnly, gstrSysName
            Else
                MsgBox objCard.���㷽ʽ & " �Ѿ�֧����,��������" & objCard.���㷽ʽ & "����֧��", vbOKOnly + vbDefaultButton1, gstrSysName
            End If
            Exit Function
        End If
        If InStr("34", objCard.��������) > 0 And mbln�������� Then
            MsgBox "��������ģʽ��,������ʹ��:" & objCard.���㷽ʽ & "���н���!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
            
        If InStr(",1,2,", "," & objCard.�������� & ",") > 0 Then
            intCount = intCount + 1
        End If
        If intCount > 1 And mbln�������� Then
            MsgBox "��������ģʽ��,������ʹ�ö��ֽ��㷽ʽ���н���!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End With

    '���ݼ��ӿ���(Ŀǰֻͬʱ֧�����ֽӿ�(��ҽ����һ�ֽӿ�)
    If dblMoney <> 0 Then If zlCheckMulitInterfaceNumValied = False Then Exit Function
        
    '1.���ѿ����
    If CheckSquareBalanceValied(objCard, tyBrushCard) = False Then
        Set objSetFocus = txtBalance(Idx_�ɿ�)
        Exit Function
    End If
     
    '2.�����ʻ����
    If CheckThreeSwapValied(objCard, dblMoney, tyBrushCard, mPatiInfor.bln�˿��־) = False Then
        Set objSetFocus = txtBalance(Idx_�ɿ�)
        Exit Function
    End If
    
    '3.һ��ͨ(�ϰ�)���
    If CheckOldOneCardIsValied(objCard, dblMoney, tyBrushCard) = False Then
        Set objSetFocus = txtBalance(Idx_�ɿ�)
        Exit Function
    End If
    
    '4.����ֽ���㷽ʽ
    If CheckCashValied(objCard) = False Then
        Set objSetFocus = txtBalance(Idx_�ɿ�)
        Exit Function
    End If
    
    
    '5.���֧Ʊ���㷽ʽ�Ƿ�Ϸ�
    If CheckChequeValied(objCard) = False Then
        Set objSetFocus = txtBalance(Idx_�ɿ�)
        Exit Function
    End If
    
    '6.����������㷽ʽ
    If CheckOtherValied(objCard) = False Then
        Set objSetFocus = txtBalance(Idx_�ɿ�)
        Exit Function
    End If
    CheckCurBalanceIsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckChequeValied(ByVal objCard As Card) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���֧Ʊ���㷽ʽ��һЩ�Ϸ�����
    '����:���ݺϷ�����true,���򷵻�False
    '����:���˺�
    '����:2014-07-08 18:21:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double
    Dim strTittle As String, dbl��ǰδ�� As Double
    Dim intMousePointer As Integer
    Dim objTempCard As Card
    Dim blnCheck As Boolean
    intMousePointer = Screen.MousePointer
    
    On Error GoTo errHandle
    
    If objCard.�������� <> 2 Or Not objCard.���㷽ʽ Like "*֧Ʊ*" Then CheckChequeValied = True: Exit Function
    
    
    dbl��ǰδ�� = RoundEx(mBalanceInfor.dblδ���ϼ� - mBalanceInfor.dbl��Ԥ���ϼ�, 5)
    
    strTittle = IIf(dbl��ǰδ�� < 0, "�˿�", "�տ�")
    dblMoney = Format(Val(txtBalance(Idx_�ɿ�).Text), "0.00")
     
    If strTittle = "�տ�" Then
    
        If RoundEx(dblMoney, 6) = 0 And Not mbln�������� Then
            Screen.MousePointer = 0
            MsgBox "δ�����տ��", vbInformation, gstrSysName
            Exit Function
        End If
        
        If dblMoney > RoundEx(dbl��ǰδ��, 2) Then
            Set objTempCard = IDKind�Ҳ�.GetCurCard
            blnCheck = False
            If objTempCard Is Nothing Then
                blnCheck = True
            Else
                If objTempCard.�ӿ���� = 1 Then blnCheck = True
            End If
            
            
            If mstr��֧Ʊ = "" And blnCheck Then
                Screen.MousePointer = 0
                MsgBox "�ڽ��㷽ʽ��û������Ӧ����Ľ��㷽ʽ,���ܽ�����֧Ʊ����", vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
        End If
        CheckChequeValied = True
        Exit Function
    End If
    
    '�˿�
    If RoundEx(dblMoney, 6) = 0 Then
        Screen.MousePointer = 0
        MsgBox "δ�����˿��", vbInformation, gstrSysName
        Exit Function
    End If
    CheckChequeValied = True
    Exit Function
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = intMousePointer
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckOtherValied(ByVal objCard As Card) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������㷽ʽ(֧Ʊ��)��һЩ�Ϸ�����
    '����:���ݺϷ�����true,���򷵻�False
    '����:���˺�
    '����:2015-01-14 15:44:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double
    Dim strTittle As String, dbl��ǰδ�� As Double
    Dim intMousePointer As Integer
    intMousePointer = Screen.MousePointer
    
    On Error GoTo errHandle
    
    If objCard.�ӿ���� > 0 Or objCard.���㷽ʽ Like "*֧Ʊ*" Or objCard.�������� = 1 Then CheckOtherValied = True: Exit Function
    
    dbl��ǰδ�� = RoundEx(mBalanceInfor.dblδ���ϼ� - mBalanceInfor.dbl��Ԥ���ϼ�, 5)
    strTittle = IIf(dbl��ǰδ�� < 0, "�˿�", "�տ�")
    dblMoney = Format(Val(txtBalance(Idx_�ɿ�).Text), "0.00")
  
    If strTittle = "�տ�" Then
        If RoundEx(dblMoney, 6) = 0 And Not mbln�������� And dbl��ǰδ�� <> 0 Then
            Screen.MousePointer = 0
            MsgBox "δ����" & strTittle & "��", vbInformation, gstrSysName
            Exit Function
        End If
        If dblMoney > RoundEx(dbl��ǰδ��, 2) Then
            Screen.MousePointer = 0
            MsgBox "ע��:" & vbCrLf & "    �����" & strTittle & "��������δ֧���Ľ��,���ܼ���!", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
        CheckOtherValied = True
        Exit Function
    End If
    
    '�˿�
    If RoundEx(dblMoney, 6) = 0 Then
        Screen.MousePointer = 0
        MsgBox "δ����" & strTittle & "��", vbInformation, gstrSysName
        Exit Function
    End If
    
    If dblMoney > RoundEx(Abs(dbl��ǰδ��), 2) Then
        Screen.MousePointer = 0
        MsgBox "ע��:" & vbCrLf & "    ������˿��������δ�˽��,���ܼ���!", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    Screen.MousePointer = 0
'    If ErrCenter() = 1 Then
'        Screen.MousePointer = intMousePointer
'        Resume
'    End If
    CheckOtherValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function CheckCashValied(ByVal objCard As Card, Optional ByVal bln�˿� As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ֽ���㷽ʽ��һЩ�Ϸ�����
    '���:objCard����ǰ֧����
    '     bln�˿�
    '����:���ݺϷ�����true,���򷵻�False
    '����:���˺�
    '����:2014-07-08 18:21:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, strTittle As String
    Dim intMousePointer As Integer
    intMousePointer = Screen.MousePointer

    
    On Error GoTo errHandle
    If objCard.�������� <> 1 Then CheckCashValied = True: Exit Function
    
    dblMoney = Format(Val(txtBalance(Idx_�ɿ�).Text), "0.00")
    If Not bln�˿� Then
        If RoundEx(dblMoney, 6) <> 0 Then
            If Val(dblMoney) < Val(lblʣ���Ը�.Caption) Then
                If mPatiInfor.bln�˿��־ Then
                    Screen.MousePointer = 0
                    MsgBox "�˿����,�벹��Ӧ�˽�" & vbCrLf & "����Ӧ��:" & lblʣ���Ը�.Caption & vbCrLf & "��ǰ�˿�" & Format(dblMoney, "0.00"), vbInformation, gstrSysName
                    Exit Function
                Else
                    Screen.MousePointer = 0
                    MsgBox "�տ����,�벹��Ӧ�ս�" & vbCrLf & "����Ӧ��:" & lblʣ���Ը�.Caption & vbCrLf & "��ǰ�տ�" & Format(dblMoney, "0.00"), vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
        '43153
        '�ɿ����:0-�����п���;1-������ȡ�ֽ�ʱ,��������ɿ�.
        If mty_ModulePara.byt�ɿ�������� = 0 Then CheckCashValied = True: Exit Function
        If mbln�������� Then CheckCashValied = True: Exit Function
        
        If txtBalance(Idx_�ɿ�).Text = "" Then
            Screen.MousePointer = 0
            MsgBox "�㻹δ����ɿ���,���ܼ���", vbExclamation, gstrSysName
            Exit Function
        End If

        CheckCashValied = True
        Exit Function
    End If
    
    '�˿��
    If dblMoney < Abs(Val(lblʣ���Ը�.Caption)) And RoundEx(dblMoney, 6) <> 0 Then
        Screen.MousePointer = 0
        MsgBox "������˿���㣡", vbInformation, gstrSysName
        Exit Function
    End If
    CheckCashValied = True
    Exit Function
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = intMousePointer
        Resume
    End If
    
    Call SaveErrLog
End Function
Private Function ExcutePatiOutHosptial() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ�в��˳�Ժ����
    '����:ִ�гɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-13 10:46:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnOut As Boolean, rsTmp As ADODB.Recordset
    Dim bln�������۲��� As Boolean
    Dim lng��ǰ����id As Long

    On Error GoTo errHandle
    If mrsInfo.State = 0 Then Exit Function
    bln�������۲��� = Val(NVL(mrsInfo!��������)) = 1
    If Not mty_ModulePara.blnAutoOut Then ExcutePatiOutHosptial = True: Exit Function
    If mEditType = g_Ed_������� And Not bln�������۲��� Or mblnCurMzBalanceNo Or mobjBalanceCon.blnCurBalanceOwnerFee Then ExcutePatiOutHosptial = True: Exit Function
    If mYBInFor.bytMCMode = 1 Then ExcutePatiOutHosptial = True: Exit Function
    
    '��Ժ�����ҳ�Ժ���ʻ���Ժ����������;���ʵ�,ֱ�ӷ���
    If bln�������۲��� Then
        If Val(NVL(mrsInfo!��Ժ)) <> 1 Then ExcutePatiOutHosptial = True: Exit Function
        lng��ǰ����id = Val(NVL(mrsInfo!��Ժ����ID))
    Else
        If Not (Not IsNull(mrsInfo!��ǰ����id) And opt��Ժ.Value) Then ExcutePatiOutHosptial = True: Exit Function
        lng��ǰ����id = Val("" & mrsInfo!��ǰ����id)
    End If
    '�Զ���Ժ(��Ժ����)
    blnOut = True
    If mYBInFor.intInsure <> 0 And Not MCPAR.δ�����Ժ Then
        Set rsTmp = GetMoneyInfo(mrsInfo!����ID, , , IIf(bln�������۲���, 1, 2))
        If Not rsTmp Is Nothing Then
            If NVL(rsTmp!�������, 0) <> 0 Then blnOut = False
        End If
    End If

    If gTy_System_Para.TY_Balance.blnҽ��������ܳ�Ժ And blnOut Then
        If Not checkҽ���´��Ժҽ��(mrsInfo!����ID, mrsInfo!��ҳID) Then blnOut = False
    End If
    If Not blnOut Then ExcutePatiOutHosptial = True: Exit Function  '�������Ժ��ֱ�ӷ���true
    
    frmAutoOut.mlng����ID = mrsInfo!����ID
    frmAutoOut.mlng��ҳID = mrsInfo!��ҳID
    frmAutoOut.mlngDepID = lng��ǰ����id
    frmAutoOut.mint���� = mYBInFor.intInsure
    frmAutoOut.mstr�Ա� = NVL(mrsInfo!�Ա�)
    frmAutoOut.Show 1, Me
    ExcutePatiOutHosptial = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function CheckNotExcuteItemValied(ByRef objSetFocus As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���δִ����Ŀ�Ƿ�Ϸ�
    '����:objSetFocus-���Ϸ�ʱ,���ع��ȱʡ��λ�ؼ�
    '����:�Ϸ����ط���true,���򷵻�False
    '����:���˺�
    '����:2015-01-12 18:33:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInfo As String
    On Error GoTo errHandle
    If gTy_System_Para.TY_Balance.byt���δִ�� = 0 Then CheckNotExcuteItemValied = True: Exit Function
    
    strInfo = ExistWaitExe(mrsInfo!����ID, NVL(mrsInfo!��ҳID, 0))
    If strInfo = "" Then CheckNotExcuteItemValied = True: Exit Function
        
    If gTy_System_Para.TY_Balance.byt���δִ�� = 1 Then
        If MsgBox("���ֲ���" & mrsInfo!���� & "������δִ����ɵ����ݣ�" & _
            vbCrLf & vbCrLf & strInfo & vbCrLf & vbCrLf & "Ҫ����������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Set objSetFocus = txtPatient
            Exit Function
        End If
    Else
        MsgBox "���ֲ���" & mrsInfo!���� & "������δִ����ɵ����ݣ�" & vbCrLf & vbCrLf & strInfo & vbCrLf & vbCrLf & "�������Ժ����.", vbInformation, gstrSysName
        Set objSetFocus = txtPatient
        Exit Function
    End If
    CheckNotExcuteItemValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckNotSendDrug() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���δ��ҩƷ�Ƿ�Ϸ�
    '����:�Ϸ�����true,���򷵻�False
    '����:���˺�
    '����:2015-01-12 18:33:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInfor As String, strInfo As String
    
    On Error GoTo errHandle
    If gTy_System_Para.TY_Balance.byt���δ��ҩ = 0 Then CheckNotSendDrug = True: Exit Function
    strInfo = ExistWaitDrug(mrsInfo!����ID, NVL(mrsInfo!��ҳID, 0), 1)
    If strInfo = "" Then CheckNotSendDrug = True: Exit Function
    If gTy_System_Para.TY_Balance.byt���δ��ҩ = 1 Then
        If MsgBox("���ֲ���" & mrsInfo!���� & strInfo & vbCrLf & vbCrLf & "Ҫ����������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
            Exit Function
        End If
    Else
        MsgBox "���ֲ���" & mrsInfo!���� & strInfo & vbCrLf & vbCrLf & "�������Ժ���ʡ�", vbInformation, gstrSysName
        If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
        Exit Function
    End If
    CheckNotSendDrug = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckMZNotExcuteItemValied(ByRef objSetFocus As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���δִ����Ŀ�Ƿ�Ϸ�
    '����:objSetFocus-���Ϸ�ʱ,���ع��ȱʡ��λ�ؼ�
    '����:�Ϸ����ط���true,���򷵻�False
    '����:���˺�
    '����:2015-01-12 18:33:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInfo As String
    On Error GoTo errHandle
    If gTy_System_Para.TY_Balance.byt������δִ�� = 0 Then CheckMZNotExcuteItemValied = True: Exit Function
    
    strInfo = ExistWaitExe(mrsInfo!����ID, NVL(mrsInfo!��ҳID, 0), 1)
    If strInfo = "" Then CheckMZNotExcuteItemValied = True: Exit Function
        
    If gTy_System_Para.TY_Balance.byt������δִ�� = 1 Then
        If MsgBox("���ֲ���" & mrsInfo!���� & "������δִ����ɵ����ݣ�" & _
            vbCrLf & vbCrLf & strInfo & vbCrLf & vbCrLf & "Ҫ����������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Set objSetFocus = txtPatient
            Exit Function
        End If
    Else
        MsgBox "���ֲ���" & mrsInfo!���� & "������δִ����ɵ����ݣ�" & vbCrLf & vbCrLf & strInfo & vbCrLf & vbCrLf & "�������������.", vbInformation, gstrSysName
        Set objSetFocus = txtPatient
        Exit Function
    End If
    CheckMZNotExcuteItemValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckMZNotSendDrug() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���δ��ҩƷ�Ƿ�Ϸ�
    '����:�Ϸ�����true,���򷵻�False
    '����:���˺�
    '����:2015-01-12 18:33:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInfor As String, strInfo As String
    
    On Error GoTo errHandle
    If gTy_System_Para.TY_Balance.byt������δ��ҩ = 0 Then CheckMZNotSendDrug = True: Exit Function
    strInfo = ExistWaitDrug(mrsInfo!����ID, NVL(mrsInfo!��ҳID, 0), 1, 1)
    If strInfo = "" Then CheckMZNotSendDrug = True: Exit Function
    If gTy_System_Para.TY_Balance.byt������δ��ҩ = 1 Then
        If MsgBox("���ֲ���" & mrsInfo!���� & strInfo & vbCrLf & vbCrLf & "Ҫ����������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
            Exit Function
        End If
    Else
        MsgBox "���ֲ���" & mrsInfo!���� & strInfo & vbCrLf & vbCrLf & "������������ʡ�", vbInformation, gstrSysName
        If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
        Exit Function
    End If
    CheckMZNotSendDrug = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckDelAppleyFeeValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�˷�������
    '����:�˷�������Ϸ�����true,���򷵻�False
    '����:���˺�
    '����:2015-01-12 18:43:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If Not mty_ModulePara.blnAutoOut Then CheckDelAppleyFeeValied = True: Exit Function
    If IsNull(mrsInfo!��ǰ����id) Or opt��Ժ.Value = False Or mYBInFor.bytMCMode = 1 Then CheckDelAppleyFeeValied = True: Exit Function
    
    If GetUnAuditReFee(mrsInfo!����ID, NVL(mrsInfo!��ҳID, 0)) Then
        If MsgBox("����" & txtPatient.Text & "�����������˷ѵ�δ��˵ļ�¼,ȷ��Ҫ���г�Ժ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Function
        End If
    End If
    CheckDelAppleyFeeValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

 
Private Function �������() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������Ч�Լ��
    '����:�Ϸ�����true,���򷵻�False
    '����:���˺�
    '����:2015-01-12 19:03:35
    '˵��:30036(bug)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str����ԭ�� As String

    On Error GoTo errHandle
    mBalanceInfor.str����ԭ�� = ""

    If Not mty_ModulePara.bln���ʼ�鲡������ Or opt��Ժ.Value = False Then ������� = True: Exit Function
    
    
    If IsCheck�����ѽ���(Val(NVL(mrsInfo!����ID)), Val(NVL(mrsInfo!��ҳID))) Then ������� = True: Exit Function
    
    If MsgBox("���ֲ���" & mrsInfo!���� & "û�н��в������," & _
        vbCrLf & "Ҫ����������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus: Exit Function
    End If
    
    str����ԭ�� = ""
    If frmInputBox.InputBox(Me, "����δ��ԭ��", "�����벡��δ��ԭ����Ϣ:", 100, 3, True, False, str����ԭ��) = False Then
        If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus: Exit Function
    End If
    mBalanceInfor.str����ԭ�� = str����ԭ��
    ������� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetSaveBalanceSQL(ByRef cllBalaceData As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���ν��ʵĽ�����ص�Sql
    '����:cllBalaceData-��������
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-13 11:10:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intInsure As Integer, strSQL As String, dblMoney As Double, dblTemp As Double
    Dim str����IDs As String, str������Ϣ As String, intMaxTime As Integer
    Dim lngTmp As String, str����ID  As String, strTemp As String, strNow As String
    Dim strסԺ���� As String, cllPartBalance As Collection, strArray() As String
    Dim dblAvail As Double, cllTemp As Collection, intCounter As Integer
    Dim i As Long, dblTotal As Double
    
    On Error GoTo errHandle
    Set cllBalaceData = New Collection
    Set cllPartBalance = New Collection
    
    If mBalanceInfor.blnSaveBill = True Then GetSaveBalanceSQL = True: Exit Function
    
    Set cllTemp = New Collection
    '��ǰ������Ϣ
    With mBalanceInfor
        .strNO = zlDatabase.GetNextNo(15)
        .lng����ID = zlDatabase.GetNextId("���˽��ʼ�¼")
        .dtBalanceDate = zlDatabase.Currentdate
    End With
    
    intInsure = mYBInFor.intInsure
    If intInsure <> 0 Then str������Ϣ = IIf(mYBInFor.intInsure = 0, " ", mYBInFor.intInsure) & "," & NVL(mrsInfo!����, " ") & "," & NVL(mrsInfo!ҽ����, " ")
    intMaxTime = 0
    intMaxTime = GetMinMaxTime(1)
    '1.���˽��ʼ�¼
    '����:25596
    ' Zl_���˽��ʼ�¼_Insert
    strSQL = "zl_���˽��ʼ�¼_Insert("
    '  Id_In           ���˽��ʼ�¼.ID%Type,
    strSQL = strSQL & "" & mBalanceInfor.lng����ID & ","
    '  ���ݺ�_In       ���˽��ʼ�¼.NO%Type,
    strSQL = strSQL & "'" & mBalanceInfor.strNO & "',"
    '  ����id_In       ���˽��ʼ�¼.����id%Type,
    strSQL = strSQL & "" & Val(NVL(mrsInfo!����ID)) & ","
    '  �շ�ʱ��_In     ���˽��ʼ�¼.�շ�ʱ��%Type,
    strSQL = strSQL & "To_Date('" & Format(mBalanceInfor.dtBalanceDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
    '  ��ʼ����_In     ���˽��ʼ�¼.��ʼ����%Type,
    strSQL = strSQL & "" & IIf(IsDate(txtPatiBegin.Text), "To_Date('" & txtPatiBegin.Text & "','YYYY-MM-DD')", "NULL") & ","
    '  ��������_In     ���˽��ʼ�¼.��������%Type,
    strSQL = strSQL & "" & IIf(IsDate(txtPatiEnd.Text), "To_Date('" & txtPatiEnd.Text & "','YYYY-MM-DD')", "NULL") & ","
    '  ��;����_In     ���˽��ʼ�¼.��;����%Type := 0,
    strSQL = strSQL & "" & IIf(mEditType = g_Ed_�������, 0, IIf(opt��;.Value, 1, 0)) & ","
    '  �ಡ�˽���_In   Number := 0,
    strSQL = strSQL & "" & 0 & ","
    '  �����ʴ���_In Number := 0,
    strSQL = strSQL & "" & ZVal(intMaxTime) & ","
    '  ��ע_In         ���˽��ʼ�¼.��ע%Type := Null
    strSQL = strSQL & "" & IIf(Trim(txtBalance(Idx_����˵��).Text) = "", "NULL", "'" & Trim(txtBalance(Idx_����˵��).Text) & "'") & ","
    '   ��Դ_In         Number := 1,1-����;2-סԺ
    strSQL = strSQL & "" & IIf(mEditType = g_Ed_�������, 1, 2) & ","
    '  ԭ��_In         ���˽��ʼ�¼.ԭ��%Type := Null
    strSQL = strSQL & "" & IIf(Trim(mBalanceInfor.str����ԭ��) = "", "NULL", "'" & Trim(mBalanceInfor.str����ԭ��) & "'") & ","
    '    ��������_In     ���˽��ʼ�¼.��������%type:=2
    strSQL = strSQL & "" & IIf(mEditType = g_Ed_�������, 1, 2) & ","
    '  ����״̬_In     ���˽��ʼ�¼.����״̬%type:=0
    '����״̬:NULL-�����Ľ�������;1-�쳣�Ľ��ʻ���������;2-�������ϵ��쳣��¼
    strSQL = strSQL & "" & 1 & ","
    ' סԺ����_In     ���˽��ʼ�¼.סԺ����%Type := Null,
    strסԺ���� = ""
    strסԺ���� = mobjBalanceCon.strTime
    If strסԺ���� = "" Then strסԺ���� = mobjBalanceAll.strAllTime
    strSQL = strSQL & "" & IIf(strסԺ���� = "", "NULL", "'" & strסԺ���� & "'") & ","
    ' ���ʽ��_In     ���˽��ʼ�¼.���ʽ��%Type := Null
    strSQL = strSQL & "" & mBalanceInfor.dbl��ǰ���� & ","
    ' Ʊ�ݺ�_In     ���˽��ʼ�¼.ʵ��Ʊ��%Type := Null
    strSQL = strSQL & IIf(mblnPrintInvoice, IIf(txtInvoice.Text = "", "Null)", "'" & txtInvoice.Text & "')"), "Null)")

    zlAddArray cllBalaceData, strSQL
    
    With vsDetailList
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("����")) <> "" _
                And Not (Val(.Cell(flexcpData, i, .ColIndex("���ʽ��"))) = 0 And Val(.Cell(flexcpData, i, .ColIndex("δ����"))) <> 0) Then
                If Val(.TextMatrix(i, .ColIndex("ID"))) = 0 Then
                    '  Zl_���ʷ��ü�¼_Insert
                    strSQL = "Zl_���ʷ��ü�¼_Insert("
                    '  Id_In       סԺ���ü�¼.ID%Type,
                    strSQL = strSQL & "" & Val(.TextMatrix(i, .ColIndex("ID"))) & ","
                    '  No_In       סԺ���ü�¼.NO%Type,
                    strSQL = strSQL & "'" & .TextMatrix(i, .ColIndex("����")) & "',"
                    '  ��¼����_In סԺ���ü�¼.��¼����%Type,
                    strSQL = strSQL & "" & Val(.TextMatrix(i, .ColIndex("��¼����"))) & ","
                    '  ��¼״̬_In סԺ���ü�¼.��¼״̬%Type,
                    strSQL = strSQL & "" & Val(.TextMatrix(i, .ColIndex("��¼״̬"))) & ","
                    '  ִ��״̬_In סԺ���ü�¼.ִ��״̬%Type,
                    strSQL = strSQL & "" & Val(.TextMatrix(i, .ColIndex("ִ��״̬"))) & ","
                    '  ���_In     סԺ���ü�¼.���%Type,
                    strSQL = strSQL & "" & Val(.TextMatrix(i, .ColIndex("���"))) & ","
                    '  ���ʽ��_In סԺ���ü�¼.���ʽ��%Type,
                    strSQL = strSQL & "" & Val(.Cell(flexcpData, i, .ColIndex("���ʽ��"))) & ","
                    '  ����id_In   סԺ���ü�¼.����id%Type
                    strSQL = strSQL & "" & mBalanceInfor.lng����ID & ")"
                    zlAddArray cllTemp, strSQL
                Else
                    If Val(.Cell(flexcpData, i, .ColIndex("���ʽ��"))) = Val(.Cell(flexcpData, i, .ColIndex("δ����"))) Then
                        str����IDs = str����IDs & Val(.TextMatrix(i, .ColIndex("ID"))) & ","
                    Else
                        '  Zl_���ʷ��ü�¼_Insert
                        strSQL = "Zl_���ʷ��ü�¼_Insert("
                        '  Id_In       סԺ���ü�¼.ID%Type,
                        strSQL = strSQL & "" & Val(.TextMatrix(i, .ColIndex("ID"))) & ","
                        '  No_In       סԺ���ü�¼.NO%Type,
                        strSQL = strSQL & "'" & .TextMatrix(i, .ColIndex("����")) & "',"
                        '  ��¼����_In סԺ���ü�¼.��¼����%Type,
                        strSQL = strSQL & "" & Val(.TextMatrix(i, .ColIndex("��¼����"))) & ","
                        '  ��¼״̬_In סԺ���ü�¼.��¼״̬%Type,
                        strSQL = strSQL & "" & Val(.TextMatrix(i, .ColIndex("��¼״̬"))) & ","
                        '  ִ��״̬_In סԺ���ü�¼.ִ��״̬%Type,
                        strSQL = strSQL & "" & Val(.TextMatrix(i, .ColIndex("ִ��״̬"))) & ","
                        '  ���_In     סԺ���ü�¼.���%Type,
                        strSQL = strSQL & "" & Val(.TextMatrix(i, .ColIndex("���"))) & ","
                        '  ���ʽ��_In סԺ���ü�¼.���ʽ��%Type,
                        strSQL = strSQL & "" & Val(.Cell(flexcpData, i, .ColIndex("���ʽ��"))) & ","
                        '  ����id_In   סԺ���ü�¼.����id%Type
                        strSQL = strSQL & "" & mBalanceInfor.lng����ID & ")"
                        zlAddArray cllBalaceData, strSQL
                    End If
                End If
            End If
        Next i
    End With
    
    While str����IDs <> ""
        If Len(str����IDs) > 3998 Then
            lngTmp = InStrRev(Mid(str����IDs, 1, 3998), ",")
            str����ID = Mid(str����IDs, 1, lngTmp - 1)
            str����IDs = Mid(str����IDs, lngTmp + 1)
        Else
            str����ID = Mid(str����IDs, 1, Len(str����IDs) - 1)
            str����IDs = ""
        End If
        strSQL = "zl_���ʷ��ü�¼_Batch('" & str����ID & "'," & mrsInfo!����ID & "," & mBalanceInfor.lng����ID & ")"
        zlAddArray cllBalaceData, strSQL
    Wend
    
    '��������-->��������-->�Ը�����������-->�ٴν���ʱ������ɲ��ܽ��ʵĴ���ԭ������12��13��¼ʱ������˽����ܶ���ʵ���Ƿ�һ�£�����δ��������Ϊ2��3�ļ�¼,���ͳ�ƵĽ��������,
    '�ִ���ʽ:����Ҫ�ȴ������δ��ģ�Ȼ���ٴ���12��13�ļ�¼
    For i = 1 To cllTemp.Count
        zlAddArray cllBalaceData, cllTemp(i)
    Next
    
    GetSaveBalanceSQL = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function



Private Function CheckBusinessRuleIsValied(ByRef objSetFocus As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ҵ������Ƿ�Ϸ�
    '����:objSetFocus-���Ϸ�ʱ,���ȱʡ��λ���ĸ��ؼ�
    '����:�Ϸ�����true,���򷵻�False
    '����:���˺�
    '����:2015-01-12 18:12:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intState As Integer, strTime As String, i As Long, strTmp As String
       
    
    On Error GoTo errHandle
    '�������,ֱ�ӷ���true
    If mYBInFor.bytMCMode <> 1 And mEditType <> g_Ed_������� Then
        intState = GetPatientState
        If mYBInFor.intInsure <> 0 And opt��Ժ.Value Then
            If MCPAR.��Ժ��������Ժ And intState <> 0 Then
                '�����:115055,����,2017/10/16,������ݺϷ���ʱ�ᱨ��
                If IsNull(mrsInfo!��ǰ����id) Then
                    MsgBox "�����ڽ����ڼ䱻������Ժ,ҽ�����˳�Ժ����ǰ�����ȳ�Ժ��", vbInformation, gstrSysName
                Else
                    MsgBox "ҽ�����˳�Ժ����ǰ�����ȳ�Ժ��", vbInformation, gstrSysName
                End If
                Exit Function
            End If
        End If
          
        '�Ƿ���Ժ
        If gTy_System_Para.TY_Balance.bln��Ժ��׼���� And opt��Ժ.Value And (intState = 1 Or intState = 2) Then '  ' 30572:Ԥ��ԺҲ����Ժ.
            MsgBox "��ǰ������Ժ���������Ժ���ʡ�", vbInformation, gstrSysName
            Exit Function
        End If
        
        '����Ƿ��д��շ���δ�˻�����
        If opt��Ժ.Value = True Then
            If PatiHaveStorage(mrsInfo!����ID) Then Exit Function
        End If
                      
        'bytAuditing:0-�����,1-��鲢��ʾ,2-��鲢��ֹ
        '����:37369:��;���ʲ����
        With gTy_System_Para.TY_Balance
            If .bytAuditing <> 0 And opt��Ժ.Value Then
                If HaveNOAuditing(mrsInfo!����ID, mobjBalanceCon.strTime) Then
                    If .bytAuditing = 1 Then
                        '�ڶ�ȡ������Ϣʱ,�Ѿ������
                    ElseIf .bytAuditing = 2 Then
                         Call MsgBox("�ò��˻�����δ��˵ļ��ʷ���,��ֹ����!", vbInformation + vbOKOnly, gstrSysName)
                         Exit Function
                    End If
                End If
            End If
        End With
          
        '��Ҫ�ٴμ��,�Է������ڼ�����˵Ĳ��˱�ȡ�����
        If (InStr(mstrPrivs, ";δ��˲�����;����;") = 0 And opt��;.Value _
            Or InStr(mstrPrivs, ";δ��˲��˳�Ժ����;") = 0 And opt��Ժ.Value) _
                And mEditType = g_Ed_סԺ���� Then
            strTime = IIf(mobjBalanceCon.strTime = "", mobjBalanceAll.strAllTime, mobjBalanceCon.strTime)
            If strTime <> "" Then
                For i = 0 To UBound(Split(strTime, ","))
                    strTmp = Split(strTime, ",")(i)
                    If Val(strTmp) <> 0 Then
                        If Not Chk�������(mrsInfo!����ID, Val(strTmp)) Then
                            MsgBox "�����ʷ����а������˵�" & strTmp & "��סԺδ��˵ķ��ü�¼��" & vbCrLf & _
                                "�㲻�ܶ�δ��˵ķ��ý��н��ʣ�", vbInformation, gstrSysName
                            If cmdRefresh.Visible And cmdRefresh.Enabled Then cmdRefresh.SetFocus
                            Exit Function
                        End If
                    End If
                Next
            End If
        End If
    End If
    
    If mEditType = g_Ed_������� Then
        If CheckMZNotExcuteItemValied(objSetFocus) = False Then Exit Function   '���δִ����Ŀ�Ƿ�Ϸ�
        If CheckMZNotSendDrug = False Then Exit Function '���δ��ҩƷ
    Else
        If opt��Ժ.Value Then
            If CheckNotExcuteItemValied(objSetFocus) = False Then Exit Function   '���δִ����Ŀ�Ƿ�Ϸ�
            If CheckNotSendDrug = False Then Exit Function '���δ��ҩƷ
        End If
    End If
    
    If Not CheckDelAppleyFeeValied Then Exit Function '����˷�����ĺϷ���
    If ������� = False Then Exit Function
    CheckBusinessRuleIsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function GetPatientState() As Integer
'����:��ȡ����״̬
'����:0-��Ժ,1-��Ժ,2-Ԥ��Ժ,-1-�������ݿ����
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    GetPatientState = -1
    On Error GoTo errH
    strSQL = "Select a.��ǰ����id, a.��ҳid As �����ҳid, b.��ҳid, b.״̬" & vbNewLine & _
            "From ������Ϣ A, ������ҳ B" & vbNewLine & _
            "Where a.����id = b.����id And Nvl(b.��ҳid, 0) = (Select Max(Column_Value) From Table(f_str2list([2]))) And b.����id = [1]"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CLng(mrsInfo!����ID), IIf(mobjBalanceCon.strTime = "", mobjBalanceAll.strAllTime, mobjBalanceCon.strTime))
    
    If rsTmp.RecordCount > 0 Then
        If Val(NVL(rsTmp!�����ҳID)) > Val(NVL(rsTmp!��ҳID)) Then
            GetPatientState = 0
        Else
            If Not IsNull(rsTmp!��ǰ����id) Then
                If Val("" & rsTmp!״̬) = 3 Then
                    GetPatientState = 2
                Else
                    GetPatientState = 1
                End If
            Else
                GetPatientState = 0
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub WriteZYInforToCard(ByVal lng����ID As Long, ByVal lng����ID As Long, Optional blnDelete As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��סԺ��Ϣд�뿨��
    '���:blnDelete-�Ƿ��˷�
    '����:���˺�
    '����:2015-01-13 11:04:01
    '����:56615
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngCardTypeID As Long, strExpend As String
    Dim objCard As Card

    On Error GoTo errHandle
        
    'δȷ��ˢ�����,ֱ���˳�
    If mEditType = g_Ed_������� Or mblnCurMzBalanceNo Then
        If InStr(1, mstrPrivs, ";������Ϣд��;") = 0 Then Exit Sub
    Else
        If InStr(1, mstrPrivs, ";סԺ��Ϣд��;") = 0 Then Exit Sub
    End If
    If lng����ID = 0 Then Exit Sub
    
    If mlngCardTypeID = 0 Then
        If blnDelete Then GoTo goDelete:
        Exit Sub
    End If
    
    If IDKIND.GetCurCard.�ӿ���� = mlngCardTypeID Then
        Set objCard = IDKIND.GetCurCard
    Else
        Set objCard = IDKIND.GetIDKindCard(mlngCardTypeID, CardTypeID)
    End If
    
    If objCard Is Nothing Then Exit Sub
    If objCard.�Ƿ�д�� = False Or objCard.�ӿ���� <= 0 Then Exit Sub '��׼д����,�����ýӿ�
    lngCardTypeID = objCard.�ӿ����
goDelete:
    If mEditType = g_Ed_������� Or mblnCurMzBalanceNo Then
        Call gobjSquare.objSquareCard.zlMzInforWriteToCard(Me, mlngModul, lngCardTypeID, _
        lng����ID, lng����ID, strExpend)
    Else
        Call gobjSquare.objSquareCard.zlZyInforWriteToCard(Me, mlngModul, lngCardTypeID, _
        lng����ID, lng����ID, strExpend)
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Public Function ExistWaitDrug(ByVal lng����ID As Long, ByVal lng��ҳID As Long, Optional ByVal lng�����Ժ��ҩ As Long = 0, Optional ByVal int�����־ As Integer) As String
'���ܣ���鲡����ҩ���Ƿ���δ��ҩ��ҩƷ������
'���أ�ҩ���ͷ��ϲ�������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select Zl_Pati_Check_Execute(1,[1],[2],-1,[3],[4]) as ���� From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ExistWaitDrug", lng����ID, lng��ҳID, lng�����Ժ��ҩ, int�����־)
    
    If Not rsTmp.EOF Then
        ExistWaitDrug = NVL(rsTmp!����)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function IsCheck�����ѽ���(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鲡���Ƿ��Ѿ�����
    '���:
    '����:
    '����:�ѽ��շ���True,���򷵻�False
    '����:���˺�
    '����:2010-05-24 16:39:47
    '˵��;
    '     ����:30036
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    strSQL = "select nvl(��Ϣֵ,0) as �������� from ������ҳ�ӱ� where ����id=[1] and ��ҳid=[2] and ��Ϣ��='��������'"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID)
    If Not rsTemp.EOF Then
            IsCheck�����ѽ��� = Val(NVL(rsTemp!��������)) = 1
    Else
            IsCheck�����ѽ��� = False
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function PatiHaveStorage(ByVal lng����ID As Long) As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strMsg As String
    
    strSQL = "Select A.���㷽ʽ,Sum(A.���) as ���" & _
        " From ����Ԥ����¼ A,���㷽ʽ B" & _
        " Where A.��¼����=1 And A.���㷽ʽ=B.���� And B.����=5 And A.����ID=[1]" & _
        " Group by A.���㷽ʽ Having Sum(A.���)<>0"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng����ID)
    If Not rsTmp.EOF Then
        Do While Not rsTmp.EOF
            strMsg = strMsg & vbCrLf & rsTmp!���㷽ʽ & "��" & Format(rsTmp!���, "0.00")
            rsTmp.MoveNext
        Loop
    End If
    If strMsg <> "" Then
        If mty_ModulePara.byt���ʼ����տ��� = 1 Then
            If MsgBox("�������´��շ���û���˻����ˣ�" & vbCrLf & strMsg & vbCrLf & vbCrLf & "Ҫ����������?", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                PatiHaveStorage = True
            Else
                PatiHaveStorage = False
            End If
        Else
            MsgBox "�������´��շ���û���˻����ˣ�" & vbCrLf & strMsg & vbCrLf & vbCrLf & "���Ƚ������˻��������ٽ��ʡ�", vbInformation, gstrSysName
            PatiHaveStorage = True
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function SaveBalaceCharge(ByVal blnԤ�� As Boolean, ByRef tyBrushCard As TY_BrushCard, _
    Optional ByRef blnChargeEnd As Boolean, _
    Optional ByRef objSetFocus As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����������
    '���:blnԤ��-��ǰ�ǽ�Ԥ����
    '����:blnChargeEnd-�շ���ɲ���(true,����շ�;False-��δ���)
    '     objSetFocus-����ʧ��ʱ,ȱʡ��λ���λ��
    '����:�������ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-14 10:35:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str���ѿ����� As String, str�շѽ��� As String, strSQL As String
    Dim strCardNo As String, strErrMsg As String
    Dim blnHaveMoney As Boolean, blnFind As Boolean, blnTrans As Boolean
    Dim dblʣ���� As Double, dblTemp As Double, dblδ����� As Double
    Dim dblMoney As Double, dbl��֧Ʊ�� As Double, bln��Ԥ�� As Boolean
    Dim i As Long, j As Long, varData As Variant, cllDeposit As Collection
    Dim cllUpdate As Collection, cllThreeSwap As Collection, cllPro As Collection
    Dim objCard As Card, lng����ID As Long
    Dim intSign As Integer, rsTmp As ADODB.Recordset
    Dim bytSign As Byte, strסԺ���� As String
    Dim intMousePointer As Integer
    Dim strArray() As String, k As Integer
    intMousePointer = Screen.MousePointer
    
    On Error GoTo errHandle
    
    '��鵱ǰ�����Ƿ���Ч
    blnChargeEnd = False
    Set objCard = IDKindPaymentsType.GetCurCard
    lng����ID = mPatiInfor.lng����ID
    
    With mBalanceInfor
        .dbl�ɿ� = 0: .dbl�Ҳ� = 0
        .dbl�ֽ� = 0
    End With
    
    If Not blnԤ�� Then
        bln��Ԥ�� = IDKind�Ҳ�.GetCurCard.�ӿ���� <> 1
    End If
    
    With mBalanceInfor
        dblδ����� = RoundEx(.dblδ���ϼ� - .dbl��Ԥ���ϼ�, 6)
    End With
    
    intSign = IIf(dblδ����� < 0, -1, 1)
    If blnԤ�� Then 'Ԥ����
        dblMoney = Val(txtBalance(Idx_��Ԥ��).Text)
        If dblMoney <> mBalanceInfor.dbl��Ԥ���ϼ� Then Exit Function
        dblʣ���� = RoundEx(mBalanceInfor.dblδ���ϼ� - mBalanceInfor.dbl��Ԥ���ϼ�, 6)
    ElseIf objCard.�������� = 1 Then '��ǰ�ֽ�֧��
        dblMoney = RoundEx(intSign * Val(txtBalance(Idx_�ɿ�).Text), 6)
        If dblMoney <> 0 Then
            mBalanceInfor.dbl�ɿ� = dblMoney
            If Not bln��Ԥ�� Then
                mBalanceInfor.dbl�Ҳ� = RoundEx(IIf(IDKind�Ҳ�.GetCurCard.���� = "�տ�", 1, -1) * Val(txtBalance(Idx_�Ҳ�).Text), 5)
            End If
        End If
        
        dblTemp = dblδ�����: dblʣ���� = 0
        dblMoney = GetCentMoney(dblTemp)
        mBalanceInfor.dbl�ֽ� = dblMoney
        
        dblʣ���� = 0
    ElseIf objCard.���� Like "*֧Ʊ*" Then
        If mbln�������� Then
            dblMoney = dblδ�����: dblʣ���� = 0
        Else
            dblMoney = RoundEx(intSign * (Val(txtBalance(Idx_�ɿ�).Text) - mPatiInfor.dblδ���ۼ�), 6)
            If dblδ����� < 0 Then
                '֧Ʊ�˿�����
                If Val(dblδ�����) > Val(dblMoney) Then
                    MsgBox "�˿�֧Ʊ��������ܳ�����ǰ�˿���!", vbInformation, gstrSysName
                    Exit Function
                End If
                dblMoney = RoundEx(intSign * Val(txtBalance(Idx_�ɿ�).Text) - mPatiInfor.dblδ���ۼ�, 6)
                dblʣ���� = RoundEx(dblδ����� - dblMoney, 6)
            Else
                dblʣ���� = RoundEx(dblδ����� - dblMoney, 6)
                If Not bln��Ԥ�� Then
                    If dblʣ���� < 0 Then
                        dbl��֧Ʊ�� = RoundEx(-1 * Val(txtBalance(Idx_�Ҳ�).Text), 5)
                        dblʣ���� = 0
                    End If
                Else
                    dblʣ���� = RoundEx(dblʣ���� + Val(txtBalance(Idx_�Ҳ�).Text), 6)
                End If
            End If
        End If
    Else  '�������㷽ʽ֧��
        If mbln�������� Then
            dblMoney = dblδ�����: dblʣ���� = 0
        Else
            dblMoney = RoundEx(intSign * Val(txtBalance(Idx_�ɿ�).Text) - mPatiInfor.dblδ���ۼ�, 6)
            dblʣ���� = RoundEx(dblδ����� - dblMoney, 6)
        End If
    End If
 

    Call Show�����(blnԤ��)
    
    '���ܴ���1.5��Ǯ
    If Abs(mBalanceInfor.dbl����) > 1.5 Then
        Screen.MousePointer = 0
        Call MsgBox("������,�����Ƿ���ȷ!", vbInformation + vbOKOnly, gstrSysName)
        Exit Function
    End If
    
    If dblʣ���� > 0 Then blnHaveMoney = True
  
 
    Set cllPro = New Collection: Set cllUpdate = New Collection
    Set cllThreeSwap = New Collection
    If Not (InStr(",1,2,", "," & objCard.�������� & ",") > 0 Or blnԤ��) Or dblʣ���� = 0 Then
        If GetSaveBalanceSQL(cllPro) = False Then Exit Function
    End If
    
    If Not blnԤ�� Then
        'ִ��һ��ͨ(�ϰ�)�ӿ�
        If ExecuteOldOneCardPayInterface(lng����ID, mBalanceInfor.lng����ID, objCard, dblMoney, tyBrushCard, cllPro) = False Then Exit Function
        'ִ�������ʻ����׽ӿ�
        If tyBrushCard.blnת�� Then
            If ExecuteThreeSwapTransferPay(objCard, dblMoney, cllPro, tyBrushCard) = False Then Exit Function
        Else
            If ExecuteThreeSwapPayInterface(lng����ID, mBalanceInfor.lng����ID, objCard, dblMoney, cllPro, tyBrushCard) = False Then Exit Function
        End If
    End If
    
    If dblʣ���� = 0 And blnԤ�� = False Then
        '��ɽ���
        With vsBlance
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("�����Ϣ")) <> "" And Val(.TextMatrix(i, .ColIndex("����״̬"))) = 0 Then
                    If Val(.TextMatrix(i, .ColIndex("�Ƿ�ת��"))) = 0 Then
                        If Val(.Cell(flexcpData, i, .ColIndex("�����Ϣ"))) = 1 Then
                            If ExecuteThreeSwapDelSingle(IDKindPaymentsType.GetIDKindCard(.TextMatrix(i, .ColIndex("�����ID")), CardTypeID), _
                                                    RoundEx(-1 * Val(.TextMatrix(i, .ColIndex("������"))), 2), .Cell(flexcpData, i, .ColIndex("����")), _
                                                    .TextMatrix(i, .ColIndex("����˵��")), .TextMatrix(i, .ColIndex("������ˮ��")), _
                                                    Val(.TextMatrix(i, .ColIndex("�����Ϣ"))), cllPro) = False Then
                                '�ӿ�ʧ��
                                For k = 1 To vsDeposit.Rows - 1
                                    If Val(vsDeposit.TextMatrix(k, vsDeposit.ColIndex("Ԥ��ID"))) = Val(.TextMatrix(i, .ColIndex("�����Ϣ"))) Then
                                        vsDeposit.Cell(flexcpForeColor, k, 0, k, vsDeposit.Cols - 1) = vbRed
                                    End If
                                Next k
                                Exit Function
                            Else
                                '�ӿڳɹ�
                                For k = 1 To vsDeposit.Rows - 1
                                    If Val(vsDeposit.TextMatrix(k, vsDeposit.ColIndex("Ԥ��ID"))) = Val(.TextMatrix(i, .ColIndex("�����Ϣ"))) Then
                                        vsDeposit.TextMatrix(k, vsDeposit.ColIndex("�༭״̬")) = 1
                                        vsDeposit.Cell(flexcpForeColor, k, 0, k, vsDeposit.Cols - 1) = vbBlack
                                    End If
                                Next k
                            End If
                        Else
                            strArray = Split(.TextMatrix(i, .ColIndex("�����Ϣ")), "|")
                            If ExecuteThreeSwapDelBatch(IDKindPaymentsType.GetIDKindCard(.TextMatrix(i, .ColIndex("�����ID")), CardTypeID), _
                                                         RoundEx(-1 * Val(.TextMatrix(i, .ColIndex("������"))), 2), .TextMatrix(i, .ColIndex("�����Ϣ")), _
                                                        cllPro) = False Then
                                '�ӿ�ʧ��
                                For j = 0 To UBound(strArray)
                                    For k = 1 To vsDeposit.Rows - 1
                                        If Val(vsDeposit.TextMatrix(k, vsDeposit.ColIndex("Ԥ��ID"))) = Val(Split(strArray(j), ",")(4)) Then
                                            vsDeposit.Cell(flexcpForeColor, k, 0, k, vsDeposit.Cols - 1) = vbRed
                                        End If
                                    Next k
                                Next j
                                Exit Function
                            Else
                                '�ӿڳɹ�
                                For j = 0 To UBound(strArray)
                                    For k = 1 To vsDeposit.Rows - 1
                                        If Val(vsDeposit.TextMatrix(k, vsDeposit.ColIndex("Ԥ��ID"))) = Val(Split(strArray(j), ",")(4)) Then
                                            vsDeposit.TextMatrix(k, vsDeposit.ColIndex("�༭״̬")) = 1
                                            vsDeposit.Cell(flexcpForeColor, k, 0, k, vsDeposit.Cols - 1) = vbBlack
                                        End If
                                    Next k
                                Next j
                            End If
                        End If
                        mBalanceInfor.blnSaveBill = True
                        .TextMatrix(i, .ColIndex("�����Ϣ")) = ""
                        .TextMatrix(i, .ColIndex("����״̬")) = 1
                        .TextMatrix(i, .ColIndex("�༭״̬")) = 0
                    Else
                        If CheckThreeSwapValied(IDKindPaymentsType.GetIDKindCard(.TextMatrix(i, .ColIndex("�����ID")), CardTypeID), Val(.TextMatrix(i, .ColIndex("������"))), tyBrushCard, True) = False Then Exit Function
                        If ExecuteThreeSwapTransferPay(IDKindPaymentsType.GetIDKindCard(.TextMatrix(i, .ColIndex("�����ID")), CardTypeID), Val(.TextMatrix(i, .ColIndex("������"))), cllPro, tyBrushCard) = False Then Exit Function
                        mBalanceInfor.blnSaveBill = True
                        .TextMatrix(i, .ColIndex("�����Ϣ")) = ""
                        .TextMatrix(i, .ColIndex("����״̬")) = 1
                        .TextMatrix(i, .ColIndex("�༭״̬")) = 0
                    End If
                End If
            Next i
        End With
        '�����Ԥ��
        If GetSaveStrickDepositSQL(cllDeposit, objSetFocus) = False Then Exit Function
        For i = 1 To cllDeposit.Count
            cllPro.Add cllDeposit(i)
        Next
        If ExcuteBalanceEnd(dbl��֧Ʊ��, cllPro) = False Then Exit Function
        
        If opt��Ժ.Value = True And mEditType = g_Ed_סԺ���� Then
            '��Ժ����,����Ƿ����
            Set rsTmp = GetMoneyInfo(mrsInfo!����ID, , , 2)
            If Not rsTmp Is Nothing Then
                '����,�����Զ����ʱ�־
                If NVL(rsTmp!�������, 0) = 0 Then
                    strסԺ���� = ""
                    strסԺ���� = mobjBalanceCon.strTime
                    If strסԺ���� = "" Then strסԺ���� = mobjBalanceAll.strAllTime
                    If strסԺ���� <> "" Then
                        strSQL = "zl_�����Զ�����_Stop(" & mrsInfo!����ID & ",'" & strסԺ���� & "')"
                        zlDatabase.ExecuteProcedure strSQL, Me.Caption
                    End If
                End If
            End If
        End If
        
        '��ӡƱ��
        Call PrintBill(mPatiInfor.lng����ID, mBalanceInfor.strNO, mBalanceInfor.lng����ID, mBalanceInfor.dtBalanceDate, mBalanceInfor.dbl�ɿ�, mBalanceInfor.dbl�Ҳ�)
        '81697:���ϴ�,2015/6/8,������
        If Not gobjPlugIn Is Nothing Then
            On Error Resume Next
            Call gobjPlugIn.InPatiCashierAfter(mPatiInfor.lng����ID, mBalanceInfor.lng����ID)
            Err.Clear
        End If
        
        If Not mbln�������� Then
            Call ExcutePatiOutHosptial '���˳�Ժ
        End If
        'סԺ��Ϣд��:56615
        Call WriteZYInforToCard(mPatiInfor.lng����ID, mBalanceInfor.lng����ID)
'        If mPatiInfor.bln��Ժ Then
        zlDatabase.SetPara "Ĭ�ϳ�Ժ����", IIf(opt��Ժ.Value = True, "1", "0"), glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0
'        End If
        
        blnChargeEnd = True
        If mEditType = g_Ed_���½��� Then Unload Me: Exit Function
        SaveBalaceCharge = True
        Exit Function
    End If
NextBalance:
    Err = 0: On Error GoTo errHandle:
    If blnԤ�� Then GoTo GoEnd:
    If dblMoney <> 0 Then
        With vsBlance
            If objCard.���ѿ� Then
                Call AddSquareBalance(objCard)
            Else
                If .Rows = 1 Then .Rows = 2
                If Not (.Rows = 2 And Trim(.TextMatrix(1, .ColIndex("���㷽ʽ"))) = "") Then
                    .Rows = .Rows + 1
'                    .RowPosition(.Rows - 1) = 1
                End If
                strCardNo = tyBrushCard.str����
                
                .TextMatrix(.Rows - 1, .ColIndex("�Ƿ�����")) = 0
                '����:0-��ͨ����;.rows-1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
                If objCard.�������� = 7 And objCard.�ӿ���� < 0 Then
                    .TextMatrix(.Rows - 1, .ColIndex("����")) = 4
                    .TextMatrix(.Rows - 1, .ColIndex("�༭״̬")) = 0 '0-��ֹɾ��;1-����༭���;2-����ɾ��
                    .TextMatrix(.Rows - 1, .ColIndex("����״̬")) = 1 '�Ƿ��ѽ���:1-�ѽ���;0-δ����
                ElseIf objCard.�ӿ���� > 0 Then
                    .TextMatrix(.Rows - 1, .ColIndex("����")) = 3
                    .TextMatrix(.Rows - 1, .ColIndex("�����ID")) = objCard.�ӿ����
                    .TextMatrix(.Rows - 1, .ColIndex("���������")) = objCard.����
                    .TextMatrix(.Rows - 1, .ColIndex("�༭״̬")) = 0 '0-��ֹɾ��;1-����༭���;2-����ɾ��
                    .TextMatrix(.Rows - 1, .ColIndex("����״̬")) = 1 '�Ƿ��ѽ���:1-�ѽ���;0-δ����
                    .TextMatrix(.Rows - 1, .ColIndex("�Ƿ�����")) = IIf(objCard.�������Ĺ��� <> "", 1, 0)
                Else
                    .TextMatrix(.Rows - 1, .ColIndex("����")) = 0
                    .TextMatrix(.Rows - 1, .ColIndex("�༭״̬")) = 2 '0-��ֹɾ��;1-����༭���;2-����ɾ��
                    .TextMatrix(.Rows - 1, .ColIndex("����״̬")) = 0 '�Ƿ��ѽ���:1-�ѽ���;0-δ����
                End If
                .TextMatrix(.Rows - 1, .ColIndex("��������")) = objCard.��������
                .TextMatrix(.Rows - 1, .ColIndex("�Ƿ�����")) = IIf(objCard.�Ƿ�����, 1, 0)
                .TextMatrix(.Rows - 1, .ColIndex("�Ƿ�ȫ��")) = IIf(objCard.�Ƿ�ȫ��, 1, 0)
                .TextMatrix(.Rows - 1, .ColIndex("У�Ա�־")) = 2
                
                .TextMatrix(.Rows - 1, .ColIndex("���㷽ʽ")) = objCard.���㷽ʽ
                .TextMatrix(.Rows - 1, .ColIndex("���㷽ʽ")) = objCard.���㷽ʽ
                .TextMatrix(.Rows - 1, .ColIndex("������")) = Format(dblMoney, "0.00")
                .TextMatrix(.Rows - 1, .ColIndex("�������")) = IIf(txtBalance(Idx_�������).Visible, Trim(txtBalance(Idx_�������).Text), "")
                .TextMatrix(.Rows - 1, .ColIndex("��ע")) = Trim(txtBalance(Idx_ժҪ).Text)
                .TextMatrix(.Rows - 1, .ColIndex("����")) = IIf(objCard.�������Ĺ��� <> "", String(Len(strCardNo), "*"), strCardNo)
                .Cell(flexcpData, .Rows - 1, .ColIndex("����")) = tyBrushCard.str����
                .TextMatrix(.Rows - 1, .ColIndex("������ˮ��")) = tyBrushCard.str������ˮ��
                .TextMatrix(.Rows - 1, .ColIndex("����˵��")) = tyBrushCard.str����˵��
                mBalanceInfor.dbl�Ѹ��ϼ� = RoundEx(mBalanceInfor.dbl�Ѹ��ϼ� + dblMoney, 6)
                mBalanceInfor.dblδ���ϼ� = RoundEx(mBalanceInfor.dblδ���ϼ� - dblMoney, 6)
            End If
            If Not mbln�������� Then
                For i = 1 To IDKindPaymentsType.ListCount
                    'ȱʡ��λ���ֽ���
                     Set objCard = IDKindPaymentsType.GetIDKindCard(i, CardTypeIndex)
                    If objCard.�������� = 1 Then IDKindPaymentsType.IDKIND = i: Exit For
                Next
            End If
        End With
    End If

GoEnd:
    Set objSetFocus = txtBalance(Idx_�ɿ�)
    txtBalance(Idx_�ɿ�).Text = ""
    Call LoadCurOwnerPayInfor
    Call LedDisplayBank
    
    SaveBalaceCharge = True
    Exit Function
errHandle:
    Screen.MousePointer = 0
    If blnTrans Then
        gcnOracle.RollbackTrans
        Call ErrCenter
    Else
        If ErrCenter() = 1 Then
            Screen.MousePointer = intMousePointer
            Resume
        End If
    End If
End Function
 

Public Function Get���ѿ����㷽ʽ() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���ѿ�������Ϣ
    '����:���˺�
    '����:2015-01-30 15:29:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str���ѿ����� As String, i As Long
    Dim objCard As Card
    Dim lngCardTypeID As Long
    
    On Error GoTo errHandle
    
    Set objCard = IDKindPaymentsType.GetCurCard
    lngCardTypeID = 0
    If Not objCard Is Nothing Then
        If objCard.�ӿ���� > 0 And objCard.���ѿ� And Val(txtBalance(Idx_�ɿ�).Text) <> 0 Then
            '��ǰˢ�����ѿ�,��ʹ���ڲ���¼��
            lngCardTypeID = objCard.�ӿ����
        End If
    End If
 
    str���ѿ����� = ""  '�����ID|����|���ѿ�ID|���ѽ��||....
    With vsBlance
       For i = 1 To .Rows - 1
           '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
           '�༭״̬:0-��ֹɾ��;1-����༭���;2-����ɾ��
           '����״̬:�Ƿ��ѽ���:1-�ѽ���;0-δ����
           If Val(.TextMatrix(i, .ColIndex("����"))) = 5 _
                And Val(.TextMatrix(i, .ColIndex("����״̬"))) = 0 _
                And lngCardTypeID <> Val(.TextMatrix(i, .ColIndex("�����ID"))) Then
                str���ѿ����� = str���ѿ����� & "||" & Val(.TextMatrix(i, .ColIndex("�����ID")))
                str���ѿ����� = str���ѿ����� & "|" & Trim(.Cell(flexcpData, i, .ColIndex("����")))
                str���ѿ����� = str���ѿ����� & "|" & Val(.TextMatrix(i, .ColIndex("���ѿ�ID")))
                str���ѿ����� = str���ѿ����� & "|" & RoundEx(Val(.TextMatrix(i, .ColIndex("������"))), 6)
           End If
       Next
    End With
    If lngCardTypeID > 0 Then
        For i = 1 To mcllCurSquareBalance.Count
            'array(�����ID,���ѿ�ID,ˢ�����,����,����,�������,�Ƿ�����))
            str���ѿ����� = str���ѿ����� & "||" & Val(mcllCurSquareBalance(i)(0))
            str���ѿ����� = str���ѿ����� & "|" & Trim(mcllCurSquareBalance(i)(3))
            str���ѿ����� = str���ѿ����� & "|" & Val(mcllCurSquareBalance(i)(1))
            str���ѿ����� = str���ѿ����� & "|" & RoundEx(Val(mcllCurSquareBalance(i)(2)), 6)
        Next
    End If
    If str���ѿ����� <> "" Then str���ѿ����� = Mid(str���ѿ�����, 3)
    Get���ѿ����㷽ʽ = str���ѿ�����
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Public Function Get��ͨ���㷽ʽ() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�շѽ�������
    '����:�շ��ý��㷽ʽ,��ʽ����:
    '       ���㷽ʽ|������|�������|����ժҪ||.....",ע���޽�������ժҪʱҪ�ÿո����
    '����:���˺�
    '����:2015-01-14 16:11:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str���㷽ʽ As String, i As Long, int���� As Integer
    Dim strBalance As String, dblMoney As Double, varData As Variant
    Dim objCard As Card, objTempCard As Card
    Dim bln��Ԥ�� As Boolean
    '���㷽ʽ|������|�������|����ժҪ||.....",ע���޽�������ժҪʱҪ�ÿո����
    '�շ����
    strBalance = ""
    With vsBlance
        For i = .Rows - 1 To 1 Step -1
            str���㷽ʽ = Trim(.TextMatrix(i, .ColIndex("���㷽ʽ")))
            '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
            int���� = Val(.TextMatrix(i, .ColIndex("����")))
            If str���㷽ʽ <> "" And int���� = 0 Then
                strBalance = strBalance & "||" & str���㷽ʽ
                strBalance = strBalance & "|" & Val(.TextMatrix(i, .ColIndex("������")))
                strBalance = strBalance & "|" & IIf(Trim(.TextMatrix(i, .ColIndex("�������"))) = "", " ", Trim(.TextMatrix(i, .ColIndex("�������"))))
                strBalance = strBalance & "|" & IIf(Trim(.TextMatrix(i, .ColIndex("��ע"))) = "", " ", Trim(.TextMatrix(i, .ColIndex("��ע"))))
            End If
        Next
        Set objCard = IDKindPaymentsType.GetCurCard
        Set objTempCard = IDKind�Ҳ�.GetCurCard
        
        bln��Ԥ�� = Not objTempCard Is Nothing
        If bln��Ԥ�� Then
            bln��Ԥ�� = objTempCard.�ӿ���� > 1
        End If
        
        If objCard.�ӿ���� <= 0 Then
            If mbln�������� Then
                dblMoney = Format(mBalanceInfor.dblδ���ϼ� - mBalanceInfor.dbl��Ԥ���ϼ�, mstrDec)
            Else
                If mPatiInfor.bln�˿��־ Then
                    dblMoney = Format(-1 * Val(txtBalance(Idx_�ɿ�).Text) - mPatiInfor.dblδ���ۼ�, mstrDec)
                Else
                    dblMoney = Format(Val(txtBalance(Idx_�ɿ�).Text) - mPatiInfor.dblδ���ۼ�, mstrDec)
                End If
            End If
            If objCard.�������� = 1 Then
                dblMoney = mBalanceInfor.dbl�ֽ�
            ElseIf bln��Ԥ�� Then
                dblMoney = RoundEx(Val(lblʣ���Ը�.Tag), 6)
            End If
            
            If dblMoney <> 0 Then
                strBalance = strBalance & "||" & objCard.���㷽ʽ
                If objCard.�������� = 1 Then
                    '�ֽ�
                    strBalance = strBalance & "|" & dblMoney
                    strBalance = strBalance & "| "
                    strBalance = strBalance & "| " & IIf(Trim(txtBalance(Idx_ժҪ).Text) = "", " ", Trim(txtBalance(Idx_ժҪ).Text))
                Else
                    strBalance = strBalance & "|" & dblMoney
                    strBalance = strBalance & "|" & IIf(Trim(txtBalance(Idx_�������).Text) = "", " ", Trim(Idx_�������))
                    strBalance = strBalance & "| " & IIf(Trim(txtBalance(Idx_ժҪ).Text) = "", " ", Trim(txtBalance(Idx_ժҪ).Text))
                End If
            End If
        End If
    End With
    If strBalance <> "" Then strBalance = Mid(strBalance, 3)
    Get��ͨ���㷽ʽ = strBalance
    
End Function
 
Private Function ExcuteBalanceEnd(ByVal dbl��֧Ʊ As Double, _
    ByVal cllBillPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ʽ�������
    '���:dbl��֧Ʊ-��ǰ��֧Ʊ���
    '     cllPro-��������
    '����:ִ�гɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-14 16:06:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, cllPro As Collection, strSQL As String
    Dim lng����ID As Long, str��ͨ���� As String, str���ѿ����� As String
    Dim dblԤ�� As Double, intԤ����� As Integer
    Dim lng����ID As Long, lng��ҳID As Long
    Dim dblMoney As Double
    
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    
    str��ͨ���� = Get��ͨ���㷽ʽ
    str���ѿ����� = Get���ѿ����㷽ʽ
    
    lng����ID = mPatiInfor.lng����ID
    lng����ID = mBalanceInfor.lng����ID
    lng��ҳID = mPatiInfor.lng��ҳID
    
    On Error GoTo errHandle
    
    If str���ѿ����� <> "" Then
        '����֮ǰ,�ȴ�������
        'Zl_���˽��ʽ���_Modify
        strSQL = "Zl_���˽��ʽ���_Modify("
        '  ��������_In     Number,
        '  --��������_In:
        '--   3-���ѿ�����:
        '--     �ٽ��㷽ʽ_IN:����һ��ˢ���ſ�,��ʽΪ:�����ID|����|���ѿ�ID|���ѽ��||."  ���ѿ�ID:Ϊ��ʱ,���ݿ����Զ���λ
        strSQL = strSQL & "3,"
        '  ����id_In       ���˽��ʼ�¼.����id%Type,
        strSQL = strSQL & "" & lng����ID & ","
        '  ����id_In       ����Ԥ����¼.����id%Type,
        strSQL = strSQL & "" & lng����ID & ","
        '  ���㷽ʽ_In     Varchar2,
        strSQL = strSQL & "'" & str���ѿ����� & "',"
        '  ��Ԥ��_In       ����Ԥ����¼.��Ԥ��%Type := Null,
        strSQL = strSQL & "NULL,"
        '  ��֧Ʊ��_In     ����Ԥ����¼.��Ԥ��%Type := Null,
        strSQL = strSQL & "NULL,"
        '  �����id_In     ����Ԥ����¼.�����id%Type := Null,
        strSQL = strSQL & "NULL,"
        '  ����_In         ����Ԥ����¼.����%Type := Null,
        strSQL = strSQL & "NULL,"
        '  ������ˮ��_In   ����Ԥ����¼.������ˮ��%Type := Null,
        strSQL = strSQL & "NULL,"
        '  ����˵��_In     ����Ԥ����¼.����˵��%Type := Null,
        strSQL = strSQL & "NULL,"
        '  �ɿ�_In         ����Ԥ����¼.�ɿ�%Type := Null,
        strSQL = strSQL & "NULL,"
        '  �Ҳ�_In         ����Ԥ����¼.�Ҳ�%Type := Null,
        strSQL = strSQL & "NULL,"
        '  �����_In     ������ü�¼.ʵ�ս��%Type := Null,
        strSQL = strSQL & "NULL,"
        '  ��������_In     Number := 2,
        strSQL = strSQL & "" & IIf(mEditType = g_Ed_������� Or mblnCurMzBalanceNo, 1, 2) & ","
        '  ȱʡ���㷽ʽ_In ���㷽ʽ.����%Type := Null,
        strSQL = strSQL & "NULL,"
        '    ����Ա���_In    ����Ԥ����¼.����Ա���%Type := Null,
        strSQL = strSQL & "'" & UserInfo.��� & "',"
        '    ����Ա����_In    ����Ԥ����¼.����Ա����%Type := Null,
        strSQL = strSQL & "'" & UserInfo.���� & "',"
        '    �տ�ʱ��_In      ����Ԥ����¼.����Ա����%Type := Null,
        strSQL = strSQL & "to_date('" & Format(mBalanceInfor.dtBalanceDate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
        '    ��Ԥ������ids_In Varchar2 := Null,
        strSQL = strSQL & "NULL,"
        '  ��ɽ���_In Number:=0
        strSQL = strSQL & "" & IIf(str��ͨ���� = "", 1, "0") & " )"
        zlAddArray cllPro, strSQL
    End If
    
    If str��ͨ���� <> "" Or str���ѿ����� = "" Then
         '����֮ǰ,�ȴ�������
         'Zl_���˽��ʽ���_Modify
        strSQL = "Zl_���˽��ʽ���_Modify("
        '  ��������_In     Number,
        '  --��������_In:
        '--   0-��ͨ�շѷ�ʽ:
        '--     �ٽ��㷽ʽ_IN:��������,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ||.." ;Ҳ�������.
        strSQL = strSQL & "0,"
        '  ����id_In       ���˽��ʼ�¼.����id%Type,
        strSQL = strSQL & "" & lng����ID & ","
        '  ����id_In       ����Ԥ����¼.����id%Type,
        strSQL = strSQL & "" & lng����ID & ","
        '  ���㷽ʽ_In     Varchar2,
        strSQL = strSQL & "'" & str��ͨ���� & "',"
        '  ��Ԥ��_In       ����Ԥ����¼.��Ԥ��%Type := Null,
        strSQL = strSQL & "NULL,"
        '  ��֧Ʊ��_In     ����Ԥ����¼.��Ԥ��%Type := Null,
        strSQL = strSQL & "" & dbl��֧Ʊ & ","
        '  �����id_In     ����Ԥ����¼.�����id%Type := Null,
        strSQL = strSQL & "NULL,"
        '  ����_In         ����Ԥ����¼.����%Type := Null,
        strSQL = strSQL & "NULL,"
        '  ������ˮ��_In   ����Ԥ����¼.������ˮ��%Type := Null,
        strSQL = strSQL & "NULL,"
        '  ����˵��_In     ����Ԥ����¼.����˵��%Type := Null,
        strSQL = strSQL & "'" & IIf(InStr(mstrForceNote, "ǿ������") + 4 = Len(mstrForceNote), "", mstrForceNote) & "',"
        '  �ɿ�_In         ����Ԥ����¼.�ɿ�%Type := Null,
        strSQL = strSQL & "" & mBalanceInfor.dbl�ɿ� & ","
        '  �Ҳ�_In         ����Ԥ����¼.�Ҳ�%Type := Null,
        strSQL = strSQL & "" & mBalanceInfor.dbl�Ҳ� & ","
        '  �����_In     ������ü�¼.ʵ�ս��%Type := Null,
        strSQL = strSQL & "" & mBalanceInfor.dbl���� & ","
        '  ��������_In     Number := 2,
        strSQL = strSQL & "" & IIf(mEditType = g_Ed_������� Or mblnCurMzBalanceNo, 1, 2) & ","
        '  ȱʡ���㷽ʽ_In ���㷽ʽ.����%Type := Null,
         strSQL = strSQL & "NULL,"
        '    ����Ա���_In    ����Ԥ����¼.����Ա���%Type := Null,
        strSQL = strSQL & "'" & UserInfo.��� & "',"
        '    ����Ա����_In    ����Ԥ����¼.����Ա����%Type := Null,
        strSQL = strSQL & "'" & UserInfo.���� & "',"
        '    �տ�ʱ��_In      ����Ԥ����¼.����Ա����%Type := Null,
        strSQL = strSQL & "to_date('" & Format(mBalanceInfor.dtBalanceDate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
        '    ��Ԥ������ids_In Varchar2 := Null,
        strSQL = strSQL & "NULL,"
        
        '  ��ɽ���_In Number:=0
        strSQL = strSQL & "1)"
        zlAddArray cllPro, strSQL
    End If
    If GetSaveAddDepositSQL(lng����ID, lng��ҳID, mBalanceInfor.lng����ID, cllPro) = False Then Exit Function
    
    '�쳣��¼ʱ�䴦��
    If mEditType = g_Ed_���½��� Then
        strSQL = "Zl_���˽����쳣_Update("
        strSQL = strSQL & "to_date('" & Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
        strSQL = strSQL & "" & lng����ID & ")"
        zlAddArray cllPro, strSQL
    End If
    
    Err = 0: On Error GoTo ErrTrans:
    zlExecuteProcedureArrAy cllPro, Me.Caption
    If mbln�������� Then
        mPatiInfor.dblδ���ۼ� = RoundEx(mPatiInfor.dblδ���ۼ� + Val(lblʣ���Ը�.Tag), 6)
        mPatiInfor.bln�������� = mbln��������
        Set mPatiInfor.objCard = IDKindPaymentsType.GetCurCard
    Else
        mPatiInfor.dblδ���ۼ� = 0
        mPatiInfor.bln�������� = False
        Set mPatiInfor.objCard = Nothing
    End If
    
    ExcuteBalanceEnd = True
    Exit Function
ErrTrans:
    gcnOracle.RollbackTrans
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetSaveAddDepositSQL(ByVal lng����ID As Long, lng��ҳID As Long, _
     ByVal lng����ID As Long, ByRef cllPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����Ԥ����SQL
    '���:lng����ID-����Ԥ����Ӧ�Ľ���ID
    '����:cllPro-������Ԥ����SQL���ӵ��ü�����
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-30 13:46:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card, intԤ����� As Integer, str���㷽ʽ As String
    Dim dblMoney As Double, strSQL As String
    
    On Error GoTo errHandle
    
    Set objCard = IDKind�Ҳ�.GetCurCard
    If objCard Is Nothing Then GetSaveAddDepositSQL = True: Exit Function
    If objCard.�ӿ���� <= 1 Then GetSaveAddDepositSQL = True: Exit Function
    If IDKindPaymentsType.GetCurCard Is Nothing Then Exit Function
        
    str���㷽ʽ = IDKindPaymentsType.GetCurCard.���㷽ʽ
    
    intԤ����� = objCard.�ӿ���� - 1    '1-����Ԥ��;2-סԺԤ��
    
    '��ΪԤ����
    dblMoney = RoundEx(Val(txtBalance(Idx_�Ҳ�).Text), 6)
    If dblMoney < 0 Then Exit Function
    
    mBalanceInfor.lngԤ��ID = zlDatabase.GetNextId("����Ԥ����¼")
    mBalanceInfor.strԤ��No = zlDatabase.GetNextNo(11)
    
    'Zl_����Ԥ����¼_Insert
    strSQL = "Zl_����Ԥ����¼_Insert("
    '  Id_In         ����Ԥ����¼.ID%Type,
    strSQL = strSQL & "" & mBalanceInfor.lngԤ��ID & ","
    '  ���ݺ�_In     ����Ԥ����¼.NO%Type,
    strSQL = strSQL & "'" & mBalanceInfor.strԤ��No & "',"
    '  Ʊ�ݺ�_In     Ʊ��ʹ����ϸ.����%Type,
    strSQL = strSQL & "'" & mstrDepositInvioce & "',"
    '  ����id_In     ����Ԥ����¼.����id%Type,
    strSQL = strSQL & "" & lng����ID & ","
    '  ��ҳid_In     ����Ԥ����¼.��ҳid%Type,:42329
    If intԤ����� = 2 Then
       strSQL = strSQL & "" & IIf(lng��ҳID = 0, "NULL", lng��ҳID) & ","
    Else
       strSQL = strSQL & "NULL,"
    End If
    '  ����id_In     ����Ԥ����¼.����id%Type,
    strSQL = strSQL & "" & UserInfo.����ID & ","
    
    '  ���_In       ����Ԥ����¼.���%Type,
    strSQL = strSQL & "" & dblMoney & ","
    '  ���㷽ʽ_In   ����Ԥ����¼.���㷽ʽ%Type,
    strSQL = strSQL & "'" & str���㷽ʽ & "',"
    '  �������_In   ����Ԥ����¼.�������%Type,
    strSQL = strSQL & "'" & txtBalance(Idx_�������).Text & "',"
    '  �ɿλ_In   ����Ԥ����¼.�ɿλ%Type,
    strSQL = strSQL & "Null,"
    '  ��λ������_In ����Ԥ����¼.��λ������%Type,
    strSQL = strSQL & "Null,"
    '  ��λ�ʺ�_In   ����Ԥ����¼.��λ�ʺ�%Type,
    strSQL = strSQL & "Null,"
    '  ժҪ_In       ����Ԥ����¼.ժҪ%Type,
    strSQL = strSQL & "'���ʴ�Ԥ��',"
    '  ����Ա���_In ����Ԥ����¼.����Ա���%Type,
    strSQL = strSQL & "'" & UserInfo.��� & "',"
    '  ����Ա����_In ����Ԥ����¼.����Ա����%Type,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '  ����id_In     Ʊ��ʹ����ϸ.����id%Type,
    strSQL = strSQL & "" & IIf(mlngԤ������ID = 0, "NULL", mlngԤ������ID) & ","
    '  Ԥ�����_In   ����Ԥ����¼.Ԥ�����%Type := Null,
    strSQL = strSQL & "" & intԤ����� & ","
    '  �����id_In   ����Ԥ����¼.�����id%Type := Null,
    strSQL = strSQL & "Null,"
    '  ���㿨���_in ����Ԥ����¼.���㿨���%type:=NULL,
    strSQL = strSQL & "Null,"
    '  ����_In       ����Ԥ����¼.����%Type := Null,
    strSQL = strSQL & "Null,"
    '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
    strSQL = strSQL & "Null,"
    '  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
    strSQL = strSQL & "Null,"
    '  ������λ_In   ����Ԥ����¼.������λ%Type := Null,
    strSQL = strSQL & "Null,"
    '  �տ�ʱ��_In   ����Ԥ����¼.�տ�ʱ��%Type := Null
    strSQL = strSQL & "to_date('" & mBalanceInfor.dtBalanceDate & "','yyyy-mm-dd hh24:mi:ss'),"
    '   ��������_In Integer:=0 :0-������Ԥ��;1-��Ϊ���۵�
    strSQL = strSQL & "0,"
    '  ����id_In     ����Ԥ����¼.����id%Type >0ʱ,��ʾĳ�ν���ʱ,ͬ��������Ԥ����¼
    strSQL = strSQL & "" & lng����ID & ","
    '  ��������_In     ����Ԥ����¼.��������%Type >0ʱ,���ʲ�����Ԥ����,����Ϊ2
    strSQL = strSQL & "" & 12 & ")"
    zlAddArray cllPro, strSQL
    GetSaveAddDepositSQL = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function zlCheckMulitInterfaceNumValied(Optional blnԤ�� As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ͬʱ�����������Ͻӿ�(��������)
    '����:�����������Ͻӿڵ�,����true,���򷵻�False
    '����:���˺�
    '����:2012-02-07 15:07:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intCount As Integer, i As Long, int���� As Integer, str���㷽ʽ As String
    Dim varData As Variant, strErrMsg As String
    Dim objCard As Card
    Dim intMousePointer As Integer
    On Error GoTo errHandle
    strErrMsg = ""
    intMousePointer = Screen.MousePointer
    Set objCard = IDKindPaymentsType.GetCurCard
        
    If blnԤ�� Or objCard.�ӿ���� <= 0 Then zlCheckMulitInterfaceNumValied = True:        Exit Function
    
   'ҽ����һ���ӿ�
   If mYBInFor.intInsure <> 0 And mBalanceInfor.blnSaveBill Then intCount = intCount + 1: strErrMsg = strErrMsg & "ҽ������:" & Format(mBalanceInfor.dblҽ��֧���ϼ�, gstrDec)
   With vsBlance
        For i = 1 To .Rows - 1
            str���㷽ʽ = Trim(.TextMatrix(i, .ColIndex("���㷽ʽ")))
            '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
 
            int���� = Val(.TextMatrix(i, .ColIndex("����")))
            If InStr("34", int����) > 0 Then
                intCount = intCount + 1: strErrMsg = strErrMsg & vbCrLf & str���㷽ʽ & ":" & .TextMatrix(i, .ColIndex("������"))
            End If
        Next
    End With
    If intCount > 2 Then
        Screen.MousePointer = 0
        Call MsgBox("ע��:" & vbCrLf & "   ��ϵͳĿǰֻ֧���������½ӿ�,�����Ѿ��������½ӿڽ���:" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName)
        Exit Function
    End If
    zlCheckMulitInterfaceNumValied = True
    Exit Function
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = intMousePointer
        Resume
    End If
End Function

Private Sub Show�����(ByVal blnԤ�� As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ�����
    '���:blnԤ��-Ԥ����
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-14 11:33:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, dbl��֧���� As Double
    Dim dblʣ���� As Double, dblTemp As Double, dblδ���� As Double
    Dim intSign As Integer, objCard As Card
    If mEditType = g_Ed_���ݲ鿴 Then Exit Sub
    
    With mBalanceInfor
        .dbl���� = 0
        dblδ���� = RoundEx(.dblδ���ϼ� - .dbl��Ԥ���ϼ�, 6)
        intSign = IIf(dblδ���� < 0, -1, 1)
    End With
    
    dblMoney = IIf(blnԤ��, Val(txtBalance(Idx_��Ԥ��).Text), intSign * Val(txtBalance(Idx_�ɿ�).Text))
    
    dbl��֧���� = 0: dblʣ���� = RoundEx(dblδ���� - dblMoney, 6)
    
    If blnԤ�� Then
        '����Ԥ��ʱ
        mBalanceInfor.dbl���� = RoundEx(dblδ���� - RoundEx(dblδ����, 2), 6): GoTo Show���:
        Exit Sub
    End If
    Set objCard = IDKindPaymentsType.GetCurCard
    If Not objCard Is Nothing Then
        If objCard.�������� = 1 Then
            dblTemp = dblδ����: dblʣ���� = 0
            dblMoney = GetCentMoney(dblTemp)
            mBalanceInfor.dbl���� = RoundEx(dblδ���� - dblMoney, 6)
            GoTo Show���:
        End If
    End If
    mBalanceInfor.dbl���� = RoundEx(dblδ���� - RoundEx(dblδ����, 2), 6): GoTo Show���:
    
    If mYBInFor.intInsure <> 0 And mBalanceInfor.blnSaveBill = False Then mBalanceInfor.dbl���� = 0
Show���:
    lbl����.Visible = mBalanceInfor.dbl���� <> 0 And mEditType <> g_Ed_ȡ������
    lbl����.Caption = "���:" & RoundEx(mBalanceInfor.dbl����, 6)
End Sub
Private Function CheckSquareBalanceValied(ByVal objCard As Card, ByRef tyBrushCard As TY_BrushCard) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ѿ����㽻�׼��
    '���:objCard-������
    '����:dblMoney-��ǰˢ�����
    '����:���׺Ϸ�����true,���򷵻�False
    '����:���˺�
    '����:2014-07-08 18:00:34
    '˵��:ͬ����֤�˽ӿں�ˢ���ӿڵ�
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, dblTemp As Double
    Dim rsMoney As ADODB.Recordset, strXMLExpend As String
    Dim strTittle As String, dbl�ʻ���� As Double
    Dim strBrushCard As TY_BrushCard, cllSquareBalance As Collection
    Dim strExpand As String, bln���� As Boolean, dblδ����� As Double
    Dim intMousePointer As Integer, strXmlIn As String
    Dim lng���ѿ�ID As Long, str���� As String, str���� As String
    Dim str������� As String, byt�Ƿ�����   As Byte
    Dim cllBushSquare As Collection, i As Long
    
    
    intMousePointer = Screen.MousePointer

    If objCard.�ӿ���� <= 0 Or objCard.���ѿ� = False Then CheckSquareBalanceValied = True: Exit Function
    
    On Error GoTo errHandle
    
    tyBrushCard = strBrushCard
    
    dblMoney = Val(txtBalance(Idx_�ɿ�).Text)
    
    If dblMoney = 0 Then CheckSquareBalanceValied = True: Exit Function
    dblδ����� = RoundEx(mBalanceInfor.dblδ���ϼ� - mBalanceInfor.dbl��Ԥ���ϼ�, 6)
    
    If dblMoney = 0 And dblδ����� <> 0 Then
        Screen.MousePointer = 0
        MsgBox "�տ���δ����,����!", vbInformation + vbOKOnly, gstrSysName
         Exit Function
    End If
    If dblMoney > Format(dblδ�����, "0.00") And dblMoney <> 0 Then
        Screen.MousePointer = 0
        MsgBox "�տ���ܴ��ڱ���δ�����:" & Format(dblδ�����, "0.00") & " ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    '�ȼ���Ӧ�Ľӿ�
    If mEditType = g_Ed_������� Or mEditType = g_Ed_סԺ���� Then
        If zlGetClassMoney(0, rsMoney) = False Then Exit Function
    Else
        If zlGetClassMoney(mBalanceInfor.lng����ID, rsMoney) = False Then Exit Function
    End If
    
     '�������ѿ���ˢ����Ϣ
     Set cllSquareBalance = New Collection
     Set mcllCurSquareBalance = New Collection
     With vsBlance
        For i = 1 To .Rows - 1
            '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
            '�༭״̬:0-��ֹɾ��;1-����༭���;2-����ɾ��
            '����״̬:�Ƿ��ѽ���:1-�ѽ���;0-δ����
            If Val(.TextMatrix(i, .ColIndex("����"))) = 5 _
                And Val(.TextMatrix(i, .ColIndex("�����ID"))) = objCard.�ӿ���� _
                And Val(.TextMatrix(i, .ColIndex("����״̬"))) = 0 Then
  
                dblTemp = RoundEx(Val(.TextMatrix(i, .ColIndex("������"))), 6)
                lng���ѿ�ID = Val(.TextMatrix(i, .ColIndex("���ѿ�ID")))
                str���� = Trim(.Cell(flexcpData, i, .ColIndex("����")))
                str���� = Trim(.Cell(flexcpData, i, .ColIndex("���ѿ�ID")))  '����
                str������� = Trim(.Cell(flexcpData, i, .ColIndex("�����ID")))  '�������
                byt�Ƿ����� = Val(.TextMatrix(i, .ColIndex("�Ƿ�����")))
                'array(�����ID,���ѿ�ID,ˢ�����,����,����,�������,�Ƿ�����)
                cllSquareBalance.Add Array(objCard.�ӿ����, lng���ѿ�ID, dblTemp, str����, str����, str�������, byt�Ƿ�����)
            End If
        Next
     End With
     For i = 1 To cllSquareBalance.Count
        mcllCurSquareBalance.Add cllSquareBalance(i)
     Next
     
    'zlBrushCard(frmMain As Object, _
        ByVal lngModule As Long, _
        ByVal rsClassMoney As ADODB.Recordset, _
        ByVal lngCardTypeID As Long, _
        ByVal bln���ѿ� As Boolean, _
        ByVal strPatiName As String, ByVal strSex As String, _
        ByVal strOld As String, ByRef dbl��� As Double, _
        Optional ByRef strCardNo As String, _
        Optional ByRef strPassWord As String, _
        Optional ByRef bln�˷� As Boolean = False, _
        Optional ByRef blnShowPatiInfor As Boolean = False, _
        Optional ByRef bln���� As Boolean = False, _
        Optional ByVal bln�����ֹ As Boolean = True, _
        Optional ByRef varSquareBalance As Variant, _
        Optional ByVal blnתԤ�� As Boolean = False, _
        Optional ByVal blnAllPay As Boolean = False, _
        Optional ByVal strXmlIn As String = "", _
        Optional ByVal bytҵ�񳡺� As Byte = 1, _
        Optional ByVal str������Դ As String, _
        Optional ByVal lng����ID As Long) As Boolean
    'varSquareBalance- Collection����,���ص�ǰˢ������(array(�����ID,���ѿ�ID,ˢ�����,����,����,�������,�Ƿ�����))
    '       str������Դ - ��ǰ֧�����õķ�����Դ�������ö��ŷָ�(ʹ�����ѿ�֧��ʱ����)
    '       lng����ID - ����ID(ʹ�����ѿ�֧��ʱ����)
    strXmlIn = "<IN><CZLX>0</CZLX></IN>"
    If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, rsMoney, _
            objCard.�ӿ����, objCard.���ѿ�, _
            "" & mPatiInfor.str����, "" & mPatiInfor.str�Ա�, "" & mPatiInfor.str����, IIf(mPatiInfor.bln�˿��־, -1, 1) * dblMoney, _
            tyBrushCard.str����, tyBrushCard.str����, _
            False, True, False, False, cllSquareBalance, False, False, strXmlIn, _
            GetFeeFromType(), mPatiInfor.lng����ID) = False Then Exit Function
    
    '����ǰ,һЩ���ݼ��
    'zlPaymentCheck(frmMain As Object, ByVal lngModule As Long, _
    ByVal strCardTypeID As Long, ByVal strCardNo As String, _
    ByVal dblMoney As Double, ByVal strNOs As String, _
    Optional ByVal strXMLExpend As String
    If gobjSquare.objSquareCard.zlPaymentCheck(Me, mlngModul, objCard.�ӿ����, _
        objCard.���ѿ�, tyBrushCard.str����, dblMoney, "", strXMLExpend) = False Then Exit Function
        
    '    zlGetAccountMoney(ByVal frmMain As Object, ByVal lngModule As Long, _
    '    ByVal strCardTypeID As Long, _
    '    ByVal strCardNo As String, strExpand As String, dblMoney As Double
    '���:frmMain-���õ�������
    '        lngModule-ģ���
    '        strCardNo-����
    '        strExpand-Ԥ����Ϊ��,�Ժ���չ
    '����:dblMoney-�����ʻ����
    'If gobjSquare.objSquareCard.zlGetAccountMoney(Me, mlngModul, objCard.�ӿ����, _
          tyBrushCard.str����, strExpand, dbl�ʻ����, objCard.���ѿ�) = False Then Exit Function
    '�Ѿ������˽�����
    
    ''array(�����ID,���ѿ�ID,ˢ�����,����,����,�������,�Ƿ�����)
    For i = 1 To cllSquareBalance.Count
       mcllCurSquareBalance.Add cllSquareBalance(i)
    Next
    
    dblMoney = 0
    For i = 1 To mcllCurSquareBalance.Count
       dblMoney = dblMoney + Val(mcllCurSquareBalance(i)(2))
    Next
        
    
    
    If RoundEx(dblMoney, 6) <> Val(txtBalance(Idx_�ɿ�).Text) Then
        txtBalance(Idx_�ɿ�).Text = Format(dblMoney, "0.00")
    End If
    CheckSquareBalanceValied = True
    Exit Function
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = intMousePointer
        Resume
    End If
    Call SaveErrLog
End Function
Private Function CheckSquareDelValied(ByVal objCard As Card, _
     ByRef tyBrushCard As TY_BrushCard, _
     Optional ByVal lng���ѿ�ID As Long, _
     Optional dblDelMoney As Double _
     ) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ѿ��˷Ѽ��
    '���:objCard-������
    '     dblDelMoney-�˿���
    '����:tyBrushCard-����ˢ������
    '����:���׺Ϸ�����true,���򷵻�False
    '����:���˺�
    '����:2015-01-23 11:07:58
    '˵��:ͬ����֤�˽ӿں�ˢ���ӿ�
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, dblTemp As Double
    Dim rsMoney As ADODB.Recordset, strXMLExpend As String
    Dim strTittle As String, dbl�ʻ���� As Double
    Dim cllSquareBalance As Collection
    Dim strExpand As String, bln���� As Boolean
    Dim dblTotal As Double, dblBrushMoney As Double
    Dim cllBalance As Collection, strXmlIn As String
    Dim varData As Variant, varTemp As Variant, i As Long, j As Long
    Dim rsBalance As ADODB.Recordset
    On Error GoTo errHandle
    
    If objCard.�ӿ���� <= 0 Or objCard.���ѿ� = False Then CheckSquareDelValied = True: Exit Function
     
    If zlGetClassMoney(mBalanceInfor.lng����ID, rsMoney) = False Then Exit Function
    If dblDelMoney = 0 Then
        If Val(txtBalance(Idx_�ɿ�).Text) = 0 Then
            MsgBox "δ�����˷ѽ��δ����,����!", vbInformation + vbOKOnly, gstrSysName
             Exit Function
        End If
    End If
     
    '�˿���
    If Not mrsOldBalance Is Nothing Then
        Set rsBalance = mrsOldBalance 'ԭ��¼��
    Else
        Set rsBalance = mrsBalance
    End If
    
    If rsBalance Is Nothing Then
        MsgBox "ע��:" & vbCrLf & "  δ�ҵ�ԭʼ�Ľ����¼,����ʹ��" & objCard.���� & "�����˿�!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If rsBalance.State <> 1 Then
        MsgBox "ע��:" & vbCrLf & "  δ�ҵ�ԭʼ�Ľ����¼,����ʹ��" & objCard.���� & "�����˿�!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
    If lng���ѿ�ID <> 0 Then
        rsBalance.Filter = "����=5 And ���㿨���=" & objCard.�ӿ���� & " And ���ѿ�ID=" & lng���ѿ�ID
    Else
        rsBalance.Filter = "����=5 And ���㿨���=" & objCard.�ӿ����
    End If
    
    If rsBalance.EOF Then
        MsgBox "ע��:" & vbCrLf & "  δ�ҵ�ԭʼ�Ľ����¼,����ʹ��" & objCard.���� & "�����˿�!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If dblDelMoney <> 0 Then
        dblMoney = dblDelMoney
    Else
        dblMoney = Val(txtBalance(Idx_�ɿ�).Text)
    End If
    
    dblTotal = 0
    Set cllSquareBalance = New Collection
    Set cllBalance = New Collection
    Set mcllCurSquareBalance = New Collection
    dblTemp = dblMoney
    
    With rsBalance
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            dblTotal = dblTotal + Val(NVL(!��Ԥ��))
            
            'dblBrushMoney = GetSquareBrushMoney(objCard.�ӿ����, Val(Nvl(!���ѿ�ID)), Nvl(!����))
            'array(�����ID,���ѿ�ID,ˢ�����, ����,����,�������,�Ƿ�����,ʣ��δ�˽��)
            cllSquareBalance.Add Array(objCard.�ӿ����, Val(NVL(!���ѿ�ID)), _
             0, NVL(!����), "", "", 0, Val(NVL(!��Ԥ��)))
            
            If dblTemp > Val(NVL(!��Ԥ��)) And dblTemp <> 0 Then
                cllBalance.Add Array(objCard.�ӿ����, Val(NVL(!���ѿ�ID)), _
                Format(Val(NVL(!��Ԥ��)), "0.00"), NVL(!����), "", "", 0)
                dblTemp = dblTemp - Val(NVL(!��Ԥ��))
            ElseIf dblTemp <> 0 Then
                cllBalance.Add Array(objCard.�ӿ����, Val(NVL(!���ѿ�ID)), _
                Format(dblTemp, "0.00"), NVL(!����), "", "", 0)
                dblTemp = 0
            End If
            .MoveNext
        Loop
    End With
    
    If RoundEx(dblTotal, 6) < RoundEx(dblMoney, 6) Then
        MsgBox "ע��:" & vbCrLf & "   ������˿��������" & objCard.���㷽ʽ & "��δ�˽��,����!" & vbCrLf & _
               "   δ�˽��:" & Format(dblTotal, "###0.00;-###0.00;;") & vbCrLf & _
               "   ��ǰ�˿�:" & Format(dblMoney, "###0.00;-###0.00;;"), vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If RoundEx(dblTotal, 6) <> RoundEx(dblMoney, 6) Then
        If objCard.�Ƿ�ȫ�� Then
            MsgBox "ע��:" & vbCrLf & "   " & objCard.���㷽ʽ & "����ȫ��,����!" & vbCrLf & _
                   "   δ�˽��:" & Format(dblTotal, "###0.00;-###0.00;;") & vbCrLf & _
                   "   ��ǰ�˿�:" & Format(dblMoney, "###0.00;-###0.00;;"), vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
    
    If gbln���ѿ��˷��鿨 Then
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
        strXmlIn = "<IN><CZLX>2</CZLX></IN>"
        If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModul, rsMoney, objCard.�ӿ����, _
            objCard.���ѿ�, mPatiInfor.str����, mPatiInfor.str�Ա�, _
            mPatiInfor.str����, dblMoney, "", "", _
            True, True, False, False, cllSquareBalance, False, False, strXmlIn) = False Then Exit Function
        Set cllBalance = cllSquareBalance
    End If
    For i = 1 To cllBalance.Count
        varData = cllBalance(i)
        dblTemp = Val(varData(2)) + dblTemp
        mcllCurSquareBalance.Add varData
    Next
    
    If dblDelMoney = 0 Then
        txtBalance(Idx_�ɿ�).Text = Format(dblTemp, "0.00")
    End If
    CheckSquareDelValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetSquareBrushMoney(ByVal lngCardTypeID As Long, ByVal lng���ѿ�ID As Long, ByVal strCardNo As String) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���ѿ���ˢ�����
    '���:lngCardTypeId-���ѿ��ӿڱ��
    '     lng���ѿ�ID-���ѿ�ID
    '     strCardNo-����
    '����:
    '����:����ˢ�����
    '����:���˺�
    '����:2014-08-12 11:51:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim j As Long, varTemp As Variant
    Dim dblMoney As Double, lngRow As Long
    Dim lng�����ID As Long, lng���ѿ�ID1 As Long, strBalance As String
    dblMoney = 0
    With vsBlance
        For lngRow = 1 To .Rows - 1
            lng�����ID = Val(.TextMatrix(lngRow, .ColIndex("�����ID")))
            lng���ѿ�ID1 = Val(.TextMatrix(lngRow, .ColIndex("���ѿ�ID")))
            strBalance = .TextMatrix(lngRow, .ColIndex("���㷽ʽ"))
            If Val(.TextMatrix(lngRow, .ColIndex("����"))) = 5 And strBalance <> "" Then
                If lngCardTypeID = lng�����ID And (lng���ѿ�ID1 = lng���ѿ�ID Or lng���ѿ�ID = 0) Then
                    dblMoney = RoundEx(dblMoney + Val(.TextMatrix(lngRow, .ColIndex("������"))), 2)
                End If
            End If
        Next
    End With
    GetSquareBrushMoney = dblMoney
End Function
Private Function zlGetClassMoney(ByRef lng����ID As Long, ByRef rsMoney As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ʱ,��ʼ��֧�����(�շ����,ʵ�ս��)
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-06-10 17:52:18
    '����:38841
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String, dblMoney As Double
    Dim dblTemp As Double
    
    On Error GoTo errHandle
    
    '��ʼ�����ݽṹ
    Set rsMoney = New ADODB.Recordset
    rsMoney.Fields.Append "�շ����", adVarChar, 10, adFldIsNullable
    rsMoney.Fields.Append "���", adDouble, , adFldIsNullable
    rsMoney.CursorLocation = adUseClient
    rsMoney.LockType = adLockOptimistic
    rsMoney.CursorType = adOpenStatic
    rsMoney.Open
        
    If lng����ID <> 0 Then
        strSQL = "" & _
        "   Select  A.�շ����,nvl(sum(A.���ʽ��) ,0) as ���   " & _
        "   From ������ü�¼ A" & _
        "   Where A.����ID=[1] Group by A.�շ���� " & _
        "   Union ALL " & _
        "   Select  A.�շ����,nvl(sum(A.���ʽ��) ,0) as ���   " & _
        "   From סԺ���ü�¼ A" & _
        "   Where A.����ID=[1] Group by A.�շ���� "
        strSQL = "Select �շ����,Sum(���) as ��� From (" & strSQL & ")  Group by  �շ����"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    
        With rsTemp
            If .RecordCount <> 0 Then .MoveFirst
            Do While Not .EOF
                rsMoney.Find "�շ����='" & NVL(!�շ����, "��") & "'", , adSearchForward, 1
                If rsMoney.EOF Then rsMoney.AddNew
                rsMoney!�շ���� = NVL(!�շ����, "��")
                rsMoney!��� = Val(NVL(rsMoney!���)) + Val(NVL(!���))
                rsMoney.Update
                .MoveNext
            Loop
        End With
        zlGetClassMoney = True
        Exit Function
    End If
  
    With mrsFeeList
        dblMoney = mBalanceInfor.dbl��ǰ����
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            dblTemp = Val(NVL(!δ����))
            If RoundEx(dblMoney - dblTemp, gbytDec) <= 0 Then
                dblTemp = dblMoney
            End If
            If dblTemp <> 0 And dblMoney <> 0 Then
                rsMoney.Find "�շ����='" & NVL(!�շ����, "��") & "'", , adSearchForward, 1
                If rsMoney.EOF Then rsMoney.AddNew
                rsMoney!�շ���� = NVL(!�շ����, "��")
                rsMoney!��� = Val(NVL(rsMoney!���)) + dblTemp
                rsMoney.Update
            End If
            dblMoney = dblMoney - dblTemp
            .MoveNext
        Loop
    End With
    zlGetClassMoney = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub LedDisplayBank(Optional ByVal blnLedAsked As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ������Ϣ
    '����:���˺�
    '����:2011-12-15 13:40:46
    '����:52117
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl���ʺϼ� As Double, i As Long
    Dim strҽ�� As String, str�������� As String, str��һ��ͨ As String, str��ͨ���� As String
    Dim varPara  As Variant, str���㷽ʽ As String, varData As Variant
    If Not gblnLED Then Exit Sub
    
    
    With vsBlance
        For i = 1 To .Rows - 1
            'ҽ������
            If .TextMatrix(i, .ColIndex("���㷽ʽ")) <> "" Then
                '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
                Select Case Val(.TextMatrix(i, .ColIndex("����")))
                Case 1 'ҽ��
                    strҽ�� = strҽ�� & "||" & .TextMatrix(i, .ColIndex("���㷽ʽ")) & ":" & Format(Val(.TextMatrix(i, .ColIndex("������"))), "0.00")
                Case 2 '�����ӿڽ���
                    str�������� = str�������� & "||" & .TextMatrix(i, .ColIndex("���㷽ʽ")) & ":" & Format(Val(.TextMatrix(i, .ColIndex("������"))), "0.00")
                Case 3   ' һ��ͨ����
                    str��һ��ͨ = str��һ��ͨ & "||" & .TextMatrix(i, .ColIndex("���㷽ʽ")) & ":" & Format(Val(.TextMatrix(i, .ColIndex("������"))), "0.00")
                Case Else
                    str��ͨ���� = str��ͨ���� & "||" & .TextMatrix(i, .ColIndex("���㷽ʽ")) & ":" & Format(Val(.TextMatrix(i, .ColIndex("������"))), "0.00")
                End Select
            End If
        Next
    End With
     
    str���㷽ʽ = ""
    If strҽ�� <> "" Then str���㷽ʽ = str���㷽ʽ & "||ҽ������:||�ʻ����:" & Format(mYBInFor.cur�������, "0.00") & strҽ��
    If str�������� <> "" Then str���㷽ʽ = str���㷽ʽ & "||һ��ͨ����:" & str��������
    If str��һ��ͨ <> "" Then str���㷽ʽ = str���㷽ʽ & "||һ��ͨ����(��):" & str��һ��ͨ
    If str��ͨ���� <> "" Then str���㷽ʽ = str���㷽ʽ & "" & str��ͨ����
    If str���㷽ʽ = "" Then Exit Sub
    str���㷽ʽ = Mid(str���㷽ʽ, 3)
    varPara = Split(str���㷽ʽ, "||")
    
    'Ŀǰ���ֻ����ʾ10������ֵ
    Select Case UBound(varPara)
    Case 0
          zl9LedVoice.DisplayBank varPara(0)
    Case 1
          zl9LedVoice.DisplayBank varPara(0), varPara(1)
    Case 2
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2)
    Case 3
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3)
    Case 4
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4)
    Case 5
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4), varPara(5)
    Case 6
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4), varPara(5), varPara(6)
    Case 7
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4), varPara(5), varPara(6), varPara(7)
    Case 8
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4), varPara(5), varPara(6), varPara(7), varPara(8)
    Case 9
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4), varPara(5), varPara(6), varPara(7), varPara(8), varPara(9)
    Case Else
        str���㷽ʽ = ""
         For i = 10 To UBound(varPara)
            str���㷽ʽ = str���㷽ʽ & ";" & varPara(i)
        Next
        If str���㷽ʽ > "" Then str���㷽ʽ = Mid(str���㷽ʽ, 2)
        zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4), varPara(5), varPara(6), varPara(7), varPara(8), varPara(9), str���㷽ʽ
    End Select
    If blnLedAsked = False Then
        If Format(mBalanceInfor.dblԤ�����ܶ�, gstrDec) <> Format(mBalanceInfor.dblҽ��֧���ϼ�, gstrDec) Then
            '����㲻һ��ʱ,��Ҫ�ٴ�����
            zl9LedVoice.Speak "#21 " & Format(Val(lblʣ���Ը�.Caption), "0.00")
        End If
    End If
End Sub

Private Sub SetOperatonCommandCaption()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ò����ؼ���Caption
    '����:���˺�
    '����:2015-01-21 16:11:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim EditType As gBalanceBill
    
    EditType = mEditType
    If chkCancel.Value = 1 Then EditType = g_Ed_��������
    
    Select Case EditType
    Case g_Ed_��������
        cmdOK.Caption = "ȷ��(&O)"
        cmdCancel.Caption = "ȡ��(&C)"
        lblBalance(3).Caption = "�� Ԥ ��"
    Case g_Ed_ȡ������
        cmdOK.Caption = "����(&O)"
        cmdCancel.Caption = "ȡ��(&C)"
        lblBalance(3).Caption = "�� Ԥ ��"
    Case g_Ed_��������
        cmdOK.Caption = "ȷ��(&O)"
        cmdCancel.Caption = "ȡ��(&C)"
        lblBalance(3).Caption = "�� Ԥ ��"
    Case Else
        cmdOK.Caption = "��ɽ���(&O)"
        cmdCancel.Caption = "ȡ������(&C)"
        lblBalance(3).Caption = "�� Ԥ ��"
    End Select
    Call picBalanceBack_Resize
End Sub
Private Function GetLocalePayCard(ByVal lng�����ID As Long, _
    ByVal bln���ѿ� As Boolean, Optional ByRef intKindIdex As Integer) As Card
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ���Ŀ�����
    '����:intKindIdex-IDkind������
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-22 15:58:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card, i As Long
    
    On Error GoTo errHandle
    intKindIdex = -1
    For i = 1 To IDKindPaymentsType.ListCount
        Set objCard = IDKindPaymentsType.GetIDKindCard(i, CardTypeIndex)
          If objCard Is Nothing Then Exit Function
        If lng�����ID = objCard.�ӿ���� And objCard.���ѿ� = bln���ѿ� Then
            intKindIdex = i
            Set GetLocalePayCard = objCard: Exit Function
        End If
    Next
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetLocaleOldOneCard(ByVal str���㷽ʽ As String, _
     Optional ByRef intKindIdex As Integer) As Card
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ������һ��ͨ�Ŀ�����
    '���:str���㷽ʽ-�ϰ�һ��ͨ�Ľ��㷽
    '����:intKindIdex-IDkind������
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-22 15:58:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card, i As Long
    On Error GoTo errHandle
    intKindIdex = -1
    For i = 1 To IDKindPaymentsType.ListCount
        Set objCard = IDKindPaymentsType.GetIDKindCard(i, CardTypeIndex)
        If objCard.���㷽ʽ = str���㷽ʽ Then
            intKindIdex = i
            Set GetLocaleOldOneCard = objCard: Exit Function
        End If
    Next
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CancelIsValied(ByVal objCard As Card, ByRef tyBrushCard As TY_BrushCard) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���Ͻ��ʼ��
    '���:objCard-��ǰ������
    '����:tyBrushCard-��ǰˢ����Ϣ
    '����:���ݺϷ�����true,���򷵻�False
    '����:���˺�
    '����:2015-01-22 15:28:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, varTemp As Variant
    Dim dblʣ���� As Double, bln�˿� As Boolean
    Dim dblMoney As Double, strBalance As String, i As Long
    Dim dblCurMoney As Double, objTemp As Card
    Dim strSquares As String, cllSquare As Collection 'array(ID,����)
    Dim lngCardTypeID As Long
    On Error GoTo errHandle
      
    If mYBInFor.intInsure > 0 Then
        If Not MCPAR.��Ժ���˽������� And mYBInFor.bytMCMode <> 1 Then
            If Not isYBPati(mPatiInfor.lng����ID, True) Then
                MsgBox "�òα������Ѿ���Ժ������ȡ���ý��ʵ���", vbInformation, gstrSysName: Exit Function
            End If
        End If
        If gclsInsure.CheckInsureValid(mYBInFor.intInsure) = False Then Exit Function
    End If
    
    With mBalanceInfor
        dblʣ���� = RoundEx(.dblδ���ϼ� - .dbl��Ԥ���ϼ�, 5)
        bln�˿� = dblʣ���� > 0
    End With
    
    dblCurMoney = Val(txtBalance(Idx_�ɿ�).Text)
    Set cllSquare = New Collection
    With vsBlance
        For i = 1 To .Rows - 1
            strBalance = Trim(.TextMatrix(i, .ColIndex("���㷽ʽ")))
            If strBalance <> "" Then
               If dblCurMoney <> 0 And objCard.���㷽ʽ = strBalance Then
                    MsgBox "���˿��б����Ѿ����ڡ�" & strBalance & "�����˿ʽ," & vbCrLf & _
                           "������ʹ�ø��˿ʽ!", vbInformation + vbOKOnly, gstrSysName
                    If txtBalance(Idx_�ɿ�).Enabled And txtBalance(Idx_�ɿ�).Visible Then txtBalance(Idx_�ɿ�).SetFocus
                    Exit Function
               End If
               
                '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
                dblMoney = RoundEx(Val(.TextMatrix(i, .ColIndex("������"))), 6)
                Select Case Val(.TextMatrix(i, .ColIndex("����")))
                Case 0 '��ͨ����
                Case 1 'Ԥ����
                Case 2 'ҽ��
                Case 3 'һ��ͨ
                    'ҽ�ƿ����ID|���ѿ�(1, 0) |���ƿ�|�Ƿ�ȫ��|�Ƿ�����|�ӿ�����
                    If Val(.TextMatrix(i, .ColIndex("����״̬"))) = 0 Then
                        Set objTemp = GetLocalePayCard(Val(.TextMatrix(i, .ColIndex("�����ID"))), False)
                        If objTemp Is Nothing Then
                            MsgBox "��ǰվ�㲻֧��" & strBalance & "��ʽ֧��!", vbInformation + vbOKOnly, gstrSysName
                            Exit Function
                        End If
                        If Val(.TextMatrix(i, .ColIndex("������"))) > 0 And (mEditType = g_Ed_�������� Or mEditType = g_Ed_��������) Then bln�˿� = True
                        If CheckThreeSwapValied(objTemp, dblMoney, tyBrushCard, bln�˿�) = False Then Exit Function
                    End If
                Case 4 'һ��ͨ(�ϰ汾)
                    Set objTemp = GetLocaleOldOneCard(strBalance)
                    If objTemp Is Nothing Then
                        MsgBox "��ǰվ�㲻֧��" & strBalance & "��ʽ֧��!", vbInformation + vbOKOnly, gstrSysName
                        Exit Function
                    End If
                    If Val(.TextMatrix(i, .ColIndex("������"))) > 0 And (mEditType = g_Ed_�������� Or mEditType = g_Ed_��������) Then bln�˿� = True
                    If CheckOldOneCardIsValied(objTemp, dblMoney, tyBrushCard, bln�˿�) = False Then Exit Function
                Case 5 '���ѿ�
                    lngCardTypeID = Val(.TextMatrix(i, .ColIndex("�����ID")))
                    If InStr(strSquares & ",", "," & lngCardTypeID & ",") = 0 Then
                        strSquares = strSquares & "," & lngCardTypeID
                        cllSquare.Add Array(lngCardTypeID, strBalance)
                    End If
                Case Else
                End Select
            End If
        Next
    End With
    For i = 1 To cllSquare.Count
        'ҽ�ƿ����ID|���ѿ�(1, 0) |���ƿ�|�Ƿ�ȫ��|�Ƿ�����|�ӿ�����
        Set objTemp = GetLocalePayCard(Val(cllSquare(i)(0)), True)
        If objTemp Is Nothing Then
            MsgBox "��ǰվ�㲻֧��" & cllSquare(i)(1) & "��ʽ֧��!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        dblMoney = GetSquareBrushMoney(Val(cllSquare(i)(0)), 0, "")
        If CheckSquareDelValied(objTemp, tyBrushCard, 0, dblMoney) = False Then Exit Function
        Call AddSquareBalance(objTemp)
    Next
    '----------------------------------------------------------------
    '��ǰˢ�����
    
    '��ǰ�Ѿ���ȫ�������,ֱ�ӷ���
    If dblCurMoney = 0 And dblʣ���� = 0 Then CancelIsValied = True: Exit Function
    
    '�ֽ���
    If CheckCashValied(objCard, bln�˿�) = False Then Exit Function
    If objCard.�������� = 1 Then CancelIsValied = True: Exit Function
    
    
    If dblCurMoney = 0 Then
        MsgBox "��ǰ" & IIf(bln�˿�, "�˿�", "�տ�") & "���δ����!", vbInformation + vbOKOnly, gstrSysName
        If txtBalance(Idx_�ɿ�).Enabled And txtBalance(Idx_�ɿ�).Visible Then txtBalance(Idx_�ɿ�).SetFocus
        Exit Function
    End If
    
    '֧Ʊ���
    If CheckChequeValied(objCard) = False Then Exit Function
    
    '���ѿ����
    If bln�˿� Then
        If CheckSquareDelValied(objCard, tyBrushCard, 0, dblCurMoney) = False Then Exit Function
    Else
        If CheckSquareBalanceValied(objCard, tyBrushCard) = False Then Exit Function
    End If
            
    '������ˢ�������ѿ�,ֱ�ӷ���true
    If objCard.�ӿ���� <= 0 Or objCard.���ѿ� Then CancelIsValied = True: Exit Function
    
    
    '�������׼��
    If CheckThreeSwapValied(objCard, dblCurMoney, tyBrushCard, bln�˿�) = False Then Exit Function
    '�ϰ�һ��ͨ���
    If CheckOldOneCardIsValied(objCard, dblCurMoney, tyBrushCard, bln�˿�) = False Then Exit Function
    CancelIsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function
Private Function GetCancelBalance(ByVal bytFun As Byte, ByRef strBalances As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�������ϵ���ͨ���㷽ʽ
    '���:bytFun-0-��ͨ;1-ҽ��;2-���ѿ�
    '����:
    '    bytfunc=0:strBalances�ĸ�ʽ:���㷽ʽ|������|�������||...
    '    bytfunc=1:strBalances�ĸ�ʽ:���㷽ʽ|������||...
    '    bytfunc=2:strBalances�ĸ�ʽ:�����ID|����|���ѿ�ID|���ѽ��||.
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-22 16:20:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strPTBalance As String, i As Long, dblMoney As Double
    Dim strYbBalance As String, strBalance As String, varData As Variant
    Dim strXFBalance As String
    
    On Error GoTo errHandle
    With vsBlance
        '�ռ��˿ʽ�����
        strPTBalance = "": strYbBalance = "": strXFBalance = ""
        For i = 1 To .Rows - 1
            dblMoney = -1 * RoundEx(Val(.TextMatrix(i, .ColIndex("������"))), 6)
            strBalance = Trim(.TextMatrix(i, .ColIndex("���㷽ʽ")))
            If strBalance <> "" Then
                '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
                Select Case Val(.TextMatrix(i, .ColIndex("����")))
                Case 0 '��ͨ����
                    '���㷽ʽ|������|�������|����ժҪ||..
                    strPTBalance = strPTBalance & "||" & strBalance
                    strPTBalance = strPTBalance & "|" & dblMoney
                    strPTBalance = strPTBalance & "|" & IIf(.TextMatrix(i, .ColIndex("�������")) = "", " ", .TextMatrix(i, .ColIndex("�������")))
                    strPTBalance = strPTBalance & "|" & IIf(.TextMatrix(i, .ColIndex("��ע")) = "", " ", .TextMatrix(i, .ColIndex("��ע")))
                Case 1 'Ԥ����
                Case 2 'ҽ��
                    '���㷽ʽ|������||...
                    strYbBalance = strYbBalance & "||" & .TextMatrix(i, .ColIndex("���㷽ʽ")) & "|" & dblMoney
                Case 3 'һ��ͨ
                Case 4 'һ��ͨ(�ϰ汾)
                Case 5 '���ѿ�
                    '�����ID|����|���ѿ�ID|���ѽ��||.
                    strXFBalance = strXFBalance & "||" & Val(.TextMatrix(i, .ColIndex("�����ID")))
                    strXFBalance = strXFBalance & "|" & Trim(.Cell(flexcpData, i, .ColIndex("����")))
                    strXFBalance = strXFBalance & "|" & Val(.TextMatrix(i, .ColIndex("���ѿ�ID")))
                    strXFBalance = strXFBalance & "|" & dblMoney
                Case Else
                End Select
            End If
        Next
    End With
    If strPTBalance <> "" Then strPTBalance = Mid(strPTBalance, 3)
    If strYbBalance <> "" Then strYbBalance = Mid(strYbBalance, 3)
    If strXFBalance <> "" Then strXFBalance = Mid(strXFBalance, 3)
    
    If bytFun = 0 Then
        strBalances = strPTBalance
    ElseIf bytFun = 1 Then
        strBalances = strYbBalance
    Else
       strBalances = strXFBalance
    End If
    GetCancelBalance = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Private Function ExecuteBalaceCancel() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ�н���ȡ������
    '����:ִ�гɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-21 16:25:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPro As Collection, objCard As Card, i As Long
    Dim strYbBalance As String '���㷽ʽ|���||...
    Dim strSQL As String, strCardNo As String
    Dim lng����ID As Long
    Dim dblʣ���� As Double, bln�˿� As Boolean
    Dim dblCurMoney As Double, dblMoney As Double
    Dim tyBrushCardInfor As TY_BrushCard
    Dim dblTemp As Double
    Dim objBackCard As Card
    
    On Error GoTo errHandle
    
    Set objCard = IDKindPaymentsType.GetCurCard
    Set objBackCard = IDKind�Ҳ�.GetCurCard
    If objCard Is Nothing Then Exit Function
    
    

    If mblnNotify = False Then
        If Not mEditType = g_Ed_�������� Then
            If MsgBox("ȷʵҪ������[" & mBalanceInfor.strNO & "]����ȡ��������?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        End If
        mblnPrintInvoice = False
        Select Case mobjRedProperty.��ӡ��ʽ
        Case 0  '����ӡ
        Case 1
            mblnPrintInvoice = True '�Զ���ӡ
        Case 2  '��ʾ��ӡ
            If MsgBox("�Ƿ��ӡ��������Ʊ��?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then mblnPrintInvoice = True
        End Select
        mblnNotify = True
    End If
    If CheckDepositFactValied = False Then Exit Function
    If CancelIsValied(objCard, tyBrushCardInfor) = False Then Exit Function
    
    
    With mBalanceInfor
        dblʣ���� = RoundEx(.dblδ���ϼ� - .dbl��Ԥ���ϼ�, 5)
        bln�˿� = dblʣ���� > 0
    End With
    
    dblCurMoney = IIf(bln�˿�, 1, -1) * Val(txtBalance(Idx_�ɿ�).Text)
    If gblnBillPrint Then
        If gobjBillPrint.zlEraseBill("", mBalanceInfor.lng����ID) = False Then Exit Function
    End If
'
    Set cllPro = New Collection
    
    If mBalanceInfor.blnSaveBill = False Then
         lng����ID = zlDatabase.GetNextId("���˽��ʼ�¼")
         With mBalanceInfor
             .lng����ID = lng����ID
             .dtBalanceDate = zlDatabase.Currentdate
        End With
        
         '���˽����¼������
         strSQL = "Zl_���˽��ʼ�¼_Cancel("
         '  No_In         ���˽��ʼ�¼.No%Type,
         strSQL = strSQL & "'" & mBalanceInfor.strNO & "',"
         '  ����id_In     ���˽��ʼ�¼.Id%Type,
         strSQL = strSQL & "" & lng����ID & ","
         '  ����Ա���_In ���˽��ʼ�¼.����Ա���%Type,
         strSQL = strSQL & "'" & UserInfo.��� & "',"
         '  ����Ա����_In ���˽��ʼ�¼.����Ա����%Type
         strSQL = strSQL & "'" & UserInfo.���� & "',"
         '  ����ʱ��_In   ���˽��ʼ�¼.�շ�ʱ��%Type := Null
         strSQL = strSQL & "to_date('" & Format(mBalanceInfor.dtBalanceDate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss')"
         strSQL = strSQL & ")"
         zlAddArray cllPro, strSQL
         'ִ��ҽ���˷Ѳ���
         If ExecuteInsureDel(cllPro) = False Then Exit Function
    End If
    
    'ִ�������ʻ�����һ��ͨ�����˷�
    If ExcuteBalanceListThreeDelSwap(cllPro) = False Then Exit Function
    
    'ִ�е�ǰ����

    If dblCurMoney <> 0 Then
       If dblCurMoney > 0 Then
            '��ǰ�˿�
            '1.ִ�е�ǰ��һ��ͨ����
            If ExecuteOneCardDelInterface(objCard, dblCurMoney, cllPro) = False Then Exit Function
            
            '2.ִ�е�ǰ�����ʻ�����
            If tyBrushCardInfor.blnת�� Then
                If ExecuteThreeSwapTransferAccount(objCard, dblCurMoney, cllPro, tyBrushCardInfor, False) = False Then Exit Function
            Else
                If mEditType = g_Ed_�������� Then
                    If ExecuteThreeSwapDelInterface(objCard, dblCurMoney, cllPro, True) = False Then Exit Function
                Else
                    If ExecuteThreeSwapDelInterface(objCard, dblCurMoney, cllPro) = False Then Exit Function
                End If
            End If
       Else
            '��ǰ�տ�
            '1.ִ�е�ǰ��һ��ͨ����
            If ExecuteOldOneCardPayInterface(mPatiInfor.lng����ID, mBalanceInfor.lng����ID, objCard, -1 * dblCurMoney, tyBrushCardInfor, cllPro) = False Then Exit Function
            '2.ִ�е�ǰ�����ʻ�����
            If ExecuteThreeSwapPayInterface(mPatiInfor.lng����ID, mBalanceInfor.lng����ID, objCard, -1 * dblCurMoney, cllPro, tyBrushCardInfor) = False Then Exit Function
       End If
    End If
    
    If objCard.�������� = 1 Then
        dblTemp = dblʣ����: dblʣ���� = 0
        mBalanceInfor.dbl�ɿ� = RoundEx(IIf(bln�˿�, -1, 1) * dblCurMoney, 5)
        mBalanceInfor.dbl�Ҳ� = RoundEx(IIf(bln�˿�, 1, -1) * Val(txtBalance(Idx_�Ҳ�).Text), 5)
        dblMoney = GetCentMoney(dblTemp)
        mBalanceInfor.dbl�ֽ� = dblMoney
    Else
        dblTemp = dblCurMoney
        If Not objBackCard Is Nothing And dblCurMoney = 0 Then
            If objBackCard.�ӿ���� <> 1 And Val(txtBalance(Idx_�Ҳ�).Text) <> 0 Then
               dblTemp = RoundEx(Val(txtBalance(Idx_�Ҳ�).Text), 6)
            End If
        End If
        dblMoney = GetCentMoney(dblTemp)
        dblʣ���� = dblʣ���� - dblCurMoney - mBalanceInfor.dbl����
    End If
    
    Call Show�����(False)
    '����˷Ѳ���
    If dblʣ���� = 0 Then
        If ExecuteOverBalanceCancel(objCard, cllPro, dblMoney) = False Then Exit Function
        mblnNotify = False
        
        strSQL = "Zl_�����Զ�����_Restore('" & mBalanceInfor.strNO & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        
        If Not gobjTax Is Nothing And gblnTax Then
            gstrTax = gobjTax.zlTaxInErase(gcnOracle, mBalanceInfor.lng����ID)
            If gstrTax <> "" Then MsgBox gstrTax, vbExclamation, gstrSysName
        End If
        
        If mblnPrintInvoice Then
            '��Ʊ��ӡ
            Call frmPrint.ReportPrint(3, mBalanceInfor.strNO, mBalanceInfor.lng����ID, mobjRedProperty, _
                mstrInvoice, mBalanceInfor.dtBalanceDate, , , mPatiInfor.lng����ID, _
                mobjRedProperty.��ӡ��ʽ, , mYBInFor.intInsure <> 0 And MCPAR.ҽ���ӿڴ�ӡƱ��)
        End If
        
        If mYBInFor.intInsure <> 0 Then
            If MCPAR.�������Ϻ��ӡ�ص� And InStr(1, mstrPrivs, ";�����˷ѻص�;") > 0 Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1137_4", Me, "����ID=" & mBalanceInfor.lng����ID, 2)
            End If
        ElseIf InStr(1, mstrPrivs, ";�����˷ѻص�;") > 0 Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1137_4", Me, "����ID=" & mBalanceInfor.lng����ID, 2)
        End If
        If mblnDepositBillPrint Then
            '��ӡԤ��Ʊ��
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103", Me, "NO=" & mBalanceInfor.strԤ��No, "����ID=" & mPatiInfor.lng����ID, "�տ�ʱ��=" & Format(mBalanceInfor.dtBalanceDate, "yyyy-mm-dd HH:MM:SS"), 2)
        End If
        Call WriteZYInforToCard(mPatiInfor.lng����ID, mBalanceInfor.lng����ID, True)
        If mintPreEditType >= 0 Then mEditType = mintPreEditType
        If mEditType = g_Ed_�������� Or mEditType = g_Ed_�������� Then
            mBalanceInfor.blnSaveBill = False
            Unload Me: ExecuteBalaceCancel = True: Exit Function
        End If
        
        mblnNotChange = True
        chkCancel.Value = 0
        Call NewBill
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        
        mblnNotChange = False
        ExecuteBalaceCancel = True
        Exit Function
    End If

    '�����˷���Ϣ
    With vsBlance
        If objCard.���ѿ� Then
            Call AddSquareBalance(objCard)
        Else
            If Not (.Rows = 2 And Trim(.TextMatrix(1, .ColIndex("���㷽ʽ"))) = "") Then
                .Rows = .Rows + 1
                .RowPosition(.Rows - 1) = 1
            End If
            '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
            strCardNo = tyBrushCardInfor.str����
            .TextMatrix(1, .ColIndex("�Ƿ�����")) = 0
            '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
            If objCard.�������� = 7 And objCard.�ӿ���� < 0 Then
                .TextMatrix(1, .ColIndex("����")) = 4
                .TextMatrix(1, .ColIndex("�༭״̬")) = 0   '0-��ֹɾ��;1-����༭���;2-����ɾ��
                .TextMatrix(1, .ColIndex("����״̬")) = 1  '�Ƿ��ѽ���:1-�ѽ���;0-δ����
            ElseIf objCard.�ӿ���� > 0 Then
                .TextMatrix(1, .ColIndex("����")) = 3
                .TextMatrix(1, .ColIndex("�����ID")) = objCard.�ӿ����
                .TextMatrix(1, .ColIndex("���������")) = objCard.����
                .TextMatrix(1, .ColIndex("�༭״̬")) = 0   '0-��ֹɾ��;1-����༭���;2-����ɾ��
                .TextMatrix(1, .ColIndex("����״̬")) = 1  '�Ƿ��ѽ���:1-�ѽ���;0-δ����
                .TextMatrix(1, .ColIndex("�Ƿ�����")) = IIf(objCard.�������Ĺ��� <> "", 1, 0)
            Else
                .TextMatrix(1, .ColIndex("����")) = 0
                .TextMatrix(1, .ColIndex("�༭״̬")) = 2   '0-��ֹɾ��;1-����༭���;2-����ɾ��
                .TextMatrix(1, .ColIndex("����״̬")) = 0  '�Ƿ��ѽ���:1-�ѽ���;0-δ����
            End If
            .TextMatrix(1, .ColIndex("���㷽ʽ")) = objCard.���㷽ʽ
            .TextMatrix(1, .ColIndex("��������")) = objCard.��������
            .TextMatrix(1, .ColIndex("�����ID")) = objCard.�ӿ����
            .TextMatrix(1, .ColIndex("���ѿ�ID")) = 0

            .TextMatrix(1, .ColIndex("������")) = Format(dblMoney, "0.00")
            .Cell(flexcpData, 1, .ColIndex("������")) = Format(dblMoney, "0.00")
            .TextMatrix(1, .ColIndex("�������")) = IIf(txtBalance(Idx_�������).Visible, Trim(txtBalance(Idx_�������).Text), "")
            .TextMatrix(1, .ColIndex("��ע")) = Trim(txtBalance(Idx_ժҪ).Text)

            If objCard.�ӿ���� > 0 Then
                .TextMatrix(1, .ColIndex("����")) = IIf(objCard.�������Ĺ��� <> "", String(Len(strCardNo), "*"), strCardNo)
                .Cell(flexcpData, 1, .ColIndex("����")) = tyBrushCardInfor.str����
                .TextMatrix(1, .ColIndex("������ˮ��")) = tyBrushCardInfor.str������ˮ��
                .TextMatrix(1, .ColIndex("����˵��")) = tyBrushCardInfor.str����˵��
                .TextMatrix(1, .ColIndex("�Ƿ�����")) = IIf(objCard.�Ƿ�����, 1, 0)
                .TextMatrix(1, .ColIndex("�Ƿ�ȫ��")) = IIf(objCard.�Ƿ�ȫ��, 1, 0)
                .TextMatrix(1, .ColIndex("���������")) = objCard.����
            End If
            mBalanceInfor.dbl�Ѹ��ϼ� = RoundEx(mBalanceInfor.dbl�Ѹ��ϼ� + dblMoney, 6)
            mBalanceInfor.dblδ���ϼ� = RoundEx(mBalanceInfor.dblδ���ϼ� - dblMoney, 6)
        End If
        For i = 1 To IDKindPaymentsType.ListCount
            'ȱʡ��λ���ֽ���
            Set objCard = IDKindPaymentsType.GetIDKindCard(i, CardTypeIndex)
            If objCard.�������� = 1 Then IDKindPaymentsType.IDKIND = i: Exit For
        Next
    End With
    
    txtBalance(Idx_�ɿ�).Text = ""
    If txtBalance(Idx_�ɿ�).Enabled And txtBalance(Idx_�ɿ�).Visible Then txtBalance(Idx_�ɿ�).SetFocus
    Call LedDisplayBank

    ExecuteBalaceCancel = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function ExecuteOverBalanceCancel(ByVal objCard As Card, _
    ByRef cllDelBalancePro As Collection, ByVal dblMoney As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ������˷Ѳ���
    '���:objCard-��ǰ֧�����
    '     cllDelBalancePro-ִ�е��˷ѵ���
    '     dblMoney-��ǰ�˿���
    '����:�˷ѳɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-23 09:31:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPro As Collection, strSQL As String
    Dim strPTBalance As String
    Dim strSquareBalance As String, bln��Ԥ�� As Boolean, i As Long
    Dim strժҪ As String
    On Error GoTo errHandle
    Set cllPro = New Collection
    For i = 1 To cllDelBalancePro.Count
        cllPro.Add cllDelBalancePro(i)
    Next
    
    If GetCancelBalance(0, strPTBalance) = False Then Exit Function
    If GetCancelBalance(2, strSquareBalance) = False Then Exit Function
    
    
    If objCard.�ӿ���� <= 0 And InStr(",1,2,", "," & objCard.�������� & ",") > 0 Then
        strPTBalance = strPTBalance & IIf(strPTBalance = "", "", "||")
        strPTBalance = strPTBalance & objCard.���㷽ʽ
        strPTBalance = strPTBalance & "|" & -1 * dblMoney
        strPTBalance = strPTBalance & "|" & IIf(txtBalance(Idx_�������).Text = "", " ", txtBalance(Idx_�������).Text)
        strPTBalance = strPTBalance & "|" & IIf(txtBalance(Idx_ժҪ).Text = "", " ", txtBalance(Idx_ժҪ).Text)
    ElseIf objCard.�ӿ���� > 0 And objCard.���ѿ� And dblMoney <> 0 Then
        For i = 1 To mcllCurSquareBalance.Count
            '�����ID|����|���ѿ�ID|���ѽ��||.
            'array(�����ID,���ѿ�ID,ˢ�����,����,����,�������,�Ƿ�����))
            strSquareBalance = strSquareBalance & IIf(strSquareBalance = "", "", "||")
            strSquareBalance = strSquareBalance & Val(mcllCurSquareBalance(i)(0))
            strSquareBalance = strSquareBalance & "|" & Trim(mcllCurSquareBalance(i)(3))
            strSquareBalance = strSquareBalance & "|" & Val(mcllCurSquareBalance(i)(1))
            strSquareBalance = strSquareBalance & "|" & -1 * Val(mcllCurSquareBalance(i)(2))
        Next
    End If
    If strSquareBalance <> "" Then
        'Zl_���˽�������_Modify
        strSQL = "Zl_���˽�������_Modify("
        '  ��������_In   Number,
        '--   1-��ͨ�˷ѷ�ʽ:
        '--     �ٽ��㷽ʽ_IN:��������,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ||.." ;Ҳ�������.
        '--   2.�������˷ѽ���:
        '--     �ٽ��㷽ʽ_IN:ֻ�ܴ���һ�����㷽ʽ,���������һЩ������Ϣ,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ"
        '--     �ڿ����ID_IN,����_IN,������ˮ��_IN,����˵��_In:��Ҫ����
        '--   3-ҽ������(�������ҽ���Ľ���,��Ҫ��ɾ��ԭҽ������,���´���ĸ���)
        '--     �ٽ��㷽ʽ_IN:��������,��ʽΪ:���㷽ʽ|������||.."
        '--   4-���ѿ�����:
        '--     �ٽ��㷽ʽ_IN:����һ��ˢ���ſ�,��ʽΪ:�����ID|����|���ѿ�ID|���ѽ��||."
        strSQL = strSQL & "" & 4 & ","
        '  ����id_In     ���˽��ʼ�¼.����id%Type,
        strSQL = strSQL & "" & mPatiInfor.lng����ID & ","
        '  ����id_In     ����Ԥ����¼.����id%Type,
        strSQL = strSQL & "" & mBalanceInfor.lng����ID & ","
        '  ���㷽ʽ_In   Varchar2,
        strSQL = strSQL & "'" & strSquareBalance & "',"
        '  �����id_In   ����Ԥ����¼.�����id%Type := Null,
        strSQL = strSQL & "NULL,"
        '  ����_In       ����Ԥ����¼.����%Type := Null,
        strSQL = strSQL & "NULL,"
        '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
        strSQL = strSQL & "NULL,"
        '  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
        strSQL = strSQL & "'" & mstrForceNote & "',"
        '  �ɿ�_In       ����Ԥ����¼.�ɿ�%Type := Null,
        strSQL = strSQL & "NULL,"
        '  �Ҳ�_In       ����Ԥ����¼.�Ҳ�%Type := Null,
        strSQL = strSQL & "NULL,"
        '  �����_In   ����Ԥ����¼.��Ԥ��%Type := Null,
        strSQL = strSQL & "NULL,"
        '  Ԥ�����_In   ����Ԥ����¼.��Ԥ��%Type := Null,
        strSQL = strSQL & "NULL,"
        '����Ա���_In    ����Ԥ����¼.����Ա���%Type := Null,
        strSQL = strSQL & "'" & UserInfo.��� & "',"
        '����Ա����_In    ����Ԥ����¼.����Ա����%Type := Null,
        strSQL = strSQL & "'" & UserInfo.���� & "',"
        '�տ�ʱ��_In      ����Ԥ����¼.����Ա����%Type := Null,
        strSQL = strSQL & "to_date('" & Format(mBalanceInfor.dtBalanceDate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
        '��Ԥ������ids_In Varchar2 := Null,
        ' ����ö��ַ���,��Ԥ��ʱ��Ч(��Ԥ�������ҵ���������һ��),��Ҫ��ʹ�ü�����Ԥ����
        strSQL = strSQL & "NULL,"
        '  �������_In Number:=0
        strSQL = strSQL & "0)"
        zlAddArray cllPro, strSQL
    End If
    
    bln��Ԥ�� = Val(txtBalance(Idx_��Ԥ��).Text) <> 0
    'Zl_���˽�������_Modify
    strSQL = "Zl_���˽�������_Modify("
    '  ��������_In   Number,
    '--   1-��ͨ�˷ѷ�ʽ:
    '--     �ٽ��㷽ʽ_IN:��������,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ||.." ;Ҳ�������.
    '--   2.�������˷ѽ���:
    '--     �ٽ��㷽ʽ_IN:ֻ�ܴ���һ�����㷽ʽ,���������һЩ������Ϣ,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ"
    '--     �ڿ����ID_IN,����_IN,������ˮ��_IN,����˵��_In:��Ҫ����
    '--   3-ҽ������(�������ҽ���Ľ���,��Ҫ��ɾ��ԭҽ������,���´���ĸ���)
    '--     �ٽ��㷽ʽ_IN:��������,��ʽΪ:���㷽ʽ|������||.."
    '--   4-���ѿ�����:
    '--     �ٽ��㷽ʽ_IN:����һ��ˢ���ſ�,��ʽΪ:�����ID|����|���ѿ�ID|���ѽ��||."
    strSQL = strSQL & "" & 1 & ","
    '  ����id_In     ���˽��ʼ�¼.����id%Type,
    strSQL = strSQL & "" & mPatiInfor.lng����ID & ","
    '  ����id_In     ����Ԥ����¼.����id%Type,
    strSQL = strSQL & "" & mBalanceInfor.lng����ID & ","
    '  ���㷽ʽ_In   Varchar2,
    strSQL = strSQL & "'" & strPTBalance & "',"
    '  �����id_In   ����Ԥ����¼.�����id%Type := Null,
    strSQL = strSQL & "NULL,"
    '  ����_In       ����Ԥ����¼.����%Type := Null,
    strSQL = strSQL & "NULL,"
    '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
    strSQL = strSQL & "NULL,"
    '  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
    strSQL = strSQL & "'" & IIf(zlCommFun.ActualLen(mstrForceNote) > 500, Mid(mstrForceNote, 1, 493) & "...", mstrForceNote) & "',"
    '  �ɿ�_In       ����Ԥ����¼.�ɿ�%Type := Null,
    strSQL = strSQL & "" & mBalanceInfor.dbl�ɿ� & ","
    '  �Ҳ�_In       ����Ԥ����¼.�Ҳ�%Type := Null,
    strSQL = strSQL & "" & mBalanceInfor.dbl�Ҳ� & ","
    '  �����_In   ����Ԥ����¼.��Ԥ��%Type := Null,
    strSQL = strSQL & "" & -1 * mBalanceInfor.dbl���� & ","
    '  Ԥ�����_In   ����Ԥ����¼.��Ԥ��%Type := Null,
    strSQL = strSQL & "" & -1 * mBalanceInfor.dbl��Ԥ���ϼ� & ","
    '����Ա���_In    ����Ԥ����¼.����Ա���%Type := Null,
    strSQL = strSQL & "'" & UserInfo.��� & "',"
    '����Ա����_In    ����Ԥ����¼.����Ա����%Type := Null,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '�տ�ʱ��_In      ����Ԥ����¼.����Ա����%Type := Null,
    strSQL = strSQL & "to_date('" & Format(mBalanceInfor.dtBalanceDate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
    '��Ԥ������ids_In Varchar2 := Null,
    ' ����ö��ַ���,��Ԥ��ʱ��Ч(��Ԥ�������ҵ���������һ��),��Ҫ��ʹ�ü�����Ԥ����
    strSQL = strSQL & "NULL,"
    '  �������_In Number:=0
    strSQL = strSQL & "2)"
    zlAddArray cllPro, strSQL
    
    If GetSaveAddDepositSQL(mPatiInfor.lng����ID, mPatiInfor.lng��ҳID, mBalanceInfor.lng����ID, cllPro) = False Then Exit Function
    
    If mEditType = g_Ed_�������� Then
        strSQL = "Zl_���˽����쳣_Update("
        strSQL = strSQL & "to_date('" & Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
        strSQL = strSQL & "" & mBalanceInfor.lng����ID & ")"
        zlAddArray cllPro, strSQL
    End If
    
    Err = 0: On Error GoTo ErrRoll:
    zlExecuteProcedureArrAy cllPro, Me.Caption
    ExecuteOverBalanceCancel = True
    Exit Function
ErrRoll:
     gcnOracle.RollbackTrans
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function


Private Function ExecuteInsureDel(ByRef cllDelBalancePro As Collection, Optional bln�쳣���� As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ��ҽ���˷Ѳ���
    '���:cllDelBalancePro-ִ�е��˷ѵ���
    '     bln�쳣����-�Ƿ��쳣����
    '����:�˷ѳɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-22 16:39:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPro As Collection, strSQL As String, strYbBalance As String
    Dim blnTransMC As Boolean, blnTrans As Boolean, i As Long
    Dim strAdvance  As String
    Dim blnReload As Boolean
    
    If mYBInFor.intInsure = 0 Then ExecuteInsureDel = True: Exit Function
    
    '��ȡҽ�����㷽ʽ
    strYbBalance = ""
    If bln�쳣���� = False Then
        If GetCancelBalance(1, strYbBalance) = False Then Exit Function
    End If
    
    On Error GoTo errHandle
    
    Set cllPro = New Collection
    For i = 1 To cllDelBalancePro.Count
        cllPro.Add cllDelBalancePro(i)
    Next
    If mYBInFor.bytMCMode = 1 Then
        If MCPAR.���ﲡ�˽������� = False Then  '��֧�������������,��ֱ�ӷ���
            If strYbBalance = "" Then ExecuteInsureDel = True: Exit Function
            MsgBox "���ڸ�ҽ����֧�������������,���,����ִ���˷Ѳ���!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    Else
        If MCPAR.סԺ�������� = False Then  '��֧�������������,��ֱ�ӷ���
            If strYbBalance = "" Then ExecuteInsureDel = True: Exit Function
            MsgBox "���ڸ�ҽ����֧��סԺ��������,���,����ִ���˷Ѳ���!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
     
    If strYbBalance <> "" Then
         'Zl_���˽�������_Modify
         strSQL = "Zl_���˽�������_Modify("
         '  ��������_In   Number,
         '--   1-��ͨ�˷ѷ�ʽ:
         '--     �ٽ��㷽ʽ_IN:��������,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ||.." ;Ҳ�������.
         '--   2.�������˷ѽ���:
         '--     �ٽ��㷽ʽ_IN:ֻ�ܴ���һ�����㷽ʽ,���������һЩ������Ϣ,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ"
         '--     �ڿ����ID_IN,����_IN,������ˮ��_IN,����˵��_In:��Ҫ����
         '--   3-ҽ������(�������ҽ���Ľ���,��Ҫ��ɾ��ԭҽ������,���´���ĸ���)
         '--     �ٽ��㷽ʽ_IN:��������,��ʽΪ:���㷽ʽ|������||.."
         '--   4-���ѿ�����:
         '--     �ٽ��㷽ʽ_IN:����һ��ˢ���ſ�,��ʽΪ:�����ID|����|���ѿ�ID|���ѽ��||."
         
        strSQL = strSQL & "" & 3 & ","
        '  ����id_In     ���˽��ʼ�¼.����id%Type,
        strSQL = strSQL & "" & mPatiInfor.lng����ID & ","
        '  ����id_In     ����Ԥ����¼.����id%Type,
        strSQL = strSQL & "" & mBalanceInfor.lng����ID & ","
        '  ���㷽ʽ_In   Varchar2,
        strSQL = strSQL & "'" & strYbBalance & "',Null,Null,Null,Null,Null,Null,Null,Null,Null,Null,Null,Null,0,1)"
        '  �����id_In   ����Ԥ����¼.�����id%Type := Null,
        '  ����_In       ����Ԥ����¼.����%Type := Null,
        '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
        '  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
        '  �ɿ�_In       ����Ԥ����¼.�ɿ�%Type := Null,
        '  �Ҳ�_In       ����Ԥ����¼.�Ҳ�%Type := Null,
        '  �����_In   ����Ԥ����¼.��Ԥ��%Type := Null,
        '  Ԥ�����_In   ����Ԥ����¼.��Ԥ��%Type := Null,
        '����Ա���_In    ����Ԥ����¼.����Ա���%Type := Null,
        '����Ա����_In    ����Ԥ����¼.����Ա����%Type := Null,
        '�տ�ʱ��_In      ����Ԥ����¼.����Ա����%Type := Null,
        '��Ԥ������ids_In Varchar2 := Null,
        ' ����ö��ַ���,��Ԥ��ʱ��Ч(��Ԥ�������ҵ���������һ��),��Ҫ��ʹ�ü�����Ԥ����
        '  �������_In Number:=0
        zlAddArray cllPro, strSQL
    End If
    
    'ִ��ҽ���˷�
    Err = 0: On Error GoTo ErrRoll:
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    blnTransMC = False
    
    If mYBInFor.bytMCMode = 1 Then
        strAdvance = mBalanceInfor.lng����ID & "|0"
        If Not gclsInsure.ClinicDelSwap(mBalanceInfor.lng����ID, , mYBInFor.intInsure, strAdvance) Then
            gcnOracle.RollbackTrans: Exit Function
        End If
        blnTransMC = True
        If zlInsureCheck(strYbBalance, strAdvance) Then
            'Zl_���˽�������_Modify
             strSQL = "Zl_���˽�������_Modify("
             '  ��������_In   Number,
             '--   1-��ͨ�˷ѷ�ʽ:
             '--     �ٽ��㷽ʽ_IN:��������,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ||.." ;Ҳ�������.
             '--   2.�������˷ѽ���:
             '--     �ٽ��㷽ʽ_IN:ֻ�ܴ���һ�����㷽ʽ,���������һЩ������Ϣ,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ"
             '--     �ڿ����ID_IN,����_IN,������ˮ��_IN,����˵��_In:��Ҫ����
             '--   3-ҽ������(�������ҽ���Ľ���,��Ҫ��ɾ��ԭҽ������,���´���ĸ���)
             '--     �ٽ��㷽ʽ_IN:��������,��ʽΪ:���㷽ʽ|������||.."
             '--   4-���ѿ�����:
             '--     �ٽ��㷽ʽ_IN:����һ��ˢ���ſ�,��ʽΪ:�����ID|����|���ѿ�ID|���ѽ��||."
             
            strSQL = strSQL & "" & 3 & ","
            '  ����id_In     ���˽��ʼ�¼.����id%Type,
            strSQL = strSQL & "" & mPatiInfor.lng����ID & ","
            '  ����id_In     ����Ԥ����¼.����id%Type,
            strSQL = strSQL & "" & mBalanceInfor.lng����ID & ","
            '  ���㷽ʽ_In   Varchar2,
            strSQL = strSQL & "'" & strAdvance & "')"
            '  �����id_In   ����Ԥ����¼.�����id%Type := Null,
            '  ����_In       ����Ԥ����¼.����%Type := Null,
            '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
            '  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
            '  �ɿ�_In       ����Ԥ����¼.�ɿ�%Type := Null,
            '  �Ҳ�_In       ����Ԥ����¼.�Ҳ�%Type := Null,
            '  �����_In   ����Ԥ����¼.��Ԥ��%Type := Null,
            '  Ԥ�����_In   ����Ԥ����¼.��Ԥ��%Type := Null,
            '����Ա���_In    ����Ԥ����¼.����Ա���%Type := Null,
            '����Ա����_In    ����Ԥ����¼.����Ա����%Type := Null,
            '�տ�ʱ��_In      ����Ԥ����¼.����Ա����%Type := Null,
            '��Ԥ������ids_In Varchar2 := Null,
            ' ����ö��ַ���,��Ԥ��ʱ��Ч(��Ԥ�������ҵ���������һ��),��Ҫ��ʹ�ü�����Ԥ����
            '  �������_In Number:=0
            zlDatabase.ExecuteProcedure strSQL, Me.Caption
            blnReload = True
        Else
            If strYbBalance <> "" Then
                'Zl_���˽�������_Modify
                 strSQL = "Zl_���˽�������_Modify("
                 '  ��������_In   Number,
                 '--   1-��ͨ�˷ѷ�ʽ:
                 '--     �ٽ��㷽ʽ_IN:��������,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ||.." ;Ҳ�������.
                 '--   2.�������˷ѽ���:
                 '--     �ٽ��㷽ʽ_IN:ֻ�ܴ���һ�����㷽ʽ,���������һЩ������Ϣ,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ"
                 '--     �ڿ����ID_IN,����_IN,������ˮ��_IN,����˵��_In:��Ҫ����
                 '--   3-ҽ������(�������ҽ���Ľ���,��Ҫ��ɾ��ԭҽ������,���´���ĸ���)
                 '--     �ٽ��㷽ʽ_IN:��������,��ʽΪ:���㷽ʽ|������||.."
                 '--   4-���ѿ�����:
                 '--     �ٽ��㷽ʽ_IN:����һ��ˢ���ſ�,��ʽΪ:�����ID|����|���ѿ�ID|���ѽ��||."
                 
                strSQL = strSQL & "" & 3 & ","
                '  ����id_In     ���˽��ʼ�¼.����id%Type,
                strSQL = strSQL & "" & mPatiInfor.lng����ID & ","
                '  ����id_In     ����Ԥ����¼.����id%Type,
                strSQL = strSQL & "" & mBalanceInfor.lng����ID & ","
                '  ���㷽ʽ_In   Varchar2,
                strSQL = strSQL & "'" & strYbBalance & "')"
                '  �����id_In   ����Ԥ����¼.�����id%Type := Null,
                '  ����_In       ����Ԥ����¼.����%Type := Null,
                '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
                '  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
                '  �ɿ�_In       ����Ԥ����¼.�ɿ�%Type := Null,
                '  �Ҳ�_In       ����Ԥ����¼.�Ҳ�%Type := Null,
                '  �����_In   ����Ԥ����¼.��Ԥ��%Type := Null,
                '  Ԥ�����_In   ����Ԥ����¼.��Ԥ��%Type := Null,
                '����Ա���_In    ����Ԥ����¼.����Ա���%Type := Null,
                '����Ա����_In    ����Ԥ����¼.����Ա����%Type := Null,
                '�տ�ʱ��_In      ����Ԥ����¼.����Ա����%Type := Null,
                '��Ԥ������ids_In Varchar2 := Null,
                ' ����ö��ַ���,��Ԥ��ʱ��Ч(��Ԥ�������ҵ���������һ��),��Ҫ��ʹ�ü�����Ԥ����
                '  �������_In Number:=0
                zlDatabase.ExecuteProcedure strSQL, Me.Caption
                blnReload = False
            End If
        End If
    Else
        strAdvance = ""
        If Not gclsInsure.SettleDelSwap(mBalanceInfor.lng����ID, mYBInFor.intInsure, strAdvance) Then
            gcnOracle.RollbackTrans:  Exit Function
        End If
        blnTransMC = True
    End If
    gcnOracle.CommitTrans: blnTrans = False
    
    If blnReload Then
        i = 1
        With vsBlance
            Do While i <= .Rows - 1
                If Val(.TextMatrix(i, .ColIndex("����"))) = 2 Then
                    Call DeletePayInfor(i, True)
                Else
                    i = i + 1
                End If
            Loop
        End With
        Call LoadBalancePayData(mPatiInfor.lng����ID, mBalanceInfor.lng����ID, , False, True, -1)
        Call LoadCurOwnerPayInfor
        MsgBox "ҽ���˿�����ѷ����仯,������µ��˿������´������ϣ�", vbInformation, gstrSysName
        mBalanceInfor.blnSaveBill = True
        Exit Function
    End If
    
    Set cllDelBalancePro = New Collection   '��ձ������Ͻ��ʵ�������
    mBalanceInfor.blnSaveBill = True
    If blnTransMC Then Call gclsInsure.BusinessAffirm(IIf(mYBInFor.bytMCMode = 1, ����Enum.Busi_ClinicDelSwap, ����Enum.Busi_SettleDelSwap), True, mYBInFor.intInsure)
    ExecuteInsureDel = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Exit Function
ErrRoll:
    If blnTrans Then gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
    
End Function

Private Function zlInsureCheck(ByVal str���ս��� As String, ByVal strAdvance As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鵱ǰ��ҽ���Ƿ���Ҫ�϶�
    '���:str���ս���-���ս���
    '       strAdvance-ҽ�����صĽ���
    '����:
    '����:��Ҫ�϶�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-20 18:03:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnMedicareCheck As Boolean, strTmp As String, i As Long, j As Long
    Dim varData As Variant, varData1 As Variant
    Dim varTemp As Variant, varTemp1 As Variant

    On Error GoTo errHandle
    If Not (strAdvance <> "" And str���ս��� <> strAdvance) Then Exit Function
    '��ʽ����ǰ��,���㷽ʽ�ͽ�����δ�����仯ʱ��У��
    blnMedicareCheck = True
    varData = Split(str���ս���, "||"): varData1 = Split(strAdvance, "||")

    If UBound(varData) = UBound(varData1) Then
        For i = 0 To UBound(varData)
            blnMedicareCheck = True
            strTmp = varData(i)
            varTemp = Split(strTmp, "|")
            For j = 0 To UBound(varData1)
                varTemp1 = Split(varData1(j), "|")
                If varTemp(0) = varTemp1(0) Then
                    If Val(varTemp(1)) = Val(varTemp1(1)) Then
                        blnMedicareCheck = False
                    End If
                End If
            Next
            If blnMedicareCheck Then Exit For
        Next
    End If
    zlInsureCheck = blnMedicareCheck
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function ExcuteBalanceListThreeDelSwap(ByRef cllDelBalancePro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ�н����б��е����������˷�
    '���:cllDelBalancePro-ִ�е��˷ѵ���
    '����:�˷ѳɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-01-22 17:20:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double
    Dim strBalance As String, objTemp As Card, i As Long
    Dim lngTypeCardTypeID As Long
    Dim strName As String
    
    On Error GoTo errHandle
  
    With vsBlance
        '�ռ��˿ʽ�����
        For i = 1 To .Rows - 1
            dblMoney = RoundEx(Val(.TextMatrix(i, .ColIndex("������"))), 6)
            strBalance = Trim(.TextMatrix(i, .ColIndex("���㷽ʽ")))
            If strBalance <> "" Then
                '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
                Select Case Val(.TextMatrix(i, .ColIndex("����")))
                Case 0 '��ͨ����
                Case 1 'Ԥ����
                Case 2 'ҽ��
                Case 3 'һ��ͨ
                    lngTypeCardTypeID = Val(.TextMatrix(i, .ColIndex("�����ID")))
                    Set objTemp = GetLocalePayCard(lngTypeCardTypeID, False)
                    If objTemp Is Nothing Then
                        strName = Trim(.TextMatrix(i, .ColIndex("���������")))
                        MsgBox "��վ�㲻֧��ʹ�á�" & IIf(strName = "", strBalance, strName) & "����ʽ�����˿�!", vbInformation + vbOKOnly, gstrSysName
                        Exit Function
                    End If
                    If Val(.TextMatrix(i, .ColIndex("����״̬"))) = 0 Then
                        If dblMoney <> 0 Then
                            'ִ���˷�
                            If Not ExecuteThreeSwapDelInterface(objTemp, dblMoney, cllDelBalancePro) Then Exit Function
                           .TextMatrix(i, .ColIndex("����״̬")) = 1
                        Else
                            
                        End If
                    End If
                Case 4 'һ��ͨ(�ϰ汾)
                    If Val(.TextMatrix(i, .ColIndex("����״̬"))) = 0 Then
                        Set objTemp = GetLocaleOldOneCard(strBalance)
                        If objTemp Is Nothing Then
                            strName = Trim(.TextMatrix(i, .ColIndex("���������")))
                            MsgBox "��վ�㲻֧��ʹ�á�" & IIf(strName = "", strBalance, strName) & "����ʽ�����˿�!", vbInformation + vbOKOnly, gstrSysName
                            Exit Function
                        End If
                        If dblMoney >= 0 Then
                            'ִ���˷���
                            If Not ExecuteOneCardDelInterface(objTemp, dblMoney, cllDelBalancePro) Then Exit Function
                            .TextMatrix(i, .ColIndex("����״̬")) = 1
                        Else
                            
                        End If
                    End If
                Case 5 '���ѿ�
                Case Else
                End Select
            End If
        Next
    End With
    ExcuteBalanceListThreeDelSwap = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function ExecuteOneCardDelInterface(ByVal objCard As Card, _
    ByVal dblDelMoney As Double, _
    ByRef cllBillPro As Collection, Optional ByVal bln�쳣���� As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:һ��ͨ�˷ѽӿ�(�ϰ�)
    '���:cllBillPro-���浥�ݵ�SQL
    '     bln�쳣����-�쳣���ϵ���(true,Ϊ�쳣���ϵ���,False-��������)
    '����:���˺�
    '����:2014-07-10 10:36:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCardNo As String, strSwap As String, strHsptCode As String 'ҽԺ����
    Dim i As Long, dblMoney As Double, strNos As String, strSQL As String
    Dim str���㷽ʽ As String
    Dim cllPro As Collection, blnTrans As Boolean
    
    '��һ��֧ͨ��,ֱ�ӷ���
    If objCard.�������� <> 7 Then ExecuteOneCardDelInterface = True: Exit Function

     mOldOneCard.rsOneCard.Filter = "���㷽ʽ='" & objCard.���㷽ʽ & "'"
    If mOldOneCard.rsOneCard.EOF Then
        MsgBox objCard.���㷽ʽ & "δ����,���ڡ������������á�����������!", vbInformation, gstrSysName
        ExecuteOneCardDelInterface = False: Exit Function
    End If
    
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    
    '�ֶ�:����,��¼����,���㷽ʽ,ժҪ,�����ID,���������,���ƿ�,���㿨���,�������,����,������ˮ��, ����˵��,�������,У�Ա�־,ҽ��,���ѿ�id
    '     �Ƿ�����,�Ƿ�ȫ��,�Ƿ�����,��Ԥ��
    '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�)
    On Error GoTo errHandle
    If mrsBalance Is Nothing Then Exit Function
    If mrsBalance.State <> 1 Then Exit Function
    mrsBalance.Filter = "����=4"
    If mrsBalance.RecordCount = 0 Then Exit Function
    With mrsBalance
        .MoveFirst
        Do While Not .EOF
            dblMoney = dblMoney + Val(NVL(mrsBalance!��Ԥ��))
            .MoveNext
        Loop
        .MoveFirst
    End With
    If RoundEx(dblMoney, 6) = 0 Then Exit Function
    
    If dblDelMoney <> dblMoney Then
        MsgBox objCard.���㷽ʽ & " ����ȫ��!" & vbCrLf & "ԭ������:" & Format(dblMoney, "0.00") & vbCrLf & " ���˿���:" & Format(dblDelMoney, "0.00"), vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    
    'һ��ͨ(��):ֻ��ʹ��һ��
    With mrsBalance
        strCardNo = NVL(!����)
        str���㷽ʽ = NVL(!���㷽ʽ)
        '���㷽ʽ|������|�������|����ժҪ||..
        str���㷽ʽ = str���㷽ʽ & "|" & -1 * dblMoney
        str���㷽ʽ = str���㷽ʽ & "|" & IIf(Trim(NVL(!�������)) = "", " ", Trim(NVL(!�������)))
        str���㷽ʽ = str���㷽ʽ & "| "
         
         
        'Zl_���˽�������_Modify
        strSQL = "Zl_���˽�������_Modify("
        '  ��������_In   Number,
        '--   1-��ͨ�˷ѷ�ʽ:
        '--     �ٽ��㷽ʽ_IN:��������,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ||.." ;Ҳ�������.
        '--   2.�������˷ѽ���:
        '--     �ٽ��㷽ʽ_IN:ֻ�ܴ���һ�����㷽ʽ,���������һЩ������Ϣ,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ"
        '--     �ڿ����ID_IN,����_IN,������ˮ��_IN,����˵��_In:��Ҫ����
        '--   3-ҽ������(�������ҽ���Ľ���,��Ҫ��ɾ��ԭҽ������,���´���ĸ���)
        '--     �ٽ��㷽ʽ_IN:��������,��ʽΪ:���㷽ʽ|������||.."
        '--   4-���ѿ�����:
        '--     �ٽ��㷽ʽ_IN:����һ��ˢ���ſ�,��ʽΪ:�����ID|����|���ѿ�ID|���ѽ��||."
        strSQL = strSQL & "" & 2 & ","
        '  ����id_In     ���˽��ʼ�¼.����id%Type,
        strSQL = strSQL & "" & mPatiInfor.lng����ID & ","
        '  ����id_In     ����Ԥ����¼.����id%Type,
        strSQL = strSQL & "" & mBalanceInfor.lng����ID & ","
        '  ���㷽ʽ_In   Varchar2,
        strSQL = strSQL & "'" & str���㷽ʽ & "',"
        '  �����id_In   ����Ԥ����¼.�����id%Type := Null,
        strSQL = strSQL & "NULL,"
        '  ����_In       ����Ԥ����¼.����%Type := Null,
        strSQL = strSQL & "'" & strCardNo & "',"
        '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
        strSQL = strSQL & "'" & NVL(!������ˮ��) & "',"
        '  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
        strSQL = strSQL & "'" & NVL(!����˵��) & "')"
        '  �ɿ�_In       ����Ԥ����¼.�ɿ�%Type := Null,
        '  �Ҳ�_In       ����Ԥ����¼.�Ҳ�%Type := Null,
        '  �����_In   ����Ԥ����¼.��Ԥ��%Type := Null,
        '  Ԥ�����_In   ����Ԥ����¼.��Ԥ��%Type := Null,
        '  �������_In Number:=0
       If Not bln�쳣���� Then zlAddArray cllPro, strSQL
    End With
    
    On Error GoTo ErrRoll:
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    Err = 0: On Error GoTo ErrRoll:
    If Not mobjICCard.ReturnSwap(strCardNo, strHsptCode, strSwap, dblMoney) Then
        gcnOracle.RollbackTrans
        MsgBox "һ��ͨ�˷ѽ��׵���ʧ��,���ܼ����˷Ѳ�����", vbExclamation, gstrSysName
        Exit Function
    End If
    gcnOracle.CommitTrans
    Set cllBillPro = New Collection
    ExecuteOneCardDelInterface = True
    mBalanceInfor.blnSaveBill = True

    Exit Function
ErrRoll:
    gcnOracle.RollbackTrans
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ExecuteThreeSwapDelBatch(ByVal objCard As Card, _
    ByVal dblMoney As Double, ByVal strInput As String, ByRef cllBillPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:һ��֧ͨ��(��������ӿ�)
    '���:dblMoney-���ν�����
    '     cllBillPro-���ݹ���(ִ��������,�Ա�����´νӿ�ʱ�ظ�ִ��)
    '����:ִ�гɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-07-09 18:14:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSwapNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim cllPro As Collection, blnTrans As Boolean
    Dim str����IDs As String, i As Long, strValue As String
    Dim cllUpdate As Collection, cllThreeSwap As Collection
    Dim strSQL As String, rsTemp As ADODB.Recordset, dblZFJE As Double
    Dim strCardNo As String, str���㷽ʽ   As String
    Dim strOutXML As String, strInXML As String, strExpend As String
    Dim objXml As New clsXML, strArray() As String, lngRow As Long
    Dim strExpendAfterXml As String, strBalanceIDs As String
    Dim j As Long
    
    Set cllUpdate = New Collection
    Set cllThreeSwap = New Collection
    
    If objCard Is Nothing Then
        If InStr(";" & mstrPrivsCard & ";", ";�����ӿ�����;") = 0 Then
            MsgBox "��û�������ӿ�����Ȩ�ޣ��޷����ýӿڲ�����", vbInformation, gstrSysName
            Exit Function
        Else
            MsgBox "δ�ҵ��˿�ӿ�,����ӿڲ�����", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '��һ��֧ͨ��,ֱ�ӷ���
    If objCard.�ӿ���� <= 0 Or objCard.���ѿ� Then ExecuteThreeSwapDelBatch = True: Exit Function
    
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    
    Err = 0: On Error GoTo ErrHand:
    
    strArray = Split(mstrBalanceLimit, "|")
    
    For i = 0 To UBound(strArray)
        If Val(Split(strArray(i), ",")(0)) = objCard.�ӿ���� Then
            If dblMoney > Abs(Val(Split(strArray(i), ",")(1))) Then
                MsgBox objCard.���㷽ʽ & " ���˿����������˿���!" & vbCrLf & "����˿���:" & Format(Val(Split(strArray(i), ",")(1)), "0.00") & vbCrLf & " ���˿���:" & Format(dblMoney, "0.00"), vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
        End If
    Next i
    
    With mrsBalance
        str���㷽ʽ = objCard.���㷽ʽ
        
        '���㷽ʽ|������|�������|����ժҪ||..
        str���㷽ʽ = str���㷽ʽ & "|" & -1 * dblMoney
        str���㷽ʽ = str���㷽ʽ & "|" & ""
        str���㷽ʽ = str���㷽ʽ & "|" & ""
        
        '����֮ǰ,�ȴ�������
        'Zl_���˽��ʽ���_Modify
        strSQL = "Zl_���˽��ʽ���_Modify("
        '  ��������_In     Number,
        '  --��������_In:
        '  --   0-��ͨ�շѷ�ʽ:
        '  --   1.����������:
        '  --     �ٽ��㷽ʽ_IN:ֻ�ܴ���һ�����㷽ʽ,���������һЩ������Ϣ,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ"
        '  --     ����֧Ʊ��_In:������
        '  --     �ۿ����ID_IN,����_IN,������ˮ��_IN,����˵��_In:��Ҫ����
        strSQL = strSQL & "1,"
        '  ����id_In       ���˽��ʼ�¼.����id%Type,
        strSQL = strSQL & "" & mPatiInfor.lng����ID & ","
        '  ����id_In       ����Ԥ����¼.����id%Type,
        strSQL = strSQL & "" & mBalanceInfor.lng����ID & ","
        '  ���㷽ʽ_In     Varchar2,
        strSQL = strSQL & "'" & str���㷽ʽ & "',"
        '  ��Ԥ��_In       ����Ԥ����¼.��Ԥ��%Type := Null,
        strSQL = strSQL & "NULL,"
        '  ��֧Ʊ��_In     ����Ԥ����¼.��Ԥ��%Type := Null,
        strSQL = strSQL & "NULL,"
        '  �����id_In     ����Ԥ����¼.�����id%Type := Null,
        strSQL = strSQL & "" & objCard.�ӿ���� & ","
        '  ����_In         ����Ԥ����¼.����%Type := Null,
        strSQL = strSQL & "" & "Null" & ","
        '  ������ˮ��_In   ����Ԥ����¼.������ˮ��%Type := Null,
        strSQL = strSQL & "" & "Null" & ","
        '  ����˵��_In     ����Ԥ����¼.����˵��%Type := Null,
        strSQL = strSQL & "" & "Null" & ","
        '  �ɿ�_In         ����Ԥ����¼.�ɿ�%Type := Null,
        strSQL = strSQL & "" & "Null" & ","
        '  �Ҳ�_In         ����Ԥ����¼.�Ҳ�%Type := Null,
        strSQL = strSQL & "" & "Null" & ","
        '  �����_In     ������ü�¼.ʵ�ս��%Type := Null,
        strSQL = strSQL & "NULL,"
        '  ��������_In     Number := 2,
        strSQL = strSQL & "" & IIf(mEditType = g_Ed_������� Or mblnCurMzBalanceNo, 1, 2) & ","
        '  ȱʡ���㷽ʽ_In ���㷽ʽ.����%Type := Null,
        strSQL = strSQL & "NULL,"
        '    ����Ա���_In    ����Ԥ����¼.����Ա���%Type := Null,
        strSQL = strSQL & "'" & UserInfo.��� & "',"
        '    ����Ա����_In    ����Ԥ����¼.����Ա����%Type := Null,
        strSQL = strSQL & "'" & UserInfo.���� & "',"
        '    �տ�ʱ��_In      ����Ԥ����¼.����Ա����%Type := Null,
        strSQL = strSQL & "to_date('" & Format(mBalanceInfor.dtBalanceDate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
        '    ��Ԥ������ids_In Varchar2 := Null,
        strSQL = strSQL & "NULL,"
        '  ��ɽ���_In Number:=0
        strSQL = strSQL & "0)"
        zlAddArray cllPro, strSQL
    End With
    
    On Error GoTo ErrRoll:
    
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    
    objXml.ClearXmlText
    Call objXml.AppendNode("JSLIST")
    strArray = Split(strInput, "|")
    For i = 0 To UBound(strArray)
        Call objXml.AppendNode("JS")
            Call objXml.appendData("KH", Split(strArray(i), ",")(0))
            Call objXml.appendData("JYLSH", TruncStringEx(Split(strArray(i), ",")(1), True))
            Call objXml.appendData("JYSM", TruncStringEx(Split(strArray(i), ",")(2), True))
            Call objXml.appendData("ZFJE", RoundEx(-1 * Val(Split(strArray(i), ",")(3)), 2))
            Call objXml.appendData("JSLX", 1)
            Call objXml.appendData("ID", Split(strArray(i), ",")(4))
        Call objXml.AppendNode("JS", True)
        
        strSQL = "Zl_�����˿���Ϣ_Insert("
        strSQL = strSQL & mBalanceInfor.lng����ID & ","
        strSQL = strSQL & Val(Split(strArray(i), ",")(4)) & ","
        strSQL = strSQL & -1 * Val(Split(strArray(i), ",")(3)) & ",'"
        strSQL = strSQL & TruncStringEx(Split(strArray(i), ",")(0), True) & "','"
        strSQL = strSQL & TruncStringEx(Split(strArray(i), ",")(1), True) & "','"
        strSQL = strSQL & TruncStringEx(Split(strArray(i), ",")(2), True) & "')"
        zlAddArray cllThreeSwap, strSQL
        strBalanceIDs = strBalanceIDs & "," & Val(Split(strArray(i), ",")(4))
    Next i
    Call objXml.AppendNode("JSLIST", True)

    strInXML = objXml.XmlText
    strExpend = objXml.XmlText
    If strBalanceIDs <> "" Then strBalanceIDs = "1|" & Mid(strBalanceIDs, 2)
    
    If gobjSquare.objSquareCard.zlReturnCheck(Me, mlngModul, objCard.�ӿ����, objCard.���ѿ�, "", strBalanceIDs, _
         dblMoney, "", "", strExpend) = False Then gcnOracle.RollbackTrans: Exit Function
    
    If gobjSquare.objSquareCard.zlReturnMultiMoney(Me, mlngModul, objCard.�ӿ����, objCard.���ѿ�, strInXML, _
         mBalanceInfor.lng����ID, strOutXML, strExpend) = False Then gcnOracle.RollbackTrans: Exit Function
    
    
    
    If strOutXML <> "" Then
        If zlXML_Init = False Then Exit Function
        If zlXML_LoadXMLToDOMDocument(strOutXML, False) = False Then Exit Function
        Call zlXML_GetChildRows("JSLIST", "JS", lngRow)
        For i = 0 To lngRow - 1
            Call zlXML_GetNodeValue("ID", i, strValue)
            strSQL = "Zl_�����˿���Ϣ_Insert("
            strSQL = strSQL & mBalanceInfor.lng����ID & ","
            strSQL = strSQL & Val(strValue) & ","
            For j = 0 To UBound(strArray)
                If Val(Split(strArray(i), ",")(4)) = Val(strValue) Then
                    dblZFJE = -1 * Val(Split(strArray(i), ",")(3))
                    Exit For
                End If
            Next j
            strSQL = strSQL & dblZFJE & ",'"
            Call zlXML_GetNodeValue("KH", i, strValue)
            strSQL = strSQL & strValue & "','"
            Call zlXML_GetNodeValue("TKLSH", i, strValue)
            strSQL = strSQL & strValue & "','"
            Call zlXML_GetNodeValue("TKSM", i, strValue)
            strSQL = strSQL & strValue & "',"
            strSQL = strSQL & 1 & ")"
            zlAddArray cllThreeSwap, strSQL
        Next i
    End If
    
    If strExpend <> "" Then
        If zlXML_Init = False Then Exit Function
        If zlXML_LoadXMLToDOMDocument(strExpend, False) = False Then Exit Function
        Call zlXML_GetChildRows("EXPENDS", "EXPEND", lngRow)
        For i = 0 To lngRow - 1
            Call zlXML_GetNodeValue("XMMC", i, strValue)
            strExpendAfterXml = strExpendAfterXml & "||" & strValue
            Call zlXML_GetNodeValue("XMNR", i, strValue)
            strExpendAfterXml = strExpendAfterXml & "|" & strValue
        Next i
    End If
    If strExpendAfterXml <> "" Then strExpendAfterXml = Mid(strExpendAfterXml, 3)
    Call zlAddThreeSwapSQLToCollection(False, mBalanceInfor.lng����ID, objCard.�ӿ����, objCard.���ѿ�, "", strExpendAfterXml, cllThreeSwap)
    zlExecuteProcedureArrAy cllUpdate, Me.Caption, , True
    ExecuteThreeSwapDelBatch = True
    
    Set cllBillPro = New Collection
    zlExecuteProcedureArrAy cllThreeSwap, Me.Caption
    
    Exit Function
ErrRoll:
    gcnOracle.RollbackTrans
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ExecuteThreeSwapDelSingle(ByVal objCard As Card, _
    ByVal dblMoney As Double, ByVal str���� As String, ByVal str����˵�� As String, _
    ByVal str������ˮ�� As String, ByVal lngԤ��ID As Long, ByRef cllBillPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:һ��֧ͨ��(���������ӿ�)
    '���:dblMoney-���ν�����
    '     cllBillPro-���ݹ���(ִ��������,�Ա�����´νӿ�ʱ�ظ�ִ��)
    '����:ִ�гɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-07-09 18:14:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSwapNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim cllPro As Collection, blnTrans As Boolean
    Dim str����IDs As String, i As Long, strValue As String
    Dim cllUpdate As Collection, cllThreeSwap As Collection
    Dim strSQL As String, rsTemp As ADODB.Recordset, dblZFJE As Double
    Dim strCardNo As String, str���㷽ʽ   As String
    Dim strOutXML As String, strInXML As String, strExpend As String
    Dim objXml As New clsXML, strArray() As String, lngRow As Long
    Dim strExpendAfterXml As String, strBalanceIDs As String
    Dim j As Long
    
    Set cllUpdate = New Collection
    Set cllThreeSwap = New Collection
    
    If objCard Is Nothing Then
        If InStr(";" & mstrPrivsCard & ";", ";�����ӿ�����;") = 0 Then
            MsgBox "��û�������ӿ�����Ȩ�ޣ��޷����ýӿڲ�����", vbInformation, gstrSysName
            Exit Function
        Else
            MsgBox "δ�ҵ��˿�ӿ�,����ӿڲ�����", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '��һ��֧ͨ��,ֱ�ӷ���
    If objCard.�ӿ���� <= 0 Or objCard.���ѿ� Then ExecuteThreeSwapDelSingle = True: Exit Function
    
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    
    Err = 0: On Error GoTo ErrHand:
    
    With mrsBalance
        str���㷽ʽ = objCard.���㷽ʽ
        
        '���㷽ʽ|������|�������|����ժҪ||..
        str���㷽ʽ = str���㷽ʽ & "|" & -1 * dblMoney
        str���㷽ʽ = str���㷽ʽ & "|" & ""
        str���㷽ʽ = str���㷽ʽ & "|" & ""
        
        '����֮ǰ,�ȴ�������
        'Zl_���˽��ʽ���_Modify
        strSQL = "Zl_���˽��ʽ���_Modify("
        '  ��������_In     Number,
        '  --��������_In:
        '  --   0-��ͨ�շѷ�ʽ:
        '  --   1.����������:
        '  --     �ٽ��㷽ʽ_IN:ֻ�ܴ���һ�����㷽ʽ,���������һЩ������Ϣ,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ"
        '  --     ����֧Ʊ��_In:������
        '  --     �ۿ����ID_IN,����_IN,������ˮ��_IN,����˵��_In:��Ҫ����
        strSQL = strSQL & "1,"
        '  ����id_In       ���˽��ʼ�¼.����id%Type,
        strSQL = strSQL & "" & mPatiInfor.lng����ID & ","
        '  ����id_In       ����Ԥ����¼.����id%Type,
        strSQL = strSQL & "" & mBalanceInfor.lng����ID & ","
        '  ���㷽ʽ_In     Varchar2,
        strSQL = strSQL & "'" & str���㷽ʽ & "',"
        '  ��Ԥ��_In       ����Ԥ����¼.��Ԥ��%Type := Null,
        strSQL = strSQL & "NULL,"
        '  ��֧Ʊ��_In     ����Ԥ����¼.��Ԥ��%Type := Null,
        strSQL = strSQL & "NULL,"
        '  �����id_In     ����Ԥ����¼.�����id%Type := Null,
        strSQL = strSQL & "" & objCard.�ӿ���� & ","
        '  ����_In         ����Ԥ����¼.����%Type := Null,
        strSQL = strSQL & "" & "Null" & ","
        '  ������ˮ��_In   ����Ԥ����¼.������ˮ��%Type := Null,
        strSQL = strSQL & "" & "Null" & ","
        '  ����˵��_In     ����Ԥ����¼.����˵��%Type := Null,
        strSQL = strSQL & "" & "Null" & ","
        '  �ɿ�_In         ����Ԥ����¼.�ɿ�%Type := Null,
        strSQL = strSQL & "" & "Null" & ","
        '  �Ҳ�_In         ����Ԥ����¼.�Ҳ�%Type := Null,
        strSQL = strSQL & "" & "Null" & ","
        '  �����_In     ������ü�¼.ʵ�ս��%Type := Null,
        strSQL = strSQL & "NULL,"
        '  ��������_In     Number := 2,
        strSQL = strSQL & "" & IIf(mEditType = g_Ed_������� Or mblnCurMzBalanceNo, 1, 2) & ","
        '  ȱʡ���㷽ʽ_In ���㷽ʽ.����%Type := Null,
        strSQL = strSQL & "NULL,"
        '    ����Ա���_In    ����Ԥ����¼.����Ա���%Type := Null,
        strSQL = strSQL & "'" & UserInfo.��� & "',"
        '    ����Ա����_In    ����Ԥ����¼.����Ա����%Type := Null,
        strSQL = strSQL & "'" & UserInfo.���� & "',"
        '    �տ�ʱ��_In      ����Ԥ����¼.����Ա����%Type := Null,
        strSQL = strSQL & "to_date('" & Format(mBalanceInfor.dtBalanceDate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
        '    ��Ԥ������ids_In Varchar2 := Null,
        strSQL = strSQL & "NULL,"
        '  ��ɽ���_In Number:=0
        strSQL = strSQL & "0)"
        zlAddArray cllPro, strSQL
    End With
    
    On Error GoTo ErrRoll:
    
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    
    strBalanceIDs = "1|" & lngԤ��ID
    If gobjSquare.objSquareCard.zlReturnCheck(Me, mlngModul, objCard.�ӿ����, False, str����, strBalanceIDs, _
         dblMoney, str������ˮ��, str����˵��, strExpend) = False Then gcnOracle.RollbackTrans: Exit Function
    
    If gobjSquare.objSquareCard.zlReturnMoney(Me, mlngModul, objCard.�ӿ����, False, str����, strBalanceIDs, _
         dblMoney, str������ˮ��, str����˵��, strExpend) = False Then gcnOracle.RollbackTrans: Exit Function
         
    strSQL = "Zl_�����˿���Ϣ_Insert("
    strSQL = strSQL & mBalanceInfor.lng����ID & ","
    strSQL = strSQL & lngԤ��ID & ","
    strSQL = strSQL & dblMoney & ",'"
    strSQL = strSQL & str���� & "','"
    strSQL = strSQL & str������ˮ�� & "','"
    strSQL = strSQL & str����˵�� & "',"
    strSQL = strSQL & 0 & ")"
    zlAddArray cllThreeSwap, strSQL
    
    Call zlAddThreeSwapSQLToCollection(False, mBalanceInfor.lng����ID, objCard.�ӿ����, objCard.���ѿ�, "", strExpend, cllThreeSwap, lngԤ��ID)
    zlExecuteProcedureArrAy cllUpdate, Me.Caption, , True
    Set cllBillPro = New Collection
    zlExecuteProcedureArrAy cllThreeSwap, Me.Caption
    ExecuteThreeSwapDelSingle = True
    Exit Function
ErrRoll:
    gcnOracle.RollbackTrans
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ExecuteThreeSwapDelInterface(ByVal objCard As Card, _
    ByVal dblDelMoney As Double, ByRef cllBillPro As Collection, _
    Optional ByVal bln�쳣���� As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:һ��֧ͨ��(�����ӿ�)
    '���:dblMoney-���ν�����
    '     cllBillPro-���ݹ���(ִ��������,�Ա�����´νӿ�ʱ�ظ�ִ��)
    '     bln�쳣����-�쳣����ʱ����:true-�쳣����;false-�������ϲ���
    '����:ִ�гɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-07-09 18:14:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSwapNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim cllPro As Collection, blnTrans As Boolean
    Dim str����IDs As String, i As Long
    Dim cllUpdate As Collection, cllThreeSwap As Collection
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strCardNo As String, dblMoney As Double, str���㷽ʽ  As String
    
    
    Set cllUpdate = New Collection
    Set cllThreeSwap = New Collection
    
    '��һ��֧ͨ��,ֱ�ӷ���
    If objCard.�ӿ���� <= 0 Or objCard.���ѿ� Then ExecuteThreeSwapDelInterface = True: Exit Function
    
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    
    Err = 0: On Error GoTo ErrHand:
    '�ֶ�:����,��¼����,���㷽ʽ,ժҪ,�����ID,���������,���ƿ�,���㿨���,�������,����,������ˮ��, ����˵��,�������,У�Ա�־,ҽ��,���ѿ�id
    '     �Ƿ�����,�Ƿ�ȫ��,�Ƿ�����,��Ԥ��
    '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�)
    
    If bln�쳣���� = True Then
        If Not mrsOldBalance Is Nothing Then
            If mrsOldBalance.State <> 1 Then Exit Function
            
            mrsOldBalance.Filter = "����=3 And �����ID=" & objCard.�ӿ����
            If mrsOldBalance.RecordCount = 0 Then Exit Function
        
            With mrsOldBalance
                .MoveFirst
                Do While Not .EOF
                    dblMoney = dblMoney + Val(NVL(mrsOldBalance!��Ԥ��))
                    .MoveNext
                Loop
                .MoveFirst
            End With
            
            If RoundEx(dblMoney, 6) = 0 Then Exit Function
            If dblDelMoney > dblMoney Then
                MsgBox objCard.���㷽ʽ & " ���˿������ԭʼ������!" & vbCrLf & "ԭ������:" & Format(dblMoney, "0.00") & vbCrLf & " ���˿���:" & Format(dblDelMoney, "0.00"), vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
            
            With mrsOldBalance
                strCardNo = NVL(!����)
                strSwapNO = NVL(!������ˮ��)
                strSwapMemo = NVL(!����˵��)
                str���㷽ʽ = NVL(!���㷽ʽ)
                
                '���㷽ʽ|������|�������|����ժҪ||..
                str���㷽ʽ = str���㷽ʽ & "|" & -1 * dblDelMoney
                str���㷽ʽ = str���㷽ʽ & "|" & IIf(Trim(NVL(!�������)) = "", " ", Trim(NVL(!�������)))
                str���㷽ʽ = str���㷽ʽ & "|" & IIf(Trim(NVL(!ժҪ)) = "", " ", Trim(NVL(!ժҪ)))
                'Zl_���˽�������_Modify
                strSQL = "Zl_���˽�������_Modify("
                '  ��������_In   Number,
                '--   1-��ͨ�˷ѷ�ʽ:
                '--     �ٽ��㷽ʽ_IN:��������,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ||.." ;Ҳ�������.
                '--   2.�������˷ѽ���:
                '--     �ٽ��㷽ʽ_IN:ֻ�ܴ���һ�����㷽ʽ,���������һЩ������Ϣ,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ"
                '--     �ڿ����ID_IN,����_IN,������ˮ��_IN,����˵��_In:��Ҫ����
                '--   3-ҽ������(�������ҽ���Ľ���,��Ҫ��ɾ��ԭҽ������,���´���ĸ���)
                '--     �ٽ��㷽ʽ_IN:��������,��ʽΪ:���㷽ʽ|������||.."
                '--   4-���ѿ�����:
                '--     �ٽ��㷽ʽ_IN:����һ��ˢ���ſ�,��ʽΪ:�����ID|����|���ѿ�ID|���ѽ��||."
                strSQL = strSQL & "" & 2 & ","
                '  ����id_In     ���˽��ʼ�¼.����id%Type,
                strSQL = strSQL & "" & mPatiInfor.lng����ID & ","
                '  ����id_In     ����Ԥ����¼.����id%Type,
                strSQL = strSQL & "" & mBalanceInfor.lng����ID & ","
                '  ���㷽ʽ_In   Varchar2,
                strSQL = strSQL & "'" & str���㷽ʽ & "',"
                '  �����id_In   ����Ԥ����¼.�����id%Type := Null,
                strSQL = strSQL & "" & objCard.�ӿ���� & ","
                '  ����_In       ����Ԥ����¼.����%Type := Null,
                strSQL = strSQL & "'" & strCardNo & "',"
                '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
                strSQL = strSQL & "'" & NVL(!������ˮ��) & "',"
                '  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
                strSQL = strSQL & "'" & NVL(!����˵��) & "')"
                '  �ɿ�_In       ����Ԥ����¼.�ɿ�%Type := Null,
                '  �Ҳ�_In       ����Ԥ����¼.�Ҳ�%Type := Null,
                '  �����_In   ����Ԥ����¼.��Ԥ��%Type := Null,
                '  Ԥ�����_In   ����Ԥ����¼.��Ԥ��%Type := Null,
                '  �������_In Number:=0
                zlAddArray cllPro, strSQL
            End With
        End If
    Else
        If Not mrsBalance Is Nothing Then
            If mrsBalance.State <> 1 Then Exit Function
            
            mrsBalance.Filter = "����=3 And �����ID=" & objCard.�ӿ����
            If mrsBalance.RecordCount = 0 Then Exit Function
        
            With mrsBalance
                .MoveFirst
                Do While Not .EOF
                    dblMoney = dblMoney + Val(NVL(mrsBalance!��Ԥ��))
                    .MoveNext
                Loop
                .MoveFirst
            End With
            
            If RoundEx(dblMoney, 6) = 0 Then Exit Function
            If dblDelMoney > dblMoney Then
                MsgBox objCard.���㷽ʽ & " ���˿������ԭʼ������!" & vbCrLf & "ԭ������:" & Format(dblMoney, "0.00") & vbCrLf & " ���˿���:" & Format(dblDelMoney, "0.00"), vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
            
            With mrsBalance
                strCardNo = NVL(!����)
                strSwapNO = NVL(!������ˮ��)
                strSwapMemo = NVL(!����˵��)
                str���㷽ʽ = NVL(!���㷽ʽ)
                
                '���㷽ʽ|������|�������|����ժҪ||..
                str���㷽ʽ = str���㷽ʽ & "|" & -1 * dblDelMoney
                str���㷽ʽ = str���㷽ʽ & "|" & IIf(Trim(NVL(!�������)) = "", " ", Trim(NVL(!�������)))
                str���㷽ʽ = str���㷽ʽ & "|" & IIf(Trim(NVL(!ժҪ)) = "", " ", Trim(NVL(!ժҪ)))
                'Zl_���˽�������_Modify
                strSQL = "Zl_���˽�������_Modify("
                '  ��������_In   Number,
                '--   1-��ͨ�˷ѷ�ʽ:
                '--     �ٽ��㷽ʽ_IN:��������,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ||.." ;Ҳ�������.
                '--   2.�������˷ѽ���:
                '--     �ٽ��㷽ʽ_IN:ֻ�ܴ���һ�����㷽ʽ,���������һЩ������Ϣ,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ"
                '--     �ڿ����ID_IN,����_IN,������ˮ��_IN,����˵��_In:��Ҫ����
                '--   3-ҽ������(�������ҽ���Ľ���,��Ҫ��ɾ��ԭҽ������,���´���ĸ���)
                '--     �ٽ��㷽ʽ_IN:��������,��ʽΪ:���㷽ʽ|������||.."
                '--   4-���ѿ�����:
                '--     �ٽ��㷽ʽ_IN:����һ��ˢ���ſ�,��ʽΪ:�����ID|����|���ѿ�ID|���ѽ��||."
                strSQL = strSQL & "" & 2 & ","
                '  ����id_In     ���˽��ʼ�¼.����id%Type,
                strSQL = strSQL & "" & mPatiInfor.lng����ID & ","
                '  ����id_In     ����Ԥ����¼.����id%Type,
                strSQL = strSQL & "" & mBalanceInfor.lng����ID & ","
                '  ���㷽ʽ_In   Varchar2,
                strSQL = strSQL & "'" & str���㷽ʽ & "',"
                '  �����id_In   ����Ԥ����¼.�����id%Type := Null,
                strSQL = strSQL & "" & objCard.�ӿ���� & ","
                '  ����_In       ����Ԥ����¼.����%Type := Null,
                strSQL = strSQL & "'" & strCardNo & "',"
                '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
                strSQL = strSQL & "'" & NVL(!������ˮ��) & "',"
                '  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
                strSQL = strSQL & "'" & NVL(!����˵��) & "')"
                '  �ɿ�_In       ����Ԥ����¼.�ɿ�%Type := Null,
                '  �Ҳ�_In       ����Ԥ����¼.�Ҳ�%Type := Null,
                '  �����_In   ����Ԥ����¼.��Ԥ��%Type := Null,
                '  Ԥ�����_In   ����Ԥ����¼.��Ԥ��%Type := Null,
                '  �������_In Number:=0
                zlAddArray cllPro, strSQL
            End With
        End If
    End If
    
    On Error GoTo ErrRoll:
    
    str����IDs = mBalanceInfor.lng����ID & IIf(mBalanceInfor.lng����ID <> 0, "," & mBalanceInfor.lng����ID, "")
    
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    
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
    '       strSwapExtendInfor-���׵���չ��Ϣ
    '           ��ʽΪ:��Ŀ����1|��Ŀ����2||��||��Ŀ����n|��Ŀ����n ÿ����Ŀ�в��ܰ���|�ַ�
    If gobjSquare.objSquareCard.zlReturnMoney(Me, mlngModul, objCard.�ӿ����, objCard.���ѿ�, strCardNo, _
        "2|" & mBalanceInfor.lng����ID, dblDelMoney, strSwapNO, strSwapMemo, strSwapExtendInfor) = False Then gcnOracle.RollbackTrans: Exit Function
    'Call zlAddUpdateSwapSQL(False, str����IDs, objCard.�ӿ����, objCard.���ѿ�, strCardNO, strSwapNO, strSwapMemo, cllUpdate, 2)
    Call zlAddThreeSwapSQLToCollection(False, mBalanceInfor.lng����ID, objCard.�ӿ����, objCard.���ѿ�, strCardNo, strSwapExtendInfor, cllThreeSwap)
    
    zlExecuteProcedureArrAy cllUpdate, Me.Caption, , True
    Set cllBillPro = New Collection
    zlExecuteProcedureArrAy cllThreeSwap, Me.Caption
    mBalanceInfor.blnSaveBill = True
    ExecuteThreeSwapDelInterface = True
    Exit Function
ErrRoll:
    gcnOracle.RollbackTrans
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ExecuteThreeSwapTransferPay(ByVal objCard As Card, _
    ByVal dblMoney As Double, ByRef cllBillPro As Collection, _
    ByRef tyBrushCard As TY_BrushCard) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:һ��ͨת��֧��(�����ӿ�)
    '���:dblMoney-���ν�����
    '     cllBillPro-���ݹ���(ִ��������,�Ա�����´νӿ�ʱ�ظ�ִ��)
    '     tyBrushCard-ת��ˢ����Ϣ
    '����:ִ�гɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-07-09 18:14:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSwapNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim cllPro As Collection, blnTrans As Boolean
    Dim str����IDs As String, i As Long
    Dim cllUpdate As Collection, cllThreeSwap As Collection
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strCardNo As String, str���㷽ʽ   As String
    Dim strXMLExpend As String
    
    Set cllUpdate = New Collection
    Set cllThreeSwap = New Collection
    
    '��һ��֧ͨ��,ֱ�ӷ���
    If objCard.�ӿ���� <= 0 Or objCard.���ѿ� Then ExecuteThreeSwapTransferPay = True: Exit Function
    If Not objCard.�Ƿ�ת�ʼ����� Then Exit Function
    
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    
    Err = 0: On Error GoTo ErrHand:
  
    With mrsBalance
        strCardNo = tyBrushCard.str����
        strSwapNO = tyBrushCard.str������ˮ��
        strSwapMemo = tyBrushCard.str����˵��
        str���㷽ʽ = objCard.���㷽ʽ
        
        '���㷽ʽ|������|�������|����ժҪ||..
        str���㷽ʽ = str���㷽ʽ & "|" & dblMoney
        str���㷽ʽ = str���㷽ʽ & "|" & " "
        str���㷽ʽ = str���㷽ʽ & "|" & "ת�ʽ���"
        
        '����֮ǰ,�ȴ�������
        'Zl_���˽��ʽ���_Modify
        strSQL = "Zl_���˽��ʽ���_Modify("
        '  ��������_In     Number,
        '  --��������_In:
        '  --   0-��ͨ�շѷ�ʽ:
        '  --   1.����������:
        '  --     �ٽ��㷽ʽ_IN:ֻ�ܴ���һ�����㷽ʽ,���������һЩ������Ϣ,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ"
        '  --     ����֧Ʊ��_In:������
        '  --     �ۿ����ID_IN,����_IN,������ˮ��_IN,����˵��_In:��Ҫ����
        strSQL = strSQL & "1,"
        '  ����id_In       ���˽��ʼ�¼.����id%Type,
        strSQL = strSQL & "" & mPatiInfor.lng����ID & ","
        '  ����id_In       ����Ԥ����¼.����id%Type,
        strSQL = strSQL & "" & mBalanceInfor.lng����ID & ","
        '  ���㷽ʽ_In     Varchar2,
        strSQL = strSQL & "'" & str���㷽ʽ & "',"
        '  ��Ԥ��_In       ����Ԥ����¼.��Ԥ��%Type := Null,
        strSQL = strSQL & "NULL,"
        '  ��֧Ʊ��_In     ����Ԥ����¼.��Ԥ��%Type := Null,
        strSQL = strSQL & "NULL,"
        '  �����id_In     ����Ԥ����¼.�����id%Type := Null,
        strSQL = strSQL & "" & objCard.�ӿ���� & ","
        '  ����_In         ����Ԥ����¼.����%Type := Null,
        strSQL = strSQL & "'" & strCardNo & "',"
        '  ������ˮ��_In   ����Ԥ����¼.������ˮ��%Type := Null,
        strSQL = strSQL & "'" & strSwapNO & "',"
        '  ����˵��_In     ����Ԥ����¼.����˵��%Type := Null,
        strSQL = strSQL & "'" & strSwapMemo & "',"
        '  �ɿ�_In         ����Ԥ����¼.�ɿ�%Type := Null,
        strSQL = strSQL & "" & mBalanceInfor.dbl�ɿ� & ","
        '  �Ҳ�_In         ����Ԥ����¼.�Ҳ�%Type := Null,
        strSQL = strSQL & "" & mBalanceInfor.dbl�Ҳ� & ","
        '  �����_In     ������ü�¼.ʵ�ս��%Type := Null,
        strSQL = strSQL & "NULL,"
        '  ��������_In     Number := 2,
        strSQL = strSQL & "" & IIf(mEditType = g_Ed_������� Or mblnCurMzBalanceNo, 1, 2) & ","
        '  ȱʡ���㷽ʽ_In ���㷽ʽ.����%Type := Null,
        strSQL = strSQL & "NULL,"
        '    ����Ա���_In    ����Ԥ����¼.����Ա���%Type := Null,
        strSQL = strSQL & "'" & UserInfo.��� & "',"
        '    ����Ա����_In    ����Ԥ����¼.����Ա����%Type := Null,
        strSQL = strSQL & "'" & UserInfo.���� & "',"
        '    �տ�ʱ��_In      ����Ԥ����¼.����Ա����%Type := Null,
        strSQL = strSQL & "to_date('" & Format(mBalanceInfor.dtBalanceDate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
        '    ��Ԥ������ids_In Varchar2 := Null,
        strSQL = strSQL & "NULL,"
        '  ��ɽ���_In Number:=0
        strSQL = strSQL & "0)"
        zlAddArray cllPro, strSQL
    End With
    
    On Error GoTo ErrRoll:
    
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    
    With mBalanceInfor
        dblMoney = RoundEx(IIf(RoundEx(.dblδ���ϼ� - .dbl��Ԥ���ϼ�, 6) < 0, -1, 1) * dblMoney, 5)
    End With
    'zlTransferAccountsMoney(ByVal frmMain As Object, ByVal lngModule As Long,
    '     ByVal lngCardTypeID As Long, _
    '    ByVal strCardNo As String, ByVal strBalanceID As String, ByVal dblMoney As Double,
    '    Optional ByRef strSwapGlideNO As String, _
    '    Optional ByRef strSwapMemo As String, Optional ByRef strSwapExtendInfor As String,
    '    Optional ByRef strXMLExpend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����ʻ�ת��
    '���:
    '   frmMain-���õ�������
    '   lngModule-HIS����ģ���
    '   lngCardTypeID-�����ID
    '   strCardNo-����
    '   strBalanceID-����ID
    '   dblMoney-ת�ʽ��
    '    strSwapExtendInfor-�˷�ҵ��ʱ�����뱾���˷ѵĳ���ID:
    '                        ��ʽ:�շ�����1|ID1,ID2��IDn||�շ�����n|ID1,ID2��IDn
    '                        �շ�����:1-Ԥ����,2-����,3-�շ�,4-�Һ�,5-ҽ�ƿ��տ�
    '   strXMLExpend-XML��:
    '       <IN>
    '             <CZLX >��������</CZLX> //0��NULL:������ҵ��;1-�������˷�ҵ��2-����ҵ��;3-�����˷�ҵ��
    '       </IN>
    '����:
    '   strSwapGlideNO-������ˮ��
    '   strSwapMemo -����˵��
    '   strSwapExtendInfor-������չ��Ϣ: ��ʽΪ:��Ŀ����1|��Ŀ����2||��||��Ŀ����n|��Ŀ����n
    '   strXMLExpend-XML��:
    '        <OUT>
    '           <ERRMSG>������Ϣ</ERRMSG >
    '        </OUT>
    '����:���˺�
    '����:2014-09-03 14:22:10
    '������:ҽ���������(����ʱ����)
    '˵��:
    '  ��. ��ҽ���������ʱ���е�����ת��ʱ��һЩ�Ϸ��Լ�飬������ת��ʱ�����Ի���֮��ĵȴ������������������ķ�����
    '  ��. �����ڼ�����Ҫ����ΪTrue�����������ת�ʹ��ܵĵ��á�
    '---------------------------------------------------------------------------------------------------------------------------------------------
    strXMLExpend = "<IN><CZLX>3</CZLX></IN>"
    strSwapExtendInfor = "2|" & mBalanceInfor.lng����ID
    If gobjSquare.objSquareCard.zlTransferAccountsMoney(Me, mlngModul, objCard.�ӿ����, _
        strCardNo, mBalanceInfor.lng����ID, Abs(dblMoney), strSwapNO, strSwapMemo, strSwapExtendInfor, strXMLExpend) = False Then
        gcnOracle.RollbackTrans: Call zlShowThreeSwapErrInfor(1, strXMLExpend): Exit Function
    End If
    
    Call zlAddUpdateSwapSQL(False, mBalanceInfor.lng����ID, objCard.�ӿ����, False, tyBrushCard.str����, strSwapNO, strSwapMemo, cllUpdate, 2)
'    strSQL = "Zl_�����˿���Ϣ_Insert(" & mBalanceInfor.lng����ID & "," & objCard.�ӿ���� & "," & -1 * dblMoney & ",'" & strCardNo & "'," & "'" & strSwapNO & "'," & "'" & strSwapMemo & "',0)"
'    zlAddArray cllUpdate, strSQL
    zlExecuteProcedureArrAy cllUpdate, Me.Caption, , True
    mBalanceInfor.blnSaveBill = True
    If strSwapExtendInfor <> "2|" & mBalanceInfor.lng����ID Then
        Call zlAddThreeSwapSQLToCollection(False, mBalanceInfor.lng����ID, objCard.�ӿ����, objCard.���ѿ�, strCardNo, strSwapExtendInfor, cllThreeSwap)
    End If
    Set cllBillPro = New Collection
    zlExecuteProcedureArrAy cllThreeSwap, Me.Caption
    ExecuteThreeSwapTransferPay = True
    
    Exit Function
ErrRoll:
    gcnOracle.RollbackTrans
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ExecuteThreeSwapTransferAccount(ByVal objCard As Card, _
    ByVal dblMoney As Double, ByRef cllBillPro As Collection, _
    ByRef tyBrushCard As TY_BrushCard, _
    Optional ByVal bln�쳣���� As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:һ��ͨת��֧��(�����ӿ�)
    '���:dblMoney-���ν�����
    '     cllBillPro-���ݹ���(ִ��������,�Ա�����´νӿ�ʱ�ظ�ִ��)
    '     tyBrushCard-ת��ˢ����Ϣ
    '����:ִ�гɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-07-09 18:14:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSwapNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim cllPro As Collection, blnTrans As Boolean
    Dim str����IDs As String, i As Long
    Dim cllUpdate As Collection, cllThreeSwap As Collection
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strCardNo As String, str���㷽ʽ   As String
    Dim strXMLExpend As String
    
    Set cllUpdate = New Collection
    Set cllThreeSwap = New Collection
    
    '��һ��֧ͨ��,ֱ�ӷ���
    If objCard.�ӿ���� <= 0 Or objCard.���ѿ� Then ExecuteThreeSwapTransferAccount = True: Exit Function
    If Not objCard.�Ƿ�ת�ʼ����� Then Exit Function
    
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    
    Err = 0: On Error GoTo ErrHand:
  
    With mrsBalance
        strCardNo = tyBrushCard.str����
        strSwapNO = tyBrushCard.str������ˮ��
        strSwapMemo = tyBrushCard.str����˵��
        str���㷽ʽ = objCard.���㷽ʽ
        
        '���㷽ʽ|������|�������|����ժҪ||..
        str���㷽ʽ = str���㷽ʽ & "|" & -1 * dblMoney
        str���㷽ʽ = str���㷽ʽ & "|" & " "
        str���㷽ʽ = str���㷽ʽ & "|" & "ת�ʽ���"
        
        'Zl_���˽�������_Modify
        strSQL = "Zl_���˽�������_Modify("
        '  ��������_In   Number,
        '--   1-��ͨ�˷ѷ�ʽ:
        '--     �ٽ��㷽ʽ_IN:��������,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ||.." ;Ҳ�������.
        '--   2.�������˷ѽ���:
        '--     �ٽ��㷽ʽ_IN:ֻ�ܴ���һ�����㷽ʽ,���������һЩ������Ϣ,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ"
        '--     �ڿ����ID_IN,����_IN,������ˮ��_IN,����˵��_In:��Ҫ����
        '--   3-ҽ������(�������ҽ���Ľ���,��Ҫ��ɾ��ԭҽ������,���´���ĸ���)
        '--     �ٽ��㷽ʽ_IN:��������,��ʽΪ:���㷽ʽ|������||.."
        '--   4-���ѿ�����:
        '--     �ٽ��㷽ʽ_IN:����һ��ˢ���ſ�,��ʽΪ:�����ID|����|���ѿ�ID|���ѽ��||."
        strSQL = strSQL & "" & 2 & ","
        '  ����id_In     ���˽��ʼ�¼.����id%Type,
        strSQL = strSQL & "" & mPatiInfor.lng����ID & ","
        '  ����id_In     ����Ԥ����¼.����id%Type,
        strSQL = strSQL & "" & mBalanceInfor.lng����ID & ","
        '  ���㷽ʽ_In   Varchar2,
        strSQL = strSQL & "'" & str���㷽ʽ & "',"
        '  �����id_In   ����Ԥ����¼.�����id%Type := Null,
        strSQL = strSQL & "" & objCard.�ӿ���� & ","
        '  ����_In       ����Ԥ����¼.����%Type := Null,
        strSQL = strSQL & "'" & strCardNo & "',"
        '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
        strSQL = strSQL & "'" & strSwapNO & "',"
        '  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
        strSQL = strSQL & "'" & strSwapMemo & "')"
        '  �ɿ�_In       ����Ԥ����¼.�ɿ�%Type := Null,
        '  �Ҳ�_In       ����Ԥ����¼.�Ҳ�%Type := Null,
        '  �����_In   ����Ԥ����¼.��Ԥ��%Type := Null,
        '  Ԥ�����_In   ����Ԥ����¼.��Ԥ��%Type := Null,
        '  �������_In Number:=0
        If bln�쳣���� = False Then zlAddArray cllPro, strSQL
    End With
    
    On Error GoTo ErrRoll:
    
    str����IDs = mBalanceInfor.lng����ID & IIf(mBalanceInfor.lng����ID <> 0, "," & mBalanceInfor.lng����ID, "")
    
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    
    'zlTransferAccountsMoney(ByVal frmMain As Object, ByVal lngModule As Long,
    '     ByVal lngCardTypeID As Long, _
    '    ByVal strCardNo As String, ByVal strBalanceID As String, ByVal dblMoney As Double,
    '    Optional ByRef strSwapGlideNO As String, _
    '    Optional ByRef strSwapMemo As String, Optional ByRef strSwapExtendInfor As String,
    '    Optional ByRef strXMLExpend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����ʻ�ת��
    '���:
    '   frmMain-���õ�������
    '   lngModule-HIS����ģ���
    '   lngCardTypeID-�����ID
    '   strCardNo-����
    '   strBalanceID-����ID
    '   dblMoney-ת�ʽ��
    '    strSwapExtendInfor-�˷�ҵ��ʱ�����뱾���˷ѵĳ���ID:
    '                        ��ʽ:�շ�����1|ID1,ID2��IDn||�շ�����n|ID1,ID2��IDn
    '                        �շ�����:1-Ԥ����,2-����,3-�շ�,4-�Һ�,5-ҽ�ƿ��տ�
    '   strXMLExpend-XML��:
    '       <IN>
    '             <CZLX >��������</CZLX> //0��NULL:������ҵ��;1-�������˷�ҵ��2-����ҵ��;3-�����˷�ҵ��
    '       </IN>
    '����:
    '   strSwapGlideNO-������ˮ��
    '   strSwapMemo -����˵��
    '   strSwapExtendInfor-������չ��Ϣ: ��ʽΪ:��Ŀ����1|��Ŀ����2||��||��Ŀ����n|��Ŀ����n
    '   strXMLExpend-XML��:
    '        <OUT>
    '           <ERRMSG>������Ϣ</ERRMSG >
    '        </OUT>
    '����:���˺�
    '����:2014-09-03 14:22:10
    '������:ҽ���������(����ʱ����)
    '˵��:
    '  ��. ��ҽ���������ʱ���е�����ת��ʱ��һЩ�Ϸ��Լ�飬������ת��ʱ�����Ի���֮��ĵȴ������������������ķ�����
    '  ��. �����ڼ�����Ҫ����ΪTrue�����������ת�ʹ��ܵĵ��á�
    '---------------------------------------------------------------------------------------------------------------------------------------------
    strXMLExpend = "<IN><CZLX>3</CZLX></IN>"
    strSwapExtendInfor = "2|" & mBalanceInfor.lng����ID
    If gobjSquare.objSquareCard.zlTransferAccountsMoney(Me, mlngModul, objCard.�ӿ����, _
        strCardNo, mBalanceInfor.lng����ID, dblMoney, strSwapNO, strSwapMemo, strSwapExtendInfor, strXMLExpend) = False Then
        gcnOracle.RollbackTrans: Call zlShowThreeSwapErrInfor(1, strXMLExpend): Exit Function
    End If
    
    Call zlAddUpdateSwapSQL(False, mBalanceInfor.lng����ID, objCard.�ӿ����, False, tyBrushCard.str����, strSwapNO, strSwapMemo, cllUpdate, 2)
    zlExecuteProcedureArrAy cllUpdate, Me.Caption, , True
    mBalanceInfor.blnSaveBill = True
    If strSwapExtendInfor <> "2|" & mBalanceInfor.lng����ID Then
        Call zlAddThreeSwapSQLToCollection(False, mBalanceInfor.lng����ID, objCard.�ӿ����, objCard.���ѿ�, strCardNo, strSwapExtendInfor, cllThreeSwap)
    End If
    Set cllBillPro = New Collection
    zlExecuteProcedureArrAy cllThreeSwap, Me.Caption
    ExecuteThreeSwapTransferAccount = True
    
    Exit Function
ErrRoll:
    gcnOracle.RollbackTrans
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub AddSquareBalance(ByVal objCard As Card)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ѿ����㷽ʽ�����㷽ʽ�б�
    '����:���˺�
    '����:2015-01-23 15:09:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllBalance As New Collection
    Dim j As Long, dblMoney As Double, strCardNo As String
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    With vsBlance
      '�����ԭʼ�����ѿ�����,�������˷�
        Call ClearSquareBalance(objCard.�ӿ����)
        Set cllBalance = mcllCurSquareBalance
        For j = 1 To cllBalance.Count
            If objCard.�ӿ���� = Val(cllBalance(j)(0)) Then
                '��ǰˢ������(array(�����ID,���ѿ�ID,ˢ�����,����,����,�������,�Ƿ�����))
                If .Rows = 1 Then .Rows = .Rows + 1
                If Not (.Rows = 2 And Trim(.TextMatrix(1, .ColIndex("���㷽ʽ"))) = "") Then
                    .Rows = .Rows + 1
'                    .RowPosition(.Rows - 1) = 1
                End If
                '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
                dblMoney = cllBalance(j)(2)
            
                .TextMatrix(.Rows - 1, .ColIndex("����")) = 5
                .TextMatrix(.Rows - 1, .ColIndex("�Ƿ�����")) = Val(cllBalance(j)(6))
                .TextMatrix(.Rows - 1, .ColIndex("��������")) = objCard.��������
                
                If zlSquareIsDelCash(objCard.�ӿ����) Then
                    .TextMatrix(.Rows - 1, .ColIndex("�༭״̬")) = 2 '0-��ֹɾ��;1-����༭���;2-����ɾ��
                Else
                    .TextMatrix(.Rows - 1, .ColIndex("�༭״̬")) = 0 '0-��ֹɾ��;1-����༭���;2-����ɾ��
                End If
                .TextMatrix(.Rows - 1, .ColIndex("����״̬")) = 0 '�Ƿ��ѽ���:1-�ѽ���;0-δ����
                .TextMatrix(.Rows - 1, .ColIndex("�����ID")) = objCard.�ӿ����
                .TextMatrix(.Rows - 1, .ColIndex("���ѿ�ID")) = Val(cllBalance(j)(1))
                .Cell(flexcpData, .Rows - 1, .ColIndex("���ѿ�ID")) = cllBalance(j)(4) '����
                .Cell(flexcpData, .Rows - 1, .ColIndex("�����ID")) = cllBalance(j)(5) '�������
                
                .TextMatrix(.Rows - 1, .ColIndex("���㷽ʽ")) = objCard.���㷽ʽ
                 strCardNo = Trim(cllBalance(j)(3))
                .TextMatrix(.Rows - 1, .ColIndex("����")) = IIf(objCard.�������Ĺ��� <> "" And objCard.�������Ĺ��� <> "0", String(Len(strCardNo), "*"), strCardNo)
                .Cell(flexcpData, .Rows - 1, .ColIndex("����")) = strCardNo
                .TextMatrix(.Rows - 1, .ColIndex("������")) = Format(dblMoney, "0.00")
                .Cell(flexcpData, .Rows - 1, .ColIndex("������")) = Format(dblMoney, "0.00")
                .TextMatrix(.Rows - 1, .ColIndex("�������")) = ""
                .TextMatrix(.Rows - 1, .ColIndex("��ע")) = ""
                .TextMatrix(.Rows - 1, .ColIndex("�Ƿ�����")) = IIf(objCard.�Ƿ�����, 1, 0)
                .TextMatrix(.Rows - 1, .ColIndex("�Ƿ�ȫ��")) = IIf(objCard.�Ƿ�ȫ��, 1, 0)
                .TextMatrix(.Rows - 1, .ColIndex("���������")) = objCard.����
                
'                If (Not (mEditType = g_Ed_�������� Or chkCancel.Value = 1)) Or ((mEditType = g_Ed_�������� Or chkCancel.Value = 1) And mBalanceInfor.blnSaveBill = False) Then
                mBalanceInfor.dbl�Ѹ��ϼ� = RoundEx(mBalanceInfor.dbl�Ѹ��ϼ� + dblMoney, 6)
                mBalanceInfor.dblδ���ϼ� = RoundEx(mBalanceInfor.dblδ���ϼ� - dblMoney, 6)
'                End If
                
            End If
        Next
    End With
End Sub

Private Sub ClearSquareBalance(ByVal lngCardTypeID As Long, _
    Optional ByVal lng���ѿ�ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ѿ�����
    '����:���˺�
    '����:2015-01-23 14:54:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, j As Long
    With vsBlance
        j = 1
        Do While j <= .Rows - 1
            If Val(.TextMatrix(j, .ColIndex("����"))) = 5 _
                And Val(.TextMatrix(j, .ColIndex("�����ID"))) = lngCardTypeID _
                And (lng���ѿ�ID = 0 Or (lng���ѿ�ID <> 0 And Val(.TextMatrix(j, .ColIndex("���ѿ�ID"))) = lng���ѿ�ID)) Then
                dblMoney = Val(.TextMatrix(j, .ColIndex("������")))
                
                mBalanceInfor.dbl�Ѹ��ϼ� = RoundEx(mBalanceInfor.dbl�Ѹ��ϼ� - dblMoney, 6)
                mBalanceInfor.dblδ���ϼ� = RoundEx(mBalanceInfor.dblδ���ϼ� + dblMoney, 6)
                If .Rows >= 2 Then
                    .RemoveItem j
                Else
                    .Rows = 2
                   .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
                   .Cell(flexcpText, 1, 0, 1, .Cols - 1) = ""
                   .RowData(1) = ""
                   j = 2
                End If
            Else
                j = j + 1
            End If
        Loop
    End With
End Sub

Private Sub vsDeposit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    vsDeposit.EditCell
    vsDeposit.EditSelStart = 0
    vsDeposit.EditSelLength = 100
End Sub

Private Sub vsDeposit_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = vsDeposit.ColIndex("��Ԥ��") Then
        If Val(vsDeposit.EditText) = Val(vsDeposit.TextMatrix(Row, Col)) Then mblnNoTrigger = True
    End If
End Sub

Private Sub vsDetailList_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    zlControl.TxtCheckKeyPress vsDeposit, KeyAscii, m���ʽ
End Sub

Private Sub vsDetailList_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long, dblMoney As Double
    With vsDetailList
        .TextMatrix(Row, Col) = Format(.TextMatrix(Row, Col), gstrDec)
        .Cell(flexcpData, Row, Col) = Val(.TextMatrix(Row, Col))
        For i = 1 To .Rows - 1
            dblMoney = dblMoney + Val(.TextMatrix(i, .ColIndex("���ʽ��")))
        Next i
    End With
    mblnNotChange = True
    txtBalance(Idx_���ν���).Text = Format(dblMoney, gstrDec)
    mblnNotChange = False
    mBalanceInfor.dbl��ǰ���� = dblMoney
    mBalanceInfor.dblδ���ϼ� = RoundEx(mBalanceInfor.dbl��ǰ���� - mBalanceInfor.dbl�Ѹ��ϼ�, 5)
    Call LoadIntendBalance
    Call LoadCurOwnerPayInfor
    If vsDetailList.Row + 1 <= vsDetailList.Rows - 1 Then
        vsDetailList.Select vsDetailList.Row + 1, vsDetailList.ColIndex("���ʽ��")
    End If
    mbln�ѱ��� = False
    
End Sub

Private Sub vsDetailList_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If mrsInfo Is Nothing Then Cancel = True: Exit Sub
    If mrsInfo.State = 0 Then Cancel = True: Exit Sub
    If mrsInfo.RecordCount = 0 Then Cancel = True: Exit Sub
    If mYBInFor.intInsure <> 0 Then Cancel = True: Exit Sub
    If InStr(mstrPrivs, ";��������;") = 0 Then Cancel = True: Exit Sub
    With vsDetailList
        If Col <> .ColIndex("���ʽ��") Then
            Cancel = True
        Else
            If .Cell(flexcpBackColor, Row, .ColIndex("���ʽ��")) = .Cell(flexcpBackColor, Row, .ColIndex("����")) _
                Or .TextMatrix(Row, .ColIndex("����")) = "" Then
                Cancel = True
            End If
            '�����������޸�
            If Val(.Cell(flexcpData, Row, .ColIndex("δ����"))) < 0 Then Cancel = True
        End If
    End With
End Sub

Private Sub vsDetailList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    vsDetailList.EditCell
    vsDetailList.EditSelStart = 0
    vsDetailList.EditSelLength = 100
End Sub

Private Sub vsDetailList_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsDetailList
        If IsNumeric(.EditText) = False And .EditText <> "" Then Cancel = True: Exit Sub
        If Val(.Cell(flexcpData, Row, .ColIndex("δ����"))) < 0 Then
            If Val(.EditText) > 0 Then Cancel = True: Exit Sub
            If Val(.EditText) < Val(.Cell(flexcpData, Row, .ColIndex("δ����"))) Then
                .EditText = Val(.Cell(flexcpData, Row, .ColIndex("δ����")))
            End If
        Else
            If Val(.EditText) < 0 Then Cancel = True: Exit Sub
            If Val(.EditText) > Val(.Cell(flexcpData, Row, .ColIndex("δ����"))) Then
                .EditText = Val(.Cell(flexcpData, Row, .ColIndex("δ����")))
            End If
        End If
    End With
End Sub

Private Sub vsFeeList_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModul, vsFeeList, Me.Name, "�����б�" & mEditType
End Sub

Private Sub vsFeeList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    zl_VsGridRowChange vsFeeList, OldRow, NewRow, OldCol, NewCol
End Sub

Private Sub vsFeeList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModul, vsFeeList, Me.Name, "�����б�" & mEditType
End Sub

Private Sub vsFeeList_GotFocus()
    zl_VsGridGotFocus vsFeeList, &HFFC0C0
End Sub
Private Sub vsFeeList_LostFocus()
   zl_VsGridLOSTFOCUS vsFeeList
End Sub

Private Sub vsDetailList_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModul, vsDetailList, Me.Name, "��ϸ�б�" & mEditType
End Sub

Private Sub vsDetailList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    On Error Resume Next
    vsDetailList.Cell(flexcpBackColor, OldRow, 0, OldRow, 3) = vbWhite
    vsDetailList.Cell(flexcpBackColor, NewRow, 0, NewRow, 3) = 16772055
'    vsDetailList.Select NewRow, 4
End Sub

Private Sub vsDetailList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModul, vsDetailList, Me.Name, "��ϸ�б�" & mEditType
End Sub

Private Sub vsDetailList_GotFocus()
    vsDetailList.Cell(flexcpBackColor, vsDetailList.Row, 0, vsDetailList.Row, 3) = 16772055
End Sub
Private Sub vsDetailList_LostFocus()
    vsDetailList.Cell(flexcpBackColor, vsDetailList.Row, 0, vsDetailList.Row, 3) = GRD_LOSTFOCUS_COLORSEL
End Sub

Private Sub vsDeposit_GotFocus()
    zl_VsGridGotFocus vsDeposit, &HFFC0C0
End Sub
Private Sub vsDeposit_LostFocus()
   zl_VsGridLOSTFOCUS vsDeposit
End Sub
Private Sub vsBlance_GotFocus()
    zl_VsGridGotFocus vsBlance, &HFFC0C0
End Sub
Private Sub vsBlance_LostFocus()
   zl_VsGridLOSTFOCUS vsBlance
End Sub
Private Function GetOldBalanceMoney(ByVal int���� As Integer, _
    ByVal objCard As Card) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ͣ�ȷ��ԭ���㷽ʽ�Ľ��
    '���:int����-����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
    '����:����ԭ������
    '����:���˺�
    '����:2015-01-30 17:33:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, rsBalance As ADODB.Recordset
    
    On Error GoTo errHandle
    If Not mrsOldBalance Is Nothing Then
        Set rsBalance = mrsOldBalance
    Else
        Set rsBalance = mrsBalance
    End If
    If rsBalance Is Nothing Then Exit Function
    If rsBalance.State <> 1 Then Exit Function
     
    If objCard.�ӿ���� > 0 Then
        If objCard.���ѿ� = False Then 'һ��ͨ
            rsBalance.Filter = "����=" & int���� & " And �����ID=" & objCard.�ӿ����
        Else '���ѿ�
            rsBalance.Filter = "����=" & int���� & " And ���㿨���=" & objCard.�ӿ����
        End If
    Else
        rsBalance.Filter = "����=" & int����
    End If
    
    If rsBalance.EOF Then
        If objCard.�Ƿ�ת�ʼ����� Then
           GetOldBalanceMoney = RoundEx(Val(lblʣ���Ը�.Caption), 6)
        End If
        rsBalance.Filter = 0: Exit Function
    End If
    
    rsBalance.MoveFirst
    Do While Not rsBalance.EOF
        dblMoney = dblMoney + Val(NVL(rsBalance!��Ԥ��))
        rsBalance.MoveNext
    Loop
    GetOldBalanceMoney = dblMoney
    rsBalance.Filter = 0
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function PatiErrBillPay(ByVal lng����ID As Long, Optional ByVal strCheckNO As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݲ���,���쳣���ݽ������½��ʻ����ϴ���
    '���:lng����ID-ָ���Ĳ���ID
    '����:�����쳣����,���ɹ���ȡ�쳣���ݷ���true,���򷵻�False
    '����:���˺�
    '����:2015-02-03 11:30:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim strNO As String, lng����ID As Long
    Dim str����Ա���� As String, strTittle As String
    Dim blnDel As Boolean, blnErrCancel As Boolean
    Dim strDelTime As String
    
    If mEditType = g_Ed_���ݲ鿴 Then Exit Function
'    If mEditType = g_Ed_������� Or mEditType <> g_Ed_סԺ���� Then Exit Function
    
    On Error GoTo errHandle
    If strCheckNO = "" Then
        strSQL = " " & _
        "    Select  a.No, a.ID, a.����Ա����, decode(��¼״̬,2,2,1) As �쳣����,A.�շ�ʱ�� " & _
        "    From ���˽��ʼ�¼ A" & _
        "    Where nvl(����״̬,0) = 1 and ����ID=[1]   And Rownum < 2 "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    Else
        strSQL = " " & _
        "    Select  a.No, a.ID, a.����Ա����, decode(��¼״̬,2,2,1) As �쳣����,A.�շ�ʱ�� " & _
        "    From ���˽��ʼ�¼ A" & _
        "    Where nvl(����״̬,0) = 1 and NO=[1]   And Rownum < 2 "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strCheckNO)
    End If
    If rsTemp.EOF Then
        If strCheckNO <> "" Then PatiErrBillPay = True
        Exit Function
    End If
    
    strNO = NVL(rsTemp!NO): lng����ID = Val(NVL(rsTemp!ID))
    blnDel = Val(NVL(rsTemp!�쳣����)) = 2
    strTittle = IIf(Not blnDel, "����", "����")
    strDelTime = Format(rsTemp!�շ�ʱ��, "yyyy-mm-dd HH:MM:SS")
    str����Ա���� = NVL(rsTemp!����Ա����)
    
    If str����Ա���� <> UserInfo.���� Then
        '100703
         If MsgBox("ע��:" & vbCrLf & _
                            "       �ò��˴����쳣��" & strTittle & "����" & IIf(str����Ա���� <> UserInfo.����, ",�õ����ǲ���Ա[" & str����Ա���� & "]��ȡ��,���޷�����," & vbCrLf, "") & " ,�Ƿ񲻶��쳣���ݽ��д���,�������н��ʲ���" & "?" & vbCrLf & vbCrLf & _
                            "���ǡ��������쳣���ݽ��д���,�������н��ʲ���. " & vbCrLf & _
                            "���񡻴�����ֹ���ʲ���.", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
            PatiErrBillPay = False
            Exit Function
         Else
            PatiErrBillPay = True
            Exit Function
         End If
    End If
    
    If MsgBox("ע��:" & vbCrLf & _
                        "       �ò��˴����쳣��" & strTittle & "����" & IIf(str����Ա���� <> UserInfo.����, ",�õ����ǲ���Ա[" & str����Ա���� & "]��ȡ��," & vbCrLf, "") & " ,�Ƿ����¶Ըõ��ݽ���" & strTittle & "?" & vbCrLf & vbCrLf & _
                        "���ǡ��������¶��쳣���� " & strTittle & vbCrLf & _
                        "���񡻴������쳣���ݽ��д���,�������н��ʲ���.", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
        Exit Function
    End If

    If strCheckNO <> "" Then
        PatiErrBillPay = True
        Exit Function
    End If
    
    mintPreEditType = mEditType
    mEditType = IIf(blnDel, g_Ed_��������, g_Ed_���½���)
    mblnViewCancel = blnDel
    Call SetFeeListColumnShow
    Call SetPatiConsControlVisible
    Call SetOperatonCommandCaption
    
    If ReadBalance(strNO) Then PatiErrBillPay = True: Exit Function
    mEditType = mintPreEditType
    Call LoadBalanceBill
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function DeleteBalance(Optional blnDelBalance As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ϴ���(�쳣����)
    '���:blnDelBalance-true-��������;false-�쳣����
    '����:���ϳɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2015-02-03 16:36:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objTemp As Card, tyBrushCard As TY_BrushCard
    Dim lng����ID As Long, lngCount As Long, dblMoney As Double
    Dim i As Long, strBalance As String, strSQL As String
    Dim cllPro As Collection
    Dim strYbBalance As String
    Dim rsTmp As ADODB.Recordset
    
    
    On Error GoTo errHandle
    If mYBInFor.intInsure > 0 Then
        If mYBInFor.bytMCMode <> 1 Then
            If Not MCPAR.��Ժ���˽������� Then
                If Not isYBPati(mPatiInfor.lng����ID, True) Then
                    MsgBox "�òα������Ѿ���Ժ������ȡ���ý��ʵ���", vbInformation, gstrSysName: Exit Function
                End If
            End If
            If MCPAR.סԺ�������� = False Then
                MsgBox "��ҽ����֧��סԺ�������ϣ�����ȡ���ý��ʵ���", vbInformation, gstrSysName: Exit Function
            End If
        ElseIf mYBInFor.bytMCMode = 1 And Not MCPAR.���ﲡ�˽������� Then
                MsgBox "��ҽ����֧������������ϣ�����ȡ���ý��ʵ���", vbInformation, gstrSysName: Exit Function
        End If
        If gclsInsure.CheckInsureValid(mYBInFor.intInsure) = False Then Exit Function
    End If
    
    Set objTemp = Nothing
    With vsBlance
        For i = 1 To .Rows - 1
            strBalance = Trim(.TextMatrix(i, .ColIndex("���㷽ʽ")))
            
            If strBalance <> "" Then
                '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
                
                Select Case Val(.TextMatrix(i, .ColIndex("����")))
                Case 0 '��ͨ����
                Case 1 'Ԥ����
                Case 2 'ҽ��
                    strYbBalance = strYbBalance & "," & strBalance
                    
                Case 3 'һ��ͨ
                    Set objTemp = GetLocalePayCard(Val(.TextMatrix(i, .ColIndex("�����ID"))), False)
                    If objTemp Is Nothing Then
                        MsgBox "��ǰվ�㲻֧��" & strBalance & "��ʽ�����˷Ѵ���,����������!", vbInformation + vbOKOnly, gstrSysName
                        Exit Function
                    End If
                     dblMoney = RoundEx(Val(.TextMatrix(i, .ColIndex("������"))), 6)
                    If CheckThreeSwapValied(objTemp, dblMoney, tyBrushCard, True) = False Then Exit Function
                    lngCount = lngCount + 1
                Case 4 'һ��ͨ(�ϰ汾)
                    dblMoney = RoundEx(Val(.TextMatrix(i, .ColIndex("������"))), 6)
                    Set objTemp = GetLocaleOldOneCard(strBalance)
                    If objTemp Is Nothing Then
                        MsgBox "��ǰվ�㲻֧��" & strBalance & "�����˷Ѵ���,����������!", vbInformation + vbOKOnly, gstrSysName
                        Exit Function
                    End If
                    If CheckOldOneCardIsValied(objTemp, dblMoney, tyBrushCard, True) = False Then Exit Function
                    lngCount = lngCount + 1
                Case 5 '���ѿ�
                Case Else
                End Select
            End If
        Next
    End With
    
    If Not mrsBalance Is Nothing Then
        mrsBalance.Filter = 0
        If mrsBalance.RecordCount <> 0 Then mrsBalance.MoveFirst
        Do While Not mrsBalance.EOF
            If Val(NVL(mrsBalance!����)) = 2 And InStr(strYbBalance & ",", "," & mrsBalance!���㷽ʽ & ",") = 0 Then
                MsgBox "��ҽ����֧�֡�" & mrsBalance!���㷽ʽ & "��ԭ���˻ش���,����������!", vbInformation + vbOKOnly, gstrSysName
                Exit Function
            End If
            mrsBalance.MoveNext
        Loop
        If mrsBalance.RecordCount <> 0 Then mrsBalance.MoveFirst
    End If
    
    strSQL = "Select 1 From �����˿���Ϣ Where ����ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mBalanceInfor.lng����ID)
    If Not rsTmp.EOF Then
        MsgBox IIf(blnDelBalance, "����", "�쳣") & "�����ݲ�֧�ְ�������˿�ӿڵĽ���,�������շѽ���!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If lngCount + IIf(mYBInFor.intInsure > 0, 1, 0) > 1 Then
        MsgBox IIf(blnDelBalance, "����", "�쳣") & "�����ݲ�֧�����ֽӿ����ϵĽ���,�������ɽ���!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    lng����ID = zlDatabase.GetNextId("���˽��ʼ�¼")
    With mBalanceInfor
        .lng����ID = lng����ID
        .dtBalanceDate = zlDatabase.Currentdate
    End With
    
    Set cllPro = New Collection
     '���˽����¼������
     strSQL = "Zl_���˽��ʼ�¼_Cancel("
     '  No_In         ���˽��ʼ�¼.No%Type,
     strSQL = strSQL & "'" & mBalanceInfor.strNO & "',"
     '  ����id_In     ���˽��ʼ�¼.Id%Type,
     strSQL = strSQL & "" & lng����ID & ","
     '  ����Ա���_In ���˽��ʼ�¼.����Ա���%Type,
     strSQL = strSQL & "'" & UserInfo.��� & "',"
     '  ����Ա����_In ���˽��ʼ�¼.����Ա����%Type
     strSQL = strSQL & "'" & UserInfo.���� & "')"
     zlAddArray cllPro, strSQL
     
     
    'Zl_���˽�������_Modify
    strSQL = "Zl_���˽�������_Modify("
    '  ��������_In   Number,
    strSQL = strSQL & "" & 0 & ","
    '  ����id_In     ���˽��ʼ�¼.����id%Type,
    strSQL = strSQL & "" & mPatiInfor.lng����ID & ","
    '  ����id_In     ����Ԥ����¼.����id%Type,
    strSQL = strSQL & "" & lng����ID & ","
    '  ���㷽ʽ_In   Varchar2,
    strSQL = strSQL & "NULL,"
    '  �����id_In   ����Ԥ����¼.�����id%Type := Null,
    strSQL = strSQL & "NULL,"
    '  ����_In       ����Ԥ����¼.����%Type := Null,
    strSQL = strSQL & "NULL,"
    '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
    strSQL = strSQL & "NULL,"
    '  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
    strSQL = strSQL & "NULL,"
    '  �ɿ�_In       ����Ԥ����¼.�ɿ�%Type := Null,
    strSQL = strSQL & "NULL,"
    '  �Ҳ�_In       ����Ԥ����¼.�Ҳ�%Type := Null,
    strSQL = strSQL & "NULL,"
    '  �����_In   ����Ԥ����¼.��Ԥ��%Type := Null,
    strSQL = strSQL & "NULL,"
    '  Ԥ�����_In   ����Ԥ����¼.��Ԥ��%Type := Null,
    strSQL = strSQL & "NULL,"
    '����Ա���_In    ����Ԥ����¼.����Ա���%Type := Null,
    strSQL = strSQL & "'" & UserInfo.��� & "',"
    '����Ա����_In    ����Ԥ����¼.����Ա����%Type := Null,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '�տ�ʱ��_In      ����Ԥ����¼.����Ա����%Type := Null,
    strSQL = strSQL & "to_date('" & Format(mBalanceInfor.dtBalanceDate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
    '��Ԥ������ids_In Varchar2 := Null,
    ' ����ö��ַ���,��Ԥ��ʱ��Ч(��Ԥ�������ҵ���������һ��),��Ҫ��ʹ�ü�����Ԥ����
    strSQL = strSQL & "NULL,"
    '  �������_In Number:=0
    strSQL = strSQL & "1)"
    zlAddArray cllPro, strSQL
    'ִ��ҽ���˷Ѳ���
    If ExecuteInsureDel(cllPro, True) = False Then Exit Function
    
    If Not objTemp Is Nothing Then
        If ExecuteThreeSwapDelInterface(objTemp, dblMoney, cllPro, True) = False Then Exit Function
        If ExecuteOneCardDelInterface(objTemp, dblMoney, cllPro, True) = False Then Exit Function
    End If
    
    strSQL = "Zl_���˽����쳣_Update("
    strSQL = strSQL & "to_date('" & Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
    strSQL = strSQL & "" & lng����ID & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    DeleteBalance = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 

Private Function CheckPatiFromZyNumIsYB(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByRef intInsure As Integer, Optional strInsureName As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ѡ��ָ��סԺ���Ƿ�Ϊҽ������
    '���:
    '����:intInsure-����ҽ�����
    '     strInsureName-ҽ������
    '����:��ҽ������true,���򷵻�False
    '����:���˺�
    '����:2017-11-13 09:53:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    intInsure = 0
    If Not mobjBalanceAll.rsAllTime Is Nothing Then
        With mobjBalanceAll.rsAllTime
            If .State = 1 Then
                .Filter = "��ҳID=" & lng��ҳID
                If Not .EOF Then
                    intInsure = Val(NVL(!����))
                    strInsureName = Trim(NVL(!��������))
                    CheckPatiFromZyNumIsYB = intInsure <> 0
                    Exit Function
                End If
            End If
        End With
    End If
    
    strSQL = "Select Nvl(a.����,0) As ����,b.���� From ������ҳ A,�������  b Where a.����=b.���(+) and A.����ID = [1] And A.��ҳID =[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(NVL(mrsInfo!����ID)), lng��ҳID)
    If rsTemp.EOF Then Exit Function
    
    intInsure = Val(NVL(rsTemp!����))
    strInsureName = Trim(NVL(rsTemp!��������))
    CheckPatiFromZyNumIsYB = intInsure <> 0
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetFeeFromType() As String
    '��ȡ�շѵ�����Դ����
    '���أ�������Դ������ö��ŷָ�
    Dim i As Long
    Dim str������Դ As String, byt������Դ As Byte
    
    On Error GoTo errHandle
    If mEditType = g_Ed_������� Or mblnCurMzBalanceNo Then '����
        str������Դ = ""
    Else 'סԺ
        GetFeeFromType = "2": Exit Function
    End If
    
    With vsDetailList
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("����")) <> "" Then
                If Not (Val(.Cell(flexcpData, i, .ColIndex("���ʽ��"))) = 0 And Val(.Cell(flexcpData, i, .ColIndex("δ����"))) <> 0) Then
                    byt������Դ = Decode(Val(.Cell(flexcpData, i, .ColIndex("���"))), 4, 3, 2, 2, 1)
                    If InStr(str������Դ, byt������Դ) = 0 Then
                        str������Դ = str������Դ & "," & byt������Դ
                    End If
                End If
            End If
        Next
    End With
    If Left(str������Դ, 1) = "," Then str������Դ = Mid(str������Դ, 2)
    GetFeeFromType = str������Դ
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function DepositMonyVerfy(Optional blnSaveCheck As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:Ԥ�����Ϸ���ٮ��
    '���:blnSaveCheck-true:������ʱ��ûЧ�Եļ��;False-�ı���У�Լ��(valied�¼�����)
    '����:
    '����:У�Գɹ�rue,���򷵻�Fale
    '����:���˺�
    '����:2017-12-28 11:31:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, blnNoRecal As Boolean
    
    On Error GoTo errHandle
    
    If chkDeposit.Visible Then DepositMonyVerfy = True: Exit Function

            
    dblMoney = RoundEx(Val(txtBalance(Idx_��Ԥ��).Text), 6)
    
    If mblnNotChange = False Then
        If Val(dblMoney) > Val(mPatiInfor.dblʵ�����) Then
            MsgBox "��ǰ����ĳ�Ԥ������Ԥ�����,���ܼ���!" & vbCrLf & "ʵ�����:" & Format(mPatiInfor.dblʵ�����, "0.00") & vbCrLf & "��Ԥ��:" & Format(Val(txtBalance(Idx_��Ԥ��).Text), "0.00")
            Exit Function
        End If
    End If
    
    blnNoRecal = dblMoney = mBalanceInfor.dbl��Ԥ���ϼ� And dblMoney <> 0
    
    If blnNoRecal = False Then
        '�����ȣ��Ͳ��������¼���
        If GetDepositTotal = dblMoney Then mBalanceInfor.dbl��Ԥ���ϼ� = dblMoney
    End If
    
    '��������(0-������г�Ԥ��;1-��ȱʡʹ��Ԥ����;2-�����ʽ������Ԥ��(��ʱ���Ⱥ�����̯��;3-ȫ��
    If dblMoney <> mBalanceInfor.dbl��Ԥ���ϼ� And mBalanceInfor.blnԤ����֤ = False Then
        If dblMoney = 0 Then
            Call RecalcDepositMoney(0)
        Else
            Call RecalcDepositMoney(2, dblMoney)
        End If
        
        mblnNotChange = True
        txtBalance(Idx_��Ԥ��).Text = Format(mBalanceInfor.dbl��Ԥ���ϼ�, "0.00")
        mblnNotChange = False
    End If
    If mblnNotChange Then DepositMonyVerfy = True: Exit Function
    
    If Not mBalanceInfor.blnԤ����֤ Then
        If CheckDepositValied(True) = False Then Exit Function
    End If
    
    If Not blnNoRecal Then
        Call LoadIntendBalance
    End If
    If blnSaveCheck And dblMoney = mBalanceInfor.dbl��Ԥ���ϼ� Then
        '���δ�����仯�����ټ��������Ϣ���
        DepositMonyVerfy = True: Exit Function
    End If
    Call LoadCurOwnerPayInfor
    Call Set�Ҳ���Ϣ
    
    DepositMonyVerfy = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetDepositTotal(Optional ByVal bln��� As Boolean = False) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��Ԥ���ܶ������ܶ�
    '���:bln���-��ȡ����ܶ�
    '����:
    '����:���س�Ԥ���ܶ������ܶ�
    '����:���˺�
    '����:2017-12-28 11:31:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intCol As Integer, i As Long
    Dim dblTemp As Double
    With vsDeposit
        dblTemp = 0
        For i = 1 To .Rows - 1
            intCol = IIf(bln���, .ColIndex("���"), .ColIndex("��Ԥ��"))
            If intCol >= 0 Then
              dblTemp = dblTemp + Val(.TextMatrix(i, intCol))
            End If
        Next i
        dblTemp = RoundEx(dblTemp, 5)
    End With
End Function



