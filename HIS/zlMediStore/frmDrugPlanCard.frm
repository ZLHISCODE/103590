VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmDrugPlanCard 
   Caption         =   "ҩƷ�ɹ��ƻ�"
   ClientHeight    =   6975
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11760
   Icon            =   "frmDrugPlanCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   11760
   StartUpPosition =   2  '��Ļ����
   Begin VB.CheckBox chk�Ƿ���ʾ������ 
      Caption         =   "��ʾ�����������"
      Height          =   240
      Left            =   6600
      TabIndex        =   58
      Top             =   6360
      Width           =   1932
   End
   Begin VB.PictureBox pic�ⷿ 
      BorderStyle     =   0  'None
      Height          =   2385
      Left            =   6600
      ScaleHeight     =   2385
      ScaleWidth      =   3855
      TabIndex        =   53
      Top             =   2760
      Visible         =   0   'False
      Width           =   3855
      Begin VB.CheckBox chk�ⷿ 
         Appearance      =   0  'Flat
         Caption         =   "ȫѡ"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   20
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   1983
         Width           =   675
      End
      Begin VB.CommandButton cmdȡ�� 
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   2720
         TabIndex        =   55
         Top             =   1920
         Width           =   1100
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   1560
         TabIndex        =   54
         Top             =   1920
         Width           =   1100
      End
      Begin MSComctlLib.ListView lvw�洢�ⷿ 
         Height          =   1935
         Left            =   0
         TabIndex        =   57
         Top             =   0
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   3413
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "img16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "����"
            Object.Width           =   3528
         EndProperty
      End
      Begin MSComctlLib.ImageList img16 
         Left            =   0
         Top             =   0
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
               Picture         =   "frmDrugPlanCard.frx":014A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDrugPlanCard.frx":69AC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDrugPlanCard.frx":D20E
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picStock 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   1680
      Left            =   2520
      ScaleHeight     =   1680
      ScaleWidth      =   8775
      TabIndex        =   44
      Top             =   2160
      Visible         =   0   'False
      Width           =   8775
      Begin VB.PictureBox picHeadStock 
         BackColor       =   &H00FFEDDD&
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   0
         ScaleHeight     =   300
         ScaleWidth      =   8535
         TabIndex        =   46
         Tag             =   "0"
         Top             =   0
         Width           =   8535
         Begin VB.CheckBox chk���пⷿ 
            BackColor       =   &H00FFEDDD&
            Caption         =   "���пⷿ"
            Height          =   180
            Left            =   1080
            TabIndex        =   50
            Top             =   48
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chk��Դ�ⷿ 
            BackColor       =   &H00FFEDDD&
            Caption         =   "��Դ�ⷿ"
            Height          =   180
            Left            =   2400
            TabIndex        =   49
            Top             =   48
            Width           =   1095
         End
         Begin VB.CheckBox chk��Դҩ�� 
            BackColor       =   &H00FFEDDD&
            Caption         =   "��Դҩ��"
            Height          =   180
            Left            =   3720
            TabIndex        =   48
            Top             =   48
            Width           =   1095
         End
         Begin VB.CommandButton cmd�ⷿ 
            Caption         =   "��"
            Height          =   285
            Left            =   5040
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   -4
            Width           =   285
         End
         Begin VB.Label lblStock 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFEDDD&
            Caption         =   "�������"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   0
            TabIndex        =   52
            Top             =   48
            Width           =   720
         End
         Begin VB.Label lbl�Զ���ⷿ 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�Զ���ⷿ"
            Height          =   180
            Left            =   5280
            TabIndex        =   51
            Top             =   45
            Width           =   975
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfStock 
         Height          =   1200
         Left            =   10
         TabIndex        =   45
         Top             =   330
         Width           =   8295
         _cx             =   14631
         _cy             =   2117
         Appearance      =   0
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
         BackColorSel    =   16761024
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   0
         GridColorFixed  =   0
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
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
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmDrugPlanCard.frx":13A70
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
   Begin VB.CheckBox chk���ؽ��ڲɹ��ƻ� 
      Caption         =   "���ؽ��ڲɹ��ƻ�"
      Height          =   240
      Left            =   4440
      TabIndex        =   43
      Top             =   6360
      Width           =   1932
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Left            =   3720
      TabIndex        =   8
      Top             =   5850
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "����(&F)"
      Height          =   350
      Left            =   2040
      TabIndex        =   7
      Top             =   5760
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   240
      TabIndex        =   6
      Top             =   5760
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6240
      TabIndex        =   4
      Top             =   5760
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7560
      TabIndex        =   5
      Top             =   5760
      Width           =   1100
   End
   Begin VB.PictureBox Pic���� 
      BackColor       =   &H80000004&
      Height          =   5325
      Left            =   0
      ScaleHeight     =   5265
      ScaleWidth      =   11655
      TabIndex        =   9
      Top             =   0
      Width           =   11715
      Begin VB.PictureBox picHis 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   1755
         Left            =   240
         ScaleHeight     =   1755
         ScaleWidth      =   8775
         TabIndex        =   32
         Top             =   1800
         Width           =   8775
         Begin VB.PictureBox picHscSend 
            BackColor       =   &H00FFEDDD&
            BorderStyle     =   0  'None
            Height          =   300
            Left            =   0
            ScaleHeight     =   300
            ScaleWidth      =   8535
            TabIndex        =   34
            Tag             =   "0"
            Top             =   0
            Width           =   8535
            Begin VB.PictureBox picColor 
               BackColor       =   &H000000FF&
               BorderStyle     =   0  'None
               Height          =   200
               Index           =   2
               Left            =   7080
               ScaleHeight     =   195
               ScaleWidth      =   255
               TabIndex        =   41
               Top             =   45
               Width           =   255
            End
            Begin VB.PictureBox picColor 
               BackColor       =   &H008080FF&
               BorderStyle     =   0  'None
               Height          =   200
               Index           =   1
               Left            =   5640
               ScaleHeight     =   195
               ScaleWidth      =   255
               TabIndex        =   39
               Top             =   45
               Width           =   255
            End
            Begin VB.PictureBox picColor 
               BackColor       =   &H00FF0000&
               BorderStyle     =   0  'None
               Height          =   200
               Index           =   0
               Left            =   4200
               ScaleHeight     =   195
               ScaleWidth      =   255
               TabIndex        =   37
               Top             =   45
               Width           =   255
            End
            Begin VB.CheckBox chkMore 
               BackColor       =   &H00FFEDDD&
               Caption         =   "����"
               Height          =   240
               Left            =   2880
               TabIndex        =   36
               Top             =   25
               Width           =   855
            End
            Begin VB.Label lblColorTxt 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFEDDD&
               Caption         =   "û��ִ��"
               Height          =   180
               Index           =   2
               Left            =   7440
               TabIndex        =   42
               Top             =   45
               Width           =   720
            End
            Begin VB.Label lblColorTxt 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFEDDD&
               Caption         =   "����ִ��"
               Height          =   180
               Index           =   1
               Left            =   6000
               TabIndex        =   40
               Top             =   45
               Width           =   720
            End
            Begin VB.Label lblColorTxt 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFEDDD&
               Caption         =   "��ȫִ��"
               Height          =   180
               Index           =   0
               Left            =   4560
               TabIndex        =   38
               Top             =   45
               Width           =   720
            End
            Begin VB.Image imgDown 
               Height          =   240
               Left            =   0
               Picture         =   "frmDrugPlanCard.frx":13B3A
               Top             =   0
               Visible         =   0   'False
               Width           =   240
            End
            Begin VB.Image imgUp 
               Height          =   240
               Left            =   0
               Picture         =   "frmDrugPlanCard.frx":13E7C
               Top             =   0
               Width           =   240
            End
            Begin VB.Label lblDiag 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFEDDD&
               Caption         =   "���ڲɹ��ƻ�ִ�����"
               ForeColor       =   &H00FF0000&
               Height          =   180
               Left            =   480
               TabIndex        =   35
               Top             =   50
               Width           =   1800
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfHisPlane 
            Height          =   1400
            Left            =   0
            TabIndex        =   33
            Top             =   315
            Width           =   2760
            _cx             =   4868
            _cy             =   2469
            Appearance      =   0
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
            BackColorSel    =   16761024
            ForeColorSel    =   -2147483640
            BackColorBkg    =   -2147483633
            BackColorAlternate=   -2147483643
            GridColor       =   0
            GridColorFixed  =   0
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   16777215
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
            Cols            =   11
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   255
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmDrugPlanCard.frx":141BE
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
      Begin ZL9BillEdit.BillEdit mshBill 
         Height          =   2805
         Left            =   195
         TabIndex        =   1
         Top             =   950
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   4948
         Appearance      =   0
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Active          =   -1  'True
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
         ColWidth0       =   1005
         BackColor       =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483634
      End
      Begin VB.TextBox txtժҪ 
         Height          =   300
         Left            =   900
         MaxLength       =   40
         TabIndex        =   3
         Top             =   4080
         Width           =   10410
      End
      Begin VB.Label lbl������ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   8745
         TabIndex        =   31
         Top             =   4500
         Width           =   540
      End
      Begin VB.Label txt������ 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   9330
         TabIndex        =   30
         Top             =   4440
         Width           =   1890
      End
      Begin VB.Label lbl�������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Left            =   8520
         TabIndex        =   29
         Top             =   4860
         Width           =   720
      End
      Begin VB.Label txt�������� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   9330
         TabIndex        =   28
         Top             =   4800
         Width           =   1890
      End
      Begin VB.Label lbl���Ʒ��� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���Ʒ���:"
         Height          =   180
         Left            =   8070
         TabIndex        =   25
         Top             =   660
         Width           =   810
      End
      Begin VB.Label txt���Ʒ��� 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ٽ��ڼ�ƽ�����շ�"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   9000
         TabIndex        =   24
         Top             =   660
         Width           =   2355
      End
      Begin VB.Label txt�ƻ����� 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   1080
         TabIndex        =   23
         Top             =   660
         Width           =   1845
      End
      Begin VB.Label lblPurchasePrice 
         AutoSize        =   -1  'True
         Caption         =   "���ϼƣ�"
         Height          =   180
         Left            =   240
         TabIndex        =   22
         Top             =   3840
         Width           =   900
      End
      Begin VB.Label Txt����� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   5400
         TabIndex        =   20
         Top             =   4440
         Width           =   1890
      End
      Begin VB.Label Txt������� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   5400
         TabIndex        =   19
         Top             =   4800
         Width           =   1890
      End
      Begin VB.Label Txt�������� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   18
         Top             =   4800
         Width           =   1890
      End
      Begin VB.Label Txt������ 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   17
         Top             =   4440
         Width           =   1890
      End
      Begin VB.Label txtNo 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   9960
         TabIndex        =   16
         Top             =   158
         Width           =   1425
      End
      Begin VB.Label LblNo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NO."
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   9480
         TabIndex        =   15
         Top             =   195
         Width           =   480
      End
      Begin VB.Label lblժҪ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ժҪ(&M)"
         Height          =   180
         Left            =   240
         TabIndex        =   2
         Top             =   4155
         Width           =   650
      End
      Begin VB.Label LblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ҩƷ�ɹ��ƻ���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   30
         TabIndex        =   14
         Top             =   120
         Width           =   11535
      End
      Begin VB.Label Lbl�ƻ����� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ƻ�����:"
         Height          =   180
         Left            =   180
         TabIndex        =   0
         Top             =   660
         Width           =   810
      End
      Begin VB.Label Lbl������ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   300
         TabIndex        =   13
         Top             =   4500
         Width           =   540
      End
      Begin VB.Label Lbl�������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Left            =   120
         TabIndex        =   12
         Top             =   4860
         Width           =   720
      End
      Begin VB.Label Lbl����� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����"
         Height          =   180
         Left            =   4800
         TabIndex        =   11
         Top             =   4500
         Width           =   540
      End
      Begin VB.Label Lbl������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         Height          =   180
         Left            =   4560
         TabIndex        =   10
         Top             =   4860
         Width           =   720
      End
   End
   Begin MSComctlLib.ImageList imghot 
      Left            =   840
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanCard.frx":14352
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanCard.frx":1456C
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanCard.frx":14786
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanCard.frx":149A0
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanCard.frx":14BBA
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanCard.frx":14DD4
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanCard.frx":14FEE
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanCard.frx":15208
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgcold 
      Left            =   120
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanCard.frx":15422
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanCard.frx":1563C
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanCard.frx":15856
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanCard.frx":15A70
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanCard.frx":15C8A
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanCard.frx":15EA4
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanCard.frx":160BE
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanCard.frx":162D8
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   26
      Top             =   6615
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDrugPlanCard.frx":164F2
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14393
            Key             =   "STOCK"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmDrugPlanCard.frx":16D86
            Key             =   "PY"
            Object.ToolTipText     =   "ƴ��(F7)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmDrugPlanCard.frx":17288
            Key             =   "WB"
            Object.ToolTipText     =   "���(F7)"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Msf��Ӧ��ѡ�� 
      Height          =   2565
      Left            =   5850
      TabIndex        =   27
      Top             =   1890
      Visible         =   0   'False
      Width           =   4785
      _ExtentX        =   8440
      _ExtentY        =   4524
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   -2147483631
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   0
      FillStyle       =   1
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblCode 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   3240
      TabIndex        =   21
      Top             =   5880
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Menu mnuCol 
      Caption         =   "����"
      Visible         =   0   'False
      Begin VB.Menu mnuColDrug 
         Caption         =   "ҩ��(���������)"
         Index           =   0
      End
      Begin VB.Menu mnuColDrug 
         Caption         =   "ҩ��(������)"
         Index           =   1
      End
      Begin VB.Menu mnuColDrug 
         Caption         =   "ҩ��(������)"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmDrugPlanCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mint�༭״̬ As Integer             '1.������2���޸ģ�3�����գ�4���鿴��5�����ˣ�6���޸�ִ������
Private mstr���ݺ� As String                '����ĵ��ݺ�;
Private mint��¼״̬ As Integer             '1:������¼;2-������¼;3-�Ѿ�������ԭ��¼
Private mblnSuccess As Boolean              'ֻҪ��һ�ųɹ�����ΪTrue������ΪFalse
Private mblnFirst As Boolean                '��һ����ʾ
Private mblnSave As Boolean                 '�Ƿ���̺����   TURE���ɹ���
Private mblnStart As Boolean
Private mfrmMain As Form
Private mintcboIndex As Integer
Private mblnEdit As Boolean                 '�Ƿ�����޸�
Private mblnChange As Boolean               '�Ƿ���й��༭
Private mintParallelRecord As Integer       '���������󵥾ݲ���ִ�еĴ��� 1���������������2���Ѿ�ɾ���ļ�¼��3���Ѿ���˵ļ�¼
Private mintPriceUnit As Integer            'ָ�������۶��۵�λ:ȱʡΪ0-���ۼ۵�λ���ۣ���ѡΪ1-��ҩ�ⵥλ���ۣ�
Private mbln���� As Boolean                 '����ȡ���ڴ������޵�ҩƷ
Private mint���� As Integer
Private mint���� As Integer
Private mbln�ƻ����� As Boolean
Private mbln�����ǿ�� As Boolean
Private mlng�ƻ�ID As Long
Private mlng�ⷿID As Long
Private mint�ƻ����� As Integer
Private mint���Ʒ��� As Integer
Private mstr��Ӧ��ID As String
Private mbln�б굥λ As Boolean
Private mbln������ʽ As Boolean             'false �����޼ƻ����� true �����޼ƻ�����
Private Str�ڼ�  As String                  '������λ��ʾ,������λ��ʾ,������λ��ʾ
Private mstrPrivs As String                     'Ȩ��
Private mblnCheckRefresh    As Boolean      '���ʱ�Ƿ�ı�ƻ�������˵��
Private mblnClearZeroPlan  As Boolean       '�Ƿ�ɾ���ƻ�����Ϊ0�ļ�¼
Private mblnBaseMedi As Boolean             '�Ƿ��������ҩ��
Private mblnOnlyBaseMedi As Boolean         '������������ҩ��
Private mintStock As Integer                '����ҩѡ��0-ֻ��ȡ����ҩ��1-ֻ��ȡ�ǳ���ҩ��2-�������Ƿ񳣱�ҩ��
Private mblnEnter As Boolean                '�Ƿ���뵥Ԫ��
Private Const MStrCaption As String = "ҩƷ�ƻ�����"
Private mintPlanPoint As Integer            'ȫԺ�ƻ�����վ�� 0-Ҫ��վ�㣬1-����վ��
Private mstrToxicologyClass As String       '�������
Private mbln�����������ƻ� As Boolean
Private mstr��Դҩ�� As String               '��ʽ:ҩ��id1,ҩ��id2...
Private mstr��Դ�ⷿ As String               '��ʽ:ҩ��id1,ҩ��id2...
Private mstrAll��Դҩ�� As String            '������Դҩ������ʽ:ҩ��id1,ҩ��id2...
Private mstrAll��Դ�ⷿ As String            '������Դҩ������ʽ:ҩ��id1,ҩ��id2...
Private mstr�Զ���ⷿ As String            '��¼�ⷿID

Private marrFrom As Variant                   '��¼�û��ָ�������и���
Private marrInitGrid As Variant                '��¼��ʼ��������и���

Private mstrBeginDate As String
Private mstrEndDate As String
Private mstrNow As String

Private mintUnit As Integer                 '��λϵ����1-�ۼ�;2-����;3-סԺ;4-ҩ��
'�Ӳ�������ȡҩƷ�۸����������С��λ��
Private mintShowCostDigit As Integer            '�ɱ���С��λ��
Private mintShowPriceDigit As Integer           '�ۼ�С��λ��
Private mintShowNumberDigit As Integer          '����С��λ��
Private mintShowMoneyDigit As Integer           '���С��λ��

Private mstrCostFormat As String
Private mstrPriceFormat As String
Private mstrNumberFormat As String
Private mstrMoneyFormat As String

Private mdatThisMondyDate As Date
Private mint�۸���ʾ As Integer             '0-��ʾ�ɱ���;1-��ʾ�ۼ�;2-��ʾ�ɱ��ۺ��ۼ�
Private mblnViewCost As Boolean             '�鿴�ɱ��� true-���Բ鿴�ɱ��� false-�����Բ鿴�ɱ���
Private mint��Ӧ��ѡ�� As Integer           '0-ȡ�ϴ���⹩Ӧ�̣�1-ȡ��ͬ��λ
Private mint��Ӧ�̷�Χ As Integer           '0-���й�Ӧ�̣�1-�б굥λ

Private mintDrugNameShow As Integer         'ҩƷ��ʾ��0����ʾ��������ƣ�1������ʾ���룻2������ʾ����

Private mrsHisPlane As ADODB.Recordset      '��¼��ʷ�ɹ��ƻ�
Private mintLastCol As Integer              '��¼�����ʾ�������
Private mstrColumn_UnSelected As String     '��¼��Щ�б�����Ϊ����ʾ

Private mlng�����̳��� As Long                 '�������ֶγ���
Private mlngԭ���س��� As Long                 'ԭ�����ֶγ���

'=========================================================================================
Private mconIntCol��� As Integer
Private mconIntColҩ�� As Integer
Private mconIntCol��Ʒ�� As Integer
Private mconIntCol��Դ As Integer
Private mconIntCol��� As Integer
Private mconIntCol������ As Integer
Private mconIntColԭ���� As Integer
Private mconIntCol��λ As Integer
Private mconIntCol����ϵ�� As Integer
Private mconIntcolҽ������ As Integer
Private mconIntColǰ������ As Integer
Private mconIntCol�������� As Integer
Private mconIntCol������� As Integer
Private mconIntCol������� As Integer
Private mconintCol������� As Integer
Private mconintCol�������� As Integer
Private mconintCol�������� As Integer
Private mconintCol�ƻ����� As Integer
Private mconintColִ������ As Integer
Private mconintColԭִ������ As Integer
Private mconintCol�ͻ���λ As Integer
Private mconintCol�ͻ����� As Integer
Private mconintCol�ͻ���װ As Integer
Private mconintCol�ɱ��� As Integer
Private mconIntCol�ɱ���� As Integer
Private mconIntCol�ۼ� As Integer
Private mconIntCol�ۼ۽�� As Integer
Private mconintCol�ϴι�Ӧ�� As Integer
Private mconintCol˵�� As Integer
Private mconIntColҩƷ��������� As Integer
Private mconIntColҩƷ���� As Integer
Private mconIntColҩƷ���� As Integer
Private mconIntCol����ҩ�� As Integer
Private mconIntCol��׼�ĺ� As Integer
Private mconIntColS   As Integer     '������
'=========================================================================================

Private Sub ClearZeroPlan()
    Dim n As Integer
    Dim i As Integer
    
    '����ƻ���Ϊ0�ļ�¼�����ݼƻ��������������жϣ�
    If mblnClearZeroPlan = False Then Exit Sub
    With mshBill
        For n = .rows - 1 To 1 Step -1
            If n = 1 And .rows = 2 And Val(.TextMatrix(n, mconintCol�ƻ�����)) = 0 Then
                For i = 0 To .Cols - 1
                    .TextMatrix(1, i) = ""
                Next
                Exit For
            End If
            If Val(.TextMatrix(n, mconintCol�ƻ�����)) = 0 Then
                .MsfObj.RemoveItem n
            End If
        Next
    End With
End Sub

Private Sub GegReg()
    mint�۸���ʾ = Val(zlDataBase.GetPara("�۸���ʾ��ʽ", glngSys, ģ���.ҩƷ�ƻ�))
    mint��Ӧ��ѡ�� = Val(zlDataBase.GetPara("��Ӧ��Ĭ��ѡ��", glngSys, ģ���.ҩƷ�ƻ�))
    mint��Ӧ�̷�Χ = Val(zlDataBase.GetPara("��Ӧ��ѡ��Χ", glngSys, ģ���.ҩƷ�ƻ�))
End Sub


Private Sub IniHisPlaneRec()
    Set mrsHisPlane = New ADODB.Recordset
    With mrsHisPlane
        If .State = 1 Then .Close
        .Fields.Append "ҩƷID", adDouble, 18, adFldIsNullable
        .Fields.Append "�ƻ�����", adDouble, 18, adFldIsNullable
        .Fields.Append "ִ������", adDouble, 18, adFldIsNullable
        .Fields.Append "NO", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "�ƻ�����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "���Ʒ���", adLongVarChar, 50, adFldIsNullable '
        .Fields.Append "������", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "��������", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "�����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "�������", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "������", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "��������", adLongVarChar, 20, adFldIsNullable
       
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Sub LoadHisPlane(ByVal lng�ⷿID As Long, ByVal lngҩƷid As Long, ByVal lngRow As Long)
    'װ����ʷ�ɹ��ƻ���¼
    Dim rsTmp As ADODB.Recordset
    Dim intExe As Integer
    Dim j As Integer
    
    On Error GoTo errHandle
    If Not mrsHisPlane Is Nothing Then
        mrsHisPlane.Filter = "ҩƷID=" & lngҩƷid
        If Not mrsHisPlane.EOF Then Exit Sub
    End If
    
    gstrSQL = "Select B.�ƻ�����, B.ִ������, A.NO, Decode(A.�ƻ�����, 1, '�¶ȼƻ�', 2, '���ȼƻ�', 3, '��ȼƻ�', '�ܼƻ�') As �ƻ�����, " & _
        " Decode(A.���Ʒ���, 1, '����ͬ�����β��շ�', 2, '�ٽ��ڼ�ƽ�����շ�', 3, 'ҩƷ����������շ�', 4, 'ҩƷ�����������շ�', '�Զ���������շ�') As ���Ʒ���, A.������, A.��������, " & _
        " A.����� , A.�������, A.������, A.�������� " & _
        " From ҩƷ�ɹ��ƻ� A, ҩƷ�ƻ����� B " & _
        " Where A.Id = B.�ƻ�id And A.����� Is Not Null And B.�ƻ�����>0 And A.�ⷿid + 0 = [1] And B.ҩƷid = [2] "
    If Trim(txtNo.Caption) <> "" Then
        gstrSQL = gstrSQL & " And A.NO <> [3]"
    End If
    gstrSQL = gstrSQL & " Order By No Desc "
    Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, "LoadHisPlane", lng�ⷿID, lngҩƷid, Trim(txtNo.Caption))
    
    With rsTmp
        If Not .EOF Then
            If NVL(!ִ������, 0) = 0 Then
                intExe = 1
            ElseIf NVL(!ִ������, 0) >= NVL(!�ƻ�����, 0) Then
                intExe = 2
            Else
                intExe = 3
            End If
        End If

        Do While Not .EOF
            mrsHisPlane.AddNew
            
            mrsHisPlane!ҩƷid = lngҩƷid
            mrsHisPlane!�ƻ����� = zlStr.FormatEx(NVL(!�ƻ�����, 0) / Val(mshBill.TextMatrix(lngRow, mconIntCol����ϵ��)), mintShowNumberDigit, , True)
            mrsHisPlane!ִ������ = zlStr.FormatEx(NVL(!ִ������, 0) / Val(mshBill.TextMatrix(lngRow, mconIntCol����ϵ��)), mintShowNumberDigit, , True)
            mrsHisPlane!NO = !NO
            mrsHisPlane!�ƻ����� = !�ƻ�����
            mrsHisPlane!���Ʒ��� = !���Ʒ���
            mrsHisPlane!������ = NVL(!������, "")
            mrsHisPlane!�������� = IIf(IsNull(!��������), "", Format(!��������, "YYYY-MM-DD"))
            mrsHisPlane!����� = NVL(!�����, "")
            mrsHisPlane!������� = IIf(IsNull(!�������), "", Format(!�������, "YYYY-MM-DD"))
            mrsHisPlane!������ = NVL(!������, "")
            mrsHisPlane!�������� = IIf(IsNull(!��������), "", Format(!��������, "YYYY-MM-DD"))
            
            .MoveNext
        Loop
    End With
    
    '�����ϴμƻ��������Ե�ǰҩƷ��ɫ
    If intExe > 0 Then
        mblnEnter = False
        With mshBill
            .Row = lngRow
            .Col = mconIntColҩ��
            j = .ColData(mconIntColҩ��)
            If .ColData(mconIntColҩ��) = 5 Then .ColData(mconIntColҩ��) = 0
            
            If intExe = 1 Then
                'δִ��
                .MsfObj.CellForeColor = vbRed
            ElseIf intExe = 2 Then
                '��ȫִ��
                .MsfObj.CellForeColor = vbBlue
            Else
                '����ִ��
                .MsfObj.CellForeColor = &H8080FF
            End If
    
            .ColData(mconIntColҩ��) = j
        End With
        mblnEnter = True
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ResizeHisPlane()
    On Error Resume Next
    
    With picHis
        If Val(picHscSend.Tag) = 1 Then
            .Height = 1755
        Else
            .Height = 300
        End If
        
        .Top = lblPurchasePrice.Top - .Height - 100
        .Left = mshBill.Left
        .Width = mshBill.Width
    End With
    
    With picStock
        If picHis.Visible Then
            .Top = picHis.Top - .Height - 60
        Else
            .Top = lblPurchasePrice.Top - .Height - 60
        End If
        
        .Left = mshBill.Left
        .Width = mshBill.Width
    End With
    
    With picHeadStock
        .Width = picStock.Width
    End With

    With vsfStock
        .Width = picStock.Width
        .Height = picStock.Height - 330
    End With
    
    With picHscSend
        .Width = picHis.Width
    End With
    
    With vsfHisPlane
        .Width = picHis.Width
    End With

    With pic�ⷿ
        .Top = picStock.Top + cmd�ⷿ.Height
        .Left = cmd�ⷿ.Left + 160
    End With
    
    With mshBill
        If picHis.Visible And picStock.Visible Then
            .Height = picStock.Top - .Top - 60
        ElseIf picHis.Visible And Not picStock.Visible Then
            .Height = picHis.Top - .Top - 60
        ElseIf Not picHis.Visible And picStock.Visible Then
            .Height = picStock.Top - .Top - 60
        Else
            .Height = lblPurchasePrice.Top - .Top - 60
        End If
    End With
End Sub

Public Sub ShowCard(FrmMain As Form, ByVal str���ݺ� As String, _
        ByVal int�༭״̬ As Integer, Optional blnSuccess As Boolean = False, Optional lng�ⷿID As Long)
    mblnSave = False
    mblnSuccess = False
    mstr���ݺ� = str���ݺ�
    mint�༭״̬ = int�༭״̬
    mintParallelRecord = 1
    mlng�ⷿID = lng�ⷿID
    mstrPrivs = GetPrivFunc(glngSys, 1330)

    mblnSuccess = blnSuccess
    mblnChange = False
    mblnFirst = True

    Set mfrmMain = FrmMain

    If mint�༭״̬ = 1 Then
        mblnEdit = True
    ElseIf mint�༭״̬ = 2 Then
        mblnEdit = True
    ElseIf mint�༭״̬ = 3 Then
        mblnEdit = False
        CmdSave.Caption = "���(&V)"
    ElseIf mint�༭״̬ = 4 Then
        mblnEdit = False
        CmdSave.Caption = "��ӡ(&P)"
        If Not zlStr.IsHavePrivs(mstrPrivs, "�ɹ��ƻ���ӡ") Then
            CmdSave.Visible = False
        Else
            CmdSave.Visible = True
        End If

    End If

    Me.Show vbModal, FrmMain
    blnSuccess = mblnSuccess
    str���ݺ� = mstr���ݺ�

End Sub

Private Sub ShowHisPlane(ByVal lngRow As Long, ByVal lngҩƷid As Long)
    '��ʾ��ʷ�ɹ��ƻ�
    
    vsfHisPlane.rows = 1
    vsfHisPlane.rows = 2
    
    lblDiag.Caption = "���ڲɹ��ƻ�ִ�����"
    
    If mrsHisPlane Is Nothing Then Exit Sub
    
    mrsHisPlane.Filter = "ҩƷID=" & lngҩƷid
    If mrsHisPlane.EOF Then Exit Sub
    
    lblDiag.Caption = lblDiag.Caption & "(" & mrsHisPlane.RecordCount & ")"
    
    With vsfHisPlane
        .Redraw = flexRDNone
        Do While Not mrsHisPlane.EOF
            .TextMatrix(.rows - 1, .ColIndex("�ƻ�����")) = zlStr.FormatEx(mrsHisPlane!�ƻ�����, mintShowNumberDigit, , True)
            .TextMatrix(.rows - 1, .ColIndex("ִ������")) = zlStr.FormatEx(NVL(mrsHisPlane!ִ������, 0), mintShowNumberDigit, , True)
            .TextMatrix(.rows - 1, .ColIndex("NO")) = mrsHisPlane!NO
            .TextMatrix(.rows - 1, .ColIndex("�ƻ�����")) = mrsHisPlane!�ƻ�����
            .TextMatrix(.rows - 1, .ColIndex("���Ʒ���")) = mrsHisPlane!���Ʒ���
            .TextMatrix(.rows - 1, .ColIndex("������")) = mrsHisPlane!������
            .TextMatrix(.rows - 1, .ColIndex("��������")) = mrsHisPlane!��������
            .TextMatrix(.rows - 1, .ColIndex("�����")) = mrsHisPlane!�����
            .TextMatrix(.rows - 1, .ColIndex("�������")) = mrsHisPlane!�������
            .TextMatrix(.rows - 1, .ColIndex("������")) = mrsHisPlane!������
            .TextMatrix(.rows - 1, .ColIndex("��������")) = mrsHisPlane!��������
            
            If NVL(mrsHisPlane!ִ������, 0) = 0 Then
                'δִ��
                .Cell(flexcpForeColor, .rows - 1, 0, .rows - 1, .Cols - 1) = vbRed
            ElseIf NVL(mrsHisPlane!ִ������, 0) >= mrsHisPlane!�ƻ����� Then
                '��ȫִ��
                .Cell(flexcpForeColor, .rows - 1, 0, .rows - 1, .Cols - 1) = vbBlue
            Else
                '����ִ��
                .Cell(flexcpForeColor, .rows - 1, 0, .rows - 1, .Cols - 1) = &H8080FF
            End If
            
            .Cell(flexcpFontBold, .rows - 1, .ColIndex("�ƻ�����")) = True
            .Cell(flexcpFontBold, .rows - 1, .ColIndex("ִ������")) = True
            
            .rows = .rows + 1
                        
            If chkMore.Value = 0 And .rows > 4 Then
                .Redraw = flexRDDirect
                Exit Sub
            End If
            
            mrsHisPlane.MoveNext
        Loop
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub chkMore_Click()
    If mshBill.TextMatrix(mshBill.Row, 0) <> "" Then
        Call ShowHisPlane(mshBill.Row, Val(mshBill.TextMatrix(mshBill.Row, 0)))
    End If
End Sub

Private Sub chk�Ƿ���ʾ������_Click()
    With picStock
        .Visible = Not .Visible
    End With
    
    Call ResizeHisPlane
End Sub
Private Sub chk��Դ�ⷿ_Click()
    If chk��Դ�ⷿ.Value = 1 Then
        If chk���пⷿ.Value = 1 Then chk���пⷿ.Value = 0
    End If
    Call ��ʾ���
End Sub
Private Sub chk��Դҩ��_Click()
    If chk��Դҩ��.Value = 1 Then
        If chk���пⷿ.Value = 1 Then chk���пⷿ.Value = 0
    End If
    Call ��ʾ���
End Sub
Private Sub chk���пⷿ_Click()
    If chk���пⷿ.Value = 1 Then
        If chk��Դҩ��.Value = 1 Then chk��Դҩ��.Value = 0
        If chk��Դ�ⷿ.Value = 1 Then chk��Դ�ⷿ.Value = 0
        If mstr�Զ���ⷿ <> "" Then mstr�Զ���ⷿ = ""
    End If
    Call ��ʾ���
End Sub
Private Sub chk���ؽ��ڲɹ��ƻ�_Click()
    With picHis
        .Visible = Not .Visible
    End With
    
    Call ResizeHisPlane
End Sub

Private Sub cmdCancel_Click()
    If MsgBox("��ȷ��Ҫ�˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Unload Me
End Sub

'����
Private Sub cmdFind_Click()

    If lblCode.Visible = False Then
        lblCode.Visible = True
        txtCode.Visible = True
        txtCode.SetFocus
    Else
        FindRow mshBill, mconIntColҩƷ���������, txtCode.Text, True
        lblCode.Visible = False
        txtCode.Visible = False
    End If
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub cmdOk_Click()
    Dim intItem As Integer
    Dim intItems As Integer
    'ȡ�ô洢�ⷿ
    mstr�Զ���ⷿ = ""
    intItems = Me.lvw�洢�ⷿ.ListItems.count
    For intItem = 1 To intItems
        If lvw�洢�ⷿ.ListItems(intItem).Checked Then
            mstr�Զ���ⷿ = mstr�Զ���ⷿ & "," & Mid(lvw�洢�ⷿ.ListItems(intItem).Key, 2)
        End If
    Next
    mstr�Զ���ⷿ = Mid(mstr�Զ���ⷿ, 2)
    
    With pic�ⷿ
        .Visible = False
    End With
    
    If mstr�Զ���ⷿ <> "" Then
        If chk���пⷿ.Value = 1 Then chk���пⷿ.Value = 0
    End If
    Call ��ʾ���
End Sub

Private Sub cmdȡ��_Click()
    With pic�ⷿ
        .Visible = False
    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 70 Or KeyCode = 102 Then
        If Shift = vbCtrlMask Then   'Ctrl+F
            cmdFind_Click
        End If
    ElseIf KeyCode = vbKeyF3 Then
        FindRow mshBill, mconIntColҩ��, txtCode.Text, False
    ElseIf KeyCode = vbKeyEscape Then
        If Msf��Ӧ��ѡ��.Visible Then
            Msf��Ӧ��ѡ��.ZOrder 1
            Msf��Ӧ��ѡ��.Visible = False
            Exit Sub
        End If
'        Call cmdCancel_Click
    ElseIf KeyCode = vbKeyF7 Then
        If staThis.Panels("PY").Bevel = sbrRaised Then
            Logogram staThis, 0
        Else
            Logogram staThis, 1
        End If
    End If
End Sub


Private Sub CmdSave_Click()
    Dim blnSuccess As Boolean
    
    If mint�༭״̬ = 4 Then    '�鿴
        '��ӡ
        Call FrmBillPrint.ShowME(Me, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1330", "zl8_bill_1330"), 0, 0, 1330, "ҩƷ�ɹ��ƻ���", txtNo.Tag, mint�۸���ʾ)
        '�˳�
        Unload Me
        Exit Sub
    End If

    If mint�༭״̬ = 3 Then        '���
        If mblnCheckRefresh Then
            If Not SaveCard Then
                Exit Sub
            End If
        End If
        If SaveCheck = True Then
            If Val(zlDataBase.GetPara("��˴�ӡ", glngSys, ģ���.ҩƷ�ƻ�)) = 1 Then
                '��ӡ
                If zlStr.IsHavePrivs(mstrPrivs, "�ɹ��ƻ���ӡ") Then
                    ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1330", "zl8_bill_1330"), Me, "���ݱ��=" & txtNo.Tag, IIf(mint�۸���ʾ = 0, "ReportFormat=1", IIf(mint�۸���ʾ = 1, "ReportFormat=2", "ReportFormat=3")), 2
                End If
            End If
            Unload Me
        End If
        Exit Sub
    End If
    
    If mint�༭״̬ = 5 Then        '����
        If SaveReCheck = True Then
            Unload Me
        End If
        Exit Sub
    End If
    
    If mint�༭״̬ = 6 Then        '�޸�ִ������
        If SaveExeAmount = True Then
            Unload Me
        End If
        Exit Sub
    End If

    If ValidData = False Then Exit Sub
    blnSuccess = SaveCard

    If blnSuccess = True Then

        If Val(zlDataBase.GetPara("���̴�ӡ", glngSys, ģ���.ҩƷ�ƻ�)) = 1 Then
            '��ӡ
            If zlStr.IsHavePrivs(mstrPrivs, "�ɹ��ƻ���ӡ") Then
                ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1330", "zl8_bill_1330"), Me, "���ݱ��=" & txtNo.Tag, IIf(mint�۸���ʾ = 0, "ReportFormat=1", IIf(mint�۸���ʾ = 1, "ReportFormat=2", "ReportFormat=3")), 2
            End If
        End If
        If mint�༭״̬ = 2 Then   '�޸�
            Unload Me
            Exit Sub
        End If
    Else
        Exit Sub
    End If

    mblnSave = False
    mblnEdit = True
    mshBill.ClearBill
    txtժҪ.Text = ""
    mblnChange = False
    If txtNo.Tag <> "" Then Me.staThis.Panels(2).Text = "��һ�ŵ��ݵ�NO�ţ�" & txtNo.Tag
End Sub

Private Sub Form_Activate()
    Dim intMonth As Integer
    Dim datCurrDate As Date
    Dim intWeekDay As Integer
    Const intMonday As Integer = vbMonday
    Dim intCountDay As Integer
    
    If mblnFirst = False Then Exit Sub

    If Not zlStr.IsHavePrivs(mstrPrivs, "���пⷿ") Then
        chk�Ƿ���ʾ������.Visible = False
    Else
        chk�Ƿ���ʾ������.Visible = True
        Call Init�洢�ⷿ
    End If
    
    '��ʼ�����뷽ʽ
    If (mint�༭״̬ = 1 Or mint�༭״̬ = 2) And gbytSimpleCodeTrans = 1 Then
        staThis.Panels("PY").Visible = True
        staThis.Panels("WB").Visible = True
        gint���뷽ʽ = Val(zlDataBase.GetPara("���뷽ʽ", , , 0))    'Ĭ��ƴ������
        Logogram staThis, gint���뷽ʽ
    Else
        staThis.Panels("PY").Visible = False
        staThis.Panels("WB").Visible = False
    End If
    
    mblnFirst = False
    If mint�༭״̬ = 1 Then
        Dim str��;ID As String, str���ͱ��� As String
        Dim lng�ⷿID As Long, int�ƻ����� As Integer, int���Ʒ��� As Integer, bln������ʽ As Boolean
        Dim strToxicologyClass As String
        
        If frmDrugPlanCondition.GetCondition(mfrmMain, str��;ID, str���ͱ���, lng�ⷿID, int�ƻ�����, _
                int���Ʒ���, mbln����, mint����, mint����, mbln�ƻ�����, _
                mstr��Ӧ��ID, mbln�б굥λ, mstrBeginDate, mstrEndDate, mbln�����ǿ��, _
                mblnClearZeroPlan, mblnBaseMedi, mintStock, bln������ʽ, mblnOnlyBaseMedi, _
                strToxicologyClass, mbln�����������ƻ�, mstr��Դҩ��, mstr��Դ�ⷿ, mstrAll��Դҩ��, mstrAll��Դ�ⷿ) = True Then
            mlng�ⷿID = lng�ⷿID
            mint�ƻ����� = int�ƻ�����
            mint���Ʒ��� = int���Ʒ���
            mbln������ʽ = bln������ʽ
            mstrToxicologyClass = strToxicologyClass
            Select Case mint�ƻ�����
                Case 1       '�¼ƻ�
                    Str�ڼ� = Format(DateAdd("m", 1, Sys.Currentdate), "yyyyMM")
                    LblTitle.Caption = GetUnitName & "(" & Mid(Str�ڼ�, 1, 4) & "��" & Right(Str�ڼ�, 2) & "��" & ") " & LblTitle.Tag & "�ɹ��ƻ�"
                    
                    mshBill.TextMatrix(0, mconintCol��������) = "��������"
                    mshBill.TextMatrix(0, mconintCol��������) = "��������"
                Case 2       '���ƻ�
                    intMonth = Month(DateAdd("Q", 1, Sys.Currentdate))
                    Str�ڼ� = Format(DateAdd("Q", 1, Sys.Currentdate), "yyyy") & IIf(intMonth <= 3, 1, IIf(intMonth >= 10, 4, IIf(intMonth <= 9 And intMonth >= 7, 3, 2)))
                    LblTitle.Caption = GetUnitName & "(" & Mid(Str�ڼ�, 1, 4) & "��" & Right(Str�ڼ�, 1) & "��" & ")" & LblTitle.Tag & "�ɹ��ƻ�"
                    
                    mshBill.TextMatrix(0, mconintCol��������) = "�ϼ�������"
                    mshBill.TextMatrix(0, mconintCol��������) = "����������"
                Case 3      '��ƻ�
                    Str�ڼ� = Format(DateAdd("yyyy", 1, Sys.Currentdate), "yyyy")
                    LblTitle.Caption = GetUnitName & "(" & Str�ڼ� & "��" & ")" & LblTitle.Tag & "�ɹ��ƻ�"
                    
                    mshBill.TextMatrix(0, mconintCol��������) = "��������"
                    mshBill.TextMatrix(0, mconintCol��������) = "��������"
                Case 4      '�ܼƻ�
                    datCurrDate = Sys.Currentdate
                    intWeekDay = Weekday(datCurrDate)
                    If intWeekDay = 1 Then
                        intCountDay = -6
                    Else
                        intCountDay = intMonday - intWeekDay
                    End If
                    mdatThisMondyDate = DateAdd("d", intCountDay, datCurrDate)
                    Str�ڼ� = Format(DateAdd("d", 7, mdatThisMondyDate), "yyyyMMDD")
                    LblTitle.Caption = GetUnitName & "(" & Mid(Str�ڼ�, 1, 4) & "��" & Mid(Str�ڼ�, 5, 2) & "��" & Right(Str�ڼ�, 2) & "��" & ")" & LblTitle.Tag & "�ɹ��ƻ�"
                    
                    mshBill.TextMatrix(0, mconintCol��������) = "��������"
                    mshBill.TextMatrix(0, mconintCol��������) = "��������"
            End Select
            
            If mint���Ʒ��� = 5 Then
                '�Զ���������Ʒ�
                mshBill.TextMatrix(0, mconIntColǰ������) = "��������"
                mshBill.TextMatrix(0, mconIntCol��������) = "��������"
                mshBill.TextMatrix(0, mconintCol��������) = "��������"
                mshBill.TextMatrix(0, mconintCol��������) = "��������"
            End If

            ReFreshALLDrug str��;ID, str���ͱ���, lng�ⷿID, int�ƻ�����, int���Ʒ���, bln������ʽ
        Else
            Unload Me
            Exit Sub
        End If
        If mshBill.Visible = True Then
            mshBill.SetFocus
        End If
    Else
        mblnChange = False
        Select Case mintParallelRecord
            Case 1
                '����
            Case 2
                '�����ѱ�ɾ��
                MsgBox "�õ����ѱ�ɾ�������飡", vbOKOnly, gstrSysName
                Unload Me
                Exit Sub
            Case 3
                '�޸ĵĵ����ѱ����
                MsgBox "�õ����ѱ���������ˣ����飡", vbOKOnly, gstrSysName
                Unload Me
                Exit Sub
        End Select
    End If

    mblnStart = True
End Sub

Private Sub ReFreshALLDrug(ByVal str��;ID, ByVal str���� As String, _
    ByVal lng�ⷿID As Long, ByVal int�ƻ����� As Integer, ByVal int���Ʒ��� As Integer, ByVal bln������ʽ As Boolean)
        '---------------------------------------------------
        '--����:������ҩƷ���мƻ�����
        '--����:
        '---------------------------------------------------
    Dim rsAllDrug As New ADODB.Recordset
    Dim rspurchase As New ADODB.Recordset
    Dim intRecord As Long
    Dim intRow  As Long
    Dim rsData As ADODB.Recordset
    Dim lng��Ӧ��ID As Long
    Dim str��Ӧ�� As String
    Dim str���� As String
    Dim strԭ���� As String
    Dim dbl������� As Double
    Dim blnOK As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim rs��ͬ��λ As ADODB.Recordset
    Dim strҩ�� As String
    Dim str���ʹ� As String
    Dim str�ͻ���λ As String
    Dim dbl�ͻ���װ As Double
    
    On Error GoTo errHandle
    Me.Refresh
    Me.MousePointer = vbHourglass
    mshBill.Redraw = False
    staThis.Panels(2).Text = "���ڼ���"
    CmdSave.Enabled = False
    CmdCancel.Enabled = False
    Pic����.Enabled = False

    str���ʹ� = Replace(str����, "'", "")
        
    'ȡָ��������ҩƷ��Ϣ
    gstrSQL = "" & _
         " SELECT DISTINCT A.ҩƷID,'[' || F.���� || ']' As ҩƷ����, F.���� As ͨ����, B.���� As ��Ʒ��,A.ҩƷ��Դ,f.���,a.����ҩ��," & _
         " Decode(" & mintUnit & ", 1, f.���㵥λ, 2, a.���ﵥλ, 3, a.סԺ��λ, a.ҩ�ⵥλ) As ��λ," & _
         " DECODE(A.�ɱ���,NULL,NVL(A.ָ��������,0),0,NVL(A.ָ��������,0),NVL(A.�ɱ���,0)) AS ����,F.����,A.ԭ����," & _
         " Decode(" & mintUnit & ", 1, 1, 2, a.�����װ, 3, a.סԺ��װ, a.ҩ���װ) As ����ϵ��," & _
         " Nvl(G.�ּ�, 0) �ۼ�,a.�ϴ��ۼ�,Nvl(F.�Ƿ���,0) �Ƿ���, a.�ͻ���λ, a.�ͻ���װ,f.��������, a.�ϴι�Ӧ��id As ��Ӧ��id, d.���� As ��Ӧ��,nvl(a.�ϴ���׼�ĺ�,a.��׼�ĺ�) as ��׼�ĺ� " & _
         " FROM ҩƷ��� A,�շ���Ŀ���� B,������ĿĿ¼ C,���Ʒ���Ŀ¼ L,�շ���ĿĿ¼ F,ҩƷ���� T, �շѼ�Ŀ G, ��Ӧ�� D "
    
    gstrSQL = gstrSQL & " WHERE A.ҩƷID=F.ID And A.ҩ��ID=C.ID and A.ҩ��ID=T.ҩ��ID And C.����ID=L.ID and L.���� in (1,2,3)" & _
         " And A.ҩƷID = B.�շ�ϸĿID(+) And B.����(+)=3 " & _
         " AND (F.����ʱ��>=TO_DATE('3000-01-01 00:00:00','YYYY-MM-DD HH24:MI:SS') OR F.����ʱ�� IS NULL)" & _
         " And A.ҩƷid = G.�շ�ϸĿid And (G.��ֹ���� Is Null Or Sysdate Between G.ִ������ And Nvl(G.��ֹ����, To_Date('3000-01-01', 'yyyy-MM-dd'))) " & _
         GetPriceClassString("G") & " And a.�ϴι�Ӧ��id = d.Id(+) "
    
    '����
    If mstrToxicologyClass <> "" Then
            gstrSQL = gstrSQL & " And " & mstrToxicologyClass
    End If
    
    If mintStock <> 2 Then
        gstrSQL = gstrSQL & " And Nvl(A.�Ƿ񳣱�, 0) = [4] "
    End If
    
    If mblnOnlyBaseMedi = True Then
        gstrSQL = gstrSQL & " and a.����ҩ�� is not null "
    End If
    
    If mblnBaseMedi = False And mblnOnlyBaseMedi = False Then
        gstrSQL = gstrSQL & " And A.����ҩ�� Is Null "
    End If
    
    If str��;ID = "" Then
        gstrSQL = gstrSQL & " And L.ID Is NULL "
    ElseIf str��;ID <> "���з���" Then
        gstrSQL = gstrSQL & " And L.ID in (select * from Table(Cast(f_Num2List([2]) As zlTools.t_NumList))) "
    End If
    
    If str���ʹ� <> "" Then
        gstrSQL = gstrSQL & " And T.ҩƷ���� in (select * from Table(Cast(f_Str2list([3]) As zlTools.t_Strlist))) "
    End If

    If lng�ⷿID = 0 Then
        '�����ȫ�ⷿ��ȡ���пⷿ��棬����ҩƷ�����ȡ�ϴι�Ӧ�̺��ϴβ���
        gstrSQL = "( " & gstrSQL & ") D," & _
                  " (Select A.ҩƷID, C.ID As ��Ӧ��ID,C.���� As ��Ӧ��, B.�ϴβ���,B.ԭ����, A.�������, A.ƽ���ۼ� " & _
                  " From (Select ҩƷid, Sum(ʵ������) As �������, " & _
                  " Decode(Sign(Sum(ʵ������)), 1, Decode(Sign(Sum(ʵ�ʽ��)), 1, Sum(ʵ�ʽ��), 0) / Sum(ʵ������), 0) ƽ���ۼ� " & _
                  " From ҩƷ��� " & _
                  " Where ���� = 1 " & _
                  " Group By ҩƷid) A, " & _
                  " ҩƷ��� B, " & _
                  " (SELECT ID,���� FROM ��Ӧ�� WHERE SUBSTR(����,1,1)=1) C " & _
                  " Where A.ҩƷid = B.ҩƷid And B.�ϴι�Ӧ��id = C.ID(+)) E "
    
    Else
        'ȡ�����������������εĹ�Ӧ�̣��ϴβ���
        gstrSQL = "( " & gstrSQL & ") D," & _
                  " (Select DISTINCT A.ҩƷID, C.ID As ��Ӧ��ID,C.���� As ��Ӧ��, B.�ϴβ���, B.ԭ����, A.�������, A.ƽ���ۼ� " & _
                  " From (Select ҩƷid, Sum(ʵ������) As �������, " & _
                  " Decode(Sign(Sum(ʵ������)), 1, Decode(Sign(Sum(ʵ�ʽ��)), 1, Sum(ʵ�ʽ��), 0) / Sum(ʵ������), 0) ƽ���ۼ� " & _
                  " From ҩƷ��� " & _
                  " Where ���� = 1 "
        If mstr��Դ�ⷿ <> "" Then
            gstrSQL = gstrSQL & " And �ⷿid In(select * from Table(Cast(f_Num2List([5]) As zlTools.t_NumList)))"
        Else
            gstrSQL = gstrSQL & " AND �ⷿID=[1] "
        End If
        
        gstrSQL = gstrSQL & " Group By ҩƷid) A, " & _
                  " (Select ҩƷid, max(�ϴι�Ӧ��id) as �ϴι�Ӧ��id, max(�ϴβ���) as �ϴβ���, max(ԭ����) as ԭ����  From ҩƷ��� " & _
                  " Where ���� = 1 AND (�ⷿID=[1] or �ⷿid In(select * from Table(Cast(f_Num2List([5]) As zlTools.t_NumList)))) " & _
                  " And (ҩƷid,Nvl(����, 0)) In " & _
                  " (Select ҩƷid,Nvl(Max(Nvl(����, 0)), 0) ���� From ҩƷ��� Where ���� = 1 AND (�ⷿID=[1] or �ⷿid In(select * from Table(Cast(f_Num2List([5]) As zlTools.t_NumList))))  Group By ҩƷid ) Group By ҩƷid) B, " & _
                  " (SELECT ID,���� FROM ��Ӧ�� WHERE SUBSTR(����,1,1)=1) C " & _
                  " Where A.ҩƷid = B.ҩƷid And B.�ϴι�Ӧ��id = C.ID(+)) E, " & _
                  " (Select distinct �շ�ϸĿid From �շ�ִ�п��� Where ִ�п���id = [1]) Z "
                  
    End If
    
    '������ȡҩƷ�����޶��SQL
    If mbln���� Then
        gstrSQL = gstrSQL & _
                ",(Select ҩƷID,sum(����) ����  " & _
                "     From ҩƷ�����޶�  " & _
                "" & IIf(lng�ⷿID = 0, "", " Where �ⷿID=[1]") & _
                "     Group By ҩƷID)     F"
    End If

    '�������У�����������ȡҩƷ�����޶�.���ޣ�
    gstrSQL = "SELECT d.ҩƷid, d.ҩƷ����, d.ͨ����, d.��Ʒ��, d.ҩƷ��Դ,d.���, " _
            & "DECODE (e.�ϴβ���, NULL, d.����, e.�ϴβ���) AS ����," _
            & "DECODE (e.ԭ����, NULL, d.ԭ����, e.ԭ����) AS ԭ����," _
            & "d.��λ,nvl(e.�������,0)/d.����ϵ�� as �������" & IIf(mbln����, ",nvl(F.����,0)/d.����ϵ�� as ����", "") & " , d.���� as ���� ,Nvl(e.��Ӧ��id, d.��Ӧ��id) As ��Ӧ��id, Nvl(e.��Ӧ��, d.��Ӧ��) As ��Ӧ��,d.����ϵ��, " _
            & " Decode(D.�Ƿ���, 0, D.�ۼ�, Decode(nvl(d.�ϴ��ۼ�,0), 0, Decode(Nvl(E.ƽ���ۼ�, 0), 0, D.�ۼ�, E.ƽ���ۼ�), d.�ϴ��ۼ�)) �ۼ�,d.�ͻ���λ,d.�ͻ���װ,d.��������,d.����ҩ��,d.��׼�ĺ� from " _
            & gstrSQL _
            & " WHERE d.ҩƷid = e.ҩƷid (+) "
    If lng�ⷿID <> 0 Then
        gstrSQL = gstrSQL & " And d.ҩƷid = z.�շ�ϸĿid "
    End If
    
    If mbln���� Then
        '���������ж�
        '���ϴ����޶���жϣ����ڴ����޶��ҩƷ����ȡ�������ɹ��ƻ�
        gstrSQL = gstrSQL & " And d.ҩƷID=F.ҩƷID(+)"
        gstrSQL = "Select * From (" & gstrSQL & ") Where (�������<���� and ����<>0)"
    End If
    gstrSQL = gstrSQL & " Order by ҩƷ����"

    If rsAllDrug.State = 1 Then rsAllDrug.Close
    
    Set rsAllDrug = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lng�ⷿID, str��;ID, str���ʹ�, mintStock, mstr��Դ�ⷿ)

    With rsAllDrug
        intRecord = .RecordCount

        If intRecord = 0 Then
            mshBill.Redraw = True
            Me.MousePointer = vbDefault
            CmdSave.Enabled = True
            CmdCancel.Enabled = True
            Pic����.Enabled = True
            Me.staThis.Panels(2).Text = ""
            Exit Sub
        End If
        .MoveFirst
        Me.Refresh
        DoEvents
        Do While Not .EOF
            dbl������� = IIf(IsNull(!�������), 0, !�������)
            lng��Ӧ��ID = NVL(!��Ӧ��id, 0)
            str��Ӧ�� = IIf(IsNull(!��Ӧ��), "", !��Ӧ��)
            str���� = IIf(IsNull(!����), "", !����)
            strԭ���� = IIf(IsNull(!ԭ����), "", !ԭ����)
            
            blnOK = True
            
            '����޿�棬���ҩƷ�����ȡ��Ӧ�̣��ϴβ���
            If IIf(IsNull(!�������), 0, !�������) = 0 Then
                If mstr��Ӧ��ID = "" Then
                    gstrSQL = "Select B.id ��Ӧ��ID, B.���� ��Ӧ��, C.�ϴβ���, C.ԭ����, 0 ������� from " & _
                          " (Select id,���� From ��Ӧ�� Where Substr(����, 1, 1) = 1)  B, ҩƷ��� C " & _
                          " Where C.�ϴι�Ӧ��id = B.ID(+) And ҩƷid = [1] "
                Else
                    gstrSQL = "Select B.id ��Ӧ��ID, B.���� ��Ӧ��, C.�ϴβ���, C.ԭ����, 0 ������� from " & _
                          " (Select A.id,A.���� From ��Ӧ�� A Where Substr(A.����, 1, 1) = 1 And A.id in (select * from Table(Cast(f_Num2List([2]) As zlTools.t_NumList))))  B, ҩƷ��� C " & _
                          " Where C.�ϴι�Ӧ��id = B.ID And ҩƷid = [1] "
                End If
                
                Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[ȡ�ϴι�Ӧ�̼�������Ϣ]", Val(rsAllDrug!ҩƷid), mstr��Ӧ��ID)
                If rsData.RecordCount > 0 Then
                    blnOK = True
                    lng��Ӧ��ID = NVL(rsData!��Ӧ��id, 0)
                    str���� = IIf(IsNull(rsData!�ϴβ���), "", rsData!�ϴβ���)
                    strԭ���� = IIf(IsNull(rsData!ԭ����), "", rsData!ԭ����)
                    str��Ӧ�� = IIf(IsNull(rsData!��Ӧ��), "", rsData!��Ӧ��)
                    dbl������� = IIf(IsNull(rsData!�������), 0, rsData!�������)
                Else
                    blnOK = False
                End If
            End If
            If mstr��Ӧ��ID <> "" Then
                If InStr("," & mstr��Ӧ��ID & ",", "," & lng��Ӧ��ID & ",") = 0 Then
                    If mbln�б굥λ Then
                        gstrSQL = "Select b.���� from ҩƷ�б굥λ a,��Ӧ�� b where  a.ҩƷID=[1] and   a.��λid=b.id and a.��λID in (select * from Table(Cast(f_Num2List([2]) As zlTools.t_NumList)))"
                        Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, Val(rsAllDrug!ҩƷid), mstr��Ӧ��ID)
                        blnOK = (rsTmp.RecordCount > 0)
                        If blnOK = True Then str��Ӧ�� = rsTmp!����
                    Else
                        blnOK = False
                    End If
                End If
            End If
            If blnOK Then
                intRow = intRow + 1
                mshBill.TextMatrix(intRow, 0) = !ҩƷid

                If gintҩƷ������ʾ = 0 Or gintҩƷ������ʾ = 2 Then
                    strҩ�� = !ͨ����
                Else
                    strҩ�� = IIf(IsNull(!��Ʒ��), !ͨ����, !��Ʒ��)
                End If
                
                mshBill.TextMatrix(intRow, mconIntColҩƷ���������) = !ҩƷ���� & strҩ��
                mshBill.TextMatrix(intRow, mconIntColҩƷ����) = !ҩƷ����
                mshBill.TextMatrix(intRow, mconIntColҩƷ����) = strҩ��
                
                If mintDrugNameShow = 1 Then
                    mshBill.TextMatrix(intRow, mconIntColҩ��) = mshBill.TextMatrix(intRow, mconIntColҩƷ����)
                ElseIf mintDrugNameShow = 2 Then
                    mshBill.TextMatrix(intRow, mconIntColҩ��) = mshBill.TextMatrix(intRow, mconIntColҩƷ����)
                Else
                    mshBill.TextMatrix(intRow, mconIntColҩ��) = mshBill.TextMatrix(intRow, mconIntColҩƷ���������)
                End If
                
                mshBill.TextMatrix(intRow, mconIntCol��Ʒ��) = IIf(IsNull(!��Ʒ��), "", !��Ʒ��)

                mshBill.TextMatrix(intRow, mconIntCol��Դ) = NVL(!ҩƷ��Դ)
                mshBill.TextMatrix(intRow, mconIntCol���) = IIf(IsNull(!���), "", !���)
                mshBill.TextMatrix(intRow, mconIntCol��λ) = IIf(IsNull(!��λ), "", !��λ)
                mshBill.TextMatrix(intRow, mconIntcolҽ������) = IIf(IsNull(!��������), "", !��������)
                mshBill.TextMatrix(intRow, mconIntCol������) = str����
                mshBill.TextMatrix(intRow, mconIntColԭ����) = strԭ����
                
                mshBill.TextMatrix(intRow, mconintCol�ϴι�Ӧ��) = str��Ӧ��
                If mint��Ӧ��ѡ�� = 1 Then
                    gstrSQL = "Select B.���� From ҩƷ��� A, ��Ӧ�� B Where A.��ͬ��λid = B.ID And A.ҩƷid = [1] "
                    Set rs��ͬ��λ = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, Val(rsAllDrug!ҩƷid))
                    If Not rs��ͬ��λ.EOF Then
                        mshBill.TextMatrix(intRow, mconintCol�ϴι�Ӧ��) = rs��ͬ��λ!����
                    End If
                End If
                
                mshBill.TextMatrix(intRow, mconintCol�������) = zlStr.FormatEx(dbl�������, mintShowNumberDigit, , True)
                mshBill.TextMatrix(intRow, mconIntCol����ϵ��) = !����ϵ��
                
                mshBill.TextMatrix(intRow, mconintCol�ɱ���) = zlStr.FormatEx(IIf(IsNull(!����), "0", !���� * !����ϵ��), mintShowPriceDigit, , True)
                mshBill.TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx(IIf(IsNull(!�ۼ�), "0", !�ۼ� * !����ϵ��), mintShowPriceDigit, , True)
                mshBill.TextMatrix(intRow, mconIntCol����ҩ��) = IIf(IsNull(!����ҩ��), "", !����ҩ��)
                mshBill.TextMatrix(intRow, mconIntCol��׼�ĺ�) = IIf(IsNull(!��׼�ĺ�), "", !��׼�ĺ�)
                
                SetNumer !ҩƷid, lng�ⷿID, IIf(IsNull(!�������), 0, !�������), intRow, int�ƻ�����, int���Ʒ���, bln������ʽ
                
                str�ͻ���λ = IIf(IsNull(!�ͻ���λ), "", !�ͻ���λ)
                dbl�ͻ���װ = IIf(IsNull(!�ͻ���װ), 0, !�ͻ���װ)
                If dbl�ͻ���װ <> 0 Then
                    mshBill.TextMatrix(intRow, mconintCol�ͻ���װ) = dbl�ͻ���װ
                    mshBill.TextMatrix(intRow, mconintCol�ͻ���λ) = str�ͻ���λ & "(1" & str�ͻ���λ & "=" & zlStr.FormatEx(dbl�ͻ���װ / !����ϵ��, 1, , True) & mshBill.TextMatrix(intRow, mconIntCol��λ) & ")"
                    mshBill.TextMatrix(intRow, mconintCol�ͻ�����) = zlStr.FormatEx(Val(mshBill.TextMatrix(intRow, mconintCol�ƻ�����)) / dbl�ͻ���װ, 1, , True)
                End If
                
                If intRow = mshBill.rows - 1 Then mshBill.rows = mshBill.rows + 1
                Call zlControl.StaShowPercent(intRow / intRecord, staThis.Panels(2), frmDrugPlanCard)
            End If
            .MoveNext
        Loop
    End With
    Call ClearZeroPlan
    Call RefreshRowNO(mshBill, mconIntCol���, 1)
    Call ��ʾ�ϼƽ��
    Me.MousePointer = vbDefault
    mshBill.Redraw = True
    CmdSave.Enabled = True
    Pic����.Enabled = True
    CmdCancel.Enabled = True
    mshBill.Col = mconintCol�ƻ�����
    Me.staThis.Panels(2).Text = ""
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetDate(ByVal intģʽ As Integer, ByVal datCurrent As Date, _
        ByRef strBegin As String, ByRef strEnd As String) As Boolean
    Dim rsdate As New Recordset

    'intģʽ=1,�¼ƻ���2�����ƻ�
    On Error GoTo errHandle
    GetDate = False
    Select Case intģʽ
    Case 1
        strBegin = Year(datCurrent) & "-" & String(2 - Len(Month(datCurrent)), "0") & Month(datCurrent) & "-01"
        gstrSQL = "select last_day(to_date([1],'yyyy-mm-dd')) from dual"
        Set rsdate = zlDataBase.OpenSQLRecord(gstrSQL, "GetDate", Format(datCurrent, "yyyy-mm-dd"))
        strEnd = Format(rsdate.Fields(0), "yyyy-mm-dd")
        rsdate.Close
    Case 2
        Select Case DatePart("Q", datCurrent)
            Case 1
                strBegin = Year(datCurrent) & "-01-01"
                strEnd = Year(datCurrent) & "-03-31"
            Case 2
                strBegin = Year(datCurrent) & "-04-01"
                strEnd = Year(datCurrent) & "-06-30"
            Case 3
                strBegin = Year(datCurrent) & "-07-01"
                strEnd = Year(datCurrent) & "-09-30"
            Case 4
                strBegin = Year(datCurrent) & "-10-01"
                strEnd = Year(datCurrent) & "-12-31"
        End Select
    Case 4
        strBegin = Format(datCurrent, "yyyy-mm-dd")
        strEnd = Format(DateAdd("d", 6, datCurrent), "yyyy-mm-dd")
    End Select
    GetDate = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'����ǰ�������������������ƻ�����,����
Private Sub SetNumer(ByVal lngҩƷid As Long, ByVal lng�ⷿID As Long, _
        ByVal num������� As Double, ByVal intCurrentRow As Integer, _
        ByVal int�ƻ����� As Integer, ByVal int���Ʒ��� As Integer, ByVal bln������ʽ As Boolean)
    '---------------------------------------------------------------------------
    '--����:ȷ�����������ͼƻ�����
    '   1 )����ͬ�����Բ��շ�������ȥǰ��ͬ��ҩƷ����������������Թ滮ԭ��Ԥ�����ģ��Աȿ������ɹ��ƻ����û��޸ĵ���
    '   2 )�ٽ��ڼ�ƽ�����շ�����ͬ���ٽ��ڼ�(ǰ�ڡ�����)��ƽ������Ԥ�����ĶԱȿ������ɹ��ƻ����û��޸ĵ�����
    '   3 )ҩƷ�������շ�������ҩƷ������������������õĲ��ΪҩƷ�ƻ��ɹ�����

    '--����:
    '       int�ƻ�����:1:�¶ȼƻ�,2.���ȼƻ�,3.��ȼƻ�,4.�ܼƻ�
    '       int���Ʒ���:1 ��ʾ����ͬ�����Բ��շ�,2 �ٽ��ڼ�ƽ�����շ�,3.�����޶�;4.��������;5-�Զ�������
    '       bln������ʽ:false ������ָ���ƻ�����  �ƻ�����=��������-�������;  true �����޼ƻ����� �ƻ�����=��������-�������
    '--����:
    '---------------------------------------------------------------------------
    Dim numǰ������ As Double
    Dim num�������� As Double
    Dim num�������� As Double
    Dim num�ƻ����� As Double
    Dim num���� As Double, num���� As Double
    Dim lng���� As Long

    Dim datǰ�� As Date
    Dim dat���� As Date
    Dim strBegin As String
    Dim strEnd As String
    Dim rsNum As New Recordset
    
    Dim str����������� As String
    Dim str�շ�����ʱ�� As String
    
    On Error GoTo errHandle
    num������� = IIf(mbln�����ǿ�� = True, 0, num�������)
    
    With mshBill
        Select Case int���Ʒ���
            Case 1      '����ͬ�����β���   ֻ���¶ȡ����ȼƻ�
                datǰ�� = DateAdd("m", Choose(int�ƻ�����, 1, 3), DateAdd("yyyy", -2, mstrNow))
                dat���� = DateAdd("m", Choose(int�ƻ�����, 1, 3), DateAdd("yyyy", -1, mstrNow))
                If lng�ⷿID = 0 Then
                    GetDate int�ƻ�����, datǰ��, strBegin, strEnd
                    gstrSQL = "SELECT ABS(SUM(NVL(����, 0))) AS ǰ������ " _
                            & " FROM ҩƷ�շ����� a, ҩƷ������ b " _
                            & " Where a.���id = b.id " _
                            & "  and ���� <>6 AND b.ϵ�� = -1 " _
                            & "  AND ҩƷid+0 = [1] " _
                            & "  AND ���� BETWEEN [2] and [3] "
                    Set rsNum = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lngҩƷid, CDate(strBegin), CDate(strEnd))
                            
                    If rsNum.EOF Then
                        numǰ������ = 0
                    Else
                        numǰ������ = IIf(IsNull(rsNum.Fields(0)), 0, rsNum.Fields(0))
                    End If
                    rsNum.Close
                    GetDate int�ƻ�����, dat����, strBegin, strEnd
                    gstrSQL = "SELECT ABS(SUM(NVL(����, 0))) AS �������� " _
                            & " FROM ҩƷ�շ����� a, ҩƷ������ b " _
                           & " Where a.���id = b.id " _
                            & "  and ���� <>6 AND b.ϵ�� = -1 " _
                            & "  AND ҩƷid+0 = [1] " _
                            & "  AND ���� BETWEEN [2] and [3] "
                    Set rsNum = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lngҩƷid, CDate(strBegin), CDate(strEnd))
                    
                    If rsNum.EOF Then
                        num�������� = 0
                    Else
                        num�������� = IIf(IsNull(rsNum.Fields(0)), 0, rsNum.Fields(0))
                    End If
                    rsNum.Close
                Else
                    GetDate int�ƻ�����, datǰ��, strBegin, strEnd
                    gstrSQL = "SELECT ABS(SUM(NVL(����, 0))) AS ǰ������ " _
                            & " FROM ҩƷ�շ����� a, ҩƷ������ b " _
                           & " Where a.���id = b.id " _
                            & "  AND b.ϵ�� = -1 " _
                            & "  and �ⷿid+0=[1] " _
                            & "  AND ҩƷid+0= [2] " _
                            & "  AND ���� BETWEEN [3] and [4] "
                    Set rsNum = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lng�ⷿID, lngҩƷid, CDate(strBegin), CDate(strEnd))
                    
                    If rsNum.EOF Then
                        numǰ������ = 0
                    Else
                        numǰ������ = IIf(IsNull(rsNum.Fields(0)), 0, rsNum.Fields(0))
                    End If
                    rsNum.Close
                    GetDate int�ƻ�����, dat����, strBegin, strEnd
                    gstrSQL = "SELECT ABS(SUM(NVL(����, 0))) AS �������� " _
                            & " FROM ҩƷ�շ����� a, ҩƷ������ b " _
                           & " Where a.���id = b.id " _
                            & "  AND b.ϵ�� = -1 " _
                            & "  and �ⷿid+0=[1] " _
                            & "  AND ҩƷid+0= [2] " _
                            & "  AND ���� BETWEEN [3] and [4] "
                    Set rsNum = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lng�ⷿID, lngҩƷid, CDate(strBegin), CDate(strEnd))
                    
                    If rsNum.EOF Then
                        num�������� = 0
                    Else
                        num�������� = IIf(IsNull(rsNum.Fields(0)), 0, rsNum.Fields(0))
                    End If
                    rsNum.Close
                End If
    
                '�Ѹ���λת����ҩ�ⵥλ��
                num�������� = num�������� / .TextMatrix(intCurrentRow, mconIntCol����ϵ��)
                numǰ������ = numǰ������ / .TextMatrix(intCurrentRow, mconIntCol����ϵ��)
                '�ƻ�����=2������������ǰ���������������
                If mbln�ƻ����� Then
                    num�ƻ����� = 2 * num�������� - numǰ������ - num�������
                    If num�ƻ����� < 0 Then num�ƻ����� = 0
                End If
                .TextMatrix(intCurrentRow, mconIntColǰ������) = zlStr.FormatEx(numǰ������, mintShowNumberDigit, , True)
                .TextMatrix(intCurrentRow, mconIntCol��������) = zlStr.FormatEx(num��������, mintShowNumberDigit, , True)
                .TextMatrix(intCurrentRow, mconintCol�ƻ�����) = IIf(zlStr.FormatEx(num�ƻ�����, mintShowNumberDigit) = 0, "", zlStr.FormatEx(num�ƻ�����, mintShowNumberDigit, , True))
                .TextMatrix(intCurrentRow, mconIntCol�ɱ����) = _
                        IIf(zlStr.FormatEx(num�ƻ����� * IIf(mshBill.TextMatrix(intCurrentRow, mconintCol�ɱ���) = "", 0, mshBill.TextMatrix(intCurrentRow, mconintCol�ɱ���)), mintShowNumberDigit, , True) = 0 _
                        , "" _
                        , zlStr.FormatEx(num�ƻ����� * IIf(mshBill.TextMatrix(intCurrentRow, mconintCol�ɱ���) = "", 0, mshBill.TextMatrix(intCurrentRow, mconintCol�ɱ���)), mintShowNumberDigit, , True))
                .TextMatrix(intCurrentRow, mconIntCol�ۼ۽��) = _
                        IIf(zlStr.FormatEx(num�ƻ����� * IIf(mshBill.TextMatrix(intCurrentRow, mconIntCol�ۼ�) = "", 0, mshBill.TextMatrix(intCurrentRow, mconIntCol�ۼ�)), mintShowNumberDigit, , True) = 0 _
                        , "" _
                        , zlStr.FormatEx(num�ƻ����� * IIf(mshBill.TextMatrix(intCurrentRow, mconIntCol�ۼ�) = "", 0, mshBill.TextMatrix(intCurrentRow, mconIntCol�ۼ�)), mintShowNumberDigit, , True))
            Case 2      '�ٽ��ڼ�ƽ�����շ�
                datǰ�� = Choose(int�ƻ�����, DateAdd("m", -2, mstrNow), DateAdd("m", -6, mstrNow), DateAdd("yyyy", -2, mstrNow), DateAdd("d", -14, mdatThisMondyDate))
                dat���� = Choose(int�ƻ�����, DateAdd("m", -1, mstrNow), DateAdd("m", -3, mstrNow), DateAdd("yyyy", -1, mstrNow), DateAdd("d", -7, mdatThisMondyDate))
                If lng�ⷿID = 0 Then
                    gstrSQL = "SELECT ABS(SUM(NVL(����, 0))) AS ǰ������ " _
                            & " FROM ҩƷ�շ����� a, ҩƷ������ b " _
                           & " Where a.���id = b.id " _
                            & "  and ���� <>6 AND b.ϵ�� = -1 " _
                            & "  AND ҩƷid+0= [1] " _
                            & "  AND ���� BETWEEN [2] and [3] "
                    Set rsNum = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lngҩƷid, CDate(Format(DateAdd(Choose(int�ƻ�����, "m", "m", "m", "d"), Choose(int�ƻ�����, -1, -3, -12, -7), datǰ��), "yyyy-mm-dd hh:mm:ss")), _
                        CDate(Format(datǰ��, "yyyy-mm-dd hh:mm:ss")))
                    
                    If rsNum.EOF Then
                        numǰ������ = 0
                    Else
                        numǰ������ = IIf(IsNull(rsNum.Fields(0)), 0, rsNum.Fields(0))
                    End If
                    rsNum.Close
    
                    gstrSQL = "SELECT ABS(SUM(NVL(����, 0))) AS �������� " _
                            & " FROM ҩƷ�շ����� a, ҩƷ������ b " _
                           & " Where a.���id = b.id " _
                            & "  and ���� <>6 AND b.ϵ�� = -1 " _
                            & "  AND ҩƷid+0= [1] " _
                            & "  AND ���� BETWEEN [2] and [3] "
                    Set rsNum = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lngҩƷid, CDate(Format(DateAdd(Choose(int�ƻ�����, "m", "m", "m", "d"), Choose(int�ƻ�����, -1, -3, -12, -7), dat����), "yyyy-mm-dd hh:mm:ss")), _
                            CDate(Format(dat����, "yyyy-mm-dd hh:mm:ss")))
                            
                    If rsNum.EOF Then
                        num�������� = 0
                    Else
                        num�������� = IIf(IsNull(rsNum.Fields(0)), 0, rsNum.Fields(0))
                    End If
                    rsNum.Close
                Else
                    gstrSQL = "SELECT ABS(SUM(NVL(����, 0))) AS ǰ������ " _
                            & " FROM ҩƷ�շ����� a, ҩƷ������ b " _
                           & " Where a.���id = b.id " _
                            & "  AND b.ϵ�� = -1 " _
                            & "  and a.�ⷿid+0=[1] " _
                            & "  AND ҩƷid+0= [2]" _
                            & "  AND ���� BETWEEN [3] and [4] "
                    Set rsNum = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lng�ⷿID, lngҩƷid, CDate(Format(DateAdd(Choose(int�ƻ�����, "m", "m", "m", "d"), Choose(int�ƻ�����, -1, -3, -12, -7), datǰ��), "yyyy-mm-dd hh:mm;ss")), _
                            CDate(Format(datǰ��, "yyyy-mm-dd hh:mm:ss")))
                            
                    If rsNum.EOF Then
                        numǰ������ = 0
                    Else
                        numǰ������ = IIf(IsNull(rsNum.Fields(0)), 0, rsNum.Fields(0))
                    End If
                    rsNum.Close
                    
                    gstrSQL = "SELECT ABS(SUM(NVL(����, 0))) AS �������� " _
                            & " FROM ҩƷ�շ����� a, ҩƷ������ b " _
                           & " Where a.���id = b.id " _
                            & "  AND b.ϵ�� = -1 " _
                            & "  and a.�ⷿid+0=[1] " _
                            & "  AND ҩƷid+0= [2] " _
                            & "  AND ���� BETWEEN [3] and [4] "
                    Set rsNum = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lng�ⷿID, lngҩƷid, CDate(Format(DateAdd(Choose(int�ƻ�����, "m", "m", "m", "d"), Choose(int�ƻ�����, -1, -3, -12, -7), dat����), "yyyy-mm-dd hh:mm:ss")), _
                            CDate(Format(dat����, "yyyy-mm-dd hh:mm:ss")))
                            
                    If rsNum.EOF Then
                        num�������� = 0
                    Else
                        num�������� = IIf(IsNull(rsNum.Fields(0)), 0, rsNum.Fields(0))
                    End If
                    rsNum.Close
                End If
    
                '�Ѹ���λת����ҩ�ⵥλ��
                num�������� = num�������� / .TextMatrix(intCurrentRow, mconIntCol����ϵ��)
                numǰ������ = numǰ������ / .TextMatrix(intCurrentRow, mconIntCol����ϵ��)
                '�ƻ����� = (ǰ������ + ��������) / 2 - �������
                If mbln�ƻ����� Then
                    num�ƻ����� = (num�������� + numǰ������) / 2 - num�������
                    If num�ƻ����� < 0 Then num�ƻ����� = 0
                End If
                .TextMatrix(intCurrentRow, mconIntColǰ������) = zlStr.FormatEx(numǰ������, mintShowNumberDigit, , True)
                .TextMatrix(intCurrentRow, mconIntCol��������) = zlStr.FormatEx(num��������, mintShowNumberDigit, , True)
                .TextMatrix(intCurrentRow, mconintCol�ƻ�����) = IIf(zlStr.FormatEx(num�ƻ�����, mintShowNumberDigit, , True) = 0, "", zlStr.FormatEx(num�ƻ�����, mintShowNumberDigit, , True))
                .TextMatrix(intCurrentRow, mconIntCol�ɱ����) = _
                        IIf(zlStr.FormatEx(num�ƻ����� * IIf(mshBill.TextMatrix(intCurrentRow, mconintCol�ɱ���) = "", 0, mshBill.TextMatrix(intCurrentRow, mconintCol�ɱ���)), mintShowNumberDigit, , True) = 0 _
                            , "" _
                            , zlStr.FormatEx(num�ƻ����� * IIf(mshBill.TextMatrix(intCurrentRow, mconintCol�ɱ���) = "", 0, mshBill.TextMatrix(intCurrentRow, mconintCol�ɱ���)), mintShowNumberDigit, , True))
                .TextMatrix(intCurrentRow, mconIntCol�ۼ۽��) = _
                        IIf(zlStr.FormatEx(num�ƻ����� * IIf(mshBill.TextMatrix(intCurrentRow, mconIntCol�ۼ�) = "", 0, mshBill.TextMatrix(intCurrentRow, mconIntCol�ۼ�)), mintShowNumberDigit, , True) = 0 _
                            , "" _
                            , zlStr.FormatEx(num�ƻ����� * IIf(mshBill.TextMatrix(intCurrentRow, mconIntCol�ۼ�) = "", 0, mshBill.TextMatrix(intCurrentRow, mconIntCol�ۼ�)), mintShowNumberDigit, , True))
            
            Case 3      'ҩƷ����������շ�
                If mstrBeginDate = "" Or mstrEndDate = "" Then
                    mstrEndDate = Format(mstrNow, "yyyy-mm-dd")
                    mstrBeginDate = Format(DateAdd("m", -1, mstrNow), "yyyy-mm-dd")
                End If
                
                gstrSQL = "Select Max(����) As ���� From ҩƷ�շ�����"
                Set rsNum = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption)
                If NVL(rsNum!����, "") = "" Then
                    str����������� = Format(DateAdd("d", 1, CDate(mstrBeginDate)), "yyyy-mm-dd")
                Else
                    str����������� = Format(DateAdd("d", 1, rsNum!����), "yyyy-mm-dd")
                End If
                
                str�շ�����ʱ�� = Format(DateAdd("d", 1, CDate(mstrEndDate)), "yyyy-mm-dd")
                
                If lng�ⷿID = 0 Then
                    gstrSQL = " Select Sum(��������) As �������� " _
                            & " From (SELECT ABS(SUM(NVL(����, 0))) AS �������� " _
                            & " FROM ҩƷ�շ����� a, ҩƷ������ b " _
                            & " Where a.���id = b.id " _
                            & "  and ���� <>6 AND b.ϵ�� = -1 " _
                            & "  AND ҩƷid+0= [1] " _
                            & "  AND ���� BETWEEN [2] and [3] " _
                            & " Union All " _
                            & " Select Abs(Sum(A.���ϵ�� * Nvl(A.ʵ������, 0) * Nvl(A.����, 1))) As �������� " _
                            & " From ҩƷ�շ���¼ A, ҩƷ������ B " _
                            & " Where A.����<>6 And A.������id = B.ID And B.ϵ�� = -1 And ҩƷid + 0 = [1] And " _
                            & " ������� >= [2] " _
                            & " And ������� Between [4] And [5])"
                    Set rsNum = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lngҩƷid, CDate(mstrBeginDate), CDate(mstrEndDate), CDate(str�����������), CDate(str�շ�����ʱ��))
                            
                    If Not rsNum.EOF Then
                        num�������� = IIf(IsNull(rsNum.Fields(0)), 0, rsNum.Fields(0)) / Val(.TextMatrix(intCurrentRow, mconIntCol����ϵ��))
                         .TextMatrix(intCurrentRow, mconIntCol��������) = zlStr.FormatEx(num��������, mintShowNumberDigit, , True)
                    End If
                    rsNum.Close
                Else
                    gstrSQL = " Select Sum(��������) As �������� " _
                            & " From (SELECT ABS(SUM(NVL(����, 0))) AS �������� " _
                            & " FROM ҩƷ�շ����� a, ҩƷ������ b " _
                            & " Where a.���id = b.id " _
                            & "  AND b.ϵ�� = -1 " _
                            & "  and a.�ⷿid+0=[1] " _
                            & "  AND ҩƷid+0= [2] " _
                            & "  AND ���� BETWEEN [3] and [4] " _
                            & " Union All " _
                            & " Select Abs(Sum(A.���ϵ�� * Nvl(A.ʵ������, 0) * Nvl(A.����, 1))) As �������� " _
                            & " From ҩƷ�շ���¼ A, ҩƷ������ B " _
                            & " Where A.������id = B.ID And B.ϵ�� = -1 And A.�ⷿid + 0 = [1] And ҩƷid + 0 = [2] And " _
                            & " ������� >= [3] " _
                            & " And ������� Between [5] And [6])"
                    Set rsNum = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lng�ⷿID, lngҩƷid, CDate(mstrBeginDate), CDate(mstrEndDate), CDate(str�����������), CDate(str�շ�����ʱ��))
                            
                    If Not rsNum.EOF Then
                        num�������� = IIf(IsNull(rsNum.Fields(0)), 0, rsNum.Fields(0)) / Val(.TextMatrix(intCurrentRow, mconIntCol����ϵ��))
                         .TextMatrix(intCurrentRow, mconIntCol��������) = zlStr.FormatEx(num��������, mintShowNumberDigit, , True)
                    End If
                    rsNum.Close
                End If
                
                If lng�ⷿID = 0 Then
                    gstrSQL = "select sum(Nvl(����,0)) as  ����,sum(Nvl(����,0)) as  ���� " _
                            & " from ҩƷ�����޶� " _
                           & " where ҩƷid=[1] "
    
                Else
                    gstrSQL = "select Nvl(����,0) As ����,Nvl(����,0) as ���� " _
                            & " from ҩƷ�����޶� " _
                           & " where ҩƷid=[1] " _
                           & "   and �ⷿid=[2]"
    
                End If
                Set rsNum = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lngҩƷid, lng�ⷿID)
                
                If rsNum.EOF Then
                    num���� = 0
                    num���� = 0
                Else
                    num���� = IIf(IsNull(rsNum!����), 0, rsNum!����)
                    num���� = IIf(IsNull(rsNum!����), 0, rsNum!����)
                End If
    
                '�Ѹ���λת����ҩ�ⵥλ��
                num���� = num���� / .TextMatrix(intCurrentRow, mconIntCol����ϵ��)
                num���� = num���� / .TextMatrix(intCurrentRow, mconIntCol����ϵ��)
                '�ƻ�����=�������ޣ��������
                If mbln�ƻ����� Then
                    If bln������ʽ = False Then
                        num�ƻ����� = IIf(num���� > num�������, num���� - num�������, 0)
                    Else
                        num�ƻ����� = IIf(num���� > num�������, num���� - num�������, 0)
                    End If
                End If
                
                .TextMatrix(intCurrentRow, mconintCol�ƻ�����) = IIf(zlStr.FormatEx(num�ƻ�����, mintShowNumberDigit) = 0, "", zlStr.FormatEx(num�ƻ�����, mintShowNumberDigit, , True))
                .TextMatrix(intCurrentRow, mconIntCol�ɱ����) = _
                        IIf(zlStr.FormatEx(num�ƻ����� * IIf(mshBill.TextMatrix(intCurrentRow, mconintCol�ɱ���) = "", 0, mshBill.TextMatrix(intCurrentRow, mconintCol�ɱ���)), mintShowNumberDigit, , True) = 0 _
                            , "" _
                            , zlStr.FormatEx(num�ƻ����� * IIf(mshBill.TextMatrix(intCurrentRow, mconintCol�ɱ���) = "", 0, mshBill.TextMatrix(intCurrentRow, mconintCol�ɱ���)), mintShowNumberDigit, , True))
                .TextMatrix(intCurrentRow, mconIntCol�ۼ۽��) = _
                        IIf(zlStr.FormatEx(num�ƻ����� * IIf(mshBill.TextMatrix(intCurrentRow, mconIntCol�ۼ�) = "", 0, mshBill.TextMatrix(intCurrentRow, mconIntCol�ۼ�)), mintShowNumberDigit, , True) = 0 _
                            , "" _
                            , zlStr.FormatEx(num�ƻ����� * IIf(mshBill.TextMatrix(intCurrentRow, mconIntCol�ۼ�) = "", 0, mshBill.TextMatrix(intCurrentRow, mconIntCol�ۼ�)), mintShowNumberDigit, , True))
            Case 4  '��������
                datǰ�� = Choose(int�ƻ�����, DateAdd("m", -2, mstrNow), DateAdd("m", -6, mstrNow), DateAdd("yyyy", -2, mstrNow), DateAdd("d", -14, mdatThisMondyDate))
                dat���� = Choose(int�ƻ�����, DateAdd("m", -1, mstrNow), DateAdd("m", -3, mstrNow), DateAdd("yyyy", -1, mstrNow), DateAdd("d", -7, mdatThisMondyDate))
                GetDate int�ƻ�����, dat����, strBegin, strEnd
                lng���� = CDate(Format(strEnd, "yyyy-MM-DD")) - CDate(Format(strBegin, "yyyy-MM-DD")) + 1
                If lng���� <= 0 Then lng���� = 1
                
                If lng�ⷿID = 0 Then
                    gstrSQL = "SELECT ABS(SUM(NVL(����, 0))) AS �������� " _
                            & " FROM ҩƷ�շ����� a, ҩƷ������ b " _
                           & " Where a.���id = b.id " _
                            & "  and ���� <>6 AND b.ϵ�� = -1 " _
                            & "  AND ҩƷid+0 = [1] " _
                            & "  AND ���� BETWEEN [2] and [3] "
                    Set rsNum = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lngҩƷid, CDate(strBegin), CDate(strEnd))
                    
                    If rsNum.EOF Then
                        num�������� = 0
                    Else
                        num�������� = IIf(IsNull(rsNum.Fields(0)), 0, rsNum.Fields(0))
                    End If
                    rsNum.Close
                Else
                    gstrSQL = "SELECT ABS(SUM(NVL(����, 0))) AS �������� " _
                            & " FROM ҩƷ�շ����� a, ҩƷ������ b " _
                           & " Where a.���id = b.id " _
                            & "  AND b.ϵ�� = -1 " _
                            & "  and �ⷿid+0=[1] " _
                            & "  AND ҩƷid+0= [2] " _
                            & "  AND ���� BETWEEN [3] and [4] "
                    Set rsNum = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lng�ⷿID, lngҩƷid, CDate(strBegin), CDate(strEnd))
                    
                    If rsNum.EOF Then
                        num�������� = 0
                    Else
                        num�������� = IIf(IsNull(rsNum.Fields(0)), 0, rsNum.Fields(0))
                    End If
                    rsNum.Close
                End If
    
                '�Ѹ���λת����ҩ�ⵥλ��
                num�������� = num�������� / .TextMatrix(intCurrentRow, mconIntCol����ϵ��)
                num���� = num�������� / lng���� * mint����
                num���� = num�������� / lng���� * mint����
                '�ƻ�����=2������������ǰ���������������
                
                If mbln�ƻ����� Then
                    If num������� < num���� Then
                        num�ƻ����� = num���� - num�������
                    Else
                        num�ƻ����� = 0
                    End If
                    If num�ƻ����� < 0 Then num�ƻ����� = 0
                End If
                .TextMatrix(intCurrentRow, mconIntColǰ������) = zlStr.FormatEx(numǰ������, mintShowNumberDigit, , True)
                .TextMatrix(intCurrentRow, mconIntCol��������) = zlStr.FormatEx(num��������, mintShowNumberDigit, , True)
                .TextMatrix(intCurrentRow, mconintCol�ƻ�����) = IIf(zlStr.FormatEx(num�ƻ�����, mintShowNumberDigit) = 0, "", zlStr.FormatEx(num�ƻ�����, mintShowNumberDigit, , True))
                .TextMatrix(intCurrentRow, mconIntCol�ɱ����) = _
                        IIf(zlStr.FormatEx(num�ƻ����� * IIf(mshBill.TextMatrix(intCurrentRow, mconintCol�ɱ���) = "", 0, mshBill.TextMatrix(intCurrentRow, mconintCol�ɱ���)), mintShowNumberDigit, , True) = 0 _
                        , "" _
                        , zlStr.FormatEx(num�ƻ����� * IIf(mshBill.TextMatrix(intCurrentRow, mconintCol�ɱ���) = "", 0, mshBill.TextMatrix(intCurrentRow, mconintCol�ɱ���)), mintShowNumberDigit, , True))
                .TextMatrix(intCurrentRow, mconIntCol�ۼ۽��) = _
                        IIf(zlStr.FormatEx(num�ƻ����� * IIf(mshBill.TextMatrix(intCurrentRow, mconIntCol�ۼ�) = "", 0, mshBill.TextMatrix(intCurrentRow, mconIntCol�ۼ�)), mintShowNumberDigit, , True) = 0 _
                        , "" _
                        , zlStr.FormatEx(num�ƻ����� * IIf(mshBill.TextMatrix(intCurrentRow, mconIntCol�ۼ�) = "", 0, mshBill.TextMatrix(intCurrentRow, mconIntCol�ۼ�)), mintShowNumberDigit, , True))
            Case 5  '�Զ�������
                If mstrBeginDate = "" Or mstrEndDate = "" Then
                    mstrEndDate = Format(mstrNow, "yyyy-mm-dd")
                    mstrBeginDate = Format(DateAdd("m", -1, mstrNow), "yyyy-mm-dd")
                End If
                
                gstrSQL = "Select Max(����) As ���� From ҩƷ�շ�����"
                Set rsNum = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption)
                If NVL(rsNum!����, "") = "" Then
                    str����������� = Format(DateAdd("d", 1, CDate(mstrBeginDate)), "yyyy-mm-dd")
                Else
                    str����������� = Format(DateAdd("d", 1, rsNum!����), "yyyy-mm-dd")
                End If
                
                str�շ�����ʱ�� = Format(DateAdd("d", 1, CDate(mstrEndDate)), "yyyy-mm-dd")
                
                If lng�ⷿID = 0 Then
                    gstrSQL = " Select Sum(��������) As �������� " _
                            & " From (SELECT ABS(SUM(NVL(����, 0))) AS �������� " _
                            & " FROM ҩƷ�շ����� a, ҩƷ������ b " _
                            & " Where a.���id = b.id " _
                            & "  and ���� <>6 AND b.ϵ�� = -1 " _
                            & "  AND ҩƷid+0= [1] " _
                            & "  AND ���� BETWEEN [2] and [3] " _
                            & " Union All " _
                            & " Select Abs(Sum(A.���ϵ�� * Nvl(A.ʵ������, 0) * Nvl(A.����, 1))) As �������� " _
                            & " From ҩƷ�շ���¼ A, ҩƷ������ B " _
                            & " Where A.����<>6 And A.������id = B.ID And B.ϵ�� = -1 And ҩƷid + 0 = [1] And " _
                            & " ������� >= [2] " _
                            & " And ������� Between [4] And [5])"
                    Set rsNum = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lngҩƷid, CDate(mstrBeginDate), CDate(mstrEndDate), CDate(str�����������), CDate(str�շ�����ʱ��))
                            
                    If rsNum.EOF Then
                        num�������� = 0
                    Else
                        num�������� = IIf(IsNull(rsNum.Fields(0)), 0, rsNum.Fields(0))
                    End If
                    rsNum.Close
                Else
                    gstrSQL = " Select Sum(��������) As �������� " _
                            & " From (SELECT ABS(SUM(NVL(����, 0))) AS �������� " _
                            & " FROM ҩƷ�շ����� a, ҩƷ������ b " _
                            & " Where a.���id = b.id " _
                            & "  AND b.ϵ�� = -1 " _
                            & "  and a.�ⷿid+0=[1] " _
                            & "  AND ҩƷid+0= [2] " _
                            & "  AND ���� BETWEEN [3] and [4] " _
                            & " Union All " _
                            & " Select Abs(Sum(A.���ϵ�� * Nvl(A.ʵ������, 0) * Nvl(A.����, 1))) As �������� " _
                            & " From ҩƷ�շ���¼ A, ҩƷ������ B " _
                            & " Where A.������id = B.ID And B.ϵ�� = -1 And A.�ⷿid + 0 = [1] And ҩƷid + 0 = [2] And " _
                            & " ������� >= [3] " _
                            & " And ������� Between [5] And [6])"
                    Set rsNum = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lng�ⷿID, lngҩƷid, CDate(mstrBeginDate), CDate(mstrEndDate), CDate(str�����������), CDate(str�շ�����ʱ��))
                            
                    If rsNum.EOF Then
                        num�������� = 0
                    Else
                        num�������� = IIf(IsNull(rsNum.Fields(0)), 0, rsNum.Fields(0))
                    End If
                    rsNum.Close
                End If
    
                '�Ѹ���λת����ҩ�ⵥλ��
                num�������� = num�������� / .TextMatrix(intCurrentRow, mconIntCol����ϵ��)
    
                If mbln�ƻ����� Then
                    If num�������� > num������� Then
                        num�ƻ����� = num�������� - num�������
                    End If
                End If
                .TextMatrix(intCurrentRow, mconIntColǰ������) = zlStr.FormatEx(num��������, mintShowNumberDigit, , True)
                .TextMatrix(intCurrentRow, mconintCol�ƻ�����) = IIf(zlStr.FormatEx(num�ƻ�����, mintShowNumberDigit, , True) = 0, "", zlStr.FormatEx(num�ƻ�����, mintShowNumberDigit, , True))
                .TextMatrix(intCurrentRow, mconIntCol�ɱ����) = _
                        IIf(zlStr.FormatEx(num�ƻ����� * IIf(mshBill.TextMatrix(intCurrentRow, mconintCol�ɱ���) = "", 0, mshBill.TextMatrix(intCurrentRow, mconintCol�ɱ���)), mintShowNumberDigit, , True) = 0 _
                        , "" _
                        , zlStr.FormatEx(num�ƻ����� * IIf(mshBill.TextMatrix(intCurrentRow, mconintCol�ɱ���) = "", 0, mshBill.TextMatrix(intCurrentRow, mconintCol�ɱ���)), mintShowNumberDigit, , True))
                .TextMatrix(intCurrentRow, mconIntCol�ۼ۽��) = _
                        IIf(zlStr.FormatEx(num�ƻ����� * IIf(mshBill.TextMatrix(intCurrentRow, mconIntCol�ۼ�) = "", 0, mshBill.TextMatrix(intCurrentRow, mconIntCol�ۼ�)), mintShowNumberDigit, , True) = 0 _
                        , "" _
                        , zlStr.FormatEx(num�ƻ����� * IIf(mshBill.TextMatrix(intCurrentRow, mconIntCol�ۼ�) = "", 0, mshBill.TextMatrix(intCurrentRow, mconIntCol�ۼ�)), mintShowNumberDigit, , True))
        End Select
        
        'ȡ�ⷿ�ϡ�����
        If lng�ⷿID = 0 Then
            gstrSQL = "select sum(Nvl(����,0)) as  ����,sum(Nvl(����,0)) as ���� " _
                    & " from ҩƷ�����޶� " _
                    & " where ҩƷid=[1] "
        
        Else
            gstrSQL = "select Nvl(����,0) As ����,Nvl(����,0) As ���� " _
                    & " from ҩƷ�����޶� " _
                    & " where ҩƷid=[1] " _
                    & "   and �ⷿid=[2]"
        End If
        Set rsNum = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lngҩƷid, lng�ⷿID)
        
        If Not rsNum.EOF Then
            .TextMatrix(intCurrentRow, mconIntCol�������) = zlStr.FormatEx(NVL(rsNum!����, 0) / Val(.TextMatrix(intCurrentRow, mconIntCol����ϵ��)), mintShowNumberDigit, , True)
            .TextMatrix(intCurrentRow, mconIntCol�������) = zlStr.FormatEx(NVL(rsNum!����, 0) / Val(.TextMatrix(intCurrentRow, mconIntCol����ϵ��)), mintShowNumberDigit, , True)
        End If
        
        '�ֱ�������ںͱ��ڵ�������
        'ȡ���ڵ����䷶Χ
        Select Case int�ƻ�����
            '1:�¶ȼƻ�,2.���ȼƻ�,3.��ȼƻ�,4.�ܼƻ�
            Case 1
                '����ʱ�䷶Χ
                strBegin = Format(DateAdd("m", -1, CDate(mstrNow)), "YYYY-MM") & "-01"
                strEnd = Format(DateAdd("d", -1, CDate(Format(CDate(mstrNow), "YYYY-MM") & "-01")), "YYYY-MM-DD") & " 23:59:59"
            Case 2
                '�ϼ���ʱ�䷶Χ
                Select Case DatePart("Q", CDate(mstrNow))
                    Case 1
                        strBegin = Format(DateAdd("yyyy", -1, CDate(mstrNow)), "YYYY") & "-10-01"
                        strEnd = Format(DateAdd("yyyy", -1, CDate(mstrNow)), "YYYY") & "-12-31 23:59:59"
                    Case 2
                        strBegin = Format(mstrNow, "YYYY") & "-01-01"
                        strEnd = Format(mstrNow, "YYYY") & "-03-31 23:59:59"
                     Case 3
                        strBegin = Format(mstrNow, "YYYY") & "-04-01"
                        strEnd = Format(mstrNow, "YYYY") & "-06-30 23:59:59"
                    Case 4
                        strBegin = Format(mstrNow, "YYYY") & "-07-01"
                        strEnd = Format(mstrNow, "YYYY") & "-09-30 23:59:59"
                End Select
            Case 3
                '�����ʱ�䷶Χ
                strBegin = Format(DateAdd("yyyy", -1, CDate(mstrNow)), "YYYY") & "-01-01"
                strEnd = Format(DateAdd("yyyy", -1, CDate(mstrNow)), "YYYY") & "-12-31 23:59:59"
            Case 4
                '����ʱ�䷶Χ�����й�ϰ�ߣ���һ��һ�����ڵĵ�һ�죩
                strBegin = Format(DateAdd("d", -DatePart("w", CDate(mstrNow), vbMonday) + 1, DateAdd("d", -7, CDate(mstrNow))), "YYYY-MM-DD")
                strEnd = Format(DateAdd("d", 6, CDate(strBegin)), "YYYY-MM-DD") & " 23:59:59"
        End Select
        
        '������������������Ҫ��ȷֵ����ҩƷ�շ�����ͳ�ƣ�
        gstrSQL = "Select -Sum(Nvl(����, 0)) As �������� " & _
            " From ҩƷ�շ�����" & _
            " Where ���� + 0 In (8, 9, 10) And ҩƷid+0=[1] And ���� Between [2] And [3] "
        If mstr��Դҩ�� <> "" Then
            gstrSQL = gstrSQL & " And �ⷿid In(select * from Table(Cast(f_Num2List([4]) As zlTools.t_NumList)))"
        End If
        
        If int���Ʒ��� = 5 Then
            '�Զ������䣬����������Ϊ��������
            strBegin = Format(DateAdd("m", -1, CDate(mstrNow)), "YYYY-MM") & "-01"
            strEnd = Format(DateAdd("d", -1, CDate(Format(CDate(mstrNow), "YYYY-MM") & "-01")), "YYYY-MM-DD") & " 23:59:59"
        End If
        
        Set rsNum = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lngҩƷid, CDate(strBegin), CDate(strEnd), mstr��Դҩ��)
        If rsNum.RecordCount > 0 Then
            .TextMatrix(intCurrentRow, mconintCol��������) = zlStr.FormatEx(NVL(rsNum!��������, 0) / Val(.TextMatrix(intCurrentRow, mconIntCol����ϵ��)), mintShowNumberDigit, , True)
        End If
        
        '����ͬ�����β���,ȡȥ��ͬ������
        If int���Ʒ��� = 1 Then
            dat���� = DateAdd("m", Choose(int�ƻ�����, 1, 3), DateAdd("yyyy", -1, mstrNow))
            GetDate int�ƻ�����, dat����, strBegin, strEnd
            
            gstrSQL = "Select -Sum(Nvl(����, 0)) As �������� " & _
                " From ҩƷ�շ�����" & _
                " Where ���� + 0 In (8, 9, 10) And ҩƷid+0=[1] And ���� Between [2] And [3] "
            If mstr��Դҩ�� <> "" Then
                gstrSQL = gstrSQL & " And �ⷿid In(select * from Table(Cast(f_Num2List([4]) As zlTools.t_NumList)))"
            End If
            
            Set rsNum = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lngҩƷid, CDate(strBegin), CDate(strEnd), mstr��Դҩ��)
            If rsNum.RecordCount > 0 Then
                num�������� = zlStr.FormatEx(NVL(rsNum!��������, 0) / Val(.TextMatrix(intCurrentRow, mconIntCol����ϵ��)), mintShowNumberDigit, , True)
            End If
        End If
        
        'ȡ���ڵ����䷶Χ
        Select Case int�ƻ�����
            '1:�¶ȼƻ�,2.���ȼƻ�,3.��ȼƻ�,4.�ܼƻ�
            Case 1
                '����ʱ�䷶Χ
                strBegin = Format(mstrNow, "YYYY-MM") & "-01"
            Case 2
                '������ʱ�䷶Χ
                Select Case DatePart("Q", CDate(mstrNow))
                    Case 1
                        strBegin = Format(mstrNow, "YYYY") & "-01-01"
                    Case 2
                        strBegin = Format(mstrNow, "YYYY") & "-04-01"
                    Case 3
                        strBegin = Format(mstrNow, "YYYY") & "-07-01"
                    Case 4
                        strBegin = Format(mstrNow, "YYYY") & "-10-01"
                End Select
            Case 3
                '�����ʱ�䷶Χ
                strBegin = Format(mstrNow, "YYYY") & "-01-01"
            Case 4
                '����ʱ�䷶Χ�����й�ϰ�ߣ���һ��һ�����ڵĵ�һ�죩
                strBegin = Format(DateAdd("d", -DatePart("w", CDate(mstrNow), vbMonday) + 1, CDate(mstrNow)), "YYYY-MM-DD")
        End Select
        
        '���ڽ���ʱ���ֹ������
        strEnd = Format(mstrNow, "YYYY-MM-DD") & " 23:59:59"
            
        '���㱾������������Ҫ��ȷֵ����ҩƷ�շ�����ͳ�ƣ�
        gstrSQL = "Select -Sum(Nvl(����, 0)) As �������� " & _
            " From ҩƷ�շ�����" & _
            " Where ���� + 0 In (8, 9, 10) And ҩƷid+0=[1] And ���� Between [2] And [3] "
        If mstr��Դҩ�� <> "" Then
            gstrSQL = gstrSQL & " And �ⷿid In(select * from Table(Cast(f_Num2List([4]) As zlTools.t_NumList)))"
        End If
        
        If int���Ʒ��� = 5 Then
            '�Զ������䣬����������Ϊ��������
            strBegin = Format(mstrNow, "YYYY-MM") & "-01"
            strEnd = Format(mstrNow, "YYYY-MM-DD")
        End If
        
        Set rsNum = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lngҩƷid, CDate(strBegin), CDate(strEnd), mstr��Դҩ��)
        If rsNum.RecordCount > 0 Then
            .TextMatrix(intCurrentRow, mconintCol��������) = zlStr.FormatEx(NVL(rsNum!��������, 0) / Val(.TextMatrix(intCurrentRow, mconIntCol����ϵ��)), mintShowNumberDigit, , True)
        End If
        
        '�Զ������䷨������������������
        If int���Ʒ��� = 5 Then
            If mstrBeginDate = "" Or mstrEndDate = "" Then
                mstrBeginDate = Format(DateAdd("m", -1, mstrNow), "yyyy-mm-dd")
                mstrEndDate = Format(mstrNow, "yyyy-mm-dd")
            End If
            
            gstrSQL = "Select -Sum(Nvl(����, 0)) As �������� " & _
                " From ҩƷ�շ�����" & _
                " Where ���� + 0 In (8, 9, 10) And ҩƷid+0=[1] And ���� Between [2] And [3] "
            If mstr��Դҩ�� <> "" Then
                gstrSQL = gstrSQL & " And �ⷿid In(select * from Table(Cast(f_Num2List([4]) As zlTools.t_NumList)))"
            End If
            
            '�Զ������䷨������������Ϊ��������
            Set rsNum = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lngҩƷid, CDate(mstrBeginDate), CDate(mstrEndDate), mstr��Դҩ��)
            If rsNum.RecordCount > 0 Then
                .TextMatrix(intCurrentRow, mconIntCol��������) = zlStr.FormatEx(NVL(rsNum!��������, 0) / Val(.TextMatrix(intCurrentRow, mconIntCol����ϵ��)), mintShowNumberDigit, , True)
            End If
        End If
        
        '�����������ƻ�����
        If mbln�����������ƻ� Then
            If mbln�ƻ����� Then
                If int���Ʒ��� = 5 Then
                    num�������� = Val(.TextMatrix(intCurrentRow, mconIntCol��������))
                ElseIf int���Ʒ��� = 1 Then
                    num�������� = num��������
                Else
                    num�������� = Val(.TextMatrix(intCurrentRow, mconintCol��������))
                End If
                
                If num�������� > num������� Then
                    num�ƻ����� = num�������� - num�������
                Else
                    num�ƻ����� = 0
                End If
                               
'                .TextMatrix(intCurrentRow, mconIntCol��������) = zlStr.FormatEx(num��������, mintShowNumberDigit, , True)
                .TextMatrix(intCurrentRow, mconintCol�ƻ�����) = IIf(zlStr.FormatEx(num�ƻ�����, mintShowNumberDigit, , True) = 0, "", zlStr.FormatEx(num�ƻ�����, mintShowNumberDigit, , True))
                .TextMatrix(intCurrentRow, mconIntCol�ɱ����) = _
                        IIf(zlStr.FormatEx(num�ƻ����� * IIf(mshBill.TextMatrix(intCurrentRow, mconintCol�ɱ���) = "", 0, mshBill.TextMatrix(intCurrentRow, mconintCol�ɱ���)), mintShowNumberDigit, , True) = 0 _
                        , "" _
                        , zlStr.FormatEx(num�ƻ����� * IIf(mshBill.TextMatrix(intCurrentRow, mconintCol�ɱ���) = "", 0, mshBill.TextMatrix(intCurrentRow, mconintCol�ɱ���)), mintShowNumberDigit, , True))
                .TextMatrix(intCurrentRow, mconIntCol�ۼ۽��) = _
                        IIf(zlStr.FormatEx(num�ƻ����� * IIf(mshBill.TextMatrix(intCurrentRow, mconIntCol�ۼ�) = "", 0, mshBill.TextMatrix(intCurrentRow, mconIntCol�ۼ�)), mintShowNumberDigit, , True) = 0 _
                        , "" _
                        , zlStr.FormatEx(num�ƻ����� * IIf(mshBill.TextMatrix(intCurrentRow, mconIntCol�ۼ�) = "", 0, mshBill.TextMatrix(intCurrentRow, mconIntCol�ۼ�)), mintShowNumberDigit, , True))
            End If
        End If
        
        'ȡ��ʷ�ɹ��ƻ�
        Call LoadHisPlane(lng�ⷿID, lngҩƷid, intCurrentRow)
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Dim rsDetail As New Recordset
    Dim bln��ҩ�ⷿ As Boolean
    Dim str�ⷿ���� As String
    Dim str������ As String
    Dim i As Integer, j As Integer
    
    On Error GoTo errHandle
    
    marrFrom = Array()
    marrInitGrid = Array()
    Call GetDefineSize
    Call GetDrugDigit(mlng�ⷿID, MStrCaption, mintUnit, mintShowCostDigit, mintShowPriceDigit, mintShowNumberDigit, mintShowMoneyDigit)
    
    mblnViewCost = zlStr.IsHavePrivs(mstrPrivs, "�鿴�ɱ���")
    '��֯��ʽ����
    mstrCostFormat = "'999999999990." & String(mintShowCostDigit, "0") & "'"
    mstrPriceFormat = "'999999999990." & String(mintShowPriceDigit, "0") & "'"
    mstrNumberFormat = "'999999999990." & String(mintShowNumberDigit, "0") & "'"
    mstrMoneyFormat = "'999999999990." & String(mintShowMoneyDigit, "0") & "'"
    mintPlanPoint = Val(zlDataBase.GetPara("ȫԺ�ƻ�����վ��", glngSys, 1330, 0))

    mintPriceUnit = GetUnit()
    txtNo = mstr���ݺ�
    txtNo.Tag = txtNo
    mblnEnter = True
    mstrNow = Format(Sys.Currentdate, "yyyy-mm-dd")
        
    mintDrugNameShow = Int(Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ�ƻ�����", "ҩƷ������ʾ��ʽ", 0)))
    If mintDrugNameShow > 2 Or mintDrugNameShow < 0 Then mintDrugNameShow = 0
    mnuColDrug.Item(mintDrugNameShow).Checked = True
    
    Call GegReg
    
    IniHisPlaneRec
    
    LblTitle.Caption = GetUnitName & LblTitle.Caption
    initCard
    
    'ֻ����ҩ��ⷿ����ʾ"ԭ����"��
    str�ⷿ���� = ""
    gstrSQL = "Select �������� From ��������˵�� Where ����id =[1]"
    Set rsDetail = zlDataBase.OpenSQLRecord(gstrSQL, "�жϿⷿ����", mlng�ⷿID)
    Do While Not rsDetail.EOF
        str�ⷿ���� = str�ⷿ���� & "," & rsDetail!��������
        rsDetail.MoveNext
    Loop
    If str�ⷿ���� Like "*��ҩ*" Or str�ⷿ���� Like "*�Ƽ���*" Then bln��ҩ�ⷿ = True
    str������ = zlDataBase.GetPara("������", glngSys, ģ���.ҩƷ�ƻ�)
    If InStr(1, "|" & str������ & "|", "|ԭ����|") = 0 Then mshBill.ColWidth(mconIntColԭ����) = IIf(bln��ҩ�ⷿ, 800, 0)
    
    For i = 1 To mconIntColS - 1
        ReDim Preserve marrInitGrid(UBound(marrInitGrid) + 1)
        marrInitGrid(UBound(marrInitGrid)) = mshBill.TextMatrix(0, i) & "|" & mshBill.ColWidth(i)
    Next
    
    '�ָ����Ի�����
    RestoreWinState Me, App.ProductName, MStrCaption

    For i = 1 To mconIntColS - 1
        ReDim Preserve marrFrom(UBound(marrFrom) + 1)
        marrFrom(UBound(marrFrom)) = mshBill.TextMatrix(0, i) & "|" & mshBill.ColWidth(i)
    Next
    
    For i = 0 To UBound(marrInitGrid)
        For j = 0 To UBound(marrFrom)
            If Split(marrInitGrid(i), "|")(0) = Split(marrFrom(j), "|")(0) And Split(marrInitGrid(i), "|")(1) * Split(marrFrom(j), "|")(1) = 0 Then
                mshBill.ColWidth(i + 1) = Split(marrInitGrid(i), "|")(1)
            End If
        Next
    Next
    
    '�ָ����Ի����ú���Ҫ���¸���Ȩ�޿������Ƿ���ʾ
    With mshBill
        If mblnViewCost = False Then
            .ColWidth(mconintCol�ɱ���) = 0
            .ColWidth(mconIntCol�ɱ����) = 0
        End If
    End With
    
    If (mint�༭״̬ = 4 Or mint�༭״̬ = 6) And Trim(Txt�����.Caption) <> "" Then
        If mshBill.ColWidth(mconintColִ������) = 0 And InStr(1, "|" & mstrColumn_UnSelected & "|", "|ִ������|") = 0 Then mshBill.ColWidth(mconintColִ������) = 1100
    Else
        mshBill.ColWidth(mconintColִ������) = 0
    End If
    
    '��Ʒ���д���
    If gintҩƷ������ʾ = 2 Then
        '��ʾ��Ʒ����
        mshBill.ColWidth(mconIntCol��Ʒ��) = IIf(mshBill.ColWidth(mconIntCol��Ʒ��) = 0, 2000, mshBill.ColWidth(mconIntCol��Ʒ��))
    Else
        '��������ʾ��Ʒ����
        mshBill.ColWidth(mconIntCol��Ʒ��) = 0
    End If
    
'    If grsMaster.State = adStateClosed Then
'        Call SetSelectorRS(1, mstrCaption, mlng�ⷿID, mlng�ⷿID)
'    End If

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub initCard()
    Dim i As Integer
    Dim rsInitCard As New Recordset
    Dim strUnitQuantity As String
    Dim intRow As Integer
    Dim intRecordCount As Integer
    Dim str��λ As String
    Dim strOrder As String, strCompare As String
    Dim strҩ�� As String
    Dim strSqlOrder As String
    Dim str�ͻ���λ As String
    Dim dbl�ͻ���װ As Double
    
    On Error GoTo errHandle
    strOrder = zlDataBase.GetPara("����", glngSys, ģ���.ҩƷ�ƻ�)
    strCompare = Mid(strOrder, 1, 1)
    
    strSqlOrder = "���"
    
    If strCompare = "0" Then
        strSqlOrder = "���"
    ElseIf strCompare = "1" Then
        strSqlOrder = "ҩƷ����"
    ElseIf strCompare = "2" Then
        If gintҩƷ������ʾ = 0 Or gintҩƷ������ʾ = 2 Then
            strSqlOrder = "ͨ����"
        Else
            strSqlOrder = "Nvl(��Ʒ��, ͨ����)"
        End If
    End If
    
    strSqlOrder = strSqlOrder & IIf(Right(strOrder, 1) = "0", " ASC", " DESC")

    '�ⷿ
    Select Case mint�༭״̬
        Case 1
            Txt������ = UserInfo.�û�����
            Txt�������� = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
            initGrid
        Case 2, 3, 4, 5, 6
            Select Case mintUnit
            Case 1
                str��λ = ",d.�ۼ۵�λ ��λ,d.�ۼ۰�װ ����ϵ��"
            Case 2
                str��λ = ",d.���ﵥλ ��λ,d.�����װ ����ϵ��"
            Case 3
                str��λ = ",d.סԺ��λ ��λ,d.סԺ��װ ����ϵ��"
            Case 4
                str��λ = ",d.ҩ�ⵥλ ��λ,d.ҩ���װ ����ϵ��"
            End Select
            
            gstrSQL = "SELECT a.id,nvl(a.�ⷿid,0) as �ⷿid,nvl(c.����,'ȫԺ') AS �ⷿ,a.ҩ��id, a.no, a.�ƻ�����,a.�ڼ�, a.���Ʒ���, a.������," _
                    & "TO_CHAR (a.��������, 'yyyy-mm-dd HH24:MI:SS') AS ��������, a.�����," _
                    & "TO_CHAR (a.�������, 'yyyy-mm-dd HH24:MI:SS') AS �������,a.������,TO_CHAR (a.��������, 'yyyy-mm-dd HH24:MI:SS') AS ��������,a.����˵��," _
                    & "b.���,b.ҩƷid,d.ҩƷ����,d.ͨ����,d.��Ʒ��,d.ҩƷ��Դ, d.���,d.����ҩ��" & str��λ & ", nvl(b.ǰ������,0) as ǰ������, nvl(b.��������,0) as ��������, " _
                    & " nvl(b.��������,0) as ��������,nvl(b.��������,0) as ��������,b.�������, b.�ƻ�����,nvl(b.ִ������,0) as ִ������,b.�ͻ�����,d.�ͻ���λ,d.�ͻ���װ, b.����, b.���, b.�ϴι�Ӧ��,b.�ϴ�������,d.ԭ����,b.˵��,b.�ۼ�,b.�ۼ۽��,d.��������,b.��׼�ĺ� " _
                    & " FROM ҩƷ�ɹ��ƻ� a, ҩƷ�ƻ����� b,���ű� c," _
                    & " (SELECT DISTINCT a.ҩƷid," _
                    & " '[' || C.���� || ']' As ҩƷ����, C.���� As ͨ����, a.ԭ����,B.���� As ��Ʒ��,a.ҩƷ��Դ,c.���,a.ҩ�ⵥλ,A.ҩ���װ,a.����ҩ��,a.���ﵥλ,A.�����װ,a.סԺ��λ,a.סԺ��װ,a.�ͻ���λ,a.�ͻ���װ,C.���㵥λ �ۼ۵�λ,1 �ۼ۰�װ,c.�������� " _
                    & " FROM ҩƷ��� a, �շ���Ŀ���� b, �շ���ĿĿ¼ c " _
                    & " WHERE a.ҩƷid = b.�շ�ϸĿID(+) and B.����(+)=3 " _
                    & "   AND a.ҩƷid = c.ID) d " _
                    & "Where a.id = b.�ƻ�id " _
                    & "  and nvl(a.�ⷿid,0)=c.id(+) " _
                    & "  and b.ҩƷid=d.ҩƷid " _
                    & "  AND a.no = [1] " & _
                    " Order by " & strSqlOrder
            Set rsInitCard = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, mstr���ݺ�)
            
            If rsInitCard.EOF Then
                mintParallelRecord = 2
                Exit Sub
            End If

            intRecordCount = rsInitCard.RecordCount

            Txt������ = rsInitCard!������
            If mint�༭״̬ = 2 Then
                Txt������ = UserInfo.�û�����
            End If
            Txt�������� = Format(rsInitCard!��������, "yyyy-mm-dd hh:mm:ss")

            Txt����� = IIf(IsNull(rsInitCard!�����), "", rsInitCard!�����)
            Txt������� = IIf(IsNull(rsInitCard!�������), "", Format(rsInitCard!�������, "yyyy-mm-dd hh:mm:ss"))
            
            txt������ = IIf(IsNull(rsInitCard!������), "", rsInitCard!������)
            txt�������� = IIf(IsNull(rsInitCard!��������), "", Format(rsInitCard!��������, "yyyy-mm-dd hh:mm:ss"))
            
            txtժҪ.Text = IIf(IsNull(rsInitCard!����˵��), "", rsInitCard!����˵��)
            txtժҪ.Tag = NVL(rsInitCard!ҩ��id)
            txt�ƻ����� = Choose(rsInitCard!�ƻ����� + 1, "��ʱ", "�¶ȼƻ�", "���ȼƻ�", "��ȼƻ�", "�ܼƻ�")
            txt���Ʒ��� = Choose(rsInitCard!���Ʒ��� + 1, "�����������", "����ͬ�����β��շ�", "�ٽ��ڼ�ƽ�����շ�", "ҩƷ����������շ�", "ҩƷ�����������շ�", "�Զ���������շ�")
            mint�ƻ����� = rsInitCard!�ƻ�����
            mint���Ʒ��� = rsInitCard!���Ʒ���
            mlng�ⷿID = rsInitCard!�ⷿid
            mlng�ƻ�ID = rsInitCard!Id

            Str�ڼ� = IIf(IsNull(rsInitCard!�ڼ�), "", rsInitCard!�ڼ�)
            Select Case mint�ƻ�����
                Case 0       '��ʱ�ƻ�
                    LblTitle.Caption = GetUnitName & rsInitCard!�ⷿ & "�ɹ��ƻ�"
                Case 1       '�¼ƻ�
                    LblTitle.Caption = GetUnitName & "(" & Mid(Str�ڼ�, 1, 4) & "��" & Right(Str�ڼ�, 2) & "��" & ") " & rsInitCard!�ⷿ & "�ɹ��ƻ�"
                Case 2       '���ƻ�
                    LblTitle.Caption = GetUnitName & "(" & Mid(Str�ڼ�, 1, 4) & "��" & Right(Str�ڼ�, 1) & "��" & ")" & rsInitCard!�ⷿ & "�ɹ��ƻ�"
                Case 3       '��ƻ�
                    LblTitle.Caption = GetUnitName & "(" & Str�ڼ� & "��" & ")" & rsInitCard!�ⷿ & "�ɹ��ƻ�"
                Case 4       '�ܼƻ�
                    LblTitle.Caption = GetUnitName & "(" & Mid(Str�ڼ�, 1, 4) & "��" & Mid(Str�ڼ�, 5, 2) & "��" & Right(Str�ڼ�, 2) & "��" & ")" & LblTitle.Tag & "�ɹ��ƻ�"
            End Select

            If (mint�༭״̬ = 2 Or mint�༭״̬ = 3) And Txt����� <> "" Then
                mintParallelRecord = 3
                Exit Sub
            End If
            
            'If IIf(IsNull(rsInitCard!����), 0, rsInitCard!����) = 0 Then
            '    mint�۸���ʾ = 1
            'Else
            '    mint�۸���ʾ = 0
            'End If
            
            initGrid
            
            
            If mint���Ʒ��� = 5 Then
                '�Զ���������Ʒ�
                mshBill.TextMatrix(0, mconIntColǰ������) = "��������"
                mshBill.TextMatrix(0, mconIntCol��������) = "��������"
                mshBill.TextMatrix(0, mconintCol��������) = "��������"
                mshBill.TextMatrix(0, mconintCol��������) = "��������"
            End If
            
            With mshBill
                For intRow = 1 To intRecordCount

                    .TextMatrix(intRow, 0) = rsInitCard!ҩƷid
                    
                    If gintҩƷ������ʾ = 0 Or gintҩƷ������ʾ = 2 Then
                        strҩ�� = rsInitCard!ͨ����
                    Else
                        strҩ�� = IIf(IsNull(rsInitCard!��Ʒ��), rsInitCard!ͨ����, rsInitCard!��Ʒ��)
                    End If
                    
                    .TextMatrix(intRow, mconIntColҩƷ���������) = rsInitCard!ҩƷ���� & strҩ��
                    .TextMatrix(intRow, mconIntColҩƷ����) = rsInitCard!ҩƷ����
                    .TextMatrix(intRow, mconIntColҩƷ����) = strҩ��
                    
                    If mintDrugNameShow = 1 Then
                        .TextMatrix(intRow, mconIntColҩ��) = .TextMatrix(intRow, mconIntColҩƷ����)
                    ElseIf mintDrugNameShow = 2 Then
                        .TextMatrix(intRow, mconIntColҩ��) = .TextMatrix(intRow, mconIntColҩƷ����)
                    Else
                        .TextMatrix(intRow, mconIntColҩ��) = .TextMatrix(intRow, mconIntColҩƷ���������)
                    End If
                    
                    .TextMatrix(intRow, mconIntCol��Ʒ��) = IIf(IsNull(rsInitCard!��Ʒ��), "", rsInitCard!��Ʒ��)

                    .TextMatrix(intRow, mconIntCol��Դ) = NVL(rsInitCard!ҩƷ��Դ)
                    .TextMatrix(intRow, mconIntCol���) = IIf(IsNull(rsInitCard!���), "", rsInitCard!���)
                    .TextMatrix(intRow, mconintCol�ϴι�Ӧ��) = IIf(IsNull(rsInitCard!�ϴι�Ӧ��), "", rsInitCard!�ϴι�Ӧ��)
                    .TextMatrix(intRow, mconIntCol������) = IIf(IsNull(rsInitCard!�ϴ�������), "", rsInitCard!�ϴ�������)
                    .TextMatrix(intRow, mconIntColԭ����) = IIf(IsNull(rsInitCard!ԭ����), "", rsInitCard!ԭ����)
                    .TextMatrix(intRow, mconIntCol��λ) = rsInitCard!��λ
                    .TextMatrix(intRow, mconIntcolҽ������) = IIf(IsNull(rsInitCard!��������), "", rsInitCard!��������)
                    .TextMatrix(intRow, mconIntCol����ϵ��) = rsInitCard!����ϵ��
                    .TextMatrix(intRow, mconIntColǰ������) = zlStr.FormatEx(rsInitCard!ǰ������ / rsInitCard!����ϵ��, mintShowNumberDigit, , True)
                    .TextMatrix(intRow, mconIntCol��������) = zlStr.FormatEx(rsInitCard!�������� / rsInitCard!����ϵ��, mintShowNumberDigit, , True)
                    .TextMatrix(intRow, mconintCol��������) = zlStr.FormatEx(rsInitCard!�������� / rsInitCard!����ϵ��, mintShowNumberDigit, , True)
                    .TextMatrix(intRow, mconintCol��������) = zlStr.FormatEx(rsInitCard!�������� / rsInitCard!����ϵ��, mintShowNumberDigit, , True)
                    .TextMatrix(intRow, mconintCol�������) = zlStr.FormatEx(rsInitCard!������� / rsInitCard!����ϵ��, mintShowNumberDigit, , True)
                    .TextMatrix(intRow, mconintCol�ƻ�����) = IIf(zlStr.FormatEx(rsInitCard!�ƻ�����, mintShowNumberDigit, , True) = 0, "", zlStr.FormatEx(rsInitCard!�ƻ����� / rsInitCard!����ϵ��, mintShowNumberDigit, , True))
                    .TextMatrix(intRow, mconintColִ������) = IIf(zlStr.FormatEx(rsInitCard!ִ������, mintShowNumberDigit, , True) = 0, "", zlStr.FormatEx(rsInitCard!ִ������ / rsInitCard!����ϵ��, mintShowNumberDigit, , True))
                    .TextMatrix(intRow, mconintColԭִ������) = IIf(zlStr.FormatEx(rsInitCard!ִ������, mintShowNumberDigit, , True) = 0, "", zlStr.FormatEx(rsInitCard!ִ������ / rsInitCard!����ϵ��, mintShowNumberDigit, , True))
                    
                    dbl�ͻ���װ = IIf(IsNull(rsInitCard!�ͻ���װ), 0, rsInitCard!�ͻ���װ)
                    str�ͻ���λ = IIf(IsNull(rsInitCard!�ͻ���λ), "", rsInitCard!�ͻ���λ)
                    If dbl�ͻ���װ <> 0 Then
                        .TextMatrix(intRow, mconintCol�ͻ�����) = IIf(IsNull(rsInitCard!�ͻ�����), "", zlStr.FormatEx(rsInitCard!�ͻ�����, 1, , True))
                        .TextMatrix(intRow, mconintCol�ͻ���λ) = str�ͻ���λ & "(1" & str�ͻ���λ & "=" & zlStr.FormatEx(dbl�ͻ���װ / rsInitCard!����ϵ��, 1, , True) & rsInitCard!��λ & ")"
                        .TextMatrix(intRow, mconintCol�ͻ���װ) = dbl�ͻ���װ
                    End If
                    
                    .TextMatrix(intRow, mconintCol�ɱ���) = zlStr.FormatEx(rsInitCard!���� * rsInitCard!����ϵ��, mintShowPriceDigit, , True)
                    .TextMatrix(intRow, mconIntCol�ɱ����) = IIf(zlStr.FormatEx(rsInitCard!���, mintShowMoneyDigit, , True) = 0, "", zlStr.FormatEx(rsInitCard!���, mintShowMoneyDigit, , True))
                    .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx(rsInitCard!�ۼ� * rsInitCard!����ϵ��, mintShowPriceDigit, , True)
                    .TextMatrix(intRow, mconIntCol�ۼ۽��) = IIf(zlStr.FormatEx(rsInitCard!�ۼ۽��, mintShowMoneyDigit) = 0, "", zlStr.FormatEx(rsInitCard!�ۼ۽��, mintShowMoneyDigit, , True))
                    
                    .TextMatrix(intRow, mconintCol˵��) = IIf(IsNull(rsInitCard!˵��), "", rsInitCard!˵��)
                    .TextMatrix(intRow, mconIntCol����ҩ��) = IIf(IsNull(rsInitCard!����ҩ��), "", rsInitCard!����ҩ��)
                    .TextMatrix(intRow, mconIntCol��׼�ĺ�) = IIf(IsNull(rsInitCard!��׼�ĺ�), "", rsInitCard!��׼�ĺ�)
                    If intRow = .rows - 1 Then .rows = .rows + 1
                    rsInitCard.MoveNext
                    
                    Call LoadHisPlane(mlng�ⷿID, Val(.TextMatrix(intRow, 0)), intRow)
                Next
            End With
            rsInitCard.Close
    End Select
    Call RefreshRowNO(mshBill, mconIntCol���, 1)
    Call ��ʾ�ϼƽ��
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'��ʼ���༭�ؼ�
Private Sub initGrid()
    Dim intCol As Integer

    Call SetColumnByUserDefine '������
    With mshBill
        .Active = True
        .Cols = mconIntColS
        .MsfObj.FixedCols = 2

        .TextMatrix(0, mconIntCol���) = "���"
        .TextMatrix(0, mconIntColҩ��) = "ҩƷ���������"
        .TextMatrix(0, mconIntCol��Ʒ��) = "��Ʒ��"
        .TextMatrix(0, mconIntCol��Դ) = "ҩƷ��Դ"
        .TextMatrix(0, mconIntCol���) = "���"
        .TextMatrix(0, mconIntCol������) = "������"
        .TextMatrix(0, mconIntColԭ����) = "ԭ����"
        .TextMatrix(0, mconIntCol��λ) = "��λ"
        .TextMatrix(0, mconIntCol����ϵ��) = "����ϵ��"
        .TextMatrix(0, mconIntcolҽ������) = "ҽ������"
        .TextMatrix(0, mconIntColǰ������) = "ǰ������"
        .TextMatrix(0, mconIntCol��������) = "��������"
        .TextMatrix(0, mconIntCol�������) = "�������"
        .TextMatrix(0, mconIntCol�������) = "�������"
        .TextMatrix(0, mconintCol�������) = "�������"
        .TextMatrix(0, mconintCol��������) = "��������"
        .TextMatrix(0, mconintCol��������) = "��������"
        .TextMatrix(0, mconintCol�ƻ�����) = "�ƻ�����"
        .TextMatrix(0, mconintColִ������) = "ִ������"
        .TextMatrix(0, mconintColԭִ������) = "ԭִ������"
        .TextMatrix(0, mconintCol�ͻ���λ) = "�ͻ���λ"
        .TextMatrix(0, mconintCol�ͻ�����) = "�ͻ�����"
        
        .TextMatrix(0, mconintCol�ɱ���) = "�ɱ���"
        .TextMatrix(0, mconIntCol�ɱ����) = "�ɱ����"
        .TextMatrix(0, mconIntCol�ۼ�) = "�ۼ�"
        .TextMatrix(0, mconIntCol�ۼ۽��) = "�ۼ۽��"
        
        .TextMatrix(0, mconintCol�ϴι�Ӧ��) = "�ϴι�Ӧ��"
        .TextMatrix(0, mconintCol˵��) = "˵��"
        .TextMatrix(0, mconIntColҩƷ���������) = "ҩƷ���������"
        .TextMatrix(0, mconIntColҩƷ����) = "ҩƷ����"
        .TextMatrix(0, mconIntColҩƷ����) = "ҩƷ����"
        .TextMatrix(0, mconIntCol����ҩ��) = "����ҩ��"
        .TextMatrix(0, mconIntCol��׼�ĺ�) = "��׼�ĺ�"
        
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, mconIntCol���) = "1"
    
        .ColWidth(mconIntCol���) = 500
        .ColWidth(mconIntColҩ��) = 2000
        .ColWidth(mconIntCol��Ʒ��) = 2000
        .ColWidth(mconIntCol��Դ) = 1000
        .ColWidth(mconIntCol���) = 900
        .ColWidth(mconIntCol������) = 800
        .ColWidth(mconIntColԭ����) = 0
        .ColWidth(mconIntCol��λ) = 500
        .ColWidth(mconIntcolҽ������) = 1000
        .ColWidth(mconIntColǰ������) = 1100
        .ColWidth(mconIntCol��������) = 1100
        .ColWidth(mconIntCol�������) = 1100
        .ColWidth(mconIntCol�������) = 1100
        .ColWidth(mconintCol�������) = 1100
        .ColWidth(mconintCol��������) = 1100
        .ColWidth(mconintCol��������) = 1100
        .ColWidth(mconintCol�ƻ�����) = 1100
        .ColWidth(mconintColִ������) = IIf(mint�༭״̬ = 6, 1100, 0)
        .ColWidth(mconintColԭִ������) = 0
        .ColWidth(mconintCol�ͻ���λ) = 1500
        .ColWidth(mconintCol�ͻ�����) = 1100
        .ColWidth(mconintCol�ͻ���װ) = 0
        
        If mint�۸���ʾ = 0 Then
            .ColWidth(mconintCol�ɱ���) = 1000
            .ColWidth(mconIntCol�ɱ����) = 1200
            .ColWidth(mconIntCol�ۼ�) = 0
            .ColWidth(mconIntCol�ۼ۽��) = 0
        ElseIf mint�۸���ʾ = 1 Then
            .ColWidth(mconintCol�ɱ���) = 0
            .ColWidth(mconIntCol�ɱ����) = 0
            .ColWidth(mconIntCol�ۼ�) = 1000
            .ColWidth(mconIntCol�ۼ۽��) = 1200
        Else
            .ColWidth(mconintCol�ɱ���) = 1000
            .ColWidth(mconIntCol�ɱ����) = 1200
            .ColWidth(mconIntCol�ۼ�) = 1000
            .ColWidth(mconIntCol�ۼ۽��) = 1200
        End If
        If mblnViewCost = False Then
            .ColWidth(mconintCol�ɱ���) = 0
            .ColWidth(mconIntCol�ɱ����) = 0
        End If
        
        .ColWidth(mconintCol�ϴι�Ӧ��) = 2000
        .ColWidth(mconIntCol����ϵ��) = 0
        .ColWidth(mconintCol˵��) = 3000
        .ColWidth(0) = 0
        .ColWidth(mconIntColҩƷ���������) = 0
        .ColWidth(mconIntColҩƷ����) = 0
        .ColWidth(mconIntColҩƷ����) = 0
        .ColWidth(mconIntCol����ҩ��) = 2000
        .ColWidth(mconIntCol��׼�ĺ�) = 2000
        
        '-1����ʾ���п���ѡ���ǲ����ͣ�"��"��" "��
        ' 0����ʾ���п���ѡ�񣬵������޸�
        ' 1����ʾ���п������룬�ⲿ��ʾΪ��ťѡ��
        ' 2����ʾ�����������У��ⲿ��ʾΪ��ťѡ�񣬵���������ѡ���
        ' 3����ʾ������ѡ���У��ⲿ��ʾΪ������ѡ��
        '4:  ��ʾ����Ϊ�������ı����û�����
        '5:  ��ʾ���в�����ѡ��
        For intCol = 0 To .Cols - 1
            .ColData(intCol) = 5
        Next
        
        .ColData(mconintCol�ͻ���λ) = 5
        .ColData(mconintCol�ͻ�����) = 5
        If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
            txtժҪ.Enabled = True
            .ColData(mconIntColҩ��) = 1
            .ColData(mconintCol�ƻ�����) = 4
            .ColData(mconintCol�ɱ���) = 4
            .ColData(mconIntCol������) = 4
            .ColData(mconIntColԭ����) = 4
            .ColData(mconintCol�ϴι�Ӧ��) = 1
            .ColData(mconintCol˵��) = 4
            .ColData(mconintCol�ͻ�����) = 4
            .ColData(mconIntCol��׼�ĺ�) = 4
        ElseIf mint�༭״̬ = 4 Then
            txtժҪ.Enabled = False
            .ColData(mconintCol�ƻ�����) = 0
        ElseIf mint�༭״̬ = 3 Then
            txtժҪ.Enabled = False
            .ColData(mconintCol�ƻ�����) = 4
            .ColData(mconintCol˵��) = 4
        ElseIf mint�༭״̬ = 6 Then
            .ColData(mconintColִ������) = 4
        End If
        
        .ColAlignment(mconIntColҩ��) = flexAlignLeftCenter
        .ColAlignment(mconIntCol��Ʒ��) = flexAlignLeftCenter
        .ColAlignment(mconIntCol��Դ) = flexAlignLeftCenter
        .ColAlignment(mconIntCol���) = flexAlignLeftCenter
        .ColAlignment(mconIntCol������) = flexAlignLeftCenter
        .ColAlignment(mconIntColԭ����) = flexAlignLeftCenter
        .ColAlignment(mconIntCol��λ) = flexAlignCenterCenter
        .ColAlignment(mconIntcolҽ������) = flexAlignLeftCenter
        .ColAlignment(mconIntColǰ������) = flexAlignRightCenter
        .ColAlignment(mconIntCol��������) = flexAlignRightCenter
        .ColAlignment(mconIntCol�������) = flexAlignRightCenter
        .ColAlignment(mconIntCol�������) = flexAlignRightCenter
        .ColAlignment(mconintCol�������) = flexAlignRightCenter
        .ColAlignment(mconintCol��������) = flexAlignRightCenter
        .ColAlignment(mconintCol��������) = flexAlignRightCenter
        .ColAlignment(mconintCol�ƻ�����) = flexAlignRightCenter
        .ColAlignment(mconintColִ������) = flexAlignRightCenter
        .ColAlignment(mconintCol�ͻ���λ) = flexAlignCenterCenter
        .ColAlignment(mconintCol�ͻ�����) = flexAlignRightCenter
        .ColAlignment(mconintCol�ɱ���) = flexAlignRightCenter
        .ColAlignment(mconIntCol�ɱ����) = flexAlignRightCenter
        .ColAlignment(mconIntCol�ۼ�) = flexAlignRightCenter
        .ColAlignment(mconIntCol�ۼ۽��) = flexAlignRightCenter
        .ColAlignment(mconintCol�ϴι�Ӧ��) = flexAlignLeftCenter
        .ColAlignment(mconintCol˵��) = flexAlignLeftCenter
        .ColAlignment(mconIntCol����ҩ��) = flexAlignLeftCenter
        .ColAlignment(mconIntCol��׼�ĺ�) = flexAlignLeftCenter
        
        .PrimaryCol = mconIntColҩ��
        .LocateCol = mconIntColҩ��
        Call SetColumnByUserDefine '������
        If InStr(1, "34", mint�༭״̬) <> 0 Then .ColData(mconIntColҩ��) = 0
    End With

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub

    With Pic����
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - IIf(staThis.Visible, staThis.Height, 0) - .Top - 100 - CmdCancel.Height - 100
    End With

    With LblTitle
        .Left = 0
        .Top = 150
        .Width = Pic����.Width
    End With


    With mshBill
        .Left = 200
        .Width = Pic����.Width - .Left * 2
    End With
    With txtNo
        .Left = mshBill.Left + mshBill.Width - .Width
        LblNo.Left = .Left - LblNo.Width - 100
        .Top = LblTitle.Top
        LblNo.Top = .Top
    End With

    txt���Ʒ���.Left = mshBill.Left + mshBill.Width - txt���Ʒ���.Width
    lbl���Ʒ���.Left = txt���Ʒ���.Left - lbl���Ʒ���.Width - 100


    Lbl�ƻ�����.Left = mshBill.Left

    txt�ƻ�����.Left = Lbl�ƻ�����.Left + Lbl�ƻ�����.Width + 100

    With Lbl��������
        .Top = Pic����.Height - 100 - .Height
        .Left = mshBill.Left + 100
    End With
    
    With Txt��������
        .Top = Lbl��������.Top - 60
        .Left = Lbl��������.Left + Lbl��������.Width + 100
    End With
    
    With Lbl������
        .Top = Lbl��������.Top - Lbl��������.Height - 180
        .Left = Lbl��������.Left
    End With
    
    With Txt������
        .Top = Lbl������.Top - 60
        .Left = Txt��������.Left
    End With
    
    With Lbl�������
        .Top = Lbl��������.Top
        .Left = mshBill.Left + (mshBill.Width - .Width - Txt�������.Width - 100) / 2
    End With
    
    With Txt�������
        .Top = Txt��������.Top
        .Left = Lbl�������.Left + Lbl�������.Width + 100
    End With
    
    With Lbl�����
        .Top = Lbl������.Top
        .Left = Lbl�������.Left
    End With
    
    With Txt�����
        .Top = Txt������.Top
        .Left = Txt�������.Left
    End With
    
    With txt��������
        .Top = Txt��������.Top
        .Left = mshBill.Left + mshBill.Width - .Width
    End With
    
    With lbl��������
        .Top = Lbl��������.Top
        .Left = txt��������.Left - 100 - .Width
    End With
    
    With txt������
        .Top = Txt������.Top
        .Left = txt��������.Left
    End With
    
    With lbl������
        .Top = Lbl������.Top
        .Left = lbl��������.Left
    End With
    
    
    With txtժҪ
        .Top = Lbl������.Top - 140 - .Height
        .Left = Txt������.Left
        .Width = mshBill.Left + mshBill.Width - .Left
    End With

    With lblժҪ
        .Top = txtժҪ.Top + 50
        .Left = txtժҪ.Left - .Width - 200
    End With

    With lblPurchasePrice
        .Left = mshBill.Left
        .Top = txtժҪ.Top - 60 - .Height
        .Width = mshBill.Width
    End With
    If mblnViewCost = False Then
        lblPurchasePrice.Visible = False
    End If
    
    With mshBill
        .Height = lblPurchasePrice.Top - .Top - 60
    End With

    With CmdCancel
        .Left = Pic����.Left + mshBill.Left + mshBill.Width - .Width
        .Top = Pic����.Top + Pic����.Height + 100
    End With

    With CmdSave
        .Left = CmdCancel.Left - .Width - 100
        .Top = CmdCancel.Top
    End With

    With cmdHelp
        .Left = Pic����.Left + mshBill.Left
        .Top = CmdCancel.Top
    End With

    With cmdFind
        .Left = cmdHelp.Left + cmdHelp.Width + 200
        .Top = CmdCancel.Top
    End With

    With lblCode
        .Left = cmdFind.Left + cmdFind.Width + 50
        .Top = CmdCancel.Top + 50
    End With
    With txtCode
        .Left = lblCode.Left + lblCode.Width + 50
        .Top = CmdCancel.Top + 30
    End With
    
    With chk���ؽ��ڲɹ��ƻ�
        .Left = txtCode.Left + txtCode.Width + 150
        .Top = CmdCancel.Top + 30
    End With

    With chk�Ƿ���ʾ������
        .Left = txtCode.Left + txtCode.Width + 150 + chk���ؽ��ڲɹ��ƻ�.Width
        .Top = CmdCancel.Top + 30
    End With
    
    Call ResizeHisPlane
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (mblnChange = False Or mint�༭״̬ = 4 Or mint�༭״̬ = 3) Then
        If MsgBox("���ݿ����Ѹı䣬��δ���̣���Ҫ�˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1
            Exit Sub
        End If
    End If
    
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\ҩƷ�ƻ�����", "ҩƷ������ʾ��ʽ", mintDrugNameShow)
    
    mstr�Զ���ⷿ = ""
    mstr��Դ�ⷿ = ""
    mstr��Դҩ�� = ""
    mstrAll��Դ�ⷿ = ""
    mstrAll��Դҩ�� = ""
    mblnStart = False
    
    Call ReleaseSelectorRS

    Set mfrmMain = Nothing
End Sub

Private Sub SetDrugName(ByVal intType As Integer)
    'ҩƷ������ʾ��
    'intType��0����ʾ��������ƣ�1������ʾ���룻2������ʾ����
    Dim lngRow As Long
    
    If intType = mintDrugNameShow Then Exit Sub
    
    mintDrugNameShow = intType
    
    With mshBill
        For lngRow = 1 To .rows - 1
            If .TextMatrix(lngRow, mconIntColҩ��) <> "" Then
                If mintDrugNameShow = 1 Then
                    .TextMatrix(lngRow, mconIntColҩ��) = .TextMatrix(lngRow, mconIntColҩƷ����)
                ElseIf mintDrugNameShow = 2 Then
                    .TextMatrix(lngRow, mconIntColҩ��) = .TextMatrix(lngRow, mconIntColҩƷ����)
                Else
                    .TextMatrix(lngRow, mconIntColҩ��) = .TextMatrix(lngRow, mconIntColҩƷ���������)
                End If
            End If
        Next
    End With
End Sub
Private Function SaveCheck() As Boolean
    Dim str����� As String

    mblnSave = False
    SaveCheck = False

    str����� = UserInfo.�û�����

    On Error GoTo errHandle
    'zl_ҩƷ�ƻ�����_VERIFY( /*ID_IN*/, /*�����_IN*/ );
    gstrSQL = "zl_ҩƷ�ƻ�����_VERIFY('" & mlng�ƻ�ID & "','" & str����� & "')"
    Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
    
    SaveCheck = True
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    Exit Function
errHandle:
    'MsgBox "���ʧ�ܣ�", vbInformation, gstrSysName
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Function

Private Function SaveReCheck() As Boolean
    Dim str����� As String

    mblnSave = False
    SaveReCheck = False

    str����� = UserInfo.�û�����

    On Error GoTo errHandle
    'zl_ҩƷ�ƻ�����_VERIFY( /*ID_IN*/, /*�����_IN*/ );
    gstrSQL = "zl_ҩƷ�ƻ�����_VERIFY('" & mlng�ƻ�ID & "','" & str����� & "',1)"
    Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
    
    SaveReCheck = True
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Function

Private Function SaveExeAmount() As Boolean
    '�޸�ִ������
    Dim strInput As String
    Dim strNo As String
    Dim intRow As Integer

    mblnSave = False
    SaveExeAmount = False
    
    strNo = Trim(txtNo.Caption)
    
    With mshBill
        For intRow = 1 To .rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                If Val(.TextMatrix(intRow, mconintColִ������)) - Val(.TextMatrix(intRow, mconintColԭִ������)) <> 0 Then
                    strInput = IIf(strInput = "", "", strInput & "|") & Val(.TextMatrix(intRow, 0)) & "," & (Val(.TextMatrix(intRow, mconintColִ������)) - Val(.TextMatrix(intRow, mconintColԭִ������))) * Val(.TextMatrix(intRow, mconIntCol����ϵ��))
                End If
            End If
        Next
    End With
    
    On Error GoTo errHandle
    
    'Zl_ҩƷ�ƻ�����_�޸�ִ������( /*No_In*/, /*Input_In*/ );
    'ִ������Ϊ����������ԭִ��������������¼������ӻ���ٵ�������
    gstrSQL = "Zl_ҩƷ�ƻ�����_�޸�ִ������("
    'No_In
    gstrSQL = gstrSQL & "'" & strNo & "'"
    'Input_In  --��ʽ:"ҩƷID1,ִ������1|ҩƷID2,ִ������2|....."
    gstrSQL = gstrSQL & ",'" & strInput & "'"
    gstrSQL = gstrSQL & ")"
    
    Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
    
    SaveExeAmount = True
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Function

Private Sub Image1_Click()

End Sub

Private Sub imgDown_Click()
    picHscSend.Tag = 0
    imgUp.Visible = True
    imgDown.Visible = False
    
    Call ResizeHisPlane
End Sub

Private Sub imgUp_Click()
    picHscSend.Tag = 1
    imgUp.Visible = False
    imgDown.Visible = True
    
    Call ResizeHisPlane
End Sub


Private Sub mnuColDrug_Click(index As Integer)
    Dim n As Integer
    
    With mnuColDrug
        For n = 0 To .count - 1
            .Item(n).Checked = False
        Next
        
        .Item(index).Checked = True
        
        Call SetDrugName(index)
    End With
End Sub

Private Sub Msf��Ӧ��ѡ��_DblClick()
    Dim blnCancel As Boolean
    With mshBill
        .Text = Msf��Ӧ��ѡ��.TextMatrix(Msf��Ӧ��ѡ��.Row, 2)
        .TextMatrix(.Row, mconintCol�ϴι�Ӧ��) = Msf��Ӧ��ѡ��.TextMatrix(Msf��Ӧ��ѡ��.Row, 2)
    End With
    Msf��Ӧ��ѡ��.Visible = False
    mshBill.SetFocus
    If mshBill.Col <> mshBill.Cols - 1 Then
        mshBill.Col = mshBill.Col + 1
    End If
End Sub

Private Sub Msf��Ӧ��ѡ��_GotFocus()
    If Msf��Ӧ��ѡ��.rows - 1 = 1 Then Call Msf��Ӧ��ѡ��_DblClick
End Sub

Private Sub Msf��Ӧ��ѡ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call Msf��Ӧ��ѡ��_DblClick
    End If
End Sub

Private Sub Msf��Ӧ��ѡ��_LostFocus()
    Msf��Ӧ��ѡ��.ZOrder 1
    Msf��Ӧ��ѡ��.Visible = False
End Sub

Private Sub mshBill_AfterAddRow(Row As Long)
    Call RefreshRowNO(mshBill, mconIntCol���, Row)
End Sub

Private Sub mshBill_AfterDeleteRow()
    Call RefreshRowNO(mshBill, mconIntCol���, mshBill.Row)
    Call ��ʾ�ϼƽ��
End Sub

Private Sub mshBill_BeforeAddRow(Row As Long)
    If mshBill.ColData(mconIntColҩ��) = 0 Then
        'Cancel = True    '�ȴ���CANCEL����
        Exit Sub
    End If
End Sub

Private Sub mshBill_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    If InStr(1, "34", mint�༭״̬) <> 0 Then
        Cancel = True
        Exit Sub
    End If
    With mshBill
        If .TextMatrix(.Row, 0) <> "" Then
            If MsgBox("��ȷʵҪɾ������ҩƷ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True
            End If
            vsfStock.rows = 1
        End If
    End With
End Sub

Private Sub mshbill_CommandClick()
    Dim sngLeft As Single, sngTop As Single
    Dim RecReturn As Recordset
    Dim strUnit As String
    Dim i As Integer
    Dim intCurRow As Integer
    Dim intOldRow As Integer
    Dim lngRow As Long
    Dim strTemp As String
    
    intOldRow = mshBill.Row
    
    On Error GoTo errHandle
    If mshBill.Col = mconIntColҩ�� Then
        mblnChange = True
'        Set RecReturn = FrmҩƷѡ����.ShowME(Me, 1, , mlng�ⷿID)
        If grsMaster.State = adStateClosed Then
           Call SetSelectorRS(1, MStrCaption, mlng�ⷿID, mlng�ⷿID)
        End If
        Set RecReturn = frmSelector.ShowME(Me, 0, 1, , , , , mlng�ⷿID, , , , , , , , mstrPrivs)
        
        If RecReturn.RecordCount > 0 Then
            '����ظ���¼ �����ظ���¼��ҩƷid���ػ���
            If RecReturn.RecordCount = 1 Then
                lngRow = CheckDouData(RecReturn)
                If lngRow > 0 Then
                    If MsgBox("��ҩƷ�Ѿ����ڣ��Ƿ���ת����¼�У�", vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                        mshBill.Row = lngRow
                        mshBill.Col = 0
                        mshBill.SetFocus
                    End If
                    Exit Sub
                End If
            Else
                Set RecReturn = CheckData(RecReturn)
            End If
        End If
                
        If RecReturn.RecordCount > 0 Then
            RecReturn.MoveFirst
            For i = 1 To RecReturn.RecordCount
                With mshBill
                    intCurRow = .Row
                    Select Case mintUnit
                    Case 1
                        strUnit = "�ۼ۵�λ"
                    Case 2
                        strUnit = "���ﵥλ"
                    Case 3
                        strUnit = "סԺ��λ"
                    Case Else
                        strUnit = "ҩ�ⵥλ"
                    End Select
                    
                    SetDrugRows RecReturn!ҩƷid, "[" & RecReturn!ҩƷ���� & "]", RecReturn!ͨ����, IIf(IsNull(RecReturn!��Ʒ��), "", RecReturn!��Ʒ��), NVL(RecReturn!ҩƷ��Դ), _
                                IIf(IsNull(RecReturn!���), "", RecReturn!���), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
                                Switch(strUnit = "�ۼ۵�λ", RecReturn!�ۼ۵�λ, strUnit = "���ﵥλ", RecReturn!���ﵥλ, strUnit = "סԺ��λ", RecReturn!סԺ��λ, strUnit = "ҩ�ⵥλ", RecReturn!ҩ�ⵥλ), RecReturn!ָ��������, _
                                Switch(strUnit = "�ۼ۵�λ", 1, strUnit = "���ﵥλ", RecReturn!�����װ, strUnit = "סԺ��λ", RecReturn!סԺ��װ, _
                                strUnit = "ҩ�ⵥλ", RecReturn!ҩ���װ), RecReturn!ԭ����
                    If (.TextMatrix(intCurRow, 0) = "" Or intCurRow = 1 Or .Row = .rows - 1) And .TextMatrix(.rows - 1, 0) <> "" Then
                        .rows = .rows + 1
                    End If
                    .Row = .rows - 1
                End With
                RecReturn.MoveNext
            Next
            RecReturn.Close
            mshBill.Row = intCurRow
            mshBill.Col = mconintCol�ƻ�����
        End If
    Else
        'ҩƷ��Ӧ�̵�ѡ��
        sngLeft = mshBill.Left + mshBill.MsfObj.CellLeft
        sngTop = mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
        If sngLeft + Msf��Ӧ��ѡ��.Width > Me.ScaleWidth Then sngLeft = Me.ScaleWidth - Msf��Ӧ��ѡ��.Width - 100

        Set RecReturn = New ADODB.Recordset
        
        '1����ȫԺ�ƻ�ʱҪ��վ��
        'ȫԺ�ƻ�δ��ѡ����ʱҪ��վ��
        If mlng�ⷿID <> 0 Or (mlng�ⷿID = 0 And mintPlanPoint = 0 And (gstrNodeNo <> "-" Or gstrNodeNo <> "0")) Then
            strTemp = "(վ�� = [2] Or վ�� is Null) And "
        End If
        
        If mint��Ӧ�̷�Χ = 1 Then
            gstrSQL = "Select A.ID,A.����,A.����,A.���� From ��Ӧ�� A,ҩƷ�б굥λ B " & _
                      "Where " & strTemp & " A.ID=B.��λID And B.ҩƷID=[1] " & _
                      "  And (To_Char(B.����ʱ��,'yyyy-MM-dd')='3000-01-01' or B.����ʱ�� is null) " & _
                      "  And A.ĩ��=1 And (substr(A.����,1,1)=1 Or Nvl(A.ĩ��,0)=0) " & _
                      "  And (To_Char(A.����ʱ��,'yyyy-MM-dd')='3000-01-01' or A.����ʱ�� is null) " & _
                      "Order By A.���� "
        Else
            gstrSQL = "Select ID,����,����,���� From ��Ӧ�� " & _
                      "Where " & strTemp & " ĩ��=1 And (substr(����,1,1)=1 Or Nvl(ĩ��,0)=0) " & _
                      "  And (To_Char(����ʱ��,'yyyy-MM-dd')='3000-01-01' or ����ʱ�� is null) " & _
                      "Order By ���� "
        End If
        Set RecReturn = zlDataBase.OpenSQLRecord(gstrSQL, "��ȡҩӦ��", Val(mshBill.TextMatrix(mshBill.Row, 0)), gstrNodeNo)
        If RecReturn.RecordCount = 0 Then
            If mint��Ӧ�̷�Χ = 1 Then
                '���û�������б굥λ������ȡ���й�Ӧ��
                gstrSQL = "Select ID,����,����,���� From ��Ӧ�� " & _
                          "Where " & strTemp & " ĩ��=1 And (substr(����,1,1)=1 Or Nvl(ĩ��,0)=0) " & _
                          "  And (To_Char(����ʱ��,'yyyy-MM-dd')='3000-01-01' or ����ʱ�� is null) " & _
                          "Order By ���� "
                Set RecReturn = zlDataBase.OpenSQLRecord(gstrSQL, "��ȡҩӦ��", Val(mshBill.TextMatrix(mshBill.Row, 0)), gstrNodeNo)
                
                If RecReturn.RecordCount = 0 Then
                    MsgBox "���ȳ�ʼ��ҩƷ��Ӧ�̣�", vbInformation, gstrSysName
                    Exit Sub
                End If
            Else
                MsgBox "���ȳ�ʼ��ҩƷ��Ӧ�̣�", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        With Msf��Ӧ��ѡ��
            .Clear
            Set .DataSource = RecReturn
            .ColWidth(0) = 0
            .ColWidth(1) = 800
            .ColWidth(2) = 3000
            .ColWidth(3) = 800

            .Row = 1
            .ColSel = .Cols - 1
        End With
        With Msf��Ӧ��ѡ��
            .Left = sngLeft
            .Top = sngTop
            .Visible = True
            .ZOrder 0
            .SetFocus
        End With
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function CheckDouData(ByVal rsData As ADODB.Recordset) As Long
    '��������Ƿ��ظ�����Χ�ظ�����������
    Dim lngRow As Long
    
    With mshBill
        For lngRow = 1 To .rows - 1
            If .TextMatrix(lngRow, 0) = rsData!ҩƷid Then
                CheckDouData = lngRow
                Exit Function
            End If
        Next
    End With
End Function

Private Sub mshbill_EditChange(curText As String)
    mblnChange = True
End Sub


Private Sub mshBill_EditKeyPress(KeyAscii As Integer)
    Dim strkey As String
    Dim intDigit As Integer
    
    With mshBill
        strkey = .Text
        If strkey = "" Then
            strkey = .TextMatrix(.Row, .Col)
        End If
        
        If .Col = mconintCol�ƻ����� Or .Col = mconintCol�ɱ��� Then
            Select Case .Col
                Case mconintCol�ƻ�����
                    intDigit = mintShowNumberDigit
                Case mconintCol�ɱ���
                    intDigit = mintShowCostDigit
            End Select
            
            If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then
                If .SelLength = Len(strkey) Then Exit Sub
                If Len(Mid(strkey, InStr(1, strkey, ".") + 1)) >= intDigit And strkey Like "*.*" Then
                    KeyAscii = 0
                    Exit Sub
                Else
                    Exit Sub
                End If
            End If
        End If
        
    End With
End Sub

Private Sub mshbill_EnterCell(Row As Long, Col As Long)
    If Not mblnEnter Then Exit Sub

    With mshBill
        If Row > 0 Then
            .SetRowColor CLng(Row), &HFFCECE, True
        End If
        Select Case .Col
            Case mconIntColҩ��
                .txtCheck = False
                .MaxLength = 40
                'ֻ��ҩ���в���ʾ�ϼ���Ϣ�Ϳ����
                Call ��ʾ�ϼƽ��
            Case mconIntCol������
                .txtCheck = False
                .MaxLength = mlng�����̳���
            Case mconIntColԭ����
                .txtCheck = False
                .MaxLength = mlngԭ���س���
            Case mconintCol�ϴι�Ӧ��
                .MaxLength = 40
                .txtCheck = False
            Case mconintCol�ƻ�����, mconintColִ������
                .txtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"
            Case mconintCol�ɱ���, mconIntCol�ۼ�
                .txtCheck = True
                .MaxLength = 11
                .TextMask = ".1234567890"
            Case mconintCol˵��
                .txtCheck = False
                .MaxLength = 50
            Case mconIntCol��׼�ĺ�
                .txtCheck = False
                .MaxLength = 40
            Case mconintCol�ͻ�����
                .txtCheck = True
                .MaxLength = 10
                .TextMask = ".1234567890"
                If .TextMatrix(Row, mconintCol�ͻ���λ) = "" Then
                    .ColData(Col) = 5
                Else
                    If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Or mint�༭״̬ = 6 Then
                        .ColData(Col) = 4
                    Else
                        .ColData(Col) = 5
                    End If
                End If
        End Select
        
        If Row > 0 Then
            If .TextMatrix(Row, 0) <> "" Then
                Call ShowHisPlane(Row, Val(.TextMatrix(Row, 0)))
            End If
        End If
    End With
    
    Call ��ʾ���
End Sub

Private Sub mshbill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strkey As String
    Dim rsDrug As New Recordset
    Dim strUnit As String
    Dim strUnitQuantity As String
    Dim lngRow As Long
    
    Dim rsTemp As Recordset
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim intCurRow As Integer
    Dim intOldRow As Integer
    Dim i As Integer
    Dim strTemp As String
    
    On Error GoTo errHandle
    If KeyCode <> vbKeyReturn Then Exit Sub
    intOldRow = mshBill.Row
    
    With mshBill
        If .Col = mconIntColҩ�� Then
            .Text = UCase(Trim(.Text))
        Else
            .Text = Trim(.Text)
        End If
        strkey = .Text

        If Mid(strkey, 1, 1) = "[" Then
            If InStr(2, strkey, "]") <> 0 Then
                strkey = Mid(strkey, 2, InStr(2, strkey, "]") - 2)
            Else
                strkey = Mid(strkey, 2)
            End If
        End If
        Select Case .Col

            Case mconIntColҩ��
                If strkey <> "" Then

                    sngLeft = Me.Left + Pic����.Left + mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
                    sngTop = Me.Top + Me.Height - Me.ScaleHeight + Pic����.Top + mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
                    If sngTop + 3630 > Screen.Height Then
                        sngTop = sngTop - mshBill.MsfObj.CellHeight - 4530
                    End If

'                    Set rsTemp = FrmҩƷ��ѡѡ����.ShowME(Me, 1, , mlng�ⷿID, , strkey, sngLeft, sngTop)
                    If grsMaster.State = adStateClosed Then
                       Call SetSelectorRS(1, MStrCaption, mlng�ⷿID, mlng�ⷿID)
                    End If
                    Set rsTemp = frmSelector.ShowME(Me, 1, 1, strkey, sngLeft, sngTop, , mlng�ⷿID, , , , , , , , mstrPrivs)
                    
                    If rsTemp.RecordCount > 0 Then
                        '����ظ���¼ �����ظ���¼��ҩƷid���ػ���
                        If rsTemp.RecordCount = 1 Then
                            lngRow = CheckDouData(rsTemp)
                            If lngRow > 0 Then
                                If MsgBox("��ҩƷ�Ѿ����ڣ��Ƿ���ת����¼�У�", vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                                    mshBill.Row = lngRow
                                    mshBill.Col = 0
                                    mshBill.SetFocus
                                End If
                                Exit Sub
                            End If
                        Else
                            Set rsTemp = CheckData(rsTemp)
                        End If
                    End If
                    
                    If rsTemp.RecordCount > 0 Then
                        rsTemp.MoveFirst
                        For i = 1 To rsTemp.RecordCount
                            With mshBill
                                intCurRow = .Row
                                Select Case mintUnit
                                Case 1
                                    strUnit = "�ۼ۵�λ"
                                Case 2
                                    strUnit = "���ﵥλ"
                                Case 3
                                    strUnit = "סԺ��λ"
                                Case Else
                                    strUnit = "ҩ�ⵥλ"
                                End Select
                                Call SetDrugRows(rsTemp!ҩƷid, _
                                        "[" & rsTemp!ҩƷ���� & "]", _
                                        rsTemp!ͨ����, _
                                        IIf(IsNull(rsTemp!��Ʒ��), "", rsTemp!��Ʒ��), _
                                        IIf(IsNull(rsTemp!ҩƷ��Դ), "", rsTemp!ҩƷ��Դ), _
                                        IIf(IsNull(rsTemp!���), "", rsTemp!���), _
                                        IIf(IsNull(rsTemp!����), "", rsTemp!����), _
                                        Switch(strUnit = "�ۼ۵�λ", rsTemp!�ۼ۵�λ, strUnit = "���ﵥλ", rsTemp!���ﵥλ, _
                                               strUnit = "סԺ��λ", rsTemp!סԺ��λ, strUnit = "ҩ�ⵥλ", rsTemp!ҩ�ⵥλ), _
                                        rsTemp!ָ��������, _
                                        Switch(strUnit = "�ۼ۵�λ", 1, strUnit = "���ﵥλ", rsTemp!�����װ, strUnit = "סԺ��λ", _
                                               rsTemp!סԺ��װ, strUnit = "ҩ�ⵥλ", rsTemp!ҩ���װ), rsTemp!ԭ����)
                                .Text = .TextMatrix(.Row, .Col)
                                If (.TextMatrix(intCurRow, 0) = "" Or intCurRow = 1 Or .Row = .rows - 1) And .TextMatrix(.rows - 1, 0) <> "" Then
                                    .rows = .rows + 1
                                End If
                                .Row = .rows - 1
                            End With
                            rsTemp.MoveNext
                        Next
                        mshBill.Row = intCurRow
                        mshBill.Col = mconintCol�ƻ�����
                    Else
                        .Text = .TextMatrix(.Row, .Col)
                        Cancel = True
                    End If
                End If
            Case mconintCol�ƻ�����
                If Not IsNumeric(strkey) And strkey <> "" Then
                    MsgBox "�Բ��𣬼ƻ���������Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If .Text = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, mconintCol�ƻ�����) = ""
                    End If
                    If .TextMatrix(.Row, mconintCol�ƻ�����) <> "" Then
                        strkey = .TextMatrix(.Row, mconintCol�ƻ�����)
                        If .TextMatrix(.Row, mconintCol�ɱ���) <> "" Then
                            .TextMatrix(.Row, mconIntCol�ɱ����) = zlStr.FormatEx(.TextMatrix(.Row, mconintCol�ɱ���) * strkey, mintShowMoneyDigit, , True)
                        End If
                        If .TextMatrix(.Row, mconIntCol�ۼ�) <> "" Then
                            .TextMatrix(.Row, mconIntCol�ۼ۽��) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol�ۼ�) * strkey, mintShowMoneyDigit, , True)
                        End If
                    End If
                    .Col = mconintCol˵��
                    Cancel = True
                End If
                
                If strkey <> "" Then
                    strkey = zlStr.FormatEx(strkey, mintShowNumberDigit, , True)
                    If Val(.TextMatrix(.Row, mconintCol�ƻ�����)) <> Val(strkey) And Not mblnCheckRefresh Then
                        mblnCheckRefresh = True
                    End If
                    .Text = strkey
                    If .TextMatrix(.Row, mconintCol�ɱ���) <> "" Then
                        .TextMatrix(.Row, mconIntCol�ɱ����) = zlStr.FormatEx(.TextMatrix(.Row, mconintCol�ɱ���) * strkey, mintShowMoneyDigit, , True)
                    End If
                    If .TextMatrix(.Row, mconIntCol�ۼ�) <> "" Then
                        .TextMatrix(.Row, mconIntCol�ۼ۽��) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol�ۼ�) * strkey, mintShowMoneyDigit, , True)
                    End If
                    If Val(.TextMatrix(.Row, mconintCol�ͻ���װ)) <> 0 Then
                        .TextMatrix(.Row, mconintCol�ͻ�����) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol����ϵ��)) / Val(.TextMatrix(.Row, mconintCol�ͻ���װ)) * Val(strkey), 1)
                    End If
                End If
                
                Call ��ʾ�ϼƽ��
            Case mconintColִ������
                If Not IsNumeric(strkey) And strkey <> "" Then
                    MsgBox "�Բ��𣬼ƻ���������Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
            Case mconintCol�ɱ���
                If .TxtVisible = False Then Exit Sub
                If Not IsNumeric(strkey) And strkey <> "" Then
                    MsgBox "�Բ��𣬵��۱���Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If .Text = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, mconintCol�ɱ���) = ""
                    End If
                    Cancel = True
                    Exit Sub
                End If

                If strkey <> "" Then
                    strkey = zlStr.FormatEx(strkey, mintShowPriceDigit, , True)
                    .Text = strkey
                    If .TextMatrix(.Row, mconintCol�ƻ�����) <> "" Then
                        .TextMatrix(.Row, mconIntCol�ɱ����) = zlStr.FormatEx(.TextMatrix(.Row, mconintCol�ƻ�����) * strkey, mintShowMoneyDigit, , True)
                    End If

                End If
                Call ��ʾ�ϼƽ��
            Case mconIntCol������
                If strkey = "" And .TextMatrix(.Row, mconIntCol������) = "" Then
                    strkey = " "
                    .Text = strkey
                    .TextMatrix(.Row, mconIntCol������) = strkey
                Else
                    If zlCommFun.StrIsValid(strkey, mlng�����̳���) = False Then
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                End If
            Case mconIntColԭ����
                If strkey = "" And .TextMatrix(.Row, mconIntColԭ����) = "" Then
                    strkey = " "
                    .Text = strkey
                    .TextMatrix(.Row, mconIntColԭ����) = strkey
                Else
                    If zlCommFun.StrIsValid(strkey, mlngԭ���س���) = False Then
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                End If
            Case mconintCol�ϴι�Ӧ��
                If .TxtVisible = False Then Exit Sub
                If strkey = "" Then
                    strkey = " "
                    .Text = strkey
                    .TextMatrix(.Row, mconintCol�ϴι�Ӧ��) = strkey
                Else
                    If zlCommFun.StrIsValid(strkey, 40) = False Then
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    strkey = UCase(strkey)
                    sngLeft = mshBill.Left + mshBill.MsfObj.CellLeft
                    sngTop = mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
                    If sngLeft + Msf��Ӧ��ѡ��.Width > Me.ScaleWidth Then sngLeft = Me.ScaleWidth - Msf��Ӧ��ѡ��.Width - 100
                    
                    '1����ȫԺ�ƻ�ʱҪ��վ��
                    'ȫԺ�ƻ�δ��ѡ����ʱҪ��վ��
                    If mlng�ⷿID <> 0 Or (mlng�ⷿID = 0 And mintPlanPoint = 0 And (gstrNodeNo <> "-" Or gstrNodeNo <> "0")) Then
                        strTemp = "(A.վ�� = '" & gstrNodeNo & "' Or A.վ�� is Null) And "
                    End If
                    
                    If mint��Ӧ�̷�Χ = 1 Then
                        gstrSQL = "Select A.ID,A.����,A.����,A.���� From ��Ӧ�� A,ҩƷ�б굥λ B Where " & strTemp & " A.ĩ��=1 And (substr(A.����,1,1)=1 Or Nvl(A.ĩ��,0)=0) ANd (To_Char(A.����ʱ��,'yyyy-MM-dd')='3000-01-01' or A.����ʱ�� is null) " & _
                            " And A.ID=B.��λID And B.ҩƷID=[2] And (To_Char(B.����ʱ��,'yyyy-MM-dd')='3000-01-01' or B.����ʱ�� is null) " & _
                            " And (upper(A.����) Like [1] Or Upper(A.����) Like [1] Or Upper(A.����) Like [1]) " & _
                            " Order By A.���� "
                    Else
                        gstrSQL = "Select A.ID,A.����,A.����,A.���� From ��Ӧ�� A Where  " & strTemp & " A.ĩ��=1 And (substr(A.����,1,1)=1 Or Nvl(A.ĩ��,0)=0) ANd (To_Char(A.����ʱ��,'yyyy-MM-dd')='3000-01-01' or A.����ʱ�� is null) " & _
                            " And (upper(A.����) Like [1] Or Upper(A.����) Like [1] Or Upper(A.����) Like [1]) " & _
                            " Order By A.���� "
                    End If
                    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[��ȡҩӦ��]", IIf(gstrMatchMethod = "0", "%", "") & strkey & "%", Val(mshBill.TextMatrix(mshBill.Row, 0)))
                    
                    If rsTemp.RecordCount = 0 Then
                        If mint��Ӧ�̷�Χ = 1 Then
                            '���û�������б굥λ������ȡ���й�Ӧ��
                            gstrSQL = "Select A.ID,A.����,A.����,A.���� From ��Ӧ�� A Where  " & strTemp & " A.ĩ��=1 And (substr(A.����,1,1)=1 Or Nvl(A.ĩ��,0)=0) ANd (To_Char(A.����ʱ��,'yyyy-MM-dd')='3000-01-01' or A.����ʱ�� is null) " & _
                                " And (upper(A.����) Like [1] Or Upper(A.����) Like [1] Or Upper(A.����) Like [1]) " & _
                                " Order By A.���� "
                            Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[��ȡҩӦ��]", IIf(gstrMatchMethod = "0", "%", "") & strkey & "%", Val(mshBill.TextMatrix(mshBill.Row, 0)))
                            
                            If rsTemp.RecordCount = 0 Then
                                MsgBox "û���ҵ����������Ĺ�Ӧ�̣�", vbInformation, gstrSysName
                                Cancel = True
                                .TxtSetFocus
                                Exit Sub
                            ElseIf rsTemp.RecordCount = 1 Then
                                .Text = rsTemp!����
                                Exit Sub
                            End If
                        Else
                            MsgBox "û���ҵ����������Ĺ�Ӧ�̣�", vbInformation, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If

                    ElseIf rsTemp.RecordCount = 1 Then
                        .Text = rsTemp!����
                        Exit Sub
                    End If
                    
                    With Msf��Ӧ��ѡ��
                        .Clear
                        Set .DataSource = rsTemp
                        .ColWidth(0) = 0
                        .ColWidth(1) = 800
                        .ColWidth(2) = 3000
                        .ColWidth(3) = 800
            
                        .Row = 1
                        .ColSel = .Cols - 1
                    End With
                    With Msf��Ӧ��ѡ��
                        .Left = sngLeft
                        .Top = sngTop
                        .Visible = True
                        .ZOrder 0
                        .SetFocus
                    End With
                    Cancel = True
                End If
            Case mconintCol˵��
                If strkey = "" And .TextMatrix(.Row, mconintCol˵��) = "" Then
                    strkey = " "
                    If Trim(.TextMatrix(.Row, mconintCol˵��)) <> Trim(strkey) And Not mblnCheckRefresh Then
                        mblnCheckRefresh = True
                    End If
                    .Text = strkey
                    .TextMatrix(.Row, mconintCol˵��) = strkey
                Else
                    If zlCommFun.StrIsValid(strkey, 50) = False Then
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                End If
            Case mconIntCol��׼�ĺ�
                If strkey = "" And .TextMatrix(.Row, mconIntCol��׼�ĺ�) = "" Then
                    strkey = " "
                    .Text = strkey
                    .TextMatrix(.Row, mconIntCol��׼�ĺ�) = strkey
                Else
                    If zlCommFun.StrIsValid(strkey, 40) = False Then
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                End If
        End Select
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mshBill_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        With mshBill
           If .Col = mconIntColҩ�� Then
                PopupMenu mnuCol, 2
            End If
        End With
    End If
End Sub

Private Sub picHscSend_Click()
    If Val(picHscSend.Tag) = "1" Then
        Call imgDown_Click
    Else
        Call imgUp_Click
    End If
End Sub

Private Sub staThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Key = "PY" And staThis.Tag <> "PY" Then
        Logogram staThis, 0
        staThis.Tag = Panel.Key
    ElseIf Panel.Key = "WB" And staThis.Tag <> "WB" Then
        Logogram staThis, 1
        staThis.Tag = Panel.Key
    End If
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
    If KeyAscii = 13 Then
        cmdFind_Click
    End If
End Sub

Private Function ValidData() As Boolean
    ValidData = False
    Dim intLop As Integer

    With mshBill
        If .TextMatrix(1, 0) <> "" Then         '�����з�����

            If LenB(StrConv(txtժҪ.Text, vbFromUnicode)) > 40 Then
                MsgBox "ժҪ����,���������20�����ֻ�40���ַ�!", vbInformation + vbOKOnly, gstrSysName
                txtժҪ.SetFocus
                Exit Function
            End If

            For intLop = 1 To .rows - 1
                If Trim(.TextMatrix(intLop, mconIntColҩ��)) <> "" Then
                    If Trim(Trim(.TextMatrix(intLop, mconintCol�ƻ�����))) <> "" Then
                        If Not IsNumeric(.TextMatrix(intLop, mconintCol�ƻ�����)) Then
                            MsgBox "��" & intLop & "��ҩƷ�ļƻ�������Ϊ�����ͣ����飡", vbInformation, gstrSysName
                            mshBill.SetFocus
                            .Row = intLop
                            .MsfObj.TopRow = intLop
                            .Col = mconintCol�ƻ�����
                            Exit Function
                        End If

                    End If
                    If Val(.TextMatrix(intLop, mconintCol�ƻ�����)) > 9999999999# Then
                        MsgBox "��" & intLop & "��ҩƷ�ļƻ��������������ݿ��ܹ������" & vbCrLf & "���Χ9999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconintCol�ƻ�����
                        Exit Function
                    End If

                    If Val(.TextMatrix(intLop, mconIntCol�ɱ����)) > 9999999999999# Then
                        MsgBox "��" & intLop & "��ҩƷ�Ľ����������ݿ��ܹ������" & vbCrLf & "���Χ9999999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconintCol�ƻ�����
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, mconIntCol�ۼ۽��)) > 9999999999999# Then
                        MsgBox "��" & intLop & "��ҩƷ���ۼ۽����������ݿ��ܹ������" & vbCrLf & "���Χ9999999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconintCol�ƻ�����
                        Exit Function
                    End If
                End If
            Next
        Else
            Exit Function
        End If
    End With

    ValidData = True
End Function

Private Function SaveCard() As Boolean
    Dim lng��� As Long
    Dim ID_IN As Long
    Dim NO_IN As Variant
    Dim �ƻ�����_IN As Integer
    Dim �ڼ�_IN As String
    Dim �ⷿID_IN As Long
    Dim ���Ʒ���_IN As Integer
    Dim ������_IN As String
    Dim ��������_IN As String
    Dim ����˵��_IN As String

    Dim ҩƷID_IN As Long
    Dim �ƻ�����_IN As Double
    Dim ����_IN As Double
    Dim ���_IN As Double
    Dim ǰ������_IN As Double
    Dim ��������_IN As Double
    Dim �������_IN As Double
    Dim �ϴι�Ӧ��_IN As String
    Dim �ϴ�������_IN As String
    Dim ˵��_IN As String
    Dim intRow As Integer
    Dim �ۼ�_IN As Double
    Dim �ۼ۽��_IN As Double
    Dim ��������_IN As Double
    Dim ��������_IN As Double
    Dim ҩ��ID_IN As Double
    Dim i As Integer
    Dim arrSql As Variant
    Dim �ͻ�����_in As Double
    Dim ��׼�ĺ�_IN As String
    
    SaveCard = False
    arrSql = Array()

    On Error GoTo errHandle
    With mshBill
        ID_IN = Sys.NextId("ҩƷ�ɹ��ƻ�")
        NO_IN = Trim(txtNo)
        If NO_IN = "" Then NO_IN = Sys.GetNextNo(32, mlng�ⷿID)
         
        If IsNull(NO_IN) Then Exit Function
        Me.txtNo.Tag = NO_IN
        �ƻ�����_IN = mint�ƻ�����
        ���Ʒ���_IN = mint���Ʒ���
        �ⷿID_IN = mlng�ⷿID
        ������_IN = UserInfo.�û�����
        ��������_IN = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        ����˵��_IN = Trim(txtժҪ.Text)
        �ڼ�_IN = Str�ڼ�
        ҩ��ID_IN = Val(txtժҪ.Tag)
        
        If mint�༭״̬ = 2 Or (mint�༭״̬ = 3 And mblnCheckRefresh) Then      '�޸�
            gstrSQL = "zl_ҩƷ�ƻ�����_DELETE('" & mlng�ƻ�ID & "')"
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = gstrSQL
        End If

        gstrSQL = "zl_ҩƷ�ƻ���������_INSERT("
        '�ƻ�ID
        gstrSQL = gstrSQL & ID_IN
        'NO
        gstrSQL = gstrSQL & ",'" & NO_IN & "'"
        '�ƻ�����
        gstrSQL = gstrSQL & "," & �ƻ�����_IN
        '�ڼ�
        gstrSQL = gstrSQL & ",'" & �ڼ�_IN & "'"
        '�ⷿID
        gstrSQL = gstrSQL & "," & IIf(�ⷿID_IN = 0, "Null", �ⷿID_IN)
        'ҩ��ID
        gstrSQL = gstrSQL & "," & IIf(ҩ��ID_IN = 0, "Null", ҩ��ID_IN)
        '���Ʒ���
        gstrSQL = gstrSQL & "," & ���Ʒ���_IN
        '������
        gstrSQL = gstrSQL & ",'" & ������_IN & "'"
        '��������
        gstrSQL = gstrSQL & ",to_date('" & ��������_IN & "','yyyy-mm-dd HH24:MI:SS')"
        '����˵��
        gstrSQL = gstrSQL & ",'" & ����˵��_IN & "'"
        '��Դ�ⷿID
        gstrSQL = gstrSQL & ",'" & IIf(mstr��Դ�ⷿ = "", IIf(mlng�ⷿID = 0, mstrAll��Դ�ⷿ, mlng�ⷿID), mstr��Դ�ⷿ) & "'"
        '��Դҩ��ID
        gstrSQL = gstrSQL & ",'" & IIf(mstr��Դҩ�� = "", mstrAll��Դҩ��, mstr��Դҩ��) & "'"
        gstrSQL = gstrSQL & ")"

        ReDim Preserve arrSql(UBound(arrSql) + 1)
        arrSql(UBound(arrSql)) = gstrSQL

        For intRow = 1 To .rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                lng��� = .TextMatrix(intRow, mconIntCol���)
                ҩƷID_IN = .TextMatrix(intRow, 0)
                
                ����_IN = .TextMatrix(intRow, mconintCol�ɱ���) / Val(.TextMatrix(intRow, mconIntCol����ϵ��))
                ���_IN = Val(.TextMatrix(intRow, mconIntCol�ɱ����))
                �ۼ�_IN = .TextMatrix(intRow, mconIntCol�ۼ�) / Val(.TextMatrix(intRow, mconIntCol����ϵ��))
                �ۼ۽��_IN = Val(.TextMatrix(intRow, mconIntCol�ۼ۽��))
            
                ǰ������_IN = Val(.TextMatrix(intRow, mconIntColǰ������)) * Val(.TextMatrix(intRow, mconIntCol����ϵ��))
                ��������_IN = Val(.TextMatrix(intRow, mconIntCol��������)) * Val(.TextMatrix(intRow, mconIntCol����ϵ��))
                �������_IN = Val(.TextMatrix(intRow, mconintCol�������)) * Val(.TextMatrix(intRow, mconIntCol����ϵ��))
                ��������_IN = Val(.TextMatrix(intRow, mconintCol��������)) * Val(.TextMatrix(intRow, mconIntCol����ϵ��))
                ��������_IN = Val(.TextMatrix(intRow, mconintCol��������)) * Val(.TextMatrix(intRow, mconIntCol����ϵ��))
                �ƻ�����_IN = Val(.TextMatrix(intRow, mconintCol�ƻ�����)) * Val(.TextMatrix(intRow, mconIntCol����ϵ��))
                �ϴι�Ӧ��_IN = .TextMatrix(intRow, mconintCol�ϴι�Ӧ��)
                �ϴ�������_IN = .TextMatrix(intRow, mconIntCol������)
                ˵��_IN = .TextMatrix(intRow, mconintCol˵��)
                �ͻ�����_in = Val(.TextMatrix(intRow, mconintCol�ͻ�����))
                ��׼�ĺ�_IN = .TextMatrix(intRow, mconIntCol��׼�ĺ�)
                
                gstrSQL = "zl_ҩƷ�ƻ�����α�_INSERT("
                '�ƻ�ID
                gstrSQL = gstrSQL & ID_IN
                'ҩƷID
                gstrSQL = gstrSQL & "," & ҩƷID_IN
                '���
                gstrSQL = gstrSQL & "," & lng���
                '�ƻ�����
                gstrSQL = gstrSQL & "," & �ƻ�����_IN
                '����
                gstrSQL = gstrSQL & "," & ����_IN
                '���
                gstrSQL = gstrSQL & "," & ���_IN
                'ǰ������
                gstrSQL = gstrSQL & "," & ǰ������_IN
                '��������
                gstrSQL = gstrSQL & "," & ��������_IN
                '�������
                gstrSQL = gstrSQL & "," & �������_IN
                '��Ӧ��
                gstrSQL = gstrSQL & ",'" & �ϴι�Ӧ��_IN & "'"
                '������
                gstrSQL = gstrSQL & ",'" & �ϴ�������_IN & "'"
                '˵��
                gstrSQL = gstrSQL & ",'" & ˵��_IN & "'"
                '�ۼ�
                gstrSQL = gstrSQL & "," & �ۼ�_IN
                '�ۼ۽��
                gstrSQL = gstrSQL & "," & �ۼ۽��_IN
                '��������
                gstrSQL = gstrSQL & "," & ��������_IN
                '��������
                gstrSQL = gstrSQL & "," & ��������_IN
                '�ͻ�����
                gstrSQL = gstrSQL & "," & �ͻ�����_in
                '��׼�ĺ�
                gstrSQL = gstrSQL & ",'" & ��׼�ĺ�_IN & "'"
                gstrSQL = gstrSQL & ")"

                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
            End If
        Next
        gcnOracle.BeginTrans
        For i = 0 To UBound(arrSql)
            Call zlDataBase.ExecuteProcedure(CStr(arrSql(i)), "SaveCard")
        Next
        gcnOracle.CommitTrans
        
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    
    If mint�༭״̬ = 3 And mblnCheckRefresh Then
        mlng�ƻ�ID = ID_IN
    End If
        
    SaveCard = True
    vsfStock.rows = 1
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    'MsgBox "����ʧ�ܣ����飡", vbInformation + vbOKOnly, gstrSysName
    Call SaveErrLog
End Function

Private Sub ��ʾ�ϼƽ��()
    Dim Dbl��� As Double, dbl�ۼ۽�� As Double
    Dim intLop As Integer

    Dbl��� = 0: dbl�ۼ۽�� = 0

    With mshBill
        For intLop = 1 To .rows - 1
            If .TextMatrix(intLop, 0) <> "" Then
                Dbl��� = Dbl��� + Val(.TextMatrix(intLop, mconIntCol�ɱ����))
                dbl�ۼ۽�� = dbl�ۼ۽�� + Val(.TextMatrix(intLop, mconIntCol�ۼ۽��))
            End If
        Next
    End With
    If mint�۸���ʾ = 0 Then
        lblPurchasePrice.Caption = "���ϼƣ�" & zlStr.FormatEx(Dbl���, mintShowMoneyDigit, , True)
    ElseIf mint�۸���ʾ = 1 Then
        lblPurchasePrice.Caption = "���ϼƣ�" & zlStr.FormatEx(dbl�ۼ۽��, mintShowMoneyDigit, , True)
    Else
        lblPurchasePrice.Caption = "�ɱ����ϼƣ�" & zlStr.FormatEx(Dbl���, mintShowMoneyDigit, , True) & "      �ۼ۽��ϼƣ�" & zlStr.FormatEx(dbl�ۼ۽��, mintShowMoneyDigit, , True)
    End If
End Sub


Private Sub ��ʾ���()
    Dim rsData As ADODB.Recordset
    Dim lngҩƷid As Long
    Dim str��λ As String
    Dim dbl��װ As Double
    Dim strSql As String
    
    If mblnStart = False Then Exit Sub
    
    On Error GoTo errHandle
    Me.staThis.Panels(2).Text = ""
    If mshBill.TextMatrix(mshBill.Row, 0) = "" Then vsfStock.rows = 1: Exit Sub
    
    lngҩƷid = Val(mshBill.TextMatrix(mshBill.Row, 0))
    str��λ = mshBill.TextMatrix(mshBill.Row, mconIntCol��λ)
    dbl��װ = Val(mshBill.TextMatrix(mshBill.Row, mconIntCol����ϵ��))
    vsfStock.Tag = 0
    vsfStock.rows = 1
    
    If txtNo <> "" Then
        strSql = "Select t.��Դ�ⷿ, t.��Դҩ�� From ҩƷ�ɹ��ƻ� T Where t.No =[1]"
        Set rsData = zlDataBase.OpenSQLRecord(strSql, "", txtNo)
        
        mstr��Դ�ⷿ = IIf(NVL(rsData!��Դ�ⷿ, 0) = 0, "", NVL(rsData!��Դ�ⷿ))
        mstr��Դҩ�� = NVL(rsData!��Դҩ��)
    End If
    
    If chk��Դҩ��.Value = 1 Or chk��Դ�ⷿ.Value = 1 Or chk���пⷿ.Value = 1 Or mstr�Զ���ⷿ <> "" Then
        
        gstrSQL = "Select B.����, A.ҩƷid, Nvl(Sum(A.��������),0) As ��������, Nvl(Sum(A.ʵ������),0) As ʵ������ " & _
            " From ҩƷ��� A, ���ű� B " & _
            " Where A.���� = 1 And A.�ⷿid + 0 = B.ID And A.ҩƷid = [1] "
            
        If chk���пⷿ.Value = 0 Then
            If chk��Դ�ⷿ.Value = 1 And chk��Դҩ��.Value = 1 And mstr�Զ���ⷿ <> "" And (mstr��Դ�ⷿ <> "" Or mstrAll��Դ�ⷿ <> "") _
                                                                                                                                     And (mstr��Դҩ�� <> "" Or mstrAll��Դҩ�� <> "") Then
                gstrSQL = gstrSQL & " and ( A.�ⷿid In(select * from Table(Cast(f_Num2List([2]) As zlTools.t_NumList))) " & _
                                                 " or A.�ⷿid In(select * from Table(Cast(f_Num2List([3]) As zlTools.t_NumList))) " & _
                                                 " or A.�ⷿid In(select * from Table(Cast(f_Num2List([4]) As zlTools.t_NumList))) )"
                
            ElseIf chk��Դ�ⷿ.Value = 1 And chk��Դҩ��.Value = 1 And mstr�Զ���ⷿ = "" And (mstr��Դ�ⷿ <> "" Or mstrAll��Դ�ⷿ <> "") _
                                                                                                                                        And (mstr��Դҩ�� <> "" Or mstrAll��Դҩ�� <> "") Then
                gstrSQL = gstrSQL & " and ( A.�ⷿid In(select * from Table(Cast(f_Num2List([2]) As zlTools.t_NumList))) " & _
                                                                 " or A.�ⷿid In(select * from Table(Cast(f_Num2List([3]) As zlTools.t_NumList))) )"
            
            ElseIf chk��Դ�ⷿ.Value = 1 And chk��Դҩ��.Value = 0 And mstr�Զ���ⷿ <> "" And (mstr��Դ�ⷿ <> "" Or mstrAll��Դ�ⷿ <> "") Then
                gstrSQL = gstrSQL & " and ( A.�ⷿid In(select * from Table(Cast(f_Num2List([2]) As zlTools.t_NumList))) " & _
                                                                 " or A.�ⷿid In(select * from Table(Cast(f_Num2List([4]) As zlTools.t_NumList))) )"
                                                                          
            ElseIf chk��Դ�ⷿ.Value = 0 And chk��Դҩ��.Value = 1 And mstr�Զ���ⷿ <> "" And (mstr��Դҩ�� <> "" Or mstrAll��Դҩ�� <> "") Then
                gstrSQL = gstrSQL & " and ( A.�ⷿid In(select * from Table(Cast(f_Num2List([3]) As zlTools.t_NumList))) " & _
                                                                 " or A.�ⷿid In(select * from Table(Cast(f_Num2List([4]) As zlTools.t_NumList))) )"
                                                                                                          
            ElseIf chk��Դ�ⷿ.Value = 1 And chk��Դҩ��.Value = 0 And mstr�Զ���ⷿ = "" And (mstr��Դ�ⷿ <> "" Or mstrAll��Դ�ⷿ <> "") Then
                gstrSQL = gstrSQL & " and A.�ⷿid In(select * from Table(Cast(f_Num2List([2]) As zlTools.t_NumList))) "
                
            ElseIf chk��Դ�ⷿ.Value = 0 And chk��Դҩ��.Value = 1 And mstr�Զ���ⷿ = "" And (mstr��Դҩ�� <> "" Or mstrAll��Դҩ�� <> "") Then
                gstrSQL = gstrSQL & " and A.�ⷿid In(select * from Table(Cast(f_Num2List([3]) As zlTools.t_NumList))) "
                
            ElseIf chk��Դ�ⷿ.Value = 0 And chk��Դҩ��.Value = 0 And mstr�Զ���ⷿ <> "" Then
                gstrSQL = gstrSQL & " and A.�ⷿid In(select * from Table(Cast(f_Num2List([4]) As zlTools.t_NumList))) "
                
            End If
        End If
        
        gstrSQL = gstrSQL & " Group By B.����, A.ҩƷid " & _
        " Order By B.����"
        
        Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, "��ʾ���", lngҩƷid, _
                                                                        IIf(mstr��Դ�ⷿ = "", IIf(mlng�ⷿID = 0, mstrAll��Դ�ⷿ, mlng�ⷿID), mstr��Դ�ⷿ), _
                                                                        IIf(mstr��Դҩ�� = "", mstrAll��Դҩ��, mstr��Դҩ��), mstr�Զ���ⷿ)
        
        Do While Not rsData.EOF
            vsfStock.rows = vsfStock.rows + 1
            vsfStock.TextMatrix(vsfStock.rows - 1, vsfStock.ColIndex("�ⷿ")) = rsData!����
            vsfStock.TextMatrix(vsfStock.rows - 1, vsfStock.ColIndex("��������")) = zlStr.FormatEx(rsData!�������� / dbl��װ, mintShowNumberDigit, , True)
            vsfStock.TextMatrix(vsfStock.rows - 1, vsfStock.ColIndex("ʵ������")) = zlStr.FormatEx(rsData!ʵ������ / dbl��װ, mintShowNumberDigit, , True)
            
            rsData.MoveNext
        Loop
    End If
            
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtժҪ_Change()
    mblnChange = True
End Sub

Private Sub txtժҪ_GotFocus()
    OS.OpenIme True
    With txtժҪ
        .SelStart = 0
        .SelLength = Len(txtժҪ.Text)
    End With
End Sub

Private Sub txtժҪ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey (vbKeyTab)
        KeyCode = 0
    End If
End Sub

Private Sub txtժҪ_LostFocus()
    OS.OpenIme
End Sub

Private Function SetDrugRows(ByVal lngҩƷid As Long, ByVal strҩƷ���� As String, ByVal strͨ���� As String, ByVal str��Ʒ�� As String, ByVal strҩƷ��Դ As String, _
        ByVal str��� As String, ByVal str���� As String, ByVal str��λ As String, _
        ByVal dblָ�������� As Double, ByVal dbl����ϵ�� As Double, ByVal strԭ���� As String) As Boolean
    Dim rsData As New Recordset
    Dim intCount As Integer
    Dim intRow As Integer
    Dim intCol As Integer

    Dim lng���� As Long
    Dim dbl������� As Double
    Dim dbl�ɱ����� As Double, dbl���۵��� As Double
    Dim rs��ͬ��λ As ADODB.Recordset
    Dim strҩ�� As String
    Dim rsTemp As ADODB.Recordset
    Dim dbl�ͻ���װ As Double
    Dim str�ͻ���λ As String
    Dim str��Ӧ�� As String
    Dim str��׼�ĺ� As String
    
    On Error GoTo errH
    SetDrugRows = False

    With mshBill
        .TextMatrix(.Row, mconIntCol���) = .Row
        .TextMatrix(.Row, mconIntCol������) = str����
        .TextMatrix(.Row, mconIntColԭ����) = strԭ����
        .TextMatrix(.Row, 0) = lngҩƷid
        .TextMatrix(.Row, mconIntCol����ϵ��) = dbl����ϵ��
        
        gstrSQL = "Select a.�ɱ���,a.ָ��������,b.���� as ��Ӧ��,nvl(a.�ϴ���׼�ĺ�,a.��׼�ĺ�) as ��׼�ĺ� From ҩƷ��� a ,��Ӧ�� b Where a.�ϴι�Ӧ��id=b.id(+) and a.ҩƷID=[1]"
        Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[ȡƽ���ɱ���]", lngҩƷid)
        dbl�ɱ����� = NVL(rsData!�ɱ���, 0)
        If dbl�ɱ����� = 0 Then dbl�ɱ����� = NVL(rsData!ָ��������, 0)
        str��Ӧ�� = NVL(rsData!��Ӧ��, "")
        str��׼�ĺ� = NVL(rsData!��׼�ĺ�, "")
        
        gstrSQL = "Select Decode(Nvl(a.�Ƿ���, 0), 0, Nvl(b.�ּ�, 0), Decode(nvl(d.�ϴ��ۼ�,0), 0, Decode(Nvl(c.ƽ���ۼ�, 0), 0, b.�ּ�, c.ƽ���ۼ�), d.�ϴ��ۼ�)) �ۼ� " & _
                 " From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ҩƷ��� D, " & _
                 " (Select ҩƷid, " & _
                 " Decode(Sign(Sum(ʵ������)), 1, Decode(Sign(Sum(ʵ�ʽ��)), 1, Sum(ʵ�ʽ��), 0) / Sum(ʵ������), 0) ƽ���ۼ� " & _
                 " From ҩƷ��� " & _
                 " Where ���� = 1 " & IIf(mlng�ⷿID = 0, "", " AND �ⷿID=[2] ") & " And ҩƷid = [1] " & _
                 " Group By ҩƷid) C " & _
                 " Where A.ID = B.�շ�ϸĿid And A.ID = C.ҩƷid(+) And A.ID = D.ҩƷid And A.ID = [1] And " & _
                 " (A.����ʱ�� >= To_Date('3000-01-01 00:00:00', 'YYYY-MM-DD HH24:MI:SS') Or A.����ʱ�� Is Null) And " & _
                 " (B.��ֹ���� Is Null Or Sysdate Between B.ִ������ And Nvl(B.��ֹ����, To_Date('3000-01-01', 'yyyy-MM-dd'))) " & _
                 GetPriceClassString("B")
                 
        Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[ȡ�ۼ�]", lngҩƷid, mlng�ⷿID)
        dbl���۵��� = rsData!�ۼ�
        
        If mlng�ⷿID = 0 Then
            '�����ȫ�ⷿ�����ҩƷ�����ȡ�����������ҩƷ�����ȡ��Ӧ�̣��ϴβ���
            gstrSQL = "Select B.���� ��Ӧ��, C.�ϴβ���, C.ԭ����,Nvl(A.�������,0) �������" & _
                      " From (Select ҩƷid, Sum(ʵ������) As ������� From ҩƷ��� " & _
                      " Where ���� = 1 And ҩƷid = [1] Group By ҩƷid) A, " & _
                      " (Select id,���� From ��Ӧ�� Where Substr(����, 1, 1) = 1) B, ҩƷ��� C " & _
                      " Where C.ҩƷid = A.ҩƷid(+) And C.�ϴι�Ӧ��id = B.ID(+) And C.ҩƷid = [1] "
            Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[ȡ�ϴι�Ӧ�̼�������Ϣ]", lngҩƷid, mlng�ⷿID)
        Else
            '���ָ���ⷿ�����ָ���ⷿȡ�����������������εĹ�Ӧ�̣��ϴβ���
            gstrSQL = " Select C.���� As ��Ӧ��, B.�ϴβ���, B.ԭ����,A.������� " & _
                      " From (Select �ⷿid, ҩƷid, Sum(ʵ������) As ������� " & _
                      " From ҩƷ��� " & _
                      " Where ���� = 1 And ҩƷid = [1] And �ⷿID=[2] " & _
                      " Group By �ⷿid, ҩƷid) A, " & _
                      " (Select �ⷿid,ҩƷid,�ϴι�Ӧ��ID,�ϴβ���,ԭ���� From ҩƷ��� " & _
                      " Where ���� = 1 And ҩƷid = [1] And �ⷿID=[2] " & _
                      " And Nvl(����, 0) = " & _
                      " (Select Nvl(Max(Nvl(����, 0)), 0) ���� From ҩƷ��� Where ���� = 1 And ҩƷid = [1] And �ⷿID=[2] )) B, " & _
                      " (SELECT id,���� FROM ��Ӧ�� WHERE SUBSTR(����,1,1)=1) C " & _
                      " Where A.�ⷿid = B.�ⷿid And A.ҩƷid = B.ҩƷid And B.�ϴι�Ӧ��id = C.ID(+) " & _
                      " And A.ҩƷid = [1] And A.�ⷿID=[2] "
            Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[ȡ�ϴι�Ӧ�̼�������Ϣ]", lngҩƷid, mlng�ⷿID)
            
            '���ָ���ⷿ�޿�棬���ҩƷ�����ȡ��Ӧ�̣��ϴβ���
            If rsData.RecordCount = 0 Then
                gstrSQL = "Select B.���� ��Ӧ��, C.�ϴβ���, C.ԭ����, 0 ������� from " & _
                          " (Select ID,���� From ��Ӧ�� Where Substr(����, 1, 1) = 1) B, ҩƷ��� C " & _
                          " Where C.�ϴι�Ӧ��id = B.ID(+) And ҩƷid = [1] "
                Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[ȡ�ϴι�Ӧ�̼�������Ϣ]", lngҩƷid, mlng�ⷿID)
            End If
        End If
       
        If Not rsData.EOF Then
            .TextMatrix(.Row, mconintCol�������) = zlStr.FormatEx(IIf(IsNull(rsData!�������), 0, rsData!�������) / dbl����ϵ��, mintShowNumberDigit, , True)
            
            .TextMatrix(.Row, mconintCol�ϴι�Ӧ��) = IIf(IsNull(rsData!��Ӧ��), str��Ӧ��, rsData!��Ӧ��)
            If mint��Ӧ��ѡ�� = 1 Then
                gstrSQL = "Select B.���� From ҩƷ��� A, ��Ӧ�� B Where A.��ͬ��λid = B.ID And A.ҩƷid = [1] "
                Set rs��ͬ��λ = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lngҩƷid)
                If Not rs��ͬ��λ.EOF Then
                    .TextMatrix(.Row, mconintCol�ϴι�Ӧ��) = rs��ͬ��λ!����
                End If
            End If
            
            .TextMatrix(.Row, mconIntCol������) = IIf(IsNull(rsData!�ϴβ���), str����, rsData!�ϴβ���)
            .TextMatrix(.Row, mconIntColԭ����) = IIf(IsNull(rsData!ԭ����), str����, rsData!ԭ����)
            SetNumer lngҩƷid, mlng�ⷿID, .TextMatrix(.Row, mconintCol�������), .Row, mint�ƻ�����, mint���Ʒ���, mbln������ʽ
        End If
        
        '���ش��װ�����Ϣ
        gstrSQL = "select a.�ͻ���λ,a.�ͻ���װ,a.����ҩ��,b.�������� from ҩƷ��� a,�շ���ĿĿ¼ b where a.ҩƷid=b.id and a.ҩƷid=[1]"
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "�����Ϣ", lngҩƷid)
        dbl�ͻ���װ = IIf(IsNull(rsTemp!�ͻ���װ), 0, rsTemp!�ͻ���װ)
        str�ͻ���λ = IIf(IsNull(rsTemp!�ͻ���λ), "", rsTemp!�ͻ���λ)
        .TextMatrix(.Row, mconIntcolҽ������) = IIf(IsNull(rsTemp!��������), "", rsTemp!��������)
        .TextMatrix(.Row, mconIntCol����ҩ��) = IIf(IsNull(rsTemp!����ҩ��), "", rsTemp!����ҩ��)
        If dbl�ͻ���װ <> 0 Then
            .TextMatrix(.Row, mconintCol�ͻ���λ) = str�ͻ���λ & "(1" & str�ͻ���λ & "=" & zlStr.FormatEx(dbl�ͻ���װ / dbl����ϵ��, 1, , True) & str��λ & ")"
            .TextMatrix(.Row, mconintCol�ͻ���װ) = dbl�ͻ���װ
            If Val(.TextMatrix(.Row, mconintCol�ƻ�����)) <> 0 Then
                .TextMatrix(.Row, mconintCol�ͻ�����) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconintCol�ƻ�����)) / dbl�ͻ���װ, 1, , True)
            End If
        End If
        
        If gintҩƷ������ʾ = 0 Or gintҩƷ������ʾ = 2 Then
            strҩ�� = strͨ����
        Else
            strҩ�� = IIf(str��Ʒ�� <> "", str��Ʒ��, strͨ����)
        End If
        
        .TextMatrix(.Row, mconIntColҩƷ���������) = strҩƷ���� & strҩ��
        .TextMatrix(.Row, mconIntColҩƷ����) = strҩƷ����
        .TextMatrix(.Row, mconIntColҩƷ����) = strҩ��
        
        If mintDrugNameShow = 1 Then
            .TextMatrix(.Row, mconIntColҩ��) = .TextMatrix(.Row, mconIntColҩƷ����)
        ElseIf mintDrugNameShow = 2 Then
            .TextMatrix(.Row, mconIntColҩ��) = .TextMatrix(.Row, mconIntColҩƷ����)
        Else
            .TextMatrix(.Row, mconIntColҩ��) = .TextMatrix(.Row, mconIntColҩƷ���������)
        End If
        
        .TextMatrix(.Row, mconIntCol��Ʒ��) = str��Ʒ��
        
        .TextMatrix(.Row, mconIntCol��Դ) = strҩƷ��Դ
        .TextMatrix(.Row, mconIntCol���) = str���
        .TextMatrix(.Row, mconIntCol��λ) = str��λ
        .TextMatrix(.Row, mconintCol�ɱ���) = zlStr.FormatEx(dbl�ɱ����� * dbl����ϵ��, mintShowPriceDigit, , True)
        .TextMatrix(.Row, mconIntCol�ۼ�) = zlStr.FormatEx(dbl���۵��� * dbl����ϵ��, mintShowPriceDigit, , True)
        .TextMatrix(.Row, mconIntCol��׼�ĺ�) = str��׼�ĺ�
    End With
    rsData.Close
    SetDrugRows = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

'ȡָ�������۶��۵�λ������ֵ��ȱʡΪ0-���ۼ۵�λ���ۣ���ѡΪ1-��ҩ�ⵥλ���ۣ�
Private Function GetUnit() As Integer
    GetUnit = gtype_UserSysParms.P29_ָ�������۶��۵�λ
End Function

Private Sub vsfHisPlane_EnterCell()
    Dim lngColor As Long
    
    With vsfHisPlane
        .ForeColorSel = &H80000008
        
        If .Row = 0 Then Exit Sub
        If .TextMatrix(.Row, .ColIndex("NO")) = "" Then Exit Sub
        
        lngColor = .Cell(flexcpForeColor, .Row, .ColIndex("�ƻ�����"))
        
        .ForeColorSel = lngColor
    End With
End Sub

Private Function CheckData(ByVal rsTemp As ADODB.Recordset) As ADODB.Recordset
    '���ܣ���������б�������ҩƷ����ѡ���ҩƷ�Ƿ��ظ���ʱ��ҩƷ�Ƿ��п��

    Dim i As Integer
    Dim strTemp As String
    Dim strInfo As String
    Dim strSql As String
    Dim strDub As String    '�ظ�ҩƷ
    Dim str�ظ�ҩ�� As String   '������¼�ظ�ѡ���˵�ҩƷ����
    
    rsTemp.MoveFirst
    strTemp = ""
    Do While Not rsTemp.EOF
        If InStr(1, strTemp, rsTemp!ҩƷid) = 0 Then
            strTemp = strTemp & rsTemp!ҩƷid & "," & rsTemp!ͨ���� & "|"
        End If
        rsTemp.MoveNext
    Loop
        
    With mshBill    '���ظ��Ĳ�ѯ����
        For i = 1 To .rows - 2
            If InStr(1, strTemp, .TextMatrix(i, 0) & ",") > 0 Then
                strInfo = strInfo & .TextMatrix(i, 0) & "," & .TextMatrix(i, mconIntColҩ��) & "|"
            End If
        Next
        
        If strInfo <> "" Then   'Ϊ��������ƴ��sql
            strDub = ""
            For i = 0 To UBound(Split(strInfo, "|")) - 1
                strDub = strDub & "ҩƷid<>" & Split(Split(strInfo, "|")(i), ",")(0) & " and "
                If UBound(Split(str�ظ�ҩ��, ",")) <= 2 Then
                    str�ظ�ҩ�� = str�ظ�ҩ�� & Split(Split(strInfo, "|")(i), ",")(1) & ","
                End If
            Next
            If strDub <> "" Then
                strDub = Mid(strDub, 1, Len(strDub) - 4)
            End If
        End If
        
        '�ж���ʲô��ʽƴ��sql
        If str�ظ�ҩ�� <> "" Then
            MsgBox str�ظ�ҩ�� & "�б����Ѿ������ˣ�" & vbCrLf & "����ҩƷ������ӣ�", vbInformation, gstrSysName
            strSql = strDub
        End If
        If strSql <> "" Then
            rsTemp.Filter = strSql
        End If
        
        Set CheckData = rsTemp
    End With
End Function

Private Sub SetColumnByUserDefine()
    Dim intColumns As Integer
    Dim strColumn_Selected As String
    Dim strColumn_All As String
    Dim arrColumn_All, arrColumn_Selected, arrColumn_UnSelected, arr����, arr��������
    Dim intCol As Integer, intCols As Integer
    Dim strAllCol As String
    
    On Error GoTo ErrHand
    
    strColumn_Selected = zlDataBase.GetPara("ѡ����", glngSys, ģ���.ҩƷ�ƻ�)
    mstrColumn_UnSelected = zlDataBase.GetPara("������", glngSys, ģ���.ҩƷ�ƻ�)
    strColumn_All = "ҩ��,2|��Ʒ��,3|ҩƷ��Դ,4|���,5|������,6|ԭ����,7|��λ,8|ҽ������,10|ǰ������,11|��������,12|�������,13|�������,14|" & _
                        "�������,15|��������,16|��������,17|�ƻ�����,18|ִ������,19|�ͻ���λ,21|�ͻ�����,22|�ɱ���,24|�ɱ����,25|�ۼ�,26|�ۼ۽��,27|�ϴι�Ӧ��,28|˵��,29|����ҩ��,33|��׼�ĺ�,34"
    
    If strColumn_Selected <> "" Then
        If mstrColumn_UnSelected <> "" Then
            strAllCol = strColumn_Selected & "|" & mstrColumn_UnSelected
        Else
            strAllCol = strColumn_Selected
        End If
        arr���� = Split(strColumn_All, "|")
        arr�������� = Split(strAllCol, "|")
        If UBound(arr����) <> UBound(arr��������) Then
            strColumn_Selected = "ҩ��|��Ʒ��|ҩƷ��Դ|���|������|ԭ����|��λ|ҽ������|ǰ������|��������|�������|�������|�������|��������|��������|�ƻ�����|ִ������|�ͻ���λ|�ͻ�����|�ɱ���|�ɱ����|�ۼ�|�ۼ۽��|�ϴι�Ӧ��|˵��|����ҩ��|��׼�ĺ�"
            mstrColumn_UnSelected = ""
            zlDataBase.SetPara "ѡ����", strColumn_Selected, glngSys, ģ���.ҩƷ�ƻ�
            zlDataBase.SetPara "������", mstrColumn_UnSelected, glngSys, ģ���.ҩƷ�ƻ�
        End If
    Else
        strColumn_Selected = "ҩ��|��Ʒ��|ҩƷ��Դ|���|������|ԭ����|��λ|ҽ������|ǰ������|��������|�������|�������|�������|��������|��������|�ƻ�����|ִ������|�ͻ���λ|�ͻ�����|�ɱ���|�ɱ����|�ۼ�|�ۼ۽��|�ϴι�Ӧ��|˵��|����ҩ��|��׼�ĺ�"
        mstrColumn_UnSelected = ""
        zlDataBase.SetPara "ѡ����", strColumn_Selected, glngSys, ģ���.ҩƷ�ƻ�
        zlDataBase.SetPara "������", mstrColumn_UnSelected, glngSys, ģ���.ҩƷ�ƻ�
    End If
    
    '����Ĭ��ֵ
    mconIntCol��� = 1
    mconIntColҩ�� = 2
    mconIntCol��Ʒ�� = 3
    mconIntCol��Դ = 4
    mconIntCol��� = 5
    mconIntCol������ = 6
    mconIntColԭ���� = 7
    mconIntCol��λ = 8
    mconIntCol����ϵ�� = 9
    mconIntcolҽ������ = 10
    mconIntColǰ������ = 11
    mconIntCol�������� = 12
    mconIntCol������� = 13
    mconIntCol������� = 14
    mconintCol������� = 15
    mconintCol�������� = 16
    mconintCol�������� = 17
    mconintCol�ƻ����� = 18
    mconintColִ������ = 19
    mconintColԭִ������ = 20
    mconintCol�ͻ���λ = 21
    mconintCol�ͻ����� = 22
    mconintCol�ͻ���װ = 23
    mconintCol�ɱ��� = 24
    mconIntCol�ɱ���� = 25
    mconIntCol�ۼ� = 26
    mconIntCol�ۼ۽�� = 27
    mconintCol�ϴι�Ӧ�� = 28
    mconintCol˵�� = 29
    mconIntColҩƷ��������� = 30
    mconIntColҩƷ���� = 31
    mconIntColҩƷ���� = 32
    mconIntCol����ҩ�� = 33
    mconIntCol��׼�ĺ� = 34
    mconIntColS = 35      '������
    mshBill.Cols = 35
    '��Ʒ���д���
    If gintҩƷ������ʾ = 2 Then
        '��ʾ��Ʒ����
        mshBill.ColWidth(mconIntCol��Ʒ��) = IIf(mshBill.ColWidth(mconIntCol��Ʒ��) = 0, 2000, mshBill.ColWidth(mconIntCol��Ʒ��))
    Else
        '��������ʾ��Ʒ����
        mshBill.ColWidth(mconIntCol��Ʒ��) = 0
    End If

    '�����û����õ�����˳��
    arrColumn_All = Split(strColumn_All, "|")
    arrColumn_Selected = Split(strColumn_Selected, "|")
    intCols = UBound(arrColumn_Selected)
    For intCol = 0 To intCols
        Call SetColumnValue(arrColumn_Selected(intCol), Split(arrColumn_All(intCol), ",")(1))
    Next
    
    '��δѡ����е��п�����Ϊ�㣬��������Ϊ5��������ѡ��
    If mstrColumn_UnSelected = "" Then Exit Sub
    intCol = intCols + 1
    intColumns = 0
    arrColumn_UnSelected = Split(mstrColumn_UnSelected, "|")
    intCols = UBound(arrColumn_All)
    For intCol = intCol To intCols
        If UBound(arrColumn_UnSelected) >= intColumns Then
            Call SetColumnValue(arrColumn_UnSelected(intColumns), Split(arrColumn_All(intCol), ",")(1), False)
            intColumns = intColumns + 1
        Else
            Call SetColumnValue(Split(arrColumn_All(intCol), ",")(0), Split(arrColumn_All(intCol), ",")(1), False)
        End If
    Next
    Exit Sub
ErrHand:
    MsgBox "�ָ�������ʱ�������������½��������ã�", vbInformation, gstrSysName
End Sub

Private Sub SetColumnValue(ByVal str���� As String, ByVal intValue As Integer, Optional ByVal blnShow As Boolean = True)
    Select Case str����
        Case "���"
            mconIntCol��� = intValue
        Case "ҩ��"
            mconIntColҩ�� = intValue
        Case "��Ʒ��"
            mconIntCol��Ʒ�� = intValue
        Case "ҩƷ��Դ"
            mconIntCol��Դ = intValue
        Case "���"
            mconIntCol��� = intValue
        Case "������"
            mconIntCol������ = intValue
        Case "ԭ����"
            mconIntColԭ���� = intValue
        Case "��λ"
            mconIntCol��λ = intValue
        Case "ҽ������"
            mconIntcolҽ������ = intValue
        Case "ǰ������"
            mconIntColǰ������ = intValue
        Case "��������"
            mconIntCol�������� = intValue
        Case "�������"
            mconIntCol������� = intValue
        Case "�������"
            mconIntCol������� = intValue
        Case "�������"
            mconintCol������� = intValue
        Case "��������"
            mconintCol�������� = intValue
        Case "��������"
            mconintCol�������� = intValue
        Case "�ƻ�����"
            mconintCol�ƻ����� = intValue
        Case "ִ������"
            mconintColִ������ = intValue
        Case "�ͻ���λ"
            mconintCol�ͻ���λ = intValue
        Case "�ͻ�����"
            mconintCol�ͻ����� = intValue
        Case "�ɱ���"
            mconintCol�ɱ��� = intValue
        Case "�ɱ����"
            mconIntCol�ɱ���� = intValue
        Case "�ۼ�"
            mconIntCol�ۼ� = intValue
        Case "�ۼ۽��"
            mconIntCol�ۼ۽�� = intValue
        Case "�ϴι�Ӧ��"
            mconintCol�ϴι�Ӧ�� = intValue
        Case "˵��"
            mconintCol˵�� = intValue
        Case "����ҩ��"
            mconIntCol����ҩ�� = intValue
        Case "��׼�ĺ�"
            mconIntCol��׼�ĺ� = intValue
    End Select
    
    If Not blnShow Then
        mshBill.ColWidth(intValue) = 0
        mshBill.ColData(intValue) = 5
    Else
        mintLastCol = intValue
    End If
End Sub

Private Sub GetDefineSize()
    '���ܣ��õ����ݿ�ı��ֶεĳ���
    On Error GoTo errHandle
    Dim rsTmp As New ADODB.Recordset
     
    gstrSQL = "Select t.�ϴβ��� as ������, t.ԭ���� as ԭ���� From ҩƷ��� T Where Rownum < 1"
    Call zlDataBase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
    
    mlng�����̳��� = rsTmp.Fields("������").DefinedSize
    mlngԭ���س��� = rsTmp.Fields("ԭ����").DefinedSize
   
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Init�洢�ⷿ()
    Dim rsDepend As New Recordset
    
    On Error GoTo errHandle
    
    gstrSQL = "SELECT DISTINCT a.id, a.���� " _
            & "FROM ��������˵�� c, �������ʷ��� b, ���ű� a " _
            & "Where c.�������� = b.���� " _
            & "  AND Instr('HIJKLMN',b.����,1) > 0 " _
            & "  AND a.id = c.����id " _
            & "  AND TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01' " _
            & IIf(zlStr.IsHavePrivs(mstrPrivs, "���пⷿ"), "", " And a.ID IN (Select ����ID From ������Ա Where ��ԱID=[2])")
    Set rsDepend = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, gstrNodeNo, UserInfo.�û�ID)
    
    If rsDepend.EOF Then
        MsgBox "û������ҩ�����ʵĲ���,��鿴���Ź���", vbInformation, gstrSysName
        rsDepend.Close
        Exit Sub
    End If
    
    lvw�洢�ⷿ.ListItems.Clear

    With rsDepend
        Do While Not .EOF
            lvw�洢�ⷿ.ListItems.Add , "K" & !Id, !����, , 2
            .MoveNext
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub Change�洢�ⷿ()
    Dim i As Integer, j As Integer
    Dim intSelect As Integer
    Dim strArr�洢�ⷿ() As String
    strArr�洢�ⷿ = Split(mstr�Զ���ⷿ, ",")
    
    For i = LBound(strArr�洢�ⷿ) To UBound(strArr�洢�ⷿ)
        For intSelect = 1 To lvw�洢�ⷿ.ListItems.count
            If strArr�洢�ⷿ(i) = Mid(lvw�洢�ⷿ.ListItems(intSelect).Key, 2) Then
                lvw�洢�ⷿ.ListItems(intSelect).Checked = True
                j = j + 1
            End If
        Next
    Next
    
    If j = lvw�洢�ⷿ.ListItems.count Then
        chk�ⷿ.Value = 1
    ElseIf j > 0 And j < lvw�洢�ⷿ.ListItems.count Then
        chk�ⷿ.Value = 2
    End If
End Sub
Private Sub cmd�ⷿ_Click()
    With pic�ⷿ
        .Visible = Not .Visible
    End With
    
    Call ResizeHisPlane
    Call Change�洢�ⷿ
End Sub

Private Sub chk�ⷿ_Click()
'�ⷿȫѡ��ť
    If chk�ⷿ.Value = 2 Then Exit Sub
    Call SetSelect(lvw�洢�ⷿ, chk�ⷿ.Value)
End Sub
Private Sub SetSelect(ByVal lvwObj As Object, Optional ByVal BlnSelect As Boolean = True)
'ȫѡ����
    Dim intSelect As Integer
    With lvwObj
        For intSelect = 1 To .ListItems.count
            .ListItems(intSelect).Checked = BlnSelect
        Next
    End With
End Sub
Private Sub lvw�洢�ⷿ_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'����ѡ��Ĵ洢�ⷿ
    Call ItemCheck(lvw�洢�ⷿ, Item, chk�ⷿ)
End Sub
Private Sub ItemCheck(ByVal lvwObj As Object, ByVal Item As MSComctlLib.ListItem, ByVal chkObj As CheckBox)
'��¼ѡ��Ŀⷿ
    Dim lngCheck As Long, intCount As Integer
    
    intCount = 0
    With lvwObj
        For lngCheck = 1 To .ListItems.count
            If .ListItems(lngCheck).Checked = True Then
                intCount = intCount + 1
            End If
        Next
        
        If intCount = lvwObj.ListItems.count Then
            chkObj.Value = 1
        ElseIf intCount > 0 Then
            chkObj.Value = 2
        Else
            chkObj.Value = 0
        End If
    End With
End Sub

