VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Begin VB.UserControl usrBodyEditor 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   9555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11415
   LockControls    =   -1  'True
   ScaleHeight     =   9555
   ScaleWidth      =   11415
   Begin VB.PictureBox picTmp 
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   4020
      ScaleHeight     =   300
      ScaleWidth      =   4590
      TabIndex        =   34
      Top             =   8850
      Width           =   4590
      Begin VB.ComboBox cboBaby 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2625
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   0
         Width           =   1920
      End
      Begin VB.OptionButton opt 
         Caption         =   "ĸ�ױ���(&0)"
         Height          =   210
         Index           =   0
         Left            =   45
         TabIndex        =   36
         Top             =   60
         Value           =   -1  'True
         Width           =   1290
      End
      Begin VB.OptionButton opt 
         Caption         =   "Ӥ��(&1)"
         Height          =   210
         Index           =   1
         Left            =   1680
         TabIndex        =   35
         Top             =   60
         Width           =   1035
      End
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   8280
      Left            =   135
      ScaleHeight     =   8280
      ScaleWidth      =   11220
      TabIndex        =   0
      Top             =   450
      Width           =   11220
      Begin MSComCtl2.FlatScrollBar hsb 
         Height          =   255
         Left            =   4815
         TabIndex        =   1
         Top             =   7200
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Arrows          =   65536
         Max             =   100
         Orientation     =   1179649
      End
      Begin MSComCtl2.FlatScrollBar vsb 
         Height          =   1155
         Left            =   7575
         TabIndex        =   2
         Top             =   6300
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   2037
         _Version        =   393216
         Appearance      =   0
         Max             =   100
         Orientation     =   1179648
      End
      Begin VB.PictureBox picCover 
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   6555
         ScaleHeight     =   660
         ScaleWidth      =   975
         TabIndex        =   3
         Top             =   7185
         Width           =   975
      End
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6810
         Left            =   240
         ScaleHeight     =   6810
         ScaleWidth      =   10920
         TabIndex        =   4
         Top             =   225
         Width           =   10920
         Begin VB.PictureBox picCard 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   690
            Index           =   0
            Left            =   90
            ScaleHeight     =   690
            ScaleWidth      =   10560
            TabIndex        =   12
            Top             =   75
            Width           =   10560
            Begin VB.TextBox txtCard 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H00FF0000&
               Height          =   210
               Index           =   7
               Left            =   4875
               Locked          =   -1  'True
               TabIndex        =   39
               TabStop         =   0   'False
               Text            =   "���"
               Top             =   375
               Width           =   2370
            End
            Begin VB.TextBox txtCard 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H00FF0000&
               Height          =   210
               Index           =   6
               Left            =   3375
               Locked          =   -1  'True
               TabIndex        =   33
               TabStop         =   0   'False
               Text            =   "����"
               Top             =   60
               Width           =   645
            End
            Begin VB.TextBox txtCard 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H00FF0000&
               Height          =   210
               Index           =   5
               Left            =   2445
               Locked          =   -1  'True
               TabIndex        =   31
               TabStop         =   0   'False
               Text            =   "�Ա�"
               Top             =   60
               Width           =   420
            End
            Begin VB.TextBox txtCard 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H00FF0000&
               Height          =   210
               Index           =   4
               Left            =   3375
               Locked          =   -1  'True
               TabIndex        =   17
               TabStop         =   0   'False
               Text            =   "12"
               Top             =   390
               Width           =   615
            End
            Begin VB.TextBox txtCard 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H00FF0000&
               Height          =   210
               Index           =   3
               Left            =   4875
               Locked          =   -1  'True
               TabIndex        =   16
               TabStop         =   0   'False
               Text            =   "��Ժ����"
               Top             =   60
               Width           =   1140
            End
            Begin VB.TextBox txtCard 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H00FF0000&
               Height          =   210
               Index           =   2
               Left            =   465
               Locked          =   -1  'True
               TabIndex        =   15
               TabStop         =   0   'False
               Text            =   "����"
               Top             =   375
               Width           =   2400
            End
            Begin VB.TextBox txtCard 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H00FF0000&
               Height          =   210
               Index           =   1
               Left            =   6645
               Locked          =   -1  'True
               TabIndex        =   14
               TabStop         =   0   'False
               Text            =   "1234567"
               Top             =   60
               Width           =   3825
            End
            Begin VB.TextBox txtCard 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H00FF0000&
               Height          =   210
               Index           =   0
               Left            =   465
               Locked          =   -1  'True
               TabIndex        =   13
               TabStop         =   0   'False
               Text            =   "������"
               Top             =   60
               Width           =   1425
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��    ��:"
               Height          =   180
               Index           =   7
               Left            =   4065
               TabIndex        =   38
               Top             =   390
               Width           =   810
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����:"
               Height          =   180
               Index           =   6
               Left            =   2910
               TabIndex        =   32
               Top             =   60
               Width           =   450
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�Ա�:"
               Height          =   180
               Index           =   4
               Left            =   1980
               TabIndex        =   30
               Top             =   60
               Width           =   450
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��Ժ����:"
               Height          =   180
               Index           =   5
               Left            =   4050
               TabIndex        =   22
               Top             =   60
               Width           =   810
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����:"
               Height          =   180
               Index           =   3
               Left            =   2910
               TabIndex        =   21
               Top             =   390
               Width           =   450
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����:"
               Height          =   180
               Index           =   2
               Left            =   0
               TabIndex        =   20
               Top             =   375
               Width           =   450
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "סԺ��:"
               Height          =   180
               Index           =   1
               Left            =   6000
               TabIndex        =   19
               Top             =   60
               Width           =   630
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����:"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   18
               Top             =   60
               Width           =   450
            End
         End
         Begin VB.PictureBox picScale 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   840
            Left            =   1230
            ScaleHeight     =   810
            ScaleWidth      =   5850
            TabIndex        =   26
            Top             =   1350
            Width           =   5880
            Begin VB.Label lblCur 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��"
               ForeColor       =   &H8000000C&
               Height          =   180
               Left            =   180
               TabIndex        =   27
               Top             =   570
               Width           =   180
            End
         End
         Begin VB.PictureBox picBack 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1950
            Left            =   1155
            ScaleHeight     =   1920
            ScaleWidth      =   5520
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   2355
            Width           =   5550
            Begin VB.PictureBox picGraph 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   4140
               Left            =   780
               ScaleHeight     =   4140
               ScaleWidth      =   5445
               TabIndex        =   25
               TabStop         =   0   'False
               Top             =   135
               Width           =   5445
               Begin VB.Line linHCur 
                  BorderStyle     =   3  'Dot
                  Visible         =   0   'False
                  X1              =   300
                  X2              =   1635
                  Y1              =   720
                  Y2              =   720
               End
               Begin VB.Line linVCur 
                  BorderStyle     =   3  'Dot
                  Visible         =   0   'False
                  X1              =   1785
                  X2              =   1785
                  Y1              =   15
                  Y2              =   690
               End
            End
         End
         Begin VB.PictureBox picLine 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   3420
            Index           =   0
            Left            =   7785
            ScaleHeight     =   3390
            ScaleWidth      =   0
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   1035
            Width           =   15
         End
         Begin zl9BodyEditorHN.VsfGrid vsf 
            Height          =   270
            Left            =   450
            TabIndex        =   5
            Top             =   4440
            Width           =   7275
            _ExtentX        =   12832
            _ExtentY        =   476
         End
         Begin VSFlex8Ctl.VSFlexGrid mshUpTab 
            Height          =   780
            Left            =   240
            TabIndex        =   6
            Top             =   810
            Width           =   5775
            _cx             =   10186
            _cy             =   1376
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
            BackColorFixed  =   -2147483643
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483632
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   3
            HighLight       =   0
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   3
            Cols            =   8
            FixedRows       =   2
            FixedCols       =   1
            RowHeightMin    =   255
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   ""
            ScrollTrack     =   0   'False
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
         Begin VSFlex8Ctl.VSFlexGrid mshDownTab 
            Height          =   1695
            Left            =   300
            TabIndex        =   7
            Top             =   5085
            Width           =   7215
            _cx             =   12726
            _cy             =   2990
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   0   'False
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
            BackColorFixed  =   -2147483643
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483632
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   3
            HighLight       =   0
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   18
            FixedRows       =   1
            FixedCols       =   4
            RowHeightMin    =   255
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   ""
            ScrollTrack     =   0   'False
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
            Begin VB.PictureBox picInput 
               Appearance      =   0  'Flat
               BackColor       =   &H80000001&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   240
               Left            =   1860
               ScaleHeight     =   240
               ScaleWidth      =   3360
               TabIndex        =   8
               Top             =   495
               Visible         =   0   'False
               Width           =   3360
               Begin VB.TextBox txtInput 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000018&
                  BorderStyle     =   0  'None
                  Height          =   240
                  Index           =   0
                  Left            =   0
                  MaxLength       =   12
                  TabIndex        =   10
                  Top             =   0
                  Width           =   1035
               End
               Begin VB.TextBox txtInput 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000018&
                  BorderStyle     =   0  'None
                  Height          =   240
                  Index           =   1
                  Left            =   1335
                  MaxLength       =   12
                  TabIndex        =   9
                  Top             =   0
                  Width           =   1035
               End
               Begin VB.Label lblInput 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "/E"
                  Height          =   180
                  Left            =   1065
                  TabIndex        =   11
                  Top             =   30
                  Width           =   180
               End
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid mshScale 
            Height          =   3090
            Left            =   150
            TabIndex        =   28
            Top             =   1185
            Width           =   6420
            _cx             =   11324
            _cy             =   5450
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
            BackColorFixed  =   -2147483643
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483632
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   3
            HighLight       =   0
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   0
            GridLinesFixed  =   0
            GridLineWidth   =   1
            Rows            =   12
            Cols            =   8
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   ""
            ScrollTrack     =   0   'False
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
            ExplorerBar     =   0
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
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
         End
         Begin VB.Label lblComment 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "˵����"
            Height          =   180
            Left            =   315
            TabIndex        =   29
            Top             =   6570
            Width           =   540
         End
      End
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "usrBodyEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################
'����
'----------------------------------------------------------------------------------------------------------------------

'�洢���µ����߲��ݵ����ݵı�־
Private Enum GraphDataRow
    ���ı�־ = 0
    �������� = 1
    �ϱ�˵�� = 2
    ������־ = 3
    ��λ��־ = 4
    ��Ժ��־ = 5
    ת�Ʊ�־ = 6
    ������־ = 7
    ��Ժ��־ = 8
    ��Ʊ�־ = 9
    ���Ա�־ = 10
    �±�˵�� = 11
    �Ͽ���־ = 12
    ������־ = 13
    ����ʱ�� = 14
    δ��˵�� = 15
End Enum

Private Enum GridDataRow
    �޸ı�־ = 0
End Enum

'��������
Private Enum OperateType
    �������� = 2                                        'ȫ���������ĵ�
    �޸Ĳ��� = 3                                        '�޸ĵĵ㣺���ܰ���ԭ�еĵ�������ĵ�
    ɾ������ = 4                                        'ɾ���ĵ�
End Enum

Private Const HOUR_STEP_Twips = 240                 '��С��Ԫ��Ŀ��
Private Const ROWHEIGHT = 39                        '��С��Ԫ��ĸ߶�*5
Private Const MAXROWS = 47                          '

'�Զ�������
'----------------------------------------------------------------------------------------------------------------------
Private Type ITEM_NO
    ��� As Long
    ��Һ As Long
    ���� As Long
    ���� As Long
    ���� As Long
    ���� As Long
    Ѫѹ As Long
    ����ѹ As Long
End Type

Private Type ITEM_SERIAL
    ������ As Integer
    ������ As Integer
    ���� As Integer
    ���� As Integer
    Ѫѹ As Integer
    ���� As Integer
    ���� As Integer
End Type

Private Type ITEM_STRUCT
    ��Ŀ���� As String
    �������� As Integer
    ���ݳ��� As Integer
    С��λ�� As Integer
    ��Сֵ As String
    ���ֵ As String
    ��¼Ƶ�� As Integer
    ���Ŀ As Boolean
    ��Ŀ��� As Long
End Type

Private Type GRAPHPOINT
    X As Single
    Y As Single
    ���� As String
    ��ɫ As Long
    ��־ As Byte
End Type

Private Type BODYFLAG
    ��Ժ As Byte
    ��� As Byte
    ת�� As Byte
    ���� As Byte
    ���� As Byte
    ��Ժ As Byte
    ���� As Byte
    ���� As Byte
End Type

'��������
'----------------------------------------------------------------------------------------------------------------------
Private mintOpDays As Integer
Private mblnStopFlag As Boolean
Private mbln�������� As Boolean
Private mblnBabys As Boolean
Private mint����Ӧ�� As Integer
Private mstr���ʷ��� As String
Private mstr��Сʱ�� As String      '��Ժʱ������ʱ��
Private mblnӤ�����µ���ʾ��Ժ As Boolean
Private mstrParam As String
Private mlngHourBegin As Long                       '����̶ȿ�ʼʱ��
Private mlngPageCur As Long                         '��¼��ǰ����һҳ
Private mblnMoved As Boolean
Private mstrEnterDate As String                      '��Ժ����
Private mstrSQL As String
Private rsTemp As New ADODB.Recordset
Private intRow As Integer
Private intCol As Integer
Private intCount As Integer                         '�������ɼ�����
Private mvarEdit As Boolean                         '�Ƿ�����༭
Private mblnNoneShow As Boolean                     'ȷ����ǰ���±��ǲ�����ʾ����
Private mrsParam As New ADODB.Recordset             '�洢��ʽ������ID,��ҳID,����ID,����ID,��Ժ,�༭
Private mstrMsgTitle As String
Private mfrmParent As Object
Private mlngLine As Long
Private mstr���²�λ As String
Private mstr������ʽ As String
Private mstr���� As String                          '�������
Private mlngNo As Long
Private mstrOpsSvr(1 To 7) As String
Private mstrOpsDays(1 To 7) As String
Private mItemSerial As ITEM_SERIAL
Private mItemNo As ITEM_NO
Private mBodyFlag As BODYFLAG
Private mItemStru() As ITEM_STRUCT
Private mItemOtherStru(0 To 1) As ITEM_STRUCT       '������0->����;1-����ѹ
Private mstrChar(2) As String                       '����Ϊ����,Ҹ��,����
Private mstrBreath As String                        '����
Private mstrPulse As String                         '����
Private mcbrToolBarҳ�� As CommandBarControl
Private mcbrToolBar As CommandBar
Private WithEvents mfrmCaseTendBodyPrint As frmCaseTendBodyPrint
Attribute mfrmCaseTendBodyPrint.VB_VarHelpID = -1

'�¼�����
Public Event Activate()
Public Event DbClickCur()
Public Event DataChanged(ByVal blnChanged As Boolean)
Public Event RButton(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event PromptInfo(ByVal strInfo As String)
Public Event SelectScale(ByVal intScale As Integer)
Public Event zlAfterPrint()

Private msinVStep As Single      '�������Ĳ���
Private msinHStep As Single      '�������Ĳ���

'API����
'----------------------------------------------------------------------------------------------------------------------
Private Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal wOptions As Long, lpRect As RECT, ByVal lpString As String, ByVal nCount As Long, lpDx As Long) As Long

'�޸�˵��
'---------------------------------------------------------------
'20090923:�������Ӻ����������������ܣ�������Ϊ�����¼��ʱ��Ĭ��Ϊ�������������Ҫ֧�ֺ�������������������Ϊ������


'�Զ��庯������������
'######################################################################################################################
Public Property Get ParentForm() As Object
    Set ParentForm = mfrmParent
End Property

Public Property Set ParentForm(objParent As Object)
    Set mfrmParent = objParent
End Property

Public Property Get ScrollBarY() As FlatScrollBar
    Set ScrollBarY = vsb
End Property

Public Property Get ScrollBarX() As FlatScrollBar
    Set ScrollBarX = hsb
End Property

Public Property Get ������Ŀ() As Boolean
    ������Ŀ = (mItemSerial.���� = Val(picGraph.Tag))
End Property

Public Property Let ���²�λ(vData As String)
    mstr���²�λ = vData
End Property

Public Property Get ������Ŀ() As Boolean
    ������Ŀ = (mItemSerial.���� = Val(picGraph.Tag))
End Property

Public Property Let ������ʽ(vData As String)
    mstr������ʽ = vData
End Property

Public Property Get ������Ŀ() As Boolean
    ������Ŀ = (mItemSerial.���� = Val(picGraph.Tag))
End Property

Public Property Let ������ʽ(vData As String)
    mstr���� = vData
End Property

Public Property Get CurPostion() As Long
    CurPostion = lblCur.Left \ HOUR_STEP_Twips
End Property

Public Property Get �Ƿ�����Ŀ() As Boolean

    �Ƿ�����Ŀ = (Val(mshDownTab.RowData(mshDownTab.Row)) = mItemNo.���)
    
End Property

Public Property Get �Ƿ��Һ��Ŀ() As Boolean

    �Ƿ��Һ��Ŀ = (Val(mshDownTab.RowData(mshDownTab.Row)) = mItemNo.��Һ)
    
End Property

Public Property Get GetPicScale() As Object
    Set GetPicScale = picScale
End Property

Public Property Get GetmshScale() As Object
    Set GetmshScale = mshScale
End Property

Public Property Get GetUpObj() As Object
    Set GetUpObj = mshUpTab
End Property

Public Property Get GetpicLine(ByVal intIndex) As Object
    Set GetpicLine = picLine(intIndex)
End Property
Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get Page() As Long
    Page = mlngPageCur
End Property

Public Property Get LineType() As Long
    LineType = mlngLine
End Property

Public Function ConvertToValue(ByVal intNo As Integer, ByVal Y As Long) As Double
    '******************************************************************************************************************
    '���ܣ� ת��������Ϊֵ
    '������
    '���أ�
    '******************************************************************************************************************
    Dim aryValue() As String
    
    aryValue = Split(picLine(intNo).Tag, ";")
    
    ConvertToValue = aryValue(0) - (Y / mshScale.ROWHEIGHT(1) - aryValue(3) + 1) * aryValue(2)
    
End Function

Public Function ConvertToY(ByVal intCol As Integer, ByVal dbValue As Double) As Long
    '******************************************************************************************************************
    '���ܣ� ת��ֵΪ������
    '������
    '���أ�
    '******************************************************************************************************************
    Dim aryValue() As String
    
    '��ȡ��Ŀ����:���ֵ����Сֵ����λֵ�������
    aryValue = Split(picLine(intCol).Tag, ";")

    '����ֵ=((���ֵ-��ǰֵ)/��λֵ+�����-1)*�и߶�
    ConvertToY = ((Val(aryValue(0)) - dbValue) / Val(aryValue(2)) + Val(aryValue(3)) - 1) * mshScale.ROWHEIGHT(1)
    
End Function

Public Function GetMaxValue(ByVal intCol As Integer) As Double
    '******************************************************************************************************************
    '���ܣ� ��ȡ���ֵ
    '������
    '���أ�
    '******************************************************************************************************************
    Dim aryValue() As String
    
    '��ȡ��Ŀ����:���ֵ����Сֵ����λֵ�������
    aryValue = Split(picLine(intCol).Tag, ";")

    GetMaxValue = Val(aryValue(0))

End Function

Public Function GetMinValue(ByVal intCol As Integer) As Double
    '******************************************************************************************************************
    '���ܣ� ��ȡ���ֵ
    '������
    '���أ�
    '******************************************************************************************************************
    Dim aryValue() As String
    
    '��ȡ��Ŀ����:���ֵ����Сֵ����λֵ�������
    aryValue = Split(picLine(intCol).Tag, ";")

    GetMinValue = Val(aryValue(1))

End Function


Public Function zlMenuClick(ByVal strMenuItem As String, Optional ByVal strParam As String) As Boolean
    '******************************************************************************************************************
    '���ܣ� �˵����ܴ�����Ҫ�����ϼ�����ӿڵ���
    '������ strMenuItem         ��������
    '       strParam            �����ַ���
    '���أ� ���óɹ�����TRUE������FALSE
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim varParam As Variant
    Dim strTmp As String
    Dim intMinCol As Long
    Dim intMaxCol As Long
    Dim aryValue As Variant
    Dim intRewrite As Integer
    Dim intNowCol As Integer
    Dim intCol As Integer
    Dim intLoop As Integer
        
    If strParam <> "" Then varParam = Split(strParam, ";")
    Select Case strMenuItem
    '------------------------------------------------------------------------------------------------------------------
    Case "��ʼ����"
        
        'strParam��ʽ������ID;��ҳID;����ID;����ID;��Ժ;�༭
        
        Set mrsParam = New ADODB.Recordset
    
        Call CreateParam(mrsParam, "����id", adBigInt)
        Call CreateParam(mrsParam, "��ҳid", adBigInt)
        Call CreateParam(mrsParam, "����id", adBigInt)
        Call CreateParam(mrsParam, "����id", adBigInt)
        Call CreateParam(mrsParam, "��Ժ", adTinyInt)
        Call CreateParam(mrsParam, "Ӥ��", adTinyInt)
        Call CreateParam(mrsParam, "�༭", adTinyInt)
        Call CreateParam(mrsParam, "������Դ", adTinyInt)
        Call CreateParam(mrsParam, "��ʼʱ��", adVarChar, 20)
        Call CreateParam(mrsParam, "����ʱ��", adVarChar, 20)
        Call CreateParam(mrsParam, "����ȼ�", adTinyInt)
        
        mrsParam.Open
        mrsParam.AddNew
                        
        mrsParam("����id").Value = Val(varParam(0))
        mrsParam("��ҳid").Value = Val(varParam(1))
        mrsParam("����id").Value = Val(varParam(2))
        mrsParam("����id").Value = Val(varParam(2))
        If UBound(varParam) >= 3 Then
            mrsParam("��Ժ").Value = Val(varParam(3))
        Else
            mrsParam("��Ժ").Value = 1
        End If
        
        If UBound(varParam) >= 4 Then
            mrsParam("�༭").Value = Val(varParam(4))
        Else
            mrsParam("�༭").Value = 0
        End If
        
        If UBound(varParam) >= 5 Then
            mrsParam("Ӥ��").Value = Val(varParam(5))
        Else
            mrsParam("Ӥ��").Value = 0
        End If
        
        mrsParam("������Դ").Value = 2
        mrsParam("����ȼ�").Value = 3
        
        gstrSQL = "Select a.���,Decode(a.Ӥ������,Null,b.����||'֮��'||Trim(To_Char(a.���,'9')),a.Ӥ������) As Ӥ������ From ������������¼ a,������Ϣ b Where a.����id=[1] And a.��ҳid=[2] And a.����id=b.����id Order By a.���"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "usrBodyEditor", Val(mrsParam("����id").Value), Val(mrsParam("��ҳid").Value))
        mblnBabys = (rs.BOF = False)
        picTmp.Visible = mblnBabys
        cboBaby.Clear
        If rs.BOF = False Then
            Do While Not rs.EOF
                cboBaby.AddItem rs("Ӥ������").Value
                cboBaby.ItemData(cboBaby.NewIndex) = rs("���").Value
                rs.MoveNext
            Loop
        End If
        If cboBaby.ListCount > 0 Then cboBaby.ListIndex = 0

        
        Call InitData
                
        Call ClearLineSelect
        Call FaceInit
        Call SetBodyMode
        
        If InitBody(Val(mrsParam("����id").Value), Val(mrsParam("��ҳid").Value), Val(mrsParam("����id").Value), Val(mrsParam("Ӥ��").Value)) = False Then Exit Function
        
    '------------------------------------------------------------------------------------------------------------------
    Case "װ������"
        
        'strParam��ʽ����ʼʱ��;����ID;��ʼʱ��;����ʱ��;ҳ��
        
        '�ж����±��Ƿ񱣴棬ѯ�ʱ������
        strTmp = isSaved()
        If strTmp <> "" Then
            If MsgBox(strTmp & "�޸���Ϣ��ʧ��" & vbCrLf & "�Ƿ�������棿", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        End If
        
        mlngNo = -1
        If strParam = "" Then
            
            '��һ���յ����±�,��û�����ڡ����ݼ�������Ϣ
            Call DrawScale
            Call DrawPaper
        Else
             mstrParam = strParam
             
'             If InitBody(Val(mrsParam("����id").Value), Val(mrsParam("��ҳid").Value), Val(mrsParam("����id").Value), Val(mrsParam("Ӥ��").Value)) = False Then Exit Function
             
             mrsParam("����id").Value = Val(varParam(1))
             mrsParam("����id").Value = Val(varParam(1))
             mrsParam("��ʼʱ��").Value = CStr(varParam(2))
             mrsParam("����ʱ��").Value = CStr(varParam(3))
             mlngNo = Val(varParam(4))
             
            '�����µ�ҳ
            mlngPageCur = Val(varParam(4))
            mstrEnterDate = CStr(varParam(0))
            
            Call ReadBodyData
            Call DrawScale
            Call DrawPaper
            Call DrawGraph
            
            '������ʾ�ؼ�
'            Call UserControl_Resize
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case "ˢ������"
        
        
        '�ж����±��Ƿ񱣴棬ѯ�ʱ������
        strTmp = isSaved()
        If strTmp <> "" Then
            If MsgBox(strTmp & "�޸���Ϣ��ʧ��" & vbCrLf & "�Ƿ�������棿", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        End If
        
        If mstrParam = "" Then
            
            '��һ���յ����±�,��û�����ڡ����ݼ�������Ϣ
            Call DrawScale
            Call DrawPaper
        Else
            
            Call ReadBodyData
            
            Call DrawScale
            Call DrawPaper
            Call DrawGraph
            
            '������ʾ�ؼ�
'            Call UserControl_Resize
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case "��������"
        
        If Val(mrsParam("�༭")) = 0 Then Exit Function

        'strParam��ʽ��
        If SaveData Then
            '��Ҫ����װ��
            zlMenuClick = True
            Call ReadBodyData
            Call DrawScale
            Call DrawPaper
            Call DrawGraph
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case "�ָ�����"
        
        If Val(mrsParam("�༭")) = 0 Then Exit Function
        
        If MsgBox("ȷʵҪ�ָ�����ǰ��������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        
        Call ReadBodyData
        Call DrawScale
        Call DrawPaper
        Call DrawGraph
    '------------------------------------------------------------------------------------------------------------------
    Case "��������"
        
        mlngLine = Val(strParam)
        
        If Val(mrsParam("�༭")) = 0 Then Exit Function
        
        'strParam��ʽ����������
                
        If Val(varParam(0)) = 0 Then
            If picGraph.Tag <> "" Then
                mshScale_MouseUp 1, 0, Val(picGraph.Tag) * mshScale.ColWidth(0) + 90, 0
                Call ClearLineSelect
            End If
        Else
            mshScale_MouseUp 1, 0, Val(varParam(0)) * mshScale.ColWidth(0) - 90, 0
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case "�����Ŀ"
        
        Dim rsData As New ADODB.Recordset
        Dim rsTmp As New ADODB.Recordset
        Dim strNotItem As String
                
        strNotItem = ""
        For intLoop = LBound(mItemStru) To UBound(mItemStru) Step -1
            
            If mItemStru(intLoop).���Ŀ Then
                strNotItem = strNotItem & "," & mItemStru(intLoop).��Ŀ���
            End If
            
        Next
        If strNotItem <> "" Then strNotItem = Mid(strNotItem, 2)
                
        Set rsData = GetGridItem(Val(mrsParam("����ȼ�").Value), Val(mrsParam("����id").Value), IIf(Val(mrsParam("Ӥ��").Value) = 0, 1, 2), 2, strNotItem)
        
        If rsData.BOF = False Then
            If ShowTxtSelDialog(mfrmParent, Nothing, "����,1500,0,1;��λ,900,0,0;��Сֵ,900,0,0;���ֵ,900,0,0", mfrmParent.Name & "\������Ŀѡ��", "�������ѡ��һ��������Ŀ��", rsData, rsTmp, 6000, 3000, , , 2, False) Then
                If rsTmp.BOF = False Then
                    If AppendGridItem(rsTmp, True) Then
                        Call picPane_Resize
                    End If
                End If
            End If
        End If


    '------------------------------------------------------------------------------------------------------------------
    Case "ɾ����Ŀ"
        
        With mshDownTab
            If .Row <= UBound(mItemStru) Then
                If mItemStru(.Row).���Ŀ Then
                    
                    '����Ƿ������ݣ����������ʱ������ɾ��
                    '���α���֮ǰ�������Լ���ǰ�����������ݣ����֮Ϊ������
                    If CheckGridData(.Row) Then
                        ShowSimpleMsg "�Բ�����Ҫɾ������������ݻ�����ǰ�����ݣ�"
                        Exit Function
                    End If
                    
                    If MsgBox("ȷʵҪɾ����ǰ�ı����Ŀ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                    
                    If DeleteActiveItem(.Row) Then
                        Call picPane_Resize
                    End If
                    
                End If
            End If
        End With
        
    '------------------------------------------------------------------------------------------------------------------
    Case "���Ժϸ�"
        
        If Val(mrsParam("�༭")) = 0 Then Exit Function
        If picScale.Tag <> "" Then
            With mshScale
                intNowCol = .FixedCols + lblCur.Left \ HOUR_STEP_Twips
                
                '�ж��Ƿ�����ֵ
                If .TextMatrix(1, intNowCol) <> "" Then
                    If Split(.TextMatrix(1, intNowCol), ";")(mItemSerial.���� + 1) <> "" Then
                        aryValue = Split(Split(.TextMatrix(1, intNowCol), ";")(mItemSerial.���� + 1), ",")
                        
                        If .TextMatrix(10, intNowCol) <> "1" And Val(aryValue(0)) > 0 Then
                            .TextMatrix(10, intNowCol) = "1"
                            RaiseEvent DataChanged(True)
                            
                            aryValue = Split(.TextMatrix(0, intNowCol), ";")
                            
                            intRewrite = Val(aryValue(mItemSerial.���� + 1))
                            Select Case intRewrite
                            Case 0
                                aryValue(mItemSerial.���� + 1) = 2
                            Case 1
                                aryValue(mItemSerial.���� + 1) = 3
                            Case 2
                                aryValue(mItemSerial.���� + 1) = 2
                            Case 3
                                aryValue(mItemSerial.���� + 1) = 3
                            Case 4
                                aryValue(mItemSerial.���� + 1) = 3
                            End Select
                            
                            .TextMatrix(0, intNowCol) = Join(aryValue, ";")
                            
                            Call DrawPaper
                            Call DrawGraph
                                            
                        End If
                    End If
                End If
            End With
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case "ȡ������"
                
        If Val(mrsParam("�༭")) = 0 Then Exit Function
        If picScale.Tag <> "" Then
                
            With mshScale
                intNowCol = .FixedCols + lblCur.Left \ HOUR_STEP_Twips
                
                
                '�ж��Ƿ��������»��µ��������������ֵʱ
                If .TextMatrix(1, intNowCol) <> "" Then
                    
                    If Split(.TextMatrix(1, intNowCol), ";")(mItemSerial.���� + 1) <> "" Then
                    
                        aryValue = Split(Split(.TextMatrix(1, intNowCol), ";")(mItemSerial.���� + 1), ",")
        
                        If .TextMatrix(10, intNowCol) = "1" And Val(aryValue(0)) > 0 Then
                        
                            .TextMatrix(10, intNowCol) = "0"
                            RaiseEvent DataChanged(True)
                            
                            aryValue = Split(.TextMatrix(0, intNowCol), ";")
                            
                            intRewrite = Val(aryValue(mItemSerial.���� + 1))
                            Select Case intRewrite
                            Case 0
                                aryValue(mItemSerial.���� + 1) = 2
                            Case 1
                                aryValue(mItemSerial.���� + 1) = 3
                            Case 2
                                aryValue(mItemSerial.���� + 1) = 2
                            Case 3
                                aryValue(mItemSerial.���� + 1) = 3
                            Case 4
                                aryValue(mItemSerial.���� + 1) = 3
                            End Select
                            
                            .TextMatrix(0, intNowCol) = Join(aryValue, ";")
                            
                            Call DrawPaper
                            Call DrawGraph
                
                        End If
                    End If
                End If
                
            End With
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case "��д������"
        
        If Val(mrsParam("�༭")) = 0 Then Exit Function
        If mshUpTab.FocusRect = flexFocusLight Then Exit Function
        
        Call mshUpTab_KeyDown(13, 0)
    '------------------------------------------------------------------------------------------------------------------
    Case "���������"
        
        If Val(mrsParam("�༭")) = 0 Then Exit Function
        Call mshUpTab_KeyDown(46, 0)
        
    '------------------------------------------------------------------------------------------------------------------
    Case "��ȡ������"
        Dim dtOperate As Date
        Dim intStart As Integer
        Dim intEnd As Integer
        For intCol = 1 To mshUpTab.Cols - 1
            Select Case mshUpTab.ColData(intCol)
            Case 0      '��������
            Case 1      'ԭ�����������գ�����Ϊɾ��������
                mshUpTab.ColData(intCol) = 3
            Case 2      '�������գ��ٴ�����Ϊ��������
                mshUpTab.ColData(intCol) = 0
            Case 3      '��ɾ���ĵ�������
            End Select
        Next
        
        '��ҽ����¼��ȡ��������������
        Set rs = GetDataFromHis(Val(mrsParam("����id")), Val(mrsParam("��ҳid")), Val(mrsParam("Ӥ��")), CDate(Split(picScale.Tag, ";")(0)), CDate(Split(picScale.Tag, ";")(1)), 1)
        If Not (rs Is Nothing) Then
            If rs.BOF = False Then
                Do While Not rs.EOF
    
                    dtOperate = Int(rs("ִ��ʱ��").Value)
                    intCol = dtOperate - Int(CDate(CDate(Split(picScale.Tag, ";")(0)))) + 1
                    If intCol >= 1 And intCol <= 7 Then
                    
                        mshUpTab.ColData(intCol) = 1
                        
                        mstrOpsDays(intCol) = Format(rs("ִ��ʱ��").Value, "yyyy-MM-dd HH:mm:ss")
                        
                        '�������ǰ�����ڵ���������������ʾ����
    
                        intStart = GetCurveColumn(CDate(Format(mstrOpsDays(intCol), "yyyy-MM-dd") & " 01:00:00"), CDate(Split(picScale.Tag, ";")(0)), mlngHourBegin) + mshScale.FixedCols - 1
                        intEnd = GetCurveColumn(CDate(Format(mstrOpsDays(intCol), "yyyy-MM-dd") & " 23:00:00"), CDate(Split(picScale.Tag, ";")(0)), mlngHourBegin) + mshScale.FixedCols - 1
                        
                        For intLoop = intStart To intEnd
                            mshScale.TextMatrix(3, intLoop) = ""
                        Next
                        
                        intLoop = GetCurveColumn(CDate(mstrOpsDays(intCol)), CDate(Split(picScale.Tag, ";")(0)), mlngHourBegin) + mshScale.FixedCols - 1
                        
                        Select Case rs("����").Value
                        Case "����"
                            If intLoop >= mshScale.FixedCols And intLoop < mshScale.Cols And mBodyFlag.���� > 0 Then
                                If mBodyFlag.���� = 2 Then
                                    mshScale.TextMatrix(3, intLoop) = "����--" & ConvertTimeToChinese(Format(mstrOpsDays(intCol), "HH:mm"))
                                Else
                                    mshScale.TextMatrix(3, intLoop) = "����"
                                End If
                            End If
                        Case "����"
                            If intLoop >= mshScale.FixedCols And intLoop < mshScale.Cols And mBodyFlag.���� > 0 Then
                                If mBodyFlag.���� = 2 Then
                                    mshScale.TextMatrix(3, intLoop) = "����--" & ConvertTimeToChinese(Format(mstrOpsDays(intCol), "HH:mm"))
                                Else
                                    mshScale.TextMatrix(3, intLoop) = "����"
                                End If
                            End If
                        End Select
                        
                        mshUpTab.Tag = "��д������"
                    End If
                    
                    rs.MoveNext
                Loop
            End If
        End If
        Call ShowOpsDays
        Call DrawPaper
        Call DrawGraph
    '------------------------------------------------------------------------------------------------------------------
    Case "��д��¼��"
        
        If Val(mrsParam("�༭")) = 0 Then Exit Function
                
        If picScale.Tag <> "" Then
            Call CalcMinMaxCol(picScale.Tag, intMinCol, intMaxCol)
                        
            If frmCaseTendBodySetLine.ShowEdit(UserControl.Extender, lblCur.Left \ HOUR_STEP_Twips, intMinCol, intMaxCol, mrsParam("����ȼ�").Value, mint����Ӧ��) Then
                RaiseEvent DataChanged(True)
            End If
            
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case "�����¼��"
        
        If Val(mrsParam("�༭")) = 0 Then Exit Function
        
        If picScale.Tag <> "" Then
                        
            If frmCaseTendBodyDelLine.ShowEdit(UserControl.Extender, lblCur.Left \ HOUR_STEP_Twips, Val(mrsParam("����ȼ�").Value), Val(mrsParam("Ӥ��").Value)) Then
                RaiseEvent DataChanged(True)
            End If
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case "��д�����"
        
        If Val(mrsParam("�༭")) = 0 Then Exit Function
        If mshDownTab.Tag = "" Or mvarEdit = False Then Exit Function
        If CheckTimeRange(mshDownTab.Col) = False Then Exit Function
        
        Call mshDownTab_DblClick
    '------------------------------------------------------------------------------------------------------------------
    Case "��������"
        
        If Val(mrsParam("�༭")) = 0 Then Exit Function
        If mshDownTab.Tag = "" Or mvarEdit = False Then Exit Function
        If CheckTimeRange(mshDownTab.Col) = False Then Exit Function

        mshDownTab.SetFocus
        Call mshDownTab_KeyUp(46, 0)
    '------------------------------------------------------------------------------------------------------------------
    Case "��������"
        
        If Val(mrsParam("�༭")) = 0 Then Exit Function
        If mshDownTab.Tag = "" Or mvarEdit = False Then Exit Function
        If CheckTimeRange(mshDownTab.Col) = False Then Exit Function
                
        If picScale.Tag <> "" And lblCur.Left >= 0 Then
                
            aryValue = Split(picScale.Tag, ";")
            strTmp = Int(CDate(aryValue(0))) + ((lblCur.Left \ HOUR_STEP_Twips) * 4) / 24
            strTmp = Format(strTmp, "yyyy-MM-DD")
            
            If MsgBox("ȷʵ��Ҫ���㡰" & strTmp & "���ڵ����������������", vbQuestion + vbDefaultButton2 + vbYesNo, gstrSysName) = vbNo Then Exit Function
            
            zlMenuClick = ReadDrink(strTmp)
            
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case "�ٸ�"
    
        If Val(mrsParam("�༭")) = 0 Then Exit Function
        If mshDownTab.Tag = "" Or mvarEdit = False Then Exit Function
        If CheckTimeRange(mshDownTab.Col) = False Then Exit Function
        
        If Val(mshDownTab.RowData(mshDownTab.Row)) <> mItemNo.��� Then Exit Function
        If picInput.Visible Then picInput.Visible = False
        
'        mbytSpecChar = 1
'        mshDownTab.Cell(flexcpData, mshDownTab.Row, mshDownTab.Col) = 1
        Call WriteDownTab(mshDownTab.Row, mshDownTab.Col, "*")

    '------------------------------------------------------------------------------------------------------------------
    Case "�೦"
    
        If Val(mrsParam("�༭")) = 0 Then Exit Function
        If mshDownTab.Tag = "" Or mvarEdit = False Then Exit Function
        If CheckTimeRange(mshDownTab.Col) = False Then Exit Function
        
        If Val(mshDownTab.RowData(mshDownTab.Row)) <> mItemNo.��� Then Exit Function
        
        If picInput.Visible Then picInput.Visible = False
'        mbytSpecChar = 2
'        mshDownTab.Cell(flexcpData, mshDownTab.Row, mshDownTab.Col) = 2
        Call WriteDownTab(mshDownTab.Row, mshDownTab.Col, "E")
        mshDownTab.SetFocus
        
        Call mshDownTab_RowColChange

    '------------------------------------------------------------------------------------------------------------------
    Case "�೦����й"
        
        If Val(mrsParam("�༭")) = 0 Then Exit Function
        If mshDownTab.Tag = "" Or mvarEdit = False Then Exit Function
        If CheckTimeRange(mshDownTab.Col) = False Then Exit Function
        
        If Val(mshDownTab.RowData(mshDownTab.Row)) <> mItemNo.��� Then Exit Function
'        mbytSpecChar = 3
'        mshDownTab.Cell(flexcpData, mshDownTab.Row, mshDownTab.Col) = 3
        
        Call WriteDownTab(mshDownTab.Row, mshDownTab.Col, "/E")
        
'        strTmp = Trim(mshDownTab.TextMatrix(mshDownTab.Row, mshDownTab.Col))

        If picInput.Visible Then
'            If Right(strTmp, 2) <> "/E" And strTmp <> "" Then
'                mshDownTab.TextMatrix(mshDownTab.Row, mshDownTab.Col) = strTmp & "/E"
'            End If

            mshDownTab.SetFocus

        Else
            Call ShowInput
        End If
        
        If txtInput(0).Visible Then
            txtInput(0).SelStart = 0
            txtInput(0).SelLength = 0
        End If
        
        Call mshDownTab_RowColChange

    '------------------------------------------------------------------------------------------------------------------
    Case "����"
        
        If Val(mrsParam("�༭")) = 0 Then Exit Function
        If mshDownTab.Tag = "" Or mvarEdit = False Then Exit Function
        If CheckTimeRange(mshDownTab.Col) = False Then Exit Function
        
        If Val(mshDownTab.RowData(mshDownTab.Row)) <> mItemNo.��Һ Then Exit Function
        
        If picInput.Visible Then picInput.Visible = False
'        mbytSpecChar = 4
'        mshDownTab.Cell(flexcpData, mshDownTab.Row, mshDownTab.Col) = 4
        Call WriteDownTab(mshDownTab.Row, mshDownTab.Col, "C")
        mshDownTab.SetFocus
        
        Call mshDownTab_RowColChange

    '------------------------------------------------------------------------------------------------------------------
    Case "��������"
        
        If Val(mrsParam("�༭")) = 0 Then Exit Function
        If mshDownTab.Tag = "" Or mvarEdit = False Then Exit Function
        If CheckTimeRange(mshDownTab.Col) = False Then Exit Function
        
        If Val(mshDownTab.RowData(mshDownTab.Row)) <> mItemNo.��Һ Then Exit Function
'        mbytSpecChar = 5
'        mshDownTab.Cell(flexcpData, mshDownTab.Row, mshDownTab.Col) = 5

        Call WriteDownTab(mshDownTab.Row, mshDownTab.Col, "1/C")
        
        strTmp = Trim(mshDownTab.TextMatrix(mshDownTab.Row, mshDownTab.Col))
        
        If picInput.Visible Then
'            If Right(strTmp, 2) <> "/C" And strTmp <> "" Then
'                If Val(strTmp) > 0 Then
'                    mshDownTab.TextMatrix(mshDownTab.Row, mshDownTab.Col) = Val(strTmp) & "/C"
'                Else
'                    mshDownTab.TextMatrix(mshDownTab.Row, mshDownTab.Col) = ""
'                End If
'            End If
            mshDownTab.SetFocus
        Else
            Call ShowInput
        End If
        If txtInput(0).Visible Then
            txtInput(0).SelStart = 0
            txtInput(0).SelLength = 0
        End If
        
        Call mshDownTab_RowColChange

    '------------------------------------------------------------------------------------------------------------------
    Case "��ʾ��������"
    
        Select Case Val(mrsParam("Ӥ��").Value)
        Case 0
            txtCard(0).Text = txtCard(0).Tag
            txtCard(7).Text = txtCard(7).Tag
        Case Else
            
            txtCard(5).Text = ""
            txtCard(6).Text = ""
            txtCard(7).Text = ""
            
            gstrSQL = "Select Decode(a.Ӥ������,Null,b.����||'֮��'||Trim(To_Char(a.���,'9')),a.Ӥ������) As Ӥ������,a.Ӥ���Ա�,a.����ʱ�� From ������������¼ a,������Ϣ b Where a.����id=[1] And a.��ҳid=[2] And a.����id=b.����id And a.���=[3]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "usrBodyEditor", Val(mrsParam("����id").Value), Val(mrsParam("��ҳid").Value), Val(mrsParam("Ӥ��").Value))
            If rs.BOF = False Then
            
                txtCard(0).Text = rs("Ӥ������").Value
                txtCard(5).Text = rs("Ӥ���Ա�").Value
                
                txtCard(6).Text = "������"
'                If IsNull(rs("����ʱ��").Value) = False Then
'                    txtCard(6).Text = DateDiff("d", rs("����ʱ��").Value, zlDatabase.Currentdate) & "��"
'                End If
                
            End If
            
        End Select
    
    End Select
        
End Function

Public Function PrintState(ByVal intPrintRange As Integer, ByVal blnPrint As Boolean, Optional lngBeginY As Long, _
    Optional ByVal intPageNo As Integer = -1, Optional ByVal strPrintDevice As String) As Boolean
    '******************************************************************************************************************
    '����:����ǰ���±��ǰ��ʼ���������±��������ӡ���ϻ�Ԥ������
    '����:blnCurState = �Ƿ�Ϊֻ��ӡ��ǰ���±�,�����ӡ�ӵ�ǰ��ʼ���������±�
    '     blnPrint    = �Ƿ��������ӡ���Ϸ��������Ԥ��������
    '******************************************************************************************************************
    
    Dim i As Long
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim strPrintName As String
    Dim intPage As Integer
    Dim blnYesPrinter As Boolean
    Dim intCol As Integer
    Dim intBeginPage As Integer
    Dim intEndPage As Integer
'    Dim intPageNo As Integer
    Dim byeReturn As Byte
    Dim strArrFromTo() As String
    Dim intOrient As Integer
    Dim intBaby As Integer
    Dim strDateFrom As String
    Dim strDateTo As String
    Dim lngIndex As Long
    
    On Error GoTo ErrHandle
    
    intBaby = Val(mrsParam("Ӥ��").Value)
    
    '------------------------------------------------------------------------------------------------------------------
    '��ӡ���ָ�������
    If Not ExistsPrinter Then
        MsgBox "ϵͳû�а�װ�κδ�ӡ�����ܼ�����ӡ�������˳���", vbInformation, gstrSysName
        Exit Function
    End If
    
    If strPrintDevice = "" Then
        If Trim(zlDatabase.GetPara("���µ���ӡ��", glngSys, 1255, "")) = "" Then
            MsgBox "û�����ô�ӡ��,��ʹ��ϵͳĬ�ϴ�ӡ�����ã�", vbInformation, gstrSysName
        Else
            strPrintName = Trim(zlDatabase.GetPara("���µ���ӡ��", glngSys, 1255, Printer.DeviceName))
            '��ӡ��
            blnYesPrinter = False
            If Printer.DeviceName <> strPrintName Then
                For i = 0 To Printers.Count - 1
                    If Printers(i).DeviceName = strPrintName Then Set Printer = Printers(i): blnYesPrinter = True: Exit For
                Next
                If blnYesPrinter = False Then
                    MsgBox "���õĴ�ӡ���Ѳ�����,��ʹ��ϵͳĬ�ϴ�ӡ�����ã�", vbInformation, gstrSysName
                End If
            End If
        End If
    Else
        strPrintName = strPrintDevice
    End If
        
    intPage = Val(zlDatabase.GetPara("���µ�ֽ��", glngSys, 1255, Printer.PaperSize))
    lngWidth = Val(zlDatabase.GetPara("���µ����", glngSys, 1255, Printer.Width))
    lngHeight = Val(zlDatabase.GetPara("���µ��߶�", glngSys, 1255, Printer.Height))
    lngLeft = Val(zlDatabase.GetPara("���µ���߾�", glngSys, 1255, OFFSET_LEFT))
    lngTop = Val(zlDatabase.GetPara("���µ��ϱ߾�", glngSys, 1255, OFFSET_TOP))
    intOrient = Val(zlDatabase.GetPara("���µ�ֽ��", glngSys, 1255, Printer.Orientation))
    
    On Error Resume Next
    'ֽ��
    If intPage = 256 Then
        Printer.PaperSize = 256
        Printer.Width = lngWidth
        Printer.Height = lngHeight
    Else
        Printer.PaperSize = intPage
    End If
    Printer.Orientation = intOrient
    
    On Error GoTo ErrHandle
    
    '------------------------------------------------------------------------------------------------------------------
    lngBeginY = IIf(lngTop > lngBeginY, lngTop, lngBeginY)
    lngIndex = mlngNo
    
    
    '��ȡ�˲��˵����µ���ҳ��
    '------------------------------------------------------------------------------------------------------------------
    strSQL = "Select ��Ժʱ��, ��Ժʱ��, 1 + Round((b.��Ժʱ�� - b.��Ժʱ��) / 7) As ҳ��" & vbNewLine & _
                "  from (Select Min(��ʼʱ��) as ��Ժʱ��," & vbNewLine & _
                "               Max(Nvl(��ֹʱ��, Sysdate)) as ��Ժʱ��" & vbNewLine & _
                "          From ���˱䶯��¼" & vbNewLine & _
                "         Where ��ʼʱ�� is Not Null And ����ID = [1] And ��ҳID = [2]) b"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, mstrMsgTitle, Val(mrsParam("����id")), Val(mrsParam("��ҳid")))
    intCount = 0
    For intCol = 0 To rsTmp("ҳ��").Value - 1
                
        strDateFrom = Format(rsTmp("��Ժʱ��").Value + 7 * intCol, "yyyy-MM-dd") & " 00:00:00"
        strDateTo = Format(rsTmp("��Ժʱ��").Value + 7 * (intCol + 1) - 1, "yyyy-MM-dd") & " 23:59:59"
        If strDateFrom < Format(rsTmp("��Ժʱ��").Value, "yyyy-MM-dd HH:mm:ss") Then
            strDateFrom = Format(rsTmp("��Ժʱ��").Value, "yyyy-MM-dd HH:mm:ss")
        End If
        
        If strDateFrom < Format(rsTmp("��Ժʱ��").Value, "yyyy-MM-dd HH:mm:ss") Then
        
            If strDateFrom < Format(rsTmp("��Ժʱ��").Value, "yyyy-MM-dd HH:mm:ss") Then strDateFrom = Format(rsTmp("��Ժʱ��").Value, "yyyy-MM-dd HH:mm:ss")
            If strDateTo > Format(rsTmp("��Ժʱ��").Value, "yyyy-MM-dd HH:mm:ss") Then strDateTo = Format(rsTmp("��Ժʱ��").Value, "yyyy-MM-dd HH:mm:ss")
            
            ReDim Preserve strArrFromTo(intCount)
            strArrFromTo(intCount) = "0;" & intCol + 1 & ";" & intCol + 1
            intCount = intCount + 1
        End If
    Next
        
    '���ֻ��ӡ��ǰ��ֻ����ʼ�ͽ���дͬһҳ��
    Set mfrmCaseTendBodyPrint = New frmCaseTendBodyPrint
    Select Case intPrintRange
    Case 0                  '��ӡ��ǰҳ
        
        If PrintOrPreviewBodyState(mfrmCaseTendBodyPrint, Val(mrsParam("����id").Value), Val(mrsParam("��ҳid").Value), intBaby, _
                Val(mrsParam("����id").Value), lngBeginY * 56.7, lngLeft, Me, False, _
                CInt(Split(strArrFromTo(lngIndex), ";")(1)), CInt(Split(strArrFromTo(lngIndex), ";")(1)), intPageNo, , mblnMoved) = True Then
                
                If blnPrint = False Then
                    mfrmCaseTendBodyPrint.Preview intPrintRange, lngBeginY, lngLeft, Me, Val(mrsParam("����id")), Val(mrsParam("��ҳid")), _
                        Val(mrsParam("����id").Value), CInt(Split(strArrFromTo(lngIndex), ";")(1)), _
                        CInt(Split(strArrFromTo(lngIndex), ";")(1)), intPageNo, strArrFromTo, lngIndex
                Else
                    Printer.PaintPicture mfrmCaseTendBodyPrint.picPage(mfrmCaseTendBodyPrint.picPage.UBound).Image, 0, 0
                    Printer.EndDoc
                End If
        Else
            MsgBox "δ֪����������µ�ʧ�ܣ�", vbExclamation, gstrSysName
        End If
        
    Case 1              '�ӵ�ǰҳ������ӡ
    
        For intCol = lngIndex To UBound(strArrFromTo)
        
            If PrintOrPreviewBodyState(mfrmCaseTendBodyPrint, Val(mrsParam("����id")), Val(mrsParam("��ҳid")), intBaby, _
                Val(mrsParam("����id").Value), lngBeginY * 56.7, lngLeft, Me, intCol <> lngIndex, _
                CInt(Split(strArrFromTo(intCol), ";")(1)), CInt(Split(strArrFromTo(intCol), ";")(1)), intPageNo, , mblnMoved) = True Then
            Else
                MsgBox "δ֪���󣬴�ӡʧ�ܣ�", vbExclamation, gstrSysName
                Exit For
            End If
            
            If blnPrint Then
                Printer.PaintPicture mfrmCaseTendBodyPrint.picPage(mfrmCaseTendBodyPrint.picPage.UBound).Image, 0, 0
                If intCol = UBound(strArrFromTo) Then
                    Printer.EndDoc
                Else
                    Printer.NewPage
                End If
            End If
        Next

        If blnPrint = False Then
            mfrmCaseTendBodyPrint.Preview intPrintRange, lngBeginY, lngLeft, Me, Val(mrsParam("����id")), Val(mrsParam("��ҳid")), _
            Val(mrsParam("����id").Value), CInt(Split(strArrFromTo(lngIndex), ";")(1)), _
                CInt(Split(strArrFromTo(lngIndex), ";")(1)), intPageNo, strArrFromTo, lngIndex
        End If
        
    Case 2          '�ӵ�һҳ������ӡ,��ȫ����ӡ
        
        For intCol = 0 To UBound(strArrFromTo)
        
            If PrintOrPreviewBodyState(mfrmCaseTendBodyPrint, Val(mrsParam("����id")), Val(mrsParam("��ҳid")), intBaby, _
                Val(mrsParam("����id").Value), lngBeginY * 56.7, lngLeft, Me, intCol <> 0, _
                CInt(Split(strArrFromTo(intCol), ";")(1)), CInt(Split(strArrFromTo(intCol), ";")(1)), intPageNo, , mblnMoved) = True Then
            Else
                MsgBox "δ֪���󣬴�ӡʧ�ܣ�", vbExclamation, gstrSysName
                Exit For
            End If
            
            If blnPrint Then
                Printer.PaintPicture mfrmCaseTendBodyPrint.picPage(mfrmCaseTendBodyPrint.picPage.UBound).Image, 0, 0
                If intCol = UBound(strArrFromTo) Then
                    Printer.EndDoc
                Else
                    Printer.NewPage
                End If
            End If
        Next

        If blnPrint = False Then
            mfrmCaseTendBodyPrint.Preview intPrintRange, lngBeginY, lngLeft, Me, Val(mrsParam("����id")), Val(mrsParam("��ҳid")), _
            Val(mrsParam("����id").Value), CInt(Split(strArrFromTo(0), ";")(1)), _
                CInt(Split(strArrFromTo(0), ";")(1)), intPageNo, strArrFromTo, 0
        End If
        
    End Select
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function AllowAudit() As Boolean
    '******************************************************************************************************************
    '���ܣ�����Ƿ��������־
    '��������
    '���أ�
    '******************************************************************************************************************
    Dim intNowCol As Integer
    Dim aryValue As Variant

    If picScale.Tag <> "" Then

        With mshScale
            intNowCol = .FixedCols + lblCur.Left \ HOUR_STEP_Twips
            
            If Split(.TextMatrix(GraphDataRow.��������, intNowCol), ";")(mItemSerial.���� + 1) <> "" Then
                aryValue = Split(Split(.TextMatrix(GraphDataRow.��������, intNowCol), ";")(mItemSerial.���� + 1), ",")
                AllowAudit = (Val(.TextMatrix(GraphDataRow.���Ա�־, intNowCol)) = 0 And Val(aryValue(0)) > 0)
            End If

        End With
    End If
End Function

Public Function AllowUnAudit() As Boolean
    '******************************************************************************************************************
    '���ܣ�����Ƿ������������־
    '��������
    '���أ�
    '******************************************************************************************************************
    Dim intNowCol As Integer
    Dim aryValue As Variant
    
    If picScale.Tag <> "" Then

        With mshScale
            intNowCol = .FixedCols + lblCur.Left \ HOUR_STEP_Twips
            
            If Split(.TextMatrix(GraphDataRow.��������, intNowCol), ";")(mItemSerial.���� + 1) <> "" Then
                aryValue = Split(Split(.TextMatrix(GraphDataRow.��������, intNowCol), ";")(mItemSerial.���� + 1), ",")
                AllowUnAudit = (.TextMatrix(GraphDataRow.���Ա�־, intNowCol) = "1" And Val(aryValue(0)) > 0)
            End If
            
        End With
    End If
End Function

Private Function InitCommandBar() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom

    '------------------------------------------------------------------------------------------------------------------
    '��ʼ����
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    cbsMain.ActiveMenuBar.Title = "�˵���"
    
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = False
        .SetIconSize False, 16, 16
        .UseSharedImageList = False 'ImageList��ʽʱ,��ͬһApp�й���,��AddImageList֮ǰ����ΪFalse
    End With

    '------------------------------------------------------------------------------------------------------------------
    '�˵�����:�����������ݣ����xtpControlPopup���͵�����ID���¸�ֵ

    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagHideWrap
    cbsMain.ActiveMenuBar.Visible = False
    
    '------------------------------------------------------------------------------------------------------------------
    '����������:������������
    
    Set mcbrToolBar = cbsMain.Add("Ӥ��", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.EnableDocking xtpFlagHideWrap
    
    Set objCustom = mcbrToolBar.Controls.Add(xtpControlCustom, conMenu_View_Option, "")
    picTmp.Visible = True
    objCustom.Handle = picTmp.hWnd

End Function

Private Function InitBody(ByVal lng����id As Long, ByVal lng��ҳid As Long, ByVal lng����id As Long, ByVal intӤ�� As Integer) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim cbrItem As CommandBarControl
    Dim intCount As Integer
    Dim strDateFrom As String
    Dim strDateTo As String
    Dim strEnterDate As String
    Dim intCol As Integer
    Dim strCaption As String
    Dim strParameter As String
    Dim strSvrCaption As String
    Dim strNow As String
    Dim strCut As String
    Dim lngLoop As Long
    Dim strTmp As String
    Dim lnglast����id As Long
    
    If lng����id = 0 Then Exit Function
    strCut = "123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    strNow = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    'ɾ������ҳ��˵���
    
    If Not mcbrToolBarҳ�� Is Nothing Then mcbrToolBarҳ��.Delete
    Set mcbrToolBarҳ�� = mcbrToolBar.Controls.Add(xtpControlPopup, conMenu_Edit_NewItem, "ҳ��"):  mcbrToolBarҳ��.BeginGroup = True
    mcbrToolBarҳ��.IconId = conMenu_Edit_Modify
    mcbrToolBarҳ��.Style = xtpButtonIconAndCaption
    
    '------------------------------------------------------------------------------------------------------------------
    strSQL = "Select Decode(c.����ʱ��,Null,b.��Ժʱ��,c.����ʱ��) As ��Ժʱ��, ��Ժʱ��, 1 + Round((b.��Ժʱ�� - Decode(c.����ʱ��,Null,b.��Ժʱ��,c.����ʱ��)) / 7) As ҳ��" & vbNewLine & _
                "  from (Select ����ID,��ҳid,Min(��ʼʱ��) as ��Ժʱ��," & vbNewLine & _
                "               Max(Nvl(��ֹʱ��, Sysdate)) as ��Ժʱ��" & vbNewLine & _
                "          From ���˱䶯��¼" & vbNewLine & _
                "         Where ��ʼʱ�� is Not Null And ����ID = [1] And ��ҳID = [2] Group By ����ID,��ҳid) b," & vbNewLine & _
                "       (Select ����ID,��ҳid,����ʱ�� From ������������¼ Where ����ID = [1] And ��ҳID = [2] And ���=[3]) c Where b.����id=c.����id(+) And b.��ҳid=c.��ҳid(+)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "usrBodyEditor", lng����id, lng��ҳid, intӤ��)
    If rsTmp.BOF Then
        MsgBox "�޲��˱���סԺ��¼��", vbExclamation, gstrSysName
        Exit Function
    End If
    
    strEnterDate = Format(rsTmp!��Ժʱ��, "yyyy-MM-dd HH:mm:ss")

    '------------------------------------------------------------------------------------------------------------------
    strSQL = "Select 1 + Round((a.��ʼʱ�� - b.��Ժʱ��) / 7) As ��ʼҳ��,1 + Round((a.��ֹʱ�� - b.��Ժʱ��) / 7) As ����ҳ��,b.��Ժʱ��," & vbNewLine & _
                "       ����id,c.����," & vbNewLine & _
                "       ��ʼʱ��," & vbNewLine & _
                "       ��ֹʱ��" & vbNewLine & _
                "  from (Select ����id," & vbNewLine & _
                "               Min(��ʼʱ��) as ��ʼʱ��," & vbNewLine & _
                "               Max(Nvl(��ֹʱ��, Sysdate)) as ��ֹʱ��" & vbNewLine & _
                "          From ���˱䶯��¼" & vbNewLine & _
                "         Where ��ʼʱ�� is Not Null And ����ID = [1] And ��ҳID = [2]" & vbNewLine & _
                "         Group by ����id) a," & vbNewLine & _
                "       (Select Decode(y.����ʱ��,Null,x.��Ժʱ��,y.����ʱ��) As ��Ժʱ�� From (Select ����ID,��ҳid,Min(��ʼʱ��) as ��Ժʱ��" & vbNewLine & _
                "          From ���˱䶯��¼" & vbNewLine & _
                "         Where ��ʼʱ�� is Not Null And ����ID = [1] And ��ҳID = [2] Group By ����id,��ҳid) x,(Select ����ID,��ҳid,����ʱ�� From ������������¼ Where ����ID = [1] And ��ҳID = [2] And ���=[3]) y Where x.����id=y.����id(+) And x.��ҳid=y.��ҳid(+) ) b,���ű� c Where c.ID=a.����id " & vbNewLine & _
                " order by a.��ʼʱ��"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, "usrBodyEditor", lng����id, lng��ҳid, intӤ��)
        
    For lngLoop = 0 To rsTmp("ҳ��").Value - 1
                
        strDateFrom = Format(rsTmp("��Ժʱ��").Value + 7 * lngLoop, "yyyy-MM-dd") & " 00:00:00"
        strDateTo = Format(rsTmp("��Ժʱ��").Value + 7 * (lngLoop + 1) - 1, "yyyy-MM-dd") & " 23:59:59"
        If strDateFrom < Format(rsTmp("��Ժʱ��").Value, "yyyy-MM-dd HH:mm:ss") Then
            strDateFrom = Format(rsTmp("��Ժʱ��").Value, "yyyy-MM-dd HH:mm:ss")
        End If
        
        If strDateFrom < Format(rsTmp("��Ժʱ��").Value, "yyyy-MM-dd HH:mm:ss") Then
        
            If strDateFrom < Format(rsTmp("��Ժʱ��").Value, "yyyy-MM-dd HH:mm:ss") Then strDateFrom = Format(rsTmp("��Ժʱ��").Value, "yyyy-MM-dd HH:mm:ss")
            If strDateTo > Format(rsTmp("��Ժʱ��").Value, "yyyy-MM-dd HH:mm:ss") Then strDateTo = Format(rsTmp("��Ժʱ��").Value, "yyyy-MM-dd HH:mm:ss")
    
            rs.Filter = ""
            rs.Filter = "��ʼҳ��<=" & lngLoop + 1 & " And ����ҳ��>=" & lngLoop + 1
            If rs.RecordCount > 0 Then rs.MoveFirst
            For intCol = 1 To rs.RecordCount
                
                If strDateFrom < Format(rs("��ʼʱ��").Value, "yyyy-MM-dd HH:mm:ss") Then
                    strTmp = Format(rs("��ʼʱ��").Value, "yyyy-MM-dd HH:mm:ss")
                Else
                    strTmp = strDateFrom
                End If
                
                If strDateTo > Format(rs("��ֹʱ��").Value, "yyyy-MM-dd HH:mm:ss") Then
                    strCaption = Format(rs("��ֹʱ��").Value, "yyyy-MM-dd HH:mm:ss")
                Else
                    strCaption = strDateTo
                End If
                
                strCaption = Format(strTmp, "yyyy-MM-dd") & "��" & Format(strCaption, "yyyy-MM-dd")
                strCaption = "��" & lngLoop + 1 & "ҳ��" & strCaption & "(" & rs("����").Value & ")"
                
                '��Ժʱ��;����id;��ʼʱ��;����ʱ��;
                Set cbrItem = mcbrToolBarҳ��.CommandBar.Controls.Add(xtpControlButton, conMenu_View_Jump, strCaption, -1, False)
                cbrItem.Parameter = strEnterDate & ";" & rs!����ID & ";" & strDateFrom & ";" & strDateTo & ";" & lngLoop
                
                lnglast����id = rs("����ID").Value
                
                rs.MoveNext
                
                strParameter = cbrItem.Parameter
                strSvrCaption = strCaption
            Next
        End If
        
    Next
    
    Call picPane_Resize
    
    If strParameter <> "" Then
        mcbrToolBarҳ��.Caption = strSvrCaption
        Call zlMenuClick("װ������", strParameter)
    End If
    
    InitBody = True
End Function


Private Function CheckTimeRange(ByVal intCol As Integer) As Boolean
    Dim strTime As String
    Dim strFrom As String
    Dim strTo As String

    Dim strEnd As String
    Dim strStart As String
    
    If picScale.Tag = "" Then Exit Function
    If InStr(picScale.Tag, ";") = 0 Then Exit Function
    
    strFrom = Split(picScale.Tag, ";")(0)
    strTo = Split(picScale.Tag, ";")(1)
    
    If strTo > Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") Then strTo = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    strTime = GetEditDateTime(intCol - mshDownTab.FixedCols + 1, CDate(strFrom))
    strStart = Split(strTime, ",")(0)
    strEnd = Split(strTime, ",")(1)
    
    CheckTimeRange = False
    
    If strStart <= strFrom And strEnd >= strTo Then
        CheckTimeRange = True
    End If
    
    If strStart >= strFrom And strStart <= strTo Then
        CheckTimeRange = True
    End If
    
    If strEnd > strFrom And strEnd < strTo Then
        CheckTimeRange = True
    End If

End Function

Private Function GetTextPos(ByVal lngHwnd As Long) As Long

    Dim lngRow As Long
    Dim lngCol As Long
    Dim lngFirst As Long
    
    lngFirst = SendMessage(lngHwnd, EM_GETFIRSTVISIBLELINE, lngRow, lngCol) + 1 '��0�п�ʼ
    Call GetCaretPos(lngHwnd, lngRow, lngCol)
    GetTextPos = lngCol
    
End Function

Private Sub GetCaretPos(ByVal TextHwnd As Long, LineNo As Long, ColNo As Long)
    Dim i As Long, j As Long, k As Long
    Dim lParam As Long, wParam As Long

    '�������ı��򴫵�EM_GETSEL��Ϣ�Ի�ȡ����ʼλ�õ�
    '�������λ�õ��ַ���
    i = SendMessage(TextHwnd, EM_GETSEL, wParam, lParam)
    j = i / 2 ^ 16
    
    '�����ı��򴫵�EM_LINEFROMCHAR��Ϣ���ݻ�õ��ַ�
    '��ȷ������Ի�ȡ��������
    LineNo = SendMessage(TextHwnd, EM_LINEFROMCHAR, j, 0) '
    LineNo = LineNo + 1
    
    '���ı��򴫵�EM_LINEINDEX��Ϣ�Ի�ȡ��������
    k = SendMessage(TextHwnd, EM_LINEINDEX, -1, 0)
    ColNo = j - k + 1
End Sub

Private Function ReadDrink(ByVal strDate As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ�
    '������
    '���أ�
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    Dim intCol As Long
    Dim strTmp As String
    Dim strFrom As String
    
    Dim strStart As String
    Dim strEnd As String
    Dim lng������id As Long
    Dim lng������id As Long
    Dim strValue As String
    Dim int��¼�� As Integer
    Dim intMax As Integer
    Dim lngCol As Long
    
    On Error GoTo errHand
    
    strFrom = CStr(mrsParam("��ʼʱ��"))
    
    strSQL = "Select A.��Ŀid From �����¼��Ŀ A Where A.��Ŀ���=[1] "
    Set rs = zlDatabase.OpenSQLRecord(strSQL, mstrMsgTitle, 6)
    If rs.BOF = False Then lng������id = zlCommFun.NVL(rs("��Ŀid"), 0)
    
    strSQL = "Select A.��Ŀid,B.��¼�� From �����¼��Ŀ A,���¼�¼��Ŀ B Where A.��Ŀ���=B.��Ŀ��� AND A.��Ŀ���=[1] "
    Set rs = zlDatabase.OpenSQLRecord(strSQL, mstrMsgTitle, 7)
    If rs.BOF = False Then
        lng������id = zlCommFun.NVL(rs("��Ŀid"), 0)
        int��¼�� = zlCommFun.NVL(rs("��¼��"), 1)
    End If
    
    If int��¼�� = 1 Then
        intMax = 6
    Else
        intMax = 2
    End If
    
    For intCol = 0 To intMax - 1
            
        strStart = Format(Int(CDate(strDate)) + intCol / intMax - (4 - mlngHourBegin) / 24, "YYYY-MM-DD hh:mm:ss")
        strEnd = Format(Int(CDate(strDate)) + intCol / intMax - (4 - mlngHourBegin) / 24 + 1 / intMax, "YYYY-MM-DD hh:mm:ss")
        
        If Int(CDate(strStart)) < Int(CDate(strDate)) Then
            strStart = Format(strDate, "yyyy-MM-dd HH:mm:ss")
        End If
        
        strSQL = "Select zl_PatitDrink([1],[2],[3],[4]) As ���� From Dual"
        
        Set rs = zlDatabase.OpenSQLRecord(strSQL, mstrMsgTitle, Val(mrsParam("����id")), Val(mrsParam("��ҳid")), CDate(strStart), CDate(strEnd))
        If rs.BOF = False Then
            
            strTmp = zlCommFun.NVL(rs("����"))
            
            If strTmp <> "" Then
                strValue = Trim(Split(strTmp, ";")(0))
                
                mstrSQL = "ZL_���ӻ����¼_UPDATE("
                mstrSQL = mstrSQL & Val(mrsParam("����id")) & ","
                mstrSQL = mstrSQL & Val(mrsParam("��ҳid")) & ","
                mstrSQL = mstrSQL & Val(mrsParam("Ӥ��")) & ","
                mstrSQL = mstrSQL & "To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'),"
                mstrSQL = mstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
                mstrSQL = mstrSQL & "1,"
                mstrSQL = mstrSQL & "7,"
                mstrSQL = mstrSQL & "0,"
                    
                mstrSQL = mstrSQL & IIf(Val(strValue) = 0, "NULL", "'" & Val(strValue) & "'")
                
                mstrSQL = mstrSQL & ")"

                Call zlDatabase.ExecuteProcedure(mstrSQL, mstrMsgTitle)
                
                If mItemSerial.������ >= 0 Then
                    
                    If int��¼�� = 2 Then
                        
                        lngCol = intCol + (Int(CDate(strStart)) - Int(CDate(strFrom)) + (4 - mlngHourBegin) / 24) * intMax + mshDownTab.FixedCols
                        
                        Call WriteDownTab(mItemSerial.������, lngCol, strValue)
                    Else
                                                
                        lngCol = intCol + (Int(CDate(strStart)) - Int(CDate(strFrom)) + (4 - mlngHourBegin) / 24) * intMax + mshScale.FixedCols
                        
                        Call WriteScaleTab(mItemSerial.������, lngCol, strValue)
                    End If
                    
                End If
                
                strValue = Trim(Split(strTmp, ";")(1))
                If UBound(Split(strTmp, ";")) > 1 Then strValue = strValue & "��"
                            
                mstrSQL = "ZL_���ӻ����¼_UPDATE("
                mstrSQL = mstrSQL & Val(mrsParam("����id")) & ","
                mstrSQL = mstrSQL & Val(mrsParam("��ҳid")) & ","
                mstrSQL = mstrSQL & Val(mrsParam("Ӥ��")) & ","
                mstrSQL = mstrSQL & "To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'),"
                mstrSQL = mstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
                mstrSQL = mstrSQL & "1,"
                mstrSQL = mstrSQL & "6,"
                mstrSQL = mstrSQL & "0,"
                    
                mstrSQL = mstrSQL & IIf(strValue = "", "NULL", "'" & strValue & "'")
                
                mstrSQL = mstrSQL & ")"

                Call zlDatabase.ExecuteProcedure(mstrSQL, mstrMsgTitle)
                
                If mItemSerial.������ >= 0 Then
                    If int��¼�� = 2 Then
                    
                        lngCol = intCol + (Int(CDate(strStart)) - Int(CDate(strFrom)) + (4 - mlngHourBegin) / 24) * intMax + mshDownTab.FixedCols
                        Call WriteDownTab(mItemSerial.������, lngCol, strValue)
                        
                    Else
                        
                        lngCol = intCol + (Int(CDate(strStart)) - Int(CDate(strFrom)) + (4 - mlngHourBegin) / 24) * intMax + mshScale.FixedCols
                        Call WriteScaleTab(mItemSerial.������, lngCol, strValue)
                        
                    End If
                End If
                
            End If
        End If
    Next
        
    ReadDrink = True
                            
    Exit Function
    
errHand:
'    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CalcScrollBarSize() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ�
    '������
    '���أ� ���óɹ�����TRUE������FALSE
    '------------------------------------------------------------------------------------------------------------------
    On Error Resume Next
    
    'ֻ����û��ʾ�������ǲ��������㲽��
    msinHStep = (pic.Width - picPane.Width) / 100
    msinVStep = (pic.Height - picPane.Height) / 100
    
    hsb.Max = 0 - Int(0 - ((pic.Width - picPane.Width) / 300))
    vsb.Max = 0 - Int(0 - ((pic.Height - picPane.Height) / 300))
    hsb.Enabled = (hsb.Max > 0)
    vsb.Enabled = (vsb.Max > 0)
    
    '�㶨Ϊ100,ֻ�ǲ��������仯
    If hsb.Enabled Then hsb.Max = 100
    If vsb.Enabled Then vsb.Max = 100
    
    CalcScrollBarSize = True
    
End Function

Private Function Check�Ƿ����(strSource As String, strTarge As String) As Boolean
    '���strSource�е�ÿһ���ַ��Ƿ���strTarge��
    Dim i As Long
    Check�Ƿ���� = False
    
    Select Case strTarge
    Case "����"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "С��"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "������"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+-_)(*&^%$#@!`~"
    Case "��С��"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/'"":;|\=+-_)(*&^%$#@!`~"
    End Select
    For i = 1 To Len(strSource)
        If InStr(strTarge, Mid(strSource, i, 1)) <= 0 Then Exit Function
    Next
    Check�Ƿ���� = True
End Function

Private Sub ClearSpecRowCol(obj As Object, ByVal intRow As Integer, Optional intCol As Variant)
    '����: ���ָ�������ָ����ָ���е�����
    '����: obj=Ҫ����������ؼ�
    '      intRow=Ҫ������к�
    '      intCol=Ҫ������к��б���Array(1,2,3),������������Ա�ʾΪArray()
    Dim i As Long
    If UBound(intCol) = -1 Then
        For i = 0 To obj.Cols - 1
            obj.TextMatrix(intRow, i) = ""
        Next
    Else
        For i = 0 To UBound(intCol)
            obj.TextMatrix(intRow, intCol(i)) = ""
        Next
    End If
    obj.RowData(intRow) = 0
End Sub

Private Sub SetColumnText(fgd As Object, intRow As Integer, ByVal varColText As Variant)
    '����: ����ָ������ؼ�����ͷ�ı�
    '����: fgd=����ؼ�
    '      intRow=�к�
    '      varColText=��ͷ�ı�����
    Dim i As Integer
    For i = 0 To fgd.Cols - 1
        fgd.TextMatrix(intRow, i) = varColText(i)
    Next
End Sub

Private Sub SetColAlignment(fgd As Object, varColAlignment As Variant)
    '����: ����ָ������ؼ����ж��뷽ʽ
    '����: fgd=����ؼ�
    '      varColAlignment=�ж��뷽ʽ����
    Dim i As Long
    For i = 0 To UBound(varColAlignment)
        fgd.ColAlignment(i) = varColAlignment(i)
    Next
End Sub

Private Sub SetColData(fgd As Object, varColData As Variant)
    '����: ����ָ������ؼ�����������Դ��ʽ
    '����: fgd=����ؼ�
    '      varColData=��������Դ��ʽ����
    Dim i As Long
    For i = 0 To UBound(varColData)
        fgd.ColData(i) = varColData(i)
    Next
End Sub

Private Sub SetFixColAlignment(fgd As Object, varFixColAlignment As Variant)
    '����: ����ָ������ؼ��Ĺ̶��ж��뷽ʽ
    '����: fgd=����ؼ�
    '      varColAlignment=�̶��ж��뷽ʽ����
    Dim i As Long
    For i = 0 To UBound(varFixColAlignment)
        fgd.ColAlignmentFixed(i) = varFixColAlignment(i)
    Next
End Sub

Private Sub SetColumnWidth(fgd As Object, ByVal varColWidth As Variant)
    '����: ����ָ������ؼ����п�
    '����: fgd=����ؼ�
    '      varColWidth=�п�����
    Dim i As Integer
    For i = 0 To fgd.Cols - 1
        fgd.ColWidth(i) = varColWidth(i)
    Next
End Sub

Public Function SetDispMode(Optional blnReadOnly As Boolean) As Boolean
    
    '�����������±�ǰ�Ǳ༭ģʽ������ʾģʽ
    
    Call SetBodyMode
    
End Function


Private Function InitData() As Boolean
    '******************************************************************************************************************
    '���ܣ������������µ��Ľӿں���
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim i As Long
    Dim strSQL As String
    Dim strTmp As String
    
    On Error GoTo ErrHandle
    '------------------------------------------------------------------------------------------------------------------
    '������ʼ��
    
    mlngLine = 0
    mlngPageCur = 1
    
    '------------------------------------------------------------------------------------------------------------------
    With pic
        .Left = 0
        .Top = 0
        .Width = UserControl.Width
        .Height = UserControl.Height
    End With
    
    mstrMsgTitle = "���±�"
    UserControl.BackColor = RGB(255, 255, 255) '�������ð�ɫ
    mblnNoneShow = False
    
    
    
    '��ȡ���±�һ�쿪ʼʱ��
    '------------------------------------------------------------------------------------------------------------------
    mlngHourBegin = zlDatabase.GetPara("���¿�ʼʱ��", glngSys, 1255, 4)
    mblnӤ�����µ���ʾ��Ժ = (zlDatabase.GetPara("Ӥ�����µ���ʾ��Ժ��Ϣ", glngSys, 1255, 1) = 1)
    
    '���˱䶯�����ʾ����
    '------------------------------------------------------------------------------------------------------------------
    strTmp = zlDatabase.GetPara("���µ����", glngSys, 1255, "1;1;1;1;1;1;1;1")
    If UBound(Split(strTmp, ";")) >= 5 Then
        mBodyFlag.��Ժ = Val(Split(strTmp, ";")(0))
        mBodyFlag.��� = Val(Split(strTmp, ";")(1))
        mBodyFlag.ת�� = Val(Split(strTmp, ";")(2))
        mBodyFlag.���� = Val(Split(strTmp, ";")(3))
        mBodyFlag.���� = Val(Split(strTmp, ";")(4))
        mBodyFlag.��Ժ = Val(Split(strTmp, ";")(5))
        If UBound(Split(strTmp, ";")) >= 6 Then mBodyFlag.���� = Val(Split(strTmp, ";")(6))
        If UBound(Split(strTmp, ";")) >= 7 Then mBodyFlag.���� = Val(Split(strTmp, ";")(7))
    End If
    
    '��ȡ����ȼ�
    '------------------------------------------------------------------------------------------------------------------
    mrsParam("����ȼ�").Value = 3
    mstrSQL = "Select zl_PatitTendGrade([1],[2]) As ����ȼ� From dual"
    Set rs = zlDatabase.OpenSQLRecord(mstrSQL, mstrMsgTitle, Val(mrsParam("����id")), Val(mrsParam("��ҳid")))
    If rs.BOF = False Then
        mrsParam("����ȼ�").Value = zlCommFun.NVL(rs("����ȼ�"), 3)
    End If
    
    '����Ƿ�������������Ŀ
    '------------------------------------------------------------------------------------------------------------------
    mstrSQL = " Select 1 From ���¼�¼��Ŀ A,����������Ŀ B,�����¼��Ŀ C " & _
                "Where C.��Ŀ���=A.��Ŀ��� " & _
                        "AND C.��ĿID=B.ID(+) " & _
                        "AND C.����ȼ�>=[1] " & _
                        "And A.��¼��=1 And RowNum<2 And C.��Ŀ���<>-1 "
                
    Set rs = zlDatabase.OpenSQLRecord(mstrSQL, mstrMsgTitle, Val(mrsParam("����ȼ�")))
    If rs.EOF Then
        ShowSimpleMsg "����Ҫ��һ���Ѽ�¼��������Ŀ��"
        
        If Val(mrsParam("�༭")) = 0 Then
            mblnNoneShow = True
            Exit Function   '��ʾģʽ������������Ŀ
        End If
    End If
    
    '�жϲ����Ƿ���ת��
    '��Ϊ�ú������ⶼ�ڵ���,�������ñ�,ֱ�Ӷ�ȡ
    '------------------------------------------------------------------------------------------------------------------
    mblnMoved = False
    If Val(mrsParam("����id")) > 0 And Val(mrsParam("��Ժ")) = 1 Then
        mstrSQL = "Select ����ת�� From ������ҳ Where ����ID=[1] And ��ҳID=[2]"
        Set rs = zlDatabase.OpenSQLRecord(mstrSQL, mstrMsgTitle, Val(mrsParam("����id")), Val(mrsParam("��ҳid")))
        mblnMoved = NVL(rs!����ת��, 0) <> 0
    End If
    If mblnMoved Or Val(mrsParam("�༭")) = 0 Then Call SetDispMode(True)
    
    
    vsf.Body.Appearance = flexFlat
    vsf.Body.RowHidden(0) = True
    vsf.Body.ColHidden(0) = True
    vsf.Body.ScrollBars = flexScrollBarNone
    vsf.Body.BorderStyle = flexBorderNone
    vsf.FixedCols = 1
    
    vsf.Rows = 2
    
    InitData = True
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Function

Private Function ReadPatiInfo() As Boolean
    '******************************************************************************************************************
    '���ܣ� ��ȡ��ǰ����mlng����ID��סԺ�����סԺ�䶯��¼��������Ϊ����ҳ���±�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo ErrHead
    
    If Val(mrsParam("����id")) = 0 Then Exit Function
    If Val(mrsParam("��ҳid")) = 0 Then Exit Function
    
    '��д����������סԺ��
    gstrSQL = "Select A.����,B.סԺ�� From ������Ϣ A,������ҳ B Where A.����ID=B.����ID And B.����id=[1] And B.��ҳID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, mstrMsgTitle, Val(mrsParam("����id")), Val(mrsParam("��ҳid")))
    If rsTmp.BOF Then
        ShowSimpleMsg "ָ���Ĳ��˲����ڣ�"
        Exit Function
    End If
    
    txtCard(0).Tag = zlCommFun.NVL(rsTmp("����").Value)
    
    Call zlMenuClick("��ʾ��������")
    
    txtCard(0).Tag = Val(mrsParam("����id"))
    txtCard(1).Text = zlCommFun.NVL(rsTmp("סԺ��").Value)
    txtCard(1).Tag = Val(mrsParam("Ӥ��").Value)
    
    gstrSQL = "Select Zl_Replace_Element_Value([1],[2],[3],2) As ������ From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, mstrMsgTitle, "������", Val(mrsParam("����id")), Val(mrsParam("��ҳid")))
    If rsTmp.BOF = False Then
        If Val(mrsParam("Ӥ��").Value) = 0 Then
            txtCard(7).Text = zlCommFun.NVL(rsTmp("������").Value)
        Else
            txtCard(7).Text = ""
        End If
    Else
        txtCard(7).Text = ""
    End If
    txtCard(7).Tag = txtCard(7).Text
    
    mstrSQL = " Select D.ID,D.����,��ʼ,��ֹ" & _
                " From ���ű� D," & _
                "   (Select ����id,Min(��ʼʱ��) as ��ʼ,Max(Nvl(��ֹʱ��,Sysdate)) as ��ֹ" & _
                "    From ���˱䶯��¼" & _
                "    Where ��ʼʱ�� is Not Null And ����ID=[1] And ��ҳID=[2]" & _
                "    Group by ����id) L" & _
                " Where L.����id=D.ID" & _
                " Order by ��ʼ"
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, mstrMsgTitle, Val(mrsParam("����id")), Val(mrsParam("��ҳid")))
    If rsTmp.BOF Then
        
        'ShowSimpleMsg "�޲��˱���סԺ��¼��"
        mblnNoneShow = True
        
        Exit Function
    End If
        
    ReadPatiInfo = True
    
    Exit Function
    
None:
    
    SetVisible
    Exit Function
ErrHead:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ClearLineSelect() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ� �����ǰѡ���������Ŀ
    '������
    '���أ�
    '------------------------------------------------------------------------------------------------------------------
    Dim intLoop As Integer
    
    picGraph.Tag = ""
    picGraph.MousePointer = 0
    linHCur.Visible = False
    linVCur.Visible = False
    
    ClearLineSelect = True
    
End Function

Private Function AddCrlf(ByVal strText As String) As String
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    Dim intLoop As Integer
    Dim strTmp As String
    
    For intLoop = 1 To Len(strText)
        strTmp = strTmp & Mid(strText, intLoop, 1) & vbCrLf
    Next
    
    AddCrlf = strTmp
    
End Function

Private Function FaceInit() As Boolean
    '******************************************************************************************************************
    '���ܣ� �������±����ã��������±�Ĳ���
    '������
    '���أ�
    '******************************************************************************************************************
    Dim i As Long
    Dim rs As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo errHand
    
    mItemSerial.���� = -1
    mItemSerial.���� = -1
    mItemSerial.���� = -1
    mItemSerial.���� = -1

'    Erase mvarDataType
    
    lblComment.Caption = "˵��:"
    mbln�������� = False
    mstr��Сʱ�� = ""
    Call Get��Ժ���ʱ��
    
    If mblnNoneShow Then Exit Function
    
    '������ʼ��
    mshUpTab.Rows = 3
    mshUpTab.Cell(flexcpAlignment, 0, 0, mshUpTab.Rows - 1, 0) = 4
    mshUpTab.Cell(flexcpText, 0, mshUpTab.FixedCols, mshUpTab.Rows - 1, mshUpTab.Cols - 1) = ""
    mshUpTab.Cell(flexcpData, 0, mshUpTab.FixedCols, mshUpTab.Rows - 1, mshUpTab.Cols - 1) = ""
    mshUpTab.Cell(flexcpForeColor, 0, mshUpTab.FixedCols, 1, mshUpTab.Cols - 1) = 16711680
    mshUpTab.Cell(flexcpForeColor, 2, mshUpTab.FixedCols, 2, mshUpTab.Cols - 1) = 255

    mshDownTab.RowHidden(0) = True
    
    mshUpTab.Redraw = False
    mshScale.Redraw = False
    mshDownTab.Redraw = False
    mvarEdit = False
    If picLine.Count > 1 Then
        For i = 1 To picLine.Count - 1
            Unload picLine(i)
        Next
    End If
    
    '��ȡ���ش�ӡ��ʼҳ��
    UserControl.BackColor = RGB(255, 255, 255)
    Call ClearSpecRowCol(mshScale, 0, Array())
    
    'Ϊ������picture�ϻ�����������������Ƚ�picScale��picGraph����Ϊ���
    picScale.Width = Screen.Width
    picScale.Height = Screen.Height
    picGraph.Left = 0
    picGraph.Top = 0
    picGraph.Width = Screen.Width
    picGraph.Height = Screen.Height
    lblCur.Top = 350
    
    '������Ŀ���õ�����ʾ����
    
    mItemSerial.������ = -1
    mItemSerial.������ = -1
    mItemNo.��Һ = 0
    mItemNo.��� = 0
    mItemNo.���� = 0
    mItemNo.���� = 0
    mItemNo.���� = 0
    
    mstrSQL = " Select ��Ŀ��� From ���¼�¼��Ŀ A Where A.��¼��=[1]"
    Set rs = zlDatabase.OpenSQLRecord(mstrSQL, mstrMsgTitle, "����")
    If rs.BOF = False Then
        mItemNo.���� = rs("��Ŀ���").Value
    End If

    Set rs = zlDatabase.OpenSQLRecord(mstrSQL, mstrMsgTitle, "����")
    If rs.BOF = False Then
        mItemNo.���� = rs("��Ŀ���").Value
    End If
        
    Set rs = zlDatabase.OpenSQLRecord(mstrSQL, mstrMsgTitle, "����")
    If rs.BOF = False Then
        mItemNo.���� = rs("��Ŀ���").Value
    End If
        
    Set rs = zlDatabase.OpenSQLRecord(mstrSQL, mstrMsgTitle, "����")
    If rs.BOF = False Then
        mItemNo.���� = rs("��Ŀ���").Value
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    '��������
    mstrSQL = " Select Max((A.���ֵ-A.��Сֵ)/Decode(A.��λֵ,0,1,A.��λֵ)+A.�����) From ���¼�¼��Ŀ A Where A.��¼��=1 "
    Set rs = zlDatabase.OpenSQLRecord(mstrSQL, mstrMsgTitle)
    If rs.BOF = False Then
        If IsNull(rs.Fields(0).Value) = False Then
            mshDownTab.Tag = "1"
            mvarEdit = True
            
            mshScale.Rows = MAXROWS
            
            mint����Ӧ�� = 2
            mstr���ʷ��� = ""
            mstrSQL = "Select a.Ӧ�÷�ʽ,b.��¼�� From �����¼��Ŀ a,���¼�¼��Ŀ b Where a.��Ŀ���=-1 And a.��Ŀ���=b.��Ŀ���"
            Set rs = zlDatabase.OpenSQLRecord(mstrSQL, mstrMsgTitle)
            If rs.BOF = False Then
                mint����Ӧ�� = zlCommFun.NVL(rs("Ӧ�÷�ʽ").Value, 2)
                mstr���ʷ��� = zlCommFun.NVL(rs("��¼��").Value, "��")
            End If
            
            '�õ�����������Ŀ
                        
            mstrSQL = " Select A.��¼��,A.��¼�� as ��Ŀ��,A.��Ŀ��� as ��Ŀ��,Nvl(B.ID,0) as ��ĿID," & _
                        " C.��Ŀ��λ As ��λ,��¼��,��Сֵ,���ֵ,��¼ɫ,1 as ��¼��,��λֵ,�����,Nvl(B.����,1) as �洢���� " & _
                        " From ���¼�¼��Ŀ A,����������Ŀ B,�����¼��Ŀ C " & _
                        " Where c.��ĿID=B.ID(+) And A.��Ŀ���=C.��Ŀ��� And A.��¼��=1 And Nvl(C.Ӧ�÷�ʽ,0)=1 AND C.����ȼ�>=[1] And Nvl(C.���ò���,0) In (0,[3]) " & _
                        " And (c.���ÿ���=1 Or (c.���ÿ���=2 And Exists (Select 1 From �������ÿ��� D Where D.��Ŀ���=c.��Ŀ��� And D.����id=[2]))) " & _
                        " Order by A.�������"
                        
            Set rs = zlDatabase.OpenSQLRecord(mstrSQL, mstrMsgTitle, Val(mrsParam("����ȼ�").Value), Val(mrsParam("����id").Value), IIf(Val(mrsParam("Ӥ��").Value) = 0, 1, 2))
            If rs.RecordCount > 0 Then rs.MoveFirst
            
            'mshScale����=�̶��� + 7�� * 6��
            
            mshScale.Cols = rs.RecordCount + (mshUpTab.Cols - 1) * 6
            mshScale.FixedCols = rs.RecordCount
            mshScale.RowHeightMin = ROWHEIGHT
            
            mshScale.Tag = ""
            Do While Not rs.EOF
                
                If rs!��Ŀ�� = "����" Then mbln�������� = True
                
                If zlCommFun.NVL(rs!��Ŀ��, 0) = 7 Then mItemSerial.������ = rs.AbsolutePosition - 1
                If zlCommFun.NVL(rs!��Ŀ��, 0) = 6 Then mItemSerial.������ = rs.AbsolutePosition - 1
                
                If rs.AbsolutePosition > picLine.Count Then
                    Load picLine(rs.AbsolutePosition - 1)
                End If
                picLine(rs.AbsolutePosition - 1).Tag = rs!���ֵ & ";" & rs!��Сֵ & ";" & rs!��λֵ & ";" & rs!�����
                picLine(rs.AbsolutePosition - 1).Visible = True
                picLine(rs.AbsolutePosition - 1).ZOrder
                
                If rs!��Ŀ�� = "����" Then mItemSerial.���� = rs.AbsolutePosition - 1
                If rs!��Ŀ�� = "����" Then mItemSerial.���� = rs.AbsolutePosition - 1
                If rs!��Ŀ�� = "����" Then mItemSerial.���� = rs.AbsolutePosition - 1
                If rs!��Ŀ�� = "����" Then mItemSerial.���� = rs.AbsolutePosition - 1
                
                '���ñ������Ŀ
                mshScale.ColWidth(rs.AbsolutePosition - 1) = IIf(mshScale.FixedCols < 4, 1200 / mshScale.FixedCols, 450)
                mshScale.ColData(rs.AbsolutePosition - 1) = Val(rs("��Ŀ��").Value)
                If zlCommFun.NVL(rs("��λ").Value) <> "" Then
                    mshScale.TextMatrix(0, rs.AbsolutePosition - 1) = rs("��Ŀ��").Value & " (" & zlCommFun.NVL(rs("��λ").Value) & ")"
                Else
                    mshScale.TextMatrix(0, rs.AbsolutePosition - 1) = rs("��Ŀ��").Value
                End If
                mshScale.Cell(flexcpAlignment, 0, rs.AbsolutePosition - 1) = flexAlignCenterTop

                mshScale.Row = 0
                mshScale.Col = rs.AbsolutePosition - 1
                mshScale.CellForeColor = rs!��¼ɫ
                If mItemSerial.���� = rs.AbsolutePosition - 1 Then
                    mshScale.Tag = mshScale.Tag & " "
                Else
                    mshScale.Tag = mshScale.Tag & zlCommFun.NVL(rs!��¼��, " ")
                End If
                
                mstrChar(0) = ""
                mstrChar(0) = ""
                mstrChar(0) = ""
                If mItemSerial.���� = rs.AbsolutePosition - 1 Then

                    Dim varTmp As Variant
                                        
                    gstrSQL = "Select ��¼�� From ���¼�¼��Ŀ Where ��Ŀ���=[1]"
                    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, mstrMsgTitle, 1)
                    If rsTmp.BOF = False Then
                        varTmp = Split(zlCommFun.NVL(rsTmp("��¼��").Value, "��,��,��"), ",")
                    Else
                        varTmp = Split("��,��,��", ",")
                    End If
                    mstrChar(0) = CStr(varTmp(0))
                    mstrChar(1) = CStr(varTmp(1))
                    mstrChar(2) = CStr(varTmp(2))
        
                    lblComment.Caption = lblComment.Caption & IIf(rs.AbsolutePosition = 1, "", "��") & rs!��Ŀ�� & "(����" & mstrChar(0) & ",Ҹ��" & mstrChar(1) & ",����" & mstrChar(2) & ")"
                ElseIf mItemSerial.���� = rs.AbsolutePosition - 1 Then
                    mstrBreath = rs!��¼��
                    lblComment.Caption = lblComment.Caption & IIf(rs.AbsolutePosition = 1, "", "��") & rs!��Ŀ�� & "(��������" & rs!��¼�� & ",������R)"
                ElseIf mItemSerial.���� = rs.AbsolutePosition - 1 Then
                    mstrPulse = rs!��¼��
                    lblComment.Caption = lblComment.Caption & IIf(rs.AbsolutePosition = 1, "", "��") & rs!��Ŀ�� & "(ȱʡ��¼��" & rs!��¼�� & ",����H)"
                Else
                    lblComment.Caption = lblComment.Caption & IIf(rs.AbsolutePosition = 1, "", "��") & rs!��Ŀ�� & "(" & rs!��¼�� & ")"
                End If
                
                For intRow = 1 To mshScale.Rows - 1
                    mshScale.Row = intRow
                    mshScale.CellForeColor = rs!��¼ɫ
                    mshScale.ROWHEIGHT(intRow) = ROWHEIGHT * 5

                    If intRow >= rs!����� And rs!���ֵ - (intRow - rs!�����) * rs!��λֵ >= rs!��Сֵ Then
                    
                        '�պ�Ϊ����ʱ�����һ�̶�ֵ
                        
                        If Int(rs("���ֵ").Value - (intRow - rs!�����) * rs("��λֵ").Value) = rs!���ֵ - (intRow - rs!�����) * rs!��λֵ Then
                        
                            Select Case rs!��Ŀ��
                            Case "����", "����", "����"
                                If (rs!���ֵ - (intRow - rs!�����) * rs!��λֵ) Mod 10 = 0 Then
                                    mshScale.TextMatrix(intRow, rs.AbsolutePosition - 1) = rs!���ֵ - (intRow - rs!�����) * rs!��λֵ
                                End If
                            Case "����"
                                mshScale.TextMatrix(intRow, rs.AbsolutePosition - 1) = CStr(rs!���ֵ - (intRow - rs!�����) * rs!��λֵ) & "��"
                            'Case "����"
                            '    mshScale.TextMatrix(intRow, rs.AbsolutePosition - 1) = rs!���ֵ - (intRow - rs!�����) * rs!��λֵ
                            End Select
                            
                        End If
                    End If
                Next
                rs.MoveNext
            Loop
        End If
    End If
    
    mshScale.Cell(flexcpAlignment, 0, 0, mshScale.Rows - 1, mshScale.FixedCols - 1) = 4
    
    With vsf
        .Cols = 0
        .NewColumn "", 0, 1
        .NewColumn "��Ŀ", mshScale.FixedCols * mshScale.ColWidth(0) + 15, 1
        For intCol = 1 To 42
            .NewColumn intCol, HOUR_STEP_Twips, 1, , 1
        Next
        .FixedCols = 2
        .Cell(flexcpAlignment, 1, 1) = flexAlignCenterCenter
        .Cell(flexcpFontName, 1, 2, 1, .Cols - 1) = "Times New Roman"
        .Cell(flexcpFontSize, 1, 2, 1, .Cols - 1) = 7.5
        .Body.Select 1, 1
        .Body.CellBorder 0, 1, 0, 0, 0, 0, 0
        .Body.Select 1, vsf.Cols - 1
        .Body.CellBorder 0, 0, 0, 1, 0, 0, 0
        .Body.BackColorFixed = .Body.BackColor
        
        For intCol = 3 To .Cols - 1 Step 2
            .Cell(flexcpBackColor, 1, intCol, 1, intCol) = &HF7ECE6
        Next
        
    End With
    
    '��ʼ��ֱ��¼����Ŀ,mshDownTab��ǰ���������ڼ�¼��Ŀ�й�����:��ʾ����;��Ŀ����;���ֵ;��Сֵ
    '------------------------------------------------------------------------------------------------------------------
    Set rsTmp = GetGridItem(Val(mrsParam("����ȼ�").Value), Val(mrsParam("����id").Value), IIf(Val(mrsParam("Ӥ��").Value) = 0, 1, 2), 1)
    With rsTmp
        If rsTmp.RecordCount > 0 Then
            rsTmp.Sort = "�������"
            rsTmp.MoveFirst
            
            mshDownTab.Tag = "1"

            '��ʼ����
            ReDim mItemStru(0 To 1)
            mshDownTab.Rows = 1
            
            For i = 1 To rsTmp.RecordCount
                
                Call AppendGridItem(rsTmp)

                rsTmp.MoveNext
            Next
        Else
            mshDownTab.Rows = 2
        End If
    End With
   
   '------------------------------------------------------------------------------------------------------------------
    '�ҵ�ϵͳ������Ŀ����Ŀ��
    '��������ϵͳ�̶�����Ŀ
    mshUpTab.RowData(0) = -1
    mshUpTab.RowData(1) = -2
    mshUpTab.RowData(2) = -3
    
    '��ʼ������
    With mshUpTab
        .ColWidth(0) = mshScale.FixedCols * mshScale.ColWidth(0)
        .TextMatrix(0, 0) = "��    ��"
        .TextMatrix(1, 0) = "סԺ����"
        .TextMatrix(2, 0) = "��/�������"
        For intCol = 1 To .Cols - 1
            .ColWidth(intCol) = HOUR_STEP_Twips * 6
        Next
    End With
    
    '��ʼ���ݱ����
    With mshScale
        .ROWHEIGHT(0) = 400 '600
        For intCol = .FixedCols To .Cols - 1
            .ColWidth(intCol) = HOUR_STEP_Twips
        Next
    End With
    
    '��ʼ����ı��
    With mshDownTab
        .ColWidth(0) = mshScale.FixedCols * mshScale.ColWidth(0)
        .ColWidth(1) = 0
        .ColWidth(2) = 0
        .ColWidth(3) = 0
        .ColAlignment(0) = 4
        .ColAlignment(1) = 4
        .ColAlignment(2) = 4
        .ColAlignment(3) = 4
        
        For intCol = .FixedCols To .Cols - 1
            .ColWidth(intCol) = HOUR_STEP_Twips * 3
            .ColAlignment(intCol) = 1
        Next
        
        
    End With
    
    '�ٸ��ݶ��������ݻ�ͼ
    mshUpTab.Redraw = True
    mshScale.Redraw = True
    mshDownTab.Redraw = True
    
    Call ReadPatiInfo
    Call SetVisible
    
    FaceInit = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function AppendGridItem(ByVal rsTmp As ADODB.Recordset, Optional ByVal blnAppend As Boolean) As Boolean
    '******************************************************************************************************************
    '���ܣ���д�����Ŀ�ı����
    '������rsTmp��Ҫ��ӵı����Ŀ
    '���أ�
    '******************************************************************************************************************
    Dim intTmp As Integer
    Dim intRow As Integer
    
    On Error GoTo errHand
            
    Select Case rsTmp("����").Value
    '------------------------------------------------------------------------------------------------------------------
    Case "����"
        
        vsf.TextMatrix(1, 1) = "����" & IIf(Not IsNull(rsTmp!��λ), "(" & rsTmp!��λ & ")", "")
        
        mItemOtherStru(0).��Ŀ���� = zlCommFun.NVL(rsTmp!����, "")
        mItemOtherStru(0).��Ŀ��� = zlCommFun.NVL(rsTmp!��Ŀ��, 0)
        mItemOtherStru(0).�������� = zlCommFun.NVL(rsTmp!�洢����, 1)
        mItemOtherStru(0).���ݳ��� = zlCommFun.NVL(rsTmp!��Ŀ����, 0)
        mItemOtherStru(0).С��λ�� = zlCommFun.NVL(rsTmp!��ĿС��, 0)
        mItemOtherStru(0).��Сֵ = zlCommFun.NVL(rsTmp!��Сֵ, "")
        mItemOtherStru(0).���ֵ = zlCommFun.NVL(rsTmp!���ֵ, "")
        mItemOtherStru(0).��¼Ƶ�� = zlCommFun.NVL(rsTmp!��¼Ƶ��, 0)
        mItemOtherStru(0).���Ŀ = False
        
        mItemNo.���� = zlCommFun.NVL(rsTmp!��Ŀ��, 0)
    '------------------------------------------------------------------------------------------------------------------
    Case "����ѹ"
        
        mItemNo.����ѹ = zlCommFun.NVL(rsTmp!��Ŀ��, 0)
        
        mItemOtherStru(1).��Ŀ���� = zlCommFun.NVL(rsTmp!����, "")
        mItemOtherStru(1).��Ŀ��� = zlCommFun.NVL(rsTmp!��Ŀ��, 0)
        mItemOtherStru(1).�������� = zlCommFun.NVL(rsTmp!�洢����, 1)
        mItemOtherStru(1).���ݳ��� = zlCommFun.NVL(rsTmp!��Ŀ����, 0)
        mItemOtherStru(1).С��λ�� = zlCommFun.NVL(rsTmp!��ĿС��, 0)
        mItemOtherStru(1).��Сֵ = zlCommFun.NVL(rsTmp!��Сֵ, "")
        mItemOtherStru(1).���ֵ = zlCommFun.NVL(rsTmp!���ֵ, "")
        mItemOtherStru(1).��¼Ƶ�� = zlCommFun.NVL(rsTmp!��¼Ƶ��, 0)
        mItemOtherStru(1).���Ŀ = False
                
    '------------------------------------------------------------------------------------------------------------------
    Case Else
                
        mshDownTab.Rows = mshDownTab.Rows + 1
        intRow = mshDownTab.Rows - 1
        
        If rsTmp("����").Value = "����ѹ" Then
            mItemNo.Ѫѹ = zlCommFun.NVL(rsTmp!��Ŀ��, 0)
            mItemSerial.Ѫѹ = intRow
            mshDownTab.TextMatrix(intRow, 0) = "Ѫѹ" & IIf(Not IsNull(rsTmp!��λ), "(" & rsTmp!��λ & ")", "")
        Else
            mshDownTab.TextMatrix(intRow, 0) = zlCommFun.NVL(rsTmp!����, "") & IIf(Not IsNull(rsTmp!��λ), "(" & rsTmp!��λ & ")", "")
        End If
        
        mshDownTab.RowData(intRow) = rsTmp("��Ŀ��").Value
        mshDownTab.ROWHEIGHT(intRow) = 255
        mshDownTab.TextMatrix(intRow, 1) = mshDownTab.TextMatrix(intRow, 0)
        mshDownTab.TextMatrix(intRow, 2) = zlCommFun.NVL(rsTmp!���ֵ, "")
        mshDownTab.TextMatrix(intRow, 3) = zlCommFun.NVL(rsTmp!��Сֵ, "")
                
        ReDim Preserve mItemStru(intRow)
        
        mItemStru(intRow).��Ŀ���� = zlCommFun.NVL(rsTmp!����, "")
        mItemStru(intRow).��Ŀ��� = zlCommFun.NVL(rsTmp!��Ŀ��, 0)
        mItemStru(intRow).�������� = zlCommFun.NVL(rsTmp!�洢����, 1)
        mItemStru(intRow).���ݳ��� = zlCommFun.NVL(rsTmp!��Ŀ����, 0)
        mItemStru(intRow).С��λ�� = zlCommFun.NVL(rsTmp!��ĿС��, 0)
        mItemStru(intRow).��Сֵ = zlCommFun.NVL(rsTmp!��Сֵ, "")
        mItemStru(intRow).���ֵ = zlCommFun.NVL(rsTmp!���ֵ, "")
        mItemStru(intRow).��¼Ƶ�� = zlCommFun.NVL(rsTmp!��¼Ƶ��, 0)
        mItemStru(intRow).���Ŀ = (zlCommFun.NVL(rsTmp!��Ŀ����, 1) = 2)
        
        Select Case zlCommFun.NVL(rsTmp!��Ŀ��, 0)
        Case 6
            mItemSerial.������ = intRow
        Case 7
            mItemSerial.������ = intRow
        Case 9
            mItemNo.��Һ = 9
        Case 10
            mItemNo.��� = 10
        End Select
                    
    End Select
    
    '����Ǻ�����ӱ����Ŀ�������Ӵ����޸ı�־
    If blnAppend Then

        With mshDownTab
            For intCol = .FixedCols To .Cols - 1
                .TextMatrix(GridDataRow.�޸ı�־, intCol) = .TextMatrix(GridDataRow.�޸ı�־, intCol) & ";"
            Next
            
        End With
        
    End If
    
    AppendGridItem = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Function

Private Function ShowOpsDays() As Boolean
    '******************************************************************************************************************
    '��ʾ��ǰ������ڵ������ձ��
    '******************************************************************************************************************
    Dim lng���� As Long
    Dim intCol As Integer
    Dim intLoop As Integer
    Dim strTmp As String
    Dim rsTmp As New ADODB.Recordset
    Dim strFrom As String
    
    On Error GoTo errHand
    
    '�������ǰ�����ڵ��������ֱ��
    
        
    For intCol = 1 To mshUpTab.Cols - 1
        mshUpTab.TextMatrix(2, intCol) = mstrOpsSvr(intCol)
    Next
    
    strFrom = Split(picScale.Tag, ";")(0)
    
    '�ҿ�ʼ����-14��ǰ����������
    mstrSQL = "SELECT Nvl(Count(a.����ʱ��),0) As ���� " & _
                "FROM ���˻����¼ a,���˻������� c " & _
                "Where a.ID = c.��¼ID " & _
                    "AND a.������Դ=2 " & _
                    "AND Nvl(a.Ӥ��,0)=[4] " & _
                    "AND a.����id=[1] " & _
                    "AND a.��ҳid=[2] " & _
                    "AND c.��¼����=4 " & _
                    "AND a.����ʱ��<[3] And c.��ֹ�汾 Is Null "
    If mblnMoved Then
        mstrSQL = Replace(mstrSQL, "���˻����¼", "H���˻����¼")
        mstrSQL = Replace(mstrSQL, "���˻�������", "H���˻�������")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, mstrMsgTitle, Val(mrsParam("����id")), Val(mrsParam("��ҳid")), CDate(strFrom), Val(mrsParam("Ӥ��")))
    If rsTmp.BOF = False Then lng���� = rsTmp("����").Value
    
    For intCol = 1 To mshUpTab.Cols - 1
    
        If DateDiff("d", CDate(strFrom), CDate(Split(picScale.Tag, ";")(1))) + 1 >= intCol Then

            If mshUpTab.ColData(intCol) = 1 Or mshUpTab.ColData(intCol) = 2 And lng���� < 12 Then
                
                lng���� = lng���� + 1
                
                strTmp = Switch(lng���� = 1, "��", _
                                lng���� = 2, "��", _
                                lng���� = 3, "��", _
                                lng���� = 4, "��", _
                                lng���� = 5, "��", _
                                lng���� = 6, "��", _
                                lng���� = 7, "��", _
                                lng���� = 8, "��", _
                                lng���� = 9, "��", _
                                lng���� = 10, "��", _
                                lng���� = 11, "��", _
                                lng���� = 12, "��")
                
                If mblnStopFlag Then
                    
                    If strTmp = "��" Then
                        mshUpTab.TextMatrix(2, intCol) = "0"
                    Else
                        mshUpTab.TextMatrix(2, intCol) = strTmp & "- 0"
                    End If
                Else
                    
                    If mshUpTab.TextMatrix(2, intCol) <> "" Then
                        mshUpTab.TextMatrix(2, intCol) = mshUpTab.TextMatrix(2, intCol) & "/" & strTmp
                    Else
                        mshUpTab.TextMatrix(2, intCol) = strTmp
                    End If
                End If
                
                For intLoop = intCol + 1 To mshUpTab.Cols - 1
                    strTmp = intLoop - intCol
                    
                    If Val(strTmp) <= mintOpDays Then
                        
                        If mblnStopFlag Then
                            mshUpTab.TextMatrix(2, intLoop) = strTmp
                        Else
                            If mshUpTab.TextMatrix(2, intLoop) <> "" Then
                                mshUpTab.TextMatrix(2, intLoop) = mshUpTab.TextMatrix(2, intLoop) & "/" & strTmp
                            Else
                                mshUpTab.TextMatrix(2, intLoop) = strTmp
                            End If
                        End If
                    End If
                    
                Next
                            
            End If
        End If
    Next
    
    ShowOpsDays = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ReadBodyData() As Boolean
    '******************************************************************************************************************
    '���ܣ� ���ݵ�ǰ�������±���ȡ���˵����±����ݣ���д����Ӧ�ĵ�Ԫ��
    '������ lng����id : ����
    '       strFrom : ��ʼʱ��
    '       strTo : ��ֹʱ��
    '���أ�
    'ע�⣺ ���˲�������������Ϊ0ʱ���������ݡ�����������ʱ�����ݲ��ҡ�������ID��Ϊ�գ�����˵������������¼�˵�ϵͳ��
    '       Ŀ��ֵ����������Ŀһ�������ڡ����˲����������������������п��Դӡ�������ID���ҵ����Ǹ���Ŀ
    '******************************************************************************************************************
    Dim lngValue As Long
    Dim lngValue2 As Long
    Dim intSvrCol As Integer
    Dim lng��λ���� As Long
    Dim dbl��� As Double
    Dim dbl��С As Double
    Dim dbl��λֵ As Double
    Dim lng����� As Long
    Dim rsTmp As New ADODB.Recordset
    Dim aryValue() As String
    Dim aryPart() As String
    Dim intMinCol As Integer
    Dim intMaxCol As Integer
    Dim blnOperate As Boolean
    Dim dtOperate As Date
    Dim strEnd As String '�Ƿ���������������
    Dim i As Long
    Dim lng����id As Long, strFrom As String, strTo As String
    Dim intColTmp As Integer
    Dim lngColor As Long
    Dim strTime As String
    Dim strStart1 As String
    Dim strEnd1 As String
    Dim strTmp As String
    Dim blnShow As Boolean          '�Ƿ���ʾ���Ժ����Ϣ
    
    Screen.MousePointer = vbHourglass
    
    On Error GoTo ErrHead
    
    '������ʼ��
    '------------------------------------------------------------------------------------------------------------------
    lng����id = Val(mrsParam("����id"))
    strFrom = CStr(mrsParam("��ʼʱ��"))
    
    If zlDatabase.GetPara("���µ���ʾ���", glngSys, 1255, 1) = 0 Then
        lbl(7).Visible = False
        txtCard(7).Visible = False
    Else
        lbl(7).Visible = True
        txtCard(7).Visible = True
    End If
    
    If CStr(mrsParam("����ʱ��")) > Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") Then
        strTo = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    Else
        strTo = CStr(mrsParam("����ʱ��"))
    End If
        
    txtCard(3).Text = ""
    
    '������������������¼���ʱ�䣬��Ӥ�����µ��Ŀ�ʼʱ��
    If Val(mrsParam("Ӥ��").Value) > 0 Then
        mstrSQL = " Select  b.����ʱ�� From ������������¼ B Where ����id=[1] And ��ҳid=[2] And ���=[3] "
        Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, mstrMsgTitle, Val(mrsParam("����id").Value), Val(mrsParam("��ҳid").Value), Val(mrsParam("Ӥ��").Value))
        If rsTmp.BOF = False Then
            mstrEnterDate = Format(zlCommFun.NVL(rsTmp("����ʱ��").Value), "yyyy-MM-dd HH:mm:ss")
            txtCard(3).Text = Format(zlCommFun.NVL(rsTmp("����ʱ��").Value), "yyyy-MM-dd")
            strFrom = mstrEnterDate
        End If
    End If
    
    '����4Сʱ��ȷ�ȣ�����ʼʱ�����ֹʱ��
    
    intCol = GetCurveColumn(CDate(strFrom), CDate(strFrom), mlngHourBegin) + mshScale.FixedCols - 1
    strFrom = Split(GetCurveDateTime(intCol - mshScale.FixedCols + 1, CDate(strFrom), mlngHourBegin), ",")(0)
    
    If Int(CDate(strFrom)) < Int(CDate(mrsParam("��ʼʱ��"))) Then
        strFrom = Format(Int(CDate(mrsParam("��ʼʱ��"))), "yyyy-MM-dd HH:mm:ss")
    End If

    intCol = GetCurveColumn(CDate(strTo), CDate(strFrom), mlngHourBegin) + mshScale.FixedCols - 1
    strTo = Split(GetCurveDateTime(intCol - mshScale.FixedCols + 1, CDate(strFrom), mlngHourBegin), ",")(1)
    
    '������Ŀ�ʼʱ�����ֹʱ����д�����tag��
    picScale.Tag = strFrom & ";" & strTo
    
    
    '��ȡ���˻�����Ϣ
    '------------------------------------------------------------------------------------------------------------------
    '��д���˲��������ŵȣ����ڲ��˿����ڿ��ڻ����ȣ�������дΪ���˵���ָ��ʱ�����󴲺�
'    lblTime.Caption = "����:" & Format(strFrom, "MM-DD") & "��" & Format(strTo, "MM-DD")
    
    '��Ժʱ��(�����ʱ��Ϊ׼)
    mstrSQL = "select ��ʼʱ�� from ���˱䶯��¼ where ����id=[1] And ��ҳid=[2] and ��ʼԭ��=2 order by ��ʼʱ��"
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, mstrMsgTitle, Val(mrsParam("����id")), Val(mrsParam("��ҳid")))
    If rsTmp.BOF = False Then
        If txtCard(3).Text = "" Then txtCard(3).Text = Format(zlCommFun.NVL(rsTmp("��ʼʱ��").Value), "yyyy-MM-dd")
    End If
    
    '��ȡ���˻�����Ϣ
    mstrSQL = " Select  b.����,A.סԺ��,b.��Ժʱ��,b.�Ա�,b.���� From ������Ϣ B,������ҳ A Where A.����ID=B.����ID And A.����id=[1] And A.��ҳID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, mstrMsgTitle, Val(mrsParam("����id")), Val(mrsParam("��ҳid")))
    If rsTmp.BOF = False Then
        txtCard(0).Text = zlCommFun.NVL(rsTmp("����").Value)
        txtCard(0).Tag = zlCommFun.NVL(rsTmp("����").Value)
        txtCard(1).Text = zlCommFun.NVL(rsTmp("סԺ��").Value)
        txtCard(5).Text = zlCommFun.NVL(rsTmp("�Ա�").Value)
        txtCard(6).Text = zlCommFun.NVL(rsTmp("����").Value)
        If txtCard(3).Text = "" Then txtCard(3).Text = Format(zlCommFun.NVL(rsTmp("��Ժʱ��").Value), "yyyy-MM-dd")
    End If
    
    Call zlMenuClick("��ʾ��������")

    '��ȡ���˿��ҡ����ŵ���Ϣ
    
    txtCard(2).Text = ""
    txtCard(4).Text = ""
    
    mstrSQL = " Select  c.���� As ����,b.���� As ����,a.����,a.��ʼԭ�� " & _
                "From ���˱䶯��¼ a,���ű� b,���ű� c " & _
                "Where a.����id=[1] And a.��ҳid=[2] And a.����id Is Not Null And a.����id=b.id and a.����id=c.id And a.��ʼʱ��-4/24<=[3] And Nvl(a.��ֹʱ��,Sysdate)>=[4] Order By a.��ʼʱ��"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, mstrMsgTitle, Val(mrsParam("����id")), Val(mrsParam("��ҳid")), CDate(strTo), CDate(strFrom))
    If rsTmp.BOF = False Then
        Do While Not rsTmp.EOF
            
            If zlCommFun.NVL(rsTmp("����").Value) <> strTmp And zlCommFun.NVL(rsTmp("����").Value) <> "" Then
            
                strTmp = zlCommFun.NVL(rsTmp("����").Value)
                
                If txtCard(2).Text = "" Then
                    txtCard(2).Text = strTmp
                Else
                    txtCard(2).Text = txtCard(2).Text & "->" & strTmp
                End If
                
            End If

            If zlCommFun.NVL(rsTmp("����").Value) <> strTime And zlCommFun.NVL(rsTmp("����").Value) <> "" Then
            
                strTime = zlCommFun.NVL(rsTmp("����").Value)
                
                If txtCard(4).Text = "" Then
                    txtCard(4).Text = strTime
                Else
                    txtCard(4).Text = txtCard(4).Text & "->" & strTime
                End If
                
            End If
                        
            rsTmp.MoveNext
        Loop
        
        If Left(txtCard(2).Text, 2) = "->" Then txtCard(2).Text = Mid(txtCard(2).Text, 3)
        If Left(txtCard(4).Text, 2) = "->" Then txtCard(4).Text = Mid(txtCard(4).Text, 3)
    End If
    
    mshUpTab.Redraw = False
    mshScale.Redraw = False
        
    '��д���ں�סԺ��������Ϣ
    '------------------------------------------------------------------------------------------------------------------
    With mshUpTab
        
        intSvrCol = .Col
        
        lngValue = 0
        mstrSQL = "Select zl_CalcInDays([1],[2],[3],[4]) As ��ʼ���� From Dual"
        Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, mstrMsgTitle, Val(mrsParam("����id").Value), Val(mrsParam("��ҳid").Value), Val(mrsParam("Ӥ��").Value), (Int(CDate(strFrom))))
        If rsTmp.BOF = False Then
            lngValue = rsTmp("��ʼ����").Value
        End If
        
        For intCol = 1 To .Cols - 1

            .ColData(intCol) = 0
            .ColAlignment(intCol) = 4
            
            strTmp = Format(CDate(strFrom) + intCol - 1, "yyyy-MM-dd")
                        
            If Right(strTmp, 5) = "01-01" Then
                'һ��ĵ�һ��
                .TextMatrix(0, intCol) = strTmp
            ElseIf strTmp = Format(mstrEnterDate, "yyyy-MM-dd") Then
                '��Ժ��һ�죬д�����
                .TextMatrix(0, intCol) = strTmp
            ElseIf intCol = 1 Then
                .TextMatrix(0, intCol) = strTmp
            ElseIf Right(strTmp, 2) = "01" Then
                .TextMatrix(0, intCol) = Right(strTmp, 5)
            Else
                .TextMatrix(0, intCol) = Right(strTmp, 2)
            End If

            .TextMatrix(1, intCol) = lngValue + (intCol - 1)
            .TextMatrix(2, intCol) = ""
        Next

    End With

    
    Dim intDays As Integer
    
    For intCol = 1 To mshUpTab.Cols - 1
         mstrOpsSvr(intCol) = ""
         mstrOpsDays(intCol) = ""
    Next
        
    '1.��ȡ��������ͼ�����ݣ�����Ϊ����ֵ��д�����У�Ϊͼ����ʾ��׼��
    '------------------------------------------------------------------------------------------------------------------
    With mshScale
        For intCol = .FixedCols To .Cols - 1
            .TextMatrix(GraphDataRow.���ı�־, intCol) = String(.FixedCols, ";")    '���ı�־
            .TextMatrix(GraphDataRow.��������, intCol) = String(.FixedCols, ";")    '��������
            .TextMatrix(GraphDataRow.�ϱ�˵��, intCol) = ""                         '˵��(�ϱ�)
            
            .Cell(flexcpData, GraphDataRow.������־, intCol, GraphDataRow.������־, intCol) = ""
            .Cell(flexcpData, GraphDataRow.��λ��־, intCol, GraphDataRow.��λ��־, intCol) = ""
            .Cell(flexcpData, GraphDataRow.��Ժ��־, intCol, GraphDataRow.��Ժ��־, intCol) = ""
            .Cell(flexcpData, GraphDataRow.ת�Ʊ�־, intCol, GraphDataRow.ת�Ʊ�־, intCol) = ""
            .Cell(flexcpData, GraphDataRow.������־, intCol, GraphDataRow.������־, intCol) = ""
            .Cell(flexcpData, GraphDataRow.��Ժ��־, intCol, GraphDataRow.��Ժ��־, intCol) = ""
            .Cell(flexcpData, GraphDataRow.��Ʊ�־, intCol, GraphDataRow.��Ʊ�־, intCol) = ""
            
            .TextMatrix(GraphDataRow.������־, intCol) = ""                         '����
            .TextMatrix(GraphDataRow.��Ժ��־, intCol) = ""                         '��Ժ
            .TextMatrix(GraphDataRow.ת�Ʊ�־, intCol) = ""                         'ת��
            .TextMatrix(GraphDataRow.������־, intCol) = ""                         '����
            .TextMatrix(GraphDataRow.��Ժ��־, intCol) = ""                         '��Ժ
            .TextMatrix(GraphDataRow.��Ʊ�־, intCol) = ""                         '���
            .TextMatrix(GraphDataRow.���Ա�־, intCol) = ""                         '���¸��Ժϸ�
            .TextMatrix(GraphDataRow.�±�˵��, intCol) = ""                         '˵��(�±�)
            .TextMatrix(GraphDataRow.�Ͽ���־, intCol) = ""                         '���������ݣ��Ͽ�
            .TextMatrix(GraphDataRow.������־, intCol) = ""                         '����
            .TextMatrix(GraphDataRow.����ʱ��, intCol) = String(.FixedCols, ";")    '����ʱ��
            .TextMatrix(GraphDataRow.δ��˵��, intCol) = String(.FixedCols, ";")    'δ��˵��
            .TextMatrix(GraphDataRow.��λ��־, intCol) = String(.FixedCols, ";")    '���²�λ
            
        Next
        
    End With
    
    mintOpDays = Val(zlDatabase.GetPara("�������ע����", glngSys, 1255, "10"))
    mblnStopFlag = (Val(zlDatabase.GetPara("�ٴ�����ֹͣǰ�α�ע", glngSys, 1255, "0")) = 1)
    
    '��ʾ��ǰ�����ǰ�������ձ��
    '------------------------------------------------------------------------------------------------------------------
    mstrSQL = "SELECT a.����ʱ�� As ʱ��,c.��Ŀ���� " & _
                "FROM ���˻����¼ a,���˻������� c " & _
                "Where a.ID = c.��¼ID " & _
                    "AND a.������Դ=2 " & _
                    "AND Nvl(a.Ӥ��,0)=[5] " & _
                    "AND a.����id=[1] " & _
                    "AND a.��ҳid=[2] " & _
                    "AND c.��¼����=4 And c.��ֹ�汾 Is Null " & _
                    "AND a.����ʱ�� Between [3] And [4] Order By a.����ʱ�� "

    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, mstrMsgTitle, Val(mrsParam("����id")), Val(mrsParam("��ҳid")), CDate(strFrom) - 14, CDate(strTo), Val(mrsParam("Ӥ��")))
    If rsTmp.BOF = False Then
        Do While Not rsTmp.EOF
            dtOperate = Int(rsTmp.Fields(0).Value)
            For intCol = 1 To mshUpTab.Cols - 1
                
                If DateDiff("d", CDate(strFrom), CDate(Split(picScale.Tag, ";")(1))) + 1 >= intCol Then
                    intDays = Val(Int(CDate(strFrom)) + intCol - 1 - dtOperate)
    
                    Select Case intDays
                    Case 0
                    
                        mshUpTab.ColData(intCol) = 1
                        mstrOpsDays(intCol) = rsTmp.Fields(0).Value
                        
                    Case 1 To mintOpDays
                    
                        If intDays >= intCol Then
                        
                            If mshUpTab.TextMatrix(2, intCol) <> "" And Not mblnStopFlag Then
                                mshUpTab.TextMatrix(2, intCol) = mshUpTab.TextMatrix(2, intCol) & "/" & intDays
                            Else
                                mshUpTab.TextMatrix(2, intCol) = intDays
                            End If
                            
                        End If
                    End Select
                    
                    mstrOpsSvr(intCol) = mshUpTab.TextMatrix(2, intCol)
                    
                    intCount = GetCurveColumn(rsTmp("ʱ��").Value, CDate(strFrom), mlngHourBegin) + mshScale.FixedCols - 1
                    
                    Select Case rsTmp("��Ŀ����").Value
                    Case "����"
                        If intCount > 0 And intCount < mshScale.Cols And mBodyFlag.���� > 0 Then
                            If mBodyFlag.���� = 2 Then
                                mshScale.TextMatrix(3, intCount) = rsTmp("��Ŀ����").Value & "--" & ConvertTimeToChinese(Format(rsTmp("ʱ��").Value, "HH:mm"))
                            Else
                                mshScale.TextMatrix(3, intCount) = rsTmp("��Ŀ����").Value
                            End If
                            
                             mshScale.Cell(flexcpData, 3, intCount, 3, intCount) = Format(rsTmp("ʱ��").Value, "HH:mm:ss")
    
                        End If
                    Case Else
                        If intCount > 0 And intCount < mshScale.Cols And mBodyFlag.���� > 0 Then
                            If mBodyFlag.���� = 2 Then
                                mshScale.TextMatrix(3, intCount) = rsTmp("��Ŀ����").Value & "--" & ConvertTimeToChinese(Format(rsTmp("ʱ��").Value, "HH:mm"))
                            Else
                                mshScale.TextMatrix(3, intCount) = rsTmp("��Ŀ����").Value
                            End If
                            
                             mshScale.Cell(flexcpData, 3, intCount, 3, intCount) = Format(rsTmp("ʱ��").Value, "HH:mm:ss")
    
                        End If
                    End Select
                End If
            Next
            rsTmp.MoveNext
        Loop
    End If

    '��ʾ��ǰ������ڵ������ձ��
    '------------------------------------------------------------------------------------------------------------------
    Call ShowOpsDays
    
    '2.��ȡ���ת�ȱ�־����
    '------------------------------------------------------------------------------------------------------------------
    Dim bytShow As Byte
    
    Set rsTmp = GetDataFromHis(Val(mrsParam("����id")), Val(mrsParam("��ҳid")), Val(mrsParam("Ӥ��")), CDate(strFrom), CDate(strTo), 2)
    If Not (rsTmp Is Nothing) Then
        If rsTmp.BOF = False Then
            Do While Not rsTmp.EOF

                intCol = GetCurveColumn(rsTmp("ʱ��").Value, CDate(strFrom), mlngHourBegin) + mshScale.FixedCols - 1
                
                If zlCommFun.NVL(rsTmp("����")) <> "" Then
                    
                    bytShow = 0
                    
                    Select Case Val(rsTmp("�к�").Value)
                    Case 5
                        bytShow = mBodyFlag.��Ժ
                    Case 6
                        bytShow = mBodyFlag.ת��
                    Case 7
                        bytShow = mBodyFlag.����
                    Case 8
                        bytShow = mBodyFlag.��Ժ
                    Case 9
                        bytShow = mBodyFlag.���
                    End Select
                    
                    If intCol >= mshScale.FixedCols And intCol < mshScale.Cols And bytShow > 0 Then
                        blnShow = True
                        If Val(rsTmp("�к�").Value) = 8 And Val(mrsParam("Ӥ��")) > 0 Then
                            blnShow = mblnӤ�����µ���ʾ��Ժ
                        End If
                        
                        If blnShow Then
                            mshScale.Cell(flexcpData, Val(rsTmp("�к�").Value), intCol, Val(rsTmp("�к�").Value), intCol) = Format(rsTmp("ʱ��").Value, "HH:mm:ss")
                            Select Case bytShow
                            Case 1
                                mshScale.TextMatrix(Val(rsTmp("�к�").Value), intCol) = rsTmp("����").Value
                            Case 2
                                mshScale.TextMatrix(Val(rsTmp("�к�").Value), intCol) = rsTmp("����").Value & "--" & ConvertTimeToChinese(Format(rsTmp("ʱ��").Value, "HH:mm"))
                            Case 3
                                mshScale.TextMatrix(Val(rsTmp("�к�").Value), intCol) = rsTmp("����").Value & rsTmp("����").Value
                            Case 4
                                mshScale.TextMatrix(Val(rsTmp("�к�").Value), intCol) = rsTmp("����").Value & rsTmp("����").Value & "--" & ConvertTimeToChinese(Format(rsTmp("ʱ��").Value, "HH:mm"))
                            End Select
                        End If
                    End If
                End If
                                            
                rsTmp.MoveNext
            Loop
        End If
    End If
    
    If Val(mrsParam("Ӥ��")) > 0 Then
        Set rsTmp = GetDataFromHis(Val(mrsParam("����id")), Val(mrsParam("��ҳid")), Val(mrsParam("Ӥ��")), CDate(strFrom), CDate(strTo), 3)
        
        If Not (rsTmp Is Nothing) Then
            If rsTmp.BOF = False Then
                Do While Not rsTmp.EOF
    
                    intCol = GetCurveColumn(rsTmp("ʱ��").Value, CDate(strFrom), mlngHourBegin) + mshScale.FixedCols - 1
                    
                    If zlCommFun.NVL(rsTmp("����")) <> "" Then
                                               
                        If intCol >= mshScale.FixedCols And intCol < mshScale.Cols And mBodyFlag.���� > 0 Then
                            
                            mshScale.Cell(flexcpData, Val(rsTmp("�к�").Value), intCol, Val(rsTmp("�к�").Value), intCol) = Format(rsTmp("ʱ��").Value, "HH:mm:ss")
                            Select Case mBodyFlag.����
                            Case 1
                                mshScale.TextMatrix(Val(rsTmp("�к�").Value), intCol) = rsTmp("����").Value
                            Case 2
                                mshScale.TextMatrix(Val(rsTmp("�к�").Value), intCol) = rsTmp("����").Value & "--" & ConvertTimeToChinese(Format(rsTmp("ʱ��").Value, "HH:mm"))
                            Case 3
                                mshScale.TextMatrix(Val(rsTmp("�к�").Value), intCol) = rsTmp("����").Value & rsTmp("����").Value
                            Case 4
                                mshScale.TextMatrix(Val(rsTmp("�к�").Value), intCol) = rsTmp("����").Value & rsTmp("����").Value & "--" & ConvertTimeToChinese(Format(rsTmp("ʱ��").Value, "HH:mm"))
                            End Select
                            
                        End If
                    End If
                                                
                    rsTmp.MoveNext
                Loop
            End If
        End If
    End If
    
    '3.�ܲ�ȱ�ע����
    '------------------------------------------------------------------------------------------------------------------
    mstrSQL = "SELECT c.��¼����,a.����ʱ�� As ʱ��,c.��¼���� As ˵��,c.��¼��� " & _
                "FROM ���˻����¼ a,���˻������� C " & _
                "Where a.ID = c.��¼ID " & _
                    "AND Nvl(a.Ӥ��,0)=[5] " & _
                    "AND a.����id=[1] " & _
                    "AND a.��ҳid=[2] " & _
                    "AND C.��¼���� In (2,6) " & _
                    "AND a.������Դ=2 And c.��ֹ�汾 Is Null " & _
                    "AND a.����ʱ�� BETWEEN [3] And [4] " & _
                "Order By a.����ʱ��"
    If mblnMoved Then
        mstrSQL = Replace(mstrSQL, "���˻����¼", "H���˻����¼")
        mstrSQL = Replace(mstrSQL, "���˻�������", "H���˻�������")
    End If
    
    
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, mstrMsgTitle, _
                                        Val(mrsParam("����id")), _
                                        Val(mrsParam("��ҳid")), _
                                        CDate(Format(strFrom, "YYYY-MM-DD")), _
                                        CDate(Format(strTo, "YYYY-MM-DD") & " 23:59:59"), Val(mrsParam("Ӥ��")))
    With rsTmp
        Do While Not .EOF

            intCol = GetCurveColumn(!ʱ��, CDate(strFrom), mlngHourBegin) + mshScale.FixedCols - 1
            
            If intCol >= mshScale.FixedCols And intCol < mshScale.Cols Then
                If zlCommFun.NVL(rsTmp("��¼����").Value, 0) = 2 Then
                    mshScale.TextMatrix(GraphDataRow.�ϱ�˵��, intCol) = zlCommFun.NVL(!˵��)
                Else
                    mshScale.TextMatrix(GraphDataRow.�±�˵��, intCol) = zlCommFun.NVL(!˵��)
                End If
                
                aryValue() = Split(mshScale.TextMatrix(GraphDataRow.���ı�־, intCol), ";")
                aryValue(0) = 1
                mshScale.TextMatrix(GraphDataRow.���ı�־, intCol) = Join(aryValue, ";")
                mshScale.TextMatrix(GraphDataRow.�Ͽ���־, intCol) = IIf(IsNull(!��¼���), "0", !��¼���)
                
            End If
            
            .MoveNext
        Loop
    End With
    
    '4.�������ݲ���
    '------------------------------------------------------------------------------------------------------------------
    Dim int�к� As Integer
    Dim aryItemName() As String
    Dim aryItemOrder() As Integer
    
    Dim strItemOrder As String
    Dim strItemName As String
    
    ReDim aryItemName(0 To mshScale.FixedCols - 1)
    ReDim aryItemOrder(0 To mshScale.FixedCols - 1)
    
    For intCol = 0 To mshScale.FixedCols - 1
        If InStr(mshScale.TextMatrix(0, intCol), "(") > 0 Then
            strTmp = Trim(Left(mshScale.TextMatrix(0, intCol), InStr(mshScale.TextMatrix(0, intCol), "(") - 1))
        Else
            strTmp = Trim(mshScale.TextMatrix(0, intCol))
        End If
        
        aryItemName(intCol) = strTmp
        aryItemOrder(intCol) = intCol + 1

    Next
    
    '45987,������,2012-09-10,��������
    '1.������ʾΪ���ߣ�2������ͳһ�ú�ɫʵ�е㣨�񣩱�ʾ���������µ�����ʾ��������
    '3. ʹ�ú������Ļ��ߣ������Ժ�R��ʾ������Ӧʱ���ں���30�κ����¶����úڱʻ�R�����ڵ�R֮���Լ�R����������֮�䲻����
    mstrSQL = "SELECT a.����ʱ�� As ʱ��,Decode(D.��Ŀ���," & mItemNo.���� & ",Decode(C.���²�λ,'������','29',C.��¼����) ,c.��¼����) As ��ֵ,c.���²�λ,c.���Ժϸ�,D.��¼��,E.������Ŀ,D.��Ŀ���,C.��¼���,C.δ��˵�� " & _
                "FROM ���˻����¼ A,���˻������� C,���¼�¼��Ŀ D,�����¼��Ŀ E " & _
                "Where a.ID = c.��¼ID " & _
                    "AND a.������Դ=2 " & _
                    "AND Nvl(a.Ӥ��,0)=[5] " & _
                    "AND a.����id=[1] " & _
                    "AND a.��ҳid=[2] " & _
                    "AND D.��Ŀ���=c.��Ŀ��� " & _
                    "AND c.��¼����=1 " & _
                    "AND E.��Ŀ���=D.��Ŀ��� " & _
                    "AND E.����ȼ�>=[6]  " & _
                    "AND a.����ʱ�� BETWEEN [3] And [4] And c.��ֹ�汾 Is Null " & _
                    "AND D.��¼��=1  " & _
                "Order By a.����ʱ��,c.��¼���"
    If mblnMoved Then
        mstrSQL = Replace(mstrSQL, "���˻����¼", "H���˻����¼")
        mstrSQL = Replace(mstrSQL, "���˻�������", "H���˻�������")
    End If
    
    
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, mstrMsgTitle, Val(mrsParam("����id")), Val(mrsParam("��ҳid")), CDate(Format(strFrom, "YYYY-MM-DD")), CDate(Format(strTo, "YYYY-MM-DD") & " 23:59:59"), Val(mrsParam("Ӥ��")), Val(mrsParam("����ȼ�").Value))
    With rsTmp
        
        Dim dtTmp As Date
        Dim blnAllow As Boolean
        Dim rsOffset As ADODB.Recordset
        
        Call InitOffset(rsOffset)
                
        Do While Not .EOF

            intCol = GetCurveColumn(!ʱ��, CDate(strFrom), mlngHourBegin) + mshScale.FixedCols - 1
            
            If (intCol - mshScale.FixedCols) < 42 Then

                strTmp = !��¼��
                Select Case mint����Ӧ��
                Case 1      '����Ӧ��
                    If zlCommFun.NVL(!��¼���, 0) = 1 And strTmp = "����" Then
                        strTmp = "����"
                    End If
                Case 2      '����ʹ��
                    If strTmp = "����" Then strTmp = "����"
                End Select
                    
                '��������������к�
                For int�к� = 0 To mshScale.FixedCols - 1
                    If strTmp = aryItemName(int�к�) Then
                        int�к� = aryItemOrder(int�к�)
                        Exit For
                    End If
                Next
                                
                '���ͬһ�����ж��ֵ����ȡ��е��ֵ��Ϊ���е�ֵ
                dtTmp = CDate(Split(GetCurveDateTime(intCol - mshScale.FixedCols + 1, CDate(strFrom), mlngHourBegin), ",")(0))
                dtTmp = DateAdd("h", 2, dtTmp)
                blnAllow = IsCenterValue(rsOffset, int�к�, intCol, !ʱ��, dtTmp)
                
                If blnAllow Then
                    
                    '��ȡ��Ŀ����:���ֵ����Сֵ����λֵ�������
                    aryValue = Split(picLine(int�к� - 1).Tag, ";")
                    dbl��� = Val(Split(picLine(int�к� - 1).Tag, ";")(0))
                    dbl��С = Val(Split(picLine(int�к� - 1).Tag, ";")(1))
                    dbl��λֵ = Val(Split(picLine(int�к� - 1).Tag, ";")(2))
                    lng����� = Val(Split(picLine(int�к� - 1).Tag, ";")(3))
                    
                    '����ܹ��ж��ٸ���λ
                    lng��λ���� = (dbl��� - dbl��С) / dbl��λֵ
                    '�����λ������������д���20����ȡ20-����У������ȡ ��λ����+�����
                    lng��λ���� = IIf(lng��λ���� + lng����� > (MAXROWS - 1), (MAXROWS - 1) - lng�����, lng��λ���� + lng�����)
                    '����ֵ=((���ֵ-��ǰֵ)/��λֵ+�����-1)*�и߶�
                    
                    If zlCommFun.NVL(!��ֵ) <> "" Then
                        lngValue2 = 0
                        If InStr(!��ֵ, ",") > 0 Then
                            lngValue = ConvertToY(int�к� - 1, Val(Mid(!��ֵ, 1, InStr(!��ֵ, ",") - 1)))
                            lngValue2 = ConvertToY(int�к� - 1, Val(Mid(!��ֵ, InStr(!��ֵ, ",") + 1)))
                        Else
                            lngValue = ConvertToY(int�к� - 1, Val(!��ֵ))
                        End If
                    
                        aryValue() = Split(mshScale.TextMatrix(GraphDataRow.��������, intCol), ";")
    
                        '��¼ͬһʱ�������������
                        
                        strStart1 = Format(Int(CDate(strFrom)) + ((intCol - mshScale.FixedCols) * 4) / 24, "YYYY-MM-DD hh:mm:ss")
                        strEnd1 = Format(Int(CDate(strFrom)) + ((intCol - mshScale.FixedCols) * 4 + 4) / 24, "YYYY-MM-DD hh:mm:ss")
                        
                        If Val(aryValue(int�к�)) > 0 And zlCommFun.NVL(!��¼���, 0) = 1 Then
                            If strTmp = "����" And mint����Ӧ�� = 1 Then
                                aryValue(int�к�) = lngValue
                            Else
                                aryValue(int�к�) = aryValue(int�к�) & "," & lngValue
                            End If
                                                    
                        Else
                            aryValue(int�к�) = lngValue
                            
                            If !��Ŀ��� = mItemNo.���� Then
                                '�������²�λ
                                aryPart = Split(mshScale.TextMatrix(GraphDataRow.��λ��־, intCol), ";")
                                aryPart(int�к�) = zlCommFun.NVL(!���²�λ, "Ҹ��")
                                mshScale.TextMatrix(GraphDataRow.��λ��־, intCol) = Join(aryPart, ";")
                                mshScale.TextMatrix(GraphDataRow.���Ա�־, intCol) = zlCommFun.NVL(!���Ժϸ�, "0")
                            ElseIf !��Ŀ��� = mItemNo.���� Then
                                aryPart = Split(mshScale.TextMatrix(GraphDataRow.��λ��־, intCol), ";")
                                aryPart(int�к�) = zlCommFun.NVL(!���²�λ, "��������")
                                mshScale.TextMatrix(GraphDataRow.��λ��־, intCol) = Join(aryPart, ";")
                            ElseIf !��Ŀ��� = mItemNo.���� Then
                                aryPart = Split(mshScale.TextMatrix(GraphDataRow.��λ��־, intCol), ";")
                                aryPart(int�к�) = zlCommFun.NVL(!���²�λ, "")
                                mshScale.TextMatrix(GraphDataRow.��λ��־, intCol) = Join(aryPart, ";")
                            End If
                    
                        End If
                        mshScale.TextMatrix(GraphDataRow.��������, intCol) = Join(aryValue, ";")
                    End If
                    
                    '��д���ı�־
                    aryValue() = Split(mshScale.TextMatrix(GraphDataRow.���ı�־, intCol), ";")
                    aryValue(int�к�) = 1
                    mshScale.TextMatrix(GraphDataRow.���ı�־, intCol) = Join(aryValue, ";")
                    
                    '��д����ʱ��
                    aryValue() = Split(mshScale.TextMatrix(GraphDataRow.����ʱ��, intCol), ";")
                    aryValue(int�к�) = Format(!ʱ��, "yyyy-MM-dd HH:mm:ss")
                    mshScale.TextMatrix(GraphDataRow.����ʱ��, intCol) = Join(aryValue, ";")
                    
                    '��дδ��˵��
                    If zlCommFun.NVL(!��ֵ) = "����" Then
                        aryValue() = Split(mshScale.TextMatrix(GraphDataRow.δ��˵��, intCol), ";")
                        aryValue(int�к�) = "����"
                        mshScale.TextMatrix(GraphDataRow.δ��˵��, intCol) = Join(aryValue, ";")
                    Else
                        If zlCommFun.NVL(!δ��˵��) <> "" Then
                            aryValue() = Split(mshScale.TextMatrix(GraphDataRow.δ��˵��, intCol), ";")
                            aryValue(int�к�) = !δ��˵��
                            mshScale.TextMatrix(GraphDataRow.δ��˵��, intCol) = Join(aryValue, ";")
                        End If
                    End If
                End If
            End If
            
            .MoveNext
        Loop
    End With
    
    '��ȡ�������±��ֱ��¼�����ݣ���д������
    '------------------------------------------------------------------------------------------------------------------
    Call ReadGridData(Val(mrsParam("����ȼ�").Value), _
                        Val(mrsParam("����id").Value), _
                        IIf(Val(mrsParam("Ӥ��").Value) = 0, 1, 2), _
                        Val(mrsParam("����id").Value), _
                        Val(mrsParam("��ҳid").Value), _
                        CDate(Format(strFrom, "YYYY-MM-DD")), _
                        CDate(Format(strTo, "YYYY-MM-DD") & " 23:59:59"), _
                        Val(mrsParam("Ӥ��").Value), _
                        mblnMoved)
    
    mshUpTab.Redraw = True
    mshScale.Redraw = True
    
    '�����Դ�������,����������
    On Error Resume Next
    Err = 0
    Debug.Print 1 / 0
    If Err <> 0 Then
        Call OutputDadaForDebug
    End If
    
    Screen.MousePointer = vbDefault
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
ErrHead:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckGridData(ByVal intIndex As Integer) As Boolean
    '******************************************************************************************************************
    '���ܣ� ���ָ���ı��¼����Ŀ�Ƿ�������
    '������ intIndex : ����
    '���أ�
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim aryTmp As Variant
    Dim intCol As Integer
    Dim intCount As Integer
    Dim strTmp As String
    
    On Error GoTo errHand
    
    CheckGridData = True
    
    With mshDownTab
        For intLoop = .FixedCols To .Cols - 1
            If Trim(.TextMatrix(intIndex, intLoop)) <> "" Then
                Exit Function
            End If
            
            aryTmp = Split(.TextMatrix(GridDataRow.�޸ı�־, intLoop), ";")
            Select Case Val(aryTmp(intIndex - 1))
            Case OperateType.��������, OperateType.�޸Ĳ���, OperateType.ɾ������
                Exit Function
            End Select
        Next
    End With
        
    CheckGridData = False
    
    Exit Function
    
    '������
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Function

Private Function DeleteActiveItem(ByVal intIndex As Integer) As Boolean
    '******************************************************************************************************************
    '���ܣ� ���ָ���Ļ��Ŀ
    '������ intIndex : ���Ŀ����
    '���أ�
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim aryTmp As Variant
    Dim intCol As Integer
    Dim intCount As Integer
    Dim strTmp As String
    
    On Error GoTo errHand
    
    For intLoop = UBound(mItemStru) To LBound(mItemStru) Step -1
        If mItemStru(intLoop).���Ŀ And intIndex = intLoop Then

            With mshDownTab
                
                '�ƶ���ص�����
                For intCol = .FixedCols To .Cols - 1
                                        
                    aryTmp = Split(.TextMatrix(GridDataRow.�޸ı�־, intCol), ";")
                    
                    For intCount = intIndex To .Rows - 2
                        aryTmp(intCount) = aryTmp(intCount + 1)
                        mItemStru(intCount).��Ŀ���� = mItemStru(intCount + 1).��Ŀ����
                        mItemStru(intCount).�������� = mItemStru(intCount + 1).��������
                        mItemStru(intCount).���ݳ��� = mItemStru(intCount + 1).���ݳ���
                        mItemStru(intCount).С��λ�� = mItemStru(intCount + 1).С��λ��
                        mItemStru(intCount).��Сֵ = mItemStru(intCount + 1).��Сֵ
                        mItemStru(intCount).���ֵ = mItemStru(intCount + 1).���ֵ
                        mItemStru(intCount).��¼Ƶ�� = mItemStru(intCount + 1).��¼Ƶ��
                        mItemStru(intCount).���Ŀ = mItemStru(intCount + 1).���Ŀ
                        mItemStru(intCount).��Ŀ��� = mItemStru(intCount + 1).��Ŀ���
                    Next
                    aryTmp(.Rows - 2) = ""
                    
                    strTmp = Join(aryTmp, ";")
                    .TextMatrix(GridDataRow.�޸ı�־, intCol) = Left(strTmp, Len(strTmp) - 1)
                Next
                
                'ɾ���м��������
                .RemoveItem intIndex
                
                intCount = UBound(mItemStru) - 1
                ReDim Preserve mItemStru(intCount)
            End With

            DeleteActiveItem = True

            Exit For
        End If
    Next
    
    Exit Function
    
    '������
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Function

Private Function ReadGridData(ByVal byt����ȼ� As Byte, _
                                ByVal lng����id As Long, _
                                ByVal byt���ò��� As Byte, _
                                ByVal lng����id As Long, _
                                ByVal lng��ҳid As Long, _
                                ByVal dt��ʼʱ�� As Date, _
                                ByVal dt����ʱ�� As Date, ByVal bytӤ�� As Byte, Optional ByVal blnMoved As Boolean) As Boolean
    '******************************************************************************************************************
    '���ܣ� ��ȡ�������±��ֱ��¼�����ݣ���д������
    '������ lng����id : ����
    '       strFrom : ��ʼʱ��
    '       strTo : ��ֹʱ��
    '���أ�
    '******************************************************************************************************************
    Dim strItemOrder As String
    Dim strItemName As String
    Dim i As Long
    Dim aryValue() As String
    Dim aryTmp As Variant
    Dim intRow As Integer
    Dim lngColor As Long
    Dim strTime As String
    Dim intColTmp As Integer
    Dim rsTmp As New ADODB.Recordset
    Dim sgl������() As Single
    Dim sgl�ų���() As Single
    Dim sgl������2() As Single
    Dim sgl�ų���2() As Single
    Dim blnChanged As Boolean
    Dim intCount As Integer
    
    On Error GoTo errHand
        
    mshDownTab.Redraw = False
    
    '���Ŀ����
    '------------------------------------------------------------------------------------------------------------------
    '��������л��Ŀ
'    intCount = 0
    For i = UBound(mItemStru) To LBound(mItemStru) Step -1
        If mItemStru(i).���Ŀ Then
'            intCount = intCount + 1
      
            Call DeleteActiveItem(i)
            blnChanged = True
        Else
            Exit For
        End If
    Next
'    If blnChanged And mshDownTab.Rows - intCount > 0 Then
'        'ɾ�������
'        mshDownTab.Rows = mshDownTab.Rows - intCount
'
'        'ɾ����ص�������
'        intCount = UBound(mItemStru) - intCount
'        ReDim Preserve mItemStru(intCount)
'    End If
    
    '�Զ���ӻ��Ŀ��ֻ�ӵ�ǰҳ�������ݵģ�
    Set rsTmp = GetGridDataItem(byt����ȼ�, lng����id, byt���ò���, lng����id, lng��ҳid, dt��ʼʱ��, dt����ʱ��, bytӤ��, blnMoved)
    If rsTmp.BOF = False Then
        blnChanged = True
        Do While Not rsTmp.EOF
            Call AppendGridItem(rsTmp)
            rsTmp.MoveNext
        Loop
    End If
    
    '���µ�������ؼ�λ��
    If blnChanged Then Call picPane_Resize
    
    '��ʼ�޸ı�־���������
    '------------------------------------------------------------------------------------------------------------------
    With mshDownTab
        For intCol = .FixedCols To .Cols - 1
            .TextMatrix(GridDataRow.�޸ı�־, intCol) = String(.Rows - 2, ";")
        Next
        For intRow = 1 To .Rows - 1
            For intCol = .FixedCols To .Cols - 1
                .TextMatrix(intRow, intCol) = ""
            Next
        Next
    End With
    
    '��ȡ��Ŀ������嵥
    '------------------------------------------------------------------------------------------------------------------
    strItemOrder = ""
    strItemName = ""
    For intRow = mshDownTab.FixedRows To mshDownTab.Rows - 1
        If Val(mshDownTab.RowData(intRow)) <> mItemNo.Ѫѹ Then
            i = InStr(1, mshDownTab.TextMatrix(intRow, 0), "(", vbTextCompare)
            If i > 0 Then
                strItemName = strItemName & ",'" & Left(mshDownTab.TextMatrix(intRow, 0), i - 1) & "'"
            Else
                strItemName = strItemName & ",'" & mshDownTab.TextMatrix(intRow, 0) & "'"
            End If
        End If
    Next
    
    
    '��ȡ����
    '------------------------------------------------------------------------------------------------------------------
    If strItemName <> "" Then
        strItemName = Mid(strItemName, 2) & ",'����','����ѹ','����ѹ'"
                        
        mstrSQL = "SELECT a.����ʱ�� As ʱ��,C.��¼���� As ���,E.������Ŀ,D.��¼��,D.��Ŀ���,D.��¼Ƶ�� " & _
                    "FROM ���˻����¼ A,���˻������� C,���¼�¼��Ŀ D,�����¼��Ŀ E " & _
                    "Where A.ID = c.��¼ID " & _
                        "AND A.������Դ=2 " & _
                        "AND Nvl(a.Ӥ��,0)=[6] " & _
                        "AND A.����id=[1] " & _
                        "AND A.��ҳid=[2] " & _
                        "AND INSTR([5],','''||D.��¼��||''',')>0 " & _
                        "AND D.��Ŀ���=C.��Ŀ��� " & _
                        "AND c.��¼����=1 " & _
                        "AND E.��Ŀ���=D.��Ŀ��� " & _
                        "AND E.����ȼ�>=[7]  " & _
                        "AND a.����ʱ�� BETWEEN [3] And [4] And c.��ֹ�汾 Is Null " & _
                        "AND D.��¼��=2 " & _
                    "Order By Decode(D.��¼��,'����ѹ',0,1)," & strItemName & ",a.����ʱ��"
                    
        If blnMoved Then
            mstrSQL = Replace(mstrSQL, "���˻����¼", "H���˻����¼")
            mstrSQL = Replace(mstrSQL, "���˻�������", "H���˻�������")
        End If
        
        Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, mstrMsgTitle, lng����id, lng��ҳid, dt��ʼʱ��, dt����ʱ��, "," & strItemName & ",", bytӤ��, byt����ȼ�)
                                                    
        Dim intColFirst1 As Integer
        Dim intColFirst2 As Integer
        
        intColFirst1 = 0
        intColFirst2 = 0
        
        With rsTmp
            vsf.Cell(flexcpText, 1, 2, 1, vsf.Cols - 1) = ""
            vsf.Cell(flexcpData, 1, 2, 1, vsf.Cols - 1) = ""
            vsf.Cell(flexcpForeColor, 1, 2, 1, vsf.Cols - 1) = 200
                
            If rsTmp.RecordCount > 0 Then
                    
                ReDim sgl������(0 To mshDownTab.Cols)
                ReDim sgl�ų���(0 To mshDownTab.Cols)
                
                rsTmp.MoveFirst
                
                For i = 0 To rsTmp.RecordCount - 1
                                        
                    Select Case !��Ŀ���
                    Case mItemNo.����
                        intCol = GetCurveColumn(!ʱ��, dt��ʼʱ��, mlngHourBegin) + vsf.FixedCols - 1
                        
                        If intCol < vsf.Cols Then
                            vsf.TextMatrix(1, intCol) = zlCommFun.NVL(!���, "")
                        End If
                        
                    Case mItemNo.����ѹ
                        
                        intCol = Int((!ʱ�� - Int(dt��ʼʱ��)) * 24) \ 12 + mshDownTab.FixedCols
                        strTime = Format(Int(dt��ʼʱ��) + (intCol - mshDownTab.FixedCols) \ 2, "YYYY-MM-DD")
                        intRow = mItemSerial.Ѫѹ
                        
                        
                        '��Ժ����ĵ�һ���Խӽ���Ժʱ��Ϊ׼��Ҳ���������Ϊ׼������ȡ���һ��
                        If Format(!ʱ��, "yyyy-MM-dd") = Format(txtCard(3).Text, "yyyy-MM-dd") Then
                            If intColFirst2 = 0 Then
                                If intCol < mshDownTab.Cols Then
                                    If mshDownTab.TextMatrix(intRow, intCol) <> "" Or zlCommFun.NVL(!���, "") <> "" Then
                                        
                                        If InStr(mshDownTab.TextMatrix(intRow, intCol), "/") = 0 Then
                                            mshDownTab.TextMatrix(intRow, intCol) = mshDownTab.TextMatrix(intRow, intCol) & "/" & zlCommFun.NVL(!���, "")
                                        Else
                                            aryTmp = Split(mshDownTab.TextMatrix(intRow, intCol), "/")
                                            aryTmp(1) = zlCommFun.NVL(!���, "")
                                            mshDownTab.TextMatrix(intRow, intCol) = aryTmp(0) & "/" & aryTmp(1)
                                        End If
                                        
                                    End If
            
                                    mshDownTab.Cell(flexcpData, intRow, intCol) = 0
                                End If
                                
                                intColFirst2 = intCol
                            ElseIf intColFirst2 <> intCol Then
                                If intCol < mshDownTab.Cols Then
                                    If mshDownTab.TextMatrix(intRow, intCol) <> "" Or zlCommFun.NVL(!���, "") <> "" Then
                                        
                                        If InStr(mshDownTab.TextMatrix(intRow, intCol), "/") = 0 Then
                                            mshDownTab.TextMatrix(intRow, intCol) = mshDownTab.TextMatrix(intRow, intCol) & "/" & zlCommFun.NVL(!���, "")
                                        Else
                                            aryTmp = Split(mshDownTab.TextMatrix(intRow, intCol), "/")
                                            aryTmp(1) = zlCommFun.NVL(!���, "")
                                            mshDownTab.TextMatrix(intRow, intCol) = aryTmp(0) & "/" & aryTmp(1)
                                        End If
                                        
                                    End If
            
                                    mshDownTab.Cell(flexcpData, intRow, intCol) = 0
                                End If
                            End If
                            
                        Else
                            intColFirst2 = intCol
                            
                            If intCol < mshDownTab.Cols Then
                                If mshDownTab.TextMatrix(intRow, intCol) <> "" Or zlCommFun.NVL(!���, "") <> "" Then
                                    
                                    If InStr(mshDownTab.TextMatrix(intRow, intCol), "/") = 0 Then
                                        mshDownTab.TextMatrix(intRow, intCol) = mshDownTab.TextMatrix(intRow, intCol) & "/" & zlCommFun.NVL(!���, "")
                                    Else
                                        aryTmp = Split(mshDownTab.TextMatrix(intRow, intCol), "/")
                                        aryTmp(1) = zlCommFun.NVL(!���, "")
                                        mshDownTab.TextMatrix(intRow, intCol) = aryTmp(0) & "/" & aryTmp(1)
                                    End If
                                    
                                End If
        
                                mshDownTab.Cell(flexcpData, intRow, intCol) = 0
                            End If
                        End If

                    Case Else

                        For intRow = 1 To mshDownTab.Rows - 1
                            If Val(mshDownTab.RowData(intRow)) = !��Ŀ��� Then
                                Exit For
                            End If
                        Next
                        
                        intCol = Int((!ʱ�� - Int(dt��ʼʱ��)) * 24) \ 12 + mshDownTab.FixedCols
                        strTime = Format(Int(dt��ʼʱ��) + (intCol - mshDownTab.FixedCols) \ 2, "YYYY-MM-DD")
                                                                                   
                        If intCol < mshDownTab.Cols Then
                            Select Case !��Ŀ���
                            Case 7          '��Һ��
                                
                                If !��¼Ƶ�� = 1 Then
                                    intColTmp = IIf(intCol Mod 2 = 0, intCol + 1, intCol)
                                Else
                                    intColTmp = intCol
                                End If
                                
                                sgl������(intColTmp) = sgl������(intColTmp) + Val(zlCommFun.NVL(!���, ""))
                                mshDownTab.TextMatrix(intRow, intColTmp) = sgl������(intColTmp)
                                
                                mshDownTab.Cell(flexcpData, intRow, intColTmp) = 0
                            Case 9          '��Һ��
                            
                                If !��¼Ƶ�� = 1 Then
                                    intColTmp = IIf(intCol Mod 2 = 0, intCol + 1, intCol)
                                Else
                                    intColTmp = intCol
                                End If
                                
                                sgl�ų���(intColTmp) = sgl�ų���(intColTmp) + Val(zlCommFun.NVL(!���))
                                
                                If Right(zlCommFun.NVL(!���), 2) = "/C" Then
                                    mshDownTab.TextMatrix(intRow, intColTmp) = sgl�ų���(intColTmp) & "/C"
                                ElseIf Right(zlCommFun.NVL(!���), 1) = "C" Then
                                    mshDownTab.TextMatrix(intRow, intColTmp) = "C"
                                Else
                                    mshDownTab.TextMatrix(intRow, intColTmp) = sgl�ų���(intColTmp)
                                End If
                                
                                mshDownTab.Cell(flexcpData, intRow, intColTmp) = 0
                                
                                If Right(mshDownTab.TextMatrix(intRow, intColTmp), 1) = "C" Then
                                    mshDownTab.Cell(flexcpData, intRow, intColTmp) = 4
                                End If
                                
                                If Right(mshDownTab.TextMatrix(intRow, intColTmp), 2) = "/C" Then
                                    mshDownTab.Cell(flexcpData, intRow, intColTmp) = 5
                                End If
                            Case 10         '������
                            
                                mshDownTab.TextMatrix(intRow, intCol) = zlCommFun.NVL(!���, "")
                                
                                mshDownTab.Cell(flexcpData, intRow, intCol) = 0
                                If Right(mshDownTab.TextMatrix(intRow, intCol), 2) = "/E" Then
                                    mshDownTab.Cell(flexcpData, intRow, intCol) = 3
                                ElseIf Right(mshDownTab.TextMatrix(intRow, intCol), 1) = "E" Then
                                    mshDownTab.Cell(flexcpData, intRow, intCol) = 2
                                ElseIf Right(mshDownTab.TextMatrix(intRow, intCol), 1) = "*" Then
                                    mshDownTab.Cell(flexcpData, intRow, intCol) = 1
                                End If
                            
                            Case mItemNo.Ѫѹ
                                                            
                                '��Ժ����ĵ�һ���Խӽ���Ժʱ��Ϊ׼��Ҳ���������Ϊ׼������ȡ���һ��
                                If Format(!ʱ��, "yyyy-MM-dd") = Format(txtCard(3).Text, "yyyy-MM-dd") Then
                                    If intColFirst1 = 0 Then
                                        mshDownTab.TextMatrix(intRow, intCol) = zlCommFun.NVL(!���, "")
                                        If zlCommFun.NVL(!���, "") <> "" Then mshDownTab.TextMatrix(intRow, intCol) = mshDownTab.TextMatrix(intRow, intCol) & "/"
                                        mshDownTab.Cell(flexcpData, intRow, intCol) = 0
                                        
                                        intColFirst1 = intCol
                                        
                                    ElseIf intColFirst1 <> intCol Then
                                    
                                        mshDownTab.TextMatrix(intRow, intCol) = zlCommFun.NVL(!���, "")
                                        If zlCommFun.NVL(!���, "") <> "" Then mshDownTab.TextMatrix(intRow, intCol) = mshDownTab.TextMatrix(intRow, intCol) & "/"
                                        mshDownTab.Cell(flexcpData, intRow, intCol) = 0
                                        
                                    End If
                                    
                                Else
                                    intColFirst1 = intCol
                                    mshDownTab.TextMatrix(intRow, intCol) = zlCommFun.NVL(!���, "")
                                    If zlCommFun.NVL(!���, "") <> "" Then mshDownTab.TextMatrix(intRow, intCol) = mshDownTab.TextMatrix(intRow, intCol) & "/"
                                    mshDownTab.Cell(flexcpData, intRow, intCol) = 0
                                End If
                                
                            Case Else
                                
                                lngColor = GridTextColor(!��¼��, zlCommFun.NVL(!���, ""))

                                mshDownTab.Cell(flexcpForeColor, intRow, intCol, intRow, intCol) = lngColor
                                mshDownTab.TextMatrix(intRow, intCol) = zlCommFun.NVL(!���, "")
                                mshDownTab.Cell(flexcpData, intRow, intCol) = 0

                            End Select
    
                            If InStr(mshDownTab.TextMatrix(0, intCol), ";") = 0 Then
                                mshDownTab.TextMatrix(0, intCol) = "1"
                            Else
                                aryValue() = Split(mshDownTab.TextMatrix(0, intCol), ";")
                                aryValue(intRow - 1) = 1
                                mshDownTab.TextMatrix(0, intCol) = Join(aryValue, ";")
                            End If
                        End If
                    End Select
                    
                    .MoveNext
                Next
            End If
        End With
    End If
    
    mshDownTab.Redraw = True
    Call mshDownTab_RowColChange
    
    ReadGridData = True
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    mshDownTab.Redraw = True
    Call SaveErrLog
    
End Function

Private Function DrawScale() As Boolean
    '******************************************************************************************************************
    '���ܣ� ��picture�ϱ�ߣ����ڽ���ʱʹ��
    '������
    '���أ�
    '******************************************************************************************************************
    Dim i As Long
    Dim strTmp As String
    Dim intMinCol As Long
    Dim intMaxCol As Long
    Dim X0 As Long
    Dim Y0 As Long
    Dim X1 As Long
    Dim Y1 As Long
    Dim lngColor As Long
    
    picScale.Cls
    
    
    '�����С���ʱ�䷶Χ
    
    If picScale.Tag <> "" Then
        Call CalcMinMaxCol(picScale.Tag, intMinCol, intMaxCol)
        lblCur.Left = intMinCol * HOUR_STEP_Twips
    End If
    
    X0 = 0
    'Y0 = mshScale.ROWHEIGHT(0) / 2 'ԭ��  mshScale.ROWHEIGHT(0)=600 ����Ϊ300
    Y0 = mshScale.ROWHEIGHT(0)
    X1 = X0 + 15000
    'DrawLine picScale, X0, Y0, X1, Y0, &H8000000A
    For intCol = 1 To mshUpTab.Cols - 1
        X0 = intCol * HOUR_STEP_Twips * 6 - 15
        'Y0 = mshScale.ROWHEIGHT(0) / 2
        Y0 = 0
        Y1 = Y0 + 800
        
        DrawLine picScale, X0 - HOUR_STEP_Twips * 5, Y0, X0 - HOUR_STEP_Twips * 5, Y1, &H8000000A
        DrawLine picScale, X0 - HOUR_STEP_Twips * 4, Y0, X0 - HOUR_STEP_Twips * 4, Y1, &H8000000A
        DrawLine picScale, X0 - HOUR_STEP_Twips * 3, Y0, X0 - HOUR_STEP_Twips * 3, Y1, &H8000000A
        DrawLine picScale, X0 - HOUR_STEP_Twips * 2, Y0, X0 - HOUR_STEP_Twips * 2, Y1, &H8000000A
        DrawLine picScale, X0 - HOUR_STEP_Twips * 1, Y0, X0 - HOUR_STEP_Twips * 1, Y1, &H8000000A
        Y0 = 0
        'DrawLine picScale, X0 - HOUR_STEP_Twips * 3, Y0, X0 - HOUR_STEP_Twips * 3, Y1, &H8000000A
        DrawLine picScale, X0, Y0, X0, Y1, &H8000000A, , 2
        '�˴���225����HOUR_STEP_Twips��Ϊ260����Ϊ�޸���
        'DrawText picScale, X0 - HOUR_STEP_Twips * 6 + 225, 80, "����", &H80000012
        'DrawText picScale, X0 - HOUR_STEP_Twips * 3 + 225, 80, "����", &H80000012
        For i = 6 To 1 Step -1
            Select Case i
            Case 6
                strTmp = mlngHourBegin + 4 * 0
                lngColor = &H8080FF
            Case 5
                strTmp = mlngHourBegin + 4 * 1
                lngColor = &H8080FF
            Case 4
                strTmp = mlngHourBegin + 4 * 2
                lngColor = &H80000012
            Case 3
                lngColor = &H80000012
                strTmp = mlngHourBegin + 4 * 3
            Case 2
                lngColor = &H80000012
                strTmp = mlngHourBegin + 4 * 4
            Case 1
                lngColor = &H8080FF
                strTmp = mlngHourBegin + 4 * 5
            End Select
            
            '�˴���135����HOUR_STEP_Twips��Ϊ260����Ϊ�޸���
            If picScale.Tag <> "" Then
                DrawText picScale, X0 - HOUR_STEP_Twips * i + 135 - picScale.TextWidth(strTmp) / 2, 100, strTmp, IIf(intCol * 6 - i >= intMinCol And intCol * 6 - i <= intMaxCol, lngColor, &H8000000A)
            End If
        Next
    Next
End Function

Public Function DrawPaper() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ� ��picture�ϻ�����ֽ�����ڽ����ˢ������֮ǰʹ��
    '������
    '���أ�
    '------------------------------------------------------------------------------------------------------------------
    Dim X0 As Long
    Dim Y0 As Long
    Dim X1 As Long
    Dim Y1 As Long
    
    picGraph.Cls
    
    '����������ͼ��
    For intCol = 1 To mshUpTab.Cols - 1
        
        X0 = intCol * HOUR_STEP_Twips * 6 - 15
        Y0 = 0
        Y1 = Y0 + 15000
        
        DrawLine picGraph, X0 - HOUR_STEP_Twips * 5, Y0, X0 - HOUR_STEP_Twips * 5, Y1, &H8000000A
        DrawLine picGraph, X0 - HOUR_STEP_Twips * 4, Y0, X0 - HOUR_STEP_Twips * 4, Y1, &H8000000A
        DrawLine picGraph, X0 - HOUR_STEP_Twips * 3, Y0, X0 - HOUR_STEP_Twips * 3, Y1, &H8000000A
        DrawLine picGraph, X0 - HOUR_STEP_Twips * 2, Y0, X0 - HOUR_STEP_Twips * 2, Y1, &H8000000A
        DrawLine picGraph, X0 - HOUR_STEP_Twips * 1, Y0, X0 - HOUR_STEP_Twips * 1, Y1, &H8000000A
        DrawLine picGraph, X0 - HOUR_STEP_Twips * 0, Y0, X0 - HOUR_STEP_Twips * 0, Y1, &H8000000A, , 2
    Next
    
    '����������ͼ��
    For intRow = 1 To mshScale.Rows - 1
        
        X0 = 0
        Y0 = (intRow - 1) * ROWHEIGHT * 5
        
        X1 = X0 + 15000

        If (intRow - 1) Mod 5 = 0 Then
            If Int((intRow - 1) / 5) = 5 Then
                DrawLine picGraph, X0, Y0 + ROWHEIGHT * 5, X1, Y0 + ROWHEIGHT * 5, &H8080FF, 0, 2
            Else
                DrawLine picGraph, X0, Y0 + ROWHEIGHT * 5, X1, Y0 + ROWHEIGHT * 5, &H8000000A, 0, 2
            End If
        Else
            DrawLine picGraph, X0, Y0 + ROWHEIGHT * 5, X1, Y0 + ROWHEIGHT * 5, &H8000000A, 0
        End If
    Next
    

End Function

Private Function Ceil(ByVal dbValue As Double) As Integer
    '******************************************************************************************************************
    '���ܣ� ת��ʱ��Ϊ��ֵ
    '������
    '���أ�
    '******************************************************************************************************************
    
    Ceil = (0 - Int(0 - dbValue))
    Ceil = Int(dbValue + 0.5)
End Function

'Public Function DrawGraph() As Boolean
'    '******************************************************************************************************************
'    '���ܣ� �����Ѿ���д�����е�������ͼ
'    '������
'    '���أ�
'    '******************************************************************************************************************
'    Dim strComment As String
'    Dim strChar As String
'    Dim dblHeight As Double
'    Dim X0 As Single, Y0 As Single
'    Dim X1 As Single, Y1 As Single
'    Dim y As Single
'    Dim aryValue() As String
'    Dim aryNote() As String
'    Dim aryDots() As String
'    Dim lngColor As Long
'    Dim dblValues As Double
'    Dim strFrom As String, i As Long
'    Dim strDate0 As String, strDate1 As String
'    Dim strtmp As String
'    Dim intPointCount As Integer
'    Dim blnStop As Boolean
'    Dim bytδ����ʾλ�� As Byte
'    Dim mpt����() As POINTAPI
'    Dim mpt����() As POINTAPI
'    ReDim mpt����(0 To mshScale.Cols - mshScale.FixedCols - 1)
'    ReDim mpt����(0 To mshScale.Cols - mshScale.FixedCols - 1)
'    Dim rsPoint As ADODB.Recordset
'    Dim rs As New ADODB.Recordset
'    Dim varNote As Variant
'
'    On Error GoTo errHand
'
'    bytδ����ʾλ�� = Val(zlDatabase.GetPara("δ��˵����ʾλ��", glngSys, 1255, "0"))
'
'
'    Call PointInit(rsPoint)
'
'    strFrom = Split(picScale.Tag, ";")(0)
'
'    With mshScale
'        '������
'        .Row = 0
'        For intCol = 0 To .FixedCols - 1
'            intPointCount = -1
'            .Col = intCol
'            strChar = Mid(.Tag, intCol + 1, 1)
'            X0 = 0: Y0 = 0: strDate0 = ""
'            blnStop = False
'
'
'            For intCount = 0 To .Cols - .FixedCols - 1
'
'                strDate1 = Format(Int(CDate(strFrom)) + (intCount * 4 + 2) / 24, "yyyy-MM-dd")
'                aryValue = Split(.TextMatrix(GraphDataRow.��������, intCount + .FixedCols), ";")
'
'                If Trim(aryValue(intCol + 1)) <> "" And Val(aryValue(intCol + 1)) > 0 Then
'
'                    X1 = HOUR_STEP_Twips * intCount + (HOUR_STEP_Twips / 2)
'                    aryDots = Split(aryValue(intCol + 1), ",")
'
'                    For i = 0 To UBound(aryDots)
'                        Y1 = aryDots(i)
'                        strChar = Mid(.Tag, intCol + 1, 1)
'                        lngColor = .CellForeColor
'
'                        If X0 <> 0 Then
'                            If i = 0 Then
'
'                                Select Case intCol
'                                Case mItemSerial.����
'                                    Select Case Split(.TextMatrix(GraphDataRow.��λ��־, intCount + .FixedCols), ";")(intCol + 1)
'                                    Case "����"
'                                        strChar = mstrChar(0)
'                                    Case "Ҹ��"
'                                        strChar = mstrChar(1)
'                                    Case "����"
'                                        strChar = mstrChar(2)
'                                    Case Else
'                                        strChar = mstrChar(1)
'                                    End Select
'
'                                    Select Case ConvertToValue(intCol, Y1)
'                                    Case Is <= GetMinValue(intCol)
'                                        '����35������
'                                        strChar = "��"
'                                        Y1 = ConvertToY(intCol, GetMinValue(intCol))
'                                        Call DrawLine(picGraph, X1, Y1, X1, Y1 + 200, lngColor, , , True)
'
'                                    Case Is >= GetMaxValue(intCol)
'                                        '����42������
'                                        strChar = "��"
'                                        Y1 = ConvertToY(intCol, GetMaxValue(intCol))
'                                        Call DrawLine(picGraph, X1, Y1, X1, Y1 - 200, lngColor, , , True)
'                                    End Select
'
'                                    If Val(.TextMatrix(10, intCount + .FixedCols)) = 1 Then
'                                        '���Ժϸ�
'                                        Call DrawText(picGraph, X1 - 50, Y1 - 250, "v", lngColor)
'                                    End If
'
'                                Case mItemSerial.����
'                                    If Split(.TextMatrix(GraphDataRow.��λ��־, intCount + .FixedCols), ";")(intCol + 1) = "" Then
'                                        strChar = mstrPulse
'                                    Else
'                                        strChar = ""
'                                    End If
'
'                                    Select Case ConvertToValue(intCol, Y1)
'                                    Case Is <= GetMinValue(intCol)
'                                        Y1 = ConvertToY(intCol, GetMinValue(intCol))
'                                        Call DrawLine(picGraph, X1, Y1, X1, Y1 + 200, lngColor, , , True)
'                                    Case Is >= GetMaxValue(intCol)
'                                        Y1 = ConvertToY(intCol, GetMaxValue(intCol))
'                                        Call DrawLine(picGraph, X1, Y1, X1, Y1 - 200, lngColor, , , True)
'                                    End Select
'
'                                    mpt����(intCount).x = X1
'                                    mpt����(intCount).y = Y1
'
'                                Case mItemSerial.����
'                                    Select Case ConvertToValue(intCol, Y1)
'                                    Case Is <= GetMinValue(intCol)
'                                        Y1 = ConvertToY(intCol, GetMinValue(intCol))
'                                        Call DrawLine(picGraph, X1, Y1, X1, Y1 + 200, lngColor, , , True)
'                                    Case Is >= GetMaxValue(intCol)
'                                        Y1 = ConvertToY(intCol, GetMaxValue(intCol))
'                                        Call DrawLine(picGraph, X1, Y1, X1, Y1 - 200, lngColor, , , True)
'                                    End Select
'
'                                    mpt����(intCount).x = X1
'                                    mpt����(intCount).y = Y1
'
'                                Case mItemSerial.����
'                                    If Split(.TextMatrix(GraphDataRow.��λ��־, intCount + .FixedCols), ";")(intCol + 1) = "��������" Then
'                                        strChar = mstrBreath
'                                    Else
'                                        strChar = ""
'                                    End If
'
'                                    Select Case ConvertToValue(intCol, Y1)
'                                    Case Is <= GetMinValue(intCol)
'                                        Y1 = ConvertToY(intCol, GetMinValue(intCol))
'                                        Call DrawLine(picGraph, X1, Y1, X1, Y1 + 200, lngColor, , , True)
'                                    Case Is >= GetMaxValue(intCol)
'                                        Y1 = ConvertToY(intCol, GetMaxValue(intCol))
'                                        Call DrawLine(picGraph, X1, Y1, X1, Y1 - 200, lngColor, , , True)
'                                    End Select
'                                End Select
'
'                                '���һ�������ݲŲ�����;2.����м���˵����Ҫ����ߵ�
'                                If (DateDiff("d", CDate(strDate0), CDate(strDate1)) <= 1) Then
'                                    If blnStop = False Then
'                                        '���δ��˵����"����",�����ϸ���㻭������
'                                        If intCol = mItemSerial.���� And Split(.TextMatrix(GraphDataRow.δ��˵��, intCount + .FixedCols), ";")(intCol + 1) = "����" Then
'                                            'nothing to do
'                                        Else
'                                            DrawLine picGraph, X0, Y0, X1, Y1, .CellForeColor
'                                        End If
'                                    End If
'                                    blnStop = False
'                                End If
'
'                            Else
'                                Select Case intCol
'                                Case mItemSerial.����
'
'                                    '������
'                                    lngColor = &HFF&
'                                    strChar = "��"
'                                    If Y1 < Y0 Then
'
'                                        '������ʧ�ܣ�������ͷ�ĺ�ɫʵ�ߣ��ַ��̶��á�
'                                        Call DrawLine(picGraph, X0, Y0, X1, Y1, lngColor, , , True)
'
'                                    ElseIf Y1 > Y0 Then
'
'                                        '�����³ɹ�������ɫ���ߣ��ַ��̶��á�
'                                        Call DrawLine(picGraph, X0, Y0, X1, Y1, lngColor, 2)
'
'                                    End If
'
'                                Case mItemSerial.����
'                                    If Y1 <> Y0 Then
'                                        lngColor = &HFF&
'                                        strChar = mstr���ʷ���
'
'                                        Select Case ConvertToValue(intCol, Y1)
'                                        Case Is <= GetMinValue(intCol)
'                                            Y1 = ConvertToY(intCol, GetMinValue(intCol))
'                                            Call DrawLine(picGraph, X1, Y1, X1, Y1 + 200, lngColor, , , True)
'                                        Case Is >= GetMaxValue(intCol)
'                                            Y1 = ConvertToY(intCol, GetMaxValue(intCol))
'                                            Call DrawLine(picGraph, X1, Y1, X1, Y1 - 200, lngColor, , , True)
'                                        End Select
'
'                                        mpt����(intCount).x = X1
'                                        mpt����(intCount).y = Y1
'
'                                    End If
'                                Case mItemSerial.����
'
'                                    Select Case ConvertToValue(intCol, Y1)
'                                    Case Is <= GetMinValue(intCol)
'                                        Y1 = ConvertToY(intCol, GetMinValue(intCol))
'                                        Call DrawLine(picGraph, X1, Y1, X1, Y1 + 200, lngColor, , , True)
'                                    Case Is >= GetMaxValue(intCol)
'                                        Y1 = ConvertToY(intCol, GetMaxValue(intCol))
'                                        Call DrawLine(picGraph, X1, Y1, X1, Y1 - 200, lngColor, , , True)
'                                    End Select
'
'                                    mpt����(intCount).x = X1
'                                    mpt����(intCount).y = Y1
'
'                                Case mItemSerial.����
'                                    If Split(.TextMatrix(GraphDataRow.��λ��־, intCount + .FixedCols), ";")(intCol + 1) = "��������" Then
'                                        strChar = mstrBreath
'                                    Else
'                                        strChar = ""
'                                    End If
'
'                                    Select Case ConvertToValue(intCol, Y1)
'                                    Case Is <= GetMinValue(intCol)
'                                        Y1 = ConvertToY(intCol, GetMinValue(intCol))
'                                        Call DrawLine(picGraph, X1, Y1, X1, Y1 + 200, lngColor, , , True)
'                                    Case Is >= GetMaxValue(intCol)
'                                        Y1 = ConvertToY(intCol, GetMaxValue(intCol))
'                                        Call DrawLine(picGraph, X1, Y1, X1, Y1 - 200, lngColor, , , True)
'                                    End Select
'                                End Select
'                            End If
'                        Else
'
'                            Select Case intCol
'                            Case mItemSerial.����
'                                Select Case Split(.TextMatrix(GraphDataRow.��λ��־, intCount + .FixedCols), ";")(intCol + 1)
'                                Case "����"
'                                    strChar = mstrChar(0)
'                                Case "Ҹ��"
'                                    strChar = mstrChar(1)
'                                Case "����"
'                                    strChar = mstrChar(2)
'                                Case Else
'                                    strChar = mstrChar(1)
'                                End Select
'
'                                Select Case ConvertToValue(intCol, Y1)
'                                Case Is <= GetMinValue(intCol)
'                                    '����35������
'                                    strChar = "��"
'                                    Y1 = ConvertToY(intCol, GetMinValue(intCol))
'                                    Call DrawLine(picGraph, X1, Y1, X1, Y1 + 200, lngColor, , , True)
'
'                                Case Is >= GetMaxValue(intCol)
'                                    '����42������
'                                    strChar = "��"
'                                    Y1 = ConvertToY(intCol, GetMaxValue(intCol))
'                                    Call DrawLine(picGraph, X1, Y1, X1, Y1 - 200, lngColor, , , True)
'
'                                End Select
'
'                                If Val(.TextMatrix(10, intCount + .FixedCols)) = 1 Then
'                                    '���Ժϸ�
'                                    Call DrawText(picGraph, X1 - 50, Y1 - 250, "v", lngColor)
'                                End If
'
'                            Case mItemSerial.����
'                                If Split(.TextMatrix(GraphDataRow.��λ��־, intCount + .FixedCols), ";")(intCol + 1) = "" Then
'                                    strChar = mstrPulse
'                                Else
'                                    strChar = ""
'                                End If
'
'                                Select Case ConvertToValue(intCol, Y1)
'                                Case Is <= GetMinValue(intCol)
'                                    Y1 = ConvertToY(intCol, GetMinValue(intCol))
'                                    Call DrawLine(picGraph, X1, Y1, X1, Y1 + 200, lngColor, , , True)
'                                Case Is >= GetMaxValue(intCol)
'                                    Y1 = ConvertToY(intCol, GetMaxValue(intCol))
'                                    Call DrawLine(picGraph, X1, Y1, X1, Y1 - 200, lngColor, , , True)
'                                End Select
'
'                                mpt����(intCount).x = X1
'                                mpt����(intCount).y = Y1
'
'                            Case mItemSerial.����
'
'                                Select Case ConvertToValue(intCol, Y1)
'                                Case Is <= GetMinValue(intCol)
'                                    Y1 = ConvertToY(intCol, GetMinValue(intCol))
'                                    Call DrawLine(picGraph, X1, Y1, X1, Y1 + 200, lngColor, , , True)
'                                Case Is >= GetMaxValue(intCol)
'                                    Y1 = ConvertToY(intCol, GetMaxValue(intCol))
'                                    Call DrawLine(picGraph, X1, Y1, X1, Y1 - 200, lngColor, , , True)
'                                End Select
'
'                                mpt����(intCount).x = X1
'                                mpt����(intCount).y = Y1
'
'                            Case mItemSerial.����
'                                If Split(.TextMatrix(GraphDataRow.��λ��־, intCount + .FixedCols), ";")(intCol + 1) = "��������" Then
'                                    strChar = mstrBreath
'                                Else
'                                    strChar = ""
'                                End If
'
'                                Select Case ConvertToValue(intCol, Y1)
'                                Case Is <= GetMinValue(intCol)
'                                    Y1 = ConvertToY(intCol, GetMinValue(intCol))
'                                    Call DrawLine(picGraph, X1, Y1, X1, Y1 + 200, lngColor, , , True)
'                                Case Is >= GetMaxValue(intCol)
'                                    Y1 = ConvertToY(intCol, GetMaxValue(intCol))
'                                    Call DrawLine(picGraph, X1, Y1, X1, Y1 - 200, lngColor, , , True)
'                                End Select
'
'                            End Select
'                        End If
'
'                        If intCol = mItemSerial.���� Then
'                            Call PointAdd(rsPoint, X1, Y1, mshScale.ColData(intCol), strChar, lngColor, intCount, mshScale.TextMatrix(4, intCount + .FixedCols))
'                        ElseIf intCol = mItemSerial.���� Then
'                            Call PointAdd(rsPoint, X1, Y1, mshScale.ColData(intCol), strChar, lngColor, intCount, mshScale.TextMatrix(4, intCount + .FixedCols), IIf(strChar = "", "BREATH", ""))
'                        ElseIf intCol = mItemSerial.���� Then
'                            Call PointAdd(rsPoint, X1, Y1, mshScale.ColData(intCol), strChar, lngColor, intCount, mshScale.TextMatrix(4, intCount + .FixedCols), IIf(strChar = "", "PACEMAKER", ""))
'                        Else
'                            Call PointAdd(rsPoint, X1, Y1, mshScale.ColData(intCol), strChar, lngColor, intCount, "")
'                        End If
'
'                        '��¼�ϴλ���λ�ú�����
'                        If X0 <> 0 And i <> 0 Then X1 = X0: Y1 = Y0 '�ӵ�һ������һ������
'                        X0 = X1: Y0 = Y1: strDate0 = strDate1
'
'                        blnStop = False
'                    Next i
'                End If
'
''                If blnStop = False Then
''                    If (.TextMatrix(GraphDataRow.�ϱ�˵��, intCount + .FixedCols) <> "" Or .TextMatrix(GraphDataRow.�±�˵��, intCount + .FixedCols) <> "") Then
''                        blnStop = (Val(.TextMatrix(GraphDataRow.�Ͽ���־, intCount + .FixedCols)) = 1)
''                    End If
''                End If
'
'                If blnStop = False Then
'
'                    aryNote = Split(.TextMatrix(GraphDataRow.δ��˵��, intCount + .FixedCols), ";")
'                    blnStop = (Trim(aryNote(intCol + 1)) <> "")
'
'                    '���±����ǿ��Բ�Ҫ�˵ģ������ǵ���ǰ������
'                    If blnStop = False Then
'                        If (.TextMatrix(GraphDataRow.�ϱ�˵��, intCount + .FixedCols) <> "" Or .TextMatrix(GraphDataRow.�±�˵��, intCount + .FixedCols) <> "") Then
'                            blnStop = (Val(.TextMatrix(GraphDataRow.�Ͽ���־, intCount + .FixedCols)) = 1)
'                        End If
'                    End If
'
'                End If
'
'            Next intCount
'
'        Next intCol
'
'
'        '������ַ���ͼ��
'        '--------------------------------------------------------------------------------------------------------------
'        Call DrawPoint(picGraph, rsPoint)
'
'        '�������������������γɶ���Σ����������ߺ����
'        '--------------------------------------------------------------------------------------------------------------
'        Call DrawPoly(picGraph, mpt����, mpt����)
'
'        Dim lngYMax As Long
'        lngYMax = ConvertToY(mItemSerial.����, 34.2)
'
'        '��ӡ���ת��־
'        '--------------------------------------------------------------------------------------------------------------
'        Dim intLoop As Integer
'        Dim rsTmp As ADODB.Recordset
'
'        '20090926:������40-42�ȼ��ӡ,����һ����Ϣ�����������С����,�ж�����Ϣ���Ӻ���һ���ӡ,��������һ���ֱ��ȫ����ӡ
'        Set rsTmp = New ADODB.Recordset
'        rsTmp.Fields.Append "�к�", adVarChar, 30
'        rsTmp.Fields.Append "ʱ��", adVarChar, 30
'        rsTmp.Fields.Append "���", adVarChar, 50
'        '20090926--
'        rsTmp.Fields.Append "��ӡ��", adVarChar, 30
'        rsTmp.Fields.Append "����", adVarChar, 50
'        '----------
'        rsTmp.Open
'
'        Dim intCharNumber As Integer
'
'        For intCol = 0 To .Cols - .FixedCols - 1
'
'            X1 = HOUR_STEP_Twips * intCol + HOUR_STEP_Twips / 2
''            Y1 = ConvertToY(mItemSerial.����, 42)
'            Y1 = 195
'            dblHeight = lngYMax - Y1
'
'            '�к�:=3��ʾ����;=5��ʾ��Ժ;=6��ʾת��;=7��ʾ����;=8��ʾ��Ժ,=13����
'            rsTmp.Filter = ""
'            For intLoop = 5 To 9
'                rsTmp.AddNew
'                rsTmp.Fields("�к�").Value = intCol + .FixedCols
'                rsTmp.Fields("ʱ��").Value = .Cell(flexcpData, intLoop, intCol + .FixedCols, intLoop, intCol + .FixedCols)
'                rsTmp.Fields("���").Value = .TextMatrix(intLoop, intCol + .FixedCols)
'            Next
'            rsTmp.AddNew
'            rsTmp.Fields("�к�").Value = intCol + .FixedCols
'            rsTmp.Fields("ʱ��").Value = .Cell(flexcpData, 3, intCol + .FixedCols, 3, intCol + .FixedCols)
'            rsTmp.Fields("���").Value = .TextMatrix(3, intCol + .FixedCols)
'
'            rsTmp.AddNew
'            rsTmp.Fields("�к�").Value = intCol + .FixedCols
'            rsTmp.Fields("ʱ��").Value = .Cell(flexcpData, 13, intCol + .FixedCols, 13, intCol + .FixedCols)
'            rsTmp.Fields("���").Value = .TextMatrix(13, intCol + .FixedCols)
'
'            'һ����Ϣ��һ��
'            strComment = ""
'            rsTmp.Filter = "�к�=" & intCol + .FixedCols
'            If rsTmp.RecordCount > 0 Then
'                rsTmp.Sort = "ʱ��"
'                rsTmp.MoveFirst
'                Do While Not rsTmp.EOF
'                    If strComment = "" Then
'                        strComment = rsTmp.Fields("���").Value
'                    Else
'                        strComment = Trim(strComment) & " " & rsTmp.Fields("���").Value
'                    End If
'                    rsTmp.MoveNext
'                Loop
'            End If
'
'            If Trim(strComment) <> "" Then
'                intCharNumber = 0
'                For intCount = 1 To Len(strComment)
'
'                    If Y1 < lngYMax Then
'                        strChar = Mid(strComment, intCount, 1)
'                        '��ɫ
'
'                        If Asc(strChar) < 0 Then
'                            If intCharNumber Mod 2 = 1 Then Y1 = Y1 + ROWHEIGHT * 2.5
'                        End If
'
'                        Call DrawRotateText(picGraph, X1 - picGraph.TextWidth(strChar) / 2, Y1 + 15, strChar, 255)
'                        If Asc(strChar) < 0 Then
'                            intCharNumber = 0
'                            Y1 = Y1 + ROWHEIGHT * 5
'                        Else
'                            Y1 = Y1 + ROWHEIGHT * 2.5
'                            intCharNumber = intCharNumber + 1
'                        End If
'                    End If
'                Next
'            End If
'
'            'δ��˵��
'            '----------------------------------------------------------------------------------------------------------
'            If bytδ����ʾλ�� = 0 Then
'                strComment = IIf(Trim(strComment) = "", "", " ")
'                strtmp = ""
'                varNote = Split(.TextMatrix(GraphDataRow.δ��˵��, intCol + .FixedCols), ";")
'                For intCount = 0 To UBound(varNote)
'                    If varNote(intCount) <> "����" Then
'                        If InStr(";" & strtmp & ";", ";" & varNote(intCount) & ";") = 0 Then
'                            strtmp = strtmp & ";" & varNote(intCount)
'                        End If
'                    End If
'                Next
'                If strtmp <> "" Then
'                    varNote = Split(strtmp, ";")
'                    For intCount = 0 To UBound(varNote)
'                        If strComment = "" Or strComment = " " Then
'                            strComment = strComment & varNote(intCount)
'                        Else
'                            strComment = strComment & " " & varNote(intCount)
'                        End If
'                    Next
'                End If
'
'                If Trim(strComment) <> "" Then
'
'                    intCharNumber = 0
'                    For intCount = 1 To Len(strComment)
'                        If Y1 <= lngYMax Then
'                            strChar = Mid(strComment, intCount, 1)
'                            '��ɫ
'
'                            If Asc(strChar) < 0 Then
'                                If intCharNumber Mod 2 = 1 Then Y1 = Y1 + ROWHEIGHT * 2.5
'                            End If
'
'                            Call DrawRotateText(picGraph, X1 - picGraph.TextWidth(strChar) / 2, Y1 + 15, strChar, -2147483635)
'
'                            If Asc(strChar) < 0 Then
'                                intCharNumber = 0
'                                Y1 = Y1 + ROWHEIGHT * 5
'                            Else
'                                Y1 = Y1 + ROWHEIGHT * 2.5
'                                intCharNumber = intCharNumber + 1
'                            End If
'                        End If
'                    Next
'                End If
'            End If
'
'            '�ϱ�˵��
'            '----------------------------------------------------------------------------------------------------------
'            strComment = IIf(Trim(strComment) = "", "", " ") & Trim(.TextMatrix(GraphDataRow.�ϱ�˵��, intCol + .FixedCols))
'            If Trim(strComment) <> "" Then
'
'                intCharNumber = 0
'                For intCount = 1 To Len(strComment)
'                    If Y1 <= lngYMax Then
'                        strChar = Mid(strComment, intCount, 1)
'                        '��ɫ
'
'                        If Asc(strChar) < 0 Then
'                            If intCharNumber Mod 2 = 1 Then Y1 = Y1 + ROWHEIGHT * 2.5
'                        End If
'
'                        Call DrawRotateText(picGraph, X1 - picGraph.TextWidth(strChar) / 2, Y1 + 15, strChar, -2147483635)
'
'                        If Asc(strChar) < 0 Then
'                            intCharNumber = 0
'                            Y1 = Y1 + ROWHEIGHT * 5
'                        Else
'                            Y1 = Y1 + ROWHEIGHT * 2.5
'                            intCharNumber = intCharNumber + 1
'                        End If
'                    End If
'                Next
'            End If
'
'            '�±�˵��
'            '----------------------------------------------------------------------------------------------------------
'
''            Y1 = ConvertToY(mItemSerial.����, 35)
'            Y1 = 7020
'            strComment = ""
'
'            'δ��˵��
'            '----------------------------------------------------------------------------------------------------------
'            If bytδ����ʾλ�� = 1 Then
'                strComment = IIf(Trim(strComment) = "", "", " ")
'                strtmp = ""
'                varNote = Split(.TextMatrix(GraphDataRow.δ��˵��, intCol + .FixedCols), ";")
'                For intCount = 0 To UBound(varNote)
'                    If varNote(intCount) <> "����" Then
'                        If InStr(";" & strtmp & ";", ";" & varNote(intCount) & ";") = 0 Then
'                            strtmp = strtmp & ";" & varNote(intCount)
'                        End If
'                    End If
'                Next
'                If strtmp <> "" Then
'                    varNote = Split(strtmp, ";")
'                    For intCount = 0 To UBound(varNote)
'                        If strComment = "" Or strComment = " " Then
'                            strComment = strComment & varNote(intCount)
'                        Else
'                            strComment = strComment & " " & varNote(intCount)
'                        End If
'                    Next
'                End If
'
'                If Trim(strComment) <> "" Then
'
'                    intCharNumber = 0
'                    For intCount = 1 To Len(strComment)
'                        If Y1 <= lngYMax Then
'                            strChar = Mid(strComment, intCount, 1)
'                            '��ɫ
'
'                            If Asc(strChar) < 0 Then
'                                If intCharNumber Mod 2 = 1 Then Y1 = Y1 + ROWHEIGHT * 2.5
'                            End If
'
'                            Call DrawRotateText(picGraph, X1 - picGraph.TextWidth(strChar) / 2, Y1 + 15, strChar, -2147483635)
'
'                            If Asc(strChar) < 0 Then
'                                intCharNumber = 0
'                                Y1 = Y1 + ROWHEIGHT * 5
'                            Else
'                                Y1 = Y1 + ROWHEIGHT * 2.5
'                                intCharNumber = intCharNumber + 1
'                            End If
'                        End If
'                    Next
'                End If
'            End If
'
'            strComment = IIf(Trim(strComment) = "", "", " ") & .TextMatrix(GraphDataRow.�±�˵��, intCol + .FixedCols)
'            If Trim(strComment) <> "" Then
'                intCharNumber = 0
'                For intCount = 1 To Len(strComment)
'                    If Y1 <= lngYMax Then
'                        strChar = Mid(strComment, intCount, 1)
'                        '��ɫ
'
'                        If Asc(strChar) < 0 Then
'                            If intCharNumber Mod 2 = 1 Then Y1 = Y1 + ROWHEIGHT * 2.5
'                        End If
'
'                        Call DrawRotateText(picGraph, X1 - picGraph.TextWidth(strChar) / 2, Y1 + 15, strChar, -2147483635)
'
'                        If Asc(strChar) < 0 Then
'                            intCharNumber = 0
'                            Y1 = Y1 + ROWHEIGHT * 5
'                        Else
'                            Y1 = Y1 + ROWHEIGHT * 2.5
'                            intCharNumber = intCharNumber + 1
'                        End If
'                    End If
'                Next
'            End If
'
'        Next
'
'    End With
'
'    Exit Function
'
'    '------------------------------------------------------------------------------------------------------------------
'errHand:
'    If ErrCenter = 1 Then
'        Resume
'    End If
'
'End Function

Public Function DrawGraph() As Boolean
    '******************************************************************************************************************
    '���ܣ� �����Ѿ���д�����е�������ͼ
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strComment As String
    Dim strChar As String, strChar1 As String
    Dim dblHeight As Double         '40-42��֮�����Ч��ӡ�߶�
    Dim X0 As Single, Y0 As Single
    Dim X1 As Single, Y1 As Single
    Dim Y As Single
    Dim aryValue() As String
    Dim aryNote() As String
    Dim aryDots() As String
    Dim lngColor As Long
    Dim dblValues As Double
    Dim strFrom As String, i As Long
    Dim strDate0 As String, strDate1 As String
    Dim strTmp As String
    Dim intPointCount As Integer
    Dim blnStop As Boolean
    Dim bytδ����ʾλ�� As Byte
    Dim mpt����() As POINTAPI
    Dim mpt����() As POINTAPI
    ReDim mpt����(0 To mshScale.Cols - mshScale.FixedCols - 1)
    ReDim mpt����(0 To mshScale.Cols - mshScale.FixedCols - 1)
    Dim rsPoint As ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim varNote As Variant
    
    On Error GoTo errHand
    
    bytδ����ʾλ�� = Val(zlDatabase.GetPara("δ��˵����ʾλ��", glngSys, 1255, "0"))
           

    Call PointInit(rsPoint)
    
    strFrom = Split(picScale.Tag, ";")(0)
    
    With mshScale
        '������
        .Row = 0
        For intCol = 0 To .FixedCols - 1
            intPointCount = -1
            .Col = intCol
            strChar = Mid(.Tag, intCol + 1, 1)
            X0 = 0: Y0 = 0: strDate0 = "": strChar1 = mstrBreath
            blnStop = False
            
            
            For intCount = 0 To .Cols - .FixedCols - 1
                
                strDate1 = Format(Int(CDate(strFrom)) + (intCount * 4 + 2) / 24, "yyyy-MM-dd")
                aryValue = Split(.TextMatrix(GraphDataRow.��������, intCount + .FixedCols), ";")

                If Trim(aryValue(intCol + 1)) <> "" And Val(aryValue(intCol + 1)) > 0 Then
                
                    X1 = HOUR_STEP_Twips * intCount + (HOUR_STEP_Twips / 2)
                    aryDots = Split(aryValue(intCol + 1), ",")
                    
                    For i = 0 To UBound(aryDots)
                        Y1 = aryDots(i)
                        strChar = Mid(.Tag, intCol + 1, 1)
                        lngColor = .CellForeColor
                        
                        If X0 <> 0 Then
                            If i = 0 Then
                                
                                Select Case intCol
                                Case mItemSerial.����
                                    Select Case Split(.TextMatrix(GraphDataRow.��λ��־, intCount + .FixedCols), ";")(intCol + 1)
                                    Case "����"
                                        strChar = mstrChar(0)
                                    Case "Ҹ��"
                                        strChar = mstrChar(1)
                                    Case "����"
                                        strChar = mstrChar(2)
                                    Case Else
                                        strChar = mstrChar(1)
                                    End Select

                                    Select Case ConvertToValue(intCol, Y1)
                                    Case Is <= GetMinValue(intCol)
                                        '����35������
                                        strChar = "��"
                                        Y1 = ConvertToY(intCol, GetMinValue(intCol))
                                        Call DrawLine(picGraph, X1, Y1, X1, Y1 + 200, lngColor, , , True)
                                        
                                    Case Is >= GetMaxValue(intCol)
                                        '����42������
                                        strChar = "��"
                                        Y1 = ConvertToY(intCol, GetMaxValue(intCol))
                                        Call DrawLine(picGraph, X1, Y1, X1, Y1 - 200, lngColor, , , True)
                                    End Select
                                    
                                    If Val(.TextMatrix(10, intCount + .FixedCols)) = 1 Then
                                        '���Ժϸ�
                                        Call DrawText(picGraph, X1 - 50, Y1 - 250, "v", lngColor)
                                    End If
                                    
                                Case mItemSerial.����
                                    If Split(.TextMatrix(GraphDataRow.��λ��־, intCount + .FixedCols), ";")(intCol + 1) = "" Then
                                        strChar = mstrPulse
                                    Else
                                        strChar = ""
                                    End If
                                    
                                    Select Case ConvertToValue(intCol, Y1)
                                    Case Is <= GetMinValue(intCol)
                                        Y1 = ConvertToY(intCol, GetMinValue(intCol))
                                        Call DrawLine(picGraph, X1, Y1, X1, Y1 + 200, lngColor, , , True)
                                    Case Is >= GetMaxValue(intCol)
                                        Y1 = ConvertToY(intCol, GetMaxValue(intCol))
                                        Call DrawLine(picGraph, X1, Y1, X1, Y1 - 200, lngColor, , , True)
                                    End Select
                                    
                                    mpt����(intCount).X = X1
                                    mpt����(intCount).Y = Y1
                                    
                                Case mItemSerial.����
                                    Select Case ConvertToValue(intCol, Y1)
                                    Case Is <= GetMinValue(intCol)
                                        Y1 = ConvertToY(intCol, GetMinValue(intCol))
                                        Call DrawLine(picGraph, X1, Y1, X1, Y1 + 200, lngColor, , , True)
                                    Case Is >= GetMaxValue(intCol)
                                        Y1 = ConvertToY(intCol, GetMaxValue(intCol))
                                        Call DrawLine(picGraph, X1, Y1, X1, Y1 - 200, lngColor, , , True)
                                    End Select
                                    
                                    mpt����(intCount).X = X1
                                    mpt����(intCount).Y = Y1
                                    
                                Case mItemSerial.����
                                    If Split(.TextMatrix(GraphDataRow.��λ��־, intCount + .FixedCols), ";")(intCol + 1) <> "������" Then
                                        strChar = mstrBreath
                                    Else
                                        strChar = ""
                                    End If
                                    
                                    Select Case ConvertToValue(intCol, Y1)
                                    Case Is <= GetMinValue(intCol)
                                        Y1 = ConvertToY(intCol, GetMinValue(intCol))
                                        Call DrawLine(picGraph, X1, Y1, X1, Y1 + 200, lngColor, , , True)
                                    Case Is >= GetMaxValue(intCol)
                                        Y1 = ConvertToY(intCol, GetMaxValue(intCol))
                                        Call DrawLine(picGraph, X1, Y1, X1, Y1 - 200, lngColor, , , True)
                                    End Select
                                End Select
                                
                                '���һ�������ݲŲ�����;2.����м���˵����Ҫ����ߵ�
                                '45987,������,2012-09-10,��������
                                '1.������ʾΪ���ߣ�2������ͳһ�ú�ɫʵ�е㣨�񣩱�ʾ���������µ�����ʾ��������
                                '3. ʹ�ú������Ļ��ߣ������Ժ�R��ʾ������Ӧʱ���ں���30�κ����¶����úڱʻ�R�����ڵ�R֮���Լ�R����������֮�䲻����
                                If (DateDiff("d", CDate(strDate0), CDate(strDate1)) <= 1) Then
                                    If blnStop = False Then
                                        '���δ��˵����"����",�����ϸ���㻭������
                                        If intCol = mItemSerial.���� And Split(.TextMatrix(GraphDataRow.δ��˵��, intCount + .FixedCols), ";")(intCol + 1) = "����" Then
                                            'nothing to do
                                        ElseIf intCol = mItemSerial.���� And (strChar = "" Or strChar1 = "") Then
                                            'nothing to do
                                        Else
                                            DrawLine picGraph, X0, Y0, X1, Y1, .CellForeColor
                                        End If
                                    End If
                                    blnStop = False
                                End If
    
                            Else
                                Select Case intCol
                                Case mItemSerial.����
                                    
                                    '������
                                    lngColor = &HFF&
                                    strChar = "��"
                                    If Y1 < Y0 Then
                                    
                                        '������ʧ�ܣ�������ͷ�ĺ�ɫʵ�ߣ��ַ��̶��á�
                                        Call DrawLine(picGraph, X0, Y0, X1, Y1, lngColor, , , True)
                                        
                                    ElseIf Y1 > Y0 Then
                                    
                                        '�����³ɹ�������ɫ���ߣ��ַ��̶��á�
                                        Call DrawLine(picGraph, X0, Y0, X1, Y1, lngColor, 2)
                                        
                                    End If
                                    
                                Case mItemSerial.����
                                    If Y1 <> Y0 Then
                                        lngColor = &HFF&
                                        strChar = mstr���ʷ���

                                        Select Case ConvertToValue(intCol, Y1)
                                        Case Is <= GetMinValue(intCol)
                                            Y1 = ConvertToY(intCol, GetMinValue(intCol))
                                            Call DrawLine(picGraph, X1, Y1, X1, Y1 + 200, lngColor, , , True)
                                        Case Is >= GetMaxValue(intCol)
                                            Y1 = ConvertToY(intCol, GetMaxValue(intCol))
                                            Call DrawLine(picGraph, X1, Y1, X1, Y1 - 200, lngColor, , , True)
                                        End Select
                                        
                                        mpt����(intCount).X = X1
                                        mpt����(intCount).Y = Y1
                                        
                                    End If
                                Case mItemSerial.����

                                    Select Case ConvertToValue(intCol, Y1)
                                    Case Is <= GetMinValue(intCol)
                                        Y1 = ConvertToY(intCol, GetMinValue(intCol))
                                        Call DrawLine(picGraph, X1, Y1, X1, Y1 + 200, lngColor, , , True)
                                    Case Is >= GetMaxValue(intCol)
                                        Y1 = ConvertToY(intCol, GetMaxValue(intCol))
                                        Call DrawLine(picGraph, X1, Y1, X1, Y1 - 200, lngColor, , , True)
                                    End Select
                                    
                                    mpt����(intCount).X = X1
                                    mpt����(intCount).Y = Y1
                                
                                Case mItemSerial.����
                                    If Split(.TextMatrix(GraphDataRow.��λ��־, intCount + .FixedCols), ";")(intCol + 1) <> "������" Then
                                        strChar = mstrBreath
                                    Else
                                        strChar = ""
                                    End If
                                    
                                    Select Case ConvertToValue(intCol, Y1)
                                    Case Is <= GetMinValue(intCol)
                                        Y1 = ConvertToY(intCol, GetMinValue(intCol))
                                        Call DrawLine(picGraph, X1, Y1, X1, Y1 + 200, lngColor, , , True)
                                    Case Is >= GetMaxValue(intCol)
                                        Y1 = ConvertToY(intCol, GetMaxValue(intCol))
                                        Call DrawLine(picGraph, X1, Y1, X1, Y1 - 200, lngColor, , , True)
                                    End Select
                                End Select
                            End If
                        Else

                            Select Case intCol
                            Case mItemSerial.����
                                Select Case Split(.TextMatrix(GraphDataRow.��λ��־, intCount + .FixedCols), ";")(intCol + 1)
                                Case "����"
                                    strChar = mstrChar(0)
                                Case "Ҹ��"
                                    strChar = mstrChar(1)
                                Case "����"
                                    strChar = mstrChar(2)
                                Case Else
                                    strChar = mstrChar(1)
                                End Select

                                Select Case ConvertToValue(intCol, Y1)
                                Case Is <= GetMinValue(intCol)
                                    '����35������
                                    strChar = "��"
                                    Y1 = ConvertToY(intCol, GetMinValue(intCol))
                                    Call DrawLine(picGraph, X1, Y1, X1, Y1 + 200, lngColor, , , True)
                                    
                                Case Is >= GetMaxValue(intCol)
                                    '����42������
                                    strChar = "��"
                                    Y1 = ConvertToY(intCol, GetMaxValue(intCol))
                                    Call DrawLine(picGraph, X1, Y1, X1, Y1 - 200, lngColor, , , True)
                                    
                                End Select
                                
                                If Val(.TextMatrix(10, intCount + .FixedCols)) = 1 Then
                                    '���Ժϸ�
                                    Call DrawText(picGraph, X1 - 50, Y1 - 250, "v", lngColor)
                                End If
                                    
                            Case mItemSerial.����
                                If Split(.TextMatrix(GraphDataRow.��λ��־, intCount + .FixedCols), ";")(intCol + 1) = "" Then
                                    strChar = mstrPulse
                                Else
                                    strChar = ""
                                End If
                                
                                Select Case ConvertToValue(intCol, Y1)
                                Case Is <= GetMinValue(intCol)
                                    Y1 = ConvertToY(intCol, GetMinValue(intCol))
                                    Call DrawLine(picGraph, X1, Y1, X1, Y1 + 200, lngColor, , , True)
                                Case Is >= GetMaxValue(intCol)
                                    Y1 = ConvertToY(intCol, GetMaxValue(intCol))
                                    Call DrawLine(picGraph, X1, Y1, X1, Y1 - 200, lngColor, , , True)
                                End Select
                                
                                mpt����(intCount).X = X1
                                mpt����(intCount).Y = Y1
                                
                            Case mItemSerial.����
                                
                                Select Case ConvertToValue(intCol, Y1)
                                Case Is <= GetMinValue(intCol)
                                    Y1 = ConvertToY(intCol, GetMinValue(intCol))
                                    Call DrawLine(picGraph, X1, Y1, X1, Y1 + 200, lngColor, , , True)
                                Case Is >= GetMaxValue(intCol)
                                    Y1 = ConvertToY(intCol, GetMaxValue(intCol))
                                    Call DrawLine(picGraph, X1, Y1, X1, Y1 - 200, lngColor, , , True)
                                End Select
                                    
                                mpt����(intCount).X = X1
                                mpt����(intCount).Y = Y1
                                
                            Case mItemSerial.����
                                If Split(.TextMatrix(GraphDataRow.��λ��־, intCount + .FixedCols), ";")(intCol + 1) <> "������" Then
                                    strChar = mstrBreath
                                Else
                                    strChar = ""
                                End If
                                
                                Select Case ConvertToValue(intCol, Y1)
                                Case Is <= GetMinValue(intCol)
                                    Y1 = ConvertToY(intCol, GetMinValue(intCol))
                                    Call DrawLine(picGraph, X1, Y1, X1, Y1 + 200, lngColor, , , True)
                                Case Is >= GetMaxValue(intCol)
                                    Y1 = ConvertToY(intCol, GetMaxValue(intCol))
                                    Call DrawLine(picGraph, X1, Y1, X1, Y1 - 200, lngColor, , , True)
                                End Select
                                
                            End Select
                        End If
                        
                        If intCol = mItemSerial.���� Then
                            Call PointAdd(rsPoint, X1, Y1, mshScale.ColData(intCol), strChar, lngColor, intCount, mshScale.TextMatrix(4, intCount + .FixedCols))
                        ElseIf intCol = mItemSerial.���� Then
                            Call PointAdd(rsPoint, X1, Y1, mshScale.ColData(intCol), strChar, lngColor, intCount, mshScale.TextMatrix(4, intCount + .FixedCols), IIf(strChar = "", "BREATH", ""))
                        ElseIf intCol = mItemSerial.���� Then
                            Call PointAdd(rsPoint, X1, Y1, mshScale.ColData(intCol), strChar, lngColor, intCount, mshScale.TextMatrix(4, intCount + .FixedCols), IIf(strChar = "", "PACEMAKER", ""))
                        Else
                            Call PointAdd(rsPoint, X1, Y1, mshScale.ColData(intCol), strChar, lngColor, intCount, "")
                        End If
                                                   
                        '��¼�ϴλ���λ�ú�����
                        If X0 <> 0 And i <> 0 Then X1 = X0: Y1 = Y0 '�ӵ�һ������һ������
                        X0 = X1: Y0 = Y1: strDate0 = strDate1: strChar1 = strChar 'strChar1����Ŀǰֻ��Ժ���
                        
                        blnStop = False
                    Next i
                End If
                    
'                If blnStop = False Then
'                    If (.TextMatrix(GraphDataRow.�ϱ�˵��, intCount + .FixedCols) <> "" Or .TextMatrix(GraphDataRow.�±�˵��, intCount + .FixedCols) <> "") Then
'                        blnStop = (Val(.TextMatrix(GraphDataRow.�Ͽ���־, intCount + .FixedCols)) = 1)
'                    End If
'                End If
                
                If blnStop = False Then
                    
                    aryNote = Split(.TextMatrix(GraphDataRow.δ��˵��, intCount + .FixedCols), ";")
                    blnStop = (Trim(aryNote(intCol + 1)) <> "")
                    
                    '���±����ǿ��Բ�Ҫ�˵ģ������ǵ���ǰ������
                    If blnStop = False Then
                        If (.TextMatrix(GraphDataRow.�ϱ�˵��, intCount + .FixedCols) <> "" Or .TextMatrix(GraphDataRow.�±�˵��, intCount + .FixedCols) <> "") Then
                            blnStop = (Val(.TextMatrix(GraphDataRow.�Ͽ���־, intCount + .FixedCols)) = 1)
                        End If
                    End If
                    
                End If
                
            Next intCount
            
        Next intCol
        
        
        '������ַ���ͼ��
        '--------------------------------------------------------------------------------------------------------------
        Call DrawPoint(picGraph, rsPoint, mItemSerial.����)
        
        '�������������������γɶ���Σ����������ߺ����
        '--------------------------------------------------------------------------------------------------------------
        Call DrawPoly(picGraph, mpt����, mpt����)

        Dim lngYMax As Long
        If mItemSerial.���� <> -1 Then
            lngYMax = ConvertToY(mItemSerial.����, 33.4)
        Else
            lngYMax = 8580
        End If
        
        '��ӡ���ת��־
        '--------------------------------------------------------------------------------------------------------------
        Dim intLoop As Integer
        Dim rsTmp As ADODB.Recordset
        
        '20090926:������40-42�ȼ��ӡ,����һ����Ϣ�����������С����,�ж�����Ϣ���Ӻ���һ���ӡ,��������һ���ֱ��ȫ����ӡ
        Set rsTmp = New ADODB.Recordset
        rsTmp.Fields.Append "�к�", adDouble, 30
        rsTmp.Fields.Append "ʱ��", adVarChar, 30
        rsTmp.Fields.Append "���", adVarChar, 50
        '20090926--
        rsTmp.Fields.Append "����", adVarChar, 50       '��¼�����ת,������Ժ,����δ��˵��,�ϱ�˵��
        rsTmp.Fields.Append "��ӡ��", adVarChar, 30
        rsTmp.Fields.Append "����", adVarChar, 30
        rsTmp.Fields.Append "�߶�", adVarChar, 30       'δ��˵�����ϱ�˵�����ùܸ߶�
        rsTmp.Fields.Append "�����С", adVarChar, 50
        '----------
        rsTmp.Open

        Dim intCharNumber As Integer
        
        For intCol = 0 To .Cols - .FixedCols - 1
            
            X1 = HOUR_STEP_Twips * intCol + HOUR_STEP_Twips / 2
'            Y1 = ConvertToY(mItemSerial.����, 42)
            Y1 = 195
            If mItemSerial.���� <> -1 Then
                dblHeight = ConvertToY(mItemSerial.����, 40) - Y1
            Else
                dblHeight = 2145 - Y1
            End If
            
            '�к�:=3��ʾ����;=5��ʾ��Ժ;=6��ʾת��;=7��ʾ����;=8��ʾ��Ժ,=13����
            rsTmp.Filter = ""
            For intLoop = 5 To 9
                If .TextMatrix(intLoop, intCol + .FixedCols) <> "" Then
                    rsTmp.AddNew
                    rsTmp.Fields("����").Value = intLoop
                    rsTmp.Fields("����").Value = X1 & ";" & Y1
                    rsTmp.Fields("�к�").Value = intCol
                    rsTmp.Fields("ʱ��").Value = .Cell(flexcpData, intLoop, intCol + .FixedCols, intLoop, intCol + .FixedCols)
                    rsTmp.Fields("���").Value = .TextMatrix(intLoop, intCol + .FixedCols)
                End If
            Next
            If .TextMatrix(������־, intCol + .FixedCols) <> "" Then
                rsTmp.AddNew
                rsTmp.Fields("����").Value = ������־
                rsTmp.Fields("����").Value = X1 & ";" & Y1
                rsTmp.Fields("�к�").Value = intCol
                rsTmp.Fields("ʱ��").Value = .Cell(flexcpData, ������־, intCol + .FixedCols, ������־, intCol + .FixedCols)
                rsTmp.Fields("���").Value = .TextMatrix(������־, intCol + .FixedCols)
            End If
            If .TextMatrix(������־, intCol + .FixedCols) <> "" Then
                rsTmp.AddNew
                rsTmp.Fields("����").Value = ������־
                rsTmp.Fields("����").Value = X1 & ";" & Y1
                rsTmp.Fields("�к�").Value = intCol
                rsTmp.Fields("ʱ��").Value = .Cell(flexcpData, ������־, intCol + .FixedCols, ������־, intCol + .FixedCols)
                rsTmp.Fields("���").Value = .TextMatrix(������־, intCol + .FixedCols)
            End If
            

            'δ��˵��
            '----------------------------------------------------------------------------------------------------------
            If bytδ����ʾλ�� = 0 Then
                strComment = ""
                strTmp = ""
                varNote = Split(.TextMatrix(GraphDataRow.δ��˵��, intCol + .FixedCols), ";")
                For intCount = 0 To UBound(varNote)
                    If varNote(intCount) <> "����" Then
                        If InStr(";" & strTmp & ";", ";" & varNote(intCount) & ";") = 0 Then
                            strTmp = strTmp & ";" & varNote(intCount)
                        End If
                    End If
                Next
                If strTmp <> "" Then
                    varNote = Split(strTmp, ";")
                    For intCount = 0 To UBound(varNote)
                        If strComment = "" Or strComment = " " Then
                            strComment = strComment & varNote(intCount)
                        Else
                            strComment = strComment & " " & varNote(intCount)
                        End If
                    Next
                End If
                If strComment <> "" Then
                    rsTmp.AddNew
                    rsTmp.Fields("����").Value = δ��˵��
                    rsTmp.Fields("����").Value = X1 & ";" & Y1
                    rsTmp.Fields("�к�").Value = intCol
                    rsTmp.Fields("ʱ��").Value = .Cell(flexcpData, δ��˵��, intCol + .FixedCols, δ��˵��, intCol + .FixedCols)
                    rsTmp.Fields("���").Value = strComment
                End If
            End If
            
            '�ϱ�˵��
            '----------------------------------------------------------------------------------------------------------
            strComment = Trim(.TextMatrix(GraphDataRow.�ϱ�˵��, intCol + .FixedCols))
            If strComment <> "" Then
                rsTmp.AddNew
                rsTmp.Fields("����").Value = �ϱ�˵��
                rsTmp.Fields("����").Value = X1 & ";" & Y1
                rsTmp.Fields("�к�").Value = intCol
                rsTmp.Fields("ʱ��").Value = .Cell(flexcpData, �ϱ�˵��, intCol + .FixedCols, �ϱ�˵��, intCol + .FixedCols)
                rsTmp.Fields("���").Value = strComment
            End If
            
            '�±�˵��
            '----------------------------------------------------------------------------------------------------------
            
'            Y1 = ConvertToY(mItemSerial.����, 35)
            Y1 = 7020
            strComment = ""
            
            'δ��˵��
            '----------------------------------------------------------------------------------------------------------
            If bytδ����ʾλ�� = 1 Then
                strComment = ""
                strTmp = ""
                varNote = Split(.TextMatrix(GraphDataRow.δ��˵��, intCol + .FixedCols), ";")
                For intCount = 0 To UBound(varNote)
                    If varNote(intCount) <> "����" Then
                        If InStr(";" & strTmp & ";", ";" & varNote(intCount) & ";") = 0 Then
                            strTmp = strTmp & ";" & varNote(intCount)
                        End If
                    End If
                Next
                If strTmp <> "" Then
                    varNote = Split(strTmp, ";")
                    For intCount = 0 To UBound(varNote)
                        If strComment = "" Or strComment = " " Then
                            strComment = strComment & varNote(intCount)
                        Else
                            strComment = strComment & " " & varNote(intCount)
                        End If
                    Next
                End If

                If Trim(strComment) <> "" Then
                    
                    intCharNumber = 0
                    For intCount = 1 To Len(strComment)
                        If Y1 <= lngYMax Then
                            strChar = Mid(strComment, intCount, 1)
                            '��ɫ
                            
                            If Asc(strChar) < 0 Then
                                If intCharNumber Mod 2 = 1 Then Y1 = Y1 + ROWHEIGHT * 2.5
                            End If
                            
                            Call DrawRotateText(picGraph, X1 - picGraph.TextWidth(strChar) / 2, Y1 + 15, strChar, -2147483635)
                            
                            If Asc(strChar) < 0 Then
                                intCharNumber = 0
                                Y1 = Y1 + ROWHEIGHT * 5
                            Else
                                Y1 = Y1 + ROWHEIGHT * 2.5
                                intCharNumber = intCharNumber + 1
                            End If
                        End If
                    Next
                End If
            End If

            strComment = IIf(Trim(strComment) = "", "", " ") & .TextMatrix(GraphDataRow.�±�˵��, intCol + .FixedCols)
            If Trim(strComment) <> "" Then
                intCharNumber = 0
                For intCount = 1 To Len(strComment)
                    If Y1 <= lngYMax Then
                        strChar = Mid(strComment, intCount, 1)
                        '��ɫ
                        
                        If Asc(strChar) < 0 Then
                            If intCharNumber Mod 2 = 1 Then Y1 = Y1 + ROWHEIGHT * 2.5
                        End If
                        
                        Call DrawRotateText(picGraph, X1 - picGraph.TextWidth(strChar) / 2, Y1 + 15, strChar, -2147483635)
                        
                        If Asc(strChar) < 0 Then
                            intCharNumber = 0
                            Y1 = Y1 + ROWHEIGHT * 5
                        Else
                            Y1 = Y1 + ROWHEIGHT * 2.5
                            intCharNumber = intCharNumber + 1
                        End If
                    End If
                Next
            End If
            
        Next

    End With
    
    Call OutputNote(picGraph, dblHeight, rsTmp)
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Function

Private Function GetFontSize(ByVal objDraw As Object, ByVal dblHeight As Double, ByVal strText As String, ByRef Y1 As Single) As Single
    Dim sinFontSize As Single
    Dim sinFontSize_Bak As Single
    Dim intCharNumber As Integer
    Dim intCount As Integer
    Dim strChar As String
    '�������������С
    
    sinFontSize_Bak = objDraw.FontSize
    For sinFontSize = objDraw.FontSize To 5 Step -1
        Y1 = 0
        intCharNumber = 0
        For intCount = 1 To Len(strText)
            strChar = Mid(strText, intCount, 1)
            
            If Asc(strChar) < 0 Then
                If intCharNumber Mod 2 = 1 Then Y1 = Y1 + ROWHEIGHT * 2.5
            End If
            
            If Asc(strChar) < 0 Then
                intCharNumber = 0
                Y1 = Y1 + ROWHEIGHT * 5
            Else
                Y1 = Y1 + ROWHEIGHT * 2.5
                intCharNumber = intCharNumber + 1
            End If
        Next
        'If Y1 <= dblHeight Then Exit For
        Exit For
    Next
    
    objDraw.FontSize = sinFontSize_Bak
    GetFontSize = sinFontSize
End Function

Private Sub OutputNote(ByVal objDraw As Object, ByVal dblHeight As Double, ByRef rsNote As ADODB.Recordset)
    '���������Ϣ:��Ժ,���,ת��,��Ժ,��������,δ��˵��,�ϱ�˵��������
    'δ��˵�����ϱ�˵��,��û�����ת�������估��������Ϣʱ,��ӡ��42-40֮��;�����40��ʼ���´�ӡ
    '��δ��˵�����ϱ�˵����,���ת����Ϣ��һ���̶ȷ������ʱ,����д������̶���,�������̶�Ҳ����Ϣ,˳��
    Dim intCol As Integer                   '��¼��ǰ�к�
    Dim intMax As Integer                   '������
    Dim intCur As Integer                   '��ǰ��¼��λ��
    Dim bln�ϱ� As Boolean
    Dim sinX1 As Single, sinY1 As Single, sinHeight As Single, sinMaxY1 As Single
    Dim rsTarget As New ADODB.Recordset
    
    '����ַ���ر�������
    Dim sinFontSize As Single
    Dim sinFontSize_Bak As Single
    Dim intCharNumber As Integer
    Dim intCount As Integer
    Dim strChar As String
    
    intMax = mshScale.Cols - mshScale.FixedCols - 1
    sinFontSize_Bak = objDraw.FontSize
    Set rsTarget = rsNote.Clone
    With rsNote
        If .RecordCount = 0 Then Exit Sub
        .Sort = "�к�,ʱ��"
        intCol = !�к�
        
        '�������ת��������ѭ��
        Do While Not .EOF
            If Trim(NVL(!���)) <> "" Then
                If Not (!���� = δ��˵�� Or !���� = �ϱ�˵��) Then
                    '������ӡ���Ƿ��Ѵ������,���������У������
                    If intCol > intMax Then intCol = intMax
                    
                    '����õ����ʵ������С���߶�
                    !�����С = GetFontSize(objDraw, dblHeight, NVL(!���), sinY1)
                    !�߶� = sinY1
                    !��ӡ�� = IIf(intCol < !�к�, !�к�, intCol)
                    .Update
                    If intCol <= !�к� Then intCol = !�к�
                    intCol = intCol + 1
                Else
                    Call GetFontSize(objDraw, dblHeight, NVL(!���), sinY1)
                    !�߶� = sinY1
                    .Update
                End If
            End If
            
            .MoveNext
        Loop
        .MoveFirst
        
        '�������ת�ȵ�������(ֻ�����һ�вŴ���һ���������)
        sinY1 = 195
        .Filter = "��ӡ��='" & intMax & "'"
        .Sort = "�к�,ʱ��"
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            'ֻ�����ת�����Ÿ����˴�ӡ��
            !���� = Split(!����, ";")(0) & ";" & sinY1
            .Update
            sinY1 = sinY1 + !�߶� + 100
            
            .MoveNext
        Loop
        .Filter = 0
        .MoveFirst
        
        '����У��δ��˵���Լ��ϱ�˵���ĸ߶�(δ��˵�����ϱ�˵��,��û�����ת�������估��������Ϣʱ,��ӡ��42-40֮��;�����40��ʼ���´�ӡ)
        Set rsTarget = .Clone
        intCol = 0
        Do While Not .EOF
            If (!���� = δ��˵�� Or !���� = �ϱ�˵��) Then
                bln�ϱ� = False
                Set rsTarget = .Clone
                rsTarget.Filter = "��ӡ��='" & !�к� & "'"
                If rsTarget.RecordCount <> 0 Then
                    '�Ѵ��ڴ�ӡ���ݵĲ�У��������
                    sinMaxY1 = Split(rsTarget!����, ";")(1)
                    Do While Not rsTarget.EOF
                        If bln�ϱ� = False Then
                            '���ǵ��ϱ��п�������40�ȿ�ʼ���,������У��һ��sinMaxY1������
                            bln�ϱ� = (rsTarget!���� = δ��˵�� Or rsTarget!���� = �ϱ�˵��)
                            If bln�ϱ� Then sinMaxY1 = Split(rsTarget!����, ";")(1)
                        End If
                        sinMaxY1 = sinMaxY1 + rsTarget!�߶� + 100
                        rsTarget.MoveNext
                    Loop
                    If mItemSerial.���� <> -1 Then
                        sinY1 = ConvertToY(mItemSerial.����, 40)
                    Else
                        sinY1 = 2145
                    End If
                    If sinY1 < sinMaxY1 Or bln�ϱ� Then sinY1 = sinMaxY1
                    sinHeight = !�߶�
                    intCol = !�к�
                Else
                    sinY1 = 195
                    intCol = !�к�
                    sinHeight = !�߶�
                End If
                rsTarget.Filter = 0
                
                !���� = Split(!����, ";")(0) & ";" & sinY1
                !��ӡ�� = !�к�                                 '��ʱ���´�ӡ��,�Ա������ѭ������
                .Update
            End If
            .MoveNext
        Loop
    
        '��ʼ�������������
        .MoveFirst
        Do While Not .EOF
            If Trim(NVL(!���)) <> "" Then
                'If (!���� = δ��˵�� Or !���� = �ϱ�˵��) Then Stop
                sinX1 = HOUR_STEP_Twips * (IIf(!��ӡ�� = "", Val(!�к�), Val(!��ӡ��))) + HOUR_STEP_Twips / 2
                sinY1 = Split(!����, ";")(1)
                intCharNumber = 0
                objDraw.FontSize = IIf(!�����С = "", 9, !�����С)
                
                For intCount = 1 To Len(!���)
                    strChar = Mid(!���, intCount, 1)
                    
                    If Asc(strChar) < 0 Then
                        If intCharNumber Mod 2 = 1 Then sinY1 = sinY1 + ROWHEIGHT * 2.5
                    End If
                    Call DrawRotateText(objDraw, sinX1 - objDraw.TextWidth(strChar) / 2, sinY1 + 15, strChar, IIf(!���� = δ��˵�� Or !���� = �ϱ�˵��, -2147483635, 255))
                    If Asc(strChar) < 0 Then
                        intCharNumber = 0
                        sinY1 = sinY1 + ROWHEIGHT * 5
                    Else
                        sinY1 = sinY1 + ROWHEIGHT * 2.5
                        intCharNumber = intCharNumber + 1
                    End If
                Next
            End If
            
            .MoveNext
        Loop
    End With
    objDraw.FontSize = sinFontSize_Bak
End Sub

Private Function isSaved() As String
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ� �ж����±��Ƿ񱣴棬���δ���淵����ʾ��Ϣ���Ѿ����淵���㳤���ַ���
    '������
    '���أ�
    '------------------------------------------------------------------------------------------------------------------
    Dim aryValue() As String
    Dim strItem As String
    On Error Resume Next
    
    With mshUpTab
        For intCol = .FixedCols To .Cols - 1
            If .ColData(intCol) = 2 Or .ColData(intCol) = 3 Then
                isSaved = "�ı��ˡ������ա���δ���档"
                Exit Function
            End If
        Next
    End With
    With mshScale
        For intCol = .FixedCols To .Cols - 1
            aryValue = Split(.TextMatrix(0, intCol), ";")
            For intCount = 0 To UBound(aryValue)
                If intCount = 0 Then
                    strItem = "˵��"
                Else
                    strItem = .TextMatrix(0, intCount - 1)
                End If
                If aryValue(intCount) = "2" Or aryValue(intCount) = "3" Or aryValue(intCount) = "4" Then
                    isSaved = "�ı��ˡ�" & strItem & "��������δ���档"
                    Exit Function
                End If
            Next
        Next
    End With
    With mshDownTab
        For intCol = .FixedCols To .Cols - 1
            aryValue = Split(.TextMatrix(0, intCol), ";")
            For intCount = 0 To UBound(aryValue)
                strItem = .TextMatrix(intCount + 1, 1)
                If aryValue(intCount) = "2" Or aryValue(intCount) = "3" Or aryValue(intCount) = "4" Then
                    isSaved = "�ı��ˡ�" & strItem & "��������δ����"
                    Exit Function
                End If
            Next
        Next
    End With
End Function


Private Function CalcMinMaxCol(ByVal strDate As String, MinCol As Long, MaxCol As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ� �����С���ʱ�䷶Χ
    '������
    '���أ�
    '------------------------------------------------------------------------------------------------------------------
    Dim aryValue() As String
    Dim dtTmp As Date
    Dim strTmp As String
    
    If mvarEdit = False Then Exit Function
    
    aryValue = Split(strDate, ";")
    
    MinCol = GetCurveColumn(CDate(aryValue(0)), CDate(aryValue(0)), mlngHourBegin) - 1
    MaxCol = GetCurveColumn(CDate(aryValue(1)), CDate(aryValue(0)), mlngHourBegin) - 1
    
End Function

Private Function SetColBkColor(ByVal Col As Long, ByVal COLOR As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ�
    '������
    '���أ�
    '------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    
    mshScale.Redraw = False
    mshScale.Col = Col
    For i = 0 To mshScale.Rows - 1
        mshScale.Row = i
        mshScale.CellBackColor = COLOR
    Next
    mshScale.Redraw = True
End Function

Private Function SetVisible() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ�
    '������
    '���أ�
    '------------------------------------------------------------------------------------------------------------------
    If Val(mrsParam("�༭")) = 0 Then
        mshUpTab.Enabled = False
        mshDownTab.Enabled = False
        picBack.Enabled = False
        picScale.Enabled = False
    Else
        mshUpTab.Enabled = True
        mshDownTab.Enabled = True
        picBack.Enabled = True
        picScale.Enabled = True
    End If
End Function

Private Function SetBodyMode() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ� �������µ�����ʾģʽ���Ǳ༭ģʽ
    '������
    '���أ�
    '------------------------------------------------------------------------------------------------------------------
    If Val(mrsParam("�༭")) = 0 Then

        mshDownTab.Enabled = False
        mshUpTab.Enabled = False
    Else

        mshDownTab.Enabled = True
        mshUpTab.Enabled = True
    End If
    
End Function

Private Function SaveData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '1������������������
    '2����������ͼ������
    '3���������±������
    '------------------------------------------------------------------------------------------------------------------
    Dim lngKey As Long, i As Long
    Dim aryValue() As String, aryData() As String
    Dim aryMakeTime() As String
    Dim strFrom As String, strTo As String
    Dim strItem As String, strTime As String
    Dim strMakeTime As String
    Dim mvarStrValue As String, dblValues As Double
    Dim lngItemCode As Long, intMode As Integer
    Dim rs As New ADODB.Recordset
    Dim intLoop As Long
    Dim strStart As String
    Dim strEnd As String
    Dim strValues As String
    Dim intTmp As Integer, intMax As Integer
    Dim lng�����ļ�id As Long
    Dim lng��������id As Long
    Dim strSQL() As String
    Dim strTmp As String
    Dim strɾ�����ʶ�׾ As String           '��֤һ��ֻɾ��һ��
    Dim blnHistoryData As Boolean           '��ʷ���ݱ������ʱ��,����ʷʱ��Ϊ׼��������
    Dim blnTrans As Boolean
    
    If Val(mrsParam("�༭")) = 0 Then Exit Function
        
    strFrom = Split(picScale.Tag, ";")(0)
    strTo = Split(picScale.Tag, ";")(1)
    Screen.MousePointer = 11
    
    ReDim Preserve strSQL(1 To 1)
    
    On Error GoTo ErrHead

    '1.����������������
    With mshUpTab
        'mshDownTab�ǰ�1Ϊ��־�����жϣ�mshUpTab�ǰ�"ɾ��������"��"��д������"�Ƿ�����˱༭
        If .Tag = "��д������" Or .Tag = "ɾ��������" Then
            For intCol = .FixedCols To .Cols - 1
                
                For intLoop = (intCol - 1) * 6 + mshScale.FixedCols To intCol * 6 + mshScale.FixedCols
                    
                    strTmp = GetCurveDateTime(intLoop - mshScale.FixedCols + 1, CDate(strFrom), mlngHourBegin)
                    strStart = Split(strTmp, ",")(0)
                    strEnd = Split(strTmp, ",")(1)

                    If Int(CDate(strStart)) < Int(CDate(strFrom)) Then
                        strStart = Format(strFrom, "yyyy-MM-dd HH:mm:ss")
                    End If
                    
                    mstrSQL = "ZL_���ӻ����¼_UPDATE("
                    mstrSQL = mstrSQL & Val(mrsParam("����id")) & ","
                    mstrSQL = mstrSQL & Val(mrsParam("��ҳid")) & ","
                    mstrSQL = mstrSQL & Val(mrsParam("Ӥ��").Value) & ","
                    mstrSQL = mstrSQL & "To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'),"
                    mstrSQL = mstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
                    mstrSQL = mstrSQL & "4,"
                    mstrSQL = mstrSQL & "0,"
                    mstrSQL = mstrSQL & "0,"
                    mstrSQL = mstrSQL & "NULL"
                    mstrSQL = mstrSQL & ")"
                    strSQL(ReDimArray(strSQL)) = mstrSQL
                Next

                strStart = mstrOpsDays(intCol)
                strEnd = mstrOpsDays(intCol)
                If strStart <> "" Then
                
                    strTmp = ""
                    
                    intTmp = GetCurveColumn(CDate(strStart), CDate(strFrom), mlngHourBegin) + mshScale.FixedCols - 1
                    If Left(mshScale.TextMatrix(3, intTmp), 4) = "��������" Then
                        strTmp = "��������"
                    ElseIf Left(mshScale.TextMatrix(3, intTmp), 2) = "����" Then
                        strTmp = "����"
                    ElseIf Left(mshScale.TextMatrix(3, intTmp), 2) = "����" Then
                        strTmp = "����"
                    End If
                    
                    mstrSQL = "ZL_���ӻ����¼_UPDATE("
                    mstrSQL = mstrSQL & Val(mrsParam("����id")) & ","
                    mstrSQL = mstrSQL & Val(mrsParam("��ҳid")) & ","
                    mstrSQL = mstrSQL & Val(mrsParam("Ӥ��").Value) & ","
                    mstrSQL = mstrSQL & "To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'),"
                    mstrSQL = mstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
                    mstrSQL = mstrSQL & "4,"
                    mstrSQL = mstrSQL & "0,"
                    mstrSQL = mstrSQL & "0,"
                    mstrSQL = mstrSQL & "'" & strTmp & "'"
                    mstrSQL = mstrSQL & ")"

                    strSQL(ReDimArray(strSQL)) = mstrSQL
                End If

            Next
        End If
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    '2.��������ͼ������
    With mshScale
    
        'ע��˵������
        For intCol = .FixedCols To .Cols - 1
            strTmp = GetCurveDateTime(intCol - mshScale.FixedCols + 1, CDate(strFrom), mlngHourBegin)
            strTime = Split(strTmp, ",")(0)
            strEnd = Split(strTmp, ",")(1)
            
            If Int(CDate(strTime)) < Int(CDate(strFrom)) Then strTime = Format(strFrom, "yyyy-MM-dd HH:mm:ss")

            aryValue = Split(.TextMatrix(GraphDataRow.���ı�־, intCol), ";")
                            
                            

            If Mid(Format(strTime, "yyyy-MM-dd HH:mm"), 12, 5) = "00:00" Then
                strMakeTime = Format(DateAdd("h", -2, CDate(strEnd)), "yyyy-MM-dd HH:mm")
                strMakeTime = Format(DateAdd("n", 1, CDate(strMakeTime)), "yyyy-MM-dd HH:mm:ss")
            Else
                strMakeTime = Format(DateAdd("h", 2, CDate(strTime)), "yyyy-MM-dd HH:mm:ss")
            End If
            strMakeTime = "To_Date('" & strMakeTime & "','yyyy-mm-dd hh24:mi:ss')"
                
            mstrSQL = "ZL_���ӻ����¼_UPDATE("
            mstrSQL = mstrSQL & Val(mrsParam("����id")) & ","
            mstrSQL = mstrSQL & Val(mrsParam("��ҳid")) & ","
            mstrSQL = mstrSQL & Val(mrsParam("Ӥ��")) & ","
            mstrSQL = mstrSQL & "To_Date('" & strTime & "','yyyy-mm-dd hh24:mi:ss'),"
            mstrSQL = mstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
            mstrSQL = mstrSQL & "2,"
            mstrSQL = mstrSQL & "0,"
            mstrSQL = mstrSQL & Val(.TextMatrix(GraphDataRow.�Ͽ���־, intCol)) & ","
            mstrSQL = mstrSQL & IIf(Val(aryValue(0)) = 4, "NULL", "'" & .TextMatrix(GraphDataRow.�ϱ�˵��, intCol) & "'")
            mstrSQL = mstrSQL & ",Null,1,1,0,0," & strMakeTime & ",Null"
            mstrSQL = mstrSQL & ")"
            strSQL(ReDimArray(strSQL)) = mstrSQL

            mstrSQL = "ZL_���ӻ����¼_UPDATE("
            mstrSQL = mstrSQL & Val(mrsParam("����id")) & ","
            mstrSQL = mstrSQL & Val(mrsParam("��ҳid")) & ","
            mstrSQL = mstrSQL & Val(mrsParam("Ӥ��")) & ","
            mstrSQL = mstrSQL & "To_Date('" & strTime & "','yyyy-mm-dd hh24:mi:ss'),"
            mstrSQL = mstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
            mstrSQL = mstrSQL & "6,"
            mstrSQL = mstrSQL & "0,"
            mstrSQL = mstrSQL & Val(.TextMatrix(GraphDataRow.�Ͽ���־, intCol)) & ","
            mstrSQL = mstrSQL & IIf(Val(aryValue(0)) = 4, "NULL", "'" & .TextMatrix(GraphDataRow.�±�˵��, intCol) & "'")
            mstrSQL = mstrSQL & ",Null,1,1,0,0," & strMakeTime & ",Null"
            mstrSQL = mstrSQL & ")"
            
            strSQL(ReDimArray(strSQL)) = mstrSQL
            
        Next
        
        '��Ŀ��ֵ����
        For intCount = 0 To .FixedCols - 1
            Dim dbl���Ӳ� As Double                                         '���浱ǰ��ķ��Ӳ�
            '��ȡָ����Ŀ���壺���ֵ����Сֵ����λֵ�������
            strItem = .TextMatrix(0, intCount)              '��ȡ����
            aryValue = Split(picLine(intCount).Tag, ";")    '��ȡ��Ŀ������
            lngItemCode = mshScale.ColData(intCount)        '��ȡ��Ŀ���
            
            '���������ж�������
            For intCol = .FixedCols To .Cols - 1
                
                '�õ�ʱ��,���û��ʱ���,�Լ���������м�ʱ��Ϊ׼(������);��������ʷʱ��Ϊ׼�������ݸ���
                strTime = Split(mshScale.TextMatrix(GraphDataRow.����ʱ��, intCol), ";")(intCount + 1)
                If strTime = "" Then
                    blnHistoryData = False
                    strTmp = GetCurveDateTime(intCol - mshScale.FixedCols + 1, CDate(strFrom), mlngHourBegin)
                    strTime = Split(strTmp, ",")(0)
                    strEnd = Split(strTmp, ",")(1)
                    If Int(CDate(strTime)) < Int(CDate(strFrom)) Then strTime = Format(strFrom, "yyyy-MM-dd HH:mm:ss")
                Else
                    strEnd = strTime
                    strMakeTime = strTime
                    blnHistoryData = True
                End If
                
                intMode = Val(Split(.TextMatrix(GraphDataRow.���ı�־, intCol), ";")(intCount + 1))
                
                If intMode = OperateType.ɾ������ Or intMode = OperateType.�޸Ĳ��� Or intMode = OperateType.�������� Then
                    'ȡÿ����м�ʱ���
                    If blnHistoryData = False Then
                        dbl���Ӳ� = DateDiff("n", CDate(Split(strTmp, ",")(0)), CDate(Split(strTmp, ",")(1)))
                        dbl���Ӳ� = dbl���Ӳ� \ 2
                        strMakeTime = DateAdd("n", dbl���Ӳ�, CDate(Split(strTmp, ",")(0)))
                        If strMakeTime < mstrEnterDate Then strMakeTime = mstrEnterDate
                    End If
                    strMakeTime = "To_Date('" & strMakeTime & "','yyyy-mm-dd hh24:mi:ss')"
                End If
                
                If intMode = OperateType.ɾ������ Or intMode = OperateType.�޸Ĳ��� Then

                    'ɾ�������µ�����
                    mstrSQL = "ZL_���ӻ����¼_UPDATE("
                    mstrSQL = mstrSQL & Val(mrsParam("����id")) & ","
                    mstrSQL = mstrSQL & Val(mrsParam("��ҳid")) & ","
                    mstrSQL = mstrSQL & Val(mrsParam("Ӥ��")) & ","
                    mstrSQL = mstrSQL & "To_Date('" & strTime & "','yyyy-mm-dd hh24:mi:ss'),"
                    mstrSQL = mstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
                    mstrSQL = mstrSQL & "1,"
                    mstrSQL = mstrSQL & lngItemCode & ","
                    mstrSQL = mstrSQL & "1,"
                    mstrSQL = mstrSQL & "Null,Null,1,1,0,0," & strMakeTime
                    mstrSQL = mstrSQL & ")"
                    strSQL(ReDimArray(strSQL)) = mstrSQL

                    'ɾ��������穵�����
                    If mint����Ӧ�� = 2 And InStr(1, strɾ�����ʶ�׾ & ",", "," & intCol & ",") = 0 Then
                        mstrSQL = "ZL_���ӻ����¼_UPDATE("
                        mstrSQL = mstrSQL & Val(mrsParam("����id")) & ","
                        mstrSQL = mstrSQL & Val(mrsParam("��ҳid")) & ","
                        mstrSQL = mstrSQL & Val(mrsParam("Ӥ��")) & ","
                        mstrSQL = mstrSQL & "To_Date('" & strTime & "','yyyy-mm-dd hh24:mi:ss'),"
                        mstrSQL = mstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
                        mstrSQL = mstrSQL & "1,"
                        mstrSQL = mstrSQL & mItemNo.���� & ","
                        mstrSQL = mstrSQL & "1,"
                        mstrSQL = mstrSQL & "Null,Null,1,1,0,0," & strMakeTime
                        mstrSQL = mstrSQL & ")"
                        strSQL(ReDimArray(strSQL)) = mstrSQL
                        strɾ�����ʶ�׾ = strɾ�����ʶ�׾ & "," & intCol
                    End If
                    
                    mstrSQL = "ZL_���ӻ����¼_UPDATE("
                    mstrSQL = mstrSQL & Val(mrsParam("����id")) & ","
                    mstrSQL = mstrSQL & Val(mrsParam("��ҳid")) & ","
                    mstrSQL = mstrSQL & Val(mrsParam("Ӥ��")) & ","
                    mstrSQL = mstrSQL & "To_Date('" & strTime & "','yyyy-mm-dd hh24:mi:ss'),"
                    mstrSQL = mstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
                    mstrSQL = mstrSQL & "1,"
                    mstrSQL = mstrSQL & lngItemCode & ","
                    mstrSQL = mstrSQL & "0,"
                    mstrSQL = mstrSQL & "Null,"
                    mstrSQL = mstrSQL & IIf(lngItemCode = mItemNo.���� Or lngItemCode = mItemNo.���� Or lngItemCode = mItemNo.����, "'" & Split(.TextMatrix(GraphDataRow.��λ��־, intCol), ";")(intCount + 1) & "'", "''")
                    mstrSQL = mstrSQL & ",1,1,0,0," & strMakeTime
                    mstrSQL = mstrSQL & ",'" & Trim(Split(.TextMatrix(GraphDataRow.δ��˵��, intCol), ";")(intCount + 1)) & "')"
                    
                    strSQL(ReDimArray(strSQL)) = mstrSQL
                End If
                
                If intMode = OperateType.�������� Or intMode = OperateType.�޸Ĳ��� Then
                    
                    'ɾ�������µ�����
                    mstrSQL = "ZL_���ӻ����¼_UPDATE("
                    mstrSQL = mstrSQL & Val(mrsParam("����id")) & ","
                    mstrSQL = mstrSQL & Val(mrsParam("��ҳid")) & ","
                    mstrSQL = mstrSQL & Val(mrsParam("Ӥ��")) & ","
                    mstrSQL = mstrSQL & "To_Date('" & strTime & "','yyyy-mm-dd hh24:mi:ss'),"
                    mstrSQL = mstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
                    mstrSQL = mstrSQL & "1,"
                    mstrSQL = mstrSQL & lngItemCode & ","
                    mstrSQL = mstrSQL & "1,"
                    mstrSQL = mstrSQL & "Null,Null,1,1,0,0," & strMakeTime
                    mstrSQL = mstrSQL & ")"
                    strSQL(ReDimArray(strSQL)) = mstrSQL

                    'ɾ��������穵�����
                    If mint����Ӧ�� = 2 And InStr(1, strɾ�����ʶ�׾ & ",", "," & intCol & ",") = 0 Then
                        mstrSQL = "ZL_���ӻ����¼_UPDATE("
                        mstrSQL = mstrSQL & Val(mrsParam("����id")) & ","
                        mstrSQL = mstrSQL & Val(mrsParam("��ҳid")) & ","
                        mstrSQL = mstrSQL & Val(mrsParam("Ӥ��")) & ","
                        mstrSQL = mstrSQL & "To_Date('" & strTime & "','yyyy-mm-dd hh24:mi:ss'),"
                        mstrSQL = mstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
                        mstrSQL = mstrSQL & "1,"
                        mstrSQL = mstrSQL & mItemNo.���� & ","
                        mstrSQL = mstrSQL & "1,"
                        mstrSQL = mstrSQL & "Null,Null,1,1,0,0," & strMakeTime
                        mstrSQL = mstrSQL & ")"
                        strSQL(ReDimArray(strSQL)) = mstrSQL
                        strɾ�����ʶ�׾ = strɾ�����ʶ�׾ & "," & intCol
                    End If
                    
                    strValues = ""
                    aryData = Split(Split(.TextMatrix(GraphDataRow.��������, intCol), ";")(intCount + 1), ",")
                    If UBound(aryData) = -1 Then
                    
                            mstrSQL = "ZL_���ӻ����¼_UPDATE("
                            mstrSQL = mstrSQL & Val(mrsParam("����id")) & ","
                            mstrSQL = mstrSQL & Val(mrsParam("��ҳid")) & ","
                            mstrSQL = mstrSQL & Val(mrsParam("Ӥ��")) & ","
                            mstrSQL = mstrSQL & "To_Date('" & strTime & "','yyyy-mm-dd hh24:mi:ss'),"
                            mstrSQL = mstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
                            mstrSQL = mstrSQL & "1,"
                            
                            Select Case lngItemCode
                            Case mItemNo.����
                                If mint����Ӧ�� = 2 Then
                                    mstrSQL = mstrSQL & IIf(strValues = "", mItemNo.����, mItemNo.����) & ","
                                Else
                                    mstrSQL = mstrSQL & mItemNo.���� & ","
                                End If
                                                            
                            Case Else
                                mstrSQL = mstrSQL & lngItemCode & ","
                            End Select
                            
                            If lngItemCode = mItemNo.���� Then
                                mstrSQL = mstrSQL & "1,"
                            Else
                                mstrSQL = mstrSQL & IIf(strValues = "", "0", "1") & ","
                            End If
    
                            mstrSQL = mstrSQL & "'',"
                            
                            mstrSQL = mstrSQL & "'',1,1,0"
                                                        
                            mstrSQL = mstrSQL & ",0," & strMakeTime
                            mstrSQL = mstrSQL & ",'" & Trim(Split(.TextMatrix(GraphDataRow.δ��˵��, intCol), ";")(intCount + 1)) & "')"
                            
                            strSQL(ReDimArray(strSQL)) = mstrSQL
                    Else
                        For i = 0 To UBound(aryData)
                            
                            dblValues = ConvertToValue(intCount, aryData(i))
                            If intCount = mItemSerial.���� Then
                                dblValues = Format(dblValues, "0.00")
                            Else
                                dblValues = Format(dblValues, "0")
                            End If
                            
                            mstrSQL = "ZL_���ӻ����¼_UPDATE("
                            mstrSQL = mstrSQL & Val(mrsParam("����id")) & ","
                            mstrSQL = mstrSQL & Val(mrsParam("��ҳid")) & ","
                            mstrSQL = mstrSQL & Val(mrsParam("Ӥ��")) & ","
                            mstrSQL = mstrSQL & "To_Date('" & strTime & "','yyyy-mm-dd hh24:mi:ss'),"
                            mstrSQL = mstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
                            mstrSQL = mstrSQL & "1,"
                            
                            Select Case lngItemCode
                            Case mItemNo.����
                                If mint����Ӧ�� = 2 Then
                                    mstrSQL = mstrSQL & IIf(strValues = "", mItemNo.����, mItemNo.����) & ","
                                Else
                                    mstrSQL = mstrSQL & mItemNo.���� & ","
                                End If
                                                            
                            Case Else
                                mstrSQL = mstrSQL & lngItemCode & ","
                            End Select
                            
                            If lngItemCode = mItemNo.���� Then
                                mstrSQL = mstrSQL & "1,"
                            Else
                                mstrSQL = mstrSQL & IIf(strValues = "", "0", "1") & ","
                            End If
                            
                            '�����������Ŀ,����ֵΪ���ұ��˵��Ϊ"����",�轫ֵ����Ϊ"����",���˵������Ϊ��
                            If CStr(Val(dblValues)) = "0" And lngItemCode = mItemNo.���� And Trim(Split(.TextMatrix(GraphDataRow.δ��˵��, intCol), ";")(intCount + 1)) = "����" Then
                                mstrSQL = mstrSQL & "'����',"
                            Else
                                mstrSQL = mstrSQL & "'" & dblValues & "',"
                            End If
                            
                            mstrSQL = mstrSQL & IIf((lngItemCode = mItemNo.���� Or lngItemCode = mItemNo.���� Or lngItemCode = mItemNo.����) And strValues = "", "'" & Split(.TextMatrix(GraphDataRow.��λ��־, intCol), ";")(intCount + 1) & "'", "''") & ",1,1,"
                            mstrSQL = mstrSQL & IIf(lngItemCode = mItemNo.���� And i = 0, Val(.TextMatrix(GraphDataRow.���Ա�־, intCol)), "0")
                            
                            mstrSQL = mstrSQL & ",0," & strMakeTime
                            '�����������Ŀ,����ֵΪ���ұ��˵��Ϊ"����",�轫ֵ����Ϊ"����",���˵������Ϊ��
                            If CStr(Val(dblValues)) = "0" And lngItemCode = mItemNo.���� And Trim(Split(.TextMatrix(GraphDataRow.δ��˵��, intCol), ";")(intCount + 1)) = "����" Then
                                mstrSQL = mstrSQL & ",'')"
                            Else
                                mstrSQL = mstrSQL & ",'" & Trim(Split(.TextMatrix(GraphDataRow.δ��˵��, intCol), ";")(intCount + 1)) & "')"
                            End If
                            
                            strSQL(ReDimArray(strSQL)) = mstrSQL
                            
                            If strValues = "" Then
                                strValues = dblValues
                            Else
                                strValues = strValues & "," & dblValues
                            End If
                        Next
                    End If
                End If
            Next
        Next
    End With
    '------------------------------------------------------------------------------------------------------------------
    '3.��������������
    With vsf
        For intCol = 2 To .Cols - 1
            
            strTmp = GetCurveDateTime(intCol - 2 + 1, CDate(strFrom), mlngHourBegin)
            strTime = Split(strTmp, ",")(0)
            strEnd = Split(strTmp, ",")(1)
            
            If Int(CDate(strTime)) < Int(CDate(strFrom)) Then
                strTime = Format(strFrom, "yyyy-MM-dd HH:mm:ss")
            End If
            
            intMode = Val(vsf.ColData(intCol))
            mvarStrValue = Trim(vsf.TextMatrix(1, intCol))
            
            If intMode = OperateType.ɾ������ Or intMode = OperateType.�޸Ĳ��� Then
            
                mstrSQL = "ZL_���ӻ����¼_UPDATE("
                mstrSQL = mstrSQL & Val(mrsParam("����id")) & ","
                mstrSQL = mstrSQL & Val(mrsParam("��ҳid")) & ","
                mstrSQL = mstrSQL & Val(mrsParam("Ӥ��")) & ","
                mstrSQL = mstrSQL & "To_Date('" & strTime & "','yyyy-mm-dd hh24:mi:ss'),"
                mstrSQL = mstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
                mstrSQL = mstrSQL & "1,"
                mstrSQL = mstrSQL & mItemNo.���� & ","
                mstrSQL = mstrSQL & "0,"
                mstrSQL = mstrSQL & "NULL"
                mstrSQL = mstrSQL & ")"
                
                strSQL(ReDimArray(strSQL)) = mstrSQL
                
            End If
            
            If intMode = OperateType.�������� Or intMode = OperateType.�޸Ĳ��� And mvarStrValue <> "" Then
                mstrSQL = "ZL_���ӻ����¼_UPDATE("
                mstrSQL = mstrSQL & Val(mrsParam("����id")) & ","
                mstrSQL = mstrSQL & Val(mrsParam("��ҳid")) & ","
                mstrSQL = mstrSQL & Val(mrsParam("Ӥ��")) & ","
                mstrSQL = mstrSQL & "To_Date('" & strTime & "','yyyy-mm-dd hh24:mi:ss'),"
                mstrSQL = mstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
                mstrSQL = mstrSQL & "1,"
                mstrSQL = mstrSQL & mItemNo.���� & ","
                mstrSQL = mstrSQL & "0,"
                mstrSQL = mstrSQL & "'" & mvarStrValue & "'"
                mstrSQL = mstrSQL & ")"
                
                strSQL(ReDimArray(strSQL)) = mstrSQL
            End If
            
        Next
    End With

    '------------------------------------------------------------------------------------------------------------------
    '3.�������±������
    With mshDownTab
        If .Tag = "1" And .RowData(1) > 0 Then
            For intCol = .FixedCols To .Cols - 1 Step 2
                
                '���ʱ��
                
                strTmp = GetEditDateTime(intCol - .FixedCols + 1, CDate(strFrom))
                
                strTime = Split(strTmp, ",")(0)
                strEnd = Split(strTmp, ",")(1)
                
                '��������б�
                
                For intCount = 0 To .Rows - 2
                    '�������
                    strItem = .TextMatrix(intCount + 1, 1)
                    
                    '��������
                    
                    For intLoop = intCol To intCol + 1
                        
                        aryValue = Split(.TextMatrix(0, intLoop), ";")
                        intMode = Val(aryValue(intCount))
                        
                        mvarStrValue = .TextMatrix(intCount + 1, intLoop)
                                                
                        If intLoop = intCol + 1 Then
                            strEnd = Format(DateAdd("d", 1, CDate(Left(strTime, 10))), "yyyy-MM-dd") & " 00:00:00"
                            strTime = Left(strTime, 10) & " 12:00:00"
                        Else
                            strTime = Left(strTime, 10) & " 00:00:00"
                            strEnd = Left(strTime, 10) & " 12:00:00"
                        End If
                        
                        strStart = Format(CDate(strTime) - (4 - 4) / 24, "YYYY-MM-DD hh:mm:ss")
                        If Int(CDate(strStart)) <> Int(CDate(strFrom)) Then strStart = Format(strTime, "yyyy-MM-dd HH:mm:ss")
                        If strStart < mstr��Сʱ�� Then strStart = mstr��Сʱ��
                        strEnd = Format(CDate(strEnd) - (4 - 4) / 24, "YYYY-MM-DD hh:mm:ss")
                        
                        If intMode = 4 Or intMode = 3 Then
                            
                            mstrSQL = "ZL_���ӻ����¼_UPDATE("
                            mstrSQL = mstrSQL & Val(mrsParam("����id")) & ","
                            mstrSQL = mstrSQL & Val(mrsParam("��ҳid")) & ","
                            mstrSQL = mstrSQL & Val(mrsParam("Ӥ��")) & ","
                            mstrSQL = mstrSQL & "To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'),"
                            mstrSQL = mstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
                            mstrSQL = mstrSQL & "1,"
                            mstrSQL = mstrSQL & Val(.RowData(intCount + 1)) & ","
                            mstrSQL = mstrSQL & "0,"
                            mstrSQL = mstrSQL & "NULL"
                            mstrSQL = mstrSQL & ")"
                            
                            strSQL(ReDimArray(strSQL)) = mstrSQL
                            
                            If Val(.RowData(intCount + 1)) = mItemNo.Ѫѹ Then
                                mstrSQL = "ZL_���ӻ����¼_UPDATE("
                                mstrSQL = mstrSQL & Val(mrsParam("����id")) & ","
                                mstrSQL = mstrSQL & Val(mrsParam("��ҳid")) & ","
                                mstrSQL = mstrSQL & Val(mrsParam("Ӥ��")) & ","
                                mstrSQL = mstrSQL & "To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'),"
                                mstrSQL = mstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
                                mstrSQL = mstrSQL & "1,"
                                mstrSQL = mstrSQL & mItemNo.����ѹ & ","
                                mstrSQL = mstrSQL & "0,"
                                mstrSQL = mstrSQL & "NULL"
                                mstrSQL = mstrSQL & ")"
                                
                                strSQL(ReDimArray(strSQL)) = mstrSQL
                            End If
                            
                        End If
                        If (intMode = 2 Or intMode = 3) And mvarStrValue <> "" Then
                            
                            
                            If Val(.RowData(intCount + 1)) = mItemNo.Ѫѹ Then
                            
                                mstrSQL = "ZL_���ӻ����¼_UPDATE("
                                mstrSQL = mstrSQL & Val(mrsParam("����id")) & ","
                                mstrSQL = mstrSQL & Val(mrsParam("��ҳid")) & ","
                                mstrSQL = mstrSQL & Val(mrsParam("Ӥ��")) & ","
                                mstrSQL = mstrSQL & "To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'),"
                                mstrSQL = mstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
                                mstrSQL = mstrSQL & "1,"
                                mstrSQL = mstrSQL & Val(.RowData(intCount + 1)) & ","
                                mstrSQL = mstrSQL & "0,"
                                mstrSQL = mstrSQL & "'" & Split(mvarStrValue, "/")(0) & "'"
                                mstrSQL = mstrSQL & ")"
                                
                                strSQL(ReDimArray(strSQL)) = mstrSQL
                            
                                mstrSQL = "ZL_���ӻ����¼_UPDATE("
                                mstrSQL = mstrSQL & Val(mrsParam("����id")) & ","
                                mstrSQL = mstrSQL & Val(mrsParam("��ҳid")) & ","
                                mstrSQL = mstrSQL & Val(mrsParam("Ӥ��")) & ","
                                mstrSQL = mstrSQL & "To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'),"
                                mstrSQL = mstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
                                mstrSQL = mstrSQL & "1,"
                                mstrSQL = mstrSQL & mItemNo.����ѹ & ","
                                mstrSQL = mstrSQL & "0,"
                                mstrSQL = mstrSQL & "'" & Split(mvarStrValue, "/")(1) & "'"
                                mstrSQL = mstrSQL & ")"
                                
                                strSQL(ReDimArray(strSQL)) = mstrSQL
                            Else
                                mstrSQL = "ZL_���ӻ����¼_UPDATE("
                                mstrSQL = mstrSQL & Val(mrsParam("����id")) & ","
                                mstrSQL = mstrSQL & Val(mrsParam("��ҳid")) & ","
                                mstrSQL = mstrSQL & Val(mrsParam("Ӥ��")) & ","
                                mstrSQL = mstrSQL & "To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'),"
                                mstrSQL = mstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
                                mstrSQL = mstrSQL & "1,"
                                mstrSQL = mstrSQL & Val(.RowData(intCount + 1)) & ","
                                mstrSQL = mstrSQL & "0,"
                                mstrSQL = mstrSQL & "'" & mvarStrValue & "'"
                                mstrSQL = mstrSQL & ",Null,1,1,0," & IIf(IsNumeric(mvarStrValue), 0, 1) & ")"
                                
                                strSQL(ReDimArray(strSQL)) = mstrSQL
                            End If
                        End If
                                            
                    Next
                Next
            Next
        End If
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    'ѭ��ִ��SQL��������
    gcnOracle.BeginTrans
    blnTrans = True
    intMax = UBound(strSQL)
    For intTmp = 1 To intMax
        If strSQL(intTmp) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(intTmp), "������������")
    Next
    gcnOracle.CommitTrans
    blnTrans = False
    SaveData = True
    
    Screen.MousePointer = 0
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
ErrHead:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
    
    If blnTrans Then gcnOracle.RollbackTrans
    Call SaveErrLog
End Function

Private Sub cboBaby_Click()
    
    If opt(1).Value = False Then Exit Sub
    
    If Val(mrsParam("Ӥ��").Value) = cboBaby.ItemData(cboBaby.ListIndex) Then Exit Sub
    mrsParam("Ӥ��").Value = cboBaby.ItemData(cboBaby.ListIndex)
    
    If InitBody(Val(mrsParam("����id").Value), Val(mrsParam("��ҳid").Value), Val(mrsParam("����id").Value), Val(mrsParam("Ӥ��").Value)) = False Then Exit Sub

    Call zlMenuClick("��ʾ��������")
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.Id
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Jump
        
        mcbrToolBarҳ��.Caption = Control.Caption
        Call zlMenuClick("װ������", Control.Parameter)
        cbsMain.RecalcLayout
        
    End Select
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long
    Dim lngTop  As Long
    Dim lngRight  As Long
    Dim lngBottom  As Long

    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)

    On Error Resume Next

    '���������ؼ�Resize����
    picPane.Move lngLeft, lngTop, lngRight - lngLeft, lngBottom - lngTop
    picPane.BackColor = pic.BackColor
    
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    Select Case Control.Id
    Case conMenu_View_Option
        Control.Visible = mblnBabys
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Jump
        
        If Control.Parameter = "" Then
            Control.Checked = True
        Else
            Control.Checked = (Val(Split(Control.Parameter, ";")(4)) = Page)
        End If
        
        
    End Select
    
End Sub

'######################################################################################################################
'�¼�

Private Sub hsb_Change()
    
    On Error Resume Next
    
'    pic.Left = 60 - hsb.Value * 300
    pic.Left = -60 - hsb.Value * msinHStep
End Sub

Private Sub mfrmCaseTendBodyPrint_AfterPrint()
    RaiseEvent zlAfterPrint
End Sub

Private Sub mshDownTab_DblClick()
    
    If Val(mrsParam("�༭")) = 0 Then Exit Sub
    If mshDownTab.Tag = "" Or mvarEdit = False Then Exit Sub
    If CheckTimeRange(mshDownTab.Col) = False Then Exit Sub

    txtInput(1).Text = ""
    Call ShowInput

End Sub

Private Function ShowInput() As Boolean
    Dim strTmp As String
    
    With mshDownTab
        
        picInput.Move .Cell(flexcpLeft, .Row, .Col), .Cell(flexcpTop, .Row, .Col), .Cell(flexcpWidth, .Row, .Col) - 15, .Cell(flexcpHeight, .Row, .Col) - 15
        picInput.Visible = True
        picInput.BackColor = .Cell(flexcpBackColor, .Row, .Col)
        picInput.Tag = .Row & ";" & .Col & ";" & ""
        txtInput(0).BackColor = picInput.BackColor
        
        If mItemNo.Ѫѹ = Val(.RowData(.Row)) Then
            txtInput(1).BackColor = picInput.BackColor
            lblInput.Caption = "/"
            
            txtInput(0).Move 0, 0, (picInput.Width - lblInput.Width) / 2, picInput.Height
            lblInput.Left = txtInput(0).Left + txtInput(0).Width
            txtInput(1).Move lblInput.Left + lblInput.Width, 0, (picInput.Width - lblInput.Width) / 2, picInput.Height
            
            strTmp = .TextMatrix(.Row, .Col)
            If InStr(strTmp, "/") > 0 Then
                
                txtInput(0).Text = Left(strTmp, InStr(strTmp, "/") - 1)
                txtInput(1).Text = Mid(strTmp, InStr(strTmp, "/") + 1)
                
            Else
                txtInput(0).Text = strTmp
            End If
            txtInput(0).Alignment = 2
            txtInput(0).SelStart = 0
            txtInput(0).SelLength = 3
            txtInput(1).Visible = True
            txtInput(1).SelStart = 0
            txtInput(1).SelLength = 3
        Else
            txtInput(1).Visible = False
            
            lblInput.Caption = ""
            txtInput(0).Alignment = 1
            txtInput(0).Move 0, 0, picInput.Width, picInput.Height
            txtInput(0).Text = .TextMatrix(.Row, .Col)
            txtInput(0).MaxLength = IIf(mItemStru(.Row).���ݳ��� < 12, 12, mItemStru(.Row).���ݳ���)
            lblInput.Move txtInput(0).Left + txtInput(0).Width, txtInput(0).Top
        
            txtInput(0).SelStart = 0
            txtInput(0).SelLength = 100
        End If

        txtInput(0).SetFocus
    End With
    
End Function

Private Sub mshDownTab_KeyPress(KeyAscii As Integer)
    If Val(mrsParam("�༭")) = 0 Then Exit Sub

    If mshDownTab.Tag = "" Or mvarEdit = False Then Exit Sub
    If mshDownTab.RowData(mshDownTab.Row) <= 0 Then Exit Sub
    If CheckTimeRange(mshDownTab.Col) = False Then Exit Sub
    
    Select Case KeyAscii
    Case 13                 'Enter�ƶ���Ԫ��
        With mshDownTab
            If .Row = .Rows - 1 Then
                If .Col < .Cols - 1 Then
                    .Col = .Col + 1
                End If
                .Row = .FixedRows
            Else
                If .RowHidden(.Row + 1) Then
                    .Row = .Row + 2
                Else
                    .Row = .Row + 1
                End If
            End If
        End With
        Call mshDownTab_RowColChange
    Case 32                 '�ո������༭
        Call mshDownTab_DblClick
    Case vbKeyDelete
        Call zlMenuClick("ɾ����Ŀ")
    Case Else

        Select Case mItemStru(mshDownTab.Row).��������
        Case 0 '��ֵ��
            
            Select Case Val(mshDownTab.RowData(mshDownTab.Row))
            Case mItemNo.���
                If Check�Ƿ����(UCase(Chr(KeyAscii)), "0123456789+/E*") = False Then KeyAscii = 0
            Case mItemNo.��Һ
                If Check�Ƿ����(UCase(Chr(KeyAscii)), "0123456789/C") = False Then KeyAscii = 0
            Case Else
                If Check�Ƿ����(UCase(Chr(KeyAscii)), "��С��") = True Then KeyAscii = 0
            End Select

        Case 1 '�ַ���
            If Check�Ƿ����(UCase(Chr(KeyAscii)), "'") = True Then KeyAscii = 0
        End Select
        
        If KeyAscii <> 0 Then
            Call mshDownTab_DblClick
            txtInput(0).Text = Chr(KeyAscii)
            txtInput(0).SelStart = Len(txtInput(0).Text)
        End If
    End Select
End Sub

Private Sub mshDownTab_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim aryValue() As String
    Dim intRewrite As Integer
    
    If Val(mrsParam("�༭")) = 0 Then Exit Sub
    
    If mshDownTab.Tag = "" Or mvarEdit = False Then Exit Sub
    If CheckTimeRange(mshDownTab.Col) = False Then Exit Sub
    
    If KeyCode = 46 Then        'Delete�����Ԫ
        With mshDownTab
            '�����Ԫ��ֵ
            If .TextMatrix(.Row, .Col) = "" Then Exit Sub
            .TextMatrix(.Row, .Col) = ""
            
            '����ɾ�����
            aryValue() = Split(.TextMatrix(0, .Col), ";")
            intRewrite = Val(aryValue(.Row - 1))
            If ((.Col + 1) - mshDownTab.FixedCols) / 2 = ((.Col + 1) - mshDownTab.FixedCols) \ 2 Then
                '���Ϊ����
                Select Case intRewrite
                Case 0
                    aryValue(.Row - 1) = 0
                Case 1
                    aryValue(.Row - 1) = 4
                Case 2
                    aryValue(.Row - 1) = 0
                Case 3
                    aryValue(.Row - 1) = 4
                Case 4
                    aryValue(.Row - 1) = 4
                End Select
            Else
                '����Ϊ����
                Select Case intRewrite
                Case 0
                    aryValue(.Row - 1) = 2
                Case 1
                    aryValue(.Row - 1) = 3
                Case 2
                    aryValue(.Row - 1) = 2
                Case 3
                    aryValue(.Row - 1) = 3
                Case 4
                    aryValue(.Row - 1) = 3
                End Select
            End If
            .TextMatrix(0, .Col) = Join(aryValue, ";")
            '���Ϊ����ͼ���������Ĳ�����־
            If ((.Col + 1) - mshDownTab.FixedCols) / 2 = ((.Col + 1) - mshDownTab.FixedCols) \ 2 Then
                aryValue() = Split(.TextMatrix(0, .Col - 1), ";")
                intRewrite = Val(aryValue(.Row - 1))
                '���ӻ��޸Ĳ���
                Select Case intRewrite
                Case 0
                    aryValue(.Row - 1) = 2
                Case 1
                    aryValue(.Row - 1) = 3
                Case 2
                    aryValue(.Row - 1) = 2
                Case 3
                    aryValue(.Row - 1) = 3
                Case 4
                    aryValue(.Row - 1) = 3
                End Select
                .TextMatrix(0, .Col - 1) = Join(aryValue, ";")
            End If
        End With
        picInput.Visible = False
        
    End If
End Sub

Private Sub mshDownTab_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Val(mrsParam("�༭")) = 0 Then Exit Sub
    
    If mshDownTab.Tag = "" Or mvarEdit = False Then Exit Sub
    If CheckTimeRange(mshDownTab.Col) = False Then Exit Sub
End Sub

Private Sub mshDownTab_RowColChange()
    
    Dim strFrom As String, strTo As String
    Dim intNowRow As Integer, intNowCol As Integer
    Dim strInfo As String
    On Error GoTo ErrHead
    
    If mshDownTab.Tag = "" Or mvarEdit = False Or picScale.Tag = "" Then Exit Sub
    If CheckTimeRange(mshDownTab.Col) = False Then
        mshDownTab.FocusRect = flexFocusLight
        Exit Sub
    Else
        mshDownTab.FocusRect = flexFocusSolid
    End If

    strFrom = Split(picScale.Tag, ";")(0)
    strTo = Split(picScale.Tag, ";")(1)
    With mshDownTab

        For intRow = .FixedRows To .Rows - 1
            For intCol = .FixedCols To .Cols - 1
                If intCol / 2 <> intCol \ 2 Then
                    '˫����ʱΪ��ɫ
                    .Cell(flexcpBackColor, intRow, intCol, intRow, intCol) = &HF7ECE6
                Else
                    '������Ϊ��ɫ
                    .Cell(flexcpBackColor, intRow, intCol, intRow, intCol) = &H80000005
                End If
            Next
        Next
    End With
    
    If (Val(mshDownTab.TextMatrix(mshDownTab.Row, 2)) <> 0 Or Val(mshDownTab.TextMatrix(mshDownTab.Row, 3)) <> 0) And mItemStru(mshDownTab.Row).�������� = 0 Then
        strInfo = "��" & mshDownTab.TextMatrix(mshDownTab.Row, 1) & "����Ŀ��Χ��" & Val(mshDownTab.TextMatrix(mshDownTab.Row, 3)) & "��" & Val(mshDownTab.TextMatrix(mshDownTab.Row, 2)) & " " & strInfo
    End If
    
    RaiseEvent PromptInfo(strInfo)
    
    Exit Sub
ErrHead:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mshScale_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    
    If Val(mrsParam("�༭")) = 0 Then Exit Sub
    If Not mvarEdit Then Exit Sub
    
    With mshScale
        intCol = (X - .Left) \ .ColWidth(0)
        If Button = 1 Then
            If intCol >= .FixedCols Then Exit Sub
            
            .Cell(flexcpBackColor, 0, 0, .Rows - 1, .FixedCols - 1) = RGB(255, 255, 255)
            
            If picGraph.Tag = CStr(intCol) Then
                Call ClearLineSelect
                mlngLine = 0
                RaiseEvent PromptInfo("")
            Else
                RaiseEvent PromptInfo("")
                mlngLine = intCol + 1

                .Cell(flexcpBackColor, 0, intCol, .Rows - 1, intCol) = RGB(0, 255, 255)
                
                picGraph.Tag = intCol
                
                picGraph.MousePointer = 2

                linHCur.BorderColor = .Cell(flexcpForeColor, 0, intCol)
                linVCur.BorderColor = linHCur.BorderColor
                
                linHCur.X1 = 0: linHCur.X2 = 0: linHCur.Y1 = 0: linHCur.Y2 = 0
                linHCur.Visible = True
                
                linVCur.X1 = 0: linVCur.X2 = 0: linVCur.Y1 = 0: linVCur.Y2 = 0
                linVCur.Visible = True
            End If
            RaiseEvent SelectScale(intCol)
        ElseIf mItemSerial.���� = intCol Or mItemSerial.���� = intCol Or mItemSerial.���� = intCol Then

            RaiseEvent RButton(Button, Shift, X, Y)
            
        End If
    End With
End Sub

Private Sub lblCur_Click()
    picScale.SetFocus
End Sub

Private Sub lblCur_DblClick()
    picScale_KeyDown 13, 0
End Sub

Private Sub mshUpTab_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strDay As String
    Dim strFrom As String
    Dim strTo As String
    
    mshUpTab.FocusRect = flexFocusLight
    If picScale.Tag = "" Then Exit Sub
    If InStr(picScale.Tag, ";") = 0 Then Exit Sub
    
    strFrom = Split(picScale.Tag, ";")(0)
    strTo = Split(picScale.Tag, ";")(1)
    
    strDay = Format(Int(CDate(strFrom) + NewCol - mshUpTab.FixedCols), "yyyy-MM-dd")
    
    If strDay >= Format(strFrom, "yyyy-MM-dd") And strDay <= Format(strTo, "yyyy-MM-dd") Then
        mshUpTab.FocusRect = flexFocusSolid
    Else
        mshUpTab.FocusRect = flexFocusLight
    End If
    
End Sub

Private Sub mshUpTab_DblClick()
    If Val(mrsParam("�༭")) = 0 Then Exit Sub
        
    Call zlMenuClick("��д������")
End Sub

Private Sub mshUpTab_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intCol As Integer
    Dim strTime As String
    Dim intStart As Integer
    Dim intEnd As Integer
    Dim intLoop As Integer
    Dim strCaption As String
    
    On Error GoTo ErrHead
    
    If Val(mrsParam("�༭")) = 0 Then Exit Sub
    
    Select Case KeyCode
    Case vbKeyReturn ' 13     '����������
        With mshUpTab
        
            If Trim(.TextMatrix(2, .Col)) = "0" Then Exit Sub       '�Ѿ����������ڣ�ֱ���˳�

            strTime = mstrOpsDays(.Col)
            If strTime = "" Then

                strTime = GetCurveDateTime((.Col - .FixedCols) * 6 + mshScale.FixedCols - 1, CDate(Split(picScale.Tag, ";")(0)), mlngHourBegin)
                strTime = Split(strTime, ",")(0)
                
            End If
            
            intCol = GetCurveColumn(CDate(strTime), CDate(Split(picScale.Tag, ";")(0)), mlngHourBegin) + mshScale.FixedCols - 1
             
            strCaption = mshScale.TextMatrix(3, intCol)
            If frmInputDate.ShowMe(strTime, Split(picScale.Tag, ";")(0), Split(picScale.Tag, ";")(1), strCaption) Then
                
                mshUpTab.Tag = "��д������"
                
                intCol = GetCurveColumn(CDate(strTime), CDate(Split(picScale.Tag, ";")(0)), mlngHourBegin) + mshScale.FixedCols - 1
                
                .Col = Int((intCol - mshScale.FixedCols) / 6) + 1

                If Trim(.TextMatrix(2, .Col)) <> "" And Trim(.TextMatrix(2, .Col)) <> "0" Then
                    If MsgBox("����" & .TextMatrix(2, .Col) & "��ǰ��������/���䣬�Ƿ��ٴ�����/���䣿", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                End If
                
                mstrOpsDays(.Col) = strTime
                
                '�������ǰ�����ڵ���������������ʾ����
                intStart = GetCurveColumn(CDate(Format(strTime, "yyyy-MM-dd") & " 00:00:00"), CDate(Split(picScale.Tag, ";")(0)), mlngHourBegin) + mshScale.FixedCols - 1
                intEnd = GetCurveColumn(CDate(Format(strTime, "yyyy-MM-dd") & " 23:00:00"), CDate(Split(picScale.Tag, ";")(0)), mlngHourBegin) + mshScale.FixedCols - 1

                For intLoop = intStart To intEnd
                    If Left(mshScale.TextMatrix(3, intLoop), 2) = "����" Then
                        mshScale.TextMatrix(3, intLoop) = ""
                    End If
                    If Left(mshScale.TextMatrix(3, intLoop), 2) = "����" Then
                        mshScale.TextMatrix(3, intLoop) = ""
                    End If
                  If Left(mshScale.TextMatrix(3, intLoop), 2) = "��������" Then
                        mshScale.TextMatrix(3, intLoop) = ""
                    End If
                Next
                                                         
                Select Case strCaption
                Case "����"
                    If mBodyFlag.���� = 2 Then
                        mshScale.TextMatrix(3, intCol) = strCaption & "--" & ConvertTimeToChinese(Format(strTime, "HH:mm"))
                    Else
                        mshScale.TextMatrix(3, intCol) = strCaption
                    End If
                Case Else
                    If mBodyFlag.���� = 2 Then
                        mshScale.TextMatrix(3, intCol) = strCaption & "--" & ConvertTimeToChinese(Format(strTime, "HH:mm"))
                    Else
                        mshScale.TextMatrix(3, intCol) = strCaption
                    End If
                End Select
                
                mshScale.Cell(flexcpData, 3, intCol, 3, intCol) = Format(strTime, "HH:mm:ss")
                
                Select Case .ColData(.Col)
                Case 0      '�������գ���дΪ��������
                    .ColData(.Col) = 2
                Case 1      'ԭ������������
                Case 2      '�Ѿ�����Ϊ��������
                Case 3      '��ɾ���ĵ������գ��ٴ�����Ϊ������
                    .ColData(.Col) = 1
                End Select
                
                Call ShowOpsDays
                Call DrawPaper
                Call DrawGraph
            
            End If

        End With
        
    Case vbKeyDelete ' 46     '���������
        With mshUpTab
'            If Trim(.TextMatrix(2, .Col)) <> "0" Then Exit Sub          '��ǰ����������
            If .ColData(.Col) <> 1 And .ColData(.Col) <> 2 Then Exit Sub
            
            If MsgBox("�Ƿ��������" & .TextMatrix(0, .Col) & "�յ������Ǽǣ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            mshUpTab.Tag = "ɾ��������"
            Select Case .ColData(.Col)
            Case 0      '��������
            Case 1      'ԭ�����������գ�����Ϊɾ��������
                .ColData(.Col) = 3
            Case 2      '�������գ��ٴ�����Ϊ��������
                .ColData(.Col) = 0
            Case 3      '��ɾ���ĵ�������
            End Select
            
            intCol = GetCurveColumn(CDate(mstrOpsDays(.Col)), CDate(Split(picScale.Tag, ";")(0)), mlngHourBegin) + mshScale.FixedCols - 1
            
            If intCol > 0 Then mshScale.TextMatrix(3, intCol) = ""
            mstrOpsDays(.Col) = ""
            
            Call ShowOpsDays
            Call DrawPaper
            Call DrawGraph

        End With
        
    End Select
    Exit Sub
ErrHead:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub opt_Click(Index As Integer)
    
    cboBaby.Enabled = (opt(1).Value = True)
    
    Select Case Index
    Case 0                  '���˱���
        
        If Val(mrsParam("Ӥ��").Value) = 0 Then Exit Sub
        mrsParam("Ӥ��").Value = 0
        
        If InitBody(Val(mrsParam("����id").Value), Val(mrsParam("��ҳid").Value), Val(mrsParam("����id").Value), Val(mrsParam("Ӥ��").Value)) = False Then Exit Sub
        
        Call zlMenuClick("��ʾ��������")
        
    Case 1                  'Ӥ��
        
        Call cboBaby_Click
        
    End Select
        
End Sub

Private Sub picCard_Paint(Index As Integer)
    Dim intLoop As Integer
    
    On Error Resume Next
    
    picCard(Index).Cls
    For intLoop = 0 To txtCard.UBound
        txtCard(intLoop).Height = 180
        If txtCard(intLoop).Visible Then
            DrawLine picCard(Index), txtCard(intLoop).Left, txtCard(intLoop).Top + txtCard(intLoop).Height + 15, txtCard(intLoop).Left + txtCard(intLoop).Width, txtCard(intLoop).Top + txtCard(intLoop).Height + 15, &H8000000C
        End If
    Next
End Sub

Private Sub picCard_Resize(Index As Integer)
    On Error Resume Next
    
    txtCard(1).Move txtCard(1).Left, txtCard(1).Top, picCard(Index).Width - txtCard(1).Left - 45
    txtCard(7).Move txtCard(7).Left, txtCard(7).Top, picCard(Index).Width - txtCard(7).Left - 45
    
End Sub

Private Sub picGraph_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo ErrHead
    '-------------------------------------------------
    '1�������������ƶ�
    '2��������ƶ������������������ƣ������ʮ�ֱ�ɲ����ƶ�����߲����ƶ�
    '-------------------------------------------------
    Dim intMinCol As Long
    Dim intMaxCol As Long
    Dim sglLeft As Single
    Dim sglRight As Single
    
    Dim aryValue() As String
    Dim aryNote() As String
    
    
    If Val(mrsParam("�༭")) = 0 Then Exit Sub
    If picGraph.Tag = "" Or picScale.Tag = "" Then Exit Sub
    
    Call CalcMinMaxCol(picScale.Tag, intMinCol, intMaxCol)

    If picGraph.MousePointer <> 2 Then picGraph.MousePointer = 2
    
    '�����С���ʱ�䷶Χ
    aryValue = Split(picScale.Tag, ";")
    
    sglLeft = intMinCol * HOUR_STEP_Twips + 30
    sglRight = (intMaxCol + 1) * HOUR_STEP_Twips - 30
    
    If X < sglLeft Then
        X = sglLeft
        If picGraph.MousePointer <> 12 Then picGraph.MousePointer = 12
    End If
    
    If X > sglRight Then
        X = sglRight
        If picGraph.MousePointer <> 12 Then picGraph.MousePointer = 12
    End If
        
    
    '��ȡ��Ŀ����:���ֵ����Сֵ����λֵ�������
    aryValue = Split(picLine(Val(picGraph.Tag)).Tag, ";")
    If Y < (aryValue(3) - 1) * mshScale.ROWHEIGHT(1) Then
        Y = (aryValue(3) - 1) * mshScale.ROWHEIGHT(1)
        
        If picGraph.MousePointer <> 12 Then picGraph.MousePointer = 12
        
    End If
    If Y > (aryValue(3) - 1 + (aryValue(0) - aryValue(1)) / aryValue(2)) * mshScale.ROWHEIGHT(1) Then
        Y = (aryValue(3) - 1 + (aryValue(0) - aryValue(1)) / aryValue(2)) * mshScale.ROWHEIGHT(1)
        If picGraph.MousePointer <> 12 Then picGraph.MousePointer = 12
    End If
    
    With linHCur
        .X1 = 0: .X2 = X:
        .Y1 = Y: .Y2 = Y
    End With
    With linVCur
        .X1 = X: .X2 = X:
        .Y1 = 0: .Y2 = Y
    End With
    
    '״̬��ʾ������ʾ����
    '------------------------------------------------------------------------------------------------------------------
    Dim intNowCol As Integer
    Dim dblValues As Single
    Dim strTmp As String
    
    intNowCol = (linVCur.X1 \ HOUR_STEP_Twips) + 1
    
    '�����ǰ���ǶϿ��ģ���������ͼ
    If Val(mshScale.TextMatrix(GraphDataRow.�Ͽ���־, intNowCol + mshScale.FixedCols - 1)) = 1 Then
        If picGraph.MousePointer <> 12 Then picGraph.MousePointer = 12
    End If
       
    aryNote = Split(mshScale.TextMatrix(GraphDataRow.δ��˵��, intNowCol + mshScale.FixedCols - 1), ";")
    If aryNote(Val(picGraph.Tag) + 1) <> "" Then
        If picGraph.MousePointer <> 12 Then picGraph.MousePointer = 12
    End If
    
    aryValue = Split(picScale.Tag, ";")
    
    strTmp = GetCurveDateTime(intNowCol, CDate(aryValue(0)), mlngHourBegin)
    dblValues = ConvertToValue(Val(picGraph.Tag), Y)
    
    If mItemSerial.���� = Val(picGraph.Tag) Then
        dblValues = Format(dblValues, "0.00")
    Else
        dblValues = Format(dblValues, "0")
    End If
    
    If strTmp <> "" Then
        strTmp = "���ڣ�" & Format(Split(strTmp, ",")(0), "yyyy-MM-dd") & " ʱ�䣺" & Format(Split(strTmp, ",")(0), "HHʱmm��") & "��" & Format(Split(strTmp, ",")(1), "HHʱmm��")
    End If
    
    RaiseEvent PromptInfo(strTmp & " " & mshScale.TextMatrix(0, Val(picGraph.Tag)) & "��" & dblValues)
    
    Exit Sub
ErrHead:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub picGraph_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo ErrHead
    '-------------------------------------------------
    '���ݰ������״̬������ֵ��¼�����������������޸Ĳ���������Ҽ�����ɾ������
    '1�������ָ��ʱ��û����ֵ�����¼���ݣ��ֱ���ǰ�ͺ�ʱ�����Ƿ������ݣ��������������ӡ��
    '2�������ָ��ʱ���Ѿ�����ֵ,�򱣴����ݣ�ͬʱ������ҳ���»滭�ķ���
    '-------------------------------------------------
    Dim aryValue() As String
    Dim aryPart() As String
    Dim intRewrite As Integer
    Dim X0 As Long, Y0 As Long
    Dim strChar As String
    
    Dim aryData() As String
    Dim i As Long
    Dim intHave As Integer
    Dim intDots As Integer
    
    If picGraph.MousePointer <> 2 Then Exit Sub
    intCol = Int(X / HOUR_STEP_Twips)

    If Val(mshScale.TextMatrix(GraphDataRow.�Ͽ���־, intCol + mshScale.FixedCols)) = 1 Then Exit Sub
    
    X = intCol * HOUR_STEP_Twips + HOUR_STEP_Twips / 2
    
    aryValue = Split(mshScale.TextMatrix(GraphDataRow.���ı�־, mshScale.FixedCols + intCol), ";")
    intRewrite = Val(aryValue(picGraph.Tag + 1))
    '------------------------------------------------------------------------------------------------------------------
    If Button = 1 Then
        Dim dblY As Double
        '���ӻ��޸Ĳ���
        Select Case intRewrite
        Case 0 '��ԭ���޵Ļ�����ɾ����
            aryValue(picGraph.Tag + 1) = 2
        Case 1 'ԭ�����е�
            aryValue(picGraph.Tag + 1) = 3
        Case 2 '������
            aryValue(picGraph.Tag + 1) = 2
        Case 3 '�޸ĵ�
            aryValue(picGraph.Tag + 1) = 3
        Case 4 '��ԭ���еĻ�����ɾ����
            aryValue(picGraph.Tag + 1) = 3
        End Select
        
        mshScale.TextMatrix(GraphDataRow.���ı�־, mshScale.FixedCols + intCol) = Join(aryValue, ";")
        
        aryValue = Split(mshScale.TextMatrix(GraphDataRow.��������, mshScale.FixedCols + intCol), ";")
        If ������Ŀ Then
            '�������Ŀ�̶�̫С�����������
            dblY = Format(ConvertToValue(mItemSerial.����, Y), "0")
            Y = ConvertToY(mItemSerial.����, dblY)
        End If
        aryValue(picGraph.Tag + 1) = Y
        mshScale.TextMatrix(GraphDataRow.��������, mshScale.FixedCols + intCol) = Join(aryValue, ";")
        
        If Val(picGraph.Tag) = mItemSerial.���� Then
            aryPart() = Split(mshScale.TextMatrix(GraphDataRow.��λ��־, mshScale.FixedCols + intCol), ";")
            aryPart(picGraph.Tag + 1) = mstr���²�λ
            mshScale.TextMatrix(GraphDataRow.��λ��־, mshScale.FixedCols + intCol) = Join(aryPart, ";")
      
        End If
        If Val(picGraph.Tag) = mItemSerial.���� Then
            aryPart = Split(mshScale.TextMatrix(GraphDataRow.��λ��־, mshScale.FixedCols + intCol), ";")
            aryPart(picGraph.Tag + 1) = mstr������ʽ
            mshScale.TextMatrix(GraphDataRow.��λ��־, mshScale.FixedCols + intCol) = Join(aryPart, ";")
        End If
        If Val(picGraph.Tag) = mItemSerial.���� Then
            aryPart = Split(mshScale.TextMatrix(GraphDataRow.��λ��־, mshScale.FixedCols + intCol), ";")
            aryPart(picGraph.Tag + 1) = mstr����
            mshScale.TextMatrix(GraphDataRow.��λ��־, mshScale.FixedCols + intCol) = Join(aryPart, ";")
        End If
    '------------------------------------------------------------------------------------------------------------------
    ElseIf Not (intRewrite = 0 Or intRewrite = 4) Then
    
        '������λ���Ƿ���һ������(����)
        
        aryData = Split(mshScale.TextMatrix(GraphDataRow.��������, mshScale.FixedCols + intCol), ";")
        aryData = Split(aryData(picGraph.Tag + 1), ",")
        intDots = UBound(aryData) + 1 '�ѻ��ĵ���
        If Abs(Val(aryData(0)) - Y) <= 60 Then intHave = 1 '�ڵ�һ������
        For i = 1 To UBound(aryData)
            If Abs(Val(aryData(i)) - Y) <= 60 Then intHave = i + 1: Exit For '�ڵ�i������
        Next
        
        If intHave = 0 Then
            '������
            Select Case intRewrite
            Case 1
                aryValue(picGraph.Tag + 1) = 3
            Case 2
                aryValue(picGraph.Tag + 1) = 2
            Case 3
                aryValue(picGraph.Tag + 1) = 3
            End Select
            mshScale.TextMatrix(GraphDataRow.���ı�־, mshScale.FixedCols + intCol) = Join(aryValue, ";")
        
            '����������
            aryValue = Split(mshScale.TextMatrix(GraphDataRow.��������, mshScale.FixedCols + intCol), ";")
            If intDots = 1 Then
                '�����ڶ���
                Select Case Val(picGraph.Tag)
                Case mItemSerial.����, mItemSerial.����
                
                    '�����������Ŀ����ʾ�����£�������Ŀ��ʾ���ʣ����������
                    If Y <> Val(aryValue(picGraph.Tag + 1)) Then
                        aryValue(picGraph.Tag + 1) = aryValue(picGraph.Tag + 1) & "," & Y
                    End If
                    
                End Select
                
            Else
                '�޸ĵڶ���
                aryData = Split(aryValue(picGraph.Tag + 1), ",")
                If Y <> Val(aryData(0)) Then
                    aryData(intDots - 1) = Y
                    aryValue(picGraph.Tag + 1) = Join(aryData, ",")
                End If
                
            End If
            mshScale.TextMatrix(GraphDataRow.��������, mshScale.FixedCols + intCol) = Join(aryValue, ";")
        Else
            '��ĳ���㸽����ɾ���õ�
            Select Case intRewrite
            Case 1
                aryValue(picGraph.Tag + 1) = IIf(intDots > 1, 3, 4)
            Case 2
                aryValue(picGraph.Tag + 1) = IIf(intDots > 1, 2, 0)
            Case 3
                aryValue(picGraph.Tag + 1) = IIf(intDots > 1, 3, 4)
            End Select
            mshScale.TextMatrix(GraphDataRow.���ı�־, mshScale.FixedCols + intCol) = Join(aryValue, ";")
            
            'ɾ���õ�����(��ɾ��,�����ÿ�)
            aryValue = Split(mshScale.TextMatrix(GraphDataRow.��������, mshScale.FixedCols + intCol), ";")
            aryData = Split(aryValue(picGraph.Tag + 1), ",")
            aryData(intHave - 1) = " "
            aryValue(picGraph.Tag + 1) = Mid(Replace("," & Join(aryData, ","), ", ", ""), 2)
            mshScale.TextMatrix(GraphDataRow.��������, mshScale.FixedCols + intCol) = Join(aryValue, ";")
        End If
    End If
        
    Call DrawPaper
    Call DrawGraph
    Exit Sub
    
    '------------------------------------------------------------------------------------------------------------------
ErrHead:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub picPane_Resize()
    On Error GoTo errHand
    
    With vsb
        .Left = picPane.Width - .Width
        .Top = 0
        .Height = picPane.Height - hsb.Height
    End With
    
    With hsb
        .Left = 0
        .Top = picPane.Height - .Height
        .Width = picPane.Width - vsb.Width
    End With
    
    With picCover
        .Left = vsb.Left
        .Top = hsb.Top
    End With
    
    picCard(0).Move 60, 0
    
    '-------------------------------------------------
    'ҳ�����
    mshUpTab.Redraw = False
    mshScale.Redraw = False
    mshDownTab.Redraw = False
    
    With mshUpTab
        .Left = 60
        .Width = .ColWidth(0) + .ColWidth(1) * (.Cols - 1) + 15
        .Top = picCard(0).Top + picCard(0).Height + 30
        .Refresh
    End With
    
    picCard(0).Left = mshUpTab.Left
    picCard(0).Width = mshUpTab.Width
    
    With mshScale
        .Left = mshUpTab.Left
        .Width = mshUpTab.Width
        .Top = mshUpTab.Top + mshUpTab.Height - 15
        .Height = .Rows * .ROWHEIGHT(.Rows - 1) + 600
        .Refresh
    End With

    With vsf
        .Left = mshUpTab.Left
        .Top = mshScale.Top + mshScale.Height
        .Width = mshUpTab.Width
        .Visible = Not mbln��������
    End With
        
    With mshDownTab
        .RowHeightMin = 255
        .Left = mshUpTab.Left
        .Top = IIf(mbln��������, vsf.Top, vsf.Top + vsf.Height) - 15
        .Width = mshUpTab.Width
        .Height = (.Rows - 1) * .ROWHEIGHT(1) + 15
    End With
    
    lblComment.Left = mshUpTab.Left
    lblComment.Top = mshDownTab.Top + mshDownTab.Height + 45
        
    For intCol = 0 To picLine.UBound
        picLine(intCol).Move mshScale.Left + mshScale.ColWidth(0) * (intCol + 1), mshScale.Top, 15, mshScale.Height - 15
    Next
    
    With picScale
        .Left = mshUpTab.ColWidth(0) + mshUpTab.Left
        .Width = mshUpTab.ColWidth(1) * (mshUpTab.Cols - 1) + 15
        .Top = mshUpTab.Top + mshUpTab.Height - 15
        .Height = mshScale.ROWHEIGHT(0) + mshScale.ROWHEIGHT(1) / 2
    End With
    
    With picBack
        .Left = picScale.Left
        .Width = picScale.Width
        .Top = picScale.Top + picScale.Height - 15
        .Height = mshScale.Top + mshScale.Height - .Top
    End With
    
    mshUpTab.Redraw = True
    mshScale.Redraw = True
    mshDownTab.Redraw = True
    
    pic.Width = mshScale.Width + vsb.Width + 45
    pic.Height = lblComment.Top + lblComment.Height + 45 + hsb.Height
    
    '���������
    Call CalcScrollBarSize
    
    hsb.Value = 0
    vsb.Value = 0
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub picScale_GotFocus()
    
    If mvarEdit = False Then Exit Sub
    
    lblCur.ForeColor = &H80000012
End Sub

Private Sub picScale_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo ErrHead
    
    Dim aryValue() As String
    Dim intMinCol As Long
    Dim intMaxCol As Long
    
    If Val(mrsParam("�༭")) = 0 Then Exit Sub
    
    '�����С���ʱ�䷶Χ
    Call CalcMinMaxCol(picScale.Tag, intMinCol, intMaxCol)
    
    Select Case KeyCode
    Case 37     '���ƶ�
        If lblCur.Left - HOUR_STEP_Twips >= intMinCol * HOUR_STEP_Twips And (lblCur.Left - HOUR_STEP_Twips) >= 0 Then lblCur.Left = lblCur.Left - HOUR_STEP_Twips
    Case 39     '���ƶ�
        If lblCur.Left + HOUR_STEP_Twips <= intMaxCol * HOUR_STEP_Twips Then lblCur.Left = lblCur.Left + HOUR_STEP_Twips
    Case 13     'Enter��������
        
        RaiseEvent DbClickCur

    End Select
    Exit Sub
ErrHead:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub picScale_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo ErrHead
    Dim aryValue() As String
    '�����С���ʱ�䷶Χ
    If mvarEdit = False Then Exit Sub
    
    aryValue = Split(picScale.Tag, ";")
    If X < (Int((CDate(aryValue(0)) - Int(CDate(aryValue(0)))) * 24) \ 4 - 1) * HOUR_STEP_Twips + 15 Then Exit Sub
    If X > (Int((CDate(aryValue(1)) - Int(CDate(aryValue(0)))) * 24) \ 4 + 1) * HOUR_STEP_Twips - 15 Then Exit Sub
    lblCur.Left = Int(X / HOUR_STEP_Twips) * HOUR_STEP_Twips
    Exit Sub
ErrHead:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtCard_GotFocus(Index As Integer)
    
    zlControl.TxtSelAll txtCard(Index)
        
End Sub

Private Sub txtCard_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txtCard_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txtCard(Index).Locked Then
        glngTXTProc = GetWindowLong(txtCard(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtCard(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtCard_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txtCard(Index).Locked Then
        Call SetWindowLong(txtCard(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txtInput_Change(Index As Integer)
    Dim blnCancel As Boolean
    
    If Val(mshDownTab.RowData(mshDownTab.Row)) <> mItemNo.Ѫѹ Then Exit Sub
    If Index <> 0 Then Exit Sub
    
    If Len(txtInput(Index).Text) = 3 And GetTextPos(txtInput(Index).hWnd) = 4 Then
        
        If CheckBlood(0) Then
            txtInput(1).SetFocus
            zlControl.TxtSelAll txtInput(1)
        Else
            txtInput(1).Text = ""
            picInput.Visible = False
            mshDownTab.SetFocus
        End If
        
    End If

End Sub

Private Sub txtInput_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim blnCancel As Boolean
    
    If Val(mshDownTab.RowData(mshDownTab.Row)) <> mItemNo.Ѫѹ Then Exit Sub
    
    Select Case KeyCode
    Case vbKeyLeft
        If Index = 1 Then
            If GetTextPos(txtInput(Index).hWnd) = 1 Then
            
                If CheckBlood(1) Then
                    txtInput(0).SetFocus
                    zlControl.TxtSelAll txtInput(0)
                Else
                    txtInput(1).Text = ""
                    picInput.Visible = False
                    mshDownTab.SetFocus
                End If
        
            End If
        End If
    Case vbKeyRight
        If Index = 0 Then
            If GetTextPos(txtInput(Index).hWnd) >= Len(txtInput(0).Text) Then
                If CheckBlood(0) Then
                    txtInput(1).SetFocus
                    zlControl.TxtSelAll txtInput(1)
                Else
                    txtInput(1).Text = ""
                    picInput.Visible = False
                    mshDownTab.SetFocus
                End If
            End If
        End If
    Case vbKeyBack
        If Index = 1 Then
            If GetTextPos(txtInput(Index).hWnd) = 1 Then

                If CheckBlood(1) Then
                    txtInput(0).SetFocus
                    If Len(txtInput(0).Text) > 0 Then
                       txtInput(0).Text = Left(txtInput(0).Text, Len(txtInput(0).Text) - 1)
                    End If
                    txtInput(0).SelStart = Len(txtInput(0).Text)
                Else
                    txtInput(1).Text = ""
                    picInput.Visible = False
                    mshDownTab.SetFocus
                End If
            End If
        End If
    End Select
End Sub

Private Sub txtInput_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim intRow As Long
    Dim intCol As Long
    Dim blnCancel As Boolean
    
    If Val(mrsParam("�༭")) = 0 Then Exit Sub
    
    If KeyAscii = Asc("'") Or KeyAscii = Asc(";") Then
        KeyAscii = 0
    End If
    
    If mshDownTab.Tag = "" Then Exit Sub
    On Error Resume Next
    intRow = Split(picInput.Tag, ";")(0)
    intCol = Split(picInput.Tag, ";")(1)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call txtInput_Validate(Index, blnCancel)
        If blnCancel Then
            picInput.Visible = False
            Exit Sub
        End If
        
        If picInput.Visible Then
            mshDownTab.SetFocus
            Exit Sub
        Else
            mshDownTab.SetFocus
            Call mshDownTab_KeyPress(13)
        End If
        
    ElseIf KeyAscii = 27 Then
        KeyAscii = 0
        picInput.Visible = False
        mshDownTab.SetFocus
    Else
    
        If Val(mshDownTab.RowData(mshDownTab.Row)) = mItemNo.Ѫѹ Then
            Select Case KeyAscii
            Case 191, vbKeyDivide, Asc("/")
                KeyAscii = 0
                If Index = 0 Then
                
                    If CheckBlood(0) Then
                        txtInput(1).SetFocus
                        zlControl.TxtSelAll txtInput(1)
                    Else
                        txtInput(1).Text = ""
                        picInput.Visible = False
                        mshDownTab.SetFocus
                    End If
                
                    End If

            End Select
        End If
    
'        Select Case mItemStru(intRow).��������
'        Case 0 '��ֵ��
'            If Check�Ƿ����(UCase(Chr(KeyAscii)), "��С��") = True Then KeyAscii = 0
'        Case 1 '�ַ���
            If Check�Ƿ����(UCase(Chr(KeyAscii)), "'") = True Then KeyAscii = 0
'        End Select
    End If

End Sub

Private Sub txtInput_LostFocus(Index As Integer)
    If Val(mshDownTab.RowData(mshDownTab.Row)) = mItemNo.Ѫѹ Then Exit Sub
    
    txtInput(Index).Text = Replace(txtInput(Index).Text, "'", "")
    picInput.Visible = False
End Sub

Private Function WriteScaleTab(ByVal intRow As Integer, ByVal intCol As Integer, ByVal strInput As String) As Boolean

    Dim aryValue() As String
    Dim intRewrite As Integer
    Dim aryPara() As String


    On Error GoTo ErrHead
        
    With mshScale
        
        '������������
        aryValue = Split(.TextMatrix(GraphDataRow.���ı�־, intCol), ";")
        intRewrite = Val(aryValue(intRow + 1))
        
        If strInput <> "" Then
            '�������ݣ��൱�����ӻ��޸Ĳ���
            Select Case intRewrite
            Case 0
                aryValue(intRow + 1) = 2
            Case 1
                aryValue(intRow + 1) = 3
            Case 2
                aryValue(intRow + 1) = 2
            Case 3
                aryValue(intRow + 1) = 3
            Case 4
                aryValue(intRow + 1) = 3
            End Select
        Else
            'û�����ݣ��൱��ɾ������
            Select Case intRewrite
            Case 0
                aryValue(intRow + 1) = 0
            Case 1
                aryValue(intRow + 1) = 4
            Case 2
                aryValue(intRow + 1) = 0
            Case 3
                aryValue(intRow + 1) = 4
            Case 4
                aryValue(intRow + 1) = 4
            End Select
        End If
        .TextMatrix(0, intCol) = Join(aryValue, ";")
        
        aryValue = Split(.TextMatrix(1, intCol), ";")
        If strInput <> "" Then
            '��ȡָ����Ŀ���壺���ֵ����Сֵ����λֵ�������
            aryPara = Split(picLine(intRow).Tag, ";")
            aryValue(intRow + 1) = ((aryPara(0) - Val(strInput)) / aryPara(2) + aryPara(3) - 1) * .ROWHEIGHT(1)
        Else
            aryValue(intRow + 1) = ""
        End If
        .TextMatrix(1, intCol) = Join(aryValue, ";")
        
    End With
    
    '�����ϼ��������ͼ�δ���
    Call DrawPaper
    Call DrawGraph
    
    WriteScaleTab = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
ErrHead:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Function

Private Function WriteDownTab(ByVal intRow As Integer, ByVal intCol As Integer, ByVal strInput As String) As Boolean
    Dim aryValue() As String
    Dim intRewrite As Integer
    Dim strUnit As String
    Dim lngColor As Long
    
    On Error GoTo ErrHead
    
    With mshDownTab

        If .TextMatrix(intRow, intCol) = strInput Then Exit Function
        
        lngColor = GridTextColor(.TextMatrix(intRow, 0), strInput)
        .Cell(flexcpForeColor, intRow, intCol, intRow, intCol) = lngColor
             
        .TextMatrix(intRow, intCol) = strInput
        
        '������ɾ�ı�־(�������û���κ���ֵ����Ϊ����ɾ������)
        If InStr(.TextMatrix(GridDataRow.�޸ı�־, intCol), ";") = 0 Then
            ReDim aryValue(0 To 0) As String
            aryValue(0) = .TextMatrix(GridDataRow.�޸ı�־, intCol)
        Else
            aryValue() = Split(.TextMatrix(GridDataRow.�޸ı�־, intCol), ";")
        End If
        
        intRewrite = Val(aryValue(intRow - 1))
        
        If Trim(strInput) <> "" Then
            '���ӻ��޸Ĳ���
            Select Case intRewrite
            Case 0
                aryValue(intRow - 1) = 2
            Case 1
                aryValue(intRow - 1) = 3
            Case 2
                aryValue(intRow - 1) = 2
            Case 3
                aryValue(intRow - 1) = 3
            Case 4
                aryValue(intRow - 1) = 3
            End Select
        Else
            If ((intCol + 1) - mshDownTab.FixedCols) / 2 <> ((intCol + 1) - mshDownTab.FixedCols) \ 2 And .TextMatrix(intRow, intCol) = .TextMatrix(intRow, intCol + 1) And .TextMatrix(intRow, intCol + 1) = "" Then
                '�����ǰ�����Ԫ��Ϊ����ʱ
                'ɾ������<---ȡ����ǰ��ɾ��Ϊ�޸Ĳ���
                Select Case intRewrite
                Case 0
                    aryValue(intRow - 1) = 0
                Case 1
                    aryValue(intRow - 1) = 4
                Case 2
                    aryValue(intRow - 1) = 0
                Case 3
                    aryValue(intRow - 1) = 4
                Case 4
                    aryValue(intRow - 1) = 4
                End Select
            Else
                '�������ͣ�2-ȫ���������ĵ�,3-�޸ĵĵ㣺���ܰ���ԭ�еĵ�������ĵ�,4-ɾ���ĵ�
                'ɾ������<---ȡ����ǰ��ɾ��Ϊ�޸Ĳ���
                Select Case intRewrite
                Case 0
                    aryValue(intRow - 1) = 0
                Case 1
                    aryValue(intRow - 1) = 3
                Case 2
                    aryValue(intRow - 1) = 2
                Case 3
                    aryValue(intRow - 1) = 3
                Case 4
                    aryValue(intRow - 1) = 3
                End Select
            End If
        End If
        .TextMatrix(GridDataRow.�޸ı�־, intCol) = Join(aryValue, ";")
        
        '�޸ĵ�����Ԫ��Ĳ�����־
        '�����ǰ�����Ԫ��Ϊ������ô���޸�����Ĳ�����־Ϊ�޸Ļ�����־
        '���������ǰ��Ԫ��Ϊ���粢��Ϊɾ��ʱ�ͽ������ɾ��������־��Ϊ�޸ı�־
        
        If ((intCol + 1) - mshDownTab.FixedCols) / 2 = ((intCol + 1) - mshDownTab.FixedCols) \ 2 Then
            If Trim(.TextMatrix(GridDataRow.�޸ı�־, intCol - 1)) = "" Then
                .TextMatrix(GridDataRow.�޸ı�־, intCol - 1) = " "
            End If
            aryValue() = Split(.TextMatrix(GridDataRow.�޸ı�־, intCol - 1), ";")
            intRewrite = Val(aryValue(intRow - 1))
            
            '���ӻ��޸Ĳ���
            Select Case intRewrite
            Case 0
                aryValue(intRow - 1) = 2
            Case 1
                aryValue(intRow - 1) = 3
            Case 2
                aryValue(intRow - 1) = 2
            Case 3
                aryValue(intRow - 1) = 3
            Case 4
                aryValue(intRow - 1) = 3
            End Select
            
            .TextMatrix(GridDataRow.�޸ı�־, intCol - 1) = Join(aryValue, ";")
            
        End If
    End With
    
    WriteDownTab = True
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
ErrHead:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckBlood(Index As Integer) As Boolean
    Dim aryValue() As String
    Dim intRewrite As Integer
    Dim aryBeforeValue() As String
    Dim strInput As String
    Dim dbMin As Double
    Dim dbMax As Double
    Dim strName As String
    
    CheckBlood = True
    
    If picInput.Visible = False Then Exit Function

    intRow = Split(picInput.Tag, ";")(0)
    intCol = Split(picInput.Tag, ";")(1)
        
    If Index = 0 Then
        '����ѹ
        dbMin = Val(mItemStru(intRow).��Сֵ)
        dbMax = Val(mItemStru(intRow).���ֵ)
        strName = "����ѹ"
    Else
        '
        dbMin = Val(mItemOtherStru(1).��Сֵ)
        dbMax = Val(mItemOtherStru(1).���ֵ)
        strName = "����ѹ"
    End If
            
    If Trim(txtInput(Index).Text) <> "" And (Val(txtInput(Index).Text) > dbMax And dbMax <> 0 Or Val(txtInput(Index).Text) < dbMin And dbMin) And mItemStru(intRow).�������� = 0 Then
        
        picInput.Visible = True
        mstrSQL = "������ֵ������" & strName & "��������Χ��" & dbMin & "��" & dbMax
        ShowSimpleMsg mstrSQL

        CheckBlood = False
        Exit Function
    End If
    
    
End Function

Private Sub txtInput_Validate(Index As Integer, Cancel As Boolean)
    Dim aryValue() As String
    Dim intRewrite As Integer
    Dim aryBeforeValue() As String
    Dim strInput As String
    Dim dbMin As Double
    Dim dbMax As Double
    Dim strName As String
    Dim strTmp As String
    Dim intPos As Integer
    
    If txtInput(Index).Visible = False Then Exit Sub
    If mvarEdit = False Or mshDownTab.Tag = "" Then Exit Sub
    
    Cancel = Not StrIsValid(txtInput(Index).Text, txtInput(Index).MaxLength)
    If Cancel Then
        On Error Resume Next
        txtInput(Index).SetFocus
        Exit Sub
    End If

    On Error GoTo ErrHead

    With mshDownTab

        If picInput.Visible = False Then Exit Sub

        intRow = Split(picInput.Tag, ";")(0)
        intCol = Split(picInput.Tag, ";")(1)
        
        dbMin = Val(.TextMatrix(intRow, 3))
        dbMax = Val(.TextMatrix(intRow, 2))
        strName = mItemStru(intRow).��Ŀ����
            
        Select Case Val(mshDownTab.RowData(mshDownTab.Row))
        Case mItemNo.Ѫѹ
            If Index = 0 Then
                '����ѹ
                dbMin = Val(mItemStru(intRow).��Сֵ)
                dbMax = Val(mItemStru(intRow).���ֵ)
                strName = "����ѹ"
            Else
                '
                dbMin = Val(mItemOtherStru(1).��Сֵ)
                dbMax = Val(mItemOtherStru(1).���ֵ)
                strName = "����ѹ"
            End If
        Case mItemNo.���
            If Index = 0 Then
                
                txtInput(Index).Text = UCase(txtInput(Index).Text)
                
                strTmp = txtInput(Index).Text
                
                If strTmp <> "" Then
                    If Check�Ƿ����(strTmp, "0123456789+/E*") = False Then
                        txtInput(Index).Text = ""
                    Else
                        intPos = InStr(strTmp, "E")
                        
                        If intPos > 0 Then
                            If Right(strTmp, 1) <> "E" Then
                                txtInput(Index).Text = ""
                            Else
                                If InStr(Mid(strTmp, 1, intPos - 1), "E") > 0 And intPos > 1 Then
                                    txtInput(Index).Text = ""
                                End If
                            End If
                            
                            intPos = InStr(strTmp, "/")
                            If intPos > 0 Then
                                If InStr(Mid(strTmp, 1, intPos - 1), "/") > 0 Then
                                    txtInput(Index).Text = ""
                                ElseIf InStr(Mid(strTmp, intPos + 1), "/") > 0 Then
                                    txtInput(Index).Text = ""
                                End If
                            End If
                            
                        ElseIf InStr(strTmp, "*") > 0 Then
                            If strTmp <> "*" Then
                                txtInput(Index).Text = ""
                            End If
                        End If
                        
                        If strTmp = "/E" Then txtInput(Index).Text = ""
                    End If
                End If
            End If
            
        Case mItemNo.��Һ
            If Index = 0 Then
            
                txtInput(Index).Text = UCase(txtInput(Index).Text)
                strTmp = txtInput(Index).Text
                
                If strTmp <> "" Then
                    If Check�Ƿ����(strTmp, "0123456789/C") = False Then
                        txtInput(Index).Text = ""
                    Else
                    
                        intPos = InStr(strTmp, "/C")
                        If intPos > 0 Then
                        
                            If strTmp = "/C" Then
                            
                                txtInput(Index).Text = ""
                                
                            ElseIf InStr(Mid(strTmp, 1, intPos - 2), "/") > 0 And intPos > 2 Then
                                
                                txtInput(Index).Text = ""
                                
                            ElseIf InStr(Mid(strTmp, 1, intPos - 2), "C") > 0 And intPos > 2 Then
                                
                                txtInput(Index).Text = ""
                                
                            ElseIf Right(strTmp, 2) <> "/C" Then
                                
                                txtInput(Index).Text = ""
                            
                            End If
                        ElseIf InStr(strTmp, "C") > 0 Then
                            If strTmp <> "C" Then
                                txtInput(Index).Text = ""
                            End If
                        End If
                    End If
                End If
            End If
        End Select
        
        If CheckStrType(txtInput(Index).Text, 99, "0123456789.") Then
            If Trim(txtInput(Index).Text) <> "" And (Val(txtInput(Index).Text) > dbMax And dbMax <> 0 Or Val(txtInput(Index).Text) < dbMin And dbMin) And mItemStru(intRow).�������� = 0 Then
                Cancel = True
                
                picInput.Visible = True
                mstrSQL = "������ֵ������" & strName & "��������Χ��" & dbMin & "��" & dbMax
                ShowSimpleMsg mstrSQL
                txtInput(Index) = ""
                txtInput(Index).SetFocus
                mshDownTab.SetFocus
                
                Exit Sub
            End If
            
            If CheckNumber(Val(txtInput(Index).Text), mItemStru(intRow).���ݳ���, mItemStru(intRow).С��λ��) = False Then
                ShowSimpleMsg "��" & strName & "��������λ�:" & mItemStru(intRow).���ݳ��� - mItemStru(intRow).С��λ�� & "��С��λ�:" & mItemStru(intRow).С��λ��
                mshDownTab.SetFocus
                Exit Sub
            End If
        End If
        
        '��д��Ԫ��ֵ
        If mItemStru(intRow).�������� = 0 Then
            If Val(mshDownTab.RowData(mshDownTab.Row)) = mItemNo.Ѫѹ Then
                strInput = txtInput(0).Text & "/" & txtInput(1).Text
                If strInput = "/" Then strInput = ""
            Else
                If Trim(txtInput(Index).Text) = "" Then
                    strInput = ""
                Else
                    If CheckStrType(txtInput(Index).Text, 99, "0123456789.") Then
                        If mItemStru(intRow).С��λ�� > 0 Then
                            strInput = Format(Val(txtInput(Index).Text), "0." & String(mItemStru(intRow).С��λ��, "0"))
                        Else
                            strInput = Format(Val(txtInput(Index).Text), "0")
                        End If
                    Else
                        strInput = txtInput(Index).Text
                    End If
                    
                End If
            End If
        ElseIf mItemStru(intRow).�������� = 1 Then
            strInput = Trim(txtInput(Index).Text)
        End If
        
        picInput.Visible = False
        
        Call WriteDownTab(intRow, intCol, strInput)
        
    End With
    
    Exit Sub
    
    '------------------------------------------------------------------------------------------------------------------
ErrHead:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub UserControl_GotFocus()
    RaiseEvent Activate

End Sub

Private Sub UserControl_Initialize()
    mstr���²�λ = "Ҹ��"
    mstr������ʽ = "��������"
    mstr���� = ""
    Call InitCommandBar
End Sub

Private Sub vsb_Change()
    On Error Resume Next
    
'    pic.Top = 0 - vsb.Value * 300
    pic.Top = 0 - vsb.Value * msinVStep
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    Dim intRewrite As Integer
    Dim strInput As String
    
    '�������ͣ�2-ȫ���������ĵ�,3-�޸ĵĵ㣺���ܰ���ԭ�еĵ�������ĵ�,4-ɾ���ĵ�
    
    strInput = vsf.TextMatrix(Row, Col)
     
    intRewrite = Val(vsf.ColData(Col))
    
    If Trim(strInput) <> "" Then
        '���ӻ��޸Ĳ���
        Select Case intRewrite
        Case 0
            vsf.ColData(Col) = 2
        Case 1
            vsf.ColData(Col) = 3
        Case 2
            vsf.ColData(Col) = 2
        Case 3
            vsf.ColData(Col) = 3
        Case 4
            vsf.ColData(Col) = 3
        End Select
    Else
        vsf.ColData(Col) = 4
    End If
    
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strInfo As String
    
    On Error Resume Next
    
    If NewCol <= 1 Then Exit Sub
    

    strInfo = "��" & vsf.TextMatrix(NewRow, 1) & "����Ŀ��Χ��" & mItemOtherStru(0).��Сֵ & "��" & mItemOtherStru(0).���ֵ

    
    RaiseEvent PromptInfo(strInfo)
End Sub

Private Sub vsf_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub vsf_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub vsf_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    Dim aryValue() As String
    Dim strStart As String
    Dim strEnd As String
    Dim strFrom As String
    Dim strTo As String
    Dim strTmp As String
    
    On Error Resume Next
    
    If Val(mrsParam("�༭")) = 0 Then
        vsf.EditMode(NewCol) = 0
        Exit Sub
    End If
    If picScale.Tag = "" Then
        vsf.EditMode(NewCol) = 0
        Exit Sub
    End If
    
    '�����С���ʱ�䷶Χ
    
    strFrom = Split(picScale.Tag, ";")(0)
    strTo = Split(picScale.Tag, ";")(1)
    
    strTmp = GetCurveDateTime(NewCol - vsf.FixedCols + 1, CDate(strFrom), mlngHourBegin)
    strStart = Split(strTmp, ",")(0)
    strEnd = Split(strTmp, ",")(1)
    
    If (strStart >= strFrom And strStart <= strTo) Or (strEnd >= strFrom And strEnd <= strTo) Then
        vsf.EditMode(NewCol) = 1
    Else
        vsf.EditMode(NewCol) = 0
    End If
    
End Sub

Private Sub vsf_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack Then
        If CheckStrType(Chr(KeyAscii), 99, "0123456789") = False Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub vsf_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then        'Delete�����Ԫ
        If vsf.Body.Editable = flexEDKbdMouse And vsf.Row = 1 And vsf.Col > 1 Then
            vsf.TextMatrix(vsf.Row, vsf.Col) = ""
        End If
    End If
End Sub

Private Sub vsf_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    Call vsf_BeforeRowColChange(0, 0, Row, Col, False)
    If vsf.EditMode(Col) = 0 Then Cancel = True
    
End Sub

Private Sub vsf_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If vsf.EditText <> "" Then
        If Val(vsf.EditText) < mItemOtherStru(0).��Сֵ Or Val(vsf.EditText) > mItemOtherStru(0).���ֵ Then
            vsf.EditText = ""
            Cancel = True
        End If
        
    End If
End Sub

Private Sub Get��Ժ���ʱ��()
    Dim rsCheck As New ADODB.Recordset
    On Error GoTo errHand
    '���ݵķ���ʱ�䲻��С�����ʱ��,������С����Ժʱ��
    
    gstrSQL = " Select MIN(��ʼʱ��) AS ʱ�� From ���˱䶯��¼ Where ����ID=[1] And ��ҳID=[2] And ����ID=[3]"
    Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "", CLng(mrsParam!����ID), CLng(mrsParam!��ҳID), CLng(mrsParam!����ID))
    mstr��Сʱ�� = Format(DateAdd("n", 1, rsCheck!ʱ��), "yyyy-MM-dd HH:mm:00")    '��Ժʱ���п��ܴ�������,���λ����
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub OutputDadaForDebug()
    Dim strRow As String
    Dim intRow As Integer, intCol As Integer
    Dim intRows As Integer, intCols As Integer
    '�������ڼ�¼������
    
    intRows = mshScale.Rows - 1
    intCols = mshScale.Cols - 1
    
    For intRow = 0 To intRows
        strRow = ""
        For intCol = 0 To intCols
            strRow = strRow & "," & mshScale.TextMatrix(intRow, intCol)
        Next
        'Debug.Print "Row:" & intRow & Space(4) & Mid(strRow, 2)
    Next
End Sub
