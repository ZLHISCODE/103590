VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmCaseTendBodyData 
   AutoRedraw      =   -1  'True
   Caption         =   "�������ݱ༭"
   ClientHeight    =   7410
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13650
   Icon            =   "frmCaseTendBodyData.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7410
   ScaleWidth      =   13650
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picNull 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   855
      ScaleHeight     =   825
      ScaleWidth      =   9975
      TabIndex        =   44
      Top             =   3540
      Width           =   10005
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ǰ�ޱ����Ŀ,���������Ŀ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1665
         TabIndex        =   45
         Top             =   270
         Width           =   6315
      End
   End
   Begin VB.PictureBox picTab 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   6375
      Left            =   720
      ScaleHeight     =   6345
      ScaleWidth      =   9465
      TabIndex        =   0
      Top             =   720
      Width           =   9495
      Begin VB.PictureBox picSplitTab 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   30
         Left            =   0
         MousePointer    =   7  'Size N S
         ScaleHeight     =   30
         ScaleWidth      =   6255
         TabIndex        =   22
         Top             =   3600
         Width           =   6255
      End
      Begin VB.PictureBox PicLst 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   8100
         ScaleHeight     =   1425
         ScaleWidth      =   1185
         TabIndex        =   37
         Top             =   2175
         Visible         =   0   'False
         Width           =   1215
         Begin VB.ListBox lstSelect 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            Index           =   0
            ItemData        =   "frmCaseTendBodyData.frx":6852
            Left            =   0
            List            =   "frmCaseTendBodyData.frx":685F
            TabIndex        =   39
            Top             =   855
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtLst 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   -30
            MultiLine       =   -1  'True
            TabIndex        =   38
            Top             =   270
            Width           =   1215
         End
         Begin VB.Label lbllst 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000018&
            Caption         =   "¼�룺"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   2
            Left            =   30
            TabIndex        =   41
            Top             =   30
            Width           =   540
         End
         Begin VB.Label lbllst 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000018&
            Caption         =   "ѡ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   3
            Left            =   15
            TabIndex        =   40
            Top             =   615
            Width           =   540
         End
      End
      Begin VB.PictureBox picEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   8160
         ScaleHeight     =   255
         ScaleWidth      =   1305
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   3600
         Visible         =   0   'False
         Width           =   1335
         Begin VB.PictureBox picHour 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   -10
            ScaleHeight     =   255
            ScaleWidth      =   465
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   -10
            Visible         =   0   'False
            Width           =   495
            Begin VB.TextBox txtHour 
               Alignment       =   2  'Center
               BackColor       =   &H80000018&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   0
               MaxLength       =   2
               TabIndex        =   34
               Top             =   15
               Width           =   315
            End
            Begin VB.Label lblHour 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "h"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   180
               Left            =   345
               TabIndex        =   35
               Top             =   45
               Width           =   105
            End
         End
         Begin VB.TextBox txtEdit 
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   240
            TabIndex        =   32
            Top             =   0
            Width           =   800
         End
         Begin VB.CommandButton cmdColor 
            Caption         =   "�E"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1080
            TabIndex        =   31
            Top             =   30
            Width           =   285
         End
         Begin VB.Label lblCheck 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   36
            Top             =   0
            Visible         =   0   'False
            Width           =   255
         End
      End
      Begin VB.ListBox lstSelect 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1080
         Index           =   1
         ItemData        =   "frmCaseTendBodyData.frx":6878
         Left            =   8100
         List            =   "frmCaseTendBodyData.frx":6885
         Style           =   1  'Checkbox
         TabIndex        =   29
         Top             =   900
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   6120
         ScaleHeight     =   1425
         ScaleWidth      =   2025
         TabIndex        =   25
         Top             =   4200
         Width           =   2055
         Begin zl9TemperatureChartJX.ColorPicker usrColor 
            Height          =   2190
            Left            =   0
            TabIndex        =   26
            Top             =   -720
            Width           =   2190
            _ExtentX        =   3863
            _ExtentY        =   3863
         End
      End
      Begin VB.Frame fraTabDetail 
         Height          =   2415
         Left            =   0
         TabIndex        =   24
         Top             =   3720
         Width           =   9495
         Begin VSFlex8Ctl.VSFlexGrid vsfTabDetail 
            Height          =   1815
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   8775
            _cx             =   15478
            _cy             =   3201
            Appearance      =   1
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
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483632
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   3
            HighLight       =   0
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   2
            RowHeightMin    =   0
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
      Begin VB.Frame fraTable 
         Height          =   3495
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Width           =   9495
         Begin VSFlex8Ctl.VSFlexGrid vsfTab 
            Height          =   2775
            Left            =   240
            TabIndex        =   27
            Top             =   360
            Width           =   8415
            _cx             =   14843
            _cy             =   4895
            Appearance      =   1
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
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   0
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   12
            FixedRows       =   1
            FixedCols       =   6
            RowHeightMin    =   0
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
      Begin VB.Label lbllst 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   75
         Index           =   1
         Left            =   8880
         TabIndex        =   43
         Top             =   2760
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Label lbllst 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   0
         Left            =   8880
         TabIndex        =   42
         Top             =   3240
         Visible         =   0   'False
         Width           =   45
      End
   End
   Begin VB.PictureBox picCurve 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H80000008&
      Height          =   5895
      Left            =   120
      ScaleHeight     =   5865
      ScaleWidth      =   9705
      TabIndex        =   1
      Top             =   720
      Width           =   9735
      Begin VB.PictureBox picSplit 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   30
         Left            =   0
         MousePointer    =   7  'Size N S
         ScaleHeight     =   30
         ScaleWidth      =   6255
         TabIndex        =   11
         Top             =   3720
         Width           =   6255
      End
      Begin VB.PictureBox picδ�� 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   6000
         ScaleHeight     =   1185
         ScaleWidth      =   1545
         TabIndex        =   15
         Top             =   960
         Width           =   1575
         Begin VB.ListBox lstδ�� 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            Height          =   1290
            ItemData        =   "frmCaseTendBodyData.frx":689E
            Left            =   0
            List            =   "frmCaseTendBodyData.frx":68A0
            TabIndex        =   16
            Top             =   0
            Width           =   2055
         End
      End
      Begin VB.PictureBox picValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   3480
         ScaleHeight     =   1425
         ScaleWidth      =   1665
         TabIndex        =   13
         Top             =   -120
         Width           =   1695
         Begin zl9TemperatureChartJX.ColorPicker usrValue 
            Height          =   2190
            Left            =   -240
            TabIndex        =   14
            Top             =   -360
            Width           =   2190
            _ExtentX        =   3863
            _ExtentY        =   3863
         End
      End
      Begin VB.Frame fraDetail 
         Height          =   1935
         Left            =   0
         TabIndex        =   10
         Top             =   3840
         Width           =   7815
         Begin zl9TemperatureChartJX.VsfGrid vsfDetail 
            Height          =   1335
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   2355
         End
      End
      Begin VB.Frame fraData 
         Height          =   2895
         Left            =   0
         TabIndex        =   9
         Top             =   600
         Width           =   7335
         Begin zl9TemperatureChartJX.VsfGrid vsfCurve 
            Height          =   2535
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   4471
         End
      End
      Begin VB.Frame fraTime 
         Height          =   615
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   9495
         Begin VB.PictureBox picToolBar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   350
            Left            =   120
            ScaleHeight     =   345
            ScaleWidth      =   2775
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   210
            Width           =   2775
            Begin VB.OptionButton OptTime 
               Caption         =   "4"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   0
               Left            =   480
               Style           =   1  'Graphical
               TabIndex        =   6
               Top             =   0
               Width           =   350
            End
            Begin VB.Label lblPtime 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ʱ��:"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Left            =   0
               TabIndex        =   7
               Top             =   45
               Width           =   450
            End
         End
         Begin VB.Label lblTime 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "00:00��05:59"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   3000
            TabIndex        =   8
            Top             =   240
            Width           =   1080
         End
      End
      Begin MSComctlLib.ImageList ImageList1 
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
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCaseTendBodyData.frx":68A2
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ilsDate 
      Left            =   7080
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaseTendBodyData.frx":6C3C
            Key             =   "preGreen"
            Object.Tag             =   "preGreen"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaseTendBodyData.frx":764E
            Key             =   "preGray"
            Object.Tag             =   "preGray"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaseTendBodyData.frx":8060
            Key             =   "nextGreen"
            Object.Tag             =   "nextGreen"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaseTendBodyData.frx":8A72
            Key             =   "nextGray"
            Object.Tag             =   "nextGray"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaseTendBodyData.frx":9484
            Key             =   "preLight"
            Object.Tag             =   "preLight"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaseTendBodyData.frx":9E96
            Key             =   "nextLight"
            Object.Tag             =   "nextLight"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picDate 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   8400
      ScaleHeight     =   360
      ScaleWidth      =   2760
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   0
      Width           =   2760
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Left            =   30
         TabIndex        =   21
         Top             =   30
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   132907011
         CurrentDate     =   42285
      End
      Begin VB.Image imgbtn 
         Height          =   240
         Index           =   1
         Left            =   1440
         Picture         =   "frmCaseTendBodyData.frx":A8A8
         Stretch         =   -1  'True
         Top             =   0
         Width           =   255
      End
      Begin VB.Image imgbtn 
         Height          =   240
         Index           =   0
         Left            =   1920
         Picture         =   "frmCaseTendBodyData.frx":B2AA
         Top             =   0
         Width           =   240
      End
      Begin VB.Image imgDefault 
         Height          =   255
         Left            =   600
         Top             =   840
         Width           =   255
      End
   End
   Begin VB.Timer tmrData 
      Interval        =   60
      Left            =   10080
      Top             =   840
   End
   Begin VB.PictureBox picStb 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      FillColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   8640
      ScaleHeight     =   360
      ScaleWidth      =   2415
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   6720
      Width           =   2415
      Begin VB.Label lblStb 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   0
         Width           =   75
      End
   End
   Begin MSComctlLib.ImageList ilsDetail 
      Left            =   7680
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaseTendBodyData.frx":BCAC
            Key             =   "detele"
            Object.Tag             =   "detele"
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.TabControl tbcThis 
      Height          =   6000
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   5055
      _Version        =   589884
      _ExtentX        =   8916
      _ExtentY        =   10583
      _StockProps     =   64
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   7050
      Width           =   13650
      _ExtentX        =   24077
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmCaseTendBodyData.frx":1250E
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   21167
            Key             =   "ZLNOTE"
            Object.ToolTipText     =   "��Ϣ��ʾ��Ϣ"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   2
            MinWidth        =   2
            Text            =   "��������"
            TextSave        =   "��������"
            Key             =   "ZLDataType"
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
   Begin MSComctlLib.ImageList ilstab 
      Left            =   6480
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaseTendBodyData.frx":12DA2
            Key             =   "mark"
            Object.Tag             =   "mark"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmCaseTendBodyData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Const mFontSize As Integer = 9 '���������ʼ��СΪ9������
Private mblnTemType As Boolean  'TRUEΪר�����µ�,FALSEΪ��׼���µ�
Private Enum TYPE_Curve
    COL_Null = 0
    COL_�༭ = 1
    COL_������ = 2
    COL_�ַ��� = 3
    COL_��Ŀ��� = 4
    COL_��Ŀ�� = 5
    col_ԭʼʱ�� = 6
    COL_�޸�״̬ = 7
    COL_��Ŀ���� = 8
    COL_��ʾ = 9
    COL_ʱ�� = 10
    COL_ԭֵ = 11
    COL_���� = 12
    COL_��ɫ = 13
    COL_���Ժϸ� = 14
    COL_��λ = 15
    Col_δ��˵�� = 16
    COL_��Դ = 17
    COL_������Դ = 18 '��ϸ�����ʾ��Դ
    COL_ɾ�� = 19
End Enum

Private Enum TYPE_Tab
    COL_TabNull = 0
    COL_tab�ַ��� = 1
    COL_tab��Ŀ��� = 2
    col_tabԭʼʱ�� = 3
    COL_tab��Ŀ�� = 4  '--��������λ
    COL_tab��Ŀ���� = 5 '- -������λ
    COL_tabDirect = 6
End Enum

Private Type Type_Item
    ���� As String
    ֵ�� As String
    ��Ŀ���� As Integer
    ��ĿС�� As Double
    ��¼Ƶ�� As Integer
    ��Ŀ��ʾ As Integer
    ��Ŀ���� As Integer
    ��Ŀ���� As Long
    ��λ As String
    ��Ŀ�� As Long
    ��Ŀ�� As String
    ��¼�� As String
    ��Ժ�ײ� As Integer
End Type

Private Type type_Patient
    lng����ID As Long
    lng��ҳID As Long
    lng�ļ�ID As Long
    lngӤ�� As Long
    lng����ID As Long
    lng����ȼ� As Long
    lng����ID As Long
    lng��ʽID As Long
End Type
Private mT_Patient As type_Patient

Private Type Type_OptRow
    �ϱ� As Integer
    �±� As Integer
End Type

Private mOptRow As Type_OptRow
    

'������:
Private mcbrToolBar As CommandBar

Private mblnStart As Boolean
Private mblnMove As Boolean
Private mblnInit As Boolean
Private mblnEdit As Boolean
Private mblnOK As Boolean
Private mblnScroll As Boolean
Private mblnResize As Boolean
Private mblnAllRefresh As Boolean
Private mint����Ӧ�� As Integer
Private mblnEdit���� As Boolean
Private mblnFileBack As Boolean  '�ļ��Ƿ�鵵
Private mbln��Ժ As Boolean '���˳�Ժ���ļ��Ѿ�����ΪTRUE
Private mbln¼��Сʱ As Boolean  'ȫ�������ʾ¼��ʱ��
Private mbln����������ʾ As Boolean '�����Ƿ���(����/����)��ʽ¼��
Private mblnRefresh���� As Boolean
Private mbln���ܵ��� As Boolean
Private mstrCurveItem As String  'ר�����µ���������Ŀ��Ϣ
Private mstrActiveItem As String '���µ����Ŀ��Ϣ
Private mstrOverDate As String '����ʵ�ʳ�Ժʱ��(�����µ�ʵ����ֹʱ��)
Private mstrBegin As String 'ĳ��ʱ���Ŀ�ʼ�ͽ���ʱ�� 00:00-05:59
Private mstrEnd As String
Private mstrDate As String '���µ���ǰҳ�ĵ�һ��ʱ��
Private mstrBTime As String  '���µ��Ŀ�ʼʱ��ͽ���ʱ��
Private mstrETime As String
Private mstrPreOutDate As String '����Ԥ��Ժʱ��
Private mstrSQL As String
Private mstrδ��˵�� As String
Private mintBigSize As Integer '�Ƿ�Ŵ�
Private mintPreDays As Integer '����¼��ʱ��
Private mlngHours As Long '���ݲ�¼ʱ��
Private marrTime() As String

'��¼��
Private mrsPart As New ADODB.Recordset '���²�λ��¼��
Private mrsNote As New ADODB.Recordset '���±����ݼ�
Private mrsCurve As New ADODB.Recordset  '�����������ݼ�¼��
Private mrsTable As New ADODB.Recordset  '���±�����ʾ�ݼ�
Private mrsTableDetail As New ADODB.Recordset '���±���������ϸ���ݼ�
Private mrsRecodeID As New ADODB.Recordset  '��¼id���ݼ�

Public Function ShowEditor(ByVal frmParent As Object, ByVal strParam As String, ByVal strTime As String, ByVal strDayTime As String, _
    ByVal int����Ӧ�� As Integer, Optional blnMove As Boolean = False, Optional ByVal bytSize As Byte = 0) As Boolean
'----------------------------------------------------------------------------------------------------------------------------------------------------------
'����:�������µ��༭����
'����:frmParent ������,strParam ��ʽ:����ID;��ҳId;�ļ�ID;Ӥ��;����ID;������ȼ�  strTime ĳ��ʱ���ʱ�䷶Χ ����:2011-01-25 00:00:00;2011-01-25 05:59:59

'     strDayTime һ�ܿ�ʼʱ��; int����Ӧ��=2 ��ʾ���������ʹ��� blnMove ��ʷ�����Ƿ�ת��
'     bytSize 0-9������ 1-12������
'----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim arrParam() As String
    Dim blnShowing As Boolean
    
    If strParam = "" Then Exit Function
    arrParam = Split(strParam, ";")
    If UBound(arrParam) < 3 Then Exit Function
        
    mblnStart = True
    mblnMove = False
    mblnInit = False
    mblnEdit = False
    mblnOK = False
    mblnResize = False
    mblnAllRefresh = False
    mbln���ܵ��� = False
    mstrOverDate = ""
    
    mT_Patient.lng����ID = 0
    mT_Patient.lng����ȼ� = 3
    
    mT_Patient.lng����ID = Val(arrParam(0))
    mT_Patient.lng��ҳID = Val(arrParam(1))
    mT_Patient.lng�ļ�ID = Val(arrParam(2))
    mT_Patient.lngӤ�� = Val(arrParam(3))
    
    If UBound(arrParam) > 3 Then mT_Patient.lng����ID = arrParam(4)
    If UBound(arrParam) > 4 Then mT_Patient.lng����ȼ� = arrParam(5)
    
    If mT_Patient.lng����ID = 0 And mT_Patient.lng��ҳID = 0 And mT_Patient.lng����ID = 0 Then
        MsgBox "�ļ�ID,����ID,��ҳID����Ϊ��,����!", vbInformation, gstrSysName
        Exit Function
    End If
    
    If Not OpenPatientInfo Then Exit Function
    
    mstrBegin = Format(Split(strTime, ";")(0), "YYYY-MM-DD HH:mm:ss")
    mstrEnd = Format(Split(strTime, ";")(1), "YYYY-MM-DD HH:mm:ss")
    mstrDate = strDayTime
    
    If Not ChekPatientOut(mT_Patient.lng�ļ�ID, mT_Patient.lng����ID, mT_Patient.lng��ҳID, mT_Patient.lngӤ��) Then Exit Function
    mintBigSize = bytSize
    Me.Font.Size = IIf(mintBigSize = 0, 9, 12)
    mint����Ӧ�� = int����Ӧ��
    mblnEdit���� = True
    mblnMove = blnMove
    
    '��̬����ʱ�㰴ť�ؼ�
    UnLoadOptTime
    LoadOptTime
    '����ļ��Ƿ�鵵
    mblnFileBack = CheckFileBack(mT_Patient.lng�ļ�ID, mblnMove)
    '��ʼ��������
    Call InitCommandBars
    '��ʼ�����
    Call GetTableRowName
    '��������
    Call zlRefreshData
    mblnInit = True
    mblnResize = True
    Me.Show 1
    
    ShowEditor = mblnOK
End Function

Public Function OpenPatientInfo() As Boolean
'------------------------------------------------------
'��ȡ���˻�����Ϣ
'------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo Errhand
    '��ȡ���˿���
    strSQL = " select ��Ժ����id from ������ҳ where ����id=[1] and ��ҳid=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mT_Patient.lng����ID, mT_Patient.lng��ҳID)
    If rsTemp.BOF = False Then
        mT_Patient.lng����ID = Val(zlStr.Nvl(rsTemp("��Ժ����ID").Value))
    End If
    
    '��ȡ���˻���ȼ�
    strSQL = "Select zl_PatitTendGrade([1],[2]) As ����ȼ� From dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mT_Patient.lng����ID, mT_Patient.lng��ҳID)
    If rsTemp.BOF = False Then
        mT_Patient.lng����ȼ� = Val(zlStr.Nvl(rsTemp("����ȼ�").Value))
    End If
    
    '��ȡ���µ�������Ϣ
    mblnTemType = False
    strSQL = "Select B.����,B.ID From ���˻����ļ� A,�����ļ��б� B Where A.��ʽID=B.ID And A.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mT_Patient.lng�ļ�ID)
    If rsTemp.BOF = False Then
        mblnTemType = (Nvl(rsTemp!����) = "1")
        mT_Patient.lng��ʽID = rsTemp!Id
    End If
    
    If mblnTemType = True Then
        gintHourBegin = T_BodyStyle.lng��ʼʱ��
    Else
        gintHourBegin = zlDatabase.GetPara("���¿�ʼʱ��", glngSys, 1255, 4)
        T_BodyStyle.lng��ʼʱ�� = gintHourBegin
        T_BodyStyle.lngʱ���� = 4
        T_BodyStyle.lng������ = 6
        T_BodyStyle.lng���� = 7
    End If
    OpenPatientInfo = True
    Exit Function
    
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



Public Function ChekPatientOut(ByVal lng�ļ�ID As Long, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal intBaby As Long) As Boolean
'-----------------------------------------------------------------------------------------------
'����:��ȡ���µ���ʼʱ��ͽ���ʱ�� ����鲡���Ƿ��Ժ
'-----------------------------------------------------------------------------------------------
    Dim strSQL As String, strNewSql As String
    Dim strBeginDate As String, strEndDate As String
    Dim rsTemp As New ADODB.Recordset
    Dim strMaxDate As String, strCurrDate As String
    Dim intDay As Integer
    mbln��Ժ = False
    On Error GoTo Errhand
    
    mintPreDays = Val(zlDatabase.GetPara("����¼�뻤����������", glngSys, 1255, "1"))
    mlngHours = Val(Mid(Val(zlDatabase.GetPara("���ݲ�¼ʱ��", glngSys)), 1, 6))
    mbln���ܵ��� = (Val(zlDatabase.GetPara("���ܲ�����ʾ��������", glngSys, 1255, 0)) = 1)
    mbln¼��Сʱ = (Val(zlDatabase.GetPara("ȫ�������ʾ¼��ʱ��", glngSys, 1255, 0)) = 1)
    mbln����������ʾ = (Val(zlDatabase.GetPara("���������(����/����)��ʽ¼��", glngSys, 1255, 0)) = 1)
    If mintPreDays < 0 Then mintPreDays = 0
    
    '��ȡ����Ԥ��Ժʱ��
    strSQL = "Select ��ʼʱ�� From ���˱䶯��¼ where ����ID=[1] and ��ҳID=[2] And ��ʼԭ��=10"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlPrint", lng����ID, lng��ҳID)
    If Not rsTemp.EOF Then mstrPreOutDate = Format(rsTemp!��ʼʱ��, "YYYY-MM-DD HH:mm:ss")
    
    '��ȡӤ��ҽ����Ϣ(ת�ƣ���Ժ),����ҽ����ҽ����ϢΪ׼��������ĸ�׳�Ժ����Ϊ׼
    strNewSql = "(SELECT " & vbNewLine & _
                "        ����ID, ��ҳID, Ӥ��ʱ��, DECODE(NVL(Ӥ��, 0), 0, DECODE(NVL(��Ժ����, ''), '', 0, 1), DECODE(NVL(Ӥ��ʱ��, ''), '', 0, 1)) ��¼" & vbNewLine & _
                "       FROM (SELECT A.����ID, A.��ҳID, B.��ʼִ��ʱ�� Ӥ��ʱ��, A.��Ժ����, B.Ӥ��" & vbNewLine & _
                "              FROM ������ҳ A," & vbNewLine & _
                "                   (SELECT B.����ID, B.��ҳID, B.Ӥ��, ��ʼִ��ʱ��" & vbNewLine & _
                "                     FROM ����ҽ����¼ B, ������ĿĿ¼ C" & vbNewLine & _
                "                     WHERE B.������ĿID + 0 = C.ID AND B.ҽ��״̬ = 8 AND NVL(B.Ӥ��, 0) <> 0 AND B.������� = 'Z' " & vbNewLine & _
                "                      AND Instr(',3,5,11,', ',' || c.�������� || ',') > 0 AND B.����ID = [2] AND B.��ҳID = [3] AND B.Ӥ��(+) = [4]) B" & vbNewLine & _
                "              WHERE A.����ID = [2] AND A.��ҳID = [3] AND A.����ID = B.����ID(+) AND A.��ҳID = B.��ҳID(+)" & vbNewLine & _
                "              ORDER BY B.��ʼִ��ʱ�� DESC)" & vbNewLine & _
                "       WHERE ROWNUM < 2) E"

    '˵��:Ŀǰ����ר�����µ������˿���ͬʱ���ڶ�����µ������µ���ʼʱ�����ֹʱ��Ĺ�������:
    '����ļ��Ŀ�ʼʱ�䲻Ϊ�ղ��Ҵ��ڵ��ڲ�����Ժʱ���Ӥ������ʱ��,���µ��Ŀ�ʼʱ�����ļ���ʼʱ��Ϊ׼,�����Բ�����Ժʱ���Ӥ������ʱ��Ϊ׼
    '����ļ�����ֹʱ�䲻Ϊ�ղ���С�ڵ��ڲ��˻�Ӥ����Ժʱ�䣨δ��Ժ���ܴ��ڵ�ǰʱ�䣩,���µ�����ʱ�����ļ���ʼʱ��Ϊ׼���������µ�����ʱ���Բ��˻�Ӥ����Ժʱ��Ϊ׼(δ��ԺΪ��ǰʱ��)
    '����ļ�����ֹʱ��Ϊ��,����ԭ�з�ʽ,��������Ѿ���Ժ�����ѳ�Ժʱ��Ϊ׼,δ��Ժ���ѵ�ǰʱ������ݽ���ʱ��Ϊ׼.
    strSQL = " SELECT  DECODE(D.��ʼʱ��,NULL,DECODE(B.����ʱ��, NULL, A.��ʼ, B.����ʱ��)," & vbNewLine & _
            "               DECODE(SIGN(D.��ʼʱ�� - DECODE(B.����ʱ��, NULL, A.��ʼ, B.����ʱ��))," & vbNewLine & _
            "                      1," & vbNewLine & _
            "                      D.��ʼʱ��," & vbNewLine & _
            "                      DECODE(B.����ʱ��, NULL, A.��ʼ, B.����ʱ��))) AS ��ʼ," & vbNewLine & _
            "       DECODE(D.����ʱ��," & vbNewLine & _
            "               NULL," & vbNewLine & _
            "               DECODE(E.��¼," & vbNewLine & _
            "                      0," & vbNewLine & _
            "                      DECODE(SIGN(NVL(E.Ӥ��ʱ��, A.��ֹ) - D.����ʱ��), 1, NVL(E.Ӥ��ʱ��, A.��ֹ), D.����ʱ��)," & vbNewLine & _
            "                      NVL(E.Ӥ��ʱ��, A.��ֹ))," & vbNewLine & _
            "               DECODE(SIGN(NVL(E.Ӥ��ʱ��, A.��ֹ) - D.����ʱ��), 1, D.����ʱ��, NVL(E.Ӥ��ʱ��, A.��ֹ))) ��ֹ," & vbNewLine & _
            "       DECODE(D.����ʱ��, NULL, E.��¼, 1) ��¼" & vbNewLine & _
            " FROM (SELECT ����ID, ��ҳID, MIN(��ʼʱ��) AS ��ʼ, MAX(NVL(��ֹʱ��, SYSDATE)) AS ��ֹ" & vbNewLine & _
            "       FROM ���˱䶯��¼" & vbNewLine & _
            "       WHERE ��ʼʱ�� IS NOT NULL AND ����ID = [2] AND ��ҳID = [3]" & vbNewLine & _
            "       GROUP BY ����ID, ��ҳID) A," & vbNewLine & _
            "     (SELECT ����ID, ��ҳID, ����ʱ�� FROM ������������¼ WHERE ����ID = [2] AND ��ҳID = [3] AND ��� = [4]) B," & vbNewLine & _
            "     (SELECT NVL(����ʱ��, SYSDATE) ����ʱ��, ��ʼʱ��, ����ʱ��" & vbNewLine & _
            "       FROM (SELECT MAX(B.����ʱ��) ����ʱ��, MAX(A.��ʼʱ��) ��ʼʱ��, MAX(A.����ʱ��) ����ʱ��" & vbNewLine & _
            "              FROM ���˻����ļ� A, ���˻������� B" & vbNewLine & _
            "              WHERE A.ID = B.�ļ�ID(+) AND A.ID = [1] AND A.����ID = [2] AND A.��ҳID = [3] AND A.Ӥ�� = [4])) D," & vbNewLine & _
            "  " & strNewSql & vbNewLine & _
            " WHERE A.����ID = E.����ID AND A.��ҳID = E.��ҳID AND A.����ID = B.����ID(+) AND A.��ҳID = B.��ҳID(+)"

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlPrint", lng�ļ�ID, lng����ID, lng��ҳID, intBaby)
    If rsTemp.RecordCount > 0 Then
        rsTemp.MoveFirst
        strBeginDate = Format(rsTemp!��ʼ, "YYYY-MM-DD HH:MM:SS")
        strEndDate = Format(rsTemp!��ֹ, "YYYY-MM-DD HH:MM:SS")
        mbln��Ժ = Not (Val(rsTemp!��¼) = 0)
    Else
        MsgBox "�޴˲��˱���סԺ��Ϣ,����!", vbInformation, gstrSysName '�������˱䶯��Ϣ�˳�
        Exit Function
    End If
    
    strCurrDate = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:mm:ss")

    mstrBTime = strBeginDate
    mstrOverDate = strEndDate
    mstrETime = strEndDate
    If CDate(mstrETime) < CDate(Format(strCurrDate, "YYYY-MM-DD HH:mm:ss")) And Not mbln��Ժ Then mstrETime = CDate(Format(strCurrDate, "YYYY-MM-DD HH:mm:ss"))
    If mstrBTime > mstrETime Then mstrBTime = mstrETime
    If mstrDate < mstrBTime Then mstrDate = mstrBTime
    
    '���˳�Ժ�Գ�Ժʱ��Ϊ��ֹʱ��
    If mbln��Ժ = True Then
        '��Ժʱ�����Ժʱ�������ͬһ�У��򽫳�Ժʱ�����һ�У���������:��ԺҲҪ¼�����£�
        mstrETime = Format(RetrunEndTimeNew(CDate(mstrBTime), CDate(mstrETime), gintHourBegin), "YYYY-MM-DD HH:mm:ss")
        strMaxDate = Format(mstrETime, "YYYY-MM-DD")
    Else
        intDay = mintPreDays - DateDiff("D", CDate(strCurrDate), CDate(mstrETime))
        If intDay < 0 Then intDay = 0
        strMaxDate = Format(DateAdd("d", intDay, CDate(mstrETime)), "yyyy-MM-dd")
        If CDate(mstrETime) < CDate(Format(DateAdd("d", mintPreDays, CDate(strCurrDate)), "yyyy-MM-dd HH:mm:ss")) Then
            mstrETime = Format(DateAdd("d", mintPreDays, CDate(strCurrDate)), "yyyy-MM-dd HH:mm:ss")
        End If
    End If
    
    mstrETime = Format(strMaxDate & " " & Format(mstrETime, "HH:mm:ss"), "yyyy-MM-DD HH:mm:ss")
    
    If Not (CDate(mstrBegin) >= CDate(mstrBTime) And CDate(mstrBegin) <= CDate(mstrETime)) Then
        If Int(CDate(mstrBTime)) = Int(CDate(mstrETime)) Then
            mstrBegin = Format(mstrBTime, "YYYY-MM-DD HH:mm:ss")
        Else
            mstrBegin = Format(Int(CDate(mstrETime)), "YYYY-MM-DD HH:mm:ss")
        End If
    End If
    If Not (CDate(mstrEnd) >= CDate(mstrBegin) And CDate(mstrEnd) <= CDate(mstrETime)) Then
        mstrEnd = Format(mstrETime, "YYYY-MM-DD HH:mm:ss")
    End If
    
    dtpDate.Value = Format(mstrBegin, "YYYY-MM-DD")
    dtpDate.MaxDate = Format(strMaxDate, "YYYY-MM-DD")
    dtpDate.MinDate = Format(mstrBTime, "YYYY-MM-DD")
    
    ChekPatientOut = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Function CheckFileBack(ByVal lngID As Long, ByVal blnMove As Boolean) As Boolean
'---------------------------------------------------------------
'����:����ļ��Ƿ�鵵
'---------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo Errhand
    
    CheckFileBack = False
    strSQL = "Select 1 From ���˻����ļ� Where Id=[1] And �鵵�� Is Not Null"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "����ļ��Ƿ�鵵", lngID)
    If blnMove = True Then
        strSQL = Replace(strSQL, "���˻����ļ�", "H���˻����ļ�")
    End If
    If rsTemp.RecordCount > 0 Then
        CheckFileBack = True
    End If
    
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CheckDateTime(ByVal lngRow As Long, ByVal strName As String, ByVal strTime As String) As Boolean
    '-----------------------------------------
    '���ܣ����¼������ʱ���Ƿ񳬹���¼��Χ
    '-----------------------------------------
    Dim strInfo As String
    Dim strErrMsg As String
    Dim strCurrDate As String
    Dim strText As String
    Dim strCenterTime As String
    Dim arrTime() As String
    
    On Error GoTo Errhand
    If lngRow <> 0 Then
        strInfo = "��" & lngRow & "��"
    ElseIf strName <> "" Then
        strInfo = strInfo & "[" & strName & "]"
    Else
        strInfo = ""
    End If
    strText = strTime
    
    arrTime = Split(Trim(strText), ":")
    
    If UBound(arrTime) <> 1 Then
        strInfo = "¼���ʱ���ʽ�Ƿ���[Сʱ:����]"
        Exit Function
    Else
        If Len(Trim(arrTime(0))) < 2 Then arrTime(0) = String(2 - Len(Trim(arrTime(0))), "0") & Trim(arrTime(0))
        If Len(Trim(arrTime(1))) < 2 Then arrTime(1) = String(2 - Len(Trim(arrTime(1))), "0") & Trim(arrTime(1))
        strText = arrTime(0) & ":" & arrTime(1)
    End If
    
    '�Ϸ��Լ��
    If IsNumeric(arrTime(0)) = False Or IsNumeric(arrTime(1)) = False Or Len(Trim(arrTime(0))) > 2 Or Len(Trim(arrTime(1))) > 2 Then
        lblStb.ForeColor = 255
        lblStb.Caption = "¼���ʱ���ʽ�Ƿ���[Сʱ:����]"
        Exit Function
    End If
    If Mid(strText, 3, 1) <> ":" Then
        lblStb.ForeColor = 255
        lblStb.Caption = "¼���ʱ���ʽ�Ƿ���[Сʱ:����]"
        Exit Function
    End If
    If Val(arrTime(0)) < 0 Or Val(arrTime(0)) > 23 Then
        lblStb.ForeColor = 255
        lblStb.Caption = "¼���ʱ���ʽ�Ƿ���[СʱӦ��0��23֮��]"
        Exit Function
    End If
    If Val(arrTime(1)) < 0 Or Val(arrTime(1)) > 59 Then
        lblStb.ForeColor = 255
        lblStb.Caption = "¼���ʱ���ʽ�Ƿ���[����Ӧ��0��59֮��]"
        Exit Function
    End If
    strTime = Format(dtpDate.Value & " " & strTime, "YYYY-MM-DD HH:mm:ss")
    If Format(mstrETime, "hh:mm:ss") < Split(lblTime, "��")(1) And mstrETime > mstrEnd Then mstrEnd = mstrETime
    If Not (CDate(Format(strTime)) >= CDate(mstrBegin) And CDate(strTime) <= CDate(mstrEnd)) Then
        lblStb.ForeColor = 255
        lblStb.Caption = "�������ʱ���� " & Format(mstrBegin, "hh:mm") & "��" & Format(mstrEnd, "hh:mm") & " ʱ���֮��"
        Exit Function
    End If
    
    If Not IsDate(strTime) Then Exit Function
    
    If DateDiff("m", CDate(Format(strTime, "YY-MM-DD hh:mm")), CDate(Format(mstrETime, "YY-MM-DD hh:mm"))) < 0 Then
        If mbln��Ժ = False Then
            strErrMsg = strInfo & "��¼����ʱ���ѳ�������[����¼��������" & mintPreDays & "��]��ָ���ķ�Χ!"
        Else
            strErrMsg = strInfo & "��¼����ʱ�䲻�ܴ���[���˳�Ժʱ����ļ�����ʱ�䣺" & Format(mstrETime, "YYYY-MM-DD HH:mm:ss") & "]!"
        End If
        lblStb.Caption = strErrMsg
        lblStb.ForeColor = 255
        Exit Function
    End If
    
    If DateDiff("m", CDate(Format(strTime, "YY-MM-DD hh:mm")), CDate(Format(mstrBTime, "YY-MM-DD hh:mm"))) > 0 Then
        strErrMsg = strInfo & "��¼����ʱ�䲻��С��[���µ���ʼʱ�䣺" & Format(mstrBTime, "YYYY-MM-DD HH:mm:ss") & "]!"
        lblStb.Caption = strErrMsg
        lblStb.ForeColor = 255
        Exit Function
    End If
    
    strCurrDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
    If Not IsAllowInput(mT_Patient.lng����ID, mT_Patient.lng��ҳID, mT_Patient.lngӤ��, strTime, strCurrDate) Then
        strErrMsg = strInfo & "��¼����ʱ��[" & strTime & "]����![�������ݲ�¼����Чʱ��:" & mlngHours & "Сʱ]"
        lblStb.Caption = strErrMsg
        lblStb.ForeColor = 255
        Exit Function
    End If
    
    CheckDateTime = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Function dtpDateChageDate(ByVal strValue As String) As Boolean
'------------------------------------------------------------------------------
'��¼ʱ��Ϸ�ʱ�������仯��ˢ������
'------------------------------------------------------------------------------
    Dim strErrMsg As String
    Dim strDate As String, strTime As String
    Dim i As Integer
    Dim strCurrDate As String
    Dim intBound As Integer
    Dim strBegin As String, strEnd As String
    Dim intCOl As Integer
    Dim strCurDate As String
    Dim intDay As Integer
    Dim strBTime As String
    
    On Error GoTo Errhand
    lblStb.Tag = lblStb.Caption
    
    If Format(strValue, "YYYY-MM-DD") > Format(mstrETime, "YYYY-MM-DD") Then
        If mbln��Ժ = False Then
            strErrMsg = "¼��������ѳ�������[����¼��������" & mintPreDays & "��]��ָ���ķ�Χ��"
        Else
            strErrMsg = "¼������ڲ��ܴ���[���˳�Ժʱ����ļ�����ʱ�䣺" & Format(mstrETime, "YYYY-MM-DD HH:mm:ss") & "]!"
        End If
        GoTo ErrInfo
    End If
    
    If Format(strValue, "YYYY-MM-DD") < Format(mstrBTime, "YYYY-MM-DD") Then
        strErrMsg = "¼������ڲ���С��[���µ���ʼʱ�䣺" & Format(mstrBTime, "YYYY-MM-DD HH:mm:ss") & "]��"
        GoTo ErrInfo
    End If
    
    strCurrDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
    strCurDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    If Format(strValue, "YYYY-MM-DD") = mstrETime Then
        strDate = Format(Format(mstrETime, "YYYY-MM-DD") & " 00:00:00", "YYYY-MM-DD HH:mm:ss")
        strTime = Format(mstrETime, "YYYY-MM-DD HH:mm:ss")
    ElseIf Format(strValue, "YYYY-MM-DD") = mstrBTime Then
        strDate = Format(mstrBTime, "YYYY-MM-DD HH:mm:ss")
        strTime = strDate
    Else
        strDate = Format(Format(strValue, "YYYY-MM-DD") & " 00:00:00", "YYYY-MM-DD HH:mm:ss")
        strTime = Format(Format(strValue, "YYYY-MM-DD") & " 23:59:00", "YYYY-MM-DD HH:mm:ss")
    End If
    
    If Not IsAllowInput(mT_Patient.lng����ID, mT_Patient.lng��ҳID, mT_Patient.lngӤ��, strTime, strCurrDate) Then
        strErrMsg = "¼���ʱ��[" & strValue & "]����[�������ݲ�¼����Чʱ��:" & mlngHours & "Сʱ]"
        GoTo ErrInfo
    End If
    
    mblnAllRefresh = True
    
    If UBound(marrTime) = -1 Then Call InitDateTimeRange(marrTime, gintHourBegin, T_BodyStyle.lng������, T_BodyStyle.lngʱ����)
    intDay = DateDiff("D", CDate(mstrBTime), CDate(strValue)) \ T_BodyStyle.lng����
    intDay = (intDay) * T_BodyStyle.lng����
    strBTime = Format(DateAdd("d", intDay, CDate(mstrBTime)), "yyyy-MM-dd") & " 00:00:00"
    
    If Format(strValue, "YYYY-MM-DD") = Format(strCurDate, "YYYY-MM-DD") Then
        If Format(strCurDate, "YYYY-MM-DD HH:mm:ss") < Format(strBTime, "YYYY-MM-DD HH:mm:ss") Then
             strDate = Format(strBTime, "YYYY-MM-DD HH:mm:ss")
        ElseIf Format(strCurDate, "YYYY-MM-DD HH:mm:ss") > Format(mstrETime, "YYYY-MM-DD HH:mm:ss") Then
             strDate = Format(mstrETime, "YYYY-MM-DD HH:mm:ss")
        End If
        intCOl = GetCurveColumnNew(strCurDate, strBTime, gintHourBegin)
        strDate = GetCurveDateNew(intCOl, strBTime, gintHourBegin)
        strDate = GetCenterTime(Split(strDate, ";")(0), Split(strDate, ";")(1))
    Else
         If Format(strValue, "YYYY-MM-DD") = Format(mstrETime, "YYYY-MM-DD") Then
            intCOl = GetCurveColumnNew(mstrETime, strBTime, gintHourBegin)
            strDate = GetCurveDateNew(intCOl, mstrBTime, gintHourBegin)
            strDate = GetCenterTime(Split(strDate, ";")(0), Split(strDate, ";")(1))
         ElseIf Format(strValue, "YYYY-MM-DD") > Format(strCurDate, "YYYY-MM-DD") And Format(strValue, "YYYY-MM-DD") < Format(mstrETime, "YYYY-MM-DD") Then
            strDate = GetCenterTime(Format(strValue, "YYYY-MM-DD 21:00:00"), Format(strValue, "YYYY-MM-DD 23:59:59"))
         End If
    End If

    For i = 0 To UBound(marrTime)
        If Format(strDate, "HH:mm:ss") >= Format(Split(marrTime(i), ",")(0), "HH:mm:ss") And Format(strDate, "HH:mm:ss") <= Format(Split(marrTime(i), ",")(1), "HH:mm:ss") Then
            Exit For
        End If
    Next i
    
    If i > UBound(marrTime) Then i = 0
    
    strBegin = Format(Format(strValue, "YYYY-MM-DD") & " " & Format(Split(marrTime(i), ",")(0), "HH:mm:ss"), "YYYY-MM-DD HH:mm:ss")
    strEnd = Format(Format(strValue, "YYYY-MM-DD") & " " & Format(Split(marrTime(i), ",")(1), "HH:mm:ss"), "YYYY-MM-DD HH:mm:ss")
    
    Call GetCenterTime(CDate(strBegin), CDate(strEnd), intBound)

    For i = 0 To OptTime.Count - 1
        OptTime(i).Caption = gintHourBegin + i * T_BodyStyle.lngʱ����
        OptTime(i).Tag = marrTime(i)
        
        If intBound > UBound(marrTime) Then intBound = 0
        If intBound = i Then
            OptTime(i).Value = 1
        End If
    Next i
    
    Call zlRefreshData(True, True)
    Call OptTime_Click(intBound)
    
    
    mblnAllRefresh = False
    dtpDateChageDate = True
    Exit Function
ErrInfo:
    If strErrMsg <> "" Then
        lblStb.Caption = strErrMsg
        lblStb.ForeColor = 255
    End If
    mblnAllRefresh = False
    
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function IsAllowInput(ByVal lng����ID As Long, ByVal lng��ҳID As Long, lngӤ�� As Long, ByVal strTime As String, ByVal strCurTime As String) As Boolean
    '----------------------------------------------
    '���ܣ�ȡ�����˷����䶯��¼��ʱ���
    '----------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strBabyOutTime As String
    
    On Error GoTo Errhand
    IsAllowInput = True
    If lngӤ�� <> 0 And mbln��Ժ = True Then
        strBabyOutTime = GetAdviceOutTime(lng����ID, lng��ҳID, lngӤ��)
        If strBabyOutTime <> "" Then
            strTime = Format(DateAdd("H", mlngHours, strBabyOutTime), "yyyy-MM-dd HH:mm")
            GoTo GONext
        End If
    End If
    gstrSQL = "" & _
              " SELECT DECODE(��ֹԭ��,1,'��Ժ',3,'ת��',10,'Ԥ��Ժ',15,'ת����',DECODE(��ʼԭ��,10,'��Ժ','δ����')) AS ����,��ֹʱ�� AS ʱ��" & _
              " From ���˱䶯��¼" & _
              " WHERE (��ֹԭ�� IN (1,3,10,15) OR ��ʼԭ��=10) And ����ID=[1] And ��ҳID=[2] And [3] <= ��ֹʱ��" & _
              " ORDER BY ��ֹʱ��"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��ָ��������ָ��ʱ��֮��ؼ����ʱ��", lng����ID, lng��ҳID, CDate(strTime))
    If rsTemp.RecordCount = 0 Then Exit Function
    'ֻȡ��һ�����ϵļ�¼
    strTime = Format(DateAdd("H", mlngHours, rsTemp!ʱ��), "yyyy-MM-dd HH:mm")
GONext:
    strCurTime = Format(strCurTime, "yyyy-MM-dd HH:mm")
    
    If strTime < strCurTime Then IsAllowInput = False
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub InitCommandBars()
'--------------------------------------------------------------------------------
'����:��ʼ��������
'--------------------------------------------------------------------------------
    Dim cbrControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    Dim cbrLable As CommandBarControl
    Dim cbrPop As CommandBarControl
    Dim cboChild As CommandBarPopup
    Dim CtlFont As stdFont
    
    On Error GoTo Errhand
    
     '��ʼ����
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    cbsMain.ActiveMenuBar.Title = "�˵���"
    cbsMain.ActiveMenuBar.Visible = False
    
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    
    With cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = False
        .SetIconSize False, 24, 24
        .SetIconSize True, 16, 16
        .UseSharedImageList = False 'ImageList��ʽʱ,��ͬһApp�й���,��AddImageList֮ǰ����ΪFalse
        Set CtlFont = .Font
        If CtlFont Is Nothing Then
            Set CtlFont = Me.Font
        End If
        CtlFont.Size = IIf(mintBigSize = 0, 9, 12)
        Set .Font = CtlFont
    End With

  '------------------------------------------------------------------------------------------------------------------
    '����������
    Set mcbrToolBar = cbsMain.Add("��׼", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    mcbrToolBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    
    With mcbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "����")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "ȡ��")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "�����Ŀ"): cbrControl.ToolTipText = "��ӻ��Ŀ": cbrControl.BeginGroup = True

        Set cbrPop = .Add(xtpControlButtonPopup, conMenu_Edit_Append, "���⴦��")
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 0, "����", -1, False): cbrControl.IconId = 1: cbrControl.Parameter = ""
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 1, "�೦[E]", -1, False):  cbrControl.IconId = 1: cbrControl.Parameter = "E"
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 2, "�೦����[/E]", -1, False):  cbrControl.IconId = 1: cbrControl.Parameter = "/E"
        Set cboChild = cbrPop.CommandBar.Controls.Add(xtpControlButtonPopup, conMenu_Edit_Append * 10 + 3, "���ʧ��", -1, False): cbrControl.IconId = 1
        Set cbrControl = cboChild.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 30, "��", -1, False):  cbrControl.Parameter = "��"
        Set cbrControl = cboChild.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 31, "*", -1, False): cbrControl.Parameter = "*"
        
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 4, "�˹�����[��]", -1, False): cbrControl.IconId = 1: cbrControl.Parameter = "��"
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 5, "����[C]", -1, False):   cbrControl.IconId = 1: cbrControl.Parameter = "C"
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 6, "��������[/C]", -1, False):   cbrControl.IconId = 1: cbrControl.Parameter = "/C"
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����")
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    '��λ������
    '------------------------------------------------------------------------------------------------------------------
    For Each cbrControl In mcbrToolBar.Controls
        If cbrControl.Type <> xtpControlCustom And cbrControl.Type <> xtpControlLabel Then
            cbrControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    With picDate
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
        .Width = (dtpDate.Width + 520) + (dtpDate.Width + 520) * mintBigSize / 3
        .Height = 300 + 300 * mintBigSize / 3
    End With
    
    With dtpDate
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
        .Width = .Width + .Width * mintBigSize / 3
        .Height = 300 + 300 * mintBigSize / 3
        .Top = 0
        .Left = 0
    End With

    With imgbtn(1)
        .Width = 240 + 240 * mintBigSize / 3
        .Height = 240 + 240 * mintBigSize / 3
        .Top = 30
        .Left = dtpDate.Width + 20
    End With
    With imgbtn(0)
        .Width = 240 + 240 * mintBigSize / 3
        .Height = 240 + 240 * mintBigSize / 3
        .Top = 30
        .Left = dtpDate.Width + imgbtn(1).Width + 30
    End With
    
    '���ڲ�¼
    '------------------------------------------------------------------------------------------------------------------
    Set cbrLable = mcbrToolBar.Controls.Add(xtpControlLabel, conMenu_View_Option, "")
    cbrLable.flags = xtpFlagRightAlign
    Set cbrCustom = mcbrToolBar.Controls.Add(xtpControlCustom, conMenu_View_Location, "")
    dtpDate.Visible = True
    cbrCustom.Handle = picDate.hWnd
    imgbtn(1).Picture = ilsDate.ListImages("preGreen").Picture
    imgbtn(0).Picture = ilsDate.ListImages("nextGreen").Picture
    cbrCustom.flags = xtpFlagRightAlign
    
'    Set cbrControl = mcbrToolBar.Controls.Add(xtpControlButton, conMenu_View_Forward, "��һ��")
'    cbrControl.flags = xtpFlagRightAlign
'    cbrControl.IconId = conMenu_View_Forward
'    If dtpDate.Value = dtpDate.MinDate Then cbrControl.Enabled = False
'    Set cbrControl = mcbrToolBar.Controls.Add(xtpControlButton, conMenu_View_Backward, "��һ��")
'    cbrControl.flags = xtpFlagRightAlign
'    cbrControl.IconId = conMenu_View_Backward
'    If dtpDate.Value = dtpDate.MaxDate Then cbrControl.Enabled = False

    '�����
    With cbsMain.KeyBindings
        .Add FALT, Asc("0"), conMenu_Edit_Append * 10
        .Add FALT, Asc("1"), (conMenu_Edit_Append * 10 + 1)
        .Add FALT, Asc("2"), (conMenu_Edit_Append * 10 + 2)
        .Add FALT, Asc("3"), (conMenu_Edit_Append * 10 + 30)
        .Add FALT, Asc("4"), (conMenu_Edit_Append * 10 + 31)
        .Add FALT, Asc("5"), (conMenu_Edit_Append * 10 + 4)
        .Add FALT, Asc("6"), (conMenu_Edit_Append * 10 + 5)
        .Add FALT, Asc("7"), (conMenu_Edit_Append * 10 + 6)
        
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem '��ӻ��Ŀ
        .Add FCONTROL, Asc("S"), conMenu_Edit_Save '����
        .Add FCONTROL, Asc("R"), conMenu_Edit_Reuse 'ȡ��
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    Call InitDateTimeRange(marrTime, gintHourBegin, T_BodyStyle.lng������, T_BodyStyle.lngʱ����)
     
    '���ر��ؼ�
    Call InitTabControl
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Function GetTableRowName() As String
'----------------------------------------------------
'��ʼ�����
'----------------------------------------------------
    Dim arrItem() As Variant
    Dim strSQL As String
    Dim strTmp As String
    Dim strֵ�� As String
    Dim strCurDate As String
    Dim strEndTime As String
    Dim strDate As String
    Dim strTmpCurve As String '������Ŀ����
    Dim strTmpTable As String '�����Ŀ����
    Dim i As Integer, intBound As Integer
    Dim intCOl As Integer
    Dim Titem As Type_Item
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo Errhand
    
    arrItem = Array()
    Call InitRecordSet
    Call GetPainDegreeNO
     '����������ʹ���ʱ�����Ƿ�ʹ����˲���
    strSQL = "select C.Ӧ�÷�ʽ From �����¼��Ŀ C where C.��Ŀ���=[1] And C.����ȼ�>=[2] And Nvl(C.���ò���,0) In (0,[3]) " & _
            " And (c.���ÿ���=1 Or (c.���ÿ���=2 And Exists (Select 1 From �������ÿ��� D Where D.��Ŀ���=c.��Ŀ��� And D.����id=[4])))"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����", -1, mT_Patient.lng����ȼ�, IIf(mT_Patient.lngӤ�� = 0, 1, 2), mT_Patient.lng����ID)
    mblnEdit���� = IIf(rsTemp.RecordCount = 0, False, True)
    If rsTemp.RecordCount > 0 Then mint����Ӧ�� = Val(zlStr.Nvl(rsTemp!Ӧ�÷�ʽ, 0))
    
    '��ʽ���Ϊ ����'ֵ��,��Ŀ����,��ĿС��,��¼Ƶ��,��Ŀ��ʾ,��Ŀ����,��Ŀ����,��λ,��Ժ�ײ�'��Ŀ��'��Ŀ��
    strTmp = "2)���±�˵��',,,,,,,,'2'�ϱ�'�ϱ�;2)���±�˵��',,,,,,,,'6'�±�'�±�"
    
    '��ȡȫ������������Ŀ
    mstrCurveItem = ""
    mstrCurveItem = T_BodyItem.str������Ŀ
    If InStr(1, "," & mstrCurveItem & ",", "," & gint���� & ",") = 0 Then
        If InStr(1, Val(T_BodyItem.str�������), gint����) > 0 Then
            mstrCurveItem = mstrCurveItem & "," & gint����
        End If
    End If
    strSQL = " Select /*+ RULE */" & _
             " a.�������, a.��¼�� ��Ŀ��, a.��Ŀ��� As ��Ŀ��, a.��¼��, a.��Ժ�ײ�, c.��Ŀֵ��, c.��Ŀ����, c.��Ŀ����, c.��ĿС��, Nvl(a.��¼Ƶ��, 2) ��¼Ƶ��, c.������, c.��Ŀ��ʾ," & _
             "  c.��Ŀ��λ " & vbNewLine & _
             "  From ���¼�¼��Ŀ A, ����������Ŀ B, �����¼��Ŀ C, Table(Cast(f_Num2list([1]) As Zltools.t_Numlist)) D " & vbNewLine & _
             " Where c.��Ŀid = b.Id(+) And a.��Ŀ��� = c.��Ŀ��� And (a.��¼�� <> 2 Or (a.��¼�� = 2 And a.��Ŀ��� = 3)) And " & vbNewLine & _
             "      Not (c.Ӧ�÷�ʽ = 2 And c.��Ŀ��� = -1) And c.��Ŀ��� = d.Column_Value " & vbNewLine & _
             " Order By Decode(a.��Ŀ���, 1, 0, 1), a.������� "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrCurveItem)
    
    strTmpCurve = ""
    With rsTemp
        Do While Not .EOF
            strֵ�� = Replace(zlStr.Nvl(!��Ŀֵ��), ":", "")
            If zlStr.Nvl(!��Ŀ����) = 0 Then
                If InStr(1, strֵ��, ";") Then strֵ�� = Split(strֵ��, ";")(0) & "��" & Split(strֵ��, ";")(1)
            End If
            strֵ�� = Replace(Replace(Replace(strֵ��, ";", ":"), "'", ""), ",", "")
            Titem.ֵ�� = strֵ��
            Titem.��Ŀ�� = Replace(Replace(zlStr.Nvl(!��Ŀ��) & IIf(zlStr.Nvl(!��Ŀ��λ, "") = "", "", "(" & !��Ŀ��λ & ")"), ";", ":"), "'", "")
            Titem.��¼�� = zlStr.Nvl(!��Ŀ��)
            Titem.��Ŀ�� = Val(zlStr.Nvl(!��Ŀ��))
            Titem.��Ժ�ײ� = Val(zlStr.Nvl(!��Ժ�ײ�, 0))
            Titem.��Ŀ���� = Val(zlStr.Nvl(!��Ŀ����, 0))
            Titem.��Ŀ���� = Val(zlStr.Nvl(!��Ŀ����, 3))
            Titem.��ĿС�� = Val(zlStr.Nvl(!��ĿС��, 0))
            Titem.��¼Ƶ�� = Val(zlStr.Nvl(!��¼Ƶ��))
            Titem.��Ŀ��ʾ = Val(zlStr.Nvl(!��Ŀ��ʾ, 0))
            If Titem.��Ŀ��ʾ = 4 Or IsWaveItem(Titem.��Ŀ��) Then
                If Titem.��¼Ƶ�� > 2 Then Titem.��¼Ƶ�� = 2
            End If
            Titem.��λ = ""
            Titem.��Ŀ���� = 1
            '��¼��Ϊ1�ͼ�¼��Ϊ2�ĺ�����ĿΪ������Ŀ
            Titem.���� = "1)����������Ŀ"
            strTmpCurve = strTmpCurve & ";" & Titem.���� & "'" & Titem.ֵ�� & "," & Titem.��Ŀ���� & "," & _
                Titem.��ĿС�� & "," & Titem.��¼Ƶ�� & "," & Titem.��Ŀ��ʾ & Titem.��Ŀ���� & Titem.��Ŀ���� & Titem.��λ & Titem.��Ժ�ײ� & "'" & _
                Titem.��Ŀ�� & "'" & Titem.��Ŀ�� & "'" & Titem.��¼��
        .MoveNext
        Loop
    End With
    
    strEndTime = DateAdd("d", T_BodyStyle.lng����, CDate(Format(Format(mstrDate, "YYYY-MM-DD") & " 23:59:59", "YYYY-MM-DD HH:mm:ss")))
    If strEndTime > mstrETime Then strEndTime = mstrETime
    mstrActiveItem = ""
    Set rsTemp = GetAppendGridItemNew(mT_Patient.lng�ļ�ID, mT_Patient.lng����ID, mT_Patient.lng��ҳID, mT_Patient.lng����ȼ�, mT_Patient.lngӤ��, _
            CDate(mstrDate), CDate(strEndTime), IIf(mT_Patient.lngӤ�� = 0, 1, 2), mT_Patient.lng����ID, T_BodyItem.str�����Ŀ, mblnMove)
    strTmpTable = ""
    With rsTemp
        Do While Not .EOF
            strֵ�� = Replace(zlStr.Nvl(!��Ŀֵ��), ":", "")
            If zlStr.Nvl(!��Ŀ����) = 0 Then
                If InStr(1, strֵ��, ";") <> 0 Then strֵ�� = Split(strֵ��, ";")(0) & "��" & Split(strֵ��, ";")(1)
            End If
            strֵ�� = Replace(Replace(Replace(strֵ��, ";", ":"), "'", ""), ",", "")
            Titem.ֵ�� = strֵ��
            Titem.���� = "2)���±����Ŀ"
            Titem.��Ŀ���� = Val(zlStr.Nvl(!��Ŀ����))
            Titem.��ĿС�� = Val(zlStr.Nvl(!��ĿС��, 0))
            Titem.��¼Ƶ�� = Val(zlStr.Nvl(!��¼Ƶ��, 2))
            Titem.��Ŀ��ʾ = Val(zlStr.Nvl(!��Ŀ��ʾ, 0))
            Titem.��Ŀ���� = Val(zlStr.Nvl(!��Ŀ����, 1))
            Titem.��Ŀ���� = zlStr.Nvl(!��Ŀ����, 3)
            Titem.��λ = Replace(Replace(Replace(zlStr.Nvl(!���²�λ), ";", ""), "'", ""), ",", "")
            Titem.��Ŀ�� = Val(zlStr.Nvl(!��Ŀ���))
            Titem.��Ŀ�� = Replace(Replace(IIf(Titem.��Ŀ�� = 4, "Ѫѹ", zlStr.Nvl(!��¼��)) & IIf(zlStr.Nvl(!��λ, "") = "", "", "(" & !��λ & ")"), ";", ":"), "'", "")
            Titem.��Ժ�ײ� = Val(zlStr.Nvl(!��Ժ�ײ�, 0))
            Titem.��¼�� = IIf(Titem.��Ŀ�� = 4, "Ѫѹ", zlStr.Nvl(!��¼��))
            
            If Titem.��Ŀ��ʾ = 4 Or IsWaveItem(Titem.��Ŀ��) Then
                If Titem.��¼Ƶ�� > 2 Then Titem.��¼Ƶ�� = 2
            End If
            
            If Titem.��Ŀ�� <> gint���� And Titem.��Ŀ�� <> 5 Then
                strTmpTable = strTmpTable & ";" & Titem.���� & "'" & Titem.ֵ�� & "," & Titem.��Ŀ���� & "," & _
                    Titem.��ĿС�� & "," & Titem.��¼Ƶ�� & "," & Titem.��Ŀ��ʾ & "," & Titem.��Ŀ���� & "," & Titem.��Ŀ���� & "," & _
                    Titem.��λ & "," & Titem.��Ժ�ײ� & "'" & Titem.��Ŀ�� & "'" & Titem.��Ŀ�� & "'" & Titem.��¼��
                '���Ŀ
                If Titem.��Ŀ���� = 2 Then
                    mstrActiveItem = mstrActiveItem & ";" & Titem.���� & "'" & Titem.ֵ�� & "," & Titem.��Ŀ���� & "," & _
                        Titem.��ĿС�� & "," & Titem.��¼Ƶ�� & "," & Titem.��Ŀ��ʾ & "," & Titem.��Ŀ���� & "," & Titem.��Ŀ���� & "," & _
                        Titem.��λ & "," & Titem.��Ժ�ײ� & "'" & Titem.��Ŀ�� & "'" & Titem.��Ŀ�� & "'" & Titem.��¼��
                End If
            End If
        .MoveNext
        Loop
    End With
    
    If Left(mstrActiveItem, 1) = ";" Then mstrActiveItem = Mid(mstrActiveItem, 2)
    If strTmp <> "" Then strTmpCurve = strTmpCurve & ";" & strTmp
    If Left(strTmpCurve, 1) = ";" Then strTmpCurve = Mid(strTmpCurve, 2)
    If Left(strTmpTable, 1) = ";" Then strTmpTable = Mid(strTmpTable, 2)
    
    '��ʼ�������������ݰ������±�
    Call InitTabCurve(strTmpCurve)
    '��ʼ�����±��
    Call InitTabTable(strTmpTable)
    '��ȡδ��˵��
    mstrδ��˵�� = ""
    mrsCurInfo.Filter = ""
    mrsCurInfo.Sort = "����"
    With mrsCurInfo
        Do While Not .EOF
            mstrδ��˵�� = IIf(mstrδ��˵�� = "", "", mstrδ��˵�� & "'") & zlStr.Nvl(!����)
            .MoveNext
        Loop
    End With
    If Left(mstrδ��˵��, 1) = "'" Then mstrδ��˵�� = Mid(mstrδ��˵��, 2)
    
    '����ѡ��ʱ�䶨λ��ǰʱ��༭״̬
    strCurDate = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:mm:ss")
    If Format(mstrBegin, "YYYY-MM-DD") = Format(strCurDate, "YYYY-MM-DD") Then
        If Format(strCurDate, "YYYY-MM-DD HH:mm:ss") < Format(mstrBegin, "YYYY-MM-DD HH:mm:ss") Then
             strCurDate = Format(mstrBegin, "YYYY-MM-DD HH:mm:ss")
        ElseIf Format(strCurDate, "YYYY-MM-DD HH:mm:ss") > Format(strEndTime, "YYYY-MM-DD HH:mm:ss") Then
             strCurDate = Format(mstrEnd, "YYYY-MM-DD HH:mm:ss")
        End If
        intCOl = GetCurveColumnNew(strCurDate, mstrBegin, gintHourBegin)
        strDate = GetCurveDateNew(intCOl, mstrBegin, gintHourBegin)
        mstrBegin = Split(strDate, ";")(0)
        mstrEnd = Split(strDate, ";")(1)
    Else
         If Format(mstrBegin, "YYYY-MM-DD") = Format(strEndTime, "YYYY-MM-DD") Then
            intCOl = GetCurveColumnNew(mstrEnd, mstrBegin, gintHourBegin)
            strDate = GetCurveDateNew(intCOl, mstrBegin, gintHourBegin)
            mstrBegin = Split(strDate, ";")(0)
            mstrEnd = Split(strDate, ";")(1)
         ElseIf Format(mstrBegin, "YYYY-MM-DD") > Format(strCurDate, "YYYY-MM-DD") And Format(mstrBegin, "YYYY-MM-DD") < Format(strEndTime, "YYYY-MM-DD") Then
            mstrBegin = Format(mstrBegin, "YYYY-MM-DD 21:00:00")
            mstrEnd = Format(mstrBegin, "YYYY-MM-DD 23:59:59")
         End If
    End If
    
    '��ȡ��ǰʱ����ڵ���ĵڼ���λ����
    Call GetCenterTime(CDate(mstrBegin), CDate(mstrEnd), intBound)
    For i = 0 To OptTime.Count - 1
        OptTime(i).Caption = gintHourBegin + i * T_BodyStyle.lngʱ����
        OptTime(i).Tag = marrTime(i)
        
        If intBound > UBound(marrTime) Then intBound = 0
        If intBound = i Then
            OptTime(i).Value = 1
        End If
    Next i
    lblTime.Caption = Format(mstrBegin, "HH:mm") & "��" & Format(mstrEnd, "HH:mm")
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function IsTatle(ByVal lngItemNO As Long) As Boolean
'---------------------------------------------------------
'����Ƿ��ǻ�����Ŀ
'---------------------------------------------------------
    If mrsCollect Is Nothing Then Exit Function
    If mrsCollect.State = adStateOpen Then
        mrsCollect.Filter = "���=" & lngItemNO
        IsTatle = mrsCollect.RecordCount > 0
    End If
End Function


Private Sub InitTabCurve(ByVal strTabName As String)
'-------------------------------------------------------
'����:��ʼ������������Ŀ
'����:���б�ͷ����Ϣ
'-------------------------------------------------------
    Dim varTabName() As String, varCode() As String
    Dim intRow As Integer, intCOl As Integer
    
    On Err GoTo Errhand
    If strTabName = "" Then Exit Sub
    varTabName = Split(strTabName, ";")
    With vsfCurve
    
        .Rows = UBound(varTabName) + 2
        .Cols = 0
        .NewColumn "", 255, 4
        .NewColumn "�༭", 900 + 900 * mintBigSize / 3, 1
        .NewColumn "������", 1500 + 1500 * mintBigSize / 3, 1
        .NewColumn "�ַ���", 900 + 900 * mintBigSize / 3, 1
        .NewColumn "��Ŀ���", 900 + 900 * mintBigSize / 3, 1
        .NewColumn "��Ŀ��", 900 + 900 * mintBigSize / 3, 1
        .NewColumn "ԭʼʱ��", 900 + 900 * mintBigSize / 3, 1
        .NewColumn "�޸�״̬", 900 + 900 * mintBigSize / 3, 1
        .NewColumn "��Ŀ����", 1200 + 1200 * mintBigSize / 3, 1
        .NewColumn "��ʾ", 900 + 900 * mintBigSize / 3, 1
        .NewColumn "ʱ��", 900 + 900 * mintBigSize / 3, 1, , 4
        .NewColumn "ԭֵ", 300 + 300 * mintBigSize / 3, 1
        .NewColumn "����", 2300 + 2300 * mintBigSize / 3, 1, , 4
        .NewColumn "����", 300 + 300 * mintBigSize / 3, 0
        .NewColumn "���Ժϸ�", 900 + 900 * mintBigSize / 3, 4
        .NewColumn "��λ", 1000 + 1000 * mintBigSize / 3, 4
        .NewColumn "δ��˵��", 1080 + 1080 * mintBigSize / 3, 4, "...", 1
        .NewColumn "��Դ", 900 + 900 * mintBigSize / 3, 4
        .NewColumn "������Դ", 300 + 300 * mintBigSize / 3, 1
        .NewColumn "ɾ��", 900 + 900 * mintBigSize / 3, 4
        .Body.RowHeight(0) = 300 + 300 * mintBigSize / 3
        .FixedCols = COL_��Ŀ���� + 1
        .FixedRows = 1
        
        .Body.Font.Size = mFontSize + mFontSize * mintBigSize / 3
        .Body.ColHidden(COL_�༭) = True
        .Body.ColHidden(COL_�ַ���) = True
        .Body.ColHidden(COL_��Ŀ���) = True
        .Body.ColHidden(COL_��Ŀ��) = True
        .Body.ColHidden(col_ԭʼʱ��) = True
        .Body.ColHidden(COL_�޸�״̬) = True
        .Body.ColHidden(COL_��ʾ) = True
        .Body.ColHidden(COL_ԭֵ) = True
        .Body.ColHidden(COL_��Դ) = True
        .Body.ColHidden(COL_������Դ) = True
        .Body.ColHidden(COL_ɾ��) = True
        .Body.WordWrap = True
        .Body.MergeCells = flexMergeRestrictColumns
        .Body.MergeCol(COL_������) = True
        .Body.MergeRow(0) = True
        
        For intRow = .FixedRows To .Rows - 1
            varCode = Split(varTabName(intRow - 1), "'")
            If UBound(varCode) > 2 Then
                .TextMatrix(intRow, COL_������) = varCode(0)
                .TextMatrix(intRow, COL_�ַ���) = varCode(1)
                .TextMatrix(intRow, COL_��Ŀ���) = varCode(2)
                .TextMatrix(intRow, COL_��Ŀ����) = varCode(3)
                .TextMatrix(intRow, COL_��Ŀ��) = varCode(4)
                .TextMatrix(intRow, COL_����) = Space(2)
                .TextMatrix(intRow, COL_��ɫ) = Space(2)
                If varCode(0) = "2)���±�˵��" Then
                    Select Case Val(varCode(2))
                        Case 2
                            mOptRow.�ϱ� = intRow
                        Case 6
                            mOptRow.�±� = intRow
                    End Select
                End If
            End If
            .Body.RowHeight(intRow) = 300 + 300 * mintBigSize / 3
            .RowData(intRow) = 0
        Next intRow
        .Body.MergeRow(intRow - 2) = True
        .Body.Cell(flexcpAlignment, 0, .FixedCols, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        .Body.Cell(flexcpText, .FixedRows, .FixedCols, .Rows - 1, .Cols - 1) = ""
    End With

    With vsfDetail
        .FixedRows = 1
        .Rows = .FixedRows + 1
        .Cols = 0
        .NewColumn "", 255, 4
        .NewColumn "�༭", 900 + 900 * mintBigSize / 3, 1
        .NewColumn "", 255, 4
        .NewColumn "�ַ���", 900 + 900 * mintBigSize / 3, 1
        .NewColumn "��Ŀ���", 900 + 900 * mintBigSize / 3, 1
        .NewColumn "��Ŀ��", 900 + 900 * mintBigSize / 3, 1
        .NewColumn "ԭʼʱ��", 900 + 900 * mintBigSize / 3, 1
        .NewColumn "�޸�״̬", 900 + 900 * mintBigSize / 3, 1
        .NewColumn "��Ŀ����", 900 + 900 * mintBigSize / 3, 1, , 4
        .NewColumn "��ʾ", 700 + 700 * mintBigSize / 3, 1, , 4
        .NewColumn "ʱ��", 900 + 900 * mintBigSize / 3, 1, , 4
        .NewColumn "ԭֵ", 900 + 900 * mintBigSize / 3, 4
        .NewColumn "����", 1200 + 1200 * mintBigSize / 3, 1, , 4
        .NewColumn "��ɫ", 900 + 900 * mintBigSize / 3, 1
        .NewColumn "���Ժϸ�", 1200 + 1200 * mintBigSize / 3, 1, , 4
        .NewColumn "��λ", 1000 + 1000 * mintBigSize / 3, 4
        .NewColumn "δ��˵��", 1080 + 1080 * mintBigSize / 3, 4, "...", 1
        .NewColumn "��Դ", 900 + 900 * mintBigSize / 3, 4
        .NewColumn "������Դ", 1400 + 1400 * mintBigSize / 3, 1, , 4
        .NewColumn "", 350 + 350 * mintBigSize / 3, 1, , 4
        .Body.RowHeightMin = 300 + 300 * mintBigSize / 3
        .Body.ColComboList(COL_��λ) = " "
        .Body.ColComboList(Col_δ��˵��) = "..."
        .Body.Font.Size = mFontSize + mFontSize * mintBigSize / 3
        .Body.ColHidden(COL_�༭) = True
        .Body.ColHidden(COL_Null) = True
        .Body.ColHidden(COL_��Ŀ���) = True
        .Body.ColHidden(COL_��Ŀ��) = True
        .Body.ColHidden(col_ԭʼʱ��) = True
        .Body.ColHidden(COL_��Ŀ����) = True
        .Body.ColHidden(COL_�ַ���) = True
        .Body.ColHidden(COL_�޸�״̬) = True
        .Body.ColHidden(COL_ԭֵ) = True
        .Body.ColHidden(COL_��ɫ) = True
        .Body.ColHidden(COL_��Դ) = True
        .Body.WordWrap = False
        .FixedCols = COL_������ + 1
        .Body.AllowUserResizing = flexResizeNone
        .Body.Cell(flexcpAlignment, 0, .FixedCols, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        .Body.Cell(flexcpText, .FixedRows, .FixedCols, .Rows - 1, .Cols - 1) = ""
    End With
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitTabTable(ByVal strTabName As String)
    '--------------------------------------
    '��ʼ�����߱�����ϸ�����ͷ
    '--------------------------------------
    Dim varTabName() As String, varCode() As String
    Dim intRow As Integer, intCOl As Integer
    Dim i As Integer, lngPreNum As Long
    
    On Error GoTo Errhand
    If strTabName = "" Then
        strTabName = "',0,0,0,0,1,0,,0'-999''"
        vsfTab.Tag = "NO"
    Else
        vsfTab.Tag = ""
    End If
    varTabName = Split(strTabName, ";")
    With vsfTab
        .Cols = 13
        .Rows = UBound(varTabName) * 2 + 3
        .FixedRows = 1
        .FixedCols = 7
        .ColHidden(COL_tab�ַ���) = True
        .ColHidden(COL_tab��Ŀ���) = True
        .ColHidden(COL_tab��Ŀ��) = True
        .ColHidden(col_tabԭʼʱ��) = True
        .WordWrap = True
        .MergeCells = flexMergeRestrictColumns
        .MergeCol(COL_tab��Ŀ����) = True
        .MergeRow(0) = True
        .ScrollBars = flexScrollBarBoth
        .SelectionMode = flexSelectionByRow
        .ColWidth(COL_TabNull) = 255
        .ColWidth(COL_tab��Ŀ����) = 1200
        .ColWidth(COL_tabDirect) = 600
        .RowHeightMin = 300 + 300 * mintBigSize / 3
        For intCOl = .FixedCols - 2 To .Cols - 1
            If intCOl < .FixedCols Then
                .TextMatrix(0, intCOl) = "����/Ƶ��"
            Else
                .TextMatrix(0, intCOl) = intCOl - .FixedCols + 1
                .ColWidth(intCOl) = 1200 + 1200 * mintBigSize / 3
            End If
        Next intCOl
        
        i = 1
        For intRow = 1 To .Rows - 1
            varCode = Split(varTabName(i - 1), "'")
            .RowData(intRow) = Split(varCode(1), ",")(3) & ";" & IIf(IsWaveItem(varCode(2)), 2, 0)
            If IsTatle(varCode(2)) Then .RowData(intRow) = Split(varCode(1), ",")(3) & ";" & "3"
            .RowData(intRow + 1) = .RowData(intRow)
            .TextMatrix(intRow, COL_tab�ַ���) = varCode(1)
            .TextMatrix(intRow, COL_tab��Ŀ���) = varCode(2)
            .TextMatrix(intRow, COL_tab��Ŀ��) = varCode(4)
            .TextMatrix(intRow, COL_TabNull) = ""
            .TextMatrix(intRow, COL_tab��Ŀ����) = varCode(3)
            If Split(varCode(1), ",")(3) > 0 Then .TextMatrix(intRow, col_tabԭʼʱ��) = Replace(Space(Split(varCode(1), ",")(3) - 1), " ", "'")
            .TextMatrix(intRow, COL_tabDirect) = "ʱ��"
            .TextMatrix(intRow + 1, COL_tab�ַ���) = varCode(1)
            .TextMatrix(intRow + 1, COL_tab��Ŀ���) = varCode(2)
            .TextMatrix(intRow + 1, COL_tab��Ŀ��) = varCode(4)
            .TextMatrix(intRow + 1, COL_TabNull) = ""
            .TextMatrix(intRow + 1, COL_tab��Ŀ����) = varCode(3)
            If Split(varCode(1), ",")(3) > 0 Then .TextMatrix(intRow + 1, col_tabԭʼʱ��) = Replace(Space(Split(varCode(1), ",")(3) - 1), " ", "'")
            .TextMatrix(intRow + 1, COL_tabDirect) = "����"
            intRow = intRow + 1
            i = i + 1
        Next intRow
        .Cell(flexcpBackColor, .FixedRows, .FixedCols, .Rows - 1, .Cols - 1) = &H80000005
        '���ݱ���Ƶ�ξ�������ɫ
        For intRow = .FixedRows To .Rows - 1
            If .FixedCols + (Val(Split(.RowData(intRow), ";")(0))) < .Cols Then
                .Cell(flexcpBackColor, intRow, .FixedCols + (Val(Split(.RowData(intRow), ";")(0))), intRow, .Cols - 1) = &H8000000F
            End If
        Next intRow
        .CellBorderRange .FixedRows, .FixedCols, .Rows - 1, .Cols - 1, .GridColor, 0, 0, 1, 0, 0, 0
        For intRow = .FixedRows To .Rows - 1
            '�������߿�
            For intCOl = 0 To (Val(Split(.RowData(intRow), ";")(0))) - 1
                .CellBorderRange intRow, .FixedCols + intCOl, intRow, .FixedCols + intCOl, .GridColor, 0, IIf(intCOl + 1 > lngPreNum And intRow > .FixedRows, 1, 0), 1, IIf(.FixedCols + intCOl < .Cols, 1, 0), 0, 0
            Next intCOl
            lngPreNum = (Val(Split(.RowData(intRow), ";")(0)))
        Next intRow
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
         
    End With

    With vsfTabDetail
        .Cols = 9
        .Rows = 2
        .FixedCols = 6
        .WordWrap = True
        .RowHeightMin = 300 + 300 * mintBigSize / 3
        .ColWidth(COL_TabNull) = 255
        .ColHidden(COL_tab�ַ���) = True
        .ColHidden(COL_tab��Ŀ���) = True
        .ColHidden(COL_tab��Ŀ��) = True
        .ColHidden(col_tabԭʼʱ��) = True
        .TextMatrix(0, .FixedCols - 1) = "����"
        .ColWidth(.FixedCols - 1) = 1500
        .TextMatrix(0, .FixedCols) = "ʱ��"
        .ColWidth(.FixedCols) = 1600
        .TextMatrix(0, .FixedCols + 1) = "����"
        .ColWidth(.FixedCols + 1) = 1200
        .TextMatrix(0, .FixedCols + 2) = "��Դ"
        .ColWidth(.FixedCols + 2) = 2500
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    End With
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub




Private Function SaveData() As Boolean
    '--------------------
    '���ݱ���
    '--------------------
    Dim lngItemCode As Long
    Dim lng��¼ID As Long, lngOld��¼ID As Long
    Dim lngRow As Long
    Dim intModify As Integer, int������ As Integer
    Dim i As Integer, int��Ŀ�״� As Integer
    Dim intԭʼ��ʾ As Integer
    Dim strValue As String, strδ�� As String
    Dim strSQL As String, strTime As String
    Dim strEnd As String, strBegin As String
    Dim strOldTime As String, strSQLShow As String
    Dim str��λ As String, strName As String, strTmp As String
    Dim strInfo As String
    Dim arrSQL() As String, arrSQLTime() As String
    Dim arrSQLShow() As String, arrTmp() As String
    Dim blnEdit As Boolean, blnSave As Boolean
    Dim blnTran As Boolean
    On Error GoTo Errhand
    
    mrsCurve.Filter = 0
    mrsCurve.Sort = "ʱ��,��Ŀ���"
    mrsTableDetail.Filter = 0
    mrsCurve.Sort = "ʱ��,��Ŀ���"
    Screen.MousePointer = 11
    ReDim Preserve arrSQL(1 To 1)
    ReDim Preserve arrSQLTime(1 To 1)
    ReDim Preserve arrSQLShow(1 To 1)
    mrsRecodeID.Filter = 0
    
    '�������߱���
    With mrsCurve
        Do While Not .EOF
            lngItemCode = Val(!��Ŀ���)
            strValue = Nvl(!��ֵ)
            mrsCurInfo.Filter = "����='" & strValue & "'"
            intModify = Val(zlStr.Nvl(!�޸�))
            blnEdit = False
            If intModify = 1 And InStr(1, ",0,3,9,", Val(zlStr.Nvl(!������Դ))) = 0 Then
                blnEdit = False
            Else
                blnEdit = True
            End If
            str��λ = Nvl(!��λ)
            blnSave = False
            If !״̬ <> 0 Then
               
                '����������Ŀ
                strTime = Format(!ʱ��, "YYYY-MM-DD HH:mm:ss")
                strOldTime = Format(!ԭʼʱ��, "YYYY-MM-DD hh:mm:ss")
                int������ = IIf(ISCheckDept(strTime) = True, 1, 0)
                strBegin = Format(mstrBegin, "YYYY-MM-DD HH:mm:ss")
                strEnd = Format(mstrEnd, "YYYY-MM-DD HH:mm:ss")
                
                
                '������ʾ״̬
                intԭʼ��ʾ = Val(Nvl(!ԭʼ��ʾ״̬))
                If intԭʼ��ʾ <> !��ʾ Then
                    strSQLShow = "Zl_���µ�����_������ʾ("
                    '����ʱ��_In In ���˻�������.����ʱ��%Type,
                     strSQLShow = strSQLShow & "To_Date('" & strTime & "','yyyy-mm-dd hh24:mi:ss')" & ","
                    '�ļ�id_In   In ���˻�������.�ļ�id%Type,
                    strSQLShow = strSQLShow & mT_Patient.lng�ļ�ID & ","
                    '��Ŀ���_In In ���˻�����ϸ.��Ŀ���%Type,
                    strSQLShow = strSQLShow & lngItemCode & ","
                    '��λ_In     In ���˻�����ϸ.���²�λ%Type,
                    strSQLShow = strSQLShow & "'" & str��λ & "',"
                    '��ʾ_In     In ���˻�����ϸ.��ʾ%Type
                    strSQLShow = strSQLShow & Val(!��ʾ) & ")"
                    
                    arrSQLShow(ReDimArray(arrSQLShow)) = strSQLShow
                End If
                
                '���޸����ݷ���ʱ��
                If strOldTime <> strTime And strOldTime <> "" Then
                    mrsRecodeID.Filter = "ʱ��='" & strOldTime & "'"
                    If mrsRecodeID.RecordCount > 0 Then
                        lng��¼ID = Val(mrsRecodeID!��¼ID)
                        '��ͬ��¼�޸ĺ����޸�
                        If lng��¼ID <> lngOld��¼ID Then
                            strSQL = "ZL_���µ�����_����ʱ��("
                            'ID_IN       IN ���˻�������.ID%TYPE,
                            strSQL = strSQL & lng��¼ID & ","
                            '����ʱ��_IN IN ���˻�������.����ʱ��%TYPE
                            strSQL = strSQL & "To_Date('" & strTime & "','yyyy-mm-dd hh24:mi:ss')" & ")"
                            
                            arrSQLTime(ReDimArray(arrSQLTime)) = strSQL
                        End If
                    End If
                End If
                lngOld��¼ID = lng��¼ID
                If strValue = "����" And lngItemCode = gint���� Then
                    strδ�� = ""
                Else
                    strδ�� = !δ��˵��
                End If
                If Val(!״̬) <> 5 And Val(!״̬) <> 6 Then '״̬Ϊ5ֻ�޸���ʱ��  ״̬Ϊ6ֻ�޸�����ʾ״̬
                    '������������������Ϣ
                    strSQL = "Zl_���µ�����_Update("
                    '�ļ�id_In   In ���˻����ļ�.Id%Type,  --���˻����ļ�ID
                    strSQL = strSQL & Val(mT_Patient.lng�ļ�ID) & ","
                    '����ʱ��_In In ���˻�������.����ʱ��%Type, --�������ݵķ���ʱ��
                    strSQL = strSQL & "To_Date('" & strTime & "','yyyy-mm-dd hh24:mi:ss')" & ","
                    '��¼����_In In ���˻�����ϸ.��¼����%Type, --������Ŀ=1���ϱ�˵��=2�����ת���=3�������ձ��=4,�±�˵��=6
                    strSQL = strSQL & "1,"
                    '��Ŀ���_In In ���˻�����ϸ.��Ŀ���%Type, --������Ŀ����ţ��ǻ�����Ŀ�̶�Ϊ0
                    strSQL = strSQL & lngItemCode & ","
                    '��¼����_In In ���˻�����ϸ.��¼����%Type := Null, --��¼���ݣ��������Ϊ�գ��������ǰ������  36��36/37
                    strSQL = strSQL & "'" & strValue & "',"
                    '���²�λ_In In ���˻�����ϸ.���²�λ%Type := Null, --ɾ������ʱ������д��λ �����Ŀ��
                    strSQL = strSQL & IIf(strValue <> "", "'" & Nvl(!��λ) & "'", "NULL") & ","
                    '���Ժϸ�_In In Number := 0,
                    strSQL = strSQL & IIf(lngItemCode = gint���� And strValue <> "", Val(!���Ժϸ�), "0") & ","
                    'δ��˵��_In In ���˻�����ϸ.δ��˵��%Type := Null, --δ��˵��
                    strSQL = strSQL & "'" & strδ�� & "',"
                    '���˼�¼_In In Number := 1,
                    strSQL = strSQL & "1,"
                    '������Դ_In In ���˻�����ϸ.������Դ%Type := 0,
                    strSQL = strSQL & IIf(Val(!������Դ) = 0, 0, !������Դ) & ","
                    '��Դid_In   In ���˻�����ϸ.��Դid%Type := Null,
                    strSQL = strSQL & IIf(Val(!��ԴID) = 0, "NULL", !��ԴID) & ","
                    '����_In     In ���˻�����ϸ.����%Type := 0,
                    strSQL = strSQL & Val(!����)
                    '  ��Ŀ�״�_In In Number := 0,--������Ŀʹ�ã���������ǰ�Ƿ���ɾ��һ��ʱ���ڵ�������Ϣ�� 1 ɾ��
                    '  ��ʼʱ��_In In ���˻�������.����ʱ��%Type := Null, --����¼��Ч��ȵĿ�ʼʱ��
                    '  ����ʱ��_In In ���˻�������.����ʱ��%Type := Null, --����¼��Ч��ȵ���ֹʱ�䣬������¼Ϊÿ���ӣ����±�Ϊ4Сʱ,ʱ�����ڵ���ͬ��Ŀ��¼Ҫɾ��
                    '  ����Ա_IN  IN ���˻�������.������%TYPE := NULL,
                    '  ������_IN IN Number :=1
                    strSQL = strSQL & ",0,NULL,NULL,NULL,"
                    strSQL = strSQL & int������ & ")"
                    
                    arrSQL(ReDimArray(arrSQL)) = strSQL
                    
                   
                End If
            End If
            .MoveNext
        Loop
    End With
    lngOld��¼ID = 0
    '�����Ŀ���棬����ʽ����������һ�� �ȴ���ʱ�䣬���޸�����
    With mrsTableDetail
        Do While Not .EOF
            lngItemCode = Val(!��Ŀ���)
            strValue = Nvl(!���)
            
            mrsCurInfo.Filter = "����='" & strValue & "'"
            If lngItemCode = 4 And zlStr.Nvl(!��Ŀ����) = "Ѫѹ" And Not mrsCurInfo.EOF Then
                strValue = Nvl(!���) & "/" & Nvl(!���)
            End If
            intModify = Val(zlStr.Nvl(!�޸�))
            blnEdit = False
            If intModify = 1 And InStr(1, ",0,3,9,", Val(zlStr.Nvl(!������Դ))) = 0 Then
                blnEdit = False
            Else
                blnEdit = True
            End If
            blnSave = False
            If !״̬ <> 0 Then
                int��Ŀ�״� = 0
                strName = zlStr.Nvl(!��Ŀ����)
                strTmp = GetItemInfo(lngItemCode, strName, lngRow)
                strTmp = strTmp & String(8 - UBound(Split(strTmp, ",")), ",")
                arrTmp = Split(strTmp, ",")
        
                strTime = Format(!ʱ��, "YYYY-MM-DD HH:mm:ss")
                strOldTime = Format(!ԭʼʱ��, "YYYY-MM-DD hh:mm:ss")
                strEnd = strTime
                '���޸����ݷ���ʱ��
                If strOldTime <> strTime And strOldTime <> "" Then
                    mrsRecodeID.Filter = "ʱ��='" & strOldTime & "'"
                    If mrsRecodeID.RecordCount > 0 Then
                        lng��¼ID = Val(mrsRecodeID!��¼ID)
                        '��ͬ��¼�޸ĺ����޸�
                        If lng��¼ID <> lngOld��¼ID Then
                            strSQL = "ZL_���µ�����_����ʱ��("
                            'ID_IN       IN ���˻�������.ID%TYPE,
                            strSQL = strSQL & lng��¼ID & ","
                            '����ʱ��_IN IN ���˻�������.����ʱ��%TYPE
                            strSQL = strSQL & "To_Date('" & strTime & "','yyyy-mm-dd hh24:mi:ss')" & ")"
                            
                            arrSQLTime(ReDimArray(arrSQLTime)) = strSQL
                        End If
                    End If
                End If
                lngOld��¼ID = lng��¼ID
                    
                If Val(!״̬) <> 3 And Val(!״̬) <> 0 Then '״̬Ϊ3����ֻ�޸���ʱ���
                    '���ڻ���������Ҫ���ݻ���ʱ��ɾ����ʱ�ε���������
                    If Val(arrTmp(4)) = 4 Then
                        strTmp = GetAnimalItemTime(lngRow, !�к�, 0, strInfo)
                        If strInfo <> "" Then Exit Function
                        strBegin = Split(strTmp, ";")(0)
                        strEnd = Split(strTmp, ";")(1)
                        If CDate(strTime) < CDate(mstrBTime) Then strTime = Format(mstrBTime, "YYYY-MM-DD HH:mm:ss")
                        If CDate(strTime) > CDate(mstrETime) Then strTime = Format(mstrETime, "YYYY-MM-DD HH:mm:ss")
                        int��Ŀ�״� = 1
                    End If
                    
                    int������ = IIf(ISCheckDept(strTime) = True, 1, 0)
                    strTime = "To_Date('" & strTime & "','yyyy-mm-dd hh24:mi:ss')"
                    
                    '����������Ϣ
                    strSQL = "Zl_���µ�����_Update("
                    '�ļ�id_In   In ���˻����ļ�.Id%Type,  --���˻����ļ�ID
                    strSQL = strSQL & Val(mT_Patient.lng�ļ�ID) & ","
                    '����ʱ��_In In ���˻�������.����ʱ��%Type, --�������ݵķ���ʱ��
                    strSQL = strSQL & strTime & ","
                    '��¼����_In In ���˻�����ϸ.��¼����%Type, --������Ŀ=1���ϱ�˵��=2�����ת���=3�������ձ��=4,�±�˵��=6
                    strSQL = strSQL & Val(Nvl(!��¼����, 1)) & ","
                    '��Ŀ���_In In ���˻�����ϸ.��Ŀ���%Type, --������Ŀ����ţ��ǻ�����Ŀ�̶�Ϊ0
                    strSQL = strSQL & lngItemCode & ","
                    '��¼����_In In ���˻�����ϸ.��¼����%Type := Null, --��¼���ݣ��������Ϊ�գ��������ǰ������  36��36/37
                    strSQL = strSQL & "'" & strValue & "',"
                    '���²�λ_In In ���˻�����ϸ.���²�λ%Type := Null, --ɾ������ʱ������д��λ �����Ŀ��
                    strSQL = strSQL & IIf(Val(arrTmp(5)) = 2, "'" & Nvl(!���²�λ) & "'", "NULL") & ","
                    '���Ժϸ�_In In Number := 0,
                    strSQL = strSQL & IIf(lngItemCode = gint���� And strValue <> "", Val(!���Ժϸ�), "0") & ","
                    'δ��˵��_In In ���˻�����ϸ.δ��˵��%Type := Null, --δ��˵��
                    If Val(arrTmp(1)) = 1 And Val(arrTmp(5)) = 2 Then
                        strSQL = strSQL & "'" & IIf(strValue = "", "", Val(!δ��˵��)) & "',"
                    Else
                        strSQL = strSQL & "NUll,"
                    End If
                    '���˼�¼_In In Number := 1,
                    strSQL = strSQL & "1,"
                    '������Դ_In In ���˻�����ϸ.������Դ%Type := 0,
                    strSQL = strSQL & Val(!������Դ) & ","
                    '��Դid_In   In ���˻�����ϸ.��Դid%Type := Null,
                    strSQL = strSQL & IIf(Val(!��ԴID) = 0, "NULL", !��ԴID) & ","
                    '����_In     In ���˻�����ϸ.����%Type := 0,
                    strSQL = strSQL & Val(!����) & ","
                    '��Ŀ�״�_In In Number := 0,--������Ŀʹ�ã���������ǰ�Ƿ���ɾ��һ��ʱ���ڵ�������Ϣ�� 1 ɾ��
                    strSQL = strSQL & int��Ŀ�״� & ","
                    '��ʼʱ��_In In ���˻�������.����ʱ��%Type := Null,
                    strSQL = strSQL & "To_Date('" & strBegin & "','yyyy-mm-dd hh24:mi:ss'),"
                    '����ʱ��_In In ���˻�������.����ʱ��%Type := Null --����¼��Ч��ȵ���ֹʱ�䣬������¼Ϊÿ���ӣ����±�Ϊ4Сʱ,ʱ�����ڵ���ͬ��Ŀ��¼Ҫɾ��
                    strSQL = strSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss')"
                    '  ����Ա_IN  IN ���˻�������.������%TYPE := NULL,
                    '  ������_IN IN Number :=1
                    strSQL = strSQL & ",NULL," & int������ & ")"
                    
                    arrSQL(ReDimArray(arrSQL)) = strSQL
                End If
            End If
            .MoveNext
        Loop
    End With
'
     '���±���Ϣ
    mrsNote.Filter = 0
    mrsNote.Sort = "ʱ��"
    With mrsNote
        Do While Not .EOF
        lngItemCode = Val(!��¼����)
        
        If Val(!״̬) <> 3 And Val(!״̬) <> 0 Then
            strTime = Format(!ʱ��, "YYYY-MM-DD HH:mm:ss")
            strValue = zlStr.Nvl(!����)
            int��Ŀ�״� = 1
            int������ = IIf(ISCheckDept(strTime) = True, 1, 0)
            strTime = "To_Date('" & strTime & "','yyyy-mm-dd hh24:mi:ss')"
            strBegin = Format(mstrBegin, "YYYY-MM-DD HH:mm:ss")
            strEnd = Format(mstrEnd, "YYYY-MM-DD HH:mm:ss")
            
             '����������Ϣ
            strSQL = "Zl_���µ�����_Update("
            '�ļ�id_In   In ���˻����ļ�.Id%Type,  --���˻����ļ�ID
            strSQL = strSQL & Val(mT_Patient.lng�ļ�ID) & ","
            '����ʱ��_In In ���˻�������.����ʱ��%Type, --�������ݵķ���ʱ��
            strSQL = strSQL & strTime & ","
            '��¼����_In In ���˻�����ϸ.��¼����%Type, --������Ŀ=1���ϱ�˵��=2�����ת���=3�������ձ��=4,�±�˵��=6
            strSQL = strSQL & lngItemCode & ","
            '��Ŀ���_In In ���˻�����ϸ.��Ŀ���%Type, --������Ŀ����ţ��ǻ�����Ŀ�̶�Ϊ0
            strSQL = strSQL & 0 & ","
            '��¼����_In In ���˻�����ϸ.��¼����%Type := Null, --��¼���ݣ��������Ϊ�գ��������ǰ������  36��36/37
            strSQL = strSQL & "'" & strValue & "',"
            '���²�λ_In In ���˻�����ϸ.���²�λ%Type := Null, --ɾ������ʱ������д��λ �����Ŀ��
            strSQL = strSQL & "NULL,"
            '���Ժϸ�_In In Number := 0,
            strSQL = strSQL & "NULL,"
            'δ��˵��_In In ���˻�����ϸ.δ��˵��%Type := Null, --δ��˵��
            strSQL = strSQL & IIf(lngItemCode <> 4, "'" & Nvl(!δ��˵��) & "'", "NULL") & ","
            '���˼�¼_In In Number := 1,
            strSQL = strSQL & "1,"
            '������Դ_In In ���˻�����ϸ.������Դ%Type := 0,
            strSQL = strSQL & Val(!������Դ) & ","
            '��Դid_In   In ���˻�����ϸ.��Դid%Type := Null,
            strSQL = strSQL & IIf(Val(!��ԴID) = 0, "NULL", !��ԴID) & ","
            '����_In     In ���˻�����ϸ.����%Type := 0,
            strSQL = strSQL & Val(!����) & ","
            '��Ŀ�״�_In In Number := 0,--������Ŀʹ�ã���������ǰ�Ƿ���ɾ��һ��ʱ���ڵ�������Ϣ�� 1 ɾ��
            strSQL = strSQL & int��Ŀ�״� & ","
            '��ʼʱ��_In In ���˻�������.����ʱ��%Type := Null,
            strSQL = strSQL & "To_Date('" & strBegin & "','yyyy-mm-dd hh24:mi:ss'),"
            '����ʱ��_In In ���˻�������.����ʱ��%Type := Null --����¼��Ч��ȵ���ֹʱ�䣬������¼Ϊÿ���ӣ����±�Ϊ4Сʱ,ʱ�����ڵ���ͬ��Ŀ��¼Ҫɾ��
            strSQL = strSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss')"
            '  ����Ա_IN  IN ���˻�������.������%TYPE := NULL,
            '  ������_IN IN Number :=1
            strSQL = strSQL & ",NULL," & int������ & ")"
            arrSQL(ReDimArray(arrSQL)) = strSQL
        End If
        .MoveNext
        Loop
    End With
    
    gcnOracle.BeginTrans
    blnTran = True
    '��ִ��ʱ��仯
    For i = 1 To UBound(arrSQLTime)
        If arrSQLTime(i) <> "" Then Call zlDatabase.ExecuteProcedure(CStr(arrSQLTime(i)), "����ʱ������"):
    Next
    '��ִ�����ݱ仯
    For i = 1 To UBound(arrSQL)
        If arrSQL(i) <> "" Then Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "������������"):
    Next
    
    '���ִ����ʾ�仯
     For i = 1 To UBound(arrSQLShow)
        If arrSQLShow(i) <> "" Then Call zlDatabase.ExecuteProcedure(CStr(arrSQLShow(i)), "������ʾ�޸�"):
    Next
    gcnOracle.CommitTrans
    blnTran = False
    mblnOK = True
    SaveData = True
    Screen.MousePointer = 0
     
    Exit Function
Errhand:
    If blnTran = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Screen.MousePointer = 0
    Call SaveErrLog
    
End Function

Private Function ISCheckDept(ByVal str����ʱ�� As String) As Boolean
    '------------------------------------------------
    '���ܣ��Ƿ���Zl_���µ�����_Update�н��п��Ҽ��
    'mstrOverDate<=mstrETime ���Ҳ����Ѿ���Ժ���϶��ǲ��˳�Ժʱ�����Ժʱ����һ�У��������Ľ��
    '------------------------------------------------
    If mbln��Ժ = True And Format(mstrOverDate, "YYYY-MM-DD HH:mm:ss") < Format(mstrETime, "YYYY-MM-DD HH:mm:ss") Then
        If Format(str����ʱ��, "YYYY-MM-DD HH:mm:ss") > Format(mstrOverDate, "YYYY-MM-DD HH:mm:ss") And Format(str����ʱ��, "YYYY-MM-DD HH:mm:ss") <= Format(mstrETime, "YYYY-MM-DD HH:mm:ss") Then
            ISCheckDept = False
        Else
            ISCheckDept = True
        End If
    Else
        ISCheckDept = True
    End If
End Function


Private Function GetItemInfo(ByVal lngItemNO As Long, ByVal strName As String, ByRef lngRow As Long) As String
'---------------------------------------------------------------
'����:��ȡ��Ŀ��Ϣ
'---------------------------------------------------------------
    Dim intRow As Integer
    Dim strValue As String
    
    On Error GoTo Errhand
    For intRow = vsfTab.FixedRows To vsfTab.Rows - 1
        If Val(vsfTab.TextMatrix(intRow, COL_tab��Ŀ���)) = lngItemNO And vsfTab.TextMatrix(intRow, COL_tab��Ŀ��) = strName And intRow Mod 2 <> 1 Then
            Exit For
        End If
    Next intRow
    
    If intRow >= vsfTab.Rows Then
        For intRow = vsfTab.FixedRows To vsfTab.Rows - 1
            If Val(vsfTab.TextMatrix(intRow, COL_tab��Ŀ���)) = lngItemNO Then
                Exit For
            End If
        Next intRow
    End If
    
    If intRow < vsfTab.Rows Then
        strValue = vsfTab.TextMatrix(intRow, COL_tab�ַ���)
    End If
    lngRow = intRow
    GetItemInfo = strValue
    
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function zlRefreshData(Optional ByVal blnCurve As Boolean = True, Optional ByVal blnTab As Boolean = True) As Boolean
    '--------------------------------------------------------------------------------
    '����:��ȡһ��ʱ���ڵ�������������
    '���� blnCurve�Ƿ�ˢ����������
    '--------------------------------------------------------------------------------
    Dim strFields As String, strValues As String, strPara As String
    Dim strSQL As String
    Dim strTime As String
    Dim strCenterTime As String '�м�ʱ��
    Dim strBTime As String '��ǰһ���ʱ��
    Dim strETime As String
    Dim dtBegin As String, dtEnd As String
    Dim strItems As String
    Dim strName As String
    Dim strItemName As String '��Ŀ���ַ���
    Dim str��Ŀ���� As String 'һ����Ŀ��
    Dim strPart As String
    Dim strParam As String
    Dim int��� As Integer, intModify As Integer
    Dim int������Դ As Integer
    Dim intRow As Integer, intNum As Integer
    Dim lng��Ŀ��� As Long, int��� As Integer
    Dim blnAdd As Boolean '�Ƿ����
    Dim rsTemp As New ADODB.Recordset   '��ѯ���ݼ�
    Dim rsCurve As New ADODB.Recordset '��ʱ��¼��
    Dim rstab As New ADODB.Recordset  '��ʱ���ݼ�
    
    On Err GoTo Errhand
    If blnCurve = False And blnTab = False Then Exit Function
    
    lblTime.Caption = Format(mstrBegin, "HH:mm") & "��" & Format(mstrEnd, "HH:mm")
    
    '��ʼ����¼��
    gstrFields = "��¼ID," & adDouble & ",18|ʱ��," & adLongVarChar & ",20"
    Call Record_Init(mrsRecodeID, gstrFields)
    
    gstrFields = "���," & adDouble & ",18|������," & adLongVarChar & ",40|��ֵ," & adLongVarChar & ",400|��λ," & adLongVarChar & ",200|" & _
         "���," & adDouble & ",1|ʱ��," & adLongVarChar & ",20|ԭʼʱ��," & adLongVarChar & ",20|��Ŀ���," & adDouble & ",18|��Ŀ����," & adLongVarChar & ",40|" & _
         "���Ժϸ�," & adDouble & ",1|δ��˵��," & adLongVarChar & ",20|������Դ," & adDouble & ",1|�޸�," & adDouble & ",1|��ʾ," & adDouble & ",1|ԭʼ��ʾ״̬," & adDouble & ",1|" & _
         "��ԴID," & adDouble & ",18|����," & adDouble & ",1|״̬," & adDouble & ",1|�к�," & adDouble & ",1|��¼����," & adDouble & ",1"
    Call Record_Init(rsCurve, gstrFields)
    Call Record_Init(mrsCurve, gstrFields)
    Call Record_Init(mrsTable, gstrFields)
    gstrFields = "ID," & adDouble & ",18|������," & adLongVarChar & ",40|���," & adLongVarChar & ",400|���²�λ," & adLongVarChar & ",200|" & _
         "���," & adDouble & ",1|ʱ��," & adLongVarChar & ",20|ԭʼʱ��," & adLongVarChar & ",20|��Ŀ���," & adDouble & ",18|��Ŀ����," & adLongVarChar & ",40|" & _
         "���Ժϸ�," & adDouble & ",1|δ��˵��," & adLongVarChar & ",20|������Դ," & adDouble & ",1|�޸�," & adDouble & ",1|��ʾ," & adDouble & ",1|ԭʼ��ʾ״̬," & adDouble & ",1|" & _
         "��ԴID," & adDouble & ",18|����," & adDouble & ",1|״̬," & adDouble & ",1|�к�," & adDouble & ",1|��¼����," & adDouble & ",1"
    Call Record_Init(mrsTableDetail, gstrFields)
    gstrFields = "���|������|��ֵ|��λ|���|ʱ��|ԭʼʱ��|��Ŀ���|��Ŀ����|���Ժϸ�|δ��˵��|������Դ|�޸�|��ʾ|ԭʼ��ʾ״̬|��ԴID|����|״̬|�к�|��¼����"
    'ˢ�������������ݺ����±���Ϣ
    If blnCurve Then
        strBTime = dtpDate.Value & " 00:00:00"
        strETime = dtpDate.Value & " 23:59:59"
        strSQL = _
            " SELECT /*+ RULE */ C.ID ���,C.��¼ID,A.����ʱ�� As ʱ��,'1)����������Ŀ' ������,C.��ʾ,c.��¼���� As ��ֵ,c.���²�λ,c.���Ժϸ�,D.��¼��,D.��Ŀ���,DECODE(D.��Ŀ���,-1,1,C.��¼���) ��¼���,C.δ��˵��,C.������Դ,C.��ԴID,C.����" & vbNewLine & _
            "                    FROM ���˻����ļ� B,���˻������� A,���˻�����ϸ C,���¼�¼��Ŀ D,�����¼��Ŀ E,Table(Cast(f_num2list([7]) As zlTools.t_Numlist)) F" & vbNewLine & _
            "                    Where B.ID=A.�ļ�ID" & vbNewLine & _
            "                        AND A.ID = C.��¼ID" & vbNewLine & _
            "                        AND B.ID=[1]" & vbNewLine & _
            "                        AND Nvl(B.Ӥ��,0)=[4]" & vbNewLine & _
            "                        AND B.����id=[2]" & vbNewLine & _
            "                        AND B.��ҳid=[3]" & vbNewLine & _
            "                        AND D.��Ŀ���=C.��Ŀ���" & vbNewLine & _
            "                        AND C.��¼����=1" & vbNewLine & _
            "                        AND E.��Ŀ���=D.��Ŀ���" & vbNewLine & _
            "                        AND E.��Ŀ���=F.COLUMN_VALUE" & vbNewLine & _
            "                        AND (NVL(D.��¼��,1)<>2 OR (NVL(D.��¼��,1)=2 And D.��Ŀ���=3))" & _
            "                        And A.����ʱ�� BETWEEN [5] And [6] And C.��ֹ�汾 Is Null" & vbNewLine & _
            "                    Order By A.����ʱ��,DECODE(D.��Ŀ���,-1,1,0),DECODE(D.��Ŀ���,-1,1,C.��¼���)"
            If mblnMove Then
                mstrSQL = Replace(strSQL, "���˻����ļ�", "H���˻����ļ�")
                mstrSQL = Replace(strSQL, "���˻�������", "H���˻�������")
                mstrSQL = Replace(strSQL, "���˻�����ϸ", "H���˻�����ϸ")
            End If
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mT_Patient.lng�ļ�ID, mT_Patient.lng����ID, mT_Patient.lng��ҳID, mT_Patient.lngӤ��, _
                 CDate(Format(strBTime, "YYYY-MM-DD HH:mm:ss")), CDate(Format(strETime, "YYYY-MM-DD HH:mm:ss")), mstrCurveItem)

    
        With rsTemp
            Do While Not .EOF
                '��Ӽ�¼��
                Call Record_Update(mrsRecodeID, "��¼ID|ʱ��", Val(Nvl(!��¼ID)) & "|" & Format(zlStr.Nvl(!ʱ��), "YYYY-MM-DD HH:mm:ss"), "��¼ID|" & Val(Nvl(!��¼ID)))
                
                intModify = 0
                If strTime = "" Then strTime = Format(zlStr.Nvl(!ʱ��), "YYYY-MM-DD HH:mm:ss")
                lng��Ŀ��� = zlStr.Nvl(!��Ŀ���)
                Select Case lng��Ŀ���
                    Case gint����
                        int��� = 1
                    Case Else
                        int��� = Val(Nvl(!��¼���))
                End Select
                intModify = IIf(InStr(1, ",0,3,9,", "," & Val(zlStr.Nvl(!������Դ)) & ",") = 0, 1, 0)
                blnAdd = True
                '���ʺ���������ʱ�����������Ӧ��ʱ���Ƿ��������
                If mint����Ӧ�� = 2 And lng��Ŀ��� = -1 Then
                    mrsCurve.Filter = "��Ŀ���=2 and ʱ��='" & Format(zlStr.Nvl(!ʱ��), "YYYY-MM-DD HH:mm:ss") & "'"
                    If mrsCurve.RecordCount > 0 Then
                        strPara = "���|" & mrsCurve("���")
                        strFields = "��ֵ|���|�޸�"
                        
                        If InStr(1, ",0,3,9,", "," & Val(zlStr.Nvl(mrsCurve!������Դ)) & ",") = 0 And InStr(1, ",0,3,9,", "," & Val(zlStr.Nvl(!������Դ)) & ",") = 0 Then
                            intModify = 1
                        Else
                            intModify = 0
                        End If
                        
                        '��������ʱ����Ϊδ��˵��ֻ��ʾ����������Ϊδ��˵��ʱ����ʾδ��˵��
                        If UBound(Split(mrsCurve("��ֵ"), "/")) <> -1 Then
                            If IsNumeric(zlStr.Nvl(!��ֵ)) Then
                                If mbln����������ʾ Then
                                    gstrValues = zlStr.Nvl(!��ֵ) & "/" & Split(mrsCurve("��ֵ"), "/")(0) & "|" & int��� & "|" & intModify
                                Else
                                    gstrValues = Split(mrsCurve("��ֵ"), "/")(0) & "/" & zlStr.Nvl(!��ֵ) & "|" & int��� & "|" & intModify
                                End If
                            Else
                                gstrValues = Split(mrsCurve("��ֵ"), "/")(0) & "|" & int��� & "|0"
                            End If
                        Else
                            gstrValues = mrsCurve("��ֵ") & "|1|0"
                        End If
                        
                        Call Record_Update(mrsCurve, strFields, gstrValues, strPara)
                        blnAdd = False
                    Else
                        lng��Ŀ��� = 2
                    End If
                End If
                
                '���������¡���ʹ��ʹ
                If (lng��Ŀ��� = gint���� Or lng��Ŀ��� = gint��ʹǿ��) And int��� = 1 Then
                    mrsCurve.Filter = "״̬<> 3 and  ��Ŀ���=" & lng��Ŀ��� & " and ʱ��='" & Format(zlStr.Nvl(!ʱ��), "YYYY-MM-DD HH:mm:ss") & "' and ���<>1"
                    If mrsCurve.RecordCount > 0 Then
                        If InStr(1, ",0,3,9,", "," & Val(zlStr.Nvl(mrsCurve!������Դ)) & ",") = 0 And InStr(1, ",0,3,9,", "," & Val(zlStr.Nvl(!������Դ)) & ",") = 0 Then
                            intModify = 1
                        Else
                            intModify = 0
                        End If
                        
                        strPara = "���|" & mrsCurve("���")
                        strFields = "��ֵ|���|�޸�"
                        gstrValues = Split(mrsCurve("��ֵ"), "/")(0) & "/" & zlStr.Nvl(!��ֵ) & "|" & int��� & "|" & intModify
                        Call Record_Update(mrsCurve, strFields, gstrValues, strPara)
                    End If
                    blnAdd = False
                End If
                
                If blnAdd Then
                    '����������ʾ����
                    strPart = GetPart(lng��Ŀ���)
                    int������Դ = Val(zlStr.Nvl(!������Դ, 0))
                    If Trim(Replace(zlStr.Nvl(!��ֵ), "/", "")) = "" Then
                        int������Դ = 0
                    End If
                    gstrValues = zlStr.Nvl(!���) & "|" & zlStr.Nvl(!������) & "|" & Trim(Replace(zlStr.Nvl(!��ֵ), "/", "")) & "|" & _
                        zlStr.Nvl(!���²�λ, strPart) & "|" & int��� & "|" & Format(zlStr.Nvl(!ʱ��), "YYYY-MM-DD HH:mm:ss") & "|" & _
                        Format(zlStr.Nvl(!ʱ��), "YYYY-MM-DD HH:mm:ss") & "|" & lng��Ŀ��� & "|" & zlStr.Nvl(!��¼��) & "|" & Val(zlStr.Nvl(!���Ժϸ�, 0)) & "|" & _
                        zlStr.Nvl(!δ��˵��) & "|" & int������Դ & "|" & intModify & "|" & Val(zlStr.Nvl(!��ʾ, 0)) & "|" & Val(zlStr.Nvl(!��ʾ, 0)) & "|" & Val(zlStr.Nvl(!��ԴID, 0)) & "|" & Val(zlStr.Nvl(!����, 0)) & "|0|0|1"
                    Call Record_Add(mrsCurve, gstrFields, gstrValues)
                End If
            .MoveNext
            Loop
        End With
        
        Call ShowCurve
        
        gstrFields = "���," & adDouble & ",18|��Ŀ���," & adDouble & ",18|ʱ��," & adLongVarChar & ",20|ԭʼʱ��," & adLongVarChar & ",20|��¼����," & adDouble & ",1|����," & _
                adLongVarChar & ",100|��Ŀ����," & adLongVarChar & ",20|δ��˵��," & adLongVarChar & ",20|��¼���," & adDouble & ",1|������Դ," & adDouble & ",1|��ʾ," & adDouble & ",1|" & _
                "��ԴID," & adDouble & ",18|����," & adDouble & ",1|״̬," & adDouble & ",1"
        Call Record_Init(mrsNote, gstrFields)
        gstrFields = "���|��Ŀ���|ʱ��|ԭʼʱ��|��¼����|����|��Ŀ����|δ��˵��|��¼���|������Դ|��ʾ|��ԴID|����|״̬"
        
        mstrSQL = "" & _
             " Select C.ID ���, B.����ʱ�� AS ʱ��,C.��¼����,C.��Ŀ���,C.δ��˵��,C.��¼����,C.��¼���,C.��Ŀ����,C.������Դ,C.��ʾ,C.��ԴID,C.����" & _
             " FROM ���˻����ļ� A, ���˻������� B, ���˻�����ϸ C" & _
             " Where A.ID=B.�ļ�ID and  B.ID = C.��¼ID AND A.ID=[1]  AND Nvl(A.Ӥ��, 0)=[4] AND a.����id=[2] AND a.��ҳid=[3] And c.��ֹ�汾 Is Null" & _
             " AND c.��¼���� in (2,6)  AND B.����ʱ�� BETWEEN [5]  And [6]"
             
        If mblnMove Then
            mstrSQL = Replace(mstrSQL, "���˻����ļ�", "H���˻����ļ�")
            mstrSQL = Replace(mstrSQL, "���˻�������", "H���˻�������")
            mstrSQL = Replace(mstrSQL, "���˻�����ϸ", "H���˻�����ϸ")
        End If
        
        Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "��ȡ���±����Ϣ", mT_Patient.lng�ļ�ID, mT_Patient.lng����ID, mT_Patient.lng��ҳID, _
            mT_Patient.lngӤ��, CDate(strBTime), CDate(strETime))
        With rsTemp
            Do While Not .EOF
                gstrValues = zlStr.Nvl(!���) & "|" & zlStr.Nvl(!��Ŀ���, 0) & "|" & Format(zlStr.Nvl(!ʱ��), "YYYY-MM-DD HH:mm:ss") & "|" & Format(zlStr.Nvl(!ʱ��), "YYYY-MM-DD HH:mm:ss") & "|" & zlStr.Nvl(!��¼����) & "|" & _
                    zlStr.Nvl(!��¼����) & "|" & zlStr.Nvl(!��Ŀ����) & "|" & Nvl(!δ��˵��) & "|" & zlStr.Nvl(!��¼���, 0) & "|" & Val(zlStr.Nvl(!������Դ, 0)) & "|" & _
                    Val(zlStr.Nvl(!��ʾ, 0)) & "|" & Val(zlStr.Nvl(!��ԴID, 0)) & "|" & Val(zlStr.Nvl(!����, 0)) & "|0"
                Call Record_Add(mrsNote, gstrFields, gstrValues)
            .MoveNext
            Loop
        End With
        
        '������±���Ϣ
        Call ShowTabUpDown
    End If
        
        '��ȡ�������
    If blnTab Then
        strItems = ""
        If vsfTab.Tag <> "NO" Then
            For intRow = vsfTab.FixedRows To vsfTab.Rows - 1
                lng��Ŀ��� = Val(vsfTab.TextMatrix(intRow, COL_tab��Ŀ���))
                If lng��Ŀ��� <> 4 Then
                    strItemName = vsfTab.TextMatrix(intRow, COL_tab��Ŀ��)
                    If InStr(1, "," & strItems & ",", ",'" & strItemName & "',") = 0 Then
                        strItems = strItems & ",'" & strItemName & "'"
                    End If
                End If
            Next
            If Left(strItems, 1) = "," Then strItems = Mid(strItems, 2)
            strItems = strItems & ",'����ѹ','����ѹ'"
            If Left(strItems, 1) = "," Then strItems = Mid(strItems, 2)
            '��ȡһ����(���ܺ��еڶ�������)���еı��������Ϣ
            mstrSQL = "SELECT C.Id,a.����ʱ�� As ʱ��,C.��¼ID,C.��¼����,C.��ʾ,C.��¼���� As ���,C.���²�λ,C.δ��˵��,nvl(C.������Դ,0) ������Դ," & vbNewLine & _
                "  DECODE(E.��Ŀ����,2,C.���²�λ || D.��¼��,D.��¼��) ��Ŀ����,D.��Ŀ���,C.��ԴID,C.����,E.��Ŀ���� " & _
                "  FROM ���˻����ļ� B, ���˻������� A,���˻�����ϸ C,���¼�¼��Ŀ D,�����¼��Ŀ E " & _
                "  Where B.ID = A.�ļ�ID" & vbNewLine & _
                "  AND A.ID = C.��¼ID" & vbNewLine & _
                "  AND B.ID = [1]" & vbNewLine & _
                "  AND Nvl(B.Ӥ��, 0) = [7]" & vbNewLine & _
                "  AND B.����id = [2]" & vbNewLine & _
                "  AND B.��ҳid = [3]" & vbNewLine & _
                "  AND INSTR([6], DECODE(E.��Ŀ����, 2,C.���²�λ || D.��¼��, D.��¼��)) > 0" & vbNewLine & _
                "  AND D.��Ŀ��� = C.��Ŀ���" & vbNewLine & _
                "  AND Mod(c.��¼����,10) = 1" & vbNewLine & _
                "  AND E.��Ŀ��� = D.��Ŀ���" & vbNewLine & _
                "  AND A.����ʱ�� BETWEEN [4] And [5]" & vbNewLine & _
                "  And C.��ֹ�汾 Is Null" & vbNewLine & _
                "  AND D.��¼�� = 2 And D.��Ŀ���<>3" & vbNewLine & _
                "  UNION ALL "
            '��ȡ�����±��Ļ�����Ŀ�����±�������Ŀ������ܴ��ڷ�������Ŀ��
            mstrSQL = mstrSQL & vbNewLine & _
                "  SELECT C.ID,a.����ʱ�� As ʱ��,C.��¼ID,C.��¼����,C.��ʾ,C.��¼���� As ���,C.���²�λ,C.δ��˵��,NVL(C.������Դ,0) ������Դ," & _
                "   Decode(d.��Ŀ����, 2, c.���²�λ || d.��Ŀ����, d.��Ŀ����) ��Ŀ����,D.��Ŀ���,C.��ԴID,C.����,D.��Ŀ����" & _
                "   FROM ���˻����ļ� B, ���˻������� A,���˻�����ϸ C,(SELECT A.��Ŀ���,A.��Ŀ����, A.��Ŀ����,B.����� FROM �����¼��Ŀ A,���������Ŀ B" & vbNewLine & _
                "       WHERE A.��Ŀ���=B.��� AND B.����� Is Not Null  " & vbNewLine & _
                "       AND NVL(A.Ӧ�÷�ʽ,0)=1 AND NVL(A.����ȼ�,0)>=[8] AND NVL(A.���ò���,0) IN (0,[9])" & vbNewLine & _
                "       AND (A.���ÿ���=1 OR (A.���ÿ���=2 AND EXISTS (SELECT 1 FROM �������ÿ��� D WHERE D.��Ŀ���=A.��Ŀ��� AND D.����ID=[10])))) D" & _
                "   Where B.ID=A.�ļ�ID And A.ID = C.��¼ID AND Instr([6], Decode(d.��Ŀ����, 2, c.���²�λ || d.��Ŀ����, d.��Ŀ����)) = 0  AND B.ID=[1]  AND Nvl(B.Ӥ��,0)=[7] " & _
                "   AND B.����id=[2]  AND B.��ҳid=[3]  AND D.��Ŀ���=C.��Ŀ���  AND C.��¼����=1" & _
                "   AND A.����ʱ�� BETWEEN [4] And [5] And C.��ֹ�汾 Is Null"
                
            mstrSQL = _
                "   Select ID,ʱ��,��¼ID,��¼����,��ʾ,���,���²�λ,δ��˵��,������Դ,��Ŀ����,��Ŀ���,��ԴID,����,��Ŀ���� From (" & mstrSQL & ")" & _
                "   Order By  Decode(��Ŀ����,'����ѹ',0,1)," & strItems & ",ʱ��"
            If mblnMove Then
                mstrSQL = Replace(mstrSQL, "���˻����ļ�", "H���˻����ļ�")
                mstrSQL = Replace(mstrSQL, "���˻�������", "H���˻�������")
                mstrSQL = Replace(mstrSQL, "���˻�����ϸ", "H���˻�����ϸ")
            End If
            
            strTime = CDate(Format(dtpDate.Value, "YYYY-MM-DD") & " 23:59:59")
            dtBegin = Int(CDate(dtpDate.Value) - 1)
            dtEnd = CDate(CDate(Format(strTime, "YYYY-MM-DD HH:mm:ss")) + 1)
            If CDate(Format(dtBegin, "YYYY-MM-DD HH:mm:ss")) < CDate(Format(mstrBTime, "YYYY-MM-DD HH:mm:ss")) Then _
                dtBegin = CDate(Format(mstrBTime, "YYYY-MM-DD HH:mm:ss"))
            If CDate(Format(dtEnd, "YYYY-MM-DD HH:mm:ss")) > CDate(Format(mstrETime, "YYYY-MM-DD HH:mm:ss")) Then _
                dtEnd = CDate(Format(mstrETime, "YYYY-MM-DD HH:mm:ss"))
            
            Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, _
                                               mT_Patient.lng�ļ�ID, _
                                               mT_Patient.lng����ID, _
                                               mT_Patient.lng��ҳID, _
                                               CDate(dtBegin), _
                                               CDate(dtEnd), _
                                               strItems, mT_Patient.lngӤ��, mT_Patient.lng����ȼ�, IIf(mT_Patient.lngӤ�� = 0, 1, 2), mT_Patient.lng����ID)
            gstrFields = "Id|������|���|���²�λ|���|ʱ��|ԭʼʱ��|��Ŀ���|��Ŀ����|���Ժϸ�|δ��˵��|������Դ|�޸�|��ʾ|��ԴID|����|״̬|�к�|��¼����"
            
            '��ϸ���ݼ�����
            With rsTemp
                .Sort = "ʱ��,��Ŀ���,id"
                Do While Not .EOF
                    '��Ӽ�¼��
                    Call Record_Update(mrsRecodeID, "��¼ID|ʱ��", Val(Nvl(!��¼ID)) & "|" & Format(zlStr.Nvl(!ʱ��), "YYYY-MM-DD HH:mm:ss"), "��¼ID|" & Val(Nvl(!��¼ID)))
                    blnAdd = False
                    intModify = IIf(InStr(1, ",0,3,9,", "," & Val(zlStr.Nvl(!������Դ)) & ",") = 0, 1, 0)
                    int��� = 0
                    If zlStr.Nvl(!Id) <> intNum Or zlStr.Nvl(!��Ŀ����) <> strName Then
                        intNum = zlStr.Nvl(!��Ŀ���)
                        strName = zlStr.Nvl(!��Ŀ����)
                        '����ѹ/����ѹ
                        If intNum = 4 Or intNum = 5 Then
                            Select Case zlStr.Nvl(!��Ŀ����)
                                Case "����ѹ"
                                    strParam = ""
                                    strParam = zlStr.Nvl(!���)
                                Case "����ѹ"
                                    If InStr(strParam, "/") > 0 Then
                                        strParam = strParam & zlStr.Nvl(!���)
                                    Else
                                        strParam = strParam & "/" & zlStr.Nvl(!���)
                                    End If
                                    mrsCurInfo.Filter = "����='" & Nvl(!���) & "'"
                                    If Not mrsCurInfo.EOF Then
                                        strParam = zlStr.Nvl(!���)
                                    End If
                                    If strParam = "/" Then strParam = ""
                                    blnAdd = True
                                    intNum = 4
                                    strName = "Ѫѹ"
                            End Select
                        Else
                            strParam = zlStr.Nvl(!���)
                            blnAdd = True
                        End If
                        
                        If !��¼���� = 11 Then blnAdd = False
                        
                        If blnAdd = True Then
                            gstrValues = zlStr.Nvl(!Id) & "|2)���±����Ŀ|" & strParam & "|" & _
                                            zlStr.Nvl(!���²�λ) & "|0|" & Format(zlStr.Nvl(!ʱ��), "YYYY-MM-DD HH:mm:ss") & "|" & _
                                            Format(zlStr.Nvl(!ʱ��), "YYYY-MM-DD HH:mm:ss") & "|" & intNum & "|" & strName & "|0|" & _
                                            zlStr.Nvl(!δ��˵��) & "|" & Val(zlStr.Nvl(!������Դ, 0)) & "|" & intModify & "|" & Val(zlStr.Nvl(!��ʾ, 0)) & "|" & _
                                            Val(zlStr.Nvl(!��ԴID, 0)) & "|" & Val(zlStr.Nvl(!����, 0)) & "|0|" & int��� & "|1"

                            Call Record_Add(mrsTableDetail, gstrFields, gstrValues)
                        End If
                    End If
                    .MoveNext
                Loop
            End With
            
            gbln��Ժ = mbln��Ժ
            Call InitTableData(rsTemp)
            Call ShowTable
        End If
    End If
        
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function InitTableData(ByVal rsTemp As ADODB.Recordset)
    '---------------------------------------------------------
    '����:��ʼ�����±������
    '---------------------------------------------------------
    Dim int��Ŀ���� As Integer, int��¼Ƶ�� As Integer, int��Ŀ��ʾ As Integer, int��Ժ�ײ� As Integer
    Dim int��� As Integer, intNum As Integer
    Dim intRow As Integer, intModify As Integer
    Dim lng��Ŀ��� As Long
    Dim blnAdd As Boolean
    Dim strPart As String '��λ
    Dim strParam As String, strFields As String, strValues As String
    Dim str��Ŀ���� As String, strName As String
    Dim rstab As New ADODB.Recordset
    
    On Error GoTo Errhand
    gstrFields = "���," & adDouble & ",18|������," & adLongVarChar & ",40|��ֵ," & adLongVarChar & ",400|��λ," & adLongVarChar & ",200|" & _
         "���," & adDouble & ",1|ʱ��," & adLongVarChar & ",20|ԭʼʱ��," & adLongVarChar & ",20|��Ŀ���," & adDouble & ",18|��Ŀ����," & adLongVarChar & ",40|" & _
         "���Ժϸ�," & adDouble & ",1|δ��˵��," & adLongVarChar & ",20|������Դ," & adDouble & ",1|�޸�," & adDouble & ",1|��ʾ," & adDouble & ",1|ԭʼ��ʾ״̬," & adDouble & ",1|" & _
         "��ԴID," & adDouble & ",18|����," & adDouble & ",1|״̬," & adDouble & ",1|�к�," & adDouble & ",1|��¼����," & adDouble & ",1"
    Call Record_Init(mrsTable, gstrFields)
    strFields = "���|������|��ֵ|��λ|���|ʱ��|ԭʼʱ��|��Ŀ���|��Ŀ����|���Ժϸ�|δ��˵��|������Դ|�޸�|��ʾ|��ԴID|����|״̬|�к�|��¼����"
    For intRow = vsfTab.FixedRows To vsfTab.Rows - 1
        int��Ŀ���� = Val(Split(vsfTab.TextMatrix(intRow, COL_tab�ַ���), ",")(5))
        int��¼Ƶ�� = Val(Split(vsfTab.TextMatrix(intRow, COL_tab�ַ���), ",")(3))
        int��Ŀ��ʾ = Val(Split(vsfTab.TextMatrix(intRow, COL_tab�ַ���), ",")(4))
        strPart = Split(vsfTab.TextMatrix(intRow, COL_tab�ַ���), ",")(7)
        int��Ժ�ײ� = Val(Split(vsfTab.TextMatrix(intRow, COL_tab�ַ���), ",")(8))
        lng��Ŀ��� = Val(vsfTab.TextMatrix(intRow, COL_tab��Ŀ���))
        str��Ŀ���� = vsfTab.TextMatrix(intRow, COL_tab��Ŀ��)
    
        intNum = 0
        strName = ""
        Set rstab = ReturnItemRecord(rsTemp, Int(CDate(Format(dtpDate.Value, "YYYY-MM-DD hh:mm:ss"))), CDate(mstrBTime), lng��Ŀ��� & ";" & str��Ŀ���� & ";" & _
                       int��¼Ƶ�� & ";" & int��Ŀ��ʾ & ";" & int��Ŀ���� & ";" & int��Ժ�ײ� & ";" & strPart, mbln���ܵ���, mbln¼��Сʱ, True)
        If rstab.RecordCount > 0 Then rstab.MoveFirst
        rstab.Sort = "ʱ��,��Ŀ���,���"
        
        With rstab
            Do While Not .EOF
                blnAdd = False
                intModify = IIf(InStr(1, ",0,3,9,", "," & Val(zlStr.Nvl(!������Դ)) & ",") = 0, 1, 0)
                If zlStr.Nvl(!���) <> intNum Or zlStr.Nvl(!��Ŀ����) <> strName Then
                    intNum = zlStr.Nvl(!��Ŀ���)
                    strName = zlStr.Nvl(!��Ŀ����)
                    '����ѹ/����ѹ
                    If lng��Ŀ��� = 4 And str��Ŀ���� = "Ѫѹ" Then
                        Select Case zlStr.Nvl(!��Ŀ����)
                            Case "����ѹ"
                                strParam = ""
                                strParam = zlStr.Nvl(!��¼����)
                            Case "����ѹ"
                                If InStr(strParam, "/") > 0 Then
                                    strParam = strParam & zlStr.Nvl(!��¼����)
                                Else
                                    strParam = strParam & "/" & zlStr.Nvl(!��¼����)
                                End If
                                'Ѫѹ��ʾ����
                                mrsCurInfo.Filter = "����='" & Nvl(!��¼����) & "'"
                                If Not mrsCurInfo.EOF Then
                                    strParam = zlStr.Nvl(!��¼����)
                                End If
                                If strParam = "/" Then strParam = ""
                                blnAdd = True
                            Case "Ѫѹ"
                                strParam = zlStr.Nvl(!��¼����)
                                blnAdd = True
                        End Select
                    Else
                        strParam = zlStr.Nvl(!��¼����)
                        blnAdd = True
                    End If
                    If blnAdd = True Then
                        '��ȡ����ʱ�Ǹ���ʱ��κ���ʾ˳������ġ����һ��ʱ����ж�������,ֻ��ȡǰһ��
                        mrsTable.Filter = "��Ŀ���=" & lng��Ŀ��� & " and ��Ŀ����='" & str��Ŀ���� & "' and �к�=" & Val(zlStr.Nvl(!���, 0))
                        If mrsTable.RecordCount = 0 Then
                            strValues = zlStr.Nvl(!Id) & "|2)���±����Ŀ|" & strParam & "|" & _
                                    zlStr.Nvl(!���²�λ) & "|0|" & Format(zlStr.Nvl(!ʱ��), "YYYY-MM-DD HH:mm:ss") & "|" & _
                                    Format(zlStr.Nvl(!ʱ��), "YYYY-MM-DD HH:mm:ss") & "|" & lng��Ŀ��� & "|" & str��Ŀ���� & "|0|" & _
                                    zlStr.Nvl(!δ��˵��) & "|" & Val(zlStr.Nvl(!������Դ, 0)) & "|" & intModify & "|" & Val(zlStr.Nvl(!��ʾ, 0)) & "|" & _
                                    Val(zlStr.Nvl(!��ԴID, 0)) & "|" & Val(zlStr.Nvl(!����, 0)) & "|0|" & zlStr.Nvl(!���, 0) & "|1"
                            Call Record_Add(mrsTable, strFields, strValues)
                        End If
                    End If
                End If
            .MoveNext
            Loop
        End With
    Next intRow
    
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Function ShowCurve()
'-----------------------------------------
'���ܣ�չʾ������������
'-----------------------------------------
    Dim intRow As Integer
    Dim strCenterTime As String
    Dim strFields As String, strValues As String, strPara As String
    Dim lngColor As Long
    Dim lng��Ŀ��� As Long
    Dim rsCompara As New ADODB.Recordset
    
    On Err GoTo Errhand
    strFields = "��Ŀ���," & adDouble & ",18|ʱ��," & adLongVarChar & ",20|��ʾ," & adDouble & ",1"
    Call Record_Init(rsCompara, strFields)
    
    With mrsCurve
        .Filter = "״̬<> 3 and ʱ�� >= '" & mstrBegin & "' and ʱ�� <=  '" & mstrEnd & "'"
        strCenterTime = GetCenterTime(mstrBegin, mstrEnd)
        Do While Not .EOF
            lng��Ŀ��� = !��Ŀ���
            rsCompara.Filter = "��Ŀ���=" & lng��Ŀ���
            strFields = "��Ŀ���|ʱ��|��ʾ"
            If rsCompara.RecordCount > 0 Then
                If !��ʾ = 1 Then
                    If rsCompara!��ʾ = 1 Then
                        If CheckShow(!ʱ��, rsCompara!ʱ��, strCenterTime) Then
                            strValues = zlStr.Nvl(!��Ŀ���) & "|" & Format(zlStr.Nvl(!ʱ��), "YYYY-MM-DD HH:mm:ss") & "|" & Val(zlStr.Nvl(!��ʾ, 0))
                            strPara = "��Ŀ���|" & lng��Ŀ���
                            rsCompara.Filter = 0
                            Call Record_Update(rsCompara, strFields, strValues, strPara)
                        End If
                    Else
                        strValues = zlStr.Nvl(!��Ŀ���) & "|" & Format(zlStr.Nvl(!ʱ��), "YYYY-MM-DD HH:mm:ss") & "|" & Val(zlStr.Nvl(!��ʾ, 0))
                        strPara = "��Ŀ���|" & lng��Ŀ���
                        rsCompara.Filter = 0
                        Call Record_Update(rsCompara, strFields, strValues, strPara)
                    End If
                Else
                    If rsCompara!��ʾ = 0 Then
                        If CheckShow(!ʱ��, rsCompara!ʱ��, strCenterTime) Then
                            strValues = lng��Ŀ��� & "|" & Format(zlStr.Nvl(!ʱ��), "YYYY-MM-DD HH:mm:ss") & "|" & Val(zlStr.Nvl(!��ʾ, 0))
                            strPara = "��Ŀ���|" & lng��Ŀ���
                            rsCompara.Filter = 0
                            Call Record_Update(rsCompara, strFields, strValues, strPara)
                        End If
                    End If
                End If
                
            Else
                strValues = lng��Ŀ��� & "|" & Format(zlStr.Nvl(!ʱ��), "YYYY-MM-DD HH:mm:ss") & "|" & Val(zlStr.Nvl(!��ʾ, 0))
                Call Record_Add(rsCompara, strFields, strValues)
            End If
            !��ʾ = 0
            .Update
            .MoveNext
        Loop
    End With
    
     With rsCompara
        .Filter = 0
        Do While Not .EOF
            mrsCurve.Filter = "״̬<> 3  and ��Ŀ���=" & !��Ŀ��� & " and ʱ�� ='" & Format(zlStr.Nvl(!ʱ��), "YYYY-MM-DD HH:mm:ss") & "'"
            mrsCurve!��ʾ = 1
            mrsCurve.Update
            .MoveNext
        Loop
    End With
    
    '��ʾ��������
    mrsCurve.Filter = "״̬<> 3 and ʱ�� >= '" & mstrBegin & "' and ʱ�� <=  '" & mstrEnd & "'"
    mrsCurve.Sort = "ʱ��"
    
    vsfCurve.Cell(flexcpText, vsfCurve.FixedRows, vsfCurve.FixedCols, vsfCurve.Rows - 1, vsfCurve.Cols - 1) = ""
    vsfCurve.Cell(flexcpForeColor, vsfCurve.FixedRows, vsfCurve.FixedCols, vsfCurve.Rows - 1, vsfCurve.Cols - 1) = &H80000012
    vsfCurve.Cell(flexcpBackColor, vsfCurve.FixedRows, vsfCurve.FixedCols, vsfCurve.Rows - 1, vsfCurve.Cols - 1) = &H80000005
    
    For intRow = vsfCurve.FixedRows To vsfCurve.Rows - 1

        vsfCurve.Body.MergeRow(intRow) = True
        vsfCurve.TextMatrix(intRow, COL_����) = Space(intRow) & IIf(vsfCurve.TextMatrix(intRow, COL_������) = "2)���±�˵��", "", "") & Space(intRow)
        vsfCurve.TextMatrix(intRow, COL_��ɫ) = Space(intRow) & IIf(vsfCurve.TextMatrix(intRow, COL_������) = "2)���±�˵��", " ", Space(intRow))
        vsfCurve.TextMatrix(intRow, col_ԭʼʱ��) = ""
        vsfCurve.TextMatrix(intRow, COL_��ʾ) = ""
        vsfCurve.TextMatrix(intRow, COL_�༭) = "0"
        vsfCurve.TextMatrix(intRow, COL_��Դ) = "0"
        vsfCurve.TextMatrix(intRow, COL_ԭֵ) = ""
        vsfCurve.TextMatrix(intRow, COL_�޸�״̬) = "0"
        If vsfCurve.TextMatrix(intRow, COL_������) = "2)���±�˵��" Then
             vsfCurve.Cell(flexcpBackColor, intRow, COL_��ɫ, intRow, COL_��ɫ) = RGB(0, 0, 255)
        End If
    Next intRow
    
    With mrsCurve
        .Filter = "״̬<> 3 and ��ʾ=1 and ʱ�� >= '" & mstrBegin & "' and ʱ�� <=  '" & mstrEnd & "'"
        Do While Not .EOF
                For intRow = vsfCurve.FixedRows To vsfCurve.Rows - 1
                    lng��Ŀ��� = Val(vsfCurve.TextMatrix(intRow, COL_��Ŀ���))
                    If !������ = vsfCurve.TextMatrix(intRow, COL_������) And !��Ŀ��� = lng��Ŀ��� Then
                        vsfCurve.TextMatrix(intRow, COL_�޸�״̬) = zlStr.Nvl(!״̬, 0)
                        vsfCurve.TextMatrix(intRow, col_ԭʼʱ��) = Format(zlStr.Nvl(!ʱ��), "YYYY-MM-DD HH:mm:ss")
                        vsfCurve.TextMatrix(intRow, COL_��ʾ) = IIf(Val(zlStr.Nvl(!��ʾ)) = 1, "��", "")
                        vsfCurve.TextMatrix(intRow, COL_ʱ��) = Format(zlStr.Nvl(!ʱ��), "HH:mm")
                        vsfCurve.TextMatrix(intRow, COL_����) = Space(intRow) & zlStr.Nvl(!��ֵ) & Space(intRow)
                        vsfCurve.TextMatrix(intRow, COL_��ɫ) = vsfCurve.TextMatrix(intRow, COL_����)
                        If Not IsNumeric(zlStr.Nvl(!��ֵ)) And zlStr.Nvl(!��ֵ) <> "����" And InStr(1, zlStr.Nvl(!��ֵ), "/") = 0 Then
                            vsfCurve.TextMatrix(intRow, COL_��λ) = ""
                            vsfCurve.TextMatrix(intRow, Col_δ��˵��) = zlStr.Nvl(!δ��˵��)
                        Else
                            vsfCurve.TextMatrix(intRow, COL_��λ) = zlStr.Nvl(!��λ)
                            vsfCurve.TextMatrix(intRow, Col_δ��˵��) = ""
                        End If
                        If lng��Ŀ��� = gint���� And (IsNumeric(zlStr.Nvl(!��ֵ)) Or zlStr.Nvl(!��ֵ) <> "����") Then
                            vsfCurve.TextMatrix(intRow, COL_���Ժϸ�) = IIf(Val(zlStr.Nvl(!���Ժϸ�)) = 1, "��", "")
                        End If
                        lngColor = 255
                        If InStr(1, ",0,3,9,", Val(zlStr.Nvl(!������Դ))) = 0 Then
                            If zlStr.Nvl(!��ֵ) = "����" And lng��Ŀ��� = gint���� Then
                                lngColor = 255
                            ElseIf lng��Ŀ��� = gint���� Or lng��Ŀ��� = gint��ʹǿ�� Or (lng��Ŀ��� = 2 And mint����Ӧ�� = 2 And mblnEdit���� = True) Then
                                If InStr(1, zlStr.Nvl(!��ֵ), "/") = 0 Then
                                    lngColor = RGB(0, 0, 255)
                                Else
                                    If Val(!�޸�) = 0 Then
                                        lngColor = RGB(0, 0, 255)
                                    Else
                                        lngColor = 255
                                    End If
                                End If
                            End If
                            vsfCurve.Cell(flexcpForeColor, intRow, COL_����, intRow, COL_����) = lngColor
                        Else
                            vsfCurve.Cell(flexcpForeColor, intRow, COL_����, intRow, COL_����) = &H80000012
                        End If
                        vsfCurve.TextMatrix(intRow, COL_��Դ) = Val(CStr(zlStr.Nvl(!������Դ)))
                        vsfCurve.TextMatrix(intRow, COL_ԭֵ) = Val(!��ֵ)
                        If lng��Ŀ��� = 2 And mbln����������ʾ And InStr(!��ֵ, "/") > 0 Then
                            vsfCurve.TextMatrix(intRow, COL_ԭֵ) = Split(!��ֵ, "/")(1)
                        End If
                        vsfCurve.TextMatrix(intRow, COL_�༭) = Val(zlStr.Nvl(!�޸�, 0))
                        
                    End If
                Next intRow
                .MoveNext
            Loop
        End With
        
         Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ShowTabUpDown()
    '--------------------
    '���ܣ��������±�
    '--------------------
    Dim intRow As Integer
    
    On Error GoTo Errhand
    mrsNote.Filter = "״̬<> 3 and ʱ�� >= '" & mstrBegin & "' and ʱ�� <=  '" & mstrEnd & "'"
    With mrsNote
        Do While Not .EOF
                If CDate(Format(!ʱ��, "YYYY-MM-DD HH:mm:ss")) >= CDate(Format(mstrBegin, "YYYY-MM-DD HH:mm:ss")) _
                    And CDate(Format(!ʱ��, "YYYY-MM-DD HH:mm:ss")) <= CDate(Format(mstrEnd, "YYYY-MM-DD HH:mm:ss")) Then
                    Select Case Val(!��¼����)
                        Case 2
                            intRow = mOptRow.�ϱ�
                        Case 6
                            intRow = mOptRow.�±�
                    End Select
                    vsfCurve.TextMatrix(intRow, COL_�޸�״̬) = zlStr.Nvl(!״̬, 0)
                    vsfCurve.TextMatrix(intRow, col_ԭʼʱ��) = Format(zlStr.Nvl(!ʱ��), "YYYY-MM-DD HH:mm:ss")
                    vsfCurve.TextMatrix(intRow, COL_ʱ��) = Format(zlStr.Nvl(!ʱ��), "hh:mm")
                    vsfCurve.TextMatrix(intRow, COL_����) = Space(intRow) & zlStr.Nvl(!����) & Space(intRow)
                    vsfCurve.Cell(flexcpBackColor, intRow, COL_��ɫ, intRow, COL_��ɫ) = IIf(IsNumeric(Nvl(!δ��˵��)) = False, 16711680, Val(Nvl(!δ��˵��)))
                    vsfCurve.TextMatrix(intRow, COL_��λ) = ""
                    vsfCurve.TextMatrix(intRow, COL_���Ժϸ�) = ""
                    vsfCurve.TextMatrix(intRow, Col_δ��˵��) = ""
                End If
        .MoveNext
        Loop
    End With
        
        Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ShowDetail(ByVal lng��Ŀ��� As Long, ByVal intNewRow As Integer)
    '--------------------------------------------------
    '���ܣ�չʾ������ϸ�������
    '--------------------------------------------------
    Dim intRow As Integer
    Dim intMarkRow As Integer
    Dim str�ַ��� As String
    Dim lngColor As Long
    
    On Err GoTo Errhand
    vsfDetail.Rows = vsfDetail.FixedRows
    vsfDetail.Rows = vsfDetail.FixedRows + 1
    If mrsCurve.State = adStateClosed Then Exit Function
    mrsCurve.Filter = "״̬<> 3 and ʱ�� > '" & mstrBegin & "' and ʱ�� <=  '" & mstrEnd & "' and ��Ŀ���=" & lng��Ŀ���
    mrsCurve.Sort = "ʱ��"
    If mblnInit Then
        vsfDetail.ColHidden(COL_���Ժϸ�) = False
        If lng��Ŀ��� <> gint���� Then vsfDetail.Body.ColHidden(COL_���Ժϸ�) = True
    End If
    If mrsCurve.RecordCount > 0 Then
        str�ַ��� = vsfCurve.TextMatrix(intNewRow, COL_�ַ���)
        intMarkRow = 0
        With mrsCurve
            intRow = vsfDetail.FixedRows
            Do While Not .EOF
                vsfDetail.TextMatrix(intRow, COL_�ַ���) = str�ַ���
                vsfDetail.TextMatrix(intRow, COL_��Ŀ���) = lng��Ŀ���
                vsfDetail.TextMatrix(intRow, COL_�޸�״̬) = !״̬
                vsfDetail.TextMatrix(intRow, COL_��ʾ) = IIf(zlStr.Nvl(!��ʾ) = 1, "��", "")
                vsfDetail.TextMatrix(intRow, COL_ʱ��) = Format(zlStr.Nvl(!ʱ��), "hh:mm")
                vsfDetail.TextMatrix(intRow, col_ԭʼʱ��) = Format(zlStr.Nvl(!ʱ��), "YYYY-MM-DD hh:mm:ss")
                vsfDetail.TextMatrix(intRow, COL_����) = zlStr.Nvl(!��ֵ)
                If Not IsNumeric(zlStr.Nvl(!��ֵ)) And zlStr.Nvl(!��ֵ) <> "����" And InStr(1, zlStr.Nvl(!��ֵ), "/") = 0 Then
                    vsfDetail.TextMatrix(intRow, COL_��λ) = ""
                    vsfDetail.TextMatrix(intRow, Col_δ��˵��) = zlStr.Nvl(!δ��˵��)
                Else
                    vsfDetail.TextMatrix(intRow, COL_��λ) = zlStr.Nvl(!��λ)
                    vsfDetail.TextMatrix(intRow, Col_δ��˵��) = ""
                End If
                If lng��Ŀ��� = gint���� And (IsNumeric(zlStr.Nvl(!��ֵ)) Or zlStr.Nvl(!��ֵ) <> "����") Then
                    vsfDetail.TextMatrix(intRow, COL_���Ժϸ�) = IIf(Val(zlStr.Nvl(!���Ժϸ�)) = 1, "��", "")
                End If
                lngColor = 255
                        If InStr(1, ",0,3,9,", Val(zlStr.Nvl(!������Դ))) = 0 Then
                            If zlStr.Nvl(!��ֵ) = "����" And lng��Ŀ��� = gint���� Then
                                lngColor = 255
                            ElseIf lng��Ŀ��� = gint���� Or lng��Ŀ��� = gint��ʹǿ�� Or (lng��Ŀ��� = 2 And mint����Ӧ�� = 2 And mblnEdit���� = True) Then
                                If InStr(1, zlStr.Nvl(!��ֵ), "/") = 0 Then
                                    lngColor = RGB(0, 0, 255)
                                Else
                                    If Val(!�޸�) = 0 Then
                                        lngColor = RGB(0, 0, 255)
                                    Else
                                        lngColor = 255
                                    End If
                                End If
                            End If
                            vsfDetail.Cell(flexcpForeColor, intRow, COL_����, intRow, COL_����) = lngColor
                        Else
                            vsfDetail.Cell(flexcpForeColor, intRow, COL_����, intRow, COL_����) = &H80000012
                        End If
                
                Select Case !������Դ
                    Case 0, 9
                        vsfDetail.TextMatrix(intRow, COL_������Դ) = "���µ�¼��"
                        vsfDetail.TextMatrix(intRow, COL_�༭) = 1
                    Case 1
                        vsfDetail.TextMatrix(intRow, COL_������Դ) = "��¼��ͬ��"
                        vsfDetail.TextMatrix(intRow, COL_�༭) = 0
                    Case 3
                        vsfDetail.TextMatrix(intRow, COL_������Դ) = "�ƶ��豸¼��"
                        vsfDetail.TextMatrix(intRow, COL_�༭) = 1
                    Case Else
                        vsfDetail.TextMatrix(intRow, COL_������Դ) = "�����豸ͬ��"
                        vsfDetail.TextMatrix(intRow, COL_�༭) = 0
                End Select
                
                vsfDetail.TextMatrix(intRow, COL_��Դ) = Val(CStr(zlStr.Nvl(!������Դ)))
                vsfDetail.TextMatrix(intRow, COL_�༭) = Val(zlStr.Nvl(!�޸�, 0))
                vsfDetail.TextMatrix(intRow, COL_ԭֵ) = Val(!��ֵ)
                If lng��Ŀ��� = gint���� And mbln����������ʾ And InStr(!��ֵ, "/") > 0 Then
                    vsfDetail.TextMatrix(intRow, COL_ԭֵ) = Split(!��ֵ, "/")(1)
                End If
                If vsfDetail.TextMatrix(intRow, COL_��ʾ) = "��" Then intMarkRow = intRow
                .MoveNext
                vsfDetail.Body.CellAlignment = flexAlignCenterCenter
                vsfDetail.Body.Cell(flexcpAlignment, intRow, COL_��ʾ, intRow, COL_��ʾ) = flexAlignCenterCenter
                vsfDetail.Body.Cell(flexcpAlignment, intRow, COL_ʱ��, intRow, COL_ʱ��) = flexAlignCenterCenter
                vsfDetail.Body.Cell(flexcpAlignment, intRow, COL_����, intRow, COL_����) = flexAlignCenterCenter
                vsfDetail.Body.Cell(flexcpAlignment, intRow, COL_���Ժϸ�, intRow, COL_���Ժϸ�) = flexAlignCenterCenter
                vsfDetail.Body.Cell(flexcpAlignment, intRow, COL_������Դ, intRow, COL_������Դ) = flexAlignCenterCenter
                intRow = intRow + 1
                vsfDetail.Rows = vsfDetail.Rows + 1
            Loop

        End With
    End If
        If intMarkRow <> 0 Then
            vsfDetail.Row = intMarkRow
            vsfDetail.Col = COL_����
        End If
     
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ShowTable()
    '--------------------------
    '���ܣ�չʾ���±������
    '--------------------------
    Dim strTime As String
    Dim strInfo As String
    Dim intRow As Integer, intƵ�� As Integer
    Dim blnAllow As Boolean, bln���� As Boolean
    Dim lngHour As String
    Dim arrOldTime() As String
    
    On Error GoTo Errhand
    mrsTable.Filter = 0
    mrsTable.Sort = "��Ŀ���,�к�,��¼���� "
    vsfTab.Cell(flexcpText, vsfTab.FixedRows, vsfTab.FixedCols, vsfTab.Rows - 1, vsfTab.Cols - 1) = ""
    strTime = ""
    With mrsTable
        Do While Not .EOF
            For intRow = vsfTab.FixedRows To vsfTab.Rows - 1
                blnAllow = False
                If vsfTab.TextMatrix(intRow, COL_tab��Ŀ���) = !��Ŀ��� Then
                    If Val(Split(vsfTab.TextMatrix(intRow, COL_tab�ַ���), ",")(5)) = 2 Then
                        If vsfTab.TextMatrix(intRow, COL_tab��Ŀ��) <> !��Ŀ���� Then
                            blnAllow = False
                        Else
                            blnAllow = True
                        End If
                    Else
                        blnAllow = True
                    End If
                End If
                intƵ�� = Val(Split(vsfTab.TextMatrix(intRow, COL_tab�ַ���), ",")(3))
                bln���� = Split(vsfTab.RowData(intRow), ";")(1) = 3
                If blnAllow = True Then
                    If InStr(vsfTab.TextMatrix(intRow, col_tabԭʼʱ��), "'") > 0 Then
                        arrOldTime = Split(vsfTab.TextMatrix(intRow, col_tabԭʼʱ��), "'")
                    Else
                        ReDim Preserve arrOldTime(0)
                    End If
                    arrOldTime(!�к� - 1) = Nvl(!ԭʼʱ��)
                    vsfTab.TextMatrix(intRow, col_tabԭʼʱ��) = Join(arrOldTime, "'")
                    If intRow Mod 2 = 0 Then '����
                        If Val(Nvl(!��¼����)) = 1 Then
                            strTime = GetAnimalItemTime(intRow, Val(!�к�), 0, strInfo)
                            If InStr(1, strTime, ";") > 0 Then lngHour = DateDiff("h", CDate(Split(strTime, ";")(0)), CDate(Split(strTime, ";")(1))) + 1
                            If lngHour > 24 Then lngHour = 24
                            If mbln¼��Сʱ And intƵ�� = 1 And bln���� And Not InStr(1, !��ֵ, ")") > 0 Then
                                vsfTab.TextMatrix(intRow, vsfTab.FixedCols + Val(!�к�) - 1) = "(" & lngHour & "h)" & !��ֵ
                            Else
                                vsfTab.TextMatrix(intRow, vsfTab.FixedCols + Val(!�к�) - 1) = !��ֵ
                            End If
                            If Val(zlStr.Nvl(!������Դ)) <> 0 Then
                                vsfTab.Cell(flexcpForeColor, intRow, vsfTab.FixedCols + Val(!�к�) - 1, intRow, vsfTab.FixedCols + Val(!�к�) - 1) = 255
                            Else
                                vsfTab.Cell(flexcpForeColor, intRow, vsfTab.FixedCols + Val(!�к�) - 1, intRow, vsfTab.FixedCols + Val(!�к�) - 1) = &H80000012
                            End If
                            If Val(Split(vsfTab.TextMatrix(intRow, COL_tab�ַ���), ",")(5)) = 2 And Val(Split(vsfTab.TextMatrix(intRow, COL_tab�ַ���), ",")(1)) = 1 And _
                                Val(Split(vsfTab.TextMatrix(intRow, COL_tab�ַ���), ",")(4)) = 0 Then
                                 vsfTab.Cell(flexcpForeColor, intRow, vsfTab.FixedCols + Val(!�к�) - 1, intRow, vsfTab.FixedCols + Val(!�к�) - 1) = Val(zlStr.Nvl(!δ��˵��))
                            End If
                        End If
                    Else
                        If Val(Nvl(!��¼����)) = 1 Then
                            strTime = GetAnimalItemTime(intRow, Val(!�к�), 0, strInfo)
                            If InStr(1, strTime, ";") > 0 Then strTime = Format(Split(strTime, ";")(0), "hh:mm") & "��" & Format(Split(strTime, ";")(1), "hh:mm")
                            vsfTab.TextMatrix(intRow, vsfTab.FixedCols + Val(!�к�) - 1) = IIf(Split(vsfTab.RowData(intRow), ";")(1) = 0, Format(!ʱ��, "hh:mm"), strTime)
                            If Val(zlStr.Nvl(!������Դ)) <> 0 Then
                                vsfTab.Cell(flexcpForeColor, intRow, vsfTab.FixedCols + Val(!�к�) - 1, intRow, vsfTab.FixedCols + Val(!�к�) - 1) = 255
                            Else
                                vsfTab.Cell(flexcpForeColor, intRow, vsfTab.FixedCols + Val(!�к�) - 1, intRow, vsfTab.FixedCols + Val(!�к�) - 1) = &H80000012
                            End If
                            If Val(Split(vsfTab.TextMatrix(intRow, COL_tab�ַ���), ",")(5)) = 2 And Val(Split(vsfTab.TextMatrix(intRow, COL_tab�ַ���), ",")(1)) = 1 And _
                                Val(Split(vsfTab.TextMatrix(intRow, COL_tab�ַ���), ",")(4)) = 0 Then
                                vsfTab.Cell(flexcpForeColor, intRow, vsfTab.FixedCols + Val(!�к�) - 1, intRow, vsfTab.FixedCols + Val(!�к�) - 1) = Val(zlStr.Nvl(!δ��˵��))
                            End If
                        End If
                    End If
                End If
            Next intRow
        .MoveNext
        Loop
    End With
    
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Function ShowTabDetail(ByVal intRow As Integer, ByVal intNewRow As Integer, ByVal intType As Integer)
'--------------------------------------------------
'���ܣ������Ŀ��ϸչʾ
'--------------------------------------------------
    Dim strTime As String
    Dim strInfo As String
    
    On Error GoTo Errhand
    With mrsTableDetail
        Do While Not .EOF
            If Val(Nvl(!��¼����)) = 11 Then
                vsfTabDetail.TextMatrix(intRow, vsfTab.FixedCols) = "(" & !��� & "h)" & vsfTabDetail.TextMatrix(intRow, vsfTabDetail.FixedCols)
            Else
                strTime = GetAnimalItemTime(vsfTab.Row, intRow, 0, strInfo)
                If InStr(1, strTime, ";") > 0 Then strTime = Format(Split(strTime, ";")(0), "hh:mm") & "��" & Format(Split(strTime, ";")(1), "hh:mm")
                If intType = 3 And strInfo <> "" Then lblStb.Caption = strInfo: lblStb.ForeColor = 255: Exit Function
                vsfTabDetail.TextMatrix(intRow, vsfTabDetail.FixedCols) = IIf(intType = 3, strTime, Format(!ʱ��, "hh:mm"))
                vsfTabDetail.TextMatrix(intRow, vsfTabDetail.FixedCols + 1) = !���
                vsfTabDetail.TextMatrix(intRow, col_tabԭʼʱ��) = Nvl(!ʱ��)
                Select Case !������Դ
                    Case 0, 9
                        vsfTabDetail.TextMatrix(intRow, vsfTabDetail.FixedCols + 2) = "���µ�¼��"
                    Case 1
                        vsfTabDetail.TextMatrix(intRow, vsfTabDetail.FixedCols + 2) = "��¼��ͬ��"
                    Case 3
                        vsfTabDetail.TextMatrix(intRow, vsfTabDetail.FixedCols + 2) = "�ƶ��豸¼��"
                    Case Else
                        vsfTabDetail.TextMatrix(intRow, vsfTabDetail.FixedCols + 2) = "�����豸ͬ��"
                End Select
                
                If Val(zlStr.Nvl(!������Դ)) <> 0 Then
                    vsfTabDetail.Cell(flexcpForeColor, intRow, vsfTabDetail.FixedCols, intRow, vsfTabDetail.FixedCols + 2) = 255
                Else
                    vsfTabDetail.Cell(flexcpForeColor, intRow, vsfTabDetail.FixedCols, intRow, vsfTabDetail.FixedCols + 2) = &H80000012
                End If
                If Val(Split(vsfTab.TextMatrix(intNewRow, COL_tab�ַ���), ",")(5)) = 2 And Val(Split(vsfTab.TextMatrix(intNewRow, COL_tab�ַ���), ",")(1)) = 1 And _
                    Val(Split(vsfTab.TextMatrix(intNewRow, COL_tab�ַ���), ",")(4)) = 0 Then
                    vsfTabDetail.Cell(flexcpForeColor, intRow, vsfTabDetail.FixedCols, intRow, vsfTabDetail.FixedCols + 2) = Val(zlStr.Nvl(!δ��˵��))
                End If
            End If
            intRow = intRow + 1
            vsfTabDetail.Rows = vsfTabDetail.Rows + 1
            vsfTabDetail.TextMatrix(intRow, COL_tab��Ŀ��) = zlStr.Nvl(!��Ŀ����)
            vsfTabDetail.TextMatrix(intRow, COL_tab��Ŀ����) = zlStr.Nvl(!��Ŀ����)
            vsfTabDetail.TextMatrix(intRow, COL_tab�ַ���) = vsfTab.TextMatrix(vsfTab.Row, COL_tab�ַ���)
            vsfTabDetail.TextMatrix(intRow, COL_tab��Ŀ���) = vsfTab.TextMatrix(vsfTab.Row, COL_tab��Ŀ���)
            .MoveNext
        Loop
    End With
    
    vsfTabDetail.Tag = vsfTabDetail.Rows - 1
    
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetAnimalItemTime(ByVal intRow As Integer, ByVal intNO As Integer, Optional IntMode As Integer = 0, Optional strInfo As String = "") As String
'--------------------------------------------------------------------------------
'����:��ȡ���±����ĿĳƵ�ε�ʱ��
'arrTime ������Ϣ ���� ��ʼʱ��  ����ʱ��
'������introw ��ǰ��,intNo ���,strInfo ������Ϣ IntMode 1 �����м��ʱ�� 0,���ؿ�ʼʱ��ͽ���ʱ��
'---------------------------------------------------------------------------------
    Dim strTmp As String, lng��Ŀ��� As Long, str��Ŀ���� As String, intƵ�� As Integer
    Dim int��Ŀ��ʾ As String, intType As Integer
    Dim arrStr() As String
    Dim rsTmp As New ADODB.Recordset
    Dim strBegin As String, strEnd As String, strTime As String, strCurrDate As String
    Dim intHour As Integer
    Dim lngRow As Long
    Dim strDate As String
    Dim strReturn As String
    Dim bln���� As Boolean

    On Error GoTo Errhand
    
    strDate = mstrBegin
    strInfo = ""
    lngRow = intRow - vsfTab.FixedRows + 1
    strTmp = vsfTab.TextMatrix(intRow, COL_tab�ַ���)
    lng��Ŀ��� = Val(vsfTab.TextMatrix(intRow, COL_tab��Ŀ���))
    str��Ŀ���� = vsfTab.TextMatrix(intRow, COL_tab��Ŀ��)
    strTmp = strTmp & String(8 - UBound(Split(strTmp, ",")), ",")
    arrStr = Split(strTmp, ",")
    intƵ�� = Val(arrStr(3))
    int��Ŀ��ʾ = Val(arrStr(4))
    
    bln���� = IsWaveItem(lng��Ŀ���)
    
    '����/���� ��Ŀ����=2
    If int��Ŀ��ʾ = 4 Or bln���� Then
        intType = 2
        If intƵ�� = 0 Then
            intƵ�� = 2
        ElseIf intƵ�� > 2 Then
            intƵ�� = 2
        End If
        
        '�ɲ���ȷ������/������Ŀ����¼����������ݻ��ǵ��������
        If Not mbln���ܵ��� Then strDate = CDate(mstrBegin) - 1
    Else
        intType = 1
    End If
    
    
    '�������ͣ�Ƶ�κ���� �������Ҳ�����Ϣ
    mrsTabTime.Filter = "����=" & intType & " and Ƶ��=" & intƵ�� & " and ���=" & intNO
    If mrsTabTime.RecordCount = 0 Then
        strInfo = "���ڻ�����Ŀ����������[" & IIf(intType = 2, "������Ŀ", "���±����Ŀ") & "]ʱ����Ϣ!"
        Exit Function
    End If
    
    With mrsTabTime
        .MoveFirst
        intHour = CInt(24 / intƵ��)
        strBegin = Format(IIf(IsDate(Trim(Nvl(!��ʼ))) = False, (Val(Nvl(!���)) - 1) * intHour & ":00:00", !��ʼ), "HH:mm:ss")
        strEnd = Format(IIf(IsDate(Trim(Nvl(!����))) = False, Val(Nvl(!���)) * intHour - 1 & ":59:59", !����), "HH:mm:ss")
        If intNO = intƵ�� Then
            If strBegin >= strEnd Then
                strBegin = Format(strDate, "YYYY-MM-DD") & " " & strBegin
                strEnd = Format(DateAdd("d", 1, CDate(strDate)), "YYYY-MM-DD") & " " & strEnd
            Else
                strBegin = Format(strDate, "YYYY-MM-DD") & " " & strBegin
                strEnd = Format(strDate, "YYYY-MM-DD") & " " & strEnd
            End If
        Else
            If strBegin >= strEnd Then
                strBegin = Format(strDate, "YYYY-MM-DD") & " " & strBegin
                strEnd = strBegin
            Else
                strBegin = Format(strDate, "YYYY-MM-DD") & " " & strBegin
                strEnd = Format(strDate, "YYYY-MM-DD") & " " & strEnd
            End If
        End If
    End With
    If strBegin < mstrBTime Then strBegin = mstrBTime
    If strEnd > mstrETime Then strEnd = mstrETime
    '��ȡ��ǰʱ��
    strCurrDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
    '��ȡ�м�ʱ��
    intHour = DateDiff("H", CDate(strBegin), CDate(strEnd) + 0.00001) / 2
    strTime = DateAdd("H", intHour, CDate(strBegin)) '�е�ʱ��
    
    If CDate(strCurrDate) >= CDate(strBegin) And CDate(strCurrDate) <= CDate(strEnd) Then
        strTime = strCurrDate
    End If
    '����δ��Ժ���ҷ����˳�Ժҽ�������Ԥ��Ժʱ���ڵ�ǰ¼�����Ӧ��ʱ�䷶Χ�ڣ���С���е�ʱ������Ԥ��Ժʱ��Ϊ׼
    If mbln��Ժ = False And IsDate(mstrPreOutDate) Then
        If Format(mstrPreOutDate, "YYYY-MM-DD HH:mm") >= Format(strBegin, "YYYY-MM-DD HH:mm") And _
            Format(mstrPreOutDate, "YYYY-MM-DD HH:mm") <= Format(strEnd, "YYYY-MM-DD HH:mm") And _
            Format(mstrPreOutDate, "YYYY-MM-DD HH:mm") < Format(strTime, "YYYY-MM-DD HH:mm") Then
            strTime = Format(mstrPreOutDate, "YYYY-MM-DD HH:mm")
        End If
    End If
    
    If CDate(strTime) < CDate(mstrBTime) Then
        strTime = Format(mstrBTime, "YYYY-MM-DD HH:mm:ss")
        If CDate(strTime) > CDate(strEnd) Then
            strInfo = "��" & lngRow & "��[" & str��Ŀ���� & "]�Ľ���ʱ�䣺" & Format(strEnd, "YYYY-MM-DD HH:mm:ss") & "������С��[���µ���ʼʱ�䣺" & Format(mstrBTime, "YYYY-MM-DD HH:mm:ss") & "]��"
            Exit Function
        End If
    End If
    
    If CDate(strTime) > CDate(mstrETime) Then
        strTime = Format(mstrETime, "YYYY-MM-DD HH:mm:ss")
        If CDate(strTime) < CDate(strBegin) Then
            If mbln��Ժ = False Then
                strInfo = "��" & lngRow & "��[" & str��Ŀ���� & "]�Ŀ�ʼʱ�䣺" & Format(strBegin, "YYYY-MM-DD HH:mm:ss") & "���ѳ�������[����¼��������" & mintPreDays & "��]��ָ���ķ�Χ��"
            Else
                strInfo = "��" & lngRow & "��[" & str��Ŀ���� & "]�Ŀ�ʼʱ�䣺" & Format(strBegin, "YYYY-MM-DD HH:mm:ss") & "�����ܴ���[���˳�Ժʱ����ļ�����ʱ�䣺" & Format(mstrETime, "YYYY-MM-DD HH:mm:ss") & "]!"
            End If
            Exit Function
        End If
    End If
    
ErrNext:
    '��鲡��ת�ƺ�Ĳ�¼ʱ��
    If Not IsAllowInput(mT_Patient.lng����ID, mT_Patient.lng��ҳID, mT_Patient.lngӤ��, strEnd, strCurrDate) Then
        strInfo = "��¼����ʱ��[" & strTime & "]����[�������ݲ�¼����Чʱ��:" & mlngHours & "Сʱ]"
        Exit Function
    End If
    
    Select Case IntMode
        Case 0
            strReturn = Format(CDate(strBegin), "YYYY-MM-DD HH:mm:ss") & ";" & Format(CDate(strEnd), "YYYY-MM-DD HH:mm:ss")
        Case 1
            strReturn = Format(CDate(strTime), "YYYY-MM-DD HH:mm:ss")
    End Select
    
    GetAnimalItemTime = strReturn
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetChildNo(ByVal lngNo As Long) As ADODB.Recordset
    '���ܣ���ȡ������Ŀ��ϸ
    Dim strFileds As String
    Dim strValue As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo Errhand
    strFileds = "���," & adDouble & ",18|�����," & adDouble & ",18"
    Call Record_Init(rsTemp, strFileds)
    strFileds = "���|�����"
    mrsCollect.Filter = "����� =" & lngNo
    If mrsCollect.RecordCount > 0 Then
        Do While Not mrsCollect.EOF
            strValue = mrsCollect!��� & "|" & mrsCollect!�����
            Call Record_Add(rsTemp, strFileds, strValue)
            mrsCollect.MoveNext
        Loop
    End If
    Set GetChildNo = rsTemp
    
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetPart(ByVal lng��Ŀ��� As Long) As String
'����:��ȡĬ�ϵ����²�λ
    Dim strPart As String
    mrsPart.Filter = "��Ŀ���=" & lng��Ŀ��� & " and ȱʡ��=1"
    If mrsPart.RecordCount > 0 Then strPart = zlStr.Nvl(mrsPart("��λ"))
    GetPart = strPart
End Function



Private Function CheckShow(ByVal strBegin As String, ByVal strEnd As String, ByVal CenterTime As String) As Boolean
'-------------------------------------------------
'���ܣ��Ա�����ʱ����Ǹ��������յ�ʱ��
'strbegin �Աȵ�ʱ��  strend��ǰʱ��
'--------------------------------------------------
    Dim strTime As String
    Dim blnAllow As Boolean
    
    If Abs(DateDiff("s", CDate(Format(strBegin, "YYYY-MM-DD HH:mm:ss")), CDate(CenterTime))) < Abs(DateDiff("s", CDate(Format(strEnd, "YYYY-MM-DD HH:mm:ss")), CDate(CenterTime))) Then
        blnAllow = True
    Else
        blnAllow = False
    End If
    
    CheckShow = blnAllow
End Function

Private Function UpdateCurveDate(ByVal vsf As Object, ByVal intRow As Integer, ByVal intCOl As Integer, ByVal intType As Integer, _
    Optional blnComList As Boolean = False) As Boolean
    '-------------------------------------------------------------------
    '���ܣ�����������Ŀ��������ϸ�����±�����ݱ��� ��¼������
    '���������������У��У���Դ����Ƿ��������б�
    '-------------------------------------------------------------------
    Dim lng��Ŀ��� As Long, strName As String, strTime As String
    Dim int���Ժϸ� As String
    Dim strEditData As String
    Dim strδ�� As String, str��λ As String
    Dim int�޸�״̬ As Integer
    Dim strData As String
    
    On Err GoTo Errhand
    If intType = 1 Or intType = 3 Then
        If vsf.EditText = vsf.Tag And vsf.EditText <> "" Then vsf.TextMatrix(intRow, COL_�޸�״̬) = 0
        If blnComList = True Then
            str��λ = vsf.EditText
            If str��λ = "" Then str��λ = vsf.TextMatrix(intRow, COL_��λ)
            If str��λ <> vsf.Tag Then vsf.TextMatrix(intRow, COL_�޸�״̬) = 2
        Else
            str��λ = vsf.TextMatrix(intRow, COL_��λ)
        End If
        lng��Ŀ��� = Val(vsfCurve.TextMatrix(vsfCurve.Row, COL_��Ŀ���))
        If intCOl = COL_���� Then
            strData = vsf.EditText
        Else
            strData = Trim(vsf.TextMatrix(intRow, COL_����))
        End If
        strTime = Trim(vsf.TextMatrix(intRow, COL_ʱ��))
        If lng��Ŀ��� = 2 And mint����Ӧ�� = 2 And mblnEdit���� = True Then
            '��ת��������������
            If mbln����������ʾ And InStr(strData, "/") > 0 Then
                strData = Split(strData, "/")(1) & "/" & Split(strData, "/")(0)
            End If
        End If
        strδ�� = Trim(vsf.TextMatrix(intRow, Col_δ��˵��))
        If strData <> "" Then strδ�� = ""
        '�������ݸ��´���
        With mrsCurve
            .Filter = "��Ŀ���=" & lng��Ŀ��� & " and ʱ��='" & Format(vsf.TextMatrix(intRow, col_ԭʼʱ��), "YYYY-MM-DD HH:mm:ss") & "'"
            int���Ժϸ� = IIf(vsf.TextMatrix(intRow, COL_���Ժϸ�) = "��", 1, 0)
            If .RecordCount <> 0 Then
                int�޸�״̬ = Val(vsf.TextMatrix(intRow, COL_�޸�״̬))
                Select Case int�޸�״̬
                    Case 0 'δ������
                    Case 1
                        !״̬ = 1
                        !��ֵ = strData
                        !��λ = str��λ
                        !��ʾ = IIf(vsf.TextMatrix(intRow, COL_��ʾ) = "��", 1, 0)
                        !���Ժϸ� = IIf(vsf.TextMatrix(intRow, COL_���Ժϸ�) = "��", 1, 0)
                        !�޸� = 0
                        !������Դ = 0
                        !δ��˵�� = strδ��
                        
                    Case 2 '�޸�
                        If !״̬ = 1 Then
                            !״̬ = 1
                        Else
                            !״̬ = 2
                        End If
                        !��ֵ = strData
                        !��λ = str��λ
                        !��ʾ = IIf(vsf.TextMatrix(intRow, COL_��ʾ) = "��", 1, 0)
                        !���Ժϸ� = IIf(vsf.TextMatrix(intRow, COL_���Ժϸ�) = "��", 1, 0)
                        !δ��˵�� = strδ��
                        !������Դ = 0
                        !�޸� = 0
                    Case 3 'ɾ��
                        !��ֵ = ""
                        !δ��˵�� = strδ��
                        !״̬ = 3
                    Case 4 '������ɾ��
                        .Delete
                    Case 5 '�޸�ʱ��
                        !ʱ�� = Format(dtpDate.Value & " " & vsf.EditText, "YYYY-MM-DD hh:mm:ss")
                        Select Case !״̬
                            Case 0
                                !״̬ = 5
                            Case 1
                                !״̬ = 1
                            Case 2
                                !״̬ = 2
                        End Select
                        vsf.TextMatrix(intRow, col_ԭʼʱ��) = !ʱ��
                    Case 6 '�޸���ʾ
                       !��ʾ = IIf(vsf.TextMatrix(intRow, COL_��ʾ) = "��", 1, 0)
                       Select Case !״̬
                            Case 0
                                !״̬ = 6
                            Case 1
                                !״̬ = 1
                            Case 2
                                !״̬ = 2
                        End Select
                End Select
                .Update
            Else
                If (strData <> "" Or strδ�� <> "") Then
                    If strTime = "" Then
                        strTime = Format(GetCenterTime(mstrBegin, mstrEnd), "YYYY-MM-DD hh:mm:ss")
                    Else
                        strTime = Format(dtpDate.Value & " " & strTime, "YYYY-MM-DD HH:mm:ss")
                    End If
                    vsf.TextMatrix(intRow, col_ԭʼʱ��) = strTime
                    gstrFields = "���|������|��ֵ|��λ|���|ʱ��|��Ŀ���|��Ŀ����|���Ժϸ�|δ��˵��|������Դ|�޸�|��ʾ|��ԴID|����|״̬|�к�|��¼����"
                    gstrValues = GetMaxNum(mrsCurve) & "|1)����������Ŀ|" & strData & "|" & str��λ & "|" & _
                        "0" & "|" & strTime & "|" & lng��Ŀ��� & "|" & strName & "|" & _
                        "0" & "|" & strδ�� & "|0|0|0|0|0|1|0|1"
                    Call Record_Add(mrsCurve, gstrFields, gstrValues)
                End If
            End If
        End With
    ElseIf intType = 2 Then
        lng��Ŀ��� = Val(vsfCurve.TextMatrix(intRow, COL_��Ŀ���))
        If intCOl = COL_���� Then
            strData = vsfCurve.EditText
        Else
            strData = Trim(vsfCurve.TextMatrix(intRow, COL_����))
        End If
        strδ�� = Trim(vsfCurve.TextMatrix(intRow, Col_δ��˵��))
        strTime = Trim(vsfCurve.TextMatrix(intRow, COL_ʱ��))
        mrsNote.Filter = "��¼����=" & lng��Ŀ��� & " and ʱ��='" & Format(vsfCurve.TextMatrix(intRow, col_ԭʼʱ��), "YYYY-MM-DD HH:mm:ss") & "'"
        If mrsNote.RecordCount <> 0 Then
            int�޸�״̬ = Val(vsfCurve.TextMatrix(intRow, COL_�޸�״̬))
            Select Case int�޸�״̬
                Case 0 'δ������
                Case 1 '�������޸�
                    mrsNote!״̬ = 2
                    mrsNote!���� = strData
                    mrsNote!δ��˵�� = IIf(mrsNote!���� = "", "", vsfCurve.Cell(flexcpBackColor, intRow, COL_��ɫ, intRow, COL_��ɫ))
                Case 2 '�޸�
                    mrsNote!״̬ = 2
                    mrsNote!���� = strData
                    mrsNote!δ��˵�� = IIf(mrsNote!���� = "", "", vsfCurve.Cell(flexcpBackColor, intRow, COL_��ɫ, intRow, COL_��ɫ))
                Case 3 'ɾ��
                    mrsNote!���� = ""
                    mrsNote!δ��˵�� = ""
                    mrsNote!״̬ = 3
                Case 4 '������ɾ��
                    mrsNote!״̬ = 4
                Case 5
                    mrsNote!ʱ�� = Format(dtpDate.Value & " " & vsfCurve.EditText, "YYYY-MM-DD hh:mm:ss")
                    Select Case mrsNote!״̬
                        Case 0
                            mrsNote!״̬ = 5
                        Case 1
                            mrsNote!״̬ = 1
                        Case 2
                            mrsNote!״̬ = 2
                            
                    End Select
            End Select
            mrsNote.Update
        Else
            If lng��Ŀ��� = 2 Then
                    strName = "�ϱ�˵��"
                ElseIf lng��Ŀ��� = 6 Then
                    strName = "�±�˵��"
                End If
            If strData <> "" Or strδ�� <> "" Then
                If strTime = "" Then
                    strTime = Format(GetCenterTime(mstrBegin, mstrEnd), "YYYY-MM-DD hh:mm:ss")
                Else
                    strTime = Format(dtpDate.Value & " " & strTime, "YYYY-MM-DD HH:mm:ss")
                End If
                vsfCurve.TextMatrix(intRow, col_ԭʼʱ��) = strTime
                gstrFields = "���|��Ŀ���|ʱ��|ԭʼʱ��|��¼����|����|��Ŀ����|δ��˵��|��¼���|������Դ|��ʾ|��ԴID|����|״̬"
                gstrValues = GetMaxNum(mrsNote) & "|" & 0 & "|" & strTime & "|" & strTime & "|" & lng��Ŀ��� & "|" & strData & "|" & strName & "|" & IIf(strδ�� = "", vsfCurve.Cell(flexcpBackColor, intRow, COL_��ɫ, intRow, COL_��ɫ), "") & "|0|0|0|0|0|1"
                Call Record_Add(mrsNote, gstrFields, gstrValues)
            End If
        End If
        
    End If
      
    UpdateCurveDate = True
    
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetMaxNum(ByVal rsTmp As ADODB.Recordset) As Long
'----------------------------------------------------
'����:��ȡ��¼mrsCurve�е�������
'----------------------------------------------------
    On Error GoTo Errhand
    rsTmp.Filter = 0
    rsTmp.Sort = "��� Desc"
    If rsTmp.RecordCount = 0 Then
        GetMaxNum = 1
    Else
        GetMaxNum = Val(rsTmp!���) + 1
    End If
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetMaxID(ByVal rsTmp As ADODB.Recordset) As Long
'----------------------------------------------------
'����:��ȡ��¼mrsCurve�е�������
'----------------------------------------------------

    On Error GoTo Errhand
    rsTmp.Filter = 0
    rsTmp.Sort = "id Desc"
    If rsTmp.RecordCount = 0 Then
        GetMaxID = 1
    Else
        GetMaxID = Val(rsTmp!Id) + 1
    End If
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub InitTabControl()
'--------------------------------------------------------------------------------
'����:��ʼ��TabControl
'--------------------------------------------------------------------------------
    Dim tabItem As TabControlItem
    Dim CtlFont As stdFont
    
    On Error GoTo Errhand
    With tbcThis
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .BoldSelected = True
            .ShowIcons = True
            .OneNoteColors = True
            .Position = xtpTabPositionTop
            .ClientFrame = xtpTabFrameSingleLine
            .DisableLunaColors = False
            .Layout = xtpTabLayoutAutoSize
            Set CtlFont = .Font
            If CtlFont Is Nothing Then
                Set CtlFont = Me.Font
            End If
            CtlFont.Size = IIf(mintBigSize = 0, 9, 12)
            Set .Font = CtlFont
        End With
        
        Set tabItem = .InsertItem(1, "��������", picCurve.hWnd, 0)
        tabItem.Tag = "����"
        Set tabItem = .InsertItem(2, "���±��", picTab.hWnd, 0)
        tabItem.Tag = "���"
        
        
        If gintEditorCurveState = 0 Then
            .Item(0).Selected = True
        Else
            .Item(1).Selected = True
        End If
    End With
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngRow As Long, lngCol As Long, lngItemNO As Long, intType As Integer, i As Integer
    Dim strValue As String, strPart As String, strPart1 As String, strName As String
    Dim strTime As String, strErrMsg As String
    Dim strTmp As String, arrStr() As String, arrTime() As String
    Dim cbrCheck As CommandBarControl
    Dim rsTemp As New ADODB.Recordset

    Select Case Control.Id
        Case conMenu_Edit_Save '����
            If Not SaveData Then Exit Sub
            Call GetTableRowName
            Call zlRefreshData
            Call SetColSelect
        Case conMenu_Edit_Reuse 'ȡ��
            Call txtEdit_KeyPress(vbKeyEscape)
            Call GetTableRowName
            Call zlRefreshData
            Call SetColSelect
        Case conMenu_Edit_NewItem '��ӻ��Ŀ
            Call txtEdit_KeyPress(vbKeyEscape)
            mblnScroll = True
            If frmCaseTendBodyActiveItem.ShowMe(vsfTab, Me, mT_Patient.lng����ȼ�, mT_Patient.lngӤ��, mT_Patient.lng����ID) Then
                vsfTab.Refresh
            End If
        Case conMenu_Edit_Append * 10, conMenu_Edit_Append * 10 + 1, conMenu_Edit_Append * 10 + 2, conMenu_Edit_Append * 10 + 30, conMenu_Edit_Append * 10 + 31, conMenu_Edit_Append * 10 + 4, conMenu_Edit_Append * 10 + 5, conMenu_Edit_Append * 10 + 6
            If vsfTab.Tag <> "" Then
                If vsfTab.Row < vsfTab.FixedRows Or vsfTab.Col < vsfTab.FixedCols Then Exit Sub
                lngRow = vsfTab.Row
                lngCol = vsfTab.Col
                lngItemNO = Val(vsfTab.TextMatrix(lngRow, COL_tab��Ŀ���))
                strName = vsfTab.TextMatrix(lngRow, COL_tab��Ŀ��)
                strValue = Trim(vsfTab.TextMatrix(lngRow, lngCol))
                strTmp = vsfTab.TextMatrix(lngRow, COL_tab�ַ���)
            Else
                If vsfTabDetail.Row < vsfTabDetail.FixedRows Or vsfTabDetail.Col < vsfTabDetail.FixedCols Or vsfTabDetail.Row > Val(vsfTabDetail.Tag) Then Exit Sub
                lngRow = vsfTabDetail.Row
                lngCol = vsfTabDetail.Col
                lngItemNO = Val(vsfTabDetail.TextMatrix(lngRow, COL_tab��Ŀ���))
                strName = vsfTabDetail.TextMatrix(lngRow, COL_tab��Ŀ��)
                strValue = Trim(vsfTabDetail.TextMatrix(lngRow, lngCol))
                strTmp = vsfTabDetail.TextMatrix(lngRow, COL_tab�ַ���)
            End If
            strTmp = strTmp & String(8 - UBound(Split(strTmp, ",")), ",")
            arrStr = Split(strTmp, ",")
            
            intType = 0
            If picEdit.Visible = True And txtEdit.Visible = True Then intType = 1
            If intType = 1 Then strValue = txtEdit.Text
            strPart = ""
            If InStr(1, "," & gint��� & "," & gint��Һ & ",", "," & lngItemNO & ",") = 0 Then Exit Sub
            Select Case Control.Id
                Case conMenu_Edit_Append * 10 + 1
                    strPart = "E"
                    If InStr(1, UCase(strValue), "/E") > 0 Then
                        strValue = Mid(strValue, 1, InStr(1, UCase(strValue), "/E") - 1)
                    End If
                    If InStr(1, UCase(strValue), "E") > 0 Then
                        strValue = Mid(strValue, 1, InStr(1, UCase(strValue), "E") - 1)
                    End If
                Case conMenu_Edit_Append * 10 + 2
                    strPart = "/E"
                    If InStr(1, UCase(strValue), "/E") > 0 Then
                        strValue = Mid(strValue, 1, InStr(1, UCase(strValue), "/E") - 1)
                    End If
                Case conMenu_Edit_Append * 10 + 30
                    strPart = "��"
                    strValue = ""
                Case conMenu_Edit_Append * 10 + 31
                    strPart = "*"
                    strValue = ""
                Case conMenu_Edit_Append * 10 + 4
                    strPart = "��"
                    strValue = ""
                Case conMenu_Edit_Append * 10 + 5
                    strPart = "C"
                    If InStr(1, UCase(strValue), "/C") > 0 Then
                        strValue = Mid(strValue, 1, InStr(1, UCase(strValue), "/C") - 1)
                    End If
                    If InStr(1, UCase(strValue), "C") > 0 Then
                        strValue = Mid(strValue, 1, InStr(1, UCase(strValue), "C") - 1)
                    End If
                Case conMenu_Edit_Append * 10 + 6
                    strPart = "/C"
                    If InStr(1, UCase(strValue), "/C") > 0 Then
                        strValue = Mid(strValue, 1, InStr(1, UCase(strValue), "/C") - 1)
                    End If
                Case conMenu_Edit_Append * 10
                    strPart = ""
                    If lngItemNO = gint��� Then
                        For i = 0 To 4
                            Select Case i
                                Case 0
                                    strPart1 = "E"
                                Case 1
                                    strPart1 = "/"
                                Case 2
                                    strPart1 = "*"
                                Case 3
                                    strPart1 = "��"
                                Case 4
                                    strPart1 = "��"
                            End Select
                            strValue = Replace(UCase(strValue), strPart1, "")
                        Next i
                    Else
                        strValue = Replace(UCase(Replace(UCase(strValue), "C", "")), "/", "")
                    End If
            End Select
            
            If IsNumeric(strValue) Then
                strValue = strValue
            Else
                strValue = ""
            End If
            strValue = strValue & Trim(strPart)
            If Left(strValue, 1) = "/" Then strValue = 1 & strValue
            
            If intType = 1 Then
                txtEdit.Text = strValue
                For Each cbrCheck In mcbrToolBar.Controls(5).Controls
                    If cbrCheck.Id = Control.Id Then
                        cbrCheck.Checked = True
                    Else
                        cbrCheck.Checked = False
                    End If
                Next

                Exit Sub
            End If
             
            '�Ǳ༭״̬��
            If IsWaveItem(lngItemNO) And InStr(1, Trim(vsfTab.TextMatrix(lngRow, lngCol)), "-") <> 0 And vsfTab.Tag <> "" Then
                strErrMsg = "������ֵ�Ѿ��γɲ�����Χ�Ĳ�����Ŀ���ܽ����޸ġ�ɾ������"
                lblStb.Caption = strErrMsg: lblStb.ForeColor = 255
                Exit Sub
            End If
            Call txtEdit_KeyPress(vbKeyEscape)
            strPart = CStr(arrStr(7))
            If vsfTab.Tag <> "" Then
                strTime = Format(Split(vsfTab.TextMatrix(vsfTab.Row, col_tabԭʼʱ��), "'")(vsfTab.Col - vsfTab.FixedCols), "YYYY-MM-DD hh:mm:ss")
            Else
                strTime = Format(vsfTabDetail.TextMatrix(vsfTabDetail.Row, col_tabԭʼʱ��), "YYYY-MM-DD hh:mm:ss")
            End If
            If strTime = "" Then
                strTime = GetAnimalItemTime(vsfTab.Row, vsfTab.Col - vsfTab.FixedCols + 1, 0, strErrMsg)
                If strErrMsg <> "" Then lblStb.Caption = strErrMsg: lblStb.ForeColor = 255: Exit Sub
                strTime = Format(DateAdd("n", DateDiff("n", Split(strTime, ";")(0), Split(strTime, ";")(1)) / 2, Split(strTime, ";")(0)), "YYYY-MM-DD hh:mm:ss")
                If IsExistData(strTime, lngItemNO) = False Then
                    Exit Sub
                End If
            End If
            
            mrsTableDetail.Filter = "��Ŀ���=" & lngItemNO & " and ��Ŀ����='" & strName & "' and ʱ��='" & strTime & "'"
            If mrsTableDetail.RecordCount > 0 Then
                If mrsTableDetail!״̬ <> 1 Then  'ԭ�е����� �޸ġ�ɾ�����״̬ʼ��Ϊ2
                    mrsTableDetail!״̬ = 2
                    mrsTableDetail!��� = strValue
                Else '�����������ݵĴ���
                    If Trim(vsfTab.TextMatrix(lngRow, lngCol)) = "" Then
                        mrsTableDetail.Delete
                    Else
                        mrsTableDetail!״̬ = 1
                        mrsTableDetail!��� = strValue
                    End If
                End If
                mrsTableDetail.Update
            Else '�����ڼ�¼����������
                If Trim(strValue) <> "" Then
                    
                    gstrFields = "id|������|���|���²�λ|���|ʱ��|��Ŀ���|��Ŀ����|���Ժϸ�|δ��˵��|������Դ|�޸�|��ʾ|��ԴID|����|״̬|�к�|��¼����"
                    gstrValues = GetMaxID(mrsTableDetail) & "|2)���±����Ŀ|" & strValue & "|" & strPart & "|" & _
                        0 & "|" & strTime & "|" & lngItemNO & "|" & strName & "|0||0|0|0|0|0|1|" & lngCol - vsfTab.FixedCols + 1 & "|1"
                    Call Record_Add(mrsTableDetail, gstrFields, gstrValues)
                    If vsfTab.Tag <> "" Then
                        arrTime = Split(vsfTab.TextMatrix(vsfTab.Row, col_tabԭʼʱ��), "'")
                        arrTime(vsfTab.Col - vsfTab.FixedCols) = strTime
                        vsfTab.TextMatrix(vsfTab.Row, col_tabԭʼʱ��) = Join(arrTime, "'")
                    Else
                        vsfTabDetail.TextMatrix(vsfTabDetail.Row, col_tabԭʼʱ��) = strTime
                    End If
                End If
            End If
            
            mrsTableDetail.Filter = "״̬<> 4 "
    
            gstrFields = "ID," & adDouble & ",18|������," & adLongVarChar & ",40|���," & adLongVarChar & ",400|���²�λ," & adLongVarChar & ",200|" & _
                 "���," & adDouble & ",1|ʱ��," & adLongVarChar & ",20|ԭʼʱ��," & adLongVarChar & ",20|��Ŀ���," & adDouble & ",18|��Ŀ����," & adLongVarChar & ",40|" & _
                 "���Ժϸ�," & adDouble & ",1|δ��˵��," & adLongVarChar & ",20|������Դ," & adDouble & ",1|�޸�," & adDouble & ",1|��ʾ," & adDouble & ",1|ԭʼ��ʾ״̬," & adDouble & ",1|" & _
                 "��ԴID," & adDouble & ",18|����," & adDouble & ",1|״̬," & adDouble & ",1|�к�," & adDouble & ",1|��¼����," & adDouble & ",1"
            Call Record_Init(rsTemp, gstrFields)
            
            Do While Not mrsTableDetail.EOF
                rsTemp.AddNew
                For i = 0 To mrsTableDetail.Fields.Count - 1
                    rsTemp.Fields(mrsTableDetail.Fields(i).Name).Value = mrsTableDetail.Fields(i).Value
                Next i
                rsTemp.Update
                mrsTableDetail.MoveNext
            Loop
            
            Call InitTableData(rsTemp)
            Call ShowTable
            If vsfTab.Tag <> "" Then
                 vsfTab.TextMatrix(lngRow, lngCol) = strValue
                Call vsfTab_AfterRowColChange(0, 0, lngRow, lngCol)
            Else
                vsfTabDetail.TextMatrix(lngRow, lngCol) = strValue
                Call vsfTabDetail_AfterRowColChange(0, 0, lngRow, lngCol)
            End If
            
        Case conMenu_Help_Help '����
            Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_File_Exit '�˳�
            Unload Me
'        Case conMenu_View_Forward
'            dtpDate.Value = dtpDate.Value - 1
'            Call dtpDate_Change
'        Case conMenu_View_Backward
'            dtpDate.Value = dtpDate.Value + 1
'            Call dtpDate_Change
    End Select
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)

    On Error Resume Next
    If stbThis.Visible = True Then Bottom = stbThis.Height
    
    With picStb
        .Top = stbThis.Top + 50
        .Left = stbThis.Panels(2).Left + 50
        .Height = stbThis.Height - 50
        .Width = stbThis.Panels(2).Width - 50
    End With

    With lblStb
        .Font.Size = 9 + 9 * mintBigSize / 3
        .Height = TextHeight("����")
        .Top = (picStb.Height - .Height) \ 2
        .Left = 10
    End With
End Sub



Private Function InitRecordSet() As Boolean
'----------------------------------------------------------------
'����:��ʼ����¼�� ������λ��Ϣ��������Ŀʱ�Σ���¼Ƶ��ʱ��
'----------------------------------------------------------------
    On Error GoTo Errhand
    '��ȡ���в�λ��Ϣ
    mstrSQL = "Select ��Ŀ���,��λ,ȱʡ�� From ���²�λ"
    Call zlDatabase.OpenRecordset(mrsPart, mstrSQL, Me.Caption)
    
    '��ȡ���ü�¼����Ϣ
    Call InitPublicData
    
    InitRecordSet = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Sub cbsMain_Resize()
    On Error Resume Next
    
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    Dim lngBottom As Long  '�ͻ�����Ĵ�С

    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    Call picToolBarReSize
    picSplit.Left = lngLeft
    picSplitTab.Left = lngLeft
    
    With tbcThis
        .Top = lngTop
        .Left = lngLeft
        .Width = lngRight - lngLeft
        .Height = lngBottom - lngTop
    End With
    picNull.Move lngLeft, lngTop + 350, lngRight, lngBottom - 350
    With lblInfo
        .Top = (picNull.Height - .Height) / 2
        .Left = (picNull.Width - .Width) / 2
    End With
    picNull.Visible = False
End Sub

Private Function GetCenterTime(ByVal dBegin As Date, ByVal dEnd As Date, Optional intBound As Integer = 0) As String
'------------------------------------------------------------------------------------
'����:��ȡĳ��ʱ����е�ʱ��,�����ǰʱ���ڱ��η�Χ�������м�ʱ�������Ե�ǰʱ��Ϊ׼
'------------------------------------------------------------------------------------
    Dim dblvalue As Double
    Dim strTime As String, strCurDate As String
    Dim i As Integer
    
    On Error GoTo Errhand
    strCurDate = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:mm:ss")
    dblvalue = DateDiff("s", dBegin, dEnd)
    strTime = Format(DateAdd("s", Fix(dblvalue / 2), dBegin), "YYYY-MM-DD HH:mm:ss")
    If strTime < mstrBTime Then
        strTime = mstrBTime
    End If
    If strTime > mstrETime Then
        strTime = mstrETime
    End If
    
    For i = 0 To UBound(marrTime)
        If Format(strTime, "HH:mm:ss") >= Format(Split(marrTime(i), ",")(0), "HH:mm:ss") And Format(strTime, "HH:mm:ss") <= Format(Split(marrTime(i), ",")(1), "HH:mm:ss") Then
            Exit For
        End If
    Next i
    If i <= UBound(marrTime) Then
        If gintHourBegin + i * T_BodyStyle.lngʱ���� = 24 Then
            strTime = Format(Format(dBegin, "YYYY-MM-DD") & " " & "23:59:59", "YYYY-MM-DD HH:mm:ss")
        Else
            strTime = Format(Format(dBegin, "YYYY-MM-DD") & " " & gintHourBegin + i * T_BodyStyle.lngʱ���� & ":00:00", "YYYY-MM-DD HH:mm:ss")
        End If
    End If
    intBound = i
    
    If CDate(strCurDate) >= dBegin And CDate(strCurDate) <= dEnd And CDate(strCurDate) < CDate(strTime) Then
        strTime = strCurDate
    End If
    
    If CDate(strTime) < CDate(mstrBTime) Then
        strTime = Format(mstrBTime, "YYYY-MM-DD HH:mm:ss")
    End If
    If CDate(strTime) > CDate(mstrETime) Then
        strTime = Format(mstrETime, "YYYY-MM-DD HH:mm:ss")
    End If
    GetCenterTime = strTime
    
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim frmMain As Form
    Dim blnEnable As Boolean
    Dim strCurrentdate As String
    
    strCurrentdate = zlDatabase.Currentdate
    Select Case Control.Id
        Case conMenu_Edit_Save, conMenu_Edit_Reuse
            picNull.Visible = (vsfTab.TextMatrix(1, COL_tab��Ŀ���) = -999 And tbcThis.Selected.Tag = "���")
            Control.Enabled = IIf(IsChange = True, True, False)
        Case conMenu_Edit_NewItem
            If tbcThis.Selected.Tag = "���" Then
                Control.Visible = True
                Control.Enabled = Not mblnFileBack
            Else
                Control.Visible = False
            End If
        Case conMenu_Edit_Append
            If tbcThis.Selected.Tag = "���" Then
                Control.Visible = True
            Else
                Control.Visible = False
            End If
        Case conMenu_Edit_Append * 10 + 0, conMenu_Edit_Append
            Control.Enabled = (is������Һ(1) Or is������Һ(2)) And Not mblnFileBack And tbcThis.Selected.Tag = "���"
        Case conMenu_Edit_Append * 10 + 1, conMenu_Edit_Append * 10 + 2, conMenu_Edit_Append * 10 + 3, conMenu_Edit_Append * 10 + 4
            Control.Enabled = is������Һ(1) And Not mblnFileBack And tbcThis.Selected.Tag = "���"
        Case conMenu_Edit_Append * 10 + 5, conMenu_Edit_Append * 10 + 6
            Control.Enabled = is������Һ(2) And Not mblnFileBack And tbcThis.Selected.Tag = "���"
    End Select
    
    If dtpDate.Value = dtpDate.MinDate Then
        imgbtn(1).Picture = ilsDate.ListImages("preGray").Picture
        imgbtn(1).Enabled = False
    Else
'        imgbtn(1).Picture = ilsDate.ListImages("preGreen").Picture
    End If
    If dtpDate.Value = dtpDate.MaxDate Then
        imgbtn(0).Picture = ilsDate.ListImages("nextGray").Picture
        imgbtn(0).Enabled = False
    Else
'        imgbtn(0).Picture = ilsDate.ListImages("nextGreen").Picture
    End If
    If Format(mstrETime, "yyyy-mm-dd") = Format(strCurrentdate, "yyyy-mm-dd") Then mstrETime = strCurrentdate
End Sub


Private Function is������Һ(ByVal intType As Integer) As Boolean
    '--------------------------------------------------
    '����Ƿ��Ǵ����Ŀ����ҹ��Ŀ  �����Ŀ���=10 ��ҹ=9
    'intType=1 Ϊ�����Ŀ ����Ϊ��Һ��Ŀ
    '--------------------------------------------------
    Dim lngItemNO As Long
    Dim strKey As String
    Dim rsObj As New ADODB.Recordset
    Dim strTmp As String, strName As String, arrStr() As String
    On Error GoTo Errhand
    
    If vsfTab.Col < vsfTab.FixedCols Or vsfTab.Row < vsfTab.FixedRows Then Exit Function
    If mblnInit = False Then Exit Function
    
    '��ȡ��Ŀ���
    lngItemNO = Val(vsfTab.TextMatrix(vsfTab.Row, COL_tab��Ŀ���))
    If intType = 1 Then
        If lngItemNO <> 10 Then Exit Function
    Else
        If lngItemNO <> 9 Then Exit Function
    End If
    strName = vsfTab.TextMatrix(vsfTab.Row, COL_tab��Ŀ��)
    strTmp = vsfTab.TextMatrix(vsfTab.Row, COL_tab�ַ���)
    strTmp = strTmp & String(8 - UBound(Split(strTmp, ",")), ",")
    arrStr = Split(strTmp, ",")
    
    '����¼Ƶ�κ���Ŀ��ʾ
    If vsfTab.Col > vsfTab.FixedCols + Val(arrStr(3)) - 1 Then Exit Function
    If InStr(1, ",2,3,5,", "," & Val(arrStr(4)) & ",") > 0 Then Exit Function
    
    '����Ƿ���ͬ��������
    If vsfTab.Tag <> "" Then
        mrsTableDetail.Filter = "��Ŀ���=" & lngItemNO & " and ��Ŀ����='" & strName & "'" & _
            "   and �к�=" & vsfTab.Col - vsfTab.FixedCols + 1
    Else
        mrsTableDetail.Filter = "��Ŀ���=" & lngItemNO & " and ��Ŀ����='" & strName & "'" & _
            "   and ʱ�� ='" & vsfTabDetail.TextMatrix(vsfTabDetail.Row, col_tabԭʼʱ��) & "'"
    End If
    If mrsTableDetail.RecordCount > 0 Then
        If InStr(1, ",0,3,9,", "," & Val(mrsTableDetail!������Դ) & ",") = 0 Then
            Exit Function
        End If
    End If
    
    is������Һ = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Sub cmdColor_Click()
    Call txtEdit_KeyDown(vbKeyDown, vbShiftMask)
End Sub

Private Sub dtpDate_Change()
    Dim strDate As String
    Dim cbrControl As CommandBarControl
     If IsChange Then
        If MsgBox("�������������Ѿ������ı�,�����Ƿ���Ҫ���棿", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Exit Sub
        End If
    End If
    If Not dtpDateChageDate(Format(dtpDate.Value, "YYYY-MM-DD")) Then Exit Sub
    imgbtn(1).Enabled = True
    imgbtn(0).Enabled = True
'    If dtpDate.Enabled = True Then dtpDate.SetFocus
'    Set cbrControl = mcbrToolBar.Controls.Find(, conMenu_View_Forward)
'    cbrControl.Enabled = True
'    Set cbrControl = mcbrToolBar.Controls.Find(, conMenu_View_Backward)
'    cbrControl.Enabled = True
    
'    If dtpDate.Value = dtpDate.MinDate Then
'        Set cbrControl = mcbrToolBar.Controls.Find(, conMenu_View_Forward)
'        cbrControl.Enabled = False
'    End If
'    If dtpDate.Value = dtpDate.MaxDate Then
'        Set cbrControl = mcbrToolBar.Controls.Find(, conMenu_View_Backward)
'        cbrControl.Enabled = False
'    End If
    If dtpDate.Value = dtpDate.MinDate Then
        imgbtn(1).Picture = ilsDate.ListImages("preGray").Picture
        imgbtn(1).Enabled = False
    Else
        imgbtn(1).Picture = ilsDate.ListImages("preGreen").Picture
    End If
    If dtpDate.Value = dtpDate.MaxDate Then
        imgbtn(0).Picture = ilsDate.ListImages("nextGray").Picture
        imgbtn(0).Enabled = False
    Else
        imgbtn(0).Picture = ilsDate.ListImages("nextGreen").Picture
    End If
End Sub

Private Sub dtpDate_CloseUp()
    Call SetColSelect
End Sub

Private Sub dtpDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call tbcThis_SelectedChanged(tbcThis.Selected)
    End If
End Sub

Private Sub dtpDate_Validate(Cancel As Boolean)
    If Not dtpDateChageDate(Format(dtpDate.Value, "YYYY-MM-DD")) Then
        Cancel = True
    End If
End Sub

Private Sub Form_Activate()
    If Not mblnStart Then Exit Sub
    mblnStart = False
    Call SetColSelect(True)
End Sub


Private Sub Form_Unload(Cancel As Integer)

    If IsChange Then
        If MsgBox("�������������Ѿ������ı�,�����Ƿ���Ҫ���棿", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Cancel = True
            Exit Sub
        End If
    End If
    mblnMove = False
    mblnInit = False
    mblnEdit = False
    mbln��Ժ = False
    mblnAllRefresh = False
    If Not (mrsCurve Is Nothing) Then Set mrsCurve = Nothing
    If Not (mrsPart Is Nothing) Then Set mrsTable = Nothing
    If Not (mrsTable Is Nothing) Then Set mrsTableDetail = Nothing
    If Not (mrsTableDetail Is Nothing) Then Set mrsPart = Nothing
    If Not (mrsNote Is Nothing) Then Set mrsNote = Nothing
    If Not (mrsRecodeID Is Nothing) Then Set mrsRecodeID = Nothing
    If Not (mcbrToolBar Is Nothing) Then Set mcbrToolBar = Nothing
    Call UnLoadOptTime
    '���洰��
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Function IsChange()
'�����Ƿ�ı�
    Dim blnChange As Boolean
    
    If mrsCurve.State = adStateOpen Then
        mrsCurve.Filter = "״̬ <> 0 "
        If mrsCurve.RecordCount > 0 Then blnChange = True
    End If
    
    If mrsNote.State = adStateOpen Then
        mrsNote.Filter = "״̬ <> 0 "
        If mrsNote.RecordCount > 0 Then blnChange = True
    End If
    
    If mrsTableDetail.State = adStateOpen Then
        mrsTableDetail.Filter = "״̬ <> 0 "
        If mrsTableDetail.RecordCount > 0 Then blnChange = True
    End If
    IsChange = blnChange
End Function


Private Sub imgbtn_Click(Index As Integer)
    Select Case Index
        Case 1
            dtpDate.Value = dtpDate.Value - 1
            Call dtpDate_Change
        Case 0
            dtpDate.Value = dtpDate.Value + 1
            Call dtpDate_Change
    End Select
End Sub


Private Sub lblCheck_DblClick()
    Call picEdit_KeyPress(vbKeySpace)
End Sub

Private Sub lstSelect_DblClick(Index As Integer)
    Call lstSelect_KeyDown(Index, vbKeyReturn, 0)
End Sub

Private Sub lstSelect_GotFocus(Index As Integer)
    Dim i As Integer, j As Integer
    PicLst.Tag = 0
    j = lstSelect(Index).ListCount - 1
    If Index = 0 And j >= 0 Then
        If lstSelect(Index).ListIndex < 0 Then lstSelect(Index).ListIndex = 0
    End If
End Sub

Private Sub lstSelect_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim strData As String
    Dim blnAllow As Boolean
    Dim i As Integer
    Dim blnTab As Boolean '�Ƿ�������
    Dim arrTag() As String
    
    strData = ""
    blnAllow = True
    If InStr(1, lbllst(Index).Tag, "|") > 0 Then arrTag = Split(lbllst(Index).Tag, "|")
    If UBound(arrTag) > 1 Then
        If Split(lbllst(Index).Tag, "|")(2) = 1 Then blnTab = True
    End If
    If KeyCode = vbKeyReturn Then
        If Shift = vbShiftMask Then Exit Sub
        For i = 0 To lstSelect(Index).ListCount - 1
            If lstSelect(Index).Selected(i) = True Then
                strData = strData & "," & Replace(lstSelect(Index).List(i), ",", "")
            End If
        Next
        If Left(strData, 1) = "," Then strData = Mid(strData, 2)
        If strData <> lstSelect(Index).Tag Then
            If blnTab Then
                blnAllow = WriteIntoVfgTab(strData, vsfTab)
            Else
                blnAllow = WriteIntoVfgTab(strData, vsfTabDetail)
            End If
        End If
        If blnAllow = True Then Call vsfTab_KeyDown(vbKeyReturn, Shift)
    ElseIf KeyCode = vbKeyLeft Then
        If blnTab Then
            Call vsfTab_KeyDown(vbKeyLeft, 0)
        Else
            Call vsfTabDetail_KeyDown(vbKeyLeft, 0)
        End If
    ElseIf KeyCode = vbKeyEscape Then
        Call txtEdit_KeyPress(vbKeyEscape)
    ElseIf Index = 0 And Shift = vbShiftMask And KeyCode = vbKeyUp Then
        KeyCode = 0
        txtLst.SetFocus
    End If
End Sub

'Private Sub imgbtn_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Select Case Index
'        Case 0
'            imgbtn(0).Picture = IIf(imgbtn(0).Enabled = True, ilsDate.ListImages("nextLight").Picture, ilsDate.ListImages("nextGray").Picture)
'        Case 1
'            imgbtn(1).Picture = IIf(imgbtn(1).Enabled = True, ilsDate.ListImages("preLight").Picture, ilsDate.ListImages("preGray").Picture)
'    End Select
'    imgDefault.Tag = Index + 1
'End Sub


Private Sub lstδ��_DblClick()
    Dim strδ�� As String
    Dim intRow As Integer
    Dim intCOl As Integer
    Dim intCount As Integer
    Dim intRows As Integer
    Dim blnAllow As Boolean
    Dim strTime As String
    Dim lng��Ŀ��� As Long
    
    If lstδ��.Tag <> "" Then
        strδ�� = lstδ��.Text
        intRow = Split(lstδ��.Tag, "|")(1)
        intCOl = Split(lstδ��.Tag, "|")(2)
        Select Case Split(lstδ��.Tag, "|")(0)
            Case 1
                vsfCurve.Row = intRow
                vsfCurve.Col = intCOl
                vsfCurve.TextMatrix(intRow, intCOl) = strδ��
                strTime = vsfCurve.TextMatrix(intRow, COL_ʱ��)
                vsfCurve.TextMatrix(intRow, COL_����) = Space(vsfCurve.Row) & Space(vsfCurve.Row)
                vsfCurve.TextMatrix(intRow, COL_��ɫ) = Space(vsfCurve.Row) & IIf(vsfCurve.TextMatrix(intRow, COL_������) = "2)���±�˵��", " ", Space(vsfCurve.Row))
                vsfCurve.TextMatrix(intRow, COL_��λ) = ""
                vsfCurve.TextMatrix(intRow, COL_���Ժϸ�) = ""
            Case 2
                strTime = IIf(vsfDetail.TextMatrix(intRow, COL_ʱ��) = "", GetCenterTime(mstrBegin, mstrEnd), vsfCurve.TextMatrix(intRow, COL_ʱ��))
                lng��Ŀ��� = vsfCurve.TextMatrix(vsfCurve.Row, COL_��Ŀ���)
                mrsCurve.Filter = "��Ŀ���=" & lng��Ŀ��� & " and  ʱ��='" & Format(strTime, "YYYY-MM-DD hh:mm:ss") & "'"
                If mrsCurve.RecordCount > 0 Then
                    lblStb.Caption = "��ǰĬ��ʱ���Ѵ������ݣ���������ʱ��"
                    lblStb.ForeColor = 255
                    picδ��.Visible = False
                    Exit Sub
                End If
                vsfDetail.Row = intRow
                vsfDetail.Col = intCOl
                vsfDetail.TextMatrix(intRow, intCOl) = strδ��
                vsfDetail.TextMatrix(vsfCurve.Row, COL_����) = ""
                vsfDetail.TextMatrix(vsfCurve.Row, COL_��λ) = ""
                vsfDetail.TextMatrix(vsfCurve.Row, COL_���Ժϸ�) = ""
        End Select
        picδ��.Visible = False
        lstδ��.Visible = False: lstδ��.Enabled = False
    End If
    
    blnAllow = True
    intCount = 0
    intRows = 0
    If Split(lstδ��.Tag, "|")(0) = 2 Then
        Call UpdateCurveDate(vsfDetail, vsfDetail.Row, vsfDetail.Col, 3)
        Call vsfDetail.SetFocus
    Else
        If Trim(vsfCurve.TextMatrix(vsfCurve.Row, COL_������)) = "1)����������Ŀ" Then
            '����������ߵ�δ������Ϊ��,ֱ�Ӹ���
            For intRow = vsfCurve.FixedRows To vsfCurve.Rows - 1
                If Trim(vsfCurve.TextMatrix(intRow, COL_������)) = "1)����������Ŀ" Then
                    If vsfCurve.TextMatrix(intRow, Col_δ��˵��) = "" And Trim(vsfCurve.TextMatrix(intRow, COL_����)) = "" Then
                        intCount = intCount + 1
                    End If
                    intRows = intRows + 1
                End If
            Next
            'ʣ�µ���Ŀ���������Ƕ�Ϊ�������
            If intCount = intRows - 1 Then
                For intRow = vsfCurve.FixedRows To vsfCurve.Rows - 1
                    If Trim(vsfCurve.TextMatrix(intRow, COL_������)) = "1)����������Ŀ" And vsfCurve.TextMatrix(intRow, Col_δ��˵��) = "" Then
                        vsfCurve.TextMatrix(intRow, Col_δ��˵��) = strδ��
                        vsfCurve.TextMatrix(intRow, COL_ʱ��) = strTime
                        vsfCurve.TextMatrix(vsfCurve.Row, COL_����) = Space(vsfCurve.Row) & Space(vsfCurve.Row)
                        vsfCurve.TextMatrix(vsfCurve.Row, COL_��ɫ) = Space(vsfCurve.Row) & IIf(vsfCurve.TextMatrix(vsfCurve.Row, COL_������) = "2)���±�˵��", " ", Space(vsfCurve.Row))
                        vsfCurve.TextMatrix(vsfCurve.Row, COL_��λ) = ""
                        vsfCurve.TextMatrix(vsfCurve.Row, COL_���Ժϸ�) = ""
                    End If
                Next
            Else
                intCount = 0
            End If
        ElseIf Trim(vsfCurve.TextMatrix(vsfCurve.Row, COL_������)) = "2)���±�˵��" Then
            blnAllow = False
        End If
        vsfCurve.Cell(flexcpAlignment, vsfCurve.FixedRows, Col_δ��˵��, vsfCurve.Rows - 1, Col_δ��˵��) = flexAlignCenterCenter
        
        If blnAllow = True Then
            If intCount = 0 Then
                Call UpdateCurveDate(vsfCurve, vsfCurve.Row, vsfCurve.Col, 1)
            ElseIf intCount = intRows - 1 Then
                For intRow = vsfCurve.FixedRows To vsfCurve.Rows - 1
                    If Trim(vsfCurve.TextMatrix(intRow, COL_������)) = "1)����������Ŀ" Then
                        Call UpdateCurveDate(vsfCurve, intRow, Col_δ��˵��, 1)
                    End If
                Next
            End If
            Call vsfCurve.SetFocus
        End If
    End If
    
End Sub

Private Sub lstδ��_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyEscape Then
        lstδ��.Visible = False: lstδ��.Enabled = False
        picδ��.Visible = False
    ElseIf KeyCode = vbKeyReturn Then
        Call lstδ��_DblClick
    End If
End Sub

Private Sub lstδ��_LostFocus()
    lstδ��.Visible = False: lstδ��.Enabled = False
    picδ��.Visible = False
End Sub

Private Sub OptTime_Click(Index As Integer)
    Dim strBegin As String, strEnd As String
    Dim blnTab As Boolean
    
    If Not mblnInit Then Exit Sub
    
    If OptTime(Index).Tag = "" Then Exit Sub
    strBegin = Split(OptTime(Index).Tag, ",")(0)
    strEnd = Split(OptTime(Index).Tag, ",")(1)
    strBegin = Format(Format(dtpDate.Value, " YYYY-MM-DD") & " " & strBegin, "YYYY-MM-DD HH:mm:ss")
    strEnd = Format(Format(dtpDate.Value, " YYYY-MM-DD") & " " & strEnd, "YYYY-MM-DD HH:mm:ss")
    
    If CDate(strBegin) < CDate(mstrBTime) Then
        strBegin = Format(mstrBTime, "YYYY-MM-DD HH:mm:ss")
    End If
    
    If CDate(strEnd) > CDate(mstrETime) Then
        strEnd = Format(mstrETime, "YYYY-MM-DD HH:mm:ss")
    End If
    
    mstrBegin = strBegin
    mstrEnd = strEnd
    lblTime.Caption = Format(mstrBegin, "HH:mm") & "��" & Format(mstrEnd, "HH:mm")
    Call ShowCurve
    Call ShowTabUpDown
    
    If mblnStart = False Then
        Call SetColSelect(True)
    End If
End Sub

Private Sub SetColSelect(Optional blnInit As Boolean = False, Optional intType As Integer = 1)
'-------------------------------------
'����:���ñ��ѡ����
'------------------------------------
    Dim intRow As Integer, intCOl As Integer
    Dim intOldRow As Integer, intOldCol As Integer
    
    On Error GoTo Errhand
    If mblnInit = False Then Exit Sub
    mblnRefresh���� = False
    If tbcThis.Selected.Tag = "����" Then
        vsfCurve.SetFocus
        If blnInit = True Then
            intOldRow = vsfCurve.Row
            intOldCol = vsfCurve.Col
            intRow = vsfCurve.Row
            intCOl = COL_����
            If intRow = vsfCurve.Row And intCOl = vsfCurve.Col Then
                vsfCurve.Col = COL_��λ
            End If
            vsfCurve.Col = COL_����
        Else
            intOldRow = vsfCurve.Row
            intOldCol = vsfCurve.Col
            intRow = vsfCurve.Row
            intCOl = vsfCurve.Col
            If intRow = vsfCurve.Row And intCOl = vsfCurve.Col Then
                If intCOl < vsfCurve.Cols - 1 Then
                    vsfCurve.Col = intCOl + 1
                Else
                    If intRow < vsfCurve.Rows - 1 Then
                        vsfCurve.Row = intRow + 1
                    Else
                        If intRow - 1 > 0 Then
                            vsfCurve.Row = intRow - 1
                        End If
                    End If
                End If
            End If
            vsfCurve.Col = intCOl
        End If
        Call vsfCurve_AfterRowColChange(intOldRow, intOldCol, intRow, intCOl)
    ElseIf tbcThis.Selected.Tag = "���" Then
        If intType = 1 Then
            vsfTab.SetFocus
            If blnInit = True Then
                intOldRow = vsfTab.Row
                intOldCol = vsfTab.Col
                intRow = vsfTab.FixedRows
                intCOl = vsfTab.FixedCols
                If intRow = vsfTab.Row And intCOl = vsfTab.Col Then
                    Call vsfTab_BeforeRowColChange(intRow, intCOl, intRow, intCOl, False)
                End If
                vsfTab.Select vsfTab.FixedRows, vsfTab.FixedCols
            Else
                intOldRow = vsfTab.Row
                intOldCol = vsfTab.Col
                intRow = vsfTab.Row
                intCOl = vsfTab.Col
                vsfTab.Select vsfTab.Row, vsfTab.Col
            End If
            Call vsfTab_AfterRowColChange(intOldRow, intOldCol, intRow, intCOl)
        Else
            vsfTabDetail.SetFocus
            If blnInit = True Then
                intOldRow = vsfTabDetail.Row
                intOldCol = vsfTabDetail.Col
                intRow = vsfTabDetail.FixedRows
                intCOl = vsfTabDetail.FixedCols
                If intRow = vsfTabDetail.Row And intCOl = vsfTabDetail.Col Then
'                    Call vsfTabDetail_BeforeRowColChange(intRow, intCOl, intRow, intCOl, False)
                End If
                vsfTabDetail.Select vsfTabDetail.FixedRows, vsfTabDetail.FixedCols
            Else
                intOldRow = vsfTabDetail.Row
                intOldCol = vsfTabDetail.Col
                intRow = vsfTabDetail.Row
                intCOl = vsfTabDetail.Col
                vsfTabDetail.Select vsfTabDetail.Row, vsfTabDetail.Col
            End If
            Call vsfTabDetail_AfterRowColChange(intOldRow, intOldCol, intRow, intCOl)
            
        End If
    End If
    
     Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Sub picColor_Resize()
    With usrValue
        .Top = -450
        .Left = 0
        .Width = picValue.Width
        .Height = picValue.Height
    End With
End Sub

Private Sub picCurve_Resize()
    Dim i As Integer
    
    On Error Resume Next
    If mblnResize = True Then picSplit.Top = ScaleHeight * 0.6: mblnResize = False

    picSplit.Width = tbcThis.Width
    
    With lblTime
        .Top = picToolBar.Top + lblPtime.Top
        .Left = 200 + picToolBar.Width + picToolBar.Left
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
    End With
    
    With fraTime
        .Top = 0
        .Left = 0
        .Width = picCurve.Width
        .Height = 100 + picToolBar.Top + picToolBar.Height
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
    End With
        
        
    With fraData
        .Left = 0
        .Width = picCurve.Width
        .Top = fraTime.Top + fraTime.Height
        .Height = picSplit.Top - .Top
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
    End With

    With vsfCurve
        .Top = 0
        .Left = 0
        .Width = fraData.Width
        .Height = fraData.Height
    End With

    With picδ��
        .Width = 1080 + 1080 * mintBigSize / 3
        .Height = 1100 + 1100 * mintBigSize / 3
        .Visible = False
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
    End With

    With lstδ��
        .Top = 0
        .Left = 0
        .Width = picδ��.Width
        .Height = picδ��.Height
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
    End With

    With fraDetail
        .Top = picSplit.Top
        .Left = 0
        .Width = picCurve.Width
        .Height = picCurve.Height - picSplit.Top - picSplit.Height
    End With
    
     With vsfDetail
        .Top = 0
        .Left = 0
        .Width = fraDetail.Width
        .Height = fraDetail.Height
    End With
    
    With picValue
        .Width = 2190
        .Height = 2190 - 450
        .Visible = False
    End With
    Call picToolBarReSize
End Sub

Private Sub picToolBarReSize()
    Dim i As Integer
    On Error Resume Next
    lblPtime.Left = 0
    lblPtime.Font.Size = mFontSize + mFontSize * mintBigSize / 3
    lblPtime.Top = 45 + 45 * mintBigSize / 3
    For i = 0 To OptTime.Count - 1
        OptTime(i).Font.Size = mFontSize + mFontSize * mintBigSize / 3
        OptTime(i).Height = 300 + 300 * mintBigSize / 3
        OptTime(i).Width = 350 + 350 * mintBigSize / 3
        OptTime(i).Left = i * OptTime(i).Width + lblPtime.Left + lblPtime.Width + 10
    Next i
    picToolBar.Top = 210
    picToolBar.Width = OptTime(OptTime.Count - 1).Left + OptTime(OptTime.Count - 1).Width
    picToolBar.Height = OptTime(OptTime.Count - 1).Top + OptTime(OptTime.Count - 1).Height
    picToolBar.Left = 50
    picToolBar.Font.Size = mFontSize + mFontSize * mintBigSize / 3
End Sub


Private Sub picEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Call txtEdit_KeyPress(KeyAscii)
    ElseIf KeyAscii = vbKeySpace Then
        If lblCheck.Caption = "��" Then
            lblCheck.Caption = ""
        Else
            lblCheck.Caption = "��"
        End If
    ElseIf KeyAscii = vbKeyReturn Then
        Call txtEdit_KeyDown(KeyAscii, 0)
    ElseIf KeyAscii = vbKeyLeft Then
        If txtEdit.Visible = False Then
            Call vsfTab_KeyDown(vbKeyLeft, 0)
        End If
    End If
End Sub

Private Sub picHour_GotFocus()
    If picHour.Visible = True Then txtHour.SetFocus
End Sub

Private Sub PicLst_GotFocus()
    If PicLst.Visible = False Then Exit Sub
    If Trim(txtLst.Text) = "" Then
        PicLst.Tag = 0
        lstSelect(0).SetFocus
    Else
        PicLst.Tag = 1
        txtLst.SetFocus
    End If
End Sub

'Private Sub picDate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'    Select Case imgDefault.Tag
'        Case 1
'            If imgbtn(0).Enabled = True Then imgbtn(0).Picture = ilsDate.ListImages("nextGreen").Picture
'        Case 2
'            If imgbtn(1).Enabled = True Then imgbtn(1).Picture = ilsDate.ListImages("preGreen").Picture
'    End Select
'
'End Sub

Private Sub picSplit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    picSplit.Tag = 1
    If picSplit.Visible = True Then picSplit.SetFocus
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    If Val(picSplit.Tag) = 0 Then Exit Sub
    
    If picSplit.Top + Y < 4000 Then
        picSplit.Top = 4000
    ElseIf Me.ScaleHeight - (picSplit.Top + Y) < Me.ScaleHeight * 0.4 Then
        picSplit.Top = Me.ScaleHeight * 0.6
    Else
        picSplit.Move picSplit.Left, picSplit.Top + Y
    End If
End Sub

Private Sub picSplit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Val(picSplit.Tag) = 1 Then Call picCurve_Resize

    picSplit.Tag = 0
End Sub



Private Sub LoadOptTime()
'-----------------------------------------
'����:���ݼ�������̬����ʱ��ѡ��ؼ�(OptTime)
'-----------------------------------------
    Dim i As Integer
    For i = 1 To T_BodyStyle.lng������ - 1
        Load OptTime(i)
        OptTime(i).Visible = True
        Set OptTime(i).Container = picToolBar
        OptTime(i).Top = OptTime(i - 1).Top
        OptTime(i).ZOrder 0
    Next i
    Call picToolBarReSize
End Sub

Private Sub UnLoadOptTime()
'------------------------------------------
'����:ж��ʱ��ѡ��ؼ�
'------------------------------------------
    Dim i As Integer
    For i = OptTime.Count - 1 To 1 Step -1
        Unload OptTime(i)
    Next i
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
    If mblnFileBack = True Then lblStb.Caption = "�������������Ѿ��鵵,��������������޸�.": lblStb.ForeColor = 255
End Sub

Private Sub picSplitTab_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    picSplitTab.Tag = 1
    If picSplit.Visible = True Then picSplitTab.SetFocus
End Sub

Private Sub picSplitTab_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    If Val(picSplitTab.Tag) = 0 Then Exit Sub
    
    If picSplitTab.Top + Y < 4000 Then
        picSplitTab.Top = 4000
    ElseIf Me.ScaleHeight - (picSplitTab.Top + Y) < Me.ScaleHeight * 0.3 Then
        picSplitTab.Top = Me.ScaleHeight * 0.7
    Else
        picSplitTab.Move picSplitTab.Left, picSplitTab.Top + Y
    End If
End Sub

Private Sub picSplitTab_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Val(picSplitTab.Tag) = 1 Then Call picTab_Resize
    picSplit.Tag = 0
End Sub

Private Sub picTab_Resize()
    On Error Resume Next
    If mblnResize = True Then picSplitTab.Top = ScaleHeight * 0.6: mblnResize = False
    picSplitTab.Width = tbcThis.Width
    
    With fraTable
         .Top = 0
        .Left = 0
        .Width = picTab.Width
        .Height = picSplitTab.Top
    End With
    
    With vsfTab
        .Top = 100
        .Left = 0
        .Width = fraTable.Width
        .Height = fraTable.Height - .Top
    End With

    picEdit.Visible = False
    txtEdit.Font.Size = mFontSize + mFontSize * mintBigSize / 3
    lblCheck.Font.Size = txtEdit.Font.Size

    With picColor
        .Width = 2190
        .Height = 2190 - 450
        .Visible = False
    End With
    
    With lstSelect(0)
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
        .Visible = False
    End With
    
    With lstSelect(1)
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
        .Visible = False
    End With
    
    With PicLst
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
        .Visible = False
    End With
    
    With txtLst
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
        .Visible = False
    End With
    
    With picHour
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
        .Visible = False
    End With
    
    With lblHour
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
        .Visible = False
    End With
    
    With txtHour
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
        .Visible = False
    End With
    
    With fraTabDetail
        .Top = picSplitTab.Top
        .Left = 0
        .Width = picTab.Width
        .Height = picTab.Height - picSplitTab.Top
    End With
    
    With vsfTabDetail
        .Top = 0
        .Left = 0
        .Width = fraTabDetail.Width
        .Height = fraTabDetail.Height
    End With


End Sub




Private Sub picToolBar_Resize()
    Dim i As Integer
    On Error Resume Next
    With lblTime
        .Top = picToolBar.Top + lblPtime.Top
        .Left = 200 + picToolBar.Width + picToolBar.Left
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
    End With
    
    For i = 0 To OptTime.Count - 1
        OptTime(i).Font.Size = mFontSize + mFontSize * mintBigSize / 3
        OptTime(i).Height = 300 + 300 * mintBigSize / 3
        OptTime(i).Width = 350 + 350 * mintBigSize / 3
        OptTime(i).Left = i * OptTime(i).Width + lblPtime.Left + lblPtime.Width + 10
    Next i
    picToolBar.Top = 210
    picToolBar.Width = OptTime(OptTime.Count - 1).Left + OptTime(OptTime.Count - 1).Width
    picToolBar.Height = OptTime(OptTime.Count - 1).Top + OptTime(OptTime.Count - 1).Height
    picToolBar.Left = 0
    picToolBar.Font.Size = mFontSize + mFontSize * mintBigSize / 3
    
End Sub



Private Sub tbcThis_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If Not mblnInit Then Exit Sub
    
    If Item.Tag = "���" Then
        picNull.Visible = (vsfTab.TextMatrix(1, COL_tab��Ŀ���) = -999)
        If picEdit.Visible = False Then
            Call SetColSelect(True)
        Else
            txtEdit_KeyPress (vbKeyEscape)
            Call SetColSelect
            
        End If
    ElseIf Item.Tag = "����" Then
        picNull.Visible = False
        If mblnStart = False Then
            Call SetColSelect
        Else
            Call SetColSelect(True)
            mblnStart = False
        End If
    End If
End Sub

Private Sub tmrData_Timer()
    Dim i As Integer
    Dim strDay As String
    
    'ˢ��ʱ�㰴ť��ʾ״̬
    
    If mstrBegin = "" Then Exit Sub
    strDay = Format(mstrBegin, "YYYY-MM-DD")
    
    If Format(mstrBegin, "YYYY-MM-DD HH:mm:ss") < Format(mstrBTime, "YYYY-MM-DD HH:mm:ss") Then mstrBegin = mstrBTime
    If Format(mstrEnd, "YYYY-MM-DD HH:mm:ss") > Format(mstrETime, "YYYY-MM-DD HH:mm:ss") Then mstrEnd = mstrETime

    If Format(mstrBegin, "YYYY-MM-DD") = Format(mstrBTime, "YYYY-MM-DD") Or Format(mstrEnd, "YYYY-MM-DD") = Format(mstrETime, "YYYY-MM-DD") Then
        For i = 0 To OptTime.Count - 1
            If OptTime(i).Tag <> "" Then
                If Format(strDay & " " & Split(OptTime(i).Tag, ",")(0), "YYYY-MM-DD HH:mm:ss") > Format(mstrETime, "YYYY-MM-DD HH:mm:ss") Or _
                    Format(strDay & " " & Split(OptTime(i).Tag, ",")(1), "YYYY-MM-DD HH:mm:ss") < Format(mstrBTime, "YYYY-MM-DD HH:mm:ss") Then
                    OptTime(i).Enabled = False
                Else
                    OptTime(i).Enabled = True
                End If
            End If
        Next i
    Else
        For i = 0 To OptTime.Count - 1
            OptTime(i).Enabled = True
        Next i
    End If
End Sub



Private Sub txtEdit_GotFocus()
    Call zlControl.TxtSelAll(txtEdit)
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intCOl As Integer, intRow As Integer
    Dim blnAllow As Boolean
    Dim strData As String
    Dim lngColor As Long
    Dim strTag As String
    
    If KeyCode = vbKeyDown Then
        If picEdit.Visible = False Then Exit Sub
        '��������Ϊ�������͵Ļ��Ŀʹ�ÿ�ݼ����Ե���������ɫ����
        If cmdColor.Visible = True And Shift = vbShiftMask And Val(Split(vsfTab.TextMatrix(vsfTab.Row, COL_tab�ַ���), ",")(1)) = 1 _
            And Val(Split(vsfTab.TextMatrix(vsfTab.Row, COL_tab�ַ���), ",")(5)) = 2 And Val(Split(vsfTab.TextMatrix(vsfTab.Row, COL_tab�ַ���), ",")(4)) = 0 Then
            With picColor
                .Top = picEdit.Top + picEdit.Height
                If .Top + .Height > vsfTab.Top + vsfTab.Height Then
                    .Top = picEdit.Top - .Height
                End If
                If .Top < vsfTab.Top Then .Top = vsfTab.Top
                .Left = picEdit.Left
                .Visible = True
                .ZOrder 0
            End With
            With usrColor
                .Left = 0
                .Top = -450
                .Visible = True
                .ZOrder 0
            End With
            picColor.SetFocus
            usrColor.Color = Val(cmdColor.Tag)
        End If
    ElseIf KeyCode = vbKeyReturn Then
        If Shift = vbShiftMask Then Exit Sub
        '������ݺϷ���
        blnAllow = True
        If picEdit.Visible = True And txtEdit.Tag <> "" Then
            intRow = Split(txtEdit.Tag, "|")(0)
            intCOl = Split(txtEdit.Tag, "|")(1)
            
            If txtEdit.Visible = True Then
                strData = IIf(picHour.Visible = True, "(" & txtHour.Text & "h)", "") & Trim(txtEdit.Text)
                lngColor = txtEdit.ForeColor
            Else
                strData = Trim(lblCheck.Caption)
                lngColor = 0
            End If
            
            If IIf(cmdColor.Visible, strData & "'" & lngColor <> picEdit.Tag, strData <> Split(picEdit.Tag, "'")(0)) Then
                If txtEdit.Tag <> "" Then strTag = txtEdit.Tag
                If Split(txtEdit.Tag, "|")(2) = 1 Then
                    blnAllow = WriteIntoVfgTab(strData, vsfTab)
                Else
                    blnAllow = WriteIntoVfgTab(strData, vsfTabDetail)
                End If
            End If
        End If
        
        If Split(IIf(strTag = "", txtEdit.Tag, strTag), "|")(2) = 1 Then
            If blnAllow = True Then
                '�ƶ�����һ��
                Call vsfTab_KeyDown(vbKeyReturn, Shift)
            Else
                Call vsfTab_EnterCell
            End If
        Else
            If blnAllow = True Then
                '�ƶ�����һ��
                Call vsfTabDetail_KeyDown(vbKeyReturn, Shift)
            Else
                Call vsfTabDetail_EnterCell
            End If
        
        End If
        
    ElseIf KeyCode = vbKeyLeft And txtEdit.SelStart = 0 Then
        If picHour.Visible = False Then
            If Split(txtEdit.Tag, "|")(2) = 1 Then
                Call vsfTab_KeyDown(vbKeyLeft, 0)
            Else
                Call vsfTabDetail_KeyDown(Left, 0)
            End If
        Else
            txtHour.SetFocus
        End If
    End If
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    Dim strVsf As String
    If KeyAscii = vbKeyEscape Then
        If InStr(txtEdit.Tag, "|") <> 0 Then strVsf = Split(txtEdit.Tag, "|")(2)
        With picEdit
            .Visible = False
            .Enabled = False
        End With
        With txtEdit
            .Visible = False
            .Enabled = False
            .Tag = ""
            .Text = ""
        End With
        With cmdColor
            .Visible = False
            .Enabled = False
            .Tag = ""
        End With
        With lstSelect(0)
            .Visible = False
            .Enabled = False
            .Tag = ""
        End With
        With lstSelect(1)
            .Visible = False
            .Enabled = False
            .Tag = ""
        End With
        
        With PicLst
            .Visible = False
            .Tag = ""
        End With
        
        With picHour
            .Visible = False
            .Enabled = False
        End With
        
        With txtHour
            .Visible = False
            .Enabled = False
            .Text = ""
        End With
        
        With lblCheck
            .Visible = False
            .Enabled = False
        End With
        mblnEdit = False
        
        If mblnAllRefresh = False And mblnStart = False Then Call SetColSelect(False, Val(strVsf))
    End If
End Sub

Private Sub txtHour_GotFocus()
    Call zlControl.TxtSelAll(txtHour)
End Sub

Private Sub txtHour_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intCOl As Integer, intRow As Integer
    Dim blnAllow As Boolean
    Dim strData As String
    Dim lngColor As Long
    
    If picHour.Visible = False Then Exit Sub
    If KeyCode = vbKeyReturn And Not (Shift = vbShiftMask) Then
        '������ݺϷ���
        blnAllow = True
        If picEdit.Visible = True And txtEdit.Tag <> "" Then
            intRow = Split(txtEdit.Tag, "|")(0)
            intCOl = Split(txtEdit.Tag, "|")(1)
            
            If txtEdit.Visible = True Then
                strData = IIf(picHour.Visible = True, "(" & txtHour.Text & "h)", "") & Trim(txtEdit.Text)
                lngColor = txtEdit.ForeColor
                If txtEdit.Text = "" Then strData = ""
            Else
                strData = Trim(lblCheck.Caption)
                lngColor = 0
            End If
            
            If strData & "'" & lngColor <> picEdit.Tag Then blnAllow = WriteIntoVfgTab(strData, vsfTab, False, False)
        End If
        If blnAllow = True Then
            '�ƶ�����һ��
            If txtEdit.Enabled = True Then
                txtEdit.SetFocus
            Else
                Call vsfTab_KeyDown(vbKeyReturn, Shift)
            End If
        Else
            txtHour.SetFocus
        End If
    ElseIf KeyCode = vbKeyLeft And txtHour.SelStart = 0 Then
        Call vsfTab_KeyDown(vbKeyLeft, 0)
    End If
End Sub

Private Sub txtHour_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Call txtEdit_KeyPress(vbKeyEscape)
    Else
        If InStr(1, "0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtHour_Validate(Cancel As Boolean)
    Dim strText As String
    strText = txtHour.Text
    If strText = "" Then Exit Sub
    If Not (Val(strText) >= 0 And strText <= 24) Then
        lblStb.Caption = "����Сʱֻ����0��24֮�䣬������¼�룡": lblStb.ForeColor = 255
        Cancel = True
    Else
        txtHour.Text = Val(strText)
    End If
End Sub

Private Sub txtLst_GotFocus()
    PicLst.Tag = 1
    Call zlControl.TxtSelAll(txtLst)
    lstSelect(0).ListIndex = -1
End Sub

Private Sub txtLst_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim blnAllow As Boolean
    
    blnAllow = True
    If KeyCode = vbKeyReturn And Shift = vbShiftMask Then Exit Sub
    If KeyCode = vbKeyReturn Then
        If Trim(txtLst.Text) <> lstSelect(0).Tag Then
            If Val(txtLst.Tag) = 1 Then
                blnAllow = WriteIntoVfgTab(txtLst.Text, vsfTab)
            Else
                blnAllow = WriteIntoVfgTab(txtLst.Text, vsfTabDetail)
            End If
        End If
        If blnAllow = True Then
            If Val(txtLst.Tag) = 1 Then
                Call vsfTab_KeyDown(vbKeyReturn, Shift)
            Else
                Call vsfTabDetail_KeyDown(vbKeyReturn, Shift)
            End If
        End If
    ElseIf KeyCode = vbKeyLeft And txtLst.SelStart = 0 Then
        If Val(txtLst.Tag) = 1 Then
                Call vsfTab_KeyDown(vbKeyLeft, 0)
            Else
                Call vsfTabDetail_KeyDown(vbKeyLeft, 0)
            End If
    ElseIf Shift = vbShiftMask And KeyCode = vbKeyDown Then
        KeyCode = 0
        lstSelect(0).SetFocus
    ElseIf KeyCode = vbKeyEscape Then
         Call txtEdit_KeyPress(vbKeyEscape)
    End If
End Sub

Private Sub usrColor_LostFocus()
    picColor.Visible = False
End Sub

Private Sub usrColor_pOK()
    Dim intRow As Integer, intCOl As Integer
    Dim strTmp As String, lng��Ŀ��� As Long, str��Ŀ���� As String
    Dim strTime As String, arrTime() As String
    
    On Error GoTo Errhand
    If Val(cmdColor.Tag) = usrColor.Color Then picColor.Visible = False:  GoTo GetSetFocus
    cmdColor.Tag = usrColor.Color
    txtEdit.ForeColor = cmdColor.Tag
    picColor.Visible = False
    
    If txtEdit.Tag <> "" Then
        intRow = Val(Split(txtEdit.Tag, "|")(0))
        intCOl = Val(Split(txtEdit.Tag, "|")(1))
    Else
        intRow = vsfTab.Row
        intCOl = vsfTab.Col
    End If
    
    lng��Ŀ��� = Val(vsfTab.TextMatrix(intRow, COL_tab��Ŀ���))
    str��Ŀ���� = vsfTab.TextMatrix(intRow, COL_tab��Ŀ��)
    If vsfTab.TextMatrix(vsfTab.Row, col_tabԭʼʱ��) <> "" Then
        arrTime = Split(vsfTab.TextMatrix(vsfTab.Row, col_tabԭʼʱ��), "'")
        If intCOl - vsfTab.FixedCols < UBound(arrTime) Then
            strTime = arrTime(intCOl - vsfTab.FixedCols)
        End If
    End If
    mrsTableDetail.Filter = "��Ŀ���=" & lng��Ŀ��� & " and ��Ŀ����='" & str��Ŀ���� & "' and ʱ��='" & strTime & "'"
    If mrsTableDetail.RecordCount > 0 Then
        mrsTableDetail!δ��˵�� = cmdColor.Tag
        If mrsTableDetail!״̬ <> 1 Then   'ԭ�е����� �޸ġ�ɾ�����״̬ʼ��Ϊ2
            mrsTableDetail!״̬ = 2
            mrsTableDetail!��� = vsfTab.TextMatrix(intRow, intCOl)
        Else '�����������ݵĴ���
            If Trim(vsfTab.TextMatrix(intRow, intCOl)) = "" Then
                mrsTableDetail.Delete
            Else
                mrsTableDetail!״̬ = 1
                mrsTableDetail!��� = vsfTab.TextMatrix(intRow, intCOl)
            End If
        End If
        mrsTableDetail.Update
    End If
    
GetSetFocus:
    If txtEdit.Visible = True Then txtEdit.SetFocus
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub usrValue_LostFocus()
    picValue.Visible = False
End Sub

Private Sub usrValue_pOK()
    If Val(vsfCurve.Cell(flexcpBackColor, usrValue.Tag, COL_��ɫ, usrValue.Tag, COL_��ɫ)) = usrValue.Color Then picValue.Visible = False: GoTo ErrNext
    vsfCurve.Cell(flexcpBackColor, usrValue.Tag, COL_��ɫ, usrValue.Tag, COL_��ɫ) = usrValue.Color
    If Trim(vsfCurve.TextMatrix(usrValue.Tag, COL_����)) = "" Then GoTo ErrNext
    If vsfCurve.TextMatrix(usrValue.Tag, COL_�޸�״̬) <> 1 Then vsfCurve.TextMatrix(usrValue.Tag, COL_�޸�״̬) = 2
    If Not UpdateCurveDate(vsfCurve, usrValue.Tag, COL_��ɫ, 2) Then vsfCurve.Cell(flexcpBackColor, usrValue.Tag, COL_��ɫ, usrValue.Tag, COL_��ɫ) = usrValue.Color
ErrNext:
    picValue.Visible = False
    If Val(usrValue.Tag) <= vsfCurve.Rows - 1 Then
        vsfCurve.Body.Select Val(usrValue.Tag), COL_����
    End If
    vsfCurve.SetFocus
End Sub

Private Sub AdjustRowFlag(ByRef objVsf As Object, ByVal intRow As Integer)
    '-----------------------------------------------------------------------------------------
    '����:���ѡ��ͼƬ
    '����:
    '-----------------------------------------------------------------------------------------
    If objVsf.FixedCols = 0 Then Exit Sub
    
    If Not (objVsf.Cell(flexcpPicture, intRow, COL_TabNull) Is Nothing) Then Exit Sub
    Set objVsf.Cell(flexcpPicture, 1, COL_TabNull, objVsf.Rows - 1, COL_TabNull) = Nothing
    Set objVsf.Cell(flexcpPicture, intRow, COL_TabNull) = ilstab.ListImages(1).Picture
    
End Sub

Private Sub vsfCurve_AfterEdit(ByVal Row As Long, ByVal Col As Long)

    vsfCurve.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignCenterCenter
    vsfCurve.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000005
    If Col = COL_���� Then
        vsfCurve.TextMatrix(Row, COL_����) = IIf(vsfCurve.EditText = "", " ", Space(Row) & vsfCurve.EditText & Space(Row))
        vsfCurve.TextMatrix(Row, COL_��ɫ) = IIf(vsfCurve.TextMatrix(Row, COL_������) = "2)���±�˵��", Space(Row), vsfCurve.TextMatrix(Row, COL_����))
    End If
    
End Sub


Private Sub vsfCurve_AfterNextRow(ByVal Row As Long, Col As Long)
    If Col = COL_ʱ�� And Row <> vsfCurve.FixedRows Then
        vsfCurve.TextMatrix(Row, COL_ʱ��) = vsfCurve.TextMatrix(Row - 1, COL_ʱ��)
    End If
End Sub

Private Sub vsfCurve_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim lng��Ŀ��� As Integer
    Dim strData As String
    Dim strTmp As String
    
    On Err GoTo Errhand
    '���ز�λ�����б�
    vsfCurve.ComboList(COL_��λ) = ""
    vsfCurve.EditMode(COL_��λ) = 0
    vsfCurve.EditMode(Col_δ��˵��) = 0
    vsfCurve.EditMode(NewCol) = 0

    lng��Ŀ��� = Val(vsfCurve.TextMatrix(NewRow, COL_��Ŀ���))
    strData = Trim(vsfCurve.TextMatrix(NewRow, COL_����))
    Select Case Trim(vsfCurve.TextMatrix(NewRow, COL_������))
        Case "1)����������Ŀ"
            vsfCurve.EditMode(Col_δ��˵��) = 1
            strTmp = GetAllPart(lng��Ŀ���)
            If strTmp <> "" Then
                If lng��Ŀ��� = 2 And InStr(1, strTmp, "|") = 0 Then
                    strTmp = " |����"
                End If
                vsfCurve.ComboList(COL_��λ) = strTmp
                vsfDetail.Body.ColComboList(COL_��λ) = strTmp
                vsfCurve.EditMode(COL_��λ) = 1
            End If
        
        If NewCol = COL_���� Or NewCol = Col_δ��˵�� Or NewCol = COL_ʱ�� Then
            '������Դ
            If InStr(1, ",0,3,9,", "," & Val(vsfCurve.TextMatrix(NewRow, COL_��Դ)) & ",") = 0 Then
                If NewCol = COL_���� Then
                    If lng��Ŀ��� = gint���� And strData = "����" Then vsfCurve.EditMode(NewCol) = 0
                    If lng��Ŀ��� = gint���� Or lng��Ŀ��� = gint��ʹǿ�� Or (lng��Ŀ��� = 2 And mint����Ӧ�� = 2 And mblnEdit���� = True) Then
                        If InStr(1, strData, "/") = 0 Then
                            vsfCurve.EditMode(NewCol) = 1
                        Else
                            If vsfCurve.TextMatrix(NewRow, COL_�༭) = 0 Then vsfCurve.EditMode(NewCol) = 1
                        End If
                    End If
                End If
            Else
                If InStr(1, ",0,3,9,", "," & Val(vsfCurve.TextMatrix(NewRow, COL_��Դ)) & ",") = 0 Then
                    vsfCurve.EditMode(NewCol) = 0
                Else
                    vsfCurve.EditMode(NewCol) = 1
                End If
           
            End If
        End If
    
        Case "2)���±�˵��"
            vsfCurve.EditMode(Col_δ��˵��) = 0
            vsfCurve.EditMode(COL_����) = 1
            vsfCurve.EditMode(COL_ʱ��) = 1
    End Select
    
    strTmp = ""
    If vsfCurve.TextMatrix(NewRow, COL_�ַ���) <> "" Then
        If Trim(Split(vsfCurve.TextMatrix(NewRow, COL_�ַ���), ",")(0)) <> "" Then
            strTmp = "���ݷ�Χ��" & Trim(Split(vsfCurve.TextMatrix(NewRow, COL_�ַ���), ",")(0)) & " "
        End If
    End If
    
    If Trim(vsfCurve.TextMatrix(NewRow, COL_������)) = "1)����������Ŀ" Then
        Select Case lng��Ŀ���
            Case 1 '����
                strTmp = strTmp & Space(4) & "�����±�ʾ��38/37"
            Case gint��ʹǿ��
                strTmp = strTmp & Space(4) & "��ʹ��ʹ��ʾ��6/2"
            Case 2
                If mint����Ӧ�� = 2 And mblnEdit���� Then strTmp = strTmp & Space(4) & "������׾��ʾ��100/130"
        End Select
    ElseIf Trim(vsfCurve.TextMatrix(NewRow, COL_������)) = "2)���±�˵��" Then
        strTmp = "�������а�SHIFT+����˫����ɫ��������ɫ����"
    End If
    lblStb.Caption = strTmp
    lblStb.ForeColor = &H80000012
    '����vsfDetail�������
    If OldRow = NewRow And mblnRefresh���� = True Then Exit Sub
    mblnRefresh���� = True
    If vsfCurve.TextMatrix(NewRow, COL_������) = "1)����������Ŀ" Then
        Call ShowDetail(lng��Ŀ���, NewRow)
    Else
        vsfDetail.Rows = vsfDetail.FixedRows
    End If
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function GetAllPart(ByVal lngNo As Long) As String
    '��ȡ���в�λ
    Dim strValue As String
    Dim strTmp As String
    
    On Error GoTo Errhand
    If Not mrsPart Is Nothing Then
        mrsPart.Filter = "��Ŀ���=" & lngNo
        mrsPart.Sort = "ȱʡ�� DESC"
        With mrsPart
            Do While Not .EOF
                strTmp = IIf(strTmp = "", zlStr.Nvl(!��λ), strTmp & "|" & zlStr.Nvl(!��λ))
            .MoveNext
            Loop
        End With
    End If
    GetAllPart = strTmp
    
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Sub vsfCurve_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngWidth As Long
    If Col = COL_��ɫ Then
        lngWidth = vsfCurve.Body.ColWidth(Col)
        vsfCurve.Body.ColWidth(COL_��ɫ) = 300
        vsfCurve.Body.ColWidth(COL_����) = vsfCurve.Body.ColWidth(COL_����) + lngWidth - 300
        If vsfCurve.Body.ColWidth(COL_����) < 500 Then vsfCurve.Body.ColWidth(COL_����) = 500
        Call vsfCurve_KeyDown(vbKeyDown, vbShiftMask)
    End If
End Sub

Private Sub vsfCurve_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim blnAllow As Boolean
    Dim intType As Integer
    Dim lng��Ŀ��� As Long
    
    On Err GoTo Errhand
    vsfCurve.Tag = vsfCurve.TextMatrix(Row, Col)

    If VsfDeleteRow(1, vsfCurve, Row, Col, Cancel) Then Exit Sub
    Call ShowCurve
    Call ShowTabUpDown
    Cancel = True
    lng��Ŀ��� = vsfCurve.TextMatrix(Row, COL_��Ŀ���)
    Call ShowDetail(lng��Ŀ���, Row)
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function VsfDeleteRow(ByVal intType As Integer, ByVal vsf As Object, ByVal Row As Integer, ByVal Col As Integer, ByRef Cancel As Boolean) As Boolean
    '-----------------------------------------
    '���ܣ�ɾ����
    '-----------------------------------------
    Dim blnAllow As Boolean
    
    On Error GoTo Errhand
    Select Case Col
        Case COL_ʱ��
            vsf.TextMatrix(Row, COL_�޸�״̬) = 2
            vsf.TextMatrix(Row, Col) = ""
            If intType = 3 Then
                intType = 3
            ElseIf Trim(vsf.TextMatrix(Row, COL_������)) = "2)���±�˵��" Then
                intType = 2
            ElseIf Trim(vsf.TextMatrix(Row, COL_������)) = "1)����������Ŀ" Then
                intType = 1
            End If
            blnAllow = True
        Case COL_��λ
            vsf.TextMatrix(Row, COL_�޸�״̬) = 2
            vsf.TextMatrix(Row, Col) = ""
            If intType = 3 Then
                intType = 3
            ElseIf Trim(vsf.TextMatrix(Row, COL_������)) = "2)���±�˵��" Then
                intType = 2
            ElseIf Trim(vsf.TextMatrix(Row, COL_������)) = "1)����������Ŀ" Then
                intType = 1
            End If
            blnAllow = True
        Case COL_����
            If vsf.TextMatrix(Row, Col) <> "" Then
                If intType = 3 Then
                    intType = 3
                    If InStr(1, ",0,3,9,", "," & Val(vsf.TextMatrix(Row, COL_��Դ)) & ",") = 0 Then
                        Cancel = True
                        lblStb.Caption = "�ɻ����¼���������ط�ͬ�����������ݲ���ɾ��."
                        lblStb.ForeColor = 255
                        vsf.Body.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000005
                        VsfDeleteRow = True
                        Exit Function
                    End If
                ElseIf Trim(vsf.TextMatrix(Row, COL_������)) = "2)���±�˵��" Then
                    intType = 2
                ElseIf Trim(vsf.TextMatrix(Row, COL_������)) = "1)����������Ŀ" Then
                    intType = 1
                    If InStr(1, ",0,3,9,", "," & Val(vsf.TextMatrix(Row, COL_��Դ)) & ",") = 0 Then
                        Cancel = True
                        lblStb.Caption = "�ɻ����¼���������ط�ͬ�����������ݲ���ɾ��."
                        lblStb.ForeColor = 255
                        vsf.Body.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000005
                        VsfDeleteRow = True
                        Exit Function
                    End If
                End If
                If CurveRowClear(intType, Row) Then blnAllow = True
            End If
        Case Col_δ��˵��
            If Trim(vsf.TextMatrix(Row, COL_������)) = "1)����������Ŀ" And vsf.TextMatrix(Row, Col) <> "" Then
                intType = 1
                If CurveRowClear(intType, Row) Then blnAllow = True
            End If
        Case COL_ɾ��
            If vsf.TextMatrix(Row, COL_����) <> "" Or vsf.TextMatrix(Row, Col_δ��˵��) <> "" Then
                intType = 3
                If InStr(1, ",0,3,9,", "," & Val(vsf.TextMatrix(Row, COL_��Դ)) & ",") = 0 Then
                    Cancel = True
                    lblStb.Caption = "�ɻ����¼���������ط�ͬ�����������ݲ���ɾ��."
                    lblStb.ForeColor = 255
                    vsf.Body.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000005
                    VsfDeleteRow = True
                    Exit Function
                End If
                If CurveRowClear(intType, Row) Then blnAllow = True
            End If
            
    End Select
    If blnAllow = True Then Call UpdateCurveDate(vsf, Row, Col, intType)
    
     Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CurveRowClear(ByVal intType As Integer, ByVal intRow As Integer) As Boolean
    '-------------------------------
    '������߱������
    '-------------------------------
    On Error Resume Next
    Select Case intType
        Case 1, 2
            vsfCurve.TextMatrix(intRow, COL_�޸�״̬) = IIf(Val(vsfCurve.TextMatrix(intRow, COL_�޸�״̬)) = 1, 4, 3)
            vsfCurve.TextMatrix(intRow, COL_ʱ��) = ""
            vsfCurve.TextMatrix(intRow, COL_����) = Space(intRow) & IIf(vsfCurve.TextMatrix(intRow, COL_������) = "2)���±�˵��", "", "") & Space(intRow)
            vsfCurve.TextMatrix(intRow, COL_��ɫ) = Space(intRow) & IIf(vsfCurve.TextMatrix(intRow, COL_������) = "2)���±�˵��", " ", "") & Space(intRow)
            vsfCurve.TextMatrix(intRow, COL_��λ) = ""
            vsfCurve.TextMatrix(intRow, COL_���Ժϸ�) = ""
            vsfCurve.TextMatrix(intRow, Col_δ��˵��) = ""
            CurveRowClear = True
        Case 3
            vsfDetail.TextMatrix(intRow, COL_�޸�״̬) = IIf(Val(vsfDetail.TextMatrix(intRow, COL_�޸�״̬)) = 1, 4, 3)
            vsfDetail.TextMatrix(intRow, COL_��ʾ) = ""
            vsfDetail.TextMatrix(intRow, COL_ʱ��) = ""
            vsfDetail.TextMatrix(intRow, COL_����) = ""
            vsfDetail.TextMatrix(intRow, COL_��λ) = ""
            vsfDetail.TextMatrix(intRow, COL_���Ժϸ�) = ""
            vsfDetail.TextMatrix(intRow, Col_δ��˵��) = ""
            vsfDetail.TextMatrix(intRow, COL_��Դ) = ""
            CurveRowClear = True
    End Select
End Function


Private Sub vsfCurve_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub vsfCurve_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim arrstrδ��˵��() As String
    Dim i As Integer
    Dim blnSelect As Boolean

    On Error GoTo Errhand
    If Trim(vsfCurve.TextMatrix(Row, COL_������)) <> "1)����������Ŀ" Then Exit Sub
    lstδ��.Tag = "1|" & Row & "|" & Col
    lstδ��.Clear
    If mstrδ��˵�� <> "" Then
        arrstrδ��˵��() = Split(mstrδ��˵��, "'")
        For i = 0 To UBound(arrstrδ��˵��)
            lstδ��.AddItem arrstrδ��˵��(i)
            If arrstrδ��˵��(i) = vsfCurve.TextMatrix(vsfCurve.Row, vsfCurve.Col) Then
                lstδ��.Selected(i) = True
                blnSelect = True
            End If
        Next
    End If
    If blnSelect = False And lstδ��.ListCount <> 0 Then lstδ��.Selected(0) = True
    
    If lstδ��.ListCount > 0 Then
        picδ��.Left = vsfCurve.CellLeft + vsfCurve.Left + 15
        picδ��.Top = fraData.Top + vsfCurve.CellTop + vsfCurve.Top + vsfCurve.CellHeight
        If lstδ��.Height < vsfCurve.CellHeight + 20 Then lstδ��.Height = vsfCurve.CellHeight + 20
        lstδ��.Width = vsfCurve.CellWidth + 20
        picδ��.Height = lstδ��.Height
        picδ��.Width = lstδ��.Width
        picδ��.Visible = True
        lstδ��.Visible = True: lstδ��.Enabled = True
        lstδ��.SetFocus
    End If
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Sub vsfCurve_ComboCloseUp(Row As Long, Col As Long, FinishEdit As Boolean)
    If Trim(vsfCurve.TextMatrix(Row, COL_������)) = "1)����������Ŀ" Then
        Call UpdateCurveDate(vsfCurve, Row, Col, 1, True)
    End If
End Sub

Private Sub vsfCurve_KeyDown(KeyCode As Integer, Shift As Integer)
    picValue.Visible = False
    picValue.Tag = ""
    With vsfCurve
        If .Col > .FixedCols - 1 And .Row > .FixedRows - 1 Then
            If KeyCode = vbKeyDown And Shift = vbShiftMask Then
                If .Col = Col_δ��˵�� Then
                    Call vsfCurve_CellButtonClick(.Row, .Col)
                ElseIf (.Col = COL_���� Or .Col = COL_��ɫ) And .TextMatrix(.Row, COL_������) = "2)���±�˵��" Then
                    vsfCurve.Tag = .TextMatrix(.Row, COL_����)
                    picValue.Top = fraData.Top + .CellTop + .CellHeight + .Top
                    If picValue.Top + picValue.Height > .Top + .Height Then
                        picValue.Top = .CellTop - picValue.Height
                    End If
                    If picValue.Top < .Top Then picValue.Top = .Top
                    picValue.Left = IIf(.Col = COL_��ɫ, .CellLeft, .CellLeft + .CellWidth) + .Left
                    picValue.Visible = True
                    picValue.ZOrder 0
         
                    usrValue.Left = 0
                    usrValue.Top = -450
                    usrValue.Visible = True
                    usrValue.ZOrder 0
                    picValue.SetFocus
                    usrValue.Color = Val(.Cell(flexcpBackColor, .Row, COL_��ɫ, .Row, COL_��ɫ))
                    picValue.Tag = Val(usrValue.Color)
                    usrValue.Tag = .Row
                End If
            End If
        End If
    End With
End Sub

Private Sub vsfCurve_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, ByVal ComboList As String, KeyCode As Integer, ByVal Shift As Integer, Cancel As Boolean)
    If KeyCode = vbKeyReturn Then
        If Col = Col_δ��˵�� Then
            If InStr(1, "," & mstrδ��˵�� & ",", "," & vsfCurve.EditText & ",") = 0 Then
                vsfCurve.TextMatrix(Row, Col) = ""
                vsfCurve.Cell(flexcpData, Row, Col) = ""
            Else
                vsfCurve.TextMatrix(Row, Col) = vsfCurve.EditText
                vsfCurve.Cell(flexcpData, Row, Col) = vsfCurve.EditText
                vsfCurve.TextMatrix(Row, COL_ʱ��) = ""
                vsfCurve.TextMatrix(Row, COL_����) = Space(Row) & IIf(vsfCurve.TextMatrix(Row, COL_������) = "2)���±�˵��", "", "") & Space(Row)
                vsfCurve.TextMatrix(Row, COL_��ɫ) = Space(Row) & IIf(vsfCurve.TextMatrix(Row, COL_������) = "2)���±�˵��", " ", "") & Space(Row)
                vsfCurve.TextMatrix(Row, COL_��λ) = ""
                vsfCurve.TextMatrix(Row, COL_���Ժϸ�) = ""
            End If
        End If
    End If
    If KeyCode = vbKeyDown And Shift = vbShiftMask And Col = COL_���� Then
        Call vsfCurve_KeyDown(KeyCode, Shift)
        Cancel = True
    End If
End Sub

Private Sub vsfCurve_KeyPress(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer, Cancel As Boolean)
    If KeyAscii = 32 Then '
        If Col = COL_���Ժϸ� Then
            If Val(vsfCurve.TextMatrix(Row, COL_����)) <> 0 And Val(vsfCurve.TextMatrix(Row, COL_��Ŀ���)) = gint���� Then
                If vsfCurve.TextMatrix(Row, COL_�޸�״̬) = 1 Then vsfCurve.TextMatrix(Row, COL_�޸�״̬) = 1
                If vsfCurve.TextMatrix(Row, COL_�޸�״̬) = 0 Then vsfCurve.TextMatrix(Row, COL_�޸�״̬) = 2
                If vsfCurve.TextMatrix(Row, Col) = "" Then
                    vsfCurve.TextMatrix(Row, Col) = "��"
                    vsfCurve.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignCenterCenter
                Else
                    vsfCurve.TextMatrix(Row, Col) = ""
                End If
                Call UpdateCurveDate(vsfCurve, Row, Col, 1)
                Call ShowDetail(gint����, Row)
            End If
        End If
        If Col = COL_��ɫ And vsfCurve.TextMatrix(Row, COL_������) = "2)���±�˵��" Then
            Call vsfCurve_KeyDown(vbKeyDown, vbShiftMask)
        End If
    End If
End Sub

Private Sub vsfCurve_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim lngNo As Long
    
    On Error Resume Next
    lngNo = Val(vsfCurve.TextMatrix(Row, COL_��Ŀ���))
    
    If KeyAscii <> vbKeyReturn Then
        If lngNo <> 0 Then
            If Col = COL_ʱ�� Then
                 If FilterKeyAscii(KeyAscii, 99, "0123456789:") = 0 Then KeyAscii = 0
            Else
                If vsfCurve.TextMatrix(Row, COL_������) = "1)����������Ŀ" Then
                    If Col <> Col_δ��˵�� Then
                        If lngNo = 2 And mint����Ӧ�� = 2 And mblnEdit���� = True Then
                            If FilterKeyAscii(KeyAscii, 99, "0123456789./") = 0 Then KeyAscii = 0
                        ElseIf lngNo = gint��ʹǿ�� Then
                            If FilterKeyAscii(KeyAscii, 99, "0123456789./") = 0 Then KeyAscii = 0
                        ElseIf lngNo = gint���� Then
                            '���²����м��
                        Else
                            If FilterKeyAscii(KeyAscii, 99, "0123456789.") = 0 Then KeyAscii = 0
                        End If
                    Else
                        If FilterKeyAscii(KeyAscii, 99, "'") > 0 Then KeyAscii = 0
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub vsfCurve_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lng��Ŀ��� As Long
    Dim strName As String
    Dim strData As String
    Dim strTime As String
    
    On Err GoTo Errhand
    lng��Ŀ��� = Val(vsfCurve.TextMatrix(Row, COL_��Ŀ���))
    strName = vsfCurve.TextMatrix(Row, COL_��Ŀ��)
    Select Case Col
        Case COL_����
            vsfCurve.TextMatrix(Row, Col) = IIf(Trim(vsfCurve.TextMatrix(Row, Col)) = "", " ", Trim(vsfCurve.TextMatrix(Row, Col)))
            If Row <> mOptRow.�ϱ� And Row <> mOptRow.�±� Then
                vsfCurve.TextMatrix(Row, COL_��ɫ) = vsfCurve.TextMatrix(Row, Col)
            Else
                vsfCurve.TextMatrix(Row, Col) = Trim(vsfCurve.TextMatrix(Row, Col))
            End If
            vsfCurve.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000018
            strData = RTrim(LTrim(vsfCurve.TextMatrix(Row, Col)))
        Case COL_ʱ��
            vsfCurve.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000018
    End Select
    
    vsfCurve.Tag = Trim(vsfCurve.TextMatrix(Row, Col))
    vsfCurve.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignCenterCenter
    If Col = COL_���� Or Col = Col_δ��˵�� Or Col = COL_ʱ�� Then
          '������Դ
        If InStr(1, ",0,3,9,", "," & Val(vsfCurve.TextMatrix(Row, COL_��Դ)) & ",") = 0 Then
            If Col = COL_���� Then
                If lng��Ŀ��� = gint���� And strData = "����" Then GoTo NotEdit
                If lng��Ŀ��� = gint���� Or lng��Ŀ��� = gint��ʹǿ�� Or (lng��Ŀ��� = 2 And mint����Ӧ�� = 2 And mblnEdit���� = True) Then
                    If InStr(1, strData, "/") = 0 Then
                        GoTo GONext
                    Else
                        If Val(vsfCurve.TextMatrix(Row, COL_�༭)) = 0 Then GoTo GONext
                    End If
                End If
            End If
NotEdit:
            Cancel = True
            vsfCurve.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000005
            vsfCurve.Body.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignCenterCenter
            If lng��Ŀ��� = gint���� Then
                lblStb.Caption = "ͬ��������[" & strName & "]����ֻ�����޸������²���."
            ElseIf lng��Ŀ��� = gint��ʹǿ�� Then
                lblStb.Caption = "ͬ��������[" & strName & "]����ֻ�����޸���ʹ��ʹ����."
            ElseIf lng��Ŀ��� = 2 And mint����Ӧ�� = 2 And mblnEdit���� = True Then
                If mbln����������ʾ Then
                    lblStb.Caption = "ͬ��������[" & strName & "]����ֻ�����޸��������Ჿ��.����/����"
                Else
                    lblStb.Caption = "ͬ��������[" & strName & "]����ֻ�����޸��������Ჿ��.����/���� "
                End If
            Else
                lblStb.Caption = "�ɻ����¼���������ط�ͬ�����������ݲ����޸�"
            End If
            lblStb.ForeColor = 255
            vsfCurve.Body.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000005
        End If
    ElseIf COL_���Ժϸ� = Col Then
        vsfCurve.Body.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignCenterCenter
    End If
GONext:
    If mblnFileBack = True Then
        Cancel = True
        vsfCurve.Body.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignCenterCenter
        lblStb.Caption = "�������������Ѿ��鵵,��������������޸�."
        lblStb.ForeColor = 255
        vsfCurve.Body.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000005
    End If
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

'Private Function StartEdit()
'    '�༭֮ǰ���ݼ�
'
'
'
'
'End Function


Private Sub vsfCurve_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lng��Ŀ��� As Long
    On Err GoTo Errhand
    If CheckPutData(1, vsfCurve, Row, Col, Cancel) Then
        lng��Ŀ��� = vsfCurve.TextMatrix(Row, COL_��Ŀ���)
        Call ShowDetail(lng��Ŀ���, Row)
    End If
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function CheckPutData(ByVal intType As Integer, ByVal vsf As Object, ByVal Row As Integer, ByVal Col As Integer, ByRef Cancel As Boolean) As Boolean
    '---------------------------------------
    '���ܣ��������ֵ��������
    '---------------------------------------
    Dim strTime As String
    Dim strMsg As String
    Dim strText As String, strData As String
    Dim strCenterTime As String
    Dim strName As String, strValue As String, strֵ�� As String, strInfo As String
    Dim intС�� As Integer
    Dim i As Integer, intCount As Integer
    Dim lng��Ŀ��� As Long
    Dim blnOK As Boolean
    Dim lngCount As Long
    Dim arrValue() As String

    On Err GoTo Errhand
    '������ݺϷ���
    
    strValue = vsf.Tag
    strֵ�� = Split(vsfCurve.TextMatrix(vsfCurve.Row, COL_�ַ���), ",")(0)
    lng��Ŀ��� = Val(vsfCurve.TextMatrix(vsfCurve.Row, COL_��Ŀ���))
    strName = vsfCurve.TextMatrix(vsfCurve.Row, COL_��Ŀ��)
    intС�� = Val(Split(vsfCurve.TextMatrix(vsfCurve.Row, COL_�ַ���), ",")(2))
    
    If vsf.Col = COL_ʱ�� Then
        strText = vsf.EditText
        strText = CToData(strText)
        If strText <> "" Then
            If Not CheckDateTime(Row, strName, strText) Then
                strMsg = lblStb.Caption
                GoTo ErrInfo
            End If
            vsf.EditText = strText
            vsf.TextMatrix(Row, COL_�޸�״̬) = 5
            Select Case vsf.TextMatrix(Row, COL_������)
                Case "1)����������Ŀ"
                    intType = 1
                Case "2)���±�˵��"
                    intType = 2
                Case Else
                    intType = 3
            End Select
            mrsCurve.Filter = "��Ŀ���=" & lng��Ŀ��� & " and  ʱ��='" & Format(dtpDate.Value & " " & strText, "YYYY-MM-DD hh:mm:ss") & "'"
            If mrsCurve.RecordCount > 0 And strValue <> strText Then
                strMsg = "��ǰʱ���Ѵ������ݣ�����������ʱ��"
                GoTo ErrInfo
            End If
            Call UpdateCurveDate(vsf, Row, Col, intType)
            CheckPutData = True
        End If
    End If
    
    If Col = COL_���� Then
    
        Select Case vsf.TextMatrix(Row, COL_������)
            Case "1)����������Ŀ"
                intType = 1
                GoTo CheckPoint
            Case "2)���±�˵��"
                If InStr(1, ",2,6,", "," & Val(vsf.TextMatrix(Row, COL_��Ŀ���)) & ",") <> 0 Then
                    picValue.Tag = vsf.Cell(flexcpBackColor, Row, COL_��ɫ, Row, COL_��ɫ)
                    intType = 2: GoTo CheckTag
                End If
            Case Else
                intType = 3
                GoTo CheckPoint
        End Select
    End If
    
    Exit Function
    
CheckPoint:
    '�������
    If Trim(vsf.EditText) <> "" And strֵ�� <> "" Then
        strInfo = vsf.EditText
        If vsf.TextMatrix(Row, COL_ʱ��) = "" Then
        mrsCurve.Filter = "��Ŀ���=" & lng��Ŀ��� & " and  ʱ��='" & Format(GetCenterTime(mstrBegin, mstrEnd), "YYYY-MM-DD hh:mm:ss") & "'"
            If mrsCurve.RecordCount > 0 Then
                strMsg = "��ǰĬ��ʱ���Ѵ������ݣ���������ʱ��"
                GoTo ErrInfo
            End If
        End If
        '���������������/��Ҫ�������������
        If lng��Ŀ��� = 2 And mint����Ӧ�� = 2 And mblnEdit���� = True Then
            If InStr(1, strInfo, "/") > 0 Then
                If Split(Trim(strInfo), "/")(1) = "" Or Split(Trim(strInfo), "/")(0) = "" Then
                    strMsg = strName & "����¼�����" & Space(4) & "��������:����/����"
                    GoTo ErrInfo
                Else
                    If Not IsNumeric(Split(Trim(strInfo), "/")(0)) Or Not IsNumeric(Split(Trim(strInfo), "/")(1)) Then
                        strMsg = strName & "����¼�����" & Space(4) & "��Ч��Χ:" & strֵ��
                        GoTo ErrInfo
                    End If
                End If
            End If
        End If
        
        If lng��Ŀ��� <> 1 And lng��Ŀ��� <> gint��ʹǿ�� And Not (lng��Ŀ��� = 2 And mint����Ӧ�� = 2 And mblnEdit���� = True) Then
            If InStr(1, strInfo, "/") Then
                strMsg = strName & "����¼�����" & Space(4) & "��Ч��Χ:" & strֵ��
                GoTo ErrInfo
            End If
        End If

        If UBound(Split(strInfo, "/")) > 1 Then
            strMsg = strName & "����¼��������飡"
            GoTo ErrInfo
        End If
        
        '�����������Ч��Χ���Ƿ���Ч
        arrValue = Split(strInfo, "/")
        lngCount = UBound(arrValue)
        For i = 0 To lngCount
            blnOK = False
            strText = arrValue(i)
            If i = 0 Then
                '����������Ŀ��Ҫ���˵�δ��˵��
                If InStr(1, strText, ";") <> 0 And UBound(arrValue) = 0 Then strText = Split(strText, ";")(1)
                If InStr(1, IIf(lng��Ŀ��� = gint����, ",����,", ""), "," & strText & ",") = 0 Then
                    blnOK = False
                Else
                    blnOK = True
                End If
            End If
            
            If Not blnOK Then
                If Not IsNumeric(strText) Then
                    strMsg = strName & "����¼�����" & Space(4) & "��Ч��Χ:" & strֵ��
                    GoTo ErrInfo
                End If
            End If
            
            If Not blnOK And strText <> "" Then
                strText = Format(Val(strText), "#0" & IIf(intС�� > 0, ".", "") & String(intС��, "0"))
                If strText = Val(strText) Then strText = Val(strText)
                If Left(strText, 1) = "." Then strText = 0 & strText
            End If
            If IsNumeric(Split(strֵ��, "��")(0)) And IsNumeric(strText) Then
                If Not (Val(strText) >= Split(strֵ��, "��")(0) And Val(strText) <= Split(strֵ��, "��")(1)) Then
                    strMsg = strName & "������Ч��Χ(" & strֵ�� & "),����!"
                    GoTo ErrInfo
                End If
            End If
            If i = 0 Then
                strInfo = strText
            Else
                strInfo = strInfo & "/" & strText
            End If
        Next i
    End If
    If strInfo <> vsf.EditText Then vsf.EditText = strInfo
    strData = vsf.EditText
    '����������Դ<>0,3,9�� ����,�������� ���б༭(�������º������������¼��������,��������)
    If InStr(1, ",0,3,9,", "," & Val(vsf.TextMatrix(Row, COL_��Դ)) & ",") = 0 Then
        If Col = COL_���� Then
            If lng��Ŀ��� = gint���� Or lng��Ŀ��� = gint��ʹǿ�� Or (lng��Ŀ��� = 2 And mint����Ӧ�� = 2 And mblnEdit���� = True) Then
                '�����������¼�뷽ʽ�Ƿ���ȷ������/����
                If lng��Ŀ��� = 2 And (InStr(strValue, "/") > 0 Or InStr(strValue, "/") = 0) And mbln����������ʾ Then
                    If InStr(1, strData, "/") <> 0 Then
                        strData = Split(strData, "/")(1)
                    Else
                        strData = strData
                    End If
                    If strData <> vsf.TextMatrix(Row, COL_ԭֵ) Then
                        If mbln����������ʾ Then
                            strMsg = "ͬ��������[" & strName & "]����ֻ�����޸��������Ჿ��.����/����"
                        Else
                            strMsg = "ͬ��������[" & strName & "]����ֻ�����޸��������Ჿ��.����/���� "
                        End If
                        vsf.TextMatrix(Row, COL_����) = Space(Row) & Trim(CStr(vsf.TextMatrix(Row, COL_ԭֵ))) & Space(Row)
                        vsf.TextMatrix(Row, COL_��ɫ) = vsf.TextMatrix(Row, COL_����)
                        GoTo ErrInfo
                    End If
                Else
                    strValue = CStr(vsf.TextMatrix(Row, COL_ԭֵ))
                    If InStr(1, strData, "/") <> 0 Then
                        strData = Split(strData, "/")(0)
                    End If
                
                    If InStr(1, vsf.TextMatrix(Row, COL_ԭֵ), "/") = 0 Then
                        If strData <> vsf.TextMatrix(Row, COL_ԭֵ) Then
                            If lng��Ŀ��� = gint���� Then
                                strMsg = "ͬ��������[" & strName & "]����ֻ�����޸������²���."
                            ElseIf lng��Ŀ��� = gint��ʹǿ�� Then
                                strMsg = "ͬ��������[" & strName & "]����ֻ�����޸���ʹ��ʹ����."
                            Else
                                If mbln����������ʾ Then
                                    strMsg = "ͬ��������[" & strName & "]����ֻ�����޸��������Ჿ��.����/����"
                                Else
                                    strMsg = "ͬ��������[" & strName & "]����ֻ�����޸��������Ჿ��.����/���� "
                                End If
                            End If
                            
                            vsf.TextMatrix(Row, COL_����) = Space(Row) & Trim(CStr(vsf.TextMatrix(Row, COL_ԭֵ))) & Space(Row)
                            vsf.TextMatrix(Row, COL_��ɫ) = vsf.TextMatrix(Row, COL_����)
                            GoTo ErrInfo
                        End If
                    Else
                        If Val(vsf.TextMatrix(Row, COL_�༭)) <> 0 Then
                            If strData <> vsf.TextMatrix(Row, COL_ԭֵ) Then
                                If lng��Ŀ��� = gint���� Then
                                    strMsg = "ͬ��������[" & strName & "]�����������������,�������޸�."
                                ElseIf lng��Ŀ��� = gint��ʹǿ�� Then
                                    strMsg = "ͬ��������[" & strName & "]�������������ʹ��ʹ,�������޸�."
                                Else
                                    strMsg = "ͬ��������[" & strName & "]�������������������,�������޸�."
                                End If
                                vsf.TextMatrix(Row, COL_����) = Space(Row) & CStr(vsf.TextMatrix(Row, COL_ԭֵ)) & Space(Row)
                                vsf.TextMatrix(Row, COL_��ɫ) = vsf.TextMatrix(Row, COL_����)
                                GoTo ErrInfo
                            End If
                        Else
                            If strData <> Split(vsf.TextMatrix(Row, COL_ԭֵ), "/")(0) Then
                                If lng��Ŀ��� = gint���� Then
                                    strMsg = "ͬ��������[" & strName & "]����ֻ�����޸������²���."
                                ElseIf lng��Ŀ��� = gint��ʹǿ�� Then
                                    strMsg = "ͬ��������[" & strName & "]����ֻ�����޸���ʹ��ʹ����."
                                Else
                                    If mbln����������ʾ Then
                                        strMsg = "ͬ��������[" & strName & "]����ֻ�����޸��������Ჿ��.����/����"
                                    Else
                                        strMsg = "ͬ��������[" & strName & "]����ֻ�����޸��������Ჿ��.����/���� "
                                    End If
                                End If
                                vsf.TextMatrix(Row, COL_����) = Space(Row) & CStr(vsf.TextMatrix(Row, COL_ԭֵ)) & Space(Row)
                                vsf.TextMatrix(Row, COL_��ɫ) = vsf.TextMatrix(Row, COL_����)
                                GoTo ErrInfo
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If

    '��ʾȱʡ��λ
    If vsf.TextMatrix(Row, COL_��λ) = "" And Trim(vsf.EditText) <> "" Then
        mrsPart.Filter = "��Ŀ���=" & lng��Ŀ��� & " and ȱʡ��=1"
        If mrsPart.RecordCount > 0 Then
            vsf.TextMatrix(Row, COL_��λ) = CStr(zlStr.Nvl(mrsPart!��λ))
        End If
    End If
    
    GoTo UpdateData
    Exit Function
    
CheckTag:
    GoTo UpdateData
    Exit Function
    
ErrInfo:
    lblStb.Caption = strMsg
    lblStb.ForeColor = 255
    vsf.TextMatrix(Row, COL_����) = Space(Row) & strValue & Space(Row)
    vsf.TextMatrix(Row, COL_��ɫ) = vsf.TextMatrix(Row, COL_����)
    Cancel = True
    Exit Function
    
UpdateData:
    
    If vsf.EditText = strValue Then Exit Function
    intCount = 0
    For i = COL_���� To Col_δ��˵��
        If Trim(vsf.TextMatrix(Row, i)) <> "" Or vsf.EditText <> "" Then
           intCount = intCount + 1
           Exit For
        End If
    Next
    If intCount = 0 Then
        If Trim(vsf.TextMatrix(Row, COL_�޸�״̬)) = 1 Then
            vsf.TextMatrix(Row, COL_�޸�״̬) = 4
        Else
            vsf.TextMatrix(Row, COL_�޸�״̬) = 3
        End If
    Else
        vsf.TextMatrix(Row, COL_�޸�״̬) = 2
    End If
    Call UpdateCurveDate(vsf, Row, Col, intType)
    CheckPutData = True
    
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CToData(ByVal strData As String) As String
'-----------------------------------------------
'���ܣ����ʱ�䣬ת��Ϊʱ�䣺����
'-----------------------------------------------
    Dim strCenterTime As String
    If InStr(1, Trim(strData), ":") = 0 And strData <> "" Then
        Select Case Len(strData)
        Case 3, 4
            strData = String(4 - Len(strData), "0") & strData
            strData = Mid(strData, 1, 2) & ":" & Mid(strData, 3)
        Case Is < 3
            strData = String(2 - Len(strData), "0") & strData
            strCenterTime = GetCenterTime(mstrBegin, mstrEnd)
            strData = Format(strCenterTime, "HH") & ":" & strData
        End Select
    End If
    CToData = strData
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub vsfDetail_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    vsfDetail.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignCenterCenter
    vsfDetail.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000005
End Sub

Private Sub vsfDetail_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim intRow As Integer
    Dim strTemp As String
    Dim lng��Ŀ��� As String
    Dim strData As String
    Dim strTmp As String
    Dim strName As String
    On Error GoTo Errhand
    
    If vsfDetail.Cols < COL_ɾ�� Or NewRow < vsfDetail.FixedRows Then Exit Sub
    If OldRow < vsfDetail.Rows Then Call vsfDetail.SelectRow(vsfDetail, OldRow, NewRow, &HFFC0C0)
    For intRow = vsfDetail.FixedRows To vsfDetail.Rows - 2
        vsfDetail.Body.Cell(flexcpPicture, intRow, COL_ɾ��, NewRow, COL_ɾ��) = Nothing
    Next
    If NewRow > vsfDetail.FixedRows - 1 And NewRow < vsfDetail.Rows - 1 Then
        vsfDetail.Body.Cell(flexcpPicture, NewRow, COL_ɾ��, NewRow, COL_ɾ��) = ilsDetail.ListImages(1).Picture
    End If
    vsfDetail.EditMode(NewCol) = 0
    
    lng��Ŀ��� = Val(vsfCurve.TextMatrix(vsfCurve.Row, COL_��Ŀ���))
    strName = Val(vsfCurve.TextMatrix(vsfCurve.Row, COL_��Ŀ��))
    strData = Trim(vsfDetail.TextMatrix(NewRow, COL_����))
    vsfDetail.EditMode(Col_δ��˵��) = 1
    vsfDetail.EditMode(COL_��λ) = 1
    
    If NewCol = COL_���� Or NewCol = Col_δ��˵�� Or NewCol = COL_ʱ�� Then
        '������Դ
        If InStr(1, ",0,3,9,", "," & Val(vsfDetail.TextMatrix(NewRow, COL_��Դ)) & ",") = 0 Then
            If NewCol = COL_���� Then
                If lng��Ŀ��� = gint���� And strData = "����" Then vsfDetail.EditMode(NewCol) = 0
                If lng��Ŀ��� = gint���� Or lng��Ŀ��� = gint��ʹǿ�� Or (lng��Ŀ��� = 2 And mint����Ӧ�� = 2 And mblnEdit���� = True) Then
                    If InStr(1, strData, "/") = 0 Then
                        vsfDetail.EditMode(NewCol) = 1
                    Else
                        If Val(vsfDetail.TextMatrix(NewRow, COL_�༭)) = 0 Then vsfDetail.EditMode(NewCol) = 1
                    End If
                End If
            End If
        Else
            If InStr(1, ",0,3,9,", "," & Val(vsfDetail.TextMatrix(NewRow, COL_��Դ)) & ",") = 0 Then
                vsfDetail.EditMode(NewCol) = 0
            Else
                vsfDetail.EditMode(NewCol) = 1
            End If
       
        End If
    End If
    
    strTmp = ""
    If vsfCurve.TextMatrix(vsfCurve.Row, COL_�ַ���) <> "" Then
        If Trim(Split(vsfCurve.TextMatrix(vsfCurve.Row, COL_�ַ���), ",")(0)) <> "" Then
            strTmp = "���ݷ�Χ��" & Trim(Split(vsfCurve.TextMatrix(vsfCurve.Row, COL_�ַ���), ",")(0)) & " "
        End If
    End If
    
    If Trim(vsfDetail.TextMatrix(NewRow, COL_������)) = "1)����������Ŀ" Then
        Select Case lng��Ŀ���
            Case 1 '����
                strTmp = strTmp & Space(4) & "�����±�ʾ��38/37"
            Case gint��ʹǿ��
                strTmp = strTmp & Space(4) & "��ʹ��ʹ��ʾ��6/2"
            Case 2
                If mint����Ӧ�� = 2 And mblnEdit���� Then strTmp = strTmp & Space(4) & "������׾��ʾ��100/130"
        End Select
    End If
    If strTmp <> "" Then lblStb.Caption = strTmp
    lblStb.ForeColor = &H80000012
    Exit Sub
    
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub




Private Sub vsfDetail_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim blnAllow As Boolean
    Dim intType As Integer
    Dim lng��Ŀ��� As Long
    
    If Not vsfDetail.TextMatrix(Row, COL_ʱ��) <> "" And (vsfDetail.TextMatrix(Row, COL_����) <> "" Or vsfDetail.TextMatrix(Row, Col_δ��˵��) <> "") Then Exit Sub
    vsfDetail.Tag = vsfDetail.TextMatrix(Row, Col)
    If VsfDeleteRow(3, vsfDetail, Row, Col, Cancel) Then Exit Sub
    Call ShowCurve
    Cancel = False
End Sub



Private Sub vsfDetail_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    If vsfDetail.TextMatrix(Row, COL_����) = "" And vsfDetail.TextMatrix(Row, Col_δ��˵��) = "" Then Cancel = True
End Sub

Private Sub vsfDetail_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim arrstrδ��˵��() As String
    Dim i As Integer
    Dim blnSelect As Boolean

    On Error GoTo Errhand
    
    lstδ��.Tag = "2|" & Row & "|" & Col
    lstδ��.Clear
    If mstrδ��˵�� <> "" Then
        arrstrδ��˵��() = Split(mstrδ��˵��, "'")
        For i = 0 To UBound(arrstrδ��˵��)
            lstδ��.AddItem arrstrδ��˵��(i)
            If arrstrδ��˵��(i) = vsfDetail.TextMatrix(vsfDetail.Row, Col_δ��˵��) Then
                lstδ��.Selected(i) = True
                blnSelect = True
            End If
        Next
    End If
    If blnSelect = False And lstδ��.ListCount <> 0 Then lstδ��.Selected(0) = True
    
    If lstδ��.ListCount > 0 Then
                picδ��.Left = vsfDetail.CellLeft + vsfDetail.Left + 15
                picδ��.Top = fraDetail.Top + vsfDetail.CellTop + vsfDetail.Top + vsfDetail.CellHeight
                If lstδ��.Height < vsfDetail.CellHeight + 20 Then lstδ��.Height = vsfDetail.CellHeight + 20
                If picδ��.Top + picδ��.Height > picCurve.Height Then picδ��.Top = fraDetail.Top + vsfDetail.Body.CellTop + vsfDetail.Top - lstδ��.Height + 20
                lstδ��.Width = vsfDetail.Body.CellWidth + 20
                picδ��.Height = lstδ��.Height
                picδ��.Width = lstδ��.Width
                picδ��.Visible = True
                lstδ��.Visible = True: lstδ��.Enabled = True
                lstδ��.SetFocus
            End If
    Exit Sub

Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Sub vsfDetail_ComboCloseUp(Row As Long, Col As Long, FinishEdit As Boolean)
    Call UpdateCurveDate(vsfDetail, Row, Col, 3, True)
End Sub

Private Sub vsfDetail_KeyDown(KeyCode As Integer, Shift As Integer)
    If vsfDetail.Col > vsfDetail.FixedCols - 1 And vsfDetail.Row > vsfDetail.FixedRows - 1 Then
        If KeyCode = vbKeyDown And Shift = vbShiftMask Then
            If vsfDetail.Col = Col_δ��˵�� Then
                Call vsfDetail_CellButtonClick(vsfDetail.Row, vsfDetail.Col)
            End If
        End If
    End If
End Sub


Private Sub vsfDetail_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, ByVal ComboList As String, KeyCode As Integer, ByVal Shift As Integer, Cancel As Boolean)
    If KeyCode = vbKeyReturn Then
        If Col = Col_δ��˵�� Then
            If InStr(1, "," & mstrδ��˵�� & ",", "," & vsfDetail.EditText & ",") = 0 Then
                vsfDetail.TextMatrix(Row, Col) = ""
                vsfDetail.Cell(flexcpData, Row, Col) = ""
            Else
                vsfCurve.TextMatrix(Row, Col) = vsfCurve.EditText
                vsfCurve.Cell(flexcpData, Row, Col) = vsfCurve.EditText
                vsfCurve.TextMatrix(Row, COL_��ʾ) = ""
                vsfCurve.TextMatrix(Row, COL_ʱ��) = ""
                vsfCurve.TextMatrix(Row, COL_����) = ""
                vsfCurve.TextMatrix(Row, COL_��λ) = ""
                vsfCurve.TextMatrix(Row, COL_���Ժϸ�) = ""
                vsfCurve.TextMatrix(Row, COL_��Դ) = ""
            End If
        End If
    End If
End Sub

Private Sub vsfDetail_KeyPress(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer, Cancel As Boolean)
    Dim intRow As Integer
    
    If Col = COL_ɾ�� Then
        Call vsfDetail_BeforeDeleteRow(Row, Col, Cancel)
        If Cancel = False Then vsfDetail.RemoveItem (Row)
    End If
    
    If KeyAscii = 32 Then '
        Select Case Col
            Case COL_���Ժϸ�
                If Trim(vsfDetail.TextMatrix(Row, COL_����)) <> "" And Val(vsfDetail.TextMatrix(Row, COL_��Ŀ���)) = gint���� Then
                    If vsfDetail.TextMatrix(Row, Col) = "" Then
                        vsfDetail.TextMatrix(Row, Col) = "��"
                        vsfDetail.Body.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignCenterCenter
                    Else
                        vsfDetail.TextMatrix(Row, Col) = ""
                    End If
                    If vsfDetail.TextMatrix(Row, COL_�޸�״̬) = 1 Then vsfDetail.TextMatrix(Row, COL_�޸�״̬) = 1
                    If vsfDetail.TextMatrix(Row, COL_�޸�״̬) = 0 Then vsfDetail.TextMatrix(Row, COL_�޸�״̬) = 2
                    Call UpdateCurveDate(vsfDetail, Row, Col, 3)
                    Call ShowCurve
                    Call ShowTabUpDown
                End If
            Case COL_��ʾ
                If Trim(vsfDetail.TextMatrix(Row, COL_����)) <> "" Or Trim(vsfDetail.TextMatrix(Row, Col_δ��˵��)) <> "" Then
                
                    For intRow = vsfDetail.FixedRows To vsfDetail.Rows - 2
                        If vsfCurve.TextMatrix(vsfCurve.Row, col_ԭʼʱ��) = vsfDetail.TextMatrix(intRow, col_ԭʼʱ��) Then
                            vsfDetail.TextMatrix(intRow, COL_�޸�״̬) = 6
                            Call UpdateCurveDate(vsfDetail, intRow, Col, 3)
                        End If
                        If vsfDetail.TextMatrix(intRow, COL_��ʾ) = "��" Then
                            vsfDetail.TextMatrix(intRow, COL_��ʾ) = ""
                            vsfDetail.TextMatrix(intRow, COL_�޸�״̬) = 6
                            Call UpdateCurveDate(vsfDetail, intRow, Col, 3)
                            Exit For
                        End If
                    Next
                    
                    vsfDetail.TextMatrix(Row, COL_�޸�״̬) = IIf(vsfDetail.TextMatrix(Row, COL_�޸�״̬) = 0, 6, 2)
                    If intRow <> Row Then
                        vsfDetail.TextMatrix(Row, COL_��ʾ) = "��"
                        vsfDetail.Body.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignCenterCenter
                    Else
                        vsfDetail.TextMatrix(Row, Col) = ""
                    End If
                    
                    Call UpdateCurveDate(vsfDetail, Row, Col, 3)
                    Call ShowCurve
                    Call ShowTabUpDown
                End If
               
        End Select
    End If
End Sub

Private Sub vsfDetail_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim lngNo As Long
    
    On Error Resume Next
    lngNo = Val(vsfDetail.TextMatrix(Row, COL_��Ŀ���))
    
    If KeyAscii <> vbKeyReturn Then
        If lngNo <> 0 Then
            If Col = COL_ʱ�� Then
                 If FilterKeyAscii(KeyAscii, 99, "0123456789:") = 0 Then KeyAscii = 0
            Else
                If vsfCurve.TextMatrix(Row, COL_������) = "1)����������Ŀ" Then
                    If Col <> Col_δ��˵�� Then
                        If lngNo = 2 And mint����Ӧ�� = 2 And mblnEdit���� = True Then
                            If FilterKeyAscii(KeyAscii, 99, "0123456789./") = 0 Then KeyAscii = 0
                        ElseIf lngNo = gint��ʹǿ�� Then
                            If FilterKeyAscii(KeyAscii, 99, "0123456789./") = 0 Then KeyAscii = 0
                        ElseIf lngNo = gint���� Then
                            '���²����м��
                        Else
                            If FilterKeyAscii(KeyAscii, 99, "0123456789.") = 0 Then KeyAscii = 0
                        End If
                    Else
                        If FilterKeyAscii(KeyAscii, 99, "'") > 0 Then KeyAscii = 0
                    End If
                End If
            End If
        End If
    End If
End Sub


Private Sub vsfDetail_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lng��Ŀ��� As Long
    Dim strName As String
    Dim strData As String
    Dim strTime As String
    
    On Err GoTo Errhand
    lng��Ŀ��� = Val(vsfDetail.TextMatrix(Row, COL_��Ŀ���))
    strName = vsfCurve.TextMatrix(vsfCurve.Row, COL_��Ŀ��)
    Select Case Col
        Case COL_����
            vsfDetail.TextMatrix(Row, Col) = Trim(vsfDetail.TextMatrix(Row, Col))
            strData = RTrim(LTrim(vsfDetail.TextMatrix(Row, Col)))
            vsfDetail.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000018
        Case COL_ʱ��
            vsfDetail.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000018
    End Select
    vsfDetail.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignCenterCenter
    vsfDetail.Tag = Trim(vsfDetail.TextMatrix(Row, Col))
    If Col = COL_���� Or Col = Col_δ��˵�� Or Col = COL_ʱ�� Then
          '������Դ
        If InStr(1, ",0,3,9,", "," & Val(vsfDetail.TextMatrix(Row, COL_��Դ)) & ",") = 0 Then
            If Col = COL_���� Then
                If lng��Ŀ��� = gint���� And strData = "����" Then GoTo NotEdit
                If lng��Ŀ��� = gint���� Or lng��Ŀ��� = gint��ʹǿ�� Or (lng��Ŀ��� = 2 And mint����Ӧ�� = 2 And mblnEdit���� = True) Then
                    If InStr(1, strData, "/") = 0 Then
                        GoTo GONext
                    Else
                        If Val(vsfDetail.TextMatrix(Row, COL_�༭)) = 0 Then GoTo GONext
                    End If
                End If
            End If
NotEdit:
            Cancel = True
            vsfDetail.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000005
            vsfDetail.Body.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignCenterCenter
            If lng��Ŀ��� = gint���� Then
                lblStb.Caption = "ͬ��������[" & strName & "]����ֻ�����޸������²���."
            ElseIf lng��Ŀ��� = gint��ʹǿ�� Then
                lblStb.Caption = "ͬ��������[" & strName & "]����ֻ�����޸���ʹ��ʹ����."
            ElseIf lng��Ŀ��� = 2 And mint����Ӧ�� = 2 And mblnEdit���� = True Then
                If mbln����������ʾ Then
                    lblStb.Caption = "ͬ��������[" & strName & "]����ֻ�����޸��������Ჿ��.����/����"
                Else
                    lblStb.Caption = "ͬ��������[" & strName & "]����ֻ�����޸��������Ჿ��.����/���� "
                End If
            Else
                lblStb.Caption = "�ɻ����¼���������ط�ͬ�����������ݲ����޸�"
            End If
            lblStb.ForeColor = 255
            vsfDetail.Body.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000005
        End If
    ElseIf COL_���Ժϸ� = Col Then
        vsfDetail.Body.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignCenterCenter
    End If
GONext:
    If mblnFileBack = True Then
        Cancel = True
        vsfDetail.Body.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignCenterCenter
        lblStb.Caption = "�������������Ѿ��鵵,��������������޸�."
        lblStb.ForeColor = 255
        vsfDetail.Body.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000005
    End If
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vsfDetail_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    On Err GoTo Errhand
    If CheckPutData(3, vsfDetail, Row, Col, Cancel) Then
        Call ShowCurve
        Call ShowDetail(vsfCurve.TextMatrix(vsfCurve.Row, COL_��Ŀ���), Row)
        vsfDetail.Row = Row
    End If
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vsfTab_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strData As String, lngType As Long
    Dim strTime As String, strInfo As String
    Dim lngNo As Long, strName As String, strTmp As String, strֵ�� As String
    Dim strChildNO As String
    Dim lngChildNO As Long
    Dim arrChildNo() As String
    Dim arrTime() As String
    Dim arrStr() As String
    Dim lng��Ŀ��� As Long
    Dim intRow As Integer
    Dim i As Integer
    Dim cbrControl As Object, cbrChild As Object
    Dim rsTotle As New ADODB.Recordset
    Dim blnCheck As Boolean
    Dim strText As String
    
    On Error GoTo Errhand
    If mblnInit = False Then Exit Sub
    If vsfTab.Tag = "NO" Then Exit Sub
    If NewRow < vsfTab.FixedRows Or NewCol < vsfTab.FixedCols Or NewCol > (Split(vsfTab.RowData(NewRow), ";")(0) + vsfTab.FixedCols - 1) Then Exit Sub
    Call AdjustRowFlag(vsfTab, NewRow)
    With vsfTab
        lngNo = Val(.TextMatrix(NewRow, COL_tab��Ŀ���))
        strName = .TextMatrix(NewRow, COL_tab��Ŀ��)
        strTmp = .TextMatrix(NewRow, COL_tab�ַ���)
        strTmp = strTmp & String(8 - UBound(Split(strTmp, ",")), ",")
        arrStr = Split(strTmp, ",")
        strֵ�� = arrStr(0)
        
        If strֵ�� = "" Then
            strInfo = ""
        Else
            strInfo = strName & "��Ч��Χ:" & strֵ��
        End If
        
        If lngNo = 4 And strName = "Ѫѹ" Then 'Ѫѹ
            strInfo = strInfo & Space(4) & "¼�����:����ѹ/����ѹ"
            mrsCurInfo.Filter = ""
            mrsCurInfo.Sort = "����"
            strTmp = ""
            Do While Not mrsCurInfo.EOF
                strTmp = strTmp & "��" & Nvl(mrsCurInfo!����)
                mrsCurInfo.MoveNext
            Loop
            strTmp = Mid(strTmp, 2)
            If strTmp <> "" Then strInfo = strInfo & "��(" & strTmp & ")"
        End If
        
        If Val(arrStr(4)) = 4 Then strInfo = strInfo & Space(4) & "������Ŀ" & Space(4) & "¼�����:����¼��" & IIf(mbln���ܵ��� = True, "����", "����") & "�����ݡ�"
    End With
    lblStb.Caption = strInfo
    lblStb.ForeColor = &H80000012
    
    strData = vsfTab.RowData(NewRow)
    lng��Ŀ��� = zlStr.Nvl(vsfTab.TextMatrix(NewRow, COL_tab��Ŀ���))
    If strData <> "" Then
        strTime = GetAnimalItemTime(vsfTab.Row, NewCol - vsfTab.FixedCols + 1, 0, strInfo)
        If strInfo <> "" Then lblStb.Caption = strInfo: lblStb.ForeColor = 255: Exit Sub
        If InStr(1, strTime, ";") > 0 Then arrTime = Split(strTime, ";")
        lngType = Val(Split(strData, ";")(1))
        With vsfTabDetail
        .Rows = .FixedRows
        .Rows = .FixedRows + 1
        vsfTabDetail.TextMatrix(.FixedRows, COL_tab�ַ���) = vsfTab.TextMatrix(NewRow, COL_tab�ַ���)
        vsfTabDetail.TextMatrix(.FixedRows, COL_tab��Ŀ���) = vsfTab.TextMatrix(NewRow, COL_tab��Ŀ���)
        vsfTabDetail.TextMatrix(.FixedRows, COL_tab��Ŀ����) = vsfTab.TextMatrix(NewRow, COL_tab��Ŀ��)
        vsfTabDetail.TextMatrix(.FixedRows, COL_tab��Ŀ��) = vsfTab.TextMatrix(NewRow, COL_tab��Ŀ��)
        vsfTabDetail.Tag = vsfTabDetail.Rows - 1
        .ColHidden(.FixedCols - 1) = lngType <> 3
        If lngType = 3 Then .ColHidden(.FixedCols - 1) = IsLastTotal(lng��Ŀ���)
        .MergeCellsFixed = flexMergeFree
        .MergeCol(.FixedCols - 1) = True
        If lngType = 3 Then
            mrsTableDetail.Filter = "��Ŀ���= " & lng��Ŀ��� & " and ʱ�� > '" & arrTime(0) & "' and ʱ�� <= '" & arrTime(1) & "' and ��¼���� <> 11 and ״̬<>4 "
            If mrsTableDetail.RecordCount > 0 Then Call ShowTabDetail(.Rows - 1, NewRow, 0)
            Set rsTotle = ReturnTotle(lng��Ŀ���, arrTime(0), arrTime(1))
            .Rows = .Rows + 1
            intRow = .Rows - 1
             rsTotle.Filter = ""
            Do While Not rsTotle.EOF
               
                vsfTabDetail.TextMatrix(intRow, vsfTabDetail.FixedCols - 1) = zlStr.Nvl(rsTotle!��Ŀ����)
                vsfTabDetail.TextMatrix(intRow, vsfTabDetail.FixedCols) = Format(Split(strTime, ";")(0), "hh:mm") & "��" & Format(Split(strTime, ";")(1), "hh:mm")
                vsfTabDetail.TextMatrix(intRow, vsfTabDetail.FixedCols + 1) = Val(Nvl(rsTotle!��ֵ))
                Select Case rsTotle!������Դ
                    Case 0, 9
                        vsfTabDetail.TextMatrix(intRow, vsfTabDetail.FixedCols + 2) = "���µ�¼��"
                    Case 1
                        vsfTabDetail.TextMatrix(intRow, vsfTabDetail.FixedCols + 2) = "��¼��ͬ��"
                    Case 3
                        vsfTabDetail.TextMatrix(intRow, vsfTabDetail.FixedCols + 2) = "�ƶ��豸¼��"
                    Case Else
                        vsfTabDetail.TextMatrix(intRow, vsfTabDetail.FixedCols + 2) = "�����豸ͬ��"
                End Select
                .RowData(intRow) = "3"
                
                .Cell(flexcpBackColor, .Rows - 1, .FixedCols, .Rows - 1, .Cols - 1) = &H8000000F
                intRow = intRow + 1
                .Rows = .Rows + 1
                rsTotle.MoveNext
            Loop
        
        .Rows = .Rows - 1
            
        Else
            mrsTableDetail.Filter = "��Ŀ��� =" & lng��Ŀ��� & " and ʱ�� >= '" & arrTime(0) & "' and ʱ�� <= '" & arrTime(1) & "' and ״̬<>4"
            If mrsTableDetail.RecordCount > 0 Then Call ShowTabDetail(.FixedRows, NewRow, lngType)
        End If

        .Cell(flexcpAlignment, 0, 2, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
    End With
    End If
     
     '��������Ƿ������޸�
    mrsTable.Filter = "��Ŀ���=" & lngNo & " and ��Ŀ����='" & strName & "'" & _
        "   and �к�=" & NewCol - vsfTab.FixedCols + 1
    If mrsTable.RecordCount > 0 Then
        If InStr(1, ",0,3,9,", "," & Val(mrsTable!������Դ) & ",") = 0 Then
            lblStb.Caption = "������Դ�ڻ����¼����PDA�����ݲ��ܽ����޸ġ�ɾ������"
            lblStb.ForeColor = 255
            Exit Sub
        End If
    End If
    
    If InStr(1, strData, ";") > 0 Then
        If Split(strData, ";")(1) = 3 And Not mbln¼��Сʱ Then
            lblStb.Caption = "�������ݽ����޸Ļ���Сʱ�����ܽ��������޸ġ�ɾ������"
            lblStb.ForeColor = 255
            Exit Sub
        End If
    End If
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If

End Sub

Private Function IsLastTotal(ByVal lngNo As Long)
    '����:�Ƿ������һ���Ļ�����Ŀ
    If mrsCollect Is Nothing Then Exit Function
    If mrsCollect.State = adStateOpen Then
        mrsCollect.Filter = "�����=" & lngNo
        If mrsCollect.RecordCount > 0 Then
            IsLastTotal = False
        Else
            IsLastTotal = True
        End If
    End If

End Function


Private Function ReturnTotle(ByVal lngItemNO As Long, ByVal strBTime As String, ByVal strETime As String) As ADODB.Recordset
    '-------------
    '���ܣ��������ϸ
    '�����������Ŀ�ĵ�һ���ӽڵ� ��Ȼ������⼶�ڵ㼰�������ӽڵ�
    '-------------
    Dim rsTemp As New ADODB.Recordset
    Dim rsTotle As New ADODB.Recordset
    Dim rsNO As ADODB.Recordset
    Dim strValue As String
    Dim strValues As String
    Dim strName As String
    Dim strFileds As String
    Dim str��Ŀ��� As String
    Dim lng��Ŀ��� As Long
    Dim dblData As Double
    Dim blnNumeric As Boolean
    
    On Error GoTo Errhand
    '��ʼ����¼��
    strFileds = "��ʼʱ��," & adLongVarChar & ",20|����ʱ��," & adLongVarChar & ",20|��Ŀ���," & adDouble & ",18|��Ŀ����," & _
                adLongVarChar & ",20|��ֵ," & adLongVarChar & ",100|������Դ," & adDouble & ",1"
    Call Record_Init(rsTotle, strFileds)
    strFileds = "��ʼʱ��|����ʱ��|��Ŀ���|��Ŀ����|��ֵ|������Դ"
    
    Set rsNO = GetChildNo(lngItemNO)
    rsNO.Filter = ""
    Do While Not rsNO.EOF
        Set rsTemp = SetCollectPItem(rsNO!���)
        mrsTableDetail.Filter = "��Ŀ���=" & rsNO!��� & " and ʱ�� > '" & strBTime & "' and ʱ�� <= '" & strETime & "'"
        lng��Ŀ��� = rsNO!���
        dblData = 0
        Do While Not mrsTableDetail.EOF
            dblData = dblData + Val(Nvl(mrsTableDetail!���))
            mrsTableDetail.MoveNext
        Loop
        mrsTableDetail.Filter = "��Ŀ��� =" & lng��Ŀ���
        strName = IIf(mrsTableDetail.RecordCount > 0, mrsTableDetail!��Ŀ����, vsfTab.TextMatrix(vsfTab.Row, COL_tab��Ŀ��))
        rsTemp.Filter = ""
        Do While Not rsTemp.EOF
            '��������ϸ����Ҫ����
            If Val(Nvl(rsTemp!���, 0)) <> lngItemNO Then
                mrsTableDetail.Filter = "��Ŀ���=" & rsTemp!��� & " and ʱ�� > '" & strBTime & "' and ʱ�� <= '" & strETime & "'"
                Do While Not mrsTableDetail.EOF
                    dblData = dblData + Val(Nvl(mrsTableDetail!���))
                    If blnNumeric = False Then blnNumeric = IsNumeric(Nvl(mrsTableDetail!���))
                    mrsTableDetail.MoveNext
                Loop
            End If
            rsTemp.MoveNext
        Loop
        strValue = IIf(dblData = 0 And blnNumeric = False, "", IIf(strValue = "", "", "(" & strValue & "h)") & IIf(Left(dblData, 1) = ".", "0", "") & dblData)
        strValues = strBTime & "|" & strETime & "|" & lng��Ŀ��� & "|" & strName & "|" & strValue & "|0"
        If Val(strValue) <> 0 Then Call Record_Add(rsTotle, strFileds, strValues)
        strValue = ""
        rsNO.MoveNext
    Loop
        
    Set ReturnTotle = rsTotle
    
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function SetCollectPItem(ByVal lngItemNO As Long) As ADODB.Recordset
'---------------------------------------------------------------------------
'����:���ݸ���ĿID������֯����Ŀ
'---------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim rsCollect As New ADODB.Recordset
    Dim strFileds As String, strValues As String
    Dim lngNo As Long
    
    On Error GoTo Errhand
    
    '��ʼ����¼��
    strFileds = "���," & adDouble & ",18|�����," & adDouble & ",18"
    Call Record_Init(rsTemp, strFileds)
    Call Record_Init(rsCollect, strFileds)
    strFileds = "���|�����"
    
    mrsCollect.Filter = 0
   '���Ƽ�¼��
    With mrsCollect
        Do While Not .EOF
            strValues = Val(Nvl(!���)) & "|" & Val(Nvl(!�����))
            Call Record_Add(rsCollect, strFileds, strValues)
            .MoveNext
        Loop
    End With
    
    rsCollect.Filter = "�����=" & lngItemNO
    With rsCollect
        Do While Not .EOF
            strValues = Val(Nvl(!���)) & "|" & lngItemNO
            Call Record_Add(rsTemp, strFileds, strValues)
            lngNo = Val(Nvl(!���))
            'ѭ���ݹ����(��ȡ���������)
            Call SetCollectCItem(rsTemp, lngItemNO, lngNo)
            .MoveNext
        Loop
    End With
    
    Set SetCollectPItem = rsTemp
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub SetCollectCItem(rsTemp As ADODB.Recordset, ByVal lngParent As Long, ByVal lngItemNO As Long)
'����: SetCollectPItem ����
    
    Dim rsCollect As New ADODB.Recordset
    Dim strFileds As String, strValues As String
    Dim lngNo As Long
    
    On Error GoTo Errhand
    '��ʼ����¼��
    strFileds = "���," & adDouble & ",18|�����," & adDouble & ",18"
    Call Record_Init(rsCollect, strFileds)
    strFileds = "���|�����"
    
    mrsCollect.Filter = 0
   '���Ƽ�¼��
    With mrsCollect
        Do While Not .EOF
            strValues = Val(Nvl(!���)) & "|" & Val(Nvl(!�����))
            Call Record_Add(rsCollect, strFileds, strValues)
            .MoveNext
        Loop
    End With
    
    rsCollect.Filter = "�����=" & lngItemNO
    With rsCollect
        Do While Not .EOF
            strValues = Val(Nvl(!���)) & "|" & lngParent
            Call Record_Add(rsTemp, strFileds, strValues)
            lngNo = Val(Nvl(!���))
            'ѭ���ݹ����(��ȡ���������)
            Call SetCollectCItem(rsTemp, lngParent, lngNo)
            .MoveNext
        Loop
    End With
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Sub vsfTab_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    mblnScroll = True
    Call vsfTab_EnterCell
    mblnScroll = False
End Sub

Private Sub vsfTab_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    With vsfTab
        If vsfTab.Tag = "NO" Then Exit Sub
        If NewRow >= .FixedRows And NewCol >= .FixedRows Then
            If NewCol < .FixedCols + (Split(.TextMatrix(NewRow, COL_tab�ַ���), ",")(3)) Then
                mrsTable.Filter = "��Ŀ���=" & Val(.TextMatrix(NewRow, COL_tab��Ŀ���)) & " and ��Ŀ����='" & .TextMatrix(NewRow, COL_tab��Ŀ��) & "'" & _
                    "   and �к�=" & NewCol - .FixedCols + 1
                If mrsTable.RecordCount > 0 Then
                    If InStr(1, ",0,3,9,", "," & Val(mrsTable!������Դ) & ",") = 0 Then
                        .FocusRect = 0
                        .HighLight = 0
                    Else
                        .FocusRect = flexFocusSolid
                    End If
                Else
                    .FocusRect = flexFocusSolid
                End If
            Else
                .FocusRect = 0
                .HighLight = 0
            End If
        Else
            .FocusRect = flexFocusNone
        End If
    End With
End Sub

Private Sub vsfTab_DblClick()
    With vsfTab
        If vsfTab.Tag = "NO" Then Exit Sub
        If .Row > .FixedRows - 1 And .Col > .FixedCols - 1 And .FocusRect = flexFocusSolid Then
            mblnEdit = True
            Call vsfTab_EnterCell
        End If
    End With
End Sub

Private Sub vsfTab_EnterCell()
    Dim strInfo As String
    Dim strValue As String, strValue1 As String
    Dim strTmp As String, strData As String
    Dim strTime As String
    Dim blnAllow As Boolean
    Dim blnEdit As Boolean
    Dim blnSelect As Boolean
    Dim intType As Integer
    Dim intƵ�� As Integer, int��Ŀ���� As Integer, int��Ŀ���� As Integer
    Dim i As Integer, j As Integer
    Dim intNum As Integer, intLen As Integer, intRow As Integer, intCOl As Integer
    Dim lngItemNO As Long, lngColor As Long
    Dim arrValue() As String, arrValue1() As String
    Dim arrStr() As String
    
    If Not mblnInit Then Exit Sub
    If vsfTab.Tag = "NO" Then Exit Sub
    blnAllow = True
    blnEdit = True
    blnSelect = True
    
    If picEdit.Visible = True And txtEdit.Tag <> "" Then
        intRow = Split(txtEdit.Tag, "|")(0)
        intCOl = Split(txtEdit.Tag, "|")(1)
        If Split(txtEdit.Tag, "|")(2) <> 1 Then txtEdit_KeyPress (vbKeyEscape): Exit Sub
        If txtEdit.Visible = True Then
            strData = IIf(picHour.Visible = True, "(" & txtHour.Text & "h)", "") & Trim(txtEdit.Text)
            lngColor = txtEdit.ForeColor
            If txtEdit.Text = "" Then strData = ""
        Else
            strData = Trim(lblCheck.Caption)
            lngColor = 0
        End If
        If Split(txtEdit.Tag, "|")(2) = 2 Then blnEdit = False
        If IIf(cmdColor.Visible, strData & "'" & lngColor <> picEdit.Tag, strData <> Split(picEdit.Tag, "'")(0)) Then blnAllow = WriteIntoVfgTab(strData, vsfTab, False, True, strInfo)
        If cmdColor.Visible = True Then vsfTab.Cell(flexcpForeColor, intRow, intCOl, intRow, intCOl) = Val(cmdColor.Tag)
        mblnEdit = blnEdit
    End If
    
    '���ݲ��Ϸ�
    If blnAllow = False Then
        If vsfTab.Row <> intRow Then vsfTab.Row = intRow
        If vsfTab.Col <> intCOl Then vsfTab.Col = intCOl
        GoTo ErrFouce
        Exit Sub
    End If
    
    If vsfTab.Row < vsfTab.FixedRows And vsfTab.Col < vsfTab.FixedCols Then Exit Sub
    If Not vsfTab.RowIsVisible(vsfTab.Row) Then Exit Sub
    If Not mblnScroll And vsfTab.Visible Then vsfTab.SetFocus
    
    '�������б༭�ؼ�
    picδ��.Visible = False
    picEdit.Visible = False
    picEdit.Tag = ""
    txtEdit.Tag = "": txtEdit.Visible = False: txtEdit.Enabled = False
    picHour.Visible = False: picHour.Enabled = False
    txtHour.Tag = "": txtHour.Visible = False: txtHour.Enabled = False
    lblCheck.Visible = False: lblCheck.Enabled = False
    cmdColor.Visible = False
    cmdColor.Enabled = False
    cmdColor.Tag = 0
    picColor.Visible = False
    PicLst.Visible = False
    PicLst.Tag = ""
    txtLst.Visible = False: txtLst.Text = ""
    lstSelect(0).Visible = False
    lstSelect(0).Enabled = False
    lstSelect(0).Tag = ""
    lstSelect(1).Visible = False
    lstSelect(1).Enabled = False
    lstSelect(1).Tag = ""
    
    If mblnFileBack = True Then
        lblStb.Caption = "�������������Ѿ��鵵,��������������޸�."
        mblnEdit = False
        GoTo ErrInfo
    End If
    
    If mblnEdit = False Then Exit Sub
    
    With vsfTab
        If .Row > .FixedRows - 1 And .Col > .FixedCols - 1 And vsfTab.Col < .FixedCols + Val(Split(vsfTab.TextMatrix(vsfTab.Row, COL_tab�ַ���), ",")(3)) Then
            intType = Val(Split(.TextMatrix(vsfTab.Row, COL_tab�ַ���), ",")(4))
            intƵ�� = Val(Split(.TextMatrix(vsfTab.Row, COL_tab�ַ���), ",")(3))
            int��Ŀ���� = Val(Split(.TextMatrix(vsfTab.Row, COL_tab�ַ���), ",")(1))
            int��Ŀ���� = Val(Split(.TextMatrix(vsfTab.Row, COL_tab�ַ���), ",")(5))
            'ʱ���м���Ƿ���Ա༭
            If .Row Mod 2 = 1 And (Split(vsfTab.RowData(.Row), ";")(1) = 2 Or Split(vsfTab.RowData(.Row), ";")(1) = 3) Then
                strInfo = "���ڻ������ݻ򲨶����ݵ�ʱ��β��ܽ����޸ġ�ɾ������"
                GoTo ErrInfo
            End If
    
            '���¼����Ŀʱ���Ƿ񳬳��û����õ�ʱ�䷶Χ���ǲ�¼��Χ
            strTime = GetAnimalItemTime(.Row, .Col - vsfTab.FixedCols + 1, 0, strInfo)
            If .Row Mod 2 = 1 And .TextMatrix(.Row, .Col) <> "" Then
               If CDate(dtpDate.Value & " " & .TextMatrix(.Row, .Col)) < CDate(Split(strTime, ";")(0)) Then strInfo = "¼��ʱ��С�����µ���ʼʱ��"
               If CDate(dtpDate.Value & " " & .TextMatrix(.Row, .Col)) > CDate(Split(strTime, ";")(1)) Then strInfo = "¼��ʱ��������µ���¼ʱ��"
            End If
            If strInfo <> "" Then
                mblnEdit = False
                GoTo ErrInfo
            End If
            '��鲨����Ŀ
            If IsWaveItem(Val(.TextMatrix(.Row, COL_tab��Ŀ���))) And InStr(1, Trim(.TextMatrix(.Row, .Col)), "-") <> 0 Then
                strInfo = "������ֵ�Ѿ��γɲ�����Χ�Ĳ�����Ŀ���ܽ����޸ġ�ɾ������"
                GoTo ErrInfo
            End If
             '���������Դ�Ƿ����Ի����¼����PDA
            mrsTable.Filter = "��Ŀ���=" & Val(.TextMatrix(.Row, COL_tab��Ŀ���)) & " and ��Ŀ����='" & .TextMatrix(.Row, COL_tab��Ŀ��) & "'" & _
                "   and �к�=" & .Col - .FixedCols + 1
            If mrsTable.RecordCount > 0 Then
                If InStr(1, ",0,3,9,", "," & Val(mrsTable!������Դ) & ",") = 0 Then
                    blnEdit = False
                End If
                cmdColor.Tag = Val(mrsTable!δ��˵��)
            End If
            
            'ȫ�������ʾ¼��ʱ��,ͬ��������Ҳ�����޸�ʱ��
            If blnEdit = False And Not (intType = 4 And intƵ�� = 1 And mbln¼��Сʱ = True) Then
                strInfo = "������Դ�ڻ����¼����PDA�����ݲ��ܽ����޸ġ�ɾ������"
                GoTo ErrInfo
            End If
            
            '��ȫ��������ݲ������޸�
            If intType = 4 And Not (intƵ�� = 1 And mbln¼��Сʱ = True) Then
                strInfo = "�������ݵĻ���ֵ���ܽ����޸ġ�ɾ������"
                GoTo ErrInfo
            End If
            
            If Not (intType = 2 Or intType = 3) Or vsfTab.Row Mod 2 = 1 Then
                picEdit.Width = .CellWidth + 10
                picEdit.Height = .CellHeight - 5
                picEdit.Top = .CellTop + .Top + 20
                picEdit.Left = .CellLeft + .Left + 15
                picEdit.Enabled = True
                picEdit.Visible = True
                picEdit.ZOrder 0
                txtEdit.Top = 0
                txtEdit.Left = 0
                txtEdit.Height = picEdit.Height
            End If
            '������Ŀ�������������͵Ļ��Ŀ����������������ɫ
            If int��Ŀ���� = 1 And intType = 0 And int��Ŀ���� = 2 And vsfTab.Row Mod 2 = 0 Then   '�ı����ͣ�� ��Ŀ
                cmdColor.Top = 0
                cmdColor.Height = picEdit.Height
                cmdColor.Width = 300
                cmdColor.Left = picEdit.Width - cmdColor.Width
                txtEdit.Width = cmdColor.Left
                cmdColor.Enabled = True
                cmdColor.Visible = True
                GoTo ShowText
            ElseIf intType = 4 And intƵ�� = 1 And mbln¼��Сʱ = True Then 'ȫ���������ʾ����ʱ��
                txtHour.Top = 10
                txtHour.Left = 10
                txtHour.Width = picHour.TextWidth("111")
                txtHour.Height = txtEdit.Height
                txtHour.MaxLength = 2
                txtHour.Visible = True
                txtHour.Enabled = True
                
                lblHour.Left = txtHour.Left + txtHour.Width
                lblHour.Top = 10
                lblHour.Visible = True
                lblHour.Enabled = True
                
                picHour.Top = -10
                picHour.Left = -10
                picHour.Width = lblHour.Left + lblHour.Width + picHour.TextWidth("1") \ 2
                picHour.Height = picEdit.Height + 20
                picHour.Visible = True
                picHour.Enabled = True
                picHour.ZOrder 0
                
                txtEdit.Top = 10
                txtEdit.Left = picHour.Left + picHour.Width + 10
                txtEdit.Width = picEdit.Width - picHour.Width + 10
                
                strTmp = vsfTab.TextMatrix(vsfTab.Row, COL_tab�ַ���)
                lngItemNO = Val(vsfTab.TextMatrix(vsfTab.Row, COL_tab��Ŀ���))
                strTmp = strTmp & String(8 - UBound(Split(strTmp, ",")), ",")
                arrStr = Split(strTmp, ",")
                intNum = Val(arrStr(2))
                intLen = Val(arrStr(6))
                
                If intLen <> 0 Then
                    If lngItemNO <> 4 Then
                        txtEdit.MaxLength = intLen + IIf(intNum = 0, 0, 1)
                    Else
                        txtEdit.MaxLength = intLen * 2 + 1 + IIf(intNum = 0, 0, 1) * 2
                    End If
                Else
                    txtEdit.MaxLength = 0
                End If
                
                If vsfTab.Row Mod 2 = 1 Then txtEdit.MaxLength = 5
                
                If InStr(1, .TextMatrix(vsfTab.Row, vsfTab.Col), ")") > 0 Then
                    txtHour.Text = Replace(Replace(Split(.TextMatrix(vsfTab.Row, vsfTab.Col), ")")(0), "(", ""), "h", "")
                    txtEdit.Text = Split(.TextMatrix(vsfTab.Row, vsfTab.Col), ")")(1)
                Else
                    txtEdit.Text = .TextMatrix(vsfTab.Row, vsfTab.Col)
                End If
                picEdit.Tag = .TextMatrix(vsfTab.Row, vsfTab.Col) & "'" & .Cell(flexcpForeColor, vsfTab.Row, vsfTab.Col)
                txtEdit.Tag = vsfTab.Row & "|" & vsfTab.Col & "|" & "1"
                txtEdit.ForeColor = cmdColor.Tag
                txtEdit.Visible = True
                txtEdit.Enabled = False
                txtEdit.ZOrder 0
                picHour.SetFocus
            ElseIf (intType = 2 Or intType = 3) And vsfTab.Row Mod 2 = 0 Then '��ѡ��ѡ
                strValue = Split(.TextMatrix(vsfTab.Row, COL_tab�ַ���), ",")(0)
                Select Case intType
                    Case 2
                        If Left(strValue, 1) <> ":" Then strValue = ":" & strValue
                        intType = 0
                    Case 3
                        intType = 1
                End Select
                
                arrValue = Split(strValue, ":")
                lstSelect(intType).Clear
                PicLst.Tag = "1"
                For i = 0 To UBound(arrValue)
                    If Left(arrValue(i), 1) = "��" Then arrValue(i) = Mid(arrValue(i), 2): strValue1 = arrValue(i)
                    lstSelect(intType).AddItem arrValue(i), i
                     
                     If intType = 0 Then
                        ReDim arrValue1(0)
                        arrValue1(0) = .TextMatrix(.Row, .Col)
                        txtLst.Text = .TextMatrix(.Row, .Col)
                        txtLst.Tag = 1
                     Else
                        arrValue1 = Split(.TextMatrix(.Row, .Col), ",")
                     End If
                     For j = 0 To UBound(arrValue1)
                        If arrValue1(j) = arrValue(i) Then
                            lstSelect(intType).Selected(i) = True
                            blnSelect = True
                        End If
                    Next j
                Next i
                
                If blnSelect = False And strValue1 <> "" And IIf(intType = 0, Trim(txtLst.Text) = "", True) Then
                    For i = 0 To lstSelect(intType).ListCount - 1
                        If lstSelect(intType).List(i) = strValue1 Then
                            lstSelect(intType).Selected(i) = True
                        End If
                    Next i
                End If
                
                If lstSelect(intType).ListIndex >= 0 Then txtLst.Text = "": PicLst.Tag = 0
                
                '�ؼ���ʾ
                If intType = 0 Then '��ѡ��Ŀ�ṩ����ѡ���¼�빦��
                    PicLst.FontName = .FontName
                    PicLst.FontSize = .FontSize
                    PicLst.Left = .CellLeft + .Left + 15
                    PicLst.Top = .CellTop + vsfTab.Top
                    PicLst.Height = 80 + (.CellHeight - 5) + PicLst.TextHeight("��") * 2 + lstSelect(intType).ListCount * (PicLst.TextHeight("��") + PicLst.TextHeight("��") \ 4)
                    If PicLst.Height < .CellHeight + 20 Then PicLst.Height = .CellHeight + 20
                    PicLst.Width = LenB(StrConv(lstSelect(intType).List(lstSelect(intType).ListCount \ 2), vbFromUnicode)) * 120 + 500    '���м���ĳ���Ϊ����
                    If PicLst.Width < .CellWidth + 20 Then PicLst.Width = .CellWidth + 20
                    If PicLst.Height > vsfTab.Height Then PicLst.Height = vsfTab.Height
                    If PicLst.Top + PicLst.Height > vsfTab.Height Then PicLst.Top = .CellTop + .Top + .CellHeight + 20 - PicLst.Height
                    If PicLst.Top < 0 Then PicLst.Top = vsfTab.Top
                    PicLst.Visible = True
                    PicLst.ZOrder 0
                    
                    lbllst(2).Left = 20
                    lbllst(2).Top = 20
                    If lbllst(2).Width > PicLst.Width Then
                        PicLst.Width = lbllst(2).Width + PicLst.TextWidth("��")
                    End If
                    lbllst(2).FontName = .FontName
                    lbllst(2).FontSize = .FontSize
                    lbllst(2).Visible = True
        
                    txtLst.Top = lbllst(2).Top + lbllst(2).Height + 20
                    txtLst.Left = -10
                    txtLst.Width = PicLst.Width
                    txtLst.Height = .CellHeight - 5
                    txtLst.FontName = .FontName
                    txtLst.FontSize = .FontSize
                    strTmp = vsfTab.TextMatrix(vsfTab.Row, COL_tab�ַ���)
                    strTmp = strTmp & String(8 - UBound(Split(strTmp, ",")), ",")
                    arrStr = Split(strTmp, ",")
                    intNum = Val(arrStr(2))
                    intLen = Val(arrStr(6))
                    txtLst.MaxLength = intLen + IIf(intNum = 0, 0, 1)
                    txtLst.Visible = True
                
                    lbllst(3).Left = 20
                    lbllst(3).Top = txtLst.Top + txtLst.Height + 20
                    lbllst(3).FontName = .FontName
                    lbllst(3).FontSize = .FontSize
                    lbllst(3).Visible = True
                    
                    lstSelect(intType).Top = lbllst(3).Top + lbllst(3).Height + 20
                    lstSelect(intType).Left = -10
                    lstSelect(intType).FontName = .FontName
                    lstSelect(intType).FontSize = .FontSize
                    lstSelect(intType).Width = PicLst.Width
                    lstSelect(intType).Height = PicLst.Height - lstSelect(intType).Top
                    lstSelect(intType).Visible = True
                    lstSelect(intType).Enabled = True
                    lstSelect(intType).ZOrder 0
                    lstSelect(intType).Tag = .TextMatrix(.Row, .Col)
                    lbllst(intType).Tag = .Row & "|" & .Col & "|" & "1"
                    txtEdit.Tag = "||1"
                     
                    If lstSelect(intType).Top + lstSelect(intType).Height <> PicLst.Height Then
                        PicLst.Height = lstSelect(intType).Top + lstSelect(intType).Height
                    End If
                    PicLst.SetFocus
                Else
                    lstSelect(intType).Top = .CellTop + vsfTab.Top
                    lstSelect(intType).Left = .CellLeft + .Left + 15
                    lstSelect(intType).FontName = .FontName
                    lstSelect(intType).FontSize = .FontSize
                    lstSelect(intType).Height = lstSelect(intType).ListCount * (PicLst.TextHeight("��") + PicLst.TextHeight("��") \ 4)
                    If lstSelect(intType).Height < .CellHeight + 20 Then lstSelect(intType).Height = .CellHeight + 20
                    lstSelect(intType).Width = LenB(StrConv(lstSelect(intType).List(lstSelect(intType).ListCount \ 2), vbFromUnicode)) * 120 + 500    '���м���ĳ���Ϊ����
                    If lstSelect(intType).Width < .CellWidth + 20 Then lstSelect(intType).Width = .CellWidth + 20
                    If lstSelect(intType).Height > vsfTab.Height Then
                        lstSelect(intType).Height = vsfTab.Height
                    End If
                    If lstSelect(intType).Top + lstSelect(intType).Height > vsfTab.Height Then
                        lstSelect(intType).Top = .CellTop + .Top + .CellHeight + 20 - lstSelect(intType).Height
                    End If
                    If lstSelect(intType).Top < 0 Then lstSelect(intType).Top = vsfTab.Top
                    
                        lstSelect(intType).Visible = True
                        lstSelect(intType).Enabled = True
                        lstSelect(intType).ZOrder 0
                        
                        lstSelect(intType).Tag = .TextMatrix(.Row, .Col)
                        lbllst(intType).Tag = .Row & "|" & .Col & "|1"
                        lstSelect(intType).SetFocus
                    End If
            ElseIf intType = 5 Then 'ѡ��
                lblCheck.Width = picEdit.Width
                lblCheck.Height = picEdit.Height
                lblCheck.Caption = .TextMatrix(vsfTab.Row, vsfTab.Col)
                picEdit.Tag = .TextMatrix(vsfTab.Row, vsfTab.Col) & "'" & .Cell(flexcpForeColor, vsfTab.Row, vsfTab.Col)
                txtEdit.Tag = vsfTab.Row & "|" & vsfTab.Col & "|" & "1"
                lblCheck.Visible = True
                lblCheck.Enabled = True
                lblCheck.ZOrder 0
                picEdit.SetFocus
            Else
                txtEdit.Width = picEdit.Width
ShowText:
                strTmp = vsfTab.TextMatrix(vsfTab.Row, COL_tab�ַ���)
                lngItemNO = Val(vsfTab.TextMatrix(vsfTab.Row, COL_tab��Ŀ���))
                strTmp = strTmp & String(8 - UBound(Split(strTmp, ",")), ",")
                arrStr = Split(strTmp, ",")
                intNum = Val(arrStr(2))
                intLen = Val(arrStr(6))
                
                If intLen <> 0 Then
                    If lngItemNO <> 4 Then
                        txtEdit.MaxLength = intLen + IIf(intNum = 0, 0, 1)
                    Else
                        txtEdit.MaxLength = intLen * 2 + 1 + IIf(intNum = 0, 0, 1) * 2
                    End If
                Else
                    txtEdit.MaxLength = 0
                End If
                
                If vsfTab.Row Mod 2 = 1 Then txtEdit.MaxLength = 5
                
                txtEdit.Text = .TextMatrix(vsfTab.Row, vsfTab.Col)
                picEdit.Tag = .TextMatrix(vsfTab.Row, vsfTab.Col) & "'" & .Cell(flexcpForeColor, vsfTab.Row, vsfTab.Col)
                txtEdit.Tag = vsfTab.Row & "|" & vsfTab.Col & "|" & "1"
                txtEdit.ForeColor = cmdColor.Tag
                txtEdit.Visible = True
                txtEdit.Enabled = True
                txtEdit.ZOrder 0
                picEdit.SetFocus
            
            End If
            
        End If
    End With
ErrFouce:
    If picEdit.Visible = True And txtEdit.Enabled = True Then txtEdit.SetFocus: Call zlControl.TxtSelAll(txtEdit)
ErrInfo:
    If strInfo <> "" Then
        lblStb.Caption = strInfo
        lblStb.ForeColor = 255
    End If
End Sub

Private Function WriteIntoVfgTab(ByVal strText As String, ByVal vsf As Object, Optional blnDelete As Boolean = False, Optional ByVal blnVisible As Boolean = True, Optional strErrMsg As String = "") As Boolean
    '-------------------------------------------------------------------------
    '����:�û��༭������д��vsfTab,������֤
    '����:strtext �༭���ı���Ϣ   blndelete �Ƿ���VsfTab��Delete ��ɾ����Ϣ
    '-------------------------------------------------------------------------
    Dim intRow As Integer
    Dim intCOl As Integer
    Dim str��Ŀ���� As String, strTmp As String, strPart As String
    Dim strֵ�� As String, strHour As String, strHourOld As String
    Dim strValue As String, strTime As String, strOldTime As String
    Dim intType As Integer, intNum As Integer, lngLen As Long, intƵ�� As Integer
    Dim int���� As Integer, int��ʾ As Integer, intIndex As Integer, int��¼���� As Integer
    Dim int״̬ As Integer  '--�����޸���Ϣ
    Dim i As Integer
    Dim lngVsfType As Integer
    Dim arrStr() As String, arrOldTime() As String
    Dim blnAllow As Boolean, blnTrue As Boolean
    Dim BlnTime As Boolean
    Dim lng��Ŀ��� As Long, lngColor As Long
    Dim rsTemp As New ADODB.Recordset
    On Err GoTo Errhand
    
    If Not blnDelete Then
        If picEdit.Visible And txtEdit.Tag <> "" Then
            intRow = Split(txtEdit.Tag, "|")(0)
            intCOl = Split(txtEdit.Tag, "|")(1)
            lngVsfType = Split(txtEdit.Tag, "|")(2)
            If lngVsfType = 1 Then
                If vsfTab.Name <> vsf.Name Then Set vsf = vsfTab
            Else
                If vsfTabDetail.Name <> vsf.Name Then Set vsf = vsfTabDetail
            End If
             '�Ƿ�������ʱ��
            If Val(lngVsfType) = 1 Then
                BlnTime = (intRow Mod 2 = 1)
            Else
                BlnTime = intCOl = vsfTabDetail.FixedCols
            End If
            If txtEdit.Visible = True Or lblCheck.Visible = True Then
                strTmp = vsf.TextMatrix(intRow, COL_tab�ַ���)
                lng��Ŀ��� = Val(vsf.TextMatrix(intRow, COL_tab��Ŀ���))
                str��Ŀ���� = vsf.TextMatrix(intRow, COL_tab��Ŀ��)
                strTmp = strTmp & String(8 - UBound(Split(strTmp, ",")), ",")
                arrStr = Split(strTmp, ",")
                strֵ�� = arrStr(0)
                intType = Val(arrStr(1))
                intNum = Val(arrStr(2))
                intƵ�� = Val(arrStr(3))
                int��ʾ = Val(arrStr(4))
                int���� = Val(arrStr(5))
                lngLen = Val(arrStr(6))
                strPart = arrStr(7)
                
                If intType = 1 Then strֵ�� = ""
                'ȫ����ܣ�����ʾ����ʱ��
                If int��ʾ = 4 And intƵ�� = 1 And mbln¼��Сʱ = True Then
                    If InStr(1, strText, ")") > 0 Then
                        strHour = Replace(Replace(Split(strText, ")")(0), "(", ""), "h", "")
                        If strHour <> "" Then
                            If Not Val(strHour) >= 0 And Val(strHour) <= 24 Then
                                lblStb.Caption = "����Сʱֻ����0��24֮�䣬������¼�룡": lblStb.ForeColor = 255
                                Exit Function
                            End If
                            strHour = "(" & strHour & "h)"
                        End If
                        strText = Split(strText, ")")(1)
                        If Trim(strText) = "" Then strHour = ""
                    End If
                End If
                If txtEdit.Enabled = True Or txtHour.Visible = True Then
                    blnAllow = CheckValidata(intRow, intCOl, lng��Ŀ���, intType, intNum, strֵ��, int��ʾ, lngLen, strText, BlnTime, strErrMsg)
                End If
            End If
            strValue = Split(IIf(Trim(picEdit.Tag) = "", "'", Trim(picEdit.Tag)), "'")(0)
        ElseIf lstSelect(0).Visible = True Or lstSelect(1).Visible = True Then
            If lstSelect(0).Visible = True Then strValue = lstSelect(0).Tag: intIndex = 0
            If lstSelect(1).Visible = True Then strValue = lstSelect(1).Tag: intIndex = 1
            intRow = Split(lbllst(intIndex).Tag, "|")(0)
            intCOl = Split(lbllst(intIndex).Tag, "|")(1)
            lng��Ŀ��� = Val(vsf.TextMatrix(intRow, COL_tab��Ŀ���))
            str��Ŀ���� = vsf.TextMatrix(intRow, COL_tab��Ŀ��)
            strTmp = vsf.TextMatrix(intRow, COL_tab�ַ���)
            strTmp = strTmp & String(8 - UBound(Split(strTmp, ",")), ",")
            arrStr = Split(strTmp, ",")
            intType = Val(arrStr(1))
            int���� = Val(arrStr(5))
            strPart = arrStr(7)
            
            blnAllow = True
        End If
    Else
        blnAllow = True
        If InStr(1, picTab.Tag, "|") = 0 Then Exit Function
        intRow = Split(picTab.Tag, "|")(0)
        intCOl = Split(picTab.Tag, "|")(1)
        lng��Ŀ��� = Val(vsf.TextMatrix(intRow, COL_tab��Ŀ���))
        str��Ŀ���� = vsf.TextMatrix(intRow, COL_tab��Ŀ��)
        strTmp = vsf.TextMatrix(intRow, COL_tab�ַ���)
        strTmp = strTmp & String(8 - UBound(Split(strTmp, ",")), ",")
        arrStr = Split(strTmp, ",")
        intType = Val(arrStr(1))
        intƵ�� = Val(arrStr(3))
        int��ʾ = Val(arrStr(4))
        int���� = Val(arrStr(5))
        strPart = arrStr(7)
        strHour = ""
        strValue = fraTable.Tag
    End If
    
    If blnAllow = True Then
        lngColor = 0
        vsf.TextMatrix(intRow, intCOl) = strHour & strText
        If cmdColor.Visible = True Then lngColor = cmdColor.Tag
        vsf.Cell(flexcpForeColor, intRow, intCOl, intRow, intCOl) = lngColor
        If vsf.Name = vsfTabDetail.Name Then
            vsf.Cell(flexcpForeColor, intRow, intCOl - 1, intRow, intCOl + 1) = lngColor
        End If
        mblnEdit = True
    Else
        If strErrMsg <> "" Then GoTo ErrInfo
        Exit Function
    End If
    
    mrsTable.Filter = ""
    int��¼���� = 1
    blnTrue = False
    '���������޸ı�־
    If blnAllow = True Then
        strHour = Replace(Replace(strHour, "(", ""), "h)", "")
        If int��ʾ = 4 And intƵ�� = 1 And mbln¼��Сʱ = True Then
            If InStr(1, strValue, ")") > 0 Then
                strHourOld = Replace(Replace(Split(strValue, ")")(0), "(", ""), "h", "")
                strValue = Split(strValue, ")")(1)
            End If
            '�����û�¼��Ļ���Сʱ��ֻ���޸û���ʱ��
            If Val(strHour) <> Val(strHourOld) Then
                blnTrue = True
                int��¼���� = 11
                GoTo DataUpdate
            End If
        End If
DataUpdate:
        If vsf.Name = "vsfTab" Then
            If InStr(vsf.TextMatrix(intRow, col_tabԭʼʱ��), "'") > 0 Then
                arrOldTime = Split(vsf.TextMatrix(intRow, col_tabԭʼʱ��), "'")
                strTime = Format(arrOldTime(intCOl - vsfTab.FixedCols), "YYYY-MM-DD hh:mm:ss")
            Else
                ReDim Preserve arrOldTime(0)
                strTime = vsf.TextMatrix(intRow, col_tabԭʼʱ��)
            End If
        Else
            strTime = Format(vsf.TextMatrix(intRow, col_tabԭʼʱ��), "YYYY-MM-DD hh:mm:ss")
        End If
        If BlnTime Then
            strText = Format(IIf(Split(vsfTab.RowData(vsfTab.Row), ";")(1) = 0 Or mbln���ܵ���, dtpDate.Value & " " & strText, dtpDate.Value - 1 & " " & strText), "YYYY-MM-DD hh:mm:ss")
        Else
            If vsf.Name = vsfTab.Name Then
                If vsf.TextMatrix(intRow - 1, intCOl) <> "" Then
                    If strTime = "" Then strTime = Format(dtpDate.Value & " " & vsf.TextMatrix(intRow - 1, intCOl), "YYYY-MM-DD hh:mm:ss")
                End If
            Else
                If vsf.TextMatrix(intRow, vsf.FixedCols) <> "" Then
                    If strTime = "" Then strTime = Format(dtpDate.Value & " " & vsf.TextMatrix(intRow - 1, intCOl), "YYYY-MM-DD hh:mm:ss")
                End If
            End If
            If strTime = "" Then
                If lngVsfType = 1 Then
                    strTime = GetAnimalItemTime(intRow, intCOl - vsfTab.FixedCols + 1, 1, strErrMsg)
                Else
                    strTime = GetAnimalItemTime(vsfTab.Row, vsfTab.Col - vsfTab.FixedCols + 1, 1, strErrMsg)
                End If
                If strErrMsg <> "" Then GoTo ErrInfo
                If IsExistData(strTime, lng��Ŀ���) = False Then
                    strErrMsg = lblStb.Caption
                    GoTo ErrInfo
                End If
            End If
        End If
        
        mrsTableDetail.Filter = "��Ŀ���=" & lng��Ŀ��� & " and ��Ŀ����='" & str��Ŀ���� & "' And ��¼����=" & int��¼���� & " and ʱ��='" & strTime & "'"
        If BlnTime Then strTime = strText: strText = ""
        If mrsTableDetail.RecordCount > 0 Then
            mrsTableDetail!δ��˵�� = lngColor
            If mrsTableDetail!״̬ <> 1 Then 'ԭ�е����� �޸ġ�ɾ�����״̬
                If BlnTime And mrsTableDetail!״̬ = 0 Then
                    mrsTableDetail!״̬ = 3
                Else
                    mrsTableDetail!״̬ = 2
                End If
                If BlnTime Then
                    mrsTableDetail!ʱ�� = strTime
                Else
                    mrsTableDetail!��� = IIf(blnTrue = True, strHour, strText)
                End If
                If strText = "" And mrsTableDetail!��� = "" Then
                    mrsTableDetail!״̬ = 4 'ɾ��
                End If
            Else
                If Trim(IIf(blnTrue = True, strHour, strText)) = "" Then '����ɾ��
                        mrsTableDetail.Delete
                    Else '����
                        mrsTableDetail!״̬ = 1
                        If BlnTime Then
                            mrsTableDetail!ʱ�� = strTime
                        Else
                            mrsTableDetail!��� = IIf(blnTrue = True, strHour, strText)
                        End If
                    End If
            End If
            mrsTableDetail.Update
        Else '����
            If Trim(strText) <> "" Then
                If strErrMsg <> "" Then GoTo ErrInfo
            End If
            strText = Replace(Replace(strText, "|", "�O"), "'", "")
            gstrFields = "id|������|���|���²�λ|���|ʱ��|��Ŀ���|��Ŀ����|���Ժϸ�|δ��˵��|������Դ|�޸�|��ʾ|��ԴID|����|״̬|�к�|��¼����"
            If BlnTime Then
                gstrValues = GetMaxID(mrsTableDetail) & "|2)���±����Ŀ|" & "" & "|" & strPart & "|" & _
                0 & "|" & strTime & "|" & lng��Ŀ��� & "|" & str��Ŀ���� & "|0|" & lngColor & "|0|0|0|0|0|1|" & vsfTab.Col - vsfTab.FixedCols + 1 & "|" & int��¼����
            Else
                gstrValues = GetMaxID(mrsTableDetail) & "|2)���±����Ŀ|" & IIf(blnTrue = True, strHour, strText) & "|" & strPart & "|" & _
                0 & "|" & strTime & "|" & lng��Ŀ��� & "|" & str��Ŀ���� & "|0|" & lngColor & "|0|0|0|0|0|1|" & vsfTab.Col - vsfTab.FixedCols + 1 & "|" & int��¼����
            End If
            Call Record_Add(mrsTableDetail, gstrFields, gstrValues)
            If lngVsfType = 1 Then
                arrOldTime(intCOl - vsfTab.FixedCols) = strTime
                vsf.TextMatrix(intRow, col_tabԭʼʱ��) = Join(arrOldTime, "'")
            Else
                vsf.TextMatrix(intRow, col_tabԭʼʱ��) = strTime
            End If
        End If
    End If
    mrsTableDetail.Filter = "״̬<> 4 "
    
    gstrFields = "ID," & adDouble & ",18|������," & adLongVarChar & ",40|���," & adLongVarChar & ",400|���²�λ," & adLongVarChar & ",200|" & _
         "���," & adDouble & ",1|ʱ��," & adLongVarChar & ",20|ԭʼʱ��," & adLongVarChar & ",20|��Ŀ���," & adDouble & ",18|��Ŀ����," & adLongVarChar & ",40|" & _
         "���Ժϸ�," & adDouble & ",1|δ��˵��," & adLongVarChar & ",20|������Դ," & adDouble & ",1|�޸�," & adDouble & ",1|��ʾ," & adDouble & ",1|ԭʼ��ʾ״̬," & adDouble & ",1|" & _
         "��ԴID," & adDouble & ",18|����," & adDouble & ",1|״̬," & adDouble & ",1|�к�," & adDouble & ",1|��¼����," & adDouble & ",1"
    Call Record_Init(rsTemp, gstrFields)
    
    Do While Not mrsTableDetail.EOF
        rsTemp.AddNew
        For i = 0 To mrsTableDetail.Fields.Count - 1
            rsTemp.Fields(mrsTableDetail.Fields(i).Name).Value = mrsTableDetail.Fields(i).Value
        Next i
        rsTemp.Update
        mrsTableDetail.MoveNext
    Loop
    
    Call InitTableData(rsTemp)
    Call ShowTable
    If blnAllow = True And blnVisible = True Then Call txtEdit_KeyPress(vbKeyEscape): mblnEdit = True
'    Call SetColSelect
    WriteIntoVfgTab = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Exit Function
ErrInfo:
    If strErrMsg <> "" Then
        lblStb.Caption = strErrMsg: lblStb.ForeColor = 255
        vsf.TextMatrix(intRow, intCOl) = strValue
    End If
End Function

Private Function CheckValidata(ByVal intRow As Integer, ByVal intCOl As Integer, ByVal lngNo As Long, ByVal intType As Integer, ByVal intС�� As Integer, ByVal strֵ�� As String, _
    ByVal int��ʾ As Integer, ByVal lngLen As Long, strInfo As String, ByVal BlnTime As Boolean, Optional strErrMsg As String = "") As Boolean
'-------------------------------------------------------------
'���ܣ�������ݺϷ��ԣ�������ݣ�
'����:introw����һ�� intCol�� ��һ��  lngNo:��Ŀ��� intype�� ��Ŀ���� 0�������� 1 �������� strֵ����Ŀֵ��
'   lngLen����Ŀ����  strInfo��ҪУ����ı�ֵ
'-------------------------------------------------------------
    Dim strName As String, strTmp As String
    Dim strTime As String
    Dim strMsg As String, strText As String
    Dim lngRow As Long, lng��Ŀ��� As Long
    Dim blnAllow As Boolean '�Ƿ��Ǵ���������Һ��
    Dim blnOK As Boolean
    Dim i As Integer
    Dim intƵ�� As Integer
    Dim arrValue() As String
    
    On Error GoTo Errhand
    strName = vsfTab.TextMatrix(vsfTab.Row, COL_tab��Ŀ��)
    lng��Ŀ��� = vsfTab.TextMatrix(vsfTab.Row, COL_tab��Ŀ���)
    lngRow = intRow - vsfTab.FixedRows + 1
    
    If strInfo = "" Then
        CheckValidata = True
        Exit Function
    End If
    If BlnTime Then
        If Len(strInfo) < 3 Or (Len(strInfo) = 3 And InStr(strInfo, ":") > 0) Or (Len(strInfo) = 5 And InStr(strInfo, ":") <= 0) Or Len(strInfo) > 5 Then
            strMsg = "��" & lngRow & "��[" & strName & "]��ʱ���ʽ¼�벻��ȷ��ӦΪ Сʱ������"
            GoTo ErrInfo
        End If
        strInfo = CToData(strInfo)
        
        If InStr(1, strInfo, ":") > 0 Then
            If Not IsDate(strInfo) Then
                 strMsg = "��" & lngRow & "��[" & strName & "]��ʱ��¼�벻��ȷ��ӦΪ Сʱ������"
                GoTo ErrInfo
            End If
            intƵ�� = Split(vsfTab.RowData(vsfTab.Row), ";")(0)
            '���¼����Ŀʱ���Ƿ񳬳��û����õ�ʱ�䷶Χ���ǲ�¼��Χ
            strTime = GetAnimalItemTime(vsfTab.Row, vsfTab.Col - vsfTab.FixedCols + 1, 0, strMsg)
            If strMsg <> "" Then GoTo ErrInfo
            If strInfo <> "" Then
                If (Format(Split(strTime, ";")(0), "hh:mm:ss") < Format(Split(strTime, ";")(1), "hh:mm:ss") And Split(vsfTab.RowData(vsfTab.Row), ";")(1) <> 0) Or (intƵ�� = 1 And Split(vsfTab.RowData(vsfTab.Row), ";")(1) = 3) Then
                    If CDate(IIf(mbln���ܵ���, dtpDate.Value, dtpDate.Value - 1) & " " & strInfo) < CDate(Split(strTime, ";")(0)) Then strMsg = "¼��ʱ��С�ڵ�������¼��ʱ��Σ�" & Split(strTime, ";")(0) & "��" & Split(strTime, ";")(1): GoTo ErrInfo
                    If CDate(IIf(mbln���ܵ���, dtpDate.Value, dtpDate.Value - 1) & " " & strInfo) > CDate(Split(strTime, ";")(1)) Then strMsg = "¼��ʱ����ڵ�������¼��ʱ��Σ�" & Split(strTime, ";")(0) & "��" & Split(strTime, ";")(1): GoTo ErrInfo
                Else
                    If CDate(dtpDate.Value & " " & strInfo) < CDate(Split(strTime, ";")(0)) Then strMsg = "¼��ʱ��С�ڵ�������¼��ʱ��Σ�" & Split(strTime, ";")(0) & "��" & Split(strTime, ";")(1): GoTo ErrInfo
                    If CDate(dtpDate.Value & " " & strInfo) > CDate(Split(strTime, ";")(1)) Then strMsg = "¼��ʱ����ڵ�������¼��ʱ��Σ�" & Split(strTime, ";")(0) & "��" & Split(strTime, ";")(1): GoTo ErrInfo
                End If
            End If
            
            If Format(IIf(mbln���ܵ���, dtpDate.Value, dtpDate.Value - 1) & " " & strInfo, "YYYY-MM-DD hh:mm:ss") < mstrBTime Then
                strMsg = "��" & lngRow & "��[" & strName & "]��ʱ��С�����µ���ʼʱ��,����!"
                GoTo ErrInfo
            End If
            If Format(IIf(mbln���ܵ���, dtpDate.Value, dtpDate.Value - 1) & " " & strInfo, "YYYY-MM-DD hh:mm:ss") > mstrETime Then
                strMsg = "��" & lngRow & "��[" & strName & "]��ʱ��������µ���¼ʱ��,����!"
                GoTo ErrInfo
            End If
            If IsExistData(Format(IIf(mbln���ܵ���, dtpDate.Value, dtpDate.Value - 1) & " " & strInfo, "YYYY-MM-DD hh:mm:ss"), lng��Ŀ���) = False Then
                strErrMsg = lblStb.Caption
                GoTo ErrInfo
            End If
        End If
    Else
        blnAllow = True
        If strName = "����" Or strName = "���" Then
            If IsNumeric(strInfo) Then
                blnAllow = True
            Else
                blnAllow = False
            End If
        End If
        '����������ҹ�����
        If blnAllow = True Then blnAllow = IIf(InStr(1, "," & gint��� & "," & gint��Һ & ",", "," & lngNo & ",") > 0, False, True)
        If Not (intType = 0 And InStr(1, "0,4", int��ʾ) <> 0) Then
            If LenB(StrConv(strInfo, vbFromUnicode)) > lngLen Then
                strMsg = "��" & lngRow & "��[" & strName & "]��ֵ����(��󳤶�:" & lngLen & "),����!"
                GoTo ErrInfo
            End If
        Else
            If intType = 0 Then
                If int��ʾ = 4 Or strֵ�� = "" Then
                    strֵ�� = "0��" & IIf(lngLen - intС�� > 0, String(lngLen - intС��, "9"), "0") & IIf(intС�� > 0, "." & String(intС��, "9"), "")
                End If
                If lngNo <> 4 And lngNo <> 5 And blnAllow = True Then
                    If Not IsNumeric(strInfo) Then
                        strMsg = strName & "����¼�����" & Space(4) & "��Ч��Χ:" & strֵ��
                        GoTo ErrInfo
                    End If
                End If
                    
                If lngNo = 4 And strName = "Ѫѹ" Then
                'Ѫѹ����¼������˵���������δ���
                    mrsCurInfo.Filter = "����='" & strInfo & "'"
                    If Not mrsCurInfo.EOF Then
                        CheckValidata = True
                        Exit Function
                    Else
                        strTmp = ""
                        mrsCurInfo.Filter = "": mrsCurInfo.Sort = "����"
                        Do While Not mrsCurInfo.EOF
                            strTmp = strTmp & "��" & Nvl(mrsCurInfo!����)
                            mrsCurInfo.MoveNext
                        Loop
                        strTmp = Mid(strTmp, 2)
                
                        If InStr(1, strInfo, "/") = 0 Then
                            strMsg = "��" & lngRow & "��[Ѫѹ]���ݵĸ�ʽ��������ѹ/����ѹ" & IIf(strTmp <> "", "��(" & strTmp & ")", "") & "��"
                            GoTo ErrInfo
                        End If
                        If Trim(Split(strInfo, "/")(0)) = "" Or Trim(Split(strInfo, "/")(1)) = "" Then
                            strMsg = "��" & lngRow & "��[Ѫѹ]����¼���������ѹ/����ѹ" & IIf(strTmp <> "", "��(" & strTmp & ")", "") & "��"
                            GoTo ErrInfo
                        End If
                    End If
                End If
                
                If UBound(Split(strInfo, "/")) > 1 And blnAllow = True Then
                    strMsg = "��" & lngRow & "��[" & strName & "]����¼��������飡"
                    GoTo ErrInfo
                End If
                
                '�����������Ч��Χ���Ƿ���Ч
                arrValue = Split(strInfo, "/")
                For i = 0 To UBound(arrValue)
                    blnOK = False
                    strText = arrValue(i)
                    If Not blnOK Then
                        If Not IsNumeric(strText) And blnAllow = True Then
                            strMsg = "��" & lngRow & "��[" & strName & "]����¼�����" & Space(4) & "��Ч��Χ:" & strֵ��
                            GoTo ErrInfo
                        End If
                    End If
                        
                    If Not blnOK And strText <> "" And blnAllow = True Then
                        strText = Format(Val(strText), "#0" & IIf(intС�� > 0, ".", "") & String(intС��, "0"))
                        '0.30תΪ0.3
                        If strText = Val(strText) Then strText = Val(strText)
                        If Left(strText, 1) = "." Then strText = 0 & strText
                    End If
                    
                    If int��ʾ <> 4 And blnAllow = True Then
                        If Len(Replace(strText, ".", "")) > lngLen Then
                            strMsg = "��" & lngRow & "��[" & strName & "]��ֵ����(��󳤶�:" & lngLen & "),����!"
                            GoTo ErrInfo
                        End If
                    End If
                    
                    If IsNumeric(Split(strֵ��, "��")(0)) And IsNumeric(strText) Then
                        If blnAllow = True Then   '��������������Ч��Χ���
                            If Not (Val(strText) >= Split(strֵ��, "��")(0) And Val(strText) <= Split(strֵ��, "��")(1)) Then
                                strMsg = strName & "������Ч��Χ(" & strֵ�� & "),����!"
                                GoTo ErrInfo
                            End If
                        End If
                    End If
                    arrValue(i) = strText
                Next
                strInfo = Join(arrValue, "/")
            End If
        End If
    End If
    
    CheckValidata = True
    Exit Function
    
ErrInfo:
    strErrMsg = strMsg
    
Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Function IsExistData(ByVal strTime As String, ByVal lngNo As Long) As Boolean
    '-----------------------------------------------
    '��鵱ǰʱ���Ƿ��Ѵ�������
    '-----------------------------------------------
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo Errhand
    
    '����޸ĵ�ʱ���Ƿ��Ѿ���������
    If mrsCurve.State = adStateOpen Then
        mrsCurve.Filter = "ʱ��= '" & strTime & "' And ��Ŀ��� = " & lngNo
        If mrsCurve.RecordCount > 0 Then lblStb.Caption = "��ǰʱ���Ѿ���������,����������ʱ��.": lblStb.ForeColor = 255: Exit Function
    End If
    If mrsTableDetail.State = adStateOpen Then
        mrsTableDetail.Filter = "ʱ��= '" & strTime & "' and ��Ŀ���=" & lngNo
        If mrsTableDetail.RecordCount > 0 Then
            If mrsTableDetail!��� <> "" Then lblStb.Caption = "��ǰʱ���Ѿ���������,����������ʱ��.": lblStb.ForeColor = 255: Exit Function
        End If
    End If
    strSQL = "select 1 From ���˻����ļ� a,���˻������� b,���˻�����ϸ c" & vbNewLine & _
        " where A.ID=B.�ļ�ID and b.id =c.��¼id and A.ID=[1] and A.����ID=[2] and A.��ҳID=[3] And nvl(A.Ӥ��,0)=[4]" & vbNewLine & _
        " and B.����ʱ��=[5] and c.��Ŀ���=[6]"
        
    If mblnMove Then
        mstrSQL = Replace(mstrSQL, "���˻����ļ�", "H���˻����ļ�")
        mstrSQL = Replace(mstrSQL, "���˻�������", "H���˻�������")
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���ʱ��", mT_Patient.lng�ļ�ID, mT_Patient.lng����ID, mT_Patient.lng��ҳID, mT_Patient.lngӤ��, CDate(strTime), lngNo)
    
    If rsTemp.RecordCount > 0 Then
        lblStb.Caption = "��ǰʱ���Ѿ���������,����������ʱ��."
        lblStb.ForeColor = 255
        Exit Function
    End If
    
    IsExistData = True
       
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub vsfTab_GotFocus()
    If vsfTab.Tag <> "NO" Then vsfTab.Tag = "1"
End Sub

Private Sub vsfTab_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intCols As Integer
    Dim intType As Integer, intƵ�� As Integer
    Dim blnTrue As Boolean
    Dim blnEdit As Boolean
    Dim strText As String
    
    If vsfTab.Tag = "NO" Then Exit Sub
    If vsfTab.Row < vsfTab.FixedRows And vsfTab.Col < vsfTab.FixedCols Then Exit Sub
    
    '���ε�ĳЩ���ܼ�
    If KeyCode = vbKeyEscape Or KeyCode = vbKeyBack Or Shift <> 0 Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyRight Then Exit Sub
    
    If KeyCode = vbKeyLeft And (picEdit.Visible = False And lstSelect(0).Visible = False And lstSelect(1).Visible = False) Then Exit Sub
    
    intCols = Val(Split(vsfTab.TextMatrix(vsfTab.Row, COL_tab�ַ���), ",")(3)) + vsfTab.FixedCols
    
    With vsfTab
        If KeyCode = vbKeyReturn Then
NextCol2: '������һ��
            If .Col < vsfTab.FixedCols Then
                .Col = .Col + 1: GoTo NextCol2
            End If
            If .Col < intCols - 1 Then
                If vsfTab.Row Mod 2 = 1 Then
                    GoTo NextRow2
                Else
                    If .Row > .FixedRows Then .Row = .Row - 1
                    .Col = .Col + 1
                    If .ColHidden(.Col) = True Then GoTo NextCol2
                End If
            Else
NextRow2: '������һ��
                If .Row < .Rows - 1 Then
                    .Row = .Row + 1
                    If .Col + 1 = intCols And .Row Mod 2 = 1 Then .Col = vsfTab.FixedCols
                    If .RowHidden(.Row) = True Then GoTo NextRow2
                Else
                    txtEdit.Tag = "||1"
                    Call txtEdit_KeyPress(vbKeyEscape)
                    .Row = .FixedRows
                    .Col = .FixedCols
                End If
            End If
            '������л��в��ɼ����Զ���ʾ����
            If .ColIsVisible(.Col) = False Then
                .LeftCol = .Col
            End If
            If .RowIsVisible(.Row) = False Then
                .TopRow = .Row
            End If
        
            Exit Sub
        End If
        '���
        If KeyCode = vbKeyLeft Then
PreCol2:
            If .Col > vsfTab.FixedCols Then
                .Col = .Col - 1
                If .ColHidden(.Col) = True Then GoTo PreCol2
            Else
PreRow2:
                If .Row > vsfTab.FixedRows Then
                    .Row = .Row - 1
                    If .RowHidden(.Row) Then GoTo PreRow2
                    .Col = Val(Split(vsfTab.TextMatrix(vsfTab.Row, COL_tab�ַ���), ",")(3)) + vsfTab.FixedCols
                    GoTo PreCol2
                End If
            End If
            '������л��в��ɼ����Զ���ʾ����
            If .ColIsVisible(.Col) = False Then
                .LeftCol = .Col
            End If
            If .RowIsVisible(.Row) = False Then
                .TopRow = .Row
            End If
            Exit Sub
        End If
        
        'ɾ����Ϣ
        If KeyCode = vbKeyDelete Then
            If Shift = 0 And .Col > .FixedCols - 1 And .Col < intCols Then
                blnEdit = True
                If .TextMatrix(.Row, .Col) <> "" Then
                    '�����Ŀ�Ƿ��ǲ�����Ŀ
                    If IsWaveItem(Val(.TextMatrix(.Row, COL_tab��Ŀ���))) And InStr(1, Trim(.TextMatrix(.Row, .Col)), "-") <> 0 Then
                        lblStb.Caption = "������ֵ�Ѿ��γɲ�����Χ�Ĳ�����Ŀ���ܽ����޸ġ�ɾ������"
                        lblStb.ForeColor = 255
                        GoTo ErrExit
                    End If
                    '���������Դ�Ƿ����Ի����¼����PDA
                    mrsCurve.Filter = "��Ŀ���=" & Val(.TextMatrix(.Row, COL_tab��Ŀ���)) & " and ��Ŀ����='" & .TextMatrix(.Row, COL_tab��Ŀ��) & "'" & _
                        "   And �к�=" & .Col - .FixedCols + 1
                    If mrsCurve.RecordCount > 0 Then
                        If InStr(1, ",0,3,9,", "," & Val(mrsCurve!������Դ) & ",") = 0 Then
                            blnEdit = False
                        End If
                    End If
                    intƵ�� = Val(Split(vsfTab.TextMatrix(vsfTab.Row, COL_tab�ַ���), ",")(3))
                    intType = Val(Split(vsfTab.TextMatrix(vsfTab.Row, COL_tab�ַ���), ",")(4))
                    If blnEdit = False And Not (intType = 4 And intƵ�� = 1 And mbln¼��Сʱ = True) Then
                        lblStb.Caption = "������Դ�ڻ����¼����PDA�����ݲ��ܽ����޸ġ�ɾ������"
                        lblStb.ForeColor = 255
                        GoTo ErrExit
                    End If
                    picTab.Tag = .Row & "|" & .Col
                    fraTable.Tag = .TextMatrix(.Row, .Col)
                    strText = ""
                    If blnEdit = False Then '������ȫ�������Ŀ������mbln¼��Сʱ=true
                        If InStr(1, .TextMatrix(.Row, .Col), ")") > 0 Then
                            strText = Split(.TextMatrix(.Row, .Col), ")")(1)
                        Else
                            GoTo ErrExit
                        End If
                    End If
                    blnTrue = WriteIntoVfgTab(strText, vsfTab, True)
                End If
            End If
ErrExit:
            mblnEdit = False
            Exit Sub
        End If
        mblnEdit = True
        Call vsfTab_EnterCell
    End With
End Sub

Private Sub vsfTab_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    vsfTab.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignLeftCenter
    If mblnFileBack = True Then
        Cancel = True
        vsfTab.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignCenterCenter
        lblStb.Caption = "�������������Ѿ��鵵,��������������޸�."
        lblStb.ForeColor = 255
        Exit Sub
    End If
    If vsfTab.Tag = "NO" Then Cancel = True
End Sub

Private Sub vsfTabDetail_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strName As String, strTmp As String, strֵ�� As String
    Dim strInfo As String
    Dim lngNo As Long
    Dim arrStr() As String
    
    If mblnInit = False Then Exit Sub
    If NewRow < vsfTabDetail.FixedRows Or NewCol < vsfTabDetail.FixedCols Or NewRow > Val(vsfTabDetail.Tag) Then Exit Sub
    Call AdjustRowFlag(vsfTabDetail, NewRow)
    With vsfTabDetail
        lngNo = Val(vsfTabDetail.TextMatrix(NewRow, COL_tab��Ŀ���))
        strName = .TextMatrix(NewRow, COL_tab��Ŀ����)
        strTmp = .TextMatrix(NewRow, COL_tab�ַ���)
        strTmp = strTmp & String(8 - UBound(Split(strTmp, ",")), ",")
        arrStr = Split(strTmp, ",")
        strֵ�� = arrStr(0)
        
        If strֵ�� = "" Then
            strInfo = ""
        Else
            strInfo = strName & "��Ч��Χ:" & strֵ��
        End If
        
        If lngNo = 4 And strName = "Ѫѹ" Then 'Ѫѹ
            strInfo = strInfo & Space(4) & "¼�����:����ѹ/����ѹ"
            mrsCurInfo.Filter = ""
            mrsCurInfo.Sort = "����"
            strTmp = ""
            Do While Not mrsCurInfo.EOF
                strTmp = strTmp & "��" & Nvl(mrsCurInfo!����)
                mrsCurInfo.MoveNext
            Loop
            strTmp = Mid(strTmp, 2)
            If strTmp <> "" Then strInfo = strInfo & "��(" & strTmp & ")"
        End If
        
        If Val(arrStr(4)) = 4 Then strInfo = strInfo & Space(4) & "������Ŀ" & Space(4) & "¼�����:����¼��" & IIf(mbln���ܵ��� = True, "����", "����") & "�����ݡ�"
        
        
    End With
    lblStb.Caption = strInfo
    lblStb.ForeColor = &H80000012
    
    mrsCurve.Filter = "��Ŀ���=" & lngNo & " and ��Ŀ����='" & strName & "'" & _
        "   and �к�=" & NewCol - vsfTab.FixedCols + 1
    If mrsCurve.RecordCount > 0 Then
        If InStr(1, ",0,3,9,", "," & Val(mrsCurve!������Դ) & ",") = 0 Then
            lblStb.Caption = "������Դ�ڻ����¼����PDA�����ݲ��ܽ����޸ġ�ɾ������"
            lblStb.ForeColor = 255
            Exit Sub
        End If
    End If
    
End Sub

Private Sub vsfTabDetail_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    mblnScroll = True
    Call vsfTabDetail_EnterCell
    mblnScroll = False
End Sub

Private Sub vsfTabDetail_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
    With vsfTabDetail
        If vsfTab.Tag = "NO" Then Exit Sub
        If .Row > .FixedRows - 1 And .Col > .FixedCols - 1 And .FocusRect = flexFocusSolid Then
            mblnEdit = True
            Call vsfTabDetail_EnterCell
        End If
    End With
End Sub

Private Sub vsfTabDetail_DblClick()
    With vsfTabDetail
        If vsfTab.Tag = "NO" Then Exit Sub
        If .Row > .FixedRows - 1 And .Col > .FixedCols - 1 And .FocusRect = flexFocusSolid And .Col < .FixedCols + 2 Then
            mblnEdit = True
            Call vsfTabDetail_EnterCell
        End If
    End With
End Sub

Private Sub vsfTabDetail_EnterCell()
    Dim strInfo As String
    Dim strValue As String, strValue1 As String
    Dim strTmp As String, strData As String
    Dim strTime As String
    Dim blnAllow As Boolean
    Dim blnEdit As Boolean
    Dim blnSelect As Boolean
    Dim intType As Integer
    Dim intƵ�� As Integer, int��Ŀ���� As Integer, int��Ŀ���� As Integer
    Dim i As Integer, j As Integer
    Dim intNum As Integer, intLen As Integer, intRow As Integer, intCOl As Integer
    Dim lngItemNO As Long, lngColor As Long
    Dim arrValue() As String, arrValue1() As String
    Dim arrStr() As String
    
    If Not mblnInit Then Exit Sub
    If vsfTab.Tag = "NO" Then Exit Sub
    If vsfTabDetail.Rows = vsfTabDetail.FixedRows Then Exit Sub
    blnAllow = True
    blnEdit = True
    blnSelect = True
    
    If picEdit.Visible = True And txtEdit.Tag <> "" Then
        intRow = Split(txtEdit.Tag, "|")(0)
        intCOl = Split(txtEdit.Tag, "|")(1)
        If Split(txtEdit.Tag, "|")(2) <> 2 Then Call txtEdit_KeyPress(vbKeyEscape): Exit Sub
        If txtEdit.Visible = True Then
            strData = IIf(picHour.Visible = True, "(" & txtHour.Text & "h)", "") & Trim(txtEdit.Text)
            lngColor = txtEdit.ForeColor
        Else
            strData = Trim(lblCheck.Caption)
            lngColor = 0
        End If
        
        If Split(txtEdit.Tag, "|")(2) = 1 Then mblnEdit = False
        If intCOl > vsfTabDetail.FixedCols + 1 Then mblnEdit = False
        
        If IIf(cmdColor.Visible, strData & "'" & lngColor <> picEdit.Tag, strData <> Split(picEdit.Tag, "'")(0)) Then blnAllow = WriteIntoVfgTab(strData, vsfTabDetail, False, True, strInfo)
        If cmdColor.Visible = True And mblnEdit = True Then vsfTabDetail.Cell(flexcpForeColor, intRow, intCOl - 1, intRow, intCOl + 1) = Val(cmdColor.Tag)
    End If
    
    '���ݲ��Ϸ�
    If blnAllow = False Then
        If vsfTabDetail.Row <> intRow Then vsfTabDetail.Row = intRow
        If vsfTabDetail.Col <> intCOl Then vsfTabDetail.Col = intCOl
        GoTo ErrFouce
        Exit Sub
    End If
    
    If vsfTabDetail.Row < vsfTabDetail.FixedRows And vsfTabDetail.Col < vsfTabDetail.FixedCols Then Exit Sub
    If Not vsfTabDetail.RowIsVisible(vsfTabDetail.Row) Then Exit Sub
    If Not mblnScroll And vsfTabDetail.Visible Then vsfTabDetail.SetFocus

    '�������б༭�ؼ�
    picδ��.Visible = False
    picEdit.Visible = False
    picEdit.Tag = ""
    txtEdit.Tag = "": txtEdit.Visible = False: txtEdit.Enabled = False
    picHour.Visible = False: picHour.Enabled = False
    txtHour.Tag = "": txtHour.Visible = False: txtHour.Enabled = False
    lblCheck.Visible = False: lblCheck.Enabled = False
    cmdColor.Visible = False
    cmdColor.Enabled = False
    cmdColor.Tag = 0
    picColor.Visible = False
    PicLst.Visible = False
    PicLst.Tag = ""
    txtLst.Visible = False: txtLst.Text = ""
    lstSelect(0).Visible = False
    lstSelect(0).Enabled = False
    lstSelect(0).Tag = ""
    lstSelect(1).Visible = False
    lstSelect(1).Enabled = False
    lstSelect(1).Tag = ""
        
    If mblnFileBack = True Then
        lblStb.Caption = "�������������Ѿ��鵵,��������������޸�."
        mblnEdit = False
        GoTo ErrInfo
    End If
    If vsfTabDetail.Row > Val(vsfTabDetail.Tag) And vsfTabDetail.Tag <> "" Then mblnEdit = False
    If mblnEdit = False Then Exit Sub
    If Not (vsfTabDetail.Row > vsfTabDetail.FixedRows - 1 And vsfTabDetail.Col > vsfTabDetail.FixedCols - 1) Then Exit Sub
    With vsfTabDetail
        If .Row > .FixedRows - 1 And .Col > .FixedCols - 1 And .Col < .FixedCols + 2 Then
            intType = Val(Split(.TextMatrix(.Row, COL_tab�ַ���), ",")(4))
            intƵ�� = Val(Split(.TextMatrix(.Row, COL_tab�ַ���), ",")(3))
            int��Ŀ���� = Val(Split(.TextMatrix(.Row, COL_tab�ַ���), ",")(1))
            int��Ŀ���� = Val(Split(.TextMatrix(.Row, COL_tab�ַ���), ",")(5))
    
            '���¼����Ŀʱ���Ƿ񳬳��û����õ�ʱ�䷶Χ���ǲ�¼��Χ
            If vsfTab.Col - vsfTab.FixedCols + 1 > Split(vsfTab.RowData(vsfTab.Row), ";")(0) Then
                strTime = GetAnimalItemTime(vsfTab.Row, Split(vsfTab.RowData(vsfTab.Row), ";")(0), 0, strInfo)
            Else
                strTime = GetAnimalItemTime(vsfTab.Row, vsfTab.Col - vsfTab.FixedCols + 1, 0, strInfo)
            End If
            If .Col = .FixedCols And .TextMatrix(.Row, .Col) <> "" Then
                If (Format(Split(strTime, ";")(0), "hh:mm:ss") < Format(Split(strTime, ";")(1), "hh:mm:ss") And Split(vsfTab.RowData(vsfTab.Row), ";")(1) <> 0) Or (intƵ�� = 1 And Split(vsfTab.RowData(vsfTab.Row), ";")(1) = 3) Then
                    If CDate(IIf(mbln���ܵ���, dtpDate.Value, dtpDate.Value - 1) & " " & .TextMatrix(.Row, .Col)) < CDate(Split(strTime, ";")(0)) Then strInfo = "¼��ʱ��С�ڱ�������¼��ʱ��Σ�" & Split(strTime, ";")(0) & "��" & Split(strTime, ";")(1)
                    If CDate(IIf(mbln���ܵ���, dtpDate.Value, dtpDate.Value - 1) & " " & .TextMatrix(.Row, .Col)) > CDate(Split(strTime, ";")(1)) Then strInfo = "¼��ʱ����ڱ�������¼��ʱ��Σ�" & Split(strTime, ";")(0) & "��" & Split(strTime, ";")(1)
                Else
                    If CDate(dtpDate.Value & " " & .TextMatrix(.Row, .Col)) < CDate(Split(strTime, ";")(0)) Then strInfo = "¼��ʱ��С�ڱ�������¼��ʱ��Σ�" & Split(strTime, ";")(0) & "��" & Split(strTime, ";")(1)
                    If CDate(dtpDate.Value & " " & .TextMatrix(.Row, .Col)) > CDate(Split(strTime, ";")(1)) Then strInfo = "¼��ʱ����ڱ�������¼��ʱ��Σ�" & Split(strTime, ";")(0) & "��" & Split(strTime, ";")(1)
                End If
            End If
            If strInfo <> "" Then
                mblnEdit = False
                GoTo ErrInfo
            End If
             '���������Դ�Ƿ����Ի����¼����PDA
            mrsTableDetail.Filter = "��Ŀ���=" & Val(.TextMatrix(.Row, COL_tab��Ŀ���)) & " and ��Ŀ����='" & .TextMatrix(.Row, COL_tab��Ŀ����) & "'" & _
                " and ʱ��='" & Format(.TextMatrix(.Row, col_tabԭʼʱ��), "YYYY-MM-DD hh:mm:ss") & "'"
            If mrsTableDetail.RecordCount > 0 Then
                If InStr(1, ",0,3,9,", "," & Val(mrsTableDetail!������Դ) & ",") = 0 Then
                    blnEdit = False
                End If
                cmdColor.Tag = Val(mrsTableDetail!δ��˵��)
            End If
            
            If blnEdit = False Then
                strInfo = "������Դ�ڻ����¼����PDA�����ݲ��ܽ����޸ġ�ɾ������"
                GoTo ErrInfo
            End If
            
            If Not (intType = 2 Or intType = 3) Or vsfTabDetail.Col = vsfTabDetail.FixedCols Then
                picEdit.Width = .CellWidth + 10
                picEdit.Height = .CellHeight - 5
                picEdit.Top = .CellTop + .Top + 20 + fraTabDetail.Top
                picEdit.Left = .CellLeft + .Left + 15
                picEdit.Enabled = True
                picEdit.Visible = True
                picEdit.ZOrder 0
                txtEdit.Top = 0
                txtEdit.Left = 0
                txtEdit.Height = picEdit.Height
            End If
            '������Ŀ�������������͵Ļ��Ŀ����������������ɫ
            If int��Ŀ���� = 1 And intType = 0 And int��Ŀ���� = 2 And vsfTabDetail.Col = vsfTabDetail.FixedCols + 1 Then '�ı����ͣ�� ��Ŀ
                cmdColor.Top = 0
                cmdColor.Height = picEdit.Height
                cmdColor.Width = 300
                cmdColor.Left = picEdit.Width - cmdColor.Width
                txtEdit.Width = cmdColor.Left
                cmdColor.Enabled = True
                cmdColor.Visible = True
                GoTo ShowText
            ElseIf intType = 4 And intƵ�� = 1 And mbln¼��Сʱ = True Then 'ȫ���������ʾ����ʱ��
                
                strTmp = vsfTabDetail.TextMatrix(vsfTabDetail.Row, COL_tab�ַ���)
                lngItemNO = Val(vsfTabDetail.TextMatrix(vsfTabDetail.Row, COL_tab��Ŀ���))
                strTmp = strTmp & String(8 - UBound(Split(strTmp, ",")), ",")
                arrStr = Split(strTmp, ",")
                intNum = Val(arrStr(2))
                intLen = Val(arrStr(6))
                
                If intLen <> 0 Then
                    If lngItemNO <> 4 Then
                        txtEdit.MaxLength = intLen + IIf(intNum = 0, 0, 1)
                    Else
                        txtEdit.MaxLength = intLen * 2 + 1 + IIf(intNum = 0, 0, 1) * 2
                    End If
                Else
                    txtEdit.MaxLength = 0
                End If
                If vsfTabDetail.Col = vsfTabDetail.FixedCols Then txtEdit.MaxLength = 5
                
                If InStr(1, .TextMatrix(vsfTabDetail.Row, vsfTabDetail.Col), ")") > 0 Then
                    txtHour.Text = Replace(Replace(Split(.TextMatrix(vsfTabDetail.Row, vsfTabDetail.Col), ")")(0), "(", ""), "h", "")
                    txtEdit.Text = Split(.TextMatrix(vsfTabDetail.Row, vsfTabDetail.Col), ")")(1)
                Else
                    txtEdit.Text = .TextMatrix(vsfTabDetail.Row, vsfTabDetail.Col)
                End If
                picEdit.Tag = .TextMatrix(vsfTabDetail.Row, vsfTabDetail.Col) & "'" & .Cell(flexcpForeColor, vsfTabDetail.Row, vsfTabDetail.Col)
                txtEdit.Tag = vsfTabDetail.Row & "|" & vsfTabDetail.Col & "|" & "2"
                txtEdit.ForeColor = cmdColor.Tag
                txtEdit.Visible = True
                txtEdit.Enabled = blnEdit
                txtEdit.ZOrder 0
'                picHour.SetFocus
            ElseIf (intType = 2 Or intType = 3) And vsfTabDetail.Col = vsfTabDetail.FixedCols + 1 Then  '��ѡ��ѡ
                strValue = Split(.TextMatrix(vsfTabDetail.Row, COL_tab�ַ���), ",")(0)
                Select Case intType
                    Case 2
                        If Left(strValue, 1) <> ":" Then strValue = ":" & strValue
                        intType = 0
                    Case 3
                        intType = 1
                End Select
                
                arrValue = Split(strValue, ":")
                lstSelect(intType).Clear
                PicLst.Tag = "1"
                For i = 0 To UBound(arrValue)
                    If Left(arrValue(i), 1) = "��" Then arrValue(i) = Mid(arrValue(i), 2): strValue1 = arrValue(i)
                    lstSelect(intType).AddItem arrValue(i), i
                     
                     If intType = 0 Then
                        ReDim arrValue1(0)
                        arrValue1(0) = .TextMatrix(.Row, .Col)
                        txtLst.Text = .TextMatrix(.Row, .Col)
                        txtLst.Tag = 2
                     Else
                        arrValue1 = Split(.TextMatrix(.Row, .Col), ",")
                     End If
                     For j = 0 To UBound(arrValue1)
                        If arrValue1(j) = arrValue(i) Then
                            lstSelect(intType).Selected(i) = True
                            blnSelect = True
                        End If
                    Next j
                Next i
                
                If blnSelect = False And strValue1 <> "" And IIf(intType = 0, Trim(txtLst.Text) = "", True) Then
                    For i = 0 To lstSelect(intType).ListCount - 1
                        If lstSelect(intType).List(i) = strValue1 Then
                            lstSelect(intType).Selected(i) = True
                        End If
                    Next i
                End If
                
                If lstSelect(intType).ListIndex >= 0 Then txtLst.Text = "": PicLst.Tag = 0
                
                '�ؼ���ʾ
                If intType = 0 Then '��ѡ��Ŀ�ṩ����ѡ���¼�빦��
                    PicLst.FontName = .FontName
                    PicLst.FontSize = .FontSize
                    PicLst.Left = .CellLeft + .Left + 15
                    PicLst.Top = picSplitTab.Top + picSplitTab.Height + vsfTabDetail.CellTop + vsfTabDetail.Top
                    PicLst.Height = 80 + (.CellHeight - 5) + PicLst.TextHeight("��") * 2 + lstSelect(intType).ListCount * (PicLst.TextHeight("��") + PicLst.TextHeight("��") \ 4)
                    If PicLst.Height < .CellHeight + 20 Then PicLst.Height = .CellHeight + 20
                    PicLst.Width = LenB(StrConv(lstSelect(intType).List(lstSelect(intType).ListCount \ 2), vbFromUnicode)) * 120 + 500    '���м���ĳ���Ϊ����
                    If PicLst.Width < .CellWidth + 20 Then PicLst.Width = .CellWidth + 20
                    If PicLst.Height > vsfTabDetail.Height Then PicLst.Height = vsfTabDetail.Height
                    If PicLst.Top + PicLst.Height > picSplitTab.Top + picSplitTab.Height + vsfTabDetail.Height Then PicLst.Top = picSplitTab.Top + picSplitTab.Height + .CellTop + .Top + .CellHeight + 20 - PicLst.Height
                    If PicLst.Top < 0 Then PicLst.Top = picSplit.Top + picSplit.Height + vsfTabDetail.Top
                    PicLst.Visible = True
                    PicLst.ZOrder 0
                    
                    lbllst(2).Left = 20
                    lbllst(2).Top = 20
                    If lbllst(2).Width > PicLst.Width Then
                        PicLst.Width = lbllst(2).Width + PicLst.TextWidth("��")
                    End If
                    lbllst(2).FontName = .FontName
                    lbllst(2).FontSize = .FontSize
                    lbllst(2).Visible = True
        
                    txtLst.Top = lbllst(2).Top + lbllst(2).Height + 20
                    txtLst.Left = -10
                    txtLst.Width = PicLst.Width
                    txtLst.Height = .CellHeight - 5
                    txtLst.FontName = .FontName
                    txtLst.FontSize = .FontSize
                    strTmp = vsfTabDetail.TextMatrix(vsfTabDetail.Row, COL_tab�ַ���)
                    strTmp = strTmp & String(8 - UBound(Split(strTmp, ",")), ",")
                    arrStr = Split(strTmp, ",")
                    intNum = Val(arrStr(2))
                    intLen = Val(arrStr(6))
                    txtLst.MaxLength = intLen + IIf(intNum = 0, 0, 1)
                    txtLst.Visible = True
                
                    lbllst(3).Left = 20
                    lbllst(3).Top = txtLst.Top + txtLst.Height + 20
                    lbllst(3).FontName = .FontName
                    lbllst(3).FontSize = .FontSize
                    lbllst(3).Visible = True
                    
                    lstSelect(intType).Top = lbllst(3).Top + lbllst(3).Height + 20
                    lstSelect(intType).Left = -10
                    lstSelect(intType).FontName = .FontName
                    lstSelect(intType).FontSize = .FontSize
                    lstSelect(intType).Width = PicLst.Width
                    lstSelect(intType).Height = PicLst.Height - lstSelect(intType).Top
                    lstSelect(intType).Visible = True
                    lstSelect(intType).Enabled = True
                    lstSelect(intType).ZOrder 0
                    lstSelect(intType).Tag = .TextMatrix(.Row, .Col)
                    lbllst(intType).Tag = .Row & "|" & .Col & "|" & "2"
                    
                    If lstSelect(intType).Top + lstSelect(intType).Height <> PicLst.Height Then
                        PicLst.Height = lstSelect(intType).Top + lstSelect(intType).Height
                    End If
                    PicLst.SetFocus
                Else
                    lstSelect(intType).Top = picSplitTab.Top + picSplitTab.Height + .CellTop + vsfTabDetail.Top
                    lstSelect(intType).Left = .CellLeft + .Left + 15
                    lstSelect(intType).FontName = .FontName
                    lstSelect(intType).FontSize = .FontSize
                    lstSelect(intType).Height = lstSelect(intType).ListCount * (PicLst.TextHeight("��") + PicLst.TextHeight("��") \ 4)
                    If lstSelect(intType).Height < .CellHeight + 20 Then lstSelect(intType).Height = .CellHeight + 20
                    lstSelect(intType).Width = LenB(StrConv(lstSelect(intType).List(lstSelect(intType).ListCount \ 2), vbFromUnicode)) * 120 + 500    '���м���ĳ���Ϊ����
                    If lstSelect(intType).Width < .CellWidth + 20 Then lstSelect(intType).Width = .CellWidth + 20
                    If lstSelect(intType).Height > vsfTabDetail.Height Then
                        lstSelect(intType).Height = vsfTabDetail.Height
                    End If
                    If lstSelect(intType).Top + lstSelect(intType).Height > picSplitTab.Top + picSplitTab.Height + vsfTabDetail.Height Then
                        lstSelect(intType).Top = picSplitTab.Top + picSplitTab.Height + .CellTop + .Top + .CellHeight + 20 - lstSelect(intType).Height
                    End If
                    If lstSelect(intType).Top < 0 Then lstSelect(intType).Top = vsfTabDetail.Top
                    
                        lstSelect(intType).Visible = True
                        lstSelect(intType).Enabled = True
                        lstSelect(intType).ZOrder 0
                        
                        lstSelect(intType).Tag = .TextMatrix(.Row, .Col)
                        lbllst(intType).Tag = .Row & "|" & .Col
                        lstSelect(intType).SetFocus
                    End If
            ElseIf intType = 5 Then 'ѡ��
                lblCheck.Width = picEdit.Width
                lblCheck.Height = picEdit.Height
                lblCheck.Caption = .TextMatrix(vsfTabDetail.Row, vsfTabDetail.Col)
                picEdit.Tag = .TextMatrix(vsfTabDetail.Row, vsfTabDetail.Col) & "'" & .Cell(flexcpForeColor, vsfTabDetail.Row, vsfTabDetail.Col)
                txtEdit.Tag = vsfTabDetail.Row & "|" & vsfTabDetail.Col & "|" & "2"
                lblCheck.Visible = True
                lblCheck.Enabled = True
                lblCheck.ZOrder 0
                picEdit.SetFocus
            Else
                txtEdit.Width = picEdit.Width
ShowText:
                strTmp = vsfTabDetail.TextMatrix(vsfTabDetail.Row, COL_tab�ַ���)
                lngItemNO = Val(vsfTabDetail.TextMatrix(vsfTabDetail.Row, COL_tab��Ŀ���))
                strTmp = strTmp & String(8 - UBound(Split(strTmp, ",")), ",")
                arrStr = Split(strTmp, ",")
                intNum = Val(arrStr(2))
                intLen = Val(arrStr(6))
                
                If intLen <> 0 Then
                    If lngItemNO <> 4 Then
                        txtEdit.MaxLength = intLen + IIf(intNum = 0, 0, 1)
                    Else
                        txtEdit.MaxLength = intLen * 2 + 1 + IIf(intNum = 0, 0, 1) * 2
                    End If
                Else
                    txtEdit.MaxLength = 0
                End If
                If .Col = .FixedCols Then txtEdit.MaxLength = 5
                txtEdit.Text = .TextMatrix(vsfTabDetail.Row, vsfTabDetail.Col)
                picEdit.Tag = .TextMatrix(vsfTabDetail.Row, vsfTabDetail.Col) & "'" & .Cell(flexcpForeColor, vsfTabDetail.Row, vsfTabDetail.Col)
                txtEdit.Tag = vsfTabDetail.Row & "|" & vsfTabDetail.Col & "|" & "2"
                txtEdit.ForeColor = cmdColor.Tag
                txtEdit.Visible = True
                txtEdit.Enabled = True
                txtEdit.ZOrder 0
                picEdit.SetFocus
            
            End If
            
        End If
    End With
ErrFouce:
    If picEdit.Visible = True And txtEdit.Enabled = True Then txtEdit.SetFocus: Call zlControl.TxtSelAll(txtEdit)
ErrInfo:
    If strInfo <> "" Then
        lblStb.Caption = strInfo
        lblStb.ForeColor = 255
    End If
End Sub

Private Sub vsfTabDetail_GotFocus()
    If vsfTab.Tag <> "NO" Then vsfTab.Tag = ""
End Sub

Private Sub vsfTabDetail_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intType As Integer
    Dim intƵ�� As Integer
    Dim strText As String
    Dim blnEdit As Boolean
    Dim blnTrue As Boolean
    
    If vsfTab.Tag = "NO" Then Exit Sub '��ϸ�б��Ǹ��������ʼ��
    If vsfTabDetail.Row < vsfTabDetail.FixedRows And vsfTabDetail.Col < vsfTabDetail.FixedCols Then Exit Sub
    
    '���ε�ĳЩ���ܼ�
    If KeyCode = vbKeyEscape Or KeyCode = vbKeyBack Or Shift <> 0 Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyRight Then Exit Sub
    
    If KeyCode = vbKeyLeft And (picEdit.Visible = False And lstSelect(0).Visible = False And lstSelect(1).Visible = False) Then Exit Sub
    
     With vsfTabDetail
        If KeyCode = vbKeyReturn Then
NextCol2:   '������һ��
            If .Col < vsfTabDetail.FixedCols Then
                .Col = .Col + 1: GoTo NextCol2
            End If
            If .Col < .Cols - 2 Then
                .Col = .Col + 1
                If .ColHidden(.Col) = True Then GoTo NextCol2
            Else
NextRow2:
                If .Row <= Val(.Tag) And .Row <> 0 Then
                    If .TextMatrix(.Row, .FixedCols + 1) <> "" And .Row = Val(.Tag) Then
                        .AddItem "", .Row + 1
                        .TextMatrix(.Row + 1, COL_tab�ַ���) = vsfTab.TextMatrix(vsfTab.Row, COL_tab�ַ���)
                        .TextMatrix(.Row + 1, COL_tab��Ŀ���) = vsfTab.TextMatrix(vsfTab.Row, COL_tab��Ŀ���)
                        .TextMatrix(.Row + 1, COL_tab��Ŀ��) = vsfTab.TextMatrix(vsfTab.Row, COL_tab��Ŀ��)
                        .TextMatrix(.Row + 1, COL_tab��Ŀ����) = vsfTab.TextMatrix(vsfTab.Row, COL_tab��Ŀ��)
                        .Cell(flexcpAlignment, .Row + 1, 0, .Row + 1, .Cols - 1) = flexAlignCenterCenter
                        .Col = vsfTabDetail.FixedCols: .Row = .Row + 1
                        If .RowHidden(.Row) = True Then GoTo NextRow2
                        .Tag = .Tag + 1
                    Else
                        If .Row < .Rows - 1 Then
                            .Col = vsfTabDetail.FixedCols: .Row = .Row + 1
                            If .RowHidden(.Row) = True Then GoTo NextRow2
                        Else
                             .Row = .FixedRows
                            .Col = .FixedCols
                        End If
                    End If
                Else
                    Call txtEdit_KeyPress(vbKeyEscape)
                    .Row = .FixedRows
                    .Col = .FixedCols
                End If
            End If
            '������л��в��ɼ����Զ���ʾ����
            If .ColIsVisible(.Col) = False Then
                .LeftCol = .Col
            End If
            If .RowIsVisible(.Row) = False Then
                .TopRow = .Row
            End If
        
            Exit Sub
        End If
        If KeyCode = vbKeyLeft Then
PreCol2:
            If .Col > vsfTabDetail.FixedCols Then
                .Col = .Col - 1
                If .ColHidden(.Col) = True Then GoTo PreCol2
            Else
PreRow2:
                If .Row > vsfTabDetail.FixedRows Then
                    .Row = .Row - 1
                    If .RowHidden(.Row) Then GoTo PreRow2
                    .Col = Val(Split(vsfTabDetail.TextMatrix(vsfTabDetail.Row, COL_tab�ַ���), ",")(3)) + vsfTabDetail.FixedCols
                    GoTo PreCol2
                End If
            End If
            '������л��в��ɼ����Զ���ʾ����
            If .ColIsVisible(.Col) = False Then
                .LeftCol = .Col
            End If
            If .RowIsVisible(.Row) = False Then
                .TopRow = .Row
            End If
            Exit Sub
        End If
        'ɾ����Ϣ
        If KeyCode = vbKeyDelete Then
            If Shift = 0 And .Col > .FixedCols - 1 And .Col < .Cols - 1 Then
                blnEdit = True
                If .TextMatrix(.Row, .Col) <> "" Then
                    '���������Դ�Ƿ����Ի����¼����PDA
                    mrsTableDetail.Filter = "��Ŀ���=" & Val(.TextMatrix(.Row, COL_tab��Ŀ���)) & " and ��Ŀ����='" & .TextMatrix(.Row, COL_tab��Ŀ��) & "'" & _
                        " and ʱ��='" & vsfTabDetail.TextMatrix(.Row, col_tabԭʼʱ��) & "'"
                    If mrsCurve.RecordCount > 0 Then
                        If InStr(1, ",0,3,9,", "," & Val(mrsCurve!������Դ) & ",") = 0 Then
                            blnEdit = False
                        End If
                    End If
                    If blnEdit = False Then
                        lblStb.Caption = "������Դ�ڻ����¼����PDA�����ݲ��ܽ����޸ġ�ɾ������"
                        lblStb.ForeColor = 255
                        GoTo ErrExit
                    End If
                    picTab.Tag = .Row & "|" & .Col
                    fraTable.Tag = .TextMatrix(.Row, .Col)
                    strText = ""
                    blnTrue = WriteIntoVfgTab(strText, vsfTabDetail, True)
                End If
            End If
ErrExit:
            mblnEdit = False
            Exit Sub
        End If
        mblnEdit = True
        Call vsfTabDetail_EnterCell
     End With
End Sub

Private Sub vsfTabDetail_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    vsfTabDetail.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignLeftCenter
    If mblnFileBack = True Then
        Cancel = True
        vsfTabDetail.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignCenterCenter
        lblStb.Caption = "�������������Ѿ��鵵,��������������޸�."
        lblStb.ForeColor = 255
        Exit Sub
    End If
    If vsfTab.Tag = "NO" Then Cancel = True
End Sub
