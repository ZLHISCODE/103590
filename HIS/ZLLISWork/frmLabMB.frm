VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~2.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~3.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "CO70B6~1.OCX"
Begin VB.Form frmLabMB 
   Caption         =   "ø����"
   ClientHeight    =   7800
   ClientLeft      =   225
   ClientTop       =   525
   ClientWidth     =   14550
   Icon            =   "frmLabMB.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   14550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin MSScriptControlCtl.ScriptControl Calc 
      Left            =   2520
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.PictureBox PicList 
      Appearance      =   0  'Flat
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   0  'None
      FillColor       =   &H00FDD6C6&
      ForeColor       =   &H80000008&
      Height          =   3195
      Left            =   90
      ScaleHeight     =   3195
      ScaleWidth      =   3195
      TabIndex        =   66
      Top             =   2160
      Width           =   3195
      Begin XtremeReportControl.ReportControl rptList 
         Height          =   2205
         Left            =   60
         TabIndex        =   67
         Top             =   360
         Width           =   1875
         _Version        =   589884
         _ExtentX        =   3307
         _ExtentY        =   3889
         _StockProps     =   0
         BorderStyle     =   2
         AutoColumnSizing=   0   'False
      End
      Begin VB.OptionButton opt���� 
         BackColor       =   &H00FDD6C6&
         Caption         =   "����"
         Height          =   255
         Index           =   3
         Left            =   2430
         TabIndex        =   73
         Top             =   60
         Width           =   735
      End
      Begin VB.OptionButton opt���� 
         BackColor       =   &H00FDD6C6&
         Caption         =   "����"
         Height          =   255
         Index           =   2
         Left            =   1640
         TabIndex        =   70
         Top             =   60
         Width           =   735
      End
      Begin VB.OptionButton opt���� 
         BackColor       =   &H00FDD6C6&
         Caption         =   "����"
         Height          =   255
         Index           =   1
         Left            =   850
         TabIndex        =   69
         Top             =   60
         Width           =   735
      End
      Begin VB.OptionButton opt���� 
         BackColor       =   &H00FDD6C6&
         Caption         =   "����"
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   68
         Top             =   60
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.ComboBox cbo�������� 
      Height          =   300
      Left            =   30
      Style           =   2  'Dropdown List
      TabIndex        =   62
      Top             =   1590
      Width           =   1935
   End
   Begin VB.PictureBox PicMain 
      BorderStyle     =   0  'None
      DrawStyle       =   5  'Transparent
      Height          =   7215
      Left            =   3330
      ScaleHeight     =   7215
      ScaleWidth      =   10755
      TabIndex        =   0
      Top             =   30
      Width           =   10755
      Begin VB.Frame fra΢�װ� 
         Caption         =   "΢�װ�"
         Height          =   3945
         Left            =   0
         TabIndex        =   45
         Top             =   3150
         Width           =   10485
         Begin VB.TextBox txt��С���Զ��� 
            Height          =   285
            Left            =   6870
            TabIndex        =   78
            Top             =   180
            Width           =   585
         End
         Begin VB.OptionButton opt��ѡ�� 
            Caption         =   "�ʿ�(QC)"
            Height          =   180
            Index           =   4
            Left            =   4410
            TabIndex        =   77
            Top             =   240
            Width           =   1035
         End
         Begin VB.OptionButton opt��ѡ�� 
            Caption         =   "����(PC)"
            Height          =   180
            Index           =   3
            Left            =   3330
            TabIndex        =   59
            Top             =   240
            Width           =   1035
         End
         Begin VB.OptionButton opt��ѡ�� 
            Caption         =   "����(NC)"
            Height          =   180
            Index           =   2
            Left            =   2250
            TabIndex        =   58
            Top             =   240
            Width           =   1035
         End
         Begin VB.OptionButton opt��ѡ�� 
            Caption         =   "�հ�(BC)"
            Height          =   180
            Index           =   1
            Left            =   1170
            TabIndex        =   57
            Top             =   240
            Width           =   1035
         End
         Begin VB.OptionButton opt��ѡ�� 
            Caption         =   "��ͨ(S)"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   56
            Top             =   240
            Value           =   -1  'True
            Width           =   945
         End
         Begin VB.Frame fra���� 
            BorderStyle     =   0  'None
            Caption         =   "Frame6"
            Height          =   375
            Left            =   60
            TabIndex        =   47
            Top             =   3480
            Width           =   10215
            Begin VB.TextBox txt���λ�� 
               Height          =   300
               Left            =   8610
               TabIndex        =   72
               Top             =   90
               Width           =   1305
            End
            Begin VB.TextBox txtCutOff 
               Height          =   300
               Left            =   6570
               Locked          =   -1  'True
               TabIndex        =   51
               Top             =   90
               Width           =   1005
            End
            Begin VB.TextBox txt���Զ��� 
               Height          =   300
               Left            =   4830
               Locked          =   -1  'True
               TabIndex        =   50
               Top             =   90
               Width           =   1005
            End
            Begin VB.TextBox txt���Զ��� 
               Height          =   300
               Left            =   2850
               Locked          =   -1  'True
               TabIndex        =   49
               Top             =   90
               Width           =   1005
            End
            Begin VB.TextBox txt�հ׶��� 
               Height          =   300
               Left            =   840
               Locked          =   -1  'True
               TabIndex        =   48
               Top             =   90
               Width           =   975
            End
            Begin VB.Label lbl���λ�� 
               AutoSize        =   -1  'True
               Caption         =   "���λ��"
               Height          =   180
               Left            =   7800
               TabIndex        =   71
               Top             =   150
               Width           =   720
            End
            Begin VB.Label lblCutOff 
               AutoSize        =   -1  'True
               Caption         =   "CutOff"
               Height          =   180
               Left            =   5970
               TabIndex        =   55
               Top             =   150
               Width           =   540
            End
            Begin VB.Label lbl���Զ��� 
               AutoSize        =   -1  'True
               Caption         =   "���Զ���"
               Height          =   180
               Left            =   4050
               TabIndex        =   54
               Top             =   150
               Width           =   720
            End
            Begin VB.Label lbl���Զ��� 
               AutoSize        =   -1  'True
               Caption         =   "���Զ���"
               Height          =   180
               Left            =   2070
               TabIndex        =   53
               Top             =   150
               Width           =   720
            End
            Begin VB.Label lbl�հ׶��� 
               AutoSize        =   -1  'True
               Caption         =   "�հ׶���"
               Height          =   180
               Left            =   60
               TabIndex        =   52
               Top             =   150
               Width           =   720
            End
         End
         Begin VB.CommandButton cmd���� 
            Caption         =   "����"
            Height          =   285
            Left            =   8940
            TabIndex        =   46
            Top             =   180
            Width           =   1155
         End
         Begin VSFlex8Ctl.VSFlexGrid vsList 
            Height          =   2835
            Left            =   120
            TabIndex        =   60
            Top             =   510
            Width           =   9855
            _cx             =   17383
            _cy             =   5001
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
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483636
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
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
            Editable        =   2
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
         Begin VB.CheckBox chk���Զ��� 
            Caption         =   "���Զ���С��       ʱ���趨ֵ����"
            Height          =   180
            Left            =   5520
            TabIndex        =   76
            Top             =   240
            Width           =   4065
         End
      End
      Begin VB.Frame fraģ�� 
         Caption         =   "ģ��"
         Height          =   615
         Left            =   0
         TabIndex        =   37
         Top             =   2520
         Width           =   10485
         Begin VB.OptionButton opt���� 
            Caption         =   "����"
            Height          =   255
            Index           =   1
            Left            =   6120
            TabIndex        =   75
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton opt���� 
            Caption         =   "����"
            Height          =   255
            Index           =   0
            Left            =   5400
            TabIndex        =   74
            Top             =   240
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.CommandButton cmdɾ��ģ�� 
            Caption         =   "ɾ��ģ��"
            Height          =   285
            Left            =   3930
            TabIndex        =   42
            Top             =   210
            Width           =   1155
         End
         Begin VB.CommandButton cmd����ģ�� 
            Caption         =   "����ģ��"
            Height          =   285
            Left            =   2670
            TabIndex        =   41
            Top             =   210
            Width           =   1155
         End
         Begin VB.ComboBox cboѡ��ģ�� 
            Height          =   300
            Left            =   990
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   210
            Width           =   1605
         End
         Begin VB.TextBox txt��ʼ�걾�� 
            Height          =   300
            Left            =   7920
            TabIndex        =   39
            ToolTipText     =   " "
            Top             =   210
            Width           =   2025
         End
         Begin VB.CommandButton cmdȷ�� 
            Caption         =   "OK"
            Height          =   285
            Left            =   9960
            TabIndex        =   38
            Top             =   210
            Width           =   375
         End
         Begin VB.Label lblѡ��ģ�� 
            AutoSize        =   -1  'True
            Caption         =   "ѡ��ģ��"
            Height          =   180
            Left            =   150
            TabIndex        =   44
            Top             =   270
            Width           =   720
         End
         Begin VB.Label lbl��ʼ�걾�� 
            AutoSize        =   -1  'True
            Caption         =   "��ʼ�걾��"
            Height          =   180
            Left            =   6990
            TabIndex        =   43
            Top             =   270
            Width           =   900
         End
      End
      Begin VB.Frame fra�������� 
         Caption         =   "��������"
         Height          =   2505
         Left            =   1380
         TabIndex        =   6
         Top             =   0
         Width           =   9105
         Begin VB.TextBox txt�Լ����� 
            Height          =   300
            Left            =   7200
            TabIndex        =   80
            Top             =   624
            Width           =   1305
         End
         Begin VB.TextBox txtCutOff��ʽ 
            Height          =   300
            Left            =   7200
            Locked          =   -1  'True
            TabIndex        =   64
            Top             =   2040
            Width           =   1635
         End
         Begin VB.TextBox txt���ʱ�� 
            Height          =   300
            Left            =   4080
            Locked          =   -1  'True
            TabIndex        =   61
            Top             =   990
            Width           =   1935
         End
         Begin VB.ComboBox cbo���� 
            Height          =   300
            Left            =   1050
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   630
            Width           =   1935
         End
         Begin VB.ComboBox cbo�ο����� 
            Height          =   300
            Left            =   4080
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   630
            Width           =   1935
         End
         Begin VB.ComboBox cbo���Ƶ�� 
            Height          =   300
            Left            =   1050
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   990
            Width           =   1935
         End
         Begin VB.ComboBox cbo���巽ʽ 
            Height          =   300
            Left            =   1050
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   1350
            Width           =   1935
         End
         Begin VB.ComboBox cbo�հ���ʽ 
            Height          =   300
            Left            =   4080
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   1350
            Width           =   1935
         End
         Begin VB.TextBox txt�Լ�Ч�� 
            Height          =   300
            Left            =   7200
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   978
            Width           =   1635
         End
         Begin VB.TextBox txt�Լ����� 
            Height          =   300
            Left            =   7200
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   1332
            Width           =   1635
         End
         Begin VB.TextBox txt���Է��� 
            Height          =   300
            Left            =   7200
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   1686
            Width           =   1635
         End
         Begin VB.TextBox txt���԰�� 
            Height          =   300
            Left            =   7200
            TabIndex        =   12
            Top             =   270
            Width           =   1635
         End
         Begin VB.OptionButton opt������� 
            Caption         =   "�������"
            Height          =   180
            Left            =   4980
            TabIndex        =   11
            Top             =   1770
            Width           =   1065
         End
         Begin VB.OptionButton opt���嵥�� 
            Caption         =   "���嵥��"
            Height          =   180
            Left            =   3870
            TabIndex        =   10
            Top             =   1770
            Value           =   -1  'True
            Width           =   1065
         End
         Begin VB.TextBox txt�����Թ�ʽ 
            Height          =   300
            Left            =   4080
            TabIndex        =   9
            Top             =   2040
            Width           =   1935
         End
         Begin VB.TextBox txt���Թ�ʽ 
            Height          =   300
            Left            =   1050
            TabIndex        =   8
            Top             =   2070
            Width           =   1935
         End
         Begin VB.ComboBox cbo������Ŀ 
            Height          =   300
            Left            =   1050
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1710
            Width           =   2805
         End
         Begin MSComCtl2.DTPicker dtp����ʱ�� 
            Height          =   300
            Left            =   1050
            TabIndex        =   21
            Top             =   270
            Width           =   4965
            _ExtentX        =   8758
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   123142147
            CurrentDate     =   39497
         End
         Begin VB.CommandButton cmdSl 
            Height          =   300
            Left            =   8520
            Picture         =   "frmLabMB.frx":6852
            Style           =   1  'Graphical
            TabIndex        =   81
            Top             =   615
            Width           =   300
         End
         Begin VB.Label lblCutOff��ʽ 
            AutoSize        =   -1  'True
            Caption         =   "CutOff��ʽ"
            Height          =   180
            Left            =   6210
            TabIndex        =   63
            Top             =   2130
            Width           =   900
         End
         Begin VB.Label lbl����ʱ�� 
            AutoSize        =   -1  'True
            Caption         =   "����ʱ��"
            Height          =   180
            Left            =   240
            TabIndex        =   36
            Top             =   330
            Width           =   720
         End
         Begin VB.Label lbl���� 
            AutoSize        =   -1  'True
            Caption         =   "��    ��"
            Height          =   180
            Left            =   240
            TabIndex        =   35
            Top             =   690
            Width           =   720
         End
         Begin VB.Label lbl�ο����� 
            AutoSize        =   -1  'True
            Caption         =   "�ο�����"
            Height          =   180
            Left            =   3300
            TabIndex        =   34
            Top             =   690
            Width           =   720
         End
         Begin VB.Label lbl���Ƶ�� 
            AutoSize        =   -1  'True
            Caption         =   "���Ƶ��"
            Height          =   180
            Left            =   240
            TabIndex        =   33
            Top             =   1050
            Width           =   720
         End
         Begin VB.Label lbl���ʱ�� 
            AutoSize        =   -1  'True
            Caption         =   "���ʱ��"
            Height          =   180
            Left            =   3300
            TabIndex        =   32
            Top             =   1050
            Width           =   720
         End
         Begin VB.Label lbl���巽ʽ 
            AutoSize        =   -1  'True
            Caption         =   "���巽ʽ"
            Height          =   180
            Left            =   240
            TabIndex        =   31
            Top             =   1410
            Width           =   720
         End
         Begin VB.Label lbl�հ���ʽ 
            AutoSize        =   -1  'True
            Caption         =   "�հ���ʽ"
            Height          =   180
            Left            =   3300
            TabIndex        =   30
            Top             =   1410
            Width           =   720
         End
         Begin VB.Label lbl���Թ�ʽ 
            AutoSize        =   -1  'True
            Caption         =   "���Թ�ʽ"
            Height          =   180
            Left            =   240
            TabIndex        =   29
            Top             =   2130
            Width           =   720
         End
         Begin VB.Label lbl�����Թ�ʽ 
            AutoSize        =   -1  'True
            Caption         =   "�����Թ�ʽ"
            Height          =   180
            Left            =   3120
            TabIndex        =   28
            Top             =   2130
            Width           =   900
         End
         Begin VB.Label lbl�Լ����� 
            AutoSize        =   -1  'True
            Caption         =   "�Լ�����"
            Height          =   180
            Left            =   6390
            TabIndex        =   27
            Top             =   1410
            Width           =   720
         End
         Begin VB.Label lbl���Է��� 
            AutoSize        =   -1  'True
            Caption         =   "���Է���"
            Height          =   180
            Left            =   6390
            TabIndex        =   26
            Top             =   1770
            Width           =   720
         End
         Begin VB.Label lbl���԰�� 
            AutoSize        =   -1  'True
            Caption         =   "���԰��"
            Height          =   180
            Left            =   6390
            TabIndex        =   25
            Top             =   330
            Width           =   720
         End
         Begin VB.Label lbl�Լ����� 
            AutoSize        =   -1  'True
            Caption         =   "�Լ�����"
            Height          =   180
            Left            =   6390
            TabIndex        =   24
            Top             =   690
            Width           =   720
         End
         Begin VB.Label lbl�Լ�Ч�� 
            AutoSize        =   -1  'True
            Caption         =   "�Լ�Ч��"
            Height          =   180
            Left            =   6390
            TabIndex        =   23
            Top             =   1050
            Width           =   720
         End
         Begin VB.Label lbl������Ŀ 
            AutoSize        =   -1  'True
            Caption         =   "������Ŀ"
            Height          =   180
            Left            =   240
            TabIndex        =   22
            Top             =   1770
            Width           =   720
         End
      End
      Begin VB.Frame fra��ʾ��ʽ 
         Caption         =   "��ʾ��ʽ"
         Height          =   2505
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   1305
         Begin VB.OptionButton opt��ʾ 
            Caption         =   "����ֵ"
            Height          =   180
            Index           =   3
            Left            =   120
            TabIndex        =   5
            Top             =   1470
            Width           =   1125
         End
         Begin VB.OptionButton opt��ʾ 
            Caption         =   "ODֵ"
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   4
            Top             =   1070
            Width           =   1125
         End
         Begin VB.OptionButton opt��ʾ 
            Caption         =   "ԭʼODֵ"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   3
            Top             =   670
            Width           =   1125
         End
         Begin VB.OptionButton opt��ʾ 
            Caption         =   "�������"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   2
            Top             =   270
            Value           =   -1  'True
            Width           =   1125
         End
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgList 
      Height          =   720
      Left            =   0
      TabIndex        =   65
      Top             =   0
      Visible         =   0   'False
      Width           =   1305
      _cx             =   2302
      _cy             =   1270
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
      BackColorFixed  =   15790320
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
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
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   79
      Top             =   7440
      Width           =   14550
      _ExtentX        =   25665
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmLabMB.frx":D0A4
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   20585
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   1800
      Top             =   210
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmLabMB.frx":D938
      Left            =   1530
      Top             =   1020
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmLabMB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const conPane_List = 201
Const conPane_Base = 202
Const conFontColor_BC = vbCyan
Const conFontColor_NC = vbBlue
Const conFontColor_PC = vbRed
Const conFontColor_QC = vbGreen
Const conFontColor_BK = vbBlack
Const conFontColor_YR = vbMagenta
Const conFontColor_YL = &H7B55DF

Private mlngEditWidth As Long       'Ϊ��Ӧ����������´�����.�ȶ��봰���С.
Private Enum mCol
    ID = 0: ���: ����ʱ��: �Լ�����: �Լ�Ч��: �Լ�����: ���Է���: ����: �ο�����: ���Ƶ��: ���ʱ��: ���巽ʽ: �հ���ʽ: OD���հ�: ���嵥��: ������Ŀ: ���Թ�ʽ: �����Թ�ʽ: CutOff��ʽ: ���Խ��: ���λ��: �Լ���¼
End Enum
Private mEditState As Integer                           '�༭״̬: 0=��� 1=���� 2=�޸�
Private mTestData(3, 1 To 8, 1 To 12) As String         ' һά����0=���;1=ԭʼOD:2=OD;3=����) ��ά��ά:(΢�װ�����)
Private mTestItem(2, 1 To 8) As String                  'ÿһ�еĹ�ʽ(0=���Թ�ʽ;1=�����Թ�ʽ;2=Cutoff��ʽ
Private mTestReagent(1 To 8) As String                  '��һ�е��Լ�ID
Private mintEditState As Integer
Private mblnShowStop As Boolean
Private mlngKey As Long                         '��ǰ��¼��ID
Private mbln_Init As Boolean                    '�����Ƿ��ʼ���ɹ�
Private mblnModify As Boolean                   '�Ƿ������޸�
Private mblnRefresh As Boolean                  '�Ƿ�ˢ������
Private mstr��ʽ As String
Private mlngMachine As Long                     '����ID
Private mrsCalc As adodb.Recordset              '��¼���㹫ʽ
Private mblnMBSelect As Boolean                 '�Ƿ�ѡ����ģ��

Private Sub cbo������Ŀ_Click()
    Dim rsTmp As New adodb.Recordset
    Dim intRow As Integer
    Dim intLoop As Integer
    On Error GoTo errH
    If Me.Visible = False Then Exit Sub
    If Me.cbo������Ŀ.ListIndex = -1 Then Exit Sub
    
    
    'ȡ������Ŀ
    mrsCalc.filter = "������ĿID=" & Val(Me.cbo������Ŀ.ItemData(Me.cbo������Ŀ.ListIndex))
    If mrsCalc.EOF = False Then
        Me.txt���Թ�ʽ.Text = mrsCalc("���Թ�ʽ") & ""
        Me.txt�����Թ�ʽ.Text = mrsCalc("�����Թ�ʽ") & ""
        Me.txtCutOff��ʽ.Text = mrsCalc("Cutoff��ʽ") & ""
    End If
    
    
    
    
'    Me.txt���Թ�ʽ.Text = mTestItem(0, Me.vsList.Row)
'    Me.txt�����Թ�ʽ.Text = mTestItem(1, Me.vsList.Row)
'    Me.txtCutOff��ʽ.Text = mTestItem(2, Me.vsList.Row)
'
'
'    If rsTmp.EOF = True Then Exit Sub
'    If Me.vsList.Row = 0 Then Me.vsList.Select 1, 1
'    If mTestItem(0, Me.vsList.Row) = "" Or mblnModify = True Then mTestItem(0, Me.vsList.Row) = Nvl(rsTmp("���Թ�ʽ"))
'    If mTestItem(1, Me.vsList.Row) = "" Or mblnModify = True Then mTestItem(1, Me.vsList.Row) = Nvl(rsTmp("�����Թ�ʽ"))
'    If mTestItem(2, Me.vsList.Row) = "" Or mblnModify = True Then mTestItem(2, Me.vsList.Row) = Nvl(rsTmp("CutOff��ʽ"))
'    vsList.TextMatrix(Me.vsList.Row, 13) = Me.cbo������Ŀ.ItemData(Me.cbo������Ŀ.ListIndex)
    With Me.vsList
        If .Cols >= 13 And .Rows > 0 Then
            If Me.opt���嵥�� Then
                For intLoop = 1 To Me.vsList.Rows - 1

                    .TextMatrix(intLoop, 13) = Me.cbo������Ŀ.ItemData(Me.cbo������Ŀ.ListIndex)
                    '���Թ�ʽ
                    If mTestItem(0, intLoop) = "" Or mblnModify = True Then mTestItem(0, intLoop) = Me.txt���Թ�ʽ.Text
                    If mTestItem(1, intLoop) = "" Or mblnModify = True Then mTestItem(1, intLoop) = Me.txt�����Թ�ʽ.Text
                    If mTestItem(2, intLoop) = "" Or mblnModify = True Then mTestItem(2, intLoop) = Me.txtCutOff��ʽ.Text
                Next
            End If
        End If
    End With
'
'    Me.txt���Թ�ʽ.Text = mTestItem(0, Me.vsList.Row)
'    Me.txt�����Թ�ʽ.Text = mTestItem(1, Me.vsList.Row)
'    Me.txtCutOff��ʽ.Text = mTestItem(2, Me.vsList.Row)
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbo������Ŀ_GotFocus()
    mblnModify = True
End Sub

Private Sub cbo������Ŀ_LostFocus()
    mblnModify = False
End Sub

Private Sub cbo��������_Click()
    Dim rsTmp As New adodb.Recordset
    Dim aItem() As String
    Dim intLoop As Integer
    
    On Error GoTo errH
    
    If Me.cbo��������.ListCount = 0 Then Exit Sub
    
    If mbln_Init Then frmLabMBControl.MB_Stop  'ֹͣ�ѳ�ʼ����������
    
    If cbo��������.ListIndex >= 0 Then
        mlngMachine = cbo��������.ItemData(cbo��������.ListIndex)
    End If
    
    gstrSql = "select ����,���Ƶ��,���ʱ��,���巽ʽ,�հ���ʽ from �������� where id = [1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.cbo��������.ItemData(Me.cbo��������.ListIndex))
    If rsTmp.EOF = True Then Exit Sub
    
'    If mEditState = 0 Then Exit Sub
    
            
    With Me.cbo����
        .Clear
        Me.cbo�ο�����.Clear
        Me.cbo�ο�����.AddItem ""
        Me.cbo�ο�����.ItemData(Me.cbo�ο�����.NewIndex) = 0
        aItem = Split(Nvl(rsTmp("����")), ";")
        For intLoop = 0 To UBound(aItem)
            .AddItem aItem(intLoop)
            Me.cbo�ο�����.AddItem aItem(intLoop)
        Next
        If .ListCount > 0 Then .ListIndex = 0
    End With
    
    With Me.cbo���Ƶ��
        .Clear
        aItem = Split(Nvl(rsTmp("���Ƶ��")), ";")
        For intLoop = 0 To UBound(aItem)
            .AddItem aItem(intLoop)
        Next
        If .ListCount > 0 Then .ListIndex = 0
    End With
    
    Me.txt���ʱ��.Text = Nvl(rsTmp("���ʱ��"))
    
    With Me.cbo���巽ʽ
        .Clear
        aItem = Split(Nvl(rsTmp("���巽ʽ")), ";")
        For intLoop = 0 To UBound(aItem)
            .AddItem aItem(intLoop)
        Next
        If .ListCount > 0 Then .ListIndex = 0
    End With
    
    With Me.cbo�հ���ʽ
        .Clear
        aItem = Split(Nvl(rsTmp("�հ���ʽ")), ";")
        For intLoop = 0 To UBound(aItem)
            .AddItem aItem(intLoop)
        Next
        If .ListCount > 0 Then .ListIndex = 0
    End With
    
    Call RefreshList
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cboѡ��ģ��_Click()
    Dim rsTmp As New adodb.Recordset
    Dim intLoop As Integer
    Dim intRow As Integer, intCol As Integer
    Dim aResult() As String
    Dim aItem() As String
    Dim blnOne As Boolean
    Dim lngItemID As Long
    
    If Me.cboѡ��ģ��.ItemData(Me.cboѡ��ģ��.ListIndex) <= 0 Then Exit Sub
    
    gstrSql = "select id,���,����,��Ŀ,���� from ����ø��ģ�� where id = [1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.cboѡ��ģ��.ItemData(Me.cboѡ��ģ��.ListIndex))
    
    If rsTmp.EOF = True Then
        MsgBox "û���ҵ�ģ���¼!", vbInformation
        Exit Sub
    End If
            
   aItem = Split(rsTmp("��Ŀ"), ";")
   aResult = Split(rsTmp("����"), "|")
   
   intLoop = Val(Me.txt��ʼ�걾��)
   
   If Me.opt����(0).Value = True Then
        '����
        For intRow = 1 To 8
            If Me.opt�������.Value = True Then intLoop = Val(Me.txt��ʼ�걾��)
            For intCol = 1 To 12
                With Me.vsList
                    If intLoop = 0 Then
                        .TextMatrix(intRow, intCol) = Split(aResult(intRow - 1), ";")(intCol - 1)
                    Else
                        If Trim(.TextMatrix(intRow, intCol)) <> "" Then
                            If IsNumeric(Split(aResult(intRow - 1), ";")(intCol - 1)) = True Then
                                .TextMatrix(intRow, intCol) = intLoop
                                intLoop = intLoop + 1
                            Else
                                .TextMatrix(intRow, intCol) = Split(aResult(intRow - 1), ";")(intCol - 1)
                            End If
                        End If
                    End If
                    mTestData(0, intRow, intCol) = .TextMatrix(intRow, intCol)
                End With
            Next
        Next
    Else
        '����
        For intCol = 1 To 12
            For intRow = 1 To 8
                With Me.vsList
                    If intLoop = 0 Then
                        .TextMatrix(intRow, intCol) = Split(aResult(intRow - 1), ";")(intCol - 1)
                    Else
                        If Trim(.TextMatrix(intRow, intCol)) <> "" Then
                            If IsNumeric(Split(aResult(intRow - 1), ";")(intCol - 1)) = True Then
                                .TextMatrix(intRow, intCol) = intLoop
                                intLoop = intLoop + 1
                            Else
                                .TextMatrix(intRow, intCol) = Split(aResult(intRow - 1), ";")(intCol - 1)
                            End If
                        End If
                    End If
                    mTestData(0, intRow, intCol) = .TextMatrix(intRow, intCol)
                End With
            Next
        Next
    End If
    
    
    For intRow = 1 To 8
        With Me.vsList
            .TextMatrix(intRow, 13) = aItem(intRow - 1)
            
            
            If lngItemID = 0 And .TextMatrix(intRow, 13) <> "" Then
                lngItemID = .TextMatrix(intRow, 13)
            End If
            If intRow > 1 Then
                If .TextMatrix(intRow, 13) <> aItem(intRow - 2) Then
                    blnOne = True
                End If
            End If
            
        End With
    Next
    If blnOne = True Then
        Me.opt�������.Value = True
    Else
        Me.opt���嵥��.Value = True
    End If
    
    Me.txt��ʼ�걾��.Text = ""
    
    For intLoop = 0 To 3
        If Me.opt��ʾ(intLoop).Value = True Then
            Call opt��ʾ_Click(intLoop)
        End If
    Next
    
    Erase mTestItem
    For intRow = 1 To 8
        For intLoop = 0 To Me.cbo������Ŀ.ListCount - 1
            If Val(aItem(intRow - 1)) = Me.cbo������Ŀ.ItemData(intLoop) Then
                Me.vsList.Row = intRow
                Me.cbo������Ŀ.ListIndex = intLoop
                If mTestItem(0, intRow) = "" Then mTestItem(0, intRow) = Me.txt���Թ�ʽ
                If mTestItem(1, intRow) = "" Then mTestItem(1, intRow) = Me.txt�����Թ�ʽ
                If mTestItem(2, intRow) = "" Then mTestItem(2, intRow) = Me.txtCutOff��ʽ
                Exit For
            End If
        Next
    Next
    
    mblnMBSelect = True
'    With Me.vsList
'        For intRow = 1 To 8
'            If Val(.TextMatrix(intRow, 13)) <> 0 Then
'                With cbo������Ŀ
'                    For intLoop = 0 To .ListCount - 1
'                        If .ItemData(intLoop) = Val(vsList.TextMatrix(intRow, 13)) Then
'                            .ListIndex = intLoop
'                        End If
'                    Next
'                End With
'            End If
'        Next
'    End With
'
'    If lngItemID <> 0 Then
'        With cbo������Ŀ
'            For intRow = 0 To .ListCount - 1
'                If .ItemData(intRow) = lngItemID Then
'                    .ListIndex = intRow
'                End If
'            Next
'        End With
'    End If
'    Me.cboѡ��ģ��.ListIndex = 0
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim cbrControl As CommandBarControl                 '�ı���ǩ
    Dim strFilter As String                             '�����ִ�
    Dim rsTmp As New adodb.Recordset
    
    Select Case Control.ID
        Case conMenu_File_PrintSet                                                          '��ӡ����
            zlPrintSet
        Case conMenu_File_Preview                                                           'Ԥ��
            Call zlRptPrint(0)
        Case conMenu_File_Print                                                             '��ӡ
            Call zlRptPrint(1)
        Case conMenu_File_Excel                                                             '�����Excel
            Call zlRptPrint(2)
        Case conMenu_File_Parameter                                                         '��������
            frmLabMBSetup.Show vbModal, Me
        Case conMenu_Edit_Save                                                              '����
            Call SaveData
        Case conMenu_Edit_Untread                                                           'ȡ��
            Call InitItem: mEditState = 0: RefreshItem (mlngKey)
        Case conMenu_File_Exit                                                              '�˳�
            Unload Me
        '----------------------------------------------------------------------------------------------------
        Case conMenu_Edit_NewItem                                                           '����
            Call AddNew
        Case conMenu_Edit_Modify                                                            '�޸�
            mEditState = 2
        Case conMenu_Edit_Delete                                                            'ɾ��
            Call DelData
        Case conMenu_Edit_Leave_Post                                                        '����
            Call CalcData
        Case conMenu_Edit_Send                                                              '����
            Call MBcontrol
        Case conMenu_LIS_MB_Connect                                                               'ѡ������(��������)
            If Me.cbo��������.ListIndex >= 0 Then
                mbln_Init = frmLabMBControl.MB_Start(Me, Me.cbo��������.ItemData(Me.cbo��������.ListIndex))
            Else
                MsgBox "��ѡ��һ��ø����", vbInformation, Me.Caption
            End If
        Case conMenu_LIS_MB_Disconnect                                                               'ȡ��ѡ��(�Ͽ���������)
            Call frmLabMBControl.MB_Stop
            mbln_Init = False
        Case conMenu_Edit_QCRes                                                             '�Լ�����
            frmLabMBReagent.Show vbModal, Me

        Case conMenu_Edit_Adjust                                                            '��������OD
            mstr��ʽ = frmLabMBcalc.ShowMe(Me)
            Call CalcData
            
        '-----------------------------------------------------------------------------------------------------
        Case conMenu_View_ToolBar_Button                                                    '��׼��ť
            Me.cbsThis(2).Visible = Not Me.cbsThis(2).Visible
            Me.cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Text                                                      '�ı���ǩ
            For Each cbrControl In Me.cbsThis(2).Controls
                cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            Me.cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Size                                                      '��ͼ��
            Me.cbsThis.Options.LargeIcons = Not Me.cbsThis.Options.LargeIcons
            Me.cbsThis.RecalcLayout
        Case conMenu_View_StatusBar                                                         '״̬��
            
        Case conMenu_View_Find                                                              '����
            strFilter = frmLabMBFilter.ShowMe(Me)
            If strFilter <> "" Then Call RefreshList(2, strFilter)
        Case conMenu_View_Refresh                                                           'ˢ��
            Call RefreshList
        '-----------------------------------------------------------------------------------------------------
        Case conMenu_Help_Help                                                              '��������
            Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_Help_Web                                                               'WEB
            Call zlHomePage(Me.hWnd)
        Case conMenu_Help_Web_Home                                                          '��ҳ
            Call zlHomePage(Me.hWnd)
        Case conMenu_Help_Web_Mail                                                          '���ͷ���
            Call zlMailTo(Me.hWnd)
        Case conMenu_Help_About                                                             '����
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    End Select
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height

End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_File_PrintSet                                                          '��ӡ����
        Case conMenu_File_Preview                                                           'Ԥ��
        Case conMenu_File_Print                                                             '��ӡ
        Case conMenu_File_Excel                                                             '�����Excel
        Case conMenu_Edit_Save                                                              '����
            Control.Enabled = (mEditState > 0)
        Case conMenu_Edit_Untread                                                           'ȡ��
            Control.Enabled = (mEditState > 0)
        '----------------------------------------------------------------------------------------------------
        Case conMenu_Edit_NewItem                                                           '����
            Control.Enabled = (mEditState = 0)
        Case conMenu_Edit_Modify                                                            '�޸�
            Control.Enabled = (mEditState = 0 And Me.rptList.Records.Count > 0)
        Case conMenu_Edit_Delete                                                            'ɾ��
            Control.Enabled = (mEditState = 0 And Me.rptList.Records.Count > 0)
        Case conMenu_Edit_Leave_Post                                                        '����
            Control.Enabled = (mEditState > 0)
        Case conMenu_Edit_Send                                                              '����
            Control.Enabled = (mEditState > 0) And mbln_Init
        Case conMenu_LIS_MB_Connect                                                               '����
            
            Control.Enabled = Not mbln_Init
        
        Case conMenu_LIS_MB_Disconnect                                                               '�Ͽ�
            Control.Enabled = mbln_Init
        
        Case conMenu_Edit_Adjust                                                            '��������
            Control.Enabled = (mEditState > 0)
        '-----------------------------------------------------------------------------------------------------
        Case conMenu_View_ToolBar_Button                                                    '��׼��ť
            Control.Checked = Me.cbsThis(2).Visible
        Case conMenu_View_ToolBar_Text                                                      '�ı���ǩ
            Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
        Case conMenu_View_ToolBar_Size                                                      '��ͼ��
            Control.Checked = Me.cbsThis.Options.LargeIcons
        Case conMenu_View_StatusBar                                                         '״̬��
        Case conMenu_View_Find                                                              '����
            
        Case conMenu_View_Refresh                                                           'ˢ��
        '-----------------------------------------------------------------------------------------------------
        Case conMenu_Help_Help                                                              '��������
        Case conMenu_Help_Web                                                               'WEB
        Case conMenu_Help_Web_Home                                                          '��ҳ
        Case conMenu_Help_Web_Mail                                                          '���ͷ���
        Case conMenu_Help_About                                                             '����
    End Select
    If mEditState = 0 Then
        Me.fra��������.Enabled = False
        Me.fraģ��.Enabled = False
'        Me.fra΢�װ�.Enabled = False
        Me.vsList.Editable = flexEDNone
        PicList.Enabled = True
    Else
        Me.fra��������.Enabled = True
        Me.fraģ��.Enabled = True
        'Me.fra΢�װ�.Enabled = True
        Me.vsList.Editable = flexEDKbdMouse
        PicList.Enabled = False
    End If
End Sub

Private Sub cmdSl_Click()
    Me.txt�Լ����� = ""
    Call SelectBatch
End Sub

Private Sub cmd����ģ��_Click()
    Call SaveTemplet
End Sub

Private Sub cmdȷ��_Click()
    Call subWriteNumber
End Sub

Private Sub cmdɾ��ģ��_Click()
    If Me.cboѡ��ģ��.Text = "" Then Exit Sub
    If MsgBox("�Ƿ�ȷ��Ҫɾ��<" & Me.cboѡ��ģ�� & ">ģ��?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
        'ɾ��
        On Error GoTo errH
        gstrSql = "Zl_����ø��ģ��_Delete(" & Me.cboѡ��ģ��.ItemData(Me.cboѡ��ģ��.ListIndex) & ")"
        zlDatabase.ExecuteProcedure gstrSql, Me.Caption
        RefreshTemplet
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmd����_Click()
    Dim intY As Integer, intX As Integer
    Erase mTestData
    Erase mTestItem
    Erase mTestReagent
    With Me.vsList
        For intY = 1 To .Rows - 1
            For intX = 1 To .Cols - 1
                .TextMatrix(intY, intX) = ""
            Next
        Next
    End With
End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_List
        Item.Handle = Me.PicList.hWnd
    Case conPane_Base
        Item.Handle = Me.PicMain.hWnd
    End Select
End Sub

Private Sub dkpMan_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Dim lngTop As Long, lngLeft As Long, lngRight As Long, lngBottom As Long
    Me.cbsThis.GetClientRect lngLeft, lngTop, lngRight, lngBottom
    Top = lngTop
    Bottom = Me.ScaleHeight - lngBottom
End Sub

Private Sub dkpMan_Resize()
    Me.cbsThis.RecalcLayout
End Sub

Private Sub Form_Load()
    Dim intLoop As Integer
    Dim intX As Integer, intY As Integer
    Dim rsTmp As New adodb.Recordset
    Dim cbrControl As CommandBarControl
    Dim cbrPopControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim cbrCustom As CommandBarControlCustom
    Dim strName As String
    Dim lngMachine As Long
    Dim rptCol As ReportColumn
    '-----------------------------------------------------
    'Ȩ�����ƴ����ƣ�����ͬʱ��������ģ�������gstrPrivs�仯�����¿�����Ч
'    mstrPrivs = gstrPrivs
    
    mlngEditWidth = Me.PicMain.Width
    
    mintEditState = 0: mblnShowStop = False
    Me.cbsThis.EnableCustomization False
    
'    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, False)
   '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Set Me.cbsThis.Icons = zlCommFun.GetPubIcons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsThis.EnableCustomization False
    '-----------------------------------------------------
    '�˵�����
    Me.cbsThis.ActiveMenuBar.Title = "�˵�"
'    Me.cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "��������(&O)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "����(&S)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Untread, "ȡ��(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��(&D)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Leave_Post, "����(&C)"): cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_LIS_MB_Connect, "��������(&N)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Send, "����(&S)")
        Set cbrControl = .Add(xtpControlButton, conMenu_LIS_MB_Disconnect, "�Ͽ�����(&N)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_QCRes, "�Լ�����(&R)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Adjust, "����OD����(&O)"): cbrControl.BeginGroup = True
        
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False
'        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Find, "����(&F)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): cbrControl.BeginGroup = True
    End With
    
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): cbrControl.BeginGroup = True
    End With
    
    Set cbrControl = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlLabel, 0, "��������")
    cbrControl.Flags = xtpFlagRightAlign

    Set cbrCustom = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlCustom, conMenu_Report_DrugQuery, "��������")
    cbrCustom.ShortcutText = "��������"
    cbrCustom.Handle = Me.cbo��������.hWnd
    cbrCustom.Flags = xtpFlagRightAlign
    cbrCustom.Style = xtpButtonIconAndCaption
    
    '�����
    With Me.cbsThis.KeyBindings
        .Add FCONTROL, Asc("S"), conMenu_Edit_Save
        .Add FCONTROL, Asc("Z"), conMenu_Edit_Untread
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add FSHIFT, VK_DELETE, conMenu_Edit_Delete
        .Add FCONTROL, Asc("B"), conMenu_Edit_Compend
        .Add FCONTROL, Asc("E"), conMenu_Edit_ApplyTo
        .Add FCONTROL, Asc("G"), conMenu_Edit_Test
        .Add FCONTROL, Asc("F"), conMenu_View_Find
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    '���ò����ò˵�
    With Me.cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_Edit_Pause
        .AddHiddenCommand conMenu_Edit_Reuse
        .AddHiddenCommand conMenu_View_Refresh
        .AddHiddenCommand conMenu_View_Option
    End With
    '-----------------------------------------------------
    '����������
    Set cbrToolBar = Me.cbsThis.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Untread, "ȡ��")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Leave_Post, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_LIS_MB_Connect, "����")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Send, "����")
        Set cbrControl = .Add(xtpControlButton, conMenu_LIS_MB_Disconnect, "�Ͽ�")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Find, "����")
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    '-----------------------------------------------------
    '���ôʾ���ʾͣ������
    Dim panSub1 As Pane, panSub2 As Pane, panSub3 As Pane

    
    Set panSub1 = dkpMan.CreatePane(conPane_List, 300, 580, DockLeftOf, Nothing)
    panSub1.Title = "���԰��б�"
    panSub1.Options = PaneNoCaption

    Set panSub2 = dkpMan.CreatePane(conPane_Base, 550, 200, DockRightOf, Nothing)
    panSub2.Title = "���ƽ���"
    panSub2.Options = PaneNoCaption

    panSub1.Select
    Me.dkpMan.SetCommandBars Me.cbsThis
    Me.dkpMan.Options.ThemedFloatingFrames = True
    Me.dkpMan.Options.HideClient = True
    
    '-----------------------------------------------------
    With Me.rptList
        Set rptCol = .Columns.Add(mCol.ID, "ID", 0, False): rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.���, "���", 80, True): .SortOrder.Add rptCol
        Set rptCol = .Columns.Add(mCol.����ʱ��, "����ʱ��", 85, True)
        Set rptCol = .Columns.Add(mCol.�Լ�����, "�Լ�����", 85, True)
        Set rptCol = .Columns.Add(mCol.�Լ�Ч��, "�Լ�Ч��", 85, True)
        Set rptCol = .Columns.Add(mCol.�Լ�����, "�Լ�����", 85, True)
        Set rptCol = .Columns.Add(mCol.���Է���, "���Է���", 85, True)
        Set rptCol = .Columns.Add(mCol.����, "����", 85, False): rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.�ο�����, "�ο�����", 85, False): rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.���Ƶ��, "���Ƶ��", 85, False): rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.���ʱ��, "���ʱ��", 85, False): rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.���巽ʽ, "���巽ʽ", 85, False): rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.�հ���ʽ, "�հ���ʽ", 85, False): rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.OD���հ�, "OD���հ�", 85, False): rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.���嵥��, "���嵥��", 85, False): rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.������Ŀ, "������Ŀ", 85, False): rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.���Թ�ʽ, "������ʽ", 85, False): rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.�����Թ�ʽ, "�����Թ�ʽ", 85, False): rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.CutOff��ʽ, "CutOff��ʽ", 85, False): rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.���Խ��, "���Խ��", 85, False): rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.���λ��, "���λ��", 85, False): rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.�Լ���¼, "�Լ���¼", 85, False): rptCol.Visible = False
        
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        With .PaintManager
            .ColumnStyle = xtpColumnShaded
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ����Ŀ..."
            .VerticalGridStyle = xtpGridSolid
        End With
    End With
    
    '-----------------------------------------------------
    '����ָ�
'    Call RestoreWinState(Me, App.ProductName)
    '-----------------------------------------------------
    With Me.vsList
        .Rows = 9
        .Cols = 14
        .FixedRows = 1
        .FixedCols = 1 '
        For intLoop = 1 To .Cols - 2
            .TextMatrix(0, intLoop) = intLoop
        Next
        .TextMatrix(0, 13) = "��Ŀ"
        
        For intLoop = 1 To .Rows - 1
            .TextMatrix(intLoop, 0) = Chr(intLoop + 64)
        Next
       
       .Select 0, 0, 8, 13

      .FillStyle = flexFillRepeat

      .CellAlignment = flexAlignCenterCenter

      'return .FillStyle to its default (if needed)

      .FillStyle = flexFillSingle
      .Select 0, 0, 0, 0
     
      .Cell(flexcpBackColor, 1, 13, 8, 13) = RGB(200, 200, 200)
    End With
    
    Call InitRecordSet(mrsCalc)
    
'    Me.chk���Զ���.Value = Mid(GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "���Զ���", "0,"), 1, 1)
'    Me.txt��С���Զ���.Text = Mid(GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "���Զ���", "0,"), 3)
    If zlDatabase.GetPara("frmLabMB_���Զ���", 100, 1208, "") = "" Then
        Me.chk���Զ���.Value = 0
        Me.txt��С���Զ���.Text = ""
    Else
        Me.chk���Զ���.Value = Mid(zlDatabase.GetPara("frmLabMB_���Զ���", 100, 1208, "0,"), 1, 1)
        Me.txt��С���Զ���.Text = Mid(zlDatabase.GetPara("frmLabMB_���Զ���", 100, 1208, "0,"), 3)
    End If
    '��������
'    lngMachine = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "����ID", 0)
'    lngMachine = zlDatabase.GetPara("frmLabMB_����ID", 100, 1208, 0)
    
    gstrSql = "select id,����,���� from �������� where  ΢���� = 2"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    With Me.cbo��������
        .Clear
        Do Until rsTmp.EOF
            .AddItem rsTmp("����") & "-" & rsTmp("����")
            .ItemData(.NewIndex) = rsTmp("ID")
            If rsTmp("ID") = mlngMachine Then
                .ListIndex = .NewIndex
            End If
            rsTmp.MoveNext
        Loop
        If .ListCount > 0 Then
            If .ListIndex < 0 Then .ListIndex = 0
        End If
    End With
    
    '���������Ŀ
    If Me.cbo��������.ListCount > 0 Then
        gstrSql = "select id,������,Ӣ���� from  ����������Ŀ a , ������Ŀ b,����������Ŀ c  where a.id = b.������Ŀid and ��Ŀ��� = 4 " & _
                   " And c.����id = [1] And a.id = c.��Ŀid "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.cbo��������.ItemData(Me.cbo��������.ListIndex)))
        With Me.cbo������Ŀ
            .Clear
            Do Until rsTmp.EOF
                .AddItem Nvl(rsTmp("������")) & "(" & Nvl(rsTmp("Ӣ����")) & ")"
                .ItemData(.NewIndex) = rsTmp("id")
                strName = strName & "|#" & Nvl(rsTmp("ID")) & ";" & Nvl(rsTmp("������"))
                rsTmp.MoveNext
            Loop
            With Me.vsList
                .ColComboList(13) = strName
            End With
        End With
        RefreshList
    End If
    mbln_Init = False
    Call RefreshTemplet
    
    Call cbo��������_Click
    
    Call RestoreWinState(Me, App.ProductName)                   '����ָ�
End Sub

Private Sub Form_Resize()
    Dim panBase As Pane
    Dim intLoop As Integer
    
    If Me.WindowState = vbMinimized Then Exit Sub
    Set panBase = Me.dkpMan.FindPane(conPane_Base)
    panBase.MinTrackSize.SetSize mlngEditWidth / Screen.TwipsPerPixelX, 265
'    panBase.MaxTrackSize.SetSize mLngEditWidth / Screen.TwipsPerPixelX, 265
    Me.dkpMan.RecalcLayout
    Me.dkpMan.NormalizeSplitters

'    panBase.MinTrackSize.SetSize 0, 0
'    panBase.MaxTrackSize.SetSize mLngEditWidth / Screen.TwipsPerPixelX, 265

    
    
    
End Sub

Private Sub fraMain_DragDrop(Source As Control, x As Single, y As Single)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    mEditState = 0
    mlngKey = 0
    Erase mTestData
    Erase mTestItem
    Erase mTestReagent
'    Call SaveSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "����ID", cbo��������.ItemData(cbo��������.NewIndex))
'    SaveSetting "ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "���Զ���", Me.chk���Զ���.Value & "," & Me.txt��С���Զ���
    
    zlDatabase.SetPara "frmLabMB_����ID", cbo��������.ItemData(cbo��������.NewIndex), 100, 1208
'    zlDatabase.SetPara "frmLabMB_���Զ���", Me.chk���Զ���.Value & "," & Me.txt��С���Զ���, 100, 1208
    mblnMBSelect = False
End Sub


Private Sub opt����_Click(Index As Integer)
    Call RefreshList(1)
End Sub

Private Sub opt��ʾ_Click(Index As Integer)
    Dim intRow As Integer, intCol As Integer
    For intRow = 1 To 8
        For intCol = 1 To 12
            With Me.vsList
                .TextMatrix(intRow, intCol) = mTestData(Index, intRow, intCol)
                                
                If InStr(mTestData(0, intRow, intCol), "BC") > 0 Then
                    '�հ�
                    .Cell(flexcpFontBold, intRow, intCol) = True
                    .Cell(flexcpForeColor, intRow, intCol) = conFontColor_BC
                ElseIf InStr(mTestData(0, intRow, intCol), "NC") > 0 Then
                    '����
                    .Cell(flexcpFontBold, intRow, intCol) = True
                    .Cell(flexcpForeColor, intRow, intCol) = conFontColor_NC
                ElseIf InStr(mTestData(0, intRow, intCol), "PC") > 0 Then
                    '����
                    .Cell(flexcpFontBold, intRow, intCol) = True
                    .Cell(flexcpForeColor, intRow, intCol) = conFontColor_PC
                ElseIf InStr(mTestData(0, intRow, intCol), "QC") > 0 Then
                    '�ʿ�
                    .Cell(flexcpFontBold, intRow, intCol) = True
                    .Cell(flexcpForeColor, intRow, intCol) = conFontColor_QC
                Else
                    If InStr(mTestData(3, intRow, intCol), "+") > 0 Then
                        .Cell(flexcpFontBold, intRow, intCol) = True
                        .Cell(flexcpForeColor, intRow, intCol) = conFontColor_YR
                    ElseIf InStr(mTestData(3, intRow, intCol), "��") > 0 Then
                        .Cell(flexcpFontBold, intRow, intCol) = True
                        .Cell(flexcpForeColor, intRow, intCol) = conFontColor_YL
                    Else
                        .Cell(flexcpFontBold, intRow, intCol) = False
                        .Cell(flexcpForeColor, intRow, intCol) = conFontColor_BK
                    End If
                End If
                
            End With
        Next
    Next
                
End Sub

Private Sub picList_Resize()
    With Me.rptList
        .Left = 10
        .Width = Me.PicList.ScaleWidth - 20
        .Height = Me.PicList.ScaleHeight - .Top - 20
        .Top = Me.opt����(0).Top + Me.opt����(1).Height + 10
    End With
End Sub

Private Sub picMain_Resize()
    Dim intLoop As Integer
    
    On Error Resume Next

    With Me.fra��������
        .Width = Me.PicMain.ScaleWidth - .Left - 50
    End With
    
    With Me.fraģ��
        .Width = Me.PicMain.ScaleWidth - .Left - 50
    End With
    
    With Me.fra΢�װ�
        .Width = Me.PicMain.ScaleWidth - .Left - 50
        .Height = Me.PicMain.ScaleHeight - .Top - 50
    End With
    
    With Me.fra����
        .Top = Me.fra΢�װ�.Height - .Height - 50
    End With
    
    With Me.vsList
        .Width = Me.fra΢�װ�.Width - .Left - 50
        .Height = Me.fra����.Top - .Top
    End With
    
    With Me.vsList
        
        For intLoop = 0 To .Rows - 1
            .RowHeight(intLoop) = (.Height - 9 * 15 - 300) / 8
        Next
        
        For intLoop = 0 To .Cols - 1
            .ColWidth(intLoop) = (.Width - 14 * 15 - 300 - 2000) / 12
        Next
        
        .ColWidth(0) = 300
        .RowHeight(0) = 300
        .ColWidth(13) = 2000
    End With
    
    With Me.txt���԰��
        .Width = Me.fra��������.Width - .Left - 100
    End With
    
    With Me.txt�Լ�����
        .Width = Me.txt���԰��.Width - Me.cmdSl.Width
        Me.cmdSl.Left = .Left + .Width
        Me.cmdSl.Top = .Top
    End With
    
    With Me.txt�Լ�Ч��
        .Width = Me.txt���԰��.Width
    End With
    
    With Me.txt�Լ�����
        .Width = Me.txt���԰��.Width
    End With
    
    With Me.txt���Է���
        .Width = Me.txt���԰��.Width
    End With
    
    With Me.txtCutOff��ʽ
        .Width = Me.txt���԰��.Width
    End With
End Sub

Private Sub rptList_SelectionChanged()
    If Me.rptList.FocusedRow Is Nothing Then Exit Sub
    mlngKey = Me.rptList.FocusedRow.Record(mCol.ID).Value
    Erase mTestData
    Erase mTestItem
    RefreshItem mlngKey
    
End Sub

Private Sub txtCutOff��ʽ_KeyPress(KeyAscii As Integer)
    Dim intRow  As Integer
    On Error Resume Next
    If KeyAscii = 13 Then
        mrsCalc.filter = "������ĿID=" & Me.cbo������Ŀ.ItemData(Me.cbo������Ŀ.ListIndex)
        If mrsCalc.EOF = False Then
            mrsCalc("CutOff��ʽ") = Me.txtCutOff��ʽ
            mrsCalc.Update
        End If
        mTestItem(2, Me.vsList.Row) = Me.txtCutOff��ʽ
        With Me.vsList
            For intRow = 1 To 8
                If .TextMatrix(intRow, 13) = Me.cbo������Ŀ.ItemData(Me.cbo������Ŀ.ListIndex) Then
                    mTestItem(2, intRow) = Me.txtCutOff��ʽ
                End If
            Next
        End With
        MsgBox "�޸Ĺ�ʽ�ɹ�!", vbInformation, Me.Caption
    End If
End Sub

Private Sub txt��ʼ�걾��_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call subWriteNumber
    End If
End Sub

Private Sub txt�����Թ�ʽ_KeyPress(KeyAscii As Integer)
    Dim intRow  As Integer
    On Error Resume Next
    If KeyAscii = 13 Then
        mTestItem(1, Me.vsList.Row) = Me.txt�����Թ�ʽ
        mrsCalc.filter = "������ĿID=" & Me.cbo������Ŀ.ItemData(Me.cbo������Ŀ.ListIndex)
        If mrsCalc.EOF = False Then
            mrsCalc("�����Թ�ʽ") = Me.txt�����Թ�ʽ
            mrsCalc.Update
        End If
        With Me.vsList
            For intRow = 1 To 8
                If .TextMatrix(intRow, 13) = Me.cbo������Ŀ.ItemData(Me.cbo������Ŀ.ListIndex) Then
                    mTestItem(1, intRow) = Me.txt�����Թ�ʽ
                End If
            Next
        End With
        MsgBox "�޸Ĺ�ʽ�ɹ�!", vbInformation, Me.Caption
    End If
End Sub

Private Sub txt�Լ�����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call SelectBatch
    End If
End Sub

Private Sub txt�Լ�����_Validate(Cancel As Boolean)
    If txt�Լ�����.Text <> "" Then
        Call SelectBatch
    Else
        txt�Լ�����.Tag = ""
        txt�Լ�Ч�� = ""
        txt�Լ����� = ""
        txt���Է��� = ""
    End If
End Sub

Private Sub txt���Թ�ʽ_KeyPress(KeyAscii As Integer)
    Dim intRow  As Integer
    On Error Resume Next
    If KeyAscii = 13 Then
        mrsCalc.filter = "������ĿID=" & Me.cbo������Ŀ.ItemData(Me.cbo������Ŀ.ListIndex)
        If mrsCalc.EOF = False Then
            mrsCalc("���Թ�ʽ") = Me.txt���Թ�ʽ
            mrsCalc.Update
        End If
        mTestItem(0, Me.vsList.Row) = Me.txt���Թ�ʽ
        With Me.vsList
            For intRow = 1 To 8
                If .TextMatrix(intRow, 13) = Me.cbo������Ŀ.ItemData(Me.cbo������Ŀ.ListIndex) Then
                    mTestItem(0, intRow) = Me.txt���Թ�ʽ
                End If
            Next
        End With
        MsgBox "�޸Ĺ�ʽ�ɹ�!", vbInformation, Me.Caption
    End If
End Sub

Private Sub vsList_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim intLoop As Integer
    Dim intRow As Integer
    If mEditState = 0 Then Exit Sub
    If Col = 13 And Row > 0 Then
        Erase mTestItem
        With Me.vsList
            For intLoop = 1 To Me.vsList.Rows - 1
                If Me.opt���嵥�� = True Then
                    .TextMatrix(intLoop, 13) = .TextMatrix(Row, Col)
                End If
                Me.vsList.Row = intLoop
                For intRow = 0 To Me.cbo������Ŀ.ListCount - 1
                    If Val(.TextMatrix(intLoop, 13)) = Me.cbo������Ŀ.ItemData(intRow) Then
                        Me.cbo������Ŀ.ListIndex = intRow
                        Call cbo������Ŀ_Click
                        mTestItem(0, intLoop) = Me.txt���Թ�ʽ
                        mTestItem(1, intLoop) = Me.txt�����Թ�ʽ
                        mTestItem(2, intLoop) = Me.txtCutOff��ʽ
                    End If
                Next
            Next
        End With
    End If
    
    For intLoop = 0 To Me.opt��ʾ.UBound
        If opt��ʾ(intLoop).Value = True Then
            Exit For
        End If
    Next
    If Row > 0 And Col > 0 And Col < 13 Then
        mTestData(intLoop, Row, Col) = Me.vsList.TextMatrix(Row, Col)
    End If
End Sub

Private Sub vsList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call CalcData
End Sub

Private Sub vsList_Click()
    Dim lngKey As Long
    Dim intLoop As Integer
    With Me.vsList
        If .Row > 0 Then
            lngKey = Val(.TextMatrix(.Row, 13))
        End If
    End With
    If lngKey <> 0 Then
        With Me.cbo������Ŀ
            For intLoop = 0 To .ListCount - 1
                If .ItemData(intLoop) = lngKey Then
                    .ListIndex = intLoop
                    Me.txt���Թ�ʽ.Text = mTestItem(0, Me.vsList.Row)
                    Me.txt�����Թ�ʽ.Text = mTestItem(1, Me.vsList.Row)
                    Me.txtCutOff��ʽ.Text = mTestItem(2, Me.vsList.Row)
                    CalcData .ItemData(intLoop)
                End If
            Next
        End With
    End If
    On Error Resume Next
    If Me.vsList.Row > 0 And Me.vsList.Row < 9 Then
        If mTestReagent(Me.vsList.Row) <> "" Then
            mblnRefresh = True
            txt�Լ�����.Text = Split(mTestReagent(Me.vsList.Row), ";")(0)
            txt�Լ�Ч��.Text = Split(mTestReagent(Me.vsList.Row), ";")(1)
            txt�Լ�����.Text = Split(mTestReagent(Me.vsList.Row), ";")(2)
            txt���Է���.Text = Split(mTestReagent(Me.vsList.Row), ";")(3)
            mblnRefresh = False
        Else
            txt�Լ����� = ""
        End If
    End If
    mblnMBSelect = False
End Sub

Private Sub vsList_DblClick()
    Dim strMaxNumber As String
    If mEditState = 0 Then Exit Sub
    With Me.vsList
        If .Row > 0 And .Col > 0 And .Col < 13 And Me.opt��ʾ(0).Value = True Then
            strMaxNumber = GetMaxNumber
            .TextMatrix(.Row, .Col) = strMaxNumber
            If InStr(strMaxNumber, "BC") > 0 Then
                '�հ�
                .Cell(flexcpFontBold, .Row, .Col) = True
                .Cell(flexcpForeColor, .Row, .Col) = conFontColor_BC
            ElseIf InStr(strMaxNumber, "NC") > 0 Then
                '����
                .Cell(flexcpFontBold, .Row, .Col) = True
                .Cell(flexcpForeColor, .Row, .Col) = conFontColor_NC
            ElseIf InStr(strMaxNumber, "PC") > 0 Then
                '����
                .Cell(flexcpFontBold, .Row, .Col) = True
                .Cell(flexcpForeColor, .Row, .Col) = conFontColor_PC
            ElseIf InStr(strMaxNumber, "QC") > 0 Then
                '�ʿ�
                .Cell(flexcpFontBold, .Row, .Col) = True
                .Cell(flexcpForeColor, .Row, .Col) = conFontColor_QC
            End If
            
        End If
    End With
End Sub

Private Sub vsList_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim intX As Integer, intY As Integer
    If KeyCode = vbKeyDelete Then
        If MsgBox("�Ƿ�ȷ��Ҫɾ��ѡ�еı��?", vbYesNo + vbDefaultButton2 + vbQuestion, Me.Caption) = vbNo Then Exit Sub
        With vsList
            For intY = .Row To .RowSel
                For intX = .Col To .ColSel
                    .TextMatrix(intY, intX) = ""
                    mTestData(0, intY, intX) = ""
                    mTestData(1, intY, intX) = ""
                    mTestData(2, intY, intX) = ""
                    mTestData(3, intY, intX) = ""
                Next
            Next
        End With
    End If
End Sub

Private Sub vsList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If mEditState = 0 Then Exit Sub
    Call subWriteNumber
End Sub

Private Sub subWriteNumber()
    '����           ��һ����д��걾���
    Dim intX As Integer, intY As Integer
    Dim intLoop As Integer
    
    If Val(Me.txt��ʼ�걾��.Text) = 0 Then Exit Sub
    If Me.opt��ʾ(0).Value = False Then Exit Sub
    
    '��ģ��ѡ��ʱֱ��ʹ��ģ��������
    If mblnMBSelect = True Then
        If Me.cboѡ��ģ��.Text <> "" Then cboѡ��ģ��_Click: Exit Sub
    End If
    
    intLoop = Val(Me.txt��ʼ�걾��.Text) - 1
    With Me.vsList
        If Me.opt����(0).Value = True Then
            For intY = .Row To .RowSel
                If Me.opt�������.Value = True Then intLoop = Val(Me.txt��ʼ�걾��.Text) - 1
                For intX = .Col To .ColSel
                    If intY > 0 And intX > 0 And intX < 13 Then
                        intLoop = intLoop + 1
                        .TextMatrix(intY, intX) = intLoop
                        mTestData(0, intY, intX) = intLoop
                        .Cell(flexcpData, intY, intX) = intLoop
                        .Cell(flexcpFontBold, intY, intX) = False
                        .Cell(flexcpForeColor, intY, intX) = vbBlack
                    End If
                Next
            Next
        Else
            For intX = .Col To .ColSel
                If Me.opt�������.Value = True Then intLoop = Val(Me.txt��ʼ�걾��.Text) - 1
                For intY = .Row To .RowSel
                    If intY > 0 And intX > 0 And intX < 13 Then
                        intLoop = intLoop + 1
                        .TextMatrix(intY, intX) = intLoop
                        mTestData(0, intY, intX) = intLoop
                        .Cell(flexcpData, intY, intX) = intLoop
                        .Cell(flexcpFontBold, intY, intX) = False
                        .Cell(flexcpForeColor, intY, intX) = vbBlack
                    End If
                Next
            Next
        End If
    End With
    Me.txt��ʼ�걾�� = ""
End Sub
Private Function GetMaxNumber() As String
    '����ȡ���б��е������
    Dim intLoop As Integer
    Dim intY  As Integer, intX As Integer
    Dim intMax As Integer
    Dim strTmp As String
    Dim strType As String
    
    '�õ�������
    For intLoop = 0 To Me.opt��ѡ��.UBound
        If Me.opt��ѡ��(intLoop).Value = True Then
            Exit For
        End If
    Next
    
    Select Case intLoop
        Case 0  '��ͨ
            strType = "S"
        Case 1  '�հ�
            strType = "BC"
        Case 2  '����
            strType = "NC"
        Case 3  '����
            strType = "PC"
        Case 4  '�ʿ�
            strType = "QC"
    End Select
    
    With Me.vsList
        For intY = 1 To .Rows - 1
            For intX = 1 To .Cols - 2
                If strType = "S" Then
                    strTmp = Val(.TextMatrix(intY, intX))
                    If Val(strTmp) <> 0 Then
                        If CInt(strTmp) >= intMax Then intMax = CInt(strTmp) + 1
                    End If
                Else
                    If InStr(.TextMatrix(intY, intX), strType) > 0 Then
                        strTmp = Val(Trim(Replace(.TextMatrix(intY, intX), strType, "")))
                        If CInt(strTmp) >= intMax Then intMax = CInt(strTmp) + 1
                    End If
                End If
            Next
        Next
    End With
    
    GetMaxNumber = Replace(strType, "S", "") & IIf(intMax = 0, 1, intMax)
End Function
Private Sub SaveTemplet()
    ''''''''''''''''''''''''''''''''
    '����   ���浽ģ��
    ''''''''''''''''''''''''''''''''
    Dim intY As Integer, intX As Integer
    Dim intRow As Integer, intCol As Integer
    Dim strNumber As String
    Dim intLoop As Integer
    Dim strResult As String
    
    '�걾�ű�ż��
    For intRow = 1 To 8
        For intCol = 1 To 12
            With Me.vsList
                strNumber = .TextMatrix(intRow, intCol)
                If Trim(strNumber) <> "" Then
                    If IsNumeric(strNumber) = True Then
                        If Len(strNumber) > 4 Then
                            MsgBox "�걾������ֻ��Ϊ<9999>�����޸ģ�", vbInformation
                            .Select intRow, intCol
                            Exit Sub
                        End If
                    Else
                        If InStr(strNumber, "BC") = 0 And InStr(strNumber, "NC") = 0 And InStr(strNumber, "PC") = 0 And InStr(strNumber, "QC") = 0 Then
                            MsgBox "�������˲���ȷ�ı��<" & strNumber & ">������!", vbInformation
                            .Select intRow, intCol
                            Exit Sub
                        End If
                    End If
                    For intY = intRow To 8
                        For intX = intCol + 1 To 12
                            If strNumber = .TextMatrix(intY, intX) Then
                                MsgBox "������ͬ�ı�ţ����޸ģ�", vbInformation
                                .Select intY, intX
                                Exit Sub
                            End If
                        Next
                    Next
                    intLoop = intLoop + 1
                End If
            End With
        Next
    Next
    If intLoop = 0 Then
        MsgBox "û��ѡ���Ų��ܱ���!", vbInformation
        Exit Sub
    End If
    
    '��֯��������
    
    '��Ŀ
    For intRow = 1 To 8
        With Me.vsList
            strNumber = .TextMatrix(intRow, 13)
            If Trim(strNumber) <> "" Then
                intLoop = intLoop + 1
            End If
            If intRow = 1 Then
                strResult = strResult & strNumber
            Else
                strResult = strResult & ";" & strNumber
            End If
            
        End With
    Next
    
    If intLoop < 8 And Me.opt�������.Value = True Then
        MsgBox "��ѡ���˵������Ŀ����������Ŀû��ѡ��", vbInformation
        Me.vsList.Select 1, 13
        Exit Sub
    End If
    
    If strResult = ";;;;;;;" Then
        MsgBox "û��ѡ�������Ŀ����ѡ������Ŀ!", vbInformation
        Me.vsList.Select 1, 13
        Exit Sub
    End If
    
    '���
    For intRow = 1 To 8
        strResult = strResult & "|"
        For intCol = 1 To 12
            With Me.vsList
                strNumber = .TextMatrix(intRow, intCol)
                If intCol = 1 Then
                    strResult = strResult & strNumber
                Else
                    strResult = strResult & ";" & strNumber
                End If
            End With
        Next
    Next
    
    frmLabMBTemplet.ShowMe Me, strResult
    RefreshTemplet
End Sub

Private Sub RefreshTemplet()
    '����   ˢ�µ�ǰģ��
    Dim rsTmp As New adodb.Recordset
    'д��ø��ģ������
    gstrSql = "select id,���,���� from ����ø��ģ�� order by ���"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    Me.cboѡ��ģ��.Clear
    Me.cboѡ��ģ��.AddItem ""
    Me.cboѡ��ģ��.ItemData(Me.cboѡ��ģ��.NewIndex) = 0
    Do Until rsTmp.EOF
        With Me.cboѡ��ģ��
            .AddItem rsTmp("���") & "-" & rsTmp("����")
            .ItemData(.NewIndex) = rsTmp("ID")
        End With
        rsTmp.MoveNext
    Loop
End Sub

Private Sub MBcontrol()
    '���� ��ø����ƽ���
    Dim strControl As String
    Dim strResult As String
    Dim intRow As Integer, intCol As Integer
    Dim aRow() As String, aCol() As String
    
    If Me.cbo��������.Text = "" Then
        MsgBox "��ѡ��һ������!", vbInformation, Me.Caption
        Me.cbo��������.SetFocus
        Exit Sub
    End If
    
    If Me.cbo����.Text = "" Then
        MsgBox "��ѡ�񲨳�!    ", vbInformation, Me.Caption
        Me.cbo����.SetFocus
        Exit Sub
    End If

    If Me.cbo���Ƶ��.Text = "" Then
        MsgBox "��ѡ�����Ƶ��!", vbInformation, Me.Caption
        Me.cbo���Ƶ��.SetFocus
        Exit Sub
    End If

    If Me.txt���ʱ��.Text = "" Then
        MsgBox "��ѡ�����ʱ��!", vbInformation, Me.Caption
        Me.cbo���Ƶ��.SetFocus
        Exit Sub
    End If

    If Me.cbo���巽ʽ.Text = "" Then
        MsgBox "��ѡ����巽ʽ!", vbInformation, Me.Caption
        Me.cbo���Ƶ��.SetFocus
        Exit Sub
    End If

    If Me.cbo�հ���ʽ.Text = "" Then
        MsgBox "��ѡ��հ���ʽ!", vbInformation, Me.Caption
        Me.cbo���Ƶ��.SetFocus
        Exit Sub
    End If
    
    Me.opt��ʾ(0).Value = True
    For intRow = 1 To 8
        For intCol = 1 To 12
            mTestData(0, intRow, intCol) = Me.vsList.TextMatrix(intRow, intCol)
        Next
    Next
    strControl = Me.cbo����.Text & ";" & Me.cbo���Ƶ��.Text & ";" & Me.txt���ʱ�� & _
                 ";" & Me.cbo���巽ʽ.Text & ";" & Me.cbo�հ���ʽ & ";" & Me.cbo�ο�����.Text
                 
    frmLabMBControl.ShowMe Me, strControl, strResult
    
    If strResult = "" Then MsgBox "û�вɼ������ݣ������²���!": Exit Sub
    
    aRow = Split(strResult, "|")
    For intRow = 1 To 8
        aCol = Split(aRow(intRow - 1), ";")
        For intCol = 1 To 12
            With Me.vsList
                If Trim(.TextMatrix(intRow, intCol)) <> "" Then
                    .Cell(flexcpData, intRow, intCol, intRow, intCol) = aCol(intCol - 1)
                    .TextMatrix(intRow, intCol) = Format(aCol(intCol - 1), "##0.000#")
                    mTestData(1, intRow, intCol) = Format(aCol(intCol - 1), "##0.000#")
                End If
            End With
        Next
    Next
    Me.opt��ʾ(1).Value = True
    Me.cboѡ��ģ��.ListIndex = 0
    '������ɾͼ���
    Call CalcData
End Sub
Private Sub AddNew()
    '����           ����һ���°�
    Dim rsTmp As New adodb.Recordset
    
    'û��ѡ������ʱ�˳�
    If Me.cbo��������.Text = "" Then MsgBox "����ѡ������!", vbInformation: Me.cbo��������.SetFocus: Exit Sub
'    ReDim mTestData(0 To 3, 1 To 8, 1 To 12)
    Erase mTestData
    Erase mTestItem
    Erase mTestReagent
    
    mEditState = 1
    
    Me.opt��ʾ(0).Value = True
    Call InitItem
    gstrSql = "select count(*) +1  from ����ø���¼ where ����ʱ�� between to_date(to_char(sysdate,'yyyy-MM-dd') || ' 00:00:00','yyyy-MM-dd HH24:mi:ss')" & vbNewLine & _
                "              and to_date(to_char(sysdate,'yyyy-MM-dd') || ' 23:59:59','yyyy-MM-dd HH24:mi:ss') and ����id = [1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngMachine)
    Me.txt���԰��.Text = Format(Now, "yyyymmdd") & "_" & rsTmp(0)
    Me.cboѡ��ģ��.ListIndex = 0
    mblnMBSelect = False
End Sub
Private Sub SaveData()
    '����           ���浱ǰ��������
    
    Dim intRow As Integer, intCol As Integer
    Dim strNumber As String             '���ڼ�������
    Dim strItem As String               '��Ŀ����
    Dim strData As String               '�������
    Dim strCalcOne As String            '���Թ�ʽ
    Dim strCalcTwo As String            '�����Թ�ʽ
    Dim strCalcThree As String          'CutOff��ʽ
    Dim str�Լ� As String               '�Լ�
    Dim lngKey As Long
    
    Dim rsTmp As New adodb.Recordset
    
    On Error GoTo errH
    
    '��������Ƿ�����
    For intRow = 1 To 8
        strData = strData & "|"
        For intCol = 1 To 13
            If intCol < 13 Then
                strNumber = strNumber & ";" & mTestData(0, intRow, intCol)
                strData = strData & ";" & mTestData(0, intRow, intCol) & "^" & mTestData(1, intRow, intCol)
            Else
                '�б���û����Ŀʱ
'                If Len(strNumber) > 1 And Me.vsList.Cell(flexcpText, intRow, intCol) = "" Then
'                    MsgBox "��ѡ��һ��������Ŀ��ȡ����ǰ�еı���!", vbInformation, Me.Caption
'                    Me.vsList.Select intRow, intCol
'                    Exit Sub
'                End If
                
                '�������Ŀʱ����Ŀû�б���ʱ
                If Len(strNumber) = 1 And Me.vsList.Cell(flexcpText, intRow, intCol) <> "" Then
                    MsgBox "��ѡ��ǰ�еı���!", vbInformation, Me.Caption
                    Me.vsList.Select intRow, 1
                    Exit Sub
                End If
                strItem = strItem & ";" & Me.vsList.Cell(flexcpText, intRow, intCol)
                
                mrsCalc.filter = "������Ŀid=" & Val(Me.vsList.TextMatrix(intRow, intCol))
                
                '���Թ�ʽ
                If mTestItem(0, intRow) <> "" Then
                    strCalcOne = strCalcOne & ";" & mTestItem(0, intRow)
                Else
                    If mrsCalc.EOF = False Then
                        strCalcOne = strCalcOne & ";" & Nvl(mrsCalc("���Թ�ʽ"))
                    Else
                        strCalcOne = strCalcOne & ";"
                    End If
                End If
                
                '�����Թ�ʽ
                If mTestItem(1, intRow) <> "" Then
                    strCalcTwo = strCalcTwo & ";" & mTestItem(1, intRow)
                Else
                    If mrsCalc.EOF = False Then
                        strCalcTwo = strCalcTwo & ";" & Nvl(mrsCalc("�����Թ�ʽ"))
                    Else
                        strCalcTwo = strCalcTwo & ";"
                    End If
                End If
                
                '���Թ�ʽ
                If mTestItem(2, intRow) <> "" Then
                    strCalcThree = strCalcThree & ";" & mTestItem(2, intRow)
                Else
                    If mrsCalc.EOF = False Then
                        strCalcThree = strCalcThree & ";" & Nvl(mrsCalc("CutOff��ʽ"))
                    Else
                        strCalcThree = strCalcThree & ";"
                    End If
                End If
            End If
        Next
    Next
    
    If mEditState = 1 Then
        lngKey = zlDatabase.GetNextId("����ø���¼")
    Else
        lngKey = mlngKey
    End If
    
    str�Լ� = Join(mTestReagent, "|")
    If Replace(str�Լ�, "|", "") = "" Then
        str�Լ� = ""
    End If
    
    '��ʼ����
    gstrSql = "Zl_����ø���¼_Insert(" & lngKey & ",'" & Me.txt���԰�� & "'," & _
                "to_date('" & Me.dtp����ʱ��.Value & "','yyyy-MM-dd HH24:MI:ss')" & ",'" & Me.cbo����.Text & "','" & _
                Me.cbo�ο�����.Text & "','" & Me.cbo���Ƶ��.Text & "','" & Me.txt���ʱ��.Text & "','" & _
                Me.cbo���巽ʽ.Text & "','" & Me.cbo�հ���ʽ.Text & "','" & txt�Լ�����.Tag & "'," & _
                IIf(Me.txt�Լ�Ч�� <> "", "to_date('" & Me.txt�Լ�Ч��.Text & "','yyyy-MM-dd HH:MI:ss')", "NULL") & _
                ",'" & Me.txt�Լ�����.Text & "','" & _
                Me.txt���Է���.Text & "'," & IIf(Me.opt���嵥��.Value = True, 1, 0) & ",'" & Me.txt���λ��.Text & "','" & Mid(strItem, 2) & "','" & _
                strCalcOne & "','" & strCalcTwo & "','" & strCalcThree & "','" & Mid(strData, 2) & "','" & str�Լ� & "'," & _
                Me.cbo��������.ItemData(Me.cbo��������.ListIndex) & ")"
    
    zlDatabase.ExecuteProcedure gstrSql, Me.Caption
    
    '���浱ǰ�������õ�ע��
    On Error Resume Next
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "����", cbo����.Text)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "�ο�����", cbo�ο�����.Text)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "���Ƶ��", cbo���Ƶ��.Text)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "���ʱ��", txt���ʱ��.Text)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "���巽ʽ", cbo���巽ʽ.Text)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "�հ���ʽ", cbo�հ���ʽ.Text)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "��ĿID", Me.cbo������Ŀ.ItemData(Me.cbo������Ŀ.NewIndex))
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "�հ���ʽ", cbo�հ���ʽ.Text)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & App.ProductName, "��������", opt����(0).Value)
    zlDatabase.SetPara "frmLabMB_����", cbo����.Text, 100, 1208
    zlDatabase.SetPara "frmLabMB_�ο�����", cbo�ο�����.Text, 100, 1208
    zlDatabase.SetPara "frmLabMB_���Ƶ��", cbo���Ƶ��.Text, 100, 1208
    zlDatabase.SetPara "frmLabMB_���ʱ��", txt���ʱ��.Text, 100, 1208
    zlDatabase.SetPara "frmLabMB_���巽ʽ", cbo���巽ʽ.Text, 100, 1208
    zlDatabase.SetPara "frmLabMB_�հ���ʽ", cbo�հ���ʽ.Text, 100, 1208
    zlDatabase.SetPara "frmLabMB_��ĿID", Me.cbo������Ŀ.ItemData(Me.cbo������Ŀ.NewIndex), 100, 1208
    On Error GoTo 0
    mEditState = 0
    mlngKey = lngKey
    Call SendData               '���͵���ʦ����վ
    Call RefreshList
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub RefreshList(Optional intType As Integer = 1, Optional strFilter As String)
    '����               ˢ������б�
    '����               intType ˢ�²���(�������ֿ��ٹ��˺͹��ˣ�
    '                       1=���ٹ���
    '                       2=����(�ڶ����еĹ����������й���) ��ʽ:"���;�Լ�����:�Ƿ�ʹ��ʱ���ѯ,��ʼʱ��,����ʱ��"
    '                       3=ͨ��ID����ID��ʹ��","���зָ�"
    '                   Strfilter �����ִ�
    
    Dim rsTmp As New adodb.Recordset
    Dim Record As ReportRecord
    Dim intLoop As Integer
    Dim strBeginDate As String, strEndDate As String
    Dim strDate As String
    Dim aItem() As String
    Dim aRow() As String
    Dim strWhere As String
    Dim strReagentNo As String          '�Լ�����
    Dim strReagentDate As String        '�Լ�Ч��
    Dim strReagentManufacturer          '����
    Dim strReagentMeans                 '����

    
    On Error GoTo errH
    
    gstrSql = "select ID,���,����ʱ��,����,�ο�����,���Ƶ��,���ʱ��,���巽ʽ,�հ���ʽ,�Լ�����,�Լ�Ч��,�Լ�����,���Է���,�Ƿ���," & vbNewLine & _
              "       ���λ��,������Ŀ,���Թ�ʽ,�����Թ�ʽ,CutOff��ʽ,���Խ��,�Լ���¼ from ����ø���¼ a "
                  
    strBeginDate = Format(GetDateTime("��  ��", 1), "yyyy-MM-dd 00:00:00")
    strEndDate = Format(GetDateTime("��  ��", 2), "yyyy-MM-dd 23:59:59")
    If intType = 1 Then
        '���ٹ���
        gstrSql = gstrSql & " Where ����ʱ�� between [1] and [2] "
        If Me.opt����(0).Value = True Then
            '����
            strBeginDate = Format(GetDateTime("��  ��", 1), "yyyy-MM-dd 00:00:00")
            strEndDate = Format(GetDateTime("��  ��", 2), "yyyy-MM-dd 23:59:59")
        ElseIf Me.opt����(1).Value = True Then
            '����
            strBeginDate = Format(GetDateTime("��  ��", 1), "yyyy-MM-dd 00:00:00")
            strEndDate = Format(GetDateTime("��  ��", 2), "yyyy-MM-dd 23:59:59")
        ElseIf Me.opt����(2).Value = True Then
            '����
            strBeginDate = Format(GetDateTime("��  ��", 1), "yyyy-MM-dd 00:00:00")
            strEndDate = Format(GetDateTime("��  ��", 2), "yyyy-MM-dd 23:59:59")
        ElseIf Me.opt����(3).Value = True Then
            '����
            strBeginDate = Format(GetDateTime("��  ��", 1), "yyyy-MM-dd 00:00:00")
            strEndDate = Format(GetDateTime("��  ��", 2), "yyyy-MM-dd 23:59:59")
        End If
        gstrSql = gstrSql & " And ����ID = [3] "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CDate(strBeginDate), CDate(strEndDate), _
        Me.cbo��������.ItemData(Me.cbo��������.ListIndex))
    ElseIf intType = 2 Then
        '����
        aItem = Split(strFilter, ";")
        
        If aItem(0) <> "" Then
            '���
            strWhere = " where ��� = [3] "
        End If
                
        If aItem(1) <> "" Then
            '�Լ�����
            If strWhere = "" Then
                strWhere = " Where �Լ����� = [4] "
            Else
                strWhere = strWhere & " And �Լ����� = [4] "
            End If
        End If
        
        If Split(aItem(2), ",")(0) = 1 Then
            '�Լ�����
            If strWhere = "" Then
                strWhere = " Where ����ʱ�� between [1] and [2] "
            Else
                strWhere = strWhere & " And ����ʱ�� between [1] and [2] "
            End If
            strBeginDate = Split(aItem(2), ",")(1)
            strEndDate = Split(aItem(2), ",")(2)
        End If
        
        gstrSql = gstrSql & strWhere
        
        
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CDate(strBeginDate), CDate(strEndDate), aItem(0), aItem(1), _
                    Me.cbo��������.ItemData(Me.cbo��������.ListIndex))
    ElseIf intType = 3 Then
        If strFilter <> "" Then
            gstrSql = gstrSql & " , (Select * From Table(Cast(f_str2list([1]) As zltools.t_strlist))) H where a.id = h.Column_Value"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CDate(strBeginDate), CDate(strEndDate), aItem(0), aItem(1))
        End If
    End If
    
    Me.rptList.Records.DeleteAll
    Do Until rsTmp.EOF
        Set Record = Me.rptList.Records.Add
        
        For intLoop = 0 To Me.rptList.Columns.Count
            Record.AddItem ""
        Next
        
        Record.Item(mCol.ID).Value = Nvl(rsTmp("ID"))
        Record.Item(mCol.���).Value = Nvl(rsTmp("���"))
        Record.Item(mCol.����ʱ��).Value = Nvl(rsTmp("����ʱ��"))
'        Record.Item(mCol.�Լ�����).Value = Nvl(rsTmp("�Լ�����"))
'        Record.Item(mCol.�Լ�Ч��).Value = Nvl(rsTmp("�Լ�Ч��"))
'        Record.Item(mCol.�Լ�����).Value = Nvl(rsTmp("�Լ�����"))
'        Record.Item(mCol.���Է���).Value = Nvl(rsTmp("���Է���"))
        Record.Item(mCol.����).Value = Nvl(rsTmp("����"))
        Record.Item(mCol.�ο�����).Value = Nvl(rsTmp("�ο�����"))
        Record.Item(mCol.���Ƶ��).Value = Nvl(rsTmp("���Ƶ��"))
        Record.Item(mCol.���ʱ��).Value = Nvl(rsTmp("���ʱ��"))
        Record.Item(mCol.���巽ʽ).Value = Nvl(rsTmp("���巽ʽ"))
        Record.Item(mCol.�հ���ʽ).Value = Nvl(rsTmp("�հ���ʽ"))
        Record.Item(mCol.���λ��).Value = Nvl(rsTmp("���λ��"))
        Record.Item(mCol.������Ŀ).Value = Nvl(rsTmp("������Ŀ"))
        Record.Item(mCol.���Թ�ʽ).Value = Nvl(rsTmp("���Թ�ʽ"))
        Record.Item(mCol.�����Թ�ʽ).Value = Nvl(rsTmp("�����Թ�ʽ"))
        Record.Item(mCol.CutOff��ʽ).Value = Nvl(rsTmp("CutOff��ʽ"))
        Record.Item(mCol.���Խ��).Value = Nvl(rsTmp("���Խ��"))
        Record.Item(mCol.�Լ���¼).Value = Nvl(rsTmp("�Լ���¼"))
        
        If Replace(Record.Item(mCol.�Լ���¼).Value, "|", "") <> "" Then
            strReagentNo = "": strReagentDate = "": strReagentManufacturer = "": strReagentMeans = ""
            'д���Լ���¼
            aRow = Split(Record.Item(mCol.�Լ���¼).Value, "|")
            
            For intLoop = 0 To UBound(aRow)
                aItem = Split(aRow(intLoop), ";")
                If UBound(aItem) >= 3 Then
                    '�Լ�
                    If InStr(strReagentNo, "<" & aItem(0) & ">") <= 0 Then
                        strReagentNo = strReagentNo & "<" & aItem(0) & ">"
                    End If
                    'Ч��
                    If InStr(strReagentDate, "<" & aItem(1) & ">") <= 0 Then
                        strReagentDate = strReagentDate & "<" & aItem(1) & ">"
                    End If
                    '����
                    If InStr(strReagentManufacturer, "<" & aItem(2) & ">") <= 0 Then
                        strReagentManufacturer = strReagentManufacturer & "<" & aItem(2) & ">"
                    End If
                    '����
                    If InStr(strReagentMeans, "<" & aItem(3) & ">") <= 0 Then
                        strReagentMeans = strReagentMeans & "<" & aItem(3) & ">"
                    End If
                End If
            Next
            Record.Item(mCol.�Լ�����).Value = strReagentNo
            Record.Item(mCol.�Լ�Ч��).Value = strReagentDate
            Record.Item(mCol.�Լ�����).Value = strReagentManufacturer
            Record.Item(mCol.���Է���).Value = strReagentMeans
        End If
        rsTmp.MoveNext
    Loop
    Me.rptList.Populate
    If mlngKey = 0 Then
        If Me.rptList.Rows.Count > 0 Then
            Call RefreshItem(Me.rptList.Rows(0).Record(mCol.ID).Value)
        End If
    Else
        Call RefreshItem(mlngKey)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub RefreshItem(lngKey As Long)
    '����           ˢ�µ�ǰ��Ŀ
    Dim rsTmp As New adodb.Recordset
    Dim lngLoop As Long, intLoop As Integer
    Dim aItem() As String
    Dim aRow() As String, aCol() As String
    Dim intRow As Integer, intCol As Integer
    Dim aRule() As String
    
    On Error GoTo errH
    
    Call InitItem
    Erase mTestItem
    gstrSql = "select ������ĿID,���Թ�ʽ,�����Թ�ʽ,CutOff��ʽ from ������Ŀ where ��Ŀ��� = 4"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    
    With Me.rptList
        For lngLoop = 0 To .Rows.Count - 1
            If .Rows(lngLoop).Record(mCol.ID).Value = lngKey Then
                '�ҵ�ID��д��
                On Error Resume Next
                dtp����ʱ�� = .Rows(lngLoop).Record(mCol.����ʱ��).Value
                cbo����.Text = .Rows(lngLoop).Record(mCol.����).Value
                cbo�ο�����.Text = .Rows(lngLoop).Record(mCol.�ο�����).Value
                cbo���Ƶ��.Text = .Rows(lngLoop).Record(mCol.���Ƶ��).Value
                txt���ʱ��.Text = .Rows(lngLoop).Record(mCol.���ʱ��).Value
                cbo���巽ʽ.Text = .Rows(lngLoop).Record(mCol.���巽ʽ).Value
                cbo�հ���ʽ.Text = .Rows(lngLoop).Record(mCol.�հ���ʽ).Value
                txt���԰��.Text = .Rows(lngLoop).Record(mCol.���).Value
'                cbo�Լ�����.Text = .Rows(lngLoop).Record(mCol.�Լ�����).Value
'                txt�Լ�Ч��.Text = .Rows(lngLoop).Record(mCol.�Լ�Ч��).Value
'                txt�Լ�����.Text = .Rows(lngLoop).Record(mCol.�Լ�����).Value
'                txt���Է���.Text = .Rows(lngLoop).Record(mCol.���Է���).Value
                txtCutOff��ʽ.Text = .Rows(lngLoop).Record(mCol.CutOff��ʽ).Value
                txt���λ��.Text = .Rows(lngLoop).Record(mCol.���λ��).Value
                
                
                
                aItem = Split(.Rows(lngLoop).Record(mCol.������Ŀ).Value, ";")
                For intLoop = 0 To UBound(aItem)
                    Me.vsList.TextMatrix(intLoop + 1, 13) = aItem(intLoop)
                    If intLoop > 0 Then
                        If Me.vsList.TextMatrix(intLoop, 13) <> Me.vsList.TextMatrix(intLoop + 1, 13) Then
                            '�ж��Ƿ���һ�����Ŀ
                            Me.opt�������.Value = True
                        End If
                    End If
                Next
                
                'ȡ������Ŀ�еĹ�ʽ
                aItem = Split(.Rows(lngLoop).Record(mCol.������Ŀ).Value, ";")
                
                
                For intRow = 1 To 8
                    '���Թ�ʽ
                    aRule = Split(Mid(.Rows(lngLoop).Record(mCol.���Թ�ʽ).Value, 2), ";")
                    mTestItem(0, intRow) = aRule(intRow - 1)
                    
                    '�����Թ�ʽ
                    aRule = Split(Mid(.Rows(lngLoop).Record(mCol.�����Թ�ʽ).Value, 2), ";")
                    mTestItem(1, intRow) = aRule(intRow - 1)
                    
                    'CutOff��ʽ
                    aRule = Split(Mid(.Rows(lngLoop).Record(mCol.CutOff��ʽ).Value, 2), ";")
                    mTestItem(2, intRow) = aRule(intRow - 1)
                    
                    rsTmp.filter = "������Ŀid=" & aItem(intRow - 1)
                    If rsTmp.EOF = False Then
                        For intLoop = 0 To Me.cbo������Ŀ.ListCount - 1
                            If Me.cbo������Ŀ.ItemData(intLoop) = rsTmp("������ĿID") Then
                                Me.vsList.Row = intRow
                                Me.cbo������Ŀ.ListIndex = intLoop
                            End If
                        Next
                    End If
                Next
                    

                On Error GoTo 0
'                If .Rows(lngLoop).Record(mCol.���Թ�ʽ).Value <> "" Then
'                    txt���Թ�ʽ.Text = Split(.Rows(lngLoop).Record(mCol.���Թ�ʽ).Value, ";")(1)
'                End If
'                If .Rows(lngLoop).Record(mCol.�����Թ�ʽ).Value <> "" Then
'                    txt�����Թ�ʽ.Text = Split(.Rows(lngLoop).Record(mCol.�����Թ�ʽ).Value, ";")(1)
'                End If
'                If .Rows(lngLoop).Record(mCol.CutOff��ʽ).Value <> "" Then
'                    txtCutOff��ʽ.Text = Split(.Rows(lngLoop).Record(mCol.CutOff��ʽ).Value, ";")(1)
'                End If
                
                
                aRow = Split(.Rows(lngLoop).Record(mCol.���Խ��).Value, "|")
                
                For intRow = 1 To 8
                    aCol = Split(aRow(intRow - 1), ";")
                    For intCol = 1 To 12
                        With Me.vsList
                            mTestData(0, intRow, intCol) = Split(aCol(intCol), "^")(0)
                            mTestData(1, intRow, intCol) = Split(aCol(intCol), "^")(1)
                        End With
                    Next
                Next
                For intLoop = 0 To opt��ʾ.Count - 1
                    If opt��ʾ(intLoop).Value = True Then
                        Call opt��ʾ_Click(intLoop)
                    End If
                Next
                Erase mTestReagent
                aItem = Split(.Rows(lngLoop).Record(mCol.�Լ���¼).Value, "|")
                For intLoop = 0 To UBound(aItem)
                    mTestReagent(intLoop + 1) = aItem(intLoop)
                Next
            End If
        Next
    End With
            
    Call CalcData
    Me.vsList.Row = 1
    Call vsList_Click
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub DelData()
    '����           ɾ������
    
    If Me.rptList.FocusedRow Is Nothing Then Exit Sub
    
    If MsgBox("�Ƿ�ȷ��Ҫɾ�����Ϊ<" & Me.rptList.FocusedRow.Record(mCol.���).Value & ">�Ľ��!", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
        Exit Sub
    End If
    
    On Error GoTo errH
    gstrSql = "Zl_����ø���¼_Delete(" & Me.rptList.FocusedRow.Record(mCol.ID).Value & ")"
    zlDatabase.ExecuteProcedure gstrSql, Me.Caption
    
    RefreshList
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
    
Private Sub CalcData(Optional OneCalc As Long)
    '����           ����հ�(BC)������(NC)������(PC)���ʿ�(QC)
    Dim intLoop As Integer, lngLoop As Long
    Dim strItem As String
    Dim intRow As Integer, intCol As Integer
    Dim strBC As String, strNC As String, strPC As String, strQC As String
    Dim aBC() As String, aNC() As String, aPC() As String, aQC() As String
    Dim dblBC As Double, dblNC As Double, dblPC As Double, dblQC As Double
    Dim str���� As String, str������ As String, strCutOff As String
    Dim rsTmp As New adodb.Recordset
    Dim aItem() As String
    Dim bln���հ׶��� As Boolean
    Dim blnС�����Զ��� As Boolean
    Dim str���Զ��� As String
    Dim intCount As Integer
    
    On Error GoTo errH
    
    '��������ODֵ
    If mstr��ʽ <> "" Then
        For intRow = 1 To 8
            For intCol = 1 To 12
                If mTestData(1, intRow, intCol) <> "" Then
                    mTestData(1, intRow, intCol) = Calc.Eval(Replace(UCase(mstr��ʽ), "R", mTestData(1, intRow, intCol)))
                    mTestData(1, intRow, intCol) = Format(mTestData(1, intRow, intCol), "##0.000#")
                End If
            Next
        Next
    End If
    mstr��ʽ = ""
    
    
'    bln���հ׶��� = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\frmLabMB", "���հ׶���", "0")
    bln���հ׶��� = zlDatabase.GetPara("frmLabMB_���հ׶���", 100, 1208, "0")
    blnС�����Զ��� = chk���Զ���.Value
    str���Զ��� = Trim(txt��С���Զ���)
    
'    gstrSql = "select ������ĿID,���Թ�ʽ,�����Թ�ʽ,CutOff��ʽ from ������Ŀ where ��Ŀ��� = 4"
'    Set rsTmp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption)
    
    '�ҳ�������Ŀ��ȥ���ظ�����Ŀ
    For intLoop = 1 To 8
        With Me.vsList
            If InStr(strItem & ";", ";" & .TextMatrix(intLoop, 13) & ";") <= 0 Then
                strItem = strItem & ";" & .TextMatrix(intLoop, 13)
            End If
        End With
    Next
    
    If OneCalc <> 0 Then strItem = ";" & OneCalc
    
    '��ʼ����
    aItem = Split(Mid(strItem, 2), ";")
    For lngLoop = 0 To UBound(aItem)
        strBC = "": strNC = "": strPC = "": strQC = ""
        dblBC = 0: dblNC = 0: dblPC = 0: dblQC = 0
        For intRow = 1 To 8
            For intCol = 1 To 12
                If Me.vsList.TextMatrix(intRow, 13) = aItem(lngLoop) Then
                    If InStr(mTestData(0, intRow, intCol), "BC") Then
                        strBC = strBC & ";" & mTestData(1, intRow, intCol)
                    ElseIf InStr(mTestData(0, intRow, intCol), "NC") Then
                        strNC = strNC & ";" & mTestData(1, intRow, intCol)
                    ElseIf InStr(mTestData(0, intRow, intCol), "PC") Then
                        strPC = strPC & ";" & mTestData(1, intRow, intCol)
                    ElseIf InStr(mTestData(0, intRow, intCol), "QC") Then
                        strQC = strQC & ";" & mTestData(1, intRow, intCol)
                    End If
                End If
            Next
        Next
        'BC
        If Trim(strBC) <> "" Then
            aBC = Split(Mid(strBC, 2), ";")
            For intLoop = 0 To UBound(aBC)
                dblBC = dblBC + Val(aBC(intLoop))
            Next
            If dblBC = 0 Then
                Me.txt�հ׶���.Text = 0
            Else
                Me.txt�հ׶���.Text = dblBC / Val((UBound(aBC)) + 1)
            End If
            Me.txt�հ׶���.Text = Format(Me.txt�հ׶���.Text, "##0.000#")
        Else
            Me.txt�հ׶���.Text = Format(0, "##0.000#")
        End If
        'NC
        If Trim(strNC) <> "" Then
            aNC = Split(Mid(strNC, 2), ";")
            For intLoop = 0 To UBound(aNC)
                dblNC = dblNC + Val(aNC(intLoop))
            Next
            If dblNC = 0 Then
                Me.txt���Զ���.Text = 0
            Else
                Me.txt���Զ���.Text = dblNC / Val((UBound(aNC)) + 1)
            End If
            Me.txt���Զ���.Text = Format(Me.txt���Զ���.Text - IIf(bln���հ׶���, Me.txt�հ׶���.Text, 0), "##0.000#")
            
        Else
            Me.txt���Զ���.Text = Format(0, "##0.000#")
        End If
        If blnС�����Զ��� = True And Val(Me.txt���Զ���.Text) <= Val(str���Զ���) And Val(str���Զ���) <> 0 Then
            Me.txt���Զ���.Text = Format(Val(str���Զ���), "##0.000#")
        End If
        'PC
        If Trim(strPC) <> "" Then
            aPC = Split(Mid(strPC, 2), ";")
            For intLoop = 0 To UBound(aPC)
                dblPC = dblPC + Val(aPC(intLoop))
            Next
            If dblPC = 0 Then
                Me.txt���Զ���.Text = 0
            Else
                Me.txt���Զ���.Text = dblPC / Val((UBound(aPC)) + 1)
            End If
            Me.txt���Զ���.Text = Format(Me.txt���Զ��� - IIf(bln���հ׶���, Me.txt�հ׶���.Text, 0), "##0.000#")
        Else
            Me.txt���Զ���.Text = Format(0, "##0.000#")
        End If
        'QC
        If Trim(strQC) <> "" Then
            aQC = Split(Mid(strQC, 2), ";")
            For intLoop = 0 To UBound(aQC)
                dblQC = dblQC + Val(aQC(intLoop))
            Next
            If dblQC = 0 Then
                dblQC = 0
            Else
                dblQC = dblQC / Val((UBound(aQC)) + 1)
            End If
            dblQC = Format(dblQC - IIf(bln���հ׶���, Me.txt�հ׶���.Text, 0), "##0.000#")
        Else
            dblQC = Format(0, "##0.000#")
        End If
        
        For intRow = 1 To 8
            If Me.vsList.TextMatrix(intRow, 13) = aItem(lngLoop) Then
                strCutOff = mTestItem(2, intRow)
            End If
        Next
        If strCutOff <> "" Then
            strCutOff = Replace(strCutOff, "BC", Me.txt�հ׶���.Text)
            strCutOff = Replace(strCutOff, "NC", Me.txt���Զ���.Text)
            strCutOff = Replace(strCutOff, "PC", Me.txt���Զ���.Text)
            strCutOff = Replace(strCutOff, "QC", dblQC)
            strCutOff = Calc.Eval(strCutOff)
            Me.txtCutOff.Text = Format(strCutOff, "##0.000#")
        End If
        
        '�������Ժ������Լ���
        For intRow = 1 To 8
            For intCol = 1 To 12
                If Me.vsList.TextMatrix(intRow, 13) = aItem(lngLoop) Then
                    If IsNumeric(mTestData(0, intRow, intCol)) = True Then
                        'ֻ������ͨ�걾
                        With Me.vsList
                            '����
'                            str���� = mTestItem(0, intRow)
                            mrsCalc.filter = "������ĿID=" & Val(Me.vsList.TextMatrix(intRow, 13))
                            If mrsCalc.EOF = False Then
                                str���� = mrsCalc("���Թ�ʽ") & ""
                            Else
                                str���� = ""
                            End If
                            If Trim(str����) <> "" Then
                                str���� = Replace(str����, "NC", Me.txt���Զ���.Text)
                                str���� = Replace(str����, "PC", Me.txt���Զ���.Text)
                                str���� = Replace(str����, "BC", Me.txt�հ׶���.Text)
                                str���� = Replace(str����, "QC", dblQC)
                                str���� = Replace(str����, "OD", Val(mTestData(1, intRow, intCol)) - IIf(bln���հ׶���, Me.txt�հ׶���.Text, 0))
                                If mTestData(1, intRow, intCol) <> "" Then
                                    mTestData(3, intRow, intCol) = IIf(Calc.Eval(str����), "����(+)", "����(-)")
                                End If
        '                        .TextMatrix(intRow, intCol) = IIf(Calc.Eval(str����), "����(+)", "����(-)")
                            End If
                            '������
                            mrsCalc.filter = "������ĿID=" & Val(Me.vsList.TextMatrix(intRow, 13))
                            If mrsCalc.EOF = False Then
                                str������ = mrsCalc("�����Թ�ʽ") & ""
                            Else
                                str������ = ""
                            End If
                                
                            If str������ <> "" And mTestData(3, intRow, intCol) <> "����(+)" Then
                                str������ = Replace(str������, "NC", Me.txt���Զ���.Text)
                                str������ = Replace(str������, "PC", Me.txt���Զ���.Text)
                                str������ = Replace(str������, "BC", Me.txt�հ׶���.Text)
                                str������ = Replace(str������, "QC", dblQC)
                                str������ = Replace(str������, "OD", Val(mTestData(1, intRow, intCol)) - IIf(bln���հ׶���, Me.txt�հ׶���.Text, 0))
                                If mTestData(1, intRow, intCol) <> "" Then
                                    mTestData(3, intRow, intCol) = IIf(Calc.Eval(str������), "������(��)", "����(-)")
                                    
                                End If
        '                        .TextMatrix(intRow, intCol) = IIf(Calc.Eval(str����), "������(+-)", "����(-)")
                            End If
                        End With
                        '�����ȥ�հ׵�OD
                        If Me.txt�հ׶���.Text <> "" And mTestData(1, intRow, intCol) <> "" Then
                            mTestData(2, intRow, intCol) = Format(mTestData(1, intRow, intCol) - IIf(bln���հ׶���, Me.txt�հ׶���.Text, 0), "##0.000#")
                        End If
                    Else
                        If mTestData(1, intRow, intCol) <> "" Then
'                            If InStr(mTestData(0, intRow, intCol), "BC") = 0 Then
                                mTestData(2, intRow, intCol) = Format(mTestData(1, intRow, intCol) - IIf(bln���հ׶���, Me.txt�հ׶���.Text, 0), "##0.000#")
                                mTestData(3, intRow, intCol) = Format(mTestData(1, intRow, intCol) - IIf(bln���հ׶���, Me.txt�հ׶���.Text, 0), "##0.000#")
'                            Else
'                                mTestData(2, intRow, intCol) = Format(mTestData(1, intRow, intCol), "##0.000#")
'                                mTestData(3, intRow, intCol) = Format(mTestData(1, intRow, intCol), "##0.000#")
'                            End If
                        End If
                    End If
                End If
            Next
        Next
    Next
    
    For intLoop = 0 To Me.opt��ʾ.Count - 1
        If Me.opt��ʾ(intLoop).Value = True Then
            Call opt��ʾ_Click(intLoop)
            Exit For
        End If
    Next
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub InitItem()
    '����   ��յ�ǰ��д�����ڵ�����
    Dim lngKey As Long
    Dim intLoop As Integer
    Dim intRow As Integer, intCol As Integer
    Dim rsTmp As New adodb.Recordset
    Dim aItem() As String
    Dim int���� As Integer
    
    Me.dtp����ʱ��.Value = Now
    On Error GoTo errH
    If Me.cbo��������.ListIndex = -1 Then Exit Sub
    
    gstrSql = "select ����,���Ƶ��,���ʱ��,���巽ʽ,�հ���ʽ from �������� where id = [1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.cbo��������.ItemData(Me.cbo��������.ListIndex))
    If rsTmp.EOF = True Then Exit Sub
    
    Me.txt�հ׶���.Text = ""
    Me.txt���Զ���.Text = ""
    Me.txt���Զ���.Text = ""
    Me.txtCutOff.Text = ""
    Me.txt���Թ�ʽ.Text = ""
    Me.txt�����Թ�ʽ.Text = ""
    Me.txtCutOff��ʽ.Text = ""
    Me.opt���嵥��.Value = True
    
    With Me.cbo����
        .Clear
        Me.cbo�ο�����.Clear
        Me.cbo�ο�����.AddItem ""
        Me.cbo�ο�����.ItemData(Me.cbo�ο�����.NewIndex) = 0
        aItem = Split(Nvl(rsTmp("����")), ";")
        For intLoop = 0 To UBound(aItem)
            .AddItem aItem(intLoop)
            Me.cbo�ο�����.AddItem aItem(intLoop)
        Next
        If .ListCount > 0 Then .ListIndex = 0
    End With
    
    With Me.cbo���Ƶ��
        .Clear
        aItem = Split(Nvl(rsTmp("���Ƶ��")), ";")
        For intLoop = 0 To UBound(aItem)
            .AddItem aItem(intLoop)
        Next
        If .ListCount > 0 Then .ListIndex = 0
    End With
    
    Me.txt���ʱ��.Text = Nvl(rsTmp("���ʱ��"))
    
    With Me.cbo���巽ʽ
        .Clear
        aItem = Split(Nvl(rsTmp("���巽ʽ")), ";")
        For intLoop = 0 To UBound(aItem)
            .AddItem aItem(intLoop)
        Next
        If .ListCount > 0 Then .ListIndex = 0
    End With
    
    With Me.cbo�հ���ʽ
        .Clear
        aItem = Split(Nvl(rsTmp("�հ���ʽ")), ";")
        For intLoop = 0 To UBound(aItem)
            .AddItem aItem(intLoop)
        Next
        If .ListCount > 0 Then .ListIndex = 0
    End With
    
    '�����ϴε�ʹ�ò���
    On Error Resume Next
    cbo����.Text = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "����", "")
    cbo�ο�����.Text = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "�ο�����", "")
    cbo���Ƶ��.Text = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "���Ƶ��", "")
    'txt���ʱ��.Text = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "���ʱ��", Me.txt���ʱ��.Text)
    If txt���ʱ��.Text = "" Then
        Me.txt���ʱ��.Text = Nvl(rsTmp("���ʱ��"))
    End If
    cbo�հ���ʽ.Text = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "�հ���ʽ", "")
    cbo���巽ʽ.Text = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "���巽ʽ", "")
    cbo�հ���ʽ.Text = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "�հ���ʽ", "")
    int���� = Val(GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName, "��������", 1))
    If int���� = 1 Then
        Me.opt����(0).Value = True
    Else
        Me.opt����(1).Value = True
    End If
    lngKey = GetSetting("ZLSOFT", "˽��ģ��\" & App.ProductName & "\" & Me.Name, "��ĿID", 0)
    
    
    
    cbo����.Text = zlDatabase.GetPara("frmLabMB_����", 100, 1208, "")
    cbo�ο�����.Text = zlDatabase.GetPara("frmLabMB_�ο�����", 100, 1208, "")
    cbo���Ƶ��.Text = zlDatabase.GetPara("frmLabMB_���Ƶ��", 100, 1208, "")
    txt���ʱ��.Text = zlDatabase.GetPara("frmLabMB_���ʱ��", 100, 1208, Me.txt���ʱ��.Text)
    If txt���ʱ��.Text = "" Then
        Me.txt���ʱ��.Text = Nvl(rsTmp("���ʱ��"))
    End If
    cbo���巽ʽ.Text = zlDatabase.GetPara("frmLabMB_���巽ʽ", 100, 1208, "")
    cbo�հ���ʽ.Text = zlDatabase.GetPara("frmLabMB_�հ���ʽ", 100, 1208, "")
    lngKey = zlDatabase.GetPara("frmLabMB_��ĿID", 100, 1208, 0)
    Me.chk���Զ���.Value = Mid(zlDatabase.GetPara("frmLabMB_���Զ���", 100, 1208, "0,"), 1, 1)
    Me.txt��С���Զ���.Text = Mid(zlDatabase.GetPara("frmLabMB_���Զ���", 100, 1208, "0,"), 3)
'    On Error GoTo 0
'    If lngKey <> 0 Then
'        For intLoop = 0 To Me.cbo������Ŀ.ListCount - 1
'            If Me.cbo������Ŀ.ItemData(intLoop) = lngKey Then
'                Me.cbo������Ŀ.ListIndex = intLoop
'                Call cbo������Ŀ_Click
'                Exit For
'            End If
'        Next
'    End If
    
    For intRow = 1 To 8
        For intCol = 1 To 13
            With Me.vsList
                .TextMatrix(intRow, intCol) = ""
            End With
        Next
    Next
    txt�Լ�����.Tag = ""
    txt�Լ�����.Text = ""
    txt�Լ�Ч��.Text = ""
    txt�Լ�����.Text = ""
    txt���Է���.Text = ""
    Erase mTestReagent
    Erase mTestItem
    For intLoop = 1 To 8
        If Me.vsList.TextMatrix(intLoop, 13) <> "" Then
            Me.vsList.Row = intLoop
            For intRow = 0 To Me.cbo������Ŀ.ListCount - 1
                If Me.cbo������Ŀ.ItemData(intRow) = Me.vsList.TextMatrix(intLoop, 13) Then
                    Call cbo������Ŀ_Click
                    
                End If
            Next
        End If
    Next
    Exit Sub
errH:
    If ErrCenter() = 0 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub zlRptPrint(ByVal bytMode As Byte)
    '����:�����ݸ��Ƶ��ɴ�ӡ�Ķ��󣬵��ô�ӡ
    '����:  bytMode��1-��ӡ;2-Ԥ��;3-�����EXCEL
'    If Me.rptList.Records.Count = 0 Then Exit Sub
'
'    '-------------------------------------------------
'    '�������ݱ��
'    If zlReportToVSFlexGrid(Me.vfgList, Me.rptList) = False Then Exit Sub
'
'    '-------------------------------------------------
'    '���ô�ӡ��������
'    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
'
'    Set objPrint.Body = Me.vfgList
'    objPrint.Title.Text = "������Ŀ�嵥"
'    Set objAppRow = New zlTabAppRow
'    Call objAppRow.Add("")
'    Call objAppRow.Add("��ӡʱ��:" & Now())
'    Call objPrint.BelowAppRows.Add(objAppRow)
'
'    If bytMode = 1 Then
'        bytMode = zlPrintAsk(objPrint)
'        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
'    Else
'        zlPrintOrView1Grd objPrint, bytMode
'    End If
    Dim intRow As Integer, intCol As Integer
    Dim strSQL1 As String, strSQL2 As String, strSQL3 As String
    
    On Error GoTo errH
    gcnOracle.BeginTrans
    For intRow = 1 To 8
        strSQL1 = "Zl_����ø����ӡ_Insert('OD','" & Chr(65 + intRow - 1)
        strSQL2 = "Zl_����ø����ӡ_Insert('����','" & Chr(65 + intRow - 1)
        strSQL3 = "Zl_����ø����ӡ_Insert('���','" & Chr(65 + intRow - 1)
        For intCol = 1 To 12
            strSQL1 = strSQL1 & "','" & mTestData(1, intRow, intCol)
            strSQL2 = strSQL2 & "','" & mTestData(3, intRow, intCol)
            strSQL3 = strSQL3 & "','" & mTestData(3, intRow, intCol)
        Next
        strSQL1 = strSQL1 & "')"
        strSQL2 = strSQL2 & "')"
        strSQL3 = strSQL3 & "')"
        zlDatabase.ExecuteProcedure strSQL1, Me.Caption
        zlDatabase.ExecuteProcedure strSQL2, Me.Caption
        zlDatabase.ExecuteProcedure strSQL3, Me.Caption
    Next
    
    Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1208_7", Me, "ø���ID=" & mlngKey, IIf(bytMode, 2, 1))
    
    gcnOracle.CommitTrans
    Exit Sub
errH:
    gcnOracle.RollbackTrans
End Sub
Private Sub SendData()
    '����               ����ø�����ݵ���ʦ����վ
    
    Dim rsTmp As New adodb.Recordset
    Dim intRow As Integer, intCol As Integer
    Dim strDate As String
    Dim lngMachine As Long
    Dim lngID As Long
    Dim lngDept As Long
    Dim strSampleType As String
    Dim strSex As String
    Dim strBirth As String
    Dim blnAuditing As Boolean
    Dim lngItemID As Long
    Dim str�ʿ�  As String
    Dim lngQCID As Long, i As Integer
    Dim strQCList() As String '������Ҫ���������
    Dim blnBegin As Boolean
    Dim astrSQL() As String
    
    
    On Error GoTo errH
        
    If Me.cbo��������.ListIndex = -1 Then Exit Sub
    lngMachine = Me.cbo��������.ItemData(Me.cbo��������.ListIndex)
    
    gstrSql = "select ʹ��С��ID from �������� where id = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngMachine)
    If rsTmp.EOF = True Then Call MsgBox("���ڼ���������ѡ��һ��ʹ��С��!", vbInformation): Exit Sub
    lngDept = rsTmp("ʹ��С��ID")
    
    strDate = zlDatabase.Currentdate
    '����Ϊ����
    ReDim strQCList(0) As String
    ReDim astrSQL(0)

    blnBegin = True
    
    For intRow = 1 To 8
        gstrSql = "select ����걾 from ���鱨����Ŀ a , ������ĿĿ¼ b where a.������Ŀid = b.id and a.������Ŀid = [1] and b.�����Ŀ = 0 "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(Val(Me.vsList.TextMatrix(intRow, 13))))
        If rsTmp.EOF = False Then strSampleType = Nvl(rsTmp("����걾"))
        '���㵱ǰ�е�CutOFFֵ,���ı�����ȡ
        Call CalcData(Val(Me.vsList.TextMatrix(intRow, 13)))
        For intCol = 1 To 12
            If (IsNumeric(mTestData(0, intRow, intCol)) = True Or UCase(Trim(mTestData(0, intRow, intCol))) Like "QC*") And IsNumeric(mTestData(1, intRow, intCol)) = True Then
                gstrSql = "Select a.*,Decode(c.�Ա�,Null,0,'��',1,'Ů',2) As �Ա�,to_char(c.��������,'yyyy-mm-dd') As �������� From ����걾��¼ a,����ҽ����¼ b,������Ϣ c " & _
                        " Where a.ҽ��id=b.id(+) And b.����id=c.����id(+)" & _
                        " And a.����ʱ�� Between [1] And [2]" & _
                        " And a.����ID=[3] And a.�걾���=[4] "
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, "��ѯ�걾��¼", CDate(Format(strDate, "yyyy-MM-dd") & " 00:00:00"), _
                        CDate(Format(strDate, "yyyy-MM-dd") & " 23:59:59"), lngMachine, mTestData(0, intRow, intCol))
                If rsTmp.EOF = True Then
                    strSex = 0: strBirth = ""
                    lngID = zlDatabase.GetNextId("����걾��¼")
                    str�ʿ� = "0"
                    If UCase(Trim(mTestData(0, intRow, intCol))) Like "QC*" Then
                        str�ʿ� = "1"
                    End If
                    ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                    astrSQL(UBound(astrSQL)) = "ZL_����걾��¼_INSERT(" & lngID & ",NULL,'" & _
                        mTestData(0, intRow, intCol) & "',NULL,NULL," & lngMachine & ",NULL," & _
                        "To_Date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'),NULL," & _
                        "To_Date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'),'" & strSampleType & "'," & _
                        "Null,To_Date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'),'" & UserInfo.���� & "','" & str�ʿ� & "'," & lngDept & ",0,0)"
                Else
                    strSex = Nvl(rsTmp("�Ա�"), 0)
                    strBirth = Nvl(rsTmp("��������"))
                    strSampleType = Nvl(rsTmp("�걾����"))
                    lngID = rsTmp("ID")
                    blnAuditing = Not IsNull(rsTmp("������"))
                    If blnAuditing = False Then
                        blnAuditing = Not IsNull(rsTmp("�����"))
                    End If
                End If
                
                'ֻ����û����˵ı걾
                If Not blnAuditing Then
'                    strItemRecords = Mid(strItemRecords, 2)
                        Dim strValue As String
                        If Val(Me.txtCutOff.Text) <> 0 Then
                            strValue = Format(Abs(Val(mTestData(2, intRow, intCol)) / Val(Me.txtCutOff.Text)), "##0.000#")
                        Else
                            strValue = "0"
                        End If
                        ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                        astrSQL(UBound(astrSQL)) = "ZL_������ͨ���_BATCHUPDATE(" & lngID & "," & _
                            lngMachine & ",'" & strSampleType & "'," & strSex & "," & _
                            IIf(strBirth = "", "Null", "To_Date('" & strBirth & "','yyyy-mm-dd hh24:mi:ss')") & ",'" & _
                            Me.vsList.TextMatrix(intRow, 13) & "^" & mTestData(3, intRow, intCol) & "^" & _
                            Format(Abs(Val(mTestData(2, intRow, intCol))), "##0.000#") & "^" & _
                            Format(Abs(Val(Me.txtCutOff.Text)), "##0.000#") & _
                            "^" & strValue & _
                           "',0," & mlngKey & ")"
                           ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                           astrSQL(UBound(astrSQL)) = "Zl_���¼�����_Cale(" & lngID & ")"
                End If
                
                
                If lngID > 0 And UCase(Trim(mTestData(0, intRow, intCol))) Like "QC*" Then
                    lngQCID = SendQC(lngID, Trim(mTestData(0, intRow, intCol)))
                    '�Զ�����
                    If lngQCID > 0 Then
                        If strQCList(UBound(strQCList)) <> "" Then ReDim Preserve strQCList(UBound(strQCList) + 1)
                        strQCList(UBound(strQCList)) = Format(CDate(strDate), "yyyy-MM-dd") & "," & CStr(lngQCID)
                    End If
                End If
            End If
        Next
    Next
'    gcnOracle.BeginTrans
'    blnBegin = True
    For i = LBound(astrSQL) To UBound(astrSQL)
        If astrSQL(i) <> "" Then
            zlDatabase.ExecuteProcedure astrSQL(i), "���͵���ʦվ"
        End If
    Next
'    gcnOracle.CommitTrans
    
'    gcnOracle.BeginTrans
'    blnBegin = True
    For i = LBound(strQCList) To UBound(strQCList)
        If InStr(strQCList(i), ",") > 0 Then
            Call AutoQCCompute(CDate(Split(strQCList(i), ",")(0)), Split(strQCList(i), ",")(1))
        End If
    Next
'    gcnOracle.CommitTrans
    
    Exit Sub
errH:
'    If blnBegin Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Sub


Private Function SendQC(ByVal lngID As Long, ByVal strSampleID As String) As Long
    '����Ϊ�ʿر걾
    
    Dim date��ǰ���� As Date, lngQCID As Long, str�걾�� As String
    Dim var�걾�� As Variant, iCoutn As Integer, lngDeviceID As Long
    Dim rsTmp As adodb.Recordset
    On Error GoTo errH
    lngQCID = 0
    date��ǰ���� = zlDatabase.Currentdate
    lngDeviceID = Me.cbo��������.ItemData(Me.cbo��������.ListIndex)
    
    gstrSql = "Select ID,�걾�� From �����ʿ�Ʒ Where [2] between ��ʼ���� and �������� And ����id = [1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, "ȡ�ʿ�Ʒ����", lngDeviceID, date��ǰ����)
    
    Do Until rsTmp.EOF Or lngQCID <> 0
        str�걾�� = "" & rsTmp.Fields("�걾��")
        If InStr(str�걾��, ",") > 0 Then
            var�걾�� = Split(str�걾��, ",")
            For iCoutn = 0 To UBound(var�걾��)
                If var�걾��(iCoutn) Like "*-*" Then
                    If strSampleID >= Val(Split(var�걾��(iCoutn), "-")(0)) And strSampleID <= Val(Split(var�걾��(iCoutn), "-")(1)) Then
                        lngQCID = rsTmp.Fields("ID")
                    End If
                Else
                    If var�걾��(iCoutn) = strSampleID Then
                        lngQCID = rsTmp.Fields("ID")
                    End If
                End If
            Next
        ElseIf str�걾�� Like "*-*" Then
            If strSampleID >= Val(Split(str�걾��, "-")(0)) And strSampleID <= Val(Split(str�걾��, "-")(1)) Then
                lngQCID = rsTmp.Fields("ID")
            End If
        Else
            If strSampleID = str�걾�� Then
                lngQCID = rsTmp.Fields("ID")
            End If
        End If
        
        rsTmp.MoveNext
    Loop
    
    If lngQCID > 0 Then
        gstrSql = "ZL_�����ʿؼ�¼_EDIT(1," & lngID & "," & lngQCID & ")"
        zlDatabase.ExecuteProcedure gstrSql, "����Ϊ�ʿ�Ʒ"
        
        SendQC = lngQCID
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub AutoQCCompute(ByVal date���� As Date, ByVal str�ʿ�Ʒ As String)

    '�Զ������ʿر걾
    ' date���� :�ʿؼ�������
    ' str�ʿ�Ʒ :�ʿ�Ʒ
    Dim rsTemp As adodb.Recordset, rsTmp As adodb.Recordset, strReturn As String
    Dim lngDeviceID As Long
    lngDeviceID = mlngMachine
    On Error GoTo errH
    gstrSql = "Select Distinct B.��Ŀid, C.����, C.������, C.Ӣ����" & vbNewLine & _
              " From �����ʿ�Ʒ A, �����ʿ�Ʒ��Ŀ B, ����������Ŀ C" & vbNewLine & _
              " Where A.ID = B.�ʿ�Ʒid And B.��Ŀid = C.ID And A.����id = [1] "
        
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, "LisComm�Զ�����", lngDeviceID)
    Do Until rsTmp.EOF
        '����һ��ʱ��
            gstrSql = "Select Zl_�����ʿؼ�¼_Compute(" & lngDeviceID & ", " & rsTmp("��ĿID") & ", To_Date('" & Format(date����, "yyyy-mm-dd") & "','yyyy-mm-dd'), '" & str�ʿ�Ʒ & "') From Dual"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "LisComm�Զ�����")

            If rsTemp.RecordCount <= 0 Then strReturn = strReturn & Format(date����, "yyyy-mm-dd") & " " & Nvl(rsTmp("������")) & "(" & Nvl(rsTmp("Ӣ����")) & ")  ������̵��ô���" & vbCrLf
            If InStr(rsTemp.Fields(0).Value, "����ʧ�أ�") > 0 Then
                strReturn = strReturn & Format(date����, "yyyy-mm-dd") & " " & Nvl(rsTmp("������")) & "(" & Nvl(rsTmp("Ӣ����")) & ")" & rsTemp.Fields(0).Value & vbCrLf

            ElseIf InStr(rsTemp.Fields(0).Value, "������ɣ�") <= 0 Then
                If InStr(rsTemp.Fields(0).Value, "������δ���־����ʧ�أ�") <= 0 Then
                strReturn = strReturn & Format(date����, "yyyy-mm-dd") & " " & Nvl(rsTmp("������")) & "(" & Nvl(rsTmp("Ӣ����")) & ")" & rsTemp.Fields(0).Value & vbCrLf
                End If
            End If
        rsTmp.MoveNext
    Loop
    If Trim(strReturn) <> "" Then
        MsgBox "�����ѱ��棬�ʿؼ���ʱ����ʧ�ػ򾯸棡", vbInformation, "��������"
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub ShowMe(objfrm As Object, lngMachine As Long)
    '�򿪴���
    mlngMachine = lngMachine
    Me.Show , objfrm
End Sub
Private Sub InitRecordSet(rsNumber As adodb.Recordset)
    '��ʼ����¼��(���ڼ�¼������Ŀ)
    Dim rsTmp As New adodb.Recordset
    
    Set rsNumber = New adodb.Recordset
    rsNumber.Fields.Append "������ĿID", adBigInt
    rsNumber.Fields.Append "���Թ�ʽ", adVarChar, 50
    rsNumber.Fields.Append "�����Թ�ʽ", adVarChar, 50
    rsNumber.Fields.Append "CutOff��ʽ", adVarChar, 50
    rsNumber.CursorLocation = adUseClient
    rsNumber.LockType = adLockOptimistic
    rsNumber.CursorType = adOpenStatic
    rsNumber.Open
    
    
    gstrSql = "select ������ĿID,���Թ�ʽ,�����Թ�ʽ,CutOff��ʽ from ������Ŀ where ��Ŀ��� = 4"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    
    Do Until rsTmp.EOF
        rsNumber.AddNew
        rsNumber("������ĿID") = rsTmp("������ĿID")
        rsNumber("���Թ�ʽ") = rsTmp("���Թ�ʽ") & ""
        rsNumber("�����Թ�ʽ") = rsTmp("�����Թ�ʽ") & ""
        rsNumber("CutOff��ʽ") = rsTmp("CutOff��ʽ") & ""
        rsNumber.Update
        rsTmp.MoveNext
    Loop
    
End Sub

Private Sub SelectBatch()
    '�Լ�����ѡ����
    Dim strReturn As String
    Dim lngItemID As Long
    Dim intLoop As Integer
    
    If Me.cbo������Ŀ.ListIndex > -1 Then
        lngItemID = Me.cbo������Ŀ.ItemData(Me.cbo������Ŀ.ListIndex)
    End If
    
    strReturn = GetBatchNo(Me.txt�Լ�����, Me.txt�Լ�����.Text, lngItemID)
    If UBound(Split(strReturn, "|")) = 4 Then
        lngItemID = Val(Split(strReturn, "|")(0))
        txt�Լ�����.Tag = Split(strReturn, "|")(1) '��ñ����������
        txt�Լ����� = Split(strReturn, "|")(1)
        txt�Լ�Ч�� = Split(strReturn, "|")(2)
        txt�Լ����� = Split(strReturn, "|")(3)
        txt���Է��� = Split(strReturn, "|")(4)
        
        If lngItemID = 0 Then
            If Me.cbo������Ŀ.ListIndex > -1 Then
                lngItemID = Me.cbo������Ŀ.ItemData(Me.cbo������Ŀ.ListIndex)
            End If
        Else
            For intLoop = 0 To Me.cbo������Ŀ.ListCount - 1
                If Val(Me.cbo������Ŀ.ItemData(intLoop)) = lngItemID Then
                    Me.cbo������Ŀ.ListIndex = intLoop
                    Exit For
                End If
            Next
        End If
        For intLoop = 1 To 8
            If Val(Me.vsList.TextMatrix(intLoop, 13)) = lngItemID Then
                mTestReagent(intLoop) = txt�Լ�����.Tag & ";" & txt�Լ�Ч�� & ";" & txt�Լ����� & ";" & txt���Է���
            End If
        Next
    Else
        txt�Լ�����.Text = ""
        txt�Լ�����.Tag = ""
        txt�Լ�Ч��.Text = ""
        txt�Լ�����.Text = ""
        txt���Է���.Text = ""
        If Me.cbo������Ŀ.ListIndex > -1 Then
            lngItemID = Val(Me.cbo������Ŀ.ItemData(Me.cbo������Ŀ.ListIndex))
        End If
        For intLoop = 1 To 8
            If Val(Me.vsList.TextMatrix(intLoop, 13)) = 0 Then
                mTestReagent(intLoop) = ""
            End If
        Next
    
    End If
End Sub
Private Function GetBatchNo(ByRef objTxt As TextBox, ByVal strInput As String, ByVal lngItemID As Long) As String
    '�Լ�����ѡ����
    Dim rsTmp As adodb.Recordset, strsql As String
    Dim objPoint As POINTAPI
    Dim sglX As Single, sglY As Single
    Dim strKey As String '���ҹؼ���
    On Error GoTo hErr
    
    strKey = DelInvalidChar(strInput) & "%"
    If lngItemID = 0 Then
    strsql = "Select Rownum As ID, a.* From (Select a.�Լ�����, a.�Լ�Ч��, a.�Լ�����, a.���Է���, b.���� As ������Ŀ, c.������Ŀid" & vbNewLine & _
            "From ����ø���Լ� A, ������ĿĿ¼ B, ���鱨����Ŀ C" & vbNewLine & _
            "Where a.������Ŀid = b.Id(+) And a.������Ŀid = c.������Ŀid(+) And b.�����Ŀ(+) = 0 And a.�Լ�Ч�� > Sysdate " & vbNewLine & _
            " And (A.�Լ����� Like [1] Or A.�Լ����� Like [2] Or A.���Է��� Like [2] ) " & vbNewLine & _
            ") A"
    Else
    strsql = "Select Rownum As ID, a.* From (Select a.�Լ�����, a.�Լ�Ч��, a.�Լ�����, a.���Է���, b.���� As ������Ŀ, c.������Ŀid" & vbNewLine & _
            "From ����ø���Լ� A, ������ĿĿ¼ B, ���鱨����Ŀ C" & vbNewLine & _
            "Where a.������Ŀid = b.Id And a.������Ŀid = c.������Ŀid And b.�����Ŀ = 0 And a.�Լ�Ч�� > Sysdate " & vbNewLine & _
            " And  C.������ĿID = [3]  And (A. �Լ����� Like [1] Or A.�Լ����� Like [2] Or A.���Է��� Like [2] ) " & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select a.�Լ�����, a.�Լ�Ч��, a.�Լ�����, a.���Է���, Null As ������Ŀ, Null As ������Ŀid" & vbNewLine & _
            "From ����ø���Լ� A" & vbNewLine & _
            "Where a.�Լ�Ч�� > Sysdate And ������Ŀid Is Null" & vbNewLine & _
            "  And (A. �Լ����� Like [1] Or A.�Լ����� Like [2] Or A.���Է��� Like [2] ) " & vbNewLine & _
            "Order By �Լ�Ч�� Desc) A"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strsql, "ȡ�Լ�����", strKey, "%" & strKey, lngItemID)
    If rsTmp.EOF Then
        GetBatchNo = strInput
    Else
        If rsTmp.RecordCount = 1 Then
            GetBatchNo = "" & rsTmp!������ĿID & "|" & rsTmp!�Լ����� & "|" & rsTmp!�Լ�Ч�� & "|" & rsTmp!�Լ����� & "|" & rsTmp!���Է���
        Else
            Call ClientToScreen(objTxt.hWnd, objPoint)
            sglX = objPoint.x * 15 - 30
            sglY = objPoint.y * 15 + objTxt.Height
            If frmSelectList.ShowSelect(Me, rsTmp, "�Լ�����,800,0,0;�Լ�Ч��,800,0,0;�Լ�����,1500,0,0;���Է���,2500,0,0;������Ŀ,5500,0,0", sglX, sglY, objTxt.Width, 2000, Me.Name & "\ø���Լ�����ѡ��", "��ѡ���Լ�����") Then
                GetBatchNo = "" & rsTmp!������ĿID & "|" & rsTmp!�Լ����� & "|" & rsTmp!�Լ�Ч�� & "|" & rsTmp!�Լ����� & "|" & rsTmp!���Է���
            Else
                GetBatchNo = strInput
            End If
        End If
    End If
    Exit Function
hErr:
    MsgBox Err.Description
End Function

