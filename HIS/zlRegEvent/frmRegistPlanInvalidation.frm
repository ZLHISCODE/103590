VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmRegistPlanInvalidation 
   Caption         =   "ͣ��ʱ������"
   ClientHeight    =   8475
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11865
   Icon            =   "frmRegistPlanInvalidation.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8475
   ScaleWidth      =   11865
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picCmd 
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   570
      ScaleHeight     =   600
      ScaleWidth      =   9225
      TabIndex        =   52
      Top             =   7890
      Width           =   9225
      Begin VB.Frame fraSplit 
         Height          =   60
         Index           =   1
         Left            =   -30
         TabIndex        =   56
         Top             =   -15
         Width           =   11805
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   6780
         TabIndex        =   55
         Top             =   75
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   8040
         TabIndex        =   54
         Top             =   60
         Width           =   1100
      End
      Begin VB.CheckBox chkClearHistory 
         Caption         =   "���������ʧЧ��ͣ��ʱ��(&S)"
         Height          =   285
         Left            =   2190
         TabIndex        =   53
         Top             =   75
         Width           =   2835
      End
   End
   Begin VB.PictureBox picStop 
      BorderStyle     =   0  'None
      Height          =   3390
      Left            =   5160
      ScaleHeight     =   3390
      ScaleWidth      =   6645
      TabIndex        =   40
      Top             =   195
      Width           =   6645
      Begin VB.CommandButton cmdSel 
         Caption         =   "&P"
         Height          =   285
         Left            =   4560
         TabIndex        =   58
         Top             =   435
         Width           =   315
      End
      Begin VB.CommandButton cmdRestore 
         Caption         =   "�ָ�ͣ�ð���(&Y)"
         Height          =   345
         Left            =   4905
         TabIndex        =   57
         Top             =   30
         Width           =   1710
      End
      Begin VB.CommandButton cmdDeleteTime 
         Caption         =   "ɾ��(&R)"
         Height          =   345
         Left            =   5775
         TabIndex        =   51
         Top             =   405
         Width           =   855
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "����(&A)"
         Height          =   345
         Left            =   4905
         TabIndex        =   48
         Top             =   405
         Width           =   855
      End
      Begin VB.TextBox txtMemo 
         Height          =   315
         Left            =   840
         MaxLength       =   100
         TabIndex        =   41
         Top             =   420
         Width           =   4050
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   285
         Left            =   840
         TabIndex        =   42
         Top             =   75
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   503
         _Version        =   393216
         CalendarTitleBackColor=   8421504
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   109445123
         CurrentDate     =   40427.6041666667
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   285
         Left            =   3045
         TabIndex        =   43
         Top             =   75
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   503
         _Version        =   393216
         CalendarTitleBackColor=   8421504
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   109445123
         CurrentDate     =   40427.0416666667
      End
      Begin VSFlex8Ctl.VSFlexGrid vsList 
         Height          =   2415
         Left            =   0
         TabIndex        =   44
         Top             =   795
         Width           =   6345
         _cx             =   11192
         _cy             =   4260
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
         GridColor       =   -2147483633
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
         Rows            =   10
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmRegistPlanInvalidation.frx":030A
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
         ExplorerBar     =   7
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
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "ͣ��ʱ��"
         Height          =   180
         Index           =   0
         Left            =   45
         TabIndex        =   47
         Top             =   120
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   2730
         TabIndex        =   46
         Top             =   105
         Width           =   225
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "��ע"
         Height          =   180
         Index           =   3
         Left            =   390
         TabIndex        =   45
         Top             =   480
         Width           =   360
      End
   End
   Begin VB.PictureBox picOthers 
      BorderStyle     =   0  'None
      Height          =   3330
      Left            =   5115
      ScaleHeight     =   3330
      ScaleWidth      =   7815
      TabIndex        =   34
      Top             =   4305
      Width           =   7815
      Begin VB.CommandButton cmdClear 
         Caption         =   "ȫ��(&U)"
         Height          =   345
         Left            =   5325
         TabIndex        =   50
         Top             =   30
         Width           =   855
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "ɾ��(&D)"
         Height          =   345
         Left            =   4425
         TabIndex        =   49
         Top             =   30
         Width           =   855
      End
      Begin VB.TextBox txtCode 
         Height          =   315
         Left            =   525
         TabIndex        =   36
         Top             =   45
         Width           =   2580
      End
      Begin VB.CommandButton cmdOthers 
         Caption         =   "��������(&O)"
         Height          =   345
         Left            =   3165
         TabIndex        =   35
         Top             =   30
         Width           =   1230
      End
      Begin VSFlex8Ctl.VSFlexGrid vsOthers 
         Height          =   2430
         Left            =   45
         TabIndex        =   38
         Top             =   465
         Width           =   6405
         _cx             =   11298
         _cy             =   4286
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
         GridColor       =   -2147483633
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
         GridLines       =   2
         GridLinesFixed  =   9
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   23
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmRegistPlanInvalidation.frx":03F9
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
         ExplorerBar     =   7
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
         Begin VB.PictureBox picImgList 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   45
            ScaleHeight     =   225
            ScaleWidth      =   210
            TabIndex        =   39
            Top             =   60
            Width           =   210
            Begin VB.Image imgColList 
               Height          =   195
               Left            =   0
               Picture         =   "frmRegistPlanInvalidation.frx":06AE
               ToolTipText     =   "ѡ����Ҫ��ʾ����(ALT+C)"
               Top             =   0
               Width           =   195
            End
         End
      End
      Begin VB.Label lblCon 
         AutoSize        =   -1  'True
         Caption         =   "�ű�"
         Height          =   180
         Index           =   4
         Left            =   90
         TabIndex        =   37
         Top             =   105
         Width           =   360
      End
   End
   Begin VB.PictureBox picBill 
      BorderStyle     =   0  'None
      Height          =   7815
      Left            =   0
      ScaleHeight     =   7815
      ScaleWidth      =   4995
      TabIndex        =   0
      Top             =   0
      Width           =   4995
      Begin VB.Frame Frame1 
         Caption         =   "������Ϣ"
         Height          =   1860
         Left            =   60
         TabIndex        =   12
         Top             =   105
         Width           =   4890
         Begin VB.TextBox txt�ű� 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   660
            MaxLength       =   5
            TabIndex        =   21
            Top             =   270
            Width           =   960
         End
         Begin VB.TextBox txt�޺� 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   3540
            MaxLength       =   5
            TabIndex        =   20
            Top             =   660
            Width           =   1215
         End
         Begin VB.ComboBox cboItem 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   660
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   1035
            Width           =   2115
         End
         Begin VB.ComboBox cboDoctor 
            Height          =   300
            Left            =   660
            TabIndex        =   18
            Top             =   1410
            Width           =   2115
         End
         Begin VB.ComboBox cbo���� 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   660
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   660
            Width           =   2115
         End
         Begin VB.TextBox txt��Լ 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   3540
            MaxLength       =   5
            TabIndex        =   16
            Top             =   1035
            Width           =   1215
         End
         Begin VB.CheckBox chk���� 
            Caption         =   "�Һ�ʱ���뽨����"
            Height          =   195
            Left            =   2985
            TabIndex        =   15
            Top             =   1463
            Width           =   1755
         End
         Begin VB.ComboBox cbo���� 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   3540
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   270
            Width           =   1230
         End
         Begin VB.CheckBox chk��ſ��� 
            Caption         =   "��ſ���"
            Height          =   255
            Left            =   1750
            TabIndex        =   13
            Top             =   293
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "�ű�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   210
            TabIndex        =   28
            Top             =   330
            Width           =   390
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "����"
            Height          =   180
            Left            =   240
            TabIndex        =   27
            Top             =   720
            Width           =   360
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "��Ŀ"
            Height          =   180
            Left            =   240
            TabIndex        =   26
            Top             =   1110
            Width           =   360
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "ҽ��"
            Height          =   180
            Left            =   240
            TabIndex        =   25
            Top             =   1485
            Width           =   360
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "�޺�"
            Height          =   180
            Left            =   3105
            TabIndex        =   24
            Top             =   720
            Width           =   360
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "��Լ"
            Height          =   180
            Left            =   3105
            TabIndex        =   23
            Top             =   1095
            Width           =   360
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "����"
            Height          =   180
            Left            =   3105
            TabIndex        =   22
            Top             =   330
            Width           =   360
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Ӧ������:"
         Height          =   2730
         Left            =   75
         TabIndex        =   6
         Top             =   5070
         Width           =   4860
         Begin VB.OptionButton opt���� 
            Caption         =   "������"
            Height          =   180
            Index           =   0
            Left            =   150
            TabIndex        =   10
            Top             =   300
            Value           =   -1  'True
            Width           =   900
         End
         Begin VB.OptionButton opt���� 
            Caption         =   "ָ������"
            Height          =   180
            Index           =   1
            Left            =   1020
            TabIndex        =   9
            Top             =   300
            Width           =   1020
         End
         Begin VB.OptionButton opt���� 
            Caption         =   "��̬����"
            Height          =   180
            Index           =   2
            Left            =   2115
            TabIndex        =   8
            Top             =   300
            Width           =   1020
         End
         Begin VB.OptionButton opt���� 
            Caption         =   "ƽ������"
            Height          =   180
            Index           =   3
            Left            =   3135
            TabIndex        =   7
            Top             =   315
            Width           =   1020
         End
         Begin MSComctlLib.ListView lvwDept 
            Height          =   2040
            Left            =   105
            TabIndex        =   11
            Top             =   615
            Width           =   4650
            _ExtentX        =   8202
            _ExtentY        =   3598
            View            =   2
            Arrange         =   2
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Ӧ��ʱ��"
         Height          =   2835
         Left            =   60
         TabIndex        =   1
         Top             =   2070
         Width           =   4890
         Begin VSFlex8Ctl.VSFlexGrid vsPlan1 
            Height          =   660
            Left            =   1200
            TabIndex        =   33
            Top             =   1305
            Width           =   3510
            _cx             =   6191
            _cy             =   1164
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
            BackColorBkg    =   -2147483634
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   2
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   300
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmRegistPlanInvalidation.frx":0BFC
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
         Begin VB.CheckBox chk��Ч�� 
            Caption         =   "��Ч��"
            Height          =   195
            Left            =   285
            TabIndex        =   29
            Top             =   2085
            Width           =   855
         End
         Begin VB.OptionButton opt�� 
            Caption         =   "ÿ��(&D)"
            Height          =   315
            Left            =   225
            TabIndex        =   5
            Top             =   285
            Width           =   960
         End
         Begin VB.OptionButton opt�� 
            Caption         =   "ÿ��(&W)"
            Height          =   315
            Left            =   225
            TabIndex        =   4
            Top             =   630
            Width           =   930
         End
         Begin VB.ComboBox cbo�� 
            Height          =   300
            Left            =   1170
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   270
            Width           =   1110
         End
         Begin VSFlex8Ctl.VSFlexGrid vsPlan 
            Height          =   660
            Left            =   1200
            TabIndex        =   2
            Top             =   690
            Width           =   3510
            _cx             =   6191
            _cy             =   1164
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
            BackColorBkg    =   -2147483636
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   7
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   300
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmRegistPlanInvalidation.frx":0C43
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
         Begin MSComCtl2.DTPicker dtpBegin 
            Height          =   300
            Left            =   1200
            TabIndex        =   30
            Top             =   2040
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   109445123
            CurrentDate     =   38091
         End
         Begin MSComCtl2.DTPicker dtpEnd 
            Height          =   300
            Left            =   1215
            TabIndex        =   31
            Top             =   2415
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   109445123
            CurrentDate     =   38091
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            Height          =   180
            Left            =   930
            TabIndex        =   32
            Top             =   2475
            Width           =   180
         End
      End
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmRegistPlanInvalidation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng����ID As Long, mblnSucces As Boolean, mblnFirst As Boolean
Private mlngModule As Long, mstrPrivs As String
Private mrsRoom As ADODB.Recordset
Private mstrDelete��� As String   'ɾ�����

Public Function ShowCard(ByVal mfrmMain As Form, ByVal lngModule As Long, ByVal strPrivs As String, _
     Optional lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ��Ҫ�޸ĵļƻ�����
    '���:mfrmMain-���õ�������
    '     lngModule-ģ���
    '     strPrivs-Ȩ�޴�
    '     lng����ID-�ҺŰ���ID.
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-09-07 10:05:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
     mlngModule = lngModule: mstrPrivs = strPrivs:  mblnSucces = False: mlng����ID = lng����ID
    Me.Show 1, mfrmMain
    ShowCard = mblnSucces
End Function
Private Function InitPanel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����������
    '����:���˺�
    '����:2010-09-08 11:41:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPane As Pane, strKey As String
    Dim objBill As Pane, lngHeight As Long, lngBillWidth As Long
    lngHeight = picCmd.Height \ Screen.TwipsPerPixelY
    lngBillWidth = picBill.Width \ Screen.TwipsPerPixelX
    With dkpMan
        Set objPane = .CreatePane(3, 400, 400, DockBottomOf, Nothing)
        objPane.Title = "��ť��"
        objPane.Options = PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable Or PaneNoCaption Or PaneNoHideable
        objPane.Handle = picCmd.Hwnd
        objPane.MaxTrackSize.Height = lngHeight
        objPane.MinTrackSize.Height = lngHeight
        
        Set objBill = .CreatePane(1, 300, 100, DockTopOf, objPane)
        objBill.Title = "��ǰ�ҺŰ���": objBill.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        objBill.Handle = picBill.Hwnd
        objBill.MaxTrackSize.Width = lngBillWidth: objBill.MinTrackSize.Width = lngBillWidth:
         Set objPane = .CreatePane(2, 400, 400, DockRightOf, objBill)
        objPane.Title = "ͣ��ʱ������"
        objPane.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        objPane.Handle = picStop.Hwnd
         Set objPane = .CreatePane(3, 400, 400, DockBottomOf, objPane)
        objPane.Title = "Ӧ���������Һ���Ŀ"
        objPane.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        objPane.Handle = picOthers.Hwnd
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = False 'ʵʱ�϶�
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
    End With
    dkpMan.RecalcLayout: DoEvents
    'zlRestoreDockPanceToReg Me, dkpMan, "����"
End Function



Private Function LoadData(Optional blnRestore As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ؼƻ�����������Ϣ
    '����:���˺�
    '����:2009-09-14 14:40:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, strSQL As String, i As Long
    Dim strCurDate As String

    Err = 0: On Error GoTo Errhand:
    strCurDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    mstrDelete��� = ""
    dtpStartDate.Value = Format(CDate(strCurDate) + 1, "yyyy-mm-dd 00:00:00")
    dtpEndDate.Value = CDate(Format(dtpStartDate.Value, "yyyy-mm-dd") & " 23:59:59")
    dtpStartDate.MinDate = CDate(strCurDate)
    dtpEndDate.MinDate = dtpStartDate.MinDate
    
   strSQL = " " & _
    "   Select A.Id as ����ID,0 as �ƻ�ID,A.����,  A.����,  A.����id,  A.��Ŀid, A.ҽ������,  A.ҽ��id,  F.�޺���,  F.��Լ��,   " & _
    "           A.����,  A.��һ,  A.�ܶ�,  A.����,  A.����,  A.����,  A.����, " & _
    "           A.��������,  A.���﷽ʽ,  A.��ſ���,  A.��ʼʱ��,  A.��ֹʱ��,B.���� As ��Ŀ,D.���� As ���� " & _
    "   From �ҺŰ��� A,�շ���ĿĿ¼ B,�ҺŰ��żƻ� C,���ű� D,�ҺŰ������� F " & _
    "   Where A.Id=C.����ID(+) And A.��Ŀid=b.Id(+) And A.����id =d.Id(+) " & _
    "         And A.Id=[1]  And a.Id = f.����id(+) And" & _
    "  Decode(To_Char([2], 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����', Null) =f.������Ŀ(+)"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, CDate(strCurDate))
    If rsTemp.EOF Then
        MsgBox "ע��:" & vbCrLf & _
        "    �ҺŰ��ſ����Ѿ�������ɾ��,�����ٽ���ͣ��ʱ������!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    '�������ݵ��ؼ���
    txt�ű�.Text = Nvl(rsTemp!����)
    cbo����.AddItem Nvl(rsTemp!����): cbo����.ListIndex = cbo����.NewIndex
    txt�޺�.Text = Nvl(rsTemp!�޺���): txt��Լ.Text = Nvl(rsTemp!��Լ��)
    chk��ſ���.Value = IIf(Val(Nvl(rsTemp!��ſ���)) = 1, 1, 0)
    chk����.Value = IIf(Val(Nvl(rsTemp!��������)) = 1, 1, 0)
    With cbo����
        .AddItem Nvl(rsTemp!����): .ItemData(.NewIndex) = Val(Nvl(rsTemp!����ID)): .ListIndex = .NewIndex
    End With
    With cboItem
        .AddItem Nvl(rsTemp!��Ŀ): .ItemData(.NewIndex) = Val(Nvl(rsTemp!��ĿID)): .ListIndex = .NewIndex
    End With
    With cboDoctor
        .AddItem Nvl(rsTemp!ҽ������): .ItemData(.NewIndex) = Val(Nvl(rsTemp!ҽ��ID)): .ListIndex = .NewIndex
    End With
    If Nvl(rsTemp!����) <> Nvl(rsTemp!��һ) Or Nvl(rsTemp!����) <> Nvl(rsTemp!�ܶ�) _
        Or Nvl(rsTemp!����) <> Nvl(rsTemp!����) Or Nvl(rsTemp!����) <> Nvl(rsTemp!����) _
        Or Nvl(rsTemp!����) <> Nvl(rsTemp!����) Or Nvl(rsTemp!����) <> Nvl(rsTemp!����) Then
        'ÿ��
        opt��.Value = True
        With vsPlan
            For i = 0 To 4
                .TextMatrix(1, i) = Nvl(rsTemp.Fields("��" & Replace(.ColKey(i), "����", "��")))  '��֪ʲôԭ��,��.colkey(i)����,Ҫ���ĳ�������.
            Next
        End With
        With vsPlan1
            For i = 0 To 1
                .TextMatrix(1, i) = Nvl(rsTemp.Fields("��" & Replace(.ColKey(i), "����", "��")))  '��֪ʲôԭ��,��.colkey(i)����,Ҫ���ĳ�������.
            Next
        End With
    Else
        'ÿ��
        opt��.Value = True:  cbo��.ListIndex = cbo.FindIndex(cbo��, Nvl(rsTemp!����), True): cbo��.Enabled = True
    End If
    '��Чʱ�䷶Χ
    dtpBegin.Value = Format(zlDatabase.Currentdate, "yyyy-MM-dd 00:00:00")
    dtpEnd.Value = CDate("3000-01-01")
    If Not IsNull(rsTemp!��ʼʱ��) Then
        chk��Ч��.Value = 1
        dtpBegin.Value = CDate(Format(rsTemp!��ʼʱ��, "yyyy-mm-dd HH:MM:SS"))
        If Not IsNull(rsTemp!��ֹʱ��) Then
            dtpEnd.Value = CDate(Format(rsTemp!��ֹʱ��, "yyyy-mm-dd HH:MM:SS"))
        End If
    End If
        
    Select Case Val(Nvl(rsTemp!���﷽ʽ))     '0-�����1-ָ�����ҡ�2-��̬���3-ƽ������,��Ӧ������������
        Case 0  '"������"
            opt����(0).Value = True
        Case 1  ' "ָ������"
            opt����(1).Value = True
        Case 2 '"��̬����"
            opt����(2).Value = True
        Case 3 ' "ƽ������"
            opt����(3).Value = True
    End Select
    
    strSQL = "Select �ű�ID,�������ҡ�From �ҺŰ������� Where �ű�ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
    Dim objItem As ListItem
    
    lvwDept.ListItems.Clear: i = 1
    Do While Not rsTemp.EOF
       Set objItem = lvwDept.ListItems.Add(, "K" & i, Nvl(rsTemp!��������))
        objItem.Checked = True
        i = i + 1
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    '���ظùҺ���Ŀ�ĵ�ͣ��ʱ����Ϣ
    strSQL = "Select ����ID,���,��ʼֹͣʱ��,����ֹͣʱ��,�ƶ���,�ƶ�����,��ע From �ҺŰ���ͣ��״̬ where ����ID=[1] Order by ��ʼֹͣʱ��,�ƶ�����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
    With vsList
        .Clear 1
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        i = 1
        Do While Not rsTemp.EOF
            .TextMatrix(i, .ColIndex("���")) = i
            .Cell(flexcpData, i, .ColIndex("���")) = Val(Nvl(rsTemp!���))
            .TextMatrix(i, .ColIndex("��ʼͣ��ʱ��")) = Format(rsTemp!��ʼֹͣʱ��, "yyyy-mm-dd HH:MM")
            .TextMatrix(i, .ColIndex("����ͣ��ʱ��")) = Format(rsTemp!����ֹͣʱ��, "yyyy-mm-dd HH:MM")
            .TextMatrix(i, .ColIndex("�ƶ���")) = Nvl(rsTemp!�ƶ���)
            .TextMatrix(i, .ColIndex("�ƶ�����")) = Format(rsTemp!�ƶ�����, "yyyy-mm-dd HH:MM:SS")
            .TextMatrix(i, .ColIndex("��ע")) = Nvl(rsTemp!��ע)
            If Format(rsTemp!����ֹͣʱ��, "yyyy-mm-dd HH:MM:SS") < strCurDate Then
                .Cell(flexcpForeColor, i, 1, i, .Cols - 1) = Me.BackColor
            Else
                .Cell(flexcpForeColor, i, 1, i, .Cols - 1) = &H8000000C
            End If
            .RowData(i) = 1
            i = i + 1
            rsTemp.MoveNext
        Loop
       zl_vsGrid_Para_Restore mlngModule, vsList, Me.Caption, "ͣ�ð���-ͣ�üƻ�", True, InStr(1, mstrPrivs, ";��������;") > 0
    End With
    If blnRestore = False Then
        vsOthers.Clear 1
        vsOthers.Rows = 2
       zl_vsGrid_Para_Restore mlngModule, vsOthers, Me.Caption, "ͣ�ð���-�ҺŰ���", True, InStr(1, mstrPrivs, ";��������;") > 0
    End If
    
    '����:43148
    gstrSQL = " Select  ����   From ����ͣ��ԭ�� where ȱʡ��־=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If Not rsTemp.EOF Then txtMemo.Text = Nvl(rsTemp!����)
    rsTemp.Close
    Set rsTemp = Nothing
    LoadData = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub cmdAdd_Click()
    Dim lngRow As Long
    If CheckStopValied = False Then Exit Sub
    With vsList
        If .TextMatrix(.Row, .ColIndex("��ʼͣ��ʱ��")) <> "" Then
            .Rows = .Rows + 1
            .Row = .Rows - 1: lngRow = .Row
        Else
            lngRow = .Row
        End If
        .RowData(lngRow) = 0
        .TextMatrix(lngRow, .ColIndex("���")) = lngRow
        .TextMatrix(lngRow, .ColIndex("��ʼͣ��ʱ��")) = Format(dtpStartDate.Value, "yyyy-mm-dd HH:MM")
        .TextMatrix(lngRow, .ColIndex("����ͣ��ʱ��")) = Format(dtpEndDate.Value, "yyyy-mm-dd HH:MM")
        .TextMatrix(lngRow, .ColIndex("��ע")) = Trim(txtMemo.Text)
    End With
End Sub
Private Sub SetCmdEnable()
    '���ð�ť�ؼ���Enabled����
    With vsPlan
        If Trim(.TextMatrix(.Row, .ColIndex("��ʼͣ��ʱ��"))) <> "" Then
            cmdDeleteTime.Enabled = True
        Else
            cmdDeleteTime.Enabled = False
        End If
    End With
    
End Sub

Private Sub cmdClear_Click()
    If MsgBox("ע��:" & vbCrLf & "  ���Ƿ�ȫ�����еĹҺŰ���?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    End If
    vsOthers.Clear 1
    vsOthers.Rows = 2
    vsOthers.Cell(flexcpData, 1, 0, 1, vsOthers.Cols - 1) = ""
End Sub

Private Sub cmdDel_Click()
        'ɾ����
        Dim lngRow As Long
        With vsOthers
            If .TextMatrix(.Row, .ColIndex("�ű�")) <> "" Then
                If MsgBox("ע��:" & vbCrLf & " ���Ƿ����Ҫ�Ƴ��ű�Ϊ��" & .TextMatrix(.Row, .ColIndex("�ű�")) & "���ĹҺŰ�����?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
            End If
            If .Rows - 1 <= 1 Then
                .Cell(flexcpText, 1, 0, 1, .Cols - 1) = ""
                .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
            Else
                lngRow = .Row
                .RemoveItem lngRow
                
                If lngRow > .Rows - 1 Then
                    .Row = lngRow - 1
                Else
                    '.Row = lngRow + 1
                End If
            End If
        End With
End Sub


Private Sub cmdDeleteTime_Click()
    'ɾ����
    Dim lngRow As Long
    With vsList
        If .TextMatrix(.Row, .ColIndex("��ʼͣ��ʱ��")) <> "" Then
            If MsgBox("ע��:" & vbCrLf & " ���Ƿ����Ҫ�Ƴ�ʱ�䷶ΧΪ" & vbCrLf & .TextMatrix(.Row, .ColIndex("��ʼͣ��ʱ��")) & "��" & .TextMatrix(.Row, .ColIndex("����ͣ��ʱ��")) & vbCrLf & "��ͣ�üƻ�������?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Sub
            End If
        End If
        
        If Val(.Cell(flexcpData, .Row, .ColIndex("���"))) <> 0 Then
                mstrDelete��� = mstrDelete��� & "," & Val(.Cell(flexcpData, .Row, .ColIndex("���")))
        End If
        
        If .Rows - 1 <= 1 Then
            .Cell(flexcpText, 1, 0, 1, .Cols - 1) = ""
            .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
        Else
            lngRow = .Row
            .RemoveItem lngRow
            If lngRow > .Rows - 1 Then
                .Row = lngRow - 1
            Else
                '.Row = lngRow + 1
            End If
        End If
    End With
    Call ReFreshNo
End Sub
Private Sub ReFreshNo()
    '����ˢ�����
    Dim i As Long
    With vsList
        For i = 1 To .Rows - 1
            If Not .RowHidden(i) Then
                .TextMatrix(i, .ColIndex("���")) = i
            End If
        Next
    End With
End Sub
Public Function GetSplitStrUnionTable(ByVal strInputSplit As String, ByVal blnNum As Boolean, _
    ByVal intBandStart As Integer, ByVal strNotSplitTable As String, ByVal strNotSplitFieldName As String, _
    ByRef OutSplitValue() As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ָ��ֵ,�ֽ����صı������
    '���:intBandStart-�󶨵���ʼ��
    '����:OutSplitValue:����0-10��ֵ,δ�������,ֱ�����IN��ʽ
    '       strNotSplitValue:δ������ʱ,����ֵ
    '����:����SQL
    '����:���˺�
    '����:2010-09-08 10:46:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, i As Long, j As Long, strTemp As String
    Dim strSubItem  As String
    strSubItem = ""
    '�ȷֽ����,�ٲ���
    varData = Split(strInputSplit, ",")
    j = intBandStart: strTemp = ""
    For i = 0 To UBound(varData)
        If Len(strTemp) > 1990 And j - intBandStart <= 10 Then
            OutSplitValue(j - intBandStart) = Mid(strTemp, 2)
            If blnNum Then
                 strSubItem = strSubItem & vbCrLf & " Union ALL " & _
                " Select Column_Value From Table(f_Num2List([" & j & "]))   "
            Else
                 strSubItem = strSubItem & vbCrLf & " Union ALL " & _
                " Select Column_Value From Table(f_Str2List([" & j & "]))  "
            End If
            j = j + 1: strTemp = ""
        End If
        strTemp = strTemp & "," & IIf(blnNum, Val(varData(i)), varData(i))
    Next
    
    If strTemp <> "" Then
        If j - intBandStart > 10 Then
            If blnNum Then
                strSubItem = strSubItem & vbCrLf & " UNION ALL Select ID From " & strNotSplitTable & " Where " & strNotSplitFieldName & " in (" & Mid(strTemp, 2) & ")"
            Else
                strTemp = "'" & Replace(Mid(strTemp, 2), ",", "','") & "'"
                strSubItem = strSubItem & vbCrLf & " UNION ALL Select ID From " & strNotSplitTable & " Where " & strNotSplitFieldName & " in (" & strTemp & ")"
            End If
        Else
            OutSplitValue(j - intBandStart) = Mid(strTemp, 2)
            If blnNum Then
                 strSubItem = strSubItem & vbCrLf & " Union ALL " & _
                " Select Column_Value From Table(f_Num2List([" & j & "]))  "
            Else
                 strSubItem = strSubItem & vbCrLf & " Union ALL " & _
                " Select Column_Value From Table(f_Str2List([" & j & "]))   "
            End If
        End If
    End If
    If strSubItem <> "" Then strSubItem = Mid(strSubItem, 13)
    GetSplitStrUnionTable = strSubItem
End Function
Private Sub cmdOthers_Click()
    Dim strType  As String, strDept   As String, str��Ŀ   As String, strҽ�� As String
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strWhere As String, strTable As String, strTemp As String
    Dim strVarDept(0 To 10) As String, strVar��Ŀ(0 To 10) As String
    Dim strVarҽ��(0 To 10) As String, strVarҽ��1(0 To 10) As String
    Dim i As Long, lngRow As Long, blnFind As Boolean, blnNotMsg As Boolean
    Dim lngCount As Long
    Dim varData As Variant
    If frmRegistPlanInvalidationCons.ShowCons(Me, mlngModule, mstrPrivs, strType, strDept, str��Ŀ, strҽ��) = False Then
        Exit Sub
    End If
    
    On Error GoTo errHandle
    
    strTable = "": strWhere = ""
    If strType <> "" Then
        strTable = strTable & ",(" & " Select Column_Value as ���� From Table(f_Str2List([1]))) J "
        strWhere = strWhere & " And A.����=J.����"
    End If
    If strDept <> "" Then
        strTemp = GetSplitStrUnionTable(strDept, True, 2, "���ű�", "ID", strVarDept)
        strTable = strTable & vbCrLf & ",(" & strTemp & ") M "
        strWhere = strWhere & " And A.����ID=M.Column_Value"
    End If
    
    If str��Ŀ <> "" Then
        strTemp = GetSplitStrUnionTable(str��Ŀ, True, 13, "�շ���ĿĿ¼", "ID", strVar��Ŀ)
        strTable = strTable & vbCrLf & ",(" & strTemp & ") Q "
        strWhere = strWhere & " And A.��ĿID=Q.Column_Value"
    End If
    If strҽ�� <> "" Then
        varData = Split(strҽ��, "||")
        For i = 0 To UBound(varData)
            If i = 0 Then
                strTemp = GetSplitStrUnionTable(varData(i), True, 24, "��Ա��", "ID", strVarҽ��)
                strTable = strTable & vbCrLf & ",(" & strTemp & ") H "
                strWhere = strWhere & " And (A.ҽ��ID=H.Column_Value  "
            ElseIf i = 1 Then 'Ժ��ҽ��
                strTemp = GetSplitStrUnionTable(varData(i), False, 24, "�ҺŰ���", "ҽ������", strVarҽ��1)
                strTable = strTable & vbCrLf & ",(" & strTemp & ") M "
                strWhere = strWhere & " Or A.ҽ������=M.Column_Value  "
            End If
        Next
        If strWhere <> "" Then strWhere = strWhere & ")"
    End If
     

  
   strSQL = "" & _
    "   Select /*+ rule */ A.Id as ����ID,0 as �ƻ�ID,A.����,  A.����,  A.����id,  A.��Ŀid, A.ҽ������,  A.ҽ��id,  F.�޺���,  F.��Լ��,   " & _
    "           A.����,  A.��һ,  A.�ܶ�,  A.����,  A.����,  A.����,  A.����, " & _
    "           A.��������,  A.���﷽ʽ,  A.��ſ���,  A.��ʼʱ��,  A.��ֹʱ��,B.���� As ��Ŀ,D.���� As ���� " & _
    "   From �ҺŰ��� A,�շ���ĿĿ¼ B,���ű� D,�ҺŰ������� F " & strTable & _
    "   Where  A.��Ŀid=b.Id(+) And A.����id =d.Id(+) And a.id=F.����ID(+) And " & vbNewLine & _
    "          Decode(To_Char(Sysdate, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����', Null) =f.������Ŀ(+)" & vbNewLine & _
    "          And A.ID <>" & mlng����ID & strWhere & _
    "   Order by A.����,A.����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strType, _
        strVarDept(0), strVarDept(1), strVarDept(2), strVarDept(3), strVarDept(4), strVarDept(5), strVarDept(6), strVarDept(7), strVarDept(8), strVarDept(9), strVarDept(10), _
        strVar��Ŀ(0), strVar��Ŀ(1), strVar��Ŀ(2), strVar��Ŀ(3), strVar��Ŀ(4), strVar��Ŀ(5), strVar��Ŀ(6), strVar��Ŀ(7), strVar��Ŀ(8), strVar��Ŀ(9), strVar��Ŀ(10), _
        strVarҽ��(0), strVarҽ��(1), strVarҽ��(2), strVarҽ��(3), strVarҽ��(4), strVarҽ��(5), strVarҽ��(6), strVarҽ��(7), strVarҽ��(8), strVarҽ��(9), strVarҽ��(10), _
        strVarҽ��1(0), strVarҽ��1(1), strVarҽ��1(2), strVarҽ��1(3), strVarҽ��1(4), strVarҽ��1(5), strVarҽ��1(6), strVarҽ��1(7), strVarҽ��1(8), strVarҽ��1(9), strVarҽ��1(10), _
        "")
    With rsTemp
        blnNotMsg = False
        lngCount = .RecordCount
        Do While Not .EOF
                With vsOthers
                    blnFind = False
                    For i = 1 To .Rows - 1
                        If Val(.TextMatrix(i, .ColIndex("ID"))) = Val(Nvl(rsTemp!����ID)) Then
                            .Row = i
                            If Not blnNotMsg Then
                                If lngCount > 1 Then
                                    If MsgBox("ע��:" & vbCrLf & "    ���롺" & .TextMatrix(i, .ColIndex("�ű�")) & "���Ѿ�����," & vbCrLf & _
                                                "�˺ű𽫲��ټ���,���������ͬ���,�Ƿ�����ʾ?" & vbCrLf & _
                                                "���ǡ���ʾ����������ظ��ĺű�,������ʾ��" & vbCrLf & _
                                                "���񡻱�ʾ����������ظ��ĺű��������ʾ��", vbQuestion + vbDefaultButton2 + vbYesNo, gstrSysName) = vbYes Then
                                            blnNotMsg = True
                                    End If
                                Else
                                    Call MsgBox("ע��:" & vbCrLf & "    ���롺" & .TextMatrix(i, .ColIndex("�ű�")) & "���Ѿ�����,�����ټ���", vbInformation + vbDefaultButton1, gstrSysName)
                                End If
                            End If
                            
                            '��������
                            blnFind = True: Exit For
                        End If
                    Next
                    If blnFind = False Then
                       If .TextMatrix(.Rows - 1, .ColIndex("ID")) <> "" Then
                        .Rows = .Rows + 1
                    End If
                    lngRow = .Rows - 1
                    .TextMatrix(lngRow, .ColIndex("ID")) = Nvl(rsTemp!����ID)
                    .TextMatrix(lngRow, .ColIndex("����")) = Nvl(rsTemp!����)
                    .TextMatrix(lngRow, .ColIndex("�ű�")) = Nvl(rsTemp!����)
                    .TextMatrix(lngRow, .ColIndex("����")) = Nvl(rsTemp!����)
                    .TextMatrix(lngRow, .ColIndex("��Ŀ")) = Nvl(rsTemp!��Ŀ)
                    .TextMatrix(lngRow, .ColIndex("ҽ��")) = Nvl(rsTemp!ҽ������)
                    .TextMatrix(lngRow, .ColIndex("�޺�")) = Nvl(rsTemp!�޺���)
                    .TextMatrix(lngRow, .ColIndex("��Լ")) = Nvl(rsTemp!��Լ��)
                    .TextMatrix(lngRow, .ColIndex("����")) = Nvl(rsTemp!����)
                    .TextMatrix(lngRow, .ColIndex("��һ")) = Nvl(rsTemp!��һ)
                    .TextMatrix(lngRow, .ColIndex("�ܶ�")) = Nvl(rsTemp!�ܶ�)
                    .TextMatrix(lngRow, .ColIndex("����")) = Nvl(rsTemp!����)
                    .TextMatrix(lngRow, .ColIndex("����")) = Nvl(rsTemp!����)
                    .TextMatrix(lngRow, .ColIndex("����")) = Nvl(rsTemp!����)
                    .TextMatrix(lngRow, .ColIndex("����")) = Nvl(rsTemp!����)
                    .TextMatrix(lngRow, .ColIndex("������")) = IIf(Val(Nvl(rsTemp!��������)) = 0, "", "��")
                    .TextMatrix(lngRow, .ColIndex("���﷽ʽ")) = Nvl(rsTemp!���﷽ʽ)
                    .TextMatrix(lngRow, .ColIndex("IDS")) = Nvl(rsTemp!����ID) & "_" & Nvl(rsTemp!��ĿID) & "_" & Nvl(rsTemp!ҽ��ID)
                    .TextMatrix(lngRow, .ColIndex("Ӧ������")) = Read����Ӧ������(Val(Nvl(rsTemp!����ID)))    ' Nvl(rsTemp!��������)
                    
                    If Not IsNull(rsTemp!��ʼʱ��) Then
                        .TextMatrix(lngRow, .ColIndex("��Ч��Χ")) = Format(rsTemp!��ʼʱ��, "yyyy-MM-dd HH:mm:ss") & _
                            "��" & Format(rsTemp!��ֹʱ��, "yyyy-MM-dd HH:mm:ss")
                        .TextMatrix(lngRow, .ColIndex("��Ч��Χ")) = Replace(.TextMatrix(lngRow, .ColIndex("��Ч��Χ")), " 00:00:00", "")
                    End If
                    .TextMatrix(lngRow, .ColIndex("��ſ���")) = IIf(Val(Nvl(rsTemp!��ſ���)) = 0, "", "��")
                    .Row = lngRow
                    lngRow = lngRow + 1
                    End If
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

Private Sub cmdRestore_Click()
    If MsgBox("ע��:" & vbCrLf & "   ִ�лָ����ܺ�,����ȡ����ǰ������,�Ƿ����?", vbQuestion + vbDefaultButton2 + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    Call LoadData(True)
End Sub

Private Sub cmdSel_Click()
    If SelectStopMemo(txtMemo, "") = False Then Exit Sub
End Sub

Private Sub Form_Load()
    Call InitPanel
    Call RestoreWinState(Me, App.ProductName)
    mblnFirst = True
End Sub
Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If LoadData = False Then Unload Me: Exit Sub
    
    Call SetCtrlEnabled
    zlControl.ControlSetFocus dtpStartDate
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub
Private Sub SetCtrlEnabled()
    '���ÿؼ���Enabled����
    Dim ctl As Control
    For Each ctl In Me.Controls
        Select Case UCase(TypeName(ctl))
        Case "TEXTBOX"
            If ctl Is txtMemo Or ctl Is txtCode Then
                ctl.Enabled = True
            Else
                ctl.Enabled = False
            End If
            zlSetCtrolBackColor ctl
        Case UCase("ComboBox")
            ctl.Enabled = False
            zlSetCtrolBackColor ctl
        Case UCase("ListView")
            ctl.Enabled = False
            zlSetCtrolBackColor ctl
        Case UCase("DTPicker")
            If ctl Is dtpStartDate Or ctl Is dtpEndDate Then
                ctl.Enabled = True
            Else
                ctl.Enabled = False
            End If
           
        Case UCase("optionbutton"), UCase("CheckBox")
            If ctl Is chkClearHistory Then
                ctl.Enabled = True
            Else
                ctl.Enabled = False
            End If
        Case Else
        End Select
    Next
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Function CheckStopValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鰲��ʱ��ͣ�õĺϷ���
    '����:�Ϸ�,����True,���򷵻�False
    '����:���˺�
    '����:2010-09-07 14:06:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    If dtpStartDate.Value > dtpEndDate.Value Then
        ShowMsgbox "ע��:" & vbCrLf & "    ��ʼͣ�����ڴ����˽���ͣ������,����!"
        If dtpEndDate.Enabled And dtpEndDate.Visible Then dtpEndDate.SetFocus
        Exit Function
    End If
    
    If dtpStartDate.Value < zlDatabase.Currentdate Then
        ShowMsgbox "ע��:" & vbCrLf & "    ��ʼͣ������С���˵�ǰϵͳʱ��,����!"
        If dtpBegin.Enabled And dtpBegin.Visible Then dtpBegin.SetFocus
        Exit Function
    End If
    If zlCommFun.ActualLen(txtMemo.Text) > 100 Then
        ShowMsgbox "ע��:" & vbCrLf & "   ��ע���������ֻ������100���ַ���50������,����!"
        If txtMemo.Enabled And txtMemo.Visible Then txtMemo.SetFocus
        Exit Function
    End If
    If InStr(1, txtMemo.Text, "'") > 0 Then
        ShowMsgbox "ע��:" & vbCrLf & "   ��ע�������뵥����,����!"
        If txtMemo.Enabled And txtMemo.Visible Then txtMemo.SetFocus
        Exit Function
    End If
    '������е������Ƿ�Ϸ�:
    With vsList
        For i = 1 To .Rows - 1
            If Format(dtpStartDate.Value, "yyyy-mm-dd HH:MM") >= Trim(.TextMatrix(i, .ColIndex("��ʼͣ��ʱ��"))) _
               And Format(dtpStartDate.Value, "yyyy-mm-dd HH:MM") <= .TextMatrix(i, .ColIndex("����ͣ��ʱ��")) Then
               ShowMsgbox "ע��:" & vbCrLf & "    ��ʼͣ��ʱ���Ѿ��ڵ�" & i & "���д���,����!"
               If dtpBegin.Enabled And dtpBegin.Visible Then dtpBegin.SetFocus
               Exit Function
            End If
            If Format(dtpEndDate.Value, "yyyy-mm-dd HH:MM") >= Trim(.TextMatrix(i, .ColIndex("��ʼͣ��ʱ��"))) _
               And Format(dtpEndDate.Value, "yyyy-mm-dd HH:MM") <= .TextMatrix(i, .ColIndex("����ͣ��ʱ��")) Then
               ShowMsgbox "ע��:" & vbCrLf & "    ����ͣ��ʱ���Ѿ��ڵ�" & i & "���д���,����!"
               If dtpEnd.Enabled And dtpEnd.Visible Then dtpEnd.SetFocus
               Exit Function
            End If
            If Format(Trim(.TextMatrix(i, .ColIndex("��ʼͣ��ʱ��"))), "yyyy-mm-dd hh:mm") <= Format(dtpStartDate.Value, "yyyy-mm-dd HH:mm") Then
                If Format(Trim(.TextMatrix(i, .ColIndex("����ͣ��ʱ��"))), "yyyy-mm-dd hh:mm") >= Format(dtpStartDate.Value, "yyyy-mm-dd HH:mm") Then
                    ShowMsgbox "ע��:" & vbCrLf & "    ��" & i & "���е�ͣ��ʱ�䷶Χ�Ѿ������ڵ�ǰ�����õ�ͣ�÷�Χ��,����!"
                    If dtpEnd.Enabled And dtpEnd.Visible Then dtpEnd.SetFocus
                    Exit Function
                End If
            Else
                If Format(Trim(.TextMatrix(i, .ColIndex("��ʼͣ��ʱ��"))), "yyyy-mm-dd hh:mm") <= Format(dtpEndDate.Value, "yyyy-mm-dd HH:mm") Then
                    ShowMsgbox "ע��:" & vbCrLf & "    ��" & i & "���е�ͣ��ʱ�䷶Χ�Ѿ������ڵ�ǰ�����õ�ͣ�÷�Χ��,����!"
                    If dtpEnd.Enabled And dtpEnd.Visible Then dtpEnd.SetFocus
                    Exit Function
                End If
            End If
        Next
    End With
     CheckStopValied = True
End Function
Private Function CheckOtherPlan(ByVal str����ID As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���Ӧ�����������ŵ������Ƿ�Ϸ�
    '���:str����ID-�������ʱ,�ö��ŷָ�
    '����:
    '����:�Ϸ�,����true, ���򷵻�False
    '����:���˺�
    '����:2010-09-07 14:22:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strValues(0 To 10) As String, strSubItem As String, varData As Variant
    Dim strTemp As String, i As Long, j As Long, strEndDate As String, strStartDate As String
    Dim strValue(0 To 10)  As String
    On Error GoTo errHandle
     '�ȷֽ����,�ٲ���
    varData = Split(str����ID, ",")
    strTemp = "": j = 1
    For i = 0 To UBound(varData)
        If Len(strTemp) > 1990 And j <= 10 Then
            strValue(j - 1) = Mid(strTemp, 2)
            strSubItem = strSubItem & " Union ALL " & _
            " Select Column_Value as ����ID From Table(f_Num2List([" & j & "])) B "
            strTemp = "," & Val(varData(i)): j = j + 1
        Else
            strTemp = strTemp & "," & Val(varData(i))
        End If
    Next
    
    If strTemp <> "" Then
        If j - 1 > 10 Then
             strSubItem = strSubItem & " UNION ALL Select ID From �ҺŰ���ͣ��״̬ Where ����ID in (" & Mid(strTemp, 2) & ")"
        Else
            strValue(j - 1) = Mid(strTemp, 2)
            strSubItem = strSubItem & " Union ALL " & _
            " Select Column_Value as ����ID From Table(f_Num2List([" & j & "])) B "
        End If
    End If
    strSQL = "" & _
       "   Select /*+ Rule*/ B.�ű�,A.��ʼֹͣʱ��,A.����ֹͣʱ��  " & _
       "   From �ҺŰ���ͣ��״̬ A,�ҺŰ��� B, (" & Mid(strSubItem, 11) & ") D" & _
       "   Where A.����ID = D.����ID and A.����ID=b.ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str����ID, strValue(0), strValue(1), strValue(2), strValue(3), strValue(4), strValue(5), strValue(6), strValue(7), strValue(8), strValue(9), strValue(10))
    With rsTemp
        Do While Not .EOF
            strStartDate = Format(rsTemp!��ʼֹͣʱ��, "yyyy-mm-dd HH:MM")
            strEndDate = Format(rsTemp!����ֹͣʱ��, "yyyy-mm-dd HH:MM")
            '������е������Ƿ�Ϸ�:
            With vsList
                For i = 1 To .Rows - 1
                    If Val(.RowData(i)) <> 1 Then
                        If Trim(.TextMatrix(i, .ColIndex("��ʼͣ��ʱ��"))) >= strStartDate _
                           And .TextMatrix(i, .ColIndex("��ʼͣ��ʱ��")) <= strEndDate Then
                           ShowMsgbox "ע��:" & vbCrLf & "    �ű�Ϊ��" & Nvl(rsTemp!�ű�) & "����ͣ��ʱ��(" & Trim(.TextMatrix(i, .ColIndex("��ʼͣ��ʱ��"))) & " ~ " & Trim(.TextMatrix(i, .ColIndex("����ͣ��ʱ��"))) & ")�Ѿ�����,����!"
                           .Row = i: .Col = .ColIndex("��ʼͣ��ʱ��")
                           If vsList.Enabled And vsList.Visible Then vsList.SetFocus
                           Exit Function
                        End If
                        If Trim(.TextMatrix(i, .ColIndex("����ͣ��ʱ��"))) >= strStartDate _
                           And .TextMatrix(i, .ColIndex("����ͣ��ʱ��")) <= strEndDate Then
                           ShowMsgbox "ע��:" & vbCrLf & "    ����ͣ��ʱ���Ѿ��ڵ�" & i & "���д���,����!"
                           .Row = i: .Col = .ColIndex("����ͣ��ʱ��")
                           If vsList.Enabled And vsList.Visible Then vsList.SetFocus
                           Exit Function
                        End If
                        
                        If strStartDate >= Trim(.TextMatrix(i, .ColIndex("��ʼͣ��ʱ��"))) And _
                          strStartDate >= Trim(.TextMatrix(i, .ColIndex("����ͣ��ʱ��"))) Then
                           ShowMsgbox "ע��:" & vbCrLf & "    ��ʼͣ��ʱ���Ѿ��ڵ�" & i & "���д���,����!"
                           .Row = i: .Col = .ColIndex("����ͣ��ʱ��")
                           If vsList.Enabled And vsList.Visible Then vsList.SetFocus
                           Exit Function
                        End If
                        If strEndDate >= Trim(.TextMatrix(i, .ColIndex("��ʼͣ��ʱ��"))) And _
                          strEndDate >= Trim(.TextMatrix(i, .ColIndex("����ͣ��ʱ��"))) Then
                           ShowMsgbox "ע��:" & vbCrLf & "    ����ͣ��ʱ���Ѿ��ڵ�" & i & "���д���,����!"
                           .Row = i: .Col = .ColIndex("����ͣ��ʱ��")
                           If vsList.Enabled And vsList.Visible Then vsList.SetFocus
                           Exit Function
                        End If
                    End If
                Next
            End With
            .MoveNext
        Loop
    End With
    CheckOtherPlan = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function SaveData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������е�����
    '����:�ɹ�,����true,���򷵻�false
    '����:���˺�
    '����:2010-09-07 14:54:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPro As Collection, strSQL As String, i As Long, j As Long, str����ID As String
    Dim strStartDate As String, strEndDate As String, str��ע As String
    Dim cll���� As Collection
    
    Set cllPro = New Collection: Set cll���� = New Collection
    With vsOthers
        str����ID = "," & mlng����ID
        For j = 1 To .Rows - 1
            If Val(.TextMatrix(j, .ColIndex("ID"))) <> 0 Then
                If Len(str����ID) > 1920 Then
                    str����ID = Mid(str����ID, 2)
                   cll����.Add str����ID
                    str����ID = ""
                End If
                str����ID = str����ID & "," & Val(.TextMatrix(j, .ColIndex("ID")))
            End If
        Next
        If str����ID <> "" Then
            str����ID = Mid(str����ID, 2)
            cll����.Add str����ID
        End If
    End With
    
    With vsList
        '�ȴ���ɾ������
        If mstrDelete��� <> "" Then
            mstrDelete��� = Mid(mstrDelete���, 2)
            For j = 1 To cll����.Count
                'Zl_�ҺŰ���ͣ��״̬_Delete
                strSQL = "Zl_�ҺŰ���ͣ��״̬_Delete("
                '  ����id_In     In �ҺŰ���ͣ��״̬.����id%Type,
                strSQL = strSQL & "" & mlng����ID & ","
                '  ���_In       In Varchar2, --�ö��ŷָ�
                strSQL = strSQL & "'" & mstrDelete��� & "',"
                '  ��������id_In In Varchar2 --�ö��ŷָ�
                strSQL = strSQL & "'" & cll����(j) & "')"
                zlAddArray cllPro, strSQL
            Next
 
        End If
        '�������ӵ�����
        For i = 1 To .Rows - 1
            If Val(.RowData(i)) = 0 And Trim(.TextMatrix(i, .ColIndex("��ʼͣ��ʱ��"))) <> "" Then '0���������ӵ�ͣ������
                str����ID = "," & mlng����ID
                strStartDate = .TextMatrix(i, .ColIndex("��ʼͣ��ʱ��"))
                strEndDate = .TextMatrix(i, .ColIndex("����ͣ��ʱ��"))
                str��ע = .TextMatrix(i, .ColIndex("��ע"))
                For j = 1 To cll����.Count
                    If chkClearHistory.Value = 1 Then
                        '    Zl_�ҺŰ���ͣ��״̬_Clear(����id_In Varchar2) Is
                        strSQL = "Zl_�ҺŰ���ͣ��״̬_Clear('" & cll����(j) & "')"
                        zlAddArray cllPro, strSQL
                    End If
                     'Zl_�ҺŰ���ͣ��״̬_Insert
                     strSQL = "Zl_�ҺŰ���ͣ��״̬_Insert("
                    '��ʼֹͣʱ��_In In �ҺŰ���ͣ��״̬.��ʼֹͣʱ��%Type,
                    strSQL = strSQL & "to_date('" & strStartDate & "','yyyy-mm-dd HH24:mi'),"
                     '����ֹͣʱ��_In In �ҺŰ���ͣ��״̬.����ֹͣʱ��%Type,
                    strSQL = strSQL & "to_date('" & strEndDate & "','yyyy-mm-dd HH24:mi'),"
                     '�ƶ���_In       In �ҺŰ���ͣ��״̬.�ƶ���%Type,
                    strSQL = strSQL & "'" & UserInfo.���� & "',"
                     '��ע_In         In �ҺŰ���ͣ��״̬.��ע%Type,
                    strSQL = strSQL & "'" & str��ע & "',"
                     '����id_In       In Varchar2 --�ö��ŷָ�
                    strSQL = strSQL & "'" & cll����(j) & "')"
                    zlAddArray cllPro, strSQL
                Next
            End If
        Next
    End With
    Err = 0: On Error GoTo Errhand:
    zlExecuteProcedureArrAy cllPro, Me.Caption
    SaveData = True
    Exit Function
Errhand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function
Private Function IsValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������������Ƿ�Ϸ�
    '���:
    '����:
    '����:�Ϸ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-09-07 16:03:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, blnFind As Boolean
    On Error GoTo errHandle
    With vsList
        For lngRow = 1 To .Rows - 1
             If .TextMatrix(lngRow, .ColIndex("��ʼͣ��ʱ��")) <> "" And InStr("0,2", Val(.RowData(lngRow))) > 0 Then
                    blnFind = True: Exit For
             End If
        Next
        If blnFind = False And mstrDelete��� = "" Then
            MsgBox "û�м���ͣ��ʱ�䣬���ܼ���!", vbInformation + vbOKOnly, gstrSysName
            If dtpStartDate.Enabled And dtpStartDate.Visible Then dtpStartDate.SetFocus
            Exit Function
        End If
    End With
    IsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub cmdOk_Click()
     If IsValied = False Then Exit Sub
     If SaveData = False Then Exit Sub
     mblnSucces = True
    Unload Me
End Sub
Public Function SelectItem(ByVal strInput As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������ֵ��ѡ����ص�����(���ڶ�ѡ)
    '���:intIndex-����
    '       strInput-�����ֵ
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-09-07 10:21:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strCode As String, blnCancel As Boolean, rsTemp As ADODB.Recordset
    Dim strDept As String, strDeptWhere As String, strTable As String
    Dim strLike As String, strWhere As String, bytCode As Byte, lngRow As Long
    Dim strTittle As String, strValue(0 To 10) As String
    Dim vRect As RECT, j As Long, i As Long
    If Trim(strInput) = "" Then Exit Function
     On Error GoTo Hd
    bytCode = Val(zlDatabase.GetPara("���뷽ʽ", , , 0))
    strLike = IIf(zlDatabase.GetPara("����ƥ��") = "0", "%", "")
    '���ܣ��๦��ѡ����,ʹ��ADO.Command��,����ʹ��[x]����
    '������
    '     frmParent=��ʾ�ĸ�����
    '     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
    '     bytStyle=ѡ�������
    '       Ϊ0ʱ:�б���:ID,��
    '       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
    '       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
    '     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
    '     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
    '     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
    '             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
    '             bytStyle=1ʱ,�����Ǳ��������
    '     strNote=ѡ������˵������
    '     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
    '     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
    '     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
    '     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
    '     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
    '     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
    '     arrInput=��Ӧ�ĸ���SQL����ֵ,��˳����,����Ϊ��ȷ����
    '���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
    '˵����
    '     1.ID���ϼ�ID����Ϊ�ַ�������
    '     2.ĩ�����ֶβ�Ҫ����ֵ
    'Ӧ�ã������ڸ������������������Ǻܴ��ѡ����,����ƥ���б��
    If strInput <> "" Then
        strCode = strLike & strInput & "%"
        If zlCommFun.IsCharAlpha(strInput) Then
            strWhere = " And ( B.���� Like upper([1]))"
        ElseIf IsNumeric(strInput) Or zlCommFun.IsNumOrChar(strInput) Then
            strWhere = " And A.���� Like upper([1])"
        Else
            strWhere = " And (A.ҽ������ Like [1] Or exists(Select 1 From ��Ա�� where ���� like [1] and a.ҽ��ID=C.id )  or B.���� like [1] or  A.���� like [1]  )"
        End If
    Else
        strWhere = ""
    End If

    strSQL = "" & _
    "   Select Distinct A.id, A.����,A.����,b.���� as ����,C.���� as ��Ŀ,nvl(D.����,A.ҽ������) as ҽ�� " & _
    "   From �ҺŰ��� A,���ű� B,�շ���ĿĿ¼ C,��Ա�� D" & _
    "   Where A.����ID=B.id  And A.��ĿID=C.id and A.ҽ��ID=D.id(+) And A.ID <>" & mlng����ID & strWhere & _
    "           And rownum<101 " & _
    "   Order by ����,����"
    
    strTittle = "�ҺŰ���"
    
    vRect = zlControl.GetControlRect(txtCode.Hwnd)
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, strTittle & "ѡ��", False, "", "��ѡ��", False, False, True, vRect.Left, vRect.Top, txtCode.Height, blnCancel, True, True, strCode)
    
    
    If blnCancel = True Then
        If txtCode.Enabled And txtCode.Visible Then txtCode.SetFocus
        Exit Function
    End If
    If rsTemp Is Nothing Then
        MsgBox "û���ҵ�����������" & strTittle & "������!", vbInformation + vbOKOnly, gstrSysName
        If txtCode.Enabled And txtCode.Visible Then txtCode.SetFocus
        Call txtCode_GotFocus
        Exit Function
    End If
    If rsTemp.State <> 1 Then
        MsgBox "û���ҵ�����������" & strTittle & "������!", vbInformation + vbOKOnly, gstrSysName
        If txtCode.Enabled And txtCode.Visible Then txtCode.SetFocus
        Call txtCode_GotFocus
        Exit Function
    End If
    Dim strValues As String, strSubItem As String, blnFind As Boolean, blnNotMsg As Boolean, lngCount As Long
    
    
    With rsTemp
        strValues = "": j = 1: lngCount = .RecordCount
        blnNotMsg = False
        Do While Not .EOF
            '�ȼ���Ƿ�������������Ѿ�������
            With vsOthers
                blnFind = False
                For i = 1 To .Rows - 1
                    If Trim(.TextMatrix(i, .ColIndex("ID"))) = Val(Nvl(rsTemp!ID)) Then
                            If Not blnNotMsg Then
                                If lngCount > 1 Then
                                    If MsgBox("ע��:" & vbCrLf & "    ���롺" & .TextMatrix(i, .ColIndex("�ű�")) & "���Ѿ�����," & vbCrLf & _
                                                "�˺ű𽫲��ټ���,���������ͬ���,�Ƿ�����ʾ?" & vbCrLf & _
                                                "���ǡ���ʾ����������ظ��ĺű�,������ʾ��" & vbCrLf & _
                                                "���񡻱�ʾ����������ظ��ĺű��������ʾ��", vbQuestion + vbDefaultButton2 + vbYesNo, gstrSysName) = vbYes Then
                                            blnNotMsg = True
                                    End If
                                Else
                                    Call MsgBox("ע��:" & vbCrLf & "    ���롺" & .TextMatrix(i, .ColIndex("�ű�")) & "���Ѿ�����,�����ټ���", vbInformation + vbDefaultButton1, gstrSysName)
                                    .Row = i
                                End If
                            End If
                           blnFind = True: Exit For
                    End If
                Next
            End With
            
            If blnFind = False Then
                If Len(strValues) > 1990 And j <= 10 Then
                    strValue(j - 1) = Mid(strValues, 2)
                    strSubItem = strSubItem & " Union ALL " & _
                    " Select Column_Value as ����ID From Table(f_Num2List([" & j & "])) B "
                    strValues = "," & Val(Nvl(rsTemp!ID)): j = j + 1
                Else
                    strValues = strValues & "," & Val(Nvl(rsTemp!ID))
                End If
            End If
            .MoveNext
        Loop
        If strValues <> "" Then
            If j - 1 > 10 Then
                 strSubItem = strSubItem & " UNION ALL Select ID From �ҺŰ��� Where ID in (" & Mid(strValues, 2) & ")"
            Else
                strValue(j - 1) = Mid(strValues, 2)
                strSubItem = strSubItem & " Union ALL " & _
                "   Select Column_Value as ����ID From Table(f_Num2List([" & j & "])) B "
            End If
        End If
    End With
    If strSubItem = "" Then Exit Function
    
    strSQL = "" & _
        "   Select /*+ rule */ A.Id as ����ID,0 as �ƻ�ID,A.����,  A.����,  A.����id,  A.��Ŀid, A.ҽ������,  A.ҽ��id, F.�޺���,  F.��Լ��,   " & _
        "           A.����,  A.��һ,  A.�ܶ�,  A.����,  A.����,  A.����,  A.����, " & _
        "           A.��������,  A.���﷽ʽ,  A.��ſ���,  A.��ʼʱ��,  A.��ֹʱ��,B.���� As ��Ŀ,D.���� As ����  " & _
        "   From �ҺŰ��� A,�շ���ĿĿ¼ B,���ű� D , �ҺŰ������� F,(" & Mid(strSubItem, 11) & ") M" & _
        "   Where  A.��Ŀid=b.Id(+) And A.����id =d.Id(+) And A.id=M.����ID  And a.Id = f.����id(+) And" & _
        "   Decode(To_Char(sysdate, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����', Null) =f.������Ŀ(+)" & vbNewLine & _
        "   Order by A.����,A.����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strValue(0), strValue(1), strValue(2), strValue(3), strValue(4), strValue(5), strValue(6), strValue(7), strValue(8), strValue(9), strValue(10))
    With rsTemp
        Do While Not .EOF
                With vsOthers
                    If .TextMatrix(.Rows - 1, .ColIndex("ID")) <> "" Then
                        .Rows = .Rows + 1
                    End If
                    lngRow = .Rows - 1
                    .TextMatrix(lngRow, .ColIndex("ID")) = Nvl(rsTemp!����ID)
                    .TextMatrix(lngRow, .ColIndex("����")) = Nvl(rsTemp!����)
                    .TextMatrix(lngRow, .ColIndex("�ű�")) = Nvl(rsTemp!����)
                    .TextMatrix(lngRow, .ColIndex("����")) = Nvl(rsTemp!����)
                    .TextMatrix(lngRow, .ColIndex("��Ŀ")) = Nvl(rsTemp!��Ŀ)
                    .TextMatrix(lngRow, .ColIndex("ҽ��")) = Nvl(rsTemp!ҽ������)
                    .TextMatrix(lngRow, .ColIndex("�޺�")) = Nvl(rsTemp!�޺���)
                    .TextMatrix(lngRow, .ColIndex("��Լ")) = Nvl(rsTemp!��Լ��)
                    .TextMatrix(lngRow, .ColIndex("����")) = Nvl(rsTemp!����)
                    .TextMatrix(lngRow, .ColIndex("��һ")) = Nvl(rsTemp!��һ)
                    .TextMatrix(lngRow, .ColIndex("�ܶ�")) = Nvl(rsTemp!�ܶ�)
                    .TextMatrix(lngRow, .ColIndex("����")) = Nvl(rsTemp!����)
                    .TextMatrix(lngRow, .ColIndex("����")) = Nvl(rsTemp!����)
                    .TextMatrix(lngRow, .ColIndex("����")) = Nvl(rsTemp!����)
                    .TextMatrix(lngRow, .ColIndex("����")) = Nvl(rsTemp!����)
                    .TextMatrix(lngRow, .ColIndex("������")) = IIf(Val(Nvl(rsTemp!��������)) = 0, "", "��")
                    .TextMatrix(lngRow, .ColIndex("���﷽ʽ")) = Nvl(rsTemp!���﷽ʽ)
                    .TextMatrix(lngRow, .ColIndex("IDS")) = Nvl(rsTemp!����ID) & "_" & Nvl(rsTemp!��ĿID) & "_" & Nvl(rsTemp!ҽ��ID)
                    .TextMatrix(lngRow, .ColIndex("Ӧ������")) = Read����Ӧ������(Val(Nvl(rsTemp!����ID)))    ' Nvl(rsTemp!��������)
                    
                    If Not IsNull(rsTemp!��ʼʱ��) Then
                        .TextMatrix(lngRow, .ColIndex("��Ч��Χ")) = Format(rsTemp!��ʼʱ��, "yyyy-MM-dd HH:mm:ss") & _
                            "��" & Format(rsTemp!��ֹʱ��, "yyyy-MM-dd HH:mm:ss")
                        .TextMatrix(lngRow, .ColIndex("��Ч��Χ")) = Replace(.TextMatrix(lngRow, .ColIndex("��Ч��Χ")), " 00:00:00", "")
                    End If
                    .TextMatrix(lngRow, .ColIndex("��ſ���")) = IIf(Val(Nvl(rsTemp!��ſ���)) = 0, "", "��")
                   ' .TextMatrix(lngRow, .ColIndex("ͣ������")) = Nvl(rsTemp!ͣ������)
'                    If Trim(.TextMatrix(lngRow, .ColIndex("ͣ������"))) <> "" Then
'                        .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = vbRed
'                    End If
                    .Row = lngRow
                    lngRow = lngRow + 1
                End With
            .MoveNext
        Loop
    End With
   Call txtCode_GotFocus
    If txtCode.Enabled And txtCode.Visible Then txtCode.SetFocus
    SelectItem = True
    Exit Function
Hd:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Function



Private Function Read����Ӧ������(ByVal lngID As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ������
    '���:lngID-ID
    '����:
    '����:
    '����:���˺�
    '����:2009-09-14 22:39:14
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim strSQL As String
    
    On Error GoTo errH
    If lngID = 0 Then Exit Function
    
    If mrsRoom Is Nothing Then
        strSQL = "Select ��������,�ű�ID From �ҺŰ�������"
        Set mrsRoom = New Recordset
        Call zlDatabase.OpenRecordset(mrsRoom, strSQL, Me.Caption)
    End If
    With mrsRoom
        .Filter = "�ű�ID=" & lngID
        If .RecordCount = 0 Then Exit Function
        Do While Not .EOF
            Read����Ӧ������ = Read����Ӧ������ & ";" & !��������
            .MoveNext
        Loop
    End With
    Read����Ӧ������ = Mid(Read����Ӧ������, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
    zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "ͣ�ð���-ͣ�üƻ�", True, , InStr(1, mstrPrivs, ";��������;") > 0
    zl_vsGrid_Para_Save mlngModule, vsOthers, Me.Caption, "ͣ�ð���-�ҺŰ���", True, , InStr(1, mstrPrivs, ";��������;") > 0
    
End Sub

Private Sub imgColList_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlControl.GetControlRect(picImgList.Hwnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgList.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsList, lngLeft, lngTop, imgColList.Height)
    zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "ͣ�ð���-�ҺŰ���", True, , InStr(1, mstrPrivs, ";��������;") > 0
End Sub
Private Sub picImgList_Click()
    Call picImgList_Click
End Sub
 
Private Sub picCmd_Resize()
    Err = 0: On Error Resume Next:
    With picCmd
        fraSplit(1).Left = .ScaleLeft: fraSplit(1).Width = .ScaleWidth
        cmdCancel.Left = .ScaleWidth - cmdCancel.Width - 100
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 50
        chkClearHistory.Left = cmdOK.Left - chkClearHistory.Width * 2
        If chkClearHistory.Left < 0 Then chkClearHistory.Left = .ScaleLeft
    End With
End Sub
Private Sub picOthers_Resize()
   Err = 0: On Error Resume Next:
    With picOthers
        vsOthers.Left = .ScaleLeft + 50: vsOthers.Width = .ScaleWidth - vsList.Left * 2
        vsOthers.Height = .ScaleHeight - vsOthers.Top
    End With
End Sub

Private Sub picStop_Resize()
   Err = 0: On Error Resume Next:
    With picStop
        vsList.Left = .ScaleLeft + 50: vsList.Width = .ScaleWidth - vsList.Left * 2
        vsList.Height = .ScaleHeight - vsList.Top
    End With
End Sub

Private Sub txtCode_GotFocus()
    zlControl.TxtSelAll txtCode
End Sub
Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If SelectItem(Trim(txtCode.Text)) = False Then
        Exit Sub
    End If
End Sub

Private Sub txtMemo_Change()
    txtMemo.Tag = ""
End Sub

Private Sub txtMemo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If txtMemo.Tag <> "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If txtMemo.Text = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If SelectStopMemo(txtMemo, Trim(txtMemo.Text)) = False Then Exit Sub
End Sub

Private Sub vsList_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "ͣ�ð���-ͣ�üƻ�", True, InStr(1, mstrPrivs, ";��������;") > 0
End Sub

Private Sub vsList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "ͣ�ð���-ͣ�üƻ�", True, InStr(1, mstrPrivs, ";��������;") > 0
End Sub

Private Sub vsOthers_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsOthers, Me.Caption, "ͣ�ð���-�ҺŰ���", True, InStr(1, mstrPrivs, ";��������;") > 0
End Sub

Private Sub vsOthers_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsOthers, Me.Caption, "ͣ�ð���-�ҺŰ���", True, InStr(1, mstrPrivs, ";��������;") > 0
End Sub
Private Function SelectStopMemo(ByVal objCtl As Control, Optional strKey As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ѡ��ͣ��ԭ��
    '���:strKey-����ֵ
    '����:
    '����:
    '����:���˺�
    '����:2011-11-08 15:00:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim vRect As RECT, strWhere As String, bytStyle As Byte
    Dim sngX As Single, sngY As Single, lngH As Long
    Dim blnCancel As Boolean
 
    On Error GoTo errH
    bytStyle = 0
    strWhere = " "
    If strKey <> "" Then
        strWhere = " Where 1=1 "
        If zlCommFun.IsCharChinese(strKey) Then
            strWhere = strWhere & " And ���� like [1]  Order by ����"
        ElseIf zlCommFun.IsCharAlpha(strKey) Then
            strWhere = strWhere & " And ���� like upper([1]) Order by ����"
        ElseIf zlCommFun.IsNumOrChar(strKey) Then
            strWhere = strWhere & " And ���� like upper([1])  Order by ����"
        Else
            strWhere = strWhere & " And  (���� like [1] or ���� like upper([1]) or ���� like upper([1])) Order by ����"
        End If
        bytStyle = 0
        strKey = gstrLike & strKey & "%"
    End If
    
    strSQL = "" & _
    "   Select Rownum as ID,����,����,����,decode(ȱʡ��־,1,'��','') as ȱʡ" & _
    "   From ����ͣ��ԭ��" & _
        strWhere
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strKey)
    
    'ShowSelect:
    '���ܣ��๦��ѡ����
    '������
    '     frmParent=��ʾ�ĸ�����
    '     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
    '     bytStyle=ѡ�������
    '       Ϊ0ʱ:�б���:ID,��
    '       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
    '       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
    '     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
    '     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
    '     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
    '             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
    '             bytStyle=1ʱ,�����Ǳ��������
    '     strNote=ѡ������˵������
    '     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
    '     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
    '     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
    '     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
    '     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
    '     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
    '���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
    '˵����
    '     1.ID���ϼ�ID����Ϊ�ַ�������
    '     2.ĩ�����ֶβ�Ҫ����ֵ
    'Ӧ�ã������ڸ������������������Ǻܴ��ѡ����,����ƥ���б�ȡ�
    vRect = zlControl.GetControlRect(objCtl.Hwnd)
    lngH = objCtl.Height
    sngX = vRect.Left - 15: sngY = vRect.Top
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, bytStyle, "ͣ��ԭ��ѡ��", IIf(bytStyle = 2, True, False), "", "��ѡ�����������ͣ��ԭ��", IIf(bytStyle = 2, True, False), True, True, sngX, sngY, lngH, blnCancel, False, True, strKey)
    If blnCancel Then
        If objCtl.Enabled And objCtl.Visible Then objCtl.SetFocus
        zlControl.TxtSelAll objCtl
        Exit Function
    End If
    If rsTemp Is Nothing Then
        ShowMsgbox "�����ڷ���������ͣ��ԭ��,����!"
        If objCtl.Enabled And objCtl.Visible Then objCtl.SetFocus
        zlControl.TxtSelAll objCtl
        Exit Function
        Exit Function
    End If
    If rsTemp.State <> 1 Then
        If objCtl.Enabled And objCtl.Visible Then objCtl.SetFocus
        zlControl.TxtSelAll objCtl
        Exit Function
        Exit Function
    End If
    With rsTemp
        objCtl.Text = Nvl(!����): objCtl.Tag = Nvl(!ID)
    End With
    If objCtl.Enabled And objCtl.Visible Then objCtl.SetFocus
    zlControl.TxtSelAll objCtl
    zlCommFun.PressKey vbKeyTab
    SelectStopMemo = True
    Exit Function
errH:
    If ErrCenter = 1 Then Resume
End Function


