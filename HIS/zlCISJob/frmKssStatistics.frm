VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmKssStatistics 
   Caption         =   "����ҩ��ͳ�Ʒ���"
   ClientHeight    =   9435
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   13905
   DrawStyle       =   6  'Inside Solid
   Icon            =   "frmKssStatistics.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9435
   ScaleWidth      =   13905
   StartUpPosition =   1  '����������
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picOtherSub 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4695
      Index           =   5
      Left            =   10770
      ScaleHeight     =   4695
      ScaleWidth      =   915
      TabIndex        =   52
      Top             =   795
      Width           =   915
      Begin VB.PictureBox picFilter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   5
         Left            =   45
         ScaleHeight     =   810
         ScaleWidth      =   11295
         TabIndex        =   173
         Top             =   45
         Width           =   11295
         Begin VB.PictureBox picOpt 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   420
            Index           =   14
            Left            =   4035
            ScaleHeight     =   420
            ScaleWidth      =   1395
            TabIndex        =   196
            TabStop         =   0   'False
            Top             =   345
            Width           =   1395
            Begin VB.OptionButton optType 
               BackColor       =   &H80000005&
               Caption         =   "����������"
               Height          =   180
               Index           =   28
               Left            =   0
               TabIndex        =   197
               Top             =   210
               Width           =   1290
            End
            Begin VB.OptionButton optType 
               BackColor       =   &H80000005&
               Caption         =   "����ϱ�׼"
               Height          =   180
               Index           =   27
               Left            =   0
               TabIndex        =   183
               Top             =   0
               Value           =   -1  'True
               Width           =   1335
            End
         End
         Begin VB.CommandButton cmdILL 
            Caption         =   "��"
            Height          =   255
            Left            =   3660
            TabIndex        =   175
            TabStop         =   0   'False
            ToolTipText     =   "Ctrl+I"
            Top             =   420
            Width           =   270
         End
         Begin VB.PictureBox picOpt 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   510
            Index           =   13
            Left            =   75
            ScaleHeight     =   510
            ScaleWidth      =   735
            TabIndex        =   194
            TabStop         =   0   'False
            Top             =   315
            Width           =   735
            Begin VB.OptionButton optType 
               BackColor       =   &H80000005&
               Caption         =   "��ҽ"
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
               Index           =   26
               Left            =   0
               TabIndex        =   195
               Top             =   225
               Width           =   690
            End
            Begin VB.OptionButton optType 
               BackColor       =   &H80000005&
               Caption         =   "��ҽ"
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
               Index           =   25
               Left            =   0
               TabIndex        =   181
               Top             =   0
               Value           =   -1  'True
               Width           =   690
            End
         End
         Begin VB.TextBox txtILL 
            Height          =   300
            IMEMode         =   1  'ON
            Left            =   825
            MaxLength       =   100
            TabIndex        =   182
            Top             =   405
            Width           =   3120
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "ͳ��(&T)"
            Height          =   300
            Index           =   9
            Left            =   9705
            TabIndex        =   187
            Top             =   405
            Width           =   960
         End
         Begin VB.PictureBox picOpt 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   9
            Left            =   8280
            ScaleHeight     =   255
            ScaleWidth      =   1380
            TabIndex        =   184
            TabStop         =   0   'False
            Top             =   465
            Width           =   1380
            Begin VB.OptionButton optType 
               BackColor       =   &H80000005&
               Caption         =   "ƽ��"
               Height          =   180
               Index           =   20
               Left            =   0
               TabIndex        =   186
               Top             =   0
               Value           =   -1  'True
               Width           =   660
            End
            Begin VB.OptionButton optType 
               BackColor       =   &H80000005&
               Caption         =   "���"
               Height          =   180
               Index           =   19
               Left            =   690
               TabIndex        =   189
               Top             =   0
               Width           =   660
            End
         End
         Begin VB.ComboBox cboTimCount 
            Height          =   300
            Index           =   5
            Left            =   825
            Style           =   2  'Dropdown List
            TabIndex        =   177
            Top             =   0
            Width           =   1365
         End
         Begin VB.TextBox txtNum 
            Height          =   300
            Index           =   3
            Left            =   6660
            MaxLength       =   3
            TabIndex        =   185
            Text            =   "10"
            Top             =   405
            Width           =   465
         End
         Begin MSComCtl2.DTPicker dtpCountE 
            Height          =   300
            Index           =   5
            Left            =   4245
            TabIndex        =   179
            Top             =   0
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   124059651
            CurrentDate     =   41774
         End
         Begin MSComCtl2.DTPicker dtpCountS 
            Height          =   300
            Index           =   5
            Left            =   2565
            TabIndex        =   178
            Top             =   0
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   124059651
            CurrentDate     =   41774
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "��������"
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
            Index           =   71
            Left            =   7440
            TabIndex        =   193
            Top             =   465
            Width           =   795
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "��                 ��"
            Height          =   180
            Index           =   69
            Left            =   2310
            TabIndex        =   192
            Top             =   60
            Width           =   2970
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "ͳ��ʱ��"
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
            Index           =   68
            Left            =   0
            TabIndex        =   191
            Top             =   60
            Width           =   795
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "��������"
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
            Index           =   60
            Left            =   5805
            TabIndex        =   190
            Top             =   450
            Width           =   795
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsIllDruUse 
         Height          =   690
         Left            =   60
         TabIndex        =   188
         Top             =   1350
         Width           =   7710
         _cx             =   13600
         _cy             =   1217
         Appearance      =   2
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
         BackColorSel    =   16444122
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmKssStatistics.frx":6852
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
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
         OwnerDraw       =   1
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "˵��"
         Height          =   180
         Index           =   67
         Left            =   285
         TabIndex        =   198
         Top             =   2265
         Width           =   360
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "ҽ������ĳ����������ҩ�ɱ�ͳ��"
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
         Index           =   7
         Left            =   195
         TabIndex        =   53
         Top             =   1005
         Width           =   3825
      End
   End
   Begin VB.PictureBox picOtherSub 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6060
      Index           =   3
      Left            =   8490
      ScaleHeight     =   6060
      ScaleWidth      =   825
      TabIndex        =   48
      Top             =   735
      Width           =   825
      Begin VB.PictureBox picFilter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   780
         Index           =   3
         Left            =   0
         ScaleHeight     =   780
         ScaleWidth      =   11865
         TabIndex        =   124
         Top             =   0
         Width           =   11865
         Begin VB.CommandButton cmdDept 
            Caption         =   "��"
            Height          =   255
            Index           =   6
            Left            =   5325
            TabIndex        =   125
            TabStop         =   0   'False
            ToolTipText     =   "Ctrl+D"
            Top             =   420
            Width           =   285
         End
         Begin VB.PictureBox picOpt 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   7
            Left            =   8115
            ScaleHeight     =   255
            ScaleWidth      =   3525
            TabIndex        =   139
            TabStop         =   0   'False
            Top             =   30
            Width           =   3525
            Begin VB.CheckBox chkType 
               BackColor       =   &H80000005&
               Caption         =   "����"
               Height          =   210
               Index           =   8
               Left            =   2625
               TabIndex        =   144
               Top             =   15
               Value           =   1  'Checked
               Width           =   690
            End
            Begin VB.CheckBox chkType 
               BackColor       =   &H80000005&
               Caption         =   "����"
               Height          =   210
               Index           =   4
               Left            =   1890
               TabIndex        =   143
               Top             =   15
               Value           =   1  'Checked
               Width           =   690
            End
            Begin VB.CheckBox chkType 
               BackColor       =   &H80000005&
               Caption         =   "����"
               Height          =   225
               Index           =   2
               Left            =   240
               TabIndex        =   141
               Top             =   15
               Value           =   1  'Checked
               Width           =   705
            End
            Begin VB.CheckBox chkType 
               BackColor       =   &H80000005&
               Caption         =   "����"
               Height          =   210
               Index           =   3
               Left            =   1035
               TabIndex        =   142
               Top             =   15
               Value           =   1  'Checked
               Width           =   750
            End
            Begin VB.Label lblN 
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   "(                                   )"
               Height          =   180
               Index           =   57
               Left            =   75
               TabIndex        =   140
               Top             =   30
               Width           =   3330
            End
         End
         Begin VB.PictureBox picOpt 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   6
            Left            =   6630
            ScaleHeight     =   255
            ScaleWidth      =   1500
            TabIndex        =   136
            TabStop         =   0   'False
            Top             =   45
            Width           =   1500
            Begin VB.OptionButton optType 
               BackColor       =   &H80000005&
               Caption         =   "����"
               Height          =   180
               Index           =   18
               Left            =   840
               TabIndex        =   138
               Top             =   0
               Value           =   -1  'True
               Width           =   660
            End
            Begin VB.OptionButton optType 
               BackColor       =   &H80000005&
               Caption         =   "������"
               Height          =   180
               Index           =   15
               Left            =   0
               TabIndex        =   137
               Top             =   0
               Width           =   840
            End
         End
         Begin VB.TextBox txtDept 
            Height          =   300
            IMEMode         =   1  'ON
            Index           =   6
            Left            =   825
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   145
            Text            =   "���п���"
            ToolTipText     =   "���п���"
            Top             =   405
            Width           =   4800
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "ͳ��(&T)"
            Height          =   300
            Index           =   7
            Left            =   9705
            TabIndex        =   150
            Top             =   405
            Width           =   960
         End
         Begin VB.PictureBox picOpt 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   8280
            ScaleHeight     =   255
            ScaleWidth      =   1380
            TabIndex        =   128
            TabStop         =   0   'False
            Top             =   465
            Width           =   1380
            Begin VB.OptionButton optType 
               BackColor       =   &H80000005&
               Caption         =   "ƽ��"
               Height          =   180
               Index           =   14
               Left            =   0
               TabIndex        =   147
               Top             =   0
               Value           =   -1  'True
               Width           =   660
            End
            Begin VB.OptionButton optType 
               BackColor       =   &H80000005&
               Caption         =   "���"
               Height          =   180
               Index           =   13
               Left            =   690
               TabIndex        =   149
               Top             =   0
               Width           =   660
            End
         End
         Begin VB.ComboBox cboTimCount 
            Height          =   300
            Index           =   3
            Left            =   825
            Style           =   2  'Dropdown List
            TabIndex        =   126
            Top             =   0
            Width           =   1365
         End
         Begin VB.TextBox txtNum 
            Height          =   300
            Index           =   1
            Left            =   6660
            MaxLength       =   4
            TabIndex        =   146
            Text            =   "10"
            Top             =   405
            Width           =   465
         End
         Begin MSComCtl2.DTPicker dtpCountE 
            Height          =   300
            Index           =   3
            Left            =   4245
            TabIndex        =   129
            Top             =   0
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   124059651
            CurrentDate     =   41774
         End
         Begin MSComCtl2.DTPicker dtpCountS 
            Height          =   300
            Index           =   3
            Left            =   2565
            TabIndex        =   127
            Top             =   0
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   124059651
            CurrentDate     =   41774
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "��������"
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
            Index           =   56
            Left            =   7440
            TabIndex        =   135
            Top             =   465
            Width           =   795
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "ͳ�ƿ���"
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
            Index           =   55
            Left            =   0
            TabIndex        =   134
            Top             =   465
            Width           =   795
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "��                 ��"
            Height          =   180
            Index           =   54
            Left            =   2310
            TabIndex        =   133
            Top             =   60
            Width           =   2970
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "ͳ��ʱ��"
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
            Index           =   53
            Left            =   0
            TabIndex        =   132
            Top             =   60
            Width           =   795
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "�п�����"
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
            Index           =   52
            Left            =   5805
            TabIndex        =   131
            Top             =   60
            Width           =   780
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "��������"
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
            Index           =   51
            Left            =   5805
            TabIndex        =   130
            Top             =   450
            Width           =   795
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsInDruUse 
         Height          =   690
         Left            =   150
         TabIndex        =   148
         Top             =   1395
         Width           =   6975
         _cx             =   12303
         _cy             =   1217
         Appearance      =   2
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
         BackColorSel    =   16444122
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmKssStatistics.frx":6926
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   150
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
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
         OwnerDraw       =   1
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VSFlex8Ctl.VSFlexGrid vsInDruAna 
         Height          =   690
         Left            =   195
         TabIndex        =   151
         Top             =   2460
         Width           =   6975
         _cx             =   12303
         _cy             =   1217
         Appearance      =   2
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
         BackColorSel    =   16444122
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmKssStatistics.frx":6A51
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
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
         OwnerDraw       =   1
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "XXX����Ժ���˿���ҩ��ʹ��ͳ�Ʒ�����"
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
         Index           =   59
         Left            =   210
         TabIndex        =   153
         Top             =   2145
         Width           =   4485
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "˵��"
         Height          =   180
         Index           =   58
         Left            =   240
         TabIndex        =   152
         Top             =   3960
         Width           =   360
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "סԺҽ��������ҩͳ��"
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
         Left            =   90
         TabIndex        =   49
         Top             =   975
         Width           =   2550
      End
   End
   Begin VB.PictureBox picOtherSub 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5670
      Index           =   2
      Left            =   7320
      ScaleHeight     =   5670
      ScaleWidth      =   990
      TabIndex        =   46
      Top             =   690
      Width           =   990
      Begin VB.PictureBox picFilter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   780
         Index           =   2
         Left            =   0
         ScaleHeight     =   780
         ScaleWidth      =   11025
         TabIndex        =   98
         Top             =   0
         Width           =   11025
         Begin VB.CommandButton cmdDept 
            Caption         =   "��"
            Height          =   255
            Index           =   5
            Left            =   5325
            TabIndex        =   99
            TabStop         =   0   'False
            ToolTipText     =   "Ctrl+D"
            Top             =   420
            Width           =   285
         End
         Begin VB.CheckBox chkType 
            BackColor       =   &H80000005&
            Caption         =   "����"
            Height          =   210
            Index           =   1
            Left            =   7395
            TabIndex        =   106
            Top             =   45
            Width           =   690
         End
         Begin VB.CheckBox chkType 
            BackColor       =   &H80000005&
            Caption         =   "����"
            Height          =   225
            Index           =   0
            Left            =   6660
            TabIndex        =   105
            Top             =   45
            Width           =   690
         End
         Begin VB.TextBox txtNum 
            Height          =   300
            Index           =   0
            Left            =   6660
            MaxLength       =   4
            TabIndex        =   108
            Text            =   "10"
            Top             =   405
            Width           =   465
         End
         Begin VB.ComboBox cboTimCount 
            Height          =   300
            Index           =   2
            Left            =   825
            Style           =   2  'Dropdown List
            TabIndex        =   102
            Top             =   0
            Width           =   1365
         End
         Begin VB.PictureBox picOpt 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   8280
            ScaleHeight     =   255
            ScaleWidth      =   1380
            TabIndex        =   100
            TabStop         =   0   'False
            Top             =   465
            Width           =   1380
            Begin VB.OptionButton optType 
               BackColor       =   &H80000005&
               Caption         =   "���"
               Height          =   180
               Index           =   10
               Left            =   690
               TabIndex        =   101
               Top             =   0
               Width           =   660
            End
            Begin VB.OptionButton optType 
               BackColor       =   &H80000005&
               Caption         =   "ƽ��"
               Height          =   180
               Index           =   6
               Left            =   0
               TabIndex        =   109
               Top             =   0
               Value           =   -1  'True
               Width           =   660
            End
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "ͳ��(&T)"
            Height          =   300
            Index           =   6
            Left            =   9705
            TabIndex        =   112
            Top             =   405
            Width           =   960
         End
         Begin VB.TextBox txtDept 
            Height          =   300
            IMEMode         =   1  'ON
            Index           =   5
            Left            =   825
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   107
            Text            =   "���п���"
            ToolTipText     =   "���п���"
            Top             =   405
            Width           =   4800
         End
         Begin MSComCtl2.DTPicker dtpCountE 
            Height          =   300
            Index           =   2
            Left            =   4245
            TabIndex        =   104
            Top             =   0
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   124059651
            CurrentDate     =   41774
         End
         Begin MSComCtl2.DTPicker dtpCountS 
            Height          =   300
            Index           =   2
            Left            =   2565
            TabIndex        =   103
            Top             =   0
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   124059651
            CurrentDate     =   41774
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "��������"
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
            Index           =   38
            Left            =   5805
            TabIndex        =   116
            Top             =   450
            Width           =   795
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "����"
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
            Index           =   37
            Left            =   6195
            TabIndex        =   115
            Top             =   60
            Width           =   405
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "ͳ��ʱ��"
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
            Index           =   34
            Left            =   0
            TabIndex        =   114
            Top             =   60
            Width           =   795
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "��                 ��"
            Height          =   180
            Index           =   33
            Left            =   2310
            TabIndex        =   113
            Top             =   60
            Width           =   2970
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "ͳ�ƿ���"
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
            Index           =   32
            Left            =   0
            TabIndex        =   111
            Top             =   465
            Width           =   795
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "��������"
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
            Index           =   3
            Left            =   7440
            TabIndex        =   110
            Top             =   465
            Width           =   795
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsCountDruUse 
         Height          =   690
         Left            =   60
         TabIndex        =   122
         Top             =   1620
         Width           =   6975
         _cx             =   12303
         _cy             =   1217
         Appearance      =   2
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
         BackColorSel    =   16444122
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmKssStatistics.frx":6A8F
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
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
         OwnerDraw       =   1
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VSFlex8Ctl.VSFlexGrid vsCountCF 
         Height          =   690
         Left            =   45
         TabIndex        =   123
         Top             =   3180
         Width           =   6975
         _cx             =   12303
         _cy             =   1217
         Appearance      =   2
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
         BackColorSel    =   16444122
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmKssStatistics.frx":6BBA
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
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
         OwnerDraw       =   1
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "˵��"
         Height          =   180
         Index           =   50
         Left            =   90
         TabIndex        =   200
         Top             =   3960
         Width           =   360
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "����������XXXX��"
         Height          =   180
         Index           =   49
         Left            =   0
         TabIndex        =   121
         Top             =   2910
         Width           =   1440
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "����������XXXX��"
         Height          =   180
         Index           =   48
         Left            =   480
         TabIndex        =   120
         Top             =   1410
         Width           =   1440
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "���ڣ�2014-12-11��2014-12-31"
         Height          =   180
         Index           =   47
         Left            =   45
         TabIndex        =   119
         Top             =   2655
         Width           =   2520
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "���ڣ�2014-12-11��2014-12-31"
         Height          =   180
         Index           =   46
         Left            =   105
         TabIndex        =   118
         Top             =   1185
         Width           =   2520
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "XXX�Ŵ���ͳ�Ʒ�����"
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
         Index           =   39
         Left            =   45
         TabIndex        =   117
         Top             =   2385
         Width           =   2445
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "��(��)�ﴦ��������ҩͳ��"
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
         Left            =   105
         TabIndex        =   47
         Top             =   885
         Width           =   3075
      End
   End
   Begin VB.PictureBox picOtherSub 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6570
      Index           =   1
      Left            =   6285
      ScaleHeight     =   6570
      ScaleWidth      =   705
      TabIndex        =   45
      Top             =   600
      Width           =   705
      Begin VB.PictureBox picFilter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   795
         Index           =   1
         Left            =   45
         ScaleHeight     =   795
         ScaleWidth      =   9120
         TabIndex        =   82
         Top             =   90
         Width           =   9120
         Begin VB.CommandButton cmdDept 
            Caption         =   "��"
            Height          =   255
            Index           =   4
            Left            =   5325
            TabIndex        =   83
            TabStop         =   0   'False
            ToolTipText     =   "Ctrl+D"
            Top             =   420
            Width           =   285
         End
         Begin VB.TextBox txtDept 
            Height          =   300
            IMEMode         =   1  'ON
            Index           =   4
            Left            =   825
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   87
            Text            =   "���п���"
            ToolTipText     =   "���п���"
            Top             =   405
            Width           =   4800
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "ͳ��(&T)"
            Height          =   300
            Index           =   5
            Left            =   8040
            TabIndex        =   89
            Top             =   405
            Width           =   960
         End
         Begin VB.PictureBox picOpt 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   6615
            ScaleHeight     =   255
            ScaleWidth      =   1380
            TabIndex        =   90
            TabStop         =   0   'False
            Top             =   465
            Width           =   1380
            Begin VB.OptionButton optType 
               BackColor       =   &H80000005&
               Caption         =   "����"
               Height          =   180
               Index           =   17
               Left            =   0
               TabIndex        =   88
               Top             =   0
               Value           =   -1  'True
               Width           =   660
            End
            Begin VB.OptionButton optType 
               BackColor       =   &H80000005&
               Caption         =   "ҽ��"
               Height          =   180
               Index           =   16
               Left            =   690
               TabIndex        =   91
               Top             =   0
               Width           =   660
            End
         End
         Begin VB.ComboBox cboTimCount 
            Height          =   300
            Index           =   1
            Left            =   825
            Style           =   2  'Dropdown List
            TabIndex        =   84
            Top             =   0
            Width           =   1365
         End
         Begin MSComCtl2.DTPicker dtpCountE 
            Height          =   300
            Index           =   1
            Left            =   4245
            TabIndex        =   86
            Top             =   0
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   124059651
            CurrentDate     =   41774
         End
         Begin MSComCtl2.DTPicker dtpCountS 
            Height          =   300
            Index           =   1
            Left            =   2565
            TabIndex        =   85
            Top             =   0
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   124059651
            CurrentDate     =   41774
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "���ܷ�ʽ"
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
            Index           =   36
            Left            =   5775
            TabIndex        =   95
            Top             =   465
            Width           =   795
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "ͳ�ƿ���"
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
            Index           =   35
            Left            =   0
            TabIndex        =   94
            Top             =   465
            Width           =   795
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "��                 ��"
            Height          =   180
            Index           =   28
            Left            =   2310
            TabIndex        =   93
            Top             =   60
            Width           =   2970
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "ͳ��ʱ��"
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
            Index           =   2
            Left            =   0
            TabIndex        =   92
            Top             =   60
            Width           =   795
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsCut 
         Height          =   915
         Left            =   210
         TabIndex        =   96
         Top             =   1890
         Width           =   6975
         _cx             =   12303
         _cy             =   1614
         Appearance      =   2
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
         BackColorSel    =   16444122
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   0
         HighLight       =   1
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
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmKssStatistics.frx":6BF8
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   110
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
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
         OwnerDraw       =   1
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "�����п�Χ����Ԥ����ҩͳ��"
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
         Index           =   76
         Left            =   450
         TabIndex        =   245
         Top             =   1095
         Width           =   3315
      End
      Begin VB.Label lblCut 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "˵��"
         Height          =   180
         Left            =   330
         TabIndex        =   97
         Top             =   3120
         Width           =   360
      End
   End
   Begin VB.PictureBox picOtherSub 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5025
      Index           =   0
      Left            =   5325
      ScaleHeight     =   5025
      ScaleWidth      =   720
      TabIndex        =   54
      Top             =   660
      Width           =   720
      Begin VB.PictureBox picFilter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1155
         Index           =   0
         Left            =   135
         ScaleHeight     =   1155
         ScaleWidth      =   10575
         TabIndex        =   57
         Top             =   60
         Width           =   10575
         Begin VB.CommandButton cmdDept 
            Caption         =   "��"
            Height          =   255
            Index           =   0
            Left            =   5325
            TabIndex        =   58
            TabStop         =   0   'False
            ToolTipText     =   "Ctrl+D"
            Top             =   780
            Width           =   285
         End
         Begin VB.ComboBox cboTimCount 
            Height          =   300
            Index           =   0
            Left            =   825
            Style           =   2  'Dropdown List
            TabIndex        =   66
            Top             =   0
            Width           =   1365
         End
         Begin VB.PictureBox picOpt 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   0
            Left            =   855
            ScaleHeight     =   240
            ScaleWidth      =   1455
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   450
            Width           =   1455
            Begin VB.OptionButton optType 
               BackColor       =   &H80000005&
               Caption         =   "����"
               Height          =   180
               Index           =   4
               Left            =   0
               TabIndex        =   65
               Top             =   0
               Width           =   660
            End
            Begin VB.OptionButton optType 
               BackColor       =   &H80000005&
               Caption         =   "סԺ"
               Height          =   180
               Index           =   5
               Left            =   720
               TabIndex        =   70
               Top             =   0
               Value           =   -1  'True
               Width           =   660
            End
         End
         Begin VB.PictureBox picOpt 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   3465
            ScaleHeight     =   255
            ScaleWidth      =   2115
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   465
            Width           =   2115
            Begin VB.OptionButton optType 
               BackColor       =   &H80000005&
               Caption         =   "ҩƷ"
               Height          =   180
               Index           =   7
               Left            =   1440
               TabIndex        =   72
               Top             =   0
               Value           =   -1  'True
               Width           =   660
            End
            Begin VB.OptionButton optType 
               BackColor       =   &H80000005&
               Caption         =   "ҽ��"
               Height          =   180
               Index           =   8
               Left            =   705
               TabIndex        =   63
               Top             =   0
               Width           =   660
            End
            Begin VB.OptionButton optType 
               BackColor       =   &H80000005&
               Caption         =   "����"
               Height          =   180
               Index           =   9
               Left            =   0
               TabIndex        =   62
               Top             =   0
               Width           =   660
            End
         End
         Begin VB.PictureBox picOpt 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   6615
            ScaleHeight     =   255
            ScaleWidth      =   1380
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   810
            Width           =   1380
            Begin VB.OptionButton optType 
               BackColor       =   &H80000005&
               Caption         =   "���"
               Height          =   180
               Index           =   11
               Left            =   690
               TabIndex        =   60
               Top             =   0
               Width           =   660
            End
            Begin VB.OptionButton optType 
               BackColor       =   &H80000005&
               Caption         =   "����"
               Height          =   180
               Index           =   12
               Left            =   0
               TabIndex        =   76
               Top             =   0
               Value           =   -1  'True
               Width           =   660
            End
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "ͳ��(&T)"
            Height          =   300
            Index           =   4
            Left            =   9555
            TabIndex        =   80
            Top             =   750
            Width           =   960
         End
         Begin VB.TextBox txtDept 
            Height          =   300
            IMEMode         =   1  'ON
            Index           =   0
            Left            =   825
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   74
            Text            =   "���п���"
            ToolTipText     =   "���п���"
            Top             =   750
            Width           =   4800
         End
         Begin VB.TextBox txtTopRan 
            Height          =   300
            Left            =   8745
            MaxLength       =   3
            TabIndex        =   78
            Text            =   "10"
            Top             =   750
            Width           =   465
         End
         Begin MSComCtl2.DTPicker dtpCountE 
            Height          =   300
            Index           =   0
            Left            =   4245
            TabIndex        =   68
            Top             =   0
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   124059651
            CurrentDate     =   41774
         End
         Begin MSComCtl2.DTPicker dtpCountS 
            Height          =   300
            Index           =   0
            Left            =   2565
            TabIndex        =   67
            Top             =   0
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   124059651
            CurrentDate     =   41774
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "ͳ��ʱ��"
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
            Index           =   8
            Left            =   0
            TabIndex        =   81
            Top             =   60
            Width           =   795
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "��                 ��"
            Height          =   180
            Index           =   9
            Left            =   2310
            TabIndex        =   79
            Top             =   60
            Width           =   2970
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "ͳ�Ƴ���"
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
            Index           =   16
            Left            =   0
            TabIndex        =   77
            Top             =   450
            Width           =   795
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "���ܷ�ʽ"
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
            Index           =   17
            Left            =   2565
            TabIndex        =   75
            Top             =   465
            Width           =   795
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "ͳ��ǰ     ��"
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
            Index           =   31
            Left            =   8130
            TabIndex        =   73
            Top             =   810
            Width           =   1380
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "ͳ�ƿ���"
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
            Index           =   18
            Left            =   0
            TabIndex        =   71
            Top             =   810
            Width           =   795
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "����ʽ"
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
            Index           =   19
            Left            =   5775
            TabIndex        =   69
            Top             =   810
            Width           =   795
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsUseRan 
         Height          =   915
         Left            =   570
         TabIndex        =   55
         Top             =   1755
         Width           =   6975
         _cx             =   12303
         _cy             =   1614
         Appearance      =   2
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
         BackColorSel    =   16444122
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmKssStatistics.frx":6CBD
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
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
         OwnerDraw       =   1
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "����ҩ��ʹ���������ͳ��"
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
         Index           =   75
         Left            =   180
         TabIndex        =   244
         Top             =   1290
         Width           =   3060
      End
      Begin VB.Label lblUse 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "˵��"
         Height          =   180
         Left            =   120
         TabIndex        =   56
         Top             =   1875
         Width           =   360
      End
   End
   Begin VB.PictureBox picReportSub 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4410
      Index           =   3
      Left            =   3735
      ScaleHeight     =   4410
      ScaleWidth      =   1125
      TabIndex        =   205
      Top             =   1455
      Width           =   1125
      Begin VB.PictureBox picFilter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   930
         Index           =   9
         Left            =   0
         ScaleHeight     =   930
         ScaleWidth      =   10125
         TabIndex        =   232
         Top             =   105
         Width           =   10125
         Begin VB.CommandButton cmdDept 
            Caption         =   "��"
            Height          =   265
            Index           =   3
            Left            =   5325
            TabIndex        =   233
            TabStop         =   0   'False
            ToolTipText     =   "Ctrl+D"
            Top             =   420
            Width           =   285
         End
         Begin VB.ComboBox cboTimRP 
            Height          =   300
            Index           =   3
            Left            =   825
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   0
            Width           =   1365
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "ͳ��(&T)"
            Height          =   300
            Index           =   3
            Left            =   5745
            TabIndex        =   31
            Top             =   405
            Width           =   960
         End
         Begin VB.TextBox txtDept 
            Height          =   300
            IMEMode         =   1  'ON
            Index           =   3
            Left            =   825
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   30
            Text            =   "���п���"
            ToolTipText     =   "���п���"
            Top             =   400
            Width           =   4800
         End
         Begin MSComCtl2.DTPicker dtpRPE 
            Height          =   300
            Index           =   3
            Left            =   4245
            TabIndex        =   29
            Top             =   0
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   124059651
            CurrentDate     =   41774
         End
         Begin MSComCtl2.DTPicker dtpRPS 
            Height          =   300
            Index           =   3
            Left            =   2565
            TabIndex        =   28
            Top             =   0
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   124059651
            CurrentDate     =   41774
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "��                 ��"
            Height          =   180
            Index           =   41
            Left            =   2310
            TabIndex        =   236
            Top             =   60
            Width           =   2970
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "ͳ��ʱ��"
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
            Index           =   40
            Left            =   0
            TabIndex        =   235
            Top             =   60
            Width           =   795
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "��������"
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
            Index           =   42
            Left            =   0
            TabIndex        =   234
            Top             =   450
            Width           =   795
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsZYYY 
         Height          =   915
         Left            =   30
         TabIndex        =   32
         Top             =   2520
         Width           =   5595
         _cx             =   9869
         _cy             =   1614
         Appearance      =   2
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
         BackColorSel    =   16444122
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
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
         OwnerDraw       =   1
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "˵����"
         Height          =   180
         Index           =   44
         Left            =   120
         TabIndex        =   231
         Top             =   3585
         Width           =   540
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "���λ�����������XXXX��"
         Height          =   180
         Index           =   45
         Left            =   30
         TabIndex        =   230
         Top             =   2280
         Width           =   1980
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "סԺ���˿�����ҩ�����"
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
         Index           =   43
         Left            =   30
         TabIndex        =   229
         Top             =   1935
         Width           =   2805
      End
   End
   Begin VB.PictureBox picReportSub 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5475
      Index           =   2
      Left            =   2355
      ScaleHeight     =   5475
      ScaleWidth      =   795
      TabIndex        =   204
      Top             =   1200
      Width           =   795
      Begin VB.PictureBox picFilter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1005
         Index           =   8
         Left            =   0
         ScaleHeight     =   1005
         ScaleWidth      =   10845
         TabIndex        =   222
         Top             =   0
         Width           =   10845
         Begin VB.CommandButton cmdDept 
            Caption         =   "��"
            Height          =   265
            Index           =   2
            Left            =   5325
            TabIndex        =   223
            TabStop         =   0   'False
            ToolTipText     =   "Ctrl+D"
            Top             =   420
            Width           =   285
         End
         Begin VB.TextBox txtDept 
            Height          =   300
            IMEMode         =   1  'ON
            Index           =   2
            Left            =   825
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   20
            Text            =   "���п���"
            ToolTipText     =   "���п���"
            Top             =   400
            Width           =   4800
         End
         Begin VB.TextBox txtCount 
            Height          =   300
            IMEMode         =   1  'ON
            Index           =   2
            Left            =   6690
            MaxLength       =   3
            TabIndex        =   21
            Text            =   "100"
            Top             =   405
            Width           =   420
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "ͳ��(&T)"
            Height          =   300
            Index           =   2
            Left            =   9570
            TabIndex        =   24
            Top             =   405
            Width           =   960
         End
         Begin VB.OptionButton optType 
            BackColor       =   &H80000005&
            Caption         =   "ƽ��"
            Height          =   180
            Index           =   3
            Left            =   8100
            TabIndex        =   22
            Top             =   450
            Width           =   660
         End
         Begin VB.OptionButton optType 
            BackColor       =   &H80000005&
            Caption         =   "���"
            Height          =   180
            Index           =   2
            Left            =   8775
            TabIndex        =   23
            Top             =   450
            Value           =   -1  'True
            Width           =   660
         End
         Begin VB.ComboBox cboTimRP 
            Height          =   300
            Index           =   2
            Left            =   825
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   0
            Width           =   1365
         End
         Begin MSComCtl2.DTPicker dtpRPE 
            Height          =   300
            Index           =   2
            Left            =   4245
            TabIndex        =   19
            Top             =   0
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   124059651
            CurrentDate     =   41774
         End
         Begin MSComCtl2.DTPicker dtpRPS 
            Height          =   300
            Index           =   2
            Left            =   2565
            TabIndex        =   18
            Top             =   0
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   124059651
            CurrentDate     =   41774
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "��������"
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
            Index           =   22
            Left            =   0
            TabIndex        =   228
            Top             =   450
            Width           =   795
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "��������"
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
            Index           =   23
            Left            =   5835
            TabIndex        =   227
            Top             =   450
            Width           =   795
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "��������"
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
            Index           =   24
            Left            =   7290
            TabIndex        =   226
            Top             =   450
            Width           =   795
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "ͳ��ʱ��"
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
            Index           =   20
            Left            =   0
            TabIndex        =   225
            Top             =   60
            Width           =   795
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "��                 ��"
            Height          =   180
            Index           =   21
            Left            =   2310
            TabIndex        =   224
            Top             =   60
            Width           =   2970
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsCF 
         Height          =   915
         Left            =   330
         TabIndex        =   26
         Top             =   3645
         Width           =   5595
         _cx             =   9869
         _cy             =   1614
         Appearance      =   2
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
         BackColorSel    =   16444122
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   2
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmKssStatistics.frx":6DD6
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VSFlex8Ctl.VSFlexGrid vsMZYY 
         Height          =   825
         Left            =   75
         TabIndex        =   25
         Top             =   1875
         Width           =   5595
         _cx             =   9869
         _cy             =   1455
         Appearance      =   2
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
         BackColorSel    =   16444122
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmKssStatistics.frx":6E12
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
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
         OwnerDraw       =   1
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "���ﲡ����ҩ��������"
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
         Index           =   25
         Left            =   60
         TabIndex        =   221
         Top             =   1095
         Width           =   2805
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "���ڣ�XXX��XXX"
         Height          =   180
         Index           =   26
         Left            =   15
         TabIndex        =   220
         Top             =   1365
         Width           =   1260
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "����������XXXX��"
         Height          =   180
         Index           =   27
         Left            =   60
         TabIndex        =   219
         Top             =   1575
         Width           =   1440
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "����������XXXX��"
         Height          =   180
         Index           =   30
         Left            =   90
         TabIndex        =   218
         Top             =   3345
         Width           =   1440
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "���ڣ�XXX��XXX"
         Height          =   180
         Index           =   29
         Left            =   -30
         TabIndex        =   217
         Top             =   3120
         Width           =   1260
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "XXX�Ŵ���ͳ�Ʒ�����"
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
         Index           =   70
         Left            =   30
         TabIndex        =   216
         Top             =   2850
         Width           =   2445
      End
      Begin VB.Label lblCFSM 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "˵��"
         Height          =   180
         Left            =   315
         TabIndex        =   215
         Top             =   4710
         Width           =   360
      End
   End
   Begin VB.PictureBox picReportSub 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5145
      Index           =   1
      Left            =   1290
      ScaleHeight     =   5145
      ScaleWidth      =   810
      TabIndex        =   203
      Top             =   1125
      Width           =   810
      Begin XtremeReportControl.ReportControl rptPati 
         Height          =   525
         Left            =   750
         TabIndex        =   16
         Top             =   1590
         Width           =   8340
         _Version        =   589884
         _ExtentX        =   14711
         _ExtentY        =   926
         _StockProps     =   0
      End
      Begin VB.PictureBox picFilter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1245
         Index           =   7
         Left            =   0
         ScaleHeight     =   1245
         ScaleWidth      =   10650
         TabIndex        =   206
         Top             =   0
         Width           =   10650
         Begin VB.CommandButton cmdDept 
            Caption         =   "��"
            Height          =   265
            Index           =   1
            Left            =   5325
            TabIndex        =   207
            TabStop         =   0   'False
            ToolTipText     =   "Ctrl+D"
            Top             =   420
            Width           =   285
         End
         Begin VB.ComboBox cboTimRP 
            Height          =   300
            Index           =   1
            Left            =   825
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   0
            Width           =   1365
         End
         Begin VB.CommandButton cmdCYEdit 
            Caption         =   "�༭�����(&E)"
            Height          =   300
            Left            =   7410
            TabIndex        =   15
            Top             =   765
            Width           =   1550
         End
         Begin VB.CommandButton cmdCYDel 
            Caption         =   "ɾ��������¼(&D)"
            Height          =   300
            Left            =   5835
            TabIndex        =   14
            Top             =   765
            Width           =   1550
         End
         Begin VB.OptionButton optType 
            BackColor       =   &H80000005&
            Caption         =   "���"
            Height          =   180
            Index           =   1
            Left            =   8775
            TabIndex        =   11
            Top             =   450
            Value           =   -1  'True
            Width           =   660
         End
         Begin VB.OptionButton optType 
            BackColor       =   &H80000005&
            Caption         =   "ƽ��"
            Height          =   180
            Index           =   0
            Left            =   8100
            TabIndex        =   10
            Top             =   450
            Width           =   660
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "����(&C)"
            Height          =   300
            Index           =   1
            Left            =   9570
            TabIndex        =   12
            Top             =   405
            Width           =   960
         End
         Begin VB.TextBox txtCount 
            Height          =   300
            IMEMode         =   1  'ON
            Index           =   1
            Left            =   6690
            MaxLength       =   4
            TabIndex        =   9
            Text            =   "15"
            Top             =   405
            Width           =   450
         End
         Begin VB.TextBox txtDept 
            Height          =   300
            IMEMode         =   1  'ON
            Index           =   1
            Left            =   825
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   8
            Text            =   "���п���"
            ToolTipText     =   "���п���"
            Top             =   400
            Width           =   4800
         End
         Begin VB.CommandButton cmdCYSel 
            Caption         =   "��"
            Height          =   265
            Left            =   5325
            TabIndex        =   208
            TabStop         =   0   'False
            ToolTipText     =   "Ctrl+R"
            Top             =   795
            Width           =   285
         End
         Begin VB.TextBox txtCYJL 
            Height          =   300
            IMEMode         =   1  'ON
            Left            =   825
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   13
            Top             =   765
            Width           =   4800
         End
         Begin MSComCtl2.DTPicker dtpRPE 
            Height          =   300
            Index           =   1
            Left            =   4245
            TabIndex        =   7
            Top             =   0
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   124059651
            CurrentDate     =   41774
         End
         Begin MSComCtl2.DTPicker dtpRPS 
            Height          =   300
            Index           =   1
            Left            =   2565
            TabIndex        =   6
            Top             =   0
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   124059651
            CurrentDate     =   41774
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "��                 ��"
            Height          =   180
            Index           =   11
            Left            =   2310
            TabIndex        =   214
            Top             =   60
            Width           =   2970
         End
         Begin VB.Label lblN 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "ͳ��ʱ��"
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
            Index           =   10
            Left            =   0
            TabIndex        =   213
            Top             =   60
            Width           =   780
         End
         Begin VB.Label lblN 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "������¼"
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
            Index           =   15
            Left            =   0
            TabIndex        =   212
            Top             =   825
            Width           =   780
         End
         Begin VB.Label lblN 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��������"
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
            Index           =   14
            Left            =   7290
            TabIndex        =   211
            Top             =   450
            Width           =   780
         End
         Begin VB.Label lblN 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��������"
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
            Index           =   13
            Left            =   5835
            TabIndex        =   210
            Top             =   450
            Width           =   780
         End
         Begin VB.Label lblN 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "��������"
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
            Index           =   12
            Left            =   0
            TabIndex        =   209
            Top             =   450
            Width           =   780
         End
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "(��)�������˿�����ҩ����������鼰���۱�"
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
         Index           =   77
         Left            =   405
         TabIndex        =   246
         Top             =   1290
         Width           =   5115
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "˵��"
         Height          =   180
         Index           =   72
         Left            =   300
         TabIndex        =   240
         Top             =   1980
         Width           =   360
      End
   End
   Begin VB.PictureBox picReportSub 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5160
      Index           =   0
      Left            =   45
      ScaleHeight     =   5160
      ScaleWidth      =   1065
      TabIndex        =   201
      Top             =   1995
      Width           =   1065
      Begin VB.PictureBox picFilter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   6
         Left            =   30
         ScaleHeight     =   495
         ScaleWidth      =   7275
         TabIndex        =   237
         Top             =   75
         Width           =   7275
         Begin VB.CommandButton cmdOK 
            Caption         =   "ͳ��(&T)"
            Height          =   300
            Index           =   0
            Left            =   5865
            TabIndex        =   3
            Top             =   0
            Width           =   960
         End
         Begin VB.ComboBox cboTimRP 
            Height          =   300
            Index           =   0
            Left            =   825
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   0
            Width           =   1365
         End
         Begin MSComCtl2.DTPicker dtpRPE 
            Height          =   300
            Index           =   0
            Left            =   4245
            TabIndex        =   2
            Top             =   0
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   124059651
            CurrentDate     =   41774
         End
         Begin MSComCtl2.DTPicker dtpRPS 
            Height          =   300
            Index           =   0
            Left            =   2565
            TabIndex        =   1
            Top             =   0
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   124059651
            CurrentDate     =   41774
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "��                 ��"
            Height          =   180
            Index           =   1
            Left            =   2310
            TabIndex        =   239
            Top             =   60
            Width           =   2970
         End
         Begin VB.Label lblN 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "ͳ��ʱ��"
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
            Index           =   0
            Left            =   0
            TabIndex        =   238
            Top             =   60
            Width           =   780
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsBill 
         Height          =   525
         Left            =   1185
         TabIndex        =   4
         Top             =   1245
         Width           =   3855
         _cx             =   6800
         _cy             =   926
         Appearance      =   2
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
         BackColorSel    =   16444122
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmKssStatistics.frx":6EB0
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
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
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "����ҩƷ���Ľ������"
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
         Index           =   74
         Left            =   420
         TabIndex        =   243
         Top             =   705
         Width           =   2805
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "˵��"
         Height          =   180
         Left            =   435
         TabIndex        =   202
         Top             =   1515
         Width           =   360
      End
   End
   Begin VB.PictureBox picOtherSub 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4050
      Index           =   4
      Left            =   9630
      ScaleHeight     =   4050
      ScaleWidth      =   900
      TabIndex        =   50
      Top             =   780
      Width           =   900
      Begin VB.PictureBox picFilter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   780
         Index           =   4
         Left            =   0
         ScaleHeight     =   780
         ScaleWidth      =   11295
         TabIndex        =   154
         Top             =   0
         Width           =   11295
         Begin VB.PictureBox picOpt 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   8
            Left            =   6645
            ScaleHeight     =   255
            ScaleWidth      =   2835
            TabIndex        =   156
            TabStop         =   0   'False
            Top             =   30
            Width           =   2835
            Begin VB.CheckBox chkType 
               BackColor       =   &H80000005&
               Caption         =   "����"
               Height          =   210
               Index           =   9
               Left            =   2175
               TabIndex        =   199
               Top             =   30
               Value           =   1  'Checked
               Width           =   690
            End
            Begin VB.CheckBox chkType 
               BackColor       =   &H80000005&
               Caption         =   "����"
               Height          =   210
               Index           =   5
               Left            =   15
               TabIndex        =   166
               Top             =   30
               Value           =   1  'Checked
               Width           =   690
            End
            Begin VB.CheckBox chkType 
               BackColor       =   &H80000005&
               Caption         =   "����"
               Height          =   225
               Index           =   6
               Left            =   705
               TabIndex        =   157
               Top             =   15
               Value           =   1  'Checked
               Width           =   705
            End
            Begin VB.CheckBox chkType 
               BackColor       =   &H80000005&
               Caption         =   "����"
               Height          =   210
               Index           =   7
               Left            =   1470
               TabIndex        =   158
               Top             =   30
               Value           =   1  'Checked
               Width           =   750
            End
         End
         Begin VB.CommandButton cmdDept 
            Caption         =   "��"
            Height          =   255
            Index           =   7
            Left            =   5325
            TabIndex        =   155
            TabStop         =   0   'False
            ToolTipText     =   "Ctrl+D"
            Top             =   420
            Width           =   285
         End
         Begin VB.TextBox txtNum 
            Height          =   300
            Index           =   2
            Left            =   6660
            MaxLength       =   3
            TabIndex        =   174
            Text            =   "10"
            Top             =   405
            Width           =   465
         End
         Begin VB.ComboBox cboTimCount 
            Height          =   300
            Index           =   4
            Left            =   825
            Style           =   2  'Dropdown List
            TabIndex        =   161
            Top             =   0
            Width           =   1365
         End
         Begin VB.PictureBox picOpt 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   10
            Left            =   8280
            ScaleHeight     =   255
            ScaleWidth      =   1380
            TabIndex        =   159
            TabStop         =   0   'False
            Top             =   465
            Width           =   1380
            Begin VB.OptionButton optType 
               BackColor       =   &H80000005&
               Caption         =   "���"
               Height          =   180
               Index           =   22
               Left            =   690
               TabIndex        =   160
               Top             =   0
               Width           =   660
            End
            Begin VB.OptionButton optType 
               BackColor       =   &H80000005&
               Caption         =   "ƽ��"
               Height          =   180
               Index           =   21
               Left            =   0
               TabIndex        =   176
               Top             =   0
               Value           =   -1  'True
               Width           =   660
            End
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "ͳ��(&T)"
            Height          =   300
            Index           =   8
            Left            =   9705
            TabIndex        =   180
            Top             =   405
            Width           =   960
         End
         Begin VB.TextBox txtDept 
            Height          =   300
            IMEMode         =   1  'ON
            Index           =   7
            Left            =   825
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   172
            Text            =   "���п���"
            ToolTipText     =   "���п���"
            Top             =   405
            Width           =   4800
         End
         Begin MSComCtl2.DTPicker dtpCountE 
            Height          =   300
            Index           =   4
            Left            =   4245
            TabIndex        =   164
            Top             =   0
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   124059651
            CurrentDate     =   41774
         End
         Begin MSComCtl2.DTPicker dtpCountS 
            Height          =   300
            Index           =   4
            Left            =   2565
            TabIndex        =   162
            Top             =   0
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   124059651
            CurrentDate     =   41774
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "��������"
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
            Index           =   66
            Left            =   5805
            TabIndex        =   170
            Top             =   450
            Width           =   795
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "�п�����"
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
            Index           =   65
            Left            =   5805
            TabIndex        =   169
            Top             =   60
            Width           =   780
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "ͳ��ʱ��"
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
            Index           =   64
            Left            =   0
            TabIndex        =   168
            Top             =   60
            Width           =   795
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "��                 ��"
            Height          =   180
            Index           =   63
            Left            =   2310
            TabIndex        =   167
            Top             =   60
            Width           =   2970
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "ͳ�ƿ���"
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
            Index           =   62
            Left            =   0
            TabIndex        =   165
            Top             =   465
            Width           =   795
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "��������"
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
            Index           =   61
            Left            =   7440
            TabIndex        =   163
            Top             =   465
            Width           =   795
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsOpeKssUse 
         Height          =   690
         Left            =   60
         TabIndex        =   171
         Top             =   1260
         Width           =   7710
         _cx             =   13600
         _cy             =   1217
         Appearance      =   2
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
         BackColorSel    =   16444122
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmKssStatistics.frx":6F2B
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
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
         OwnerDraw       =   1
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "˵��"
         Height          =   180
         Index           =   73
         Left            =   180
         TabIndex        =   241
         Top             =   2610
         Width           =   360
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "���󿹾�ҩ��ʹ�ó�N��ͳ��"
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
         Left            =   75
         TabIndex        =   51
         Top             =   825
         Width           =   3195
      End
   End
   Begin VB.PictureBox picOther 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   5190
      ScaleHeight     =   855
      ScaleWidth      =   8865
      TabIndex        =   35
      Top             =   165
      Width           =   8865
      Begin XtremeSuiteControls.TabControl tbcOther 
         Height          =   960
         Left            =   225
         TabIndex        =   37
         Top             =   75
         Width           =   8490
         _Version        =   589884
         _ExtentX        =   14975
         _ExtentY        =   1693
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picDept 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2880
      Left            =   1890
      ScaleHeight     =   2850
      ScaleWidth      =   4890
      TabIndex        =   38
      Top             =   7380
      Visible         =   0   'False
      Width           =   4920
      Begin VB.TextBox txtFind 
         Height          =   270
         Left            =   50
         TabIndex        =   42
         Top             =   75
         Width           =   1575
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "����"
         Height          =   270
         Left            =   1740
         TabIndex        =   41
         Top             =   75
         Width           =   615
      End
      Begin VB.CommandButton cmdFindOk 
         Caption         =   "ȷ��"
         Height          =   270
         Left            =   3480
         TabIndex        =   40
         Top             =   75
         Width           =   615
      End
      Begin VB.CommandButton cmdFindCancle 
         Caption         =   "ȡ��"
         Height          =   270
         Left            =   4200
         TabIndex        =   39
         Top             =   75
         Width           =   615
      End
      Begin MSComctlLib.ListView lvwItems 
         Height          =   2280
         Left            =   75
         TabIndex        =   43
         ToolTipText     =   "ȫѡCtrl+A��ȫ��Ctrl+R"
         Top             =   510
         Width           =   4755
         _ExtentX        =   8387
         _ExtentY        =   4022
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "img16"
         SmallIcons      =   "img16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
   End
   Begin VB.PictureBox picReport 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   165
      ScaleHeight     =   810
      ScaleWidth      =   4935
      TabIndex        =   34
      Top             =   675
      Width           =   4935
      Begin XtremeSuiteControls.TabControl tbcReport 
         Height          =   2535
         Left            =   120
         TabIndex        =   36
         Top             =   75
         Width           =   6975
         _Version        =   589884
         _ExtentX        =   12303
         _ExtentY        =   4471
         _StockProps     =   64
      End
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   840
      Top             =   15
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKssStatistics.frx":702E
            Key             =   "Path"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKssStatistics.frx":75C8
            Key             =   "PatiMan"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKssStatistics.frx":7B62
            Key             =   "PatiWoman"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKssStatistics.frx":80FC
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKssStatistics.frx":8696
            Key             =   "UnCheck"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKssStatistics.frx":8C30
            Key             =   "������"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKssStatistics.frx":F492
            Key             =   "Dept"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKssStatistics.frx":15CF4
            Key             =   "printer"
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.TabControl tbcSub 
      Height          =   525
      Left            =   1770
      TabIndex        =   33
      Top             =   45
      Width           =   660
      _Version        =   589884
      _ExtentX        =   1164
      _ExtentY        =   926
      _StockProps     =   64
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   44
      Top             =   9075
      Width           =   13905
      _ExtentX        =   24527
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmKssStatistics.frx":1C556
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   20003
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   " "
            TextSave        =   " "
            Key             =   "������ɫ"
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
   Begin VSFlex8Ctl.VSFlexGrid vsTmp 
      Height          =   465
      Left            =   3495
      TabIndex        =   242
      TabStop         =   0   'False
      Top             =   75
      Visible         =   0   'False
      Width           =   1200
      _cx             =   2117
      _cy             =   820
      Appearance      =   2
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
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
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
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   120
      Top             =   90
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmKssStatistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mEnumPanel
    '�ϱ�����
    PanelItem_����ҩƷ���Ľ������ = 0         'picReportSub(0)  picBill
    PanelItem_���˿�����ҩ����������鼰���۱� = 1 'picReportSub(1)  picCY----(��)�������˿�����ҩ����������鼰���۱�
    PanelItem_���ﴦ��������ҩ����� = 2         'picReportSub(2)  picCF
    PanelItem_סԺ���˿�����ҩ����� = 3         'picReportSub(3)  picYZ

    'ͳ������
    PanelItem_����ҩ��ʹ���������ͳ�� = 0       'picOtherSub(0) ͳ��
    PanelItem_�����п�Χ����Ԥ����ҩͳ�� = 1     'picOtherSub(1)
    PanelItem_�ż��ﴦ��������ҩͳ�� = 2         'picOtherSub(2) ----��(��)�ﴦ��������ҩͳ��
    PanelItem_סԺҽ��������ҩͳ�� = 3           'picOtherSub(3)
    PanelItem_���󿹾�ҩ��ʹ�ó�N��ͳ�� = 4      'picOtherSub(4)
    PanelItem_ҽ������ĳ����������ҩ�ɱ�ͳ�� = 5 'picOtherSub(e_C3_lblN_ͳ�Ʊ�_����_5)ͳ��
End Enum

'����ؼ����±꣬��ʽ˵��  e_R0_cbo_12_ͳ��ʱ��,R0��ʾ�ϱ����ݵĵ�һ�����棬R1�ϱ��ڶ������棬C0ͳ�����ݵ�һ�����棬C1ͳ�����ݵڶ�������
Private Enum mCtlID
    e_R0_cboTimRP_ͳ��ʱ��_0 = 0
    e_R0_dtpRPS_��ʼʱ��_0 = 0
    e_R0_dtpRPE_����ʱ��_0 = 0
    e_R0_cmdOK_ͳ��_0 = 0
    e_R0_picFilter_��������_6 = 6
    e_R0_lblN_����_74 = 74
    
    e_R1_cboTimRP_ͳ��ʱ��_1 = 1
    e_R1_dtpRPS_��ʼʱ��_1 = 1
    e_R1_dtpRPE_����ʱ��_1 = 1
    e_R1_picFilter_��������_7 = 7
    e_R1_lblN_�׶�˵��_72 = 72
    e_R1_txtCount_��������_1 = 1
    e_R1_optType_��������_ƽ��_0 = 0
    e_R1_optType_��������_���_1 = 1
    e_R1_cmdOK_����_1 = 1
    e_R1_cmdDept_����ѡ����_1 = 1
    e_R1_txtDept_��������_1 = 1
    e_R1_lblN_������������_13 = 13
    e_R1_lblN_����_77 = 77
    
    e_R2_cboTimRP_ͳ��ʱ��_2 = 2
    e_R2_dtpRPS_��ʼʱ��_2 = 2
    e_R2_dtpRPE_����ʱ��_2 = 2
    e_R2_cmdDept_����ѡ����_2 = 2
    e_R2_txtDept_��������_2 = 2
    e_R2_txtCount_��������_2 = 2
    e_R2_optType_��������_ƽ��_3 = 3
    e_R2_optType_��������_���_2 = 2
    e_R2_cmdOK_ͳ��_2 = 2
    e_R2_picFilter_��������_8 = 8
    e_R2_lblN_�����_����_25 = 25
    e_R2_lblN_�����_����_26 = 26
    e_R2_lblN_�����_��������_27 = 27
    e_R2_lblN_������_����_70 = 70
    e_R2_lblN_������_����_29 = 29
    e_R2_lblN_������_��������_30 = 30
    
    e_R3_cboTimRP_ͳ��ʱ��_3 = 3
    e_R3_dtpRPS_��ʼʱ��_3 = 3
    e_R3_dtpRPE_����ʱ��_3 = 3
    e_R3_cmdDept_����ѡ����_3 = 3
    e_R3_txtDept_��������_3 = 3
    e_R3_cmdOK_ͳ��_3 = 3
    e_R3_picFilter_��������_9 = 9
    e_R3_lblN_�����_����_43 = 43
    e_R3_lblN_�����_��������_45 = 45
    e_R3_lblN_�׶�˵��_44 = 44
    
    e_C0_cboTimCount_ͳ��ʱ��_0 = 0
    e_C0_dtpCountS_��ʼʱ��_0 = 0
    e_C0_dtpCountE_����ʱ��_0 = 0
    e_C0_txtDept_��������_0 = 0
    e_C0_optType_ͳ�Ƴ���_סԺ_5 = 5
    e_C0_optType_ͳ�Ƴ���_����_4 = 4
    e_C0_optType_���ܷ�ʽ_����_9 = 9
    e_C0_optType_���ܷ�ʽ_ҽ��_8 = 8
    e_C0_optType_���ܷ�ʽ_ҩƷ_7 = 7
    e_C0_optType_����ʽ_����_12 = 12
    e_C0_optType_����ʽ_���_11 = 11
    e_C0_cmdOK_ͳ��_4 = 4
    e_C0_cmdDept_����ѡ����_0 = 0
    e_C0_picFilter_��������_0 = 0
    e_C0_lblN_����_75 = 75
    
    e_C1_cboTimCount_ͳ��ʱ��_1 = 1
    e_C1_dtpCountS_��ʼʱ��_1 = 1
    e_C1_dtpCountE_����ʱ��_1 = 1
    e_C1_txtDept_ͳ�ƿ���_4 = 4
    e_C1_cmdDept_����ѡ����_4 = 4
    e_C1_optType_���ܷ�ʽ_����_17 = 17
    e_C1_optType_���ܷ�ʽ_ҽ��_16 = 16
    e_C1_cmdOK_ͳ��_5 = 5
    e_C1_picFilter_��������_1 = 1
    e_C1_lblN_����_76 = 76
    
    e_C2_cboTimCount_ͳ��ʱ��_2 = 2
    e_C2_dtpCountS_��ʼʱ��_2 = 2
    e_C2_dtpCountE_����ʱ��_2 = 2
    e_C2_chkType_����_����_0 = 0
    e_C2_chkType_����_����_1 = 1
    e_C2_txtDept_ͳ�ƿ���_5 = 5
    e_C2_cmdDept_����ѡ����_5 = 5
    e_C2_txtNum_ͳ�ƿ���_0 = 0
    e_C2_optType_��������_ƽ��_6 = 6
    e_C2_optType_��������_���_10 = 10
    e_C2_cmdOK_ͳ��_6 = 6
    e_C2_picFilter_��������_2 = 2
    e_C2_lblN_ͳ�Ʊ�_����_4 = 4
    e_C2_lblN_ͳ�Ʊ�_����_46 = 46
    e_C2_lblN_ͳ�Ʊ�_��������_48 = 48
    e_C2_lblN_������_����_39 = 39
    e_C2_lblN_������_����_47 = 47
    e_C2_lblN_������_��������_49 = 49
    e_C2_lblN_�׶�˵��_50 = 50
    
    e_C3_cboTimCount_ͳ��ʱ��_3 = 3
    e_C3_dtpCountS_��ʼʱ��_3 = 3
    e_C3_dtpCountE_����ʱ��_3 = 3
    e_C3_txtDept_ͳ�ƿ���_6 = 6
    e_C3_cmdDept_����ѡ����_6 = 6
    e_C3_txtNum_��������_1 = 1
    e_C3_cmdOK_ͳ��_7 = 7
    e_C3_optType_�п�����_������_15 = 15
    e_C3_optType_�п�����_����_18 = 18
    e_C3_optType_��������_ƽ��_14 = 14
    e_C3_optType_��������_���_13 = 13
    e_C3_chkType_�п�����_����_2 = 2
    e_C3_chkType_�п�����_����_3 = 3
    e_C3_chkType_�п�����_����_4 = 4
    e_C3_chkType_�п�����_����_8 = 8
    e_C3_picFilter_��������_3 = 3
    e_C3_lblN_ͳ�Ʊ�_����_5 = 5
    e_C3_lblN_������_����_59 = 59
    e_C3_lblN_�׶�˵��_58 = 58
    
    e_C4_cboTimCount_ͳ��ʱ��_4 = 4
    e_C4_dtpCountS_��ʼʱ��_4 = 4
    e_C4_dtpCountE_����ʱ��_4 = 4
    e_C4_chkType_�п�����_����_5 = 5
    e_C4_chkType_�п�����_����_6 = 6
    e_C4_chkType_�п�����_����_7 = 7
    e_C4_chkType_�п�����_����_9 = 9
    e_C4_txtDept_ͳ�ƿ���_7 = 7
    e_C4_cmdDept_����ѡ����_7 = 7
    e_C4_txtNum_��������_2 = 2
    e_C4_optType_��������_ƽ��_21 = 21
    e_C4_optType_��������_���_22 = 22
    e_C4_cmdOK_ͳ��_8 = 8
    e_C4_picFilter_��������_4 = 4
    e_C4_lblN_ͳ�Ʊ�_����_6 = 6
    e_C4_lblN_�׶�˵��_73 = 73
    
    e_C5_cboTimCount_ͳ��ʱ��_5 = 5
    e_C5_dtpCountS_��ʼʱ��_5 = 5
    e_C5_dtpCountE_����ʱ��_5 = 5
    e_C5_optType_��ҽ_25 = 25
    e_C5_optType_��ҽ_26 = 26
    e_C5_optType_�����_27 = 27
    e_C5_optType_������_28 = 28
    e_C5_txtNum_��������_3 = 3
    e_C5_optType_��������_ƽ��_20 = 20
    e_C5_optType_��������_���_19 = 19
    e_C5_cmdOK_ͳ��_9 = 9
    e_C5_picFilter_��������_5 = 5
    e_C5_lblN_������_����_7 = 7
    e_C5_lblN_�׶�˵��_67 = 67
    
End Enum


Private Enum COL_VSBILL '�ϱ�����  ����ҩƷ���Ľ������  vsBill ������
    COL_ͳ����Ŀ = 0
    COL_��� = 1
    COL_��λ = 2
    COL_��ע = 3
End Enum

Private Enum ROW_VSBILL '�ϱ�����  ����ҩƷ���Ľ������   vsBill  ������
    ROW_��ҽԺ������ = 1
    ROW_�������� = 2
    ROW_��ҩƷ������ = 3
    ROW_ҩƷռҽԺ��������� = 4
    ROW_ҩƷ����������� = 5
    ROW_ҩƷ�����������ռҽԺ��������� = 6
    ROW_��ҩȫ��ʹ�ý�� = 7
    ROW_������ҩ�� = 8
    ROW_סԺ��ҩ�� = 9
    ROW_����ҩ��ȫ��ʹ�ý�� = 10
    ROW_������ҩ���� = 11
    ROW_סԺ��ҩ���� = 12
    ROW_����ҩ��ռҩƷ��������� = 13
End Enum

Private Enum PATIREPORT_COLUMN '�ϱ�����  ���˿�����ҩ����������鼰���۱�  rptPati
    COL_�༭ = 0
    COL_��ӡ
    
    col_����
    col_����
    col_�Ա�
    col_����
    col_סԺ��
    col_����
    col_סԺҽʦ
    col_��Ժ����
        
    col_����Id
    col_��ҳID
    COL_����ID
    COL_���
    COL_����ID
End Enum

Private Enum COL_CFADVICE '�ϱ����� ���ﴦ��������ҩ����� �ϰ벿�ֱ�� vsCF / ͳ������ �ż��ﴦ��������ҩͳ�� �ϰ벿�ֱ�� vsCountDruUse
    COL_CF��� = 0
    COL_CF�����
    COL_CF��������
    COL_CF��������
    COL_CF����ҽ��
    COL_CF����
    COL_CF��������
    COL_CF���
    COL_CFҩƷƷ����
    COL_CF��ҩƷ����
    COL_CFע���
    COL_CF��ҩƷ����
    COL_CFͨ����
    COL_CF���
    COL_CF����
    COL_CF���
    COL_CF�÷�����
    COL_CF��ҩ;��
    COL_CF�������
    COL_CFҩƷ���
    COL_CF��ҩ���
    
    COL_CF����ID
    COL_CF�Һ�ID
    COL_CF�Һŵ�
End Enum

Private Enum COL_YZADVICE '�ϱ�����   סԺ���˿�����ҩ�����  vsZYYY
    COL_YZ��� = 0
    COL_YZҩƷͨ���� = 1
    COL_YZ���� = 2
    COL_YZ��� = 3
    COL_YZ��λ = 4
    COL_YZ���� = 5
    COL_YZ�ܷ��� = 6
End Enum

Private Enum COL_VSUSERAN_DRUG 'ͳ������  ����ҩ��ʹ���������ͳ��  vsUseRan ��ҩƷ��ʽ����
    COL_D���� = 0
    COL_D�ܽ��
    COL_Dʹ������
    COL_D����������
    COL_Dÿ��ƽ�����
    COL_DDDDs
    COL_Dʹ��ǿ��
    COL_DռҩƷ������
End Enum

Private Enum COL_VSUSERAN_NUDRUG 'ͳ������  ����ҩ��ʹ���������ͳ��  vsUseRan �����һ�ҽ����ʽ����
    COL_UD��� = 0
    COL_UDҩƷ����
    COL_UD����
    COL_UD���
    COL_UD����
    COL_UD�ܽ��
    COL_UDʹ������
    COL_UD����������
    COL_UDÿ��ƽ�����
    COL_UDDDDs
    COL_UDʹ��ǿ��
    COL_UDռҩƷ�ܽ�����
End Enum

Private Enum COL_VSCUT 'ͳ������  �����п�Χ����Ԥ����ҩͳ�� vsCUT
    COL_CUT���� = 0
    COL_CUTʹ���˴�
    COL_CUTʹ����
    COL_CUT�п���
    COL_CUT��������
    COL_CUT�п�ʹ����
    COL_CUT��ǰ��ҩ
    COL_CUTƽ����ҩ
    COL_CUTƷ����
End Enum

Private Enum COL_VSINDRUUSE 'ͳ������  סԺҽ��������ҩͳ��  vsInDruUse
    COL_DRU��� = 0
    COL_DRUסԺ��
    COL_DRU��������
    COL_DRU��Ժ����
    COL_DRU����ҽ��
    COL_DRU����
    COL_DRU��Ժ���
    COL_DRU��������
    COL_DRUסԺ����
    COL_DRU���ƽ��
    COL_DRU�վ����ƽ��
    COL_DRU�п�����
    COL_DRUҩƷ����
    COL_DRU����ҩ��Ʒ����
    COL_DRUҩƷ���
    COL_DRU����ҩ����
    COL_DRU������ҩ
    
    COL_DRUҩƷ����
    COL_DRU����
    COL_DRU���
    COL_DRU�÷�����
    COL_DRU��ҩ����
    COL_DRU��ҩ;��
    COL_DRU��ҩĿ��
    
    COL_DRU����id
    COL_DRU��ҳid
End Enum

Private Enum COL_VSILLDRUUSE 'ͳ������  ҽ������ĳ����������ҩ�ɱ�ͳ��  vsIllDruUse
    COL_ILL����ҽ�� = 0
    COL_ILL��������
    COL_ILL�ÿ�ҩ����
    COL_ILL������
    COL_ILL�ܽ��
    COL_ILLҩƷ���
    COL_ILL�˾����ƶ�
    COL_ILL�˾��ս��
    COL_ILL��ҩ���
    COL_ILL��ҩƷ����
    
    COL_ILL����
    COL_ILL��ת
    COL_ILLδ��
    COL_ILL����
    COL_ILL����
    
    COL_ILLID
End Enum

Private Enum COL_VSOPEKSSUSE 'ͳ������  ���󿹾�ҩ��ʹ�ó�N��ͳ��  vsOpeKssUse
    COL_OPEסԺ�� = 0
    COL_OPE��������
    COL_OPE����
    COL_OPE��������
    COL_OPE�п�����
    COL_OPE������ҩ����
    
    COL_OPE����id
    COL_OPE��ҳid
End Enum

Private mstrPrivs As String
Private mlngModul As Long
Private mlngFind As Long
Private mdatCurr As Date
Private mstrMatch As String
Private mint���� As Integer '����ƥ�䷽ʽ��0-ƴ��,1-���
Private mlng����ID As Long
Private mlng���������� As Long
Private mlng������������ As Long
Private mlngRP���ﴦ���������� As Long '�ϱ����ݽ��棬����������ʵ������
Private mlngOT���ﴦ���������� As Long 'ͳ�����ݽ��棬����������ʵ������
Private mbln�������� As Boolean

Private Sub dtpRPS_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_Load()
    Dim tpGroup As TaskPanelGroup
    Dim tpGroupItem As TaskPanelGroupItem
    Dim i As Long
    Dim strCaption As String
    Dim strTmp As String, str��� As String
    Dim varArr As Variant
    
    mstrPrivs = gstrPrivs
    mlngModul = glngModul
    
    'TabControl
    '-----------------------------------------------------
    With Me.tbcSub
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        '���Ӵ���ʱ��Form_Load�����Զ�ѡ�е�һ������Ŀ�Ƭ
        '������õ�ǰ��Ƭ����,�򲻻��Զ��л�ѡ��,����ʾ����δ��
        '����ָ����������Ч�����ձ�Ϊ0-N��ֻ�ǿ��ܸı����˳��
        .InsertItem(0, "�ϱ�����ͳ��", picReport.hwnd, 0).Tag = "�ϱ�����ͳ��"
        .InsertItem(1, "��������ͳ��", picOther.hwnd, 0).Tag = "��������ͳ��"
        .Item(0).Selected = True
    End With
    strCaption = "����ҩƷ���Ľ������;(��)�������˿�����ҩ����������鼰���۱�;���ﴦ��������ҩ�����;סԺ���˿�����ҩ�����"
    With tbcReport
        With .PaintManager
            .Appearance = xtpTabAppearanceVisualStudio
            .Color = xtpTabColorOffice2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
        End With
        For i = 0 To UBound(Split(strCaption, ";"))
            .InsertItem(i, Split(strCaption, ";")(i), picReportSub(i).hwnd, 0).Tag = Split(strCaption, ";")(i)
        Next
    End With
    
    tbcReport.Item(tbcReport.ItemCount - 1).Selected = True
    tbcReport.Item(mEnumPanel.PanelItem_����ҩƷ���Ľ������).Selected = True
    tbcReport.Item(mEnumPanel.PanelItem_���˿�����ҩ����������鼰���۱�).Tag = "���˿�����ҩ����������鼰���۱�" 'Tag����һ��
    
    strCaption = "����ҩ��ʹ���������ͳ��;�����п�Χ����Ԥ����ҩͳ��;��(��)�ﴦ��������ҩͳ��;סԺҽ��������ҩͳ��;���󿹾�ҩ��ʹ�ó�N��ͳ��;ҽ������ĳ����������ҩ�ɱ�ͳ��"
    With tbcOther
        With .PaintManager
            .Appearance = xtpTabAppearanceVisualStudio
            .Color = xtpTabColorOffice2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
        End With
        For i = 0 To UBound(Split(strCaption, ";"))
            .InsertItem(i, Split(strCaption, ";")(i), picOtherSub(i).hwnd, 0).Tag = Split(strCaption, ";")(i)
        Next
    End With
    tbcOther.Item(tbcOther.ItemCount - 1).Selected = True
    tbcOther.Item(mEnumPanel.PanelItem_����ҩ��ʹ���������ͳ��).Selected = True
    tbcOther.Item(mEnumPanel.PanelItem_�ż��ﴦ��������ҩͳ��).Tag = "�ż��ﴦ��������ҩͳ��" 'Tag����һ��
    
    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '����VisualTheme����Ч
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    Call MainDefCommandBar
    
    mstrMatch = IIf(Val(zlDatabase.GetPara("����ƥ��", , , True)) = 0, "%", "")
    str��� = zlDatabase.GetPara("���Ƽ�������", glngSys, 1269, "")
    mint���� = Val(zlDatabase.GetPara("���뷽ʽ"))
    If InStr(str���, "|") > 0 Then
        varArr = Split(str���, "|")
        optType(e_C5_optType_��ҽ_25).Value = Val(varArr(0)) = 1
        optType(e_C5_optType_������_28).Value = Val(varArr(1)) = 1
        txtILL.Tag = Val(varArr(2))
        strTmp = varArr(0) & "|" & varArr(1) & "|" & varArr(2) & "|"
        txtILL.Text = Replace(str���, strTmp, "")
        cmdILL.Tag = txtILL.Text
    End If
    
    Me.lvwItems.ListItems.Clear
    With Me.lvwItems.ColumnHeaders
        .Clear
        .Add , "����", "����", 1500
        .Add , "����", "����", 900
    End With
    With Me.lvwItems
        .ColumnHeaders("����").Position = 1
        .SortKey = .ColumnHeaders("����").Index - 1
        .SortOrder = lvwAscending
        .Width = 3000
    End With
    
    Call LoadDept
    
    Call InitReportColumn
    
    Call InitVS���
    
    Call InitVS�������(vsMZYY, vsCF)
    Call InitVS�������(vsCountDruUse, vsCountCF)
    
    mdatCurr = zlDatabase.Currentdate
    Call InitTimeList
    Call Rest���ڷ�Χ
    Call RestoreWinState(Me, App.ProductName)

End Sub

Private Sub Rest���ڷ�Χ()
'���ܣ�ȡ��һ�ε����ڷ�Χ��ע�����ȡ
    Dim strTmp As String
    
    strTmp = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\���ڷ�Χ", "����ҩƷ���Ľ������", "")
    If strTmp <> "" Then
        If strTmp <> dtpRPS(e_R0_dtpRPS_��ʼʱ��_0).Value & "," & dtpRPE(e_R0_dtpRPE_����ʱ��_0).Value Then
            dtpRPS(e_R0_dtpRPS_��ʼʱ��_0).Value = CDate(Split(strTmp, ",")(0))
            dtpRPE(e_R0_dtpRPE_����ʱ��_0).Value = CDate(Split(strTmp, ",")(1))
            Call Cbo.Locate(cboTimRP(e_R0_dtpRPS_��ʼʱ��_0), "�Զ���", False)
        End If
    End If
    
    strTmp = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\���ڷ�Χ", "���˿�����ҩ����������鼰���۱�", "")
    If strTmp <> "" Then
        If strTmp <> dtpRPS(e_R1_dtpRPS_��ʼʱ��_1).Value & "," & dtpRPE(e_R1_dtpRPE_����ʱ��_1).Value Then
            dtpRPS(e_R1_dtpRPS_��ʼʱ��_1).Value = CDate(Split(strTmp, ",")(0))
            dtpRPE(e_R1_dtpRPE_����ʱ��_1).Value = CDate(Split(strTmp, ",")(1))
            Call Cbo.Locate(cboTimRP(e_R1_dtpRPS_��ʼʱ��_1), "�Զ���", False)
        End If
    End If
    
    strTmp = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\���ڷ�Χ", "���ﴦ��������ҩ�����", "")
    If strTmp <> "" Then
        If strTmp <> dtpRPS(e_R2_dtpRPS_��ʼʱ��_2).Value & "," & dtpRPE(e_R2_dtpRPE_����ʱ��_2).Value Then
            dtpRPS(e_R2_dtpRPS_��ʼʱ��_2).Value = CDate(Split(strTmp, ",")(0))
            dtpRPE(e_R2_dtpRPE_����ʱ��_2).Value = CDate(Split(strTmp, ",")(1))
            Call Cbo.Locate(cboTimRP(e_R2_dtpRPS_��ʼʱ��_2), "�Զ���", False)
        End If
    End If
    
    strTmp = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\���ڷ�Χ", "סԺ���˿�����ҩ�����", "")
    If strTmp <> "" Then
        If strTmp <> dtpRPS(e_R3_dtpRPS_��ʼʱ��_3).Value & "," & dtpRPE(e_R3_dtpRPE_����ʱ��_3).Value Then
            dtpRPS(e_R3_dtpRPS_��ʼʱ��_3).Value = CDate(Split(strTmp, ",")(0))
            dtpRPE(e_R3_dtpRPE_����ʱ��_3).Value = CDate(Split(strTmp, ",")(1))
            Call Cbo.Locate(cboTimRP(e_R3_dtpRPS_��ʼʱ��_3), "�Զ���", False)
        End If
    End If
    
    
    strTmp = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\���ڷ�Χ", "����ҩ��ʹ���������ͳ��", "")
    If strTmp <> "" Then
        If strTmp <> dtpCountS(e_C0_dtpCountS_��ʼʱ��_0).Value & "," & dtpCountE(e_C0_dtpCountE_����ʱ��_0).Value Then
            dtpCountS(e_C0_dtpCountS_��ʼʱ��_0).Value = CDate(Split(strTmp, ",")(0))
            dtpCountE(e_C0_dtpCountE_����ʱ��_0).Value = CDate(Split(strTmp, ",")(1))
            Call Cbo.Locate(cboTimCount(e_C0_dtpCountS_��ʼʱ��_0), "�Զ���", False)
        End If
    End If
    
    strTmp = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\���ڷ�Χ", "�����п�Χ����Ԥ����ҩͳ��", "")
    If strTmp <> "" Then
        If strTmp <> dtpCountS(e_C1_dtpCountS_��ʼʱ��_1).Value & "," & dtpCountE(e_C1_dtpCountE_����ʱ��_1).Value Then
            dtpCountS(e_C1_dtpCountS_��ʼʱ��_1).Value = CDate(Split(strTmp, ",")(0))
            dtpCountE(e_C1_dtpCountE_����ʱ��_1).Value = CDate(Split(strTmp, ",")(1))
            Call Cbo.Locate(cboTimCount(e_C1_dtpCountS_��ʼʱ��_1), "�Զ���", False)
        End If
    End If
    
    
    strTmp = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\���ڷ�Χ", "�ż��ﴦ��������ҩͳ��", "")
    If strTmp <> "" Then
        If strTmp <> dtpCountS(e_C2_dtpCountS_��ʼʱ��_2).Value & "," & dtpCountE(e_C2_dtpCountE_����ʱ��_2).Value Then
            dtpCountS(e_C2_dtpCountS_��ʼʱ��_2).Value = CDate(Split(strTmp, ",")(0))
            dtpCountE(e_C2_dtpCountE_����ʱ��_2).Value = CDate(Split(strTmp, ",")(1))
            Call Cbo.Locate(cboTimCount(e_C2_dtpCountS_��ʼʱ��_2), "�Զ���", False)
        End If
    End If
    
    strTmp = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\���ڷ�Χ", "סԺҽ��������ҩͳ��", "")
    If strTmp <> "" Then
        If strTmp <> dtpCountS(e_C3_dtpCountS_��ʼʱ��_3).Value & "," & dtpCountE(e_C3_dtpCountE_����ʱ��_3).Value Then
            dtpCountS(e_C3_dtpCountS_��ʼʱ��_3).Value = CDate(Split(strTmp, ",")(0))
            dtpCountE(e_C3_dtpCountE_����ʱ��_3).Value = CDate(Split(strTmp, ",")(1))
            Call Cbo.Locate(cboTimCount(e_C3_dtpCountS_��ʼʱ��_3), "�Զ���", False)
        End If
    End If
    
    strTmp = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\���ڷ�Χ", "���󿹾�ҩ��ʹ�ó�N��ͳ��", "")
    If strTmp <> "" Then
        If strTmp <> dtpCountS(e_C4_dtpCountS_��ʼʱ��_4).Value & "," & dtpCountE(e_C4_dtpCountE_����ʱ��_4).Value Then
            dtpCountS(e_C4_dtpCountS_��ʼʱ��_4).Value = CDate(Split(strTmp, ",")(0))
            dtpCountE(e_C4_dtpCountE_����ʱ��_4).Value = CDate(Split(strTmp, ",")(1))
            Call Cbo.Locate(cboTimCount(e_C4_dtpCountS_��ʼʱ��_4), "�Զ���", False)
        End If
    End If
    
    strTmp = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\���ڷ�Χ", "ҽ������ĳ����������ҩ�ɱ�ͳ��", "")
    If strTmp <> "" Then
        If strTmp <> dtpCountS(e_C5_dtpCountS_��ʼʱ��_5).Value & "," & dtpCountE(e_C5_dtpCountE_����ʱ��_5).Value Then
            dtpCountS(e_C5_dtpCountS_��ʼʱ��_5).Value = CDate(Split(strTmp, ",")(0))
            dtpCountE(e_C5_dtpCountE_����ʱ��_5).Value = CDate(Split(strTmp, ",")(1))
            Call Cbo.Locate(cboTimCount(e_C5_dtpCountS_��ʼʱ��_5), "�Զ���", False)
        End If
    End If
    
End Sub

Private Sub InitTimeList()
    Dim strDate As String
    Dim strTmp As String
    Dim lngTmp As Long
    Dim str���һ�� As String
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim rs��� As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo errH
    
    strDate = Format(mdatCurr, "yyyy-MM-dd hh:mm:ss")
    lngTmp = Val(Split(strDate, "-")(0)) - 1
    
    
    strSql = "Select Distinct Substr(�ڼ�, 1, 4)||'��' As ��� From �ڼ�� Order By ���"
    Set rs��� = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    
    With cboTimRP(e_R0_cboTimRP_ͳ��ʱ��_0)
        .Clear
        For i = 1 To rs���.RecordCount
            .AddItem rs���!��� & ""
            rs���.MoveNext
        Next
        .AddItem "�Զ���"
    End With
    Call Cbo.Locate(cboTimRP(e_R0_cboTimRP_ͳ��ʱ��_0), lngTmp & "��", False)
    
    
    strDate = Format(DateAdd("m", -1, mdatCurr), "yyyy-MM-dd hh:mm:ss")
    
    strSql = "Select ��ʼ����, ��ֹ���� From �ڼ�� Where �ڼ� = [1]"
    strTmp = Split(strDate, "-")(0) & Split(strDate, "-")(1)
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strTmp)
    
    With cboTimRP(e_R1_cboTimRP_ͳ��ʱ��_1)
        .Clear
        If Not rsTmp.EOF Then
            .Tag = Format(rsTmp!��ʼ���� & "", "yyyy-MM-dd") & "," & Format(rsTmp!��ֹ���� & "", "yyyy-MM-dd")
        End If
        .AddItem Split(strDate, "-")(0) & "��" & Split(strDate, "-")(1) & "��"
        .AddItem "�Զ���"
        .ListIndex = 0
    End With
    
    With cboTimRP(e_R2_cboTimRP_ͳ��ʱ��_2)
        .Clear
        .Tag = cboTimRP(1).Tag
        .AddItem Split(strDate, "-")(0) & "��" & Split(strDate, "-")(1) & "��"
        .AddItem "�Զ���"
        .ListIndex = 0
    End With
    
    '��һ���ȣ�1�£�3�µڶ����ȣ�4�£�6�µ������ȣ�7�£�9�µ��ļ��ȣ�10�£�12��
    strDate = Format(DateAdd("m", -3, mdatCurr), "yyyy-MM-dd hh:mm:ss")
    lngTmp = Val(Split(strDate, "-")(1))
    strTmp = Split(strDate, "-")(0)
    If lngTmp >= 1 And lngTmp <= 3 Then
        cboTimRP(e_R3_cboTimRP_ͳ��ʱ��_3).Tag = strTmp & "-01-01," & strTmp & "-03-31"
        strTmp = strTmp & "��1����"
    ElseIf lngTmp >= 4 And lngTmp <= 6 Then
        cboTimRP(e_R3_cboTimRP_ͳ��ʱ��_3).Tag = strTmp & "-04-01," & strTmp & "-06-30"
        strTmp = strTmp & "��2����"
    ElseIf lngTmp >= 7 And lngTmp <= 9 Then
        cboTimRP(e_R3_cboTimRP_ͳ��ʱ��_3).Tag = strTmp & "-07-01," & strTmp & "-09-30"
        strTmp = strTmp & "��3����"
    ElseIf lngTmp >= 10 And lngTmp <= 12 Then
        cboTimRP(e_R3_cboTimRP_ͳ��ʱ��_3).Tag = strTmp & "-10-01," & strTmp & "-12-31"
        strTmp = strTmp & "��4����"
    End If
    
    With cboTimRP(e_R3_cboTimRP_ͳ��ʱ��_3)
        .Clear
        .AddItem strTmp
        .AddItem "�Զ���"
        .ListIndex = 0
    End With
    
    For i = 0 To 5
        With cboTimCount(i)
            .Clear
            .Tag = cboTimRP(e_R1_cboTimRP_ͳ��ʱ��_1).Tag
            .AddItem "���һ��"
            .AddItem "�Զ���"
            .ListIndex = 0
        End With
    Next
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub dtpRPE_Change(Index As Integer)
    Call Cbo.Locate(cboTimRP(Index), "�Զ���", False)
End Sub

Private Sub dtpRPS_Change(Index As Integer)
    Call Cbo.Locate(cboTimRP(Index), "�Զ���", False)
End Sub

Private Sub dtpCountE_Change(Index As Integer)
    Call Cbo.Locate(cboTimCount(Index), "�Զ���", False)
End Sub

Private Sub dtpCountS_Change(Index As Integer)
    Call Cbo.Locate(cboTimCount(Index), "�Զ���", False)
End Sub

Private Sub cboTimCount_Click(Index As Integer)
    With cboTimCount(Index)
        If .ListIndex = 0 And .Tag <> "" Then
            dtpCountS(Index).Value = CDate(Split(.Tag, ",")(0))
            dtpCountE(Index).Value = CDate(Split(.Tag, ",")(1))
        End If
    End With
End Sub

Private Sub cboTimRP_Click(Index As Integer)
    Dim strTmp As String
    Dim lngYear As Long
    
    'ȡ��ǰ���ڵ����
    strTmp = Format(mdatCurr, "yyyy-MM-dd")
    lngYear = Val(Split(strTmp, "-")(0)) - 1
    
    Select Case Index
    Case e_R0_cboTimRP_ͳ��ʱ��_0
        With cboTimRP(Index)
            If .Text <> "�Զ���" Then
                strTmp = Mid(.Text, 1, 4) & "-01-01"
                dtpRPS(Index).Value = CDate(strTmp)
                
                strTmp = Mid(.Text, 1, 4) & "-12-31"
                dtpRPE(Index).Value = CDate(strTmp)
            End If
        End With
    Case Else
        With cboTimRP(Index)
            If .ListIndex = 0 And .Tag <> "" Then
                dtpRPS(Index).Value = CDate(Split(.Tag, ",")(0))
                dtpRPE(Index).Value = CDate(Split(.Tag, ",")(1))
            End If
        End With
    End Select
End Sub

Private Sub InitReportColumn()
    Dim objCol As ReportColumn

    With rptPati
        Set objCol = .Columns.Add(COL_�༭, "�༭", 30, False)
        objCol.Sortable = False
        objCol.AllowDrag = False
        objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Columns.Add(COL_��ӡ, "��ӡ", 30, False)
        objCol.Sortable = False
        objCol.AllowDrag = False
        objCol.Alignment = xtpAlignmentCenter
       
        Set objCol = .Columns.Add(col_����, "����", 50, False)
            objCol.Visible = False
  
        Set objCol = .Columns.Add(col_����, "����", 600, True)
            objCol.Groupable = False
            
        Set objCol = .Columns.Add(col_�Ա�, "�Ա�", 400, True)
            objCol.Alignment = xtpAlignmentLeft
            
        Set objCol = .Columns.Add(col_����, "����", 400, True)
            objCol.Alignment = xtpAlignmentLeft
            
        Set objCol = .Columns.Add(col_סԺ��, "סԺ��", 600, True)
            objCol.Alignment = xtpAlignmentLeft
            
        Set objCol = .Columns.Add(col_����, "����", 600, True)
            objCol.Alignment = xtpAlignmentLeft
            
        Set objCol = .Columns.Add(col_סԺҽʦ, "סԺҽʦ", 600, True)
            objCol.Alignment = xtpAlignmentLeft
            
        Set objCol = .Columns.Add(col_��Ժ����, "��Ժ����", 800, True)
            objCol.Alignment = xtpAlignmentCenter
        
        For Each objCol In .Columns
            objCol.Editable = False
            objCol.Groupable = objCol.Index = col_����
        Next
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ�Ĳ���..."
        End With
        .AutoColumnSizing = False
        .AllowColumnRemove = False
        .ShowGroupBox = True
        .ShowItemsInGroups = False
        .PreviewMode = True
        .MultipleSelection = False '������SelectionChanged�¼�
        .SetImageList Me.img16
        
        .GroupsOrder.Add .Columns(col_����)
        .GroupsOrder(0).SortAscending = True '����֮��,��������в���ʾ,�����е������ǲ����
        
        '����֮�����ʧȥ��¼���е�˳��,���ǿ�м���������
        .SortOrder.Add .Columns(col_����)
        .SortOrder(0).SortAscending = True
        .SortOrder.Add .Columns(col_��Ժ����)
        .SortOrder(1).SortAscending = True
    End With
End Sub

Private Sub MainDefCommandBar()
'���ܣ������ڲ˵����岿��
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom
    
    Dim lngCount As Long
    
    '�˵�����
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��")

        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)")
            objControl.BeginGroup = True
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "������(&T)")
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Archive, "���Ӳ�������(&I)")
            objControl.BeginGroup = True
            
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)")
            objControl.BeginGroup = True
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "��̳(&F)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��")
            objControl.BeginGroup = True
    End With


    '����������:������������
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("������", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")


        Set objControl = .Add(xtpControlButton, conMenu_Tool_Archive, "����")
            objControl.BeginGroup = True
            
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "����")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With

    '����һЩ�������ȼ���
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyAdd, conMenu_View_Expend_AllExpend 'չ��������
        .Add FCONTROL, vbKeySubtract, conMenu_View_Expend_AllCollapse '�۵�������
        .Add FCONTROL, vbKeyP, conMenu_File_Print '��ӡ
        .Add 0, vbKeyF5, conMenu_View_Refresh 'ˢ��
        .Add 0, vbKeyF1, conMenu_Help_Help '����
    End With
    
    '�ָ����̶���һЩ�˵�����
    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.SetIconSize 16, 16
    For lngCount = 2 To cbsMain.Count
        cbsMain(lngCount).ContextMenuPresent = False
        cbsMain(lngCount).ShowTextBelowIcons = False
        cbsMain(lngCount).EnableDocking xtpFlagHideWrap Or xtpFlagStretched
        For Each objControl In cbsMain(lngCount).Controls
            objControl.Style = xtpButtonIconAndCaption
        Next
    Next
    
    '��ȡ��������ģ��ı���(��������ģ���)
    '-----------------------------------------------------
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModul, mstrPrivs, "ZL1_INSIDE_1269_1", "ZL1_INSIDE_1269_2")
End Sub

Private Sub InitTable(ByRef vsgInfo As VSFlexGrid, ByVal strHead As String)
    Dim arrHead As Variant, i As Long
    
    arrHead = Split(strHead, ";")
    With vsgInfo
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
    End With
End Sub

Private Sub InitVS���()
'���ܣ���ʼ��VS��񣬹̶��е����õ�
    Dim strTmp As String
    Dim i As Integer
    
    '�ϱ����� ����ҩƷ���Ľ������
    strTmp = "ͳ����Ŀ,3570,1;���,1900,7;��λ,740,4;��ע,2890,1"
    Call InitTable(vsBill, strTmp)
    
    '����һ������
    With vsBill
        .Rows = .FixedRows
        .Rows = 14
        .RowHeight(0) = 450
        For i = 1 To 13
            .RowHeight(i) = 450
        Next
        .TextMatrix(ROW_��ҽԺ������, COL_ͳ����Ŀ) = "һ����ҽԺ�����루��"
        .TextMatrix(ROW_��������, COL_ͳ����Ŀ) = "�������������"
        .TextMatrix(ROW_��ҩƷ������, COL_ͳ����Ŀ) = "������ҩƷ�����루��"
        .TextMatrix(ROW_ҩƷռҽԺ���������, COL_ͳ����Ŀ) = "�ġ�ҩƷռҽԺ���������"
        .TextMatrix(ROW_ҩƷ�����������, COL_ͳ����Ŀ) = "�塢ҩƷ����������루��"
        .TextMatrix(ROW_ҩƷ�����������ռҽԺ���������, COL_ͳ����Ŀ) = "����ҩƷ�����������ռҽԺ���������"
        .TextMatrix(ROW_��ҩȫ��ʹ�ý��, COL_ͳ����Ŀ) = "�ߡ���ҩȫ��ʹ�ý����ۼۣ�"
        .TextMatrix(ROW_������ҩ��, COL_ͳ����Ŀ) = "    ���У�������ҩ��"
        .TextMatrix(ROW_סԺ��ҩ��, COL_ͳ����Ŀ) = "          סԺ��ҩ��"
        .TextMatrix(ROW_����ҩ��ȫ��ʹ�ý��, COL_ͳ����Ŀ) = "�ˡ�����ҩ��ȫ��ʹ�ý����ۼۣ�"
        .TextMatrix(ROW_������ҩ����, COL_ͳ����Ŀ) = "    ���У�������ҩ��"
        .TextMatrix(ROW_סԺ��ҩ����, COL_ͳ����Ŀ) = "          סԺ��ҩ��"
        .TextMatrix(ROW_����ҩ��ռҩƷ���������, COL_ͳ����Ŀ) = "�š�����ҩ��ռҩƷ���������"
    End With
    
    
    '�ϱ����� סԺ���˿�����ҩ�����
    strTmp = "���,1500,4;ҩƷͨ����,2200,1;����,1600,4;���,2800,1;��λ,700,4;����,800,7;�ܷ���(Ԫ),1200,7"
    Call InitTable(vsZYYY, strTmp)
    
    'ͳ�Ʋ��� ��ʼ��ͳ�Ʋ��ֱ��  ����ҩ��ʹ���������ͳ��  ����
    strTmp = "���,3100,4;ҩƷ����,3600,4;����,2000,4;���,2800,4;����,1020,4;�ܽ��(Ԫ),1200,4;ʹ������,530,4;ÿ��ƽ�����(Ԫ),760,4;DDDs,500,4;ʹ��ǿ��,550,4;ռҩƷ�ܽ�����(%),780,4"
    Call InitTable(vsUseRan, strTmp)
    
    'ͳ�Ʋ��� ��ʼ�� �����п�Χ����Ԥ����ҩͳ�� �����񣭣�vsCut
    With vsCut
        .Clear
        .FixedRows = 2: .FixedCols = 0
        .MergeCellsFixed = flexMergeFixedOnly
        .MergeRow(0) = True
        .MergeCol(0) = True
        .Cols = 9
        For i = COL_CUT���� To COL_CUTƷ����
            .ColAlignment(i) = flexAlignCenterCenter
            .ColWidth(i) = 2000
        Next
        
        .Cell(flexcpText, 0, 0, 1, 0) = "��������"
        .Cell(flexcpText, 0, 1, 0, 2) = "����ҩ��ʹ�����"
        .Cell(flexcpText, 0, COL_CUT�п���, 0, COL_CUTƷ����) = "�����п�Χ����Ԥ��ʹ�ÿ���ҩ�����"
        .TextMatrix(1, COL_CUTʹ���˴�) = "ʹ���˴�(��)"
        .TextMatrix(1, COL_CUTʹ����) = "ʹ����(%)"
        .TextMatrix(1, COL_CUT�п���) = "�����п���(��)"
        .TextMatrix(1, COL_CUT��������) = "����ҩԤ��ʹ����(��)"
        .TextMatrix(1, COL_CUT�п�ʹ����) = "ʹ����(%)"
        .TextMatrix(1, COL_CUT��ǰ��ҩ) = "��ǰ��ҩ��"
        .TextMatrix(1, COL_CUTƽ����ҩ) = "ƽ����ҩ����"
        .TextMatrix(1, COL_CUTƷ����) = "ƽ����ҩƷ����"
    End With
    
    'ͳ�Ʋ��� ����   סԺҽ��������ҩͳ��  �������ʼ��
    With vsInDruUse
        .Clear
        .FixedRows = 2: .FixedCols = 0
        .MergeCells = flexMergeFixedOnly
        .MergeRow(0) = True
        .Cols = 26
        For i = COL_DRU��� To COL_DRU��ҩĿ��
            .ColAlignment(i) = flexAlignLeftCenter
            .MergeCol(i) = True
        Next
        .ColWidth(COL_DRU���) = 450
        .ColAlignment(COL_DRU���) = flexAlignCenterCenter
        
        .ColWidth(COL_DRUסԺ��) = 800
        .ColWidth(COL_DRU��������) = 930
        .ColWidth(COL_DRU����ҽ��) = .ColWidth(COL_DRU��������)
        .ColWidth(COL_DRU����) = 1400
        .ColWidth(COL_DRU��ҩ;��) = .ColWidth(COL_DRU��������)
        
        .ColWidth(COL_DRU��Ժ����) = 990
        .ColAlignment(COL_DRU��Ժ����) = flexAlignCenterCenter
        
        .ColWidth(COL_DRU��Ժ���) = 2300
        
        .ColWidth(COL_DRUסԺ����) = 600
        .ColAlignment(COL_DRUסԺ����) = flexAlignRightCenter
        
        .ColWidth(COL_DRU���ƽ��) = 1060
        .ColAlignment(COL_DRU���ƽ��) = flexAlignRightCenter
        
        .ColWidth(COL_DRU�վ����ƽ��) = 900
        .ColAlignment(COL_DRU�վ����ƽ��) = flexAlignRightCenter
        
        .ColWidth(COL_DRU�п�����) = 450
        
        .ColWidth(COL_DRUҩƷ����) = .ColWidth(COL_DRU���)
        .ColAlignment(COL_DRUҩƷ����) = flexAlignRightCenter
        
        .ColWidth(COL_DRU����ҩ��Ʒ����) = .ColWidth(COL_DRU���)
        .ColAlignment(COL_DRU����ҩ��Ʒ����) = flexAlignRightCenter
        
        .ColWidth(COL_DRUҩƷ���) = 1080
        .ColAlignment(COL_DRUҩƷ���) = flexAlignRightCenter
        
        .ColWidth(COL_DRU����ҩ����) = 1035
        .ColAlignment(COL_DRU����ҩ����) = flexAlignRightCenter
        
        .ColWidth(COL_DRU������ҩ) = 820
        .ColWidth(COL_DRUҩƷ����) = 2500
        .ColWidth(COL_DRU����) = 700
        .ColAlignment(COL_DRU����) = flexAlignCenterCenter
        
        .ColWidth(COL_DRU���) = 2800
        .ColWidth(COL_DRU�÷�����) = 1260
        .ColWidth(COL_DRU��ҩ����) = 450
        .ColAlignment(COL_DRU��ҩ����) = flexAlignRightCenter
        
        .ColWidth(COL_DRU��ҩ;��) = 1125
        
        .ColWidth(COL_DRU��ҩĿ��) = 600
        .ColAlignment(COL_DRU��ҩĿ��) = flexAlignCenterCenter
        
        .Cell(flexcpText, 0, COL_DRU���, 1, COL_DRU���) = "���"
        .Cell(flexcpText, 0, COL_DRUסԺ��, 1, COL_DRUסԺ��) = "סԺ��"
        .Cell(flexcpText, 0, COL_DRU��������, 1, COL_DRU��������) = "��������"
        .Cell(flexcpText, 0, COL_DRU��Ժ����, 1, COL_DRU��Ժ����) = "��Ժ����"
        .Cell(flexcpText, 0, COL_DRU����ҽ��, 1, COL_DRU����ҽ��) = "����ҽ��"
        .Cell(flexcpText, 0, COL_DRU����, 1, COL_DRU����) = "����"
        .Cell(flexcpText, 0, COL_DRU��Ժ���, 1, COL_DRU��Ժ���) = "��Ժ���"
        .Cell(flexcpText, 0, COL_DRU��������, 1, COL_DRU��������) = "��������"
        .Cell(flexcpText, 0, COL_DRUסԺ����, 1, COL_DRUסԺ����) = "סԺ����"
        .Cell(flexcpText, 0, COL_DRU���ƽ��, 1, COL_DRU���ƽ��) = "���ƽ��(Ԫ)"
        .Cell(flexcpText, 0, COL_DRU�վ����ƽ��, 1, COL_DRU�վ����ƽ��) = "�վ����ƽ��(Ԫ)"
        .Cell(flexcpText, 0, COL_DRU�п�����, 1, COL_DRU�п�����) = "�п�����"
        .Cell(flexcpText, 0, COL_DRUҩƷ����, 1, COL_DRUҩƷ����) = "ҩƷ����"
        .Cell(flexcpText, 0, COL_DRU����ҩ��Ʒ����, 1, COL_DRU����ҩ��Ʒ����) = "����ҩ��Ʒ����"
        .Cell(flexcpText, 0, COL_DRUҩƷ���, 1, COL_DRUҩƷ���) = "ҩƷ���(Ԫ)"
        .Cell(flexcpText, 0, COL_DRU����ҩ����, 1, COL_DRU����ҩ����) = "����ҩ����(Ԫ)"
        .Cell(flexcpText, 0, COL_DRU������ҩ, 1, COL_DRU������ҩ) = "������ҩ"
        
        .Cell(flexcpText, 0, COL_DRUҩƷ����, 0, COL_DRU��ҩĿ��) = "����ҩʹ�����(�����÷�)"
        
        .TextMatrix(1, COL_DRUҩƷ����) = "ҩƷ����"
        .TextMatrix(1, COL_DRU����) = "����"
        .TextMatrix(1, COL_DRU���) = "���"
        .TextMatrix(1, COL_DRU�÷�����) = "�÷�����"
        .TextMatrix(1, COL_DRU��ҩ����) = "����"
        .TextMatrix(1, COL_DRU��ҩ;��) = "��ҩ;��"
        .TextMatrix(1, COL_DRU��ҩĿ��) = "Ŀ��"
        
        .ColHidden(COL_DRU����id) = True
        .ColHidden(COL_DRU��ҳid) = True
        
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, COL_DRU��ҩĿ��) = flexAlignCenterCenter
        
        .WordWrap = True
    End With
    
    With vsInDruAna
        .Clear
        .Cols = 4: .Rows = 0
        .RowHeightMin = 300
        For i = 0 To .Cols - 1
            .ColWidth(i) = 5000
        Next
        .WordWrap = True
        .AddItem "" '��0��
        .TextMatrix(.Rows - 1, 0) = "A(��ҩ��Ʒ����)=0��"
        .TextMatrix(.Rows - 1, 1) = "B(ƽ����ҩƷ����A/C)=0��"
        .TextMatrix(.Rows - 1, 2) = "C(ʹ�ÿ���ҩ���Ʒ����)=0��"
        .TextMatrix(.Rows - 1, 3) = "D(ʹ�ÿ���ҩ��İٷ���C/A)*100%=0%"
        
        .AddItem "" '��1��
        .TextMatrix(.Rows - 1, 0) = "E(ʹ�ÿ�ҩ��Ĳ�����)=0��"
        .TextMatrix(.Rows - 1, 1) = "F(��Ժ���˿���ҩ��ʹ����E/ʵ������)*100%=0%"
        .TextMatrix(.Rows - 1, 2) = "G(�����ܽ��)=0Ԫ"
        .TextMatrix(.Rows - 1, 3) = "H(����ƽ�����ƽ��G/ʵ������)=0Ԫ"
        
        .AddItem "" '��2��
        .TextMatrix(.Rows - 1, 0) = "I(ҩƷ�ܽ��)=0Ԫ"
        .TextMatrix(.Rows - 1, 1) = "L(ҩƷ�ܽ��ռ�����ܽ��İٷ���I/G)*100%=0%"
        .TextMatrix(.Rows - 1, 2) = "K(����ҩ���ܽ��)=0Ԫ"
        .TextMatrix(.Rows - 1, 3) = "J(����ҩ���ܽ��ռҩƷ�ܽ��İٷ���K/I)*100%=0%"
        
        .AddItem "" '��3��
        .TextMatrix(.Rows - 1, 0) = "M(���ÿ���ҩ��Ĳ�����)=0��"
        .TextMatrix(.Rows - 1, 1) = "O(���ÿ���ҩ���ʹ����M/E)*100%=0%"
        .TextMatrix(.Rows - 1, 2) = "P(����ʹ�ÿ���ҩ��Ĳ�����)=0��"
        .TextMatrix(.Rows - 1, 3) = "Q(����ʹ�ÿ���ҩ���ʹ����P/E)*100%��0%"
        
        .AddItem "" '��4��
        .TextMatrix(.Rows - 1, 0) = "R(����ʹ�ÿ���ҩ��Ĳ�����)=0��"
        .TextMatrix(.Rows - 1, 1) = "S(����ʹ�ÿ���ҩ���ʹ����R/E)*100%��0%"
        .TextMatrix(.Rows - 1, 2) = "T(����ʹ�ÿ���ҩ��Ĳ�����)=0��"
        .TextMatrix(.Rows - 1, 3) = "U(����ʹ�ÿ���ҩ���ʹ����T/E)*100%��0%"
        
        .AddItem "" '��5��
        .TextMatrix(.Rows - 1, 0) = "V(Ԥ��ʹ�ÿ���ҩ��Ĳ�����)=0��"
        .TextMatrix(.Rows - 1, 1) = "W(Ԥ��ʹ�ÿ���ҩ�ﹹ�ɱ�V/E)100%=0%"
        .TextMatrix(.Rows - 1, 2) = "X(����ʹ�ÿ���ҩ��Ĳ�����)=0��"
        .TextMatrix(.Rows - 1, 3) = "Y(����ʹ�ÿ���ҩ�ﹹ�ɱ�Y/E)*100%=0%"
        
    End With
    
    'ͳ�Ʋ��� ���󿹾�ҩ��ʹ�ó�N��ͳ�� ���� vsOpeKssUse
    strTmp = "סԺ��,1500,1;��������,2000,4;����,2000,4;��������,5000,1;�п�����,1000,4;������ҩ����,2000,4;����id;��ҳid"
    Call InitTable(vsOpeKssUse, strTmp)
    
    'ͳ�Ʋ���   ҽ������ĳ����������ҩ�ɱ�ͳ�� ���棬 vsIllDruUse
    With vsIllDruUse
        .Clear
        .FixedRows = 2: .FixedCols = 0
        .MergeCells = flexMergeFixedOnly
        .MergeRow(0) = True
        .Cols = 16
        For i = COL_ILL����ҽ�� To COL_ILL����
            .ColAlignment(i) = flexAlignRightCenter
            .MergeCol(i) = True
        Next
        .ColAlignment(COL_ILL����ҽ��) = flexAlignCenterCenter
        .ColHidden(COL_ILLID) = True
        .ColWidth(COL_ILL����ҽ��) = 1500
        .ColWidth(COL_ILL��������) = 1270
        .ColWidth(COL_ILL�ÿ�ҩ����) = 1270
        .ColWidth(COL_ILL������) = 1070
        
        .ColWidth(COL_ILL�ܽ��) = 1080
        .ColWidth(COL_ILLҩƷ���) = 1080
        .ColWidth(COL_ILL�˾����ƶ�) = 1300
        .ColWidth(COL_ILL�˾��ս��) = 1300
        .ColWidth(COL_ILL��ҩ���) = 1300
        .ColWidth(COL_ILL��ҩƷ����) = 830
        .ColWidth(COL_ILL����) = 550
        .ColWidth(11) = 550
        .ColWidth(12) = 550
        .ColWidth(13) = 550
        .ColWidth(14) = 550
        
        .Cell(flexcpText, 0, COL_ILL����ҽ��, 1, COL_ILL����ҽ��) = "����ҽ��"
        .Cell(flexcpText, 0, COL_ILL��������, 1, COL_ILL��������) = "��������"
        .Cell(flexcpText, 0, COL_ILL�ÿ�ҩ����, 1, COL_ILL�ÿ�ҩ����) = "ʹ�ÿ���ҩ�������"
        .Cell(flexcpText, 0, COL_ILL������, 1, COL_ILL������) = "������(%)"
        .Cell(flexcpText, 0, COL_ILL�ܽ��, 1, COL_ILL�ܽ��) = "�ܽ��(Ԫ)"
        .Cell(flexcpText, 0, COL_ILLҩƷ���, 1, COL_ILLҩƷ���) = "ҩƷ�ܽ��(Ԫ)"
        .Cell(flexcpText, 0, COL_ILL�˾����ƶ�, 1, COL_ILL�˾����ƶ�) = "�˾����ƽ��(Ԫ)"
        .Cell(flexcpText, 0, COL_ILL�˾��ս��, 1, COL_ILL�˾��ս��) = "�˾��ս��(Ԫ)"
        .Cell(flexcpText, 0, COL_ILL��ҩ���, 1, COL_ILL��ҩ���) = "����ҩ���ܽ��(Ԫ)"
        .Cell(flexcpText, 0, COL_ILL��ҩƷ����, 1, COL_ILL��ҩƷ����) = "����ҩ��Ʒ����"
        
        .Cell(flexcpText, 0, COL_ILL����, 0, COL_ILL����) = "���ƽ��"
        
        .TextMatrix(1, COL_ILL����) = "����"
        .TextMatrix(1, COL_ILL��ת) = "��ת"
        .TextMatrix(1, COL_ILLδ��) = "δ��"
        .TextMatrix(1, COL_ILL����) = "����"
        .TextMatrix(1, COL_ILL����) = "����"
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
    End With
End Sub

Private Sub InitVS�������(ByRef vsgInfo1 As VSFlexGrid, ByRef vsgInfo2 As VSFlexGrid)
'���ܣ���ʼ����������ͷ���ϱ����ŵ����ﴦ�����ϱ����ݺ�ͳ�����ݹ���
'������ vsgInfo2 ��ϸ��� �ϱ����ݣ���vsMZYY   ͳ�� ���� vsCountDruUse��   vsgInfo2 �������  �ϱ����ݣ���  vsCF   ͳ�� ����  vsCountCF
    Dim i As Integer
 
    With vsgInfo1
        .Clear
        .FixedRows = 2: .FixedCols = 0
        .Rows = 3: .Cols = 24
        .MergeCells = flexMergeFixedOnly
        .MergeRow(0) = True
        For i = COL_CF��� To COL_CF��ҩ���
            .ColAlignment(i) = flexAlignLeftCenter '��ͳһ�������
            .MergeCol(i) = True
        Next
        .ColWidth(COL_CF���) = 450
        .ColAlignment(COL_CF���) = flexAlignCenterCenter
        
        .ColWidth(COL_CF�����) = 1110
        .ColWidth(COL_CF��������) = 1050
        .ColWidth(COL_CF��������) = 1080
        .ColAlignment(COL_CF��������) = flexAlignCenterCenter
        
        .ColWidth(COL_CF����ҽ��) = 900
        .ColWidth(COL_CF����) = 1400
        .ColWidth(COL_CF��������) = 540
        .ColWidth(COL_CF���) = 2210
        .ColWidth(COL_CFҩƷƷ����) = 480
        .ColAlignment(COL_CFҩƷƷ����) = flexAlignRightCenter
        .ColWidth(COL_CF��ҩƷ����) = 480
        .ColAlignment(COL_CF��ҩƷ����) = flexAlignRightCenter
        .ColWidth(COL_CFע���) = 630
        .ColAlignment(COL_CFע���) = flexAlignCenterCenter
        
        .ColWidth(COL_CF��ҩƷ����) = 480
        .ColAlignment(COL_CF��ҩƷ����) = flexAlignRightCenter
        
        .ColWidth(COL_CFͨ����) = 2310
        .ColWidth(COL_CF���) = 2800
        .ColWidth(COL_CF����) = 615
        .ColAlignment(COL_CF����) = flexAlignRightCenter
        
        .ColWidth(COL_CF���) = 870
        .ColAlignment(COL_CF���) = flexAlignRightCenter
        .ColWidth(COL_CF�÷�����) = 1260
        .ColWidth(COL_CF��ҩ;��) = 1020
        .ColWidth(COL_CF�������) = 780
        .ColAlignment(COL_CF�������) = flexAlignRightCenter
        .ColWidth(COL_CFҩƷ���) = 780
        .ColAlignment(COL_CFҩƷ���) = flexAlignRightCenter
        .ColWidth(COL_CF��ҩ���) = 840
        .ColAlignment(COL_CF��ҩ���) = flexAlignRightCenter
        
        .ColHidden(COL_CF����ID) = True
        .ColHidden(COL_CF�Һ�ID) = True
        .ColHidden(COL_CF�Һŵ�) = True
        
        .Cell(flexcpText, 0, COL_CF���, 1, COL_CF���) = "���"
        .Cell(flexcpText, 0, COL_CF�����, 1, COL_CF�����) = "�����"
        .Cell(flexcpText, 0, COL_CF��������, 1, COL_CF��������) = "��������"
        .Cell(flexcpText, 0, COL_CF��������, 1, COL_CF��������) = "��������"
        .Cell(flexcpText, 0, COL_CF����ҽ��, 1, COL_CF����ҽ��) = "����ҽ��"
        .Cell(flexcpText, 0, COL_CF����, 1, COL_CF����) = "����"
        .Cell(flexcpText, 0, COL_CF��������, 1, COL_CF��������) = "����"
        .Cell(flexcpText, 0, COL_CF���, 1, COL_CF���) = "���"
        .Cell(flexcpText, 0, COL_CFҩƷƷ����, 1, COL_CFҩƷƷ����) = "ҩƷ����"
        .Cell(flexcpText, 0, COL_CF��ҩƷ����, 1, COL_CF��ҩƷ����) = "����ҩ����"
        .Cell(flexcpText, 0, COL_CFע���, 1, COL_CFע���) = "ע�����/��"
        .Cell(flexcpText, 0, COL_CF��ҩƷ����, 1, COL_CF��ҩƷ����) = "����ҩ����"
        
        .Cell(flexcpText, 0, COL_CFͨ����, 0, COL_CF��ҩ;��) = "����ҩ��ʹ�����(�����÷�)"
    
        .TextMatrix(1, COL_CFͨ����) = "ͨ����"
        .TextMatrix(1, COL_CF���) = "���"
        .TextMatrix(1, COL_CF����) = "����"
        .TextMatrix(1, COL_CF���) = "���(Ԫ)"
        .TextMatrix(1, COL_CF�÷�����) = "�÷�����"
        .TextMatrix(1, COL_CF��ҩ;��) = "��ҩ;��"
        .Cell(flexcpText, 0, COL_CF�������, 1, COL_CF�������) = "�������(Ԫ)"
        .Cell(flexcpText, 0, COL_CFҩƷ���, 1, COL_CFҩƷ���) = "ҩƷ���(Ԫ)"
        .Cell(flexcpText, 0, COL_CF��ҩ���, 1, COL_CF��ҩ���) = "����ҩ����(Ԫ)"
        
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, COL_CF��ҩ���) = flexAlignCenterCenter
        
        .WordWrap = True
    End With
    
    With vsgInfo2
        .Clear
        .Cols = 4: .Rows = 0
        .RowHeightMin = 300
        For i = 0 To .Cols - 1
            .ColWidth(i) = 5000
        Next
        
        .AddItem "" '��һ��
        .TextMatrix(.Rows - 1, 0) = "A(������ҩ��Ʒ����)=0��"
        .TextMatrix(.Rows - 1, 1) = "B(ƽ����ҩƷ����A/������)=0��"
        .TextMatrix(.Rows - 1, 2) = "C(ʹ�ÿ���ҩ���Ʒ����)=0��"
        .TextMatrix(.Rows - 1, 3) = "D(����ʹ�ÿ���ҩ��İٷ���C/A*100%)=0%"
        
        .AddItem "" '�ڶ���
        .TextMatrix(.Rows - 1, 0) = "E(ʹ��ע����Ĵ�����)=0��"
        .TextMatrix(.Rows - 1, 1) = "F(����ʹ��ע����Ĵ����İٷ���E/������*100%)=0%"
        .TextMatrix(.Rows - 1, 2) = "G(ʹ�ÿ���ҩ��Ĵ�����)=0��"
        .TextMatrix(.Rows - 1, 3) = "H(����ʹ�ÿ���ҩ��Ĵ����İٷ��� G/100)=0%"
                
        .AddItem "" '������
        .TextMatrix(.Rows - 1, 0) = "I(�����ܽ��)=0Ԫ"
        .TextMatrix(.Rows - 1, 1) = "J(����ƽ����� I/100)=0Ԫ"
        .TextMatrix(.Rows - 1, 2) = "K(ʹ�ÿ���ҩ����ܽ��)=0Ԫ"
        .TextMatrix(.Rows - 1, 3) = "L(����ҩ���ܽ��ռ�����ܽ��ı��� K/I)=0%"
        
        .AddItem "" '������
        .TextMatrix(.Rows - 1, 0) = "M(ʹ�ÿ���ҩ��Ĵ����ܽ��)=0Ԫ"
        .TextMatrix(.Rows - 1, 1) = "N(ÿ�ſ���ҩ����ƽ����� M/G)=0Ԫ"
        .TextMatrix(.Rows - 1, 2) = "O(ʹ�û���ҩ���Ʒ����)=0��"
        .TextMatrix(.Rows - 1, 3) = "P(����ʹ�û���ҩ��İٷ��� O/A)=0%"
        
        .AddItem "" '������
        .TextMatrix(.Rows - 1, 0) = "Q(ʹ�ÿ���ҩ��Ĵ�������)=0��"
        .TextMatrix(.Rows - 1, 1) = "R(����ʹ�ÿ���ҩ�ﴦ���İٷ��� Q/100)=0%"
    End With
End Sub

Private Sub Load��������(ByRef vsgInfo1 As VSFlexGrid, ByRef vsgInfo2 As VSFlexGrid, ByVal blnRP As Boolean)
'���ܣ����ܴ���ͳ�Ʒ�������
'������ vsgInfo1 �����󴦷��������ϱ����ݣ���vsMZYY  ����ͳ�� vsCountDruUse ��
'       vsgInfo2 ��������                      vsCF             vsCountCF
'       blnRP ���棬true �ϱ����ݣ�false ����ͳ��
    Dim lng�������� As Long
    Dim lng����ҩ���� As Long
    Dim lngע������� As Long
    Dim lngʹ�ÿ���ҩ������ As Long
    Dim dbl�����ܽ�� As Double
    Dim dblʹ�ÿ���ҩ��� As Double
    Dim dblʹ�ÿ���ҩ������� As Double
    Dim lng����ҩ���� As Long
    
    
    Dim strDec As String
    Dim lngʵ������ As Long
    Dim lng��ҩƷ�� As Long
    Dim strTmp As String
    Dim dblTmp As Double
    Dim lngTmp As Long
    
    Dim i As Long
    
    strDec = "0.00"
    
    With vsgInfo1
        'ȡʵ�ʴ�����������100�ŵ�ʵ�ʿ��ܲ���100�ţ�
        For i = .Rows - 1 To .FixedRows Step -1
            If Val(.TextMatrix(i, COL_CF���)) <> 0 Then
                lngʵ������ = Val(.TextMatrix(i, COL_CF���))
                Exit For
            End If
        Next
        
        If lngʵ������ = 0 Then Exit Sub
        
        For i = .FixedRows To .Rows - 1
            'A(������ҩ��Ʒ����)  ÿ�������õ�ҩƷ�����
            lng��ҩƷ�� = lng��ҩƷ�� + Val(.TextMatrix(i, COL_CFҩƷƷ����))
            
            'C(ʹ�ÿ���ҩ���Ʒ����)  ÿ�������õĿ���ҩҩƷ�����
            lng����ҩ���� = lng����ҩ���� + Val(.TextMatrix(i, COL_CF��ҩƷ����))
            
            If .TextMatrix(i, COL_CFע���) = "��" Then lngע������� = lngע������� + 1
            
            If Val(.TextMatrix(i, COL_CF��ҩƷ����)) <> 0 Then
                lngʹ�ÿ���ҩ������ = lngʹ�ÿ���ҩ������ + 1
                dblʹ�ÿ���ҩ������� = dblʹ�ÿ���ҩ������� + Val(.TextMatrix(i, COL_CF�������))
            End If
            
            dbl�����ܽ�� = dbl�����ܽ�� + Val(.TextMatrix(i, COL_CF�������))
            dblʹ�ÿ���ҩ��� = dblʹ�ÿ���ҩ��� + Val(.TextMatrix(i, COL_CF��ҩ���))
            lng����ҩ���� = lng����ҩ���� + Val(.TextMatrix(i, COL_CF��ҩƷ����))
        Next
    End With
    
    With vsgInfo2
'        .AddItem "" '��һ��
        .TextMatrix(lngTmp, 0) = "A(������ҩ��Ʒ����)=" & lng��ҩƷ�� & "��"
        
        dblTmp = lng��ҩƷ�� / lngʵ������
        .TextMatrix(lngTmp, 1) = "B(ƽ����ҩƷ���� A/" & lngʵ������ & ")=" & Format(dblTmp, strDec) & "��": dblTmp = 0
        
        .TextMatrix(lngTmp, 2) = "C(ʹ�ÿ���ҩ���Ʒ����)=" & lng����ҩ���� & "��"
    
        If lng��ҩƷ�� <> 0 Then dblTmp = lng����ҩ���� * 100 / lng��ҩƷ��
        .TextMatrix(lngTmp, 3) = "D(����ʹ�ÿ���ҩ��İٷ��� C/A)=" & Format(dblTmp, strDec) & "%": dblTmp = 0
        
'        .AddItem "" '�ڶ���
        lngTmp = lngTmp + 1
        .TextMatrix(lngTmp, 0) = "E(ʹ��ע����Ĵ�����)=" & lngע������� & "��"
        
        dblTmp = lngע������� * 100 / lngʵ������
        .TextMatrix(lngTmp, 1) = "F(����ʹ��ע����Ĵ����İٷ��� E/" & lngʵ������ & ")=" & Format(dblTmp, strDec) & "%": dblTmp = 0
        .TextMatrix(lngTmp, 2) = "G(ʹ�ÿ���ҩ��Ĵ�����)=" & lngʹ�ÿ���ҩ������ & "��"
        
        dblTmp = lngʹ�ÿ���ҩ������ * 100 / lngʵ������
        .TextMatrix(lngTmp, 3) = "H(����ʹ�ÿ���ҩ��Ĵ����İٷ��� G/" & lngʵ������ & ")=" & Format(dblTmp, strDec) & "%": dblTmp = 0
                
'        .AddItem "" '������
        lngTmp = lngTmp + 1
        .TextMatrix(lngTmp, 0) = "I(�����ܽ��)=" & Format(dbl�����ܽ��, strDec) & "Ԫ"
        
        dblTmp = dbl�����ܽ�� / lngʵ������
        .TextMatrix(lngTmp, 1) = "J(����ƽ����� I/" & lngʵ������ & ")=" & Format(dblTmp, strDec) & "Ԫ": dblTmp = 0
        
        .TextMatrix(lngTmp, 2) = "K(ʹ�ÿ���ҩ����ܽ��)=" & Format(dblʹ�ÿ���ҩ���, strDec) & "Ԫ"
        
        If dbl�����ܽ�� <> 0 Then dblTmp = dblʹ�ÿ���ҩ��� / dbl�����ܽ��
        
        .TextMatrix(lngTmp, 3) = "L(����ҩ���ܽ��ռ�����ܽ��ı��� K/I)=" & Format(dblTmp, strDec): dblTmp = 0
        
'        .AddItem "" '������
        lngTmp = lngTmp + 1
        .TextMatrix(lngTmp, 0) = "M(ʹ�ÿ���ҩ��Ĵ����ܽ��)=" & Format(dblʹ�ÿ���ҩ�������, strDec) & "Ԫ"
        
        If lngʹ�ÿ���ҩ������ <> 0 Then dblTmp = dblʹ�ÿ���ҩ������� / lngʹ�ÿ���ҩ������
        .TextMatrix(lngTmp, 1) = "N(ÿ�ſ���ҩ����ƽ����� M/G)=" & Format(dblTmp, strDec) & "Ԫ": dblTmp = 0
        
        .TextMatrix(lngTmp, 2) = "O(ʹ�û���ҩ���Ʒ����)=" & lng����ҩ���� & "��"
        
        If lng��ҩƷ�� <> 0 Then dblTmp = lng����ҩ���� * 100 / lng��ҩƷ��
        .TextMatrix(lngTmp, 3) = "P(����ʹ�û���ҩ��İٷ��� O/A * 100%)=" & Format(dblTmp, strDec) & "%": dblTmp = 0
        
'        .AddItem "" '������
        lngTmp = lngTmp + 1
        .TextMatrix(lngTmp, 0) = "Q(ʹ�ÿ���ҩ��Ĵ�������)=" & lngʹ�ÿ���ҩ������ & "��"
        
        dblTmp = 100 * lngʹ�ÿ���ҩ������ / lngʵ������
        .TextMatrix(lngTmp, 1) = "R(����ʹ�ÿ���ҩ�ﴦ���İٷ��� Q/100)=" & Format(dblTmp, strDec) & "%": dblTmp = 0
    End With

End Sub

Private Sub LoadvsUseRan()
'���ܣ�����   ����ҩ��ʹ���������ͳ��  ������
    Dim strSql As String, strPar1 As String, strPar2 As String
    Dim strSQLDetail As String
    Dim rsDetail As ADODB.Recordset
    Dim strDept As String
    Dim rsTmp As ADODB.Recordset
    Dim rsTmpOther As ADODB.Recordset
    Dim strDec As String
    Dim strTmp As String
    Dim lngTmp As Long
    Dim dblTmp As Double
    Dim lng���������� As Long
    Dim lng�������˴� As Long
    Dim dbl��ҩ�� As Double
    Dim i As Long
    Dim rs��Ժ���� As ADODB.Recordset
    Dim strSQL��Ժ���� As String
    Dim dat��Ժ��ʼ As Date
    Dim dat��Ժ���� As Date
    Dim strWhere���� As String
    Dim lng���� As Long
    Dim lng���� As Long

    On Error GoTo errH
    Screen.MousePointer = 11
    Call zlCommFun.ShowFlash("���ڶ�ȡ���ݣ����Ժ�...")
    
    strDec = "0.00"
    
    '���ڷ�Χ��������
    strPar1 = "To_Date('" & Format(dtpCountS(e_C0_dtpCountS_��ʼʱ��_0).Value, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')"
    strPar2 = "To_Date('" & Format(dtpCountE(e_C0_dtpCountE_����ʱ��_0).Value, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
    strDept = IIf(txtDept(e_C0_txtDept_��������_0).Tag = "", "", " and a.��������id in (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))")
    
    If optType(e_C0_optType_ͳ�Ƴ���_סԺ_5).Value Then
        'ͳ�Ƴ��� סԺ
        If optType(e_C0_optType_���ܷ�ʽ_����_9).Value Then '���ܷ�ʽ ����
            strSql = "Select /*+ rule*/  m.����,m.��������id, m.Ddds, m.����, m.���� as ������, round(m.����) as ����" & vbNewLine & _
                "From (Select n.����,m.��������id, Sum(m.Ddds) As Ddds, Sum(m.����) As ����, Sum(m.����) As ����, Count(1) As ����" & vbNewLine & _
                "       From (Select m.��������id, m.����id, m.��ҳid, Sum(m.Ddds) As Ddds, Sum(m.����) As ����, Nvl(g.��Ժ����,Sysdate) -g.��Ժ���� As ����" & vbNewLine & _
                "              From (Select a.��������id,a.����id, a.��ҳid, Sum(a.���ʽ��) As ����," & vbNewLine & _
                "                            Sum(Decode(Nvl(b.Dddֵ, 0), 0, 0, a.���� * b.����ϵ�� / b.Dddֵ)) As Ddds" & vbNewLine & _
                "                     From סԺ���ü�¼ A, ҩƷ��� B,ҩƷ���� D" & vbNewLine & _
                "                     Where a.�շ���� = '5' And a.��¼״̬ <> 0 And  a.����ʱ�� Between " & strPar1 & " And " & strPar2 & strDept & vbNewLine & _
                "                           And a.�շ�ϸĿid = b.ҩƷid And b.ҩ��id = d.ҩ��id And Nvl(d.������, 0) <> 0" & vbNewLine & _
                "                     Group By a.����id, a.��ҳid,a.��������id) M, ������ҳ G" & vbNewLine & _
                "          where  g.����id = m.����id And g.��ҳid = m.��ҳid  Group By m.����id, m.��ҳid, m.��������id,g.��Ժ����,g.��Ժ����) M,���ű� n" & vbNewLine & _
                "       where m.��������id=n.id Group By m.��������id,n.���� having Sum(m.����) >0 " & vbNewLine & _
                "       Order By Sum(m.����) Desc) M" & vbNewLine & _
                "Where Rownum <" & Val(txtTopRan.Text) + 1
            strSQLDetail = "select a.��������id as id,sum(a.���ʽ��) as �ܷ��� from סԺ���ü�¼ a" & vbNewLine & _
                "where a.�շ���� in ('5','6','7') and a.��¼״̬ <> 0 " & vbNewLine & _
                "and a.����ʱ�� between " & strPar1 & " And " & strPar2 & vbNewLine & _
                "and a.��������id in (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))" & vbNewLine & _
                "group by a.��������id having sum(a.���ʽ��)>0"
        ElseIf optType(e_C0_optType_���ܷ�ʽ_ҽ��_8).Value Then
            '���ܷ�ʽ ҽ��
            strSql = "Select /*+ rule*/  nvl(m.������,'��') as ������, m.Ddds, m.����, m.���� as ������,round(m.����) as ����" & vbNewLine & _
                "From (Select m.������, Sum(m.Ddds) As Ddds, Sum(m.����) As ����, Count(1) As ����, Sum(m.����) As ����" & vbNewLine & _
                "       From (Select m.������, m.����id, m.��ҳid, Sum(m.Ddds) As Ddds, Sum(m.����) As ����,Nvl(g.��Ժ����,Sysdate) -g.��Ժ���� As ����" & vbNewLine & _
                "              From (Select a.������,a.����id, a.��ҳid," & vbNewLine & _
                "                   Sum(Decode(Nvl(b.Dddֵ, 0), 0, 0, a.���� * b.����ϵ�� / b.Dddֵ)) As Ddds, Sum(a.���ʽ��) As ����" & vbNewLine & _
                "                     From סԺ���ü�¼ A, ҩƷ��� B,ҩƷ���� D" & vbNewLine & _
                "                     Where a.�շ���� = '5' And a.��¼״̬ <> 0 And a.����ʱ�� Between " & strPar1 & " And " & strPar2 & strDept & vbNewLine & _
                "                           And a.�շ�ϸĿid = b.ҩƷid And b.ҩ��id = d.ҩ��id and  Nvl(d.������, 0) <> 0" & vbNewLine & _
                "                     Group By a.����id, a.��ҳid,a.������) M, ������ҳ G" & vbNewLine & _
                "   where  g.����id = m.����id And g.��ҳid = m.��ҳid   Group By m.����id, m.��ҳid, m.������,g.��Ժ����,g.��Ժ����) M" & vbNewLine & _
                "       Group By m.������ having Sum(m.����) >0 " & vbNewLine & _
                "       Order By Sum(m.����) Desc) M" & vbNewLine & _
                "Where Rownum <" & Val(txtTopRan.Text) + 1
            strSQLDetail = "Select nvl(a.������,'��') as ������, Sum(a.���ʽ��) As �ܷ��� From סԺ���ü�¼ A" & vbNewLine & _
                "Where a.�շ���� In ('5', '6', '7') And a.��¼״̬ <> 0 And a.������ is not null And" & vbNewLine & _
                " a.����ʱ�� Between " & strPar1 & " And " & strPar2 & "and instr(',[1],',','||nvl(a.������,'��')||',')>0" & vbNewLine & _
                "Group By a.������ having sum(a.���ʽ��)>0"
        Else
            strDept = IIf(txtDept(e_C0_txtDept_��������_0).Tag = "", "", " and a.��Ժ����id in (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))")
            
            strSQL��Ժ���� = "select count(1) as ����,round(Sum(m.����)) As ����,min(m.��Ժ����) as ��ʼ,max(m.��Ժ����) as ���� from" & vbNewLine & _
                "(select a.��Ժ����-a.��Ժ���� As ����,a.��Ժ����,a.��Ժ����  from ������ҳ a where a.��Ժ���� Between [2] And [3] " & strDept & " ) m"
            strPar1 = Format(dtpCountS(e_C0_dtpCountS_��ʼʱ��_0).Value, "yyyy-MM-dd 00:00:00")
            strPar2 = Format(dtpCountE(e_C0_dtpCountE_����ʱ��_0).Value, "yyyy-MM-dd 23:59:59")
            dat��Ժ��ʼ = CDate(strPar1)
            dat��Ժ���� = CDate(strPar2)
            Set rs��Ժ���� = zlDatabase.OpenSQLRecord(strSQL��Ժ����, Me.Caption, txtDept(e_C0_txtDept_��������_0).Tag, dat��Ժ��ʼ, dat��Ժ����)
            
            If Val("" & rs��Ժ����!����) = 0 Then
                Screen.MousePointer = 0
                Call zlCommFun.StopFlash
                MsgBox "��ǰ������δ�ҵ��κ����ݣ�����������ͳ�Ʋ�����", vbInformation, gstrSysName
                Exit Sub
            End If
            
            lng���� = Val("" & rs��Ժ����!����)
            lng���� = Val("" & rs��Ժ����!����)
            
            strDept = IIf(txtDept(e_C0_txtDept_��������_0).Tag = "", "", " and g.��Ժ����id in (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))")
            strWhere���� = " and g.��Ժ���� Between [2] And [3]" & strDept
            
            
            strPar1 = "To_Date('" & Format(rs��Ժ����!��ʼ, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')"
            strPar2 = "To_Date('" & Format(rs��Ժ����!����, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
            
            '���ܷ�ʽ ҩƷ
            strSql = "Select /*+ rule*/  m.���, m.ҩƷͨ����, m.����, m.���, m.Ddds, m.����, m.����," & lng���� & " As ������," & lng���� & " as ����, m.סԺ��λ As ��λ" & vbNewLine & _
                "From (Select m.�շ�ϸĿid, m.���, m.ҩƷͨ����, m.����, m.���, Sum(m.Ddds) As Ddds, Sum(m.����) As ����,Sum(m.����) As ����, m.סԺ��λ" & vbNewLine & _
                "       From (Select m.����id, m.��ҳid, m.�շ�ϸĿid, m.���, m.ҩƷͨ����, m.����, m.���, Sum(m.Ddds) As Ddds, Sum(m.����) As ����," & vbNewLine & _
                "                 Sum(m.����) As ����, m.סԺ��λ" & vbNewLine & _
                "              From (Select a.����id, a.��ҳid, a.�շ�ϸĿid, e.���� As ���, c.���� As ҩƷͨ����," & vbNewLine & _
                "                            d.ҩƷ���� As ����, c.��� || c.���� As ���," & vbNewLine & _
                "                            Sum(Decode(Nvl(b.Dddֵ, 0), 0, 0, a.���� * b.����ϵ�� / b.Dddֵ)) As Ddds, Sum(a.���ʽ��) As ����, b.סԺ��λ," & vbNewLine & _
                "                            Sum(a.����) As ����" & vbNewLine & _
                "                     From סԺ���ü�¼ A, ҩƷ��� B, �շ���ĿĿ¼ C, ҩƷ���� D, ���Ʒ���Ŀ¼ E, ������ĿĿ¼ F" & vbNewLine & _
                "                     Where a.�շ���� = '5' And a.��¼״̬ <> 0 And" & vbNewLine & _
                "                           a.����ʱ�� Between " & strPar1 & " And " & strPar2 & vbNewLine & _
                "                           And a.�շ�ϸĿid = b.ҩƷid And b.ҩƷid = c.id And d.ҩ��id = b.ҩ��id And d.ҩ��id = f.Id And f.����id = e.Id And Nvl(d.������,0)<>0" & vbNewLine & _
                "                     Group By a.�շ�ϸĿid, a.����id, a.��ҳid,e.����, c.����, d.ҩƷ����, c.���, c.����," & vbNewLine & _
                "                              b.סԺ��λ) M, ������ҳ G where g.����id = m.����id And g.��ҳid = m.��ҳid" & strWhere���� & vbNewLine & _
                "              Group By m.����id, m.��ҳid, m.�շ�ϸĿid, m.���, m.ҩƷͨ����, m.����, m.���, m.סԺ��λ,g.��Ժ����,g.��Ժ����) M" & vbNewLine & _
                "       Group By m.�շ�ϸĿid, m.���, m.ҩƷͨ����, m.����, m.���, m.סԺ��λ having Sum(m.����) >0 " & vbNewLine & _
                "       Order By Sum(m." & IIf(optType(e_C0_optType_����ʽ_����_12).Value, "����", "����") & ") desc) M" & vbNewLine & _
                "Where Rownum < " & Val(txtTopRan.Text) + 1

            strSQLDetail = "Select Sum(a.���ʽ��) As ��ҩ�� From סԺ���ü�¼ A,������ҳ g" & vbNewLine & _
                " Where a.�շ���� In ('5', '6', '7') And a.��¼״̬ <> 0 And" & vbNewLine & _
                " a.����ʱ�� Between " & strPar1 & " And " & strPar2 & " and a.����id+0 = g.����id And a.��ҳid+0= g.��ҳid" & strWhere����
        End If
    Else    'ͳ�Ƴ��� ����
        If optType(e_C0_optType_���ܷ�ʽ_����_9).Value Then
            '���ܷ�ʽ ����
            strSql = "Select /*+ rule*/  m.����, m.��������id, m.Ddds, m.����, m.������ as ������" & vbNewLine & _
                "From (Select n.����, m.��������id, Sum(m.Ddds) As Ddds, Sum(m.����) As ����, Sum(m.������) As ������" & vbNewLine & _
                "       From (Select m.��������id, Sum(m.Ddds) As Ddds, Sum(m.����) As ����, Count(1) As ������" & vbNewLine & _
                "              From (Select m.��������id, Sum(m.Ddds) As Ddds, Sum(m.����) As ����" & vbNewLine & _
                "                     From (Select a.��������id, a.No, Sum(Decode(Nvl(b.Dddֵ, 0), 0, 0, a.���� * b.����ϵ�� / b.Dddֵ)) As Ddds," & vbNewLine & _
                "                                   Sum(a.���ʽ��) As ����" & vbNewLine & _
                "                            From ������ü�¼ A, ҩƷ��� B,ҩƷ���� D" & vbNewLine & _
                "                            Where a.�շ���� = '5' And a.��¼״̬ <> 0 And a.����ʱ�� Between " & strPar1 & " And " & strPar2 & strDept & vbNewLine & _
                "                                  And a.�շ�ϸĿid = b.ҩƷid And b.ҩ��id = d.ҩ��id And Nvl(d.������, 0) <> 0" & vbNewLine & _
                "                            Group By a.No,a.��������id) M" & vbNewLine & _
                "                     Group By m.No, m.��������id) M" & vbNewLine & _
                "              Group By m.��������id) M, ���ű� N" & vbNewLine & _
                "       Where m.��������id = n.Id" & vbNewLine & _
                "       Group By m.��������id, n.���� having Sum(m.����) >0 " & vbNewLine & _
                "       Order By Sum(m.����) Desc) M" & vbNewLine & _
                "Where Rownum <" & Val(txtTopRan.Text) + 1
                
            strSQLDetail = "select a.��������id as id,sum(a.���ʽ��) as �ܷ��� from ������ü�¼ a" & vbNewLine & _
                "where a.�շ���� in ('5','6','7') and a.��¼״̬ <> 0" & vbNewLine & _
                "and a.����ʱ�� between " & strPar1 & " And " & strPar2 & vbNewLine & _
                "and a.��������id in (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))" & vbNewLine & _
                "group by a.��������id having sum(a.���ʽ��)>0"

        ElseIf optType(e_C0_optType_���ܷ�ʽ_ҽ��_8).Value Then
            '���ܷ�ʽ ҽ��
            strSql = "Select /*+ rule*/  nvl(m.������,'��') as ������, m.Ddds, m.����, m.������ as ������" & vbNewLine & _
                "From (Select m.������, Sum(m.Ddds) As Ddds, Sum(m.����) As ����, Count(1) As ������" & vbNewLine & _
                "       From (Select m.������, Sum(m.Ddds) As Ddds, Sum(m.����) As ����" & vbNewLine & _
                "              From (Select a.������, a.NO, Sum(Decode(Nvl(b.Dddֵ, 0), 0, 0, a.���� * b.����ϵ�� / b.Dddֵ)) As Ddds," & vbNewLine & _
                "                            Sum(a.���ʽ��) As ����" & vbNewLine & _
                "                     From ������ü�¼ A, ҩƷ��� B,ҩƷ���� D" & vbNewLine & _
                "                     Where a.�շ���� = '5' And a.��¼״̬ <> 0 And a.����ʱ�� Between " & strPar1 & " And " & strPar2 & strDept & vbNewLine & _
                "                           And a.�շ�ϸĿid = b.ҩƷid And b.ҩ��id = d.ҩ��id And  Nvl(d.������, 0) <> 0" & vbNewLine & _
                "                     Group By a.No,a.������) M" & vbNewLine & _
                "              Group By m.No, m.������) M" & vbNewLine & _
                "       Group By m.������  having Sum(m.����) >0" & vbNewLine & _
                "       Order By Sum(m.����) Desc) M" & vbNewLine & _
                "Where Rownum <" & Val(txtTopRan.Text) + 1
            
            strSQLDetail = "Select nvl(a.������,'��') as ������, Sum(a.���ʽ��) As �ܷ��� From ������ü�¼ A" & vbNewLine & _
                "Where a.�շ���� In ('5', '6', '7') And a.��¼״̬ <> 0 And a.������ is not null And " & vbNewLine & _
                " a.����ʱ�� Between " & strPar1 & " And " & strPar2 & "and instr(',[1],',','||nvl(a.������,'��')||',')>0" & vbNewLine & _
                "Group By a.������ having Sum(a.���ʽ��)>0"
        Else
            '���ܷ�ʽ ҩƷ
            strSql = "Select /*+ rule*/  m.���, m.ҩƷͨ����, m.����, m.���, m.Ddds, m.����, m.������ as ������, m.����, m.���ﵥλ as ��λ" & vbNewLine & _
                "From (Select m.�շ�ϸĿid, m.���, m.ҩƷͨ����, m.����, m.���,m.���ﵥλ,Sum(m.Ddds) As Ddds, Sum(m.����) As ����, Count(1) As ������, Sum(m.����) As ����" & vbNewLine & _
                "       From (Select m.�շ�ϸĿid, m.No, m.���, m.ҩƷͨ����, m.����, m.���,m.���ﵥλ, Sum(m.Ddds) As Ddds, Sum(m.����) As ����, Sum(m.����) As ����" & vbNewLine & _
                "              From (Select a.No || '' As NO, a.�շ�ϸĿid, e.���� As ���, c.���� As ҩƷͨ����, d.ҩƷ���� As ����, c.��� || c.���� As ���," & vbNewLine & _
                "                            Sum(Decode(Nvl(b.Dddֵ, 0), 0, 0, a.���� * b.����ϵ�� / b.Dddֵ)) As Ddds, Sum(a.���ʽ��) As ����," & vbNewLine & _
                "                            Sum(a.����) As ����, b.���ﵥλ" & vbNewLine & _
                "                     From ������ü�¼ A, ҩƷ��� B, �շ���ĿĿ¼ C, ҩƷ���� D, ���Ʒ���Ŀ¼ E, ������ĿĿ¼ F" & vbNewLine & _
                "                     Where a.�շ���� = '5' And a.��¼״̬ <> 0 And  a.����ʱ�� Between " & strPar1 & " And " & strPar2 & strDept & vbNewLine & _
                "                            And a.�շ�ϸĿid = b.ҩƷid And b.ҩƷid = c.id And" & vbNewLine & _
                "                           d.ҩ��id = b.ҩ��id And d.ҩ��id = f.Id And f.����id = e.Id And Nvl(d.������, 0) <> 0" & vbNewLine & _
                "                     Group By a.�շ�ϸĿid, a.No, e.����, c.����, d.ҩƷ����, c.���, c.����,b.���ﵥλ) M" & vbNewLine & _
                "       Group By m.�շ�ϸĿid,m.No, m.���, m.ҩƷͨ����, m.����, m.���, m.���ﵥλ having Sum(m.����) >0) M" & vbNewLine & _
                "       group by  m.�շ�ϸĿid, m.���, m.ҩƷͨ����, m.����, m.���, m.���ﵥλ" & vbNewLine & _
                "       Order By Sum(m." & IIf(optType(e_C0_optType_����ʽ_����_12).Value, "����", "����") & ") Desc) M" & vbNewLine & _
                "Where Rownum < " & Val(txtTopRan.Text) + 1
            strSQLDetail = "Select Sum(a.���ʽ��) As ��ҩ�� From ������ü�¼ A" & vbNewLine & _
                " Where a.�շ���� In ('5', '6', '7') And a.��¼״̬ <> 0 And" & vbNewLine & _
                " a.����ʱ�� Between " & strPar1 & " And " & strPar2 & strDept
        End If
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, txtDept(e_C0_txtDept_��������_0).Tag, dat��Ժ��ʼ, dat��Ժ����)
    
    '���Ҫ���³�ʼ�� �� ���ݼ���
    vsUseRan.Rows = vsUseRan.FixedRows
    vsUseRan.Rows = vsUseRan.FixedRows + 1
    
    If rsTmp.EOF Then
        Screen.MousePointer = 0
        Call zlCommFun.StopFlash
        MsgBox "��ǰ������δ�ҵ��κ����ݣ�����������ͳ�Ʋ�����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If optType(e_C0_optType_���ܷ�ʽ_����_9).Value Or optType(e_C0_optType_���ܷ�ʽ_ҽ��_8).Value Then '  ���һ�ҽ��
        strTmp = IIf(optType(e_C0_optType_���ܷ�ʽ_ҽ��_8).Value, "ҽ������", "��������") & ",2000,1;�ܽ��(Ԫ),1510,7;ʹ������(����),1440,7;����������,1080,7;ÿ��ƽ�����(Ԫ),1720,7;DDDs,1000,7;ʹ��ǿ��,1000,7;ռҩƷ������(%),2100,7"
        Call InitTable(vsUseRan, strTmp)
        vsUseRan.RowHeight(0) = 600
        If rsTmp.RecordCount > 0 Then
            strTmp = ""
            For i = 1 To rsTmp.RecordCount
                If optType(e_C0_optType_���ܷ�ʽ_����_9).Value Then '�����һ���
                    strTmp = strTmp & "," & rsTmp!��������id
                Else
                    strTmp = strTmp & ",'" & rsTmp!������ & ","
                End If
                rsTmp.MoveNext
            Next
            strTmp = Mid(strTmp, 2)
            rsTmp.MoveFirst
            
            Set rsDetail = zlDatabase.OpenSQLRecord(strSQLDetail, Me.Caption, strTmp)
            
            With vsUseRan
                .Rows = rsTmp.RecordCount + 1
                For i = 1 To rsTmp.RecordCount
                    If optType(e_C0_optType_���ܷ�ʽ_ҽ��_8).Value Then
                        .TextMatrix(i, COL_D����) = rsTmp!������ & "" 'ҽ������
                    Else
                        .TextMatrix(i, COL_D����) = rsTmp!���� & "" '��������
                    End If
                    
                    .TextMatrix(i, COL_D�ܽ��) = Format(Val(rsTmp!���� & ""), strDec)       '�ܽ��(Ԫ) ����ҩ��
                    .TextMatrix(i, COL_Dʹ������) = Val(rsTmp!������ & "") ' ʹ������(����)
                    .TextMatrix(i, COL_DDDDs) = Format(Val(rsTmp!Ddds & ""), strDec)
                    
                    If optType(e_C0_optType_ͳ�Ƴ���_סԺ_5).Value Then
                        .TextMatrix(i, COL_D����������) = Val(rsTmp!���� & "")
                    End If
                    
                    dblTmp = 0
                    If Val(rsTmp!������ & "") <> 0 Then dblTmp = Val(rsTmp!���� & "") / Val(rsTmp!������ & "")
                    .TextMatrix(i, COL_Dÿ��ƽ�����) = Format(dblTmp, strDec)   'ÿ��ƽ�����(Ԫ)
                    
                    dblTmp = 0
                    If optType(e_C0_optType_ͳ�Ƴ���_סԺ_5).Value Then '�����סԺ��ʹ��ǿ�ȼ��㷽ʽ��ͬ  optN02--סԺ
                        If Val(rsTmp!���� & "") <> 0 Then dblTmp = Val(rsTmp!Ddds & "") * 100 / Val(rsTmp!���� & "")
                    Else
                        If Val(rsTmp!������ & "") <> 0 Then dblTmp = Val(rsTmp!Ddds & "") * 100 / Val(rsTmp!������ & "")
                    End If
                    .TextMatrix(i, COL_Dʹ��ǿ��) = Format(dblTmp, strDec) 'ʹ��ǿ��
                    
                    dblTmp = 0: dbl��ҩ�� = 0: rsDetail.Filter = 0
                    
                    If optType(e_C0_optType_���ܷ�ʽ_ҽ��_8).Value Then
                        rsDetail.Filter = "������='" & rsTmp!������ & "'"
                    Else
                        rsDetail.Filter = "id=" & rsTmp!��������id
                    End If
                    
                    If Not rsDetail.EOF Then dbl��ҩ�� = Val(rsDetail!�ܷ��� & "")
                    If dbl��ҩ�� <> 0 Then dblTmp = Val(rsTmp!���� & "") * 100 / dbl��ҩ��
                    .TextMatrix(i, COL_DռҩƷ������) = Format(dblTmp, strDec) & "%"  'ռҩƷ������(%)
                    
                    rsTmp.MoveNext
                Next
                .ColHidden(COL_D����������) = Not optType(e_C0_optType_ͳ�Ƴ���_סԺ_5).Value '��ͳ�Ƴ�����סԺʱ����ʾ  ����������
            End With
        End If
    Else
        '���ܷ�ʽ ҩƷ
        strTmp = "���,1480,4;ҩƷ����,2480,1;����,1530,4;���,2800,1;����,1020,7;�ܽ��(Ԫ),1000,7;ʹ������,530,7;����������,760,7;ÿ��ƽ�����(Ԫ),930,7;DDDs,750,7;ʹ��ǿ��,750,7;ռҩƷ�ܽ�����(%),1000,7"
        Call InitTable(vsUseRan, strTmp)
        vsUseRan.RowHeight(0) = 600
        If rsTmp.RecordCount > 0 Then
            Set rsDetail = zlDatabase.OpenSQLRecord(strSQLDetail, Me.Caption, txtDept(0).Tag, dat��Ժ��ʼ, dat��Ժ����)
            If Not rsDetail.EOF Then dbl��ҩ�� = Val(rsDetail!��ҩ�� & "")
            With vsUseRan
                .Rows = rsTmp.RecordCount + 1
                
                For i = 1 To rsTmp.RecordCount
                    .TextMatrix(i, COL_UD���) = rsTmp!��� & ""
                    .TextMatrix(i, COL_UDҩƷ����) = rsTmp!ҩƷͨ���� & ""
                    .TextMatrix(i, COL_UD����) = rsTmp!���� & ""
                    .TextMatrix(i, COL_UD���) = rsTmp!��� & ""
                    .TextMatrix(i, COL_UD����) = Val(rsTmp!���� & "") & rsTmp!��λ '����
                    .TextMatrix(i, COL_UD�ܽ��) = Format(rsTmp!���� & "", strDec)
                    .TextMatrix(i, COL_UDʹ������) = rsTmp!������ & ""
                    .TextMatrix(i, COL_UDDDDs) = Format(rsTmp!Ddds & "", strDec)
                    
                    '��ͳ�Ƴ�����סԺʱ����ʾ  ����������  ��
                    If optType(e_C0_optType_ͳ�Ƴ���_סԺ_5).Value Then .TextMatrix(i, COL_UD����������) = rsTmp!���� & ""
                    
                    dblTmp = 0
                    If Val(rsTmp!������ & "") <> 0 Then dblTmp = Val(rsTmp!���� & "") / Val(rsTmp!������ & "")
                    .TextMatrix(i, COL_UDÿ��ƽ�����) = Format(dblTmp, strDec) 'ƽ�����
                    
                    dblTmp = 0 '�����סԺ��ʹ��ǿ�ȼ��㷽ʽ��ͬ����������Ϊ��ÿ����һ���ͬ�� ���� �� ������
                    If optType(e_C0_optType_ͳ�Ƴ���_סԺ_5).Value Then
                        If Val(rsTmp!���� & "") <> 0 Then dblTmp = Val(rsTmp!Ddds & "") * 100 / Val(rsTmp!���� & "")
                    Else
                        If Val(rsTmp!������ & "") <> 0 Then dblTmp = Val(rsTmp!Ddds & "") * 100 / Val(rsTmp!������ & "")
                    End If
                    .TextMatrix(i, COL_UDʹ��ǿ��) = Format(dblTmp, strDec) 'ʹ��ǿ��
                    
                    dblTmp = 0
                    If dbl��ҩ�� <> 0 Then dblTmp = Val(rsTmp!���� & "") * 100 / dbl��ҩ��
                    .TextMatrix(i, COL_UDռҩƷ�ܽ�����) = Format(dblTmp, strDec) & "%" 'ռҩƷ�ܽ�����
                    
                    rsTmp.MoveNext
                Next
                .ColHidden(COL_UD���) = False '�������ʾ
                .ColHidden(COL_UD����������) = optType(e_C0_optType_ͳ�Ƴ���_����_4).Value   '��ͳ�Ƴ�����סԺʱ����ʾ  ����������  ��
            End With
        End If
    End If
    
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\���ڷ�Χ", "����ҩ��ʹ���������ͳ��", dtpCountS(e_C0_dtpCountS_��ʼʱ��_0).Value & "," & dtpCountE(e_C0_dtpCountE_����ʱ��_0).Value
    
    Screen.MousePointer = 0
    Call zlCommFun.StopFlash
    Exit Sub
errH:
    Screen.MousePointer = 0
    Call zlCommFun.StopFlash
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadvsCut()
'���ܣ�����  �����п�Χ����Ԥ����ҩͳ�� ����������
    Dim strSql As String, strPar As String
    Dim rs���� As ADODB.Recordset
    Dim rs��ҩ���� As ADODB.Recordset
    Dim rsһ�п����� As ADODB.Recordset
    Dim rs��ǰ��ҩ���� As ADODB.Recordset
    Dim rs�������� As ADODB.Recordset
    Dim rs���� As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim lng�п�Ԥ������ As Long
    Dim lng������ As Long
    Dim lngҩƷ�� As Long
    Dim lng��ǰ�� As Long
    Dim lng���� As Long
    Dim lngTmp As Long
    Dim dblTmp As Double
    Dim lngRow As Long
    Dim strPatis As String
    Dim arrTmp As Variant
    
    Dim rsҩƷ As ADODB.Recordset
    Dim rs�пڿ���ҩ As ADODB.Recordset
 
    Dim strWhere As String
    Dim strTmp As String
    Dim str���� As String
 
    Dim int��ǰ As Integer
    Dim i As Long, j As Long, k As Long, m As Long
    
    '��������
    strPar = "To_Date('" & Format(dtpCountS(e_C1_dtpCountS_��ʼʱ��_1).Value, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')" & _
        " and To_Date('" & Format(dtpCountE(e_C1_dtpCountS_��ʼʱ��_1).Value, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')" & _
        IIf(txtDept(e_C1_txtDept_ͳ�ƿ���_4).Tag = "", "", " and a.��Ժ����id in (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))")
 
    On Error GoTo errH
    Screen.MousePointer = 11
    Call zlCommFun.ShowFlash("���ڶ�ȡ���ݣ����Ժ�...")
    
    vsCut.Rows = vsCut.FixedRows
    vsCut.Rows = vsCut.FixedRows + 1
    vsCut.Cell(flexcpText, 0, COL_CUT����, 1, COL_CUT����) = IIf(optType(e_C1_optType_���ܷ�ʽ_����_17).Value, "��������", "סԺҽʦ")
    
    
    '����Ժ���˽��г�������ȡһ����Ժ��������ʱ��Ϳ��Ҽ���
    '������
    If optType(e_C1_optType_���ܷ�ʽ_����_17).Value Then '���һ���
        strSql = "Select a.��Ժ����id as ����id,b.����,Count(1) As ���� From ������ҳ A, ���ű� B Where a.��Ժ���� Between " & strPar & " And a.��Ժ����id = b.Id Group By a.��Ժ����id,b.����"
    Else   'ҽ������
        strSql = "Select 0 as ����id,Nvl(a.סԺҽʦ,'��') As ����, Count(1) As ���� From ������ҳ A Where a.��Ժ���� Between " & strPar & " Group By Nvl(a.סԺҽʦ,'��')"
    End If
    Set rs���� = zlDatabase.OpenSQLRecord(strSql, Me.Caption, txtDept(e_C1_txtDept_ͳ�ƿ���_4).Tag)
    
    If rs����.EOF Then
        Screen.MousePointer = 0
        Call zlCommFun.StopFlash
        MsgBox "��ǰ������δ�ҵ��κ����ݣ�����������ͳ�Ʋ�����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '����ҩ����
    If optType(e_C1_optType_���ܷ�ʽ_����_17).Value Then '���һ���
        strSql = "Select a.��Ժ����id as ����id, Count(1) As ����" & vbNewLine & _
            " From ������ҳ A Where a.��Ժ���� Between " & strPar & " And Exists" & vbNewLine & _
            " (Select 1 From סԺ���ü�¼ B, ҩƷ��� C, ҩƷ���� D" & vbNewLine & _
            " Where a.����id = b.����id And a.��ҳid = b.��ҳid And b.��¼״̬ <> 0 And b.�շ���� = '5' And b.�շ�ϸĿid = c.ҩƷid And" & vbNewLine & _
            " c.ҩ��id = d.ҩ��id And Nvl(d.������, 0) <> 0) Group By a.��Ժ����id"

    Else   'ҽ������
        strSql = "Select Nvl(a.סԺҽʦ,'��') as ����, Count(1) As ����" & vbNewLine & _
            " From ������ҳ A Where a.��Ժ���� Between " & strPar & " And Exists" & vbNewLine & _
            " (Select 1 From סԺ���ü�¼ B, ҩƷ��� C, ҩƷ���� D" & vbNewLine & _
            " Where a.����id = b.����id And a.��ҳid = b.��ҳid And b.��¼״̬ <> 0 And b.�շ���� = '5' And b.�շ�ϸĿid = c.ҩƷid And" & vbNewLine & _
            " c.ҩ��id = d.ҩ��id And Nvl(d.������, 0) <> 0) Group By Nvl(a.סԺҽʦ,'��')"
    End If
    Set rs��ҩ���� = zlDatabase.OpenSQLRecord(strSql, Me.Caption, txtDept(e_C1_txtDept_ͳ�ƿ���_4).Tag)
    
    'һ���п�����
    If optType(e_C1_optType_���ܷ�ʽ_����_17).Value Then '���һ���
        strSql = "Select a.��Ժ����id as ����id, Count(1) As ����" & vbNewLine & _
            " From ������ҳ A Where a.��Ժ���� Between " & strPar & " And Exists" & vbNewLine & _
            " (select 1 from ���������¼ b where a.����id = b.����id And a.��ҳid = b.��ҳid and b.�п�='��') Group By a.��Ժ����id"

    Else   'ҽ������
        strSql = "Select Nvl(a.סԺҽʦ,'��') as ����, Count(1) As ����" & vbNewLine & _
            " From ������ҳ A Where a.��Ժ���� Between " & strPar & " And Exists" & vbNewLine & _
            " (select 1 from ���������¼ b where a.����id = b.����id And a.��ҳid = b.��ҳid and b.�п�='��') Group By Nvl(a.סԺҽʦ,'��')"
    End If
    Set rsһ�п����� = zlDatabase.OpenSQLRecord(strSql, Me.Caption, txtDept(e_C1_txtDept_ͳ�ƿ���_4).Tag)
    
    '���������ͳ���û��ҲӦ���˳�
    If rs��ҩ����.EOF And rsһ�п�����.EOF Then
        Screen.MousePointer = 0
        Call zlCommFun.StopFlash
        MsgBox "��ǰ������δ�ҵ��κ����ݣ�����������ͳ�Ʋ�����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    'һ���п�ʹ�ÿ���ҩĿ��ΪԤ������ϸ����ҩĿ�ģ�1 �������������ȷ��Ϊ����ҩ
    If optType(e_C1_optType_���ܷ�ʽ_����_17).Value Then '���һ���
        strSql = "Select a.��Ժ����id as ����id,a.����id,a.��ҳid" & vbNewLine & _
            " From ������ҳ A Where a.��Ժ���� Between " & strPar & " And Exists" & vbNewLine & _
            " (select 1 from ����ҽ����¼ c where c.ҽ��״̬ in (8,9) and c.����id=a.����id and c.��ҳid=a.��ҳid and c.��ҩĿ��=1 and c.�������='5')" & _
            " and Exists (select 1 from ���������¼ b where a.����id = b.����id And a.��ҳid = b.��ҳid and b.�п�='��')" & _
            " Group By a.��Ժ����id,a.����id, a.��ҳid"
    Else   'ҽ������
        strSql = "Select Nvl(a.סԺҽʦ,'��') as ����,a.����id,a.��ҳid" & vbNewLine & _
            " From ������ҳ A Where a.��Ժ���� Between " & strPar & " And Exists" & vbNewLine & _
            " (select 1 from ����ҽ����¼ c where c.ҽ��״̬ in (8,9) and c.����id=a.����id and c.��ҳid=a.��ҳid and c.��ҩĿ��=1 and c.�������='5')" & _
            " and Exists (select 1 from ���������¼ b where a.����id = b.����id And a.��ҳid = b.��ҳid and b.�п�='��')" & _
            " Group By Nvl(a.סԺҽʦ,'��'),a.����id, a.��ҳid"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, txtDept(e_C1_txtDept_ͳ�ƿ���_4).Tag)
    
    lng�п�Ԥ������ = rsTmp.RecordCount
    
    For i = 1 To rsTmp.RecordCount
        strPatis = strPatis & "," & rsTmp!����ID & ":" & rsTmp!��ҳID
        rsTmp.MoveNext
    Next
    strPatis = Mid(strPatis, 2) '�õ�����
    
    If strPatis <> "" Then
        '��ǰ��ҩ����
        strSql = "Select /*+ rule*/ a.����id,a.��ҳid,count(1) as ��ҩ����,To_Char(min(a.��ʼִ��ʱ��), 'YYYY-MM-DD HH24:MI:SS') as ��ʼִ��ʱ��," & _
            " to_char(max(a.������ʼʱ��), 'YYYY-MM-DD HH24:MI:SS') as ������ʼʱ��" & _
            " from ( select a.����id,a.��ҳid,b.������Ŀid,min(b.��ʼִ��ʱ��) as ��ʼִ��ʱ��,max(c.������ʼʱ��) as ������ʼʱ��" & _
            " From סԺ���ü�¼ A, ����ҽ����¼ B,���������¼ c" & vbNewLine & _
            " Where a.��¼״̬ <> 0 And a.�շ���� = '5' And a.ҽ����� = b.Id and a.����id=c.����id and a.��ҳid=c.��ҳid and c.�п�='��' and b.��ҩĿ��=1" & _
            " and (a.����id,a.��ҳid) In (Select C1, C2 From Table(f_Num2list2([1]))) group by a.����id,a.��ҳid,b.������Ŀid) a group by a.����id,a.��ҳid"
        Set rs��ǰ��ҩ���� = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strPatis)
        
        '�������ݼ�������
        strSql = "select ����id,��ҳid,Zl_Adviceexetimes(Id,��ʼִ��ʱ��,Nvl(Nvl(�ϴ�ִ��ʱ��,ִ����ֹʱ��),ͣ��ʱ��)," & _
            " ִ��ʱ�䷽��,��ʼִ��ʱ��,��ʼִ��ʱ��-1,Ƶ�ʼ��,�����λ,ҽ����Ч) as �ֽ�ʱ�� From ����ҽ����¼" & _
            " where (����id,��ҳid) In (Select To_Number(C1), C2 From Table(f_Str2list2([1]))) and ��ҩĿ��=1 and �������='5' and ҽ��״̬ in (8,9)"
        Set rs�������� = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strPatis)
        
        '��ҩ����������� ���˿����ؼ�¼ ��ȡ����Ϊ���������ڼ���
        strSql = "select a.����id,a.��ҳid,sum(a.ʹ������) as ���� from ���˿����ؼ�¼ a where (a.����id,a.��ҳid) In (Select To_Number(C1), C2 From Table(f_Str2list2([1]))) group by a.����id,a.��ҳid"
        Set rs���� = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strPatis)
    End If
    
    '��������
    With vsCut
        .Rows = .FixedRows
        For i = 1 To rs����.RecordCount
NextRow:
            If optType(e_C1_optType_���ܷ�ʽ_����_17).Value Then '���һ���
                rs��ҩ����.Filter = "����id=" & Val(rs����!����ID & "")
                rsһ�п�����.Filter = "����id=" & Val(rs����!����ID & "")
                rsTmp.Filter = "����id=" & Val(rs����!����ID & "")
            Else
                rs��ҩ����.Filter = "����='" & rs����!���� & "'"
                rsһ�п�����.Filter = "����='" & rs����!���� & "'"
                rsTmp.Filter = "����='" & rs����!���� & "'"
            End If
            
            lng������ = 0
            If Not rsһ�п�����.EOF Then lng������ = Val(rsһ�п�����!���� & "")
            lngTmp = 0
            If Not rs��ҩ����.EOF Then lngTmp = Val(rs��ҩ����!���� & "")
            
            If lng������ = 0 And lngTmp = 0 Then
                i = i + 1
                rs����.MoveNext
                If rs����.EOF Then Exit For
                GoTo NextRow
            End If
            
            .AddItem "": lngRow = .Rows - 1
            
            .TextMatrix(lngRow, COL_CUT����) = rs����!���� & ""
            lng������ = Val(rs����!���� & "")
            
            If Not rs��ҩ����.EOF Then lngTmp = Val(rs��ҩ����!���� & "")
            .TextMatrix(lngRow, COL_CUTʹ���˴�) = lngTmp
            
            If lng������ <> 0 Then dblTmp = lngTmp * 100 / lng������
            .TextMatrix(lngRow, COL_CUTʹ����) = Format(dblTmp, "0.00") & "%"
            lngTmp = 0: dblTmp = 0
            
            If Not rsһ�п�����.EOF Then lngTmp = Val(rsһ�п�����!���� & "")
            .TextMatrix(lngRow, COL_CUT�п���) = lngTmp: lngTmp = 0
 
            If Not rsTmp.EOF Then lngTmp = rsTmp.RecordCount
            .TextMatrix(lngRow, COL_CUT��������) = lngTmp: lngTmp = 0
    
            If Val(.TextMatrix(lngRow, COL_CUT�п���)) <> 0 Then dblTmp = Val(.TextMatrix(lngRow, COL_CUT��������)) * 100 / Val(rsһ�п�����!���� & "")
            .TextMatrix(lngRow, COL_CUT�п�ʹ����) = Format(dblTmp, "0.00") & "%": dblTmp = 0
            
            For j = 1 To rsTmp.RecordCount
                If strPatis <> "" Then
                    rs��ǰ��ҩ����.Filter = "����id=" & rsTmp!����ID & " and ��ҳid=" & rsTmp!��ҳID
                    If Not rs��ǰ��ҩ����.EOF Then
                        lngTmp = lngTmp + 1 '�п�Ԥ���ÿ�������
                        lngҩƷ�� = lngҩƷ�� + Val(rs��ǰ��ҩ����!��ҩ���� & "") '�п�Ԥ���ÿ�����ҩ��Ʒ����
                        If rs��ǰ��ҩ����!��ʼִ��ʱ�� & "" <> "" And rs��ǰ��ҩ����!������ʼʱ�� & "" <> "" Then '�п�Ԥ���ÿ�����ǰʹ������
                            If rs��ǰ��ҩ����!��ʼִ��ʱ�� & "" < rs��ǰ��ҩ����!������ʼʱ�� & "" Then lng��ǰ�� = lng��ǰ�� + 1
                        End If
                    End If
                    '����
                    rs��������.Filter = "����id=" & rsTmp!����ID & " and ��ҳid=" & rsTmp!��ҳID
                    If Not rs��������.EOF Then
                        For k = 1 To rs��������.RecordCount
                            strTmp = rs��������!�ֽ�ʱ�� & ""
                            If strTmp <> "" Then
                                arrTmp = Split(strTmp, ",")
                                For m = 0 To UBound(arrTmp)
                                    strTmp = Split(arrTmp(m), " ")(0)
                                    If InStr("," & str���� & ",", "," & strTmp & ",") = 0 Then
                                        str���� = str���� & "," & strTmp
                                    End If
                                Next
                            End If
                            rs��������.MoveNext
                        Next
                        
                        If str���� <> "" Then
                            str���� = Mid(str����, 2)
                            lngTmp = UBound(Split(str����, ",")) + 1
                        End If
                        
                        lng���� = lng���� + lngTmp
                        lngTmp = 0: str���� = "": strTmp = ""
                    Else
                        rs����.Filter = "����id=" & rsTmp!����ID & " and ��ҳid=" & rsTmp!��ҳID
                        If Not rs����.EOF Then
                            lng���� = lng���� + Val(rs����!���� & "")
                        Else
                            lng���� = lng���� + 1
                        End If
                    End If
                    
                End If
                rsTmp.MoveNext
            Next
            
            .TextMatrix(lngRow, COL_CUT��������) = lngTmp: lngTmp = 0 '�п�Ԥ���ÿ�������
            
            .TextMatrix(lngRow, COL_CUT��ǰ��ҩ) = lng��ǰ��: lng��ǰ�� = 0 '�п�Ԥ���ÿ�����ǰʹ������
            
            .TextMatrix(lngRow, COL_CUTƽ����ҩ) = lng����: lng���� = 0 '����
            
            .TextMatrix(lngRow, COL_CUTƷ����) = lngҩƷ��: lngҩƷ�� = 0 '�п�Ԥ���ÿ�����ҩ��Ʒ����
            rs����.MoveNext
        Next
    End With
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\���ڷ�Χ", "�����п�Χ����Ԥ����ҩͳ��", dtpCountS(e_C1_dtpCountS_��ʼʱ��_1).Value & "," & dtpCountE(e_C1_dtpCountE_����ʱ��_1).Value
    Screen.MousePointer = 0
    Call zlCommFun.StopFlash
    Exit Sub
errH:
    Screen.MousePointer = 0
    Call zlCommFun.StopFlash
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadvsInDruUse()
'���ܣ�����   סԺҽ��������ҩͳ��  ���������ݼ���
    Dim strSql As String, strPar As String, strDept As String
    Dim rsTmp As ADODB.Recordset
    Dim strҽ��IDs As String
    Dim rs��� As ADODB.Recordset       '��Ժ��� һ�Զ�
    Dim rs���� As ADODB.Recordset       '��ȡ �������� �� �п����� ���� һ�Զ�
    Dim rs��� As ADODB.Recordset       '���ƽ�ҩƷ��� һ��һ
    Dim rs���� As ADODB.Recordset       'ҩƷ����������ҩ��Ʒ���� һ��һ
    Dim rs���� As ADODB.Recordset
    Dim lngBaseRow As Long, lngTmpRow As Long
    Dim dblTmp As Double, strDec As String
    Dim rs��ҩ��ϸ As ADODB.Recordset   '����ҩ��ʹ�����
    Dim lng������ As Long
    Dim str�п� As String
    Dim strPatis As String '������Ϣ ��ʽ��"����id1:��ҳid1,����id2:��ҳid2,......."
    Dim strParTable As String, strTable As String
    Dim varArr As Variant
    Dim strFilter As String
    Dim strTmp As String
    Dim strTmp1 As String
    Dim lngTmp As Long
    Dim i As Long, j As Long, k As Long
    
    'ͳ�Ʒ���
    Dim lngҩƷ�� As Long
    Dim lng��ҩ�� As Long
    Dim lng�ÿ���ҩ���� As Long
    Dim lngʵ������ As Long
    Dim dbl�ܽ�� As Double
    Dim dblҩƷ��� As Double
    Dim dbl��ҩ��� As Double
    Dim lng�������� As Long
    Dim lng�������� As Long
    Dim lng�������� As Long
    Dim lng�������� As Long
    Dim lngԤ������ As Long
    Dim lng�������� As Long
    Dim str��ҩ�� As String
    
    On Error GoTo errH
    Screen.MousePointer = 11
    Call zlCommFun.ShowFlash("���ڶ�ȡ���ݣ����Ժ�...")
    
    lng������ = Val(txtNum(e_C3_txtNum_��������_1).Text)
    strDec = "0.00"
    
    '���ڷ�Χ��������
    strPar = "To_Date('" & Format(dtpCountS(e_C3_dtpCountS_��ʼʱ��_3).Value, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')" & _
        " and To_Date('" & Format(dtpCountE(e_C3_dtpCountE_����ʱ��_3).Value, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
    strDept = IIf(txtDept(e_C3_txtDept_ͳ�ƿ���_6).Tag = "", "", " and a.��Ժ����id in (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))")
     
    '����SQL��ָ��ʱ���ڣ�ָ�������ڣ��ù�����û�ù�����ҩ�Ĳ��˶�Ӧ�ó���
    strSql = "Select a.����id,a.��ҳid,a.סԺ��,a.����,To_Char(a.��Ժ����, 'YYYY-MM-DD HH24:MI:SS') as ��Ժ����,a.סԺҽʦ as ����ҽ��,c.���� as ����,a.סԺ����" & _
        " From ������ҳ A,���ű� c" & _
        " Where a.��Ժ����id =c.id" & _
        " And a.��Ժ���� Between " & strPar & strDept
        
    If optType(e_C3_optType_�п�����_������_15).Value Then '������
        If optType(e_C3_optType_��������_ƽ��_14).Value Then
            'ƽ������
            strSql = strSql & " And Not Exists (Select 1 From ���������¼ X Where a.����id=x.����id And a.��ҳid=x.��ҳid) Order By a.��Ժ���� desc"
            
            strSql = "select m.����id,m.��ҳid,m.סԺ��,m.����,m.��Ժ����,m.����ҽ��,m.����,m.סԺ���� from (" & _
                "select m.����id,m.��ҳid,m.סԺ��,m.����,m.��Ժ����,m.����ҽ��,m.����,m.סԺ���� from (" & _
                    "select m.����id,m.��ҳid,m.סԺ��,m.����,m.��Ժ����,m.����ҽ��,m.����,m.סԺ����,Mod(Rownum,[2]) M from (" & strSql & ") m  Order By M) M " & _
                     " Where Rownum <([2]+1)) M Order By m.��Ժ���� Desc"
        Else
            '�������
            strSql = strSql & " And Not Exists (Select 1 From ���������¼ X Where a.����id=x.����id And a.��ҳid=x.��ҳid) Order By Dbms_Random.Value"
            strSql = "select ����id,��ҳid,סԺ��,����,��Ժ����,����ҽ��,����,סԺ���� from (" & strSql & ") where rownum < ([2] + 1) Order By ��Ժ���� desc"
        End If
    Else
        str�п� = str�п� & IIf(chkType(e_C3_chkType_�п�����_����_2).Value = 1, "��,", "")
        str�п� = str�п� & IIf(chkType(e_C3_chkType_�п�����_����_3).Value = 1, "��,", "")
        str�п� = str�п� & IIf(chkType(e_C3_chkType_�п�����_����_4).Value = 1, "��,", "")
        str�п� = str�п� & IIf(chkType(e_C3_chkType_�п�����_����_8).Value = 1, "��,", "")
        
        If str�п� = "��,��,��,��," Then str�п� = ""
         
        If optType(e_C3_optType_��������_ƽ��_14).Value Then
            'ƽ������
            strSql = strSql & " And Exists (Select 1 From ���������¼ B Where a.����id=b.����id And a.��ҳid=b.��ҳid" & _
            IIf(str�п� = "", "", " And Instr([3],b.�п�)>0") & ") Order By a.��Ժ���� desc"
            
            strSql = "select m.����id,m.��ҳid,m.סԺ��,m.����,m.��Ժ����,m.����ҽ��,m.����,m.סԺ���� from (" & _
                " select m.����id,m.��ҳid,m.סԺ��,m.����,m.��Ժ����,m.����ҽ��,m.����,m.סԺ���� from (" & _
                " select m.����id,m.��ҳid,m.סԺ��,m.����,m.��Ժ����,m.����ҽ��,m.����,m.סԺ����,Mod(Rownum,[2]) M from (" & strSql & ") m  Order By M) M " & _
                " Where Rownum <([2]+1)) M Order By m.��Ժ���� Desc"
        Else
            '�������
            strSql = strSql & " And Exists (Select 1 From ���������¼ B  Where a.����id=b.����id And a.��ҳid=b.��ҳid" & IIf(str�п� = "", "", " And Instr([3],b.�п�)>0") & ") Order By Dbms_Random.Value"
            strSql = "select ����id,��ҳid,סԺ��,����,��Ժ����,����ҽ��,����,סԺ���� from (" & strSql & ") where rownum < ([2] + 1) Order By ��Ժ����"
        End If
    End If
    
    '�������
    vsInDruUse.Rows = vsInDruUse.FixedRows
    vsInDruUse.Rows = vsInDruUse.FixedRows + 1
    '����ʾ������
    vsInDruUse.ColHidden(COL_DRU��������) = optType(e_C3_optType_�п�����_������_15).Value 'optType(15).Value ��true ������ false ����
    vsInDruUse.ColHidden(COL_DRU�п�����) = optType(e_C3_optType_�п�����_������_15).Value
    lblN(e_C3_lblN_������_����_59).Caption = "0����Ժ���˿���ҩ��ʹ��ͳ�Ʒ�����"
    
    '���������ң����������п�
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, txtDept(e_C3_txtDept_ͳ�ƿ���_6).Tag, lng������, str�п�)
    
    If rsTmp.EOF Then
        Screen.MousePointer = 0
        Call zlCommFun.StopFlash
        MsgBox "���յ�ǰ���õ�����δ�����κ����ݣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    lngʵ������ = rsTmp.RecordCount '�������������ܱȽ�����д������С
    lblN(e_C3_lblN_������_����_59).Caption = lngʵ������ & "����Ժ���˿���ҩ��ʹ��ͳ�Ʒ�����"
    
    For i = 1 To rsTmp.RecordCount
        strPatis = strPatis & "," & rsTmp!����ID & ":" & rsTmp!��ҳID
        rsTmp.MoveNext
    Next

    strPatis = Mid(strPatis, 2) '�õ�����
    
    strParTable = "Select C1, C2 From Table(f_Num2list2([1]))"
    strTable = strParTable
    
    If Len(strPatis) >= 4000 Then
        varArr = Array()
        varArr = GetParTable(strPar, strParTable, strTable)
    End If
    
    'ȡ��ҳ�ĵ�һ����Ժ��ϣ���������ҽ
    strSql = "select a.����id,a.��ҳid,a.������� as ��� from ������ϼ�¼ a where a.��¼��Դ=3 And NVL(A.�������,1) = 1 and a.��ϴ���=1 and a.������� in (3,13) and (a.����id,a.��ҳid) In (" & strTable & ")"
    
    If Len(strPatis) >= 4000 Then
        Set rs��� = zlDatabase.OpenSQLRecord(strSql, Me.Caption, CStr(varArr(0)), CStr(varArr(1)), CStr(varArr(2)), CStr(varArr(3)), CStr(varArr(4)), CStr(varArr(5)), _
            CStr(varArr(6)), CStr(varArr(7)), CStr(varArr(8)), CStr(varArr(9)))
    Else
        Set rs��� = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strPatis)
    End If
    
    If Not optType(e_C3_optType_�п�����_������_15).Value Then '��������
        strSql = "select a.����id,a.��ҳid,a.�������� as ����,a.�п� from ���������¼ a where (a.����id,a.��ҳid) In (" & strTable & ")"
      
        If Len(strPatis) >= 4000 Then
            Set rs���� = zlDatabase.OpenSQLRecord(strSql, Me.Caption, CStr(varArr(0)), CStr(varArr(1)), CStr(varArr(2)), CStr(varArr(3)), CStr(varArr(4)), CStr(varArr(5)), _
                CStr(varArr(6)), CStr(varArr(7)), CStr(varArr(8)), CStr(varArr(9)))
        Else
            Set rs���� = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strPatis)
        End If

    End If
    
    '����ҩ���������ȡ
    strSql = "select a.����id,a.��ҳid,sum(a.���ʽ��) as ���ƽ��,Sum(Decode(a.�շ����,'5',a.���ʽ��,'6',a.���ʽ��,'7',a.���ʽ��, 0)) As ҩƷ���" & _
        " from סԺ���ü�¼ a where a.��¼״̬<>0 and (a.����id,a.��ҳid) In (" & strTable & ") group by a.����id,a.��ҳid"
        
    If Len(strPatis) >= 4000 Then
        Set rs��� = zlDatabase.OpenSQLRecord(strSql, Me.Caption, CStr(varArr(0)), CStr(varArr(1)), CStr(varArr(2)), CStr(varArr(3)), CStr(varArr(4)), CStr(varArr(5)), _
            CStr(varArr(6)), CStr(varArr(7)), CStr(varArr(8)), CStr(varArr(9)))
    Else
        Set rs��� = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strPatis)
    End If
    
    'ҩƷ���� ����ȡ�� ����ҩ����
    strSql = "Select a.����id, a.��ҳid, Count(1) As ҩƷ����, Sum(a.����ҩ) As ����ҩ����, Sum(a.����ҩ���) As ����ҩ���" & vbNewLine & _
        "From (Select a.����id, a.��ҳid, c.ҩ��id, Decode(Nvl(c.������, 0), 0, 0, 1) As ����ҩ," & vbNewLine & _
        "              Sum(Decode(Nvl(c.������, 0), 0, 0, a.���ʽ��)) As ����ҩ���" & vbNewLine & _
        "       From סԺ���ü�¼ A, ҩƷ��� B, ҩƷ���� C" & vbNewLine & _
        "       Where a.�շ���� In ('5', '6', '7') And a.�շ�ϸĿid = b.ҩƷid And b.ҩ��id = c.ҩ��id And a.��¼״̬ <> 0 And" & vbNewLine & _
        "             (a.����id, a.��ҳid) In (" & strTable & ")" & vbNewLine & _
        "       Group By a.����id, a.��ҳid, c.ҩ��id, Nvl(c.������, 0)) A" & vbNewLine & _
        "Group By a.����id, a.��ҳid"
    
    
    If Len(strPatis) >= 4000 Then
        Set rs���� = zlDatabase.OpenSQLRecord(strSql, Me.Caption, CStr(varArr(0)), CStr(varArr(1)), CStr(varArr(2)), CStr(varArr(3)), CStr(varArr(4)), CStr(varArr(5)), _
            CStr(varArr(6)), CStr(varArr(7)), CStr(varArr(8)), CStr(varArr(9)))
    Else
        Set rs���� = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strPatis)
    End If

    '��ҩ��ϸ
    strSql = "select e.����id,e.��ҳid,e.id,e.���id,e.ҽ����Ч,a.ҩƷ����,a.����,a.���,e.ִ��Ƶ��,e.��������,f.���㵥λ," & vbNewLine & _
        "e.����,g.ҽ������ as ��ҩ;��,decode(e.��ҩĿ��,1,'Ԥ��',2,'����','') as Ŀ��" & vbNewLine & _
        "from (Select a.ҽ����� as ҽ��id,d.���� as ҩƷ����,c.ҩƷ���� as ����,d.��� || d.���� As ���" & vbNewLine & _
        "From סԺ���ü�¼ A, ҩƷ��� B, ҩƷ���� C,�շ���ĿĿ¼ d" & vbNewLine & _
        "Where a.�շ���� = '5' And a.�շ�ϸĿid = b.ҩƷid And b.ҩ��id = c.ҩ��id" & vbNewLine & _
        "and b.ҩƷid=d.id  And Nvl(c.������, 0) <> 0 And a.��¼״̬ <> 0 And" & vbNewLine & _
        "      (a.����id, a.��ҳid) In (" & strTable & ")" & vbNewLine & _
        "group by a.ҽ�����,d.����,c.ҩƷ����,d.���,d.����) a,����ҽ����¼ e,������ĿĿ¼ f,����ҽ����¼ g" & vbNewLine & _
        "where a.ҽ��id=e.id and e.������Ŀid=f.id and e.���id=g.id"
    
    If Len(strPatis) >= 4000 Then
        Set rs��ҩ��ϸ = zlDatabase.OpenSQLRecord(strSql, Me.Caption, CStr(varArr(0)), CStr(varArr(1)), CStr(varArr(2)), CStr(varArr(3)), CStr(varArr(4)), CStr(varArr(5)), _
            CStr(varArr(6)), CStr(varArr(7)), CStr(varArr(8)), CStr(varArr(9)))
    Else
        Set rs��ҩ��ϸ = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strPatis)
    End If
    
    '��ҩ������ֻ���㳤��ҽ��
    rs��ҩ��ϸ.Filter = "ҽ����Ч=0"
    If Not rs��ҩ��ϸ.EOF Then
        For i = 1 To rs��ҩ��ϸ.RecordCount
            If InStr("," & strҽ��IDs & ",", "," & rs��ҩ��ϸ!���ID & ",") = 0 Then
                strҽ��IDs = strҽ��IDs & "," & rs��ҩ��ϸ!���ID
            End If
            rs��ҩ��ϸ.MoveNext
        Next
        strҽ��IDs = Mid(strҽ��IDs, 2)
        If strҽ��IDs <> "" Then
            strSql = "select id as ҽ��id,Zl_Adviceexetimes(Id,��ʼִ��ʱ��,Nvl(�ϴ�ִ��ʱ��,ִ����ֹʱ��),ִ��ʱ�䷽��,��ʼִ��ʱ��,��ʼִ��ʱ��-1,Ƶ�ʼ��,�����λ,ҽ����Ч) as �ֽ�ʱ��" & vbNewLine & _
                "From ����ҽ����¼ where id in (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))"
            Set rs���� = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strҽ��IDs)
        End If
        rs��ҩ��ϸ.Filter = 0
    End If
    
    '��������
    rsTmp.MoveFirst
    
    With vsInDruUse
        .Rows = .FixedRows
        For i = 1 To rsTmp.RecordCount
            .AddItem ""
            lngBaseRow = .Rows - 1
            .TextMatrix(lngBaseRow, COL_DRU����id) = rsTmp!����ID
            .TextMatrix(lngBaseRow, COL_DRU��ҳid) = rsTmp!��ҳID
            .TextMatrix(lngBaseRow, COL_DRU���) = i
            .TextMatrix(lngBaseRow, COL_DRUסԺ��) = rsTmp!סԺ�� & ""
            .TextMatrix(lngBaseRow, COL_DRU��������) = rsTmp!���� & ""
            .TextMatrix(lngBaseRow, COL_DRU��Ժ����) = Format(rsTmp!��Ժ���� & "", "YYYY-MM-DD")
            .TextMatrix(lngBaseRow, COL_DRU����ҽ��) = rsTmp!����ҽ�� & ""
            .TextMatrix(lngBaseRow, COL_DRU����) = rsTmp!���� & ""
            .TextMatrix(lngBaseRow, COL_DRUסԺ����) = rsTmp!סԺ���� & ""
            
            strFilter = "����id=" & rsTmp!����ID & " And ��ҳid=" & rsTmp!��ҳID
            '���
            rs���.Filter = strFilter
            If Not rs���.EOF Then rs���.MoveFirst
            strTmp = ""
            Do Until rs���.EOF
                If InStr(";" & strTmp & ";", ";" & rs���!��� & ";") = 0 Then
                    strTmp = strTmp & ";" & rs���!���
                End If
                rs���.MoveNext
            Loop
            .TextMatrix(lngBaseRow, COL_DRU��Ժ���) = Mid(strTmp, 2)
            strTmp = "": rs���.Filter = 0
            '�������п�
            If Not optType(e_C3_optType_�п�����_������_15).Value Then
                rs����.Filter = strFilter
                If Not rs����.EOF Then rs����.MoveFirst
                Do Until rs����.EOF
                    strTmp = strTmp & ";" & rs����!����
                    If InStr("," & strTmp1 & ",", "," & rs����!�п� & ",") = 0 Then
                        strTmp1 = strTmp1 & "," & rs����!�п�
                    End If
                    rs����.MoveNext
                Loop
                .TextMatrix(lngBaseRow, COL_DRU��������) = Mid(strTmp, 2)
                .TextMatrix(lngBaseRow, COL_DRU�п�����) = Mid(strTmp1, 2)
                strTmp = "": strTmp1 = "": rs����.Filter = 0
            End If
            '���
            rs���.Filter = strFilter
            If Not rs���.EOF Then
                rs���.MoveFirst
                .TextMatrix(lngBaseRow, COL_DRU���ƽ��) = Format(rs���!���ƽ�� & "", strDec)
                .TextMatrix(lngBaseRow, COL_DRUҩƷ���) = Format(rs���!ҩƷ��� & "", strDec)
                If Val(rsTmp!סԺ���� & "") <> 0 Then dblTmp = Val(rs���!���ƽ�� & "") / Val(rsTmp!סԺ���� & "")
                .TextMatrix(lngBaseRow, COL_DRU�վ����ƽ��) = Format(rs���!ҩƷ��� & "", strDec): dblTmp = 0
                
                dbl�ܽ�� = dbl�ܽ�� + Val(rs���!���ƽ�� & "")
                dblҩƷ��� = dblҩƷ��� + Val(rs���!ҩƷ��� & "")
                
            End If
            rs���.Filter = 0
            '����
            rs����.Filter = strFilter
            If Not rs����.EOF Then
                .TextMatrix(lngBaseRow, COL_DRUҩƷ����) = rs����!ҩƷ���� & ""
                .TextMatrix(lngBaseRow, COL_DRU����ҩ��Ʒ����) = rs����!����ҩ���� & ""
                .TextMatrix(lngBaseRow, COL_DRU����ҩ����) = Format(rs����!����ҩ��� & "", strDec)
                .TextMatrix(lngBaseRow, COL_DRU������ҩ) = Decode(Val(rs����!����ҩ���� & ""), 0, "", 1, "����", 2, "����", 3, "����", 4, "����", ">����")
                
                lngҩƷ�� = lngҩƷ�� + Val(rs����!ҩƷ���� & "")
                lng��ҩ�� = lng��ҩ�� + Val(rs����!����ҩ���� & "")
                
                dbl��ҩ��� = dbl��ҩ��� + Val(rs����!����ҩ��� & "")
                
                If Val(rs����!����ҩ���� & "") <> 0 Then
                    lng�ÿ���ҩ���� = lng�ÿ���ҩ���� + 1
                    
                    Select Case Val(rs����!����ҩ���� & "")
                        Case 1
                            lng�������� = lng�������� + 1
                        Case 2
                            lng�������� = lng�������� + 1
                        Case 3
                            lng�������� = lng�������� + 1
                        Case 4
                            lng�������� = lng�������� + 1
                    End Select
                End If
            End If
            rs����.Filter = 0
 
            '��ϸ��ҩ
            rs��ҩ��ϸ.Filter = strFilter
            If Not rs��ҩ��ϸ.EOF Then
                For j = 1 To rs��ҩ��ϸ.RecordCount
                    
                    '��ȡ�÷�����������strTmp ��
                    strTmp = rs��ҩ��ϸ!�������� & ""
                    If Mid(strTmp, 1, 1) = "." Then strTmp = "0" & strTmp
                    strTmp = rs��ҩ��ϸ!ִ��Ƶ�� & "," & strTmp & rs��ҩ��ϸ!���㵥λ
                
                    If j = 1 Then
                        lngTmpRow = lngBaseRow
                        .TextMatrix(lngBaseRow, COL_DRUҩƷ����) = rs��ҩ��ϸ!ҩƷ���� & ""
                        .TextMatrix(lngBaseRow, COL_DRU����) = rs��ҩ��ϸ!���� & ""
                        .TextMatrix(lngBaseRow, COL_DRU���) = rs��ҩ��ϸ!��� & ""
                        
                        .TextMatrix(lngTmpRow, COL_DRU�÷�����) = strTmp
                        
                        .TextMatrix(lngBaseRow, COL_DRU��ҩ����) = Val(rs��ҩ��ϸ!���� & "")
                        .TextMatrix(lngBaseRow, COL_DRU��ҩ;��) = rs��ҩ��ϸ!��ҩ;�� & ""
                        .TextMatrix(lngBaseRow, COL_DRU��ҩĿ��) = IIf("" = rs��ҩ��ϸ!Ŀ�� & "", "��", rs��ҩ��ϸ!Ŀ��)
                    Else
                        .AddItem ""
                        lngTmpRow = .Rows - 1
                        .TextMatrix(lngTmpRow, COL_DRU����id) = rsTmp!����ID
                        .TextMatrix(lngTmpRow, COL_DRU��ҳid) = rsTmp!��ҳID
                        
                        .TextMatrix(lngTmpRow, COL_DRUҩƷ����) = rs��ҩ��ϸ!ҩƷ���� & ""
                        .TextMatrix(lngTmpRow, COL_DRU����) = rs��ҩ��ϸ!���� & ""
                        .TextMatrix(lngTmpRow, COL_DRU���) = rs��ҩ��ϸ!��� & ""
                        .TextMatrix(lngTmpRow, COL_DRU�÷�����) = strTmp ' rs��ҩ��ϸ!�÷����� & ""
                        .TextMatrix(lngTmpRow, COL_DRU��ҩ����) = Val(rs��ҩ��ϸ!���� & "")
                        .TextMatrix(lngTmpRow, COL_DRU��ҩ;��) = rs��ҩ��ϸ!��ҩ;�� & ""
                        .TextMatrix(lngTmpRow, COL_DRU��ҩĿ��) = IIf("" = rs��ҩ��ϸ!Ŀ�� & "", "��", rs��ҩ��ϸ!Ŀ��)
                    End If
                    
                    '��ҩ����
                    If Not rs���� Is Nothing Then
                        rs����.Filter = "ҽ��id=" & rs��ҩ��ϸ!���ID
                        
                        If Not rs����.EOF Then
                            strTmp = rs����!�ֽ�ʱ�� & ""
                            If strTmp <> "" Then .TextMatrix(lngTmpRow, COL_DRU��ҩ����) = UBound(Split(strTmp, ",")) + 1
                        End If
                    End If
                    
                    If InStr("," & str��ҩ�� & ",", "," & .TextMatrix(lngTmpRow, COL_DRU��ҩĿ��) & ",") = 0 Then
                        str��ҩ�� = str��ҩ�� & "," & .TextMatrix(lngTmpRow, COL_DRU��ҩĿ��)
                    End If
                    rs��ҩ��ϸ.MoveNext
                Next
            End If
            rs��ҩ��ϸ.Filter = 0
            
            If InStr("," & str��ҩ�� & ",", ",Ԥ��,") > 0 Then lngԤ������ = lngԤ������ + 1
            If InStr("," & str��ҩ�� & ",", ",����,") > 0 Then lng�������� = lng�������� + 1
            
            str��ҩ�� = ""
'-----------------------------------------------------------------------------
            rsTmp.MoveNext
        Next
    End With
        
    '����ͳ�Ʒ������
    With vsInDruAna
'        .AddItem "" '��0��
        lngTmp = 0
        .TextMatrix(0, 0) = "A(��ҩ��Ʒ����)=" & lngҩƷ�� & "��"
    
        If lng�ÿ���ҩ���� <> 0 Then dblTmp = lngҩƷ�� / lng�ÿ���ҩ����
        .TextMatrix(0, 1) = "B(ƽ����ҩƷ����A/E)=" & Format(dblTmp, strDec) & "��": dblTmp = 0
        
        .TextMatrix(0, 2) = "C(ʹ�ÿ���ҩ���Ʒ����)=" & lng��ҩ�� & "��"
        
        dblTmp = lng�ÿ���ҩ���� * 100 / lngҩƷ��
        .TextMatrix(0, 3) = "D(ʹ�ÿ���ҩ��İٷ���E/A)*100%=" & Format(dblTmp, strDec) & "%": dblTmp = 0
        
'        .AddItem "" '��1��
        lngTmp = 1
        .TextMatrix(lngTmp, 0) = "E(ʹ�ÿ�ҩ��Ĳ�����)=" & lng�ÿ���ҩ���� & "��"
        
        dblTmp = lng�ÿ���ҩ���� * 100 / lngʵ������
        .TextMatrix(lngTmp, 1) = "F(��Ժ���˿���ҩ��ʹ����E/ʵ������)*100%=" & Format(dblTmp, strDec) & "%": dblTmp = 0
        
        .TextMatrix(lngTmp, 2) = "G(�����ܽ��)=" & dbl�ܽ�� & "Ԫ"
        dblTmp = dbl�ܽ�� / lngʵ������
        .TextMatrix(lngTmp, 3) = "H(����ƽ�����ƽ��G/ʵ������)=" & Format(dblTmp, strDec) & "Ԫ": dblTmp = 0
        
'        .AddItem "" '��2��
        lngTmp = 2
        .TextMatrix(lngTmp, 0) = "I(ҩƷ�ܽ��)=" & dblҩƷ��� & "Ԫ"
        
        If dbl�ܽ�� <> 0 Then dblTmp = dblҩƷ��� * 100 / dbl�ܽ��
        .TextMatrix(lngTmp, 1) = "L(ҩƷ�ܽ��ռ�����ܽ��İٷ���I/G)*100%=" & Format(dblTmp, strDec) & "%": dblTmp = 0
        
        .TextMatrix(lngTmp, 2) = "K(����ҩ���ܽ��)=" & dbl��ҩ��� & "Ԫ"
        
        If dblҩƷ��� <> 0 Then dblTmp = dbl��ҩ��� * 100 / dblҩƷ���
        .TextMatrix(lngTmp, 3) = "J(����ҩ���ܽ��ռҩƷ�ܽ��İٷ���K/I)*100%=" & Format(dblTmp, strDec) & "%": dblTmp = 0
        
'        .AddItem "" '��3��
        lngTmp = 3
        .TextMatrix(lngTmp, 0) = "M(���ÿ���ҩ��Ĳ�����)=" & lng�������� & "��"
        
        If lng�ÿ���ҩ���� <> 0 Then dblTmp = lng�������� * 100 / lng�ÿ���ҩ����
        .TextMatrix(lngTmp, 1) = "O(���ÿ���ҩ���ʹ����M/E)*100%=" & Format(dblTmp, strDec) & "%": dblTmp = 0
        
        .TextMatrix(lngTmp, 2) = "P(����ʹ�ÿ���ҩ��Ĳ�����)=" & lng�������� & "��"
        
        If lng�ÿ���ҩ���� <> 0 Then dblTmp = lng�������� * 100 / lng�ÿ���ҩ����
        .TextMatrix(lngTmp, 3) = "Q(����ʹ�ÿ���ҩ���ʹ����P/E)*100%=" & Format(dblTmp, strDec) & "%": dblTmp = 0

'        .AddItem "" '��4��
        lngTmp = 4
        .TextMatrix(lngTmp, 0) = "R(����ʹ�ÿ���ҩ��Ĳ�����)=" & lng�������� & "��"
        
        If lng�ÿ���ҩ���� <> 0 Then dblTmp = lng�������� * 100 / lng�ÿ���ҩ����
        .TextMatrix(lngTmp, 1) = "S(����ʹ�ÿ���ҩ���ʹ����R/E)*100%=" & Format(dblTmp, strDec) & "%": dblTmp = 0
        
        .TextMatrix(lngTmp, 2) = "T(����ʹ�ÿ���ҩ��Ĳ�����)=" & lng�������� & "��"
        
        If lng�ÿ���ҩ���� <> 0 Then dblTmp = lng�������� * 100 / lng�ÿ���ҩ����
        .TextMatrix(lngTmp, 3) = "U(����ʹ�ÿ���ҩ���ʹ����T/E)*100%=" & Format(dblTmp, strDec) & "%": dblTmp = 0

'        .AddItem "" '��5��
        lngTmp = 5
        .TextMatrix(lngTmp, 0) = "V(Ԥ��ʹ�ÿ���ҩ��Ĳ�����)=" & lngԤ������ & "��"
        
        If lng�ÿ���ҩ���� <> 0 Then dblTmp = lngԤ������ * 100 / lng�ÿ���ҩ����
        .TextMatrix(lngTmp, 1) = "W(Ԥ��ʹ�ÿ���ҩ�ﹹ�ɱ�V/E)100%=" & Format(dblTmp, strDec) & "%": dblTmp = 0
        
        .TextMatrix(lngTmp, 2) = "X(����ʹ�ÿ���ҩ��Ĳ�����)=" & lng�������� & "��"
        
        If lng�ÿ���ҩ���� <> 0 Then dblTmp = lng�������� * 100 / lng�ÿ���ҩ����
        .TextMatrix(lngTmp, 3) = "Y(����ʹ�ÿ���ҩ�ﹹ�ɱ�Y/E)*100%=" & Format(dblTmp, strDec) & "%": dblTmp = 0

    End With
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\���ڷ�Χ", "סԺҽ��������ҩͳ��", dtpCountS(e_C3_dtpCountS_��ʼʱ��_3).Value & "," & dtpCountE(e_C3_dtpCountE_����ʱ��_3).Value
    Screen.MousePointer = 0
    Call zlCommFun.StopFlash
    Exit Sub
errH:
    Screen.MousePointer = 0
    Call zlCommFun.StopFlash
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadvsOpeKssUse()
'���ܣ�����   ���󿹾�ҩ��ʹ�ó�N��ͳ��  ���������
    Dim strSql As String, strPar As String, strDept As String
    Dim rsTmp As ADODB.Recordset
    Dim rs���� As ADODB.Recordset
    Dim str�п� As String
    Dim lng������ As Long
    Dim strPatis As String
    Dim strFilter As String
    Dim strTmp As String, strTmp1 As String
    Dim i As Long
    Dim intD As Integer, intC As Integer
    Dim lngTmp As Long
    Dim strParTable As String, strTable As String
    Dim varArr As Variant
      
    '����п�ȫѡ����ȫ��ѡ����Ϊ�ǲ������п�����
    str�п� = str�п� & IIf(chkType(e_C4_chkType_�п�����_����_5).Value = 1, "��,", "")
    str�п� = str�п� & IIf(chkType(e_C4_chkType_�п�����_����_6).Value = 1, "��,", "")
    str�п� = str�п� & IIf(chkType(e_C4_chkType_�п�����_����_7).Value = 1, "��,", "")
    str�п� = str�п� & IIf(chkType(e_C4_chkType_�п�����_����_9).Value = 1, "��,", "")
    
    If str�п� = "��,��,��,��," Then str�п� = ""
    
    lng������ = Val(txtNum(e_C4_txtNum_��������_2).Text)
    
    '���ڷ�Χ��������
    strPar = "To_Date('" & Format(dtpCountS(e_C4_dtpCountS_��ʼʱ��_4).Value, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')" & _
        " and To_Date('" & Format(dtpCountE(e_C4_dtpCountE_����ʱ��_4).Value, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
    strDept = IIf(txtDept(e_C4_txtDept_ͳ�ƿ���_7).Tag = "", "", " and a.��Ժ����id in (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))")
    
    strSql = "select a.����id,a.��ҳid,a.סԺ��,a.���� as ��������,d.���� as ����,c.ʹ������ as ������ҩ����" & _
        " from ������ҳ a,���˿����ؼ�¼ c,���ű� d" & _
        " where a.��Ժ����id =d.id and a.����id=c.����id and a.��ҳid=c.��ҳid" & _
        " And a.��Ժ���� Between " & strPar & strDept & _
        " and c.ʹ�ý׶�='����'" & IIf(str�п� = "", "", " and exists (select 1 from ���������¼ M where a.����id=m.����id and a.��ҳid=m.��ҳid and Instr([2],m.�п�)>0)")
    
    On Error GoTo errH
    Screen.MousePointer = 11
    Call zlCommFun.ShowFlash("���ڶ�ȡ���ݣ����Ժ�...")
    
    If optType(e_C4_optType_��������_���_22).Value Then '�������
        strSql = "select ����id,��ҳid,סԺ��,��������,����,������ҩ���� from (" & strSql & " Order By Dbms_Random.Value) where rownum < ([3] + 1) Order By ������ҩ���� Desc"
    Else
        strSql = "Select m.����id, m.��ҳid, m.סԺ��, m.��������, m.����, m.������ҩ����" & vbNewLine & _
            "From (Select m.����id, m.��ҳid, m.סԺ��, m.��������, m.����, m.������ҩ����" & vbNewLine & _
            "       From (Select m.����id, m.��ҳid, m.סԺ��, m.��������, m.����, m.������ҩ����, Mod(Rownum,[3]) M" & vbNewLine & _
            "              From (" & strSql & " Order By a.��Ժ���� desc) M" & vbNewLine & _
            "              Order By M) M" & vbNewLine & _
            "       Where Rownum <([3]+1)) M" & vbNewLine & _
            "Order By m.������ҩ���� Desc"
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, txtDept(e_C4_txtDept_ͳ�ƿ���_7).Tag, str�п�, lng������)
    
    If rsTmp.EOF Then
        Screen.MousePointer = 0
        Call zlCommFun.StopFlash
        MsgBox "��ǰ������δ�ҵ��κ����ݣ����������ó���ͳ�Ʋ�����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    With vsOpeKssUse
        .Rows = rsTmp.RecordCount + 1
        For i = 1 To rsTmp.RecordCount
            strPatis = strPatis & "," & rsTmp!����ID & ":" & rsTmp!��ҳID
            
            .TextMatrix(i, COL_OPE����id) = rsTmp!����ID
            .TextMatrix(i, COL_OPE��ҳid) = rsTmp!��ҳID
             
            .TextMatrix(i, COL_OPEסԺ��) = rsTmp!סԺ�� & ""
            .TextMatrix(i, COL_OPE��������) = rsTmp!�������� & ""
            .TextMatrix(i, COL_OPE����) = rsTmp!���� & ""
            .TextMatrix(i, COL_OPE������ҩ����) = rsTmp!������ҩ���� & ""
            rsTmp.MoveNext
        Next
        
        strPatis = Mid(strPatis, 2)

        strParTable = "Select C1, C2 From Table(f_Num2list2([1]))"
        strTable = strParTable
        
        If Len(strPatis) >= 4000 Then
            varArr = Array()
            varArr = GetParTable(strPar, strParTable, strTable)
        End If
        
        strSql = "select a.����id,a.��ҳid,a.�������� as ����,a.�п� from ���������¼ a where (a.����id,a.��ҳid) In (" & strTable & ")"
        
        If Len(strPatis) >= 4000 Then
            Set rs���� = zlDatabase.OpenSQLRecord(strSql, Me.Caption, CStr(varArr(0)), CStr(varArr(1)), CStr(varArr(2)), CStr(varArr(3)), CStr(varArr(4)), CStr(varArr(5)), _
                    CStr(varArr(6)), CStr(varArr(7)), CStr(varArr(8)), CStr(varArr(9)))
        Else
            Set rs���� = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strPatis)
        End If
 
        For i = .FixedRows To .Rows - 1
            strFilter = "����id=" & Val(.TextMatrix(i, COL_OPE����id)) & " And ��ҳid=" & Val(.TextMatrix(i, COL_OPE��ҳid))
        
            strTmp = "": strTmp1 = ""
            rs����.Filter = strFilter
            If Not rs����.EOF Then rs����.MoveFirst
            Do Until rs����.EOF
                strTmp = strTmp & ";" & rs����!����
                If rs����!�п� & "" <> "" And InStr("," & strTmp1 & ",", "," & rs����!�п� & ",") = 0 Then
                    strTmp1 = strTmp1 & "," & rs����!�п�
                End If
                rs����.MoveNext
            Loop
            .TextMatrix(i, COL_OPE��������) = Mid(strTmp, 2)
            .TextMatrix(i, COL_OPE�п�����) = Mid(strTmp1, 2)
        Next
    End With
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\���ڷ�Χ", "���󿹾�ҩ��ʹ�ó�N��ͳ��", dtpCountS(e_C4_dtpCountS_��ʼʱ��_4).Value & "," & dtpCountE(e_C4_dtpCountE_����ʱ��_4).Value
    Screen.MousePointer = 0
    Call zlCommFun.StopFlash
    Exit Sub
errH:
    Screen.MousePointer = 0
    Call zlCommFun.StopFlash
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadvsIllDruUse()
'���ܣ�����   ҽ������ĳ����������ҩ�ɱ�ͳ�� �����������
    Dim strSql As String, strPar As String
    Dim rsTmp As ADODB.Recordset
    
    Dim strParTable As String, strTable As String
    Dim varArr As Variant
    
    Dim rs���� As ADODB.Recordset
    Dim rs����ҩ�� As ADODB.Recordset
    Dim rs���ƽ�� As ADODB.Recordset
    
    Dim strPatis As String, strDec As String
    Dim strTmp As String, strFilter As String
    Dim lng������ As Long
    Dim lngTmp As Long
    Dim dblTmp As Double
    Dim i As Long
    
    lng������ = Val(txtNum(e_C5_txtNum_��������_3).Text)
    
    strTmp = IIf(optType(e_C5_optType_��ҽ_25).Value, 1, 0) & "|" & IIf(optType(e_C5_optType_������_28).Value, 1, 0) & "|" & txtILL.Tag & "|" & txtILL.Text
    Call zlDatabase.SetPara("���Ƽ�������", strTmp, glngSys, 1269)
    
    '���ڷ�Χ��������
    strPar = "To_Date('" & Format(dtpCountS(e_C5_dtpCountS_��ʼʱ��_5).Value, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')" & _
        " and To_Date('" & Format(dtpCountE(e_C5_dtpCountE_����ʱ��_5).Value, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
    strDec = "0.00"
    On Error GoTo errH
    Screen.MousePointer = 11
    Call zlCommFun.ShowFlash("���ڶ�ȡ���ݣ����Ժ�...")
    
    strSql = "select a.����id,a.��ҳid from ������ҳ a where a.��Ժ���� between " & strPar & "and exists (" & _
        "select 1 from ������ϼ�¼ b where b.����id=a.����id and b.��ҳid=a.��ҳid and b.��¼��Դ=3 and b.��ϴ���=1 And NVL(B.�������,1) = 1" & _
        " and b.�������=[1]" & IIf(optType(e_C5_optType_������_28).Value, " and b.����id=[2]", " and b.���id=[2]") & _
        ")"

    If optType(e_C5_optType_��������_���_19).Value Then ' �������
        strSql = "select ����id,��ҳid from (" & strSql & "order by Dbms_Random.Value) where rownum<([3] + 1)"
    Else
        strSql = "Select m.����id,m.��ҳid From (Select m.����id,m.��ҳid,Mod(Rownum,[3]) M From (" & strSql & " order by a.��Ժ����) m Order By M) m Where Rownum <([3] + 1)"
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(IIf(optType(e_C5_optType_��ҽ_25).Value, 3, 13)), txtILL.Tag, lng������)

    If rsTmp.EOF Then
       Screen.MousePointer = 0
       Call zlCommFun.StopFlash
       MsgBox "��ǰ������δ�ҵ��κ����ݣ����������ó���ͳ�Ʋ�����", vbInformation, gstrSysName
       Exit Sub
    End If
    
    For i = 1 To rsTmp.RecordCount
        strPatis = strPatis & "," & rsTmp!����ID & ":" & rsTmp!��ҳID
        rsTmp.MoveNext
    Next
     
    strPatis = Mid(strPatis, 2)
    
    strParTable = "Select C1, C2 From Table(f_Num2list2([3]))"
    strTable = strParTable
    
    If Len(strPatis) >= 4000 Then
        varArr = Array()
        varArr = GetParTable(strPar, strParTable, strTable)
    End If

    strSql = "select nvl(������,'��') as ����ҽ��,count(1) as ��������,Sum(sign(����ҩ����)) as ����ҩ����,sum(סԺ����) as סԺ����,sum(�ܷ���) as �ܷ���,sum(��ҩ��) as ��ҩ��,sum(����ҩ��) as ����ҩ��" & _
        " from (select b.������,b.����id,b.��ҳid,max(a.סԺ����) as סԺ����,sum(b.���ʽ��) as �ܷ���,sum(decode(Nvl(e.������, 0),0,0,b.���ʽ��)) as ����ҩ��," & _
        " sum(decode(b.�շ����,'5',b.���ʽ��,'6',b.���ʽ��,'7',b.���ʽ��,0)) as ��ҩ��,sum(decode(Nvl(e.������, 0),0,0,1)) as ����ҩ����" & _
        " from ������ҳ a,סԺ���ü�¼ b,ҩƷĿ¼ d,ҩƷ���� e" & _
        " where  a.����id=b.����id and a.��ҳid=b.��ҳid and b.��¼״̬<>0 and b.�շ�ϸĿid=d.ҩƷid(+) and d.ҩ��id=e.ҩ��id(+)" & _
        " and (a.����id,a.��ҳid) In (" & strTable & ")" & _
        " group by b.������,b.����id,b.��ҳid,a.��Ժ����,a.��Ժ����)" & _
        " group by ������ order by ��������"
    
    If Len(strPatis) >= 4000 Then
        Set rs���� = zlDatabase.OpenSQLRecord(strSql, Me.Caption, "", "", CStr(varArr(0)), CStr(varArr(1)), CStr(varArr(2)), CStr(varArr(3)), CStr(varArr(4)), CStr(varArr(5)), _
            CStr(varArr(6)), CStr(varArr(7)), CStr(varArr(8)), CStr(varArr(9)))
    Else
        Set rs���� = zlDatabase.OpenSQLRecord(strSql, Me.Caption, "", "", strPatis)
    End If
    
    '��Ժ���
    strSql = "select /*+ rule*/ nvl(������,'��') as ����ҽ��,sum(sign(����)) as ����,sum(sign(��ת)) as ��ת,sum(sign(δ��)) as δ��,sum(sign(����)) as ����,sum(sign(����)) as ����" & _
        " from (select a.������,a.����id,a.��ҳid,sum(decode(��Ժ���,'����',1,0)) as ����,sum(decode(��Ժ���,'��ת',1,0)) as ��ת," & _
        " sum(decode(��Ժ���,'δ��',1,0)) as δ��,sum(decode(��Ժ���,'����',1,0)) as ����," & _
        " sum(decode(��Ժ���,'����',0,'��ת',0,'����',0,'δ��',0,1)) as ����" & _
        " from סԺ���ü�¼ a,������ϼ�¼ b where a.����id=b.����id and a.��¼״̬<>0 and a.��ҳid=b.��ҳid and b.��¼��Դ=3 and b.��ϴ���=1  And NVL(B.�������,1) = 1" & _
        " and b.�������=[1]" & IIf(optType(e_C5_optType_������_28).Value, " and b.����id=[2]", " and b.���id=[2]") & _
        " and (a.����id,a.��ҳid) In (" & strTable & ")" & _
        " group by a.������,a.����id,a.��ҳid) group by ������"
    If Len(strPatis) >= 4000 Then
        Set rs���ƽ�� = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(IIf(optType(e_C5_optType_��ҽ_25).Value, 3, 13)), txtILL.Tag, CStr(varArr(0)), CStr(varArr(1)), CStr(varArr(2)), _
                CStr(varArr(3)), CStr(varArr(4)), CStr(varArr(5)), CStr(varArr(6)), CStr(varArr(7)), CStr(varArr(8)), CStr(varArr(9)))
    Else
        Set rs���ƽ�� = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(IIf(optType(e_C5_optType_��ҽ_25).Value, 3, 13)), txtILL.Tag, strPatis)
    End If
    
    strSql = "select nvl(������,'��') as ����ҽ��,count(1) as ����ҩ����" & _
        " from (select b.������,d.ҩ��ID" & _
        " from סԺ���ü�¼ b,ҩƷĿ¼ d,ҩƷ���� e" & _
        " where b.�շ�ϸĿid=d.ҩƷid(+) and d.ҩ��id=e.ҩ��id(+)" & _
        " and Nvl(e.������, 0)<>0 and b.�շ����='5'" & _
        " and (b.����id,b.��ҳid) In (" & strTable & ")" & _
        " group by b.������,d.ҩ��ID ) group by ������"
        
    If Len(strPatis) >= 4000 Then
        Set rs����ҩ�� = zlDatabase.OpenSQLRecord(strSql, Me.Caption, "", "", CStr(varArr(0)), CStr(varArr(1)), CStr(varArr(2)), CStr(varArr(3)), CStr(varArr(4)), CStr(varArr(5)), _
            CStr(varArr(6)), CStr(varArr(7)), CStr(varArr(8)), CStr(varArr(9)))
    Else
        Set rs����ҩ�� = zlDatabase.OpenSQLRecord(strSql, Me.Caption, "", "", strPatis)
    End If
    
    With vsIllDruUse
        .Rows = 2
        For i = 1 To rs����.RecordCount
            .AddItem ""
            
            strFilter = "����ҽ��='" & rs����!����ҽ�� & "'"
            rs���ƽ��.Filter = strFilter
            rs����ҩ��.Filter = strFilter
            
            .TextMatrix(.Rows - 1, COL_ILL����ҽ��) = rs����!����ҽ�� & ""
            .TextMatrix(.Rows - 1, COL_ILL��������) = Val(rs����!�������� & "")
            .TextMatrix(.Rows - 1, COL_ILL�ÿ�ҩ����) = Val(rs����!����ҩ���� & "")
            
            .TextMatrix(.Rows - 1, COL_ILL����) = 0
            .TextMatrix(.Rows - 1, COL_ILL��ת) = 0
            .TextMatrix(.Rows - 1, COL_ILLδ��) = 0
            .TextMatrix(.Rows - 1, COL_ILL����) = 0
            .TextMatrix(.Rows - 1, COL_ILL����) = 0
            
            strTmp = "": lngTmp = 0
            If Not rs���ƽ��.EOF Then rs���ƽ��.MoveFirst
            Do Until rs���ƽ��.EOF
                lngTmp = Val(rs���ƽ��!���� & "")
                .TextMatrix(.Rows - 1, COL_ILL����) = Val(rs���ƽ��!���� & "")
                .TextMatrix(.Rows - 1, COL_ILL��ת) = Val(rs���ƽ��!��ת & "")
                .TextMatrix(.Rows - 1, COL_ILLδ��) = Val(rs���ƽ��!δ�� & "")
                .TextMatrix(.Rows - 1, COL_ILL����) = Val(rs���ƽ��!���� & "")
                .TextMatrix(.Rows - 1, COL_ILL����) = Val(rs���ƽ��!���� & "")
                rs���ƽ��.MoveNext
            Loop
            
            If Val(rs����!�������� & "") <> 0 Then dblTmp = lngTmp * 100 / Val(rs����!�������� & "")
 
            .TextMatrix(.Rows - 1, COL_ILL������) = Format(dblTmp, strDec) & "%": dblTmp = 0
            
            dblTmp = Val(rs����!�ܷ��� & "")
            .TextMatrix(.Rows - 1, COL_ILL�ܽ��) = Format(dblTmp, strDec): dblTmp = 0
            
            dblTmp = Val(rs����!��ҩ�� & "")
            .TextMatrix(.Rows - 1, COL_ILLҩƷ���) = Format(dblTmp, strDec): dblTmp = 0
            
            If Val(rs����!�������� & "") <> 0 Then dblTmp = Val(rs����!�ܷ��� & "") / Val(rs����!�������� & "")
            .TextMatrix(.Rows - 1, COL_ILL�˾����ƶ�) = Format(dblTmp, strDec): dblTmp = 0
            
            If Val(rs����!סԺ���� & "") <> 0 Then dblTmp = Val(.TextMatrix(.Rows - 1, COL_ILL�˾����ƶ�)) / Val(rs����!סԺ���� & "")
            .TextMatrix(.Rows - 1, COL_ILL�˾��ս��) = Format(dblTmp, strDec): dblTmp = 0
            
            dblTmp = Val(rs����!����ҩ�� & "")
            .TextMatrix(.Rows - 1, COL_ILL��ҩ���) = Format(dblTmp, strDec): dblTmp = 0
            
            If Not rs����ҩ��.EOF Then rs����ҩ��.MoveFirst: lngTmp = 0
            Do Until rs����ҩ��.EOF
                lngTmp = Val(rs����ҩ��!����ҩ���� & "")
                rs����ҩ��.MoveNext
            Loop
            .TextMatrix(.Rows - 1, COL_ILL��ҩƷ����) = lngTmp
            rs����.MoveNext
        Next
    End With
    
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\���ڷ�Χ", "ҽ������ĳ����������ҩ�ɱ�ͳ��", dtpCountS(e_C5_dtpCountS_��ʼʱ��_5).Value & "," & dtpCountE(e_C5_dtpCountE_����ʱ��_5).Value
    
    Screen.MousePointer = 0
    Call zlCommFun.StopFlash
    Exit Sub
errH:
    Screen.MousePointer = 0
    Call zlCommFun.StopFlash
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdILL_Click()
'���ܣ����ѡ���������ý���   ҽ������ĳ����������ҩ�ɱ�ͳ��
    Dim rsTmp As ADODB.Recordset
     
    If Not optType(e_C5_optType_��ҽ_25).Value Then
        If optType(e_C5_optType_�����_27).Value Then
            '���������:��ҽ���ݣ�һ����Ͽ������ڶ������
            Set rsTmp = zlDatabase.ShowILLSelect(Me, "2", 0, , True, False)
        Else
            'B-��ҽ��������
            Set rsTmp = zlDatabase.ShowILLSelect(Me, "B", 0, , True)
        End If
    Else
        If optType(e_C5_optType_�����_27).Value Then
            '���������:��ҽ���ݣ�һ����Ͽ������ڶ������
            Set rsTmp = zlDatabase.ShowILLSelect(Me, "1", 0, , True, False)
        Else
            'D-ICD-10��������
            Set rsTmp = zlDatabase.ShowILLSelect(Me, "D", 0, , True)
        End If
    End If
     If Not rsTmp Is Nothing Then
        txtILL.Text = "(" & rsTmp!���� & ")" & rsTmp!����
        cmdILL.Tag = txtILL.Text
        txtILL.Tag = rsTmp!��ĿID
    End If
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    Dim i As Long, strTmp As String
    
    Select Case Control.ID
    Case conMenu_View_ToolBar_Button '������
        For i = 2 To cbsMain.Count
            Me.cbsMain(i).Visible = Not Me.cbsMain(i).Visible
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Text '��ť����
        For i = 2 To cbsMain.Count
            For Each objControl In Me.cbsMain(i).Controls
                objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Size '��ͼ��
        Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
        Me.cbsMain.RecalcLayout
    Case conMenu_View_StatusBar '״̬��
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsMain.RecalcLayout
    Case conMenu_View_Refresh
        If GetCurҳ�� = "���˿�����ҩ����������鼰���۱�" Then Call LoadPati
    Case conMenu_Tool_Archive '���Ӳ�������
        Call Show���Ӳ�������
    Case conMenu_Help_Help '����
        Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_File_Exit '�˳�
        Unload Me
    Case conMenu_File_Print
        Call zlRptPrint(1)
    Case conMenu_File_Preview
        Call zlRptPrint(2)
    Case conMenu_File_Excel
        Call zlRptPrint(3)
    Case conMenu_Help_Web_Home 'Web�ϵ�����
        Call zlHomePage(Me.hwnd)
    Case conMenu_Help_Web_Forum '������̳
        Call zlWebForum(Me.hwnd)
    Case conMenu_Help_Web_Mail '���ͷ���
        Call zlMailTo(Me.hwnd)
    Case conMenu_Help_About '����
        Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case conMenu_File_PrintSet
        SwitchPrintSet glngSys & "\" & 1269
        Call zlPrintSet
        SwitchPrintSet glngSys & "\" & 1269, True
    Case Else
        If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            'ִ�з�������ǰģ��ı���
            strTmp = Split(Control.Parameter, ",")(1)
                Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), strTmp, Me)
        End If
    End Select
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_View_ToolBar_Button '������
        If cbsMain.Count >= 2 Then
            Control.Checked = Me.cbsMain(2).Visible
        End If
    Case conMenu_View_ToolBar_Text 'ͼ������
        If cbsMain.Count >= 2 Then
            Control.Checked = Not (Me.cbsMain(2).Controls(GetFirstCommandBar(cbsMain(2).Controls)).Style = xtpButtonIcon)
        End If
    Case conMenu_View_ToolBar_Size '��ͼ��
        Control.Checked = Me.cbsMain.Options.LargeIcons
    Case conMenu_View_StatusBar '״̬��
        Control.Checked = Me.stbThis.Visible
    Case conMenu_Tool_Archive
        Control.Enabled = (tbcSub.Selected.Index = 0 And tbcReport.Selected.Index = 1)
    End Select
End Sub

Private Sub cmdOK_Click(Index As Integer)
'���ܣ���� ͳ�� ��ť
    
    If Not CheckData() Then Exit Sub
    
    If MsgBox("����������ǳ���ʱ�����ҿ���Ӱ��ϵͳ���������ܣ�������ҵ�����ʱ�����У���ȷ��Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    If PanelItem_����ҩƷ���Ľ������ = Index Then
        
        Call LoadBill
        
    ElseIf PanelItem_���˿�����ҩ����������鼰���۱� = Index Then
    
        If Save������¼ Then
            Call LoadPati
        Else
            rptPati.Records.DeleteAll
            rptPati.Populate
            txtCYJL.Tag = ""
            txtCYJL.Text = ""
        End If
        
    ElseIf PanelItem_���ﴦ��������ҩ����� = Index Then
    
        Call Load��������(vsMZYY, True)
        Call Load��������(vsMZYY, vsCF, True)
        
    ElseIf PanelItem_סԺ���˿�����ҩ����� = Index Then
    
        Call LoadInKssAdvice
        
    ElseIf PanelItem_����ҩ��ʹ���������ͳ�� = (Index - 4) Then
    
        Call LoadvsUseRan
        
    ElseIf PanelItem_�����п�Χ����Ԥ����ҩͳ�� = (Index - 4) Then
    
        Call LoadvsCut
        
    ElseIf PanelItem_�ż��ﴦ��������ҩͳ�� = (Index - 4) Then
    
        Call Load��������(vsCountDruUse, False)
        Call Load��������(vsCountDruUse, vsCountCF, False)
        
    ElseIf PanelItem_סԺҽ��������ҩͳ�� = (Index - 4) Then
    
        Call LoadvsInDruUse
        
    ElseIf PanelItem_���󿹾�ҩ��ʹ�ó�N��ͳ�� = (Index - 4) Then
    
        Call LoadvsOpeKssUse
        
    ElseIf PanelItem_ҽ������ĳ����������ҩ�ɱ�ͳ�� = (Index - 4) Then
    
        Call LoadvsIllDruUse
    End If
End Sub

Private Function CheckData() As Boolean
'���ܣ���������ǰ�ļ��

    Select Case GetCurҳ��
        Case "����ҩƷ���Ľ������"
            If dtpRPS(e_R0_dtpRPS_��ʼʱ��_0).Value >= dtpRPE(e_R0_dtpRPE_����ʱ��_0).Value Then
                MsgBox "��ʼʱ��Ӧ��С�ڽ���ʱ�䡣", vbInformation, gstrSysName
                dtpRPE(e_R0_dtpRPE_����ʱ��_0).SetFocus
                Exit Function
            End If
        Case "���˿�����ҩ����������鼰���۱�"
            '����ʱ���һ����飬��ʼʱ��С�ڽ���ʱ��
            If dtpRPS(e_R1_dtpRPS_��ʼʱ��_1).Value >= dtpRPE(e_R1_dtpRPE_����ʱ��_1).Value Then
                MsgBox "��ʼʱ��Ӧ��С�ڽ���ʱ�䡣", vbInformation, gstrSysName
                dtpRPE(e_R1_dtpRPE_����ʱ��_1).SetFocus
                Exit Function
            End If
    
            If Val(txtCount(e_R1_txtCount_��������_1).Text) = 0 Then
                MsgBox "������������Ϊ�㡣", vbInformation, gstrSysName
                txtCount(e_R1_txtCount_��������_1).SetFocus
                Exit Function
            End If
        Case "���ﴦ��������ҩ�����"
            If dtpRPS(e_R2_dtpRPS_��ʼʱ��_2).Value > dtpRPE(e_R2_dtpRPE_����ʱ��_2).Value Then
                MsgBox "��ʼʱ��С�ڽ���ʱ�䡣", vbInformation, gstrSysName
                dtpRPE(e_R2_dtpRPE_����ʱ��_2).SetFocus
                Exit Function
            End If
            
            If Val(txtCount(e_R2_txtCount_��������_2).Text) = 0 Then
                MsgBox "������������Ϊ�㡣", vbInformation, gstrSysName
                txtCount(e_R2_txtCount_��������_2).SetFocus
                Exit Function
            End If
        Case "סԺ���˿�����ҩ�����"
            If dtpRPS(e_R3_dtpRPS_��ʼʱ��_3).Value > dtpRPE(e_R3_dtpRPE_����ʱ��_3).Value Then
                MsgBox "��ʼʱ��С�ڽ���ʱ�䡣", vbInformation, gstrSysName
                dtpRPE(e_R3_dtpRPE_����ʱ��_3).SetFocus
                Exit Function
            End If
        Case "����ҩ��ʹ���������ͳ��"
            If dtpCountS(e_C0_dtpCountS_��ʼʱ��_0).Value > dtpCountE(e_C0_dtpCountE_����ʱ��_0).Value Then
                MsgBox "��ʼʱ��С�ڽ���ʱ�䡣", vbInformation, gstrSysName
                dtpCountE(e_C0_dtpCountE_����ʱ��_0).SetFocus
                Exit Function
            ElseIf Not IsNumeric(txtTopRan.Text) Then
                MsgBox "ͳ�����α����Ǵ������������", vbInformation, gstrSysName
                txtTopRan.SetFocus
                Exit Function
            ElseIf Val(txtTopRan.Text) <= 0 Then
                MsgBox "ͳ�����α����Ǵ������������", vbInformation, gstrSysName
                txtTopRan.SetFocus
                Exit Function
            End If
        Case "�����п�Χ����Ԥ����ҩͳ��"
            If dtpCountS(e_C1_dtpCountS_��ʼʱ��_1).Value > dtpCountE(e_C1_dtpCountE_����ʱ��_1).Value Then
                MsgBox "��ʼʱ��С�ڽ���ʱ�䡣", vbInformation, gstrSysName
                dtpCountE(e_C1_dtpCountE_����ʱ��_1).SetFocus
                Exit Function
            End If
        Case "�ż��ﴦ��������ҩͳ��"
            If dtpCountS(e_C2_dtpCountS_��ʼʱ��_2).Value > dtpCountE(e_C2_dtpCountE_����ʱ��_2).Value Then
                MsgBox "��ʼʱ��С�ڽ���ʱ�䡣", vbInformation, gstrSysName
                dtpCountE(e_C2_dtpCountE_����ʱ��_2).SetFocus
                Exit Function
            End If
            
            If Val(txtNum(e_C2_txtNum_ͳ�ƿ���_0).Text) = 0 Then
                MsgBox "������������Ϊ�㡣", vbInformation, gstrSysName
                txtNum(e_C2_txtNum_ͳ�ƿ���_0).SetFocus
                Exit Function
            End If
        Case "סԺҽ��������ҩͳ��"
            If dtpCountS(e_C3_dtpCountS_��ʼʱ��_3).Value > dtpCountE(e_C3_dtpCountE_����ʱ��_3).Value Then
                MsgBox "��ʼʱ��С�ڽ���ʱ�䡣", vbInformation, gstrSysName
                dtpCountE(e_C3_dtpCountE_����ʱ��_3).SetFocus
                Exit Function
            End If
            If Val(txtNum(e_C3_txtNum_��������_1).Text) = 0 Then
                MsgBox "������������Ϊ�㡣", vbInformation, gstrSysName
                txtNum(e_C3_txtNum_��������_1).SetFocus
                Exit Function
            End If
            
            If optType(e_C3_optType_�п�����_����_18).Value Then
                If chkType(e_C3_chkType_�п�����_����_2).Value <> 1 And chkType(e_C3_chkType_�п�����_����_3).Value <> 1 And _
                    chkType(e_C3_chkType_�п�����_����_4).Value <> 1 And chkType(e_C3_chkType_�п�����_����_8).Value <> 1 Then
                    MsgBox "��ѡ��һ���п����͡�", vbInformation, gstrSysName
                End If
            End If
            
        Case "���󿹾�ҩ��ʹ�ó�N��ͳ��"
            If Val(txtNum(e_C4_txtNum_��������_2).Text) = 0 Then
                MsgBox "������������Ϊ�㡣", vbInformation, gstrSysName
                txtNum(e_C4_txtNum_��������_2).SetFocus
                Exit Function
            End If
            
        Case "ҽ������ĳ����������ҩ�ɱ�ͳ��"
            If dtpCountS(e_C5_dtpCountS_��ʼʱ��_5).Value > dtpCountE(e_C5_dtpCountE_����ʱ��_5).Value Then
                MsgBox "��ʼʱ��С�ڽ���ʱ�䡣", vbInformation, gstrSysName
                dtpCountE(e_C5_dtpCountE_����ʱ��_5).SetFocus
                Exit Function
            End If
            
            If Val(txtNum(e_C5_txtNum_��������_3).Text) = 0 Then
                MsgBox "������������Ϊ�㡣", vbInformation, gstrSysName
                txtNum(e_C5_txtNum_��������_3).SetFocus
                Exit Function
            End If
            
            If txtILL.Text = "" Then
                MsgBox "��ѡ��һ�ּ�����", vbInformation, gstrSysName
                txtILL.SetFocus
                Exit Function
            End If
    End Select
    
    CheckData = True
End Function

Private Sub cmdCYDel_Click()
    Dim blnTrans As Boolean
    
    If txtCYJL.Tag = "" Then Exit Sub
    If MsgBox("��ȷ��Ҫɾ�����γ�����¼��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    End If
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    Call zlDatabase.ExecuteProcedure("Zl_����ҩ�������¼_Delete(" & txtCYJL.Tag & ")", Me.Caption)
    gcnOracle.CommitTrans: blnTrans = False
    txtCYJL.Tag = ""
    txtCYJL.Text = ""
    rptPati.Records.DeleteAll
    rptPati.Populate
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdCYEdit_Click()
    
    If rptPati.SelectedRows.Count = 0 Then
        MsgBox "δѡ����һ�����ˣ�", vbInformation, gstrSysName
        Exit Sub
    End If

    If rptPati.SelectedRows(0).GroupRow Then
        MsgBox "δѡ����һ�����ˣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If rptPati.SelectedRows.Count > 0 Then Call Edit���˵���
  
End Sub

Private Sub Edit���˵���()
'���ܣ�������������

    Dim bln�༭ As Boolean
    Dim bln��ӡ As Boolean
    Dim blnTmp As Boolean
    
    With rptPati.SelectedRows(0)
        If Not .GroupRow Then
            If .Record(COL_�༭).Value = "��" Then bln�༭ = True
            If .Record(COL_��ӡ).Value = "��" Then bln��ӡ = True
            blnTmp = frmKssSurveyEdit.ShowMe(Me, .Record(COL_����ID).Value, .Record(col_����Id).Value, .Record(col_��ҳID).Value, .Record(COL_���).Value, _
                IIf(Val(.Record(COL_����ID).Value) = 0, mlng������������, mlng����������), .Record(col_����).Value, Val(.Record(COL_����ID).Value) > 0, bln�༭, bln��ӡ)
            If blnTmp Then
                If bln�༭ Then .Record(COL_�༭).Value = "��"
                If bln��ӡ Then .Record(COL_��ӡ).Value = "��"
                rptPati.Populate
            End If
        End If
    End With
End Sub

Private Sub Show���Ӳ�������()
    If rptPati.SelectedRows.Count = 0 Then
        MsgBox "δѡ����һ�����ˣ�", vbInformation, gstrSysName
        Exit Sub
    End If

    If rptPati.SelectedRows(0).GroupRow Then
        MsgBox "δѡ����һ�����ˣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If rptPati.SelectedRows.Count > 0 Then
        With rptPati.SelectedRows(0)
            If Not .GroupRow Then
                Call frmArchiveView.ShowArchive(Me, Val(.Record(col_����Id).Value), Val(.Record(col_��ҳID).Value))
            End If
        End With
    End If
End Sub

Private Function Save������¼() As Boolean
    Dim strSql As String
    Dim blnTrans As Boolean
    Dim strCurDate As String
    
    On Error GoTo errH
    
    mlng����ID = zlDatabase.GetNextId("����ҩ�������¼")
    mdatCurr = zlDatabase.Currentdate
    strCurDate = Format(mdatCurr, "yyyy-MM-dd hh:mm:ss")
    
    strSql = "Zl_����ҩ�������¼_Insert(" & mlng����ID & ",'" & UserInfo.���� & "'," & _
       "to_date('" & Format(dtpRPS(e_R1_dtpRPS_��ʼʱ��_1).Value, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')," & _
       "to_date('" & Format(dtpRPE(e_R1_dtpRPE_����ʱ��_1).Value, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')," & _
       Val(txtCount(e_R1_txtCount_��������_1).Text) & "," & IIf(optType(e_R1_optType_��������_ƽ��_0).Value, 0, 1) & "," & _
       IIf(txtDept(e_R1_txtDept_��������_1).Tag = "", "NULL,", "'" & txtDept(e_R1_txtDept_��������_1).Tag & "',") & _
       "to_date('" & strCurDate & "','YYYY-MM-DD HH24:MI:SS'))"
    
    gcnOracle.BeginTrans: blnTrans = True
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    gcnOracle.CommitTrans: blnTrans = False
    
    MsgBox "���γ�����ɣ�", vbInformation, gstrSysName
    
    txtCYJL.Text = "����ʱ�䣺" & strCurDate & "  �����ˣ�" & UserInfo.����
    txtCYJL.Tag = mlng����ID
    
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\���ڷ�Χ", "���˿�����ҩ����������鼰���۱�", dtpRPS(e_R1_dtpRPS_��ʼʱ��_1).Value & "," & dtpRPE(e_R1_dtpRPE_����ʱ��_1).Value
    
    Save������¼ = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub LoadBill()
'���ܣ�����   ����ҩƷ���Ľ������   �����ݣ���ѯ��ʱ�������Զ�������������ص������ϣ��ý������ʾ
    Dim strSql As String, strPar As String
    Dim rsTmp As ADODB.Recordset
    Dim dblTmp As Double
    Dim strDec As String '���õľ��ȣ�4λС��
    Dim i As Long
    
    Dim dblTotal As Double ' "һ����ҽԺ�����루��"
    Dim dbl��� As Double  '"�塢ҩƷ����������루��"
    Dim dbl��ҩ�� As Double
    Dim dblסԺ��ҩ�� As Double
    Dim dbl������ҩ�� As Double
    Dim dblסԺ��ҩ�ѿ� As Double
    Dim dbl������ҩ�ѿ� As Double
    
    On Error GoTo errH
    Screen.MousePointer = 11
    Call zlCommFun.ShowFlash("���ڶ�ȡ���ݣ����Ժ�...")
    
    strDec = "0.0000"
    
    '�������������
    With vsBill
        For i = 1 To 13
            .TextMatrix(i, COL_��λ) = "��Ԫ"
        Next
        .TextMatrix(ROW_��ҽԺ������, COL_��ע) = "������������"
        .TextMatrix(ROW_ҩƷռҽԺ���������, COL_��λ) = "%"
        .TextMatrix(ROW_ҩƷ�����������ռҽԺ���������, COL_��λ) = "%"
        .TextMatrix(ROW_����ҩ��ռҩƷ���������, COL_��λ) = "%"
        .Cell(flexcpText, 1, COL_���, 13, COL_���) = strDec 'δ���ɽ��ʱ��ֵ����Ϊ 0.00
    End With
    
    '���ڷ�Χ��������
    strPar = "To_Date('" & Format(dtpRPS(e_R0_dtpRPS_��ʼʱ��_0).Value, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')" & _
        " and To_Date('" & Format(dtpRPE(e_R0_dtpRPE_����ʱ��_0).Value, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
    
    '�ӻ��ܱ��в�ѯ��ҽԺ�ܷ���  "һ����ҽԺ�����루��"
    strSql = "select sum(a.���ʽ��)/10000 as ������ from ���˷��û��� a  where a.���� between " & strPar
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    If Not rsTmp.EOF Then dblTotal = Val(rsTmp!������ & "")
    
    vsBill.TextMatrix(ROW_��ҽԺ������, COL_���) = Format(dblTotal, strDec)
    
    '���   ��ҩ��
    strSql = "select sum(a.���)/10000 as ���,-1*sum(decode(a.����,8,a.���,9,a.���,10,a.���,0))/10000 as ҩ�� from ҩƷ�շ����� a where a.����<14 and a.���� Between " & strPar
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    If Not rsTmp.EOF Then
        dbl��� = Val(rsTmp!��� & "")
        dbl��ҩ�� = Val(rsTmp!ҩ�� & "")
    End If
    vsBill.TextMatrix(ROW_ҩƷ�����������, COL_���) = Format(dbl���, strDec)
    vsBill.TextMatrix(ROW_��ҩƷ������, COL_���) = Format(dbl��ҩ��, strDec)
    
    'ҩƷ�����������ռҽԺ���������
    If dblTotal <> 0 Then
        dblTmp = dbl��� * 100 / dblTotal
        If dblTmp <> 0 Then
            vsBill.TextMatrix(ROW_ҩƷ�����������ռҽԺ���������, COL_���) = Format(dblTmp, strDec)
        End If
    End If
    
    'ҩƷռҽԺ���������
    If dblTotal <> 0 Then
        dblTmp = dbl��ҩ�� * 100 / dblTotal
        If dblTmp <> 0 Then
            vsBill.TextMatrix(ROW_ҩƷռҽԺ���������, COL_���) = Format(dblTmp, strDec)
        End If
    End If
    
    '������ü�¼   ��ҩ��  ����ҩ��
    strSql = "select sum(a.ҩ��)/10000 as ������ҩ��,Sum(Decode(Nvl(c.������, 0), 0, 0, a.ҩ��))/10000 As ������ҩ����ҩ��" & _
        " from (Select x.�շ�ϸĿid,Sum(x.���ʽ��) As ҩ�� From ������ü�¼ X Where x.����ʱ�� Between " & strPar & _
        " And x.��¼״̬ <> 0 And x.�շ����='5' group by x.�շ�ϸĿid) a," & _
        " ҩƷ��� B, ҩƷ���� C where a.�շ�ϸĿid = b.ҩƷid And b.ҩ��id = c.ҩ��id"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    
    If Not rsTmp.EOF Then
        dbl������ҩ�� = Val(rsTmp!������ҩ�� & "")
        dbl������ҩ�ѿ� = Val(rsTmp!������ҩ����ҩ�� & "")
    End If
    
    vsBill.TextMatrix(ROW_������ҩ��, COL_���) = Format(dbl������ҩ��, strDec)
    vsBill.TextMatrix(ROW_������ҩ����, COL_���) = Format(dbl������ҩ�ѿ�, strDec)
    
    'סԺ���ü�¼   ��ҩ��  ����ҩ��
    strSql = "select sum(a.ҩ��)/10000 as סԺ��ҩ��,Sum(Decode(Nvl(c.������, 0), 0, 0, a.ҩ��))/10000 As סԺ��ҩ����ҩ��" & _
        " from (Select x.�շ�ϸĿid,Sum(x.���ʽ��) As ҩ�� From סԺ���ü�¼ X Where x.����ʱ�� Between " & strPar & _
        " And x.��¼״̬ <> 0 And x.�շ����='5' group by x.�շ�ϸĿid) a," & _
        " ҩƷ��� B, ҩƷ���� C where a.�շ�ϸĿid = b.ҩƷid And b.ҩ��id = c.ҩ��id"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    If Not rsTmp.EOF Then
        dblסԺ��ҩ�� = Val(rsTmp!סԺ��ҩ�� & "")
        dblסԺ��ҩ�ѿ� = Val(rsTmp!סԺ��ҩ����ҩ�� & "")
    End If
    
    vsBill.TextMatrix(ROW_סԺ��ҩ��, COL_���) = Format(dblסԺ��ҩ��, strDec)
    vsBill.TextMatrix(ROW_סԺ��ҩ����, COL_���) = Format(dblסԺ��ҩ�ѿ�, strDec)
    
    '��ҩȫ��ʹ�ý��
    dblTmp = dblסԺ��ҩ�� + dbl������ҩ��
    vsBill.TextMatrix(ROW_��ҩȫ��ʹ�ý��, COL_���) = Format(dblTmp, strDec)
    
    '����ҩ��ȫ��ʹ�ý��
    dblTmp = dblסԺ��ҩ�ѿ� + dbl������ҩ�ѿ�
    vsBill.TextMatrix(ROW_����ҩ��ȫ��ʹ�ý��, COL_���) = Format(dblTmp, strDec)
    
    '����ҩ��ռҩƷ���������
    If dbl��ҩ�� <> 0 Then
        dblTmp = dblTmp * 100 / dbl��ҩ��
        If dblTmp <> 0 Then
            vsBill.TextMatrix(ROW_����ҩ��ռҩƷ���������, COL_���) = Format(dblTmp, strDec)
        End If
    End If
    
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\���ڷ�Χ", "����ҩƷ���Ľ������", dtpRPS(e_R0_dtpRPS_��ʼʱ��_0).Value & "," & dtpRPE(e_R0_dtpRPE_����ʱ��_0).Value
    
    Screen.MousePointer = 0
    Call zlCommFun.StopFlash
    Exit Sub
errH:
    Screen.MousePointer = 0
    Call zlCommFun.StopFlash
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function LoadPati()
'���ܣ��ӽس��������б�
    Dim rsTmp As ADODB.Recordset
    Dim objParent As ReportRecord
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim strTmp As String
    Dim objRow As ReportRow
    Dim lngCount As Long
    Dim strSql As String
    Dim i As Long
    
    rptPati.Records.DeleteAll
    
    strSql = "Select b.����id,b.���,b.����id,b.��ҳid,b.�Ƿ��ӡ,b.�Ƿ�༭,a.����,a.�Ա�,a.����,a.סԺ��,a.��Ժ����id,d.���� As ��Ժ����,a.סԺҽʦ,a.��Ժ����,max(e.Id) as ����id" & _
        " From ������ҳ A, ����ҩ�������ϸ B,���ű� D, ���������¼ E" & vbNewLine & _
        " Where a.����id = b.����id And a.��ҳid = b.��ҳid And b.����id =[1] And a.��Ժ����id = d.Id And a.����id = e.����id(+)" & _
        " And a.��ҳid = e.��ҳid(+)" & _
        " group by b.����id,b.���,b.����id,b.��ҳid,a.����,a.�Ա�,a.����,a.סԺ��,a.��Ժ����id,d.����,a.סԺҽʦ,a.��Ժ����,b.�Ƿ��ӡ,b.�Ƿ�༭ Order By b.���"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(txtCYJL.Tag))
    
    '��ʵ�ʳ���������
    lblN(e_R1_lblN_������������_13).Tag = rsTmp.RecordCount
    
    For i = 1 To rsTmp.RecordCount
        Set objRecord = Me.rptPati.Records.Add()
        objRecord.Tag = CStr(rsTmp!����ID & "," & rsTmp!��ҳID) '���ڲ��˶�λ
        
        If Val(rsTmp!����id & "") > 0 Then lngCount = lngCount + 1
        
        strTmp = IIf(Val(rsTmp!�Ƿ�༭ & "") = 0, "��", "��")
        objRecord.AddItem strTmp '�Ƿ�༭��������
        
        strTmp = IIf(Val(rsTmp!�Ƿ��ӡ & "") = 0, "��", "��")
        objRecord.AddItem strTmp '�Ƿ��ӡ��ͼ��
        
        Set objItem = objRecord.AddItem(IIf(Val(rsTmp!����id & "") > 0, "����", "������"))   '������Value��������
            objItem.Caption = IIf(Val(rsTmp!����id & "") > 0, "��������", "����������")
        
        objRecord.AddItem CStr(Nvl(rsTmp!����))
        objRecord.AddItem CStr(Nvl(rsTmp!�Ա�))
        objRecord.AddItem CStr(Nvl(rsTmp!����))
        objRecord.AddItem CStr(Nvl(rsTmp!סԺ��))
        objRecord.AddItem CStr(Nvl(rsTmp!��Ժ����))
        objRecord.AddItem CStr(Nvl(rsTmp!סԺҽʦ))
        objRecord.AddItem Format(rsTmp!��Ժ���� & "", "yyyy-mm-dd hh:mm:ss")
        
        objRecord.AddItem Val(rsTmp!����ID)
        objRecord.AddItem Val(rsTmp!��ҳID)
        objRecord.AddItem Val(rsTmp!����ID)
        objRecord.AddItem Val(rsTmp!���)
        objRecord.AddItem Val(rsTmp!����id & "")
        
        rsTmp.MoveNext
    Next
    
    mlng���������� = lngCount
    mlng������������ = rsTmp.RecordCount - lngCount
    
    rptPati.Populate
    
    With rptPati.Columns
        .Column(col_����).Width = 60
        .Column(col_�Ա�).Width = 30
        .Column(col_����).Width = 60
        .Column(col_סԺ��).Width = 70
        .Column(col_����).Width = 100
        .Column(col_סԺҽʦ).Width = 60
        .Column(col_��Ժ����).Width = 140
        .Column(col_��Ժ����).Alignment = xtpAlignmentCenter
        .Column(col_����).Alignment = xtpAlignmentLeft
    End With
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Load��������(ByRef vsgInfo As VSFlexGrid, ByVal blnRP As Boolean)
'���ܣ����ﴦ�������������Ǽ���
'������ vsgInfo �������ϱ����ݣ���vsMZYY  ����ͳ�� vsCountDruUse
'       blnRP ���棬true �ϱ����ݣ�false ����ͳ��
    Dim strSql As String, strPar As String
    Dim str���� As String, strDept  As String
    Dim lngBaseRow As Long, strDec As String
    Dim i As Long, j As Long, k As Long
    Dim strTableIn As String, strTableOut As String
    Dim varArr As Variant
    Dim rs���� As ADODB.Recordset
    Dim rs���� As ADODB.Recordset
    Dim rs������� As ADODB.Recordset
    Dim rsҩƷ���� As ADODB.Recordset
    Dim rs��� As ADODB.Recordset
    Dim rsҽ������ As ADODB.Recordset
    Dim rs���ô��� As ADODB.Recordset
    Dim rs���ÿ�ҩ��ϸ As ADODB.Recordset
    Dim rsҽ����ҩ��ϸ As ADODB.Recordset
    Dim strParҽ������ As String
    Dim strPar���ô��� As String
    Dim strDeptIDs As String
    Dim strTmp As String
    Dim lng�������� As Long
    Dim bln������ʽ As Boolean 'ƽ�����������, bln������ʽ true ƽ������ false �������
    Dim blnע�� As Boolean
    Dim lngTmp As Long
    Dim dblTmp As Double
    
    strDec = "0.00"
    
    If blnRP Then   '���ڷ�Χ��������
        strPar = "To_Date('" & Format(dtpRPS(e_R2_dtpRPS_��ʼʱ��_2).Value, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')" & _
            " and To_Date('" & Format(dtpRPE(e_R2_dtpRPE_����ʱ��_2).Value, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
        
        lblN(e_R2_lblN_�����_��������_27).Caption = "����������0��"
        lblN(e_R2_lblN_������_��������_30).Caption = lblN(e_R2_lblN_�����_��������_27).Caption
        lblN(e_R2_lblN_������_����_70).Caption = "0�Ŵ���ͳ�Ʒ�����"
        lblN(e_R2_lblN_�����_����_26).Caption = "���ڣ�" & Format(dtpRPS(e_R2_dtpRPS_��ʼʱ��_2).Value, "YYYY-MM-DD") & "��" & Format(dtpRPE(e_R2_dtpRPE_����ʱ��_2).Value, "YYYY-MM-DD")
        lblN(e_R2_lblN_������_����_29).Caption = lblN(e_R2_lblN_�����_����_26).Caption
    Else
        strPar = "To_Date('" & Format(dtpCountS(e_C2_dtpCountS_��ʼʱ��_2).Value, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')" & _
            " and To_Date('" & Format(dtpCountE(e_C2_dtpCountE_����ʱ��_2).Value, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
        
        lblN(e_C2_lblN_ͳ�Ʊ�_��������_48).Caption = "����������0��"
        lblN(e_C2_lblN_������_��������_49).Caption = lblN(e_C2_lblN_ͳ�Ʊ�_��������_48).Caption
        lblN(e_C2_lblN_������_����_39).Caption = "0�Ŵ���ͳ�Ʒ�����"
        lblN(e_C2_lblN_ͳ�Ʊ�_����_46).Caption = "���ڣ�" & Format(dtpCountS(e_C2_dtpCountS_��ʼʱ��_2).Value, "YYYY-MM-DD") & "��" & Format(dtpCountE(e_C2_dtpCountE_����ʱ��_2).Value, "YYYY-MM-DD")
        lblN(e_C2_lblN_������_����_47).Caption = lblN(e_C2_lblN_ͳ�Ʊ�_����_46).Caption
    End If
    
    strDeptIDs = IIf(blnRP, txtDept(e_R2_txtDept_��������_2).Tag, txtDept(e_C2_txtDept_ͳ�ƿ���_5).Tag)
    lng�������� = IIf(blnRP, Val(txtCount(e_R2_txtCount_��������_2).Text), Val(txtNum(e_C2_txtNum_ͳ�ƿ���_0).Text))
    bln������ʽ = IIf(blnRP, optType(e_R2_optType_��������_ƽ��_3).Value, optType(e_C2_optType_��������_ƽ��_6).Value)     ' ƽ������
    strDept = IIf(strDeptIDs = "", "", " and a.��������id in (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))")
    On Error GoTo errH
    Screen.MousePointer = 11
    Call zlCommFun.ShowFlash("���ڶ�ȡ���ݣ����Ժ�...")
    '���ԭ������
    vsgInfo.Rows = vsgInfo.FixedRows
    vsgInfo.Rows = vsgInfo.FixedRows + 1

    '����SQL��֯��ʽ��������������������ּ���ʱ������Ϊ1�ֻ࣬�鼱��ʱ����2�࣬���������Ϊ3�ࡣ�����ദ��ͳһ��Ϊ�ǷǼ��ﴦ����
    '����������ʹ�ã������ദ����Ϊ�������ң�ҽ���ദ����Ϊ ִ�в��� ���˹Һż�¼.ִ�в���id
    str���� = ""
    If Not blnRP Then
        If chkType(e_C2_chkType_����_����_0).Value = 1 And chkType(e_C2_chkType_����_����_1).Value = 0 Then        '3�� ��������
            str���� = " and (nvl(a.�Ƿ���,0)<>1 and (a.ҽ����� is null or exists (select 1 from ����ҽ����¼ B, ���˹Һż�¼ C where a.ҽ����� = b.Id and b.�Һŵ� = c.No and nvl(c.����,0)<>1)))"
        ElseIf chkType(e_C2_chkType_����_����_0).Value = 0 And chkType(e_C2_chkType_����_����_1).Value = 1 Then '2�� ֻ������
            str���� = " and (nvl(a.�Ƿ���,0)=1 or exists (select 1 from ����ҽ����¼ B, ���˹Һż�¼ C where a.ҽ����� = b.Id and  b.�Һŵ� = c.No and nvl(c.����,0)=1))"
        End If
    End If
 
    If bln������ʽ Then 'ƽ������
        strSql = "select  a.��ʶ��,a.ҽ��,a.�����,a.��������,a.��������,a.����ҽ��,a.����id,a.����,a.ҩ��" & vbNewLine & _
            "from (select a.��ʶ��,a.ҽ��,a.�����,a.��������,a.��������,a.����ҽ��,a.����id,a.����,a.ҩ��,Mod(Rownum,[2]) M" & vbNewLine & _
            "from (Select a.No As ��ʶ��, Decode(Nvl(Max(a.ҽ�����), 0), 0, 0, 1) As ҽ��, a.��ʶ�� As �����, a.���� As ��������," & vbNewLine & _
            "       To_Char(Min(a.����ʱ��), 'YYYY-MM-DD HH24:MI:SS') As ��������, a.������ As ����ҽ��, a.��������id As ����id, a.����,sum(a.���ʽ��) as ҩ��" & vbNewLine & _
            "From ������ü�¼ A" & vbNewLine & _
            "Where a.��¼״̬ <> 0 And a.�շ���� In ('5','6','7') And" & vbNewLine & _
            "      a.����ʱ�� Between " & strPar & strDept & str���� & vbNewLine & _
            "Group By a.No, a.��ʶ��, a.����, a.������, a.��������id, a.���� having sum(a.���ʽ��)>0 order by Min(a.����ʱ��) desc) a" & vbNewLine & _
            "order by M) a where rownum<([2]+1)"
    Else
        strSql = "select  a.��ʶ��,a.ҽ��,a.�����,a.��������,a.��������,a.����ҽ��,a.����id,a.����,a.ҩ��" & vbNewLine & _
            "from (select a.��ʶ��,a.ҽ��,a.�����,a.��������,a.��������,a.����ҽ��,a.����id,a.����,a.ҩ��" & vbNewLine & _
            "from (select a.��ʶ��,a.ҽ��,a.�����,a.��������,a.��������,a.����ҽ��,a.����id,a.����,a.ҩ��" & vbNewLine & _
            "from (Select a.No As ��ʶ��, Decode(Nvl(Max(a.ҽ�����), 0), 0, 0, 1) As ҽ��, a.��ʶ�� As �����, a.���� As ��������," & vbNewLine & _
            "       To_Char(Min(a.����ʱ��), 'YYYY-MM-DD HH24:MI:SS') As ��������, a.������ As ����ҽ��, a.��������id As ����id, a.����,sum(a.���ʽ��) as ҩ��" & vbNewLine & _
            "From ������ü�¼ A" & vbNewLine & _
            "Where a.��¼״̬ <> 0 And a.�շ���� In ('5','6','7') And" & vbNewLine & _
            "      a.����ʱ�� Between " & strPar & strDept & str���� & vbNewLine & _
            "Group By a.No, a.��ʶ��, a.����, a.������, a.��������id, a.���� having sum(a.���ʽ��)>0 ) a" & vbNewLine & _
            "order by Dbms_Random.Value) a where rownum<([2]+1)) a order by a.�������� desc"
    End If
    
    Set rs���� = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strDeptIDs, lng��������)

    If rs����.EOF Then
        Screen.MousePointer = 0
        Call zlCommFun.StopFlash
        MsgBox "��ǰ������δ�ҵ��κ����ݣ����������ó���ͳ�Ʋ�����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If blnRP Then
        lblN(e_R2_lblN_�����_��������_27).Caption = "����������" & rs����.RecordCount & "��"
        lblN(e_R2_lblN_������_��������_30).Caption = lblN(e_R2_lblN_�����_��������_27).Caption
        lblN(e_R2_lblN_������_����_70).Caption = rs����.RecordCount & "�Ŵ���ͳ�Ʒ�����"
    Else
        lblN(e_C2_lblN_ͳ�Ʊ�_��������_48).Caption = "����������" & rs����.RecordCount & "��"
        lblN(e_C2_lblN_������_��������_49).Caption = lblN(e_C2_lblN_ͳ�Ʊ�_��������_48).Caption
        lblN(e_C2_lblN_������_����_39).Caption = rs����.RecordCount & "�Ŵ���ͳ�Ʒ�����"
    End If
    
    '������ϸSQL��������ɺ����ݱ���൱��
    '�����ռ�
    strPar = "": strDeptIDs = ""
    For i = 1 To rs����.RecordCount
        If Val(rs����!ҽ�� & "") = 0 Then
            strPar���ô��� = strPar���ô��� & "," & rs����!��ʶ��
        Else
            strParҽ������ = strParҽ������ & "," & rs����!��ʶ��
        End If
        strPar = strPar & "," & rs����!��ʶ��
        If InStr("," & strDeptIDs & ",", "," & rs����!����ID & ",") = 0 Then
            strDeptIDs = strDeptIDs & "," & rs����!����ID
        End If
        rs����.MoveNext
    Next
    rs����.MoveFirst
    
    strDeptIDs = Mid(strDeptIDs, 2) '����id���ǲ��ᳬ����������
    strSql = "select id as ����id,���� as ���� from ���ű� where id in (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))"
    Set rs���� = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strDeptIDs)

    'NO �ϳɵĲ�������Ҫ��������Ҫ������
    strPar = Mid(strPar, 2)
    strTableIn = "Select Column_Value From Table(Cast(f_Str2list([1]) As zlTools.t_Strlist))"
    varArr = Array()
    varArr = GetParTable(strPar, strTableIn, strTableOut)
    strSql = "select a.No As ��ʶ��,Sum(a.���ʽ��) As ������� From ������ü�¼ A where a.��¼״̬ <> 0 And a.no in (" & strTableOut & ") Group By a.No"
    Set rs������� = zlDatabase.OpenSQLRecord(strSql, Me.Caption, CStr(varArr(0)), CStr(varArr(1)), CStr(varArr(2)), CStr(varArr(3)), CStr(varArr(4)), CStr(varArr(5)), _
            CStr(varArr(6)), CStr(varArr(7)), CStr(varArr(8)), CStr(varArr(9)))
    
    strSql = "Select a.��ʶ��, Sum(Sign(a.����ҩ)) As ����ҩ����, Sum(Sign(a.����ҩ)) As ����ҩ����, Sum(Sign(a.ҩƷ)) As ҩƷ����, Sum(a.����ҩ��) As ����ҩ��" & vbNewLine & _
        "From (Select a.��ʶ��, a.ҩ��id, Sum(a.����ҩ) As ����ҩ, Sum(a.����ҩ) As ����ҩ, Sum(a.ҩƷ) As ҩƷ, Sum(a.����ҩ��) As ����ҩ��" & vbNewLine & _
        "       From (Select a.No As ��ʶ��, c.ҩ��id, Decode(Nvl(b.����ҩ��, '0'), '0', 0, 1) As ����ҩ, Decode(Nvl(c.������, 0), 0, 0, 1) As ����ҩ," & vbNewLine & _
        "                     Decode(a.�շ����, '5', 1, '6', 1, '7', 1, 0) As ҩƷ, Sum(Decode(Nvl(c.������, 0), 0, 0, a.���ʽ��)) As ����ҩ��" & vbNewLine & _
        "              From ������ü�¼ A, ҩƷ��� B, ҩƷ���� C" & vbNewLine & _
        "              Where a.�շ�ϸĿid = b.ҩƷid And b.ҩ��id = c.ҩ��id And a.��¼״̬ <> 0 And a.�շ���� In ('5', '6', '7') and a.no in (" & strTableOut & ")" & vbNewLine & _
        "              Group By a.No, c.ҩ��id, a.�շ����, Nvl(b.����ҩ��, '0'), Nvl(c.������, 0)) A" & vbNewLine & _
        "       Group By a.��ʶ��, a.ҩ��id) A" & vbNewLine & _
        "Group By a.��ʶ��"
    Set rsҩƷ���� = zlDatabase.OpenSQLRecord(strSql, Me.Caption, CStr(varArr(0)), CStr(varArr(1)), CStr(varArr(2)), CStr(varArr(3)), CStr(varArr(4)), CStr(varArr(5)), _
            CStr(varArr(6)), CStr(varArr(7)), CStr(varArr(8)), CStr(varArr(9)))
    
    If strPar���ô��� <> "" Then
    
        strPar���ô��� = Mid(strPar���ô���, 2)
        varArr = Array()
        varArr = GetParTable(strPar���ô���, strTableIn, strTableOut)
        
        '���ô�������������ʹ����ϸ�����÷�����
        strSql = "Select a.No As ��ʶ��, 0 As ҽ��, a.�շ�ϸĿid, f.���� As ҩƷͨ����, f.��� || f.���� As ���, Sum(a.���ʽ��) As ����," & vbNewLine & _
            "       Sum(a.����) * b.����ϵ�� || b.���ﵥλ As ����" & vbNewLine & _
            "From ������ü�¼ A, ҩƷ��� B, ҩƷ���� C, �շ���ĿĿ¼ F" & vbNewLine & _
            "Where a.��¼״̬ <> 0 And a.�շ�ϸĿid = b.ҩƷid And b.ҩ��id = c.ҩ��id And b.ҩƷid = f.Id And" & vbNewLine & _
            "      a.No In (" & strTableOut & ") And a.�շ���� = '5' And Nvl(c.������, 0) <> 0 And a.ҽ����� Is Null" & vbNewLine & _
            "Group By a.No, a.�շ�ϸĿid, f.����, f.���, f.����, b.����ϵ��, b.���ﵥλ"

        Set rs���ÿ�ҩ��ϸ = zlDatabase.OpenSQLRecord(strSql, Me.Caption, CStr(varArr(0)), CStr(varArr(1)), CStr(varArr(2)), CStr(varArr(3)), CStr(varArr(4)), CStr(varArr(5)), _
            CStr(varArr(6)), CStr(varArr(7)), CStr(varArr(8)), CStr(varArr(9)))
    End If
    
    If strParҽ������ <> "" Then
        strParҽ������ = Mid(strParҽ������, 2)
        varArr = Array()
        varArr = GetParTable(strParҽ������, strTableIn, strTableOut)
        '����ҩ����ҩ����ϸ
        strSql = "Select a.No As ��ʶ��, 1 As ҽ��, f.���� As ҩƷͨ����, f.��� || f.���� As ���, Sum(a.���ʽ��) As ����, Sum(a.����) * b.����ϵ�� || b.���ﵥλ As ����," & vbNewLine & _
            "       e.ִ��Ƶ��,e.��������,i.���㵥λ,g.ҽ������ As ��ҩ;��" & vbNewLine & _
            "From ������ü�¼ A, ҩƷ��� B, ҩƷ���� C, �շ���ĿĿ¼ F,����ҽ����¼ E, ������ĿĿ¼ I, ����ҽ����¼ G" & vbNewLine & _
            "Where a.��¼״̬ <> 0 And a.�շ�ϸĿid = b.ҩƷid And b.ҩ��id = c.ҩ��id And b.ҩƷid = f.Id And e.���id = g.Id And" & vbNewLine & _
            "      e.������Ŀid = i.Id And a.No In (" & strTableOut & ")  And" & vbNewLine & _
            "      e.Id = a.ҽ����� And a.�շ���� = '5' And Nvl(c.������, 0) <> 0" & vbNewLine & _
            "Group By a.No, f.����, f.���, f.����, b.����ϵ��, b.���ﵥλ, e.ִ��Ƶ��, e.��������, i.���㵥λ, g.ҽ������"

        Set rsҽ����ҩ��ϸ = zlDatabase.OpenSQLRecord(strSql, Me.Caption, CStr(varArr(0)), CStr(varArr(1)), CStr(varArr(2)), CStr(varArr(3)), CStr(varArr(4)), CStr(varArr(5)), _
            CStr(varArr(6)), CStr(varArr(7)), CStr(varArr(8)), CStr(varArr(9)))
        
        '��ȡ���
        strSql = "select a.no as ��ʶ��,d.�������" & vbNewLine & _
            "from ������ü�¼ a,����ҽ����¼ b,���˹Һż�¼ c,������ϼ�¼ d" & vbNewLine & _
            "where a.ҽ�����=b.id and b.�Һŵ�=c.no and c.����id=d.����id and c.id=d.��ҳid and a.��¼״̬ <> 0 and a.no in (" & strTableOut & ")"
        Set rs��� = zlDatabase.OpenSQLRecord(strSql, Me.Caption, CStr(varArr(0)), CStr(varArr(1)), CStr(varArr(2)), CStr(varArr(3)), CStr(varArr(4)), CStr(varArr(5)), _
            CStr(varArr(6)), CStr(varArr(7)), CStr(varArr(8)), CStr(varArr(9)))
    End If
    '--------------------------��������������ȡ���--------------------------------------------------------------
    With vsgInfo
        .Rows = vsgInfo.FixedRows
        For i = 1 To rs����.RecordCount
            .AddItem ""
            lngBaseRow = .Rows - 1
            
            .TextMatrix(lngBaseRow, COL_CF�Һŵ�) = rs����!��ʶ��
            
            .TextMatrix(lngBaseRow, COL_CF���) = i
            .TextMatrix(lngBaseRow, COL_CF�����) = IIf("" = rs����!����� & "", rs����!��ʶ�� & "", rs����!����� & "")
            .TextMatrix(lngBaseRow, COL_CF��������) = Format(rs����!�������� & "", "YYYY-MM-DD")
            .TextMatrix(lngBaseRow, COL_CF��������) = rs����!�������� & ""
            .TextMatrix(lngBaseRow, COL_CF����ҽ��) = rs����!����ҽ�� & ""
            rs����.Filter = "����id=" & Val(rs����!����ID & "")
            If Not rs����.EOF Then .TextMatrix(lngBaseRow, COL_CF����) = rs����!���� & ""
            .TextMatrix(lngBaseRow, COL_CF��������) = rs����!���� & ""
            .TextMatrix(lngBaseRow, COL_CFҩƷ���) = Format(rs����!ҩ�� & "", strDec)
            
            rs�������.Filter = "��ʶ��='" & rs����!��ʶ�� & "'"
            If Not rs�������.EOF Then dblTmp = Val(rs�������!������� & "")
            .TextMatrix(lngBaseRow, COL_CF�������) = Format(dblTmp, strDec): dblTmp = 0
            
            rsҩƷ����.Filter = "��ʶ��='" & rs����!��ʶ�� & "'"
            If Not rsҩƷ����.EOF Then
                .TextMatrix(lngBaseRow, COL_CFҩƷƷ����) = rsҩƷ����!ҩƷ���� & ""
                .TextMatrix(lngBaseRow, COL_CF��ҩƷ����) = rsҩƷ����!����ҩ���� & ""
                .TextMatrix(lngBaseRow, COL_CF��ҩƷ����) = rsҩƷ����!����ҩ���� & ""
                .TextMatrix(lngBaseRow, COL_CF��ҩ���) = Format(rsҩƷ����!����ҩ�� & "", strDec)
            End If
            
            If Val(rs����!ҽ�� & "") = 0 Then
                rs���ÿ�ҩ��ϸ.Filter = 0
                rs���ÿ�ҩ��ϸ.Filter = "��ʶ��='" & rs����!��ʶ�� & "'"
                If Not rs���ÿ�ҩ��ϸ.EOF Then
                    For j = 1 To rs���ÿ�ҩ��ϸ.RecordCount
                        If j = 1 Then
                            .TextMatrix(lngBaseRow, COL_CFͨ����) = rs���ÿ�ҩ��ϸ!ҩƷͨ���� & ""
                            .TextMatrix(lngBaseRow, COL_CF���) = rs���ÿ�ҩ��ϸ!��� & ""
                            .TextMatrix(lngBaseRow, COL_CF����) = rs���ÿ�ҩ��ϸ!���� & ""
                            .TextMatrix(lngBaseRow, COL_CF���) = Format(rs���ÿ�ҩ��ϸ!���� & "", strDec)
                        Else
                            .AddItem ""
                            lngTmp = .Rows - 1
                            For k = COL_CF��� To COL_CF�Һŵ�
                                .TextMatrix(lngTmp, k) = .TextMatrix(lngBaseRow, k)
                            Next
                            .TextMatrix(lngTmp, COL_CFͨ����) = rs���ÿ�ҩ��ϸ!ҩƷͨ���� & ""
                            .TextMatrix(lngTmp, COL_CF���) = rs���ÿ�ҩ��ϸ!��� & ""
                            .TextMatrix(lngTmp, COL_CF����) = rs���ÿ�ҩ��ϸ!���� & ""
                            .TextMatrix(lngTmp, COL_CF���) = Format(rs���ÿ�ҩ��ϸ!���� & "", strDec)
                        End If
                        rs���ÿ�ҩ��ϸ.MoveNext
                    Next
                End If
            Else
                blnע�� = False
                rsҽ����ҩ��ϸ.Filter = 0
                rsҽ����ҩ��ϸ.Filter = "��ʶ��='" & rs����!��ʶ�� & "'"
                If Not rsҽ����ҩ��ϸ.EOF Then
                    For j = 1 To rsҽ����ҩ��ϸ.RecordCount
                        
                        '��ȡ�÷�����������strTmp ��
                        strTmp = rsҽ����ҩ��ϸ!�������� & ""
                        If Mid(strTmp, 1, 1) = "." Then strTmp = "0" & strTmp
                        strTmp = rsҽ����ҩ��ϸ!ִ��Ƶ�� & "," & strTmp & rsҽ����ҩ��ϸ!���㵥λ
                        
                        If j = 1 Then
                            .TextMatrix(lngBaseRow, COL_CFͨ����) = rsҽ����ҩ��ϸ!ҩƷͨ���� & ""
                            .TextMatrix(lngBaseRow, COL_CF���) = rsҽ����ҩ��ϸ!��� & ""
                            .TextMatrix(lngBaseRow, COL_CF����) = rsҽ����ҩ��ϸ!���� & ""
                            .TextMatrix(lngBaseRow, COL_CF���) = Format(rsҽ����ҩ��ϸ!���� & "", strDec)
                            .TextMatrix(lngBaseRow, COL_CF�÷�����) = strTmp
                            .TextMatrix(lngBaseRow, COL_CF��ҩ;��) = rsҽ����ҩ��ϸ!��ҩ;�� & ""
                        Else
                            .AddItem ""
                            lngTmp = .Rows - 1
                            For k = COL_CF��� To COL_CF�Һŵ�
                                .TextMatrix(lngTmp, k) = .TextMatrix(lngBaseRow, k)
                            Next
                            .TextMatrix(lngTmp, COL_CFͨ����) = rsҽ����ҩ��ϸ!ҩƷͨ���� & ""
                            .TextMatrix(lngTmp, COL_CF���) = rsҽ����ҩ��ϸ!��� & ""
                            .TextMatrix(lngTmp, COL_CF����) = rsҽ����ҩ��ϸ!���� & ""
                            .TextMatrix(lngTmp, COL_CF���) = Format(rsҽ����ҩ��ϸ!���� & "", strDec)
                            .TextMatrix(lngTmp, COL_CF�÷�����) = strTmp
                            .TextMatrix(lngTmp, COL_CF��ҩ;��) = rsҽ����ҩ��ϸ!��ҩ;�� & ""
                        End If
                        If InStr(rsҽ����ҩ��ϸ!��ҩ;�� & "", "ע��") > 0 Then blnע�� = True
                        rsҽ����ҩ��ϸ.MoveNext
                    Next
                End If
                .TextMatrix(lngBaseRow, COL_CFע���) = IIf(blnע��, "��", "��")
                strTmp = "": rs���.Filter = 0
                rs���.Filter = "��ʶ��='" & rs����!��ʶ�� & "'"
                If Not rs���.EOF Then
                    For j = 1 To rs���.RecordCount
                        If InStr("," & strTmp & ",", "," & rs���!������� & ",") = 0 Then
                            strTmp = strTmp & "," & rs���!�������
                        End If
                        rs���.MoveNext
                    Next
                End If
                .TextMatrix(lngBaseRow, COL_CF���) = Mid(strTmp, 2)
            End If
            rs����.MoveNext
        Next
    End With
    
    If blnRP Then
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\���ڷ�Χ", "���ﴦ��������ҩ�����", dtpRPS(e_R2_dtpRPS_��ʼʱ��_2).Value & "," & dtpRPE(e_R2_dtpRPE_����ʱ��_2).Value
    Else
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\���ڷ�Χ", "�ż��ﴦ��������ҩͳ��", dtpCountS(e_C2_dtpCountS_��ʼʱ��_2).Value & "," & dtpCountE(e_C2_dtpCountE_����ʱ��_2).Value
    End If
    
    Screen.MousePointer = 0
    Call zlCommFun.StopFlash
    Exit Sub
errH:
    Screen.MousePointer = 0
    Call zlCommFun.StopFlash
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub optType_Click(Index As Integer)
    Select Case Index
    Case e_C3_optType_�п�����_������_15, e_C3_optType_�п�����_����_18
        chkType(e_C3_chkType_�п�����_����_2).Enabled = Not optType(e_C3_optType_�п�����_������_15).Value
        chkType(e_C3_chkType_�п�����_����_3).Enabled = Not optType(e_C3_optType_�п�����_������_15).Value
        chkType(e_C3_chkType_�п�����_����_4).Enabled = Not optType(e_C3_optType_�п�����_������_15).Value
        chkType(e_C3_chkType_�п�����_����_8).Enabled = Not optType(e_C3_optType_�п�����_������_15).Value
    Case e_C0_optType_���ܷ�ʽ_����_9, e_C0_optType_���ܷ�ʽ_ҽ��_8, e_C0_optType_���ܷ�ʽ_ҩƷ_7
        If Index = e_C0_optType_���ܷ�ʽ_ҩƷ_7 Then '7-��ҩƷ����
            optType(e_C0_optType_����ʽ_����_12).Enabled = True
            optType(e_C0_optType_����ʽ_���_11).Enabled = True
        Else            '��9-���Һ�8-ҽ������
            optType(e_C0_optType_����ʽ_����_12).Enabled = False
            optType(e_C0_optType_����ʽ_���_11).Value = True
            optType(e_C0_optType_����ʽ_���_11).Enabled = False
        End If
    End Select
End Sub

Private Sub picDept_Resize()
    On Error Resume Next
    lvwItems.ColumnHeaders(1).Width = (picDept.Width - 1300)
End Sub

Private Sub picOther_Resize()
    On Error Resume Next
    
    tbcOther.Move 0, 0, picOther.Width, picOther.Height
End Sub

Private Sub picReport_Resize()
    On Error Resume Next
    
    tbcReport.Move 0, 0, picReport.Width, picReport.Height
End Sub

Private Sub picReportSub_Resize(Index As Integer)
    Dim lngT As Long '���붥���ľ���
    Dim lngL As Long '������˵ľ���
    Dim lngW As Long
    Dim lngH As Long
    Dim lngTmp As Long
    Dim lngW˵������ As Long '�ײ�˵�����ֵĿ�ȣ���һ����180�����㣬���������˵�����ֵ�����Ҫͬ������
    Dim int���� As Integer

    lngT = 50: lngL = 60
    lngW = picReportSub(Index).Width - lngL
    lngH = picReportSub(Index).Height
    lngW˵������ = 180

    On Error Resume Next
     
    Select Case Index
        Case PanelItem_����ҩƷ���Ľ������
            
            int���� = 1
            
            picFilter(e_R0_picFilter_��������_6).Left = lngL
            picFilter(e_R0_picFilter_��������_6).Top = lngT
            picFilter(e_R0_picFilter_��������_6).Height = cmdOK(e_R0_cmdOK_ͳ��_0).Top + cmdOK(e_R0_cmdOK_ͳ��_0).Height
            picFilter(e_R0_picFilter_��������_6).Width = lngW
            
            lblN(e_R0_lblN_����_74).Left = (lngW - lblN(e_R0_lblN_����_74).Width) / 2
            lblN(e_R0_lblN_����_74).Top = picFilter(e_R0_picFilter_��������_6).Height + picFilter(e_R0_picFilter_��������_6).Top + 100
            
            vsBill.Left = lngL
            vsBill.Top = lblN(e_R0_lblN_����_74).Top + lblN(e_R0_lblN_����_74).Height + 50
            vsBill.Width = lngW
            vsBill.Height = lngH - vsBill.Top - lngW˵������ * int���� - 120
            
            lblInfo.Left = lngL
            lblInfo.Top = lngH - lngW˵������ * int���� - 60
            
        Case PanelItem_���˿�����ҩ����������鼰���۱�
            
            int���� = 1
            
            picFilter(e_R1_picFilter_��������_7).Left = lngL
            picFilter(e_R1_picFilter_��������_7).Top = lngT
            picFilter(e_R1_picFilter_��������_7).Height = txtCYJL.Top + txtCYJL.Height
            picFilter(e_R1_picFilter_��������_7).Width = lngW
            
            lngTmp = txtCYJL.Top + txtCYJL.Height + 30 + lngT
            
            lblN(e_R1_lblN_����_77).Left = (lngW - lblN(e_R1_lblN_����_77).Width) / 2
            lblN(e_R1_lblN_����_77).Top = picFilter(e_R1_picFilter_��������_7).Height + picFilter(e_R1_picFilter_��������_7).Top + 100
            
            rptPati.Left = lngL
            rptPati.Top = lblN(e_R1_lblN_����_77).Height + lblN(e_R1_lblN_����_77).Top + 50
            rptPati.Width = lngW
            rptPati.Height = lngH - rptPati.Top - lngW˵������ * int���� - 120
            
            lblN(e_R1_lblN_�׶�˵��_72).Left = lngL
            lblN(e_R1_lblN_�׶�˵��_72).Top = lngH - lblN(e_R1_lblN_�׶�˵��_72).Height - 60
            
        Case PanelItem_���ﴦ��������ҩ�����
        
            int���� = 3
            
            picFilter(e_R2_picFilter_��������_8).Left = lngL
            picFilter(e_R2_picFilter_��������_8).Top = lngT
            picFilter(e_R2_picFilter_��������_8).Height = txtDept(e_R2_txtDept_��������_2).Top + txtDept(e_R2_txtDept_��������_2).Height
            picFilter(e_R2_picFilter_��������_8).Width = lngW
            
            lblN(e_R2_lblN_�����_����_25).Top = picFilter(e_R2_picFilter_��������_8).Height + picFilter(e_R2_picFilter_��������_8).Top + 100
            lblN(e_R2_lblN_�����_����_25).Left = (lngW - lblN(e_R2_lblN_�����_����_25).Width) / 2
            
            lblN(e_R2_lblN_�����_����_26).Top = lblN(e_R2_lblN_�����_����_25).Top + lblN(e_R2_lblN_�����_����_25).Height + 20
            lblN(e_R2_lblN_�����_����_26).Left = (lngW - lblN(e_R2_lblN_�����_����_26).Width) / 2
            
            lblN(e_R2_lblN_�����_��������_27).Top = lblN(e_R2_lblN_�����_����_26).Top
            lblN(e_R2_lblN_�����_��������_27).Left = lngW - lblN(e_R2_lblN_�����_��������_27).Width - 100
            
            vsMZYY.Left = lngL: vsMZYY.Width = lngW
            vsMZYY.Top = lblN(e_R2_lblN_�����_��������_27).Top + lblN(e_R2_lblN_�����_��������_27).Height + 50
            
            vsCF.Left = lngL: vsCF.Width = lngW
            vsCF.Height = 1800
            vsCF.Top = lngH - vsCF.Height - lngW˵������ * int���� - 120
            
            lblN(e_R2_lblN_������_����_29).Top = vsCF.Top - 50 - lblN(e_R2_lblN_������_����_29).Height
            lblN(e_R2_lblN_������_����_29).Left = (lngW - lblN(e_R2_lblN_������_����_29).Width) / 2
            
            lblN(e_R2_lblN_������_��������_30).Left = lngW - lblN(e_R2_lblN_������_��������_30).Width - 100
            lblN(e_R2_lblN_������_��������_30).Top = lblN(e_R2_lblN_������_����_29).Top
            
            lblN(e_R2_lblN_������_����_70).Top = lblN(e_R2_lblN_������_����_29).Top - 20 - lblN(e_R2_lblN_������_����_70).Height
            lblN(e_R2_lblN_������_����_70).Left = (lngW - lblN(e_R2_lblN_������_����_70).Width) / 2
            
            lblCFSM.Left = lngL
            lblCFSM.Top = lngH - lngW˵������ * int���� - 60
            
            vsMZYY.Height = lblN(e_R2_lblN_������_����_70).Top - vsMZYY.Top - 50
            
        Case PanelItem_סԺ���˿�����ҩ�����
            
            int���� = 1
            
            picFilter(e_R3_picFilter_��������_9).Left = lngL
            picFilter(e_R3_picFilter_��������_9).Top = lngT
            picFilter(e_R3_picFilter_��������_9).Height = txtDept(e_R3_txtDept_��������_3).Top + txtDept(e_R3_txtDept_��������_3).Height
            picFilter(e_R3_picFilter_��������_9).Width = lngW
        
            lblN(e_R3_lblN_�����_����_43).Left = (lngW - lblN(e_R3_lblN_�����_����_43).Width) / 2
            lblN(e_R3_lblN_�����_����_43).Top = picFilter(e_R3_picFilter_��������_9).Height + picFilter(e_R3_picFilter_��������_9).Top + 100
            
            lblN(e_R3_lblN_�����_��������_45).Left = lngW - lblN(e_R3_lblN_�����_��������_45).Width - 300
            lblN(e_R3_lblN_�����_��������_45).Top = lblN(e_R3_lblN_�����_����_43).Top + lblN(e_R3_lblN_�����_����_43).Height + 20
                
            vsZYYY.Left = lngL
            vsZYYY.Width = lngW
            vsZYYY.Top = lblN(e_R3_lblN_�����_��������_45).Top + lblN(e_R3_lblN_�����_��������_45).Height + 60
            vsZYYY.Height = lngH - vsZYYY.Top - lngW˵������ * int���� - 120
            
            lblN(e_R3_lblN_�׶�˵��_44).Left = lngL
            lblN(e_R3_lblN_�׶�˵��_44).Top = lngH - lngW˵������ * int���� - 60
            
    End Select
End Sub

Private Sub picOtherSub_Resize(Index As Integer)
    Dim lngL As Long
    Dim lngT As Long
    Dim lngW As Long
    Dim lngH As Long
    Dim lngW˵������ As Long '�ײ�˵�����ֵĿ�ȣ���һ����180�����㣬���������˵�����ֵ�����Ҫͬ������
    Dim int���� As Integer
    
    On Error Resume Next
    
    lngL = 60: lngT = 50
    lngW = picOtherSub(Index).Width
    lngH = picOtherSub(Index).Height
    lngW˵������ = 180
    
    picFilter(Index).Left = lngL
    picFilter(Index).Top = lngT
    picFilter(Index).Width = lngW - lngL
 
    Select Case Index
    
    Case PanelItem_����ҩ��ʹ���������ͳ��
        
        int���� = 1
        
        picFilter(e_C0_picFilter_��������_0).Height = cmdOK(e_C0_cmdOK_ͳ��_4).Top + cmdOK(e_C0_cmdOK_ͳ��_4).Height
        
        lblN(e_C0_lblN_����_75).Left = (lngW - lngL - lblN(e_C0_lblN_����_75).Width) / 2
        lblN(e_C0_lblN_����_75).Top = picFilter(e_C0_picFilter_��������_0).Top + picFilter(e_C0_picFilter_��������_0).Height + 100
        
        vsUseRan.Left = lngL
        vsUseRan.Width = lngW - lngL
        vsUseRan.Top = lblN(e_C0_lblN_����_75).Top + lblN(e_C0_lblN_����_75).Height + 50
        vsUseRan.Height = lngH - vsUseRan.Top - lngW˵������ * int���� - 120
        
        lblUse.Left = lngL
        
        lblUse.Top = lngH - lngW˵������ * int���� - 60
        
    Case PanelItem_�����п�Χ����Ԥ����ҩͳ��
        
        int���� = 1
        
        picFilter(e_C1_picFilter_��������_1).Height = cmdOK(e_C1_cmdOK_ͳ��_5).Top + cmdOK(e_C1_cmdOK_ͳ��_5).Height
        
        lblN(e_C1_lblN_����_76).Left = (lngW - lngL - lblN(e_C1_lblN_����_76).Width) / 2
        lblN(e_C1_lblN_����_76).Top = picFilter(e_C1_picFilter_��������_1).Top + picFilter(e_C1_picFilter_��������_1).Height + 100
        vsCut.Left = lngL
        vsCut.Top = lblN(e_C1_lblN_����_76).Height + lblN(e_C1_lblN_����_76).Top + 50
        vsCut.Width = lngW - lngL
        vsCut.Height = lngH - vsCut.Top - lngW˵������ * int���� - 120
        
        lblCut.Left = lngL
        lblCut.Top = lngH - lngW˵������ * int���� - 60
    
    Case PanelItem_�ż��ﴦ��������ҩͳ��
    
        int���� = 4
        
        picFilter(e_C2_picFilter_��������_2).Height = cmdOK(e_C2_cmdOK_ͳ��_6).Top + cmdOK(e_C2_cmdOK_ͳ��_6).Height
        
        lblN(e_C2_lblN_ͳ�Ʊ�_����_4).Left = (lngW - lngL - lblN(e_C2_lblN_ͳ�Ʊ�_����_4).Width) / 2
        lblN(e_C2_lblN_ͳ�Ʊ�_����_4).Top = picFilter(e_C2_picFilter_��������_2).Top + picFilter(e_C2_picFilter_��������_2).Height + 100
        
        lblN(e_C2_lblN_ͳ�Ʊ�_����_46).Left = (lngW - lngL - lblN(e_C2_lblN_ͳ�Ʊ�_����_46).Width) / 2
        lblN(e_C2_lblN_ͳ�Ʊ�_����_46).Top = lblN(e_C2_lblN_ͳ�Ʊ�_����_4).Top + lblN(e_C2_lblN_ͳ�Ʊ�_����_4).Height + 20
        
        lblN(e_C2_lblN_ͳ�Ʊ�_��������_48).Left = lngW - lblN(e_C2_lblN_ͳ�Ʊ�_��������_48).Width - 100
        lblN(e_C2_lblN_ͳ�Ʊ�_��������_48).Top = lblN(e_C2_lblN_ͳ�Ʊ�_����_46).Top
        
        vsCountDruUse.Left = lngL
        vsCountDruUse.Width = lngW - lngL
        vsCountDruUse.Top = lblN(e_C2_lblN_ͳ�Ʊ�_��������_48).Top + lblN(e_C2_lblN_ͳ�Ʊ�_��������_48).Height + 50
        
        vsCountCF.Left = lngL
        vsCountCF.Width = lngW - lngL
        vsCountCF.Height = 1800
        vsCountCF.Top = lngH - vsCountCF.Height - lngW˵������ * int���� - 120
        
        lblN(e_C2_lblN_�׶�˵��_50).Left = lngL
        lblN(e_C2_lblN_�׶�˵��_50).Top = lngH - lngW˵������ * int���� - 60

        lblN(e_C2_lblN_������_����_47).Left = (lngW - lngL - lblN(e_C2_lblN_������_����_47).Width) / 2
        lblN(e_C2_lblN_������_����_47).Top = vsCountCF.Top - lblN(e_C2_lblN_������_����_47).Height - 50
                
        lblN(e_C2_lblN_������_��������_49).Left = lngW - lblN(e_C2_lblN_������_��������_49).Width - 100
        lblN(e_C2_lblN_������_��������_49).Top = lblN(e_C2_lblN_������_����_47).Top
        
        
        lblN(e_C2_lblN_������_����_39).Left = (lngW - lngL - lblN(e_C2_lblN_������_����_39).Width) / 2
        lblN(e_C2_lblN_������_����_39).Top = lblN(e_C2_lblN_������_��������_49).Top - lblN(e_C2_lblN_������_����_39).Height - 20
        
        vsCountDruUse.Height = lblN(e_C2_lblN_������_����_39).Top - vsCountDruUse.Top - 50
        
    Case PanelItem_סԺҽ��������ҩͳ��
        
        int���� = 2
        
        picFilter(e_C3_picFilter_��������_3).Height = cmdOK(e_C3_cmdOK_ͳ��_7).Top + cmdOK(e_C3_cmdOK_ͳ��_7).Height
        
        lblN(e_C3_lblN_ͳ�Ʊ�_����_5).Left = (lngW - lngL - lblN(e_C3_lblN_ͳ�Ʊ�_����_5).Width) / 2
        lblN(e_C3_lblN_ͳ�Ʊ�_����_5).Top = picFilter(e_C3_picFilter_��������_3).Top + picFilter(e_C3_picFilter_��������_3).Height + 100
        
        vsInDruUse.Left = lngL
        vsInDruUse.Top = lblN(e_C3_lblN_ͳ�Ʊ�_����_5).Top + lblN(e_C3_lblN_ͳ�Ʊ�_����_5).Height + 50
        vsInDruUse.Width = lngW - lngL
        
        vsInDruAna.Left = lngL
        vsInDruAna.Width = lngW - lngL
        vsInDruAna.Height = 2110
        vsInDruAna.Top = lngH - vsInDruAna.Height - lngW˵������ * int���� - 120
        
        lblN(e_C3_lblN_������_����_59).Left = (lngW - lngL - lblN(e_C3_lblN_������_����_59).Width) / 2
        lblN(e_C3_lblN_������_����_59).Top = vsInDruAna.Top - lblN(e_C3_lblN_������_����_59).Height - 50
        
        lblN(e_C3_lblN_�׶�˵��_58).Left = lngL
        lblN(e_C3_lblN_�׶�˵��_58).Top = lngH - lngW˵������ * int���� - 60
        
        
        vsInDruUse.Height = lblN(e_C3_lblN_������_����_59).Top - vsInDruUse.Top - 50
        
    Case PanelItem_���󿹾�ҩ��ʹ�ó�N��ͳ��
        
        int���� = 1
        
        picFilter(e_C4_picFilter_��������_4).Height = cmdOK(e_C4_cmdOK_ͳ��_8).Top + cmdOK(e_C4_cmdOK_ͳ��_8).Height
        
        lblN(e_C4_lblN_ͳ�Ʊ�_����_6).Left = (lngW - lngL - lblN(e_C4_lblN_ͳ�Ʊ�_����_6).Width) / 2
        lblN(e_C4_lblN_ͳ�Ʊ�_����_6).Top = picFilter(e_C4_picFilter_��������_4).Top + picFilter(e_C4_picFilter_��������_4).Height + 100
        
        vsOpeKssUse.Left = lngL
        vsOpeKssUse.Width = lngW - lngL
        vsOpeKssUse.Top = lblN(e_C4_lblN_ͳ�Ʊ�_����_6).Top + lblN(e_C4_lblN_ͳ�Ʊ�_����_6).Height + 50
        vsOpeKssUse.Height = lngH - vsOpeKssUse.Top - lngW˵������ * int���� - 120
        
        lblN(e_C4_lblN_�׶�˵��_73).Left = lngL
        lblN(e_C4_lblN_�׶�˵��_73).Top = lngH - lngW˵������ * int���� - 60
        
    Case PanelItem_ҽ������ĳ����������ҩ�ɱ�ͳ��
        
        int���� = 3
        
        picFilter(e_C5_picFilter_��������_5).Height = cmdOK(e_C5_cmdOK_ͳ��_9).Top + cmdOK(e_C5_cmdOK_ͳ��_9).Height + 70
        
        lblN(e_C5_lblN_������_����_7).Left = (lngW - lngL - lblN(7).Width) / 2
        lblN(e_C5_lblN_������_����_7).Top = picFilter(e_C5_picFilter_��������_5).Top + picFilter(e_C5_picFilter_��������_5).Height + 30
        
        vsIllDruUse.Left = lngL
        vsIllDruUse.Width = lngW - lngL
        vsIllDruUse.Top = lblN(e_C5_lblN_������_����_7).Top + lblN(e_C5_lblN_������_����_7).Height + 50
        vsIllDruUse.Height = lngH - vsIllDruUse.Top - lngW˵������ * int���� - 120
        
        lblN(e_C5_lblN_�׶�˵��_67).Left = lngL
        lblN(e_C5_lblN_�׶�˵��_67).Top = lngH - lngW˵������ * int���� - 60
    End Select
End Sub

Private Sub txtCount_Validate(Index As Integer, Cancel As Boolean)
    If Val(txtCount(Index).Text) = 0 Then txtCount(Index).Text = 0
End Sub

Private Sub txtILL_GotFocus()
    Call zlControl.TxtSelAll(txtILL)
End Sub

Private Sub txtILL_KeyPress(KeyAscii As Integer)
    Dim blnCancel As Boolean

    If KeyAscii = 13 Then
        Call txtILL_Validate(blnCancel)
        If Not blnCancel Then Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub vsInDruUse_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
'���ܣ���������
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim i As Long
    Dim vRect As RECT
    
    If Row < 2 Then Exit Sub
    
    With vsInDruUse
        If Col >= COL_DRU��� And Col <= COL_DRU������ҩ Then
            lngBegin = Row: lngEnd = Row
            
            For i = Row - 1 To .FixedRows Step -1
                If Val(.TextMatrix(Row, COL_DRU����id)) = Val(.TextMatrix(i, COL_DRU����id)) And Val(.TextMatrix(Row, COL_DRU��ҳid)) = Val(.TextMatrix(i, COL_DRU��ҳid)) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            
            For i = Row + 1 To .Rows - 1
                If Val(.TextMatrix(Row, COL_DRU����id)) = Val(.TextMatrix(i, COL_DRU����id)) And Val(.TextMatrix(Row, COL_DRU��ҳid)) = Val(.TextMatrix(i, COL_DRU��ҳid)) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
            
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
            
        End If
    End With
End Sub

Private Sub vsMZYY_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    vsMZYY.RowHeightMax = 2000
    vsMZYY.RowHeightMin = 250
    vsMZYY.AutoSize 0, vsMZYY.Cols - 1
End Sub

Private Sub vsCountCF_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    vsCountCF.RowHeightMax = 2000
    vsCountCF.RowHeightMin = 250
    vsCountCF.AutoSize 0, vsCountCF.Cols - 1
End Sub

Private Sub vsInDruUse_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    vsInDruUse.RowHeightMax = 2000
    vsInDruUse.RowHeightMin = 250
    vsInDruUse.AutoSize 0, vsInDruUse.Cols - 1
End Sub

Private Sub vsCountDruUse_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
'���ܣ��������ֱ߿���
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim i As Long
    Dim vRect As RECT
    
    With vsCountDruUse
        If Row < 3 Then Exit Sub
        If Col >= COL_CF��� And Col <= COL_CF��ҩƷ���� Or Col >= COL_CF������� And Col <= COL_CF��ҩ��� Then
            lngBegin = Row: lngEnd = Row
            
            For i = Row - 1 To .FixedRows Step -1
                If .TextMatrix(i, COL_CF�Һŵ�) <> .TextMatrix(Row, COL_CF�Һŵ�) Then
                    Exit For
                Else
                    lngBegin = i
                End If
            Next
            
            For i = Row + 1 To .Rows - 1
                If .TextMatrix(i, COL_CF�Һŵ�) <> .TextMatrix(Row, COL_CF�Һŵ�) Then
                    Exit For
                Else
                    lngEnd = i
                End If
            Next
            
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
        End If
    End With
End Sub

Private Sub vsMZYY_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
'���ܣ��������ֱ߿���
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim i As Long
    Dim vRect As RECT
    
    With vsMZYY
        If Row < 3 Then Exit Sub
        If Col >= COL_CF��� And Col <= COL_CF��ҩƷ���� Or Col >= COL_CF������� And Col <= COL_CF��ҩ��� Then
            lngBegin = Row: lngEnd = Row
            
            For i = Row - 1 To .FixedRows Step -1
                If .TextMatrix(i, COL_CF�Һŵ�) <> .TextMatrix(Row, COL_CF�Һŵ�) Then
                    Exit For
                Else
                    lngBegin = i
                End If
            Next
            
            For i = Row + 1 To .Rows - 1
                If .TextMatrix(i, COL_CF�Һŵ�) <> .TextMatrix(Row, COL_CF�Һŵ�) Then
                    Exit For
                Else
                    lngEnd = i
                End If
            Next
            
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
        End If
    End With
End Sub

Private Sub LoadInKssAdvice()
'���ܣ����� סԺ���˿�����ҩ����� ��������
    Dim strSql As String, strPar As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    Dim strValue As String
    Dim dblTmp As Double
    Dim i As Long
    Dim j As Long
    
    On Error GoTo errH
    Screen.MousePointer = 11
    Call zlCommFun.ShowFlash("���ڶ�ȡ���ݣ����Ժ�...")
    
    '���ڷ�Χ��������
    strPar = "To_Date('" & Format(dtpRPS(e_R3_dtpRPS_��ʼʱ��_3).Value, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')" & _
    " and To_Date('" & Format(dtpRPE(e_R3_dtpRPE_����ʱ��_3).Value, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
    
    '���ر������
    strSql = "select /*+ rule*/ e.���� as ���,c.���� as ҩƷͨ����,d.ҩƷ���� as ����,c.���||c.���� as ���,b.סԺ��λ as ��λ,a.����,a.����" & vbNewLine & _
        "from (select x.�շ�ϸĿid,sum(x.���ʽ��) as ����,sum(x.����) as ���� from סԺ���ü�¼ X,������ҳ Y where x.�շ����='5'and x.��¼״̬ <> 0" & vbNewLine & _
        "and y.����id=x.����id and y.��ҳid=x.��ҳid" & _
        " and y.��Ժ���� between " & strPar & vbNewLine & _
        IIf(txtDept(e_R3_txtDept_��������_3).Tag = "", "", " and y.��Ժ����id in (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))") & vbNewLine & _
        "group by x.�շ�ϸĿid) a,ҩƷ��� b,�շ���ĿĿ¼ c,ҩƷ���� d,���Ʒ���Ŀ¼ e,������ĿĿ¼ f" & vbNewLine & _
        "where a.�շ�ϸĿid=b.ҩƷid and b.ҩƷid=c.id and d.ҩ��id=b.ҩ��ID and d.ҩ��id=f.id and f.����id=e.id and nvl(d.������,0)<>0"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, txtDept(e_R3_txtDept_��������_3).Tag)

    If rsTmp.RecordCount > 0 Then
        With vsZYYY
            .Rows = vsZYYY.FixedRows
            For i = 1 To rsTmp.RecordCount
                '�������
                strTmp = ""
                For j = 0 To rsTmp.Fields.Count - 1
                    If j = rsTmp.Fields.Count - 1 Then
                        strValue = Replace(Nvl(rsTmp.Fields(j).Value), vbTab, "")
                        strTmp = strTmp & vbTab & IIf(Mid(strValue, 1, 1) = ".", "0" & strValue, strValue)
                    Else
                        strTmp = strTmp & vbTab & Replace(Nvl(rsTmp.Fields(j).Value), vbTab, "")
                    End If
                Next
                .AddItem Mid(strTmp, 2)
                rsTmp.MoveNext
            Next
            .SubtotalPosition = flexSTBelow
            .Subtotal flexSTSum, -1, COL_YZ�ܷ���, "#######" & gstrDec, , vbBlack, False, "�ϼ�"
            
            .MergeCellsFixed = flexMergeFree
            .MergeCol(0) = True
            
            '��ʽ�������ͽ�����������λС��
            strTmp = "0.00"
            
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, COL_YZ����) = Format(.TextMatrix(i, COL_YZ����), gstrDec)
                .TextMatrix(i, COL_YZ�ܷ���) = Format(.TextMatrix(i, COL_YZ�ܷ���), gstrDec)
            Next
        End With
        
        '���� ����������  ҽԺ���λ�����������ҽԺ���ȳ�Ժ������������ͬ��ƽ��סԺ��������ҽԺͳ�Ʋ����ṩ��
        strSql = "Select sum(סԺ����) as ������ From ������ҳ Where ��Ժ���� between " & strPar & _
            IIf(txtDept(e_R3_txtDept_��������_3).Tag = "", "", " and ��Ժ����id in (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))")
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, txtDept(e_R3_txtDept_��������_3).Tag)
        If Not rsTmp.EOF > 0 Then dblTmp = Val(rsTmp!������ & "")
        lblN(e_R3_lblN_�����_��������_45).Caption = "���λ�����������" & Round(dblTmp) & "��"
    Else
        vsZYYY.Rows = vsZYYY.FixedRows
        vsZYYY.Rows = vsZYYY.Rows + 1
        Screen.MousePointer = 0
        Call zlCommFun.StopFlash
        MsgBox "��ǰ������δ�ҵ��κ����ݣ����������ó���ͳ�Ʋ�����", vbInformation, gstrSysName
        Exit Sub
    End If
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\���ڷ�Χ", "סԺ���˿�����ҩ�����", dtpRPS(e_R3_dtpRPS_��ʼʱ��_3).Value & "," & dtpRPE(e_R3_dtpRPE_����ʱ��_3).Value
    Screen.MousePointer = 0
    Call zlCommFun.StopFlash
    Exit Sub
errH:
    Screen.MousePointer = 0
    Call zlCommFun.StopFlash
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub zlRptPrint(ByVal bytMode As Byte, Optional ByVal intTabNum As Integer = 1)
'����:��¼���ӡ
'����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL,intTabNumҳ���еĵڼ������Ĭ��ֻ��һ�����������������ʱѭ�����
    Dim objPrint As New zlPrint1Grd
    Dim objAppRow As New zlTabAppRow
    Dim objVSF As VSFlexGrid
    Dim objTable As Object
    Dim objReport As ReportControl
    Dim blnIsRPT As Boolean   'True-��ReportControl������Ҫת����VSF����
    Dim varArr As Variant
    Dim i As Integer
    Dim strTmp As String
    Dim strButtom As String
    Dim strFace As String
    
    strFace = GetCurҳ��
    
    Select Case strFace
    
    Case "����ҩƷ���Ľ������"
    
        objPrint.Title.Text = "����ҩƷ���Ľ������"
        
        Set objAppRow = New zlTabAppRow
        objAppRow.Add "ͳ��ʱ�䣺" & dtpRPS(e_R0_dtpRPS_��ʼʱ��_0).Value & " �� " & dtpRPE(e_R0_dtpRPE_����ʱ��_0).Value
        objPrint.UnderAppRows.Add objAppRow
        
        Set objTable = vsBill
        
        strButtom = lblInfo.Caption
    Case "���˿�����ҩ����������鼰���۱�"
    
        objPrint.Title.Text = "���������б�"
        
    
        Set objAppRow = New zlTabAppRow
        objAppRow.Add "ͳ��ʱ�䣺" & dtpRPS(e_R1_dtpRPS_��ʼʱ��_1).Value & " �� " & dtpRPE(e_R1_dtpRPE_����ʱ��_1).Value
        objPrint.UnderAppRows.Add objAppRow
        
        Set objAppRow = New zlTabAppRow
        objAppRow.Add "�������ң�" & txtCYJL.Text
        objAppRow.Add "����" & Val(lblN(e_R1_lblN_������������_13).Tag) & "�ˡ�"
        objPrint.UnderAppRows.Add objAppRow
        
        Set objTable = rptPati
        blnIsRPT = True
        strButtom = vbCrLf & lblN(e_R1_lblN_�׶�˵��_72).Caption
        
    Case "���ﴦ��������ҩ�����"
        Set objAppRow = New zlTabAppRow
        objAppRow.Add "ͳ��ʱ�䣺" & dtpRPS(e_R2_dtpRPS_��ʼʱ��_2).Value & " �� " & dtpRPE(e_R2_dtpRPE_����ʱ��_2).Value
        objAppRow.Add lblN(e_R2_lblN_�����_��������_27).Caption
        objPrint.UnderAppRows.Add objAppRow
            
        If intTabNum = 1 Then
            objPrint.Title.Text = lblN(e_R2_lblN_�����_����_25).Caption
            Set objTable = vsMZYY
        ElseIf intTabNum = 2 Then
            objPrint.Title.Text = lblN(e_R2_lblN_������_����_70).Caption
            Set objTable = vsCF
        End If
        
        strButtom = lblCFSM.Caption
    Case "סԺ���˿�����ҩ�����"
    
        objPrint.Title.Text = lblN(e_R3_lblN_�����_����_43).Caption
        
        Set objAppRow = New zlTabAppRow
        objAppRow.Add "ͳ��ʱ�䣺" & dtpRPS(e_R3_dtpRPS_��ʼʱ��_3).Value & " �� " & dtpRPE(e_R3_dtpRPE_����ʱ��_3).Value
        objAppRow.Add lblN(e_R3_lblN_�����_��������_45).Caption
        objPrint.UnderAppRows.Add objAppRow
        
        Set objTable = vsZYYY
        
        strButtom = lblN(e_R3_lblN_�׶�˵��_44).Caption
        
    Case "����ҩ��ʹ���������ͳ��"
        objPrint.Title.Text = "����ҩ��ʹ���������ͳ��"
        
        Set objAppRow = New zlTabAppRow
        objAppRow.Add "ͳ��ʱ�䣺" & dtpCountS(e_C0_dtpCountS_��ʼʱ��_0).Value & " �� " & dtpCountE(e_C0_dtpCountE_����ʱ��_0).Value
        objPrint.UnderAppRows.Add objAppRow
        
        Set objTable = vsUseRan
        
        strButtom = lblUse.Caption
    Case "�����п�Χ����Ԥ����ҩͳ��"
        objPrint.Title.Text = "�����п�Χ����Ԥ����ҩͳ��"
        
        Set objAppRow = New zlTabAppRow
        objAppRow.Add "ͳ��ʱ�䣺" & dtpCountS(e_C1_dtpCountS_��ʼʱ��_1).Value & " �� " & dtpCountE(e_C1_dtpCountE_����ʱ��_1).Value
        objPrint.UnderAppRows.Add objAppRow
        
        Set objTable = vsCut
        
        strButtom = lblCut.Caption
        
    Case "�ż��ﴦ��������ҩͳ��"
        Set objAppRow = New zlTabAppRow
        objAppRow.Add "ͳ��ʱ�䣺" & dtpCountS(e_C2_dtpCountS_��ʼʱ��_2).Value & " �� " & dtpCountE(e_C2_dtpCountE_����ʱ��_2).Value
        objAppRow.Add lblN(e_C2_lblN_ͳ�Ʊ�_��������_48).Caption
        objPrint.UnderAppRows.Add objAppRow
            
        If intTabNum = 1 Then
            objPrint.Title.Text = lblN(e_C2_lblN_ͳ�Ʊ�_����_4).Caption
            Set objTable = vsCountDruUse
        ElseIf intTabNum = 2 Then
            objPrint.Title.Text = lblN(e_C2_lblN_������_����_39).Caption
            Set objTable = vsCountCF
        End If
        
        strButtom = lblN(e_C2_lblN_�׶�˵��_50).Caption
    
    Case "סԺҽ��������ҩͳ��"
        Set objAppRow = New zlTabAppRow
        objAppRow.Add "ͳ��ʱ�䣺" & dtpCountS(e_C3_dtpCountS_��ʼʱ��_3).Value & " �� " & dtpCountE(e_C3_dtpCountE_����ʱ��_3).Value
        objPrint.UnderAppRows.Add objAppRow
            
        If intTabNum = 1 Then
            objPrint.Title.Text = lblN(e_C3_lblN_ͳ�Ʊ�_����_5).Caption
            Set objTable = vsInDruUse
        ElseIf intTabNum = 2 Then
            objPrint.Title.Text = lblN(e_C3_lblN_������_����_59).Caption
            Set objTable = vsInDruAna
        End If
        
        strButtom = lblN(e_C3_lblN_�׶�˵��_58).Caption
        
    
    Case "���󿹾�ҩ��ʹ�ó�N��ͳ��"
        objPrint.Title.Text = "���󿹾�ҩ��ʹ�ó�N��ͳ��"
        Set objAppRow = New zlTabAppRow
        objAppRow.Add "ͳ��ʱ�䣺" & dtpCountS(e_C4_dtpCountS_��ʼʱ��_4).Value & " �� " & dtpCountE(e_C4_dtpCountE_����ʱ��_4).Value
        objPrint.UnderAppRows.Add objAppRow
        Set objTable = vsOpeKssUse
        strButtom = lblN(e_C4_lblN_�׶�˵��_73).Caption
    Case "ҽ������ĳ����������ҩ�ɱ�ͳ��"
        objPrint.Title.Text = "ҽ������ĳ����������ҩ�ɱ�ͳ��"
        Set objAppRow = New zlTabAppRow
        objAppRow.Add "ͳ��ʱ�䣺" & dtpCountS(e_C5_dtpCountS_��ʼʱ��_5).Value & " �� " & dtpCountE(e_C5_dtpCountE_����ʱ��_5).Value
        objAppRow.Add "��ϣ�" & txtILL.Text
        objPrint.UnderAppRows.Add objAppRow
        Set objTable = vsIllDruUse
        strButtom = lblN(e_C5_lblN_�׶�˵��_67).Caption
    End Select
    
    '�������ݱ��
    If blnIsRPT Then
        Set objReport = objTable
        If objReport.Records.Count = 0 Then Exit Sub
        If Not zlReportToVSFlexGrid(vsTmp, objReport) Then Exit Sub
        blnIsRPT = False
    Else
        Set objVSF = objTable
        If Not zlCopyVSFlexGrid(vsTmp, objVSF) Then Exit Sub
    End If
    
    '���ô�ӡ��������
    '---------------------------------------
    Set objPrint.Body = Me.vsTmp
    '����
    If strButtom <> "" Then
        varArr = Split(strButtom, vbCrLf)
        For i = 0 To UBound(varArr)
            Set objAppRow = New zlTabAppRow
            strTmp = varArr(i)
            Call objAppRow.Add(strTmp)
            Call objPrint.BelowAppRows.Add(objAppRow)
        Next
        strButtom = ""
    End If
    
    '��ӡ��ʱ�����Ϣ
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("��ӡ��:" & UserInfo.����)
    Call objAppRow.Add("��ӡʱ��:" & Format(Now, "yyyy-MM-dd HH:mm"))
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
    
    If (strFace = "���ﴦ��������ҩ�����" Or strFace = "�ż��ﴦ��������ҩͳ��" Or strFace = "סԺҽ��������ҩͳ��") And intTabNum = 1 Then
        Call zlRptPrint(bytMode, 2)
    End If
    
End Sub

Private Sub Set����˵��(ByVal strCaption As String)
'���ܣ����ý����·���˵����Ϣ
    Select Case strCaption
        Case "����ҩƷ���Ľ������"
            lblInfo.Caption = "˵������������Ϊ�㣬ϵͳ�в���¼�˷��ã������롢ҩƷ������Ͳ����Դ�ڻ��ܼ�¼����������Դ�ڷ�����ϸ��¼����������"
        Case "���˿�����ҩ����������鼰���۱�"
            lblN(e_R1_lblN_�׶�˵��_72).Caption = "˵��������Ժʱ��ͳ�Ժ���ҹ��˳�Ժ���ˣ�����δʹ�ÿ�ҩ�ĳ�Ժ���ˡ�"
        Case "���ﴦ��������ҩ�����"
            lblCFSM.Caption = "˵��������ָһ�Ű���ҩƷ���õ��շѵ��ݡ�ͨ��ҽ�������ķ��õ��ݲ�����ҩ��ϸ��ϵȣ�ֱ���շѲ����ĵ���ֻ��ҩƷ��Ϣ��" & vbCrLf & _
                "       �����÷���ʱ��Ϳ����ҿƳ������������������ʾΪ���õĵ��ݺţ���������ָ���÷���ʱ�䣬����ҽ��ָ�����ˣ�����ָ�������ҡ�" & vbCrLf & _
                "       ��ҩƷƷ�ּ������ҩƷ������ͬһҩƷ��ͬ�����һ��ҩ��"
        Case "סԺ���˿�����ҩ�����"
            lblN(e_R3_lblN_�׶�˵��_44).Caption = "˵��������Ժʱ��ͳ�Ժ���ҹ���ʹ���˿���ҩ�ĳ�Ժ���ˣ����λ���������=��Ժ����������ͬ��ƽ��סԺ����������סԺ������"
        Case "����ҩ��ʹ���������ͳ��"
            lblUse.Caption = "˵����ͳ��ʱ��ָ���÷���ʱ�䣻��������ָ�������ң���סԺ����ͳ��ʱ������һ��סԺ��һ�Σ�ͳ���������ʱ�Դ���Ϊ��λ��"
        Case "�����п�Χ����Ԥ����ҩͳ��"
            lblCut.Caption = "˵����ͨ��ʱ��Ϳ��ҹ��˳�Ժ���ˣ�����ҩ��ҽ����ʽ�´����ͳ�ƣ�ҩƷ������Ʒ��ͳ�ƣ�������������ҽ��ִ�з�������õ���������ҳ�в��˿����ؼ�¼��ȡ�����û����Ĭ��Ϊһ�졣"
        Case "�ż��ﴦ��������ҩͳ��"
            lblN(e_C2_lblN_�׶�˵��_50).Caption = "˵��������ָһ�Ű���ҩƷ���õ��շѵ��ݡ�ͨ��ҽ�������ķ��ò�����ҩ��ϸ��ϵȣ�ֱ���շѲ����ĵ���ֻ��ҩƷ��Ϣ��" & vbCrLf & _
                "       �������֣��ӷ��õ��������Ƿ�������ҽ�������ĵ������һ���жϲ��˹Һ��Ƿ��" & vbCrLf & _
                "       �����÷���ʱ��Ϳ����ҿƳ�������������������ʾΪ���õĵ��ݺţ���������ָ�÷ѷ���ʱ�䣬����ҽ��ָ�����ˣ�����ָ�������ҡ�" & vbCrLf & _
                "       ��ҩƷƷ�ּ������ҩƷ������ͬһҩƷ��ͬ�����һ��ҩ��"
        Case "סԺҽ��������ҩͳ��"
            lblN(e_C3_lblN_�׶�˵��_58).Caption = "˵���������˳�Ժʱ��ͳ�Ժ���ҳ������ˣ��������ҳ�е���(��)ҽ��Ҫ��Ժ��ϣ�ҩƷ�����ǰ�Ʒ��ͳ�ƣ�ͬ��ҩ��ͬ�����һ��ҩ��" & vbCrLf & _
                "       ����ҩʹ����ϸָ���Բ��˵�ҽ����¼����������ҩ��������ҽ���´�ʱ���Ϸ����Ӧ��Ϊ�㡣"
        Case "���󿹾�ҩ��ʹ�ó�N��ͳ��"
            lblN(e_C4_lblN_�׶�˵��_73).Caption = "˵����ͳ�Ƴ�Ժ���ˣ�Ҫ����ҳ����д�������������ͳ��"
        Case "ҽ������ĳ����������ҩ�ɱ�ͳ��"
            lblN(e_C5_lblN_�׶�˵��_67).Caption = "˵����ͳ�ƶ���Ϊ����������ȫ����Ժ���ˣ����˳�Ժʱ����ָ����Χ������ҳ����д�ĳ�Ժ��Ҫ���Ϊ������ѡ���ָ����ϡ�" & vbCrLf & _
                "       ����ҽ��ָ����סԺҽʦ���û��סԺҽʦ����ʾΪ�գ�������(%)=��������/�����������˾����ƽ��=�ܽ��/����������" & vbCrLf & _
                "       �˾��ս��=�ܽ��/���Ʋ���סԺ����֮�ͣ�����ҩ��Ʒ��������ͬҩƷ���Ʋ�ͬ����ʱֻ��һ�֡�"
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub
 
Private Sub rptPati_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Dim lngIndex As Long
    If rptPati.SelectedRows.Count > 0 Then Call Edit���˵���
End Sub
 
Private Sub tbcOther_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Call Set����˵��(Item.Tag)
End Sub
 
Private Sub tbcReport_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Call Set����˵��(Item.Tag)
End Sub

Private Sub txtCount_KeyPress(Index As Integer, KeyAscii As Integer)
'���ܣ�����������ֻ����������
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtNum_KeyPress(Index As Integer, KeyAscii As Integer)
'���ܣ�����������ֻ����������
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtTopRan_KeyPress(KeyAscii As Integer)
'���ܣ�����������ֻ����������
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub LoadDept()
'����ѡ����------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim objItem As ListItem
    
    On Error GoTo errH
    
    strSql = "select distinct ID,����,����" & _
        " from ���ű� D,��������˵�� T" & _
        " where D.ID=T.����ID and ��������=[1] " & _
        " and (D.����ʱ�� is null or D.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))" & _
        " order by ����"
                
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, "�ٴ�")
    
    Me.lvwItems.ListItems.Clear
    
    Me.lvwItems.Checkboxes = True
   
    Do Until rsTmp.EOF
        Set objItem = Me.lvwItems.ListItems.Add(, "_" & rsTmp!ID, rsTmp!����)
        objItem.Icon = "Dept": objItem.SmallIcon = "Dept"
        objItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) = rsTmp!����
        objItem.Checked = False
        rsTmp.MoveNext
    Loop
    
    'û��ʱ�˳�
    If Me.lvwItems.ListItems.Count = 0 Then Exit Sub
    
    Me.lvwItems.ListItems(1).Selected = True
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub lvwItems_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mlngFind = Item.Index + 1
End Sub

Private Sub lvwItems_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then 'ȫѡ Ctrl+A
        Call SetSelect(lvwItems, True)
    End If
    
    If KeyCode = vbKeyR And Shift = vbCtrlMask Then     'ȫ�� Ctrl+R
        Call SetSelect(lvwItems, False)
    End If
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call cmdFind_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And picDept.Visible Then
        picDept.Visible = False
        txtFind.Text = ""
    End If
    If Me.ActiveControl.Name = "txtILL" Then Exit Sub
    If TypeName(Me.ActiveControl) <> "CommandButton" And KeyCode = vbKeyReturn Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
    
    If Not picDept.Visible And KeyCode = vbKeyD And Shift = vbCtrlMask Then
        
        Select Case GetCurҳ��
        
        Case "���˿�����ҩ����������鼰���۱�"
            Call cmdDept_Click(e_R1_cmdDept_����ѡ����_1)
  
        Case "���ﴦ��������ҩ�����"
            Call cmdDept_Click(e_R2_cmdDept_����ѡ����_2)
            
        Case "סԺ���˿�����ҩ�����"
        
            Call cmdDept_Click(e_R3_cmdDept_����ѡ����_3)
            
        Case "����ҩ��ʹ���������ͳ��"
            Call cmdDept_Click(e_C0_cmdDept_����ѡ����_0)
            
        Case "�����п�Χ����Ԥ����ҩͳ��"
            Call cmdDept_Click(e_C1_cmdDept_����ѡ����_4)
            
        Case "�ż��ﴦ��������ҩͳ��"
            Call cmdDept_Click(e_C2_cmdDept_����ѡ����_5)
        
        Case "סԺҽ��������ҩͳ��"
            Call cmdDept_Click(e_C3_cmdDept_����ѡ����_6)
            
        Case "���󿹾�ҩ��ʹ�ó�N��ͳ��"
            Call cmdDept_Click(e_C4_cmdDept_����ѡ����_7)
            
        End Select
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        If GetCurҳ�� = "���˿�����ҩ����������鼰���۱�" Then
            If picDept.Visible = False Then Call cmdCYSel_Click
        End If
    ElseIf KeyCode = vbKeyI And Shift = vbCtrlMask Then
        If GetCurҳ�� = "ҽ������ĳ����������ҩ�ɱ�ͳ��" Then
            Call cmdILL_Click
        End If
    End If
End Sub

Private Sub cmdFind_Click()
    Dim strFind As String
    Dim i As Long
    Dim blnIsFind As Boolean
    
    strFind = UCase(Trim(txtFind.Text))
    If strFind = "" Then Exit Sub
    For i = mlngFind To lvwItems.ListItems.Count
        If zlCommFun.SpellCode(Mid(lvwItems.ListItems(i).Text, InStr(lvwItems.ListItems(i).Text, "-") + 1)) Like UCase(IIf(mstrMatch <> "", "*", "") & strFind & "*") Or _
                UCase(lvwItems.ListItems(i).Text) Like UCase(IIf(mstrMatch <> "", "*", "") & strFind & "*") Then
            lvwItems.ListItems(i).Selected = True
            lvwItems.ListItems(i).EnsureVisible
            blnIsFind = True
            mlngFind = i + 1
            Exit For
        End If
    Next
    If blnIsFind = False Then
        If mlngFind = 1 Then
            MsgBox "û���ҵ������ҵĿ��ҡ�", vbInformation, Me.Caption
        Else
            MsgBox "�Ѿ������һ�������ˡ�", vbInformation, Me.Caption
            mlngFind = 1
        End If
    End If
End Sub

Private Sub cmdFindCancle_Click()
    Call lvwItems_KeyPress(vbKeyEscape)
End Sub

Private Sub cmdFindOk_Click()
    Call lvwItems_DblClick
End Sub

Private Sub lvwItems_LostFocus()
    Call picDept_LostFocus
End Sub

Private Sub lvwItems_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn, vbKeySpace
        If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
        If lvwItems.SelectedItem.Checked = False And KeyAscii = vbKeyReturn Then
            lvwItems.SelectedItem.Checked = Not lvwItems.SelectedItem.Checked
            Exit Sub
        End If
        If lvwItems.Checkboxes = True And KeyAscii = vbKeySpace Then Exit Sub
        Call lvwItems_DblClick
    Case vbKeyEscape
        picDept.Visible = False
        txtFind.Text = ""
    End Select
End Sub

Private Sub SetSelect(ByVal lvwObj As Object, Optional ByVal blnSelect As Boolean = True)
    Dim i As Integer
    
    With lvwObj
        For i = 1 To .ListItems.Count
            .ListItems(i).Checked = blnSelect
        Next
    End With
End Sub

Private Sub lvwItems_DblClick()
    Dim i As Integer
    Dim m As Integer
    Dim blnBatch As Boolean
    Dim str���� As String
    Dim str����IDs As String
    Dim strTmp As String
    Dim varArr As Variant
    Dim n As Integer
    Dim strNew As String
    Dim blnNew As Boolean
    
    Dim intIndex As Integer
        
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    
    Select Case GetCurҳ��
        Case "���˿�����ҩ����������鼰���۱�"
            intIndex = 1
        Case "���ﴦ��������ҩ�����"
            intIndex = 2
        Case "סԺ���˿�����ҩ�����"
            intIndex = 3
        Case "����ҩ��ʹ���������ͳ��"
            intIndex = 0
        Case "�����п�Χ����Ԥ����ҩͳ��"
            intIndex = 4
        Case "�ż��ﴦ��������ҩͳ��"
            intIndex = 5
        Case "סԺҽ��������ҩͳ��"
            intIndex = 6
        Case "���󿹾�ҩ��ʹ�ó�N��ͳ��"
            intIndex = 7
        Case "ҽ������ĳ����������ҩ�ɱ�ͳ��"
            intIndex = 8
    End Select
   
    For i = 1 To lvwItems.ListItems.Count
        If lvwItems.ListItems(i).Checked Then
            strTmp = Mid(lvwItems.ListItems(i).Key, 2) & "," & lvwItems.ListItems(i).Text
            If InStr(str����, strTmp) = 0 Then str���� = str���� & ";" & strTmp
        End If
    Next
    If str���� = "" Then
        txtDept(intIndex).Text = "���п���"
        txtDept(intIndex).ToolTipText = "���п���"
        txtDept(intIndex).Tag = ""
        picDept.Visible = False
        txtFind.Text = ""
        Exit Sub
    End If
    str���� = Mid(str����, 2)
    
    varArr = Split(str����, ";"): strTmp = ""
    
    For i = 0 To UBound(varArr)
        strTmp = strTmp & "," & Split(varArr(i), ",")(1)
        str����IDs = str����IDs & "," & Split(varArr(i), ",")(0)
    Next
    
    txtDept(intIndex).Text = Mid(strTmp, 2)
    txtDept(intIndex).ToolTipText = txtDept(intIndex).Text
    txtDept(intIndex).Tag = Mid(str����IDs, 2)
    picDept.Visible = False
    txtFind.Text = ""
End Sub

Private Sub picDept_LostFocus()
    Dim strActive As String
    
    strActive = UCase(Me.ActiveControl.Name)
    
    If InStr(1, "CMDFINDCANCLE,LVWITEMS,PICDEPT,TXTFIND,CMDFIND,CMDFINDOK", strActive) <> 0 Then
        Exit Sub
    End If

    picDept.Visible = False
    txtFind.Text = ""
    mlngFind = 1
End Sub

Private Sub lvwItems_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwItems.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwItems.SortOrder = IIf(Me.lvwItems.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwItems.SortKey = ColumnHeader.Index - 1
        Me.lvwItems.SortOrder = lvwAscending
    End If
End Sub

Private Sub cmdDept_Click(Index As Integer)
'���ܣ���ʾ����ѡ����
    Dim rsTmp As New ADODB.Recordset
    Dim objItem As ListItem
    Dim lngTmp  As Long
    Dim i As Integer
    
    With Me.picDept
        .Left = txtDept(Index).Left
        .Width = txtDept(Index).Width + 700
        .Top = txtDept(Index).Top + txtDept(Index).Height + picReportSub(PanelItem_����ҩƷ���Ľ������).Top + picReport.Top + 950
        cmdFind.Visible = True
        txtFind.Visible = True
        cmdFindOk.Visible = True
        cmdFindCancle.Visible = True
        .ZOrder 0
        .Visible = True
    End With

    With Me.lvwItems
        .Left = 0
        .Top = txtFind.Height + 100
        .Width = Me.picDept.Width
        .Height = Me.picDept.Height - txtFind.Height - 50 - 50
        txtFind.Top = 50
        cmdFind.Top = 50
        cmdFindOk.Left = .Width + .Left - cmdFind.Width - 80 - cmdFindCancle.Width
        cmdFindCancle.Left = .Width + .Left - cmdFind.Width - 50
        cmdFindOk.Top = cmdFind.Top
        cmdFindCancle.Top = cmdFind.Top
        .SetFocus
        .Refresh
    End With
    
    Call SetSelect(lvwItems, False)
    If txtDept(Index).Tag = "" Then Exit Sub
   
    For i = 1 To lvwItems.ListItems.Count
        lngTmp = Val(Mid(lvwItems.ListItems(i).Key, 2))
        Me.lvwItems.ListItems(i).Checked = InStr("," & txtDept(Index).Tag & ",", "," & lngTmp & ",") > 0
    Next
End Sub

Private Sub cmdCYSel_Click()
'���ܣ�������¼ѡ����---------------------------------------
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim blnCanle As Boolean
    Dim x As Long, y As Long
    
    x = Me.Left + tbcSub.Left + 1150
    y = Me.Top + tbcSub.Top + 1900
            
    strSql = "Select ID,������,To_Char(����ʱ��, 'YYYY-MM-DD HH24:MI:SS') as ����ʱ�� From ����ҩ�������¼ order by ����ʱ�� desc"
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, "����ҩ�������¼", False, "", "", False, False, True, x, y, txtCYJL.Height, blnCanle, False, True)
    If blnCanle Then Exit Sub
    If rsTmp Is Nothing Then
        MsgBox "Ŀǰû�г�����¼������ִ�г�����", vbInformation, gstrSysName
        Exit Sub
    End If
    txtCYJL.Text = "����ʱ�䣺" & rsTmp!����ʱ�� & "  �����ˣ�" & rsTmp!������
    txtCYJL.Tag = rsTmp!ID
    Call LoadPati
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    If Not Me.Visible Then Exit Sub
    
    With Me.tbcSub
    
        .Left = lngLeft
        .Top = lngTop
        .Height = lngBottom - lngTop
        .Width = lngRight - lngLeft
    
    End With
    Me.Refresh
End Sub

Private Function GetDiagSQL(ByVal strInput As String) As String
'���ܣ���ò�ѯ��ϵ�SQL
'������strInput-��ѯ����,strsql--���ص�SQL
'���أ�strsql--��ѯ��ҽ��ϵ�SQL
    Dim strSql As String
    
    If optType(26).Value Then  '��ҽ���
        If optType(27).Value Then    ' ����ϱ�׼
            '���������:��ҽ���ݣ�һ����Ͽ������ڶ������
            If zlCommFun.IsCharChinese(strInput) Then
                strSql = "B.���� Like [2]" '���뺺��ʱֻƥ������
            Else
                strSql = "A.���� Like [1] Or B.���� Like [2] Or B.���� Like [2]"
            End If
           strSql = _
                " Select Distinct A.ID,A.ID as ��ĿID,A.����,Null as ���,A.����,A.˵��,A.����," & vbNewLine & _
                " Decode(b.����, [4], 1, Decode(b.����,[4],1,decode(a.����,[4],1,NULL))) As ����1ID,Decode(d.���id, Null, Decode(c.���id, Null, Null, 2), 1) As ����2ID," & vbNewLine & _
                " Decode(Substr(b.����, 1, Length([4])), [4], 1, Decode(Substr(b.����, 1, Length([4])),[4],1,decode(Substr(a.����, 1, Length([4])),[4],1,NULL))) As ����3ID" & _
                " From �������Ŀ¼ A,������ϱ��� B, ������Ͽ��� C, ������Ͽ��� D" & _
                " Where A.ID=B.���ID And c.���id(+) = a.Id And d.���id(+) = a.Id And A.���=2" & _
                " And B.����=[3] And d.��Աid(+) = [5] And (c.����id In (Select ����id From ������Ա Where ��Աid = [5]) Or c.����id Is Null) " & _
                " And (" & strSql & ")" & _
                " Order by ����1ID, ����2ID, ����3ID,A.����"
                '����˳��������ȫƥ��(���ơ����롢���룩�������ղء�����ǿ����ղء�Ȼ������ƥ��(���ơ����롢���룩�������˫��ƥ��
        Else
            'B-��ҽ��������
            If zlCommFun.IsCharChinese(strInput) Then
                strSql = "A.���� Like [2]" '���뺺��ʱֻƥ������
            Else
                strSql = "A.���� Like [1] Or A.���� Like [2] Or " & IIf(mint���� = 0, "A.����", "A.�����") & " Like [2]"
            End If
            strSql = _
                "Select Distinct a.Id, a.Id As ��Ŀid, a.����, a.���, a.����, a.����," & IIf(mint���� = 0, "A.����", "A.����� as ����") & ", a.˵��," & _
                " Decode(a.����, [4], 1, Decode(" & IIf(mint���� = 0, "A.����", "A.�����") & ",[4],1,decode(a.����,[4],1,NULL))) As ����1ID," & vbNewLine & _
                "                Decode(d.����id, Null, Decode(c.����id, Null, Null, 2), 1) As ����2ID," & vbNewLine & _
                "                Decode(Substr(a.����, 1, Length([4])), [4], 1, Decode(Substr(" & IIf(mint���� = 0, "A.����", "A.�����") & ", 1, Length([4])),[4],1,decode(Substr(a.����, 1, Length([4])),[4],1,NULL))) As ����3ID" & vbNewLine & _
                "From ��������Ŀ¼ A, ����������� C, ����������� D" & vbNewLine & _
                "Where a.��� = 'B' And (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And c.����id(+) = a.Id And" & vbNewLine & _
                "      d.����id(+) = a.Id And (c.����id In (Select ����id From ������Ա Where ��Աid = [5]) Or c.����id Is Null) And d.��Աid(+) = [5] And (" & strSql & ")" & _
                "Order By ����1ID, ����2ID, ����3ID, ����"
        End If
    Else
        If optType(27).Value Then
            '���������:��ҽ���ݣ�һ����Ͽ������ڶ������
            If zlCommFun.IsCharChinese(strInput) Then
                strSql = "B.���� Like [2]" '���뺺��ʱ,ֻƥ������
            Else
                strSql = "A.���� Like [1] Or B.���� Like [2] Or B.���� Like [2]"
            End If
            strSql = _
                " Select Distinct A.ID,A.ID as ��ĿID,A.����,Null as ���,A.����,A.˵��,A.����," & vbNewLine & _
                " Decode(b.����, [4], 1, Decode(b.����,[4],1,decode(a.����,[4],1,NULL))) As ����1ID,Decode(d.���id, Null, Decode(c.���id, Null, Null, 2), 1) As ����2ID," & vbNewLine & _
                " Decode(Substr(b.����, 1, Length([4])), [4], 1, Decode(Substr(b.����, 1, Length([4])),[4],1,decode(Substr(a.����, 1, Length([4])),[4],1,NULL))) As ����3ID" & _
                " From �������Ŀ¼ A,������ϱ��� B, ������Ͽ��� C, ������Ͽ��� D" & _
                " Where A.ID=B.���ID And c.���id(+) = a.Id And d.���id(+) = a.Id And A.���=1" & _
                " And B.����=[3] And d.��Աid(+) = [5] And (c.����id In (Select ����id From ������Ա Where ��Աid = [5]) Or c.����id Is Null) " & _
                " And (" & strSql & ")" & _
                " Order by ����1ID, ����2ID, ����3ID,A.����"
                '����˳��������ȫƥ��(���ơ����롢���룩�������ղء�����ǿ����ղء�Ȼ������ƥ��(���ơ����롢���룩�������˫��ƥ��
        Else
            'D-ICD-10��������
            If zlCommFun.IsCharChinese(strInput) Then
                strSql = "A.���� Like [2]" '���뺺��ʱ,ֻƥ������
            Else
                strSql = "A.���� Like [1] Or A.���� Like [2] Or " & IIf(mint���� = 0, "A.����", "A.�����") & " Like [2]"
            End If
            strSql = _
                "Select Distinct a.Id, a.Id As ��Ŀid, a.����, a.���, a.����, a.����," & IIf(mint���� = 0, "A.����", "A.����� as ����") & ", a.˵��," & _
                " Decode(a.����, [4], 1, Decode(" & IIf(mint���� = 0, "A.����", "A.�����") & ",[4],1,decode(a.����,[4],1,NULL))) As ����1ID," & vbNewLine & _
                "                Decode(d.����id, Null, Decode(c.����id, Null, Null, 2), 1) As ����2ID," & vbNewLine & _
                "                Decode(Substr(a.����, 1, Length([4])), [4], 1, Decode(Substr(" & IIf(mint���� = 0, "A.����", "A.�����") & ", 1, Length([4])),[4],1,decode(Substr(a.����, 1, Length([4])),[4],1,NULL))) As ����3ID" & vbNewLine & _
                "From ��������Ŀ¼ A, ����������� C, ����������� D" & vbNewLine & _
                "Where a.��� = 'D' And (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And c.����id(+) = a.Id And" & vbNewLine & _
                "      d.����id(+) = a.Id And (c.����id In (Select ����id From ������Ա Where ��Աid = [5]) Or c.����id Is Null) And d.��Աid(+) = [5] And (" & strSql & ")" & _
                "Order By ����1ID, ����2ID, ����3ID, ����"

        End If
    End If
    GetDiagSQL = strSql
End Function

Private Sub txtILL_Validate(Cancel As Boolean)
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    
    If Trim(txtILL.Text) = "" Then
        txtILL.Tag = ""
        Exit Sub
    End If
    strTmp = Trim(txtILL.Text)
    If strTmp = cmdILL.Tag Then Exit Sub
    On Error GoTo errH
    strSql = GetDiagSQL(txtILL.Text)
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strTmp, strTmp, mint���� + 1, strTmp, UserInfo.ID)
    If rsTmp.RecordCount = 1 Then
        txtILL.Tag = rsTmp!��ĿID
        txtILL.Text = "(" & rsTmp!���� & ")" & rsTmp!����
        cmdILL.Tag = txtILL.Text
    Else
        MsgBox "δ�ҵ���Ӧ��Ŀ��", vbInformation, gstrSysName
        Cancel = True
        Call zlControl.TxtSelAll(txtILL)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Function zlReportToVSFlexGrid(vfgList As VSFlexGrid, rptList As ReportControl) As Boolean
    '-------------------------------------------------
    '��ȫ����ǿ��չ��,�������ݱ��
    Dim rptCol As ReportColumn
    Dim rptRcd As ReportRecord
    Dim rptItem As ReportRecordItem
    Dim rptRow As ReportRow
    
    Dim lngCol As Long, lngRow As Long
    
    On Error GoTo errH:
    For Each rptRow In rptList.Rows
        If rptRow.GroupRow Then rptRow.Expanded = True
    Next
    
    With vfgList
        .Clear
        .Rows = rptList.Records.Count + 1
        .Cols = 0: .Cols = rptList.Columns.Count
        .FixedCols = rptList.GroupsOrder.Count
        
        '�����и���
        .Row = 0
        lngCol = 0
        For Each rptCol In rptList.GroupsOrder
            .TextMatrix(0, lngCol) = rptCol.Caption
            .ColData(lngCol) = rptCol.ItemIndex
            Select Case rptCol.Alignment
            Case xtpAlignmentLeft: .FixedAlignment(lngCol) = flexAlignLeftCenter
            Case xtpAlignmentCenter: .FixedAlignment(lngCol) = flexAlignCenterCenter
            Case xtpAlignmentRight:  .FixedAlignment(lngCol) = flexAlignRightCenter
            End Select
            .Cell(flexcpAlignment, 0, lngCol, .FixedRows - 1) = flexAlignCenterCenter
            .Cell(flexcpAlignment, .FixedRows, lngCol, .Rows - 1) = .FixedAlignment(lngCol)
            .ColWidth(lngCol) = rptCol.Width * 15
            .MergeCol(lngCol) = True
            lngCol = lngCol + 1
        Next
        For Each rptCol In rptList.Columns
            If rptCol.Visible Then
                .TextMatrix(0, lngCol) = rptCol.Caption
                .ColData(lngCol) = rptCol.ItemIndex
                Select Case rptCol.Alignment
                Case xtpAlignmentLeft: .ColAlignment(lngCol) = flexAlignLeftCenter
                Case xtpAlignmentCenter: .ColAlignment(lngCol) = flexAlignCenterCenter
                Case xtpAlignmentRight: .ColAlignment(lngCol) = flexAlignRightCenter
                End Select
                .Cell(flexcpAlignment, 0, lngCol, .FixedRows - 1) = flexAlignCenterCenter
                .Cell(flexcpAlignment, .FixedRows, lngCol, .Rows - 1) = .ColAlignment(lngCol)
                If rptCol.Width < 20 Then
                    .ColWidth(lngCol) = 0
                Else
                    .ColWidth(lngCol) = rptCol.Width * 15
                End If
                lngCol = lngCol + 1
            End If
        Next
        vfgList.Cols = lngCol
        
        '�����и���
        lngRow = 0
        For Each rptRow In rptList.Rows
            If rptRow.GroupRow = False Then
                lngRow = lngRow + 1
                For lngCol = 0 To .Cols - 1
                    .TextMatrix(lngRow, lngCol) = rptRow.Record(.ColData(lngCol)).Value
                Next
            End If
        Next
        
        .RowHeight(-1) = .RowHeightMin
    End With
    zlReportToVSFlexGrid = True
    Exit Function
errH:
    zlReportToVSFlexGrid = False

End Function

Private Function zlCopyVSFlexGrid(vsgTemp As VSFlexGrid, vsgCopy As VSFlexGrid) As Boolean
'����: ��vsgCopy�Ŀɼ����е����ݸ��Ƶ�vsgTemp�� , ����Excel���
'����:
'     vsgTemp-���ƺ�Ķ���
'     vsgCopy-�����ƵĶ���
'     strMsg -��ʾ��Ϣ
    Dim i As Long
    Dim j As Long
    Dim lngCol As Long
    Dim lngRow As Long
    Dim lngTmp As Long
    
    On Error GoTo errH:
    
    With vsgTemp
        .Rows = 0: .Cols = 0
        .Rows = vsgCopy.Rows
        .FixedRows = vsgCopy.FixedRows
        .MergeCells = vsgCopy.MergeCells
        
        '����
        lngCol = 0
        For i = 0 To vsgCopy.Cols - 1 '��
            If Not vsgCopy.ColHidden(i) Then
                
                .Cols = .Cols + 1
                .ColWidth(lngCol) = vsgCopy.ColWidth(i)
                lngRow = 0: lngTmp = 0
                
                For j = 0 To vsgCopy.Rows - 1 '��
                    If Not vsgCopy.RowHidden(j) Then
                        .ColAlignment(lngCol) = vsgCopy.ColAlignment(i)
                        .Cell(flexcpAlignment, lngRow, lngCol) = vsgCopy.Cell(flexcpAlignment, j, i)  '���뷽ʽ
                        .TextMatrix(lngRow, lngCol) = vsgCopy.TextMatrix(j, i)
                        lngRow = lngRow + 1
                    Else
                        lngTmp = lngTmp + 1  '��¼������
                    End If
                Next
                lngCol = lngCol + 1
            End If
        Next
        '
        .Rows = .Rows - lngTmp 'ɾ��������
        .FixedCols = vsgCopy.FixedCols
        .RowHeight(-1) = vsgCopy.RowHeightMin
    End With
    
    zlCopyVSFlexGrid = True
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetCurҳ��() As String
'���ܣ���ǰ���Ǹ�ҳ��
    Select Case tbcSub.Selected.Index
        Case 0
            Select Case tbcReport.Selected.Index
                Case 0
                    GetCurҳ�� = "����ҩƷ���Ľ������"
                Case 1
                    GetCurҳ�� = "���˿�����ҩ����������鼰���۱�"
                Case 2
                    GetCurҳ�� = "���ﴦ��������ҩ�����"
                Case 3
                    GetCurҳ�� = "סԺ���˿�����ҩ�����"
            End Select
        Case 1
            Select Case tbcOther.Selected.Index
                Case 0
                    GetCurҳ�� = "����ҩ��ʹ���������ͳ��"
                Case 1
                    GetCurҳ�� = "�����п�Χ����Ԥ����ҩͳ��"
                Case 2
                    GetCurҳ�� = "�ż��ﴦ��������ҩͳ��"
                Case 3
                    GetCurҳ�� = "סԺҽ��������ҩͳ��"
                Case 4
                    GetCurҳ�� = "���󿹾�ҩ��ʹ�ó�N��ͳ��"
                Case 5
                    GetCurҳ�� = "ҽ������ĳ����������ҩ�ɱ�ͳ��"
            End Select
    End Select
End Function

