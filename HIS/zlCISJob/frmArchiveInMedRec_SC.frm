VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmArchiveInMedRec_SC 
   BorderStyle     =   0  'None
   ClientHeight    =   8370
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10395
   LinkTopic       =   "Form1"
   ScaleHeight     =   8370
   ScaleWidth      =   10395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8235
      Left            =   120
      ScaleHeight     =   8235
      ScaleWidth      =   10245
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   120
      Width           =   10245
      Begin VB.Frame fraVH 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   9840
         TabIndex        =   61
         Top             =   7920
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.VScrollBar vsc 
         Height          =   7800
         Left            =   9840
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.HScrollBar hsc 
         Height          =   255
         Left            =   90
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   7920
         Visible         =   0   'False
         Width           =   9675
      End
      Begin MSComctlLib.ImageList imgSize 
         Left            =   960
         Top             =   5190
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   9
         ImageHeight     =   9
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmArchiveInMedRec_SC.frx":0000
               Key             =   "-"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmArchiveInMedRec_SC.frx":04EA
               Key             =   "+"
            EndProperty
         EndProperty
      End
      Begin VB.Frame fraBack 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   7800
         Left            =   90
         TabIndex        =   62
         Top             =   0
         Width           =   9675
         Begin VB.Frame fraInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "   �����뻯�� "
            ForeColor       =   &H00FF0000&
            Height          =   5010
            Index           =   5
            Left            =   120
            TabIndex        =   153
            Tag             =   "5010"
            Top             =   120
            Width           =   9495
            Begin VB.PictureBox picSize 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   165
               Index           =   5
               Left            =   180
               ScaleHeight     =   135
               ScaleWidth      =   135
               TabIndex        =   154
               TabStop         =   0   'False
               Top             =   0
               Width           =   165
            End
            Begin VSFlex8Ctl.VSFlexGrid vsChemotherapy 
               Height          =   1635
               Left            =   120
               TabIndex        =   186
               Top             =   480
               Width           =   9240
               _cx             =   16298
               _cy             =   2884
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
               BackColor       =   16777215
               ForeColor       =   -2147483640
               BackColorFixed  =   16777215
               ForeColorFixed  =   4210752
               BackColorSel    =   4210752
               ForeColorSel    =   -2147483634
               BackColorBkg    =   16777215
               BackColorAlternate=   16777215
               GridColor       =   -2147483636
               GridColorFixed  =   -2147483636
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483643
               FocusRect       =   3
               HighLight       =   2
               AllowSelection  =   0   'False
               AllowBigSelection=   0   'False
               AllowUserResizing=   1
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   1
               GridLineWidth   =   1
               Rows            =   3
               Cols            =   7
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   250
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmArchiveInMedRec_SC.frx":09D4
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
               AllowUserFreezing=   0
               BackColorFrozen =   0
               ForeColorFrozen =   0
               WallPaperAlignment=   9
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   24
            End
            Begin VSFlex8Ctl.VSFlexGrid vsRadiotherapy 
               Height          =   2205
               Left            =   120
               TabIndex        =   187
               Top             =   2640
               Width           =   9240
               _cx             =   16298
               _cy             =   3889
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
               BackColor       =   16777215
               ForeColor       =   -2147483640
               BackColorFixed  =   16777215
               ForeColorFixed  =   4210752
               BackColorSel    =   4210752
               ForeColorSel    =   -2147483634
               BackColorBkg    =   16777215
               BackColorAlternate=   16777215
               GridColor       =   -2147483636
               GridColorFixed  =   -2147483636
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483643
               FocusRect       =   3
               HighLight       =   2
               AllowSelection  =   0   'False
               AllowBigSelection=   0   'False
               AllowUserResizing=   1
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   1
               GridLineWidth   =   1
               Rows            =   3
               Cols            =   7
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   250
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmArchiveInMedRec_SC.frx":0AEA
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
               AllowUserFreezing=   0
               BackColorFrozen =   0
               ForeColorFrozen =   0
               WallPaperAlignment=   9
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   24
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "���Ƽ�¼��Ϣ"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   138
               Left            =   120
               TabIndex        =   180
               Top             =   2400
               Width           =   1080
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "���Ƽ�¼��Ϣ"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   137
               Left            =   120
               TabIndex        =   179
               Top             =   240
               Width           =   1080
            End
         End
         Begin VB.Frame fraInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "   סԺ��� "
            ForeColor       =   &H00FF0000&
            Height          =   7650
            Index           =   4
            Left            =   120
            TabIndex        =   218
            Tag             =   "7650"
            Top             =   120
            Width           =   9495
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "ʾ�̲���"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   9
               Left            =   540
               TabIndex        =   336
               Top             =   1469
               Width           =   1020
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   98
               Left            =   4800
               Locked          =   -1  'True
               TabIndex        =   271
               Top             =   4087
               Width           =   720
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   99
               Left            =   6360
               Locked          =   -1  'True
               TabIndex        =   270
               Top             =   4087
               Width           =   2400
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "�������"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   16
               Left            =   180
               TabIndex        =   269
               Top             =   4080
               Width           =   1020
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   92
               Left            =   5415
               Locked          =   -1  'True
               TabIndex        =   268
               Top             =   2964
               Width           =   3480
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   91
               Left            =   975
               Locked          =   -1  'True
               TabIndex        =   267
               Top             =   2964
               Width           =   2640
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   121
               Left            =   4320
               Locked          =   -1  'True
               TabIndex        =   266
               Top             =   7080
               Width           =   1335
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   120
               Left            =   1035
               Locked          =   -1  'True
               TabIndex        =   265
               Top             =   7080
               Width           =   1335
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   96
               Left            =   3720
               Locked          =   -1  'True
               TabIndex        =   264
               Top             =   3708
               Width           =   5295
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   95
               Left            =   915
               Locked          =   -1  'True
               TabIndex        =   263
               Top             =   3708
               Width           =   1335
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   82
               Left            =   4320
               Locked          =   -1  'True
               TabIndex        =   262
               Top             =   1104
               Width           =   1200
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   80
               Left            =   4320
               Locked          =   -1  'True
               TabIndex        =   261
               Top             =   732
               Width           =   1200
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   78
               Left            =   4320
               Locked          =   -1  'True
               TabIndex        =   260
               Top             =   360
               Width           =   1200
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   77
               Left            =   960
               Locked          =   -1  'True
               TabIndex        =   259
               Top             =   360
               Width           =   1440
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   106
               Left            =   2190
               Locked          =   -1  'True
               MaxLength       =   9
               TabIndex        =   258
               Text            =   "��"
               Top             =   4839
               Width           =   360
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   107
               Left            =   3120
               Locked          =   -1  'True
               TabIndex        =   257
               Top             =   4839
               Width           =   5940
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   89
               Left            =   3690
               Locked          =   -1  'True
               TabIndex        =   256
               Top             =   2592
               Width           =   1080
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   108
               Left            =   1155
               Locked          =   -1  'True
               TabIndex        =   255
               Top             =   5211
               Width           =   720
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   105
               Left            =   8070
               Locked          =   -1  'True
               TabIndex        =   254
               Top             =   4467
               Width           =   435
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   104
               Left            =   7140
               Locked          =   -1  'True
               TabIndex        =   253
               Top             =   4467
               Width           =   435
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   103
               Left            =   6180
               Locked          =   -1  'True
               TabIndex        =   252
               Top             =   4467
               Width           =   480
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   102
               Left            =   4620
               Locked          =   -1  'True
               TabIndex        =   251
               Top             =   4467
               Width           =   435
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   100
               Left            =   2940
               Locked          =   -1  'True
               TabIndex        =   250
               Top             =   4467
               Width           =   480
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   118
               Left            =   4320
               Locked          =   -1  'True
               TabIndex        =   249
               Top             =   6720
               Width           =   1335
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   122
               Left            =   7650
               Locked          =   -1  'True
               TabIndex        =   248
               Top             =   7080
               Width           =   1335
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   117
               Left            =   1035
               Locked          =   -1  'True
               TabIndex        =   247
               Top             =   6720
               Width           =   1335
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   83
               Left            =   4320
               Locked          =   -1  'True
               TabIndex        =   246
               Top             =   1476
               Width           =   1200
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   87
               Left            =   4320
               Locked          =   -1  'True
               TabIndex        =   245
               Top             =   2220
               Width           =   1200
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   86
               Left            =   915
               TabIndex        =   244
               Top             =   2220
               Width           =   1440
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   88
               Left            =   915
               Locked          =   -1  'True
               TabIndex        =   243
               Top             =   2592
               Width           =   1440
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   85
               Left            =   4320
               Locked          =   -1  'True
               TabIndex        =   242
               Top             =   1848
               Width           =   1200
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   84
               Left            =   915
               Locked          =   -1  'True
               TabIndex        =   241
               Top             =   1848
               Width           =   1440
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   110
               Left            =   7770
               Locked          =   -1  'True
               TabIndex        =   240
               Top             =   5583
               Width           =   1440
            End
            Begin VB.CheckBox chkInfo 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "����"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   17
               Left            =   3480
               TabIndex        =   239
               Top             =   5576
               Width           =   660
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   79
               Left            =   915
               Locked          =   -1  'True
               TabIndex        =   238
               Top             =   732
               Width           =   1440
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   81
               Left            =   915
               Locked          =   -1  'True
               TabIndex        =   237
               Top             =   1104
               Width           =   1440
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   109
               Left            =   975
               Locked          =   -1  'True
               TabIndex        =   236
               Top             =   5583
               Width           =   1335
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   111
               Left            =   1680
               Locked          =   -1  'True
               TabIndex        =   235
               Top             =   6015
               Width           =   1335
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   114
               Left            =   915
               Locked          =   -1  'True
               TabIndex        =   234
               Top             =   6375
               Width           =   1335
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   112
               Left            =   4320
               Locked          =   -1  'True
               TabIndex        =   233
               Top             =   6015
               Width           =   1335
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   115
               Left            =   4320
               Locked          =   -1  'True
               TabIndex        =   232
               Top             =   6375
               Width           =   1335
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   113
               Left            =   7650
               Locked          =   -1  'True
               TabIndex        =   231
               Top             =   6030
               Width           =   1575
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   116
               Left            =   7650
               Locked          =   -1  'True
               TabIndex        =   230
               Top             =   6360
               Width           =   1335
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   119
               Left            =   7680
               Locked          =   -1  'True
               TabIndex        =   229
               Top             =   6735
               Width           =   1335
            End
            Begin VB.PictureBox picSize 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   165
               Index           =   4
               Left            =   180
               ScaleHeight     =   135
               ScaleWidth      =   135
               TabIndex        =   228
               TabStop         =   0   'False
               Top             =   0
               Width           =   165
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "���в���"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   18
               Left            =   1860
               TabIndex        =   227
               Top             =   1469
               Width           =   1020
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   101
               Left            =   3750
               Locked          =   -1  'True
               TabIndex        =   226
               Top             =   4467
               Width           =   435
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   97
               Left            =   2760
               Locked          =   -1  'True
               TabIndex        =   225
               Top             =   4087
               Width           =   720
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "����·��"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   19
               Left            =   1560
               TabIndex        =   224
               Top             =   3329
               Width           =   1020
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "���·��"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   20
               Left            =   2640
               TabIndex        =   223
               Top             =   3329
               Width           =   1020
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "����"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   21
               Left            =   6000
               TabIndex        =   222
               Top             =   3329
               Width           =   780
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   93
               Left            =   4560
               Locked          =   -1  'True
               TabIndex        =   221
               Top             =   3336
               Width           =   1335
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   90
               Left            =   6060
               Locked          =   -1  'True
               TabIndex        =   220
               Top             =   2592
               Width           =   1080
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   94
               Left            =   7560
               Locked          =   -1  'True
               TabIndex        =   219
               Top             =   3330
               Width           =   1335
            End
            Begin VSFlex8Ctl.VSFlexGrid vsCheck 
               Height          =   2055
               Left            =   6480
               TabIndex        =   272
               Top             =   360
               Width           =   2415
               _cx             =   4260
               _cy             =   3625
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
               BackColor       =   16777215
               ForeColor       =   -2147483640
               BackColorFixed  =   16777215
               ForeColorFixed  =   4210752
               BackColorSel    =   4210752
               ForeColorSel    =   -2147483634
               BackColorBkg    =   16777215
               BackColorAlternate=   16777215
               GridColor       =   -2147483636
               GridColorFixed  =   -2147483636
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483643
               FocusRect       =   3
               HighLight       =   2
               AllowSelection  =   -1  'True
               AllowBigSelection=   -1  'True
               AllowUserResizing=   1
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   1
               GridLineWidth   =   1
               Rows            =   8
               Cols            =   2
               FixedRows       =   0
               FixedCols       =   1
               RowHeightMin    =   250
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmArchiveInMedRec_SC.frx":0C10
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
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ժ�ڻ���         ��   ��Ժ����          �Σ�����"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   97
               Left            =   2010
               TabIndex        =   326
               Top             =   4087
               Width           =   4320
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   92
               X1              =   5415
               X2              =   9000
               Y1              =   3145
               Y2              =   3145
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����ҽѧ��ʾ"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   92
               Left            =   4320
               TabIndex        =   325
               Top             =   2964
               Width           =   1080
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   91
               X1              =   975
               X2              =   3840
               Y1              =   3145
               Y2              =   3145
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ҽѧ��ʾ"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   91
               Left            =   180
               TabIndex        =   324
               Top             =   2964
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   121
               X1              =   4320
               X2              =   5745
               Y1              =   7260
               Y2              =   7260
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��������"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   121
               Left            =   3540
               TabIndex        =   323
               Top             =   7080
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   120
               X1              =   1035
               X2              =   2460
               Y1              =   7260
               Y2              =   7260
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�ʿ�����"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   120
               Left            =   180
               TabIndex        =   322
               Top             =   7080
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ת�����"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   96
               Left            =   2940
               TabIndex        =   321
               Top             =   3708
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��Ժ��ʽ"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   95
               Left            =   180
               TabIndex        =   320
               Top             =   3708
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   95
               X1              =   915
               X2              =   2400
               Y1              =   3885
               Y2              =   3885
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   96
               X1              =   3690
               X2              =   9120
               Y1              =   3885
               Y2              =   3885
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   82
               X1              =   4320
               X2              =   5640
               Y1              =   1290
               Y2              =   1290
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����״��"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   82
               Left            =   3540
               TabIndex        =   319
               Top             =   1104
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   80
               X1              =   4320
               X2              =   5640
               Y1              =   915
               Y2              =   915
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����ʱ��"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   80
               Left            =   3540
               TabIndex        =   318
               Top             =   732
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   78
               X1              =   4320
               X2              =   5640
               Y1              =   540
               Y2              =   540
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��Ѫǰ9����"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   78
               Left            =   3090
               TabIndex        =   317
               Top             =   360
               Width           =   1170
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   77
               X1              =   915
               X2              =   2400
               Y1              =   535
               Y2              =   535
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��������"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   77
               Left            =   180
               TabIndex        =   316
               Top             =   360
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   107
               X1              =   3000
               X2              =   9120
               Y1              =   5020
               Y2              =   5020
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��Ժ31��������Ժ�ƻ�"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   106
               Left            =   180
               TabIndex        =   315
               Top             =   4839
               Width           =   1800
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   106
               X1              =   2070
               X2              =   2650
               Y1              =   5020
               Y2              =   5020
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ŀ��"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   107
               Left            =   2715
               TabIndex        =   314
               Top             =   4839
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ml"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   189
               Left            =   4800
               TabIndex        =   313
               Top             =   2595
               Width           =   180
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�������"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   89
               Left            =   2940
               TabIndex        =   312
               Top             =   2595
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   140
               X1              =   3690
               X2              =   4800
               Y1              =   2775
               Y2              =   2775
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   108
               X1              =   1155
               X2              =   2040
               Y1              =   5385
               Y2              =   5385
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "������ʹ��"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   108
               Left            =   180
               TabIndex        =   311
               Top             =   5211
               Width           =   900
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Сʱ"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   208
               Left            =   2145
               TabIndex        =   310
               Top             =   5211
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   205
               Left            =   8535
               TabIndex        =   309
               Top             =   4467
               Width           =   360
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   105
               X1              =   7935
               X2              =   8530
               Y1              =   4645
               Y2              =   4645
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Сʱ"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   105
               Left            =   7575
               TabIndex        =   308
               Top             =   4467
               Width           =   360
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   104
               X1              =   6990
               X2              =   7560
               Y1              =   4645
               Y2              =   4645
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   104
               Left            =   6840
               TabIndex        =   307
               Top             =   4467
               Width           =   180
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   103
               X1              =   6180
               X2              =   6760
               Y1              =   4645
               Y2              =   4645
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��Ժ��"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   103
               Left            =   5640
               TabIndex        =   306
               Top             =   4467
               Width           =   540
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   202
               Left            =   5175
               TabIndex        =   305
               Top             =   4467
               Width           =   360
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   102
               X1              =   4545
               X2              =   5115
               Y1              =   4645
               Y2              =   4645
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   101
               X1              =   3660
               X2              =   4260
               Y1              =   4645
               Y2              =   4645
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   101
               Left            =   3480
               TabIndex        =   304
               Top             =   4467
               Width           =   180
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   100
               X1              =   2940
               X2              =   3495
               Y1              =   4645
               Y2              =   4645
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "­�����˻��߻���ʱ��;   ��Ժǰ"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   100
               Left            =   180
               TabIndex        =   303
               Top             =   4467
               Width           =   2700
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "���λ�ʿ"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   118
               Left            =   3540
               TabIndex        =   302
               Top             =   6720
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   118
               X1              =   4320
               X2              =   5745
               Y1              =   6900
               Y2              =   6900
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   122
               X1              =   7650
               X2              =   9240
               Y1              =   7260
               Y2              =   7260
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   117
               X1              =   1035
               X2              =   2460
               Y1              =   6900
               Y2              =   6900
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�ʿ�ҽʦ"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   117
               Left            =   180
               TabIndex        =   301
               Top             =   6699
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�ʿػ�ʿ"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   122
               Left            =   6900
               TabIndex        =   300
               Top             =   7080
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��Ѫ��Ӧ"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   83
               Left            =   3540
               TabIndex        =   299
               Top             =   1476
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   83
               X1              =   4320
               X2              =   5640
               Y1              =   1650
               Y2              =   1650
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ʵϰҽʦ"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   119
               Left            =   6900
               TabIndex        =   298
               Top             =   6735
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����ҽʦ"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   115
               Left            =   3540
               TabIndex        =   297
               Top             =   6375
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����ҽʦ"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   114
               Left            =   180
               TabIndex        =   296
               Top             =   6327
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "סԺҽʦ"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   116
               Left            =   6900
               TabIndex        =   295
               Top             =   6375
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����ҽʦ"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   112
               Left            =   3540
               TabIndex        =   294
               Top             =   6015
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����(������)ҽʦ"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   111
               Left            =   180
               TabIndex        =   293
               Top             =   5955
               Width           =   1440
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "������"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   113
               Left            =   7080
               TabIndex        =   292
               Top             =   6030
               Width           =   540
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����ҽʦ"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   109
               Left            =   180
               TabIndex        =   291
               Top             =   5583
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "������"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   87
               Left            =   3720
               TabIndex        =   290
               Top             =   2220
               Width           =   540
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ml"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   186
               Left            =   2520
               TabIndex        =   289
               Top             =   2220
               Width           =   180
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��ȫѪ"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   86
               Left            =   360
               TabIndex        =   288
               Top             =   2220
               Width           =   540
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ml"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   188
               Left            =   2520
               TabIndex        =   287
               Top             =   2640
               Width           =   180
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��Ѫ��"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   88
               Left            =   360
               TabIndex        =   286
               Top             =   2592
               Width           =   540
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��λ"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   185
               Left            =   5640
               TabIndex        =   285
               Top             =   1848
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��ѪС��"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   85
               Left            =   3540
               TabIndex        =   284
               Top             =   1848
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��λ"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   184
               Left            =   2520
               TabIndex        =   283
               Top             =   1845
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "���ϸ��"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   84
               Left            =   180
               TabIndex        =   282
               Top             =   1848
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Rh"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   81
               Left            =   720
               TabIndex        =   281
               Top             =   1104
               Width           =   180
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ѫ��"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   79
               Left            =   540
               TabIndex        =   280
               Top             =   732
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��������"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   110
               Left            =   7020
               TabIndex        =   279
               Top             =   5583
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   110
               X1              =   7770
               X2              =   9345
               Y1              =   5760
               Y2              =   5760
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   79
               X1              =   915
               X2              =   2400
               Y1              =   910
               Y2              =   910
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   84
               X1              =   915
               X2              =   2400
               Y1              =   2025
               Y2              =   2025
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   86
               X1              =   915
               X2              =   2400
               Y1              =   2400
               Y2              =   2400
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   81
               X1              =   915
               X2              =   2400
               Y1              =   1290
               Y2              =   1290
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   85
               X1              =   4320
               X2              =   5640
               Y1              =   2025
               Y2              =   2025
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   87
               X1              =   4320
               X2              =   5640
               Y1              =   2400
               Y2              =   2400
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   88
               X1              =   915
               X2              =   2400
               Y1              =   2775
               Y2              =   2775
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   109
               X1              =   975
               X2              =   2400
               Y1              =   5760
               Y2              =   5760
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   111
               X1              =   1680
               X2              =   3105
               Y1              =   6195
               Y2              =   6195
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   114
               X1              =   915
               X2              =   2340
               Y1              =   6555
               Y2              =   6555
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   112
               X1              =   4320
               X2              =   5745
               Y1              =   6195
               Y2              =   6195
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   115
               X1              =   4320
               X2              =   5745
               Y1              =   6555
               Y2              =   6555
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   113
               X1              =   7650
               X2              =   9240
               Y1              =   6210
               Y2              =   6210
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   116
               X1              =   7650
               X2              =   9240
               Y1              =   6555
               Y2              =   6555
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   119
               X1              =   7650
               X2              =   9240
               Y1              =   6915
               Y2              =   6915
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Сʱ"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   102
               Left            =   4215
               TabIndex        =   278
               Top             =   4467
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�ٴ�·����Ϣ��"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   139
               Left            =   180
               TabIndex        =   277
               Top             =   3336
               Width           =   1260
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�˳�ԭ��"
               Height          =   180
               Index           =   93
               Left            =   3720
               TabIndex        =   276
               Top             =   3336
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����ԭ��"
               Height          =   180
               Index           =   94
               Left            =   6840
               TabIndex        =   275
               Top             =   3336
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   90
               X1              =   6060
               X2              =   7230
               Y1              =   2770
               Y2              =   2770
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��׵���"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   90
               Left            =   5280
               TabIndex        =   274
               Top             =   2592
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "g"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   190
               Left            =   7320
               TabIndex        =   273
               Top             =   2595
               Width           =   90
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   93
               X1              =   4440
               X2              =   5880
               Y1              =   3510
               Y2              =   3510
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   94
               X1              =   7560
               X2              =   9000
               Y1              =   3525
               Y2              =   3525
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   97
               X1              =   2760
               X2              =   3480
               Y1              =   4265
               Y2              =   4265
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   98
               X1              =   4800
               X2              =   5520
               Y1              =   4265
               Y2              =   4265
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   99
               X1              =   6360
               X2              =   8880
               Y1              =   4265
               Y2              =   4265
            End
         End
         Begin VB.Frame fraInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "   ��ҽ��� "
            ForeColor       =   &H00FF0000&
            Height          =   5835
            Index           =   1
            Left            =   120
            TabIndex        =   114
            Tag             =   "5835"
            Top             =   120
            Width           =   9495
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "סԺ�ڼ�没�ػ�Σ"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   15
               Left            =   240
               TabIndex        =   210
               Top             =   5400
               Width           =   2325
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "���Ѳ���"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   14
               Left            =   3000
               TabIndex        =   209
               Top             =   5400
               Width           =   1020
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   56
               Left            =   7155
               Locked          =   -1  'True
               TabIndex        =   177
               Top             =   3876
               Width           =   1980
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   49
               Left            =   4080
               Locked          =   -1  'True
               TabIndex        =   176
               Top             =   3120
               Width           =   1515
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   59
               Left            =   4560
               Locked          =   -1  'True
               TabIndex        =   174
               Top             =   4627
               Width           =   4530
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   57
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   172
               Top             =   4248
               Width           =   1635
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   53
               Left            =   7080
               Locked          =   -1  'True
               TabIndex        =   170
               Top             =   3504
               Width           =   1875
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   48
               Left            =   960
               Locked          =   -1  'True
               TabIndex        =   168
               Top             =   3132
               Width           =   1695
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   45
               Left            =   960
               Locked          =   -1  'True
               TabIndex        =   166
               Top             =   2760
               Width           =   1660
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "ҽԺ��Ⱦ����ԭѧ���"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   3
               Left            =   6960
               TabIndex        =   155
               Top             =   4241
               Width           =   2150
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   47
               Left            =   7080
               Locked          =   -1  'True
               TabIndex        =   151
               Top             =   2760
               Width           =   2115
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   62
               Left            =   4335
               Locked          =   -1  'True
               TabIndex        =   143
               Top             =   5010
               Width           =   4770
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   58
               Left            =   3720
               Locked          =   -1  'True
               TabIndex        =   140
               Top             =   4248
               Width           =   2970
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "�·�����"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   5
               Left            =   1680
               TabIndex        =   139
               Top             =   4620
               Width           =   1020
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "��������ʬ��"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   4
               Left            =   240
               TabIndex        =   138
               Top             =   4620
               Width           =   1485
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   61
               Left            =   2910
               Locked          =   -1  'True
               TabIndex        =   38
               Top             =   5010
               Width           =   510
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   60
               Left            =   975
               Locked          =   -1  'True
               TabIndex        =   37
               Top             =   5010
               Width           =   870
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "�Ƿ�ȷ��"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   2
               Left            =   2760
               TabIndex        =   35
               Top             =   2753
               Width           =   1020
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   46
               Left            =   4695
               Locked          =   -1  'True
               TabIndex        =   36
               Top             =   2760
               Width           =   1680
            End
            Begin VB.PictureBox picSize 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   165
               Index           =   1
               Left            =   180
               ScaleHeight     =   135
               ScaleWidth      =   135
               TabIndex        =   115
               TabStop         =   0   'False
               Top             =   0
               Width           =   165
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   50
               Left            =   7080
               Locked          =   -1  'True
               TabIndex        =   41
               Top             =   3132
               Width           =   1755
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   55
               Left            =   4080
               Locked          =   -1  'True
               TabIndex        =   43
               Top             =   3876
               Width           =   1455
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   52
               Left            =   4080
               Locked          =   -1  'True
               TabIndex        =   40
               Top             =   3504
               Width           =   1575
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   54
               Left            =   960
               Locked          =   -1  'True
               TabIndex        =   42
               Top             =   3876
               Width           =   1755
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   51
               Left            =   960
               Locked          =   -1  'True
               TabIndex        =   39
               Top             =   3504
               Width           =   1755
            End
            Begin VSFlex8Ctl.VSFlexGrid vsDiagXY 
               Height          =   2385
               Left            =   135
               TabIndex        =   34
               Top             =   270
               Width           =   9240
               _cx             =   16298
               _cy             =   4207
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
               BackColor       =   16777215
               ForeColor       =   -2147483640
               BackColorFixed  =   16777215
               ForeColorFixed  =   4210752
               BackColorSel    =   4210752
               ForeColorSel    =   -2147483634
               BackColorBkg    =   16777215
               BackColorAlternate=   16777215
               GridColor       =   -2147483636
               GridColorFixed  =   -2147483636
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483643
               FocusRect       =   3
               HighLight       =   2
               AllowSelection  =   0   'False
               AllowBigSelection=   0   'False
               AllowUserResizing=   1
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   1
               GridLineWidth   =   1
               Rows            =   9
               Cols            =   9
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   250
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmArchiveInMedRec_SC.frx":0C88
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
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   56
               X1              =   7065
               X2              =   9240
               Y1              =   4050
               Y2              =   4050
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��ǰ������"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   56
               Left            =   6180
               TabIndex        =   178
               Top             =   3876
               Width           =   900
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   59
               X1              =   4605
               X2              =   9120
               Y1              =   4800
               Y2              =   4800
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ҽԺ��Ⱦ��ԭѧ���"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   59
               Left            =   2960
               TabIndex        =   175
               Top             =   4627
               Width           =   1620
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����ʱ��"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   57
               Left            =   240
               TabIndex        =   173
               Top             =   4248
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   57
               X1              =   960
               X2              =   2760
               Y1              =   4420
               Y2              =   4420
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��������Ժ"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   53
               Left            =   6180
               TabIndex        =   171
               Top             =   3504
               Width           =   900
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   53
               X1              =   7080
               X2              =   9240
               Y1              =   3690
               Y2              =   3690
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����������"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   49
               Left            =   2960
               TabIndex        =   169
               Top             =   3135
               Width           =   1080
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   49
               X1              =   4080
               X2              =   5640
               Y1              =   3315
               Y2              =   3315
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   48
               X1              =   960
               X2              =   2745
               Y1              =   3310
               Y2              =   3310
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�ֻ��̶�"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   48
               Left            =   240
               TabIndex        =   167
               Top             =   3132
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   45
               X1              =   960
               X2              =   2745
               Y1              =   2940
               Y2              =   2940
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��Ժ���"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   45
               Left            =   240
               TabIndex        =   165
               Top             =   2760
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   47
               X1              =   7080
               X2              =   9315
               Y1              =   2940
               Y2              =   2940
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�����"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   47
               Left            =   6540
               TabIndex        =   152
               Top             =   2760
               Width           =   540
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����ԭ��"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   62
               Left            =   3480
               TabIndex        =   144
               Top             =   5010
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   62
               X1              =   4245
               X2              =   9120
               Y1              =   5190
               Y2              =   5190
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����ԭ��"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   58
               Left            =   2960
               TabIndex        =   141
               Top             =   4248
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   58
               X1              =   3720
               X2              =   6840
               Y1              =   4420
               Y2              =   4420
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�ɹ�����"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   61
               Left            =   2055
               TabIndex        =   123
               Top             =   5010
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "���ȴ���"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   60
               Left            =   240
               TabIndex        =   122
               Top             =   5010
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ȷ������"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   46
               Left            =   3855
               TabIndex        =   121
               Top             =   2760
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   46
               X1              =   4605
               X2              =   6390
               Y1              =   2940
               Y2              =   2940
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   60
               X1              =   960
               X2              =   1845
               Y1              =   5190
               Y2              =   5190
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   61
               X1              =   2820
               X2              =   3420
               Y1              =   5190
               Y2              =   5190
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   50
               X1              =   7080
               X2              =   9240
               Y1              =   3310
               Y2              =   3310
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   55
               X1              =   4080
               X2              =   5640
               Y1              =   4050
               Y2              =   4050
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   52
               X1              =   4080
               X2              =   5640
               Y1              =   3690
               Y2              =   3690
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   54
               X1              =   960
               X2              =   2760
               Y1              =   4050
               Y2              =   4050
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   51
               X1              =   960
               X2              =   2760
               Y1              =   3690
               Y2              =   3690
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�������Ժ"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   51
               Left            =   60
               TabIndex        =   120
               Top             =   3504
               Width           =   900
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��Ժ���Ժ"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   52
               Left            =   3140
               TabIndex        =   119
               Top             =   3504
               Width           =   900
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�����벡��"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   50
               Left            =   6180
               TabIndex        =   118
               Top             =   3132
               Width           =   900
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�ٴ��벡��"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   54
               Left            =   60
               TabIndex        =   117
               Top             =   3876
               Width           =   900
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�ٴ���ʬ��"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   55
               Left            =   3140
               TabIndex        =   116
               Top             =   3876
               Width           =   900
            End
         End
         Begin VB.Frame fraInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "   ���������� "
            ForeColor       =   &H00FF0000&
            Height          =   3345
            Index           =   3
            Left            =   120
            TabIndex        =   99
            Tag             =   "3705"
            Top             =   120
            Width           =   9495
            Begin VB.PictureBox picSize 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   165
               Index           =   3
               Left            =   180
               ScaleHeight     =   135
               ScaleWidth      =   135
               TabIndex        =   100
               TabStop         =   0   'False
               Top             =   0
               Width           =   165
            End
            Begin VSFlex8Ctl.VSFlexGrid vsOPS 
               Height          =   1335
               Left            =   165
               TabIndex        =   57
               Top             =   1800
               Width           =   9180
               _cx             =   16192
               _cy             =   2355
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
               BackColor       =   16777215
               ForeColor       =   -2147483640
               BackColorFixed  =   16777215
               ForeColorFixed  =   4210752
               BackColorSel    =   4210752
               ForeColorSel    =   -2147483634
               BackColorBkg    =   16777215
               BackColorAlternate=   16777215
               GridColor       =   -2147483636
               GridColorFixed  =   -2147483636
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483643
               FocusRect       =   3
               HighLight       =   2
               AllowSelection  =   0   'False
               AllowBigSelection=   0   'False
               AllowUserResizing=   1
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   1
               GridLineWidth   =   1
               Rows            =   2
               Cols            =   24
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   250
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"frmArchiveInMedRec_SC.frx":0DD9
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
               AllowUserFreezing=   0
               BackColorFrozen =   0
               ForeColorFrozen =   0
               WallPaperAlignment=   9
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   24
            End
            Begin VSFlex8Ctl.VSFlexGrid vsAller 
               Height          =   1335
               Left            =   165
               TabIndex        =   56
               Top             =   300
               Width           =   9180
               _cx             =   16192
               _cy             =   2355
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
               BackColor       =   16777215
               ForeColor       =   -2147483640
               BackColorFixed  =   16777215
               ForeColorFixed  =   4210752
               BackColorSel    =   4210752
               ForeColorSel    =   -2147483634
               BackColorBkg    =   16777215
               BackColorAlternate=   16777215
               GridColor       =   -2147483636
               GridColorFixed  =   -2147483636
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483643
               FocusRect       =   3
               HighLight       =   2
               AllowSelection  =   0   'False
               AllowBigSelection=   0   'False
               AllowUserResizing=   1
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   1
               GridLineWidth   =   1
               Rows            =   2
               Cols            =   3
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   250
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmArchiveInMedRec_SC.frx":10B5
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
         Begin VB.Frame fraInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "   ��ҳ2"
            ForeColor       =   &H00FF0000&
            Height          =   6090
            Index           =   7
            Left            =   120
            TabIndex        =   188
            Tag             =   "6090"
            Top             =   120
            Width           =   9495
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "סԺ�ڼ�����Լ��"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   13
               Left            =   5520
               TabIndex        =   335
               Top             =   1680
               Width           =   1740
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   130
               Left            =   7935
               Locked          =   -1  'True
               TabIndex        =   333
               Top             =   2040
               Width           =   1455
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   129
               Left            =   6255
               Locked          =   -1  'True
               TabIndex        =   331
               Top             =   1320
               Width           =   1455
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   128
               Left            =   6255
               Locked          =   -1  'True
               TabIndex        =   329
               Top             =   960
               Width           =   1455
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   127
               Left            =   6255
               Locked          =   -1  'True
               TabIndex        =   327
               Top             =   600
               Width           =   1455
            End
            Begin VB.Frame fraSplit 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00FFFFFF&
               Height          =   75
               Index           =   1
               Left            =   1320
               TabIndex        =   198
               Top             =   3240
               Width           =   7815
            End
            Begin VB.Frame fraAdvEvent 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "�����¼�"
               ForeColor       =   &H80000008&
               Height          =   2715
               Left            =   240
               TabIndex        =   192
               Top             =   240
               Width           =   4335
               Begin VB.TextBox txtinfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   180
                  IMEMode         =   3  'DISABLE
                  Index           =   126
                  Left            =   1440
                  Locked          =   -1  'True
                  TabIndex        =   202
                  Top             =   2220
                  Width           =   1995
               End
               Begin VB.TextBox txtinfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   180
                  IMEMode         =   3  'DISABLE
                  Index           =   125
                  Left            =   1440
                  Locked          =   -1  'True
                  TabIndex        =   201
                  Top             =   1860
                  Width           =   1995
               End
               Begin VB.TextBox txtinfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   180
                  IMEMode         =   3  'DISABLE
                  Index           =   124
                  Left            =   2880
                  Locked          =   -1  'True
                  TabIndex        =   200
                  Top             =   1500
                  Width           =   915
               End
               Begin VB.TextBox txtinfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   180
                  IMEMode         =   3  'DISABLE
                  Index           =   123
                  Left            =   1200
                  Locked          =   -1  'True
                  TabIndex        =   199
                  Top             =   1500
                  Width           =   975
               End
               Begin VB.ListBox lstAdvEvent 
                  Height          =   960
                  ItemData        =   "frmArchiveInMedRec_SC.frx":1122
                  Left            =   120
                  List            =   "frmArchiveInMedRec_SC.frx":1124
                  TabIndex        =   193
                  Top             =   240
                  Width           =   3765
               End
               Begin VB.Line linInfo 
                  BorderColor     =   &H00808080&
                  Index           =   126
                  X1              =   1440
                  X2              =   3840
                  Y1              =   2400
                  Y2              =   2400
               End
               Begin VB.Line linInfo 
                  BorderColor     =   &H00808080&
                  Index           =   125
                  X1              =   1440
                  X2              =   3840
                  Y1              =   2040
                  Y2              =   2040
               End
               Begin VB.Line linInfo 
                  BorderColor     =   &H00808080&
                  Index           =   124
                  X1              =   2895
                  X2              =   3845
                  Y1              =   1680
                  Y2              =   1680
               End
               Begin VB.Line linInfo 
                  BorderColor     =   &H00808080&
                  Index           =   123
                  X1              =   1200
                  X2              =   2260
                  Y1              =   1680
                  Y2              =   1680
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "����"
                  Height          =   180
                  Index           =   124
                  Left            =   2520
                  TabIndex        =   197
                  Top             =   1500
                  Width           =   360
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "ѹ�������ڼ�"
                  Height          =   180
                  Index           =   123
                  Left            =   120
                  TabIndex        =   196
                  Top             =   1500
                  Width           =   1080
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "������׹��ԭ��"
                  Height          =   180
                  Index           =   126
                  Left            =   120
                  TabIndex        =   195
                  Top             =   2220
                  Width           =   1260
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "������׹���˺�"
                  Height          =   180
                  Index           =   125
                  Left            =   120
                  TabIndex        =   194
                  Top             =   1860
                  Width           =   1260
               End
            End
            Begin VB.PictureBox picSize 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   165
               Index           =   7
               Left            =   180
               ScaleHeight     =   135
               ScaleWidth      =   135
               TabIndex        =   189
               TabStop         =   0   'False
               Top             =   0
               Width           =   165
            End
            Begin VSFlex8Ctl.VSFlexGrid vsfMain 
               Height          =   2490
               Left            =   120
               TabIndex        =   191
               Top             =   3480
               Width           =   9165
               _cx             =   16166
               _cy             =   4392
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
               BackColor       =   16777215
               ForeColor       =   -2147483640
               BackColorFixed  =   16777215
               ForeColorFixed  =   4210752
               BackColorSel    =   4210752
               ForeColorSel    =   -2147483634
               BackColorBkg    =   16777215
               BackColorAlternate=   16777215
               GridColor       =   -2147483636
               GridColorFixed  =   -2147483636
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483643
               FocusRect       =   3
               HighLight       =   2
               AllowSelection  =   0   'False
               AllowBigSelection=   0   'False
               AllowUserResizing=   1
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   1
               GridLineWidth   =   1
               Rows            =   50
               Cols            =   10
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   250
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   ""
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
               AllowUserFreezing=   0
               BackColorFrozen =   0
               ForeColorFrozen =   0
               WallPaperAlignment=   9
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   24
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   130
               X1              =   7935
               X2              =   9420
               Y1              =   2220
               Y2              =   2220
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��Ժʱ͸����Ѫ͸����͸�����ص�ֵ"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   130
               Left            =   5040
               TabIndex        =   334
               Top             =   2040
               Width           =   2880
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   129
               X1              =   6255
               X2              =   7740
               Y1              =   1500
               Y2              =   1500
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�ٴ�����"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   129
               Left            =   5520
               TabIndex        =   332
               Top             =   1320
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   128
               X1              =   6255
               X2              =   7740
               Y1              =   1140
               Y2              =   1140
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����ҩ��"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   128
               Left            =   5520
               TabIndex        =   330
               Top             =   960
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   127
               X1              =   6255
               X2              =   7740
               Y1              =   780
               Y2              =   780
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��Һ��Ӧ"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   127
               Left            =   5520
               TabIndex        =   328
               Top             =   600
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����������Ŀ"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   140
               Left            =   240
               TabIndex        =   190
               Top             =   3180
               Width           =   1080
            End
         End
         Begin VB.Frame fraInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "   ��ҳ1 "
            ForeColor       =   &H00FF0000&
            Height          =   7170
            Index           =   6
            Left            =   120
            TabIndex        =   181
            Tag             =   "7170"
            Top             =   120
            Width           =   9495
            Begin VB.PictureBox picSize 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   165
               Index           =   6
               Left            =   180
               ScaleHeight     =   135
               ScaleWidth      =   135
               TabIndex        =   182
               TabStop         =   0   'False
               Top             =   0
               Width           =   165
            End
            Begin VSFlex8Ctl.VSFlexGrid vsKSS 
               Height          =   1095
               Left            =   120
               TabIndex        =   183
               Top             =   480
               Width           =   9120
               _cx             =   16087
               _cy             =   1931
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
               BackColor       =   16777215
               ForeColor       =   -2147483640
               BackColorFixed  =   16777215
               ForeColorFixed  =   4210752
               BackColorSel    =   4210752
               ForeColorSel    =   -2147483634
               BackColorBkg    =   16777215
               BackColorAlternate=   16777215
               GridColor       =   -2147483636
               GridColorFixed  =   -2147483636
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483643
               FocusRect       =   3
               HighLight       =   2
               AllowSelection  =   0   'False
               AllowBigSelection=   0   'False
               AllowUserResizing=   1
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   1
               GridLineWidth   =   1
               Rows            =   2
               Cols            =   7
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   250
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmArchiveInMedRec_SC.frx":1126
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
               AllowUserFreezing=   0
               BackColorFrozen =   0
               ForeColorFrozen =   0
               WallPaperAlignment=   9
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   24
            End
            Begin VSFlex8Ctl.VSFlexGrid vsFlxAddICU 
               Height          =   1305
               Left            =   120
               TabIndex        =   211
               Top             =   2040
               Width           =   9120
               _cx             =   16087
               _cy             =   2302
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
               BackColor       =   16777215
               ForeColor       =   -2147483640
               BackColorFixed  =   16777215
               ForeColorFixed  =   4210752
               BackColorSel    =   4210752
               ForeColorSel    =   -2147483634
               BackColorBkg    =   16777215
               BackColorAlternate=   16777215
               GridColor       =   -2147483636
               GridColorFixed  =   -2147483636
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483643
               FocusRect       =   3
               HighLight       =   2
               AllowSelection  =   0   'False
               AllowBigSelection=   0   'False
               AllowUserResizing=   1
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   1
               GridLineWidth   =   1
               Rows            =   5
               Cols            =   6
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   250
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmArchiveInMedRec_SC.frx":120D
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
               AllowUserFreezing=   0
               BackColorFrozen =   0
               ForeColorFrozen =   0
               WallPaperAlignment=   9
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   24
            End
            Begin VSFlex8Ctl.VSFlexGrid vsICUInstruments 
               Height          =   1305
               Left            =   120
               TabIndex        =   213
               Top             =   3840
               Width           =   9120
               _cx             =   16087
               _cy             =   2302
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
               BackColor       =   16777215
               ForeColor       =   -2147483640
               BackColorFixed  =   16777215
               ForeColorFixed  =   4210752
               BackColorSel    =   4210752
               ForeColorSel    =   -2147483634
               BackColorBkg    =   16777215
               BackColorAlternate=   16777215
               GridColor       =   -2147483636
               GridColorFixed  =   -2147483636
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483643
               FocusRect       =   3
               HighLight       =   2
               AllowSelection  =   0   'False
               AllowBigSelection=   0   'False
               AllowUserResizing=   1
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   1
               GridLineWidth   =   1
               Rows            =   5
               Cols            =   5
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   250
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmArchiveInMedRec_SC.frx":12D7
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
               AllowUserFreezing=   0
               BackColorFrozen =   0
               ForeColorFrozen =   0
               WallPaperAlignment=   9
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   24
            End
            Begin VSFlex8Ctl.VSFlexGrid vsInfect 
               Height          =   1305
               Left            =   120
               TabIndex        =   215
               Top             =   5520
               Width           =   3360
               _cx             =   5927
               _cy             =   2302
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
               BackColor       =   16777215
               ForeColor       =   -2147483640
               BackColorFixed  =   16777215
               ForeColorFixed  =   4210752
               BackColorSel    =   4210752
               ForeColorSel    =   -2147483634
               BackColorBkg    =   16777215
               BackColorAlternate=   16777215
               GridColor       =   -2147483636
               GridColorFixed  =   -2147483636
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483643
               FocusRect       =   3
               HighLight       =   2
               AllowSelection  =   0   'False
               AllowBigSelection=   0   'False
               AllowUserResizing=   1
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   1
               GridLineWidth   =   1
               Rows            =   5
               Cols            =   3
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   250
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmArchiveInMedRec_SC.frx":139A
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
               AllowUserFreezing=   0
               BackColorFrozen =   0
               ForeColorFrozen =   0
               WallPaperAlignment=   9
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   24
            End
            Begin VSFlex8Ctl.VSFlexGrid vsSample 
               Height          =   1305
               Left            =   4200
               TabIndex        =   217
               Top             =   5520
               Width           =   5040
               _cx             =   8890
               _cy             =   2302
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
               BackColor       =   16777215
               ForeColor       =   -2147483640
               BackColorFixed  =   16777215
               ForeColorFixed  =   4210752
               BackColorSel    =   4210752
               ForeColorSel    =   -2147483634
               BackColorBkg    =   16777215
               BackColorAlternate=   16777215
               GridColor       =   -2147483636
               GridColorFixed  =   -2147483636
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483643
               FocusRect       =   3
               HighLight       =   2
               AllowSelection  =   0   'False
               AllowBigSelection=   0   'False
               AllowUserResizing=   1
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   1
               GridLineWidth   =   1
               Rows            =   5
               Cols            =   3
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   250
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmArchiveInMedRec_SC.frx":140A
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
               AllowUserFreezing=   0
               BackColorFrozen =   0
               ForeColorFrozen =   0
               WallPaperAlignment=   9
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   24
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�걾��Դ"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   136
               Left            =   4200
               TabIndex        =   216
               Top             =   5280
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ҽԺ��Ⱦ���"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   135
               Left            =   120
               TabIndex        =   214
               Top             =   5280
               Width           =   1080
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "������ס��֢�໤���ڼ���еʹ�����"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   134
               Left            =   120
               TabIndex        =   212
               Top             =   3600
               Width           =   3060
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����ҩ��ʹ���������DDD���������У�"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   142
               Left            =   120
               TabIndex        =   185
               Top             =   240
               Width           =   3150
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "������ס��֢�໤������¼"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   141
               Left            =   120
               TabIndex        =   184
               Top             =   1800
               Width           =   2160
            End
         End
         Begin VB.Frame fraInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "   ��ҽ��� "
            ForeColor       =   &H00FF0000&
            Height          =   4170
            Index           =   2
            Left            =   120
            TabIndex        =   101
            Tag             =   "4170"
            Top             =   120
            Width           =   9495
            Begin VB.Frame fraSub 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   " ���Ʒ��� "
               ForeColor       =   &H00404040&
               Height          =   1320
               Index           =   2
               Left            =   4320
               TabIndex        =   108
               Top             =   2580
               Width           =   4905
               Begin VB.TextBox txtinfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   180
                  Index           =   73
                  Left            =   3990
                  Locked          =   -1  'True
                  TabIndex        =   149
                  Top             =   960
                  Width           =   555
               End
               Begin VB.TextBox txtinfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   180
                  Index           =   72
                  Left            =   3990
                  Locked          =   -1  'True
                  TabIndex        =   147
                  Top             =   645
                  Width           =   555
               End
               Begin VB.TextBox txtinfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   180
                  Index           =   71
                  Left            =   3990
                  Locked          =   -1  'True
                  TabIndex        =   145
                  Top             =   330
                  Width           =   555
               End
               Begin VB.TextBox txtinfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   180
                  Index           =   68
                  Left            =   1185
                  Locked          =   -1  'True
                  TabIndex        =   53
                  Top             =   330
                  Width           =   1035
               End
               Begin VB.TextBox txtinfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   180
                  Index           =   69
                  Left            =   1185
                  Locked          =   -1  'True
                  TabIndex        =   54
                  Top             =   645
                  Width           =   1035
               End
               Begin VB.TextBox txtinfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   180
                  Index           =   70
                  Left            =   1545
                  Locked          =   -1  'True
                  TabIndex        =   55
                  Top             =   960
                  Width           =   1035
               End
               Begin VB.Line linInfo 
                  BorderColor     =   &H00808080&
                  Index           =   73
                  X1              =   3960
                  X2              =   4545
                  Y1              =   1140
                  Y2              =   1140
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "��֤ʩ��"
                  ForeColor       =   &H00404040&
                  Height          =   180
                  Index           =   73
                  Left            =   3240
                  TabIndex        =   150
                  Top             =   960
                  Width           =   720
               End
               Begin VB.Line linInfo 
                  BorderColor     =   &H00808080&
                  Index           =   72
                  X1              =   3960
                  X2              =   4545
                  Y1              =   825
                  Y2              =   825
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "ʹ����ҽ���Ƽ���"
                  ForeColor       =   &H00404040&
                  Height          =   180
                  Index           =   72
                  Left            =   2520
                  TabIndex        =   148
                  Top             =   645
                  Width           =   1440
               End
               Begin VB.Line linInfo 
                  BorderColor     =   &H00808080&
                  Index           =   71
                  X1              =   3960
                  X2              =   4545
                  Y1              =   510
                  Y2              =   510
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "ʹ����ҽ�����豸"
                  ForeColor       =   &H00404040&
                  Height          =   180
                  Index           =   71
                  Left            =   2520
                  TabIndex        =   146
                  Top             =   330
                  Width           =   1440
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "�������"
                  ForeColor       =   &H00404040&
                  Height          =   180
                  Index           =   68
                  Left            =   315
                  TabIndex        =   111
                  Top             =   330
                  Width           =   720
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "���ȷ���"
                  ForeColor       =   &H00404040&
                  Height          =   180
                  Index           =   69
                  Left            =   315
                  TabIndex        =   110
                  Top             =   645
                  Width           =   720
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "������ҩ�Ƽ�"
                  ForeColor       =   &H00404040&
                  Height          =   180
                  Index           =   70
                  Left            =   315
                  TabIndex        =   109
                  Top             =   960
                  Width           =   1080
               End
               Begin VB.Line linInfo 
                  BorderColor     =   &H00808080&
                  Index           =   68
                  X1              =   1095
                  X2              =   2220
                  Y1              =   510
                  Y2              =   510
               End
               Begin VB.Line linInfo 
                  BorderColor     =   &H00808080&
                  Index           =   69
                  X1              =   1095
                  X2              =   2220
                  Y1              =   825
                  Y2              =   825
               End
               Begin VB.Line linInfo 
                  BorderColor     =   &H00808080&
                  Index           =   70
                  X1              =   1455
                  X2              =   2580
                  Y1              =   1140
                  Y2              =   1140
               End
            End
            Begin VB.Frame fraSub 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   " סԺ�ڼ䲡�� "
               ForeColor       =   &H00404040&
               Height          =   1320
               Index           =   0
               Left            =   165
               TabIndex        =   107
               Top             =   2580
               Width           =   1485
               Begin VB.CheckBox chkInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  Caption         =   "����"
                  ForeColor       =   &H00404040&
                  Height          =   195
                  Index           =   8
                  Left            =   405
                  TabIndex        =   49
                  Top             =   960
                  Width           =   660
               End
               Begin VB.CheckBox chkInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  Caption         =   "��֢"
                  ForeColor       =   &H00404040&
                  Height          =   195
                  Index           =   7
                  Left            =   405
                  TabIndex        =   48
                  Top             =   645
                  Width           =   660
               End
               Begin VB.CheckBox chkInfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  Caption         =   "Σ��"
                  ForeColor       =   &H00404040&
                  Height          =   195
                  Index           =   6
                  Left            =   405
                  TabIndex        =   47
                  Top             =   330
                  Width           =   660
               End
            End
            Begin VB.PictureBox picSize 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   165
               Index           =   2
               Left            =   180
               ScaleHeight     =   135
               ScaleWidth      =   135
               TabIndex        =   106
               TabStop         =   0   'False
               Top             =   0
               Width           =   165
            End
            Begin VB.Frame fraSub 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   " ׼ȷ�� "
               ForeColor       =   &H00404040&
               Height          =   1320
               Index           =   1
               Left            =   2032
               TabIndex        =   102
               Top             =   2580
               Width           =   1905
               Begin VB.TextBox txtinfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   180
                  Index           =   67
                  Left            =   600
                  Locked          =   -1  'True
                  TabIndex        =   52
                  Top             =   960
                  Width           =   1035
               End
               Begin VB.TextBox txtinfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   180
                  Index           =   66
                  Left            =   600
                  Locked          =   -1  'True
                  TabIndex        =   51
                  Top             =   645
                  Width           =   1035
               End
               Begin VB.TextBox txtinfo 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   180
                  Index           =   65
                  Left            =   600
                  Locked          =   -1  'True
                  TabIndex        =   50
                  Top             =   330
                  Width           =   1035
               End
               Begin VB.Line linInfo 
                  BorderColor     =   &H00808080&
                  Index           =   67
                  X1              =   630
                  X2              =   1755
                  Y1              =   1140
                  Y2              =   1140
               End
               Begin VB.Line linInfo 
                  BorderColor     =   &H00808080&
                  Index           =   66
                  X1              =   630
                  X2              =   1755
                  Y1              =   825
                  Y2              =   825
               End
               Begin VB.Line linInfo 
                  BorderColor     =   &H00808080&
                  Index           =   65
                  X1              =   630
                  X2              =   1755
                  Y1              =   510
                  Y2              =   510
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "��ҩ"
                  ForeColor       =   &H00404040&
                  Height          =   180
                  Index           =   67
                  Left            =   210
                  TabIndex        =   105
                  Top             =   960
                  Width           =   360
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "�η�"
                  ForeColor       =   &H00404040&
                  Height          =   180
                  Index           =   66
                  Left            =   210
                  TabIndex        =   104
                  Top             =   645
                  Width           =   360
               End
               Begin VB.Label lblInfo 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "��֤"
                  ForeColor       =   &H00404040&
                  Height          =   180
                  Index           =   65
                  Left            =   210
                  TabIndex        =   103
                  Top             =   330
                  Width           =   360
               End
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   63
               Left            =   1425
               Locked          =   -1  'True
               TabIndex        =   45
               Top             =   2190
               Width           =   915
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   64
               Left            =   4020
               Locked          =   -1  'True
               TabIndex        =   46
               Top             =   2190
               Width           =   915
            End
            Begin VSFlex8Ctl.VSFlexGrid vsDiagZY 
               Height          =   1710
               Left            =   165
               TabIndex        =   44
               Top             =   270
               Width           =   9180
               _cx             =   16192
               _cy             =   3016
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
               BackColor       =   16777215
               ForeColor       =   -2147483640
               BackColorFixed  =   16777215
               ForeColorFixed  =   4210752
               BackColorSel    =   4210752
               ForeColorSel    =   -2147483634
               BackColorBkg    =   16777215
               BackColorAlternate=   16777215
               GridColor       =   -2147483636
               GridColorFixed  =   -2147483636
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483643
               FocusRect       =   3
               HighLight       =   2
               AllowSelection  =   0   'False
               AllowBigSelection=   0   'False
               AllowUserResizing=   1
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   1
               GridLineWidth   =   1
               Rows            =   5
               Cols            =   7
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   250
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"frmArchiveInMedRec_SC.frx":147B
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
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��Ժ���Ժ"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   64
               Left            =   3000
               TabIndex        =   113
               Top             =   2190
               Width           =   900
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�������Ժ"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   63
               Left            =   390
               TabIndex        =   112
               Top             =   2190
               Width           =   900
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   63
               X1              =   1335
               X2              =   2465
               Y1              =   2370
               Y2              =   2370
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   64
               X1              =   3930
               X2              =   5015
               Y1              =   2370
               Y2              =   2370
            End
         End
         Begin VB.Frame fraInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "   ������Ϣ "
            ForeColor       =   &H00FF0000&
            Height          =   6195
            Index           =   0
            Left            =   120
            TabIndex        =   63
            Tag             =   "6195"
            Top             =   120
            Width           =   9495
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   132
               Left            =   6960
               Locked          =   -1  'True
               TabIndex        =   207
               Top             =   4305
               Width           =   1860
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   133
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   205
               Top             =   4305
               Width           =   1860
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   131
               Left            =   4815
               Locked          =   -1  'True
               TabIndex        =   203
               Top             =   2160
               Width           =   2805
            End
            Begin VB.TextBox txtinfo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   10
               Left            =   2640
               Locked          =   -1  'True
               TabIndex        =   164
               Top             =   1065
               Width           =   425
            End
            Begin VB.TextBox txtinfo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   9
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   163
               Top             =   1065
               Width           =   425
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   32
               Left            =   5760
               Locked          =   -1  'True
               TabIndex        =   160
               Top             =   3945
               Width           =   3075
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   37
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   158
               Top             =   5385
               Width           =   1455
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "����Ժ"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   0
               Left            =   5640
               TabIndex        =   156
               Top             =   338
               Width           =   915
            End
            Begin VB.CheckBox chkInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "��Ժǰ����Ժ����"
               ForeColor       =   &H00404040&
               Height          =   195
               Index           =   1
               Left            =   6780
               TabIndex        =   142
               Top             =   5018
               Width           =   1740
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   23
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   135
               Top             =   2865
               Width           =   2805
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   24
               Left            =   4815
               Locked          =   -1  'True
               TabIndex        =   134
               Top             =   2865
               Width           =   1530
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   15
               Left            =   4815
               Locked          =   -1  'True
               TabIndex        =   132
               Top             =   1785
               Width           =   1530
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   13
               Left            =   6960
               Locked          =   -1  'True
               TabIndex        =   129
               Top             =   1425
               Width           =   1740
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   12
               Left            =   4455
               Locked          =   -1  'True
               TabIndex        =   128
               Top             =   1425
               Width           =   1050
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   11
               Left            =   2400
               Locked          =   -1  'True
               TabIndex        =   126
               Top             =   1425
               Width           =   810
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   3
               Left            =   7260
               Locked          =   -1  'True
               TabIndex        =   124
               Top             =   345
               Width           =   1395
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   0
               Left            =   1305
               Locked          =   -1  'True
               TabIndex        =   1
               Top             =   345
               Width           =   900
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   40
               Left            =   7560
               Locked          =   -1  'True
               TabIndex        =   29
               Top             =   5385
               Width           =   1290
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   39
               Left            =   5445
               Locked          =   -1  'True
               TabIndex        =   28
               Top             =   5385
               Width           =   1290
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   38
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   27
               Top             =   5385
               Width           =   1290
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   44
               Left            =   7200
               Locked          =   -1  'True
               TabIndex        =   33
               Top             =   5745
               Width           =   975
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   43
               Left            =   5160
               Locked          =   -1  'True
               TabIndex        =   32
               Top             =   5745
               Width           =   945
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   42
               Left            =   3270
               Locked          =   -1  'True
               TabIndex        =   31
               Top             =   5745
               Width           =   1335
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   41
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   30
               Top             =   5745
               Width           =   1455
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   36
               Left            =   5040
               Locked          =   -1  'True
               TabIndex        =   26
               Top             =   5025
               Width           =   945
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   35
               Left            =   3150
               Locked          =   -1  'True
               TabIndex        =   25
               Top             =   5025
               Width           =   1335
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   34
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   24
               Top             =   5025
               Width           =   1335
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   31
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   23
               Top             =   3945
               Width           =   4200
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   30
               Left            =   4815
               Locked          =   -1  'True
               TabIndex        =   22
               Top             =   3585
               Width           =   1530
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   28
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   20
               Top             =   3585
               Width           =   1035
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   27
               Left            =   6960
               Locked          =   -1  'True
               TabIndex        =   19
               Top             =   3225
               Width           =   1695
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   26
               Left            =   4815
               Locked          =   -1  'True
               TabIndex        =   18
               Top             =   3225
               Width           =   1530
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   25
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   17
               Top             =   3225
               Width           =   2805
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   22
               Left            =   6960
               Locked          =   -1  'True
               TabIndex        =   16
               Top             =   2505
               Width           =   1815
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   21
               Left            =   4815
               Locked          =   -1  'True
               TabIndex        =   15
               Top             =   2505
               Width           =   1530
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   20
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   14
               Top             =   2505
               Width           =   2805
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   17
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   13
               Top             =   2145
               Width           =   2805
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   14
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   12
               Top             =   1785
               Width           =   2805
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   7
               Left            =   6330
               Locked          =   -1  'True
               TabIndex        =   6
               Top             =   705
               Width           =   690
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   6
               Left            =   4545
               Locked          =   -1  'True
               TabIndex        =   5
               Top             =   705
               Width           =   1260
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               Index           =   4
               Left            =   1180
               Locked          =   -1  'True
               TabIndex        =   3
               Top             =   705
               Width           =   860
            End
            Begin VB.TextBox txtinfo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   2
               Left            =   4635
               Locked          =   -1  'True
               TabIndex        =   2
               Top             =   345
               Width           =   285
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   1
               Left            =   3210
               Locked          =   -1  'True
               TabIndex        =   0
               Top             =   345
               Width           =   1050
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   5
               Left            =   2760
               Locked          =   -1  'True
               TabIndex        =   4
               Top             =   705
               Width           =   645
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   19
               Left            =   6240
               Locked          =   -1  'True
               TabIndex        =   7
               Top             =   1065
               Width           =   975
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   16
               Left            =   7560
               Locked          =   -1  'True
               TabIndex        =   10
               Top             =   1065
               Width           =   1215
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   8
               Left            =   7650
               Locked          =   -1  'True
               TabIndex        =   9
               Top             =   705
               Width           =   1095
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   18
               Left            =   4095
               Locked          =   -1  'True
               TabIndex        =   8
               Top             =   1065
               Width           =   1530
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   33
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   11
               Top             =   4665
               Width           =   1500
            End
            Begin VB.TextBox txtinfo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   180
               IMEMode         =   3  'DISABLE
               Index           =   29
               Left            =   2880
               Locked          =   -1  'True
               TabIndex        =   21
               Top             =   3585
               Width           =   1095
            End
            Begin VB.PictureBox picSize 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   165
               Index           =   0
               Left            =   180
               ScaleHeight     =   135
               ScaleWidth      =   135
               TabIndex        =   64
               TabStop         =   0   'False
               Top             =   0
               Width           =   165
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   132
               X1              =   6960
               X2              =   8880
               Y1              =   4485
               Y2              =   4485
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "QQ"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   132
               Left            =   6780
               TabIndex        =   208
               Top             =   4305
               Width           =   180
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   133
               X1              =   1080
               X2              =   3000
               Y1              =   4485
               Y2              =   4485
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Email"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   133
               Left            =   600
               TabIndex        =   206
               Top             =   4305
               Width           =   450
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   131
               X1              =   4815
               X2              =   7695
               Y1              =   2340
               Y2              =   2340
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����֤��"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   131
               Left            =   4080
               TabIndex        =   204
               Top             =   2160
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   10
               X1              =   2670
               X2              =   3120
               Y1              =   1240
               Y2              =   1240
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   9
               X1              =   1060
               X2              =   1580
               Y1              =   1240
               Y2              =   1240
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   32
               X1              =   5760
               X2              =   8880
               Y1              =   4125
               Y2              =   4125
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   37
               X1              =   1080
               X2              =   2640
               Y1              =   5560
               Y2              =   5560
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "���      cm"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   10
               Left            =   2265
               TabIndex        =   162
               Top             =   1065
               Width           =   1080
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����      kg"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   9
               Left            =   720
               TabIndex        =   161
               Top             =   1065
               Width           =   1080
            End
            Begin VB.Label lblInfo 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "����"
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   32
               Left            =   5400
               TabIndex        =   159
               Top             =   3945
               Width           =   360
            End
            Begin VB.Label lblInfo 
               BackStyle       =   0  'Transparent
               Caption         =   "���ʱ��"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   37
               Left            =   330
               TabIndex        =   157
               Top             =   5385
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   24
               X1              =   4845
               X2              =   6380
               Y1              =   3040
               Y2              =   3040
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   23
               X1              =   1080
               X2              =   3960
               Y1              =   3040
               Y2              =   3040
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "���ڵ�ַ"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   23
               Left            =   330
               TabIndex        =   137
               Top             =   2865
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�ʱ�"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   24
               Left            =   4440
               TabIndex        =   136
               Top             =   2865
               Width           =   360
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   15
               X1              =   4845
               X2              =   6375
               Y1              =   1960
               Y2              =   1960
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   16
               X1              =   7560
               X2              =   8760
               Y1              =   1245
               Y2              =   1245
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   15
               Left            =   4440
               TabIndex        =   133
               Top             =   1785
               Width           =   360
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   13
               X1              =   6960
               X2              =   8760
               Y1              =   1605
               Y2              =   1605
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   12
               X1              =   4365
               X2              =   5520
               Y1              =   1600
               Y2              =   1600
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����������"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   12
               Left            =   3480
               TabIndex        =   131
               Top             =   1425
               Width           =   900
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��������Ժ����"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   13
               Left            =   5700
               TabIndex        =   130
               Top             =   1425
               Width           =   1260
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   11
               X1              =   2280
               X2              =   3360
               Y1              =   1600
               Y2              =   1600
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�����䲻��һ����ģ� ����"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   11
               Left            =   90
               TabIndex        =   127
               Top             =   1425
               Width           =   2250
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   3
               X1              =   7170
               X2              =   8760
               Y1              =   525
               Y2              =   525
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "������"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   3
               Left            =   6600
               TabIndex        =   125
               Top             =   345
               Width           =   540
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "סԺ����"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   44
               Left            =   6480
               TabIndex        =   98
               Top             =   5745
               Width           =   720
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��������"
               ForeColor       =   &H00808080&
               Height          =   180
               Index           =   40
               Left            =   6780
               TabIndex        =   97
               Top             =   5385
               Width           =   720
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��������"
               ForeColor       =   &H00808080&
               Height          =   180
               Index           =   39
               Left            =   4680
               TabIndex        =   96
               Top             =   5385
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ת�����"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   38
               Left            =   2640
               TabIndex        =   95
               Top             =   5385
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   43
               Left            =   4680
               TabIndex        =   94
               Top             =   5745
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   42
               Left            =   2805
               TabIndex        =   93
               Top             =   5745
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��Ժʱ��"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   41
               Left            =   330
               TabIndex        =   92
               Top             =   5745
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   36
               Left            =   4560
               TabIndex        =   91
               Top             =   5025
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   35
               Left            =   2685
               TabIndex        =   90
               Top             =   5025
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��Ժʱ��"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   34
               Left            =   330
               TabIndex        =   89
               Top             =   5025
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��ϵ�˵�ַ"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   31
               Left            =   150
               TabIndex        =   88
               Top             =   3945
               Width           =   900
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�绰"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   30
               Left            =   4440
               TabIndex        =   87
               Top             =   3585
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��ϵ"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   29
               Left            =   2400
               TabIndex        =   86
               Top             =   3585
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��ϵ������"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   28
               Left            =   150
               TabIndex        =   85
               Top             =   3585
               Width           =   900
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�ʱ�"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   27
               Left            =   6600
               TabIndex        =   84
               Top             =   3225
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�绰"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   26
               Left            =   4440
               TabIndex        =   83
               Top             =   3225
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "������λ"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   25
               Left            =   330
               TabIndex        =   82
               Top             =   3225
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�ʱ�"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   22
               Left            =   6600
               TabIndex        =   81
               Top             =   2505
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�绰"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   21
               Left            =   4440
               TabIndex        =   80
               Top             =   2505
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��סַ"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   20
               Left            =   510
               TabIndex        =   79
               Top             =   2505
               Width           =   540
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "���֤��"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   17
               Left            =   330
               TabIndex        =   78
               Top             =   2145
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�����ص�"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   14
               Left            =   330
               TabIndex        =   77
               Top             =   1785
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   16
               Left            =   7200
               TabIndex        =   76
               Top             =   1065
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   8
               Left            =   7170
               TabIndex        =   75
               Top             =   690
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��Ժ;��"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   33
               Left            =   330
               TabIndex        =   74
               Top             =   4665
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ְҵ"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   18
               Left            =   3720
               TabIndex        =   73
               Top             =   1065
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   19
               Left            =   5880
               TabIndex        =   72
               Top             =   1065
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   7
               Left            =   5940
               TabIndex        =   71
               Top             =   690
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��������"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   6
               Left            =   3690
               TabIndex        =   70
               Top             =   690
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�Ա�"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   5
               Left            =   2265
               TabIndex        =   69
               Top             =   690
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   4
               Left            =   690
               TabIndex        =   68
               Top             =   705
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ҽ�Ƹ��ѷ�ʽ"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   0
               Left            =   90
               TabIndex        =   67
               Top             =   345
               Width           =   1080
            End
            Begin VB.Label lblInfo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��    ��סԺ"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   2
               Left            =   4425
               TabIndex        =   66
               Top             =   345
               Width           =   1080
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��������"
               ForeColor       =   &H00404040&
               Height          =   180
               Index           =   1
               Left            =   2370
               TabIndex        =   65
               Top             =   345
               Width           =   720
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   1
               X1              =   3120
               X2              =   4320
               Y1              =   525
               Y2              =   525
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   2
               X1              =   4635
               X2              =   4925
               Y1              =   525
               Y2              =   525
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   0
               X1              =   1215
               X2              =   2280
               Y1              =   525
               Y2              =   525
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   4
               X1              =   1080
               X2              =   2040
               Y1              =   880
               Y2              =   880
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   7
               X1              =   6330
               X2              =   7080
               Y1              =   880
               Y2              =   880
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   8
               X1              =   7560
               X2              =   8760
               Y1              =   885
               Y2              =   885
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   5
               X1              =   2670
               X2              =   3480
               Y1              =   880
               Y2              =   880
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   19
               X1              =   6240
               X2              =   7200
               Y1              =   1245
               Y2              =   1245
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   6
               X1              =   4455
               X2              =   5760
               Y1              =   880
               Y2              =   880
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   18
               X1              =   4125
               X2              =   5655
               Y1              =   1240
               Y2              =   1240
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   33
               X1              =   1080
               X2              =   2670
               Y1              =   4840
               Y2              =   4840
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   14
               X1              =   1080
               X2              =   3975
               Y1              =   1960
               Y2              =   1960
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   17
               X1              =   1080
               X2              =   3960
               Y1              =   2320
               Y2              =   2320
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   20
               X1              =   1080
               X2              =   3975
               Y1              =   2680
               Y2              =   2680
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   25
               X1              =   1080
               X2              =   3975
               Y1              =   3400
               Y2              =   3400
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   31
               X1              =   1080
               X2              =   5280
               Y1              =   4120
               Y2              =   4120
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   21
               X1              =   4845
               X2              =   6380
               Y1              =   2680
               Y2              =   2680
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   26
               X1              =   4845
               X2              =   6360
               Y1              =   3400
               Y2              =   3400
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   30
               X1              =   4845
               X2              =   6360
               Y1              =   3760
               Y2              =   3760
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   22
               X1              =   6960
               X2              =   8760
               Y1              =   2685
               Y2              =   2685
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   27
               X1              =   6960
               X2              =   8760
               Y1              =   3400
               Y2              =   3400
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   28
               X1              =   1080
               X2              =   2205
               Y1              =   3760
               Y2              =   3760
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   29
               X1              =   2790
               X2              =   3975
               Y1              =   3760
               Y2              =   3760
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   34
               X1              =   1080
               X2              =   2640
               Y1              =   5205
               Y2              =   5205
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   41
               X1              =   1080
               X2              =   2700
               Y1              =   5920
               Y2              =   5920
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   35
               X1              =   3075
               X2              =   4560
               Y1              =   5205
               Y2              =   5205
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   42
               X1              =   3195
               X2              =   4560
               Y1              =   5920
               Y2              =   5920
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   36
               X1              =   4965
               X2              =   5965
               Y1              =   5205
               Y2              =   5205
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   43
               X1              =   5160
               X2              =   6190
               Y1              =   5920
               Y2              =   5920
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   44
               X1              =   7200
               X2              =   8880
               Y1              =   5925
               Y2              =   5925
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   38
               X1              =   3360
               X2              =   4680
               Y1              =   5565
               Y2              =   5565
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   39
               X1              =   5400
               X2              =   6720
               Y1              =   5560
               Y2              =   5560
            End
            Begin VB.Line linInfo 
               BorderColor     =   &H00808080&
               Index           =   40
               X1              =   7560
               X2              =   8880
               Y1              =   5560
               Y2              =   5560
            End
         End
      End
   End
End
Attribute VB_Name = "frmArchiveInMedRec_SC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'˵����Ϊ�˱��ֽ���Ŀ�ά���ԣ��������ؼ�ʱ��ע�Ᵽ��ÿ����Ϣ��Ŀ������lblInfo��linInfo,txtinfo ��index��ͬ��
'      ��������Ϣ��Ŀ����2��lblinfo������һ��lblinfo��indexΪtxtinfo.index+100

'�ϴ�ˢ������ʱ�Ĳ�����Ϣ
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mblnMoved As Boolean
Private mblnCheck As Boolean
Private mbln���� As Boolean

Private Enum Fra�˵�
    FRA_������Ϣ = 0
    FRA_��ҽ��� = 1
    FRA_��ҽ��� = 2
    FRA_���������� = 3
    FRA_סԺ��� = 4
    FRA_�����뻯�� = 5
    FRA_��ҳ1 = 6
    FRA_��ҳ2 = 7
End Enum

Private Enum ������Ϣ
    txt���ʽ = 0
    txt�������� = 1
    txtסԺ���� = 2
    chk����Ժ = 0
    txt������ = 3
    txt���� = 4
    txt�Ա� = 5
    txt�������� = 6
    txt���� = 7
    txt���� = 8
    txt���� = 9
    txt��� = 10
    txt������������ = 11
    txt���������� = 12
    txt��������Ժ���� = 13
    txt�����ص� = 14
    txt���� = 15
    txt���� = 16
    txt���֤�� = 17
    txtְҵ = 18
    txt���� = 19
    txt��ͥ��ַ = 20
    txt��ͥ�绰 = 21
    txt��ͥ�ʱ� = 22
    txt���ڵ�ַ = 23
    txt�����ʱ� = 24
    txt������λ = 25
    txt��λ�绰 = 26
    txt��λ�ʱ� = 27
    txt��ϵ������ = 28
    txt��ϵ�˹�ϵ = 29
    txt��ϵ�˵绰 = 30
    txt��ϵ�˵�ַ = 31
    txt���� = 32
    txt��Ժ;�� = 33
    txt��Ժʱ�� = 34
    txt��Ժ���� = 35
    txt��Ժ���� = 36
    chk��Ժǰ����Ժ���� = 1
    txt���ʱ�� = 37
    txtת��1 = 38
    txtת��2 = 39
    txtת��3 = 40
    txt��Ժʱ�� = 41
    txt��Ժ���� = 42
    txt��Ժ���� = 43
    txtסԺ���� = 44
    txt����֤�� = 131
    txtEmail = 133
    txtqq = 132
End Enum

Private Enum ��ҽ���
    txt��Ժ��� = 45
    chk�Ƿ�ȷ�� = 2
    txtȷ������ = 46
    txt����� = 47
    txt�ֻ��̶� = 48
    txt���������� = 49
    txt�����벡�� = 50
    txt�������Ժ = 51
    txt��Ժ���Ժ = 52
    txt��������Ժ = 53
    txt�ٴ��벡�� = 54
    txt�ٴ���ʬ�� = 55
    txt��ǰ������ = 56
    txt����ʱ�� = 57
    txt����ԭ�� = 58
    chkҽԺ��Ⱦ����ԭѧ��� = 3
    chk��������ʬ�� = 4
    chk�·����� = 5
    txtҽԺ��Ⱦ��ԭѧ��� = 59
    txt���ȴ��� = 60
    txt�ɹ����� = 61
    txt����ԭ�� = 62
    chkסԺ�ڼ�没�ػ�Σ = 15
    chk���Ѳ��� = 14
End Enum

Private Enum ��ҽ���
    txt��ҽ�������Ժ = 63
    txt��ҽ��Ժ���Ժ = 64
    chkΣ�� = 6
    chk��֢ = 7
    chk���� = 8
    txt��֤ = 65
    txt�η� = 66
    txt��ҩ = 67
    txt������� = 68
    txt���ȷ��� = 69
    txt������ҩ = 70
    txt��ҽ�豸 = 71
    txt��ҽ���� = 72
    txt��֤ʩ�� = 73
End Enum

Private Enum סԺ���
    txt�������� = 77
    txt��Ѫǰ9���� = 78
    txtѪ�� = 79
    txt����ʱ�� = 80
    txtRh = 81
    txt����״�� = 82
    txt��Ѫ��Ӧ = 83
    chkʾ�̲��� = 9
    chk���в��� = 18
    txt���ϸ�� = 84
    txt��ѪС�� = 85
    txt��ȫѪ = 86
    txt������ = 87
    txt��Ѫ�� = 88
    txt������� = 89
    txt��׵��� = 90
    txtҽѧ��ʾ = 91
    txt����ҽѧ��ʾ = 92
    chk����·�� = 19
    txt�˳�ԭ�� = 93
    chk���� = 21
    txt����ԭ�� = 94
    chk���·�� = 20
    txt��Ժ��ʽ = 95
    txtת����� = 96
    chk������� = 16
    txtԺ�ڻ������ = 97
    txt��Ժ������� = 98
    txt����������� = 99
    txt��Ժǰ�� = 100
    txt��ԺǰСʱ = 101
    txt��Ժǰ���� = 102
    txt��Ժ���� = 103
    txt��Ժ��Сʱ = 104
    txt��Ժ����� = 105
    txt����Ժ���� = 106
    txt31��Ŀ�� = 107
    txt������Сʱ = 108
    chk���� = 17
    txt����ҽʦ = 109
    txt�������� = 110
    txt����ҽʦ = 111
    txt����ҽʦ = 112
    txt������ = 113
    txt����ҽʦ = 114
    txt����ҽʦ = 115
    txtסԺҽʦ = 116
    txt�ʿ�ҽʦ = 117
    txt���λ�ʿ = 118
    txtʵϰҽʦ = 119
    txt�ʿ����� = 120
    txt�������� = 121
    txt�ʿػ�ʿ = 122
End Enum

Private Enum ������
    col������� = 0
    col������� = 1
    col��ҽ֤�� = 2
    col��ע = 3
    col��Ժ���� = 4
    col��Ժ��� = 5
    colzy���� = 6
    col�Ƿ�δ�� = 6
    col�Ƿ����� = 7
    col���� = 8
End Enum

Private Enum �������
    col��ʼ���� = 0
    col�������� = 1
    col������ҩʱ�� = 2
    COL������� = 3
    col׼������ = 4
    col�������� = 5
    col�������� = 6
    col�ٴ����� = 7
    col����ҽʦ = 8
    col������ʿ = 9
    col����1 = 10
    col����2 = 11
    col����ʼʱ�� = 12
    col�������� = 13
    colASA�ּ� = 14
    colNNIS�ּ� = 15
    col�������� = 16
    col����ҽʦ = 17
    col�п����� = 18
    col�пڲ�λ = 19
    col�ط������Ҽƻ� = 20
    col�ط�������Ŀ�� = 21
    col�пڸ�Ⱦ = 22
    col����֢ = 23
End Enum

Private Enum �������
    col����ʱ�� = 0
    col����ҩ�� = 1
    col������Ӧ = 2
End Enum

Private Enum ���Ƽ�¼
    col���Ʊ��� = 0
    COL���ƿ�ʼ���� = 1
    col���ƽ������� = 2
    col�����Ƴ��� = 3
    col���Ʒ��� = 4
    col�������� = 5
    col����Ч�� = 6
End Enum

Private Enum ���Ƽ�¼
    col���Ʊ��� = 0
    COL���ƿ�ʼ���� = 1
    col���ƽ������� = 2
    col��Ұ��λ = 3
    col������� = 4
    col�����ۼƼ��� = 5
    col����Ч�� = 6
End Enum

Private Enum ������
    rowct = 0
    rowPETCT = 1
    row˫ԴCT = 2
    rowXƬ = 3
    rowB�� = 4
    row�����Ķ�ͼ = 5
    rowMRI = 6
    rowͬλ�ؼ�� = 7
End Enum

Private Enum �걾��Դ
    SC_�걾 = 0
    SC_��ԭѧ���뼰���� = 1
    SC_�ͼ���ǰ = 2
End Enum

Private Enum ҽԺ��Ⱦ
    IC_ȷ������ = 0
    IC_��Ⱦ��λ = 1
    IC_ҽԺ��Ⱦ���� = 2
End Enum

Private Enum ICU��е
    IIC_ICU���� = 0
    IIC_��е�򵼹����� = 1
    IIC_��ʼʹ��ʱ�� = 2
    IIC_����ʹ��ʱ�� = 3
    IIC_�ۼ�ʱ�� = 4
End Enum

Private Enum ICU�����¼
    IRC_��� = 0
    IRC_ICU���� = 1
    IRC_��סʱ�� = 2
    IRC_ת��ʱ�� = 3
    IRC_����ס�ƻ� = 4
    IRC_����סԭ�� = 4
End Enum

Private Enum ������
    kss���� = 0
    kss��ҩĿ�� = 1
    kssʹ�ý׶� = 2
    kssʹ������ = 3
    KSSһ���п�Ԥ���� = 4
    KSSDDD�� = 5
    KSS������ҩ = 6
End Enum

Private Enum ����
     txtѹ�������ڼ� = 123
     txtѹ������ = 124
     txt������׹���˺� = 125
     txt������׹��ԭ�� = 126
     txt��Һ��Ӧ = 127
     txt����ҩ�� = 128
     txt�ٴ����� = 129
     chkסԺ�ڼ�����Լ�� = 13
     txt��Ժ͸�����ص�ֵ = 130
End Enum


Public Function zlRefresh(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal blnMoved As Boolean) As Boolean
'���ܣ�ˢ�»����ҽ���嵥
    Dim rsTmp As New ADODB.Recordset
    Dim StrSQL As String, lng����ID As Long
    Dim bln��ҽ As Boolean
    
    mlng����ID = lng����ID: mlng��ҳID = lng��ҳID: mblnMoved = blnMoved
    
    On Error GoTo errH
    
    StrSQL = "Select ��Ժ����ID From ������ҳ Where ����id=[1] And ��ҳid=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
    If Not rsTmp.EOF Then lng����ID = Nvl(rsTmp!��Ժ����ID, 0)
    bln��ҽ = Have��������(lng����ID, "��ҽ��")
    fraInfo(FRA_��ҽ���).Visible = bln��ҽ
    fraInfo(FRA_��ҽ���).Enabled = bln��ҽ
    mbln���� = CheckShare(300) '����ϵͳ
    fraInfo(FRA_�����뻯��).Visible = mbln����
    fraInfo(FRA_�����뻯��).Enabled = mbln����
    
    Call SetPageHeight
    Call SetScrollbar
    
    Call ClearPageData
    If mlng����ID <> 0 Then Call LoadPageData
    
    Call Form_Resize
    zlRefresh = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub chkInfo_Click(Index As Integer)
    If Not mblnCheck Then
        mblnCheck = True
        chkInfo(Index).Value = IIf(chkInfo(Index).Value = 1, 0, 1)
        mblnCheck = False
    End If
End Sub

Private Sub Form_Activate()
    Call Form_Resize
End Sub

Private Sub Form_Load()
    '�������ߴ�
    vsc.Width = GetSystemMetrics(SM_CXVSCROLL) * Screen.TwipsPerPixelX
    hsc.Height = GetSystemMetrics(SM_CXHSCROLL) * Screen.TwipsPerPixelY
    fraVH.Width = vsc.Width: fraVH.Height = hsc.Height
    fraBack.Left = 0: fraBack.Top = 0
    picBack.BackColor = fraBack.BackColor
End Sub

Private Sub SetPageHeight()
'���ܣ�����ҳ��������չ��״̬���ý���ߴ�
'˵����Tag=1��ʾ����
    Dim i As Long, intCurIdx As Integer
    
    For i = 0 To fraInfo.UBound
        If Val(picSize(i).Tag) = 0 Then
            fraInfo(i).Height = Val(fraInfo(i).Tag)
            Set picSize(i).Picture = imgSize.ListImages("-").Picture
        Else
            fraInfo(i).Height = 225
            Set picSize(i).Picture = imgSize.ListImages("+").Picture
        End If
    Next
    
    intCurIdx = 0
    For i = 1 To fraInfo.UBound
        If fraInfo(i).Enabled Then
            fraInfo(i).Top = fraInfo(intCurIdx).Top + fraInfo(intCurIdx).Height + 100
            intCurIdx = i
        End If
    Next
    fraBack.Height = fraInfo(intCurIdx).Top + fraInfo(intCurIdx).Height + fraInfo(0).Top
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    picBack.Left = 0
    picBack.Top = 0
    picBack.Width = Me.ScaleWidth
    picBack.Height = Me.ScaleHeight
    
    Call SetScrollbar
    
    If hsc.Visible Then
        hsc.Left = 0
        hsc.Top = picBack.ScaleHeight - hsc.Height
        hsc.Width = picBack.ScaleWidth - IIf(vsc.Visible, vsc.Width, 0)
    End If
    If vsc.Visible Then
        vsc.Top = 0
        vsc.Left = picBack.ScaleWidth - vsc.Width
        vsc.Height = picBack.ScaleHeight - IIf(hsc.Visible, hsc.Height, 0)
    End If
    If fraVH.Visible Then
        fraVH.Left = vsc.Left
        fraVH.Top = hsc.Top
        fraVH.Refresh
    End If
End Sub

Private Sub SetScrollbar()
'���ܣ����ݵ�ǰ����ߴ����ù������ɼ��Լ��������
    If fraBack.Width + IIf(vsc.Visible, vsc.Width, 0) <= picBack.ScaleWidth Then
        hsc.Visible = False
    Else
        hsc.Min = 0
        hsc.SmallChange = 5
        hsc.LargeChange = 50
        If Not hsc.Visible Then hsc.Value = 0
        hsc.Visible = True
    End If
    
    If fraBack.Height + IIf(hsc.Visible, hsc.Height, 0) <= picBack.ScaleHeight Then
        vsc.Visible = False
    Else
        vsc.Min = 0
        vsc.SmallChange = 5
        vsc.LargeChange = 50
        If Not vsc.Visible Then vsc.Value = 0
        vsc.Visible = True
    End If
    
    If hsc.Visible Then
        hsc.Max = (picBack.ScaleWidth - fraBack.Width - IIf(vsc.Visible, vsc.Width, 0)) / Screen.TwipsPerPixelX
    End If
    
    If vsc.Visible Then
        vsc.Max = (picBack.ScaleHeight - fraBack.Height - IIf(hsc.Visible, hsc.Height, 0)) / Screen.TwipsPerPixelY
    End If
    
    fraVH.Visible = vsc.Visible And hsc.Visible
End Sub

Private Sub hsc_Change()
    Call hsc_Scroll
End Sub

Private Sub picSize_Click(Index As Integer)
    picSize(Index).Tag = IIf(Val(picSize(Index).Tag) = 0, 1, 0)
    Call SetPageHeight
    Call Form_Resize
    If Not vsc.Visible Then fraBack.Top = 0
    If Not hsc.Visible Then fraBack.Left = 0
End Sub

Private Sub txtInfo_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txtInfo(Index))
End Sub

Private Sub vsc_Change()
    Call vsc_Scroll
End Sub

Private Sub hsc_Scroll()
    fraBack.Left = hsc.Value * Screen.TwipsPerPixelX
End Sub

Private Sub vsc_Scroll()
    fraBack.Top = vsc.Value * Screen.TwipsPerPixelY
End Sub

Private Sub vsDiagXY_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call vsDiagXY.ShowCell(NewRow, NewCol)
End Sub

Private Sub vsDiagZY_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call vsDiagZY.ShowCell(NewRow, NewCol)
End Sub

Private Sub vsfMain__AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call vsfMain.ShowCell(NewRow, NewCol)
End Sub

Private Sub VSFlexGrid1_Click()

End Sub

Private Sub vsOPS_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call vsOPS.ShowCell(NewRow, NewCol)
End Sub

Private Sub vsRadiotherapy_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call vsRadiotherapy.ShowCell(NewRow, NewCol)
End Sub

Private Sub vsChemotherapy_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call vsChemotherapy.ShowCell(NewRow, NewCol)
End Sub


Private Sub ClearPageData()
'���ܣ������ҳ�е�����
    Dim objTmp As Object
    Dim i As Long, j As Long
    
    mblnCheck = True
    
    For Each objTmp In Me.Controls
        If TypeName(objTmp) = "TextBox" Then
            objTmp.Text = ""
        ElseIf TypeName(objTmp) = "CheckBox" Then
            objTmp.Value = 0
        End If
    Next
    
    With vsDiagXY
        .Cell(flexcpText, .FixedRows, .FixedCols, .Rows - 1, .Cols - 1) = ""
        For i = .Rows - 1 To .FixedRows Step -1
            If .TextMatrix(i, col�������) = "" Then
                .RemoveItem i
            End If
        Next
    End With
    With vsDiagZY
        .Cell(flexcpText, .FixedRows, .FixedCols, .Rows - 1, .Cols - 1) = ""
        For i = .Rows - 1 To .FixedRows Step -1
            If .TextMatrix(i, col�������) = "" Then
                .RemoveItem i
            End If
        Next
    End With
    With vsOPS
        .Rows = .FixedRows
        .Rows = .FixedRows + 1
    End With
    With vsAller
        .Rows = .FixedRows
        .Rows = .FixedRows + 1
    End With
    
    With vsChemotherapy
        .Rows = .FixedRows
        .Rows = .FixedRows + 1
    End With
    
    With vsFlxAddICU
        .Rows = .FixedRows
        .Rows = .FixedRows + 1
    End With
    
    With vsfMain
        .Rows = .FixedRows
        .Rows = .FixedRows + 10
        .Cols = .FixedCols
        .Cols = .FixedCols + 10
    End With
    
    With vsRadiotherapy
        .Rows = .FixedRows
        .Rows = .FixedRows + 1
    End With
    
    lstAdvEvent.Clear
        
    mblnCheck = False
End Sub

Private Function GetRow(ByVal lng������� As Long) As Long
'���ܣ�����ָ��������͵ĵ�һ�����
    If InStr(",11,12,13,", "," & lng������� & ",") > 0 Then
        GetRow = vsDiagZY.FindRow(CStr(lng�������), , colzy����)
    Else
        GetRow = vsDiagXY.FindRow(CStr(lng�������), , col����)
    End If
End Function

Private Function LoadPageData() As Boolean
'���ܣ���ȡ���˵���ҳ��Ϣ
    Dim rsTmp As New ADODB.Recordset
    Dim StrSQL As String, i As Long, j As Long
    Dim lngRow As Long, varTmp As Variant
    Dim strTmp As String
    Dim bln��ҳ��� As Boolean, bln�ֻ��̶� As Boolean
    Dim blnѹ�� As Boolean, bln����׹�� As Boolean
    
    On Error GoTo errH

    Screen.MousePointer = 11
    mblnCheck = True
    
    '��ʼ������������Ŀ
    Call FillVsf
    
    '������Ϣ����
    '---------------------------------------------------------------
    StrSQL = "Select סԺ��,����,�Ա�,��������,�����ص�,���֤��,����֤��,����,����,������,����,email,QQ From ������Ϣ Where ����ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID)

    txtInfo(txt��������).Text = Nvl(rsTmp!������)
    txtInfo(txtסԺ����).Text = mlng��ҳID
    txtInfo(txt����).Text = Nvl(rsTmp!����)
    txtInfo(txt�Ա�).Text = Nvl(rsTmp!�Ա�)
    txtInfo(txt����Ժ����).Text = "��"
    If Format(rsTmp!��������, "HH:mm") <> "00:00" Then
        txtInfo(txt��������).Text = Format(rsTmp!��������, "yyyy-MM-dd HH:mm")
    Else
        txtInfo(txt��������).Text = Format(rsTmp!��������, "yyyy-MM-dd")
    End If

    txtInfo(txt�����ص�).Text = Nvl(rsTmp!�����ص�)
    txtInfo(txt����).Text = Nvl(rsTmp!����)
    txtInfo(txt���֤��).Text = Nvl(rsTmp!���֤��)
    txtInfo(txt����).Text = Nvl(rsTmp!����)
    txtInfo(txt����).Text = Nvl(rsTmp!����)
    txtInfo(txt����֤��).Text = Nvl(rsTmp!����֤��)
    txtInfo(txtEmail).Text = Nvl(rsTmp!Email)
    txtInfo(txtqq).Text = Nvl(rsTmp!QQ)
    '�����Ŷ�ȡ
    StrSQL = "select ������ from סԺ������¼ where ����ID=[1] and ��ҳID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
    If rsTmp.RecordCount <> 0 Then
        txtInfo(txt������).Text = Nvl(rsTmp!������)
    End If
    '������ҳ����
    '---------------------------------------------------------------
    StrSQL = "Select A.*,B.���� as ��Ժ����,C.���� as ��Ժ����" & _
        " From ������ҳ A,���ű� B,���ű� C" & _
        " Where A.��Ժ����ID=B.ID And A.��Ժ����ID=C.ID" & _
        " And A.����ID=[1] And A.��ҳID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)

    txtInfo(txt���ʽ).Text = Nvl(rsTmp!ҽ�Ƹ��ʽ)
    '���۲�����סԺ��
    If Nvl(rsTmp!��������, 0) <> 0 Then
        lblInfo(txt��������).Visible = False
        txtInfo(txt��������).Visible = False
    End If
    chkInfo(chk����Ժ).Value = Nvl(rsTmp!����Ժ, 0)
    txtInfo(txt������).Text = Nvl(rsTmp!������)
    
    txtInfo(txt����).Text = Nvl(rsTmp!����)
    txtInfo(txt����).Text = Nvl(rsTmp!����)
    '�������
    txtInfo(txt���).Text = IIf(rsTmp!��� & "" = "0", "", rsTmp!��� & "")
    txtInfo(txt����).Text = IIf(rsTmp!���� & "" = "0", "", rsTmp!���� & "")
    txtInfo(txtְҵ).Text = Nvl(rsTmp!ְҵ)
    txtInfo(txt����).Text = Nvl(rsTmp!����״��)
    txtInfo(txt��ͥ��ַ).Text = Nvl(rsTmp!��ͥ��ַ)
    txtInfo(txt��ͥ�绰).Text = Nvl(rsTmp!��ͥ�绰)
    txtInfo(txt��ͥ�ʱ�).Text = Nvl(rsTmp!��ͥ��ַ�ʱ�)
    txtInfo(txt���ڵ�ַ).Text = Nvl(rsTmp!���ڵ�ַ)
    txtInfo(txt�����ʱ�).Text = Nvl(rsTmp!���ڵ�ַ�ʱ�)
    
    txtInfo(txt������λ).Text = Nvl(rsTmp!��λ��ַ)
    txtInfo(txt��λ�绰).Text = Nvl(rsTmp!��λ�绰)
    txtInfo(txt��λ�ʱ�).Text = Nvl(rsTmp!��λ�ʱ�)
    txtInfo(txt��ϵ������).Text = Nvl(rsTmp!��ϵ������)
    txtInfo(txt��ϵ�˹�ϵ).Text = Nvl(rsTmp!��ϵ�˹�ϵ)
    txtInfo(txt��ϵ�˵绰).Text = Nvl(rsTmp!��ϵ�˵绰)
    txtInfo(txt��ϵ�˵�ַ).Text = Nvl(rsTmp!��ϵ�˵�ַ)
    If Not IsNull(rsTmp!����) Then
        txtInfo(txt����).Text = Nvl(rsTmp!����)
    End If

    txtInfo(txt��Ժ;��).Text = Nvl(rsTmp!��Ժ��ʽ)
    txtInfo(txt��Ժʱ��).Text = Format(rsTmp!��Ժ����, "yyyy-MM-dd HH:mm")
    txtInfo(txt��Ժ����).Text = rsTmp!��Ժ����
    
    txtInfo(txt��Ժʱ��).Text = Format(Nvl(rsTmp!��Ժ����), "yyyy-MM-dd HH:mm")
    txtInfo(txt��Ժ����).Text = rsTmp!��Ժ����
    If Not IsNull(rsTmp!��Ժ����) Then
        txtInfo(txtסԺ����).Text = DateDiff("d", rsTmp!��Ժ����, rsTmp!��Ժ����)
    Else
        txtInfo(txtסԺ����).Text = DateDiff("d", rsTmp!��Ժ����, zlDatabase.Currentdate)
    End If
    If Val(txtInfo(txtסԺ����).Text) = 0 Then txtInfo(txtסԺ����).Text = "1"
    
     txtInfo(txt��Ժ���).Text = Nvl(rsTmp!��Ժ����)
    chkInfo(chk�Ƿ�ȷ��).Value = Nvl(rsTmp!�Ƿ�ȷ��, 0)
    If chkInfo(chk�Ƿ�ȷ��).Value = 1 Then
        txtInfo(txtȷ������).Text = Format(Nvl(rsTmp!ȷ������), "yyyy-MM-dd HH:mm")
    End If
    chkInfo(chk��������ʬ��).Value = Nvl(rsTmp!ʬ���־, 0)
    chkInfo(chk�·�����).Value = Nvl(rsTmp!�·�����, 0)
    txtInfo(txt���ȴ���).Text = Nvl(rsTmp!���ȴ���)
    If Val(txtInfo(txt���ȴ���).Text) <> 0 Then
        txtInfo(txt�ɹ�����).Text = Nvl(rsTmp!�ɹ�����)
    End If
    
    txtInfo(txt�������).Text = Nvl(rsTmp!��ҽ�������)
    
    txtInfo(txtѪ��).Text = Nvl(rsTmp!Ѫ��)
    chkInfo(chk����).Value = IIf(Nvl(rsTmp!�����־, 0) = 0, 0, 1)
    If chkInfo(chk����).Value = 1 Then
        txtInfo(txt��������).Text = IIf(Nvl(rsTmp!�����־, 0) = 9, "", Nvl(rsTmp!��������, 0)) & _
            Decode(Nvl(rsTmp!�����־, 0), 1, "��", 2, "��", 3, "��", 4, "��", 9, "����")
    End If
    txtInfo(txt����ҽʦ).Text = Nvl(rsTmp!����ҽʦ)
    txtInfo(txtסԺҽʦ).Text = Nvl(rsTmp!סԺҽʦ)
    txtInfo(txt���λ�ʿ).Text = Nvl(rsTmp!���λ�ʿ)
    '���ʱ��
    If Nvl(rsTmp!״̬, 0) = 1 Then
        txtInfo(txt���ʱ��).Text = "��δ���"
    Else
        StrSQL = "Select ��ʼʱ�� From ���˱䶯��¼" & _
            " Where ����ID=[1] And ��ҳID=[2] And ��ʼԭ�� IN(2,1) And ��ʼʱ�� is Not Null Order by ��ʼԭ�� Desc"
        Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
        If Not rsTmp.EOF Then
            txtInfo(txt���ʱ��).Text = Format(rsTmp!��ʼʱ��, "yyyy-MM-dd HH:mm")
        End If
    End If
    
    '�����ӱ���
    '---------------------------------------------------------------
    StrSQL = "Select a.����ID,a.��ҳID,a.��Ϣ��,a.��Ϣֵ,b.���� From ������ҳ�ӱ� a " & _
            ",������Ŀ b" & " where a.��Ϣ��=b.����(+) And a.����ID=[1] And a.��ҳID=[2] Order by a.��Ϣ��"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
    For i = 1 To rsTmp.RecordCount
        Select Case UCase(Nvl(rsTmp!��Ϣ��))
            Case "������������"
                txtInfo(txt������������).Text = Nvl(rsTmp!��Ϣֵ)
            Case "��������������"
                txtInfo(txt����������).Text = Nvl(rsTmp!��Ϣֵ) & IIf(Nvl(rsTmp!��Ϣֵ) = "", "", " ��")
            Case "��������Ժ����"
                txtInfo(txt��������Ժ����).Text = Nvl(rsTmp!��Ϣֵ) & IIf(Nvl(rsTmp!��Ϣֵ) = "", "", " ��")
            Case "����"
                txtInfo(txt����).Text = Nvl(rsTmp!��Ϣֵ)
            Case "��Ժ����"
                txtInfo(txt��Ժ����).Text = Nvl(rsTmp!��Ϣֵ)
            Case "��Ժǰ����Ժ����"
                chkInfo(chk��Ժǰ����Ժ����).Value = Val(Nvl(rsTmp!��Ϣֵ, 0))
            Case "ת�Ƽ�¼"
                varTmp = Split(Nvl(rsTmp!��Ϣֵ), ",")
                If UBound(varTmp) >= 0 Then txtInfo(txtת��1).Text = varTmp(0)
                If UBound(varTmp) >= 1 Then txtInfo(txtת��2).Text = varTmp(1)
                If UBound(varTmp) >= 2 Then txtInfo(txtת��3).Text = varTmp(2)
            Case "��Ժ����"
                txtInfo(txt��Ժ����).Text = Nvl(rsTmp!��Ϣֵ)
            Case "�����"
                txtInfo(txt�����).Text = Nvl(rsTmp!��Ϣֵ)
            Case "�ֻ��̶�"
                If Nvl(rsTmp!��Ϣֵ) <> "" Then
                    txtInfo(txt�ֻ��̶�).Text = Nvl(rsTmp!��Ϣֵ)
                End If
            Case "����������"
                If Nvl(rsTmp!��Ϣֵ) <> "" Then
                    txtInfo(txt����������).Text = Nvl(rsTmp!��Ϣֵ)
                End If
            Case "��ԭѧ���"
                chkInfo(chkҽԺ��Ⱦ����ԭѧ���).Value = Val(Nvl(rsTmp!��Ϣֵ, 0))
            Case "����ʱ��"
                If Not (IsNull(rsTmp!��Ϣֵ) Or Not IsDate(rsTmp!��Ϣֵ)) Then
                    txtInfo(txt����ʱ��).Text = rsTmp!��Ϣֵ
                End If
            Case "��������ԭ��"
                txtInfo(txt����ԭ��).Text = Nvl(rsTmp!��Ϣֵ)
            Case "���Ȳ���"
                txtInfo(txt����ԭ��).Text = Nvl(rsTmp!��Ϣֵ)
            Case "��ҽΣ��"
                chkInfo(chkΣ��).Value = Val(Nvl(rsTmp!��Ϣֵ, 0))
            Case "��ҽ��֢"
                chkInfo(chk��֢).Value = Val(Nvl(rsTmp!��Ϣֵ, 0))
            Case "��ҽ����"
                chkInfo(chk����).Value = Val(Nvl(rsTmp!��Ϣֵ, 0))
            Case "��ҽ���ȷ���"
                txtInfo(txt���ȷ���).Text = Nvl(rsTmp!��Ϣֵ)
            Case "������ҩ�Ƽ�"
                txtInfo(txt������ҩ).Text = Nvl(rsTmp!��Ϣֵ)
            Case "��ҽ�豸"
                txtInfo(txt��ҽ�豸).Text = Nvl(rsTmp!��Ϣֵ)
            Case "��ҽ����"
                txtInfo(txt��ҽ����).Text = Nvl(rsTmp!��Ϣֵ)
            Case "��֤ʩ��"
                txtInfo(txt��֤ʩ��).Text = Nvl(rsTmp!��Ϣֵ)
            Case "��������"
                txtInfo(txt��������).Text = GetNameByCode("��������", Nvl(rsTmp!��Ϣֵ))
'            Case UCase("HBsAg")
'                txtinfo(txtHBsAg).Text = Nvl(rsTmp!��Ϣֵ)
'            Case UCase("HCV-Ab")
'                txtinfo(txtHCVAb).Text = Nvl(rsTmp!��Ϣֵ)
'            Case UCase("HIV-Ab")
'                txtinfo(txtHIVAb).Text = Nvl(rsTmp!��Ϣֵ)
            Case UCase("Rh")
                txtInfo(txtRh).Text = Nvl(rsTmp!��Ϣֵ)
            Case "��Ѫ���"
                txtInfo(txt��Ѫǰ9����).Text = Nvl(rsTmp!��Ϣֵ)
            Case "����ʱ��"
                If Nvl(rsTmp!��Ϣֵ) <> "" Then
                    If Format(rsTmp!��Ϣֵ, "HH:mm") <> "00:00" Then
                        txtInfo(txt����ʱ��).Text = Format(rsTmp!��Ϣֵ, "yyyy-MM-dd HH:mm")
                    Else
                        txtInfo(txt����ʱ��).Text = Format(rsTmp!��Ϣֵ, "yyyy-MM-dd")
                    End If
                End If
            Case "����״��"
                txtInfo(txt����״��).Text = Decode(Val(Nvl(rsTmp!��Ϣֵ, 0)), 0, "δ����", 1, "����1̥", 2, "����2̥������", 4, "4-����")
            Case "��Һ��Ӧ"
                txtInfo(txt��Һ��Ӧ).Text = Nvl(rsTmp!��Ϣֵ)
            Case "��Ѫ��Ӧ"
                txtInfo(txt��Ѫ��Ӧ).Text = Decode(Val(Nvl(rsTmp!��Ϣֵ, 0)), 0, "��", 1, "��", 2, "δ��")
            Case "ʾ�̲���"
                chkInfo(chkʾ�̲���).Value = Val(Nvl(rsTmp!��Ϣֵ, 0))
            Case "���в���"
                chkInfo(chk���в���).Value = Val(Nvl(rsTmp!��Ϣֵ, 0))
            Case "���ϸ��"
                txtInfo(txt���ϸ��).Text = Nvl(rsTmp!��Ϣֵ)
            Case "��ѪС��"
                txtInfo(txt��ѪС��).Text = Nvl(rsTmp!��Ϣֵ)
            Case "��Ѫ��"
                txtInfo(txt��Ѫ��).Text = Nvl(rsTmp!��Ϣֵ)
            Case "��ȫѪ"
                txtInfo(txt��ȫѪ).Text = Nvl(rsTmp!��Ϣֵ)
            Case "������"
                txtInfo(txt������).Text = Nvl(rsTmp!��Ϣֵ)
            Case "�������"
                txtInfo(txt�������).Text = Nvl(rsTmp!��Ϣֵ)
            Case "��׵���"
                txtInfo(txt��׵���).Text = Nvl(rsTmp!��Ϣֵ)
            Case "ҽѧ��ʾ"
                txtInfo(txtҽѧ��ʾ).Text = Nvl(rsTmp!��Ϣֵ)
            Case "����ҽѧ��ʾ"
                txtInfo(txt����ҽѧ��ʾ).Text = Nvl(rsTmp!��Ϣֵ)
            Case "��Ժ��ʽ"
                txtInfo(txt��Ժ��ʽ).Text = Nvl(rsTmp!��Ϣֵ)
            Case "��Ժת��"
                txtInfo(txtת�����).Text = Nvl(rsTmp!��Ϣֵ)
            Case "����ʱ��"
                '�����ʽ:��Ժǰ(�죬Сʱ,����)|��Ժ��(�죬Сʱ,����)
                txtInfo(txt��Ժǰ��).Text = Split(Split(Nvl(rsTmp!��Ϣֵ), "|")(0) & ",", ",")(0)
                txtInfo(txt��ԺǰСʱ).Text = Split(Split(Nvl(rsTmp!��Ϣֵ), "|")(0) & ",", ",")(1)
                txtInfo(txt��Ժǰ����).Text = Split(Split(Nvl(rsTmp!��Ϣֵ), "|")(0) & ",", ",")(2)
                txtInfo(txt��Ժ����).Text = Split(Split(Nvl(rsTmp!��Ϣֵ), "|")(1) & ",", ",")(0)
                txtInfo(txt��Ժ��Сʱ).Text = Split(Split(Nvl(rsTmp!��Ϣֵ) & "|", "|")(1) & ",", ",")(1)
                txtInfo(txt��Ժ�����).Text = Split(Split(Nvl(rsTmp!��Ϣֵ) & "|", "|")(1) & ",", ",")(2)
            Case "����Ժ�ƻ�����"
                lblInfo(txt����Ժ����).Caption = "��Ժ" & IIf(Nvl(rsTmp!��Ϣֵ, "0") = "0", "31", "7") & "��������Ժ�ƻ�"
            Case "31������סԺ"
                If Nvl(rsTmp!��Ϣֵ) <> "" Then
                    txtInfo(txt31��Ŀ��).Text = Nvl(rsTmp!��Ϣֵ)
                    txtInfo(txt����Ժ����).Text = "��"
                Else
                    txtInfo(txt����Ժ����).Text = "��"
                End If
            Case "������ʹ��ʱ��"
                txtInfo(txt������Сʱ).Text = Nvl(rsTmp!��Ϣֵ)
            Case "������"
                txtInfo(txt������).Text = Nvl(rsTmp!��Ϣֵ)
            Case "����ҽʦ"
                txtInfo(txt����ҽʦ).Text = Nvl(rsTmp!��Ϣֵ)
            Case "����ҽʦ"
                txtInfo(txt����ҽʦ).Text = Nvl(rsTmp!��Ϣֵ)
            Case "����ҽʦ"
                txtInfo(txt����ҽʦ).Text = Nvl(rsTmp!��Ϣֵ)
            Case "����ҽʦ"
                txtInfo(txt����ҽʦ).Text = Nvl(rsTmp!��Ϣֵ)
            Case "ʵϰҽʦ"
                txtInfo(txtʵϰҽʦ).Text = Nvl(rsTmp!��Ϣֵ)
            Case "�ʿ�ҽʦ"
                txtInfo(txt�ʿ�ҽʦ).Text = Nvl(rsTmp!��Ϣֵ)
            Case "�ʿػ�ʿ"
                txtInfo(txt�ʿػ�ʿ).Text = Nvl(rsTmp!��Ϣֵ)
            Case "��������"
                txtInfo(txt��������).Text = Nvl(rsTmp!��Ϣֵ)
            Case "��ҳ��������"
                txtInfo(txt�ʿ�����).Text = Nvl(rsTmp!��Ϣֵ)
'            Case "CT"
'                chkInfo(chkCT).Value = Val(Nvl(rsTmp!��Ϣֵ, 0))
'            Case "MRI"
'                chkInfo(chkMRI).Value = Val(Nvl(rsTmp!��Ϣֵ, 0))
'            Case "��ɫ������"
'                chkInfo(chk��ɫ������).Value = Val(Nvl(rsTmp!��Ϣֵ, 0))
'            Case "������4"
'                vsTSJC.TextMatrix(vsTSJC.FixedRows + 0, 1) = Nvl(rsTmp!��Ϣֵ)
'                vsTSJC.Cell(flexcpData, vsTSJC.FixedRows + 0, 1) = vsTSJC.TextMatrix(vsTSJC.FixedRows + 0, 1)
'            Case "������5"
'                vsTSJC.TextMatrix(vsTSJC.FixedRows + 1, 1) = Nvl(rsTmp!��Ϣֵ)
'                vsTSJC.Cell(flexcpData, vsTSJC.FixedRows + 1, 1) = vsTSJC.TextMatrix(vsTSJC.FixedRows + 1, 1)
'            Case "������6"
'                vsTSJC.TextMatrix(vsTSJC.FixedRows + 2, 1) = Nvl(rsTmp!��Ϣֵ)
'                vsTSJC.Cell(flexcpData, vsTSJC.FixedRows + 2, 1) = vsTSJC.TextMatrix(vsTSJC.FixedRows + 2, 1)
            Case "ѹ�������ڼ�"
                txtInfo(txtѹ�������ڼ�).Text = Nvl(rsTmp!��Ϣֵ, " ")
            Case "ѹ������"
                txtInfo(txtѹ������).Text = Nvl(rsTmp!��Ϣֵ, " ")
            Case "������׹���˺�"
                txtInfo(txt������׹���˺�).Text = Nvl(rsTmp!��Ϣֵ, " ")
            Case "������׹��ԭ��"
                txtInfo(txt������׹��ԭ��).Text = Nvl(rsTmp!��Ϣֵ, " ")
            Case "�没�ز�Σ"
                chkInfo(chkסԺ�ڼ�没�ػ�Σ).Value = Val(Nvl(rsTmp!��Ϣֵ))
            Case "���Ѳ���"
                chkInfo(chk���Ѳ���).Value = Val(Nvl(rsTmp!��Ϣֵ))
            Case "�ٴ�·��" '���ݺ���ʡ��ҳ
                chkInfo(chk����·��).Value = IIf(Val(Nvl(rsTmp!��Ϣֵ)) >= 1, 1, 0)
            Case "�˳�ԭ��"
                If Nvl(rsTmp!��Ϣֵ) = "1" Then
                    chkInfo(chk���·��).Value = 1
                Else
                    chkInfo(chk���·��).Value = 0
                    txtInfo(txt�˳�ԭ��).Text = Nvl(rsTmp!��Ϣֵ)
                End If
            Case "����ԭ��"
                If Nvl(rsTmp!��Ϣֵ) = "0" Then
                    chkInfo(chk����).Value = 0
                Else
                    chkInfo(chk����).Value = 1
                    txtInfo(txt����ԭ��).Text = Trim(Nvl(rsTmp!��Ϣֵ))
                End If
            Case "��Ժ����"
                chkInfo(chk�������).Value = 1
                txtInfo(txt��Ժ�������).Text = Val(Nvl(rsTmp!��Ϣֵ))
            Case "Ժ�ڻ���"
                chkInfo(chk�������).Value = 1
                txtInfo(txtԺ�ڻ������).Text = Val(Nvl(rsTmp!��Ϣֵ))
            Case "�������"
                If Nvl(rsTmp!��Ϣֵ) = "0" Then
                    chkInfo(chk�������).Value = 0
                Else
                    chkInfo(chk�������).Value = 1
                    txtInfo(txt�����������).Text = Trim(Nvl(rsTmp!��Ϣֵ))
                End If
            Case "CT"
                 vsCheck.TextMatrix(rowct, 1) = Decode(Val(Nvl(rsTmp!��Ϣֵ)), 1, "1-����", 2, "2-����", 3, "3-δ��", "")
            Case "PETCT"
                vsCheck.TextMatrix(rowPETCT, 1) = Decode(Val(Nvl(rsTmp!��Ϣֵ)), 1, "1-����", 2, "2-����", 3, "3-δ��", "")
            Case "˫ԴCT"
                vsCheck.TextMatrix(row˫ԴCT, 1) = Decode(Val(Nvl(rsTmp!��Ϣֵ)), 1, "1-����", 2, "2-����", 3, "3-δ��", "")
            Case "XƬ"
                vsCheck.TextMatrix(rowXƬ, 1) = Decode(Val(Nvl(rsTmp!��Ϣֵ)), 1, "1-����", 2, "2-����", 3, "3-δ��", "")
            Case "B��"
                vsCheck.TextMatrix(rowB��, 1) = Decode(Val(Nvl(rsTmp!��Ϣֵ)), 1, "1-����", 2, "2-����", 3, "3-δ��", "")
            Case "�����Ķ�ͼ"
                vsCheck.TextMatrix(row�����Ķ�ͼ, 1) = Decode(Val(Nvl(rsTmp!��Ϣֵ)), 1, "1-����", 2, "2-����", 3, "3-δ��", "")
            Case "MRI"
                vsCheck.TextMatrix(rowMRI, 1) = Decode(Val(Nvl(rsTmp!��Ϣֵ)), 1, "1-����", 2, "2-����", 3, "3-δ��", "")
            Case "ͬλ�ؼ��"
                vsCheck.TextMatrix(rowͬλ�ؼ��, 1) = Decode(Val(Nvl(rsTmp!��Ϣֵ)), 1, "1-����", 2, "2-����", 3, "3-δ��", "")
            Case "��Һҩ��"
                txtInfo(txt����ҩ��).Text = Nvl(rsTmp!��Ϣֵ)
            Case "��Һ����"
                txtInfo(txt�ٴ�����).Text = Nvl(rsTmp!��Ϣֵ)
            Case "����Լ��"
                chkInfo(chkסԺ�ڼ�����Լ��).Value = Val(Nvl(rsTmp!��Ϣֵ))
            Case "͸�����ص�ֵ"
                txtInfo(txt��Ժ͸�����ص�ֵ).Text = Nvl(rsTmp!��Ϣֵ)
            Case Else
                '�������������
                If Left(Nvl(rsTmp!��Ϣ��), 3) = "������" And Not IsNull(rsTmp!��Ϣֵ) Then
                    With vsKSS
                        For j = .FixedRows To .Rows - 1
                            .RowData(j) = GetIDTmp(rsTmp!��Ϣֵ)
                            If .RowData(j) <> 0 Then
                                .TextMatrix(j, 1) = rsTmp!��Ϣֵ
                                Exit For
                            End If
                        Next
                        If j > .Rows - 1 Then
                            .AddItem ""
                            .RowData(.Rows - 1) = GetIDTmp(rsTmp!��Ϣֵ)
                            If .RowData(.Rows - 1) <> 0 Then
                                .TextMatrix(.Rows - 1, 1) = rsTmp!��Ϣֵ
                            End If
                        End If
                    End With
                Else
                    '������Ŀ
                    If Not IsNull(rsTmp("����")) Then
                        With vsfMain
                            For j = 0 To vsfMain.Cols - 1 Step 3
                                lngRow = vsfMain.FindRow(rsTmp("��Ϣ��"), , j)
                                If lngRow >= 0 Then
                                    If vsfMain.TextMatrix(lngRow, j) = rsTmp("��Ϣ��") Then
                                        If vsfMain.TextMatrix(lngRow, j + 2) = "�Ƿ�" Then
                                            vsfMain.Cell(flexcpChecked, lngRow, j + 1) = IIf(rsTmp("��Ϣֵ") = 0, 2, 1)
                                            Exit For
                                        Else
                                            vsfMain.TextMatrix(lngRow, j + 1) = rsTmp("��Ϣֵ") & ""
                                            Exit For
                                        End If
                                    End If
                                End If
                            Next j
                        End With
                    End If
                End If
        End Select
        rsTmp.MoveNext
    Next
    
    '�Զ���ȡת�ƿ��Ҽ��������(�����)
    '---------------------------------------------------------------
    If txtInfo(txtת��1).Text = "" And txtInfo(txtת��2).Text = "" And txtInfo(txtת��3).Text = "" Then
        StrSQL = _
            " Select B.����" & _
            " From ���˱䶯��¼ A,���ű� B" & _
            " Where A.����ID=[1] And A.��ҳID=[2]" & _
            " And A.����ID=B.ID And A.��ʼԭ��=3 And A.��ʼʱ�� is Not NULL" & _
            " Order by A.��ʼʱ��"
        Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
        For i = 1 To rsTmp.RecordCount
            If i = 1 Then
                txtInfo(txtת��1).Text = rsTmp!����
            ElseIf i = 2 Then
                txtInfo(txtת��2).Text = rsTmp!����
            ElseIf i = 3 Then
                txtInfo(txtת��3).Text = rsTmp!����
                Exit For
            End If
            rsTmp.MoveNext
        Next
    End If

    If txtInfo(txt��Ժ����).Text = "" Then
        StrSQL = "Select B.�����" & _
            " From ������ҳ A,��λ״����¼ B" & _
            " Where A.����ID=[1] And A.��ҳID=[2]" & _
            " And A.��Ժ����ID=B.����ID And A.��Ժ����=B.����"
        Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
        If Not rsTmp.EOF Then txtInfo(txt��Ժ����).Text = Nvl(rsTmp!�����)
    End If

    If txtInfo(txt��Ժ����).Text = "" Then
        StrSQL = "Select B.�����" & _
            " From ������ҳ A,��λ״����¼ B" & _
            " Where A.����ID=[1] And A.��ҳID=[2]" & _
            " And A.��ǰ����ID=B.����ID And A.��Ժ����=B.����"
        Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
        If Not rsTmp.EOF Then txtInfo(txt��Ժ����).Text = Nvl(rsTmp!�����)
    End If
    
    '��ҽ���
    '---------------------------------------------------------------
'    str���ƽ�� = Get���ƽ��
'    vsDiagXY.ColData(col��Ժ���) = str���ƽ��

    '�ж���ҳ�Ƿ�������
    StrSQL = "Select 1 From ������ϼ�¼ Where ����ID=[1] And ��ҳID=[2] And ��¼��Դ=3  And RowNum<2"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
    bln��ҳ��� = rsTmp.RecordCount > 0
    If bln��ҳ��� Then
        strTmp = " and a.��¼��Դ=3 "
    Else
        strTmp = " And a.��¼��Դ IN(1,2,3,4) "
    End If
    'ȱʡ����ʼ��
    With vsDiagXY
        '1-��ҽ�������;2-��ҽ��Ժ���;3-��Ժ���(�������);5-Ժ�ڸ�Ⱦ;6-�������;7-�����ж���;10-����֢
        .TextMatrix(1, col����) = 1
        .TextMatrix(2, col����) = 2
        .TextMatrix(3, col����) = 3
        .TextMatrix(4, col����) = 3
        .TextMatrix(5, col����) = 5
        .TextMatrix(6, col����) = 10
        .TextMatrix(7, col����) = 6
        .TextMatrix(8, col����) = 7
    End With

    '��ȡ������Դ�����
    StrSQL = "Select a.��ע,a.ID,a.����ID,a.��ҳID,a.ҽ��ID,a.��¼��Դ,a.��ϴ���,a.�������,a.����ID,a.�������,a.����ID,a.��Ժ����," & _
        " a.���ID,a.֤��ID,a.�������,a.��Ժ���,a.�Ƿ�δ��,a.�Ƿ�����,a.��¼����,a.��¼��,a.ȡ��ʱ��,a.ȡ����,a.����ID, b.���� As ��������, c.���� As ��ϱ��� " & _
        " From ������ϼ�¼ A, ��������Ŀ¼ B, �������Ŀ¼ C" & _
        " Where a.����id = b.Id(+) And a.���id = c.Id(+)  And a.������� IN(1,2,3,5,6,7,10,21)" & _
        strTmp & _
        " And a.ȡ��ʱ�� is Null And a.����ID=[1] And a.��ҳID=[2]" & _
        " Order by a.�������,a.��¼��Դ Desc,a.��ϴ���,a.ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
    If Not rsTmp.EOF Then
        With vsDiagXY
            StrSQL = "1,2,3,5,6,7,10,21"
            For i = 0 To UBound(Split(StrSQL, ","))
                rsTmp.Filter = "��¼��Դ=3 And �������=" & Split(StrSQL, ",")(i)
                If Val(Split(StrSQL, ",")(i)) <> 21 Then
                    If rsTmp.EOF Then
                        rsTmp.Filter = "��¼��Դ=2 And �������=" & Split(StrSQL, ",")(i)
                    End If
                    If rsTmp.EOF Then
                        rsTmp.Filter = "��¼��Դ=1 And �������=" & Split(StrSQL, ",")(i)
                    End If
                End If
                If rsTmp.EOF Then
                    rsTmp.Filter = "��¼��Դ=4 And �������=" & Split(StrSQL, ",")(i)
                End If

                If Val(Split(StrSQL, ",")(i)) = 21 Then
                    '21-��ԭѧ���
                    If Not rsTmp.EOF Then
                        txtInfo(txtҽԺ��Ⱦ��ԭѧ���).Text = Nvl(rsTmp!�������)
                    End If
                Else
                    Do While Not rsTmp.EOF
                        'ȷ����ǰ��ʾ��
                        lngRow = .FindRow(CStr(Split(StrSQL, ",")(i)), , col����)
                        For j = lngRow To .Rows - 1
                            If Val(.TextMatrix(j, col����)) = Val(Split(StrSQL, ",")(i)) Then
                                lngRow = j
                                If .TextMatrix(j, col�������) = "" Then Exit For
                            Else
                                Exit For
                            End If
                        Next
                        If .TextMatrix(lngRow, col�������) <> "" Then
                            lngRow = lngRow + 1: .AddItem "", lngRow
                            .TextMatrix(lngRow, col����) = Split(StrSQL, ",")(i)
                        End If

                        '�ֻ��̶Ⱥ�����������
                        If Val("" & rsTmp!�������) = 3 And Val("" & rsTmp!��ϴ���) = 1 Then
                            If Trim(Nvl(rsTmp!��������)) = "" Then
                                bln�ֻ��̶� = False
                            Else
                                bln�ֻ��̶� = ((InStr("C", UCase(Left(Nvl(rsTmp!��������), 1)))) > 0) Or ((InStr("D0", UCase(Left(Nvl(rsTmp!��������), 2)))) > 0) Or ((InStr("D32.,D33.,", UCase(Left(Nvl(rsTmp!��������), 4)))) > 0)
                            End If
                        End If

                        txtInfo(txt�ֻ��̶�).Enabled = bln�ֻ��̶�
                        lblInfo(txt�ֻ��̶�).Enabled = bln�ֻ��̶�
                        lblInfo(txt����������).Enabled = bln�ֻ��̶�
                        txtInfo(txt����������).Enabled = bln�ֻ��̶�
                        .TextMatrix(lngRow, col�������) = Nvl(rsTmp!�������)
                        .TextMatrix(lngRow, col��ע) = Nvl(rsTmp!��ע)
                        .TextMatrix(lngRow, col��Ժ���) = Nvl(rsTmp!��Ժ���)
                        .TextMatrix(lngRow, col��Ժ����) = Nvl(rsTmp!��Ժ����)
                        .TextMatrix(lngRow, col�Ƿ�δ��) = IIf(Nvl(rsTmp!�Ƿ�δ��, 0) = 1, "��", "")
                        .TextMatrix(lngRow, col�Ƿ�����) = IIf(Nvl(rsTmp!�Ƿ�����, 0) = 1, "��", "")
                        rsTmp.MoveNext
                    Loop
                End If
            Next
        End With
    End If

    vsDiagXY.Cell(flexcpForeColor, 1, col�Ƿ�����, vsDiagXY.Rows - 1, col�Ƿ�����) = vbRed
    vsDiagXY.Cell(flexcpBackColor, GetRow(3), vsDiagXY.FixedRows, GetRow(3), vsDiagXY.Cols - 1) = &HC0FFC0
    vsDiagXY.Row = 1: vsDiagXY.Col = col�������
    If vsDiagXY.TextMatrix(GetRow(6), col�������) <> "" Then
        txtInfo(txt�����).Enabled = True
        txtInfo(txt�����).BackColor = vbWindowBackground
    End If

    '��Ϸ������
    '---------------------------------------------------------------
    StrSQL = "Select ��������,������� From ��Ϸ������ Where ����ID=[1] And ��ҳID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
    Do While Not rsTmp.EOF
        Select Case rsTmp!��������
        Case 1 '�������Ժ
            txtInfo(txt�������Ժ).Text = Decode(Nvl(rsTmp!�������, 0), 1, "����", 2, "������", 3, "���϶�", "")
        Case 2 '��Ժ���Ժ
            txtInfo(txt��Ժ���Ժ).Text = Decode(Nvl(rsTmp!�������, 0), 1, "����", 2, "������", 3, "���϶�", "")
        Case 3 '�����벡��
            txtInfo(txt�����벡��).Text = Decode(Nvl(rsTmp!�������, 0), 1, "����", 2, "������", 3, "���϶�", "")
        Case 4 '�ٴ��벡��
            txtInfo(txt�ٴ��벡��).Text = Decode(Nvl(rsTmp!�������, 0), 1, "����", 2, "������", 3, "���϶�", "")
        Case 5 '�ٴ���ʬ��
            txtInfo(txt�ٴ���ʬ��).Text = Decode(Nvl(rsTmp!�������, 0), 1, "����", 2, "������", 3, "���϶�", "")
        Case 6 '��ǰ������
            txtInfo(txt��ǰ������).Text = Decode(Nvl(rsTmp!�������, 0), 1, "����", 2, "������", 3, "���϶�", "")
        Case 7 '��������Ժ
             txtInfo(txt��������Ժ).Text = Decode(Nvl(rsTmp!�������, 0), 1, "����", 2, "������", 3, "���϶�", "")
        Case 11 '��ҽ�������Ժ
            txtInfo(txt��ҽ�������Ժ).Text = Decode(Nvl(rsTmp!�������, 0), 1, "����", 2, "������", 3, "���϶�", "")
        Case 12 '��ҽ��Ժ���Ժ
            txtInfo(txt��ҽ��Ժ���Ժ).Text = Decode(Nvl(rsTmp!�������, 0), 1, "����", 2, "������", 3, "���϶�", "")
        Case 13 '��ҽ��֤
            txtInfo(txt��֤).Text = Decode(Nvl(rsTmp!�������, 0), 1, "׼ȷ", 2, "����׼ȷ", 3, "�ش�ȱ��", 4, "����", "")
        Case 14 '��ҽ�η�
            txtInfo(txt�η�).Text = Decode(Nvl(rsTmp!�������, 0), 1, "׼ȷ", 2, "����׼ȷ", 3, "�ش�ȱ��", 4, "����", "")
        Case 15 '��ҽ��ҩ
            txtInfo(txt��ҩ).Text = Decode(Nvl(rsTmp!�������, 0), 1, "׼ȷ", 2, "����׼ȷ", 3, "�ش�ȱ��", 4, "����", "")
        End Select
        rsTmp.MoveNext
    Loop

    '��ҽ���
    '---------------------------------------------------------------
    'ȱʡ����ʼ��
    With vsDiagZY
        '11-��ҽ�������;12-��ҽ��Ժ���;13-��ҽ��Ժ���(��Ҫ��ϡ��������)
        .TextMatrix(1, colzy����) = 11
        .TextMatrix(2, colzy����) = 12
        .TextMatrix(3, colzy����) = 13
        .TextMatrix(4, colzy����) = 13
    End With

    If bln��ҳ��� Then
        strTmp = " and a.��¼��Դ=3 "
    Else
        strTmp = " And a.��¼��Դ IN(1,2,3,4) "
    End If

    '��ȡ������Դ�����
    StrSQL = "Select a.��ע, a.Id, a.����id, a.��ҳid, a.ҽ��id, a.��¼��Դ, a.��ϴ���, a.�������, a.����id, a.�������,a.��Ժ����," & _
        " a.����id, a.���id, a.֤��id, a.�������,a.��Ժ���, a.�Ƿ�δ��, a.�Ƿ�����, a.��¼����, a.��¼��, a.ȡ��ʱ��," & _
        " a.ȡ����, a.����id, b.���� As ��������, c.���� As ��ϱ���,d.���� as ֤����� From ������ϼ�¼ A, ��������Ŀ¼ B, �������Ŀ¼ C,��������Ŀ¼ D" & _
        " Where a.����id = b.Id(+) And a.���id = c.Id(+) And a.֤��ID=d.ID(+) And a.������� IN(11,12,13)" & _
        strTmp & _
        " And ȡ��ʱ�� Is Null And ����ID=[1] And ��ҳID=[2]" & _
        " Order by a.�������,a.��¼��Դ Desc,a.��ϴ���,a.�������,a.ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
    
    If Not rsTmp.EOF Then
        With vsDiagZY
            StrSQL = "11,12,13"
            For i = 0 To UBound(Split(StrSQL, ","))
                rsTmp.Filter = "��¼��Դ=3 And �������=" & Split(StrSQL, ",")(i)
                If rsTmp.EOF Then
                    rsTmp.Filter = "��¼��Դ=2 And �������=" & Split(StrSQL, ",")(i)
                End If
                If rsTmp.EOF Then
                    rsTmp.Filter = "��¼��Դ=1 And �������=" & Split(StrSQL, ",")(i)
                End If
                If rsTmp.EOF Then
                    rsTmp.Filter = "��¼��Դ=4 And �������=" & Split(StrSQL, ",")(i)
                End If

                Do While Not rsTmp.EOF
                    'ȷ����ǰ��ʾ��
                    lngRow = .FindRow(CStr(Split(StrSQL, ",")(i)), , colzy����)
                    For j = lngRow To .Rows - 1
                        If Val(.TextMatrix(j, colzy����)) = Val(Split(StrSQL, ",")(i)) Then
                            lngRow = j
                            If .TextMatrix(j, col�������) = "" Then Exit For
                        Else
                            Exit For
                        End If
                    Next
                    If .TextMatrix(lngRow, col�������) <> "" Then
                        lngRow = lngRow + 1: .AddItem "", lngRow
                        .TextMatrix(lngRow, colzy����) = Split(StrSQL, ",")(i)
                    End If
                    .TextMatrix(lngRow, col��ע) = Nvl(rsTmp!��ע)
                    .TextMatrix(lngRow, col�������) = Nvl(rsTmp!�������)
                    .TextMatrix(lngRow, col��Ժ���) = Nvl(rsTmp!��Ժ���)
                    .TextMatrix(lngRow, col��Ժ����) = Nvl(rsTmp!��Ժ����)
                    'ȡ֤������
                    If InStr(.TextMatrix(lngRow, col�������), "(") > 0 And InStr(.TextMatrix(lngRow, col�������), ")") > 0 Then
                        strTmp = Mid(.TextMatrix(lngRow, col�������), InStrRev(.TextMatrix(lngRow, col�������), "(") + 1)
                        strTmp = Mid(strTmp, 1, Len(strTmp) - 1)
                        '��ȡ֤��
                        .TextMatrix(lngRow, col��ҽ֤��) = strTmp
                        'ȥ�����������֤��
                        .TextMatrix(lngRow, col�������) = Mid(.TextMatrix(lngRow, col�������), 1, InStrRev(.TextMatrix(lngRow, col�������), "(") - 1)
                    Else
                       .TextMatrix(lngRow, col��ҽ֤��) = ""
                    End If
                    
                    rsTmp.MoveNext
                Loop
            Next
        End With
    End If
    vsDiagZY.Cell(flexcpBackColor, GetRow(13), vsDiagZY.FixedRows, GetRow(13), vsDiagZY.Cols - 1) = &HC0FFC0
    vsDiagZY.Row = 1: vsDiagZY.Col = col�������

    '������Ϣ:����סԺ��,������
    '---------------------------------------------------------------
    StrSQL = "Select ��¼��Դ,NVL(����ʱ��,��¼ʱ��) as ����ʱ��,ҩ��ID,ҩ����,������Ӧ From ���˹�����¼ A" & _
        " Where ���=1 And ����ID=[1] And ��ҳID=[2]" & _
        " And Not Exists(Select ҩ��ID From ���˹�����¼" & _
            " Where (Nvl(ҩ��ID,0)=Nvl(A.ҩ��ID,0) Or Nvl(ҩ����,'Null')=Nvl(A.ҩ����,'Null'))" & _
            " And Nvl(���,0)=0 And ��¼ʱ��>A.��¼ʱ�� And ����ID=[1] And ��ҳID=[2])" & _
        " Order by NVL(����ʱ��,��¼ʱ��),ҩ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
    If Not rsTmp.EOF Then
        rsTmp.Filter = "��¼��Դ=3" '��ҳ������д��
        If rsTmp.EOF Then rsTmp.Filter = "��¼��Դ<>3" '������Դ����Ϊȱʡ��ʾ
        With vsAller
            .Rows = rsTmp.RecordCount + 1 '�̶���+����
            For i = 1 To rsTmp.RecordCount
                '������Դ�Ŀ������ظ�
                lngRow = -1
                If Not IsNull(rsTmp!ҩ��ID) Then
                    lngRow = .FindRow(CLng(rsTmp!ҩ��ID))
                ElseIf Not IsNull(rsTmp!ҩ����) Then
                    lngRow = .FindRow(CStr(rsTmp!ҩ����), , col����ҩ��)
                End If
                If lngRow = -1 Then
                    .RowData(i) = CLng(Nvl(rsTmp!ҩ��ID, 0))
                    .TextMatrix(i, col����ʱ��) = Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm")
                    .TextMatrix(i, col����ҩ��) = Nvl(rsTmp!ҩ����)
                    .TextMatrix(i, col������Ӧ) = Nvl(rsTmp!������Ӧ)
                End If
                rsTmp.MoveNext
            Next
            If .Rows = .FixedRows Then .Rows = .FixedRows + 1
        End With
    End If
    vsAller.Row = 1: vsAller.Col = col����ҩ��

    '�������
    '---------------------------------------------------------------
    '�׶�ȡ��ҳ�����������
    StrSQL = "Select a.������ʿ,a.�������,a.�п�,a.����,NVl(B.����,C.����) as ��������,a.������ʼʱ��,a.��������ʱ��,a.��������,a.����ҽʦ,a.��һ����,a.�ڶ�����,a.����ʼʱ��,a.��������,a.����ҽʦ,a.ASA�ּ�,a.�ٴ�����,a.NNIS�ּ�,decode(a.��������,1,'һ������',2,'��������',3,'��������',4,'�ļ�����',' ') as ��������,a.׼������,a.������ҩʱ��,a.�пڲ�λ,a.�ط��ƻ�,a.�ط�Ŀ��,a.�пڸ�Ⱦ,a.����֢" & _
            " From ���������¼  A,��������Ŀ¼ B,������ĿĿ¼ C Where c.ID(+)=a.������ĿID And A.��������ID=B.ID(+) and ����ID=[1] And ��ҳID=[2] And ��¼��Դ=3 Order by A.ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
    If rsTmp.EOF Then 'û��ʱ��ȡ������Դ�����
        '��������������ʱ��дȡ��
        StrSQL = "Select Max(��¼����) From ���������¼" & _
            " Where ����ID=" & mlng����ID & " And ��ҳID=" & mlng��ҳID & _
            " And ��¼��Դ=1 And ȡ��ʱ�� is NULL"
         StrSQL = "Select a.������ʿ,a.�������,a.�п�,a.����,a.ȡ����,NVl(B.����,C.����) as ��������,a.������ʼʱ��,a.��������ʱ��,a.��������,a.����ҽʦ,a.��һ����,a.�ڶ�����,a.����ʼʱ��,a.��������,a.����ҽʦ,a.ASA�ּ�,a.�ٴ�����,a.NNIS�ּ�,decode(a.��������,1,'һ������',2,'��������',3,'��������',4,'�ļ�����',' ') as ��������,a.׼������,a.������ҩʱ��,a.�пڲ�λ,a.�ط��ƻ�,a.�ط�Ŀ��,a.�пڸ�Ⱦ,a.����֢" & _
            " From ���������¼  A,��������Ŀ¼ B,������ĿĿ¼ C Where c.ID(+)=a.������ĿID And " & _
            " A.��������ID=B.ID(+) and ����ID=[1] And ��ҳID=[2]" & _
            " And ��¼��Դ=1 And ȡ��ʱ�� is NULL And ��¼����=(" & StrSQL & ")" & _
            " Order by A.ID"
        Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
        If rsTmp.EOF Then '����
            StrSQL = "Select a.������ʿ,a.�������,a.�п�,a.����,NVl(B.����,C.����) as ��������,a.������ʼʱ��,a.��������ʱ��,a.��������,a.����ҽʦ,a.��һ����,a.�ڶ�����,a.����ʼʱ��,a.��������,a.����ҽʦ,a.ASA�ּ�,a.�ٴ�����,a.NNIS�ּ�,decode(a.��������,1,'һ������',2,'��������',3,'��������',4,'�ļ�����',' ') as ��������,a.׼������,a.������ҩʱ��,a.�пڲ�λ,a.�ط��ƻ�,a.�ط�Ŀ��,a.�пڸ�Ⱦ,a.����֢" & _
                " From ���������¼  A,��������Ŀ¼ B,������ĿĿ¼ C Where c.ID(+)=a.������ĿID And  A.��������ID=B.ID(+) and ����ID=[1] And ��ҳID=[2] And ��¼��Դ=4 Order by A.ID"
            Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
        End If
    End If
    If Not rsTmp.EOF Then
        With vsOPS
            .Rows = .FixedRows + rsTmp.RecordCount + 1
            For i = 1 To rsTmp.RecordCount
                .TextMatrix(i, col��ʼ����) = Format(Nvl(rsTmp!������ʼʱ��), "yyyy-MM-dd HH:mm")
                .TextMatrix(i, col��������) = Format(Nvl(rsTmp!��������ʱ��), "yyyy-MM-dd HH:mm")
                .TextMatrix(i, col��������) = Nvl(rsTmp!��������)
                .TextMatrix(i, col��������) = Nvl(rsTmp!��������)
                .TextMatrix(i, col����ҽʦ) = Nvl(rsTmp!����ҽʦ)
                .TextMatrix(i, col������ʿ) = Nvl(rsTmp!������ʿ)
                .TextMatrix(i, col����1) = Nvl(rsTmp!��һ����)
                .TextMatrix(i, col����2) = Nvl(rsTmp!�ڶ�����)
                .TextMatrix(i, col����ҽʦ) = Nvl(rsTmp!����ҽʦ)
                If Not IsNull(rsTmp!�п�) And Not IsNull(rsTmp!����) Then
                    .TextMatrix(i, col�п�����) = rsTmp!�п� & "/" & rsTmp!����
                End If
                .TextMatrix(i, COL�������) = Nvl(rsTmp!�������)
                .TextMatrix(i, col��������) = Nvl(rsTmp!��������)
                .TextMatrix(i, colASA�ּ�) = Nvl(rsTmp!asa�ּ�)
                .TextMatrix(i, colNNIS�ּ�) = Nvl(rsTmp!NNIS�ּ�)
                .TextMatrix(i, col��������) = Nvl(rsTmp!��������)
                .TextMatrix(i, col�ٴ�����) = IIf(Val(rsTmp!�ٴ����� & "") = 1, -1, 0)
                .TextMatrix(i, col׼������) = Nvl(rsTmp!׼������)
                .TextMatrix(i, col������ҩʱ��) = Format(Nvl(rsTmp!������ҩʱ��), "yyyy-MM-dd HH:mm")
                .TextMatrix(i, col����ʼʱ��) = Format(Nvl(rsTmp!����ʼʱ��), "yyyy-MM-dd HH:mm")
                .TextMatrix(i, col�пڲ�λ) = Nvl(rsTmp!�пڲ�λ)
                .TextMatrix(i, col�ط�������Ŀ��) = Nvl(rsTmp!�ط�Ŀ��)
                .TextMatrix(i, col�ط������Ҽƻ�) = IIf(Val(rsTmp!�ط��ƻ� & "") = 1, -1, 0)
                .TextMatrix(i, col�пڸ�Ⱦ) = IIf(Val(rsTmp!�пڸ�Ⱦ & "") = 1, -1, 0)
                .TextMatrix(i, col����֢) = IIf(Val(rsTmp!����֢ & "") = 1, -1, 0)
                rsTmp.MoveNext
            Next
        End With
    End If

    '--------------------------------------------------------------
    '����ҩ��
    StrSQL = "Select a.ҩ��id, a.��ҩĿ��, a.ʹ�ý׶�, a.ʹ������,a.ҩƷ���� ����,һ���п�Ԥ����,DDD��,������ҩ " & vbNewLine & _
            " From ���˿����ؼ�¼ A" & vbNewLine & _
            " Where a.����id = [1] And a.��ҳid = [2] Order By DDD�� Desc"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)

    Do While Not rsTmp.EOF
        With vsKSS
            For j = .FixedRows To .Rows - 1
                If .TextMatrix(j, 1) = "" Then
                    .RowData(j) = Val(rsTmp!ҩ��id & "")
                    If .RowData(j) <> 0 Then
                        .TextMatrix(j, kss����) = Nvl(rsTmp!����)
                        .TextMatrix(j, kss��ҩĿ��) = Nvl(rsTmp!��ҩĿ��)
                        .TextMatrix(j, kssʹ�ý׶�) = Nvl(rsTmp!ʹ�ý׶�)
                        .TextMatrix(j, kssʹ������) = IIf(Val(rsTmp!ʹ������ & "") = 0, "", Val(rsTmp!ʹ������ & "") & "")
                        .Cell(flexcpChecked, j, KSSһ���п�Ԥ����) = Val(rsTmp!һ���п�Ԥ���� & "")
                        .TextMatrix(j, KSSDDD��) = IIf(Val(rsTmp!DDD�� & "") > 0 And Val(rsTmp!DDD�� & "") < 1, "0", "") & Val(rsTmp!DDD�� & "")
                        .TextMatrix(j, KSS������ҩ) = rsTmp!������ҩ & ""
                    End If
                    Exit For
                ElseIf .RowData(j) = Val(rsTmp!ҩ��id & "") Then
                '�ų��ظ�ֵ��������ظ��ģ��򽫺�����е���Ϣ���ϡ�
                    If .RowData(j) <> 0 Then
                        .TextMatrix(j, kss����) = Nvl(rsTmp!����)
                        .TextMatrix(j, kss��ҩĿ��) = Nvl(rsTmp!��ҩĿ��)
                        .TextMatrix(j, kssʹ�ý׶�) = Nvl(rsTmp!ʹ�ý׶�)
                        .TextMatrix(j, kssʹ������) = IIf(Val(rsTmp!ʹ������ & "") = 0, "", Val(rsTmp!ʹ������ & "") & "")
                        .Cell(flexcpChecked, j, KSSһ���п�Ԥ����) = Val(rsTmp!һ���п�Ԥ���� & "")
                        .TextMatrix(j, KSSDDD��) = IIf(Val(rsTmp!DDD�� & "") > 0 And Val(rsTmp!DDD�� & "") < 1, "0", "") & Val(rsTmp!DDD�� & "")
                        .TextMatrix(j, KSS������ҩ) = rsTmp!������ҩ & ""
                    End If
                    Exit For
                End If
            Next
            '���û������û�п����ˣ�������һ��
            If j > .Rows - 1 Then
                .AddItem ""
                .RowData(.Rows - 1) = Val(rsTmp!ҩ��id & "")
                If .RowData(.Rows - 1) <> 0 Then
                    .TextMatrix(.Rows - 1, kss����) = rsTmp!����
                    .TextMatrix(.Rows - 1, kss��ҩĿ��) = Nvl(rsTmp!��ҩĿ��)
                    .TextMatrix(.Rows - 1, kssʹ�ý׶�) = Nvl(rsTmp!ʹ�ý׶�)
                    .TextMatrix(.Rows - 1, kssʹ������) = IIf(Val(rsTmp!ʹ������ & "") = 0, "", Val(rsTmp!ʹ������ & "") & "")
                    .Cell(flexcpChecked, .Rows - 1, KSSһ���п�Ԥ����) = Val(rsTmp!һ���п�Ԥ���� & "")
                    .TextMatrix(.Rows - 1, KSSDDD��) = IIf(Val(rsTmp!DDD�� & "") > 0 And Val(rsTmp!DDD�� & "") < 1, "0", "") & Val(rsTmp!DDD�� & "")
                    .TextMatrix(.Rows - 1, KSS������ҩ) = rsTmp!������ҩ & ""
                End If
            End If
        End With
        rsTmp.MoveNext
    Loop
    
    If mbln���� Then
        '���ƻ���
        Call Load���������(mlng����ID, mlng��ҳID)
    End If
    Call Load��ҳ����(mlng����ID, mlng��ҳID)
    
    '������Ϣ
    '---------------------------------------------------------------
    '�����¼�
    lstAdvEvent.Clear
    
    
    blnѹ�� = False: bln����׹�� = False
    StrSQL = "Select ����, ����" & vbNewLine & _
            "From �����¼� A," & vbNewLine & _
            "     (Select Decode(��Ϣֵ, Null, Null, ',' || ��Ϣֵ || ',') ��Ϣֵ" & vbNewLine & _
            "       From ������ҳ�ӱ�" & vbNewLine & _
            "       Where ����id = [1] And ��ҳid = [2] And ��Ϣ�� = '�����¼�') B" & vbNewLine & _
            "Where Instr(b.��Ϣֵ , chr(44)|| a.���� ||chr(44) ) > 0"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
    For i = 1 To rsTmp.RecordCount
        lstAdvEvent.AddItem Nvl(rsTmp!����)
        If Nvl(rsTmp!����) = "ѹ��" Then
            blnѹ�� = True
        ElseIf Nvl(rsTmp!����) = "ҽԺ�ڵ���/׹��" Then 'ѹ�� ����׹��
            bln����׹�� = True
        End If
        rsTmp.MoveNext
    Next

    txtInfo(txtѹ�������ڼ�).Enabled = blnѹ��
    txtInfo(txtѹ������).Enabled = blnѹ��
    lblInfo(txtѹ�������ڼ�).Enabled = blnѹ��
    lblInfo(txtѹ������).Enabled = blnѹ��

    txtInfo(txt������׹��ԭ��).Enabled = bln����׹��
    txtInfo(txt������׹���˺�).Enabled = bln����׹��
    lblInfo(txt������׹��ԭ��).Enabled = bln����׹��
    lblInfo(txt������׹���˺�).Enabled = bln����׹��
    
    mblnCheck = False
    Screen.MousePointer = 0
    LoadPageData = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Load���������(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:���ط����뻯����Ϣ
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-10-21 15:55:27
    '����:13999
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long
    Dim StrSQL As String
    
    Err = 0: On Error GoTo Errhand:
    StrSQL = " " & _
    "   Select A.����id, A.��ҳid, A.���, A.����id, A.��ʼ����, A.��������, A.�Ƴ���, A.����, A.���Ʒ���, A.����Ч��, " & _
    "          B.���� || '-' || B.���� As ������Ϣ " & _
    "   From �������Ƽ�¼ A, ��������Ŀ¼ B " & _
    "   Where A.����id = B.Id And a.����id=[1] And a.��ҳid=[2] " & _
    "   Order By ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, lng����ID, lng��ҳID)
    With vsChemotherapy
        .Rows = 2
        .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
        .Clear 1
        If rsTemp.RecordCount <> 0 Then .Rows = rsTemp.RecordCount + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("��ѧ���Ʊ���")) = Nvl(rsTemp!������Ϣ)
            .TextMatrix(lngRow, .ColIndex("��ʼ����")) = Format(rsTemp!��ʼ����, "yyyy-MM-DD")
            .TextMatrix(lngRow, .ColIndex("��������")) = Format(rsTemp!��������, "yyyy-MM-DD")
            .TextMatrix(lngRow, .ColIndex("�Ƴ���")) = Format(Val(Nvl(rsTemp!�Ƴ���)), "###;-###;;")
            .TextMatrix(lngRow, .ColIndex("����")) = Format(Val(Nvl(rsTemp!����)), "###;-###;;")
            .TextMatrix(lngRow, .ColIndex("���Ʒ���")) = Trim(Nvl(rsTemp!���Ʒ���))
            .TextMatrix(lngRow, .ColIndex("����Ч��")) = Trim(Nvl(rsTemp!����Ч��))
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    StrSQL = " " & _
    "   Select A.����id, A.��ҳid, A.���, A.����id, A.��ʼ����, A.��������,A.��Ұ��λ, A.�������, A.�ۼ���, A.����Ч��, " & _
    "          B.���� || '-' || B.���� As ������Ϣ " & _
    "   From �������Ƽ�¼ A, ��������Ŀ¼ B " & _
    "   Where A.����id = B.Id And a.����id=[1] And a.��ҳid=[2] " & _
    "   Order By ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, lng����ID, lng��ҳID)
    With vsRadiotherapy
        .Rows = 2
        .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
        .Clear 1
        If rsTemp.RecordCount <> 0 Then .Rows = rsTemp.RecordCount + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("�������Ʊ���")) = Nvl(rsTemp!������Ϣ)
            .TextMatrix(lngRow, .ColIndex("��ʼ����")) = Format(rsTemp!��ʼ����, "yyyy-MM-DD")
            .TextMatrix(lngRow, .ColIndex("��������")) = Format(rsTemp!��������, "yyyy-MM-DD")
            .TextMatrix(lngRow, .ColIndex("�������")) = Format(Val(Nvl(rsTemp!�������)), "###;-###;;")
            .TextMatrix(lngRow, .ColIndex("�ۼ���")) = Format(Val(Nvl(rsTemp!�ۼ���)), "###;-###;;")
            .TextMatrix(lngRow, .ColIndex("��Ұ��λ")) = Trim(Nvl(rsTemp!��Ұ��λ))
            .TextMatrix(lngRow, .ColIndex("����Ч��")) = Trim(Nvl(rsTemp!����Ч��))
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    Load��������� = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function Load��ҳ����(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    '-------------------------------------------------------------------------------------------------------------------------
    '����:���ظ�ҳ����
    '����:lng����id-����id
    '     lng��ҳid -��ҳid
    '����:���سɹ�,����true,���򷵻�False
    '-------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long
    Dim StrSQL As String
    Dim i As Long
    
    Err = 0: On Error GoTo Errhand:
     
    
    '����֢���
    StrSQL = "" & _
        "Select ���, b.���� || '.' || B.���� As �໤������, To_Char(����ʱ��, 'yyyy-mm-dd HH24:mi') As ����ʱ��, To_Char(�˳�ʱ��, 'yyyy-mm-dd HH24:mi') As �˳�ʱ��, ����ס�ƻ�," & vbNewLine & _
        "       ����סԭ��" & vbNewLine & _
        "From ������֢�໤��� a, Icu���� b" & vbNewLine & _
        "Where a.�໤������ = b.����(+) And ����id = [1] And ��ҳid = [2]" & vbNewLine & _
        "Order By ���"

    Set rsTemp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, lng����ID, lng��ҳID)
    With vsFlxAddICU
        If rsTemp.RecordCount <> 0 Then
            .Rows = rsTemp.RecordCount + 1
        Else
            .Rows = 2
        End If
        lngRow = 1
        Do While Not rsTemp.EOF
            .Cell(flexcpText, lngRow, 0) = Nvl(rsTemp!���)
            .Cell(flexcpText, lngRow, 1) = Nvl(rsTemp!�໤������)
            .Cell(flexcpText, lngRow, 2) = Nvl(rsTemp!����ʱ��)
            .Cell(flexcpText, lngRow, 3) = Nvl(rsTemp!�˳�ʱ��)
            .Cell(flexcpChecked, lngRow, 4) = Nvl(rsTemp!����ס�ƻ�)
            .Cell(flexcpText, lngRow, 5) = Nvl(rsTemp!����סԭ��)
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    For i = 1 To vsFlxAddICU.Rows - 1
        vsFlxAddICU.TextMatrix(i, 0) = i
    Next
    '��е�������
    StrSQL = "" & _
        "Select a.���,��� ||'-' ||a.�໤������ As �໤������,C.����||'.'||C.���� As ��е������, To_Char(��ʼʹ��ʱ��, 'yyyy-mm-dd HH24:mi') As ��ʼʹ��ʱ��, To_Char(����ʹ��ʱ��, 'yyyy-mm-dd HH24:mi') As ����ʹ��ʱ��," & vbNewLine & _
        "       ��Ⱦ�ۼ�ʱ��" & vbNewLine & _
        "From ��е����ʹ����� a, ��е����Ŀ¼ c" & vbNewLine & _
        "Where a.��е������ = c.����(+) And ����id = [1] And ��ҳid = [2]" & vbNewLine & _
        "Order By ���"

    Set rsTemp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, lng����ID, lng��ҳID)
    With vsICUInstruments
        If rsTemp.RecordCount <> 0 Then
            .Rows = rsTemp.RecordCount + 1
        Else
            .Rows = 2
        End If
        lngRow = 1
        Do While Not rsTemp.EOF
            .Cell(flexcpText, lngRow, 0) = Nvl(rsTemp!�໤������)
            .Cell(flexcpText, lngRow, 1) = Nvl(rsTemp!��е������)
            .Cell(flexcpText, lngRow, 2) = Nvl(rsTemp!��ʼʹ��ʱ��)
            .Cell(flexcpText, lngRow, 3) = Nvl(rsTemp!����ʹ��ʱ��)
            .Cell(flexcpText, lngRow, 4) = Nvl(rsTemp!��Ⱦ�ۼ�ʱ��)
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    
    ' ���ظ�Ⱦ���
    StrSQL = "" & _
            "Select To_Char(a.ȷ������, 'yyyy-mm-dd') As ȷ������, b.���� || '.' || a.��Ⱦ��λ As ��Ⱦ��λ, a.��Ⱦ����, c.����" & vbNewLine & _
            " From ���˸�Ⱦ��¼ a, ��Ⱦ��λ b, ҽԺ��ȾĿ¼ c" & vbNewLine & _
            " Where a.��Ⱦ��λ = b.����(+) And a.��Ⱦ���� = c.����(+) And a.����id = [1] And a.��ҳid = [2]" & _
            " Order By a.���"
    Set rsTemp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, lng����ID, lng��ҳID)
    With vsInfect
        If rsTemp.RecordCount <> 0 Then
            .Rows = rsTemp.RecordCount + 1
        Else
            .Rows = 2
        End If
        lngRow = 1
        Do While Not rsTemp.EOF
            .Cell(flexcpText, lngRow, 0) = Nvl(rsTemp!ȷ������)
            .Cell(flexcpText, lngRow, 1) = Nvl(rsTemp!��Ⱦ��λ)
            .Cell(flexcpText, lngRow, 2) = Nvl(rsTemp!����)
            .RowData(lngRow) = Nvl(rsTemp!��Ⱦ����)
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    
    ' ���ر걾��Դ
    StrSQL = "" & _
            "Select a.�걾, a.��ԭѧ���� || '-' || b.���� As ��ԭѧ����, To_Char(a.�ͼ�����, 'yyyy-mm-dd') As �ͼ�����" & vbNewLine & _
            " From ���˲�ԭѧ��� a, ��ԭѧĿ¼ b" & vbNewLine & _
            " Where a.��ԭѧ���� = b.����(+) And a.����id = [1] And a.��ҳid = [2]" & _
            " Order By a.���"

    Set rsTemp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, lng����ID, lng��ҳID)
    With vsSample
        If rsTemp.RecordCount <> 0 Then
            .Rows = rsTemp.RecordCount + 1
        Else
            .Rows = 2
        End If
        lngRow = 1
        Do While Not rsTemp.EOF
            .Cell(flexcpText, lngRow, 0) = Decode(Nvl(rsTemp!�걾), "1", "1.ѪҺ", "2", "2.��Һ", "3", "3.���", "4", "4.̵Һ", "5", "5.����������")
            .Cell(flexcpText, lngRow, 1) = Nvl(rsTemp!��ԭѧ����)
            .Cell(flexcpText, lngRow, 2) = Nvl(rsTemp!�ͼ�����)
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With

    
    Load��ҳ���� = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub FillVsf()
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long
    Dim lngCol As Long
    Dim StrSQL As String
    
    On Error GoTo errH
    StrSQL = "select ����,���� from ������Ŀ order by ����"
    vsfMain.Clear
    
    Call zlDatabase.OpenRecordset(rsTemp, StrSQL, Me.Caption)
    If rsTemp.RecordCount = 0 Then vsfMain.Rows = 1: vsfMain.Cols = 1: Exit Sub
    If (rsTemp.RecordCount Mod 2) <> 0 Then
        vsfMain.Rows = rsTemp.RecordCount \ 2 + 2
    Else
        vsfMain.Rows = rsTemp.RecordCount \ 2 + 1
    End If
    With vsfMain
        .Cols = 6
        For lngRow = 0 To 3 Step 3
            .TextMatrix(0, lngRow) = "��Ŀ"
            .TextMatrix(0, lngRow + 1) = "����"
            .ColWidth(0 + lngRow) = 1500
            .ColWidth(1 + lngRow) = 1200
            .ColWidth(2 + lngRow) = 0
        Next lngRow
        .Cell(flexcpAlignment, 0, 0, 0, vsfMain.Cols - 1) = 4
        .Cell(flexcpBackColor, 1, 0, .Rows - 1, 0) = &HFCE7D8
        .Cell(flexcpBackColor, 1, 3, .Rows - 1, 3) = &HFCE7D8
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignCenterCenter
    End With
    lngRow = 1
    lngCol = 0
    While Not rsTemp.EOF
        If lngCol < 4 Then
            With vsfMain
                .TextMatrix(lngRow, lngCol + 0) = rsTemp!����
                .TextMatrix(lngRow, lngCol + 2) = rsTemp!���� & ""
                If InStr(rsTemp!����, "�Ƿ�") > 0 Then
                    vsfMain.TextMatrix(lngRow, lngCol + 1) = "��"
                    vsfMain.Cell(flexcpChecked, lngRow, lngCol + 1) = 2
                End If
            End With
            lngCol = lngCol + 3
            rsTemp.MoveNext
        Else
            lngCol = 0
            lngRow = lngRow + 1
        End If
    Wend
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetIDTmp(ByVal strName As String) As Long
'���ܣ��������ڽ�������ҳ�ӱ�Ŀ����� �Ƶ����±� ���˿����ؼ�¼�У���ǰû�м�¼ҩƷid�����ڸ������ƽ�id�����
    Dim rsTmp As Recordset, StrSQL As String
    
    On Error GoTo errH
    StrSQL = "Select Distinct a.Id" & vbNewLine & _
                "From ������ĿĿ¼ A, ������Ŀ���� B, ҩƷ���� C" & vbNewLine & _
                "Where a.Id = b.������Ŀid And a.Id = c.ҩ��id And Nvl(c.������, 0) <> 0 And A.����=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, strName)
    If rsTmp.RecordCount > 0 Then
        GetIDTmp = Val(rsTmp!ID)
    Else
        GetIDTmp = 0
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


