VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Begin VB.Form frmItemDeliveryEdit 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "#"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9675
   Icon            =   "frmItemDeliveryEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   9675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   8325
      Index           =   0
      Left            =   0
      ScaleHeight     =   8325
      ScaleWidth      =   9630
      TabIndex        =   10
      Top             =   390
      Width           =   9630
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   3
         Left            =   4095
         TabIndex        =   1
         Top             =   90
         Width           =   5445
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   2
         Left            =   705
         TabIndex        =   20
         Top             =   90
         Width           =   2580
      End
      Begin VB.Frame fra 
         Height          =   7320
         Index           =   0
         Left            =   45
         TabIndex        =   11
         Top             =   375
         Width           =   4980
         Begin VB.PictureBox picKind 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   945
            Index           =   1
            Left            =   2985
            ScaleHeight     =   915
            ScaleWidth      =   2085
            TabIndex        =   21
            Top             =   5805
            Visible         =   0   'False
            Width           =   2115
            Begin VB.TextBox txt 
               BorderStyle     =   0  'None
               Height          =   825
               Index           =   4
               Left            =   105
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   22
               Top             =   60
               Width           =   1110
            End
         End
         Begin VB.PictureBox picKind 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   6315
            Index           =   0
            Left            =   75
            ScaleHeight     =   6285
            ScaleWidth      =   4770
            TabIndex        =   23
            Top             =   525
            Visible         =   0   'False
            Width           =   4800
            Begin VB.CommandButton cmdExpand 
               Caption         =   "+"
               Height          =   345
               Left            =   3945
               TabIndex        =   45
               ToolTipText     =   "ȫ��չ��"
               Top             =   360
               Width           =   375
            End
            Begin VB.CommandButton cmdCollapse 
               Caption         =   "-"
               Height          =   345
               Left            =   4365
               TabIndex        =   44
               ToolTipText     =   "ȫ������"
               Top             =   360
               Width           =   375
            End
            Begin VB.OptionButton opt 
               Caption         =   "��Ʒ����"
               Height          =   195
               Index           =   0
               Left            =   855
               TabIndex        =   42
               Top             =   90
               Value           =   -1  'True
               Width           =   1050
            End
            Begin VB.OptionButton opt 
               Caption         =   "��Ϣ����"
               Height          =   195
               Index           =   1
               Left            =   2025
               TabIndex        =   41
               Top             =   90
               Width           =   1155
            End
            Begin VB.TextBox txt 
               Height          =   300
               Index           =   0
               Left            =   855
               TabIndex        =   39
               Top             =   390
               Width           =   2625
            End
            Begin VB.CommandButton cmdFind 
               Height          =   330
               Left            =   3510
               Picture         =   "frmItemDeliveryEdit.frx":000C
               Style           =   1  'Graphical
               TabIndex        =   38
               ToolTipText     =   "������Ĳ������ݽ��в��Ҷ�λ"
               Top             =   375
               Width           =   375
            End
            Begin VB.PictureBox picBack 
               Appearance      =   0  'Flat
               ForeColor       =   &H80000008&
               Height          =   945
               Index           =   2
               Left            =   2085
               ScaleHeight     =   915
               ScaleWidth      =   2085
               TabIndex        =   34
               Top             =   4755
               Visible         =   0   'False
               Width           =   2115
               Begin VSFlex8Ctl.VSFlexGrid vsf 
                  Height          =   750
                  Index           =   2
                  Left            =   60
                  TabIndex        =   35
                  Top             =   30
                  Width           =   1755
                  _cx             =   3096
                  _cy             =   1323
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
                  BackColorSel    =   16772055
                  ForeColorSel    =   0
                  BackColorBkg    =   -2147483643
                  BackColorAlternate=   -2147483643
                  GridColor       =   -2147483638
                  GridColorFixed  =   -2147483632
                  TreeColor       =   -2147483638
                  FloodColor      =   192
                  SheetBorder     =   -2147483643
                  FocusRect       =   1
                  HighLight       =   1
                  AllowSelection  =   0   'False
                  AllowBigSelection=   0   'False
                  AllowUserResizing=   1
                  SelectionMode   =   1
                  GridLines       =   8
                  GridLinesFixed  =   8
                  GridLineWidth   =   1
                  Rows            =   2
                  Cols            =   5
                  FixedRows       =   1
                  FixedCols       =   0
                  RowHeightMin    =   270
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
                  AutoSizeMode    =   1
                  AutoSearch      =   0
                  AutoSearchDelay =   2
                  MultiTotals     =   -1  'True
                  SubtotalPosition=   1
                  OutlineBar      =   6
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
            Begin VB.PictureBox picBack 
               Appearance      =   0  'Flat
               ForeColor       =   &H80000008&
               Height          =   1050
               Index           =   3
               Left            =   120
               ScaleHeight     =   1020
               ScaleWidth      =   2115
               TabIndex        =   32
               Top             =   3465
               Visible         =   0   'False
               Width           =   2145
               Begin VSFlex8Ctl.VSFlexGrid vsf 
                  Height          =   960
                  Index           =   3
                  Left            =   60
                  TabIndex        =   33
                  Top             =   30
                  Width           =   1725
                  _cx             =   3043
                  _cy             =   1693
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
                  BackColorSel    =   16772055
                  ForeColorSel    =   0
                  BackColorBkg    =   -2147483643
                  BackColorAlternate=   -2147483643
                  GridColor       =   -2147483638
                  GridColorFixed  =   -2147483632
                  TreeColor       =   -2147483638
                  FloodColor      =   192
                  SheetBorder     =   -2147483643
                  FocusRect       =   1
                  HighLight       =   1
                  AllowSelection  =   0   'False
                  AllowBigSelection=   0   'False
                  AllowUserResizing=   1
                  SelectionMode   =   1
                  GridLines       =   8
                  GridLinesFixed  =   8
                  GridLineWidth   =   1
                  Rows            =   2
                  Cols            =   5
                  FixedRows       =   1
                  FixedCols       =   0
                  RowHeightMin    =   270
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
                  AutoSizeMode    =   1
                  AutoSearch      =   0
                  AutoSearchDelay =   2
                  MultiTotals     =   -1  'True
                  SubtotalPosition=   1
                  OutlineBar      =   6
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
            Begin VB.PictureBox picBack 
               Appearance      =   0  'Flat
               ForeColor       =   &H80000008&
               Height          =   1125
               Index           =   4
               Left            =   90
               ScaleHeight     =   1095
               ScaleWidth      =   2205
               TabIndex        =   30
               Top             =   4650
               Visible         =   0   'False
               Width           =   2235
               Begin VSFlex8Ctl.VSFlexGrid vsf 
                  Height          =   960
                  Index           =   4
                  Left            =   60
                  TabIndex        =   31
                  Top             =   30
                  Width           =   1890
                  _cx             =   3334
                  _cy             =   1693
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
                  BackColorSel    =   16772055
                  ForeColorSel    =   0
                  BackColorBkg    =   -2147483643
                  BackColorAlternate=   -2147483643
                  GridColor       =   -2147483638
                  GridColorFixed  =   -2147483632
                  TreeColor       =   -2147483638
                  FloodColor      =   192
                  SheetBorder     =   -2147483643
                  FocusRect       =   1
                  HighLight       =   1
                  AllowSelection  =   0   'False
                  AllowBigSelection=   0   'False
                  AllowUserResizing=   1
                  SelectionMode   =   1
                  GridLines       =   8
                  GridLinesFixed  =   8
                  GridLineWidth   =   1
                  Rows            =   2
                  Cols            =   5
                  FixedRows       =   1
                  FixedCols       =   0
                  RowHeightMin    =   270
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
                  AutoSizeMode    =   1
                  AutoSearch      =   0
                  AutoSearchDelay =   2
                  MultiTotals     =   -1  'True
                  SubtotalPosition=   1
                  OutlineBar      =   6
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
            Begin VB.PictureBox picBack 
               Appearance      =   0  'Flat
               ForeColor       =   &H80000008&
               Height          =   1125
               Index           =   5
               Left            =   2145
               ScaleHeight     =   1095
               ScaleWidth      =   2205
               TabIndex        =   28
               Top             =   1035
               Visible         =   0   'False
               Width           =   2235
               Begin VSFlex8Ctl.VSFlexGrid vsf 
                  Height          =   960
                  Index           =   5
                  Left            =   60
                  TabIndex        =   29
                  Top             =   15
                  Width           =   1890
                  _cx             =   3334
                  _cy             =   1693
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
                  BackColorSel    =   16772055
                  ForeColorSel    =   0
                  BackColorBkg    =   -2147483643
                  BackColorAlternate=   -2147483643
                  GridColor       =   -2147483638
                  GridColorFixed  =   -2147483632
                  TreeColor       =   -2147483638
                  FloodColor      =   192
                  SheetBorder     =   -2147483643
                  FocusRect       =   1
                  HighLight       =   1
                  AllowSelection  =   0   'False
                  AllowBigSelection=   0   'False
                  AllowUserResizing=   1
                  SelectionMode   =   1
                  GridLines       =   8
                  GridLinesFixed  =   8
                  GridLineWidth   =   1
                  Rows            =   2
                  Cols            =   5
                  FixedRows       =   1
                  FixedCols       =   0
                  RowHeightMin    =   270
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
                  AutoSizeMode    =   1
                  AutoSearch      =   0
                  AutoSearchDelay =   2
                  MultiTotals     =   -1  'True
                  SubtotalPosition=   1
                  OutlineBar      =   6
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
            Begin VB.PictureBox picBack 
               Appearance      =   0  'Flat
               ForeColor       =   &H80000008&
               Height          =   1125
               Index           =   6
               Left            =   2175
               ScaleHeight     =   1095
               ScaleWidth      =   2205
               TabIndex        =   26
               Top             =   2250
               Visible         =   0   'False
               Width           =   2235
               Begin VSFlex8Ctl.VSFlexGrid vsf 
                  Height          =   960
                  Index           =   6
                  Left            =   60
                  TabIndex        =   27
                  Top             =   30
                  Width           =   1890
                  _cx             =   3334
                  _cy             =   1693
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
                  BackColorSel    =   16772055
                  ForeColorSel    =   0
                  BackColorBkg    =   -2147483643
                  BackColorAlternate=   -2147483643
                  GridColor       =   -2147483638
                  GridColorFixed  =   -2147483632
                  TreeColor       =   -2147483638
                  FloodColor      =   192
                  SheetBorder     =   -2147483643
                  FocusRect       =   1
                  HighLight       =   1
                  AllowSelection  =   0   'False
                  AllowBigSelection=   0   'False
                  AllowUserResizing=   1
                  SelectionMode   =   1
                  GridLines       =   8
                  GridLinesFixed  =   8
                  GridLineWidth   =   1
                  Rows            =   2
                  Cols            =   5
                  FixedRows       =   1
                  FixedCols       =   0
                  RowHeightMin    =   270
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
                  AutoSizeMode    =   1
                  AutoSearch      =   0
                  AutoSearchDelay =   2
                  MultiTotals     =   -1  'True
                  SubtotalPosition=   1
                  OutlineBar      =   6
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
            Begin VB.PictureBox picBack 
               Appearance      =   0  'Flat
               ForeColor       =   &H80000008&
               Height          =   1125
               Index           =   7
               Left            =   2190
               ScaleHeight     =   1095
               ScaleWidth      =   2205
               TabIndex        =   24
               Top             =   3555
               Visible         =   0   'False
               Width           =   2235
               Begin VSFlex8Ctl.VSFlexGrid vsf 
                  Height          =   960
                  Index           =   7
                  Left            =   60
                  TabIndex        =   25
                  Top             =   30
                  Width           =   1890
                  _cx             =   3334
                  _cy             =   1693
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
                  BackColorSel    =   16772055
                  ForeColorSel    =   0
                  BackColorBkg    =   -2147483643
                  BackColorAlternate=   -2147483643
                  GridColor       =   -2147483638
                  GridColorFixed  =   -2147483632
                  TreeColor       =   -2147483638
                  FloodColor      =   192
                  SheetBorder     =   -2147483643
                  FocusRect       =   1
                  HighLight       =   1
                  AllowSelection  =   0   'False
                  AllowBigSelection=   0   'False
                  AllowUserResizing=   1
                  SelectionMode   =   1
                  GridLines       =   8
                  GridLinesFixed  =   8
                  GridLineWidth   =   1
                  Rows            =   2
                  Cols            =   5
                  FixedRows       =   1
                  FixedCols       =   0
                  RowHeightMin    =   270
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
                  AutoSizeMode    =   1
                  AutoSearch      =   0
                  AutoSearchDelay =   2
                  MultiTotals     =   -1  'True
                  SubtotalPosition=   1
                  OutlineBar      =   6
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
            Begin VB.PictureBox picBack 
               Appearance      =   0  'Flat
               ForeColor       =   &H80000008&
               Height          =   5550
               Index           =   1
               Left            =   30
               ScaleHeight     =   5520
               ScaleWidth      =   4695
               TabIndex        =   36
               Top             =   750
               Width           =   4725
               Begin VSFlex8Ctl.VSFlexGrid vsf 
                  Height          =   2130
                  Index           =   1
                  Left            =   120
                  TabIndex        =   37
                  Top             =   15
                  Width           =   1290
                  _cx             =   2275
                  _cy             =   3757
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
                  BackColorSel    =   16772055
                  ForeColorSel    =   0
                  BackColorBkg    =   -2147483643
                  BackColorAlternate=   -2147483643
                  GridColor       =   -2147483638
                  GridColorFixed  =   -2147483632
                  TreeColor       =   -2147483638
                  FloodColor      =   192
                  SheetBorder     =   -2147483643
                  FocusRect       =   1
                  HighLight       =   1
                  AllowSelection  =   0   'False
                  AllowBigSelection=   0   'False
                  AllowUserResizing=   1
                  SelectionMode   =   1
                  GridLines       =   8
                  GridLinesFixed  =   8
                  GridLineWidth   =   1
                  Rows            =   2
                  Cols            =   5
                  FixedRows       =   1
                  FixedCols       =   0
                  RowHeightMin    =   270
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
                  AutoSizeMode    =   1
                  AutoSearch      =   0
                  AutoSearchDelay =   2
                  MultiTotals     =   -1  'True
                  SubtotalPosition=   1
                  OutlineBar      =   6
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
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "������Դ"
               Height          =   180
               Index           =   1
               Left            =   45
               TabIndex        =   43
               Top             =   90
               Width           =   720
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��λ����"
               Height          =   180
               Index           =   3
               Left            =   30
               TabIndex        =   40
               Top             =   435
               Width           =   720
            End
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   0
            Left            =   900
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   180
            Width           =   4005
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "��� >"
            Height          =   345
            Left            =   3795
            TabIndex        =   4
            ToolTipText     =   "��ӵ�ǰѡ�е�Ŀ������"
            Top             =   6885
            Width           =   1100
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ͷ�ݶ���"
            Height          =   180
            Index           =   2
            Left            =   105
            TabIndex        =   2
            Top             =   240
            Width           =   720
         End
      End
      Begin VB.Frame fra 
         Height          =   7320
         Index           =   1
         Left            =   5055
         TabIndex        =   12
         Top             =   375
         Width           =   4515
         Begin VB.CommandButton cmdFindSel 
            Height          =   345
            Left            =   3105
            Picture         =   "frmItemDeliveryEdit.frx":685E
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "������Ĳ������ݽ��в��Ҷ�λ"
            Top             =   585
            Width           =   390
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   1
            Left            =   855
            TabIndex        =   17
            Top             =   630
            Width           =   2235
         End
         Begin VB.CommandButton cmdExpandSel 
            Caption         =   "+"
            Height          =   345
            Left            =   3600
            TabIndex        =   15
            ToolTipText     =   "ȫ��չ��"
            Top             =   600
            Width           =   375
         End
         Begin VB.CommandButton cmdCollapseSel 
            Caption         =   "-"
            Height          =   345
            Left            =   4035
            TabIndex        =   14
            ToolTipText     =   "ȫ������"
            Top             =   600
            Width           =   375
         End
         Begin VB.CommandButton cmdRemove 
            Caption         =   "< �Ƴ�"
            Height          =   350
            Left            =   75
            TabIndex        =   6
            ToolTipText     =   "ɾ����ǰѡ�е�Ŀ������"
            Top             =   6900
            Width           =   1100
         End
         Begin VSFlex8Ctl.VSFlexGrid vsf 
            Height          =   5835
            Index           =   0
            Left            =   75
            TabIndex        =   5
            Top             =   1005
            Width           =   4335
            _cx             =   7646
            _cy             =   10292
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
            BackColorSel    =   16772055
            ForeColorSel    =   0
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483638
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483638
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   8
            GridLinesFixed  =   8
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   270
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
            AutoSizeMode    =   1
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   6
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
         Begin VB.Image img 
            Height          =   480
            Index           =   0
            Left            =   90
            Picture         =   "frmItemDeliveryEdit.frx":D0B0
            Top             =   60
            Width           =   480
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��λ����"
            Height          =   180
            Index           =   5
            Left            =   90
            TabIndex        =   16
            Top             =   675
            Width           =   720
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "��ѡ���Ŀ�꣬���Ŀ��ʱ��Ҫͬʱ���㡣ͬһĿ��Ϊ���й�ϵ"
            Height          =   405
            Index           =   4
            Left            =   630
            TabIndex        =   13
            Top             =   165
            Width           =   3810
         End
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   8445
         TabIndex        =   8
         Top             =   7860
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   7260
         TabIndex        =   7
         Top             =   7845
         Width           =   1100
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ʶ(&C)"
         Height          =   180
         Index           =   6
         Left            =   60
         TabIndex        =   19
         Top             =   150
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����(&N)"
         Height          =   180
         Index           =   0
         Left            =   3390
         TabIndex        =   0
         Top             =   150
         Width           =   630
      End
   End
   Begin VB.TextBox txtLocation 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   2370
      TabIndex        =   9
      Top             =   75
      Width           =   1575
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
Attribute VB_Name = "frmItemDeliveryEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################

Private mfrmParent As Object
Private mbytMode As Byte
Private mblnDataChanged As Boolean
Private mblnReading As Boolean
Private mobjFindKey As CommandBarControl
Private mstrFindKey As String
Private mrsPara As ADODB.Recordset
Private mstrDataKey As String
Private mstrItemDataKey As String
Private mlngModualCode As Long
Private mblnContiune As Boolean
Private mclsVsf(7) As zlVSFlexGrid.clsVsf
Private mrsSelTarget As ADODB.Recordset
Private mblnOutline(13) As Boolean
Public Event AfterNewData(ByVal DataKey As String)
Public Event AfterModifyData(ByVal DataKey As String)
Public Event AfterDeleteData(ByVal DataKey As String)
Public Event Forward(ByRef DataKey As String, ByRef Cancel As Boolean)
Public Event Backward(ByRef DataKey As String, ByRef Cancel As Boolean)

'######################################################################################################################

Public Function InitDialog(ByVal frmParent As Object, Optional ByVal lngModualCode As Long) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Set mfrmParent = frmParent
    mlngModualCode = lngModualCode
    InitDialog = True
    
End Function

Public Sub NewData(ByVal strItemDataKey As String)
    '******************************************************************************************************************
    '���ܣ�����Ͷ��Ŀ��
    '������
    '���أ�
    '******************************************************************************************************************
    mbytMode = 1
    Me.Caption = "����Ͷ��Ŀ��"
    mstrItemDataKey = strItemDataKey
    mstrDataKey = ""
    
    Call InitGrid
    Call InitData
    Call InitCommandBar
    
    mblnDataChanged = False
        
    Me.Show 1, mfrmParent
    
End Sub

Public Sub ModifyData(ByVal strItemDataKey As String, ByVal strDataKey As String)
    '******************************************************************************************************************
    '���ܣ��޸�Ͷ��Ŀ��
    '������
    '���أ�
    '******************************************************************************************************************
    mbytMode = 2
    mstrDataKey = strDataKey
    mstrItemDataKey = strItemDataKey
    Me.Caption = "�޸�Ͷ��Ŀ��"
    
    Call InitGrid
    Call InitData
    Call InitCommandBar
    
    Call ReadData(mstrDataKey)
    
    mblnDataChanged = False
    
    Me.Show 1, mfrmParent
    
End Sub

Public Sub ViewData(ByVal strItemDataKey As String, ByVal strDataKey As String)
    '******************************************************************************************************************
    '���ܣ�����Ͷ��Ŀ��
    '������
    '���أ�
    '******************************************************************************************************************
    mbytMode = 0
    mstrDataKey = strDataKey
    mstrItemDataKey = strItemDataKey
    Me.Caption = "����Ͷ��Ŀ��"
    
    Call InitGrid
    Call InitData
    Call InitCommandBar
    
    cmdOK.Enabled = False
    txt(2).Enabled = False
    txt(3).Enabled = False
    cmdAdd.Enabled = False
    cmdRemove.Enabled = False
    
    Call ReadData(mstrDataKey)
    
    mblnDataChanged = False
    
    Me.Show 1, mfrmParent
    
End Sub

Public Sub DeleteData(ByVal strDataKey As String)
    '******************************************************************************************************************
    '���ܣ�ɾ��Ͷ��Ŀ��
    '������
    '���أ�
    '******************************************************************************************************************
    mbytMode = 3
    If strDataKey = "" Then Exit Sub
    mstrDataKey = strDataKey
        
    Set mrsPara = zlCommFun.CreateParameter
    Call zlCommFun.SetParameter(mrsPara, "id", mstrDataKey)
        
    If gclsBusiness.ItemDeliverEdit("Delete", mrsPara) Then
        RaiseEvent AfterDeleteData(mstrDataKey)
    End If
End Sub

'######################################################################################################################
Private Function InitGrid() As Boolean
    '******************************************************************************************************************
    '���ܣ���ʼ�����ؼ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim intRow As Integer
    Dim intCol As Integer
    Dim intStartRow As Integer
    
    picKind(1).Move picKind(0).Left, picKind(0).Top, picKind(0).Width, picKind(0).Height
    
    For intLoop = 2 To 7
        picBack(intLoop).Left = picBack(1).Left
        picBack(intLoop).Top = picBack(1).Top
        picBack(intLoop).Width = picBack(1).Width
        picBack(intLoop).Height = picBack(1).Height
    Next
       
    For intLoop = 0 To 1
        picKind(intLoop).BorderStyle = 0
    Next
    
    For intLoop = 1 To 7
        picBack(intLoop).BorderStyle = 0
    Next
    
    Set mclsVsf(0) = New zlVSFlexGrid.clsVsf
    With mclsVsf(0)
        Call .Initialize(Me.Controls, vsf(0), True, False, gclsBusiness.GetImageList(16))
        Call .ClearColumn
        
        Call .AppendColumn("", 270, flexAlignCenterCenter, flexDTString, , "[���]", False, False, False)
        Call .AppendColumn("id", 0, flexAlignLeftCenter, flexDTString, , "", True, , , True)
        Call .AppendColumn("�ϼ�id", 0, flexAlignLeftCenter, flexDTString, , "", True, , , True)
        Call .AppendColumn("״̬", 0, flexAlignLeftCenter, flexDTString, , "", True, , , True)
        Call .AppendColumn("����", 1500, flexAlignLeftCenter, flexDTString, , "", True)
                
        .IndicatorMode = 2
        .IndicatorCol = .ColIndex("���")
        .ConstCol = .ColIndex("���")
        .VsfObject.OutlineCol = .ColIndex("����")
        .VsfObject.RowHidden(0) = True
        
    End With
    
    With vsf(0)
        .MergeCells = flexMergeFree
        .MergeCol(.ColIndex("����")) = True
    End With
    mclsVsf(0).AppendRows = False
    
    '------------------------------------------------------------------------------------------------------------------
    Set mclsVsf(1) = New zlVSFlexGrid.clsVsf
    With mclsVsf(1)
        Call .Initialize(Me.Controls, vsf(1), True, False, gclsBusiness.GetImageList(16))
        Call .ClearColumn
        Call .AppendColumn("", 270, flexAlignCenterCenter, flexDTString, , "[���]", False, False, False)
        Call .AppendColumn("id", 0, flexAlignLeftCenter, flexDTString, , "", True, , , True)
        Call .AppendColumn("�ϼ�id", 0, flexAlignLeftCenter, flexDTString, , "", True, , , True)
        Call .AppendColumn("����", 990, flexAlignLeftCenter, flexDTString, , "", True)
                        
        .IndicatorMode = 2
        .IndicatorCol = .ColIndex("���")
        .ConstCol = .ColIndex("���")
        .VsfObject.OutlineCol = .ColIndex("����")
        .VsfObject.RowHidden(0) = True
        .AppendRows = False
        
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    Set mclsVsf(2) = New zlVSFlexGrid.clsVsf
    With mclsVsf(2)
        Call .Initialize(Me.Controls, vsf(2), True, False, gclsBusiness.GetImageList(16))
        Call .ClearColumn
        Call .AppendColumn("", 270, flexAlignCenterCenter, flexDTString, , "[���]", False, False, False)
        
        Call .AppendColumn("id", 0, flexAlignLeftCenter, flexDTString, , "", True, , , True)
        Call .AppendColumn("�ϼ�id", 0, flexAlignLeftCenter, flexDTString, , "", True, , , True)
        Call .AppendColumn("����", 3000, flexAlignLeftCenter, flexDTString, , "", True)
        
        .VsfObject.OutlineCol = .ColIndex("����")
        .IndicatorMode = 2
        .IndicatorCol = .ColIndex("���")
        .ConstCol = .ColIndex("���")
        
        vsf(2).RowHidden(0) = True
        
        .AppendRows = False
        
        
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    Set mclsVsf(3) = New zlVSFlexGrid.clsVsf
    With mclsVsf(3)
        Call .Initialize(Me.Controls, vsf(3), True, False, gclsBusiness.GetImageList(16))
        Call .ClearColumn
        Call .AppendColumn("", 270, flexAlignCenterCenter, flexDTString, , "[���]", False, False, False)
        Call .AppendColumn("id", 0, flexAlignLeftCenter, flexDTString, , "", True, , , True)
        Call .AppendColumn("����", 1500, flexAlignLeftCenter, flexDTString, , "", True)
                
        .IndicatorMode = 2
        .IndicatorCol = .ColIndex("���")
        .ConstCol = .ColIndex("���")
        .VsfObject.RowHidden(0) = True
        .AppendRows = False
        
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    Set mclsVsf(4) = New zlVSFlexGrid.clsVsf
    With mclsVsf(4)
        Call .Initialize(Me.Controls, vsf(4), True, False, gclsBusiness.GetImageList(16))
        Call .ClearColumn
        Call .AppendColumn("", 270, flexAlignCenterCenter, flexDTString, , "[���]", False, False, False)
        Call .AppendColumn("id", 0, flexAlignLeftCenter, flexDTString, , "", True, , , True)
        Call .AppendColumn("�ϼ�id", 0, flexAlignLeftCenter, flexDTString, , "", True, , , True)
        Call .AppendColumn("����", 1500, flexAlignLeftCenter, flexDTString, , "", True)
        
        .IndicatorMode = 2
        .IndicatorCol = .ColIndex("���")
        .ConstCol = .ColIndex("���")
        .VsfObject.OutlineCol = .ColIndex("����")
        .VsfObject.RowHidden(0) = True
        
        .AppendRows = False
        
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    Set mclsVsf(5) = New zlVSFlexGrid.clsVsf
    With mclsVsf(5)
        Call .Initialize(Me.Controls, vsf(5), True, False, gclsBusiness.GetImageList(16))
        Call .ClearColumn
        Call .AppendColumn("", 270, flexAlignCenterCenter, flexDTString, , "[���]", False, False, False)
        Call .AppendColumn("id", 0, flexAlignLeftCenter, flexDTString, , "", True, , , True)
        Call .AppendColumn("�ϼ�id", 0, flexAlignLeftCenter, flexDTString, , "", True, , , True)
        Call .AppendColumn("����", 2100, flexAlignLeftCenter, flexDTString, , "", True)
'        Call .AppendColumn("����", 1500, flexAlignLeftCenter, flexDTString, , "", True)
        
        .VsfObject.OutlineCol = .ColIndex("����")
        .IndicatorMode = 2
        .IndicatorCol = .ColIndex("���")
        .ConstCol = .ColIndex("���")
        .VsfObject.RowHidden(0) = True
        .AppendRows = False
        
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    Set mclsVsf(6) = New zlVSFlexGrid.clsVsf
    With mclsVsf(6)
        Call .Initialize(Me.Controls, vsf(6), True, False, gclsBusiness.GetImageList(16))
        Call .ClearColumn
        Call .AppendColumn("", 270, flexAlignCenterCenter, flexDTString, , "[���]", False, False, False)
        Call .AppendColumn("id", 0, flexAlignLeftCenter, flexDTString, , "", True, , , True)
        Call .AppendColumn("�ϼ�id", 0, flexAlignLeftCenter, flexDTString, , "", True, , , True)
        Call .AppendColumn("����", 3000, flexAlignLeftCenter, flexDTString, , "", True)
                
        .VsfObject.OutlineCol = .ColIndex("����")
        .IndicatorMode = 2
        .IndicatorCol = .ColIndex("���")
        .ConstCol = .ColIndex("���")
        .VsfObject.RowHidden(0) = True
        .AppendRows = False
        
    End With
     
     
    '------------------------------------------------------------------------------------------------------------------
    Set mclsVsf(7) = New zlVSFlexGrid.clsVsf
    With mclsVsf(7)
        Call .Initialize(Me.Controls, vsf(7), True, False, GetImageList(16))
        Call .ClearColumn
        
        Call .AppendColumn("", 720, flexAlignCenterCenter, flexDTString, , "[���]", False, False, False)
        Call .AppendColumn("id", 0, flexAlignLeftCenter, flexDTString, , "", True, , , True)
        Call .AppendColumn("parent_id", 0, flexAlignLeftCenter, flexDTString, , "", True, , , True)
        Call .AppendColumn("�ڵ����", 2100, flexAlignLeftCenter, flexDTString, , "", True)
        Call .AppendColumn("�����ظ�", 1500, flexAlignLeftCenter, flexDTString, , "", True)
                                        
        .VsfObject.OutlineCol = .ColIndex("�ڵ����")
        .IndicatorMode = 2
        .IndicatorCol = .ColIndex("���")
        .ConstCol = .ColIndex("���")
        
    End With
    
End Function

Private Function InitData() As Boolean
    '******************************************************************************************************************
    '���ܣ���ʼ������
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo errHand
    
    mblnContiune = False
    mblnReading = True
    
    Set mrsSelTarget = New ADODB.Recordset
    With mrsSelTarget
        .Fields.Append "id", adVarChar, 100, adFldKeyColumn
        .Fields.Append "�ϼ�id", adVarChar, 200
        .Fields.Append "״̬", adTinyInt
        .Fields.Append "����", adVarChar, 200
        .Fields.Append "����1", adBigInt
        .Fields.Append "����2", adVarChar, 200
        .Open
    End With
    
    Set rsTmp = gclsBusiness.ItemDeliverStruct()
    If Not (rsTmp Is Nothing) Then
        txt(3).MaxLength = rsTmp("deliver_title").DefinedSize
    End If
'
    With cbo(0)
        .Clear
        For i = 1 To 7
            .AddItem i & " - " & GetTargetTitle(i)
            .ItemData(.NewIndex) = i
        Next
        .ListIndex = 0
    End With
        
    mblnReading = False
    
    Call cbo_Click(0)
        
    InitData = True
    '------------------------------------------------------------------------------------------------------------------
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

'Private Function GetTargetCode(ByVal strTargetTitle As String) As Integer
'
'    If InStr(strTargetTitle, "-") > 0 Then strTargetTitle = Trim(Mid(strTargetTitle, 4))
'    Select Case strTargetTitle
'    Case "��Ʒ��Ա", "��Ϣ��Ա"
'        GetTargetCode = 1
'    Case "��Ʒ����", "��Ϣ����"
'        GetTargetCode = 2
'    Case "��Ʒ��Ա����", "��Ϣ����"
'        GetTargetCode = 3
'    Case "��Ʒ��ɫ", "��Ϣ��ɫ"
'        GetTargetCode = 4
'    Case "��Ʒվ��", "��Ϣվ��"
'        GetTargetCode = 5
'    Case "��Ʒģ��", "��Ϣģ��"
'        GetTargetCode = 6
'    Case "��Ϣ�û�"
'        GetTargetCode = 7
'    End Select
'End Function

Private Function GetTargetTitle(ByVal intTargetType As Integer, Optional ByVal bytSource As Byte = 1) As String
    Select Case intTargetType
    Case 1
        GetTargetTitle = IIf(bytSource = 1, "��Ա", "��Ϣ��Ա")
    Case 2
        GetTargetTitle = IIf(bytSource = 1, "����", "��Ϣ����")
    Case 3
        GetTargetTitle = IIf(bytSource = 1, "��Ա����", "��Ϣ����")
    Case 4
        GetTargetTitle = IIf(bytSource = 1, "��ɫ", "��Ϣ��ɫ")
    Case 5
        GetTargetTitle = IIf(bytSource = 1, "վ��", "��Ϣվ��")
    Case 6
        GetTargetTitle = IIf(bytSource = 1, "ģ��", "��Ϣģ��")
    Case 7
        GetTargetTitle = "��Ϣ�û�"
    End Select
End Function

Private Function LoadTargetTypeData(ByVal strTargetType As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�װ��Ŀ�����ʹ�ѡ����
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rsTmp As ADODB.Recordset
    Dim intLoop As Integer
    Dim intCount As Integer
    Dim intMaxOutlineLevel As Integer
    
    On Error GoTo errHand
    
    Select Case strTargetType
    '------------------------------------------------------------------------------------------------------------------
    Case "��Ա"
        Set rsTmp = gclsBusiness.PersonRead()
        If Not (rsTmp Is Nothing) Then
            With mclsVsf(1)
                Call .LoadGrid(rsTmp)
                intMaxOutlineLevel = .ShowOutline(.ColIndex("id"), .ColIndex("�ϼ�id"))
                For intLoop = intMaxOutlineLevel To 1 Step -1
                    Call .OutLine(intLoop)
                Next
            End With
            Call ExpandAllOutline(vsf(1))
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case "����"
        Set rsTmp = gclsBusiness.DeptRead()
        If Not (rsTmp Is Nothing) Then
            With mclsVsf(2)
                Call .LoadGrid(rsTmp)
                intMaxOutlineLevel = .ShowOutline(.ColIndex("id"), .ColIndex("�ϼ�id"))
                For intLoop = intMaxOutlineLevel To 1 Step -1
                    Call .OutLine(intLoop)
                Next
            End With
            Call ExpandAllOutline(vsf(2))
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case "��Ա����"
        Set rsTmp = gclsBusiness.PersonPropertyRead()
        If Not (rsTmp Is Nothing) Then
            Call mclsVsf(3).LoadDataSource(rsTmp)
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case "��ɫ"
        Set rsTmp = gclsBusiness.RoleRead()
        If Not (rsTmp Is Nothing) Then
            With mclsVsf(4)
                Call .LoadGrid(rsTmp)
                intMaxOutlineLevel = .ShowOutline(.ColIndex("id"), .ColIndex("�ϼ�id"))
                For intLoop = intMaxOutlineLevel To 1 Step -1
                    Call .OutLine(intLoop)
                Next
            End With
            Call ExpandAllOutline(vsf(4))
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case "վ��"
        Set rsTmp = gclsBusiness.StationRead()
        If Not (rsTmp Is Nothing) Then
            With mclsVsf(5)
                Call .LoadGrid(rsTmp)
                intMaxOutlineLevel = .ShowOutline(.ColIndex("id"), .ColIndex("�ϼ�id"))
                For intLoop = intMaxOutlineLevel To 1 Step -1
                    Call .OutLine(intLoop)
                Next
            End With
            Call ExpandAllOutline(vsf(5))
        End If
    
    '------------------------------------------------------------------------------------------------------------------
    Case "ģ��"
        Set rsTmp = gclsBusiness.ModuleRead()
        If Not (rsTmp Is Nothing) Then
            With mclsVsf(6)
                Call .LoadGrid(rsTmp)
                intMaxOutlineLevel = .ShowOutline(.ColIndex("id"), .ColIndex("�ϼ�id"))
'                For intLoop = intMaxOutlineLevel To 1 Step -1
'                    Call .Outline(intLoop)
'                Next
            End With
    '        Call OutlineExpand(intMaxOutlineLevel)
            Call ExpandAllOutline(vsf(6))
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case "��Ϣ����"
    
        intMaxOutlineLevel = 0
        
        Set mrsPara = zlCommFun.CreateCondition
        Call zlCommFun.SetCondition(mrsPara, "item_id", mstrItemDataKey)
        
        With mclsVsf(7)
            Call .LoadGrid(gclsBusiness.ItemConfigRead("item_id", mrsPara))
            intMaxOutlineLevel = .ShowOutline(.ColIndex("id"), .ColIndex("parent_id"))
                    
'            For intLoop = intMaxOutlineLevel To 1 Step -1
'                Call .Outline(intLoop)
'            Next
        End With
        Call ExpandAllOutline(vsf(7))
    End Select
    
    LoadTargetTypeData = True
    '------------------------------------------------------------------------------------------------------------------
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ReadData(ByVal strDataKey As String) As Boolean
    '******************************************************************************************************************
    '���ܣ���ȡͶ��Ŀ������
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strDeliveobject As String
    Dim rsTmp As ADODB.Recordset
    Dim rsCondition As ADODB.Recordset
    Dim objXML As New clsMessageXML
    Dim strTitle As String
    Dim strKey As String
    Dim strNodeName As String
    Dim strSys As String
    Dim strPrivilige As String
    
    On Error GoTo errHand
    
    mclsVsf(0).ClearGrid
    
    Set rsCondition = zlCommFun.CreateCondition
    Call zlCommFun.SetCondition(rsCondition, "id", strDataKey)

    mblnReading = True
    Set rsTmp = gclsBusiness.ItemDeliverRead("id", rsCondition)
    If rsTmp.BOF = False Then
        txt(3).Text = zlCommFun.NVL(rsTmp("deliver_title").Value)
        txt(2).Text = zlCommFun.NVL(rsTmp("deliver_code").Value)
        strDeliveobject = zlCommFun.NVL(rsTmp("deliver_object").Value)
        
        Call gclsBusiness.GetDeliveryTree(strDeliveobject, mrsSelTarget)
        Call UpdateTargetGrid
        
    End If
    
    mblnReading = False
    mblnDataChanged = False
    
    ReadData = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub UpdateTargetGrid()
    Dim intMaxOutlineLevel As Integer
    Dim intLoop As Integer
        
    With mclsVsf(0)
        .ClearGrid
        
        mrsSelTarget.Filter = ""
        mrsSelTarget.Sort = "����1"
        If mrsSelTarget.RecordCount > 0 Then
            mrsSelTarget.MoveFirst
        
            Call .LoadGrid(mrsSelTarget)
            intMaxOutlineLevel = .ShowOutline(.ColIndex("id"), .ColIndex("�ϼ�id"))
        
            Call UpdateCollapseState
            .VsfObject.ShowCell .VsfObject.Row, .VsfObject.ColIndex("����")
        
        End If
    
    End With
End Sub

Private Sub UpdateCollapseState()
    Dim lngRow As Integer
    
    With vsf(0)
        For lngRow = 1 To .Rows - 1
            .IsCollapsed(lngRow) = IIf(Val(.TextMatrix(lngRow, .ColIndex("״̬"))) = 1, flexOutlineExpanded, flexOutlineCollapsed)
        Next
    End With
End Sub

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
    Dim cbrCustom As CommandBarControlCustom
    Dim objFindKey As CommandBarControl
    
    On Error GoTo errHand
    
    '------------------------------------------------------------------------------------------------------------------
    '��ʼ����
    Call zlCommFun.CommandBarInit(cbsMain)
    cbsMain.VisualTheme = xtpThemeWhidbey
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    cbsMain.Options.LargeIcons = False
    
    '------------------------------------------------------------------------------------------------------------------
    '�˵�����:�����������ݣ����xtpControlPopup���͵�����ID���¸�ֵ

    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.Visible = False
    
    '------------------------------------------------------------------------------------------------------------------
    '����������:������������
    
    
    Set objBar = cbsMain.Add("������", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
        

    
    mstrFindKey = zlDataBase.GetPara("��λ����", ParamInfo.ϵͳ��, mlngModualCode, "����")
    If mstrFindKey = "" Then mstrFindKey = "����"

    Set mobjFindKey = zlCommFun.NewToolBar(objBar, xtpControlPopup, conMenu_View_LocationItem, mstrFindKey, True, , xtpButtonIconAndCaption)
    mobjFindKey.IconId = conMenu_View_Find
    mobjFindKey.flags = xtpFlagRightAlign
    mobjFindKey.Style = xtpButtonIconAndCaption
    Set objControl = mobjFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_View_LocationItem, "&1.����"): objControl.Parameter = "����"
    objControl.IconId = 1
'    Set objControl = mobjFindKey.CommandBar.Controls.Add(xtpControlButton, conMenu_View_LocationItem, "&2.����"): objControl.Parameter = "����"
'    objControl.IconId = 1

    Set cbrCustom = zlCommFun.NewToolBar(objBar, xtpControlCustom, 0, "")
    cbrCustom.Handle = txtLocation.hWnd
        
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_View_Filter, "����")
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_View_Forward, "��һ��")
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_View_Backward, "��һ��")
        
    
    Set objControl = zlCommFun.NewToolBar(objBar, xtpControlButton, conMenu_View_Option, IIf(mbytMode = 1, "ȷ��֮��������", "ȷ��֮�����޸�"), False)
    objControl.IconId = conMenu_View_UnCheck
    If mbytMode <> 1 Then objControl.flags = xtpFlagRightAlign
    
    txtLocation.Visible = (mbytMode = 2)
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
    
End Function

Private Function ValidData() As Boolean
    '******************************************************************************************************************
    '���ܣ�У��༭���ݵ���Ч��
    '������
    '���أ�
    '******************************************************************************************************************
    
    If Len(txt(3).Text) = 0 Then
        ShowSimpleMsg "Ͷ����������Ϊ�գ�"
        Call LocationObj(txt(3))
        Exit Function
    End If
    
'    With vsf(0)
'        If .TextMatrix(1, .ColIndex("����")) = "" Then
'            ShowSimpleMsg "Ͷ��Ŀ�����ò���Ϊ�գ�"
'            Call LocationObj(vsf(0))
'            Exit Function
'        End If
'    End With
    
    ValidData = True
    
End Function

Private Function SaveData(ByRef strDataKey As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rsPara As ADODB.Recordset
    Dim strDeliverObject As String
    Dim strDeliverProduct As String
    Dim intRow As Integer
    Dim intCol As Integer
    Dim strTargetType As String
    Dim strTemp As String
    Dim aryTemp As Variant
    
    On Error GoTo errHand
        
    Set rsPara = zlCommFun.CreateParameter
    Call zlCommFun.SetParameter(rsPara, "id", strDataKey)
    Call zlCommFun.SetParameter(rsPara, "item_id", mstrItemDataKey)
    Call zlCommFun.SetParameter(rsPara, "deliver_title", Trim(txt(3).Text))
    Call zlCommFun.SetParameter(rsPara, "deliver_code", Trim(txt(2).Text))
    
    With vsf(0)
        strDeliverProduct = ""
                
        For intRow = 1 To .Rows - 1
            
            If .TextMatrix(intRow, .ColIndex("�ϼ�id")) = "" Then
                Select Case strTargetType
                Case "��Ա"
                    strDeliverProduct = strDeliverProduct & "</persons>"
                Case "����"
                    strDeliverProduct = strDeliverProduct & "</depts>"
                Case "��Ա����"
                    strDeliverProduct = strDeliverProduct & "</personpropertys>"
                Case "��ɫ"
                    strDeliverProduct = strDeliverProduct & "</roles>"
                Case "վ��"
                    strDeliverProduct = strDeliverProduct & "</stations>"
                Case "ģ��"
                    strDeliverProduct = strDeliverProduct & "</modules>"
                Case "��Ϣ�û�"
                    strDeliverProduct = strDeliverProduct & "</mipusers>"
                End Select
                
                strTargetType = .TextMatrix(intRow, .ColIndex("id"))
                Select Case strTargetType
                Case "��Ա"
                    strDeliverProduct = strDeliverProduct & "<persons>"
                Case "����"
                    strDeliverProduct = strDeliverProduct & "<depts>"
                Case "��Ա����"
                    strDeliverProduct = strDeliverProduct & "<personpropertys>"
                Case "��ɫ"
                    strDeliverProduct = strDeliverProduct & "<roles>"
                Case "վ��"
                    strDeliverProduct = strDeliverProduct & "<stations>"
                Case "ģ��"
                    strDeliverProduct = strDeliverProduct & "<modules>"
                Case "��Ϣ�û�"
                    strDeliverProduct = strDeliverProduct & "<mipusers>"
                End Select
            Else
                                
                Select Case Left(.TextMatrix(intRow, .ColIndex("id")), 6)
                Case "��Ʒ(��Ա)"
                    strDeliverProduct = strDeliverProduct & "<person>"
                    strDeliverProduct = strDeliverProduct & "<title>" & .TextMatrix(intRow, .ColIndex("����")) & "</title>"
                    strDeliverProduct = strDeliverProduct & "<key>" & Mid(.TextMatrix(intRow, .ColIndex("id")), 8) & "</key>"
                    strDeliverProduct = strDeliverProduct & "</person>"
                Case "��Ʒ(����)"
                    strDeliverProduct = strDeliverProduct & "<dept>"
                    strDeliverProduct = strDeliverProduct & "<title>" & .TextMatrix(intRow, .ColIndex("����")) & "</title>"
                    strDeliverProduct = strDeliverProduct & "<key>" & Mid(.TextMatrix(intRow, .ColIndex("id")), 8) & "</key>"
                    strDeliverProduct = strDeliverProduct & "</dept>"
                Case "��Ʒ(����)"
                    strDeliverProduct = strDeliverProduct & "<personproperty>" & Mid(.TextMatrix(intRow, .ColIndex("id")), 8) & "</personproperty>"
                Case "��Ʒ(��ɫ)"
                    strDeliverProduct = strDeliverProduct & "<role>" & Mid(.TextMatrix(intRow, .ColIndex("id")), 8) & "</role>"
                Case "��Ʒ(վ��)"
                    strDeliverProduct = strDeliverProduct & "<station>" & Mid(.TextMatrix(intRow, .ColIndex("id")), 8) & "</station>"
                Case "��Ʒ(�û�)"
                    strDeliverProduct = strDeliverProduct & "<mipuser>" & Mid(.TextMatrix(intRow, .ColIndex("id")), 8) & "</mipuser>"
                Case "��Ʒ(ģ��)"
                    'ϵͳ��_���_����
                    aryTemp = Split(strTemp, "_")
                    strDeliverProduct = strDeliverProduct & "<module>"
                    strDeliverProduct = strDeliverProduct & "<title>" & .TextMatrix(intRow, .ColIndex("����")) & "</title>"
                    strDeliverProduct = strDeliverProduct & "<key>" & Mid(.TextMatrix(intRow, .ColIndex("id")), 8) & "</key>"
                    strDeliverProduct = strDeliverProduct & "</module>"
                Case "��Ϣ(��Ա)"
                    strDeliverProduct = strDeliverProduct & "<message>" & .TextMatrix(intRow, .ColIndex("����")) & "</message>"
                Case "��Ϣ(����)"
                    strDeliverProduct = strDeliverProduct & "<message>" & .TextMatrix(intRow, .ColIndex("����")) & "</message>"
                Case "��Ϣ(����)"
                    strDeliverProduct = strDeliverProduct & "<message>" & .TextMatrix(intRow, .ColIndex("����")) & "</message>"
                Case "��Ϣ(��ɫ)"
                    strDeliverProduct = strDeliverProduct & "<message>" & .TextMatrix(intRow, .ColIndex("����")) & "</message>"
                Case "��Ϣ(վ��)"
                    strDeliverProduct = strDeliverProduct & "<message>" & .TextMatrix(intRow, .ColIndex("����")) & "</message>"
                Case "��Ϣ(ģ��)"
                    strDeliverProduct = strDeliverProduct & "<message>" & .TextMatrix(intRow, .ColIndex("����")) & "</message>"
                Case "��Ϣ(�û�)"
                    strDeliverProduct = strDeliverProduct & "<message>" & .TextMatrix(intRow, .ColIndex("����")) & "</message>"
                End Select
            
            End If
            
        Next
    End With
    
    If strTargetType <> "" Then
        Select Case strTargetType
        Case "��Ա"
            strDeliverProduct = strDeliverProduct & "</persons>"
        Case "����"
            strDeliverProduct = strDeliverProduct & "</depts>"
        Case "��Ա����"
            strDeliverProduct = strDeliverProduct & "</personpropertys>"
        Case "��ɫ"
            strDeliverProduct = strDeliverProduct & "</roles>"
        Case "վ��"
            strDeliverProduct = strDeliverProduct & "</stations>"
        Case "ģ��"
            strDeliverProduct = strDeliverProduct & "</modules>"
        Case "��Ϣ�û�"
            strDeliverProduct = strDeliverProduct & "</mipusers>"
        End Select
    End If
    
    If strDeliverProduct <> "" Then
        strDeliverObject = "<deliverobject>" & strDeliverProduct & "</deliverobject>"
    End If
        
    Call zlCommFun.SetParameter(rsPara, "deliver_object", strDeliverObject)

    Select Case mbytMode
    '------------------------------------------------------------------------------------------------------------------
    Case 1          '����
        strDataKey = zlCommFun.GetGUID
        Call zlCommFun.SetParameter(rsPara, "id", strDataKey)
        SaveData = gclsBusiness.ItemDeliverEdit("INSERT", rsPara)
    '------------------------------------------------------------------------------------------------------------------
    Case 2          '�޸�
        SaveData = gclsBusiness.ItemDeliverEdit("UPDATE", rsPara)
    End Select
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    
    If zlComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub cbo_Click(Index As Integer)
    Dim intLoop As Integer
    Dim intIndex As Integer
    
    If mblnReading = True Then Exit Sub
    
    With cbo(Index)
        Select Case .ItemData(.ListIndex)
        Case 1, 2, 3, 4, 5, 6
            picKind(0).Visible = True
            picKind(1).Visible = False
        Case 7
            picKind(0).Visible = False
            picKind(1).Visible = True
        End Select
    End With
    
    If opt(0).Value = True Then
        For intLoop = 1 To 7
            picBack(intLoop).Visible = False
        Next
        intIndex = cbo(Index).ItemData(cbo(Index).ListIndex)
        picBack(intIndex).Visible = True
    Else
        If picBack(7).Visible = False Then
            For intLoop = 1 To 7
                picBack(intLoop).Visible = False
            Next
            
        End If
        intIndex = 7
    End If
    picBack(intIndex).Visible = True
    
    '��ȡ����
    If picBack(intIndex).Tag = "" Then
        If intIndex = 7 Then
            Call LoadTargetTypeData("��Ϣ����")
        Else
            Call LoadTargetTypeData(GetTargetTitle(cbo(0).ItemData(cbo(0).ListIndex)))
        End If
        picBack(intIndex).Tag = "loaded"
    End If
    
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    Dim blnCancel As Boolean
    Dim strDataKey As String
    
    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Forward               '��һ��
        
        strDataKey = mstrDataKey
        
        RaiseEvent Forward(strDataKey, blnCancel)
        If blnCancel = False Then
        
            mstrDataKey = strDataKey
            Call ReadData(mstrDataKey)
    
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Backward               '��һ��
        
        strDataKey = mstrDataKey
        
        RaiseEvent Backward(strDataKey, blnCancel)
        If blnCancel = False Then
            
            mstrDataKey = strDataKey
            Call ReadData(mstrDataKey)
            
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_LocationItem
        mstrFindKey = Control.Parameter
        mobjFindKey.Caption = mstrFindKey
        cbsMain.RecalcLayout
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Filter
        
        Dim strText As String
        Dim rsCondition As ADODB.Recordset
        Dim rsData As ADODB.Recordset
        Dim rs As ADODB.Recordset
        
        If txtLocation.Text <> "" Then
            
            txtLocation.Tag = ""
            
            
            Set rsCondition = zlCommFun.CreateCondition
            
            Call zlCommFun.SetCondition(rsCondition, "FilterStyle", mstrFindKey)
            Call zlCommFun.SetCondition(rsCondition, "FilterText", txtLocation.Text)
            Set rsData = gclsBusiness.EventRead("FilterData", rsCondition)
            
            If zlCommFun.ShowPubSelect(Me, txtLocation, 2, "����,1500,0,1;����,1500,0,0;����,1500,0,0;�豸,1500,0,0", Me.Name & "\ҵ���¼�����", "����±���ѡ��һ��ҵ���¼�", rsData, rs, , , , , , True) = 1 Then
                mstrDataKey = rs("id").Value
                Call ReadData(mstrDataKey)
                txtLocation.Tag = ""
            Else
                txtLocation.Tag = ""
                Call LocationObj(txtLocation, True)
                Exit Sub
            End If
                        
            Call LocationObj(txtLocation, True)
        End If
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_LocationItem
    
        mstrFindKey = Control.Parameter
        mobjFindKey.Caption = mstrFindKey
        cbsMain.RecalcLayout
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Option
        mblnContiune = Not mblnContiune
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
    picPane(0).Move lngLeft, lngTop, lngRight - lngLeft, lngBottom - lngTop
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_View_Filter, conMenu_View_LocationItem, conMenu_View_Backward, conMenu_View_Forward, 0
        Control.Visible = (mbytMode = 2)
        Control.Enabled = Not mblnDataChanged
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Option
        Control.Checked = mblnContiune
        Control.IconId = IIf(mblnContiune = True, conMenu_View_Check, conMenu_View_UnCheck)
    End Select
End Sub

Private Sub cmdAdd_Click()
    Dim strKey As String
    Dim strContent As String
    Dim intRow As Integer
    Dim strParentKey As String
            
    If cbo(0).ItemData(cbo(0).ListIndex) = 7 Then
        If Trim(txt(4).Text) <> "" Then
            strKey = "��Ʒ(�û�)_" & txt(4).Text
            If mclsVsf(0).CheckHave(strKey, False, mclsVsf(0).ColIndex("id")) = False Then
                Call gclsBusiness.InsertDeliveryTree("��Ϣ�û�", strKey, txt(4).Text, mrsSelTarget)
                Call UpdateTargetGrid
                mblnDataChanged = True
            End If
        End If
    Else
        If opt(0).Value = True Then
            Select Case cbo(0).ItemData(cbo(0).ListIndex)
            '------------------------------------------------------------------------------------------------------------------
            Case 1
                With vsf(1)
                    If .TextMatrix(.Row, .ColIndex("�ϼ�id")) <> "" Then
                        strKey = "��Ʒ(��Ա)_" & Mid(.TextMatrix(.Row, .ColIndex("id")), 2)       '��Աid
                        If mclsVsf(0).CheckHave(strKey, False, mclsVsf(0).ColIndex("id")) = False Then
                            Call gclsBusiness.InsertDeliveryTree("��Ա", strKey, .TextMatrix(.Row, .ColIndex("����")), mrsSelTarget)
                            Call UpdateTargetGrid
                            mblnDataChanged = True
                        End If
                    End If
                End With
            '------------------------------------------------------------------------------------------------------------------
            Case 2
                With vsf(2)
                    If .TextMatrix(.Row, .ColIndex("�ϼ�id")) <> "" Then
                        strKey = "��Ʒ(����)_" & Mid(.TextMatrix(.Row, .ColIndex("id")), 2)               '����id
                        If mclsVsf(0).CheckHave(strKey, False, mclsVsf(0).ColIndex("id")) = False Then
                            Call gclsBusiness.InsertDeliveryTree("����", strKey, .TextMatrix(.Row, .ColIndex("����")), mrsSelTarget)
                            Call UpdateTargetGrid
                            mblnDataChanged = True
                        End If
                    End If
                End With
            '------------------------------------------------------------------------------------------------------------------
            Case 3
                With vsf(3)
                    strKey = "��Ʒ(����)_" & .TextMatrix(.Row, .ColIndex("id"))                   '��Ա����
                    If mclsVsf(0).CheckHave(strKey, False, mclsVsf(0).ColIndex("id")) = False Then
                        Call gclsBusiness.InsertDeliveryTree("��Ա����", strKey, .TextMatrix(.Row, .ColIndex("����")), mrsSelTarget)
                        Call UpdateTargetGrid
                        mblnDataChanged = True
                    End If
                End With
            '------------------------------------------------------------------------------------------------------------------
            Case 4
                With vsf(4)
'                    If .TextMatrix(.Row, .ColIndex("�ϼ�id")) <> "" Then
                    strKey = "��Ʒ(��ɫ)_" & Mid(.TextMatrix(.Row, .ColIndex("id")), 2)       '��ɫ����(ȥ��ZL_)
                    If mclsVsf(0).CheckHave(strKey, False, mclsVsf(0).ColIndex("id")) = False Then
                        Call gclsBusiness.InsertDeliveryTree("��ɫ", strKey, .TextMatrix(.Row, .ColIndex("����")), mrsSelTarget)
                        Call UpdateTargetGrid
                        mblnDataChanged = True
                    End If
'                    End If
                End With
            '------------------------------------------------------------------------------------------------------------------
            Case 5
                With vsf(5)
                    If .TextMatrix(.Row, .ColIndex("�ϼ�id")) <> "" Then
                        strKey = "��Ʒ(վ��)_" & Mid(.TextMatrix(.Row, .ColIndex("id")), 2)       'ip
                        If mclsVsf(0).CheckHave(strKey, False, mclsVsf(0).ColIndex("id")) = False Then
                            Call gclsBusiness.InsertDeliveryTree("վ��", strKey, .TextMatrix(.Row, .ColIndex("����")), mrsSelTarget)
                            Call UpdateTargetGrid
                            mblnDataChanged = True
                        End If
                    End If
                End With
            '------------------------------------------------------------------------------------------------------------------
            Case 6
                With vsf(6)
                    If .TextMatrix(.Row, .ColIndex("�ϼ�id")) <> "" Then
                        strKey = "��Ʒ(ģ��)_" & Mid(.TextMatrix(.Row, .ColIndex("id")), 2)          'ϵͳ��_���_����
                        If mclsVsf(0).CheckHave(strKey, False, mclsVsf(0).ColIndex("id")) = False Then
                            Call gclsBusiness.InsertDeliveryTree("ģ��", strKey, .TextMatrix(.Row, .ColIndex("����")), mrsSelTarget)
                            Call UpdateTargetGrid
                            mblnDataChanged = True
                        End If
                    End If
                End With
            End Select
        Else
            With vsf(7)
                            
                strParentKey = .TextMatrix(.Row, .ColIndex("parent_id"))
                strContent = .TextMatrix(.Row, .ColIndex("�ڵ����"))
                Do While strParentKey <> ""
                    intRow = mclsVsf(7).FindRow(strParentKey, .ColIndex("id"))
                    If intRow > 0 Then
                        strParentKey = .TextMatrix(intRow, .ColIndex("parent_id"))
                        strContent = .TextMatrix(intRow, .ColIndex("�ڵ����")) & "/" & strContent
                    Else
                        strParentKey = ""
                    End If
                Loop
                            
                If strContent <> "" Then
                    strContent = "/" & strContent
                    
                    strKey = "��Ϣ(" & GetTargetTitle(cbo(0).ItemData(cbo(0).ListIndex)) & ")_" & strContent
                                
                    If mclsVsf(0).CheckHave(strKey, False, mclsVsf(0).ColIndex("id")) = False Then
                        Call gclsBusiness.InsertDeliveryTree(GetTargetTitle(cbo(0).ItemData(cbo(0).ListIndex)), strKey, strContent, mrsSelTarget)
                        Call UpdateTargetGrid
                        mblnDataChanged = True
                    End If
                End If
            End With
            
            
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    '
    Unload Me
End Sub


Private Sub cmdCollapse_Click()
    If opt(0).Value Then
        Call CollapseAllOutline(vsf(cbo(0).ItemData(cbo(0).ListIndex)))
    Else
        Call CollapseAllOutline(vsf(7))
    End If
End Sub

Private Sub cmdCollapseSel_Click()
    mrsSelTarget.Filter = ""
    If mrsSelTarget.RecordCount > 0 Then
        Do While Not mrsSelTarget.EOF
            mrsSelTarget.Update "״̬", 2
            mrsSelTarget.MoveNext
        Loop
    End If
    
    Call UpdateTargetGrid
    Call UpdateCollapseState
End Sub

Private Sub cmdExpand_Click()
    If opt(0).Value Then
        Call ExpandAllOutline(vsf(cbo(0).ItemData(cbo(0).ListIndex)))
    Else
        Call ExpandAllOutline(vsf(7))
    End If
End Sub

Private Sub cmdExpandSel_Click()
    mrsSelTarget.Filter = ""
    If mrsSelTarget.RecordCount > 0 Then
        Do While Not mrsSelTarget.EOF
            mrsSelTarget.Update "״̬", 1
            mrsSelTarget.MoveNext
        Loop
    End If
    Call UpdateTargetGrid
    Call UpdateCollapseState
End Sub

Private Sub cmdFind_Click()
    Dim lngRow As Long
    Dim intIndex As Integer
    
    If opt(0).Value Then
        intIndex = cbo(0).ItemData(cbo(0).ListIndex)
    Else
        intIndex = 7
    End If
    With mclsVsf(intIndex)
        
        lngRow = .FindRow(txt(0).Text, .ColIndex("����"), 2, .VsfObject.Row + 1)
        If lngRow = -1 Then
            lngRow = .FindRow(txt(0).Text, .ColIndex("����"), 2)
        End If
        
        If lngRow > 0 Then
            .VsfObject.Row = lngRow
            .VsfObject.ShowCell lngRow, .ColIndex("����")
        End If
        
        Call LocationObj(txt(0), True)
    End With
End Sub

Private Sub cmdFindSel_Click()
    Dim lngRow As Long
    
    With mclsVsf(0)
        
        lngRow = .FindRow(txt(1).Text, .ColIndex("����"), 2, .VsfObject.Row + 1)
        If lngRow = -1 Then
            lngRow = .FindRow(txt(1).Text, .ColIndex("����"), 2)
        End If
        
        If lngRow > 0 Then
            .VsfObject.Row = lngRow
            .VsfObject.ShowCell lngRow, .ColIndex("����")
        End If
        
        Call LocationObj(txt(1))
    End With
End Sub

Private Sub cmdOK_Click()
        
    If mblnDataChanged = True Then
        If ValidData = True Then
                
            If SaveData(mstrDataKey) = True Then
                
                Select Case mbytMode
                Case 1
                    RaiseEvent AfterNewData(mstrDataKey)
                Case 2
                    RaiseEvent AfterModifyData(mstrDataKey)
                End Select
                
                If mblnContiune = False Then
                    mblnDataChanged = False
                    Unload Me
                Else
                    '���û�����������һ������״̬
                    If mbytMode = 1 Then
                        mstrDataKey = ""
                        txt(3).Text = ""
                        mclsVsf(0).ClearGrid
                    End If
                    Call LocationObj(txt(3))
                    mblnDataChanged = False
                End If
                
            End If
        End If
    End If
    
End Sub


Private Sub cmdRemove_Click()
    Dim lngRow As Long
    Dim strUpKey As String
    
    With vsf(0)
        
        lngRow = .Row
        
        If .TextMatrix(.Row, .ColIndex("�ϼ�id")) <> "" Then
            strUpKey = .TextMatrix(.Row, .ColIndex("�ϼ�id"))
            mrsSelTarget.Filter = ""
            mrsSelTarget.Filter = "id='" & .TextMatrix(.Row, .ColIndex("id")) & "'"
            If mrsSelTarget.RecordCount > 0 Then
                mrsSelTarget.Delete adAffectCurrent
                mrsSelTarget.Filter = ""
                mrsSelTarget.Filter = "�ϼ�id='" & strUpKey & "'"
                If mrsSelTarget.RecordCount = 0 Then
                    mrsSelTarget.Filter = ""
                    mrsSelTarget.Filter = "id='" & strUpKey & "'"
                    If mrsSelTarget.RecordCount > 0 Then
                        mrsSelTarget.Delete adAffectCurrent
                    End If
                End If
                
                Call UpdateTargetGrid
                mblnDataChanged = True
                
                If .Rows > lngRow Then
                    .Row = lngRow
                Else
                    .Row = .Rows - 1
                End If
                .ShowCell .Row, .ColIndex("����")
            End If
        End If
    End With
'
'    With vsf(0)
'        If .Rows = 2 Then
'            mclsVsf(0).ClearGrid
'        Else
'            vsf(0).RemoveItem vsf(0).Row
'            mclsVsf(0).AppendRows = True
'            mclsVsf(0).UpdateSerial
'        End If
'
'        mblnDataChanged = True
'    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnDataChanged Then
        Cancel = (MsgBox("�������޸ĵ����ݱ��뱣������Ч���Ƿ񲻱�����˳���", vbYesNo + vbQuestion + vbDefaultButton2, ParamInfo.ϵͳ����) = vbNo)
        If Cancel Then Exit Sub
    End If
        
    If Not (mclsVsf(0) Is Nothing) Then Set mclsVsf(0) = Nothing
    If Not (mclsVsf(1) Is Nothing) Then Set mclsVsf(1) = Nothing
    If Not (mclsVsf(2) Is Nothing) Then Set mclsVsf(2) = Nothing
    If Not (mclsVsf(3) Is Nothing) Then Set mclsVsf(3) = Nothing
    If Not (mclsVsf(4) Is Nothing) Then Set mclsVsf(4) = Nothing
    If Not (mclsVsf(5) Is Nothing) Then Set mclsVsf(5) = Nothing
    If Not (mclsVsf(6) Is Nothing) Then Set mclsVsf(6) = Nothing
    If Not (mclsVsf(7) Is Nothing) Then Set mclsVsf(7) = Nothing
        
    Set mrsPara = Nothing
    Set mfrmParent = Nothing
    Set mobjFindKey = Nothing
    Set mrsSelTarget = Nothing
    
End Sub


Private Sub opt_Click(Index As Integer)
    Call cbo_Click(0)
End Sub

Private Sub opt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub picBack_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 0
        txt(4).Move 15, 15, picBack(Index).Width - 30, picBack(Index).Height - 30
    Case 1
        vsf(1).Move 15, 15, picBack(Index).Width - 30, picBack(Index).Height - 30
    Case 2
        vsf(2).Move 15, 15, picBack(Index).Width - 30, picBack(Index).Height - 30
    Case 3
        vsf(3).Move 15, 15, picBack(Index).Width - 30, picBack(Index).Height - 30
    Case 4
        vsf(4).Move 15, 15, picBack(Index).Width - 30, picBack(Index).Height - 30
    Case 5
        vsf(5).Move 15, 15, picBack(Index).Width - 30, picBack(Index).Height - 30
    Case 6
        vsf(6).Move 15, 15, picBack(Index).Width - 30, picBack(Index).Height - 30
    Case 7
        vsf(7).Move 15, 15, picBack(Index).Width - 30, picBack(Index).Height - 30
    End Select
    
End Sub

Private Sub picKind_Paint(Index As Integer)
    zlControl.PicShowFlat picKind(Index), -1
End Sub

Private Sub picKind_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 1
        txt(4).Move 15, 15, picKind(Index).Width - 30, picKind(Index).Height - 30
    End Select
End Sub

Private Sub txt_Change(Index As Integer)
    
    If mblnReading Then Exit Sub
    
    Select Case Index
    Case 0, 1
        Exit Sub
    End Select
    
    mblnDataChanged = True
End Sub

Private Sub txt_GotFocus(Index As Integer)
    
    zlControl.TxtSelAll txt(Index)
    
    Select Case Index
    Case 0, 1, 3
        zlCommFun.OpenIme True
    End Select
        
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
        
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Select Case Index
        Case 0
            Call cmdFind_Click
        Case 1
            Call cmdFindSel_Click
        Case Else
            zlCommFun.PressKey vbKeyTab
        End Select
        
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)

    Select Case Index
    Case 0, 3
        zlCommFun.OpenIme False
    End Select

End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        glngTXTProc = GetWindowLong(txt(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub


Private Sub vsf_AfterCollapse(Index As Integer, ByVal Row As Long, ByVal State As Integer)
    
    With vsf(Index)
        If Index = 0 Then
            mrsSelTarget.Filter = ""
            mrsSelTarget.Filter = "id='" & .TextMatrix(Row, .ColIndex("id")) & "'"
            If mrsSelTarget.RecordCount > 0 Then
                mrsSelTarget.Update "״̬", IIf(State = 0, 1, 2)
            End If
        End If
    End With
    
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsf(Index).AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsf_DblClick(Index As Integer)
    Select Case Index
    Case 1, 2, 3, 4, 5, 6, 7
        If cmdAdd.Enabled Then Call cmdAdd_Click
    Case 0
        If cmdRemove.Enabled Then Call cmdRemove_Click
    End Select
End Sub
