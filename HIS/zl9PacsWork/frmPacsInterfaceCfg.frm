VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPacsInterfaceCfg 
   Caption         =   "������ù���"
   ClientHeight    =   10155
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   16020
   Icon            =   "frmPacsInterfaceCfg.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   10155
   ScaleWidth      =   16020
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picAppCfg 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   8895
      Left            =   5760
      ScaleHeight     =   8865
      ScaleWidth      =   10020
      TabIndex        =   4
      Top             =   360
      Width           =   10050
      Begin VB.Frame fraAppFuns 
         BorderStyle     =   0  'None
         Caption         =   "��������"
         Height          =   7380
         Left            =   240
         TabIndex        =   14
         Top             =   1320
         Width           =   9915
         Begin VB.Frame fraVBS 
            Caption         =   "VBS�ű�"
            Height          =   3315
            Left            =   5880
            TabIndex        =   23
            Top             =   3840
            Width           =   3555
            Begin VB.CheckBox chkModify 
               Caption         =   "�ֶ�����"
               Height          =   255
               Left            =   960
               TabIndex        =   28
               Top             =   0
               Width           =   1095
            End
            Begin VB.TextBox txtVBS 
               Height          =   2055
               Left            =   240
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   26
               Top             =   360
               Width           =   2655
            End
         End
         Begin VB.Frame fraFuncs 
            Caption         =   "�����б�"
            Height          =   3495
            Left            =   240
            TabIndex        =   19
            Top             =   240
            Width           =   9135
            Begin VB.CommandButton cmdTestFunc 
               Caption         =   "������֤"
               Height          =   375
               Left            =   3000
               TabIndex        =   32
               Top             =   3000
               Width           =   1215
            End
            Begin VB.CommandButton cmdDelFun 
               Caption         =   "ɾ������"
               Height          =   375
               Left            =   7080
               TabIndex        =   31
               Top             =   2880
               Width           =   1215
            End
            Begin VB.CommandButton cmdAddFunc 
               Caption         =   "��ӹ���"
               Height          =   375
               Left            =   120
               TabIndex        =   21
               Top             =   3000
               Width           =   1215
            End
            Begin VB.CommandButton cmdDelFunc 
               Caption         =   "ɾ������"
               Height          =   375
               Left            =   1560
               TabIndex        =   20
               Top             =   3000
               Width           =   1215
            End
            Begin VSFlex8Ctl.VSFlexGrid vsfAppFuns 
               Height          =   2580
               Left            =   180
               TabIndex        =   22
               Top             =   300
               Width           =   8415
               _cx             =   14843
               _cy             =   4551
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
               ForeColorSel    =   0
               BackColorBkg    =   -2147483636
               BackColorAlternate=   -2147483643
               GridColor       =   -2147483633
               GridColorFixed  =   -2147483632
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483642
               FocusRect       =   1
               HighLight       =   1
               AllowSelection  =   0   'False
               AllowBigSelection=   -1  'True
               AllowUserResizing=   0
               SelectionMode   =   1
               GridLines       =   2
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   50
               Cols            =   3
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   360
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
               Begin VSFlex8Ctl.VSFlexGrid vsfFuncs 
                  Height          =   2445
                  Left            =   4320
                  TabIndex        =   30
                  Top             =   0
                  Width           =   915
                  _cx             =   1614
                  _cy             =   4313
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
                  BackColorSel    =   16761024
                  ForeColorSel    =   0
                  BackColorBkg    =   -2147483636
                  BackColorAlternate=   -2147483643
                  GridColor       =   -2147483633
                  GridColorFixed  =   -2147483632
                  TreeColor       =   -2147483632
                  FloodColor      =   192
                  SheetBorder     =   -2147483642
                  FocusRect       =   1
                  HighLight       =   1
                  AllowSelection  =   0   'False
                  AllowBigSelection=   -1  'True
                  AllowUserResizing=   0
                  SelectionMode   =   1
                  GridLines       =   2
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   10
                  Cols            =   3
                  FixedRows       =   0
                  FixedCols       =   0
                  RowHeightMin    =   360
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
            End
         End
         Begin VB.Frame fraFuncParas 
            Caption         =   "�����б�"
            Height          =   3495
            Left            =   240
            TabIndex        =   15
            Top             =   3840
            Width           =   4215
            Begin VB.CommandButton cmdDelPara 
               Caption         =   "ɾ������"
               Height          =   375
               Left            =   1680
               TabIndex        =   17
               Top             =   3000
               Width           =   1215
            End
            Begin VB.CommandButton cmdAddPara 
               Caption         =   "��Ӳ���"
               Height          =   375
               Left            =   120
               TabIndex        =   16
               Top             =   3000
               Width           =   1215
            End
            Begin VSFlex8Ctl.VSFlexGrid vsfAppfunPara 
               Height          =   2460
               Left            =   180
               TabIndex        =   18
               Top             =   480
               Width           =   3060
               _cx             =   5397
               _cy             =   4339
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
               ForeColorSel    =   0
               BackColorBkg    =   -2147483636
               BackColorAlternate=   -2147483643
               GridColor       =   -2147483633
               GridColorFixed  =   -2147483632
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483642
               FocusRect       =   1
               HighLight       =   1
               AllowSelection  =   0   'False
               AllowBigSelection=   -1  'True
               AllowUserResizing=   0
               SelectionMode   =   1
               GridLines       =   2
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   50
               Cols            =   3
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   360
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
               Begin VB.CommandButton cmdConfigWindow 
                  Caption         =   "����"
                  Height          =   375
                  Left            =   2280
                  TabIndex        =   27
                  Top             =   1080
                  Visible         =   0   'False
                  Width           =   735
               End
            End
         End
      End
      Begin VB.Frame fraAppInfo 
         Caption         =   "������Ϣ"
         Height          =   1150
         Left            =   180
         TabIndex        =   5
         Top             =   120
         Width           =   8955
         Begin VB.TextBox txtAppName 
            Height          =   315
            Left            =   1140
            TabIndex        =   24
            Top             =   300
            Width           =   2895
         End
         Begin VB.ComboBox cboType 
            Height          =   300
            ItemData        =   "frmPacsInterfaceCfg.frx":6852
            Left            =   1140
            List            =   "frmPacsInterfaceCfg.frx":685F
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   720
            Width           =   2895
         End
         Begin VB.ComboBox cboClasses 
            Height          =   300
            Left            =   5220
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   720
            Width           =   2355
         End
         Begin VB.CheckBox chkUseThisApp 
            Caption         =   "����"
            Height          =   255
            Left            =   7800
            TabIndex        =   8
            Top             =   740
            Width           =   675
         End
         Begin VB.TextBox txtAppDir 
            Height          =   350
            Left            =   5220
            Locked          =   -1  'True
            TabIndex        =   7
            Tag             =   "VBS��̬�ű�"
            Top             =   285
            Width           =   3390
         End
         Begin VB.CommandButton cmdSelectApp 
            Caption         =   "��"
            Height          =   350
            Left            =   8595
            TabIndex        =   6
            Top             =   270
            Width           =   260
         End
         Begin VB.Label lblAppName 
            AutoSize        =   -1  'True
            Caption         =   "������ƣ�"
            Height          =   180
            Left            =   240
            TabIndex        =   25
            Top             =   360
            Width           =   900
         End
         Begin VB.Label lblType 
            AutoSize        =   -1  'True
            Caption         =   "ִ�����ͣ�"
            Height          =   180
            Left            =   240
            TabIndex        =   13
            Top             =   760
            Width           =   900
         End
         Begin VB.Label lblAppDir 
            AutoSize        =   -1  'True
            Caption         =   "����·����"
            Height          =   180
            Left            =   4140
            TabIndex        =   12
            Top             =   360
            Width           =   900
         End
         Begin VB.Label lblClasses 
            AutoSize        =   -1  'True
            Caption         =   "���򼯺ϣ�"
            Height          =   180
            Left            =   4140
            TabIndex        =   11
            Top             =   780
            Width           =   900
         End
      End
   End
   Begin VB.PictureBox picApp 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7755
      Left            =   360
      ScaleHeight     =   7725
      ScaleWidth      =   4545
      TabIndex        =   0
      Top             =   1680
      Width           =   4575
      Begin VB.ComboBox cboStation 
         Height          =   300
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1260
         Width           =   2115
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfApp 
         Height          =   3840
         Left            =   960
         TabIndex        =   2
         Top             =   2280
         Width           =   3540
         _cx             =   6244
         _cy             =   6773
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
         ForeColorSel    =   0
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   2
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   360
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
      Begin VB.Label lblStation 
         AutoSize        =   -1  'True
         Caption         =   "����վ��"
         Height          =   180
         Left            =   780
         TabIndex        =   3
         Top             =   1320
         Width           =   720
      End
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   780
      Top             =   660
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   29
      Top             =   9795
      Width           =   16020
      _ExtentX        =   28258
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   4154
            MinWidth        =   4154
            Picture         =   "frmPacsInterfaceCfg.frx":6881
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14288
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
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
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   2880
      Top             =   900
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmPacsInterfaceCfg.frx":7115
      Left            =   2100
      Top             =   1140
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPacsInterfaceCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const C_STR_CUSTOMPARAS = "[[ϵͳ��]]|[[ģ���]]|[[����ID]]|[[����ID]]|[[ҽ��ID]]|[[����]]|[[�����]]|[[סԺ��]]|[[���֤��]]|[[Ӱ�����]]|[[�û���]]|[[�˺���]]|[[��ǰ���ھ��]]"

Private Enum mAppCol
    ��� = 0: ��������: ����汾: ����·��: ����ID: ����: ִ������: �Ƿ�����: ����ģ��
End Enum

Private Enum mAppFuncCol
    ��� = 0: ��������:  ���ù���: �����Ҽ��˵�: ���빤����: �Զ�ִ��ʱ��: ��Ӧ����: ��������: VBS�ű�: ����ID: ��֤ͨ��
End Enum

Private Enum mAppFuncsCol
    ������� = 0: ���ܷ���: ��������
End Enum

Private Enum mAppFuncParaCol
    ��� = 0: ��������: ��������: ��������
End Enum

Private Enum mExecuteType
    ��̬���� = 1: Shell����: API����
End Enum

'�˵�����ö�ٶ���
Private Enum TMenuType
    mtFile = 1
    mtSave
    mtCancel
    mtQuit
    
    mtEdit
    mtAdd
    mtMod
    mtDel
    mtUse
    mtCheck
    mtRefresh
End Enum

Private mintTestSta As Integer  '����״̬ 0 δ���� 1 ͨ��  2 δͨ��
Private mstr�������� As String
Private mblnConfiging As Boolean
Private mblnIsAddCfg As Boolean

Private mlngModule As Long
Private mstrPrivs As String
Private mlngAdviceID As Long
Private mlngSendNo As Long
Private mlngPatId As Long

Public Function ShowPacsInterfaceCfg(objParent As Object, ByVal lngModule As Long, ByVal strPrivs As String, _
                                ByVal lngAdviceID As Long, ByVal lngSendNO As Long, ByVal lngPatId As Long) As Boolean
    
    mlngModule = lngModule
    mstrPrivs = strPrivs
    mlngAdviceID = lngAdviceID
    mlngSendNo = lngSendNO
    mlngPatId = lngPatId
    
    Call Me.Show(1, objParent)
End Function

Private Sub cboClasses_Click()
On Error GoTo ErrorHand
    
    Call FuncsFaceEnabled(mblnConfiging)
    Call ParasFaceEnabled(mblnConfiging)
    Call LoadAllClassFunc(txtAppDir.Text, cboClasses.Text)
    
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
    err.Clear
End Sub

Private Sub cboStation_Click()
On Error GoTo ErrorHand
    
    Call ClearAllCfg
    Call LoadAppInfo
    
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
    err.Clear
End Sub

Private Sub cbrMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
On Error GoTo ErrorHnad
    
    Select Case Control.ID
        Case TMenuType.mtSave
            Call SaveAppCfg
            
        Case TMenuType.mtCancel
            Call CancelAppCfg
            
        Case TMenuType.mtAdd
            Call AddAppCfg
            
        Case TMenuType.mtMod
            Call ModAppCfg
            
        Case TMenuType.mtDel
            Call DelAppCfg
            
        Case TMenuType.mtUse
            Call UseAppCfg(Control)
            
        Case TMenuType.mtRefresh
            Call ClearAllCfg
            Call LoadAppInfo
            
        Case TMenuType.mtQuit
            Call Unload(Me)
            
'---------------------------�鿴----------------
        Case conMenu_View_ToolBar_Button '������
            Call Menu_View_ToolBar_Button_click(Control)
        Case conMenu_View_ToolBar_Text '��ť����
            Call Menu_View_ToolBar_Text_click(Control)
        Case conMenu_View_ToolBar_Size '��ͼ��
            Call Menu_View_ToolBar_Size_click(Control)
        Case conMenu_View_StatusBar '״̬��
            Call Menu_View_StatusBar_click(Control)

'--------------------------����-----------------
        Case conMenu_Help_Help
            Call Menu_Help_Help_click
        Case conMenu_Help_Web_Forum
            Call Menu_Help_Web_Forum_click
        Case conMenu_Help_Web_Home
            Call Menu_Help_Web_Home_click
        Case conMenu_Help_Web_Mail
            Call Menu_Help_Web_Mail_click
        Case conMenu_Help_About
            Call Menu_Help_About_click
    
    End Select
    
    Exit Sub
ErrorHnad:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub

Private Function ValidData() As Boolean
'------------------------------------------------
'���ܣ�����������ݵĺϷ���
'������ ��
'���أ�True--��������ϸ񣬿��Լ�����False --���������벻�ϸ���Ҫ�޸�����
'------------------------------------------------
On Error GoTo ErrorHnad
    
    ValidData = False
    
    '������Ϣ
    If Trim(txtAppName.Text) = "" Then
        MsgBox "�������Ʋ���Ϊ�գ������룡", vbExclamation, gstrSysName
        txtAppName.SetFocus
        Exit Function
        
    ElseIf Trim(txtAppDir.Text) = "" Then
        MsgBox "����·������Ϊ�գ������룡", vbExclamation, gstrSysName
        txtAppDir.SetFocus
        Exit Function
        
    ElseIf cboType.ListIndex < 0 Then
        MsgBox "����ִ�����Ͳ���Ϊ�գ������룡", vbExclamation, gstrSysName
        cboType.SetFocus
        Exit Function
        
    ElseIf cboClasses.ListIndex < 0 Then
        MsgBox "���򼯺ϲ���Ϊ�գ������룡", vbExclamation, gstrSysName
        cboClasses.SetFocus
        Exit Function
    End If
    
    If CheckAppFuns() = False Then Exit Function
    If CheckAppfunPara() = False Then Exit Function
    
    ValidData = True
    
    Exit Function
ErrorHnad:
    MsgBox err.Description, vbExclamation, gstrSysName
End Function

Private Function CheckAppFuns() As Boolean
'���ܲ����б�
    Dim i As Integer
    
    CheckAppFuns = True
    
    With vsfAppFuns
        If .Rows <= 1 Then
            MsgBox "��������ع��ܣ�", vbExclamation, gstrSysName
            CheckAppFuns = False
            Exit Function
        End If
        
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, mAppFuncCol.��������)) = "" Then
                MsgBox "�������Ʋ���Ϊ�գ������룡", vbExclamation, gstrSysName
                CheckAppFuns = False
                Exit Function
            End If
            
            If Trim(vsfFuncs.TextMatrix(0, 0)) = "" Then
                MsgBox "���ܶ�Ӧ��������Ϊ�գ������룡", vbExclamation, gstrSysName
                CheckAppFuns = False
                Exit Function
            End If
        Next
    End With
End Function

Private Function CheckAppfunPara() As Boolean
'���������б�
    Dim i As Integer
    
    CheckAppfunPara = True
    
    With vsfAppfunPara
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, mAppFuncParaCol.��������)) = "" Then
                MsgBox "�������Ͳ���Ϊ�գ������룡", vbExclamation, gstrSysName
                CheckAppfunPara = False
                Exit Function
            End If
            
            If Trim(.TextMatrix(i, mAppFuncParaCol.��������)) = "" And Trim(.TextMatrix(i, mAppFuncParaCol.��������)) <> "�ַ���" Then
                MsgBox "�������첻��Ϊ�գ������룡", vbExclamation, gstrSysName
                CheckAppfunPara = False
                Exit Function
            End If
        Next
    End With
End Function

Private Sub SaveAppCfg()
'------------------------------------------------
'���ܣ�����������Ϣ
'��������
'���أ���
'------------------------------------------------
    Dim i As Integer
    Dim lngAppId As Long
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strParaInfo As String
    Dim strVBS As String
    Dim intType As Integer 'ִ��ʱ��
    
    
    '������Ϣ����Ч�Լ��
    If Not ValidData Then Exit Sub

    '�ж�VBS�Ƿ��޸Ĺ������޸Ĺ�Ҫ�ȸ��µ��б���
    If chkModify.value = 1 Then
        If Not (txtVBS.Text = vsfAppFuns.TextMatrix(vsfAppFuns.RowSel, mAppFuncCol.VBS�ű�)) Then
            If CheckAppCfg(txtVBS.Text) Then
                vsfAppFuns.TextMatrix(vsfAppFuns.RowSel, mAppFuncCol.VBS�ű�) = txtVBS.Text
            Else
                Exit Sub
            End If
        End If
    End If
    
    If Not DoBeforeSave() Then Exit Sub
    
    mblnConfiging = False
    
    Call InputFaceEnabled(False)
    Call AppFaceEnabled(True)
    chkModify.value = 0
    
    '����ʱ����ȡ����ID
    If mblnIsAddCfg Then
        strSql = "Select Nvl(Max(ID), 0) + 1 as ����ID From Ӱ�����ҽ�"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "")
        If rsTemp.RecordCount > 0 Then lngAppId = Val(rsTemp!����ID)
    Else
        lngAppId = vsfApp.TextMatrix(vsfApp.RowSel, mAppCol.����ID)
    End If
    
    '����Ӱ�����ҽ���Ϣ
    strSql = "ZL_Ӱ�����ҽ�_Update(" & lngAppId & ",'" & _
                                         txtAppName.Text & "','" & _
                                         txtAppDir.tag & "','" & _
                                         txtAppDir.Text & "','" & _
                                         cboClasses.Text & "'," & _
                                         cboType.ItemData(cboType.ListIndex) & "," & _
                                         chkUseThisApp.value & "," & _
                                         cboStation.ItemData(cboStation.ListIndex) & ")"
    
    Call zlDatabase.ExecuteProcedure(strSql, "")
    
    '�������������Ϣ
    strSql = "ZL_Ӱ��������_Delete(" & lngAppId & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "")
    
    For i = 1 To vsfAppFuns.Rows - 1
        intType = CInt(convertInterfaceTime(vsfAppFuns.TextMatrix(i, mAppFuncCol.�Զ�ִ��ʱ��), True))
        strSql = "ZL_Ӱ��������_Update(" & lngAppId & ",'" & _
                                             vsfAppFuns.TextMatrix(i, mAppFuncCol.��������) & "','" & _
                                             vsfAppFuns.TextMatrix(i, mAppFuncCol.��Ӧ����) & "','" & _
                                             vsfAppFuns.TextMatrix(i, mAppFuncCol.��������) & "'," & _
                                             IIf(vsfAppFuns.Cell(flexcpChecked, i, mAppFuncCol.���ù���) = 1, 1, 0) & "," & _
                                             IIf(vsfAppFuns.Cell(flexcpChecked, i, mAppFuncCol.�����Ҽ��˵�) = 1, 1, 0) & "," & _
                                             IIf(vsfAppFuns.Cell(flexcpChecked, i, mAppFuncCol.���빤����) = 1, 1, 0) & "," & _
                                             intType & ",'" & _
                                             vsfAppFuns.TextMatrix(i, mAppFuncCol.VBS�ű�) & "')"
        
        Call zlDatabase.ExecuteProcedure(strSql, "")
    Next
    
    If mblnIsAddCfg Then
        For i = 1 To vsfApp.Rows - 1
            With vsfApp
                If vsfApp.TextMatrix(i, mAppCol.��������) = "" Then
                    .TextMatrix(i, mAppCol.��������) = txtAppName.Text
                    .TextMatrix(i, mAppCol.����汾) = txtAppDir.tag
                    .TextMatrix(i, mAppCol.����·��) = txtAppDir.Text
                    .TextMatrix(i, mAppCol.����ID) = lngAppId
                    .TextMatrix(i, mAppCol.����) = cboClasses.Text
                    .TextMatrix(i, mAppCol.ִ������) = cboType.ItemData(cboType.ListIndex)
                    .TextMatrix(i, mAppCol.�Ƿ�����) = chkUseThisApp.value
                    .TextMatrix(i, mAppCol.����ģ��) = cboStation.ItemData(cboStation.ListIndex)
                    
                    Exit For
                End If
            End With
        Next
    Else
         With vsfApp
            .TextMatrix(.RowSel, mAppCol.��������) = txtAppName.Text
            .TextMatrix(.RowSel, mAppCol.����汾) = txtAppDir.tag
            .TextMatrix(.RowSel, mAppCol.����·��) = txtAppDir.Text
            .TextMatrix(.RowSel, mAppCol.����ID) = lngAppId
            .TextMatrix(.RowSel, mAppCol.����) = cboClasses.Text
            .TextMatrix(.RowSel, mAppCol.ִ������) = cboType.ItemData(cboType.ListIndex)
            .TextMatrix(.RowSel, mAppCol.�Ƿ�����) = chkUseThisApp.value
            .TextMatrix(.RowSel, mAppCol.����ģ��) = cboStation.ItemData(cboStation.ListIndex)
        End With
    End If
End Sub

Private Sub CancelAppCfg()
    mblnConfiging = False
    If mblnIsAddCfg Then txtAppName.Text = ""
    
    Call vsfApp_SelChange
    chkModify = 0
    Call InputFaceEnabled(False)
    Call AppFaceEnabled(True)
    
End Sub

Private Sub AddAppCfg()
    mblnConfiging = True
    mblnIsAddCfg = True
    txtAppName.Text = ""
    
    Call ClearInputCfg
    Call AppFaceEnabled(False)
    Call AppInfoFaceEnabled(True)
End Sub

Private Sub ModAppCfg()
    mblnConfiging = True
    mblnIsAddCfg = False
    
    Call AppFaceEnabled(False)
    Call InputFaceEnabled(True)
    cboType.Enabled = False
End Sub

Private Sub DelAppCfg()
    Dim strSql As String
    Dim lngAppId As Long
    
    lngAppId = vsfApp.TextMatrix(vsfApp.RowSel, mAppCol.����ID)
    If lngAppId <= 0 Then Exit Sub
    
    If MsgBox("ȷ��Ҫɾ���˳���������", vbExclamation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    strSql = "ZL_Ӱ�����ҽ�_Delete(" & lngAppId & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "")
    
    Call ReLoadVSFList(vsfApp)
    Call vsfApp_SelChange
End Sub

Private Sub ReLoadVSFList(vsfList As VSFlexGrid)
    Dim i As Integer, j As Integer
    
    On Error GoTo ErrorHand
    
    If vsfList Is Nothing Then Exit Sub
    
    With vsfList
        For i = vsfList.RowSel To vsfList.Rows - 2
            For j = 1 To vsfList.Cols - 1
                vsfList.TextMatrix(i, j) = vsfList.TextMatrix(i + 1, j)
            Next
        Next
    End With
    
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
    err.Clear
End Sub

Private Sub UseAppCfg(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngAppId As Long
    Dim strSql As String
    
    lngAppId = vsfApp.TextMatrix(vsfApp.RowSel, mAppCol.����ID)
    
    If lngAppId <= 0 Then Exit Sub
    
    chkUseThisApp.value = IIf(Control.Caption = "����", 1, 0)
    vsfApp.TextMatrix(vsfApp.RowSel, mAppCol.�Ƿ�����) = IIf(Control.Caption = "����", 1, 0)
    
    strSql = "ZL_Ӱ�����ҽ�_Update(" & lngAppId & ",'" & _
                                         txtAppName.Text & "','" & _
                                         txtAppDir.tag & "','" & _
                                         txtAppDir.Text & "','" & _
                                         cboClasses.Text & "'," & _
                                         cboType.ItemData(cboType.ListIndex) & "," & _
                                         chkUseThisApp.value & "," & _
                                         cboStation.ItemData(cboStation.ListIndex) & ")"
    
    Call zlDatabase.ExecuteProcedure(strSql, "")
End Sub

Private Function CheckAppCfg(ByVal strVBS As String, Optional ByVal blTest As Boolean = False) As Boolean
'����vbs�ű�ʵ�ֹ���,blTest �Ƿ��ܲ��ԣ����ǣ�����Ҫ�����Ĳ��Թ��ܡ�
    Dim i As Integer
    Dim lngStart As Long, lngEnd As Long
    Dim ary() As String
    Dim strTmpVBS As String, strParaName As String, strParaVal As String
    Dim objCall As Object
    
On Error GoTo ErrorHnad
    
    ary = Split(strVBS, vbCrLf)
    
    For i = 0 To UBound(ary)
        '����Ԥ����������ڲ���ֵ
        strTmpVBS = ary(i)
        
        Do While InStr(strTmpVBS, "[[") > 0
            lngStart = InStr(strTmpVBS, "[[")
            lngEnd = InStr(strTmpVBS, "]]") + 2
            
            strParaName = Mid(strTmpVBS, lngStart, lngEnd - lngStart)
            
            Select Case strParaName
                Case "[[�û���]]"
                    strParaVal = "ZLHIS"
                                        
                Case "[[�˺���]]"
                    strParaVal = "ZLHIS"
                                        
                Case "[[ϵͳ��]]"
                    strParaVal = "100"
                    
                Case "[[ģ���]]"
                    strParaVal = "1291"
                    
                Case "[[����ID]]"
                    strParaVal = "64"
                
                Case "[[����ID]]"
                    strParaVal = "1"
                    
                Case "[[ҽ��ID]]"
                    strParaVal = "101"
                    
                Case "[[����]]"
                    strParaVal = "110"
                    
                Case "[[�����]]"
                    strParaVal = "1"
                
                Case "[[סԺ��]]"
                    strParaVal = "110"
                    
                Case "[[���֤��]]"
                    strParaVal = "500105190001010000"
                    
                Case "[[Ӱ�����]]"
                    strParaVal = "CT"
                                        
                Case "[[��ǰ���ھ��]]"
                    strParaVal = Me.hWnd
                                        
                Case Else
                    MsgBox "���ֲ���ʶ���Ԥ�������������", vbExclamation, gstrSysName
                    CheckAppCfg = False
                    Exit Function
            End Select
            
            If strParaVal <> "------" Then strVBS = Replace(strVBS, strParaName, strParaVal)
            
            strTmpVBS = Trim(Mid(strTmpVBS, lngEnd))
        Loop
    Next
    
    CheckAppCfg = ExecuteSub(strVBS, Me, blTest)
    
    Exit Function
ErrorHnad:
    MsgBox err.Description, vbExclamation, gstrSysName
    CheckAppCfg = False
End Function

Public Function ExecuteSub(ByVal strVBS As String, ByVal objParent As Object, Optional ByVal blTest As Boolean = False) As Boolean
'����vbs�ű�ʵ�ֹ���
    Dim objCall As Object
    Dim strTempVBS As String
    
On Error GoTo ErrorHnad
    
    '�����ű�ִ�ж���
    Set objCall = CreateObject("ScriptControl")
    objCall.Timeout = 60000
    objCall.Language = "vbscript"
    
    Call objCall.AddCode(strVBS)
    
    If blTest Then
        Call objCall.Run(Trim("ExcuteSub"))
    End If
    ExecuteSub = True
    
    Exit Function
ErrorHnad:
    MsgBox err.Description, vbExclamation, gstrSysName
End Function

Private Sub cbrMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
On Error GoTo ErrorHnad
    
    Select Case Control.ID
        Case TMenuType.mtRefresh
            Control.Enabled = Not mblnConfiging
                        
        Case TMenuType.mtSave
            Control.Enabled = mblnConfiging
            
        Case TMenuType.mtCancel
            Control.Enabled = mblnConfiging
            
        Case TMenuType.mtAdd
            Control.Enabled = Not mblnConfiging
            
        Case TMenuType.mtMod
            Control.Enabled = vsfApp.RowSel > 0 And vsfApp.TextMatrix(vsfApp.RowSel, mAppCol.��������) <> "" And Not mblnConfiging
            
        Case TMenuType.mtDel
            Control.Enabled = vsfApp.RowSel > 0 And vsfApp.TextMatrix(vsfApp.RowSel, mAppCol.��������) <> "" And Not mblnConfiging
            
        Case TMenuType.mtUse
            Control.Enabled = vsfApp.RowSel > 0 And vsfApp.TextMatrix(vsfApp.RowSel, mAppCol.��������) <> "" And Not mblnConfiging
            Control.Caption = IIf(Val(vsfApp.TextMatrix(vsfApp.RowSel, mAppCol.�Ƿ�����)) = 1, "����", "����")
                        
            Control.IconId = IIf(Val(vsfApp.TextMatrix(vsfApp.RowSel, mAppCol.�Ƿ�����)) = 1, 3006, 3009)
            
            '���ڼ�ʱˢ�°�ť״̬
            Control.Enabled = Not Control.Enabled
            Control.Enabled = Not Control.Enabled
        
        Case TMenuType.mtQuit
            Control.Enabled = Not mblnConfiging
    End Select
    
    Exit Sub
ErrorHnad:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub

Private Sub chkModify_Click()
    txtVBS.Enabled = (chkModify.value = 1)
    cmdDelFun.Enabled = (chkModify.value = 0)
End Sub

Private Sub cmdAddFunc_Click()
    Dim i As Integer

On Error GoTo ErrorHand

    '�������
    If vsfAppFuns.Rows > 1 Then
        If Trim(vsfAppFuns.TextMatrix(vsfAppFuns.Rows - 1, mAppFuncCol.��������)) = "" Then
            MsgBox "�������Ʋ���Ϊ�գ�������!", vbExclamation, gstrSysName
            Exit Sub
        ElseIf Trim(vsfAppFuns.TextMatrix(vsfAppFuns.Rows - 1, mAppFuncCol.��Ӧ����)) = "" Then
            MsgBox "���ܷ�������Ϊ�գ�������!", vbExclamation, gstrSysName
            Exit Sub
        End If
        
        '�Թ��ܵĲ������м��
        If vsfAppfunPara.Rows <= 1 Then
            'û�����ò���������Ҫ������
            If cboType.ItemData(cboType.ListIndex) <> mExecuteType.��̬���� Then
                If MsgBox("��û�жԴ˹��ܷ������в������ã�ȷ��Ҫ������", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
            End If
        End If
        
        '��VBS�ű�������֤
        If Not CheckAppCfg(txtVBS.Text) Then Exit Sub
    End If
    
    vsfAppFuns.Rows = vsfAppFuns.Rows + 1
    vsfAppFuns.TextMatrix(vsfAppFuns.Rows - 1, mAppFuncCol.���) = vsfAppFuns.Rows - 1
    vsfAppFuns.TextMatrix(vsfAppFuns.Rows - 1, mAppFuncCol.���ù���) = 1
    vsfAppFuns.Select vsfAppFuns.Rows - 1, 1
    vsfAppFuns.EditCell
    
    If cboType.Text = "Shell����" Then vsfFuncs.ColComboList(mAppFuncsCol.���ܷ���) = txtAppDir.Text
    
    Call ParasFaceEnabled(True)
    Call FuncsFaceEnabled(True)
    Call VBSFaceEnabled(True)
    Call ClearFuncs
    
    '��ӹ��ܺ󣬽����������ִ������
    cboType.Enabled = False
    
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
    err.Clear
End Sub

Private Sub cmdAddPara_Click()
On Error GoTo ErrorHand
    
    vsfAppfunPara.Rows = vsfAppfunPara.Rows + 1
    vsfAppfunPara.TextMatrix(vsfAppfunPara.Rows - 1, mAppFuncParaCol.���) = vsfAppfunPara.Rows - 1
    vsfAppfunPara.Select vsfAppfunPara.Rows - 1, 1
    vsfAppfunPara.EditCell
    vsfAppfunPara.Cell(flexcpAlignment, 0, 0, vsfAppfunPara.Rows - 1, vsfAppfunPara.Cols - 1) = flexAlignLeftCenter
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
    err.Clear
End Sub

Private Sub cmdConfigWindow_Click()
    Dim strResult As String
    Dim strTxtValueOld
    
    On Error GoTo ErrorHand
    
    strTxtValueOld = vsfAppfunPara.TextMatrix(vsfAppfunPara.RowSel, mAppFuncParaCol.��������)
    strResult = frmPacsInterfaceParEdit.EditPara(vsfAppfunPara.TextMatrix(vsfAppfunPara.RowSel, mAppFuncParaCol.��������), Me, (mblnConfiging And chkModify.value = 0))
    
    If mblnConfiging Then
        vsfAppfunPara.TextMatrix(vsfAppfunPara.RowSel, mAppFuncParaCol.��������) = strResult
        Call vsfAppfunPara_AfterEdit(0, mAppFuncParaCol.��������)
    End If
    
    Call RefreshCfg
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
    err.Clear
End Sub

Private Sub cmdDelFun_Click()
On Error GoTo ErrorHand
    Dim LngSel As Integer

    LngSel = vsfFuncs.Row

    If vsfFuncs.RowSel + 1 > vsfFuncs.Rows Then Exit Sub
    If vsfFuncs.TextMatrix(vsfFuncs.RowSel, mAppFuncsCol.���ܷ���) = "" Then Exit Sub
    
    vsfFuncs.TextMatrix(vsfFuncs.RowSel, mAppFuncsCol.��������) = ""
    vsfFuncs.TextMatrix(vsfFuncs.RowSel, mAppFuncsCol.���ܷ���) = ""
    
    Call ReLoadVSFList(vsfFuncs)
    
    vsfFuncs.RowSel = 0
    vsfFuncs.Row = LngSel
    Call LoadFuncParaCfg
    Call DoFraFuncParasCaption

    Call RefreshCfg
    Call CreateVBS
    
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
    err.Clear
End Sub

Private Sub cmdDelFunc_Click()
On Error GoTo ErrorHand
    Dim i As Long, j As Long
    Dim lngRow As Long
    Dim blOneRow As Boolean '�Ƿ�ɾ��ǰֻ��һ������
'''''ɾ������ǰ����ѡ����һ����Ч�Ĺ��ܣ��ڹ����б��У������������ݣ���ѡ�к�һ�����ݣ����ҴӺ�һ�����ݿ�ʼ��������һ�С�������û�����ݣ���ѡ��ǰһ������
''''���Ѿ���Ψһ���ݣ�����ʾ������ɾ������ֱ���޸�

    blOneRow = False
    If vsfAppFuns.Rows = 1 Then Exit Sub
    If vsfAppFuns.Rows = 2 Then blOneRow = True

    
    lngRow = vsfAppFuns.Row
    If lngRow = 0 Then lngRow = 1
    
    Call vsfAppFuns.RemoveItem(lngRow)
    
    If Not blOneRow Then
    'ɾ��ǰ��ֻһ������
        For i = lngRow To vsfAppFuns.Rows - 1
            vsfAppFuns.TextMatrix(i, mAppFuncsCol.�������) = i
        Next
    
        If lngRow = vsfAppFuns.Rows Then
            '�Ѿ������һ����ѡ��ǰ��һ��
            vsfAppFuns.Row = lngRow - 1
            vsfAppFuns.RowSel = lngRow - 1
        Else
            '�������һ����ѡ��ǰ
            vsfAppFuns.Row = lngRow
            vsfAppFuns.RowSel = lngRow
        End If
    
        Call vsfAppFuns_SelChange
        Call ParasFaceEnabled(vsfAppFuns.Rows > 1)
        Call VBSFaceEnabled(vsfAppFuns.Rows > 1)
    Else
    'ɾ��ǰֻ��һ������

        Call ClearFuncs
        Call ClearAppFuncParaCfg
        Call DoFraFuncParasCaption
        txtVBS.Text = ""
        Call ParasFaceEnabled(vsfAppFuns.Rows > 1)
        Call VBSFaceEnabled(vsfAppFuns.Rows > 1)
        Call FuncsFaceEnabled(mblnConfiging)
    End If
    
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
    err.Clear
End Sub

Private Sub cmdDelPara_Click()
    If vsfAppfunPara.Rows <= 1 Then Exit Sub
    
    Call vsfAppfunPara.RemoveItem(vsfAppfunPara.Rows - 1)
    Call RefreshCfg
    
End Sub

Private Sub cmdSelectApp_Click()
On Error GoTo ErrorHand
    Dim strFilePath As String, strFileName As String
    
    dlgFile.Filter = "(*.exe)|*.exe|(*.dll)|*.dll|(*.ocx)|*.ocx|(*.tlb)|*.tlb|(*.*)|*.*"
    dlgFile.ShowOpen
    
    strFilePath = dlgFile.Filename
    strFileName = dlgFile.FileTitle
    cboType.Enabled = True
    
    Call FuncsFaceEnabled(False)
    Call ParasFaceEnabled(False)
    Call VBSFaceEnabled(False)
    Call ClearInputCfg
    Call ClearFuncs
    
    Call LoadAppConfig(strFilePath, strFileName)
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
    err.Clear
End Sub

Private Sub LoadAppConfig(ByVal strFilePath As String, ByVal strFileName As String)
    Dim objFSO As New FileSystemObject
    
    If strFilePath = "" Then Exit Sub
    
    txtAppDir.Text = strFilePath
    
    txtAppDir.tag = objFSO.GetFileVersion(strFilePath)
    cboClasses.tag = strFileName
    
    vsfFuncs.ColComboList(mAppFuncsCol.���ܷ���) = ""
    
    '���س�������й��ܺͲ���
    Call LoadAllClass(strFilePath, strFileName)
    Call picAppCfg_Resize
End Sub

Private Sub LoadAllClass(ByVal strFilePath As String, ByVal strFileName As String)
'����dll��ocx��������������ĳ���
    Dim i As Integer
    Dim objClassInfo As TypeLibInfo
    Dim objInterfaceInfo As InterfaceInfo

On Error GoTo ErrorHand
    cboClasses.Clear
    
    Set objClassInfo = TypeLibInfoFromFile(strFilePath)
    
    cboType.Clear
    cboType.AddItem "��̬����"
    cboType.ItemData(cboType.NewIndex) = mExecuteType.��̬����
    cboType.Text = "��̬����"
    cboType.Enabled = False
    cboClasses.Enabled = True
    
    For Each objInterfaceInfo In objClassInfo.Interfaces
        If Not objInterfaceInfo.VTableInterface Is Nothing Then
            cboClasses.AddItem objInterfaceInfo.Parent & "." & Mid(objInterfaceInfo.Name, 2)
            If objInterfaceInfo.Parent & "." & Mid(objInterfaceInfo.Name, 2) = strFileName Then
                cboClasses.ListIndex = cboClasses.NewIndex
            End If
        End If
    Next
    
    Exit Sub
ErrorHand:
    cboType.Clear
    cboType.AddItem "Shell����"
    cboType.ItemData(cboType.NewIndex) = mExecuteType.Shell����
    cboType.AddItem "API����"
    cboType.ItemData(cboType.NewIndex) = mExecuteType.API����
    
    If vsfApp.TextMatrix(vsfApp.RowSel, mAppCol.��������) <> "" And vsfApp.RowSel > 0 Then
        If vsfApp.TextMatrix(vsfApp.RowSel, mAppCol.ִ������) = 2 Then
            cboType.ListIndex = 0
        Else
            cboType.ListIndex = 1
        End If
    Else
        cboType.ListIndex = 0
        cboType.Enabled = True
    End If
    
    cboClasses.AddItem strFileName
    cboClasses.Text = strFileName
    cboClasses.Enabled = False
End Sub

Private Sub LoadAllClassFunc(ByVal strFileName As String, ByVal strClassName As String)
'���ݳ��򼯼��ض�Ӧ�ķ���
    Dim objClassInfo As TypeLibInfo
    Dim objInterfaceInfo As InterfaceInfo
    Dim objMemberInfo As MemberInfo
    Dim strFuncs As String
    
    If cboType.ItemData(cboType.ListIndex) = mExecuteType.��̬���� Then
        Set objClassInfo = TypeLibInfoFromFile(strFileName)
        
        For Each objInterfaceInfo In objClassInfo.Interfaces
            If Not objInterfaceInfo.VTableInterface Is Nothing Then
                If objInterfaceInfo.Parent & "." & Mid(objInterfaceInfo.Name, 2) = strClassName Then
                    For Each objMemberInfo In objInterfaceInfo.Members
                        If objMemberInfo.InvokeKind = INVOKE_FUNC Then
                            '����Ƿ�������ص������б�
                            If objMemberInfo.Name <> "QueryInterface" And objMemberInfo.Name <> "AddRef" _
                                And objMemberInfo.Name <> "Release" And objMemberInfo.Name <> "GetTypeInfoCount" _
                                And objMemberInfo.Name <> "GetTypeInfo" And objMemberInfo.Name <> "GetIDsOfNames" _
                                And objMemberInfo.Name <> "Invoke" Then
                                strFuncs = strFuncs & "|" & objMemberInfo.Name
                            End If
                        End If
                    Next
                End If
            End If
        Next
        
        vsfFuncs.ColComboList(mAppFuncsCol.���ܷ���) = IIf(strFuncs <> "", Mid(strFuncs, 2), "")
    Else
        vsfFuncs.ColComboList(mAppFuncsCol.���ܷ���) = ""
        If cboType.Text = "Shell����" Then vsfFuncs.ColComboList(mAppFuncsCol.���ܷ���) = txtAppDir.Text
    End If
End Sub

Private Sub LoadParasWithFunc(ByVal strFileName As String, ByVal strClassName As String, ByVal strFuncName As String)
'����ѡ��ķ�����ȡ��Ӧ�Ĳ���
    Dim objClassInfo As TypeLibInfo
    Dim objInterfaceInfo As InterfaceInfo
    Dim objMemberInfo As MemberInfo
    Dim objParameterInfo As ParameterInfo
    Dim strParas As String
    
    Set objClassInfo = TypeLibInfoFromFile(strFileName)
    
    For Each objInterfaceInfo In objClassInfo.Interfaces
        If Not objInterfaceInfo.VTableInterface Is Nothing Then
            If objInterfaceInfo.Parent & "." & Mid(objInterfaceInfo.Name, 2) = strClassName Then
                For Each objMemberInfo In objInterfaceInfo.Members
                    If objMemberInfo.InvokeKind = INVOKE_FUNC Then
                        '����Ƿ���
                        If objMemberInfo.Name = strFuncName Then
                            For Each objParameterInfo In objMemberInfo.Parameters
                                Call AddParaToParasList(objParameterInfo.Name)
                            Next
                        End If
                    End If
                Next
            End If
        End If
    Next
End Sub

Private Sub AddParaToParasList(ByVal strParaName As String)
    With vsfAppfunPara
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, mAppFuncParaCol.���) = .Rows - 1
        .TextMatrix(.Rows - 1, mAppFuncParaCol.��������) = strParaName
    End With
End Sub

Private Sub cmdTestFunc_Click()
On Error GoTo ErrorHand
    Dim strVBS As String

    strVBS = txtVBS.Text
    
    If InStr(strVBS, "[[") = 0 Then
        '������Ԥ���������ֱ�ӽ�����֤
        If CheckAppCfg(strVBS, True) Then
            mintTestSta = ͨ��
        Else
            mintTestSta = δͨ��
            Exit Sub
        End If
    Else
        '����Ԥ�����������Ҫ�����������֤��
        If CheckAppCfg(strVBS, False) Then
            mintTestSta = frmPacsInterfaceVBSTest.zlShowMe(strVBS, Me)
        Else
            mintTestSta = δͨ��
            Exit Sub
        End If
    End If
    
    If mintTestSta = δͨ�� Then MsgBox "��֤ʧ�ܣ����顣", vbExclamation, gstrSysName
    If mintTestSta = ͨ�� Then
        If (MsgBox("������֤������" & vbLf & "�����ʵ������ж��Ƿ�������������ѡ���ǡ���", vbYesNo, "���Խ��")) = vbYes Then
            vsfAppFuns.TextMatrix(vsfAppFuns.RowSel, mAppFuncCol.��֤ͨ��) = 1
            vsfAppFuns.Cell(flexcpBackColor, vsfAppFuns.RowSel, 0) = vsfAppFuns.BackColorFixed
        Else
            mintTestSta = δͨ��
            vsfAppFuns.Cell(flexcpBackColor, vsfAppFuns.RowSel, 0) = &HC0C0FF
        End If
    End If
    
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub

Private Sub Form_Load()
On Error GoTo ErrorHand
    
    Call InitCommandBars
    Call InitFaceScheme
    
    Call InitAppfuncParaList
    Call InitAppFunsList
    Call InitAppList
    Call InitAppFuncs
    
    Call InitEdit
    Call InputFaceEnabled(False)
    Call RestoreWinState(Me)
    
    stbThis.Panels(4).Text = "������:" & UserInfo.����
    
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
    err.Clear
End Sub

Private Sub LoadAppInfo()
    Dim i As Integer
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    strSql = "Select ID,����,�汾,·��,����,ִ������,�Ƿ�����,����ģ�� " & _
             "From Ӱ�����ҽ� Where ����ģ�� = [1] Order By ID"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "", cboStation.ItemData(cboStation.ListIndex))
    
    If rsData.RecordCount <= 0 Then Exit Sub
    
    With vsfApp
        For i = 1 To rsData.RecordCount
            .TextMatrix(i, mAppCol.��������) = zlCommFun.NVL(rsData!����)
            .TextMatrix(i, mAppCol.����汾) = zlCommFun.NVL(rsData!�汾)
            .TextMatrix(i, mAppCol.����·��) = zlCommFun.NVL(rsData!·��)
            .TextMatrix(i, mAppCol.����ID) = zlCommFun.NVL(rsData!ID)
            .TextMatrix(i, mAppCol.����) = zlCommFun.NVL(rsData!����)
            .TextMatrix(i, mAppCol.ִ������) = zlCommFun.NVL(rsData!ִ������)
            .TextMatrix(i, mAppCol.�Ƿ�����) = zlCommFun.NVL(rsData!�Ƿ�����)
            .TextMatrix(i, mAppCol.����ģ��) = zlCommFun.NVL(rsData!����ģ��)
            
            rsData.MoveNext
        Next
        
        If .Rows > 1 Then
            If .Row = 1 Then
                .RowSel = 1
                Call vsfApp_SelChange
            Else
                .Row = 1
                .RowSel = 1
            End If
        End If
    End With
End Sub

Private Sub InitEdit()
    cboStation.Clear
    cboStation.AddItem "ȫ��"
    cboStation.ItemData(cboStation.NewIndex) = 0
    
    cboStation.AddItem "Ӱ��ҽ������վ"
    cboStation.ItemData(cboStation.NewIndex) = 1290
    
    cboStation.AddItem "Ӱ��ɼ�����վ"
    cboStation.ItemData(cboStation.NewIndex) = 1291
    
    cboStation.AddItem "Ӱ������վ"
    cboStation.ItemData(cboStation.NewIndex) = 1294
    
    cboStation.ListIndex = 0
End Sub

Private Sub InitAppList()
    Dim i As Integer
    
On Error GoTo ErrorHand

    With vsfApp
        .Cols = 9
        .Rows = 51
        
        .ColWidth(mAppCol.���) = 300
        .ColWidth(mAppCol.��������) = 1300
        .ColWidth(mAppCol.����汾) = 900
        
        .TextMatrix(0, mAppCol.���) = "��"
        .TextMatrix(0, mAppCol.��������) = "�������"
        .TextMatrix(0, mAppCol.����汾) = "����汾"
        .TextMatrix(0, mAppCol.����·��) = "����·��"
        
        .TextMatrix(0, mAppCol.����ID) = "����ID"
        .TextMatrix(0, mAppCol.����) = "����"
        .TextMatrix(0, mAppCol.ִ������) = "ִ������"
        .TextMatrix(0, mAppCol.�Ƿ�����) = "�Ƿ�����"
        .TextMatrix(0, mAppCol.����ģ��) = "����ģ��"
        
        .ExtendLastCol = True
        
        For i = 4 To .Cols - 1
            .ColHidden(i) = True
        Next
        
        For i = 1 To .Rows - 1
            .TextMatrix(i, mAppCol.���) = i
            .TextMatrix(i, mAppCol.����ID) = 0
        Next
        
        .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = flexAlignLeftCenter
        .RowSel = 0
    End With
    
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
    err.Clear
End Sub

Private Sub InitAppFunsList()
On Error GoTo ErrorHand
    Dim i As Integer
    

    
    With vsfAppFuns
        .Cols = 11
        .Rows = 1
        .WordWrap = True '�Զ�������ʾ
        .RowHeight(0) = 500 '�������Ϊһ����Ԫ����ʾ2��
        .ColWidth(mAppFuncCol.��������) = 1000
        .ColWidth(mAppFuncCol.���) = 300
        .ColWidth(mAppFuncCol.��������) = 1000
        .ColWidth(mAppFuncCol.���ù���) = 500
        .ColWidth(mAppFuncCol.�����Ҽ��˵�) = 500
        .ColWidth(mAppFuncCol.���빤����) = 500
        .ColWidth(mAppFuncCol.�Զ�ִ��ʱ��) = 1400
        .ColWidth(mAppFuncCol.��Ӧ����) = 1000
        
        
        .TextMatrix(0, mAppFuncCol.���) = "��"
        .TextMatrix(0, mAppFuncCol.��������) = "��������"
        .TextMatrix(0, mAppFuncCol.���ù���) = "���ù���"
        .TextMatrix(0, mAppFuncCol.�����Ҽ��˵�) = "�Ҽ��˵�"
        .TextMatrix(0, mAppFuncCol.���빤����) = "������"
        .TextMatrix(0, mAppFuncCol.�Զ�ִ��ʱ��) = "�Զ�ִ��ʱ��"
        .TextMatrix(0, mAppFuncCol.��Ӧ����) = "��Ӧ����"
        .TextMatrix(0, mAppFuncCol.��������) = "��������"
        .TextMatrix(0, mAppFuncCol.VBS�ű�) = "VBS�ű�"
        .TextMatrix(0, mAppFuncCol.����ID) = "����ID"
        .TextMatrix(0, mAppFuncCol.��֤ͨ��) = "��֤ͨ��"
          
        .ColHidden(mAppFuncCol.��������) = True
        .ColHidden(mAppFuncCol.VBS�ű�) = True
        .ColHidden(mAppFuncCol.����ID) = True
        .ColHidden(mAppFuncCol.��֤ͨ��) = True
        
        .ExtendLastCol = True
        .ColDataType(2) = flexDTBoolean
        .ColDataType(3) = flexDTBoolean
        .ColDataType(4) = flexDTBoolean
        .ColComboList(5) = C_STR_INTERFACE_0 & "|" & C_STR_INTERFACE_1 & "|" & C_STR_INTERFACE_2 & "|" & C_STR_INTERFACE_3 & "|" & _
                                      C_STR_INTERFACE_4 & "|" & C_STR_INTERFACE_5 & "|" & C_STR_INTERFACE_6 & "|" & C_STR_INTERFACE_7 & "|" & _
                                      C_STR_INTERFACE_11 & "|" & C_STR_INTERFACE_12 & "|" & C_STR_INTERFACE_13 & "|" & C_STR_INTERFACE_14 & "|" & _
                                      C_STR_INTERFACE_15 & "|" & C_STR_INTERFACE_16 & "|" & C_STR_INTERFACE_17 & "|" & C_STR_INTERFACE_21 & "|" & _
                                      C_STR_INTERFACE_22
        
        For i = 1 To .Rows - 1
            .TextMatrix(i, mAppFuncCol.���) = i
            .TextMatrix(i, mAppFuncCol.����ID) = 0
        Next
        
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignLeftCenter
    End With
    
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
    err.Clear
End Sub

Private Sub InitAppFuncs()
    Dim i As Integer
    With vsfFuncs
        .Clear
        
        .ExtendLastCol = True
        .Cols = 3
        .Rows = 10
        .FixedCols = 1
        .FixedRows = 0
        
        .ColWidth(mAppFuncsCol.�������) = 300
        .ColWidthMax = 300 '����������ʹ ������� ���Ϊ300
        .ColHidden(mAppFuncsCol.��������) = True
        
        For i = 0 To .Rows - 1
            .TextMatrix(i, mAppFuncsCol.�������) = i + 1
        Next
        
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignLeftCenter
        
    End With
End Sub

Private Sub InitAppfuncParaList()
    Dim i As Integer

On Error GoTo ErrorHand
    
    With vsfAppfunPara
        .Cols = 4
        .Rows = 1 '�������
        
        .ColWidth(mAppFuncParaCol.���) = 300
        .ColWidth(mAppFuncParaCol.��������) = 1100
        .ColWidth(mAppFuncParaCol.��������) = 1100
        
        .TextMatrix(0, mAppFuncParaCol.���) = "��"
        .TextMatrix(0, mAppFuncParaCol.��������) = "��������"
        .TextMatrix(0, mAppFuncParaCol.��������) = "��������"
        .TextMatrix(0, mAppFuncParaCol.��������) = "��������"
        
        .ColComboList(mAppFuncParaCol.��������) = "Ԥ����|�ַ���|������|������"
        
        .ExtendLastCol = True
        
        For i = 1 To .Rows - 1
            .TextMatrix(i, mAppFuncParaCol.���) = i
        Next
        
    End With
    
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
    err.Clear
End Sub

Private Sub InputFaceEnabled(ByVal blnEnabled As Boolean)
    Call AppInfoFaceEnabled(blnEnabled)
    Call FuncsFaceEnabled(blnEnabled)
    Call ParasFaceEnabled(blnEnabled)
    Call VBSFaceEnabled(blnEnabled)
End Sub

Private Sub AppFaceEnabled(ByVal blnEnabled As Boolean)
    lblStation.Enabled = blnEnabled
    cboStation.Enabled = blnEnabled
    vsfApp.Enabled = blnEnabled
End Sub

Private Sub AppInfoFaceEnabled(ByVal blnEnabled As Boolean)
    lblAppDir.Enabled = blnEnabled
    txtAppDir.Enabled = blnEnabled
    cmdSelectApp.Enabled = blnEnabled
    
    lblType.Enabled = blnEnabled
    cboType.Enabled = blnEnabled
    
    lblClasses.Enabled = blnEnabled
    cboClasses.Enabled = blnEnabled
    
    lblAppName.Enabled = blnEnabled
    txtAppName.Enabled = blnEnabled
    
    chkUseThisApp.Enabled = blnEnabled
End Sub

Private Sub FuncsFaceEnabled(ByVal blnEnabled As Boolean)
    Dim blHaveFunc As Boolean 'vsfAppFuns�Ƿ�����Ч����
    Dim blHaveFun As Boolean 'vsfFuncs�Ƿ�����Ч����
    Dim i As Long
    
    blHaveFunc = False
    blHaveFun = False
    
    blHaveFunc = vsfAppFuns.Rows > 1
    
    For i = 1 To vsfFuncs.Rows - 1
        If vsfFuncs.TextMatrix(1, mAppFuncsCol.�������) <> "" Then
            blHaveFun = True
            Exit For
        End If
    Next

    
    vsfAppFuns.Editable = IIf(blnEnabled, flexEDKbdMouse, flexEDNone)
    cmdAddFunc.Enabled = blnEnabled
    
    cmdDelFunc.Enabled = blnEnabled And blHaveFunc
    cmdDelFun.Enabled = blnEnabled And blHaveFunc And blHaveFun
    
    cmdTestFunc.Enabled = blnEnabled And blHaveFunc
End Sub

Private Sub ParasFaceEnabled(ByVal blnEnabled As Boolean)
    vsfAppfunPara.Editable = IIf(blnEnabled, flexEDKbdMouse, flexEDNone)
    cmdAddPara.Enabled = blnEnabled
    cmdDelPara.Enabled = blnEnabled
    
    If vsfAppFuns.Rows <= 1 Then
        vsfFuncs.Editable = flexEDNone
    Else
        vsfFuncs.Editable = IIf(blnEnabled, flexEDKbdMouse, flexEDNone)
    End If
End Sub

Private Sub VBSFaceEnabled(ByVal blnEnabled As Boolean)
    fraVBS.Enabled = blnEnabled
    chkModify.Enabled = blnEnabled
    txtVBS.Enabled = blnEnabled And chkModify.value = 1
End Sub

Private Sub InitCommandBars()
    '���ܴ���������
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrPopControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    
    '���ò˵����͹��������
    With cbrMain.Options
        .ShowExpandButtonAlways = False                         '�����ڹ������Ҳ���ʾѡ�ť,��ʹ�������㹻��
        .ToolBarAccelTips = True                                '��ʾ��ť��ʾ
        .AlwaysShowFullMenus = False                            '�����õĲ˵���������
        .UseFadedIcons = False                                  'ͼ����ʾΪ��ɫЧ��
        .IconsWithShadow = True                                 '���ָ�������ͼ����ʾ��ӰЧ��
        .UseDisabledIcons = True                                '��������ť����ʱͼ����ʾΪ������ʽ
        .LargeIcons = True                                      '��������ʾΪ��ͼ��
        .SetIconSize True, 24, 24                               '���ô�ͼ��ĳߴ�
        .SetIconSize False, 16, 16                              '����Сͼ��ĳߴ�
    End With
    With cbrMain
        .VisualTheme = xtpThemeOffice2003                      '���ÿؼ���ʾ���
        .EnableCustomization False                             '�Ƿ������Զ�������
        Set .Icons = zlCommFun.GetPubIcons                     '���ù�����ͼ��ؼ�
    End With

    Me.cbrMain.EnableCustomization False
    Me.cbrMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    '�˵�����
'Begin------------------------�༭�˵�--------------------------------------Ĭ�Ͽɼ�
    cbrMain.ActiveMenuBar.Title = "�˵�"
    
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, TMenuType.mtFile, "�ļ�(&F)")
    
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtSave, "����(&S)"): cbrControl.IconId = 3091
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtCancel, "ȡ��(&C)"): cbrControl.IconId = 3565
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtQuit, "�˳�(&Q)"): cbrControl.IconId = 2613: cbrControl.BeginGroup = True
    
    
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, TMenuType.mtEdit, "�༭(&E)")
    
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtAdd, "����(&N)"): cbrControl.IconId = 4010
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtMod, "�޸�(&M)"): cbrControl.IconId = 3003
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtDel, "ɾ��(&D)"): cbrControl.IconId = 4008
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtUse, "����(&A)"): cbrControl.IconId = 3006
    Set cbrControl = cbrMenuBar.CommandBar.Controls.Add(xtpControlButton, TMenuType.mtRefresh, "ˢ��(&R)"): cbrControl.IconId = 3823: cbrControl.BeginGroup = True
    cbrControl.ShortcutText = "F5"
    
    'Begin----------------------�鿴�˵�--------------------------------------
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, 1, "�鿴(V)")
    Call CreateViewAndHelpMenu(cbrMenuBar, Nothing)

    'Begin----------------------�����˵�--------------------------------------Ĭ�Ͽɼ�
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, 2, "����(H)")
    Call CreateViewAndHelpMenu(Nothing, cbrMenuBar)
    
    
    '---------------------����������------------------------------------------
    Set cbrToolBar = Me.cbrMain.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = True
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtSave, "����", "����"): cbrControl.IconId = 3091
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtCancel, "ȡ��", "ȡ��"): cbrControl.IconId = 3565
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtAdd, "����", "����"): cbrControl.IconId = 4010: cbrControl.BeginGroup = True
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtMod, "�޸�", "�޸�"): cbrControl.IconId = 3003
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtDel, "ɾ��", "ɾ��"): cbrControl.IconId = 4008
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtUse, "����", "����"): cbrControl.IconId = 3006
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtRefresh, "ˢ��", "ˢ��"): cbrControl.IconId = 791: cbrControl.BeginGroup = True

    Set cbrControl = cbrToolBar.Controls.Add(xtpControlButton, TMenuType.mtQuit, "�˳�", "�˳�"): cbrControl.IconId = 2613: cbrControl.BeginGroup = True
    cbrControl.BeginGroup = True
    
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
End Sub

Private Sub InitFaceScheme()
    Dim Pane1 As Pane, Pane2 As Pane
    
     With Me.dkpMain
        .VisualTheme = ThemeOffice2003
        .SetCommandBars cbrMain
        .Options.HideClient = True
        .Options.UseSplitterTracker = False 'ʵʱ�϶�
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
        
        .PanelPaintManager.BoldSelected = True
        .TabPaintManager.Position = xtpTabPositionLeft  'TAB�ŵ������ʾ
        .TabPaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .TabPaintManager.BoldSelected = True
        dkpMain.Options.DefaultPaneOptions = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
        
        Set Pane1 = .CreatePane(1, 300, 100, DockLeftOf)
        Pane1.Handle = picApp.hWnd
        
        Set Pane2 = .CreatePane(2, 500, 100, DockRightOf, Pane1)
        Pane2.Handle = picAppCfg.hWnd
    End With
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.Width < 15000 Then Me.Width = 15000
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrorHand

    mblnConfiging = False
    Call SaveWinState(Me)
    
    Exit Sub
ErrorHand:
End Sub

Private Sub picApp_Resize()
    On Error Resume Next
    
    lblStation.Left = 120
    lblStation.Top = 240
    
    cboStation.Left = lblStation.Left + lblStation.Width + 240
    cboStation.Top = 200
    cboStation.Width = picApp.Width - cboStation.Left - 120
    
    vsfApp.Left = lblStation.Left
    vsfApp.Top = cboStation.Top + cboStation.Height + 120
    vsfApp.Width = picApp.Width - vsfApp.Left * 2
    vsfApp.Height = picApp.Height - vsfApp.Top - 360
End Sub

Private Sub picAppCfg_Resize()
    On Error Resume Next
    
    '������Ϣ
    fraAppInfo.Left = 120
    fraAppInfo.Top = 120
    fraAppInfo.Width = picAppCfg.Width - fraAppInfo.Left * 2
    
    lblAppDir.Left = txtAppName.Left + txtAppName.Width + 300
    txtAppDir.Width = fraAppInfo.Width - lblAppDir.Left - lblAppDir.Width - cmdSelectApp.Width - 120
    cmdSelectApp.Left = txtAppDir.Left + txtAppDir.Width
    
    lblClasses.Left = cboType.Left + cboType.Width + 300
    cboClasses.Width = fraAppInfo.Width - cboClasses.Left - chkUseThisApp.Width - 300
    chkUseThisApp.Left = fraAppInfo.Width - chkUseThisApp.Width - 60
    
    '��������
    fraAppFuns.Left = 120
    fraAppFuns.Top = fraAppInfo.Top + fraAppInfo.Height + 120
    fraAppFuns.Width = fraAppInfo.Width
    fraAppFuns.Height = picAppCfg.Height - fraAppFuns.Top - 120
    
    Call fraFuncs.Move(0, 0, fraAppFuns.Width)
    Call vsfAppFuns.Move(120, 300, fraAppFuns.Width - 240)
    
    vsfFuncs.Left = vsfAppFuns.Cell(flexcpLeft, 0, mAppFuncCol.��Ӧ����) - 10
    vsfFuncs.Top = vsfAppFuns.Cell(flexcpTop, 0, mAppFuncCol.��Ӧ����) + vsfAppFuns.Cell(flexcpHeight, 0, mAppFuncCol.��Ӧ����) - 10
    vsfFuncs.Width = vsfAppFuns.Cell(flexcpWidth, 0, mAppFuncCol.��Ӧ����)
    vsfFuncs.Height = vsfAppFuns.Height - vsfFuncs.Top
    
    If fraFuncs.Width - cmdDelFun.Width - 120 > (cmdTestFunc.Left + cmdTestFunc.Width + 300) Then
        cmdDelFun.Left = (fraFuncs.Width - cmdDelFun.Width) - 120
    Else
        cmdDelFun.Left = cmdTestFunc.Left + cmdTestFunc.Width + 300
    End If
'
    cmdDelFun.Top = cmdAddFunc.Top
    
    Call fraFuncParas.Move(0, fraFuncs.Height + 240, fraAppFuns.Width * 0.5, fraAppFuns.Height - fraFuncs.Height - 240)
        
    Call vsfAppfunPara.Move(120, 300, fraFuncParas.Width - 240, fraFuncParas.Height - 600)
    If cboType.ItemData(cboType.ListIndex) = mExecuteType.��̬���� Then
        Call vsfAppfunPara.Move(120, 300, fraFuncParas.Width - 240, fraFuncParas.Height - 600)
        cmdDelPara.Visible = False
        cmdAddPara.Visible = False
    Else
        Call cmdAddPara.Move(120, fraFuncParas.Height - cmdAddPara.Height - 360)
        Call cmdDelPara.Move(1560, fraFuncParas.Height - cmdDelPara.Height - 360)
        Call vsfAppfunPara.Move(120, 300, fraFuncParas.Width - 240, fraFuncParas.Height - 1200)
        cmdDelPara.Visible = True
        cmdAddPara.Visible = True
    End If
    
    Call fraVBS.Move(fraFuncParas.Left + fraFuncParas.Width, fraFuncParas.Top, fraAppFuns.Width * 0.5, fraAppFuns.Height - fraFuncs.Height - 120)
        
    Call txtVBS.Move(60, 300, fraFuncParas.Width - 120, fraFuncParas.Height - 540)

    
    Call ShowButton
End Sub

Private Sub ClearAllCfg()
    Call ClearEdit
    Call ClearAppCfg
    Call ClearAppFuncCfg
    Call ClearFuncs
    Call ClearAppFuncParaCfg
End Sub

Private Sub ClearInputCfg()
    Call ClearAppFuncCfg
    Call ClearFuncs
    Call ClearAppFuncParaCfg
End Sub

Private Sub ClearAppCfg()
    Dim i As Integer, j As Integer
    
    For i = 1 To vsfApp.Rows - 1
        For j = 1 To vsfApp.Cols - 1
            vsfApp.TextMatrix(i, j) = ""
        Next
    Next
End Sub

Private Sub ClearFuncs()
    Dim i As Integer, j As Integer
    
    For i = 0 To vsfFuncs.Rows - 1
        For j = 1 To vsfFuncs.Cols - 1
            vsfFuncs.TextMatrix(i, j) = ""
        Next
    Next
End Sub

Private Sub ClearEdit()
    txtAppDir.Text = ""
    chkUseThisApp.value = 1
    cboType.Clear
    cboClasses.Clear
End Sub


Private Sub ClearAppFuncCfg()
    vsfAppFuns.Rows = 1
    vsfAppFuns.RowSel = 0
    If mblnIsAddCfg Then
        Call ClearEdit
    End If
    
    txtVBS.Text = ""
End Sub

Private Sub ClearAppFuncParaCfg()
    vsfAppfunPara.Rows = 1
    vsfAppfunPara.RowSel = 0
End Sub

Private Sub txtVBS_KeyUp(KeyCode As Integer, Shift As Integer)
 On Error GoTo ErrorHand
    vsfAppFuns.TextMatrix(vsfAppFuns.RowSel, mAppFuncCol.VBS�ű�) = txtVBS.Text
    If chkModify.value = 1 Then
        vsfAppFuns.TextMatrix(vsfAppFuns.RowSel, mAppFuncCol.��֤ͨ��) = 0
        Call VerifyAllFuns
    End If
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub

Private Sub txtVBS_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 On Error GoTo ErrorHand
    vsfAppFuns.TextMatrix(vsfAppFuns.RowSel, mAppFuncCol.VBS�ű�) = txtVBS.Text
    If chkModify.value = 1 And Button = 2 Then
        vsfAppFuns.TextMatrix(vsfAppFuns.RowSel, mAppFuncCol.��֤ͨ��) = 0
        Call VerifyAllFuns
    End If
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub

Private Sub vsfApp_SelChange()
    Dim lngAppId As Long
    
On Error GoTo ErrorHand

    txtAppName.Text = ""
    Call ClearEdit
    Call ClearAppFuncCfg
    Call ClearFuncs
    Call ClearAppFuncParaCfg
    
    If vsfApp.RowSel <= 0 Then Exit Sub

    lngAppId = Val(vsfApp.TextMatrix(vsfApp.RowSel, mAppCol.����ID))
    If lngAppId <= 0 Then Exit Sub
    
    chkUseThisApp.value = IIf(vsfApp.TextMatrix(vsfApp.RowSel, mAppCol.�Ƿ�����) = 1, 1, 0)
    txtAppName.Text = vsfApp.TextMatrix(vsfApp.RowSel, mAppCol.��������)
    
    Call LoadAppConfig(vsfApp.TextMatrix(vsfApp.RowSel, mAppCol.����·��), vsfApp.TextMatrix(vsfApp.RowSel, mAppCol.����))
    Call LoadCfgData(lngAppId)
    
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
    err.Clear
End Sub

Private Sub LoadCfgData(ByVal lngAppId As Long)
    Dim i As Integer
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    strSql = "Select �������,����,����,��������,�Ƿ�����,�Ƿ�����Ҽ��˵�,�Ƿ���빤����,�Զ�ִ��ʱ��,VBS�ű� From Ӱ�������� Where ���ID = [1] Order By �������"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "", lngAppId)
    
    If rsTemp.RecordCount <= 0 Then Exit Sub
    
    vsfAppFuns.Rows = rsTemp.RecordCount + 1
    vsfAppFuns.RowSel = 0
    
    For i = 1 To rsTemp.RecordCount
        vsfAppFuns.TextMatrix(i, mAppFuncCol.���) = i
        vsfAppFuns.TextMatrix(i, mAppFuncCol.��������) = zlCommFun.NVL(rsTemp!����)
        vsfAppFuns.TextMatrix(i, mAppFuncCol.���ù���) = zlCommFun.NVL(rsTemp!�Ƿ�����)
        vsfAppFuns.TextMatrix(i, mAppFuncCol.�����Ҽ��˵�) = zlCommFun.NVL(rsTemp!�Ƿ�����Ҽ��˵�)
        vsfAppFuns.TextMatrix(i, mAppFuncCol.���빤����) = zlCommFun.NVL(rsTemp!�Ƿ���빤����)
        vsfAppFuns.TextMatrix(i, mAppFuncCol.�Զ�ִ��ʱ��) = convertInterfaceTime(zlCommFun.NVL(rsTemp!�Զ�ִ��ʱ��), False)
        vsfAppFuns.TextMatrix(i, mAppFuncCol.��Ӧ����) = zlCommFun.NVL(rsTemp!����)
        vsfAppFuns.TextMatrix(i, mAppFuncCol.��������) = zlCommFun.NVL(rsTemp!��������)
        vsfAppFuns.TextMatrix(i, mAppFuncCol.VBS�ű�) = zlCommFun.NVL(rsTemp!VBS�ű�)
        vsfAppFuns.TextMatrix(i, mAppFuncCol.����ID) = zlCommFun.NVL(rsTemp!�������, 0)
        vsfAppFuns.TextMatrix(i, mAppFuncCol.��֤ͨ��) = 1
        
        rsTemp.MoveNext
    Next
   
    vsfAppFuns.Cell(flexcpAlignment, 1, 0, vsfAppFuns.Rows - 1, vsfAppFuns.Cols - 1) = flexAlignLeftCenter
    If vsfAppFuns.Rows > 1 Then
        vsfAppFuns.RowSel = 1
        txtVBS.Text = vsfAppFuns.TextMatrix(1, mAppFuncCol.VBS�ű�)
    End If
End Sub

Private Sub CreateVBS()
'�������ù���VBS�ű�
On Error GoTo ErrorHand
    Dim i As Integer, j As Integer
    Dim strVBS As String
    Dim strFuncs As String, strParas As String
    
    Dim strParaVal As String
    Dim strDefine As String, strReg As String, strDefines As String
    
    Dim strParasType As String
    Dim strReturn As String
    
    Dim strFuncInfo As String, strParaInfo As String
    Dim strFuncName As String, strFuncPara As String
    Dim strParaName As String, strParaType As String, strParaValu As String
    
    If vsfAppFuns.Rows < 2 Then Exit Sub
    
    '��������,�����ɲ�������Ԥ����|�ַ���|������|������
    strFuncInfo = vsfAppFuns.TextMatrix(vsfAppFuns.RowSel, mAppFuncCol.��Ӧ����)
    strParaInfo = vsfAppFuns.TextMatrix(vsfAppFuns.RowSel, mAppFuncCol.��������)
    
    For i = 0 To UBound(Split(strFuncInfo, "��"))
        If strFuncInfo <> "" Then
            strParas = ""
            strParasType = ""
            strDefine = ""
                
            If strParaInfo <> "" Then
                strFuncPara = Split(strParaInfo, "��")(i)
                 
                For j = 0 To UBound(Split(strFuncPara, ""))
                    strParaName = Split(Split(strFuncPara, "")(j), "��")(0)
                    strParaType = Split(Split(strFuncPara, "")(j), "��")(1)
                    strParaValu = Split(Split(strFuncPara, "")(j), "��")(2)
                    
                    Select Case strParaType
                        Case "Ԥ����"
                            If cboType.ItemData(cboType.ListIndex) = mExecuteType.Shell���� Then
                                strParas = strParas & " " & strParaValu
                            Else
                                strParas = strParas & ", """ & strParaValu & """"
                            End If
                            
                            strParasType = strParasType & "s"
                            
                        Case "�ַ���"
                            If cboType.ItemData(cboType.ListIndex) = mExecuteType.Shell���� Then
                                strParas = strParas & " " & strParaValu
                            Else
                                strParas = strParas & ", """ & strParaValu & """"
                            End If
                            
                            strParasType = strParasType & "s"
                            
                        Case "������"
                            strParas = strParas & ", " & Val(strParaValu)
                            
                            strParasType = strParasType & "l"
                            
                        Case "������"
                            strParas = strParas & ", " & IIf(strParaValu = "", "False", strParaValu)
                            
                            strParasType = strParasType & "s"
                            
                    End Select
                Next
            End If
            
            strFuncName = Split(strFuncInfo, "��")(i)
            strReturn = "s"
            
            If strDefine <> "" Then strDefines = strDefines & strDefine
        
            Select Case cboType.Text
                Case "��̬����"
                    strFuncs = strFuncs & "    Call objExecute." & strFuncName & IIf(strParas = "", "", "(" & Mid(strParas, 2) & ")") & vbCrLf
    
                Case "Shell����"
                    strFuncs = strFuncs & "    Call objExecute.exec (""" & strFuncName & IIf(strParas = "", """)", " " & Mid(strParas, 2) & """)") & vbCrLf
                    
                Case "API����"
                    strReg = strReg & "    objExecute.Register """ & cboClasses.Text & """, """ & strFuncName & """, ""i=" & strParasType & """, ""R=" & strReturn & """" & vbCrLf
                    strFuncs = strFuncs & "    Call objExecute." & strFuncName & IIf(strParas = "", "", "(" & Mid(strParas, 2) & ")") & vbCrLf
                    
            End Select
        End If
    Next
    
    '���ɽű�
    Select Case cboType.Text
        Case "��̬����"
            strVBS = "Sub ExcuteSub()" & vbCrLf & strDefines & _
                     "    Dim objExecute" & vbCrLf & _
                     "                " & vbCrLf & _
                     "    Set objExecute = CreateObject(""" & cboClasses.Text & """)" & vbCrLf & strFuncs & _
                     "End Sub"
        
        Case "Shell����"
            strVBS = "Sub ExcuteSub()" & vbCrLf & strDefines & _
                     "    Dim objExecute" & vbCrLf & _
                     "                " & vbCrLf & _
                     "    Set objExecute = CreateObject(""wscript.shell"")" & vbCrLf & strFuncs & _
                     "End Sub"
        
        Case "API����"
            strVBS = "Sub ExcuteSub()" & vbCrLf & strDefines & _
                     "    Dim objExecute" & vbCrLf & _
                     "                " & vbCrLf & _
                     "    Set objExecute = CreateObject(""DynamicWrapper"")" & vbCrLf & strReg & _
                     "                " & vbCrLf & strFuncs & _
                     "End Sub"
    End Select
    
    txtVBS.Text = strVBS
    vsfAppFuns.TextMatrix(vsfAppFuns.RowSel, mAppFuncCol.VBS�ű�) = strVBS
    vsfAppFuns.TextMatrix(vsfAppFuns.RowSel, mAppFuncCol.��֤ͨ��) = 0
    
    Call VerifyAllFuns
    
    Exit Sub
ErrorHand:
    err.Raise -1, "CreateVBS", "[GetSelectRowAdviceID]" & vbCrLf & err.Description
End Sub

Private Sub vsfAppfunPara_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
On Error GoTo ErrorHand
    Call ShowButton
Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
    err.Clear
End Sub

Private Sub vsfAppfunPara_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    If chkModify.value = 1 Then
        If Col = mAppFuncParaCol.�������� Or Col = mAppFuncParaCol.�������� Or Col = mAppFuncParaCol.�������� Then Cancel = True
    End If
    If Col = mAppFuncParaCol.�������� Then mstr�������� = vsfAppfunPara.TextMatrix(vsfAppfunPara.RowSel, mAppFuncParaCol.��������)
End Sub


Private Sub vsfAppfunPara_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    On Error Resume Next
    
    If NewRow = 0 Or OldRow = NewRow Or vsfAppfunPara.Rows - 1 < OldRow Then Exit Sub
    
    Cancel = True
    
    If Trim(vsfAppfunPara.TextMatrix(OldRow, mAppFuncParaCol.��������)) = "" Then
        MsgBox "�������Ͳ���Ϊ�գ������룡", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    If Trim(vsfAppfunPara.TextMatrix(OldRow, mAppFuncParaCol.��������)) = "" And _
       Trim(vsfAppfunPara.TextMatrix(OldRow, mAppFuncParaCol.��������)) <> "�ַ���" Then
        MsgBox "�������첻��Ϊ�գ������룡", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    Cancel = False
End Sub

Private Sub vsfAppfunPara_SelChange()
On Error GoTo ErrorHand
    Call ShowButton
Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
    err.Clear
End Sub

Private Sub vsfAppFuns_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = mAppFuncCol.��Ӧ���� Then
        Call CreateVBS
    End If
End Sub

Private Sub vsfAppFuns_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
'104536 ����vsfAppFunsˮƽ��������ͬ������vsfAppFuns�ڲ�vsfFuncs��λ��
    On Error Resume Next
    vsfFuncs.Left = vsfAppFuns.Cell(flexcpLeft, 0, mAppFuncCol.��Ӧ����) - 10
    vsfFuncs.Width = vsfAppFuns.Cell(flexcpWidth, 0, mAppFuncCol.��Ӧ����)
End Sub

Private Sub vsfAppFuns_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If chkModify.value = 1 Then
        If Col = mAppFuncCol.��Ӧ���� Then Cancel = True
    End If
End Sub

Private Sub vsfAppFuns_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    On Error Resume Next
    
    If OldRow < 1 Or NewRow = 0 Or OldRow = NewRow Or vsfAppFuns.Rows - 1 < OldRow Then Exit Sub
    
    Cancel = True
    
    If Trim(vsfAppFuns.TextMatrix(OldRow, mAppFuncCol.��������)) = "" Then
        MsgBox "�������Ʋ���Ϊ�գ������룡", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    If Trim(vsfFuncs.TextMatrix(0, mAppFuncsCol.���ܷ���)) = "" And vsfAppFuns.Rows > 2 Then
        MsgBox "���ܶ�Ӧ��������Ϊ�գ������룡", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    Cancel = False
End Sub

Private Sub vsfAppFuns_SelChange()
On Error GoTo ErrorHand

    If vsfAppFuns.Row > 0 Then txtVBS.Text = vsfAppFuns.TextMatrix(vsfAppFuns.Row, mAppFuncCol.VBS�ű�)

    Call LoadFuncFunCfg
    Call DoFraFuncParasCaption
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
    err.Clear
End Sub

Private Sub LoadFuncFunCfg()
'���ض�Ӧ���ܵķ���
On Error GoTo ErrorHand
    Dim i As Integer
    Dim strFuncInfo As String
    Dim strParaInfo As String
    
    If vsfAppFuns.RowSel <= 0 Then Exit Sub
    
    Call ClearFuncs
    
    If vsfAppFuns.TextMatrix(vsfAppFuns.RowSel, mAppFuncCol.��Ӧ����) <> "" Then
        strFuncInfo = vsfAppFuns.TextMatrix(vsfAppFuns.RowSel, mAppFuncCol.��Ӧ����)
        strParaInfo = vsfAppFuns.TextMatrix(vsfAppFuns.RowSel, mAppFuncCol.��������)
        
        For i = 0 To UBound(Split(strFuncInfo, "��"))
            vsfFuncs.TextMatrix(i, mAppFuncsCol.���ܷ���) = Split(strFuncInfo, "��")(i)
        Next
        
        For i = 0 To UBound(Split(strParaInfo, "��"))
            vsfFuncs.TextMatrix(i, mAppFuncsCol.��������) = Split(strParaInfo, "��")(i)
        Next
    End If
    
    If vsfFuncs.RowSel <> 0 Then
        vsfFuncs.Row = 0
        vsfFuncs.RowSel = 0
    Else
        Call LoadFuncParaCfg
    End If
    
    Exit Sub
ErrorHand:
    err.Raise -1, "LoadFuncFunCfg", "[GetSelectRowAdviceID]" & vbCrLf & err.Description
End Sub

Private Sub LoadFuncParaCfg()
'���ط�����Ӧ�Ĳ���
On Error GoTo ErrorHand
    Dim strParaInfo As String
    Dim i As Integer
    
    If vsfFuncs.TextMatrix(vsfFuncs.RowSel, mAppFuncsCol.��������) <> "" Then
        strParaInfo = vsfFuncs.TextMatrix(vsfFuncs.RowSel, mAppFuncsCol.��������)
        
        If strParaInfo <> "" Then
            vsfAppfunPara.Rows = UBound(Split(strParaInfo, "")) + 2
            
            For i = 0 To UBound(Split(strParaInfo, ""))
                vsfAppfunPara.TextMatrix(i + 1, mAppFuncParaCol.���) = i + 1
                vsfAppfunPara.TextMatrix(i + 1, mAppFuncParaCol.��������) = Split(Split(strParaInfo, "")(i), "��")(0)
                vsfAppfunPara.TextMatrix(i + 1, mAppFuncParaCol.��������) = Split(Split(strParaInfo, "")(i), "��")(1)
                vsfAppfunPara.TextMatrix(i + 1, mAppFuncParaCol.��������) = Split(Split(strParaInfo, "")(i), "��")(2)
            Next
            
            vsfAppfunPara.Cell(flexcpAlignment, 1, 0, vsfAppfunPara.Rows - 1, vsfAppfunPara.Cols - 1) = flexAlignLeftCenter
        End If
    Else
        For i = 1 To vsfAppfunPara.Rows - 1
            vsfAppfunPara.TextMatrix(i, mAppFuncParaCol.���) = i
            vsfAppfunPara.TextMatrix(i, mAppFuncParaCol.��������) = ""
            vsfAppfunPara.TextMatrix(i, mAppFuncParaCol.��������) = ""
            vsfAppfunPara.TextMatrix(i, mAppFuncParaCol.��������) = ""
        Next
    
    End If
    

    Exit Sub
ErrorHand:
    err.Raise -1, "LoadFuncParaCfg", "[GetSelectRowAdviceID]" & vbCrLf & err.Description
End Sub

Private Sub vsfAppfunPara_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrorHand

    Dim str�������� As String
        
    str�������� = vsfAppfunPara.TextMatrix(vsfAppfunPara.RowSel, mAppFuncParaCol.��������)
        
    If Col = mAppFuncParaCol.�������� Then
        If str�������� = "" Then
            MsgBox "����ѡ�������ֵ����!", vbExclamation, gstrSysName
            vsfAppfunPara.TextMatrix(vsfAppfunPara.RowSel, mAppFuncParaCol.��������) = ""
            vsfAppfunPara.EditCell
        End If
    End If
    
    If Col = mAppFuncParaCol.�������� Then
        
        If mstr�������� <> str�������� Then
                
            vsfAppfunPara.TextMatrix(vsfAppfunPara.RowSel, mAppFuncParaCol.��������) = ""
            
            If str�������� = "������" Then
                vsfAppfunPara.TextMatrix(vsfAppfunPara.RowSel, mAppFuncParaCol.��������) = "0"
            ElseIf str�������� = "�ַ���" Then
                vsfAppfunPara.TextMatrix(vsfAppfunPara.RowSel, mAppFuncParaCol.��������) = ""
            ElseIf str�������� = "������" Then
                vsfAppfunPara.TextMatrix(vsfAppfunPara.RowSel, mAppFuncParaCol.��������) = "False"
            End If
                        
        End If
                
    End If
    
    vsfAppfunPara.Cell(flexcpAlignment, 0, 0, vsfAppfunPara.Rows - 1, vsfAppfunPara.Cols - 1) = flexAlignLeftCenter
    
    Call RefreshCfg
    
    If Col = mAppFuncParaCol.�������� Or Col = mAppFuncParaCol.�������� Then
        Call CreateVBS
    End If
    
    Call ShowButton

    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
    err.Clear
End Sub

Private Sub RefreshCfg()
    Dim i As Integer
    Dim strFuncInfo As String
    Dim strParaInfo As String
    
    For i = 1 To vsfAppfunPara.Rows - 1
        If vsfAppfunPara.TextMatrix(i, mAppFuncParaCol.��������) = "" And vsfAppfunPara.TextMatrix(i, mAppFuncParaCol.��������) <> "" Then
            MsgBox "�����ڲ��������б���ѡ���Ӧ������ֵ���ͣ�", vbExclamation, gstrSysName
            Exit Sub
        End If
        
        strParaInfo = strParaInfo & "" & vsfAppfunPara.TextMatrix(i, mAppFuncParaCol.��������) & "��" & vsfAppfunPara.TextMatrix(i, mAppFuncParaCol.��������) & "��" & vsfAppfunPara.TextMatrix(i, mAppFuncParaCol.��������)
    Next
    
    vsfFuncs.TextMatrix(vsfFuncs.RowSel, mAppFuncsCol.��������) = IIf(strParaInfo <> "", Mid(strParaInfo, 2), "")
    
    strFuncInfo = ""
    strParaInfo = ""
    
    For i = 0 To vsfFuncs.Rows - 1
        If vsfFuncs.TextMatrix(i, mAppFuncsCol.���ܷ���) <> "" Then
            strFuncInfo = strFuncInfo & "��" & vsfFuncs.TextMatrix(i, mAppFuncsCol.���ܷ���)
            strParaInfo = strParaInfo & "��" & vsfFuncs.TextMatrix(i, mAppFuncsCol.��������)
        End If
    Next
    
    If vsfAppFuns.Rows > 1 Then
        vsfAppFuns.TextMatrix(vsfAppFuns.RowSel, mAppFuncCol.��Ӧ����) = IIf(strFuncInfo <> "", Mid(strFuncInfo, 2), "")
        vsfAppFuns.TextMatrix(vsfAppFuns.RowSel, mAppFuncCol.��������) = IIf(strParaInfo <> "", Mid(strParaInfo, 2), "")
    End If

End Sub

Private Sub vsfAppfunPara_EnterCell()
On Error GoTo ErrorHand
    If vsfAppfunPara.ColSel = mAppFuncParaCol.�������� Then
        If vsfAppfunPara.TextMatrix(vsfAppfunPara.RowSel, mAppFuncParaCol.��������) = "Ԥ����" Then
            '?���Զ�̬��Ӳ������ͣ�
            vsfAppfunPara.ColComboList(mAppFuncParaCol.��������) = C_STR_CUSTOMPARAS
        ElseIf vsfAppfunPara.TextMatrix(vsfAppfunPara.RowSel, mAppFuncParaCol.��������) = "������" Then
            vsfAppfunPara.ColComboList(mAppFuncParaCol.��������) = "True|False"
        Else
            vsfAppfunPara.ColComboList(mAppFuncParaCol.��������) = ""
        End If
        
    End If
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
    err.Clear
End Sub

Private Sub ShowButton()
'��ָ����Ԫ����ʾ���ð�ť
    cmdConfigWindow.Visible = False
    
    With vsfAppfunPara
        If .RowSel < 1 Then Exit Sub
        
        cmdConfigWindow.Left = .Cell(flexcpLeft, .RowSel, mAppFuncParaCol.��������) + .Cell(flexcpWidth, .RowSel, mAppFuncParaCol.��������) - cmdConfigWindow.Width
        cmdConfigWindow.Top = .Cell(flexcpTop, .RowSel, mAppFuncParaCol.��������)
        cmdConfigWindow.Height = .Cell(flexcpHeight, .RowSel, mAppFuncParaCol.��������)
        
        If cmdConfigWindow.Top < .RowHeight(0) Then Exit Sub
    
        cmdConfigWindow.Visible = .TextMatrix(.RowSel, mAppFuncParaCol.��������) = "�ַ���"
    End With
End Sub

Private Sub vsfFuncs_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrorHand
    If Col = mAppFuncsCol.���ܷ��� Then
        If cboType.ItemData(cboType.ListIndex) = mExecuteType.��̬���� Then
            '���ط�����Ӧ�Ĳ���
            Call ClearAppFuncParaCfg  '�������
            Call LoadParasWithFunc(txtAppDir.Text, cboClasses.Text, vsfFuncs.TextMatrix(Row, Col))
        End If
    End If
    
    vsfFuncs.TextMatrix(vsfFuncs.RowSel, mAppFuncsCol.��������) = ""
    
    If Col = mAppFuncsCol.���ܷ��� Then
        Call CreateVBS
    End If
    
    Call RefreshCfg
    
    
    If Col = mAppFuncsCol.���ܷ��� Then
        Call DoFraFuncParasCaption
    End If
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
    err.Clear
End Sub

Private Sub vsfFuncs_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrorHand
    If chkModify.value = 1 Then
        If Col = mAppFuncsCol.�������� Or Col = mAppFuncsCol.���ܷ��� Then Cancel = True
    End If
    
    If vsfAppFuns.Rows <= 0 Then
        Cancel = True
        Exit Sub
    End If
    
    If Row = 0 Then Exit Sub
    
    If vsfFuncs.TextMatrix(Row - 1, mAppFuncsCol.���ܷ���) = "" Then Cancel = True
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
    err.Clear
End Sub

Private Sub vsfFuncs_ComboCloseUp(ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)
On Error Resume Next
    If Col = mAppFuncsCol.���ܷ��� Then
        SendKeys "{ENTER}"
    End If
    err.Clear
End Sub

Private Sub vsfFuncs_SelChange()
On Error GoTo ErrorHand
    Call LoadFuncParaCfg
    Call DoFraFuncParasCaption
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
    err.Clear
End Sub

Private Sub Menu_Help_Web_Mail_click()
On Error GoTo errHandle
    zlMailTo hWnd
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Help_About_click()
On Error GoTo errHandle
    ShowAbout Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Help_Help_click()
'���ܣ����ð�������
On Error GoTo errHandle
    ShowHelp App.ProductName, Me.hWnd, Me.Name
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Help_Web_Forum_click()
On Error GoTo errHandle
    Call zlWebForum(Me.hWnd)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Help_Web_Home_click()
On Error GoTo errHandle
    zlHomePage hWnd
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Menu_View_StatusBar_click(ByVal Control As XtremeCommandBars.ICommandBarControl)
On Error GoTo errHandle
    
    Control.Checked = Not Control.Checked
    
    Me.cbrMain.RecalcLayout
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Menu_View_ToolBar_Button_click(ByVal Control As XtremeCommandBars.ICommandBarControl)
On Error GoTo errHandle
    Dim i As Integer
    
    For i = 2 To cbrMain.Count
        Me.cbrMain(i).Visible = Not Me.cbrMain(i).Visible
    Next

    Control.Checked = Not Control.Checked
    Me.cbrMain.RecalcLayout
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Menu_View_ToolBar_Size_click(ByVal Control As XtremeCommandBars.ICommandBarControl)
On Error GoTo errHandle
    Me.cbrMain.Options.LargeIcons = Not Me.cbrMain.Options.LargeIcons
    Control.Checked = Not Control.Checked
    
    Me.cbrMain.RecalcLayout
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_View_ToolBar_Text_click(ByVal Control As XtremeCommandBars.ICommandBarControl)
On Error GoTo errHandle
    Dim i As Integer, cbrControl As CommandBarControl
    Dim intStyle As Integer

    For i = 2 To cbrMain.Count
        If Me.cbrMain(i).Controls.Count >= 1 Then
            intStyle = Me.cbrMain(i).Controls(i).Style
            If intStyle = xtpButtonIconAndCaption Then
                intStyle = xtpButtonIcon
                Me.cbrMain(i).ShowTextBelowIcons = False
            Else
                intStyle = xtpButtonIconAndCaption
                Me.cbrMain(i).ShowTextBelowIcons = True
            End If
        End If
        
        For Each cbrControl In Me.cbrMain(i).Controls
            cbrControl.Style = intStyle
        Next
    Next
    
    Control.Checked = Not Control.Checked
    Me.cbrMain.RecalcLayout
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Function convertInterfaceTime(ByVal strText As String, ByVal intConvetType As Boolean) As String
'��ִ��ʱ���ɺ���ת�������֣�Ĭ��Ϊ0(��ִ��)
'intConvetType˵�� : true ��strתint  false��intתstr
    If intConvetType Then
        convertInterfaceTime = 0
        
        Select Case strText
            Case C_STR_INTERFACE_0
                convertInterfaceTime = EInterfaceExeTime.���Զ�ִ��
            Case C_STR_INTERFACE_1
                convertInterfaceTime = EInterfaceExeTime.�ǼǺ�
            Case C_STR_INTERFACE_2
                convertInterfaceTime = EInterfaceExeTime.������
            Case C_STR_INTERFACE_3
                convertInterfaceTime = EInterfaceExeTime.��ͼ��
            Case C_STR_INTERFACE_4
                convertInterfaceTime = EInterfaceExeTime.���汣���
            Case C_STR_INTERFACE_5
                convertInterfaceTime = EInterfaceExeTime.����ǩ����
            Case C_STR_INTERFACE_6
                convertInterfaceTime = EInterfaceExeTime.������˺�
            Case C_STR_INTERFACE_7
                convertInterfaceTime = EInterfaceExeTime.�����ɺ�
            Case C_STR_INTERFACE_11
                convertInterfaceTime = EInterfaceExeTime.ȡ���Ǽ�ʱ
            Case C_STR_INTERFACE_12
                convertInterfaceTime = EInterfaceExeTime.ȡ������ʱ
            Case C_STR_INTERFACE_13
                convertInterfaceTime = EInterfaceExeTime.ɾ��ͼ��ʱ
            Case C_STR_INTERFACE_14
                convertInterfaceTime = EInterfaceExeTime.ȡ������ʱ
            Case C_STR_INTERFACE_15
                convertInterfaceTime = EInterfaceExeTime.ȡ��ǩ��ʱ
            Case C_STR_INTERFACE_16
                convertInterfaceTime = EInterfaceExeTime.ȡ�����ʱ
            Case C_STR_INTERFACE_17
                convertInterfaceTime = EInterfaceExeTime.ȡ�����ʱ
            Case C_STR_INTERFACE_21
                convertInterfaceTime = EInterfaceExeTime.����л���
            Case C_STR_INTERFACE_22
                convertInterfaceTime = EInterfaceExeTime.���沵�غ�
        End Select
        
    Else
        convertInterfaceTime = C_STR_INTERFACE_0
        
        Select Case strText
            Case EInterfaceExeTime.���Զ�ִ��
                convertInterfaceTime = C_STR_INTERFACE_0
            Case EInterfaceExeTime.�ǼǺ�
                convertInterfaceTime = C_STR_INTERFACE_1
            Case EInterfaceExeTime.������
                convertInterfaceTime = C_STR_INTERFACE_2
            Case EInterfaceExeTime.��ͼ��
                convertInterfaceTime = C_STR_INTERFACE_3
            Case EInterfaceExeTime.���汣���
                convertInterfaceTime = C_STR_INTERFACE_4
            Case EInterfaceExeTime.����ǩ����
                convertInterfaceTime = C_STR_INTERFACE_5
            Case EInterfaceExeTime.������˺�
                convertInterfaceTime = C_STR_INTERFACE_6
            Case EInterfaceExeTime.�����ɺ�
                convertInterfaceTime = C_STR_INTERFACE_7
            Case EInterfaceExeTime.ȡ���Ǽ�ʱ
                convertInterfaceTime = C_STR_INTERFACE_11
            Case EInterfaceExeTime.ȡ������ʱ
                convertInterfaceTime = C_STR_INTERFACE_12
            Case EInterfaceExeTime.ɾ��ͼ��ʱ
                convertInterfaceTime = C_STR_INTERFACE_13
            Case EInterfaceExeTime.ȡ������ʱ
                convertInterfaceTime = C_STR_INTERFACE_14
            Case EInterfaceExeTime.ȡ��ǩ��ʱ
                convertInterfaceTime = C_STR_INTERFACE_15
            Case EInterfaceExeTime.ȡ�����ʱ
                convertInterfaceTime = C_STR_INTERFACE_16
            Case EInterfaceExeTime.ȡ�����ʱ
                convertInterfaceTime = C_STR_INTERFACE_17
            Case EInterfaceExeTime.����л���
                convertInterfaceTime = C_STR_INTERFACE_21
            Case EInterfaceExeTime.���沵�غ�
                convertInterfaceTime = C_STR_INTERFACE_22
                
        End Select
    End If
    

End Function

Private Sub DoFraFuncParasCaption()
On Error GoTo errH
    Dim strCaption  As String
    strCaption = "�����б�"
    
    If Len(vsfAppFuns.TextMatrix(vsfAppFuns.RowSel, mAppFuncCol.��������)) > 0 And Len(vsfFuncs.TextMatrix(vsfFuncs.RowSel, mAppFuncsCol.���ܷ���)) > 0 Then
        strCaption = strCaption & "[" & vsfAppFuns.TextMatrix(vsfAppFuns.RowSel, mAppFuncCol.��������) & " - "
        strCaption = strCaption & vsfFuncs.TextMatrix(vsfFuncs.RowSel, mAppFuncsCol.���ܷ���) & "]"
    ElseIf Len(vsfAppFuns.TextMatrix(vsfAppFuns.RowSel, mAppFuncCol.��������)) > 0 And Len(vsfFuncs.TextMatrix(vsfFuncs.RowSel, mAppFuncsCol.���ܷ���)) = 0 Then
        strCaption = strCaption & "[" & vsfAppFuns.TextMatrix(vsfAppFuns.RowSel, mAppFuncCol.��������) & "]"
    End If
    
    fraFuncParas.Caption = strCaption
    Exit Sub
errH:
    fraFuncParas.Caption = "�����б�"
End Sub

Private Sub VerifyAllFuns()
'����Ƿ����й����Ѿ�ͨ����֤
'&HC0C0FF �ۺ�ɫ
On Error GoTo errH
    Dim i As Long

    With vsfAppFuns
        For i = 1 To vsfAppFuns.Rows - 1
            If vsfAppFuns.TextMatrix(i, mAppFuncCol.��������) = "" Then Exit Sub

            If vsfAppFuns.TextMatrix(i, mAppFuncCol.��֤ͨ��) <> "1" Then
                .Cell(flexcpBackColor, i, 0) = &HC0C0FF
            Else
                .Cell(flexcpBackColor, i, 0) = .BackColorFixed
            End If
        Next
    End With
    Exit Sub
errH:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub

Private Function DoBeforeSave() As Boolean
'�������ǰ�Ĵ���������δ��֤�Ĺ��ܣ�����ʾ��Ҫ����֤
On Error GoTo errH
    Dim i As Long

    DoBeforeSave = True
    With vsfAppFuns
        For i = 1 To vsfAppFuns.Rows - 1

            If .Cell(flexcpBackColor, i, 0) = &HC0C0FF Then
                If (MsgBox("����δ������֤����֤δͨ���Ĺ�����ʱ�������棬ѡ���ǡ�������֤�����������棻ѡ�񡮷��Ƚ��й�����֤��", vbYesNo, gstrSysName)) = vbYes Then
                    DoBeforeSave = True
                Else
                    DoBeforeSave = False
                End If
                Exit Function
            End If

        Next
    End With
    Exit Function
errH:
    DoBeforeSave = False
    MsgBox err.Description, vbExclamation, gstrSysName
End Function


