VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClinicChargeBalance 
   Caption         =   "�����շѽ���"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10365
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "����"
      Size            =   15.75
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmClinicChargeBalance.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   10365
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdExit 
      Caption         =   "����¼��(&X)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   8250
      TabIndex        =   37
      Top             =   2265
      Width           =   2055
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "�����շ�(&J)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   8235
      TabIndex        =   32
      Top             =   915
      Width           =   2055
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   3090
      Left            =   45
      ScaleHeight     =   3090
      ScaleWidth      =   7995
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   990
      Width           =   7995
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         ForeColor       =   &H80000008&
         Height          =   1350
         Left            =   45
         ScaleHeight     =   1320
         ScaleWidth      =   3060
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1650
         Width           =   3090
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
            Left            =   2040
            TabIndex        =   30
            Top             =   615
            Width           =   1005
         End
         Begin XtremeSuiteControls.ShortcutCaption stcTittleTotal 
            Height          =   420
            Left            =   15
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   30
            Width           =   3045
            _Version        =   589884
            _ExtentX        =   5371
            _ExtentY        =   741
            _StockProps     =   6
            Caption         =   "�Ը��ϼ�"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   15.76
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
         End
      End
      Begin VB.PictureBox picPay 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         ForeColor       =   &H80000008&
         Height          =   2910
         Left            =   3195
         ScaleHeight     =   2880
         ScaleWidth      =   4710
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   90
         Width           =   4740
         Begin VB.TextBox txt��Ԥ�� 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            IMEMode         =   3  'DISABLE
            Left            =   1380
            MaxLength       =   10
            TabIndex        =   3
            Top             =   165
            Width           =   3240
         End
         Begin VB.ComboBox cbo֧����ʽ 
            BackColor       =   &H8000000F&
            ForeColor       =   &H8000000D&
            Height          =   435
            Left            =   1380
            Style           =   2  'Dropdown List
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   720
            Width           =   1245
         End
         Begin VB.TextBox txt������� 
            Height          =   480
            IMEMode         =   3  'DISABLE
            Left            =   1380
            MaxLength       =   30
            TabIndex        =   10
            Top             =   1815
            Width           =   3225
         End
         Begin VB.TextBox txt�ɿ� 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            IMEMode         =   3  'DISABLE
            Left            =   2700
            MaxLength       =   12
            TabIndex        =   6
            Top             =   735
            Width           =   1920
         End
         Begin VB.TextBox txtժҪ 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   1380
            MaxLength       =   50
            MultiLine       =   -1  'True
            TabIndex        =   12
            Top             =   2385
            Width           =   3210
         End
         Begin VB.TextBox txt�Ҳ� 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            IMEMode         =   3  'DISABLE
            Left            =   1380
            Locked          =   -1  'True
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   1275
            Width           =   3225
         End
         Begin VB.Label lbl��Ԥ�� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " Ԥ���"
            Height          =   315
            Left            =   180
            TabIndex        =   2
            Top             =   225
            Width           =   1170
         End
         Begin VB.Label lblPayType 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ɡ���"
            Height          =   315
            Left            =   360
            TabIndex        =   4
            Top             =   765
            Width           =   990
         End
         Begin VB.Label lbl������� 
            AutoSize        =   -1  'True
            Caption         =   "�������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   90
            TabIndex        =   9
            Top             =   1950
            Width           =   1260
         End
         Begin VB.Label lbl�Ҳ� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ҡ���"
            Height          =   315
            Left            =   360
            TabIndex        =   7
            Top             =   1350
            Width           =   990
         End
         Begin VB.Label lblժҪ 
            AutoSize        =   -1  'True
            Caption         =   "ժ  Ҫ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   390
            TabIndex        =   11
            Top             =   2460
            Width           =   960
         End
      End
      Begin VB.PictureBox picTotal 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         ForeColor       =   &H80000008&
         Height          =   1395
         Left            =   45
         ScaleHeight     =   1365
         ScaleWidth      =   3060
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   90
         Width           =   3090
         Begin XtremeSuiteControls.ShortcutCaption stcTittile 
            Height          =   450
            Left            =   15
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   30
            Width           =   3045
            _Version        =   589884
            _ExtentX        =   5371
            _ExtentY        =   794
            _StockProps     =   6
            Caption         =   "��ǰδ��"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   15.76
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
            ForeColor       =   &H000000FF&
            Height          =   615
            Left            =   2055
            TabIndex        =   15
            Top             =   585
            Width           =   1005
         End
      End
   End
   Begin VB.TextBox txt�ϼ� 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   6435
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   285
      Width           =   1575
   End
   Begin VB.TextBox txtҽ�� 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   1500
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   315
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   60
      Left            =   0
      TabIndex        =   23
      Top             =   900
      Width           =   8100
   End
   Begin VB.Frame fraSplitLeft 
      Caption         =   "Frame3"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4275
      Left            =   8100
      TabIndex        =   20
      Top             =   -180
      Width           =   30
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   21
      Top             =   6120
      Width           =   10365
      _ExtentX        =   18283
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   3572
            MinWidth        =   882
            Picture         =   "frmClinicChargeBalance.frx":08CA
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8599
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   582
            MinWidth        =   2
            Object.Tag             =   "�����շ�Ԥ�������ʾ"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   582
            MinWidth        =   2
            Object.Tag             =   "�����շ�������������ʾ"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   882
            MinWidth        =   882
            Picture         =   "frmClinicChargeBalance.frx":115E
            Key             =   "Calc"
            Object.ToolTipText     =   "������:ALT+?"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1693
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
            Object.Width           =   1693
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picBlance 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   2025
      Left            =   90
      ScaleHeight     =   1995
      ScaleWidth      =   11325
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   4095
      Width           =   11355
      Begin VB.CommandButton cmdDel 
         Caption         =   "ɾ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   8745
         TabIndex        =   31
         Top             =   60
         Width           =   1080
      End
      Begin VSFlex8Ctl.VSFlexGrid vsBlance 
         Height          =   1815
         Left            =   15
         TabIndex        =   18
         Top             =   495
         Width           =   9915
         _cx             =   17489
         _cy             =   3201
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
         BackColorSel    =   16761024
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
         FormatString    =   $"frmClinicChargeBalance.frx":1838
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
      Begin VB.Label lbl�ѽ� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "�Ѹ��ϼ�:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   4305
         TabIndex        =   17
         Top             =   120
         Width           =   1365
      End
      Begin VB.Label Label4 
         Caption         =   "����֧�����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   0
         TabIndex        =   16
         Top             =   98
         Width           =   2145
      End
   End
   Begin VB.PictureBox pic��� 
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   8145
      ScaleHeight     =   1140
      ScaleWidth      =   2040
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   2865
      Width           =   2040
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0111"
         Height          =   285
         Left            =   135
         TabIndex        =   35
         Top             =   600
         Width           =   1890
      End
      Begin VB.Label lbl��� 
         Caption         =   "�������"
         Height          =   315
         Left            =   105
         TabIndex        =   34
         Top             =   90
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "����շ�(&O)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   8208
      TabIndex        =   19
      Top             =   228
      Width           =   2055
   End
   Begin VB.CommandButton cmdYBBalance 
      Caption         =   "ҽ������(&Y)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   8220
      TabIndex        =   36
      Top             =   255
      Width           =   2055
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "�շѺϼ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   5085
      TabIndex        =   14
      Top             =   360
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ҽ��֧��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   120
      TabIndex        =   13
      Top             =   375
      Width           =   1260
   End
End
Attribute VB_Name = "frmClinicChargeBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Enum gChargePayType
    EM_FUN_�շ� = 0
    EM_FUN_���� = 1
    EM_FUN_���� = 2
End Enum
Public Enum gExitMode
    EM_EX_��� = 0
    EM_EX_��ͣ = 1
    EM_EX_���� = 2
    EM_EX_���� = 3
    EM_EX_�˳� = 4
End Enum
Private mbytFunc As gChargePayType  '0-�շ�;1-����
Private mfrmMain As frmClinicCharge
Private mbytReturnMode As gExitMode
Private mbln�쳣���� As Boolean
Private mblnYB�˿� As Boolean 'ҽ������������˵��ݽ�����
Private mbln�ֵ��ݽ������ȫ�� As Boolean '��ǰ�쳣�����Ƿ�Ϊ�ֵ��ݽ������ȫ�������
Private mblnElsePersonErrBill As Boolean '�Ƿ������˵��쳣����
'------------------------------------------------------------------------------------------
'���������ر���
Private mobjChargeInfor As clsClinicChargeInfor
Private mlngModule As Long, mstrPrivs As String
Private mstrYBPati As String
Private mblnOK As Boolean
Private mbln�������� As Boolean
Private mblnCur���� As Boolean
Private mlngR As Long
Private mlngBrushCardTypeID As Long '����������ˢ���Ŀ����ID,�Ա�ȱʡ��λ�ڸ�֧�������
Private mblnUnloaded  As Boolean
Private mblnLoad As Boolean
Private mstr��֧Ʊ As String
Private mCurCardPay As gTY_PayMoney '���ο�֧��
Private mcolCardPayMode As Collection
Private Type TY_ChargeMoney
    dbl����ʵ�� As Double
    dbl����Ӧ�� As Double
    dbl����ҽ��֧�� As Double
    dbl�����Ѹ��ϼ� As Double
    dbl���γ�Ԥ��  As Double
    dbl��ǰδ�� As Double
    dblԤ����� As Double
    dbl������� As Double
    dbl����Ԥ�� As Double
    dblӦ���ۼ� As Double
    dbl�������� As Double
End Type

Private mCurCarge As TY_ChargeMoney
'------------------------------------------------------------------------------------------
'�ֲ�����
Private mblnFirst As Boolean
Private mblnUnLoad As Boolean '�Ƿ�Unload����
Private mbln�ѱ��� As Boolean
Private mblnҽ���ѱ��� As Boolean
Private mstrҽ������ As String
Private mblnYbBalanced As Boolean 'ҽ���Ѿ�����
Private mblnThreeInterface As Boolean '�Ѿ����������ӿ�
Private mcur������� As Currency
Private mblnSaveBill As Boolean '���ݱ���ɹ�
Private mblnCommitBill As Boolean '�����Ƿ��Ѿ��ύ��
Private mblnPriceBillCommit As Boolean '���۵��Ƿ��Ѿ��ύ
Private mcllPriceSQL As Collection 'ֱ���շ�ʱ���ȱ���Ϊ���۵�

Private mcllOverPro As Collection
Private mblnSavePrice As Boolean '����ҽ������Ϊ���۵�
Private mrsBalance As ADODB.Recordset   '������Ϣ
Private mstrTittle As String '�������
'----------------------------------------------------------------------------------------------
'ҽ�����
'��ǰ���������ҽ��֧�ֲ���
Private Type TYPE_MedicarePAR
    �����ѽɿ���� As Boolean    '27536
    ҽ���ӿڴ�ӡƱ�� As Boolean
    ���������շ� As Boolean
    �ֱҴ��� As Boolean
    �൥�ݷֵ��ݽ��� As Boolean '86321
    һ�ν���ֵ����˷� As Boolean '91602
End Type
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private mInsurePara As TYPE_MedicarePAR
Private mrsOneCard As ADODB.Recordset
Private mrsBlance As ADODB.Recordset
'---------------------------------------------------------------------------------
Private mbln�����շ� As Boolean
'---------------------------------------------------------------------------------
Private mdbl�ֽ� As Double, mdblԭδ�� As Double
Private Const VK_RETURN = &HD
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private mblnCacheKeyReturn As Boolean   '41025:�Ƿ񻺴��˻س���,���ܴ������շѽ���ˢ���б�������˻س�,�����Ҫ�ж�
Private mrsClassMoney As ADODB.Recordset
Private mcllSquareBalance As Collection '���ѿ�������Ϣ
Private mcllCurSquareBalance As Collection '��ǰ���ѿ�ˢ����Ϣ
Private mblnNotChange As Boolean
Private mblnCurBrushPrepay   As Boolean '��ǰ�Ƿ�ˢ��Ԥ����
    
Public Function zlChargeWin(ByVal frmMain As Object, ByVal bytFunc As gChargePayType, _
    ByVal lngModule As Long, ByVal strPrivs As String, _
    ByRef objChargeInfor As clsClinicChargeInfor, _
    Optional bytReturnMode As gExitMode = EM_EX_���, _
    Optional bln�������� As Boolean, _
    Optional lngBrushCardTypeID As Long = 0, _
    Optional bln�쳣���� As Boolean = False, _
    Optional blnElsePersonErrBill As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������:��ʾ����֧�����㴰��
    '���:frmMain-���õ�������
    '       bytFunc-0-�շ�;1-����
    '       lngModule -ģ���
    '       strPrivs-Ȩ�޴�
    '       objChargeInfor-������Ϣ
    '       lngBrushCardTypeID-ȱʡ��ˢ�����ID
    '       bln�쳣����-�쳣�������ϴ���(�쳣����ʱ����):���Ϊtrue,��ʾ������ϵ��쳣���ݽ�������
    '       blnElsePersonErrBill-�Ƿ������˵��쳣����
    '����:objChargeInfor.�ɿ���-����Ľɿ�����Ҳ����(���ֽ�ʱ,����)
    '     objChargeInfor.����Ӧ��-ҽ������,�������շ������,��Ҫ���¼��㱾�ε�Ӧ�ɶ�
    '     objChargeInfor.�շѽ���:���ر����շѵĽ��㷽ʽ,��ʽ����:
    '                       ���:�ɿ��־(1-�ɿ�;2-�Ҳ�)|���㷽ʽ1:���1:�ɿ��־(1-�ɿ�;2-�Ҳ�)|...
    '        bln��������-�Ƿ����¼���Ʊ��
    '        bytReturnMode-���ز���ģʽ(0-�����շ����,1-��ͣ�շ�;2-���������շ�;3-��������)
    '����:����շ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-12 09:59:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mobjChargeInfor = objChargeInfor: Set mcllOverPro = Nothing
    Set mrsClassMoney = Nothing: Set mrsBalance = Nothing
    mblnYbBalanced = False: mblnThreeInterface = False: mblnOK = False
    mblnUnLoad = False: mblnUnloaded = False: mblnSaveBill = False
    mblnCommitBill = bln�쳣����: mblnElsePersonErrBill = blnElsePersonErrBill
    mbln�ֵ��ݽ������ȫ�� = False
    mblnPriceBillCommit = False
    
    mstrPrivs = strPrivs: mlngModule = lngModule
    mlngBrushCardTypeID = lngBrushCardTypeID: Set mfrmMain = frmMain
    mbln�쳣���� = bln�쳣����
    mbytFunc = bytFunc: mbytReturnMode = EM_EX_���
    
    mCurCarge.dblӦ���ۼ� = mobjChargeInfor.Ӧ���ۼ�
    mbln�������� = mobjChargeInfor.Ӧ���ۼ� <> 0
    mblnOK = False
    On Error Resume Next
    Me.Show IIf(gfrmMain Is Nothing, 0, 1), frmMain
    bln�������� = mbln��������: bytReturnMode = mbytReturnMode
    'Set objChargeInfor = mobjChargeInfor
    zlChargeWin = mblnOK
End Function

Private Function SaveFeeBilL() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������õ�������
    '���:lng����ID-��������
    '     cllPro-ִ�е���ع���
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-06-12 14:36:17
    '˵��:
    '   ���ô˹���ʱ,����Ҫ��ʼ����,�쳣ʱ,���ݻ���,����ɹ�ʱ,δ�ύ����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPro As Collection, lng����ID As Long, strNos As String
    Dim cllItem As Collection, blnTransMedicare As Boolean
    Dim cllPriceSQL As Collection
    On Error GoTo errHandle
    
    If (mblnSaveBill And mblnCommitBill) Or mbytFunc = EM_FUN_���� Then
        gcnOracle.BeginTrans
        SaveFeeBilL = True: Exit Function
    End If
    
    If mfrmMain.zlGetSaveBillSQL(lng����ID, cllPriceSQL, strNos, cllPro, mcllOverPro) = False Then Exit Function
    mobjChargeInfor.����ID = lng����ID
    mobjChargeInfor.������� = -1 * lng����ID
    mobjChargeInfor.Nos = strNos
    Set mcllPriceSQL = cllPriceSQL
    
    '���ύ���۵����Ա㲻����ҩƷ��棩
    blnTransMedicare = True
    If mblnPriceBillCommit = False Then
        gcnOracle.BeginTrans
        For Each cllItem In cllPriceSQL
            ExecuteProcedureArrAy cllItem, Me.Caption, True, True
        Next
        gcnOracle.CommitTrans
        mblnPriceBillCommit = True
    End If
    
    gcnOracle.BeginTrans
    For Each cllItem In cllPro '91665
        ExecuteProcedureArrAy cllItem, Me.Caption, True, True
    Next
    
    mblnSaveBill = True: SaveFeeBilL = True
    Exit Function
errHandle:
    If blnTransMedicare Then gcnOracle.RollbackTrans
    If Err.Description Like "*��ǰ���㵥�۲�һ��*" Then
        If MsgBox("ĳЩ����ҩƷ�۸��ѷ����仯��Ҫ�Զ�����۸���", vbYesNo + vbQuestion + vbDefaultButton1, gstrSysName) = vbYes Then
            Call mfrmMain.zlReCalcMoney(mobjChargeInfor)
            Call SetControlProperty
            Call SetCtrlVisible
            Call SetControlEnabled
            Exit Function
        End If
        Call SaveErrLog
        Exit Function
    End If
    If ErrCenter() = 1 Then
'        Resume
    End If
End Function

Private Sub DelMedicareTempNOs()
    '����:ֱ���շ�ʱ,ɾ��ǰһ�������ύ�Ļ��۵�
    Dim i As Integer, varNos As Variant
    Dim strSQL As String
    
    On Error GoTo errHandle
    If mcllPriceSQL Is Nothing Then Exit Sub
    varNos = Split(mobjChargeInfor.Nos, ",")
    For i = 0 To UBound(varNos)
        If CollectionExitsValue(mcllPriceSQL, varNos(i)) Then
            strSQL = "zl_���ﻮ�ۼ�¼_DELETE('" & varNos(i) & "')"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        End If
    Next
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub zlInitTotalStru()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ�����ܽ��
    '����:���˺�
    '����:2011-12-26 13:19:04
    '����:44944
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mbln�������� And Not grsTotal Is Nothing Then Exit Sub
    Set grsTotal = New ADODB.Recordset
    grsTotal.Fields.Append "����", adBigInt, , adFldIsNullable
    grsTotal.Fields.Append "���㷽ʽ", adVarChar, 60, adFldIsNullable
    grsTotal.Fields.Append "������", adDouble, , adFldIsNullable
    grsTotal.CursorLocation = adUseClient
    grsTotal.LockType = adLockOptimistic
    grsTotal.CursorType = adOpenStatic
    grsTotal.Open
End Sub

Private Sub WhriteTotalDataToReCord(ByVal dblԤ�� As Double, _
    ByVal dblMoney As Double, ByVal dbl��֧Ʊ As Double)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݱ���,�Ա��ۼƻ�������
    '����:���˺�
    '����:2011-12-26 22:25:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, str���㷽ʽ As String, dbl�ɿ� As Double, dbl�Ҳ� As Double
    Dim int���� As Integer
    If grsTotal Is Nothing Then Call zlInitTotalStru
    If grsTotal.State <> 1 Then Call zlInitTotalStru
    
    If (mCurCardPay.int���� = 1 Or mCurCardPay.int���� = 2) And mblnCur���� = False Then
        dbl�ɿ� = Val(txt�ɿ�.Text)
        dbl�Ҳ� = Val(txt�Ҳ�.Text)
    End If
    If dbl�ɿ� = 0 Then
        dbl�ɿ� = 0: dbl�Ҳ� = 0
    End If
    On Error GoTo errHandle
    
    With vsBlance
        If grsTotal.RecordCount <> 0 Then grsTotal.MoveFirst
        If dbl�ɿ� <> 0 Then
            grsTotal.Find "���㷽ʽ='���νɿ�'", , adSearchForward, 1
            If grsTotal.EOF Then grsTotal.AddNew
            grsTotal!���� = -99
            grsTotal!���㷽ʽ = "�ɿ�"
            grsTotal!������ = dbl�ɿ�
        End If
        
        If dbl�Ҳ� <> 0 Then
            grsTotal.Find "���㷽ʽ='" & IIf(mCurCardPay.bln֧Ʊ, "��֧Ʊ", "�Ҳ�") & "'", , adSearchForward, 1
            If grsTotal.EOF Then grsTotal.AddNew
            grsTotal!���� = -98
            grsTotal!���㷽ʽ = IIf(mCurCardPay.bln֧Ʊ, "��֧Ʊ", "�Ҳ�")
            grsTotal!������ = dbl�Ҳ�
        End If
        
        For i = 1 To .Rows - 1
            str���㷽ʽ = Trim(.TextMatrix(i, .ColIndex("֧����ʽ")))
            int���� = Val(.RowData(i))
            If str���㷽ʽ <> "" Then
                '.rowdata:  0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
                '����:-99-�ɿ�;-98-�Ҳ�,0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
                grsTotal.Find "���㷽ʽ='" & str���㷽ʽ & "'", , adSearchForward, 1
                
                If grsTotal.EOF Then grsTotal.AddNew
                grsTotal!���� = int����
                grsTotal!���㷽ʽ = str���㷽ʽ
                grsTotal!������ = Val(Nvl(grsTotal!������)) + Val(.TextMatrix(i, .ColIndex("֧�����")))
                grsTotal.Update
            End If
        Next
        
        If dblԤ�� <> 0 Then
            grsTotal.Find "���㷽ʽ='Ԥ���'", , adSearchForward, 1
            If grsTotal.EOF Then grsTotal.AddNew
            grsTotal!���� = 1
            grsTotal!���㷽ʽ = "Ԥ���"
            grsTotal!������ = Val(Nvl(grsTotal!������)) + dblԤ��
            grsTotal.Update
        End If
        If mCurCardPay.bln���ѿ� And Not mcllCurSquareBalance Is Nothing Then
            For i = 1 To mcllCurSquareBalance.Count
                '��ǰˢ������(array(�����ID,���ѿ�ID,ˢ�����,����,����,�������,�Ƿ�����))
                grsTotal.Find "���㷽ʽ='" & mCurCardPay.str���㷽ʽ & "'", , adSearchForward, 1
                If grsTotal.EOF Then grsTotal.AddNew
                grsTotal!���� = IIf(mCurCardPay.blnOneCard, 4, 5)
                grsTotal!���㷽ʽ = mCurCardPay.str���㷽ʽ
                grsTotal!������ = Val(Nvl(grsTotal!������)) + Val(mcllCurSquareBalance(i)(2))
                grsTotal.Update
            Next
        Else
            grsTotal.Find "���㷽ʽ='" & mCurCardPay.str���㷽ʽ & "'", , adSearchForward, 1
            If grsTotal.EOF Then grsTotal.AddNew
            ''1-�ֽ���㷽ʽ,2-������ҽ������,3-ҽ�������ʻ�,4-ҽ������ͳ��,5-���տ���,6-�����ۿ�,7-һ��ͨ����,8-���㿨����;<0 ��ʾ������֧��
            '.rowdata:  0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
             '����:99-�ɿ�;98-�Ҳ�,0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
        
            Select Case mCurCardPay.int����
            Case 1, 2
                grsTotal!���� = 0
            Case 3, 4
                grsTotal!���� = 2
            Case 7, 8
                grsTotal!���� = IIf(mCurCardPay.blnOneCard, 4, 3)
            Case Else
                grsTotal!���� = 0
            End Select
            
            grsTotal!���㷽ʽ = mCurCardPay.str���㷽ʽ
            grsTotal!������ = Val(Nvl(grsTotal!������)) + dblMoney
            grsTotal.Update
            If dbl��֧Ʊ <> 0 Then
                grsTotal.Find "���㷽ʽ='" & mstr��֧Ʊ & "'", , adSearchForward, 1
                If grsTotal.EOF Then grsTotal.AddNew
                grsTotal!���� = 0
                grsTotal!���㷽ʽ = mstr��֧Ʊ
                grsTotal!������ = Val(Nvl(grsTotal!������)) + dbl��֧Ʊ
                grsTotal.Update
            End If
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub initInsure()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ҽ������
    '����:���˺�
    '����:2011-08-21 18:55:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjChargeInfor.intInsure = 0 Then Exit Sub
    mInsurePara.ҽ���ӿڴ�ӡƱ�� = gclsInsure.GetCapability(supportҽ���ӿڴ�ӡƱ��, mobjChargeInfor.����ID, mobjChargeInfor.intInsure)
    mInsurePara.���������շ� = gclsInsure.GetCapability(support���������շ�, mobjChargeInfor.����ID, mobjChargeInfor.intInsure)
    '���˺�:27536 20100119
    mInsurePara.�����ѽɿ���� = gclsInsure.GetCapability(support�����ѽɿ����, mobjChargeInfor.����ID, mobjChargeInfor.intInsure)
    mInsurePara.�ֱҴ��� = gclsInsure.GetCapability(support�ֱҴ���, mobjChargeInfor.����ID, mobjChargeInfor.intInsure)
    mInsurePara.�൥�ݷֵ��ݽ��� = gclsInsure.GetCapability(support�൥�ݷֵ��ݽ���, mobjChargeInfor.����ID, mobjChargeInfor.intInsure)
    mInsurePara.һ�ν���ֵ����˷� = gclsInsure.GetCapability(supportһ�ν���ֵ����˷�, mobjChargeInfor.����ID, mobjChargeInfor.intInsure)
End Sub
Private Sub InitBalanceData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����������
    '����:���˺�
    '����:2012-02-05 16:02:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call ClearBanalce
    With mCurCarge
          .dbl����ʵ�� = mobjChargeInfor.ʵ�ս��
          .dbl����ҽ��֧�� = mobjChargeInfor.ҽ��������
          .dbl�����Ѹ��ϼ� = 0
          .dbl����Ӧ�� = mobjChargeInfor.Ӧ�ս��
          .dbl��ǰδ�� = .dbl����ʵ�� - .dbl����ҽ��֧��
          .dbl���γ�Ԥ�� = 0
          .dbl�������� = 0
      End With
      '����Ԥ����δ������������������бȽϣ�ȷ���Ƿ��ظ�����
      mdblԭδ�� = mCurCarge.dbl��ǰδ��
      mblnYB�˿� = mCurCarge.dbl��ǰδ�� + mCurCarge.dblӦ���ۼ� < 0
      
End Sub
Private Sub ClearBanalce()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����������
    '����:���˺�
    '����:2012-02-05 16:02:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With mCurCarge
        .dbl����ʵ�� = 0
        .dbl����ҽ��֧�� = 0
        .dbl�����Ѹ��ϼ� = 0
        .dbl����Ӧ�� = 0
        .dbl��ǰδ�� = 0
        .dbl���γ�Ԥ�� = 0
        .dbl�������� = 0
    End With
    With vsBlance
        .Clear 1: .Rows = 2
    End With
    txtҽ��.Text = "0.00"
    txt�ϼ�.Text = "0.00"
End Sub

Private Sub LoadData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������
    '����:���˺�
    '����:2011-08-20 19:49:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim i As Long, bln���ѿ� As Boolean, lng�����ID As Long
    Dim strCardNo As String
    Dim blnYb As Boolean
    
    On Error GoTo errHandle
    
    Call ClearBanalce
    
    '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
    gstrSQL = "" & _
    "   Select  A.ID, " & _
    "        Case when Mod(A.��¼����,10)=1 then 1  " & _
    "             when B.���� is not null then  2 " & _
    "             when nvl(A.�����ID,0)<>0  then  3 " & _
    "             when J.���㷽ʽ is not null   then  4 " & _
    "             else 0 end as ����, " & _
    "        Mod(A.��¼����,10) as ��¼����,A.���㷽ʽ,A.��Ԥ��,A.ժҪ, " & _
    "        A.�����ID,A.���㿨���, " & _
    "        A.�������,A.����,A.������ˮ��,nvl(C.�Ƿ�����,0) as ���ƿ�, " & _
    "        nvl(C.�Ƿ�����,0) as �Ƿ�����,nvl(C.�Ƿ�ȫ��,0) as �Ƿ�ȫ��, " & _
    "        Decode(C.��������,NULL,0,1) as  �Ƿ�����,nvl(C.�Ƿ��˿��鿨,0) as �Ƿ��˿��鿨," & _
    "        C.���� as ���������,A.����˵��,A.�������,A.У�Ա�־, " & _
    "        decode(B.����,Null,0,1) as ҽ��,0 as ���ѿ�id" & _
    "   From ����Ԥ����¼ A ,ҽ�ƿ���� C,һ��ͨĿ¼ J, " & _
    "        (Select ���� From ���㷽ʽ where ���� in (3,4)) B" & _
    "   Where A.����ID= [1] " & _
    "         And A.���㷽ʽ=J.���㷽ʽ(+) And A.�����ID=C.ID(+) " & _
    "         And A.���㷽ʽ=B.����(+)  " & _
    "         And (a.��¼���� In (1, 11) Or Nvl(a.���㿨���, 0) = 0)"
       
    gstrSQL = gstrSQL & " Union ALL " & _
    "   Select A.ID,5 as  ����,Mod(A.��¼����,10) as ��¼����,A.���㷽ʽ,-1*nvl(b.Ӧ�ս��,0) as ��Ԥ��,A.ժҪ,A.�����ID,A.���㿨���," & _
    "        A.�������,B.����,B.������ˮ��,nvl( M.���ƿ�,0) as ���ƿ�, " & _
    "        nvl( M.�Ƿ�����,0) as �Ƿ�����,nvl(M.�Ƿ�ȫ��,0) as �Ƿ�ȫ��, " & _
    "        nvl(M.�Ƿ�����,0) as  �Ƿ�����,0 as �Ƿ��˿��鿨," & _
    "        M.���� as ���������,A.����˵��,A.�������,A.У�Ա�־,0 as ҽ��,B.���ѿ�id" & _
    "   From ����Ԥ����¼ A ,���˿������¼ B, ���ѿ����Ŀ¼ M " & _
    "   Where  a.Id = b.����id And a.���㿨��� = m.���  " & _
    "        And A.����ID = [1] and Mod(A.��¼����,10)<>1 "
    
   gstrSQL = "" & _
   "   Select  ����,��¼����,���㷽ʽ,ժҪ,�����ID,���������,���ƿ�,���㿨���,�������,����,������ˮ��, ����˵��,�������,У�Ա�־,ҽ��,���ѿ�id," & _
   "         max(�Ƿ�����) as �Ƿ�����,max(�Ƿ��˿��鿨) as �Ƿ��˿��鿨," & _
   "         max(�Ƿ�ȫ��) as �Ƿ�ȫ��,max(�Ƿ�����) as �Ƿ����� , nvl(sum(��Ԥ��),0) as ��Ԥ��" & _
   "   From (" & gstrSQL & ") " & _
   "   Group by ����,��¼����,���㷽ʽ,ժҪ,�����ID,���������,���ƿ�,���㿨���,�������,����,������ˮ��, ����˵��,�������,У�Ա�־,ҽ��,���ѿ�id"
    
    Set mrsBalance = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mobjChargeInfor.����ID)
    With mrsBalance
        i = 1: blnYb = False
        Do While Not .EOF
            Select Case Nvl(!����)
            Case 1 'Ԥ����
                mCurCarge.dbl���γ�Ԥ�� = RoundEx(mCurCarge.dbl���γ�Ԥ�� + Val(Nvl(!��Ԥ��)), 6)
                mCurCarge.dbl�����Ѹ��ϼ� = RoundEx(mCurCarge.dbl�����Ѹ��ϼ� + Val(Nvl(!��Ԥ��)), 6)
            Case 2, 3, 5 'ҽ��,һ��ͨ,���ѿ�
                If Nvl(!����) = 2 Then
                    mCurCarge.dbl����ҽ��֧�� = RoundEx(mCurCarge.dbl����ҽ��֧�� + Nvl(!��Ԥ��, 0), 6)
                    blnYb = True
                End If
                If Val(Nvl(mrsBalance!У�Ա�־, 0)) = 2 Then
                    With vsBlance
                        If .TextMatrix(i, .ColIndex("֧����ʽ")) <> "" Then
                            .Rows = .Rows + 1
                            i = i + 1
                        End If
                        .RowData(i) = Nvl(mrsBalance!����)
                        strCardNo = Nvl(mrsBalance!����)
                        lng�����ID = Val(Nvl(mrsBalance!���㿨���))
                        If Nvl(mrsBalance!����) = 5 Then
                            If mcllSquareBalance Is Nothing Then Set mcllSquareBalance = New Collection
                            'array(�����ID,���ѿ�ID,ˢ�����,����,����,�������,�Ƿ�����)
                            mcllSquareBalance.Add Array(lng�����ID, Val(Nvl(mrsBalance!���ѿ�ID)), _
                            Format(Val(Nvl(mrsBalance!��Ԥ��)), "0.00"), strCardNo, "", "", Val(Nvl(mrsBalance!�Ƿ�����)))
                        End If
                        .TextMatrix(i, .ColIndex("֧����ʽ")) = Nvl(mrsBalance!���㷽ʽ)
                        ' ҽ�ƿ����ID|���ѿ�(1, 0) |���ƿ�|�Ƿ�ȫ��|�Ƿ�����|�ӿ�����
                        .Cell(flexcpData, i, .ColIndex("֧����ʽ")) = lng�����ID & "|" & IIf(Val(Nvl(mrsBalance!����)) = 5, 1, 0) & "|" & Val(Nvl(mrsBalance!���ƿ�)) & "|" & Val(Nvl(mrsBalance!�Ƿ�ȫ��)) & "|" & Val(Nvl(mrsBalance!�Ƿ�����)) & "|" & Nvl(mrsBalance!���������)
                        .TextMatrix(i, .ColIndex("֧�����")) = Format(Val(Nvl(mrsBalance!��Ԥ��)), "0.00")
                        .TextMatrix(i, .ColIndex("�������")) = Nvl(mrsBalance!�������)
                        .TextMatrix(i, .ColIndex("��ע")) = Nvl(mrsBalance!ժҪ)
                        .TextMatrix(i, .ColIndex("������ˮ��")) = Nvl(mrsBalance!������ˮ��)
                        .TextMatrix(i, .ColIndex("����˵��")) = Nvl(mrsBalance!����˵��)
                        .TextMatrix(i, .ColIndex("����")) = IIf(Val(Nvl(mrsBalance!�Ƿ�����)) = 1, String(Len(strCardNo), "*"), strCardNo)
                        .Cell(flexcpData, i, .ColIndex("����")) = Nvl(mrsBalance!����)
                        mCurCarge.dbl�����Ѹ��ϼ� = RoundEx(mCurCarge.dbl�����Ѹ��ϼ� + Val(Nvl(mrsBalance!��Ԥ��)), 6)
                    End With
                End If
            Case Else '0-��ͨ����
                With vsBlance
                   If .TextMatrix(i, .ColIndex("֧����ʽ")) <> "" And Nvl(mrsBalance!���㷽ʽ) <> "" Then
                       .Rows = .Rows + 1
                       i = i + 1
                   End If
                   If Nvl(mrsBalance!���㷽ʽ) <> "" Then
                        .RowData(i) = Nvl(mrsBalance!����)
                        .TextMatrix(i, .ColIndex("֧����ʽ")) = Nvl(mrsBalance!���㷽ʽ)
                        .TextMatrix(i, .ColIndex("֧�����")) = Format(Val(Nvl(mrsBalance!��Ԥ��)), "0.00")
                        .TextMatrix(i, .ColIndex("�������")) = Nvl(mrsBalance!�������)
                        .TextMatrix(i, .ColIndex("��ע")) = Nvl(mrsBalance!ժҪ)
                        .TextMatrix(i, .ColIndex("������ˮ��")) = Nvl(mrsBalance!������ˮ��)
                        .TextMatrix(i, .ColIndex("����˵��")) = Nvl(mrsBalance!����˵��)
                        .TextMatrix(i, .ColIndex("����")) = IIf(Val(Nvl(mrsBalance!�Ƿ�����)) = 1, String(Len(strCardNo), "*"), strCardNo)
                        .Cell(flexcpData, i, .ColIndex("����")) = Nvl(mrsBalance!����)
                        mCurCarge.dbl�����Ѹ��ϼ� = RoundEx(mCurCarge.dbl�����Ѹ��ϼ� + Val(Nvl(mrsBalance!��Ԥ��)), 6)
                    End If
                End With
            End Select
            .MoveNext
        Loop
    End With
                   
    gstrSQL = "" & _
    "   Select  B.NO,B.����ID, Nvl(Sum(Nvl(B.Ӧ�ս��, 0)), 0)  As ����Ӧ�պϼ�, " & _
    "       Nvl(Sum(Nvl(B.ʵ�ս��, 0)), 0)  As ����ʵ�պϼ� " & _
    "   From ������ü�¼ B  " & _
    "   Where B.����id =[1]  " & _
    "   Group by B.NO,B.����ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mobjChargeInfor.����ID)
    With mCurCarge
         .dbl����ʵ�� = 0:
         .dbl����Ӧ�� = 0
        Do While Not rsTemp.EOF
            .dbl����ʵ�� = RoundEx(.dbl����ʵ�� + Val(Nvl(rsTemp!����ʵ�պϼ�)), 6)
            .dbl����Ӧ�� = RoundEx(.dbl����Ӧ�� + Val(Nvl(rsTemp!����Ӧ�պϼ�)), 6)
            rsTemp.MoveNext
        Loop
        .dbl��ǰδ�� = RoundEx(.dbl����ʵ�� - .dbl�����Ѹ��ϼ�, 6)
        If .dbl���γ�Ԥ�� <> 0 Then
            With vsBlance
                If .Rows = 2 Then .Row = 1
                If .Row < 0 Then .Row = 1
                i = .Row
                If Trim(.TextMatrix(.Row, .ColIndex("֧����ʽ"))) <> "" Then
                    .Rows = .Rows + 1
                    i = .Rows - 1
                End If
                ' 0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
                .RowData(i) = 1
                .TextMatrix(i, .ColIndex("֧����ʽ")) = "Ԥ���"
                .TextMatrix(i, .ColIndex("֧�����")) = Format(mCurCarge.dbl���γ�Ԥ��, "0.00")
            End With
        End If
        mblnYB�˿� = mCurCarge.dbl��ǰδ�� + mCurCarge.dblӦ���ۼ� < 0 And blnYb
    End With
    
    vsBlance_AfterRowColChange 0, 0, vsBlance.Row, vsBlance.Col
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub InitFace()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���ؼ�
    '����:���˺�
    '����:2011-06-13 14:09:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, dbl��� As Double, rsTemp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
    With vsBlance
        .Cell(flexcpFontBold, 1, 0, 1, .COLS - 1) = True
        .Clear 1: .Rows = 2
    End With
    With mCurCarge
        .dbl���γ�Ԥ�� = 0
        .dbl����ʵ�� = 0
        .dbl����ҽ��֧�� = 0
        .dbl�����Ѹ��ϼ� = 0
        .dbl����Ӧ�� = 0
        .dbl��ǰδ�� = 0
        .dbl������� = 0
        .dbl����Ԥ�� = 0
        .dblԤ����� = 0
    End With
    With mCurCardPay
        .lng���ѿ�ID = 0
        .str������� = ""
        .dbl��ˢ��� = 0
        .strˢ������ = ""
        .strˢ������ = ""
    End With
    mstr��֧Ʊ = ""
    strSQL = " " & _
    " Select B.���� " & _
    " From ���㷽ʽӦ�� A, ���㷽ʽ B " & _
    " Where A.Ӧ�ó��� = '�շ�' And B.���� = A.���㷽ʽ And a.���ʽ Is Null" & _
    "       And Nvl(B.Ӧ����, 0) = 1 And Rownum <= 1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If Not rsTemp.EOF Then
        mstr��֧Ʊ = Nvl(rsTemp!����)
    End If
    
    Call initInsure
    If mbytFunc = EM_FUN_�շ� Then
        Call InitBalanceData
    Else
        If mobjChargeInfor.intInsure <> 0 Then   'ҽ������ʱ,�쳣����һ�㶼�ǽ����˵�.
            strSQL = "Select 1" & _
                    " From ����Ԥ����¼ A, ���㷽ʽ B" & _
                    " Where a.���㷽ʽ = b.���� And b.���� In (3, 4) And ����id = [1] " & _
                    "       And Nvl(У�Ա�־, 0) = 1 And Rownum < 2"
            strSQL = strSQL & "Union All" & _
                    " Select 1" & _
                    " From ���ս����¼" & _
                    " Where ��¼id = [1] " & _
                    "       And Not Exists(Select 1 From ����Ԥ����¼ A, ���㷽ʽ B" & _
                    "                       Where a.���㷽ʽ = b.���� And b.���� In (3, 4) And a.����id = ��¼id)" & _
                    "       And Rownum < 2"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mobjChargeInfor.����ID)
            'У�Ա�־����2���ѳɹ�����
            '91914,�൥�ݷֵ��ݽ��㲻֧��Ԥ����ʱ����Ԥ����¼���п���û��ҽ��������Ϣ
            mblnYbBalanced = rsTemp.EOF
        End If
        Call LoadData
    End If
    Call Load֧����ʽ: Call LoadPatiInfor
    Call SetDeleteVisible '����������ʱɾ����ťӦ�ø��������ʾ
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub SetControlProperty(Optional blnԤ�� As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ÿؼ�����
    '����:blnԤ��-�Ƿ���������Ԥ��
    '����:���˺�
    '����:2011-08-12 10:43:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim sngTop As Long, sngSplitHeight As Single, dbl�ֽ� As Double
    Dim bln�ֱ� As Boolean, dblMoney As Double
    Dim bln�˿� As Boolean '��Ҫ��ҽ����ؽ�������˵����շ�
    Dim blnVisible As Boolean
    
    sngSplitHeight = 80
    
    '51670
    If mobjChargeInfor.����ID = 0 Or mbln�������� Then
        lbl��Ԥ��.Visible = False
        txt��Ԥ��.Visible = False
        txt��Ԥ��.Text = "0"
    End If
    
    blnVisible = mbytFunc = EM_FUN_�շ�  '����Ϊ�����շ�
    ' 0-�������нɿ�������ۼƿ���,1-��������ɿ��Ž��������ۼ�(�ı䲡�˳���)��2-�շ�ʱ����Ҫ����ɿ���
    ' 3-�շ�ʱ,�������˽����ۼ�(���ǰ����շѵ�����շѹ��ܻ�ı䲡��ʱ)
    blnVisible = blnVisible And (gTy_Module_Para.byt�ɿ���� = 1 Or gTy_Module_Para.byt�ɿ���� = 3)
    blnVisible = blnVisible And Val(txt��Ԥ��.Text) = 0 'δʹ��Ԥ����
    blnVisible = blnVisible And Val(txt�ɿ�.Text) = 0 'δ����ɿ���
    blnVisible = blnVisible And (mCurCarge.dbl�����Ѹ��ϼ� - mCurCarge.dbl����ҽ��֧��) = 0 'ȫ��Ϊҽ��֧��ʱ
    
    'δʹ������������
    blnVisible = blnVisible And mCurCardPay.lngҽ�ƿ����ID = 0 And mCurCardPay.blnOneCard = False
    '��ͨ���˻��ֻʹ����ҽ������
    blnVisible = blnVisible And (mobjChargeInfor.intInsure = 0 Or mobjChargeInfor.intInsure <> 0 And mblnYbBalanced)
    cmdNext.Visible = blnVisible
        
    lbl�ѽ�.Caption = "�Ѹ��ϼ�:" & Format(mCurCarge.dbl�����Ѹ��ϼ�, "###0.00;-###0.00;0.00;0.00;")
    
    If mCurCardPay.int���� = 1 And blnԤ�� = False Then
        dblMoney = mCurCarge.dbl��ǰδ�� + mCurCarge.dblӦ���ۼ�
        If mobjChargeInfor.intInsure > 0 Then  '����:43855,44069
            If mInsurePara.�ֱҴ��� Then
                bln�ֱ� = True
                dbl�ֽ� = CentMoney(CCur(dblMoney))
            Else
                dbl�ֽ� = Format(dblMoney, "0.00")
            End If
        Else
             bln�ֱ� = True
            dbl�ֽ� = RoundEx(CentMoney(CCur(dblMoney)), 6)
        End If
        lblʣ���Ը�.Caption = Format(dbl�ֽ�, "0.00")
    Else
        lblʣ���Ը�.Caption = Format(mCurCarge.dbl��ǰδ�� + mCurCarge.dblӦ���ۼ�, "0.00")
    End If
    
    '����:58344
    '   ����Ƿ�ǰ֧�����Ϊ����,�Ǹ���ʱ,��Ҫ���Ѳ���Ա(��Ҫ��ҽ������ʱ���ܴ��ڱ����ݵķ���)
    If Not mblnYB�˿� Then
        lblPayType.Caption = "�ɡ���"
        lblPayType.ForeColor = &H80000008
        cbo֧����ʽ.ForeColor = &H80000008
        txt�ɿ�.ForeColor = &H80000008
        lbl�Ҳ�.Caption = "��  ��"
    Else
        lblPayType.Caption = "�ˡ���"
        lblPayType.ForeColor = vbRed
        cbo֧����ʽ.ForeColor = vbRed
        txt�ɿ�.ForeColor = vbRed
        lbl�Ҳ�.Caption = "��  ��"
        '�˿�ʱ��������Ԥ��
        txt��Ԥ��.Visible = False: lbl��Ԥ��.Visible = False
        mblnNotChange = True
        txt��Ԥ��.Text = "0"
        mblnNotChange = False
    End If
    
    If blnԤ�� Then
        'Ԥ���Ĵ���
        lbl�Ҳ�.Visible = False: txt�Ҳ�.Visible = False
        txt�Ҳ�.Text = 0
    ElseIf mCurCardPay.int���� = 1 Then
        lbl�Ҳ�.Visible = True: txt�Ҳ�.Visible = True
'        lbl�Ҳ�.Caption = "�ҡ���"
        If IIf(mblnYB�˿�, -1, 1) * Val(txt�ɿ�.Text) >= dbl�ֽ� Then
            lbl�Ҳ�.ForeColor = &H80000008
            txt�Ҳ�.ForeColor = &H80000008
        Else
            lbl�Ҳ�.ForeColor = vbRed
            txt�Ҳ�.ForeColor = vbRed
        End If
        
        If bln�ֱ� Then
            dblMoney = CentMoney(CCur(mCurCarge.dbl��ǰδ��))
        Else
            dblMoney = mCurCarge.dbl��ǰδ��
        End If
        '61611
        txt�Ҳ�.Text = Format(IIf(mblnYB�˿�, -1, 1) * Val(txt�ɿ�.Text) - dblMoney - mCurCarge.dblӦ���ۼ�, "0.00")
        txt�������.Visible = False: lbl�������.Visible = False
        
    ElseIf mCurCardPay.bln֧Ʊ Then
        If mblnYB�˿� Then
            '58344
            lbl�Ҳ�.Visible = False
            txt�Ҳ�.Visible = False
            txt�Ҳ�.Text = 0
        Else
            If RoundEx(Val(txt�ɿ�.Text), 6) > RoundEx(mCurCarge.dbl��ǰδ�� + mCurCarge.dblӦ���ۼ�, 6) Then
                  lbl�Ҳ�.Visible = True: txt�Ҳ�.Visible = True
                  lbl�Ҳ�.Caption = "  �� ֧ Ʊ"
                  txt�Ҳ�.Text = Format(Val(txt�ɿ�.Text) - RoundEx(mCurCarge.dbl��ǰδ��, 2) - mCurCarge.dblӦ���ۼ�, "0.00")
                  txt�Ҳ�.ForeColor = vbRed
                  lbl�Ҳ�.ForeColor = vbRed
            Else
                  lbl�Ҳ�.Visible = False
                  txt�Ҳ�.Visible = False
                  txt�Ҳ�.Text = 0
            End If
        End If
         txt�������.Visible = True
         lbl�������.Visible = True
    ElseIf cbo֧����ʽ.Text Like "*��*" And mCurCardPay.lngҽ�ƿ����ID = 0 Then
         txt�������.Visible = True
         lbl�������.Visible = True
        lbl�Ҳ�.Visible = False
        txt�Ҳ�.Visible = False
        txt�Ҳ�.Text = 0
    Else
        lbl�Ҳ�.Visible = False
        txt�Ҳ�.Visible = False
        txt�������.Visible = False: lbl�������.Visible = False
    End If
    sngTop = txt��Ԥ��.Top
    If txt��Ԥ��.Visible Then
        sngTop = txt��Ԥ��.Top + txt��Ԥ��.Height + sngSplitHeight
    End If
    cbo֧����ʽ.Top = sngTop
    txt�ɿ�.Top = sngTop
    lblPayType.Top = sngTop + (cbo֧����ʽ.Height - lblPayType.Height) \ 2
    sngTop = sngTop + cbo֧����ʽ.Height + sngSplitHeight
    If lbl�Ҳ�.Visible Then
        txt�Ҳ�.Top = sngTop
        lbl�Ҳ�.Top = sngTop + (txt�Ҳ�.Height - lbl�Ҳ�.Height) \ 2
        sngTop = sngTop + txt�Ҳ�.Height + sngSplitHeight
    End If
    If txt�������.Visible Then
        txt�������.Top = sngTop
        lbl�������.Top = sngTop + (txt�������.Height - lbl�������.Height) \ 2
        sngTop = sngTop + txt�������.Height + sngSplitHeight
    End If
     
    txtժҪ.Top = sngTop
    lblժҪ.Top = sngTop + 25
    txtժҪ.Height = picPay.Height - sngTop - 100
    If mbytFunc = 1 Then
        txt��Ԥ��.BackColor = Me.BackColor
        txt�ɿ�.BackColor = Me.BackColor
        txt�������.BackColor = Me.BackColor
        txtժҪ.BackColor = Me.BackColor
        cbo֧����ʽ.BackColor = Me.BackColor
        txt�Ҳ�.BackColor = Me.BackColor
        txt�Ҳ�.Text = ""
    End If
 
End Sub
Private Sub cbo֧����ʽ_Click()
    Dim i As Long, varData As Variant, varTemp As Variant
    Dim lngIndex As Long
    
    If mblnFirst Then Exit Sub
    txt�ɿ�.Text = "": txt�ɿ�.Locked = False
    With mCurCardPay
       .lngҽ�ƿ����ID = 0
       .bln���ѿ� = False
       .str���㷽ʽ = ""
       .lng���ѿ�ID = 0
       .str���� = ""
       .strˢ������ = ""
       .strˢ������ = ""
       .lngID = 0
       .strNo = ""
       .str���� = ""
       .bln�������� = False
       .intҽ�ƿ����� = 0
       .bln���� = False
       .bln֧Ʊ = False
       .blnOneCard = False
       .int���� = 0
       .bln���ƿ� = False
       .str������ˮ�� = ""
       .str����˵�� = ""
    End With
    With cbo֧����ʽ
        If .ListIndex = -1 Then GoTo SetProperty:
        lngIndex = .ListIndex + 1
        mCurCardPay.int���� = .ItemData(.ListIndex)
        mCurCardPay.blnOneCard = .ItemData(.ListIndex) = 7
        mCurCardPay.bln֧Ʊ = False
        If .ItemData(.ListIndex) = 2 And cbo֧����ʽ.Text Like "*֧Ʊ*" Then
             mCurCardPay.bln֧Ʊ = True
        End If
    End With
    '��|ȫ��|ˢ����־|�����ID(���ѿ����)|����|�Ƿ����ѿ�|���㷽ʽ|���Ĺ���|�Ƿ����ƿ�;��
    If Not mcolCardPayMode Is Nothing Then
        With mCurCardPay
            .lngҽ�ƿ����ID = Val(mcolCardPayMode(lngIndex)(3))
            .bln���ѿ� = Val(mcolCardPayMode(lngIndex)(5)) = 1
            .str���㷽ʽ = Trim(mcolCardPayMode(lngIndex)(6))
            .str���� = Trim(mcolCardPayMode(lngIndex)(1))
            .bln���� = Val(mcolCardPayMode(lngIndex)(2)) = 0
            If .lngҽ�ƿ����ID <> 0 Then .bln֧Ʊ = False: .blnOneCard = False
            .bln���ƿ� = Val(mcolCardPayMode(lngIndex)(8)) = 1
            .bln�������� = Trim(mcolCardPayMode(lngIndex)(7)) <> "" And Trim(mcolCardPayMode(lngIndex)(7)) <> "0"
            If .bln���ѿ� Or (.int���� <> 1 And mblnYB�˿�) Then
                '57682:ȱʡΪ����֧�����
                txt�ɿ�.Text = Format(IIf(mblnYB�˿�, -1, 1) * Val(lblʣ���Ը�.Caption), "0.00")
            End If
         End With
     Else
         mCurCardPay.str���㷽ʽ = zlstr.NeedName(cbo֧����ʽ.Text)
     End If
     If mCurCardPay.blnOneCard Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = New clsICCard
            Call mobjICCard.SetParent(Me.hWnd)
            Set mobjICCard.gcnOracle = gcnOracle
        End If
     End If
SetProperty:
     Call SetControlProperty
     Call Show�����(False)
     If Not mcolCardPayMode Is Nothing Then
        If mCurCardPay.bln���ѿ� Or (mCurCardPay.int���� <> 1 And mblnYB�˿�) Then
            '57682:ȱʡΪ����֧�����
            txt�ɿ�.Text = Format(IIf(mblnYB�˿�, -1, 1) * Val(lblʣ���Ը�.Caption), "0.00")
        ElseIf mCurCardPay.lngҽ�ƿ����ID > 0 And Not mCurCardPay.bln���ѿ� Then
            If gTy_Module_Para.bytˢ��ȱʡ������ <> 0 Then
                txt�ɿ�.Text = Format(IIf(mblnYB�˿�, -1, 1) * Val(lblʣ���Ը�.Caption), "0.00")
                '�������޸�
                If gTy_Module_Para.bytˢ��ȱʡ������ = 2 Then txt�ɿ�.Locked = True
            End If
        End If
     End If
     If txt�ɿ�.Enabled Then txt�ɿ�.SetFocus
End Sub

Private Function CheckOneCard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���һ��ͨ�Ƿ���ȷ
    '����:һ��ͨ��֤��ȷ���һ��ͨ,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-23 17:07:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim CurOneCard As Currency, dblMoney As Double
    
    If mCurCardPay.blnOneCard = False Then CheckOneCard = True: Exit Function
    If mobjICCard Is Nothing Then
        Set mobjICCard = New clsICCard
        Call mobjICCard.SetParent(Me.hWnd)
        Set mobjICCard.gcnOracle = gcnOracle
    End If
    If mobjICCard Is Nothing Then
        MsgBox "һ��ͨ�ӿڴ���ʧ��!", vbOKOnly, gstrSysName
        Exit Function
    End If
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
        Optional ByRef bln���� As Boolean = False, _
        Optional ByVal bln�����ֹ As Boolean = True, _
        Optional ByRef varSquareBalance As Variant, _
        Optional ByVal blnתԤ�� As Boolean = False, _
        Optional ByVal blnAllPay As Boolean = False, _
        Optional ByVal strXmlIn As String = "") As Boolean
        '       strXmlIn-����������XML���,Ŀǰ��ʽ����:
        '       <IN>
        '           <CZLX>0</CZLX>    //��������,0-��������ˢ��,1-ת�˵���ˢ��,2-�˿����ˢ��
        '       </IN>
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ָ��֧�����,����ˢ������
    '���:rsClassMoney:�շ����,���
    '        lngCardTypeID-Ϊ��ʱ,Ϊ��һ��ͨˢ��
    '       bln�����ֹ-Ŀǰֻ������ѿ�,��ʾ����ʱ,��ֹ��������,������������֧��
    dblMoney = Val(txt�ɿ�.Text)
    If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModule, Nothing, 0, False, _
    mobjChargeInfor.����, mobjChargeInfor.�Ա�, mobjChargeInfor.����, dblMoney, mCurCardPay.strˢ������, mCurCardPay.strˢ������, _
    False, True, False, False, Nothing, False, False, "<IN><CZLX>2</CZLX></IN>") = False Then Exit Function
 
    CurOneCard = mobjICCard.GetSpare
    If CurOneCard < Val(txt�ɿ�.Text) Then
        MsgBox "������֧��,����!" & vbCrLf & vbCrLf & _
        "   �� ��  ��" & Format(CurOneCard, "0.00") & vbCrLf & _
        "   ����֧��" & Format(Val(txt�ɿ�.Text), "0.00"), vbInformation, gstrSysName
        Exit Function
    End If
    
    stbThis.Panels(4).Text = Format(CurOneCard, "0.00")
    stbThis.Panels(4).ToolTipText = mCurCardPay.str���㷽ʽ & "���ʻ����:" & Format(CurOneCard, "0.00")
    '�Ѿ�������֧�����
    If dblMoney <> Val(txt�ɿ�.Text) Then
        txt�ɿ�.Text = Format(dblMoney, "0.00")
    End If
    CheckOneCard = True
End Function

Private Function CheckPrepayMoneyIsValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���Ԥ�����������Ƿ�Ϸ�
    '����:�Ϸ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-24 10:36:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    '������Ӧ��
    If Not mbln�ѱ��� Then Call LedVoiceSpeak
    
    If BrushcardStrikePrepay = False Then
        If txt��Ԥ��.Enabled And txt��Ԥ��.Visible Then txt��Ԥ��.SetFocus
        zlControl.TxtSelAll txt��Ԥ��
        Call SetControlProperty(True)
        Exit Function
    End If
    CheckPrepayMoneyIsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function isValied(Optional bln���� As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����շ�����ʱ����Ч��,������Ч,����true,���򷵻�False
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-13 16:30:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    
    If Not CheckTextLength("�������", txt�������) Then Exit Function
    If Not CheckTextLength("ժҪ", txtժҪ) Then Exit Function
    
    '������Ӧ��
    If Not mbln�ѱ��� Then Call LedVoiceSpeak
    txt�ɿ�.Text = Format(Val(txt�ɿ�.Text), "0.00")
    If BrushCardThreeSwapCheck = False Then
        If txt�ɿ�.Enabled And txt�ɿ�.Visible Then txt�ɿ�.SetFocus
        zlControl.TxtSelAll txt�ɿ�
        Exit Function
    End If
    'һ��ͨˢ��
    If CheckOneCard = False Then
        If txt�ɿ�.Enabled And txt�ɿ�.Visible Then txt�ɿ�.SetFocus
        zlControl.TxtSelAll txt�ɿ�
        Exit Function
    End If
    
    If CheckInterfaceNumIsValied = False Then
        If txt�ɿ�.Enabled And txt�ɿ�.Visible Then txt�ɿ�.SetFocus
        zlControl.TxtSelAll txt�ɿ�
        Exit Function
    End If
    If mCurCardPay.int���� = 1 Then
        '�ֽ�֧��,��Ҫ���ڱ���¼����
        'ֻ���ֽ�Ŵ���,�Ҳ�
        '�����շ�:
        '���˺�:22343,�ɿ������
        Select Case gTy_Module_Para.byt�ɿ����
        Case 1, 3 '1-�ಡ�ɿ�;3�����˽ɿ��ۼ�
            If mblnCur���� = False Then
                If RoundEx(mCurCarge.dbl��ǰδ��, 2) > 0 And RoundEx(Val(txt�ɿ�.Text), 2) = 0 Then
                   If MsgBox("ע��:" & vbCrLf & "    �ò���δ����ɿ���,�Ƿ�����շ�? ", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                       If txt�ɿ�.Enabled And txt�ɿ�.Visible Then txt�ɿ�.SetFocus
                       zlControl.TxtSelAll txt�ɿ�
                       Exit Function
                   End If
                End If
            End If
        Case 2  '2-�շ�ʱ����Ҫ����ɿ���
            If RoundEx(mCurCarge.dbl��ǰδ��, 2) > 0 And RoundEx(Val(txt�ɿ�.Text), 2) = 0 Then
                MsgBox "ע��:" & vbCrLf & _
                "    �ò���δ����ɿ���,���ܽ����շ�!", vbInformation + vbDefaultButton1, gstrSysName
                If txt�ɿ�.Enabled And txt�ɿ�.Visible Then txt�ɿ�.SetFocus
                zlControl.TxtSelAll txt�ɿ�
                Exit Function
            End If
        Case Else   ',0-�������нɿ�������ۼƿ���
            'ҽ������ɿ���:Ҫ�ɶ�δ��ʱ,�Խɿ���Ϊ������������,��Ϊ��ǿ������0�����ɿ��
            If mstrYBPati <> "" And Not mInsurePara.���������շ� And RoundEx(mCurCarge.dbl��ǰδ��, 6) > 0 And Val(txt�ɿ�.Text) = 0 Then
                '���˺�:27536 20100119
                If mInsurePara.�����ѽɿ���� = False Then MsgBox "������:" & vbCrLf & vbTab & "��ҽ�����˵ķ���δȫ�����㣬��ע����ȡ���˽ɿ", vbInformation, gstrSysName
            End If
        End Select
        
        If Val(txt�ɿ�.Text) <> 0 Then
            If mblnYB�˿� Then
                If CSng(txt�Ҳ�.Text) > 0 Then
                    MsgBox "�˿���㣬�벹���˿��", vbInformation, gstrSysName
                    txt�ɿ�.SetFocus: zlControl.TxtSelAll txt�ɿ�
                    Exit Function
                End If
            Else
                If CSng(txt�Ҳ�.Text) < 0 Then
                    MsgBox "�ɿ���㣬�벹��ɿ��", vbInformation, gstrSysName
                    txt�ɿ�.SetFocus: zlControl.TxtSelAll txt�ɿ�
                    Exit Function
                End If
            End If
        End If
    Else
        If mblnCur���� = False Then
            If Val(txt�ɿ�) = 0 Then
                MsgBox "δ���뽻�׽��,����!", vbInformation + vbOKOnly, gstrSysName
                If txt�ɿ�.Enabled And txt�ɿ�.Visible Then txt�ɿ�.SetFocus
                zlControl.TxtSelAll txt�ɿ�: Exit Function
            End If
            If Not IsNumeric(txt�ɿ�.Text) And txt�ɿ�.Text <> "" Then
                MsgBox "��Ч��ֵ��", vbInformation, gstrSysName
                If txt�ɿ�.Enabled And txt�ɿ�.Visible Then txt�ɿ�.SetFocus
                zlControl.TxtSelAll txt�ɿ�: Exit Function
            ElseIf Val(txt�ɿ�.Text) < 0 Then
                MsgBox "���׽���Ϊ����", vbInformation, gstrSysName
                If txt�ɿ�.Enabled And txt�ɿ�.Visible Then txt�ɿ�.SetFocus
                zlControl.TxtSelAll txt�ɿ�: Exit Function
            End If
        End If
        If Not mCurCardPay.bln֧Ʊ Then
            '����:42793
            '�������㷽ʽ,����Ľ��ܴ���δ������
            If RoundEx(Abs(Val(txt�ɿ�.Text)), 2) > RoundEx(Abs(mCurCarge.dbl��ǰδ�� + mCurCarge.dblӦ���ۼ�), 2) Then
                MsgBox "ע��:" & vbCrLf & "    �����" & IIf(mblnYB�˿�, "�˿�", "�ɿ�") & "��������δ" & IIf(mblnYB�˿�, "��", "֧��") & "�Ľ����ܼ�����", vbOKOnly + vbInformation, gstrSysName
                txt�ɿ�.SetFocus: zlControl.TxtSelAll txt�ɿ�
                Exit Function
            End If
        End If
        If Val(txt�ɿ�.Text) <> 0 And mCurCarge.dblӦ���ۼ� <> 0 Then '��������շ�ʱ���������С��δ�����
            If RoundEx(Abs(Val(txt�ɿ�.Text)), 2) < RoundEx(Abs(mCurCarge.dbl��ǰδ�� + mCurCarge.dblӦ���ۼ�), 2) Then
                MsgBox IIf(mblnYB�˿�, "�˿�", "�ɿ�") & "���㣬�벹��" & IIf(mblnYB�˿�, "�˿�", "�ɿ�") & "��", vbInformation, gstrSysName
                txt�ɿ�.SetFocus: zlControl.TxtSelAll txt�ɿ�
                Exit Function
            End If
        End If
    End If

    '��鵱ǰ�����Ƿ�������ִ�����,��Ҫ�ǲ���ԭ����м��
    '��ֹ��������Ա����:
    '45186
    gstrSQL = "" & _
    "   Select  1  From ����Ԥ����¼ A " & _
    "   Where   A.����ID=[1] and nvl(A.У�Ա�־,0)<>0 and Rownum =1 and A.��¼״̬=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mobjChargeInfor.����ID)
    If rsTemp.EOF Then
        '�����Ǳ�����ִ��,������Ҫ����Ƿ�����ִ��
        gstrSQL = "Select ��¼״̬, ����Ա����,����״̬ From ������ü�¼ Where ����ID=[1] And rownum=1"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mobjChargeInfor.����ID)
        
        If Not rsTemp.EOF Then
            If Val(Nvl(rsTemp!��¼״̬)) <> 1 Then
                MsgBox "�õ����Ѿ�����������Ա����,�����ٽ����շ�!", vbOKOnly + vbInformation, gstrSysName
                'ִ���շ�
                Unload Me
                Exit Function
            End If
            
            If Val(Nvl(rsTemp!����״̬)) <> 1 Then
                MsgBox "�ô��շ��Ѿ��������շ�,�����ٽ����շ�!", vbOKOnly + vbInformation, gstrSysName
                'ִ���շ�
                Unload Me
                Exit Function
            End If
            
            If Nvl(rsTemp!����Ա����) <> UserInfo.���� Then
                MsgBox "�õ��ݲ��Ǳ����շѵ�,������ȡ��������Ա�ĵ���!", vbOKOnly + vbInformation, gstrSysName
                'ִ���շ�
                Unload Me
                Exit Function
            End If
        End If
    End If
    
    isValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckInterfaceNumIsValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ӿ������Ƿ񳬹�2������
    '����:δ����2������,����true,���򷵻�False
    '����:���˺�
    '����:2012-02-27 15:23:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngCount As Long, varData As Variant
    Dim strNames As String, i As Long
    
    On Error GoTo errHandle
    
    lngCount = IIf(mobjChargeInfor.intInsure <> 0, 1, 0)   'ҽ����һ������
    If mCurCardPay.lngҽ�ƿ����ID = 0 Or (mCurCardPay.bln���ѿ� And mCurCardPay.bln���ƿ�) Then CheckInterfaceNumIsValied = True: Exit Function
    With vsBlance
        strNames = IIf(mobjChargeInfor.intInsure <> 0, vbCrLf & "ҽ������", "")
        For i = 1 To .Rows - 1
            '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
            If Val(.RowData(i)) = 3 Or Val(.RowData(i)) = 4 Or Val(.RowData(i)) = 5 Then
                ' ҽ�ƿ����ID|���ѿ�(1, 0) |���ƿ�|�Ƿ�ȫ��|�Ƿ�����|�ӿ�����
                 varData = Split(.Cell(flexcpData, i, .ColIndex("֧����ʽ")) & "|||||", "|")
                 If Val(varData(0)) <> 0 Then
                    If Val(varData(1)) <> 1 Then
                        lngCount = lngCount + 1
                        strNames = strNames & vbCrLf & varData(5)
                    ElseIf Val(varData(2)) = 0 Then
                        '���ѿ�Ҳ�ǽӿڵ�,�������������ӿ�
                        lngCount = lngCount + 1
                        strNames = strNames & vbCrLf & varData(5)
                    End If
                End If
            End If
        Next
    End With
    If lngCount >= 2 Then
        MsgBox "  ϵͳ��ֻ֧���������ڵĽӿ�,������ˢ������," & vbCrLf & "  ����Ϊ��ǰ�Ѿ�ˢ�Ľӿ�!" & vbCrLf & strNames, vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    CheckInterfaceNumIsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckDelValied(ByRef blnExistThreeSwap As Boolean, _
    ByRef blnȫ�� As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����˷�
    '����:blnExistThreeSwap-�Ƿ���������ӿ�
    '        blnȫ��-���������ӿ��Ƿ����ȫ��
    '����:�Ϸ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-25 16:14:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double
    
    blnȫ�� = False: blnExistThreeSwap = False
    On Error GoTo errHandle
    
    mrsBalance.Filter = "  ����=3   OR ����=4    "
    If mrsBalance.EOF Then mrsBalance.Filter = 0: CheckDelValied = True: Exit Function
    With mrsBalance
        Do While Not .EOF
            dblMoney = RoundEx(Val(Nvl(!��Ԥ��)), 6)
            '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
            Select Case Nvl(!����)
            Case 3  'һ��ͨ(��)
                If Val(Nvl(!У�Ա�־)) = 2 Then
                    '����
                    If zlCheckDelValied(Val(Nvl(!�����ID)), CStr(Nvl(!���������)), False, _
                        Nvl(!����), Nvl(!������ˮ��), Nvl(!����˵��), mobjChargeInfor.����ID, dblMoney, _
                        Val(Nvl(!�Ƿ��˿��鿨)) = 1) = False Then Exit Function
                    If Not blnȫ�� Then blnȫ�� = Val(Nvl(!�Ƿ�ȫ��)) = 1
                    
                End If
            Case 4  'һ��ͨ(��)
                If Val(Nvl(!У�Ա�־)) = 2 Then
                    If CheckDelOneCardValied(Nvl(!����), dblMoney) = False Then Exit Function
                    If Not blnȫ�� Then blnȫ�� = True
                End If
            End Select
            .MoveNext
        Loop
    End With
    blnExistThreeSwap = True

    CheckDelValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckDelOneCardValied(ByVal strԭ���� As String, ByVal dblMoney As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���һ��ͨ�˷ѵ���Ч��
    '����:�Ϸ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-25 16:48:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCardNo As String
    On Error GoTo errHandle
    If mobjICCard Is Nothing Then
        On Error Resume Next
        Set mobjICCard = CreateObject("zlICCard.clsICCard")
        On Error GoTo 0
    End If
    If mobjICCard Is Nothing Then
        MsgBox "һ��ͨ�ӿڴ���ʧ��,���ܽ����˷�!����ӿ��ļ�.", vbInformation, gstrSysName
        Exit Function
    End If
    strCardNo = mobjICCard.Read_Card(Me)
    If strCardNo = "" Then
        MsgBox "һ��ͨ����ʧ��,�뽫IC�����ڶ�������", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    If strCardNo <> strԭ���� Then
        MsgBox "��ǰ������ۿ�Ų�һ��,���ܽ����˷�.", vbInformation, gstrSysName
        Exit Function
    End If
    CheckDelOneCardValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function

Private Sub cbo֧����ʽ_GotFocus()
    ClearԤ����
End Sub

Private Sub cbo֧����ʽ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Function CancelBalance(ByRef blnUnload As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�쳣����
    '����:blnUnload-�Ƿ�ִ��unload me
    '����:���˺�
    '����:2014-06-19 14:42:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dtDelDate As Date, lng����ID As Long
    Dim strInvoice As String, strNo As String, strSQL As String
    Dim cllPro As Collection, varData As Variant, i As Long
    Dim blnIsExiseThreeSwap As Boolean, blnȫ�� As Boolean
    Dim blnCommit As Boolean
    
    dtDelDate = zlDatabase.Currentdate
    blnUnload = False
    'һ��ͨ;���������׵ļ��
    If CheckDelValied(blnIsExiseThreeSwap, blnȫ��) = False Then
        If MsgBox("ע��:" & vbCrLf & "���������Ľ��е����������˷�,�Ƿ���ͣ����?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        blnUnload = True: Exit Function
    End If
    
    If mobjChargeInfor.intInsure <> 0 And mInsurePara.ҽ���ӿڴ�ӡƱ�� Then
        If zlCheckInvoiceValied(lng����ID, 1, , mobjChargeInfor.ShareUserID, mobjChargeInfor.PatiUseType) = False Then
            If MsgBox("ע��:" & vbCrLf & "    ����ЧƱ��,�Ƿ���ͣ����?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            blnUnload = True: Exit Function
        End If
        strInvoice = GetNextBill(lng����ID)
    End If
    
    mobjChargeInfor.�շѽ��� = "": mbln�������� = False
    '�������ϴ���
    Set cllPro = New Collection
    If mobjChargeInfor.Nos = "" And mobjChargeInfor.����ID <> 0 Then
        mobjChargeInfor.Nos = zlGetBalanceNos(1, mobjChargeInfor.����ID, False)
     End If
    
    varData = Split(Replace(mobjChargeInfor.Nos, "'", ""), ",")
    If Not mbln�쳣���� And Not mblnCommitBill Then
        mobjChargeInfor.����ID = zlDatabase.GetNextId("���˽��ʼ�¼")
        
        For i = UBound(varData) To 0 Step -1
            strNo = varData(i)
            'Zl_�����շѼ�¼_����
            strSQL = "Zl_�����շѼ�¼_����("
            '  No_In         ������ü�¼.No%Type,
            strSQL = strSQL & "'" & varData(i) & "',"
            '  ����Ա���_In ������ü�¼.����Ա���%Type,
            strSQL = strSQL & "'" & UserInfo.��� & "',"
            '  ����Ա����_In ������ü�¼.����Ա����%Type,
            strSQL = strSQL & "'" & UserInfo.���� & "',"
            '  ���_In       Varchar2 := Null,
            strSQL = strSQL & "NULL,"
            '  �˷�ʱ��_In   ������ü�¼.�Ǽ�ʱ��%Type := Null,
            strSQL = strSQL & "To_Date('" & Format(dtDelDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
            '  �˷�ժҪ_In   ������ü�¼.ժҪ%Type := Null,
            strSQL = strSQL & "'��������',"
            '  ����id_In     ����Ԥ����¼.����id%Type := Null,
            strSQL = strSQL & mobjChargeInfor.����ID & ","
            '  ����Ʊ��_In Number:=0
            strSQL = strSQL & 1 & ")"
            zlAddArray cllPro, strSQL
        Next
        '�Ȳ���Ʊ�ݣ�ҽ���ӿڲ���ȡ��
        If mInsurePara.ҽ���ӿڴ�ӡƱ�� And mobjChargeInfor.intInsure <> 0 Then
            strSQL = "zl_�����շѼ�¼_RePrint('" & strNo & "','" & strInvoice & "'," & ZVal(lng����ID) & ",'" & UserInfo.���� & "'," & _
                "To_Date('" & Format(dtDelDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),1,1)"
            zlAddArray cllPro, strSQL
        End If
        
        'ԭ����
        'Zl_�����˷ѽ���_Modify
        strSQL = "Zl_�����˷ѽ���_Modify("
        '  ��������_In   Number,
        strSQL = strSQL & "" & 0 & ","
        '  ����id_In     ������ü�¼.����id%Type,
        strSQL = strSQL & "" & mobjChargeInfor.����ID & ","
        '  ����id_In     ����Ԥ����¼.����id%Type,
        strSQL = strSQL & "" & mobjChargeInfor.����ID & ","
        '  ���㷽ʽ_In   Varchar2,
        strSQL = strSQL & "NULL)"
        zlAddArray cllPro, strSQL
 
    End If
    
    Err = 0: On Error GoTo Errhand:
    
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
     
    If mobjChargeInfor.intInsure <> 0 Then
        If ExcuteInsureDel(blnCommit) = False Then
            If blnCommit Then
                mblnCommitBill = True
                cmdExit.Visible = False
            End If
            Exit Function
        End If
        '�޸�У�Ա�־
        ' Zl_���������շ�_ҽ������
        strSQL = "Zl_���������շ�_ҽ������("
        '  ����id_In   ������ü�¼.����id%Type,
        strSQL = strSQL & mobjChargeInfor.����ID & ","
        '  �������_In ����Ԥ����¼.�������%Type,
        strSQL = strSQL & "Null,"
        '  ���ս���_In Varchar2
        strSQL = strSQL & "Null)"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        mblnCommitBill = True
        gcnOracle.CommitTrans: gcnOracle.BeginTrans
    End If
    
    On Error GoTo ErrInterface:
    
    '�����������˽���

    If ExcuteThreeSwapDel(mobjChargeInfor.����ID, mobjChargeInfor.����ID) = False Then
        If MsgBox("ע��:" & vbCrLf & "���������Ľ��е����������˷�,�Ƿ���ͣ����?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        blnUnload = True: Exit Function
    End If
    mblnCommitBill = True
    gcnOracle.CommitTrans: gcnOracle.BeginTrans
      
    If ExcuteOverFeeDel(mobjChargeInfor.����ID, mobjChargeInfor.����ID) = False Then Exit Function
    gcnOracle.CommitTrans
    
    mbytReturnMode = EM_EX_����
    CancelBalance = True
    blnUnload = True: Exit Function
Errhand:
    gcnOracle.RollbackTrans
ErrInterface:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ExcuteOverFeeDel(ByVal lng����ID As Long, ByVal lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����˷��շ�
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-29 14:50:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo errHandle
      
    'Zl_�����˷ѽ���_Modify
    strSQL = "Zl_�����˷ѽ���_Modify("
    '  ��������_In   Number,
    strSQL = strSQL & "" & 1 & ","
    '  ����id_In     ������ü�¼.����id%Type,
    strSQL = strSQL & "" & lng����ID & ","
    '  ����id_In     ����Ԥ����¼.����id%Type,
    strSQL = strSQL & "" & lng����ID & ","
    '  ���㷽ʽ_In   Varchar2,
    strSQL = strSQL & "NULL,"
    '  ��Ԥ��_In     ����Ԥ����¼.��Ԥ��%Type := Null,
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
    '  �����_In   ������ü�¼.ʵ�ս��%Type := Null,
    strSQL = strSQL & "NULL,"
    '  ����˷�_In   Number := 0,
    '  -- ����˷�_In:0-δ����˷�;1-�쳣����˷�;2-����˷�
    strSQL = strSQL & "1)"
    '  ԭ����id_In   ����Ԥ����¼.����id%Type := Null
    '�쳣����,����ҲӦ��Ϊ�쳣����
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    ExcuteOverFeeDel = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
End Function

Private Function ExcuteThreeSwapDel(ByVal lng����ID As Long, ByVal lngԭ����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�˷ѽ���(һ��ͨ���������㽻��)
    '����:���׳ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-25 17:29:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim cllUpdate As Collection, cllThreeSwap As Collection
    Dim strCardNo As String, i As Long, strSQL As String, strErrMsg As String
    Dim strSwapNO As String, strSwapMemo As String, varData As Variant
    Dim lng�����ID As Long, bln���ѿ� As Boolean, strTemp As String
    Dim st��������� As String, blnTrans As Boolean, dblMoney As Double
    Dim strҽԺ���� As String, rsTemp As ADODB.Recordset
    
    gstrSQL = "" & _
    "   Select A.���㷽ʽ,A.ժҪ, " & _
    "             A.�����ID,A.�������,A.����,A.������ˮ��, " & _
    "             nvl(C.�Ƿ�����,0) as ���ƿ�, " & _
    "             C.���� as ����,A.����˵��," & _
    "             Sum(A.��Ԥ��) as ��Ԥ��" & _
    "   From ����Ԥ����¼ A ,ҽ�ƿ���� C" & _
    "   Where A.����ID=[1] And nvl(A.У�Ա�־,0)=1  " & _
    "         And A.�����ID=C.ID" & _
    "   Group by A.���㷽ʽ,A.ժҪ,A.�����ID , A.�������,A.����,A.������ˮ��, " & _
    "           nvl(C.�Ƿ�����,0),C.����,A.����˵��,A.�������"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng����ID)
    Set cllUpdate = New Collection: Set cllThreeSwap = New Collection
    With rsTemp
        Do While Not .EOF
                lng�����ID = Val(Nvl(!�����ID))
                bln���ѿ� = False
                st��������� = Nvl(!����)
                strSwapNO = Nvl(!������ˮ��)
                strSwapMemo = Nvl(!����˵��)
                strCardNo = Nvl(!����)
                dblMoney = Nvl(!��Ԥ��)
                
                'Zl_����Ԥ����¼_����У�Ա�־
                strSQL = "Zl_����Ԥ����¼_����У�Ա�־("
                '  ����id_In     ������ü�¼.����id%Type,
                strSQL = strSQL & "" & lng����ID & ","
                '  �շѽ���_In   Varchar2,
                strSQL = strSQL & "'" & Nvl(!���㷽ʽ) & "',"
                '  �����id_In   ����Ԥ����¼.�����id%Type := Null,
                strSQL = strSQL & "" & lng�����ID & ","
                '  ���ѿ�_In     Integer := 0,
                strSQL = strSQL & "" & IIf(bln���ѿ�, 1, 0) & ","
                '  ����_In       ����Ԥ����¼.����%Type := Null,
                strSQL = strSQL & "'" & strCardNo & "',"
                '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
                strSQL = strSQL & "'" & strSwapNO & "',"
                '  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
                strSQL = strSQL & "'" & strSwapMemo & "',"
                '  У�Ա�־_In   ����Ԥ����¼.У�Ա�־%Type := 0
                strSQL = strSQL & "2)"
                zlDatabase.ExecuteProcedure strSQL, Me.Caption
                If CallBackBalanceInterface(lng����ID, lngԭ����ID, lng�����ID, bln���ѿ�, dblMoney, strCardNo, strSwapNO, strSwapMemo, cllUpdate, cllThreeSwap, strErrMsg) = False Then
                    gcnOracle.RollbackTrans: Exit Function
                End If
                gcnOracle.CommitTrans
                zlExecuteProcedureArrAy cllUpdate, Me.Caption
                zlExecuteProcedureArrAy cllThreeSwap, Me.Caption
                gcnOracle.BeginTrans
            .MoveNext
        Loop
    End With
    ExcuteThreeSwapDel = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog:
End Function

Private Function zlDelOneCard(ByVal strCardNo As String, ByVal strҽԺ���� As String, _
    ByVal str������ˮ�� As String, ByVal dblMoney As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��һ��ͨ����
    '����:strErrMsg-���صĴ�����Ϣ
    '����:
    '����:���˺�
    '����:2011-08-25 17:38:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If Not mobjICCard.ReturnSwap(strCardNo, strҽԺ����, str������ˮ��, dblMoney) Then
            MsgBox "һ��ͨ�˷ѽ��׵���ʧ��,�˷Ѳ���ʧ�ܣ�", vbExclamation, gstrSysName
            Exit Function
    End If
    zlDelOneCard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function ExcuteInsureDel(ByRef blnCommit As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ҽ���˷ѽӿ�
    '����:���óɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-25 12:17:56
    '˵��:��Ҫ�������������,�����˷Ѻ�,�ù��̲��ύ,��Ҫ�������ύ;
    '     ���ʧ��,�����񽫻���(��Ҫ�Ǳ��ⵯ�������������)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strAdvance As String, blnTransMedicare As Boolean
    Dim i As Long, p As Integer, strAdvanceOld As String
    Dim colBalance As Collection '��¼���ŵ��ݱ��ս���
    Dim strSQL As String
    Dim rsCharge As ADODB.Recordset, strNo As String
    Dim str���㷽ʽ As String, strDel���㷽ʽ As String
    On Error GoTo errHandle
    
    strAdvance = mobjChargeInfor.����ID & "|" & "0"
    
    blnTransMedicare = False
    If Not (mInsurePara.�൥�ݷֵ��ݽ��� Or mInsurePara.һ�ν���ֵ����˷�) Then
        If mblnCommitBill Then ExcuteInsureDel = True: Exit Function
        If Not gclsInsure.ClinicDelSwap(mobjChargeInfor.����ID, , mobjChargeInfor.intInsure, strAdvance) Then
             gcnOracle.RollbackTrans: Exit Function
        End If
        
        blnTransMedicare = True
        If strAdvance = mobjChargeInfor.����ID & "|" & "0" Or strAdvance = "" Then
            Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, True, mobjChargeInfor.intInsure)
            ExcuteInsureDel = True
            Exit Function
        End If
    Else
        Set colBalance = New Collection
        strAdvanceOld = strAdvance
        
        '93337,�˷�ʱ�����ݺŵ�����нӿڵ���
        strSQL = "Select Distinct NO From ������ü�¼ Where ����id = [1] Order By No Desc"
        Set rsCharge = zlDatabase.OpenSQLRecord(strSQL, "��ȡԭʼ���õ��ݺ�", mobjChargeInfor.����ID)
        
        p = 1
        Do While Not rsCharge.EOF
            colBalance.Add Array()
            strDel���㷽ʽ = "": str���㷽ʽ = ""
            strNo = Nvl(rsCharge!NO)
            '1.�����ŵ����Ƿ���Ҫ����ҽ������
            str���㷽ʽ = zlGetYBBalanceNo(mobjChargeInfor.����ID, strNo, mobjChargeInfor.����ID, _
                                    mobjChargeInfor.intInsure, True)
            
            '2.���õ����Ƿ���ҽ�����ϣ����ܴ��ڶ��ҽ���������ϣ�
            '������óɹ����ӿڣ���û���κ�ҽ�����ϣ������һ�ε���ҽ���ӿڣ���Ϊ�޷�ȷ���Ƿ���óɹ���
            strDel���㷽ʽ = zlGetYBBalanceNo(mobjChargeInfor.����ID, strNo)
            Call SetBalanceVal(colBalance, p, strDel���㷽ʽ)
                
            '3.����ҽ���˷ѽӿڣ��ύ����
            If str���㷽ʽ <> "" And strDel���㷽ʽ = "" Then
                '    Zl_ҽ��������ϸ_Insert(
                strSQL = "Zl_ҽ��������ϸ_Insert("
                '      ����id_In   ҽ��������ϸ.����id%Type,
                strSQL = strSQL & "" & mobjChargeInfor.����ID & ","
                '      No_In       ҽ��������ϸ.No%Type,
                strSQL = strSQL & "'" & strNo & "',"
                '      ���㷽ʽ_In Varchar2,
                strSQL = strSQL & "'" & str���㷽ʽ & "')"
                '      ��ע_In     ҽ��������ϸ.��ע%Type := Null
                zlDatabase.ExecuteProcedure strSQL, Me.Caption
                
                strAdvance = strAdvanceOld & "|" & strNo
                '��Ϊ�����̶�Ϊҽ������,�������ƹ̶�Ϊҽ������(����ͳ�ﲻ��ȷ��),�Ժ�Ӧȥ���ò���
                If Not gclsInsure.ClinicDelSwap(mobjChargeInfor.����ID, True, mobjChargeInfor.intInsure, _
                                                strAdvance) Then gcnOracle.RollbackTrans: Exit Function
                If strAdvance = strAdvanceOld & "|" & strNo Then strAdvance = ""
                
                If zlInsureCheck(str���㷽ʽ, strAdvance) Then
                    str���㷽ʽ = strAdvance
                    '    Zl_ҽ��������ϸ_Insert(
                    strSQL = "Zl_ҽ��������ϸ_Insert("
                    '      ����id_In   ҽ��������ϸ.����id%Type,
                    strSQL = strSQL & "" & mobjChargeInfor.����ID & ","
                    '      No_In       ҽ��������ϸ.No%Type,
                    strSQL = strSQL & "'" & strNo & "',"
                    '      ���㷽ʽ_In Varchar2,
                    strSQL = strSQL & "'" & strAdvance & "')"
                    '      ��ע_In     ҽ��������ϸ.��ע%Type := Null
                    zlDatabase.ExecuteProcedure strSQL, Me.Caption
                End If
                gcnOracle.CommitTrans: blnCommit = True
                
                Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, True, mobjChargeInfor.intInsure)
                Call SetBalanceVal(colBalance, p, str���㷽ʽ)
                
                gcnOracle.BeginTrans
            End If
            
            p = p + 1
            rsCharge.MoveNext
        Loop
        
        'ȫ���ɹ��������ܵĽ��㷽ʽ
        strAdvance = GetMedicareStr(colBalance)
    End If
    
    '���ݷ��صĽ�����Ϣ������Ԥ����¼��strAdvance���ظ�ʽ:���㷽ʽ1|���||���㷽ʽ2|���...
    If InStr(strAdvance, "|") > 0 Then
        Call ҽ�����ݸ���(mobjChargeInfor.����ID, mobjChargeInfor.����ID, strAdvance, True, Nothing)
    End If
    If Not (mInsurePara.�൥�ݷֵ��ݽ��� Or mInsurePara.һ�ν���ֵ����˷�) Then
        Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, True, mobjChargeInfor.intInsure)
    End If
    ExcuteInsureDel = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    If blnTransMedicare Then Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, False, mobjChargeInfor.intInsure)
    Call ErrCenter
End Function

Private Sub cmdDel_Click()
    Dim dblMoney As Double, strSQL As String
    Dim byt�������� As Byte
    Dim str���㷽ʽ As String
    
    ClearԤ����
    If mbytFunc = EM_FUN_���� Then Exit Sub
    'ɾ����صķ���
    With vsBlance
        If .Row < 0 Then Exit Sub
        '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
        Select Case Val(.RowData(.Row))
        Case 1  'Ԥ���
            byt�������� = 1
            str���㷽ʽ = ""
        Case 0  '��ͨ�Ľ��㷽ʽ
            byt�������� = 0
            str���㷽ʽ = .TextMatrix(.Row, .ColIndex("֧����ʽ"))
        Case Else
            ' 2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
             '����ֱ��ɾ��
            Exit Sub
        End Select
        dblMoney = Val(.TextMatrix(.Row, .ColIndex("֧�����")))
        
        mCurCarge.dbl��ǰδ�� = RoundEx(mCurCarge.dbl��ǰδ�� + dblMoney, 6)
        mCurCarge.dbl�����Ѹ��ϼ� = RoundEx(mCurCarge.dbl�����Ѹ��ϼ� - dblMoney, 6)
        Call SetControlProperty
        If Val(.RowData(.Row)) = 1 Then
            txt��Ԥ��.Enabled = True: txt��Ԥ��.BackColor = txt�ɿ�.BackColor
            If txt��Ԥ��.Enabled And txt��Ԥ��.Visible Then txt��Ԥ��.SetFocus
            zlControl.TxtSelAll txt��Ԥ��
            txt��Ԥ��.Tag = "": lbl��Ԥ��.Tag = ""
        End If
        If .Rows <= 2 Then
            .Clear 1
            .RowData(1) = ""
            .Cell(flexcpData, 1, 0, 1, .COLS - 1) = ""
        Else
            vsBlance.RemoveItem .Row
        End If
    End With
    Call SetDeleteVisible
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdExit_Click()
    ClearԤ����
    mblnOK = False: mbytReturnMode = EM_EX_�˳�
    Call ExcuteMainReshData(EM_EX_�˳�)
    Unload Me
End Sub

Private Sub cmdNext_Click()
    Dim blnUnload As Boolean
    '������һ�ŵ��ݵ�¼��
    '�����ϴ�֧����ʽ
    ClearԤ����
    If mCurCarge.dbl���γ�Ԥ�� <> 0 Then
        MsgBox "ʹ����Ԥ�����,���������շ�!", vbOKOnly + vbInformation, gstrSysName
        Exit Sub
    End If
    gtyPrePatiPay = mCurCardPay: mblnCur���� = True
    If Not Check�ɿ�(2) Then GoTo GoOver
    '�ȴ���Ԥ��
    If BrushcardStrikePrepay = False Then GoTo GoOver
    '�ٴ�������
    If isValied(True) = False Then GoTo GoOver
    If SaveCharge(, blnUnload) = False Then
        If mblnPriceBillCommit And mblnCommitBill = False Then
            'ֱ���շ�ʱ,ɾ��ǰһ�������ύ�Ļ��۵�
            Call DelMedicareTempNOs
            mblnPriceBillCommit = False
        End If
        GoTo GoOver
    End If
    
    mbln�������� = True
    'ˢ��������
    ExcuteMainReshData EM_EX_����
    mbytReturnMode = EM_EX_����
    Unload Me
GoOver:
    mobjChargeInfor.�շѽ��� = ""
    mblnCur���� = False
End Sub

Private Sub cmdOK_Click()
    Dim blnUnload As Boolean
    
    If mbytFunc = EM_FUN_���� Or mbytFunc = EM_FUN_���� Then
        '�������
        If zlIsCheckExistErrBill(mobjChargeInfor.�������) = False Then
            MsgBox "��ǰ�쳣�����ѱ������㲻�ܼ�����", vbInformation, gstrSysName
            Exit Sub
        End If
        If zlCheckOtherSessionDoing(mobjChargeInfor.�������) Then
            MsgBox "��ǰ�������������շѴ����н��д����㲻�ܼ�����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If zlIsCheckExiseSingularity(mobjChargeInfor.�������) Then
            If mbytFunc = EM_FUN_���� Then
                MsgBox "���쳣�����Ѿ������ϣ���ˣ�������" & IIf(mbytFunc = EM_FUN_����, "�����շ�", "��������") & "����ˢ�·����б�", vbInformation, gstrSysName
                Call cmdExit_Click: Exit Sub
            End If
        End If
        If Not zlIsCheckExistErrBill(mobjChargeInfor.�������) Then
            MsgBox "���쳣�����Ѿ��������շѣ���ˣ�������" & IIf(mbytFunc = EM_FUN_����, "�����շ�", "��������") & "����ˢ�·����б�", vbInformation, gstrSysName
            Call cmdExit_Click: Exit Sub
        End If
    End If
    If mbytFunc = EM_FUN_���� Then
        mblnOK = CancelBalance(blnUnload)
        If blnUnload = True Then
            ExcuteMainReshData (EM_EX_����)
            Unload Me
        End If
        Exit Sub
    End If
   '���ݽ��水�˻س���
   If mblnCacheKeyReturn Then mblnCacheKeyReturn = False: Exit Sub
    '�ȴ���Ԥ��
    mbln�������� = False
    mblnCurBrushPrepay = False
    If BrushcardStrikePrepay = False Then
       If txt��Ԥ��.Enabled And txt��Ԥ��.Visible Then txt��Ԥ��.SetFocus
        Exit Sub
    End If
    If mblnCurBrushPrepay Then
        If mblnUnloaded Then
            'ˢ����������Ϣ
            ExcuteMainReshData EM_EX_���
            Unload Me
        End If
        Exit Sub
    End If
    
    '�ٴ�������
    If isValied = False Then Exit Sub
    If txt�ɿ�.Text <> "0.00" Then
        'LED��ʾ
        Call ShowLedInfor
    End If
    If SaveCharge(, blnUnload) = False Then
        If mblnPriceBillCommit And mblnCommitBill = False Then
            'ֱ���շ�ʱ,ɾ��ǰһ�������ύ�Ļ��۵�
            Call DelMedicareTempNOs
            mblnPriceBillCommit = False
        End If
        Exit Sub
    End If
    If blnUnload Then
        'ˢ����������Ϣ
        ExcuteMainReshData EM_EX_���
        Unload Me
    End If
End Sub

Private Sub ExcuteMainReshData(ByVal bytExitMode As gExitMode)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ���������ˢ������
    '����:���˺�
    '����:2014-06-17 15:09:44
    '˵��:��Ҫ��Ӧ��ҽ��ˢ��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mblnOK Then zlAutoPayDrugAndStuff mcllOverPro  '�Զ�����
    If Not gfrmMain Is Nothing Then Exit Sub
    Call mfrmMain.zlExeBalanceWinRefrshData(mblnOK, bytExitMode, mbln�����շ�, mobjChargeInfor)
End Sub

Private Function zlAutoPayDrugAndStuff(ByRef cllDrugAndStuff As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����Զ�����
    '����:���ϳɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2012-02-06 14:55:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllItem As Collection, blnTrans As Boolean
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strWhere As String
    Dim strDrugSql As String, strStuffSql As String
    
    On Error GoTo errHandle
    If mbytFunc = EM_FUN_���� And (gbln�շѺ��Զ���ҩ Or gbln�����Զ�����) Then '104017
        '�쳣����ʱ�����Զ���ҩ/����
        Set cllDrugAndStuff = New Collection
        If gbln�շѺ��Զ���ҩ Then
            strDrugSql = _
                " Select Distinct 1 As ����, a.No, a.ִ�в���id, a.������" & vbNewLine & _
                " From ������ü�¼ A" & vbNewLine & _
                " Where a.��¼���� = 1 And Nvl(a.ִ�в���id, 0) <> 0" & vbNewLine & _
                "       And a.�շ���� In ('5', '6', '7') And a.����ID = [1]"
        End If
        
        If gbln�����Զ����� Then
            strStuffSql = _
                " Select Distinct 2 As ����, a.No, a.ִ�в���id, a.������" & vbNewLine & _
                " From ������ü�¼ A, �������� B" & vbNewLine & _
                " Where a.��¼���� = 1 And a.�շ�ϸĿid = b.����id(+) And Nvl(a.ִ�в���id, 0) <> 0" & vbNewLine & _
                "       And a.�շ���� = '4' And b.�������� = 1 And a.����ID = [1]"
        End If
        strSQL = strDrugSql & vbNewLine & _
            IIf(gbln�����Զ�����, " Union All" & vbNewLine & strStuffSql, "")
        
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯҩƷ�����ķ���", mobjChargeInfor.����ID)
        Do While Not rsTemp.EOF
            If Val(Nvl(rsTemp!����)) = 1 Then 'ҩƷ
                strSQL = "ZL_ҩƷ�շ���¼_������ҩ(" & Val(Nvl(rsTemp!ִ�в���ID)) & ",8,'" & Nvl(rsTemp!NO) & "','" & UserInfo.���� & "','" & UserInfo.���� & "','" & Nvl(rsTemp!������) & "')"
            Else '����
                '24-�շѴ������ϣ�25-���ʵ���������
                strSQL = "zl_�����շ���¼_��������(" & Val(Nvl(rsTemp!ִ�в���ID)) & ",24,'" & Nvl(rsTemp!NO) & "','" & UserInfo.���� & "','" & UserInfo.���� & "','" & UserInfo.���� & "',1,Sysdate)"
            End If
            zlAddArray cllDrugAndStuff, strSQL
            rsTemp.MoveNext
        Loop
        
        If cllDrugAndStuff.Count = 0 Then zlAutoPayDrugAndStuff = True: Exit Function
        blnTrans = True
        ExecuteProcedureArrAy cllDrugAndStuff, Me.Caption
        blnTrans = False
        zlAutoPayDrugAndStuff = True
        Exit Function
    End If
    
    If cllDrugAndStuff Is Nothing Then zlAutoPayDrugAndStuff = True: Exit Function
    
    blnTrans = True
    gcnOracle.BeginTrans
    For Each cllItem In cllDrugAndStuff '91665
        ExecuteProcedureArrAy cllItem, Me.Caption, True, True
    Next
    gcnOracle.CommitTrans
    blnTrans = False
    zlAutoPayDrugAndStuff = True
    Exit Function
errHandle:
    If blnTrans Then
        gcnOracle.RollbackTrans
        Call ErrCenter '��������
    Else
        If ErrCenter() = 1 Then
            Resume
        End If
    End If
End Function

Private Sub SetCtrlVisible()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ÿؼ�����ʾ״̬
    '����:���˺�
    '����:2012-02-03 13:58:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnTemp As Boolean
    If mbytFunc = EM_FUN_�շ� Then
        'ҽ����ҽ��δ���н���ʱ,����ʾ
        cmdYBBalance.Visible = mobjChargeInfor.intInsure <> 0 And Not mblnYbBalanced
        'ҽ�����н����˵�,���ҽ����,��ʾ����շ�
        cmdOK.Visible = (mobjChargeInfor.intInsure = 0 Or mobjChargeInfor.intInsure <> 0 And mblnYbBalanced)
        'ҽ�������˽����,�����˳�
        cmdExit.Visible = mobjChargeInfor.intInsure = 0 And Not (mblnThreeInterface Or mblnCommitBill) _
                          Or mobjChargeInfor.intInsure <> 0 And Not mblnYbBalanced
        '�����շ�
        blnTemp = gTy_Module_Para.byt�ɿ���� = 1 Or gTy_Module_Para.byt�ɿ���� = 3 '�Ƿ���������շ�
        '��ͨ�շѻ�ҽ���Ѿ�����
        blnTemp = blnTemp And (mobjChargeInfor.intInsure = 0 Or mobjChargeInfor.intInsure <> 0 And mblnYbBalanced)
        blnTemp = blnTemp And Val(txt��Ԥ��.Text) = 0 'δ��Ԥ�����
        cmdNext.Visible = blnTemp And (mCurCarge.dbl����ʵ�� = mCurCarge.dbl��ǰδ��)
        If (gTy_Module_Para.byt�ɿ���� = 1 Or gTy_Module_Para.byt�ɿ���� = 3) And mbln�������� Then
            cbo֧����ʽ.Locked = True
        End If
        Exit Sub
     End If
     
     If mbytFunc = EM_FUN_���� Then
        cmdExit.Caption = "�˳�(&E)"
        cmdOK.Visible = mblnYbBalanced Or mobjChargeInfor.intInsure = 0
        cmdYBBalance.Visible = Not mblnYbBalanced And mobjChargeInfor.intInsure <> 0
        cmdExit.Visible = True
        cmdNext.Visible = False
     End If
     If mbytFunc = EM_FUN_���� Then
        cmdOK.Caption = "���Ͻ���(&O)"
        cmdExit.Caption = "�˳�(&E)"
        cmdOK.Visible = True
        cmdYBBalance.Visible = False
        cmdExit.Visible = True
        cmdNext.Visible = False
     End If
End Sub

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
        '��Ҫ������������
        'Zl_�����շѽ���_Modify
        strSQL = "Zl_�����շѽ���_Modify("
        '  ��������_In   Number,
        strSQL = strSQL & "" & 2 & ","
        '  ����id_In     ������ü�¼.����id%Type,
        strSQL = strSQL & "" & lng����ID & ","
        '  ����id_In     ����Ԥ����¼.����id%Type,
        strSQL = strSQL & "" & lng����ID & ","
        '  ���㷽ʽ_In   Varchar2,
        strSQL = strSQL & "'" & strҽ������ & "')"
        '  ��Ԥ��_In     ����Ԥ����¼.��Ԥ��%Type,
        '  ��֧Ʊ��_In   ����Ԥ����¼.��Ԥ��%Type,
        '  �����id_In   ����Ԥ����¼.�����id%Type := Null,
        '  ����_In       ����Ԥ����¼.����%Type := Null,
        '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
        '  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
        '  �ɿ�_In       ����Ԥ����¼.�ɿ�%Type := Null,
        '  �Ҳ�_In       ����Ԥ����¼.�Ҳ�%Type := Null,
        '  �����_In   ������ü�¼.ʵ�ս��%Type := Null,
        '  ��ɽ���_In Number:=0
        ') As
        '  ------------------------------------------------------------------------------------------------------------------------------
        '  --����:�շѽ���ʱ,�޸Ľ���������Ϣ
        '  --��������_In:
        '  --   0-��ͨ�շѷ�ʽ:
        '  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ||.." ;Ҳ�������.
        '  --     �ڳ�Ԥ��_In:����漰Ԥ����,���뱾�εĳ�Ԥ��,�������շ�ʱ,������
        '  --     ����֧Ʊ��_In:����漰��֧Ʊ,���뱾�ε���֧Ʊ��,�������շ�ʱ,������
        '  --   1.����������:
        '  --     �ٽ��㷽ʽ_IN:ֻ�ܴ���һ�����㷽ʽ,���������һЩ������Ϣ,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ"
        '  --     �ڳ�Ԥ��_In: ������
        '  --     ����֧Ʊ��_In:������
        '  --     �ܿ����ID_IN,����_IN,������ˮ��_IN,����˵��_In:��Ҫ����
        '  --   2-ҽ������(�������ҽ���Ľ���,��Ҫ��ɾ��ԭҽ������,���´���ĸ���)
        '  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:���㷽ʽ|������||.."
        '  --     �ڳ�Ԥ��_In: ������
        '  --     ����֧Ʊ��_In:������
        '  --   3-���ѿ�����:
        '  --     �ٽ��㷽ʽ_IN:����һ��ˢ���ſ�,��ʽΪ:�����ID|����|���ѿ�ID|���ѽ��||."
        '  --     �ڳ�Ԥ��_In: ������
        '  --     ����֧Ʊ��_In:������
        '  -- �����_In:��������ʱ,����
        '  -- ��ɽ���_In:1-����շ�;0-δ����շ�
        '  ------------------------------------------------------------------------------------------------------------------------------
    End If
    If cllPro Is Nothing Then
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Else
        zlAddArray cllPro, strSQL
    End If
    
    ҽ�����ݸ��� = True
End Function

Private Function ҽ������϶�() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Ƿ����ҽ��У��(��Ҫ��ҽ�����׵��óɹ���,����ҽ�����ݵĽ϶�)
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2012-02-07 16:21:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strҽ������ As String, i As Long, strShowMsg As String
     Dim strSQL As String, lng����ID As Long
    Dim cllBalance As Collection
    
    On Error GoTo errHandle
    lng����ID = IIf(mbln�쳣����, mobjChargeInfor.����ID, mobjChargeInfor.����ID)
    If lng����ID = 0 Then ҽ������϶� = True: Exit Function
    If mobjChargeInfor.intInsure = 0 Then ҽ������϶� = True: Exit Function
    
    '108630,���ٸ���"���ս����¼.У��"���жϣ�ֻҪ���쳣���ݶ�ҪУ��
'    '0-����;1-��У��;2-���У��;3-���ӣ�ָ���ص�������ҽ��֧���ĸ��ֽ��㷽ʽ
'    gstrSQL = "" & _
'    "   Select /*+ rule */ A.��¼ID,A.У��  " & _
'    "   From ���ս����¼ A" & _
'    "   Where A.��¼ID=[1] And nvl(A.У��,0)=1 "
'    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng����ID)
'    If rsTemp.EOF Then ҽ������϶� = True: Exit Function
    
    '��ͨ����ҽ��������ϸ������У��
    strSQL = "Select 1" & _
            " From ����Ԥ����¼ A, ���㷽ʽ B" & _
            " Where a.���㷽ʽ = b.���� And b.���� In (3, 4) And ����id = [1] " & _
            "       And Nvl(У�Ա�־, 0) = 1 And Rownum < 2"
    strSQL = strSQL & "Union All" & _
            " Select 1" & _
            " From ���ս����¼" & _
            " Where ��¼id = [1] " & _
            "       And Not Exists(Select 1 From ����Ԥ����¼ A, ���㷽ʽ B" & _
            "                       Where a.���㷽ʽ = b.���� And b.���� In (3, 4) And a.����id = ��¼id)" & _
            "       And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    If rsTemp.EOF Then ҽ������϶� = True: Exit Function
    
    strҽ������ = zlGetYBBalanceNo(lng����ID)
    
    '���ҽ���˶Ա��޼�¼���˳�
    'Select ����ID,���㷽ʽ,��� From ���ս�����ϸ Where ��־=1
    strSQL = "Select A.����ID,a.���㷽ʽ,a.���" & _
            " From ���ս�����ϸ A ,���㷽ʽ C" & _
            " Where A.����id =[1] and A.��־=1 and A.���㷽ʽ=C.���� And C.���� in (3,4) " & _
            " Order by A.���㷽ʽ"
    'ҽ���ܿصĹ��̶̹�д����һ��"�ֽ�",�����ſ���ҽ����Ľ��㷽ʽ
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���ս������", lng����ID)
    'δ�к˶�����,ֱ�ӷ���
    If rsTemp.RecordCount = 0 And strҽ������ = "" Then ҽ������϶� = True: Exit Function
    
    If rsTemp.RecordCount > 0 Then
        strҽ������ = "" '���㷽ʽ|������||
        Set cllBalance = New Collection
        For i = 1 To rsTemp.RecordCount
            strҽ������ = strҽ������ & "||" & Nvl(rsTemp!���㷽ʽ) & "|" & Val(Nvl(rsTemp!���))
            rsTemp.MoveNext
        Next
        If strҽ������ <> "" Then strҽ������ = Mid(strҽ������, 3)
    End If
    If strҽ������ = "" Then ҽ������϶� = True: Exit Function
    
    strShowMsg = Replace(Replace(strҽ������, "||", vbCrLf), "|", "��")
    MsgBox "ע�⣺" & vbCrLf & "    ҽ��" & IIf(mbln�쳣����, "�˷�", "") & _
        "���������ѱ���������ݿ��ܲ�һ�£���У�Խ������ݡ�" & vbCrLf & _
        "����Ϊ��ȷ��" & IIf(mbln�쳣����, "�˷�", "") & "�������ݣ�" & vbCrLf & _
        strShowMsg, vbInformation + vbOKOnly, gstrSysName
    Call ҽ�����ݸ���(mobjChargeInfor.����ID, lng����ID, strҽ������, mbln�쳣����, Nothing)
    
    '�޸�У�Ա�־,ҽ���϶�����ɹ�
    If mbln�쳣���� = False And mInsurePara.�൥�ݷֵ��ݽ��� And gTy_Module_Para.blnֻ��ҽ������ɹ������շ� Then
        'ͨ��"ҽ��������ϸ"����Ƿ��ǡ�ֻ��ҽ������ɹ������շѡ����쳣����
        strSQL = "Select 1" & vbNewLine & _
                " From ������ü�¼ A, ҽ��������ϸ B" & vbNewLine & _
                " Where a.����id = b.����id(+) And a.No = b.No(+) And a.����id = [1] And b.No Is Null And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
        mbln�ֵ��ݽ������ȫ�� = Not rsTemp.EOF
        If mbln�ֵ��ݽ������ȫ�� = False Then
            '�޸�У�Ա�־
            ' Zl_���������շ�_ҽ������
            strSQL = "Zl_���������շ�_ҽ������("
            '  ����id_In   ������ü�¼.����id%Type,
            strSQL = strSQL & lng����ID & ","
            '  �������_In ����Ԥ����¼.�������%Type,
            strSQL = strSQL & "Null,"
            '  ���ս���_In Varchar2
            strSQL = strSQL & "Null)"
            zlDatabase.ExecuteProcedure strSQL, Me.Caption
        End If
    End If
    '���¼�������
    Call LoadData
    Call LoadPatiInfor
    Call SetControlProperty
    
    ҽ������϶� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Exit Function
ErrCommit:
    Call ErrCenter
    Resume  '����ִ�����������ִ��
End Function

Private Sub cmdYBBalance_Click()
    Dim blnCancel As Boolean, strNos As String
    Dim lng����ID As Long, blnCommit As Boolean
    
    If mbytFunc = EM_FUN_���� Or mbytFunc = EM_FUN_���� Then
        '�������
        If zlIsCheckExistErrBill(mobjChargeInfor.�������) = False Then
            MsgBox "��ǰ�쳣�����ѱ������㲻�ܼ�����", vbInformation, gstrSysName
            Exit Sub
        End If
        If zlCheckOtherSessionDoing(mobjChargeInfor.�������) Then
            MsgBox "��ǰ�������������շѴ����н��д����㲻�ܼ�����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If zlIsCheckExiseSingularity(mobjChargeInfor.�������) Then
            If mbytFunc = EM_FUN_���� Then
                MsgBox "���쳣�����Ѿ������ϣ���ˣ�����" & IIf(mbytFunc = EM_FUN_����, "�����շ�", "��������") & "����ˢ�·����б�", vbInformation, gstrSysName
                Call cmdExit_Click: Exit Sub
            End If
        End If
        If Not zlIsCheckExistErrBill(mobjChargeInfor.�������) Then
            MsgBox "���쳣�����Ѿ��������շѣ���ˣ�����" & IIf(mbytFunc = EM_FUN_����, "�����շ�", "��������") & "����ˢ�·����б�", vbInformation, gstrSysName
            Call cmdExit_Click: Exit Sub
        End If
    End If
    If mInsurePara.�൥�ݷֵ��ݽ��� And gTy_Module_Para.blnֻ��ҽ������ɹ������շ� And mbln�ֵ��ݽ������ȫ�� = False Then
        If mfrmMain.zlSaveBillAndClinicSwapByNo(lng����ID, strNos, mcllOverPro, mobjChargeInfor, blnCommit) = False Then
            If blnCommit = False Then Exit Sub
            mobjChargeInfor.����ID = lng����ID
            mobjChargeInfor.������� = CStr(-1 * lng����ID)
            mobjChargeInfor.Nos = strNos
            
            '���ύ�ĵ�����ֱ��У�Խ�����Ϣ����
            Call ҽ������϶�
            '���¼�������
            Call LoadData
            Call LoadPatiInfor
            Call SetControlProperty
        Else
            mobjChargeInfor.����ID = lng����ID
            mobjChargeInfor.������� = CStr(-1 * lng����ID)
            mobjChargeInfor.Nos = strNos
        End If

        mblnSaveBill = True
        mblnYbBalanced = True: mblnCommitBill = True
        cmdExit.Visible = False
    Else
        '���ݱ���
        If SaveFeeBilL = False Then
            If mblnPriceBillCommit And mblnCommitBill = False Then
                'ֱ���շ�ʱ,ɾ��ǰһ�������ύ�Ļ��۵�
                Call DelMedicareTempNOs
                mblnPriceBillCommit = False
            End If
            Exit Sub
        End If
        '����ҽ������
        If zlInsureClinicSwap = False Then
            If mblnPriceBillCommit And mblnCommitBill = False Then
                'ֱ���շ�ʱ,ɾ��ǰһ�������ύ�Ļ��۵�
                Call DelMedicareTempNOs
                mblnPriceBillCommit = False
            End If
            Exit Sub
        End If
        If mblnElsePersonErrBill Then
            If UpdateElsePersonErrBill() = False Then Exit Sub
        End If
        mblnElsePersonErrBill = False '�Ѹ���
    End If

    Call LoadData
    'ҽ��:58344
    mblnYB�˿� = mCurCarge.dbl��ǰδ�� + mCurCarge.dblӦ���ۼ� < 0
    
    Call LoadPatiInfor
    Call SetControlProperty
    '���ҽ������,��Ҫ�������ð�ť
    Call SetCtrlVisible
    Call SetControlEnabled
    '��궨λ
    '����ʹ��Ԥ��
    If txt��Ԥ��.Visible And txt��Ԥ��.Enabled And gblnPrePayPriority Then
        txt��Ԥ��.SetFocus
        Call SetControlProperty(True)
        Call Show�����(True)
    Else
        mblnNotChange = True
        txt��Ԥ��.Text = ""
        mblnNotChange = False
        '70430,Ƚ����,2014-4-24,�ڽ���Ԥ����ʱ��ʾ�ɿ������ҽ������ʱ�ٴ���ʾ��ͬ�ɿ������ظ���ʾ��
        If txt�ɿ�.Enabled And txt�ɿ�.Visible Then
            mblnҽ���ѱ��� = True '�������ѱ���Ϊtrue,����txt�ɿ��ý��������
            txt�ɿ�.SetFocus
        End If
        Call Show�����(False)
    End If
    Call LedDisplayBank
    If mCurCarge.dbl��ǰδ�� = 0 And cmdOK.Visible And cmdOK.Enabled Then
        'ҽ��ȫ������,ֱ��ȷ�����:63773
        Call cmdOK_Click
    End If
End Sub

Private Sub Form_Activate()
    Dim blnSetFocus As Boolean
    
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
    Call cbo֧����ʽ_Click
    Call SetControlProperty
    Call SetCtrlVisible
    Call SetControlEnabled
    
    If txt��Ԥ��.Visible Then txt��Ԥ��.Enabled = True
    '��궨λ
    If gTy_Module_Para.blnҽ��������ȱʡ��λ Then
        If cmdYBBalance.Visible And cmdYBBalance.Enabled Then
            cmdYBBalance.SetFocus: blnSetFocus = True
        End If
    End If
    If blnSetFocus = False Then
        If Val(txt��Ԥ��.Text) <> 0 And txt��Ԥ��.Enabled Then
            If txt��Ԥ��.Enabled And txt��Ԥ��.Visible Then txt��Ԥ��.SetFocus
            Call Show�����(True)
        Else
            If txt�ɿ�.Enabled And txt�ɿ�.Visible Then txt�ɿ�.SetFocus
            Call Show�����(False)
        End If
    End If
    mblnLoad = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Select Case KeyCode
    Case vbKeyAdd, vbKeyF4
        If (gTy_Module_Para.byt�ɿ���� = 1 Or gTy_Module_Para.byt�ɿ���� = 3) And mbln�������� Then Exit Sub
        '47457
        If gTy_Module_Para.blnʹ�üӼ��л� = False And KeyCode = vbKeyAdd Then Exit Sub
        If Me.ActiveControl Is txt�ɿ� Then
            i = cbo֧����ʽ.ListIndex
            If i >= cbo֧����ʽ.ListCount - 1 Then
                i = 0
            Else
                i = i + 1
            End If
            cbo֧����ʽ.ListIndex = i
        End If
    Case vbKeySubtract
        If (gTy_Module_Para.byt�ɿ���� = 1 Or gTy_Module_Para.byt�ɿ���� = 3) And mbln�������� Then Exit Sub
        '47457
        If gTy_Module_Para.blnʹ�üӼ��л� = False And KeyCode = vbKeySubtract Then Exit Sub
        If Me.ActiveControl Is txt�ɿ� Then
            i = cbo֧����ʽ.ListIndex
            If i <= 0 Then
                i = cbo֧����ʽ.ListCount - 1
            Else
                i = i - 1
            End If
            cbo֧����ʽ.ListIndex = i
        End If
     Case vbKeyF12
            If Shift = vbCtrlMask Then
                'ǿ����LED����,(�ϼ�)
                 Call LedVoiceSpeak
            End If
    Case vbKeyF2
        'ǿ�����
        If mobjChargeInfor.intInsure <> 0 And mblnYbBalanced = False Then
            Call cmdYBBalance_Click
        Else
            cmdOK_Click '43169
        End If
    Case vbKeyReturn
      '      zlCommFun.PressKey vbKeyTab
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr(":'|", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub Form_Load()
    'ѡ������������Ƿ����˻س�����
    mblnCacheKeyReturn = False
    mblnCacheKeyReturn = (GetAsyncKeyState(VK_RETURN) And &H1) <> 0
    mstrTittle = "�����շѽ���"
    RestoreWinState Me, App.ProductName, mstrTittle
    Call zlInitTotalStru
    Call SetWindowsSize
    Set mrsOneCard = GetOneCard
    zlControl.CboSetWidth cbo֧����ʽ.hWnd, cbo֧����ʽ.Width * 2
    txt��Ԥ��.Enabled = False
    mblnFirst = True: mblnLoad = True
    mblnUnLoad = False
    zlControl.PicShowFlat picTotal, -1, , taCenterAlign
    zlControl.PicShowFlat Picture1, -1, , taCenterAlign
    zlControl.PicShowFlat picPay, -1, , taCenterAlign
    
    Call InitFace
    If ҽ������϶� = False Then Unload Me: Exit Sub
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    'If Me.Width < 10530 Then Me.Width = 10530
    'If Me.Height < 7035 Then Me.Height = 7035
    With picBlance
        .Width = ScaleWidth - .Left * 2
        .Height = ScaleHeight - stbThis.Height - .Top
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    'ֱ���շ�ʱ,������۵����ύ������û���շѣ���Ҫɾ��ǰһ�������ύ�Ļ��۵�
    If mblnPriceBillCommit And mblnCommitBill = False Then
        Call DelMedicareTempNOs
        mblnPriceBillCommit = False
    End If
    With mCurCarge
           .dbl���γ�Ԥ�� = 0
           .dbl����ʵ�� = 0
           .dbl����ҽ��֧�� = 0
           .dbl�����Ѹ��ϼ� = 0
           .dbl����Ӧ�� = 0
           .dbl��ǰδ�� = 0
           .dbl������� = 0
           .dbl����Ԥ�� = 0
           .dblԤ����� = 0
    End With
    mblnYB�˿� = False
    If Not mobjICCard Is Nothing Then
        Call mobjICCard.SetEnabled(False)
        Set mobjICCard = Nothing
    End If
    Set mrsClassMoney = Nothing
    With mCurCardPay
        .lng���ѿ�ID = 0
        .str������� = ""
        .dbl��ˢ��� = 0
    End With
    SaveWinState Me, App.ProductName, mstrTittle
End Sub

Private Sub picBlance_Resize()
    Err = 0: On Error Resume Next
    With vsBlance
'        fraSplitBottom.Left = 0
'        fraSplitBottom.Width = picBlance.ScaleWidth + 50
        .Left = picBlance.ScaleLeft
        .Width = picBlance.ScaleWidth
        .Height = picBlance.ScaleHeight - .Top
    End With
End Sub

Private Sub setDefaultPrepayMoney()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ȱʡԤ�����
    '����:���˺�
    '����:2011-08-13 17:21:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With mCurCarge
         txt��Ԥ��.Text = "0.00"
         If Not mblnLoad Or (mblnLoad And gblnPrePayPriority) Then
            If .dbl����Ԥ�� > 0 Then
                txt��Ԥ��.Text = Format(IIf(.dbl����Ԥ�� > .dbl��ǰδ��, .dbl��ǰδ��, .dbl����Ԥ��), "###0.00;###0.00;0.00;0.00")
                txt�ɿ�.Text = Format(.dbl��ǰδ�� - IIf(.dbl����Ԥ�� > .dbl��ǰδ��, .dbl��ǰδ��, .dbl����Ԥ��), "###0.00;###0.00;0.00;0.00")
            End If
        End If
    End With
End Sub

Private Sub LoadPatiInfor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ز�����Ϣ
    '����:���˺�
    '����:2011-08-13 10:52:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    stbThis.Panels(2).Text = mobjChargeInfor.����
    Set rsTemp = GetMoneyInfo(mobjChargeInfor.����ID, 0, False, 1, False, 0, True)
    Dim dbl������� As Double
    With mCurCarge
        .dblԤ����� = 0
        .dbl������� = 0
        Do While Not rsTemp.EOF
            .dblԤ����� = RoundEx(.dblԤ����� + Val(Nvl(rsTemp!Ԥ�����)), 6)
            .dbl������� = RoundEx(.dbl������� + Val(Nvl(rsTemp!�������)), 6)
            If Nvl(rsTemp!����, 0) = 1 Then
                dbl������� = RoundEx(Val(Nvl(rsTemp!Ԥ�����)) - Val(Nvl(rsTemp!�������)), 6)
            End If
            rsTemp.MoveNext
        Loop
        .dbl����Ԥ�� = RoundEx(.dblԤ����� - .dbl�������, 6)
    End With
    If RoundEx(mCurCarge.dbl����Ԥ��, 6) = 0 And RoundEx(dbl�������, 6) = 0 Then
        stbThis.Panels(3).Visible = False
    Else
        stbThis.Panels(3).Visible = True
        stbThis.Panels(3).Text = "Ԥ��:" & Format(mCurCarge.dbl����Ԥ��, "0.00") & _
            IIf(dbl������� > 0, "(������:" & Format(dbl�������, "0.00") & ")", "")
    End If
    
    txtҽ��.Text = Format(mCurCarge.dbl����ҽ��֧��, "###0.00;-###0.00;0.00;0.00;")
    txt�ϼ�.Text = Format(mCurCarge.dbl����ʵ��, "###0.00;-###0.00;0.00;0.00;")
    lbl�Ը��ϼ�.Caption = Format(mCurCarge.dbl����ʵ�� - mCurCarge.dbl����ҽ��֧��, "###0.00;-###0.00;0.00;0.00")
    Call setDefaultPrepayMoney
    If mCurCarge.dbl���γ�Ԥ�� <> 0 Then
        txt��Ԥ��.Text = Format(mCurCarge.dbl���γ�Ԥ��, "0.00")
        txt��Ԥ��.Tag = "1"
        txt��Ԥ��.BackColor = Me.BackColor
        lbl��Ԥ��.Tag = "1"
        txt��Ԥ��.Enabled = False
    End If
End Sub

Private Sub LedVoiceSpeak()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������
    '����:���˺�
    '����:2011-08-13 16:38:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    'If mCurCardPay.int���� <> 1 Then Exit Sub
    If gblnLED = False Then Exit Sub
    If mobjChargeInfor.intInsure <> 0 And mblnYbBalanced = False Then Exit Sub
    
'    If mCurCarge.dbl����ʵ�� = 0 Then Exit Sub
'    If mCurCarge.dbl��ǰδ�� = 0 Then Exit Sub
    zl9LedVoice.Speak "#21 " & Format(lblʣ���Ը�.Caption, "0.00")
    mbln�ѱ��� = True
End Sub

Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
   If Panel.Key = "Calc" Then
        mlngR = FindWindow("SciCalc", "������")
        If mlngR <> 0 Then
            BringWindowToTop mlngR
        Else
            On Error Resume Next
            Shell "calc.exe", vbNormalFocus
        End If
  End If
End Sub

Private Sub stcTittile_GotFocus()
  ClearԤ����
End Sub

Private Sub stcTittleTotal_GotFocus()
  ClearԤ����
End Sub

Private Sub txt��Ԥ��_Change()
    lbl��Ԥ��.Tag = "": txt��Ԥ��.Tag = ""
    txt��Ԥ��.BackColor = txt�ɿ�.BackColor
    If mblnNotChange Then Exit Sub
    Call SetControlProperty(True)
    Call Show�����(True)
End Sub

Private Sub txt��Ԥ��_GotFocus()
    If Val(txt��Ԥ��.Text) = 0 And mblnLoad = False Then
           Call setDefaultPrepayMoney
    End If
    zlControl.TxtSelAll txt��Ԥ��
    Call SetControlProperty(True)
    Call Show�����(True)
    'If Val(txt��Ԥ��.Tag) = Val(txt��Ԥ��.Text) Then Exit Sub
    
    '�Զ����ۻ��ֹ�����ʱ���ȼ�����
    'Call LedVoiceSpeak
End Sub

Private Sub txt��Ԥ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If Val(txt��Ԥ��.Tag) = Val(txt��Ԥ��.Text) Then GoTo SendKeyTab:
    If Val(txt��Ԥ��.Text) = 0 Then GoTo SendKeyTab:
    If CheckPrepayMoneyIsValied = False Then Exit Sub
    If mblnUnloaded Then
        'ˢ����������Ϣ
        ExcuteMainReshData EM_EX_���
        Unload Me: Exit Sub
    End If
    Exit Sub
SendKeyTab:
    zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txt��Ԥ��_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txt��Ԥ��, KeyAscii, m���ʽ
End Sub

Private Sub txt��Ԥ��_LostFocus()
    If mblnLoad Then Exit Sub
    If Val(txt��Ԥ��.Text) = 0 Then txt�ɿ�.Text = ""
End Sub

Private Sub txt��Ԥ��_Validate(Cancel As Boolean)
    If lbl��Ԥ��.Tag = "1" Then Exit Sub
    If mobjChargeInfor.����ID = 0 Then Exit Sub
    
    If Val(txt��Ԥ��.Tag) = Val(txt��Ԥ��.Text) Then Exit Sub
    
    If txt��Ԥ��.Text = "" Then
        txt��Ԥ��.Text = "0.00"
    ElseIf Not IsNumeric(txt��Ԥ��.Text) And txt��Ԥ��.Text <> "" Then
        MsgBox "��Ч��ֵ��", vbInformation, gstrSysName
        If txt��Ԥ��.Enabled And txt��Ԥ��.Visible Then txt��Ԥ��.SetFocus
        zlControl.TxtSelAll txt��Ԥ��: Cancel = True: Exit Sub
    ElseIf Val(txt��Ԥ��.Text) < 0 Then
        MsgBox "Ԥ��������Ϊ����", vbInformation, gstrSysName
        Call setDefaultPrepayMoney
        If txt��Ԥ��.Enabled And txt��Ԥ��.Visible Then txt��Ԥ��.SetFocus
        zlControl.TxtSelAll txt��Ԥ��: Cancel = True: Exit Sub
    ElseIf Val(txt��Ԥ��.Text) > 0 And mCurCarge.dbl����ʵ�� < 0 Then
        MsgBox "����Ӧ�����Ϊ��ʱ����ʹ��Ԥ��", vbInformation, gstrSysName
        txt��Ԥ��.Text = "0.00"
        If txt��Ԥ��.Enabled And txt��Ԥ��.Visible Then txt��Ԥ��.SetFocus
        zlControl.TxtSelAll txt��Ԥ��:   Exit Sub
    Else
        txt��Ԥ��.Text = Format(Val(txt��Ԥ��.Text), "0.00")
    End If
    If Val(txt��Ԥ��.Text) > mCurCarge.dbl����Ԥ�� Then
        MsgBox "Ԥ�������ܳ������˵�Ԥ�����:" & Format(mCurCarge.dbl����Ԥ��, "0.00") & " ��", vbInformation, gstrSysName
        Call setDefaultPrepayMoney
        If txt��Ԥ��.Enabled And txt��Ԥ��.Visible Then txt��Ԥ��.SetFocus
        zlControl.TxtSelAll txt��Ԥ��: Cancel = True: Exit Sub
    ElseIf Val(txt��Ԥ��.Text) > Format(mCurCarge.dbl��ǰδ��, "0.00") And Val(txt��Ԥ��.Text) <> 0 Then
        MsgBox "Ԥ��������ܴ���Ӧ�����:" & Format(mCurCarge.dbl��ǰδ��, "0.00") & " ��", vbInformation, gstrSysName
        Call setDefaultPrepayMoney
        If txt��Ԥ��.Enabled And txt��Ԥ��.Visible Then txt��Ԥ��.SetFocus
        zlControl.TxtSelAll txt��Ԥ��: Cancel = True: Exit Sub
    Else
        txt��Ԥ��.Text = Format(Val(txt��Ԥ��.Text), "0.00")
    End If
   ' If CheckPrepayMoneyIsValied = False Then Cancel = True: Exit Sub
End Sub

Private Function BrushcardStrikePrepay() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��֤ˢ����Ԥ��
    '����:�����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-14 14:35:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If Val(lbl��Ԥ��.Tag) = 1 Then BrushcardStrikePrepay = True: Exit Function
    If Val(txt��Ԥ��) = 0 Then BrushcardStrikePrepay = True: Exit Function
    mblnCurBrushPrepay = True

    If Not IsNumeric(txt��Ԥ��.Text) And txt��Ԥ��.Text <> "" Then
        MsgBox "��Ч��ֵ��", vbInformation, gstrSysName
        If txt��Ԥ��.Enabled And txt��Ԥ��.Visible Then txt��Ԥ��.SetFocus
        zlControl.TxtSelAll txt��Ԥ��: Exit Function
    ElseIf Val(txt��Ԥ��.Text) < 0 Then
        MsgBox "Ԥ��������Ϊ����", vbInformation, gstrSysName
        Call setDefaultPrepayMoney
        If txt��Ԥ��.Enabled And txt��Ԥ��.Visible Then txt��Ԥ��.SetFocus
        zlControl.TxtSelAll txt��Ԥ��: Exit Function
    ElseIf Val(txt��Ԥ��.Text) > 0 And mCurCarge.dbl����ʵ�� < 0 Then
        MsgBox "����Ӧ�����Ϊ��ʱ����ʹ��Ԥ��", vbInformation, gstrSysName
        txt��Ԥ��.Text = "0.00"
        If txt��Ԥ��.Enabled And txt��Ԥ��.Visible Then txt��Ԥ��.SetFocus
        zlControl.TxtSelAll txt��Ԥ��:   Exit Function
    Else
        txt��Ԥ��.Text = Format(Val(txt��Ԥ��.Text), "0.00")
    End If
    If Val(txt��Ԥ��.Text) > mCurCarge.dbl����Ԥ�� Then
        MsgBox "Ԥ�������ܳ������˵�Ԥ�����:" & Format(mCurCarge.dbl����Ԥ��, "0.00") & " ��", vbInformation, gstrSysName
        Call setDefaultPrepayMoney
        If txt��Ԥ��.Enabled And txt��Ԥ��.Visible Then txt��Ԥ��.SetFocus
        zlControl.TxtSelAll txt��Ԥ��: Exit Function
    ElseIf Val(txt��Ԥ��.Text) > Format(mCurCarge.dbl��ǰδ��, "0.00") And Val(txt��Ԥ��.Text) <> 0 Then
        MsgBox "Ԥ��������ܴ���Ӧ�����:" & Format(mCurCarge.dbl��ǰδ��, "0.00") & " ��", vbInformation, gstrSysName
        Call setDefaultPrepayMoney
        If txt��Ԥ��.Enabled And txt��Ԥ��.Visible Then txt��Ԥ��.SetFocus
        zlControl.TxtSelAll txt��Ԥ��: Exit Function
    End If
    
    'ˢ��ȷ��
    'frmParent As Object, ByVal lngSys As Long, _
    ByVal lng����ID As Long, ByVal cur��� As Currency, _
    Optional lngModul As Long = 0, _
    Optional bytOperationType As Byte = 0
    Dim str����IDs As String
    If zlDatabase.PatiIdentify(Me, glngSys, mobjChargeInfor.����ID, Val(txt��Ԥ��), mlngModule, 1, mlngBrushCardTypeID, _
            IIf(-1 * gdblԤ��������鿨 >= Val(txt��Ԥ��), False, True), True, str����IDs, (gdblԤ��������鿨 <> 0), (gdblԤ��������鿨 = 2)) Then
        mobjChargeInfor.����IDs = str����IDs
        lbl��Ԥ��.Tag = "1"
       ' txt��Ԥ��.ForeColor = d
       txt��Ԥ��.BackColor = Me.BackColor
       txt��Ԥ��.Tag = Val(txt��Ԥ��)
       txt��Ԥ��.Enabled = False
        If SaveCharge(True) = False Then
            If mblnPriceBillCommit And mblnCommitBill = False Then
                'ֱ���շ�ʱ,ɾ��ǰһ�������ύ�Ļ��۵�
                Call DelMedicareTempNOs
                mblnPriceBillCommit = False
            End If
            txt��Ԥ��.Enabled = True
            txt��Ԥ��.BackColor = txt�ɿ�.BackColor
             lbl��Ԥ��.Tag = ""
            Exit Function
        End If
         BrushcardStrikePrepay = True
        If mblnUnloaded Then Exit Function
    Else
        lbl��Ԥ��.Tag = ""
        txt��Ԥ��.Enabled = True
        Call SetControlProperty
       Exit Function
    End If
    Call SetControlProperty
    BrushcardStrikePrepay = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function BrushCardThreeSwapCheck() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ˢ����֤
    '����:����true,���򷵻�False
    '����:���˺�
    '����:2011-08-14 14:35:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsMoney As ADODB.Recordset, strXMLExpend As String
    Dim dblMoney  As Double, dblBrushCardMoneyed As Double '��ˢ���ѿ����
    Dim cllSquareBalance As Collection
    On Error GoTo errHandle
    If mCurCardPay.lngҽ�ƿ����ID = 0 Then BrushCardThreeSwapCheck = True: Exit Function
    If mblnYB�˿� Then
        MsgBox "��ǰΪ�˿�ģʽ��Ŀǰϵͳ�ݲ�֧�ֽ��˿���˸� " & mCurCardPay.str���㷽ʽ & "��", vbInformation + vbOKOnly, gstrSysName
        If txt�ɿ�.Enabled And txt�ɿ�.Visible Then txt�ɿ�.SetFocus
        zlControl.TxtSelAll txt�ɿ�: Exit Function
    End If
    If Val(txt�ɿ�) = 0 Then
        MsgBox "δ���뽻�׽����飡", vbInformation + vbOKOnly, gstrSysName
        If txt�ɿ�.Enabled And txt�ɿ�.Visible Then txt�ɿ�.SetFocus
        zlControl.TxtSelAll txt�ɿ�: Exit Function
    End If
    If Not IsNumeric(txt�ɿ�.Text) Then
        MsgBox "��Ч��ֵ��", vbInformation + vbOKOnly, gstrSysName
        If txt�ɿ�.Enabled And txt�ɿ�.Visible Then txt�ɿ�.SetFocus
        zlControl.TxtSelAll txt�ɿ�: Exit Function
    ElseIf Val(txt�ɿ�.Text) > Format(Abs(mCurCarge.dbl��ǰδ��), "0.00") Then
        MsgBox "���׽��ܴ��ڱ���δ�����:" & Format(mCurCarge.dbl��ǰδ��, "0.00") & " ��", vbInformation, gstrSysName
        If txt�ɿ�.Enabled And txt�ɿ�.Visible Then txt�ɿ�.SetFocus
        zlControl.TxtSelAll txt�ɿ�: Exit Function
    End If
    If mCurCardPay.lngҽ�ƿ����ID > 0 And Not mCurCardPay.bln���ѿ� Then
        If Val(txt�ɿ�.Text) <> Format(Abs(mCurCarge.dbl��ǰδ��), "0.00") Then
            If gTy_Module_Para.bytˢ��ȱʡ������ = 1 Then
                If MsgBox("���׽��(" & Format(Val(txt�ɿ�.Text), "0.00") & ")�뱾��δ�����(" & Format(mCurCarge.dbl��ǰδ��, "0.00") & _
                    ")��ͬ���Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            ElseIf gTy_Module_Para.bytˢ��ȱʡ������ = 2 Then
                MsgBox "���׽��(" & Format(Val(txt�ɿ�.Text), "0.00") & ")�뱾��δ�����(" & Format(mCurCarge.dbl��ǰδ��, "0.00") & _
                    ")��ͬ�����ܼ�����", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    If zlGetClassMoney(mobjChargeInfor.����ID, rsMoney) = False Then Exit Function
    '   zlBrushCard(frmMain As Object, _
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
    Optional ByVal str������Դ As String, _
    Optional ByVal lng����ID As Long) As Boolean
    '       strXmlIn-����������XML���,Ŀǰ��ʽ����:
    '       <IN>
    '           <CZLX>0</CZLX>    //��������,0-��������ˢ��,1-ת�˵���ˢ��,2-�˿����ˢ��
    '       </IN>
    '       varSquareBalance- Collection����,���ص�ǰˢ������(array(�����ID,���ѿ�ID,ˢ�����,����,����,�������,�Ƿ�����))
    '       str������Դ - ��ǰ֧�����õķ�����Դ�������ö��ŷָ�(ʹ�����ѿ�֧��ʱ����)
    '       lng����ID - ����ID(ʹ�����ѿ�֧��ʱ����)
    Set cllSquareBalance = Nothing
    Set mcllCurSquareBalance = Nothing
    If mCurCardPay.bln���ѿ� Then
        '�������ѿ���ˢ����Ϣ
       Set cllSquareBalance = mcllSquareBalance
     End If
     
    dblMoney = Val(txt�ɿ�.Text)
    If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModule, rsMoney, _
        mCurCardPay.lngҽ�ƿ����ID, mCurCardPay.bln���ѿ�, _
    mobjChargeInfor.����, mobjChargeInfor.�Ա�, mobjChargeInfor.����, dblMoney, mCurCardPay.strˢ������, mCurCardPay.strˢ������, _
    False, True, False, False, cllSquareBalance, False, False, "<IN><CZLX>0</CZLX></IN>", mobjChargeInfor.������Դ, mobjChargeInfor.����ID) = False Then Exit Function
    '���ѿ���ֵ
    If mCurCardPay.bln���ѿ� Then
        Set mcllCurSquareBalance = cllSquareBalance
    End If
    
    '����ǰ,һЩ���ݼ��
    'zlPaymentCheck(frmMain As Object, ByVal lngModule As Long, _
    ByVal strCardTypeID As Long, ByVal strCardNo As String, _
    ByVal dblMoney As Double, ByVal strNOs As String, _
    Optional ByVal strXMLExpend As String
    'mobjChargeInfor.strNOs:��������ʱ,û�����ʱ,����Ϊ��.
    If gobjSquare.objSquareCard.zlPaymentCheck(Me, mlngModule, mCurCardPay.lngҽ�ƿ����ID, _
        mCurCardPay.bln���ѿ�, mCurCardPay.strˢ������, dblMoney, mobjChargeInfor.Nos, strXMLExpend) = False Then Exit Function
'    zlGetAccountMoney(ByVal frmMain As Object, ByVal lngModule As Long, _
'    ByVal strCardTypeID As Long, _
'    ByVal strCardNo As String, strExpand As String, dblMoney As Double
    '���:frmMain-���õ�������
    '        lngModule-ģ���
    '        strCardNo-����
    '        strExpand-Ԥ����Ϊ��,�Ժ���չ
    '����:dblMoney-�����ʻ����
    Dim strExpand As String, dbl�ʻ���� As Double
    If gobjSquare.objSquareCard.zlGetAccountMoney(Me, mlngModule, mCurCardPay.lngҽ�ƿ����ID, _
          mCurCardPay.strˢ������, strExpand, dbl�ʻ����, mCurCardPay.bln���ѿ�) = False Then Exit Function
    stbThis.Panels(4).Text = Format(dbl�ʻ����, "0.00")
    stbThis.Panels(4).ToolTipText = mCurCardPay.str���㷽ʽ & "���ʻ����:" & Format(dbl�ʻ����, "0.00")
    mCurCardPay.dbl�ʻ���� = RoundEx(dbl�ʻ����, 2)
    '�Ѿ�������֧�����
    If dblMoney <> Val(txt�ɿ�.Text) Then
        txt�ɿ�.Text = Format(dblMoney, "0.00")
    End If
    BrushCardThreeSwapCheck = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function zlGetClassMoney(ByRef lng����ID As Long, ByRef rsMoney As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ʱ,��ʼ��֧�����(�շ����,ʵ�ս��)
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-06-10 17:52:18
    '����:38841
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    On Error GoTo errHandle
    If Not mrsClassMoney Is Nothing Then
        Set rsMoney = mrsClassMoney: zlGetClassMoney = True: Exit Function
    End If
    
    '��ʼ�����ݽṹ
    Set mrsClassMoney = New ADODB.Recordset
    mrsClassMoney.Fields.Append "�շ����", adVarChar, 10, adFldIsNullable
    mrsClassMoney.Fields.Append "���", adDouble, , adFldIsNullable
    mrsClassMoney.CursorLocation = adUseClient
    mrsClassMoney.LockType = adLockOptimistic
    mrsClassMoney.CursorType = adOpenStatic
    mrsClassMoney.Open
    If lng����ID = 0 And mbytFunc = EM_FUN_�շ� Then
        Call mfrmMain.zlGetClassMoney(rsTemp)
    Else
        strSQL = "" & _
        "   Select  A.�շ����,nvl(sum(ʵ�ս��) ,0) as ���   " & _
        "   From ������ü�¼ A,(Select ����ID From ����Ԥ����¼ where ����ID=[1] ) B " & _
        "   Where A.����ID=B.����ID " & _
        "   Group by �շ����"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    End If
    With rsTemp
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            mrsClassMoney.Find "�շ����='" & Nvl(!�շ����, "��") & "'", , adSearchForward, 1
            If mrsClassMoney.EOF Then mrsClassMoney.AddNew
            mrsClassMoney!�շ���� = Nvl(!�շ����, "��")
            mrsClassMoney!��� = Val(Nvl(mrsClassMoney!���)) + Val(Nvl(!���))
            mrsClassMoney.Update
            .MoveNext
        Loop
    End With
    Set rsMoney = mrsClassMoney
    zlGetClassMoney = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub txt�ϼ�_GotFocus()
  ClearԤ����
End Sub

Private Sub txt�ɿ�_Change()
    Call SetControlProperty
    Call Show�����(False)
End Sub

Private Sub txt�ɿ�_GotFocus()
    'ֻ�Խɿ���Ϊ�շѽ�������ʱ,��������ɿ��0
    '���˺�:22343
    Call ClearԤ����
    If gTy_Module_Para.byt�ɿ���� = 1 _
        Or gTy_Module_Para.byt�ɿ���� = 3 _
        Or gTy_Module_Para.byt�ɿ���� = 2 Then
        If Val(txt�ɿ�.Text) = 0 And Me.ActiveControl Is txt�ɿ� Then
            txt�ɿ�.Text = ""
        End If
    End If
    With mCurCardPay
       If .bln���ѿ� Or (.int���� <> 1 And mblnYB�˿�) Then
           '57682:ȱʡΪ����֧�����
           txt�ɿ�.Text = Format(IIf(mblnYB�˿�, -1, 1) * Val(lblʣ���Ը�.Caption), "0.00")
        ElseIf mCurCardPay.lngҽ�ƿ����ID > 0 And Not mCurCardPay.bln���ѿ� Then
            If gTy_Module_Para.bytˢ��ȱʡ������ <> 0 Then
                txt�ɿ�.Text = Format(IIf(mblnYB�˿�, -1, 1) * Val(lblʣ���Ը�.Caption), "0.00")
            End If
        End If
    End With
    Call SetControlProperty
    Call Show�����(False)
  '  Call zlControl.TxtSelAll(txt�ɿ�)
    '�Զ����ۻ��ֹ�����ʱ���ȼ�����
    If mblnҽ���ѱ��� Then
        mblnҽ���ѱ��� = False
    Else
        Call LedVoiceSpeak
    End If
    zlControl.TxtSelAll txt�ɿ�
End Sub

Private Sub ShowLedInfor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾLed��Ϣ
    '����:���˺�
    '����:2011-08-13 15:25:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gblnLED = False Then Exit Sub
'    If mCurCarge.dbl����ʵ�� = 0 Then Exit Sub
    
    'ֻ�н��ֲ���ʾ
    If Val(txt��Ԥ��.Text) = 0 And mCurCardPay.int���� = 1 Then
        zl9LedVoice.DispCharge mCurCarge.dbl��ǰδ�� - mCurCarge.dbl�������� + mCurCarge.dblӦ���ۼ�, Val(txt�ɿ�.Text), Val(txt�Ҳ�.Text)
    Else '����֧���ֽ�ʱ�Ĵ���
        Call zl9LedVoice.DisplayBank( _
            "�ϼ�:" & txt�ϼ�.Text & "Ԫ,Ӧ��:" & lblʣ���Ը�.Caption & "Ԫ", _
            "����:" & txt�ɿ�.Text & "Ԫ" & IIf(Val(txt�Ҳ�.Text) = 0, "", ",����:" & Val(txt�Ҳ�.Text) & "Ԫ"))
    End If
    zl9LedVoice.Speak "#22 " & Val(txt�ɿ�.Text)
    zl9LedVoice.Speak "#23 " & Val(txt�Ҳ�.Text)
    zl9LedVoice.Speak "#3"
End Sub

Private Sub LedDisplayBank()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ������Ϣ
    '����:���˺�
    '����:2011-12-15 13:40:46
    '����:52117
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl���ʺϼ� As Double, i As Long
    Dim strҽ�� As String, str�������� As String, str��һ��ͨ As String, str��ͨ���� As String
    Dim varPara  As Variant, str���㷽ʽ As String
    If Not gblnLED Then Exit Sub
    
    With vsBlance
        For i = 1 To .Rows - 1
            'ҽ������
            If .TextMatrix(i, .ColIndex("֧����ʽ")) <> "" Then
                Select Case .RowData(i)
                Case 1 'ҽ��
                    strҽ�� = strҽ�� & "||" & .TextMatrix(i, .ColIndex("֧����ʽ")) & ":" & Format(Val(.TextMatrix(i, .ColIndex("֧�����"))), "0.00")
                Case 2 '�����ӿڽ���
                    str�������� = str�������� & "||" & .TextMatrix(i, .ColIndex("֧����ʽ")) & ":" & Format(Val(.TextMatrix(i, .ColIndex("֧�����"))), "0.00")
                Case 3   ' һ��ͨ����
                    str��һ��ͨ = str��һ��ͨ & "||" & .TextMatrix(i, .ColIndex("֧����ʽ")) & ":" & Format(Val(.TextMatrix(i, .ColIndex("֧�����"))), "0.00")
                Case Else
                    str��ͨ���� = str��ͨ���� & "||" & .TextMatrix(i, .ColIndex("֧����ʽ")) & ":" & Format(Val(.TextMatrix(i, .ColIndex("֧�����"))), "0.00")
                End Select
            End If
        Next
    End With
     
    str���㷽ʽ = ""
    If strҽ�� <> "" Then str���㷽ʽ = str���㷽ʽ & "||ҽ������:||�ʻ����:" & Format(mcur�������, "0.00") & strҽ��
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

    '70430,Ƚ����,2014-4-24,�ڽ���Ԥ����ʱ��ʾ�ɿ������ҽ������ʱ�ٴ���ʾ��ͬ�ɿ������ظ���ʾ��
    If Format(mdblԭδ��, gstrDec) <> Format(Val(lblʣ���Ը�.Caption), gstrDec) Then
        zl9LedVoice.Speak "#21 " & Format(Val(lblʣ���Ը�.Caption), "0.00")
        mbln�ѱ��� = True
    End If
End Sub

Private Function Check�ɿ�(ByVal bytMode As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ɿ���
    '���:bytMode-0-�ڽɿ�س�����;1-���������;2-�����Ǽ�����������
    '����:
    '����:����ᷨ,����true,���򷵻�False
    '����:���˺�
    '����:2011-09-06 10:47:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If cbo֧����ʽ.ListIndex < 0 Then Exit Function
    If txt�ɿ�.Text <> "" Then
        txt�ɿ�.Text = Format(Val(txt�ɿ�.Text), "0.00")
        If Abs(Val(txt�ɿ�.Text)) > 99999999 Then
            MsgBox "����Ľɿ������,����ܳ���99999999!", vbOKOnly, gstrSysName
            Exit Function
        End If
        If Val(txt�ɿ�.Text) = 0 Then
            If cbo֧����ʽ.ItemData(cbo֧����ʽ.ListIndex) = -1 Then
                '��Ҫ�ų������ӿڽ���
                MsgBox "δ����ɿ���,������" & cbo֧����ʽ.Text & "֧��,����!", vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
            If mCurCardPay.blnOneCard Then
                '��Ҫ�ų������ӿڽ���
                MsgBox "δ����ɿ���,������һ��ͨ����֧��,����!", vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
        End If
        Check�ɿ� = True
        Exit Function
    End If
    
    'δ����ɿ�����
    '�ɿ����:0-�������нɿ�������ۼƿ���,1-��������ɿ��Ž��������ۼ�
    '       2-�շ�ʱ����Ҫ����ɿ���
    Select Case gTy_Module_Para.byt�ɿ����
    Case 1, 3 '1-�ಡ���ۼ�; 3-�������ۼƻ�
        If cbo֧����ʽ.ItemData(cbo֧����ʽ.ListIndex) = -1 Then
            '��Ҫ�ų������ӿڽ���
            MsgBox "δ����ɿ���,������" & cbo֧����ʽ.Text & "֧��,����!", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
        If mCurCardPay.blnOneCard Then
            '��Ҫ�ų������ӿڽ���
            MsgBox "δ����ɿ���,������һ��ͨ����֧��,����!", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
    Case 2  '�շ�ʱ����Ҫ����ɿ���
            MsgBox "δ����ɿ���,����", vbOKOnly + vbInformation, gstrSysName
            txt�ɿ�.SetFocus: Exit Function
    End Select
    
    Check�ɿ� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub txt�ɿ�_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txt�ɿ�, KeyAscii, m���ʽ
    If KeyAscii <> 13 Then Exit Sub
    If mblnCacheKeyReturn = True Then mblnCacheKeyReturn = False
    KeyAscii = 0
    If Check�ɿ�(0) = False Then Exit Sub
    
    If (gTy_Module_Para.byt�ɿ���� = 1 Or gTy_Module_Para.byt�ɿ���� = 3) And txt�ɿ�.Text = "" Then
        If cmdNext.Visible And cmdNext.Enabled Then cmdNext.SetFocus
          Exit Sub
    End If
    
    'ֻ�Խɿ���Ϊ�շѽ�������ʱ,��������ɿ��0
    If gTy_Module_Para.byt�ɿ���� = 1 _
        Or gTy_Module_Para.byt�ɿ���� = 3 _
        Or gTy_Module_Para.byt�ɿ���� = 2 Then
        If txt�ɿ�.Text = "" Then Exit Sub
    End If
    
    If mCurCardPay.int���� <> 1 Then
        If mCurCardPay.bln֧Ʊ Or (cbo֧����ʽ.Text Like "*��*" And mCurCardPay.lngҽ�ƿ����ID = 0) Then
            zlCommFun.PressKey vbKeyTab
        Else
            Call cmdOK_Click
            Call txt�ɿ�_GotFocus   '47147
        End If
        Exit Sub
    End If
    
    If Val(txt�ɿ�.Text) = 0 Then txt�ɿ�.Text = "0.00"
    If txt�ɿ�.Text <> "0.00" Then
        If mblnYB�˿� Then
            If CSng(txt�Ҳ�.Text) <= 0 Then
                'LED��ʾ
                'Call ShowLedInfor
                'ȷ��
                 Call cmdOK_Click
            Else
                MsgBox "�˿����,�벹���˿��", vbInformation, gstrSysName
                txt�ɿ�.SetFocus: zlControl.TxtSelAll txt�ɿ�
            End If
        Else
            If CSng(txt�Ҳ�.Text) >= 0 Then
                'LED��ʾ
                'Call ShowLedInfor
                'ȷ��
                 Call cmdOK_Click
            Else
                MsgBox "�ɿ����,�벹��Ӧ�ɽ�", vbInformation, gstrSysName
                txt�ɿ�.SetFocus: zlControl.TxtSelAll txt�ɿ�
            End If
        End If
        Exit Sub
    End If
    Call cmdOK_Click
   ' Call zlCommFun.PressKey(vbKeyTab) '�����ۼӽɿ�
End Sub

Private Sub Load֧����ʽ()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������Ч��֧����ʽ
    '����:���˺�
    '����:2011-07-08 11:41:36
    '˵����
    '   ȱʡ���㷽ʽ�Ĺ�������˳�����£�
    '   1.����������շѣ��ϴ�ѡ��Ľ��㷽ʽ����
    '   2.ҽ�Ƹ��ʽ���õ�ȱʡ���㷽ʽ
    '   3.ģ�����"ȱʡ���㷽ʽ"���õĽ��㷽ʽ
    '   4.���㷽ʽȱʡΪ�������е�ˢ�����
    '   5.���㷽ʽӦ�������õ�ȱʡ��
    '   6.����Ϊ"1-�ֽ���㷽ʽ"�Ľ��㷽ʽ
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim j As Long, strPayType As String, varData As Variant, varTemp As Variant, i As Long
    Dim rsTemp As ADODB.Recordset, blnFind As Boolean
    Dim strSQL As String
    
    Set rsTemp = Get���㷽ʽ("�շ�")
    Set mcolCardPayMode = New Collection
    '��|ȫ��|ˢ����־|�����ID(���ѿ����)|����|�Ƿ����ѿ�|���㷽ʽ|�Ƿ�����|�Ƿ����ƿ�;��
    If Not gobjSquare Is Nothing Then
        strPayType = gobjSquare.objSquareCard.zlGetAvailabilityCardType
    End If
    varData = Split(strPayType, ";")
    With cbo֧����ʽ
        .Clear: j = 0
        Do While Not rsTemp.EOF
            blnFind = False
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "||||||", "|")
                If varTemp(6) = Nvl(rsTemp!����) Then blnFind = True: Exit For
            Next
            If Not blnFind Then
                If Not (Val(Nvl(rsTemp!����)) = 3 Or Val(Nvl(rsTemp!����)) = 4 _
                    Or Val(Nvl(rsTemp!����)) = 7 Or Val(Nvl(rsTemp!����)) = 8 Or Val(Nvl(rsTemp!Ӧ����)) = 1) Then
                    '������ҽ���Ľ��㷽ʽ
                    .AddItem Nvl(rsTemp!����)
                    .ItemData(.NewIndex) = Val(Nvl(rsTemp!����))
                    mcolCardPayMode.Add Array("", Nvl(rsTemp!����), 0, 0, 0, 0, Nvl(rsTemp!����), 0, 0), "K" & j
                    
                    If rsTemp!ȱʡ = 1 Then .ListIndex = .NewIndex
                    If Val(Nvl(rsTemp!����)) = 1 And .ListIndex < 0 Then .ListIndex = .NewIndex
                    If mbln�������� Then
                        If gtyPrePatiPay.str���㷽ʽ = Nvl(rsTemp!����) Then
                             .ListIndex = .NewIndex
                        End If
                    End If
                    j = j + 1
              End If
            End If
            rsTemp.MoveNext
        Loop
        
        For i = 0 To UBound(varData)
            varTemp = Split(varData(i) & "||||||", "|")
            rsTemp.Filter = "����='" & varTemp(6) & "'" '���㷽ʽҪ������"����"Ӧ�ó��ϲ���ʹ��
            If Not rsTemp.EOF Then
                .AddItem varTemp(1)
                .ItemData(.NewIndex) = -1
                mcolCardPayMode.Add varTemp, "K" & j
                
                If mbln�������� Then
                    '��|ȫ��|ˢ����־|�����ID(���ѿ����)|����|�Ƿ����ѿ�|���㷽ʽ|�Ƿ�����|�Ƿ����ƿ�;��
                    If gtyPrePatiPay.lngҽ�ƿ����ID = Val(varTemp(3)) _
                        And gtyPrePatiPay.bln���ѿ� And Val(varTemp(5)) = 1 Then
                         .ListIndex = .NewIndex
                    ElseIf gtyPrePatiPay.lngҽ�ƿ����ID = Val(varTemp(3)) _
                        And gtyPrePatiPay.bln���ѿ� = False And Val(varTemp(5)) = 0 Then
                         .ListIndex = .NewIndex
                    End If
                Else
                    'ȱʡΪ�������е�ˢ�����
                    If mlngBrushCardTypeID = Val(varTemp(3)) And Val(varTemp(5)) <> 1 Then .ListIndex = .NewIndex
                End If
                j = j + 1
            End If
        Next
        
        '����ȱʡ��֧�����
        'ע�⣬���㷽ʽҽ�ƿ���ʾ���ǿ�������ƣ������ǽ��㷽ʽ
        If Not mbln�������� Then
            If gstr���㷽ʽ <> "" Then
                '60574,���ݲ�������ȱʡ��֧����𣬶���ҽ�ƿ���������������ƣ������ǽ��㷽ʽ
                For j = 0 To .ListCount - 1
                    If .List(j) = gstr���㷽ʽ Then .ListIndex = j: Exit For
                Next
            End If
            
            If mobjChargeInfor.ȱʡ���㷽ʽ <> "" Then
                '����ҽ�Ƹ��ʽ��ȱʡ���㷽ʽ����ȱʡ��֧�����
                '��|ȫ��|ˢ����־|�����ID(���ѿ����)|����|�Ƿ����ѿ�|���㷽ʽ|�Ƿ�����|�Ƿ����ƿ�;��
                For j = 1 To mcolCardPayMode.Count
                    If mcolCardPayMode(j)(6) = mobjChargeInfor.ȱʡ���㷽ʽ Then .ListIndex = j - 1: Exit For
                Next
            End If
        End If
        If .ListCount > 0 And .ListIndex < 0 Then .ListIndex = 0
        If .ListCount = 0 Then
            MsgBox "Ԥ������û�п��õĽ��㷽ʽ,���ȵ����㷽ʽ���������á�", vbExclamation, gstrSysName
            mblnUnLoad = True: Exit Sub
        End If
    End With
End Sub

Private Sub txt�ɿ�_LostFocus()
    txt�ɿ�.Text = Format(Val(txt�ɿ�.Text), "0.00")
End Sub

Private Sub txt�ɿ�_Validate(Cancel As Boolean)
    txt�ɿ�.Text = Format(Val(txt�ɿ�.Text), "0.00")
End Sub

Private Sub txt�������_GotFocus()
   zlControl.TxtSelAll txt�������
   ClearԤ����
End Sub

Private Sub txt�������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txt�������_KeyPress(KeyAscii As Integer)
    If InStr(":'|", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    zlControl.TxtCheckKeyPress txt�������, KeyAscii, m�ı�ʽ
End Sub

Private Sub txtҽ��_GotFocus()
    ClearԤ����
End Sub

Private Sub txtժҪ_GotFocus()
    zlCommFun.OpenIme True
    zlControl.TxtSelAll txtժҪ
    ClearԤ����
End Sub

Private Sub txtժҪ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If cmdOK.Visible And cmdOK.Enabled Then cmdOK.SetFocus
    End If
End Sub

Private Sub txtժҪ_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub txt�Ҳ�_GotFocus()
    zlControl.TxtSelAll txt�Ҳ�
    ClearԤ����
End Sub

Private Function zlOneCardPrayMoney(ByVal dblMoney As Double, ByRef strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:һ��֧ͨ��
    '����:һ��֧ͨ���ɹ�,����true,���򷵻�false
    '����:���˺�
    '����:2011-08-23 17:57:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl��� As Double, strҽԺ���� As String
    On Error GoTo errHandle

    If mCurCardPay.blnOneCard = False Then zlOneCardPrayMoney = True: Exit Function
    
    mrsOneCard.Filter = "���㷽ʽ='" & mCurCardPay.str���㷽ʽ & "'"
    If mrsOneCard.EOF Then zlOneCardPrayMoney = True: Exit Function
    
    'һ��ͨ���㣨�޸ĵ���ʱ��Ϊû�ж������޷�ȷ��ʹ��������һ��ͨ�����Բ�֧���޸Ĺ���)
    Dim intCardType As Integer, strSwapNO As String
    If Not mobjICCard.PaymentSwap(dblMoney, dbl���, intCardType, Val("" & mrsOneCard!ҽԺ����), mCurCardPay.strˢ������, strSwapNO, mobjChargeInfor.����ID, mobjChargeInfor.����ID) Then
        gcnOracle.RollbackTrans
        MsgBox mCurCardPay.str���㷽ʽ & "����ʧ��!", vbOKOnly, gstrSysName
        Exit Function
    End If
    mblnThreeInterface = True
    gstrSQL = "zl_һ��ͨ����_Update(" & 0 & ",'" & mCurCardPay.str���㷽ʽ & "','" & mCurCardPay.strˢ������ & "','" & intCardType & "','" & strSwapNO & "'," & dbl��� & "," & mobjChargeInfor.����ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    zlOneCardPrayMoney = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
 End Function
 
Private Function zlInterfacePrayMoney(ByRef cllUpdate As Collection, _
    ByRef cllThreeSwap As Collection, dblMoney As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ӿ�֧�����
    '����:cllPro-�޸�������������
    '        cll��������-����������������
    '����:֧���ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-17 13:34:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng����ID As Long, strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String
    
    On Error GoTo errHandle
    If mCurCardPay.lngҽ�ƿ����ID = 0 And mCurCardPay.lngҽ�ƿ����ID = 0 Then zlInterfacePrayMoney = True: Exit Function
    If cbo֧����ʽ.ItemData(cbo֧����ʽ.ListIndex) <> -1 Then zlInterfacePrayMoney = True: Exit Function
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
    '       dblMoney-֧�����
    '����:strSwapGlideNO-������ˮ��
    '       strSwapMemo-����˵��
    '       strSwapExtendInfor-������չ��Ϣ: ��ʽΪ:��Ŀ����1|��Ŀ����2||��||��Ŀ����n|��Ŀ����n
    '����:�ۿ�ɹ�,����true,���򷵻�Flase
    '˵��:
    '   ��������Ҫ�ۿ�ĵط����øýӿ�,Ŀǰ�滮��:�շ��ң��Һ���;������ѯ��;ҽ������վ��ҩ���ȡ�
    '   һ����˵���ɹ��ۿ�󣬶�Ӧ�ô�ӡ��صĽ���Ʊ�ݣ����Է��ڴ˽ӿڽ��д���.
    '   �ڿۿ�ɹ��󣬷��ؽ�����ˮ�ź���ر�ע˵���������������������Ϣ�����Է��ڽ���˵�����Ա��˷�.
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gobjSquare.objSquareCard.zlPaymentMoney(Me, mlngModule, mCurCardPay.lngҽ�ƿ����ID, mCurCardPay.bln���ѿ�, _
                mCurCardPay.strˢ������, mobjChargeInfor.����ID, mCurCardPay.strNo, dblMoney, strSwapGlideNO, _
                strSwapMemo, strSwapExtendInfor) = False Then gcnOracle.RollbackTrans: Exit Function
    '����������������
    mblnThreeInterface = True
    
    If mCurCardPay.lngҽ�ƿ����ID <> 0 And mobjChargeInfor.����ID <> 0 And cbo֧����ʽ.Visible Then
        mCurCardPay.str������ˮ�� = strSwapGlideNO
        mCurCardPay.str����˵�� = strSwapMemo
        If mCurCardPay.bln���ѿ� = False Then
            Call zlAddUpdateSwapSQL(False, mobjChargeInfor.����ID, mCurCardPay.lngҽ�ƿ����ID, mCurCardPay.bln���ѿ�, mCurCardPay.strˢ������, strSwapGlideNO, strSwapMemo, cllUpdate, 2)
        End If
        Call zlAddThreeSwapSQLToCollection(False, mobjChargeInfor.����ID, mCurCardPay.lngҽ�ƿ����ID, mCurCardPay.bln���ѿ�, mCurCardPay.strˢ������, strSwapExtendInfor, cllThreeSwap)
    End If
    zlInterfacePrayMoney = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ChargeOver(ByVal dbl��֧Ʊ�� As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�շ����
    '���:blnNotCommit-�Ƿ�û�н��������ύ�����ʱ���ύ����(ԭ���Ƕ���ͨ���˽���һ���ύ)
    '����:���˺�
    '����:2011-08-15 15:50:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim dbl�ɿ� As Double, dbl�Ҳ� As Double
    Dim str�շѽ��� As String, dblԤ��� As Double
    dblԤ��� = 0
    str�շѽ��� = Get�շѽ���(dblԤ���)
    On Error GoTo errHandle
    If mCurCardPay.int���� = 1 And mblnCur���� = False Then
        dbl�ɿ� = Val(txt�ɿ�.Text)
        dbl�Ҳ� = Val(txt�Ҳ�.Text)
    End If
    If dbl�ɿ� = 0 Then
        dbl�ɿ� = 0: dbl�Ҳ� = 0
    End If
    
    ' Zl_�����շѽ���_Modify
    strSQL = "Zl_�����շѽ���_Modify("
    '    ��������_In   Number,
    '    --��������_In:
    '    --   0-��ͨ�շѷ�ʽ:
    '    --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ||.." ;Ҳ�������.
    '    --     �ڳ�Ԥ��_In:����漰Ԥ����,���뱾�εĳ�Ԥ��,�������շ�ʱ,������
    '    --     ����֧Ʊ��_In:����漰��֧Ʊ,���뱾�ε���֧Ʊ��,�������շ�ʱ,������
    '    --   1.����������:
    '    --     �ٽ��㷽ʽ_IN:ֻ�ܴ���һ�����㷽ʽ,���������һЩ������Ϣ,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ"
    '    --     �ڳ�Ԥ��_In: ������
    '    --     ����֧Ʊ��_In:������
    '    --     �ܿ����ID_IN,����_IN,������ˮ��_IN,����˵��_In:��Ҫ����
    '    --   2-ҽ������(�������ҽ���Ľ���,��Ҫ��ɾ��ԭҽ������,���´���ĸ���)
    '    --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:���㷽ʽ|������||.."
    '    --     �ڳ�Ԥ��_In: ������
    '    --     ����֧Ʊ��_In:������
    '    --   3-���ѿ�����:
    '    --     �ٽ��㷽ʽ_IN:����һ��ˢ���ſ�,��ʽΪ:�����ID|����|���ѿ�ID|���ѽ��||."
    '    --     �ڳ�Ԥ��_In: ������
    '    --     ����֧Ʊ��_In:������
    strSQL = strSQL & 0 & ","
    '    ����id_In     ������ü�¼.����id%Type,
    strSQL = strSQL & mobjChargeInfor.����ID & ","
    '    ����id_In     ����Ԥ����¼.����id%Type,
    strSQL = strSQL & mobjChargeInfor.����ID & ","
    '    ���㷽ʽ_In   Varchar2,
    strSQL = strSQL & "'" & str�շѽ��� & "'" & ","
    '    ��Ԥ��_In     ����Ԥ����¼.��Ԥ��%Type := Null,
    strSQL = strSQL & dblԤ��� & ","
    '    ��֧Ʊ��_In   ����Ԥ����¼.��Ԥ��%Type := Null,
    strSQL = strSQL & dbl��֧Ʊ�� & ","
    '    �����id_In   ����Ԥ����¼.�����id%Type := Null,
    strSQL = strSQL & "NULL,"
    '    ����_In       ����Ԥ����¼.����%Type := Null,
    strSQL = strSQL & "NULL,"
    '    ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
    strSQL = strSQL & "NULL,"
    '    ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
    strSQL = strSQL & "NULL,"
    '    �ɿ�_In       ����Ԥ����¼.�ɿ�%Type := Null,
    strSQL = strSQL & "" & dbl�ɿ� & ","
    '    �Ҳ�_In       ����Ԥ����¼.�Ҳ�%Type := Null,
    strSQL = strSQL & "" & dbl�Ҳ� & ","
    '    �����_In   ������ü�¼.ʵ�ս��%Type := Null,
    '    -- �����_In:��������ʱ,����
    strSQL = strSQL & "" & mCurCarge.dbl�������� & ","
    '    ��ɽ���_In Number:=0
    '    -- ��ɽ���_In:1-����շ�;0-δ����շ�
    strSQL = strSQL & "1,"
    '77141,Ƚ����,2014-8-26,������ò����շ�/�˷Ѻ�,û�н�����Ϣ
    'ȱʡ���㷽ʽ_In ���㷽ʽ.����%Type := Null
    strSQL = strSQL & "'" & Trim(cbo֧����ʽ.Text) & "',"
    '79868,Ƚ����,2015-06-10,ʹ�ò��˼���Ԥ��
    '��Ԥ������ids_In Varchar2:=Null
    strSQL = strSQL & "'" & mobjChargeInfor.����ID & "," & mobjChargeInfor.����IDs & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    mobjChargeInfor.�ɿ� = dbl�ɿ�: mobjChargeInfor.�Ҳ� = dbl�Ҳ�
    ChargeOver = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Show�����(ByVal blnԤ�� As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ�����
    '���:blnԤ��-Ԥ����
    '����:���˺�
    '����:2011-09-30 15:40:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, dbl��֧Ʊ�� As Double
    Dim dblʣ���� As Double, dblTemp As Double
    
    mCurCarge.dbl�������� = 0
    dblMoney = IIf(blnԤ��, Val(txt��Ԥ��.Text), IIf(mblnYB�˿�, -1, 1) * Val(txt�ɿ�.Text))
    dbl��֧Ʊ�� = 0
    dblʣ���� = RoundEx(mCurCarge.dbl��ǰδ�� - dblMoney, 6)
    If blnԤ�� Then
        dblMoney = Val(txt��Ԥ��.Text)
        mCurCarge.dbl�������� = RoundEx(mCurCarge.dbl����ʵ�� - mCurCarge.dbl�����Ѹ��ϼ� - RoundEx(mCurCarge.dbl��ǰδ��, 2), 6)
    ElseIf mCurCardPay.int���� = 1 Then
        dblTemp = IIf(dblMoney = 0, dblʣ����, mCurCarge.dbl��ǰδ��): dblʣ���� = 0
        If mobjChargeInfor.intInsure > 0 Then  '����:43855
            If mInsurePara.�ֱҴ��� Then
                dblMoney = CentMoney(CCur(dblTemp))
            Else
                dblMoney = Format(dblTemp, "0.00")
            End If
        Else
            dblMoney = CentMoney(CCur(dblTemp))
        End If
        mCurCarge.dbl�������� = mCurCarge.dbl����ʵ�� - mCurCarge.dbl�����Ѹ��ϼ� - dblMoney
    Else
        mCurCarge.dbl�������� = mCurCarge.dbl����ʵ�� - mCurCarge.dbl�����Ѹ��ϼ� - RoundEx(mCurCarge.dbl��ǰδ��, 2)
    End If
    
    '����:47637
    'δ����ҽ������ǰ,����ʾ���
    If mobjChargeInfor.intInsure <> 0 And mblnYbBalanced = False Then mCurCarge.dbl�������� = 0
    mCurCarge.dbl�������� = RoundEx(mCurCarge.dbl��������, 6)
    pic���.Visible = mCurCarge.dbl�������� <> 0
    lbl����.Caption = FormatEx(mCurCarge.dbl��������, 6)
End Sub

Private Function zlCheckMulitInterfaceNumValied(Optional blnԤ�� As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ͬʱ�����������Ͻӿ�(��������)
    '����:�����������Ͻӿڵ�,����true,���򷵻�False
    '����:���˺�
    '����:2012-02-07 15:07:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intCount As Integer, i As Long, int���� As Integer, str���㷽ʽ As String
    Dim varData As Variant, strErrMsg As String
    On Error GoTo errHandle
    strErrMsg = ""
    
    If blnԤ�� Or mCurCardPay.lngҽ�ƿ����ID = 0 Or cbo֧����ʽ.ItemData(cbo֧����ʽ.ListIndex) <> -1 Then
        zlCheckMulitInterfaceNumValied = True
        Exit Function
    End If
    
    'ҽ����һ���ӿ�
    If mobjChargeInfor.intInsure <> 0 And mblnYbBalanced Then intCount = intCount + 1: strErrMsg = strErrMsg & "ҽ������:" & txtҽ��.Text
   With vsBlance
        For i = 1 To .Rows - 1
            str���㷽ʽ = Trim(.TextMatrix(i, .ColIndex("֧����ʽ")))
            int���� = Val(.RowData(i))
            'rowdata:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
            If InStr("34", int����) > 0 Then
                If int���� = 4 Then intCount = intCount + 1
                If int���� = 3 Then '�����ӿ�
                    intCount = intCount + 1: strErrMsg = strErrMsg & vbCrLf & str���㷽ʽ & ":" & .TextMatrix(i, .ColIndex("֧�����"))
                End If
            End If
        Next
    End With
    If intCount > 2 Then
        Call MsgBox("ע��:" & vbCrLf & "   ��ϵͳĿǰֻ֧���������½ӿ�,�����Ѿ��������½ӿڽ���:" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName)
        Exit Function
    End If
    zlCheckMulitInterfaceNumValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function SaveCharge(Optional blnԤ�� As Boolean, _
    Optional ByRef blnUnload As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����������
    '����:blnUnload-�Ƿ��շ���ɣ��˳��󣬽�Unload����
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-14 17:38:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str���ѿ����� As String, str�շѽ��� As String, strSQL As String
    Dim strCardNo As String, strErrMsg As String
    Dim blnHaveMoney As Boolean, blnFind As Boolean, blnTrans As Boolean
    Dim dblʣ���� As Double, dblTemp As Double
    Dim dblMoney As Double, dbl��֧Ʊ�� As Double
    Dim i As Integer, j As Long
    Dim cllUpdate As Collection, cllThreeSwap As Collection, cllPro As Collection
    Dim objCard As Card, dblCheckMoney As Double
    
    On Error GoTo errHandle
    
    blnUnload = False
    If zlCheckMulitInterfaceNumValied = False Then Exit Function
    
    mobjChargeInfor.�շѽ��� = "" '����:42791
    
    mdbl�ֽ� = 0
    dblMoney = IIf(blnԤ��, Val(txt��Ԥ��.Text), IIf(mblnYB�˿�, -1, 1) * Val(txt�ɿ�.Text)) - IIf(mblnCur����, 0, mCurCarge.dblӦ���ۼ�)
    dbl��֧Ʊ�� = 0
    dblʣ���� = mCurCarge.dbl��ǰδ�� - dblMoney
    
    '��������Ľ���ʵ�ʱ���Ľ��ֿ�����Ҫ��Ϊ�����շ�ʱ��Ӧ�ɽ���ʵ�����ݱ����һ�£�
    If Val(txt�ɿ�.Text) = 0 Then
        dblCheckMoney = mCurCarge.dbl��ǰδ�� + IIf(mblnCur����, 0, mCurCarge.dblӦ���ۼ�)
    Else
        dblCheckMoney = IIf(mblnYB�˿�, -1, 1) * Val(txt�ɿ�.Text)
    End If
    If mblnCur���� = False And mCurCarge.dblӦ���ۼ� <> 0 Then
        dblMoney = IIf(dblMoney = 0, dblʣ����, mCurCarge.dbl��ǰδ��)
    End If
    
    If blnԤ�� Then
        dblMoney = Val(txt��Ԥ��.Text)
        mobjChargeInfor.�շѽ��� = mobjChargeInfor.�շѽ��� & "|��Ԥ��:" & dblMoney
        '����:58344
        '����Ƿ�ǰ֧�����Ϊ����,�Ǹ���ʱ,��Ҫ���Ѳ���Ա(��Ҫ��ҽ������ʱ���ܴ��ڱ����ݵķ���)
        If mblnYB�˿� And mblnCur���� = False Then
              Call MsgBox("ע��:" & vbCrLf & "    ��ǰ�����˿ʽ,������ʹ��Ԥ����!", vbExclamation + vbOKOnly + vbDefaultButton2, gstrSysName)
              Exit Function
        End If
        
    ElseIf mCurCardPay.int���� = 1 Then
        dblTemp = IIf(dblMoney = 0, dblʣ����, mCurCarge.dbl��ǰδ��): dblʣ���� = 0
        If mobjChargeInfor.intInsure > 0 Then  '����:43855
            If gclsInsure.GetCapability(support�ֱҴ���, , mobjChargeInfor.intInsure) Then
                dblMoney = CentMoney(CCur(dblTemp))
                dblCheckMoney = CentMoney(CCur(dblCheckMoney))
            Else
                dblMoney = Format(dblTemp, "0.00")
                dblCheckMoney = Format(dblCheckMoney, "0.00")
            End If
        Else
            dblMoney = CentMoney(CCur(dblTemp))
            dblCheckMoney = CentMoney(CCur(dblCheckMoney))
        End If
        '����:58344
        '����Ƿ�ǰ֧�����Ϊ����,�Ǹ���ʱ,��Ҫ���Ѳ���Ա(��Ҫ��ҽ������ʱ���ܴ��ڱ����ݵķ���)
        If mblnYB�˿� And mblnCur���� = False Then
             If MsgBox("ע��:" & vbCrLf & "    δ������Ϊ�˿�,���Ƿ����Ҫ�ˡ�" & mCurCardPay.str���㷽ʽ & ":" & Abs(dblCheckMoney) & "��������?" & vbCrLf & IIf(Val(txt�ɿ�.Text) <> 0, "  ��ǰ�˸������ܶ�:" & txt�ɿ�.Text & vbCrLf & "  ��ǰӦ�ջ��ܶ�:" & Abs(txt�Ҳ�.Text), ""), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
             If Abs(dblCheckMoney) < Abs(lblʣ���Ը�.Caption) Then
                Call MsgBox("ע��:" & vbCrLf & "    δ������Ϊ�˿�,�㲻�ܽ��ж���˿����," & vbCrLf & "��ǰ�˽��(" & Format(dblCheckMoney, "0.00") & ")�������ʣ����(" & lblʣ���Ը�.Caption & ")!", vbInactiveBorder + vbOKOnly + vbDefaultButton2, gstrSysName)
                Exit Function
             End If
        End If
        mdbl�ֽ� = dblMoney
        If Val(txt�ɿ�.Text) <> 0 Then
            mobjChargeInfor.�շѽ��� = mobjChargeInfor.�շѽ��� & "|�ɿ�:" & IIf(mblnYB�˿�, -1, 1) * Val(txt�ɿ�.Text) & ":1"
            mobjChargeInfor.�շѽ��� = mobjChargeInfor.�շѽ��� & "|�Ҳ�:" & IIf(mblnYB�˿�, -1, 1) * Val(txt�Ҳ�.Text) & ":2"
        End If
        mobjChargeInfor.�շѽ��� = mobjChargeInfor.�շѽ��� & "|" & mCurCardPay.str���㷽ʽ & ":" & dblMoney
    ElseIf mCurCardPay.bln֧Ʊ Then
        mobjChargeInfor.�շѽ��� = mobjChargeInfor.�շѽ��� & "|" & mCurCardPay.str���㷽ʽ & ":" & dblMoney
        '����:58344
        '����Ƿ�ǰ֧�����Ϊ����,�Ǹ���ʱ,��Ҫ���Ѳ���Ա(��Ҫ��ҽ������ʱ���ܴ��ڱ����ݵķ���)
        If mblnYB�˿� And mblnCur���� = False Then
             If MsgBox("ע��:" & vbCrLf & "    δ������Ϊ�˿�,���Ƿ����Ҫ�ˡ�" & mCurCardPay.str���㷽ʽ & ":" & Abs(dblCheckMoney) & "��������?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
             If Abs(dblCheckMoney) <> Abs(lblʣ���Ը�.Caption) Then
                Call MsgBox("ע��:" & vbCrLf & "    δ������Ϊ�˿�,��ǰ�˽��(" & Format(Abs(dblCheckMoney), "0.00") & ")�������ʣ����(" & Abs(Val(lblʣ���Ը�.Caption)) & ")!", vbInactiveBorder + vbOKOnly + vbDefaultButton2, gstrSysName)
                Exit Function
             End If
        Else
            If dblʣ���� < 0 Then
                If mstr��֧Ʊ = "" Then
                    MsgBox "�ڽ��㷽ʽ��û������Ӧ����Ľ��㷽ʽ,���ܽ�����֧Ʊ����", vbOKOnly + vbInformation, gstrSysName
                    Exit Function
                End If
                dbl��֧Ʊ�� = -1 * Val(txt�Ҳ�.Text)
                mobjChargeInfor.�շѽ��� = mobjChargeInfor.�շѽ��� & "|" & mstr��֧Ʊ & ":" & -1 * dbl��֧Ʊ�� & ":2"
            End If
        End If
    Else
        '����:58344
        '����Ƿ�ǰ֧�����Ϊ����,�Ǹ���ʱ,��Ҫ���Ѳ���Ա(��Ҫ��ҽ������ʱ���ܴ��ڱ����ݵķ���)
        If mblnYB�˿� And mblnCur���� = False Then
             If MsgBox("ע��:" & vbCrLf & "    δ������Ϊ�˿�,���Ƿ����Ҫ�ˡ�" & mCurCardPay.str���㷽ʽ & ":" & Abs(dblCheckMoney) & "��������?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
             If Abs(dblCheckMoney) <> Abs(lblʣ���Ը�.Caption) Then
                Call MsgBox("ע��:" & vbCrLf & "    δ������Ϊ�˿�,��ǰ�˽��(" & Format(Abs(dblCheckMoney), "0.00") & ")�������ʣ����(" & Abs(Val(lblʣ���Ը�.Caption)) & ")!", vbInactiveBorder + vbOKOnly + vbDefaultButton2, gstrSysName)
                Exit Function
             End If
        End If
        mobjChargeInfor.�շѽ��� = mobjChargeInfor.�շѽ��� & "|" & mCurCardPay.str���㷽ʽ & ":" & dblMoney
    End If
    
    If mblnCur���� And Val(txt�ɿ�.Text) = 0 Then
        If mCurCardPay.int���� <> 1 Or dblMoney = 0 Then
            dblMoney = mCurCarge.dbl��ǰδ��
        End If
        dblʣ���� = 0
    End If
    
    Call Show�����(blnԤ��)
    If mCurCardPay.int���� = 1 Then
        If Abs(mCurCarge.dbl��������) > 1.5 Then
            Call MsgBox("������,�����Ƿ���ȷ!", vbInformation + vbOKOnly, gstrSysName)
            Exit Function
        End If
    End If
    
    'mCurCarge.dbl�������� = Format(-1 * (mCurCarge.dbl����ʵ�� - mCurCarge.dbl�����Ѹ��ϼ� - dblMoney - dbl��֧Ʊ��), gstrDec)
    '���ܴ���10��Ǯ
    If RoundEx(dblʣ����, 2) > 0 Then blnHaveMoney = True
    With vsBlance
        blnFind = False
        For i = 1 To .Rows - 1
            ' '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
            If blnԤ�� Then
                If Val(.RowData(i)) = 1 Then blnFind = True
            ElseIf mCurCardPay.bln���ѿ� And mCurCardPay.bln���ƿ� Then
                '���ѿ�,�Ѿ����,�����ٴ���
            Else
                If .TextMatrix(i, .ColIndex("֧����ʽ")) = mCurCardPay.str���㷽ʽ Then
                    blnFind = True
                End If
            End If
            mobjChargeInfor.�շѽ��� = mobjChargeInfor.�շѽ��� & "|" & .TextMatrix(i, .ColIndex("֧����ʽ")) & ":" & .TextMatrix(i, .ColIndex("֧�����"))
        Next
        
        If blnFind Then
            If blnԤ�� Then
                MsgBox "�Ѿ���Ԥ���֧��,ֻ��ɾ��Ԥ�������֧��!", vbOKOnly, gstrSysName
            Else
                MsgBox mCurCardPay.str���㷽ʽ & " �Ѿ�֧����,��������" & mCurCardPay.str���㷽ʽ & "����֧��", vbOKOnly + vbDefaultButton1, gstrSysName
            End If
            Exit Function
        End If
    End With
    
    If blnHaveMoney = False And dblMoney = 0 Then
        GoTo GoOver:
    End If
    
    Set cllPro = New Collection
    Set cllUpdate = New Collection
    Set cllThreeSwap = New Collection
    str���ѿ����� = ""  '�����ID|����|���ѿ�ID|���ѽ��||....
    If mCurCardPay.bln���ѿ� And Not blnԤ�� Then
        If mcllCurSquareBalance Is Nothing Then Exit Function
        If mcllCurSquareBalance.Count = 0 Then Exit Function
        For j = 1 To mcllCurSquareBalance.Count
            ' array(�����ID,���ѿ�ID,ˢ�����,����,����,�������,�Ƿ�����)
            str���ѿ����� = str���ѿ����� & "||" & Val(mcllCurSquareBalance(j)(0))
            str���ѿ����� = str���ѿ����� & "|" & mcllCurSquareBalance(j)(3)
            str���ѿ����� = str���ѿ����� & "|" & Val(mcllCurSquareBalance(j)(1))
            str���ѿ����� = str���ѿ����� & "|" & Val(mcllCurSquareBalance(j)(2))
        Next
        If str���ѿ����� <> "" Then str���ѿ����� = Mid(str���ѿ�����, 3)
    End If
    
    If Not (blnԤ�� Or mCurCardPay.lngҽ�ƿ����ID = 0 Or cbo֧����ʽ.ItemData(cbo֧����ʽ.ListIndex) <> -1) Then
        '�������ӿڵ���ؽ���,��Ҫ�ȴ���ӿ�����
        
        '�ȱ��浥��
        blnTrans = True
        If SaveFeeBilL = False Then blnTrans = False: Exit Function
        If mblnElsePersonErrBill Then
            If UpdateElsePersonErrBill() = False Then blnTrans = False: Exit Function
        End If
        
        ' Zl_�����շѽ���_Modify
        strSQL = "Zl_�����շѽ���_Modify("
        '    ��������_In   Number,
        '    --��������_In:
        '    --   0-��ͨ�շѷ�ʽ:
        '    --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ||.." ;Ҳ�������.
        '    --     �ڳ�Ԥ��_In:����漰Ԥ����,���뱾�εĳ�Ԥ��,�������շ�ʱ,������
        '    --     ����֧Ʊ��_In:����漰��֧Ʊ,���뱾�ε���֧Ʊ��,�������շ�ʱ,������
        '    --   1.����������:
        '    --     �ٽ��㷽ʽ_IN:ֻ�ܴ���һ�����㷽ʽ,���������һЩ������Ϣ,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ"
        '    --     �ڳ�Ԥ��_In: ������
        '    --     ����֧Ʊ��_In:������
        '    --     �ܿ����ID_IN,����_IN,������ˮ��_IN,����˵��_In:��Ҫ����
        '    --   2-ҽ������(�������ҽ���Ľ���,��Ҫ��ɾ��ԭҽ������,���´���ĸ���)
        '    --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:���㷽ʽ|������||.."
        '    --     �ڳ�Ԥ��_In: ������
        '    --     ����֧Ʊ��_In:������
        '    --   3-���ѿ�����:
        '    --     �ٽ��㷽ʽ_IN:����һ��ˢ���ſ�,��ʽΪ:�����ID|����|���ѿ�ID|���ѽ��||."
        '    --     �ڳ�Ԥ��_In: ������
        '    --     ����֧Ʊ��_In:������
        If mCurCardPay.bln���ѿ� Then
            strSQL = strSQL & "3" & ","
        Else
            strSQL = strSQL & "1" & ","
        End If
        '    ����id_In     ������ü�¼.����id%Type,
        strSQL = strSQL & mobjChargeInfor.����ID & ","
        '    ����id_In     ����Ԥ����¼.����id%Type,
        strSQL = strSQL & mobjChargeInfor.����ID & ","
        '    ���㷽ʽ_In   Varchar2,
        If mCurCardPay.bln���ѿ� Then
            strSQL = strSQL & "'" & str���ѿ����� & "'" & ","
        Else
            '"���㷽ʽ|������|�������|����ժҪ||.." ;Ҳ�������.
            str�շѽ��� = mCurCardPay.str���㷽ʽ
            str�շѽ��� = str�շѽ��� & "|" & dblMoney
            str�շѽ��� = str�շѽ��� & "|" & IIf(txt�������.Text = "", " ", txt�������.Text)
            str�շѽ��� = str�շѽ��� & "|" & IIf(txtժҪ.Text = "", " ", txtժҪ.Text)
            strSQL = strSQL & "'" & str�շѽ��� & "'" & ","
        End If
        '    ��Ԥ��_In     ����Ԥ����¼.��Ԥ��%Type := Null,
        strSQL = strSQL & "NULL,"
        '    ��֧Ʊ��_In   ����Ԥ����¼.��Ԥ��%Type := Null,
        strSQL = strSQL & "NULL,"
        '    �����id_In   ����Ԥ����¼.�����id%Type := Null,
        strSQL = strSQL & "" & IIf(mCurCardPay.lngҽ�ƿ����ID = 0, "NULL", mCurCardPay.lngҽ�ƿ����ID) & ","
        '    ����_In       ����Ԥ����¼.����%Type := Null,
        strSQL = strSQL & "" & IIf(mCurCardPay.strˢ������ <> "", "'" & mCurCardPay.strˢ������ & "'", "NULL") & ","
        '    ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
        strSQL = strSQL & "NULL,"
        '    ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
        strSQL = strSQL & "NULL,"
        '    �ɿ�_In       ����Ԥ����¼.�ɿ�%Type := Null,
        strSQL = strSQL & "NULL,"
        '    �Ҳ�_In       ����Ԥ����¼.�Ҳ�%Type := Null,
        strSQL = strSQL & "NULL,"
        '    �����_In   ������ü�¼.ʵ�ս��%Type := Null,
        '    -- �����_In:��������ʱ,����
        strSQL = strSQL & "NULL,"
        '    ��ɽ���_In Number:=0
        '    -- ��ɽ���_In:1-����շ�;0-δ����շ�
        strSQL = strSQL & "0)"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        If Not mCurCardPay.bln���ѿ� Then
             If zlInterfacePrayMoney(cllUpdate, cllThreeSwap, dblMoney) = False Then blnTrans = False: Exit Function
        End If
        
        'һ��ͨ����(�ϰ�)
        If zlOneCardPrayMoney(dblMoney, strErrMsg) = False Then
            gcnOracle.RollbackTrans
            MsgBox strErrMsg, vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
        
        gcnOracle.CommitTrans:  mblnCommitBill = True
        mblnElsePersonErrBill = False '�Ѹ���
        Call zlExecuteProcedureArrAy(cllUpdate, Me.Caption)
        
        blnTrans = False
        Call SetCtrlVisible
         blnTrans = True
        Call zlExecuteProcedureArrAy(cllThreeSwap, Me.Caption)
         blnTrans = False
    End If
GoOver:
    If mobjChargeInfor.intInsure <> 0 Then
        If Not (blnԤ�� Or mCurCardPay.lngҽ�ƿ����ID <> 0 _
            Or mCurCardPay.blnOneCard) Then
            'ֻ��ҽ�����˲Ż�������½϶Ե����,��˲Ż����¼��㱾��Ӧ�ɵ����
            '��Ҫ�Ǹ��������շѵ�����
            mobjChargeInfor.����Ӧ�� = mobjChargeInfor.����Ӧ�� + dblMoney
        End If
    End If
    
    If Not blnHaveMoney Then
        blnTrans = True
        If SaveFeeBilL = False Then blnTrans = False: Exit Function
        If mblnElsePersonErrBill Then
            If UpdateElsePersonErrBill() = False Then blnTrans = False: Exit Function
        End If
        If ChargeOver(dbl��֧Ʊ��) = False Then blnTrans = False:    Exit Function
        gcnOracle.CommitTrans:  mblnCommitBill = True
        mblnElsePersonErrBill = False '�Ѹ���
        blnTrans = False
        Call WhriteTotalDataToReCord(IIf(blnԤ��, dblMoney, 0), IIf(Not blnԤ��, dblMoney, 0), dbl��֧Ʊ��)
        mblnOK = True: SaveCharge = True: mblnUnloaded = True
        blnUnload = True
         Exit Function
    End If
    
    mobjChargeInfor.�շѽ��� = ""
    If Not blnԤ�� And mCurCardPay.int���� = 1 Then
       '�ֽ�
        SaveCharge = True: Exit Function
    End If
    
    Err = 0: On Error GoTo errHandle:
    With vsBlance
        If mCurCardPay.bln���ѿ� And Not blnԤ�� Then
            If mcllSquareBalance Is Nothing Then Set mcllSquareBalance = New Collection
            
            For j = 1 To mcllCurSquareBalance.Count
                '��ǰˢ������(array(�����ID,���ѿ�ID,ˢ�����,����,����,�������,�Ƿ�����))
                mcllSquareBalance.Add mcllCurSquareBalance(j)
                If Not (.Rows = 2 And Trim(.TextMatrix(1, .ColIndex("֧����ʽ"))) = "") Then
                    .Rows = .Rows + 1
                    .RowPosition(.Rows - 1) = 1
                End If
                '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
                .RowData(1) = 5
                .TextMatrix(1, .ColIndex("֧����ʽ")) = mCurCardPay.str���㷽ʽ
                 ' ҽ�ƿ����ID|���ѿ�(1, 0) |���ƿ�| �ӿ�����
                .Cell(flexcpData, 1, .ColIndex("֧����ʽ")) = Val(mcllCurSquareBalance(j)(0)) & "|" & 1 & "|" & IIf(mCurCardPay.bln���ƿ�, 1, 0) & "|" & mCurCardPay.str����
                .RowData(1) = 2
                strCardNo = Trim(mcllCurSquareBalance(j)(3))
                .TextMatrix(1, .ColIndex("����")) = IIf(mCurCardPay.bln��������, String(Len(strCardNo), "*"), strCardNo)
                .Cell(flexcpData, 1, .ColIndex("����")) = strCardNo
                .TextMatrix(1, .ColIndex("֧�����")) = Format(Val(mcllCurSquareBalance(j)(2)), "0.00")
                .TextMatrix(1, .ColIndex("�������")) = ""
                .TextMatrix(1, .ColIndex("��ע")) = ""
                mCurCarge.dbl�����Ѹ��ϼ� = RoundEx(mCurCarge.dbl�����Ѹ��ϼ� + Val(mcllCurSquareBalance(j)(2)), 6)
                mCurCarge.dbl��ǰδ�� = RoundEx(mCurCarge.dbl��ǰδ�� - Val(mcllCurSquareBalance(j)(2)), 6)
            Next
        Else
            If Not (.Rows = 2 And Trim(.TextMatrix(1, .ColIndex("֧����ʽ"))) = "") Then
                .Rows = .Rows + 1
                .RowPosition(.Rows - 1) = 1
            End If
            '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
            .RowData(1) = 0
            strCardNo = mCurCardPay.strˢ������
            If blnԤ�� Then
                .TextMatrix(1, .ColIndex("֧����ʽ")) = "Ԥ���"
                .RowData(1) = 1
            ElseIf mCurCardPay.lngҽ�ƿ����ID <> 0 Then
                Set objCard = GetPayCard(mCurCardPay.lngҽ�ƿ����ID, mCurCardPay.bln���ѿ�, False)
                .TextMatrix(1, .ColIndex("֧����ʽ")) = objCard.���㷽ʽ
                If Not objCard Is Nothing Then
                    'ҽ�ƿ����ID|���ѿ�(1, 0) |���ƿ�|�Ƿ�ȫ��|�Ƿ�����|�ӿ�����
                    .Cell(flexcpData, 1, .ColIndex("֧����ʽ")) = mCurCardPay.lngҽ�ƿ����ID & "|" & IIf(mCurCardPay.bln���ѿ�, 1, 0) & "|" & IIf(mCurCardPay.bln���ƿ�, 1, 0) & "|" & IIf(objCard.�Ƿ�ȫ��, 1, 0) & "|" & IIf(objCard.�Ƿ�����, 1, 0) & "|" & mCurCardPay.str����
                Else
                    .Cell(flexcpData, 1, .ColIndex("֧����ʽ")) = mCurCardPay.lngҽ�ƿ����ID & "|" & IIf(mCurCardPay.bln���ѿ�, 1, 0) & "|" & IIf(mCurCardPay.bln���ƿ�, 1, 0) & "|" & 0 & "|" & 0 & "|" & mCurCardPay.str����
                End If
                .RowData(1) = 3
                strCardNo = gobjSquare.objSquareCard.zlGetCardNODencode(mCurCardPay.strˢ������, mCurCardPay.lngҽ�ƿ����ID, mCurCardPay.bln���ѿ�)
            ElseIf mCurCardPay.blnOneCard Then
                
                .TextMatrix(1, .ColIndex("֧����ʽ")) = mCurCardPay.str���㷽ʽ
                .RowData(1) = 4
            Else
                .TextMatrix(1, .ColIndex("֧����ʽ")) = mCurCardPay.str���㷽ʽ
            End If
            .TextMatrix(1, .ColIndex("֧�����")) = Format(dblMoney, "0.00")
            If Not blnԤ�� Then
                .TextMatrix(1, .ColIndex("�������")) = IIf(txt�������.Visible, Trim(txt�������.Text), "")
                .TextMatrix(1, .ColIndex("��ע")) = Trim(txtժҪ.Text)
                
                .TextMatrix(1, .ColIndex("����")) = IIf(mCurCardPay.bln��������, String(Len(strCardNo), "*"), strCardNo)
                .Cell(flexcpData, 1, .ColIndex("����")) = mCurCardPay.strˢ������
                .TextMatrix(1, .ColIndex("������ˮ��")) = mCurCardPay.str������ˮ��
                .TextMatrix(1, .ColIndex("����˵��")) = mCurCardPay.str����˵��
            End If
            mCurCarge.dbl�����Ѹ��ϼ� = RoundEx(mCurCarge.dbl�����Ѹ��ϼ� + dblMoney, 6)
            mCurCarge.dbl��ǰδ�� = RoundEx(mCurCarge.dbl��ǰδ�� - dblMoney, 6)
        End If
        
        For i = 0 To cbo֧����ʽ.ListCount
            'ȱʡ��λ���ֽ���
            If cbo֧����ʽ.ItemData(i) = 1 Then cbo֧����ʽ.ListIndex = i: Exit For
        Next
        Call SetControlProperty
        txt�ɿ�.Text = ""
        If txt�ɿ�.Enabled And txt�ɿ�.Visible Then txt�ɿ�.SetFocus
        Call LedDisplayBank
    End With
    Call SetDeleteVisible
    SaveCharge = True
    Exit Function
errHandle:
    If blnTrans Then
        gcnOracle.RollbackTrans
        Call ErrCenter
    Else
        If ErrCenter() = 1 Then
            Resume
        End If
    End If
End Function

Private Function UpdateElsePersonErrBill()
    '�����쳣���ݣ����ⲿ�ֵ��ݸ���Ϊ��ǰ����Ա
    Dim strSQL As String
    
    On Error GoTo errHandler
    'Zl_�����쳣�շ�_���²���Ա
    strSQL = "Zl_�����쳣�շ�_���²���Ա("
    '����id_In     ������ü�¼.����id%Type,
    strSQL = strSQL & "" & mobjChargeInfor.����ID & ","
    '����Ա���_In ������ü�¼.����Ա���%Type,
    strSQL = strSQL & "'" & UserInfo.��� & "',"
    '����Ա����_In ������ü�¼.����Ա����%Type,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '�������_In   ����Ԥ����¼.�������%Type
    strSQL = strSQL & mobjChargeInfor.������� & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    UpdateElsePersonErrBill = True
    Exit Function
errHandler:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub txt�Ҳ�_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txt�Ҳ�_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lbl�Ҳ�.Caption <> "�Ҳ�" Then
      ''  zlCommFun.ShowTipInfo txt�Ҳ�.hWnd, mstrӦ������㷽ʽ, False
    Else
        zlCommFun.ShowTipInfo txt�Ҳ�.hWnd, "", False
    End If
End Sub

Private Sub vsBlance_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow = NewRow Then Exit Sub
    If NewRow < 0 Then Exit Sub
    Call SetDeleteVisible
End Sub

Private Sub SetDeleteVisible()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ɾ���ؼ���visible����
    '����:���˺�
    '����:2011-09-20 10:42:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim int���� As Integer
    If vsBlance.Row < 1 Then
        int���� = -1
    ElseIf vsBlance.TextMatrix(vsBlance.Row, vsBlance.ColIndex("֧����ʽ")) = "" Then
        int���� = -1
    Else
        int���� = Val(vsBlance.RowData(vsBlance.Row))
    End If
     '.rowdata: '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
    cmdDel.Visible = (int���� = 0 Or int���� = 1) And mbytFunc <> EM_FUN_����
End Sub

Private Sub SetWindowsSize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ô����С
    '����:���˺�
    '����:2011-09-15 11:26:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If OS.IsDesinMode Then Exit Sub
    '��С����ߴ�
    With gWinRect
        .MaxW = Me.Width
        .MaxH = Screen.Height * Screen.TwipsPerPixelY
        .MinH = Me.Height
        .MinW = Me.Width
    End With
    glngOld = GetWindowLong(hWnd, GWL_WNDPROC)
    Call SetWindowLong(hWnd, GWL_WNDPROC, AddressOf SetWindowResizeWndMessage)
End Sub

Private Function zlCheckDelValied(ByVal lng�����ID As Long, _
     ByVal strName As String, _
     ByVal bln���ѿ� As Boolean, ByVal strCardNo As String, _
     ByVal strSwapNO As String, strSwapMemo As String, _
     ByRef lng����ID As Long, ByVal dbl�˿��� As Double, _
     ByVal bln�Ƿ��˿��鿨 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����˷ѽ��׽ӿ�
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2010-02-08 16:40:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strXMLExend As String
    If lng�����ID = 0 Then zlCheckDelValied = True: Exit Function
    
    If gobjSquare.objSquareCard Is Nothing Then
        MsgBox "ע��:" & vbCrLf & _
                     "      ��ǰ�շ��ǰ�" & strName & " �շѵ�,�������ڲ�������ز���,�����˿�,����ϵͳ����Ա��ϵ!", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
        Exit Function
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
    '       strXMLExpend    XML IN  ��ѡ����(��չ��).��δ����
    '����:�˿�Ϸ�,����true,���򷵻�Flase
      If gobjSquare.objSquareCard.zlReturnCheck(Me, mlngModule, lng�����ID, bln���ѿ�, strCardNo, _
        "3|" & lng����ID, dbl�˿���, strSwapNO, strSwapMemo, strXMLExend) = False Then
          zlCheckDelValied = False
          Exit Function
     End If
     
     If bln�Ƿ��˿��鿨 Then
       '����ˢ������
        'zlBrushCard(frmMain As Object, _
        ByVal lngModule As Long, _
        ByVal rsClassMoney As ADODB.Recordset, _
        ByVal lngCardTypeID As Long, _
        ByVal bln���ѿ� As Boolean, _
        ByVal strPatiName As String, ByVal strSex As String, _
        ByVal strOld As String, ByVal dbl��� As Double, _
        Optional ByRef strCardNo As String, _
        Optional ByRef strPassWord As String, _
        Optional ByRef bln�˷� As Boolean = False, _
        Optional ByRef blnShowPatiInfor As Boolean = False, _
        Optional ByRef bln���� As Boolean = False, _
        Optional ByVal bln�����ֹ As Boolean = True, _
        Optional ByRef varSquareBalance As Variant, _
        Optional ByVal blnתԤ�� As Boolean = False, _
        Optional ByVal blnAllPay As Boolean = False, _
        Optional ByVal strXmlIn As String = "") As Boolean
        '       strXmlIn-����������XML���,Ŀǰ��ʽ����:
        '       <IN>
        '           <CZLX>0</CZLX>    //��������,0-��������ˢ��,1-ת�˵���ˢ��,2-�˿����ˢ��
        '       </IN>
        Dim strPassWord As String
        If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModule, Nothing, lng�����ID, _
            bln���ѿ�, mobjChargeInfor.����, mobjChargeInfor.�Ա�, mobjChargeInfor.����, dbl�˿���, strCardNo, strPassWord, _
            True, True, False, True, Nothing, False, False, "<IN><CZLX>2</CZLX></IN>") = False Then Exit Function
    End If
goEnd:
    zlCheckDelValied = True
    Exit Function
End Function

Private Function CallBackBalanceInterface(ByVal lng����ID As Long, ByVal lngԭ����ID As Long, _
    ByVal lng�����ID As Long, _
    ByVal bln���ѿ� As Boolean, _
    ByVal dblMoney As Double, _
    ByVal strCardNo As String, _
    ByVal strSwapNO As String, _
    ByVal strSwapMemo As String, _
    ByRef cllUpdate As Collection, _
    ByRef cllThreeSwap As Collection, ByRef strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���û��˽ӿ�
    '���:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-07-13 10:33:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strSwapExtendInfor As String, strTemp As String
    
    Err = 0: On Error GoTo Errhand:
    
    '�����ID,����,�Ƿ����ѿ�(1-��;0-��),������ˮ��,����˵��,strNO
    If lng�����ID = 0 Then CallBackBalanceInterface = True: Exit Function
    '81489,Ƚ����,2015-1-22,�˷Ѵ������ID
    strSwapExtendInfor = "3|" & lng����ID: strTemp = strSwapExtendInfor
    
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
    '       strSwapExtendInfor-���룬�����˷ѵĳ���ID��
    '                           ��ʽ:�շ�����1|ID1,ID2��IDn||�շ�����n|ID1,ID2��IDn
    '                           �շ�����:1-Ԥ����,2-����,3-�շ�,4-�Һ�,5-ҽ�ƿ��տ�
    '       strSwapExtendInfor-���������׵���չ��Ϣ
    '           ��ʽΪ:��Ŀ����1|��Ŀ����2||��||��Ŀ����n|��Ŀ����n ÿ����Ŀ�в��ܰ���|�ַ�
    If gobjSquare.objSquareCard.zlReturnMoney(Me, mlngModule, lng�����ID, bln���ѿ�, strCardNo, "3|" & lngԭ����ID, dblMoney, strSwapNO, strSwapMemo, strSwapExtendInfor) = False Then Exit Function
    
    Set cllUpdate = New Collection: Set cllThreeSwap = New Collection
    Call zlAddUpdateSwapSQL(False, lng����ID, lng�����ID, bln���ѿ�, strCardNo, strSwapNO, strSwapMemo, cllUpdate)
    If strTemp <> strSwapExtendInfor Then
        Call zlAddThreeSwapSQLToCollection(False, lng����ID, lng�����ID, bln���ѿ�, strCardNo, strSwapExtendInfor, cllThreeSwap)
    End If
    CallBackBalanceInterface = True
Errhand:
End Function

Public Function Getʵ�ս��(ByVal strNo As String) As Double
    Dim i As Long
    On Error GoTo errHandle
    If Not mrsBlance Is Nothing Then
        gstrSQL = "" & _
        "   Select NO,nvl(sum(A.��Ԥ��),0) as ��Ԥ��" & _
        "   From ����Ԥ����¼ A,���㷽ʽ B " & _
        "   Where a.��¼����=3 And A.�������=[1]   " & _
        "               And ( ���㷽ʽ=b.���� and b.���� in (3,4) OR ���㷽ʽ is null ) "
        Set mrsBlance = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mobjChargeInfor.����ID)
    End If
    mrsBlance.Filter = "NO='" & strNo & "'"
    If Not mrsBlance.EOF Then
        Getʵ�ս�� = Val(Nvl(mrsBlance!��Ԥ��))
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub SetControlEnabled()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ÿؼ�������
    '����:���˺�
    '����:2012-02-03 15:08:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnEdit As Boolean
    blnEdit = (mobjChargeInfor.intInsure = 0 Or mobjChargeInfor.intInsure <> 0 And mblnYbBalanced) And mbytFunc <> EM_FUN_����
    blnEdit = blnEdit Or (mbytFunc = EM_FUN_���� And (mblnYbBalanced Or mobjChargeInfor.intInsure = 0))
    picPay.Enabled = blnEdit
    txt��Ԥ��.Enabled = blnEdit
    txt�ɿ�.Enabled = blnEdit
    txt�Ҳ�.Enabled = blnEdit
    txt�������.Enabled = blnEdit
    txtժҪ.Enabled = blnEdit
    
    '������ʾ��ɫ
    txt��Ԥ��.BackColor = IIf(txt��Ԥ��.Enabled, &H80000005, Me.BackColor)
    txt�ɿ�.BackColor = IIf(txt�ɿ�.Enabled, &H80000005, Me.BackColor)
    txt�Ҳ�.BackColor = IIf(txt�Ҳ�.Enabled, &H80000005, Me.BackColor)
    txt�������.BackColor = IIf(txt�������.Enabled, &H80000005, Me.BackColor)
    txtժҪ.BackColor = IIf(txtժҪ.Enabled, &H80000005, Me.BackColor)
End Sub

Public Function Get�շѽ���(ByRef dblԤ��� As Double) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�շѽ�������
    '����:dblԤ���-���ر���֧����Ԥ��
    '����:�շ��ý��㷽ʽ,��ʽ����:
    '       ���㷽ʽ|������|�������|����ժҪ||.....",ע���޽�������ժҪʱҪ�ÿո����
    '����:���˺�
    '����:2012-02-06 10:58:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str���㷽ʽ As String, i As Integer, int���� As Integer
    Dim str�շѽ��� As String
    Dim dblMoney As Double
    
    '���㷽ʽ|������|�������|����ժҪ||.....",ע���޽�������ժҪʱҪ�ÿո����
    '�շ����
    str�շѽ��� = ""
    With vsBlance
        dblԤ��� = Val(txt��Ԥ��.Text)
        
        For i = .Rows - 1 To 1 Step -1
            str���㷽ʽ = Trim(.TextMatrix(i, .ColIndex("֧����ʽ")))
            ' 0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
            int���� = Val(.RowData(i))
            If str���㷽ʽ <> "" And int���� = 0 Then
                '.rowdata:0-��ͨ�Ľ��㷽ʽ-1-ҽ������;2-�����ӿڽ���;3-һ��ͨ����;4-Ԥ���
                str�շѽ��� = str�շѽ��� & "||" & str���㷽ʽ
                str�շѽ��� = str�շѽ��� & "|" & Val(.TextMatrix(i, .ColIndex("֧�����")))
                str�շѽ��� = str�շѽ��� & "|" & IIf(Trim(.TextMatrix(i, .ColIndex("�������"))) = "", " ", Trim(.TextMatrix(i, .ColIndex("�������"))))
                str�շѽ��� = str�շѽ��� & "|" & IIf(Trim(.TextMatrix(i, .ColIndex("��ע"))) = "", " ", Trim(.TextMatrix(i, .ColIndex("��ע"))))
            End If
        Next
        
        If (mCurCardPay.lngҽ�ƿ����ID = 0 Or cbo֧����ʽ.ItemData(cbo֧����ʽ.ListIndex) <> -1) Then
            dblMoney = IIf(mblnYB�˿�, -1, 1) * Val(txt�ɿ�.Text) - IIf(mblnCur����, 0, mCurCarge.dblӦ���ۼ�)
            If mCurCardPay.int���� = 1 Then
                dblMoney = mdbl�ֽ�
            ElseIf mblnCur���� And dblMoney = 0 Then
                dblMoney = mCurCarge.dbl��ǰδ��
            End If
            If dblMoney <> 0 Then
                str�շѽ��� = str�շѽ��� & "||" & mCurCardPay.str���㷽ʽ
                If mCurCardPay.int���� = 1 Then
                    '�ֽ�
                    str�շѽ��� = str�շѽ��� & "|" & dblMoney
                    str�շѽ��� = str�շѽ��� & "| "
                    '77183:������,2014-08-27,�ֽ����ʱû����ժҪ������
                    str�շѽ��� = str�շѽ��� & "| " & IIf(Trim(txtժҪ) = "", " ", Trim(txtժҪ))
                Else
                    str�շѽ��� = str�շѽ��� & "|" & dblMoney
                    str�շѽ��� = str�շѽ��� & "|" & IIf(Trim(txt�������) = "", " ", Trim(txt�������))
                    str�շѽ��� = str�շѽ��� & "|" & IIf(Trim(txtժҪ) = "", " ", Trim(txtժҪ))
                End If
            End If
        End If
    End With
    If str�շѽ��� <> "" Then str�շѽ��� = Mid(str�շѽ���, 3)
    Get�շѽ��� = str�շѽ���
End Function

Private Function zlInsureClinicSwap() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ҽ������
    '���:blnModifyBill-�Ƿ��޸ĵ���
    '       strBalanceIDs:���ν��ʵ�ID,�ֱ��ö��ŷ���
    '       strSaveNos-����ĵ��ݺ�
    '����:strSaveNos-�����Ѿ�����ɹ��ĵ��ݺ�
    '       blnAffair-�Ƿ�������
    '       strSaveSucessNos-����ɹ���Ʊ��(�Ի�����Ч)
    '����:ҽ�����óɹ����ҽ��,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-20 17:15:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varNos As Variant, strSQL As String
    Dim strBillNO As String, blnTrans As Boolean, blnTransMedicare As Boolean
    Dim p As Integer, strAdvance As String
    Dim strTmp As String, i As Long
 
    On Error GoTo errHandle
    If mobjChargeInfor.intInsure = 0 Then zlInsureClinicSwap = True: Exit Function
    blnTrans = True
'    '1. ����Ϊ���۵�
'    If mblnSavePrice Then
'        '����Ϊ���۵�
'        '���������ҽ��,�շ�ȷ��ʱʵ��ȴ����Ϊ���۵�:�����۵���ϸ,����Oracle������ִ��
'        varNos = Split(mobjChargeInfor.Nos, ",")
'        For p = 1 To UBound(varNos)
'            strBillNO = mobjChargeInfor(p)
'            If Not gclsInsure.TranChargeDetail(1, strBillNO, 1, 0, "", , mobjChargeInfor.intInsure) Then
'                'ɾ�����۵�(��������)
'                Call DelMedicareTempNO(True, strBillNO)
'                gcnOracle.RollbackTrans: Exit Function
'            End If
'        Next
'        mblnYbBalanced = True   'ҽ���Ѿ�����
'        zlInsureClinicSwap = True
'        Exit Function
'    End If
      
    If mInsurePara.ҽ���ӿڴ�ӡƱ�� And mobjChargeInfor.ҽ������Ʊ�� = False Then
        '���ϸ����Ʊ��ʱ���浱ǰƱ��
        If Not gblnStrictCtrl Then
            zlDatabase.SetPara "��ǰ�շ�Ʊ�ݺ�", mobjChargeInfor.��ǰ��Ʊ��, glngSys, 1121, zlstr.IsHavePrivs(mstrPrivs, "��������")
        End If
    End If
    
    strAdvance = CStr(-1 * mobjChargeInfor.����ID)
    Dim blnCommit As Boolean '����ִ�гɹ���ȫ��ִ�гɹ������ύ����
    If Not mfrmMain.zlInsureClinicSwap(mobjChargeInfor.����ID, mobjChargeInfor.intInsure, strAdvance, blnCommit) Then
        '�쳣����ʱ������û�������ύ�������أ������շѽ���ֻ��δ�ύ����ǰ��������
        cmdExit.Visible = (cmdExit.Visible And Not blnCommit) Or mbytFunc = EM_FUN_����
        mblnCommitBill = mblnCommitBill Or blnCommit
        gcnOracle.RollbackTrans: Exit Function
    End If
    
    mblnYbBalanced = True   'ҽ���Ѿ�����
    blnTransMedicare = True
    
    If strAdvance = CStr(-1 * mobjChargeInfor.����ID) Then strAdvance = ""
    
    If Not zlInsureCheck(mobjChargeInfor.Ԥ�����, strAdvance) Or strAdvance = "" Then
        '�޸�У�Ա�־
        ' Zl_���������շ�_ҽ������
        strSQL = "Zl_���������շ�_ҽ������("
        '  ����id_In   ������ü�¼.����id%Type,
        strSQL = strSQL & mobjChargeInfor.����ID & ","
        '  �������_In ����Ԥ����¼.�������%Type,
        strSQL = strSQL & "Null,"
        '  ���ս���_In Varchar2
        strSQL = strSQL & "Null)"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        gcnOracle.CommitTrans: blnTrans = False: mblnCommitBill = True
        If Not mInsurePara.�൥�ݷֵ��ݽ��� Then Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicSwap, True, mobjChargeInfor.intInsure)
        zlInsureClinicSwap = True: Exit Function
    End If
    
    Call ҽ�����ݸ���(mobjChargeInfor.����ID, mobjChargeInfor.����ID, strAdvance, False, Nothing)
    '�޸�У�Ա�־
    ' Zl_���������շ�_ҽ������
    strSQL = "Zl_���������շ�_ҽ������("
    '  ����id_In   ������ü�¼.����id%Type,
    strSQL = strSQL & mobjChargeInfor.����ID & ","
    '  �������_In ����Ԥ����¼.�������%Type,
    strSQL = strSQL & "Null,"
    '  ���ս���_In Varchar2
    strSQL = strSQL & "Null)"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    gcnOracle.CommitTrans: blnTrans = False: mblnCommitBill = True
    If Not mInsurePara.�൥�ݷֵ��ݽ��� Then Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicSwap, True, mobjChargeInfor.intInsure)
    zlInsureClinicSwap = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    If blnTrans Then
        'ҽ����HIS����ͬһ������,HIS����ʧ��,��ҽ���������ϴ�,������Ҫ��"ȡ������"�ӿ�
        If blnTransMedicare Then Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicSwap, False, mobjChargeInfor.intInsure)
    End If
'    If blnTransMedicare = False Then    '���ҽ���ɹ��ˣ���ɾ�����۵�������ʧ�ܿ�������
'        Call DelMedicareTempNO(False, strBillNO)
'    End If
    Call SaveErrLog
End Function

Private Sub DelMedicareTempNO(ByVal blnPriceSaved As Boolean, ByVal strBillNO As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ҽ��ֱ���շ�ʱ,ɾ��ǰһ�������ύ�Ļ��۵�
    '����:���˺�
    '����:2014-06-06 18:20:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If Not blnPriceSaved Then Exit Sub
    
    gstrSQL = "zl_���ﻮ�ۼ�¼_DELETE('" & strBillNO & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsBlance_GotFocus()
    ClearԤ����
End Sub

Private Sub ClearԤ����()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���Ԥ����
    '����:���˺�
    '����:2014-08-07 15:22:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not txt��Ԥ��.Enabled Then Exit Sub
    If Not txt��Ԥ��.Visible Then Exit Sub
    If Val(lbl��Ԥ��.Tag) = 1 Then Exit Sub
    If Val(txt��Ԥ��) = 0 Then Exit Sub
    txt��Ԥ��.Text = ""
    txt�ɿ�.Text = ""
End Sub

Private Function GetPayCard(ByVal lngCardTypeID As Long, ByVal bln���ѿ� As Boolean, Optional bln������ As Boolean = True) As Card
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�����ID
    '���:lngCardTypeID-�����ID
    '����:����Card����
    '����:���˺�
    '����:2014-07-31 15:11:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card
    On Error GoTo errHandle
    If Not gobjSquare.objSquareCard Is Nothing Then
        'zlGetCard:(ByVal lngCardTypeID As Long, ByVal bln���ѿ� As Boolean, ByRef objCard As Card)
        If gobjSquare.objSquareCard.zlGetCard(lngCardTypeID, bln���ѿ�, objCard) = False Then Exit Function
        Set GetPayCard = objCard
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
