VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChargePayMentWin 
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
   Icon            =   "frmChargePaymentWin.frx":0000
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
      Left            =   8235
      TabIndex        =   36
      Top             =   255
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
         Begin XtremeSuiteControls.ShortcutCaption ShortcutCaption2 
            Height          =   420
            Left            =   15
            TabIndex        =   29
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
         Begin XtremeSuiteControls.ShortcutCaption ShortcutCaption1 
            Height          =   450
            Left            =   15
            TabIndex        =   27
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
      Height          =   510
      Left            =   0
      TabIndex        =   21
      Top             =   6030
      Width           =   10365
      _ExtentX        =   18283
      _ExtentY        =   900
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   3572
            MinWidth        =   882
            Picture         =   "frmChargePaymentWin.frx":08CA
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7461
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1138
            MinWidth        =   1146
            Object.Tag             =   "�����շ�Ԥ�������ʾ"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1164
            MinWidth        =   1162
            Object.Tag             =   "�����շ�������������ʾ"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   882
            MinWidth        =   882
            Picture         =   "frmChargePaymentWin.frx":115E
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
         ForeColorSel    =   -2147483634
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
         FormatString    =   $"frmChargePaymentWin.frx":1838
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
      Left            =   8220
      TabIndex        =   19
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
Attribute VB_Name = "frmChargePayMentWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Enum PayChargeType
    EM_�����շ� = 0
    EM_�쳣���� = 1
    EM_�����շ� = 2
End Enum
Public Enum ExitMode
    EM_�շ���� = 0
    EM_��ͣ�շ� = 1
    EM_�������� = 2
    EM_�������� = 3
    EM_�˳��շ� = 4
End Enum
Private mbytFunc As PayChargeType  '0-�շ�;1-����
Private mfrmMain As Object
Private mbytReturnMode As ExitMode
Private mbln�쳣���� As Boolean
Private mblnYB�˿� As Boolean 'ҽ������������˵��ݽ�����
'------------------------------------------------------------------------------------------
'���������ر���
Private mlngModule As Long, mstrPrivs As String
Private mintInsure As Integer, mlng����ID As Long
Private mlng����ID As Long, mstr����IDs As String
Private mstr����IDs  As String  'Ŀǰֻ���쳣������Ч
Private mstrNOs As String
Private mstrYBPati As String
Private mstrPatiInfo As String '������Ϣ
Private mlngShareUseID As Long
Private mstrUseType As String 'ʹ�����
Private mblnOK As Boolean
Private mstr���� As String, mstr�Ա� As String, mstr���� As String, mstr�ѱ� As String
Private mbln�������� As Boolean
Private mblnCur���� As Boolean
Private mlngR As Long
Private mlngBrushCardTypeID As Long '����������ˢ���Ŀ����ID,�Ա�ȱʡ��λ�ڸ�֧�������
Private mblnUnloaded  As Boolean
Private mblnLoad As Boolean
'����:42791
Private mstrBalances As String   '��ǰ�Ľ����:���㷽ʽ:���:�ɿ��־(1-�ɿ�;2-�Ҳ�)|���㷽ʽ1:���1:�ɿ��־(1-�ɿ�;2-�Ҳ�)|...
Private mstr��֧Ʊ As String
Private mCurCardPay As gTY_PayMoney '���ο�֧��
Private mdbl����Ӧ�� As Double  '����Ӧ�ɽ��(�������۳�Ԥ����Ǯ)
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
Private mstrҽ������ As String
Private mblnYbBalanced As Boolean 'ҽ���Ѿ�����
Private mblnThreeInterface As Boolean '�Ѿ����������ӿ�
Private mcur������� As Currency
'----------------------------------------------------------------------------------------------
'ҽ�����
'��ǰ���������ҽ��֧�ֲ���
Private Type TYPE_MedicarePAR
    ��������ҽ����Ŀ As Boolean
    �����շѴ�Ϊ���۵� As Boolean
    �����ѽɿ���� As Boolean    '27536
    ������봫����ϸ As Boolean
    ҽ���ӿڴ�ӡƱ�� As Boolean
    ҽ��ȷ���������� As Boolean
    �൥��һ�ν��� As Boolean
    ����������� As Boolean
    ���������շ� As Boolean
    ����Ԥ���� As Boolean
    �൥���շ� As Boolean
    �ֱҴ��� As Boolean
    ʵʱ��� As Boolean
    ���Ը� As Boolean
    ȫ�Ը� As Boolean
    blnOnlyBjYb As Boolean '���ؽ�֧�ֱ���ҽ��:���˺�
    �˷Ѻ��ӡ�ص� As Boolean '
    �൥�ݵ�һ�ν��� As Boolean
End Type
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private mInsurePara As TYPE_MedicarePAR
Private mrsOneCard As ADODB.Recordset
Private mrsBlance As ADODB.Recordset
Private mdbl�ɿ��� As Double, mdbl�Ҳ� As Double
'---------------------------------------------------------------------------------
Private mbln�����շ� As Boolean
'---------------------------------------------------------------------------------
Private mdbl�ֽ� As Double, mdblԭδ�� As Double
Private Const VK_RETURN = &HD
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private mblnCacheKeyReturn As Boolean   '41025:�Ƿ񻺴��˻س���,���ܴ������շѽ���ˢ���б�������˻س�,�����Ҫ�ж�
Public Event zlSaveData(ByRef lng������� As Long, ByRef str����IDs As String, ByRef strSaveNos As String, ByRef blnNotCommit As Boolean, ByRef blnCancel As Boolean)
Private mrsClassMoney As ADODB.Recordset
Private mcllSquareBalance As Collection '���ѿ�������Ϣ
Private mcllCurSquareBalance As Collection '��ǰ���ѿ�ˢ����Ϣ
Private mblnNotChange As Boolean
Private mstrTitle As String '���ڴ�����Ի�����Ĵ�����

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
            grsTotal!���� = 0
            grsTotal!���㷽ʽ = "�ɿ�"
            grsTotal!������ = dbl�ɿ�
        End If
        
        If dbl�Ҳ� <> 0 Then
            grsTotal.Find "���㷽ʽ='" & IIf(mCurCardPay.bln֧Ʊ, "��֧Ʊ", "�Ҳ�") & "'", , adSearchForward, 1
            If grsTotal.EOF Then grsTotal.AddNew
            grsTotal!���� = 1
            grsTotal!���㷽ʽ = IIf(mCurCardPay.bln֧Ʊ, "��֧��", "�Ҳ�")
            grsTotal!������ = dbl�Ҳ�
        End If
        
        For i = 1 To .Rows - 1
            str���㷽ʽ = Trim(.TextMatrix(i, .ColIndex("֧����ʽ")))
            int���� = Val(.RowData(i))
            If str���㷽ʽ <> "" Then
                '.rowdata:0-��ͨ�Ľ��㷽ʽ-1-ҽ������;2-�����ӿڽ���;3-һ��ͨ����;4-Ԥ���
                '����:0-�ɿ�;1-�Ҳ�,2-��Ԥ��;����(mod 10:0-��ͨ����;1-ҽ������;2-������Ʒ;3-һ��ͨ)
                grsTotal.Find "���㷽ʽ='" & str���㷽ʽ & "'", , adSearchForward, 1
                If grsTotal.EOF Then grsTotal.AddNew
                grsTotal!���� = IIf(int���� + 10 = 14, 2, int���� + 10)
                grsTotal!���㷽ʽ = str���㷽ʽ
                grsTotal!������ = Val(Nvl(grsTotal!������)) + Val(.TextMatrix(i, .ColIndex("֧�����")))
                grsTotal.Update
            End If
        Next
        
        If dblԤ�� <> 0 Then
            grsTotal.Find "���㷽ʽ='Ԥ���'", , adSearchForward, 1
            If grsTotal.EOF Then grsTotal.AddNew
            grsTotal!���� = 2
            grsTotal!���㷽ʽ = "Ԥ���"
            grsTotal!������ = Val(Nvl(grsTotal!������)) + dblԤ��
            grsTotal.Update
        End If
        If mCurCardPay.bln���ѿ� Then
            For i = 1 To mcllCurSquareBalance.Count
                '��ǰˢ������(array(�����ID,���ѿ�ID,ˢ�����,����,����,�������,�Ƿ�����))
                grsTotal.Find "���㷽ʽ='" & mCurCardPay.str���㷽ʽ & "'", , adSearchForward, 1
                If grsTotal.EOF Then grsTotal.AddNew
                grsTotal!���� = IIf(mCurCardPay.blnOneCard, 13, 12)
                grsTotal!���㷽ʽ = mCurCardPay.str���㷽ʽ
                grsTotal!������ = Val(Nvl(grsTotal!������)) + Val(mcllCurSquareBalance(i)(2))
                grsTotal.Update
            Next
        Else
            grsTotal.Find "���㷽ʽ='" & mCurCardPay.str���㷽ʽ & "'", , adSearchForward, 1
            If grsTotal.EOF Then grsTotal.AddNew
            ''1-�ֽ���㷽ʽ,2-������ҽ������,3-ҽ�������ʻ�,4-ҽ������ͳ��,5-���տ���,6-�����ۿ�,7-һ��ͨ����,8-���㿨����;<0 ��ʾ������֧��
            '����:0-�ɿ�;1-�Ҳ�,2-��Ԥ��;����(mod 10:0-��ͨ����;1-ҽ������;2-������Ʒ;3-һ��ͨ)
            Select Case mCurCardPay.int����
            Case 1, 2
                grsTotal!���� = 10
            Case 3, 4
                grsTotal!���� = 11
            Case 7, 8
                grsTotal!���� = IIf(mCurCardPay.blnOneCard, 13, 12)
            Case Else
                grsTotal!���� = 10
            End Select
            grsTotal!���㷽ʽ = mCurCardPay.str���㷽ʽ
            grsTotal!������ = Val(Nvl(grsTotal!������)) + dblMoney
            grsTotal.Update
            If dbl��֧Ʊ <> 0 Then
                grsTotal.Find "���㷽ʽ='" & mstr��֧Ʊ & "'", , adSearchForward, 1
                If grsTotal.EOF Then grsTotal.AddNew
                grsTotal!���� = 10
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
    If mintInsure = 0 Then Exit Sub
'    mInsurePara.��������ҽ����Ŀ = gclsInsure.GetCapability(support��������ҽ����Ŀ, mlng����ID, mintInsure)
'    mInsurePara.�����շѴ�Ϊ���۵� = gclsInsure.GetCapability(support�����շѴ�Ϊ���۵�, mlng����ID, mintInsure)
'    mInsurePara.������봫����ϸ = gclsInsure.GetCapability(support������봫����ϸ, mlng����ID, mintInsure)
'    mInsurePara.ҽ��ȷ���������� = gclsInsure.GetCapability(supportҽ��ȷ����������, mlng����ID, mintInsure)
     mInsurePara.ҽ���ӿڴ�ӡƱ�� = gclsInsure.GetCapability(supportҽ���ӿڴ�ӡƱ��, mlng����ID, mintInsure)
    mInsurePara.�൥��һ�ν��� = gclsInsure.GetCapability(support�൥��һ�ν���, mlng����ID, mintInsure)
    mInsurePara.���������շ� = gclsInsure.GetCapability(support���������շ�, mlng����ID, mintInsure)
    '���˺�:27536 20100119
    mInsurePara.�����ѽɿ���� = gclsInsure.GetCapability(support�����ѽɿ����, mlng����ID, mintInsure)
    mInsurePara.����������� = gclsInsure.GetCapability(support�����������, , mintInsure)
    mInsurePara.�൥���շ� = gclsInsure.GetCapability(support�൥���շ�, mlng����ID, mintInsure)
    mInsurePara.����Ԥ���� = gclsInsure.GetCapability(support����Ԥ��, mlng����ID, mintInsure)
    mInsurePara.�ֱҴ��� = gclsInsure.GetCapability(support�ֱҴ���, mlng����ID, mintInsure)
'    mInsurePara.���Ը� = gclsInsure.GetCapability(support�շ��ʻ������Ը�, mlng����ID, mintInsure)
'    mInsurePara.ȫ�Ը� = gclsInsure.GetCapability(support�շ��ʻ�ȫ�Է�, mlng����ID, mintInsure)
'    mInsurePara.ʵʱ��� = gclsInsure.GetCapability(supportʵʱ���, mlng����ID, mintInsure)
    'mInsurePara.�˷Ѻ��ӡ�ص� = gclsInsure.GetCapability(support�˷Ѻ��ӡ�ص�, mlng����ID, mintInsure)
     mInsurePara.�൥�ݵ�һ�ν��� = gclsInsure.GetCapability(support����_���ֵ��ݽ���, mlng����ID, mintInsure)
End Sub
Private Sub InitBalanceData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����������
    '����:���˺�
    '����:2012-02-05 16:02:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call ClearBanalce
    With mCurCarge
          .dbl����ʵ�� = mfrmMain.zlGetToTatal
          .dbl����ҽ��֧�� = mfrmMain.GetMedicareSum
          .dbl�����Ѹ��ϼ� = 0
          .dbl����Ӧ�� = mfrmMain.GetBillSum(True)
          .dbl��ǰδ�� = .dbl����ʵ�� - .dbl����ҽ��֧��
          .dbl���γ�Ԥ�� = 0
          .dbl�������� = 0
      End With
      '����Ԥ����δ������������������бȽϣ�ȷ���Ƿ��ظ�����
      mdblԭδ�� = mCurCarge.dbl��ǰδ��
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
 
    gstrSQL = "" & _
    "   Select  A.ID,Mod(A.��¼����,10) as ��¼����,A.���㷽ʽ,A.��Ԥ��,A.ժҪ,A.�����ID,A.���㿨���, " & _
    "               A.�������,A.����,A.������ˮ��,nvl(C.�Ƿ�����,0) as ���ƿ�, " & _
    "               nvl(C.�Ƿ�����,0) as �Ƿ�����, " & _
    "               nvl(C.�Ƿ�ȫ��,0) as �Ƿ�ȫ��, " & _
    "               decode(C.��������,NULL,0,1) as  �Ƿ�����," & _
    "               C.���� as ���������,A.����˵��,A.�������,A.У�Ա�־,decode(B.����,Null,0,1) as ҽ��,0 as ���ѿ�id" & _
    "   From ����Ԥ����¼ A ,ҽ�ƿ���� C" & IIf(mbln�쳣����, ",Table( f_Num2list( [3])) Q ", "") & _
    "           ,(Select ���� From ���㷽ʽ where ���� in (3,4)) B" & _
    "   Where  " & IIf(mbln�쳣����, "A.����ID=Q.Column_Value", "A.������� = [1] ") & _
    "                And A.�����ID=C.ID(+) And A.���㷽ʽ=B.����(+) And nvl(A.���㿨���,0)=0"
    
 gstrSQL = gstrSQL & " Union ALL " & _
    "   Select   A.ID,Mod(A.��¼����,10) as ��¼����,A.���㷽ʽ,-1*nvl(b.Ӧ�ս��,0) as ��Ԥ��,A.ժҪ,A.�����ID,A.���㿨���, " & _
    "           A.�������,B.����,B.������ˮ��,nvl( M.���ƿ�,0) as ���ƿ�, " & _
    "           nvl( M.�Ƿ�����,0) as �Ƿ�����, " & _
    "           nvl(M.�Ƿ�ȫ��,0) as �Ƿ�ȫ��, " & _
    "           nvl(M.�Ƿ�����,0) as  �Ƿ�����," & _
    "           M.���� as ���������,A.����˵��,A.�������,A.У�Ա�־,0 as ҽ��,B.���ѿ�id" & _
    "   From ����Ԥ����¼ A ,���˿������¼ B, " & _
    "              ���ѿ����Ŀ¼ M" & IIf(mbln�쳣����, ",Table( f_Num2list( [3])) Q ", "") & _
    "   Where  a.Id = b.����id And a.���㿨��� = m.���  " & _
                  IIf(mbln�쳣����, "And A.����ID=Q.Column_Value", " And A.������� = [1] ")
   gstrSQL = "" & _
   "    Select   /*+ rule */    ��¼����,���㷽ʽ,ժҪ,�����ID,���������,���ƿ�,���㿨���,�������,����,������ˮ��, ����˵��,�������,У�Ա�־,ҽ��,���ѿ�id," & _
   "               max(�Ƿ�����) as �Ƿ�����,max(�Ƿ�ȫ��) as �Ƿ�ȫ��,max(�Ƿ�����) as �Ƿ����� , nvl(sum(��Ԥ��),0) as ��Ԥ��" & _
   "    From (" & gstrSQL & ") " & _
   "   Group by ��¼����,���㷽ʽ,ժҪ,�����ID,���������,���ƿ�,���㿨���,�������,����,������ˮ��, ����˵��,�������,У�Ա�־,ҽ��,���ѿ�id"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng����ID, IIf(mbln�쳣����, 2, 1), mstr����IDs)
    With rsTemp
        i = 1
        blnYb = False
        Do While Not .EOF
            If Nvl(rsTemp!ժҪ) = "���ս���" Or Nvl(rsTemp!ҽ��) = "1" Then
                    mCurCarge.dbl����ҽ��֧�� = RoundEx(mCurCarge.dbl����ҽ��֧�� + Nvl(rsTemp!��Ԥ��, 0), 6)
                    blnYb = True
            End If
            If Val(Nvl(rsTemp!У�Ա�־, 0)) = 2 Then
                With vsBlance
                    If .TextMatrix(i, .ColIndex("֧����ʽ")) <> "" Then
                        .Rows = .Rows + 1
                        i = i + 1
                    End If
                    .RowData(i) = 0
                    strCardNo = Nvl(rsTemp!����)
                    lng�����ID = Val(Nvl(rsTemp!���㿨���))
                    bln���ѿ� = lng�����ID <> 0
                    If bln���ѿ� Then
                        If mcllSquareBalance Is Nothing Then Set mcllSquareBalance = New Collection
                        'array(�����ID,���ѿ�ID,ˢ�����,����,����,�������,�Ƿ�����
                        mcllSquareBalance.Add Array(lng�����ID, Val(Nvl(rsTemp!���ѿ�ID)), _
                        Format(Val(Nvl(rsTemp!��Ԥ��)), "0.00"), strCardNo, "", "", Val(Nvl(rsTemp!�Ƿ�����)))
                    End If
                    
                    If Not bln���ѿ� Then lng�����ID = Val(Nvl(rsTemp!�����ID))
                    
                    If lng�����ID <> 0 Then .RowData(i) = 2
                    If lng�����ID <> 0 Then
                        strCardNo = gobjSquare.objSquareCard.zlGetCardNODencode(strCardNo, lng�����ID, bln���ѿ�)
                    End If
                    If Nvl(rsTemp!ժҪ) = "���ս���" Or Val(Nvl(rsTemp!ҽ��)) = 1 Then
                        .RowData(i) = 1 'ҽ������
                        If InStr(1, mstrҽ������, "," & Nvl(rsTemp!���㷽ʽ)) = 0 Then
                            mstrҽ������ = mstrҽ������ & "," & Nvl(rsTemp!���㷽ʽ)
                        End If
                    ElseIf lng�����ID <> 0 Then
                        '�����ӿڽ���
                        .RowData(i) = 2 '�����ӿڽ���
                    Else
                        '�Ƿ�һ��ͨ����
                        mrsOneCard.Filter = "���㷽ʽ='" & Nvl(rsTemp!���㷽ʽ) & "'"
                        If Not mrsOneCard.EOF Then
                            .RowData(i) = 3 'һ��ͨ����
                        End If
                        mrsOneCard.Filter = 0
                    End If
                    
                    .TextMatrix(i, .ColIndex("֧����ʽ")) = Nvl(rsTemp!���㷽ʽ)
                    ' ҽ�ƿ����ID|���ѿ�(1, 0) |���ƿ�|�Ƿ�ȫ��|�Ƿ�����|�ӿ�����
                    .Cell(flexcpData, i, .ColIndex("֧����ʽ")) = lng�����ID & "|" & IIf(bln���ѿ�, 1, 0) & "|" & Val(Nvl(rsTemp!���ƿ�)) & "|" & Val(Nvl(rsTemp!�Ƿ�ȫ��)) & "|" & Val(Nvl(rsTemp!�Ƿ�����)) & "|" & Nvl(rsTemp!���������)
                    
                    .TextMatrix(i, .ColIndex("֧�����")) = Format(Val(Nvl(rsTemp!��Ԥ��)), "0.00")
                    .TextMatrix(i, .ColIndex("�������")) = Nvl(rsTemp!�������)
                    .TextMatrix(i, .ColIndex("��ע")) = Nvl(rsTemp!ժҪ)
                    .TextMatrix(i, .ColIndex("������ˮ��")) = Nvl(rsTemp!������ˮ��)
                    .TextMatrix(i, .ColIndex("����˵��")) = Nvl(rsTemp!����˵��)
                    .TextMatrix(i, .ColIndex("����")) = IIf(Val(Nvl(rsTemp!�Ƿ�����)) = 1, String(Len(strCardNo), "*"), strCardNo)
                    .Cell(flexcpData, i, .ColIndex("����")) = Nvl(rsTemp!����)
      
                    mCurCarge.dbl�����Ѹ��ϼ� = RoundEx(mCurCarge.dbl�����Ѹ��ϼ� + Val(Nvl(rsTemp!��Ԥ��)), 6)
                End With
            ElseIf Val(Nvl(rsTemp!��¼����)) = 1 Or Val(Nvl(rsTemp!��¼����)) = 11 Then
                mCurCarge.dbl���γ�Ԥ�� = RoundEx(mCurCarge.dbl���γ�Ԥ�� + Val(Nvl(rsTemp!��Ԥ��)), 6)
                mCurCarge.dbl�����Ѹ��ϼ� = RoundEx(mCurCarge.dbl�����Ѹ��ϼ� + Val(Nvl(rsTemp!��Ԥ��)), 6)
            End If
            .MoveNext
        Loop
    End With

                    
    If mbln�쳣���� Then
         gstrSQL = "" & _
         "   Select /*+ rule */ B.NO,B.����ID, Nvl(Sum(Nvl(B.Ӧ�ս��, 0)), 0)  As ����Ӧ�պϼ�, " & _
         "       Nvl(Sum(Nvl(B.ʵ�ս��, 0)), 0)  As ����ʵ�պϼ� " & _
         "   From ������ü�¼ B , Table( f_Num2list( [2])) Q  " & _
        "    Where B.����ID=Q.Column_Value " & _
        "    Group by B.NO,B.����ID"
    Else
         gstrSQL = "" & _
         "   Select  /*+ rule */ B.NO,B.����ID, Nvl(Sum(Nvl(B.Ӧ�ս��, 0)), 0)  As ����Ӧ�պϼ�, " & _
         "       Nvl(Sum(Nvl(B.ʵ�ս��, 0)), 0)  As ����ʵ�պϼ� " & _
         "   From ������ü�¼ B  " & _
        "    Where B.����id in (Select ����ID From ����Ԥ����¼ where �������=[1] )  " & _
        "    Group by B.NO,B.����ID"
    End If
   Set mrsBlance = Nothing
   Set mrsBlance = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng����ID, mstr����IDs)
   With mCurCarge
         .dbl����ʵ�� = 0:
         .dbl����Ӧ�� = 0
        Do While Not mrsBlance.EOF
            .dbl����ʵ�� = RoundEx(.dbl����ʵ�� + Val(Nvl(mrsBlance!����ʵ�պϼ�)), 6)
            .dbl����Ӧ�� = RoundEx(.dbl����Ӧ�� + Val(Nvl(mrsBlance!����Ӧ�պϼ�)), 6)
            mrsBlance.MoveNext
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
                .TextMatrix(i, .ColIndex("֧����ʽ")) = "Ԥ���"
                .RowData(i) = 4
                .TextMatrix(i, .ColIndex("֧�����")) = Format(mCurCarge.dbl���γ�Ԥ��, "0.00")
            End With
        End If
        mblnYB�˿� = mCurCarge.dbl��ǰδ�� < 0 And blnYb
   End With
   
   vsBlance_AfterRowColChange 0, 0, vsBlance.Row, vsBlance.Col
   Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Public Function zlChargeWin(ByVal frmMain As Object, ByVal bytFunc As PayChargeType, _
    ByVal lngModule As Long, ByVal strPrivs As String, _
    ByVal lngShareUseID As Long, ByVal strUseType As String, _
    ByVal lng����ID As Long, ByVal str����IDs As String, _
    ByVal strNos As String, _
    ByVal lng����ID As Long, ByVal intInsure As Integer, _
    Optional ByVal str���� As String = "", Optional ByVal str�Ա� As String, _
    Optional str���� As String, Optional str�ѱ� As String = "", _
    Optional dbl�ɿ��� As Double, Optional dbl�Ҳ� As Double, _
    Optional bytReturnMode As ExitMode = EM_�շ����, _
    Optional dblӦ���ۼ� As Double, _
    Optional bln�������� As Boolean, _
    Optional lngBrushCardTypeID As Long = 0, _
    Optional dbl����Ӧ�� As Double = 0, _
    Optional strBalance As String = "", Optional bln�쳣���� As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������:��ʾ����֧�����㴰��
    '���:frmMain-���õ�������
    '       bytFunc-0-�շ�;1-����
    '       lngModule -ģ���
    '       strPrivs-Ȩ�޴�
    '       mlng����ID:�൥�ݽ���ʱ,�Թ����Ľ���IDΪ׼.����Ϊ����Id
    '       strNos-���ݺ�:�Զ��ŷ���,��"AAAA,BBBBB"
    '       dblPayMoney-���������ܶ�
    '       dblYbMoney-ҽ��֧�����
    '       lngBrushCardTypeID-ȱʡ��ˢ�����ID
    '       bln�쳣����-�쳣�������ϴ���(�쳣����ʱ����):���Ϊtrue,��ʾ������ϵ��쳣���ݽ�������
    '����:dbl�ɿ���-����Ľɿ�����Ҳ����(���ֽ�ʱ,����)
    '        bln��������-�Ƿ����¼���Ʊ��
    '        bytReturnMode-���ز���ģʽ(0-�����շ����,1-��ͣ�շ�;2-���������շ�;3-��������)
    '        dbl����Ӧ��-ҽ������,�������շ������,��Ҫ���¼��㱾�ε�Ӧ�ɶ�
    '       strBalance-���ر����շѵĽ��㷽ʽ,��ʽ����:
    '                       ���:�ɿ��־(1-�ɿ�;2-�Ҳ�)|���㷽ʽ1:���1:�ɿ��־(1-�ɿ�;2-�Ҳ�)|...
    '����:����շ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-12 09:59:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Set mrsClassMoney = Nothing
    mblnYbBalanced = False: mblnThreeInterface = False: mblnOK = False
    mlngBrushCardTypeID = lngBrushCardTypeID: Set mfrmMain = frmMain
    mintInsure = intInsure: mlngShareUseID = lngShareUseID: mstrUseType = strUseType
    mlng����ID = lng����ID: mlng����ID = lng����ID: mstrPrivs = strPrivs
    mstr����IDs = "": mstr����IDs = str����IDs: mlngModule = lngModule
    
    mstr���� = str����: mstr�Ա� = str�Ա�: mstr���� = str����: mstr�ѱ� = str�ѱ�
    mstrNOs = strNos: mdbl����Ӧ�� = 0: mbln�쳣���� = bln�쳣����
    mstrPatiInfo = str����
   ' mstrPatiInfo = mstrPatiInfo & "�Ա�:" & str�Ա� & Space(4)
    'mstrPatiInfo = mstrPatiInfo & "����:" & str���� & Space(4)
    'mstrPatiInfo = mstrPatiInfo & "�ѱ�:" & str�ѱ� & Space(4)
    mdbl�ɿ��� = 0: mdbl�Ҳ� = 0: mblnUnLoad = False: mblnUnloaded = False
    mCurCarge.dblӦ���ۼ� = dblӦ���ۼ�
    mbln�������� = dblӦ���ۼ� <> 0
    mstrBalances = ""
    mbytFunc = bytFunc: mbytReturnMode = EM_�շ����
    
    If bln�쳣���� Then
        mstr����IDs = mstr����IDs
        mstr����IDs = zlGetԭ����IDs(mstr����IDs)
    End If
    
    mblnOK = False
    Me.Show 1, frmMain
    bln�������� = mbln��������: dbl����Ӧ�� = mdbl����Ӧ��
    dbl�ɿ��� = mdbl�ɿ���: dbl�Ҳ� = mdbl�Ҳ�
    strBalance = mstrBalances
    bytReturnMode = mbytReturnMode
    zlChargeWin = mblnOK
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
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
    "         Select B.���� " & _
    "         From ���㷽ʽӦ�� A, ���㷽ʽ B " & _
    "         Where A.Ӧ�ó��� = '�շ�' And B.���� = A.���㷽ʽ And Nvl(B.Ӧ����, 0) = 1 And a.���ʽ Is Null And Rownum <= 1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If Not rsTemp.EOF Then
        mstr��֧Ʊ = Nvl(rsTemp!����)
    End If
    Call initInsure
    If mbytFunc = EM_�����շ� Then
        Call InitBalanceData
    Else
        Call LoadData
    End If
    Call Load֧����ʽ: Call LoadPatiInfor
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
    
    sngSplitHeight = 80
    
    '51670
    If mlng����ID = 0 Or mbln�������� Then
        lbl��Ԥ��.Visible = False
        txt��Ԥ��.Visible = False
        txt��Ԥ��.Text = "0"
    End If
    
    cmdNext.Visible = Val(txt��Ԥ��.Text) = 0 And Val(txt�ɿ�.Text) = 0 And mbytFunc = EM_�����շ� And _
        (mCurCarge.dbl�����Ѹ��ϼ� - mCurCarge.dbl����ҽ��֧��) = 0 _
        And (gTy_Module_Para.byt�ɿ���� = 1 Or gTy_Module_Para.byt�ɿ���� = 3) And _
        mCurCardPay.lngҽ�ƿ����ID = 0 And mCurCardPay.blnOneCard = False _
        And (mintInsure = 0 Or mintInsure <> 0 And mblnYbBalanced)
        
    lbl�ѽ�.Caption = "�Ѹ��ϼ�:" & Format(mCurCarge.dbl�����Ѹ��ϼ�, "###0.00;-###0.00;0.00;0.00;")
    
    If mCurCardPay.int���� = 1 And blnԤ�� = False Then
        dblMoney = mCurCarge.dbl��ǰδ�� + mCurCarge.dblӦ���ۼ�
        If mintInsure > 0 Then  '����:43855,44069
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
    Else
        lblPayType.Caption = "�ˡ���"
        lblPayType.ForeColor = vbRed
        cbo֧����ʽ.ForeColor = vbRed
        txt�ɿ�.ForeColor = vbRed
        '�˿�ʱ��������Ԥ��
        txt��Ԥ��.Visible = False: lbl��Ԥ��.Visible = False
        txt��Ԥ��.Text = 0
    End If
    
    If blnԤ�� Then
        'Ԥ���Ĵ���
        lbl�Ҳ�.Visible = False: txt�Ҳ�.Visible = False
        txt�Ҳ�.Text = 0
    ElseIf mCurCardPay.int���� = 1 Then
        lbl�Ҳ�.Visible = True: txt�Ҳ�.Visible = True
        lbl�Ҳ�.Caption = "�ҡ���"
        If IIf(mblnYB�˿� < 0, -1, 1) * Val(txt�ɿ�.Text) >= dbl�ֽ� Then
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
        'IIf(mblnYB�˿�, -1, 1) * (IIf(mblnYB�˿�, -1, 1) * Val(txt�ɿ�.Text) - dblMoney - mCurCarge.dblӦ���ۼ�), "0.00")
        txt�Ҳ�.Text = Format(Val(txt�ɿ�.Text) - dblMoney - mCurCarge.dblӦ���ۼ�, "0.00")
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
    txt�ɿ�.Text = ""
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
    '    Optional ByRef bln���� As Boolean = False, _
    '    Optional ByVal bln�����ֹ As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ָ��֧�����,����ˢ������
    '���:rsClassMoney:�շ����,���
    '        lngCardTypeID-Ϊ��ʱ,Ϊ��һ��ͨˢ��
    '       bln�����ֹ-Ŀǰֻ������ѿ�,��ʾ����ʱ,��ֹ��������,������������֧��
    dblMoney = Val(txt�ɿ�.Text)
    If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModule, Nothing, 0, False, _
    mstr����, mstr�Ա�, mstr����, dblMoney, mCurCardPay.strˢ������, mCurCardPay.strˢ������, _
    False, True, False, False) = False Then Exit Function
 
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
    '������Ӧ��
    If Not mbln�ѱ��� Then Call LedVoiceSpeak
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
            If CSng(txt�Ҳ�.Text) < 0 Then
                MsgBox "�ɿ����,�벹��Ӧ�ɽ�", vbInformation, gstrSysName
                txt�ɿ�.SetFocus: zlControl.TxtSelAll txt�ɿ�
                Exit Function
            End If
        End If
    ElseIf Not mCurCardPay.bln֧Ʊ Then
            '����:42793
            '�������㷽ʽ,����Ľ��ܴ���δ������
            If RoundEx(Val(txt�ɿ�.Text), 2) > RoundEx(mCurCarge.dbl��ǰδ��, 2) Then
                MsgBox "ע��:" & vbCrLf & "    ����Ľɿ��������δ֧���Ľ��,���ܼ���!", vbOKOnly + vbInformation, gstrSysName
                txt�ɿ�.SetFocus: zlControl.TxtSelAll txt�ɿ�
                Exit Function
            End If
    End If

    '��鵱ǰ�����Ƿ�������ִ�����,��Ҫ�ǲ���ԭ����м��
    '��ֹ��������Ա����:
    '45186
    gstrSQL = "" & _
    "   Select  1  From ����Ԥ����¼ A " & _
    "   Where   A.�������=[1] and nvl(A.У�Ա�־,0)<>0 and Rownum =1 and A.��¼״̬=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng����ID)
    If rsTemp.EOF Then
        '�����Ǳ�����ִ��,������Ҫ����Ƿ�����ִ��
        gstrSQL = "Select ��¼״̬, ����Ա����,ִ��״̬ From ������ü�¼ Where ����ID=[1] And rownum=1"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng����ID)
        
        If Not rsTemp.EOF Then
            If Val(Nvl(rsTemp!��¼״̬)) <> 1 Then
                MsgBox "�õ����Ѿ�����������Ա����,�����ٽ����շ�!", vbOKOnly + vbInformation, gstrSysName
                'ִ���շ�
                Unload Me
                Exit Function
            End If
            
            If Val(Nvl(rsTemp!ִ��״̬)) <> 9 Then
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
    
    lngCount = IIf(mintInsure <> 0, 1, 0)   'ҽ����һ������
    If mCurCardPay.lngҽ�ƿ����ID = 0 Then CheckInterfaceNumIsValied = True: Exit Function
    With vsBlance
        strNames = IIf(mintInsure <> 0, vbCrLf & "ҽ������", "")
        For i = 1 To .Rows - 1
            If Val(.RowData(i)) = 2 Or Val(.RowData(i)) = 3 Then
                '�����ӿڻ�һ��ͨ(�ϰ�)
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

Private Function CheckDelValied(ByRef blnExistThreeSwap As Boolean, ByRef blnȫ�� As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����˷�
    '����:blnExistThreeSwap-�Ƿ���������ӿ�
    '        blnȫ��-���������ӿ��Ƿ����ȫ��
    '����:�Ϸ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-25 16:14:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCardNo As String, i As Long
    Dim strSwapNO As String, strSwapMemo As String, varData As Variant
    Dim lng�����ID As Long, bln���ѿ� As Boolean, strTemp As String
    Dim st��������� As String, dblMoney As Double
    blnȫ�� = False: blnExistThreeSwap = False
    With vsBlance
        For i = 1 To .Rows - 1
            dblMoney = Val(.TextMatrix(i, .ColIndex("֧�����")))
            Select Case Val(.RowData(i))
            Case 2 '��������
                ' ҽ�ƿ����ID|���ѿ�(1, 0) |���ƿ�|�Ƿ�ȫ��|�Ƿ�����|�ӿ�����
                strTemp = .Cell(flexcpData, i, .ColIndex("֧����ʽ"))
                If strTemp <> "" Then
                    varData = Split(strTemp & "||||", "|")
                    lng�����ID = Val(varData(0))
                    bln���ѿ� = Val(varData(1)) = 1
                    st��������� = varData(5)
                    strSwapNO = Trim(.TextMatrix(i, .ColIndex("������ˮ��")))
                    strSwapMemo = Trim(.TextMatrix(i, .ColIndex("����˵��")))
                    strCardNo = .Cell(flexcpData, i, .ColIndex("����"))
                    If bln���ѿ� And Val(varData(2)) <> 1 Then
                        blnExistThreeSwap = True
                        blnȫ�� = Val(varData(3)) = 1
                    ElseIf Not bln���ѿ� Then
                        blnExistThreeSwap = True
                        blnȫ�� = Val(varData(3)) = 1
                    End If
                    
                    If zlCheckDelValied(lng�����ID, st���������, bln���ѿ�, strCardNo, strSwapNO, strSwapMemo, mstr����IDs, Val(.TextMatrix(i, .ColIndex("֧�����")))) = False Then
                        Exit Function
                    End If
                End If
             Case 3 'һ��ͨ����
                strCardNo = .Cell(flexcpData, i, .ColIndex("����"))
                 If CheckDelOneCardValied(strCardNo, dblMoney) = False Then Exit Function
                blnExistThreeSwap = True
                blnȫ�� = True
             Case Else
             End Select
        Next
    End With
    CheckDelValied = True
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

Private Sub cbo֧����ʽ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
Private Function Getҽ������ID() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������ҽ������(����ID)
    '����:����IDs
    '����:���˺�
    '����:2012-01-05 19:02:52
    '����:45217
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, str����ID As String
    On Error GoTo errHandle
    If mbln�쳣���� Then
        gstrSQL = "" & _
        "   Select Distinct A.����id " & _
        "   From ������ü�¼ A, ����Ԥ����¼ B, ������ü�¼ D, ����Ԥ����¼ C, " & _
        "           (Select ���� From ���㷽ʽ Where ���� In (3, 4)) U " & _
        "   Where A.����id = B.����id And B.���㷽ʽ = U.���� And Nvl(B.У�Ա�־, 0) = 2 And A.NO = D.NO And " & _
        "         A.��¼���� = D.��¼���� And A.��¼״̬ In (1, 3) And D.����id = C.����id And Nvl(C.У�Ա�־, 0) = 1 And " & _
        "         C.������� = [1]"
    Else
        gstrSQL = "" & _
        "   Select  distinct A.����ID" & _
        "   From ����Ԥ����¼ A,(Select ���� From ���㷽ʽ where ���� in (3,4)) B" & _
        "   Where A.������� = [1]  And A.���㷽ʽ=B.����(+) and nvl(A.У�Ա�־,0)=2"
    End If
   Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng����ID)
    With rsTemp
        Do While Not .EOF
            str����ID = str����ID & "," & Val(Nvl(rsTemp!����ID))
            .MoveNext
        Loop
    End With
    rsTemp.Close
    Set rsTemp = Nothing
    Getҽ������ID = str����ID
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function DelInsureSingle(ByVal blnExistThreeBalance As Boolean, ByRef strSaveCussNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ֵ��ŵ��ݽ���ҽ���˷�
    '���:blnExistThreeBalance-�Ƿ���ڵ���������
    '����:strSaveCussNo-���ʳɹ��ĵ���
    '����:ҽ�����׳ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2012-02-29 15:51:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, varData As Variant
    Dim cllPro As Collection, i As Long, dbl����� As Double
    Dim DateDel As Date, lng����ID As Long, strInvoice As String, strNo As String
    Dim blnCommit As Boolean, blnAffaired As Boolean
    Dim strNos As String, strYB�˷�ID As String
    Dim cllProNO As Collection, lng������� As Long, lng����ID As String, str����IDs As String
    Dim blnCallInsure As Boolean  '�Ƿ�Ҫ��ҽ��
    Dim blnTrans As Boolean
    Dim varBalance  As Variant, j As Long, intPage As Integer
    Dim intPages As Integer, strAdvance As String, lng����ID As Long
    Dim strYB�˷�IDs As String, strSuccesNo As String
    
    
    DateDel = zlDatabase.Currentdate
    If mintInsure = 0 _
        Or (mintInsure <> 0 And (mInsurePara.�൥�ݵ�һ�ν��� Or mInsurePara.�൥��һ�ν���)) _
        Then DelInsureSingle = True: Exit Function
    strYB�˷�ID = Getҽ������ID
    varBalance = Split(mstr����IDs, ",")
    
    '�������ϴ���
    varData = Split(Replace(mstrNOs, "'", ""), ",")
    intPages = UBound(varData) + 1
    For i = UBound(varData) To 0 Step -1
        Set cllPro = New Collection
        strNo = varData(i)
        lng����ID = zlDatabase.GetNextId("���˽��ʼ�¼")
        str����IDs = str����IDs & "," & lng����ID
        If lng������� = 0 Then lng������� = lng����ID
        
        'Zl_�����շѼ�¼_Delete
        strSQL = "zl_�����շѼ�¼_DELETE("
        '  No_In           ������ü�¼.NO%Type,
        strSQL = strSQL & "'" & varData(i) & "',"
        '  ����Ա���_In   ������ü�¼.����Ա���%Type,
        strSQL = strSQL & "'" & UserInfo.��� & "',"
        '  ����Ա����_In   ������ü�¼.����Ա����%Type,
        strSQL = strSQL & "'" & UserInfo.���� & "',"
        '  ҽ�����㷽ʽ_In Varchar2 := Null,
        strSQL = strSQL & "NULL,"
        '  ���_In         Varchar2 := Null,
        strSQL = strSQL & "NULL,"
        '  ���㷽ʽ_In     ����Ԥ����¼.���㷽ʽ%Type := Null,
        strSQL = strSQL & "NULL,"
        '  ���_In         ������ü�¼.ʵ�ս��%Type := 0,
        strSQL = strSQL & "" & dbl����� & ","
        '  �˷�ʱ��_In     ������ü�¼.�Ǽ�ʱ��%Type := Null,
        strSQL = strSQL & "To_Date('" & Format(DateDel, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
        '  ����Ʊ��_In     Number := 0,
        strSQL = strSQL & "1,"
        '  �˷�ժҪ_In     ������ü�¼.ժҪ%Type := Null
        strSQL = strSQL & "'��������',"
        '     У�Ա�־_In: 0-����Ҫ�϶�;1-��϶�(��������Ա�ɿ����,������Ʊ��)
        strSQL = strSQL & "1,"
        '  ����id_In       ����Ԥ����¼.����id%Type := Null,
        strSQL = strSQL & lng����ID & ","
        '  �������_In     ����Ԥ����¼.�������%Type := Null
        strSQL = strSQL & lng������� & ","
          'һ��ͨ����_In   Varchar2 := Null
        strSQL = strSQL & "NULL,"
        '  �˿����_In     Number := 0,
        strSQL = strSQL & "0,"
        '  �൥��ȫ��_In   Number := 0,
        strSQL = strSQL & "0)"
        zlAddArray cllPro, strSQL
        If dbl����� <> 0 Then
            strSQL = "zl_�����շ����_Insert('" & varData(i) & "'," & dbl����� & ",1,0)"
            zlAddArray cllPro, strSQL
        End If
        '�Ȳ���Ʊ�ݣ�ҽ���ӿڲ���ȡ��
        If mInsurePara.ҽ���ӿڴ�ӡƱ�� Then
            strSQL = "zl_�����շѼ�¼_RePrint('" & strNo & "','" & strInvoice & "'," & ZVal(lng����ID) & ",'" & UserInfo.���� & "'," & _
                "To_Date('" & Format(DateDel, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),1,1)"
            zlAddArray cllPro, strSQL
        End If
        '������������
        'strAdvance = ҳ�� & "|��ǰҳ��"
        For j = 0 To UBound(varData)
            If varData(j) = varData(i) Then intPage = j + 1: Exit For
        Next
        '���˺�:ҽ����strAdvancey����:�����˷�������|��ǰ�˷ѵڼ���:27231
        strAdvance = intPages & "|" & intPage
        lng����ID = Val(varBalance(i))
        blnCallInsure = False
        If InStr(1, "," & strYB�˷�ID & ",", "," & lng����ID & ",") > 0 Then
            ' Zl_�������_�϶Ա�־_Update
            strSQL = "Zl_�������_�϶Ա�־_Update("
            '  ����id_In     ������ü�¼.����id%Type,
            strSQL = strSQL & "" & lng����ID & ","
            '  �������id_In ����Ԥ����¼.�������%Type,
            strSQL = strSQL & "NULL,"
            '  �շѽ���_In   Varchar2,
            strSQL = strSQL & "'" & mstrҽ������ & "',"
            '  �����id_In   ����Ԥ����¼.�����id%Type := Null,
            strSQL = strSQL & "NULL,"
            '  ���ѿ�_In     Integer := 0,
            strSQL = strSQL & "0,"
            '  ����_In       ����Ԥ����¼.����%Type := Null,
            strSQL = strSQL & "NULL,"
            '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
            strSQL = strSQL & "NULL,"
            '  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
            strSQL = strSQL & "NULL,"
            '  У�Ա�־_In   ����Ԥ����¼.У�Ա�־%Type := 0
            strSQL = strSQL & "2)"
            zlAddArray cllPro, strSQL
            blnCallInsure = True
         End If
        
        '���ݴ���
        Err = 0: On Error GoTo Errhand:
        blnCommit = False
        gcnOracle.BeginTrans: blnTrans = True
        zlExecuteProcedureArrAy cllPro, Me.Caption, True, True
        If blnCallInsure Then
            '����ҽ���ӿ�
            If Not gclsInsure.ClinicDelSwap(lng����ID, , mintInsure, strAdvance) Then
                gcnOracle.RollbackTrans
                If strSuccesNo <> "" Then strSuccesNo = Mid(strSuccesNo, 2)
                If blnExistThreeBalance Then
                    '���ڵ������ӿ�δ�˷����,��Ҫ���⴦��
                    Call MsgBox("ע��:" & vbCrLf & "    ����Ϊ" & varData(i) & "���շѵ�������ʧ��,��ע�����쳣��������������!" & vbCrLf & _
                                      IIf(strSuccesNo <> "", vbCrLf & "�������µ���ҽ�����ϳɹ�,�����������˷ѻ�δ����:" & strSuccesNo, "") & vbCrLf, vbInformation + vbOKCancel + vbDefaultButton1, gstrSysName)
                Else
                    Call MsgBox("ע��:" & vbCrLf & "    ����Ϊ" & varData(i) & "���շѵ�������ʧ��,��ע�����쳣��������������!" & vbCrLf & _
                                      IIf(strSuccesNo <> "", vbCrLf & "�������µ������ϳɹ�:" & strSuccesNo, "") & vbCrLf, vbInformation + vbOKCancel + vbDefaultButton1, gstrSysName)
                End If
                Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, False, mintInsure)
                Exit Function
            End If
            strSuccesNo = strSuccesNo & "," & strNo
            Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, True, mintInsure)
        End If
        If Not blnCommit Then gcnOracle.CommitTrans
        blnTrans = False
        If Not blnExistThreeBalance Then
            If OverFeeDel(lng����ID, mlng����ID, True) = False Then
                Call MsgBox("ע��:" & vbCrLf & "    ����Ϊ" & varData(i) & "���շѵ���ҽ�����ϳɹ�,��HIS����ʧ��,��ע�����쳣��������������!" & vbCrLf & _
                                  IIf(strSuccesNo <> "", vbCrLf & "�������µ������ϳɹ�:" & strSuccesNo, "") & vbCrLf, vbInformation + vbOKCancel + vbDefaultButton1, gstrSysName)
                Exit Function
            End If
        End If
    Next
    
    If blnExistThreeBalance Then
        '�����������˽���
        blnCommit = True
        If DelSwapThree(str����IDs, lng�������, blnCommit) = False Then
            If Not blnCommit Then gcnOracle.RollbackTrans
            If MsgBox("ע��:" & vbCrLf & "���������Ľ��е����������˷�,�Ƿ���ͣ����?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Function
             Exit Function
        End If
        If Left(str����IDs, 1) = "," Then str����IDs = Mid(str����IDs, 2)
        If OverFeeDel(str����IDs, mlng����ID, blnCommit) = False Then
            If Not blnCommit Then gcnOracle.RollbackTrans
            Exit Function
        End If
    End If
    mbytReturnMode = 2
    DelInsureSingle = True
    Exit Function
Errhand:
    
    gcnOracle.RollbackTrans
ErrInterface:
    Call ErrCenter
    Call SaveErrLog
End Function
Public Function zlGetԭ����IDs(ByVal str�˷�IDs As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݽ�����Ż�ȡ����ID
    '����:�Զ��ŷָ����˷ѵĽ���ID,��:123,23,...
    '����:���˺�
    '����:2012-03-02 10:06:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim str����ID As String, i As Long
    Dim strSQL As String, varData As Variant
    
    On Error GoTo errHandle
    strSQL = "" & _
    "   Select /*+ rule */   Distinct A.����id,A.NO " & _
    "   From   ������ü�¼ A,������ü�¼  B,Table( f_Num2list( [1])) C " & _
    "   Where A.NO=B.NO and A.��¼����=B.��¼���� and A.��¼״̬=3" & _
    "               And B.����ID=C.Column_Value" & _
    "   Order by ����ID desc "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡԭ����ID����", str�˷�IDs)
    varData = Split(Replace(mstrNOs, "'", ""), ",")
    For i = 0 To UBound(varData)
        '������mstrNo��λ��һ��,��Ȼ��ȡ��Ӧ���ݵĽ���IDʱ,�������
        rsTemp.Find " NO='" & varData(i) & "'", , , 1
        If Not rsTemp.EOF Then
            str����ID = str����ID & "," & Val(Nvl(rsTemp!����ID))
        Else
            str����ID = str����ID & "," & "0"
        End If
    Next
    If str����ID <> "" Then str����ID = Mid(str����ID, 2)
    rsTemp.Close
    Set rsTemp = Nothing
    zlGetԭ����IDs = str����ID
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cmdCancelClick()
    Dim strSQL As String, varData As Variant
    Dim cllPro As Collection, i As Long, dbl����� As Double
    Dim DateDel As Date, lng����ID As Long, strInvoice As String, strNo As String
    Dim blnCommit As Boolean, blnAffaired As Boolean
    Dim strNos As String, strYB�˷�ID As String
    Dim cllProNO As Collection, lng������� As Long, lng����ID As String, str����IDs As String
    Dim blnIsExiseThreeSwap As Boolean, blnȫ�� As Boolean
    
    DateDel = zlDatabase.Currentdate
    'һ��ͨ;���������׵ļ��
    If CheckDelValied(blnIsExiseThreeSwap, blnȫ��) = False Then
        If MsgBox("ע��:" & vbCrLf & "���������Ľ��е����������˷�,�Ƿ���ͣ����?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
        Unload Me: Exit Sub
    End If
    If mintInsure <> 0 And mInsurePara.ҽ���ӿڴ�ӡƱ�� Then
        If zlGetInvoiceGroupUseID(lng����ID) = False Then
            If MsgBox("ע��:" & vbCrLf & "    ����ЧƱ��,�Ƿ���ͣ����?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            Unload Me: Exit Sub
        End If
        strInvoice = GetNextBill(lng����ID)
    End If
    mstrBalances = "": mbln�������� = False
    If mintInsure <> 0 And Not mbln�쳣���� And Not (mInsurePara.�൥��һ�ν��� Or mInsurePara.�൥�ݵ�һ�ν���) And (Not blnIsExiseThreeSwap Or blnIsExiseThreeSwap And blnȫ�� = False) Then
        If DelInsureSingle(blnIsExiseThreeSwap, "") = False Then Unload Me: Exit Sub
        Unload Me
        Exit Sub
    End If
    
    If mintInsure <> 0 Then strYB�˷�ID = Getҽ������ID
    '�������ϴ���
    Set cllPro = New Collection
    varData = Split(Replace(mstrNOs, "'", ""), ",")
    If Not mbln�쳣���� Then
        For i = UBound(varData) To 0 Step -1
            strNo = varData(i)
            lng����ID = zlDatabase.GetNextId("���˽��ʼ�¼")
            str����IDs = str����IDs & "," & lng����ID
            If lng������� = 0 Then lng������� = lng����ID
            'Zl_�����շѼ�¼_Delete
            strSQL = "zl_�����շѼ�¼_DELETE("
            '  No_In           ������ü�¼.NO%Type,
            strSQL = strSQL & "'" & varData(i) & "',"
            '  ����Ա���_In   ������ü�¼.����Ա���%Type,
            strSQL = strSQL & "'" & UserInfo.��� & "',"
            '  ����Ա����_In   ������ü�¼.����Ա����%Type,
            strSQL = strSQL & "'" & UserInfo.���� & "',"
            '  ҽ�����㷽ʽ_In Varchar2 := Null,
            strSQL = strSQL & "NULL,"
            '  ���_In         Varchar2 := Null,
            strSQL = strSQL & "NULL,"
            '  ���㷽ʽ_In     ����Ԥ����¼.���㷽ʽ%Type := Null,
            strSQL = strSQL & "NULL,"
            '  ���_In         ������ü�¼.ʵ�ս��%Type := 0,
            strSQL = strSQL & "" & dbl����� & ","
            '  �˷�ʱ��_In     ������ü�¼.�Ǽ�ʱ��%Type := Null,
            strSQL = strSQL & "To_Date('" & Format(DateDel, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
            '  ����Ʊ��_In     Number := 0,
            strSQL = strSQL & "1,"
            '  �˷�ժҪ_In     ������ü�¼.ժҪ%Type := Null
            strSQL = strSQL & "'��������',"
            '     У�Ա�־_In: 0-����Ҫ�϶�;1-��϶�(��������Ա�ɿ����,������Ʊ��)
            strSQL = strSQL & "1,"
            '  ����id_In       ����Ԥ����¼.����id%Type := Null,
            strSQL = strSQL & lng����ID & ","
            '  �������_In     ����Ԥ����¼.�������%Type := Null
            strSQL = strSQL & lng������� & ","
            'һ��ͨ����_In   Varchar2 := Null
            strSQL = strSQL & "NULL,"
            '  �˿����_In     Number := 0,
            strSQL = strSQL & "0,"
            '  �൥��ȫ��_In   Number := 0,
            strSQL = strSQL & "0)"
            zlAddArray cllPro, strSQL
            If dbl����� <> 0 Then
                strSQL = "zl_�����շ����_Insert('" & varData(i) & "'," & dbl����� & ",1,0)"
                zlAddArray cllPro, strSQL
            End If
            '�Ȳ���Ʊ�ݣ�ҽ���ӿڲ���ȡ��
            If mInsurePara.ҽ���ӿڴ�ӡƱ�� Then
                strSQL = "zl_�����շѼ�¼_RePrint('" & strNo & "','" & strInvoice & "'," & ZVal(lng����ID) & ",'" & UserInfo.���� & "'," & _
                    "To_Date('" & Format(DateDel, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),1,1)"
                zlAddArray cllPro, strSQL
            End If
           
        Next
        If mstrҽ������ <> "" Then
            ' Zl_�������_�϶Ա�־_Update
            strSQL = "Zl_�������_�϶Ա�־_Update("
            '  ����id_In     ������ü�¼.����id%Type,
            strSQL = strSQL & "NULL,"
            '  �������id_In ����Ԥ����¼.�������%Type,
            strSQL = strSQL & "" & mlng����ID & ","
            '  �շѽ���_In   Varchar2,
            strSQL = strSQL & "'" & mstrҽ������ & "',"
            '  �����id_In   ����Ԥ����¼.�����id%Type := Null,
            strSQL = strSQL & "NULL,"
            '  ���ѿ�_In     Integer := 0,
            strSQL = strSQL & "0,"
            '  ����_In       ����Ԥ����¼.����%Type := Null,
            strSQL = strSQL & "NULL,"
            '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
            strSQL = strSQL & "NULL,"
            '  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
            strSQL = strSQL & "NULL,"
            '  У�Ա�־_In   ����Ԥ����¼.У�Ա�־%Type := 0
            strSQL = strSQL & "2)"
            zlAddArray cllPro, strSQL
        End If
        'ȫ��
        Err = 0: On Error GoTo Errhand:
        zlExecuteProcedureArrAy cllPro, Me.Caption, True
    Else
        str����IDs = mstr����IDs: lng������� = mlng����ID
        gcnOracle.BeginTrans
    End If
    
    On Error GoTo ErrInterface:
    blnCommit = False
    If mintInsure <> 0 And mstrҽ������ <> "" Then
        If mInsurePara.�൥��һ�ν��� Then
            If DelInsureMulitOneSwap(varData, DateDel, blnCommit) = False Then
                If blnCommit = False Then
                    gcnOracle.RollbackTrans: Exit Sub
                End If
            End If
        ElseIf mInsurePara.�൥�ݵ�һ�ν��� Then
              If DelInsureMulitCallOneInterfrace(varData, blnCommit) = False Then
                    If blnCommit = False Then gcnOracle.RollbackTrans
                    Exit Sub
              End If
        Else
            'ѭ�����ýӿ�
            If InsureCallInterface(varData, strYB�˷�ID, blnCommit) = False Then
                If blnCommit = False Then gcnOracle.RollbackTrans
                Unload Me
                Exit Sub
            End If
        End If
    End If
    'blnAffaired = mstrҽ������ <> ""    '�Ѿ�������������
    '�����������˽���
    If DelSwapThree(str����IDs, lng�������, blnCommit) = False Then
        If Not blnCommit Then gcnOracle.RollbackTrans
        If MsgBox("ע��:" & vbCrLf & "���������Ľ��е����������˷�,�Ƿ���ͣ����?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
        Unload Me: Exit Sub
    End If
    
    If OverFeeDel(str����IDs, mlng����ID, blnCommit) = False Then
        If Not blnCommit Then gcnOracle.RollbackTrans
        Exit Sub
    End If
    mbytReturnMode = 2: mblnOK = True
    Unload Me
    Exit Sub
Errhand:
    gcnOracle.RollbackTrans
ErrInterface:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function OverFeeDel(ByVal str����IDs As String, ByVal lng����ID As Long, ByRef blnCommited As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����˷��շ�
    '���:strNos-����շѵĵ���(����Ϊ����,��Ŀǰֻ��һ�ŵ���)
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-29 14:50:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo errHandle
    If Left(str����IDs, 1) = "," Then str����IDs = Mid(str����IDs, 2)

    ' Zl_�����շѽ���_����˷�
    strSQL = "Zl_�����շѽ���_����˷�("
    '  ����id_In       ������ü�¼.����id%Type,
    strSQL = strSQL & "" & lng����ID & ","
    '  �˷ѽ������_In ����Ԥ����¼.�������%Type,
    strSQL = strSQL & "NULL,"
    '  ����ids_In      Varchar2,
    strSQL = strSQL & "'" & str����IDs & "',"
    '  ����Ա����_In   ����Ԥ����¼.����Ա����%Type := Null
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '������־_In     Integer := 0:
    '0-���½ɿ�����Ԥ�����;1-�����½ɿ�����Ԥ�����,2-��������Ա�ɿ����,ֻ����Ԥ�����
    strSQL = strSQL & "2,"
    '�쳣����_In     Number := 0
    strSQL = strSQL & "1)"
    '�쳣����,����ҲӦ��Ϊ�쳣����
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    If Not blnCommited Then
        gcnOracle.CommitTrans: blnCommited = True
    End If
    OverFeeDel = True
    Exit Function
errHandle:
    If Not blnCommited Then gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
    blnCommited = True
End Function
Private Function DelSwapThree(ByVal str����IDs As String, ByVal lng�˷ѽ������ As String, blnCommit As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�˷ѽ���(һ��ͨ���������㽻��)
    '���:blnCommit -�Ѿ��������������
    '����:���׳ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-25 17:29:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim strCardNo As String, i As Long, strSQL As String, strErrMsg As String
    Dim strSwapNO As String, strSwapMemo As String, varData As Variant
    Dim lng�����ID As Long, bln���ѿ� As Boolean, strTemp As String
    Dim st��������� As String, blnTrans As Boolean, dblMoney As Double
    Dim cllUpdate As Collection, cllThreeSwap As Collection
    Dim strҽԺ���� As String, rsTemp As ADODB.Recordset
    
    gstrSQL = "" & _
    "   Select A.���㷽ʽ,A.ժҪ, " & _
    "             nvl(A.�����ID,nvl(A.���㿨���,0)) as �����ID,Decode(nvl(A.���㿨���,0),0,0,1) as ���ѿ�," & _
    "             A.�������,A.����,A.������ˮ��, " & _
    "             nvl(C.�Ƿ�����,M.���ƿ�) as ���ƿ�, " & _
    "             nvl(C.����,M.����) as ����,A.����˵��,A.�������," & _
    "             Sum(A.��Ԥ��) as ��Ԥ��" & _
    "   From ����Ԥ����¼ A ,ҽ�ƿ���� C,���ѿ����Ŀ¼ M" & _
    "   Where A.�������=[1] And nvl(A.У�Ա�־,0)=1  " & _
    "                And A.�����ID=C.ID(+) and A.���㿨���=M.���(+)   " & _
    "                And nvl(A.�����ID,nvl(A.���㿨���,0))<>0 " & _
    "   Group by A.���㷽ʽ,A.ժҪ,nvl(A.�����ID,nvl(A.���㿨���,0)),Decode(nvl(A.���㿨���,0),0,0,1) ," & _
    "             A.�������,A.����,A.������ˮ��, " & _
    "             nvl(C.�Ƿ�����,M.���ƿ�) , " & _
    "             nvl(C.����,M.����),A.����˵��,A.�������"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng�˷ѽ������)
    Set cllUpdate = New Collection: Set cllThreeSwap = New Collection
    With rsTemp
        Do While Not .EOF
                lng�����ID = Val(Nvl(!�����ID))
                bln���ѿ� = Val(Nvl(!���ѿ�)) = 1
                st��������� = Nvl(!����)
                strSwapNO = Nvl(!������ˮ��)
                strSwapMemo = Nvl(!����˵��)
                strCardNo = Nvl(!����)
                dblMoney = Nvl(!��Ԥ��)
                
               ' Zl_�������_�϶Ա�־_Update
                strSQL = "Zl_�������_�϶Ա�־_Update("
                '  ����id_In     ������ü�¼.����id%Type,
                strSQL = strSQL & "" & "NULL" & ","
                '  �������id_In ����Ԥ����¼.�������%Type,
                strSQL = strSQL & "" & lng�˷ѽ������ & ","
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
                
                '61688
                If blnCommit Then
                    gcnOracle.BeginTrans
                End If
                 blnTrans = True
                zlDatabase.ExecuteProcedure strSQL, Me.Caption
                If CallBackBalanceInterface(str����IDs, lng�����ID, bln���ѿ�, dblMoney, strCardNo, strSwapNO, strSwapMemo, cllUpdate, cllThreeSwap, strErrMsg) = False Then
                    blnCommit = True
                    gcnOracle.RollbackTrans: Exit Function
                End If
                gcnOracle.CommitTrans: blnTrans = False: blnCommit = True
                zlExecuteProcedureArrAy cllUpdate, Me.Caption
                zlExecuteProcedureArrAy cllThreeSwap, Me.Caption
            .MoveNext
        Loop
    End With
    DelSwapThree = True
    Exit Function
errHandle:
    If blnCommit = False Then gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog:     blnCommit = True
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

Private Function InsureCallInterface(ByVal varNos As Variant, ByVal strYB�˷�IDs As String, Optional blnCommited As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����ݺŵ��ýӿ�
    '����:strYB�˷�IDs-����ҽ���ķ���ID,�ö��ŷ���
    '����:���óɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-25 12:21:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnHaveInterface As Boolean, strAdvance As String
    Dim intPages As Integer, j As Integer
    Dim i As Long, intPage As Integer, varBalance As Variant
    Dim lng����ID As Long, blnTrans As Boolean, strSQL As String
    Dim strsuccesNOs As String
    
    'ҽ��Ҫ������һ�ſ�ʼ��
    varBalance = Split(mstr����IDs, ",")
    intPages = UBound(varNos) + 1
    blnTrans = False: strsuccesNOs = ""
    For i = UBound(varNos) To 0 Step -1
        blnTrans = False
        'strAdvance = ҳ�� & "|��ǰҳ��"
        For j = 0 To UBound(varNos)
            If varNos(j) = varNos(i) Then intPage = j + 1: Exit For
        Next
        
        '���˺�:ҽ����strAdvancey����:�����˷�������|��ǰ�˷ѵڼ���:27231
        strAdvance = intPages & "|" & intPage
TORe:
        lng����ID = Val(varBalance(i))
        ' Zl_�������_�϶Ա�־_Update
        strSQL = "Zl_�������_�϶Ա�־_Update("
        '  ����id_In     ������ü�¼.����id%Type,
        strSQL = strSQL & "" & lng����ID & ","
        '  �������id_In ����Ԥ����¼.�������%Type,
        strSQL = strSQL & "NULL,"
        '  �շѽ���_In   Varchar2,
        strSQL = strSQL & "'" & mstrҽ������ & "',"
        '  �����id_In   ����Ԥ����¼.�����id%Type := Null,
        strSQL = strSQL & "NULL,"
        '  ���ѿ�_In     Integer := 0,
        strSQL = strSQL & "0,"
        '  ����_In       ����Ԥ����¼.����%Type := Null,
        strSQL = strSQL & "NULL,"
        '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
        strSQL = strSQL & "NULL,"
        '  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
        strSQL = strSQL & "NULL,"
        '  У�Ա�־_In   ����Ԥ����¼.У�Ա�־%Type := 0
        strSQL = strSQL & "2)"
        If InStr(1, "," & strYB�˷�IDs & ",", "," & lng����ID & ",") > 0 Then
            If blnCommited Then gcnOracle.BeginTrans
            blnTrans = True
            zlDatabase.ExecuteProcedure strSQL, Me.Caption
            If Not gclsInsure.ClinicDelSwap(lng����ID, , mintInsure, strAdvance) Then
                gcnOracle.RollbackTrans: blnCommited = True
                Call MsgBox("ע��:" & vbCrLf & "    ����Ϊ" & varNos(i) & " ���շѵ��ݽ���ҽ���˷�ʱʧ��,������ڡ��쳣���ݡ������½������ϴ������ϵͳ����Ա��ϵ!" & vbCrLf & _
                                   IIf(strsuccesNOs <> "", "����Ϊҽ���Ѿ��˷ѳɹ��ĵ���:" & vbCrLf & strsuccesNOs, "") & vbCrLf & _
                                   "" & vbCrLf, vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName)
                      '  GoTo TORe:
                 Exit Function
            End If
            gcnOracle.CommitTrans: blnCommited = True
            strsuccesNOs = strsuccesNOs & "," & varNos(i)
            Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, True, mintInsure)
            blnHaveInterface = True
        End If
    Next
    InsureCallInterface = True
End Function

Private Function DelInsureMulitCallOneInterfrace(ByVal varNos As Variant, ByRef blnCommit As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�൥�ݵ���һ�νӿ�
    '����:�൥�ݵ���һ�νӿڳɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-25 12:17:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strAdvance As String, varBalance As Variant, lng����ID As Long
    Dim strSQL As String, blnTransMedicare As Boolean
    Dim dbl������ As Double, dbl�ɷ���� As Double, dbl�˿�ϼ� As Double, dbl��� As Double
    Dim dbl����� As Double
    Dim str���㷽ʽ As String, strBalance As String
    Dim arrData As Variant, blnTrans As Boolean
    Dim cllPro As Collection, rsTmp As ADODB.Recordset
    Dim k As Long, j As Long, i As Long
    
    blnCommit = False
    If mInsurePara.�൥�ݵ�һ�ν��� = False Then DelInsureMulitCallOneInterfrace = True: Exit Function
    On Error GoTo errHandle
    varBalance = Split(mstr����IDs, ",")
    strAdvance = mstr����IDs
    lng����ID = Val(varBalance(UBound(varBalance)))
    blnTransMedicare = False
    If Not gclsInsure.ClinicDelSwap(lng����ID, , mintInsure, strAdvance) Then
         blnCommit = True
         gcnOracle.RollbackTrans
         Exit Function
    End If
    blnTransMedicare = True
    If strAdvance = mstr����IDs Or strAdvance = "" Then
        Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, True, mintInsure)
        If Not blnCommit Then gcnOracle.CommitTrans: blnCommit = True
        DelInsureMulitCallOneInterfrace = True
        Exit Function
    End If
    '���ݷ��صĽ�����Ϣ������Ԥ����¼��strAdvance���ظ�ʽ:���㷽ʽ1|���||���㷽ʽ2:���...
    '�ȷ�̯��ÿ�ŵ�����
    '1.��̯��ҽ��
    Set mrsBlance = Nothing
    Set rsTmp = GetBalanceSet
    varBalance = Split(strAdvance, "||")
    For i = 0 To UBound(varBalance)
        str���㷽ʽ = Split(varBalance(i), "|")(0)
        dbl������ = -1 * Val(Split(varBalance(i), "|")(1))
        For k = 0 To UBound(varNos)
            dbl�ɷ���� = Getʵ�ս��(varNos(k))
            rsTmp.Filter = "�������=" & k
            For j = 1 To rsTmp.RecordCount
                dbl�ɷ���� = dbl�ɷ���� - rsTmp!������
                rsTmp.MoveNext
            Next
            If dbl�ɷ���� > 0 Then
                If dbl�ɷ���� <= dbl������ Then
                    dbl������ = dbl������ - dbl�ɷ����
                Else
                    dbl�ɷ���� = dbl������
                    dbl������ = 0
                End If
                rsTmp.AddNew
                rsTmp!������� = k
                rsTmp!���㷽ʽ = str���㷽ʽ
                If k = UBound(varNos) Then  'δ��̯���,�������һ�ŵ�����
                    rsTmp!������ = dbl�ɷ���� + dbl������
                End If
                rsTmp.Update
                If dbl������ = 0 Then Exit For
            ElseIf k = UBound(varNos) Then  'δ��̯���,�������һ�ŵ�����
                rsTmp.AddNew
                rsTmp!������� = k
                rsTmp!���㷽ʽ = str���㷽ʽ
                rsTmp!������ = dbl������
                rsTmp.Update
            End If
        Next
    Next
    For k = 0 To UBound(varNos)
        strBalance = ""
        dbl����� = 0
        dbl��� = Getʵ�ս��(varNos(k))
        rsTmp.Filter = "�������=" & k
        For i = 1 To rsTmp.RecordCount
            strBalance = IIf(strBalance = "", "", strBalance & "||") & rsTmp!���㷽ʽ & "|" & -1 * rsTmp!������
            dbl��� = dbl��� - rsTmp!������
            rsTmp.MoveNext
        Next
        dbl�˿�ϼ� = dbl�˿�ϼ� + dbl���
        lng����ID = GetDelBalanceID(varNos(k))
        'Zl_ҽ������У��_Update
        strSQL = "Zl_ҽ������У��_Update("
        '  ����id_In   ������ü�¼.����id%Type,
        strSQL = strSQL & lng����ID & ","
        '  ���ս���_In Varchar2
        strSQL = strSQL & strBalance & ")"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    Next
    gcnOracle.CommitTrans: blnTrans = False
    blnCommit = True
    Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, True, mintInsure)

    MsgBox "Ӧ�˽��" & vbCrLf & zlstr.NeedName(cbo֧����ʽ.Text) & "��" & Format(dbl�˿�ϼ�, "0.00") & "Ԫ", vbInformation + vbOKOnly, gstrSysName
    
    DelInsureMulitCallOneInterfrace = True
    Exit Function
errHandle:
    '����:50134
    If blnTrans Then gcnOracle.RollbackTrans
    If blnTransMedicare Then Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, False, mintInsure)
    Call ErrCenter
End Function

Private Function DelInsureMulitOneSwap(ByVal varNos As Variant, _
    ByVal dtDate As Date, Optional blnCommit As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�൥��һ�ν���
    '����:blnCommit-�Ƿ��Ѿ��ύ
    '����:�൥��һ�ν����Ƕ൥��һ�ν���,�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-25 10:45:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strAdvance As String, varBalance As Variant, lng����ID As Long
    Dim strSQL As String, blnTransMedicare As Boolean
    Dim dbl������ As Double, dbl�ɷ���� As Double, dbl�˿�ϼ� As Double, dbl��� As Double
    Dim dbl����� As Double
    Dim str���㷽ʽ As String, strBalance As String
    Dim arrData As Variant, blnTrans As Boolean
    Dim cllPro As Collection, rsTmp As ADODB.Recordset
    Dim k As Long, j As Long, i As Long
    
    blnTrans = True
    blnCommit = False
    On Error GoTo errHandle
    If mintInsure = 0 Then DelInsureMulitOneSwap = True: Exit Function
    If Not mInsurePara.�൥��һ�ν��� Then DelInsureMulitOneSwap = True: Exit Function
    
    varBalance = Split(mstr����IDs, ",")
    strAdvance = mstr����IDs
    lng����ID = Val(varBalance(UBound(varBalance)))
    blnTransMedicare = False
    If Not gclsInsure.ClinicDelSwap(lng����ID, , mintInsure, strAdvance) Then
         blnCommit = True
         gcnOracle.RollbackTrans
         Exit Function
    End If
    blnTransMedicare = True
    If strAdvance = mstr����IDs Or strAdvance = "" Then
        Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, True, mintInsure)
        gcnOracle.CommitTrans: blnCommit = True
        DelInsureMulitOneSwap = True
        Exit Function
    End If
    '���ݷ��صĽ�����Ϣ������Ԥ����¼��strAdvance���ظ�ʽ:���㷽ʽ1|���||���㷽ʽ2:���...
    '�ȷ�̯��ÿ�ŵ�����
    '1.��̯��ҽ��
    Set mrsBlance = Nothing
    Set rsTmp = GetBalanceSet
    varBalance = Split(strAdvance, "||")
    For i = 0 To UBound(varBalance)
        str���㷽ʽ = Split(varBalance(i), "|")(0)
        dbl������ = -1 * Val(Split(varBalance(i), "|")(1))
        For k = 0 To UBound(varNos)
            dbl�ɷ���� = Getʵ�ս��(varNos(k))
            rsTmp.Filter = "�������=" & k
            For j = 1 To rsTmp.RecordCount
                dbl�ɷ���� = dbl�ɷ���� - rsTmp!������
                rsTmp.MoveNext
            Next
            If dbl�ɷ���� > 0 Then
                If dbl�ɷ���� <= dbl������ Then
                    dbl������ = dbl������ - dbl�ɷ����
                Else
                    dbl�ɷ���� = dbl������
                    dbl������ = 0
                End If
                rsTmp.AddNew
                rsTmp!������� = k
                rsTmp!���㷽ʽ = str���㷽ʽ
                If k = UBound(varNos) Then  'δ��̯���,�������һ�ŵ�����
                    rsTmp!������ = dbl�ɷ���� + dbl������
                End If
                rsTmp.Update
                If dbl������ = 0 Then Exit For
            ElseIf k = UBound(varNos) Then  'δ��̯���,�������һ�ŵ�����
                rsTmp.AddNew
                rsTmp!������� = k
                rsTmp!���㷽ʽ = str���㷽ʽ
                rsTmp!������ = dbl������
                rsTmp.Update
            End If
        Next
    Next
    For k = 0 To UBound(varNos)
        strBalance = ""
        dbl����� = 0
        dbl��� = Getʵ�ս��(varNos(k))
        rsTmp.Filter = "�������=" & k
        For i = 1 To rsTmp.RecordCount
            strBalance = IIf(strBalance = "", "", strBalance & "||") & rsTmp!���㷽ʽ & "|" & -1 * rsTmp!������
            dbl��� = dbl��� - rsTmp!������
            rsTmp.MoveNext
        Next
        dbl�˿�ϼ� = dbl�˿�ϼ� + dbl���
        lng����ID = GetDelBalanceID(varNos(k))
        'Zl_ҽ������У��_Update
        strSQL = "Zl_ҽ������У��_Update("
        '  ����id_In   ������ü�¼.����id%Type,
        strSQL = strSQL & lng����ID & ","
        '  ���ս���_In Varchar2
        strSQL = strSQL & strBalance & ")"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    Next
    gcnOracle.CommitTrans: blnTrans = False
    blnCommit = True
    Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, True, mintInsure)

    MsgBox "Ӧ�˽��" & vbCrLf & zlstr.NeedName(cbo֧����ʽ.Text) & "��" & Format(dbl�˿�ϼ�, "0.00") & "Ԫ", vbInformation + vbOKOnly, gstrSysName
    
    DelInsureMulitOneSwap = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    If blnTrans Then gcnOracle.RollbackTrans
    If blnTransMedicare Then Call gclsInsure.BusinessAffirm(����Enum.Busi_ClinicDelSwap, False, mintInsure)

End Function
 

Private Sub cmdDel_Click()
    Dim dblMoney As Double, strSQL As String
    Dim byt�������� As Byte
    Dim str���㷽ʽ As String
    If mbytFunc = EM_�쳣���� Then Exit Sub
    'ɾ����صķ���
    With vsBlance
        If .Row < 0 Then Exit Sub
        '.rowdata:0-��ͨ�Ľ��㷽ʽ-1-ҽ������;2-�����ӿڽ���;3-һ��ͨ����;4-Ԥ���
        Select Case Val(.RowData(.Row))
        Case 1, 2, 3    '1-ҽ������;2-�����ӿڽ���;3-һ��ͨ����
            '����ֱ��ɾ��
            Exit Sub
        Case 4  'Ԥ���
            byt�������� = 1
            str���㷽ʽ = ""
        Case 0  '��ͨ�Ľ��㷽ʽ
            byt�������� = 0
            str���㷽ʽ = .TextMatrix(.Row, .ColIndex("֧����ʽ"))
        Case Else
            Exit Sub
        End Select
        dblMoney = Val(.TextMatrix(.Row, .ColIndex("֧�����")))
        If Not (byt�������� = 0 Or byt�������� = 1) Then
            '�����д���
            'Zl_�����շѽ���_Ԥ��_Del
            strSQL = " Zl_�����շѽ���_Ԥ��_Del("
            '  ��������_In   Number,0-�����շ�;1-��Ԥ��
            strSQL = strSQL & "" & byt�������� & ","
            '  �������id_In ����Ԥ����¼.�������%Type,
            strSQL = strSQL & "" & mlng����ID & ","
            '  ���㷽ʽ_In   Varchar2,
            strSQL = strSQL & "" & IIf(str���㷽ʽ = "", "NULL", "'" & str���㷽ʽ & "'") & ","
            '  ������_In   ����Ԥ����¼.��Ԥ��%Type
            strSQL = strSQL & dblMoney & ")"
            Err = 0: On Error GoTo Errhand:
            zlDatabase.ExecuteProcedure strSQL, Me.Caption
        End If
        mCurCarge.dbl��ǰδ�� = RoundEx(mCurCarge.dbl��ǰδ�� + dblMoney, 6)
        mCurCarge.dbl�����Ѹ��ϼ� = RoundEx(mCurCarge.dbl�����Ѹ��ϼ� - dblMoney, 6)
        Call SetControlProperty
        If Val(.RowData(.Row)) = 4 Then
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
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub cmdExit_Click()
    mblnOK = False: mbytReturnMode = EM_�˳��շ�
    Unload Me
End Sub

Private Sub cmdNext_Click()
    '������һ�ŵ��ݵ�¼��
    '�����ϴ�֧����ʽ
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
    If SaveCharge = False Then GoTo GoOver
    mbln�������� = True
    mbytReturnMode = 3
GoOver:
    mstrBalances = ""
    mblnCur���� = False
End Sub

Private Sub cmdOK_Click()
    '�������
    If mbytFunc = EM_�쳣���� Or mbytFunc = EM_�����շ� Then
        If zlIsCheckExistErrBill(mlng����ID) = False Then
            MsgBox "��ǰ�쳣�����ѱ������㲻�ܼ�����", vbInformation, gstrSysName
            Exit Sub
        End If
        If zlCheckOtherSessionDoing(mlng����ID) Then
            MsgBox "��ǰ�������������շѴ����н��д����㲻�ܼ�����", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
   If mbytFunc = EM_�쳣���� Then
     Call cmdCancelClick
     Exit Sub
   End If
   '���ݽ��水�˻س���
   If mblnCacheKeyReturn Then mblnCacheKeyReturn = False: Exit Sub
    '�ȴ���Ԥ��
    mbln�������� = False
    If BrushcardStrikePrepay = False Then Exit Sub
    '�ٴ�������
    If isValied = False Then Exit Sub
    If txt�ɿ�.Text <> "0.00" Then
        'LED��ʾ
        Call ShowLedInfor
    End If
    If SaveCharge = False Then Exit Sub
End Sub
Private Sub SetCtrlVisible()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ÿؼ�����ʾ״̬
    '����:���˺�
    '����:2012-02-03 13:58:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnTemp As Boolean
    If mbytFunc = EM_�����շ� Then
        'ҽ����ҽ��δ���н���ʱ,����ʾ
        cmdYBBalance.Visible = mintInsure <> 0 And mblnYbBalanced = False
        'ҽ�����н����˵�,���ҽ����,��ʾ����շ�
        cmdOK.Visible = (mintInsure = 0 Or mintInsure <> 0 And mblnYbBalanced)
        'ҽ�������˽����,�����˳�
        cmdExit.Visible = mintInsure = 0 And Not mblnThreeInterface Or mintInsure <> 0 And mblnYbBalanced = False
        '�����շ�
        blnTemp = gTy_Module_Para.byt�ɿ���� = 1 Or gTy_Module_Para.byt�ɿ���� = 3 '�Ƿ���������շ�
        '��ͨ�շѻ�ҽ���Ѿ�����
        blnTemp = blnTemp And (mintInsure = 0 Or mintInsure <> 0 And mblnYbBalanced)
        blnTemp = blnTemp And Val(txt��Ԥ��.Text) = 0 'δ��Ԥ�����
        cmdNext.Visible = blnTemp
        If (gTy_Module_Para.byt�ɿ���� = 1 Or gTy_Module_Para.byt�ɿ���� = 3) And mbln�������� Then
            cbo֧����ʽ.Locked = True
        End If
        Exit Sub
     End If
     
     If mbytFunc = EM_�����շ� Then
        cmdExit.Caption = "�˳�(&E)"
        cmdOK.Visible = True: cmdYBBalance.Visible = False
        mblnYbBalanced = mintInsure <> 0    'ҽ������ʱ,�쳣����һ�㶼�ǽ����˵�.
        cmdExit.Visible = True: cmdNext.Visible = False
     End If
     If mbytFunc = EM_�쳣���� Then
        cmdOK.Caption = "���Ͻ���(&O)"
        cmdExit.Caption = "�˳�(&E)"
        cmdOK.Visible = True: cmdYBBalance.Visible = False
        cmdExit.Visible = True: cmdNext.Visible = False
     End If
End Sub
Private Function SaveBill(Optional blnNotCommit As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���浥�ݴ���
    '���:�����������ύ(��Ҫ�Ǵ�����ͨ�����շѣ��������һ�������н��д��������쳣���ݵĳ���)
    '����:����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2012-02-05 16:50:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
   Dim blnCancel As Boolean, strNos As String
    '���ݱ���
    RaiseEvent zlSaveData(mlng����ID, mstr����IDs, strNos, blnNotCommit, blnCancel)
    mstrNOs = strNos
    If blnCancel Then Exit Function
    SaveBill = True
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
    Dim strTemp As String, dblMoney As Double
    Dim lng����ID As Long
    On Error GoTo errHandle
    If mbytFunc = EM_�����շ� Then ҽ������϶� = True: Exit Function
    If mintInsure = 0 Then ҽ������϶� = True: Exit Function
    If mstr����IDs = "" Then Exit Function
    
    '0-����;1-��У��;2-���У��;3-���ӣ�ָ���ص�������ҽ��֧���ĸ��ֽ��㷽ʽ
    gstrSQL = "" & _
    "   Select /*+ rule */ A.��¼ID,A.У��  " & _
    "   From ���ս����¼ A,Table( f_Num2list([1]))  B " & _
    "   Where A.��¼ID=B.Column_Value And nvl(A.У��,0)=1 "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstr����IDs)
    If rsTemp.EOF Then ҽ������϶� = True: Exit Function
    '���ҽ���˶Ա��޼�¼���˳�
    'Select ����ID,���㷽ʽ,��� From ���ս�����ϸ Where ��־=1
    gstrSQL = "" & _
    "   Select /*+ rule */  A.����ID,a.���㷽ʽ,a.���" & _
    "    From ���ս�����ϸ A,Table( f_Num2list([1])) B ,���㷽ʽ C" & _
    "   Where A.����id =B.Column_Value and A.��־=1 and A.���㷽ʽ=C.���� And C.���� in (3,4) " & _
    "   Order by A.���㷽ʽ"
    'ҽ���ܿصĹ��̶̹�д����һ��"�ֽ�",�����ſ���ҽ����Ľ��㷽ʽ
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���ս������", mstr����IDs)
        'δ�к˶�����,ֱ�ӷ���
    If rsTemp.RecordCount = 0 Then ҽ������϶� = True: Exit Function
    
    strҽ������ = ""   '���㷽ʽ|������||
    strShowMsg = ""
    strTemp = "": dblMoney = 0
    For i = 1 To rsTemp.RecordCount
        If strTemp <> Nvl(rsTemp!���㷽ʽ, " ") Then
            If strTemp <> "" And dblMoney <> 0 Then
                 strҽ������ = strҽ������ & "||" & strTemp & "|" & dblMoney
                 strShowMsg = strShowMsg & vbCrLf & strTemp & ":" & dblMoney
            End If
            strTemp = Nvl(rsTemp!���㷽ʽ, " ")
            dblMoney = 0
        End If
        dblMoney = dblMoney + Val(Nvl(rsTemp!���))
        rsTemp.MoveNext
    Next
    If strTemp <> "" And dblMoney <> 0 Then
         strҽ������ = strҽ������ & "||" & strTemp & "|" & dblMoney
        strShowMsg = strShowMsg & vbCrLf & strTemp & ":" & dblMoney
    End If
    If strҽ������ <> "" Then strҽ������ = Mid(strҽ������, 3)
    MsgBox "ע��:" & vbCrLf & "  �ڽ���ҽ������ʱ,ҽ��Ԥ��������ʽ���㲻һ��,��У�Ա��ս�������,����Ϊ��ȷ�Ľ�������:" & vbCrLf & strShowMsg, vbInformation + vbOKOnly, gstrSysName
    If mInsurePara.�൥�ݵ�һ�ν��� Or mInsurePara.�൥��һ�ν��� Then
        'Zl_���������շ�_ҽ������
        gstrSQL = "Zl_���������շ�_ҽ������("
        '����id_In   ������ü�¼.����id%Type,
        gstrSQL = gstrSQL & "NULL,"
        '�������_In ����Ԥ����¼.�������%Type,
        gstrSQL = gstrSQL & "" & mlng����ID & ","
        '���ս���_In Varchar2
        gstrSQL = gstrSQL & "'" & strҽ������ & "')"
        Err = 0: On Error GoTo ErrCommit:
        zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
        '���¼�������
        Call LoadData
        Call LoadPatiInfor
        Call SetControlProperty
        ҽ������϶� = True
        Exit Function
    End If
    lng����ID = 0: strTemp = ""
    '��������Ч��
    rsTemp.Sort = "����ID"
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    With rsTemp
        strҽ������ = ""
        Do While Not .EOF
            If lng����ID <> Val(Nvl(!����ID)) Then
                If lng����ID <> 0 Then
                    strҽ������ = Mid(strҽ������, 3)
                    '�϶�����
                    'Zl_���������շ�_ҽ������
                    gstrSQL = "Zl_���������շ�_ҽ������("
                    '����id_In   ������ü�¼.����id%Type,
                    gstrSQL = gstrSQL & "" & lng����ID & ","
                    '�������_In ����Ԥ����¼.�������%Type,
                    gstrSQL = gstrSQL & "NULL,"
                    '���ս���_In Varchar2
                    gstrSQL = gstrSQL & "'" & strҽ������ & "')"
                    Err = 0: On Error GoTo ErrCommit:
                    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
                End If
                lng����ID = Val(Nvl(!����ID))
                strҽ������ = ""
            End If
            strTemp = Trim(Nvl(rsTemp!���㷽ʽ, " "))
            If strTemp <> "" Then
                strҽ������ = strҽ������ & "||" & strTemp & "|" & Val(Nvl(rsTemp!���))
            End If
            .MoveNext
        Loop
        If strҽ������ <> "" And lng����ID <> 0 Then
            strҽ������ = Mid(strҽ������, 3)
            '�϶�����
            'Zl_���������շ�_ҽ������
            gstrSQL = "Zl_���������շ�_ҽ������("
            '����id_In   ������ü�¼.����id%Type,
            gstrSQL = gstrSQL & "" & lng����ID & ","
            '�������_In ����Ԥ����¼.�������%Type,
            gstrSQL = gstrSQL & "NULL,"
            '���ս���_In Varchar2
            gstrSQL = gstrSQL & "'" & strҽ������ & "')"
            Err = 0: On Error GoTo ErrCommit:
            zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
        End If
    End With
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
    
    '�������
    If mbytFunc = EM_�쳣���� Or mbytFunc = EM_�����շ� Then
        If zlIsCheckExistErrBill(mlng����ID) = False Then
            MsgBox "��ǰ�쳣�����ѱ������㲻�ܼ�����", vbInformation, gstrSysName
            Exit Sub
        End If
        If zlCheckOtherSessionDoing(mlng����ID) Then
            MsgBox "��ǰ�������������շѴ����н��д����㲻�ܼ�����", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    '���ݱ���
    If SaveBill = False Then Exit Sub
    mblnYbBalanced = True   'ҽ���Ѿ�����
    Call LoadData
    'ҽ��:58344
    mblnYB�˿� = mCurCarge.dbl��ǰδ�� < 0
    
    Call LoadPatiInfor
    Call SetControlProperty
    '���ҽ������,��Ҫ�������ð�ť
    Call SetCtrlVisible
    Call SetControlEnabled
    '��궨λ
    '����ʹ��Ԥ��
    If txt��Ԥ��.Visible And txt��Ԥ��.Enabled And gblnPrePayPriority Then
        txt��Ԥ��.SetFocus
        Call SetControlProperty(True): mbln�ѱ��� = True
        Call Show�����(True)
    Else
        mblnNotChange = True
        txt��Ԥ��.Text = ""
        mblnNotChange = False
        '70430,Ƚ����,2014-4-24,�ڽ���Ԥ����ʱ��ʾ�ɿ������ҽ������ʱ�ٴ���ʾ��ͬ�ɿ������ظ���ʾ��
        If txt�ɿ�.Enabled And txt�ɿ�.Visible Then
            mbln�ѱ��� = True '�������ѱ���Ϊtrue,����txt�ɿ��ý��������
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
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
    Call cbo֧����ʽ_Click
    Call SetControlProperty
    Call SetCtrlVisible
    Call SetControlEnabled
    If ҽ������϶� = False Then Unload Me: Exit Sub
    If txt��Ԥ��.Visible Then txt��Ԥ��.Enabled = True
    '��궨λ
    If Val(txt��Ԥ��.Text) <> 0 And txt��Ԥ��.Enabled Then
        If txt��Ԥ��.Enabled And txt��Ԥ��.Visible Then txt��Ԥ��.SetFocus
        Call Show�����(True)
    Else
        If txt�ɿ�.Enabled And txt�ɿ�.Visible Then txt�ɿ�.SetFocus
        Call Show�����(False)
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
        If mintInsure <> 0 And mblnYbBalanced = False Then
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
    mstrTitle = "�����շѽ���"
    RestoreWinState Me, App.ProductName, mstrTitle
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
    mlng����ID = 0
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
    SaveWinState Me, App.ProductName, mstrTitle
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
            If .dbl����Ԥ�� <> 0 Then
                txt��Ԥ��.Text = Format(IIf(.dbl����Ԥ�� > .dbl��ǰδ��, .dbl��ǰδ��, .dbl����Ԥ��), "###0.00;###0.00;0.00;0.00")
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
    stbThis.Panels(2).Text = mstrPatiInfo
    Set rsTemp = GetMoneyInfo(mlng����ID, 0, False, 1, False)
    With mCurCarge
        .dblԤ����� = 0
        .dbl������� = 0
        If Not rsTemp.EOF Then
            .dblԤ����� = RoundEx(Val(Nvl(rsTemp!Ԥ�����)), 6)
            .dbl������� = RoundEx(Val(Nvl(rsTemp!�������)), 6)
        End If
        .dbl����Ԥ�� = RoundEx(.dblԤ����� - .dbl�������, 6)
        If .dbl����Ԥ�� < 0 Then .dbl����Ԥ�� = 0
    End With
    txtҽ��.Text = Format(mCurCarge.dbl����ҽ��֧��, "###0.00;-###0.00;0.00;0.00;")
    txt�ϼ�.Text = Format(mCurCarge.dbl����ʵ��, "###0.00;-###0.00;0.00;0.00;")
    stbThis.Panels(3).Text = Format(mCurCarge.dbl����Ԥ��, "####0.00;-####0.00;0.00;0.00")
    
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
    If mintInsure <> 0 And mblnYbBalanced = False Then Exit Sub
    
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
     
    'If Val(txt��Ԥ��.Tag) = Val(txt��Ԥ��.Text) Then Exit Sub
    
    '�Զ����ۻ��ֹ�����ʱ���ȼ�����
    'Call LedVoiceSpeak
   
End Sub

Private Sub txt��Ԥ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
     zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txt��Ԥ��_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txt��Ԥ��, KeyAscii, m���ʽ
End Sub
Private Sub txt��Ԥ��_LostFocus()
      If mblnLoad Then Exit Sub
      
      If Val(txt��Ԥ��.Tag) = Val(txt��Ԥ��.Text) Then Exit Sub
      If Val(txt��Ԥ��.Text) = 0 Then Exit Sub
      If CheckPrepayMoneyIsValied = False Then Exit Sub
      
End Sub

Private Sub txt��Ԥ��_Validate(Cancel As Boolean)
    If lbl��Ԥ��.Tag = "1" Then Exit Sub
    If mlng����ID = 0 Then Exit Sub
    
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
    ElseIf Val(txt��Ԥ��.Text) > mCurCarge.dbl����Ԥ�� Then
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
    ElseIf Val(txt��Ԥ��.Text) > mCurCarge.dbl����Ԥ�� Then
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
    If zlDatabase.PatiIdentify(Me, glngSys, mlng����ID, Val(txt��Ԥ��), mlngModule, 1, mlngBrushCardTypeID, _
            IIf(-1 * gdblԤ��������鿨 >= Val(txt��Ԥ��), False, True), , , (gdblԤ��������鿨 <> 0), (gdblԤ��������鿨 = 2)) Then
        lbl��Ԥ��.Tag = "1"
       ' txt��Ԥ��.ForeColor = d
       txt��Ԥ��.BackColor = Me.BackColor
       txt��Ԥ��.Tag = Val(txt��Ԥ��)
       txt��Ԥ��.Enabled = False
        If SaveCharge(True) = False Then
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
    If Val(txt�ɿ�) = 0 Then
        MsgBox "δ���뽻�׽��,����!", vbInformation + vbOKOnly
         Exit Function
    End If
    If Not IsNumeric(txt�ɿ�.Text) And txt�ɿ�.Text <> "" Then
        MsgBox "��Ч��ֵ��", vbInformation, gstrSysName
        If txt�ɿ�.Enabled And txt�ɿ�.Visible Then txt�ɿ�.SetFocus
        zlControl.TxtSelAll txt�ɿ�: Exit Function
    ElseIf Val(txt�ɿ�.Text) < 0 Then
        MsgBox "���׽���Ϊ����", vbInformation, gstrSysName
        If txt�ɿ�.Enabled And txt�ɿ�.Visible Then txt�ɿ�.SetFocus
        zlControl.TxtSelAll txt�ɿ�: Exit Function
    ElseIf Abs(Val(txt�ɿ�.Text)) > Format(Abs(mCurCarge.dbl��ǰδ��), "0.00") And Val(txt�ɿ�.Text) <> 0 Then
        MsgBox "���׽��ܴ��ڱ���δ�����:" & Format(mCurCarge.dbl��ǰδ��, "0.00") & " ��", vbInformation, gstrSysName
        If txt�ɿ�.Enabled And txt�ɿ�.Visible Then txt�ɿ�.SetFocus
        zlControl.TxtSelAll txt�ɿ�: Exit Function
    End If
    If mCurCardPay.bln���ѿ� And mblnYB�˿� Then
        MsgBox "��ǰΪ�˿�ģʽ,Ŀǰϵͳ�ݲ�֧�ֽ��˿���˸�" & mCurCardPay.str���㷽ʽ, vbInformation + vbOKOnly, gstrSysName
        If txt�ɿ�.Enabled And txt�ɿ�.Visible Then txt�ɿ�.SetFocus
        zlControl.TxtSelAll txt�ɿ�: Exit Function
        Exit Function
    End If
    If zlGetClassMoney(mlng����ID, rsMoney) = False Then Exit Function
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
    Optional ByRef varSquareBalance As Variant) As Boolean
    '       varSquareBalance- Collection����,���ص�ǰˢ������(array(�����ID,���ѿ�ID,ˢ�����,����,����,�������,�Ƿ�����))
    Set cllSquareBalance = Nothing
    Set mcllCurSquareBalance = Nothing
    If mCurCardPay.bln���ѿ� Then
        '�������ѿ���ˢ����Ϣ
       Set cllSquareBalance = mcllSquareBalance
     End If
     
    dblMoney = Val(txt�ɿ�.Text)
    If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModule, rsMoney, _
        mCurCardPay.lngҽ�ƿ����ID, mCurCardPay.bln���ѿ�, _
    mstr����, mstr�Ա�, mstr����, dblMoney, mCurCardPay.strˢ������, mCurCardPay.strˢ������, _
    False, True, False, False, cllSquareBalance) = False Then Exit Function
    '���ѿ���ֵ
    If mCurCardPay.bln���ѿ� Then
        Set mcllCurSquareBalance = cllSquareBalance
    End If
    
    '����ǰ,һЩ���ݼ��
    'zlPaymentCheck(frmMain As Object, ByVal lngModule As Long, _
    ByVal strCardTypeID As Long, ByVal strCardNo As String, _
    ByVal dblMoney As Double, ByVal strNOs As String, _
    Optional ByVal strXMLExpend As String
    'mstrNOs:��������ʱ,û�����ʱ,����Ϊ��.
    If gobjSquare.objSquareCard.zlPaymentCheck(Me, mlngModule, mCurCardPay.lngҽ�ƿ����ID, _
        mCurCardPay.bln���ѿ�, mCurCardPay.strˢ������, dblMoney, mstrNOs, strXMLExpend) = False Then Exit Function
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
Private Function zlGetClassMoney(ByRef lng������� As Long, ByRef rsMoney As ADODB.Recordset) As Boolean
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
    If lng������� = 0 And mbytFunc = EM_�����շ� Then
        Call mfrmMain.zlGetClassMoney(rsTemp)
    Else
        strSQL = "" & _
        "   Select  A.�շ����,nvl(sum(ʵ�ս��) ,0) as ���   " & _
        "   From ������ü�¼ A,(Select ����ID From ����Ԥ����¼ where �������=[1] ) B " & _
        "   Where A.����ID=B.����ID " & _
        "   Group by �շ����"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng�������)
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
Private Sub txt�ɿ�_Change()
    Call SetControlProperty
    Call Show�����(False)
End Sub
Private Sub txt�ɿ�_GotFocus()
    'ֻ�Խɿ���Ϊ�շѽ�������ʱ,��������ɿ��0
    '���˺�:22343
    If gTy_Module_Para.byt�ɿ���� = 1 _
        Or gTy_Module_Para.byt�ɿ���� = 3 _
        Or gTy_Module_Para.byt�ɿ���� = 2 Then
        If Val(txt�ɿ�.Text) = 0 And Me.ActiveControl Is txt�ɿ� Then
            txt�ɿ�.Text = ""
        End If
    End If
    Call SetControlProperty
  '  Call zlControl.TxtSelAll(txt�ɿ�)
    '�Զ����ۻ��ֹ�����ʱ���ȼ�����
    If Not mbln�ѱ��� Then Call LedVoiceSpeak
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
        zl9LedVoice.DispCharge mCurCarge.dbl��ǰδ�� + mCurCarge.dblӦ���ۼ�, Val(txt�ɿ�.Text), Val(txt�Ҳ�.Text)
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
        If CSng(txt�Ҳ�.Text) >= 0 Then
            'LED��ʾ
            'Call ShowLedInfor
            'ȷ��
             Call cmdOK_Click
        Else
            MsgBox "�ɿ����,�벹��Ӧ�ɽ�", vbInformation, gstrSysName
            txt�ɿ�.SetFocus: zlControl.TxtSelAll txt�ɿ�
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
                varTemp = Split(varData(i) & "|||||", "|")
                If varTemp(6) = Nvl(rsTemp!����) Then
                    blnFind = True
                    Exit For
                End If
            Next
            If Not blnFind Then
                If Not (Val(Nvl(rsTemp!����)) = 3 Or Val(Nvl(rsTemp!����)) = 4 Or Val(Nvl(rsTemp!Ӧ����)) = 1) Then
                    '������ҽ���Ľ��㷽ʽ
                    .AddItem Nvl(rsTemp!����)
                    mcolCardPayMode.Add Array("", Nvl(rsTemp!����), 0, 0, 0, 0, Nvl(rsTemp!����), 0, 0), "K" & j
                    If rsTemp!ȱʡ = 1 Then .ListIndex = .NewIndex
                    If Val(Nvl(rsTemp!����)) = 1 And .ListIndex < 0 Then .ListIndex = .NewIndex
                    .ItemData(.NewIndex) = Val(Nvl(rsTemp!����))
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
            If InStr(1, varData(i), "|") <> 0 Then
                varTemp = Split(varData(i), "|")
                mcolCardPayMode.Add varTemp, "K" & j
                .AddItem varTemp(1): .ItemData(.NewIndex) = -1
                If mbln�������� Then
                    '   '��|ȫ��|ˢ����־|�����ID(���ѿ����)|����|�Ƿ����ѿ�|���㷽ʽ|�Ƿ�����|�Ƿ����ƿ�;��
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
        If .ListCount > 0 And .ListIndex < 0 Then .ListIndex = 0
        If Not mbln�������� And gstr���㷽ʽ <> "" Then
            '60574
            '���ݲ�������ȱʡ��֧�����
            For j = 0 To .ListCount - 1
                If .List(j) = gstr���㷽ʽ Then
                    .ListIndex = j: Exit For
                End If
            Next
        End If
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
End Sub
Private Sub txt�������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txt�������_KeyPress(KeyAscii As Integer)
    If InStr(":'|", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    zlControl.TxtCheckKeyPress txt�������, KeyAscii, m�ı�ʽ
End Sub

Private Sub txtժҪ_GotFocus()
    zlControl.TxtSelAll txtժҪ
End Sub

Private Sub txtժҪ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If cmdOK.Visible And cmdOK.Enabled Then cmdOK.SetFocus
    End If
End Sub
Private Sub txt�Ҳ�_GotFocus()
    zlControl.TxtSelAll txt�Ҳ�
End Sub

Private Function zlOneCardPrayMoney(ByVal dblMoney As Double, ByRef strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:һ��֧ͨ��
    '����:һ��֧ͨ���ɹ�,����true,���򷵻�false
    '����:���˺�
    '����:2011-08-23 17:57:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl��� As Double, strҽԺ���� As String
    If mCurCardPay.blnOneCard = False Then zlOneCardPrayMoney = True: Exit Function
    mrsOneCard.Filter = "���㷽ʽ='" & mCurCardPay.str���㷽ʽ & "'"
    If mrsOneCard.EOF Then
         strErrMsg = "δ�ҵ����㷽ʽΪ" & mCurCardPay.str���㷽ʽ & "��һ��ͨ!"
         Exit Function
    End If
    'һ��ͨ���㣨�޸ĵ���ʱ��Ϊû�ж������޷�ȷ��ʹ��������һ��ͨ�����Բ�֧���޸Ĺ���)
    Dim intCardType As Integer, strSwapNO As String
    If Not mobjICCard.PaymentSwap(dblMoney, dbl���, intCardType, Val("" & mrsOneCard!ҽԺ����), mCurCardPay.strˢ������, strSwapNO, mlng����ID, mlng����ID) Then
         strErrMsg = "һ��ͨ����ʧ��"
        Exit Function
    End If
    gstrSQL = "zl_һ��ͨ����_Update(" & 0 & ",'" & mCurCardPay.str���㷽ʽ & "','" & mCurCardPay.strˢ������ & "','" & intCardType & "','" & strSwapNO & "'," & dbl��� & "," & mlng����ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    zlOneCardPrayMoney = True
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
    If gobjSquare.objSquareCard.zlPaymentMoney(Me, mlngModule, mCurCardPay.lngҽ�ƿ����ID, mCurCardPay.bln���ѿ�, mCurCardPay.strˢ������, mstr����IDs, mCurCardPay.strNo, dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor) = False Then Exit Function
    '����������������
     If mCurCardPay.lngҽ�ƿ����ID <> 0 And mlng����ID <> 0 And cbo֧����ʽ.Visible Then
        mCurCardPay.str������ˮ�� = strSwapGlideNO
        mCurCardPay.str����˵�� = strSwapMemo
        If mCurCardPay.bln���ѿ� = False Then
            Call zlAddUpdateSwapSQL(False, mstr����IDs, mCurCardPay.lngҽ�ƿ����ID, mCurCardPay.bln���ѿ�, mCurCardPay.strˢ������, strSwapGlideNO, strSwapMemo, cllUpdate, 2)
        End If
        Call zlAddThreeSwapSQLToCollection(False, mstr����IDs, mCurCardPay.lngҽ�ƿ����ID, mCurCardPay.bln���ѿ�, mCurCardPay.strˢ������, strSwapExtendInfor, cllThreeSwap)
    End If
    zlInterfacePrayMoney = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function ChargeOver(ByVal blnNotCommit As Boolean, ByVal dbl��֧Ʊ�� As Double) As Boolean
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
    'Zl_�����շѽ���_����շ�
    strSQL = "Zl_�����շѽ���_����շ�("
    '  ����id_In     ������ü�¼.����id%Type,
    strSQL = strSQL & "" & mlng����ID & ","
    '  �������id_In ����Ԥ����¼.�������%Type,
    strSQL = strSQL & "" & mlng����ID & ","
    '  �ɿ�_In       ����Ԥ����¼.�ɿ�%Type := Null,
    strSQL = strSQL & "" & dbl�ɿ� & ","
    '  �Ҳ�_In       ����Ԥ����¼.�Ҳ�%Type := Null,
    strSQL = strSQL & "" & dbl�Ҳ� & ","
    '  �����_In   ������ü�¼.ʵ�ս��%Type := Null,
    strSQL = strSQL & "" & mCurCarge.dbl�������� & ","
    '  ���㷽ʽ_In   ����Ԥ����¼.���㷽ʽ%Type := Null,
    strSQL = strSQL & "'" & mCurCardPay.str���㷽ʽ & "',"
    '  Ԥ���_In     ����Ԥ����¼.��Ԥ��%Type := Null,
    strSQL = strSQL & "" & dblԤ��� & ","
    '  ��֧Ʊ��_In   ����Ԥ����¼.��Ԥ��%Type := Null,
    strSQL = strSQL & "" & dbl��֧Ʊ�� & ","
    '  �շѽ���_In Varchar2:=Null
    strSQL = strSQL & "'" & str�շѽ��� & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    mdbl�ɿ��� = dbl�ɿ�: mdbl�Ҳ� = dbl�Ҳ�
    If blnNotCommit Then gcnOracle.CommitTrans
    ChargeOver = True
    Exit Function
errHandle:
    If blnNotCommit Then gcnOracle.RollbackTrans
    Call ErrCenter
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
        mCurCarge.dbl�������� = -1 * (mCurCarge.dbl����ʵ�� - mCurCarge.dbl�����Ѹ��ϼ� - RoundEx(mCurCarge.dbl��ǰδ��, 2))
    ElseIf mCurCardPay.int���� = 1 Then
        dblTemp = IIf(dblMoney = 0, dblʣ����, mCurCarge.dbl��ǰδ��): dblʣ���� = 0
        If mintInsure > 0 Then  '����:43855
            If mInsurePara.�ֱҴ��� Then
                dblMoney = CentMoney(CCur(dblTemp))
            Else
                dblMoney = Format(dblTemp, "0.00")
            End If
        Else
                dblMoney = CentMoney(CCur(dblTemp))
        End If
        mCurCarge.dbl�������� = -1 * (mCurCarge.dbl����ʵ�� - mCurCarge.dbl�����Ѹ��ϼ� - dblMoney)
    ElseIf mCurCardPay.bln֧Ʊ Then
        'ֻ���ֽ��������
'        If dblʣ���� < 0 Then
'            dbl��֧Ʊ�� = -1 * Val(txt�Ҳ�.Text)
'            mCurCarge.dbl�������� = Format(-1 * (mCurCarge.dbl����ʵ�� - mCurCarge.dbl�����Ѹ��ϼ� - dblMoney - dbl��֧Ʊ��), gstrDec)
'        End If
    Else
        'ֻ���ֽ��������
        'mCurCarge.dbl�������� = Format(-1 * (mCurCarge.dbl����ʵ�� - mCurCarge.dbl�����Ѹ��ϼ� - RoundEx(mCurCarge.dbl��ǰδ��, 2)), gstrDec)
    End If
    If mblnCur���� And Val(txt�ɿ�.Text) = 0 Then
'        dblMoney = mCurCarge.dbl��ǰδ��
'        mCurCarge.dbl�������� = Format(-1 * (mCurCarge.dbl����ʵ�� - mCurCarge.dbl�����Ѹ��ϼ� - RoundEx(mCurCarge.dbl��ǰδ��, 2)), gstrDec)
'        dblʣ���� = 0
    End If
    '����:47637
    'δ����ҽ������ǰ,����ʾ���
    If mintInsure <> 0 And mblnYbBalanced = False Then mCurCarge.dbl�������� = 0
    mCurCarge.dbl�������� = Format(mCurCarge.dbl��������, gstrDec)
    pic���.Visible = mCurCarge.dbl�������� <> 0
    lbl����.Caption = Format(mCurCarge.dbl��������, gstrDec)
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
    If mintInsure <> 0 And mblnYbBalanced Then intCount = intCount + 1: strErrMsg = strErrMsg & "ҽ������:" & txtҽ��.Text
   With vsBlance
        For i = 1 To .Rows - 1
            str���㷽ʽ = Trim(.TextMatrix(i, .ColIndex("֧����ʽ")))
            int���� = Val(.RowData(i))
            '.rowdata:0-��ͨ�Ľ��㷽ʽ-1-ҽ������;2-�����ӿڽ���;3-һ��ͨ����;4-Ԥ���
            If InStr("23", int����) > 0 Then
                If int���� = 3 Then intCount = intCount + 1:
                If int���� = 2 Then '�����ӿ�
                    ' ' ҽ�ƿ����ID|���ѿ�(1, 0) |���ƿ�| �ӿ�����
                    varData = Split(.Cell(flexcpData, i, .ColIndex("֧����ʽ")) & "|||||", "|")
                    If Val(varData(1)) = 1 Then '���ѿ�
                        '���ƿ�,��������
                        If Val(varData(2)) = 0 Then intCount = intCount + 1:  strErrMsg = strErrMsg & vbCrLf & varData(3) & ":" & .TextMatrix(i, .ColIndex("֧�����"))
                    Else
                         intCount = intCount + 1: strErrMsg = strErrMsg & vbCrLf & varData(3) & ":" & .TextMatrix(i, .ColIndex("֧�����"))
                    End If
                End If
            End If
        Next
    End With
    If intCount > 2 Then
        Call MsgBox("ע��:" & vbCrLf & "   ��ϵͳĿǰֻ֧���������½ӿ�,�����Ѿ��������½ӿڽ���:" & vbCrLf & strErrMsg, vbInformation + vbOKOnly)
        Exit Function
    End If
    zlCheckMulitInterfaceNumValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function
Private Function SaveCharge(Optional blnԤ�� As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����������
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-14 17:38:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnHaveMoney As Boolean, dblʣ���� As Double, strSQL As String
    Dim dblMoney As Double, strErrMsg As String, dbl��֧Ʊ�� As Double
    Dim i As Integer, blnFind As Boolean, cllPro As Collection
    Dim str���ѿ����� As String, j As Long
    Dim strCardNo As String, dblTemp As Double, blnNotCommit As Boolean '�����������ύ
    Dim cllUpdate As Collection, cllThreeSwap As Collection
    Dim blnSaveBilling As Boolean   '��ǰ���񱣴浥��
    
    On Error GoTo errHandle
    blnSaveBilling = False
    If zlCheckMulitInterfaceNumValied = False Then Exit Function
    mstrBalances = "" '����:42791
    mdbl�ֽ� = 0
    dblMoney = IIf(blnԤ��, Val(txt��Ԥ��.Text), IIf(mblnYB�˿�, -1, 1) * Val(txt�ɿ�.Text))
    dbl��֧Ʊ�� = 0
    dblʣ���� = mCurCarge.dbl��ǰδ�� - dblMoney
    If blnԤ�� Then
        dblMoney = Val(txt��Ԥ��.Text)
        mstrBalances = mstrBalances & "|��Ԥ��:" & dblMoney
        '����:58344
        '����Ƿ�ǰ֧�����Ϊ����,�Ǹ���ʱ,��Ҫ���Ѳ���Ա(��Ҫ��ҽ������ʱ���ܴ��ڱ����ݵķ���)
        If mblnYB�˿� Then
              Call MsgBox("ע��:" & vbCrLf & "    ��ǰ�����˿ʽ,������ʹ��Ԥ����!", vbExclamation + vbOKOnly + vbDefaultButton2, gstrSysName)
              Exit Function
        End If
    ElseIf mCurCardPay.int���� = 1 Then
        dblTemp = IIf(dblMoney = 0, dblʣ����, mCurCarge.dbl��ǰδ��): dblʣ���� = 0
        If mintInsure > 0 Then  '����:43855
            If gclsInsure.GetCapability(support�ֱҴ���, , mintInsure) Then
                dblMoney = CentMoney(CCur(dblTemp))
            Else
                dblMoney = Format(dblTemp, "0.00")
            End If
        Else
                dblMoney = CentMoney(CCur(dblTemp))
        End If
        '����:58344
        '����Ƿ�ǰ֧�����Ϊ����,�Ǹ���ʱ,��Ҫ���Ѳ���Ա(��Ҫ��ҽ������ʱ���ܴ��ڱ����ݵķ���)
        If mblnYB�˿� Then
             If MsgBox("ע��:" & vbCrLf & "    δ������Ϊ�˿�,���Ƿ����Ҫ�ˡ�" & mCurCardPay.str���㷽ʽ & ":" & Abs(dblMoney) & "��������?" & vbCrLf & IIf(Val(txt�ɿ�.Text) <> 0, "  ��ǰ�˸������ܶ�:" & txt�ɿ�.Text & vbCrLf & "  ��ǰӦ�ջ��ܶ�:" & txt�Ҳ�.Text, ""), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
             If Abs(dblMoney) < Abs(lblʣ���Ը�.Caption) Then
                Call MsgBox("ע��:" & vbCrLf & "    δ������Ϊ�˿�,�㲻�ܽ��ж���˿����," & vbCrLf & "��ǰ�˽��(" & Format(dblMoney, "0.00") & ")�������ʣ����(" & lblʣ���Ը�.Caption & ")!", vbInactiveBorder + vbOKOnly + vbDefaultButton2, gstrSysName)
                Exit Function
             End If
        End If
        mdbl�ֽ� = dblMoney
        If Val(txt�ɿ�.Text) <> 0 Then
            mstrBalances = mstrBalances & "|�ɿ�:" & IIf(mblnYB�˿�, -1, 1) * Val(txt�ɿ�.Text) & ":1"
            mstrBalances = mstrBalances & "|�Ҳ�:" & IIf(mblnYB�˿�, -1, 1) * Val(txt�Ҳ�.Text) & ":2"
        End If
        mstrBalances = mstrBalances & "|" & mCurCardPay.str���㷽ʽ & ":" & dblMoney
    ElseIf mCurCardPay.bln֧Ʊ Then
        mstrBalances = mstrBalances & "|" & mCurCardPay.str���㷽ʽ & ":" & dblMoney
        '����:58344
        '����Ƿ�ǰ֧�����Ϊ����,�Ǹ���ʱ,��Ҫ���Ѳ���Ա(��Ҫ��ҽ������ʱ���ܴ��ڱ����ݵķ���)
        If mblnYB�˿� Then
             If MsgBox("ע��:" & vbCrLf & "    δ������Ϊ�˿�,���Ƿ����Ҫ�ˡ�" & mCurCardPay.str���㷽ʽ & ":" & Abs(dblMoney) & "��������?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
             If Abs(dblMoney) <> Abs(lblʣ���Ը�.Caption) Then
                Call MsgBox("ע��:" & vbCrLf & "    δ������Ϊ�˿�,��ǰ�˽��(" & Format(Abs(dblMoney), "0.00") & ")�������ʣ����(" & Abs(Val(lblʣ���Ը�.Caption)) & ")!", vbInactiveBorder + vbOKOnly + vbDefaultButton2, gstrSysName)
                Exit Function
             End If
        Else
            
            If dblʣ���� < 0 Then
                If mstr��֧Ʊ = "" Then
                    MsgBox "�ڽ��㷽ʽ��û������Ӧ����Ľ��㷽ʽ,���ܽ�����֧Ʊ����", vbOKOnly + vbInformation, gstrSysName
                    Exit Function
                End If
                dbl��֧Ʊ�� = -1 * Val(txt�Ҳ�.Text)
                mstrBalances = mstrBalances & "|" & mstr��֧Ʊ & ":" & -1 * dbl��֧Ʊ�� & ":2"
            End If
        End If
    Else
        '����:58344
        '����Ƿ�ǰ֧�����Ϊ����,�Ǹ���ʱ,��Ҫ���Ѳ���Ա(��Ҫ��ҽ������ʱ���ܴ��ڱ����ݵķ���)
        If mblnYB�˿� Then
             If MsgBox("ע��:" & vbCrLf & "    δ������Ϊ�˿�,���Ƿ����Ҫ�ˡ�" & mCurCardPay.str���㷽ʽ & ":" & Abs(dblMoney) & "��������?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
             If Abs(dblMoney) <> Abs(lblʣ���Ը�.Caption) Then
                Call MsgBox("ע��:" & vbCrLf & "    δ������Ϊ�˿�,��ǰ�˽��(" & Format(Abs(dblMoney), "0.00") & ")�������ʣ����(" & Abs(Val(lblʣ���Ը�.Caption)) & ")!", vbInactiveBorder + vbOKOnly + vbDefaultButton2, gstrSysName)
                Exit Function
             End If
        End If
        mstrBalances = mstrBalances & "|" & mCurCardPay.str���㷽ʽ & ":" & dblMoney
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
    If dblʣ���� > 0 Then blnHaveMoney = True
    With vsBlance
        blnFind = False
        For i = 1 To .Rows - 1
            If blnԤ�� Then
                If Val(.RowData(i)) = 4 Then blnFind = True
            ElseIf mCurCardPay.bln���ѿ� And mCurCardPay.bln���ƿ� Then
                '���ѿ�,�Ѿ����,�����ٴ���
            Else
                If .TextMatrix(i, .ColIndex("֧����ʽ")) = mCurCardPay.str���㷽ʽ Then
                    blnFind = True
                End If
            End If
            mstrBalances = mstrBalances & "|" & .TextMatrix(i, .ColIndex("֧����ʽ")) & ":" & .TextMatrix(i, .ColIndex("֧�����"))
        Next
        
        If blnFind Then
            If blnԤ�� Then
                MsgBox "�Ѿ���Ԥ���֧��,ֻ��ɾ��Ԥ�������֧��!", vbOKOnly + gstrSysName
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
    If mCurCardPay.bln���ѿ� Then
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
    Err = 0: On Error GoTo ErrCommit:
    If Not (blnԤ�� Or mCurCardPay.lngҽ�ƿ����ID = 0 Or cbo֧����ʽ.ItemData(cbo֧����ʽ.ListIndex) <> -1) Then
        '�������ӿڵ���ؽ���,��Ҫ�ȴ���ӿ�����
        blnNotCommit = False
        If Not mblnYbBalanced And mlng����ID = 0 Then
            blnNotCommit = True
            If SaveBill(blnNotCommit) = False Then
                blnNotCommit = False: mlng����ID = 0: Exit Function
            End If
            blnSaveBilling = True
        End If
        
        'Zl_�����շѽ���_Modify
        strSQL = "Zl_�����շѽ���_Modify("
        '  ��������_In   Number,
        '--��������_In:0-�����շ�;
        '--            1-��Ԥ��(���㷽ʽΪNULL,������<>0);
        '--            2-ҽ������:�����ҽ������,���㷽ʽ_IN����Ϊ���;
        '--            3-���ѿ���������(���㷽ʽ_IN��ʽΪ:�����ID|����|���ѿ�ID|���ѽ��||....)
        strSQL = strSQL & IIf(blnԤ��, "1", IIf(mCurCardPay.bln���ѿ�, "3", "0")) & ","
        '  ����id_In     ������ü�¼.����id%Type,
        strSQL = strSQL & mlng����ID & ","
        '  �������id_In ����Ԥ����¼.�������%Type,
        strSQL = strSQL & mlng����ID & ","
        '  ���㷽ʽ_In   Varchar2,
        If blnԤ�� Then
            strSQL = strSQL & "NULL" & ","
        ElseIf mCurCardPay.bln���ѿ� Then
            strSQL = strSQL & "'" & str���ѿ����� & "'" & ","
        Else
            strSQL = strSQL & "'" & mCurCardPay.str���㷽ʽ & "'" & ","
        End If
        '  ������_In   ����Ԥ����¼.��Ԥ��%Type,
        strSQL = strSQL & dblMoney & ","
        ' ��֧Ʊ��_In   ����Ԥ����¼.��Ԥ��%Type,
        strSQL = strSQL & dbl��֧Ʊ�� & ","
        '  ժҪ_In       ����Ԥ����¼.ժҪ%Type := Null,
        strSQL = strSQL & "'" & Trim(txtժҪ.Text) & "',"
        '  �������_In   ����Ԥ����¼.�������%Type := Null,
        strSQL = strSQL & "" & IIf(mCurCardPay.bln֧Ʊ Or txt�������.Visible, "'" & Trim(txt�������.Text) & "'", "NULL") & ","
        '  �����id_In   ����Ԥ����¼.�����id%Type := Null,
        strSQL = strSQL & "" & IIf(mCurCardPay.lngҽ�ƿ����ID = 0, "NULL", mCurCardPay.lngҽ�ƿ����ID) & ","
        '  ���ѿ�_In     Integer := 0,
        strSQL = strSQL & "" & IIf(mCurCardPay.bln���ѿ�, 1, 0) & ","
        '  ����_In       ����Ԥ����¼.����%Type := Null,
        strSQL = strSQL & "" & IIf(mCurCardPay.strˢ������ <> "", "'" & mCurCardPay.strˢ������ & "'", "NULL") & ","
        '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
        strSQL = strSQL & "NULL" & ","
        '  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null
        strSQL = strSQL & "NULL" & ")"
        zlAddArray cllPro, strSQL
        Call zlExecuteProcedureArrAy(cllPro, Me.Caption, True, blnNotCommit)
        If Not mCurCardPay.bln���ѿ� Then
            '���ѿ����ٵ��ýӿ�
             If zlInterfacePrayMoney(cllUpdate, cllThreeSwap, dblMoney) = False Then
                 '����:47637
                  If Not (mblnYbBalanced Or mblnThreeInterface) Or blnSaveBilling Then mlng����ID = 0
                    gcnOracle.RollbackTrans: Exit Function
            End If
        End If
        
        Err = 0: On Error GoTo ErrUpdate:
        'һ��ͨ����
        If zlOneCardPrayMoney(dblMoney, strErrMsg) = False Then
            gcnOracle.RollbackTrans
            MsgBox strErrMsg, vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
        gcnOracle.CommitTrans
        Call zlExecuteProcedureArrAy(cllUpdate, Me.Caption)
        blnNotCommit = False: mblnThreeInterface = True
        Call SetCtrlVisible
        On Error GoTo ErrOthers:
        Call zlExecuteProcedureArrAy(cllThreeSwap, Me.Caption)
    End If
GoOver:
    If mintInsure <> 0 Then
        If Not (blnԤ�� Or mCurCardPay.lngҽ�ƿ����ID <> 0 _
            Or mCurCardPay.blnOneCard) Then
            'ֻ��ҽ�����˲Ż�������½϶Ե����,��˲Ż����¼��㱾��Ӧ�ɵ����
            '��Ҫ�Ǹ��������շѵ�����
            mdbl����Ӧ�� = mdbl����Ӧ�� + dblMoney
        End If
    End If
    
    If Not blnHaveMoney Then
        If mlng����ID = 0 Then
            blnNotCommit = True
            If SaveBill(blnNotCommit) = False Then
                blnNotCommit = False: mlng����ID = 0: Exit Function
            End If
            blnSaveBilling = True
        End If
        If ChargeOver(blnNotCommit, dbl��֧Ʊ��) = False Then
            If blnNotCommit Or blnSaveBilling Then mlng����ID = 0
            Exit Function
        End If
        Call WhriteTotalDataToReCord(IIf(blnԤ��, dblMoney, 0), IIf(Not blnԤ��, dblMoney, 0), dbl��֧Ʊ��)
        mblnOK = True
        SaveCharge = True: mblnUnloaded = True
        
        Unload Me:
        Exit Function
    End If
    mstrBalances = ""
    If Not blnԤ�� And mCurCardPay.int���� = 1 Then
       '�ֽ�
        SaveCharge = True: Exit Function
    End If
    Err = 0: On Error GoTo errHandle:
    With vsBlance
        If mCurCardPay.bln���ѿ� Then
            If mcllSquareBalance Is Nothing Then Set mcllSquareBalance = New Collection
            For j = 1 To mcllCurSquareBalance.Count
                '��ǰˢ������(array(�����ID,���ѿ�ID,ˢ�����,����,����,�������,�Ƿ�����))
                mcllSquareBalance.Add mcllCurSquareBalance(j)
                If Not (.Rows = 2 And Trim(.TextMatrix(1, .ColIndex("֧����ʽ"))) = "") Then
                    .Rows = .Rows + 1
                    .RowPosition(.Rows - 1) = 1
                End If
                .RowData(1) = 0
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
            .RowData(1) = 0
            strCardNo = mCurCardPay.strˢ������
            If blnԤ�� Then
                .TextMatrix(1, .ColIndex("֧����ʽ")) = "Ԥ���"
                .RowData(1) = 4
            ElseIf mCurCardPay.lngҽ�ƿ����ID <> 0 Then
                .TextMatrix(1, .ColIndex("֧����ʽ")) = mCurCardPay.str���㷽ʽ
                 ' ҽ�ƿ����ID|���ѿ�(1, 0) |���ƿ�| �ӿ�����
                .Cell(flexcpData, 1, .ColIndex("֧����ʽ")) = mCurCardPay.lngҽ�ƿ����ID & "|" & IIf(mCurCardPay.bln���ѿ�, 1, 0) & "|" & IIf(mCurCardPay.bln���ƿ�, 1, 0) & "|" & mCurCardPay.str����
                .RowData(1) = 2
                strCardNo = gobjSquare.objSquareCard.zlGetCardNODencode(mCurCardPay.strˢ������, mCurCardPay.lngҽ�ƿ����ID, mCurCardPay.bln���ѿ�)
            ElseIf mCurCardPay.blnOneCard Then
                .TextMatrix(1, .ColIndex("֧����ʽ")) = mCurCardPay.str���㷽ʽ
                .RowData(1) = 3
            Else
                .TextMatrix(1, .ColIndex("֧����ʽ")) = mCurCardPay.str���㷽ʽ
            End If
            .TextMatrix(1, .ColIndex("֧�����")) = Format(dblMoney, "0.00")
            .TextMatrix(1, .ColIndex("�������")) = IIf(txt�������.Visible, Trim(txt�������.Text), "")
            .TextMatrix(1, .ColIndex("��ע")) = Trim(txtժҪ.Text)
            
            .TextMatrix(1, .ColIndex("����")) = IIf(mCurCardPay.bln��������, String(Len(strCardNo), "*"), strCardNo)
            .Cell(flexcpData, 1, .ColIndex("����")) = mCurCardPay.strˢ������
            .TextMatrix(1, .ColIndex("������ˮ��")) = mCurCardPay.str������ˮ��
            .TextMatrix(1, .ColIndex("����˵��")) = mCurCardPay.str����˵��
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
ErrCommit:
    gcnOracle.RollbackTrans
    If blnSaveBilling Then mlng����ID = 0
    
ErrUpdate:
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    
    Exit Function
ErrOthers:
    '����������Ϣ,�ܱ�����������,��������������.
    If ErrCenter = 1 Then
        gcnOracle.RollbackTrans
        Resume
    End If
    gcnOracle.CommitTrans
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
     If vsBlance.Row < 0 Then
        int���� = -1
     Else
        int���� = Val(vsBlance.RowData(vsBlance.Row))
    End If
     '.rowdata:0-��ͨ�Ľ��㷽ʽ-1-ҽ������;2-�����ӿڽ���;3-һ��ͨ����;4-Ԥ���
    cmdDel.Visible = (int���� = 0 Or int���� = 4) And mbytFunc <> EM_�쳣����
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
     ByRef str����ID As String, _
    ByVal dbl�˿��� As Double) As Boolean
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
        "3|" & str����ID, dbl�˿���, strSwapNO, strSwapMemo, strXMLExend) = False Then
          zlCheckDelValied = False
          Exit Function
     End If
goEnd:
    zlCheckDelValied = True
    Exit Function
End Function

Private Function CallBackBalanceInterface(ByVal str����IDs As String, _
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
    Dim strԭ����IDs As String, strSwapExtendInfor As String, strTemp As String
    Err = 0: On Error GoTo Errhand:
    
    '�����ID,����,�Ƿ����ѿ�(1-��;0-��),������ˮ��,����˵��,strNO
    If lng�����ID = 0 Then CallBackBalanceInterface = True: Exit Function
    If Left(str����IDs, 1) = "," Then str����IDs = Mid(str����IDs, 2)
    strSQL = "" & _
    "    Select A.NO From ������ü�¼ A,����Ԥ����¼ M,Table( f_Num2list([1]))  B   " & _
    "    Where  A.��¼���� = 1 And A.��¼״̬=2  " & _
    "               And A.����ID=M.����ID  " & IIf(bln���ѿ�, " And nvl(M.���㿨���,0)=[2]", " And nvl(M.�����ID,0)=[2] ") & _
    "           And A.����ID=B.Column_Value " & _
    "      "
   strSQL = "" & _
   "    Select /*+ RULE */ distinct  ����ID From ������ü�¼ Q, (" & strSQL & ") M  " & _
   "    Where Q.NO=M.NO and Q.��¼����=1 and Q.��¼״̬=3  "
   '61688
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str����IDs, lng�����ID)
    If rsTemp.EOF Then
        strErrMsg = "δ�ҵ��������㽻����Ϣ���˷�ʧ��": Exit Function
    End If
    With rsTemp
        strԭ����IDs = ""
        Do While Not .EOF
            strԭ����IDs = strԭ����IDs & "," & Nvl(!����ID)
            .MoveNext
        Loop
    End With
    If strԭ����IDs <> "" Then strԭ����IDs = Mid(strԭ����IDs, 2)
    '81489,Ƚ����,2015-1-22,�˷Ѵ������ID
    strSwapExtendInfor = "3|" & str����IDs: strTemp = strSwapExtendInfor
    
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
    If gobjSquare.objSquareCard.zlReturnMoney(Me, mlngModule, lng�����ID, bln���ѿ�, strCardNo, "3|" & strԭ����IDs, dblMoney, strSwapNO, strSwapMemo, strSwapExtendInfor) = False Then Exit Function
    Set cllUpdate = New Collection: Set cllThreeSwap = New Collection
    Call zlAddUpdateSwapSQL(False, str����IDs, lng�����ID, bln���ѿ�, strCardNo, strSwapNO, strSwapMemo, cllUpdate)
    If strTemp <> strSwapExtendInfor Then
        Call zlAddThreeSwapSQLToCollection(False, str����IDs, lng�����ID, bln���ѿ�, strCardNo, strSwapExtendInfor, cllThreeSwap)
    End If
    CallBackBalanceInterface = True
Errhand:
End Function
Private Function zlGetInvoiceGroupUseID(ByRef lng����ID As Long, _
    Optional intNum As Integer = 1, Optional strInvoiceNO As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡƱ�ݵ�����ID
    '���:lng����ID-����id
    '       intNum-ҳ��
    '       strInvoiceNO-����ķ�Ʊ��
    '����:lng����ID-����ID
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-04-29 15:36:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    lng����ID = GetInvoiceGroupID(1, intNum, lng����ID, mlngShareUseID, strInvoiceNO, mstrUseType)
    If lng����ID <= 0 Then
        Select Case lng����ID
            Case 0 '����ʧ��
            Case -1
                If Trim(mstrUseType) = "" Then
                    MsgBox "��û�����ú͹��õ��շ�Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                Else
                    MsgBox "��û�����ú͹��õġ�" & mstrUseType & "���շ�Ʊ��,��������һ��Ʊ�ݻ����ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                End If
                Exit Function
            Case -2
                If Trim(mstrUseType) = "" Then
                    MsgBox "���صĹ���Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                Else
                    MsgBox "���صĹ���Ʊ�ݵġ�" & mstrUseType & "���շ�Ʊ���Ѿ�����,��������һ��Ʊ�ݻ��������ñ��ع���Ʊ�ݣ�", vbInformation, gstrSysName
                End If
                Exit Function
        End Select
    End If
    zlGetInvoiceGroupUseID = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
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
        Set mrsBlance = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng����ID)
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
    blnEdit = (mintInsure = 0 Or mintInsure <> 0 And mblnYbBalanced = True) And mbytFunc <> EM_�쳣����
    blnEdit = blnEdit Or mbytFunc = EM_�����շ�
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
            dblMoney = IIf(mblnYB�˿�, -1, 1) * Val(txt�ɿ�.Text)
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
                    str�շѽ��� = str�շѽ��� & "|" & IIf(Trim(txtժҪ) = "", " ", Trim(txtժҪ))
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
 
