VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDockDiagEdit 
   BorderStyle     =   0  'None
   Caption         =   "���뵥���"
   ClientHeight    =   4125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   9990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picXY 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   120
      ScaleHeight     =   3495
      ScaleWidth      =   9855
      TabIndex        =   0
      Top             =   360
      Width           =   9855
      Begin VB.Frame fraMore 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   0
         TabIndex        =   40
         Top             =   0
         Visible         =   0   'False
         Width           =   225
         Begin VB.Image imgMore 
            Height          =   225
            Left            =   0
            Picture         =   "frmDockDiagEdit.frx":0000
            Top             =   0
            Width           =   225
         End
      End
      Begin VB.OptionButton optDiag 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "���ݼ�����������(&2)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   6720
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   60
         Width           =   2010
      End
      Begin VB.OptionButton optDiag 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "������ϱ�׼����(&1)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   4680
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   60
         Value           =   -1  'True
         Width           =   2010
      End
      Begin VB.Frame frmXY 
         BorderStyle     =   0  'None
         Height          =   1815
         Left            =   0
         TabIndex        =   1
         Top             =   1680
         Visible         =   0   'False
         Width           =   10095
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   300
            Index           =   60
            ItemData        =   "frmDockDiagEdit.frx":0401
            Left            =   1425
            List            =   "frmDockDiagEdit.frx":0403
            Locked          =   -1  'True
            TabIndex        =   13
            TabStop         =   0   'False
            Text            =   "cboBaseInfo"
            Top             =   1380
            Width           =   675
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   14
            ItemData        =   "frmDockDiagEdit.frx":0405
            Left            =   7920
            List            =   "frmDockDiagEdit.frx":0407
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   135
            Width           =   1470
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   15
            ItemData        =   "frmDockDiagEdit.frx":0409
            Left            =   1425
            List            =   "frmDockDiagEdit.frx":040B
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   540
            Width           =   1470
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   19
            ItemData        =   "frmDockDiagEdit.frx":040D
            Left            =   1425
            List            =   "frmDockDiagEdit.frx":040F
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   945
            Width           =   1470
         End
         Begin VB.ComboBox cboBaseInfo 
            BackColor       =   &H8000000F&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   12
            ItemData        =   "frmDockDiagEdit.frx":0411
            Left            =   1425
            List            =   "frmDockDiagEdit.frx":0413
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   135
            Width           =   1470
         End
         Begin VB.ComboBox cboBaseInfo 
            BackColor       =   &H8000000F&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   13
            ItemData        =   "frmDockDiagEdit.frx":0415
            Left            =   4680
            List            =   "frmDockDiagEdit.frx":0417
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   135
            Width           =   1470
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   18
            ItemData        =   "frmDockDiagEdit.frx":0419
            Left            =   7920
            List            =   "frmDockDiagEdit.frx":041B
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   540
            Width           =   1470
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   16
            ItemData        =   "frmDockDiagEdit.frx":041D
            Left            =   4680
            List            =   "frmDockDiagEdit.frx":041F
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   540
            Width           =   1470
         End
         Begin VB.ComboBox cboBaseInfo 
            BackColor       =   &H8000000F&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   21
            ItemData        =   "frmDockDiagEdit.frx":0421
            Left            =   4680
            List            =   "frmDockDiagEdit.frx":0423
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   945
            Width           =   1470
         End
         Begin VB.TextBox txtInfo 
            BackColor       =   &H8000000F&
            Height          =   300
            Index           =   20
            Left            =   6690
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   1380
            Width           =   2775
         End
         Begin VB.TextBox txtDateInfo 
            BackColor       =   &H8000000F&
            Height          =   300
            Index           =   5
            Left            =   3480
            Locked          =   -1  'True
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   1380
            Width           =   1830
         End
         Begin VB.CommandButton cmdInfo 
            Caption         =   "��"
            Enabled         =   0   'False
            Height          =   240
            Index           =   20
            Left            =   9180
            TabIndex        =   2
            TabStop         =   0   'False
            ToolTipText     =   "ѡ��(*)"
            Top             =   1410
            Width           =   270
         End
         Begin MSMask.MaskEdBox mskDateInfo 
            Height          =   300
            Index           =   5
            Left            =   3480
            TabIndex        =   14
            TabStop         =   0   'False
            Tag             =   "####-##-## ##:##:##"
            Top             =   1380
            Visible         =   0   'False
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   529
            _Version        =   393216
            BackColor       =   -2147483633
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   19
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "####-##-## ##:##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label lblBaseInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��������ʬ��(&R)"
            Height          =   180
            Index           =   60
            Left            =   45
            TabIndex        =   25
            Top             =   1440
            Width           =   1350
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�������Ժ(&H)"
            Height          =   180
            Index           =   14
            Left            =   6720
            TabIndex        =   24
            Top             =   180
            Width           =   1170
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Ժ���Ժ(&I)"
            Height          =   180
            Index           =   15
            Left            =   225
            TabIndex        =   23
            Top             =   600
            Width           =   1170
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�����벡��(&L)"
            Height          =   180
            Index           =   18
            Left            =   6720
            TabIndex        =   22
            Top             =   600
            Width           =   1170
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ٴ��벡��(&M)"
            Height          =   180
            Index           =   19
            Left            =   225
            TabIndex        =   21
            Top             =   1005
            Width           =   1170
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ֻ��̶�(&F)"
            Height          =   180
            Index           =   12
            Left            =   405
            TabIndex        =   20
            Top             =   195
            Width           =   990
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����������(&G)"
            Height          =   180
            Index           =   13
            Left            =   3285
            TabIndex        =   19
            Top             =   195
            Width           =   1350
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��������Ժ(&J)"
            Height          =   180
            Index           =   16
            Left            =   3465
            TabIndex        =   18
            Top             =   600
            Width           =   1170
         End
         Begin VB.Label lblDateInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ʱ��(&N)"
            Height          =   180
            Index           =   5
            Left            =   2400
            TabIndex        =   17
            Top             =   1440
            Width           =   990
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ԭ��(&P)"
            Height          =   180
            Index           =   20
            Left            =   5670
            TabIndex        =   16
            Top             =   1440
            Width           =   990
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ٴ���ʬ��(&T)"
            Height          =   180
            Index           =   21
            Left            =   3450
            TabIndex        =   15
            Top             =   1005
            Width           =   1170
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsDiagXY 
         Height          =   1245
         Left            =   0
         TabIndex        =   28
         Top             =   360
         Width           =   9855
         _cx             =   17383
         _cy             =   2196
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
         BackColorFixed  =   14811105
         ForeColorFixed  =   -2147483630
         BackColorSel    =   4210752
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
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
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   25
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmDockDiagEdit.frx":0425
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
   Begin VB.PictureBox picZY 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   120
      ScaleHeight     =   3495
      ScaleWidth      =   9855
      TabIndex        =   31
      Top             =   360
      Width           =   9855
      Begin VB.Frame frmZY 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   0
         TabIndex        =   35
         Top             =   1560
         Visible         =   0   'False
         Width           =   9975
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   22
            ItemData        =   "frmDockDiagEdit.frx":06C5
            Left            =   1320
            List            =   "frmDockDiagEdit.frx":06C7
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   120
            Width           =   1395
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   23
            ItemData        =   "frmDockDiagEdit.frx":06C9
            Left            =   4185
            List            =   "frmDockDiagEdit.frx":06CB
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   120
            Width           =   1395
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�������Ժ(&A)"
            Height          =   180
            Index           =   22
            Left            =   120
            TabIndex        =   39
            Top             =   180
            Width           =   1170
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Ժ���Ժ(&B)"
            Height          =   180
            Index           =   23
            Left            =   2985
            TabIndex        =   38
            Top             =   180
            Width           =   1170
         End
      End
      Begin VB.OptionButton optDiag 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "���ݼ�����������(&4)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   3
         Left            =   6600
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   60
         Width           =   2010
      End
      Begin VB.OptionButton optDiag 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "������ϱ�׼����(&3)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   2
         Left            =   4560
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   60
         Value           =   -1  'True
         Width           =   2010
      End
      Begin VSFlex8Ctl.VSFlexGrid vsDiagZY 
         Height          =   1155
         Left            =   0
         TabIndex        =   34
         Top             =   360
         Width           =   10215
         _cx             =   18018
         _cy             =   2037
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
         BackColorFixed  =   14811105
         ForeColorFixed  =   -2147483630
         BackColorSel    =   4210752
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
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
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   25
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmDockDiagEdit.frx":06CD
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
   Begin VB.Frame fraSplit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   15
      Index           =   0
      Left            =   0
      TabIndex        =   29
      Top             =   240
      Width           =   9255
   End
   Begin MSComctlLib.TabStrip tabFunc 
      Height          =   345
      Left            =   56
      TabIndex        =   30
      Tag             =   "��ҽ���"
      Top             =   0
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   609
      Style           =   2
      TabFixedWidth   =   2027
      TabFixedHeight  =   617
      Separators      =   -1  'True
      TabMinWidth     =   0
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "��ҽ���"
            Key             =   "��ҽ���"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "��ҽ���"
            Key             =   "��ҽ���"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Image imgButtonNew 
      Height          =   240
      Left            =   480
      Picture         =   "frmDockDiagEdit.frx":098B
      Top             =   4200
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgButtonDel 
      Height          =   240
      Left            =   0
      Picture         =   "frmDockDiagEdit.frx":0F15
      Top             =   4200
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmDockDiagEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrOldOutWay As String          '�洢��Ժ��ʽ�ı���

Public Function ShowMe() As Boolean
'���أ� ShowDiagEdit= ��ȷ������ȡ��
    Show 1, gclsPros.MainForm
    ShowMe = gclsPros.IsOK
End Function

Private Sub cmdCancel_Click()
    Call CmdCancelClick
End Sub

Private Sub Form_Load()
    'סԺ��ҳ���
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim lngColWidth As Long
    On Error GoTo errH
    gclsPros.IsOK = False
    If gclsPros.PatiType = PF_סԺ Then
        Call InitControlData
        gclsPros.IsSigned = IsSignature
    Else
        strSQL = "Select 1 from ���˹Һż�¼ where Rownum<2 And ID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ���ﲡ���Ƿ�����ʷ��", gclsPros.��ҳID)
        If Not rsTmp.EOF Then
            gclsPros.Moved = False
        Else
            strSQL = "Select 1 from H���˹Һż�¼ where Rownum<2 And ID = [1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ���ﲡ���Ƿ�����ʷ��", gclsPros.��ҳID)
            gclsPros.Moved = Not rsTmp.EOF
        End If
    End If
    strSQL = "Select A.����,Nvl(A.·��״̬,-1) ·��״̬,A.�Ա�" & _
        " From ������ҳ A" & _
        " Where A.����ID=[1] And A.��ҳID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, gclsPros.����ID, gclsPros.��ҳID)
    If Not rsTmp.EOF Then
        gclsPros.InsureType = NVL(rsTmp!����, 0)
        '-1:δ����,0-�����ϵ���������1-ִ���У�2-����������3-�������
        gclsPros.PathState = Val(rsTmp!·��״̬ & "")
        gclsPros.Sex = rsTmp!�Ա� & ""
    End If
    
    If gclsPros.PatiType = PF_סԺ Then
        strSQL = "Select 1 From ���������¼  A Where  A.����ID=[1] And A.��ҳID=[2] "
        If gclsPros.Moved Then
            strSQL = Replace(strSQL, "���������¼", "H���������¼")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, gclsPros.����ID, gclsPros.��ҳID)
        gclsPros.Have���� = Not rsTmp.EOF
        '��ȡ�ٴ�·����Ϣ
        Call GetPatiPathInfo
    End If
    If gclsPros.DiagRowIDs = "" Then
        With gclsPros.DiagConn
            .Filter = "��ʶID=" & gclsPros.AplyMark
            .Sort = "���ID"
            Do While Not .EOF
                gclsPros.DiagRowIDs = gclsPros.DiagRowIDs & IIf(gclsPros.DiagRowIDs = "", "", ",") & !���id
                .MoveNext
            Loop
        End With
    End If
    '���ý���
    Call InitDaigSel
    '��������
    Call LoadData
    Call gclsPros.InitFacePara '���ý�������ؼ�״̬
    If gclsPros.BlnICDEleven Then
        gclsPros.CurrentForm.vsDiagXY.ColHidden(DI_��ע) = True
        gclsPros.CurrentForm.vsDiagXY.ColHidden(DI_ǰע��) = True
        gclsPros.CurrentForm.vsDiagXY.ColHidden(DI_��ע��) = True
        gclsPros.CurrentForm.vsDiagXY.ColHidden(DI_ICD����) = True
        gclsPros.CurrentForm.vsDiagXY.ColWidth(DI_��ϱ���) = gclsPros.CurrentForm.vsDiagXY.ColWidth(DI_��ϱ���) + gclsPros.CurrentForm.vsDiagXY.ColWidth(DI_ICD����) - IIf(gclsPros.PatiType = PF_����, 800, 300)
        gclsPros.CurrentForm.vsDiagXY.ColWidth(DI_�������) = gclsPros.CurrentForm.vsDiagXY.ColWidth(DI_�������) + gclsPros.CurrentForm.vsDiagXY.ColWidth(DI_��ע) - IIf(gclsPros.PatiType = PF_����, 800, 300)
        
        gclsPros.CurrentForm.vsDiagZY.ColHidden(DI_��ע) = True
        gclsPros.CurrentForm.vsDiagZY.ColHidden(DI_ǰע��) = True
        gclsPros.CurrentForm.vsDiagZY.ColHidden(DI_��ע��) = True
        gclsPros.CurrentForm.vsDiagZY.ColHidden(DI_��ҽ֤��) = True
        gclsPros.CurrentForm.vsDiagZY.ColWidth(DI_��ϱ���) = gclsPros.CurrentForm.vsDiagZY.ColWidth(DI_��ϱ���) + gclsPros.CurrentForm.vsDiagZY.ColWidth(DI_��ע) - IIf(gclsPros.PatiType = PF_����, 350, 300)
        gclsPros.CurrentForm.vsDiagZY.ColWidth(DI_�������) = gclsPros.CurrentForm.vsDiagZY.ColWidth(DI_�������) + gclsPros.CurrentForm.vsDiagZY.ColWidth(DI_��ҽ֤��) - IIf(gclsPros.PatiType = PF_����, 350, 300)
    Else
        gclsPros.CurrentForm.vsDiagXY.ColHidden(DI_��ע) = False
        gclsPros.CurrentForm.vsDiagXY.ColHidden(DI_ǰע��) = IIf(gclsPros.PatiType = PF_����, True, Not gclsPros.AddAnnotation)
        gclsPros.CurrentForm.vsDiagXY.ColHidden(DI_��ע��) = IIf(gclsPros.PatiType = PF_����, True, Not gclsPros.AddAnnotation)
        gclsPros.CurrentForm.vsDiagXY.ColHidden(DI_ICD����) = False
        gclsPros.CurrentForm.vsDiagXY.ColWidth(DI_��ϱ���) = IIf(gclsPros.PatiType = PF_����, 900, 850)
        gclsPros.CurrentForm.vsDiagXY.ColWidth(DI_�������) = IIf(gclsPros.PatiType = PF_����, 4000, 2500)
        
        gclsPros.CurrentForm.vsDiagZY.ColHidden(DI_��ע) = False
        gclsPros.CurrentForm.vsDiagZY.ColHidden(DI_ǰע��) = True
        gclsPros.CurrentForm.vsDiagZY.ColHidden(DI_��ע��) = True
        gclsPros.CurrentForm.vsDiagZY.ColHidden(DI_��ҽ֤��) = False
        gclsPros.CurrentForm.vsDiagZY.ColWidth(DI_��ϱ���) = IIf(gclsPros.PatiType = PF_����, 900, 850)
        gclsPros.CurrentForm.vsDiagXY.ColWidth(DI_�������) = IIf(gclsPros.PatiType = PF_����, 2900, 1900)
    End If
    If gclsPros.PatiType = PF_סԺ And gclsPros.FuncType <> f���뵥��� Then
        '��ȡԭ�еĳ�Ժ��ʽ
        mstrOldOutWay = vsDiagXY.TextMatrix(DT_��Ժ���XY, DI_��Ժ���)
        If gclsPros.Have��ҽ And mstrOldOutWay = "" Then
            mstrOldOutWay = vsDiagZY.TextMatrix(DT_��Ժ���XY, DI_��Ժ���)
        End If

        Call ChangeOutInfo(zlStr.NeedName(mstrOldOutWay))
        '������Ϸ���������ݲ�����
        Call CacheLoadDiagMatchData(GetDiagMatchData(gclsPros.����ID, gclsPros.��ҳID))
        '����ǩ�������ý���״̬
        Call SetControlState(gclsPros.IsSigned)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    fraSplit(0).Left = 0
    fraSplit(0).Top = tabFunc.Top + tabFunc.Height + 15
    fraSplit(0).Visible = tabFunc.Visible
    fraSplit(0).Width = Me.ScaleWidth
    If tabFunc.Visible Then
        picZY.Top = fraSplit(0).Top + fraSplit(0).Height + 15
    Else
        picZY.Top = tabFunc.Top
    End If
    picZY.Left = tabFunc.Left
    picZY.Height = vsDiagZY.Top + vsDiagZY.Height + 15
    picZY.Width = Me.ScaleWidth - picZY.Left * 3

    picXY.Left = picZY.Left
    picXY.Top = picZY.Top
    picXY.Height = vsDiagXY.Top + vsDiagXY.Height + 15
    picXY.Width = Me.ScaleWidth - picXY.Left * 3
End Sub

Private Sub optDiag_Click(Index As Integer)
    Call optDiagClick(Index)
End Sub

Private Sub picXY_Resize()
    Dim lngWidth As Long
    Dim lngColsWidth As Long
    Dim i As Long
    On Error Resume Next
    vsDiagXY.Height = picXY.ScaleHeight - vsDiagXY.Top - 120 - IIf(gclsPros.PatiType <> PF_����, frmXY.Height, 0)
    vsDiagXY.Width = picXY.ScaleWidth - vsDiagXY.Left * 2
    optDiag(1).Left = picXY.ScaleWidth - optDiag(1).Width - 120
    optDiag(0).Left = optDiag(1).Left - optDiag(0).Width - 120
    lngWidth = vsDiagXY.Width
    For i = 0 To vsDiagXY.Cols
        If Not vsDiagXY.ColHidden(i) And i <> DI_������� Then
            lngColsWidth = lngColsWidth + vsDiagXY.ColWidth(i)
        End If
    Next
    If lngWidth > lngColsWidth Then
        lngColsWidth = lngWidth - lngColsWidth
        vsDiagXY.ColWidth(DI_�������) = lngColsWidth - 400
    End If
    If gclsPros.PatiType = PF_סԺ Then frmXY.Top = vsDiagXY.Top + vsDiagXY.Height + 120
End Sub

Private Sub picZY_Resize()
    Dim lngWidth As Long
    Dim lngColsWidth As Long
    Dim i As Long
    On Error Resume Next
    vsDiagZY.Height = picZY.ScaleHeight - vsDiagZY.Top - 120 - IIf(gclsPros.PatiType <> PF_����, frmXY.Height, 0)
    vsDiagZY.Width = picZY.ScaleWidth - vsDiagZY.Left * 2
    optDiag(3).Left = picZY.ScaleWidth - optDiag(3).Width - 120
    optDiag(2).Left = optDiag(3).Left - optDiag(2).Width - 120
    lngWidth = vsDiagZY.Width
    For i = 0 To vsDiagZY.Cols
        If Not vsDiagZY.ColHidden(i) And i <> DI_������� Then
            lngColsWidth = lngColsWidth + vsDiagZY.ColWidth(i)
        End If
    Next
    If lngWidth > lngColsWidth Then
        lngColsWidth = lngWidth - lngColsWidth
        vsDiagZY.ColWidth(DI_�������) = lngColsWidth - 400
    End If
    If gclsPros.PatiType = PF_סԺ Then frmXY.Top = vsDiagXY.Top + vsDiagXY.Height + 120
End Sub

Private Sub Form_Unload(Cancel As Integer)
    gclsPros.IsOK = True
    Call FormUnLoad(Cancel)
End Sub

Private Sub tabFunc_Click()
    If tabFunc.Visible Then
        picXY.Visible = tabFunc.SelectedItem.Key = "��ҽ���"
        picZY.Visible = tabFunc.SelectedItem.Key = "��ҽ���"
    End If
End Sub

Private Sub vsDiagXY_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call DiagAfterEdit(vsDiagXY, Row, Col)
End Sub

Private Sub vsDiagXY_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call DiagAfterRowColChange(vsDiagXY, OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsDiagXY_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call DiagBeforeUserResize(vsDiagXY, Row, Col, Cancel)
End Sub

Private Sub vsDiagXY_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Call DiagCellButtonClick(vsDiagXY, Row, Col)
    Call EnableWindow(Me.hwnd, True)
End Sub

Private Sub vsDiagXY_Click()
    Call DiagClick(vsDiagXY)
    Call EnableWindow(Me.hwnd, True)
End Sub

Private Sub vsDiagXY_ComboDropDown(ByVal Row As Long, ByVal Col As Long)
    Call DiagComboDropDown(vsDiagXY, Row, Col)
End Sub

Private Sub vsDiagXY_DblClick()
    Call DiagDblClick(vsDiagXY)
    Call EnableWindow(Me.hwnd, True)
End Sub

Private Sub vsDiagXY_KeyDown(KeyCode As Integer, Shift As Integer)
    Call DiagKeyDown(vsDiagXY, KeyCode, Shift)
End Sub

Private Sub vsDiagXY_KeyPress(KeyAscii As Integer)
    Call DiagKeyPress(vsDiagXY, KeyAscii)
    Call EnableWindow(Me.hwnd, True)
End Sub

Private Sub vsDiagXY_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Call DiagKeyPressEdit(vsDiagXY, Row, Col, KeyAscii)
End Sub

Private Sub vsDiagXY_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    Call DiagSetupEditWindow(vsDiagXY, Row, Col, EditWindow, IsCombo)
End Sub

Private Sub vsDiagXY_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call DiagStartEdit(vsDiagXY, Row, Col, Cancel)
End Sub

Private Sub vsDiagXY_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call DiagValidateEdit(vsDiagXY, Row, Col, Cancel)
    Call EnableWindow(Me.hwnd, True)
End Sub

Private Sub vsDiagZY_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call DiagAfterEdit(vsDiagZY, Row, Col)
End Sub

Private Sub vsDiagZY_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call DiagAfterRowColChange(vsDiagZY, OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsDiagZY_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call DiagBeforeUserResize(vsDiagZY, Row, Col, Cancel)
End Sub

Private Sub vsDiagZY_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Call DiagCellButtonClick(vsDiagZY, Row, Col)
    Call EnableWindow(Me.hwnd, True)
End Sub

Private Sub vsDiagZY_Click()
    Call DiagClick(vsDiagZY)
    Call EnableWindow(Me.hwnd, True)
End Sub

Private Sub vsDiagZY_ComboDropDown(ByVal Row As Long, ByVal Col As Long)
    Call DiagComboDropDown(vsDiagZY, Row, Col)
End Sub

Private Sub vsDiagZY_DblClick()
    Call DiagDblClick(vsDiagZY)
    Call EnableWindow(Me.hwnd, True)
End Sub

Private Sub vsDiagZY_KeyDown(KeyCode As Integer, Shift As Integer)
    Call DiagKeyDown(vsDiagZY, KeyCode, Shift)
End Sub

Private Sub vsDiagZY_KeyPress(KeyAscii As Integer)
    Call DiagKeyPress(vsDiagZY, KeyAscii)
    Call EnableWindow(Me.hwnd, True)
End Sub

Private Sub vsDiagZY_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Call DiagKeyPressEdit(vsDiagZY, Row, Col, KeyAscii)
End Sub

Private Sub vsDiagZY_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    Call DiagSetupEditWindow(vsDiagZY, Row, Col, EditWindow, IsCombo)
End Sub

Private Sub vsDiagZY_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call DiagStartEdit(vsDiagZY, Row, Col, Cancel)
End Sub

Private Sub vsDiagZY_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call DiagValidateEdit(vsDiagZY, Row, Col, Cancel)
    Call EnableWindow(Me.hwnd, True)
End Sub

Private Sub LoadData()
    Dim rsTmp As ADODB.Recordset
    Dim strDiagFilter As String
    
    On Error GoTo errH
    '��ȡ���
    Set rsTmp = GetPatiDiagData(gclsPros.����ID, gclsPros.��ҳID, IIf(gclsPros.PatiType <> PF_����, 1, 0), , , gclsPros.Moved)
    rsTmp.Filter = "��¼��Դ=" & IIf(gclsPros.FuncType = f������ҳ, 4, 3)
    strDiagFilter = rsTmp.Filter
    rsTmp.Filter = "�������='D'"
    If rsTmp.EOF Then
        rsTmp.Filter = "�������='E'"
        If Not rsTmp.EOF Then
            gclsPros.BlnICDEleven = True
        Else
            If gclsPros.PatiType = PF_סԺ Then
                If gclsPros.InICDEleven Then
                    gclsPros.BlnICDEleven = True
                Else
                    gclsPros.BlnICDEleven = False
                End If
            ElseIf gclsPros.PatiType = PF_���� Then
                If gclsPros.OutICDEleven Then
                    gclsPros.BlnICDEleven = True
                Else
                    gclsPros.BlnICDEleven = False
                End If
            End If
        End If
    Else
        gclsPros.BlnICDEleven = False
    End If
    rsTmp.Filter = IIf(strDiagFilter <> "0", strDiagFilter, "")
    '2��������ҽ
    '   1-��ҽ�������;2-��ҽ��Ժ���;3-��Ժ���(�������);5-Ժ�ڸ�Ⱦ;6-�������;7-�����ж���;10-����֢
    Call CacheLoadVsDiagData(vsDiagXY, rsTmp, IIf(gclsPros.PatiType <> PF_����, "2", "1"), , -1)
    '3��������ҽ���
    '   11-��ҽ�������;12-��ҽ��Ժ���;13-��ҽ��Ժ���(��Ҫ��ϡ��������)
    If gclsPros.Have��ҽ Then
        Call CacheLoadVsDiagData(vsDiagZY, rsTmp, IIf(gclsPros.PatiType <> PF_����, "12", "11"), , -1)
    End If
    '����ȷ�ϴ�Ⱦ�����
    If gclsPros.IsComfirmInfect Then
        vsDiagXY.ColHidden(DI_����) = True
        vsDiagXY.ColWidth(DI_�������) = vsDiagXY.ColWidth(DI_�������) + vsDiagXY.ColWidth(DI_����)
        Call LoadInfeciousDiseases
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadInfeciousDiseases()
'���ܣ�����ȷ�ϴ�Ⱦ�����
    Dim lngStart As Long, dtType As DiagType
    Dim blnAdd As Boolean, LngRow As Long, j As Long
    Dim strSQL As String, str�Ա� As String
    Dim rsInput As ADODB.Recordset
    
    On Error GoTo errH
    If gclsPros.Sex Like "*��*" Then
        str�Ա� = "��"
    ElseIf gclsPros.Sex Like "*Ů*" Then
        str�Ա� = "Ů"
    End If
    With gclsPros.DiagConn
        .Filter = "����=1"
        dtType = IIf(gclsPros.PatiType = PF_����, DT_�������XY, DT_��Ժ���XY)
        lngStart = vsDiagXY.FindRow(dtType, , DI_��Ϸ���, , True)
        Do While Not .EOF
            blnAdd = True: LngRow = lngStart
            '���ڼ���ID�����ID�Ž��д���
            If Val(!���Ŀ¼ID & "") <> 0 Or Val(!����Ŀ¼ID & "") <> 0 Then
                For j = LngRow To vsDiagXY.Rows - 1
                    If Val(vsDiagXY.TextMatrix(j, DI_��Ϸ���)) = dtType Then
                        LngRow = j
                        If vsDiagXY.TextMatrix(j, DI_�������) = "" Then Exit For
                        If Val(vsDiagXY.TextMatrix(j, DI_����ID)) = Val(!����Ŀ¼ID & "") And Val(!����Ŀ¼ID & "") <> 0 Then
                            blnAdd = False: Exit For
                        ElseIf Val(vsDiagXY.TextMatrix(j, DI_���ID)) = Val(!���Ŀ¼ID & "") And Val(!���Ŀ¼ID & "") <> 0 Then
                            blnAdd = False: Exit For
                        End If
                    Else
                        Exit For
                    End If
                Next
                If blnAdd Then
                    If Val(!���Ŀ¼ID & "") <> 0 And (gclsPros.DiagInputXY = 0 Or Val(!����Ŀ¼ID & "") = 0) Then
                        strSQL = "Select Distinct a.Id, a.��Ŀid, a.����, b.���, b.����, Null ����id, Null ��������, a.����, a.˵��, Null ����, a.����, a.��Ч����, a.����, a.�Ƿ���," & vbNewLine & _
                                    "                b.���� ��������, b.Id ����id, b.��� �������, a.���id" & vbNewLine & _
                                    "From (Select a.Id, a.Id ��Ŀid, a.����, Null ���, Null ����, Null ����id, Null ��������, a.����, a.˵��, a.����, b.����, 0 ��Ч����, 0 ����, 0 �Ƿ���," & vbNewLine & _
                                    "              Max(d.����id) ����id, a.Id ���id" & vbNewLine & _
                                    "       From �������Ŀ¼ a, ������ϱ��� b, ������϶��� d" & vbNewLine & _
                                    "       Where a.Id = [1] And a.Id = b.���id And a.Id = d.���id(+) And a.��� = 1 And b.���� = [4]" & vbNewLine & _
                                    " And (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd'))" & vbNewLine & _
                                    "       Group By a.Id, a.����, a.����, a.˵��, a.����, b.����) a, ��������Ŀ¼ b, ������Ͽ��� c, ������Ͽ��� d" & vbNewLine & _
                                    "Where a.����id = b.Id(+) And c.���id(+) = a.Id And d.���id(+) = a.Id And c.����id(+) = [5] And d.��Աid(+) = [6]" & vbNewLine & _
                                    "Order By a.����"
                    Else
                        strSQL = "Select Distinct a.Id, a.��Ŀid, a.����, a.���, a.����, a.����id, a.��������, a.����, a.˵��, a.����, a.����id, a.����, a.��Ч����, a.����, a.�Ƿ���," & vbNewLine & _
                                    "                a.��������, a.����id, a.�������, a.���id" & vbNewLine & _
                                    "From (Select a.Id, a.Id ��Ŀid, a.����, a.���, a.����, Null ����id, Null ��������, a.����, a.˵��, Null ����, a.����id, a.����� ����, a.��Ч����, a.����," & vbNewLine & _
                                    "              c.�Ƿ���, a.���� ��������, a.Id ����id, a.��� �������, Max(b.���id) ���id" & vbNewLine & _
                                    "       From ��������Ŀ¼ a, ������϶��� b, ����������� c" & vbNewLine & _
                                    "       Where a.Id = [2] And a.Id = b.����id(+) And a.����id = c.Id(+) And a.���='D' And" & vbNewLine & _
                                    IIf(str�Ա� <> "", "  (A.�Ա�����=[3] Or A.�Ա����� is NULL) And ", " ") & " (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
                                    "       Group By a.Id, a.����, a.���, a.����, a.����, a.˵��, a.����id, a.�����, a.��Ч����, a.����, a.���, c.�Ƿ���) a, ����������� c, ����������� d" & vbNewLine & _
                                    "Where c.����id(+) = a.Id And d.����id(+) = a.Id And c.����id(+) = [5] And d.��Աid(+) = [6]" & vbNewLine & _
                                    "Order By a.����"
                    End If
                    Set rsInput = zlDatabase.OpenSQLRecord(strSQL, "ȷ�ϴ�Ⱦ��", Val(!���Ŀ¼ID & ""), Val(!����Ŀ¼ID & ""), str�Ա�, gclsPros.BriefCode + 1, gclsPros.��Ժ����ID, UserInfo.ID)
                    If rsInput.RecordCount > 0 Then
                        '������
                         If vsDiagXY.TextMatrix(LngRow, DI_�������) <> "" Then
                             LngRow = LngRow + 1: vsDiagXY.AddItem "", LngRow
                             vsDiagXY.TextMatrix(LngRow, DI_��Ϸ���) = dtType
                         End If
                         Call SetDiagInput(vsDiagXY, LngRow, rsInput)
                    End If
                End If
            End If
            .MoveNext
        Loop
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub InitDaigSel(Optional ByVal blnAfterLoad As Boolean)
'��ʼ�����ѡ�����
'������blnAfterLoad=�Ƿ����ݼ���֮���ʼ��
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim lngColWidth As Long, LngRow As Long
    
    tabFunc.Visible = gclsPros.Have��ҽ
    If tabFunc.Visible Then
        picXY.Visible = True
        picXY.ZOrder
        picZY.Visible = False
    End If
    frmXY.Visible = False
    frmZY.Visible = False
    Call InitTableDiag
End Sub

Private Function IsSignature() As Boolean
'���ܣ���ȡ��ǰ���˵�ҽʦ��ǩ�����
'���أ������Ƿ���ǩ��
    Dim rsTmp As ADODB.Recordset
    Dim intCurr As Integer, intHave As Integer
    Dim strSQL As String, blnReadOnly As Boolean
    Dim i As Integer
    '˵����arrInfos �����Ԫ��һһ��Ӧ����Ա����ӵ͵���
    Dim arrInfos() As Variant '����ǩ������Ϣ��
    Dim arrSgnIdxs() As Variant '����ǩ������Ϣ��
    Dim arrName() As Variant
    On Error GoTo errH
    blnReadOnly = False: intCurr = -1: intHave = -1
    arrSgnIdxs = Array("סԺҽʦǩ��", "����ҽʦǩ��", "����ҽʦǩ��", "������ǩ��")
    arrInfos = Array("סԺҽʦ", "����ҽʦ", "����ҽʦ", "������")
    arrName = Array("", "", "", "")
    
    strSQL = "select 'סԺҽʦ' as ��Ϣ��, A.סԺҽʦ as ��Ϣֵ from ������ҳ A where a.����id = [1]  And a.��ҳid = [2]" & vbNewLine & _
             "union all" & vbNewLine & _
             "select A.��Ϣ�� , A.��Ϣֵ from ������ҳ�ӱ� A where  A.����id = [1] And A.��ҳid = [2]"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ����ҽʦ���", gclsPros.����ID, gclsPros.��ҳID)
    
    For i = LBound(arrInfos) To UBound(arrInfos)
        rsTmp.Filter = "��Ϣ��='" & arrInfos(i) & "'"
        If Not rsTmp.EOF Then
            arrName(i) = rsTmp!��Ϣֵ & ""
        End If
    Next
    For i = LBound(arrName) To UBound(arrName)
        If arrName(i) = UserInfo.���� Then
            intCurr = i
        End If
        gclsPros.AuxiInfo.Filter = "��Ϣ��='" & arrSgnIdxs(i) & "'"
        If Not gclsPros.AuxiInfo.EOF Then
            intHave = i
        End If
    Next

    '�����ǰ��Աǩ�����𲻸�����ǩ�������򲻿ɱ༭
    If intCurr <= intHave And intHave >= 0 Then
        blnReadOnly = True
    End If
    IsSignature = blnReadOnly
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function InitControlData() As Boolean
'���ܣ���ʼ������ؼ�����
    Dim i As Integer
    On Error GoTo errH
    If gclsPros.PatiType = PF_סԺ Then
        Call SetCboFromList(Array("��", "��"), Array(cboBaseInfo(BCC_��������ʬ��)), 0)
        Call SetCboFromRec(Array(BCC_�ֻ��̶�, BCC_����������), 0)
        Call SetCboFromList(Array("0-δ��", "1-����", "2-������", "3-���϶�"), Array(cboBaseInfo(BCC_�������ԺXY), cboBaseInfo(BCC_��������Ժ), cboBaseInfo(BCC_��Ժ���ԺXY), cboBaseInfo(BCC_�����벡��), cboBaseInfo(BCC_�ٴ��벡��), _
         cboBaseInfo(BCC_�������ԺZY), cboBaseInfo(BCC_��Ժ���ԺZY), cboBaseInfo(BCC_�ٴ���ʬ��)))
    End If
       
    Set gclsPros.PatiInfo = GetPatiMainInfoData(gclsPros.����ID, gclsPros.��ҳID)
    '���ز�����Ϣ
    If Not gclsPros.PatiInfo.EOF Then
        For i = 0 To gclsPros.PatiInfo.Fields.Count - 1
             Call SetCtrlValues(UCase(gclsPros.PatiInfo.Fields(i).Name & ""), gclsPros.PatiInfo.Fields(i).Value & "", , True)
        Next
    End If

    Set gclsPros.AuxiInfo = GetPatiAuxiInfoData(gclsPros.����ID, gclsPros.��ҳID)   '�ӱ���Ϣ
    If Not gclsPros.AuxiInfo.EOF Then
        gclsPros.AuxiInfo.MoveFirst
        For i = 1 To gclsPros.AuxiInfo.RecordCount
            Call SetCtrlValues(gclsPros.AuxiInfo!��Ϣ�� & "", gclsPros.AuxiInfo!��Ϣֵ & "", gclsPros.AuxiInfo!���� & "")
            gclsPros.AuxiInfo.MoveNext
        Next
    End If
    InitControlData = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SetControlState(ByVal blnState As Boolean) As Boolean
'���ܣ����ݵ�ǰ���˵�ҽʦ��ǩ�������ȷ��ǩ�����������ݵĿɱ༭��
    Dim objControl As Object
    Dim strName As String
    Dim vsTmp As VSFlexGrid
    On Error GoTo errH
    With gclsPros.CurrentForm
        If blnState Then
            For Each objControl In .Controls
                strName = objControl.Name
                If strName = "cboBaseInfo" Or strName = "chkInfo" Or strName = "cmdInfo" Or strName = "mskDateInfo" Or strName = "txtInfo" Then
                    Call SetCtrlLocked(objControl, blnState)
                ElseIf strName = "vsDiagXY" Or strName = "vsDiagZY" Then
                    Set vsTmp = objControl
                    vsTmp.BackColorBkg = &H8000000F
                    vsTmp.Cell(flexcpBackColor, 0, DI_�������, vsTmp.Rows - 1, vsTmp.Cols - 1) = &H8000000F
                    vsTmp.Cell(flexcpBackColor, 0, DI_����, vsTmp.Rows - 1, DI_����) = &H80000005
                End If
            Next
        End If
    End With
    SetControlState = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cboBaseInfo_Click(Index As Integer)
    Call CboBaseInfoClick(Index)
End Sub

Private Sub cboBaseInfo_GotFocus(Index As Integer)
    Call CboBaseInfoGotFocus(Index)
End Sub

Private Sub cboBaseInfo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CboBaseInfoKeyDown(Index, KeyCode, Shift)
End Sub

Private Sub cboBaseInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    Call CboBaseInfoKeyPress(Index, KeyAscii)
End Sub

Private Sub cboBaseInfo_Validate(Index As Integer, Cancel As Boolean)
    Call cboBaseInfoValidate(Index, Cancel)
End Sub

Private Sub cmdInfo_Click(Index As Integer)
    Call CmdInfoClick(Index)
End Sub

Private Sub PopPatiOtherSQL(ByRef arrSQL As Variant)
'���ܣ�����Ժ��ʽ�����SQL��������
    Dim strNewOutWay As String
    Dim strValue As String
    On Error GoTo errH
    strNewOutWay = vsDiagXY.TextMatrix(DT_��Ժ���XY, DI_��Ժ���)
    If gclsPros.Have��ҽ And strNewOutWay = "" Then
       strNewOutWay = vsDiagZY.TextMatrix(DT_��Ժ���XY, DI_��Ժ���)
    End If
    
    If (mstrOldOutWay <> "����" And strNewOutWay = "����") Or (mstrOldOutWay = "����" And strNewOutWay <> "����") Then
        strNewOutWay = IIf(strNewOutWay = "����", "����", "����")
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_������ҳ_��ҳ����EX(" & gclsPros.����ID & "," & gclsPros.��ҳID & ",'��Ժ��ʽ','" & strNewOutWay & "')"
        If strNewOutWay = "����" Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_������ҳ�ӱ�_��ҳ����(" & gclsPros.����ID & "," & gclsPros.��ҳID & ",'��Ժת��', NULL)"
        End If
    End If
    
    gclsPros.MainInfoRec.Filter = "�Ƿ�ı�=1 and ��Ϣ��='ʬ���־'"
    If Not gclsPros.MainInfoRec.EOF Then
        strValue = gclsPros.MainInfoRec!��Ϣ��ֵ & ""
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_������ҳ_��ҳ����EX(" & gclsPros.����ID & "," & gclsPros.��ҳID & ",'ʬ���־','" & strValue & "')"
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub imgMore_Click()
    Call ImgMoreClick
End Sub

Private Sub vsDiagXY_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DiagMouseMove(vsDiagXY, Button, Shift, X, Y)
End Sub






