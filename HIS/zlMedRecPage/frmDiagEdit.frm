VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDiagEdit 
   BackColor       =   &H80000004&
   Caption         =   "���ѡ�񼰱༭"
   ClientHeight    =   7335
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10440
   Icon            =   "frmDiagEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleMode       =   0  'User
   ScaleWidth      =   10653.07
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picXY 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6015
      Left            =   120
      ScaleHeight     =   6015
      ScaleWidth      =   10335
      TabIndex        =   3
      Top             =   480
      Width           =   10335
      Begin VB.Frame frmXY 
         BorderStyle     =   0  'None
         Height          =   1815
         Left            =   0
         TabIndex        =   28
         Top             =   4200
         Width           =   10095
         Begin VB.CommandButton cmdInfo 
            Caption         =   "��"
            Enabled         =   0   'False
            Height          =   240
            Index           =   20
            Left            =   9180
            TabIndex        =   18
            TabStop         =   0   'False
            ToolTipText     =   "ѡ��(*)"
            Top             =   1410
            Width           =   270
         End
         Begin VB.TextBox txtDateInfo 
            BackColor       =   &H8000000F&
            Height          =   300
            Index           =   5
            Left            =   3480
            Locked          =   -1  'True
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   1380
            Width           =   1830
         End
         Begin VB.TextBox txtInfo 
            BackColor       =   &H8000000F&
            Height          =   300
            Index           =   20
            Left            =   6690
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   1380
            Width           =   2775
         End
         Begin VB.ComboBox cboBaseInfo 
            BackColor       =   &H8000000F&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   21
            ItemData        =   "frmDiagEdit.frx":6852
            Left            =   4680
            List            =   "frmDiagEdit.frx":6854
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   945
            Width           =   1470
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   16
            ItemData        =   "frmDiagEdit.frx":6856
            Left            =   4680
            List            =   "frmDiagEdit.frx":6858
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   540
            Width           =   1470
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   18
            ItemData        =   "frmDiagEdit.frx":685A
            Left            =   7920
            List            =   "frmDiagEdit.frx":685C
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   540
            Width           =   1470
         End
         Begin VB.ComboBox cboBaseInfo 
            BackColor       =   &H8000000F&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   13
            ItemData        =   "frmDiagEdit.frx":685E
            Left            =   4680
            List            =   "frmDiagEdit.frx":6860
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   135
            Width           =   1470
         End
         Begin VB.ComboBox cboBaseInfo 
            BackColor       =   &H8000000F&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   12
            ItemData        =   "frmDiagEdit.frx":6862
            Left            =   1425
            List            =   "frmDiagEdit.frx":6864
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   135
            Width           =   1470
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   19
            ItemData        =   "frmDiagEdit.frx":6866
            Left            =   1425
            List            =   "frmDiagEdit.frx":6868
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   945
            Width           =   1470
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   15
            ItemData        =   "frmDiagEdit.frx":686A
            Left            =   1425
            List            =   "frmDiagEdit.frx":686C
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   540
            Width           =   1470
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   14
            ItemData        =   "frmDiagEdit.frx":686E
            Left            =   7920
            List            =   "frmDiagEdit.frx":6870
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   135
            Width           =   1470
         End
         Begin VB.ComboBox cboBaseInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   300
            Index           =   60
            ItemData        =   "frmDiagEdit.frx":6872
            Left            =   1425
            List            =   "frmDiagEdit.frx":6874
            Locked          =   -1  'True
            TabIndex        =   15
            TabStop         =   0   'False
            Text            =   "cboBaseInfo"
            Top             =   1380
            Width           =   675
         End
         Begin MSMask.MaskEdBox mskDateInfo 
            Height          =   300
            Index           =   5
            Left            =   3480
            TabIndex        =   29
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
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ٴ���ʬ��(&T)"
            Height          =   180
            Index           =   21
            Left            =   3450
            TabIndex        =   40
            Top             =   1005
            Width           =   1170
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ԭ��(&P)"
            Height          =   180
            Index           =   20
            Left            =   5670
            TabIndex        =   39
            Top             =   1440
            Width           =   990
         End
         Begin VB.Label lblDateInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ʱ��(&N)"
            Height          =   180
            Index           =   5
            Left            =   2400
            TabIndex        =   38
            Top             =   1440
            Width           =   990
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��������Ժ(&J)"
            Height          =   180
            Index           =   16
            Left            =   3465
            TabIndex        =   37
            Top             =   600
            Width           =   1170
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����������(&G)"
            Height          =   180
            Index           =   13
            Left            =   3285
            TabIndex        =   36
            Top             =   195
            Width           =   1350
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ֻ��̶�(&F)"
            Height          =   180
            Index           =   12
            Left            =   405
            TabIndex        =   35
            Top             =   195
            Width           =   990
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ٴ��벡��(&M)"
            Height          =   180
            Index           =   19
            Left            =   225
            TabIndex        =   34
            Top             =   1005
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
            TabIndex        =   33
            Top             =   600
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
            TabIndex        =   32
            Top             =   600
            Width           =   1170
         End
         Begin VB.Label lblBaseInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�������Ժ(&H)"
            Height          =   180
            Index           =   14
            Left            =   6720
            TabIndex        =   31
            Top             =   180
            Width           =   1170
         End
         Begin VB.Label lblBaseInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��������ʬ��(&R)"
            Height          =   180
            Index           =   60
            Left            =   45
            TabIndex        =   30
            Top             =   1440
            Width           =   1350
         End
      End
      Begin VB.OptionButton optDiag 
         BackColor       =   &H00EFF0E0&
         Caption         =   "������ϱ�׼����(&1)"
         ForeColor       =   &H00004000&
         Height          =   180
         Index           =   0
         Left            =   4680
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   60
         Value           =   -1  'True
         Width           =   2010
      End
      Begin VB.OptionButton optDiag 
         BackColor       =   &H00EFF0E0&
         Caption         =   "���ݼ�����������(&2)"
         ForeColor       =   &H00004000&
         Height          =   180
         Index           =   1
         Left            =   6720
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   60
         Width           =   2010
      End
      Begin VSFlex8Ctl.VSFlexGrid vsDiagXY 
         Height          =   3795
         Left            =   0
         TabIndex        =   6
         Top             =   360
         Width           =   10095
         _cx             =   17806
         _cy             =   6694
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
         Rows            =   9
         Cols            =   25
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmDiagEdit.frx":6876
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
      Height          =   6015
      Left            =   120
      ScaleHeight     =   6015
      ScaleWidth      =   10335
      TabIndex        =   19
      Top             =   480
      Width           =   10335
      Begin VB.Frame frmZY 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   0
         TabIndex        =   41
         Top             =   4320
         Width           =   9975
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   22
            ItemData        =   "frmDiagEdit.frx":6B5C
            Left            =   1320
            List            =   "frmDiagEdit.frx":6B5E
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   120
            Width           =   1395
         End
         Begin VB.ComboBox cboBaseInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   23
            ItemData        =   "frmDiagEdit.frx":6B60
            Left            =   4185
            List            =   "frmDiagEdit.frx":6B62
            Style           =   2  'Dropdown List
            TabIndex        =   24
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
            TabIndex        =   43
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
            TabIndex        =   42
            Top             =   180
            Width           =   1170
         End
      End
      Begin VB.OptionButton optDiag 
         BackColor       =   &H00EFF0E0&
         Caption         =   "������ϱ�׼����(&3)"
         ForeColor       =   &H00004000&
         Height          =   180
         Index           =   2
         Left            =   4560
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   60
         Value           =   -1  'True
         Width           =   2010
      End
      Begin VB.OptionButton optDiag 
         BackColor       =   &H00EFF0E0&
         Caption         =   "���ݼ�����������(&4)"
         ForeColor       =   &H00004000&
         Height          =   180
         Index           =   3
         Left            =   6600
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   60
         Width           =   2010
      End
      Begin VSFlex8Ctl.VSFlexGrid vsDiagZY 
         Height          =   3915
         Left            =   0
         TabIndex        =   22
         Top             =   360
         Width           =   10215
         _cx             =   18018
         _cy             =   6906
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
         Rows            =   5
         Cols            =   25
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmDiagEdit.frx":6B64
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
      Index           =   1
      Left            =   0
      TabIndex        =   27
      Top             =   6600
      Width           =   10095
   End
   Begin VB.Frame fraSplit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   15
      Index           =   0
      Left            =   120
      TabIndex        =   26
      Top             =   360
      Width           =   9975
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   10440
      TabIndex        =   0
      Top             =   6735
      Width           =   10440
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   8400
         TabIndex        =   2
         Top             =   150
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   7080
         TabIndex        =   1
         Top             =   150
         Width           =   1100
      End
      Begin VB.Image imgButtonNew 
         Height          =   240
         Left            =   720
         Picture         =   "frmDiagEdit.frx":6E22
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgButtonDel 
         Height          =   240
         Left            =   0
         Picture         =   "frmDiagEdit.frx":73AC
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin MSComctlLib.TabStrip tabFunc 
      Height          =   345
      Left            =   176
      TabIndex        =   25
      Tag             =   "��ҽ���"
      Top             =   120
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
End
Attribute VB_Name = "frmDiagEdit"
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

Private Sub cmdOK_Click()
    If CheckData() Then
        Call SaveData
        gclsPros.IsOK = True
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    'סԺ��ҳ���
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim lngColWidth As Long
    On Error GoTo errH
    gclsPros.IsOK = False
    If gclsPros.PatiType = PF_סԺ Then
        Call InitControlData
        gclsPros.IsSigned = IsSignature
    Else
        strSql = "Select 1 from ���˹Һż�¼ where Rownum<2 And ID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ѯ���ﲡ���Ƿ�����ʷ��", gclsPros.��ҳID)
        If Not rsTmp.EOF Then
            gclsPros.Moved = False
        Else
            strSql = "Select 1 from H���˹Һż�¼ where Rownum<2 And ID = [1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ѯ���ﲡ���Ƿ�����ʷ��", gclsPros.��ҳID)
            gclsPros.Moved = Not rsTmp.EOF
        End If
    End If
    Call gclsPros.InitFacePara '���ý�������ؼ�״̬
    strSql = "Select A.����,Nvl(A.·��״̬,-1) ·��״̬,A.�Ա�" & _
        " From ������ҳ A" & _
        " Where A.����ID=[1] And A.��ҳID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, gclsPros.����ID, gclsPros.��ҳID)
    If Not rsTmp.EOF Then
        gclsPros.InsureType = Nvl(rsTmp!����, 0)
        '-1:δ����,0-�����ϵ���������1-ִ���У�2-����������3-�������
        gclsPros.PathState = Val(rsTmp!·��״̬ & "")
        gclsPros.Sex = rsTmp!�Ա� & ""
    End If
    
    If gclsPros.PatiType = PF_סԺ Then
        strSql = "Select 1 From ���������¼  A Where  A.����ID=[1] And A.��ҳID=[2] "
        If gclsPros.Moved Then
            strSql = Replace(strSql, "���������¼", "H���������¼")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, gclsPros.����ID, gclsPros.��ҳID)
        gclsPros.Have���� = Not rsTmp.EOF
        '��ȡ�ٴ�·����Ϣ
        Call GetPatiPathInfo
    End If
    If gclsPros.DiagRowIDs = "" Then
        With gclsPros.DiagConn
            .Filter = "��ʶID=" & gclsPros.AplyMark
            .Sort = "���ID"
            Do While Not .EOF
                gclsPros.DiagRowIDs = gclsPros.DiagRowIDs & IIf(gclsPros.DiagRowIDs = "", "", ",") & !���ID
                .MoveNext
            Loop
        End With
    End If
    '���ý���
    Call InitDaigSel
    '��������
    Call LoadData
    If gclsPros.PatiType = PF_סԺ Then
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
    If Me.Width > 20000 Then
        Me.Width = 20000
    End If
    If Me.Height > 12000 Then
        Me.Height = 12000
    End If
    If Me.Width < 6000 Then
        Me.Width = 6000
    End If
    If Me.Height < 5000 Then
        Me.Height = 5000
    End If
    picBottom.Height = 600
    picBottom.Top = Me.ScaleHeight - 600
    fraSplit(0).Left = 0:    fraSplit(1).Left = 0
    fraSplit(0).Top = tabFunc.Top + tabFunc.Height + 15
    fraSplit(0).Visible = tabFunc.Visible
    fraSplit(1).Top = picBottom.Top - 15 - fraSplit(1).Height
    fraSplit(0).Width = Me.ScaleWidth:  fraSplit(1).Width = fraSplit(0).Width
    If tabFunc.Visible Then
        picZY.Top = fraSplit(0).Top + fraSplit(0).Height + 15
    Else
        picZY.Top = tabFunc.Top
    End If
    picZY.Left = tabFunc.Left
    picZY.Height = fraSplit(1).Top - 15 - picZY.Top
    picZY.Width = Me.ScaleWidth - picZY.Left * 3
    
    picXY.Left = picZY.Left
    picXY.Top = picZY.Top
    picXY.Height = picZY.Height
    picXY.Width = picZY.Width
    
    cmdCancel.Left = picXY.Left + picXY.Width - cmdCancel.Width - 60
    cmdOK.Left = cmdCancel.Left - 60 - cmdOK.Width
End Sub

Private Sub optDiag_Click(Index As Integer)
    Call optDiagClick(Index)
End Sub

Private Sub picXY_Resize()
    On Error Resume Next
    vsDiagXY.Height = picXY.ScaleHeight - vsDiagXY.Top - 120 - IIf(gclsPros.PatiType <> PF_����, frmXY.Height, 0)
    vsDiagXY.Width = picXY.ScaleWidth - vsDiagXY.Left * 2
    optDiag(1).Left = picXY.ScaleWidth - optDiag(1).Width - 120
    optDiag(0).Left = optDiag(1).Left - optDiag(0).Width - 120
    If gclsPros.PatiType = PF_סԺ Then frmXY.Top = vsDiagXY.Top + vsDiagXY.Height + 120
End Sub

Private Sub picZY_Resize()
    On Error Resume Next
    vsDiagZY.Height = picZY.ScaleHeight - vsDiagZY.Top - 120 - IIf(gclsPros.PatiType <> PF_����, frmXY.Height, 0)
    vsDiagZY.Width = picZY.ScaleWidth - vsDiagZY.Left * 2
    optDiag(3).Left = picZY.ScaleWidth - optDiag(3).Width - 120
    optDiag(2).Left = optDiag(3).Left - optDiag(2).Width - 120
    If gclsPros.PatiType = PF_סԺ Then frmXY.Top = vsDiagXY.Top + vsDiagXY.Height + 120
End Sub

Private Sub Form_Unload(Cancel As Integer)
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
End Sub

Private Sub vsDiagXY_Click()
    Call DiagClick(vsDiagXY)
End Sub

Private Sub vsDiagXY_ComboDropDown(ByVal Row As Long, ByVal Col As Long)
    Call DiagComboDropDown(vsDiagXY, Row, Col)
End Sub

Private Sub vsDiagXY_DblClick()
    Call DiagDblClick(vsDiagXY)
End Sub

Private Sub vsDiagXY_KeyDown(KeyCode As Integer, Shift As Integer)
    Call DiagKeyDown(vsDiagXY, KeyCode, Shift)
End Sub

Private Sub vsDiagXY_KeyPress(KeyAscii As Integer)
    Call DiagKeyPress(vsDiagXY, KeyAscii)
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
End Sub

Private Sub vsDiagZY_Click()
    Call DiagClick(vsDiagZY)
End Sub

Private Sub vsDiagZY_ComboDropDown(ByVal Row As Long, ByVal Col As Long)
    Call DiagComboDropDown(vsDiagZY, Row, Col)
End Sub

Private Sub vsDiagZY_DblClick()
    Call DiagDblClick(vsDiagZY)
End Sub

Private Sub vsDiagZY_KeyDown(KeyCode As Integer, Shift As Integer)
    Call DiagKeyDown(vsDiagZY, KeyCode, Shift)
End Sub

Private Sub vsDiagZY_KeyPress(KeyAscii As Integer)
    Call DiagKeyPress(vsDiagZY, KeyAscii)
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
End Sub

Private Sub LoadData()
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errH
    '��ȡ���
    Set rsTmp = GetPatiDiagData(gclsPros.����ID, gclsPros.��ҳID, IIf(gclsPros.PatiType <> PF_����, 1, 0), , , gclsPros.Moved)
    rsTmp.Filter = "��¼��Դ=" & IIf(gclsPros.FuncType = f������ҳ, 4, 3)
    '2��������ҽ
    '   1-��ҽ�������;2-��ҽ��Ժ���;3-��Ժ���(�������);5-Ժ�ڸ�Ⱦ;6-�������;7-�����ж���;10-����֢
    Call CacheLoadVsDiagData(vsDiagXY, rsTmp, IIf(gclsPros.PatiType <> PF_����, "1,2,3,5,6,7,10", "1"), , -1)
    '3��������ҽ���
    '   11-��ҽ�������;12-��ҽ��Ժ���;13-��ҽ��Ժ���(��Ҫ��ϡ��������)
    If gclsPros.Have��ҽ Then
        Call CacheLoadVsDiagData(vsDiagZY, rsTmp, IIf(gclsPros.PatiType <> PF_����, "11,12,13", "11"), , -1)
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
    Dim strSql As String, str�Ա� As String
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
                        strSql = "Select Distinct a.Id, a.��Ŀid, a.����, b.���, b.����, Null ����id, Null ��������, a.����, a.˵��, Null ����, a.����, a.��Ч����, a.����, a.�Ƿ���," & vbNewLine & _
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
                        strSql = "Select Distinct a.Id, a.��Ŀid, a.����, a.���, a.����, a.����id, a.��������, a.����, a.˵��, a.����, a.����id, a.����, a.��Ч����, a.����, a.�Ƿ���," & vbNewLine & _
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
                    Set rsInput = zlDatabase.OpenSQLRecord(strSql, "ȷ�ϴ�Ⱦ��", Val(!���Ŀ¼ID & ""), Val(!����Ŀ¼ID & ""), str�Ա�, gclsPros.BriefCode + 1, gclsPros.��Ժ����ID, UserInfo.ID)
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

Private Sub SaveData()
    Dim arrSQL() As Variant
    Dim blnTrans As Boolean
    Dim i As Long
    Dim datCur As Date
    arrSQL = Array()
    If gclsPros.InfosChange Then
        datCur = zlDatabase.Currentdate
        Call PopPatiDiagSQL(arrSQL, datCur)
        If gclsPros.PatiType = PF_סԺ Then
            '�ӱ���Ϣ����
            Call PopPatiAuxiSQL(arrSQL, gclsPros.Is��ʿվ)
            '��Ϸ����������
            Call PopDiagMatchSQL(arrSQL)
            '���˵ĳ�Ժ��ʽ��ʬ����Ϣ�ı���
            Call PopPatiOtherSQL(arrSQL)
        End If
        Screen.MousePointer = 11
        On Error GoTo errH
        gcnOracle.BeginTrans: blnTrans = True
        For i = 0 To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
        Next
        gcnOracle.CommitTrans: blnTrans = False
        Call SendMsgDiag(datCur)
        On Error GoTo 0
        Screen.MousePointer = 0
    End If
    Exit Sub
errH:
    Screen.MousePointer = 0
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function CheckData(Optional ByVal blnDiagnose As Boolean) As Boolean
    Dim i As Long
    Dim j As Long
    Dim curDate As Date
    Dim blnHaveSel As Boolean
    Dim lngSize As Long
    gclsPros.InfosChange = False
    If Not CheckDiagData(zlDatabase.Currentdate, blnHaveSel) Then Exit Function
    '�������
    Call CacheLoadVsDiagData(vsDiagXY, , , True)
    Call CacheLoadVsDiagData(vsDiagZY, , , True)
    If gclsPros.PatiType = PF_סԺ Then
        Call CacheLoadDiagMatchData(, True)
        Call CacheCtrlValues
    End If
    gclsPros.DiagSel = blnHaveSel
    '���ѡ���˹�������������Ϸ���ҳ��ϣ������������±���
    Call UpdateCacheRecInfo(2)
    gclsPros.MainInfoRec.Filter = "�Ƿ�ı�=1"
    gclsPros.InfosChange = Not gclsPros.MainInfoRec.EOF
    
    CheckData = True
End Function

Private Sub InitDaigSel(Optional ByVal blnAfterLoad As Boolean)
'��ʼ�����ѡ�����
'������blnAfterLoad=�Ƿ����ݼ���֮���ʼ��
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim lngColWidth As Long, LngRow As Long
    
    tabFunc.Visible = gclsPros.Have��ҽ
    If tabFunc.Visible Then
        picXY.Visible = True
        picXY.ZOrder
        picZY.Visible = False
    End If
    frmXY.Visible = (gclsPros.PatiType = PF_סԺ)
    frmZY.Visible = (gclsPros.PatiType = PF_סԺ)
    Call InitTableDiag
End Sub

Private Function IsSignature() As Boolean
'���ܣ���ȡ��ǰ���˵�ҽʦ��ǩ�����
'���أ������Ƿ���ǩ��
    Dim rsTmp As ADODB.Recordset
    Dim intCurr As Integer, intHave As Integer
    Dim strSql As String, blnReadOnly As Boolean
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
    
    strSql = "select 'סԺҽʦ' as ��Ϣ��, A.סԺҽʦ as ��Ϣֵ from ������ҳ A where a.����id = [1]  And a.��ҳid = [2]" & vbNewLine & _
             "union all" & vbNewLine & _
             "select A.��Ϣ�� , A.��Ϣֵ from ������ҳ�ӱ� A where  A.����id = [1] And A.��ҳid = [2]"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ѯ����ҽʦ���", gclsPros.����ID, gclsPros.��ҳID)
    
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

