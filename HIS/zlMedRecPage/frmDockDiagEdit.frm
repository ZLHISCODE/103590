VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDockDiagEdit 
   BorderStyle     =   0  'None
   Caption         =   "申请单诊断"
   ClientHeight    =   4125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   9990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
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
         Caption         =   "根据疾病编码输入(&2)"
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
         Caption         =   "根据诊断标准输入(&1)"
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
            Caption         =   "…"
            Enabled         =   0   'False
            Height          =   240
            Index           =   20
            Left            =   9180
            TabIndex        =   2
            TabStop         =   0   'False
            ToolTipText     =   "选择(*)"
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
               Name            =   "宋体"
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
            Caption         =   "死亡患者尸检(&R)"
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
            Caption         =   "门诊与出院(&H)"
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
            Caption         =   "入院与出院(&I)"
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
            Caption         =   "放射与病理(&L)"
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
            Caption         =   "临床与病理(&M)"
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
            Caption         =   "分化程度(&F)"
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
            Caption         =   "最高诊断依据(&G)"
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
            Caption         =   "门诊与入院(&J)"
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
            Caption         =   "死亡时间(&N)"
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
            Caption         =   "死亡原因(&P)"
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
            Caption         =   "临床与尸检(&T)"
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
            Name            =   "宋体"
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
            Caption         =   "门诊与出院(&A)"
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
            Caption         =   "入院与出院(&B)"
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
         Caption         =   "根据疾病编码输入(&4)"
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
         Caption         =   "根据诊断标准输入(&3)"
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
            Name            =   "宋体"
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
      Tag             =   "西医诊断"
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
            Caption         =   "西医诊断"
            Key             =   "西医诊断"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "中医诊断"
            Key             =   "中医诊断"
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

Private mstrOldOutWay As String          '存储出院方式的变量

Public Function ShowMe() As Boolean
'返回： ShowDiagEdit= 是确定还是取消
    Show 1, gclsPros.MainForm
    ShowMe = gclsPros.IsOK
End Function

Private Sub cmdCancel_Click()
    Call CmdCancelClick
End Sub

Private Sub Form_Load()
    '住院首页相关
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim lngColWidth As Long
    On Error GoTo errH
    gclsPros.IsOK = False
    If gclsPros.PatiType = PF_住院 Then
        Call InitControlData
        gclsPros.IsSigned = IsSignature
    Else
        strSQL = "Select 1 from 病人挂号记录 where Rownum<2 And ID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "查询门诊病人是否在历史库", gclsPros.主页ID)
        If Not rsTmp.EOF Then
            gclsPros.Moved = False
        Else
            strSQL = "Select 1 from H病人挂号记录 where Rownum<2 And ID = [1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "查询门诊病人是否在历史库", gclsPros.主页ID)
            gclsPros.Moved = Not rsTmp.EOF
        End If
    End If
    strSQL = "Select A.险类,Nvl(A.路径状态,-1) 路径状态,A.性别" & _
        " From 病案主页 A" & _
        " Where A.病人ID=[1] And A.主页ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, gclsPros.病人ID, gclsPros.主页ID)
    If Not rsTmp.EOF Then
        gclsPros.InsureType = NVL(rsTmp!险类, 0)
        '-1:未导入,0-不符合导入条件，1-执行中，2-正常结束，3-变异结束
        gclsPros.PathState = Val(rsTmp!路径状态 & "")
        gclsPros.Sex = rsTmp!性别 & ""
    End If
    
    If gclsPros.PatiType = PF_住院 Then
        strSQL = "Select 1 From 病人手麻记录  A Where  A.病人ID=[1] And A.主页ID=[2] "
        If gclsPros.Moved Then
            strSQL = Replace(strSQL, "病人手麻记录", "H病人手麻记录")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, gclsPros.病人ID, gclsPros.主页ID)
        gclsPros.Have手术 = Not rsTmp.EOF
        '获取临川路径信息
        Call GetPatiPathInfo
    End If
    If gclsPros.DiagRowIDs = "" Then
        With gclsPros.DiagConn
            .Filter = "标识ID=" & gclsPros.AplyMark
            .Sort = "诊断ID"
            Do While Not .EOF
                gclsPros.DiagRowIDs = gclsPros.DiagRowIDs & IIf(gclsPros.DiagRowIDs = "", "", ",") & !诊断id
                .MoveNext
            Loop
        End With
    End If
    '设置界面
    Call InitDaigSel
    '加载数据
    Call LoadData
    Call gclsPros.InitFacePara '设置界面参数控件状态
    If gclsPros.BlnICDEleven Then
        gclsPros.CurrentForm.vsDiagXY.ColHidden(DI_备注) = True
        gclsPros.CurrentForm.vsDiagXY.ColHidden(DI_前注释) = True
        gclsPros.CurrentForm.vsDiagXY.ColHidden(DI_后注释) = True
        gclsPros.CurrentForm.vsDiagXY.ColHidden(DI_ICD附码) = True
        gclsPros.CurrentForm.vsDiagXY.ColWidth(DI_诊断编码) = gclsPros.CurrentForm.vsDiagXY.ColWidth(DI_诊断编码) + gclsPros.CurrentForm.vsDiagXY.ColWidth(DI_ICD附码) - IIf(gclsPros.PatiType = PF_门诊, 800, 300)
        gclsPros.CurrentForm.vsDiagXY.ColWidth(DI_诊断描述) = gclsPros.CurrentForm.vsDiagXY.ColWidth(DI_诊断描述) + gclsPros.CurrentForm.vsDiagXY.ColWidth(DI_备注) - IIf(gclsPros.PatiType = PF_门诊, 800, 300)
        
        gclsPros.CurrentForm.vsDiagZY.ColHidden(DI_备注) = True
        gclsPros.CurrentForm.vsDiagZY.ColHidden(DI_前注释) = True
        gclsPros.CurrentForm.vsDiagZY.ColHidden(DI_后注释) = True
        gclsPros.CurrentForm.vsDiagZY.ColHidden(DI_中医证候) = True
        gclsPros.CurrentForm.vsDiagZY.ColWidth(DI_诊断编码) = gclsPros.CurrentForm.vsDiagZY.ColWidth(DI_诊断编码) + gclsPros.CurrentForm.vsDiagZY.ColWidth(DI_备注) - IIf(gclsPros.PatiType = PF_门诊, 350, 300)
        gclsPros.CurrentForm.vsDiagZY.ColWidth(DI_诊断描述) = gclsPros.CurrentForm.vsDiagZY.ColWidth(DI_诊断描述) + gclsPros.CurrentForm.vsDiagZY.ColWidth(DI_中医证候) - IIf(gclsPros.PatiType = PF_门诊, 350, 300)
    Else
        gclsPros.CurrentForm.vsDiagXY.ColHidden(DI_备注) = False
        gclsPros.CurrentForm.vsDiagXY.ColHidden(DI_前注释) = IIf(gclsPros.PatiType = PF_门诊, True, Not gclsPros.AddAnnotation)
        gclsPros.CurrentForm.vsDiagXY.ColHidden(DI_后注释) = IIf(gclsPros.PatiType = PF_门诊, True, Not gclsPros.AddAnnotation)
        gclsPros.CurrentForm.vsDiagXY.ColHidden(DI_ICD附码) = False
        gclsPros.CurrentForm.vsDiagXY.ColWidth(DI_诊断编码) = IIf(gclsPros.PatiType = PF_门诊, 900, 850)
        gclsPros.CurrentForm.vsDiagXY.ColWidth(DI_诊断描述) = IIf(gclsPros.PatiType = PF_门诊, 4000, 2500)
        
        gclsPros.CurrentForm.vsDiagZY.ColHidden(DI_备注) = False
        gclsPros.CurrentForm.vsDiagZY.ColHidden(DI_前注释) = True
        gclsPros.CurrentForm.vsDiagZY.ColHidden(DI_后注释) = True
        gclsPros.CurrentForm.vsDiagZY.ColHidden(DI_中医证候) = False
        gclsPros.CurrentForm.vsDiagZY.ColWidth(DI_诊断编码) = IIf(gclsPros.PatiType = PF_门诊, 900, 850)
        gclsPros.CurrentForm.vsDiagXY.ColWidth(DI_诊断描述) = IIf(gclsPros.PatiType = PF_门诊, 2900, 1900)
    End If
    If gclsPros.PatiType = PF_住院 And gclsPros.FuncType <> f申请单诊断 Then
        '获取原有的出院方式
        mstrOldOutWay = vsDiagXY.TextMatrix(DT_出院诊断XY, DI_出院情况)
        If gclsPros.Have中医 And mstrOldOutWay = "" Then
            mstrOldOutWay = vsDiagZY.TextMatrix(DT_出院诊断XY, DI_出院情况)
        End If

        Call ChangeOutInfo(zlStr.NeedName(mstrOldOutWay))
        '加载诊断符合情况数据并缓存
        Call CacheLoadDiagMatchData(GetDiagMatchData(gclsPros.病人ID, gclsPros.主页ID))
        '根据签名后设置界面状态
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
    vsDiagXY.Height = picXY.ScaleHeight - vsDiagXY.Top - 120 - IIf(gclsPros.PatiType <> PF_门诊, frmXY.Height, 0)
    vsDiagXY.Width = picXY.ScaleWidth - vsDiagXY.Left * 2
    optDiag(1).Left = picXY.ScaleWidth - optDiag(1).Width - 120
    optDiag(0).Left = optDiag(1).Left - optDiag(0).Width - 120
    lngWidth = vsDiagXY.Width
    For i = 0 To vsDiagXY.Cols
        If Not vsDiagXY.ColHidden(i) And i <> DI_诊断描述 Then
            lngColsWidth = lngColsWidth + vsDiagXY.ColWidth(i)
        End If
    Next
    If lngWidth > lngColsWidth Then
        lngColsWidth = lngWidth - lngColsWidth
        vsDiagXY.ColWidth(DI_诊断描述) = lngColsWidth - 400
    End If
    If gclsPros.PatiType = PF_住院 Then frmXY.Top = vsDiagXY.Top + vsDiagXY.Height + 120
End Sub

Private Sub picZY_Resize()
    Dim lngWidth As Long
    Dim lngColsWidth As Long
    Dim i As Long
    On Error Resume Next
    vsDiagZY.Height = picZY.ScaleHeight - vsDiagZY.Top - 120 - IIf(gclsPros.PatiType <> PF_门诊, frmXY.Height, 0)
    vsDiagZY.Width = picZY.ScaleWidth - vsDiagZY.Left * 2
    optDiag(3).Left = picZY.ScaleWidth - optDiag(3).Width - 120
    optDiag(2).Left = optDiag(3).Left - optDiag(2).Width - 120
    lngWidth = vsDiagZY.Width
    For i = 0 To vsDiagZY.Cols
        If Not vsDiagZY.ColHidden(i) And i <> DI_诊断描述 Then
            lngColsWidth = lngColsWidth + vsDiagZY.ColWidth(i)
        End If
    Next
    If lngWidth > lngColsWidth Then
        lngColsWidth = lngWidth - lngColsWidth
        vsDiagZY.ColWidth(DI_诊断描述) = lngColsWidth - 400
    End If
    If gclsPros.PatiType = PF_住院 Then frmXY.Top = vsDiagXY.Top + vsDiagXY.Height + 120
End Sub

Private Sub Form_Unload(Cancel As Integer)
    gclsPros.IsOK = True
    Call FormUnLoad(Cancel)
End Sub

Private Sub tabFunc_Click()
    If tabFunc.Visible Then
        picXY.Visible = tabFunc.SelectedItem.Key = "西医诊断"
        picZY.Visible = tabFunc.SelectedItem.Key = "中医诊断"
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
    '读取诊断
    Set rsTmp = GetPatiDiagData(gclsPros.病人ID, gclsPros.主页ID, IIf(gclsPros.PatiType <> PF_门诊, 1, 0), , , gclsPros.Moved)
    rsTmp.Filter = "记录来源=" & IIf(gclsPros.FuncType = f病案首页, 4, 3)
    strDiagFilter = rsTmp.Filter
    rsTmp.Filter = "编码类别='D'"
    If rsTmp.EOF Then
        rsTmp.Filter = "编码类别='E'"
        If Not rsTmp.EOF Then
            gclsPros.BlnICDEleven = True
        Else
            If gclsPros.PatiType = PF_住院 Then
                If gclsPros.InICDEleven Then
                    gclsPros.BlnICDEleven = True
                Else
                    gclsPros.BlnICDEleven = False
                End If
            ElseIf gclsPros.PatiType = PF_门诊 Then
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
    '2、加载西医
    '   1-西医门诊诊断;2-西医入院诊断;3-出院诊断(其他诊断);5-院内感染;6-病理诊断;7-损伤中毒码;10-并发症
    Call CacheLoadVsDiagData(vsDiagXY, rsTmp, IIf(gclsPros.PatiType <> PF_门诊, "2", "1"), , -1)
    '3、加载中医诊断
    '   11-中医门诊诊断;12-中医入院诊断;13-中医出院诊断(主要诊断、其它诊断)
    If gclsPros.Have中医 Then
        Call CacheLoadVsDiagData(vsDiagZY, rsTmp, IIf(gclsPros.PatiType <> PF_门诊, "12", "11"), , -1)
    End If
    '加载确认传染病诊断
    If gclsPros.IsComfirmInfect Then
        vsDiagXY.ColHidden(DI_关联) = True
        vsDiagXY.ColWidth(DI_诊断描述) = vsDiagXY.ColWidth(DI_诊断描述) + vsDiagXY.ColWidth(DI_关联)
        Call LoadInfeciousDiseases
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadInfeciousDiseases()
'功能：生成确认传染病诊断
    Dim lngStart As Long, dtType As DiagType
    Dim blnAdd As Boolean, LngRow As Long, j As Long
    Dim strSQL As String, str性别 As String
    Dim rsInput As ADODB.Recordset
    
    On Error GoTo errH
    If gclsPros.Sex Like "*男*" Then
        str性别 = "男"
    ElseIf gclsPros.Sex Like "*女*" Then
        str性别 = "女"
    End If
    With gclsPros.DiagConn
        .Filter = "类型=1"
        dtType = IIf(gclsPros.PatiType = PF_门诊, DT_门诊诊断XY, DT_出院诊断XY)
        lngStart = vsDiagXY.FindRow(dtType, , DI_诊断分类, , True)
        Do While Not .EOF
            blnAdd = True: LngRow = lngStart
            '存在疾病ID与诊断ID才进行处理
            If Val(!诊断目录ID & "") <> 0 Or Val(!疾病目录ID & "") <> 0 Then
                For j = LngRow To vsDiagXY.Rows - 1
                    If Val(vsDiagXY.TextMatrix(j, DI_诊断分类)) = dtType Then
                        LngRow = j
                        If vsDiagXY.TextMatrix(j, DI_诊断描述) = "" Then Exit For
                        If Val(vsDiagXY.TextMatrix(j, DI_疾病ID)) = Val(!疾病目录ID & "") And Val(!疾病目录ID & "") <> 0 Then
                            blnAdd = False: Exit For
                        ElseIf Val(vsDiagXY.TextMatrix(j, DI_诊断ID)) = Val(!诊断目录ID & "") And Val(!诊断目录ID & "") <> 0 Then
                            blnAdd = False: Exit For
                        End If
                    Else
                        Exit For
                    End If
                Next
                If blnAdd Then
                    If Val(!诊断目录ID & "") <> 0 And (gclsPros.DiagInputXY = 0 Or Val(!疾病目录ID & "") = 0) Then
                        strSQL = "Select Distinct a.Id, a.项目id, a.编码, b.序号, b.附码, Null 附码id, Null 附码名称, a.名称, a.说明, Null 编者, a.简码, a.疗效限制, a.分娩, a.是否病人," & vbNewLine & _
                                    "                b.编码 疾病编码, b.Id 疾病id, b.类别 疾病类别, a.诊断id" & vbNewLine & _
                                    "From (Select a.Id, a.Id 项目id, a.编码, Null 序号, Null 附码, Null 附码id, Null 附码名称, a.名称, a.说明, a.编者, b.简码, 0 疗效限制, 0 分娩, 0 是否病人," & vbNewLine & _
                                    "              Max(d.疾病id) 疾病id, a.Id 诊断id" & vbNewLine & _
                                    "       From 疾病诊断目录 a, 疾病诊断别名 b, 疾病诊断对照 d" & vbNewLine & _
                                    "       Where a.Id = [1] And a.Id = b.诊断id And a.Id = d.诊断id(+) And a.类别 = 1 And b.码类 = [4]" & vbNewLine & _
                                    " And (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd'))" & vbNewLine & _
                                    "       Group By a.Id, a.编码, a.名称, a.说明, a.编者, b.简码) a, 疾病编码目录 b, 疾病诊断科室 c, 疾病诊断科室 d" & vbNewLine & _
                                    "Where a.疾病id = b.Id(+) And c.诊断id(+) = a.Id And d.诊断id(+) = a.Id And c.科室id(+) = [5] And d.人员id(+) = [6]" & vbNewLine & _
                                    "Order By a.编码"
                    Else
                        strSQL = "Select Distinct a.Id, a.项目id, a.编码, a.序号, a.附码, a.附码id, a.附码名称, a.名称, a.说明, a.编者, a.分类id, a.简码, a.疗效限制, a.分娩, a.是否病人," & vbNewLine & _
                                    "                a.疾病编码, a.疾病id, a.疾病类别, a.诊断id" & vbNewLine & _
                                    "From (Select a.Id, a.Id 项目id, a.编码, a.序号, a.附码, Null 附码id, Null 附码名称, a.名称, a.说明, Null 编者, a.分类id, a.五笔码 简码, a.疗效限制, a.分娩," & vbNewLine & _
                                    "              c.是否病人, a.编码 疾病编码, a.Id 疾病id, a.类别 疾病类别, Max(b.诊断id) 诊断id" & vbNewLine & _
                                    "       From 疾病编码目录 a, 疾病诊断对照 b, 疾病编码分类 c" & vbNewLine & _
                                    "       Where a.Id = [2] And a.Id = b.疾病id(+) And a.分类id = c.Id(+) And a.类别='D' And" & vbNewLine & _
                                    IIf(str性别 <> "", "  (A.性别限制=[3] Or A.性别限制 is NULL) And ", " ") & " (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
                                    "       Group By a.Id, a.编码, a.序号, a.附码, a.名称, a.说明, a.分类id, a.五笔码, a.疗效限制, a.分娩, a.类别, c.是否病人) a, 疾病编码科室 c, 疾病编码科室 d" & vbNewLine & _
                                    "Where c.疾病id(+) = a.Id And d.疾病id(+) = a.Id And c.科室id(+) = [5] And d.人员id(+) = [6]" & vbNewLine & _
                                    "Order By a.编码"
                    End If
                    Set rsInput = zlDatabase.OpenSQLRecord(strSQL, "确认传染病", Val(!诊断目录ID & ""), Val(!疾病目录ID & ""), str性别, gclsPros.BriefCode + 1, gclsPros.出院科室ID, UserInfo.ID)
                    If rsInput.RecordCount > 0 Then
                        '新增行
                         If vsDiagXY.TextMatrix(LngRow, DI_诊断描述) <> "" Then
                             LngRow = LngRow + 1: vsDiagXY.AddItem "", LngRow
                             vsDiagXY.TextMatrix(LngRow, DI_诊断分类) = dtType
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
'初始化诊断选择界面
'参数：blnAfterLoad=是否数据加载之后初始化
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim lngColWidth As Long, LngRow As Long
    
    tabFunc.Visible = gclsPros.Have中医
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
'功能：获取当前病人的医师及签名情况
'返回：界面是否已签名
    Dim rsTmp As ADODB.Recordset
    Dim intCurr As Integer, intHave As Integer
    Dim strSQL As String, blnReadOnly As Boolean
    Dim i As Integer
    '说明：arrInfos 数组的元素一一对应，人员级别从低到高
    Dim arrInfos() As Variant '各类签名的信息名
    Dim arrSgnIdxs() As Variant '各类签名的信息名
    Dim arrName() As Variant
    On Error GoTo errH
    blnReadOnly = False: intCurr = -1: intHave = -1
    arrSgnIdxs = Array("住院医师签名", "主治医师签名", "主任医师签名", "科主任签名")
    arrInfos = Array("住院医师", "主治医师", "主任医师", "科主任")
    arrName = Array("", "", "", "")
    
    strSQL = "select '住院医师' as 信息名, A.住院医师 as 信息值 from 病案主页 A where a.病人id = [1]  And a.主页id = [2]" & vbNewLine & _
             "union all" & vbNewLine & _
             "select A.信息名 , A.信息值 from 病案主页从表 A where  A.病人id = [1] And A.主页id = [2]"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "查询病人医师情况", gclsPros.病人ID, gclsPros.主页ID)
    
    For i = LBound(arrInfos) To UBound(arrInfos)
        rsTmp.Filter = "信息名='" & arrInfos(i) & "'"
        If Not rsTmp.EOF Then
            arrName(i) = rsTmp!信息值 & ""
        End If
    Next
    For i = LBound(arrName) To UBound(arrName)
        If arrName(i) = UserInfo.姓名 Then
            intCurr = i
        End If
        gclsPros.AuxiInfo.Filter = "信息名='" & arrSgnIdxs(i) & "'"
        If Not gclsPros.AuxiInfo.EOF Then
            intHave = i
        End If
    Next

    '如果当前人员签名级别不高于已签名级别，则不可编辑
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
'功能：初始化界面控件数据
    Dim i As Integer
    On Error GoTo errH
    If gclsPros.PatiType = PF_住院 Then
        Call SetCboFromList(Array("否", "是"), Array(cboBaseInfo(BCC_死亡患者尸检)), 0)
        Call SetCboFromRec(Array(BCC_分化程度, BCC_最高诊断依据), 0)
        Call SetCboFromList(Array("0-未做", "1-符合", "2-不符合", "3-不肯定"), Array(cboBaseInfo(BCC_门诊与出院XY), cboBaseInfo(BCC_门诊与入院), cboBaseInfo(BCC_入院与出院XY), cboBaseInfo(BCC_放射与病理), cboBaseInfo(BCC_临床与病理), _
         cboBaseInfo(BCC_门诊与出院ZY), cboBaseInfo(BCC_入院与出院ZY), cboBaseInfo(BCC_临床与尸检)))
    End If
       
    Set gclsPros.PatiInfo = GetPatiMainInfoData(gclsPros.病人ID, gclsPros.主页ID)
    '加载病人信息
    If Not gclsPros.PatiInfo.EOF Then
        For i = 0 To gclsPros.PatiInfo.Fields.Count - 1
             Call SetCtrlValues(UCase(gclsPros.PatiInfo.Fields(i).Name & ""), gclsPros.PatiInfo.Fields(i).Value & "", , True)
        Next
    End If

    Set gclsPros.AuxiInfo = GetPatiAuxiInfoData(gclsPros.病人ID, gclsPros.主页ID)   '从表信息
    If Not gclsPros.AuxiInfo.EOF Then
        gclsPros.AuxiInfo.MoveFirst
        For i = 1 To gclsPros.AuxiInfo.RecordCount
            Call SetCtrlValues(gclsPros.AuxiInfo!信息名 & "", gclsPros.AuxiInfo!信息值 & "", gclsPros.AuxiInfo!编码 & "")
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
'功能：根据当前病人的医师及签名情况，确定签名及界面数据的可编辑性
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
                    vsTmp.Cell(flexcpBackColor, 0, DI_诊断类型, vsTmp.Rows - 1, vsTmp.Cols - 1) = &H8000000F
                    vsTmp.Cell(flexcpBackColor, 0, DI_关联, vsTmp.Rows - 1, DI_关联) = &H80000005
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
'功能：将出院方式情况的SQL加入数组
    Dim strNewOutWay As String
    Dim strValue As String
    On Error GoTo errH
    strNewOutWay = vsDiagXY.TextMatrix(DT_出院诊断XY, DI_出院情况)
    If gclsPros.Have中医 And strNewOutWay = "" Then
       strNewOutWay = vsDiagZY.TextMatrix(DT_出院诊断XY, DI_出院情况)
    End If
    
    If (mstrOldOutWay <> "死亡" And strNewOutWay = "死亡") Or (mstrOldOutWay = "死亡" And strNewOutWay <> "死亡") Then
        strNewOutWay = IIf(strNewOutWay = "死亡", "死亡", "正常")
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_病案主页_首页整理EX(" & gclsPros.病人ID & "," & gclsPros.主页ID & ",'出院方式','" & strNewOutWay & "')"
        If strNewOutWay = "死亡" Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_病案主页从表_首页整理(" & gclsPros.病人ID & "," & gclsPros.主页ID & ",'出院转入', NULL)"
        End If
    End If
    
    gclsPros.MainInfoRec.Filter = "是否改变=1 and 信息名='尸检标志'"
    If Not gclsPros.MainInfoRec.EOF Then
        strValue = gclsPros.MainInfoRec!信息现值 & ""
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_病案主页_首页整理EX(" & gclsPros.病人ID & "," & gclsPros.主页ID & ",'尸检标志','" & strValue & "')"
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






