VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmAdviceReport 
   AutoRedraw      =   -1  'True
   Caption         =   "打印执行单"
   ClientHeight    =   8625
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   16590
   Icon            =   "frmAdviceReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8625
   ScaleWidth      =   16590
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fraDetail 
      Height          =   8160
      Left            =   120
      TabIndex        =   11
      Top             =   -15
      Width           =   16425
      Begin VB.PictureBox PicView 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   7875
         Left            =   7200
         ScaleHeight     =   7875
         ScaleWidth      =   9135
         TabIndex        =   27
         Top             =   180
         Width           =   9135
         Begin VB.CommandButton cmdRePrint 
            Caption         =   "重打停止到上次打印前的医嘱"
            Height          =   345
            Left            =   6645
            TabIndex        =   34
            Top             =   18
            Width           =   2495
         End
         Begin VB.CommandButton cmdSelect 
            Caption         =   "查询(&Q)"
            Height          =   350
            Left            =   5595
            TabIndex        =   33
            Top             =   15
            Width           =   975
         End
         Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
            Height          =   7380
            Left            =   0
            TabIndex        =   28
            Top             =   420
            Width           =   9135
            _cx             =   16113
            _cy             =   13017
            Appearance      =   1
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
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16771802
            ForeColorSel    =   -2147483640
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   2
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   250
            RowHeightMax    =   2000
            ColWidthMin     =   0
            ColWidthMax     =   5000
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmAdviceReport.frx":014A
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
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin MSComCtl2.DTPicker dtpViewBegin 
            Height          =   300
            Left            =   1200
            TabIndex        =   29
            Top             =   45
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   529
            _Version        =   393216
            CalendarTitleBackColor=   8388608
            CalendarTitleForeColor=   16777215
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   123994115
            CurrentDate     =   37953
         End
         Begin MSComCtl2.DTPicker dtpViewEnd 
            Height          =   300
            Left            =   3480
            TabIndex        =   30
            Top             =   45
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   529
            _Version        =   393216
            CalendarTitleBackColor=   8388608
            CalendarTitleForeColor=   16777215
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   123994115
            CurrentDate     =   37953
         End
         Begin VB.Label lblTo 
            Caption         =   "~"
            Height          =   135
            Left            =   3345
            TabIndex        =   32
            Top             =   165
            Width           =   135
         End
         Begin VB.Label lblView 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "已打执行时间"
            Height          =   180
            Left            =   80
            TabIndex        =   31
            Top             =   105
            Width           =   1080
         End
      End
      Begin VB.Frame fraPati 
         BorderStyle     =   0  'None
         Height          =   3960
         Left            =   120
         TabIndex        =   12
         Top             =   4080
         Width           =   7005
         Begin VB.CheckBox chkOut 
            Caption         =   "包含出院病人"
            Height          =   375
            Left            =   0
            TabIndex        =   39
            Top             =   705
            Width           =   915
         End
         Begin VB.Frame fraBaby 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   3600
            TabIndex        =   23
            Top             =   60
            Visible         =   0   'False
            Width           =   3195
            Begin VB.OptionButton optBaby 
               Caption         =   "病人医嘱"
               Height          =   180
               Index           =   1
               Left            =   1080
               TabIndex        =   26
               Top             =   0
               Width           =   1020
            End
            Begin VB.OptionButton optBaby 
               Caption         =   "所有医嘱"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   25
               Top             =   0
               Value           =   -1  'True
               Width           =   1020
            End
            Begin VB.OptionButton optBaby 
               Caption         =   "婴儿医嘱"
               Height          =   180
               Index           =   2
               Left            =   2175
               TabIndex        =   24
               Top             =   0
               Width           =   1020
            End
         End
         Begin VB.ComboBox cboUnit 
            Height          =   300
            Left            =   1035
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   0
            Width           =   2535
         End
         Begin VB.CommandButton cmdNoPati 
            Caption         =   "全清"
            Height          =   330
            Left            =   90
            TabIndex        =   6
            TabStop         =   0   'False
            ToolTipText     =   "Ctrl + R"
            Top             =   3525
            Width           =   870
         End
         Begin VB.CommandButton cmdAllPati 
            Caption         =   "全选"
            Height          =   330
            Left            =   90
            TabIndex        =   5
            TabStop         =   0   'False
            ToolTipText     =   "Ctrl + A"
            Top             =   3150
            Width           =   870
         End
         Begin MSComctlLib.ImageList img16 
            Left            =   240
            Top             =   1680
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   1
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAdviceReport.frx":0273
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ImageList img32 
            Left            =   255
            Top             =   1215
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   32
            ImageHeight     =   32
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   2
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAdviceReport.frx":03CD
                  Key             =   "Left"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAdviceReport.frx":0CA7
                  Key             =   "Right"
               EndProperty
            EndProperty
         End
         Begin VSFlex8Ctl.VSFlexGrid vsPati 
            Bindings        =   "frmAdviceReport.frx":1581
            Height          =   3515
            Left            =   1035
            TabIndex        =   37
            Top             =   360
            Width           =   5775
            _cx             =   10186
            _cy             =   6200
            Appearance      =   2
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
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16444122
            ForeColorSel    =   -2147483640
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   0
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   12
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   250
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmAdviceReport.frx":1595
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
         Begin MSComctlLib.ImageList imgPati 
            Left            =   240
            Top             =   2280
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   1
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmAdviceReport.frx":16E2
                  Key             =   "Child"
               EndProperty
            EndProperty
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "住院病人(&I)"
            Height          =   180
            Left            =   0
            TabIndex        =   4
            Top             =   435
            Width           =   990
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "住院病区(&U)"
            Height          =   180
            Left            =   0
            TabIndex        =   2
            Top             =   60
            Width           =   990
         End
      End
      Begin VB.Frame fraline 
         Height          =   30
         Left            =   120
         MousePointer    =   7  'Size N S
         TabIndex        =   22
         Top             =   3960
         Width           =   6855
      End
      Begin VB.Frame fraReport 
         BorderStyle     =   0  'None
         Height          =   3615
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   6975
         Begin VB.Frame fraCondition 
            BorderStyle     =   0  'None
            Height          =   3600
            Left            =   4680
            TabIndex        =   15
            Top             =   15
            Width           =   2325
            Begin VB.CheckBox ChkWaitPrint 
               Caption         =   "只显示待打印的病人"
               Height          =   195
               Left            =   0
               TabIndex        =   36
               Top             =   3360
               Width           =   1980
            End
            Begin VB.CommandButton cmdView 
               Caption         =   "隐藏已打印信息"
               Height          =   350
               Left            =   0
               Style           =   1  'Graphical
               TabIndex        =   35
               Top             =   0
               Width           =   2055
            End
            Begin VB.CheckBox chk期效 
               Caption         =   "长期(&L)"
               Height          =   195
               Index           =   0
               Left            =   0
               TabIndex        =   18
               Top             =   1665
               Value           =   1  'Checked
               Width           =   930
            End
            Begin VB.CheckBox chk期效 
               Caption         =   "临时(&T)"
               Height          =   195
               Index           =   1
               Left            =   1080
               TabIndex        =   17
               Top             =   1665
               Value           =   1  'Checked
               Width           =   930
            End
            Begin VB.CheckBox chk重复打印 
               Caption         =   "包含已打印过的(&A)"
               Height          =   195
               Left            =   0
               TabIndex        =   16
               Top             =   1965
               Width           =   2295
            End
            Begin MSComCtl2.DTPicker dtpBegin 
               Height          =   300
               Left            =   0
               TabIndex        =   19
               Top             =   825
               Width           =   2070
               _ExtentX        =   3651
               _ExtentY        =   529
               _Version        =   393216
               CalendarTitleBackColor=   8388608
               CalendarTitleForeColor=   16777215
               CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
               Format          =   123994115
               CurrentDate     =   37953
            End
            Begin MSComCtl2.DTPicker dtpEnd 
               Height          =   300
               Left            =   0
               TabIndex        =   20
               Top             =   1185
               Width           =   2070
               _ExtentX        =   3651
               _ExtentY        =   529
               _Version        =   393216
               CalendarTitleBackColor=   8388608
               CalendarTitleForeColor=   16777215
               CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
               Format          =   123994115
               CurrentDate     =   37953
            End
            Begin VB.Label lbl执行时间 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "待打执行时间(&E)"
               Height          =   180
               Left            =   0
               TabIndex        =   21
               Top             =   480
               Width           =   1350
            End
         End
         Begin VB.CommandButton cmdSetup 
            Caption         =   "打印设置"
            Height          =   330
            Left            =   0
            TabIndex        =   10
            TabStop         =   0   'False
            ToolTipText     =   "Ctrl + S"
            Top             =   405
            Width           =   990
         End
         Begin MSComctlLib.ListView lvwReport 
            Height          =   3600
            Left            =   1035
            TabIndex        =   1
            Top             =   0
            Width           =   3540
            _ExtentX        =   6244
            _ExtentY        =   6350
            View            =   3
            Arrange         =   2
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            AllowReorder    =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            SmallIcons      =   "img16"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "报表"
               Object.Width           =   6615
            EndProperty
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "执行单(&R)"
            Height          =   180
            Left            =   180
            TabIndex        =   0
            Top             =   60
            Width           =   810
         End
      End
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   16590
      TabIndex        =   13
      Top             =   8190
      Width           =   16590
      Begin VB.CommandButton cmdGoOn 
         Caption         =   "续打(&A)"
         Height          =   350
         Left            =   1560
         TabIndex        =   38
         Top             =   0
         Width           =   1100
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "打印(&P)"
         Height          =   350
         Left            =   3870
         TabIndex        =   8
         Top             =   0
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "退出(&X)"
         Height          =   350
         Left            =   5235
         TabIndex        =   9
         Top             =   0
         Width           =   1100
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "预览(&V)"
         Height          =   350
         Left            =   2760
         TabIndex        =   7
         Top             =   0
         Width           =   1100
      End
   End
End
Attribute VB_Name = "frmAdviceReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mobjReport As clsReport
Attribute mobjReport.VB_VarHelpID = -1

Private mMainPrivs As String 'IN:调用主界面所具有的权限,注意非内部模块权限
Private mlng病区ID As Long 'IN
Private mlng病人ID As Long 'IN
Private mblnOnePati As Boolean 'IN，单病人模式
Private mstr当前病人 As String '逐个打印时当前打印的:病人ID,主页ID
Private mstrPrintedID As String '打印过的医嘱ID串
Private mlng医嘱处理范围 As Long    '0-所有医嘱，1-病人医嘱，2-婴儿医嘱(报表中为-1，0，婴儿序号)
Private mlngLastRow As Long     '打印的最后一行
Private mintType As Long    '打印类型，0-打印，1-重打，2-续打
Private mbln逐个病人换页打印 As Boolean
Private mbln产科 As Boolean
Private mbln婴儿过滤 As Boolean

Private Enum AdviceCol
    col姓名 = 0
    COL床位 = 1
    COL婴儿 = 2
    col期效 = 3
    col医嘱内容 = 4
    COL频率 = 5
    COL上次打印时间 = 6
    COL停止时间 = 7
    col类别 = 8
    col相关ID = 9
End Enum

Private Enum PatiCol
    COL_选择 = 0
    COL_姓名 = 1
    COL_住院号 = 2
    COL_床位 = 3
    COL_住院医师 = 4
    COL_费别 = 5
    COL_护理等级 = 6
    COL_科室 = 7
    COL_入院时间 = 8
    COL_病人类型 = 9
    COL_主页ID = 10
    COL_婴儿 = 11
End Enum

Public Function ShowMe(frmParent As Object, ByVal MainPrivs As String, _
    ByVal lng病区ID As Long, ByVal lng病人ID As Long, _
     ByVal blnOnePati As Boolean, Optional ByVal lng医护科室ID As Long, Optional ByVal lng婴儿病区ID As Long) As Boolean
'参数：
    mMainPrivs = MainPrivs
    
    mlng病人ID = lng病人ID
    mlng病区ID = lng病区ID
    If lng婴儿病区ID <> 0 Then
        If lng婴儿病区ID = lng医护科室ID Then
            mlng病区ID = lng婴儿病区ID
        End If
    End If
   
    mblnOnePati = blnOnePati
        
    Me.Show 1, frmParent
End Function

Private Sub cboUnit_Click()
'功能：读取指定范围内的病人列表
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    Dim i As Integer, j As Integer, k As Integer
    Dim str病人IDs As String, lng病区ID As Long, lngUnitID As Long
    Dim lngColor As Long, lng医嘱处理范围 As Long
    Dim blnIsWowen As Boolean
        
    On Error GoTo errH
    
    lngUnitID = cboUnit.ItemData(cboUnit.ListIndex)
    blnIsWowen = DeptIsWoman(0, Get科室IDs(lngUnitID))
    mbln产科 = blnIsWowen
    If blnIsWowen Then
        '医嘱处理范围
        optBaby(mlng医嘱处理范围).value = True
        fraBaby.Visible = True
        '必须勾选逐个病人打印才显示婴儿
        If mbln逐个病人换页打印 Then
            lng医嘱处理范围 = IIF(optBaby(0).value, -1, IIF(optBaby(1).value, 0, 1))
            mbln婴儿过滤 = True
        Else
            lng医嘱处理范围 = 0
            mbln婴儿过滤 = False
        End If
    Else
        fraBaby.Visible = False
        optBaby(0).value = True
        lng医嘱处理范围 = 0
    End If
    Call SetBabyVisible
    vsPati.Rows = 1
    
    str病人IDs = IIF(mblnOnePati, "", zldatabase.GetPara("发送病人", glngSys, p住院医嘱发送))
    If str病人IDs <> "" And InStr(str病人IDs, ":") > 0 Then
        lng病区ID = Val(Split(str病人IDs, ":")(0))
        str病人IDs = Split(str病人IDs, ":")(1)
    End If
    With vsPati
        Set rsTmp = GetPatiRsByUnit(lngUnitID, 0, False, False, chkOut.value = 1, True, lng医嘱处理范围)
        vsPati.Rows = rsTmp.RecordCount + 1
        For i = 1 To rsTmp.RecordCount
            If Not (blnIsWowen And mbln逐个病人换页打印 And rsTmp!婴儿病区ID <> "" And (rsTmp!婴儿病区ID & "") = mlng病区ID And rsTmp!婴儿姓名 & "" = "") And Not (mblnOnePati And rsTmp!病人ID <> mlng病人ID) Then
                .RowData(i) = Val(rsTmp!病人ID & "")
                .TextMatrix(i, COL_姓名) = IIF(rsTmp!婴儿姓名 & "" = "", rsTmp!姓名 & "", rsTmp!婴儿姓名 & "")
                If rsTmp!婴儿姓名 & "" <> "" Then .Cell(flexcpPicture, i, COL_姓名) = imgPati.ListImages("Child").Picture
                .TextMatrix(i, COL_住院号) = IIF(IsNull(rsTmp!住院号), "", rsTmp!住院号)
                .TextMatrix(i, COL_床位) = IIF(IsNull(rsTmp!床号), "", rsTmp!床号)
                .TextMatrix(i, COL_住院医师) = IIF(IsNull(rsTmp!住院医师), "", rsTmp!住院医师)
                .TextMatrix(i, COL_费别) = IIF(IsNull(rsTmp!费别), "", rsTmp!费别)
                .TextMatrix(i, COL_护理等级) = IIF(IsNull(rsTmp!护理等级), "", rsTmp!护理等级)
                .TextMatrix(i, COL_科室) = IIF(IsNull(rsTmp!科室), "", rsTmp!科室)
                .TextMatrix(i, COL_入院时间) = Format(rsTmp!入院日期, "yyyy-MM-dd HH:mm")
                .TextMatrix(i, COL_病人类型) = NVL(rsTmp!病人类型)
                .TextMatrix(i, COL_主页ID) = rsTmp!主页ID & ""
                .TextMatrix(i, COL_婴儿) = rsTmp!婴儿序号 & ""
                
                '病人颜色
                lngColor = zldatabase.GetPatiColor(NVL(rsTmp!病人类型))
                .Cell(flexcpForeColor, i, COL_住院号, i, COL_住院号) = lngColor
                .Cell(flexcpForeColor, i, COL_病人类型, i, COL_病人类型) = lngColor
                
                '上次是否选择
                If lngUnitID = lng病区ID And str病人IDs <> "" Then
                    If InStr("," & str病人IDs & ",", "," & rsTmp!病人ID & ",") > 0 Then
                        .TextMatrix(i, COL_选择) = -1
                        If k = 0 Then '为了看到有选择的
                            .ShowCell i, COL_姓名
                            k = 1
                        End If
                    End If
                ElseIf rsTmp!病人ID = mlng病人ID Then
                    .TextMatrix(i, COL_选择) = -1
                    .ShowCell i, COL_姓名
                End If
            End If
            rsTmp.MoveNext
        Next
        For i = .Rows - 1 To 1 Step -1
            If Val(.RowData(i) & "") = 0 Then .RemoveItem i
        Next
    End With
    mlng病区ID = lngUnitID
    '勾选了才处理
    If ChkWaitPrint.value = 1 Then
        Call LoadWaitPrint(ChkWaitPrint.value)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ChkWaitPrint_Click()
    Call LoadWaitPrint(ChkWaitPrint.value = 1)
    Call SetBabyVisible
End Sub

Private Sub LoadWaitPrint(ByVal blnIsLoadWart As Boolean)
'功能：过滤待打印的病人
'参数：LoadWaitPrint-true过滤待打印的，否则显示全部
    Dim i As Long
    Dim strFiter As String, str期效 As String
    Dim strSql As String, rsTmp As Recordset
    
    If blnIsLoadWart Then
        '删除不需要打印的
        '婴儿
        If optBaby(1).value = True Then
            strFiter = " And Nvl(a.婴儿, 0) = 0 "
        ElseIf optBaby(2).value = True Then
            strFiter = " And Nvl(a.婴儿, 0) <> 0 "
        End If
        '报表
        Select Case UCase(Mid(lvwReport.SelectedItem.Key, 2))
        Case "ZL1_INSIDE_1254_4" '服药单
            strFiter = strFiter & " And B.操作类型='2' And B.执行分类=4"
        Case "ZL1_INSIDE_1254_5" '注射单
            strFiter = strFiter & " And B.操作类型='2' And B.执行分类=2"
        Case "ZL1_INSIDE_1254_6" '输液单
            strFiter = strFiter & " And B.操作类型='2' And B.执行分类=1"
        End Select
        '期效
        If chk期效(0).value = 1 And chk期效(1).value = 1 Then
            strFiter = strFiter & " And a.医嘱期效 In(0,1) "
        ElseIf chk期效(0).value = 1 And chk期效(1).value = 0 Then
            strFiter = strFiter & " And a.医嘱期效 =0 "
        ElseIf chk期效(0).value = 0 And chk期效(1).value = 1 Then
            strFiter = strFiter & " And a.医嘱期效 =1 "
        End If
        strSql = "Select distinct a.病人id, a.主页id," & IIF(optBaby(1).value = False And fraBaby.Visible And mbln逐个病人换页打印, "NVL(A.婴儿,0)", "0") & " as 婴儿" & vbNewLine & _
                "From 病人医嘱记录 A, 诊疗项目目录 B ,病人信息 C,病案主页 D,在院病人 R" & vbNewLine & _
                "Where a.诊疗项目id = b.Id And a.相关id Is Null And a.校对时间 Is Not Null AND C.病人ID=D.病人ID and C.主页ID=D.主页ID And a.医嘱状态 <> 4 And ([1]" & vbNewLine & _
                "       <= a.执行终止时间 Or a.执行终止时间 Is Null) And [2]" & vbNewLine & _
                "      >= a.开始执行时间  And" & vbNewLine & _
                "      a.病人id=c.病人id and a.主页id=c.主页id and C.病人ID=R.病人ID And C.当前病区ID=R.病区ID and (R.病区id = [4] or D.婴儿病区ID = [4]) And" & vbNewLine & _
                "      Zl_Adviceexecount(a.Id, [1], [2], [5], [3]) > 0" & strFiter
        On Error GoTo errH
        Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, CDate(dtpBegin.value), CDate(dtpEnd.value), lvwReport.SelectedItem.Tag, mlng病区ID, chk重复打印.value)
        
        For i = 1 To vsPati.Rows - 1
            rsTmp.Filter = "病人ID=" & vsPati.RowData(i) & " And 主页ID=" & Val(vsPati.TextMatrix(i, COL_主页ID)) & " And 婴儿=" & Val(vsPati.TextMatrix(i, COL_婴儿))
            If rsTmp.RecordCount = 0 Then
                '隐藏
                vsPati.RowHidden(i) = True
            Else
                '显示
                vsPati.RowHidden(i) = False
            End If
        Next
    Else
        For i = 1 To vsPati.Rows - 1
            '显示
            vsPati.RowHidden(i) = False
        Next
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub chk期效_Click(Index As Integer)
    If chk期效((Index + 1) Mod 2).value = 0 And chk期效(Index).value = 0 Then chk期效(Index).value = 1
    '勾选了才处理
    If ChkWaitPrint.value = 1 Then
        Call LoadWaitPrint(ChkWaitPrint.value)
    End If
End Sub

Private Sub chk重复打印_Click()
    '勾选了才处理
    If ChkWaitPrint.value = 1 Then
        Call LoadWaitPrint(ChkWaitPrint.value)
    End If
End Sub

Private Sub cmdAllPati_Click()
    Call SelectLVW(vsPati, True)
    vsPati.SetFocus
End Sub

Private Sub SelectLVW(objVsg As Object, ByVal blnCheck As Boolean)
    Dim i As Long
    For i = 1 To objVsg.Rows - 1
        objVsg.TextMatrix(i, COL_选择) = IIF(blnCheck, -1, 0)
    Next
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGoOn_Click()
    Call PrintOrPreview(2, 2) '续打
End Sub

Private Sub cmdNoPati_Click()
    Call SelectLVW(vsPati, False)
    vsPati.SetFocus
End Sub

Private Sub cmdPreview_Click()
    Call PrintOrPreview(1) '预览
End Sub

Private Sub cmdPrint_Click()
    Call PrintOrPreview(2) '打印
End Sub

Private Sub PrintOrPreview(ByVal bytMode As Byte, Optional ByVal bytType As Byte)
'参数：bytMode=1-预览,2-打印
'      bytType=0-打印所有医嘱，1-只打印停止到上次打印时间之前的医嘱,2-续打单个病人的执行单
    Dim curDate As Date, strTmp As String, i As Long
    Dim arrPati As Variant, str病人IDs As String
    Dim str期效 As String, str病人 As String
    Dim str重复打印 As String
    Dim datBegin As Date, datEnd As Date
    Dim strReports As String, j As Long, k As Long, z As Long
    Dim lng起始行号 As Long, str起始行号 As String
    Dim strRPTNO As String '1-没有勾选，则取光标定位行；2-有勾选，光标定位行未勾选，则取第一个勾选行；3-有勾选，光标定位在勾选行上，则取光标定位行
    
    mintType = bytType
    If bytType = 1 Then
        datBegin = dtpViewBegin.value: datEnd = dtpViewEnd.value
    Else
        datBegin = dtpBegin.value: datEnd = dtpEnd.value
    End If
    
    If datBegin >= datEnd Then
        MsgBox "开始时间应小于结束时间。", vbInformation, gstrSysName
        IIF(bytType = 1, dtpViewBegin, dtpBegin).SetFocus: Exit Sub
    End If
    
    mstrPrintedID = ""
    
    '保存报表病人设置
    str病人IDs = ""
    arrPati = Array()
    For i = 1 To vsPati.Rows - 1
        If Val(vsPati.TextMatrix(i, COL_选择)) = -1 And vsPati.RowHidden(i) = False Then
            str病人IDs = str病人IDs & "," & vsPati.RowData(i)
            ReDim Preserve arrPati(UBound(arrPati) + 1)
            arrPati(UBound(arrPati)) = vsPati.RowData(i) & "," & vsPati.TextMatrix(i, COL_主页ID) & "," & Val(vsPati.TextMatrix(i, COL_婴儿))
        End If
    Next
    str病人IDs = Mid(str病人IDs, 2)
    For i = 1 To lvwReport.ListItems.Count
        If lvwReport.ListItems(i).Checked Then
            strReports = strReports & "," & lvwReport.ListItems(i).Tag
            strRPTNO = strRPTNO & "," & Mid(lvwReport.ListItems(i).Key, 2)
        End If
    Next
    strReports = Mid(strReports, 2)
    
    If strRPTNO <> "" Then
        If InStr(strRPTNO & ",", "," & Mid(lvwReport.SelectedItem.Key, 2) & ",") > 0 Then
            strRPTNO = Mid(lvwReport.SelectedItem.Key, 2)
        Else
            strRPTNO = Split(strRPTNO, ",")(1)
        End If
    Else
        strRPTNO = Mid(lvwReport.SelectedItem.Key, 2)
    End If

    
    '如果选择行未勾选，则提示用户只打印勾选了的
    If bytMode = 2 And bytType = 0 Then
        If strReports <> "" And lvwReport.SelectedItem.Checked = False Then
            If MsgBox("当前勾选了" & UBound(Split(strReports, ",")) + 1 & "张报表，本次打印只打印勾选了的报表,是否继续？", vbInformation + vbYesNo + vbDefaultButton1, Me.Caption) = vbNo Then
                Exit Sub
            End If
        End If
    End If
    
    '续打检查，必须是逐个病人打印，且单个病人，单个执行单
    If bytType = 2 Then
        If mbln逐个病人换页打印 = False Then
            MsgBox "必须启用""逐个病人换页打印""才允许续打。", vbInformation, gstrSysName
            Exit Sub
        ElseIf UBound(Split(str病人IDs, ",")) <> 0 Then
            MsgBox "只能选择一个病人进行续打。", vbInformation, gstrSysName
            vsPati.SetFocus: Exit Sub
        End If
    End If
    If str病人IDs = "" Then
        MsgBox "请至少选择一个住院病人。", vbInformation, gstrSysName
        vsPati.SetFocus: Exit Sub
    End If

    '逐个打印婴儿医嘱时不保存参数
    If bytType = 0 Then
        If UBound(Split(str病人IDs, ",")) = 0 And Val(str病人IDs) = mlng病人ID Then
            Call zldatabase.SetPara("发送病人", "", glngSys, p住院医嘱发送)
        Else
            Call zldatabase.SetPara("发送病人", cboUnit.ItemData(cboUnit.ListIndex) & ":" & str病人IDs, glngSys, p住院医嘱发送)
        End If
        Call zldatabase.SetPara("执行单打印报表", strReports, glngSys, p住院医嘱发送)
    End If
    '保存设置
    curDate = zldatabase.Currentdate
    If bytType = 0 Then
        Call zldatabase.SetPara("常用报表期效", chk期效(0).value & chk期效(1).value, glngSys, p住院医嘱发送)
        Call zldatabase.SetPara("常用报表开始时点", Format(dtpBegin.value, "HH:mm:ss"), glngSys, p住院医嘱发送)
        Call zldatabase.SetPara("常用报表开始间隔", Int(CDate(Format(dtpBegin.value, "yyyy-MM-dd")) - CDate(Format(curDate, "yyyy-MM-dd"))), glngSys, p住院医嘱发送)
        Call zldatabase.SetPara("常用报表结束时点", Format(dtpEnd.value, "HH:mm:ss"), glngSys, p住院医嘱发送)
        Call zldatabase.SetPara("常用报表结束间隔", Int(CDate(Format(dtpEnd.value, "yyyy-MM-dd")) - CDate(Format(curDate, "yyyy-MM-dd"))), glngSys, p住院医嘱发送)
    End If
    
    '只显示待打印的病人
    '不使用参数保存，1、功能不常用，2、且必须根据执行时间期效进行过滤，下次进入时时间又发生了变化。
    
    '期效条件
    If chk期效(0).value = 1 And chk期效(1).value = 1 Then
        str期效 = "0,1"
    ElseIf chk期效(0).value = 1 Then
        str期效 = "0"
    Else
        str期效 = "1"
    End If
    
    str重复打印 = IIF(chk重复打印.Visible, chk重复打印.value, 0)
    If bytType = 1 Then str重复打印 = "2"
    
    '病人条件
    If mbln逐个病人换页打印 = False Then
        If UBound(arrPati) = 0 Then
            '单个病人
            str病人 = "(" & Mid(arrPati(0), 1, Len(arrPati(0)) - 2) & ")"
        Else
            '多个病人
            strTmp = ""
            For i = 0 To UBound(arrPati)
                strTmp = strTmp & "," & Replace(Mid(arrPati(i), 1, Len(arrPati(i)) - 2), ",", ":")
            Next
            strTmp = Mid(strTmp, 2)
            str病人 = " Select  C1 As 病人ID,C2 As 主页ID From Table(f_Num2list2('" & strTmp & "')) "
        End If
        
        '执行
        If strReports = "" Or bytMode = 1 Then ' 只预览当前行
            Call mobjReport.ReportOpen(gcnOracle, glngSys, strRPTNO, Me, _
                "开始时间=" & Format(datBegin, "yyyy-MM-dd HH:mm:ss"), _
                "结束时间=" & Format(datEnd, "yyyy-MM-dd HH:mm:ss"), _
                "期效=" & str期效, "病人=" & str病人, "起始行号=1", "续打病人ID=0", "重复打印=" & str重复打印, "报表ID=" & lvwReport.SelectedItem.Tag, "医嘱处理范围=" & IIF(optBaby(0).value, -1, IIF(optBaby(1).value, 0, -2)), bytMode)
        Else
            '批量打印
            For i = 1 To lvwReport.ListItems.Count
                If lvwReport.ListItems(i).Checked Then
                    Call mobjReport.ReportOpen(gcnOracle, glngSys, Mid(lvwReport.ListItems(i).Key, 2), Me, _
                        "开始时间=" & Format(datBegin, "yyyy-MM-dd HH:mm:ss"), _
                        "结束时间=" & Format(datEnd, "yyyy-MM-dd HH:mm:ss"), _
                        "期效=" & str期效, "病人=" & str病人, "续打病人ID=0", "起始行号=1", "重复打印=" & str重复打印, "报表ID=" & lvwReport.ListItems(i).Tag, "医嘱处理范围=" & IIF(optBaby(0).value, -1, IIF(optBaby(1).value, 0, -2)), bytMode)
                End If
            Next
        End If
    Else
        '逐个病人进行
        Screen.MousePointer = 11
        If vsPati.Visible Then vsPati.SetFocus: Me.Refresh
        '批量打印
        '如果是预览或是一个报表都未勾选，则默认选中行报表
        For z = 1 To IIF(strReports = "" Or bytMode = 1 Or bytType = 2, 1, lvwReport.ListItems.Count)
            If lvwReport.ListItems(z).Checked Or (strReports = "" Or bytMode = 1 Or bytType = 2) Then
                For i = 0 To UBound(arrPati)
                    str病人 = "_" & Split(arrPati(i), ",")(0)
                    j = vsPati.FindRow(Val(Split(arrPati(i), ",")(0)))
                    If (optBaby(2).value Or optBaby(0).value) And mbln逐个病人换页打印 And j <> -1 Then
                        '婴儿列表
                        For k = j To vsPati.Rows - 1
                            If Val(Split(arrPati(i), ",")(0)) = vsPati.RowData(k) And Val(vsPati.TextMatrix(k, COL_婴儿)) = Val(Split(arrPati(i), ",")(2)) Then
                                j = k: Exit For
                            End If
                        Next
                    End If
                    If j <> -1 Then
                        If Val(Split(arrPati(i), ",")(2)) <> Val(vsPati.TextMatrix(j, COL_婴儿)) Then
                            
                            j = vsPati.FindRow(Val(Split(arrPati(i), ",")(0)), j + 1)
                        End If
                        vsPati.TextMatrix(j, COL_选择) = -1
                        Call vsPati.ShowCell(j, COL_姓名)
                        vsPati.Refresh: Me.Refresh
                        
                        str病人 = "(" & Mid(arrPati(i), 1, Len(arrPati(i)) - 2) & ")"
                        mstr当前病人 = Mid(arrPati(i), 1, Len(arrPati(i)) - 2)
                        '续打时提示
                        lng起始行号 = 1
                        If bytType = 2 Then
                            lng起始行号 = Get起始行号(Val(Split(arrPati(i), ",")(0)), Val(Split(arrPati(i), ",")(1)), Val(Split(arrPati(i), ",")(2)), Val(lvwReport.SelectedItem.Tag))
                            str起始行号 = lng起始行号
                            '确定起始行号
                            If zlCommFun.ShowMsgBox("续打起始行号", "当前将从第" & str起始行号 & "行开始续打，请确认。" & vbCrLf & "如有误请输入正确的起始行号。", "!确定(&O),?取消(&C)", Me, vbInformation, , , , , , "起始行号", 10, str起始行号) <> "确定" Then
                                Exit Sub
                            End If
                            lng起始行号 = Val(str起始行号)
                            If lng起始行号 = 0 Then
                                MsgBox "输入的行数不正确。", vbInformation, Me.Caption
                                Exit Sub
                            End If
                        End If
                        '预览、未勾选报表、续打都针对当前选择的报表，而不是勾选的报表。
                        If strReports = "" Or bytMode = 1 Or bytType = 2 Then ' 只预览当前行
                            Call mobjReport.ReportOpen(gcnOracle, glngSys, strRPTNO, Me, _
                                "开始时间=" & Format(datBegin, "yyyy-MM-dd HH:mm:ss"), _
                                "结束时间=" & Format(datEnd, "yyyy-MM-dd HH:mm:ss"), _
                                "期效=" & str期效, "病人=" & str病人, "重复打印=" & str重复打印, _
                                "报表ID=" & lvwReport.SelectedItem.Tag, "PressWorkFirst=" & IIF(lng起始行号 = 1, 0, 1), "起始行号=" & lng起始行号, _
                                "续打病人ID=" & Val(Split(arrPati(i), ",")(0)), "医嘱处理范围=" & IIF(optBaby(1).value, 0, Val(Split(arrPati(i), ",")(2))), bytMode)
                        Else
                            Call mobjReport.ReportOpen(gcnOracle, glngSys, Mid(lvwReport.ListItems(z).Key, 2), Me, _
                                "开始时间=" & Format(datBegin, "yyyy-MM-dd HH:mm:ss"), _
                                "结束时间=" & Format(datEnd, "yyyy-MM-dd HH:mm:ss"), _
                                "期效=" & str期效, "病人=" & str病人, "重复打印=" & str重复打印, _
                                "报表ID=" & lvwReport.ListItems(z).Tag, "PressWorkFirst=" & IIF(lng起始行号 = 1, 0, 1), "起始行号=" & lng起始行号, _
                                "续打病人ID=" & Val(Split(arrPati(i), ",")(0)), "医嘱处理范围=" & IIF(optBaby(1).value, 0, Val(Split(arrPati(i), ",")(2))), bytMode)
                        End If
                        '只预览第一个病人的数据
                        If bytMode = 1 And i = 0 Then Exit For
                    End If
                Next
            End If
        Next
        Screen.MousePointer = 0
    End If
    If bytMode = 2 Then
        '勾选了才处理
        If ChkWaitPrint.value = 1 Then
            Call LoadWaitPrint(ChkWaitPrint.value)
        End If
    End If
End Sub

Private Function Get起始行号(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal int婴儿 As Integer, ByVal lng报表ID As Long) As Long
'功能：取对应报表对应病人（或婴儿）的上次打印的末次行号
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, lngRow As Long
    
    On Error GoTo errH
    
    strSql = "select 末页末行号 from 病人执行单打印 where 病人ID=[1] And 主页ID=[2]  And 报表ID=[4] And 婴儿=[3]"
    Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, lng病人ID, lng主页ID, int婴儿, lng报表ID)
    If rsTmp.RecordCount > 0 Then lngRow = Val(rsTmp!末页末行号 & "")
    Get起始行号 = lngRow + 1
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
                                
Private Sub cmdRePrint_Click()
     Call PrintOrPreview(2, 1) '打印停止到上次打印时间之前的医嘱
End Sub

Private Sub cmdSelect_Click()
    Call LoadAdvice
End Sub

Private Sub cmdSetup_Click()
    Call mobjReport.ReportPrintSet(gcnOracle, glngSys, Mid(lvwReport.SelectedItem.Key, 2), Me)
End Sub

Private Sub cmdView_Click()
    If cmdView.Caption = "隐藏已打印信息" Then
        cmdView.Caption = "显示已打印信息"
        PicView.Visible = False
        If Me.WindowState = 0 Then Me.Width = fraCondition.Left + fraCondition.Width + 650
    Else
        cmdView.Caption = "隐藏已打印信息"
        PicView.Visible = True
        If Me.WindowState = 0 Then Me.Width = PicView.Left + PicView.Width + 580
    End If
    Call Form_Resize
End Sub

Private Sub dtpBegin_Change()
    '勾选了才处理
    If ChkWaitPrint.value = 1 Then
        Call LoadWaitPrint(ChkWaitPrint.value)
    End If
End Sub

Private Sub dtpEnd_Change()
    '勾选了才处理
    If ChkWaitPrint.value = 1 Then
        Call LoadWaitPrint(ChkWaitPrint.value)
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long, j As Long
    
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        If cmdAllPati.Visible Then Call cmdAllPati_Click
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        If cmdNoPati.Visible Then Call cmdNoPati_Click
    ElseIf KeyCode = vbKeyS And Shift = vbCtrlMask Then
        Call cmdSetup_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_Load()
    Dim curDate As Date
    Dim strTmp As String, lngTmp As Long
        
    Call InitReports '读取报表
    If lvwReport.ListItems.Count = 0 Then
        MsgBox "你没有权限打印任何一张报表，请与系统管理员联系。", vbInformation, gstrSysName
        Unload Me: Exit Sub
    End If
    
    If mobjReport Is Nothing Then Set mobjReport = New clsReport
    '缺省结束时间
    curDate = zldatabase.Currentdate
    
    '缺省医嘱期效
    strTmp = zldatabase.GetPara("常用报表期效", glngSys, p住院医嘱发送, "11", Array(chk期效(0), chk期效(1)))
    chk期效(0).value = Val(Left(strTmp, 1))
    chk期效(1).value = Val(Right(strTmp, 1))

    
    strTmp = zldatabase.GetPara("常用报表开始时点", glngSys, p住院医嘱发送, "00:00:00", Array(dtpBegin))
    lngTmp = Val(zldatabase.GetPara("常用报表开始间隔", glngSys, p住院医嘱发送, "0", Array(dtpBegin)))
    dtpBegin.value = Format(curDate + lngTmp, "yyyy-MM-dd " & strTmp)
    dtpViewBegin.value = dtpBegin.value
    
    strTmp = zldatabase.GetPara("常用报表结束时点", glngSys, p住院医嘱发送, "23:59:59", Array(dtpEnd))
    lngTmp = Val(zldatabase.GetPara("常用报表结束间隔", glngSys, p住院医嘱发送, "0", Array(dtpEnd)))
    dtpEnd.value = Format(curDate + lngTmp, "yyyy-MM-dd " & strTmp)
    dtpViewEnd.value = dtpEnd.value
    
    If mblnOnePati = False Then mbln逐个病人换页打印 = Val(zldatabase.GetPara("常用报表逐个打印", glngSys, p住院医嘱发送, "0")) = 1
    mlng医嘱处理范围 = Val(zldatabase.GetPara("医嘱处理范围", glngSys, p住院医嘱发送, "0"))
    If mblnOnePati Then
        cboUnit.Enabled = False
        mbln逐个病人换页打印 = True
        ChkWaitPrint.Visible = False
    End If
    Call InitUnits '读取病区/病人
    
    Call zlControl.LvwFlatColumnHeader(lvwReport)
    '支持排序
    vsPati.ExplorerBar = flexExSort
    vsPati.Editable = flexEDKbdMouse

    If Val(zldatabase.GetPara("显示已打印信息", glngSys, p住院医嘱发送, "0")) = 0 Then
        Call cmdView_Click
    End If

    
    Call RestoreWinState(Me, App.ProductName, IIF(mblnOnePati, "OnePati", ""))
    
End Sub

Private Function InitUnits() As Boolean
'功能：初始化住院临床科室
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long, strSql As String
    
    On Error GoTo errH
    
    '包含门诊观察室
    If InStr(mMainPrivs, "全院病人") > 0 Then
        strSql = _
            " Select Distinct A.ID,A.编码,A.名称" & _
            " From 部门表 A,部门性质说明 B " & _
            " Where A.ID=B.部门ID And B.服务对象 in(1,2,3) And B.工作性质='护理'" & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
            " And (A.撤档时间 is NULL or Trunc(A.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " Order by A.编码"
    Else
        '求有权病区：直接所在病区+所在科室所属病区
        strSql = _
            " Select A.ID,A.编码,A.名称,Nvl(C.缺省,0) as 缺省" & _
            " From 部门表 A,部门性质说明 B,部门人员 C" & _
            " Where A.ID=B.部门ID And A.ID=C.部门ID And C.人员ID=[1]" & _
            " And B.服务对象 in(1,2,3) And B.工作性质='护理'" & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
            " And (A.撤档时间 is NULL or Trunc(A.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))"
        strSql = strSql & " Union " & _
            " Select C.ID,C.编码,C.名称,Nvl(B.缺省,0) as 缺省" & _
            " From 病区科室对应 A,部门人员 B,部门表 C" & _
            " Where A.病区ID=C.ID And B.部门ID=A.科室ID And B.人员ID=[1]" & _
            " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
            " And (C.撤档时间 is NULL or Trunc(C.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))"
        strSql = "Select ID,编码,名称,Max(缺省) as 缺省 From (" & strSql & ") Group by ID,编码,名称 Order by 编码"
    End If
    
    cboUnit.Clear
    Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.ID)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cboUnit.AddItem rsTmp!编码 & "-" & rsTmp!名称
            cboUnit.ItemData(cboUnit.NewIndex) = rsTmp!ID
            If rsTmp!ID = mlng病区ID Then cboUnit.ListIndex = cboUnit.NewIndex
            rsTmp.MoveNext
        Next
    End If
    InitUnits = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.Width < 14205 And PicView.Visible Then Me.Width = 14205
    cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width - 120
    cmdPrint.Left = cmdCancel.Left - cmdPrint.Width - 240
    cmdPreview.Left = cmdPrint.Left - cmdPreview.Width - 30
    cmdGoOn.Left = cmdPreview.Left - cmdGoOn.Width - 120
    
    fraDetail.Width = Me.ScaleWidth - 240
    fraDetail.Height = Me.ScaleHeight - picBottom.Height - 120
    
    fraReport.Width = fraDetail.Width - IIF(PicView.Visible, PicView.Width, 0) - 240
    lvwReport.Width = fraDetail.Width - IIF(PicView.Visible, PicView.Width, 0) - lvwReport.Left - fraCondition.Width - 300
    lvwReport.ColumnHeaders(1).Width = lvwReport.Width - 140
        
    lvwReport.Height = fraReport.Height - lvwReport.Top - 60
    fraCondition.Left = lvwReport.Left + lvwReport.Width + 120
    
    fraline.Top = fraReport.Top + fraReport.Height
    fraline.Width = fraDetail.Width - IIF(PicView.Visible, PicView.Width, 0) - fraline.Left - 120
    
    PicView.Left = fraCondition.Width + fraCondition.Left + 100
    PicView.Height = Me.Height - 1400
    If vsAdvice.Visible Then vsAdvice.Height = Me.Height - 1300
    
    fraPati.Width = fraDetail.Width - IIF(PicView.Visible, PicView.Width, 0) - fraPati.Left - 120
    vsPati.Width = fraPati.Width - vsPati.Left
    
    fraPati.Top = fraline.Top + fraline.Height + 60
    fraPati.Height = fraDetail.Height - fraline.Top - 120
    vsPati.Height = fraPati.Height - vsPati.Top - 60
    
    cmdNoPati.Top = vsPati.Top + vsPati.Height - 30 - cmdNoPati.Height
    cmdAllPati.Top = cmdNoPati.Top - cmdAllPati.Height - 30
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mMainPrivs = ""
    mlng病人ID = 0
    mlng病区ID = 0
    Call zldatabase.SetPara("显示已打印信息", IIF(PicView.Visible, 1, 0), glngSys, p住院医嘱发送)
    'Set mobjReport = Nothing '自动缓存以便报表部件中的缓存能重复使用
    
    Call SaveWinState(Me, App.ProductName, IIF(mblnOnePati, "OnePati", ""))
End Sub

Private Function InitReports() As Boolean
'功能：读取可用报表
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Long
    Dim objItem As ListItem
    Dim strReports As String
    
    On Error GoTo errH
    
    strReports = zldatabase.GetPara("执行单打印报表", glngSys, p住院医嘱发送)
    strSql = "Select ID,编号,名称,功能 From zlReports Where 系统=[1] And 编号 IN('ZL1_INSIDE_1254_4','ZL1_INSIDE_1254_5','ZL1_INSIDE_1254_6'" & _
         ",'ZL1_INSIDE_1254_7','ZL1_INSIDE_1254_8','ZL1_INSIDE_1254_9','ZL1_INSIDE_1254_10','ZL1_INSIDE_1254_11'" & _
         ",'ZL1_INSIDE_1254_12','ZL1_INSIDE_1254_13','ZL1_INSIDE_1254_14','ZL1_INSIDE_1254_15','ZL1_INSIDE_1254_16') Order by ID"
    Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, glngSys)
    Do While Not rsTmp.EOF
        If InStr(GetInsidePrivs(p住院医嘱发送), ";" & rsTmp!功能 & ";") > 0 Then
            Set objItem = lvwReport.ListItems.Add(, "_" & rsTmp!编号, rsTmp!名称, , 1)
            objItem.Tag = Val(rsTmp!ID)
            If InStr("," & strReports & ",", "," & Val(rsTmp!ID) & ",") > 0 And strReports <> "" Then
                objItem.Checked = True
            End If
        End If
        rsTmp.MoveNext
    Loop
    InitReports = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub fraline_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        On Error Resume Next
        If fraReport.Height + Y < 1000 Or fraReport.Height - Y < 500 Then Exit Sub
        If fraReport.Height + Y > (fraDetail.Height - cmdAllPati.Height * 7) Then Exit Sub
        
        fraline.Top = fraline.Top + Y
        fraReport.Height = fraReport.Height + Y
        fraPati.Top = fraPati.Top + Y
        fraPati.Height = fraPati.Height - Y
        
        Call Form_Resize
    End If
End Sub

Private Sub lvwReport_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Item.Selected = True
End Sub

Private Sub lvwReport_ItemClick(ByVal Item As MSComctlLib.ListItem)
    '勾选了才处理
    If ChkWaitPrint.value = 1 Then
        Call LoadWaitPrint(ChkWaitPrint.value)
    End If
End Sub

Private Sub mobjReport_PrintSheetRow(ByVal ReportNum As String, Sheet As Object, ByVal Page As Integer, ByVal Row As Long, ByVal ID As Long)
'功能：报表数据打印事件，记录医嘱打印行数据
'说明：当表格行无数据要打印时，是不会激活该事件的
    If ID <> 0 Then
        If InStr(mstrPrintedID & ",", "," & ID & ",") = 0 Then
            mstrPrintedID = mstrPrintedID & "," & ID
        End If
        mlngLastRow = Row
    End If
End Sub

Private Sub mobjReport_AfterPrint(ByVal ReportNum As String)
'功能：打印之后，更新医嘱的上次打印时间
    Dim rsTmp As ADODB.Recordset
    Dim arrPati As Variant, arrSQL As Variant
    Dim strSql As String, i As Long
    Dim strSQLPati As String, strPatis As String, strTemp As String
    Dim strThis As String, p As Long, n As Long, lngParStar As Long
    Dim varPar(0 To 10) As String, blnTrans As Boolean, lngReportID As Long
    
    On Error GoTo errH
    
    If mstrPrintedID <> "" Then
        mstrPrintedID = Mid(mstrPrintedID, 2)
        n = 0
        Do While True
            If Len(mstrPrintedID) < 4000 Then
                p = Len(mstrPrintedID) + 1
            Else
                p = InStrRev(Mid(mstrPrintedID, 1, 4000), ",")
            End If
            strThis = Mid(mstrPrintedID, 1, p - 1)
            
            If n > 10 Then
                '太长不再处理，使之报错
                varPar(10) = varPar(10) & "," & strThis
            Else
                varPar(n) = strThis
            End If
            
            n = n + 1
            mstrPrintedID = Mid(mstrPrintedID, p + 1)
            If mstrPrintedID = "" Then Exit Do
        Loop
        For i = 1 To lvwReport.ListItems.Count
            If Mid(lvwReport.ListItems(i).Key, 2) = ReportNum Then lngReportID = Val(lvwReport.ListItems(i).Tag): Exit For
        Next
        arrSQL = Array()
        For i = 0 To UBound(varPar)
            If varPar(i) <> "" Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_医嘱执行单_打印('" & varPar(i) & "'," & lngReportID & "," & _
                    "To_Date('" & Format(IIF(mintType = 1, dtpViewBegin.value, dtpBegin.value), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                    "To_Date('" & Format(IIF(mintType = 1, dtpViewEnd.value, dtpEnd.value), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & IIF(mbln逐个病人换页打印, mlngLastRow, 0) & ")"
            End If
        Next
    
    Else
        '医嘱期效
        If chk期效(0).value = 1 And chk期效(1).value = 1 Then
            strSql = strSql & " And A.医嘱期效 IN(0,1)"
        ElseIf chk期效(0).value = 1 Then
            strSql = strSql & " And A.医嘱期效=0"
        Else
            strSql = strSql & " And A.医嘱期效=1"
        End If
        
        '治疗类别
        Select Case UCase(Mid(lvwReport.SelectedItem.Key, 2))
        Case "ZL1_INSIDE_1254_4" '服药单
            strSql = " And B.执行分类=4"
        Case "ZL1_INSIDE_1254_5" '注射单
            strSql = " And B.执行分类=2"
        Case "ZL1_INSIDE_1254_6" '输液单
            strSql = " And B.执行分类=1"
        Case Else
            '其他的报表是自定义的，不记录上次打印时间，不支持避免重复打印
            Exit Sub
        End Select
        
        '病人选择
        arrPati = Array()
        For i = 1 To vsPati.Rows - 1
            If Val(vsPati.TextMatrix(i, COL_选择)) = -1 And vsPati.RowHidden(i) = False Then
                ReDim Preserve arrPati(UBound(arrPati) + 1)
                arrPati(UBound(arrPati)) = vsPati.RowData(i) & "," & vsPati.TextMatrix(i, COL_主页ID)
            End If
        Next
        If mbln逐个病人换页打印 Then
            strSql = strSql & " And A.病人ID = [4] And A.主页ID = [5]"
            varPar(0) = Split(mstr当前病人, ",")(0)
            varPar(1) = Split(mstr当前病人, ",")(1)
        Else
            If UBound(arrPati) = 0 Then '单个病人
                strSql = strSql & " And (A.病人ID,A.主页ID) IN((" & arrPati(0) & "))"
            Else
                For i = 0 To UBound(arrPati)
                    strPatis = strPatis & "," & Replace(arrPati(i), ",", ":")
                Next
                        
                strPatis = Mid(strPatis, 2)   '去掉前置,号
                strTemp = "Select a.C1 As 病人ID,a.C2 As 主页ID From Table(f_Num2list2([1])) a"
                n = 0
                lngParStar = 3
                Do While True
                    If Len(strPatis) < 4000 Then
                        p = Len(strPatis) + 1
                    Else
                        p = InStrRev(Mid(strPatis, 1, 4000), ",")
                    End If
                    strThis = Mid(strPatis, 1, p - 1)
                    
                    If n > 10 Then
                        strSQLPati = strSQLPati & vbNewLine & " Union All " & Replace(strTemp, "[1]", "'" & strThis & "'")
                    Else
                        varPar(n) = strThis
                        strSQLPati = IIF(strSQLPati = "", "", strSQLPati & vbNewLine & " Union All ") & Replace(strTemp, "[1]", "[" & (lngParStar + n + 1) & "]")
                    End If
                    
                    n = n + 1
                    strPatis = Mid(strPatis, p + 1)
                    If strPatis = "" Then Exit Do
                Loop
                
                strSql = strSql & " And (A.病人ID,A.主页ID) IN(" & strSQLPati & ")"
            End If
        End If
            
        '读取本次打印的医嘱
        strSql = _
            " Select /*+ Rule*/A.ID,zl_AdviceExeCount(A.Id,[1],[2]) As 次数" & _
            " From 病人医嘱记录 A,诊疗项目目录 B" & _
            " Where A.诊疗项目ID=B.ID And A.诊疗类别='E' And B.操作类型='2'" & _
            " And A.校对时间 Is Not Null And A.医嘱状态<>4" & strSql & _
            " And ([1]<=执行终止时间 Or 执行终止时间 Is Null) And [2]>=开始执行时间"
        strSql = "Select * From (" & strSql & ") Where 次数>0"
        
        Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, dtpBegin.value, dtpEnd.value, mlng病区ID, varPar(0), varPar(1), varPar(2), varPar(3), varPar(4), varPar(5), varPar(6), varPar(7), varPar(8), varPar(9), varPar(10))
        
        arrSQL = Array()
        Do While Not rsTmp.EOF
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_治疗医嘱执行_打印(" & rsTmp!ID & "," & _
                "To_Date('" & Format(IIF(mintType = 1, dtpViewBegin.value, dtpBegin.value), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                "To_Date('" & Format(IIF(mintType = 1, dtpViewEnd.value, dtpEnd.value), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
            rsTmp.MoveNext
        Loop
    End If
    mlngLastRow = 0
    '执行提交数据
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSQL)
        zldatabase.ExecuteProcedure CStr(arrSQL(i)), Me.Caption
    Next
    gcnOracle.CommitTrans: blnTrans = False
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub optBaby_Click(Index As Integer)
    If fraBaby.Visible Then
        mlng医嘱处理范围 = Index
        Call cboUnit_Click
    End If
End Sub

Private Sub chkOut_Click()
'显示出院病人
    Call cboUnit_Click
End Sub

Private Sub LoadAdvice()
'功能: 加载医嘱
'参数: 是否加载已执行医嘱 , 为空为加载待审核医嘱
    Dim rsTmp As Recordset
    Dim strSql As String
    Dim i As Long, j As Long
    Dim intBedLen As Integer
    Dim str病人IDs As String
    Dim blnDo As String
    Dim bln给药途径 As Boolean
    
    intBedLen = GetMaxBedLen(mlng病区ID, False)
    str病人IDs = ""
    For i = 1 To vsPati.Rows - 1
        If Val(vsPati.TextMatrix(i, COL_选择)) = -1 And vsPati.RowHidden(i) = False Then
            If InStr(str病人IDs & ",", "," & vsPati.RowData(i) & ",") = 0 Then
                str病人IDs = str病人IDs & "," & vsPati.RowData(i)
            End If
        End If
    Next
    str病人IDs = Mid(str病人IDs, 2)
    If str病人IDs = "" Then Exit Sub
    strSql = "Select /*+ Rule*/ b.Id, b.相关id, e.婴儿姓名,b.诊疗类别, b.姓名, Decode(b.医嘱期效, 0, '长嘱', '临嘱') As 期效, LPAD(c.出院病床," & intBedLen & ",' ') as 床号," & vbNewLine & _
            "       Decode(b.相关id, Null, b.医嘱内容 || ' ' || b.执行频次, b.医嘱内容) As 医嘱内容, b.执行频次, a.上次打印时间, b.执行终止时间" & vbNewLine & _
            "From 医嘱执行打印 A, 病人医嘱记录 B, 病案主页 C,病人信息 D,病人新生儿记录 E,在院病人 R" & vbNewLine & _
            "Where a.医嘱id = NVL(B.相关ID,b.Id)  And e.病人ID(+)=b.病人ID and e.主页ID(+)=b.主页ID And e.序号(+)=b.婴儿 And b.病人id = c.病人id And b.主页id = c.主页id And c.病人id=d.病人id and c.主页id=d.主页id" & _
            " And (R.病区ID=[1] OR C.婴儿病区ID=[1]) and D.病人ID=R.病人ID And D.当前病区ID=R.病区ID And zl_AdviceExeCount(b.Id,[2],[3],1)>0 " & _
            " And R.病人ID In(Select Column_Value From Table(Cast(f_Str2List([4]) As zlTools.t_StrList)))  And A.报表ID=[5] "
            

    strSql = strSql & " Order By 床号,Nvl(b.相关id, b.Id),b.序号"
    

    On Error GoTo errH
    Set rsTmp = zldatabase.OpenSQLRecord(strSql, Me.Caption, mlng病区ID, CDate(dtpViewBegin.value), CDate(dtpViewEnd.value), str病人IDs, Val(lvwReport.SelectedItem.Tag))

    With vsAdvice
        .Redraw = flexRDNone
        .Rows = 1
        If rsTmp.RecordCount > 0 Then
            i = 1
            Do While Not rsTmp.EOF
                .AddItem ""
                .RowData(i) = rsTmp!ID & ""
                .TextMatrix(i, col姓名) = rsTmp!姓名 & ""
                .TextMatrix(i, col期效) = rsTmp!期效 & ""
                .TextMatrix(i, COL床位) = rsTmp!床号 & ""
                .TextMatrix(i, COL婴儿) = rsTmp!婴儿姓名 & ""
                .TextMatrix(i, COL频率) = rsTmp!执行频次 & ""
                .TextMatrix(i, COL上次打印时间) = Format(rsTmp!上次打印时间 & "", "yyyy-MM-dd HH:mm")
                .TextMatrix(i, COL停止时间) = Format(rsTmp!执行终止时间 & "", "yyyy-MM-dd HH:mm")
                If Format(rsTmp!执行终止时间 & "", "yyyy-MM-dd HH:mm:ss") < Format(rsTmp!上次打印时间 & "", "yyyy-MM-dd HH:mm:ss") And rsTmp!执行终止时间 & "" <> "" Then
                    '停止在上次打印时间之前的医嘱用绿色标注
                    .Cell(flexcpBackColor, i, col姓名, i, COL停止时间) = &HE1FFE1
                End If
                .TextMatrix(i, col类别) = rsTmp!诊疗类别 & ""
                .TextMatrix(i, col医嘱内容) = rsTmp!医嘱内容 & ""
                .TextMatrix(i, col相关ID) = rsTmp!相关ID & ""
                bln给药途径 = False
                If IsNull(rsTmp!相关ID) And rsTmp!诊疗类别 & "" = "E" Then
                    If Val(.TextMatrix(i - 1, col相关ID)) = .RowData(i) Then
                        If InStr(",5,6,", .TextMatrix(i - 1, col类别)) > 0 Then
                            bln给药途径 = True
                        End If
                    End If
                End If
                '隐蔽一些附加行
                If (InStr(",F,G,D,7,E,C,", rsTmp!诊疗类别) > 0 And Not IsNull(rsTmp!相关ID)) Or bln给药途径 Then
                    .RemoveItem i
                    i = i - 1
                End If
                rsTmp.MoveNext
                i = i + 1
            Loop
        Else
            .AddItem ""
        End If
        vsAdvice.ColHidden(COL婴儿) = Not fraBaby.Visible
        '自动调整行高
        .AutoSize col医嘱内容
        .Redraw = flexRDDirect
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsAdvice_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
        '说明：1.OwnerDraw要设置为Over(画出单元所有内容)
'      2.Cell的GridLine从上下左右向内都是从第1根线开始
'      3.Cell的Border从左上是从第2根线开始,右下是从第1根线开始
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsAdvice
        lngLeft = col期效: lngRight = col期效
        If Not Between(Col, lngLeft, lngRight) Then
            lngLeft = COL频率: lngRight = COL频率
        End If
        If Not Between(Col, lngLeft, lngRight) Then Exit Sub
        If Not RowIn一并给药(Row, lngBegin, lngEnd, vsAdvice) Then Exit Sub
        
        vRect.Left = Left '擦除左边表格线
        vRect.Right = Right - 1 '保留右边表格线
        If Row = lngBegin Then
            vRect.Top = Bottom - 1
            vRect.Bottom = Bottom
        Else
            If Row = lngEnd Then
                vRect.Top = Top
                vRect.Bottom = Bottom - 2 '底行保留下边线(本窗体中用到下边线粗为2)
            Else
                vRect.Top = Top
                vRect.Bottom = Bottom
            End If
        End If
        
        If Between(Row, .Row, .RowSel) Then
            SetBkColor hDC, OS.SysColor2RGB(.BackColorSel)
        Else
            If .Cell(flexcpBackColor, Row, col姓名) = &HE1FFE1 Then
                SetBkColor hDC, OS.SysColor2RGB(&HE1FFE1)
            Else
                SetBkColor hDC, OS.SysColor2RGB(.BackColor)
            End If
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Done = True
    End With
End Sub

Private Function RowIn一并给药(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long, vsTmp As VSFlexGrid) As Boolean
'功能：判断指定行是否在一并给药的范围中,如果是,同时返回行号范围
    Dim i As Long, blnTmp As Boolean
    
    With vsTmp
        If .TextMatrix(lngRow, col类别) = "" Then Exit Function
        If .TextMatrix(lngRow, col类别) = "类别" Then Exit Function
        If Val(.TextMatrix(lngRow - 1, col相关ID)) = Val(.TextMatrix(lngRow, col相关ID)) And Val(.TextMatrix(lngRow, col相关ID)) <> 0 Or Val(.TextMatrix(lngRow - 1, col相关ID)) = Val(.RowData(lngRow)) Or Val(.RowData(lngRow - 1)) = Val(.TextMatrix(lngRow, col相关ID)) And Val(.RowData(lngRow - 1)) <> 0 Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(lngRow + 1, col相关ID)) = Val(.TextMatrix(lngRow, col相关ID)) And Val(.TextMatrix(lngRow + 1, col相关ID)) <> 0 Or Val(.RowData(lngRow + 1)) = Val(.TextMatrix(lngRow, col相关ID)) Or Val(.TextMatrix(lngRow + 1, col相关ID)) = Val(.RowData(lngRow)) Then
                blnTmp = True
            End If
        End If
        lngBegin = lngRow
        For i = lngRow - 1 To .FixedRows Step -1
            If Val(.TextMatrix(i, col相关ID)) = Val(.TextMatrix(lngRow, col相关ID)) And Val(.TextMatrix(lngRow, col相关ID)) <> 0 And Val(.RowData(i)) <> Val(.RowData(lngRow)) Or Val(.TextMatrix(i, col相关ID)) = Val(.RowData(lngRow)) Or Val(.RowData(i)) = Val(.TextMatrix(lngRow, col相关ID)) And Val(.RowData(i)) <> 0 Then
                lngBegin = i
            Else
                Exit For
            End If
        Next
        lngEnd = lngRow
        For i = lngRow + 1 To .Rows - 1
            If Val(.TextMatrix(i, col相关ID)) = Val(.TextMatrix(lngRow, col相关ID)) And Val(.TextMatrix(lngRow, col相关ID)) <> 0 And Val(.RowData(i)) <> Val(.RowData(lngRow)) Or Val(.RowData(i)) = Val(.TextMatrix(lngRow, col相关ID)) Or Val(.TextMatrix(i, col相关ID)) = Val(.RowData(lngRow)) Then
                lngEnd = i
            Else
                Exit For
            End If
        Next
        RowIn一并给药 = blnTmp
    End With
End Function

Private Sub vsPati_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> COL_选择 Then Cancel = True
End Sub

Private Sub vsPati_KeyDown(KeyCode As Integer, Shift As Integer)
    If vsPati.Row > 0 And KeyCode = vbKeySpace And vsPati.Col <> COL_选择 Then
        vsPati.TextMatrix(vsPati.Row, COL_选择) = IIF(Val(vsPati.TextMatrix(vsPati.Row, COL_选择)) = -1, 0, -1)
    End If
End Sub

Private Sub SetBabyVisible()
'功能：设置婴儿过滤条件的可见性
    Dim blnTmp As Boolean
    If mbln产科 Then
        If mbln婴儿过滤 Then
            blnTmp = True
        Else
            If ChkWaitPrint.value = 1 Then
                blnTmp = True
            End If
        End If
    End If
    fraBaby.Visible = blnTmp
End Sub
