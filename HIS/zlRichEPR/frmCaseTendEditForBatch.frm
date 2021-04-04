VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCaseTendEditForBatch 
   Caption         =   "批量录入"
   ClientHeight    =   8700
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15615
   Icon            =   "frmCaseTendEditForBatch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8700
   ScaleWidth      =   15615
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView lvw房间号 
      Height          =   1725
      Left            =   8520
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   660
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   3043
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgRow"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "名称"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.PictureBox picQuery 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2475
      Left            =   0
      ScaleHeight     =   2475
      ScaleWidth      =   3675
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   5850
      Width           =   3675
      Begin zlRichEPR.usrTendEditor mfrmCaseTendEditForSinglePerson 
         Height          =   1875
         Left            =   420
         TabIndex        =   25
         Top             =   300
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   3307
      End
   End
   Begin MSComctlLib.ImageList imgRow 
      Left            =   4950
      Top             =   2970
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaseTendEditForBatch.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaseTendEditForBatch.frx":686E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picLocate 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   60
      ScaleHeight     =   315
      ScaleWidth      =   2115
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   720
      Width           =   2115
      Begin VB.TextBox txt床号 
         Height          =   300
         Left            =   960
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   0
         Width           =   1155
      End
      Begin VB.Label lbl定位 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "按床号定位"
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   15
         TabIndex        =   15
         Top             =   60
         Width           =   900
      End
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4485
      Left            =   0
      ScaleHeight     =   4485
      ScaleWidth      =   10545
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1350
      Width           =   10545
      Begin MSComctlLib.ListView lvwMultiSel 
         Height          =   1725
         Left            =   3090
         TabIndex        =   21
         Top             =   1350
         Visible         =   0   'False
         Width           =   3285
         _ExtentX        =   5794
         _ExtentY        =   3043
         View            =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "名称"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.PictureBox picInput 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   500
         Left            =   30
         ScaleHeight     =   495
         ScaleWidth      =   945
         TabIndex        =   17
         Top             =   3810
         Visible         =   0   'False
         Width           =   945
         Begin VB.CommandButton cmd未记说明 
            Caption         =   "‥"
            Height          =   225
            Left            =   630
            TabIndex        =   19
            Top             =   30
            Width           =   255
         End
         Begin VB.ComboBox cbo部位 
            Height          =   300
            Left            =   0
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   0
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.TextBox txt数据 
            Height          =   500
            Left            =   0
            TabIndex        =   18
            Top             =   0
            Width           =   945
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid Vsf 
         Height          =   4335
         Left            =   0
         TabIndex        =   22
         Top             =   30
         Width           =   10425
         _cx             =   18389
         _cy             =   7646
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
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16764057
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   600
         RowHeightMax    =   2000
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmCaseTendEditForBatch.frx":D0D0
         ScrollTrack     =   -1  'True
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
         Begin VB.ListBox lst体温标识 
            Height          =   600
            ItemData        =   "frmCaseTendEditForBatch.frx":D132
            Left            =   900
            List            =   "frmCaseTendEditForBatch.frx":D134
            TabIndex        =   27
            Top             =   600
            Width           =   915
         End
      End
   End
   Begin VB.PictureBox picCond 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   15015
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   15015
      Begin VB.OptionButton optLevel 
         Caption         =   "纵向"
         Height          =   255
         Index           =   1
         Left            =   14160
         TabIndex        =   29
         Top             =   33
         Width           =   735
      End
      Begin VB.OptionButton optLevel 
         Caption         =   "横向"
         Height          =   255
         Index           =   0
         Left            =   13320
         TabIndex        =   28
         Top             =   33
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.PictureBox PicPati 
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   6240
         ScaleHeight     =   300
         ScaleWidth      =   1500
         TabIndex        =   7
         Top             =   0
         Width           =   1500
         Begin VB.ComboBox cbo病人 
            Height          =   300
            Left            =   400
            Style           =   2  'Dropdown List
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   0
            Width           =   1065
         End
         Begin VB.Label lbl病人 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "病人"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   0
            TabIndex        =   8
            Top             =   60
            Width           =   360
         End
      End
      Begin VB.CommandButton cmd房间号 
         Caption         =   "房间号"
         Height          =   320
         Left            =   7740
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   0
         Width           =   765
      End
      Begin VB.CommandButton cmd刷新 
         Caption         =   "刷新(&R)"
         Height          =   320
         Left            =   11220
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   0
         Width           =   915
      End
      Begin VB.ComboBox cbo护理等级 
         Height          =   300
         Left            =   4800
         Style           =   2  'Dropdown List
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   0
         Width           =   1365
      End
      Begin VB.ComboBox cbo科室 
         Height          =   300
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   0
         Width           =   1425
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   300
         Left            =   450
         TabIndex        =   2
         Top             =   0
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   174063619
         UpDown          =   -1  'True
         CurrentDate     =   38702
      End
      Begin VB.Label lblEntry 
         AutoSize        =   -1  'True
         Caption         =   "录入方式"
         Height          =   180
         Left            =   12360
         TabIndex        =   30
         Top             =   65
         Width           =   720
      End
      Begin VB.Label lbl房间号清单 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "所有房间"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   8580
         TabIndex        =   11
         Top             =   60
         Width           =   2505
      End
      Begin VB.Label lbl等级 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "等级"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4380
         TabIndex        =   5
         Top             =   60
         Width           =   360
      End
      Begin VB.Label lbl科室 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "科室"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2400
         TabIndex        =   3
         Top             =   60
         Width           =   360
      End
      Begin VB.Label lbl时间 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "时间"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   60
         TabIndex        =   1
         Top             =   60
         Width           =   360
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   24
      Top             =   8340
      Width           =   15615
      _ExtentX        =   27543
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmCaseTendEditForBatch.frx":D136
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   22463
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmCaseTendEditForBatch.frx":D9C8
      Left            =   450
      Top             =   30
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmCaseTendEditForBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmParent As Object
Private mblnInit As Boolean
Private mblnData As Boolean                 '本次是否提取出数据?
Private mstrSel As String                   '复制行:1;复制某单元格:1.1
Private mblnShow As Boolean                 '是否显示录入框
Private mblnChange As Boolean               '是否修改数据
Private mintPreDays As Long
Private mstrMaxDate As String
Private mstrSelItems As String              '保存用户本次增加的列，以免刷新后重新设置
Private mstrTime As String                  '保存读取数据后的有效时间,以便检查时间是否在提取数据后进行了修改
Private mstr房间号 As String
Private mlng病人数 As Long                  '当前显示的病人数
Private mbyt病人 As Integer                     '-1 所有,0 母亲 ,1 婴儿 (妇产科可以选择，其他科室默认为所有)
Private mblnRefresh As Boolean              '是否刷新过数据
Private mblnCheckVersion As Boolean
Private mobjExtendedBar As CommandBar
Private mstrScope As String
Private mdtOutbegin As Date, mdtOutEnd As Date

'以下变量的值,在ENTERCELL中更新
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mlng科室ID As Long
Private mlng病区ID As Long
Private mbyt护理等级 As Long
Private mint婴儿 As Integer
Private mbln心率 As Boolean                 '是否需要录入心率
Private mstrPrivs As String

Private mlngOper As Long                    '手术列号
Private mlngSigner As Long                  '签名人
Private mlngSignTime As Long                '签名时间
Private mlngRecord As Long                  '记录ID
Private mlngGroup As Long                   '组号
Private mlngCert As Long                    '证书ID

Private mrsItems As New ADODB.Recordset             '所有护理记录项目清单
Private mrsSelItems As New ADODB.Recordset          '当前录入的护理记录项目清单
Private mrsPatient As New ADODB.Recordset           '当前录入的护理记录项目清单

Private Const madLongVarCharDefault As Integer = 10          '字符型字段缺省长度
Private Const madDoubleDefault As Integer = 18               '数字型字段缺省长度
Private Const madDbDateDefault As Integer = 20               '日期型字段缺省长度
Private Const p住院护士站 As Long = 1262

Public Event AfterDataChanged()
Public Event AfterArchiveChanged()
Public Event AfterRefresh()
Public Event AfterSelChange(ByVal lngCert As Long)

Dim strFields As String
Dim strValues As String
Dim blnScroll As Boolean

'记录上次选择行,顶行,以便刷新后重新定位
Dim lngLastRow As Long
Dim lngLastTopRow As Long
Dim lngLastPatientID As Long

Private Enum 病人信息
    姓名 = 1
    病人ID
    主页ID
    性别
    住院号
    床号
    护理等级
    体温标识
    有效数据项
End Enum

Private Sub cbo部位_Click()
    If txt数据.Enabled = False Or Val(cbo部位.Tag) = 1 Then txt数据.Text = cbo部位.Text
End Sub

Private Sub cbo护理等级_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmd房间号.SetFocus
End Sub

Private Sub cbo科室_Click()
    Dim arrCode
    Dim blnVisible As Boolean
    Dim lngCount As Long
    
    If cbo科室.ListCount = 0 Then Exit Sub
    If cbo科室.Tag = "" Then cbo科室.Tag = String(cbo科室.ListCount, "[LPF]0"): cbo科室.Tag = Mid(cbo科室.Tag, 4)
    arrCode = Split(cbo科室.Tag, "[LPF]")
    blnVisible = (Val(arrCode(cbo科室.ListIndex)) = 1)
    PicPati.Enabled = blnVisible
    PicPati.Visible = blnVisible
    '设置相关控件坐标
    If blnVisible = False Then
        cmd房间号.Left = PicPati.Left
    Else
        cmd房间号.Left = PicPati.Left + PicPati.Width + 20
    End If
    lbl房间号清单.Left = cmd房间号.Left + cmd房间号.Width + 10
    lvw房间号.Left = lbl房间号清单.Left
    cmd刷新.Left = lbl房间号清单.Left + lbl房间号清单.Width + 50
    lblEntry.Left = cmd刷新.Left + cmd刷新.Width + 50
    optLevel(0).Left = lblEntry.Left + lblEntry.Width + 50
    optLevel(1).Left = optLevel(0).Left + optLevel(0).Width + 20
    picCond.Width = optLevel(1).Left + optLevel(1).Width + 10
    
    '菜单
    Dim cbrControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    Dim objExtendedBar As CommandBar
    
    If mobjExtendedBar Is Nothing Then Exit Sub
    '删除条件菜单项
    For lngCount = mobjExtendedBar.Controls.Count To 1 Step -1
        mobjExtendedBar.Controls(lngCount).Delete
    Next
    Set objExtendedBar = mobjExtendedBar
    With objExtendedBar.Controls
        Set cbrCustom = .Add(xtpControlCustom, 0, "")
        cbrCustom.flags = xtpFlagAlignLeft
        cbrCustom.Handle = Me.picCond.hWnd
        cbrCustom.ToolTipText = "条件"
        
        Set cbrCustom = .Add(xtpControlCustom, 0, "")
        cbrCustom.flags = xtpFlagAlignLeft
        cbrCustom.Handle = Me.picLocate.hWnd
        cbrCustom.ToolTipText = "定位"
    End With
End Sub

Private Sub cbo科室_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cbo护理等级.SetFocus
End Sub


'没有内容的情况下按下键,则应弹出未记说明
'固定/表示录入脉搏短拙与物理降温
'在输入内容后按下键则弹出部位或方式
'用*或小键盘上其它字符代替下键

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then
        Bottom = stbThis.Height
    Else
        Bottom = 0
    End If
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnEnable As Boolean
    Dim blnData As Boolean      '有数据则为真
    Dim blnClear As Boolean     '清除吗?
    Dim strRecord As String
    Dim strSymbol As String
    Dim strSelItems As String
    Dim strDelItem As String
    Dim lngOrder As Long
    Dim introw As Integer, intCol As Integer, intRowSel As Integer, intColSel As Integer
    Dim intRow_ As Integer, intCol_ As Integer
    
    Select Case Control.ID
    Case conMenu_Edit_Copy
        mstrSel = mfrmCaseTendEditForSinglePerson.GetCopyData
    Case conMenu_Edit_PASTE
        Call PasteData
    Case conMenu_Edit_Clear
        '依次将有数据的列找出来,将其colData设置为1,然后将所选单元格的内容清空
        blnEnable = picInput.Visible
        introw = vsf.ROW
        intCol = vsf.Col
        intRowSel = vsf.RowSel
        intColSel = vsf.ColSel
        
        If vsf.ROW > vsf.RowSel Then introw = vsf.RowSel: intRowSel = vsf.ROW
        If vsf.Col > vsf.ColSel Then intCol = vsf.ColSel: intColSel = vsf.Col
        If intColSel >= mlngSigner Then intColSel = mlngSigner - 1
        
        For intRow_ = introw To intRowSel
            For intCol_ = intCol To intColSel
                If vsf.TextMatrix(intRow_, intCol_) <> "" Then
                    '只有记录ID为空的行,才允许删除整行;否则,只能清除除日期,时间外的数据
                    If Not (Val(vsf.TextMatrix(intRow_, mlngRecord)) <> 0 And intCol_ <= 2) Then
                        blnClear = CheckVersion(intRow_, intCol_)
                        
                        If blnClear Then
                            vsf.Cell(flexcpData, intRow_, intCol_) = 1
                            vsf.Cell(flexcpText, intRow_, intCol_) = ""
                            vsf.RowData(intRow_) = 1
                            mblnChange = True
                        End If
                    End If
                End If
            Next
        Next
        
        '对于记录ID为空,且整行无数据的无效行,删除掉
        intRowSel = vsf.Rows - 1        '最后一行永远不删,当做新增空白行,留给用户录入
        intColSel = mlngSigner - 1
        For introw = intRowSel To 1 Step -1
            blnData = False
            For intCol = IIf(Val(vsf.RowData(introw)) = 0, 1, 有效数据项) To intColSel
                If vsf.TextMatrix(introw, intCol) <> "" Then
                    blnData = True
                    Exit For
                End If
            Next
            If Not blnData Then
                If Val(vsf.TextMatrix(introw, mlngRecord)) <> 0 Then   '历史数据隐藏
                    vsf.RowHidden(introw) = True
                Else
                    If introw <> vsf.Rows - 1 Then
                        vsf.RemoveItem introw               '新记录删除
                        '如果删除的行是新产生行,同步删除内部记录集(如果有的话是粘贴时产生的)
                        mrsPatient.Filter = "行=" & introw
                        If mrsPatient.RecordCount <> 0 Then
                            mrsPatient.Delete
                            mrsPatient.Filter = ""
                            If mrsPatient.RecordCount <> 0 Then mrsPatient.MoveFirst
                            Do While Not mrsPatient.EOF
                                If mrsPatient!行 > introw Then
                                    mrsPatient!行 = mrsPatient!行 - 1
                                End If
                                mrsPatient.MoveNext
                            Loop
                            mrsPatient.UpdateBatch
                        End If
                    End If
                End If
            End If
        Next
        
        mrsPatient.Filter = 0
        mblnShow = False
        picInput.Visible = False
        
        '清除选择区域
        vsf.RowSel = vsf.ROW
        vsf.ColSel = vsf.Col
        If vsf.Enabled And vsf.Visible Then vsf.SetFocus
        If blnEnable Then Call Vsf_EnterCell
    Case conMenu_Edit_SPECIALCHAR
        strSymbol = frmInsSymbol.ShowMe(False, 0)
        Me.txt数据.Text = Me.txt数据.Text & strSymbol
    Case conMenu_Edit_Append
        '手术列与签名人之间的列,都是临时添加的项目,这部分项目是按项目序号大小顺序添加的,因此,在手工添加时,也应该保证此顺序,避免刷新后列顺序发生变化
        With mrsSelItems
            '得到已选择项目的序号清单
            If .RecordCount <> 0 Then .MoveFirst
            Do While Not .EOF
                strSelItems = strSelItems & "," & !项目序号
                .MoveNext
            Loop
            If .RecordCount <> 0 Then .MoveFirst
        End With
        strSelItems = strSelItems & ","
        
        '因是多病人,所以此处护理等级始终传-1,婴儿传-1
        strSelItems = frmTendItemChoose.ShowSelect(strSelItems, -1, -1, mlng科室ID)
        If strSelItems = "" Then Exit Sub
        mstrSelItems = mstrSelItems & IIf(mstrSelItems = "", "", vbCrLf) & strSelItems
        
        Call InsertColumn(strSelItems)
    Case conMenu_Edit_Delete
        '如果查询列表中有数据则不允许删除
        intCol = vsf.Col
        intRowSel = vsf.Rows - 1
        For introw = vsf.ROW To intRowSel
            If vsf.TextMatrix(introw, intCol) <> "" Or vsf.Cell(flexcpData, introw, intCol) <> 0 Then
                MsgBox "当前项目有数据，不允许删除！", vbInformation, gstrSysName
                Exit Sub
            End If
        Next
        
        Call DeleteColumn(intCol)
    Case conMenu_Edit_NewItem   '分组
        '增加新行(分组数据)
        '定位病人所在行
        mrsPatient.Filter = "行=" & vsf.ROW
        If mrsPatient.RecordCount <> 0 Then
            strRecord = mrsPatient!病人ID & "|" & mrsPatient!主页ID & "|" & mrsPatient!科室ID & "|" & mrsPatient!婴儿 & "|" & mrsPatient!护理等级 & "|" & mrsPatient!护理等级名称 & "|" & NVL(mrsPatient!匹配列)
        End If
        mrsPatient.Filter = 0
        
        '增加新行并复制当前行的病人基本信息
        introw = vsf.ROW + 1
        If Val(vsf.TextMatrix(introw - 1, 病人ID)) = 0 Then Exit Sub
        
        vsf.Rows = vsf.Rows + 1
        vsf.RowPosition(vsf.Rows - 1) = introw
        '同一个病人的多行数据,只有第一行才显示病人的信息
'        Vsf.TextMatrix(intRow, 姓名) = Vsf.TextMatrix(intRow - 1, 姓名)
'        Vsf.TextMatrix(intRow, 性别) = Vsf.TextMatrix(intRow - 1, 性别)
'        Vsf.TextMatrix(intRow, 住院号) = Vsf.TextMatrix(intRow - 1, 住院号)
'        Vsf.TextMatrix(intRow, 床号) = Vsf.TextMatrix(intRow - 1, 床号)
        vsf.TextMatrix(introw, 病人ID) = vsf.TextMatrix(introw - 1, 病人ID)
        vsf.TextMatrix(introw, 主页ID) = vsf.TextMatrix(introw - 1, 主页ID)
        vsf.Cell(flexcpAlignment, introw, 1, introw, 有效数据项 - 1) = flexAlignLeftCenter
        'Vsf.Cell(flexcpAlignment, intRow, 有效数据项, intRow, Vsf.Cols - 1) = flexAlignCenterCenter
        
        '更新内存记录集
        With mrsPatient
            .Filter = "行>=" & introw
            Do While Not .EOF
                !行 = !行 + 1
                .Update
                .MoveNext
            Loop
            .Filter = 0
        End With
        '添加当前行
        strFields = "行|病人ID|主页ID|科室ID|婴儿|护理等级|护理等级名称|匹配列"
        strValues = introw & "|" & strRecord
        Call Record_Add(mrsPatient, strFields, strValues)
        
    Case conMenu_Edit_Transf_Save '保存
        If SaveME Then Call ShowMe(mfrmParent, mlng病区ID, mstrPrivs, False, False)
    Case conMenu_Edit_Transf_Cancle '取消
        mstrSel = ""
        mblnShow = False
        picInput.Visible = False
        cbsThis.ActiveMenuBar.Visible = False
        cbsThis.RecalcLayout
        
        Call ReadData(True)
        mblnChange = False
        
        Call vsf_AfterRowColChange(1, 1, vsf.ROW, 1)
    Case conMenu_Tool_Sign          '签名
        Call SignMe
    Case conMenu_Manage_ThingDel    '取消签名
        Call UnSignMe
    Case conMenu_Edit_ApplyTo       '新增空行
        vsf.Rows = vsf.Rows + 1
        If Not vsf.RowIsVisible(vsf.Rows - 1) Then
            vsf.TopRow = vsf.Rows - 1
            vsf.ROW = vsf.Rows - 1
        End If
        vsf.Col = 姓名
    End Select
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If mblnInit = False Then Exit Sub
    
    Select Case Control.ID
    Case conMenu_Edit_Copy
        Control.Enabled = (vsf.TextMatrix(vsf.ROW, vsf.Cols - 1) = "") And mlng病人数 <> 0
    Case conMenu_Edit_PASTE
        Control.Enabled = (mstrSel <> "") And (vsf.TextMatrix(vsf.ROW, vsf.Cols - 1) = "")
    Case conMenu_Edit_SPECIALCHAR, conMenu_Edit_Append
        Control.Enabled = (vsf.TextMatrix(vsf.ROW, vsf.Cols - 1) = "")
    Case conMenu_Edit_Clear '签名的数据不允许清除
        Control.Enabled = (vsf.TextMatrix(vsf.ROW, vsf.Cols - 1) = "") And mblnCheckVersion
        
        '如果是多选,则允许清除
        If vsf.RowSel <> vsf.ROW Or vsf.ColSel <> vsf.Col Then Control.Enabled = True
    Case conMenu_Edit_Delete
        Dim blnDel As Boolean
        If mrsSelItems.State = 1 Then
            mrsSelItems.Filter = "列=" & vsf.Col
            If mrsSelItems.RecordCount <> 0 Then
                blnDel = (mrsSelItems!固定 = 0)
            End If
            mrsSelItems.Filter = 0
        End If
        Control.Enabled = (vsf.TextMatrix(vsf.ROW, vsf.Cols - 1) = "") And blnDel
    Case conMenu_Edit_NewItem   '分组
        Control.Enabled = (vsf.TextMatrix(vsf.ROW, vsf.Cols - 1) = "")  '没有归档就允许
    Case conMenu_Edit_ApplyTo   '新增行
        Control.Enabled = (vsf.TextMatrix(vsf.ROW, vsf.Cols - 1) = "")  '没有归档就允许
    Case conMenu_Edit_Transf_Save '保存
        Control.Enabled = mblnChange Or (Format(dtp.Value, "yyyy-MM-dd HH:mm:ss") <> mstrTime And mblnData)
    Case conMenu_Edit_Transf_Cancle '取消
        Control.Enabled = mblnChange
    Case conMenu_Tool_Sign          '签名
        Control.Enabled = Not mblnChange And (vsf.TextMatrix(vsf.ROW, mlngSigner) = "") And (vsf.TextMatrix(vsf.ROW, vsf.Cols - 1) = "") And Val(vsf.TextMatrix(vsf.ROW, mlngRecord)) <> 0
    Case conMenu_Manage_ThingDel    '取消签名
        Control.Enabled = Not mblnChange And (vsf.TextMatrix(vsf.ROW, mlngSigner) <> "") And (vsf.TextMatrix(vsf.ROW, vsf.Cols - 1) = "") And Val(vsf.TextMatrix(vsf.ROW, mlngRecord)) <> 0
    End Select
End Sub

Private Sub cmd房间号_Click()
    Dim lvwItem As ListItem
    Dim intDo As Integer, intMax As Integer
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '根据当前选择的科室,提取房间信息
    
    gstrSQL = " Select distinct 房间号 From 床位状况记录 Where 病区ID=[1] " & IIf(CLng(cbo科室.ItemData(Me.cbo科室.ListIndex)) = -1, "", " And 科室ID=[2]") & " And 房间号 is not null"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng病区ID, CLng(cbo科室.ItemData(Me.cbo科室.ListIndex)))
    With rsTemp
        lvw房间号.ListItems.Clear
        lvw房间号.ListItems.Add , "K0", "所有房间", , 2
        
        Do While Not .EOF
            Set lvwItem = lvw房间号.ListItems.Add(, "K" & .AbsolutePosition, !房间号, , 2)
            '根据用户的选择显示
            If InStr(1, "," & lbl房间号清单.Caption & ",", "," & !房间号 & ",") <> 0 Then lvwItem.Checked = True
            .MoveNext
        Loop
        If InStr(1, "," & lbl房间号清单.Caption & ",", ",所有房间,") <> 0 Then
            lvw房间号.ListItems(1).Checked = True
            Call lvw房间号_ItemCheck(lvw房间号.ListItems(1))
        End If
        
        lvw房间号.Move lbl房间号清单.Left + 100, 900
        lvw房间号.Visible = True
        lvw房间号.SetFocus
    End With
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmd刷新_Click()
    If mblnInit = False Then Exit Sub
    
    mlng科室ID = Me.cbo科室.ItemData(Me.cbo科室.ListIndex)
    mbyt护理等级 = Me.cbo护理等级.ItemData(Me.cbo护理等级.ListIndex)
    mstr房间号 = Me.lbl房间号清单.Caption
    If PicPati.Visible = True Then
        mbyt病人 = cbo病人.ItemData(cbo病人.ListIndex)
    Else
        mbyt病人 = -1
    End If
    mblnRefresh = True
    
    Call ReadData
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = picMain.hWnd
    Case 2
        Item.Handle = picQuery.hWnd
    End Select
End Sub

Private Sub dkpMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Bottom = stbThis.Height
    
End Sub

Private Sub dtp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If vsf.Visible And vsf.Enabled Then vsf.SetFocus
    End If
End Sub

Private Sub dtp_LostFocus()
    If Me.dtp.Tag = Me.dtp.Value Then Exit Sub
    Me.dtp.Tag = Me.dtp.Value
    
    If mblnRefresh Then
        '以前刷新后改条件,则更新条件值
        mlng科室ID = Me.cbo科室.ItemData(Me.cbo科室.ListIndex)
        mbyt护理等级 = Me.cbo护理等级.ItemData(Me.cbo护理等级.ListIndex)
        mstr房间号 = Me.lbl房间号清单.Caption
    End If
    
    Call ReadData
End Sub

Private Sub cmd未记说明_Click()
    If cbo部位.Visible Then
        If Val(cbo部位.Tag) = 0 Then
            Call txt数据_KeyDown(vbKeyDown, vbShiftMask)
        Else
            Call txt数据_KeyDown(vbKeyDown, 0)
            txt数据.Text = ""
            txt数据.SetFocus
        End If
    Else
        Call txt数据_KeyDown(vbKeyW, vbCtrlMask)
    End If
End Sub

Private Sub Form_Activate()
    Call Vsf_EnterCell
    Call vsf_AfterRowColChange(1, 1, vsf.ROW, 1)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If InStr(1, "TXT数据,CBO部位", UCase(Me.ActiveControl.Name)) <> 0 Then
            mblnShow = False
            picInput.Visible = False
            vsf.SetFocus
        End If
    End If
End Sub

Private Sub Form_Load()
    mstrSel = ""
    mstrSelItems = ""
    mstr房间号 = ""
    mblnShow = False
    mblnChange = False
    mblnInit = False
    mblnRefresh = False
    mlng科室ID = 0
    mlng病人数 = 0
    mintPreDays = Val(zlDatabase.GetPara("超期录入护理数据天数", glngSys, 1255, "1"))
    mstrMaxDate = Format(DateAdd("d", mintPreDays, zlDatabase.Currentdate), "yyyy-MM-dd")
    dtp.MaxDate = mstrMaxDate & " 23:59:59"
    dtp.Value = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
    Set mobjExtendedBar = Nothing
    
    Call InitMenuBar
    cbsThis.ActiveMenuBar.Visible = False
    cbsThis.RecalcLayout
    
    Call InitPanelMain
    Call InitEnv            '初始化环境
    
    lst体温标识.Clear
    lst体温标识.AddItem "1次/日"
    lst体温标识.AddItem "2次/日"
    lst体温标识.AddItem "3次/日"
    lst体温标识.AddItem "4次/日"
    lst体温标识.AddItem "5次/日"
    lst体温标识.AddItem "6次/日"
End Sub

Private Sub InitEnv()
    Dim curDate As Date, intDay As Integer
    Dim rsTemp As New ADODB.Recordset
    Dim blnVisible As Boolean
    Dim blnType As Boolean
    On Error GoTo errHand
    
    glngHours = Val(zlDatabase.GetPara("数据补录时限", glngSys))
    mstrScope = zlDatabase.GetPara("病人显示范围", glngSys, p住院护士站, "10000")
    '出院病人时间范围
    curDate = zlDatabase.Currentdate
    intDay = Val(zlDatabase.GetPara("出院病人结束间隔", glngSys, p住院护士站, 7))
    mdtOutEnd = Format(curDate + intDay, "yyyy-MM-dd 23:59:59")
    intDay = Val(zlDatabase.GetPara("出院病人开始间隔", glngSys, p住院护士站, 30))
    mdtOutbegin = Format(mdtOutEnd - intDay, "yyyy-MM-dd 00:00:00")
    
    blnType = Val(GetSetting("ZLSOFT", "私有模块\frmCaseTendEditForBatch\" & gstrUserName, "Value")) = 0
    If blnType Then
        optLevel(0).Value = True
    Else
        optLevel(1).Value = True
    End If
    
    '提取当前病区下的所有科室
    gstrSQL = " Select distinct B.ID,B.编码||'-'||B.名称 AS 科室,decode(nvl(E.工作性质,''),'产科',1,0) 性质" & _
              " From 病区科室对应 A,部门表 B,部门人员 C,人员表 D,部门性质说明 E" & _
              " Where A.科室ID = b.ID And A.科室ID=C.部门ID And C.人员ID=D.ID And A.病区ID = [1]" & _
              IIf(InStr(1, mstrPrivs, "当前病区") <> 0, "", " And D.ID=[2]") & _
              " And B.ID=E.部门ID(+) And E.工作性质(+)='产科'" & _
              " Order by B.编码||'-'||B.名称"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng病区ID, glngUserId)
    With Me.cbo科室
        .Clear
        .Tag = ""
        If InStr(1, mstrPrivs, "当前病区") <> 0 Then
            .AddItem "所有科室"
            .ItemData(.NewIndex) = -1
        End If
        Do While Not rsTemp.EOF
            .AddItem rsTemp!科室
            .ItemData(.NewIndex) = rsTemp!ID
            .Tag = .Tag & "[LPF]" & rsTemp!性质
            If blnVisible = False Then blnVisible = (Val(rsTemp!性质) = 1)
            rsTemp.MoveNext
        Loop
        .Tag = IIf(blnVisible = True, 1, 0) & .Tag
        If Left(.Tag, 5) = "[LPF]" Then .Tag = Mid(.Tag, 6)
        .ListIndex = 0
    End With
    
    '提取所有护理等级
    With Me.cbo护理等级
        .Clear
        .AddItem "所有"
        .ItemData(.NewIndex) = -1
        .AddItem "三级护理"
        .ItemData(.NewIndex) = 3
        .AddItem "二级护理"
        .ItemData(.NewIndex) = 2
        .AddItem "一级护理"
        .ItemData(.NewIndex) = 1
        .AddItem "特级护理"
        .ItemData(.NewIndex) = 0
        .ListIndex = 0
    End With
    
    '添加病人
    With Me.cbo病人
        .Clear
        .AddItem "所有"
        .ItemData(.NewIndex) = -1
        .AddItem "母亲"
        .ItemData(.NewIndex) = 0
        .AddItem "婴儿"
        .ItemData(.NewIndex) = 1
        .ListIndex = 0
    End With
    '打开现存在的所有护理记录项目
    gstrSQL = " Select 项目序号,项目名称,项目类型,项目性质,项目长度,项目小数,项目表示,项目单位,项目值域,护理等级,应用方式" & _
              " From 护理记录项目 B" & _
              " Where B.应用方式<>0 " & _
              " Order by 项目序号"
    Set mrsItems = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitPanelMain()
    Dim objPane As Pane
    
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.UseSplitterTracker = False '实时拖动
    dkpMain.Options.AlphaDockingContext = True
    dkpMain.Options.CloseGroupOnButtonClick = True
    dkpMain.Options.HideClient = True
    
    dkpMain.SetCommandBars cbsThis
    
    Set objPane = dkpMain.CreatePane(1, 100, 200, DockTopOf, Nothing): objPane.Title = "编辑": objPane.Options = PaneNoCaption
    Set objPane = dkpMain.CreatePane(2, 100, 100, DockBottomOf, objPane): objPane.Title = "查询": objPane.Options = PaneNoCaption
End Sub

Private Function InitMenuBar() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    Dim cbrPop As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim rs As ADODB.Recordset
    Dim objExtendedBar As CommandBar
    
    On Error GoTo errHand
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    cbsThis.ActiveMenuBar.Title = "菜单栏"
    cbsThis.ActiveMenuBar.Visible = False
    
    Set cbsThis.Icons = zlCommFun.GetPubIcons
        With cbsThis.Options
        .AlwaysShowFullMenus = False
        .ShowExpandButtonAlways = False
        .UseDisabledIcons = True
        .SetIconSize True, 24, 24
        .LargeIcons = True
    End With

    '------------------------------------------------------------------------------------------------------------------
    '菜单定义
    cbsThis.ActiveMenuBar.Title = "菜单"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    
    
     '快键绑定
    With cbsThis.KeyBindings

        .Add FCONTROL, Asc("S"), conMenu_Edit_Transf_Save
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F2, conMenu_Edit_Transf_Save
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    '工具栏定义
    Set cbrToolBar = cbsThis.Add("标准", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagHideWrap
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Copy, "复制"): cbrControl.ToolTipText = "复制(Ctrl+C)"
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_PASTE, "粘贴"):  cbrControl.ToolTipText = "粘贴(Ctrl+V)"
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Clear, "清除"):   cbrControl.ToolTipText = "清除"
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_SPECIALCHAR, "特殊符号"):  cbrControl.ToolTipText = "插入特殊符号(Ctrl+D)"

        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "分组"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "增加分组(Alt+G)"
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Append, "添加"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "添加项目(Alt+A)"
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除"):  cbrControl.ToolTipText = "删除项目(Alt+D)"
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ApplyTo, "新增"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "新增行(Ctrl+A)"
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Transf_Save, "保存"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "保存(Alt+S)"
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Transf_Cancle, "取消")
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Sign, "签名"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "记录签名(Alt+R)"
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ThingDel, "取消"):  cbrControl.ToolTipText = "取消签名(Alt+U)"
    End With
    
    '过滤条件
    '------------------------------------------------------------------------------------------------------------------
    Set objExtendedBar = cbsThis.Add("条件", xtpBarTop)
    objExtendedBar.ContextMenuPresent = False
    objExtendedBar.ShowTextBelowIcons = False
    objExtendedBar.EnableDocking xtpFlagHideWrap
    With objExtendedBar.Controls
        Set cbrCustom = .Add(xtpControlCustom, 0, "")
        cbrCustom.flags = xtpFlagAlignLeft
        cbrCustom.Handle = Me.picCond.hWnd
        cbrCustom.ToolTipText = "条件"
        
        Set cbrCustom = .Add(xtpControlCustom, 0, "")
        cbrCustom.flags = xtpFlagAlignLeft
        cbrCustom.Handle = Me.picLocate.hWnd
        cbrCustom.ToolTipText = "定位"
    End With
    
    Set mobjExtendedBar = objExtendedBar
    
    For Each cbrControl In cbrToolBar.Controls
        If cbrControl.Type <> xtpControlCustom And cbrControl.Type <> xtpControlLabel Then
            cbrControl.STYLE = xtpButtonIconAndCaption
        End If
    Next
    
     '快键绑定
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("A"), conMenu_Edit_ApplyTo
        .Add FCONTROL, Asc("C"), conMenu_Edit_Copy
        .Add FCONTROL, Asc("V"), conMenu_Edit_PASTE
        .Add FCONTROL, Asc("D"), conMenu_Edit_SPECIALCHAR
        .Add FALT, Asc("C"), conMenu_Edit_Audit
        .Add FALT, Asc("N"), conMenu_Edit_NewItem
        .Add FALT, Asc("A"), conMenu_Edit_Append
        .Add FALT, Asc("D"), conMenu_Edit_Delete
        .Add FALT, Asc("S"), conMenu_Edit_Transf_Save
        .Add FALT, Asc("R"), conMenu_Tool_Sign
        .Add FALT, Asc("U"), conMenu_Edit_Untread
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    InitMenuBar = True
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub InitBill()
    Dim blnLocate As Boolean            '是否找到当前护士的主管科室
    Dim intCol As Integer, intCols As Integer
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    '初始化内存记录集
    'mrsPatient
    strFields = "行," & adDouble & ",18|病人ID," & adDouble & ",18|主页ID," & adDouble & ",18|科室ID," & adDouble & ",18|" & _
                "婴儿," & adDouble & ",18|护理等级," & adDouble & ",18|护理等级名称," & adLongVarChar & ",50|匹配列," & adLongVarChar & ",500"
    Call Record_Init(mrsPatient, strFields)
    'mrsSelItems
    strFields = "列," & adDouble & ",18|项目序号," & adDouble & ",18|项目名称," & adLongVarChar & ",20|固定," & adDouble & ",2"
    Call Record_Init(mrsSelItems, strFields)
    strFields = "列|项目序号|项目名称|固定"
    
    '先添加模板设定的项目
    strSQL = " Select B.项目序号,B.项目名称,B.项目单位,B.项目类型,1 AS 固定" & _
             " From 护理项目模板 A,护理记录项目 B" & _
             " Where a.项目序号 = b.项目序号 And B.应用方式<>0 And A.科室ID=[1] And A.护理等级=-1 And B.适用病人 IN (0,1,2)" & _
             " Order by A.排列序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Me.cbo科室.ItemData(Me.cbo科室.ListIndex))
    If rsTemp.RecordCount = 0 Then
        '按以前的规则提取项目清单供录入
        strSQL = " Select B.项目序号,B.项目名称,B.项目单位,B.项目类型,0 AS 固定" & _
                 " From 护理记录项目 B" & _
                 " Where B.应用方式<>0 And B.适用病人 IN (0,1,2)" & _
                 " Order by B.项目序号"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    End If
    
    With vsf
        intCols = .Cols - 1
        For intCol = 1 To intCols
            .ColHidden(intCol) = False
        Next
        
        .Clear
        .Rows = 2
        .FixedCols = 1
        .Cols = rsTemp.RecordCount + .FixedCols + 有效数据项     '加上姓名性别床号住院号列,再加上固定的手术列
        .RowHeightMin = 600
        .AllowUserResizing = flexResizeColumns
        .ExplorerBar = flexExNone
        .WordWrap = True
        
        .TextMatrix(0, 姓名) = "姓名"
        .TextMatrix(0, 性别) = "性别"
        .TextMatrix(0, 住院号) = "住院号"
        .TextMatrix(0, 床号) = "床号"
        .TextMatrix(0, 护理等级) = "护理等级"
        .TextMatrix(0, 病人ID) = "病人ID"
        .TextMatrix(0, 主页ID) = "主页ID"
        .TextMatrix(0, 体温标识) = "体温标识"
        .ColWidth(0) = 300
        .ColWidth(姓名) = 1700
        .ColWidth(性别) = 500
        .ColWidth(住院号) = 1000
        .ColWidth(床号) = 800
        .ColWidth(护理等级) = 1000
        .ColWidth(病人ID) = 0
        .ColWidth(主页ID) = 0
        .ColWidth(体温标识) = 1000
        
        intCol = 有效数据项
        Do While Not rsTemp.EOF
            If rsTemp!项目名称 Like "舒张压*" And .TextMatrix(0, intCol - 1) Like "收缩压*" Then
                .TextMatrix(0, intCol - 1) = "血压" & IIf(NVL(rsTemp!项目单位) = "", "", vbCrLf & "(" & rsTemp!项目单位 & ")")
                .Cols = .Cols - 1
                intCol = intCol - 1
            Else
                .TextMatrix(0, intCol) = rsTemp!项目名称 & IIf(NVL(rsTemp!项目单位) = "", "", vbCrLf & "(" & rsTemp!项目单位 & ")")
            End If
            .ColWidth(intCol) = 900
            .ColAlignment(intCol) = IIf(rsTemp!项目类型 = 0, flexAlignCenterCenter, flexAlignLeftTop)       '数字则居中显示,非数字以用户录入的数据显示
            
            '将目前已选择的项目加入内存记录集中
            strFields = "列|项目序号|项目名称|固定"
            strValues = intCol & "|" & rsTemp!项目序号 & "|" & rsTemp!项目名称 & "|" & rsTemp!固定
            Call Record_Add(mrsSelItems, strFields, strValues)
            
            intCol = intCol + 1
            rsTemp.MoveNext
        Loop
        '.Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        .MergeCells = flexMergeFree
        .WordWrap = True
        
        '将目前已选择的项目加入内存记录集中
        strFields = "列|项目序号|项目名称|固定"
        strValues = .Cols - 1 & "|0|手术|1"
        Call Record_Add(mrsSelItems, strFields, strValues)
        
        mlngOper = .Cols - 1
        .TextMatrix(0, .Cols - 1) = "手术"
    End With
    
    '检查是否需要录入心率
    mrsSelItems.Filter = "项目序号=-1"
    mbln心率 = (mrsSelItems.RecordCount <> 0)
    mrsSelItems.Filter = 0
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function AddPatient(ByVal StrKey As String, Optional ByVal rsPatient As ADODB.Recordset) As Boolean
    Dim lngRow As Long
    Dim strFind As String
    Dim strItems As String
    Dim strStart As String
    Dim intCol As Integer, intCols As Integer
    Dim int护理等级 As Integer, int婴儿 As Integer, lng科室ID As Long
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    If rsPatient Is Nothing Then
        strStart = Format(dtp.Value, "yyyy-MM-dd HH:mm:ss")
        '单病人,查找到该病人的信息,如果表格中不存在则加入(只能通过点击工具栏的分组增加同一病人的多组数据)
        lngRow = vsf.ROW
        StrKey = UCase(StrKey)
        
        Select Case Left(StrKey, 1)
        Case "-"    '病人ID
            StrKey = Mid(StrKey, 2)
            strFind = " 病人ID=" & StrKey
        Case "+"    '住院号
            StrKey = Mid(StrKey, 2)
            strFind = " 住院号=" & StrKey
        Case Else   '床号
            strFind = " 床号='" & StrKey & "'"
        End Select
        '73204:刘鹏飞,2014-06-09,无法过滤入院入科的病人
        '73097:刘鹏飞,2014-06-09,多次住院的病人会在录入列表出现多行(添加a.住院次数=b.主页ID)
        '58890:刘鹏飞,2013-02-26,在院病人读取性能优化(关联在院病人表进行查询)
        '34版本病人信息添加主页ID
        '提取病人列表
        gstrSQL = " SELECT B.病人ID, B.主页ID, NVL(B.姓名,A.姓名) 姓名,NVL(B.性别,A.性别) 性别,G.信息值 AS 体温标识, B.住院号, B.出院病床 AS 床号,zl_PatitTendGrade(B.病人ID,B.主页ID) AS 护理等级,D.名称 AS 护理等级名称,B.出院科室ID AS 科室ID,0 AS 婴儿" & _
                  " FROM 病人信息 A,病案主页 B,部门表 C,护理等级 D,病人变动记录 F,病案主页从表 G,在院病人 R" & _
                  " Where A.病人ID = b.病人ID And A.主页ID=B.主页ID And NVL(b.主页ID, 0) <> 0 And b.出院科室ID = C.ID " & _
                  " And A.病人ID=F.病人ID And A.主页ID=F.主页ID And (F.开始原因=2 OR (F.开始原因=1 And Nvl(B.状态,0)<>1 And NOT Exists(Select 病人ID From 病人变动记录 Where 开始原因=2 and 病人ID=F.病人ID And 主页ID=F.主页ID))) And F.开始时间<=[5]" & _
                  " And B.病人ID=G.病人ID(+) And B.主页ID=G.主页ID(+) And G.信息名(+)='体温标识' " & _
                  " AND Nvl(B.病案状态,0)<>5 AND B.封存时间 is NULL And B.护理等级ID=D.序号(+) And R.病人ID=A.病人ID And R.病区ID=[3] " & _
                  IIf(mlng科室ID = -1, "", " And R.科室ID=[4]")
        If Val(Mid(mstrScope, 2, 1)) <> 0 Then
            gstrSQL = gstrSQL & _
                  " Union" & _
                  " SELECT B.病人ID, B.主页ID, NVL(B.姓名,A.姓名) 姓名,NVL(B.性别,A.性别) 性别,G.信息值 AS 体温标识, B.住院号, B.出院病床 AS 床号,zl_PatitTendGrade(B.病人ID,B.主页ID) AS 护理等级,D.名称 AS 护理等级名称,B.出院科室ID AS 科室ID,0 AS 婴儿" & _
                  " FROM 病人信息 A,病案主页 B,部门表 C,护理等级 D,病人变动记录 F,病案主页从表 G" & _
                  " Where A.病人ID = b.病人ID And A.主页ID=B.主页ID And NVL(b.主页ID, 0) <> 0 And b.出院科室ID = C.ID And b.当前病区ID + 0 = [3]" & _
                  " And A.病人ID=F.病人ID And A.主页ID=F.主页ID And (F.开始原因=2 OR (F.开始原因=1 And NOT Exists(Select 病人ID From 病人变动记录 Where 开始原因=2 and 病人ID=F.病人ID And 主页ID=F.主页ID))) And F.开始时间<=[5]" & _
                  " And B.病人ID=G.病人ID(+) And B.主页ID=G.主页ID(+) And G.信息名(+)='体温标识' " & _
                  " AND B.出院日期 BETWEEN [1] AND [2] AND Nvl(B.病案状态,0)<>5 AND B.封存时间 is NULL And B.护理等级ID=D.序号(+)" & _
                IIf(mlng科室ID = -1, "", " And B.出院科室ID=[4]")
        End If
        '提取新生儿列表
        gstrSQL = gstrSQL & _
                  " UNION " & _
                  " Select B.病人ID,B.主页ID,NVL(A.婴儿姓名,B.姓名||'之子'||A.序号) AS 姓名,B.性别,G.信息值 AS 体温标识,B.住院号,B.床号,B.护理等级,B.护理等级名称,B.科室ID AS 科室ID,A.序号 AS 婴儿" & _
                  " From 病人新生儿记录 A,(" & gstrSQL & ") B,病案主页从表 G" & _
                  " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID" & _
                  " And A.病人ID=G.病人ID(+) And A.主页ID=G.主页ID(+) And G.信息名(+)='体温标识'||DECODE(A.序号,0,'',A.序号) "
        
        gstrSQL = " Select * From (" & gstrSQL & ") " & _
                  " Where " & strFind
        Set rsPatient = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mdtOutbegin, mdtOutEnd, mlng病区ID, mlng科室ID, CDate(strStart))
        If rsPatient.RecordCount = 0 Then
            MsgBox "没有找到该病人！", vbInformation, gstrSysName
            Exit Function
        End If
    Else
        '批量添加,只可能发生在初始化的时候
        lngRow = vsf.Rows - 1
    End If
    
    '将病人添加到表格中
    intCols = vsf.Cols - 1
    With rsPatient
        Do While Not .EOF
            If lngRow > vsf.Rows - 1 Then vsf.Rows = vsf.Rows + 1
            
            vsf.TextMatrix(lngRow, 姓名) = IIf(!婴儿 <> 0, Space(4), "") & !姓名
            vsf.TextMatrix(lngRow, 性别) = !性别
            vsf.TextMatrix(lngRow, 住院号) = NVL(!住院号)
            vsf.TextMatrix(lngRow, 床号) = NVL(!床号)
            vsf.TextMatrix(lngRow, 护理等级) = NVL(!护理等级名称)
            vsf.TextMatrix(lngRow, 病人ID) = !病人ID
            vsf.TextMatrix(lngRow, 主页ID) = !主页ID
            vsf.TextMatrix(lngRow, 体温标识) = NVL(!体温标识)
            
            '把其它列清空(避免在确定病人信息前录入了数据,再确定病人信息,导致一些不该录入的项目有了数据
            vsf.Cell(flexcpData, lngRow, 有效数据项, lngRow, vsf.Cols - 1) = 0
            vsf.Cell(flexcpText, lngRow, 有效数据项, lngRow, vsf.Cols - 1) = ""
            
            If !护理等级 <> int护理等级 Or !婴儿 <> int婴儿 Or lng科室ID <> !科室ID Then
                strItems = ""
                int护理等级 = !护理等级
                int婴儿 = !婴儿
                lng科室ID = !科室ID
                
                '提取各病人允许编辑的项目
                gstrSQL = " Select B.项目序号" & _
                          " From 护理记录项目 B" & _
                          " Where B.应用方式<>0 And B.护理等级 >= [1] And B.适用病人 IN (0,[2])" & _
                          " And (B.适用科室=1 Or (B.适用科室=2 And Exists (Select 1 From 护理适用科室 D Where D.项目序号=B.项目序号 And D.科室id=[3])))"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CInt(!护理等级), CInt(IIf(!婴儿 = 0, 1, 2)), CLng(!科室ID))
                With rsTemp
                    Do While Not .EOF
                        strItems = strItems & IIf(strItems = "", "", ",") & !项目序号
                        
                        .MoveNext
                    Loop
                End With
            End If
            
            '将信息更新到内存记录集中
            strFields = "行|病人ID|主页ID|科室ID|婴儿|护理等级|护理等级名称|匹配列"
            strValues = lngRow & "|" & !病人ID & "|" & !主页ID & "|" & !科室ID & "|" & !婴儿 & "|" & !护理等级 & "|" & !护理等级名称 & "|" & strItems
            Call Record_Add(mrsPatient, strFields, strValues)
            
            Call DrawBackColor(lngRow)
            
            AddPatient = True
            lngRow = lngRow + 1
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
    End With
    
    vsf.Cell(flexcpAlignment, 1, 姓名, vsf.Rows - 1, 有效数据项 - 1) = flexAlignLeftCenter
    If mblnInit Then Call vsf_AfterRowColChange(vsf.ROW + 1, 1, vsf.ROW, 1)
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub DrawBackColor(ByVal lngRow As Long)
    Dim intCol As Integer
    '将不允许录入的项目列设置为灰色
    
    mrsPatient.Filter = "行=" & lngRow
    If mrsPatient.RecordCount <> 0 Then
        For intCol = 有效数据项 To vsf.Cols - 1
            mrsSelItems.Filter = "列=" & intCol
            If mrsSelItems.RecordCount <> 0 Then
                If intCol <> mlngOper And InStr(1, "," & mrsPatient!匹配列 & ",", "," & mrsSelItems!项目序号 & ",") = 0 Then
                    vsf.Cell(flexcpBackColor, lngRow, intCol) = &HE0E0E0
                End If
            End If
        Next
    End If

    mrsSelItems.Filter = 0
    mrsPatient.Filter = 0
End Sub

Private Sub ReadData(Optional ByVal blnCancel As Boolean = False)
    Dim arrColumn
    Dim intStart As Integer, intEnd As Integer
    
    Dim int心率应用 As Integer
    Dim strPatient As String, strChild As String
    Dim strStart As String, strEnd As String
    Dim rsData As New ADODB.Recordset
    Dim rsPatient As New ADODB.Recordset
    On Error GoTo errHand
    '读取近期多少天的数据
    
    If mblnChange And blnCancel = False Then
        If MsgBox("当前数据还未保存，点“是”进行保存，点“否”将放弃本次修改！", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Call Vsf_EnterCell
            Call SaveData
        End If
    End If
    mblnShow = False
    picInput.Visible = False
    mblnInit = False
    
    mrsItems.Filter = "项目序号=-1"
    If mrsItems.RecordCount <> 0 Then
        int心率应用 = mrsItems!应用方式
    End If
    mrsItems.Filter = 0
    strStart = Format(dtp.Value, "yyyy-MM-dd HH:mm:ss")
    
    '1、先提取出查询时间范围内自己添加的项目,依次加到表格中
    Call InitBill
    Call AddColumns
    
    '将用户选择的列添加进去
    If mstrSelItems <> "" Then
        arrColumn = Split(mstrSelItems, vbCrLf)
        intEnd = UBound(arrColumn)
        For intStart = 0 To intEnd
            Call InsertColumn(arrColumn(intStart))
        Next
    End If
    vsf.Cell(flexcpAlignment, 0, 0, 0, vsf.Cols - 1) = flexAlignCenterCenter
    '73204:刘鹏飞,2014-06-09,无法过滤入院入科的病人
    '73097:刘鹏飞,2014-06-09,多次住院的病人会在录入列表出现多行(添加a.住院次数=b.主页ID)
    '58890:刘鹏飞,2013-02-26,在院病人读取性能优化(关联在院病人表进行查询)
    '提取病人列表
    strPatient = " SELECT B.病人ID, B.主页ID, NVL(B.姓名,A.姓名) 姓名,NVL(B.性别,A.性别) 性别,G.信息值 AS 体温标识, B.住院号, B.出院病床 AS 床号,zl_PatitTendGrade(B.病人ID,B.主页ID) AS 护理等级,D.名称 AS 护理等级名称,B.出院科室ID AS 科室ID,0 AS 婴儿" & _
              " FROM 病人信息 A,病案主页 B,部门表 C,护理等级 D,床位状况记录 E,病人变动记录 F,病案主页从表 G,在院病人 R" & _
              " Where A.病人ID = B.病人ID And A.主页ID=B.主页ID And NVL(b.主页ID, 0) <> 0 And b.出院科室ID = C.ID " & _
              " AND Nvl(B.病案状态,0)<>5 AND B.封存时间 IS NULL And B.护理等级ID=D.序号(+)" & _
              " And A.病人ID=F.病人ID And A.主页ID=F.主页ID And (F.开始原因=2 OR (F.开始原因=1 And Nvl(B.状态,0)<>1 And NOT Exists(Select 病人ID From 病人变动记录 Where 开始原因=2 and 病人ID=F.病人ID And 主页ID=F.主页ID))) And F.开始时间<=[7]" & _
              " And B.病人ID=G.病人ID(+) And B.主页ID=G.主页ID(+) And G.信息名(+)='体温标识' " & _
              " AND A.病人ID=E.病人ID(+) And R.病人ID=A.病人ID And R.病区ID=[3] " & IIf(mlng科室ID = -1, "", " And R.科室ID=[4]") & IIf(lbl房间号清单.Caption = "所有房间", "", " And instr([6],','||E.房间号||',')<>0")
    '提取新生儿列表
    strChild = " Select B.病人ID,B.主页ID,NVL(A.婴儿姓名,B.姓名||'之子'||A.序号) AS 姓名,B.性别,G.信息值 AS 体温标识,B.住院号,B.床号,B.护理等级,B.护理等级名称,B.科室ID AS 科室ID,A.序号 AS 婴儿" & _
              " From 病人新生儿记录 A,(" & strPatient & ") B,病案主页从表 G" & _
              " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID" & _
              " And A.病人ID=G.病人ID(+) And A.主页ID=G.主页ID(+) And G.信息名(+)='体温标识'||DECODE(A.序号,0,'',A.序号) "
    If mbyt病人 = 0 Then '母亲
        strPatient = strPatient
    ElseIf mbyt病人 = 1 Then '婴儿
        strPatient = strChild
    Else
        strPatient = strPatient & " UNION " & strChild
    End If
'    strPatient = strPatient & _
'              " UNION " & _
'              " Select B.病人ID,B.主页ID,NVL(A.婴儿姓名,B.姓名||'之子'||A.序号) AS 姓名,B.性别,G.信息值 AS 体温标识,B.住院号,B.床号,B.护理等级,B.护理等级名称,B.科室ID AS 科室ID,A.序号 AS 婴儿" & _
'              " From 病人新生儿记录 A,(" & strPatient & ") B,病案主页从表 G" & _
'              " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID" & _
'              " And A.病人ID=G.病人ID(+) And A.主页ID=G.主页ID(+) And G.信息名(+)='体温标识'||DECODE(A.序号,0,'',A.序号) "
    
    strPatient = "SELECT * FROM (" & strPatient & ") " & IIf(Me.cbo护理等级.ListIndex = 0, "", " WHERE 护理等级=[5]") & " Order by Lpad(床号,10,' '),婴儿"
    Set rsPatient = zlDatabase.OpenSQLRecord(strPatient, Me.Caption, mdtOutbegin, mdtOutEnd, mlng病区ID, mlng科室ID, mbyt护理等级, "," & lbl房间号清单.Caption & ",", CDate(strStart))
    Call AddPatient("", rsPatient)
    mlng病人数 = rsPatient.RecordCount
    
    '2、提取数据
    gstrSQL = " Select X.* From ("
    If int心率应用 = 2 Then
        gstrSQL = gstrSQL & _
                    "Select C.病人ID,C.主页ID,C.婴儿,A.项目序号,DECODE(A.记录类型,4,A.项目名称, A.记录内容) As 记录结果, " & _
                        "D.项目ID AS 证书ID,Nvl(A.终止版本,A.开始版本) AS 实际版本,D.记录人 AS 签名人,D.项目名称 As 签名时间," & _
                        "Decode(a.记录内容,Null,'',A.体温部位) As 部位,b.记录内容 As 标记,b.记录标记," & _
                        "C.发生时间 As 完成日期,A.记录id,A.记录组号,a.未记说明,C.归档人,a.记录人 " & _
                    " From 病人护理内容 A, 病人护理内容 B,病人护理记录 C,病人护理内容 D " & _
                    " Where C.ID = A.记录id And b.记录id(+)=a.记录id And b.记录组号(+)=a.记录组号 And b.记录标记(+) =1 " & _
                         " AND A.记录类型 =1 AND C.病人来源 = 2 AND NVL(A.记录标记,0) <> 1 " & _
                         " And D.记录类型(+)=5 And D.记录ID(+)=C.ID And D.终止版本(+) Is NULL" & _
                         " AND C.发生时间 = [7] "
    Else
        gstrSQL = gstrSQL & _
                    "Select C.病人ID,C.主页ID,C.婴儿,A.项目序号,DECODE(A.记录类型,4,A.项目名称, A.记录内容) As 记录结果, " & _
                        "D.项目ID AS 证书ID,Nvl(A.终止版本,A.开始版本) AS 实际版本,D.记录人 AS 签名人,D.项目名称 As 签名时间, " & _
                        "Decode(a.记录内容,Null,'',A.体温部位) As 部位,Decode(a.项目序号,2,'',-1,'',b.记录内容) As 标记,Decode(a.项目序号,2,0,-1,0,b.记录标记) As 记录标记," & _
                        "C.发生时间 As 完成日期,A.记录id,A.记录组号,a.未记说明,C.归档人,a.记录人 " & _
                    " From 病人护理内容 A, 病人护理内容 B,病人护理记录 C,病人护理内容 D " & _
                    " Where C.ID = A.记录id And b.记录id(+)=a.记录id And b.记录组号(+)=a.记录组号 And b.记录标记(+) =1 " & _
                         " AND A.记录类型 =1 AND C.病人来源 = 2 AND ((NVL(A.记录标记,0) <> 1 And a.项目序号>0) or a.项目序号=-1 or (a.项目序号=0 and a.记录类型=4)) " & _
                         " And D.记录类型(+)=5 And D.记录ID(+)=C.ID And D.终止版本(+) Is NULL" & _
                         " AND C.发生时间 = [7] "
    End If
    gstrSQL = gstrSQL & _
                "       And a.终止版本 Is Null And b.终止版本 Is Null " & _
                "       And Decode(a.项目序号,2,-1,a.项目序号)=b.项目序号(+)) X,护理记录项目 Y " & _
                "Where Y.项目序号 = X.项目序号 And Nvl(y.应用方式,0)=1 "
    
    '加上手术项目
    gstrSQL = gstrSQL & _
                " UNION " & _
                " Select C.病人ID,C.主页ID,C.婴儿,A.项目序号,DECODE(A.记录类型,4,A.项目名称, A.记录内容) As 记录结果, " & _
                    "D.项目ID AS 证书ID,Nvl(A.终止版本,A.开始版本) AS 实际版本,D.记录人 AS 签名人,D.项目名称 As 签名时间, " & _
                    "Decode(a.记录内容,Null,'',A.体温部位) As 部位,Decode(a.项目序号,2,'',-1,'',b.记录内容) As 标记,Decode(a.项目序号,2,0,-1,0,b.记录标记) As 记录标记," & _
                    "C.发生时间 As 完成日期,A.记录id,A.记录组号,a.未记说明,C.归档人,a.记录人" & _
                " From 病人护理内容 A, 病人护理内容 B,病人护理记录 C,病人护理内容 D " & _
                " Where C.ID = A.记录id And b.记录id(+)=a.记录id And b.记录组号(+)=a.记录组号 And b.记录标记(+) =1 " & _
                "      And a.终止版本 Is Null And b.终止版本 Is Null And D.终止版本(+) Is NULL" & _
                     " AND A.记录类型 =4 AND C.病人来源 = 2 And D.记录类型(+)=5 And D.记录ID(+)=C.ID AND C.发生时间 = [7] "
    
    gstrSQL = " Select A.*,B.科室ID,B.护理等级,B.姓名,B.性别,B.住院号,B.床号 From (" & gstrSQL & ") A,(" & strPatient & ") B" & _
              " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID And A.婴儿=B.婴儿" & _
              " Order By A.病人ID,A.主页ID,A.婴儿,A.记录组号,DECODE(A.项目序号,0,999,A.项目序号)"
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mdtOutbegin, mdtOutEnd, mlng病区ID, mlng科室ID, mbyt护理等级, "," & lbl房间号清单.Caption & ",", CDate(strStart))
    
    '准备添加数据(遇到没有的项目,直接在表格中增加该列,同时处理内部记录集
    Call ShowData(rsData)
    mstrTime = Format(dtp.Value, "yyyy-MM-dd HH:mm:ss")
    mblnData = (rsData.RecordCount)
    'Call OutputRsData(mrsSelItems)
    
    mblnInit = True
    If mlng病人数 <> 0 Then Call vsf_AfterRowColChange(2, 2, 1, 1)
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub DeleteColumn(ByVal intCol As Integer)
    Dim lngOrder As Long
    Dim strName As String
    Dim arrColumn
    Dim intStart As Integer, intEnd As Integer
    '删除指定的列
    
    mrsSelItems.Filter = "列=" & intCol
    lngOrder = mrsSelItems!项目序号
    strName = mrsSelItems!项目名称
    mrsSelItems.Filter = 0
    
    '删除列
    vsf.ColPosition(intCol) = vsf.Cols - 1
    vsf.Cols = vsf.Cols - 1
    '处理内部记录集
    With mrsSelItems
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If !列 > intCol Then
                !列 = !列 - 1
                .Update
            ElseIf !列 = intCol Then
                .Delete
            Else
            End If
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
    End With
    '相关模块变量的更新
    If mlngOper > intCol Then mlngOper = mlngOper - 1
    mlngSigner = mlngSigner - 1
    mlngSignTime = mlngSignTime - 1
    mlngRecord = mlngRecord - 1
    mlngGroup = mlngGroup - 1
    mlngCert = mlngCert - 1
    
    arrColumn = Split(mstrSelItems, vbCrLf)
    intEnd = UBound(arrColumn)
    mstrSelItems = ""
    For intStart = 0 To intEnd
        If Val(Split(arrColumn(intStart), "|")(0)) <> lngOrder Then
            mstrSelItems = mstrSelItems & IIf(mstrSelItems = "", "", vbCrLf) & arrColumn(intStart)
        End If
    Next
End Sub

Private Sub InsertColumn(ByVal strSelItems As String)
    Dim lngOrder As Long
    Dim lngRow As Long, lngRows As Long
    
    '如果已存在该列则退出
    mrsSelItems.Filter = "项目序号=" & Val(Split(strSelItems, "|")(0))
    If mrsSelItems.RecordCount <> 0 Then
        mrsSelItems.Filter = 0
        Exit Sub
    End If
    
    '将用户选择的项目添加到表格中
    mrsItems.Filter = "项目序号=" & Val(Split(strSelItems, "|")(0))
    vsf.Cols = vsf.Cols + 1
    vsf.TextMatrix(0, vsf.Cols - 1) = Split(strSelItems, "|")(1) & IIf(NVL(mrsItems!项目单位) = "", "", vbCrLf & "(" & mrsItems!项目单位 & ")")
    vsf.ColAlignment(vsf.Cols - 1) = IIf(mrsItems!项目类型 = 0, flexAlignCenterCenter, flexAlignLeftTop)       '数字则居中显示,非数字以用户录入的数据显示
    mrsItems.Filter = 0
    'Vsf.Cell(flexcpAlignment, 0, Vsf.Cols - 1, Vsf.Rows - 1, Vsf.Cols - 1) = flexAlignCenterCenter  '整列进行设置
        
    '取手术后那列的项目序号
    With mrsSelItems
        .Filter = "列>" & mlngOper
        .Sort = "列"
        Do While Not .EOF
            If !项目序号 > Val(Split(strSelItems, "|")(0)) Then
                lngOrder = !列
                Exit Do
            End If
            .MoveNext
        Loop
        If lngOrder = 0 Then lngOrder = mlngSigner  '没找着,说明没得添加项目,取签名列
        
        .Filter = 0
        If .RecordCount <> 0 Then .MoveFirst
    End With
    
    vsf.ColPosition(vsf.Cols - 1) = lngOrder      '签名人列开始往后移
    '处理内部记录集
    With mrsSelItems
        Do While Not .EOF
            If !列 >= lngOrder Then
                !列 = !列 + 1
                .Update
            End If
            .MoveNext
        Loop
    End With
    
    strFields = "列|项目序号|项目名称|固定"
    strValues = lngOrder & "|" & Split(strSelItems, "|")(0) & "|" & Split(strSelItems, "|")(1) & "|0"
    Call Record_Add(mrsSelItems, strFields, strValues)
    
    '相关模块变量的更新
    mlngSigner = mlngSigner + 1
    mlngSignTime = mlngSignTime + 1
    mlngRecord = mlngRecord + 1
    mlngGroup = mlngGroup + 1
    mlngCert = mlngCert + 1
    
    '根据病人设置此列的背景色
    lngRows = vsf.Rows - 1
    For lngRow = 1 To lngRows
        Call DrawBackColor(lngRow)
    Next
End Sub

Private Sub AddColumns(Optional ByVal rsColumns As ADODB.Recordset)
    Dim blnAdd As Boolean
    '将历史数据中存在的多余列添加到表格中
    If Not rsColumns Is Nothing Then
        If rsColumns.State = 1 Then
            If rsColumns.RecordCount <> 0 Then
                blnAdd = True
            End If
        End If
    End If
    
    If blnAdd Then
        With rsColumns
            Do While Not .EOF
                mrsSelItems.Filter = "项目序号=" & !项目序号
                If mrsSelItems.RecordCount = 0 Then
                    mrsItems.Filter = "项目序号=" & !项目序号
                    vsf.Cols = vsf.Cols + 1
                    vsf.TextMatrix(0, vsf.Cols - 1) = .Fields("项目名称").Value & IIf(NVL(mrsItems!项目单位) = "", "", vbCrLf & "(" & mrsItems!项目单位 & ")")
                    vsf.ColAlignment(vsf.Cols - 1) = IIf(.Fields("项目类型").Value = 0, flexAlignCenterCenter, flexAlignLeftTop)
                    mrsItems.Filter = 0
                    
                    strFields = "列|项目序号|项目名称|固定"
                    strValues = vsf.Cols - 1 & "|" & !项目序号 & "|" & !项目名称 & "|0"
                    Call Record_Add(mrsSelItems, strFields, strValues)
                End If
                .MoveNext
            Loop
        End With
    End If
    
    '固定加入签名人,签名时间列
    With vsf
        .Cols = .Cols + 1
        .TextMatrix(0, .Cols - 1) = "签名人"
        .ColAlignment(vsf.Cols - 1) = flexAlignCenterCenter
        mlngSigner = .Cols - 1
        .Cols = .Cols + 1
        .TextMatrix(0, .Cols - 1) = "签名时间"
        .ColAlignment(vsf.Cols - 1) = flexAlignCenterCenter
        mlngSignTime = .Cols - 1
        .Cols = .Cols + 1
        .TextMatrix(0, .Cols - 1) = "证书ID"
        .ColAlignment(vsf.Cols - 1) = flexAlignCenterCenter
        .ColHidden(.Cols - 1) = True
        mlngCert = .Cols - 1
        .Cols = .Cols + 1
        .TextMatrix(0, .Cols - 1) = "记录ID"
        .ColAlignment(vsf.Cols - 1) = flexAlignCenterCenter
        .ColHidden(.Cols - 1) = True
        mlngRecord = .Cols - 1
        .Cols = .Cols + 1
        .TextMatrix(0, .Cols - 1) = "组号"
        .ColAlignment(vsf.Cols - 1) = flexAlignCenterCenter
        .ColHidden(.Cols - 1) = True
        mlngGroup = .Cols - 1
        .Cols = .Cols + 1
        .TextMatrix(0, .Cols - 1) = "记录人"
        .ColAlignment(vsf.Cols - 1) = flexAlignCenterCenter
        .Cols = .Cols + 1
        .TextMatrix(0, .Cols - 1) = "归档人"
        .ColAlignment(vsf.Cols - 1) = flexAlignCenterCenter
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
    End With
    mrsSelItems.Filter = 0
End Sub

Private Sub ShowData(ByVal rsData As ADODB.Recordset)
    On Error GoTo errHand
    Dim lngRow As Long
    Dim blnNewGroup As Boolean
    Dim lngGroup As Long        '组号
    Dim lngRecord As Long       '记录ID
    Dim strData As String       '暂存数据
    Dim strRecord As String     '记录内存记录集中的数据(对应读取病人信息后)
    Dim str护理等级名称 As String
    Dim lng终止版本 As Long, bln上色 As Boolean
    Dim rsTemp As New ADODB.Recordset   '提取当前记录最大的终止版本
    
    strFields = "行|病人ID|主页ID|科室ID|婴儿|护理等级|护理等级名称|匹配列"
    
    With rsData
        Do While Not .EOF
            '记录ID不同,就说明是不同的病人了
            If (lngRecord <> !记录ID Or lngGroup <> !记录组号) Then
                '提取当前记录最大的终止版本
                gstrSQL = " Select max(开始版本),Max(终止版本) From 病人护理内容 Where 记录ID=[1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取当前记录最大的终止版本", CLng(!记录ID))
                lng终止版本 = NVL(rsTemp.Fields(0).Value, 1)
                If lng终止版本 < NVL(rsTemp.Fields(1).Value, 1) Then lng终止版本 = NVL(rsTemp.Fields(1).Value, 1)
                
                blnNewGroup = False
                If lngRecord <> !记录ID Then
                    '定位病人所在行
                    mrsPatient.Filter = "病人ID=" & !病人ID & " And 主页ID=" & !主页ID & " And 科室ID=" & !科室ID & " And 婴儿=" & !婴儿
                    str护理等级名称 = "三级护理"
                    If mrsPatient.RecordCount <> 0 Then
                        mrsPatient.MoveLast
                        lngRow = mrsPatient!行
                        strRecord = mrsPatient!病人ID & "|" & mrsPatient!主页ID & "|" & mrsPatient!科室ID & "|" & mrsPatient!婴儿 & "|" & mrsPatient!护理等级 & "|" & mrsPatient!护理等级名称 & "|" & NVL(mrsPatient!匹配列)
                        str护理等级名称 = mrsPatient!护理等级名称
                    End If
                    mrsPatient.Filter = 0
                Else
                    '增加新行(分组数据)
                    blnNewGroup = True
                    lngRow = lngRow + 1
                    vsf.Rows = vsf.Rows + 1
                    vsf.RowPosition(vsf.Rows - 1) = lngRow
                    '更新内存记录集
                    With mrsPatient
                        .Filter = "行>=" & lngRow
                        Do While Not .EOF
                            !行 = !行 + 1
                            .Update
                            .MoveNext
                        Loop
                        .Filter = 0
                    End With
                    '添加当前行
                    strValues = lngRow & "|" & strRecord
                    Call Record_Add(mrsPatient, strFields, strValues)
                    
                    Call DrawBackColor(lngRow)
                End If
                
                '先写入签名人及签名时间
                lngRecord = !记录ID
                lngGroup = !记录组号
                bln上色 = True
                If Not IsNull(!签名人) Then
                    bln上色 = False
                    vsf.Cell(flexcpPicture, lngRow, 0) = imgRow.ListImages(1).Picture
                End If
                vsf.Cell(flexcpPictureAlignment, lngRow, 0) = flexAlignCenterCenter
                
'                If Not blnNewGroup Then
                    vsf.TextMatrix(lngRow, 姓名) = IIf(!婴儿 <> 0, Space(4), "") & CStr(.Fields("姓名").Value)
                    vsf.TextMatrix(lngRow, 性别) = CStr(.Fields("性别").Value)
                    If NVL(.Fields("住院号").Value, 0) = 0 Then
                        vsf.TextMatrix(lngRow, 住院号) = ""
                    Else
                        vsf.TextMatrix(lngRow, 住院号) = CLng(NVL(.Fields("住院号").Value, 0))
                    End If
                    vsf.TextMatrix(lngRow, 床号) = NVL(.Fields("床号").Value)
                    vsf.TextMatrix(lngRow, 护理等级) = str护理等级名称
'                End If
                vsf.TextMatrix(lngRow, 病人ID) = Val(.Fields("病人ID").Value)
                vsf.TextMatrix(lngRow, 主页ID) = Val(.Fields("主页ID").Value)
                vsf.TextMatrix(lngRow, mlngCert) = Val(NVL(.Fields("证书ID").Value, 0))
                vsf.TextMatrix(lngRow, mlngSigner) = NVL(.Fields("签名人").Value)
                vsf.TextMatrix(lngRow, mlngSignTime) = Format(.Fields("签名时间").Value, "yyyy-MM-dd HH:mm:ss")
                vsf.TextMatrix(lngRow, mlngRecord) = CLng(.Fields("记录ID").Value)
                vsf.TextMatrix(lngRow, mlngGroup) = CLng(.Fields("记录组号").Value)
                vsf.TextMatrix(lngRow, vsf.Cols - 2) = NVL(.Fields("记录人").Value)
                vsf.TextMatrix(lngRow, vsf.Cols - 1) = NVL(.Fields("归档人").Value)
                vsf.RowData(lngRow) = 0
                
                If bln上色 Then '签名人为空,且终止版本大于1,才说明需要上色;排开初步产生的数据不需要上色的情况
                    bln上色 = (lng终止版本 > 1)
                End If
            End If
            
            '先写入普通的护理项目
            If !项目序号 <> 0 Then
                '如果未记说明不为空,显示未记说明
                If Not IsNull(.Fields("未记说明").Value) Then
                    strData = .Fields("未记说明").Value
                Else
                    strData = NVL(.Fields("记录结果").Value)
                    If Not IsNull(.Fields("标记").Value) Then
                        strData = strData & "/" & .Fields("标记").Value
                    End If
                    If Not IsNull(.Fields("部位").Value) Then
                        strData = .Fields("部位").Value & ":" & strData
                    ElseIf !项目序号 = 1 Then
                        strData = "腋温:" & strData
                    End If
                End If
                
                mrsSelItems.Filter = "项目序号=" & !项目序号
                If mrsSelItems.RecordCount <> 0 Then
                    If !项目序号 = 5 Then   '收缩压,如果对应单元格有内容,则说明已填入舒张压,以/组合显示
                        If vsf.TextMatrix(lngRow, mrsSelItems!列) <> "" Then
                            vsf.TextMatrix(lngRow, mrsSelItems!列) = vsf.TextMatrix(lngRow, mrsSelItems!列) & "/" & strData
                        Else
                            vsf.TextMatrix(lngRow, mrsSelItems!列) = strData
                        End If
                    Else
                        vsf.TextMatrix(lngRow, mrsSelItems!列) = strData
                    End If
                End If
            Else
                '再写入手术
                strData = NVL(.Fields("记录结果").Value)
                mrsSelItems.Filter = "项目序号=0"
                If mrsSelItems.RecordCount <> 0 Then
                    vsf.TextMatrix(lngRow, mrsSelItems!列) = strData
                End If
            End If
            
            '上色(手术除外)
            If !实际版本 = lng终止版本 And bln上色 Then
                vsf.Cell(flexcpForeColor, lngRow, mrsSelItems!列) = &HFF&
            End If
            
            .MoveNext
        Loop
    End With
    mrsSelItems.Filter = 0
    
    '使用CellData来保存修改标志
    vsf.Cell(flexcpAlignment, 1, 1, vsf.Rows - 1, 有效数据项 - 1) = flexAlignLeftCenter
    'Vsf.Cell(flexcpAlignment, 1, 有效数据项, Vsf.Rows - 1, Vsf.Cols - 1) = flexAlignCenterCenter
    vsf.Cell(flexcpData, 1, 1, vsf.Rows - 1, vsf.Cols - 1) = 0
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    mrsSelItems.Filter = 0
End Sub

Private Sub PasteData()
    Dim arrOrder, arrData
    Dim intCol As Integer, intCols As Integer
    '根据项目序号进行复制
    
    arrOrder = Split(mstrSel, "|")
    arrData = Split(arrOrder(1), ",")
    arrOrder = Split(arrOrder(0), ",")
    intCols = UBound(arrOrder)
    
    mrsPatient.Filter = "行=" & vsf.ROW
    If mrsPatient.RecordCount <> 0 Then
        For intCol = 0 To intCols
            '首先检查该列是否匹配
            If InStr(1, "," & mrsPatient!匹配列 & ",", "," & arrOrder(intCol) & ",") <> 0 Then
                '更新数据
                mrsSelItems.Filter = "项目序号=" & arrOrder(intCol)
                If mrsSelItems.RecordCount <> 0 Then
                    If arrData(intCol) <> "" Or vsf.TextMatrix(vsf.ROW, mrsSelItems!列) <> "" Then
                        vsf.TextMatrix(vsf.ROW, mrsSelItems!列) = arrData(intCol)
                        
                        '更新修改标志
                        vsf.RowData(vsf.ROW) = 1
                        vsf.Cell(flexcpData, vsf.ROW, mrsSelItems!列) = 1
                        mblnChange = True   '有可能复制项目的护理等级不适用病人当前护理等级，因此在循环内设置
                    End If
                End If
            End If
        Next
    End If
    
    mrsPatient.Filter = 0
    mrsSelItems.Filter = 0
End Sub

Private Function WriteIntoVsf(Optional ByVal strInfo As String) As Boolean
    Dim blnAllow As Boolean
    Dim StrText As String
    Dim lngRecord As Long
    Dim lngRow As Long, lngCol As Long
    Dim lngRows As Long, lngCols As Long
    Dim intType As Integer, lngOrder As Long, lngClass As Long, strName As String, lngLength As Long, str值域 As String
    
    lngRow = Split(txt数据.Tag, "|")(0)
    lngCol = Split(txt数据.Tag, "|")(1)
    
    If picInput.Visible Then
        If lngCol = 姓名 Then
            '录入病人信息
            If AddPatient(txt数据.Text) = False Then
                '还原
                vsf.TextMatrix(lngRow, lngCol) = picInput.Tag
                txt数据.Tag = ""
                picInput.Visible = False
            End If
        ElseIf txt数据.Enabled Then
            '检查数据合法性
             '检查数据合法性
            If Val(cbo部位.Tag) = 0 Then
                If txt数据.Text <> "" Then
                    StrText = IIf(cbo部位.Visible And Trim(cbo部位.Text) <> "", cbo部位.Text & ":", "") & Trim(txt数据.Text)
                End If
            Else
                StrText = IIf(Trim(txt数据.Text) <> "", Trim(txt数据.Text), cbo部位.Text)
            End If
    
            '定位列对应的护理记录进行检查
            mrsSelItems.Filter = "列=" & lngCol
            mrsItems.Filter = "项目序号=" & mrsSelItems!项目序号
            
            intType = mrsItems!项目类型     '0-数值；1-文字
            lngClass = mrsItems!项目性质
            lngOrder = mrsItems!项目序号
            strName = mrsItems!项目名称
            lngLength = mrsItems!项目长度 + IIf(NVL(mrsItems!项目小数, 0) = 0, 0, NVL(mrsItems!项目小数, 0) + 1)
            If intType = 0 Then
                str值域 = NVL(mrsItems!项目值域)
            Else
                str值域 = ""
                StrText = txt数据.Text      '非数字型项目,以用户原始录入为准
            End If
            
            blnAllow = CheckValid(StrText, lngOrder, lngClass, strName, lngLength, lngRow, lngCol, str值域)
            mrsItems.Filter = 0
            mrsSelItems.Filter = 0
            If blnAllow Then vsf.TextMatrix(lngRow, lngCol) = StrText
        Else
            blnAllow = True
            vsf.TextMatrix(lngRow, lngCol) = txt数据.Text
        End If
    Else
        blnAllow = True
        lngRow = Split(lvwMultiSel.Tag, "|")(0)
        lngCol = Split(lvwMultiSel.Tag, "|")(1)
        vsf.TextMatrix(lngRow, lngCol) = strInfo
    End If

    txt数据.Tag = ""
    cbo部位.Visible = False
    txt数据.Height = picInput.Height
    picInput.Visible = False
    lvwMultiSel.Visible = False
    
    '更新修改标志
    If blnAllow Then
        If picInput.Tag <> vsf.TextMatrix(lngRow, lngCol) Then
            '如果是修改的时间,需要把记录ID相同的所有记录的时间全部修改了
            If lngCol <= 2 And Val(vsf.TextMatrix(lngRow, mlngRecord)) <> 0 Then
                lngRows = vsf.Rows - 1
                lngRecord = Val(vsf.TextMatrix(lngRow, mlngRecord))
                For lngRow = 1 To lngRows
                    If Val(vsf.TextMatrix(lngRow, mlngRecord)) = lngRecord Then
                        vsf.TextMatrix(lngRow, lngCol) = StrText
                        '修改标志
                        vsf.RowData(lngRow) = 1
                        vsf.Cell(flexcpData, lngRow, lngCol) = 1
                    End If
                Next
            Else
                '修改标志
                lngRecord = Val(vsf.TextMatrix(lngRow, mlngRecord))
                vsf.RowData(lngRow) = 1
                vsf.Cell(flexcpData, lngRow, lngCol) = 1
            End If
            mblnChange = True
        End If

        WriteIntoVsf = True
        If mblnChange Then RaiseEvent AfterDataChanged
    End If
End Function

Private Sub lst体温标识_DblClick()
    Dim lng病人ID As Long, lng主页ID As Long, int婴儿 As Integer
    On Error GoTo errHand
    '保存病人体温标识
    mrsPatient.Filter = "行=" & vsf.ROW
    If mrsPatient.RecordCount <> 0 Then
        '先定位修改过的行,再在列中循环找到修改过的列
        lng病人ID = mrsPatient!病人ID
        lng主页ID = mrsPatient!主页ID
        int婴儿 = mrsPatient!婴儿
        
        gstrSQL = "ZL_病人体温标识_Update(" & lng病人ID & "," & lng主页ID & "," & _
            int婴儿 & ",'" & lst体温标识.Text & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存病人体温标识")
        vsf.TextMatrix(vsf.ROW, 体温标识) = lst体温标识.Text
        
        Call vsf_KeyDown(vbKeyReturn, 0)
    End If
    mrsPatient.Filter = 0
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    mrsPatient.Filter = 0
End Sub

Private Sub lst体温标识_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call lst体温标识_DblClick
End Sub

Private Sub lvwMultiSel_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strData As String
    Dim intCol As Integer, intMax As Integer
    Dim blnAllow As Boolean
    
    If KeyCode = vbKeyReturn Then
        intMax = lvwMultiSel.ListItems.Count
        For intCol = 1 To intMax
            If lvwMultiSel.ListItems(intCol).Checked Then
                strData = strData & IIf(strData = "", "", ",") & lvwMultiSel.ListItems(intCol).Text
            End If
        Next
        blnAllow = WriteIntoVsf(strData)
        Call vsf_KeyDown(vbKeyReturn, Shift)
'    ElseIf KeyCode = vbKeyLeft Then
'        Call vsf_KeyDown(KeyCode, Shift)
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mblnChange Then
        If MsgBox("当前数据还未保存，点“是”进行保存，点“否”将放弃本次修改！", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Call Vsf_EnterCell
            Call SaveData
        End If
    End If
End Sub

Private Sub lvw房间号_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Static blnExit As Boolean
    Dim blnFind As Boolean
    Dim blnCheck As Boolean
    Dim intDo As Integer, intMax As Integer
    
    If blnExit Then Exit Sub
    intMax = lvw房间号.ListItems.Count
    blnCheck = Item.Checked
    
    If Item.Text = "所有房间" Then
        blnExit = True
        For intDo = 2 To intMax
            lvw房间号.ListItems(intDo).Checked = blnCheck
        Next
    Else
        blnExit = True
        For intDo = 2 To intMax
            If lvw房间号.ListItems(intDo).Checked = Not blnCheck Then
                blnFind = True
                Exit For
            End If
        Next
        If blnFind Then
            lvw房间号.ListItems(1).Checked = False
        Else
            lvw房间号.ListItems(1).Checked = blnCheck
        End If
    End If
    
    blnExit = False
End Sub

Private Sub lvw房间号_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim blnCancel As Boolean
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    Call lvw房间号_Validate(blnCancel)
    cmd刷新.SetFocus
End Sub

Private Sub lvw房间号_Validate(Cancel As Boolean)
    Dim str清单 As String
    Dim intDo As Integer, intMax As Integer
    
    If lvw房间号.ListItems(1).Checked Then
        lvw房间号.Visible = False
        Me.lbl房间号清单.Caption = "所有房间"
        Exit Sub
    End If
    
    intMax = lvw房间号.ListItems.Count
    For intDo = 2 To intMax
        If lvw房间号.ListItems(intDo).Checked Then str清单 = str清单 & IIf(str清单 = "", "", ",") & lvw房间号.ListItems(intDo).Text
    Next
    
    If str清单 = "" Then
        Cancel = True
        MsgBox "必须选择一个房间！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    lvw房间号.Visible = False
    lbl房间号清单.Caption = str清单
    cmd刷新.SetFocus
End Sub

Private Sub mfrmCaseTendEditForSinglePerson_DBCLICK(ByVal strData As String)
    Dim StrText As String
    
    If vsf.TextMatrix(vsf.ROW, vsf.Cols - 1) <> "" Then Exit Sub
    
    mrsSelItems.Filter = "项目序号=" & Val(Split(strData, "|")(0))
    If mrsSelItems.RecordCount <> 0 Then
        StrText = vsf.TextMatrix(vsf.ROW, mrsSelItems!列)
        
        '更新修改标志
        If StrText <> Split(strData, "|")(1) Then
            '修改标志
            vsf.RowData(vsf.ROW) = 1
            vsf.Cell(flexcpData, vsf.ROW, mrsSelItems!列) = 1
            vsf.TextMatrix(vsf.ROW, mrsSelItems!列) = Split(strData, "|")(1)
            
            mblnChange = True
        End If
        
'        '移动到下一行
'        Call vsf_KeyDown(vbKeyReturn, 0)
    End If
    mrsSelItems.Filter = 0
End Sub


Private Sub optLevel_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
     If optLevel(0).Value Then
         SaveSetting "ZLSOFT", "私有模块\frmCaseTendEditForBatch\" & gstrUserName, "Value", 0
    Else
        SaveSetting "ZLSOFT", "私有模块\frmCaseTendEditForBatch\" & gstrUserName, "Value", 1
    End If
End Sub

Private Sub picMain_Resize()
    picMain.Left = 0
    vsf.Width = picMain.Width
    vsf.Height = picMain.Height - vsf.Top
End Sub

Private Sub cbo部位_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call txt数据_KeyDown(vbKeyReturn, 0): Exit Sub
End Sub

Private Sub picQuery_Resize()
    mfrmCaseTendEditForSinglePerson.Left = 0
    mfrmCaseTendEditForSinglePerson.Top = 0
    mfrmCaseTendEditForSinglePerson.Width = picQuery.Width
    mfrmCaseTendEditForSinglePerson.Height = picQuery.Height
End Sub

Private Sub txt床号_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngRow As Long
    '定位到当前床号上
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(txt床号.Text) = "" Then Exit Sub
    
    lngRow = vsf.FindRow(Trim(txt床号.Text), , 床号)
    If lngRow < 1 Then Exit Sub
    vsf.ROW = lngRow
    If vsf.RowIsVisible(lngRow) = False Then vsf.TopRow = lngRow
End Sub

Private Sub txt数据_GotFocus()
    Call zlControl.TxtSelAll(txt数据)
End Sub

Private Sub txt数据_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim StrText As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    If KeyCode = vbKeyDown And InStr(1, "体温脉搏呼吸手术", Mid(vsf.TextMatrix(0, vsf.Col), 1, 2)) <> 0 Then
        If Shift = 0 Then
            cbo部位.Tag = 0
            cbo部位.Clear
            If Mid(vsf.TextMatrix(0, vsf.Col), 1, 2) = "体温" Then
                cbo部位.AddItem "腋温"
                cbo部位.AddItem "口温"
                cbo部位.AddItem "肛温"
            ElseIf Mid(vsf.TextMatrix(0, vsf.Col), 1, 2) = "脉搏" Then
                cbo部位.AddItem ""
                cbo部位.AddItem "起搏器"
            ElseIf Mid(vsf.TextMatrix(0, vsf.Col), 1, 2) = "呼吸" Then
                cbo部位.AddItem "自主呼吸"
                cbo部位.AddItem "呼吸机"
            ElseIf Mid(vsf.TextMatrix(0, vsf.Col), 1, 2) = "手术" Then
                cbo部位.AddItem "手术"
                cbo部位.AddItem "分娩"
                cbo部位.AddItem "手术分娩"
            End If
            If cbo部位.ListCount <> 0 Then cbo部位.ListIndex = 0
            cmd未记说明.ToolTipText = IIf(Val(cbo部位.Tag) = 0, "切换到未记说明", "切换到部位")
        ElseIf Shift = vbShiftMask Then
            gstrSQL = " Select 名称 From 常用体温说明 Order by 编码"
            Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "提取未记说明")
            With rsTemp
                Me.cbo部位.Clear
                Do While Not .EOF
                    Me.cbo部位.AddItem !名称
                    .MoveNext
                Loop
                Me.cbo部位.ListIndex = 0
                cbo部位.Tag = 1
            End With
        End If
        
        With cbo部位
            .Top = picInput.Height - .Height
            .Width = picInput.Width
            .Visible = True
            .ZOrder 0
            .SetFocus
        End With
        txt数据.Height = picInput.Height - cbo部位.Height
        If cbo部位.Tag = 1 Then txt数据.Text = cbo部位.Text
        cmd未记说明.ToolTipText = IIf(Val(cbo部位.Tag) = 0, "切换到未记说明", "切换到部位")
    ElseIf KeyCode = vbKeyReturn Then
        Dim strData As String
        Dim lngCol As Long
        Dim blnAllow As Boolean
        
        blnAllow = True
        If Shift = vbCtrlMask Then Exit Sub
        If picInput.Visible And txt数据.Tag <> "" Then
            lngCol = Split(txt数据.Tag, "|")(1)
            If InStr(1, "体温脉搏呼吸", Mid(vsf.TextMatrix(0, lngCol), 1, 2)) <> 0 Then
                '检查数据合法性
                If cbo部位.Tag = 0 Then
                    If txt数据.Text <> "" Then
                        strData = IIf(cbo部位.Visible And Trim(cbo部位.Text) <> "", cbo部位.Text & ":", "") & Trim(txt数据.Text)
                    End If
                Else
                    strData = IIf(Trim(txt数据.Text) <> "", Trim(txt数据.Text), cbo部位.Text)
                End If
            Else
                strData = Trim(txt数据.Text)
            End If
            If strData <> picInput.Tag Then blnAllow = WriteIntoVsf
        End If
        
        If blnAllow Then
            Call vsf_KeyDown(vbKeyReturn, Shift)
        Else
            Call Vsf_EnterCell
        End If
    ElseIf KeyCode = vbKeyLeft Then
        If txt数据.SelStart = 0 Then Call vsf_KeyDown(KeyCode, Shift)
    ElseIf KeyCode = vbKeyW And Shift = vbCtrlMask Then
        Dim lng病人ID As Long, lng主页ID As Long
        If Not (cmd未记说明.Visible And cbo部位.Visible = False) Then Exit Sub
        
        mrsPatient.Filter = "行=" & vsf.ROW
        If mrsPatient.RecordCount <> 0 Then
            lng病人ID = mrsPatient!病人ID
            lng主页ID = mrsPatient!主页ID
        End If
        mrsPatient.Filter = 0
        
        If lng病人ID = 0 Then Exit Sub
        StrText = frmWordsEditor.ShowMe(Me, lng病人ID, lng主页ID, txt数据.Text)
        If StrText = "" Then Exit Sub
        txt数据.Text = StrText
        
        Call txt数据_KeyDown(vbKeyReturn, 0)
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim blnDo As Boolean
    Dim lng证书ID As Long, str签名人 As String, str签名时间 As String
    Dim lng病人ID As Long, lng主页ID As Long, int婴儿 As Integer
    '显示指定病人的历史数据
    
    If mblnInit = False Then Exit Sub
    '显示当前项目的相关信息
    mrsSelItems.Filter = "列=" & NewCol
    If mrsSelItems.RecordCount <> 0 Then
        mrsItems.Filter = "项目序号=" & mrsSelItems!项目序号
        If mrsItems.RecordCount <> 0 Then
            If NVL(mrsItems!项目值域) <> "" Then
                If mrsItems!项目类型 = 0 Then
                    stbThis.Panels(2).Text = "有效范围:" & Split(mrsItems!项目值域, ";")(0) & "～" & Split(mrsItems!项目值域, ";")(1)
                Else
                    stbThis.Panels(2).Text = "有效范围:" & mrsItems!项目值域
                End If
            Else
                stbThis.Panels(2).Text = ""
            End If
            
            If mrsSelItems!项目序号 = 1 Then
                stbThis.Panels(2).Text = stbThis.Panels(2).Text & Space(5) & "物理降温表示法:39/37.5"
            ElseIf mrsSelItems!项目序号 = 3 Then
                If mbln心率 = False Then stbThis.Panels(2).Text = stbThis.Panels(2).Text & Space(5) & "脉搏短拙表示法:130/120"
            ElseIf vsf.TextMatrix(0, NewCol) Like "血压*" Then
                stbThis.Panels(2).Text = stbThis.Panels(2).Text & Space(5) & "录入规则:收缩压/舒张压"
            End If
            
            If mrsSelItems!项目序号 >= 1 And mrsSelItems!项目序号 <= 3 Then
                stbThis.Panels(2).Text = stbThis.Panels(2).Text & Space(5) & "按↓进行部位选择;按SHIFT+↓进行未记说明的选择"
            End If
        End If
    End If
    mrsSelItems.Filter = 0
    mrsItems.Filter = 0
    
    '----------------------------------------
    '如果选择行发生变化则提取该病人的历史数据
    If OldRow = NewRow Then Exit Sub
    '如果病人没有发生变化,也不做任何处理
    mrsPatient.Filter = "行=" & OldRow
    If mrsPatient.RecordCount <> 0 Then
        lng病人ID = mrsPatient!病人ID
        lng主页ID = mrsPatient!主页ID
        int婴儿 = mrsPatient!婴儿
    End If
    
    mrsPatient.Filter = "行=" & NewRow
    If mrsPatient.RecordCount <> 0 Then
        If lng病人ID <> mrsPatient!病人ID Or lng主页ID <> mrsPatient!主页ID Or int婴儿 <> mrsPatient!婴儿 Then
            blnDo = True
        End If
    End If
    
    '告诉查询子窗体更新数据
    If blnDo Then
        Call mfrmCaseTendEditForSinglePerson.ShowMe(Me, mrsPatient!病人ID, mrsPatient!主页ID, mrsPatient!科室ID, mrsPatient!婴儿, mrsPatient!护理等级, "", False, False)
        
        '根据归档与否,实时更新归档人
        mrsPatient.Filter = "病人ID=" & mrsPatient!病人ID & " And 主页ID=" & mrsPatient!主页ID & " And 婴儿=" & mrsPatient!婴儿
        Do While Not mrsPatient.EOF
            vsf.TextMatrix(mrsPatient!行, vsf.Cols - 1) = mfrmCaseTendEditForSinglePerson.mstrPigeonhole
            mrsPatient.MoveNext
        Loop
    End If
    
    mrsPatient.Filter = 0
End Sub

Private Sub vsf_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    blnScroll = True
    Call Vsf_EnterCell
    blnScroll = False
End Sub

Private Sub vsf_AfterUserResize(ByVal ROW As Long, ByVal Col As Long)
    Call Vsf_EnterCell
End Sub

Private Sub vsf_DblClick()
    mblnShow = True
    Call Vsf_EnterCell
End Sub

Private Sub Vsf_EnterCell()
    Dim arrData
    Dim strData As String
    Dim intIndex As Integer
    Dim lngRow As Long
    Dim lngCol As Long
    Dim blnAllow As Boolean, blnWords As Boolean
    Dim intCol As Integer, intMax As Integer
    
    If mblnInit = False Then Exit Sub
    
    '如果已录入数据则保存
    blnAllow = True
    If picInput.Visible And txt数据.Tag <> "" Then
        lngRow = Split(txt数据.Tag, "|")(0)
        lngCol = Split(txt数据.Tag, "|")(1)
        If InStr(1, "体温脉搏呼吸", Mid(vsf.TextMatrix(0, lngCol), 1, 2)) <> 0 Then
            '检查数据合法性
            If cbo部位.Tag = 0 Then
                If txt数据.Text <> "" Then
                    strData = IIf(cbo部位.Visible And Trim(cbo部位.Text) <> "", cbo部位.Text & ":", "") & txt数据.Text
                End If
            Else
                strData = IIf(Trim(txt数据.Text) <> "", Trim(txt数据.Text), cbo部位.Text)
            End If
        Else
            strData = txt数据.Text
        End If
        If strData <> picInput.Tag Then blnAllow = WriteIntoVsf
    ElseIf lvwMultiSel.Visible Then
        intMax = lvwMultiSel.ListItems.Count
        For intCol = 1 To intMax
            If lvwMultiSel.ListItems(intCol).Checked Then
                strData = strData & IIf(strData = "", "", ",") & lvwMultiSel.ListItems(intCol).Text
            End If
        Next
        blnAllow = WriteIntoVsf(strData)
    ElseIf vsf.Col = 体温标识 Then
        blnAllow = True
    End If
    Call vsf.AutoSize(0, vsf.Cols - 1)
    picInput.Visible = False
    lvwMultiSel.Visible = False
    lst体温标识.Visible = False
    If blnAllow = False Then
        If vsf.ROW <> lngRow Then vsf.ROW = lngRow
        If vsf.Col <> lngCol Then vsf.Col = lngCol
        Exit Sub
    End If
    
    RaiseEvent AfterSelChange(IIf(Trim(vsf.TextMatrix(vsf.ROW, mlngSigner)) <> "", 1, 0))
    
    mblnCheckVersion = CheckVersion
    If mblnShow = False Or (vsf.TextMatrix(vsf.ROW, vsf.Cols - 1) <> "") Then Exit Sub
    If vsf.Col = 0 Or vsf.ROW = 0 Then Exit Sub
    If vsf.Col = mlngOper And mblnCheckVersion = False Then Exit Sub
    If vsf.Col <> 体温标识 Then
        If vsf.Col >= mlngSigner Or (vsf.Col < 有效数据项 And (vsf.Col <> 姓名 And Val(vsf.TextMatrix(vsf.ROW, mlngRecord)) = 0)) Then Exit Sub   '签名人,签名时间以及组号不允许编辑,组号隐藏
    End If
    If vsf.RowIsVisible(vsf.ROW) = False Then Exit Sub
    '检查此列是否允许编辑(存在该病人,找到当前列的项目序号进行检查)
    blnAllow = False
    If vsf.Col <> 1 Then
        If vsf.Col = 体温标识 Then
            With lst体温标识
                .Left = vsf.CellLeft
                .Top = vsf.CellTop
                .Width = vsf.CellWidth
                .Visible = True
                .ZOrder 0
            End With
        Else
            mrsPatient.Filter = "行=" & vsf.ROW
            If mrsPatient.RecordCount <> 0 Then
                If vsf.Col = 姓名 Or vsf.Col = mlngOper Then
                    blnAllow = True
                Else
                    mrsSelItems.Filter = "列=" & vsf.Col
                    If mrsSelItems.RecordCount <> 0 Then
                        If InStr(1, "," & NVL(mrsPatient!匹配列) & ",", "," & mrsSelItems!项目序号 & ",") <> 0 Then
                            blnAllow = True
                        End If
                    End If
                End If
            End If
            mrsPatient.Filter = 0
            mrsSelItems.Filter = 0
        End If
    Else
        '新增的空白行,只允许编辑病人
        blnAllow = True
    End If
    If Not blnAllow Then Exit Sub
    If Not blnScroll And vsf.Visible And vsf.Enabled Then vsf.SetFocus
    
    '准备显示
    With picInput
        .Tag = vsf.TextMatrix(vsf.ROW, vsf.Col)             '保存编辑前的数据
        .Left = vsf.ColPos(vsf.Col) + vsf.Left
        .Top = vsf.RowPos(vsf.ROW) + vsf.Top
        .Width = vsf.ColWidth(vsf.Col)
        If vsf.ROW = vsf.Rows - 1 Then
            .Height = vsf.ROWHEIGHT(vsf.ROW)    '取其行高
        Else
            .Height = vsf.RowPos(vsf.ROW + 1) - vsf.RowPos(vsf.ROW)
        End If
        If .Height > vsf.RowHeightMax Then .Height = vsf.RowHeightMax
        If .Height < 600 Then .Height = 600
        .ZOrder 0
        .Visible = True
    End With
    With cbo部位
        .Visible = False
        .Clear
        .Tag = 0
        blnAllow = True
        If Mid(vsf.TextMatrix(0, vsf.Col), 1, 2) = "体温" Then
            .AddItem "腋温"
            .AddItem "口温"
            .AddItem "肛温"
            .Visible = True
        ElseIf Mid(vsf.TextMatrix(0, vsf.Col), 1, 2) = "脉搏" Then
            .AddItem ""
            .AddItem "起搏器"
            .Visible = True
        ElseIf Mid(vsf.TextMatrix(0, vsf.Col), 1, 2) = "呼吸" Then
            .AddItem "自主呼吸"
            .AddItem "呼吸机"
            .Visible = True
        ElseIf Mid(vsf.TextMatrix(0, vsf.Col), 1, 2) = "手术" Then
            .AddItem "手术"
            .AddItem "分娩"
            .AddItem "手术分娩"
            .Visible = True
            blnAllow = False
        Else
            '定位列,如果是单选,则将值域加入下拉框
            mrsSelItems.Filter = "列=" & vsf.Col
            If mrsSelItems.RecordCount <> 0 Then
                mrsItems.Filter = "项目序号=" & mrsSelItems!项目序号
                If mrsItems.RecordCount <> 0 Then
                    If mrsItems!项目表示 = 2 Then
                        '单选
                        .AddItem " "
                        arrData = Split(NVL(mrsItems!项目值域), ";")
                        intMax = UBound(arrData)
                        For intCol = 0 To intMax
                            If Mid(arrData(intCol), 1, 1) = "√" Then intIndex = intCol
                            .AddItem Replace(arrData(intCol), "√", "")
                        Next
                        blnAllow = False
                    ElseIf mrsItems!项目表示 = 3 Then
                        '多选
                        picInput.Visible = False
                        lvwMultiSel.Left = picInput.Left + picInput.Width - lvwMultiSel.Width
                        lvwMultiSel.Top = picInput.Top + picInput.Height
                        lvwMultiSel.Visible = True
                        If lvwMultiSel.Top + lvwMultiSel.Height > picMain.Height Then lvwMultiSel.Top = picInput.Top - lvwMultiSel.Height
                        
                        '加入数据
                        lvwMultiSel.ListItems.Clear
                        arrData = Split(NVL(mrsItems!项目值域), ";")
                        intMax = UBound(arrData)
                        For intCol = 0 To intMax
                            strData = Replace(arrData(intCol), "√", "")
                            lvwMultiSel.ListItems.Add , "K" & intCol, strData
                            If Mid(arrData(intCol), 1, 1) = "√" Then lvwMultiSel.ListItems(intCol + 1).Selected = True
                            If InStr(1, "," & vsf.TextMatrix(vsf.ROW, vsf.Col) & ",", "," & strData & ",") <> 0 Then lvwMultiSel.ListItems(intCol + 1).Checked = True
                        Next
                        lvwMultiSel.Tag = vsf.ROW & "|" & vsf.Col
                        lvwMultiSel.SetFocus
                    ElseIf mrsItems!项目类型 = 1 And mrsItems!项目长度 >= 200 Then
                        blnWords = True
                    End If
                End If
            End If
            mrsSelItems.Filter = 0
            mrsItems.Filter = 0
        End If
        If .ListCount <> 0 Then .ListIndex = 0
    End With
    
    With txt数据
        .Enabled = blnAllow          '如果当前列是手术列或单选项,则不允许录入
        .Text = vsf.TextMatrix(vsf.ROW, vsf.Col)
        If .Enabled Then
            If InStr(1, .Text, ":") <> 0 And cbo部位.ListCount > 0 Then
                With cbo部位
                    If InStr(1, txt数据.Text, ":") <> 0 Then
                        Call zlControl.CboLocate(cbo部位, Split(txt数据.Text, ":")(0))
                    End If
                    '.Top = picInput.Height - .Height
                    .Width = picInput.Width
                    .Visible = True
                    .ZOrder 0
                End With
                .Text = Split(.Text, ":")(1)
            End If
        Else
            If .Text <> "" Then Call zlControl.CboLocate(cbo部位, .Text)
            With cbo部位
                '.Top = picInput.Height - .Height
                .Width = picInput.Width
                .Visible = True
                .ZOrder 0
            End With
        End If
        .Width = picInput.Width
        .Height = picInput.Height - IIf(cbo部位.Visible, cbo部位.Height, 0)
        .Tag = vsf.ROW & "|" & vsf.Col
    End With
    
    If cbo部位.Enabled Then
        cbo部位.Top = picInput.Height - cbo部位.Height
        cbo部位.Width = txt数据.Width
    End If
    
    cmd未记说明.Visible = (InStr(1, "体温脉搏呼吸", Mid(vsf.TextMatrix(0, vsf.Col), 1, 2)) <> 0) Or blnWords
    If cmd未记说明.Visible Then
        '如果是体温曲线项目,如果录入的数据不是数值型,则将标志改为1
        If InStr(1, txt数据.Text, "/") = 0 Then
            If Trim(Split(txt数据.Text & "|", "|")(0)) <> "" And Trim(Split(txt数据.Text & "|", "|")(0)) <> "不升" Then
                If Not IsNumeric(Split(txt数据.Text & "|", "|")(0)) Then
                    strData = Split(txt数据.Text & "|", "|")(0)
                    Call txt数据_KeyDown(vbKeyDown, vbShiftMask)
                    txt数据.Text = strData
                End If
            End If
        End If
        If blnWords Then
            cmd未记说明.ToolTipText = "可以按Ctrl+W调出词句编辑器"
        Else
            cmd未记说明.ToolTipText = IIf(Val(cbo部位.Tag) = 0, "切换到未记说明", "切换到部位")
        End If
        cmd未记说明.Left = txt数据.Width - cmd未记说明.Width
    End If
    
    On Error Resume Next
    If txt数据.Enabled Then
        txt数据.SetFocus
    Else
        cbo部位.SetFocus
    End If
End Sub

Private Sub SetDockRight(BarToDock As CommandBar, BarOnLeft As CommandBar)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Dim Left As Long
    Dim Top As Long
    Dim Right As Long
    Dim Bottom As Long
    
    cbsThis.RecalcLayout
    BarOnLeft.GetWindowRect Left, Top, Right, Bottom
    
    cbsThis.DockToolBar BarToDock, Right, (Bottom + Top) / 2, BarOnLeft.Position

End Sub

Private Sub vsf_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strCompare As String
    Dim blnNextRow As Boolean
    
    '如果是上下左右,吃掉
    If KeyCode = vbKeyEscape Or KeyCode = vbKeyBack Or Shift <> 0 _
        Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyRight Then
        Exit Sub
    End If
    If KeyCode = vbKeyLeft And (picInput.Visible = False And lvwMultiSel.Visible = False) Then Exit Sub
    
    blnNextRow = Val(GetSetting("ZLSOFT", "私有模块\frmCaseTendEditForBatch\" & gstrUserName, "Value")) = 0
    
    If KeyCode = vbKeyDelete Then
        '清除当前单元格的内容
        vsf.TextMatrix(vsf.ROW, vsf.Col) = ""
        Me.cbo部位.Visible = False
        Me.txt数据.Text = ""
        Me.txt数据.Height = picInput.Height
    End If
    
    If KeyCode = vbKeyReturn Then
        '问题号:56592,李涛,批量录入纵向跳转.
        If blnNextRow = False Then
toNextRow2:
            If vsf.ROW < vsf.Rows - 1 Then
                vsf.ROW = vsf.ROW + 1
                vsf.Cell(flexcpAlignment, 1, 姓名, vsf.Rows - 1, 有效数据项 - 1) = flexAlignLeftCenter
                If vsf.RowHidden(vsf.ROW) Then GoTo toNextRow2
            Else
toNextCol2:
                If vsf.Col < mlngSigner Then
                    vsf.ROW = 1
                    vsf.Col = vsf.Col + 1
                    If vsf.Col = mlngSigner Then GoTo toNextCol2
                    If vsf.ColHidden(vsf.Col) Or vsf.Col < 有效数据项 Then GoTo toNextCol2
                    
                    '不允许录入的列直接跳过
                    If vsf.Col <> mlngOper Then
                        mrsPatient.Filter = "行=" & vsf.ROW
                        If mrsPatient.RecordCount <> 0 Then
                            strCompare = mrsPatient!匹配列
                            mrsSelItems.Filter = "列=" & vsf.Col
                            If strCompare <> "" Then    '为空则说明该列是新加行,还没有录入病人信息
                                If InStr(1, "," & strCompare & ",", "," & mrsSelItems!项目序号 & ",") = 0 Then GoTo toNextCol2
                            End If
                        End If
                        mrsPatient.Filter = 0
                        mrsSelItems.Filter = 0
                    End If
                Else
                    vsf.ROW = 1
                    vsf.Col = mlngSigner - 1
                End If
            End If

            If vsf.ColIsVisible(vsf.Col) = False Then
                vsf.LeftCol = vsf.Col
            End If
            If vsf.RowIsVisible(vsf.ROW) = False Then
                vsf.TopRow = vsf.ROW
            End If
            Exit Sub

        
        Else
        
    
        '跳到下一个有效单元格
toNextCol:
            If vsf.Col < mlngSigner Then
                vsf.Col = vsf.Col + 1
                If vsf.Col = mlngSigner Then GoTo toNextCol
                If vsf.ColHidden(vsf.Col) Or vsf.Col < 有效数据项 Then GoTo toNextCol
                
                '不允许录入的列直接跳过
                If vsf.Col <> mlngOper Then
                    mrsPatient.Filter = "行=" & vsf.ROW
                    If mrsPatient.RecordCount <> 0 Then
                        strCompare = mrsPatient!匹配列
                        mrsSelItems.Filter = "列=" & vsf.Col
                        If strCompare <> "" Then    '为空则说明该列是新加行,还没有录入病人信息
                            If InStr(1, "," & strCompare & ",", "," & mrsSelItems!项目序号 & ",") = 0 Then GoTo toNextCol
                        End If
                    End If
                    mrsPatient.Filter = 0
                    mrsSelItems.Filter = 0
                End If
            Else
toNextRow:
                If vsf.ROW = vsf.Rows - 1 Then
                    vsf.Rows = vsf.Rows + 1
                    vsf.Cell(flexcpAlignment, 1, 姓名, vsf.Rows - 1, 有效数据项 - 1) = flexAlignLeftCenter
                    'Vsf.Cell(flexcpAlignment, Vsf.Rows - 1, 有效数据项, Vsf.Rows - 1, Vsf.Cols - 1) = flexAlignCenterCenter
                End If
                vsf.ROW = vsf.ROW + 1
                If vsf.RowHidden(vsf.ROW) Then GoTo toNextRow
                vsf.Col = 1
            End If
            If vsf.ColIsVisible(vsf.Col) = False Then
                vsf.LeftCol = vsf.Col
            End If
            If vsf.RowIsVisible(vsf.ROW) = False Then
                vsf.TopRow = vsf.ROW
            End If
            Exit Sub
        End If
    End If
    
    If KeyCode = vbKeyLeft Then
        '跳到上一个有效单元格
toPreCol:
        If vsf.Col > 1 Then
            vsf.Col = vsf.Col - 1
            If vsf.Col >= mlngSigner Then GoTo toPreCol
            If vsf.Col = mlngOper Then GoTo toPreCol
            If vsf.Col <> 1 And vsf.Col < 有效数据项 Then GoTo toPreCol
            If vsf.ColHidden(vsf.Col) Then GoTo toPreCol
            
            '不允许录入的列直接跳过
            If vsf.Col <> 1 Then
                mrsPatient.Filter = "行=" & vsf.ROW
                If mrsPatient.RecordCount <> 0 Then
                    strCompare = mrsPatient!匹配列
                    mrsSelItems.Filter = "列=" & vsf.Col
                    If strCompare <> "" Then    '为空则说明该列是新加行,还没有录入病人信息
                        If InStr(1, "," & strCompare & ",", "," & mrsSelItems!项目序号 & ",") = 0 Then GoTo toPreCol
                    End If
                End If
                mrsPatient.Filter = 0
                mrsSelItems.Filter = 0
            End If
        Else
toPreRow:
            If vsf.ROW > 1 Then
                vsf.ROW = vsf.ROW - 1
                vsf.Col = vsf.Cols - 1
                GoTo toPreCol
            Else
                vsf.ROW = 1
            End If
            If vsf.RowHidden(vsf.ROW) Then GoTo toPreRow
            vsf.Col = 1
        End If
        If vsf.ColIsVisible(vsf.Col) = False Then
            vsf.LeftCol = vsf.Col
        End If
        If vsf.RowIsVisible(vsf.ROW) = False Then
            vsf.TopRow = vsf.ROW
        End If
        Exit Sub
    End If
    
    mblnShow = True
    Call Vsf_EnterCell
End Sub

Private Sub vsf_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim introw As Integer, intCol As Integer
    If Button <> 1 Then Exit Sub
    
    intCol = vsf.MouseCol
    introw = vsf.MouseRow
    If introw = 0 And intCol = 0 Then
        Call vsf.Select(0, 0, vsf.Rows - 1, vsf.Cols - 1)
    ElseIf intCol = 0 Then
        Call vsf.Select(introw, 0, introw, vsf.Cols - 1)
    ElseIf introw = 0 Then
        Call vsf.Select(0, intCol, vsf.Rows - 1, intCol)
    End If
End Sub

Public Sub SignMe()
    Dim blnSign As Boolean          '是否签名成功
    Dim strSignTime As String       '保证所有签名的签名时间一致,便于取消签名时按签名时间统一取消
    Dim lngRecord As Long
    Dim str状态 As String           '保存签名选项,避免循环签名时不停的弹出签名窗口
    Dim introw As Integer, intRows As Integer
    On Error GoTo errHand
    '按发生时间循环进行签名
    
    If mblnInit = False Then Exit Sub
    
    intRows = vsf.Rows - 1
    strSignTime = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    For introw = 1 To intRows
        If vsf.TextMatrix(introw, mlngSigner) = "" And vsf.TextMatrix(introw, vsf.Cols - 2) = gstrUserName Then
            If lngRecord <> Val(vsf.TextMatrix(introw, mlngRecord)) Then
                lngRecord = Val(vsf.TextMatrix(introw, mlngRecord))
                If SignName(introw, strSignTime, str状态) = False Then Exit For
                blnSign = True
            End If
        End If
    Next
    
    If blnSign Then Call ShowMe(mfrmParent, mlng病区ID, mstrPrivs, False, False)
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub UnSignMe()
    Dim blnUnSign As Boolean
    Dim strTime As String               '记录时间
    Dim strSignTime As String           '签名时间
    Dim introw As Integer, intRows As Integer
    Dim blnClear As Boolean             '取消签名时是否清除该版本的数据回退到上次签名后的状态
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    If mblnInit = False Then Exit Sub
    strSignTime = vsf.TextMatrix(vsf.ROW, mlngSignTime)
    blnClear = (MsgBox("取消签名时是否该版本的数据回退到上次签名后的状态？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes)
    
    '把同一签名时间的数据提取出来,依次取消签名
    gstrSQL = " Select A.病人ID,A.主页ID,A.婴儿,A.发生时间,B.记录人 AS 签名人 From 病人护理记录 A,病人护理内容 B" & _
              " Where A.ID=B.记录ID And A.病人来源=2 And B.记录类型=5 And B.项目名称=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strSignTime)
  
    With rsTemp
        Do While Not .EOF
            If rsTemp!签名人 = gstrUserName Then
                If UnSignName(!病人ID, !主页ID, !婴儿, Format(!发生时间, "yyyy-MM-dd HH:mm:ss"), blnClear) = False Then Exit Sub
                blnUnSign = True
            End If
            .MoveNext
        Loop
    End With
    If blnUnSign Then Call ShowMe(mfrmParent, mlng病区ID, mstrPrivs, False, False)
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function SignName(ByVal lngRow As Long, ByVal strSignTime As String, str状态 As String) As Boolean
    Dim rs As New ADODB.Recordset
    Dim oSign As cEPRSign
    Dim strStart As String
    Dim strSource As String
    Dim lngLoop As Long
    Dim lng科室ID As Long, lng病人ID As Long, lng主页ID As Long, int婴儿 As Integer
    
    On Error GoTo errHand
    
    '初始处理
    '------------------------------------------------------------------------------------------------------------------
    strSource = ""
    strStart = Format(dtp.Value, "yyyy-MM-dd HH:mm:ss")
    
    mrsPatient.Filter = "行=" & lngRow
    If mrsPatient.RecordCount = 0 Then Exit Function
    '先定位修改过的行,再在列中循环找到修改过的列
    lng病人ID = mrsPatient!病人ID
    lng主页ID = mrsPatient!主页ID
    lng科室ID = mrsPatient!科室ID
    int婴儿 = mrsPatient!婴儿
    
    '检查当前是否已经签名了
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select 1 From 病人护理内容 a,病人护理记录 b Where b.病人id=[1] And b.主页id=[2] And b.发生时间=[3] And Nvl(b.婴儿,0)=[4] And a.记录id=b.ID And a.记录类型=5 And Nvl(a.开始版本,1)=Nvl(b.最后版本,1)"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng病人ID, lng主页ID, CDate(strStart), int婴儿)
    If rs.BOF = False Then
        MsgBox "当前没有需要签名的信息！", vbInformation, gstrSysName
        Exit Function
    End If
        
    '获取要签名的内容
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select a.记录类型,a.项目分组,a.项目序号,a.项目名称,a.项目类型,a.记录内容,a.项目单位,a.记录标记,a.体温部位,a.记录组号,a.复试合格,a.未记说明,a.记录人,a.修改时间" & vbNewLine & _
             " From 病人护理内容 a,病人护理记录 b " & vbNewLine & _
             " Where b.病人id=[1] And b.主页id=[2] And b.发生时间=[3] And Nvl(b.婴儿,0)=[4] And a.记录id=b.ID And a.终止版本 Is Null" & vbNewLine & _
             " Order by A.项目序号"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng病人ID, lng主页ID, CDate(strStart), int婴儿)
    If rs.BOF = False Then
        Do While Not rs.EOF
            For lngLoop = 0 To rs.Fields.Count - 1
                strSource = strSource & CStr(zlCommFun.NVL(rs.Fields(lngLoop).Value, ""))
            Next
            rs.MoveNext
        Loop
    End If
    If strSource = "" Then
        MsgBox "当前没有需要签名的信息！", vbInformation, gstrSysName
        Exit Function
    End If
    '76223:刘鹏飞,2014-08-05,电子签名添加时间戳信息
    '------------------------------------------------------------------------------------------------------------------
    Set oSign = frmCaseTendSign.ShowMe(Me, mstrPrivs, strSource, lng病人ID, lng主页ID, mlng病区ID, str状态)
    If Not oSign Is Nothing Then
        gstrSQL = "ZL_电子护理记录_SignName("
        gstrSQL = gstrSQL & lng病人ID & "," & lng主页ID & "," & int婴儿 & ","
        gstrSQL = gstrSQL & "To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'),"
        gstrSQL = gstrSQL & "'" & oSign.姓名 & "',"
        gstrSQL = gstrSQL & "'" & oSign.签名信息 & "',"
        gstrSQL = gstrSQL & oSign.证书ID & ","
        gstrSQL = gstrSQL & oSign.签名方式 & ",'" & oSign.时间戳 & "','" & oSign.时间戳信息 & "')"

        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        SignName = True
    End If
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CheckSigned(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal int婴儿 As Integer, ByVal str发生时间 As String, _
    Optional ByRef lng证书ID As Long, Optional ByRef str签名人 As String, Optional ByRef str签名时间 As String, Optional ByVal blnCheck As Boolean = True) As Boolean
    Dim rs As New ADODB.Recordset
    On Error GoTo errHand
    
    '检查当前是否已经签名了
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select A.项目ID AS 证书ID,A.记录人 AS 签名人,A.项目名称 AS 签名时间 From 病人护理内容 a,病人护理记录 b Where b.病人id=[1] And b.主页id=[2] And b.发生时间=[3] And Nvl(b.婴儿,0)=[4] And a.记录id=b.ID And a.记录类型=5"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng病人ID, lng主页ID, CDate(str发生时间), int婴儿)
    If rs.BOF Then
        If blnCheck Then MsgBox "当前没有需要取消的签名！", vbInformation, gstrSysName
        Exit Function
    End If
    
    lng证书ID = NVL(rs!证书ID, 0)
    str签名人 = NVL(rs!签名人)
    str签名时间 = NVL(rs!签名时间)
    CheckSigned = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function UnSignName(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal int婴儿 As Integer, ByVal str发生时间 As String, ByVal blnClear As Boolean) As Boolean
    '******************************************************************************************************************
    '功能:
    '
    '
    '******************************************************************************************************************
    Dim lng证书ID As Long
    Dim strSource As String
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    '检查当前是否已经签名了
    '------------------------------------------------------------------------------------------------------------------
    If Not CheckSigned(lng病人ID, lng主页ID, int婴儿, str发生时间, lng证书ID) Then Exit Function
    
    '如果是电子签名,则需要验证
    '------------------------------------------------------------------------------------------------------------------
    If lng证书ID > 0 Then
        '数字签名验证
        Err.Clear
        If gobjTendESign Is Nothing Then
            On Error Resume Next
            Set gobjTendESign = CreateObject("zl9ESign.clsESign")
            If Err <> 0 Then Err.Clear
            On Error GoTo 0
            If Not gobjTendESign Is Nothing Then Call gobjTendESign.Initialize(gcnOracle, glngSys)
        End If
        If Not gobjTendESign Is Nothing Then
            If Not gobjTendESign.CheckCertificate(gstrDBUser) Then Exit Function
        Else
            MsgBox "电子签名部件未能正确安装，回退操作不能继续！", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
    End If

    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Zl_电子护理记录_Unsignname("
    gstrSQL = gstrSQL & lng病人ID & ","
    gstrSQL = gstrSQL & lng主页ID & ","
    gstrSQL = gstrSQL & int婴儿 & ","
    gstrSQL = gstrSQL & "To_Date('" & str发生时间 & "','yyyy-mm-dd hh24:mi:ss')," & _
                      IIf(blnClear, "1", "0") & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    UnSignName = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function SaveME() As Boolean
    If Not CheckData Then Exit Function
    If Not SaveData Then Exit Function
    mblnShow = False
    picInput.Visible = False
    SaveME = True
End Function

Public Sub ShowMe(ByVal frmParent As Form, ByVal lng病区ID As Long, Optional ByVal strPrivs As String, _
    Optional ByVal blnCancel As Boolean = False, Optional ByVal blnShow As Boolean = True)
    '******************************************************************************************************************
    '功能： 显示护理记录文件内容
    '参数： frmParent           上级窗体对象
    '       lngPatiID           病人id
    '       lngPageID           主页id
    '       lngDeptID           要显示护理记录的科室
    '       intBaby             婴儿标志
    '返回： 无
    '******************************************************************************************************************
'    Dim bln护理级别 As Boolean
    
    Err = 0
    Dim lngRow As Long
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    mstrPrivs = strPrivs
    mlng病区ID = lng病区ID
    Set mfrmParent = frmParent
    If mlng病区ID = 0 Then Exit Sub

    mstrSel = ""
    mblnShow = False
    picInput.Visible = False

    Call ReadData
    
    mblnChange = False
    RaiseEvent AfterRefresh

    Call vsf_AfterRowColChange(2, 2, 1, 1)
    Call dkpMain.RecalcLayout
    
    '设置某些列不移动
    vsf.FrozenCols = 有效数据项 - 1
    vsf.SheetBorder = &HC0C0FF
    
    If blnShow Then Me.Show 1, frmParent
    Exit Sub
    
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function CheckData() As Boolean
    Dim StrText As String
    Dim strMaxDate As String, str值域 As String
    Dim lngRow As Long, lngRows As Long, lngCol As Long
    Dim intType As Integer, lngOrder As Long, lngClass As Long, strName As String, lngLength As Long
    On Error GoTo errHand
    '检查数据录入合法性
    
    lngRows = vsf.Rows - 1
    
    '依次检查各个项目的录入合法性
    With mrsSelItems
        .MoveFirst
        Do While Not .EOF
            mrsItems.Filter = "项目序号=" & !项目序号
            If mrsItems.RecordCount <> 0 Then
                lngCol = !列
                intType = mrsItems!项目类型     '0-数值；1-文字
                lngClass = mrsItems!项目性质
                lngOrder = mrsItems!项目序号
                strName = mrsItems!项目名称
                lngLength = mrsItems!项目长度 + IIf(NVL(mrsItems!项目小数, 0) = 0, 0, NVL(mrsItems!项目小数, 0) + 1)
                If intType = 0 Then
                    str值域 = NVL(mrsItems!项目值域)
                Else
                    str值域 = ""
                End If
                '数值项目:只有体温,呼吸与脉搏,以及血压才存在/录入
                '文本项目:只检查是否超长
                
                For lngRow = 1 To lngRows
                    If Val(vsf.Cell(flexcpData, lngRow, lngCol)) = 1 Then
                        StrText = vsf.TextMatrix(lngRow, lngCol)
                        If Trim(StrText) <> "" Then
                            If Not CheckValid(StrText, lngOrder, lngClass, strName, lngLength, lngRow, lngCol, str值域) Then
                                vsf.ROW = lngRow
                                If vsf.RowIsVisible(vsf.ROW) Then vsf.TopRow = vsf.ROW
                                Exit Function
                            End If
                        End If
                    End If
                Next
            End If
            
            .MoveNext
        Loop
    End With
    
    mrsItems.Filter = 0
    CheckData = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    mrsItems.Filter = 0
End Function

Private Function CheckValid(ByRef StrText As String, ByVal lngOrder As Long, ByVal lngClass As Long, ByVal strCap As String, _
    ByVal lngLength As Long, ByVal lngRow As Long, ByVal lngCol As Long, Optional ByVal str值域 As String) As Boolean
    Dim arrData
    Dim intDo As Integer, intCount As Integer
    Dim strPart As String, strValue1 As String, strValue2 As String, strTextClone As String
    
    If StrText = "" Then
        CheckValid = True
        Exit Function
    End If
    
    '先取出部位,上/下数据
    strTextClone = StrText
    If InStr(1, strTextClone, ":") <> 0 Then
        strPart = Split(strTextClone, ":")(0)
        strTextClone = Split(strTextClone, ":")(1)
    End If
    If InStr(1, strTextClone, "/") <> 0 Then
        strValue1 = Split(strTextClone, "/")(0)
        strValue2 = Split(strTextClone, "/")(1)
    Else
        strValue1 = strTextClone
    End If
    
    If lngClass = 2 Then '如果是活动项目则可能存在部位,把部位提出来,只检查录入的数据是否超过限制
        If InStr(1, StrText, ":") <> 0 Then
            StrText = Split(StrText, ":")(1)
        End If
    End If
    
'    If str值域 = "" Then  '普通项目
'        If Not (lngOrder = 9 Or lngOrder = 10) Then '大便次数与排出量不进行有效范围检查
'            If LenB(StrConv(strText, vbFromUnicode)) > lngLength Then
'                MsgBox "第" & lngRow & "行的" & strCap & "超长，请检查！", vbInformation, gstrSysName
'                Exit Function
'            End If
'        End If
'    Else                    '体温脉搏呼吸以及血压
        '没有心率的时候，才允许录入脉搏
        If lngOrder = 2 And mbln心率 Then
            If InStr(1, StrText, "/") <> 0 Then
                MsgBox "请将测得的心率数据录入单独的心率单元格中！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        If lngOrder = 3 Then
            If InStr(1, StrText, "/") <> 0 Then
                MsgBox "呼吸数据录入错误！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        If lngOrder = 4 Or lngOrder = 5 Then
            '血压值必须含/
            If vsf.TextMatrix(0, lngCol) Like "血压*" Then
                If InStr(1, StrText, "/") = 0 Then
                    MsgBox "血压数据的格式错误：收缩压/舒张压！", vbInformation, gstrSysName
                    Exit Function
                End If
                If Trim(Split(StrText, "/")(0)) = "" Or Trim(Split(StrText, "/")(1)) = "" Then
                    MsgBox "血压数据错误：收缩压/舒张压！", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
        If UBound(Split(StrText, "/")) > 1 Then
            MsgBox "第" & lngRow & "行的" & strCap & "数据录入错误，请检查！", vbInformation, gstrSysName
            Exit Function
        End If
        
        arrData = Split(StrText, "/")
        intCount = UBound(arrData)
        For intDo = 0 To intCount
            StrText = arrData(intDo)
            If InStr(1, StrText, ":") <> 0 Then StrText = Split(StrText, ":")(1)
            '曲线项目不检查是否超长
            If lngOrder > 3 Then
                If LenB(StrConv(StrText, vbFromUnicode)) > lngLength Then
                    MsgBox "第" & lngRow & "行的" & strCap & "超长，请检查！", vbInformation, gstrSysName
                    vsf.TopRow = lngRow
                    Exit Function
                End If
            End If
            If IsNumeric(StrText) Then    '有效范围与当前录入值都是数值型才检查,否则当成是未记说明
                If Not (lngOrder = 9 Or lngOrder = 10) Then '大便次数与排出量不进行有效范围检查
                    If str值域 <> "" Then
                        If IsNumeric(Split(str值域, ";")(0)) Then
                            If Not (Val(StrText) >= Split(str值域, ";")(0) And Val(StrText) <= Split(str值域, ";")(1)) Then
                                MsgBox "第" & lngRow & "行的" & strCap & "超出有效范围（" & Split(str值域, ";")(0) & "-" & Split(str值域, ";")(1) & "），请检查！", vbInformation, gstrSysName
                                Exit Function
                            End If
                        End If
                    End If
                End If
                If mrsItems!项目类型 = 0 Then
                    If NVL(mrsItems!项目小数, 0) <> 0 Then
                        If intDo = 0 Then
                            strValue1 = Format(StrText, "#0." & String(mrsItems!项目小数, "0"))
                        Else
                            strValue2 = Format(StrText, "#0." & String(mrsItems!项目小数, "0"))
                        End If
                    Else
                        If intDo = 0 Then
                            strValue1 = Format(StrText, "#0")
                        Else
                            strValue2 = Format(StrText, "#0")
                        End If
                    End If
                End If
            End If
        Next
'    End If
    
    '拼装输入串
    StrText = IIf(strPart <> "", strPart & ":", "") & strValue1 & IIf(strValue2 <> "", "/" & strValue2, "")
    CheckValid = True
End Function

Private Function SaveData() As Boolean
    Dim blnTrans As Boolean, blnOper As Boolean         '指定某个时间段里是否出现手术
    Dim lngOrder As Long, lng科室ID As Long, lng病人ID As Long, lng主页ID As Long, int婴儿 As Integer
    Dim strTmp As String
    Dim intAllow As Integer, intType As Integer, lngClass As Long
    Dim str内容 As String, str标记 As String, str部位 As String, str未记说明 As String 'str标记:只保存特殊降温或脉搏短拙
    Dim lngRecord As Long, lngGroup As Long
    Dim lngRow As Long, lngRows As Long, lngCol As Long, lngCols As Long
    Dim strDate As String, strStart As String, strEnd As String, strSQLtmp As String
    Dim rsTemp As New ADODB.Recordset
    
    Dim intPos As Integer, intMax As Integer
    Dim strSQL() As String
    On Error GoTo errHand
    '同一个时间里(同一条记录ID),不允许出现多组手术,也就是只允许一个组号里有手术的存在
    
    ReDim Preserve strSQL(1 To 1)
    lngRows = vsf.Rows - 1
    lngCols = mlngSigner - 1         '后面的签名人,签名时间,记录ID,组号不处理
    intAllow = IIf(InStr(mstrPrivs, "他人护理记录") > 0, 1, 0)
    strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    '准备保存数据
    lng病人ID = 0
    lng主页ID = 0
    lng科室ID = 0
    int婴儿 = 0
    For lngRow = 1 To lngRows
        mrsPatient.Filter = "行=" & lngRow
        If mrsPatient.RecordCount <> 0 Then
            '先定位修改过的行,再在列中循环找到修改过的列
            If lng病人ID <> mrsPatient!病人ID Or lng主页ID <> mrsPatient!主页ID Or int婴儿 <> mrsPatient!婴儿 Then
                lng病人ID = mrsPatient!病人ID
                lng主页ID = mrsPatient!主页ID
                lng科室ID = mrsPatient!科室ID
                int婴儿 = mrsPatient!婴儿
                blnOper = False
            End If
            
            strStart = Format(Me.dtp.Value, "yyyy-MM-dd HH:mm:ss")
            strEnd = Format(DateAdd("n", 1, CDate(strStart)), "yyyy-MM-dd HH:mm:ss")
            
            '数据发生时间不能在当前操作员所属科室的有效时间以前
            If Val(vsf.RowData(lngRow)) = 1 Then
                If Not CheckTime(lngRow, lng病人ID, lng主页ID, Mid(strStart, 1, 16), Mid(strDate, 1, 16)) Then Exit Function
            End If
            
            '如果提取数据后修改了时间,需要更新时间
            If lngRecord <> Val(vsf.TextMatrix(lngRow, mlngRecord)) And Val(vsf.TextMatrix(lngRow, mlngRecord)) <> 0 Then
                lngRecord = Val(vsf.TextMatrix(lngRow, mlngRecord))
                gstrSQL = "Zl_病人护理记录_UpdateReplace(" & lngRecord & ",0," & int婴儿 & ",To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'))"
                strSQL(ReDimArray(strSQL)) = gstrSQL
            End If
            
            If Val(vsf.RowData(lngRow)) = 1 Then
                '有组号则取组号，无组号，则取当前最大组号
                lngRecord = Val(vsf.TextMatrix(lngRow, mlngRecord))
                lngGroup = Val(vsf.TextMatrix(lngRow, mlngGroup))
                '有可能原来的数据中的组号不是按顺序增加的,因此此段进行校正
                If lngGroup = 0 Then
                    '取最大的组号
                    gstrSQL = " select max(记录组号) AS 组号 " & _
                              " From 病人护理内容" & _
                              " where 记录ID=(" & _
                              "     select ID from 病人护理记录" & _
                              "     where 病人ID=[1] and 主页ID=[2] and 婴儿=[3] and 科室ID=[4] and 发生时间=[5])"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng病人ID, lng主页ID, int婴儿, lng科室ID, CDate(strStart))
                    lngGroup = NVL(rsTemp!组号, 0) + 1
                End If
                
                '一个元素一个元素的处理
                For lngCol = 有效数据项 To lngCols
                    If Val(vsf.Cell(flexcpData, lngRow, lngCol)) = 1 Then
                        
                        '此数据进行了新增或修改操作
                        gstrSQL = "Zl_病人护理记录_UpdateRecord("
                        gstrSQL = gstrSQL & mrsPatient!病人ID & "," & mrsPatient!主页ID & "," & mrsPatient!婴儿 & ","
                        gstrSQL = gstrSQL & "To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'),"
                        gstrSQL = gstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
                        gstrSQL = gstrSQL & IIf(lngCol <> mlngOper, 1, 4) & ","
                        
                        lngOrder = 0
                        If lngCol <> mlngOper Then
                            mrsSelItems.Filter = "列=" & lngCol
                            mrsItems.Filter = "项目序号=" & mrsSelItems!项目序号
                            lngClass = mrsItems!项目性质
                            intType = mrsItems!项目类型
                            lngOrder = mrsItems!项目序号
                        End If
                        strSQLtmp = gstrSQL
                        gstrSQL = gstrSQL & lngOrder & ","
                        
                        str部位 = "": str标记 = "": str未记说明 = ""
                        str内容 = vsf.TextMatrix(lngRow, lngCol)
                        If (lngOrder = 1 Or lngOrder = 2 Or lngOrder = 3) Or lngClass = 2 Then
                            If InStr(1, str内容, ":") <> 0 Then
                                str部位 = Trim(Split(str内容, ":")(0))
                                str内容 = Trim(Split(str内容, ":")(1))
                            End If
                            If InStr(1, str内容, "/") <> 0 Then
                                str标记 = Trim(Split(str内容, "/")(1))
                                str内容 = Trim(Split(str内容, "/")(0))
                            End If
                        ElseIf lngOrder = 4 Then        '因为是按列循环,所以只会处理一次,如果是合并录入收缩压与舒张压,则在保存后再处理下
                            If InStr(1, str内容, "/") <> 0 Then
                                str内容 = Split(str内容, "/")(lngOrder - 4)
                            End If
                        End If
                        '只有曲线项目才存在未记说明的概念
                        If lngOrder <= 3 And Not IsNumeric(str内容) And lngCol <> mlngOper Then
                            If (lngOrder = 1 And str内容 <> "不升") Or lngOrder <> 1 Then
                                str未记说明 = str内容
                                str内容 = ""
                            End If
                        End If
                        
                        '体温脉搏项目,如果有/填1
                        If lngOrder = -1 Then
                            gstrSQL = gstrSQL & "1,"
                        Else
                            gstrSQL = gstrSQL & "0,"
                        End If
                        
                        If lngCol <> mlngOper Or blnOper = False Then
                            gstrSQL = gstrSQL & "'" & str内容 & "','" & str部位 & "'," & intAllow & "," & IIf(IsNumeric(str内容), 0, 1) & "," & lngGroup & ",'" & str未记说明 & "')"
                            strSQL(ReDimArray(strSQL)) = gstrSQL
                        
                            '如果是血压
                            If lngOrder = 4 And vsf.TextMatrix(0, lngCol) Like "血压*" Then
                                If str内容 <> "" Then str内容 = Split(vsf.TextMatrix(lngRow, lngCol), "/")(1)       '不为空时进行赋值,为空则说明现在是清除数据
                                strSQLtmp = strSQLtmp & "5,0,"
                                gstrSQL = strSQLtmp & "'" & str内容 & "','" & str部位 & "'," & intAllow & "," & IIf(IsNumeric(str内容), 0, 1) & "," & lngGroup & ",'" & str未记说明 & "')"
                                strSQL(ReDimArray(strSQL)) = gstrSQL
                            End If
                                
                            If lngCol = mlngOper Then blnOper = True
                        End If
                        
                        '----------------------------------------------------------------------------
                        '没有选择心率,就允许他在脉搏处同时录入(如果都为空,完成标记部分数据清除的功能)
                        If (lngOrder = 1 Or lngOrder = 2 And mbln心率 = False) Then
            
                            gstrSQL = "Zl_病人护理记录_UpdateRecord("
                            gstrSQL = gstrSQL & lng病人ID & "," & lng主页ID & "," & int婴儿 & ","
                            gstrSQL = gstrSQL & "To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'),"
                            gstrSQL = gstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
                            gstrSQL = gstrSQL & "1,"
                            gstrSQL = gstrSQL & IIf(lngOrder = 2, -1, lngOrder) & ","
                            gstrSQL = gstrSQL & "1,"
                                                            
                            If str标记 <> "" And str内容 <> "" Then
                                Select Case intType
                                Case 0          '数值
                                    strTmp = Val(str标记)
                                Case 1          '文本
                                    strTmp = str标记
                                End Select
                                gstrSQL = gstrSQL & "'" & strTmp & "','" & str部位 & "'," & intAllow & "," & IIf(IsNumeric(strTmp), 0, 1) & "," & lngGroup & ",Null)"
                            Else
                                gstrSQL = gstrSQL & "NULL,'" & str部位 & "'," & intAllow & ",0," & lngGroup & ",Null)"
                            End If
                            strSQL(ReDimArray(strSQL)) = gstrSQL
                        End If
                    End If
                Next
            End If
        End If
    Next
    
    '循环执行SQL保存数据
    gcnOracle.BeginTrans
    blnTrans = True
    intMax = UBound(strSQL)
    For intPos = 1 To intMax
        If strSQL(intPos) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(intPos), Me.Caption)
    Next
    SaveData = True
    gcnOracle.CommitTrans
    blnTrans = False
    
    mblnChange = False
    mrsItems.Filter = 0
    mrsSelItems.Filter = 0
    mrsPatient.Filter = 0
    
    RaiseEvent AfterDataChanged
    Exit Function
    
errHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    mrsItems.Filter = 0
    mrsSelItems.Filter = 0
    mrsPatient.Filter = 0
    lng科室ID = Me.cbo科室.ItemData(Me.cbo科室.ListIndex)  '必须设置
End Function


'---------------------------------------------------------------------------------
'以下是基础函数或过程
'---------------------------------------------------------------------------------
Public Sub Record_Add(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String)
    Dim arrFields, arrValues, intField As Integer
    '添加记录
    'strFields:字段名|字段名
    'strValues:值|值
    
    '例子：
    'Dim strFields As String, strValues As String
    'strFields = "RecordID|科目ID|摘要"
    'strValues = "5188|6666|科目名称"
    'Call Record_Update(rsVoucher, strFields, strValues)

    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField = 0 Then Exit Sub

    With rsObj
        .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Public Sub Record_Update(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String, ByVal strPrimary As String, Optional ByVal blnDelete As Boolean = False)
    Dim arrFields, arrValues, intField As Integer
    '更新记录,如果不存在,则新增
    'strPrimary:字段名,值
    'strFields:字段名|字段名
    'strValues:值|值
    
    '例子：
    'Dim strFields As String, strValues As String, strPrimary As String
    'strFields = "RecordID|科目ID|摘要"
    'strValues = "5188|6666|科目名称"
    'strPrimary = "RecordID,5188"
    'Call Record_Update(rsVoucher, strFields, strValues, strPrimary, True)

    If strValues = "" Then strValues = " "
    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField < 0 Then Exit Sub

    With rsObj
        If Record_Locate(rsObj, strPrimary, blnDelete) = False Then .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Public Function Record_Locate(ByRef rsObj As ADODB.Recordset, ByVal strPrimary As String, Optional ByVal blnDelete As Boolean = False) As Boolean
    Dim arrTmp
    '定位到指定记录
    'strPrimary:主健,值
    'blnDelete=True,则该记录集存在"删除"字段
    Record_Locate = False
    
    arrTmp = Split(strPrimary, "|")
    With rsObj
        If .RecordCount = 0 Then Exit Function
        .MoveFirst
        .Find arrTmp(0) & "='" & arrTmp(1) & "'"
        If .EOF Then Exit Function
        If blnDelete Then
            Do While Not .EOF
                If !删除 = 0 Then Record_Locate = True: Exit Do
                .MoveNext
            Loop
        Else
            Record_Locate = True
        End If
    End With
End Function

Public Sub Record_Init(ByRef rsObj As ADODB.Recordset, ByVal strFields As String)
    Dim arrFields, intField As Integer
    Dim strFieldName As String, intType As Integer, lngLength As Long
    '初始化映射记录集
    'strFields:字段名,类型,长度|字段名,类型,长度    如果长度为零,则取默认长度
    '字符型:adLongVarChar;数字型:adDouble;日期型:adDBDate
    
    '例子：
    'Dim rsVoucher As New ADODB.Recordset, strFields As String
    'strFields = "RecordID," & adDouble & ",18|科目ID," & adDouble & ",18|摘要, " & adLongVarChar & ",50|" & _
    '"删除," & adDouble & ",1"
    'Call Record_Init(rsVoucher, strFields)

    arrFields = Split(strFields, "|")
    Set rsObj = New ADODB.Recordset

    With rsObj
        If .State = 1 Then .Close
        For intField = 0 To UBound(arrFields)
            strFieldName = Split(arrFields(intField), ",")(0)
            intType = Split(arrFields(intField), ",")(1)
            lngLength = Split(arrFields(intField), ",")(2)

            '获取字段缺省长度
            If lngLength = 0 Then
                Select Case intType
                Case adDouble
                    lngLength = madDoubleDefault
                Case adVarChar
                    lngLength = madLongVarCharDefault
                Case adLongVarChar
                    lngLength = madLongVarCharDefault
                Case Else
                    lngLength = madDbDateDefault
                End Select
            End If
            .Fields.Append strFieldName, intType, lngLength, adFldIsNullable
        Next
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Sub OutputRsData(ByVal rsObj As ADODB.Recordset)
    Dim intCol As Integer, intCols As Integer
    With rsObj
        Do While Not .EOF
            Debug.Print !列 & "," & !项目序号 & "," & !项目名称
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
    End With
End Sub

Private Function CheckVersion(Optional ByVal lngRow As Long = 0, Optional ByVal lngCol As Long = 0) As Boolean
    Dim lng项目序号 As Long
    Dim lng当前版本 As Long
    Dim lng最高版本 As Long
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '因手术只产生一条记录,只有当签名记录的最大版本小于手术数据的开始版本时,才允许进行编辑(含清除功能)
    '如果要清除一行,该行存在手术记录,如果不允许对手术列进行编辑,则取消该操作
    
    If lngRow = 0 Then lngRow = vsf.ROW
    If lngCol = 0 Then lngCol = vsf.Col
    If Val(vsf.TextMatrix(lngRow, mlngRecord)) = 0 Then CheckVersion = True: Exit Function      '新记录直接退出
    If vsf.Cell(flexcpData, lngRow, lngCol) <> 0 Then CheckVersion = True: Exit Function                              '本次新增的数据允许清除
    
    '取当前单元格的项目序号
    mrsSelItems.Filter = "列=" & lngCol
    If mrsSelItems.RecordCount <> 0 Then
        lng项目序号 = mrsSelItems!项目序号
    Else
        mrsSelItems.Filter = 0
        Exit Function
    End If
    mrsSelItems.Filter = 0
    
    '取当前记录+组号的最大版本
    gstrSQL = " Select Max(开始版本) AS 最高版本 From 病人护理内容 Where 记录ID=[1] And 记录类型=5"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取当前记录+组号的最大版本", Val(vsf.TextMatrix(lngRow, mlngRecord)), Val(vsf.TextMatrix(lngRow, mlngGroup)))
    lng最高版本 = NVL(rsTemp!最高版本, 0)
    
    '取当前项目的当前版本
    gstrSQL = " Select MAX(开始版本) AS 当前版本 From 病人护理内容 Where 记录ID=[1] And 记录组号=[2]" & IIf(lngCol = mlngOper, " And 记录类型=4", " And 项目序号=[3]")
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取当前记录+组号的最大版本", Val(vsf.TextMatrix(lngRow, mlngRecord)), Val(vsf.TextMatrix(lngRow, mlngGroup)), lng项目序号)
    lng当前版本 = NVL(rsTemp!当前版本, 1)
    
    '只有当前版本大于最高版本,才允许清除(签名的数据也不允许清除)
    '同时如果最高版本=1,且签名人为空,也允许清除
    CheckVersion = ((lng当前版本 > lng最高版本) Or (lng最高版本 = 1 And vsf.Cell(flexcpForeColor, lngRow, lngCol) = &HFF&))
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CheckTime(ByVal lngRow As Long, ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
    ByVal strTime As String, ByVal strCurTime As String) As Boolean
    Dim blnExist As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '数据发生时间必须在当前科室的有效时间范围内
    
    gstrSQL = " Select 开始原因,病区ID,to_char(开始时间,'yyyy-MM-dd hh24:mi') AS 开始时间,to_char( nvl(终止时间,sysDate+" & mintPreDays & "),'yyyy-MM-dd hh24:mi') AS 终止时间 " & _
              " From 病人变动记录 " & _
              " Where 病人ID=[1] And 主页ID=[2]" & _
              " Order by 开始时间,开始原因"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取当前科室有效时间范围", lng病人ID, lng主页ID)
    With rsTemp
        .Filter = "病区ID=" & mlng病区ID
        Do While Not .EOF
            If strTime >= !开始时间 And strTime <= !终止时间 Then
                blnExist = True
                Exit Do
            End If
            .MoveNext
        Loop
        .Filter = 0
        '找到了就退出
        If blnExist Then
            If Not IsAllowInput(lng病人ID, lng主页ID, strTime, strCurTime) Then
                MsgBox "第" & lngRow & "行的发生时间" & strTime & "有误！[超过数据补录的有效时限:" & glngHours & "小时]", vbInformation, gstrSysName
                Exit Function
            End If
            
            CheckTime = True
            Exit Function
        End If
        
        '没找到,就整理原因进行准确性提示
        .Filter = "开始原因=1"
        If .RecordCount <> 0 Then
            If !开始原因 = 1 And strTime < !开始时间 Then
                MsgBox "第" & lngRow & "行的发生时间" & strTime & "有误！[发生时间不能小于病人入院时间:" & !开始时间 & "]", vbInformation, gstrSysName
                GoTo exitHand
            End If
        End If
        .Filter = "开始原因=2"
        If .RecordCount <> 0 Then
            If !开始原因 = 2 And strTime < !开始时间 Then
                MsgBox "第" & lngRow & "行的发生时间" & strTime & "有误！[发生时间不能小于病人入科时间:" & !开始时间 & "]", vbInformation, gstrSysName
                GoTo exitHand
            End If
        End If
        .Filter = "开始原因=10"
        If .RecordCount <> 0 Then
            If !开始原因 = 10 And strTime > !终止时间 Then
                MsgBox "第" & lngRow & "行的发生时间" & strTime & "有误！[发生时间不能大于出院时间:" & !终止时间 & "]", vbInformation, gstrSysName
                GoTo exitHand
            End If
        End If
        .Filter = 0
        '其他情况说明
        MsgBox "第" & lngRow & "行的发生时间" & strTime & "有误！[不在当前病区的有效时间范围内]", vbInformation, gstrSysName
        GoTo exitHand
    End With
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
exitHand:
    rsTemp.Filter = 0
End Function
