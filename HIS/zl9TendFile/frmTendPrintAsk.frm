VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTendPrintAsk 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "记录单打印"
   ClientHeight    =   4500
   ClientLeft      =   2550
   ClientTop       =   2625
   ClientWidth     =   6165
   HelpContextID   =   10322
   Icon            =   "frmTendPrintAsk.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin TabDlg.SSTab SSTPrint 
      Height          =   4380
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   7726
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "打印选项"
      TabPicture(0)   =   "frmTendPrintAsk.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "picPrint"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "清除打印"
      TabPicture(1)   =   "frmTendPrintAsk.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraPrint(0)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "打印参数"
      TabPicture(2)   =   "frmTendPrintAsk.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "chkPrintSet(0)"
      Tab(2).Control(1)=   "chkPrintSet(1)"
      Tab(2).Control(2)=   "chkPrintSet(2)"
      Tab(2).ControlCount=   3
      Begin VB.CheckBox chkPrintSet 
         Caption         =   "预览、打印时数据未满页部分固定输出表格"
         Height          =   255
         Index           =   0
         Left            =   -74850
         TabIndex        =   13
         Top             =   1230
         Width           =   3795
      End
      Begin VB.CheckBox chkPrintSet 
         Caption         =   "预览、打印时数据满页才进行输出(文件未结束有效)"
         Height          =   255
         Index           =   1
         Left            =   -74850
         TabIndex        =   12
         Top             =   810
         Width           =   4440
      End
      Begin VB.CheckBox chkPrintSet 
         Caption         =   "打印时数据页奇偶输出(不勾则按页号顺序输出)"
         Height          =   255
         Index           =   2
         Left            =   -74850
         TabIndex        =   11
         Top             =   1680
         Width           =   4080
      End
      Begin VB.Frame fraPrint 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1785
         Index           =   0
         Left            =   -74940
         TabIndex        =   6
         Tag             =   "清除重打"
         Top             =   930
         Width           =   4575
         Begin VB.TextBox txtClear 
            BackColor       =   &H00FFFFFF&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1710
            MaxLength       =   5
            TabIndex        =   8
            Top             =   300
            Width           =   1035
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "清除"
            Height          =   350
            Left            =   2760
            TabIndex        =   7
            Top             =   270
            Width           =   705
         End
         Begin VB.Label lblTag 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "有效页码范围:第1页 ～ 第5页"
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   0
            Left            =   1095
            TabIndex        =   14
            Top             =   810
            Width           =   2430
         End
         Begin VB.Label lblTag 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "清除从起始页开始的所有打印数据，可重新打印"
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   3
            Left            =   390
            TabIndex        =   10
            Top             =   1290
            Width           =   3780
         End
         Begin VB.Label lblPage 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "起始页"
            Height          =   180
            Left            =   1110
            TabIndex        =   9
            Top             =   360
            Width           =   540
         End
      End
      Begin VB.PictureBox picPrint 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3870
         Left            =   30
         ScaleHeight     =   3870
         ScaleWidth      =   4605
         TabIndex        =   1
         Top             =   450
         Width           =   4605
         Begin VSFlex8Ctl.VSFlexGrid vfgPrint 
            Height          =   2655
            Left            =   90
            TabIndex        =   2
            Top             =   675
            Width           =   3405
            _cx             =   6006
            _cy             =   4683
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
            BackColorSel    =   -2147483643
            ForeColorSel    =   0
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   7
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   255
            RowHeightMax    =   5000
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmTendPrintAsk.frx":0496
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
            WordWrap        =   0   'False
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            DataMode        =   0
            VirtualData     =   -1  'True
            DataMember      =   ""
            ComboSearch     =   3
            AutoSizeMouse   =   0   'False
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
         Begin XtremeCommandBars.CommandBars cbsMain 
            Left            =   90
            Top             =   60
            _Version        =   589884
            _ExtentX        =   635
            _ExtentY        =   635
            _StockProps     =   0
         End
      End
   End
   Begin VB.CommandButton cmdEXCEL 
      Caption         =   "输出到&Excel"
      Height          =   350
      Left            =   4830
      TabIndex        =   5
      Top             =   4080
      Width           =   1245
   End
   Begin VB.CommandButton cmdPreView 
      Caption         =   "预览(&V)"
      Height          =   350
      Left            =   4830
      TabIndex        =   3
      Top             =   390
      Width           =   1245
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "打印(&P)"
      Height          =   350
      Left            =   4830
      TabIndex        =   4
      Top             =   870
      Width           =   1245
   End
   Begin MSComDlg.CommonDialog comDlg 
      Bindings        =   "frmTendPrintAsk.frx":058D
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgData 
      Left            =   660
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTendPrintAsk.frx":05A1
            Key             =   "已打"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTendPrintAsk.frx":0B3B
            Key             =   "未打"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTendPrintAsk.frx":10D5
            Key             =   "续打"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTendPrintAsk.frx":166F
            Key             =   "重打"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   3810
      Top             =   1170
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
            Picture         =   "frmTendPrintAsk.frx":1C09
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmTendPrintAsk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Public mbytRunMode As Byte        '执行方式
Public mintPageRows As Integer
Public mstrPrintPages As String  '格式：页码;标识(续打或正常打印),页码;标识......

Private mrsData As New ADODB.Recordset
Private mstrSQL As String
Private mbytFileState As Byte
Private Type Type_DataState
    bln待打 As Boolean
    bln重打 As Boolean
    bln所有 As Boolean
End Type
Private mDataState As Type_DataState

Private Enum E_CommandBarId
    ID_待打 = 1
    ID_重打 = 2
    ID_所有 = 3
End Enum

Public Property Get FileID() As Long
    FileID = glng文件ID
End Property

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim intPrintTag As Integer
    Select Case Control.ID
        Case ID_待打
            mbytFileState = 0
            Call LoadPrintData
        Case ID_重打
            mbytFileState = 1
            Call LoadPrintData
        Case ID_所有
            mbytFileState = 2
            Call LoadPrintData
        Case conMenu_View_Refresh
            Call zlRefresh(glng文件ID)
        Case conMenu_Edit_SelAll * 100# + 1, conMenu_Edit_SelAll * 100# + 2, conMenu_Edit_SelAll * 100# + 3, conMenu_Edit_SelAll * 100# + 4
            If Control.ID = conMenu_Edit_SelAll * 100# + 1 Then
                intPrintTag = 0
            ElseIf Control.ID = conMenu_Edit_SelAll * 100# + 2 Then
                intPrintTag = -1
            ElseIf Control.ID = conMenu_Edit_SelAll * 100# + 3 Then
                intPrintTag = 1
            Else
                intPrintTag = 2
            End If
            Call RevfgPrint(1, intPrintTag)
        Case conMenu_Edit_SelAll * 100# + 5
            Call RevfgPrint(3, intPrintTag)
        Case conMenu_Edit_SelAll * 100# + 6
            Call RevfgPrint(4, intPrintTag)
    End Select
End Sub

Private Sub cbsMain_Resize()
    Dim lngTop As Long, lngLeft As Long, lngRight As Long, lngBottom As Long
    On Error Resume Next
    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    With vfgPrint
        .Left = lngLeft
        .Top = lngTop
        .Width = lngRight - lngLeft
        .Height = lngBottom - lngTop
    End With
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim intPrintTag As Integer
    Select Case Control.ID
        Case ID_待打
            Control.Enabled = mDataState.bln待打
            Control.IconId = IIf(mbytFileState = 0, 90004, 90003)
        Case ID_重打
            Control.Enabled = mDataState.bln重打
            Control.IconId = IIf(mbytFileState = 1, 90004, 90003)
        Case ID_所有
            Control.Enabled = mDataState.bln所有
            Control.IconId = IIf(mbytFileState = 2, 90004, 90003)
        Case conMenu_View_Refresh
            Control.Visible = False
        Case conMenu_Edit_SelAll * 100# + 1, conMenu_Edit_SelAll * 100# + 2, conMenu_Edit_SelAll * 100# + 3, conMenu_Edit_SelAll * 100# + 4
            If Control.ID = conMenu_Edit_SelAll * 100# + 1 Then
                intPrintTag = 0
            ElseIf Control.ID = conMenu_Edit_SelAll * 100# + 2 Then
                intPrintTag = -1
            ElseIf Control.ID = conMenu_Edit_SelAll * 100# + 3 Then
                intPrintTag = 1
            Else
                intPrintTag = 2
            End If
            Control.Enabled = RevfgPrint(2, intPrintTag)
        Case conMenu_Edit_SelAll * 100# + 5, conMenu_Edit_SelAll * 100# + 6
            Control.Enabled = mDataState.bln所有
    End Select
End Sub

Private Sub cmdClear_Click()
    Dim arrPage() As String
    On Error GoTo ErrHand
    
    If Not vfgPrint.Rows > vfgPrint.FixedRows Then Exit Sub
    If Trim(txtClear.Text) = "" Then Exit Sub
    If Not IsNumeric(txtClear.Text) Then
        MsgBox "起始页号含有非法字符,请检查！", vbInformation, gstrSysName
        If txtClear.Enabled And txtClear.Visible Then txtClear.SetFocus
        Exit Sub
    End If
    If txtClear.Tag = "" Then txtClear.Tag = "0-0"
    arrPage = Split(txtClear.Tag, "-")
    If Not (Val(txtClear.Text) >= Val(arrPage(0)) And Val(txtClear.Text) <= Val(arrPage(1))) Then
        MsgBox "输入的页号不在有效页码范围内,请检查！", vbInformation, gstrSysName
        If txtClear.Enabled And txtClear.Visible Then txtClear.SetFocus
        Exit Sub
    End If
    
    Call zlDatabase.ExecuteProcedure("ZL_病人护理打印_CLEAR(0,0,0," & glng文件ID & "," & Val(txtClear.Text) & ")", "清除打印数据")
    MsgBox "清除成功！", vbInformation, gstrSysName
    Call zlRefresh(glng文件ID)
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdPrint_Click()
    If Not PrePrint Then Exit Sub
    mbytRunMode = 1
    Me.Hide
End Sub

Private Sub cmdEXCEL_Click()
    If Not PrePrint Then Exit Sub
    mbytRunMode = 3
    Me.Hide
End Sub

Private Sub cmdPreView_Click()
    If Not PrePrint Then Exit Sub
    mbytRunMode = 2
    Me.Hide
End Sub

Private Function PrePrint() As Boolean
    Dim lngRow As Long, lngPreRow As Long
    Dim strTmp As String, strTag As String
    Dim blnTrue As Boolean
    
    mstrPrintPages = ""
    With vfgPrint
        blnTrue = True
        lngPreRow = -1
        For lngRow = .FixedRows To .Rows - 1
            If Val(.TextMatrix(lngRow, .ColIndex("选择"))) <> 0 Then
                mstrPrintPages = mstrPrintPages & "," & Val(.TextMatrix(lngRow, .ColIndex("页码"))) & ";" & IIf(.RowData(lngRow) = -1, IIf(.TextMatrix(lngRow, .ColIndex("状态")) = "待续打", 1, 2), 2)
                If .RowData(lngRow) = -1 Then
                    If .TextMatrix(lngRow, .ColIndex("状态")) = "待续打" Then
                        strTag = "续打"
                    Else
                        strTag = "重打"
                    End If
                    strTmp = strTmp & vbCrLf & "第[" & Val(.TextMatrix(lngRow, .ColIndex("页码"))) & "]页【" & strTag & "】"
                End If
                If lngPreRow > -1 And lngPreRow + 1 <> lngRow And blnTrue = True Then blnTrue = False
                lngPreRow = lngRow
            End If
        Next lngRow
    End With
    '选择奇偶打印页码必须连续
    If chkPrintSet(2).Value <> 0 And blnTrue = False Then
        If MsgBox("本次打印勾选了参数【奇偶打印】，但您选择的页码不连续，将不能使用奇偶打印，请问您是否继续？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            Exit Function
        End If
    End If
        
    If strTmp <> "" Then
        If MsgBox("由于本次打印包含了之前未满页但已打印的页,请您对打印状态进行核对：" & strTmp & vbCrLf & _
            "请问您是否继续？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            Exit Function
        End If
    End If
    mstrPrintPages = Mid(mstrPrintPages, 2)
    
    If mstrPrintPages = "" Then
        MsgBox "请您选择页码！", vbInformation, gstrSysName
        If vfgPrint.Enabled And vfgPrint.Visible Then vfgPrint.SetFocus
        Exit Function
    End If
    Call SaveParam ' 保存参数设置
    PrePrint = True
End Function

Private Sub Form_Load()
    Dim objTool As CommandBar
    Dim objControl As CommandBarControl
    Dim objChildControl As CommandBarControl
    On Error GoTo ErrHand
    '初始化菜单
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    With cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize False, 24, 24
        .SetIconSize True, 16, 16
    End With
    cbsMain.VisualTheme = xtpThemeOffice2003
    cbsMain.EnableCustomization False
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    cbsMain.ActiveMenuBar.Visible = False
    
    '工具栏定义
    '-----------------------------------------------------
    Set objTool = cbsMain.Add("工具栏", xtpBarTop)      '固有
    objTool.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    objTool.ModifyStyle XTP_CBRS_GRIPPER, 0
    'objTool.Closeable = False
    With objTool.Controls
        .Add xtpControlLabel, conMenu_View_Show, "显示："
        Set objControl = .Add(xtpControlButton, ID_待打, "待打"): objControl.Style = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, ID_重打, "重打"):   objControl.Style = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, ID_所有, "所有"):   objControl.Style = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新"):   objControl.Style = xtpButtonIconAndCaption
        objControl.Flags = xtpFlagRightAlign
        Set objControl = .Add(xtpControlButtonPopup, conMenu_Edit_SelAll, "快速选择↓")
        objControl.Flags = xtpFlagRightAlign: objControl.BeginGroup = True: objControl.Style = xtpButtonIconAndCaption
        With objControl.CommandBar.Controls
            Set objChildControl = .Add(xtpControlButton, conMenu_Edit_SelAll * 100# + 1, "待打印(&1)"): objChildControl.ToolTipText = "选择所有未打的页": objChildControl.Style = xtpButtonCaption
            Set objChildControl = .Add(xtpControlButton, conMenu_Edit_SelAll * 100# + 2, "待续打(&2)"): objChildControl.ToolTipText = "选择所有续打的页": objChildControl.Style = xtpButtonCaption
            Set objChildControl = .Add(xtpControlButton, conMenu_Edit_SelAll * 100# + 3, "已打印(&3)"): objChildControl.ToolTipText = "选择所有已打的页": objChildControl.Style = xtpButtonCaption
            Set objChildControl = .Add(xtpControlButton, conMenu_Edit_SelAll * 100# + 4, "待重打(&4)"): objChildControl.ToolTipText = "选择所有重打的页": objChildControl.Style = xtpButtonCaption
            Set objChildControl = .Add(xtpControlButton, conMenu_Edit_SelAll * 100# + 5, "全选(&A)"): objChildControl.ToolTipText = "选择所有页": objChildControl.Style = xtpButtonCaption: objChildControl.BeginGroup = True
            Set objChildControl = .Add(xtpControlButton, conMenu_Edit_SelAll * 100# + 6, "全清(&C)"): objChildControl.ToolTipText = "清除对所有页的选择": objChildControl.Style = xtpButtonCaption
        End With
    End With
    cbsMain.KeyBindings.Add 0, VK_F5, conMenu_View_Refresh
    Call zlRefresh(glng文件ID)
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function zlRefresh(ByVal lngFileID As Long) As Boolean
    Dim rsData As New ADODB.Recordset
    Dim lng开始页号 As Long, lngPati As Long, lngPage As Long, lngBaby As Long
    Dim strSQLNew As String
    On Error GoTo ErrHand
    mintPageRows = 0
    '读取文件信息
    mstrSQL = " Select 病人ID,主页ID,NVL(婴儿,0) 婴儿,文件名称 From 病人护理文件 Where ID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(mstrSQL, "读取所有护理文件", lngFileID)
    Me.Caption = NVL(rsData!文件名称)
    lngPati = rsData!病人ID
    lngPage = rsData!主页ID
    lngBaby = rsData!婴儿
    '读取该文件的开始页码
    mstrSQL = "Select Min(开始页号) 开始页号 From 病人护理打印 Where 文件ID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(mstrSQL, "读取该文件的开始页码", lngFileID)
    lng开始页号 = Val(NVL(rsData!开始页号, 0))
    
    '读取文件数据行
    mstrSQL = " Select  d.内容文本" & vbNewLine & _
             " From 病历文件结构 d, 病历文件结构 p,病人护理文件 c" & vbNewLine & _
             " Where p.Id = d.父id And p.文件id = c.格式ID and C.ID=[1] And p.对象类型 = 1 And p.内容文本 = '表格样式' and d.要素名称='有效数据行'"
    Set rsData = zlDatabase.OpenSQLRecord(mstrSQL, "读取最大数据行", lngFileID)
    If rsData.RecordCount <> 0 Then
        mintPageRows = NVL(rsData!内容文本, 0)
    End If
    mbytFileState = 0
    mDataState.bln待打 = False: mDataState.bln重打 = False: mDataState.bln所有 = False
    '读取打印情况列表
    If lng开始页号 = 0 Then
        strSQLNew = ""
    Else
        strSQLNew = _
        "       Union" & vbNewLine & _
        "       Select 开始页号, 打印页号, 数据行, 打印标识" & vbNewLine & _
        "       From (With 病人护理文件_F1 As (Select Id, 续打id From 病人护理文件 Where 病人id = [2] And 主页id = [3] And Nvl(婴儿, 0) = [4])" & vbNewLine & _
        "               Select 结束页号 开始页号, Decode(打印结束页号, 结束页号, 结束页号, Null) 打印页号, 结束行号 数据行," & vbNewLine & _
        "                   Decode(打印结束页号, 结束页号, Decode(打印人, Null, 2, 1), 0) 打印标识" & vbNewLine & _
        "               From 病人护理打印 a, (Select Id From 病人护理文件_F1 Start With 续打id = [1] Connect By Prior Id = 续打id) b" & vbNewLine & _
        "               Where a.文件id = b.Id And a.结束页号 = [5])"
    End If
    mstrSQL = "Select 开始页号, Max(打印页号) 打印页号,Max(数据行) 数据行, Decode(Max(打印标识), 1, Decode(Min(打印标识), 0, -1, 1), Max(打印标识)) 打印标识" & vbNewLine & _
        " From (Select 开始页号, 打印页号,开始行号+行数-1 数据行," & vbNewLine & _
        "              Decode(打印页号," & vbNewLine & _
        "                      Null," & vbNewLine & _
        "                      0," & vbNewLine & _
        "                      Decode(打印页号, 开始页号, Decode(打印行号, 开始行号, Decode(打印人, Null, 2, 1), 2), 2)) 打印标识" & vbNewLine & _
        "       From 病人护理打印" & vbNewLine & _
        "       Where 文件id = [1]" & vbNewLine & _
        "       Union" & vbNewLine & _
        "       Select 结束页号 开始页号, Decode(打印结束页号, 结束页号, 结束页号, Null) 打印页号,结束行号 数据行," & vbNewLine & _
        "              Decode(打印结束页号, 结束页号, Decode(打印人, Null, 2, 1), 0) 打印标识" & vbNewLine & _
        "       From 病人护理打印" & vbNewLine & _
        "       Where 文件id = [1] And 结束页号 > 开始页号" & vbNewLine & _
        "       " & strSQLNew & ")" & vbNewLine & _
        " Group By 开始页号" & vbNewLine & _
        " Order By 开始页号"
    Set mrsData = zlDatabase.OpenSQLRecord(mstrSQL, "提取打印信息", lngFileID, lngPati, lngPage, lngBaby, lng开始页号)
    mrsData.Filter = ""
    mDataState.bln所有 = mrsData.RecordCount > 0
    mrsData.Filter = "打印标识=-1 OR 打印标识=0"
    mDataState.bln待打 = mrsData.RecordCount > 0
    mrsData.Filter = "打印标识=1 OR 打印标识=2"
    mDataState.bln重打 = mrsData.RecordCount > 0
    
    If mDataState.bln待打 = True Then
        mbytFileState = 0
    ElseIf mDataState.bln重打 = True Then
        mbytFileState = 1
    Else
        mbytFileState = 2
    End If
    cmdPreView.Enabled = mDataState.bln所有
    cmdPrint.Enabled = cmdPreView.Enabled
    cmdEXCEL.Enabled = cmdPreView.Enabled
    
    txtClear.Tag = "0-0"
    If mDataState.bln所有 = False Then
        lblTag(0).Caption = "该文件还未录入数据"
        cmdClear.Enabled = False
    Else
        mrsData.Filter = ""
        txtClear.Tag = Val(NVL(mrsData!开始页号))
        lblTag(0).Caption = "有效页码范围:第" & mrsData!开始页号 & "页"
        mrsData.MoveLast
        txtClear.Tag = txtClear.Tag & "-" & Val(NVL(mrsData!开始页号))
        lblTag(0).Caption = lblTag(0).Caption & " ～ 第" & mrsData!开始页号 & "页"
        cmdClear.Enabled = True
    End If
    lblTag(0).Left = (fraPrint(0).Width - lblTag(0).Width) \ 2
    If lblTag(0).Left < 0 Then lblTag(0).Left = 0
    
    Call LoadPrintData '显示打印列表数据
    Call LoadParam '加载参数
    
    zlRefresh = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function LoadPrintData() As Boolean
'功能:加载文件信息以及打印情况信息
    Dim lngRow As Long
    Dim strTag As String
    Dim stdPic As StdPicture
    
    On Error GoTo ErrHand
    
    Select Case mbytFileState
        Case 0
            mrsData.Filter = "打印标识=-1 OR 打印标识=0"
        Case 1
            mrsData.Filter = "打印标识=1 OR 打印标识=2"
        Case 2
            mrsData.Filter = ""
    End Select
    
    With vfgPrint
        .FixedRows = 1
        .FixedCols = 1
        .Rows = .FixedRows
        .Cols = 7
        .Editable = flexEDKbdMouse
        .MergeCells = flexMergeFixedOnly
        .MergeCellsFixed = flexMergeRestrictColumns
        .MergeRow(0) = True
        .MergeCol(.ColIndex("图片")) = True
        .MergeCol(.ColIndex("选择")) = True
        .ColHidden(.ColIndex("是否满页")) = True
         Set .Cell(flexcpPicture, 0, 0, .Rows - 1, .Cols - 1) = Nothing
        Do While Not mrsData.EOF
            Select Case Val(mrsData!打印标识)
                Case -1 '续打页
                    strTag = "待续打"
                    Set stdPic = imgData.ListImages("续打").Picture
                Case 0  '未打页
                    strTag = "待打印"
                    Set stdPic = imgData.ListImages("未打").Picture
                Case 1 '已打页
                    strTag = "已打印"
                    Set stdPic = imgData.ListImages("已打").Picture
                Case 2 '重打页
                    strTag = "待重打"
                    Set stdPic = imgData.ListImages("重打").Picture
            End Select
            If mrsData.AbsolutePosition + .FixedRows > .Rows Then .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("选择")) = 0
            .TextMatrix(.Rows - 1, .ColIndex("页码")) = CStr(mrsData!开始页号)
            .TextMatrix(.Rows - 1, .ColIndex("状态")) = strTag
            .TextMatrix(.Rows - 1, .ColIndex("是否满页")) = IIf(Val(NVL(mrsData!数据行, 0)) >= mintPageRows, "是", "否")
            .TextMatrix(.Rows - 1, .ColIndex("打印页码")) = CStr(NVL(mrsData!打印页号))
            .RowData(.Rows - 1) = Val(mrsData!打印标识)
            Set .Cell(flexcpPicture, .Rows - 1, 1) = stdPic
        mrsData.MoveNext
        Loop
        
        If .FixedRows < .Rows Then
            .Cell(flexcpBackColor, .FixedRows, .ColIndex("选择"), .Rows - 1, .Cols - 1) = &H80000005
            .RowSel = .FixedRows
            .ColSel = .ColIndex("选择")
        End If
        If .Enabled = True And .Visible = True Then .SetFocus
    End With
    
    LoadPrintData = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub AdjustRowFlag(ByRef objVsf As Object, ByVal intRow As Integer)
    '-----------------------------------------------------------------------------------------
    '功能:
    '参数:
    '-----------------------------------------------------------------------------------------
    If objVsf.FixedCols = 0 Then Exit Sub
    If objVsf.Rows <= objVsf.FixedRows Then Exit Sub
    If Not (objVsf.Cell(flexcpPicture, intRow, 0) Is Nothing) Then Exit Sub
    Set objVsf.Cell(flexcpPicture, 0, 0, objVsf.Rows - 1, 0) = Nothing
    Set objVsf.Cell(flexcpPicture, intRow, 0) = ils16.ListImages(1).Picture
End Sub

Private Function RevfgPrint(ByVal byt场合 As Byte, Optional ByVal intPrintTag As Integer = 0) As Boolean
'byt场合:
'       1:根据intPrintTag进行某项标识的选择
'       2:根据intPrintTag判断对应的标识是否存在,返回TRUE OR False
'       3:全选
'       4:全清
'intPrintTag:打印标识 -1 -续打；0 -未打;1 -已打;2 -重打   byt场合=3,4不传
    Dim lngRow As Long
    Dim blnOK As Boolean
    
    blnOK = False
    With vfgPrint
        For lngRow = .FixedRows To .Rows - 1
            Select Case byt场合
                Case 1
                    If Val(.RowData(lngRow)) = intPrintTag Then
                        .TextMatrix(lngRow, .ColIndex("选择")) = 1
                        .Cell(flexcpBackColor, lngRow, .ColIndex("选择"), lngRow, .Cols - 1) = RGB(135, 206, 235)
                    Else
                        .TextMatrix(lngRow, .ColIndex("选择")) = 0
                        .Cell(flexcpBackColor, lngRow, .ColIndex("选择"), lngRow, .Cols - 1) = &H80000005
                    End If
                Case 2
                    If Val(.RowData(lngRow)) = intPrintTag And .TextMatrix(lngRow, .ColIndex("页码")) <> "" Then
                        blnOK = True
                        Exit For
                    End If
                Case 3
                    .TextMatrix(lngRow, .ColIndex("选择")) = 1
                    .Cell(flexcpBackColor, lngRow, .ColIndex("选择"), lngRow, .Cols - 1) = RGB(135, 206, 235)
                Case 4
                    .TextMatrix(lngRow, .ColIndex("选择")) = 0
                    .Cell(flexcpBackColor, lngRow, .ColIndex("选择"), lngRow, .Cols - 1) = &H80000005
            End Select
            
        Next lngRow
        If byt场合 <> 2 Then
            If .RowSel >= .FixedRows And .Rows > .FixedRows Then
                .BackColorSel = .Cell(flexcpBackColor, .RowSel, .ColIndex("选择"), .RowSel, .Cols - 1)
            End If
        End If
    End With
    If byt场合 = 2 Then
        RevfgPrint = blnOK
    Else
        RevfgPrint = True
    End If
End Function
 
            
Private Sub Form_Unload(Cancel As Integer)
    Call SaveParam
    mbytRunMode = 0
End Sub


Private Sub SaveParam()
'保存打印相关参数
    '--56134:刘鹏飞,2012-12-19,记录单打印时,数据未满页空白部分输出表格
    Call zlDatabase.SetPara("记录单未满页打印表格", chkPrintSet(0).Value, glngSys, 1255)
    '--46506:刘鹏飞,2012-12-19,记录单打印时，数据满页才进行输出(文件为结束时有效)
    Call zlDatabase.SetPara("记录单满页打印", chkPrintSet(1).Value, glngSys, 1255)
    '--49753:刘鹏飞,2012-12-19,记录单打印时，数据页奇偶输出
    Call zlDatabase.SetPara("记录单奇偶打印", chkPrintSet(2).Value, glngSys, 1255)
End Sub

Private Sub LoadParam()
'加载打印相关参数
    '--56134:刘鹏飞,2012-12-19,记录单打印时,数据未满页空白部分输出表格
    chkPrintSet(0).Value = Val(zlDatabase.GetPara("记录单未满页打印表格", glngSys, 1255, "0", Array(chkPrintSet(0)), True))
    '--46506:刘鹏飞,2012-12-19,记录单打印时，数据满页才进行输出(文件为结束时有效)
    chkPrintSet(1).Value = Val(zlDatabase.GetPara("记录单满页打印", glngSys, 1255, "0", Array(chkPrintSet(1)), True))
    '--49753:刘鹏飞,2012-12-19,记录单打印时，数据页奇偶输出
    chkPrintSet(2).Value = Val(zlDatabase.GetPara("记录单奇偶打印", glngSys, 1255, "0", Array(chkPrintSet(2)), True))
End Sub

Private Sub Label1_Click(Index As Integer)

End Sub

Private Sub txtClear_KeyPress(KeyAscii As Integer)
    Call zlControl.TxtCheckKeyPress(txtClear, KeyAscii, m数字式)
End Sub

Private Sub vfgPrint_AfterEdit(ByVal ROW As Long, ByVal COL As Long)
    Dim intValue As Integer
    Dim lngRow As Long, lngStartRow As Long, lngEndRow As Long
    With vfgPrint
        If ROW < .FixedRows Then Exit Sub
        If COL = .ColIndex("选择") Then
            intValue = Val(.TextMatrix(ROW, COL))
            .Cell(flexcpBackColor, ROW, COL, ROW, .Cols - 1) = IIf(intValue = 0, &H80000005, RGB(135, 206, 235))
            .BackColorSel = IIf(intValue = 0, &H80000005, RGB(135, 206, 235))
            '判断Shift键是否按起，如果按起则进行批量选(类似windows文件选择)
            If (GetAsyncKeyState(vbKeyShift) And &H8000) = &H8000 And intValue <> 0 Then
                lngStartRow = -1
                lngEndRow = -1
                For lngRow = ROW - 1 To .FixedRows Step -1
                    If Val(.TextMatrix(lngRow, COL)) <> 0 Then
                        lngStartRow = lngRow
                        lngEndRow = ROW
                        Exit For
                    End If
                Next lngRow
                If lngStartRow = -1 Then
                    For lngRow = ROW + 1 To .Rows - 1
                        If Val(.TextMatrix(lngRow, COL)) <> 0 Then
                            lngStartRow = ROW
                            lngEndRow = lngRow
                            Exit For
                        End If
                    Next lngRow
                End If
                If lngStartRow < lngEndRow Then
                    For lngRow = lngStartRow To lngEndRow
                        .TextMatrix(lngRow, COL) = 1
                        .Cell(flexcpBackColor, lngRow, COL, lngRow, .Cols - 1) = RGB(135, 206, 235)
                    Next
                End If
            End If
        ElseIf COL = .ColIndex("状态") Then
            If .TextMatrix(ROW, COL) = "" Then .TextMatrix(ROW, COL) = "待续打"
        End If
    End With
End Sub

Private Sub vfgPrint_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim blnCancle As Boolean
    vfgPrint.ColComboList(NewCol) = ""
    
    Call AdjustRowFlag(vfgPrint, NewRow)
    Call vfgPrint_StartEdit(NewRow, NewCol, blnCancle)
    If blnCancle = False And vfgPrint.ColIndex("状态") = NewCol Then
        vfgPrint.ColComboList(NewCol) = "待续打|待重打"
    End If
End Sub

Private Sub vfgPrint_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
    vfgPrint.BackColorSel = vfgPrint.Cell(flexcpBackColor, NewRowSel, vfgPrint.ColIndex("选择"), NewRowSel, vfgPrint.Cols - 1)
End Sub

Private Sub vfgPrint_DblClick()
    Call vfgPrint_KeyPress(vbKeySpace)
End Sub

Private Sub vfgPrint_KeyPress(KeyAscii As Integer)
    Dim intValue As Integer
    If KeyAscii = vbKeySpace And vfgPrint.ROW >= vfgPrint.FixedRows And vfgPrint.ROW < vfgPrint.Rows And vfgPrint.COL <> vfgPrint.ColIndex("选择") Then
        intValue = Val(vfgPrint.TextMatrix(vfgPrint.ROW, vfgPrint.ColIndex("选择")))
        vfgPrint.TextMatrix(vfgPrint.ROW, vfgPrint.ColIndex("选择")) = IIf(intValue = 0, 1, 0)
        intValue = Val(vfgPrint.TextMatrix(vfgPrint.ROW, vfgPrint.ColIndex("选择")))
        vfgPrint.Cell(flexcpBackColor, vfgPrint.ROW, vfgPrint.ColIndex("选择"), vfgPrint.ROW, vfgPrint.Cols - 1) = IIf(intValue = 0, &H80000005, RGB(135, 206, 235))
        vfgPrint.BackColorSel = IIf(intValue = 0, &H80000005, RGB(135, 206, 235))
    End If
End Sub

Private Sub vfgPrint_StartEdit(ByVal ROW As Long, ByVal COL As Long, Cancel As Boolean)
    With vfgPrint
        If ROW >= .FixedRows Then
            If COL = .ColIndex("选择") Then
                Cancel = False
            ElseIf COL = .ColIndex("状态") And .RowData(ROW) = -1 Then
                Cancel = False
            Else
                Cancel = True
            End If
        Else
            Cancel = True
        End If
    End With
End Sub
