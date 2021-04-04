VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "CODEJO~1.OCX"
Begin VB.Form frmRunLimitManage 
   BackColor       =   &H80000005&
   Caption         =   "功能限时管理"
   ClientHeight    =   8655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12315
   ControlBox      =   0   'False
   Icon            =   "frmRunLimitManage.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "form3"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Picture         =   "frmRunLimitManage.frx":6852
   ScaleHeight     =   8655
   ScaleWidth      =   12315
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picTop 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   0
      ScaleHeight     =   3015
      ScaleWidth      =   12255
      TabIndex        =   4
      Top             =   1290
      Width           =   12255
      Begin VSFlex8Ctl.VSFlexGrid vsfModuleList 
         Height          =   2175
         Left            =   45
         TabIndex        =   9
         Top             =   0
         Width           =   11955
         _cx             =   21087
         _cy             =   3836
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483634
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   260
         RowHeightMax    =   260
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmRunLimitManage.frx":6D4B
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   3
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
         ExplorerBar     =   1
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
   Begin VB.CommandButton cmdPlanSet 
      Caption         =   "方案管理(&S)"
      Height          =   350
      Left            =   10995
      TabIndex        =   8
      Top             =   690
      Width           =   1200
   End
   Begin MSComctlLib.ImageList imgPlanDetail 
      Left            =   11625
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   97
      ImageHeight     =   30
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRunLimitManage.frx":6E38
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picBottom 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   3885
      Left            =   0
      ScaleHeight     =   3885
      ScaleWidth      =   12255
      TabIndex        =   2
      Top             =   4590
      Width           =   12255
      Begin VB.PictureBox picPlanDetail 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   2415
         Left            =   0
         ScaleHeight     =   2415
         ScaleWidth      =   8130
         TabIndex        =   5
         Top             =   1035
         Width           =   8130
         Begin VSFlex8Ctl.VSFlexGrid vsfPlanDetail 
            Height          =   2115
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Width           =   7905
            _cx             =   13944
            _cy             =   3731
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   16774866
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   16774866
            GridColor       =   -2147483633
            GridColorFixed  =   15984570
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483633
            FocusRect       =   1
            HighLight       =   0
            AllowSelection  =   0   'False
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   0
            GridLineWidth   =   1
            Rows            =   3
            Cols            =   8
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmRunLimitManage.frx":B10E
            ScrollTrack     =   0   'False
            ScrollBars      =   1
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
      Begin XtremeSuiteControls.TabControl tbcPage 
         Height          =   960
         Left            =   45
         TabIndex        =   3
         Top             =   0
         Width           =   1755
         _Version        =   589884
         _ExtentX        =   3096
         _ExtentY        =   1693
         _StockProps     =   64
      End
   End
   Begin VB.Frame frmTopMidSplit 
      Height          =   45
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   12360
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   11010
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRunLimitManage.frx":B200
            Key             =   "system"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRunLimitManage.frx":11A62
            Key             =   "function"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRunLimitManage.frx":182C4
            Key             =   "program"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picTopBottom 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   45
      MousePointer    =   7  'Size N S
      ScaleHeight     =   135
      ScaleWidth      =   4515
      TabIndex        =   10
      Top             =   4365
      Width           =   4515
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   $"frmRunLimitManage.frx":1EB26
      Height          =   360
      Left            =   1050
      TabIndex        =   7
      Top             =   675
      Width           =   7740
   End
   Begin VB.Image imgDescription 
      Height          =   720
      Left            =   150
      Picture         =   "frmRunLimitManage.frx":1EBD8
      Top             =   495
      Width           =   720
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "功能限时管理"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   195
      TabIndex        =   0
      Top             =   135
      Width           =   1440
   End
End
Attribute VB_Name = "frmRunLimitManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrLastPlan As String '记录上一次选择的方案名称
Private mrsPlanList As ADODB.Recordset
Private Const vsfTitleBackColor = &HF0E5BD  '方案内容表格标题背景颜色
Private Const vsfContentBackColor = &HFFFAE4 '方案内容表格内容部分浅色背景色
Private Const vsfTitleHeight = 500
Private Const vsfRowHeight = 1000
Private Enum ModuleList
    ML_序号 = 0
    ML_系统 = 1
    ML_模块 = 2
    ML_功能 = 3
    ML_限时方案 = 4
    ML_方案说明 = 5
    ML_操作选项 = 6
    ML_限时原因 = 7
End Enum
Private Enum PlanDetailTitle
    PDT_星期 = 0
    PDT_时间段1 = 1
    PDT_时间段扩展 = 2
End Enum
Private Enum PlanDetail
    PD_标题 = 0
    PD_星期日 = 1
    PD_星期一 = 2
    PD_星期二 = 3
    PD_星期三 = 4
    PD_星期四 = 5
    PD_星期五 = 6
    PD_星期六 = 7
End Enum

Public Function SupportPrint() As Boolean
'返回本窗口是否支持打印，供主窗口调用
    SupportPrint = False
End Function

Private Sub cmdPlanSet_Click()
    Call frmRunLimitPlanManage.ShowMe(vsfPlanDetail.Tag)
    Call FillModuleData(vsfModuleList.Row)
End Sub

Private Function SaveData(ByVal lngRow As Long) As Boolean
'保存模块的方案选择，操作选项以及限时原因等信息
    On Error GoTo errH
    If InStr(vsfModuleList.TextMatrix(lngRow, ML_限时原因), "'") > 0 Then
        MsgBox "“限时原因”中含有单引号，请重新填写！", vbInformation, gstrSysName
        Exit Function
    ElseIf LenB(StrConv(vsfModuleList.TextMatrix(lngRow, ML_限时原因), vbFromUnicode)) > 250 Then
        MsgBox "“限时原因”内容不能操作125个汉字或250个字符，请重新填写！"
        Exit Function
    Else
        Call ExecuteProcedure("Zl_ZlRunLimitSet_Update(" & vsfModuleList.TextMatrix(lngRow, ML_序号) & _
                            "," & Val(vsfPlanDetail.Tag) & "," & IIf(vsfModuleList.TextMatrix(lngRow, ML_操作选项) = "禁止", 0, 1) & _
                            ",'" & vsfModuleList.TextMatrix(lngRow, ML_限时原因) & "')", "保存模块方案信息")
    End If
    vsfModuleList.Tag = vsfModuleList.TextMatrix(vsfModuleList.RowSel, ML_限时方案) & "_" & _
                        vsfModuleList.TextMatrix(vsfModuleList.RowSel, ML_操作选项) & "_" & _
                        vsfModuleList.TextMatrix(vsfModuleList.RowSel, ML_限时原因)
    SaveData = True
    Exit Function
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    '禁止输入单引号
    If InStr("'", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    picTopBottom.Top = Val(GetSetting("ZLSOFT", "公共模块\服务器管理工具\功能限时管理", "picTopBottom_Top", "5000"))
    '对tabControl控件进行初始化
    Call InitTabControl
    '初始化表格格式
    Call FormatPlanDetail
    '填充数据
    Call FillModuleData
End Sub

'==============================================================================
'=功能： 初始Tab控件
'==============================================================================
Private Function InitTabControl() As Boolean
    Dim objTabItem As TabControlItem
    
    On Error GoTo errH
    With tbcPage
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .BoldSelected = True
            .ClientFrame = xtpTabFrameSingleLine
            .ShowIcons = True
            .OneNoteColors = True
            .DisableLunaColors = True
        End With
        '第一页
        Set objTabItem = .InsertItem(0, "预设方案", picPlanDetail.hwnd, 0)
    End With

    InitTabControl = True

    Exit Function
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    err.Clear
End Function

Private Sub FormatPlanDetail()
    '设置下方方案展示表格格式
    With vsfPlanDetail
        .Cell(flexcpPicture, 0, 0) = imgPlanDetail.ListImages(1).Picture
        .GridLines = flexGridNone
        .rowHeight(PD_标题) = vsfTitleHeight
        .rowHeight(PDT_时间段1) = vsfRowHeight
    End With
End Sub

'填充上方模块功能列表中的方案及操作选项下拉框
Private Sub FormatPlanList()
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strComboList As String
    
    On Error GoTo errH
    strSql = "Select 名称, 描述 From Zlrunlimit Where 是否启用 = 1 Order by 序号"
    Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSql, Me.Caption)
    With rsTemp
        Do While Not .EOF
            strComboList = strComboList & "|" & !名称
            .MoveNext
        Loop
    End With
    vsfModuleList.ColComboList(ML_限时方案) = "[无方案设置]|" & strComboList
    vsfModuleList.ColComboList(ML_操作选项) = "提醒|禁止"
    Exit Sub
errH:
    MsgBox err.Description, vbInformation
End Sub

'填充上方模块功能及其限时信息
Private Sub FillModuleData(Optional ByVal lngRow As Long)
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Long
    
    On Error GoTo errH
    Call FormatPlanList
    strSql = "Select 0 系统, a.模块, b.标题 模块名称, a.序号, a.功能, a.操作选项, c.名称 方案, '服务器管理工具' 系统名称, a.限时原因" & vbNewLine & _
            "From Zlrunlimitset A, zlSvrTools B, Zlrunlimit C" & vbNewLine & _
            "Where a.模块 = b.编号 And a.系统 Is Null And a.方案序号 = c.序号(+)" & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select a.系统, a.模块, c.标题 模块名称, a.序号, a.功能, a.操作选项, d.名称 方案, b.名称 系统名称, a.限时原因" & vbNewLine & _
            "From Zlrunlimitset A, zlSystems B, zlPrograms C, Zlrunlimit D" & vbNewLine & _
            "Where a.系统 = b.编号 And a.模块 = c.序号 And a.方案序号 = d.序号(+)"
    Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSql, Me.Caption)
    '填充模块功能列表
    rsTemp.Sort = "系统,模块,序号"
    With rsTemp
        vsfModuleList.Rows = .RecordCount + 1
        For i = 1 To .RecordCount
            vsfModuleList.TextMatrix(i, ML_序号) = !序号
            vsfModuleList.TextMatrix(i, ML_系统) = !系统名称
            vsfModuleList.TextMatrix(i, ML_模块) = !模块名称
            vsfModuleList.TextMatrix(i, ML_功能) = !功能
            vsfModuleList.TextMatrix(i, ML_限时方案) = Nvl(!方案, "[无方案设置]")
            vsfModuleList.TextMatrix(i, ML_操作选项) = IIf(!操作选项 = 0, "禁止", "提醒")
            vsfModuleList.TextMatrix(i, ML_限时原因) = !限时原因
            .MoveNext
        Next
        If .RecordCount > 0 Then
            vsfModuleList.MergeCol(ML_系统) = True
            vsfModuleList.MergeCol(ML_模块) = True
            If lngRow = 0 Then
                vsfModuleList.Row = 1
            Else
                vsfModuleList.Row = lngRow
            End If
            Call vsfModuleList_Click
        End If
    End With
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    cmdPlanSet.Left = Me.ScaleWidth - cmdPlanSet.Width - 200
    frmTopMidSplit.Width = Me.ScaleWidth
    picTop.Height = picTopBottom.Top - picTop.Top
    picTop.Width = Me.ScaleWidth - 60
    picTopBottom.Width = picTop.Width - 45
    picBottom.Top = picTop.Top + picTop.Height + 60
    picBottom.Width = picTop.Width
    picBottom.Height = Me.ScaleHeight - picBottom.Top
    If err.Number <> 0 Then err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "ZLSOFT", "公共模块\服务器管理工具\功能限时管理", "picTopBottom_Top", picTopBottom.Top
    mstrLastPlan = ""
End Sub

Private Sub picTopBottom_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If picTopBottom.Top >= 9000 And Y > 0 Then Exit Sub
        If picTopBottom.Top <= 5000 And Y < 0 Then Exit Sub
        picTopBottom.Top = picTopBottom.Top + Y
        picTop.Height = picTopBottom.Top - picTop.Top
        picBottom.Top = picTop.Top + picTop.Height + 60
        picBottom.Height = Me.ScaleHeight - picBottom.Top
    End If
End Sub

Private Sub picTop_Resize()
    On Error Resume Next
    vsfModuleList.Width = picTop.Width
    vsfModuleList.Height = picTop.Height
    If err.Number <> 0 Then err.Clear
End Sub

Private Sub picPlanDetail_Resize()
    On Error Resume Next
    vsfPlanDetail.Width = picPlanDetail.Width
    vsfPlanDetail.Height = picPlanDetail.Height
    Call AdjustFormDisplay
    If err.Number <> 0 Then err.Clear
End Sub

Private Sub AdjustFormDisplay()
    With vsfPlanDetail
        .Select 0, 0, .Rows - 1, .Cols - 1
        .CellBorder &HE9D2A5, 1, 2, 1, 0, 2, 2
        .Cell(flexcpBackColor, PDT_星期, PD_标题, .Rows - 1, 0) = vsfTitleBackColor
        .Cell(flexcpBackColor, PDT_星期, PD_标题, 0, .Cols - 1) = vsfTitleBackColor
        .Cell(flexcpBackColor, PDT_时间段1, PD_星期日, .Rows - 1, PD_星期日) = vsfContentBackColor
        .Cell(flexcpBackColor, PDT_时间段1, PD_星期二, .Rows - 1, PD_星期二) = vsfContentBackColor
        .Cell(flexcpBackColor, PDT_时间段1, PD_星期四, .Rows - 1, PD_星期四) = vsfContentBackColor
        .Cell(flexcpBackColor, PDT_时间段1, PD_星期六, .Rows - 1, PD_星期六) = vsfContentBackColor
        .rowHeight(.Rows - 1) = picBottom.Height - (.Rows - 1) * .rowHeight(PDT_时间段1) + 200
    End With
End Sub

Private Sub picBottom_Resize()
    tbcPage.Width = picBottom.Width
    tbcPage.Height = picBottom.Height
End Sub

Private Sub FillPlanDetail(ByVal strPlanName As String)
'填充详细方案信息
'strPlanName = 方案名称
    Dim j As Long  '表示时间段
    Dim lngLastWeekNo As Long
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errH
        '将老的方案信息清空
        Call ClearPlanDetail
        If strPlanName = "[无方案设置]" Then
            tbcPage.Item(0).Caption = "无方案"
        Else
            tbcPage.Item(0).Caption = strPlanName
        End If
        mstrLastPlan = strPlanName
        '填充新方案
        strSql = "Select b.序号, a.星期, To_Char(a.开始时间, 'HH24:MI:SS') 开始时间, To_Char(a.结束时间, 'HH24:MI:SS') 结束时间, b.描述" & vbNewLine & _
                "From Zlrunlimittime A, Zlrunlimit B" & vbNewLine & _
                "Where a.方案 = b.序号 And b.名称 = [1]" & vbNewLine & _
                "Order By a.星期, a.开始时间"
        Set rsTemp = gclsBase.OpenSQLRecord(gcnOracle, strSql, Me.Caption, strPlanName)
        With rsTemp
            If .RecordCount > 0 Then
                vsfModuleList.TextMatrix(vsfModuleList.RowSel, ML_方案说明) = !描述 & ""
                vsfPlanDetail.Tag = !序号
            Else
                vsfModuleList.TextMatrix(vsfModuleList.RowSel, ML_方案说明) = ""
                vsfPlanDetail.Tag = 1
            End If
            Do While Not .EOF
                If !星期 = lngLastWeekNo Then
                    j = j + 1
                    If j + 2 > vsfPlanDetail.Rows Then
                        vsfPlanDetail.Rows = j + 2
                        vsfPlanDetail.rowHeight(j) = vsfPlanDetail.rowHeight(PDT_时间段1)
                        vsfPlanDetail.TextMatrix(j, 0) = "时间段" & j
                        vsfPlanDetail.ColAlignment(j) = flexAlignCenterCenter
                    End If
                Else
                    j = 1
                End If
                vsfPlanDetail.TextMatrix(j, !星期 + 1) = "起 " & !开始时间 & vbNewLine & vbNewLine & "止 " & !结束时间
                lngLastWeekNo = !星期
                .MoveNext
            Loop
            Call AdjustFormDisplay
        End With
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    If 0 = 1 Then
        Resume
    End If
End Sub

'将老的方案信息清空
Private Sub ClearPlanDetail()
    Dim i As Long
    
    With vsfPlanDetail
        .Rows = 3
        .TextMatrix(PDT_时间段扩展, 0) = ""
        For i = PD_星期日 To PD_星期六
            .TextMatrix(PDT_时间段1, i) = ""
            .TextMatrix(PDT_时间段扩展, i) = ""
        Next
        Call AdjustFormDisplay
    End With
End Sub


Private Sub vsfModuleList_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsfModuleList
        If .Tag <> .TextMatrix(Row, ML_限时方案) & "_" & .TextMatrix(Row, ML_操作选项) & "_" & .TextMatrix(Row, ML_限时原因) Then
            Call SaveData(Row)
        End If
    End With
End Sub

Private Sub vsfModuleList_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = ML_系统 Or Col = ML_模块 Or Col = ML_功能 Then Cancel = True
End Sub

'当点击某一个模块功能时，将在下方展示出该模块功能选择的限时方案的详细信息
Private Sub vsfModuleList_Click()
    With vsfModuleList
        .Tag = .TextMatrix(.RowSel, ML_限时方案) & "_" & .TextMatrix(.RowSel, ML_操作选项) & "_" & .TextMatrix(.RowSel, ML_限时原因)
        If mstrLastPlan = .TextMatrix(.RowSel, ML_限时方案) And .MouseRow = .Row Then Exit Sub
        Call FillPlanDetail(.TextMatrix(.RowSel, ML_限时方案))
    End With
End Sub

Private Sub vsfModuleList_ComboCloseUp(ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)
    If Col = ML_限时方案 Then
        If mstrLastPlan = vsfModuleList.EditText Then Exit Sub
        Call FillPlanDetail(vsfModuleList.EditText)
    End If
End Sub

Private Sub vsfModuleList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sinRight As Single
    Dim sinLeftPlan As Single, sinLeftReason As Single
    Dim strTip As String
    Dim lngRow As Long
    
    lngRow = Int(Y / 260)
    With vsfModuleList
        If lngRow > .Rows - 1 Or lngRow = 0 Then
            Call ShowTipInfo(.hwnd, "")
            Exit Sub
        End If
        sinLeftPlan = .ColWidth(ML_系统) + .ColWidth(ML_模块) + .ColWidth(ML_功能)
        sinRight = .ColWidth(ML_系统) + .ColWidth(ML_模块) + .ColWidth(ML_功能) + .ColWidth(ML_限时方案)
        sinLeftReason = .ColWidth(ML_系统) + .ColWidth(ML_模块) + .ColWidth(ML_功能) + .ColWidth(ML_限时方案) + .ColWidth(ML_操作选项)
        If X >= sinLeftPlan And X <= sinRight Then
            strTip = .TextMatrix(lngRow, ML_方案说明)
        ElseIf X > sinLeftReason Then
            strTip = .TextMatrix(lngRow, ML_限时原因)
        Else
            strTip = ""
        End If
        Call ShowTipInfo(.hwnd, strTip, True)
    End With
End Sub

Private Sub vsfPlanDetail_DblClick()
    Call cmdPlanSet_Click
End Sub
