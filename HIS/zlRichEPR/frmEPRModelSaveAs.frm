VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "*\A..\zlRichEditor\zlRichEdit.vbp"
Begin VB.Form frmEPRModelSaveAs 
   Caption         =   "另存为片段..."
   ClientHeight    =   7110
   ClientLeft      =   2835
   ClientTop       =   3825
   ClientWidth     =   10755
   Icon            =   "frmEPRModelSaveAs.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   10755
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   5865
      Index           =   2
      Left            =   6645
      ScaleHeight     =   5865
      ScaleWidth      =   3900
      TabIndex        =   4
      Top             =   300
      Width           =   3900
      Begin VB.Frame fra 
         Height          =   6015
         Index           =   0
         Left            =   0
         TabIndex        =   9
         Top             =   300
         Width           =   3900
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   0
            Left            =   840
            TabIndex        =   19
            Top             =   1470
            Width           =   2940
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   0
            Left            =   840
            TabIndex        =   18
            Top             =   435
            Width           =   2940
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   1
            Left            =   840
            TabIndex        =   17
            Top             =   780
            Width           =   2940
         End
         Begin VB.OptionButton opt范围 
            Caption         =   "&1)全院通用"
            Height          =   180
            Index           =   0
            Left            =   825
            TabIndex        =   16
            Top             =   1905
            Width           =   1215
         End
         Begin VB.OptionButton opt范围 
            Caption         =   "&2)科内通用"
            Height          =   180
            Index           =   1
            Left            =   825
            TabIndex        =   15
            Top             =   2205
            Width           =   1215
         End
         Begin VB.OptionButton opt范围 
            Caption         =   "&3)个人使用"
            Height          =   180
            Index           =   2
            Left            =   825
            TabIndex        =   14
            Top             =   2505
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   1
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   4305
            Width           =   2940
         End
         Begin VB.TextBox txt 
            Height          =   735
            Index           =   3
            Left            =   840
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   12
            Top             =   2775
            Width           =   2940
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   2
            Left            =   840
            MaxLength       =   10
            TabIndex        =   11
            Top             =   1125
            Width           =   2940
         End
         Begin VB.CheckBox chkAdd 
            Caption         =   "新增(&A)"
            Enabled         =   0   'False
            Height          =   240
            Left            =   855
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   150
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "分类(&F)"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   120
            TabIndex        =   27
            Top             =   1515
            Width           =   630
         End
         Begin VB.Label lbl编号 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "编号(&B)"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   120
            TabIndex        =   26
            Top             =   495
            Width           =   630
         End
         Begin VB.Label lbl名称 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "名称(&N)"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   120
            TabIndex        =   25
            Top             =   825
            Width           =   630
         End
         Begin VB.Label lbl范围 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "使用(&U)"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   120
            TabIndex        =   24
            Top             =   1920
            Width           =   630
         End
         Begin VB.Label lbl人员 
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   840
            TabIndex        =   23
            Top             =   4650
            Width           =   2940
         End
         Begin VB.Label lbl科室 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "制作(&R)"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   105
            TabIndex        =   22
            Top             =   4350
            Width           =   630
         End
         Begin VB.Label lbl说明 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "说明(&M)"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   105
            TabIndex        =   21
            Top             =   2820
            Width           =   630
         End
         Begin VB.Label lbl简码 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "简码(&D)"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   120
            TabIndex        =   20
            Top             =   1185
            Width           =   630
         End
      End
   End
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   1935
      Index           =   1
      Left            =   750
      ScaleHeight     =   1935
      ScaleWidth      =   3630
      TabIndex        =   3
      Top             =   3390
      Width           =   3630
      Begin VB.Frame fra 
         Height          =   1905
         Index           =   2
         Left            =   705
         TabIndex        =   7
         Top             =   45
         Width           =   6585
         Begin VSFlex8Ctl.VSFlexGrid vfgTerm 
            Height          =   750
            Left            =   285
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   180
            Width           =   4080
            _cx             =   7197
            _cy             =   1323
            Appearance      =   2
            BorderStyle     =   0
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
            BackColor       =   16777215
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            BackColorBkg    =   -2147483643
            BackColorAlternate=   16777215
            GridColor       =   -2147483643
            GridColorFixed  =   -2147483643
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   0
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   -1  'True
            AllowUserResizing=   2
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   4
            Cols            =   1
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   ""
            ScrollTrack     =   -1  'True
            ScrollBars      =   2
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
         End
      End
   End
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   3075
      Index           =   0
      Left            =   675
      ScaleHeight     =   3075
      ScaleWidth      =   3630
      TabIndex        =   2
      Top             =   105
      Width           =   3630
      Begin VB.Frame fra 
         Height          =   3885
         Index           =   1
         Left            =   315
         TabIndex        =   5
         Top             =   180
         Width           =   6795
         Begin XtremeReportControl.ReportControl rptList 
            Height          =   2970
            Left            =   75
            TabIndex        =   6
            Top             =   150
            Width           =   4515
            _Version        =   589884
            _ExtentX        =   7964
            _ExtentY        =   5239
            _StockProps     =   0
            BorderStyle     =   2
            MultipleSelection=   0   'False
            EditOnClick     =   0   'False
         End
      End
   End
   Begin zlRichEditor.Editor Editor1 
      Height          =   2985
      Left            =   10680
      TabIndex        =   0
      Top             =   4170
      Visible         =   0   'False
      Width           =   6480
      _extentx        =   11430
      _extenty        =   5265
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   5160
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEPRModelSaveAs.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEPRModelSaveAs.frx":0B24
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEPRModelSaveAs.frx":10BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEPRModelSaveAs.frx":1458
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEPRModelSaveAs.frx":1D32
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEPRModelSaveAs.frx":20CC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   6735
      Width           =   10755
      _ExtentX        =   18971
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmEPRModelSaveAs.frx":2466
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16087
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
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmEPRModelSaveAs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'（１）窗体级变量定义
'######################################################################################################################

Private Enum mLCol
    图标 = 0: 性质: ID: 科室ID: 人员ID: 分类: 编号: 名称: 简码: 说明: 部门: 人员
End Enum

Private mbytFromTab As Byte     '来源表类型:1-病历范文目录,2-电子病历记录
Private mlngFromId As Long      '来源记录id
Private mbytPower As Integer     '用户权限级别

Private mlngFileID As Long      '文件ID
Private mstrCompends As String  '提纲id
Private mlngDemoId As Long      '另存的示范id
Private mblnOK As Boolean

Private mlngSelfId As Long      '当前用户的人员id
Private mstrSelfName As String  '当前用户的人员姓名
Private mblnDataChanged As Boolean



'（２）自定义过程或函数
'######################################################################################################################
Public Function ShowMe(ByVal bytFromTab As Byte, ByVal lngFromId As Long, Optional strCompends As String) As Long
    '******************************************************************************************************************
    '功能：上级程序调用本窗体的，传递参数，并显示窗体
    '参数： bytFromTab-来源表类型,1-病历范文目录,2-电子病历记录
    '       lngFromId-来源的记录id
    '       strCompends-另存的提纲id串，未传递时表示另存为范文，否则另存片段
    '返回：确定返回新增或修改的ID；取消返回0
    '******************************************************************************************************************
    
    mbytFromTab = bytFromTab
    mlngFromId = lngFromId
    mstrCompends = Trim(strCompends)
    
    If ExecuteCommand("初始控件") = False Then Unload Me: Exit Function
    If ExecuteCommand("读注册表") = False Then Unload Me: Exit Function
    If ExecuteCommand("初始数据") = False Then Unload Me: Exit Function

    Call ExecuteCommand("刷新数据")
    
    DataChanged = False
    
    '默认为新增
    Call chkAdd_Click
        
    '显示窗体
    
    Me.Show vbModal
    
    If mblnOK Then
        ShowMe = mlngDemoId
    Else
        ShowMe = 0
    End If

End Function

Private Property Let DataChanged(ByVal blnData As Boolean)
    mblnDataChanged = blnData
End Property

Private Property Get DataChanged() As Boolean
    DataChanged = mblnDataChanged
End Property

Private Function ExecuteCommand(ByVal strCommand As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
Dim intLoop As Integer
Dim rsTemp As New ADODB.Recordset
Dim rsSQL As New ADODB.Recordset
Dim strSQL As String
Dim strTmp As String
Dim lngTMP As Long
Dim byt应用范围 As Byte
Dim lngCount As Long
Dim rptCol As ReportColumn
Dim rptRcd As ReportRecord
Dim rptItem As ReportRecordItem
Dim rptRow As ReportRow
Dim mobjDoc As New cEPRDocument
    
    On Error GoTo errHand

    
    Select Case strCommand
    '--------------------------------------------------------------------------------------------------------------
    Case "初始控件"
        
        Call InitCommandBar
        
        '表格
        With rptList
                
            Set rptCol = .Columns.Add(mLCol.图标, "", 18, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Alignment = xtpAlignmentCenter
            Set rptCol = .Columns.Add(mLCol.性质, "", 18, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Alignment = xtpAlignmentCenter
            Set rptCol = .Columns.Add(mLCol.ID, "ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
            Set rptCol = .Columns.Add(mLCol.科室ID, "科室id", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
            Set rptCol = .Columns.Add(mLCol.人员ID, "人员id", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
            Set rptCol = .Columns.Add(mLCol.分类, "分类", 0, False): rptCol.Editable = False: rptCol.Groupable = False:  rptCol.Visible = False
            Set rptCol = .Columns.Add(mLCol.编号, "编号", 49, False): rptCol.Editable = False: rptCol.Groupable = False: .SortOrder.Add rptCol
            Set rptCol = .Columns.Add(mLCol.名称, "名称", 100, True): rptCol.Editable = False: rptCol.Groupable = False
            Set rptCol = .Columns.Add(mLCol.简码, "简码", 60, False): rptCol.Editable = False: rptCol.Groupable = False
            Set rptCol = .Columns.Add(mLCol.说明, "说明", 200, True): rptCol.Editable = False: rptCol.Groupable = False
            Set rptCol = .Columns.Add(mLCol.部门, "部门", 70, True): rptCol.Editable = False: rptCol.Groupable = True
            Set rptCol = .Columns.Add(mLCol.人员, "编制人", 50, False): rptCol.Editable = False: rptCol.Groupable = True
            
            .SetImageList Me.imgList
            .AllowColumnRemove = False
            .MultipleSelection = False
            .ShowItemsInGroups = False
            With .PaintManager
                .ColumnStyle = xtpColumnFlat
                .GridLineColor = RGB(225, 225, 225)
                .NoGroupByText = "拖动列标题到这里,按该列分组..."
                .NoItemsText = "没有可显示的项目..."
                .VerticalGridStyle = xtpGridSolid
            End With
            .GroupsOrder.DeleteAll
            .GroupsOrder.Add .Columns.Find(mLCol.分类)
            .GroupsOrder(0).SortAscending = True
            .SortOrder.Add .Columns.Find(mLCol.编号)
        End With
        
        txt(0).MaxLength = GetMaxLength("病历范文目录", "编号")
        txt(1).MaxLength = GetMaxLength("病历范文目录", "名称")
        txt(2).MaxLength = GetMaxLength("病历范文目录", "简码")
        txt(3).MaxLength = GetMaxLength("病历范文目录", "说明")
        
    '--------------------------------------------------------------------------------------------------------------
    Case "初始数据"
         
        
        If InStr(1, gstrPrivsEpr, "全院病历范文") <> 0 Then
            mbytPower = 0
        ElseIf InStr(1, gstrPrivsEpr, "科室病历范文") <> 0 Then
            mbytPower = 1
            opt范围(0).Enabled = False
        ElseIf InStr(1, gstrPrivsEpr, "个人病历范文") <> 0 Then
            mbytPower = 2
            opt范围(0).Enabled = False
            opt范围(1).Enabled = False
        Else
            mbytPower = -1
            MsgBox "对不起，你不具备范文编辑权限！", vbInformation, gstrSysName
            Exit Function
        End If
        
        '选择可选择的缺省
        If opt范围(2).Enabled Then
            opt范围(2).Value = 1
        Else
            If opt范围(1).Enabled Then
                opt范围(2).Value = 1
            ElseIf opt范围(0).Enabled Then
                opt范围(0).Value = 1
            End If
        End If
    
        Me.Caption = "另存为" & IIf(mstrCompends = "", "范文", "片段") & "..."
    
        '获取文件定义id
        If mbytFromTab = 1 Then
            gstrSQL = "Select f.Id, f.编号, f.名称 From 病历文件列表 f, 病历范文目录 s Where f.Id = s.文件id And s.Id = [1]"
        Else
            gstrSQL = "Select f.Id, f.编号, f.名称 From 病历文件列表 f, 电子病历记录 s Where f.Id = s.文件id And s.Id = [1]"
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFromId)
        If rsTemp.RecordCount <= 0 Then
            MsgBox "对应文件定义丢失，不能另存" & IIf(mstrCompends = "", "范文", "片段") & "！", vbExclamation, gstrSysName
            Exit Function
        End If
        
        Me.Caption = "保存为“" & rsTemp!编号 & "-" & rsTemp!名称 & "”的" & IIf(mstrCompends = "", "范文", "片段") & ":"
        mlngFileID = rsTemp!ID
        
        '基本数据信息
        gstrSQL = "Select Distinct D.ID, D.编码, D.名称, R.缺省, R.人员id, P.姓名" & vbNewLine & _
                "From 部门表 D, 部门人员 R, 人员表 P, 上机人员表 U, 部门性质说明 C," & vbNewLine & _
                "     (Select 种类, 通用 From 病历文件列表 Where ID = [1]) L" & vbNewLine & _
                "Where D.ID = R.部门id And R.人员id = P.ID And P.ID = U.人员id And U.用户名 = User And D.ID = C.部门id And" & vbNewLine & _
                "      C.工作性质 In ('临床', '检查', '检验', '手术', '治疗', '护理', '营养', '体检') And" & vbNewLine & _
                "      (Nvl(L.通用, 0) <> 2 Or L.种类 = 7 Or" & vbNewLine & _
                "      L.种类 <> 7 And L.通用 = 2 And D.ID In (Select 科室id From 病历应用科室 Where 文件id = [1]))" & vbNewLine & _
                "Order By R.缺省 Desc, D.编码"
                
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
        With rsTemp
            If .RecordCount <= 0 Then
                MsgBox "你目前不属于该病历应用科室范围，不能管理" & IIf(mstrCompends = "", "范文", "片段") & "！", vbExclamation, gstrSysName
                Exit Function
            End If
            Do While Not .EOF
                cbo(1).AddItem !编码 & "-" & !名称
                cbo(1).ItemData(cbo(1).NewIndex) = !ID
                If !缺省 = 1 Then cbo(1).ListIndex = cbo(1).NewIndex
                mlngSelfId = !人员ID: mstrSelfName = !姓名
                .MoveNext
            Loop
            If cbo(1).ListIndex = -1 Then cbo(1).ListIndex = 0
        End With
        
        cbo(0).Clear
        cbo(0).AddItem ""
        gstrSQL = "Select a.分类 From 病历范文目录 a Where a.文件id=[1] And a.分类 Is Not Null"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
        If rsTemp.BOF = False Then
            Do While Not rsTemp.EOF
                cbo(0).AddItem rsTemp("分类").Value
                rsTemp.MoveNext
            Loop
        End If
        cbo(0).ListIndex = 0

    '--------------------------------------------------------------------------------------------------------------
    Case "刷新数据"

        Call ExecuteCommand("读取范文")
        Call ExecuteCommand("读取条件")
            
    '--------------------------------------------------------------------------------------------------------------
    Case "读取范文"
        
        Dim objItem As ReportRecordItem
        
        gstrSQL = "Select L.ID, L.编号, L.名称, L.简码, Nvl(L.分类, '未分类') As 分类, L.性质, L.说明, L.通用级," & vbNewLine & _
                    "       L.科室id, L.人员id, D.名称 As 科室, P.姓名 As 人员,Decode(L.分类, Null, 1, 2) As 排序" & vbNewLine & _
                    "From 病历范文目录 L, 部门表 D, 人员表 P" & vbNewLine & _
                    "Where L.科室id = D.ID And L.人员id = P.ID And L.文件id =[1] And Nvl(L.性质, 0) =[2] And L.通用级 >=[3]" & vbNewLine & _
                    Decode(mbytPower, 0, "", 1, " And 科室ID In (Select R.部门id From 部门人员 R, 上机人员表 U Where R.人员id = U.人员id And U.用户名 = User) ", 2, " And 人员ID=[5] ") & vbNewLine & _
                    "Order By Decode(L.分类, Null, 1, 2), L.分类, L.编号"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID, IIf(mstrCompends = "", 0, 1), mbytPower, glngDeptId, glngUserId)
        
        rptList.Records.DeleteAll
        Do While Not rsTemp.EOF
            Set rptRcd = rptList.Records.Add()
            Set rptItem = rptRcd.AddItem(CInt(IIf(IsNull(rsTemp!通用级), 0, rsTemp!通用级))): rptItem.Icon = rptItem.Value
            Set rptItem = rptRcd.AddItem(CInt(Val("" & rsTemp!性质))): rptItem.Icon = IIf(rptItem.Value = 0, 4, 5)
            rptRcd.AddItem CStr(rsTemp!ID)
            rptRcd.AddItem zlCommFun.NVL(rsTemp!科室ID, 0)
            rptRcd.AddItem zlCommFun.NVL(rsTemp!人员ID, 0)
            Set objItem = rptRcd.AddItem(Val(rsTemp!排序) & CStr(rsTemp!分类))
            objItem.Caption = CStr(rsTemp!分类)
            rptRcd.AddItem CStr(rsTemp!编号)
            rptRcd.AddItem CStr(rsTemp!名称)
            rptRcd.AddItem CStr("" & rsTemp!简码)
            rptRcd.AddItem CStr("" & rsTemp!说明)
            rptRcd.AddItem CStr("" & rsTemp!科室)
            rptRcd.AddItem CStr("" & rsTemp!人员)
            rsTemp.MoveNext
        Loop
        rptList.Populate
        
        If rptList.Rows.Count > 0 Then
            For Each rptRow In Me.rptList.Rows
                If Not (rptRow.Record Is Nothing) Then
                    If mlngDemoId = rptRow.Record(mLCol.ID).Value Then Set rptList.FocusedRow = rptRow: Exit For
                End If
            Next
            If rptList.FocusedRow Is Nothing Then Set rptList.FocusedRow = rptList.Rows(0)
        End If

    '--------------------------------------------------------------------------------------------------------------
    Case "读取条件"
    
        With vfgTerm
            .Clear
            .Rows = .FixedRows
            Set .Cell(flexcpPicture, .FixedRows - 1, 0) = imgList.ListImages(4).Picture
                        
            If Not (rptList.FocusedRow Is Nothing) Then
                If Not (rptList.FocusedRow.Record Is Nothing) Then
                    lngTMP = rptList.FocusedRow.Record.Item(mLCol.ID).Value
                End If
            End If
                        
            gstrSQL = "Select 名称 As 条件项, 简码 As 条件值" & vbNewLine & _
                    "From Table(Cast(f_Segment_条件项([1]) As " & gstrDbOwner & ".t_Dic_Rowset))" & vbNewLine & _
                    "Where 简码 Is Not Null"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngTMP)
            
            If rsTemp.RecordCount <= 0 Then
                .TextMatrix(.FixedRows - 1, 0) = "无使用限制条件。"
            Else
                .TextMatrix(.FixedRows - 1, 0) = "在以下条件满足时可以使用："
            End If
            
            Do While Not rsTemp.EOF
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = Space(2) & .Rows - 1 & ")" & rsTemp!条件项 & "为'" & Replace(rsTemp!条件值, vbTab, "'或'") & "'"
                rsTemp.MoveNext
            Loop

            .AutoSize 0
        End With
    '--------------------------------------------------------------------------------------------------------------
    Case "读取明细"
        
        If Not (rptList.FocusedRow Is Nothing) Then
            If Not (rptList.FocusedRow.Record Is Nothing) Then
            
                
                '先判断是否有权限修改
                byt应用范围 = rptList.FocusedRow.Record.Item(mLCol.图标).Value
                
                If opt范围(byt应用范围).Enabled = False Then
                    MsgBox "对不起，你不能更改“" & rptList.FocusedRow.Record.Item(mLCol.名称).Value & "”" & IIf(mstrCompends = "", "片段", "范文") & "！", vbInformation, gstrSysName
                    Exit Function
                End If
                
                lngTMP = rptList.FocusedRow.Record.Item(mLCol.ID).Value
                
                chkAdd.Tag = lngTMP
                
                chkAdd.Value = vbUnchecked
                chkAdd.Enabled = True
                
                txt(0).Text = rptList.FocusedRow.Record.Item(mLCol.编号).Value
                txt(1).Text = rptList.FocusedRow.Record.Item(mLCol.名称).Value
                txt(2).Text = rptList.FocusedRow.Record.Item(mLCol.简码).Value
                txt(3).Text = rptList.FocusedRow.Record.Item(mLCol.说明).Value
                cbo(0).Text = rptList.FocusedRow.Record.Item(mLCol.分类).Caption
                
                opt范围(byt应用范围).Value = True
                                
                For lngCount = 0 To cbo(1).ListCount - 1
                    If cbo(1).ItemData(lngCount) = rptList.FocusedRow.Record.Item(mLCol.科室ID).Value Then
                        cbo(1).ListIndex = lngCount
                        Exit For
                    End If
                Next
                
                lbl人员.Tag = rptList.FocusedRow.Record.Item(mLCol.人员ID).Value
                lbl人员.Caption = rptList.FocusedRow.Record.Item(mLCol.人员).Value

            End If
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case "删除范文"
        
        If Not (rptList.FocusedRow Is Nothing) Then
            If Not (rptList.FocusedRow.Record Is Nothing) Then
            
                strTmp = "真的删除" & IIf(mstrCompends = "", "范文", "片段") & "“" & rptList.FocusedRow.Record.Item(mLCol.名称).Value & "”吗？"
                If MsgBox(strTmp, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    
                    lngTMP = rptList.FocusedRow.Record.Item(mLCol.ID).Value
                    
                    gstrSQL = "zl_病历范文目录_delete('" & lngTMP & "')"
                    
                    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                    
                    rptList.Records.RemoveAt rptList.FocusedRow.Record.Index
                    rptList.Populate
                    
                    If lngTMP = Val(chkAdd.Tag) And lngTMP > 0 Then
                        Call chkAdd_Click
                    End If
                    
                    Call ExecuteCommand("读取条件")
                End If
                
            End If
        End If

    '--------------------------------------------------------------------------------------------------------------
    Case "校验数据"
                
        If Trim(txt(0).Text) = "" Then MsgBox "请输入编号！", vbInformation, gstrSysName: txt(0).SetFocus: Exit Function
        If Trim(txt(1).Text) = "" Then MsgBox "请输入名称！", vbInformation, gstrSysName: txt(1).SetFocus: Exit Function
        If cbo(1).ListIndex = -1 Then MsgBox "请输入科室！", vbInformation, gstrSysName: cbo(1).SetFocus: Exit Function
        
        If Val(chkAdd.Tag) > 0 Then
            
            gstrSQL = "Select 通用级,科室ID,名称,人员ID From 病历范文目录 Where ID=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(chkAdd.Tag))
            If rsTemp.BOF = False Then
                
                Select Case zlCommFun.NVL(rsTemp("通用级").Value, 0)
                Case 0            '全院通用
                    
                Case 1            '科室通用
                    '本科室
                    If zlCommFun.NVL(rsTemp("科室id").Value, 0) <> glngDeptId Then
                        '禁止
                        Call MsgBox("你无权覆盖已存在的“" & zlCommFun.NVL(rsTemp("名称").Value) & "”" & IIf(mstrCompends = "", "范文", "片段") & "！", vbInformation, gstrSysName)
                        
                        Exit Function
                    End If
                Case 2            '个人通用
                    '本人
                    If zlCommFun.NVL(rsTemp("人员id").Value, 0) <> glngUserId Then
                        '禁止
                        Call MsgBox("你无权覆盖已存在的“" & zlCommFun.NVL(rsTemp("名称").Value) & "”" & IIf(mstrCompends = "", "范文", "片段") & "！", vbInformation, gstrSysName)
                                                
                        Exit Function
                    End If
                End Select
                
                If MsgBox("你选择了覆盖已存在的“" & zlCommFun.NVL(rsTemp("名称").Value) & "”" & IIf(mstrCompends = "", "范文", "片段") & "，是否继续？", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            
            End If
            
        End If
        
    '--------------------------------------------------------------------------------------------------------------
    Case "保存数据"

        If chkAdd.Value = vbChecked Then
            '新增范文/片段
            mlngDemoId = zlDatabase.GetNextId("病历范文目录")
        Else
            '修改原有的范文/片段
            mlngDemoId = Val(chkAdd.Tag)
        End If
        
        byt应用范围 = IIf(opt范围(0).Value, 0, IIf(opt范围(1).Value, 1, 2))

        If mstrCompends <> "" Then
            '如果是片段，直接调用存储过程转储
            
            gstrSQL = mlngDemoId & ",'" & _
                        Trim(txt(0).Text) & "','" & _
                        Trim(Me.txt(1).Text) & "','" & _
                        Trim(Me.txt(2).Text) & "','" & _
                        Replace(Trim(Me.txt(3).Text), Chr(vbKeyReturn), "") & "'," & _
                        byt应用范围 & "," & _
                        cbo(1).ItemData(cbo(1).ListIndex) & "," & _
                        Val(lbl人员.Tag) & "," & _
                        mbytFromTab & "," & _
                        mlngFromId & ",'" & _
                        mstrCompends & "','" & _
                        cbo(0).Text & "'"
                                
            gstrSQL = "Zl_病历范文目录_Segment(" & gstrSQL & ")"
        
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        Else
            '如果是范文，则在保存目录后，导入文件进行保存
            
            gstrSQL = mlngDemoId & IIf(Me.chkAdd.Value = vbChecked, "," & mlngFileID, "")
            gstrSQL = gstrSQL & ",'" & Trim(Me.txt(0).Text) & "','" & Trim(Me.txt(1).Text) & "','" & Trim(Me.txt(2).Text) & "'"
            gstrSQL = gstrSQL & IIf(Me.chkAdd.Value = vbChecked, ",0", "") & ",'" & Replace(Trim(Me.txt(3).Text), Chr(vbKeyReturn), "") & "'"
            gstrSQL = gstrSQL & "," & byt应用范围

            gstrSQL = gstrSQL & "," & cbo(1).ItemData(cbo(1).ListIndex) & IIf(chkAdd.Value = vbChecked, "," & lbl人员.Tag, "") & ",'" & cbo(0).Text & "'"
            gstrSQL = IIf(chkAdd.Value = vbChecked, "Zl_病历范文目录_Insert", "Zl_病历范文目录_Update") & "(" & gstrSQL & ")"
            
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        
            '导入指定文件到控件,并保存
            
            
            mobjDoc.InitEPRDoc cprEM_修改, cprET_全文示范编辑, mlngDemoId
            If mbytFromTab = 1 Then
                mobjDoc.ImportEPRDemo Editor1, mlngFromId, True
            Else
                Editor1.AuditMode = True
                mobjDoc.ImportOldEPRFile Editor1, mlngFromId, True
                Editor1.AuditMode = False
            End If
            If mobjDoc.SaveEPRDoc(Editor1) = False Then
                MsgBox "虽然范文目录保存，但在内容保存中发生错误！", vbExclamation, gstrSysName
            End If
            Set mobjDoc = Nothing
        End If
    
    End Select
    ExecuteCommand = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function InitCommandBar() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl

    '------------------------------------------------------------------------------------------------------------------
    '初始设置
    Call CommandBarInit(cbsMain)

    '------------------------------------------------------------------------------------------------------------------
    '菜单定义:包括公共部份，请对xtpControlPopup类型的命令ID重新赋值

    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.Options.LargeIcons = True
    cbsMain.ActiveMenuBar.Visible = False
    
    '------------------------------------------------------------------------------------------------------------------
    '工具栏定义:包括公共部份

    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
    
    
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Save, "保存(&S)")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Delete, "删除(&D)", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Help_Help, "帮助(&H)", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Exit, "退出(&X)")
    
    '------------------------------------------------------------------------------------------------------------------
    '命令的快键绑定:公共部份主界面已处理

    With cbsMain.KeyBindings
        .Add 0, vbKeyF1, conMenu_Help_Help                  '帮助
        .Add 0, vbKeyF2, conMenu_Edit_Save                  '保存
    End With

End Function

'以下为控件事件处理
'######################################################################################################################

Private Sub cbo_Change(Index As Integer)
    
End Sub

Private Sub cbo_Click(Index As Integer)
    DataChanged = True
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)

    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Save                        '保存
        If ExecuteCommand("校验数据") And DataChanged Then
            If ExecuteCommand("保存数据") Then
                mblnOK = True
                Unload Me
            End If
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete
        
        Call ExecuteCommand("删除范文")
        
    '--------------------------------------------------------------------------------------------------------------
    Case Else

         '与业务无关的功能，公共的功能
        Call CommandBarExecutePublic(Control, Me)

    End Select
End Sub


Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long
    Dim lngTop  As Long
    Dim lngRight  As Long
    Dim lngBottom  As Long

    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)

    On Error Resume Next
    
    picPane(0).Move lngLeft, lngTop, lngRight - lngLeft - picPane(2).Width, lngBottom - lngTop - picPane(1).Height
    picPane(1).Move lngLeft, picPane(0).Top + picPane(0).Height, picPane(0).Width
    picPane(2).Move picPane(1).Left + picPane(1).Width, lngTop, picPane(2).Width, lngBottom - lngTop
    
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)

    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Save
        Control.Enabled = DataChanged
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete

        If Not (rptList.FocusedRow Is Nothing) Then
            If Not (rptList.FocusedRow.Record Is Nothing) Then
                Control.Enabled = rptList.FocusedRow.Record.Item(mLCol.ID).Value > 0
            Else
                Control.Enabled = False
            End If
        Else
            Control.Enabled = False
        End If
            
    '--------------------------------------------------------------------------------------------------------------
    Case Else

        '与业务无关的功能，公共的功能
        Call CommandBarUpdatePublic(Control, Me)

    End Select
End Sub

Private Sub chkAdd_Click()
    Dim rsTemp As New ADODB.Recordset
    
    If Me.chkAdd.Value <> vbChecked Then Exit Sub

    txt(0).Text = GetMax("病历范文目录", "编号", txt(0).MaxLength, " Where 文件id=" & mlngFileID)
    txt(1).Text = "新" & IIf(mstrCompends = "", "范文-", "片段-") & Me.txt(0).Text
    txt(2).Text = Left(zlCommFun.SpellCode(txt(1).Text), 10)
    lbl人员.Tag = mlngSelfId: Me.lbl人员.Caption = mstrSelfName
        
    If txt(0).Visible Then txt(0).SetFocus
    
    Me.chkAdd.Enabled = False
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub opt范围_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
       
    Select Case Index
    Case 0
        fra(1).Move 0, -75, picPane(Index).Width, picPane(Index).Height + 75
        rptList.Move 15, 105, fra(1).Width - 45, fra(1).Height - 105 - 30
    Case 1
    
        fra(2).Move 0, -75, picPane(Index).Width, picPane(Index).Height + 75
        vfgTerm.Move 15, 105, fra(2).Width - 45, fra(2).Height - 105 - 30
        
    Case 2
        fra(0).Move 0, -75, picPane(Index).Width, picPane(Index).Height + 75
        txt(3).Move txt(3).Left, txt(3).Top, txt(3).Width, fra(0).Height - txt(3).Top - 810
        
        cbo(1).Move cbo(1).Left, txt(3).Top + txt(3).Height + 45
        lbl科室.Top = cbo(1).Top + 45
        
        lbl人员.Move lbl人员.Left, cbo(1).Top + cbo(1).Height + 45
    End Select
    
End Sub

Private Sub rptList_KeyDown(KeyCode As Integer, Shift As Integer)


    If KeyCode = vbKeyDelete Then
        Call ExecuteCommand("删除范文")
    End If

End Sub

Private Sub rptList_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    
    Call ExecuteCommand("读取明细")
    
End Sub

Private Sub rptList_SelectionChanged()
    
    Call ExecuteCommand("读取条件")
    
End Sub

Private Sub txt_Change(Index As Integer)
    DataChanged = True
End Sub

Private Sub txt_GotFocus(Index As Integer)
    
    zlControl.TxtSelAll txt(Index)
    
    Select Case Index
    Case 1, 3
        zlCommFun.OpenIme True
    End Select
        
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0

        zlCommFun.PressKey vbKeyTab

    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
        Select Case Index
            Case 0
                If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
            Case 2
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End Select
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)

    Select Case Index
    Case 1
        zlCommFun.OpenIme False
        If InStr(txt(Index).Text, "'") = 0 Then txt(2).Text = zlGetSymbol(txt(Index).Text)

    Case 3
        zlCommFun.OpenIme False
        txt(Index) = Replace(Me.txt(Index).Text, Chr(vbKeyReturn), "")
    End Select

End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        glngTXTProc = GetWindowLong(txt(Index).Hwnd, GWL_WNDPROC)
        Call SetWindowLong(txt(Index).Hwnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        Call SetWindowLong(txt(Index).Hwnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    '包括%符号
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
End Sub
