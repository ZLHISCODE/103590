VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmStandardPathRef 
   Caption         =   "标准路径参考"
   ClientHeight    =   8730
   ClientLeft      =   6345
   ClientTop       =   2085
   ClientWidth     =   13755
   Icon            =   "frmStandardPathRef.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8730
   ScaleWidth      =   13755
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picSTPathList 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8010
      Left            =   0
      ScaleHeight     =   8010
      ScaleWidth      =   4005
      TabIndex        =   10
      Top             =   0
      Width           =   4005
      Begin XtremeReportControl.ReportControl rptStPath 
         Height          =   1695
         Left            =   240
         TabIndex        =   11
         Top             =   5760
         Width           =   2820
         _Version        =   589884
         _ExtentX        =   4974
         _ExtentY        =   2990
         _StockProps     =   0
         SkipGroupsFocus =   0   'False
      End
      Begin VB.PictureBox picFind 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   3855
         TabIndex        =   13
         Top             =   120
         Width           =   3855
         Begin VB.CommandButton cmdImport 
            Caption         =   "导入路径"
            Height          =   300
            Left            =   2710
            TabIndex        =   15
            Top             =   360
            Width           =   1100
         End
         Begin VB.TextBox txtInput 
            Height          =   300
            Left            =   435
            MaxLength       =   100
            TabIndex        =   14
            Top             =   0
            Width           =   3375
         End
         Begin VB.Label lblFind 
            BackColor       =   &H00FFFFFF&
            Caption         =   "查找"
            Height          =   255
            Left            =   0
            TabIndex        =   16
            Top             =   30
            Width           =   375
         End
      End
      Begin XtremeSuiteControls.TabControl tbcPathName 
         Height          =   1455
         Left            =   120
         TabIndex        =   12
         Top             =   4200
         Width           =   3015
         _Version        =   589884
         _ExtentX        =   5318
         _ExtentY        =   2566
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picStPathDetial 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   7335
      Left            =   5280
      ScaleHeight     =   7335
      ScaleWidth      =   7575
      TabIndex        =   1
      Top             =   0
      Width           =   7575
      Begin VB.PictureBox picPathTable 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2295
         Left            =   240
         ScaleHeight     =   2295
         ScaleWidth      =   6255
         TabIndex        =   3
         Top             =   1800
         Width           =   6255
         Begin VB.Frame fraSplitNS 
            BackColor       =   &H00F0F4E4&
            BorderStyle     =   0  'None
            Height          =   100
            Left            =   0
            TabIndex        =   9
            Top             =   1200
            Width           =   6255
         End
         Begin VSFlex8Ctl.VSFlexGrid vsPathTable 
            Height          =   975
            Left            =   0
            TabIndex        =   4
            Top             =   1320
            Width           =   3585
            _cx             =   6324
            _cy             =   1720
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
            BackColor       =   16777215
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   16777215
            BackColorAlternate=   16777215
            GridColor       =   32768
            GridColorFixed  =   32768
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   3
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   3
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   4
            Cols            =   4
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   20
            RowHeightMax    =   5000
            ColWidthMin     =   20
            ColWidthMax     =   9000
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmStandardPathRef.frx":058A
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
         Begin VB.Frame fra表头 
            BackColor       =   &H00F0F4E4&
            BorderStyle     =   0  'None
            Height          =   1095
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   6255
            Begin VB.Label lbl表头 
               AutoSize        =   -1  'True
               BackColor       =   &H00F0F4E4&
               Height          =   180
               Left            =   120
               TabIndex        =   6
               Top             =   0
               Width           =   90
            End
         End
      End
      Begin XtremeSuiteControls.TabControl tbcStPath 
         Height          =   7335
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   6615
         _Version        =   589884
         _ExtentX        =   11668
         _ExtentY        =   12938
         _StockProps     =   64
      End
      Begin VB.PictureBox picPathCourse 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3735
         Left            =   0
         ScaleHeight     =   3735
         ScaleWidth      =   6255
         TabIndex        =   7
         Top             =   480
         Width           =   6255
         Begin RichTextLib.RichTextBox rtfPathCourse 
            Height          =   4095
            Left            =   120
            TabIndex        =   8
            Top             =   0
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   7223
            _Version        =   393217
            BackColor       =   16777215
            BorderStyle     =   0
            ReadOnly        =   -1  'True
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"frmStandardPathRef.frx":05F6
         End
      End
   End
   Begin VB.Frame fraSplit 
      Caption         =   "Frame1"
      Height          =   7335
      Left            =   5200
      MousePointer    =   9  'Size W E
      TabIndex        =   0
      Top             =   0
      Width           =   45
   End
End
Attribute VB_Name = "frmStandardPathRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrs表单         As New ADODB.Recordset      '标准路径表单
Private mrs表头信息     As New ADODB.Recordset      '标准路径表单的表头以及分类信息等信息
Private mlngStPathID    As Long                     '选中的标准路径的ID
Private mlngFunc        As Long                     '0-标准路径表单查看,1-将标准路径表单导成临床路径表单
Private mblnOK          As Boolean                  'mblnFunc=1:导入成功 返回True
Private mintMode        As Integer                  '0-住院；1-门诊
Private Const M_INT_STEPNUM = 3

Private Enum PathListCols
    COL_ID = 0
    COL_科室名称 = 1
    COL_编码 = 2
    COL_路径名称 = 3
    COL_版本说明 = 4
    COL_疾病编码 = 5
End Enum

Private Enum CATE_TYPE
    IX_路径名称 = 0
    IX_疾病编码 = 1
End Enum

Public Function ShowMe(frmMain As Object, ByVal lngStPathID As Long, Optional ByVal lngFunc As Long = 0, Optional ByVal intMode As Integer)
'参数：lngStPathID 选中的标准路径
    mblnOK = False
    mlngStPathID = lngStPathID
    mlngFunc = lngFunc
    mintMode = intMode
    Me.Show 1, frmMain
    
    ShowMe = mblnOK
End Function

Private Sub FuncFindSTPath(Optional ByVal blnNext As Boolean)
'功能:根据输入文本内容,定位路径表单
    Dim i As Long
    Dim blnHave As Boolean
    Dim blnReStart As Boolean
    Dim objRow As Object
    
    Call zlControl.TxtSelAll(txtInput)
    '开始查找行
    If rptStPath.SelectedRows.count > 0 Then blnHave = True
    If Not blnNext Or blnReStart Or Not blnHave Then
        i = 0    'ReportControl的索引从是0开始
    Else
        i = rptStPath.SelectedRows(0).Index + 1
    End If

    For Each objRow In rptStPath.Rows
        objRow.Expanded = True
    Next
    '查找路径
    For i = i To rptStPath.Rows.count - 1
        With rptStPath.Rows(i)
            If .Record.Tag <> "" Then
                If zlStr.IsCharChinese(Trim(txtInput.Text)) Then
                    If .Record(COL_路径名称).Value Like "*" & Trim(txtInput.Text) & "*" Then
                        Exit For
                    End If
                Else '疾病编码
                    If .Record(COL_疾病编码).Value Like "*" & UCase(Trim(txtInput.Text)) & "*" Then
                        Exit For
                    End If
                End If
        
            End If
        End With
    Next
    
    If i <= rptStPath.Rows.count - 1 Then
        blnReStart = False
        '该行选中且显示在可见区域,并引发SelectionChanged事件
        Set rptStPath.FocusedRow = rptStPath.Rows(i)

        If rptStPath.Visible Then rptStPath.SetFocus
    Else
        blnReStart = True
        MsgBox IIf(blnNext, "后面已", "") & "找不到符合条件的标准路径。", vbInformation, gstrSysName
    End If
End Sub

Private Sub cmdImport_Click()
'功能:将标准路径表单导成临床路径表单
    Dim strSql As String

    On Error GoTo errH
    If mlngStPathID <> 0 Then
        If mintMode = 1 Then
            strSql = "zl_门诊路径导入_Import(" & mlngStPathID & ")"
        Else
            strSql = "zl_临床路径导入_Import(" & mlngStPathID & ")"
        End If
        zlDatabase.ExecuteProcedure strSql, Me.Caption
        MsgBox "导入成功!", vbInformation + vbOKOnly, gstrSysName
        mblnOK = True
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'mlngFunc=1 支持查找功能
    If mlngFunc = 1 Then
        If KeyCode = vbKeyF And Shift = vbCtrlMask Then
            txtInput.SetFocus
            If Trim(txtInput.Text) <> "" Then
                Call FuncFindSTPath
            End If
        ElseIf KeyCode = vbKeyF3 Then
            If Trim(txtInput.Text) <> "" Then
                FuncFindSTPath (True)
            End If
            txtInput.SetFocus
        End If
    End If
End Sub

Private Sub Form_Load()
'功能：初始化tbcControl 与reportControl
    Dim objCol        As ReportColumn

    If mlngFunc = 0 Then
        If mintMode = 1 Then
            Me.Caption = "标准门诊路径参考"
        Else
            Me.Caption = "标准路径参考"
        End If
        picFind.Visible = False
    Else
        If mintMode = 1 Then
            Me.Caption = "导入标准门诊路径"
        Else
            Me.Caption = "导入标准路径"
        End If
        picFind.Visible = True
    End If
    'tbcPathName路径参考
    With Me.tbcPathName
        With .PaintManager
            .Appearance = xtpTabAppearanceExcel
            .Color = xtpTabColorOffice2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        
        .InsertItem 0, "西医参考", rptStPath.Hwnd, 0
        .InsertItem 1, "中医参考", rptStPath.Hwnd, 0
        
        .Item(1).Selected = True
        .Item(0).Selected = True
        
    End With
    
    With rptStPath
        '初始化Report控件的列与属性
        Set objCol = .Columns.Add(PathListCols.COL_ID, "ID", 20, False)
            objCol.Alignment = xtpAlignmentCenter: objCol.Resizable = True: objCol.AllowDrag = False: objCol.Visible = False
        Set objCol = .Columns.Add(PathListCols.COL_科室名称, "科室名称", 80, False)
            objCol.Resizable = True: objCol.Alignment = xtpAlignmentLeft: objCol.AllowDrag = False: objCol.TreeColumn = True: objCol.Groupable = True
        Set objCol = .Columns.Add(PathListCols.COL_编码, "编码", 50, False)
            objCol.Alignment = xtpAlignmentLeft: objCol.Resizable = True: objCol.AllowDrag = False
        Set objCol = .Columns.Add(PathListCols.COL_路径名称, "路径名称", 200, False)
            objCol.Alignment = xtpAlignmentLeft: objCol.Resizable = True: objCol.AllowDrag = False
        Set objCol = .Columns.Add(PathListCols.COL_版本说明, "版本说明", 70, False)
            objCol.Alignment = xtpAlignmentLeft: objCol.Resizable = True: objCol.AllowDrag = False
        Set objCol = .Columns.Add(PathListCols.COL_疾病编码, "疾病编码", 200, False)
            objCol.Alignment = xtpAlignmentLeft: objCol.Resizable = True: objCol.AllowDrag = False
            
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoItemsText = "没有可显示的项目."
        End With
        
        .AutoColumnSizing = False
        .AllowColumnRemove = False
        .ShowGroupBox = False
        .ShowItemsInGroups = False
        .PreviewMode = True
        .MultipleSelection = False '会引发SelectionChanged事件
    End With
    
    '初始化tbcControl
    With tbcStPath
        With .PaintManager
            .Appearance = xtpTabAppearanceVisualStudio
            .Color = xtpTabColorOffice2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
        End With
        
        .AllowReorder = False
        '初次加载数据只加载选项卡以及标准流程
        If mintMode = 1 Then
            Call .InsertItem(0, "标准门诊流程", picPathCourse.Hwnd, 0)
        Else
            Call .InsertItem(0, "标准住院流程", picPathCourse.Hwnd, 0)
        End If
        .Item(0).Selected = True '默认选择标准流程
    End With
    
    '加载标准路径目录
    Call LoadStPathList(0, True)
    '根据选择的标准路径ID加载路径流程，路径表单，表单表头
    Call LoadPathByID(mlngStPathID, True, 0)
End Sub

Private Sub Form_Resize()
'功能：设置tbcPathName与picStPathDetial的位置大小
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.WindowState <> vbMaximized And Me.WindowState <> vbMinimized Then
        Me.Height = IIf(Me.Height < 9000, 9000, Me.Height)
        Me.Width = IIf(Me.Width < 12000, 12000, Me.Width)
    End If

    picSTPathList.Move 0, 0, Me.ScaleWidth * 0.3, Me.ScaleHeight
   
    fraSplit.Left = picSTPathList.Left + picSTPathList.Width + 30
    fraSplit.Height = Me.ScaleHeight
    
    picStPathDetial.Left = fraSplit.Left + fraSplit.Width + 30
    picStPathDetial.Width = Me.ScaleWidth - picStPathDetial.ScaleLeft
    picStPathDetial.Height = Me.ScaleHeight - picStPathDetial.ScaleTop
 
End Sub

Private Sub fraSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'功能：实现标准路径清单与标准路径内容自由拖动大小
    If Button = 1 Then
        If picSTPathList.Width + X > 11000 Or picSTPathList.Width + X < 2000 Then Exit Sub
        
        fraSplit.Left = fraSplit.Left + X
        picSTPathList.Width = fraSplit.Left - 30 - picSTPathList.Left
        picStPathDetial.Left = fraSplit.Left + 30
        picStPathDetial.Width = Me.ScaleWidth - picStPathDetial.Left
        
        Me.Refresh
    End If
    
End Sub

Private Sub picPathCourse_Resize()
'功能：实现标准路径流程内容的大小设置

    rtfPathCourse.Width = picPathCourse.Width - rtfPathCourse.Left - 120
    rtfPathCourse.Height = picPathCourse.Height - rtfPathCourse.Top
    
End Sub

Private Sub picPathTable_Resize()
'功能：设置表单表头与表单所在控件的位置与大小

    fra表头.Height = lbl表头.Height + 60
    fra表头.Width = picPathTable.Width
    lbl表头.Width = fra表头.Width
    lbl表头.Width = fra表头.Width - lbl表头.Left
    
    fraSplitNS.Top = fra表头.Top + fra表头.Height
    fraSplitNS.Width = picPathTable.Width
    
    vsPathTable.Top = fraSplitNS.Top + fraSplitNS.Height
    vsPathTable.Height = picPathTable.Height - vsPathTable.Top - 120
    vsPathTable.Width = picPathTable.Width - vsPathTable.Left - 120
    
End Sub

Private Sub picStPathDetial_Resize()
'功能：标准路径内容区的大小设置

    tbcStPath.Width = picStPathDetial.Width
    tbcStPath.Height = picStPathDetial.Height
    picPathTable.Width = tbcStPath.Width '触发picStPathDetial_Resize
    picPathCourse.Width = tbcStPath.Width
    
End Sub

Private Sub picSTPathList_Resize()
    On Error Resume Next
    If mlngFunc = 0 Then
        tbcPathName.Top = 0
        tbcPathName.Left = 0
    Else
        picFind.Move 120, 120, picSTPathList.Width - 240, 850
        tbcPathName.Top = picFind.Height + picFind.Top
        tbcPathName.Left = 0
    End If
    
    tbcPathName.Width = picSTPathList.Width
    tbcPathName.Height = picSTPathList.Height - tbcPathName.Top

End Sub

Private Sub rptStPath_SelectionChanged()
'功能：保存选择的路径ID,并根据ID加载标准路径流程以及表单

    If Me.Visible Then
        If mlngStPathID <> Val(rptStPath.SelectedRows(0).Record.Tag) And Val(rptStPath.SelectedRows(0).Record.Tag) <> 0 Then
            mlngStPathID = Val(rptStPath.SelectedRows(0).Record.Tag)
            tbcPathName.Item(tbcPathName.Selected.Index).Tag = mlngStPathID
            Call LoadPathByID(mlngStPathID, True, 0)
        End If
        
        cmdImport.Enabled = (mlngStPathID <> 0 And mlngFunc = 1)
    End If
    
End Sub

Private Sub tbcPathName_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
'功能:根据选项卡选择加载内容
    If Me.Visible Then
        mlngStPathID = 0
        Call LoadStPathList(Item.Index)
        Call LoadPathByID(mlngStPathID, True, 0)
    End If
End Sub

Private Sub tbcStPath_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
'功能：选择表单时加载表单内容

    If Me.Visible Then
        Call LoadPathByID(mlngStPathID, False, Item.Index)
        picPathCourse.Visible = Item.Index = 0
        picPathTable.Visible = Item.Index <> 0
    End If
End Sub

Private Sub LoadStPathList(ByVal lngIndex As Long, Optional ByVal blnFirst As Boolean)
'功能：加载标准路径目录
'参数:lngIndex 0-西医参考,1-中医参考
'     blnFirst True-首次加载,False-非首次加载
    Dim objRecord     As ReportRecord
    Dim objPreRecord     As ReportRecord
    Dim objItem       As ReportRecordItem
    Dim i As Long, strDept As String
    Dim rsTemp As ADODB.Recordset
    Dim strSql As String
    
On Error GoTo errH
    '清空记录,避免内容重复
    rptStPath.Records.DeleteAll
    
    If blnFirst And mlngStPathID <> 0 Then
        If mintMode = 1 Then
            strSql = "Select Nvl(t.类别, 0) as 类别 From 标准门诊路径目录 T Where t.Id = [1]"
        Else
            strSql = "Select Nvl(t.类别, 0) as 类别 From 标准路径目录 T Where t.Id = [1]"
        End If
        
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngStPathID)
        lngIndex = Val(rsTemp!类别 & "")
        tbcPathName.Item(lngIndex).Selected = True '选中指定项
    End If

    If mintMode = 1 Then
        strSql = " Select a.Id, a.科室名称, a.编码, a.路径名称, a.版本说明, b.疾病编码" & vbNewLine & _
                 " From 标准门诊路径目录 A, 标准门诊路径病种 B" & vbNewLine & _
                 " Where a.Id = b.标准路径id  and Nvl(a.类别,0)=[1] order by 科室名称,ID "
    Else
        strSql = " Select a.Id, a.科室名称, a.编码, a.路径名称, a.版本说明, b.疾病编码" & vbNewLine & _
                 " From 标准路径目录 A, 标准路径病种 B" & vbNewLine & _
                 " Where a.Id = b.标准路径id and Nvl(a.类别,0)=[1] order by 科室名称,ID "
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngIndex)
    
    For i = 0 To rsTemp.RecordCount - 1
        '在每个科室开始的地方添加分组行
        If strDept <> CStr(rsTemp!科室名称) Then
            Set objPreRecord = rptStPath.Records.Add()
            Set objItem = objPreRecord.AddItem(CStr(""))
            Set objItem = objPreRecord.AddItem(CStr(rsTemp!科室名称))
            Set objItem = objPreRecord.AddItem("")
            Set objItem = objPreRecord.AddItem("")
            Set objItem = objPreRecord.AddItem("")
            Set objItem = objPreRecord.AddItem("")
            objPreRecord.Tag = ""
            objPreRecord.Expanded = False
            '加载子记录
            Set objRecord = objPreRecord.Childs.Add()
            Set objItem = objRecord.AddItem(CStr(rsTemp!ID))
            Set objItem = objRecord.AddItem("")
            Set objItem = objRecord.AddItem(CStr(rsTemp!编码 & ""))
            Set objItem = objRecord.AddItem(CStr(rsTemp!路径名称 & ""))
            Set objItem = objRecord.AddItem(CStr(rsTemp!版本说明 & ""))
            Set objItem = objRecord.AddItem(CStr(rsTemp!疾病编码 & ""))
            objRecord.Tag = CStr(rsTemp!ID)
            strDept = CStr(rsTemp!科室名称)
            
            If mlngStPathID = 0 Then
                mlngStPathID = rsTemp!ID
                objPreRecord.Expanded = True
            Else
                If rsTemp!ID = mlngStPathID Then
                    objPreRecord.Expanded = True
                End If
            End If

            rsTemp.MoveNext
        Else
            Set objRecord = objPreRecord.Childs.Add()
            Set objItem = objRecord.AddItem(CStr(rsTemp!ID))
            Set objItem = objRecord.AddItem("")
            Set objItem = objRecord.AddItem(CStr(rsTemp!编码 & ""))
            Set objItem = objRecord.AddItem(CStr(rsTemp!路径名称 & ""))
            Set objItem = objRecord.AddItem(CStr(rsTemp!版本说明 & ""))
            Set objItem = objRecord.AddItem(CStr(rsTemp!疾病编码 & ""))
            objRecord.Tag = CStr(rsTemp!ID)
            If rsTemp!ID = mlngStPathID Then
                objPreRecord.Expanded = True
            End If
            strDept = CStr(rsTemp!科室名称)
            rsTemp.MoveNext
        End If
    Next
    rptStPath.Populate
    '定位到选择的标准路径
    For i = 0 To rptStPath.Rows.count - 1
        If Val(rptStPath.Rows(i).Record.Tag) = mlngStPathID Then
            rptStPath.Rows(i).Selected = True
            Exit For
        End If
    Next
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadPathByID(ByVal lngId As Long, Optional ByVal blnReadData As Boolean, Optional ByVal lng序号 As Long)
'功能：根据选择的标准路径ID读取数据，并根据表单序号加载路径流程，路径表单，表单表头
'参数：lngID   选择的路径ID
'      blnReadData 是否读取标准路径信息（在标准路径初次加载或者标准路径切换时是均需要读取）
'      lng序号  0 标准路径流程，1 表单1，2，表单2...
    Dim rsTmp As ADODB.Recordset
    Dim i As Long, j As Long, k As Long
    Dim strSql As String, strFilter As String
    Dim strTilePos As String '记录标题行位置格式为，标题1起始位置，长度;标题2起始位置，长度
    Dim lngColCount As Long, lng表单行数 As Long, lngBeginRow As Long
    Dim lngRowCount As Long
    Dim strContent As String
    
    On Error GoTo errH
    
    If blnReadData Then
        '删除选项卡，清空vs数据
        vsPathTable.Delete
        For i = tbcStPath.ItemCount - 1 To 1 Step -1
            tbcStPath.RemoveItem (i)
        Next
        
        '加载标准流程
        rtfPathCourse.Visible = False
        rtfPathCourse.Text = ""
        If mintMode = 1 Then
            strSql = "Select 标题, 内容 From 标准门诊路径流程 Where 标准路径id = [1] Order By 序号"
        Else
            strSql = "Select 标题, 内容 From 标准路径流程 Where 标准路径id = [1] Order By 序号"
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngId)
        If rsTmp.RecordCount <> 0 Then
            For i = 1 To rsTmp.RecordCount
                strTilePos = strTilePos & ";" & Len(strContent) & "," & Len(rsTmp!标题)
                strContent = strContent & rsTmp!标题 & vbNewLine & vbNewLine & rsTmp!内容 & vbNewLine & vbNewLine
                rsTmp.MoveNext
            Next
            rtfPathCourse.Text = strContent
        End If
               
        
        Call SetStPathCourceFont(Mid(strTilePos, 2)) '设置字体
        rtfPathCourse.Visible = True
        
        '读取表单总体信息
        If mintMode = 1 Then
            strSql = "Select a.表单序号 表单序号, b.表单名称, b.表单表头, a.行数, a.列数" & vbNewLine & _
                    "From (Select 表单序号, Max(分类序号) 行数, Max(阶段序号) 列数 From 标准门诊路径表单 Where 标准路径id = [1] Group By 表单序号) A, 标准门诊路径表单 B" & vbNewLine & _
                    "Where b.标准路径id =[1] And a.表单序号 = b.表单序号 And b.分类序号 = 1 And b.阶段序号 = 1" & vbNewLine & _
                    "Order By 表单序号"
        Else
            strSql = "Select a.表单序号 表单序号, b.表单名称, b.表单表头, a.行数, a.列数" & vbNewLine & _
                    "From (Select 表单序号, Max(分类序号) 行数, Max(阶段序号) 列数 From 标准路径表单 Where 标准路径id = [1] Group By 表单序号) A, 标准路径表单 B" & vbNewLine & _
                    "Where b.标准路径id =[1] And a.表单序号 = b.表单序号 And b.分类序号 = 1 And b.阶段序号 = 1" & vbNewLine & _
                    "Order By 表单序号"
        End If

        Set mrs表头信息 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngId)
        
        '加载标准路径表单选项
        If mrs表头信息.RecordCount > 0 Then
            j = mrs表头信息.RecordCount
            For i = 1 To j
                mrs表头信息.Filter = "表单序号 =" & i
                Call tbcStPath.InsertItem(i, mrs表头信息!表单名称, picPathTable.Hwnd, 0)
            Next
            '读取表单数据
            If mintMode = 1 Then
                strSql = "Select  表单序号, 表单名称, 表单表头, 分类序号, 分类名称, 阶段序号, 阶段名称, 路径内容" & vbNewLine & _
                    "From   标准门诊路径表单" & vbNewLine & _
                    "where 标准路径id=[1]"
            Else
                strSql = "Select  表单序号, 表单名称, 表单表头, 分类序号, 分类名称, 阶段序号, 阶段名称, 路径内容" & vbNewLine & _
                    "From   标准路径表单" & vbNewLine & _
                    "where 标准路径id=[1]"
            End If
            Set mrs表单 = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngId)
        End If
    End If
    
    If lng序号 <> 0 Then
        '没有表单信息数据则不加载表单信息
        mrs表单.Filter = ""
        If mrs表单.RecordCount = 0 Then tbcStPath.Item(0).Selected = True: Exit Sub
        mrs表头信息.Filter = " 表单序号 =" & lng序号
        If mrs表头信息.RecordCount = 0 Then tbcStPath.Item(0).Selected = True: Exit Sub
        
        '加载表单表头
        lbl表头.Caption = ""
        lbl表头.Caption = vbNewLine & mrs表头信息!表单表头
        
        With vsPathTable
            .Redraw = False
            .Rows = 0
            .Cols = 0
            '确定行数
            lngColCount = Val(mrs表头信息!列数 & "") - 1
            '确定总行数
            lng表单行数 = Val(mrs表头信息!行数 & "")
            lngRowCount = IntEx(lngColCount / M_INT_STEPNUM) * lng表单行数 + IntEx(lngColCount / M_INT_STEPNUM) - 1
            If lngRowCount = 1 And lngColCount = 1 Then
                .Rows = 0
                .Cols = 0
                Call SetVsStyle
                Call picPathTable_Resize '由于lbl表头是autoSize的因此需要调用resize
                tbcStPath.Item(lng序号).Selected = True
                Exit Sub
            Else
                .Rows = lngRowCount
                .Cols = IIf(lngColCount > M_INT_STEPNUM, M_INT_STEPNUM + 1, lngColCount + 1)
            End If
    
            For k = 1 To IntEx(lngColCount / M_INT_STEPNUM)
                lngBeginRow = (k - 1) * lng表单行数 + (k - 1)
                For i = lngBeginRow To lngBeginRow + lng表单行数 - 1
                    For j = 0 To .Cols - 1
                        '每个表单表格区域的第一个单元格为时间
                        If i = lngBeginRow And j = 0 Then
                            .TextMatrix(i, j) = "时间"
                        Else
                            If Not (i = lngBeginRow Or j = 0) Then
                                strFilter = "表单序号=" & lng序号 & " and 分类序号=" & i - lngBeginRow + 1 & " and 阶段序号=" & (k - 1) * 3 + j + 1
                                mrs表单.Filter = strFilter
                                If mrs表单.RecordCount = 1 Then
                                    .TextMatrix(i, j) = Nvl(mrs表单!路径内容, " ")
                                    .TextMatrix(i, 0) = Replace(Replace(Replace(mrs表单!分类名称 & "", Chr(13), ""), Chr(10), ""), " ", "")
                                    .TextMatrix(lngBeginRow, j) = mrs表单!阶段名称 & ""
                                End If
                            End If
                        End If
                    Next
                Next
            Next
            
            Call SetVsStyle
            .Redraw = True
            Call picPathTable_Resize '由于lbl表头是autoSize的因此需要调用resize
        
            
        End With
    End If
    
    tbcStPath.Item(lng序号).Selected = True
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetStPathCourceFont(ByVal strTilePos As String)
'功能：对RichTextBox进行字体设置
'参数 strTilePos 记录标题行位置格式为，标题1起始位置，标题长度;标题2起始位置，标题长度
    Dim arrtmp As Variant, i As Long
    
    On Error Resume Next
    
    If Len(Trim(strTilePos)) = 0 Then Exit Sub
    arrtmp = Split(Trim(strTilePos), ";")

    With rtfPathCourse

        For i = LBound(arrtmp) To UBound(arrtmp)
            .SelStart = Split(arrtmp(i), ",")(0)
            .SelLength = Split(arrtmp(i), ",")(1)
            .SelFontSize = 14
            .SelFontName = "黑体"
            .SelBold = True
            .SelLength = 0
        Next

        .SelStart = 0 '光标移动到开始
    End With
End Sub

Private Sub SetVsStyle()
'功能：根据内容设置表单表格的单元格的高度与宽度,以及内容颜色等，以及单元格的合并等
    Dim i As Long, j As Long
    Dim lngmaxHeight As Long
    Dim strTmp As String

    With vsPathTable

        '修改分类名称，阶段，分类加粗居中
        .Cell(flexcpFontBold, 0, 0, .Rows - 1, 0) = True
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, 0) = 4 '居中
        .Cell(flexcpBackColor, 0, 0, .Rows - 1, 0) = &HE1FFE1

        .AutoResize = False
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .Cols - 1, False, 0) '自动调整大小
        '设置阶段字体，颜色，对齐方式
        For i = 0 To .Rows - 1
            If .TextMatrix(i, 0) = "时间" Then
                .Cell(flexcpAlignment, i, 0, i, .Cols - 1) = 4
                .Cell(flexcpFontBold, i, 0, i, .Cols - 1) = False '设置加粗前要先清除加粗
                .Cell(flexcpFontBold, i, 0, i, .Cols - 1) = True
                .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &HE1FFE1
            Else
                .Cell(flexcpAlignment, i, 1, i, .Cols - 1) = 0
            End If
        Next

        '获取同一行最高的单元格高度赋值给行高
        For i = 0 To .Rows - 1
            If .TextMatrix(i, 0) <> "" Then
                For j = 0 To .Cols - 1
                    If j = 0 Then
                        lngmaxHeight = ComputerLines(.TextMatrix(i, j))
                    Else
                        lngmaxHeight = IIf(lngmaxHeight > ComputerLines(.TextMatrix(i, j)), lngmaxHeight, ComputerLines(.TextMatrix(i, j)))
                    End If
                Next
                .RowHeight(i) = lngmaxHeight * Me.TextHeight("字") * 1.5
            Else
                For j = 0 To .Cols - 1
                    .TextMatrix(i, j) = " " '为了合并单元格
                Next
            End If
        Next
        '分割行单元格合并，以及边框颜色设置
        .MergeCells = flexMergeFree
        For i = 0 To .Rows - 1
            If .TextMatrix(i, 0) = " " Then
                Call .CellBorderRange(i, 0, i, .Cols - 1, &HFFFFFF, 1, 0, 1, 0, 1, 0)
                .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &HFFFFFF
                .MergeRow(i) = True
            End If
        Next
        For i = 1 To .Cols - 1
            .ColWidth(i) = 4000
        Next
        .ColWidth(0) = 1500
        '实现自由拖动列宽
        .FixedRows = 1
        Call .CellBorderRange(0, 0, 0, .Cols - 1, &H8000&, 0, 0, 1, 1, 1, 1)
    End With
End Sub

Private Function ComputerLines(ByVal strInput As String) As Long
'功能：计算输入文本中回车符的个数
'参数：  strInput   要计算回车符的字符串
'返回：   回车符的个数

    Dim strTmp As String
    Dim count  As Long, lngPos As Long, lngLen As Long
    
    lngPos = InStr(strInput, Chr(13))
    lngLen = Len(strInput)
    strTmp = strInput
    
    Do While lngPos <> 0
        If Trim(strTmp) = "" Then Exit Do
        If lngPos + 1 <= lngLen Then
            strTmp = Mid(strTmp, lngPos + 1)
            count = count + 1
            lngPos = InStr(strTmp, Chr(13))
            lngLen = Len(strTmp)
        End If
    Loop
    
    ComputerLines = count + 2
End Function

Private Sub txtInput_GotFocus()
    Call zlControl.TxtSelAll(txtInput)
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call FuncFindSTPath
        txtInput.SetFocus
    End If
End Sub

Private Sub txtInput_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strTip As String
    
    strTip = "按路径名称\疾病编码查找" & vbCrLf & "查找(Ctrl+F)" & vbCrLf & "查找下一个(F3)"
    zlCommFun.ShowTipInfo txtInput.Hwnd, strTip, True
End Sub


