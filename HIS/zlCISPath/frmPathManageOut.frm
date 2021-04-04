VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPathManageOut 
   AutoRedraw      =   -1  'True
   Caption         =   "门诊临床路径管理"
   ClientHeight    =   7950
   ClientLeft      =   225
   ClientTop       =   525
   ClientWidth     =   10890
   Icon            =   "frmPathManageOut.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7950
   ScaleWidth      =   10890
   Begin VB.TextBox txtFind 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   8385
      TabIndex        =   12
      Top             =   180
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog cdgFile 
      Left            =   3180
      Top             =   285
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picTemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   2250
      ScaleHeight     =   600
      ScaleWidth      =   660
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   210
      Visible         =   0   'False
      Width           =   660
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   1200
      Top             =   300
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
            Picture         =   "frmPathManageOut.frx":058A
            Key             =   "Path"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathManageOut.frx":0B24
            Key             =   "File"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathManageOut.frx":10BE
            Key             =   "branch"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathManageOut.frx":7920
            Key             =   "Merge"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6285
      Left            =   60
      ScaleHeight     =   6285
      ScaleWidth      =   3615
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1065
      Width           =   3615
      Begin XtremeReportControl.ReportControl rptPath 
         Height          =   3975
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   3375
         _Version        =   589884
         _ExtentX        =   5953
         _ExtentY        =   7011
         _StockProps     =   0
         BorderStyle     =   2
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin VSFlex8Ctl.VSFlexGrid vsFile 
         Height          =   810
         Left            =   225
         TabIndex        =   2
         Top             =   5265
         Width           =   3150
         _cx             =   5556
         _cy             =   1429
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
         BackColorSel    =   16571840
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   4
         GridLinesFixed  =   5
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   285
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmPathManageOut.frx":E182
         ScrollTrack     =   -1  'True
         ScrollBars      =   0
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
   End
   Begin VB.Frame fraLR 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5400
      Left            =   3720
      MousePointer    =   9  'Size W E
      TabIndex        =   6
      Top             =   1485
      Width           =   45
   End
   Begin XtremeSuiteControls.TabControl tbcContent 
      Height          =   4155
      Left            =   3930
      TabIndex        =   4
      Top             =   3225
      Width           =   6735
      _Version        =   589884
      _ExtentX        =   11880
      _ExtentY        =   7329
      _StockProps     =   64
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2385
      Left            =   3930
      ScaleHeight     =   2385
      ScaleWidth      =   6720
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   840
      Width           =   6720
      Begin VSFlex8Ctl.VSFlexGrid vsgIllness 
         Height          =   855
         Left            =   480
         TabIndex        =   13
         Top             =   1440
         Width           =   5535
         _cx             =   9763
         _cy             =   1508
         Appearance      =   0
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
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
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
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   0
         GridLinesFixed  =   0
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
         FillStyle       =   1
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
      Begin VB.Label lbl科室 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "●适用科室："
         Height          =   180
         Index           =   0
         Left            =   165
         TabIndex        =   10
         Top             =   555
         Width           =   1080
      End
      Begin VB.Label lbl科室 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "………………………………………………………………………………"
         Height          =   180
         Index           =   1
         Left            =   330
         MouseIcon       =   "frmPathManageOut.frx":E1BF
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   780
         Width           =   5475
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbl病种 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "●对应病种："
         Height          =   180
         Index           =   0
         Left            =   165
         TabIndex        =   8
         Top             =   1215
         Width           =   1080
      End
      Begin VB.Label lbl说明 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "说明：…………………………………………………………………………………"
         Height          =   180
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   6210
         WordWrap        =   -1  'True
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   7590
      Width           =   10890
      _ExtentX        =   19209
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPathManageOut.frx":E311
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16298
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
      Left            =   270
      Top             =   165
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPathManageOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mfrmDesign   As frmPathDesignOut
Attribute mfrmDesign.VB_VarHelpID = -1
Private WithEvents mfrmContent  As frmPathDesignOut
Attribute mfrmContent.VB_VarHelpID = -1
Private WithEvents mfrmEdit     As frmPathEditOut       '新增、修改路径的窗体
Attribute mfrmEdit.VB_VarHelpID = -1
Private mstr疾病编码            As String               '当前选择的临床路径的对应疾病编码
Private mstrPrivs               As String
Private mstrDictPrivs           As String
Private mlngModul               As Long
Private zlAppTool               As Object

Private Enum COL_LIST
    COL_ID = 0
    COL_图标 = 1
    COL_分支 = 2
    COL_行号 = 3
    COL_分类 = 4
    COL_编码 = 5
    COL_名称 = 6
    COL_适用性别 = 7
    COL_适用年龄 = 8
    COL_说明 = 9
    COL_通用 = 10
    COL_最新版本 = 11
    COL_拼音简码 = 12
End Enum

Private Sub FuncPathNew()
'功能: 新增门诊临床路径
    Dim str分类 As String

    If InStr(mstrPrivs, "不限制路径数量") > 0 Then
        
    ElseIf InStr(mstrPrivs, "30个以下路径") > 0 Then
        If rptPath.Records.count >= 30 Then
            MsgBox "已达到最大授权允许的路径数量，不允许再新增。", vbInformation, gstrSysName
            Exit Sub
        End If
    ElseIf InStr(mstrPrivs, "5个以下路径") > 0 Then
        If rptPath.Records.count >= 5 Then
            MsgBox "已达到最大授权允许的路径数量，不允许再新增。", vbInformation, gstrSysName
            Exit Sub
        End If
    Else
        MsgBox "不能明确授权允许的路径数量，请与系统管理员联系。", vbInformation, gstrSysName
        Exit Sub
    End If

    If rptPath.SelectedRows.count > 0 Then
        If rptPath.SelectedRows(0).GroupRow Then
            str分类 = rptPath.SelectedRows(0).Childs(0).Record(COL_分类).Value
        Else
            str分类 = rptPath.SelectedRows(0).Record(COL_分类).Value
        End If
    End If
    mfrmEdit.ShowEdit Me, mstrPrivs, , str分类                                      '事件中已刷新
End Sub

Private Sub FuncPathModify()
'功能: 修改门诊临床路径
    mfrmEdit.ShowEdit Me, mstrPrivs, rptPath.SelectedRows(0).Record(COL_ID).Value   '事件中已刷新
End Sub

Private Sub FuncPathDelete()
'功能: 删除门诊临床路径
    Dim strSql As String

    With rptPath.SelectedRows(0)
        If MsgBox("确实要删除临床路径""" & .Record(COL_名称).Value & """吗？", vbQuestion + vbDefaultButton2 + vbYesNo, gstrSysName) = vbNo Then Exit Sub

        strSql = "Zl_门诊路径目录_Delete(" & .Record(COL_ID).Value & ")"
        On Error GoTo errH
        zlDatabase.ExecuteProcedure strSql, Me.Caption
        
        Call RefreshData
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncPathFileDelete()
'功能: 删除门诊临床路径文件
    Dim strSql As String

    With vsFile
        If MsgBox("确实要删除文件""" & .TextMatrix(.Row, 1) & """吗？", vbQuestion + vbDefaultButton2 + vbYesNo, gstrSysName) = vbNo Then Exit Sub

        strSql = "Zl_门诊路径文件_Delete(" & rptPath.SelectedRows(0).Record(COL_ID).Value & ",'" & .TextMatrix(.Row, 1) & "')"
        On Error GoTo errH
        zlDatabase.ExecuteProcedure strSql, Me.Caption
        On Error GoTo 0

        .RemoveItem .Row
        .Height = .Height - .RowHeightMin
        Call picList_Resize
        Call vsFile_AfterRowColChange(0, 0, .Row, .Col)
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FuncPathFileView()
'功能：打开临床路径文件查看
    Dim strFile As String
    Dim lngRetu As Long, strInfo As String

    Screen.MousePointer = 11
    
    On Error GoTo errH
    strFile = gobjFile.GetSpecialFolder(TemporaryFolder) & "\" & vsFile.TextMatrix(vsFile.Row, 1)
    If gobjFile.FileExists(strFile) Then gobjFile.DeleteFile strFile, True

    strFile = Sys.ReadLob(glngSys, 26, rptPath.SelectedRows(0).Record(COL_ID).Value & "," & vsFile.TextMatrix(vsFile.Row, 1), strFile)
    If Not gobjFile.FileExists(strFile) Then
        MsgBox "文件内容读取失败！", vbInformation, gstrSysName:
        Screen.MousePointer = 0: Exit Sub
    End If

    lngRetu = ShellExecute(Me.Hwnd, "open", strFile, "", "", SW_SHOWNORMAL)
    If lngRetu <= 32 Then
        Select Case lngRetu
            Case 2
                strInfo = "错误的关联"
            Case 29
                strInfo = "关联失败"
            Case 30
                strInfo = "关联应用程式忙碌中..."
            Case 31
                strInfo = "没有关联任何应用程式"
            Case Else
                strInfo = "无法识别的错误"
        End Select
        MsgBox "文件打开时出错：" & vbCrLf & vbCrLf & vbTab & strInfo, vbExclamation, gstrSysName
    End If

    Screen.MousePointer = 0
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FuncPathFileNew(ByVal lngModle As Long)
'参数：lngModle=0 新增普通文件，=1 新增患者版路径表
    Dim arrSQL() As String
    Dim strFile As String
    Dim strFileName As String
    Dim i As Long
    Dim blnTrans As Boolean

    cdgFile.DialogTitle = "选择要添加的门诊临床路径文件"
    If lngModle = 0 Then
        cdgFile.Filter = "所有文件|*.*"
    Else
        cdgFile.Filter = "Word文档(*.doc;*.docx)|*.doc;*.docx"
        For i = 1 To vsFile.Rows - 1
            If vsFile.Cell(flexcpForeColor, i, 1) = &HFF0000 Then
                MsgBox "当前路径已经存在患者版路径表文件，请删除后再进行添加。", vbInformation, Me.Caption
                Exit Sub
            End If
        Next
    End If
    cdgFile.Flags = &H80000 Or &H4 Or &H1000 Or &H200000 Or &H800
    cdgFile.InitDir = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "门诊路径文件目录")
    cdgFile.CancelError = True
    On Error Resume Next
    cdgFile.ShowOpen
    If Err.Number <> 0 Then
        Err.Clear: Exit Sub
    End If
    On Error GoTo 0
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "门诊路径文件目录", gobjFile.GetFile(cdgFile.FileName).ParentFolder.Path
    strFile = cdgFile.FileName                              '包含路径
    strFileName = gobjFile.GetFile(cdgFile.FileName).Name

    '检查文件大小不超过3M
    If gobjFile.GetFile(strFile).Size / 1024 / 1024 > 3 Then
        MsgBox "文件尺寸太大(超过3M)，请对文件进行适当的整理后再添加。", vbInformation, gstrSysName
        Exit Sub
    End If

    Screen.MousePointer = 11

    ReDim arrSQL(0)
    arrSQL(0) = "Zl_门诊路径文件_Insert(" & rptPath.SelectedRows(0).Record(COL_ID).Value & ",'" & strFileName & "'," & lngModle & ")"
    If Not Sys.GetLobSql(glngSys, 26, rptPath.SelectedRows(0).Record(COL_ID).Value & "," & strFileName, strFile, arrSQL()) Then
        MsgBox "文件添加失败！", vbExclamation, gstrSysName
        Screen.MousePointer = 0
        Exit Sub
    End If

    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    For i = LBound(arrSQL) To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(arrSQL(i), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTrans = False
    On Error GoTo 0

    Call LoadPathFile(rptPath.SelectedRows(0).Record(COL_ID).Value) '刷新

    Screen.MousePointer = 0
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FuncPathTableOutputAll()
'功能：输出全部临床路径表到Excel
    Dim lngCount As Long, i As Long
    Dim objRow As ReportRow
    Dim objControl As CommandBarControl

    lngCount = rptPath.Records.count
    If lngCount = 0 Then
        MsgBox "当前没有可以输出的路径表。", vbInformation, gstrSysName
    Else
        If MsgBox("共有" & lngCount & "个路径表，你确定要全部输出到Excel吗？", vbQuestion + vbOKCancel + vbDefaultButton1, gstrSysName) = vbCancel Then
            Exit Sub
        End If
    End If

    For i = 0 To rptPath.Rows.count - 1
        If Not rptPath.Rows(i).GroupRow Then
            Set objRow = rptPath.Rows(i)
            Set rptPath.FocusedRow = objRow                 '该行选中且显示在可见区域,并引发SelectionChanged事件

            Set objControl = cbsMain.FindControl(, conMenu_File_Excel, True, True)
            If Not objControl Is Nothing Then
                Call mfrmContent.zlExecuteCommandBars(objControl, True)
            End If
        End If
    Next
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    Dim objRow As ReportRow, i As Long
    Dim lng路径ID As Long
    Dim blnTmp As Boolean
    Dim str分类 As String
    Dim str编码 As String
    Dim frmSub As Form

    If Control.ID <> 0 And Control.ID <> conMenu_View_FindNext Then
        If cbsMain.FindControl(, Control.ID, True, True) Is Nothing Then Exit Sub
    End If

    Select Case Control.ID
        Case conMenu_File_ExportToXML * 10# + 1    '批量导出XML
            Call FuncExportToXMLBatch
        Case conMenu_File_ExportToXML * 10# + 2    '批量导入XML
            Call FuncImportFromXMLBatch
        Case conMenu_File_ExportToXML * 10# + 3    '导入标准路径
            Set frmSub = New frmStandardPathRef
            If frmSub.ShowMe(gfrmMain, 0, 1, 1) Then
                '刷新
                Call RefreshData
            End If
        Case conMenu_File_BatPrint              '全部输出到Excel
            Call FuncPathTableOutputAll
        Case conMenu_Edit_NewItem               '新增
            Call FuncPathNew
        Case conMenu_Edit_Modify                '修改
            Call FuncPathModify
        Case conMenu_Edit_Delete                '删除
            Call FuncPathDelete
        Case conMenu_Edit_Archive * 10# + 1     '新增文件
            Call FuncPathFileNew(0)
        Case conMenu_Edit_Archive * 10# + 2     '新增文件
            Call FuncPathFileNew(1)
        Case conMenu_Edit_Archive * 10# + 3     '查看文件
            Call FuncPathFileView
        Case conMenu_Edit_Archive * 10# + 4     '删除文件
            Call FuncPathFileDelete
        Case conMenu_Tool_Define                '图标设置
            If frmIconManage.ShowMe(Me) Then
                Call rptPath_SelectionChanged
            End If
        Case conMenu_Tool_Option                '字典管理
            If zlAppTool Is Nothing Then Set zlAppTool = CreateObject("zl9AppTool.clsAppTool")
            Call zlAppTool.zlAppointDict("路径结果性质,路径常见结果,门诊变异常见原因", glngSys)
        Case conMenu_Edit_Report                '出径登记表
            Call frmPathOutDefinition.ShowMe(Me, rptPath.SelectedRows(0).Record(COL_ID).Value, rptPath.SelectedRows(0).Record(COL_名称).Value, 1)
        Case conMenu_Edit_Compend               '设计
            If InStr(mstrPrivs, "不限制路径数量") > 0 Then
                'Do Nothing
            ElseIf InStr(mstrPrivs, "30个以下路径") > 0 Then
                If rptPath.Records.count > 30 Then
                    MsgBox "已达到最大授权允许的路径数量，请先删除多余的临床路径。", vbInformation, gstrSysName
                    Exit Sub
                End If
            ElseIf InStr(mstrPrivs, "5个以下路径") > 0 Then
                If rptPath.Records.count > 5 Then
                    MsgBox "已达到最大授权允许的路径数量，请先删除多余的临床路径。", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
            Call mfrmDesign.ShowDesign(Me, rptPath.SelectedRows(0).Record(COL_ID).Value, mstrPrivs, mstr疾病编码)
            Call RefreshData    '为了设计界面新增版本后刷新路径目录中未审核路径的颜色
'        Case conMenu_Edit_Adjust  ' 辅助改进
'            If rptPath.SelectedRows(0).Record(COL_最新版本).Value = 0 Then
'                MsgBox "该路径为未审核启用过的新建路径,不能执行辅助改进功能。", vbInformation, gstrSysName
'                Exit Sub
'            End If
'            str分类 = rptPath.SelectedRows(0).Record(COL_分类).Value
'            Call frmPathImprove.ShowMe(Me, rptPath.SelectedRows(0).Record(COL_ID).Value, str分类, str编码, blnTmp)
'            If blnTmp Then Call RefreshData(str分类, str编码)  '刷新数据
'        Case conMenu_Edit_BatExecute                            '批量调整
'            Call frmPathItemBatReplace.ShowMe(Me, mstrPrivs)
        Case conMenu_View_Find                                  '查找
            If Me.ActiveControl Is txtFind Then
                txtFind.SetFocus                                '有时需要定位一下
                If txtFind.Text <> "" Then
                    Call FuncFindPath
                End If
            Else
                txtFind.SetFocus
            End If
        Case conMenu_View_FindNext          '查找下一个
            If txtFind.Text = "" Then
                txtFind.SetFocus
            Else
                Call FuncFindPath(True)
            End If
        Case conMenu_View_ToolBar_Button    '工具栏
            For i = 2 To cbsMain.count
                Me.cbsMain(i).Visible = Not Me.cbsMain(i).Visible
            Next
            Me.cbsMain.RecalcLayout
        Case conMenu_View_ToolBar_Text    '按钮文字
            For i = 2 To cbsMain.count
                For Each objControl In Me.cbsMain(i).Controls
                    objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                Next
            Next
            Me.cbsMain.RecalcLayout
        Case conMenu_View_ToolBar_Size    '大图标
            Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
            Me.cbsMain.RecalcLayout
        Case conMenu_View_StatusBar    '状态栏
            Me.stbThis.Visible = Not Me.stbThis.Visible
            Me.cbsMain.RecalcLayout
        Case conMenu_View_StPath    '查看标准路径参考
            Call frmStPathList.ShowMe(Me, mstr疾病编码, 1)
        Case conMenu_View_Expend_CurCollapse    '折叠当前组
            If rptPath.SelectedRows.count > 0 Then
                If rptPath.SelectedRows(0).GroupRow Then
                    rptPath.SelectedRows(0).Expanded = False
                ElseIf Not rptPath.SelectedRows(0).ParentRow Is Nothing Then
                    If rptPath.SelectedRows(0).ParentRow.GroupRow Then
                        rptPath.SelectedRows(0).ParentRow.Expanded = False
                    End If
                End If
            End If
            '因折叠定位到分组上,不会自动激活该事件
            Call rptPath_SelectionChanged
        Case conMenu_View_Expend_CurExpend    '展开当前组
            If rptPath.SelectedRows.count > 0 Then
                rptPath.SelectedRows(0).Expanded = True
            End If
        Case conMenu_View_Expend_AllCollapse    '折叠所有组
            For Each objRow In rptPath.Rows
                If objRow.GroupRow Then objRow.Expanded = False
            Next
            '因折叠定位到分组上,不会自动激活该事件
            Call rptPath_SelectionChanged
        Case conMenu_View_Expend_AllExpend    '展开所有组
            For Each objRow In rptPath.Rows
                If objRow.GroupRow Then objRow.Expanded = True
            Next
        Case conMenu_View_Refresh           '刷新
            Call RefreshData
        Case conMenu_Help_Web_Home          'Web上的中联
            Call zlHomePage(Me.Hwnd)
        Case conMenu_Help_Web_Forum         '中联论坛
            Call zlWebForum(Me.Hwnd)
        Case conMenu_Help_Web_Mail          '发送反馈
            Call zlMailTo(Me.Hwnd)
        Case conMenu_Help_About             '关于
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case conMenu_Help_Help              '帮助
            Call ShowHelp(App.ProductName, Me.Hwnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_File_Exit              '退出
            Unload Me
        Case Else
            If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
                If rptPath.SelectedRows.count > 0 Then
                    If Not rptPath.SelectedRows(0).GroupRow Then
                        lng路径ID = rptPath.SelectedRows(0).Record(COL_ID).Value
                    End If
                End If
                '执行发布到当前模块的报表
                Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, "路径ID=" & lng路径ID)
            Else
                Call mfrmContent.zlExecuteCommandBars(Control)
                Select Case Control.ID
                    Case conMenu_Edit_Audit, conMenu_Edit_Untread    '审核,取消审核
                        Call RefreshData
                    Case conMenu_Edit_Stop, conMenu_Edit_Reuse      '停用,取消停用
                        Call RefreshData
                End Select
            End If
    End Select
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long

    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)

    On Error Resume Next

    With Me.picList
        .Left = lngLeft: .Top = lngTop
        .Height = lngBottom - lngTop
    End With

    With Me.fraLR
        .Left = Me.picList.Left + Me.picList.Width
        .Top = Me.picList.Top
        .Height = Me.picList.Height
    End With

    With Me.PicInfo
        .Left = fraLR.Left + fraLR.Width
        .Top = fraLR.Top
        .Width = lngRight - .Left
    End With
    Call ResizeInfoPane

    With Me.tbcContent
        .Left = PicInfo.Left
        .Top = PicInfo.Top + PicInfo.Height
        .Width = PicInfo.Width
        .Height = lngBottom - .Top
    End With

    Me.Refresh
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnEnabled As Boolean
    Dim lng路径ID As Long

    If rptPath.SelectedRows.count > 0 Then
        If Not rptPath.SelectedRows(0).GroupRow Then
            lng路径ID = rptPath.SelectedRows(0).Record(COL_ID).Value
        End If
    End If

    Select Case Control.ID
        '单个导出XML在子窗体中判断
        Case conMenu_File_ExportToXML * 10# + 1    '批量导出XML
            If InStr(mstrPrivs, "导出XML") = 0 Then
                Control.Visible = False
            Else
                Control.Enabled = rptPath.Records.count > 0
            End If
        Case conMenu_File_ExportToXML * 10# + 2    '批量导入XML
            If InStr(mstrPrivs, "导入XML") = 0 Then
                Control.Visible = False
            End If
        Case conMenu_Edit_NewItem    '新增
            If InStr(mstrPrivs, "增删改") = 0 Then
                Control.Visible = False
            End If
        Case conMenu_Edit_Modify    '修改
            If InStr(mstrPrivs, "增删改") = 0 Then
                Control.Visible = False
            Else
                Control.Enabled = lng路径ID <> 0
            End If
        Case conMenu_Edit_Delete    '删除
            If InStr(mstrPrivs, "增删改") = 0 Then
                Control.Visible = False
            Else
                Control.Enabled = lng路径ID <> 0
            End If
        Case conMenu_Edit_Archive * 10# + 1, conMenu_Edit_Archive * 10# + 2    '增加文件,患者临床路径表
            If InStr(mstrPrivs, "增删改") = 0 Then
                Control.Visible = False
            Else
                Control.Enabled = lng路径ID <> 0
            End If
        Case conMenu_Edit_Archive * 10# + 3    '查看文件
            Control.Enabled = lng路径ID <> 0 And vsFile.Rows > vsFile.FixedRows And vsFile.Row >= vsFile.FixedRows
        Case conMenu_Edit_Archive * 10# + 4    '删除文件
            Control.Enabled = lng路径ID <> 0 And vsFile.Rows > vsFile.FixedRows And vsFile.Row >= vsFile.FixedRows
        Case conMenu_Edit_Report  '出径登记表
            If InStr(mstrPrivs, "出径登记表设计") = 0 Then
                Control.Visible = False
            End If
            blnEnabled = False
            If rptPath.SelectedRows.count > 0 Then
                If Not rptPath.SelectedRows(0).GroupRow Then
                   blnEnabled = True
                End If
            End If
            Control.Enabled = blnEnabled
'        Case conMenu_Edit_Adjust, conMenu_Edit_BatExecute '辅助改进,'批量调整
'            If Control.ID = conMenu_Edit_Adjust Then
'                If InStr(mstrPrivs, "临床路径表辅助改进") = 0 Then
'                    Control.Visible = False
'                End If
'            ElseIf Control.ID = conMenu_Edit_BatExecute Then
'                If InStr(mstrPrivs, "路径表设计") = 0 Then
'                    Control.Visible = False
'                End If
'            End If
'            If Control.Visible Then
'                If rptPath.SelectedRows.count > 0 Then
'                    If Not rptPath.SelectedRows(0).GroupRow Then
'                        Control.Enabled = True
'                    Else
'                        Control.Enabled = False
'                    End If
'                End If
'            End If
        Case conMenu_Tool_Define    '图标设置
            If InStr(mstrPrivs, "图标设置") = 0 Then
                Control.Visible = False
            End If
        Case conMenu_Tool_Option    '字典管理
            If InStr(mstrDictPrivs, "基本") = 0 Then
                Control.Visible = False
            End If
        Case conMenu_View_ToolBar_Button    '工具栏
            If cbsMain.count >= 2 Then
                Control.Checked = Me.cbsMain(2).Visible
            End If
        Case conMenu_View_ToolBar_Text    '图标文字
            If cbsMain.count >= 2 Then
                Control.Checked = Not (Me.cbsMain(2).Controls(1).Style = xtpButtonIcon)
            End If
        Case conMenu_View_ToolBar_Size    '大图标
            Control.Checked = Me.cbsMain.Options.LargeIcons
        Case conMenu_View_FindNext    '查找下一个
            Control.Visible = False
        Case conMenu_View_StatusBar    '状态栏
            Control.Checked = Me.stbThis.Visible
        Case conMenu_View_Expend_CurExpend    '展开当前组
            blnEnabled = False
            If rptPath.SelectedRows.count > 0 Then
                If rptPath.SelectedRows(0).GroupRow Then
                    blnEnabled = Not rptPath.SelectedRows(0).Expanded
                End If
            End If
            Control.Enabled = blnEnabled
        Case conMenu_View_Expend_CurCollapse    '折叠当前组
            blnEnabled = False
            If rptPath.SelectedRows.count > 0 Then
                If rptPath.SelectedRows(0).GroupRow Then
                    blnEnabled = rptPath.SelectedRows(0).Expanded
                ElseIf Not rptPath.SelectedRows(0).ParentRow Is Nothing Then
                    If rptPath.SelectedRows(0).ParentRow.GroupRow Then
                        blnEnabled = rptPath.SelectedRows(0).ParentRow.Expanded
                    End If
                End If
            End If
            Control.Enabled = blnEnabled
        Case conMenu_View_Expend    '折叠/展开组
            Control.Enabled = rptPath.GroupsOrder.count > 0 And rptPath.Rows.count > 0
        Case Else
            Call mfrmContent.zlUpdateCommandBars(Control)
    End Select
End Sub

Private Sub Form_Load()
    mstrPrivs = gstrPrivs
    mlngModul = glngModul
    '获取字典管理权限串
    mstrDictPrivs = GetPrivFunc(0, 11)
    Call zlCommFun.SetWindowsInTaskBar(Me.Hwnd, False)

    Set mfrmEdit = New frmPathEditOut
    Set mfrmDesign = New frmPathDesignOut
    Set mfrmContent = New frmPathDesignOut
    'ReportControl
    '-----------------------------------------------------
    Call InitReportColumn
    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True    '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    Call MainDefCommandBar

    'TabControl
    '-----------------------------------------------------
    With Me.tbcContent
        With .PaintManager
            .Appearance = xtpTabAppearanceVisio
            .Color = xtpTabColorOffice2003
        End With
        .InsertItem 0, "临床路径表", mfrmContent.Hwnd, 0
    End With

    '附件表格
    '-----------------------------------------------------
    With vsFile
        .RowHeight(0) = 315
        Set .Cell(flexcpPicture, 0, 0) = img16.ListImages("File").Picture
        .Cell(flexcpPictureAlignment, 0, 0) = 7
        .TextMatrix(0, 1) = "路径文件附件,蓝色表示路径表(患者版)"
        .Cell(flexcpFontBold, 0, 1) = True
    End With

    '对应病种
    '---------------------------------------------------------
    Call InitVsgIllness                         '初始化对应病种的VSF控件
    '
    Call RestoreWinState(Me, App.ProductName)

    Call RefreshData                            '根据当前设置的条件读取临床路径目录数据
End Sub

Private Sub MainDefCommandBar()
'功能：主窗口菜单定义部份
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom
    Dim lngCount As Long

    '菜单定义
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")
        Set objControl = .Add(xtpControlButton, conMenu_File_BatPrint, "全部输出到&Excel…")
        Set objControl = .Add(xtpControlButton, conMenu_File_ExportToXML, "导出&XML文件…")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_ExportToXML * 10# + 1, "批量导出XML文件…")
        Set objControl = .Add(xtpControlButton, conMenu_File_ExportToXML * 10# + 2, "批量导入XML文件…")
        'Set objControl = .Add(xtpControlButton, conMenu_File_ExportToXML * 10# + 3, "导入标准路径")
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)")
        objControl.BeginGroup = True
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增(&N)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改(&M)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&D)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_Archive, "文件(&F)")
        objPopup.BeginGroup = True
        objPopup.IconId = conMenu_Manage_Report
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Archive * 10# + 1, "添加普通文件(&1)")
            objControl.IconId = conMenu_Edit_NewItem
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Archive * 10# + 2, "添加路径表(患者版)(&2)")
            objControl.IconId = 3205
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Archive * 10# + 3, "查看文件(&3)")
            objControl.IconId = conMenu_Tool_Search
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Archive * 10# + 4, "删除文件(&4)")
            objControl.IconId = conMenu_Edit_Delete
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Compend, "设计(&S)")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Report, "出径登记表设计(&P)")
'        Set objControl = .Add(xtpControlButton, conMenu_Edit_Adjust, "辅助改进(&U)")
'        objControl.BeginGroup = True
'        Set objControl = .Add(xtpControlButton, conMenu_Edit_BatExecute, "批量调整(&B)")
'        objControl.IconId = conMenu_Apply_AllCard
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Audit, "审核(&A)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "取消审核(&R)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Stop, "停用(&S)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "取消停用(&Z)")
        objControl.IconId = conMenu_Edit_Untread
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "工具(&T)", -1, False)
    objMenu.ID = conMenu_ToolPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Define, "图标设置(&I)")
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Option, "字典管理(&D)")
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)")
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_StPath, "标准路径参考")
        objControl.BeginGroup = True
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_Expend, "展开/折叠组(&X)"):
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_AllCollapse, "折叠所有组(&L)", -1, False)
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_AllExpend, "展开所有组(&X)", -1, False)
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_CurCollapse, "折叠当前组(&C)", -1, False)
            objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_CurExpend, "展开当前组(&E)", -1, False)
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_FindNext, "查找下一个(&C)")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)")
        objControl.BeginGroup = True
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "论坛(&F)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…")
        objControl.BeginGroup = True
    End With

    '主菜单右侧的查找
    With cbsMain.ActiveMenuBar.Controls
        Set objControl = .Add(xtpControlLabel, 0, "查找")
        objControl.IconId = conMenu_View_Find
        objControl.Flags = xtpFlagRightAlign
        Set objCustom = .Add(xtpControlCustom, conMenu_View_Find, "")
        objCustom.Handle = txtFind.Hwnd
        objCustom.Flags = xtpFlagRightAlign
    End With

    '工具栏定义:包括公共部份
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Compend, "设计")
'        Set objControl = .Add(xtpControlButton, conMenu_Edit_Adjust, "辅助改进")
'        objControl.BeginGroup = True
'        Set objControl = .Add(xtpControlButton, conMenu_Edit_BatExecute, "批量调整")
'        objControl.IconId = conMenu_Apply_AllCard
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Audit, "审核"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "取消审核")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Stop, "停用"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "取消停用")
        objControl.IconId = conMenu_Edit_Untread
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With

    '设置一些公共的热键绑定
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyA, conMenu_Edit_NewItem                     '新增
        .Add FCONTROL, vbKeyM, conMenu_Edit_Modify                      '修改
        .Add FSHIFT, vbKeyDelete, conMenu_Edit_Delete                   '删除
        .Add FCONTROL, vbKeyD, conMenu_Edit_Compend                     '设计/修订
        .Add FCONTROL, vbKeyU, conMenu_Edit_Audit                       '审核
        .Add FCONTROL, vbKeyR, conMenu_Edit_Stop                        '停用
        .Add FCONTROL, vbKeyF, conMenu_View_Find                        '查找
        .Add 0, vbKeyF3, conMenu_View_FindNext                          '查找下一个
        .Add FCONTROL, vbKeyAdd, conMenu_View_Expend_AllExpend          '展开所有组
        .Add FCONTROL, vbKeySubtract, conMenu_View_Expend_AllCollapse   '折叠所有组
        .Add FCONTROL, vbKeyP, conMenu_File_Print                       '打印
        .Add 0, vbKeyF5, conMenu_View_Refresh                           '刷新
        .Add 0, vbKeyF1, conMenu_Help_Help                              '帮助
    End With

    '设置一些公共的不常用命令
    '-----------------------------------------------------
    With cbsMain.Options
        .AddHiddenCommand conMenu_File_PrintSet         '打印设置
        .AddHiddenCommand conMenu_File_Excel            '输出到Excel
    End With

    '恢复及固定的一些菜单设置
    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.SetIconSize 16, 16
    For lngCount = 2 To cbsMain.count
        cbsMain(lngCount).ContextMenuPresent = False
        cbsMain(lngCount).ShowTextBelowIcons = False
        cbsMain(lngCount).EnableDocking xtpFlagHideWrap
        For Each objControl In cbsMain(lngCount).Controls
            objControl.Style = xtpButtonIconAndCaption
        Next
    Next

    '读取发布到该模块的报表(不含虚拟模块的)
    '-----------------------------------------------------
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModul, mstrPrivs)
End Sub

Private Sub InitReportColumn()
'初始化ReportControl控件
    Dim objCol As ReportColumn

    With rptPath
        '当列顺序或数量(代码或人为隐藏)改变后,要用Find(列号)或ItemIndex查找列,但仍可用Record(列号)访问数据行
        Set objCol = .Columns.Add(COL_ID, "", 0, False)
        objCol.Visible = False
        Set objCol = .Columns.Add(COL_图标, "", 18, False)
        objCol.Sortable = False
        objCol.AllowDrag = False
        objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Columns.Add(COL_分支, "", 18, False)
        objCol.Sortable = False
        objCol.AllowDrag = False
        objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Columns.Add(COL_行号, "行号", 35, True)
        objCol.AllowDrag = False
        objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Columns.Add(COL_分类, "分类", 80, True)
        objCol.Visible = False
        Set objCol = .Columns.Add(COL_编码, "编码", 35, True)
        objCol.Groupable = False
        Set objCol = .Columns.Add(COL_名称, "名称", 150, True)
        objCol.Groupable = False
        Set objCol = .Columns.Add(COL_适用性别, "适用性别", 55, True)
        objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Columns.Add(COL_适用年龄, "适用年龄", 55, True)
        Set objCol = .Columns.Add(COL_说明, "", 0, False)
        objCol.Visible = False
        Set objCol = .Columns.Add(COL_通用, "", 0, False)
        objCol.Visible = False
        Set objCol = .Columns.Add(COL_最新版本, "", 0, False)
        objCol.Visible = False
        Set objCol = .Columns.Add(COL_拼音简码, "", 0, False)
        objCol.Visible = False

        For Each objCol In .Columns
            objCol.Editable = False
        Next

        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的临床路径..."
        End With
        .AutoColumnSizing = False
        .AllowColumnRemove = False
        .ShowGroupBox = True
        .ShowItemsInGroups = False
        .PreviewMode = True
        .MultipleSelection = False              '会引发SelectionChanged事件
        .SetImageList Me.img16

        .GroupsOrder.Add .Columns(COL_分类)
        .GroupsOrder(0).SortAscending = True    '分组之后,如果分组列不显示,分组列的排序是不变的

        '分组之后可能失去记录集中的顺序,因此强行加入排序列
        .SortOrder.Add .Columns(COL_分类)
        .SortOrder(0).SortAscending = True
        .SortOrder.Add .Columns(COL_编码)
        .SortOrder(1).SortAscending = True
    End With
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    Call cbsMain_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)

    Unload mfrmDesign
    Set mfrmDesign = Nothing
    
    Unload mfrmContent
    Set mfrmContent = Nothing

    Unload mfrmEdit
    Set mfrmEdit = Nothing
End Sub

Private Sub fraLR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next

    If Button = 1 Then
        If picList.Width + X < 2000 Or PicInfo.Width - X < 3000 Then Exit Sub

        fraLR.Left = fraLR.Left + X
        picList.Width = picList.Width + X

        Call Form_Resize
    End If
End Sub

Private Sub lbl科室_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'科室，鼠标移上去之后加上下划线，颜色变蓝
    If Index = 1 And lbl科室(Index).Caption <> "" Then
        Me.lbl科室(1).Font.Underline = True
        Me.lbl科室(1).ForeColor = RGB(0, 0, 128)
    End If
End Sub

Private Sub mfrmDesign_DataChanged(ByVal 路径ID As Long)
'刷新路径表信息
    Call mfrmContent.zlRefresh(路径ID, mstrPrivs, lbl科室(1).Caption, vsgIllness.Tag)
End Sub

Private Sub mfrmEdit_AfterSave(ByVal 分类 As String, ByVal 编码 As String)
    Call RefreshData(分类, 编码)
End Sub

Private Sub picInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'鼠标移上去会后取消下划线和颜色恢复
    Me.lbl科室(1).Font.Underline = False
    Me.lbl科室(1).ForeColor = lbl科室(0).ForeColor
    vsgIllness.FontUnderline = False
    vsgIllness.ForeColor = lbl病种(0).ForeColor
End Sub

Private Sub picInfo_Resize()
    On Error Resume Next
    
    lbl说明.Width = PicInfo.ScaleWidth - lbl说明.Left * 2
    lbl科室(1).Width = PicInfo.ScaleWidth - lbl科室(1).Left - lbl说明.Left
    vsgIllness.Left = lbl病种(0).Left
    vsgIllness.Width = PicInfo.ScaleWidth - vsgIllness.Left - lbl说明.Left
End Sub

Private Sub picList_Resize()
    On Error Resume Next

    rptPath.Left = 0
    rptPath.Top = 0
    rptPath.Width = picList.ScaleWidth
    rptPath.Height = picList.ScaleHeight - vsFile.Height

    vsFile.Left = 0
    vsFile.Top = rptPath.Top + rptPath.Height
    vsFile.Width = picList.ScaleWidth
End Sub

Private Sub rptPath_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim objControl As CommandBarControl
    Dim objRow As ReportRow

    If KeyCode = vbKeyReturn And Shift = 0 Then
        If rptPath.SelectedRows.count > 0 Then
            If Not rptPath.SelectedRows(0).GroupRow Then
                Set objRow = rptPath.SelectedRows(0)
            End If
        End If
        If Not objRow Is Nothing Then
            Set objControl = cbsMain.FindControl(, conMenu_Edit_Modify, True, True)
            If Not objControl Is Nothing Then objControl.Execute
        End If
    End If
End Sub

Private Sub rptPath_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim objHitTest As ReportHitTestInfo
    Dim objPopup As CommandBarPopup

    If Button = 2 Then
        Set objHitTest = rptPath.HitTest(X, Y)
        If objHitTest.ht = xtpHitTestReportArea And Not objHitTest.Row Is Nothing Then
            If objHitTest.Row.GroupRow Then
                Set objPopup = cbsMain.FindControl(, conMenu_View_Expend, , True)
            ElseIf objHitTest.Row.Childs.count = 0 Then
                Set objPopup = cbsMain.ActiveMenuBar.Controls(2)
            End If
        End If

        rptPath.SetFocus
        If Not objPopup Is Nothing Then objPopup.CommandBar.ShowPopup
    End If
End Sub

Private Sub rptPath_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Dim objControl As CommandBarControl

    If Not Row.GroupRow Then
        Set objControl = cbsMain.FindControl(, conMenu_Edit_Modify, True, True)
        If Not objControl Is Nothing Then objControl.Execute
    End If
End Sub

Private Sub rptPath_SelectionChanged()
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, strTmp As String
    Dim arrStr As Variant
    Dim intRowNum As Integer: Dim intColNum As Integer
    Dim i As Long

    On Error GoTo errH

    If rptPath.SelectedRows.count = 0 Then
        Call ClearSubData
    ElseIf rptPath.SelectedRows(0).GroupRow Then
        Call ClearSubData
    Else
        With rptPath.SelectedRows(0)
            lbl说明.Caption = "说明：" & .Record(COL_说明).Value

            '对应科室信息
            If .Record(COL_通用).Value = 1 Then
                lbl科室(1).Caption = "该临床路径适用于所有门诊临床科室"
            Else
                strSql = "Select B.编码,B.名称 From 门诊路径科室 A,部门表 B Where A.科室ID=B.ID And A.路径ID=[1] Order by B.编码"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(.Record(COL_ID).Value))
                strTmp = ""
                Do While Not rsTmp.EOF
                    strTmp = strTmp & "," & rsTmp!编码 & "-" & rsTmp!名称
                    rsTmp.MoveNext
                Loop
                If strTmp <> "" Then
                    lbl科室(1).Caption = Mid(strTmp, 2)
                Else
                    lbl科室(1).Caption = "<该临床路径尚未设置所适用的科室>"
                End If
            End If

            '对应病种信息
            vsgIllness.Clear

            strSql = " Select Decode(B.编码,NULL,'['||C.编码||']'||C.名称,'['||B.编码||']'||B.名称) as 名称 ,B.编码 " & _
                     " From 门诊路径病种 A,疾病编码目录 B,疾病诊断目录 C" & _
                     " Where A.疾病ID=B.ID(+) And A.诊断ID=C.ID(+) And A.路径ID=[1] " & _
                     " Order by B.编码,C.编码"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(.Record(COL_ID).Value))
            strTmp = ""
            mstr疾病编码 = ""
            Do While Not rsTmp.EOF
                strTmp = strTmp & "," & rsTmp!名称
                If rsTmp!编码 & "" <> "" Then
                    mstr疾病编码 = mstr疾病编码 & "," & rsTmp!编码
                End If
                rsTmp.MoveNext
            Loop
            If strTmp <> "" Then
                With vsgIllness
                    arrStr = Split(Mid(strTmp, 2), ",")
                    .Cols = 3: .Rows = ((UBound(arrStr) + 1) + (.Cols - 1)) \ .Cols
                    .Tag = Mid(strTmp, 2)
                    For i = 0 To UBound(arrStr)
                        intRowNum = i \ .Cols
                        intColNum = i Mod .Cols
                        .TextMatrix(intRowNum, intColNum) = arrStr(i)
                    Next i
                    mstr疾病编码 = Mid(mstr疾病编码, 2)
                End With
            Else
                vsgIllness.Rows = 1: vsgIllness.Cols = 1
                vsgIllness.TextMatrix(0, 0) = "<该临床路径尚未设置所对应的病种>"
            End If

            '对应文件信息
            Call LoadPathFile(Val(.Record(COL_ID).Value))

            '路径表信息
            Call mfrmContent.zlRefresh(Val(.Record(COL_ID).Value), mstrPrivs, lbl科室(1).Caption, vsgIllness.Tag)
        End With

        Call Form_Resize
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function LoadPathFile(ByVal lng路径ID As Long) As Boolean
'功能：显示临床路径文件内容和患者临床路径
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, i As Long

    On Error GoTo errH

    strSql = "Select 文件名,创建人,创建时间,类别 From 门诊路径文件 Where 路径ID=[1] Order by 创建时间"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng路径ID)
    With vsFile
        .Rows = .FixedRows '先清除
        .Rows = .FixedRows + rsTmp.RecordCount
        .Height = .RowHeight(0) + .RowHeightMin * (.Rows - 1) + Screen.TwipsPerPixelY * 2
        For i = 1 To rsTmp.RecordCount
            .TextMatrix(i, 1) = rsTmp!文件名
            Set .Cell(flexcpPicture, i, 0) = zlCommFun.GetFileIcon(rsTmp!文件名, True, App.hInstance)
            .Cell(flexcpPictureAlignment, i, 0) = 7

            '删除之前的临时文件
            If gobjFile.FileExists(gobjFile.GetSpecialFolder(TemporaryFolder) & "\" & rsTmp!文件名) Then
                gobjFile.DeleteFile gobjFile.GetSpecialFolder(TemporaryFolder) & "\" & rsTmp!文件名, True
            End If
            If rsTmp!类别 = 1 Then .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = &HFF0000
            rsTmp.MoveNext
        Next
    End With

    Call picList_Resize
    LoadPathFile = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ResizeInfoPane()
'功能：根据当前信息内容，调整信息面板内信息项及面板尺寸和位置
'说明：利用Label的AutoSize属性自动调整标签高度
    lbl科室(0).Top = lbl说明.Top + lbl说明.Height + Screen.TwipsPerPixelY * 6
    lbl科室(1).Top = lbl科室(0).Top + lbl科室(0).Height + Screen.TwipsPerPixelY * 3
    lbl病种(0).Top = lbl科室(1).Top + lbl科室(1).Height + Screen.TwipsPerPixelY * 6

    vsgIllness.Top = lbl病种(0).Top + lbl病种(0).Height + Screen.TwipsPerPixelY * 3
    '根据对应病种行数动态显示对应病种信息，最大显示5行
    vsgIllness.Height = vsgIllness.RowHeightMin * IIf(vsgIllness.Rows > 5, 5, vsgIllness.Rows)
    vsgIllness.ColWidthMin = vsgIllness.Width / vsgIllness.Cols
    PicInfo.Height = vsgIllness.Top + vsgIllness.Height + Screen.TwipsPerPixelY * 6
End Sub

Private Function RefreshData(Optional ByVal str分类 As String, Optional ByVal str编码 As String) As Boolean
'功能：根据当前设置的条件读取临床路径目录数据
'参数：用于定位
    Dim rsTmp       As ADODB.Recordset
    Dim strSql      As String
    Dim objRecord   As ReportRecord
    Dim objItem     As ReportRecordItem
    Dim objRow As ReportRow, i As Long
    Dim lngPreID As Long, lngPreIdx As Long
    Dim intTypeNum  As Integer
    Dim lngPathColor As Long                '未审核路径目录前景颜色值

    Screen.MousePointer = 11

    On Error GoTo errH

    'SQL中不排序提高效率,ReportControl有排序处理
    strSql = "Select Distinct a.Id, a.分类, a.编码, a.名称, a.适用性别, a.适用年龄, a.说明, a.通用, a.最新版本, Min(Decode(c.审核时间, Null, 0, 1)) As 已审核" & vbNewLine & _
             "From 门诊路径目录 A, 门诊路径版本 C" & vbNewLine & _
             "Where a.Id = c.路径id(+)"

    If InStr(mstrPrivs, "全院路径") = 0 Then
        '没有权限时，只能对只应用于本科的路径进行处理
        strSql = strSql & _
                 " And A.通用 = 2 And Exists" & vbNewLine & _
                 "      (Select 1 From 部门人员 C,门诊路径科室 D  " & vbNewLine & _
                 "       Where C.人员id = [1] And D.科室id = C.部门id And 路径id = A.ID)"
    End If

    strSql = strSql & " Group By a.Id, a.分类, a.编码, a.名称, a.适用性别, a.适用年龄, a.说明, a.通用, a.最新版本 "

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.ID)

    '记录现在选中的反馈
    If rptPath.SelectedRows.count > 0 Then
        If Not rptPath.SelectedRows(0).GroupRow Then
            lngPreIdx = rptPath.SelectedRows(0).Index    '用于快速重新定位
            lngPreID = rptPath.SelectedRows(0).Record(COL_ID).Value
        End If
    End If

    rptPath.Records.DeleteAll
    Do While Not rsTmp.EOF
        Set objRecord = Me.rptPath.Records.Add()
        Set objItem = objRecord.AddItem(Val(rsTmp!ID))
        Set objItem = objRecord.AddItem("")
        objItem.Icon = img16.ListImages("Path").Index - 1
        Set objItem = objRecord.AddItem("")
        Set objItem = objRecord.AddItem("")
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!分类, "<未指定分类>")))
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!编码)))
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!名称)))
        Set objItem = objRecord.AddItem(CStr(Decode(NVL(rsTmp!适用性别, 0), 0, "", 1, "男", 2, "女")))
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!适用年龄)))
        Set objItem = objRecord.AddItem(CStr("" & rsTmp!说明))
        Set objItem = objRecord.AddItem(Val(NVL(rsTmp!通用, 1)))
        Set objItem = objRecord.AddItem(Val(NVL(rsTmp!最新版本, 0)))
        Set objItem = objRecord.AddItem(zlCommFun.SpellCode(NVL(rsTmp!名称) & "※0"))

        lngPathColor = IIf(Val(rsTmp!已审核) = 1, vbBlack, &H80&)
        For i = COL_行号 To COL_通用
            objRecord.Item(i).ForeColor = lngPathColor
        Next

        rsTmp.MoveNext
    Loop

    rptPath.Populate

    '分类有多个时，显示行号列
    If rptPath.Rows.count - rptPath.Records.count > 1 Then
        rptPath.Columns(COL_行号).Visible = True
        rptPath.Columns(COL_行号).SortAscending = True
    Else
        rptPath.Columns(COL_行号).Visible = False
    End If

    '行号赋值
    For i = 0 To rptPath.Rows.count - 1
        With rptPath.Rows(i)
            If .GroupRow Then intTypeNum = intTypeNum + 1
            If Not .GroupRow Then
                .Record(COL_行号).Value = i - intTypeNum + 1
            End If
        End With
    Next

    If rptPath.Rows.count = 0 Then
        Call ClearSubData
    Else
        If str分类 <> "" And str编码 <> "" Then
            For i = 0 To rptPath.Rows.count - 1
                If Not rptPath.Rows(i).GroupRow Then
                    If rptPath.Rows(i).Record(COL_分类).Value = str分类 _
                       And rptPath.Rows(i).Record(COL_编码).Value = str编码 Then
                        Set objRow = rptPath.Rows(i): Exit For
                    End If
                End If
            Next
        Else
            If lngPreID <> 0 Then
                '先快速定位
                If lngPreIdx <= rptPath.Rows.count - 1 Then
                    If Not rptPath.Rows(lngPreIdx).GroupRow Then
                        If rptPath.Rows(lngPreIdx).Record(COL_ID).Value = lngPreID Then
                            Set objRow = rptPath.Rows(lngPreIdx)
                        End If
                    End If
                End If
                '再进行查找
                If objRow Is Nothing Then
                    For i = 0 To rptPath.Rows.count - 1
                        If Not rptPath.Rows(i).GroupRow Then
                            If rptPath.Rows(i).Record(COL_ID).Value = lngPreID Then
                                Set objRow = rptPath.Rows(i): Exit For
                            End If
                        End If
                    Next
                End If
            End If
            '取第一个非分组行
            If objRow Is Nothing Then
                For i = 0 To rptPath.Rows.count - 1
                    If Not rptPath.Rows(i).GroupRow Then Set objRow = rptPath.Rows(i): Exit For
                Next
            End If
        End If

        Set rptPath.FocusedRow = objRow    '该行选中且显示在可见区域,并引发SelectionChanged事件
        Me.stbThis.Panels(2).Text = "共有 " & rptPath.Records.count & " 个临床路径"
    End If

    Screen.MousePointer = 0
    RefreshData = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ClearSubData()
'清楚原有的数据信息
    Dim i As Integer

    lbl说明.Caption = "说明："

    lbl科室(1).Caption = ""

    vsgIllness.Rows = 0
    vsgIllness.Rows = 5
    vsFile.Rows = vsFile.FixedRows
    vsFile.Height = vsFile.RowHeight(0) + vsFile.RowHeightMin * (vsFile.Rows - 1) + Screen.TwipsPerPixelY * 2

    Me.stbThis.Panels(2).Text = ""

    Call mfrmContent.zlRefresh(0, mstrPrivs, lbl科室(1).Caption, vsgIllness.Tag)

    Call Form_Resize
    Call picList_Resize
End Sub

Private Sub txtFind_GotFocus()
    Call zlControl.TxtSelAll(txtFind)
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
'按下Enter键进行查找
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call FuncFindPath
    End If
End Sub

Private Sub FuncFindPath(Optional ByVal blnNext As Boolean)
'参数：blnNext=是否查找下一个
    Static blnReStart As Boolean
    Dim blnHave As Boolean, i As Long

    Call zlControl.TxtSelAll(txtFind)
    '开始查找行
    If rptPath.SelectedRows.count > 0 Then blnHave = True
    If Not blnNext Or blnReStart Or Not blnHave Then
        i = 0    'ReportControl的索引从是0开始
    Else
        i = rptPath.SelectedRows(0).Index + 1
    End If

    '查找路径
    For i = i To rptPath.Rows.count - 1
        With rptPath.Rows(i)
            If Not .GroupRow Then
                '根据输入的内容判断，如果是字母则查找简码，是汉字则查找名称，是数字则查找行号
                If zlCommFun.IsNumOrChar(Trim(txtFind.Text)) Then
                    '字母或数字
                    If zlCommFun.IsCharAlpha(Trim(txtFind.Text)) Then
                        '全是字母查找拼音简码
                        If .Record(COL_拼音简码).Value Like "*" & UCase(Trim(txtFind.Text)) & "*" Then
                            Exit For
                        End If
                    Else
                        '全是数字查找行号
                        If .Record(COL_行号).Value Like "*" & Trim(txtFind.Text) & "*" Then
                            Exit For
                        End If
                    End If
                ElseIf zlCommFun.IsCharChinese(Trim(txtFind.Text)) Then
                    '包含汉字 查找名称
                    If .Record(COL_名称).Value Like "*" & Trim(txtFind.Text) & "*" Then
                        Exit For
                    End If
                End If
            End If
        End With
    Next

    If i <= rptPath.Rows.count - 1 Then
        blnReStart = False
        '该行选中且显示在可见区域,并引发SelectionChanged事件
        Set rptPath.FocusedRow = rptPath.Rows(i)

        If rptPath.Visible Then rptPath.SetFocus
    Else
        blnReStart = True
        MsgBox IIf(blnNext, "后面已", "") & "找不到符合条件的门诊临床路径。", vbInformation, gstrSysName
    End If
End Sub

Private Sub txtFind_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'鼠标经过查找框的时候浮现的提示
    Dim strTip As String

    strTip = "查找(Ctrl+F)" & vbCrLf & "查找下一个(F3)"
    zlCommFun.ShowTipInfo txtFind.Hwnd, strTip, True
End Sub

Private Sub vsFile_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    On Error Resume Next
    vsFile.ForeColorSel = vsFile.Cell(flexcpForeColor, NewRow, 0)
End Sub

Private Sub vsFile_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If NewCol = 0 Then Cancel = True
End Sub

Private Sub vsFile_DblClick()
'鼠标双击查看
    If vsFile.MouseRow >= vsFile.FixedRows Then
        Call vsFile_KeyPress(13)
    End If
End Sub

Private Sub vsFile_KeyDown(KeyCode As Integer, Shift As Integer)
'按下Delete键删除
    If KeyCode = vbKeyDelete Then
        If vsFile.Row >= vsFile.FixedRows Then Call FuncPathFileDelete
    End If
End Sub

Private Sub vsFile_KeyPress(KeyAscii As Integer)
'按下Enter键查看
    If KeyAscii = 13 Then
        KeyAscii = 0
        If vsFile.Row >= vsFile.FixedRows Then Call FuncPathFileView
    End If
End Sub

Private Sub vsFile_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'弹出右键菜单
    Dim lngRow As Long
    Dim objPopup As CommandBarPopup

    lngRow = vsFile.MouseRow

    If Button = 2 And lngRow <> -1 Then
        vsFile.SetFocus
        If lngRow <= vsFile.Rows - 1 And lngRow >= vsFile.FixedRows Then
            vsFile.Row = lngRow
        End If

        Set objPopup = cbsMain.FindControl(, conMenu_Edit_Archive, True, True)
        If Not objPopup Is Nothing Then objPopup.CommandBar.ShowPopup
    End If
End Sub

Private Sub FuncExportToXMLBatch()
'功能：批量导出为XML文件
    Dim strPath As String, strFile As String
    Dim strFail As String, intCount As Integer
    Dim strMsg As String, i As Long

    If MsgBox("本功能将导出所有临床路径的最新已审核版本，" & _
        vbCrLf & "每个导出文件的命名规则为""分类-路径名称.xml""，" & vbCrLf & "如果在导出目标位置有相同名称的文件则将被覆盖。" & _
        vbCrLf & vbCrLf & "要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub

    strPath = zlCommFun.OpenDir(Me.Hwnd, "门诊临床路径导出目录", GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "门诊临床路径XML目录"))
    If strPath = "" Then Exit Sub
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "门诊临床路径XML目录", strPath

    Screen.MousePointer = 11
    For i = 0 To rptPath.Records.count - 1
        With rptPath.Records(i)
            Call zlCommFun.ShowFlash(i + 1 & "/" & rptPath.Records.count & "，正在导出门诊临床路径""" & .Item(COL_名称).Value & """ ...", Me)
            If .Item(COL_最新版本).Value > 0 Then
                strFile = strPath & "\" & .Item(COL_分类).Value & "-" & .Item(COL_名称).Value & ".xml"
                If ExportOutPathToXML(.Item(COL_ID).Value, .Item(COL_最新版本).Value, strFile) Then
                    intCount = intCount + 1
                Else
                    strFail = strFail & vbCrLf & strFile
                End If
            End If
        End With
    Next
    Call zlCommFun.StopFlash
    Screen.MousePointer = 0

    strMsg = "导出完成，共成功导出 " & intCount & " 个门诊临床路径文件。" & _
        IIf(strFail <> "", "以下门诊临床路径导出失败：" & vbCrLf & strFail, "")
    MsgBox strMsg, vbInformation, gstrSysName
End Sub

Private Sub FuncImportFromXMLBatch()
'功能：批量导入XML文件
    Dim arrFile() As String
    Dim strFail As String, strMsg As String
    Dim intCount As Long, i As Long
    Dim intLimit As Integer, strLimit As String, blnLimit As Boolean

    If InStr(mstrPrivs, "不限制路径数量") > 0 Then
        intLimit = 0
    ElseIf InStr(mstrPrivs, "30个以下路径") > 0 Then
        intLimit = 30
    ElseIf InStr(mstrPrivs, "5个以下路径") > 0 Then
        intLimit = 5
    Else
        MsgBox "不能明确授权允许的路径数量，请与系统管理员联系。", vbInformation, gstrSysName
        Exit Sub
    End If

    strMsg = "批量导入多个门诊临床路径XML文件时：" & vbCrLf & vbCrLf & _
            "　1.如果系统中不存在相同分类和名称的门诊临床路径，则导入将新增加该门诊临床路径。" & vbCrLf & _
            "　2.如果系统中已存在相同分类和名称的门诊临床路径，则将导入到该门诊临床路径新的版本中。" & vbCrLf & _
            "　　如果该门诊临床路径存在未审核的版本，则导入将覆盖该版本的内容。" & vbCrLf & vbCrLf & _
            "要继续吗？"
    If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub

    cdgFile.DialogTitle = "导入临床路径"
    cdgFile.Filter = "XML文件|*.xml"
    cdgFile.Flags = &H200 Or &H80000 Or &H4 Or &H1000 Or &H200000 Or &H800
    cdgFile.InitDir = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "门诊临床路径XML目录")
    cdgFile.FileName = ""
    cdgFile.MaxFileSize = 25600 '多选时，所有文件名长度有限制(Byte)
    cdgFile.CancelError = True
    On Error Resume Next
    cdgFile.ShowOpen
    If Err.Number <> 0 Then
        Err.Clear: Exit Sub
    End If
    On Error GoTo 0

    If InStr(cdgFile.FileName, Chr(0)) > 0 Then
        ReDim arrFile(UBound(Split(cdgFile.FileName, Chr(0))) - 1)
        For i = 0 To UBound(arrFile)
            arrFile(i) = Split(cdgFile.FileName, Chr(0))(0) & "\" & Split(cdgFile.FileName, Chr(0))(i + 1)
        Next
    Else
        ReDim arrFile(0)
        arrFile(0) = cdgFile.FileName
    End If
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "门诊临床路径XML目录", gobjFile.GetParentFolderName(arrFile(0))

    Screen.MousePointer = 11
    For i = 0 To UBound(arrFile)
        Call zlCommFun.ShowFlash(i + 1 & "/" & UBound(arrFile) + 1 & "，正在导入文件""" & gobjFile.GetFileName(arrFile(i)) & """ ...", Me)
        If ImportOutPathFromXML(arrFile(i), , , intLimit, blnLimit) Then
            intCount = intCount + 1
        Else
            If blnLimit Then
                strLimit = strLimit & vbCrLf & arrFile(i)
            Else
                strFail = strFail & vbCrLf & arrFile(i)
            End If
        End If
    Next
    Call zlCommFun.StopFlash
    Call RefreshData
    Screen.MousePointer = 0

    strMsg = "导入完成，共成功导入 " & intCount & " 个门诊临床路径文件。" & _
        IIf(strFail <> "", "以下门诊临床路径文件导入失败：" & vbCrLf & strFail, "") & _
        IIf(strLimit <> "", "以下路径因授权数量限制未导入：" & vbCrLf & strLimit, "")
    MsgBox strMsg, vbInformation, gstrSysName
End Sub

Private Sub InitVsgIllness()
'功能:初始化对应病种的VSF控件
    With vsgIllness
        .Cols = 3
        .Rows = 5
        .FixedCols = 0
        .FixedRows = 0
        .AllowSelection = False
        .BackColorBkg = vbWhite
        .RowHeightMin = 300
        .Appearance = flexXPThemes
        .BorderStyle = flexBorderNone
        .ScrollBars = flexScrollBarVertical
        .GridLines = flexGridNone
        .ColWidthMin = .Width / .Cols
    End With
End Sub

Private Sub vsgIllness_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'病种，鼠标移上去之后加上下划线，颜色变蓝
    vsgIllness.FontUnderline = True
    vsgIllness.ForeColor = RGB(0, 0, 128)
    vsgIllness.ToolTipText = vsgIllness.Text
End Sub
