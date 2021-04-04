VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEPRFileMan 
   Caption         =   "病历文件管理"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8640
   Icon            =   "frmEPRFileMan.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   8640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox PicFileTab 
      BorderStyle     =   0  'None
      Height          =   5010
      Left            =   195
      ScaleHeight     =   5010
      ScaleWidth      =   4410
      TabIndex        =   2
      Top             =   660
      Width           =   4410
      Begin XtremeReportControl.ReportControl rptList 
         Height          =   4440
         Left            =   15
         TabIndex        =   3
         Top             =   510
         Width           =   3975
         _Version        =   589884
         _ExtentX        =   7011
         _ExtentY        =   7832
         _StockProps     =   0
         BorderStyle     =   2
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
      End
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   870
         TabIndex        =   4
         Top             =   75
         Width           =   3105
      End
      Begin VB.Label lblFind 
         Caption         =   "查找(&V)"
         Height          =   405
         Left            =   135
         TabIndex        =   5
         Top             =   105
         Width           =   945
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6015
      Width           =   8640
      _ExtentX        =   15240
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmEPRFileMan.frx":058A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12330
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
   Begin MSComctlLib.ImageList imgList 
      Left            =   2730
      Top             =   60
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
            Picture         =   "frmEPRFileMan.frx":0E1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEPRFileMan.frx":13B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEPRFileMan.frx":1950
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEPRFileMan.frx":1EEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEPRFileMan.frx":2484
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEPRFileMan.frx":2A1E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vgdList 
      Height          =   900
      Left            =   300
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5730
      Visible         =   0   'False
      Width           =   1080
      _cx             =   1905
      _cy             =   1587
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   0   'False
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
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
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
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
   Begin MSComDlg.CommonDialog dlgThis 
      Left            =   2070
      Top             =   90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   270
      Top             =   105
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmEPRFileMan.frx":2FB8
      Left            =   960
      Top             =   210
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmEPRFileMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-----------------------------------------------------
'常量
'-----------------------------------------------------
Private Enum mCol
    图标 = 0: ID: 种类: 编号: 名称: 说明: 保留: 页面: 子类: 简码
End Enum
Const conPane_FileTab = 1
Const conPane_Request = 2
Const conPane_Compend = 3

'-----------------------------------------------------
'窗体变量
'-----------------------------------------------------
Private mstrPrivs As String     '当前使用者权限串
Private mstrKinds As String     '当前允许定义的病历类型串

Private WithEvents mfrmRequest As frmEPRFileRequest     '应用要求窗格
Attribute mfrmRequest.VB_VarHelpID = -1
Private WithEvents mfrmContent As frmEPRFileContent     '病历提纲窗格
Attribute mfrmContent.VB_VarHelpID = -1
Private mObjTabEpr As cTableEPR

Private mintCurKind As Integer      '病历种类
Private mlngCurFileId As Long       '当前文件ID
Private mstrCurFixed As String      '保留病历特性
Private mblnPartogram As Boolean    '是否是产程文件 (种类=3 And 保留=1)

Private mblnFindTag As Boolean      '搜索框焦点判断
Private mintLastRows As Integer     '搜索最后定位行位置

Public Sub RefreshList()
    Call mfrmContent.zlRefresh(mlngCurFileId)
End Sub

Public Function zlRefList(Optional lngFileID As Long) As Long
    '功能：刷新装入指定种类的病历文件清单，并定位到指定的文件上
Dim strGroups As String
Dim rsTemp As New ADODB.Recordset
Dim rptRcd As ReportRecord
Dim rptItem As ReportRecordItem
Dim rptRow As ReportRow


    Me.rptList.Tag = ""
    gstrSQL = "Select l.Id, l.种类, l.编号, l.名称, l.说明, Nvl(l.保留, 0) As 保留, Decode(f.独立, 1, '单独页面', f.名称) As 页面,l.子类" & _
            " From 病历文件列表 l," & _
            "      (Select f.种类, f.编号, f.名称, Count(l.ID) As 独立" & _
            "        From 病历页面格式 f, 病历文件列表 l" & _
            "        Where f.种类 = l.种类 And f.编号 = l.页面 And f.种类 In (" & mstrKinds & ")" & _
            "        Group By f.种类, f.编号, f.名称) f" & _
            " Where l.种类 = f.种类 And l.页面 = f.编号 and l.保留<>4"
    
    Err = 0: On Error GoTo errHand
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    Me.rptList.Records.DeleteAll
    With rsTemp
        strGroups = ""
        Do While Not .EOF
            If InStr(1, strGroups, !种类) = 0 Then strGroups = strGroups & "," & !种类
            Set rptRcd = Me.rptList.Records.Add()
            Set rptItem = rptRcd.AddItem(CStr(!种类)): rptItem.Icon = rptItem.Value - 1
            rptRcd.AddItem CStr(!ID)
            Select Case !种类
            Case 1: rptRcd.AddItem CStr("1-门诊病历")
            Case 2: rptRcd.AddItem CStr("2-住院病历")
            Case 3: rptRcd.AddItem CStr("3-护理记录")
            Case 4: rptRcd.AddItem CStr("4-护理病历")
            Case 5 And !保留 <> 4: rptRcd.AddItem CStr("5-疾病证明报告")
            Case 6: rptRcd.AddItem CStr("6-知情文件")
            Case Else: rptRcd.AddItem ""
            End Select
            rptRcd.AddItem CStr(!编号)
            rptRcd.AddItem CStr(!名称)
            rptRcd.AddItem CStr("" & !说明)
            Select Case !保留
            Case 0: rptRcd.AddItem ""
            Case 1: rptRcd.AddItem CStr("保留")
            Case 2: rptRcd.AddItem CStr("表格")
            Case 3: rptRcd.AddItem CStr("快捷")
            Case Else
                If NVL(!种类) = 3 And NVL(!保留) = -1 Then
                    rptRcd.AddItem "曲线"
                Else
                    rptRcd.AddItem CStr("特殊")
                End If
            End Select
            rptRcd.AddItem CStr(!页面)
            rptRcd.AddItem CStr(NVL(!子类))
            rptRcd.AddItem zl9ComLib.zlStr.PinYinCode(CStr(!名称))
            .MoveNext
        Loop
        If strGroups <> "" Then strGroups = Mid(strGroups, 2)
    End With
    With Me.rptList
        If UBound(Split(strGroups, ",")) < 1 Then
            .GroupsOrder.DeleteAll
        ElseIf .GroupsOrder.Count = 0 Then
            .GroupsOrder.Add .Columns.Find(mCol.种类)
            .GroupsOrder(0).SortAscending = True
        End If
        .Populate
    End With
    
    If lngFileID <> 0 Then
        For Each rptRow In Me.rptList.Rows
            If rptRow.GroupRow = False Then
                If Val(rptRow.Record(mCol.ID).Value) = lngFileID Then
                    Set Me.rptList.FocusedRow = rptRow: Exit For
                End If
            End If
        Next
    End If
    If Me.rptList.Rows.Count > 0 Then
        If Me.rptList.FocusedRow Is Nothing Then Set Me.rptList.FocusedRow = Me.rptList.Rows(0)
        If Me.rptList.FocusedRow.GroupRow Then
            lngFileID = 0
        Else
            lngFileID = Me.rptList.FocusedRow.Record.Item(mCol.ID).Value
        End If
    Else
        lngFileID = 0
    End If
    
    zlRefList = Me.rptList.Records.Count
    Exit Function

errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    zlRefList = Me.rptList.Records.Count
    lngFileID = 0
End Function

Public Sub zlRptPrint(ByVal bytMode As Byte)
    '功能:将数据复制到可打印的对象，调用打印
    '参数:  bytMode，1-打印;2-预览;3-输出到EXCEL
    If Me.rptList.Records.Count = 0 Then Exit Sub

    '-------------------------------------------------
    '复制数据表格
    If zlReportToVSFlexGrid(Me.vgdList, Me.rptList) = False Then Exit Sub

    '-------------------------------------------------
    '调用打印部件处理
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow

    Set objPrint.Body = Me.vgdList
    objPrint.Title.Text = "病历文件清单"
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("打印时间:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)

    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngFileID As Long, lngCopyId As Long
    Dim cbrControl As CommandBarControl
    Dim str编号 As String, str名称 As String
    
    Select Case Control.ID
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Preview: Call zlRptPrint(0)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    Case conMenu_File_ExportToXML + 1
       frmFileExportOrImport.ShowMe Me, 1
    Case conMenu_File_ExportToXML + 2
        frmFileExportOrImport.ShowMe Me, 2
    Case conMenu_File_ExportToXML
        '导出到XML文件
        If Me.rptList.FocusedRow Is Nothing Then Exit Sub
        If Me.rptList.FocusedRow.GroupRow = True Then Exit Sub
        Dim strF As String
        lngFileID = Me.rptList.FocusedRow.Record.Item(mCol.ID).Value
        '指定保存的文件路径
        On Error Resume Next
        dlgThis.Filename = "定义_" & Me.rptList.FocusedRow.Record.Item(mCol.名称).Value & ".xml"
        dlgThis.Filter = "*.XML|*.xml|*.*|*.*"
        dlgThis.CancelError = True
        dlgThis.ShowSave
        If Err.Number = 32755 Then Err.Clear: Exit Sub
        strF = dlgThis.Filename
        On Error GoTo errHand
        If gobjFSO.FileExists(strF) Then
            DoEvents
            If MsgBox("该文件已经存在，是否覆盖？", vbOKCancel + vbQuestion, gstrSysName) = vbCancel Then Exit Sub
        End If
        
        If mstrCurFixed = "表格" Then '表格式病历导出
            mObjTabEpr.InitOpenEPR Me, cprEM_修改, cprET_病历文件定义, lngFileID, False, 0
            If mObjTabEpr.zlExportXML(strF) Then
                MsgBox "成功导出为XML文件！" & vbCrLf & "文件名:" & strF, vbOKOnly + vbInformation, gstrSysName
            End If
        Else
            Dim DocXML As New cEPRDocument
            '普通住院病历
            DocXML.InitEPRDoc cprEM_修改, cprET_病历文件定义, lngFileID
            DocXML.KeepRTF = True
            DocXML.OpenEPRDoc DocXML.frmEditor.Editor1
            If DocXML.ExportToXMLFile(DocXML.frmEditor.Editor1, strF) Then
                DoEvents
            End If
        End If
    Case conMenu_File_Exit: Unload Me
    
    Case conMenu_Edit_NewItem
        If Me.rptList.FocusedRow Is Nothing Then
            lngCopyId = 0
        ElseIf Me.rptList.FocusedRow.GroupRow = True Then
            lngCopyId = 0
        Else
            lngCopyId = Me.rptList.FocusedRow.Record.Item(mCol.ID).Value
        End If
        lngFileID = frmEPRFileEdit.ShowMe(Me, mstrKinds, True, lngCopyId)
        If lngFileID <> 0 Then Call Me.zlRefList(lngFileID)
    
    Case conMenu_Edit_Modify
        If Me.rptList.FocusedRow Is Nothing Then Exit Sub
        If Me.rptList.FocusedRow.GroupRow = True Then Exit Sub
        lngFileID = Me.rptList.FocusedRow.Record.Item(mCol.ID).Value
        lngFileID = frmEPRFileEdit.ShowMe(Me, mstrKinds, False, lngFileID)
        If lngFileID <> 0 Then Call Me.zlRefList(lngFileID)
    
    Case conMenu_Edit_Delete
        With Me.rptList
            If .FocusedRow Is Nothing Then Exit Sub
            If .FocusedRow.GroupRow Then Exit Sub
            '如果要删除的是专科体温单，检查是否还存其它专科体温单文件,最后一份文件不允许删除
            If mintCurKind = 3 And mstrCurFixed = "曲线" And Me.rptList.FocusedRow.Record(mCol.子类).Value = "1" Then
                If IsLastWaveFile = True Then
                    MsgBox "专科体温单文件至少要保留一份,改份文件不能删除。", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
            If MsgBox("真的删除该文件吗？" & vbCrLf & "——" & .FocusedRow.Record(mCol.名称).Value, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                gstrSQL = "Zl_病历文件列表_Delete(" & .FocusedRow.Record(mCol.ID).Value & ")"
                Err = 0: On Error GoTo errHand
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                Err = 0: On Error GoTo 0
                lngCopyId = .FocusedRow.Record.Index
                Call .Records.RemoveAt(.FocusedRow.Record.Index)
                .Populate
                If .Records.Count <> 0 Then
                    If lngCopyId >= .Records.Count Then lngCopyId = 0
                    lngFileID = .Records(lngCopyId).Item(mCol.ID).Value
                Else
                    lngFileID = 0
                End If
                Call Me.zlRefList(lngFileID)
            End If
        End With
    Case conMenu_Edit_ApplyTo
        If mlngCurFileId = 0 Then Exit Sub
        If frmEPRFileApplyTo.ShowMe(Me, mlngCurFileId) Then Call mfrmRequest.zlRefresh(mlngCurFileId)
    Case conMenu_Edit_Request
        If mlngCurFileId = 0 Then Exit Sub
        Select Case mintCurKind
        Case 1, 2, 4
            If frmEPRFileTimeout.ShowMe(Me, mlngCurFileId) Then Call mfrmRequest.zlRefresh(mlngCurFileId)
        Case 5
            If frmEPRFileDisease.ShowMe(Me, mlngCurFileId) Then Call mfrmRequest.zlRefresh(mlngCurFileId)
        Case 6
            If frmEPRFileMeasure.ShowMe(Me, mlngCurFileId) Then Call mfrmRequest.zlRefresh(mlngCurFileId)
        End Select
    Case conMenu_Edit_Compend
        If mlngCurFileId = 0 Then Exit Sub
        If mintCurKind = 3 Then
            '护理记录样式定义
            If mstrCurFixed = "曲线" Then
                If Me.rptList.FocusedRow Is Nothing Then Exit Sub
                If Me.rptList.FocusedRow.GroupRow Then Exit Sub
                If Me.rptList.FocusedRow.Record(mCol.子类).Value = "1" Then
                    If frmTendWaveStyle.ShowMe(Me, mlngCurFileId) = True Then
                        Me.rptList.Tag = ""
                        Call rptList_SelectionChanged
                    End If
                Else
                    Call frmTendWavePrintSet.ShowMe(Me, mlngCurFileId)
                End If
            ElseIf mstrCurFixed = "保留" Then
                '产程图
                If frmTendPartogramStyle.ShowMe(Me, mlngCurFileId) Then
                    Me.rptList.Tag = ""
                    Call rptList_SelectionChanged
                End If
            Else
                
                If frmTendFileStyle.ShowMe(Me, mlngCurFileId) Then
                    Me.rptList.Tag = ""
                    Call rptList_SelectionChanged
                End If
            End If
            
        ElseIf mintCurKind = 2 And mstrCurFixed = "特殊" Then
        
        ElseIf mstrCurFixed = "表格" Then
            On Error GoTo errHand
            mObjTabEpr.InitOpenEPR Me, cprEM_修改, cprET_病历文件定义, mlngCurFileId
        Else
            Dim Doc As New cEPRDocument
            If mlngCurFileId = 0 Then Exit Sub
            Doc.InitEPRDoc cprEM_修改, cprET_病历文件定义, mlngCurFileId
            Doc.ShowEPREditor Me
        End If
    Case conMenu_Edit_ElementChange
        frmElementChange.ShowMe Me, mlngCurFileId
    Case conMenu_Edit_Privacy
        '隐私保护设置
        Dim frmP As New frmPrivacyProtect
        frmP.ShowMe Me
    Case conMenu_View_ToolBar_Button
        Me.cbsThis(2).Visible = Not Me.cbsThis(2).Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Text
        For Each cbrControl In Me.cbsThis(2).Controls
            cbrControl.STYLE = IIf(cbrControl.STYLE = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Size
        Me.cbsThis.Options.LargeIcons = Not Me.cbsThis.Options.LargeIcons
        Me.cbsThis.RecalcLayout
    Case conMenu_View_StatusBar
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsThis.RecalcLayout
    
    Case conMenu_View_Jump
'        If Screen.ActiveForm.Name = mfrmFileTab.Name Then
'            Call Me.dkpMan.Panes(conPane_Request).Select
'        ElseIf Screen.ActiveForm.Name = mfrmRequest.Name Then
'            Call Me.dkpMan.Panes(conPane_Compend).Select
'        Else
'            Call Me.dkpMan.Panes(conPane_FileTab).Select
'        End If
    Case conMenu_View_LocationItem
        txtFind.SetFocus
    Case conMenu_View_Refresh
        Call zlRefList(mlngCurFileId)
    
    Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Forum '中联论坛
        Call zlWebForum(Me.hWnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision)
    Case Else
        '执行发布到当前模块的报表
        If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            If rptList.SelectedRows.Count > 0 Then
                If Not rptList.SelectedRows(0).GroupRow Then
                    str编号 = rptList.SelectedRows(0).Record(mCol.编号).Value
                    str名称 = rptList.SelectedRows(0).Record(mCol.名称).Value
                End If
            End If
            If str名称 <> "" Then
                Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, "编号=" & str编号, "名称=" & str名称)
            Else
                Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me)
            End If
        End If
    End Select

    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnZWave As Boolean
    If Me.Visible = False Then Exit Sub
        
    If mblnFindTag = True Then
        txtFind.ForeColor = vbBlack
        If txtFind.Text = "请输入名称或拼音简码" Then txtFind.Text = ""
    Else
        If txtFind.Text = "" Then txtFind.ForeColor = vbGrayText: txtFind.Text = "请输入名称或拼音简码"
    End If
    
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = (mstrKinds <> "")
        End Select
    End If
    
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel: Control.Enabled = (Me.rptList.Records.Count <> 0)
    Case conMenu_File_ExportToXML
        Control.Enabled = (Me.rptList.Records.Count <> 0)
    Case conMenu_Edit_NewItem: Control.Enabled = (mstrKinds <> "" And InStr(1, mstrPrivs, "文件增删改") > 0)
    Case conMenu_Edit_Modify: Control.Enabled = (mlngCurFileId <> 0 And InStr(1, mstrPrivs, "文件增删改") > 0) ' And mstrCurFixed <> "特殊"
    Case conMenu_Edit_Delete
        '专科体温单允许删除,标准体温单不能删除
        blnZWave = False
        If Not rptList.FocusedRow Is Nothing Then
            If Not rptList.FocusedRow.GroupRow Then
                 If mstrCurFixed = "曲线" And mintCurKind = 3 And rptList.FocusedRow.Record(mCol.子类).Value = "1" Then
                    blnZWave = True
                 End If
            End If
        End If
        Control.Enabled = (mlngCurFileId <> 0 And InStr(1, mstrPrivs, "文件增删改") > 0) And (Trim(mstrCurFixed) = "" Or mstrCurFixed = "表格" Or mstrCurFixed = "快捷" Or blnZWave = True)
    Case conMenu_Edit_ApplyTo: Control.Enabled = (mlngCurFileId <> 0 And Not mblnPartogram And InStr(1, mstrPrivs, "适用科室") > 0)
    Case conMenu_Edit_Request: Control.Enabled = (mlngCurFileId <> 0 And InStr(1, mstrPrivs, "限制要求") > 0) And mintCurKind <> 3
    Case conMenu_Edit_Compend
        Control.Enabled = (mlngCurFileId <> 0 And InStr(1, mstrPrivs, "样式构造") > 0)
        If Control.Enabled Then Control.Enabled = (mintCurKind <> 3 Or mintCurKind = 3 And mstrCurFixed <> "特殊")
        If Control.Enabled Then Control.Enabled = mstrCurFixed <> "特殊"
    Case conMenu_Edit_Privacy: Control.Enabled = (InStr(1, mstrPrivs, "隐私设置") > 0)
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).STYLE = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
    Case conMenu_Edit_ElementChange: Control.Enabled = (mlngCurFileId <> 0) And Not (mstrCurFixed = "表格" Or mstrCurFixed = "快捷" Or mstrCurFixed = "特殊" Or mintCurKind = 3)
    End Select
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_FileTab
        Item.Handle = Me.PicFileTab.hWnd
    Case conPane_Request
        If mfrmRequest Is Nothing Then Set mfrmRequest = New frmEPRFileRequest
        Item.Handle = mfrmRequest.hWnd
    Case conPane_Compend
        If mfrmContent Is Nothing Then Set mfrmContent = New frmEPRFileContent
        Item.Handle = mfrmContent.hWnd
    End Select
End Sub

Private Sub Form_Load()
Dim cbrControl As CommandBarControl
Dim cbrMenuBar As CommandBarPopup
Dim cbrToolBar As CommandBar
Dim rptCol As ReportColumn
Dim lngCount As Long

    '-----------------------------------------------------
    '权限限制串复制，避免同时进入其他模块而导致gstrPrivs变化，导致控制无效
    mstrPrivs = gstrPrivs
    mstrKinds = ""
    mblnPartogram = False
    If InStr(1, mstrPrivs, "门诊病历") > 0 Then mstrKinds = mstrKinds & ",1"
    If InStr(1, mstrPrivs, "住院病历") > 0 Then mstrKinds = mstrKinds & ",2"
    If InStr(1, mstrPrivs, "护理记录") > 0 Then mstrKinds = mstrKinds & ",3"
    If InStr(1, mstrPrivs, "护理病历") > 0 Then mstrKinds = mstrKinds & ",4"
    If InStr(1, mstrPrivs, "疾病证明报告") > 0 Then mstrKinds = mstrKinds & ",5"
    If InStr(1, mstrPrivs, "知情文件") > 0 Then mstrKinds = mstrKinds & ",6"
    If mstrKinds <> "" Then mstrKinds = Mid(mstrKinds, 2)
    
    Call ZLCommFun.SetWindowsInTaskBar(Me.hWnd, gblnShowInTaskBar)
    
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Set Me.cbsThis.Icons = ZLCommFun.GetPubIcons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsThis.EnableCustomization False
    
    '-----------------------------------------------------
    '菜单定义
    Me.cbsThis.ActiveMenuBar.Title = "菜单"
    Me.cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_ExportToXML, "导出为XML文件(&L)…")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_ExportToXML + 1, "批量导出XML文件(&E)…")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_ExportToXML + 2, "批量导入XML文件(&I)…")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&D)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ApplyTo, "适用科室(&T)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Request, "限制要求(&R)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Compend, "样式构造(&F)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ElementChange, "要素联动设置(&E)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Privacy, "隐私项目设置(&P)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_View_LocationItem, "查找(&F)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Jump, "窗格跳转(&J)")
    End With
    
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "论坛(&F)", -1, False  '固有
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): cbrControl.BeginGroup = True
    End With
    
    '快键绑定
    With Me.cbsThis.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add 0, VK_DELETE, conMenu_Edit_Delete
        .Add FCONTROL, Asc("T"), conMenu_Edit_ApplyTo
        .Add FCONTROL, Asc("R"), conMenu_Edit_Request
        .Add FCONTROL, Asc("D"), conMenu_Edit_Compend
        .Add FCONTROL, Asc("F"), conMenu_View_LocationItem
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F6, conMenu_View_Jump
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    '设置不常用菜单
    With Me.cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Jump
        .AddHiddenCommand conMenu_View_Refresh
    End With
    '-----------------------------------------------------
    '工具栏定义
    Set cbrToolBar = Me.cbsThis.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ApplyTo, "使用科室"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Request, "限制要求")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Compend, "样式构造")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.STYLE = xtpButtonIconAndCaption
    Next
    
    '读取发布到该模块的报表:因为是一次性读取,全局变量可用
    '---------------------------------------------------------------
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, gstrPrivs)
    
    '-----------------------------------------------------
    '设置词句显示停靠窗格
    If mfrmRequest Is Nothing Then Set mfrmRequest = New frmEPRFileRequest
    If mfrmContent Is Nothing Then Set mfrmContent = New frmEPRFileContent
    If mObjTabEpr Is Nothing Then Set mObjTabEpr = New cTableEPR
    mObjTabEpr.InitTableEPR gcnOracle, glngSys, gstrDbOwner
    
    Dim panFileTab As Pane, panRequest As Pane, panCompend As Pane
    Set panFileTab = dkpMan.CreatePane(conPane_FileTab, 180, 400, DockLeftOf, Nothing)
    panFileTab.Title = "文件列表"
    panFileTab.Options = PaneNoCaption
    
    Set panRequest = dkpMan.CreatePane(conPane_Request, 400, 200, DockRightOf, Nothing)
    panRequest.Title = "应用要求"
    panRequest.Options = PaneNoCloseable Or PaneNoFloatable 'Or PaneNoHideable

    Set panCompend = dkpMan.CreatePane(conPane_Compend, 400, 300, DockBottomOf, panRequest)
    panCompend.Title = "文件样式"
    panCompend.Options = PaneNoCloseable Or PaneNoFloatable 'Or PaneNoHideable
    
    Me.dkpMan.SetCommandBars Me.cbsThis
    Me.dkpMan.Options.ThemedFloatingFrames = True
    Me.dkpMan.Options.HideClient = True
    
    '-----------------------------------------------------
    With Me.rptList
        Set rptCol = .Columns.Add(mCol.图标, "", 18, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mCol.ID, "ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.种类, "种类", 90, False): rptCol.Editable = False: rptCol.Groupable = True: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.编号, "编号", 50, False): rptCol.Editable = False: rptCol.Groupable = False: .SortOrder.Add rptCol
        Set rptCol = .Columns.Add(mCol.名称, "名称", 120, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.说明, "说明", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.保留, "类型", 30, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.页面, "页面", 80, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.子类, "子类", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.简码, "简码", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        
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
    End With
    
    '-----------------------------------------------------
    '查询框初始化
    mblnFindTag = False
    txtFind.ForeColor = vbGrayText
    txtFind.Text = "请输入名称或拼音简码"
    mintLastRows = 0
    
    '-----------------------------------------------------
    '界面恢复
    Call RestoreWinState(Me, App.ProductName)
    '-----------------------------------------------------
    '数据装入
    If mstrKinds = "" Then
        DoEvents
        Me.stbThis.Panels(2).Text = "你不具备病历文件定义管理权限"
    Else
        lngCount = Me.zlRefList()
        Me.stbThis.Panels(2).Text = "共有" & lngCount & "个病历文件"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload mfrmRequest
    Unload mfrmContent
    Set mfrmRequest = Nothing
    Set mfrmContent = Nothing
    Set mObjTabEpr = Nothing
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub mfrmContent_DblClick()
    Dim cbrControl As CommandBarControl
    If mlngCurFileId = 0 Or mstrCurFixed = "曲线" Then Exit Sub
    Set cbrControl = Me.cbsThis.FindControl(, conMenu_Edit_Compend)
    If cbrControl Is Nothing Then Exit Sub
    If cbrControl.Visible = False Or cbrControl.Enabled = False Then Exit Sub
    Call cbsThis_Execute(cbrControl)
End Sub
Private Sub mfrmRequest_DblClick(lngWhere As zlEnumDClick)
Dim cbrControl As CommandBarControl
    If mlngCurFileId = 0 Then Exit Sub
    Select Case lngWhere
    Case cprEmDClickApplyTo: Set cbrControl = Me.cbsThis.FindControl(, conMenu_Edit_ApplyTo)
    Case cprEmDClickRequest: Set cbrControl = Me.cbsThis.FindControl(, conMenu_Edit_Request)
    Case Else: Set cbrControl = Nothing
    End Select
    If cbrControl Is Nothing Then Exit Sub
    If cbrControl.Visible = False Or cbrControl.Enabled = False Then Exit Sub
    Call cbsThis_Execute(cbrControl)
End Sub


Private Sub PicFileTab_Resize()
    lblFind.Move 70, 90, lblFind.Width, lblFind.Height
    If PicFileTab.Width > 800 Then txtFind.Move 800, 50, PicFileTab.Width - 800, 300
    If PicFileTab.Height > 400 Then rptList.Move 0, 400, PicFileTab.Width, PicFileTab.Height - 400
End Sub

Private Sub rptList_KeyDown(KeyCode As Integer, Shift As Integer)
    If Me.rptList.Visible = False Then Exit Sub
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Me.rptList.FocusedRow Is Nothing Then Exit Sub
    If Me.rptList.FocusedRow.GroupRow Then Exit Sub
    Call rptList_RowDblClick(Me.rptList.FocusedRow, Me.rptList.FocusedRow.Record.Item(mCol.编号))
End Sub

Private Sub rptList_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
Dim cbrControl As CommandBarControl
Dim cbrMenuBar As CommandBarPopup
Dim cbrPopupBar As CommandBar
Dim cbrPopupItem As CommandBarControl
    
    If Button <> vbRightButton Then Exit Sub
    If Me.cbsThis.ActiveMenuBar.Controls(2).Visible = False Then Exit Sub

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls(2)
    Set cbrPopupBar = Me.cbsThis.Add("弹出菜单", xtpBarPopup)
    For Each cbrControl In cbrMenuBar.CommandBar.Controls
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, cbrControl.ID, cbrControl.Caption)
        cbrPopupItem.BeginGroup = cbrControl.BeginGroup
    Next
    cbrPopupBar.ShowPopup
End Sub

Private Sub rptList_RowDblClick(ByVal ROW As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
Dim cbrControl As CommandBarControl

    With Me.rptList
        If .FocusedRow Is Nothing Then
            mintCurKind = 0: mlngCurFileId = 0: mstrCurFixed = "": mblnPartogram = False
        ElseIf .FocusedRow.GroupRow = True Then
            mintCurKind = .FocusedRow.Childs.ROW(0).Record.Item(mCol.图标).Value: mlngCurFileId = 0: mstrCurFixed = "": mblnPartogram = False
        Else
            mintCurKind = .FocusedRow.Record.Item(mCol.图标).Value
            mlngCurFileId = .FocusedRow.Record.Item(mCol.ID).Value
            mstrCurFixed = .FocusedRow.Record.Item(mCol.保留).Value
            mblnPartogram = ((mstrCurFixed = "保留") And (mintCurKind = 3))
        End If
    End With
    If mlngCurFileId = 0 Then Exit Sub
    
    Set cbrControl = Me.cbsThis.FindControl(, conMenu_Edit_Modify)
    If cbrControl Is Nothing Then Exit Sub
    If cbrControl.Visible = False Or cbrControl.Enabled = False Then Exit Sub
    Call cbsThis_Execute(cbrControl)

End Sub

Private Sub rptList_SelectionChanged()
    With Me.rptList
        If .FocusedRow Is Nothing Then
            mintCurKind = 0: mlngCurFileId = 0: mstrCurFixed = "": mblnPartogram = False
        ElseIf .FocusedRow.GroupRow = True Then
            mintCurKind = .FocusedRow.Childs.ROW(0).Record.Item(mCol.图标).Value: mlngCurFileId = 0: mstrCurFixed = "": mblnPartogram = False
        Else
            mintCurKind = .FocusedRow.Record.Item(mCol.图标).Value
            mlngCurFileId = .FocusedRow.Record.Item(mCol.ID).Value
            mstrCurFixed = .FocusedRow.Record.Item(mCol.保留).Value
            mblnPartogram = ((mstrCurFixed = "保留") And (mintCurKind = 3))
            If Val(Me.rptList.Tag) <> Me.rptList.FocusedRow.Index Then
                Call mfrmRequest.zlRefresh(mlngCurFileId)
                Call mfrmContent.zlRefresh(mlngCurFileId)
                Me.rptList.Tag = Me.rptList.FocusedRow.Index
            End If
        End If
    End With
End Sub

Private Function IsLastWaveFile() As Boolean
'功能:检查专科体温单文件是否是最后一份
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    gstrSQL = "Select Count(1) 数目 From 病历文件列表 where 种类=3 And 保留=-1 And 子类='1'"
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, "病历文件列表")
    IsLastWaveFile = rsTemp!数目 = 1
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub txtFind_Change()
    mintLastRows = 0
End Sub

Private Sub txtFind_GotFocus()
    mblnFindTag = True
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyBack Then txtFind.SetFocus '防止按删除键焦点转移bug
End Sub

Private Sub txtFind_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim intCount As Integer

    If KeyCode = vbKeyReturn And txtFind.Text <> "" Then
        For intCount = mintLastRows + 1 To Me.rptList.Rows.Count - 1
            If Me.rptList.Rows(intCount).GroupRow = False Then
                If InStr(Me.rptList.Rows(intCount).Record(mCol.名称).Value, txtFind.Text) Or InStr(Me.rptList.Rows(intCount).Record(mCol.简码).Value, UCase(txtFind.Text)) Then
                    Set Me.rptList.FocusedRow = Me.rptList.Rows(intCount)
                    mintLastRows = intCount
                    Exit For
                End If
            End If
        Next
        If intCount = Me.rptList.Rows.Count And mintLastRows = 0 Then
            Call MsgBox("未找到与“" & txtFind.Text & "”匹配的病历，请重新输入名称或简码。", vbInformation, gstrSysName)
            txtFind.Text = ""
        End If
    End If
    txtFind.SetFocus
End Sub

Private Sub txtFind_LostFocus()
    mblnFindTag = False
End Sub
