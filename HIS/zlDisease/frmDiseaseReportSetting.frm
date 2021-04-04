VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDiseaseReportSetting 
   Caption         =   "疾病报告设置"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8640
   Icon            =   "frmDiseaseReportSetting.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   8640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin XtremeReportControl.ReportControl rptList 
      Height          =   840
      Left            =   45
      TabIndex        =   6
      Top             =   1680
      Width           =   4170
      _Version        =   589884
      _ExtentX        =   7355
      _ExtentY        =   1482
      _StockProps     =   0
      BorderStyle     =   2
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
   End
   Begin VB.PictureBox PicLeftHead 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   45
      ScaleHeight     =   1335
      ScaleWidth      =   4170
      TabIndex        =   7
      Top             =   2640
      Width           =   4170
      Begin XtremeSuiteControls.TabControl tabMain 
         Height          =   1095
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   3810
         _Version        =   589884
         _ExtentX        =   6720
         _ExtentY        =   1931
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picParameter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1275
      Left            =   3600
      ScaleHeight     =   1275
      ScaleWidth      =   4170
      TabIndex        =   1
      Top             =   120
      Width           =   4170
      Begin VB.CheckBox chkOneCard 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "传染病报告卡一病一卡"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   480
         TabIndex        =   10
         Top             =   840
         Width           =   2415
      End
      Begin VB.CheckBox chkIDNO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "报告卡中有效证件号必须填写"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   620
         Width           =   3255
      End
      Begin VB.OptionButton optParameter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "提示编辑报告卡"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   465
         TabIndex        =   4
         Top             =   315
         Value           =   -1  'True
         Width           =   1605
      End
      Begin VB.OptionButton optParameter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "弹出编辑报告卡"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   3
         Top             =   315
         Width           =   1680
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "首页整理后:"
         Height          =   225
         Left            =   90
         TabIndex        =   2
         Top             =   45
         Width           =   1110
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
            Picture         =   "frmDiseaseReportSetting.frx":058A
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
   Begin MSComDlg.CommonDialog dlgThis 
      Left            =   2070
      Top             =   90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VSFlex8Ctl.VSFlexGrid vgdList 
      Height          =   900
      Left            =   120
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4200
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
   Begin MSComctlLib.ImageList imgList 
      Left            =   120
      Top             =   5160
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
            Picture         =   "frmDiseaseReportSetting.frx":0E1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiseaseReportSetting.frx":13B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiseaseReportSetting.frx":1950
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiseaseReportSetting.frx":1EEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiseaseReportSetting.frx":2484
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiseaseReportSetting.frx":2A1E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   300
      Top             =   105
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmDiseaseReportSetting.frx":2FB8
      Left            =   960
      Top             =   210
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmDiseaseReportSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'窗体变量
'-----------------------------------------------------
Private Enum mCol
    图标 = 0: ID: 种类: 编号: 名称: 说明: 保留: 子类
End Enum

Const conPane_Parameter = 1
Const conPane_Request = 2
Const conPane_Compend = 3

Private mstrPrivs As String                             '当前使用者权限串
Private mfrmRequest As Object                           '应用要求窗格
Attribute mfrmRequest.VB_VarHelpID = -1
Private mfrmContent As Object                           '病历提纲窗格
Attribute mfrmContent.VB_VarHelpID = -1
Private WithEvents mDockDisease As zlRichEPR.cDockDisease
Attribute mDockDisease.VB_VarHelpID = -1
Private mstrKinds As String                             '当前允许定义的病历类型串
Private mblnFileList As Boolean                         '是否显示传染病病例文件列表
Private mObjTabEpr As cTableEPR
Private mlngFileID As Long
Private mlngCurFileId As Long                           '病例列表中的当前文件ID
Private mlngFixedFileID As Long                         '定制的中华人民共和国传染病报告卡
Private mstrCurFixed As String                          '保留病历特性

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngFileID As Long, lngCopyId As Long
    Dim cbrControl As CommandBarControl
    Dim strSQL As String

    Select Case Control.ID
        Case conMenu_File_PrintSet
            Call zlPrintSet
        Case conMenu_File_Preview
            Call zlRptPrint(0)
        Case conMenu_File_Print
            Call zlRptPrint(1)
        Case conMenu_File_Excel
            Call zlRptPrint(3)
        Case conMenu_File_ExportToXML + 1
            mDockDisease.zlGetFrmFileExportOrImport.ShowMe Me, 1
        Case conMenu_File_ExportToXML + 2
            mDockDisease.zlGetFrmFileExportOrImport.ShowMe Me, 2
        Case conMenu_File_ExportToXML
            '导出到XML文件
            If Me.rptList.FocusedRow Is Nothing Then Exit Sub
            If Me.rptList.FocusedRow.GroupRow = True Then Exit Sub
            Dim strF As String
            lngFileID = Me.rptList.FocusedRow.Record.Item(mCol.ID).Value
            '指定保存的文件路径
            On Error Resume Next
            dlgThis.FileName = "定义_" & Me.rptList.FocusedRow.Record.Item(mCol.名称).Value & ".xml"
            dlgThis.Filter = "*.XML|*.xml|*.*|*.*"
            dlgThis.CancelError = True
            dlgThis.ShowSave
            If Err.Number = 32755 Then Err.Clear: Exit Sub
            strF = dlgThis.FileName
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
        Case conMenu_File_Exit
            Unload Me
        Case conMenu_Edit_NewItem
            If Me.rptList.FocusedRow Is Nothing Then
                lngCopyId = 0
            ElseIf Me.rptList.FocusedRow.GroupRow = True Then
                lngCopyId = 0
            Else
                lngCopyId = Me.rptList.FocusedRow.Record.Item(mCol.ID).Value
            End If
            lngFileID = mDockDisease.zlGetFrmEPRFileEdit.ShowMe(Me, mstrKinds, True, lngCopyId)
            If lngFileID <> 0 Then Call zlRefList(lngFileID)
        Case conMenu_Edit_Modify
            If Me.rptList.FocusedRow Is Nothing Then Exit Sub
            If Me.rptList.FocusedRow.GroupRow = True Then Exit Sub
            lngFileID = Me.rptList.FocusedRow.Record.Item(mCol.ID).Value
            lngFileID = mDockDisease.zlGetFrmEPRFileEdit.ShowMe(Me, mstrKinds, False, lngFileID)
            If lngFileID <> 0 Then Call zlRefList(lngFileID)
        Case conMenu_Edit_Delete
            With Me.rptList
                If .FocusedRow Is Nothing Then Exit Sub
                If .FocusedRow.GroupRow Then Exit Sub
                If MsgBox("真的删除该文件吗？" & vbCrLf & "——" & .FocusedRow.Record(mCol.名称).Value, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    strSQL = "Zl_病历文件列表_Delete(" & .FocusedRow.Record(mCol.ID).Value & ")"
                    Err = 0: On Error GoTo errHand
                    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
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
                    Call zlRefList(lngFileID)
                End If
            End With
        Case conMenu_Edit_ApplyTo
            If mlngFileID = 0 Then Exit Sub
            If mDockDisease.zlGetFrmEPRFileApplyTo.ShowMe(Me, mlngFileID) Then Call mfrmRequest.zlRefresh(mlngFileID)
        Case conMenu_Edit_Request
            If mlngFileID = 0 Then Exit Sub
            If mDockDisease.zlGetFrmEPRFileDisease.ShowMe(Me, mlngFileID) Then Call mfrmRequest.zlRefresh(mlngFileID)
        Case conMenu_Edit_Compend
            If mlngCurFileId = 0 Then Exit Sub
            If mstrCurFixed = "表格" Then
                On Error GoTo errHand
                mObjTabEpr.InitOpenEPR Me, cprEM_修改, cprET_病历文件定义, mlngCurFileId
            Else
                Dim Doc As New cEPRDocument
                If mlngCurFileId = 0 Then Exit Sub
                Doc.InitEPRDoc cprEM_修改, cprET_病历文件定义, mlngCurFileId
                Doc.ShowEPREditor Me
            End If
        Case conMenu_Edit_ElementChange
            mDockDisease.zlGetFrmElementChange.ShowMe Me, mlngCurFileId
        Case conMenu_Edit_Privacy
            '隐私保护设置
            mDockDisease.zlGetFrmPrivacyProtect.ShowMe Me, glngModul
        Case conMenu_View_ToolBar_Button
            Me.cbsThis(2).Visible = Not Me.cbsThis(2).Visible
            Me.cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Text
            For Each cbrControl In Me.cbsThis(2).Controls
                cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            Me.cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Size
            Me.cbsThis.Options.LargeIcons = Not Me.cbsThis.Options.LargeIcons
            Me.cbsThis.RecalcLayout
        Case conMenu_View_StatusBar
            Me.stbThis.Visible = Not Me.stbThis.Visible
            Me.cbsThis.RecalcLayout
        Case conMenu_View_Refresh
            If mblnFileList Then
                 Call zlRefList(mlngCurFileId)
            Else
                Call mfrmRequest.zlRefresh(mlngFixedFileID)
                Call mfrmContent.zlRefresh(mlngFixedFileID)
            End If
        Case conMenu_Help_Help
            Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_Help_Web_Home
            Call zlHomePage(Me.hwnd)
        Case conMenu_Help_Web_Forum '中联论坛
            Call zlWebForum(Me.hwnd)
        Case conMenu_Help_Web_Mail
            Call zlMailTo(Me.hwnd)
        Case conMenu_Help_About
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
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
    If Me.Visible = False Then Exit Sub
    On Error Resume Next
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
            Case conMenu_EditPopup
                Control.Visible = mblnFileList
        End Select
    End If

    Err = 0: On Error Resume Next
    Select Case Control.ID
        Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel, conMenu_File_ExportToXML
            Control.Visible = mblnFileList
            If Control.Visible Then Control.Enabled = Me.rptList.Records.Count <> 0
        Case conMenu_Edit_NewItem
            Control.Visible = mblnFileList
            If Control.Visible Then Control.Enabled = InStr(1, mstrPrivs, "文件增删改") > 0
        Case conMenu_Edit_Modify
            Control.Visible = mblnFileList
            If Control.Visible Then Control.Enabled = (mlngCurFileId <> 0 And InStr(1, mstrPrivs, "文件增删改") > 0)
        Case conMenu_Edit_Delete
            Control.Visible = mblnFileList
            If Control.Visible Then Control.Enabled = (mlngCurFileId <> 0 And InStr(1, mstrPrivs, "文件增删改") > 0) And (Trim(mstrCurFixed) = "" Or mstrCurFixed = "表格" Or mstrCurFixed = "快捷")
        Case conMenu_Edit_ApplyTo
            Control.Enabled = (mlngFileID <> 0 And InStr(1, mstrPrivs, "适用科室") > 0)
        Case conMenu_Edit_Request
            Control.Enabled = (mlngFileID <> 0 And InStr(1, mstrPrivs, "限制要求") > 0)
        Case conMenu_Edit_Compend
            Control.Visible = mblnFileList
            If Control.Visible Then Control.Enabled = (mlngCurFileId <> 0 And InStr(1, mstrPrivs, "样式构造") > 0)
            If Control.Enabled Then Control.Enabled = mstrCurFixed <> "特殊"
        Case conMenu_Edit_Privacy
             Control.Visible = mblnFileList
            If Control.Visible Then Control.Enabled = (InStr(1, mstrPrivs, "隐私设置") > 0)
        Case conMenu_View_ToolBar_Button
            Control.Checked = Me.cbsThis(2).Visible
        Case conMenu_View_ToolBar_Text
            Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
        Case conMenu_View_ToolBar_Size
            Control.Checked = Me.cbsThis.Options.LargeIcons
        Case conMenu_View_StatusBar
            Control.Checked = Me.stbThis.Visible
        Case conMenu_Edit_ElementChange
            Control.Visible = mblnFileList
            If Control.Visible Then Control.Enabled = (mlngCurFileId <> 0) And Not (mstrCurFixed = "表格" Or mstrCurFixed = "快捷" Or mstrCurFixed = "特殊")
    End Select
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
        Case conPane_Parameter
            Item.Handle = PicLeftHead.hwnd
        Case conPane_Request
            If mfrmRequest Is Nothing Then Set mfrmRequest = mDockDisease.zlGetFrmEPRFileRequest
            Item.Handle = mfrmRequest.hwnd
        Case conPane_Compend
            If mfrmContent Is Nothing Then Set mfrmContent = mDockDisease.zlGetFrmEPRFileContent
            Item.Handle = mfrmContent.hwnd
    End Select
End Sub

Private Sub Form_Load()
    Dim panParameter As Pane, panRequest As Pane, panCompend As Pane

    '只显示疾病报告病历文件，种类为5
    mstrKinds = ",5,"
    mblnFileList = False
    Set mDockDisease = New zlRichEPR.cDockDisease
    Call gobjComlib.ZLCommFun.SetWindowsInTaskBar(Me.hwnd, gblnShowInTaskBar)
    '--------------------------------------------------------------------------------
    '读取发布到该模块的报表:因为是一次性读取,全局变量可用
    '--------------------------------------------------------------------------------
    Call gobjComlib.zlDatabase.ShowReportMenu(Me, glngSys, glngModul, gstrPrivs)
    '--------------------------------------------------------------------------------
    If mfrmRequest Is Nothing Then Set mfrmRequest = mDockDisease.zlGetFrmEPRFileRequest
    If mfrmContent Is Nothing Then Set mfrmContent = mDockDisease.zlGetFrmEPRFileContent
    If mObjTabEpr Is Nothing Then Set mObjTabEpr = New cTableEPR
    mObjTabEpr.InitTableEPR gcnOracle, glngSys, gstrDBOwer
    
    '初始化菜单控件
    Call InitCommandBar
    Call InitTabContol
    Call InitReportControl
    
    Call zlRefList

    Set panParameter = dkpMan.CreatePane(conPane_Parameter, 100, 30, DockTopOf, Nothing)
    panParameter.Title = "参数设置"
    panParameter.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoCaption

    Set panRequest = dkpMan.CreatePane(conPane_Request, 100, 90, DockBottomOf, Nothing)
    panRequest.Title = "应用要求"
    panRequest.Options = PaneNoCloseable Or PaneNoFloatable

    Set panCompend = dkpMan.CreatePane(conPane_Compend, Me.ScaleX(Screen.Width, vbTwips, vbPixels) - 400, 100, DockRightOf, Nothing)
    panCompend.Title = "文件样式"
    panCompend.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoCaption

    If gobjComlib.zlDatabase.GetPara("首页整理后编辑疾控报告卡", glngSys, 1277, "0") = 0 Then
        optParameter(0).Value = True
    Else
        optParameter(1).Value = True
    End If
    
    If 1 = Val(gobjComlib.zlDatabase.GetPara("传染病报告身份证号码必填", glngSys, 1277, "0")) Then
        chkIDNO.Value = 1
    Else
        chkIDNO.Value = 0
    End If
    
    If 1 = Val(gobjComlib.zlDatabase.GetPara("传染病报告卡一病一卡", glngSys, 1277, "0")) Then
        chkOneCard.Value = 1
    Else
        chkOneCard.Value = 0
    End If
    
    Me.dkpMan.SetCommandBars Me.cbsThis
    Me.dkpMan.Options.ThemedFloatingFrames = True
    Me.dkpMan.Options.HideClient = True
    mstrPrivs = gstrPrivs
    
    '初始化传染病报告卡
    mlngFileID = mlngFixedFileID
    
    '界面恢复
    Call gobjComlib.RestoreWinState(Me, App.ProductName)
    Me.WindowState = vbMaximized
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mfrmRequest Is Nothing Then
        Unload mfrmRequest
        Set mfrmRequest = Nothing
    End If
    If Not mfrmContent Is Nothing Then
        Unload mfrmContent
        Set mfrmContent = Nothing
    End If
    Set mDockDisease = Nothing
    Set mObjTabEpr = Nothing
    Call gobjComlib.SaveWinState(Me, App.ProductName)
End Sub


Private Sub mDockDisease_EPRFileRequestDblClick(lngWhere As Integer)
    Dim cbrControl As CommandBarControl

    Select Case lngWhere
        Case 1: Set cbrControl = Me.cbsThis.FindControl(, conMenu_Edit_ApplyTo)
        Case 2: Set cbrControl = Me.cbsThis.FindControl(, conMenu_Edit_Request)
        Case Else: Set cbrControl = Nothing
    End Select
    If cbrControl Is Nothing Then Exit Sub
    If cbrControl.Visible = False Or cbrControl.Enabled = False Then Exit Sub
    Call cbsThis_Execute(cbrControl)
End Sub

Private Sub optParameter_Click(Index As Integer)
    Call gobjComlib.zlDatabase.SetPara("首页整理后编辑疾控报告卡", CStr(Index), glngSys, 1277)
End Sub

Private Sub chkIDNO_Click()
    Call gobjComlib.zlDatabase.SetPara("传染病报告身份证号码必填", chkIDNO.Value, glngSys, 1277)
    Call mfrmContent.SetCaption身份证
End Sub

Private Sub chkOneCard_Click()
    Call gobjComlib.zlDatabase.SetPara("传染病报告卡一病一卡", chkOneCard.Value, glngSys, 1277)
End Sub

Private Sub PicLeftHead_Resize()
    On Error Resume Next
    tabMain.Move PicLeftHead.ScaleLeft, PicLeftHead.ScaleTop, PicLeftHead.ScaleWidth, PicLeftHead.ScaleHeight
End Sub

Private Sub InitCommandBar()
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar

    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Set Me.cbsThis.Icons = gobjComlib.ZLCommFun.GetPubIcons
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
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): cbrControl.BeginGroup = True
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
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    '设置不常用菜单
    With Me.cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
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
        cbrControl.Style = xtpButtonIconAndCaption
    Next
End Sub

Private Sub InitTabContol()
'功能：初始化TabControl控件
    With tabMain
        With .PaintManager
            .Appearance = xtpTabAppearanceExcel
            .Color = xtpTabColorOffice2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With

        .InsertItem(0, "传染病报告卡", picParameter.hwnd, 0).Tag = "传染病报告卡"
        .InsertItem(1, "疾病证明报告", rptList.hwnd, 0).Tag = "疾病证明报告"
        .Item(1).Selected = True
        .Item(0).Selected = True
    End With
End Sub

Private Sub InitReportControl()
'功能：初始化ReportControl控件
    Dim rptCol As ReportColumn
    With rptList
        Set rptCol = .Columns.Add(mCol.图标, "", 18, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Sortable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mCol.ID, "ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.种类, "种类", 90, False): rptCol.Editable = False: rptCol.Groupable = True: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.编号, "编号", 50, False): rptCol.Editable = False: rptCol.Groupable = False: .SortOrder.Add rptCol
        Set rptCol = .Columns.Add(mCol.名称, "名称", 120, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.说明, "说明", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.保留, "类型", 30, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.子类, "子类", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        
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
End Sub

Private Function zlRefList(Optional lngFileID As Long) As Long
'功能：刷新装入指定种类的病历文件清单，并定位到指定的文件上
    Dim rsTemp As New ADODB.Recordset
    Dim rptRcd As ReportRecord
    Dim rptItem As ReportRecordItem
    Dim rptRow As ReportRow
    Dim strSQL As String

    Me.rptList.Tag = "-1"
    strSQL = "Select l.Id, l.种类, l.编号, l.名称, l.说明, Nvl(l.保留, 0) As 保留,l.子类" & _
            " From 病历文件列表 l  Where l.种类 = 5"

    Err = 0: On Error GoTo errHand
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    rsTemp.Filter = "保留 = 4"
    If rsTemp.RecordCount > 0 Then
        mlngFixedFileID = NVL(rsTemp!ID, 0)
    End If
    rsTemp.Filter = "保留 <> 4"
    rptList.Records.DeleteAll
    With rsTemp
        Do While Not .EOF
            Set rptRcd = Me.rptList.Records.Add()
            Set rptItem = rptRcd.AddItem(CStr(!种类)): rptItem.Icon = rptItem.Value - 1
            rptRcd.AddItem CStr(!ID)
            
            Select Case !种类
                Case 5 And !保留 <> 4: rptRcd.AddItem CStr("疾病证明报告")
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
                        rptRcd.AddItem CStr("特殊")
            End Select
            rptRcd.AddItem CStr(NVL(!子类))
            .MoveNext
        Loop
    End With
    Me.rptList.Populate

    If mblnFileList Then
        If lngFileID <> 0 Then
            For Each rptRow In rptList.Rows
                If rptRow.GroupRow = False Then
                    If Val(rptRow.Record(mCol.ID).Value) = lngFileID Then
                        Set Me.rptList.FocusedRow = rptRow: Exit For
                    End If
                End If
            Next
        End If
        
        If Me.rptList.Rows.Count > 0 Then
            If Me.rptList.FocusedRow Is Nothing Then
                Set Me.rptList.FocusedRow = Me.rptList.Rows(0)
            End If
            If Me.rptList.FocusedRow.GroupRow Then
                lngFileID = 0
            Else
                lngFileID = Me.rptList.FocusedRow.Record.Item(mCol.ID).Value
            End If
        Else
            lngFileID = 0
        End If
        Call mfrmRequest.zlRefresh(lngFileID)
        Call mfrmContent.zlRefresh(lngFileID)
    Else
        Call mfrmRequest.zlRefresh(mlngFixedFileID)
        Call mfrmContent.zlRefresh(mlngFixedFileID)
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

Private Sub rptList_KeyDown(KeyCode As Integer, Shift As Integer)
    If Me.rptList.Visible = False Then Exit Sub
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Me.rptList.FocusedRow Is Nothing Then Exit Sub
    If Me.rptList.FocusedRow.GroupRow Then Exit Sub
    Call rptList_RowDblClick(Me.rptList.FocusedRow, Me.rptList.FocusedRow.Record.Item(mCol.编号))
End Sub

Private Sub rptList_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
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

Private Sub rptList_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Dim cbrControl As CommandBarControl

    With Me.rptList
        If .FocusedRow Is Nothing Then
            mlngCurFileId = 0
            mstrCurFixed = ""
        ElseIf .FocusedRow.GroupRow = True Then
            mlngCurFileId = 0
            mstrCurFixed = ""
        Else
            mlngCurFileId = .FocusedRow.Record.Item(mCol.ID).Value
            mstrCurFixed = .FocusedRow.Record.Item(mCol.保留).Value
        End If
    End With
    If mlngCurFileId = 0 Then Exit Sub

    Set cbrControl = Me.cbsThis.FindControl(, conMenu_Edit_Modify)
    If cbrControl Is Nothing Then Exit Sub
    If cbrControl.Visible = False Or cbrControl.Enabled = False Then Exit Sub
    Call cbsThis_Execute(cbrControl)
End Sub

Private Sub rptList_SelectionChanged()
    With rptList
        If .FocusedRow Is Nothing Then
            mlngCurFileId = 0
            mstrCurFixed = ""
            Call mfrmRequest.zlRefresh(mlngCurFileId)
            Call mfrmContent.zlRefresh(mlngCurFileId)
        ElseIf .FocusedRow.GroupRow = True Then
            mlngCurFileId = 0
            mstrCurFixed = ""
            Call mfrmRequest.zlRefresh(mlngCurFileId)
            Call mfrmContent.zlRefresh(mlngCurFileId)
        Else
            mlngCurFileId = .FocusedRow.Record.Item(mCol.ID).Value
            mstrCurFixed = .FocusedRow.Record.Item(mCol.保留).Value
            If Val(Me.rptList.Tag) <> Me.rptList.FocusedRow.Index Then
                Call mfrmRequest.zlRefresh(mlngCurFileId)
                Call mfrmContent.zlRefresh(mlngCurFileId)
                Me.rptList.Tag = Me.rptList.FocusedRow.Index
            End If
        End If
        mlngFileID = mlngCurFileId
    End With
End Sub

Private Sub tabMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If Not Me.Visible Then Exit Sub

    If Item.Tag = "传染病报告卡" Then
        mlngFileID = mlngFixedFileID
        mblnFileList = False
    ElseIf Item.Tag = "疾病证明报告" Then
        mlngFileID = mlngCurFileId
        mblnFileList = True
    End If
    Call mfrmRequest.zlRefresh(mlngFileID)
    Call mfrmContent.zlRefresh(mlngFileID)
End Sub

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
