VERSION 5.00
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{FBAFE9A8-8B26-4559-9D12-D70E36A97BE3}#2.1#0"; "zlRichEditor.ocx"
Begin VB.Form frmDockEPRContent 
   BackColor       =   &H80000003&
   BorderStyle     =   0  'None
   Caption         =   "病历文件提纲"
   ClientHeight    =   4080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6915
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   0
      ScaleHeight     =   3615
      ScaleWidth      =   6375
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      Begin VB.PictureBox picRich 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   3150
         Left            =   585
         ScaleHeight     =   3150
         ScaleWidth      =   4830
         TabIndex        =   1
         Top             =   0
         Width           =   4830
         Begin zlRichEditor.Editor edtThis 
            Height          =   2580
            Left            =   0
            TabIndex        =   2
            Top             =   0
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   4551
            WithViewButtonas=   0   'False
            ShowRuler       =   0   'False
         End
      End
      Begin XtremeDockingPane.DockingPane dkpMan 
         Left            =   0
         Top             =   2865
         _Version        =   589884
         _ExtentX        =   450
         _ExtentY        =   423
         _StockProps     =   0
         VisualTheme     =   5
      End
   End
End
Attribute VB_Name = "frmDockEPRContent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-----------------------------------------------------
'窗体事件
'-----------------------------------------------------
Public Event DblClick()                                                 '返回双击操作事件

Private Enum FileType
    conPane_RichEpr = 1
    conPane_TablEpr = 2
    conPane_Feedback = 3
    conPane_Infection = 4
End Enum
'-----------------------------------------------------
'窗体变量
'-----------------------------------------------------
Private mlngPatientID As Long       '病人ID
Private mlngRecordId As Long        '病历记录ID
Private mfrmReport  As frmDiseaseRegist    '传染病阳性结果反馈单
Private mObjTabEprView As cTableEPR      '表格病历
Private mobjInfection As Object          '中华人民共和国传染病报告卡

Public mIsShowAnnex As Boolean

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_RichEpr
        Item.Handle = picRich.hWnd
    Case conPane_Feedback
        Item.Handle = mfrmReport.hWnd
    Case conPane_TablEpr
        Item.Handle = mObjTabEprView.zlGetForm.hWnd
    Case conPane_Infection
        Item.Handle = mobjInfection.zlGetForm.hWnd
    End Select
End Sub

Private Sub edtThis_DblClick(ViewMode As zlRichEditor.ViewModeEnum)
    RaiseEvent DblClick
End Sub

Private Sub Form_Load()
    Dim Pane1 As Pane, pane2 As Pane, pane3 As Pane, pane4 As Pane
    On Error GoTo errHand
    Set mObjTabEprView = New cTableEPR
    mObjTabEprView.InitTableEPR gcnOracle, glngSys, gstrDBOwer
    Set mfrmReport = New frmDiseaseRegist
    Call mfrmReport.SetFrmInset(True)
    
    Set Pane1 = dkpMan.CreatePane(conPane_RichEpr, 1200, 200, DockTopOf, Nothing)
    Pane1.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set pane2 = dkpMan.CreatePane(conPane_TablEpr, 1200, 200, DockTopOf, Nothing)
    pane2.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    pane2.Close
    
    Set pane3 = dkpMan.CreatePane(conPane_Feedback, 1200, 200, DockTopOf, Nothing)
    pane3.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    pane3.Close
    
    Set pane3 = dkpMan.CreatePane(conPane_Infection, 1200, 200, DockTopOf, Nothing)
    pane3.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    pane3.Close
    
    With dkpMan
        .VisualTheme = ThemeOffice2003
        .Options.HideClient = True
        .Options.UseSplitterTracker = True
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = False
    End With
        
    Set mobjInfection = DynamicCreate("zlDisReportCard.clsDisReportCard", "传染病报告卡", True)
    If Not mobjInfection Is Nothing Then
        mobjInfection.Init gcnOracle, glngSys
    End If
    
    mlngRecordId = -1
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

'-----------------------------------------------------
'窗体公共方法
'-----------------------------------------------------

Public Sub zlRefresh(ByVal lngRecordId As Long, strAnnexRight As String, Optional ByVal blnPrivacyProtect As Boolean, _
                Optional ByVal blnMoved As Boolean, Optional ByRef blnViewFile As Boolean, Optional ByVal byteEdit As Byte, _
                Optional ByVal blnAllowDelete As Boolean)
'功能：刷新病历显示内容；
'参数：lngRecordId：电子病历记录ID；blnPrivacyProtect：是否启用隐私保护;strAnnexRight-附件操作权限,byteEdit=0 RichEdit =1 表格式病历;blnViewFile 是否可以预览
    Dim blnPrivacy As Boolean, Elements As New cEPRElements
    Dim rs As New ADODB.Recordset, lngKey As Long
    Dim strSQL As String
    
    On Error GoTo errHand
    If blnPrivacyProtect = True Then
        blnPrivacy = InStr(gstrPrivs, ";忽略隐私保护;") = 0     '保护隐私项目
    End If
    
    mlngRecordId = lngRecordId
    dkpMan.FindPane(conPane_RichEpr).Close
    dkpMan.FindPane(conPane_TablEpr).Close
    dkpMan.FindPane(conPane_Feedback).Close
    dkpMan.FindPane(conPane_Infection).Close
    
    If byteEdit = 1 Then
        dkpMan.ShowPane conPane_TablEpr
        Call mObjTabEprView.InitOpenEPR(Me, cprEM_修改, cprET_单病历审核, mlngRecordId, False, 0)
        Call mObjTabEprView.zlRefreshDockfrm
        blnViewFile = True
    ElseIf byteEdit = 2 Then '传染病报告卡专用编辑器
        dkpMan.ShowPane conPane_Infection
        strSQL = "Select ID,病人ID,主页ID From 电子病历记录 Where ID=[1]"
        If blnMoved Then strSQL = Replace(strSQL, "电子病历记录", "H电子病历记录")
        Set rs = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngRecordId)
        mobjInfection.zlRefresh rs!病人ID, rs!主页ID, lngRecordId, blnMoved
    ElseIf byteEdit = 3 Then '传染病阳性结果反馈单
        dkpMan.ShowPane conPane_Feedback
        Call mfrmReport.zlRefresh(mlngRecordId)
        Call mfrmReport.SetReportTop(100)
    Else
        dkpMan.ShowPane conPane_RichEpr
        Dim strTemp As String, strZipFile As String
        Me.edtThis.Freeze
        Me.edtThis.ReadOnly = False
        Me.edtThis.NewDoc
        strZipFile = zlBlobRead(5, lngRecordId, , blnMoved)
        If gobjFSO.FileExists(strZipFile) Then
            strTemp = zlFileUnzip(strZipFile)
            If gobjFSO.FileExists(strTemp) Then
                '打开文件
                Me.edtThis.OpenDoc strTemp
                '设置替换项目
                If blnPrivacy Then
                    '读取所有的要素
                    strSQL = "Select A.ID,A.对象标记 From 电子病历内容 A, 隐私保护项目 B,诊治所见项目 C " & _
                        "Where A.对象类型 = 4 And A.替换域 = 1 And A.文件id = [1] And A.对象序号 > 0 and B.项目id = C.ID And A.要素名称 =C.中文名 And C.替换域 = 1 "
                    If blnMoved Then strSQL = Replace(strSQL, "电子病历内容", "H电子病历内容")
                    Set rs = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngRecordId)
                    If Not rs.EOF Then
                        Do While Not rs.EOF
                            lngKey = Elements.Add(Nvl(rs("对象标记"), 0))
                            Elements("K" & lngKey).GetElementFromDB cprET_单病历编辑, rs("ID"), True, IIf(blnMoved, "H电子病历内容", "电子病历内容")
                            '替换要素内容
                            Elements("K" & lngKey).内容文本 = String(Len(Elements("K" & lngKey).内容文本), "*")
                            Elements("K" & lngKey).Refresh Me.edtThis
                            rs.MoveNext
                        Loop
                    End If
                    rs.Close
                End If
                gobjFSO.DeleteFile strTemp, True
            End If
            gobjFSO.DeleteFile strZipFile, True
            Me.edtThis.SelStart = 0
            blnViewFile = True
        Else
            blnViewFile = False
        End If
        If lngRecordId > 0 Then
            '设置页面格式
            Dim mEPRFileInfo As New cEPRFileDefineInfo
            strSQL = "Select c.ID, a.格式,c.病人ID From   病历页面格式 a, 病历文件列表 b, 电子病历记录 c " & _
                    " Where  c.文件id = b.id And a.种类 = b.种类 And a.编号 = b.页面 And c.ID = [1]"
            If blnMoved Then strSQL = Replace(strSQL, "电子病历记录", "H电子病历记录")
            Set rs = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngRecordId)
            If Not rs.EOF Then
                mlngPatientID = rs!病人ID
                mEPRFileInfo.格式 = gobjComlib.zlCommFun.Nvl(rs("格式").Value)
                mEPRFileInfo.SetFormat Me.edtThis, mEPRFileInfo.格式
                Me.edtThis.ResetWYSIWYG
            End If
            Set mEPRFileInfo = Nothing
        End If
        Call RefreshObject(lngRecordId, blnMoved)
        Me.edtThis.SelStart = 0
        Me.edtThis.UnFreeze
        edtThis.RefreshTargetDC
        Me.edtThis.ViewMode = cprNormal
        Me.edtThis.ReadOnly = True
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    picMain.Move Me.ScaleLeft, Me.ScaleTop, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Unload mfrmReport
    Set mfrmReport = Nothing
    Unload mObjTabEprView.zlGetForm
    Set mObjTabEprView = Nothing
    Unload mobjInfection.zlGetForm
    Set mobjInfection.zlGetForm = Nothing
    Set mobjInfection = Nothing
End Sub

Private Sub picRich_Resize()
On Error Resume Next
    edtThis.Top = 0: edtThis.Left = 0
    edtThis.Width = picRich.ScaleWidth: edtThis.Height = picRich.Height
End Sub

Private Sub RefreshObject(ByVal lngRecordId As Long, ByVal blnMoved As Boolean)
'刷新界面上的图片,目前只刷新图片，有需要时再调整刷新表格
    Dim Pictures As New cEPRPictures, rsTemp As New ADODB.Recordset, lngKey As Long, Tables As New cEPRTables
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean, blnForce As Boolean
    Dim strSQL As String
    '读取所有的图片
    strSQL = "Select ID, 文件id,开始版, 终止版,父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行,预制提纲ID " & _
        "From 电子病历内容 " & _
        "Where 文件id = [1] And 对象类型 in(3,5) And 对象序号 Is Not Null" '不显示表格中的图片
    If blnMoved Then strSQL = Replace(strSQL, "电子病历内容", "H电子病历内容")
    Set rsTemp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngRecordId)
    Do Until rsTemp.EOF
        If rsTemp!对象类型 = 5 Then
            lngKey = Pictures.Add(Nvl(rsTemp!对象标记, 0))
            Call Pictures("K" & lngKey).FillPictureMember(rsTemp, IIf(blnMoved, "H电子病历内容", "电子病历内容"))
            Call Pictures("K" & lngKey).DeleteFromEditor(edtThis)
            Call Pictures("K" & lngKey).InsertIntoEditor(edtThis, -1, True)
        ElseIf rsTemp!对象类型 = 3 Then
            lngKey = Tables.Add(Nvl(rsTemp!对象标记, 0))
            Call Tables("K" & lngKey).FillTableMember(rsTemp, IIf(blnMoved, "H电子病历内容", "电子病历内容"))
            
            If Tables("K" & lngKey).Cells.Count = 1 Then
                '一个单元格，可能是PACS编辑器书写的内容
                If FindKey(edtThis, "T", lngKey, lKSS, lKSE, lKES, lKEE, True) Then
                    '先删除
                    Call Tables("K" & lngKey).DeleteFromEditor(edtThis)
                    With edtThis
                        blnForce = .ForceEdit
                        .InProcessing = True
                        .Tag = "TableSingleCell:InsertIntoEditor"
                        .ForceEdit = True
                        .Range(lKSS, lKSS).Font.Protected = False
                        .Range(lKSS, lKSS).Font.Hidden = False
                        .Range(lKSS, lKSS) = Tables("K" & lngKey).Cells(1).内容文本
                        .ForceEdit = blnForce
                        .UnFreeze
                        .InProcessing = False
                        .Tag = ""
                    End With
                End If
            Else
                '多个单元格
                '先删除
                Call Tables("K" & lngKey).DeleteFromEditor(edtThis)
                Call Tables("K" & lngKey).InsertIntoEditor(edtThis, -1)
            End If
        End If
        rsTemp.MoveNext
    Loop
End Sub





