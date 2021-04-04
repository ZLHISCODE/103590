VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{FBAFE9A8-8B26-4559-9D12-D70E36A97BE3}#2.1#0"; "zlRichEditor.ocx"
Begin VB.Form frmDockEPRContent 
   BackColor       =   &H80000003&
   BorderStyle     =   0  'None
   Caption         =   "病历文件提纲"
   ClientHeight    =   3585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5940
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picRich 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   3150
      Left            =   705
      ScaleHeight     =   3150
      ScaleWidth      =   4830
      TabIndex        =   0
      Top             =   135
      Width           =   4830
      Begin zlRichEditor.Editor edtThis 
         Height          =   2580
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   4551
         WithViewButtonas=   0   'False
         ShowRuler       =   0   'False
      End
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   120
      Top             =   2985
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   45
      Top             =   45
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
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
    conPane_Infection = 3
    conPane_Annex = 4
End Enum
'-----------------------------------------------------
'窗体变量
'-----------------------------------------------------
Private mlngPatientID As Long       '病人ID
Private mlngRecordId As Long        '病历记录ID
Private mfrmAnnex As frmDockAnnex    '病历附件窗体
Private mObjTabEprView As cTableEPR      '表格病历
Private mobjInfection As Object
Public mIsShowAnnex As Boolean

Private Function CopyEnable() As Integer
On Error GoTo errHand
Dim rsTemp As New ADODB.Recordset
    gstrSQL = "Select Zl_Fun_CopyEnable([1]) CopyEnable From Dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngRecordId)
    If rsTemp!CopyEnable = 1 Then
        CopyEnable = 1
    Else
        CopyEnable = 0
    End If
    
    Exit Function
errHand:
    CopyEnable = 0
End Function

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case ID_EDIT_COPY
        If CopyEnable() = 1 Then
            If Control.Enabled And Control.Visible Then '快捷键执行时需要判断
                gstrCopyPID = CStr(mlngPatientID)
                Me.edtThis.Copy
            End If
        Else
            MsgBox "选定的病历不允许复制", vbInformation, gstrSysName
        End If
    End Select
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case ID_EDIT_COPY
        Control.Enabled = edtThis.Selection.EndPos <> edtThis.Selection.StartPos
        Control.Enabled = edtThis.Selection.getType <> cprSTPicture
        Control.Visible = InStr(gstrPrivsEpr, "内容复制") > 0
    End Select
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_RichEpr
        Item.Handle = picRich.hWnd
    Case conPane_Annex
        If Not mIsShowAnnex Then
             Item.Handle = mfrmAnnex.hWnd
        End If
    Case conPane_TablEpr
        Item.Handle = mObjTabEprView.zlGetForm.hWnd
    Case conPane_Infection
        Item.Handle = mobjInfection.zlGetForm.hWnd
    End Select
End Sub

Private Sub edtThis_RequestRightMenu(ViewMode As zlRichEditor.ViewModeEnum, Shift As Integer, X As Single, Y As Single)
    '没有内容复制权限不允许复制
    If InStr(gstrPrivsEpr, "内容复制") = 0 Then Exit Sub
    
    Dim Popup As CommandBar
    Dim Control As CommandBarControl
    
    Set Popup = cbsThis.Add("Popup", xtpBarPopup)
    With Popup.Controls
        Set Control = .Add(xtpControlButton, ID_EDIT_COPY, "复制(&C)")
        Popup.ShowPopup
    End With
End Sub

Private Sub edtThis_DblClick(ViewMode As zlRichEditor.ViewModeEnum)
    RaiseEvent DblClick
End Sub
Private Sub Form_Load()
Dim Pane1 As Pane, pane2 As Pane, pane3 As Pane, Pane4 As Pane
    On Error GoTo errHand
    Set mObjTabEprView = New cTableEPR
    mObjTabEprView.InitTableEPR gcnOracle, glngSys, gstrDbOwner
    Set mfrmAnnex = New frmDockAnnex
    
    Set Pane1 = dkpMan.CreatePane(conPane_RichEpr, 1200, 200, DockTopOf, Nothing)
    Pane1.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set pane2 = dkpMan.CreatePane(conPane_TablEpr, 1200, 200, DockTopOf, Nothing)
    pane2.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    pane2.Close
    
    Set pane3 = dkpMan.CreatePane(conPane_Infection, 1200, 200, DockTopOf, Nothing)
    pane3.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    pane3.Close
    
    If Not mIsShowAnnex Then
        Set Pane4 = dkpMan.CreatePane(conPane_Annex, 1200, 15, DockBottomOf, Nothing)
        Pane4.MinTrackSize.Height = 360 / Screen.TwipsPerPixelY: Pane4.MaxTrackSize.Height = 360 / Screen.TwipsPerPixelY
        Pane4.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    End If
    
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
    
    cbsThis.ActiveMenuBar.Visible = False
    cbsThis.KeyBindings.Add FCONTROL, Asc("C"), ID_EDIT_COPY
    mlngRecordId = -1
    Exit Sub

errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub Clear()
    On Error Resume Next
    dkpMan.FindPane(conPane_TablEpr).Close
    dkpMan.FindPane(conPane_Infection).Close
    dkpMan.ShowPane conPane_RichEpr
    
    edtThis.Freeze
    edtThis.ReadOnly = False
    edtThis.ForceEdit = True
    edtThis.InProcessing = True
    edtThis.Tag = "LoadFile"
    edtThis.NewDoc
    
    edtThis.UnFreeze
    edtThis.RefreshTargetDC
    edtThis.ViewMode = cprNormal
    edtThis.ReadOnly = True
    edtThis.ForceEdit = False
    edtThis.InProcessing = False
    edtThis.Tag = ""
End Sub

Public Sub zlRefresh(ByVal lngRecordId As Long, strAnnexRight As String, Optional ByVal blnPrivacyProtect As Boolean, _
                Optional ByVal blnMoved As Boolean, Optional ByRef blnViewFile As Boolean, Optional ByVal byteEdit As Byte, _
                Optional ByVal blnAllowDelete As Boolean, Optional ByVal blnClearMode As Boolean)
'功能：刷新病历显示内容；
'参数：lngRecordId：电子病历记录ID；blnPrivacyProtect：是否启用隐私保护;strAnnexRight-附件操作权限,byteEdit=0 RichEdit =1 表格式病历;blnViewFile 是否可以预览
    Dim blnPrivacy As Boolean, Elements As New cEPRElements
    Dim rs As New ADODB.Recordset, lngKey As Long
    
    On Error GoTo errHand
    If blnPrivacyProtect = True Then
        blnPrivacy = InStr(gstrPrivsEpr, ";忽略隐私保护;") = 0     '保护隐私项目
    End If
    
    mlngRecordId = lngRecordId
    dkpMan.FindPane(conPane_RichEpr).Close
    dkpMan.FindPane(conPane_TablEpr).Close
    dkpMan.FindPane(conPane_Infection).Close
    If byteEdit = 1 Then
        dkpMan.ShowPane conPane_TablEpr
        Call mObjTabEprView.InitOpenEPR(Me, cprEM_修改, cprET_单病历审核, mlngRecordId, False, 0, , , , , , , , blnMoved)
        Call mObjTabEprView.zlRefreshDockfrm
        blnViewFile = True
    ElseIf byteEdit = 2 Then '传染病报告卡专用编辑器
        dkpMan.ShowPane conPane_Infection
        gstrSQL = "Select ID,病人ID,主页ID From 电子病历记录 Where ID=[1]"
        If blnMoved Then gstrSQL = Replace(gstrSQL, "电子病历记录", "H电子病历记录")
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngRecordId)
        mobjInfection.zlRefresh rs!病人ID, rs!主页ID, lngRecordId, blnMoved
    ElseIf byteEdit = 0 Then
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
                    gstrSQL = "Select A.ID,A.对象标记 From 电子病历内容 A, 隐私保护项目 B,诊治所见项目 C " & _
                        "Where A.对象类型 = 4 And A.替换域 = 1 And A.文件id = [1] And A.对象序号 > 0 and B.项目id = C.ID And A.要素名称 =C.中文名 And C.替换域 = 1 "
                    If blnMoved Then gstrSQL = Replace(gstrSQL, "电子病历内容", "H电子病历内容")
                    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngRecordId)
                    If Not rs.EOF Then
                        Do While Not rs.EOF
                            lngKey = Elements.Add(NVL(rs("对象标记"), 0))
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
            Call BuildRTF(edtThis, lngRecordId, blnMoved)
            blnViewFile = True
        End If
        
        ' 对最终文档的处理
        If blnClearMode Then
            edtThis.AuditMode = True
            edtThis.AcceptAuditText    '清洁模式 过滤审阅修订痕迹
        End If
        
        If lngRecordId > 0 Then
            '设置页面格式
            Dim mEPRFileInfo As New cEPRFileDefineInfo
            gstrSQL = "Select c.ID, a.格式,c.病人ID From   病历页面格式 a, 病历文件列表 b, 电子病历记录 c " & _
                    " Where  c.文件id = b.id And a.种类 = b.种类 And a.编号 = b.页面 And c.ID = [1]"
            If blnMoved Then gstrSQL = Replace(gstrSQL, "电子病历记录", "H电子病历记录")
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngRecordId)
            If Not rs.EOF Then
                mlngPatientID = rs!病人ID
                mEPRFileInfo.格式 = zlCommFun.NVL(rs("格式").Value)
                mEPRFileInfo.SetFormat Me.edtThis, mEPRFileInfo.格式
                Me.edtThis.ResetWYSIWYG
            End If
            Set mEPRFileInfo = Nothing
        End If
        If strZipFile <> "" Then '有RTF文件才刷新对象
            Call RefreshObject(lngRecordId, blnMoved)
        End If
        Me.edtThis.SelStart = 0
        Me.edtThis.UnFreeze
        Me.edtThis.RefreshTargetDC
        Me.edtThis.ViewMode = cprNormal
        Me.edtThis.ReadOnly = True
        
    End If
    '调用附件列表
    Call mfrmAnnex.zlRefresh(mlngRecordId, strAnnexRight, blnMoved, blnAllowDelete)
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Unload mfrmAnnex
    Set mfrmAnnex = Nothing
    Unload mObjTabEprView.zlGetForm
    Set mObjTabEprView = Nothing
    Unload mobjInfection.zlGetForm
    Set mobjInfection.zlGetForm = Nothing
    Set mobjInfection = Nothing
End Sub

Private Sub picRich_Resize()
    edtThis.Top = 0: edtThis.Left = 0
    edtThis.Width = picRich.ScaleWidth: edtThis.Height = picRich.Height
End Sub
Private Sub RefreshObject(ByVal lngRecordId As Long, ByVal blnMoved As Boolean)
'刷新界面上的图片,目前只刷新图片，有需要时再调整刷新表格
Dim Pictures As New cEPRPictures, rsTemp As New ADODB.Recordset, lngKey As Long, Tables As New cEPRTables
Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean, blnForce As Boolean

    '读取所有的图片
    gstrSQL = "Select ID, 文件id,开始版, 终止版,父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行,预制提纲ID " & _
        "From 电子病历内容 " & _
        "Where 文件id = [1] And 对象类型 in(3,5) And 对象序号 Is Not Null" '不显示表格中的图片
    If blnMoved Then gstrSQL = Replace(gstrSQL, "电子病历内容", "H电子病历内容")
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngRecordId)
    Do Until rsTemp.EOF
        If rsTemp!对象类型 = 5 Then
            lngKey = Pictures.Add(NVL(rsTemp!对象标记, 0))
            Call Pictures("K" & lngKey).FillPictureMember(rsTemp, IIf(blnMoved, "H电子病历内容", "电子病历内容"))
            Call Pictures("K" & lngKey).DeleteFromEditor(edtThis)
            Call Pictures("K" & lngKey).InsertIntoEditor(edtThis, -1, True)
        ElseIf rsTemp!对象类型 = 3 Then
            lngKey = Tables.Add(NVL(rsTemp!对象标记, 0))
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
