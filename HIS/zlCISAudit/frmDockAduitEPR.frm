VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FBAFE9A8-8B26-4559-9D12-D70E36A97BE3}#2.1#0"; "zlRichEditor.ocx"
Begin VB.Form frmDockAduitEPR 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6045
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picThis 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   630
      Left            =   5865
      ScaleHeight     =   630
      ScaleWidth      =   1395
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   705
      Index           =   1
      Left            =   1080
      ScaleHeight     =   705
      ScaleWidth      =   6135
      TabIndex        =   2
      Top             =   3285
      Visible         =   0   'False
      Width           =   6135
      Begin MSComctlLib.ListView lvwThis 
         Height          =   330
         Left            =   495
         TabIndex        =   3
         Top             =   0
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   582
         View            =   1
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FlatScrollBar   =   -1  'True
         _Version        =   393217
         Icons           =   "imgThis"
         SmallIcons      =   "imgThis"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "文件"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Label lblThis 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "附件:"
         Height          =   180
         Left            =   0
         TabIndex        =   4
         Top             =   75
         Width           =   450
      End
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   1980
      Index           =   0
      Left            =   1350
      ScaleHeight     =   1980
      ScaleWidth      =   2685
      TabIndex        =   0
      Top             =   705
      Width           =   2685
      Begin zlRichEditor.Editor edtThis 
         Height          =   1245
         Left            =   315
         TabIndex        =   1
         Top             =   315
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   2196
      End
   End
   Begin MSComctlLib.ImageList imgThis 
      Left            =   5610
      Top             =   885
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   360
      Top             =   1905
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
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
Attribute VB_Name = "frmDockAduitEPR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################
Private mblnAddition As Boolean
Private mlngKey As Long

Private Enum FileType
    conPane_RichEpr = 1
    conPane_TablEpr = 2
    conPane_Annex = 3
End Enum
Private Enum ICON_SIZE
    ICON_SMALL = 16
    ICON_LARGE = 32
End Enum

Public Event PrintEpr(ByVal lngRecordId As Long)

Private mObjTabEprView As cTableEPR      '表格病历
Private WithEvents mfrmPrintPreview As frmPrintPreview
Attribute mfrmPrintPreview.VB_VarHelpID = -1
Private mblnDataMove As Boolean
Private mfrmMain As Object
Private mfrmAnnex As New frmDockAnnex
Private Function SetRichDocsPos(ByVal lngRecordId As Long) As Boolean
    '通过ID先定位，无法定位时再加载
    Dim lngKSS As Long, lngKSE As Long, lngKES As Long, lngKEE As Long, blnNeed As Boolean, lngKey As Long, lngLen As Long, i As Integer
    lngLen = Len(edtThis.Text)
    For i = 0 To lngLen
        If FindNextKey(edtThis, i, "F", lngKey, lngKSS, lngKSE, lngKES, lngKEE, blnNeed) Then
            If edtThis.Range(lngKSE, lngKES).Text = lngRecordId Then
                edtThis.Range(lngKEE + 1, lngKEE + 1).Selected
                SetRichDocsPos = True
                Exit Function
            End If
            i = lngKEE
        Else
            Exit Function
        End If
    Next
End Function
Public Sub zlRefreshData(ByVal lngRecordId As Long, strAnnexRight As String, _
                Optional ByVal blnMoved As Boolean, Optional ByRef blnViewFile As Boolean, Optional ByVal byteEdit As Byte, _
                Optional ByVal blnAllowDelete As Boolean)
'功能：显示共享病历，刷新病历显示内容；
'参数：lngRecordId：电子病历记录ID；blnPrivacyProtect：是否启用隐私保护;strAnnexRight-附件操作权限,byteEdit=0 RichEdit =1 表格式病历;blnViewFile 是否可以预览
    Dim blnPrivacy As Boolean
    Dim rs As New ADODB.Recordset
    Dim collFile As New Collection, lngLen1 As Long, lngLen2 As Long, i As Integer, lngFileID As Long, strIDs As String, lngStart As Long, StrKey As String
    
    On Error GoTo errHand
    If lngRecordId = 0 Then Exit Sub
    
    If byteEdit = 1 Then
        dkpMain.FindPane(conPane_RichEpr).Close
        dkpMain.ShowPane conPane_TablEpr
        Call mObjTabEprView.InitOpenEPR(Me, cprEM_修改, cprET_单病历审核, lngRecordId, False, 0)
        Call mObjTabEprView.zlRefreshDockfrm
        blnViewFile = True
    Else
        dkpMain.FindPane(conPane_TablEpr).Close
        dkpMain.ShowPane conPane_RichEpr
        If SetRichDocsPos(lngRecordId) Then Exit Sub
        
        '共享文档加载
        gstrSQL = "Select Count(C.Id) As 数目, c.病人ID,c.主页ID, c.文件id, c.创建时间" & vbNewLine & _
                "From 病历文件列表 F, 病历文件列表 B, 电子病历记录 C" & vbNewLine & _
                "Where f.种类 = b.种类 And f.页面 = b.页面 And b.Id = c.文件id And c.Id = [1]" & vbNewLine & _
                "Group By c.病人ID,c.主页ID, c.文件id, c.创建时间"
        If blnMoved Then gstrSQL = Replace(gstrSQL, "电子病历记录", "H电子病历记录")
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngRecordId)
        lngFileID = rs!文件ID
        edtThis.Freeze
        edtThis.ReadOnly = False
        edtThis.ForceEdit = True
        edtThis.InProcessing = True
        edtThis.Tag = "LoadFile"
        edtThis.NewDoc
        
        If rs!数目 = 1 Then
            '读取RTF文件
            Call ReadRTF(edtThis, lngRecordId, True, blnMoved)
        Else
            zlCommFun.ShowFlash "请稍待，正在加载病历内容！"
            '读取共享页面文件ID
            strIDs = GetFileRange(rs!文件ID, lngRecordId, Format(rs!创建时间, "yyyy-MM-dd HH:mm:ss"), 2, rs!病人ID, rs!主页ID, blnMoved)
            gfrmPublic.edtBuff.Freeze
            gfrmPublic.edtBuff.ReadOnly = False
            gfrmPublic.edtBuff.ForceEdit = True
            gfrmPublic.edtBuff.InProcessing = True
            gfrmPublic.edtBuff.Tag = "LoadFile"
            For i = 0 To UBound(Split(strIDs, ","))
                '读取RTF文件
                Call ReadRTF(gfrmPublic.edtBuff, Split(strIDs, ",")(i), True, blnMoved)
                
                '记录文件ID
                StrKey = "FS(" & Format(i, "00000000") & ",1,0)" & Split(strIDs, ",")(i) & "FE(" & Format(i, "00000000") & ",1,0)"
                'lngLen2 = Len(edtThis.Text) '将文件添加到主文档末尾
                gfrmPublic.edtBuff.Range(0, 0).Selected
                gfrmPublic.edtBuff.Range(0, 0).Text = StrKey
                gfrmPublic.edtBuff.Range(0, 0 + Len(StrKey) + Len(Split(strIDs, ",")(i))).Font.Protected = True
                gfrmPublic.edtBuff.Range(0, 0 + Len(StrKey) + Len(Split(strIDs, ",")(i))).Font.Hidden = True
                
                '追加RTF文件
                lngLen1 = Len(gfrmPublic.edtBuff.Text) '记录临时文件开始、结束位置
                lngLen2 = Len(edtThis.Text) '将文件添加到主文档末尾
                edtThis.Range(lngLen2, lngLen2).Font.Protected = False
                edtThis.Range(lngLen2, lngLen2).Selected
                gfrmPublic.edtBuff.SelectAll
                gfrmPublic.edtBuff.CopyWithFormat
                edtThis.PasteWithFormat
                lngStart = Len(edtThis.Text)
                If i < UBound(Split(strIDs, ",")) Then
                    '只要不是最后一份文件，末尾保证有一个回车，以备追加下一个文件
                    If edtThis.Range(lngStart - 2, lngStart) = vbCrLf Then
                        edtThis.Range(lngStart - 2, lngStart).Font.Hidden = False
                    Else
                        edtThis.Range(lngStart, lngStart).Text = vbCrLf
                        edtThis.Range(lngStart, lngStart + 2).Font.Hidden = False
                    End If
                End If
                edtThis.TOM.TextDocument.Range(lngStart, lngStart).Para = gfrmPublic.edtBuff.TOM.TextDocument.Range(lngLen1, lngLen1).Para '.Duplicate
            Next
        End If
        
        If lngRecordId > 0 Then
            '设置页面格式
            Dim mEPRFileInfo As New cEPRFileDefineInfo
            gstrSQL = "Select a.格式 From 病历页面格式 a, 病历文件列表 b" & _
                    " Where b.id=[1] And a.种类 = b.种类 And a.编号 = b.页面"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngFileID)
            If Not rs.EOF Then
                mEPRFileInfo.格式 = zlCommFun.NVL(rs("格式").Value)
                mEPRFileInfo.SetFormat Me.edtThis, mEPRFileInfo.格式
                Me.edtThis.ResetWYSIWYG
            End If
            Set mEPRFileInfo = Nothing
        End If
        gfrmPublic.edtBuff.UnFreeze
        gfrmPublic.edtBuff.ForceEdit = False
        edtThis.SelStart = 0
        edtThis.UnFreeze
        edtThis.RefreshTargetDC
        edtThis.ViewMode = cprNormal
        edtThis.ReadOnly = True
        edtThis.ForceEdit = False
        edtThis.InProcessing = False
        edtThis.Tag = ""
        Call SetRichDocsPos(lngRecordId)
    End If
    '调用附件列表
    Call mfrmAnnex.zlRefresh(lngRecordId, strAnnexRight, blnMoved, blnAllowDelete)
    zlCommFun.StopFlash
    Exit Sub
errHand:
    zlCommFun.StopFlash
    If ErrCenter = 1 Then
        Resume
    End If
    On Error Resume Next
    gfrmPublic.edtBuff.UnFreeze
    gfrmPublic.edtBuff.ForceEdit = False
    edtThis.SelStart = 0
    edtThis.UnFreeze
    edtThis.RefreshTargetDC
    edtThis.ViewMode = cprNormal
    edtThis.ReadOnly = True
    edtThis.ForceEdit = False
    edtThis.InProcessing = False
    edtThis.Tag = ""
    Err.Clear
End Sub
'######################################################################################################################
Public Function zlRefresh(ByVal frmMain As Object, ByVal lngKey As Long, ByVal blnDataMove As Boolean) As Boolean
    Dim rs As New ADODB.Recordset
    Set mfrmMain = frmMain
    mlngKey = lngKey
    mblnDataMove = blnDataMove
    
    LockWindowUpdate Me.hWnd

    gstrSQL = "Select 1 From 电子病历格式 Where 文件id = [1] And 内容 Is Not Null"
    If mblnDataMove Then gstrSQL = Replace(gstrSQL, "电子病历格式", "H电子病历格式")
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
    If Not rs.EOF Then
        Call OpenEPR(lngKey)
    Else
        Call OpenSignleEPR(lngKey)
    End If
    
    '检查是否有病历附件
    gstrSQL = "Select 1 From 电子病历附件 Where 病历id = [1]"
    If mblnDataMove Then gstrSQL = Replace(gstrSQL, "电子病历附件", "H电子病历附件")
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
    If rs.BOF = False Then
        Call ShowAdition(lngKey)
        
        dkpMain.ShowPane conPane_Annex
    Else
        dkpMain.FindPane(conPane_Annex).Close
    End If
    
    LockWindowUpdate 0
End Function

Public Function zlPrintDocument(ByVal eDocType As EPRDocTypeEnum, Optional ByVal bytMode As Byte = 2, Optional ByVal lngKey As Long, Optional ByVal strPrintDeviceName As String) As Boolean
    
    '1-预览,2-打印
    Dim lngEPRKey As Long
    Dim rs As New ADODB.Recordset
    Dim frmP As New frmPrintPreview
    Dim strReportCode As String
    Dim lng医嘱id As Long
    Dim lng病人ID As Long
    Dim rsTemp As New ADODB.Recordset
    Dim intOutMode As Integer
    Dim strExseNo As String, intExseKind As Integer

    If eDocType = cpr诊疗报告 Then
    
         gstrSQL = "Select f.通用, f.编号 From 电子病历记录 l, 病历文件列表 f Where l.文件id = f.Id And l.Id = [1]"
         Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
         If rsTemp.RecordCount > 0 Then
            intOutMode = Val("" & rsTemp!通用)
         End If
        If mblnDataMove Then
            gstrSQL = Replace(gstrSQL, "病人医嘱记录", "H病人医嘱记录")
            gstrSQL = Replace(gstrSQL, "病人医嘱发送", "H病人医嘱发送")
        End If
        gstrSQL = "select b.记录性质,b.no,b.医嘱id,c.病人id from 病人医嘱报告 a,病人医嘱发送 b ,病人医嘱记录 c where a.病历id=[1] and a.医嘱id=b.医嘱id and c.id=b.医嘱id"
        If mblnDataMove Then gstrSQL = Replace(gstrSQL, "病人医嘱发送", "H病人医嘱发送")
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
    End If

    If intOutMode = 2 Then
        '采用自定义报表打印
        strReportCode = "ZLCISBILL" & Format(rsTemp!编号, "00000") & "-2"
        If rs.RecordCount > 0 Then
            strExseNo = "" & rs!NO
            intExseKind = Val("" & rs!记录性质)
        End If
        If Check检验行(lngKey, lng医嘱id, lng病人ID) Then
            If Open_LIS_Report(Me, lng医嘱id, lng病人ID, False, IIf(bytMode = 1, False, True)) Then Exit Function
        End If
        Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\zl9Report\LocalSet\" & strReportCode, "Printer", strPrintDeviceName)
        Call ReportOpen(gcnOracle, glngSys, strReportCode, Me, "NO=" & CStr(strExseNo), "性质=" & Val(intExseKind), bytMode)
    Else
        'EPR打印
        gstrSQL = "select a.病人ID,a.主页ID,a.病历种类,a.病人来源,a.编辑方式 from 电子病历记录 a where a.ID=[1] "
        If mblnDataMove Then gstrSQL = Replace(gstrSQL, "电子病历记录", "H电子病历记录")

        lngEPRKey = IIf(lngKey > 0, lngKey, mlngKey)

        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngEPRKey)
        If Not rs.EOF Then
            If NVL(rs!编辑方式, 0) = 0 Then
                eDocType = Val(rs!病历种类)
                If strPrintDeviceName <> "" Then
                    frmP.DoMultiDocPreview mfrmMain, eDocType, zlCommFun.NVL(rs("病人ID").Value, 0), zlCommFun.NVL(rs("主页ID").Value, 0), zlCommFun.NVL(rs("病历种类").Value, 1), "", lngEPRKey, IIf(bytMode = 1, False, True), False, True, mblnDataMove, , strPrintDeviceName
                Else
                    frmP.DoMultiDocPreview mfrmMain, eDocType, zlCommFun.NVL(rs("病人ID").Value, 0), zlCommFun.NVL(rs("主页ID").Value, 0), zlCommFun.NVL(rs("病历种类").Value, 1), "", lngEPRKey, IIf(bytMode = 1, False, True), False, False, mblnDataMove
                End If
                Unload frmP 'ByZT:窗体Load了未显示，没有人为关闭的情况下VB不会自动Unload
                Set frmP = Nothing
            Else
                Call mObjTabEprView.InitOpenEPR(Me, cprEM_修改, cprET_单病历审核, lngEPRKey, False, 0, rs!病人来源)
                mObjTabEprView.zlPrintDoc Me, bytMode = 1
            End If
        End If
    End If
    

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
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    cbsMain.ActiveMenuBar.Title = "菜单栏"
    
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeBlue
    cbsMain.VisualTheme = xtpThemeOffice2003
    With cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = False
        .SetIconSize False, 16, 16
        .UseSharedImageList = False 'ImageList方式时,因同一App中共享,在AddImageList之前设置为False
    End With

    '------------------------------------------------------------------------------------------------------------------
    '菜单定义:包括公共部份，请对xtpControlPopup类型的命令ID重新赋值

    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.Visible = False
    
    '------------------------------------------------------------------------------------------------------------------
    '工具栏定义:包括公共部份
    
    Set objBar = cbsMain.Add("标准", xtpBarTop)
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
    
    Set objControl = objBar.Controls.Add(xtpControlButton, conMenu_File_Open, "打开查阅详细内容...")
    cbsMain.KeyBindings.Add FCONTROL, Asc("C"), ID_EDIT_COPY
    
End Function
Private Function OpenSignleEPR(ByVal lngEPRid As Long) As Boolean
    Dim strPath As String, strFile As String, lngs As Long, lnge As Long
    Dim rs As New ADODB.Recordset
    Dim Doc As New cEPRDocument, Elements As New cEPRElements
    Dim lng病人ID As Long, lng主页ID As Long, byt病历种类 As EPRDocTypeEnum, lng编辑方式 As Long, lngKey As Long, blnPrivacy As Boolean
    
    On Error GoTo errHand
    
    zlCommFun.ShowFlash "请稍候..."
    lngs = GetTickCount
    Screen.MousePointer = vbHourglass
    DoEvents
    LockWindowUpdate Me.hWnd

    gstrSQL = "select 病人ID,主页ID,病历种类,编辑方式 from 电子病历记录 where ID=[1]"
    If mblnDataMove Then
        gstrSQL = Replace(gstrSQL, "电子病历记录", "H电子病历记录")
    End If
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngEPRid)
    If Not rs.EOF Then
        lng病人ID = NVL(rs("病人ID").Value, 0)
        lng主页ID = NVL(rs("主页ID").Value, 0)
        byt病历种类 = NVL(rs("病历种类").Value, 1)
        lng编辑方式 = NVL(rs("编辑方式").Value, 0)
    End If
    rs.Close
    
    edtThis.ForceEdit = True
    
    '保存临时文件
    strPath = IIf(Environ$("tmp") <> vbNullString, Environ$("tmp"), Environ$("temp"))
    strFile = strPath & "\" & App.hInstance & CLng(Timer) & ".TMP"
    
    Doc.InitEPRDoc cprEM_修改, cprET_单病历审核, lngEPRid, byt病历种类, lng病人ID, CStr(lng主页ID), , , , mblnDataMove
    Doc.OpenEPRDoc Doc.frmEditor.Editor1        '打开该文件
    
    '设置替换项目
    If blnPrivacy Then
        '读取所有的要素
        gstrSQL = "Select A.ID,A.对象标记 From 电子病历内容 A, 隐私保护项目 B,诊治所见项目 C " & _
            "Where A.对象类型 = 4 And A.替换域 = 1 And A.文件id = [1] And A.对象序号 > 0 and B.项目id = C.ID And A.要素名称 =C.中文名 And C.替换域 = 1 "
            
        If mblnDataMove Then
            gstrSQL = Replace(gstrSQL, "电子病历内容", "H电子病历内容")
        End If
        
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngEPRid)
        If Not rs.EOF Then
            Do While Not rs.EOF
                lngKey = Elements.Add(NVL(rs("对象标记"), 0))
                Elements("K" & lngKey).GetElementFromDB cprET_单病历编辑, rs("ID"), True, "电子病历内容"
                '替换要素内容
                Elements("K" & lngKey).内容文本 = String(Len(Elements("K" & lngKey).内容文本), "*")
                Elements("K" & lngKey).Refresh Doc.frmEditor.Editor1
                rs.MoveNext
            Loop
        End If
        rs.Close
    End If
    Doc.frmEditor.SaveDocToFile strFile, False     '存储非清洁临时文件

    With edtThis
        If lng编辑方式 = 0 Then
            dkpMain.FindPane(conPane_TablEpr).Close
            dkpMain.ShowPane conPane_RichEpr
        Else
            dkpMain.FindPane(conPane_RichEpr).Close
            dkpMain.ShowPane conPane_TablEpr
        End If
        .Freeze
        .ReadOnly = False
        .NewDoc
        .ForceEdit = True
        .ViewMode = cprNormal
        .OpenDoc strFile

        '设置页眉页脚
        Set .Picture = Doc.frmEditor.Editor1.Picture
        .Head = Doc.frmEditor.Editor1.Head
        .Foot = Doc.frmEditor.Editor1.Foot
        
'        Doc.frmEditor.Editor1.DocHeadCopyWithFormat         '从类编辑器中Copy富格式页眉页脚
'        .DocHeadPasteWithFormat                             '控件页眉页脚粘贴
'        Doc.frmEditor.Editor1.DocFootCopyWithFormat
'        .DocFootPasteWithFormat
'        Call Doc.GetReplacedHeadFootString(edThis)          '用类方法处理控件页眉页脚中的要素
                
        .PaperWidth = Doc.frmEditor.Editor1.PaperWidth
        .PaperHeight = Doc.frmEditor.Editor1.PaperHeight
        .MarginLeft = Doc.frmEditor.Editor1.MarginLeft
        .MarginRight = Doc.frmEditor.Editor1.MarginRight
        .MarginTop = Doc.frmEditor.Editor1.MarginTop
        .MarginBottom = Doc.frmEditor.Editor1.MarginBottom

        '设置页面格式
        Doc.EPRFileInfo.SetFormat edtThis, Doc.EPRFileInfo.格式
        edtThis.ResetWYSIWYG    '刷新所见即所得（WYSIWYG）显示

        '分页
        .SelectAll
        .AuditMode = True
        .AcceptAuditText
        .ViewMode = cprNormal
        .Range(0, 0).Selected
        .ForceEdit = False
        .UnFreeze
        .ReadOnly = True
    End With

    If gobjFSO.FileExists(strFile) Then gobjFSO.DeleteFile strFile    '删除临时文件
 
    Doc.frmEditor.Editor1.Modified = False
    
    Set rs = Nothing
    lnge = GetTickCount
    Debug.Print "读取耗时" & lnge - lngs
    LockWindowUpdate 0
    zlCommFun.StopFlash
    Screen.MousePointer = vbDefault
    
    OpenSignleEPR = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function OpenEPR(ByVal lngEPRid As Long) As Boolean
'******************************************************************************************************************
'功能：刷新病历显示内容；
'参数：lngEPRId-电子病历记录ID
'******************************************************************************************************************

Dim mstrPrivs As String, blnPrivacy As Boolean, Elements As New cEPRElements
Dim rs As New ADODB.Recordset, lngKey As Long
Dim strTemp As String, strZipFile As String
    
    zlCommFun.ShowFlash "请稍等，正在读取病历文件...", Me
    gstrSQL = "Select 编辑方式 From 电子病历记录 Where ID=[1]"
    If mblnDataMove Then gstrSQL = Replace(gstrSQL, "电子病历记录", "H电子病历记录")
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Name, lngEPRid)
    
    If NVL(rs!编辑方式, 0) = 1 Then
        dkpMain.FindPane(conPane_RichEpr).Close
        dkpMain.ShowPane conPane_TablEpr
        Call mObjTabEprView.InitOpenEPR(Me, cprEM_修改, cprET_单病历审核, lngEPRid, False, 0)
        Call mObjTabEprView.zlRefreshDockfrm
    Else
        dkpMain.FindPane(conPane_TablEpr).Close
        dkpMain.ShowPane conPane_RichEpr
        edtThis.Freeze
        edtThis.ReadOnly = False
        edtThis.ForceEdit = True
        edtThis.NewDoc
        Call ReadRTFFile(lngEPRid)

        If lngEPRid > 0 Then
            '设置页面格式
            Dim mEPRFileInfo As New cEPRFileDefineInfo
            gstrSQL = "Select c.ID, a.格式 From   病历页面格式 a, 病历文件列表 b, 电子病历记录 c " & _
                    " Where  c.文件id = b.id And a.种类 = b.种类 And a.编号 = b.页面 And c.ID = [1]"
                    
            If mblnDataMove Then gstrSQL = Replace(gstrSQL, "电子病历记录", "H电子病历记录")
            
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngEPRid)
            If Not rs.EOF Then
                mEPRFileInfo.格式 = zlCommFun.NVL(rs("格式").Value)
                mEPRFileInfo.SetFormat edtThis, mEPRFileInfo.格式
                edtThis.ResetWYSIWYG
            End If
            Set mEPRFileInfo = Nothing
        End If
        Call RefreshObject(lngEPRid, mblnDataMove)
        edtThis.ForceEdit = False
        edtThis.SelStart = 0
        edtThis.UnFreeze
        edtThis.RefreshTargetDC
        edtThis.ViewMode = cprNormal
        edtThis.ReadOnly = True
    End If
    
    OpenEPR = True
    zlCommFun.StopFlash
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    zlCommFun.StopFlash
    Call SaveErrLog
End Function
Private Sub ReadRTFFile(ByVal lngID As Long)
Dim rs As New ADODB.Recordset, strFile As String, strRtf As String, lngLen1 As Long, lngLen2 As Long, lngStart As Long
Dim strIDs As String
    On Error GoTo errHand
    gstrSQL = "Select Count(C.Id) As 数目, c.Id, c.病历名称, c.文件id, c.创建时间,c.病人ID,c.主页ID, B.页面" & vbNewLine & _
                "From 病历文件列表 F, 病历文件列表 B, 电子病历记录 C" & vbNewLine & _
                "Where f.种类 = b.种类 And f.页面 = b.页面 And b.Id = c.文件id And c.Id = [1]" & vbNewLine & _
                "Group By c.Id, c.病历名称, c.文件id, c.创建时间, c.病人id, c.主页id, B.页面"
    If mblnDataMove Then gstrSQL = Replace(gstrSQL, "电子病历记录", "H电子病历记录")
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngID)
    
    If rs!数目 = 1 Then '独立页面直接打印
        strFile = zlBlobRead(5, lngID)
        If gobjFSO.FileExists(strFile) Then
            strRtf = zlFileUnzip(strFile)
            If gobjFSO.FileExists(strRtf) Then
                edtThis.OpenDoc strRtf '打开文件
                gobjFSO.DeleteFile strRtf, True
            End If
            gobjFSO.DeleteFile strFile, True
        End If
    Else
        '读取共享页面的文件ID
        strIDs = GetFileRange(rs!文件ID, lngID, Format(rs!创建时间, "yyyy-MM-dd HH:mm:ss"), cpr住院病历, rs!病人ID, rs!主页ID, mblnDataMove)
        '读取共享页面的文件ID
        gstrSQL = "Select /*+ rule*/ a.Id, a.文件id, a.病历名称, a.最后版本, a.保存人, a.完成时间, a.保存时间" & vbNewLine & _
                "From 电子病历记录 A, Table(Cast(f_Num2list([1]) As Zltools.t_Numlist)) B" & vbNewLine & _
                "Where a.Id = b.Column_Value" & vbNewLine & _
                "Order By a.序号, a.创建时间"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strIDs)
        gfrmPublic.edtBuff.Freeze
        gfrmPublic.edtBuff.ForceEdit = True
        Do Until rs.EOF
            strFile = zlBlobRead(5, rs!ID)
            If gobjFSO.FileExists(strFile) Then
                strRtf = zlFileUnzip(strFile)
                If gobjFSO.FileExists(strRtf) Then
                    gfrmPublic.edtBuff.OpenDoc strRtf '打开文件
                    gobjFSO.DeleteFile strRtf, True
                End If
                gobjFSO.DeleteFile strFile, True
            End If
            '追加RTF文件
            lngLen1 = Len(gfrmPublic.edtBuff.Text) '记录临时文件开始、结束位置
            lngLen2 = Len(edtThis.Text) '将文件添加到主文档末尾
            edtThis.Range(lngLen2, lngLen2).Selected
            gfrmPublic.edtBuff.SelectAll
            gfrmPublic.edtBuff.CopyWithFormat
            edtThis.PasteWithFormat
            lngStart = Len(edtThis.Text)
            If rs.AbsolutePosition < rs.RecordCount Then
                '只要不是最后一份文件，末尾保证有一个回车，以备追加下一个文件
                If edtThis.Range(lngStart - 2, lngStart) = vbCrLf Then
                    edtThis.Range(lngStart - 2, lngStart).Font.Hidden = False
                Else
                    edtThis.Range(lngStart, lngStart).Text = vbCrLf
                    edtThis.Range(lngStart, lngStart + 2).Font.Hidden = False
                End If
            End If
            edtThis.TOM.TextDocument.Range(lngStart, lngStart).Para = gfrmPublic.edtBuff.TOM.TextDocument.Range(lngLen1, lngLen1).Para '.Duplicate
            rs.MoveNext
        Loop
        gfrmPublic.edtBuff.UnFreeze
        gfrmPublic.edtBuff.ForceEdit = False
        Unload gfrmPublic
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Function ShowAdition(ByVal lngRecordId As Long) As Boolean
    '******************************************************************************************************************
    '功能：刷新病历附件列表；
    '参数：lngRecordId：电子病历记录ID
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim objItem As ListItem
    Dim objIcon As StdPicture
    
    Set Me.lvwThis.Icons = Nothing
    Set lvwThis.SmallIcons = Nothing
    lvwThis.ListItems.Clear
    imgThis.ListImages.Clear
    
    gstrSQL = "Select 序号, 文件名, 大小, 创建人, 日期 From 电子病历附件 Where 病历id = [1]"
    If mblnDataMove Then
        gstrSQL = Replace(gstrSQL, "电子病历附件", "H电子病历附件")
    End If
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngRecordId)
    With rs
        Do While Not .EOF
            Set objIcon = GetFileIcon(!文件名, ICON_SMALL, True)
            imgThis.ListImages.Add , , objIcon
            
            Set lvwThis.Icons = imgThis
            Set lvwThis.SmallIcons = imgThis
            
            Set objItem = lvwThis.ListItems.Add(, "_" & !序号, !文件名 & "(" & !大小 & "KB)")
            objItem.Tag = !文件名
            objItem.Icon = imgThis.ListImages.Count
            objItem.SmallIcon = objItem.Icon
            .MoveNext
        Loop
        If lvwThis.ListItems.Count > 0 Then lvwThis.ListItems(1).Selected = True
    End With
    
    Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Function

Private Function GetFileIcon(ByVal strFile As String, ByVal intSize As ICON_SIZE, Optional blnUntrue As Boolean) As StdPicture
    '******************************************************************************************************************
    '功能：返回指定文件的大图标或小图标
    '说明：需要一个PictureBox控件，无边框，AutoRedraw = True
    '参数： strFile，包含后缀的文件名，当文件真实文件时，应该包含完整的路径名
    '       intSize，获取图标的大小
    '       blnUntrue，非真实文件，这时需要创建文件来获得相关信息
    '******************************************************************************************************************
    Dim fInfo As SHFILEINFO
    Dim lngRetu As Long
    
    If blnUntrue Then
        strFile = IIf(Environ$("tmp") <> vbNullString, Environ$("tmp"), Environ$("temp")) & "\" & App.hInstance & CLng(Timer) & strFile
        gobjFSO.CreateTextFile strFile, False
    End If
    If intSize = ICON_LARGE Then
        lngRetu = SHGetFileInfo(strFile, 0, fInfo, Len(fInfo), SHGFI_SHELLICONSIZE Or SHGFI_ICON Or SHGFI_LARGEICON)
    Else
        lngRetu = SHGetFileInfo(strFile, 0, fInfo, Len(fInfo), SHGFI_SHELLICONSIZE Or SHGFI_ICON Or SHGFI_SMALLICON)
    End If
    If blnUntrue Then gobjFSO.DeleteFile strFile, True
    
    picThis.Width = intSize * Screen.TwipsPerPixelX
    picThis.Height = intSize * Screen.TwipsPerPixelY
    picThis.Cls
    If lngRetu <> 0 Then
        DrawIconEx picThis.hDC, 0, 0, fInfo.hIcon, intSize, intSize, 0, 0, DI_NORMAL
        DestroyIcon fInfo.hIcon
    End If
    Set GetFileIcon = Me.picThis.Image
End Function

Private Sub OpenFile()
    '功能：打开播放附件
    Dim strFile As String
    Dim varRetu As Variant, strInfo As String
    
    Screen.MousePointer = vbHourglass
    strFile = IIf(Environ$("tmp") <> vbNullString, Environ$("tmp"), Environ$("temp")) & "\" & lvwThis.SelectedItem.Tag
    If gobjFSO.FileExists(strFile) Then gobjFSO.DeleteFile strFile, True
    If zlBlobRead(8, mlngKey & "," & Mid(lvwThis.SelectedItem.Key, 2), strFile) = "" Then
        MsgBox "文件读取失败，请确认附件的有效性！", vbInformation, gstrSysName:
        Screen.MousePointer = vbDefault: Exit Sub
    End If
    varRetu = ShellExecute(Me.hWnd, "open", strFile, "", "", SW_SHOWNORMAL)
    If varRetu <= 32 Then
        Select Case varRetu
        Case 2: strInfo = "错误的关联"
        Case 29: strInfo = "关联失败"
        Case 30: strInfo = "关联应用程式忙碌中..."
        Case 31: strInfo = "没有关联任何应用程式"
        Case Else: strInfo = "无法识别的错误"
        End Select
        MsgBox "附件打开发生：" & strInfo, vbExclamation, gstrSysName
    End If
    
    Screen.MousePointer = vbDefault
End Sub

'######################################################################################################################

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Open
        Dim frm As New frmEPRView
        frm.ShowMe Me, mlngKey
    Case ID_EDIT_COPY
        If Control.Enabled And Control.Visible Then '快捷键执行时需要判断
            Me.edtThis.Copy
        End If
    End Select
End Sub
 
Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    On Error GoTo errHand

    Select Case Control.ID
    Case conMenu_File_Open
        Control.Enabled = (mlngKey > 0)
    Case ID_EDIT_COPY
        Control.Enabled = edtThis.Selection.EndPos <> edtThis.Selection.StartPos
        Control.Enabled = edtThis.Selection.getType <> cprSTPicture
        Control.Visible = InStr(gstrPrivsEpr, "内容复制") > 0
    End Select

errHand:

End Sub


Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_RichEpr
        Item.Handle = picPane(0).hWnd
    Case conPane_Annex
        Item.Handle = picPane(1).hWnd
    Case conPane_TablEpr
        Item.Handle = mObjTabEprView.zlGetForm.hWnd
    End Select
End Sub

Private Sub edtThis_RequestRightMenu(ViewMode As zlRichEditor.ViewModeEnum, Shift As Integer, X As Single, Y As Single)
    '没有内容复制权限不允许复制
    If InStr(gstrPrivsEpr, "内容复制") = 0 Then Exit Sub
    
    Dim Popup As CommandBar
    Dim Control As CommandBarControl
    
    Set Popup = cbsMain.Add("Popup", xtpBarPopup)
    With Popup.Controls
        Set Control = .Add(xtpControlButton, ID_EDIT_COPY, "复制(&C)")
        Popup.ShowPopup
    End With
End Sub

Private Sub Form_Load()
Dim Pane1 As Pane, pane2 As Pane, pane3 As Pane
    Set mObjTabEprView = New cTableEPR
    mObjTabEprView.InitTableEPR gcnOracle, glngSys, gstrDbOwner
    
    Set Pane1 = dkpMain.CreatePane(conPane_RichEpr, 1200, 200, DockTopOf, Nothing)
    Pane1.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set pane2 = dkpMain.CreatePane(conPane_TablEpr, 1200, 200, DockTopOf, Nothing)
    pane2.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    pane2.Close
    
    Set pane3 = dkpMain.CreatePane(conPane_Annex, 1200, 15, DockBottomOf, Nothing)
    pane3.MinTrackSize.Height = 0: pane3.MaxTrackSize.Height = 360 / Screen.TwipsPerPixelY
    pane3.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    With dkpMain
        .VisualTheme = ThemeOffice2003
        .Options.HideClient = True
        .Options.UseSplitterTracker = True
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = False
    End With

    Call InitCommandBar
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Unload mfrmPrintPreview
    Set mfrmPrintPreview = Nothing
    Unload mObjTabEprView.zlGetForm
    Set mObjTabEprView = Nothing
    Set mfrmMain = Nothing
    Unload mfrmAnnex
    Set mfrmAnnex = Nothing
End Sub

Private Sub lvwThis_DblClick()
    Call OpenFile
End Sub

Private Sub mfrmPrintPreview_PrintEpr(ByVal lngRecordId As Long)
    '
    RaiseEvent PrintEpr(lngRecordId)
    
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 0
        edtThis.Move 0, 0, picPane(Index).Width, picPane(Index).Height
    Case 1
        lblThis.Top = (picPane(Index).Height - Me.lblThis.Height) / 2
        lvwThis.Move lvwThis.Left, 15, picPane(Index).Width - lvwThis.Left, picPane(Index).Height - 30
    End Select
    
End Sub

Private Sub RefreshObject(ByVal lngRecordId As Long, ByVal blnMoved As Boolean)
'刷新界面上的图片,目前只刷新图片，有需要时再调整刷新表格
Dim Pictures As New cEPRPictures, rsTemp As New ADODB.Recordset, lngKey As Long, Tables As New cEPRTables
Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean, blnForce As Boolean

    '读取所有的图片
    gstrSQL = "Select ID, 文件id,开始版, 终止版,父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行,预制提纲ID " & _
        "From 电子病历内容 " & _
        "Where 文件id = [1] And 对象类型 In (3,5) And 对象序号 Is Not Null" '不显示表格中的图片
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

Private Function Check检验行(ByVal lng病历ID As Long, ByRef lng医嘱id As Long, ByRef lng病人ID As Long) As Boolean
'功能：判断指定行是否检验组合行
'说明：指定行为显示行,且类别="E"
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
        
    On Error GoTo ErrH
    
    strSQL = "Select A.相关ID,A.病人ID From 病人医嘱记录 A,病人医嘱报告 B Where Rownum=1 And A.相关ID=B.医嘱ID And A.诊疗类别='C'And B.病历id=[1]"
    If mblnMoved_HL Then
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, lng病历ID)
    If Not rsTmp.EOF Then
        lng医嘱id = rsTmp!相关ID
        lng病人ID = rsTmp!病人ID
        Check检验行 = True
    End If
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function Open_LIS_Report(ByVal frmParent As Object, ByVal lng医嘱id As Long, ByVal lng病人ID As Long, ByVal blnCurrMoved As Boolean, ByVal blnPrint As Boolean) As Boolean
    '调用LiwWork打印带图形的LIS报表
    Dim strChart(0 To 8) As String
    Dim strReportCode As String
    Dim strReportParaNo As String
    Dim bytReportParaMode As Byte
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim intLoop As Integer
    Dim objLisWork As Object
    Dim lng发送号 As Long, lng标本id As Long
                    
    On Error GoTo ErrHandle
    Set objLisWork = CreateObject("zl9LisWork.clsLISImg")
    
    strSQL = "select 发送号 from 病人医嘱发送 a , 病人医嘱记录 b where b.id = a.医嘱id and b.id = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Open_LIS_Report", lng医嘱id)
    If Not rsTmp.EOF Then
        lng发送号 = NVL(rsTmp!发送号, 0)
    End If
    strSQL = "select ID from 检验标本记录 b where b.医嘱id = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "Open_LIS_Report", lng医嘱id)
    If Not rsTmp.EOF Then
        lng标本id = NVL(rsTmp!ID, 0)
    End If
    If lng发送号 = 0 Or lng标本id = 0 Then Exit Function
    
    strSQL = "select id from 检验图像结果 where 标本id = [1] order by ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "LISWork.LIS_Report", lng标本id)
    intLoop = 0
    Do Until rsTmp.EOF
        If Not objLisWork Is Nothing Then
            If objLisWork.Get_Chart2d_File(App.Path, rsTmp("ID")) Then
                strChart(intLoop) = App.Path & "\" & rsTmp("ID") & ".cht"
            End If
        End If
        intLoop = intLoop + 1
        rsTmp.MoveNext
    Loop
    
    If Not objLisWork Is Nothing Then
        If objLisWork.Get_ReportCode(lng医嘱id, lng发送号, strReportCode, strReportParaNo, bytReportParaMode, blnCurrMoved) Then
            Call ReportOpen(gcnOracle, glngSys, strReportCode, frmParent, "NO=" & strReportParaNo, "性质=" & bytReportParaMode, "医嘱ID=" & lng医嘱id, _
                            "病人ID=" & lng病人ID, "标本ID=" & lng标本id, _
                            "图形1=" & strChart(0), "图形2=" & strChart(1), "图形3=" & strChart(2), "图形4=" & strChart(3), _
                            "图形5=" & strChart(4), "图形6=" & strChart(5), "图形7=" & strChart(6), "图形8=" & strChart(7), _
                            "图形9=" & strChart(8), IIf(blnPrint, 2, 1))
        End If
    End If
    '删除图形文件
    For intLoop = 0 To 8
        If strChart(intLoop) <> "" Then
            If Dir(strChart(intLoop)) <> "" Then Kill strChart(intLoop)
        End If
    Next
    
    Open_LIS_Report = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

