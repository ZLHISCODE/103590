VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{FBAFE9A8-8B26-4559-9D12-D70E36A97BE3}#2.1#0"; "zlRichEditor.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEPRView 
   Caption         =   "单病历查阅"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9705
   Icon            =   "frmEPRView.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7800
   ScaleWidth      =   9705
   StartUpPosition =   3  '窗口缺省
   Begin zlRichEditor.Editor edtOrig 
      Height          =   2310
      Left            =   225
      TabIndex        =   0
      Top             =   1305
      Visible         =   0   'False
      Width           =   4020
      _ExtentX        =   7091
      _ExtentY        =   4075
      Title           =   ""
      ShowRuler       =   0   'False
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   7425
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9763
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2716
            MinWidth        =   2716
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1658
            MinWidth        =   1658
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
   Begin zlRichEditor.Editor edtClear 
      Height          =   2625
      Left            =   225
      TabIndex        =   2
      Top             =   3735
      Visible         =   0   'False
      Width           =   4020
      _ExtentX        =   7091
      _ExtentY        =   4630
      Title           =   ""
      ShowRuler       =   0   'False
   End
   Begin MSComDlg.CommonDialog dlgThis 
      Left            =   4770
      Top             =   135
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   45
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      ScaleMode       =   1
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmEPRView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'文件 "File"
Private Const ID_File_SaveCopy = 302    '保存副本(A)...
Private Const ID_File_SaveTxt = 303     '保存为文本(V)...
Private Const ID_FILE_PRINT = 304       '打印(P)...
Private Const ID_FILE_Copy = 305        '复制到剪贴板(C)
Private Const ID_FILE_EXIT = 306        '退出(X)

'视图 "View"
Private Const ID_View_Mode = 311        '显示状态(&S)
Private Const ID_View_Mode_Orig = 312   '原始状态(&O)
Private Const ID_View_Mode_Clear = 313  '清洁状态(&C)
Private Const ID_View_StatusBar = 314   '状态栏(S)
Private Const ID_View_Ref = 315        '诊断参考(D)

'帮助 "Help"
Private Const ID_HELP_CONTENT = 500     '帮助主题
Private Const ID_HELP_CONTACT = 502     '发送反馈
Private Const ID_HELP_ONLINE = 503      '在线医业
Private Const ID_HELP_ABOUT = 504       '关于...

Private mlng记录ID As Long              '记录ID
Private mlngPatiId As Long, mlngPageId As Long '主页ID
Private mlngFileType  As Integer           '病历种类
Private mlngMode As Long                '显示模式:0~Orig; 1~Clear
Private blnPrivacyProtect As Boolean    '是否启用隐私保护

Public Tables As cEPRTables             '表格集合
Public Pictures As cEPRPictures         '图片集合
Public Compends As cEPRCompends         '提纲集合
Public Elements As cEPRElements         '诊治要素集合
Public Signs As cEPRSigns               '签名组集合

Private mblnShowModeless                '以非模态方式显示
Private mstrItems As String             '向诊疗参考传递的串
Private mblnChildMode As Boolean        '是否是嵌入编辑的子窗体
Private mblnCanPrint As Boolean         '是否可以打印
Private mlngAdviceID As Long            '医嘱ID
Private mfrmParent As Object            '调用窗体

'Public Property Get ChildMode() As Boolean
'    ChildMode = mblnChildMode
'End Property
'
'Public Property Let ChildMode(vData As Boolean)
'    mblnChildMode = vData
'    If mblnChildMode Then
'        Me.BorderStyle = 0
'        SetWindowLong Me.hWnd, GWL_STYLE, GetWindowLong(Me.hWnd, GWL_STYLE) Xor WS_BORDER Xor WS_THICKFRAME Xor WS_DLGFRAME
'    Else
'        Me.BorderStyle = 2
'    End If
'End Property

Public Property Get CanPrint() As Boolean
    CanPrint = mblnCanPrint
End Property

Public Property Let CanPrint(vData As Boolean)
    mblnCanPrint = vData
End Property

'################################################################################################################
'## 功能：  显示病历文件查阅窗体
'##
'## 参数：  frmParent       ：父窗体
'##         lng记录ID       ：记录ID
'##         blnPrivacyOn    ：是否启用隐私保护
'##         blnCanPrint     ：是否允许打印
'##         blnChildMode    ：是否是嵌入方式
'################################################################################################################
Public Sub ShowMe(ByRef frmParent As Object, ByVal lng记录ID As Long, _
    Optional blnPrivacyOn As Boolean = False, _
    Optional blnCanPrint As Boolean = True, _
    Optional blnChildMode As Boolean = False, _
    Optional lngAdviceID As Long)
    
    Dim objControl As CommandBarControl
    Set mfrmParent = frmParent
    mlngFileType = 0
    blnPrivacyProtect = blnPrivacyOn
    mblnCanPrint = blnCanPrint
    mlngAdviceID = lngAdviceID
'    Me.ChildMode = blnChildMode
    
    Call InitCommandBars    '工具栏初始化
    
    gobjComLib.zlCommFun.ShowFlash "请稍候..."
    Screen.MousePointer = vbHourglass
    mlng记录ID = lng记录ID      '记录ID
    mlngMode = 1                '清洁模式
    
    Call OpenSignleEPR
    Call ShowEPRFile
    
    gobjComLib.zlCommFun.StopFlash
    Screen.MousePointer = vbDefault
    mblnShowModeless = (frmParent.BorderStyle = vbSizable Or frmParent.BorderStyle = vbBSNone)
    
    Me.Show IIf(mblnShowModeless, vbModeless, vbModal), frmParent
    Exit Sub
LL:
    Unload Me
    MsgBox "无法打开该文件", vbOKOnly + vbInformation, gstrSysName
End Sub

Private Sub ShowEPRFile()
    edtOrig.Visible = (mlngMode = 0)
    edtClear.Visible = (mlngMode = 1)
    Call cbsThis_Resize
End Sub

Public Sub OpenSignleEPR()
    gobjComLib.zlCommFun.ShowFlash "请稍候..."
    Screen.MousePointer = vbHourglass
'    DoEvents
'    LockWindowUpdate Me.hWnd
    '=================================================================================================
    Dim i As Long, strPath As String, strF As String
    Dim rs As New ADODB.Recordset
    Dim Doc As New cEPRDocument, Elements As New cEPRElements
    Dim lngStart As Long, lngLen  As Long
    Dim lng病人ID As Long, lng主页ID As Long, 病历种类 As EPRDocTypeEnum
    Dim lngKey As Long, blnPrivacy As Boolean

    mstrItems = ""
'    If blnPrivacyProtect = True Then
'        blnPrivacy = InStr(gstrPrivsEpr, ";忽略隐私保护;") = 0     '保护隐私项目
'    End If
    On Error GoTo errHand

    gstrSQL = "select 病人ID,主页ID,病历种类 from 电子病历记录 where ID=[1]"
    Set rs = gobjComLib.zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng记录ID)
    If Not rs.EOF Then
        mlngPatiId = Nvl(rs("病人ID"), 0)
        mlngPageId = Nvl(rs("主页ID"), 0)
        mlngFileType = Nvl(rs("病历种类"), 1)
    End If
    rs.Close
    If mlngFileType = cpr诊疗报告 Then
        gstrSQL = "Select b.Id, b.诊疗项目id From 病人医嘱报告 A, 病人医嘱记录 B Where a.病历id = [1] And a.医嘱id = b.Id Order By b.Id"
        Set rs = gobjComLib.zlDatabase.OpenSQLRecord(gstrSQL, "提取诊疗项目ID", mlng记录ID)
        Do Until rs.EOF
            mstrItems = "," & rs!诊疗项目ID
                        rs.MoveNext
        Loop
        If mstrItems <> "" Then mstrItems = Mid(mstrItems, 2)
    End If

    edtOrig.ForceEdit = True
    edtClear.ForceEdit = True
    '保存临时文件
    strPath = IIf(Environ$("tmp") <> vbNullString, Environ$("tmp"), Environ$("temp"))
    strF = strPath & "\" & App.hInstance & CLng(Timer) & ".TMP"
    Doc.InitEPRDoc cprEM_修改, cprET_单病历审核, mlng记录ID, IIf(mlngFileType = 2, 2, 1), lng病人ID, CStr(lng主页ID), 0, UserInfo.部门ID, mlngAdviceID
    Doc.OpenEPRDoc Doc.frmEditor.Editor1         '打开该文件
    '设置替换项目
    If blnPrivacy Then
        '读取所有的要素
        gstrSQL = "Select A.ID,A.对象标记 From 电子病历内容 A, 隐私保护项目 B,诊治所见项目 C " & _
            "Where A.对象类型 = 4 And A.替换域 = 1 And A.文件id = [1] And A.对象序号 > 0 and B.项目id = C.ID And A.要素名称 =C.中文名 And C.替换域 = 1 "
        Set rs = gobjComLib.zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng记录ID)
        If Not rs.EOF Then
            Do While Not rs.EOF
                lngKey = Elements.Add(Nvl(rs("对象标记"), 0))
                Elements("K" & lngKey).GetElementFromDB cprET_单病历编辑, rs("ID"), True, "电子病历内容"
                '替换要素内容
                Elements("K" & lngKey).内容文本 = String(Len(Elements("K" & lngKey).内容文本), "*")
                Elements("K" & lngKey).Refresh Doc.frmEditor.Editor1
                rs.MoveNext
            Loop
        End If
        rs.Close
    End If
    Doc.frmEditor.SaveDocToFile strF, False     '存储非清洁临时文件

    With edtOrig
        .NewDoc
        .ForceEdit = True
        .ViewMode = cprNormal
        .OpenDoc strF

        '设置页眉页脚
        Set .Picture = Doc.frmEditor.Editor1.Picture
        .HeadFileTextRTF = Doc.frmEditor.Editor1.HeadFileTextRTF
        .FootFileTextRTF = Doc.frmEditor.Editor1.FootFileTextRTF

        Call Doc.GetReplacedHeadFootString(edtOrig)
        '设置页面格式
        Doc.EPRFileInfo.SetFormat edtOrig, Doc.EPRFileInfo.格式
        edtOrig.ResetWYSIWYG    '刷新所见即所得（WYSIWYG）显示

        '分页
        .ViewMode = cprNormal
        .AuditMode = True
        .Range(0, 0).Selected
        .ForceEdit = False
        .ReadOnly = True
    End With

    With edtClear
        .NewDoc
        .ForceEdit = True
        .ViewMode = cprNormal
        .OpenDoc strF

        '设置页眉页脚
        Set .Picture = Doc.frmEditor.Editor1.Picture
        .HeadFileTextRTF = Doc.frmEditor.Editor1.HeadFileTextRTF
        .FootFileTextRTF = Doc.frmEditor.Editor1.FootFileTextRTF

        Call Doc.GetReplacedHeadFootString(edtClear)
        '设置页面格式
        Doc.EPRFileInfo.SetFormat edtClear, Doc.EPRFileInfo.格式
        edtClear.ResetWYSIWYG    '刷新所见即所得（WYSIWYG）显示

        '分页
        .SelectAll
        .AuditMode = True
        .AcceptAuditText
        .ViewMode = cprNormal
        .Range(0, 0).Selected
        .ForceEdit = False
        .ReadOnly = True
    End With
    If gobjFSO.FileExists(strF) Then gobjFSO.DeleteFile strF    '删除临时文件

    Doc.frmEditor.Editor1.Modified = False

    Set rs = Nothing

    '=================================================================================================
'    LockWindowUpdate 0
    gobjComLib.zlCommFun.StopFlash
    Screen.MousePointer = vbDefault
    Exit Sub
errHand:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
On Error GoTo ErrH
    Dim strSQL As String
    
    Select Case Control.ID
    Case ID_FILE_Copy
        If Control.Enabled And Control.Visible Then '快捷键执行时需要判断
'            gstrCopyPID = CStr(mlngPatiId)
            If Me.edtOrig.Visible Then
                edtOrig.Copy
            Else
                edtClear.Copy
            End If
        End If
        
    Case ID_FILE_PRINT
        '打印(P)...
        If Me.edtOrig.Visible Then
            If edtOrig.PrintDoc(False, 0, 0, "", 1) = False Then Exit Sub
        Else
            If edtClear.PrintDoc(False, 0, 0, "", 1) = False Then Exit Sub
        End If
        '刷新打印数量
        strSQL = "ZL_影像报告打印_Update(" & mlngAdviceID & ")"
        gobjComLib.zlDatabase.ExecuteProcedure strSQL, "更新打印标记"
        If mfrmParent Is Nothing Then Exit Sub
        Unload Me
    Case ID_FILE_EXIT
        '退出(X)
        Unload Me
    Case ID_View_Mode_Orig
        '原始状态
        mlngMode = 0
        Call ShowEPRFile
    Case ID_View_Mode_Clear
        '最终状态
        mlngMode = 1
        Call ShowEPRFile
    
    Case ID_View_Ref
        Call ShowReference
    Case Else
    End Select
    Exit Sub
ErrH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height / Screen.TwipsPerPixelY
End Sub

Private Sub cbsThis_Resize()
    On Error Resume Next
    Dim Left As Long
    Dim Top As Long
    Dim Right As Long
    Dim Bottom As Long
    
    Me.cbsThis.GetClientRect Left, Top, Right, Bottom
    edtOrig.Width = 0: edtOrig.Height = 0
    edtOrig.Move Left * Screen.TwipsPerPixelX, Top * Screen.TwipsPerPixelY, _
        (Right - Left) * Screen.TwipsPerPixelX, (Bottom - Top) * Screen.TwipsPerPixelY
    edtClear.Width = 0: edtClear.Height = 0
    edtClear.Move Left * Screen.TwipsPerPixelX, Top * Screen.TwipsPerPixelY, _
        (Right - Left) * Screen.TwipsPerPixelX, (Bottom - Top) * Screen.TwipsPerPixelY
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case ID_FILE_Copy
        If Me.edtOrig.Visible Then
            Control.Enabled = (Trim(Me.edtOrig.SelText) <> "" And Me.edtOrig.ViewMode <> cprPaper)
        Else
            Control.Enabled = (Trim(Me.edtClear.SelText) <> "" And Me.edtClear.ViewMode <> cprPaper)
        End If
        Control.Visible = True
    Case ID_FILE_PRINT
        '打印(P)...
        Control.Enabled = mblnCanPrint
    Case ID_FILE_EXIT
        '退出(X)
    Case ID_View_Mode_Orig
        '原始状态
        Control.Checked = (mlngMode = 0)
    Case ID_View_Mode_Clear
        '最终状态
        Control.Checked = (mlngMode = 1)
    Case ID_View_Ref
        Control.Visible = (mlngFileType = cpr诊疗报告)
    Case Else
        '关于...
    End Select
End Sub

Private Sub InitCommandBars()
Dim BarMain As CommandBar
Dim cbp文件 As CommandBarPopup      '文件菜单
Dim cbp视图 As CommandBarPopup      '视图菜单
Dim cbp帮助 As CommandBarPopup      '帮助菜单
    '窗体位置恢复
'    Call RestoreWinState(Me, App.ProductName)
    '## 菜单初始化
    Dim cbpPopup As CommandBarPopup                     '临时对象
    Dim cbpPopupSub As CommandBarPopup                  '临时对象
    Dim objControl As CommandBarControl                 '工具栏控件
    Dim objCustControl As CommandBarControlCustom       '自定义控件
    
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Set Me.cbsThis.Icons = gobjComLib.zlCommFun.GetPubIcons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True         '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsThis.EnableCustomization False
    
    cbsThis.ActiveMenuBar.Title = "菜单栏"
    Set cbp文件 = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "文件(&F)")
    With cbp文件.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, ID_File_SaveCopy, "保存副本(&A)..."): objControl.IconId = 104
        .Add xtpControlButton, ID_File_SaveTxt, "另存为文本(&T)..."
        Set objControl = .Add(xtpControlButton, ID_FILE_Copy, "复制文本(&C)")
        objControl.Visible = False
        
        Set objControl = .Add(xtpControlButton, ID_FILE_PRINT, "打印(&P)..."): objControl.IconId = 103
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_FILE_EXIT, "退出(&X)"): objControl.IconId = 191
        objControl.BeginGroup = True
    End With
    
    Set cbp视图 = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "视图(&V)")
    With cbp视图.CommandBar.Controls
        Set cbpPopup = .Add(xtpControlPopup, 0, "工具栏(&T)")
        cbpPopup.BeginGroup = True
        cbpPopup.CommandBar.Controls.Add xtpControlButton, 59392, "工具栏列表"
        .Add xtpControlButton, ID_View_StatusBar, "状态栏(&S)"
        Set objControl = .Add(xtpControlButton, ID_View_Mode_Orig, "原始状态(&O)"): objControl.BeginGroup = True: objControl.Style = xtpButtonIconAndCaption: objControl.IconId = 817
        Set objControl = .Add(xtpControlButton, ID_View_Mode_Clear, "最终状态(&C)"): objControl.Style = xtpButtonIconAndCaption: objControl.IconId = 4010
        Set objControl = .Add(xtpControlButton, ID_View_Ref, "诊疗参考(&D)"): objControl.Style = xtpButtonIconAndCaption: objControl.IconId = 801
    End With
    
    
    Set BarMain = cbsThis.Add("工具栏", xtpBarTop)
    With BarMain.Controls
        Set objControl = .Add(xtpControlButton, ID_View_Mode_Orig, "原始状态(F5)"): objControl.BeginGroup = True: objControl.Style = xtpButtonIconAndCaption: objControl.IconId = 817
        Set objControl = .Add(xtpControlButton, ID_View_Mode_Clear, "最终状态(F6)"): objControl.Style = xtpButtonIconAndCaption: objControl.IconId = 4010
        Set objControl = .Add(xtpControlButton, ID_View_Ref, "诊疗参考(&D)"): objControl.Style = xtpButtonIconAndCaption: objControl.IconId = 801
'        Set objControl = .Add(xtpControlButton, ID_HELP_CONTENT, "帮助"): objControl.IconId = conMenu_Help_Help
        objControl.BeginGroup = True
        objControl.Style = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, ID_FILE_EXIT, "关闭")
        objControl.BeginGroup = True: objControl.IconId = 191
        objControl.Style = xtpButtonIconAndCaption
    End With
    
    cbsThis.KeyBindings.Add 8, Asc("C"), ID_FILE_Copy
End Sub

Private Sub edtClear_RequestRightMenu(ViewMode As zlRichEditor.ViewModeEnum, Shift As Integer, X As Single, Y As Single)
    '没有内容复制权限不允许复制
'    If InStr(gstrPrivsEpr, "内容复制") = 0 Then Exit Sub

    Dim Popup As CommandBar
    Dim Control As CommandBarControl

    Set Popup = cbsThis.Add("Popup", xtpBarPopup)
    With Popup.Controls
        Set Control = .Add(xtpControlButton, ID_FILE_Copy, "复制(&C)")
        Popup.ShowPopup
    End With
End Sub

Private Sub edtOrig_RequestRightMenu(ViewMode As zlRichEditor.ViewModeEnum, Shift As Integer, X As Single, Y As Single)
    '没有内容复制权限不允许复制
'    If InStr(gstrPrivsEpr, "内容复制") = 0 Then Exit Sub

    Dim Popup As CommandBar
    Dim Control As CommandBarControl

    Set Popup = cbsThis.Add("Popup", xtpBarPopup)
    With Popup.Controls
        Set Control = .Add(xtpControlButton, ID_FILE_Copy, "复制(&C)")
        Popup.ShowPopup
    End With
End Sub

Private Sub Form_Load()
    Set Tables = New cEPRTables
    Set Pictures = New cEPRPictures
    Set Compends = New cEPRCompends
    Set Elements = New cEPRElements
    Set Signs = New cEPRSigns
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    Call SaveWinState(Me, App.ProductName)
    Set Tables = Nothing
    Set Pictures = Nothing
    Set Compends = Nothing
    Set Elements = Nothing
    Set Signs = Nothing
    Set mfrmParent = Nothing
End Sub

Private Sub ShowReference()
On Error GoTo ErrH
    Dim objReference As Object, lngItem As Long
    
    Dim blnDo As Boolean
    Dim lngPatiID As Long
    Dim lngPatiFrom As Long
    Dim lngFlagNo As Long
    Dim lngClinicID As Long
    
    Dim strSQL As String
    Dim rs As Recordset
    Dim strItem As String
    Dim strarr() As String
    Dim strTmp As String
    Dim i As Integer

    gstrSQL = "SELECT 诊疗项目ID,医嘱内容  FROM 病人医嘱记录 where ID=[1]"
    Set rs = gobjComLib.zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngAdviceID)
    
    On Error Resume Next
    strItem = rs!诊疗项目ID & ""
    strTmp = rs!医嘱内容 & ""
    
    
    If strItem = "" Or strTmp = "" Then GoTo NextStep
    strarr = Split(strTmp, ",")
    
    For i = 0 To UBound(strarr)
        If InStr(strarr(i), "(") > 0 Then
            strTmp = Mid(strarr(i), 1, InStr(strarr(i), "(") - 1)
        Else
            strTmp = strarr(i)
        End If
        strItem = strItem & "<SP2>" & strTmp
    Next
    
NextStep:
    If CreatePlugInOK() Then
        If GetClinicHelpInfo(mlngAdviceID, lngPatiID, lngPatiFrom, lngFlagNo, lngClinicID) Then
            On Error Resume Next
            If Not gobjPlugIn Is Nothing Then
                blnDo = gobjPlugIn.ShowClinicHelp(Me.hWnd, 1, lngPatiFrom, lngPatiID, lngFlagNo, lngClinicID, strItem)
                Call zlPlugInErrH(err, "ShowClinicHelp")
            End If
            If err.Number <> 0 Then
                blnDo = False '出错时返回False
                err.Clear
            End If
        End If
        On Error GoTo ErrH
    End If

    If blnDo = True Then Exit Sub
    
    On Error Resume Next
    lngItem = Val(Split(mstrItems, ",")(0))
    Set objReference = CreateObject("zlPublicAdvice.clsPublicAdvice")
    If err.Number <> 0 Then Exit Sub
    Call objReference.InitCommon(gcnOracle, 100)
    Call objReference.ShowClincHelp(IIf(mblnShowModeless, vbModeless, vbModal), Me, lngItem, False, mstrItems, strItem)
    Exit Sub
ErrH:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub
