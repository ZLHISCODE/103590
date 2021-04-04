VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Begin VB.Form frmResult 
   Caption         =   "校对结果"
   ClientHeight    =   8715
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14940
   Icon            =   "frmResult.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   14940
   StartUpPosition =   2  '屏幕中心
   Begin XtremeReportControl.ReportControl rtcResult 
      Height          =   5415
      Left            =   720
      TabIndex        =   1
      Top             =   720
      Width           =   11775
      _Version        =   589884
      _ExtentX        =   20770
      _ExtentY        =   9551
      _StockProps     =   0
      SkipGroupsFocus =   0   'False
   End
   Begin VB.Timer timerPopulate 
      Interval        =   100
      Left            =   6120
      Top             =   120
   End
   Begin VB.CheckBox chkCur 
      Caption         =   "仅显示本次校对"
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   6240
      Width           =   1935
   End
   Begin VB.PictureBox picBox 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   720
      ScaleHeight     =   735
      ScaleWidth      =   11295
      TabIndex        =   2
      Top             =   6600
      Width           =   11295
      Begin VB.PictureBox picHint 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   8055
         TabIndex        =   3
         Top             =   375
         Width           =   8055
      End
      Begin VB.Label lblHint 
         Height          =   375
         Left            =   0
         TabIndex        =   4
         Top             =   120
         Width           =   8055
      End
   End
   Begin MSComctlLib.StatusBar staPane 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   8340
      Width           =   14940
      _ExtentX        =   26353
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   25823
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "发现校对失败的图像，请联系管理员处理。"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   1080
      TabIndex        =   6
      Top             =   240
      Width           =   4560
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager ImgMain 
      Bindings        =   "frmResult.frx":6852
      Left            =   720
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmResult.frx":6866
   End
End
Attribute VB_Name = "frmResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnSelect As Boolean
Private mblnOk As Boolean
Private mblnDo As Boolean
Private mstrDept As String
Private mstrCurValid As String
Private mblnShow As Boolean
Private mlngIndex As Long
Private mstrAdvice As String

Private Enum TColName
    tc医嘱ID = 0
    tc患者信息 = 1
    tc图像UID = 2
    tcIp = 3
    tcFTP路径 = 4
    tc设备 = 5
    tc执行间 = 6
    
    tc采集时间 = 7
    tc校对时间 = 8
    tc校对结果 = 9
    tc本地路径 = 10
End Enum

Public Event OnValid(rsResult As Recordset, ByRef lngResult As emResult, ByRef strFtpDef As String)
Public Event OnUnload()

Public Sub ShowMe(strDept As String, Optional strCurValid As String)
    mblnDo = False
    mstrDept = strDept
    mstrCurValid = strCurValid
    
    Me.Show
    
    If Not mblnDo Then
        Call GetStadyInfo
    End If
End Sub

Private Sub DoCloseRtc()
    Dim i As Long
    
    If rtcResult.Rows.Count <= 0 Then Exit Sub

    For i = 0 To rtcResult.Rows.Count - 1
        If i > rtcResult.Rows.Count - 1 Then Exit Sub
        If rtcResult.Rows(i).GroupRow Then
            rtcResult.Rows(i).Expanded = False
        End If
    Next
    
End Sub

Private Sub cbrMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngResult As emResult
    Dim strAdvice As String
    Dim lngRow As Long
    
    On Error GoTo errHandle
    
    If rtcResult.Rows.Count = 0 And Control.ID <> conMenu_Process_Exit Then Exit Sub
    
    If Not rtcResult.FocusedRow Is Nothing Then
        lngRow = rtcResult.FocusedRow.Index
    End If
    
    Select Case Control.ID
        Case conMenu_Process_ReValidAll
            
            mblnOk = False
            strAdvice = GetAdvice
            Call ReValid(True)
            
            DoWork strAdvice, lngRow
            
            
        Case conMenu_Process_UpDownALL
            mblnOk = False
            
            If MsgBox("是否重新上传所有文件(需本地缓存存在对应文件)？", vbYesNo, "提示") = vbNo Then
                Exit Sub
            End If
            strAdvice = GetAdvice
            Call UpLoad(True)
            
            DoWork strAdvice, lngRow
            
        Case conMenu_Process_ReValid
            If Not DoCheck Then
                MsgBox "请先选择需要操作的患者。", vbInformation, Me.Caption
                Exit Sub
            End If
            mblnOk = False
            strAdvice = GetAdvice
            Call ReValid(False)
            
            DoWork strAdvice, lngRow
            
        Case conMenu_Process_UpDown
            If Not DoCheck Then
                MsgBox "请先选择需要操作的患者。", vbInformation, Me.Caption
                Exit Sub
            End If
            mblnOk = False
            If MsgBox("是否重新上传所选文件(需本地缓存存在对应文件)？", vbYesNo, "提示") = vbNo Then
                Exit Sub
            End If
            strAdvice = GetAdvice
            Call UpLoad(False)
            
            DoWork strAdvice, lngRow

        Case conMenu_Process_Ignore  '忽略校对结果
            If rtcResult.FocusedRow Is Nothing Then
                MsgBox "请选择需要忽略的图像。", vbInformation, "提示"
                Exit Sub
            End If
            
            If rtcResult.FocusedRow.GroupRow Then
                MsgBox "请选择需要忽略的图像。", vbInformation, "提示"
                Exit Sub
            End If
            
            If MsgBox("是否忽略图像【" & rtcResult.FocusedRow.Record(tc图像UID).Value & "】的校对结果？", vbYesNo, "提示") = vbNo Then
                Exit Sub
            End If
            
            strAdvice = GetAdvice
            If IgnoreResult Then
                DoWork strAdvice, lngRow
            End If
        Case conMenu_Process_Exit   '退出
            Unload Me
    End Select
    
    If Control.ID <> conMenu_Process_Exit Then
        If rtcResult.Rows.Count = 0 Then
            lbl.Visible = False
        Else
            lbl.Visible = True
        End If
    End If
    Exit Sub
errHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub DoWork(ByVal strAdvice As String, ByVal lngRow As Long)
    If mblnOk Then
        Call GetStadyInfo
        Call RefreshImagesInfo(strAdvice)
        mblnOk = False
    End If
    
    rtcResult.Populate
    Call DoDold(strAdvice)
    
    If lngRow > 0 Then
        If lngRow <= rtcResult.Rows.Count - 1 Then
            rtcResult.FocusedRow = rtcResult.Rows(lngRow)
        Else
            If rtcResult.Rows.Count > 0 Then
                rtcResult.FocusedRow = rtcResult.Rows(rtcResult.Rows.Count - 1)
            End If
        End If
    End If
End Sub

Private Sub ReValid(blnAll As Boolean)
    Dim rsRecord As Recordset
    Dim lngResult As emResult
    Dim lngRow As Long
    Dim lngDefault As Long
    Dim lngSuceed As Long
    Dim strFtpDef As String
    Dim strFtpConnErr As String
    Dim lngUnValid As Long
    Dim lngCurIndex As Long
    Dim lngCount As Long

    mstrCurValid = ""
    Set rsRecord = GetRecord(blnAll)
    
    If rsRecord Is Nothing Then Exit Sub
    If rsRecord.RecordCount < 1 Then Exit Sub
    
    lngCount = rsRecord.RecordCount
    lngCurIndex = 0
    chkCur.Visible = False
    picBox.Height = 735
    Call Form_Resize
    Do While Not rsRecord.EOF
        lngCurIndex = lngCurIndex + 1
        staPane.Panels(1).Text = "正在校对：" & lngCurIndex & "/" & lngCount & "。已发现" & lngDefault & "个文件校对失败。"
        lblHint.Caption = "正在校对：" & NVL(IIf(Len(NVL(rsRecord("设备名1"))) = 0, NVL(rsRecord("Root2")), NVL(rsRecord("Root1"))) & NVL(rsRecord("URL")))
        lblHint.Refresh
        picHint.Width = picBox.Width / lngCount * lngCurIndex
        picHint.Refresh
        
        If InStr(strFtpConnErr, "[" & IIf(Len(rsRecord!Host1) = 0, rsRecord!Host2, rsRecord!Host1) & "]") = 0 Then
            RaiseEvent OnValid(rsRecord, lngResult, strFtpDef)
            
            If Len(strFtpDef) = 0 Then
'                lngRow = GetIndex(rsRecord!图像Uid)
                If lngResult = etSucceed Then
                    mblnOk = True
                    lngSuceed = lngSuceed + 1
                Else
                    lngDefault = lngDefault + 1
                    If InStr(mstrCurValid, "[" & rsRecord("医嘱ID") & "]") = 0 Then
                        mstrCurValid = mstrCurValid & "[" & rsRecord("医嘱ID") & "]"
                    End If
'                    rtcResult.Rows(lngRow).Record(tc校对结果).Value = GetResult(lngResult)
                End If
            Else
                strFtpConnErr = strFtpConnErr & "[" & strFtpDef & "]"
                lngUnValid = lngUnValid + 1
            End If
        Else
            lngUnValid = lngUnValid + 1
        End If
        rsRecord.MoveNext
    Loop
    
    lblHint.Caption = ""
    picBox.Height = 0
    chkCur.Visible = True
    Call Form_Resize
    picHint.Width = 0
    staPane.Panels(1).Text = "校对完成。本次共" & lngCount & "个文件，有" & lngDefault & "个文件校对失败" & IIf(lngUnValid > 0, "，" & lngUnValid & "个未校对(FTP连接失败)。", "。")
End Sub

Private Sub chkCur_Click()
    Dim strAdvice As String
    
    On Error GoTo errHandle
    
    strAdvice = GetAdvice
    
    Call GetStadyInfo
    Call DoDold(strAdvice)
    strAdvice = GetAdvice
    Call RefreshImagesInfo(strAdvice)
    rtcResult.Populate
    Call DoDold(strAdvice)
    Exit Sub
errHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub Form_Load()
    On Error GoTo errHandle
    
    picHint.BackColor = &H8000000D
    picHint.Width = 0
    picHint.Left = -15
    picBox.Height = 0
    Call InitCommandBars
'    Call InitGrid
    Call InitRtcResult
''    Call InitData
    Call GetStadyInfo
    Call DoCloseRtc
    
    If rtcResult.Rows.Count = 0 Then
        lbl.Visible = False
    Else
        lbl.Visible = True
    End If
    
    mblnDo = True
    Exit Sub
errHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub InitCommandBars()
    '功能创建工具条
    Dim cbrControl As CommandBarControl
    Dim objControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrMain.VisualTheme = xtpThemeOffice2003
    Set Me.cbrMain.Icons = ImgMain.Icons
    
    With Me.cbrMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    Me.cbrMain.EnableCustomization False
    Me.cbrMain.ActiveMenuBar.Visible = False
    
    '图像操作工具栏定义
    Set cbrToolBar = Me.cbrMain.Add("图像操作栏", xtpBarTop)
'    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = True '文本显示在图标下方
    cbrToolBar.Closeable = False
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_ReValid, "重新校对"): cbrControl.ToolTipText = "重新校对所选失败图像"
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_UpDown, "重新上传"): cbrControl.ToolTipText = "从本地重新上传所选失败图像到FTP"
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_ReValidAll, "校对所有"): cbrControl.ToolTipText = "重新校对所有失败图像"
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_UpDownALL, "上传所有"): cbrControl.ToolTipText = "从本地重新上传所有失败图像到FTP"
        If IsDBA Then
            Set cbrControl = .Add(xtpControlButton, conMenu_Process_Ignore, "忽略结果"): cbrControl.ToolTipText = "忽略校对结果，对失败的图像不再提示"
            cbrControl.BeginGroup = True
        End If
        Set cbrControl = .Add(xtpControlButton, conMenu_Process_Exit, "退出"): cbrControl.ToolTipText = "退出"
        cbrControl.BeginGroup = True
    End With
    For Each cbrControl In cbrToolBar.Controls
         cbrControl.Style = xtpButtonIconAndCaption
         cbrControl.Category = "Main" '设置成主界面菜单
    Next
    cbrToolBar.Position = xtpBarTop
End Sub

Private Function GetResult(lngResult As emResult) As String
    Select Case lngResult
        Case etUndetected    '未校对或未校对出错
            GetResult = "未校对"
        Case etFileMiss          '文件缺失
            GetResult = "文件缺失"
        Case etFileNull         '文件大小为0
            GetResult = "文件大小为0"
        Case etReadError        '读取异常
            GetResult = "读取异常"
        Case etRoadError        '路径错误
            GetResult = "路径错误"
        Case etSucceed       '校对成功
            GetResult = "校对成功"
    End Select
End Function

Private Sub Form_Resize()
On Error Resume Next
    Dim lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long
    
    cbrMain.RecalcLayout
    cbrMain.GetClientRect lngLeft, lngTop, lngRight, lngBottom
    
    lbl.Left = lngLeft + (lngRight - lbl.Width) / 2
    lbl.Top = lngTop + 100
    rtcResult.Left = 0
    rtcResult.Top = lngTop + lbl.Height + 200
    rtcResult.Width = Me.ScaleWidth
    rtcResult.Height = Me.ScaleHeight - rtcResult.Top - staPane.Height - picBox.Height - chkCur.Height - 120
    
    chkCur.Left = rtcResult.Left + 120
    chkCur.Top = rtcResult.Top + rtcResult.Height + 120
    
    picBox.Left = 0
    picBox.Top = chkCur.Top + chkCur.Height + 120
    picBox.Width = rtcResult.Width
End Sub

Private Sub GetStadyInfo()
'获取校对失败的检查信息
On Error GoTo errH
    Dim strSql As String
    Dim rsRecord As New ADODB.Recordset
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim blnContinue As Boolean
    
    rtcResult.Records.DeleteAll
    rtcResult.Populate
    
    strSql = "Select a.医嘱id, a.影像类别, a.姓名, 性别, a.年龄, a.检查uid, a.检查号, b.名称" & vbNewLine & _
            "From 影像检查记录 a, 部门表 b" & vbNewLine & _
            "Where a.执行科室id = b.Id And 校对状态 = [1] " & IIf(Len(mstrDept) > 0, " and b.名称 in " & mstrDept, "")
        
    Set rsRecord = gobjComlib.zlDatabase.OpenSQLRecord(strSql, "获取校对失败的检查信息", 2)
    
    If rsRecord.RecordCount <= 0 Then Exit Sub

    Do While Not rsRecord.EOF
        blnContinue = False
        If chkCur.Value = 1 Then
            If InStr(mstrCurValid, "[" & NVL(rsRecord!医嘱ID) & "]") > 0 Then
                blnContinue = True
            End If
        Else
            blnContinue = True
        End If
        If blnContinue Then
            Set objRecord = Me.rtcResult.Records.Add()
            
            Set objItem = objRecord.AddItem(NVL(rsRecord!医嘱ID))
            Set objItem = objRecord.AddItem("患者姓名:" & NVL(rsRecord!姓名) & "   性别:" & NVL(rsRecord!性别) & "   年龄:" & NVL(rsRecord!年龄) & "   病人科室:" & NVL(rsRecord!名称) & "   检查号:" & NVL(rsRecord!检查号) & "【" & NVL(rsRecord!影像类别) & "-" & Val(rsRecord!医嘱ID) & "】")
            Set objItem = objRecord.AddItem("0")
            Set objItem = objRecord.AddItem("0")
            Set objItem = objRecord.AddItem("0")
            Set objItem = objRecord.AddItem("0")
            Set objItem = objRecord.AddItem("0")
            Set objItem = objRecord.AddItem("0")
            Set objItem = objRecord.AddItem("0")
            Set objItem = objRecord.AddItem("0")
        End If
        
        rsRecord.MoveNext
    Loop
    
    rtcResult.Populate
    Exit Sub
errH:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub


Private Function GetWhere(blnAll As Boolean) As String
    Dim i As Long
    On Error GoTo errH
    
    GetWhere = ""
    With rtcResult
        For i = .Rows.Count - 1 To 0 Step -1
            If i > rtcResult.Rows.Count - 1 Then Exit For
            
            
            If IIf(blnAll, True, .Rows(i).Selected = True) And .Rows(i).GroupRow Then
                
                .Rows(i).Expanded = True
            End If
        Next
        For i = 0 To .Rows.Count - 1
'            If .Rows(i).Selected = True Then
'                If Not .Rows(i).GroupRow Then
'
'                    GetWhere = GetWhere & IIf(Len(GetWhere) = 0, "'", ",'") & .Rows(i).Record(tc图像UID).Value & "'"
'                End If
'            Else
'                If Not .Rows(i).GroupRow Then
'                    If .Rows(i).ParentRow.Selected Then
'
'                        GetWhere = GetWhere & IIf(Len(GetWhere) = 0, "'", ",'") & .Rows(i).Record(tc图像UID).Value & "'"
'                    End If
'                End If
'            End If
            If Not .Rows(i).GroupRow Then
'                If InStr(strAdvice, .Rows(i).Record(tc医嘱ID).Value) > 0 Then
                If IIf(blnAll, True, .Rows(i).ParentRow.Selected) Then
                    GetWhere = GetWhere & IIf(Len(GetWhere) = 0, "'", ",'") & .Rows(i).Record(tc医嘱ID).Value & "'"
                End If
'                End If
            End If
        Next
        
        If Len(GetWhere) > 0 Then GetWhere = "c.医嘱ID in (" & GetWhere & ")"
    End With
    Exit Function
errH:
    MsgBox Err.Description, vbCritical, Me.Caption
End Function

Private Function GetRecord(blnAll As Boolean, Optional strImageUid As String) As Recordset
On Error GoTo errH
    Dim strWhere As String
    Dim strSql As String
    
    If Len(strImageUid) > 0 Then
        strWhere = "a.图像Uid = '" & strImageUid & "'"
    Else
        strWhere = GetWhere(blnAll)
    End If
    
    If Len(strWhere) = 0 Then
        Exit Function
    End If
    If Len(strWhere) > 0 Then strWhere = " and " & strWhere
    
    strSql = "Select Rownum As 顺序号,c.医嘱ID,c.姓名, c.性别, c.年龄,c.影像类别,c.检查号, a.图像号, a.采集时间,c.接收日期, d.Ftp用户名 As User1, d.Ftp密码 As Pwd1, d.Ip地址 As Host1," & vbNewLine & _
                "       '/' || d.Ftp目录 || '/' As Root1, d.共享目录 As 共享目录1, d.共享目录用户名 As 共享目录用户名1, d.共享目录密码 As 共享目录密码1," & vbNewLine & _
                "       Decode(c.接收日期, Null, '', To_Char(c.接收日期, 'YYYYMMDD') || '/') || c.检查uid || '/' || a.图像uid As Url, d.设备号 As 设备号1," & vbNewLine & _
                "       d.设备名 As 设备名1, e.Ftp用户名 As User2, e.Ftp密码 As Pwd2, e.Ip地址 As Host2, '/' || e.Ftp目录 || '/' As Root2," & vbNewLine & _
                "       e.共享目录 As 共享目录2, e.共享目录用户名 As 共享目录用户名2, e.共享目录密码 As 共享目录密码2, e.设备号 As 设备号2, e.设备名 As 设备名2, a.图像uid, c.检查uid,f.名称,g.执行间," & vbNewLine & _
                "       b.序列uid, a.动态图, a.编码名称, a.录制长度, c.校对日期, a.校对结果" & vbNewLine & _
                "From 影像检查图象 a, 影像检查序列 b, 影像检查记录 c, 影像设备目录 d, 影像设备目录 e ,部门表 f,病人医嘱发送 g" & vbNewLine & _
                "Where a.序列uid = b.序列uid And b.检查uid = c.检查uid And c.位置一 = d.设备号(+) And c.位置二 = e.设备号(+) and c.校对状态 = 2 and c.执行科室id = f.id and c.医嘱ID = g.医嘱ID and nvl(a.动态图,0) = 0 and Nvl(a.校对结果,0) > 0 and Nvl(a.校对结果,0) < 5" & strWhere & vbNewLine & _
                "Order by a.图像UID"
    
    Set GetRecord = gobjComlib.zlDatabase.OpenSQLRecord(strSql, "获取校对失败图像")
    Exit Function
errH:
    MsgBox Err.Description, vbCritical, Me.Caption
End Function

Private Function UpLoad(blnAll As Boolean) As Boolean
    Dim Inet1 As New clsFtp
    Dim Inet2 As New clsFtp
    Dim objFile As New Scripting.FileSystemObject
    Dim strTmpFile As String
    Dim strCachePath As String
    Dim strImgInstanceUid As String
    Dim rsRecord As Recordset
    Dim blnResult As Boolean
    Dim lngCount As Long
    Dim lngRow As Long
    Dim strSql As String
    Dim strFtpDef As String
    Dim strFtpConnErr As String
    Dim lngDefult As Long
    Dim lngCurIndex As Long
    Dim dcmImage As DicomImage
    
    Set rsRecord = GetRecord(blnAll)
    
    If rsRecord Is Nothing Then Exit Function
    If rsRecord.RecordCount < 1 Then Exit Function
    
    strFtpConnErr = ""
    lngCurIndex = 0

    picBox.Height = 735
    chkCur.Visible = False
    Call Form_Resize
    
    Do While Not rsRecord.EOF
        lngCurIndex = lngCurIndex + 1
        
        staPane.Panels(1).Text = "正在上传：" & lngCurIndex & "/" & lngCount & "。已发现" & lngDefult & "个文件上传失败。"
        lblHint.Caption = "正在上传：" & NVL(IIf(Len(NVL(rsRecord("设备名1"))) = 0, NVL(rsRecord("Root2")), NVL(rsRecord("Root1"))) & NVL(rsRecord("URL")))
        lblHint.Refresh
        picHint.Width = picBox.Width / rsRecord.RecordCount * lngCurIndex
        picHint.Refresh
            
        If InStr(strFtpConnErr, "[" & IIf(Len(rsRecord!Host1) = 0, rsRecord!Host2, rsRecord!Host1) & "]") = 0 Then
            strFtpDef = ""
            blnResult = False
            strCachePath = GetCacheDir
            strImgInstanceUid = Trim(NVL(rsRecord!图像UID))
    
            strTmpFile = strCachePath & NVL(rsRecord("URL"))
        
        
            strTmpFile = Replace(Trim(strTmpFile), "/", "\")
            
            If Dir(strTmpFile) <> vbNullString Then
                Set dcmImage = ReadViewImage(strTmpFile)
                
                If Not dcmImage Is Nothing Then
                    '建立FTP连接
                    If NVL(rsRecord!设备号1) <> vbNullString And Inet1.hConnection = 0 Then
                        If Inet1.FuncFtpConnect(NVL(rsRecord!Host1), NVL(rsRecord!User1), NVL(rsRecord!Pwd1)) = 0 Then
                            If NVL(rsRecord!设备号2) <> vbNullString Then
                                If Inet2.FuncFtpConnect(NVL(rsRecord!Host2), NVL(rsRecord!User2), NVL(rsRecord!Pwd2)) = 0 Then
    
                                    strFtpDef = rsRecord("Host2")
                                End If
                            Else
                                strFtpDef = rsRecord("Host1")
                            End If
                        End If
                    End If
                    
                    If Len(strFtpDef) = 0 Then
                        If Inet1.FuncUploadFile(objFile.GetParentFolderName(NVL(rsRecord!Root1) & rsRecord!Url), strTmpFile, objFile.GetFileName(strTmpFile)) <> 0 Then
                            If NVL(rsRecord!设备号2) <> vbNullString Then
                                If Inet2.hConnection = 0 Then Inet2.FuncFtpConnect NVL(rsRecord!Host2), NVL(rsRecord!User2), NVL(rsRecord!Pwd2)
                                If Inet1.FuncUploadFile(objFile.GetParentFolderName(NVL(rsRecord!Host2) & rsRecord!Url), strTmpFile, objFile.GetFileName(strTmpFile)) = 0 Then
                                    blnResult = True
                                End If
                            End If
                        Else
                            blnResult = True
                        End If
                    End If
                End If
                If Len(strFtpDef) = 0 Then
                    strFtpConnErr = strFtpConnErr & "[" & strFtpDef & "]"
                End If
            End If
            If blnResult Then
                ' 记录到数据库
                strSql = "zl_影像检查图象_校对('" & rsRecord("医嘱ID") & "','" & rsRecord("图像UID") & "',to_date('" & gobjComlib.zlDatabase.Currentdate & "','yyyy-mm-dd hh24:mi:ss'),5)"
                Call gobjComlib.zlDatabase.ExecuteProcedure(strSql, "保存校对结果")
                
                lngCount = lngCount + 1
                mblnOk = True
                
            Else
                lngDefult = lngDefult + 1
            End If
        Else
            lngDefult = lngDefult + 1
        End If
        rsRecord.MoveNext
    Loop
    staPane.Panels(1).Text = "上传完成。本次共" & rsRecord.RecordCount & "个文件。" & lngCount & "个上传成功，" & lngDefult & "个上传失败。"
    picHint.Width = 0
    picBox.Height = 0
    chkCur.Visible = True
    Call Form_Resize
    lblHint.Caption = ""
End Function


Private Function ServeValid() As Boolean
'检查zlPacsServeCenter服务失败目录
    Dim strServePath As String
    Dim objFile As New Scripting.FileSystemObject
    Dim strFile As String
    Dim strFileContent As String
    Dim blnTag As Boolean
    
    strServePath = GetAppRoot & "\Pacs\PacsServerCenter\FileCache\Abandon\"
        
    strFile = Dir(strServePath & "*.XML")
    Do While strFile <> ""   ' 开始循环。
        strFileContent = AnalysisXML(strServePath & strFile)
        blnTag = UCase(Mid(strServePath & strFile, 1, 2)) = "U_"
        Call DoServerAbandon(strFileContent, blnTag)
        strFile = Dir   ' 查找下一个目录。
    Loop
    
End Function

Private Sub DoServerAbandon(strTemp As String, blnTag As Boolean)
'将后台传输服务失败的图像做颜色标记
    Dim strFileName As String
    
    
    If rtcResult.Rows.Count <= 0 Then Exit Sub
    
    strFileName = GetFileInfo("文件名称", strTemp)
    
    If InStr(UCase(strFileName), ".AVI") > 0 Or InStr(UCase(strFileName), ".WAV") > 0 Or InStr(UCase(strFileName), ".JPG") > 0 Then
        Exit Sub
    Else
        
        If Not DoColor(strFileName, blnTag) Then
'            AddNew GetRecord(strFileName), etFileMiss, Replace(GetFileInfo("本地目录", strTemp) & "\" & strFileName, "\\", "\")
            Call DoColor(strFileName, blnTag)
        End If
    End If
End Sub

Private Function DoColor(strFileName As String, blnTag As Boolean) As Boolean
    Dim i As Long
    
    For i = 0 To rtcResult.Rows.Count - 1
        If Not rtcResult.Rows(i).GroupRow Then
            If rtcResult.Rows(i).Record(tc图像UID).Value = strFileName Then
                ChangeCorlor i, blnTag
                DoColor = True
                Exit For
            End If
        End If
    Next
End Function

Private Sub ChangeCorlor(lngRow As Long, blnTag As Boolean)
    Dim i As Long
    
    For i = 0 To 9
        rtcResult.Rows(lngRow).Record(i).BackColor = IIf(blnTag, vbRed, vbGreen)
    Next
End Sub

Private Function AnalysisXML(ByVal FilePath As String) As String
'读取XML文件内容
    Dim strContent As String
    
    strContent = OpenFile(FilePath)

    AnalysisXML = strContent
End Function

Private Function GetFileInfo(strItem As String, strContent As String) As String
    Dim strInfor As String
    
    strInfor = Mid(strContent, InStr(strContent, "<" & strItem & ">") + 6, InStr(strContent, "</" & strItem & ">") - InStr(strContent, "<" & strItem & ">") - 6)
    
    GetFileInfo = strInfor
End Function

Private Function OpenFile(ByVal strFile As String) As String
    Dim strFileLine As String
    Dim curByte() As Byte
    Dim lngSzie As Long
    Dim lngFileNum As Long
    
    lngFileNum = FreeFile
    
    lngSzie = FileLen(strFile)
    
    ReDim curByte(lngSzie) As Byte
    Open strFile For Binary As #lngFileNum
    
    Get #lngFileNum, , curByte()

    Close #lngFileNum
    
    OpenFile = Unicode8Decode(curByte)
End Function

Public Function ByteArrayToString(hex() As Byte, size As Long) As String
    Dim i As Long, c As Byte, str As String
    str = ""
    For i = 0 To size - 1
            str = str & ChrB(hex(i))
    Next i
    ByteArrayToString = str
End Function

Private Sub InitRtcResult()
    Dim objCol As ReportColumn
        
    '初始化排队队列显示字段
    Call Me.rtcResult.Columns.DeleteAll
    With Me.rtcResult.Columns
        rtcResult.AutoColumnSizing = True
        rtcResult.AllowColumnRemove = False
        rtcResult.ShowItemsInGroups = False
        rtcResult.SkipGroupsFocus = True
        rtcResult.MultipleSelection = True
        rtcResult.AutoColumnSizing = False
        
        With rtcResult.PaintManager
            .ColumnStyle = xtpColumnShaded
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "将列标题拖动到此,可按该列分组..."
            .NoItemsText = "没有可显示的项目..."
            .VerticalGridStyle = xtpGridSolid
        End With
        
        Set objCol = .Add(tc医嘱ID, "医嘱ID", 0, False)
        objCol.Visible = False
        
        Set objCol = .Add(tc患者信息, "患者信息", 0, True)
        objCol.Visible = False
        objCol.Sortable = False
        objCol.Editable = False
        
        Set objCol = .Add(tc图像UID, "文件名", 180, True)
        objCol.Sortable = False
        objCol.Editable = False
        
        Set objCol = .Add(tc设备, "设备", 100, True)
        objCol.Sortable = False
        objCol.Editable = False
        objCol.Groupable = False
        
        Set objCol = .Add(tcIp, "IP地址", 100, True)
        objCol.Sortable = False
        objCol.Editable = False
        objCol.Groupable = False
        
        Set objCol = .Add(tcFTP路径, "FTP路径", 200, True)
        objCol.Sortable = False
        objCol.Editable = False
        objCol.Groupable = False
        
        
        Set objCol = .Add(tc执行间, "检查房间", 100, True)
        objCol.Sortable = False
        objCol.Editable = False
        objCol.Groupable = False
        
        Set objCol = .Add(tc采集时间, "采集时间", 100, True)
        objCol.Sortable = False
        objCol.Editable = False
        objCol.Groupable = False
        
        Set objCol = .Add(tc校对时间, "校对时间", 100, True)
        objCol.Sortable = False
        objCol.Editable = False
        objCol.Groupable = False
        
        Set objCol = .Add(tc校对结果, "校对结果", 80, True)
        objCol.Sortable = False
        objCol.Editable = False
        objCol.Groupable = False
    End With
    
    With rtcResult
        .GroupsOrder.DeleteAll
        .GroupsOrder.Add .Columns(tc患者信息)
        .GroupsOrder.Column(0).Caption = ""
        
        .SortOrder.DeleteAll
        .SortOrder.Add .Columns(tc图像UID)
        .SortOrder(0).SortAscending = True
    End With
    
End Sub


Private Function GetIndex(strValue As String) As Long
'跟据图像UID获取行数
    Dim i As Long
    
    GetIndex = 0
    For i = 1 To rtcResult.Rows.Count - 1
        If Not rtcResult.Rows(i).GroupRow Then
            If rtcResult.Rows(i).Record(tc图像UID).Value = strValue Then
                GetIndex = i
                Exit Function
            End If
        End If
    Next
End Function


Public Sub DeleteRow(ByVal lngRowIndex As Long)
'删除队列记录数据
    Dim lngRecordIndex As Long
    
    lngRecordIndex = rtcResult.Rows(lngRowIndex).Record.Index
    rtcResult.Rows(lngRowIndex).Selected = False
    
    Call rtcResult.Records.RemoveAt(lngRecordIndex)
    Call rtcResult.Populate
    
    If rtcResult.Rows.Count > lngRowIndex Then
        rtcResult.Rows(lngRowIndex).Selected = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    RaiseEvent OnUnload
End Sub

Private Sub picBox_Resize()
    On Error Resume Next
    
    lblHint.Width = picBox.Width
End Sub


Private Sub rtcResult_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
On Error GoTo errH
    Dim lngRow As Long
    Dim lngAdvice As Long
    Dim strAdvice As String
    
    If mblnSelect And mblnShow Then
        lngRow = mlngIndex
        If rtcResult.Rows(lngRow).GroupRow Then
            If rtcResult.Rows(lngRow).Childs(0).Record(tc图像UID).Value = "0" Then
                lngAdvice = rtcResult.Rows(lngRow).Childs(0).Record(tc医嘱ID).Value
                strAdvice = GetAdvice
                Call GetImageInfo(lngAdvice)
                
                mstrAdvice = strAdvice
                
                rtcResult.FocusedRow = rtcResult.Rows(lngRow)
            End If
        End If
        mblnShow = False
        mblnSelect = False
        mlngIndex = 0
    End If
    Exit Sub
errH:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub rtcResult_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngRow As Long
    Dim lngAdvice As Long

    On Error GoTo errHandle
    If (KeyCode = 13 Or KeyCode = 0 Or KeyCode = 39 Or KeyCode = 37) And mblnSelect Then
        If rtcResult.FocusedRow Is Nothing Then Exit Sub
        lngRow = rtcResult.FocusedRow.Index

        If rtcResult.Rows(lngRow).GroupRow Then
'            If rtcResult.Rows(lngRow).Expanded = True Then
                If rtcResult.Rows(lngRow).Childs(0).Record(tc图像UID).Value = "0" Then
                    mblnShow = True
                    mlngIndex = lngRow
                End If
'            End If
        End If
    End If
    Exit Sub
errHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub rtcResult_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim lngRow As Long
    Dim lngAdvice As Long
    
    On Error GoTo errHandle
    If Button = 1 And Shift = 0 Then

        If rtcResult.FocusedRow Is Nothing Then Exit Sub
        lngRow = rtcResult.FocusedRow.Index

        If rtcResult.Rows(lngRow).GroupRow Then
'            If rtcResult.Rows(lngRow).Expanded = True Then
                If rtcResult.Rows(lngRow).Childs(0).Record(tc图像UID).Value = "0" Then
                    mblnShow = True
                    mlngIndex = lngRow
                End If
'            End If
        End If
    End If
    
    If Button = 2 And Not rtcResult.FocusedRow Is Nothing Then
        Dim objPopup As CommandBar
        Dim objControl As CommandBarControl

        Set cbrMain.Icons = ImgMain.Icons
        Set objPopup = cbrMain.Add("右键菜单", xtpBarPopup)
        With objPopup.Controls
            Set objControl = .Add(xtpControlButton, conMenu_Process_ReValid, "重新校对")
            Set objControl = .Add(xtpControlButton, conMenu_Process_UpDown, "重新上传")
            Set objControl = .Add(xtpControlButton, conMenu_Process_ReValidAll, "校对所有")
            Set objControl = .Add(xtpControlButton, conMenu_Process_UpDownALL, "上传所有")
            
            If IsDBA Then
                Set objControl = .Add(xtpControlButton, conMenu_Process_Ignore, "忽略结果")
            End If
        End With
        objPopup.ShowPopup
    End If
    Exit Sub
errHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
    mblnSelect = False
End Sub

Private Sub rtcResult_SelectionChanged()
    mblnSelect = True
End Sub


Private Sub GetImageInfo(Optional lngAdvice As Long, Optional blnOpen As Boolean)
'获取校对失败的图像信息
    Dim strSql As String
    Dim rsRecord As New ADODB.Recordset
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim lngRow As Long
    Dim lngCur As Long
    Dim i As Long

    strSql = "Select Rownum As 顺序号,c.医嘱ID,c.姓名, c.性别, c.年龄,c.影像类别,c.检查号, a.图像号, a.采集时间,c.接收日期, d.Ftp用户名 As User1, d.Ftp密码 As Pwd1, d.Ip地址 As Host1," & vbNewLine & _
                "       '/' || d.Ftp目录 || '/' As Root1, d.共享目录 As 共享目录1, d.共享目录用户名 As 共享目录用户名1, d.共享目录密码 As 共享目录密码1," & vbNewLine & _
                "       Decode(c.接收日期, Null, '', To_Char(c.接收日期, 'YYYYMMDD') || '/') || c.检查uid || '/' || a.图像uid As Url, d.设备号 As 设备号1," & vbNewLine & _
                "       d.设备名 As 设备名1, e.Ftp用户名 As User2, e.Ftp密码 As Pwd2, e.Ip地址 As Host2, '/' || e.Ftp目录 || '/' As Root2," & vbNewLine & _
                "       e.共享目录 As 共享目录2, e.共享目录用户名 As 共享目录用户名2, e.共享目录密码 As 共享目录密码2, e.设备号 As 设备号2, e.设备名 As 设备名2, a.图像uid, c.检查uid,f.名称,g.执行间," & vbNewLine & _
                "       b.序列uid, a.动态图, a.编码名称, a.录制长度, c.校对日期, a.校对结果" & vbNewLine & _
                "From 影像检查图象 a, 影像检查序列 b, 影像检查记录 c, 影像设备目录 d, 影像设备目录 e ,部门表 f,病人医嘱发送 g" & vbNewLine & _
                "Where a.序列uid = b.序列uid And b.检查uid = c.检查uid And c.位置一 = d.设备号(+) And c.位置二 = e.设备号(+) and c.校对状态 = 2 and c.执行科室id = f.id and c.医嘱ID = g.医嘱ID and nvl(a.动态图,0) = 0  and (a.校对结果 > 0 and a.校对结果 < 5) " & IIf(lngAdvice <> 0, "and c.医嘱ID = [1]", "") & vbNewLine & _
                "Order by a.图像UID"

    If lngAdvice = 0 Then
        Set rsRecord = gobjComlib.zlDatabase.OpenSQLRecord(strSql, "获取校对失败的检查信息")
    Else
        Set rsRecord = gobjComlib.zlDatabase.OpenSQLRecord(strSql, "获取校对失败的检查信息", lngAdvice)
    End If

    If rsRecord.RecordCount <= 0 Then Exit Sub

    Do While Not rsRecord.EOF
        Set objRecord = Me.rtcResult.Records.Add()

        Set objItem = objRecord.AddItem(NVL(rsRecord!医嘱ID))
        Set objItem = objRecord.AddItem("患者姓名:" & NVL(rsRecord!姓名) & "   性别:" & NVL(rsRecord!性别) & "   年龄:" & NVL(rsRecord!年龄) & "   病人科室:" & NVL(rsRecord!名称) & "   检查号:" & NVL(rsRecord!检查号) & "【" & NVL(rsRecord!影像类别) & "-" & Val(rsRecord!医嘱ID) & "】")
        Set objItem = objRecord.AddItem(NVL(rsRecord!图像UID))
        Set objItem = objRecord.AddItem(IIf(Len(NVL(rsRecord!Host1)) <> 0, NVL(rsRecord!Host1), NVL(rsRecord!Host2)))
        Set objItem = objRecord.AddItem(NVL(IIf(Len(NVL(rsRecord("设备名1"))) = 0, NVL(rsRecord("Root2")), NVL(rsRecord("Root1"))) & NVL(rsRecord("URL"))))
        Set objItem = objRecord.AddItem(NVL(IIf(Len(NVL(rsRecord("设备名1"))) = 0, NVL(rsRecord("设备名2")), NVL(rsRecord("设备名1")))))
        Set objItem = objRecord.AddItem(NVL(rsRecord!执行间))
        Set objItem = objRecord.AddItem(NVL(rsRecord("采集时间")))
        Set objItem = objRecord.AddItem(NVL(rsRecord("校对日期")))
        Set objItem = objRecord.AddItem(GetResult(NVL(rsRecord("校对结果"))))

        rsRecord.MoveNext
    Loop

'    rtcResult.Populate

'    lngCur = GetIndexRow(lngAdvice)
    
    Call DeleteBlankRow(lngAdvice)
    
    rtcResult.SortOrder(0).SortAscending = True
'    Call DoCloseRtc
'
    
'    lngCur = GetIndexRow(lngAdvice)
'    Call DoCloseRtc
'
'    rtcResult.Rows(lngCur).Record(tc图像UID).Record.Tag = 1
End Sub

'Private Sub DeleteBlankRow(lngRow As Long)
'    Dim i As Long
'
'    For i = lngRow To rtcResult.Rows.Count - 1
'        If rtcResult.Rows(i).GroupRow Then Exit Sub
'
'        If rtcResult.Rows(i).Record(tc图像UID).Value = "0" Then
'            rtcResult.Records.RemoveAt rtcResult.Rows(i).Index
'            Exit Sub
'        End If
'
'    Next
'End Sub

Private Sub DeleteBlankRow(lngAdvice As Long)
    Dim i As Long

    For i = 0 To rtcResult.Records.Count - 1

        If rtcResult.Records.Record(i).Item(tc图像UID).Value = "0" And rtcResult.Records.Record(i).Item(tc医嘱ID).Value = lngAdvice Then
            rtcResult.Records.RemoveAt i
            Exit Sub
        End If
    Next
    
End Sub

Private Function GetIndexRow(lngAdvice As Long) As Long
    Dim i As Long
    
    For i = 0 To rtcResult.Rows.Count - 1
        If Not rtcResult.Rows(i).GroupRow Then
            If rtcResult.Rows(i).Record(tc医嘱ID).Value = lngAdvice Then
                GetIndexRow = i
                Exit Function
            End If
        End If
    Next
End Function

Private Function ExpandRow() As Long()
    Dim i As Long
    Dim arrRow() As Long
    
    ReDim arrRow(0)
    For i = 0 To rtcResult.Rows.Count - 1
        If i > rtcResult.Rows.Count - 1 Then
            ExpandRow = arrRow
            Exit Function
        End If
        If rtcResult.Rows(i).GroupRow Then
            If rtcResult.Rows(i).Expanded Then
                ReDim Preserve arrRow(UBound(arrRow) + 1)
                arrRow(UBound(arrRow)) = i
                
                rtcResult.Rows(i).Expanded = False
            End If
        End If
    Next
    ExpandRow = arrRow
End Function

Private Sub OpenRow(arrRow() As Long)
    Dim i As Long
    Dim j As Long
    
    For i = UBound(arrRow) To 1 Step -1
        If arrRow(i) <= rtcResult.Rows.Count - 1 Then
            If rtcResult.Rows(arrRow(i)).GroupRow Then
                rtcResult.Rows(arrRow(i)).Expanded = True
            End If
        End If
    Next
End Sub

'Private Sub CheckDo(ByRef strAdvice As String)
'    Dim i As Long
'    Dim arrAdvice() As Long
'
'    ReDim arrAdvice(0)
'    For i = rtcResult.Rows.Count - 1 To 0 Step -1
'        If rtcResult.Rows(i).GroupRow Then
'            If rtcResult.Rows(i).Selected Then
'                If Not rtcResult.Rows(i).Expanded Then
'                    rtcResult.Rows(i).Expanded = True
'                End If
'
'                strAdvice = strAdvice & "|" & rtcResult.Rows(i).Childs(0).Record(tc医嘱ID).Value
'                If rtcResult.Rows(i).Childs(0).Record(tc图像UID).Value = "0" Then
'                    ReDim Preserve arrAdvice(UBound(arrAdvice) + 1)
'                    arrAdvice(UBound(arrAdvice)) = rtcResult.Rows(i).Childs(0).Record(tc医嘱ID).Value
'                End If
'            End If
'        End If
'    Next
'
'    For i = 1 To UBound(arrAdvice)
'        Call GetImageInfo(arrAdvice(i))
'    Next
'
'End Sub
'
'
'Private Function GetDoAdvice() As String
'    Dim i As Long
'    Dim arrAdvice() As Long
'
'    ReDim arrAdvice(0)
'    For i = rtcResult.Rows.Count - 1 To 0 Step -1
'        If rtcResult.Rows(i).GroupRow Then
'            If Not rtcResult.Rows(i).Expanded Then
'                rtcResult.Rows(i).Expanded = True
'            End If
'
'            GetDoAdvice = GetDoAdvice & "|" & rtcResult.Rows(i).Childs(0).Record(tc医嘱ID).Value
'
'            If rtcResult.Rows(i).Childs(0).Record(tc图像UID).Value = "0" Then
'                ReDim Preserve arrAdvice(UBound(arrAdvice) + 1)
'                arrAdvice(UBound(arrAdvice)) = rtcResult.Rows(i).Childs(0).Record(tc医嘱ID).Value
'            End If
'        End If
'    Next
'
'    For i = 1 To UBound(arrAdvice)
'        Call GetImageInfo(arrAdvice(i))
'    Next
'End Function

Private Function GetAdvice() As String
    Dim i As Long
    Dim strAdvice As String
    
    For i = rtcResult.Rows.Count - 1 To 0 Step -1
        If rtcResult.Rows(i).GroupRow Then
            If Not rtcResult.Rows(i).Expanded Then
                rtcResult.Rows(i).Expanded = True
            Else
                If InStr(strAdvice, "[" & rtcResult.Rows(i).Childs(0).Record(tc医嘱ID).Value & "]") = 0 Then
                    strAdvice = strAdvice & "[" & rtcResult.Rows(i).Childs(0).Record(tc医嘱ID).Value & "]"
                End If
            End If
            
        End If
    Next
    
    GetAdvice = strAdvice
End Function

Private Sub DoDold(strAdvice As String)
On Error GoTo errH
    Dim i As Long
    
    For i = rtcResult.Rows.Count - 1 To 0 Step -1
        If rtcResult.Rows(i).GroupRow Then
            If InStr(strAdvice, "[" & rtcResult.Rows(i).Childs(0).Record(tc医嘱ID).Value & "]") = 0 Then
                rtcResult.Rows(i).Expanded = False
            End If
        End If
    Next
    Exit Sub
errH:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Function DoCheck() As Boolean
    Dim i As Long
    
    If rtcResult.SelectedRows.Count <= 0 Then Exit Function
    
    For i = 0 To rtcResult.SelectedRows.Count - 1
        If rtcResult.SelectedRows(i).GroupRow Then
            DoCheck = True
            Exit Function
        End If
    Next
End Function

Private Sub RefreshImagesInfo(ByVal strAdvice As String)
    Dim i As Long
    Dim arrAdvice() As String
    
    strAdvice = Replace(strAdvice, "][", "-")
    strAdvice = Replace(strAdvice, "]", "")
    strAdvice = Replace(strAdvice, "[", "")
    
    arrAdvice = Split(strAdvice, "-")
    
    For i = 0 To UBound(arrAdvice)
        If Val(arrAdvice(i)) > 0 Then
            Call GetImageInfo(Val(arrAdvice(i)), True)
        End If
    Next
End Sub

Private Function IgnoreResult() As Boolean
    Dim strSql As String
    
    
    With rtcResult.FocusedRow
        strSql = "zl_影像检查图象_校对('" & .Record(tc医嘱ID).Value & "','" & .Record(tc图像UID).Value & "',to_date('" & gobjComlib.zlDatabase.Currentdate & "','yyyy-mm-dd hh24:mi:ss'),6)"
    End With
    
    Call gobjComlib.zlDatabase.ExecuteProcedure(strSql, "忽略校对结果")
    mblnOk = True
    IgnoreResult = True
End Function

Private Function IsDBA() As Boolean
    Dim strSql As String
    Dim rsTmp As Recordset
    
    strSql = "select 所有者 from ZLSystems where 编号 = 100 and 名称 = '医院系统标准版'"
    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSql, "获取所有者")
    
    If rsTmp.RecordCount <= 0 Then Exit Function
    
    If Len(gstrUserName) = 0 Then
        Call getUser(gcnOracle.ConnectionString)
    End If
    
    If UCase(gstrUserName) = UCase(rsTmp("所有者")) Then
        IsDBA = True
    End If
End Function

Private Sub timerPopulate_Timer()
On Error GoTo errH
    If mstrAdvice <> "" And Not mblnShow And Not mblnSelect Then
        rtcResult.Populate
        
        Call DoDold(mstrAdvice)
        mstrAdvice = ""
        
    End If
    Exit Sub
errH:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub
