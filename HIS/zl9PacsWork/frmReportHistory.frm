VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmReportHistory 
   Caption         =   "报告修订历史"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9735
   Icon            =   "frmReportHistory.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6270
   ScaleWidth      =   9735
   StartUpPosition =   1  '所有者中心
   Begin RichTextLib.RichTextBox rtfEPR 
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   5520
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmReportHistory.frx":0CCA
   End
   Begin RichTextLib.RichTextBox rtxtReport 
      Height          =   4455
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   7858
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmReportHistory.frx":0D67
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   120
      Top             =   240
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmReportHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngAdviceID As Long        '医嘱ID
Private mlngPatientId As Long       '病人ID
Private mlngCur科室ID As Long       '科室ID
Private mlngReportID As Long        '报告ID
Private mlngMode As Long            '报告查看状态，0-修订状态，1-最终状态
Private mintReportCount As Integer  '历史报告的总数
Private mlngViewReportID As Long    '当前查看的报告ID
Private mlngViewAdviceID As Long    '当前查看的医嘱ID
Private mstrOffset As String        '当前行左边的缩进

Private mobjReport As zlRichEPR.cDockReport    '报告对象



Public Sub zlShowMe(frmParent As Object, lngAdviceID As Long, lngReportID As Long)
    mlngAdviceID = lngAdviceID
    mlngReportID = lngReportID
    mlngViewReportID = mlngReportID
    mlngViewAdviceID = mlngAdviceID
    Me.Show 0, frmParent
End Sub

Private Sub cbrMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_PacsReport_Mode_Orig                   '原始状态
            If mlngMode <> 0 Then
                ShowModeOrig mlngViewReportID, mlngViewAdviceID
            End If
            mlngMode = 0
            Me.cbrMain.FindControl(, conMenu_PacsReport_Mode_Clear, , True).Checked = False
            Control.Checked = True
        Case conMenu_PacsReport_Mode_Clear                  '最终状态
            If mlngMode <> 1 Then
                ShowModeClear mlngViewReportID, mlngViewAdviceID
            End If
            mlngMode = 1
            Me.cbrMain.FindControl(, conMenu_PacsReport_Mode_Orig, , True).Checked = False
            Control.Checked = True
        Case conMenu_File_Preview                           '报告预览
            If mlngViewReportID = 0 Then Exit Sub
            mobjReport.zlRefresh 0, 0
            mobjReport.zlRefresh mlngViewAdviceID, UserInfo.部门ID
            mobjReport.zlExecuteCommandBars Control
        Case conMenu_File_Exit                              '   退出
                Unload Me
        Case Else
            ShowHistory Control.ID
    End Select
End Sub

Private Sub cbrMain_Resize()
    Dim iLeft As Long, iTop As Long, iRight As Long, iBottom As Long
    cbrMain.GetClientRect iLeft, iTop, iRight, iBottom
    rtxtReport.Left = iLeft
    rtxtReport.Top = iTop
    rtxtReport.Width = Abs(iRight - iLeft)
    rtxtReport.Height = Abs(iBottom - iTop)
End Sub

Private Sub ShowHistory(iIndex As Integer)
    Dim lngReportID As Long
    Dim lngAdviceID As Long
    Dim strTemp As String
    
    If iIndex > conMenu_PacsReport_History_Times And iIndex <= conMenu_PacsReport_History_Times + mintReportCount Then
        strTemp = Me.cbrMain.FindControl(, iIndex, , True).DescriptionText
        If InStr(strTemp, "-") <> 0 Then
            lngReportID = Val(Split(strTemp, "-")(1))
            lngAdviceID = Val(Split(strTemp, "-")(0))
            mlngViewReportID = lngReportID
            mlngViewAdviceID = lngAdviceID
            If mlngMode = 0 Then
                Call ShowModeOrig(mlngViewReportID, mlngViewAdviceID)
            ElseIf mlngMode = 1 Then
                Call ShowModeClear(mlngViewReportID, mlngViewAdviceID)
            End If
        End If
    End If
End Sub

Private Sub ShowTitle(lngReportID As Long, lngAdviceID As Long)
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strSeparator1 As String
    Dim strSeparator2 As String
    Dim lngStart As Long
    Dim strTitle As String
    Dim strWriter As String
    Dim lng病人ID As Long
    Dim lng主页ID As Long
    Dim int婴儿 As Integer
    Dim rsBaby As ADODB.Recordset
    
    If lngReportID = 0 Then Exit Sub
    
    strSeparator1 = mstrOffset & "==================================================" & vbCrLf
    strSeparator2 = mstrOffset & "-------------------" & vbCrLf
    
    strSql = "Select a.姓名,a.检查号,b.开嘱时间,b.医嘱内容,a.报告人,a.复核人,nvl(b.婴儿,0) as 婴儿,a.接收日期 as 检查时间, " _
            & "b.病人ID, nvl(b.主页ID,0) as 主页ID From 影像检查记录 a,病人医嘱记录 b Where a.医嘱id = b.Id And b.Id = [1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngAdviceID)
    If rsTemp.EOF = True Then Exit Sub
    
    strTitle = mstrOffset & Nvl(rsTemp!医嘱内容) & vbCrLf
    
    lngStart = Len(rtxtReport.Text)
    rtxtReport.SelStart = lngStart
    rtxtReport.SelLength = 0
    rtxtReport.SelText = strTitle
    
    rtxtReport.SelStart = lngStart
    rtxtReport.SelLength = Len(strTitle)
    rtxtReport.SelFontName = "宋体"
    rtxtReport.SelFontSize = 16
    rtxtReport.SelBold = True
    rtxtReport.SelColor = vbBlue
    
    '婴儿的姓名需要特殊显示
    If rsTemp!婴儿 = 0 Then
        strWriter = vbCrLf & mstrOffset & "姓名：" & Nvl(rsTemp!姓名) & "      检查号：" & Nvl(rsTemp!检查号) & vbCrLf _
               & mstrOffset & "报告人：" & Nvl(rsTemp!报告人) & "      审核人：" & Nvl(rsTemp!复核人) & vbCrLf _
               & mstrOffset & "开嘱时间： " & Nvl(rsTemp!开嘱时间) & "      检查时间：" & Nvl(rsTemp!检查时间) & vbCrLf
    Else
        lng病人ID = rsTemp!病人ID
        lng主页ID = rsTemp!主页ID
        int婴儿 = rsTemp!婴儿
        strSql = "Select Decode(a.婴儿姓名,Null,b.姓名||'之子'||Trim(To_Char(a.序号,'9')),a.婴儿姓名) As 婴儿姓名,婴儿性别,出生时间 From 病人新生儿记录 a,病人信息 b Where a.病人id=[1] And a.主页id=[2] And a.病人id=b.病人id And a.序号=[3]"
        Set rsBaby = zlDatabase.OpenSQLRecord(strSql, "查找婴儿信息", lng病人ID, lng主页ID, int婴儿)
        
        strWriter = vbCrLf & mstrOffset & "姓名：" & rsBaby!婴儿姓名 & "      检查号：" & Nvl(rsTemp!检查号) & vbCrLf _
               & mstrOffset & "报告人：" & Nvl(rsTemp!报告人) & "      审核人：" & Nvl(rsTemp!复核人) & vbCrLf _
               & mstrOffset & "开嘱时间： " & Nvl(rsTemp!开嘱时间) & "      检查时间：" & Nvl(rsTemp!检查时间) & vbCrLf
    
    End If
'    '病历信息
'    strSQL = "Select 病历名称 From 电子病历记录  Where Id =  [1] "
'    Set rsTemp = OpenSQLRecord(strSQL, Me.Caption, lngReportID)
'    If rsTemp.EOF = True Then Exit Sub
'
'    strTitle = mstrOffset & Nvl(rsTemp!病历名称) & vbCrLf
'
'    lngStart = Len(rtxtReport.Text)
'    rtxtReport.SelStart = lngStart
'    rtxtReport.SelLength = 0
'    rtxtReport.SelText = strTitle
'
'    rtxtReport.SelStart = lngStart
'    rtxtReport.SelLength = Len(strTitle)
'    rtxtReport.SelFontName = "宋体"
'    rtxtReport.SelFontSize = 14
'    rtxtReport.SelBold = False
'    rtxtReport.SelColor = vbBlue
    
    '显示创建人
    strWriter = strWriter
    
    lngStart = Len(rtxtReport.Text)
    rtxtReport.SelStart = lngStart
    rtxtReport.SelLength = 0
    rtxtReport.SelText = strWriter
    
    rtxtReport.SelStart = lngStart
    rtxtReport.SelLength = Len(strWriter)
    rtxtReport.SelFontName = "宋体"
    rtxtReport.SelFontSize = 14
    rtxtReport.SelBold = False
    rtxtReport.SelColor = vbBlue
    
    '显示横线
    lngStart = Len(rtxtReport.Text)
    rtxtReport.SelStart = lngStart
    rtxtReport.SelLength = 0
    rtxtReport.SelText = strSeparator1
    
    rtxtReport.SelStart = lngStart
    rtxtReport.SelLength = Len(strSeparator1)
    rtxtReport.SelFontName = "宋体"
    rtxtReport.SelFontSize = 14
    rtxtReport.SelBold = False
    rtxtReport.SelColor = vbBlack
    
'    '签名信息
'    strSQL = "Select 内容文本 As 签名人 ,要素名称 As 签名前缀,对象标记 From 电子病历内容 b Where  b.对象类型=8 And 文件ID= [1] Order By 对象标记 "
'    Set rsTemp = OpenSQLRecord(strSQL, Me.Caption, lngReportID)
'    If rsTemp.EOF = True Then Exit Sub
'
'    strTitle = mstrOffset & "签名人：" & Nvl(rsTemp!签名前缀) & Nvl(rsTemp!签名人) & vbCrLf
'    rsTemp.MoveNext
'    While Not rsTemp.EOF
'        strTitle = strTitle & mstrOffset & "        " & Nvl(rsTemp!签名前缀) & Nvl(rsTemp!签名人) & vbCrLf
'        rsTemp.MoveNext
'    Wend
'    strTitle = strTitle & strSeparator1
'
'    lngStart = Len(rtxtReport.Text)
'    rtxtReport.SelStart = lngStart
'    rtxtReport.SelLength = 0
'    rtxtReport.SelText = strTitle
'
'    rtxtReport.SelStart = lngStart
'    rtxtReport.SelLength = Len(strTitle)
'    rtxtReport.SelFontName = "宋体"
'    rtxtReport.SelFontSize = 14
'    rtxtReport.SelBold = False
'    rtxtReport.SelColor = vbBlue
End Sub

Private Sub Form_Load()
    
    mlngMode = 1
    mintReportCount = 0
    mstrOffset = "  "
    Set mobjReport = New zlRichEPR.cDockReport      '电子病历报告
    
    Call RestoreWinState(Me, App.ProductName)
    
    Call InitCommandBars '初始化菜单
    
    If mlngReportID = 0 Then    '当前报告没有保存，直接显示最近的一次历史报告
        If mintReportCount >= 1 Then
            ShowHistory conMenu_PacsReport_History_Times + mintReportCount
        End If
    Else
        ShowModeClear mlngViewReportID, mlngViewAdviceID
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Unload mobjReport.zlGetForm        '电子病历报告
    '保存窗体位置
    Call SaveWinState(Me, App.ProductName)
End Sub


Private Sub InitCommandBars()
    '功能创建工具条
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim cbrPopControl As CommandBarControl
    Dim strSql  As String
    Dim strSQLBak As String
    Dim rsTemp As ADODB.Recordset
    
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrMain.VisualTheme = xtpThemeOffice2003
    Set Me.cbrMain.Icons = zlCommFun.GetPubIcons
    
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
    
    '采集工具栏定义
    Set cbrToolBar = Me.cbrMain.Add("报告历史", xtpBarLeft)
    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.Closeable = False
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_Mode_Orig, "原始状态")
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_Mode_Clear, "最终状态")
        cbrControl.Checked = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "报告预览")
        cbrControl.IconId = 102
        cbrControl.Style = xtpButtonIconAndCaption
        cbrControl.BeginGroup = True
        
        '增加历史报告的菜单，只有有历史报告的时候，才增加这个菜单
        strSql = "Select 病人ID,执行科室ID From 病人医嘱记录  Where Id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "", mlngAdviceID)
        If rsTemp.EOF = False Then
            mlngPatientId = Nvl(rsTemp!病人ID, 0)
            mlngCur科室ID = Nvl(rsTemp!执行科室ID, 0)
            
            strSql = "Select c.Id As 医嘱id,c.开嘱时间 As 开嘱时间,c.医嘱内容,b.病历Id  From 影像检查记录 a,病人医嘱报告 b,病人医嘱记录 c" _
                    & " Where a.医嘱id = c.Id And b.医嘱ID= c.Id And c.病人ID=[1] And c.相关ID Is Null And c.执行科室ID  in " _
                    & " (Select 部门ID From 部门人员 Where 人员ID =[2]) Order By 开嘱时间 Asc"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "", mlngPatientId, UserInfo.ID)
            
            If rsTemp.RecordCount > 1 Or (mlngReportID = 0 And rsTemp.RecordCount = 1) Then
                Set cbrControl = .Add(xtpControlPopup, conMenu_PacsReport_History_Times, "报告历史"): cbrControl.ID = conMenu_PacsReport_History_Times
                
                Do Until rsTemp.EOF
                   Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_PacsReport_History_Times + rsTemp.AbsolutePosition, "第" & rsTemp.AbsolutePosition & "次(" & Format(rsTemp!开嘱时间, "yyyy-mm-dd") & ") " & rsTemp!医嘱内容)
                   cbrPopControl.DescriptionText = rsTemp!医嘱ID & "-" & rsTemp!病历ID
                   rsTemp.MoveNext
                Loop
'                '如果当前正在编辑的报告还没有创建，则增加当前报告的菜单
'                If mlngReportID = 0 Then
'                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_PacsReport_History_Times + rsTemp.RecordCount + 1, "当前报告")
'                   cbrPopControl.DescriptionText = mlngAdviceID & "-0"
'                End If
                mintReportCount = rsTemp.RecordCount
            End If
        End If
        
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
        cbrControl.Style = xtpButtonIconAndCaption
      
    End With
    cbrToolBar.Position = xtpBarTop
End Sub

Public Sub ShowModeOrig(lngReportID As Long, lngAdviceID As Long)
    
    rtxtReport.Text = ""
    Call ShowTitle(lngReportID, lngAdviceID)
    Call ShowReportText(lngReportID, "检查所见")
    Call ShowReportText(lngReportID, "诊断意见")
    Call ShowReportText(lngReportID, "建议")
    
    rtxtReport.SelStart = 0
    rtxtReport.SelLength = 0
End Sub

Private Sub ShowReportText(lngReportID As Long, strType As String)
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim lngStart As Long
    Dim strText As String
    Dim strSeparator2 As String
    Dim strSeparator1 As String
    
    strSeparator1 = vbCrLf & mstrOffset & "-------" & vbCrLf
    strSeparator2 = vbCrLf ' & mstrOffset & "------------" & vbCrLf
    
    
    '读取报告的内容
    strSql = "Select a.内容文本 As 标题, b.对象属性, b.内容文本 As 正文,b.开始版 as 版本 From 电子病历内容 a,电子病历内容 b " & _
             " Where a.文件id = [1] And a.对象类型 = 3 And a.Id = b.父ID And b.对象类型 = 2 and a.内容文本 =[2] order by b.开始版  "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngReportID, strType)
    
    If rsTemp.EOF = False Then
        lngStart = Len(rtxtReport.Text)
        Select Case strType
            Case "检查所见"
                strText = strSeparator2 & mstrOffset & pReport_CheckViewName & strSeparator2
            Case "诊断意见"
                strText = vbCrLf & strSeparator2 & mstrOffset & pReport_ResultName & strSeparator2
            Case "建议"
                strText = vbCrLf & strSeparator2 & mstrOffset & pReport_AdviceName & strSeparator2
        End Select
        
        rtxtReport.SelStart = lngStart
        rtxtReport.SelLength = 0
        rtxtReport.SelText = strText
        
        rtxtReport.SelStart = lngStart
        rtxtReport.SelLength = Len(strText)
        rtxtReport.SelFontName = "宋体"
        rtxtReport.SelFontSize = 14
        rtxtReport.SelColor = vbBlue
        rtxtReport.SelBold = True
    End If
    
    While Not rsTemp.EOF
        lngStart = Len(rtxtReport.Text)
        strText = strSeparator1 & mstrOffset & "第 " & Nvl(rsTemp!版本) & " 版：" & strSeparator1 & mstrOffset & Nvl(rsTemp!正文) & vbCrLf
        rtxtReport.SelStart = lngStart
        rtxtReport.SelLength = 0
        rtxtReport.SelText = strText
        
        rtxtReport.SelStart = lngStart
        rtxtReport.SelLength = Len(strText)
        rtxtReport.SelFontName = "宋体"
        rtxtReport.SelFontSize = 14
        rtxtReport.SelColor = vbBlack
        rtxtReport.SelBold = False
        
        rsTemp.MoveNext
    Wend
End Sub

Public Sub ShowModeClear(lngReportID As Long, lngAdviceID As Long)
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim lngStart As Long
    Dim strText As String
    Dim strTitle As String
    Dim strSeparator2 As String
    Dim blnShow As Boolean
    
    strSeparator2 = vbCrLf 'vbCrLf & mstrOffset & "------------" & vbCrLf
    rtxtReport.Text = ""
    
    Call ShowTitle(lngReportID, lngAdviceID)
    
    '读取报告的内容
    strSql = "Select a.内容文本 As 标题, b.对象属性, b.内容文本 As 正文,b.开始版 as 版本 From 电子病历内容 a,电子病历内容 b " & _
             " Where a.文件id = [1] And a.对象类型 = 3 And a.Id = b.父ID And b.对象类型 = 2 and b.终止版=0  "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngReportID)
    
    While Not rsTemp.EOF
        blnShow = False
        Select Case rsTemp!标题
            Case "检查所见"
                strTitle = strSeparator2 & mstrOffset & pReport_CheckViewName & strSeparator2
                strText = vbCrLf & mstrOffset & Nvl(rsTemp!正文) & vbCrLf
                blnShow = True
            Case "诊断意见"
                strTitle = strSeparator2 & mstrOffset & pReport_ResultName & strSeparator2
                strText = vbCrLf & mstrOffset & Nvl(rsTemp!正文) & vbCrLf
                blnShow = True
            Case "建议"
                strTitle = strSeparator2 & mstrOffset & pReport_AdviceName & strSeparator2
                strText = vbCrLf & mstrOffset & Nvl(rsTemp!正文) & vbCrLf
                blnShow = True
        End Select
        
        If blnShow = True Then
            lngStart = Len(rtxtReport.Text)
            rtxtReport.SelStart = lngStart
            rtxtReport.SelLength = 0
            rtxtReport.SelText = strTitle
            
            rtxtReport.SelStart = lngStart
            rtxtReport.SelLength = Len(strTitle)
            rtxtReport.SelFontName = "宋体"
            rtxtReport.SelFontSize = 14
            rtxtReport.SelColor = vbBlue
            rtxtReport.SelBold = True
            
            lngStart = Len(rtxtReport.Text)
            rtxtReport.SelStart = lngStart
            rtxtReport.SelLength = 0
            rtxtReport.SelText = strText
            
            rtxtReport.SelStart = lngStart
            rtxtReport.SelLength = Len(strText)
            rtxtReport.SelFontName = "宋体"
            rtxtReport.SelFontSize = 14
            rtxtReport.SelColor = vbBlack
            rtxtReport.SelBold = False
        End If
            
        rsTemp.MoveNext
    Wend
    
    rtxtReport.SelStart = 0
    rtxtReport.SelLength = 0
    
    If Not blnShow Then
    'blnShow=true 说明存在表格，不填充电子病历内容
        Call FillERPWord
    End If
End Sub

Private Sub FillERPWord()
'填充电子病历格式的内容
On Error GoTo errH
    Dim strZipFile As String
    Dim strReportFormatFile As String
    Dim strTemp As String

    strReportFormatFile = ""
    
    strZipFile = zlBlobRead(5, mlngReportID, strReportFormatFile)
    strTemp = zlFileUnzip(strZipFile)
    rtfEPR.Filename = strTemp
    
    Call DoEPRReportFormat(rtfEPR)
    rtxtReport.Text = rtxtReport.Text & vbCrLf & "  " & rtfEPR.Text
    
    Kill strZipFile
    Exit Sub
errH:
    Kill strZipFile
    Call err.Raise(0, , "FillERPWord异常-" & err.Description)
    Resume
End Sub

Private Sub DoEPRReportFormat(ByRef rtfEPR As RichTextBox)
    '处理电子病历格式
On Error GoTo errH
    Dim i As Long
    Dim lngStart As Long
    Dim lngEnd As Long
    Dim blnContinu As Boolean
    Dim strNew As String
    Dim strOld As String
    Dim lngStartNext As Long ' 一组需要去掉的字符的开始位置
    Dim lngFlagPos As Long '"00000"的位置
    
    strOld = rtfEPR.TextRTF
    strNew = strOld
    blnContinu = True
    
    lngFlagPos = InStr(1, strNew, "00000")
    If lngFlagPos > 5 Then
        lngStartNext = lngFlagPos - 5
    Else
        blnContinu = False
    End If
    
    While blnContinu
        '去掉形如 ES(00000007,0,0)的部分 关键点  XX前面必定有一个空格
        lngStart = InStr(lngStartNext, strNew, " ")
        
        If lngStart > 0 Then
            lngStartNext = lngStart
        Else
            lngStartNext = Len(strNew) - 1
            blnContinu = False
        End If
        lngEnd = InStr(lngStart, strNew, ")")
        
        '去掉从空格后面一个到 ) 的内容
        If lngStart > 0 And lngEnd > 0 And lngEnd - lngStart > 0 And lngEnd - lngStart < 20 Then
            strNew = Mid(strNew, 1, lngStart) & Mid(strNew, lngEnd + 1)
        Else
            lngFlagPos = InStr(lngFlagPos + 10, strNew, "00000")
            If lngFlagPos < 5 Then
                blnContinu = False
            Else
                lngStartNext = lngFlagPos - 5
            End If
        End If
    Wend

    strNew = Replace(strNew, "\par ", "\par   ")
    rtfEPR.Text = ""
    rtfEPR.TextRTF = strNew
    Exit Sub
errH:
    rtfEPR.TextRTF = strOld
    If App.LogMode = 0 Then MsgBox "DoEPRReportFormat调试错误" & err.Description
End Sub
