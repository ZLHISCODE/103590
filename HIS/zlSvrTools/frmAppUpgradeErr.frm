VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{CA73588D-282F-4592-9369-A61CC244FADA}#15.3#0"; "Codejock.SyntaxEdit.v15.3.1.ocx"
Begin VB.Form frmAppUpgradeErr 
   AutoRedraw      =   -1  'True
   Caption         =   "错误"
   ClientHeight    =   7845
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   11250
   Icon            =   "frmAppUpgradeErr.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   11250
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fraSplit 
      Height          =   30
      Index           =   1
      Left            =   -60
      TabIndex        =   12
      Top             =   7200
      Width           =   11292
   End
   Begin VB.Frame fraSplit 
      Height          =   30
      Index           =   0
      Left            =   -60
      MousePointer    =   7  'Size N S
      TabIndex        =   11
      Top             =   3720
      Width           =   11292
   End
   Begin VB.PictureBox picErrInfo 
      BorderStyle     =   0  'None
      Height          =   3612
      Left            =   120
      ScaleHeight     =   3615
      ScaleWidth      =   11055
      TabIndex        =   8
      Top             =   60
      Width           =   11052
      Begin RichTextLib.RichTextBox rtfErr 
         Height          =   3336
         Left            =   0
         TabIndex        =   9
         Top             =   276
         Width           =   11052
         _ExtentX        =   19500
         _ExtentY        =   5874
         _Version        =   393217
         BackColor       =   -2147483633
         BorderStyle     =   0
         Enabled         =   -1  'True
         HideSelection   =   0   'False
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmAppUpgradeErr.frx":6852
      End
      Begin VB.Label lblErrInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "错误描述："
         ForeColor       =   &H00404040&
         Height          =   180
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   900
      End
   End
   Begin VB.PictureBox picModify 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3252
      Left            =   120
      ScaleHeight     =   3255
      ScaleWidth      =   11055
      TabIndex        =   4
      Top             =   3840
      Width           =   11052
      Begin XtremeSyntaxEdit.SyntaxEdit synModiSQL 
         Height          =   1812
         Left            =   -240
         TabIndex        =   6
         Top             =   600
         Width           =   11172
         _Version        =   983043
         _ExtentX        =   19706
         _ExtentY        =   3196
         _StockProps     =   84
         Text            =   "Drop Index 病人手麻记录_IX_主页ID;"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         EnableSyntaxColorization=   -1  'True
         ShowLineNumbers =   -1  'True
         ShowSelectionMargin=   -1  'True
         ShowScrollBarVert=   -1  'True
         ShowScrollBarHorz=   0   'False
         EnableVirtualSpace=   0   'False
         EnableAutoIndent=   -1  'True
         ShowWhiteSpace  =   0   'False
         ShowCollapsibleNodes=   -1  'True
         AutoCompleteWndWidth=   160
      End
      Begin VB.Label lblSQLErr 
         Caption         =   "执行结果："
         Height          =   612
         Left            =   0
         TabIndex        =   7
         Top             =   2520
         Width           =   10932
      End
      Begin VB.Label lblModify 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAppUpgradeErr.frx":6BF0
         ForeColor       =   &H00404040&
         Height          =   540
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   10860
      End
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   11250
      TabIndex        =   0
      Top             =   7260
      Width           =   11256
      Begin VB.Timer tmrRefresh 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   0
         Top             =   0
      End
      Begin VB.Timer tmrThis 
         Enabled         =   0   'False
         Interval        =   60000
         Left            =   3000
         Top             =   120
      End
      Begin VB.CommandButton cmdAbort 
         Caption         =   "中止(&A)"
         Height          =   350
         Left            =   9876
         TabIndex        =   3
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdIgnore 
         Caption         =   "忽略(&I)"
         Height          =   350
         Left            =   8772
         TabIndex        =   2
         Tag             =   "8175"
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdRetry 
         Caption         =   "重试(&R)"
         Height          =   350
         Left            =   7680
         TabIndex        =   1
         Tag             =   "7080"
         Top             =   120
         Width           =   1100
      End
   End
End
Attribute VB_Name = "frmAppUpgradeErr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'错误信息
Private Type ErrInfo
    ErrNum              As Long
    ErrDesc             As String
    ErrAdvice           As String
    ErrAdviceType       As Integer
    ErrPos              As String
    ErrSQL              As String
    ErrOparate          As VbMsgBoxResult
    ErrOparateModi      As VbMsgBoxResult   '填写修正SQL后的建议
    ErrModiSQL          As String           '错误的修正SQL
End Type
Private merrCur         As ErrInfo
Private mstrUser        As String
Private mcnThis         As ADODB.Connection
Private mfrmParent      As Object
Private mobjSQL         As clsSQLInfo
Private mobjPreSQL      As clsSQLInfo '上一个SQL
Private mblnIgnoreErr   As Boolean
Private mclsrun         As clsRunScript
Private mintTimes       As Long
Private mblnAuto        As Boolean '是否是Timer时间自动执行
Private mblnModify      As Boolean '错误修正模式
Private mblnShut        As Boolean  '判断是否直接关闭
Public Function ShowError(ByVal cnThis As ADODB.Connection, ByVal lngErrNum As Long, ByVal strErrInfo As String, ByVal objSQL As clsSQLInfo, frmParent As Object, Optional ByVal blnIgnoreErr As Boolean = True, Optional ByRef blnSysIgnore As Boolean, Optional ByVal clsRun As clsRunScript, Optional ByRef blnErrRepaired As Boolean) As VbMsgBoxResult
'blnErrRepaired=错误是否被该窗体自动修复，如果修复了，外面不再写日志
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errH
    blnErrRepaired = False
    merrCur.ErrNum = lngErrNum
    merrCur.ErrDesc = strErrInfo
    merrCur.ErrOparate = 0
    merrCur.ErrOparateModi = 0
    merrCur.ErrModiSQL = ""
    Set mobjSQL = objSQL
    If Not clsRun Is mclsrun Then Set mclsrun = clsRun
    Set mfrmParent = frmParent
    mblnIgnoreErr = blnIgnoreErr
    If Not cnThis Is mcnThis Then
        Set mcnThis = cnThis
        strSQL = "Select User From Dual"
        Set rsTmp = gclsBase.OpenSQLRecord(cnThis, strSQL, App.Title)
        mstrUser = rsTmp!User
    End If
    On Error Resume Next
    Call GetAdviceFromError
    '判断错误的是否可以自动忽略
    If blnIgnoreErr Then
        If merrCur.ErrOparate = vbIgnore Then
            ShowError = merrCur.ErrOparate
            blnSysIgnore = True        '系统建议忽略
            Unload Me
            Exit Function
        '生成修正SQL后，建议忽略
        ElseIf ModifyErrors(True) Then '执行修正SQL成功，则自动忽略
            blnErrRepaired = True
            merrCur.ErrOparate = vbIgnore
            ShowError = merrCur.ErrOparate
            blnSysIgnore = True        '系统建议忽略
            Unload Me
            Exit Function
        End If
    End If
    On Error GoTo errH
    mblnShut = True
    Me.Show 1, frmParent
    ShowError = merrCur.ErrOparate
    Exit Function
errH:
     MsgBox err.Description, vbInformation, App.Title
     If 0 = 1 Then
        Resume
     End If
End Function

Private Sub cmdAbort_Click()
    If MsgBox("系统必须要完整地进行升迁之后才能正常使用。确实要中止升迁操作吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    merrCur.ErrOparate = vbAbort
    mblnShut = False
    Unload Me
End Sub

Private Sub cmdIgnore_Click()
    '以避免连续击键误操作
    If cmdIgnore.Tag = "1" Then Exit Sub
    cmdIgnore.Tag = "1"
    If ModifyErrors Or mblnAuto Then
        merrCur.ErrOparate = vbIgnore
        mblnShut = False
        cmdIgnore.Tag = ""
        Unload Me
    Else
        cmdIgnore.Tag = ""
    End If
End Sub

Private Sub cmdRetry_Click()
    '以避免连续击键误操作
    If cmdRetry.Tag = "1" Then Exit Sub
    cmdIgnore.Tag = "1"
    '没有错误则直接重试
    If ModifyErrors Or mblnAuto Then
        merrCur.ErrOparate = vbRetry
        mblnShut = False
        cmdRetry.Tag = ""
        Unload Me
    Else
        cmdRetry.Tag = ""
    End If
End Sub

Private Function GetFormatSQL(ByVal strSQL As String) As String
    Dim arrFMT As Variant, i As Long, strReturn As String
    '获取用于分析的标准SQL串
    arrFMT = Split(Replace(Replace(strSQL, vbCrLf, vbCr), vbLf, vbCr), vbCr)
    For i = 0 To UBound(arrFMT)
        strReturn = strReturn & " " & TrimComment(arrFMT(i))
    Next
    strReturn = UCase(TrimEx(strReturn))
    GetFormatSQL = strReturn
End Function

Private Function ModifyErrors(Optional ByVal blnErrAutoAdjust As Boolean) As Boolean
'功能：执行修正SQL
    Dim strSQL As String, strErr As String
    Dim strLine As String, i As Long
    Dim strLogSQL As String
    Dim objScript As clsRunScript
    Dim blnHaveErr As Boolean, blnHaveSQL As Boolean
    Dim lngAffect As Long
    Dim strOldModfy As String
    Dim datBegin As Date, datEnd As Date, lngSQLTime As Long
    
    '执行修正脚本。
    If Not blnErrAutoAdjust Then
        If mblnModify Then
            Set objScript = New clsRunScript
            For i = Val(synModiSQL.Tag) + 1 To synModiSQL.RowsCount
                strSQL = strSQL & IIf(strSQL = "", "", vbNewLine) & synModiSQL.RowText(i)
            Next
            If objScript.AnalysisSQLString(strSQL, Val(synModiSQL.Tag) + 1) Then
                On Error Resume Next
                Do While Not objScript.EOF
                    strLogSQL = GetLogSQL(objScript.SQLInfo): blnHaveSQL = True
                    lblSQLErr.Caption = "正在执行SQL(" & objScript.Line & "行)：" & strLogSQL
                    If Not mclsrun Is Nothing Then mclsrun.WriteLog String(17, " ") & "错误中心(人工)修正SQL：" & strLogSQL
                    err.Clear
                    datBegin = Now: datEnd = Now
                    DoEvents
                    mcnThis.Execute objScript.SQLInfo.SQL, lngAffect, adCmdText
                    If err.Number <> 0 Then
                        If mcnThis.Errors.Count > 0 Then
                            lblSQLErr.Caption = lblSQLErr.Caption & vbNewLine & "执行出错，错误信息：" & mcnThis.Errors(0).Description
                            If Not mclsrun Is Nothing Then mclsrun.WriteLog String(17, " ") & "错误中心(人工)错误：" & mcnThis.Errors(0).Description
                        Else
                             lblSQLErr.Caption = lblSQLErr.Caption & vbNewLine & "执行出错，错误信息：" & err.Description
                            If Not mclsrun Is Nothing Then mclsrun.WriteLog String(17, " ") & "错误中心(人工)错误：" & err.Description
                        End If
                        blnHaveErr = True: err.Clear
                        Exit Do
                    Else
                        lblSQLErr.Caption = lblSQLErr.Caption & vbNewLine & "执行成功！"
                        If Not mclsrun Is Nothing Then mclsrun.WriteLog String(17, " ") & "错误中心(人工)结果：执行成功" & IIf(lngAffect > 0, "," & lngAffect & " 行数据生效", ",0行数据生效")
                        lblSQLErr.Tag = objScript.Line
                    End If
                    If mclsrun.SQLRecTime <> 0 Then
                        lngSQLTime = DateDiff("n", datBegin, datEnd)
                        If lngSQLTime >= mclsrun.SQLRecTime Then
                            mclsrun.WriteLog String(17, " ") & "错误中心(人工)SQL处理耗时：" & lngSQLTime & "分钟"
                        End If
                    End If
                    objScript.ReadNextSQL
                Loop
            End If
        End If
    'showErr自动执行SQL
    Else
        If merrCur.ErrModiSQL = "" Then
            Exit Function
        End If
        strOldModfy = merrCur.ErrModiSQL
        '执行修正SQL
        strSQL = merrCur.ErrModiSQL
        Set objScript = New clsRunScript
        If objScript.AnalysisSQLString(strSQL, 0) Then
            On Error Resume Next
            Do While Not objScript.EOF
                strLogSQL = GetLogSQL(objScript.SQLInfo): blnHaveSQL = True
                If Not mclsrun Is Nothing Then mclsrun.WriteLog String(17, " ") & "错误中心(系统)修正SQL：" & strLogSQL
                err.Clear
                datBegin = Now: datEnd = Now
                DoEvents
                mcnThis.Execute objScript.SQLInfo.SQL, lngAffect, adCmdText
                If err.Number <> 0 Then
                    If mcnThis.Errors.Count > 0 Then
                        If Not mclsrun Is Nothing Then mclsrun.WriteLog String(17, " ") & "错误中心(系统)错误：" & mcnThis.Errors(0).Description
                    Else
                        If Not mclsrun Is Nothing Then mclsrun.WriteLog String(17, " ") & "错误中心(系统)错误：" & err.Description
                    End If
                    blnHaveErr = True: err.Clear
                    Exit Do
                Else
                    If Not mclsrun Is Nothing Then mclsrun.WriteLog String(17, " ") & "错误中心(系统)结果：执行成功" & IIf(lngAffect > 0, "," & lngAffect & " 行数据生效", ",0行数据生效")
                End If
                If mclsrun.SQLRecTime <> 0 Then
                    lngSQLTime = DateDiff("n", datBegin, datEnd)
                    If lngSQLTime >= mclsrun.SQLRecTime Then
                        mclsrun.WriteLog String(17, " ") & "错误中心(系统)SQL处理耗时：" & lngSQLTime & "分钟"
                    End If
                End If
                objScript.ReadNextSQL
            Loop
        End If
        If Not blnHaveErr Then
            '执行修正后重试原来的SQL,在这里处理，防止发生死循环
            If merrCur.ErrOparateModi = vbRetry Then
                If err.Number <> 0 Then err.Clear
                mclsrun.WriteLog String(17, " ") & "错误中心(系统)操作：执行修正SQL后重试原SQL"
                err.Clear: On Error Resume Next
                datBegin = Now: datEnd = Now
                DoEvents
                mcnThis.Execute mobjSQL.SQL, lngAffect, adCmdText
                datEnd = Now
                If err.Number = 0 Then
                    mclsrun.WriteLog String(17, " ") & "错误中心(系统)结果:成功" & IIf(lngAffect > 0, "," & lngAffect & " 行数据生效", ",0行数据生效")
                    If mclsrun.SQLRecTime <> 0 Then
                        lngSQLTime = DateDiff("n", datBegin, datEnd)
                        If lngSQLTime >= mclsrun.SQLRecTime Then
                            mclsrun.WriteLog String(17, " ") & "错误中心(系统)SQL处理耗时：" & lngSQLTime & "分钟"
                        End If
                    End If
                    blnHaveErr = False
                Else
                    '再次查看错误建议
                    If mcnThis.Errors.Count > 0 Then
                        merrCur.ErrNum = mcnThis.Errors(0).NativeError
                        merrCur.ErrDesc = mcnThis.Errors(0).Description
                    Else
                        merrCur.ErrNum = err.Number
                        merrCur.ErrDesc = err.Description
                    End If
                    merrCur.ErrDesc = Replace(merrCur.ErrDesc, "[Microsoft][ODBC driver for Oracle][Oracle]", "")
                    merrCur.ErrModiSQL = ""
                    mclsrun.WriteLog String(17, " ") & "错误中心(系统)结果:" & merrCur.ErrDesc
                    Call GetAdviceFromError
                     '错误修正SQL不能解决问题，则自动清空，建议变为原来的
                    If strOldModfy = merrCur.ErrModiSQL Then
                        merrCur.ErrModiSQL = ""
                        merrCur.ErrOparateModi = merrCur.ErrOparate
                    End If
                    '再次检查发现变为可以自动忽略
                    If merrCur.ErrOparate = vbIgnore Then
                        blnHaveErr = False
                    Else
                        blnHaveErr = True
                    End If
                End If
            '修正错误后忽略原始SQL
            ElseIf merrCur.ErrOparateModi = vbIgnore Then
                blnHaveErr = False
            End If
        Else
            '错误修正SQL不能解决问题，则自动清空，建议变为原来的
            merrCur.ErrModiSQL = ""
            merrCur.ErrOparateModi = merrCur.ErrOparate
        End If
    End If
    '没有错误则直接重试
    If blnHaveErr Then
    ElseIf mblnModify And Not blnHaveSQL Then
        blnHaveErr = True
    End If
    If Not blnErrAutoAdjust Then Call RefreshButton
    ModifyErrors = Not blnHaveErr
End Function

Private Sub RefreshButton()
'功能：刷新按钮显示文字
    Dim blnDo As Boolean
    blnDo = Trim(synModiSQL.RowText(Val(synModiSQL.Tag) + 1)) <> "" Or synModiSQL.RowsCount > Val(synModiSQL.Tag) + 1
    lblSQLErr.Visible = blnDo
    synModiSQL.Height = lblSQLErr.Top - 60 - synModiSQL.Top + IIf(blnDo, 0, lblSQLErr.Height)
    If blnDo Then
        cmdRetry.Width = 2800
        cmdIgnore.Width = 2800
        cmdRetry.Caption = "执行修正脚本并重试原脚本(&R)"
        cmdIgnore.Caption = "执行修正脚本并忽略原脚本(&I)"
        mblnModify = True
    Else
        cmdRetry.Width = 1100
        cmdIgnore.Width = 1100
        cmdRetry.Caption = "重试(&R)"
        cmdIgnore.Caption = "忽略(&I)"
        Me.Refresh
        lblSQLErr.Caption = ""
        mblnModify = False
    End If
    cmdRetry.Left = picBottom.ScaleWidth - (cmdRetry.Width * 2 + cmdAbort.Width + 100 + 30 * 2)
    Call SetCtrlPosOnLine(False, 0, cmdRetry, 30, cmdIgnore, 30, cmdAbort)
End Sub

Private Sub Form_Activate()
    rtfErr.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If Not Me.ActiveControl Is synModiSQL Then
        If UCase(Chr(KeyAscii)) = "R" Then
            KeyAscii = 0
            If cmdRetry.Enabled Then cmdRetry_Click
        ElseIf UCase(Chr(KeyAscii)) = "I" Then
            KeyAscii = 0
            If cmdIgnore.Enabled Then cmdIgnore_Click
        ElseIf UCase(Chr(KeyAscii)) = "A" Then
            KeyAscii = 0
            If cmdAbort.Enabled Then cmdAbort_Click
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim strPos As String, strTemp As String
    Dim strColor As String
    Dim objFSO As New Scripting.FileSystemObject
    Dim strPath As String
    
    '显示出错执行用户
    Me.Caption = "错误 - " & mstrUser
    '错误窗体位置处理
    Me.Left = mfrmParent.Left + (mfrmParent.Width - Me.Width) / 2
    Me.Top = mfrmParent.Top + (mfrmParent.Height - Me.Height) / 2
    strTemp = GetSetting("ZLSOFT", "私有模块\" & gstrUserName & "\界面设置\" & App.ProductName & Me.name & "\Form", "相对位置", "")
    If UBound(Split(strTemp, ",")) = 1 Then
        Me.Left = mfrmParent.Left + Val(Split(strTemp, ",")(0))
        Me.Top = mfrmParent.Top + Val(Split(strTemp, ",")(1))
    End If
    If Me.Left < 0 Then Me.Left = 0
    If Me.Top < 0 Then Me.Top = 0
    
    merrCur.ErrSQL = IIf(mobjSQL.Tip <> "", mobjSQL.Tip & vbCrLf, "") & mobjSQL.SQL
    If mobjSQL.File = "" Then
        merrCur.ErrPos = ""
    Else
        merrCur.ErrPos = "文件：" & mobjSQL.File & "  " & "行号：" & mobjSQL.FileLine
    End If
    
    rtfErr.Text = ""
    '编号信息
    rtfErr.Text = rtfErr.Text & "【编号】": rtfErr.SelStart = 1: rtfErr.SelLength = Len("【编号】"): rtfErr.SelFontSize = 14: rtfErr.SelBold = True
    rtfErr.Text = rtfErr.Text & merrCur.ErrNum & vbNewLine: strPos = Len(rtfErr.Text)
    '错误信息
    rtfErr.Text = rtfErr.Text & "【信息】": rtfErr.SelStart = Len(rtfErr.Text) + 1: rtfErr.SelLength = Len("【信息】"): rtfErr.SelFontSize = 14: rtfErr.SelBold = True
    rtfErr.Text = rtfErr.Text & merrCur.ErrDesc & vbNewLine: strPos = strPos & "," & Len(rtfErr.Text)
    '建议信息
    rtfErr.Text = rtfErr.Text & "【建议】": rtfErr.SelStart = Len(rtfErr.Text) + 1: rtfErr.SelLength = Len("【建议】"): rtfErr.SelFontSize = 14: rtfErr.SelBold = True
    rtfErr.Text = rtfErr.Text & merrCur.ErrAdvice & vbNewLine: strPos = strPos & "," & Len(rtfErr.Text)
    '建议信息
    rtfErr.Text = rtfErr.Text & "【位置】": rtfErr.SelStart = Len(rtfErr.Text) + 1: rtfErr.SelLength = Len("【位置】"): rtfErr.SelFontSize = 14: rtfErr.SelBold = True
    rtfErr.Text = rtfErr.Text & merrCur.ErrPos & vbNewLine: strPos = strPos & "," & Len(rtfErr.Text)
    'SQL信息
    rtfErr.Text = rtfErr.Text & "【SQL 】": rtfErr.SelStart = Len(rtfErr.Text) + 1: rtfErr.SelLength = Len("【位置】"): rtfErr.SelFontSize = 14: rtfErr.SelBold = True
    rtfErr.Text = rtfErr.Text & merrCur.ErrSQL: strPos = strPos & "," & Len(rtfErr.Text)
    
    '语法控件颜色方案
    synModiSQL.Font.name = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\SQLFont", "FontName", "Fixedsys")
    synModiSQL.Font.Size = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\SQLFont", "FontSize", 12)
    synModiSQL.Font.Underline = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\SQLFont", "FontUnderline", 0)
    synModiSQL.Font.Italic = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\SQLFont", "FontItalic", 0)
    synModiSQL.Font.Bold = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\SQLFont", "FontBold", 0)
    synModiSQL.Font.Strikethrough = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\SQLFont", "FontStrikethru", 0)
    synModiSQL.BorderStyle = xtpBorderClientEdge
    
    '设置控件的显示颜色方案为：SQL
    If Not gblnInIDE Then '增加多环境支持
        strPath = App.Path & "\PUBLIC\_sql.schclass"
    Else
        strPath = objFSO.GetParentFolderName(GetSetting("ZLSOFT", "公共全局", "程序路径")) & "\PUBLIC\_sql.schclass"
    End If
    If Not objFSO.FileExists(strPath) Then
        strPath = "C:\Appsoft\PUBLIC\_sql.schclass"
    End If
    If objFSO.FileExists(strPath) Then
        strColor = ReadFileToString(strPath)
    Else
        strColor = ""
    End If
    synModiSQL.SyntaxSet "[Schemes]" & vbCrLf & "SQL" & vbCrLf & "[Themes]" & vbCrLf & "Default" & vbCrLf & "Alternative" & vbCrLf
    synModiSQL.SyntaxScheme = strColor
    synModiSQL.Text = " "
    synModiSQL.Tag = synModiSQL.RowsCount - 1
    lblSQLErr.Tag = synModiSQL.Tag
    lblSQLErr.Caption = ""
    synModiSQL.CurrPos.Row = synModiSQL.RowsCount
    synModiSQL.Text = merrCur.ErrModiSQL
    If merrCur.ErrOparate = vbRetry Or merrCur.ErrOparateModi = vbRetry Then
        cmdRetry.FontBold = True
    ElseIf merrCur.ErrOparate = vbIgnore Or merrCur.ErrOparateModi = vbIgnore Then
        cmdIgnore.FontBold = True
    ElseIf merrCur.ErrOparate = vbAbort Or merrCur.ErrOparateModi = vbAbort Then
        cmdAbort.FontBold = True
    End If
    tmrThis.Interval = glngInterval
    tmrThis.Enabled = glngAtuoErr > 1
    tmrRefresh.Enabled = True
    mblnAuto = False
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    '窗体大小位置调整
    If Me.WindowState <> vbMinimized Then
        If Me.Height <= 6000 Then
            Me.Height = 6000
        End If
        If Me.Width < 8000 Then
            Me.Width = 8000
        End If
    End If
    If Me.WindowState <> vbMaximized Then
        If Me.Left + Me.Width > Screen.Width Then Me.Left = Screen.Width - Me.Width
        If Me.Top + Me.Height > Screen.Height Then Me.Top = Screen.Height - Me.Height
    End If
    'picBottom定位
    picBottom.Top = Me.ScaleHeight - picBottom.Height - 30
    If Me.WindowState <> vbMinimized Then
        picBottom.ScaleWidth = Me.ScaleWidth
    End If
    '先摆正两个分割线,以及下方PIC的位置
    fraSplit(1).Top = picBottom.Top - fraSplit(1).Height - 15
    picModify.Top = fraSplit(0).Top + fraSplit(0).Height + 30
    '调整上下两个pic已经分割线的高度与宽度
    picErrInfo.Height = fraSplit(0).Top - picErrInfo.Top - 30
    picModify.Height = fraSplit(1).Top - picModify.Top - 30
    picErrInfo.Width = Me.ScaleWidth - picErrInfo.Left
    picModify.Width = Me.ScaleWidth - picModify.Left
    fraSplit(0).Width = Me.ScaleWidth - fraSplit(0).Left
    fraSplit(1).Width = Me.ScaleWidth - fraSplit(1).Left
    '调正底部按钮位置
    Call RefreshButton
    If err.Number <> 0 Then err.Clear
    On Error GoTo 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strTemp As String
    
    tmrRefresh.Enabled = False
    If mblnShut Then '直接关闭窗体，默认采取中止操作
        If MsgBox("系统必须要完整地进行升迁之后才能正常使用。确实要中止升迁操作吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1
            Exit Sub
        Else
            merrCur.ErrOparate = vbAbort
        End If
    End If
    If glngAtuoErr > 0 Then Set mobjPreSQL = mobjSQL.CopyMe
    Set mobjSQL = Nothing
    '保存窗体位置
    strTemp = Me.Left - mfrmParent.Left & "," & Me.Top - mfrmParent.Top
    SaveSetting "ZLSOFT", "私有模块\" & gstrUserName & "\界面设置\" & App.ProductName & Me.name & "\Form", "相对位置", strTemp
End Sub

Private Sub GetAdviceFromError()
'功能：根据不同的Oracle错误，设置相应的操作建议说明
    Dim strSQL As String
    Dim strOwner As String, strName As String, strType As String
    Dim strTemp As String

    '首先确定是Oracle的错误才进行处理
    If InStr(merrCur.ErrDesc, "ORA-") = 0 Then
        merrCur.ErrAdvice = "升迁工具内部错误，请尝试重试操作"
        merrCur.ErrOparate = vbRetry
        Exit Sub
    End If
    '块语句
    If mobjSQL.Block Then
        merrCur.ErrAdvice = "请仔细检查，并作必要的处理后尝试重试操作"
        merrCur.ErrOparate = vbRetry
        Select Case merrCur.ErrNum
            Case 1 'ORA-00001: 违反唯一约束条件 (ZLTOOLS.XXX_PK)(ZLTOOLS.XXX_UQ_YYY)
                CheckAdjustSequence (merrCur.ErrDesc)
            Case 1502
                'ORA-01502: index 'XXXX' or partition of such index is in unusable state
                'ORA-01502: 索引 'XXXX' 或这类索引的分区处于不可用状态
                merrCur.ErrAdvice = "请重建索引后重试。"
                merrCur.ErrOparate = vbRetry
                Call CheckAdjustIndex(merrCur.ErrDesc)
            Case 6575
                'ORA-06575: 程序包或函数 TTT 处于无效状态
                'ORA-06575: Package or function TTT is in an invalid state
                merrCur.ErrAdvice = "请先正确编译对应的函数或过程后再重试。"
                merrCur.ErrOparate = vbRetry
                Call CheckAdjustProcedure(merrCur.ErrDesc)
            Case 12899 'ORA-12899: 列 "ZLHIS"."体检任务人员"."所属团体" 的值太大 (实际值: 62, 最大值: 60)
                'ORA-12899: value too large for column "SYSTEM"."STUDENTINFO"."SNAME" (actual: 78, maximum: 30)
                merrCur.ErrAdvice = "请先调整字段到合适的精度后再重试。"
                merrCur.ErrOparate = vbRetry
                Call CheckAdjust12899(merrCur.ErrDesc)
        End Select
    Else
        Select Case merrCur.ErrNum
            Case 1 'ORA-00001: 违反唯一约束条件 (ZLTOOLS.XXX_PK)(ZLTOOLS.XXX_UQ_YYY)
                merrCur.ErrOparate = vbRetry
                If mobjSQL.PartSQL Like "INSERT INTO *" Then
                    Call CheckAdjustSequence(merrCur.ErrDesc)
                    Call CheckAdjustTableData(mobjSQL.SQL, IIf(mobjSQL.PartSQL Like "INSERT INTO *VALUES*", 0, 1), merrCur.ErrDesc)
                ElseIf mobjSQL.PartSQL Like "UPDATE *" Then
                    merrCur.ErrAdvice = "请检查要更新数据的正确性。"
                Else
                    merrCur.ErrAdvice = "可能是语句重复运行出错，请检查。"
                End If
            Case 955 'ORA-00955: 名称已被现有对象占用
                strSQL = GetFormatSQL(mobjSQL.SQL)
                Call GetCreateName(strSQL, strOwner, strName, strType)
                merrCur.ErrAdvice = "可能是语句重复运行出错，一般情况下可以忽略该错误。"
                merrCur.ErrOparate = vbIgnore
                If strName <> "" Then
                    If Not ObjectExists(strOwner, strName, strType) Then
                        merrCur.ErrAdvice = "SQL语句要创建的对象已被其他类型的同名对象占用，请手工处理后重试。"
                        merrCur.ErrOparate = vbRetry
                    ElseIf strType = "TABLE" And Not strSQL Like "* AS SELECT *" Then  'Create Table XXX As Select方式，不作仔细分析
                        strTemp = CheckCreateTabCol(strSQL, strOwner, strName)
                        If strTemp <> "" Then
                            merrCur.ErrAdvice = "SQL语句要创建的对象已被其他类型的同名对象占用,但是两者结构存在差异，请手工处理后重试。"
                            merrCur.ErrOparate = vbRetry
                        End If
                    End If
                Else
                    Call CheckAdjustConstraint(strSQL)
                End If
            Case 1430 'ORA-01430: 表中已经存在要添加的列
                strSQL = GetFormatSQL(mobjSQL.SQL)
                merrCur.ErrAdvice = "可能是语句重复运行出错，一般情况下可以忽略该错误。"
                merrCur.ErrOparate = vbIgnore
                Call CheckTabChangeCol(strSQL)
            Case 2260, 2261
                strSQL = GetFormatSQL(mobjSQL.SQL)
                'ORA-02260: 表只能具有一个主关键字
                'ORA-02261: 表中已存在这样的唯一关键字或主关键字
                merrCur.ErrAdvice = "可能是语句重复运行出错，一般情况下可以忽略该错误。"
                merrCur.ErrOparate = vbIgnore
                Call CheckUniqueKeyCol(strSQL)
            Case 2291
                'ORA-02291: 违反完整约束条件 (ZLTOOLS.XXX_FK_YYID) - 未找到父项关键字
                If mobjSQL.PartSQL Like "INSERT INTO ZLRPT*" Then
                    merrCur.ErrAdvice = "可能是语句重复运行出错，一般情况下可以忽略该错误。"
                    merrCur.ErrOparate = vbIgnore
                Else
                    merrCur.ErrAdvice = "请检查所要求的父项数据的正确性。"
                    merrCur.ErrOparate = vbRetry
                End If
            Case 2298 '无法验证 (ZLHIS.XXX_FK_YY) - 未找到父项关键字
                'ORA-02298: 无法验证 (ZLHIS.XXX_FK_YY) - 未找到父项关键字
                If mobjSQL.PartSQL Like "ALTER TABLE * ADD CONSTRAINT * FOREIGN KEY *" Then
                    merrCur.ErrAdvice = "可能是数据错误，父项数据缺失。"
                    merrCur.ErrModiSQL = mobjSQL.SQL & " enable novalidate;"
                    merrCur.ErrOparateModi = vbRetry
                    merrCur.ErrOparate = vbRetry
                Else
                    merrCur.ErrAdvice = "请检查所要求的父项数据的正确性。"
                    merrCur.ErrOparate = vbRetry
                End If
            Case 957
                'ORA-00957:重复的列名
                If mobjSQL.PartSQL Like "ALTER TABLE * RENAME COLUMN * TO *" Then
                    merrCur.ErrAdvice = "可能是语句重复运行出错，一般情况下可以忽略该错误。"
                    merrCur.ErrOparate = vbIgnore
                ElseIf strSQL Like "CREATE TABLE*" Then
                    merrCur.ErrAdvice = "请检查SQL脚本的正确性。"
                    merrCur.ErrOparate = vbRetry
                End If
            Case 904
                'ORA-00904: 无效列名
                strSQL = GetFormatSQL(mobjSQL.SQL)
                If mobjSQL.PartSQL Like "ALTER TABLE * DROP COLUMN *" Then
                    merrCur.ErrAdvice = "可能是语句重复运行出错，一般情况下可以忽略该错误。"
                    merrCur.ErrOparate = vbIgnore
                Else
                    merrCur.ErrAdvice = "请先补充正确的表列字段后再重试。"
                    merrCur.ErrOparate = vbRetry
                End If
            Case 942
                strSQL = GetFormatSQL(mobjSQL.SQL)
                'ORA-00942: 表或视图不存在
                If mobjSQL.PartSQL Like "DROP TABLE*" Or mobjSQL.PartSQL Like "DROP VIEW*" Or mobjSQL.PartSQL Like "DROP MATERIALIZED VIEW*" Then
                    merrCur.ErrAdvice = "可能是语句重复运行出错，一般情况下可以忽略该错误。"
                    merrCur.ErrOparate = vbIgnore
                Else
                    merrCur.ErrAdvice = "请先补充创建对应的表或视图后再重试。"
                    merrCur.ErrOparate = vbRetry
                End If
            Case 1442
                'ORA-01442: 要修改为 NOT NULL 的列已经是 NOT NULL
                If mobjSQL.PartSQL Like "ALTER TABLE * MODIFY * CONSTRAINT * NOT NULL*" Then
                    merrCur.ErrAdvice = "可能是语句重复运行出错，一般情况下可以忽略该错误。"
                    merrCur.ErrOparate = vbIgnore
                Else
                    merrCur.ErrAdvice = "请检查修改列是否已经修改。"
                    merrCur.ErrOparate = vbRetry
                End If
            Case 8002
                'ORA-08002: 序列ZLRPTDATAS_ID.CURRVAL 尚未在此进程中定义
                If mobjSQL.PartSQL Like "INSERT INTO ZLRPT*" Then
                    merrCur.ErrAdvice = "可能是语句重复运行出错，一般情况下可以忽略该错误。"
                    merrCur.ErrOparate = vbIgnore
                Else
                    merrCur.ErrAdvice = "请检查之前的相关语句是否正确运行。"
                    merrCur.ErrOparate = vbRetry
                End If
                Call CheckAdjustSequnceVali(merrCur.ErrDesc)
            '-----------------------------------------------------------------------
            Case 1418
                strSQL = GetFormatSQL(mobjSQL.SQL)
                'ORA-01418: 指定的索引不存在
                Call CheckErr1418(strSQL)
            Case 4043, 4080, 2289
                'ORA-04043: 对象 XXX 不存在
                'ORA-04080: 触发器 'XXX' 不存在
                'ORA-02289: 序列（号）不存在
                If mobjSQL.PartSQL Like "DROP *" Then
                    merrCur.ErrAdvice = "可能是语句重复运行出错，一般情况下可以忽略该错误。"
                    merrCur.ErrOparate = vbIgnore
                Else
                    merrCur.ErrAdvice = "请先补充创建相应的对象后重试。"
                    merrCur.ErrOparate = vbRetry
                End If
            Case 1432, 1434, 1543, 1919, 1921, 1927, 1952, 2264, 2275, 2443, 4081, 12003, 12006
                'ORA-01432: 要删除的公用同义词不存在
                'ORA-01434: 要删除的隐含同义词不存在
                'ORA-01543: 表空间'XXX'已经存在
                'ORA-01919: 作用'XXX'不存在
                'ORA-01921: 作用名'XXX'与另一个用户名或作用名发生冲突
                'ORA-01927: 无法 REVOKE 您未授权的权限
                'ORA-01952: 系统权限未授予'ZLHIS'
                'ORA-02264: 名称已被一现有约束条件占用
                'ORA-02275: 此表中已经存在这样的引用约束条件
                'ORA-02443: 无法删除约束条件 - 不存在约束条件
                'ORA-04081: 触发器 'XXX' 已经存在
                'ORA-12003: 实体化视图(快照) "SYS"."TESTVIEW" 不存在
                'ORA-12006: 具有相同用户名的快照已经存在
                merrCur.ErrAdvice = "可能是语句重复运行出错，一般情况下可以忽略该错误。"
                merrCur.ErrOparate = vbIgnore
            '-----------------------------------------------------------------------
            Case 900, 907, 936
                'ORA-00900: 无效 SQL 语句
                'ORA-00907: 缺少右括号
                'ORA-00936: 缺少表达式
                merrCur.ErrAdvice = "请检查SQL脚本的正确性。"
                merrCur.ErrOparate = vbRetry
            Case 959
                'ORA-00959: 表空间'XXX'不存在
                merrCur.ErrAdvice = "请先补充创建对应的表空间后再重试。"
                merrCur.ErrOparate = vbRetry
            Case 1031
                'ORA-01031: 权限不足
                merrCur.ErrAdvice = "请先在数据库中授予当前用户相应角色权限后再重试。"
                merrCur.ErrOparate = vbRetry
            Case 1400, 1407, 2290
                'ORA-01400: 无法将 NULL 插入 ("ZLHIS"."XXX"."YYY")
                'ORA-01407: 无法更新 ("ZLHIS"."XXX"."YYY") 为 NULL
                'ORA-02290: 违反检查约束条件 (ZLHIS.地区_CK_缺省标志)
                merrCur.ErrAdvice = "请先检查并将错误的约束条件修正后再重试。"
                merrCur.ErrOparate = vbRetry
            Case 1408
                'ORA-01408: 此列列表已索引
                strSQL = GetFormatSQL(mobjSQL.SQL)
                merrCur.ErrAdvice = "请先检查并将错误的约束条件修正后再重试。"
                merrCur.ErrOparate = vbRetry
                Call CheckIndexCol(strSQL)
            Case 1401, 1438
                'ORA-01401: 插入的值对于列过大
                'ORA-01438: 值大于此列指定的允许精确度
                merrCur.ErrAdvice = "请先调整字段到合适的精度后再重试。"
                merrCur.ErrOparate = vbRetry
            Case 1439, 1440
                'ORA-01439: 要更改数据类型，则要修改的列必须为空 (empty)
                'ORA-01440: 要减小精确度或标度，则要修改的列必须为空 (empty)
                merrCur.ErrAdvice = "请先备份对应的表列数据，并清空后再重试。"
                merrCur.ErrOparate = vbRetry
            Case 1502
                'ORA-01502: index 'XXXX' or partition of such index is in unusable state
                'ORA-01502: 索引 'XXXX' 或这类索引的分区处于不可用状态
                merrCur.ErrAdvice = "请重建索引后重试。"
                merrCur.ErrOparate = vbRetry
                Call CheckAdjustIndex(merrCur.ErrDesc)
            Case 1775
                strSQL = GetFormatSQL(mobjSQL.SQL)
                'ORA-01775: 同义词的循环链
                merrCur.ErrAdvice = "同义词所对应的对象不存在，请查找是否是升级脚本删除对象时没有删除同义词所导致的，若是请忽略。"
                merrCur.ErrOparate = vbRetry
                Call CheckErr1775(strSQL)
            Case 2270
                strSQL = GetFormatSQL(mobjSQL.SQL)
                'ORA-02270: 此列列表的唯一或主键不匹配
                merrCur.ErrAdvice = "请对要创建外键的主表引用字段补充主键或唯一键后再重试。"
                merrCur.ErrOparate = vbRetry
                Call CheckErr2270(strSQL)
            Case 2273
                'ORA-02273: 此唯一/主键已被某些外部关键字引用
                merrCur.ErrAdvice = "请先删除从表的外键约束引用后再重试。"
                merrCur.ErrOparate = vbRetry
            Case 2091, 2292 '2091是延迟约束
                'ORA-02292: 违反完整约束条件 (ZLHIS.XXX_FK_YYY) - 已找到子记录日志
                merrCur.ErrAdvice = "请先删除从表的关联数据后再重试。"
                merrCur.ErrOparate = vbRetry
            Case 2303
                'ORA-02303: 无法使用类型或表的相关性来删除或取代一个类型
                merrCur.ErrAdvice = "请先取消引用该类型的相关表列或类型之后再重试。"
                merrCur.ErrOparate = vbRetry
            Case 2299, 2437
                'ORA-02437: 无法验证 (ZLHIS.XXX_PK) - 违反主键
                'ORA-02299: 无法验证 (ZLHIS.XXX_UQ_YYY) - 未找到重复关键字
                merrCur.ErrAdvice = "请对表中对应字段的数据进行重复检查处理后再重试。"
                merrCur.ErrOparate = vbRetry
            Case 6575
                'ORA-06575: 程序包或函数 TTT 处于无效状态
                'ORA-06575: Package or function TTT is in an invalid state
                merrCur.ErrAdvice = "请先正确编译对应的函数或过程后再重试。"
                merrCur.ErrOparate = vbRetry
                Call CheckAdjustProcedure(merrCur.ErrDesc)
            Case 6576
                'ORA-06576: 不是有效的函数或过程名
                merrCur.ErrAdvice = "请先正确创建对应的函数或过程后再重试。"
                merrCur.ErrOparate = vbRetry
            Case 12899 'ORA-12899: 列 "ZLHIS"."体检任务人员"."所属团体" 的值太大 (实际值: 62, 最大值: 60)
                'ORA-12899: value too large for column "SYSTEM"."STUDENTINFO"."SNAME" (actual: 78, maximum: 30)
                merrCur.ErrAdvice = "请先调整字段到合适的精度后再重试。"
                merrCur.ErrOparate = vbRetry
                Call CheckAdjust12899(merrCur.ErrDesc)
            Case 19001
                strSQL = GetFormatSQL(mobjSQL.SQL)
                'ORA-19001指定的存储选项无效
                If strSQL Like "CREATE TABLE*STORE AS SECUREFILE BINARY XML*" Then
                    If GetOracleVersion(True, True) < 11 Then 'racle版本低于11g时
                        merrCur.ErrAdvice = "Oracle版本低于11G不支持该存储选项，后续有低版本的结构创建语句，可以忽略。"
                        merrCur.ErrOparate = vbIgnore
                    Else
                        merrCur.ErrAdvice = "请检查SQL是否书写正确或当前Oracle版本是否支持该存储选项。"
                        merrCur.ErrOparate = vbRetry
                    End If
                Else
                    merrCur.ErrAdvice = "请检查SQL是否书写正确或当前Oracle版本是否支持该存储选项。"
                    merrCur.ErrOparate = vbRetry
                End If
            Case 22858 'ORA-22858: 数据类型的更改无效;一般将普通类型修改为大对象
                strSQL = GetFormatSQL(mobjSQL.SQL)
                Call CheckTabChangeCol(strSQL)
            Case 22859 'ORA-22859: 无效的列修改;一般将大对象修改为普通类型
                strSQL = GetFormatSQL(mobjSQL.SQL)
                Call CheckTabChangeCol(strSQL)
            Case 23292 'ORA-23292: 约束条件不存在
                strSQL = GetFormatSQL(mobjSQL.SQL)
                Call CheckErr23292(strSQL)
            Case Else
                merrCur.ErrAdvice = "请仔细检查，并作必要的处理后尝试重试操作"
                merrCur.ErrOparate = vbRetry
        End Select
    End If
End Sub

Private Sub CheckTabChangeCol(ByVal strSQL As String)
'功能：检查表添加列的内容与数据库是否一致
'参数：strSQL=已格式化为标准大写的SQL语句
'返回：
    Dim rsTemp As New ADODB.Recordset
    Dim strOwner As String, strName As String
    Dim strCol As String, arrCol As Variant
    Dim strType As String, intMatch As Integer
    Dim intLen As Integer, intDigit As Integer
    Dim strError As String, i As Long
    Dim strModifySQL As String
    Dim blnModify As Boolean
    
    If Not (strSQL Like "ALTER TABLE * ADD*" Or strSQL Like "ALTER TABLE * MODIFY*") Then Exit Sub

    '表名
    strName = Split(Mid(strSQL, InStr(strSQL, "ALTER TABLE ") + Len("ALTER TABLE ")), " ")(0)
    If InStr(strName, ".") > 0 Then
        strOwner = Split(strName, ".")(0)
        strName = Split(strName, ".")(1)
    End If

    '取出SQL语句中的列定义
    If strSQL Like "ALTER TABLE * ADD*" Then
        strSQL = Trim(Mid(strSQL, InStr(strSQL, strName & " ADD") + Len(strName & " ADD")))
        blnModify = False
    Else
        strSQL = Trim(Mid(strSQL, InStr(strSQL, strName & " MODIFY") + Len(strName & " MODIFY")))
        blnModify = True
    End If
    If Left(strSQL, 1) = "(" Then
        intMatch = 1
        For i = 2 To Len(strSQL)
            If Mid(strSQL, i, 1) = "(" Then
                intMatch = intMatch + 1
            ElseIf Mid(strSQL, i, 1) = ")" Then
                intMatch = intMatch - 1
                If intMatch = 0 Then Exit For
            End If
            If Mid(strSQL, i, 1) = "," And intMatch = 1 Then
                strCol = strCol & "|" '不在外层括号中的",",如Number(16,5)
            Else
                strCol = strCol & Mid(strSQL, i, 1)
            End If
        Next
    Else
        strCol = strSQL
    End If
    arrCol = Split(strCol, "|")

    '将SQL中的列定义与数据库中的进行比较
    On Error Resume Next
    strSQL = "Select Column_Name,Data_Type,Data_Length,Data_Precision,Data_Scale From ALL_Tab_Columns" & _
        " Where OWNER=" & IIf(strOwner = "", "User", "'" & strOwner & "'") & " And Table_Name='" & strName & "'"
    rsTemp.CursorLocation = adUseClient
    rsTemp.Open strSQL, mcnThis, adOpenKeyset
    For i = 0 To UBound(arrCol)
        arrCol(i) = Trim(arrCol(i))

        strCol = Left(arrCol(i), InStr(arrCol(i), " ") - 1) '名称 Number ( 16, 5) Not Null Default 1.23
        strType = Mid(arrCol(i), InStr(arrCol(i), " ") + 1)

        rsTemp.Filter = "Column_Name='" & strCol & "'"
        If rsTemp.EOF Then
            strModifySQL = strModifySQL & IIf(strModifySQL = "", "", vbNewLine) & "ALter Table " & IIf(strOwner = "", mstrUser, strOwner) & "." & strName & " Add " & arrCol(i) & ";"
            strError = strError & "," & strCol
        Else
            If strType Like rsTemp!DATA_TYPE & "*" Then
                If rsTemp!DATA_TYPE = "NUMBER" Then
                    If InStr(strType, ",") > 0 Then 'Number(16,5)
                        intLen = Val(Split(Replace(Split(Split(strType, "(")(1), ")")(0), " ", ""), ",")(0))
                        intDigit = Val(Split(Replace(Split(Split(strType, "(")(1), ")")(0), " ", ""), ",")(1))
                    Else 'Number(18)
                        intLen = Val(Replace(Split(Split(strType, "(")(1), ")")(0), " ", ""))
                        intDigit = 0
                    End If
                    If rsTemp!Data_Precision < intLen Or rsTemp!Data_Scale < intDigit Then
                        strError = strError & "," & strCol
                        strModifySQL = strModifySQL & IIf(strModifySQL = "", "", vbNewLine) & "ALter Table " & IIf(strOwner = "", mstrUser, strOwner) & "." & strName & " Modify " & arrCol(i) & ";"
                    End If
                ElseIf rsTemp!DATA_TYPE = "VARCHAR2" Then 'Varchar2(50)
                    intLen = Val(Replace(Split(Split(strType, "(")(1), ")")(0), " ", ""))
                    If rsTemp!Data_Length < intLen Then
                        strError = strError & "," & strCol
                        strModifySQL = strModifySQL & IIf(strModifySQL = "", "", vbNewLine) & "ALter Table " & IIf(strOwner = "", mstrUser, strOwner) & "." & strName & " Modify " & arrCol(i) & ";"
                    End If
                ElseIf rsTemp!DATA_TYPE = strType Then
                '数据类型正常
                End If
            Else
                strError = strError & "," & strCol
            End If
        End If
    Next
    strError = Mid(strError, 2)

    If strError <> "" Then
        merrCur.ErrAdvice = "表中缺少SQL语句要添加的其他列，或已经存在的列类型与SQL语句不符，请手工处理后重试。"
        merrCur.ErrOparate = vbRetry
        merrCur.ErrModiSQL = strModifySQL
        merrCur.ErrOparateModi = vbRetry
    Else
        merrCur.ErrAdvice = "表中已经存在要添加的列或者列类型已经修改，请确认后忽略。"
        merrCur.ErrOparate = vbIgnore
        merrCur.ErrModiSQL = ""
        merrCur.ErrOparateModi = vbIgnore
    End If
End Sub

Private Sub CheckErr1775(ByVal strSQL As String)
'ORA-01775: 同义词的循环链
'功能：管理工具删除表"zlPDASynch"、"zlStreamTabs"
'参数：strSQL=已格式化为标准大写的SQL语句
'返回：建议提示内容(strAdvice)及缺省操作按钮值(intAdvice)
'管理工具版本9.41.0
'Drop Table zlPDASynch;
'drop table zlStreamTabs;
    If strSQL Like "* ZLPDASYNCH*" Or strSQL Like "* ZLSTREAMTABS*" Then
        merrCur.ErrAdvice = "该对象已经在管理工具版本9.41.0中删除。"
        merrCur.ErrOparate = vbIgnore
    End If
End Sub

Private Sub CheckErr1418(ByVal strSQL As String)
'ORA-01418: 指定的索引不存在
'功能：删除索引自动忽略，重命名索引，查看索引是否存在，存在则自动忽略
'参数：strSQL=已格式化为标准大写的SQL语句
    Dim strIndexName As String, arrTmp As Variant
    merrCur.ErrAdvice = "请先补充创建相应的对象后重试。"
    merrCur.ErrOparate = vbRetry
    If strSQL Like "DROP *" Then
        merrCur.ErrAdvice = "可能是语句重复运行出错，一般情况下可以忽略该错误。"
        merrCur.ErrOparate = vbIgnore
    ElseIf strSQL Like "ALTER INDEX * RENAME TO *" Then
        arrTmp = Split(strSQL, "RENAME TO")
        If UBound(arrTmp) < 1 Then Exit Sub
        strIndexName = UCase(Trim(Split(Trim(arrTmp(1)), " ")(0)))
        If strIndexName = "" Then Exit Sub
        If ObjectExists(UCase(mstrUser), strIndexName, "INDEX") Then
            merrCur.ErrAdvice = "可能是语句重复运行出错，一般情况下可以忽略该错误。"
            merrCur.ErrOparate = vbIgnore
        End If
    End If
End Sub

Private Sub CheckErr23292(ByVal strSQL As String)
'ORA -23292: 约束条件不存在
'功能：删除索引自动忽略，重命名索引，查看索引是否存在，存在则自动忽略
'参数：strSQL=已格式化为标准大写的SQL语句
    Dim strConName As String, arrTmp As Variant, strTableName As String
    Dim arrTmp1 As Variant
    Dim strOwner As String
    
    merrCur.ErrAdvice = "请先补充创建相应的对象后重试。"
    merrCur.ErrOparate = vbRetry
    If strSQL Like "ALTER TABLE * RENAME CONSTRAINT * TO *" Then
        arrTmp = Split(strSQL, " TO ")
        If UBound(arrTmp) < 1 Then Exit Sub
        strTableName = arrTmp(0)
        strTableName = Trim(Mid(Split(strTableName, "RENAME")(0), Len("ALTER TABLE ")))
        If strTableName = "" Then Exit Sub
        arrTmp1 = Split(strTableName, ".")
        If UBound(arrTmp1) = 1 Then
            strOwner = arrTmp1(0)
            strTableName = arrTmp1(1)
        Else
            strTableName = arrTmp1(0)
        End If
        If strTableName = "" Then Exit Sub
        strConName = UCase(Trim(Split(Trim(arrTmp(1)), " ")(0)))
        If strConName = "" Then Exit Sub
        arrTmp1 = Split(strConName, ".")
        If UBound(arrTmp1) = 1 Then
            strOwner = arrTmp1(0)
            strConName = arrTmp1(1)
        Else
            strConName = arrTmp1(0)
        End If
        If strConName = "" Then Exit Sub
        If strOwner = "" Then strOwner = mstrUser
        If ObjectExists(UCase(strOwner), strConName, "CONSTRAINT", strTableName) Then
            merrCur.ErrAdvice = "可能是语句重复运行出错，一般情况下可以忽略该错误。"
            merrCur.ErrOparate = vbIgnore
        End If
    End If
End Sub

Private Sub CheckErr2270(ByVal strSQL As String)
'ORA-02270: 此列列表的唯一或主键不匹配
'功能：处理由于存在主键/唯一键，但是不存在主键/唯一键索引导致的错误
'参数：strSQL=已格式化为标准大写的SQL语句
'返回：建议提示内容(strAdvice)及缺省操作按钮值(intAdvice)
    Dim strTable As String, strRTable As String, strCols As String, strRCols As String
    Dim strColsInfo As String, strRColsInfo As String, strTmp As String, strPreCon As String
    Dim strOwner As String, strROwner As String
    Dim rsRTable As New ADODB.Recordset, rsTable As New ADODB.Recordset
    Dim cllColInfo As Collection
    Dim i As Long, arrTmp As Variant
    Dim strModifySQL As String
    
    '该错误有以下原因：
    '1、引用表字段未创建主键或唯一键
    '2、引用表字段创建的主键或唯一键字段的类型，个数与要创建外键的字段存在差异
    '3、引用表的主键唯一键没有索引
    'Alter Table 保险支付比例 Add Constraint 保险支付比例_FK_年龄段 Foreign Key (险类,中心,在职,年龄段) References 保险年龄段(险类,中心,在职,年龄段) On Delete Cascade;
    If Not strSQL Like "ALTER TABLE * ADD CONSTRAINT * FOREIGN KEY * REFERENCES *" Then Exit Sub
    '解析SQL中的信息
    arrTmp = Split(strSQL, "ADD CONSTRAINT")
    strTable = Trim(Split(arrTmp(0), "ALTER TABLE")(1))
    arrTmp = Split(Split(arrTmp(1), "FOREIGN KEY")(1), "REFERENCES")
    strCols = UCase(Trim(Replace(Replace(Replace(arrTmp(0), "(", ""), ")", ""), " ", ""))) '去除两侧括号以及其中的空格
    arrTmp = Split(Split(arrTmp(1), ")")(0), "(") '以括号做分割符
    strRTable = Trim(arrTmp(0))
    strRCols = UCase(Trim(Replace(arrTmp(1), " ", "")))  '去除其中的空格
    
    If InStr(strTable, ".") > 0 Then
        strOwner = UCase(Split(strTable, ".")(0))
        strTable = UCase(Split(strTable, ".")(1))
    End If
    If InStr(strRTable, ".") > 0 Then
        strROwner = UCase(Split(strRTable, ".")(0))
        strRTable = UCase(Split(strRTable, ".")(1))
    End If
    
    '获取引用表的主键唯一键信息以及其字段信息以及对应索引等
    strSQL = "Select a.Constraint_Name, a.Column_Name, a.Position, b.Data_Type, b.Data_Length, b.Data_Precision, b.Data_Scale," & vbNewLine & _
                    "       c.Constraint_Type, c.Index_Name" & vbNewLine & _
                    "From (Select a.Owner, a.Constraint_Name, a.Table_Name, a.Column_Name, Nvl(a.Position, 1) Position" & vbNewLine & _
                    "       From All_Cons_Columns A" & vbNewLine & _
                    "       Where a.Table_Name = '" & strRTable & "') A," & vbNewLine & _
                    "     (Select a.Owner, a.Table_Name, a.Column_Name, a.Data_Type, a.Data_Length, a.Data_Precision, a.Data_Scale" & vbNewLine & _
                    "       From All_Tab_Columns A" & vbNewLine & _
                    "       Where a.Table_Name = '" & strRTable & "') B," & vbNewLine & _
                    "     (Select a.Owner, a.Constraint_Name, a.Constraint_Type, a.Index_Name" & vbNewLine & _
                    "       From All_Constraints A" & vbNewLine & _
                    "       Where a.Table_Name = '" & strRTable & "' And a.Constraint_Type In ('P', 'U')) C" & vbNewLine & _
                    "Where a.Owner = b.Owner And a.Column_Name = b.Column_Name And a.Constraint_Name = c.Constraint_Name And" & vbNewLine & _
                    "      a.Owner = c.Owner And a.Owner " & IIf(strROwner <> "", "=" & strROwner, IIf(mstrUser = "ZLTOOLS", "= User", " In (User, 'ZLTOOLS')")) & vbNewLine & _
                    " Order By a.Constraint_Name,a.Position"
    Set rsRTable = gclsBase.OpenSQLRecord(mcnThis, strSQL, "错误处理-获取主表信息")
    '1、引用表字段未创建主键或唯一键
    If rsRTable.RecordCount = 0 Then Exit Sub
    '2、引用表字段创建的主键或唯一键字段的类型，个数与要创建外键的字段存在差异
    strTmp = ""
    Set cllColInfo = New Collection
    With rsRTable
        Do While Not .EOF
            If strPreCon <> !Constraint_Name Then
                If strTmp <> "" Then cllColInfo.Add strPreCon & "=" & strRColsInfo, strTmp
                strPreCon = !Constraint_Name
                strRColsInfo = !DATA_TYPE & "," & !Data_Length & "," & !Data_Precision & "," & !Data_Scale
                strTmp = !Column_Name
            Else
                strRColsInfo = strRColsInfo & "|" & !DATA_TYPE & "," & !Data_Length & "," & !Data_Precision & "," & !Data_Scale
                strTmp = strTmp & "," & !Column_Name
            End If
            .MoveNext
        Loop
        If strTmp <> "" Then
            cllColInfo.Add strPreCon & "=" & strRColsInfo, strTmp
        End If
    End With
    strTmp = ""
    '检查引用字段上是否有主键或唯一键约束
    On Error Resume Next
    strTmp = cllColInfo(strRCols)
    If err.Number <> 0 Then
        '没有获取到对应约束
        err.Clear: Exit Sub
    End If
    On Error GoTo 0
    '获取到约束，则对比字段类型
    strSQL = "Select a.Column_Name, a.Data_Type, a.Data_Length, a.Data_Precision, a.Data_Scale" & vbNewLine & _
                    "From All_Tab_Columns A" & vbNewLine & _
                    "Where a.Owner " & IIf(strOwner <> "", "=" & strOwner, IIf(mstrUser = "ZLTOOLS", "= User", " In (User, 'ZLTOOLS')")) & " And a.Table_Name = '" & strTable & "' And a.Column_Name In ('" & Replace(strCols, ",", "','") & "')"
    Set rsTable = gclsBase.OpenSQLRecord(mcnThis, strSQL, "错误处理-获取从表信息")
    arrTmp = Split(strCols, ",")
    For i = LBound(arrTmp) To UBound(arrTmp)
        rsTable.Filter = "Column_Name='" & arrTmp(i) & "'"
        If Not rsTable.EOF Then
            strColsInfo = strColsInfo & IIf(strColsInfo = "", "", "|") & rsTable!DATA_TYPE & "," & rsTable!Data_Length & "," & rsTable!Data_Precision & "," & rsTable!Data_Scale
        End If
    Next
    arrTmp = Split(strTmp, "=")
    strTmp = arrTmp(0)
    strRColsInfo = arrTmp(1)
    If strColsInfo <> strRColsInfo Then
        Exit Sub '类型不匹配
    End If
    rsRTable.Filter = "Constraint_Name='" & strTmp & "'"
    If rsRTable!Index_Name & "" = "" Then
        '优先判断唯一约束，主键约束同名索引是否存在
        strSQL = "Select a.Status" & vbNewLine & _
                        "From All_Indexes A" & vbNewLine & _
                        "Where a.Table_Owner " & IIf(strROwner <> "", "=" & strROwner, IIf(mstrUser = "ZLTOOLS", "= User", " In (User, 'ZLTOOLS')")) & " And a.Uniqueness = 'UNIQUE' And a.Table_Name = [1] And a.Index_Name =[2]"
        Set rsTable = gclsBase.OpenSQLRecord(mcnThis, strSQL, "错误处理-获取从表信息", strRTable, strTmp)
        If rsTable.RecordCount = 0 Then
            strModifySQL = strModifySQL & IIf(strModifySQL = "", "", vbNewLine) & "Create Index " & strTmp & " On " & strRTable & "(" & strRCols & ")   ;"
        Else
            If rsTable!Status <> "VALID" Then
                strModifySQL = strModifySQL & IIf(strModifySQL = "", "", vbNewLine) & "ALter Index " & strTmp & "  Rebuild;"
            End If
        End If
        strModifySQL = strModifySQL & IIf(strModifySQL = "", "", vbNewLine) & "ALter  Table  " & strRTable & " Modify Constraint " & strTmp & " Using Index " & strTmp & ";"
        merrCur.ErrAdvice = "主表主键或唯一键缺少索引，请创建后重试。"
        merrCur.ErrOparate = vbRetry
        merrCur.ErrOparateModi = vbRetry
        merrCur.ErrModiSQL = strModifySQL
    End If
End Sub

Private Function CheckUniqueKeyCol(ByVal strSQL As String, Optional ByVal blnCreateTab As Boolean) As String
'功能：检查SQL主键或唯一约束与数据库中是否一致
'参数：strSQL=已格式化为标准大写的SQL语句
'      blnCreateTab=是否创建表调用，该调用不设置错误SQL与错误建议，只返回修正SQL
'返回：建议提示内容(strAdvice)及缺省操作按钮值(intAdvice)
    Dim rsTemp As New ADODB.Recordset
    Dim strCol As String, strOwner As String
    Dim strTab As String, strName As String
    Dim strType As String, strTmp As String
    Dim strPreConName As String, strLikeConName As String, strSameConName As String
    Dim blnLike As Boolean, strSameOwner As String, strLikeOwner As String
    Dim strSameConType As String
    
    
    If Not (strSQL Like "ALTER TABLE * ADD CONSTRAINT * PRIMARY KEY*" Or _
        strSQL Like "ALTER TABLE * ADD CONSTRAINT * UNIQUE*") Then Exit Function

    strTab = Split(Mid(strSQL, InStr(strSQL, "ALTER TABLE ") + Len("ALTER TABLE ")), " ")(0)
    strName = Split(Mid(strSQL, InStr(strSQL, "ADD CONSTRAINT ") + Len("ADD CONSTRAINT ")), " ")(0)
    If strSQL Like "*PRIMARY KEY*" Then
        strCol = Mid(strSQL, InStr(strSQL, "PRIMARY KEY") + Len("PRIMARY KEY"))
        strType = "P"
    Else
        strCol = Mid(strSQL, InStr(strSQL, "UNIQUE") + Len("UNIQUE"))
        strType = "U"
    End If
    strCol = Mid(strCol, InStr(strCol, "(") + 1)
    strCol = Left(strCol, InStr(strCol, ")") - 1)
    strCol = Replace(Trim(strCol), " ", "")

    On Error Resume Next
    If InStr(strTab, ".") > 0 Then
        strOwner = Split(strTab, ".")(0)
        strTab = Split(strTab, ".")(1)
    Else
        strOwner = mstrUser
    End If
    strSQL = "Select Column_Name" & vbNewLine & _
            "From All_Cons_Columns a" & vbNewLine & _
            "Where (a.Owner, a.Table_Name, a.Constraint_Name) In" & vbNewLine & _
            "      (Select b.Owner, b.Table_Name, b.Constraint_Name" & vbNewLine & _
            "       From All_Constraints b" & vbNewLine & _
            "       Where b.Owner = " & IIf(strOwner = "", "User", "'" & strOwner & "'") & " And b.Table_Name = '" & strTab & "' And b.Constraint_Name = '" & strName & "' And b.Constraint_Type = '" & strType & "')" & vbNewLine & _
            "Order By Position"
    Set rsTemp = gclsBase.OpenSQLRecord(mcnThis, strSQL, App.Title)
    strSQL = ""
    Do While Not rsTemp.EOF
        strSQL = strSQL & "," & rsTemp!Column_Name
        rsTemp.MoveNext
    Loop
    strSQL = Mid(strSQL, 2)
    '可能是其他的约束名或约束类型
    If strSQL = "" Then
        strTmp = Replace(UCase(strCol), ",", "','")
        strSQL = "Select a.Owner, a.Constraint_Name, a.Column_Name, b.Constraint_Type" & vbNewLine & _
                "From All_Cons_Columns a, All_Constraints b" & vbNewLine & _
                "Where b.Owner = " & IIf(strOwner = "", "User", "'" & strOwner & "'") & " And b.Table_Name = '" & strTab & "' And a.Owner = b.Owner And a.Table_Name = b.Table_Name And" & vbNewLine & _
                "      a.Constraint_Name = b.Constraint_Name And a.Column_Name In ('" & strTmp & "') And b.Constraint_Type In('P','U')" & vbNewLine & _
                "Order By a.Position"
        Set rsTemp = gclsBase.OpenSQLRecord(mcnThis, strSQL, App.Title)
        strSQL = "": strTmp = "": blnLike = True
        If Not rsTemp Is Nothing Then
            '查看同列索引
            Do While Not rsTemp.EOF
                If strPreConName <> rsTemp!Constraint_Name Then
                    If strPreConName <> "" Then
                        If strTmp = strCol Then
                            strSameConName = strPreConName
                            strSQL = strTmp
                            Exit Do
                        ElseIf strLikeConName <> "" And blnLike Then
                            strSQL = strTmp
                            blnLike = False
                        End If
                    End If
                    If strCol Like rsTemp!Column_Name & ",*" And blnLike Then
                        strLikeConName = rsTemp!Constraint_Name
                        strLikeOwner = rsTemp!Owner
                    End If
                    strPreConName = rsTemp!Constraint_Name
                    strSameOwner = rsTemp!Owner
                    strSameConType = rsTemp!constraint_type
                    strTmp = rsTemp!Column_Name
                Else
                    strTmp = strTmp & "," & rsTemp!Column_Name
                End If
                rsTemp.MoveNext
            Loop
            If strPreConName <> "" Then
                If strTmp = strCol Then
                    strSameConName = strPreConName
                    strSQL = strTmp
                ElseIf strLikeConName <> "" And blnLike Then
                    strSQL = strTmp
                    blnLike = False
                End If
            End If
        End If
    End If
    
    If strSQL <> strCol Then
        If strLikeConName <> "" Then
            If blnCreateTab Then
                 CheckUniqueKeyCol = "ALTER TABLE " & strOwner & "." & strTab & " Drop CONSTRAINT " & strLikeConName & ";"
            Else
                merrCur.ErrAdvice = "该索引的列上已经存在约束""" & strLikeConName & """,但是两者约束列或属于类型有差异，请手工处理后重试。"
                merrCur.ErrOparate = vbRetry
                merrCur.ErrModiSQL = "ALTER TABLE " & strLikeOwner & "." & strTab & " Drop CONSTRAINT " & strLikeConName & ";"
                merrCur.ErrOparateModi = vbRetry
            End If
        Else
            If blnCreateTab Then
                CheckUniqueKeyCol = "ALTER TABLE " & strOwner & "." & strTab & " Drop CONSTRAINT " & strName & ";"
            Else
                merrCur.ErrAdvice = "已经存在的主键或唯一约束的字段或字段顺序与SQL语句不符，请手工处理后重试。"
                merrCur.ErrOparate = vbRetry
                merrCur.ErrModiSQL = "ALTER TABLE " & strOwner & "." & strTab & " Drop CONSTRAINT " & strName & ";"
                merrCur.ErrOparateModi = vbRetry
            End If
        End If
    ElseIf strSameConName <> "" Then
        If strSameConType <> strType Then
            If blnCreateTab Then
                CheckUniqueKeyCol = "ALTER TABLE " & strOwner & "." & strTab & " Drop CONSTRAINT " & strSameConName & ";"
            Else
                merrCur.ErrOparate = vbRetry
                merrCur.ErrAdvice = "已经存在相同约束列上的约束，但是约束类型不同的约束" & strSameConName & "，请手工处理后重试。"
                merrCur.ErrModiSQL = "ALTER TABLE " & strOwner & "." & strTab & " Drop CONSTRAINT " & strSameConName & ";"
                merrCur.ErrOparateModi = vbIgnore
            End If
        Else
            If blnCreateTab Then
                CheckUniqueKeyCol = "alter table " & strOwner & "." & strTab & " rename constraint  " & strSameConName & " to " & strName & " ;" & vbNewLine & _
                                    "alter index " & strSameConName & " rename to " & strName & " ;"
            Else
                merrCur.ErrOparate = vbRetry
                merrCur.ErrAdvice = "已经存相同列上的约束，但是名称与SQL不符，请手工处理后重试。"
                merrCur.ErrModiSQL = "alter table " & strOwner & "." & strTab & " rename constraint  " & strSameConName & " to " & strName & " ;" & vbNewLine & _
                                    "alter index " & strSameConName & " rename to " & strName & " ;"
                merrCur.ErrOparateModi = vbIgnore
            End If
        End If
    Else
        merrCur.ErrAdvice = "已经存相同约束，可以忽略该错误。"
        merrCur.ErrOparate = vbIgnore
        merrCur.ErrModiSQL = ""
        merrCur.ErrOparateModi = vbIgnore
    End If
End Function

Private Function CheckIndexCol(ByVal strSQL As String, Optional ByVal blnCreateTab As Boolean) As String
'功能：检查SQL主键或唯一约束与数据库中是否一致
'参数：strSQL=已格式化为标准大写的SQL语句
'      blnCreateTab=是否创建表调用，该调用不设置错误SQL与错误建议，只返回修正SQL
'返回：建议提示内容(strAdvice)及缺省操作按钮值(intAdvice)
    Dim rsTemp As ADODB.Recordset
    Dim strCol As String, strOwner As String
    Dim strTab As String, strName As String
    Dim strTmp As String, strPreIndName As String, strLikeIndName As String, strSameIndName As String
    Dim blnLike As Boolean, strSameOwner As String, strLikeOwner As String
    
    If Not strSQL Like "CREATE INDEX * ON *" Then Exit Function
    strName = Split(Mid(strSQL, InStr(strSQL, "CREATE INDEX ") + Len("CREATE INDEX ")), " ")(0)
    strCol = Split(Mid(strSQL, InStr(strSQL, "ON ") + Len("ON ")), ")")(0)
    strTab = Split(strCol, "(")(0)
    strCol = Split(strCol, "(")(1)
    strCol = Replace(Trim(strCol), " ", "")

    On Error Resume Next
    If InStr(strTab, ".") > 0 Then
        strOwner = Split(strTab, ".")(0)
        strTab = Split(strTab, ".")(1)
    Else
        strOwner = mstrUser
    End If
    strSQL = "Select a.Column_Name" & vbNewLine & _
            "From All_Ind_Columns a" & vbNewLine & _
            "Where a.Table_Owner = " & IIf(strOwner = "", "User", "'" & strOwner & "'") & " And a.Table_Name = '" & strTab & "' And a.Index_Name = '" & strName & "'" & vbNewLine & _
            "Order By a.Column_Position"
    Set rsTemp = gclsBase.OpenSQLRecord(mcnThis, strSQL, App.Title)
    strSQL = ""
    If Not rsTemp Is Nothing Then
        Do While Not rsTemp.EOF
            strSQL = strSQL & "," & rsTemp!Column_Name
            rsTemp.MoveNext
        Loop
    End If
    strSQL = Mid(strSQL, 2)
    '可能是其他的索引名
    If strSQL = "" Then
        strTmp = Replace(UCase(strCol), ",", "','")
        strSQL = "Select a.Index_Name, a.Column_Name ,a.Index_Owner" & vbNewLine & _
                "From All_Ind_Columns a" & vbNewLine & _
                "Where (a.Index_Name, a.Index_Owner) In" & vbNewLine & _
                "      (Select Distinct b.Index_Name, b.Index_Owner" & vbNewLine & _
                "       From All_Ind_Columns b" & vbNewLine & _
                "       Where b.TABLE_OWNER = " & IIf(strOwner = "", "User", "'" & strOwner & "'") & " And b.Column_Name In ('" & strTmp & "') And b.Table_Name = '" & strTab & "')" & vbNewLine & _
                "Order By a.Index_Name, a.Column_Position"
        Set rsTemp = gclsBase.OpenSQLRecord(mcnThis, strSQL, App.Title)
        strSQL = "": strTmp = "": blnLike = True
        If Not rsTemp Is Nothing Then
            '查看同列索引
            Do While Not rsTemp.EOF
                If strPreIndName <> rsTemp!Index_Name Then
                    If strPreIndName <> "" Then
                        If strTmp = strCol Then
                            strSameIndName = strPreIndName
                            strSQL = strTmp
                            Exit Do
                        ElseIf strLikeIndName <> "" And blnLike Then
                            strSQL = strTmp
                            blnLike = False
                        End If
                    End If
                    If strCol Like rsTemp!Column_Name & ",*" And blnLike Then
                        strLikeIndName = rsTemp!Index_Name
                        strLikeOwner = rsTemp!Index_Owner
                    End If
                    strPreIndName = rsTemp!Index_Name
                    strSameOwner = rsTemp!Index_Owner
                    strTmp = rsTemp!Column_Name
                Else
                    strTmp = strTmp & "," & rsTemp!Column_Name
                End If
                rsTemp.MoveNext
            Loop
            If strPreIndName <> "" Then
                If strTmp = strCol Then
                    strSameIndName = strPreIndName
                    strSQL = strTmp
                ElseIf strLikeIndName <> "" And blnLike Then
                    strSQL = strTmp
                    blnLike = False
                End If
            End If
        End If
    End If
    If strSQL <> strCol Then
        If strLikeIndName <> "" Then
            merrCur.ErrAdvice = "该索引的列上已经存在索引""" & strLikeIndName & """,但是两者索引列有差异，请手工处理后重试。"
            merrCur.ErrOparate = vbRetry
            merrCur.ErrModiSQL = "Drop Index " & strLikeOwner & "." & strLikeIndName & ";"
            merrCur.ErrOparateModi = vbRetry
        Else
            merrCur.ErrAdvice = "已经存在的主键或唯一约束的字段或字段顺序与SQL语句不符，请手工处理后重试。"
            merrCur.ErrOparate = vbRetry
            merrCur.ErrModiSQL = "Drop Index " & strOwner & "." & strName & ";"
            merrCur.ErrOparateModi = vbRetry
        End If
    ElseIf strSameIndName <> "" Then
        merrCur.ErrAdvice = "已经存相同列上的索引名称与SQL不符，请手工处理后重试。"
        merrCur.ErrOparate = vbRetry
        merrCur.ErrModiSQL = "Alter Index " & strSameOwner & "." & strSameIndName & " Rename to " & strName & ";"
        merrCur.ErrOparateModi = vbIgnore
    Else
        merrCur.ErrAdvice = "已经存相同索引，可以忽略该错误。"
        merrCur.ErrOparate = vbIgnore
        merrCur.ErrModiSQL = ""
        merrCur.ErrOparateModi = vbIgnore
    End If
End Function

Private Function CheckCreateTabCol(ByVal strSQL As String, ByVal strOwner As String, ByVal strName As String) As String
'功能：检查SQL创建表的列与当前数据库中的是否一致
'参数：strSQL=已格式化为标准大写的SQL语句
'返回：数据库中比SQL中要少的字段
    Dim rsTemp As New ADODB.Recordset
    Dim intMatch As Integer, i As Long
    Dim arrCol As Variant, strError As String
    Dim strCol As String, strType As String
    Dim intLen As Integer, intDigit As Integer
    Dim strModifySQL As String
    Dim strTmpSQL As String, strConsSQL As String
    
    '取出SQL语句中的列定义
    intMatch = 1
    For i = InStr(strSQL, "(") + 1 To Len(strSQL)
        If Mid(strSQL, i, 1) = "(" Then
            intMatch = intMatch + 1
        ElseIf Mid(strSQL, i, 1) = ")" Then
            intMatch = intMatch - 1
            If intMatch = 0 Then Exit For
        End If
        If Mid(strSQL, i, 1) = "," And intMatch = 1 Then
            strCol = strCol & "|" '不在外层括号中的",",如Number(16,5)
        Else
            strCol = strCol & Mid(strSQL, i, 1)
        End If
    Next
    arrCol = Split(strCol, "|")

    '将SQL中的列定义与数据库中的进行比较
    On Error Resume Next
    strSQL = "Select Column_Name,Data_Type,Data_Length,Data_Precision,Data_Scale From ALL_Tab_Columns" & _
        " Where OWNER=" & IIf(strOwner = "", "User", "'" & strOwner & "'") & " And Table_Name='" & strName & "'"
    rsTemp.CursorLocation = adUseClient
    rsTemp.Open strSQL, mcnThis, adOpenKeyset
    For i = 0 To UBound(arrCol)
        arrCol(i) = Trim(arrCol(i))

        strCol = Left(arrCol(i), InStr(arrCol(i), " ") - 1) '名称 Number ( 16, 5) Not Null Default 1.23
        strType = Mid(arrCol(i), InStr(arrCol(i), " ") + 1)
        If arrCol(i) Like "* PRIMARY KEY*" Or arrCol(i) Like "* UNIQUE*" And strCol = "CONSTRAINT" Then
            '检查约束列
            strTmpSQL = "ALTER TABLE " & IIf(strOwner = "", mstrUser, strOwner) & "." & strName & " ADD " & arrCol(i)
            strTmpSQL = CheckUniqueKeyCol(strTmpSQL, True)
            If strTmpSQL <> "" Then
                strConsSQL = strConsSQL & IIf(strConsSQL = "", "", vbNewLine) & strTmpSQL
            End If
        Else
            rsTemp.Filter = "Column_Name='" & strCol & "'"
            If rsTemp.EOF Then
                strModifySQL = strModifySQL & IIf(strModifySQL = "", "", vbNewLine) & "ALter Table " & IIf(strOwner = "", mstrUser, strOwner) & "." & strName & " Add " & arrCol(i) & ";"
                strError = strError & "," & strCol
            Else
                If strType Like rsTemp!DATA_TYPE & "*" Then
                    If rsTemp!DATA_TYPE = "NUMBER" Then
                        If InStr(strType, ",") > 0 Then 'Number(16,5)
                            intLen = Val(Split(Replace(Split(Split(strType, "(")(1), ")")(0), " ", ""), ",")(0))
                            intDigit = Val(Split(Replace(Split(Split(strType, "(")(1), ")")(0), " ", ""), ",")(1))
                        Else 'Number(18)
                            intLen = Val(Replace(Split(Split(strType, "(")(1), ")")(0), " ", ""))
                            intDigit = 0
                        End If
                        If rsTemp!Data_Precision < intLen Or rsTemp!Data_Scale < intDigit Then
                            strModifySQL = strModifySQL & IIf(strModifySQL = "", "", vbNewLine) & "ALter Table " & IIf(strOwner = "", mstrUser, strOwner) & "." & strName & " Modify " & arrCol(i) & ";"
                            strError = strError & "," & strCol
                        End If
                    ElseIf rsTemp!DATA_TYPE = "VARCHAR2" Then 'Varchar2(50)
                        intLen = Val(Replace(Split(Split(strType, "(")(1), ")")(0), " ", ""))
                        If rsTemp!Data_Length < intLen Then
                            strModifySQL = strModifySQL & IIf(strModifySQL = "", "", vbNewLine) & "ALter Table " & IIf(strOwner = "", mstrUser, strOwner) & "." & strName & " Modify " & arrCol(i) & ";"
                            strError = strError & "," & strCol
                        End If
                    ElseIf strType = rsTemp!DATA_TYPE Then
                        '不处理
                    End If
                Else
                    strError = strError & "," & strCol
                End If
            End If
        End If
    Next
    strError = Mid(strError, 2)
    If strConsSQL = "" Then
        merrCur.ErrAdvice = "已经存在的表""" & strName & """的字段""" & strError & """与SQL语句不符，请手工处理后重试。"
        CheckCreateTabCol = strError
        merrCur.ErrModiSQL = strModifySQL
    Else
        merrCur.ErrAdvice = "1、已经存在的表""" & strName & """的字段""" & strError & """与SQL语句不符，请手工处理后重试。"
        merrCur.ErrAdvice = "2、已经存在的主键或唯一约束的字段或字段顺序与SQL语句不符，请手工处理后重试。"""
        CheckCreateTabCol = strError & IIf(strError = "", "", vbNewLine) & strConsSQL
        merrCur.ErrModiSQL = strModifySQL & IIf(strModifySQL = "", "", vbNewLine) & strConsSQL
    End If
    merrCur.ErrOparateModi = vbRetry
End Function

Private Sub GetCreateName(ByVal strSQL As String, strOwner As String, strName As String, strType As String)
'功能：从Create SQL语句中返回所创建的对象名和类型
'参数：strSQL=已格式化为标准大写的SQL语句
'返回：创建的对象名,可能包含所有者,如"ZLHIS.部门表"
    strOwner = "": strName = "": strType = ""

    If strSQL Like "CREATE *" Then
        '只包含了User_Objects中的类型,没有包含Create Role,Tablespace
        If strSQL Like "* TABLE *" Then
            strType = "TABLE"
        ElseIf strSQL Like "* INDEX *" Then
            strType = "INDEX"
        ElseIf strSQL Like "* SEQUENCE *" Then
            strType = "SEQUENCE"
        ElseIf strSQL Like "* SYNONYM *" Then
            strType = "SYNONYM"
        ElseIf strSQL Like "* MATERIALIZED VIEW *" Then
            strType = "MATERIALIZED VIEW"
        ElseIf strSQL Like "* VIEW *" Then
            strType = "VIEW"
        ElseIf strSQL Like "* TRIGGER *" Then
            strType = "TRIGGER"
        ElseIf strSQL Like "* FUNCTION *" Then
            strType = "FUNCTION"
        ElseIf strSQL Like "* PROCEDURE *" Then
            strType = "PROCEDURE"
        ElseIf strSQL Like "* TYPE BODY *" Then
            strType = "TYPE BODY"
        ElseIf strSQL Like "* TYPE *" Then
            strType = "TYPE"
        ElseIf strSQL Like "* PACKAGE BODY *" Then
            strType = "PACKAGE BODY"
        ElseIf strSQL Like "* PACKAGE *" Then
            strType = "PACKAGE"
        End If
        strName = Split(Mid(strSQL, InStr(strSQL, " " & strType & " ") + Len(strType) + 2), " ")(0)
        If InStr(strName, "(") > 0 Then
            strName = Left(strName, InStr(strName, "(") - 1)
        End If
        If InStr(strName, ".") > 0 Then
            strOwner = Split(strName, ".")(0)
            strName = Split(strName, ".")(1)
        End If
    End If
End Sub


Private Function ObjectExists(ByVal strOwner As String, ByVal strName As String, ByVal strType As String, Optional ByVal strTableName As String) As Boolean
'功能：判断指定的对象是否存在
'说明：当为约束时，必须传strTableName
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String

    On Error Resume Next
    If strType <> "CONSTRAINT" Then
        strSQL = "Select Object_Name From All_Objects Where Owner=" & IIf(strType = "SYNONYM", "'PUBLIC'", IIf(strOwner = "", "User", "'" & strOwner & "'")) & " And Object_Type='" & strType & "' And Object_Name='" & strName & "'"
    Else
        strSQL = "Select Constraint_Name" & vbNewLine & _
        "From All_Constraints a" & vbNewLine & _
        "Where a.Constraint_Name = '" & strName & "' And a.Table_Name = '" & strTableName & "' And a.Owner = " & IIf(strOwner = "", "User", "'" & strOwner & "'")
    End If
    Set rsTemp = gclsBase.OpenSQLRecord(mcnThis, strSQL, App.Title)
    ObjectExists = Not rsTemp.EOF
End Function


Private Sub CheckAdjustSequence(ByVal strErr As String)
'功能：检查并修正序列
 '说明，错误为违反唯一约束的才进行检查
 '          [Microsoft][ODBC driver for Oracle][Oracle]ORA-00001: 违反唯一约束条件 (ZLTOOLS.ZLPROGPRIVS_PK)
    Dim strConstraint As String, strUser As String, strTable As String, strSeqName As String
    Dim strSeqCol As String, intCurMax As Long
    Dim arrTmp As Variant
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    '获取约束名称以及所有者
    strErr = UCase(strErr)
    arrTmp = Split(strErr, "ORA-00001:")
    If UBound(arrTmp) <> 1 Then Exit Sub
    strErr = arrTmp(1)
    If Not strErr Like "*(*)*" Then Exit Sub
    strConstraint = Split(Split(strErr, ")")(0), "(")(1)
    arrTmp = Split(strConstraint, ".")
    strConstraint = arrTmp(1)
    strUser = arrTmp(0)
    '获取约束的详细信息
    strSQL = "Select Constraint_Name, Table_Name" & vbNewLine & _
                    "From All_Constraints a" & vbNewLine & _
                    "Where A.Owner =[1] And A.Constraint_Name = [2]"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "错误处理", strUser, strConstraint)
    If rsTmp.EOF Then Exit Sub
    strTable = rsTmp!Table_Name & ""
    '判断是否存在该表的序列
    strSQL = "Select B.Sequence_Name, Substr(Sequence_Name, 1, Instr(Sequence_Name, '_') - 1) Table_Name," & vbNewLine & _
                    "       Substr(Sequence_Name, Instr(Sequence_Name, '_') + 1) Column_Name, B.Last_Number" & vbNewLine & _
                    "From All_Sequences b" & vbNewLine & _
                    "Where B.Sequence_Owner =[1] And B.Sequence_Name Like [2]"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "错误处理", strUser, strTable & "_%")
    If rsTmp.EOF Then Exit Sub
    strSeqName = rsTmp!Sequence_Name: strSeqCol = rsTmp!Column_Name: intCurMax = Val(rsTmp!Last_Number & "")
    '判断约束列中存在不存在序列的列
    strSQL = "Select 1" & vbNewLine & _
                "From All_Cons_Columns a" & vbNewLine & _
                "Where A.Owner =[1] And A.Constraint_Name =[2] And A.Column_Name = [3]"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "错误处理", strUser, strConstraint, strSeqCol)
    If rsTmp.EOF Then Exit Sub
    
    '序列修正匿名块
    strSQL = "Declare" & vbNewLine & _
                    "  N_Maxid Number(18);" & vbNewLine & _
                    "  N_Curid Number(18);" & vbNewLine & _
                    "  N_Incre Number(18);" & vbNewLine & _
                    "  N_Tmp   Number(18);" & vbNewLine & _
                    "Begin" & vbNewLine & _
                    "  --获取数据表中该列最大值"
    If strTable = "门诊费用记录" Or strTable = "住院费用记录" Or strTable = "病人费用记录" Then
'        strSQL = strSQL & vbNewLine & _
'                    "  Select Max(Id) Into N_Maxid From [*所有者*].[*表名*];"
        strSQL = strSQL & vbNewLine & _
                    "  Select Max(Mid)" & vbNewLine & _
                    "  Into N_Maxid" & vbNewLine & _
                    "  From (Select Max(" & strSeqCol & ") As Mid" & vbNewLine & _
                    "         From [*所有者*].门诊费用记录" & vbNewLine & _
                    "         Union All" & vbNewLine & _
                    "         Select Max(" & strSeqCol & ") As Mid From [*所有者*].住院费用记录);"
    Else
        strSQL = strSQL & vbNewLine & "  Select Max(" & strSeqCol & ") Into N_Maxid From [*所有者*].[*表名*];"
    End If
    strSQL = strSQL & vbNewLine & _
                    "  N_Maxid := Nvl(N_Maxid, 0);" & vbNewLine & _
                    "  --获取当期序列值" & vbNewLine & _
                    "  Select [*序列名*].Nextval Into N_Curid From Dual;" & vbNewLine & _
                    "  --修正序列" & vbNewLine & _
                    "  If N_Maxid - N_Curid > 0 Then" & vbNewLine & _
                    "    --获取序列当前增量" & vbNewLine & _
                    "    Select Increment_By" & vbNewLine & _
                    "    Into N_Incre" & vbNewLine & _
                    "    From All_Sequences" & vbNewLine & _
                    "    Where Sequence_Owner = '[*所有者*]' And Sequence_Name = '[*序列名*]';" & vbNewLine & _
                    "    N_Incre := Nvl(N_Incre, 1);" & vbNewLine & _
                    "    --修正成反向增量" & vbNewLine & _
                    "    Execute Immediate 'Alter Sequence [*所有者*].[*序列名*] Increment By ' ||(N_Maxid - N_Curid);" & vbNewLine & _
                    "    --移动一次增量" & vbNewLine & _
                    "    Select [*序列名*].Nextval Into N_Tmp From Dual;" & vbNewLine & _
                    "    --恢复原始增量" & vbNewLine & _
                    "    Execute Immediate 'Alter Sequence [*所有者*].[*序列名*] Increment By ' ||N_Incre;" & vbNewLine & _
                    "  End If;" & vbNewLine & _
                    "End;" & vbNewLine & _
                    "/"
    strSQL = Replace(Replace(Replace(strSQL, "[*所有者*]", strUser), "[*表名*]", strTable), "[*序列名*]", strSeqName)
    merrCur.ErrOparateModi = vbRetry
    merrCur.ErrModiSQL = strSQL
End Sub

Private Sub CheckAdjustConstraint(ByVal strSQL As String)
'功能：检查并修正约束
 'ORA-00955: 名称已被现有对象占用
 '              1. ALTER TABLE * ADD CONSTRAINT * PRIMARY KEY
'               2. ALTER TABLE * ADD CONSTRAINT * UNIQUE*
    Dim strTable As String, strConstraint As String
    Dim strUser As String
    Dim arrTmp As Variant
    
    If Not (strSQL Like "ALTER TABLE * ADD CONSTRAINT * PRIMARY KEY*" Or _
        strSQL Like "ALTER TABLE * ADD CONSTRAINT * UNIQUE*") Then Exit Sub
    strTable = Split(Mid(strSQL, InStr(strSQL, "ALTER TABLE ") + Len("ALTER TABLE ")), " ")(0)
    strConstraint = Split(Mid(strSQL, InStr(strSQL, "ADD CONSTRAINT ") + Len("ADD CONSTRAINT ")), " ")(0)
    On Error Resume Next
    arrTmp = Split(strTable, ".")
    If UBound(arrTmp) > 0 Then
        strUser = arrTmp(0)
        strTable = arrTmp(1)
    Else
        strUser = mstrUser
    End If
    '存在同名索引
    If ObjectExists(strUser, strConstraint, "INDEX") Then
        merrCur.ErrAdvice = "约束被同名索引占用，请执行删除索引后重试。"
        merrCur.ErrOparate = vbRetry
        merrCur.ErrModiSQL = "Drop Index " & strConstraint & ";"
        merrCur.ErrOparateModi = vbRetry
    End If
End Sub

Private Sub CheckAdjust12899(ByVal strErr As String)
'ORA-12899: 列 "ZLHIS"."体检任务人员"."所属团体" 的值太大 (实际值: 62, 最大值: 60)
'ORA-12899: value too large for column "SYSTEM"."STUDENTINFO"."SNAME" (actual: 78, maximum: 30)
    Dim arrTmp As Variant
    Dim strColLen As String
    Dim strOwner As String, strTable As String, strColName As String
    Dim strInfo As String, intNewLen As Integer, intOldLen As Integer
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    arrTmp = Split(strErr, "(")
    '获取列精度信息
    If UBound(arrTmp) < 1 Then Exit Sub
    strColLen = arrTmp(1)
    strColLen = Split(strColLen, ")")(0)
    '获取列信息
    strInfo = Replace(arrTmp(0), """.""", ".")
    arrTmp = Split(strInfo, """")
    If UBound(arrTmp) < 1 Then Exit Sub
    strInfo = Trim(arrTmp(1)) 'Ownere.Talbe.Col
    arrTmp = Split(strInfo, ".")
    If UBound(arrTmp) < 2 Then Exit Sub
    strOwner = UCase(Trim(arrTmp(0)))
    strTable = UCase(Trim(arrTmp(1)))
    strColName = UCase(Trim(arrTmp(2)))
    '获取新长度
    arrTmp = Split(strColLen, ":")
    intOldLen = Val(arrTmp(UBound(arrTmp)))
    intNewLen = Val(arrTmp(1))
    If intNewLen = 0 Or intOldLen = 0 Then Exit Sub
    strSQL = "Select Data_Type, Data_Length, Data_Precision, Data_Scale" & vbNewLine & _
            "From All_Tab_Columns" & vbNewLine & _
            "Where Owner = [1] And Table_Name = [2] And Column_Name = [3]"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "CheckAdjust12899", strOwner, strTable, strColName)
    If Not rsTmp.EOF Then
        If rsTmp!DATA_TYPE Like "*CHAR*" Or rsTmp!DATA_TYPE = "RAW" Then
            merrCur.ErrOparate = vbRetry
            merrCur.ErrOparateModi = vbRetry
            merrCur.ErrModiSQL = "alter table " & strOwner & "." & strTable & " modify " & strColName & " " & rsTmp!DATA_TYPE & "(" & intNewLen & ");"
        ElseIf rsTmp!DATA_TYPE = "NUMBER" Then
            merrCur.ErrOparate = vbRetry
            merrCur.ErrOparateModi = vbRetry
            intNewLen = Val(rsTmp!Data_Precision & "") + intNewLen - intOldLen
            If Val(rsTmp!Data_Scale & "") = 0 Then
                merrCur.ErrModiSQL = "alter table " & strOwner & "." & strTable & " modify " & strColName & " " & rsTmp!DATA_TYPE & "(" & intNewLen & ");"
            Else
                merrCur.ErrModiSQL = "alter table " & strOwner & "." & strTable & " modify " & strColName & " " & rsTmp!DATA_TYPE & "(" & intNewLen & "," & rsTmp!Data_Scale & ");"
            End If
        End If
    End If
End Sub

Private Sub CheckAdjustIndex(ByVal strErr As String)
'ORA-01502: index 'XXX' or partition of such index is in unusable state
    Dim arrTmp As Variant
    Dim strIndex As String
    Dim strUser As String
    
    arrTmp = Split(strErr, "'")
    strIndex = arrTmp(1)
    arrTmp = Split(strIndex, ".")
    If UBound(arrTmp) = 0 Then
        strUser = mstrUser
        strIndex = arrTmp(0)
    Else
        strUser = arrTmp(0)
        strIndex = arrTmp(1)
    End If
    merrCur.ErrOparate = vbRetry
    merrCur.ErrModiSQL = "ALter Index " & strUser & "." & strIndex & "  Rebuild;"
    merrCur.ErrOparateModi = vbRetry
End Sub

Private Sub CheckAdjustProcedure(ByVal strErr As String)
        'ORA-06575: 程序包或函数 TTT 处于无效状态
        'ORA-06575: Package or function TTT is in an invalid state
'PROCEDURE', 'FUNCTION', 'VIEW', 'MATERIALIZED VIEW', 'TRIGGER', 'PACKAGE', 'PACKAGE BODY', 'TYPE', 'TYPE BODY'
    Dim arrTmp As Variant
    Dim strProcedure As String
    Dim strType As String, strUser As String

    arrTmp = Split(strErr, "ORA-06575:")
    If UBound(arrTmp) <> 1 Then Exit Sub
    arrTmp = Split(arrTmp(1), " ")
    If UBound(arrTmp) = 3 Then
        strProcedure = arrTmp(2)
    ElseIf UBound(arrTmp) = 9 Then
        strProcedure = arrTmp(4)
    End If
    If strProcedure <> "" Then
        arrTmp = Split(strProcedure, ".")
        If UBound(arrTmp) = 0 Then
            strUser = mstrUser
            strProcedure = arrTmp(0)
        Else
            strUser = arrTmp(0)
            strProcedure = arrTmp(1)
        End If
        strType = GetObjectType(strUser, strProcedure)
        If strType <> "" Then
            merrCur.ErrOparateModi = vbRetry
            merrCur.ErrModiSQL = "alter " & strType & " " & strUser & "." & strProcedure & " compile;"
        End If
    End If
End Sub

Private Sub CheckAdjustSequnceVali(ByVal strErr As String)
'ORA-08002:sequence string.CURRVAL is not yet defined in this session
 'ORA-08002: 序列 ZLRPTDATAS_ID.CURRVAL 尚未在此进程中定义
    Dim strSeq As String
    Dim arrTmp As Variant
    
    strErr = UCase(strErr)
    If strErr Like "ORA-08002: 序列*" Then
        strSeq = Mid(strErr, Len("ORA-08002: 序列") + 1)
        strSeq = Trim(Mid(strSeq, 1, Len(strSeq) + Len("尚未在此进程中定义")))
    ElseIf strErr Like "ORA-08002: SEQUENCE *" Then
        strSeq = Mid(strErr, Len("ORA-08002: SEQUENCE ") + 1)
        strSeq = Split(Trim(strSeq), " ")(0)
    End If
    If strSeq Like "*.CURRVAL" Then
        strSeq = Replace(strSeq, ".CURRVAL", ".Nextval")
        merrCur.ErrOparateModi = vbRetry
        merrCur.ErrModiSQL = "Select " & strSeq & " From Dual;"
    End If
End Sub

Private Function GetObjectType(ByVal strOwner As String, strName As String) As String
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    strSQL = "Select a.OBJECT_TYPE" & vbNewLine & _
                "From All_Objects a" & vbNewLine & _
                "Where A.Status <> 'VALID' And A.Object_Name =[1] And  A.Owner =[2]"
    Set rsTmp = gclsBase.OpenSQLRecord(mcnThis, strSQL, App.Title, UCase(strName), UCase(strOwner))
    If Not rsTmp.EOF Then
        GetObjectType = rsTmp!Object_Type & ""
    End If
End Function

'Call CheckAdjustTableData(mobjSQL.SQL, 1, merrCur.ErrDesc)
Private Function CheckAdjustTableData(ByVal strInputSQL As String, ByVal bytMode As Byte, ByVal strErrDesc As String) As Boolean
'对于字典表自动修复
'错误:ORA-00001: 违反唯一约束条件 (ZLHIS.感染因素_PK)
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strOwner As String, strTableName As String, strConsCols As String, strConsColDataType As String
    Dim strSeqName As String, strSeqCol As String, blnMultiCon As Boolean
    Dim strDataSQL As String, strUpdateCols As String, strCols As String
    Dim arrConsCols As Variant, arrConsColDataType As Variant, arrUpdateCol As Variant
    Dim strWhereSQL As String, strTmp As String, strDataCol As String, strAdjustSQL As String
    Dim i As Integer, blnCanAdjust As Boolean, strHint As String
    '解析错误中的信息，并获取表的各种约束序列信息
    blnCanAdjust = True
    If Not GetConstraintInfo(strErrDesc, strOwner, strTableName, _
                    strConsCols, strConsColDataType, strSeqName, strSeqCol, blnMultiCon, strHint) Then
        blnCanAdjust = False
    ElseIf strHint = "" Then
        '从SQL中解析产生不包含序列的SQL,并获取插入的数据列
        If Not GetSQLData(bytMode, strInputSQL, strSeqName, strSeqCol, strOwner & "." & strTableName, strDataSQL, strCols) Then
            blnCanAdjust = False
        End If
    End If
    If blnCanAdjust And strHint = "" Then
        arrConsCols = Split(strConsCols, ",")
        arrConsColDataType = Split(strConsColDataType, ",")
    '    strWhereSQL = "Select 1 From " & strOwner & "." & strTableName & " b Where "
        strWhereSQL = ""
        strUpdateCols = "," & strCols & ","
        '从插入列中剔除所有的主键与唯一约束列
        For i = LBound(arrConsCols) To UBound(arrConsCols)
            If InStr(strUpdateCols, "," & arrConsCols(i) & ",") = 0 Then
                mclsrun.WriteLog String(17, " ") & "错误中心(系统)警告：" & strOwner & "." & strTableName & "的插入SQL中没有包含所有的唯一或主键约束(缺失列：" & arrConsCols(i) & ")，无法自动修复。"
                blnCanAdjust = False
            Else
                strUpdateCols = Replace(strUpdateCols, "," & arrConsCols(i) & ",", ",")
            End If
            If arrConsColDataType(i) <> "-1" Then
                If strTableName = "ZLPROGPRIVS" Then '对象权限特殊处理
                    If InStr("所有者,对象,权限", arrConsCols(i)) > 0 Then
                        strWhereSQL = strWhereSQL & " Upper(a." & arrConsCols(i) & ")=Upper(b." & arrConsCols(i) & ") And "
                    Else
                        strWhereSQL = strWhereSQL & " Nvl(a." & arrConsCols(i) & ",0)=NVL(b." & arrConsCols(i) & ",0) And "
                    End If
                Else
                    If arrConsColDataType(i) = 0 Then
                        strWhereSQL = strWhereSQL & " Nvl(a." & arrConsCols(i) & ",'NONEDATA')=NVL(b." & arrConsCols(i) & ",'NONEDATA') And "
                    ElseIf arrConsColDataType(i) = 1 Then
                        strWhereSQL = strWhereSQL & " Nvl(a." & arrConsCols(i) & ",0)=NVL(b." & arrConsCols(i) & ",0) And "
                    ElseIf arrConsColDataType(i) = 2 Then
                        strWhereSQL = strWhereSQL & " Nvl(a." & arrConsCols(i) & ",SYSDATE)=NVL(b." & arrConsCols(i) & ",SYSDATE) And "
                    End If
                End If
            Else
                mclsrun.WriteLog String(17, " ") & "错误中心(系统)警告：" & strOwner & "." & strTableName & "." & arrConsCols(i) & "的数据类型无法提供自动修正，请联系开发人员"
                blnCanAdjust = False
            End If
        Next
        If blnCanAdjust Then
            strWhereSQL = Mid(strWhereSQL, 1, Len(strWhereSQL) - Len(" And "))
            If strUpdateCols = "," Then
                strUpdateCols = ""
            ElseIf strUpdateCols <> "" Then
                strUpdateCols = Mid(strUpdateCols, 2, Len(strUpdateCols) - 2)
            End If
            If strTableName = "ZLPARAMETERS" Then '判断该参数是35.0以上还是以下，尽管当前是35代码，但是仍旧需要进行参数说明更新。因为多系统共享，可能部分系统还是低于35
                If InStr(strCols, "影响控制说明") > 0 Then
                    strUpdateCols = "影响控制说明,参数值含义,关联说明,适用说明,警告说明"
                Else
                    strUpdateCols = "参数说明"
                End If
            End If
            '生成插入SQL
            strSQL = "Select " & strCols & " From (" & strDataSQL & ") a Where Not Exists(Select 1 From " & strOwner & "." & strTableName & " b Where " & strWhereSQL & ")"
            On Error Resume Next
            Set rsTmp = gclsBase.OpenSQLRecord(mcnThis, strSQL, App.Title)
            If rsTmp Is Nothing Then
                mclsrun.WriteLog String(17, " ") & "错误中心(系统)警告：获取插入数据出错，该错误无法自动修复,信息(" & err.Description & ",SQL:" & strSQL & ")。"
                err.Clear
                blnCanAdjust = False
            Else
                Do While Not rsTmp.EOF
                    strDataCol = ""
                    For i = 0 To rsTmp.Fields.Count - 1
                        If IsNull(rsTmp.Fields(i).value) Then
                            If IsType(rsTmp.Fields(i).Type, adNumeric) Then
                                strDataCol = strDataCol & ",-Null"
                            Else
                                strDataCol = strDataCol & ",Null"
                            End If
                        ElseIf IsType(rsTmp.Fields(i).Type, adDate) Then
                            strDataCol = strDataCol & "," & "To_Date('" & Format(rsTmp.Fields(i).value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                        ElseIf IsType(rsTmp.Fields(i).Type, adVarChar) Then
                            strDataCol = strDataCol & "," & SQLAdjust(rsTmp.Fields(i).value)
                        ElseIf IsType(rsTmp.Fields(i).Type, adNumeric) Then
                            strDataCol = strDataCol & "," & rsTmp.Fields(i).value
                        Else
                            mclsrun.WriteLog String(17, " ") & "错误中心(系统)警告：" & strOwner & "." & strTableName & "." & rsTmp.Fields(i).name & "的数据类型无法提供自动修正，请联系开发人员"
                        End If
                    Next
                    strAdjustSQL = strAdjustSQL & vbNewLine & "Insert Into " & strOwner & "." & strTableName & _
                                "(" & IIf(strSeqCol <> "", strSeqCol & ",", "") & strCols & ")" & _
                                "Select " & IIf(strSeqName <> "", strSeqName & ".Nextval,", "") & Mid(strDataCol, 2) & " From Dual;"
                    mclsrun.WriteLog String(17, " ") & "错误中心(系统)新增数据：" & strCols & "(" & Mid(strDataCol, 2) & ")"
                    rsTmp.MoveNext
                Loop
            End If
            If blnCanAdjust Then
                '生成更新SQL
                strSQL = "Select " & AddTablePreSubfix(strCols, "A") & IIf(strUpdateCols = "", "", "," & AddTablePreSubfix(strUpdateCols, "B", "B")) & " From (" & strDataSQL & ") a ," & strOwner & "." & strTableName & " b Where " & strWhereSQL
                Set rsTmp = gclsBase.OpenSQLRecord(mcnThis, strSQL, App.Title)
                If rsTmp Is Nothing Then
                    mclsrun.WriteLog String(17, " ") & "错误中心(系统)警告：获取更新数据出错，该错误无法自动修复,信息(" & err.Description & ",SQL:" & strSQL & ")。"
                    err.Clear
                    blnCanAdjust = False
                Else
                    
                    arrUpdateCol = Split(strUpdateCols, ",")
                    Do While Not rsTmp.EOF
                        strDataCol = "": strTmp = ""
                        If strUpdateCols <> "" Then '存在更新字段,则获取更新字段SQL
                            For i = LBound(arrUpdateCol) To UBound(arrUpdateCol)
                                strTmp = strTmp & "," & arrUpdateCol(i) & ":"
                                If IsNull(rsTmp.Fields(arrUpdateCol(i)).value) Then
                                    strDataCol = strDataCol & "," & arrUpdateCol(i) & "=Null"
                                    strTmp = strTmp & "Null"
                                ElseIf IsType(rsTmp.Fields(arrUpdateCol(i)).Type, adDate) Then
                                    strDataCol = strDataCol & "," & arrUpdateCol(i) & "=" & "To_Date('" & Format(rsTmp.Fields(arrUpdateCol(i)).value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                                    strTmp = strTmp & Format(rsTmp.Fields(arrUpdateCol(i)).value, "yyyy-MM-dd HH:mm:ss")
                                ElseIf IsType(rsTmp.Fields(arrUpdateCol(i)).Type, adVarChar) Then
                                    strDataCol = strDataCol & "," & arrUpdateCol(i) & "=" & SQLAdjust(rsTmp.Fields(arrUpdateCol(i)).value)
                                    strTmp = strTmp & rsTmp.Fields(arrUpdateCol(i)).value
                                ElseIf IsType(rsTmp.Fields(arrUpdateCol(i)).Type, adNumeric) Then
                                    strDataCol = strDataCol & "," & arrUpdateCol(i) & "=" & rsTmp.Fields(arrUpdateCol(i)).value
                                    strTmp = strTmp & rsTmp.Fields(arrUpdateCol(i)).value
                                Else
                                    mclsrun.WriteLog String(17, " ") & "错误中心(系统)警告：" & strOwner & "." & strTableName & "." & arrUpdateCol(i) & "的数据类型无法提供自动修正，请联系开发人员"
                                End If
                                strTmp = strTmp & "->"
                                '增加数据字段值记录
                                If IsNull(rsTmp.Fields(arrUpdateCol(i) & "B").value) Then
                                    strTmp = strTmp & "Null"
                                ElseIf IsType(rsTmp.Fields(arrUpdateCol(i) & "B").Type, adDate) Then
                                    strTmp = strTmp & Format(rsTmp.Fields(arrUpdateCol(i) & "B").value, "yyyy-MM-dd HH:mm:ss")
                                ElseIf IsType(rsTmp.Fields(arrUpdateCol(i)).Type, adVarChar) Then
                                    strTmp = strTmp & rsTmp.Fields(arrUpdateCol(i) & "B").value
                                ElseIf IsType(rsTmp.Fields(arrUpdateCol(i) & "B").Type, adNumeric) Then
                                    strTmp = strTmp & rsTmp.Fields(arrUpdateCol(i) & "B").value
                                End If
                                
                            Next
                        End If
                        strWhereSQL = ""
                        For i = LBound(arrConsCols) To UBound(arrConsCols) '获取更新的约束条件
                            If IsNull(rsTmp.Fields(arrConsCols(i)).value) Then
                                strWhereSQL = strWhereSQL & " And " & arrConsCols(i) & " Is Null "
                            Else
                                If arrConsColDataType(i) = 0 Then
                                    strWhereSQL = strWhereSQL & " And " & arrConsCols(i) & " = " & SQLAdjust(rsTmp.Fields(arrConsCols(i)).value)
                                ElseIf arrConsColDataType(i) = 1 Then
                                    strWhereSQL = strWhereSQL & " And " & arrConsCols(i) & " = " & rsTmp.Fields(arrConsCols(i)).value
                                ElseIf arrConsColDataType(i) = 2 Then
                                    strWhereSQL = strWhereSQL & " And " & arrConsCols(i) & " = " & "To_Date('" & Format(rsTmp.Fields(arrConsCols(i)).value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                                End If
                            End If
                        Next
                        If strUpdateCols <> "" And blnMultiCon Then
                            strAdjustSQL = strAdjustSQL & vbNewLine & " Update " & strOwner & "." & strTableName & " Set " & Mid(strDataCol, 2) & " Where " & Mid(strWhereSQL, Len(" And  ")) & ";"
                            mclsrun.WriteLog String(17, " ") & "错误中心(系统)更新数据：" & Mid(strWhereSQL, Len(" And  ")) & "([字段名:SQL中数值->数据库数值]" & Mid(strTmp, 2) & ")"
                        Else
                            mclsrun.WriteLog String(17, " ") & "错误中心(系统)已经存在数据-" & IIf(strUpdateCols = "", "无更新列", "单主键或唯一键") & "：" & Mid(strWhereSQL, Len(" And  ")) & IIf(strUpdateCols = "", "", "([字段名:SQL中数值->数据库数值]" & Mid(strTmp, 2) & ")")
                        End If
                        rsTmp.MoveNext
                    Loop
                End If
            End If
        End If
    '可以通过11g新特性避免重复数据导致的插入失败
    ElseIf strHint <> "" Then
        strInputSQL = Trim(strInputSQL)
        strAdjustSQL = "Insert " & strHint & Mid(strInputSQL, Len("Insert") + 1) & ";"
        mclsrun.WriteLog String(17, " ") & "错误中心(系统)处理：通过11g新特性避免插入数据失败" & strHint
    End If
    If strAdjustSQL <> "" And blnCanAdjust Then
        strSQL = Mid(strAdjustSQL, Len(vbNewLine) + 1)
        merrCur.ErrAdvice = "可能这些字段未更新或部分数据已经存在，请检查确认处理！"
        merrCur.ErrOparate = vbRetry
        merrCur.ErrOparateModi = vbIgnore
        merrCur.ErrModiSQL = merrCur.ErrModiSQL & vbNewLine & strAdjustSQL
        CheckAdjustTableData = True
    ElseIf Not blnCanAdjust Then
        If bytMode = 0 Then
            merrCur.ErrAdvice = "可能是语句重复运行出错，一般情况下可以忽略该错误。"
            merrCur.ErrOparate = vbIgnore
            merrCur.ErrOparateModi = vbIgnore
        Else
            merrCur.ErrAdvice = "无法自动检查数据，请检查确认SQL和数据库一致！"
            merrCur.ErrOparate = vbRetry
            merrCur.ErrOparateModi = vbIgnore
        End If
    Else
        merrCur.ErrAdvice = "脚本中数据已经存在，请检查确认处理！"
        merrCur.ErrOparate = vbIgnore
        merrCur.ErrOparateModi = vbIgnore
        CheckAdjustTableData = True
    End If
End Function

Private Function AddTablePreSubfix(ByVal strCols As String, Optional ByVal strPresubfix As String, Optional ByVal strAlisSubFix As String) As String
'功能：给字段增加增加表前缀或者别名
'参数：strCols-字段集合，以逗号分割
'      strPresubfix=表名前缀
'      strAlisSubFix=字段的别名后最，如 COLA会生成，strPresubfix.COLA  COLA&strAlisSubFix
'返回：生成后的列名
    Dim arrTmp As Variant, i As Integer
    Dim strReturn As String
    
    If strCols = "" Then Exit Function
    If strPresubfix = "" And strAlisSubFix = "" Then
        AddTablePreSubfix = strCols
        Exit Function
    End If
    arrTmp = Split(strCols, ",")
    For i = LBound(arrTmp) To UBound(arrTmp)
        strReturn = strReturn & "," & IIf(strPresubfix <> "", strPresubfix & ".", "") & arrTmp(i)
        If strAlisSubFix <> "" Then
            strReturn = strReturn & " " & arrTmp(i) & strAlisSubFix
        End If
    Next
    strReturn = Mid(strReturn, 2)
    AddTablePreSubfix = strReturn
End Function


Private Function GetConstraintInfo(ByVal strErrDesc As String, ByRef strOwner As String, ByRef strTableName As String, ByRef strConsCols As String, ByRef strConsColDataType As String, ByRef strSeqName As String, ByRef strSeqCol As String, ByRef blnMultiCon As Boolean, ByRef strHint As String) As Boolean
'功能：通过错误描述获取表名,表的所有者，所有唯一主键约束列,该表存在的序列等信息
'参数：strErrDesc=数据插入产生的错误信息
'返回：是否获取成功
'strOwner=获取的表所有者
'strTableName=获取的表名
'strConsCols=该表上存在的所有唯一约束与主键约束的列的合集，以逗号分割列。注意，噶合集排除了序列对应列
'strConsColDataType=约束列的列类型，-1-不能进行脚本自动修正的类型，0-Char,1-NUMber,2-date
'strSeqName=该表上存在的序列
'strSeqCols=该表上序列对应的列
'blnMultiCon=是否存在多个约束，排除序列对应主键
'说明：仅有病人医嘱发送、病人医嘱记录、排队叫号队列、病理归档信息存在2个序列，其余表只有一个，因此只考虑单表单序列
    Dim arrTmp As Variant, strConName As String
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strTmp As String
    On Error Resume Next
    '前三种错误较多，因此手工指定
    strHint = ""
    If strErrDesc Like "*ZLTOOLS.ZLPARAMETERS_UQ_*" Or strErrDesc Like "*ZLTOOLS.ZLPARAMETERS_PK*" Then
        strOwner = "ZLTOOLS"
        strTableName = "ZLPARAMETERS"
        strConsCols = "系统,模块,参数号,参数名"
        strConsColDataType = "1,1,1,0"
        strSeqName = "ZLPARAMETERS_ID"
        strSeqCol = "ID"
        blnMultiCon = True
    ElseIf merrCur.ErrDesc Like "*ZLTOOLS.ZLPROGFUNCS_PK*" Then
        strOwner = "ZLTOOLS"
        strTableName = "ZLPROGFUNCS"
        strConsCols = "系统,序号,功能"
        strConsColDataType = "1,1,0"
        strSeqName = ""
        strSeqCol = ""
        If Not gblnClose11g Then
            If GetOracleVersion(True, True) >= 11 Then strHint = "/*+ IGNORE_ROW_ON_DUPKEY_INDEX(ZLPROGFUNCS,ZLPROGFUNCS_PK)*/ "
        End If
    ElseIf merrCur.ErrDesc Like "*ZLTOOLS.ZLPROGPRIVS_PK*" Then
        strOwner = "ZLTOOLS"
        strTableName = "ZLPROGPRIVS"
        strConsCols = "系统,序号,功能,所有者,对象,权限"
        strConsColDataType = "1,1,0,0,0,0"
        strSeqName = ""
        strSeqCol = ""
        If Not gblnClose11g Then
            If GetOracleVersion(True, True) >= 11 Then strHint = "/*+ IGNORE_ROW_ON_DUPKEY_INDEX(ZLPROGPRIVS,ZLPROGPRIVS_PK)*/ "
        End If
    Else
        '获取错误描述中的违反的约束对象
        arrTmp = Split(strErrDesc, "(")
        If UBound(arrTmp) < 1 Then
            mclsrun.WriteLog String(17, " ") & "错误中心(系统)警告：(1)无法解析错误信息，该错误无法自动修复。"
            Exit Function
        End If
        strTmp = arrTmp(1)
        arrTmp = Split(strTmp, ")")
        If UBound(arrTmp) < 1 Then
            mclsrun.WriteLog String(17, " ") & "错误中心(系统)警告：(2)无法解析错误信息，该错误无法自动修复。"
            Exit Function
        End If
        strTmp = arrTmp(0)
        arrTmp = Split(UCase(strTmp), ".")
         If UBound(arrTmp) < 1 Then
            mclsrun.WriteLog String(17, " ") & "错误中心(系统)警告：(3)无法解析错误信息，该错误无法自动修复。"
            Exit Function
        End If
        strOwner = Trim(arrTmp(0))
        strConName = Trim(arrTmp(1))
        '获取约束列与约束类型，约束不是唯一或者主键约束则退出，无法获取约束列也退出
        strSQL = "Select a.Constraint_Type, a.Table_Name" & vbNewLine & _
                "From All_Constraints a" & vbNewLine & _
                "Where a.Owner = [1] And a.Constraint_Name = [2] And a.Constraint_Type In ('P', 'U')"
        Set rsTmp = gclsBase.OpenSQLRecord(mcnThis, strSQL, "错误中心", strOwner, strConName)
        If rsTmp Is Nothing Then
            mclsrun.WriteLog String(17, " ") & "错误中心(系统)警告：获取违反约束的表出错，该错误无法自动修复,信息(" & err.Description & ",SQL:" & strSQL & ")。"
            err.Clear
            Exit Function
        ElseIf rsTmp.EOF Then
            mclsrun.WriteLog String(17, " ") & "错误中心(系统)警告：未找到违反约束的表，该错误无法自动修复。"
            Exit Function
        End If
        strTableName = rsTmp!Table_Name
        '只要不是业务表即ZLbakTale与ZLBigTables,则可以默认当作自动修正
        strSQL = "Select Count(1) 计数" & vbNewLine & _
                "From (Select 表名 From Zltools.Zlbaktables Union All Select 表名 From Zlbigtables) a" & vbNewLine & _
                "Where a.表名 = [1]"
        Set rsTmp = gclsBase.OpenSQLRecord(mcnThis, strSQL, "错误中心", strOwner, strTableName)
        If rsTmp Is Nothing Then
            mclsrun.WriteLog String(17, " ") & "错误中心(系统)警告：违反约束的表是否是业务表检查出错，该错误无法自动修复,信息(" & err.Description & ",SQL:" & strSQL & ")。"
            err.Clear
            Exit Function
        ElseIf rsTmp!计数 > 0 Then
            mclsrun.WriteLog String(17, " ") & "错误中心(系统)提示：违反约束的表为业务数据表，该错误无法自动修复。"
            Exit Function
        End If
        '获取序列
        strSQL = "Select a.SEQUENCE_NAME" & vbNewLine & _
                "From All_Sequences a" & vbNewLine & _
                "Where a.Sequence_Owner =[1] and a.SEQUENCE_NAME like [2]"
        Set rsTmp = gclsBase.OpenSQLRecord(mcnThis, strSQL, "错误中心", strOwner, strTableName & "_%")
        strSeqName = "": strSeqCol = ""
        If Not rsTmp Is Nothing Then
            If rsTmp.RecordCount > 1 Then
                mclsrun.WriteLog String(17, " ") & "错误中心(系统)提示：违反约束的表存在多个序列，该错误无法自动修复。"
                Exit Function '多序列不处理
            End If
            If Not rsTmp.EOF Then
                strSeqName = rsTmp!Sequence_Name
                strSeqCol = UCase(Mid(strSeqName, Len(strTableName & "_%")))
            End If
        Else
            mclsrun.WriteLog String(17, " ") & "错误中心(系统)警告：违反约束的表对应序列的检查出错，该错误无法自动修复,信息(" & err.Description & ",SQL:" & strSQL & ")。"
            err.Clear
            Exit Function
        End If
        If Not gblnClose11g Then
            If GetOracleVersion(True, True) >= 11 Then
                strSQL = "Select a.Table_Name, a.Index_Name" & vbNewLine & _
                        "From All_Indexes A" & vbNewLine & _
                        "Where a.Uniqueness = 'UNIQUE' And a.Table_Owner = [1] And a.Table_Name = [2]"
                Set rsTmp = gclsBase.OpenSQLRecord(mcnThis, strSQL, "错误中心", strOwner, strTableName)
                If rsTmp.RecordCount = 1 Then
                    Do While Not rsTmp.EOF
                        strHint = strHint & " IGNORE_ROW_ON_DUPKEY_INDEX(" & rsTmp!Table_Name & "," & rsTmp!Index_Name & ")"
                        rsTmp.MoveNext
                    Loop
                End If
                If strHint <> "" Then strHint = "/*+ " & strHint & "*/ "
            End If
        End If
        '获取唯一索引列与类型,直接写所有者，表名速度更快
        strSQL = "Select a.Column_Name, b.Data_Type,count(1) 计数" & vbNewLine & _
                "From all_ind_columns a, All_Tab_Columns b, all_indexes c" & vbNewLine & _
                "Where a.TABLE_OWNER = [1] And a.Table_Name = [2] And b.Owner = [1] And b.Table_Name = [2] And" & vbNewLine & _
                "      c.TABLE_OWNER = [1] And c.Table_Name = [2] And c.INDEX_NAME= a.INDEX_NAME And" & vbNewLine & _
                "      a.Column_Name = b.Column_Name And c.UNIQUENESS='UNIQUE'" & vbNewLine & _
                "group by   a.Column_Name, b.Data_Type"
        Set rsTmp = gclsBase.OpenSQLRecord(mcnThis, strSQL, "错误中心", strOwner, strTableName)
        If rsTmp Is Nothing Then
            mclsrun.WriteLog String(17, " ") & "错误中心(系统)警告：违反约束的表获取全部约束列出错，该错误无法自动修复,信息(" & err.Description & ",SQL:" & strSQL & ")。"
            err.Clear
            Exit Function
        End If
        blnMultiCon = False
        Do While Not rsTmp.EOF
            If strSeqCol <> rsTmp!Column_Name Then
                strConsCols = strConsCols & "," & rsTmp!Column_Name
                If rsTmp!DATA_TYPE Like "*CHAR*" Then
                    strConsColDataType = strConsColDataType & "," & 0
                Else
                    strConsColDataType = strConsColDataType & "," & Decode(rsTmp!DATA_TYPE, "NUMBER", 1, "DATE", 2, -1)
                End If
                If Not blnMultiCon Then blnMultiCon = Val(rsTmp!计数) > 1
            Else
                If Not blnMultiCon Then blnMultiCon = Val(rsTmp!计数) > 2
            End If
            rsTmp.MoveNext
        Loop
        If strConsCols = "" And strSeqName = "" Then
            mclsrun.WriteLog String(17, " ") & "错误中心(系统)警告：违反约束的表无法获取对应的约束列，该错误无法自动修复。"
            Exit Function
        End If
        strConsCols = Mid(strConsCols, 2)
        strConsColDataType = Mid(strConsColDataType, 2)
    End If
    If err.Number <> 0 Then err.Clear
    GetConstraintInfo = True
End Function

Private Function GetSQLData(ByVal bytMode As Byte, ByVal strSQL As String, ByVal strSeqName As String, ByVal strSeqCol As String, ByVal strTable As String, ByRef strDataSQL As String, ByRef strCols As String) As Boolean
'功能：分离SQL中的序列，并获取SQL中产生数据的SQL,与插入数据的列
'参数：bytMode=0-Insert Values方式 1-Insert Select 方式
'      strSQL=需要分离的SQL
'      strSeqName=表存在的序列名
'      strSeqCol=序列对应的列
'      strTable=所有者前缀的表
'返回：GetSQLData=是否分离成功
'      strDataSQL=产生数据的SQL,序列已经被提出。Values方式已经被改写为Select  From Dual方式
'      strCols=数据插入的列，已经提出序列
    Dim lngPos As Long, arrTmp As Variant, strTmp As String

    On Error GoTo errH
    If bytMode = 0 Then
        arrTmp = Split(strSQL, "values", , vbTextCompare)
        lngPos = InStrRev(arrTmp(0), ")")
        strCols = Mid(arrTmp(0), 1, lngPos - 1)
        lngPos = InStr(strCols, "(")
        strCols = Mid(strCols, lngPos + 1)
        strCols = CutSegByInfo(strCols, strSeqCol)
        lngPos = InStrRev(arrTmp(1), ")")
        strTmp = Mid(arrTmp(1), 1, lngPos - 1)
        lngPos = InStr(strTmp, "(")
        strTmp = Mid(strTmp, lngPos + 1)
        strTmp = CutSegByInfo(strTmp, strSeqName & ".Nextval")
        strTmp = "Select " & strTmp & " From Dual"
    Else
        lngPos = InStr(UCase(strSQL), "SELECT")
        strTmp = Mid(strSQL, lngPos + Len("SELECT") + 1)
        strCols = Mid(strSQL, 1, lngPos - 1)
        lngPos = InStrRev(strCols, ")")
        strCols = Mid(strCols, 1, lngPos - 1)
        lngPos = InStr(strCols, "(")
        strCols = Mid(strCols, lngPos + 1)
        strCols = CutSegByInfo(strCols, strSeqCol)
        strTmp = "SELECT " & CutSegByInfo(strTmp, strSeqName & ".Nextval")
    End If
    strCols = UCase(Replace(TrimEx(strCols, True), " ", ""))
    strDataSQL = "Select " & strCols & " From " & strTable & " Where 0=1 Union All " & vbNewLine & _
                strTmp
    
    GetSQLData = True
    Exit Function
errH:
    mclsrun.WriteLog String(17, " ") & "错误中心(系统)警告：获取插入数据的SQL出错，该错误无法自动修复,信息(" & err.Description & ")。"
    err.Clear
End Function

Private Function CutSegByInfo(ByVal strInput As String, ByVal strKeyInfo As String)
'功能：去掉以逗号分割的一个字符串中指定的一个段的信息。
    Dim arrTmp As Variant, lngCout As Long, i As Long
    Dim strReturn As String
    Dim arrCols As Variant

    arrTmp = Split(strInput, ",")
    For i = LBound(arrTmp) To UBound(arrTmp)
        lngCout = lngCout + Len(arrTmp(i)) - Len(Replace(arrTmp(i), "'", ""))
        If lngCout Mod 2 = 0 Then
            arrCols = Split(Trim(arrTmp(i)) & " ", " ")
            '防止如 Zlparameters_Id.Nextval AS id这样的写法
            If UCase(Trim(arrCols(0))) = UCase(Trim(strKeyInfo)) Then
                '定位到目标，不进行处理
            Else
                strReturn = strReturn & "," & arrTmp(i)
            End If
        Else
            strReturn = strReturn & "," & arrTmp(i)
        End If
    Next
    If strReturn <> "" Then
        strReturn = Mid(strReturn, 2)
    End If
    CutSegByInfo = strReturn
End Function

Private Function GetInsertColData(ByVal strSQL As String, ByVal bytMode As Byte, Optional ByRef strCols As String, Optional ByVal strRemoveCols As String) As String
'功能：从InsertSQL中解析出Insert语句的插入列与数据。
'暂时未使用
'参数：strSQL=解析的语句
'      strCols=语句的插入列
'      bytMode=语句格式0-Insert into values形式,1-Insert Into Select 形式
'      strRemoveCols=需要移除的列，如ID,部门

'返回：插入的列数据
    Dim strTmp As String, strTmpCols As String, strData As String
    Dim arrTmp As Variant, arrData As Variant, arrReMoveIndex As Variant
    Dim cllStr As Collection, i As Long, j As Long, lngCout As Long
    Dim strFTMSQL As String
    
    '分解列与数据
    strFTMSQL = UCase(TrimEx(GetFMTSQLStr(strSQL, cllStr), True))
    If bytMode = 0 Then
        arrTmp = Split(strFTMSQL, "VALUES")
        If UBound(arrTmp) < 1 Then Exit Function
        strData = GetInfoInsideBracket(arrTmp(1))
        strTmpCols = GetInfoInsideBracket(arrTmp(0))
        If strData = "" Or strTmpCols = "" Then Exit Function
        If strRemoveCols <> "" Then
            arrTmp = Split(strTmpCols, ",")
            strTmpCols = ""
            arrReMoveIndex = Split(strRemoveCols, ",")
            strRemoveCols = "," & strRemoveCols & ","
            For i = LBound(arrTmp) To UBound(arrTmp)
                If InStr(strRemoveCols, "," & arrTmp(i) & ",") Then
                    arrReMoveIndex = i
                Else
                    strTmpCols = strTmpCols & "," & arrTmp(i)
                End If
            Next
            strTmpCols = Mid(strTmpCols, 2)
        End If
        strData = "SELECT " & strTmp & " FROM DUAL"
    Else
        arrTmp = Split(strFTMSQL, "SELECT")
        strTmpCols = GetInfoInsideBracket(arrTmp(0))
        strData = Mid(strSQL, Len(arrTmp(0)) + 1)
    End If
    '移除某些列
    If strRemoveCols <> "" Then
        arrTmp = Split(strTmpCols, ",")
        strTmpCols = ""
        arrReMoveIndex = Split(strRemoveCols, ",")
        strRemoveCols = "," & strRemoveCols & ","
        arrReMoveIndex = Array()
        For i = LBound(arrTmp) To UBound(arrTmp)
            If InStr(strRemoveCols, "," & arrTmp(i) & ",") Then
                 ReDim Preserve arrReMoveIndex(UBound(arrReMoveIndex) + 1)
                arrReMoveIndex(UBound(arrReMoveIndex)) = i
            Else
                strTmpCols = strTmpCols & "," & arrTmp(i)
            End If
        Next
        strTmpCols = Mid(strTmpCols, 2)
        '移除数据列
        If UBound(arrReMoveIndex) > -1 Then
            arrTmp = Split(strData, " FROM ")
            For i = LBound(arrTmp) To UBound(arrTmp)
                arrData = Split(arrTmp(i), ",")
            Next
        End If
    End If
    strCols = strTmpCols
End Function

Private Sub fraSplit_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
     If Button = 1 And Index = 0 Then fraSplit(Index).Top = fraSplit(Index).Top + y
End Sub

Private Sub fraSplit_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Index = 1 Then Exit Sub
    If fraSplit(Index).Top <= 2500 Then fraSplit(Index).Top = 2500
    If picBottom.Top - fraSplit(Index).Top <= 3000 Then fraSplit(Index).Top = picBottom.Top - 3000
    Call Form_Resize
End Sub

Private Sub picErrInfo_Resize()
    rtfErr.Height = picErrInfo.ScaleHeight - rtfErr.Top
    rtfErr.Width = picErrInfo.ScaleWidth - rtfErr.Left
End Sub

Private Sub picModify_Resize()
    lblSQLErr.Top = picModify.ScaleHeight - lblSQLErr.Height - 30
    synModiSQL.Height = lblSQLErr.Top - synModiSQL.Top
    lblSQLErr.Width = picModify.ScaleWidth - lblSQLErr.Left
    synModiSQL.Width = picModify.ScaleWidth - synModiSQL.Left
    lblModify.Width = picModify.ScaleWidth - lblModify.Left
End Sub

Private Sub synModiSQL_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strTmp As String
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        synModiSQL.SelectAll 'Ctrl+A
    ElseIf KeyCode = vbKeyZ And Shift = vbCtrlMask Then
        If Val(synModiSQL.Tag) >= synModiSQL.RowsCount Then Exit Sub
        If synModiSQL.CurrPos.Row > Val(synModiSQL.Tag) Then
            synModiSQL.UnDo
        End If
    ElseIf KeyCode = vbKeyC And Shift = vbCtrlMask Then
        synModiSQL.Copy
    ElseIf KeyCode = vbKeyV And Shift = vbCtrlMask Then
        If Val(synModiSQL.Tag) >= synModiSQL.RowsCount Then Exit Sub
        If synModiSQL.CurrPos.Row > Val(synModiSQL.Tag) Then
            synModiSQL.Paste
        End If
    Else
        If synModiSQL.CurrPos.Row <= Val(synModiSQL.Tag) Then
            KeyCode = 0
        ElseIf synModiSQL.RowsCount <= Val(synModiSQL.Tag) + 1 Then
            If synModiSQL.RowsCount < Val(synModiSQL.Tag) + 1 Then
                synModiSQL.Text = synModiSQL.Text & vbNewLine & ""
            ElseIf Trim(synModiSQL.RowText(Val(synModiSQL.Tag) + 1)) = "" Then
                KeyCode = 0
            End If
        End If
    End If
End Sub


Private Sub synModiSQL_TextChanged(ByVal nRowFrom As Long, ByVal nRowTo As Long, ByVal nActions As Long)
    Call RefreshButton
End Sub

Private Sub tmrRefresh_Timer()
    Me.Refresh
End Sub

Private Sub tmrThis_Timer()
    mblnAuto = True
    If mobjSQL.IsSameTo(mobjPreSQL) Then
        mintTimes = mintTimes + 1
    Else
        mintTimes = 0
    End If
    
    If mintTimes >= glngAtuoErr Then
        Call cmdIgnore_Click
    Else
        Call cmdRetry_Click
    End If
End Sub

