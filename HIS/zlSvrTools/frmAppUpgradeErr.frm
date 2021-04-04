VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{CA73588D-282F-4592-9369-A61CC244FADA}#15.3#0"; "Codejock.SyntaxEdit.v15.3.1.ocx"
Begin VB.Form frmAppUpgradeErr 
   AutoRedraw      =   -1  'True
   Caption         =   "����"
   ClientHeight    =   7845
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   11250
   Icon            =   "frmAppUpgradeErr.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   11250
   StartUpPosition =   2  '��Ļ����
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
         Caption         =   "����������"
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
         Text            =   "Drop Index ���������¼_IX_��ҳID;"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "΢���ź�"
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
         Caption         =   "ִ�н����"
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
         Caption         =   "��ֹ(&A)"
         Height          =   350
         Left            =   9876
         TabIndex        =   3
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdIgnore 
         Caption         =   "����(&I)"
         Height          =   350
         Left            =   8772
         TabIndex        =   2
         Tag             =   "8175"
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdRetry 
         Caption         =   "����(&R)"
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
'������Ϣ
Private Type ErrInfo
    ErrNum              As Long
    ErrDesc             As String
    ErrAdvice           As String
    ErrAdviceType       As Integer
    ErrPos              As String
    ErrSQL              As String
    ErrOparate          As VbMsgBoxResult
    ErrOparateModi      As VbMsgBoxResult   '��д����SQL��Ľ���
    ErrModiSQL          As String           '���������SQL
End Type
Private merrCur         As ErrInfo
Private mstrUser        As String
Private mcnThis         As ADODB.Connection
Private mfrmParent      As Object
Private mobjSQL         As clsSQLInfo
Private mobjPreSQL      As clsSQLInfo '��һ��SQL
Private mblnIgnoreErr   As Boolean
Private mclsrun         As clsRunScript
Private mintTimes       As Long
Private mblnAuto        As Boolean '�Ƿ���Timerʱ���Զ�ִ��
Private mblnModify      As Boolean '��������ģʽ
Private mblnShut        As Boolean  '�ж��Ƿ�ֱ�ӹر�
Public Function ShowError(ByVal cnThis As ADODB.Connection, ByVal lngErrNum As Long, ByVal strErrInfo As String, ByVal objSQL As clsSQLInfo, frmParent As Object, Optional ByVal blnIgnoreErr As Boolean = True, Optional ByRef blnSysIgnore As Boolean, Optional ByVal clsRun As clsRunScript, Optional ByRef blnErrRepaired As Boolean) As VbMsgBoxResult
'blnErrRepaired=�����Ƿ񱻸ô����Զ��޸�������޸��ˣ����治��д��־
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
    '�жϴ�����Ƿ�����Զ�����
    If blnIgnoreErr Then
        If merrCur.ErrOparate = vbIgnore Then
            ShowError = merrCur.ErrOparate
            blnSysIgnore = True        'ϵͳ�������
            Unload Me
            Exit Function
        '��������SQL�󣬽������
        ElseIf ModifyErrors(True) Then 'ִ������SQL�ɹ������Զ�����
            blnErrRepaired = True
            merrCur.ErrOparate = vbIgnore
            ShowError = merrCur.ErrOparate
            blnSysIgnore = True        'ϵͳ�������
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
    If MsgBox("ϵͳ����Ҫ�����ؽ�����Ǩ֮���������ʹ�á�ȷʵҪ��ֹ��Ǩ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    merrCur.ErrOparate = vbAbort
    mblnShut = False
    Unload Me
End Sub

Private Sub cmdIgnore_Click()
    '�Ա����������������
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
    '�Ա����������������
    If cmdRetry.Tag = "1" Then Exit Sub
    cmdIgnore.Tag = "1"
    'û�д�����ֱ������
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
    '��ȡ���ڷ����ı�׼SQL��
    arrFMT = Split(Replace(Replace(strSQL, vbCrLf, vbCr), vbLf, vbCr), vbCr)
    For i = 0 To UBound(arrFMT)
        strReturn = strReturn & " " & TrimComment(arrFMT(i))
    Next
    strReturn = UCase(TrimEx(strReturn))
    GetFormatSQL = strReturn
End Function

Private Function ModifyErrors(Optional ByVal blnErrAutoAdjust As Boolean) As Boolean
'���ܣ�ִ������SQL
    Dim strSQL As String, strErr As String
    Dim strLine As String, i As Long
    Dim strLogSQL As String
    Dim objScript As clsRunScript
    Dim blnHaveErr As Boolean, blnHaveSQL As Boolean
    Dim lngAffect As Long
    Dim strOldModfy As String
    Dim datBegin As Date, datEnd As Date, lngSQLTime As Long
    
    'ִ�������ű���
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
                    lblSQLErr.Caption = "����ִ��SQL(" & objScript.Line & "��)��" & strLogSQL
                    If Not mclsrun Is Nothing Then mclsrun.WriteLog String(17, " ") & "��������(�˹�)����SQL��" & strLogSQL
                    err.Clear
                    datBegin = Now: datEnd = Now
                    DoEvents
                    mcnThis.Execute objScript.SQLInfo.SQL, lngAffect, adCmdText
                    If err.Number <> 0 Then
                        If mcnThis.Errors.Count > 0 Then
                            lblSQLErr.Caption = lblSQLErr.Caption & vbNewLine & "ִ�г���������Ϣ��" & mcnThis.Errors(0).Description
                            If Not mclsrun Is Nothing Then mclsrun.WriteLog String(17, " ") & "��������(�˹�)����" & mcnThis.Errors(0).Description
                        Else
                             lblSQLErr.Caption = lblSQLErr.Caption & vbNewLine & "ִ�г���������Ϣ��" & err.Description
                            If Not mclsrun Is Nothing Then mclsrun.WriteLog String(17, " ") & "��������(�˹�)����" & err.Description
                        End If
                        blnHaveErr = True: err.Clear
                        Exit Do
                    Else
                        lblSQLErr.Caption = lblSQLErr.Caption & vbNewLine & "ִ�гɹ���"
                        If Not mclsrun Is Nothing Then mclsrun.WriteLog String(17, " ") & "��������(�˹�)�����ִ�гɹ�" & IIf(lngAffect > 0, "," & lngAffect & " ��������Ч", ",0��������Ч")
                        lblSQLErr.Tag = objScript.Line
                    End If
                    If mclsrun.SQLRecTime <> 0 Then
                        lngSQLTime = DateDiff("n", datBegin, datEnd)
                        If lngSQLTime >= mclsrun.SQLRecTime Then
                            mclsrun.WriteLog String(17, " ") & "��������(�˹�)SQL�����ʱ��" & lngSQLTime & "����"
                        End If
                    End If
                    objScript.ReadNextSQL
                Loop
            End If
        End If
    'showErr�Զ�ִ��SQL
    Else
        If merrCur.ErrModiSQL = "" Then
            Exit Function
        End If
        strOldModfy = merrCur.ErrModiSQL
        'ִ������SQL
        strSQL = merrCur.ErrModiSQL
        Set objScript = New clsRunScript
        If objScript.AnalysisSQLString(strSQL, 0) Then
            On Error Resume Next
            Do While Not objScript.EOF
                strLogSQL = GetLogSQL(objScript.SQLInfo): blnHaveSQL = True
                If Not mclsrun Is Nothing Then mclsrun.WriteLog String(17, " ") & "��������(ϵͳ)����SQL��" & strLogSQL
                err.Clear
                datBegin = Now: datEnd = Now
                DoEvents
                mcnThis.Execute objScript.SQLInfo.SQL, lngAffect, adCmdText
                If err.Number <> 0 Then
                    If mcnThis.Errors.Count > 0 Then
                        If Not mclsrun Is Nothing Then mclsrun.WriteLog String(17, " ") & "��������(ϵͳ)����" & mcnThis.Errors(0).Description
                    Else
                        If Not mclsrun Is Nothing Then mclsrun.WriteLog String(17, " ") & "��������(ϵͳ)����" & err.Description
                    End If
                    blnHaveErr = True: err.Clear
                    Exit Do
                Else
                    If Not mclsrun Is Nothing Then mclsrun.WriteLog String(17, " ") & "��������(ϵͳ)�����ִ�гɹ�" & IIf(lngAffect > 0, "," & lngAffect & " ��������Ч", ",0��������Ч")
                End If
                If mclsrun.SQLRecTime <> 0 Then
                    lngSQLTime = DateDiff("n", datBegin, datEnd)
                    If lngSQLTime >= mclsrun.SQLRecTime Then
                        mclsrun.WriteLog String(17, " ") & "��������(ϵͳ)SQL�����ʱ��" & lngSQLTime & "����"
                    End If
                End If
                objScript.ReadNextSQL
            Loop
        End If
        If Not blnHaveErr Then
            'ִ������������ԭ����SQL,�����ﴦ����ֹ������ѭ��
            If merrCur.ErrOparateModi = vbRetry Then
                If err.Number <> 0 Then err.Clear
                mclsrun.WriteLog String(17, " ") & "��������(ϵͳ)������ִ������SQL������ԭSQL"
                err.Clear: On Error Resume Next
                datBegin = Now: datEnd = Now
                DoEvents
                mcnThis.Execute mobjSQL.SQL, lngAffect, adCmdText
                datEnd = Now
                If err.Number = 0 Then
                    mclsrun.WriteLog String(17, " ") & "��������(ϵͳ)���:�ɹ�" & IIf(lngAffect > 0, "," & lngAffect & " ��������Ч", ",0��������Ч")
                    If mclsrun.SQLRecTime <> 0 Then
                        lngSQLTime = DateDiff("n", datBegin, datEnd)
                        If lngSQLTime >= mclsrun.SQLRecTime Then
                            mclsrun.WriteLog String(17, " ") & "��������(ϵͳ)SQL�����ʱ��" & lngSQLTime & "����"
                        End If
                    End If
                    blnHaveErr = False
                Else
                    '�ٴβ鿴������
                    If mcnThis.Errors.Count > 0 Then
                        merrCur.ErrNum = mcnThis.Errors(0).NativeError
                        merrCur.ErrDesc = mcnThis.Errors(0).Description
                    Else
                        merrCur.ErrNum = err.Number
                        merrCur.ErrDesc = err.Description
                    End If
                    merrCur.ErrDesc = Replace(merrCur.ErrDesc, "[Microsoft][ODBC driver for Oracle][Oracle]", "")
                    merrCur.ErrModiSQL = ""
                    mclsrun.WriteLog String(17, " ") & "��������(ϵͳ)���:" & merrCur.ErrDesc
                    Call GetAdviceFromError
                     '��������SQL���ܽ�����⣬���Զ���գ������Ϊԭ����
                    If strOldModfy = merrCur.ErrModiSQL Then
                        merrCur.ErrModiSQL = ""
                        merrCur.ErrOparateModi = merrCur.ErrOparate
                    End If
                    '�ٴμ�鷢�ֱ�Ϊ�����Զ�����
                    If merrCur.ErrOparate = vbIgnore Then
                        blnHaveErr = False
                    Else
                        blnHaveErr = True
                    End If
                End If
            '������������ԭʼSQL
            ElseIf merrCur.ErrOparateModi = vbIgnore Then
                blnHaveErr = False
            End If
        Else
            '��������SQL���ܽ�����⣬���Զ���գ������Ϊԭ����
            merrCur.ErrModiSQL = ""
            merrCur.ErrOparateModi = merrCur.ErrOparate
        End If
    End If
    'û�д�����ֱ������
    If blnHaveErr Then
    ElseIf mblnModify And Not blnHaveSQL Then
        blnHaveErr = True
    End If
    If Not blnErrAutoAdjust Then Call RefreshButton
    ModifyErrors = Not blnHaveErr
End Function

Private Sub RefreshButton()
'���ܣ�ˢ�°�ť��ʾ����
    Dim blnDo As Boolean
    blnDo = Trim(synModiSQL.RowText(Val(synModiSQL.Tag) + 1)) <> "" Or synModiSQL.RowsCount > Val(synModiSQL.Tag) + 1
    lblSQLErr.Visible = blnDo
    synModiSQL.Height = lblSQLErr.Top - 60 - synModiSQL.Top + IIf(blnDo, 0, lblSQLErr.Height)
    If blnDo Then
        cmdRetry.Width = 2800
        cmdIgnore.Width = 2800
        cmdRetry.Caption = "ִ�������ű�������ԭ�ű�(&R)"
        cmdIgnore.Caption = "ִ�������ű�������ԭ�ű�(&I)"
        mblnModify = True
    Else
        cmdRetry.Width = 1100
        cmdIgnore.Width = 1100
        cmdRetry.Caption = "����(&R)"
        cmdIgnore.Caption = "����(&I)"
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
    
    '��ʾ����ִ���û�
    Me.Caption = "���� - " & mstrUser
    '������λ�ô���
    Me.Left = mfrmParent.Left + (mfrmParent.Width - Me.Width) / 2
    Me.Top = mfrmParent.Top + (mfrmParent.Height - Me.Height) / 2
    strTemp = GetSetting("ZLSOFT", "˽��ģ��\" & gstrUserName & "\��������\" & App.ProductName & Me.name & "\Form", "���λ��", "")
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
        merrCur.ErrPos = "�ļ���" & mobjSQL.File & "  " & "�кţ�" & mobjSQL.FileLine
    End If
    
    rtfErr.Text = ""
    '�����Ϣ
    rtfErr.Text = rtfErr.Text & "����š�": rtfErr.SelStart = 1: rtfErr.SelLength = Len("����š�"): rtfErr.SelFontSize = 14: rtfErr.SelBold = True
    rtfErr.Text = rtfErr.Text & merrCur.ErrNum & vbNewLine: strPos = Len(rtfErr.Text)
    '������Ϣ
    rtfErr.Text = rtfErr.Text & "����Ϣ��": rtfErr.SelStart = Len(rtfErr.Text) + 1: rtfErr.SelLength = Len("����Ϣ��"): rtfErr.SelFontSize = 14: rtfErr.SelBold = True
    rtfErr.Text = rtfErr.Text & merrCur.ErrDesc & vbNewLine: strPos = strPos & "," & Len(rtfErr.Text)
    '������Ϣ
    rtfErr.Text = rtfErr.Text & "�����顿": rtfErr.SelStart = Len(rtfErr.Text) + 1: rtfErr.SelLength = Len("�����顿"): rtfErr.SelFontSize = 14: rtfErr.SelBold = True
    rtfErr.Text = rtfErr.Text & merrCur.ErrAdvice & vbNewLine: strPos = strPos & "," & Len(rtfErr.Text)
    '������Ϣ
    rtfErr.Text = rtfErr.Text & "��λ�á�": rtfErr.SelStart = Len(rtfErr.Text) + 1: rtfErr.SelLength = Len("��λ�á�"): rtfErr.SelFontSize = 14: rtfErr.SelBold = True
    rtfErr.Text = rtfErr.Text & merrCur.ErrPos & vbNewLine: strPos = strPos & "," & Len(rtfErr.Text)
    'SQL��Ϣ
    rtfErr.Text = rtfErr.Text & "��SQL ��": rtfErr.SelStart = Len(rtfErr.Text) + 1: rtfErr.SelLength = Len("��λ�á�"): rtfErr.SelFontSize = 14: rtfErr.SelBold = True
    rtfErr.Text = rtfErr.Text & merrCur.ErrSQL: strPos = strPos & "," & Len(rtfErr.Text)
    
    '�﷨�ؼ���ɫ����
    synModiSQL.Font.name = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\SQLFont", "FontName", "Fixedsys")
    synModiSQL.Font.Size = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\SQLFont", "FontSize", 12)
    synModiSQL.Font.Underline = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\SQLFont", "FontUnderline", 0)
    synModiSQL.Font.Italic = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\SQLFont", "FontItalic", 0)
    synModiSQL.Font.Bold = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\SQLFont", "FontBold", 0)
    synModiSQL.Font.Strikethrough = GetSetting("ZLSOFT", "����ģ��\" & App.ProductName & "\SQLFont", "FontStrikethru", 0)
    synModiSQL.BorderStyle = xtpBorderClientEdge
    
    '���ÿؼ�����ʾ��ɫ����Ϊ��SQL
    If Not gblnInIDE Then '���Ӷ໷��֧��
        strPath = App.Path & "\PUBLIC\_sql.schclass"
    Else
        strPath = objFSO.GetParentFolderName(GetSetting("ZLSOFT", "����ȫ��", "����·��")) & "\PUBLIC\_sql.schclass"
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
    '�����Сλ�õ���
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
    'picBottom��λ
    picBottom.Top = Me.ScaleHeight - picBottom.Height - 30
    If Me.WindowState <> vbMinimized Then
        picBottom.ScaleWidth = Me.ScaleWidth
    End If
    '�Ȱ��������ָ���,�Լ��·�PIC��λ��
    fraSplit(1).Top = picBottom.Top - fraSplit(1).Height - 15
    picModify.Top = fraSplit(0).Top + fraSplit(0).Height + 30
    '������������pic�Ѿ��ָ��ߵĸ߶�����
    picErrInfo.Height = fraSplit(0).Top - picErrInfo.Top - 30
    picModify.Height = fraSplit(1).Top - picModify.Top - 30
    picErrInfo.Width = Me.ScaleWidth - picErrInfo.Left
    picModify.Width = Me.ScaleWidth - picModify.Left
    fraSplit(0).Width = Me.ScaleWidth - fraSplit(0).Left
    fraSplit(1).Width = Me.ScaleWidth - fraSplit(1).Left
    '�����ײ���ťλ��
    Call RefreshButton
    If err.Number <> 0 Then err.Clear
    On Error GoTo 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strTemp As String
    
    tmrRefresh.Enabled = False
    If mblnShut Then 'ֱ�ӹرմ��壬Ĭ�ϲ�ȡ��ֹ����
        If MsgBox("ϵͳ����Ҫ�����ؽ�����Ǩ֮���������ʹ�á�ȷʵҪ��ֹ��Ǩ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1
            Exit Sub
        Else
            merrCur.ErrOparate = vbAbort
        End If
    End If
    If glngAtuoErr > 0 Then Set mobjPreSQL = mobjSQL.CopyMe
    Set mobjSQL = Nothing
    '���洰��λ��
    strTemp = Me.Left - mfrmParent.Left & "," & Me.Top - mfrmParent.Top
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrUserName & "\��������\" & App.ProductName & Me.name & "\Form", "���λ��", strTemp
End Sub

Private Sub GetAdviceFromError()
'���ܣ����ݲ�ͬ��Oracle����������Ӧ�Ĳ�������˵��
    Dim strSQL As String
    Dim strOwner As String, strName As String, strType As String
    Dim strTemp As String

    '����ȷ����Oracle�Ĵ���Ž��д���
    If InStr(merrCur.ErrDesc, "ORA-") = 0 Then
        merrCur.ErrAdvice = "��Ǩ�����ڲ������볢�����Բ���"
        merrCur.ErrOparate = vbRetry
        Exit Sub
    End If
    '�����
    If mobjSQL.Block Then
        merrCur.ErrAdvice = "����ϸ��飬������Ҫ�Ĵ���������Բ���"
        merrCur.ErrOparate = vbRetry
        Select Case merrCur.ErrNum
            Case 1 'ORA-00001: Υ��ΨһԼ������ (ZLTOOLS.XXX_PK)(ZLTOOLS.XXX_UQ_YYY)
                CheckAdjustSequence (merrCur.ErrDesc)
            Case 1502
                'ORA-01502: index 'XXXX' or partition of such index is in unusable state
                'ORA-01502: ���� 'XXXX' �����������ķ������ڲ�����״̬
                merrCur.ErrAdvice = "���ؽ����������ԡ�"
                merrCur.ErrOparate = vbRetry
                Call CheckAdjustIndex(merrCur.ErrDesc)
            Case 6575
                'ORA-06575: ��������� TTT ������Ч״̬
                'ORA-06575: Package or function TTT is in an invalid state
                merrCur.ErrAdvice = "������ȷ�����Ӧ�ĺ�������̺������ԡ�"
                merrCur.ErrOparate = vbRetry
                Call CheckAdjustProcedure(merrCur.ErrDesc)
            Case 12899 'ORA-12899: �� "ZLHIS"."���������Ա"."��������" ��ֵ̫�� (ʵ��ֵ: 62, ���ֵ: 60)
                'ORA-12899: value too large for column "SYSTEM"."STUDENTINFO"."SNAME" (actual: 78, maximum: 30)
                merrCur.ErrAdvice = "���ȵ����ֶε����ʵľ��Ⱥ������ԡ�"
                merrCur.ErrOparate = vbRetry
                Call CheckAdjust12899(merrCur.ErrDesc)
        End Select
    Else
        Select Case merrCur.ErrNum
            Case 1 'ORA-00001: Υ��ΨһԼ������ (ZLTOOLS.XXX_PK)(ZLTOOLS.XXX_UQ_YYY)
                merrCur.ErrOparate = vbRetry
                If mobjSQL.PartSQL Like "INSERT INTO *" Then
                    Call CheckAdjustSequence(merrCur.ErrDesc)
                    Call CheckAdjustTableData(mobjSQL.SQL, IIf(mobjSQL.PartSQL Like "INSERT INTO *VALUES*", 0, 1), merrCur.ErrDesc)
                ElseIf mobjSQL.PartSQL Like "UPDATE *" Then
                    merrCur.ErrAdvice = "����Ҫ�������ݵ���ȷ�ԡ�"
                Else
                    merrCur.ErrAdvice = "����������ظ����г������顣"
                End If
            Case 955 'ORA-00955: �����ѱ����ж���ռ��
                strSQL = GetFormatSQL(mobjSQL.SQL)
                Call GetCreateName(strSQL, strOwner, strName, strType)
                merrCur.ErrAdvice = "����������ظ����г���һ������¿��Ժ��Ըô���"
                merrCur.ErrOparate = vbIgnore
                If strName <> "" Then
                    If Not ObjectExists(strOwner, strName, strType) Then
                        merrCur.ErrAdvice = "SQL���Ҫ�����Ķ����ѱ��������͵�ͬ������ռ�ã����ֹ���������ԡ�"
                        merrCur.ErrOparate = vbRetry
                    ElseIf strType = "TABLE" And Not strSQL Like "* AS SELECT *" Then  'Create Table XXX As Select��ʽ��������ϸ����
                        strTemp = CheckCreateTabCol(strSQL, strOwner, strName)
                        If strTemp <> "" Then
                            merrCur.ErrAdvice = "SQL���Ҫ�����Ķ����ѱ��������͵�ͬ������ռ��,�������߽ṹ���ڲ��죬���ֹ���������ԡ�"
                            merrCur.ErrOparate = vbRetry
                        End If
                    End If
                Else
                    Call CheckAdjustConstraint(strSQL)
                End If
            Case 1430 'ORA-01430: �����Ѿ�����Ҫ��ӵ���
                strSQL = GetFormatSQL(mobjSQL.SQL)
                merrCur.ErrAdvice = "����������ظ����г���һ������¿��Ժ��Ըô���"
                merrCur.ErrOparate = vbIgnore
                Call CheckTabChangeCol(strSQL)
            Case 2260, 2261
                strSQL = GetFormatSQL(mobjSQL.SQL)
                'ORA-02260: ��ֻ�ܾ���һ�����ؼ���
                'ORA-02261: �����Ѵ���������Ψһ�ؼ��ֻ����ؼ���
                merrCur.ErrAdvice = "����������ظ����г���һ������¿��Ժ��Ըô���"
                merrCur.ErrOparate = vbIgnore
                Call CheckUniqueKeyCol(strSQL)
            Case 2291
                'ORA-02291: Υ������Լ������ (ZLTOOLS.XXX_FK_YYID) - δ�ҵ�����ؼ���
                If mobjSQL.PartSQL Like "INSERT INTO ZLRPT*" Then
                    merrCur.ErrAdvice = "����������ظ����г���һ������¿��Ժ��Ըô���"
                    merrCur.ErrOparate = vbIgnore
                Else
                    merrCur.ErrAdvice = "������Ҫ��ĸ������ݵ���ȷ�ԡ�"
                    merrCur.ErrOparate = vbRetry
                End If
            Case 2298 '�޷���֤ (ZLHIS.XXX_FK_YY) - δ�ҵ�����ؼ���
                'ORA-02298: �޷���֤ (ZLHIS.XXX_FK_YY) - δ�ҵ�����ؼ���
                If mobjSQL.PartSQL Like "ALTER TABLE * ADD CONSTRAINT * FOREIGN KEY *" Then
                    merrCur.ErrAdvice = "���������ݴ��󣬸�������ȱʧ��"
                    merrCur.ErrModiSQL = mobjSQL.SQL & " enable novalidate;"
                    merrCur.ErrOparateModi = vbRetry
                    merrCur.ErrOparate = vbRetry
                Else
                    merrCur.ErrAdvice = "������Ҫ��ĸ������ݵ���ȷ�ԡ�"
                    merrCur.ErrOparate = vbRetry
                End If
            Case 957
                'ORA-00957:�ظ�������
                If mobjSQL.PartSQL Like "ALTER TABLE * RENAME COLUMN * TO *" Then
                    merrCur.ErrAdvice = "����������ظ����г���һ������¿��Ժ��Ըô���"
                    merrCur.ErrOparate = vbIgnore
                ElseIf strSQL Like "CREATE TABLE*" Then
                    merrCur.ErrAdvice = "����SQL�ű�����ȷ�ԡ�"
                    merrCur.ErrOparate = vbRetry
                End If
            Case 904
                'ORA-00904: ��Ч����
                strSQL = GetFormatSQL(mobjSQL.SQL)
                If mobjSQL.PartSQL Like "ALTER TABLE * DROP COLUMN *" Then
                    merrCur.ErrAdvice = "����������ظ����г���һ������¿��Ժ��Ըô���"
                    merrCur.ErrOparate = vbIgnore
                Else
                    merrCur.ErrAdvice = "���Ȳ�����ȷ�ı����ֶκ������ԡ�"
                    merrCur.ErrOparate = vbRetry
                End If
            Case 942
                strSQL = GetFormatSQL(mobjSQL.SQL)
                'ORA-00942: �����ͼ������
                If mobjSQL.PartSQL Like "DROP TABLE*" Or mobjSQL.PartSQL Like "DROP VIEW*" Or mobjSQL.PartSQL Like "DROP MATERIALIZED VIEW*" Then
                    merrCur.ErrAdvice = "����������ظ����г���һ������¿��Ժ��Ըô���"
                    merrCur.ErrOparate = vbIgnore
                Else
                    merrCur.ErrAdvice = "���Ȳ��䴴����Ӧ�ı����ͼ�������ԡ�"
                    merrCur.ErrOparate = vbRetry
                End If
            Case 1442
                'ORA-01442: Ҫ�޸�Ϊ NOT NULL �����Ѿ��� NOT NULL
                If mobjSQL.PartSQL Like "ALTER TABLE * MODIFY * CONSTRAINT * NOT NULL*" Then
                    merrCur.ErrAdvice = "����������ظ����г���һ������¿��Ժ��Ըô���"
                    merrCur.ErrOparate = vbIgnore
                Else
                    merrCur.ErrAdvice = "�����޸����Ƿ��Ѿ��޸ġ�"
                    merrCur.ErrOparate = vbRetry
                End If
            Case 8002
                'ORA-08002: ����ZLRPTDATAS_ID.CURRVAL ��δ�ڴ˽����ж���
                If mobjSQL.PartSQL Like "INSERT INTO ZLRPT*" Then
                    merrCur.ErrAdvice = "����������ظ����г���һ������¿��Ժ��Ըô���"
                    merrCur.ErrOparate = vbIgnore
                Else
                    merrCur.ErrAdvice = "����֮ǰ���������Ƿ���ȷ���С�"
                    merrCur.ErrOparate = vbRetry
                End If
                Call CheckAdjustSequnceVali(merrCur.ErrDesc)
            '-----------------------------------------------------------------------
            Case 1418
                strSQL = GetFormatSQL(mobjSQL.SQL)
                'ORA-01418: ָ��������������
                Call CheckErr1418(strSQL)
            Case 4043, 4080, 2289
                'ORA-04043: ���� XXX ������
                'ORA-04080: ������ 'XXX' ������
                'ORA-02289: ���У��ţ�������
                If mobjSQL.PartSQL Like "DROP *" Then
                    merrCur.ErrAdvice = "����������ظ����г���һ������¿��Ժ��Ըô���"
                    merrCur.ErrOparate = vbIgnore
                Else
                    merrCur.ErrAdvice = "���Ȳ��䴴����Ӧ�Ķ�������ԡ�"
                    merrCur.ErrOparate = vbRetry
                End If
            Case 1432, 1434, 1543, 1919, 1921, 1927, 1952, 2264, 2275, 2443, 4081, 12003, 12006
                'ORA-01432: Ҫɾ���Ĺ���ͬ��ʲ�����
                'ORA-01434: Ҫɾ��������ͬ��ʲ�����
                'ORA-01543: ��ռ�'XXX'�Ѿ�����
                'ORA-01919: ����'XXX'������
                'ORA-01921: ������'XXX'����һ���û�����������������ͻ
                'ORA-01927: �޷� REVOKE ��δ��Ȩ��Ȩ��
                'ORA-01952: ϵͳȨ��δ����'ZLHIS'
                'ORA-02264: �����ѱ�һ����Լ������ռ��
                'ORA-02275: �˱����Ѿ���������������Լ������
                'ORA-02443: �޷�ɾ��Լ������ - ������Լ������
                'ORA-04081: ������ 'XXX' �Ѿ�����
                'ORA-12003: ʵ�廯��ͼ(����) "SYS"."TESTVIEW" ������
                'ORA-12006: ������ͬ�û����Ŀ����Ѿ�����
                merrCur.ErrAdvice = "����������ظ����г���һ������¿��Ժ��Ըô���"
                merrCur.ErrOparate = vbIgnore
            '-----------------------------------------------------------------------
            Case 900, 907, 936
                'ORA-00900: ��Ч SQL ���
                'ORA-00907: ȱ��������
                'ORA-00936: ȱ�ٱ��ʽ
                merrCur.ErrAdvice = "����SQL�ű�����ȷ�ԡ�"
                merrCur.ErrOparate = vbRetry
            Case 959
                'ORA-00959: ��ռ�'XXX'������
                merrCur.ErrAdvice = "���Ȳ��䴴����Ӧ�ı�ռ�������ԡ�"
                merrCur.ErrOparate = vbRetry
            Case 1031
                'ORA-01031: Ȩ�޲���
                merrCur.ErrAdvice = "���������ݿ������赱ǰ�û���Ӧ��ɫȨ�޺������ԡ�"
                merrCur.ErrOparate = vbRetry
            Case 1400, 1407, 2290
                'ORA-01400: �޷��� NULL ���� ("ZLHIS"."XXX"."YYY")
                'ORA-01407: �޷����� ("ZLHIS"."XXX"."YYY") Ϊ NULL
                'ORA-02290: Υ�����Լ������ (ZLHIS.����_CK_ȱʡ��־)
                merrCur.ErrAdvice = "���ȼ�鲢�������Լ�����������������ԡ�"
                merrCur.ErrOparate = vbRetry
            Case 1408
                'ORA-01408: �����б�������
                strSQL = GetFormatSQL(mobjSQL.SQL)
                merrCur.ErrAdvice = "���ȼ�鲢�������Լ�����������������ԡ�"
                merrCur.ErrOparate = vbRetry
                Call CheckIndexCol(strSQL)
            Case 1401, 1438
                'ORA-01401: �����ֵ�����й���
                'ORA-01438: ֵ���ڴ���ָ��������ȷ��
                merrCur.ErrAdvice = "���ȵ����ֶε����ʵľ��Ⱥ������ԡ�"
                merrCur.ErrOparate = vbRetry
            Case 1439, 1440
                'ORA-01439: Ҫ�����������ͣ���Ҫ�޸ĵ��б���Ϊ�� (empty)
                'ORA-01440: Ҫ��С��ȷ�Ȼ��ȣ���Ҫ�޸ĵ��б���Ϊ�� (empty)
                merrCur.ErrAdvice = "���ȱ��ݶ�Ӧ�ı������ݣ�����պ������ԡ�"
                merrCur.ErrOparate = vbRetry
            Case 1502
                'ORA-01502: index 'XXXX' or partition of such index is in unusable state
                'ORA-01502: ���� 'XXXX' �����������ķ������ڲ�����״̬
                merrCur.ErrAdvice = "���ؽ����������ԡ�"
                merrCur.ErrOparate = vbRetry
                Call CheckAdjustIndex(merrCur.ErrDesc)
            Case 1775
                strSQL = GetFormatSQL(mobjSQL.SQL)
                'ORA-01775: ͬ��ʵ�ѭ����
                merrCur.ErrAdvice = "ͬ�������Ӧ�Ķ��󲻴��ڣ�������Ƿ��������ű�ɾ������ʱû��ɾ��ͬ��������µģ���������ԡ�"
                merrCur.ErrOparate = vbRetry
                Call CheckErr1775(strSQL)
            Case 2270
                strSQL = GetFormatSQL(mobjSQL.SQL)
                'ORA-02270: �����б��Ψһ��������ƥ��
                merrCur.ErrAdvice = "���Ҫ������������������ֶβ���������Ψһ���������ԡ�"
                merrCur.ErrOparate = vbRetry
                Call CheckErr2270(strSQL)
            Case 2273
                'ORA-02273: ��Ψһ/�����ѱ�ĳЩ�ⲿ�ؼ�������
                merrCur.ErrAdvice = "����ɾ���ӱ�����Լ�����ú������ԡ�"
                merrCur.ErrOparate = vbRetry
            Case 2091, 2292 '2091���ӳ�Լ��
                'ORA-02292: Υ������Լ������ (ZLHIS.XXX_FK_YYY) - ���ҵ��Ӽ�¼��־
                merrCur.ErrAdvice = "����ɾ���ӱ�Ĺ������ݺ������ԡ�"
                merrCur.ErrOparate = vbRetry
            Case 2303
                'ORA-02303: �޷�ʹ�����ͻ����������ɾ����ȡ��һ������
                merrCur.ErrAdvice = "����ȡ�����ø����͵���ر��л�����֮�������ԡ�"
                merrCur.ErrOparate = vbRetry
            Case 2299, 2437
                'ORA-02437: �޷���֤ (ZLHIS.XXX_PK) - Υ������
                'ORA-02299: �޷���֤ (ZLHIS.XXX_UQ_YYY) - δ�ҵ��ظ��ؼ���
                merrCur.ErrAdvice = "��Ա��ж�Ӧ�ֶε����ݽ����ظ���鴦��������ԡ�"
                merrCur.ErrOparate = vbRetry
            Case 6575
                'ORA-06575: ��������� TTT ������Ч״̬
                'ORA-06575: Package or function TTT is in an invalid state
                merrCur.ErrAdvice = "������ȷ�����Ӧ�ĺ�������̺������ԡ�"
                merrCur.ErrOparate = vbRetry
                Call CheckAdjustProcedure(merrCur.ErrDesc)
            Case 6576
                'ORA-06576: ������Ч�ĺ����������
                merrCur.ErrAdvice = "������ȷ������Ӧ�ĺ�������̺������ԡ�"
                merrCur.ErrOparate = vbRetry
            Case 12899 'ORA-12899: �� "ZLHIS"."���������Ա"."��������" ��ֵ̫�� (ʵ��ֵ: 62, ���ֵ: 60)
                'ORA-12899: value too large for column "SYSTEM"."STUDENTINFO"."SNAME" (actual: 78, maximum: 30)
                merrCur.ErrAdvice = "���ȵ����ֶε����ʵľ��Ⱥ������ԡ�"
                merrCur.ErrOparate = vbRetry
                Call CheckAdjust12899(merrCur.ErrDesc)
            Case 19001
                strSQL = GetFormatSQL(mobjSQL.SQL)
                'ORA-19001ָ���Ĵ洢ѡ����Ч
                If strSQL Like "CREATE TABLE*STORE AS SECUREFILE BINARY XML*" Then
                    If GetOracleVersion(True, True) < 11 Then 'racle�汾����11gʱ
                        merrCur.ErrAdvice = "Oracle�汾����11G��֧�ָô洢ѡ������еͰ汾�Ľṹ������䣬���Ժ��ԡ�"
                        merrCur.ErrOparate = vbIgnore
                    Else
                        merrCur.ErrAdvice = "����SQL�Ƿ���д��ȷ��ǰOracle�汾�Ƿ�֧�ָô洢ѡ�"
                        merrCur.ErrOparate = vbRetry
                    End If
                Else
                    merrCur.ErrAdvice = "����SQL�Ƿ���д��ȷ��ǰOracle�汾�Ƿ�֧�ָô洢ѡ�"
                    merrCur.ErrOparate = vbRetry
                End If
            Case 22858 'ORA-22858: �������͵ĸ�����Ч;һ�㽫��ͨ�����޸�Ϊ�����
                strSQL = GetFormatSQL(mobjSQL.SQL)
                Call CheckTabChangeCol(strSQL)
            Case 22859 'ORA-22859: ��Ч�����޸�;һ�㽫������޸�Ϊ��ͨ����
                strSQL = GetFormatSQL(mobjSQL.SQL)
                Call CheckTabChangeCol(strSQL)
            Case 23292 'ORA-23292: Լ������������
                strSQL = GetFormatSQL(mobjSQL.SQL)
                Call CheckErr23292(strSQL)
            Case Else
                merrCur.ErrAdvice = "����ϸ��飬������Ҫ�Ĵ���������Բ���"
                merrCur.ErrOparate = vbRetry
        End Select
    End If
End Sub

Private Sub CheckTabChangeCol(ByVal strSQL As String)
'���ܣ���������е����������ݿ��Ƿ�һ��
'������strSQL=�Ѹ�ʽ��Ϊ��׼��д��SQL���
'���أ�
    Dim rsTemp As New ADODB.Recordset
    Dim strOwner As String, strName As String
    Dim strCol As String, arrCol As Variant
    Dim strType As String, intMatch As Integer
    Dim intLen As Integer, intDigit As Integer
    Dim strError As String, i As Long
    Dim strModifySQL As String
    Dim blnModify As Boolean
    
    If Not (strSQL Like "ALTER TABLE * ADD*" Or strSQL Like "ALTER TABLE * MODIFY*") Then Exit Sub

    '����
    strName = Split(Mid(strSQL, InStr(strSQL, "ALTER TABLE ") + Len("ALTER TABLE ")), " ")(0)
    If InStr(strName, ".") > 0 Then
        strOwner = Split(strName, ".")(0)
        strName = Split(strName, ".")(1)
    End If

    'ȡ��SQL����е��ж���
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
                strCol = strCol & "|" '������������е�",",��Number(16,5)
            Else
                strCol = strCol & Mid(strSQL, i, 1)
            End If
        Next
    Else
        strCol = strSQL
    End If
    arrCol = Split(strCol, "|")

    '��SQL�е��ж��������ݿ��еĽ��бȽ�
    On Error Resume Next
    strSQL = "Select Column_Name,Data_Type,Data_Length,Data_Precision,Data_Scale From ALL_Tab_Columns" & _
        " Where OWNER=" & IIf(strOwner = "", "User", "'" & strOwner & "'") & " And Table_Name='" & strName & "'"
    rsTemp.CursorLocation = adUseClient
    rsTemp.Open strSQL, mcnThis, adOpenKeyset
    For i = 0 To UBound(arrCol)
        arrCol(i) = Trim(arrCol(i))

        strCol = Left(arrCol(i), InStr(arrCol(i), " ") - 1) '���� Number ( 16, 5) Not Null Default 1.23
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
                '������������
                End If
            Else
                strError = strError & "," & strCol
            End If
        End If
    Next
    strError = Mid(strError, 2)

    If strError <> "" Then
        merrCur.ErrAdvice = "����ȱ��SQL���Ҫ��ӵ������У����Ѿ����ڵ���������SQL��䲻�������ֹ���������ԡ�"
        merrCur.ErrOparate = vbRetry
        merrCur.ErrModiSQL = strModifySQL
        merrCur.ErrOparateModi = vbRetry
    Else
        merrCur.ErrAdvice = "�����Ѿ�����Ҫ��ӵ��л����������Ѿ��޸ģ���ȷ�Ϻ���ԡ�"
        merrCur.ErrOparate = vbIgnore
        merrCur.ErrModiSQL = ""
        merrCur.ErrOparateModi = vbIgnore
    End If
End Sub

Private Sub CheckErr1775(ByVal strSQL As String)
'ORA-01775: ͬ��ʵ�ѭ����
'���ܣ�������ɾ����"zlPDASynch"��"zlStreamTabs"
'������strSQL=�Ѹ�ʽ��Ϊ��׼��д��SQL���
'���أ�������ʾ����(strAdvice)��ȱʡ������ťֵ(intAdvice)
'�����߰汾9.41.0
'Drop Table zlPDASynch;
'drop table zlStreamTabs;
    If strSQL Like "* ZLPDASYNCH*" Or strSQL Like "* ZLSTREAMTABS*" Then
        merrCur.ErrAdvice = "�ö����Ѿ��ڹ����߰汾9.41.0��ɾ����"
        merrCur.ErrOparate = vbIgnore
    End If
End Sub

Private Sub CheckErr1418(ByVal strSQL As String)
'ORA-01418: ָ��������������
'���ܣ�ɾ�������Զ����ԣ��������������鿴�����Ƿ���ڣ��������Զ�����
'������strSQL=�Ѹ�ʽ��Ϊ��׼��д��SQL���
    Dim strIndexName As String, arrTmp As Variant
    merrCur.ErrAdvice = "���Ȳ��䴴����Ӧ�Ķ�������ԡ�"
    merrCur.ErrOparate = vbRetry
    If strSQL Like "DROP *" Then
        merrCur.ErrAdvice = "����������ظ����г���һ������¿��Ժ��Ըô���"
        merrCur.ErrOparate = vbIgnore
    ElseIf strSQL Like "ALTER INDEX * RENAME TO *" Then
        arrTmp = Split(strSQL, "RENAME TO")
        If UBound(arrTmp) < 1 Then Exit Sub
        strIndexName = UCase(Trim(Split(Trim(arrTmp(1)), " ")(0)))
        If strIndexName = "" Then Exit Sub
        If ObjectExists(UCase(mstrUser), strIndexName, "INDEX") Then
            merrCur.ErrAdvice = "����������ظ����г���һ������¿��Ժ��Ըô���"
            merrCur.ErrOparate = vbIgnore
        End If
    End If
End Sub

Private Sub CheckErr23292(ByVal strSQL As String)
'ORA -23292: Լ������������
'���ܣ�ɾ�������Զ����ԣ��������������鿴�����Ƿ���ڣ��������Զ�����
'������strSQL=�Ѹ�ʽ��Ϊ��׼��д��SQL���
    Dim strConName As String, arrTmp As Variant, strTableName As String
    Dim arrTmp1 As Variant
    Dim strOwner As String
    
    merrCur.ErrAdvice = "���Ȳ��䴴����Ӧ�Ķ�������ԡ�"
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
            merrCur.ErrAdvice = "����������ظ����г���һ������¿��Ժ��Ըô���"
            merrCur.ErrOparate = vbIgnore
        End If
    End If
End Sub

Private Sub CheckErr2270(ByVal strSQL As String)
'ORA-02270: �����б��Ψһ��������ƥ��
'���ܣ��������ڴ�������/Ψһ�������ǲ���������/Ψһ���������µĴ���
'������strSQL=�Ѹ�ʽ��Ϊ��׼��д��SQL���
'���أ�������ʾ����(strAdvice)��ȱʡ������ťֵ(intAdvice)
    Dim strTable As String, strRTable As String, strCols As String, strRCols As String
    Dim strColsInfo As String, strRColsInfo As String, strTmp As String, strPreCon As String
    Dim strOwner As String, strROwner As String
    Dim rsRTable As New ADODB.Recordset, rsTable As New ADODB.Recordset
    Dim cllColInfo As Collection
    Dim i As Long, arrTmp As Variant
    Dim strModifySQL As String
    
    '�ô���������ԭ��
    '1�����ñ��ֶ�δ����������Ψһ��
    '2�����ñ��ֶδ�����������Ψһ���ֶε����ͣ�������Ҫ����������ֶδ��ڲ���
    '3�����ñ������Ψһ��û������
    'Alter Table ����֧������ Add Constraint ����֧������_FK_����� Foreign Key (����,����,��ְ,�����) References ���������(����,����,��ְ,�����) On Delete Cascade;
    If Not strSQL Like "ALTER TABLE * ADD CONSTRAINT * FOREIGN KEY * REFERENCES *" Then Exit Sub
    '����SQL�е���Ϣ
    arrTmp = Split(strSQL, "ADD CONSTRAINT")
    strTable = Trim(Split(arrTmp(0), "ALTER TABLE")(1))
    arrTmp = Split(Split(arrTmp(1), "FOREIGN KEY")(1), "REFERENCES")
    strCols = UCase(Trim(Replace(Replace(Replace(arrTmp(0), "(", ""), ")", ""), " ", ""))) 'ȥ�����������Լ����еĿո�
    arrTmp = Split(Split(arrTmp(1), ")")(0), "(") '���������ָ��
    strRTable = Trim(arrTmp(0))
    strRCols = UCase(Trim(Replace(arrTmp(1), " ", "")))  'ȥ�����еĿո�
    
    If InStr(strTable, ".") > 0 Then
        strOwner = UCase(Split(strTable, ".")(0))
        strTable = UCase(Split(strTable, ".")(1))
    End If
    If InStr(strRTable, ".") > 0 Then
        strROwner = UCase(Split(strRTable, ".")(0))
        strRTable = UCase(Split(strRTable, ".")(1))
    End If
    
    '��ȡ���ñ������Ψһ����Ϣ�Լ����ֶ���Ϣ�Լ���Ӧ������
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
    Set rsRTable = gclsBase.OpenSQLRecord(mcnThis, strSQL, "������-��ȡ������Ϣ")
    '1�����ñ��ֶ�δ����������Ψһ��
    If rsRTable.RecordCount = 0 Then Exit Sub
    '2�����ñ��ֶδ�����������Ψһ���ֶε����ͣ�������Ҫ����������ֶδ��ڲ���
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
    '��������ֶ����Ƿ���������Ψһ��Լ��
    On Error Resume Next
    strTmp = cllColInfo(strRCols)
    If err.Number <> 0 Then
        'û�л�ȡ����ӦԼ��
        err.Clear: Exit Sub
    End If
    On Error GoTo 0
    '��ȡ��Լ������Ա��ֶ�����
    strSQL = "Select a.Column_Name, a.Data_Type, a.Data_Length, a.Data_Precision, a.Data_Scale" & vbNewLine & _
                    "From All_Tab_Columns A" & vbNewLine & _
                    "Where a.Owner " & IIf(strOwner <> "", "=" & strOwner, IIf(mstrUser = "ZLTOOLS", "= User", " In (User, 'ZLTOOLS')")) & " And a.Table_Name = '" & strTable & "' And a.Column_Name In ('" & Replace(strCols, ",", "','") & "')"
    Set rsTable = gclsBase.OpenSQLRecord(mcnThis, strSQL, "������-��ȡ�ӱ���Ϣ")
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
        Exit Sub '���Ͳ�ƥ��
    End If
    rsRTable.Filter = "Constraint_Name='" & strTmp & "'"
    If rsRTable!Index_Name & "" = "" Then
        '�����ж�ΨһԼ��������Լ��ͬ�������Ƿ����
        strSQL = "Select a.Status" & vbNewLine & _
                        "From All_Indexes A" & vbNewLine & _
                        "Where a.Table_Owner " & IIf(strROwner <> "", "=" & strROwner, IIf(mstrUser = "ZLTOOLS", "= User", " In (User, 'ZLTOOLS')")) & " And a.Uniqueness = 'UNIQUE' And a.Table_Name = [1] And a.Index_Name =[2]"
        Set rsTable = gclsBase.OpenSQLRecord(mcnThis, strSQL, "������-��ȡ�ӱ���Ϣ", strRTable, strTmp)
        If rsTable.RecordCount = 0 Then
            strModifySQL = strModifySQL & IIf(strModifySQL = "", "", vbNewLine) & "Create Index " & strTmp & " On " & strRTable & "(" & strRCols & ")   ;"
        Else
            If rsTable!Status <> "VALID" Then
                strModifySQL = strModifySQL & IIf(strModifySQL = "", "", vbNewLine) & "ALter Index " & strTmp & "  Rebuild;"
            End If
        End If
        strModifySQL = strModifySQL & IIf(strModifySQL = "", "", vbNewLine) & "ALter  Table  " & strRTable & " Modify Constraint " & strTmp & " Using Index " & strTmp & ";"
        merrCur.ErrAdvice = "����������Ψһ��ȱ���������봴�������ԡ�"
        merrCur.ErrOparate = vbRetry
        merrCur.ErrOparateModi = vbRetry
        merrCur.ErrModiSQL = strModifySQL
    End If
End Sub

Private Function CheckUniqueKeyCol(ByVal strSQL As String, Optional ByVal blnCreateTab As Boolean) As String
'���ܣ����SQL������ΨһԼ�������ݿ����Ƿ�һ��
'������strSQL=�Ѹ�ʽ��Ϊ��׼��д��SQL���
'      blnCreateTab=�Ƿ񴴽�����ã��õ��ò����ô���SQL������飬ֻ��������SQL
'���أ�������ʾ����(strAdvice)��ȱʡ������ťֵ(intAdvice)
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
    '������������Լ������Լ������
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
            '�鿴ͬ������
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
                merrCur.ErrAdvice = "�������������Ѿ�����Լ��""" & strLikeConName & """,��������Լ���л����������в��죬���ֹ���������ԡ�"
                merrCur.ErrOparate = vbRetry
                merrCur.ErrModiSQL = "ALTER TABLE " & strLikeOwner & "." & strTab & " Drop CONSTRAINT " & strLikeConName & ";"
                merrCur.ErrOparateModi = vbRetry
            End If
        Else
            If blnCreateTab Then
                CheckUniqueKeyCol = "ALTER TABLE " & strOwner & "." & strTab & " Drop CONSTRAINT " & strName & ";"
            Else
                merrCur.ErrAdvice = "�Ѿ����ڵ�������ΨһԼ�����ֶλ��ֶ�˳����SQL��䲻�������ֹ���������ԡ�"
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
                merrCur.ErrAdvice = "�Ѿ�������ͬԼ�����ϵ�Լ��������Լ�����Ͳ�ͬ��Լ��" & strSameConName & "�����ֹ���������ԡ�"
                merrCur.ErrModiSQL = "ALTER TABLE " & strOwner & "." & strTab & " Drop CONSTRAINT " & strSameConName & ";"
                merrCur.ErrOparateModi = vbIgnore
            End If
        Else
            If blnCreateTab Then
                CheckUniqueKeyCol = "alter table " & strOwner & "." & strTab & " rename constraint  " & strSameConName & " to " & strName & " ;" & vbNewLine & _
                                    "alter index " & strSameConName & " rename to " & strName & " ;"
            Else
                merrCur.ErrOparate = vbRetry
                merrCur.ErrAdvice = "�Ѿ�����ͬ���ϵ�Լ��������������SQL���������ֹ���������ԡ�"
                merrCur.ErrModiSQL = "alter table " & strOwner & "." & strTab & " rename constraint  " & strSameConName & " to " & strName & " ;" & vbNewLine & _
                                    "alter index " & strSameConName & " rename to " & strName & " ;"
                merrCur.ErrOparateModi = vbIgnore
            End If
        End If
    Else
        merrCur.ErrAdvice = "�Ѿ�����ͬԼ�������Ժ��Ըô���"
        merrCur.ErrOparate = vbIgnore
        merrCur.ErrModiSQL = ""
        merrCur.ErrOparateModi = vbIgnore
    End If
End Function

Private Function CheckIndexCol(ByVal strSQL As String, Optional ByVal blnCreateTab As Boolean) As String
'���ܣ����SQL������ΨһԼ�������ݿ����Ƿ�һ��
'������strSQL=�Ѹ�ʽ��Ϊ��׼��д��SQL���
'      blnCreateTab=�Ƿ񴴽�����ã��õ��ò����ô���SQL������飬ֻ��������SQL
'���أ�������ʾ����(strAdvice)��ȱʡ������ťֵ(intAdvice)
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
    '������������������
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
            '�鿴ͬ������
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
            merrCur.ErrAdvice = "�������������Ѿ���������""" & strLikeIndName & """,���������������в��죬���ֹ���������ԡ�"
            merrCur.ErrOparate = vbRetry
            merrCur.ErrModiSQL = "Drop Index " & strLikeOwner & "." & strLikeIndName & ";"
            merrCur.ErrOparateModi = vbRetry
        Else
            merrCur.ErrAdvice = "�Ѿ����ڵ�������ΨһԼ�����ֶλ��ֶ�˳����SQL��䲻�������ֹ���������ԡ�"
            merrCur.ErrOparate = vbRetry
            merrCur.ErrModiSQL = "Drop Index " & strOwner & "." & strName & ";"
            merrCur.ErrOparateModi = vbRetry
        End If
    ElseIf strSameIndName <> "" Then
        merrCur.ErrAdvice = "�Ѿ�����ͬ���ϵ�����������SQL���������ֹ���������ԡ�"
        merrCur.ErrOparate = vbRetry
        merrCur.ErrModiSQL = "Alter Index " & strSameOwner & "." & strSameIndName & " Rename to " & strName & ";"
        merrCur.ErrOparateModi = vbIgnore
    Else
        merrCur.ErrAdvice = "�Ѿ�����ͬ���������Ժ��Ըô���"
        merrCur.ErrOparate = vbIgnore
        merrCur.ErrModiSQL = ""
        merrCur.ErrOparateModi = vbIgnore
    End If
End Function

Private Function CheckCreateTabCol(ByVal strSQL As String, ByVal strOwner As String, ByVal strName As String) As String
'���ܣ����SQL����������뵱ǰ���ݿ��е��Ƿ�һ��
'������strSQL=�Ѹ�ʽ��Ϊ��׼��д��SQL���
'���أ����ݿ��б�SQL��Ҫ�ٵ��ֶ�
    Dim rsTemp As New ADODB.Recordset
    Dim intMatch As Integer, i As Long
    Dim arrCol As Variant, strError As String
    Dim strCol As String, strType As String
    Dim intLen As Integer, intDigit As Integer
    Dim strModifySQL As String
    Dim strTmpSQL As String, strConsSQL As String
    
    'ȡ��SQL����е��ж���
    intMatch = 1
    For i = InStr(strSQL, "(") + 1 To Len(strSQL)
        If Mid(strSQL, i, 1) = "(" Then
            intMatch = intMatch + 1
        ElseIf Mid(strSQL, i, 1) = ")" Then
            intMatch = intMatch - 1
            If intMatch = 0 Then Exit For
        End If
        If Mid(strSQL, i, 1) = "," And intMatch = 1 Then
            strCol = strCol & "|" '������������е�",",��Number(16,5)
        Else
            strCol = strCol & Mid(strSQL, i, 1)
        End If
    Next
    arrCol = Split(strCol, "|")

    '��SQL�е��ж��������ݿ��еĽ��бȽ�
    On Error Resume Next
    strSQL = "Select Column_Name,Data_Type,Data_Length,Data_Precision,Data_Scale From ALL_Tab_Columns" & _
        " Where OWNER=" & IIf(strOwner = "", "User", "'" & strOwner & "'") & " And Table_Name='" & strName & "'"
    rsTemp.CursorLocation = adUseClient
    rsTemp.Open strSQL, mcnThis, adOpenKeyset
    For i = 0 To UBound(arrCol)
        arrCol(i) = Trim(arrCol(i))

        strCol = Left(arrCol(i), InStr(arrCol(i), " ") - 1) '���� Number ( 16, 5) Not Null Default 1.23
        strType = Mid(arrCol(i), InStr(arrCol(i), " ") + 1)
        If arrCol(i) Like "* PRIMARY KEY*" Or arrCol(i) Like "* UNIQUE*" And strCol = "CONSTRAINT" Then
            '���Լ����
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
                        '������
                    End If
                Else
                    strError = strError & "," & strCol
                End If
            End If
        End If
    Next
    strError = Mid(strError, 2)
    If strConsSQL = "" Then
        merrCur.ErrAdvice = "�Ѿ����ڵı�""" & strName & """���ֶ�""" & strError & """��SQL��䲻�������ֹ���������ԡ�"
        CheckCreateTabCol = strError
        merrCur.ErrModiSQL = strModifySQL
    Else
        merrCur.ErrAdvice = "1���Ѿ����ڵı�""" & strName & """���ֶ�""" & strError & """��SQL��䲻�������ֹ���������ԡ�"
        merrCur.ErrAdvice = "2���Ѿ����ڵ�������ΨһԼ�����ֶλ��ֶ�˳����SQL��䲻�������ֹ���������ԡ�"""
        CheckCreateTabCol = strError & IIf(strError = "", "", vbNewLine) & strConsSQL
        merrCur.ErrModiSQL = strModifySQL & IIf(strModifySQL = "", "", vbNewLine) & strConsSQL
    End If
    merrCur.ErrOparateModi = vbRetry
End Function

Private Sub GetCreateName(ByVal strSQL As String, strOwner As String, strName As String, strType As String)
'���ܣ���Create SQL����з����������Ķ�����������
'������strSQL=�Ѹ�ʽ��Ϊ��׼��д��SQL���
'���أ������Ķ�����,���ܰ���������,��"ZLHIS.���ű�"
    strOwner = "": strName = "": strType = ""

    If strSQL Like "CREATE *" Then
        'ֻ������User_Objects�е�����,û�а���Create Role,Tablespace
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
'���ܣ��ж�ָ���Ķ����Ƿ����
'˵������ΪԼ��ʱ�����봫strTableName
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
'���ܣ���鲢��������
 '˵��������ΪΥ��ΨһԼ���ĲŽ��м��
 '          [Microsoft][ODBC driver for Oracle][Oracle]ORA-00001: Υ��ΨһԼ������ (ZLTOOLS.ZLPROGPRIVS_PK)
    Dim strConstraint As String, strUser As String, strTable As String, strSeqName As String
    Dim strSeqCol As String, intCurMax As Long
    Dim arrTmp As Variant
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    '��ȡԼ�������Լ�������
    strErr = UCase(strErr)
    arrTmp = Split(strErr, "ORA-00001:")
    If UBound(arrTmp) <> 1 Then Exit Sub
    strErr = arrTmp(1)
    If Not strErr Like "*(*)*" Then Exit Sub
    strConstraint = Split(Split(strErr, ")")(0), "(")(1)
    arrTmp = Split(strConstraint, ".")
    strConstraint = arrTmp(1)
    strUser = arrTmp(0)
    '��ȡԼ������ϸ��Ϣ
    strSQL = "Select Constraint_Name, Table_Name" & vbNewLine & _
                    "From All_Constraints a" & vbNewLine & _
                    "Where A.Owner =[1] And A.Constraint_Name = [2]"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "������", strUser, strConstraint)
    If rsTmp.EOF Then Exit Sub
    strTable = rsTmp!Table_Name & ""
    '�ж��Ƿ���ڸñ������
    strSQL = "Select B.Sequence_Name, Substr(Sequence_Name, 1, Instr(Sequence_Name, '_') - 1) Table_Name," & vbNewLine & _
                    "       Substr(Sequence_Name, Instr(Sequence_Name, '_') + 1) Column_Name, B.Last_Number" & vbNewLine & _
                    "From All_Sequences b" & vbNewLine & _
                    "Where B.Sequence_Owner =[1] And B.Sequence_Name Like [2]"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "������", strUser, strTable & "_%")
    If rsTmp.EOF Then Exit Sub
    strSeqName = rsTmp!Sequence_Name: strSeqCol = rsTmp!Column_Name: intCurMax = Val(rsTmp!Last_Number & "")
    '�ж�Լ�����д��ڲ��������е���
    strSQL = "Select 1" & vbNewLine & _
                "From All_Cons_Columns a" & vbNewLine & _
                "Where A.Owner =[1] And A.Constraint_Name =[2] And A.Column_Name = [3]"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "������", strUser, strConstraint, strSeqCol)
    If rsTmp.EOF Then Exit Sub
    
    '��������������
    strSQL = "Declare" & vbNewLine & _
                    "  N_Maxid Number(18);" & vbNewLine & _
                    "  N_Curid Number(18);" & vbNewLine & _
                    "  N_Incre Number(18);" & vbNewLine & _
                    "  N_Tmp   Number(18);" & vbNewLine & _
                    "Begin" & vbNewLine & _
                    "  --��ȡ���ݱ��и������ֵ"
    If strTable = "������ü�¼" Or strTable = "סԺ���ü�¼" Or strTable = "���˷��ü�¼" Then
'        strSQL = strSQL & vbNewLine & _
'                    "  Select Max(Id) Into N_Maxid From [*������*].[*����*];"
        strSQL = strSQL & vbNewLine & _
                    "  Select Max(Mid)" & vbNewLine & _
                    "  Into N_Maxid" & vbNewLine & _
                    "  From (Select Max(" & strSeqCol & ") As Mid" & vbNewLine & _
                    "         From [*������*].������ü�¼" & vbNewLine & _
                    "         Union All" & vbNewLine & _
                    "         Select Max(" & strSeqCol & ") As Mid From [*������*].סԺ���ü�¼);"
    Else
        strSQL = strSQL & vbNewLine & "  Select Max(" & strSeqCol & ") Into N_Maxid From [*������*].[*����*];"
    End If
    strSQL = strSQL & vbNewLine & _
                    "  N_Maxid := Nvl(N_Maxid, 0);" & vbNewLine & _
                    "  --��ȡ��������ֵ" & vbNewLine & _
                    "  Select [*������*].Nextval Into N_Curid From Dual;" & vbNewLine & _
                    "  --��������" & vbNewLine & _
                    "  If N_Maxid - N_Curid > 0 Then" & vbNewLine & _
                    "    --��ȡ���е�ǰ����" & vbNewLine & _
                    "    Select Increment_By" & vbNewLine & _
                    "    Into N_Incre" & vbNewLine & _
                    "    From All_Sequences" & vbNewLine & _
                    "    Where Sequence_Owner = '[*������*]' And Sequence_Name = '[*������*]';" & vbNewLine & _
                    "    N_Incre := Nvl(N_Incre, 1);" & vbNewLine & _
                    "    --�����ɷ�������" & vbNewLine & _
                    "    Execute Immediate 'Alter Sequence [*������*].[*������*] Increment By ' ||(N_Maxid - N_Curid);" & vbNewLine & _
                    "    --�ƶ�һ������" & vbNewLine & _
                    "    Select [*������*].Nextval Into N_Tmp From Dual;" & vbNewLine & _
                    "    --�ָ�ԭʼ����" & vbNewLine & _
                    "    Execute Immediate 'Alter Sequence [*������*].[*������*] Increment By ' ||N_Incre;" & vbNewLine & _
                    "  End If;" & vbNewLine & _
                    "End;" & vbNewLine & _
                    "/"
    strSQL = Replace(Replace(Replace(strSQL, "[*������*]", strUser), "[*����*]", strTable), "[*������*]", strSeqName)
    merrCur.ErrOparateModi = vbRetry
    merrCur.ErrModiSQL = strSQL
End Sub

Private Sub CheckAdjustConstraint(ByVal strSQL As String)
'���ܣ���鲢����Լ��
 'ORA-00955: �����ѱ����ж���ռ��
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
    '����ͬ������
    If ObjectExists(strUser, strConstraint, "INDEX") Then
        merrCur.ErrAdvice = "Լ����ͬ������ռ�ã���ִ��ɾ�����������ԡ�"
        merrCur.ErrOparate = vbRetry
        merrCur.ErrModiSQL = "Drop Index " & strConstraint & ";"
        merrCur.ErrOparateModi = vbRetry
    End If
End Sub

Private Sub CheckAdjust12899(ByVal strErr As String)
'ORA-12899: �� "ZLHIS"."���������Ա"."��������" ��ֵ̫�� (ʵ��ֵ: 62, ���ֵ: 60)
'ORA-12899: value too large for column "SYSTEM"."STUDENTINFO"."SNAME" (actual: 78, maximum: 30)
    Dim arrTmp As Variant
    Dim strColLen As String
    Dim strOwner As String, strTable As String, strColName As String
    Dim strInfo As String, intNewLen As Integer, intOldLen As Integer
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    arrTmp = Split(strErr, "(")
    '��ȡ�о�����Ϣ
    If UBound(arrTmp) < 1 Then Exit Sub
    strColLen = arrTmp(1)
    strColLen = Split(strColLen, ")")(0)
    '��ȡ����Ϣ
    strInfo = Replace(arrTmp(0), """.""", ".")
    arrTmp = Split(strInfo, """")
    If UBound(arrTmp) < 1 Then Exit Sub
    strInfo = Trim(arrTmp(1)) 'Ownere.Talbe.Col
    arrTmp = Split(strInfo, ".")
    If UBound(arrTmp) < 2 Then Exit Sub
    strOwner = UCase(Trim(arrTmp(0)))
    strTable = UCase(Trim(arrTmp(1)))
    strColName = UCase(Trim(arrTmp(2)))
    '��ȡ�³���
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
        'ORA-06575: ��������� TTT ������Ч״̬
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
 'ORA-08002: ���� ZLRPTDATAS_ID.CURRVAL ��δ�ڴ˽����ж���
    Dim strSeq As String
    Dim arrTmp As Variant
    
    strErr = UCase(strErr)
    If strErr Like "ORA-08002: ����*" Then
        strSeq = Mid(strErr, Len("ORA-08002: ����") + 1)
        strSeq = Trim(Mid(strSeq, 1, Len(strSeq) + Len("��δ�ڴ˽����ж���")))
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
'�����ֵ���Զ��޸�
'����:ORA-00001: Υ��ΨһԼ������ (ZLHIS.��Ⱦ����_PK)
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strOwner As String, strTableName As String, strConsCols As String, strConsColDataType As String
    Dim strSeqName As String, strSeqCol As String, blnMultiCon As Boolean
    Dim strDataSQL As String, strUpdateCols As String, strCols As String
    Dim arrConsCols As Variant, arrConsColDataType As Variant, arrUpdateCol As Variant
    Dim strWhereSQL As String, strTmp As String, strDataCol As String, strAdjustSQL As String
    Dim i As Integer, blnCanAdjust As Boolean, strHint As String
    '���������е���Ϣ������ȡ��ĸ���Լ��������Ϣ
    blnCanAdjust = True
    If Not GetConstraintInfo(strErrDesc, strOwner, strTableName, _
                    strConsCols, strConsColDataType, strSeqName, strSeqCol, blnMultiCon, strHint) Then
        blnCanAdjust = False
    ElseIf strHint = "" Then
        '��SQL�н����������������е�SQL,����ȡ�����������
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
        '�Ӳ��������޳����е�������ΨһԼ����
        For i = LBound(arrConsCols) To UBound(arrConsCols)
            If InStr(strUpdateCols, "," & arrConsCols(i) & ",") = 0 Then
                mclsrun.WriteLog String(17, " ") & "��������(ϵͳ)���棺" & strOwner & "." & strTableName & "�Ĳ���SQL��û�а������е�Ψһ������Լ��(ȱʧ�У�" & arrConsCols(i) & ")���޷��Զ��޸���"
                blnCanAdjust = False
            Else
                strUpdateCols = Replace(strUpdateCols, "," & arrConsCols(i) & ",", ",")
            End If
            If arrConsColDataType(i) <> "-1" Then
                If strTableName = "ZLPROGPRIVS" Then '����Ȩ�����⴦��
                    If InStr("������,����,Ȩ��", arrConsCols(i)) > 0 Then
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
                mclsrun.WriteLog String(17, " ") & "��������(ϵͳ)���棺" & strOwner & "." & strTableName & "." & arrConsCols(i) & "�����������޷��ṩ�Զ�����������ϵ������Ա"
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
            If strTableName = "ZLPARAMETERS" Then '�жϸò�����35.0���ϻ������£����ܵ�ǰ��35���룬�����Ծ���Ҫ���в���˵�����¡���Ϊ��ϵͳ�������ܲ���ϵͳ���ǵ���35
                If InStr(strCols, "Ӱ�����˵��") > 0 Then
                    strUpdateCols = "Ӱ�����˵��,����ֵ����,����˵��,����˵��,����˵��"
                Else
                    strUpdateCols = "����˵��"
                End If
            End If
            '���ɲ���SQL
            strSQL = "Select " & strCols & " From (" & strDataSQL & ") a Where Not Exists(Select 1 From " & strOwner & "." & strTableName & " b Where " & strWhereSQL & ")"
            On Error Resume Next
            Set rsTmp = gclsBase.OpenSQLRecord(mcnThis, strSQL, App.Title)
            If rsTmp Is Nothing Then
                mclsrun.WriteLog String(17, " ") & "��������(ϵͳ)���棺��ȡ�������ݳ����ô����޷��Զ��޸�,��Ϣ(" & err.Description & ",SQL:" & strSQL & ")��"
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
                            mclsrun.WriteLog String(17, " ") & "��������(ϵͳ)���棺" & strOwner & "." & strTableName & "." & rsTmp.Fields(i).name & "�����������޷��ṩ�Զ�����������ϵ������Ա"
                        End If
                    Next
                    strAdjustSQL = strAdjustSQL & vbNewLine & "Insert Into " & strOwner & "." & strTableName & _
                                "(" & IIf(strSeqCol <> "", strSeqCol & ",", "") & strCols & ")" & _
                                "Select " & IIf(strSeqName <> "", strSeqName & ".Nextval,", "") & Mid(strDataCol, 2) & " From Dual;"
                    mclsrun.WriteLog String(17, " ") & "��������(ϵͳ)�������ݣ�" & strCols & "(" & Mid(strDataCol, 2) & ")"
                    rsTmp.MoveNext
                Loop
            End If
            If blnCanAdjust Then
                '���ɸ���SQL
                strSQL = "Select " & AddTablePreSubfix(strCols, "A") & IIf(strUpdateCols = "", "", "," & AddTablePreSubfix(strUpdateCols, "B", "B")) & " From (" & strDataSQL & ") a ," & strOwner & "." & strTableName & " b Where " & strWhereSQL
                Set rsTmp = gclsBase.OpenSQLRecord(mcnThis, strSQL, App.Title)
                If rsTmp Is Nothing Then
                    mclsrun.WriteLog String(17, " ") & "��������(ϵͳ)���棺��ȡ�������ݳ����ô����޷��Զ��޸�,��Ϣ(" & err.Description & ",SQL:" & strSQL & ")��"
                    err.Clear
                    blnCanAdjust = False
                Else
                    
                    arrUpdateCol = Split(strUpdateCols, ",")
                    Do While Not rsTmp.EOF
                        strDataCol = "": strTmp = ""
                        If strUpdateCols <> "" Then '���ڸ����ֶ�,���ȡ�����ֶ�SQL
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
                                    mclsrun.WriteLog String(17, " ") & "��������(ϵͳ)���棺" & strOwner & "." & strTableName & "." & arrUpdateCol(i) & "�����������޷��ṩ�Զ�����������ϵ������Ա"
                                End If
                                strTmp = strTmp & "->"
                                '���������ֶ�ֵ��¼
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
                        For i = LBound(arrConsCols) To UBound(arrConsCols) '��ȡ���µ�Լ������
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
                            mclsrun.WriteLog String(17, " ") & "��������(ϵͳ)�������ݣ�" & Mid(strWhereSQL, Len(" And  ")) & "([�ֶ���:SQL����ֵ->���ݿ���ֵ]" & Mid(strTmp, 2) & ")"
                        Else
                            mclsrun.WriteLog String(17, " ") & "��������(ϵͳ)�Ѿ���������-" & IIf(strUpdateCols = "", "�޸�����", "��������Ψһ��") & "��" & Mid(strWhereSQL, Len(" And  ")) & IIf(strUpdateCols = "", "", "([�ֶ���:SQL����ֵ->���ݿ���ֵ]" & Mid(strTmp, 2) & ")")
                        End If
                        rsTmp.MoveNext
                    Loop
                End If
            End If
        End If
    '����ͨ��11g�����Ա����ظ����ݵ��µĲ���ʧ��
    ElseIf strHint <> "" Then
        strInputSQL = Trim(strInputSQL)
        strAdjustSQL = "Insert " & strHint & Mid(strInputSQL, Len("Insert") + 1) & ";"
        mclsrun.WriteLog String(17, " ") & "��������(ϵͳ)����ͨ��11g�����Ա����������ʧ��" & strHint
    End If
    If strAdjustSQL <> "" And blnCanAdjust Then
        strSQL = Mid(strAdjustSQL, Len(vbNewLine) + 1)
        merrCur.ErrAdvice = "������Щ�ֶ�δ���»򲿷������Ѿ����ڣ�����ȷ�ϴ���"
        merrCur.ErrOparate = vbRetry
        merrCur.ErrOparateModi = vbIgnore
        merrCur.ErrModiSQL = merrCur.ErrModiSQL & vbNewLine & strAdjustSQL
        CheckAdjustTableData = True
    ElseIf Not blnCanAdjust Then
        If bytMode = 0 Then
            merrCur.ErrAdvice = "����������ظ����г���һ������¿��Ժ��Ըô���"
            merrCur.ErrOparate = vbIgnore
            merrCur.ErrOparateModi = vbIgnore
        Else
            merrCur.ErrAdvice = "�޷��Զ�������ݣ�����ȷ��SQL�����ݿ�һ�£�"
            merrCur.ErrOparate = vbRetry
            merrCur.ErrOparateModi = vbIgnore
        End If
    Else
        merrCur.ErrAdvice = "�ű��������Ѿ����ڣ�����ȷ�ϴ���"
        merrCur.ErrOparate = vbIgnore
        merrCur.ErrOparateModi = vbIgnore
        CheckAdjustTableData = True
    End If
End Function

Private Function AddTablePreSubfix(ByVal strCols As String, Optional ByVal strPresubfix As String, Optional ByVal strAlisSubFix As String) As String
'���ܣ����ֶ��������ӱ�ǰ׺���߱���
'������strCols-�ֶμ��ϣ��Զ��ŷָ�
'      strPresubfix=����ǰ׺
'      strAlisSubFix=�ֶεı�������� COLA�����ɣ�strPresubfix.COLA  COLA&strAlisSubFix
'���أ����ɺ������
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
'���ܣ�ͨ������������ȡ����,��������ߣ�����Ψһ����Լ����,�ñ���ڵ����е���Ϣ
'������strErrDesc=���ݲ�������Ĵ�����Ϣ
'���أ��Ƿ��ȡ�ɹ�
'strOwner=��ȡ�ı�������
'strTableName=��ȡ�ı���
'strConsCols=�ñ��ϴ��ڵ�����ΨһԼ��������Լ�����еĺϼ����Զ��ŷָ��С�ע�⣬���ϼ��ų������ж�Ӧ��
'strConsColDataType=Լ���е������ͣ�-1-���ܽ��нű��Զ����������ͣ�0-Char,1-NUMber,2-date
'strSeqName=�ñ��ϴ��ڵ�����
'strSeqCols=�ñ������ж�Ӧ����
'blnMultiCon=�Ƿ���ڶ��Լ�����ų����ж�Ӧ����
'˵�������в���ҽ�����͡�����ҽ����¼���ŶӽкŶ��С�����鵵��Ϣ����2�����У������ֻ��һ�������ֻ���ǵ�������
    Dim arrTmp As Variant, strConName As String
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strTmp As String
    On Error Resume Next
    'ǰ���ִ���϶࣬����ֹ�ָ��
    strHint = ""
    If strErrDesc Like "*ZLTOOLS.ZLPARAMETERS_UQ_*" Or strErrDesc Like "*ZLTOOLS.ZLPARAMETERS_PK*" Then
        strOwner = "ZLTOOLS"
        strTableName = "ZLPARAMETERS"
        strConsCols = "ϵͳ,ģ��,������,������"
        strConsColDataType = "1,1,1,0"
        strSeqName = "ZLPARAMETERS_ID"
        strSeqCol = "ID"
        blnMultiCon = True
    ElseIf merrCur.ErrDesc Like "*ZLTOOLS.ZLPROGFUNCS_PK*" Then
        strOwner = "ZLTOOLS"
        strTableName = "ZLPROGFUNCS"
        strConsCols = "ϵͳ,���,����"
        strConsColDataType = "1,1,0"
        strSeqName = ""
        strSeqCol = ""
        If Not gblnClose11g Then
            If GetOracleVersion(True, True) >= 11 Then strHint = "/*+ IGNORE_ROW_ON_DUPKEY_INDEX(ZLPROGFUNCS,ZLPROGFUNCS_PK)*/ "
        End If
    ElseIf merrCur.ErrDesc Like "*ZLTOOLS.ZLPROGPRIVS_PK*" Then
        strOwner = "ZLTOOLS"
        strTableName = "ZLPROGPRIVS"
        strConsCols = "ϵͳ,���,����,������,����,Ȩ��"
        strConsColDataType = "1,1,0,0,0,0"
        strSeqName = ""
        strSeqCol = ""
        If Not gblnClose11g Then
            If GetOracleVersion(True, True) >= 11 Then strHint = "/*+ IGNORE_ROW_ON_DUPKEY_INDEX(ZLPROGPRIVS,ZLPROGPRIVS_PK)*/ "
        End If
    Else
        '��ȡ���������е�Υ����Լ������
        arrTmp = Split(strErrDesc, "(")
        If UBound(arrTmp) < 1 Then
            mclsrun.WriteLog String(17, " ") & "��������(ϵͳ)���棺(1)�޷�����������Ϣ���ô����޷��Զ��޸���"
            Exit Function
        End If
        strTmp = arrTmp(1)
        arrTmp = Split(strTmp, ")")
        If UBound(arrTmp) < 1 Then
            mclsrun.WriteLog String(17, " ") & "��������(ϵͳ)���棺(2)�޷�����������Ϣ���ô����޷��Զ��޸���"
            Exit Function
        End If
        strTmp = arrTmp(0)
        arrTmp = Split(UCase(strTmp), ".")
         If UBound(arrTmp) < 1 Then
            mclsrun.WriteLog String(17, " ") & "��������(ϵͳ)���棺(3)�޷�����������Ϣ���ô����޷��Զ��޸���"
            Exit Function
        End If
        strOwner = Trim(arrTmp(0))
        strConName = Trim(arrTmp(1))
        '��ȡԼ������Լ�����ͣ�Լ������Ψһ��������Լ�����˳����޷���ȡԼ����Ҳ�˳�
        strSQL = "Select a.Constraint_Type, a.Table_Name" & vbNewLine & _
                "From All_Constraints a" & vbNewLine & _
                "Where a.Owner = [1] And a.Constraint_Name = [2] And a.Constraint_Type In ('P', 'U')"
        Set rsTmp = gclsBase.OpenSQLRecord(mcnThis, strSQL, "��������", strOwner, strConName)
        If rsTmp Is Nothing Then
            mclsrun.WriteLog String(17, " ") & "��������(ϵͳ)���棺��ȡΥ��Լ���ı�����ô����޷��Զ��޸�,��Ϣ(" & err.Description & ",SQL:" & strSQL & ")��"
            err.Clear
            Exit Function
        ElseIf rsTmp.EOF Then
            mclsrun.WriteLog String(17, " ") & "��������(ϵͳ)���棺δ�ҵ�Υ��Լ���ı��ô����޷��Զ��޸���"
            Exit Function
        End If
        strTableName = rsTmp!Table_Name
        'ֻҪ����ҵ���ZLbakTale��ZLBigTables,�����Ĭ�ϵ����Զ�����
        strSQL = "Select Count(1) ����" & vbNewLine & _
                "From (Select ���� From Zltools.Zlbaktables Union All Select ���� From Zlbigtables) a" & vbNewLine & _
                "Where a.���� = [1]"
        Set rsTmp = gclsBase.OpenSQLRecord(mcnThis, strSQL, "��������", strOwner, strTableName)
        If rsTmp Is Nothing Then
            mclsrun.WriteLog String(17, " ") & "��������(ϵͳ)���棺Υ��Լ���ı��Ƿ���ҵ���������ô����޷��Զ��޸�,��Ϣ(" & err.Description & ",SQL:" & strSQL & ")��"
            err.Clear
            Exit Function
        ElseIf rsTmp!���� > 0 Then
            mclsrun.WriteLog String(17, " ") & "��������(ϵͳ)��ʾ��Υ��Լ���ı�Ϊҵ�����ݱ��ô����޷��Զ��޸���"
            Exit Function
        End If
        '��ȡ����
        strSQL = "Select a.SEQUENCE_NAME" & vbNewLine & _
                "From All_Sequences a" & vbNewLine & _
                "Where a.Sequence_Owner =[1] and a.SEQUENCE_NAME like [2]"
        Set rsTmp = gclsBase.OpenSQLRecord(mcnThis, strSQL, "��������", strOwner, strTableName & "_%")
        strSeqName = "": strSeqCol = ""
        If Not rsTmp Is Nothing Then
            If rsTmp.RecordCount > 1 Then
                mclsrun.WriteLog String(17, " ") & "��������(ϵͳ)��ʾ��Υ��Լ���ı���ڶ�����У��ô����޷��Զ��޸���"
                Exit Function '�����в�����
            End If
            If Not rsTmp.EOF Then
                strSeqName = rsTmp!Sequence_Name
                strSeqCol = UCase(Mid(strSeqName, Len(strTableName & "_%")))
            End If
        Else
            mclsrun.WriteLog String(17, " ") & "��������(ϵͳ)���棺Υ��Լ���ı��Ӧ���еļ������ô����޷��Զ��޸�,��Ϣ(" & err.Description & ",SQL:" & strSQL & ")��"
            err.Clear
            Exit Function
        End If
        If Not gblnClose11g Then
            If GetOracleVersion(True, True) >= 11 Then
                strSQL = "Select a.Table_Name, a.Index_Name" & vbNewLine & _
                        "From All_Indexes A" & vbNewLine & _
                        "Where a.Uniqueness = 'UNIQUE' And a.Table_Owner = [1] And a.Table_Name = [2]"
                Set rsTmp = gclsBase.OpenSQLRecord(mcnThis, strSQL, "��������", strOwner, strTableName)
                If rsTmp.RecordCount = 1 Then
                    Do While Not rsTmp.EOF
                        strHint = strHint & " IGNORE_ROW_ON_DUPKEY_INDEX(" & rsTmp!Table_Name & "," & rsTmp!Index_Name & ")"
                        rsTmp.MoveNext
                    Loop
                End If
                If strHint <> "" Then strHint = "/*+ " & strHint & "*/ "
            End If
        End If
        '��ȡΨһ������������,ֱ��д�����ߣ������ٶȸ���
        strSQL = "Select a.Column_Name, b.Data_Type,count(1) ����" & vbNewLine & _
                "From all_ind_columns a, All_Tab_Columns b, all_indexes c" & vbNewLine & _
                "Where a.TABLE_OWNER = [1] And a.Table_Name = [2] And b.Owner = [1] And b.Table_Name = [2] And" & vbNewLine & _
                "      c.TABLE_OWNER = [1] And c.Table_Name = [2] And c.INDEX_NAME= a.INDEX_NAME And" & vbNewLine & _
                "      a.Column_Name = b.Column_Name And c.UNIQUENESS='UNIQUE'" & vbNewLine & _
                "group by   a.Column_Name, b.Data_Type"
        Set rsTmp = gclsBase.OpenSQLRecord(mcnThis, strSQL, "��������", strOwner, strTableName)
        If rsTmp Is Nothing Then
            mclsrun.WriteLog String(17, " ") & "��������(ϵͳ)���棺Υ��Լ���ı��ȡȫ��Լ���г����ô����޷��Զ��޸�,��Ϣ(" & err.Description & ",SQL:" & strSQL & ")��"
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
                If Not blnMultiCon Then blnMultiCon = Val(rsTmp!����) > 1
            Else
                If Not blnMultiCon Then blnMultiCon = Val(rsTmp!����) > 2
            End If
            rsTmp.MoveNext
        Loop
        If strConsCols = "" And strSeqName = "" Then
            mclsrun.WriteLog String(17, " ") & "��������(ϵͳ)���棺Υ��Լ���ı��޷���ȡ��Ӧ��Լ���У��ô����޷��Զ��޸���"
            Exit Function
        End If
        strConsCols = Mid(strConsCols, 2)
        strConsColDataType = Mid(strConsColDataType, 2)
    End If
    If err.Number <> 0 Then err.Clear
    GetConstraintInfo = True
End Function

Private Function GetSQLData(ByVal bytMode As Byte, ByVal strSQL As String, ByVal strSeqName As String, ByVal strSeqCol As String, ByVal strTable As String, ByRef strDataSQL As String, ByRef strCols As String) As Boolean
'���ܣ�����SQL�е����У�����ȡSQL�в������ݵ�SQL,��������ݵ���
'������bytMode=0-Insert Values��ʽ 1-Insert Select ��ʽ
'      strSQL=��Ҫ�����SQL
'      strSeqName=����ڵ�������
'      strSeqCol=���ж�Ӧ����
'      strTable=������ǰ׺�ı�
'���أ�GetSQLData=�Ƿ����ɹ�
'      strDataSQL=�������ݵ�SQL,�����Ѿ��������Values��ʽ�Ѿ�����дΪSelect  From Dual��ʽ
'      strCols=���ݲ�����У��Ѿ��������
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
    mclsrun.WriteLog String(17, " ") & "��������(ϵͳ)���棺��ȡ�������ݵ�SQL�����ô����޷��Զ��޸�,��Ϣ(" & err.Description & ")��"
    err.Clear
End Function

Private Function CutSegByInfo(ByVal strInput As String, ByVal strKeyInfo As String)
'���ܣ�ȥ���Զ��ŷָ��һ���ַ�����ָ����һ���ε���Ϣ��
    Dim arrTmp As Variant, lngCout As Long, i As Long
    Dim strReturn As String
    Dim arrCols As Variant

    arrTmp = Split(strInput, ",")
    For i = LBound(arrTmp) To UBound(arrTmp)
        lngCout = lngCout + Len(arrTmp(i)) - Len(Replace(arrTmp(i), "'", ""))
        If lngCout Mod 2 = 0 Then
            arrCols = Split(Trim(arrTmp(i)) & " ", " ")
            '��ֹ�� Zlparameters_Id.Nextval AS id������д��
            If UCase(Trim(arrCols(0))) = UCase(Trim(strKeyInfo)) Then
                '��λ��Ŀ�꣬�����д���
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
'���ܣ���InsertSQL�н�����Insert���Ĳ����������ݡ�
'��ʱδʹ��
'������strSQL=���������
'      strCols=���Ĳ�����
'      bytMode=����ʽ0-Insert into values��ʽ,1-Insert Into Select ��ʽ
'      strRemoveCols=��Ҫ�Ƴ����У���ID,����

'���أ������������
    Dim strTmp As String, strTmpCols As String, strData As String
    Dim arrTmp As Variant, arrData As Variant, arrReMoveIndex As Variant
    Dim cllStr As Collection, i As Long, j As Long, lngCout As Long
    Dim strFTMSQL As String
    
    '�ֽ���������
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
    '�Ƴ�ĳЩ��
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
        '�Ƴ�������
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

