VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImp 
   BackColor       =   &H80000005&
   Caption         =   "数据导入"
   ClientHeight    =   6990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8085
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "frmImp.frx":0000
   ScaleHeight     =   6990
   ScaleWidth      =   8085
   WindowState     =   2  'Maximized
   Begin VB.Frame fra 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "程序辅助选项"
      ForeColor       =   &H80000008&
      Height          =   1665
      Index           =   1
      Left            =   4980
      TabIndex        =   21
      Top             =   1950
      Width           =   1935
      Begin VB.CheckBox chk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "删除原有表(&T)"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   9
         Left            =   150
         TabIndex        =   22
         Top             =   300
         Width           =   1575
      End
      Begin VB.CheckBox chk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "清空表数据(&A)"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   7
         Left            =   420
         TabIndex        =   23
         Top             =   690
         Value           =   1  'Checked
         Width           =   1485
      End
      Begin VB.CheckBox chk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "使约束无效(&D)"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   8
         Left            =   420
         TabIndex        =   24
         Top             =   1110
         Value           =   1  'Checked
         Width           =   1485
      End
   End
   Begin VB.TextBox txtFile 
      Appearance      =   0  'Flat
      Height          =   630
      Index           =   2
      Left            =   1020
      Locked          =   -1  'True
      MaxLength       =   256
      MultiLine       =   -1  'True
      TabIndex        =   27
      Top             =   4770
      Width           =   5925
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "执行(&E)"
      Height          =   350
      Left            =   5820
      TabIndex        =   25
      Top             =   3960
      Width           =   1100
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "…"
      Height          =   300
      Index           =   1
      Left            =   6570
      TabIndex        =   8
      Top             =   1560
      Width           =   300
   End
   Begin VB.TextBox txtFile 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   1
      Left            =   2070
      MaxLength       =   256
      TabIndex        =   7
      Top             =   1560
      Width           =   4485
   End
   Begin VB.ComboBox cmbSystem 
      Height          =   300
      Left            =   2070
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   720
      Width           =   4815
   End
   Begin VB.Frame fra 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "导入选项"
      ForeColor       =   &H80000008&
      Height          =   2415
      Index           =   0
      Left            =   1020
      TabIndex        =   9
      Top             =   1920
      Width           =   3795
      Begin VB.TextBox txtBuffer 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   3120
         MaxLength       =   5
         TabIndex        =   20
         Text            =   "300"
         ToolTipText     =   "请注意此值不要超过当前可用内存大小"
         Top             =   2040
         Width           =   555
      End
      Begin VB.TextBox txtUser 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1440
         MaxLength       =   256
         TabIndex        =   18
         Top             =   1680
         Width           =   2235
      End
      Begin VB.CheckBox chk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "忽略对象创建错误(&R)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   390
         TabIndex        =   16
         Top             =   1350
         Value           =   1  'Checked
         Width           =   2025
      End
      Begin VB.CheckBox chk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "数据每行提交(&M)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   2100
         TabIndex        =   15
         Top             =   1005
         Value           =   1  'Checked
         Width           =   1665
      End
      Begin VB.CheckBox chk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "只显示内容(&W)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   390
         TabIndex        =   14
         Top             =   1005
         Width           =   1515
      End
      Begin VB.CheckBox chk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "导入权限(&G)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   2100
         TabIndex        =   13
         Top             =   660
         Value           =   1  'Checked
         Width           =   1305
      End
      Begin VB.CheckBox chk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "导入索引(&I)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   2100
         TabIndex        =   11
         Top             =   330
         Value           =   1  'Checked
         Width           =   1305
      End
      Begin VB.CheckBox chk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "导入约束(&C)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   390
         TabIndex        =   12
         Top             =   660
         Value           =   1  'Checked
         Width           =   1305
      End
      Begin VB.CheckBox chk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "导入表数据(&R)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   390
         TabIndex        =   10
         Top             =   330
         Value           =   1  'Checked
         Width           =   1485
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "数据缓冲区大小(单位:KB)(&B)"
         Height          =   180
         Index           =   5
         Left            =   360
         TabIndex        =   19
         Top             =   2100
         Width           =   2340
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "导出用户(&U)"
         Height          =   180
         Index           =   4
         Left            =   360
         TabIndex        =   17
         Top             =   1740
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "…"
      Height          =   300
      Index           =   0
      Left            =   6570
      TabIndex        =   5
      Top             =   1140
      Width           =   300
   End
   Begin VB.TextBox txtFile 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   0
      Left            =   2070
      MaxLength       =   256
      TabIndex        =   4
      Top             =   1155
      Width           =   4485
   End
   Begin MSComDlg.CommonDialog cmmFile 
      Left            =   5220
      Top             =   30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "命令行文本"
      Height          =   180
      Index           =   3
      Left            =   1020
      TabIndex        =   28
      Top             =   4500
      Width           =   900
   End
   Begin VB.Label lbl说明 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   1020
      TabIndex        =   26
      Top             =   5640
      Width           =   6825
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "记录日志(&L)"
      Height          =   180
      Index           =   2
      Left            =   1020
      TabIndex        =   6
      Top             =   1590
      Width           =   990
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "导入系统(&S)"
      Height          =   180
      Index           =   1
      Left            =   1020
      TabIndex        =   1
      Top             =   780
      Width           =   990
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "导入文件(&F)"
      Height          =   180
      Index           =   0
      Left            =   1020
      TabIndex        =   3
      Top             =   1200
      Width           =   990
   End
   Begin VB.Image imgICO 
      Height          =   480
      Left            =   240
      Picture         =   "frmImp.frx":04F9
      Top             =   690
      Width           =   480
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "数据导入"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   195
      TabIndex        =   0
      Top             =   105
      Width           =   960
   End
End
Attribute VB_Name = "frmImp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mrsSystem As New ADODB.Recordset
Dim mstr所有者 As String '保存当前系统的所有者名
Dim mstrVer As String

Private Enum conCheck
    Rows = 0
    Indexes = 1
    Constraints = 2
    Grants = 3
    OnlyShow = 4
    Commit = 5
    Ignore = 6
    Clear = 7
    Disable = 8
    Drop = 9
End Enum

Private Function GetCommand() As String
    Dim strFromUser As String
    
    strFromUser = Trim(txtUser.Text)
    If strFromUser = "" Then strFromUser = mstr所有者
    
    GetCommand = "IMP" & mstrVer & " USERID=" & gstrUserName & "/" & IIf(gstrUserName <> gstrLoginUserName, "*****", gstrPassword) & IIf(gstrServer = "", "", "@" & gstrServer) _
        & " FROMUSER=(" & strFromUser & ")  TOUSER=(" & mstr所有者 & ") BUFFER=" & IIf(IsNumeric(txtBuffer.Text), CStr(Val(txtBuffer.Text) * 1024), "30720") _
        & " FILE=" & Trim(txtFile(0).Text) & IIf(Trim(txtFile(1).Text) = "", "", " LOG=" & Trim(txtFile(1).Text)) _
        & " ROWS=" & IIf(chk(Rows).value = 1, "Y", "N") & " INDEXES=" & IIf(chk(Indexes).value = 1, "Y", "N") _
        & IIf(chk(Constraints).Enabled, " CONSTRAINTS =" & IIf(chk(Constraints).value = 1, "Y", "N"), "") & " GRANTS =" & IIf(chk(Grants).value = 1, "Y", "N") _
        & " SHOW=" & IIf(chk(OnlyShow).value = 1, "Y", "N" & " COMMIT =" & IIf(chk(Commit).value = 1, "Y", "N") _
                                                       & " IGNORE=" & IIf(chk(Ignore).value = 1, "Y", "N"))
End Function

Private Sub cmdExecute_Click()
    Dim strDMPFile As String
    Dim strLogFile As String
    Dim lngProcess As Long
    Dim lngTemp As Long
    Dim strCommand As String
    Dim varTime As Variant
    Dim rsTemp As New ADODB.Recordset
    Dim rsCons As New ADODB.Recordset
    Dim blnSuccess As Boolean
    Dim intVer As Integer
    Dim strNote As String
    
    intVer = GetOracleVersion
    
    '对文件名的合法性进行判断
    strDMPFile = Trim(txtFile(0).Text)
    strLogFile = Trim(txtFile(1).Text)
    If strDMPFile = "" Then
        txtFile(0).SetFocus
        MsgBox "请输入导入文件名。", vbExclamation, gstrSysName
        Exit Sub
    End If
    If strLogFile = strDMPFile Then
        txtFile(1).SetFocus
        MsgBox "导入文件和日志记录文件不能是同一个文件。", vbExclamation, gstrSysName
        Exit Sub
    End If
    If Dir(strDMPFile) = "" Then
        MsgBox "请输入一个正确的导入文件名。", vbExclamation, gstrSysName
        txtFile(0).SetFocus
        Exit Sub
    End If
    If Dir(strLogFile) <> "" And strLogFile <> "" Then
        If MsgBox("记录日志已经存在，是否覆盖？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            txtFile(1).SetFocus
            Exit Sub
        End If
    End If
    If strLogFile <> "" Then
        On Error Resume Next
        lngTemp = FreeFile
        '记录日志文件可以为空
        Open strLogFile For Binary As lngTemp
        If err <> 0 Then
            MsgBox "记录日志文件名非法。", vbExclamation, gstrSysName
            txtFile(1).SetFocus
            Exit Sub
        End If
        Close lngTemp
    End If
    
    '执行导入操作
    
    If MsgBox("真的要进行导入操作吗？" & vbCrLf & "这会对现有的数据库对象产生影响的。", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    On Error GoTo errHandle
    SetEnable False
    
    frmWait.BeginWait "开始时间:" & Now() & ".正在清除数据……"
    strCommand = GetCommand()
    If gstrUserName <> gstrLoginUserName Then
        strCommand = Replace(strCommand, "*****", gstrPassword)
    End If
    
    varTime = Now() '记录下开始导出的时间
    
    On Error Resume Next
    rsTemp.CursorLocation = adUseClient
    rsCons.CursorLocation = adUseClient
    If chk(Drop).value = 1 Then
        '删除所有表
        gstrSQL = "select TABLE_NAME from all_tables where OWNER='" & mstr所有者 & "' And Instr(Table_NAME,'BIN$')<=0 "
        rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
        Do Until rsTemp.EOF
            '--- 2007-03-07 删除表时,如果是10g,则加上Purge 关键字,不放入回收站
            gcnOracle.Execute "Drop Table " & mstr所有者 & "." & rsTemp("TABLE_NAME") & " cascade constraints" & IIf(intVer >= 100, " Purge", "")
            rsTemp.MoveNext
        Loop
        
        If rsTemp.State = adStateOpen Then rsTemp.Close
        
        '删除所有视图
        gstrSQL = "Select View_name From All_Views Where Owner = '" & mstr所有者 & "'"
        rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
        Do Until rsTemp.EOF
            gcnOracle.Execute "Drop View " & mstr所有者 & "." & rsTemp("View_name")
            rsTemp.MoveNext
        Loop
        
        '删除所有序列
        gstrSQL = "select SEQUENCE_NAME from all_sequences where SEQUENCE_OWNER='" & mstr所有者 & "'"
        rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
        Do Until rsTemp.EOF
            gcnOracle.Execute "Drop Sequence " & mstr所有者 & "." & rsTemp("SEQUENCE_NAME")
            rsTemp.MoveNext
        Loop
    Else
        If chk(Disable).value = 1 Then
            gstrSQL = "select CONSTRAINT_NAME,CONSTRAINT_TYPE,TABLE_NAME from all_constraints where OWNER='" & mstr所有者 & "' And Instr(Table_NAME,'BIN$')<=0"
            rsCons.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
            '首先使外键约束无效
            rsCons.Filter = "CONSTRAINT_TYPE='R'"
            Do Until rsCons.EOF
                gcnOracle.Execute "Alter Table " & mstr所有者 & "." & rsCons("TABLE_NAME") & " disable constraint " & rsCons("CONSTRAINT_NAME")
                rsCons.MoveNext
            Loop
            '再使其它类型的约束无效
            rsCons.Filter = "CONSTRAINT_TYPE<>'R'"
            Do Until rsCons.EOF
                gcnOracle.Execute "Alter Table " & mstr所有者 & "." & rsCons("TABLE_NAME") & " disable constraint " & rsCons("CONSTRAINT_NAME")
                rsCons.MoveNext
            Loop
        End If
        If chk(Clear).value = 1 Then
            gstrSQL = "select TABLE_NAME from all_tables where OWNER='" & mstr所有者 & "' And Instr(Table_NAME,'BIN$')<=0"
            rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
            Do Until rsTemp.EOF
                gcnOracle.Execute "truncate Table " & mstr所有者 & "." & rsTemp("TABLE_NAME") & "  drop storage"
                rsTemp.MoveNext
            Loop
        End If
    End If
    
    '执行Import命令
    '显示正在导入数据
    frmWait.lbl内容 = Replace(frmWait.lbl内容, "清除", "导入")
    err.Clear
    lngTemp = Shell(strCommand, vbHide)
    If err <> 0 Then
        err.Clear
        MsgBox "目前的系统不能正确完成数据恢复，请检查：" & _
            vbCrLf & "   1） 是否存在imp" & mstrVer & ".exe文件；" & _
            vbCrLf & "   2） Set Path是否指向其存在的目录；" & _
            vbCrLf & "   3） 导入文件是由同版本的Export程序导出的。", vbExclamation, gstrSysName
        frmWait.EndWait
        SetEnable True
        Exit Sub
    End If
    
    On Error GoTo errHandle
        
    lngProcess = OpenProcess(Process_Query_Information, False, lngTemp)
    Do
        Sleep 100
        GetExitCodeProcess lngProcess, lngTemp
        DoEvents
    Loop While lngTemp = Still_Active
    CloseHandle lngProcess
    Call AdjustSequence(mstr所有者, gcnOracle)
    
    If lngTemp <> 0 And lngTemp <> 1 Then
        frmWait.EndWait
        MsgBox "数据导入程序运行失败。如果需要看到详细错误信息，可以运行“命令行文本”。" & vbCrLf & _
            "请检查：" & vbCrLf & _
            "   1） 所选的数据文件是否有效的导出文件；" & vbCrLf & _
            "   2） 导出文件与要导入的数据库版本相同，不能将8.0导出的文件导入到8i中；" & vbCrLf & _
            "   3） 导入用户的权限是否符合要求，不能将DBA用户导出的数据用普通用户导入；" & vbCrLf & _
            "   4） 导出用户名是否正确，如果与当前用户名不同，请在“导出用户”输入正确的用户名。", vbExclamation, gstrSysName
        SetEnable True
        Exit Sub
    End If
    '执行完成
errHandle:
    If err = 0 Then blnSuccess = True
    
    On Error Resume Next
    strLogFile = "" '用这个变量来保存未成功恢复的约束
    If chk(Disable).value = 1 And chk(Drop).value = 0 Then
        '首先使其它类型的约束有效
        rsCons.MoveFirst
        rsCons.Filter = "CONSTRAINT_TYPE<>'R'"
        Do Until rsCons.EOF
            gcnOracle.Execute "Alter Table " & mstr所有者 & "." & rsCons("TABLE_NAME") & " Enable constraint " & rsCons("CONSTRAINT_NAME")
            If err <> 0 Then
                err.Clear
                strLogFile = strLogFile & vbCrLf & rsCons("CONSTRAINT_NAME")
            End If
            rsCons.MoveNext
        Loop
        '然后使外键约束有效
        rsCons.Filter = "CONSTRAINT_TYPE='R'"
        Do Until rsCons.EOF
            gcnOracle.Execute "Alter Table " & mstr所有者 & "." & rsCons("TABLE_NAME") & " Enable constraint " & rsCons("CONSTRAINT_NAME")
            If err <> 0 Then
                err.Clear
                strLogFile = strLogFile & vbCrLf & rsCons("CONSTRAINT_NAME")
            End If
            rsCons.MoveNext
        Loop
    End If
    '恢复序列
    Call AdjustSequence(mstr所有者, gcnOracle)
    frmWait.EndWait
    If blnSuccess = True Then
        MsgBox "数据恢复完成！" & vbCrLf & vbCrLf & _
            "共耗时" & Format(CDate(Now - varTime), "hh时mm分ss秒。") & _
            IIf(strLogFile = "", "", "但以下约束不能正常启用：" & strLogFile), vbExclamation, gstrSysName
        If chk(Rows).value = 1 Then strNote = ",表数据"
        If chk(Indexes).value = 1 Then strNote = ",索引"
        If chk(Constraints).value = 1 Then strNote = ",约束"
        If chk(Grants).value = 1 Then strNote = ",权限"
        '插入重要操作日志
        Call SaveAuditLog(2, "执行", "成功将本地文件" & Right(strDMPFile, Len(strDMPFile) - InStrRev(strDMPFile, "\")) & "中的" & Mid(strNote, 2) & "导入到“" & Split(cmbSystem.Text, " ")(0) & "”中")
    Else
        MsgBox "数据导入失败。" & IIf(strLogFile = "", "", "并且以下约束不能正常启用：" & strLogFile), vbExclamation, gstrSysName
    End If
    SetEnable True
End Sub

Private Sub SetEnable(ByVal blnEnable As Boolean)
    frmMDIMain.Enabled = blnEnable
    cmbSystem.Enabled = blnEnable
    cmdExecute.Enabled = blnEnable
    fra(0).Enabled = blnEnable
    fra(1).Enabled = blnEnable
End Sub

Private Sub chk_Click(Index As Integer)
    Dim i As Integer
    If Index = OnlyShow Then
        '当导入过程只完成显示时，有些选项是不可用的
        If chk(Index).value = 1 Then
            For i = 5 To 9
                chk(i).Enabled = False
                chk(i).value = 0
            Next
        Else
            For i = 5 To 9
                chk(i).Enabled = True
            Next
        End If
    ElseIf Index = Drop Then
        '当导入过程要删除原有表时，有些选项是不可用的
        If chk(Index).value = 1 Then
            For i = 7 To 8
                chk(i).Enabled = False
                chk(i).value = 0
            Next
        Else
            For i = 7 To 8
                chk(i).Enabled = True
            Next
        End If
    ElseIf Index = Clear Then
        If chk(Index).value = 1 Then
            chk(Disable).Enabled = False
            chk(Disable).value = 1
        Else
            chk(Disable).Enabled = True
        End If
    End If
    txtFile(2).Text = GetCommand()
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    SendKeys "{TAB}"
End Sub

Private Sub cmdFile_Click(Index As Integer)
    cmmFile.FileName = txtFile(Index).Text
    If Index = 0 Then
        cmmFile.Filter = "导入文件(*.dmp)|*.dmp"
        cmmFile.ShowOpen
    Else
        cmmFile.Filter = "记录日志(*.log)|*.log"
        cmmFile.ShowSave
    End If
    If cmmFile.FileName <> "" Then txtFile(Index).Text = cmmFile.FileName
End Sub

Private Sub cmbSystem_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then txtFile(0).SetFocus
End Sub


Private Sub txtBuffer_Change()
    txtFile(2).Text = GetCommand()
End Sub

Private Sub txtBuffer_GotFocus()
    txtBuffer.SelStart = 0
    txtBuffer.SelLength = Len(txtBuffer.Text)
End Sub

Private Sub txtBuffer_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtBuffer_LostFocus()
    If Not IsNumeric(txtBuffer) Then
        txtBuffer.SetFocus
    End If
End Sub


Private Sub txtFile_Change(Index As Integer)
    If Index <> 2 Then
        txtFile(2).Text = GetCommand()
    End If
End Sub

Private Sub txtFile_GotFocus(Index As Integer)
    'txtFile(Index).SetFocus
End Sub

Private Sub txtFile_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Index = 0 Then
            txtFile(1).SetFocus
        Else
            chk(Rows).SetFocus
        End If
    End If

End Sub

Private Sub Form_Load()
    lbl说明.Caption = "提示：" & vbCrLf & _
                    "     导入数据要经过一个很漫长的过程才能完成。在这段时间内服务器对客户的响应会变得迟钝，因此最好在服务器空闲时完成本操作。" & vbCrLf & _
                    "     在导入过程中或导入完成后，可以通过记录日志文件了解导入的一些具体情况。" & vbCrLf & _
                    "     你如果对导入命令熟悉，也可以直接在Windows运行窗口执行命令行文本。"
    
    Dim intVer As Integer
    
    intVer = GetOracleVersion
    
    If intVer < 80 Then
        MsgBox "该Oracle版本可能由于过旧，本程序可能不能正常运行，" & vbCr _
            & "请考虑将BIN目录中的[IMP+版本号.EXE]改为[IMP.EXE]再执行。", vbExclamation, gstrSysName
        mstrVer = ""
    ElseIf intVer = 80 Then            'Oracle8.0
        mstrVer = "80"
        chk(Constraints).value = 0
        chk(Constraints).Enabled = False
    Else
        mstrVer = ""
    End If
    Call FillSystem
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mrsSystem.State = 1 Then mrsSystem.Close
    Set mrsSystem = Nothing
    mstr所有者 = ""
End Sub

Private Sub Form_Resize()
    Dim sngTemp As Single
    
    On Error Resume Next
    sngTemp = IIf(ScaleWidth > 5000, ScaleWidth, 5000)
    cmbSystem.Width = sngTemp - cmbSystem.Left - 200
    cmdFile(0).Left = sngTemp - cmdFile(0).Width - 200
    cmdFile(1).Left = cmdFile(0).Left
    txtFile(0).Width = cmdFile(0).Left - 15 - txtFile(1).Left
    txtFile(1).Width = txtFile(0).Width
    txtFile(2).Width = cmbSystem.Left + cmbSystem.Width - txtFile(2).Left
    
    lbl说明.Width = ScaleWidth - 200 - lbl说明.Left
    lbl说明.Height = ScaleHeight - 200 - lbl说明.Top
    
End Sub

Private Sub cmbSystem_Click()
    If cmbSystem.ItemData(cmbSystem.ListIndex) = -1 Then
        cmdExecute.Enabled = True
        mstr所有者 = "ZLTOOLS"
    Else
        mrsSystem.Filter = "编号=" & cmbSystem.ItemData(cmbSystem.ListIndex)
        If mrsSystem.RecordCount = 0 Then
            cmdExecute.Enabled = False
        Else
            cmdExecute.Enabled = True
            mstr所有者 = mrsSystem("所有者")
        End If
    End If
    txtFile(2).Text = GetCommand()
End Sub

Private Sub FillSystem()
    '显示所有可显示的系统
    On Error GoTo errHandle
    If gblnDBA = True Then
        Set mrsSystem = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_Zlsystems", "")
    Else
        Set mrsSystem = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_Zlsystems", UCase(gstrUserName))
    End If
    
    Do Until mrsSystem.EOF
        cmbSystem.AddItem mrsSystem("名称") & " v" & mrsSystem("版本号") & "（" & mrsSystem("编号") & "）"
        cmbSystem.ItemData(cmbSystem.NewIndex) = mrsSystem("编号")
        mrsSystem.MoveNext
    Loop
    If gblnDBA = True Then
        cmbSystem.AddItem "管理工具"
        cmbSystem.ItemData(cmbSystem.NewIndex) = -1
    End If
    If mrsSystem.RecordCount > 0 Then
        cmbSystem.ListIndex = 0
    Else
        cmdExecute.Enabled = False
    End If
    Exit Sub
errHandle:
    MsgBox err.Description, vbCritical, Me.Caption
End Sub

Public Function SupportPrint() As Boolean
'返回本窗口是否支持打印，供主窗口调用
    SupportPrint = False
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'供主窗口调用，实现具体的打印工作
'如果没有可打印的，就留下一个空的接口

End Sub

Private Sub txtUser_Change()
    txtFile(2).Text = GetCommand()
End Sub
