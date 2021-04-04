VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmAutoJobs 
   BackColor       =   &H80000005&
   Caption         =   "后台作业管理"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6465
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Picture         =   "frmAutoJobs.frx":0000
   ScaleHeight     =   6450
   ScaleWidth      =   6465
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdDel 
      Caption         =   "删除(&D)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   4965
      TabIndex        =   13
      Top             =   4830
      Width           =   945
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "增加(&A)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   4020
      TabIndex        =   12
      Top             =   4830
      Width           =   945
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "运行设置(&T)…"
      Height          =   350
      Left            =   2355
      TabIndex        =   11
      Top             =   4830
      Width           =   1395
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hgdJobs 
      Height          =   2325
      Left            =   975
      TabIndex        =   8
      Top             =   1290
      Width           =   4920
      _ExtentX        =   8678
      _ExtentY        =   4101
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.OptionButton optKind 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "用户自定(&3)"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   2
      Left            =   4575
      TabIndex        =   7
      Top             =   1065
      Width           =   1305
   End
   Begin VB.OptionButton optKind 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "数据转移(&2)"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   1
      Left            =   3255
      TabIndex        =   6
      Top             =   1065
      Width           =   1305
   End
   Begin VB.OptionButton optKind 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "系统设定(&1)"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   0
      Left            =   1935
      TabIndex        =   5
      Top             =   1065
      Value           =   -1  'True
      Width           =   1305
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "测试执行(&T)…"
      Height          =   350
      Left            =   975
      TabIndex        =   4
      Top             =   4830
      Width           =   1395
   End
   Begin VB.ComboBox cmbSystem 
      Height          =   300
      Left            =   1740
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   615
      Width           =   4185
   End
   Begin VB.Frame fraComment 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1125
      Left            =   975
      TabIndex        =   9
      Top             =   3510
      Width           =   4920
      Begin VB.Label lbl说明 
         BackStyle       =   0  'Transparent
         Height          =   525
         Left            =   690
         TabIndex        =   17
         Top             =   210
         Width           =   1965
      End
      Begin VB.Label lblComment 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "说明："
         Height          =   180
         Left            =   120
         TabIndex        =   15
         Top             =   195
         Width           =   540
      End
      Begin VB.Label lblPara 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "参数："
         Height          =   180
         Left            =   120
         TabIndex        =   14
         Top             =   825
         Width           =   540
      End
   End
   Begin VB.Label lbl附加 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   1080
      TabIndex        =   16
      Top             =   5850
      Width           =   4890
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblJobs 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "作业目录："
      Height          =   180
      Left            =   975
      TabIndex        =   10
      Top             =   1065
      Width           =   900
   End
   Begin VB.Label lblMain 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   180
      Left            =   975
      TabIndex        =   3
      Top             =   5400
      Width           =   4890
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblSys 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "应用系统"
      Height          =   180
      Left            =   975
      TabIndex        =   1
      Top             =   675
      Width           =   720
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "后台作业管理"
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
      Width           =   1440
   End
   Begin VB.Image imgMain 
      Height          =   480
      Left            =   240
      Picture         =   "frmAutoJobs.frx":04F9
      Top             =   645
      Width           =   480
   End
End
Attribute VB_Name = "frmAutoJobs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mlngMaxJobs As Long '本数据库最多可设的作业数
Private mstrSystem As String '记录当前选择的系统
Private mstrDirectory As String '记录当前作业目录

Private Sub cmbSystem_Click()
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHandle
    If cmbSystem.ListCount = 0 Then Exit Sub
    If cmbSystem.ItemData(cmbSystem.ListIndex) = 0 Then
        cmbSystem.Tag = "ZLTOOLS"
    Else
        cmbSystem.Tag = GetOwnerName(Val(cmbSystem.ItemData(cmbSystem.ListIndex)), gcnOracle)
    End If
    
    With rsTemp
        gstrSQL = "select 序号,说明,参数,内容,作业号," & _
                "      名称,decode(作业号,0,'否',null,'否','是') as 自动执行," & _
                "      decode(J.BROKEN,null,'未知','Y','无效','有效') as 状态," & _
                "      执行时间,间隔时间||时间单位 as 间隔时间" & _
                " From zlAutoJobs Z," & IIf(gblnDBA, "dba_jobs", "user_jobs") & " J" & _
                " where Z.作业号=J.JOB(+)" & _
                "   and Nvl(Z.系统,0)=" & cmbSystem.ItemData(cmbSystem.ListIndex)
        If optKind(2).value Then
            gstrSQL = gstrSQL & "   and 类型=3"
        ElseIf optKind(1).value Then
            gstrSQL = gstrSQL & "   and 类型=2"
        Else
            gstrSQL = gstrSQL & "   and 类型=1"
            optKind(0).value = True
        End If
        If .State = adStateOpen Then .Close
        .Open gstrSQL, gcnOracle, adOpenKeyset
    End With
    If cmbSystem.Text = mstrSystem And optKind(1).Tag = mstrDirectory Then
        hgdJobs.Tag = hgdJobs.Row
    Else
        hgdJobs.Tag = ""
    End If
    If Not rsTemp.EOF Then
        Set hgdJobs.Recordset = rsTemp
    Else
        hgdJobs.Cols = 10
        hgdJobs.Rows = 1
        hgdJobs.Rows = 2
        hgdJobs.FixedRows = 1
    End If
    
    With hgdJobs
        .ColWidth(0) = 450
        .ColWidth(1) = 0
        .ColWidth(2) = 0
        .ColWidth(3) = 0
        .ColWidth(4) = 0
        .ColWidth(5) = 2000
        .ColWidth(6) = 800
        .ColAlignment(6) = 4
        .ColWidth(7) = 600
        .ColWidth(8) = 1900
        .ColAlignment(8) = 4
        .ColWidth(9) = 900
        If Val(hgdJobs.Tag) >= .Rows Then hgdJobs.Tag = Val(hgdJobs.Tag) - 1
        .Row = IIf(hgdJobs.Tag = "", 1, hgdJobs.Tag)
        .Col = 0
        .RowSel = .Row
        .ColSel = .Cols - 1
    
        .TextMatrix(0, 0) = "序号"
        .TextMatrix(0, 5) = "名称"
        .TextMatrix(0, 6) = "自动执行"
        .TextMatrix(0, 7) = "状态"
        .TextMatrix(0, 8) = "执行时间"
        .TextMatrix(0, 9) = "间隔时间"
    End With
    
    If mlngMaxJobs > 0 Then
        If Not rsTemp.EOF And Not rsTemp.BOF Then
            cmdTest.Enabled = True
            cmdSet.Enabled = True
            
            If optKind(2).value Then
                cmdAdd.Enabled = True
                cmdDel.Enabled = True
            Else
                cmdAdd.Enabled = False
                cmdDel.Enabled = False
            End If
        Else
            cmdTest.Enabled = False
            cmdSet.Enabled = False
            
            If optKind(2).value Then
                cmdAdd.Enabled = True
                cmdDel.Enabled = False
            Else
                cmdAdd.Enabled = False
                cmdDel.Enabled = False
            End If
        End If
    End If
    Call hgdJobs_RowColChange
    mstrSystem = cmbSystem.Text
    mstrDirectory = optKind(1).Tag
    Exit Sub
ErrHandle:
    MsgBox err.Description, vbCritical, Me.Caption
End Sub

Private Sub cmdAdd_Click()
    With frmAutoJobset
        .Tag = "ADD"                                           '类型必然=3
        .lblSys.Tag = cmbSystem.ItemData(cmbSystem.ListIndex)  '系统
        .imgMain.Tag = cmbSystem.Tag                           '所有者
        .cmdWhat.Enabled = True
        .chk规则.Visible = True
        .txtJobComment.Locked = False
        .txtJobComment.ForeColor = Me.ForeColor
        .Height = .fraCycle.Top + .fraCycle.Height + 600
        .fraPara.Visible = False
        .dtpStart.value = CurrentDate()
        .Show 1, frmMDIMain
    End With
    Call cmbSystem_Click
End Sub

Private Sub cmdDel_Click()
    Dim lngSystem As Long
    Dim cnTools As ADODB.Connection
    Dim strRemarks As String
    
    '验证身份并输入操作说明
    strRemarks = "删除作业：" & hgdJobs.TextMatrix(hgdJobs.Row, 5)
    If Not CheckAuditStatus("0303", "删除", strRemarks) Then Exit Sub
    With cmbSystem
        lngSystem = .ItemData(.ListIndex) \ 100
        If Val(hgdJobs.TextMatrix(hgdJobs.Row, 4)) <> 0 Then
            If .Tag = "ZLTOOLS" Then
                Set cnTools = GetConnection("ZLTOOLS")
                If cnTools Is Nothing Then Exit Sub
            Else
                Set cnTools = gcnOracle
            End If
            gstrSQL = "zl_JobRemove(" & IIf(.ItemData(.ListIndex) = 0, "Null", .ItemData(.ListIndex)) & ",3," & hgdJobs.TextMatrix(hgdJobs.Row, 0) & ")"
            err = 0
            cnTools.Execute gstrSQL, , adCmdStoredProc
            If err <> 0 Then
                MsgBox "作业删除失败！", vbExclamation, gstrSysName
                Exit Sub
            End If
        End If
        gstrSQL = "delete zlAutoJobs" & _
                " where Nvl(系统,0)=" & .ItemData(.ListIndex) & " and 类型=3" & _
                " and 序号=" & hgdJobs.TextMatrix(hgdJobs.Row, 0)
        err = 0
        On Error Resume Next
        gcnOracle.Execute gstrSQL
        If err <> 0 Then
            MsgBox "作业删除失败！", vbExclamation, gstrSysName
            Exit Sub
        End If
        '插入重要操作日志
        Call SaveAuditLog(3, "删除", "删除“" & Split(cmbSystem.Text, " ")(0) & "”中的自动作业“" & hgdJobs.TextMatrix(hgdJobs.Row, 5) & "”", strRemarks)
    End With
    Call cmbSystem_Click
End Sub

Private Sub cmdSet_Click()
    Dim strParas As String
    Dim aryPara() As String
    Dim intCount As Integer
    On Error GoTo ErrHandle
    If Val(hgdJobs.TextMatrix(hgdJobs.Row, 0)) = 0 Then Exit Sub
    
    With hgdJobs
        frmAutoJobset.lblSys.Tag = cmbSystem.ItemData(cmbSystem.ListIndex)  '系统
        frmAutoJobset.txtJobName.Tag = .TextMatrix(.Row, 0)                 '序号
        frmAutoJobset.txtJobName.Text = .TextMatrix(.Row, 5)                    '名称
        frmAutoJobset.chkAutoJob.value = IIf(.TextMatrix(.Row, 6) = "是", 1, 0) '自动执行
        If .TextMatrix(.Row, 7) = "未知" Then
            frmAutoJobset.chkAutoJob.Tag = 0                                '作业号
        Else
            frmAutoJobset.chkAutoJob.Tag = .TextMatrix(.Row, 4)             '作业号
        End If
        frmAutoJobset.lblJobWhat.Caption = .TextMatrix(.Row, 3)             '内容
        frmAutoJobset.txtJobComment.Text = .TextMatrix(.Row, 1)             '说明
        frmAutoJobset.dtpStart.value = IIf(.TextMatrix(.Row, 8) = "", CurrentDate(), .TextMatrix(.Row, 8)) '执行时间
        frmAutoJobset.dtpStart.Tag = frmAutoJobset.dtpStart.value
        frmAutoJobset.txtCycle.Text = Val(.TextMatrix(.Row, 9))  '间隔时间
        frmAutoJobset.cboCycle.Tag = Replace(.TextMatrix(.Row, 9), frmAutoJobset.txtCycle.Text, "") '时间单位
        strParas = Trim(.TextMatrix(.Row, 2))
    End With
    
    With frmAutoJobset
        .imgMain.Tag = cmbSystem.Tag                             '所有者
        If optKind(2).value = True Then
            .Tag = 3                                             '类型
            .cmdWhat.Enabled = True
            .txtJobComment.Locked = False
            .txtJobComment.ForeColor = Me.ForeColor
        ElseIf optKind(1).value = True Then
            .Tag = 2
            .fraPara.Enabled = False
        Else
            .Tag = 1
        End If
        
        If strParas = "" Then
            .Height = .fraCycle.Top + .fraCycle.Height + 600
            .fraPara.Visible = False
        Else
            .fraPara.Visible = True
            aryPara = Split(strParas, ";")
            For intCount = 0 To UBound(aryPara)
                If intCount > .lblPara.UBound Then Load .lblPara(intCount)
                If intCount > .txtPara.UBound Then Load .txtPara(intCount)
                .lblPara(intCount).Top = intCount * 400 + 375
                .txtPara(intCount).Top = intCount * 400 + 315
                .lblPara(intCount).Left = .txtPara(0).Left - .lblPara(intCount).Width - 45
                .txtPara(intCount).Left = .txtPara(0).Left
                .lblPara(intCount).Caption = Left(aryPara(intCount), InStr(1, aryPara(intCount), ",") - 1)
                .txtPara(intCount).Text = Mid(aryPara(intCount), InStr(1, aryPara(intCount), ",") + 1)
                .lblPara(intCount).Visible = True
                .txtPara(intCount).Visible = True
            Next
            .fraPara.Height = (UBound(aryPara) + 1) * 400 + 375
            .Height = .fraPara.Top + .fraPara.Height + 600
        End If
        .Show 1, frmMDIMain
    End With
    
    Call cmbSystem_Click
    Exit Sub
ErrHandle:
    MsgBox err.Description, vbCritical, Me.Caption
End Sub

Private Sub cmdTest_Click()
    Dim lngSystem As Long
    Dim cnTools As ADODB.Connection
    
    With cmbSystem
        lngSystem = .ItemData(.ListIndex) \ 100
        If Val(hgdJobs.TextMatrix(hgdJobs.Row, 4)) <> 0 Then
            If .Tag = "ZLTOOLS" Then
                Set cnTools = GetConnection("ZLTOOLS")
                If cnTools Is Nothing Then Exit Sub
            Else
                Set cnTools = gcnOracle
            End If
            If optKind(2).value = True Then
                gstrSQL = "zl_JobRun(" & IIf(.ItemData(.ListIndex) = 0, "Null", .ItemData(.ListIndex)) & ",3," & hgdJobs.TextMatrix(hgdJobs.Row, 0) & ")"
            ElseIf optKind(1).value = True Then
                gstrSQL = "zl_JobRun(" & IIf(.ItemData(.ListIndex) = 0, "Null", .ItemData(.ListIndex)) & ",2," & hgdJobs.TextMatrix(hgdJobs.Row, 0) & ")"
            Else
                gstrSQL = "zl_JobRun(" & IIf(.ItemData(.ListIndex) = 0, "Null", .ItemData(.ListIndex)) & ",1," & hgdJobs.TextMatrix(hgdJobs.Row, 0) & ")"
            End If
            err = 0
            On Error Resume Next
            cnTools.Execute gstrSQL, , adCmdStoredProc
            If err <> 0 Then
                MsgBox "测试过程发生意外错误！" & vbNewLine & err.Description, vbExclamation, gstrSysName
                Exit Sub
            End If
            MsgBox "测试执行完成，如果该作业状态变为“有效”，说明执行成功。", vbInformation, gstrSysName
        End If
    End With
    Call cmbSystem_Click
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo ErrHandle
    '转入不存在的数据转移记录作为作业记录
    gstrSQL = "INSERT INTO zlAutoJobs (系统,类型,序号,名称,说明,内容,参数,执行时间,间隔时间)" & _
            " SELECT 系统,2,组号,组名,说明,'zl'||floor(系统/100)||'_DataMoveOut'||组号,日期字段||','||转出描述,to_date('2000-01-01 01:00:00','YYYY-MM-DD HH24:MI:SS'),30" & _
            " FROM zlDataMove" & _
            " WHERE (系统,组号) not in( select 系统,序号 from zlAutoJobs where 类型=2)"
    gcnOracle.Execute gstrSQL
    
    lblMain.Caption = "该工具提供用户设置调整数据库后台自动作业，主要用于需要定期批量进行的数据计算和数据修改等事务的处理。" & _
        vbCrLf & vbCrLf & "建议在系统较空闲时执行后台作业，以减少和其他任务的资源竞争，保证前台事务的运行速度。"
    
    With rsTemp
        gstrSQL = "select value" & _
                " from v$parameter" & _
                " where name='job_queue_processes'"
        If .State = adStateOpen Then .Close
        .Open gstrSQL, gcnOracle, adOpenKeyset
        mlngMaxJobs = 0
        If Not .EOF Then
            mlngMaxJobs = .Fields(0).value
            If mlngMaxJobs > 0 Then
                lbl附加.Caption = "根据数据库设置，目前最多可设置" & mlngMaxJobs & "个自动作业"
            Else
                lbl附加.Caption = "当前不能设置自动作业，如有必要，请修改数据库参数job_queue_processes"
            End If
        End If
        If mlngMaxJobs = 0 Then
            cmdTest.Enabled = False
            cmdSet.Enabled = False
            cmdAdd.Enabled = False
            cmdDel.Enabled = False
        End If
     End With
     
        'DBA可以设置zlTools的Job
    If gblnDBA Then
        Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_Reginfo", "版本号")
        If Not rsTemp.EOF Then
            cmbSystem.addItem "服务器管理工具 v" & rsTemp!内容
        Else
            cmbSystem.addItem "服务器管理工具"
        End If
    End If
        
    If gblnDBA Then
        Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_Zlsystems", "")
    Else
        Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_Zlsystems", UCase(gstrUserName))
    End If

    Do Until rsTemp.EOF
        cmbSystem.addItem rsTemp!名称 & " v" & rsTemp!版本号 & "（" & rsTemp!编号 & "）"
        cmbSystem.ItemData(cmbSystem.NewIndex) = rsTemp!编号
        rsTemp.MoveNext
    Loop
    If cmbSystem.ListCount = 0 Then cmdTest.Enabled = False
    If cmbSystem.ListCount = 1 Then cmbSystem.Locked = True
    '缺省不为服务器管理工具
    If cmbSystem.ListCount > 0 Then
        For i = 0 To cmbSystem.ListCount - 1
            If cmbSystem.ItemData(i) <> 0 Then
                cmbSystem.ListIndex = i: Exit For
            End If
        Next
        If cmbSystem.ListIndex = -1 Then cmbSystem.ListIndex = 0
    End If
    Exit Sub
ErrHandle:
    MsgBox err.Description, vbCritical, Me.Caption
 
End Sub

Private Sub Form_Resize()
    Dim sngBottom As Single
    
    On Error Resume Next
    With imgMain
        .Top = 700
        .Left = ScaleLeft + 200
    End With
    With lblMain
        .Left = lblSys.Left
        .Width = ScaleWidth - .Left - imgMain.Left
        
        lbl附加.Left = lblSys.Left
        lbl附加.Width = .Width
    End With
    
    With hgdJobs
        If ScaleWidth - .Left - 200 > cmbSystem.Left + cmbSystem.Width - .Left Then
            .Width = ScaleWidth - .Left - 200
        Else
            .Width = cmbSystem.Left + cmbSystem.Width - .Left
        End If
        sngBottom = ScaleHeight - lblMain.Height - 420 - cmdTest.Height - fraComment.Height - lbl附加.Height
        .Height = IIf(sngBottom - .Top > 2500, sngBottom - .Top, 2500)
    End With
    
    fraComment.Width = hgdJobs.Width
    fraComment.Top = hgdJobs.Top + hgdJobs.Height
    lbl说明.Width = fraComment.Width - lbl说明.Left - 300
    
    cmdDel.Left = hgdJobs.Left + hgdJobs.Width - cmdDel.Width
    cmdAdd.Left = cmdDel.Left - cmdAdd.Width
    cmdTest.Top = fraComment.Top + fraComment.Height + 60
    cmdSet.Top = cmdTest.Top
    cmdAdd.Top = cmdTest.Top
    cmdDel.Top = cmdTest.Top
    
    lblMain.Top = cmdTest.Top + cmdTest.Height + 200
    lbl附加.Top = lblMain.Top + lblMain.Height + 60
    
End Sub

Private Sub hgdJobs_RowColChange()
    With hgdJobs
        lbl说明.Caption = .TextMatrix(.Row, 1)
        lblPara.Caption = "参数：" & .TextMatrix(.Row, 2)
        If .TextMatrix(.Row, 6) = "是" Then
            cmdTest.Enabled = True
        Else
            cmdTest.Enabled = False
        End If
    End With
End Sub

Private Sub optKind_Click(Index As Integer)
    optKind(1).Tag = Index
    Call cmbSystem_Click
End Sub

Public Function SupportPrint() As Boolean
'返回本窗口是否支持打印，供主窗口调用
    SupportPrint = True
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'供主窗口调用，实现具体的打印工作
'如果没有可打印的，就留下一个空的接口

'参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    
    On Error GoTo ErrHandle
    objPrint.Title.Text = "后台作业"
    
    objRow.Add "应用系统：" & cmbSystem.Text
    objRow.Add "作业目录：" & Switch(optKind(0).value, "系统设定", optKind(1).value, "数据转移", True, "用户自定")
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "打印时间：" & Format(CurrentDate, "yyyy年MM月dd日")
    Set objPrint.Body = hgdJobs
    objPrint.BelowAppRows.Add objRow
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
    Exit Sub
ErrHandle:
    MsgBox err.Description, vbCritical, Me.Caption
End Sub
