VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAutoJobset 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "自动作业设置"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7035
   ControlBox      =   0   'False
   Icon            =   "frmAutoJobset.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox pic背景 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   960
      ScaleHeight     =   2265
      ScaleWidth      =   4050
      TabIndex        =   19
      Top             =   1170
      Visible         =   0   'False
      Width           =   4080
      Begin VB.Label lbl说明 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "编号100的系统中增加一个自动调价的作业"
         Height          =   180
         Index           =   2
         Left            =   450
         TabIndex        =   27
         Top             =   1740
         Width           =   3330
      End
      Begin VB.Label lbl说明 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ZL100_USERJOB自动调价"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   1
         Left            =   450
         TabIndex        =   26
         Top             =   1980
         Width           =   1890
      End
      Begin VB.Label lbl说明 
         BackStyle       =   0  'Transparent
         Caption         =   "命名规则中蓝体部分由用户输入；对服务器管理用户，不需要系统号。"
         Height          =   345
         Index           =   0
         Left            =   480
         TabIndex        =   25
         Top             =   1005
         Width           =   3345
      End
      Begin VB.Label lbl标题 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "举例"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   2
         Left            =   90
         TabIndex        =   24
         Top             =   1440
         Width           =   390
      End
      Begin VB.Label lbl标题 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "说明"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   1
         Left            =   90
         TabIndex        =   23
         Top             =   810
         Width           =   390
      End
      Begin VB.Label lbl用户 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[系统号]        功能"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   690
         TabIndex        =   22
         Top             =   450
         Width           =   2100
      End
      Begin VB.Label lbl固定 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ZL        _USERJOB"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   480
         TabIndex        =   21
         Top             =   450
         Width           =   1890
      End
      Begin VB.Label lbl标题 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "用户自定作业命名规则"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   20
         Top             =   150
         Width           =   1950
      End
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "更新参数"
      Height          =   350
      Left            =   5370
      TabIndex        =   28
      Top             =   1560
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CheckBox chk规则 
      Caption         =   "命名规则"
      Height          =   350
      Left            =   5370
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1170
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.TextBox txtJobName 
      Height          =   300
      Left            =   1320
      TabIndex        =   0
      Top             =   75
      Width           =   3810
   End
   Begin VB.TextBox txtJobComment 
      ForeColor       =   &H00808080&
      Height          =   1230
      Left            =   900
      Locked          =   -1  'True
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   870
      Width           =   4215
   End
   Begin VB.CommandButton cmdWhat 
      Caption         =   "…"
      Enabled         =   0   'False
      Height          =   300
      Left            =   4830
      TabIndex        =   1
      Top             =   450
      Width           =   285
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5370
      TabIndex        =   15
      Top             =   480
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5370
      TabIndex        =   14
      Top             =   60
      Width           =   1100
   End
   Begin VB.Frame fraPara 
      Caption         =   "执行参数"
      Height          =   840
      Left            =   900
      TabIndex        =   12
      Top             =   3690
      Width           =   4215
      Begin VB.TextBox txtPara 
         Height          =   300
         Index           =   0
         Left            =   1035
         TabIndex        =   6
         Top             =   315
         Width           =   2010
      End
      Begin VB.Label lblPara 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "登记时间"
         Height          =   180
         Index           =   0
         Left            =   285
         TabIndex        =   13
         Top             =   375
         Width           =   720
      End
   End
   Begin VB.Frame fraCycle 
      Caption         =   "执行周期"
      Height          =   1080
      Left            =   900
      TabIndex        =   9
      Top             =   2535
      Width           =   4215
      Begin MSComCtl2.DTPicker dtpStart 
         Height          =   300
         Left            =   2100
         TabIndex        =   4
         Top             =   645
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "HH:mm:ss"
         Format          =   106168323
         UpDown          =   -1  'True
         CurrentDate     =   37031.0416666667
      End
      Begin VB.ComboBox cboMonth 
         Height          =   300
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   645
         Width           =   900
      End
      Begin VB.ComboBox cboDay 
         Height          =   300
         Left            =   2370
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   645
         Width           =   1030
      End
      Begin VB.ComboBox cboWeek 
         Height          =   300
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   645
         Width           =   1030
      End
      Begin VB.ComboBox cboCycle 
         Height          =   300
         Left            =   2505
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   225
         Width           =   720
      End
      Begin VB.TextBox txtCycle 
         Height          =   300
         Left            =   1080
         TabIndex        =   5
         Top             =   225
         Width           =   1425
      End
      Begin VB.Label lblCycle 
         AutoSize        =   -1  'True
         Caption         =   "循环时间"
         Height          =   180
         Left            =   285
         TabIndex        =   11
         Top             =   270
         Width           =   720
      End
      Begin VB.Label lblStart 
         AutoSize        =   -1  'True
         Caption         =   "执行时间"
         Height          =   180
         Left            =   285
         TabIndex        =   10
         Top             =   705
         Width           =   720
      End
   End
   Begin VB.CheckBox chkAutoJob 
      Caption         =   "设置为后台自动作业(&A)"
      Height          =   210
      Left            =   900
      TabIndex        =   3
      Top             =   2190
      Width           =   2850
   End
   Begin VB.Label lblComment 
      AutoSize        =   -1  'True
      Caption         =   "说明"
      Height          =   180
      Left            =   450
      TabIndex        =   17
      Top             =   900
      Width           =   360
   End
   Begin VB.Label lblJobWhat 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   1320
      TabIndex        =   7
      Top             =   450
      Width           =   3525
   End
   Begin VB.Label lblWhat 
      AutoSize        =   -1  'True
      Caption         =   "内容"
      Height          =   180
      Left            =   900
      TabIndex        =   16
      Top             =   510
      Width           =   360
   End
   Begin VB.Image imgMain 
      Height          =   480
      Left            =   255
      Picture         =   "frmAutoJobset.frx":000C
      Top             =   195
      Width           =   480
   End
   Begin VB.Label lblSys 
      AutoSize        =   -1  'True
      Caption         =   "作业"
      Height          =   180
      Left            =   900
      TabIndex        =   8
      Top             =   150
      Width           =   360
   End
   Begin VB.Menu mnuProcedures 
      Caption         =   "Procedure"
      Visible         =   0   'False
      Begin VB.Menu mnuWhat 
         Caption         =   "mnuWhat"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmAutoJobset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsTemp As New ADODB.Recordset
Dim strSQL As String
Dim intCount As Integer
Dim strOrder As String, strParas As String
Dim aryPara() As String
Private mdateNow As Date

Private Enum DateUnit
    DU_天 = 0
    DU_周 = 1
    DU_月 = 2
    DU_季度 = 3
End Enum

Private Sub cboCycle_Click()
    Dim rsTmp As ADODB.Recordset
    Dim lngDay As Long
    Dim lngMonth As Long
    Dim lngMaxDay As Long
    
    Select Case cboCycle.ListIndex
    Case DU_天
        cboMonth.Visible = False
        cboWeek.Visible = False
        cboDay.Visible = False
        dtpStart.Width = 2145
        txtCycle.Width = 1425
        cboCycle.Left = txtCycle.Left + txtCycle.Width
        dtpStart.Left = txtCycle.Left
        
        If cboCycle.Text = cboCycle.Tag Then
            dtpStart.value = dtpStart.Tag
        Else
            dtpStart.value = "2001/5/20 1:00:00"
        End If
    Case DU_周
        cboMonth.Visible = False
        cboWeek.Visible = True
        cboDay.Visible = False
        dtpStart.Width = 1125
        txtCycle.Width = 1425
        cboCycle.Left = txtCycle.Left + txtCycle.Width
        cboWeek.Left = txtCycle.Left
        dtpStart.Left = cboWeek.Left + cboWeek.Width - 20
        
        If cboCycle.Text = cboCycle.Tag Then
            cboWeek.ListIndex = Weekday(CDate(dtpStart.Tag)) - 1
            dtpStart.value = dtpStart.Tag
        Else
            cboWeek.ListIndex = 1
            dtpStart.value = "2001/5/20 1:00:00"
        End If
    Case DU_月
        cboMonth.Visible = False
        cboWeek.Visible = False
        cboDay.Visible = True
        dtpStart.Width = 1125
        txtCycle.Width = 1425
        cboCycle.Left = txtCycle.Left + txtCycle.Width
        cboDay.Left = txtCycle.Left
        dtpStart.Left = cboDay.Left + cboDay.Width - 20
        
        If cboCycle.Text = cboCycle.Tag Then
            '获取指定月最大天数
            lngMaxDay = Right(DateSerial(Year(dtpStart.Tag), Month(dtpStart.Tag) + 1, 0), 2)
            lngDay = Format(dtpStart.Tag, "d")
            If lngDay <= 28 Then
                cboDay.Text = lngDay & "日"
            ElseIf lngDay = lngMaxDay Then
                cboDay.Text = "月末"
            ElseIf lngDay = lngMaxDay - 1 Then
                cboDay.Text = "月末-1"
            ElseIf lngDay = lngMaxDay - 2 Then
                cboDay.Text = "月末-2"
            End If
            dtpStart.value = dtpStart.Tag
        Else
            cboDay.ListIndex = 0
            dtpStart.value = "2001/5/20 1:00:00"
        End If
    Case DU_季度
        cboWeek.Visible = False
        cboMonth.Visible = True
        cboDay.Visible = True
        dtpStart.Width = 1125
        txtCycle.Width = 2310
        cboCycle.Left = txtCycle.Left + txtCycle.Width
        cboMonth.Left = txtCycle.Left
        cboDay.Left = cboMonth.Left + cboMonth.Width - 20
        dtpStart.Left = cboDay.Left + cboDay.Width - 20
        
        If cboCycle.Text = cboCycle.Tag Then
            '获得指定月是第几个月
            lngMonth = Format(dtpStart.Tag, "M") Mod 3 - 1
            If lngMonth = 0 Then
                cboMonth.Text = "第一月"
            ElseIf lngMonth = 1 Then
                cboMonth.Text = "第二月"
            Else
                lngMonth = 2
                cboMonth.Text = "第三月"
            End If
            '获取指定月最大天数
            lngMaxDay = Right(DateSerial(Year(CDate(dtpStart.Tag)), Month(CDate(dtpStart.Tag)) + 1, 0), 2)
            lngDay = Format(dtpStart.Tag, "d")
            If lngDay <= 28 Then
                cboDay.Text = lngDay & "日"
            ElseIf lngDay = lngMaxDay Then
                cboDay.Text = "月末"
            ElseIf lngDay = lngMaxDay - 1 Then
                cboDay.Text = "月末-1"
            ElseIf lngDay = lngMaxDay - 2 Then
                cboDay.Text = "月末-2"
            End If
            dtpStart.value = dtpStart.Tag
        Else
            cboMonth.ListIndex = 0
            cboDay.ListIndex = 0
            dtpStart.value = "2001/5/20 1:00:00"
        End If
        
        '存入当前季度中第一月的月份
        cboMonth.Tag = Format(mdateNow, "M") - lngMonth
    End Select
End Sub

Private Sub chk规则_Click()
    pic背景.Visible = chk规则.value = 1
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strExecuteTime As String, strQuarterly As String
    Dim rsTmp As ADODB.Recordset
    Dim lngMaxDay As Long
    Dim cnTools As ADODB.Connection
    
    If Trim(lblJobWhat.Caption) = "" Then
        MsgBox "未设置作业内容！", vbExclamation, gstrSysName
        Exit Sub
    End If
    If Val(txtCycle.Text) = 0 Then
        MsgBox "未正确设置作业循环时间！", vbExclamation, gstrSysName
        txtCycle.SetFocus: Exit Sub
    End If
    
    strParas = ""
    If fraPara.Visible Then
        For intCount = 0 To lblPara.UBound
            If lblPara(intCount).Visible = False Then Exit For
            If Trim(txtPara(intCount).Text) = "" Then
                MsgBox lblPara(intCount).Caption & " 参数未指定值！", vbExclamation, gstrSysName
                Exit Sub
            End If
            strParas = strParas & ";" & lblPara(intCount).Caption & "," & txtPara(intCount).Text
        Next
    End If
    If strParas <> "" Then strParas = Mid(strParas, 2)
    
    '将获取到的执行日期信息转换为具体的日期
    Select Case cboCycle.ListIndex
    Case DU_天
        strExecuteTime = Format(mdateNow, "yyyy-MM-dd") & " " & Format(dtpStart.value, "HH:mm:ss")
    Case DU_周
        strExecuteTime = Format(DateAdd("d", cboWeek.ListIndex + 1 - Weekday(mdateNow), mdateNow), "yyyy-MM-dd") & " " & Format(dtpStart.value, "HH:mm:ss")
    Case DU_月
        If cboDay.ListIndex <= 27 Then
            strExecuteTime = Format(mdateNow, "yyyy-MM") & "-" & Val(cboDay.Text) & " " & Format(dtpStart.value, "HH:mm:ss")
        Else
            lngMaxDay = Right(DateSerial(Year(mdateNow), Month(mdateNow) + 1, 0), 2)
            strExecuteTime = Format(mdateNow, "yyyy-MM") & "-" & lngMaxDay - (cboDay.ListCount - cboDay.ListIndex - 1) & " " & Format(dtpStart.value, "HH:mm:ss")
        End If
    Case DU_季度
        If cboDay.ListIndex <= 27 Then
            strExecuteTime = Format(mdateNow, "yyyy") & "-" & cboMonth.Tag + cboMonth.ListIndex & "-" & Val(cboDay.Text) & " " & Format(dtpStart.value, "HH:mm:ss")
        Else
            strQuarterly = Format(mdateNow, "yyyy") & "-" & cboMonth.Tag + cboMonth.ListIndex & "-" & "01 11:11:11"
            lngMaxDay = Right(DateSerial(Year(CDate(strQuarterly)), Month(CDate(strQuarterly)) + 1, 0), 2)
            strExecuteTime = Format(mdateNow, "yyyy") & "-" & cboMonth.Tag + cboMonth.ListIndex & "-" & lngMaxDay - (cboDay.ListCount - cboDay.ListIndex - 1) & " " & Format(dtpStart.value, "HH:mm:ss")
        End If
    End Select
    
    If Tag = "ADD" Then
        Dim rsOut As New ADODB.Recordset
        '取ZlAutoJob序列号
        Set rsOut = OpenCursor(gcnOracle, "ZLTOOLS.B_Runmana.Get_Job_number", Val(lblSys.Tag))
        If rsOut.RecordCount > 0 Then
            strOrder = Nvl(Val(rsOut.Fields(0)), 1)
        Else
            strOrder = 1
        End If
        strSQL = "insert into zlAutoJobs(系统,类型,序号,名称,说明,内容,参数,执行时间,间隔时间,时间单位)" & _
                " values (" & IIf(Val(lblSys.Tag) = 0, "Null", lblSys.Tag) & ",3," & Val(strOrder) & "," & _
                "       '" & txtJobName.Text & "'," & _
                "       '" & txtJobComment.Text & "'," & _
                "       '" & lblJobWhat.Caption & "'," & _
                "       '" & strParas & "'," & _
                "       to_date('" & strExecuteTime & "','yyyy-MM-dd HH24:MI:SS')," & _
                "       " & Val(txtCycle.Text) & _
                "       ,'" & cboCycle.Text & "')"
    Else
        strSQL = "update zlAutoJobs" & _
                " set 名称='" & txtJobName.Text & "'," & _
                "     说明='" & txtJobComment.Text & "'," & _
                "     内容='" & lblJobWhat.Caption & "'," & _
                "     参数='" & strParas & "'," & _
                "     执行时间=to_date('" & strExecuteTime & "','yyyy-MM-dd HH24:MI:SS')," & _
                "     间隔时间=" & Val(txtCycle.Text) & "," & _
                "     时间单位='" & cboCycle.Text & "'" & _
                " Where Nvl(系统,0)=" & Val(lblSys.Tag) & _
                "     and 类型=" & Tag & _
                "     and 序号=" & txtJobName.Tag
    End If
    err = 0
    On Error Resume Next
    gcnOracle.Execute strSQL
    If err <> 0 Then
        MsgBox "作业设置保存失败，请检查设置情况！" & vbNewLine & err.Description, vbExclamation, gstrSysName
        Exit Sub
    End If
    If Tag = "ADD" Then
        '插入重要操作日志
        Call SaveAuditLog(1, "增加", "在“" & Split(frmAutoJobs.cmbSystem.Text, " ")(0) & "”添加自动作业“" & txtJobName.Text & "”")
    Else
        '插入重要操作日志
        Call SaveAuditLog(2, "运行设置", "修改“" & Split(frmAutoJobs.cmbSystem.Text, " ")(0) & "”中的自动作业“" & txtJobName.Text & "”")
    End If
    err = 0
    If imgMain.Tag = "ZLTOOLS" Then
        Set cnTools = GetConnection("ZLTOOLS")
        If cnTools Is Nothing Then Exit Sub
    Else
        Set cnTools = gcnOracle
    End If
    If chkAutoJob.value = 1 Then
        If Tag = "ADD" Then                      '新作业
            strSQL = "zl" & "_JobSubmit(" & IIf(Val(lblSys.Tag) = 0, "Null", lblSys.Tag) & ",3," & Val(strOrder) & ")"
        ElseIf Val(chkAutoJob.Tag) = 0 Then      '首次设置为自动作业
            strSQL = "zl" & "_JobSubmit(" & IIf(Val(lblSys.Tag) = 0, "Null", lblSys.Tag) & "," & Tag & "," & txtJobName.Tag & ")"
        Else                                        '修改已经启用的作业
            strSQL = "zl" & "_JobChange(" & IIf(Val(lblSys.Tag) = 0, "Null", lblSys.Tag) & "," & Tag & "," & txtJobName.Tag & ")"
        End If
        cnTools.Execute strSQL, , adCmdStoredProc
    Else
        If Val(chkAutoJob.Tag) <> 0 Then         '取消自动作业
            strSQL = "zl" & "_JobRemove(" & IIf(Val(lblSys.Tag) = 0, "Null", lblSys.Tag) & "," & Tag & "," & txtJobName.Tag & ")"
            cnTools.Execute strSQL, , adCmdStoredProc
        End If
    End If
    If err <> 0 Then
        MsgBox "虽然作业设置保存，但未能成功设置为自动作业。请检查数据库系统！", vbExclamation, gstrSysName
    End If
    
    Unload Me
End Sub

Private Sub cmdUpdate_Click()
    Dim rsTemp As New ADODB.Recordset
On Error GoTo errHandle
    
    If MsgBox("是否根据数据归档转移处设置的时间更新参数？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub

    Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.B_Runmana.Get_depict", Val(lblSys.Tag), Val(txtJobName.Tag))
    If rsTemp.RecordCount > 0 Then
        txtPara(0).Text = Val(IIf(IsNull(rsTemp.Fields(0)), "150", rsTemp.Fields(0)))
    Else
        txtPara(0).Text = 150
    End If
    Exit Sub
errHandle:
    MsgBox err.Description, vbCritical, Me.Caption
End Sub

Private Sub cmdWhat_Click()
   Dim cnTools As ADODB.Connection
On Error GoTo errHandle
    If Val(cmdWhat.Tag) = 0 Then
        If imgMain.Tag = "ZLTOOLS" Then
            Set cnTools = GetConnection("ZLTOOLS")
            If cnTools Is Nothing Then Exit Sub
        Else
            Set cnTools = gcnOracle
        End If
        Set rsTemp = cnTools.Execute("SELECT Object_Name  From All_Objects " & vbNewLine & _
                                      "WHERE Object_Type = 'PROCEDURE' AND Object_Name LIKE 'ZL" & CStr(IIf(Val(lblSys.Tag) = 0, "", lblSys.Tag)) & "_USERJOB%' " & vbNewLine & _
                                      " AND Status = 'VALID' AND Owner = '" & CStr(imgMain.Tag) & "'")
        With rsTemp
            Do While Not .EOF
                If .AbsolutePosition - 1 > mnuWhat.UBound Then Load mnuWhat(.AbsolutePosition - 1)
                mnuWhat(.AbsolutePosition - 1).Caption = .Fields(0).value
                mnuWhat(.AbsolutePosition - 1).Visible = True
                .MoveNext
            Loop
            cmdWhat.Tag = .RecordCount
        End With
    End If
    If Val(cmdWhat.Tag) > 0 Then
        PopupMenu mnuProcedures, 2
    Else
        MsgBox "没有可选的存储过程", vbExclamation, gstrSysName
    End If
    Exit Sub
errHandle:
    MsgBox err.Description, vbCritical, Me.Caption
End Sub

Private Sub Form_Activate()
    Dim i As Long
    
    If frmAutoJobset.Tag = "2" Then cmdUpdate.Visible = True
    cboCycle.Clear
    cboCycle.addItem "天"
    cboCycle.addItem "周"
    cboCycle.addItem "月"
    cboCycle.addItem "季度"
    cboWeek.Clear
    cboWeek.addItem "星期日"
    cboWeek.addItem "星期一"
    cboWeek.addItem "星期二"
    cboWeek.addItem "星期三"
    cboWeek.addItem "星期四"
    cboWeek.addItem "星期五"
    cboWeek.addItem "星期六"
    cboMonth.Clear
    cboMonth.addItem "第一月"
    cboMonth.addItem "第二月"
    cboMonth.addItem "第三月"
    cboDay.Clear
    For i = 1 To 28
        cboDay.addItem i & "日"
    Next
    cboDay.addItem "月末-2"
    cboDay.addItem "月末-1"
    cboDay.addItem "月末"
    
    '将当前数据库时间存入变量
    mdateNow = CurrentDate()
    
    cboCycle.Text = IIf(cboCycle.Tag = "", "天", cboCycle.Tag)
End Sub

Private Sub mnuWhat_Click(Index As Integer)
    On Error GoTo errHandle
    lblJobWhat.Caption = mnuWhat(Index).Caption
    With rsTemp
        If gblnDBA Then
            strSQL = "select rtrim(ltrim(upper(text))) from dba_source where name='" & mnuWhat(Index).Caption & "' and OWNER='" & imgMain.Tag & "'"
        Else
            strSQL = "select rtrim(ltrim(upper(text))) from user_source where name='" & mnuWhat(Index).Caption & "'"
        End If
        If .State = adStateOpen Then .Close
        .Open strSQL, gcnOracle, adOpenKeyset
        strSQL = ""
        Do While Not .EOF
            strSQL = strSQL & " " & Replace(Replace(Replace(Replace(Trim(.Fields(0).value), vbCrLf, " "), vbCr, " "), vbLf, " "), vbTab, " ")
            If InStr(1, strSQL, " AS ") > 0 Then Exit Do
            If InStr(1, strSQL, " IS ") > 0 Then Exit Do
            If InStr(1, strSQL, ")AS ") > 0 Then Exit Do
            If InStr(1, strSQL, ")IS ") > 0 Then Exit Do
            If Right(strSQL, 3) = " AS" Then Exit Do
            If Right(strSQL, 3) = " IS" Then Exit Do
            If Right(strSQL, 3) = ")AS" Then Exit Do
            If Right(strSQL, 3) = ")IS" Then Exit Do
            .MoveNext
        Loop
        strSQL = Replace(Replace(Replace(Replace(strSQL, vbCrLf, " "), vbCr, " "), vbLf, " "), vbTab, " ")
        If InStr(1, strSQL, "(") > 0 Then
            strSQL = Mid(strSQL, InStr(1, strSQL, "(") + 1)
            strSQL = Left(strSQL, InStr(1, strSQL, ")") - 1)
        Else
            strSQL = ""
        End If
        
        For intCount = 0 To lblPara.UBound
            lblPara(intCount).Visible = False
            txtPara(intCount).Visible = False
        Next
    
        If strSQL = "" Then
            Height = fraCycle.Top + fraCycle.Height + 600
            fraPara.Visible = False
        Else
            fraPara.Visible = True
            aryPara = Split(strSQL, ",")
            For intCount = 0 To UBound(aryPara)
                aryPara(intCount) = Trim(aryPara(intCount))
                If intCount > lblPara.UBound Then Load lblPara(intCount)
                If intCount > txtPara.UBound Then Load txtPara(intCount)
                lblPara(intCount).Top = intCount * 400 + 375
                txtPara(intCount).Top = intCount * 400 + 315
                lblPara(intCount).Left = txtPara(0).Left - lblPara(intCount).Width - 45
                txtPara(intCount).Left = txtPara(0).Left
                lblPara(intCount).Caption = Left(aryPara(intCount), InStr(1, aryPara(intCount), " ") - 1)
                txtPara(intCount).Text = ""
                lblPara(intCount).Visible = True
                txtPara(intCount).Visible = True
            Next
            fraPara.Height = (UBound(aryPara) + 1) * 400 + 375
            Height = fraPara.Top + fraPara.Height + 600
        End If
    
    End With
    Exit Sub
errHandle:
    MsgBox err.Description, vbCritical, Me.Caption
End Sub

Private Sub txtCycle_KeyPress(KeyAscii As Integer)
    If Not (InStr("0123456789", Chr(KeyAscii)) > 0 Or KeyAscii = 8) Then
        KeyAscii = 0
    End If
End Sub
