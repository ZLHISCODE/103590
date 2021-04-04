VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDiffPriceRecalCard 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "药品差价计算"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   Icon            =   "frmDiffPriceRecalCard.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdInput 
      Caption         =   "录入差价(&E)"
      Height          =   350
      Left            =   3525
      TabIndex        =   22
      Top             =   3840
      Visible         =   0   'False
      Width           =   1200
   End
   Begin MSComCtl2.DTPicker dtpBegin 
      Height          =   285
      Left            =   1245
      TabIndex        =   20
      Top             =   3270
      Visible         =   0   'False
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "yyyy年MM月dd日 HH:mm:ss"
      Format          =   114491395
      CurrentDate     =   36444
      MaxDate         =   401768
   End
   Begin VB.CommandButton cmdIni 
      Caption         =   "初始结存(&I)"
      Height          =   350
      Left            =   45
      TabIndex        =   17
      Top             =   3840
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "查看(&B)"
      Height          =   350
      Left            =   2430
      TabIndex        =   14
      Top             =   3840
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CommandButton cmdVerify 
      Caption         =   "审核结存(&V)"
      Height          =   350
      Left            =   45
      TabIndex        =   13
      Top             =   3840
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "取消结存(&D)"
      Height          =   350
      Left            =   1245
      TabIndex        =   12
      Top             =   3840
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.ComboBox cbo核算方法 
      Height          =   300
      Left            =   5250
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   2895
      Width           =   1605
   End
   Begin VB.CommandButton CmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   7065
      TabIndex        =   5
      Top             =   3840
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   30
      Left            =   -840
      TabIndex        =   4
      Top             =   3615
      Width           =   9060
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5970
      TabIndex        =   2
      Top             =   3840
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "核算(&O)"
      Height          =   350
      Left            =   4860
      TabIndex        =   1
      Top             =   3840
      Width           =   1100
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   4395
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDiffPriceRecalCard.frx":000C
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9816
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
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
   Begin VB.ComboBox cbo库房 
      Height          =   300
      Left            =   1245
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2910
      Width           =   1605
   End
   Begin MSComCtl2.DTPicker dtpTime 
      Height          =   285
      Left            =   5250
      TabIndex        =   16
      Top             =   3255
      Visible         =   0   'False
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "yyyy年MM月dd日 HH:mm:ss"
      Format          =   114491395
      CurrentDate     =   36444
      MaxDate         =   401768
   End
   Begin MSComCtl2.DTPicker dtpEnd 
      Height          =   285
      Left            =   5250
      TabIndex        =   21
      Top             =   3255
      Visible         =   0   'False
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "yyyy年MM月dd日 HH:mm:ss"
      Format          =   114491395
      CurrentDate     =   36444
      MaxDate         =   401768
   End
   Begin VB.Label lblEnd 
      AutoSize        =   -1  'True
      Caption         =   "结束时间"
      Height          =   180
      Left            =   4290
      TabIndex        =   19
      Top             =   3315
      Width           =   720
   End
   Begin VB.Label lblBegin 
      AutoSize        =   -1  'True
      Caption         =   "开始时间"
      Height          =   180
      Left            =   420
      TabIndex        =   18
      Top             =   3345
      Width           =   720
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      Caption         =   "本次结存时间"
      Height          =   180
      Left            =   4125
      TabIndex        =   15
      Top             =   3315
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Label lbl上次结存 
      Caption         =   "2007-01-01 22:00:00(未审核)"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   1245
      TabIndex        =   11
      Top             =   3345
      Visible         =   0   'False
      Width           =   6660
   End
   Begin VB.Label lbl结存 
      AutoSize        =   -1  'True
      Caption         =   "上次结存时间"
      Height          =   180
      Left            =   75
      TabIndex        =   10
      Top             =   3360
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "核算方法"
      Height          =   180
      Left            =   4290
      TabIndex        =   8
      Top             =   2955
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "库房"
      Height          =   180
      Left            =   450
      TabIndex        =   6
      Top             =   2985
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   90
      Picture         =   "frmDiffPriceRecalCard.frx":08A0
      Top             =   105
      Width           =   480
   End
   Begin VB.Label lblMemo 
      Caption         =   $"frmDiffPriceRecalCard.frx":0CE2
      ForeColor       =   &H00C00000&
      Height          =   2505
      Left            =   630
      TabIndex        =   3
      Top             =   75
      Width           =   7620
   End
End
Attribute VB_Name = "frmDiffPriceRecalCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mstr上次结存时间 As String
Dim mbln是否审核 As Boolean
Dim mlng库房ID As Long
Dim mint核算方法 As Integer
Dim mint单位系数 As Integer
Dim mbln仅初始结存 As Boolean

Private Const intIni As Integer = 6

Private Enum con核算方法
    type_移动平均 = 1
    type_全月平均 = 2
    type_先进先出 = 3
End Enum
Private Sub GetUnit(ByVal lng库房ID As Long)
    Dim strUnit As String
    strUnit = GetDrugUnit(lng库房ID, Me.Caption)
    Select Case strUnit
        Case "住院单位"
            mint单位系数 = 4
        Case "门诊单位"
            mint单位系数 = 3
        Case "药库单位"
            mint单位系数 = 2
        Case "售价单位"
            mint单位系数 = 1
    End Select
End Sub

Private Sub Get上次结存(ByVal lng库房ID As Long)
    Dim rsTmp As New ADODB.Recordset
    
    mbln仅初始结存 = False
    
    On Error GoTo errHandle
    '如果选择先进先出，则显示上次结存信息
    If cbo核算方法.ListIndex = 1 Then
        lblTime.Visible = True
        dtpTime.Visible = True
        dtpTime.Value = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        
        '检查是否仅存在初始结存
        gstrSQL = "Select nvl(是否初始,0) 是否初始 From 药品结存 Where Nvl(是否初始, 0) = 1 And 库房id = [1] And Rownum = 1" & _
                " Union All " & _
                " Select nvl(是否初始,0) 是否初始 From 药品结存 Where Nvl(是否初始, 0) = 0 And 库房id = [1] And Rownum = 1"
        Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "-检查是否存在初始结存", lng库房ID)
        If rsTmp.RecordCount = 1 Then
            If rsTmp!是否初始 = 1 Then
                mbln仅初始结存 = True
            End If
        End If
        cmdInput.Visible = mbln仅初始结存
        
        '检查是否存在上次结存
        gstrSQL = "Select Max(结存日期) 结存时间 From 药品结存  Where 库房id=[1] "
        Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "-检查是否存在上次结存", lng库房ID)
        
        mstr上次结存时间 = Format(rsTmp!结存时间, "YYYY-MM-DD HH:MM:SS")
        
        If mstr上次结存时间 = "" Then
            lbl结存.Visible = True
            lbl上次结存.Visible = True
            cmdIni.Visible = True
            lbl上次结存.Caption = "无初始结存信息，请初始化结存！"
            Exit Sub
        End If
        
        '取上次结存信息
        gstrSQL = "Select Nvl(结存标志, 0) 结存标志 From 药品结存 " & _
             " Where 库房id = [1] And 结存日期 = [2] And Rownum = 1"
        Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "-取上次结存", lng库房ID, CDate(Format(mstr上次结存时间, "yyyy-mm-dd hh:mm:ss")))
        
        If rsTmp.RecordCount = 0 Then
            mstr上次结存时间 = ""
            lbl结存.Visible = True
            lbl上次结存.Visible = True
            cmdIni.Visible = True
            lbl上次结存.Caption = "无初始结存信息，请初始化结存！"
            Exit Sub
        Else
            mbln是否审核 = (rsTmp!结存标志 = 1)
        End If
        
        '显示上次结存信息
        If mstr上次结存时间 <> "" Then
            lbl结存.Visible = True
            lbl上次结存.Visible = True
            cmdVerify.Visible = True
            cmdDel.Visible = True
            cmdBrowse.Visible = True
            lbl上次结存.Caption = mstr上次结存时间 & IIf(mbln是否审核, "", "(未审核)")
            cmdVerify.Enabled = Not mbln是否审核
        End If
        
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub IniControl()
    lbl结存.Visible = False
    lbl上次结存.Visible = False
    cmdVerify.Visible = False
    cmdDel.Visible = False
    cmdBrowse.Visible = False
    lblTime.Visible = False
    dtpTime.Visible = False
    cmdIni.Visible = False
    
    lblBegin.Visible = False
    lblEnd.Visible = False
    dtpBegin.Visible = False
    dtpEnd.Visible = False
End Sub

Private Sub RefreshNow(ByVal lng库房ID As Long)
    Call IniControl
    
    If mint核算方法 = type_全月平均 Then
        lblBegin.Visible = True
        lblEnd.Visible = True
        dtpBegin.Visible = True
        dtpEnd.Visible = True
        
        dtpBegin.Value = Format(Sys.Currentdate, "yyyy-mm") & "-01 00:00:00"
        dtpEnd.Value = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        Exit Sub
    End If
    
    If mint核算方法 = type_先进先出 Then
        Call Get上次结存(lng库房ID)
    End If
End Sub

Private Sub cbo核算方法_Click()
    Select Case cbo核算方法.ListIndex
        Case 0
            '全月平均
            mint核算方法 = type_全月平均
        Case 1
            '先进先出
            mint核算方法 = type_先进先出
    End Select
    
    Call RefreshNow(mlng库房ID)
End Sub


Private Sub cbo库房_Click()
    If Cbo库房.ItemData(Cbo库房.ListIndex) <> mlng库房ID Then
        mlng库房ID = Cbo库房.ItemData(Cbo库房.ListIndex)
        Call GetUnit(mlng库房ID)
        Call RefreshNow(mlng库房ID)
    End If
End Sub


Private Sub cmdBrowse_Click()
    Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1308", Me, "库房=" & Cbo库房.Text & "|" & IIf(Cbo库房.ItemData(Cbo库房.ListIndex) = 0, " is not null ", "=" & Cbo库房.ItemData(Cbo库房.ListIndex)), "结存日期=" & CDate(mstr上次结存时间), "单位=" & mint单位系数)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDel_Click()
    If MsgBox("是否删除上次结存？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    End If
    
    gstrSQL = "Zl_药品结存_Delete(to_date('" & Format(mstr上次结存时间, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')," & mlng库房ID & ",2)"
    Me.staThis.Panels(2).Text = "正在删除上次结存，请等待。。。！"
    
    Me.MousePointer = vbHourglass
    
    Call zlDataBase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    Me.MousePointer = vbDefault
    MsgBox "删除成功！", vbOKOnly + vbInformation, gstrSysName
    
    Call IniControl
    Call Get上次结存(mlng库房ID)
    
    DoEvents
    Me.staThis.Panels(2).Text = ""
    
    Exit Sub
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub cmdIni_Click()
    '如果采用先进先出法，但未有初始积存，则要生成初始结存
    If mstr上次结存时间 = "" And DateDiff("m", dtpTime.Value, Sys.Currentdate) > intIni Then
        MsgBox "期初结存日期不能早于" & intIni & "个月。"
        dtpTime.Value = Sys.Currentdate
        Exit Sub
    End If
    
    If mint核算方法 = type_先进先出 And mstr上次结存时间 = "" Then
        gstrSQL = "Zl_药品结存_Insert(to_date('" & Format(dtpTime.Value, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss') ," & mlng库房ID & ",NULL)"
    Else
        Exit Sub
    End If
    
    On Error GoTo errHandle
    
    Me.staThis.Panels(2).Text = "正在初始药品结存，请等待。。。！"
    
    Me.MousePointer = vbHourglass
    
    Call zlDataBase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    Me.MousePointer = vbDefault
    MsgBox "初始药品结存成功！", vbOKOnly + vbInformation, gstrSysName
    
    Call IniControl
    Call Get上次结存(mlng库房ID)
    
    DoEvents
    Me.staThis.Panels(2).Text = ""
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdInput_Click()
    If MsgBox("录入差价必须所有库房都完成初始结存，是否现在就录入差价？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    End If
    
    frmDiffPriceRecal.ShowCard Me, 2
    
End Sub
Private Sub CmdSave_Click()
    Dim strFirstTime As String
    Dim strSecondTime As String
    
    On Error GoTo errHandle
    
    If mint核算方法 = type_先进先出 And mstr上次结存时间 = "" Then
        MsgBox "该库房还没有进行初始结存，请按初始结存按钮进行结存！", vbOKOnly + vbInformation, gstrSysName
        Exit Sub
    End If
    
    If mint核算方法 = type_先进先出 And mstr上次结存时间 <> "" And mbln是否审核 = False Then
        MsgBox "该库房上次结存还没有审核，请先审核上次结存！", vbOKOnly + vbInformation, gstrSysName
        Exit Sub
    End If
    
    If mint核算方法 = type_先进先出 And mstr上次结存时间 <> "" And mbln是否审核 = True Then
        If mstr上次结存时间 > Format(dtpTime.Value, "yyyy-mm-dd hh:mm:ss") Then
            MsgBox "当前的结存日期小于了上次结存日期，请重新设置结存日期，或者取消上次结存！", vbOKOnly + vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    If mint核算方法 = type_先进先出 Then
        strFirstTime = Format(dtpTime.Value, "yyyy-mm-dd hh:mm:ss")
        strSecondTime = mstr上次结存时间
    End If
    
    If mint核算方法 = type_全月平均 Then
        strFirstTime = Format(dtpBegin.Value, "yyyy-mm-dd hh:mm:ss")
        strSecondTime = Format(dtpEnd.Value, "yyyy-mm-dd hh:mm:ss")
    End If
    
    gstrSQL = "zl_药品差价重整_UPDATE(to_date('" & Format(strFirstTime, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss') ," & mlng库房ID & ","

    If mstr上次结存时间 = "" Then
        gstrSQL = gstrSQL & "NULL,"
    Else
        gstrSQL = gstrSQL & "to_date('" & Format(strSecondTime, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),"
    End If

    gstrSQL = gstrSQL & mint核算方法 & ")"
      
    Me.staThis.Panels(2).Text = "正在计算差价，请等待。。。！"
    
    Me.MousePointer = vbHourglass
    
    Call zlDataBase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    Me.MousePointer = vbDefault
    MsgBox "差价重算成功！", vbOKOnly + vbInformation, gstrSysName
    
    Call RefreshNow(mlng库房ID)
    
    DoEvents
    Me.staThis.Panels(2).Text = ""
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdVerify_Click()
    gstrSQL = "Zl_药品结存_Verify(to_date('" & Format(mstr上次结存时间, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')," & mlng库房ID & ")"
    Me.staThis.Panels(2).Text = "正在审核上次结存，请等待。。。！"
    
    Me.MousePointer = vbHourglass
    
    Call zlDataBase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    Me.MousePointer = vbDefault
    MsgBox "审核成功！", vbOKOnly + vbInformation, gstrSysName

    Call IniControl
    Call Get上次结存(mlng库房ID)
    
    DoEvents
    Me.staThis.Panels(2).Text = ""
    
    Exit Sub
End Sub

Private Sub dtpBegin_Change()
    If DateDiff("s", dtpEnd.Value, dtpBegin.Value) > 0 Then
        dtpBegin.Value = Format(Sys.Currentdate, "yyyy-mm") & "-01 00:00:00"
    End If
End Sub


Private Sub dtpEnd_Change()
    If DateDiff("s", dtpEnd.Value, dtpBegin.Value) > 0 Then
        dtpEnd.Value = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
    End If
End Sub

Private Sub Form_Load()
    Dim rsTmp As New Recordset
    
    On Error GoTo errHandle
    RestoreWinState Me, App.Title
    
    '载入库房
    gstrSQL = "Select Distinct A.ID, A.名称 " & _
            " From 部门性质说明 C, 部门性质分类 B, 部门表 A " & _
            " Where (a.站点 = [1] Or a.站点 is Null) And C.工作性质 = B.名称 And A.ID = C.部门id And " & _
            " To_Char(A.撤档时间, 'yyyy-MM-dd') = '3000-01-01' And " & _
            " Instr('HIJKLMN', B.编码, 1) > 0"
    Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "-载入所有库房", gstrNodeNo)
    
    With Cbo库房
        .Clear
        Do While Not rsTmp.EOF
            .AddItem rsTmp!名称
            .ItemData(.NewIndex) = rsTmp!id
            rsTmp.MoveNext
        Loop
        rsTmp.Close
        'If .ListIndex = -1 Then .ListIndex = 0
        If .ListCount < 1 Then
            MsgBox "至少应该设置一个有药库性质，药房性质，或者制剂室性质的部门，请查看部门管理！", vbInformation, gstrSysName
            Unload Me
            Exit Sub
        Else
            .ListIndex = 0
        End If
    End With
    
    '载入核算方法
    With cbo核算方法
        .Clear
        .AddItem "全月平均"
        .AddItem "先进先出"
        If .ListIndex = -1 Then .ListIndex = 0
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.Title
End Sub

