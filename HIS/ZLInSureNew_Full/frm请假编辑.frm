VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm请假编辑 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "请假编辑"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6555
   Icon            =   "frm请假编辑.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSComCtl2.DTPicker dtp开始日期 
      Height          =   300
      Left            =   2820
      TabIndex        =   5
      Top             =   3210
      Visible         =   0   'False
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy年MM月dd日"
      Format          =   90046467
      CurrentDate     =   38433
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5730
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm请假编辑.frx":000C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "退出(&X)"
      Height          =   350
      Left            =   5070
      TabIndex        =   4
      Top             =   3210
      Width           =   1100
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "请假结束(&E)"
      Enabled         =   0   'False
      Height          =   405
      Left            =   150
      TabIndex        =   3
      Top             =   3150
      Width           =   1245
   End
   Begin VB.CommandButton cmdADD 
      Caption         =   "请假申请(&A)"
      Height          =   405
      Left            =   1500
      TabIndex        =   2
      Top             =   3150
      Width           =   1245
   End
   Begin MSComctlLib.ListView lvw请假记录 
      Height          =   2325
      Left            =   150
      TabIndex        =   1
      Top             =   720
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   4101
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "请假交易流水号"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "开始日期"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "结束日期"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   180
      Picture         =   "frm请假编辑.frx":1D16
      Top             =   150
      Width           =   480
   End
   Begin VB.Label lblPatient 
      Caption         =   "姓名:性别:医保号:入院日期"
      Height          =   195
      Left            =   780
      TabIndex        =   0
      Top             =   270
      Width           =   5955
   End
End
Attribute VB_Name = "frm请假编辑"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim blnStart As Boolean
Private mstrInput As String
Private mlng病人ID As Long
Private mstr流水号 As String
Private rsTemp As New ADODB.Recordset
Private cn银海 As New ADODB.Connection

Public Sub ShowEditor(ByVal lng病人ID As Long)
    mlng病人ID = lng病人ID
    Me.Show 1
End Sub

Private Sub cmdADD_Click()
    Dim blnEnabled As Boolean
    cmdEdit.Enabled = False
    
    With lvw请假记录
        If .ListItems.Count <> 0 Then
            If Not .SelectedItem Is Nothing Then
                blnEnabled = (.SelectedItem.SubItems(2) = "")
            End If
        End If
    End With
    
    If dtp开始日期.Visible Then
        dtp开始日期.Visible = False
        cmdEdit.Enabled = blnEnabled
    Else
        dtp开始日期.Visible = True
    End If
End Sub

Private Sub cmdEdit_Click()
    Dim blnTrans As Boolean
    Dim str开始日期 As String
    Dim datCurr As Date
    Dim str交易流水号 As String
    On Error GoTo errHand
    
    str交易流水号 = lvw请假记录.SelectedItem.Text
    str开始日期 = lvw请假记录.SelectedItem.SubItems(1)
    
    cn银海.BeginTrans
    blnTrans = True
    
    '产生新的请假登记记录，并调用接口
    gstrSQL = "zl_请假登记记录_END('" & mstr流水号 & "','" & str交易流水号 & "')"
    cn银海.Execute gstrSQL, , adCmdStoredProc
    
    mstrInput = mstr流水号 & "|" & str交易流水号 & "|" & str开始日期 & "|" & Format(datCurr, "yyyyMMdd")
    Call 调用接口_准备_重庆银海版("33", mstrInput)
    If Not 调用接口_重庆银海版 Then
        cn银海.RollbackTrans
        Exit Sub
    End If
    
    cn银海.CommitTrans
    blnTrans = False
    
    Call RefreshData
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTrans Then cn银海.RollbackTrans
End Sub

Private Sub cmdExit_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub dtp开始日期_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim datCurr As Date
    Dim blnTrans As Boolean
    Dim str开始日期 As String
    Dim str交易流水号 As String
    On Error GoTo errHand
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If MsgBox("你确定要进行请假开始登记吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    '检查日期
    If Format(dtp开始日期.Value, "yyyy-MM-dd") > Format(zlDatabase.Currentdate, "yyyy-MM-dd") Then
        MsgBox "请假开始日期不能大于当前日期！", vbInformation, gstrSysName
        Exit Sub
    End If
    str开始日期 = Format(dtp开始日期.Value, "yyyy-MM-dd")
    
    gstrSQL = "Select 1 From 请假登记记录 Where 结束日期 Is NULL"
    Call OpenRecordset(rsTemp, "检查是否存在请假未结束的记录", gstrSQL, cn银海)
    If rsTemp.RecordCount <> 0 Then
        MsgBox "当前病人还处于请假状态中,不能继续进行请假登记！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '请假开始,检查是否存在请假未结束的记录
    cn银海.BeginTrans
    blnTrans = True
    
    '产生新的请假登记记录，并调用接口
    datCurr = zlDatabase.Currentdate()
    gstrSQL = "zl_请假登记记录_START('" & mstr流水号 & "','" & Format(datCurr, "yyyyMMddHHmmss") & "',to_Date('" & str开始日期 & "','yyyy-MM-dd'))"
    cn银海.Execute gstrSQL, , adCmdStoredProc
    
    mstrInput = mstr流水号 & "|" & Format(datCurr, "yyyyMMddHHmmss") & "|" & Format(str开始日期, "yyyyMMdd") & "|"
    Call 调用接口_准备_重庆银海版("33", mstrInput)
    If Not 调用接口_重庆银海版 Then
        cn银海.RollbackTrans
        Exit Sub
    End If
    
    cn银海.CommitTrans
    blnTrans = False
    
    Call cmdADD_Click
    Call RefreshData
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTrans Then cn银海.RollbackTrans
End Sub

Private Sub Form_Activate()
    If blnStart = False Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Dim strPatient As String
    Dim strUser As String, strServer As String, strPass As String
    Dim rsTmp As ADODB.Recordset
    blnStart = False
    
    If Not Init医保 Then Exit Sub
    
    '取病人的流水号
    gstrSQL = "Select 版本号 From zlSystems Where 编号 = 100"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "HIS版本号")
    If Split(rsTmp!版本号, ".")(0) = 10 And Split(rsTmp!版本号, ".")(1) >= 34 Then
        gstrSQL = " Select A.流水号,B.姓名,B.性别,A.医保号,C.入院日期 " & _
              " From 保险帐户 A,病人信息 B,病案主页 C" & _
              " Where A.险类=" & TYPE_重庆银海版 & " And A.病人ID=" & mlng病人ID & _
              " And A.病人ID=B.病人ID And B.病人ID=C.病人ID And B.主页ID=C.主页ID"
    Else
        gstrSQL = " Select A.流水号,B.姓名,B.性别,A.医保号,C.入院日期 " & _
              " From 保险帐户 A,病人信息 B,病案主页 C" & _
              " Where A.险类=" & TYPE_重庆银海版 & " And A.病人ID=" & mlng病人ID & _
              " And A.病人ID=B.病人ID And B.病人ID=C.病人ID And B.住院次数=C.主页ID"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取病人的就诊流水号")
    mstr流水号 = Nvl(rsTemp!流水号)
    strPatient = "姓名:" & rsTemp!姓名 & "|" & "性别:" & Nvl(rsTemp!性别) & "|" & "医保号:" & rsTemp!医保号 & "|" & "入院日期:" & Format(rsTemp!入院日期, "yyyy-MM-dd")
    lblPatient.Caption = strPatient
    Me.dtp开始日期.Value = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    
    Call RefreshData
    
    blnStart = True
End Sub

Private Sub RefreshData()
    Dim lvwItem As ListItem
    '提取出该病人本次住院的所有请假记录
    gstrSQL = "Select 请假交易流水号,开始日期,结束日期 From 请假登记记录 Where 就诊流水号='" & mstr流水号 & "' Order By 请假交易流水号"
    Call OpenRecordset(rsTemp, "提取出该病人本次住院的所有请假记录", gstrSQL, cn银海)
    With rsTemp
        lvw请假记录.ListItems.Clear
        Do While Not .EOF
            Set lvwItem = lvw请假记录.ListItems.Add(, "K_" & .AbsolutePosition, !请假交易流水号, , 1)
            lvwItem.SubItems(1) = Format(!开始日期, "yyyyMMdd")
            If Not IsNull(!结束日期) Then
                lvwItem.SubItems(2) = Format(!结束日期, "yyyyMMdd")
            End If
            .MoveNext
        Loop
        
        If .RecordCount <> 0 Then
            lvw请假记录.ListItems(1).Selected = True
            lvw请假记录.SelectedItem.Selected = True
            Call lvw请假记录_ItemClick(lvw请假记录.SelectedItem)
        End If
    End With
End Sub

Private Sub lvw请假记录_ItemClick(ByVal Item As MSComctlLib.ListItem)
    '任何情况都允许新增
    '当前记录无结束时间时,允许结束
    cmdEdit.Enabled = False
    If dtp开始日期.Visible Then Exit Sub
    
    With lvw请假记录
        If .ListItems.Count = 0 Then Exit Sub
        If .SelectedItem Is Nothing Then Exit Sub
        
        cmdEdit.Enabled = (Item.SubItems(2) = "")
    End With
End Sub

Private Function Init医保() As Boolean
    Dim strUser As String, strServer As String, strPass As String
    
    '读出连接医保服务器的配置
    gstrSQL = "select 参数名,参数值 from 保险参数 where 参数名 like '医保%' and 险类=" & TYPE_重庆银海版
    Call OpenRecordset(rsTemp, "泸州医保")
    
    Do Until rsTemp.EOF
        Select Case rsTemp("参数名")
            Case "医保用户名"
                strUser = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
            Case "医保服务器"
                strServer = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
            Case "医保用户密码"
                strPass = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
        End Select
        rsTemp.MoveNext
    Loop
    
    If OraDataOpen(cn银海, strServer, strUser, strPass, False) = False Then
        MsgBox "无法连接到中间库，请检查保险参数是否设置正确！", vbInformation, gstrSysName
        Exit Function
    End If
    Init医保 = True
End Function
