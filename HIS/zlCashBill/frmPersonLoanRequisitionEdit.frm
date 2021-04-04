VERSION 5.00
Begin VB.Form frmPersonLoanRequisitionEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "借款申请"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7590
   Icon            =   "frmPersonLoanRequisitionEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   7590
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   420
      Left            =   6225
      TabIndex        =   18
      Top             =   4365
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   420
      Left            =   4875
      TabIndex        =   17
      Top             =   4365
      Width           =   1200
   End
   Begin VB.Frame fraSplit 
      Height          =   120
      Index           =   1
      Left            =   -45
      TabIndex        =   21
      Top             =   3990
      Width           =   7830
   End
   Begin VB.Frame fraSplit 
      Height          =   120
      Index           =   0
      Left            =   30
      TabIndex        =   20
      Top             =   900
      Width           =   7830
   End
   Begin VB.Frame fra申请 
      BorderStyle     =   0  'None
      Height          =   2940
      Left            =   105
      TabIndex        =   19
      Top             =   1035
      Width           =   7530
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   6
         Left            =   870
         TabIndex        =   4
         Top             =   810
         Width           =   2490
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   5
         Left            =   870
         TabIndex        =   16
         Top             =   2430
         Width           =   6375
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   4
         Left            =   870
         TabIndex        =   14
         Top             =   2025
         Width           =   2490
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   3
         Left            =   870
         TabIndex        =   8
         Top             =   1230
         Width           =   6330
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   2
         Left            =   4695
         TabIndex        =   12
         Top             =   1620
         Width           =   2490
      End
      Begin VB.ComboBox cbo借出人 
         Height          =   300
         Left            =   870
         TabIndex        =   10
         Text            =   "cbo借出人"
         Top             =   1605
         Width           =   2490
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   1
         Left            =   4695
         TabIndex        =   6
         Top             =   765
         Width           =   2490
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   870
         TabIndex        =   2
         Top             =   390
         Width           =   2490
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "借款金额"
         Height          =   180
         Index           =   7
         Left            =   75
         TabIndex        =   3
         Top             =   870
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "取消原因"
         Height          =   180
         Index           =   6
         Left            =   75
         TabIndex        =   15
         Top             =   2520
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "取消时间"
         Height          =   180
         Index           =   5
         Left            =   75
         TabIndex        =   13
         Top             =   2085
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "备注"
         Height          =   180
         Index           =   4
         Left            =   435
         TabIndex        =   7
         Top             =   1260
         Width           =   360
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "借出时间"
         Height          =   180
         Index           =   3
         Left            =   3855
         TabIndex        =   11
         Top             =   1680
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "借出人"
         Height          =   240
         Index           =   2
         Left            =   255
         TabIndex        =   9
         Top             =   1650
         Width           =   540
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "申请时间"
         Height          =   180
         Index           =   1
         Left            =   3855
         TabIndex        =   5
         Top             =   825
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "借款人"
         Height          =   180
         Index           =   0
         Left            =   255
         TabIndex        =   1
         Top             =   450
         Width           =   540
      End
   End
   Begin VB.Label lbl 
      Caption         =   $"frmPersonLoanRequisitionEdit.frx":058A
      Height          =   885
      Left            =   1035
      TabIndex        =   0
      Top             =   135
      Width           =   6435
   End
   Begin VB.Image img 
      Height          =   720
      Left            =   135
      Picture         =   "frmPersonLoanRequisitionEdit.frx":0675
      Top             =   150
      Width           =   720
   End
End
Attribute VB_Name = "frmPersonLoanRequisitionEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnFirst As Boolean, mblnChange As Boolean, mlngID As Long
Private mlngModule As Long, mstrPrivs As String
Private mblnSucceed As Boolean '操作标志:true,成功,否则False
Private mrs人员 As ADODB.Recordset

Public Enum gEditTypeLoan
    FN_申请 = 0
    FN_修改 = 1
    FN_借出 = 2
    FN_取消借出 = 3
    FN_查询 = 4
End Enum

Private mEditType As gEditTypeLoan
Private Enum mIdxTxt
    idx_借款人 = 0
    idx_申请时间 = 1
    idx_借出时间 = 2
    idx_备注 = 3
    idx_取消时间 = 4
    idx_取消原因 = 5
    idx_借款金额 = 6
End Enum
Public Function ShowEdit(ByVal frmMain As Form, ByVal EditType As gEditTypeLoan, ByVal strPrivs As String, ByVal lngModule As Long, Optional lngID As Long = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:程序编辑入口:表示显示或申请等操作
    '入参:frmMain-主窗体
    '     EditType-编辑类型
    '     strPrivs-权限串
    '     lngModule－模块号
    '     lngID-非申请时，有效,传入借款ID
    '出参:
    '返回:操作成功，返回ture,否则返回False
    '编制:刘兴洪
    '日期:2009-09-08 11:54:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule: mstrPrivs = strPrivs: mEditType = EditType: mblnSucceed = False: mlngID = lngID
    Me.Show 1, frmMain
    ShowEdit = mblnSucceed
End Function

Private Function CheckDepend() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查相关的依赖关系
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-09-08 12:00:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    gstrSQL = "" & _
    "   Select Distinct B.ID, B.编号, B.姓名, B.别名, B.简码, B.出生日期, B.性别, B.办公室电话 " & _
    "   From 人员性质说明 A, 人员表 B " & _
    "   Where A.人员id = B.ID And A.人员性质 In ('门诊挂号员', '门诊收费员', '预交收款员', '住院结帐员') " & _
    "   Order By 编号"
    Set mrs人员 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If mrs人员.RecordCount = 0 Then
        MsgBox "注意：" & vbCrLf & _
               "　　没有一个人员为“门诊挂号员、门诊收费员、预交收款员、住院结帐员”" & vbCrLf & _
               " 请在“人员管理”中设置！", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If mEditType = FN_申请 Or mEditType = FN_修改 Then
        With cbo借出人
            .Clear: mrs人员.MoveFirst
            Do While Not mrs人员.EOF
                If Nvl(mrs人员!姓名) <> UserInfo.姓名 Then
                    .AddItem Nvl(mrs人员!姓名)
                    .ItemData(.NewIndex) = Nvl(Val(mrs人员!ID))
                End If
                mrs人员.MoveNext
            Loop
        End With
    End If
    CheckDepend = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub SetCtrolEnabled()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置控件的编辑属性
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-09-08 14:51:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    txtEdit(mIdxTxt.idx_借款人).Enabled = False
    txtEdit(mIdxTxt.idx_申请时间).Enabled = False
    txtEdit(mIdxTxt.idx_借出时间).Enabled = False
    txtEdit(mIdxTxt.idx_备注).Enabled = False
    txtEdit(mIdxTxt.idx_取消时间).Enabled = False
    txtEdit(mIdxTxt.idx_取消原因).Enabled = False
    txtEdit(mIdxTxt.idx_借款金额).Enabled = False
    cbo借出人.Enabled = False
    Select Case mEditType
    Case gEditTypeLoan.FN_申请, gEditTypeLoan.FN_修改
        txtEdit(mIdxTxt.idx_备注).Enabled = True
        txtEdit(mIdxTxt.idx_借款金额).Enabled = True
        cbo借出人.Enabled = True
    Case gEditTypeLoan.FN_取消借出
        txtEdit(mIdxTxt.idx_取消原因).Enabled = True
    Case Else
    End Select
End Sub
Private Sub ClearCtrl()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除数据
    '编制:刘兴洪
    '日期:2009-09-08 14:27:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim ctl As Control
    For Each ctl In Me.Controls
        Select Case UCase(TypeName(ctl))
        Case "TEXTBOX"
            ctl.Text = ""
            If mEditType = FN_申请 Then
                If ctl.Index = mIdxTxt.idx_借款人 Then ctl.Text = UserInfo.姓名
                If ctl.Index = mIdxTxt.idx_申请时间 Then ctl.Text = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
            End If
        Case Else
        End Select
        Call zlSetCrlEnbled(ctl, ctl.Enabled)
    Next
End Sub

Private Function LoadDataToCtrol() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载数据到控件
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-09-08 14:26:36
    '---------------------------------------------------------------------------------------------------------------------------------------------\
    Dim rsTemp As ADODB.Recordset, i As Long
    
    Err = 0: On Error GoTo ErrHand:
    Call ClearCtrl
    If mEditType = FN_申请 Then LoadDataToCtrol = True: Exit Function
    
  
    gstrSQL = " " & _
    "    Select Id, 借款金额, 备注, 借款人, to_char(申请时间,'yyyy-mm-dd hh24:mi:ss') as 申请时间 ,  " & _
    "           借出人, to_char(借出时间,'yyyy-mm-dd hh24:mi:ss') as 借出时间, " & _
    "           to_char(取消时间,'yyyy-mm-dd hh24:mi:ss') as 取消时间, 取消原因 " & _
    "    From 人员借款记录 " & _
    "    Where ID=[1] " & _
    "    Order by 借出人,申请时间"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngID)
    If rsTemp.RecordCount = 0 Then
        MsgBox "注意：" & vbCrLf & _
               "    该借款记录可能已经被他人删除，不能继续操作!", vbOKOnly + vbDefaultButton1 + vbInformation, gstrSysName
        Exit Function
    End If
    
    txtEdit(mIdxTxt.idx_借款人).Text = Nvl(rsTemp!借款人)
    txtEdit(mIdxTxt.idx_申请时间).Text = Nvl(rsTemp!申请时间)
    txtEdit(mIdxTxt.idx_备注).Text = Nvl(rsTemp!备注)
    txtEdit(mIdxTxt.idx_借款金额).Text = Format(Val(Nvl(rsTemp!借款金额)), "####0.00;-###0.00;;")
    
    If mEditType = FN_修改 Then
        If Nvl(rsTemp!借出时间) <> "" Then
            MsgBox "注意：" & vbCrLf & _
                   "    该借款记录已经被他人借出，不能再进行修改操作!", vbOKOnly + vbDefaultButton1 + vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    If mEditType = FN_借出 Then
        If Nvl(rsTemp!借出时间) <> "" Then
            MsgBox "注意：" & vbCrLf & _
                   "    该借款记录已经被他人借出，不能再进行借出操作!", vbOKOnly + vbDefaultButton1 + vbInformation, gstrSysName
            Exit Function
        End If
        txtEdit(mIdxTxt.idx_借出时间).Text = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    Else
        txtEdit(mIdxTxt.idx_借出时间).Text = Nvl(rsTemp!借出时间)
    End If
    
    If mEditType = FN_取消借出 Then
        If Trim(Nvl(rsTemp!借出时间)) = "" Then
            MsgBox "注意：" & vbCrLf & _
                   "    该借款记录还未确认借出，不能进行取消借出操作!", vbOKOnly + vbDefaultButton1 + vbInformation, gstrSysName
            Exit Function
        End If
        If Trim(Nvl(rsTemp!取消时间)) <> "" Then
            MsgBox "注意：" & vbCrLf & _
                   "    该借款记录已经被他人取消，不能再进行取消借出操作!", vbOKOnly + vbDefaultButton1 + vbInformation, gstrSysName
            Exit Function
        End If
        txtEdit(mIdxTxt.idx_取消时间).Text = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    Else
        txtEdit(mIdxTxt.idx_取消时间).Text = Nvl(rsTemp!取消时间)
    End If
    With cbo借出人
        .ListIndex = -1
        For i = 0 To .ListCount - 1
            If .List(i) = Nvl(rsTemp!借出人) Then .ListIndex = i: Exit For
        Next
        If .ListIndex < 0 Then
            .AddItem Nvl(rsTemp!借出人)
            .ListIndex = .NewIndex
        End If
    End With
    
    txtEdit(mIdxTxt.idx_取消原因).Text = Nvl(rsTemp!取消原因)
    LoadDataToCtrol = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Private Sub SetDefaultInputLen()
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    gstrSQL = "Select 备注,取消原因  From 人员借款记录 where id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, -1)
    txtEdit(mIdxTxt.idx_备注).MaxLength = rsTemp.Fields("备注").DefinedSize
    txtEdit(mIdxTxt.idx_取消原因).MaxLength = rsTemp.Fields("取消原因").DefinedSize
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Function isValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查输入的数据是否合法
    '返回:合法，返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-09-08 15:25:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim ctl As Object
    
    Err = 0: On Error GoTo ErrHand:
    
    Set ctl = txtEdit(mIdxTxt.idx_备注)
    If zlCommFun.ActualLen(ctl.Text) > ctl.MaxLength Then
        MsgBox "备注只能输入" & ctl.MaxLength & " 个字符或" & ctl.MaxLength \ 2 & "个汉字,请检查!", vbInformation + vbOKOnly, gstrSysName
        Call zlcontrol.ControlSetFocus(ctl, True)
        Exit Function
    End If
    Set ctl = txtEdit(mIdxTxt.idx_取消原因)
    If zlCommFun.ActualLen(ctl.Text) > ctl.MaxLength Then
        MsgBox "取消原因只能输入" & ctl.MaxLength & " 个字符或" & ctl.MaxLength \ 2 & "个汉字,请检查!", vbInformation + vbOKOnly, gstrSysName
        Call zlcontrol.ControlSetFocus(ctl, True)
        Exit Function
    End If
    
    Set ctl = txtEdit(mIdxTxt.idx_借款金额)
    If Val(ctl.Text) > 10 ^ 12 - 1 Then
        MsgBox "借款金额必须小于" & CStr(10 ^ 12 - 1) & " ,请检查!", vbInformation + vbOKOnly, gstrSysName
        Call zlcontrol.ControlSetFocus(ctl, True)
        Exit Function
    End If
    If Val(ctl.Text) <= 0 Then
        MsgBox "借款金额必须大于零 ,请检查!", vbInformation + vbOKOnly, gstrSysName
        Call zlcontrol.ControlSetFocus(ctl, True)
        Exit Function
    End If
        
    Set ctl = txtEdit(mIdxTxt.idx_借款人)
    If ctl.Text = cbo借出人.Text Then
        MsgBox "借款人与借出人是同一人,不能继续!", vbInformation + vbOKOnly, gstrSysName
        Call zlcontrol.ControlSetFocus(cbo借出人, True)
        Exit Function
    End If
    If Trim(ctl.Text) = "" Then
        MsgBox "未设置人员的对照关系,请与系统管理员联系!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If Trim(cbo借出人.Text) = "" Or cbo借出人.ListIndex < 0 Then
        MsgBox "借出人未选择,不能继续!", vbInformation + vbOKOnly, gstrSysName
        Call zlcontrol.ControlSetFocus(cbo借出人, True)
        Exit Function
    End If
    
    isValied = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function SaveData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存数据
    '返回:保存成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-09-08 15:57:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long
    Err = 0: On Error GoTo ErrHand:
    If mEditType = FN_修改 Then
        lngID = mlngID
    Else
        lngID = zlDatabase.GetNextId("人员借款记录")
    End If
    'Zl_人员借款记录_Insert
    gstrSQL = IIf(mEditType = FN_修改, "Zl_人员借款记录_Update(", "Zl_人员借款记录_Insert(")
    '  Id_In       In 人员借款记录.ID%Type,
    gstrSQL = gstrSQL & "" & lngID & ","
    '  借款金额_In In 人员借款记录.借款金额%Type,
    gstrSQL = gstrSQL & "" & Val(txtEdit(mIdxTxt.idx_借款金额)) & ","
    '  备注_In     In 人员借款记录.备注%Type,
    gstrSQL = gstrSQL & "'" & Trim(txtEdit(mIdxTxt.idx_备注)) & "',"
    '  借款人_In   In 人员借款记录.借款人%Type,
    gstrSQL = gstrSQL & "'" & Trim(txtEdit(mIdxTxt.idx_借款人)) & "',"
    '  申请时间_In In 人员借款记录.申请时间%Type,
    gstrSQL = gstrSQL & "to_date('" & Trim(txtEdit(mIdxTxt.idx_申请时间)) & "','yyyy-mm-dd hh24:mi:ss'),"
    '  借出人_In   In 人员借款记录.借出人%Type
    gstrSQL = gstrSQL & "'" & Trim(cbo借出人.Text) & "')"
    
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    mlngID = lngID
    SaveData = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Private Function SaveLoanOut(ByVal lngID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:借出保存
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-09-08 16:28:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    'Zl_人员借款记录_借出(Id_In In 人员借款记录.ID%Type) Is
    Err = 0: On Error GoTo ErrHand:
    gstrSQL = "Zl_人员借款记录_借出(" & lngID & ")"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    SaveLoanOut = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Private Function SaveCancelLoanOut() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:借出保存
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-09-08 16:28:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    ' Zl_人员借款记录_取消借出
    gstrSQL = "Zl_人员借款记录_取消借出("
    '  Id_In       In 人员借款记录.ID%Type,
    gstrSQL = gstrSQL & "" & mlngID & ","
    '  取消原因_In In 人员借款记录.取消原因%Type,
    gstrSQL = gstrSQL & "'" & txtEdit(mIdxTxt.idx_取消原因).Text & "',"
    '  取消时间_In In 人员借款记录.取消时间%Type
    gstrSQL = gstrSQL & "to_date('" & txtEdit(mIdxTxt.idx_取消时间).Text & "','yyyy-mm-dd hh24:mi:ss'))"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    SaveCancelLoanOut = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Private Sub cbo借出人_GotFocus()
    zlcontrol.TxtSelAll cbo借出人
    zlCommFun.OpenIme False
End Sub

Private Sub cbo借出人_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If cbo借出人.ListIndex >= 0 Then zlCommFun.PressKey vbKeyTab: Exit Sub
    '选择:
    If Select人员选择器(cbo借出人, Trim(cbo借出人.Text)) Then Exit Sub
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If mEditType = FN_查询 Then Unload Me: Exit Sub
    
    If isValied = False Then Exit Sub
    
    If mEditType = FN_借出 Then
        If SaveLoanOut(mlngID) = False Then Exit Sub
        If IIf(Val(zlDatabase.GetPara("借出打印", glngSys, mlngModule, "0")) = 1, 1, 0) = 1 Then
            '打印
            If InStr(mstrPrivs, "借款单") <> 0 Then
                ReportOpen gcnOracle, glngSys, "zl1_bill_1502", Me, "ID=" & mlngID, 2
            End If
        End If
        mblnChange = False: mblnSucceed = True: Unload Me: Exit Sub
    End If
    If mEditType = FN_取消借出 Then
        If SaveCancelLoanOut = False Then Exit Sub
        If IIf(Val(zlDatabase.GetPara("借出打印", glngSys, mlngModule, "0")) = 1, 1, 0) = 1 Then
            '打印
            If InStr(mstrPrivs, "借款单") <> 0 Then
                ReportOpen gcnOracle, glngSys, "zl1_bill_1502", Me, "ID=" & mlngID, 2
            End If
        End If
        mblnChange = False: mblnSucceed = True: Unload Me: Exit Sub
    End If
    
    If SaveData = False Then Exit Sub
    
    If IIf(Val(zlDatabase.GetPara("申请打印", glngSys, mlngModule, "0")) = 1, 1, 0) = 1 Then
        '打印
        If InStr(mstrPrivs, "借款单") <> 0 Then
            ReportOpen gcnOracle, glngSys, "zl1_bill_1502", Me, "ID=" & mlngID, 2
        End If
    End If
    
    mblnSucceed = True: mblnChange = False
    If mEditType = FN_修改 Then Unload Me: Exit Sub
    Call ClearCtrl
    zlcontrol.ControlSetFocus txtEdit(mIdxTxt.idx_借款金额), True
    mlngID = 0
    mblnSucceed = True: mblnChange = False
End Sub


Private Sub cbo借出人_Change()
    mblnChange = True
    
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If CheckDepend = False Then Unload Me: Exit Sub
    Call SetDefaultInputLen
    mblnChange = False
    Call SetCtrolEnabled
    '加载数据
    If LoadDataToCtrol = False Then
        mblnChange = False
        Unload Me: Exit Sub
    End If
    If mEditType = FN_修改 Or mEditType = FN_申请 Then
        zlcontrol.ControlSetFocus txtEdit(mIdxTxt.idx_借款金额), True
        cbo借出人.SelLength = 0
    ElseIf mEditType = FN_取消借出 Then
        zlcontrol.ControlSetFocus txtEdit(mIdxTxt.idx_取消原因), True
    Else
        zlcontrol.ControlSetFocus cmdOK, True
    End If
    mblnChange = False
End Sub
Private Sub Form_Load()
    mblnFirst = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange Then
        If MsgBox("数据可能已改变，但未存盘，真要退出吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1
            Exit Sub
        End If
    End If
End Sub

Private Sub txtEdit_Change(Index As Integer)
    If Index <> mIdxTxt.idx_借款金额 Then
        mblnChange = True
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlcontrol.TxtSelAll txtEdit(Index)
    Select Case Index
    Case mIdxTxt.idx_备注, mIdxTxt.idx_取消原因
        zlCommFun.OpenIme True
    Case Else
        zlCommFun.OpenIme False
    End Select
End Sub

Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab: Exit Sub
    mblnChange = True
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = mIdxTxt.idx_借款金额 Then
        zlcontrol.TxtCheckKeyPress txtEdit(Index), KeyAscii, m金额式
    Else
        zlcontrol.TxtCheckKeyPress txtEdit(Index), KeyAscii, m文本式
    End If
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
     zlCommFun.OpenIme False
     If Index = mIdxTxt.idx_借款金额 Then
        txtEdit(Index).Text = Format(Val(txtEdit(Index).Text), "###0.00;-###0.00;;")
     End If
End Sub

Private Function Select人员选择器(ByVal objCtl As Control, ByVal strSearch As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:部门选择器
    '入参::objCtl-指定控件
    '     strSearch-要搜索的条件
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-05-01 14:18:58
    '-----------------------------------------------------------------------------------------------------------
    Dim blnCancel As Boolean, strKey As String, strTittle As String, lngH As Long, strTemp As String
    Dim vRect As RECT
    Dim rsTemp  As ADODB.Recordset
    Dim i As Long
    'zlDatabase.ShowSelect
    '功能：
    '参数：
    '     frmParent=显示的父窗体
    '     strSQL=数据来源,不同风格的选择器对SQL中的字段有不同要求
    '     bytStyle=选择器风格
    '       为0时:列表风格:ID,…
    '       为1时:树形风格:ID,上级ID,编码,名称(如果bln末级，则需要末级字段)
    '       为2时:双表风格:ID,上级ID,编码,名称,末级…；ListView只显示末级=1的项目
    '     strTitle=选择器功能命名,也用于个性化区分
    '     bln末级=当树形选择器(bytStyle=1)时,是否只能选择末级为1的项目
    '     strSeek=当bytStyle<>2时有效,缺省定位的项目。
    '             bytStyle=0时,以ID和上级ID之后的第一个字段为准。
    '             bytStyle=1时,可以是编码或名称
    '     strNote=选择器的说明文字
    '     blnShowSub=当选择一个非根结点时,是否显示所有下级子树中的项目(项目多时较慢)
    '     blnShowRoot=当选择根结点时,是否显示所有项目(项目多时较慢)
    '     blnNoneWin,X,Y,txtH=处理成非窗体风格,X,Y,txtH表示调用界面输入框的坐标(相对于屏幕)和高度
    '     Cancel=返回参数,表示是否取消,主要用于blnNoneWin=True时
    '     blnMultiOne=当bytStyle=0时,是否将对多行相同记录当作一行判断
    '     blnSearch=是否显示行号,并可以输入行号定位
    '返回：取消=Nothing,选择=SQL源的单行记录集
    '说明：
    '     1.ID和上级ID可以为字符型数据
    '     2.末级等字段不要带空值
    '应用：可用于各个程序中数据量不是很大的选择器,输入匹配列表等。
    
    strTittle = "人员选择器"
    vRect = zlcontrol.GetControlRect(objCtl.hwnd)
    lngH = objCtl.Height
    
  
    gstrSQL = "" & _
    "   Select Distinct B.ID, B.编号, B.姓名, B.别名, B.简码, B.出生日期, B.性别, B.办公室电话 " & _
    "   From 人员性质说明 A, 人员表 B " & _
    "   Where A.人员id = B.ID And A.人员性质 In ('门诊挂号员', '门诊收费员', '预交收款员', '住院结帐员') and B.id <>[2] " & _
    "         and (b.编号 like upper([1]) or b.姓名 like [1] or b.简码 like upper([1]) or b.别名 like [1]) " & _
    "   Order By b.编号"
    
    strKey = GetMatchingSting(strSearch, False)
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, strTittle, False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False, strKey, UserInfo.ID)
 
    If blnCancel = True Then
       zlcontrol.ControlSetFocus objCtl, True
        Exit Function
    End If
    
    If rsTemp Is Nothing Then
        ShowMsgbox "没有满足条件的人员信息,请检查!"
       zlcontrol.ControlSetFocus objCtl, True
        Exit Function
    End If
    zlcontrol.ControlSetFocus objCtl, True
    Dim blnHaveData As Boolean
    With objCtl
        For i = 0 To .ListCount - 1
            If Nvl(rsTemp!姓名) = .List(i) Then
                .ListIndex = i: Exit For
                blnHaveData = True
            End If
        Next
        If blnHaveData = False Then
            .AddItem Nvl(rsTemp!姓名)
            .ItemData(.NewIndex) = Val(Nvl(rsTemp!ID)): .ListIndex = .NewIndex
        End If
    End With
    zlCommFun.PressKey vbKeyTab
    Select人员选择器 = True
End Function

