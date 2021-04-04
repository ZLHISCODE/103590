VERSION 5.00
Begin VB.Form frmDepositNew 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "门诊预交转住院预交"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5760
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtPrePay 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   2190
      TabIndex        =   9
      Text            =   "0.00"
      Top             =   870
      Width           =   2355
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4230
      TabIndex        =   13
      Top             =   1770
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2940
      TabIndex        =   12
      Top             =   1770
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   330
      TabIndex        =   11
      Top             =   1770
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   45
      Left            =   30
      TabIndex        =   14
      Top             =   540
      Width           =   5775
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "住院号："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   210
      Index           =   4
      Left            =   3840
      TabIndex        =   6
      Top             =   180
      Width           =   840
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2145"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Index           =   14
      Left            =   4650
      TabIndex        =   7
      Top             =   180
      Width           =   420
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "12岁"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Index           =   13
      Left            =   3120
      TabIndex        =   5
      Top             =   180
      Width           =   420
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "年龄："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   210
      Index           =   3
      Left            =   2580
      TabIndex        =   4
      Top             =   180
      Width           =   630
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "男"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Index           =   12
      Left            =   2100
      TabIndex        =   3
      Top             =   180
      Width           =   210
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "性别："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   210
      Index           =   2
      Left            =   1500
      TabIndex        =   2
      Top             =   180
      Width           =   630
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "王二小"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Index           =   11
      Left            =   780
      TabIndex        =   1
      Top             =   180
      Width           =   630
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "病人："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   210
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   630
   End
   Begin VB.Label lblPrePay 
      AutoSize        =   -1  'True
      Caption         =   "金额(&T):"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1050
      TabIndex        =   8
      Top             =   900
      Width           =   1020
   End
   Begin VB.Label lblremark 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "说明:"
      Height          =   180
      Left            =   2580
      TabIndex        =   10
      Top             =   1320
      Width           =   465
   End
End
Attribute VB_Name = "frmDepositNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mdblPrePay As Double
Private mstrRemark As String
Private mblnOK As Boolean

Private Enum idx_Lable
    lblName = 1
    txtName = 11
    lblSex = 2
    txtSex = 12
    lblAge = 3
    txtAge = 13
    lblInNumber = 4
    txtInNumber = 14
End Enum

Private mpatiInfo As clsPatientInfo '病人信息

Public Function ShowMe(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    '功能:新门诊病人门诊预交转住院预交
    '入参:
    '   lng病人ID - 病人ID
    '   lng主页ID - 主页ID
    '返回:成功返回True,否则返回False
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    mblnOK = False
    
    On Error Resume Next
    Me.Show vbModal
    ShowMe = mblnOK

End Function

Private Sub Form_Load()
    Dim strData As String
    
    zlCommFun.ShowFlash "正在获取可转入的门诊预交款，请稍后...", Me
    If GetBillData(mlng病人ID, strData) = False Then GoTo ErrExit:
    If InitData(mlng病人ID, mlng主页ID) = False Then GoTo ErrExit:
    If InitFace() = False Then GoTo ErrExit:
    zlCommFun.StopFlash
    Exit Sub
ErrExit:
    zlCommFun.StopFlash
    Unload Me: Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    mlng病人ID = 0
    mlng主页ID = 0
    mdblPrePay = 0
    mstrRemark = ""

    Set mpatiInfo = Nothing

End Sub

Private Sub cmdOK_Click()
    Dim strJsonIn As String
    Dim strData As String
    Dim blnTrans As Boolean
    
    On Error GoTo ErrHander
    cmdOk.Enabled = False
    zlCommFun.ShowFlash "正在进行门诊预交金额转住院处理，请稍后...", Me
    
    If CheckPrePayValid = False Then cmdOk.Enabled = True: Exit Sub
    '保存预交数据
    gcnOracle.BeginTrans: blnTrans = True
    If SaveDate() = False Then gcnOracle.RollbackTrans: cmdOk.Enabled = True: Exit Sub
    
    '门诊预交金额转住院确认
    '输入    编码             名称      说明                数据类型        备注
    '        pid              患者ID                         Number(18)       非空
    '        prepaid_payment  预交金                        Number(18,2)    非空
    '
    '输出    编码             名称      说明                数据类型        备注
    '        result           执行结果  1-成功；-1-失败     Number(1)       非空
    '        errmsg           错误消息  失败时返回错误消息  Varchar2(200)
    strJsonIn = "{""head"":{""bizno"":""RJ005"",""sysno"":""ZLDAYROOM"",""time"":"""",""action_no"":"""",""tarno"":""03""}"
    strJsonIn = "{""input"":" & strJsonIn & ",""pid"":" & mlng病人ID & ",""prepaid_payment"":" & mdblPrePay & "}}"
    
    Call Sys.NewSystemSvr("新门诊系统", "门诊预交金额转住院确认", strJsonIn, strData)

    If strData = "" Then strData = "{}"
    If Val(zlStr.JSONParse("result", strData)) <> 1 Then
        gcnOracle.RollbackTrans
        MsgBox zlStr.JSONParse("errmsg", strData), vbInformation, gstrSysName
        zlCommFun.StopFlash: mblnOK = True
        Exit Sub
    End If
    gcnOracle.CommitTrans: blnTrans = False
    zlCommFun.StopFlash
    cmdOk.Enabled = True
    mblnOK = True
    Unload Me
    Exit Sub
ErrHander:
    zlCommFun.StopFlash
    cmdOk.Enabled = True
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Function InitFace() As Boolean
    '初始化界面
    On Error GoTo ErrHandler
    
    lbl(txtName).Caption = mpatiInfo.姓名
    lbl(txtSex).Caption = mpatiInfo.性别
    lbl(txtAge).Caption = mpatiInfo.年龄
    lbl(txtInNumber).Caption = mpatiInfo.住院号
    
    txtPrePay.Text = Format(mdblPrePay, "0.00")
    txtPrePay.Tag = Nvl(mdblPrePay)
    If Nvl(mstrRemark) = "" Then
        lblremark.Visible = False
    Else
        lblremark.Caption = "说明:" & mstrRemark
        If LenB(lblremark.Caption) > 50 Then lblremark.Caption = MidB(lblremark.Caption, 1, 50) & "……"
    End If
    Call SetPatiControl
    InitFace = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function InitData(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Boolean
    '初始化数据
    On Error GoTo ErrHandler
    '获取病人信息
     If GetPatiInfo(lng病人ID, lng主页ID, mpatiInfo) = False Then
        MsgBox "未找到病人信息！", vbInformation, gstrSysName
        Exit Function
    End If
        
    InitData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetPatiControl()
    '设置病人信息控件位置
    Dim sngSplit As Single
    
    sngSplit = 200
    On Error Resume Next
    lbl(txtName).Left = lbl(lblName).Left + lbl(lblName).Width
    
    lbl(lblSex).Left = lbl(txtName).Left + lbl(txtName).Width + sngSplit
    lbl(txtSex).Left = lbl(lblSex).Left + lbl(lblSex).Width
    
    lbl(lblAge).Left = lbl(txtSex).Left + lbl(txtSex).Width + sngSplit
    lbl(txtAge).Left = lbl(lblAge).Left + lbl(lblAge).Width
    
    lbl(lblInNumber).Left = lbl(txtAge).Left + lbl(txtAge).Width + sngSplit
    lbl(txtInNumber).Left = lbl(lblInNumber).Left + lbl(lblInNumber).Width
End Sub

Private Function GetBillData(ByVal lng病人ID As Long, ByRef strData As String) As Boolean
    '通过服务获取数据
    Dim strJsonIn As String
    
    On Error GoTo ErrHandler
    
    '调用新门诊“门诊预交金额转住院”服务
    '    输入    编码               名称      说明                  数据类型        备注
    '            pid                患者ID                          Number(18)      非空
    '
    '    输出    编码               名称       说明                 数据类型        备注
    '            result             执行结果   1-成功；-1-失败      Number(1)       非空
    '            errmsg             错误消息   失败时返回错误消息   Varchar2(200)
    '            prepaid_payment    预交金                          Number(18,2)    非空
    '            remark             备注       备注信息，如:“新门诊现金转入”，存入ZLHIS“病人预交记录.摘要”中
    '                                                               VARCHAR2(50)
                         
    strJsonIn = "{""head"":{""bizno"":""RJ004"",""sysno"":""ZLDAYROOM"",""time"":"""",""action_no"":"""",""tarno"":""03""}"
    strJsonIn = "{""input"":" & strJsonIn & ",""pid"":" & lng病人ID & "}}"
    Call Sys.NewSystemSvr("新门诊系统", "门诊预交金额转住院", strJsonIn, strData)
    If strData = "" Then strData = "{}"
    If Val(zlStr.JSONParse("result", strData)) <> 1 Then
        MsgBox "获取可门诊预交金额转住院信息时出错！" & vbCrLf & _
            zlStr.JSONParse("errmsg", strData), vbInformation, gstrSysName
        Exit Function
    End If
    mstrRemark = Nvl(zlStr.JSONParse("remark", strData))
    mdblPrePay = Val(zlStr.JSONParse("prepaid_payment", strData))
    If Nvl(mdblPrePay) = 0 Then
        MsgBox "该病人无可用门诊预交转住院预交金额！", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    
    GetBillData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SaveDate() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:对当前输入的预交款单据存盘
    '   lng病人ID - 病人ID
    '   lng主页ID - 主页ID
    '返回:成功返回True,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNO As String, strSQL As String
    Dim lng预交ID As Long
    
    strNO = zlDatabase.GetNextNo(11)
    lng预交ID = zlDatabase.GetNextId("病人预交记录")
    
    'Zl_病人预交记录_Insert_S
    strSQL = "Zl_病人预交记录_Insert_S("
    '  Id_In         病人预交记录.ID%Type,
    strSQL = strSQL & "" & lng预交ID & ","
    '  单据号_In     病人预交记录.NO%Type,
    strSQL = strSQL & "'" & strNO & "',"
    '  票据号_In     票据使用明细.号码%Type,
    strSQL = strSQL & "NULL,"
    '  病人id_In     病人预交记录.病人id%Type,
    strSQL = strSQL & "" & ZVal(mlng病人ID) & ","
    '  主页id_In     病人预交记录.主页id%Type,
    strSQL = strSQL & "" & ZVal(mlng主页ID) & ","
    '  姓名_In         病人预交记录.姓名%Type,
    strSQL = strSQL & "'" & mpatiInfo.姓名 & "',"
    '  性别_In         病人预交记录.性别%Type,
    strSQL = strSQL & "'" & mpatiInfo.性别 & "',"
    '  年龄_In         病人预交记录.年龄%Type,
    strSQL = strSQL & "'" & mpatiInfo.年龄 & "',"
    '  门诊号_In       病人预交记录.门诊号%Type,
    strSQL = strSQL & "NULL,"
    '  住院号_In       病人预交记录.住院号%Type,
    strSQL = strSQL & ZVal(mpatiInfo.住院号) & ","
    '  付款方式名称_In 病人预交记录.付款方式名称%Type,
    strSQL = strSQL & "'" & mpatiInfo.医疗付款方式 & "',"
    '  科室id_In     病人预交记录.科室id%Type,
    strSQL = strSQL & "NULL,"
    '  金额_In       病人预交记录.金额%Type,
    strSQL = strSQL & "" & mdblPrePay & ","
    '  结算方式_In   病人预交记录.结算方式%Type,
    strSQL = strSQL & "'" & "现金" & "',"
    '  结算号码_In   病人预交记录.结算号码%Type,
    strSQL = strSQL & "NULL,"
    '  缴款单位_In   病人预交记录.缴款单位%Type,
    strSQL = strSQL & "NULL,"
    '  单位开户行_In 病人预交记录.单位开户行%Type,
    strSQL = strSQL & "NULL,"
    '  单位帐号_In   病人预交记录.单位帐号%Type,
    strSQL = strSQL & "NULL,"
    '  摘要_In       病人预交记录.摘要%Type,
    strSQL = strSQL & "'" & mstrRemark & "',"
    '  操作员编号_In 病人预交记录.操作员编号%Type,
    strSQL = strSQL & "'" & UserInfo.编号 & "',"
    '  操作员姓名_In 病人预交记录.操作员姓名%Type,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '  领用id_In     票据使用明细.领用id%Type,
    strSQL = strSQL & "NULL,"
    '  预交类别_In   病人预交记录.预交类别%Type := Null,
    strSQL = strSQL & " 2)"
    On Error GoTo errH
    
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    SaveDate = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub txtPrePay_Change()
    If Val(txtPrePay.Text) <> 0 Then mdblPrePay = Val(txtPrePay.Text)
End Sub

Private Sub txtPrePay_GotFocus()
    zlControl.TxtSelAll txtPrePay
End Sub

Private Sub txtPrePay_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtPrePay_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtPrePay, KeyAscii, m金额式
End Sub

Private Function CheckPrePayValid() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查门诊预交金额有效性
    '返回:成功返回True,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    If Val(txtPrePay.Text) > Val(txtPrePay.Tag) Then
        MsgBox "该病人门诊预交金额转住院最大不能超过" & txtPrePay.Tag & "!", vbOKOnly + vbInformation, gstrSysName
        txtPrePay.Text = Format(Val(txtPrePay.Tag), "0.00")
        txtPrePay.SetFocus: zlControl.TxtSelAll txtPrePay: Exit Function
    
    ElseIf Val(txtPrePay.Text) <= 0 Then
        MsgBox "门诊预交金额转住院金额无效,必须在0到" & txtPrePay.Tag & "之间!", vbOKOnly + vbInformation, gstrSysName
        txtPrePay.Text = Format(Val(txtPrePay.Tag), "0.00")
        txtPrePay.SetFocus: zlControl.TxtSelAll txtPrePay: Exit Function
    End If
    
    CheckPrePayValid = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetPatiInfo(ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
                          ByRef PatiPageInfo As clsPatientInfo) As Boolean
    '功能：根据病人id和主页id获取病人信息和病案主页中的信息
    '入参：lng病人id-病人id；lng主页id-主页id
    '出参：PatiPageInfo-病案主页中的信息
    '返回：获取成功返回true,否则返回false
    Dim str病人id As String
    
    On Error GoTo errHandle
  
    '读取指定住院次数住院的信息
    str病人id = lng病人ID & ":" & lng主页ID
    Call GetPatiPageInforByID(str病人id, PatiPageInfo, False)
    If PatiPageInfo.病人ID = 0 Then Exit Function
      
    GetPatiInfo = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
