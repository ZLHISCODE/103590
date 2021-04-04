VERSION 5.00
Object = "*\A..\zlIDKind\zlIDKind.vbp"
Begin VB.Form frmBindPatientNo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "门诊号绑定"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4815
   Icon            =   "frmBindPatientNO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   4815
   StartUpPosition =   1  '所有者中心
   Begin zlIDKind.IDKindNew IDKind 
      Height          =   375
      Left            =   960
      TabIndex        =   16
      Top             =   240
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      Appearance      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   10.5
      FontName        =   "宋体"
      IDKind          =   -1
      BackColor       =   -2147483633
   End
   Begin VB.TextBox txt门诊号 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   960
      TabIndex        =   15
      Top             =   735
      Width           =   3615
   End
   Begin VB.Frame Frame1 
      Height          =   90
      Left            =   0
      TabIndex        =   14
      Top             =   2640
      Width           =   4905
   End
   Begin VB.TextBox txtPatient 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1590
      TabIndex        =   3
      ToolTipText     =   "热键:F11"
      Top             =   240
      Width           =   2400
   End
   Begin VB.CommandButton cmdYb 
      Caption         =   "医保"
      Height          =   345
      Left            =   4020
      TabIndex        =   2
      Top             =   240
      Width           =   555
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3555
      TabIndex        =   1
      ToolTipText     =   "热键：F2"
      Top             =   2880
      Width           =   1110
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2385
      TabIndex        =   0
      Top             =   2880
      Width           =   1110
   End
   Begin VB.Label txt身份证 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   960
      TabIndex        =   13
      Top             =   1740
      Width           =   3600
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "身份证"
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
      Left            =   240
      TabIndex        =   12
      Top             =   1800
      Width           =   630
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "年龄"
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
      Left            =   2835
      TabIndex        =   11
      Top             =   2250
      Width           =   420
   End
   Begin VB.Label lbl险类 
      AutoSize        =   -1  'True
      Caption         =   "医保"
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
      Left            =   450
      TabIndex        =   10
      Top             =   1290
      Width           =   420
   End
   Begin VB.Label txt险类 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   960
      TabIndex        =   9
      Top             =   1230
      Width           =   3615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "性别"
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
      Left            =   450
      TabIndex        =   8
      Top             =   2250
      Width           =   420
   End
   Begin VB.Label txt性别 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   960
      TabIndex        =   7
      Top             =   2190
      Width           =   1200
   End
   Begin VB.Label txt年龄 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3345
      TabIndex        =   6
      Top             =   2190
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "门诊号"
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
      Left            =   240
      TabIndex        =   5
      Top             =   810
      Width           =   630
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "病人"
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
      Left            =   450
      TabIndex        =   4
      Top             =   300
      Width           =   420
   End
End
Attribute VB_Name = "frmBindPatientNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mobjIDCard As zlIDCard.clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private mrsInfo As ADODB.Recordset
Private mblnOlnyBJYB As Boolean
Private mint预约失效次数 As Integer
Private mstrPrivs As String, mintIDKind As Integer
Private mstrNo As String
Private mblnOk As Boolean
Private mbln允许住院病人挂号 As Boolean
Private mblnNotClick As Boolean
Private mstrYBPati   As String
Private mintInsure   As Integer
Private mlngPatient As Long
Private Const mlngModule = 1111
'-----------------------------------------------------------------------------------
'结算卡相关
Private mstrPassWord As String

Private Sub cmdEnter_Click()
       Dim blnNoPrint   As Boolean
       If mrsInfo Is Nothing Then Exit Sub
       If mrsInfo.RecordCount = 0 Then Exit Sub
       If txtPatient.Text = "" Or mrsInfo!病人ID <> txtPatient.Tag Then Exit Sub
       cmdEnter.Enabled = False
       If SaveData() = False Then
            cmdEnter.Enabled = True
            Exit Sub
       End If
       Select Case gByt打印病人条码
       Case 0: blnNoPrint = True
       Case 1: blnNoPrint = False
       Case 2:
              If MsgBox("是否需要打印病人条码？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    blnNoPrint = True
              End If
       End Select
       If Not blnNoPrint Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1111_2", Me, "病人ID=" & Val(Nvl(mrsInfo!病人ID)), 2)
       End If
       
       Call ClearControlContent
End Sub

Private Sub ClearControlContent()
        Me.txtPatient.Tag = ""
        mlngPatient = 0
        Me.txt门诊号.Tag = ""
        txt险类.Caption = ""
        txt性别.Caption = ""
        txt年龄.Caption = ""
        txtPatient.Text = ""
        Me.txt门诊号.Text = ""
        txt身份证.Caption = ""
        Set mrsInfo = Nothing
       cmdEnter.Enabled = True
End Sub
'-----------------------------------------------------------------------------------
Private Sub cmdYb_Click()
     '医保身份证验证
     Call zlInusreIdentify
     
End Sub
Private Sub zlInusreIdentify()
    '------------------------------------------------------------------------------------------------------------------------
    '功能：医保身份验卡
    '编制：刘兴洪
    '日期：2010-07-14 11:32:08
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim lng病人ID As Long
    Dim str病人类型 As String
    If mrsInfo Is Nothing Then
        lng病人ID = 0
        str病人类型 = ""
    Else
        lng病人ID = Val(Nvl(mrsInfo!病人ID))
        str病人类型 = Nvl(mrsInfo!病人类型)
    End If
     
    mstrYBPati = gclsInsure.Identify(3, lng病人ID, mintInsure)
    If mstrYBPati <> "" Then
        '空或：0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);8病人ID
        If UBound(Split(mstrYBPati, ";")) >= 8 Then
            If IsNumeric(Split(mstrYBPati, ";")(8)) Then lng病人ID = Val(Split(mstrYBPati, ";")(8))
        End If
        If lng病人ID <> 0 Then
            '问题:29283
            '  -- 参数:调用场合-1-挂号;2-收费
            '  --        病人id_In-病人ID(未建档的,传入零)
            '  --        卡号_In: 刷卡卡号;未刷卡时,为空
            '  --         刷卡方式_In:  1-普能刷卡;2-医保刷卡
            txtPatient.Text = "-" & lng病人ID
            txtPatient_KeyPress (13)
            If str病人类型 = "" Then txtPatient.ForeColor = vbRed
            Me.txtPatient.SetFocus
        Else
            mstrYBPati = "": mintInsure = 0: txtPatient.SetFocus
        End If
    Else
        '修改问题：38917 作者：冉勇
        If Not txtPatient.Enabled Then txtPatient.Enabled = True
         mstrYBPati = "": mintInsure = 0: txtPatient.SetFocus
    End If
End Sub

 
Private Sub cmdCancel_Click()
    mblnOk = False: Unload Me
    
End Sub

Private Sub cmdOK_Click()
     If Me.txtPatient.Text = "" Or Me.txtPatient.Text <> Me.txtPatient.Tag Then Exit Sub
        
End Sub

Private Function CheckPatient(ByVal str门诊号 As String, ByVal lng病人ID As Long) As Boolean

'功能：判断指定门诊号是否已经存在于数据库中
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strMsg As String
    Dim lng合并ID As Long
    On Error GoTo errH
    strSQL = "    " & vbNewLine & " Select B.病人id As ID, B.病人id, B.姓名, B.性别, B.年龄, B.门诊号, B.出生日期, B.身份证号, B.家庭地址, B.工作单位,"
    strSQL = strSQL & vbNewLine & "      A.名称 险类名称"
    strSQL = strSQL & vbNewLine & " From 病人信息 B, 保险类别 A"
    strSQL = strSQL & vbNewLine & " Where B.险类 = A.序号(+) And B.停用时间 Is Null  "
    strSQL = strSQL & vbNewLine & " And b.门诊号=[1] And 病人ID<>[2]"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlRegEvent", str门诊号, lng病人ID)
    If rsTmp.RecordCount > 0 Then
       If Nvl(mrsInfo!姓名) <> Nvl(rsTmp!姓名) Then
            strMsg = "当前病人姓名[" & Nvl(mrsInfo!姓名) & "]与现有门诊号[" & str门诊号 & "]的病人[" & Nvl(rsTmp!姓名) & "]的病人不一致!" & vbCrLf & _
                    "是否把姓名为[" & Nvl(mrsInfo!姓名) & "]的信息合并到姓名为[" & rsTmp!姓名 & "]中？"
            If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            lng合并ID = Val(rsTmp!病人ID)
            If zlPatiMerge(lng病人ID, lng合并ID, False) = False Then Exit Function
       Else
            strMsg = "当前系统已经存在姓名为[" & mrsInfo!姓名 & "]并且门诊号为[" & str门诊号 & "]的病人,是否将当前病人合并到门诊号为[" & str门诊号 & "]的病人中?"
            If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            lng合并ID = Val(rsTmp!病人ID)
            If zlPatiMerge(lng病人ID, lng合并ID, False) = False Then Exit Function
            Call GetPatient("-" & lng合并ID)
       End If
        strSQL = "    " & vbNewLine & " Select B.病人id As ID, B.病人id, B.姓名, B.性别, B.年龄, B.门诊号, B.出生日期, B.身份证号, B.家庭地址, B.工作单位,"
        strSQL = strSQL & vbNewLine & "      A.名称 险类名称"
        strSQL = strSQL & vbNewLine & " From 病人信息 B, 保险类别 A"
        strSQL = strSQL & vbNewLine & " Where B.险类 = A.序号(+) And B.停用时间 Is Null  "
        strSQL = strSQL & vbNewLine & " And b.病人ID=[1] "
        Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng合并ID)
        mlngPatient = lng合并ID
    End If
      CheckPatient = True
       Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function
Private Function SaveData() As Boolean
    '
    Dim strSQL     As String
    Dim lng病人ID  As Long
    Dim str医保    As String
    Dim str门诊号  As String
    Dim strDat     As String
    Dim Datsys     As Date
    Dim intInsure  As Integer
    On Error GoTo Hd
    
    If txt门诊号.Text = "" Then
        MsgBox "门诊号不能为空!", vbInformation, Me.Caption
        Exit Function
    End If
    str门诊号 = txt门诊号.Text
    If CheckPatient(str门诊号, mlngPatient) = False Then Exit Function
    If Exist门诊号(Me.txt门诊号.Text, mlngPatient) Then
        MsgBox "门诊号已经被使用!", vbInformation, Me.Caption
        Exit Function
    End If
    Datsys = zlDatabase.Currentdate
    strDat = "to_date('" & Format(Datsys, "yyyy-MM-dd hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')"
    If Trim(Me.txt险类.Caption) <> "" And mstrYBPati <> "" Then
        str医保 = Split(mstrYBPati, ";")(1)
        intInsure = mintInsure
    End If
  'Zl_病人信息_绑定门诊号(
  '病人id_In   病人信息.病人id%Type,
  '门诊号_In   病人信息.门诊号%Type,
  '登记时间_In 病人信息.登记时间%Type,
  '医保号_In   病人信息.医保号%Type := Null,
  '险类_In     病人信息.险类%Type := Null,
  '处理类型_In Number:=0
  '--功能：处理挂号病人信息门诊号 绑定
  '--参数：
  '--处理类型：
  '--             0=仅更新门诊号  不对病人的费用记录 和挂号记录 门诊号 进行更新
  '--             1=更新门诊号  同时处理 病人的 费用信息
    strSQL = "Zl_病人信息_绑定门诊号"
    strSQL = strSQL & "(" & mlngPatient & ","
    strSQL = strSQL & "'" & str门诊号 & "'" & ","
    strSQL = strSQL & strDat & ","
    strSQL = strSQL & IIf(str医保 = "", "NUll", "'" & str医保 & "'") & ","
    strSQL = strSQL & IIf(intInsure = 0, "NULL", intInsure) & ","
    strSQL = strSQL & "1)"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    SaveData = True
    Exit Function
Hd:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim blnCancel As Boolean
    Select Case KeyCode
        Case vbKeyF4
            If Shift = vbCtrlMask Then
                If IDKind.Enabled Then IDKind.IDKind = IDKind.GetKindIndex("IC卡号"): Call IDKind_Click(IDKind.GetCurCard)
            ElseIf Me.ActiveControl Is txtPatient Then
                If IDKind.Enabled Then
                    If Shift = vbShiftMask Then
                        IDKind.IDKind = IIf(IDKind.IDKind = 0, UBound(Split(IDKind.IDkindStr, ";")), IDKind.IDKind - 1)
                    Else
                        IDKind.IDKind = IIf(IDKind.IDKind = UBound(Split(IDKind.IDkindStr, ";")), 0, IDKind.IDKind + 1)
                    End If
                End If
            End If
        Case vbKeyF11
            If txtPatient.Enabled And txtPatient.Visible And Not txtPatient.Locked Then
                If Me.ActiveControl Is txtPatient Then
                    IDKind.IDKind = IIf(IDKind.IDKind = UBound(Split(IDKind.IDkindStr, ";")), IDKind.GetKindIndex("姓名"), IDKind.IDKind + 1)
                Else
                    txtPatient.SetFocus
                End If
            End If
        Case vbKeyReturn
       
    End Select
End Sub
Private Sub Form_Load()
    Dim strTemp As String
    Set mobjIDCard = New clsIDCard
    Set mobjICCard = New clsICCard
    Call mobjIDCard.SetParent(Me.Hwnd)
    Call mobjICCard.SetParent(Me.Hwnd)
    If gobjSquare.objSquareCard Is Nothing Then
        CreateSquareCardObject gfrmMain, mlngModule
    End If
    InitIDKind
    Set mobjICCard.gcnOracle = gcnOracle
    Call GetRegInFor(g私有模块, Me.Name, "idkind", strTemp)
    mintIDKind = Val(strTemp)
    If mintIDKind > 0 And mintIDKind <= IDKind.ListCount Then IDKind.IDKind = mintIDKind
    mbln允许住院病人挂号 = zlDatabase.GetPara("允许住院病人挂号", glngSys, mlngModule, 0) = "1"
    Me.Icon = frmRegist.Icon
End Sub

Private Function InitIDKind() As Boolean
    Dim objCard As Card
    Dim lngCardID As Long
    Call IDKind.zlInit(Me, glngSys, mlngModule, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", txtPatient)
    lngCardID = Val(zlDatabase.GetPara("缺省医疗卡类别", glngSys, mlngModule, 0))
    If lngCardID <> 0 Then
        IDKind.DefaultCardType = lngCardID
    End If
    Set objCard = IDKind.GetfaultCard
    Set gobjSquare.objDefaultCard = objCard
    IDKind.ShowPropertySet = InStr(";" & mstrPrivs & ";", "参数设置") > 0
    If IDKind.Cards.按缺省卡查找 And Not objCard Is Nothing Then
        gobjSquare.bln缺省卡号密文 = objCard.卡号密文规则 <> ""
        gobjSquare.int缺省卡号长度 = objCard.卡号长度
         
       
    Else
        gobjSquare.bln缺省卡号密文 = IDKind.Cards.加密显示
        gobjSquare.int缺省卡号长度 = 100
    End If
    
        
End Function

Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    If Not mobjIDCard Is Nothing Then
        Call mobjIDCard.SetEnabled(False)
        Set mobjIDCard = Nothing
    End If
    If Not mobjICCard Is Nothing Then
        Call mobjICCard.SetEnabled(False)
        Set mobjICCard = Nothing
    End If
    Set mrsInfo = Nothing
    mintIDKind = IDKind.IDKind
    Call SaveRegInFor(g私有模块, Me.Name, "idkind", mintIDKind)
     
End Sub


Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng卡类别ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
    
    If IsCardType(IDKind, "IC卡号") Then
        If Not mobjICCard Is Nothing Then
            txtPatient.Text = mobjICCard.Read_Card()
            If txtPatient.Text <> "" Then
                Call GetPatient(Trim(txtPatient))
            End If
        End If
        Exit Sub
    End If
    lng卡类别ID = IDKind.GetCurCard.接口序号
    If lng卡类别ID = 0 Then Exit Sub
    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModule, lng卡类别ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtPatient.Text = strOutCardNO
    If txtPatient.Text <> "" Then
    Call GetPatient(Trim(txtPatient.Text))
    End If
End Sub

Private Sub IDKind_ItemClick(index As Integer, objCard As zlIDKind.Card)
     '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|
    '是否存在帐户(1-存在帐户;0-不存在帐户)|卡号密文(第几位至第几位加密,空为不加密)
     
    Set gobjSquare.objCurCard = objCard
     
    If objCard.接口序号 > 0 Then
        txtPatient.MaxLength = IDKind.GetCardNoLen
        txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
    Else
        txtPatient.MaxLength = 0: txtPatient.PasswordChar = ""
    End If
    
    '需要清除信息,避免刷卡后,再切换,造成密文显示失去意义
    If txtPatient.Text <> "" And Not mblnNotClick Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    zlControl.TxtSelAll txtPatient
End Sub

Private Sub IDKind_KeyPress(KeyAscii As Integer)
'
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    Dim lngPreIDKind As Long, intIndex As Integer
    Dim dtDate As Date
    Dim blnNew As Boolean

    If txtPatient.Locked Or txtPatient.Text <> "" Then Exit Sub 'Or Not Me.ActiveControl Is txtPatient
    mblnNotClick = True

    intIndex = IDKind.GetKindIndex(objCard.名称)
    lngPreIDKind = IDKind.IDKind
    If intIndex > 0 Then IDKind.IDKind = intIndex

    txtPatient.Text = objPatiInfor.卡号
    Call txtPatient_KeyPress(vbKeyReturn)
    If mrsInfo Is Nothing Then
        blnNew = True
    ElseIf mrsInfo.State <> 1 Then
        blnNew = True
    End If
    IDKind.IDKind = lngPreIDKind
    mblnNotClick = False
End Sub
Private Sub txtPatient_Change()
    txtPatient.Tag = ""
  
    txtPatient.ForeColor = &H80000008
    If Me.ActiveControl Is txtPatient Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
        IDKind.SetAutoReadCard txtPatient.Text = ""
    End If
   
End Sub

Private Sub txtPatient_GotFocus()
   zlControl.TxtSelAll txtPatient
      If txtPatient.Text = "" And Not txtPatient.Locked Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(True)
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(True)
        IDKind.SetAutoReadCard True
    End If
  Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim lngID As Long, lngUnit As Long, i As Long
    Dim rsTmp As ADODB.Recordset, strInfo As String
    Dim strSQL As String, curTotal As Currency
    Dim blnCard As Boolean, blnICCard As Boolean
    
    On Error GoTo errH
    If txtPatient.Locked Then Exit Sub
    
    If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    
    If IsCardType(IDKind, "姓名") Then
        blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, glngSys)
    ElseIf IsCardType(IDKind, "门诊号") Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If Not (IsNumeric(Chr(KeyAscii)) Or Chr(KeyAscii) = "-") Then KeyAscii = 0: Exit Sub
        End If
    End If
    If blnCard And Len(txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Trim(txtPatient.Text) <> "" Then
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        ElseIf IsNumeric(txtPatient.Tag) Then
            KeyAscii = 0
            'If txtPatient.Tag <> "" Then
            '刷新病人信息:"-病人ID"
            If Val(txtPatient.Tag) <> 0 Then
                Call zlCommFun.PressKey(vbKeyTab)
                Exit Sub
            End If
            Call GetPatient(txtPatient.Tag, False)
            Exit Sub
        End If
        KeyAscii = 0
        If IsCardType(IDKind, "IC卡号") Then blnICCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
        Call GetPatient(txtPatient.Text, blnCard)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub GetPatient(ByVal strInput As String, Optional blnCard As Boolean)
    '------------------------------------------------------------------------------------------------------------------------
    '功能：获取病人信息
    '入参：blnCard=是否就诊卡刷卡
    '出参：
    '返回：
    '编制：刘兴洪
    '日期：2010-07-16 14:24:14
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strTemp As String
    Dim blnSame As Boolean, blnCancel As Boolean
    Dim cur余额 As Currency, curMoney As Currency
    Dim i As Integer, strPati As String
    Dim vRect As RECT, str非在院 As String
    Dim strSQL As String, lng卡类别ID As Long, strPassWord As String, strErrMsg As String
    Dim strTmp As String
    Dim lng病人ID As Long, blnHavePassWord As Boolean
    Dim blnIDCard As Boolean
    On Error GoTo errH
    If Not mbln允许住院病人挂号 Then
        str非在院 = " And Not Exists(Select 1 From 病案主页 Where 病人ID=B.病人ID And 主页ID=B.主页ID And Nvl(病人性质,0)=0 And 出院日期 is Null)"
    End If
    
    strSQL = ""
    
    If Not (blnCard Or IDKind.GetCurCard.接口序号 = IDKind.GetfaultCard.接口序号) _
         And IDKind.GetCurCard.名称 Like "*姓名" And Not (Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2))) Then
      
        If IDKind.Cards.按缺省卡查找 And Not IDKind.GetfaultCard Is Nothing Then
            lng卡类别ID = IDKind.GetfaultCard.接口序号
        ElseIf IDKind.GetCurCard.接口序号 > 0 Then
            lng卡类别ID = IDKind.GetCurCard.接口序号
        Else
            lng卡类别ID = -1
        End If
        
        '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|是否存在帐户(1-存在帐户;0-不存在帐户);…
        If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
        If lng病人ID <= 0 Then lng病人ID = 0
        strInput = "-" & lng病人ID
        blnHavePassWord = True
        strSQL = strSQL & " And B.病人ID=[2] " & str非在院
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then
        '门诊号
        strSQL = strSQL & " And B.门诊号=[2]" & str非在院
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then
        '病人ID
        strSQL = strSQL & " And B.病人ID=[2]" & str非在院
    Else
        Select Case IDKind.GetCurCard.名称
            Case "姓名", "姓名或就诊卡"
                '姓名
                blnSame = False
                If Not mrsInfo Is Nothing Then
                    If txtPatient.Text = mrsInfo!姓名 Then blnSame = True
                End If
                 
                
                If Not blnSame Then
                    If (Not gblnSeekName) Or (gblnSeekName And Len(txtPatient.Text) < 2) Then
                        txt险类.Caption = "": txt性别 = "": txt年龄 = ""
                        txtPatient.Text = "": Me.txt门诊号.Text = ""
                        txt年龄.Caption = "0"
                        txt身份证.Caption = ""
                        txt性别.Caption = ""
                        Set mrsInfo = Nothing: Exit Sub
                    Else
                       strSQL = strSQL & " And  B.姓名 Like [3]"
                       
                    End If
                Else
                    strSQL = strSQL & " And B.病人ID=[2]"
                    strInput = "-" & Val(mrsInfo!病人ID)
                End If
            Case "医保号"
                strInput = UCase(strInput)
                If mblnOlnyBJYB And zlCommFun.ActualLen(strInput) >= 9 Then
                    '仅北京医保才有效:见问题:问题:26982
                    strSQL = strSQL & " And B.医保号 like [3] " & str非在院
                    strTemp = Left(strInput, 9) & "%"
                Else
                    strSQL = strSQL & " And B.医保号=[1]" & str非在院
                End If
            Case "身份证号", "身份证", "二代身份证"
'                 strInput = UCase(strInput)
'                If gobjSquare.objSquareCard.zlGetPatiID("身份证", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
'                strSQL = strSQL & " And B.病人ID=[2]" & str非在院
'                strInput = "-" & lng病人ID
                 blnIDCard = True
                 strSQL = strSQL & " And B.身份证号=[1] " & str非在院
            Case "IC卡号", "IC卡"
                strInput = UCase(strInput)
                If gobjSquare.objSquareCard.zlGetPatiID("IC卡", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                strSQL = strSQL & " And B.病人ID=[2]" & str非在院
                strInput = "-" & lng病人ID
            Case "门诊号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = strSQL & " And B.门诊号=[1]" & str非在院
             Case Else
                '其他类别的,获取相关的病人ID
                If Val(IDKind.GetCurCard.接口序号) >= 0 Then
                    lng卡类别ID = Val(IDKind.GetCurCard.接口序号)
                    If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                    If lng病人ID = 0 Then lng病人ID = 0
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(IDKind.GetCurCard.名称, strInput, False, lng病人ID, _
                        strPassWord, strErrMsg) = False Then lng病人ID = 0
                End If
                If lng病人ID <= 0 Then lng病人ID = 0
                strSQL = strSQL & " And B.病人ID=[2]" & str非在院
                strInput = "-" & lng病人ID
                blnHavePassWord = True
        End Select
    End If
    strTmp = strSQL
    strSQL = "    " & vbNewLine & " Select /*+Rule */distinct  B.病人id As ID, Decode(sign(nvl(ylkxx.病人id,0)),0,'','√') as 三方账户, B.病人id,B.姓名, B.性别, B.年龄, B.门诊号, B.出生日期, B.身份证号, B.家庭地址, B.工作单位,"
    strSQL = strSQL & vbNewLine & "      A.名称 险类名称,B.病人类型"
    strSQL = strSQL & vbNewLine & " From 病人信息 B, 保险类别 A,医疗卡类别 YLK,病人医疗卡信息 YLKXX"
    strSQL = strSQL & vbNewLine & " Where B.险类 = A.序号(+) and b.病人id=ylkxx.病人id(+) and ylkxx.状态(+)=0 and  ylkxx.卡类别id=ylk.id(+)  and ylk.是否自制(+)=0 And B.停用时间 Is Null   "
    strSQL = strSQL & vbNewLine & strTmp
     
    On Error GoTo errH
     If Not blnIDCard Then
        vRect = zlControl.GetControlRect(txtPatient.Hwnd)
        Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput, Mid(strInput, 2), strInput & "%")
     Else
        vRect = zlControl.GetControlRect(txtPatient.Hwnd)
     Set mrsInfo = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "病人查找", 1, "√", "请选择病人", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, False, True, strInput, CStr(Mid(strInput, 2)), strInput & "%")
     End If
     If Not mrsInfo Is Nothing And Not blnCancel Then  'And Not blnCancel
        If mrsInfo.RecordCount = 0 Then
            Set mrsInfo = Nothing
            txt险类.Caption = "": txt性别 = "": txt年龄 = ""
            txtPatient.Text = "": Me.txt门诊号.Text = ""
            txt年龄.Caption = "0"
            txt身份证.Caption = ""
            txt性别.Caption = ""
            Exit Sub
        ElseIf mrsInfo!ID = 0 Then  '没有找到病人信息
            Set mrsInfo = Nothing
            txt险类.Caption = "": txt性别 = "": txt年龄 = ""
            txtPatient.Text = "": Me.txt门诊号.Text = ""
            txt年龄.Caption = "0"
            txt身份证.Caption = ""
            txt性别.Caption = ""
            Exit Sub
        Else '获取到病人信息
          
            Me.txt门诊号.Tag = Nvl(mrsInfo!门诊号)
            txt险类.Caption = Nvl(mrsInfo!险类名称)
            txt性别.Caption = Nvl(mrsInfo!性别)
            txt年龄.Caption = Nvl(mrsInfo!年龄)
            txtPatient.Text = Nvl(mrsInfo!姓名)
            Me.txtPatient.Tag = Nvl(mrsInfo!ID)
            Me.txt门诊号.Text = Nvl(mrsInfo!门诊号)
            txt身份证.Caption = Nvl(mrsInfo!身份证号)
            mlngPatient = Val(Nvl(mrsInfo!ID))
            '74428:李南春，2014-7-7，病人姓名颜色处理
            Call SetPatiColor(txtPatient, Nvl(mrsInfo!病人类型), IIf(Trim(txt险类.Caption) <> "", vbRed, txtPatient.ForeColor))
        End If
    Else '取消选择
        txt险类.Caption = "": txt性别 = "": txt年龄 = ""
        txtPatient.Text = "": Me.txt门诊号.Text = ""
        mlngPatient = 0
        txt年龄.Caption = "0"
        txt身份证.Caption = ""
        txt性别.Caption = ""
        Set mrsInfo = Nothing: Exit Sub
    End If
    
'    ''''''''''''''''
'
'     If Not mrsInfo Is Nothing And Not blnCancel Then
'        If mrsInfo!ID = 0 Then '没有找到病人信息
'            Set mrsInfo = Nothing
'            txt险类.Caption = "": txt性别 = "": txt年龄 = ""
'            txtPatient.Text = "": Me.txt门诊号.Text = ""
'            txt年龄.Caption = "0"
'            txt身份证.Caption = ""
'            txt性别.Caption = ""
'            Exit Sub
'        Else '获取到病人信息
'
'          Me.txt门诊号.Tag = Nvl(mrsInfo!门诊号)
'          txt险类.Caption = Nvl(mrsInfo!险类名称)
'          txt性别.Caption = Nvl(mrsInfo!性别)
'          txt年龄.Caption = Nvl(mrsInfo!年龄)
'          txtPatient.Text = Nvl(mrsInfo!姓名)
'          Me.txtPatient.Tag = Nvl(mrsInfo!ID)
'          Me.txt门诊号.Text = Nvl(mrsInfo!门诊号)
'          txt身份证.Caption = Nvl(mrsInfo!身份证号)
'          mlngPatient = Val(Nvl(mrsInfo!ID))
'          If Trim(txt险类.Caption) <> "" Then Me.txtPatient.ForeColor = vbRed
'        End If
'    Else '取消选择
'        txt险类.Caption = "": txt性别 = "": txt年龄 = ""
'        txtPatient.Text = "": Me.txt门诊号.Text = ""
'        mlngPatient = 0
'        txt年龄.Caption = "0"
'        txt身份证.Caption = ""
'        txt性别.Caption = ""
'        Set mrsInfo = Nothing: Exit Sub
'    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
 
 
 

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    Dim lngPreIDKind As Long
    If Not txtPatient.Locked And txtPatient.Text = "" And Me.ActiveControl Is txtPatient Then
        mblnNotClick = True
        lngPreIDKind = IDKind.IDKind
        IDKind.IDKind = IDKind.GetKindIndex("身份证号")
        txtPatient.Text = strID
        Call GetPatient(Trim(txtPatient.Text))
        IDKind.IDKind = lngPreIDKind
        mblnNotClick = False
    End If
End Sub


Private Sub mobjICCard_ShowICCardInfo(ByVal strNO As String)
    Dim lngPreIDKind As Long
    If Not txtPatient.Locked And txtPatient.Text = "" And Me.ActiveControl Is txtPatient Then
        lngPreIDKind = IDKind.IDKind
        mblnNotClick = True
        IDKind.IDKind = IDKind.GetKindIndex("IC卡号")
        txtPatient.Text = strNO
        If txtPatient.Text <> "" Then
            Call GetPatient(Trim(txtPatient.Text))
        Else
            Call mobjICCard.SetEnabled(False) '如果不符合发卡条件，禁用继续自动读取
        End If
        IDKind.IDKind = lngPreIDKind
        mblnNotClick = False
        If Not txtPatient.Locked And Me.ActiveControl Is txtPatient Then mobjICCard.SetEnabled (txtPatient.Text = "")
    End If
End Sub


Private Sub txtPatient_LostFocus()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled False
End Sub

Private Sub txt门诊号_Change()
   Me.txt门诊号.Tag = ""
End Sub

Private Sub txt门诊号_GotFocus()
   zlControl.TxtSelAll txt门诊号
End Sub

 
'Private Function CheckPatiValid(ByVal strCard As String) As Boolean
'    '------------------------------------------------------------------------------------------------------------------------
'    '功能：检查指定输入的卡号是否合法
'    '入参：strCard-指定的卡号
'    '返回：合法,返回True,否则返回False
'    '编制：刘兴洪
'    '日期：2010-07-19 10:14:31
'    '说明：31182
'    '------------------------------------------------------------------------------------------------------------------------
'   Dim rsTmp As ADODB.Recordset, strSQL As String, lng病人ID As Long
'
'    strSQL = "Select Nvl(就诊状态,0) 就诊状态,病人ID,姓名,性别 From 病人信息 Where 就诊卡号 = [1]"
'    On Error GoTo errH
'    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strCard)
'    If rsTmp.RecordCount = 0 Then CheckPatiValid = True: Exit Function
'
'    '1.检查状态:原来主要是在输就诊卡时进行检查的,由于txt卡号_Validate事情,不一定能检查到,因此,本增加在按确定时,增加该检查
'    If Val(Nvl(rsTmp!就诊状态)) <> 0 Then
'        MsgBox "卡号为" & strCard & "的病人正在就诊或等待就诊,不能绑定该卡号.", vbInformation, gstrSysName
'        Exit Function
'    End If
'
'    '2.检查是否病人姓名相同
'    If Nvl(rsTmp!姓名) <> Trim(txtPatient.Text) And Val(txt卡号.Tag) = 0 Then
'       If MsgBox("持卡病人『" & Nvl(rsTmp!姓名) & "』与输入的病人『" & Trim(txtPatient.Text) & "』不一致,是否继续?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
'    End If
'
'    '3.挂号病人与刷就诊卡得出的病人是两个不同建档的病人
'    lng病人ID = Val(Nvl(rsTmp!病人ID))
'    If Val(txt卡号.Tag) <> lng病人ID And Val(txt卡号.Tag) <> 0 Then
'        If Nvl(rsTmp!姓名) <> Trim(txtPatient.Text) Then
'            If MsgBox("注意: " & vbCrLf & _
'                             "     持卡病人『" & Nvl(rsTmp!姓名) & "』与输入的病人『" & Trim(txtPatient.Text) & "』不一致," & vbCrLf & _
'                             "     但同时都是建档病人,是否将病人『" & Trim(txtPatient.Text) & "』合并到病人『" & Nvl(rsTmp!姓名) & "』中?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
'            '合并
'            If zlPatiMerge(Val(txt卡号.Tag), lng病人ID, True) = False Then Exit Function
'        Else '病人姓名相同,自动进行合并
'            '自动合并
'            If zlPatiMerge(Val(txt卡号.Tag), lng病人ID, False) = False Then Exit Function
'        End If
'        '重新刷新相关的数据
'        RaiseEvent PatiMerged(lng病人ID)
'
'    End If
'    CheckPatiValid = True
'    Exit Function
'errH:
'    If errCenter() = 1 Then Resume
'    Call SaveErrLog
'End Function

Private Sub txt门诊号_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call zlCommFun.PressKey(vbKeyTab)
        End If
End Sub


'获取idkind的默认kind值
Private Function IDKindDefaultKind() As Long
    Dim lngIndex As Long
    'IDkind的默认Kind
    If IDKind.DefaultCardType = "" Then
        lngIndex = -1
    Else
        If IsNumeric(IDKind.DefaultCardType) Then
           lngIndex = IDKind.GetKindIndex(IDKind.GetfaultCard.名称)
        Else
           lngIndex = IDKind.GetKindIndex(IDKind.DefaultCardType)
        End If
    End If
    IDKindDefaultKind = lngIndex
End Function

'控件名称是否匹配
Private Function IsCardType(ByVal IDKindCtl As IDKindNew, ByVal strCardName As String) As Boolean
    If IDKindCtl Is Nothing Then Exit Function
    If UCase(TypeName(IDKindCtl)) <> "IDKINDNEW" Then Exit Function
    Select Case strCardName
     Case "姓名", "姓名或就诊卡"
          IsCardType = IDKindCtl.GetCurCard.名称 Like "姓名*"
     Case "身份证", "身份证号", "二代身份证"
          IsCardType = IDKindCtl.GetCurCard.名称 Like "*身份证*"
     Case "IC卡号", "IC卡"
          IsCardType = IDKindCtl.GetCurCard.名称 Like "IC卡*"
     Case "医保号"
          IsCardType = IDKindCtl.GetCurCard.名称 = "医保号"
     Case "门诊号"
          IsCardType = IDKindCtl.GetCurCard.名称 = "门诊号"
     Case Else
            If IDKindCtl.GetCurCard Is Nothing Then Exit Function
            If Not IsNumeric(strCardName) Or Val(strCardName) <= 0 Then Exit Function
            If IDKindCtl.GetCurCard.接口序号 <= 0 Then Exit Function
            IsCardType = IDKindCtl.GetCurCard.接口序号 = Val(strCardName)
     End Select
End Function
                

