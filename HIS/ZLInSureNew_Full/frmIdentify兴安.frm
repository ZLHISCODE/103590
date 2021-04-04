VERSION 5.00
Begin VB.Form frmIdentify兴安 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "身份验证"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7080
   Icon            =   "frmIdentify兴安.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox TxtEdit 
      Height          =   300
      Index           =   1
      Left            =   4635
      MaxLength       =   20
      TabIndex        =   3
      Top             =   945
      Width           =   2295
   End
   Begin VB.ComboBox cbo交易类型 
      Height          =   300
      Left            =   4635
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   2527
      Width           =   2295
   End
   Begin VB.CommandButton cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5730
      TabIndex        =   33
      Top             =   4605
      Width           =   1100
   End
   Begin VB.CommandButton cmd确定 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4440
      TabIndex        =   32
      Top             =   4635
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   75
      Index           =   1
      Left            =   -75
      TabIndex        =   35
      Top             =   585
      Width           =   8340
   End
   Begin VB.Frame fra 
      Height          =   75
      Index           =   0
      Left            =   0
      TabIndex        =   34
      Top             =   4380
      Width           =   8340
   End
   Begin VB.TextBox TxtEdit 
      Height          =   285
      Index           =   0
      Left            =   1005
      MaxLength       =   20
      TabIndex        =   1
      Top             =   953
      Width           =   2265
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   15
      Left            =   4635
      TabIndex        =   31
      Top             =   4050
      Width           =   2325
   End
   Begin VB.Label lbl 
      Caption         =   "当年住院次数"
      Height          =   180
      Index           =   15
      Left            =   3525
      TabIndex        =   30
      Top             =   4110
      Width           =   1080
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   14
      Left            =   1365
      TabIndex        =   29
      Top             =   4065
      Width           =   1905
   End
   Begin VB.Label lbl 
      Caption         =   "已用慢病统筹"
      Height          =   180
      Index           =   14
      Left            =   225
      TabIndex        =   28
      Top             =   4110
      Width           =   1080
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   13
      Left            =   4635
      TabIndex        =   27
      Top             =   3675
      Width           =   2325
   End
   Begin VB.Label lbl 
      Caption         =   "当年已用自负段"
      Height          =   180
      Index           =   13
      Left            =   3360
      TabIndex        =   26
      Top             =   3735
      Width           =   1260
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   12
      Left            =   1365
      TabIndex        =   25
      Top             =   3690
      Width           =   1905
   End
   Begin VB.Label lbl 
      Caption         =   "当日已用药品"
      Height          =   180
      Index           =   12
      Left            =   240
      TabIndex        =   24
      Top             =   3735
      Width           =   1080
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "交易类型"
      Height          =   180
      Index           =   11
      Left            =   3885
      TabIndex        =   18
      Top             =   2587
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "病种名称"
      Height          =   180
      Index           =   9
      Left            =   255
      TabIndex        =   20
      Top             =   2962
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   1005
      TabIndex        =   21
      Top             =   2910
      Width           =   5925
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "病种编码"
      Height          =   180
      Index           =   8
      Left            =   255
      TabIndex        =   16
      Top             =   2587
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   1005
      TabIndex        =   17
      Top             =   2535
      Width           =   2265
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "帐户余额"
      Height          =   180
      Index           =   7
      Left            =   3885
      TabIndex        =   14
      Top             =   2182
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   7
      Left            =   4635
      TabIndex        =   15
      Top             =   2122
      Width           =   2295
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "帐户状态"
      Height          =   180
      Index           =   6
      Left            =   255
      TabIndex        =   12
      Top             =   2182
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   6
      Left            =   1005
      TabIndex        =   13
      Top             =   2130
      Width           =   2265
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "人员类别"
      Height          =   180
      Index           =   5
      Left            =   3885
      TabIndex        =   10
      Top             =   1755
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   5
      Left            =   4635
      TabIndex        =   11
      Top             =   1695
      Width           =   2295
   End
   Begin VB.Label lblInfor 
      AutoSize        =   -1  'True
      Caption         =   "通过滋卡验证人员身份，并将验证结果信息显示出来。"
      Height          =   180
      Left            =   675
      TabIndex        =   36
      Top             =   345
      Width           =   4320
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   105
      Picture         =   "frmIdentify兴安.frx":030A
      Top             =   75
      Width           =   480
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "医保卡号"
      Height          =   180
      Index           =   0
      Left            =   255
      TabIndex        =   0
      Top             =   1005
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "医保证号"
      Height          =   180
      Index           =   1
      Left            =   3885
      TabIndex        =   2
      Top             =   1005
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "姓名"
      Height          =   180
      Index           =   2
      Left            =   615
      TabIndex        =   4
      Top             =   1387
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "性别"
      Height          =   180
      Index           =   3
      Left            =   4245
      TabIndex        =   6
      Top             =   1387
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "年龄"
      Height          =   180
      Index           =   4
      Left            =   615
      TabIndex        =   8
      Top             =   1755
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "单位名称"
      Height          =   180
      Index           =   10
      Left            =   255
      TabIndex        =   22
      Top             =   3352
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   1005
      TabIndex        =   5
      Top             =   1335
      Width           =   2265
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   4635
      TabIndex        =   7
      Top             =   1335
      Width           =   975
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   1005
      TabIndex        =   9
      Top             =   1703
      Width           =   1455
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   10
      Left            =   1005
      TabIndex        =   23
      Top             =   3300
      Width           =   5940
   End
End
Attribute VB_Name = "frmIdentify兴安"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbytType As Byte            '0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐

Private mlng病人ID As Long
Private mstrReturn As String
Private mintPreCol As Integer, mintsort As Integer
Dim mblnChange As Boolean

Private Sub cbo交易类型_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub
Private Sub txtEdit_Change(Index As Integer)
    If Index = 0 And mblnChange = False Then
        g病人身份_兴安.个人编号 = ""
        g病人身份_兴安.卡号 = ""
    End If
End Sub

Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim strCurrDate As String
    Dim rsTemp As New ADODB.Recordset
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    g病人身份_兴安.byt类型 = mbytType
    mblnChange = True
    If Index = 0 Then
        SetOKCtrl False
        mblnChange = True
        
        If txtEdit(0).Text = "" Then
            ShowMsgbox "请输入卡号"
            txtEdit(0).SetFocus
            Exit Sub
        End If
        
        g病人身份_兴安.卡号 = Mid(txtEdit(0).Text, 4, 16)
        If 身份鉴别_兴安 = False Then Exit Sub
        Call LoadCtrlData
        txtEdit(0).Text = g病人身份_兴安.卡号
        SetOKCtrl True
    Else
        If mbytType = 0 Then
        Else
            SetOKCtrl False
            mblnChange = True
            g病人身份_兴安.卡号 = txtEdit(1)
            g病人身份_兴安.个人编号 = txtEdit(1)
            
            If 身份鉴别_兴安 = False Then Exit Sub
            Call LoadCtrlData
            SetOKCtrl True
        End If
    End If
    SetCboListIndex
    zlCommFun.PressKey vbKeyTab
End Sub

Private Sub SetOKCtrl(ByVal blnEn As Boolean)
    cmd确定.Enabled = blnEn
End Sub

Private Function IsValid() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:验证数据的合法性
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    IsValid = False
    If mbytType = 0 Then
        If Trim(txtEdit(0).Text) = "" Then
            MsgBox "还没有输入医保卡号！", vbInformation, gstrSysName
            txtEdit(0).SetFocus
            Exit Function
        End If
        If Trim(txtEdit(1).Text) <> g病人身份_兴安.个人编号 Then
            ShowMsgbox "医保证号不正确定,请检查"
            txtEdit(1).SetFocus
            Exit Function
        End If
        If Trim(g病人身份_兴安.姓名) = "" Then
            MsgBox "还没进行身份验证！", vbInformation, gstrSysName
            txtEdit(0).SetFocus
            Exit Function
        End If
    
    Else
        If Trim(txtEdit(1).Text) = "" Then
            MsgBox "还没有输入医保证号！", vbInformation, gstrSysName
            If txtEdit(1).Enabled Then txtEdit(1).SetFocus
            Exit Function
        End If
        If Trim(g病人身份_兴安.姓名) = "" Then
            MsgBox "还没进行身份验证！", vbInformation, gstrSysName
            If txtEdit(1).Enabled Then txtEdit(1).SetFocus
            Exit Function
        End If
    
    End If
    
    If cbo交易类型.Text = "" Then
        ShowMsgbox "交易类别未选择"
        Exit Function
    End If
    
    If mbytType <> 2 Then
        If mbytType = 4 Then
            '不检查当前着态
        Else
            '检查病人状态
            gstrSQL = "select nvl(当前状态,0) as 状态 from 保险帐户 where 险类=[1] and 医保号=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_兴安, g病人身份_兴安.个人编号)
            If rsTemp.RecordCount > 0 Then
                If rsTemp("状态") > 0 Then
                    MsgBox "该病人已经在院，不能通过身份验证。", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
        If mbytType = 0 Or mbytType = 3 Then
            '设置
        End If
    Else
        '不区分门诊和住院的，只是刷卡显示一下内容而已，不保存
        Unload Me
        Exit Function
    End If
    IsValid = True
End Function

Private Sub cmd取消_Click()
    Unload Me
End Sub

Private Sub cmd确定_Click()
    Dim lng疾病ID As Long
    
    Dim strIdentify As String, strAddition As String
    Dim rsTemp As New ADODB.Recordset
    Dim str类别 As String
    Dim int当前状态 As Integer
    
    If IsValid = False Then Exit Sub
    If mbytType = 0 Then
        Dim StrInput As String, strOutput As String
        StrInput = InitInfor_兴安.医院编码 & vbTab
        StrInput = StrInput & UserInfo.编号 & vbTab
        StrInput = StrInput & UserInfo.姓名 & vbTab
        StrInput = StrInput & cbo交易类型.ItemData(cbo交易类型.ListIndex)
        If 业务请求_兴安(操作员注册, StrInput, strOutput) = False Then Exit Sub
        
        g病人身份_兴安.门诊流水号 = Split(strOutput, vbTab)(0)
    End If
    
    g病人身份_兴安.交易类型 = cbo交易类型.Text
    
    int当前状态 = 0
    
    If mbytType = 4 Then
        '需确定当前状态,因为当前状态是不能改变的
        gstrSQL = "Select * from 保险帐户 where 险类=" & TYPE_兴安 & " and  医保号='" & g病人身份_兴安.个人编号 & "'"
        
        zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
        If Not rsTemp.EOF Then
            mlng病人ID = Nvl(rsTemp!病人ID, 0)
            int当前状态 = Nvl(rsTemp!当前状态, 0)
        End If
        rsTemp.Close
    End If
    
    '建立病人档案信息，传入格式：
    '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);
    '8.中心代码;9.顺序号;10人员身份;11帐户余额;12当前状态;13病种ID;14在职(0,1);15退休证号;16年龄段;17灰度级
    '18帐户增加累计,19帐户支出累计,20进入统筹累计,21统筹报销累计,22住院次数累计
    With g病人身份_兴安
        
        strIdentify = .卡号                               '0卡号
        strIdentify = strIdentify & ";" & .个人编号           '1医保号
        strIdentify = strIdentify & ";" & ""                 '2密码
        strIdentify = strIdentify & ";" & .姓名               '3姓名
        strIdentify = strIdentify & ";" & .性别                 '4性别
        strIdentify = strIdentify & ";" & ""                    '5出生日期
        strIdentify = strIdentify & ";" & ""                      '6身份证
        strIdentify = strIdentify & ";" & .单位名称                 '7.单位名称(编码)
        strAddition = ";0"                                          '8.中心代码
        strAddition = strAddition & ";" & .住院登记号                            '9.顺序号
        strAddition = strAddition & ";" & .人员类别                 '10人员身份
        strAddition = strAddition & ";" & .帐户余额                 '11帐户余额
        
        strAddition = strAddition & ";" & int当前状态                            '12当前状态
        strAddition = strAddition & ";"                             '13病种ID
        strAddition = strAddition & ";1"                            '14在职(1,2,3)
        strAddition = strAddition & ";" & IIf(.病种代码 = "", "", .病种代码 & "-" & .病种名称)                          '15退休证号
        strAddition = strAddition & ";" & .年龄                     '16年龄段
        strAddition = strAddition & ";"                             '17灰度级
        strAddition = strAddition & ";" & .帐户余额                 '18帐户增加累计
        strAddition = strAddition & ";0"                            '19帐户支出累计
        strAddition = strAddition & ";0"                            '20上年工资总额
        strAddition = strAddition & ";"                             '21住院次数累计
    End With
    
    mlng病人ID = BuildPatiInfo(0, strIdentify & strAddition, mlng病人ID, TYPE_兴安)
    
    '保险帐户:增加字段:当日已用药品,当年已用自负段,当年已用慢病,本年已用基本统筹,本年已用大病统筹,当年住院次数,本次住院起付标准
    If mbytType = 0 Then
        '门诊:
        '更新保险帐户的相关信息
        gstrSQL = "zl_保险帐户_更新信息(" & mlng病人ID & "," & TYPE_兴安 & ",'当日已用药品','" & g病人身份_兴安.当日已用药品金额 & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存当日已用药品金额")
        gstrSQL = "zl_保险帐户_更新信息(" & mlng病人ID & "," & TYPE_兴安 & ",'当年已用自负段','" & g病人身份_兴安.当年已用自负段 & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存当年已用自负段")
        gstrSQL = "zl_保险帐户_更新信息(" & mlng病人ID & "," & TYPE_兴安 & ",'当年已用慢病','" & g病人身份_兴安.当年已用慢病统筹 & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存当年已用慢病统筹")
        gstrSQL = "zl_保险帐户_更新信息(" & mlng病人ID & "," & TYPE_兴安 & ",'门诊流水号','''" & g病人身份_兴安.门诊流水号 & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存门诊流水号")
    ElseIf mbytType = 1 Then
        gstrSQL = "zl_保险帐户_更新信息(" & mlng病人ID & "," & TYPE_兴安 & ",'本年已用基本统筹','" & g病人身份_兴安.本年已用基本统筹 & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存本年已用基本统筹")
        gstrSQL = "zl_保险帐户_更新信息(" & mlng病人ID & "," & TYPE_兴安 & ",'本年已用大病统筹','" & g病人身份_兴安.本年已用大病统筹 & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存本年已用大病统筹")
        gstrSQL = "zl_保险帐户_更新信息(" & mlng病人ID & "," & TYPE_兴安 & ",'当年住院次数','" & g病人身份_兴安.当年第几次住院 & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存当年第几次住院")
        gstrSQL = "zl_保险帐户_更新信息(" & mlng病人ID & "," & TYPE_兴安 & ",'本次住院起付标准','" & g病人身份_兴安.本次住院起付标准 & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存本次住院起付标准")
    End If
    '返回格式:中间插入病人ID
    If mlng病人ID > 0 Then
        mstrReturn = strIdentify & ";" & mlng病人ID & strAddition
    End If
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub
Public Function GetPatient(Optional bytType As Byte, Optional lng病人ID As Long = 0) As String
    mbytType = bytType
    mlng病人ID = lng病人ID
    mstrReturn = ""
    DebugTool "进入身份验证,并开始加入基本信息"
    
    If LoadBaseData = False Then
        DebugTool "加入失败(身份验证)"
        Exit Function
    End If
    DebugTool "加入成功(身份验证)"
    
    Me.Show 1
    lng病人ID = mlng病人ID
    GetPatient = mstrReturn
End Function
Private Function LoadBaseData() As Boolean
    '加载基础数据
    Dim rsTemp As New ADODB.Recordset
    LoadBaseData = False
    On Error GoTo errHand:
    
    If mbytType = 0 Then
        cbo交易类型.AddItem "普通医保门诊"
        cbo交易类型.ListIndex = cbo交易类型.NewIndex
        cbo交易类型.ItemData(cbo交易类型.NewIndex) = 1
        cbo交易类型.AddItem "特殊医保门诊"
        cbo交易类型.ItemData(cbo交易类型.NewIndex) = 2
        txtEdit(0).Enabled = True
        txtEdit(1).Enabled = True
    Else
        cbo交易类型.AddItem "医保住院"
        cbo交易类型.ListIndex = cbo交易类型.NewIndex
        txtEdit(0).Enabled = False
        txtEdit(1).Enabled = True
        cbo交易类型.Enabled = False
        lbl(12).Caption = "本次起付标准"
        lbl(13).Caption = "已用基本统筹"
        lbl(14).Caption = "已用大病统筹"
    End If
    
    LoadBaseData = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function
Private Sub LoadCtrlData()
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:填充数据
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    With g病人身份_兴安
       ' txtEdit(1).Text = .个人编号
        lblEdit(2).Caption = .姓名
        lblEdit(3).Caption = .性别
        lblEdit(4).Caption = .年龄
        lblEdit(5).Caption = .人员类别
        lblEdit(6).Caption = .帐户状态
        lblEdit(7).Caption = .帐户余额
        lblEdit(8).Caption = .病种代码
        lblEdit(9).Caption = .病种名称
        lblEdit(10).Caption = .单位名称
        
        lblEdit(15).Caption = .当年第几次住院
        
        If mbytType = 0 Then
            lblEdit(12).Caption = .当日已用药品金额
            lblEdit(13).Caption = .当年已用自负段
            lblEdit(14).Caption = .当年已用慢病统筹
        ElseIf mbytType = 1 Then
            lblEdit(12).Caption = .本次住院起付标准
            lblEdit(13).Caption = .本年已用基本统筹
            lblEdit(14).Caption = .本年已用大病统筹
        End If
    End With
End Sub
Private Sub SetCboListIndex()
    '设置控件属性
    Dim i As Long
    If mbytType <> 0 Then Exit Sub
    If InStr(1, "离休二乙", g病人身份_兴安.人员类别) <> 0 Then
       For i = 0 To cbo交易类型.ListCount - 1
            If cbo交易类型.ItemData(i) = 2 Then
                cbo交易类型.ListIndex = i
            End If
       Next
       cbo交易类型.Enabled = False
    Else
       cbo交易类型.Enabled = True
    End If
    
End Sub
