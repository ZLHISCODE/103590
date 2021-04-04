VERSION 5.00
Begin VB.Form frmIdentify渝北农医 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "身份验证"
   ClientHeight    =   4635
   ClientLeft      =   6615
   ClientTop       =   5505
   ClientWidth     =   5880
   Icon            =   "frmIdentify渝北农医.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txt家庭帐户余额 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1650
      TabIndex        =   15
      Top             =   3030
      Width           =   2355
   End
   Begin VB.ComboBox cbo医疗类别 
      Height          =   300
      Left            =   1650
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   3420
      Width           =   2385
   End
   Begin VB.TextBox txt疾病信息 
      Height          =   300
      Left            =   1650
      TabIndex        =   19
      Top             =   3810
      Width           =   2355
   End
   Begin VB.TextBox txt并发症 
      Height          =   300
      Left            =   1650
      TabIndex        =   21
      Top             =   4200
      Width           =   2355
   End
   Begin VB.CommandButton cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4590
      TabIndex        =   24
      Top             =   840
      Width           =   1100
   End
   Begin VB.CommandButton cmd确定 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4590
      TabIndex        =   23
      Top             =   360
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   4515
      Left            =   4290
      TabIndex        =   22
      Top             =   -30
      Width           =   30
   End
   Begin VB.TextBox txt个人帐户余额 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1650
      TabIndex        =   13
      Top             =   2640
      Width           =   2355
   End
   Begin VB.TextBox txt身份证号 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1650
      TabIndex        =   11
      Top             =   2250
      Width           =   2355
   End
   Begin VB.TextBox txt出生日期 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1650
      TabIndex        =   9
      Top             =   1860
      Width           =   1365
   End
   Begin VB.TextBox txt性别 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1650
      TabIndex        =   7
      Top             =   1470
      Width           =   855
   End
   Begin VB.TextBox txt姓名 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1650
      TabIndex        =   5
      Top             =   1080
      Width           =   2355
   End
   Begin VB.TextBox txt个人编码 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1650
      TabIndex        =   3
      Top             =   690
      Width           =   2355
   End
   Begin VB.TextBox txt医疗证号 
      Height          =   300
      Left            =   1650
      MaxLength       =   25
      TabIndex        =   1
      Top             =   300
      Width           =   2355
   End
   Begin VB.Label lbl家庭帐户余额 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "家庭帐户余额(&F)"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   210
      TabIndex        =   14
      Top             =   3090
      Width           =   1350
   End
   Begin VB.Label lbl医疗类别 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "医疗类别(&T)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   570
      TabIndex        =   16
      Top             =   3480
      Width           =   990
   End
   Begin VB.Label lbl疾病信息 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "疾病信息(&D)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   570
      TabIndex        =   18
      Top             =   3870
      Width           =   990
   End
   Begin VB.Label lbl并发症 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "并发症(&Z)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   750
      TabIndex        =   20
      Top             =   4260
      Width           =   810
   End
   Begin VB.Label lbl人帐余额 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "个人帐户余额(&Y)"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   210
      TabIndex        =   12
      Top             =   2700
      Width           =   1350
   End
   Begin VB.Label lbl身份证号 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "身份证号(&I)"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   570
      TabIndex        =   10
      Top             =   2310
      Width           =   990
   End
   Begin VB.Label lbl出生日期 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "出生日期(&B)"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   570
      TabIndex        =   8
      Top             =   1920
      Width           =   990
   End
   Begin VB.Label lbl性别 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "性别(&S)"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   930
      TabIndex        =   6
      Top             =   1530
      Width           =   630
   End
   Begin VB.Label lbl姓名 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "姓名(&N)"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   930
      TabIndex        =   4
      Top             =   1140
      Width           =   630
   End
   Begin VB.Label lbl个人编码 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "个人编码(&P)"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   570
      TabIndex        =   2
      Top             =   750
      Width           =   990
   End
   Begin VB.Label lbl医疗证号 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "医疗证号(&L)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   570
      TabIndex        =   0
      Top             =   360
      Width           =   990
   End
End
Attribute VB_Name = "frmIdentify渝北农医"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytType As Byte            '0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号
Private mlng病人ID As Long
Private mstrReturn As String

Private Sub cbo医疗类别_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txt并发症_GotFocus()
    Call zlControl.TxtSelAll(txt并发症)
End Sub

Private Sub txt并发症_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txt疾病信息_GotFocus()
    Call zlControl.TxtSelAll(txt疾病信息)
End Sub

Private Sub txt疾病信息_KeyPress(KeyAscii As Integer)
    Dim strLike As String
    Dim StrInput As String
    Dim str性别 As String
    Dim blnCancel As Boolean
    Dim rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    If KeyAscii <> vbKeyReturn Then Exit Sub
    
    If txt疾病信息.Text = lbl疾病信息.Tag And txt疾病信息.Text <> "" Then
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf txt疾病信息.Text = "" Then
        txt疾病信息.Tag = "": lbl疾病信息.Tag = ""
        Call zlCommFun.PressKey(vbKeyTab) '允许不输入
    Else
        strLike = IIf(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = 0, "%", "")
        StrInput = UCase(txt疾病信息.Text)
        str性别 = txt性别.Text
        If str性别 = "男" Then
            str性别 = " And (A.性别限制='男' Or A.性别限制 is NULL)"
        ElseIf str性别 = "女" Then
            str性别 = " And (A.性别限制='女' Or A.性别限制 is NULL)"
        Else
            str性别 = ""
        End If
        gstrSQL = "Select A.ID,A.编码,A.附码,A.名称,A.简码,A.说明,A.性别限制,B.类别" & _
            " From 疾病编码目录 A,疾病编码类别 B" & _
            " Where A.类别=B.编码 And A.类别 Not IN('B','Z')" & _
            " And (A.编码 Like '" & StrInput & "%'" & _
            " Or Upper(A.名称) Like '" & strLike & StrInput & "%'" & _
            " Or Upper(A.简码) Like '" & strLike & StrInput & "%'" & _
            " Or Upper(A.附码) Like '" & strLike & StrInput & "%')" & _
            " And Rownum<=100" & str性别 & _
            " Order by A.类别,A.编码"
        Set rsTemp = zlDatabase.ShowSelect(Me, gstrSQL, 0, "疾病编码Input", , , , , , True, _
            txt疾病信息.Left + Me.Left, _
            txt疾病信息.Top + Me.Top, txt疾病信息.Height, blnCancel, , True)
        If Not rsTemp Is Nothing Then
            txt疾病信息.Tag = rsTemp!ID
            txt疾病信息.Text = "(" & rsTemp!编码 & ")" & rsTemp!名称
            lbl疾病信息.Tag = txt疾病信息.Text '用于恢复显示
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            If Not blnCancel Then
                MsgBox "没有找到匹配的疾病编码。", vbInformation, gstrSysName
            End If
            If lbl疾病信息.Tag <> "" Then txt疾病信息.Text = lbl疾病信息.Tag
            txt疾病信息.SetFocus
        End If
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmd取消_Click()
    Unload Me
End Sub

Private Sub cmd确定_Click()
    Dim lng疾病ID As Long
    Dim str出生日期 As String
    Dim strIdentify As String, strAddition As String
    Dim rsTemp As New ADODB.Recordset
    
    If Trim(txt医疗证号.Text) = "" Then
        MsgBox "还没有输入医疗证号！", vbInformation, gstrSysName
        txt医疗证号.SetFocus
        Exit Sub
    End If
    If Trim(txt姓名.Text) = "" Then
        MsgBox "还没有获取该医保病人的身份信息，不能通过验证！", vbInformation, gstrSysName
        txt医疗证号.SetFocus
        Exit Sub
    End If
    If mbytType <> 3 Then
        If txt疾病信息.Tag = "" Then
            MsgBox "请输入病人的疾病信息！", vbInformation, gstrSysName
            txt疾病信息.SetFocus
            Exit Sub
        End If
    
        '获取病种ID
        gstrSQL = "Select ID From 疾病编码目录 Where 编码=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取保险病种", CStr(txt疾病信息.Tag))
        If Not rsTemp.EOF Then
            lng疾病ID = rsTemp!ID
        End If
    End If
    
    If mbytType <> 2 Then
        '检查病人状态
        gstrSQL = "select nvl(当前状态,0) as 状态 from 保险帐户 where 险类=[1] and 医保号=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_渝北农医, CStr(txt医疗证号.Text))
        If rsTemp.RecordCount > 0 Then
            If rsTemp("状态") > 0 Then
                MsgBox "该病人已经在院，不能通过身份验证。", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    Else
        '不区分门诊和住院的，只是刷卡显示一下内容而已，不保存
        Unload Me
        Exit Sub
    End If
    
    '建立病人档案信息，传入格式：
    '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);
    '8.中心代码;9.顺序号;10人员身份;11帐户余额;12当前状态;13病种ID;14在职(0,1);15退休证号;16年龄段;17灰度级
    '18帐户增加累计,19帐户支出累计,20进入统筹累计,21统筹报销累计,22住院次数累计
    strIdentify = txt医疗证号.Text                              '0卡号
    strIdentify = strIdentify & ";" & txt个人编码.Text          '1医保号
    strIdentify = strIdentify & ";"                             '2密码
    strIdentify = strIdentify & ";" & txt姓名.Text              '3姓名
    strIdentify = strIdentify & ";" & txt性别.Text              '4性别
    strIdentify = strIdentify & ";" & txt出生日期.Text          '5出生日期
    strIdentify = strIdentify & ";" & txt身份证号.Text          '6身份证
    strIdentify = strIdentify & ";"                             '7.单位名称(编码)
    strAddition = ";0"                                          '8.中心代码
    strAddition = strAddition & ";"                             '9.顺序号
    strAddition = strAddition & ";"                             '10人员身份
    strAddition = strAddition & ";" & Val(txt个人帐户余额.Text)     '11帐户余额
    strAddition = strAddition & ";0"                            '12当前状态
    strAddition = strAddition & ";" & lng疾病ID                 '13病种ID
    strAddition = strAddition & ";1"                            '14在职(1,2,3)
    strAddition = strAddition & ";"                             '15退休证号
    strAddition = strAddition & ";"                             '16年龄段
    strAddition = strAddition & ";"                             '17灰度级
    strAddition = strAddition & ";" & Val(txt个人帐户余额.Text)     '18帐户增加累计
    strAddition = strAddition & ";0"                            '19帐户支出累计
    strAddition = strAddition & ";0"                            '20上年工资总额
    strAddition = strAddition & ";"                             '21住院次数累计
    
    mlng病人ID = BuildPatiInfo(0, strIdentify & strAddition, mlng病人ID, TYPE_渝北农医)
    '返回格式:中间插入病人ID
    If mlng病人ID > 0 Then
        mstrReturn = strIdentify & ";" & mlng病人ID & strAddition
    End If
    
    '更新家庭帐户余额
    gstrSQL = "zl_保险帐户_更新信息(" & mlng病人ID & "," & TYPE_渝北农医 & ",'家庭帐户余额','''" & Val(txt家庭帐户余额.Text) & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "更新家庭帐户余额")
    
    With gComInfo_渝北农医
        .医疗证号 = txt医疗证号.Text
        .个人编号 = txt个人编码.Text
        .疾病编码 = txt疾病信息.Tag
        .并发症 = txt并发症.Text
        .业务类型 = Me.cbo医疗类别.ItemData(Me.cbo医疗类别.ListIndex)
    End With
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
'      11  医疗类别    普通门诊
'      12  医疗类别    特殊门诊
'      13  医疗类别    急诊抢救
'      21  医疗类别    普通住院
'      22  医疗类别    转入医院住院
    With Me.cbo医疗类别
        .Clear
        '加入门诊业务
        If mbytType = 0 Or mbytType = 2 Or mbytType = 3 Then
            .AddItem "普通门诊"
            .ItemData(.NewIndex) = 11
            .AddItem "特殊门诊"
            .ItemData(.NewIndex) = 12
            .AddItem "急诊抢救"
            .ItemData(.NewIndex) = 13
        End If
        '加入住院业务
        If mbytType = 1 Or mbytType = 2 Then
            .AddItem "普通住院"
            .ItemData(.NewIndex) = 21
            .AddItem "转入医院住院"
            .ItemData(.NewIndex) = 22
        End If
        .ListIndex = 0
    End With
    
    '挂号可以不必输入疾病与并发症信息
    If mbytType = 3 Then
        txt疾病信息.Enabled = False
        txt并发症.Enabled = False
    End If
End Sub

Private Sub txt医疗证号_GotFocus()
    Call zlControl.TxtSelAll(txt医疗证号)
End Sub

Public Function GetPatient(Optional bytType As Byte, Optional lng病人ID As Long = 0) As String
    mbytType = bytType
    mlng病人ID = lng病人ID
    mstrReturn = ""
    Me.Show 1
    lng病人ID = mlng病人ID
    GetPatient = mstrReturn
End Function

Private Sub txt医疗证号_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim StrInput As String
    Dim arrOutput
    If KeyCode <> vbKeyReturn Then Exit Sub
    StrInput = Trim(txt医疗证号.Text)
    If Trim(StrInput) = "" Then Exit Sub
    
    Call 调用接口_准备_渝北农医(读取个人信息_渝北农医, StrInput)
    If Not 调用接口_渝北农医 Then Exit Sub
    
    arrOutput = Split(gstrOutput_渝北农医, gstrSplit_Col_重庆银海版)
    If Val(arrOutput(7)) = 1 Then
        MsgBox "该卡已被锁定，不允许再次办理就诊登记！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    txt个人编码.Text = arrOutput(1)
    txt姓名.Text = arrOutput(2)
    txt性别.Text = IIf(Val(arrOutput(3)) = 0, "男", "女")
    txt出生日期.Text = Format(arrOutput(4), "yyyy-MM-dd")
    txt身份证号.Text = arrOutput(5)
    txt个人帐户余额.Text = Format(arrOutput(6), "#0.00")
    txt家庭帐户余额.Text = Format(arrOutput(12), "#0.00")
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub
