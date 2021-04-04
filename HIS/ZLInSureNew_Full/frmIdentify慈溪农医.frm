VERSION 5.00
Object = "{E5918DE2-9E1A-472B-96C6-5AE5994F9138}#1.0#0"; "ReadBarComm.dll"
Begin VB.Form frmIdentify慈溪农医 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "身份验证"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8445
   Icon            =   "frmIdentify慈溪农医.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   8445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chk自输卡号 
      Caption         =   "手工输入卡号"
      Height          =   375
      Left            =   1200
      TabIndex        =   40
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CheckBox Chk事后结算 
      Caption         =   "事后结算"
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
      Height          =   255
      Left            =   1440
      TabIndex        =   39
      Top             =   240
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CheckBox chk小儿用药 
      Alignment       =   1  'Right Justify
      Caption         =   "小儿用药"
      Height          =   255
      Left            =   600
      TabIndex        =   34
      Top             =   4170
      Width           =   1335
   End
   Begin VB.TextBox txt承担比例 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   300
      Left            =   6090
      MaxLength       =   3
      TabIndex        =   31
      Top             =   3390
      Width           =   2025
   End
   Begin VB.ComboBox cbo医疗类别 
      Height          =   300
      Left            =   1410
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   3390
      Width           =   2025
   End
   Begin VB.TextBox txt疾病信息 
      Height          =   300
      Left            =   1410
      TabIndex        =   33
      Top             =   3780
      Width           =   6735
   End
   Begin VB.CommandButton cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6960
      TabIndex        =   37
      Top             =   4680
      Width           =   1100
   End
   Begin VB.CommandButton cmd确定 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5700
      TabIndex        =   36
      Top             =   4680
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   30
      Left            =   30
      TabIndex        =   35
      Top             =   4500
      Width           =   8835
   End
   Begin VB.TextBox txt卡状态 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1410
      TabIndex        =   13
      Top             =   3000
      Width           =   2025
   End
   Begin VB.TextBox txt病种名称 
      Enabled         =   0   'False
      Height          =   300
      Left            =   6090
      TabIndex        =   17
      Top             =   1050
      Width           =   2025
   End
   Begin VB.TextBox txt帐户余额 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   300
      Left            =   6090
      TabIndex        =   27
      Top             =   3000
      Width           =   2025
   End
   Begin VB.TextBox txt帐户累计支付 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   300
      Left            =   6090
      TabIndex        =   25
      Top             =   2610
      Width           =   2025
   End
   Begin VB.TextBox txt住院累计实报 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   300
      Left            =   6090
      TabIndex        =   23
      Top             =   2220
      Width           =   2025
   End
   Begin VB.TextBox txt特殊门诊累计报销 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   300
      Left            =   6090
      TabIndex        =   21
      Top             =   1830
      Width           =   2025
   End
   Begin VB.TextBox txt门诊累计报销 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   300
      Left            =   6090
      TabIndex        =   19
      Top             =   1440
      Width           =   2025
   End
   Begin VB.TextBox txt病种代码 
      Enabled         =   0   'False
      Height          =   300
      Left            =   6090
      TabIndex        =   15
      Top             =   660
      Width           =   1395
   End
   Begin VB.TextBox txt出生日期 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1410
      TabIndex        =   11
      Top             =   2610
      Width           =   2025
   End
   Begin VB.TextBox txt身份证号 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1410
      TabIndex        =   9
      Top             =   2220
      Width           =   2025
   End
   Begin VB.TextBox txt性别 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1410
      TabIndex        =   7
      Top             =   1830
      Width           =   555
   End
   Begin VB.TextBox txt姓名 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1410
      TabIndex        =   5
      Top             =   1440
      Width           =   1185
   End
   Begin VB.TextBox txt医疗证号 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1410
      TabIndex        =   3
      Top             =   1050
      Width           =   1575
   End
   Begin VB.TextBox txt卡证号码 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1410
      MaxLength       =   30
      TabIndex        =   1
      Top             =   660
      Width           =   2835
   End
   Begin READBARCOMMLibCtl.ReadBar2Comm ReadCard 
      Height          =   375
      Left            =   1320
      OleObjectBlob   =   "frmIdentify慈溪农医.frx":000C
      TabIndex        =   38
      Top             =   5400
      Width           =   5175
   End
   Begin VB.Label lbl承担比例 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "承担比例"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5280
      TabIndex        =   30
      Top             =   3450
      Width           =   720
   End
   Begin VB.Label lbl医疗类别 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "医疗类别"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   600
      TabIndex        =   28
      Top             =   3450
      Width           =   720
   End
   Begin VB.Label lbl疾病信息 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "疾病信息"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   600
      TabIndex        =   32
      Top             =   3840
      Width           =   720
   End
   Begin VB.Label lbl卡状态 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "卡状态"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   780
      TabIndex        =   12
      Top             =   3060
      Width           =   540
   End
   Begin VB.Label lbl病种名称 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "特殊病种名称"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   4920
      TabIndex        =   16
      Top             =   1110
      Width           =   1080
   End
   Begin VB.Label lbl帐户余额 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "帐户余额"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5280
      TabIndex        =   26
      Top             =   3060
      Width           =   720
   End
   Begin VB.Label lbl帐户累计支付 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "帐户累计支付"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   4920
      TabIndex        =   24
      Top             =   2670
      Width           =   1080
   End
   Begin VB.Label lbl住院累计实报 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "住院累计实报"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   4920
      TabIndex        =   22
      Top             =   2280
      Width           =   1080
   End
   Begin VB.Label lbl特殊门诊累计报销 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "特殊门诊累计报销"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   4560
      TabIndex        =   20
      Top             =   1890
      Width           =   1440
   End
   Begin VB.Label lbl门诊累计报销 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "门诊累计报销"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   4920
      TabIndex        =   18
      Top             =   1500
      Width           =   1080
   End
   Begin VB.Label lbl病种代码 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "特殊病种代码"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   4920
      TabIndex        =   14
      Top             =   720
      Width           =   1080
   End
   Begin VB.Label lbl出生日期 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "出生日期"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   600
      TabIndex        =   10
      Top             =   2670
      Width           =   720
   End
   Begin VB.Label lbl身份证号 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "身份证号"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   600
      TabIndex        =   8
      Top             =   2280
      Width           =   720
   End
   Begin VB.Label lbl性别 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "性别"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   960
      TabIndex        =   6
      Top             =   1890
      Width           =   360
   End
   Begin VB.Label lbl姓名 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "姓名"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   960
      TabIndex        =   4
      Top             =   1500
      Width           =   360
   End
   Begin VB.Label lbl医疗证号 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "医疗证号"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   600
      TabIndex        =   2
      Top             =   1110
      Width           =   720
   End
   Begin VB.Label lbl卡证号码 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "卡证号码"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   720
   End
End
Attribute VB_Name = "frmIdentify慈溪农医"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytType As Byte            '0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号
Private mlng病人ID As Long
Private mstrReturn As String

Private Sub cbo医疗类别_Click()
    Me.txt承担比例.Enabled = False
    If cbo医疗类别.ItemData(cbo医疗类别.ListIndex) = 22 Then
        '交通事故
        Me.txt承担比例.Enabled = True
        Me.txt承担比例.SetFocus
    End If
End Sub

Private Sub cbo医疗类别_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Chk事后结算_Click()
Dim strDate As String
strDate = Format(zlDatabase.Currentdate, "yymmddhhmmss")
    If Chk事后结算.Value = 1 Then
       txt卡证号码.Text = UserInfo.用户名 + strDate
       txt医疗证号.Text = "事后补报" + strDate
       txt姓名.Enabled = True
       txt性别.Enabled = True
       txt身份证号.Enabled = True
       txt出生日期.Enabled = True
       txt门诊累计报销.Text = 0
       txt特殊门诊累计报销.Text = 0
       txt帐户余额.Text = 0
      chk自输卡号.Enabled = False
     Else
        txt卡证号码.Text = ""
       txt医疗证号.Text = ""
       txt姓名.Enabled = False
       txt性别.Enabled = False
       txt身份证号.Enabled = False
       txt出生日期.Enabled = False
       txt门诊累计报销.Text = 0
       txt特殊门诊累计报销.Text = 0
       txt帐户余额.Text = 0
        chk自输卡号.Enabled = True
    End If
End Sub

Private Sub cmd读卡_Click()

End Sub

Private Sub chk自输卡号_Click()
If chk自输卡号.Value = 1 Then
   txt卡证号码.Enabled = True
Else
   txt卡证号码.Enabled = False
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
    
    If Trim(txt卡证号码.Text) = "" Then
        MsgBox "请读卡！", vbInformation, gstrSysName
        txt卡证号码.SetFocus
        Exit Sub
    End If
    If Trim(txt姓名.Text) = "" Then
        MsgBox "还没有获取该医保病人的身份信息，不能通过验证！", vbInformation, gstrSysName
        txt卡证号码.SetFocus
        Exit Sub
    End If
    If Me.cbo医疗类别.ItemData(Me.cbo医疗类别.ListIndex) = 22 Then
        If Val(txt承担比例.Text) < 0 Then
            MsgBox "承担比例不能小于零！", vbInformation, gstrSysName
            txt承担比例.SetFocus
            Exit Sub
        End If
        If Val(txt承担比例.Text) > 100 Then
            MsgBox "承担比例不能大于一百！", vbInformation, gstrSysName
            txt承担比例.SetFocus
            Exit Sub
        End If
    End If
  ' txt疾病信息.Tag = 124
    If mbytType <> 3 Then
     
    
       If Val(txt疾病信息.Tag) = 0 Then
           ' MsgBox "请输入病人的疾病信息！", vbInformation, gstrSysName
           'txt疾病信息.SetFocus
         ' Exit Sub
          txt疾病信息.Tag = 999
      End If
    
        
    End If
    
    If mbytType <> 2 Then
        '检查病人状态
        gstrSQL = "select nvl(当前状态,0) as 状态 from 保险帐户 where 险类=[1] and 医保号=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_慈溪农医, txt医疗证号)
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
    strIdentify = txt卡证号码.Text                              '0卡号
    strIdentify = strIdentify & ";" & txt医疗证号.Text          '1医保号
    strIdentify = strIdentify & ";"                             '2密码
    strIdentify = strIdentify & ";" & txt姓名.Text              '3姓名
    strIdentify = strIdentify & ";" & txt性别.Text              '4性别
    strIdentify = strIdentify & ";" & txt出生日期.Text          '5出生日期
    strIdentify = strIdentify & ";" & txt身份证号.Text          '6身份证
    strIdentify = strIdentify & ";"                             '7.单位名称(编码)
    strAddition = ";0"                                          '8.中心代码
    strAddition = strAddition & ";"                             '9.顺序号
    strAddition = strAddition & ";" '10人员身份
 
  strAddition = strAddition & ";" & Val(txt帐户余额.Text)     '11帐户余额
    strAddition = strAddition & ";0"                            '12当前状态
    strAddition = strAddition & ";" & Val(txt疾病信息.Tag)                 '13病种ID
    strAddition = strAddition & ";1"                            '14在职(1,2,3)
    strAddition = strAddition & ";" & Val(txt承担比例.Text)     '15退休证号
    strAddition = strAddition & ";"                             '16年龄段
    strAddition = strAddition & ";"                             '17灰度级
    strAddition = strAddition & ";" & Val(txt帐户余额.Text) + Val(txt帐户累计支付.Text)   '18帐户增加累计
    strAddition = strAddition & ";" & Val(txt帐户累计支付.Text)                           '19帐户支出累计
    strAddition = strAddition & ";0"                            '20上年工资总额
    strAddition = strAddition & ";"                             '21住院次数累计
    
    mlng病人ID = BuildPatiInfo(0, strIdentify & strAddition, mlng病人ID, TYPE_慈溪农医)
    '返回格式:中间插入病人ID
    If mlng病人ID > 0 Then
        mstrReturn = strIdentify & ";" & mlng病人ID & strAddition
    End If
    
    With gComInfo_慈溪农医
        .医疗证号 = txt医疗证号.Text
        .个人编号 = txt卡证号码.Text
        .业务类型 = Me.cbo医疗类别.ItemData(Me.cbo医疗类别.ListIndex)
    End With
    If Chk事后结算.Value = 1 Then
       gComInfo_慈溪农医.结算类型 = "事后补报"
    Else
       gComInfo_慈溪农医.结算类型 = "实时结算"
    End If
    
    '更新保险帐户相关信息（业务类型）
    gstrSQL = "zl_保险帐户_更新信息(" & mlng病人ID & "," & TYPE_慈溪农医 & ",'小儿用药','" & chk小儿用药.Value & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存业务类型")
    
    Unload Me

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim strMsg As String
    Dim IntPort As Integer
    Dim lngReturn As Long
    Dim rsTemp As New ADODB.Recordset
    
    
    With Me.cbo医疗类别
        .Clear
        '加入门诊业务
        If mbytType = 0 Or mbytType = 2 Or mbytType = 3 Then
             .AddItem "农保门诊"
            .ItemData(.NewIndex) = 11
            .AddItem "特殊门诊"
            .ItemData(.NewIndex) = 12
            Chk事后结算.Visible = True
            
        End If
        '加入住院业务
        If mbytType = 1 Or mbytType = 2 Then
            .AddItem "普通住院"
            .ItemData(.NewIndex) = 21
            .AddItem "交通事故"
            .ItemData(.NewIndex) = 22
            .AddItem "大病救助"
            .ItemData(.NewIndex) = 23
            .AddItem "难产"
            .ItemData(.NewIndex) = 24
            .AddItem "其他"
            .ItemData(.NewIndex) = 25
            
            chk小儿用药.Enabled = True
        End If
        .ListIndex = 0
    End With
    
    '挂号可以不必输入疾病与并发症信息
    If mbytType = 3 Then
        txt疾病信息.Enabled = False
    End If
    
    '取IC端口号
    gstrSQL = "Select 参数值 From 保险参数 Where 险类=[1] ANd 参数名='IC端口号'"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取IC端口号", TYPE_慈溪农医)
    If rsTemp.RecordCount = 0 Then
        IntPort = 1
    Else
        IntPort = Nvl(rsTemp!参数值, 1)
    End If
    
    '初始化读卡部件
    lngReturn = ReadCard.OpenPort(IntPort)
    Select Case lngReturn
    Case 0, 1
        strMsg = ""
    Case -1
        strMsg = "打开串口失败(一般由于端口号不存在或被占用)"
    Case -2
        strMsg = "获取串口状态失败"
    Case -3
        strMsg = "设置端口状态失败"
    Case -4
        strMsg = "创建工作者线程失败(系统资源不足)"
    Case -5
        strMsg = "内部错误"
    End Select
    If strMsg <> "" Then
        MsgBox strMsg, vbInformation, gstrSysName
        Exit Sub
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReadCard.ClosePort
End Sub

Private Sub ReadCard_OnComm(ByVal strData As String, ByVal lValidData As Long)
    Me.txt卡证号码.Text = Mid(strData, 1, 20)
    If Trim(txt卡证号码.Text) <> "" Then Call txt卡证号码_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub txt承担比例_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub txt卡证号码_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strReturn As String
    Dim str就诊卡号 As String   '03-25应军杰加
    Dim str农保卡号 As String   '03-25应军杰加
   
    
    Dim rsItem As New ADODB.Recordset '03-25应军杰加
    Dim arrReturn
    Const Returncode As Integer = 0         '如果返回0表示成功
    Const Returninfo As Integer = 1         '对应的错误提示
    Const Hzylhm As Integer = 2             '合作医疗号码
    Const Cyxm As Integer = 3               '成员姓名
    Const Cyxb As Integer = 4               '成员性别
    Const Sfzhm As Integer = 5              '身份证号码
    Const Jtdz As Integer = 6               '家庭地址
    Const Csrq As Integer = 7               '出生日期
    Const Kzt As Integer = 8                '卡状态
    Const Tsbzdm As Integer = 9             '特殊病种代码
    Const Tsbzmc As Integer = 10            '特殊病种名称
    Const Mzljsb As Integer = 11            '门诊累计实报
    Const Tsmzljsb As Integer = 12          '特殊门诊累计实报
    Const Zyljsb As Integer = 13            '住院累计实报
    Const Zhljzf As Integer = 14            '帐户累计支付
    Const Zhye As Integer = 15              '帐户余额
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    '03-25 应军杰加
   
    
    If Trim(txt卡证号码.Text) = "" Then Exit Sub
    Call 调用接口_准备_慈溪农医(gstrFunc慈溪农医_GetPersonalInfo, "Kzhm=" & txt卡证号码.Text)
    If Not 调用接口_慈溪农医 Then Exit Sub
    
    strReturn = gstrOutput_慈溪农医
    arrReturn = Split(strReturn, "&")
    Me.txt医疗证号.Text = Trim(Split(arrReturn(Hzylhm), "=")(1))
    Me.txt姓名.Text = Trim(Split(arrReturn(Cyxm), "=")(1))
    Me.txt性别.Text = Trim(Split(arrReturn(Cyxb), "=")(1))
    Me.txt身份证号.Text = Trim(Split(arrReturn(Sfzhm), "=")(1))
    Me.txt出生日期.Text = Format(Trim(Split(arrReturn(Csrq), "=")(1)), "yyyy-MM-dd")
    Me.txt病种代码.Text = Trim(Split(arrReturn(Tsbzdm), "=")(1))
    Me.txt病种名称.Text = Trim(Split(arrReturn(Tsbzmc), "=")(1))
    Me.txt门诊累计报销.Text = Trim(Split(arrReturn(Mzljsb), "=")(1))
    Me.txt特殊门诊累计报销.Text = Trim(Split(arrReturn(Tsmzljsb), "=")(1))
    Me.txt住院累计实报.Text = Trim(Split(arrReturn(Zyljsb), "=")(1))
    Me.txt帐户累计支付.Text = Trim(Split(arrReturn(Zhljzf), "=")(1))
    Me.txt帐户余额.Text = Trim(Split(arrReturn(Zhye), "=")(1))
    Me.txt卡状态.Text = Trim(Split(arrReturn(Kzt), "=")(1))
    Call zlCommFun.PressKey(vbKeyTab)
End Sub

Public Function GetPatient(Optional bytType As Byte, Optional lng病人ID As Long = 0) As String
    mbytType = bytType
    mlng病人ID = lng病人ID
    mstrReturn = ""
    Me.Show 1
    lng病人ID = mlng病人ID
    GetPatient = mstrReturn
End Function

