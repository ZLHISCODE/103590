VERSION 5.00
Begin VB.Form frmIdentify重庆银海版 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "身份验证"
   ClientHeight    =   5625
   ClientLeft      =   5310
   ClientTop       =   3135
   ClientWidth     =   7725
   Icon            =   "frm重庆银海版.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   7725
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd确定 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4920
      TabIndex        =   46
      Top             =   5160
      Width           =   1100
   End
   Begin VB.CommandButton cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6180
      TabIndex        =   47
      Top             =   5160
      Width           =   1100
   End
   Begin VB.Frame Frame2 
      Caption         =   "累计信息(元)"
      Height          =   1965
      Left            =   150
      TabIndex        =   29
      Top             =   3060
      Width           =   7425
      Begin VB.TextBox txt帐户余额 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1590
         TabIndex        =   31
         Tag             =   "13#"
         Top             =   300
         Width           =   1845
      End
      Begin VB.TextBox txt住院历史未补助 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5190
         TabIndex        =   45
         Tag             =   "25#"
         Top             =   1470
         Width           =   1845
      End
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1590
         TabIndex        =   43
         Tag             =   "24#"
         Top             =   1470
         Width           =   1845
      End
      Begin VB.TextBox txt特殊症状起付 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5190
         TabIndex        =   41
         Tag             =   "23#"
         Top             =   1080
         Width           =   1845
      End
      Begin VB.TextBox txt符合公务员门诊 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1590
         TabIndex        =   39
         Tag             =   "17#"
         Top             =   1080
         Width           =   1845
      End
      Begin VB.TextBox txt特病门诊医保费 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5190
         TabIndex        =   37
         Tag             =   "16#"
         Top             =   690
         Width           =   1845
      End
      Begin VB.TextBox txt特病起付累计 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1590
         TabIndex        =   35
         Tag             =   "15#"
         Top             =   690
         Width           =   1845
      End
      Begin VB.TextBox txt统筹支付累计 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5190
         TabIndex        =   33
         Tag             =   "14#"
         Top             =   300
         Width           =   1845
      End
      Begin VB.Label lbl帐户余额 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "帐户余额"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   810
         TabIndex        =   30
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lbl住院历史未补助 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "住院历史未补助"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3840
         TabIndex        =   44
         Top             =   1530
         Width           =   1260
      End
      Begin VB.Label lbl特病历史未补助 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "特病历史未补助"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   210
         TabIndex        =   42
         Top             =   1530
         Width           =   1260
      End
      Begin VB.Label lbl特殊症状起付 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "特殊症状起付"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4020
         TabIndex        =   40
         Top             =   1140
         Width           =   1080
      End
      Begin VB.Label lbl符合公务员门诊 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "符合公务员门诊"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   210
         TabIndex        =   38
         Top             =   1140
         Width           =   1260
      End
      Begin VB.Label lbl特病门诊医保费 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "特病门诊医保费"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3840
         TabIndex        =   36
         Top             =   750
         Width           =   1260
      End
      Begin VB.Label lbl特病门诊起付 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "特病门诊起付"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   390
         TabIndex        =   34
         Top             =   750
         Width           =   1080
      End
      Begin VB.Label lbl统筹支付累计 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "统筹支付"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4380
         TabIndex        =   32
         Top             =   360
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "基本信息"
      Height          =   2745
      Left            =   150
      TabIndex        =   0
      Top             =   90
      Width           =   7425
      Begin VB.TextBox txt并发症 
         Height          =   300
         Left            =   4770
         TabIndex        =   28
         Top             =   2250
         Width           =   2355
      End
      Begin VB.TextBox txt疾病信息 
         Height          =   300
         Left            =   1080
         TabIndex        =   25
         Top             =   2250
         Width           =   2055
      End
      Begin VB.TextBox txt住院次数 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4770
         TabIndex        =   22
         Tag             =   "18"
         Top             =   1860
         Width           =   555
      End
      Begin VB.CommandButton cmd疾病信息 
         Caption         =   "…"
         Height          =   300
         Left            =   3150
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   2250
         Width           =   285
      End
      Begin VB.TextBox txt身份证号 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1080
         TabIndex        =   12
         Tag             =   "0"
         Top             =   1080
         Width           =   2355
      End
      Begin VB.ComboBox cbo医疗类别 
         Height          =   300
         Left            =   4770
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   300
         Width           =   2355
      End
      Begin VB.TextBox txt单位名称 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1080
         TabIndex        =   20
         Tag             =   "8+9"
         Top             =   1860
         Width           =   2355
      End
      Begin VB.TextBox txt统筹区号 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4770
         TabIndex        =   18
         Tag             =   "7-"
         Top             =   1470
         Width           =   2355
      End
      Begin VB.CheckBox chk享受公务员补助 
         Alignment       =   1  'Right Justify
         Caption         =   "享受公务员补助"
         Enabled         =   0   'False
         Height          =   225
         Left            =   5490
         TabIndex        =   23
         Tag             =   "6"
         Top             =   1890
         Width           =   1635
      End
      Begin VB.TextBox txt行政级别 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1080
         TabIndex        =   16
         Tag             =   "5-"
         Top             =   1470
         Width           =   2355
      End
      Begin VB.TextBox txt人员类别 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4770
         TabIndex        =   14
         Tag             =   "6-"
         Top             =   1080
         Width           =   2355
      End
      Begin VB.TextBox txt年龄 
         Enabled         =   0   'False
         Height          =   300
         Left            =   6570
         TabIndex        =   10
         Tag             =   "12"
         Top             =   690
         Width           =   555
      End
      Begin VB.TextBox txt性别 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4770
         TabIndex        =   8
         Tag             =   "2"
         Top             =   690
         Width           =   1095
      End
      Begin VB.TextBox txt姓名 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1080
         TabIndex        =   6
         Tag             =   "1"
         Top             =   690
         Width           =   1365
      End
      Begin VB.TextBox txt医疗证号 
         Height          =   300
         Left            =   1080
         TabIndex        =   2
         Top             =   300
         Width           =   2355
      End
      Begin VB.Label lbl并发症 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "并发症"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4140
         TabIndex        =   27
         Top             =   2310
         Width           =   540
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
         Left            =   300
         TabIndex        =   24
         Top             =   2310
         Width           =   720
      End
      Begin VB.Label lbl住院次数 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "住院次数"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3960
         TabIndex        =   21
         Top             =   1920
         Width           =   720
      End
      Begin VB.Label lbl身份证号 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "身份证号"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   300
         TabIndex        =   11
         Top             =   1140
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
         Left            =   3960
         TabIndex        =   3
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lbl单位名称 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "单位名称"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   300
         TabIndex        =   19
         Top             =   1920
         Width           =   720
      End
      Begin VB.Label lbl统筹区号 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "统筹区号"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3960
         TabIndex        =   17
         Top             =   1530
         Width           =   720
      End
      Begin VB.Label lbl行政级别 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "行政级别"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   300
         TabIndex        =   15
         Top             =   1530
         Width           =   720
      End
      Begin VB.Label lbl人员类别 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "人员类别"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3960
         TabIndex        =   13
         Top             =   1140
         Width           =   720
      End
      Begin VB.Label lbl年龄 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "年龄"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6120
         TabIndex        =   9
         Top             =   750
         Width           =   360
      End
      Begin VB.Label lbl性别 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "性别"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4320
         TabIndex        =   7
         Top             =   750
         Width           =   360
      End
      Begin VB.Label lbl姓名 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "姓名"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   660
         TabIndex        =   5
         Top             =   750
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
         Left            =   300
         TabIndex        =   1
         Top             =   360
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmIdentify重庆银海版"
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

Private Sub cmd疾病信息_Click()
    Dim str病种编码 As String, str并发症 As String
    Dim rsTemp As New ADODB.Recordset
    str病种编码 = txt疾病信息.Tag
    str并发症 = txt并发症.Text
    If frm病种选择_重庆银海版.ShowSelect(Me, Me.cbo医疗类别.ItemData(Me.cbo医疗类别.ListIndex), str病种编码, str并发症) = True Then
        gstrSQL = "Select 名称 From 保险病种 Where 编码=[1] And 险类=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取病种的名称", str病种编码, TYPE_重庆银海版)
        txt疾病信息.Tag = str病种编码
        txt疾病信息.Text = "(" & str病种编码 & ")" & rsTemp!名称
        lbl疾病信息.Tag = txt疾病信息.Text '用于恢复显示
        txt并发症.SetFocus
    End If
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
    Dim rsTemp As New ADODB.Recordset, strText As String
    Dim blnReturn As Boolean, lst As ListItem
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If txt疾病信息.Text = "" And txt疾病信息.Tag <> "" Then Exit Sub
    
    On Error GoTo errHandle
    
    strText = txt疾病信息.Text
    If InStr(1, strText, "(") <> 0 Then
        If InStr(1, strText, ")") <> 0 Then
            strText = Mid(strText, 2, InStr(1, strText, ")") - 2)
        End If
    End If
    gstrSQL = "Select A.ID,A.编码,A.名称,A.简码" & _
             "   FROM 保险病种 A WHERE A.险类=[1] And (" & _
             "A.编码 like [2] || '%' or A.名称 like [2] || '%' or A.简码 like [2] || '%')" & Get病种类别
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_重庆银海版, strText)
    If rsTemp.RecordCount = 0 Then
        MsgBox "不存在该病种，请重新输入！", vbInformation, gstrSysName
        txt疾病信息.Text = lbl疾病信息.Tag
        zlControl.TxtSelAll txt疾病信息
        Exit Sub
    Else
        '出现选择器
        If rsTemp.RecordCount > 1 Then
            '对于字段大于3的，即使只有一条记录把该对话框显示出来，以便让用户得到更多的信息
            blnReturn = frmListSel.ShowSelect(TYPE_重庆银海版, rsTemp, "ID", "医保病种选择", "请选择医保病种：")
        Else
            blnReturn = True
        End If
    End If
    
    If blnReturn = False Then
        '记录集中没有可选择的数据
        txt疾病信息.Text = lbl疾病信息.Tag
        zlControl.TxtSelAll txt疾病信息
        Exit Sub
    Else
        '肯定是有记录集的
        txt疾病信息.Tag = rsTemp!编码
        txt疾病信息.Text = "(" & rsTemp!编码 & ")" & rsTemp!名称
        lbl疾病信息.Tag = txt疾病信息.Text '用于恢复显示
    End If
    
    Call zlCommFun.PressKey(vbKeyTab)
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
    If txt疾病信息.Tag = "" Then
        MsgBox "请输入病人的疾病信息！", vbInformation, gstrSysName
        txt疾病信息.SetFocus
        Exit Sub
    End If
    
    '获取病种ID
    gstrSQL = "Select ID From 保险病种 Where 编码=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取保险病种", CStr(txt疾病信息.Tag))
    If Not rsTemp.EOF Then
        lng疾病ID = rsTemp!ID
    End If
    
    If mbytType <> 2 Then
        '检查病人状态
        gstrSQL = "select nvl(当前状态,0) as 状态 from 保险帐户 where 险类=[1] and 医保号=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_重庆银海版, CStr(txt医疗证号.Text))
        If rsTemp.RecordCount > 0 Then
            If rsTemp("状态") > 0 Then
                MsgBox "该病人已经在院，不能通过身份验证。", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        '暂不支持挂号
        If mbytType = 3 Then
            Unload Me
            Exit Sub
        End If
    Else
        '不区分门诊和住院的，只是刷卡显示一下内容而已，不保存
        Unload Me
        Exit Sub
    End If
    
    '根据身份证号得到出生日期
    If Len(txt身份证号) >= 15 Then
        If Len(txt身份证号) = 15 Then
            str出生日期 = "19" & Mid(txt身份证号, 7, 2) & "-" & Mid(txt身份证号, 9, 2) & "-" & Mid(txt身份证号, 11, 2)
        ElseIf Len(txt身份证号) = 18 Then
            str出生日期 = Mid(txt身份证号, 7, 4) & "-" & Mid(txt身份证号, 11, 2) & "-" & Mid(txt身份证号, 13, 2)
        End If
        If Not IsDate(str出生日期) Then str出生日期 = ""
    End If
    
    '建立病人档案信息，传入格式：
    '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);
    '8.中心代码;9.顺序号;10人员身份;11帐户余额;12当前状态;13病种ID;14在职(0,1);15退休证号;16年龄段;17灰度级
    '18帐户增加累计,19帐户支出累计,20进入统筹累计,21统筹报销累计,22住院次数累计
    strIdentify = txt医疗证号.Text                              '0卡号
    strIdentify = strIdentify & ";" & txt医疗证号.Text          '1医保号
    strIdentify = strIdentify & ";"                             '2密码
    strIdentify = strIdentify & ";" & txt姓名.Text              '3姓名
    strIdentify = strIdentify & ";" & txt性别.Text              '4性别
    strIdentify = strIdentify & ";" & str出生日期               '5出生日期
    strIdentify = strIdentify & ";" & txt身份证号.Text          '6身份证
    strIdentify = strIdentify & ";" & txt单位名称.Text          '7.单位名称(编码)
    strAddition = ";0"                                          '8.中心代码
    strAddition = strAddition & ";"                             '9.顺序号
    strAddition = strAddition & ";" & Split(txt人员类别.Text, "-")(0)         '10人员身份
    strAddition = strAddition & ";" & Val(txt帐户余额.Text)     '11帐户余额
    strAddition = strAddition & ";0"                            '12当前状态
    strAddition = strAddition & ";" & lng疾病ID                 '13病种ID
    strAddition = strAddition & ";1"                            '14在职(1,2,3)
    strAddition = strAddition & ";"                             '15退休证号
    strAddition = strAddition & ";"                             '16年龄段
    strAddition = strAddition & ";"                             '17灰度级
    strAddition = strAddition & ";" & Val(txt帐户余额.Text)     '18帐户增加累计
    strAddition = strAddition & ";0"                            '19帐户支出累计
    strAddition = strAddition & ";0"                            '20上年工资总额
    strAddition = strAddition & ";" & Val(txt住院次数.Text)     '21住院次数累计
    
    mlng病人ID = BuildPatiInfo(0, strIdentify & strAddition, mlng病人ID, TYPE_重庆银海版)
    '返回格式:中间插入病人ID
    If mlng病人ID > 0 Then
        mstrReturn = strIdentify & ";" & mlng病人ID & strAddition
    End If
    
    With gComInfo_重庆银海版
        .个人编号 = txt医疗证号.Text
        .疾病编码 = txt疾病信息.Tag
        .并发症 = txt并发症.Text
        .统筹区号 = Split(txt统筹区号.Text, "-")(0)
        .业务类型 = Me.cbo医疗类别.ItemData(Me.cbo医疗类别.ListIndex)
        .帐户余额 = Val(txt帐户余额.Text)
    End With
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
'      10  医疗类别    药店购药
'      11  医疗类别    普通门诊
'      13  医疗类别    特殊病门诊
'      14  医疗类别    急诊抢救
'      15  医疗类别    白内障门诊手术
'      21  医疗类别    普通住院
'      22  医疗类别    转入医院住院
'      23  医疗类别    出院家庭病床
'      24  医疗类别    高龄家庭病床
'      25  医疗类别    急诊转住院
    With Me.cbo医疗类别
        .Clear
        '加入门诊业务
        If mbytType = 0 Or mbytType = 2 Then
'            If glngSys \ 100 <> 1 Then
'                .AddItem "药店购药"
'                .ItemData(.NewIndex) = 10
'            Else
                .AddItem "普通门诊"
                .ItemData(.NewIndex) = 11
'            End If
            .AddItem "特殊病门诊"
            .ItemData(.NewIndex) = 13
            .AddItem "急诊抢救"
            .ItemData(.NewIndex) = 14
            .AddItem "白内障门诊手术"
            .ItemData(.NewIndex) = 15
        End If
        '加入住院业务
        If mbytType = 1 Or mbytType = 2 Then
            .AddItem "普通住院"
            .ItemData(.NewIndex) = 21
            .AddItem "转入医院住院"
            .ItemData(.NewIndex) = 22
            .AddItem "高龄家庭病床"
            .ItemData(.NewIndex) = 24
            .AddItem "急诊转住院"
            .ItemData(.NewIndex) = 25
        End If
        .ListIndex = 0
    End With
End Sub

Private Sub txt医疗证号_GotFocus()
    Call zlControl.TxtSelAll(txt医疗证号)
End Sub

Private Sub txt医疗证号_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim objControl As Control           '某个控件
    Dim strTag As String            '保存控件的Tag值
    Dim strReturn As String         '接口返回的数据
    Dim arrReturn, arrTag                  '分解返回的数据而创建的数组
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(txt医疗证号.Text) = "" Then Exit Sub
    
    '调用获取参保病人的基本信息接口
    Call 调用接口_准备_重庆银海版("07", txt医疗证号.Text)
    If Not 调用接口_重庆银海版 Then Exit Sub
    strReturn = gstrReturn_重庆银海版
    arrReturn = Split(strReturn, gstrSplit_Col_重庆银海版)
    
    '序号    数据类型    长度    精度    说明
    '1.      string  18      身份号码
    '2.      string  20      姓名
    '3.      string  10      性别，返回中文含意
    '4.      string  20      民族，返回中文含意
    '5.      string  3       人员类别，见代码表
    '6.      string  3       行政级别，见代码表
    '7.      string  3       享受公务员补助标志，见代码表
    '8.      string  3       参保人员所在的统筹区号，见代码表
    '9.      string  14      单位编号
    '10.     string  50      单位名称
    '11.     string  3   0   当前待遇封锁类别，见代码表
    '12.     string  128     当前封锁原因
    '13.     number  3   0   实足年龄
    '14.     number  8   2   账户余额
    '15.     number  8   2   统筹支付累计
    '16.     number  8   2   特病门诊起付标准支付累计
    '17.     number  8   2   特病门诊医保费累计
    '18.     number  8   2   符合公务员范围门诊费用累计
    '19.     number  3   0   本年住院次数
    '20.     string  3       当前住院状态，见代码表
    '21.     string  3       本年特殊症状是否已住院标志，见代码表
    '22.     number  3       本年首次特殊症状住院次数
    '23.     string  3       本年已发生特殊症状住院最高医院等级，见代码表
    '24.     number  8   2   本年特殊症状起伏标准累计
    '25.     number  8   2   特病历史未补助自付金额
    '26.     number  8   2   住院历史未补助自付金额
    
    'Tag值的含义说明（数字表示数组下标，-表示该数据需要通过转换得到，+表示其值需要两个数组元素组合而成,第一个元素需要用()，#表示需要格式化为两位小数）
    For Each objControl In Controls
        '调试重庆医保银海版 204-04-07
        If objControl.Name <> "txt疾病信息" And objControl.Name <> "lbl疾病信息" Then
            strTag = objControl.Tag
            If Trim(strTag) <> "" Then
                If UCase(TypeName(objControl)) = "TEXTBOX" Then
                    If InStr(1, strTag, "+") <> 0 Then
                        arrTag = Split(strTag, "+")
                        objControl.Text = arrReturn(arrTag(1)) & "(" & arrReturn(arrTag(0)) & ")"
                    ElseIf InStr(1, strTag, "-") <> 0 Then
                        strTag = Replace(strTag, "-", "")
                        objControl.Text = arrReturn(Val(strTag)) & "-" & Exchange(arrReturn(Val(strTag)), strTag)
                    ElseIf InStr(1, strTag, "#") <> 0 Then
                        strTag = Replace(strTag, "#", "")
                        objControl.Text = Format(arrReturn(Val(strTag)), "#####0.00;-#####0.00; ;")
                    Else
                        objControl.Text = arrReturn(Val(strTag))
                    End If
                Else    '只可能是Checkbox
                    chk享受公务员补助.Value = Val(arrReturn(Val(strTag)))
                End If
            End If
        End If
    Next
    
    Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Function Exchange(ByVal strData As String, ByVal intOrder As Integer) As String
    Dim intCol As Integer, intCols As Integer
    Dim arrData
    Const str人员类别 As String = "31,红军时期参加工作的离休干部|32,抗战时期参加工作的离休干部|33,解放战争时期参加工作的离休干部" & _
            "|21,退休|22,退休异地居住|24,退职|25,退职异地居住|26,退职缴费人员|27,破产企业退休职工|28,建国前参加革命工作的老工人" & _
            "|11,在职|12,在职驻外|14,临时用工|15,自主择业军转干|16,再就业中心下岗职工|19,退休缴费人员"
    Const str行政级别 As String = "030,省(部)级|033,相当省(部)级|040,副省(部)级|043,相当副省(部)级|050,厅(局)级|053,相当厅(局)级" & _
            "|060,副厅(局)级|063,相当副厅(局)级|070,处级|073,相当处级|080,相当副处级|083,科级|090,相当科级|093,副科级|100,相当副科级" & _
            "|110,员级|200,其它"
    Const str统筹区号 As String = "03,渝中区|04,大渡口区|05,江北区|06,沙坪坝区|07,九龙坡区|08,南岸区|14,经开区|15,高新区|20,离休干部"
    '目前只有人员类别[6]、统筹区号[7]及行政级别[5]需要转换
    
    Select Case intOrder
    Case 5  '行政级别
        arrData = Split(str行政级别, "|")
    Case 6  '人员类别
        arrData = Split(str人员类别, "|")
    Case Else '统筹区号
        arrData = Split(str统筹区号, "|")
    End Select
    intCols = UBound(arrData)
    
    For intCol = 0 To intCols
        If strData = Split(arrData(intCol), ",")(0) Then
            Exchange = Split(arrData(intCol), ",")(1)
            Exit For
        End If
    Next
End Function

Public Function GetPatient(Optional bytType As Byte, Optional lng病人ID As Long = 0) As String
    mbytType = bytType
    mlng病人ID = lng病人ID
    mstrReturn = ""
    Me.Show 1
    lng病人ID = mlng病人ID
    GetPatient = mstrReturn
End Function

Private Function Get病种类别() As String
    '如果是门诊急诊,仅允许选择门诊急诊
    '如果是门诊特殊,仅允许选择门诊特殊
    '其他情况允许选择除门诊急诊和门诊特殊外的病种
    If Me.cbo医疗类别.ItemData(Me.cbo医疗类别.ListIndex) = 13 Then
        Get病种类别 = " And 类别 In ('1')"
    ElseIf Me.cbo医疗类别.ItemData(Me.cbo医疗类别.ListIndex) = 14 Then
        Get病种类别 = " And 类别 In ('2')"
    Else
        Get病种类别 = " And 类别 In ('0','3','4')"
    End If
End Function
