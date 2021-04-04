VERSION 5.00
Begin VB.Form frmIdentify南充阆中 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病人身份验证"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7770
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   7770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdChangePassWord 
      Caption         =   "修改密码(&P)"
      Height          =   350
      Left            =   45
      TabIndex        =   30
      Top             =   5415
      Width           =   1380
   End
   Begin VB.TextBox txtPassWord 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   5100
      MaxLength       =   6
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   900
      Width           =   2535
   End
   Begin VB.TextBox txtEdit 
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      Height          =   300
      Left            =   5100
      MaxLength       =   25
      TabIndex        =   7
      Tag             =   "社会保障号"
      Top             =   1305
      Width           =   2535
   End
   Begin VB.CommandButton cmd验卡 
      Caption         =   "重新读卡(&R)"
      Height          =   350
      Left            =   1455
      TabIndex        =   27
      Top             =   5415
      Width           =   1305
   End
   Begin VB.ComboBox cbo社保 
      Height          =   300
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   900
      Width           =   2805
   End
   Begin VB.CommandButton cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6600
      TabIndex        =   25
      Top             =   5415
      Width           =   1100
   End
   Begin VB.CommandButton cmd确定 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5445
      TabIndex        =   24
      Top             =   5415
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   75
      Index           =   1
      Left            =   -60
      TabIndex        =   28
      Top             =   615
      Width           =   8340
   End
   Begin VB.Frame fra 
      Height          =   75
      Index           =   0
      Left            =   -555
      TabIndex        =   26
      Top             =   5220
      Width           =   8340
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "工作年月"
      Height          =   180
      Index           =   29
      Left            =   6090
      TabIndex        =   68
      ToolTipText     =   "年内住院次数"
      Top             =   4935
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   27
      Left            =   6855
      TabIndex        =   67
      Top             =   4890
      Width           =   780
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "退休年月"
      Height          =   180
      Index           =   28
      Left            =   4350
      TabIndex        =   66
      ToolTipText     =   "年内住院次数"
      Top             =   4935
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   26
      Left            =   5100
      TabIndex        =   65
      Top             =   4875
      Width           =   885
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "提取时间"
      Height          =   180
      Index           =   27
      Left            =   2175
      TabIndex        =   64
      ToolTipText     =   "年内住院次数"
      Top             =   4935
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   25
      Left            =   2895
      TabIndex        =   63
      Top             =   4890
      Width           =   900
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "缴费年限"
      Height          =   180
      Index           =   26
      Left            =   240
      TabIndex        =   62
      ToolTipText     =   "缴费年限"
      Top             =   4935
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   24
      Left            =   960
      TabIndex        =   61
      Top             =   4890
      Width           =   990
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "已报销额"
      Height          =   180
      Index           =   25
      Left            =   4350
      TabIndex        =   60
      ToolTipText     =   "年内已报销金额"
      Top             =   4575
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   23
      Left            =   5100
      TabIndex        =   59
      Top             =   4515
      Width           =   2535
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "住院次数"
      Height          =   180
      Index           =   24
      Left            =   2175
      TabIndex        =   58
      ToolTipText     =   "年内住院次数"
      Top             =   4575
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   22
      Left            =   2895
      TabIndex        =   57
      Top             =   4530
      Width           =   900
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "公务员待遇"
      Height          =   180
      Index           =   23
      Left            =   60
      TabIndex        =   56
      Top             =   4575
      Width           =   900
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   21
      Left            =   960
      TabIndex        =   55
      Top             =   4530
      Width           =   990
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "补助待遇"
      Height          =   180
      Index           =   22
      Left            =   6090
      TabIndex        =   54
      Top             =   4200
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   20
      Left            =   6855
      TabIndex        =   53
      Top             =   4155
      Width           =   780
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "基本待遇"
      Height          =   180
      Index           =   21
      Left            =   4350
      TabIndex        =   52
      Top             =   4200
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   19
      Left            =   5100
      TabIndex        =   51
      Top             =   4140
      Width           =   885
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "公务员状态"
      Height          =   180
      Index           =   20
      Left            =   1995
      TabIndex        =   50
      Top             =   4200
      Width           =   900
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   18
      Left            =   2895
      TabIndex        =   49
      Top             =   4155
      Width           =   900
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "工龄"
      Height          =   180
      Index           =   19
      Left            =   600
      TabIndex        =   48
      Top             =   4200
      Width           =   360
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   17
      Left            =   960
      TabIndex        =   47
      Top             =   4155
      Width           =   990
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "基本医疗"
      Height          =   180
      Index           =   18
      Left            =   4350
      TabIndex        =   46
      Top             =   3810
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   16
      Left            =   6855
      TabIndex        =   45
      Top             =   3765
      Width           =   780
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "补充医疗"
      Height          =   180
      Index           =   17
      Left            =   6090
      TabIndex        =   44
      Top             =   3810
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   15
      Left            =   5100
      TabIndex        =   43
      Top             =   3750
      Width           =   885
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "异地居住"
      Height          =   180
      Index           =   16
      Left            =   2175
      TabIndex        =   42
      Top             =   3810
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   14
      Left            =   2895
      TabIndex        =   41
      Top             =   3765
      Width           =   900
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "个人ID"
      Height          =   180
      Index           =   15
      Left            =   420
      TabIndex        =   40
      Top             =   3810
      Width           =   540
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   13
      Left            =   960
      TabIndex        =   39
      Top             =   3765
      Width           =   990
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   12
      Left            =   6855
      TabIndex        =   38
      Top             =   3368
      Width           =   780
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   11
      Left            =   5100
      TabIndex        =   37
      Top             =   3360
      Width           =   885
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "住院状态"
      Height          =   180
      Index           =   14
      Left            =   6090
      TabIndex        =   36
      Top             =   3420
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "人员分类"
      Height          =   180
      Index           =   13
      Left            =   4350
      TabIndex        =   35
      Top             =   3420
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   10
      Left            =   2895
      TabIndex        =   34
      Top             =   3368
      Width           =   900
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   9
      Left            =   960
      TabIndex        =   33
      Top             =   3368
      Width           =   990
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "职称级别"
      Height          =   180
      Index           =   12
      Left            =   2175
      TabIndex        =   32
      Top             =   3420
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "职务级别"
      Height          =   180
      Index           =   11
      Left            =   240
      TabIndex        =   31
      Top             =   3420
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "密码"
      Height          =   180
      Index           =   9
      Left            =   4710
      TabIndex        =   2
      Top             =   960
      Width           =   360
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "记录号"
      Height          =   180
      Index           =   2
      Left            =   4530
      TabIndex        =   8
      Top             =   1785
      Width           =   540
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "出生日期"
      Height          =   180
      Index           =   4
      Left            =   4350
      TabIndex        =   20
      Top             =   2625
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   6
      Left            =   5100
      TabIndex        =   21
      Top             =   2565
      Width           =   2535
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   0
      Left            =   960
      TabIndex        =   5
      Top             =   1305
      Width           =   2805
   End
   Begin VB.Label lblInfor 
      AutoSize        =   -1  'True
      Caption         =   "医保病人基本信息显示，可以通过[重新读卡]按钮重新进行读取病人基本信息。"
      Height          =   180
      Left            =   630
      TabIndex        =   29
      Top             =   360
      Width           =   6300
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   60
      Picture         =   "frmIdentify南充阆中.frx":0000
      Top             =   105
      Width           =   480
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "卡号"
      Height          =   180
      Index           =   0
      Left            =   600
      TabIndex        =   4
      Top             =   1365
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "姓名"
      Height          =   180
      Index           =   1
      Left            =   600
      TabIndex        =   10
      Top             =   1785
      Width           =   360
   End
   Begin VB.Label lblInf 
      AutoSize        =   -1  'True
      Caption         =   "医保证号"
      Height          =   180
      Left            =   4350
      TabIndex        =   6
      Top             =   1365
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "性别"
      Height          =   180
      Index           =   3
      Left            =   600
      TabIndex        =   12
      Top             =   2190
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "身份证号"
      Height          =   180
      Index           =   5
      Left            =   240
      TabIndex        =   18
      Top             =   2625
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "帐户余额"
      Height          =   180
      Index           =   6
      Left            =   4350
      TabIndex        =   16
      Top             =   2190
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "社保机构"
      Height          =   180
      Index           =   7
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "年龄"
      Height          =   180
      Index           =   8
      Left            =   2535
      TabIndex        =   14
      Top             =   2190
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "单位名称"
      Height          =   180
      Index           =   10
      Left            =   240
      TabIndex        =   22
      Top             =   3015
      Width           =   720
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   1
      Left            =   960
      TabIndex        =   11
      Top             =   1725
      Width           =   2805
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   2
      Left            =   5100
      TabIndex        =   9
      Top             =   1725
      Width           =   2535
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   960
      TabIndex        =   13
      Top             =   2130
      Width           =   990
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   4
      Left            =   2895
      TabIndex        =   15
      Top             =   2145
      Width           =   870
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   5
      Left            =   960
      TabIndex        =   19
      Top             =   2565
      Width           =   2805
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   8
      Left            =   960
      TabIndex        =   23
      Top             =   2970
      Width           =   6675
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   7
      Left            =   5100
      TabIndex        =   17
      Top             =   2130
      Width           =   2535
   End
End
Attribute VB_Name = "frmIdentify南充阆中"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbytType As Byte            '0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐,88-按住院信息进行查询

Private mlng病人ID As Long
Private mstrReturn As String
Private mintPreCol As Integer, mintsort As Integer
Private mblnFirst As Boolean        '第一次起动系统时调用
Private mblnChange As Boolean
Private Sub cbo社保_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdChangePassWord_Click()
    Dim strOldPassWord As String
    Dim strNewPassWord As String
    Dim StrInput As String
    Dim strOutput As String
    
    If cbo社保.ListIndex < 0 Then
        ShowMsgbox "未选择医保机构编码,请选择!"
        Exit Sub
    End If
    If Split(cbo社保.Text, "_")(0) = "" Then
        ShowMsgbox "医保机构编码为空了,请重新选择!"
        Exit Sub
    End If
    g病人身份_南充阆中.机构编码 = Split(cbo社保.Text, "-")(0)
    strNewPassWord = frm修改密码.ChangePassword(strOldPassWord, strOldPassWord)
    If strOldPassWord = strNewPassWord Then Exit Sub
    If strNewPassWord = "" Then Exit Sub
    
    '    YBJGBH  PCHAR   保险机构编号
    '    COLDPASS    PCHAR   旧密码
    '    CNEWPAS PCHAR   新密码

    StrInput = g病人身份_南充阆中.机构编码
    StrInput = StrInput & vbTab & strOldPassWord
    StrInput = StrInput & vbTab & strNewPassWord
    If 业务请求_南充阆中(修改密码_旺苍, StrInput, strOutput) = False Then Exit Sub
    MsgBox "密码修改成功!", vbInformation + vbDefaultButton1, gstrSysName
    
End Sub

Private Sub cmd验卡_Click()

    If mbytType = 1 Or mbytType = 4 Or mbytType = 88 Then
        If 获取参保人员信息_住院 = False Then
            cmd确定.Enabled = False
            Call ClearData
            Exit Sub
        End If
        Call LoadCtrlData
        cmd确定.Enabled = True
        Exit Sub
    End If
    If 获取参保人员信息 = False Then
         cmd确定.Enabled = False
         Call ClearData
         Exit Sub
     End If
     Call LoadCtrlData
     cmd确定.Enabled = True
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
    Call ClearData
    
    cmd确定.Enabled = False
    If mbytType = 1 Or mbytType = 4 Or mbytType = 88 Then
        txtPassWord.Enabled = False
        txtPassWord.BackColor = lblEdit(0).BackColor
        txtEdit.Enabled = True
        txtEdit.BackColor = cbo社保.BackColor
        '曾明春:20050420 由于是读IC卡，无密码，故不提供密码的修改功能
        cmdChangePassWord.Enabled = False
    Else
        txtPassWord.Enabled = True
        txtPassWord.BackColor = cbo社保.BackColor
    End If
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
    Dim StrInput As String, strOutput As String
    Dim lng状态 As Long
    
    IsValid = False
    If Trim(g病人身份_南充阆中.姓名) = "" Then
        MsgBox "还没进行身份验证！", vbInformation, gstrSysName
        If cmd验卡.Enabled Then cmd验卡.SetFocus
        Exit Function
    End If
    
     If cbo社保.Text = "" Then
        ShowMsgbox "社保机构还未选择"
        Exit Function
    End If
      
    If mbytType <> 2 And mbytType <> 88 Then
        If mbytType = 4 Then
            '不检查当前状态
        Else
            '检查病人状态
            gstrSQL = "select nvl(当前状态,0) as 状态 from 保险帐户 where 险类=[1] and 医保号=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_南充阆中, g病人身份_南充阆中.医保证号)
            If rsTemp.RecordCount > 0 Then
                If rsTemp("状态") > 0 Then
                    MsgBox "该病人已经在院，不能通过身份验证。", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
    Else
        '不区分门诊和住院的，只是刷卡显示一下内容而已，不保存
         '需确定当前状态,因为当前状态是不能改变的
        gstrSQL = "Select * from 保险帐户 where 险类=[1] and  医保号=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_南充阆中, g病人身份_南充阆中.医保证号)
        If Not rsTemp.EOF Then
            mlng病人ID = Nvl(rsTemp!病人ID, 0)
        Else
            mlng病人ID = 0
        End If
        mstrReturn = mlng病人ID
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
    Dim StrInput  As String, strOutput As String
    Dim strIdentify As String, strAddition As String
    Dim rsTemp As New ADODB.Recordset
    Dim str社保 As String
    Dim int当前状态 As Integer
    Dim lng状态 As Long
    
    
    g病人身份_南充阆中.机构编码 = Split(cbo社保.Text, "-")(0)
    g病人身份_南充阆中.社保中心 = cbo社保.ItemData(cbo社保.ListIndex)
    If IsValid = False Then Exit Sub
    
    int当前状态 = 0
    If mbytType = 4 Then
        '需确定当前状态,因为当前状态是不能改变的
        gstrSQL = "Select * from 保险帐户 where 险类=" & TYPE_南充阆中 & " and  医保号='" & g病人身份_南充阆中.医保证号 & "'"
        
        zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
        If Not rsTemp.EOF Then
            mlng病人ID = Nvl(rsTemp!病人ID, 0)
            int当前状态 = Nvl(rsTemp!当前状态, 0)
        End If
        rsTemp.Close
    End If
    g病人身份_南充阆中.密码 = txtPassWord.Text
    
    '建立病人档案信息，传入格式：
    '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);
    '8.中心代码;9.顺序号;10人员身份;11帐户余额;12当前状态;13病种ID;14在职(0,1);15退休证号;16年龄段;17灰度级
    '18帐户增加累计,19帐户支出累计,20进入统筹累计,21统筹报销累计,22住院次数累计
    With g病人身份_南充阆中
        
        strIdentify = .医保卡号                                '0卡号
        strIdentify = strIdentify & ";" & .医保证号             '1医保号
        strIdentify = strIdentify & ";"                    '2密码
        strIdentify = strIdentify & ";" & .姓名               '3姓名
        strIdentify = strIdentify & ";" & Decode(.性别, "1", "男", "2", "女", .性别)              '4性别
        strIdentify = strIdentify & ";" & .出生日期                '5出生日期
        strIdentify = strIdentify & ";" & .身份证号码            '6身份证
        strIdentify = strIdentify & ";" & .单位名称     '7.单位名称(编码)
        strAddition = ";0" & .社保中心                                           '8.中心代码
        strAddition = strAddition & ";" & .记录号                               '9.顺序号
        strAddition = strAddition & ";" & .人员分类                                  '10人员身份
        strAddition = strAddition & ";" & .帐户余额                  '11帐户余额
        
        strAddition = strAddition & ";" & int当前状态                            '12当前状态
        strAddition = strAddition & ";"             '13病种ID
        strAddition = strAddition & ";1"                        '14在职(1,2,3)
        strAddition = strAddition & ";" & .机构编码            '15退休证号
        strAddition = strAddition & ";" & .年龄                     '16年龄段
        strAddition = strAddition & ";"                         '17灰度级
        strAddition = strAddition & ";" & .帐户余额                           '18帐户增加累计
        strAddition = strAddition & ";0"                            '19帐户支出累计
        strAddition = strAddition & ";0"                            '20上年工资总额
        strAddition = strAddition & ";"                             '21住院次数累计
    End With
    
    mlng病人ID = BuildPatiInfo(0, strIdentify & strAddition, mlng病人ID, TYPE_南充阆中)
    If mlng病人ID = 0 Then Exit Sub
    
    If mbytType = 1 Or mbytType = 4 Then
        '更新附加信息(过程参数如下:)
        '    病人ID_IN,个人ID_IN,参加工作年月_IN,退休年月_IN,职务级别_IN,职称级别_IN,异地居住标志_IN
        '    单位ID_IN,年月_IN,住院性质_IN,基本医疗标志_IN,补充医疗标志_IN,公务员标志_IN,基本待遇状态_IN,补充待遇状态_IN
        '    公务员待遇状态_IN ,年内住院次数_IN,年内已报销金额_IN,缴费年限_IN,提取时间_IN,住院记录号_IN

        gstrSQL = "zl_保险帐户补充信息_Update("
        gstrSQL = gstrSQL & "" & mlng病人ID & ","
        gstrSQL = gstrSQL & "'" & g病人身份_南充阆中.个人ID & "',"
        gstrSQL = gstrSQL & "'" & g病人身份_南充阆中.参加工作年月 & "',"
        gstrSQL = gstrSQL & "'" & g病人身份_南充阆中.退休年月 & "',"
        gstrSQL = gstrSQL & "'" & g病人身份_南充阆中.职务级别 & "',"
        gstrSQL = gstrSQL & "'" & g病人身份_南充阆中.职称级别 & "',"
        gstrSQL = gstrSQL & "'" & g病人身份_南充阆中.异地居住标志 & "',"
        gstrSQL = gstrSQL & "'" & g病人身份_南充阆中.单位ID & "',"
        gstrSQL = gstrSQL & "'" & g病人身份_南充阆中.年月 & "',"
        gstrSQL = gstrSQL & "'" & g病人身份_南充阆中.住院性质 & "',"
        gstrSQL = gstrSQL & "'" & g病人身份_南充阆中.基本医疗标志 & "',"
        gstrSQL = gstrSQL & "'" & g病人身份_南充阆中.补充医疗标志 & "',"
        gstrSQL = gstrSQL & "'" & g病人身份_南充阆中.公务员标志 & "',"
        gstrSQL = gstrSQL & "'" & g病人身份_南充阆中.基本待遇状态 & "',"
        gstrSQL = gstrSQL & "'" & g病人身份_南充阆中.补充待遇状态 & "',"
        gstrSQL = gstrSQL & "'" & g病人身份_南充阆中.公务员待遇状态 & "',"
        gstrSQL = gstrSQL & "'" & g病人身份_南充阆中.年内住院次数 & "',"
        gstrSQL = gstrSQL & "'" & g病人身份_南充阆中.年内已报销金额 & "',"
        gstrSQL = gstrSQL & "'" & g病人身份_南充阆中.缴费年限 & "',"
        gstrSQL = gstrSQL & "'" & g病人身份_南充阆中.提取时间 & "',"
        gstrSQL = gstrSQL & "'" & g病人身份_南充阆中.住院记录号 & "')"
        ExecuteProcedure_南充阆中 "保存帐户附加信息"
    Else
    End If
    g病人身份_南充阆中.病人ID = mlng病人ID
    
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
    If Load社保机构 = False Then
        DebugTool "加入失败(身份验证)"
        Exit Function
    End If
    DebugTool "加入成功(身份验证)"
    
    Me.Show 1
    lng病人ID = mlng病人ID
    GetPatient = mstrReturn
End Function
Private Sub LoadCtrlData()
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:填充数据
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    With g病人身份_南充阆中
        lblEdit(0).Caption = .医保卡号
        txtEdit.Text = .医保证号
        lblEdit(1).Caption = .姓名
        lblEdit(2).Caption = IIf(mbytType = 1 Or mbytType = 4, .住院记录号, .记录号)
        lblEdit(3).Caption = Decode(.性别, "1", "男", "2", "女", .性别)
        lblEdit(4).Caption = .年龄
        lblEdit(5).Caption = .身份证号码
        lblEdit(6).Caption = .出生日期
        lblEdit(7).Caption = Format(.帐户余额, "####0.00;-####0.00;;")
        lblEdit(8).Caption = .单位名称
        lblEdit(9).Caption = .职务级别
        lblEdit(10).Caption = .职称级别
        lblEdit(11).Caption = .人员分类
        lblEdit(12).Caption = ""
        
        lblEdit(13).Caption = .个人ID
        lblEdit(14).Caption = .异地居住标志
        lblEdit(15).Caption = .基本医疗标志
        lblEdit(16).Caption = .补充医疗标志
        lblEdit(17).Caption = .年月
        lblEdit(18).Caption = .公务员标志
        lblEdit(19).Caption = .基本待遇状态
        lblEdit(20).Caption = .补充待遇状态
        lblEdit(21).Caption = .公务员待遇状态
        lblEdit(22).Caption = .年内住院次数
        lblEdit(23).Caption = .年内已报销金额
        
        lblEdit(24).Caption = .缴费年限
        lblEdit(25).Caption = .提取时间
        
        lblEdit(26).Caption = .退休年月
        lblEdit(27).Caption = .参加工作年月
        
    End With
End Sub

Private Sub Form_Load()
        mblnFirst = True
End Sub

Private Function Load社保机构() As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "Select * From 保险中心目录 where 险类=" & TYPE_南充阆中 & " and 序号<>0 order by 编码"

    Err = 0
    On Error GoTo errHand:
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption & "社保机构目录"
    If rsTemp.RecordCount = 0 Then
        ShowMsgbox "不存社保机构目录，请在参数中下载机构!"
        Exit Function
    End If
    
    With rsTemp
        cbo社保.Clear
        Do While Not .EOF
            cbo社保.AddItem Nvl(!编码) & "--" & Nvl(!名称)
            cbo社保.ItemData(cbo社保.NewIndex) = Nvl(!序号, 0)
            .MoveNext
        Loop
    End With
    cbo社保.ListIndex = 0
    SetDefaultSel
    Load社保机构 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function
Private Function SetDefaultSel() As Boolean
    Dim strReg As String
    Dim i As Integer
    
    SetDefaultSel = False
    Err = 0: On Error GoTo errHand:
    Call GetRegInFor(g公共模块, "医保", "社保机构代码", strReg)
    If cbo社保.ListCount = 0 Then Exit Function
    For i = 0 To cbo社保.ListCount - 1
        If Split(cbo社保.List(i) & "--", "--")(0) = strReg Then
            cbo社保.ListIndex = i
            Exit For
        End If
    Next
    If cbo社保.ListIndex < 0 Then
        cbo社保.ListIndex = 0
    End If
    SetDefaultSel = True
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function
Private Function 获取参保人员信息_住院() As Boolean
    '功能:获取参保人员信息(住院部份)
    Dim StrInput As String, strOutput As String
    Dim bln输入 As Boolean
    Dim strArr As Variant
    
    获取参保人员信息_住院 = False
    Err = 0: On Error GoTo errHand:
    If Trim(txtEdit.Text) <> "" Then
        StrInput = Trim(txtEdit.Text) & vbTab
        bln输入 = True
    End If
    StrInput = StrInput & Split(cbo社保.Text, "--")(0)
    '曾明春:20050420 如果是读IC卡,首先根据IC卡中读出的医保证号,再取参保人员的基本资料
    If bln输入 = False Then
        If 业务请求_南充阆中(获得人员资料_读卡_旺苍, StrInput, strOutput) = False Then
        Exit Function
        Else: strArr = Split(strOutput, "||")
              StrInput = Split(strArr(0), "--")(0) & vbTab & Split(cbo社保.Text, "--")(0)
              txtEdit.Text = strArr(0)
              lblEdit(0).Caption = strArr(1)
        End If
    End If
    If 业务请求_南充阆中(获得人员资料_医保号_旺苍, StrInput, strOutput) = False Then
       Exit Function
    End If
    strArr = Split(strOutput, "||")
    
    '个人ID||社保编号||姓名||性别||出生日期（格式：YYYY-MM-DD）||参加工作年月||退休年月||职务级别||职称级别||人员分类||
    '异地居住标志||单位ID||单位名称||年龄||年月||医保证号||住院性质||基本医疗标志||补充医疗标志||公务员标志||基本医疗待遇状态||
    '补充医疗待遇状态||公务员待遇状态||年内往院次料||年内已报销金额||缴费年限||提取时间||住院记录号||

    With g病人身份_南充阆中
        .医保卡号 = lblEdit(0).Caption
        .医保证号 = txtEdit.Text
        .记录号 = strArr(0)
        .姓名 = strArr(2)
        .身份证号码 = strArr(1)
        .单位名称 = strArr(12)
        .性别 = strArr(3)
        .出生日期 = zlCommFun.AddDate(strArr(4))
        .年龄 = Val(strArr(14))
        .机构编码 = Split(cbo社保.Text, "--")(0)
        .社保中心 = cbo社保.ItemData(cbo社保.ListIndex)
        .个人ID = strArr(0)
        .参加工作年月 = strArr(5)
        
        .退休年月 = strArr(6)
        .职务级别 = strArr(7)
        .职称级别 = strArr(8)
        .人员分类 = strArr(9)
        .异地居住标志 = strArr(10)
        .单位ID = strArr(11)
        .年月 = strArr(13)   '工龄
        .住院性质 = strArr(16)
        .基本医疗标志 = "" 'strArr(17)
        .补充医疗标志 = "" 'strArr(18)
        .公务员标志 = "" 'strArr(19)
        .基本待遇状态 = "" 'strArr(20)
        .补充待遇状态 = "" 'strArr(21)
        .公务员待遇状态 = "" 'strArr(22)
        .年内住院次数 = "" 'strArr(23)
        .年内已报销金额 = "" 'strArr(24)
        .缴费年限 = "" 'strArr(25)
        .提取时间 = "" 'strArr(26)
        .住院记录号 = "" 'strArr(27)
        .str住院信息 = strOutput
    End With
    获取参保人员信息_住院 = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Function 获取参保人员信息() As Boolean
    '获取参保人员信息
    Dim StrInput As String
    Dim strOutput As String
    Dim strArr
    
    获取参保人员信息 = False
    
    Err = 0
    On Error GoTo errHand:
    If cbo社保.ListIndex < 0 Then
        MsgBox "未选择社保机构编码,请选择!", vbInformation + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    With g病人身份_南充阆中
        .机构编码 = Split(cbo社保.Text, "--")(0)
    End With
    If txtPassWord.Text = "" Then
           MsgBox "你未输入密码,请输入!", vbInformation + vbDefaultButton1, gstrSysName
           If txtPassWord.Enabled Then txtPassWord.SetFocus
           g病人身份_南充阆中.密码 = ""
           Exit Function
    End If
    g病人身份_南充阆中.密码 = txtPassWord.Text
    If 业务请求_南充阆中(获得参保人员资料_旺苍, "", strOutput) = False Then
        Call ClearData
        Exit Function
    End If
    strArr = Split(strOutput, "||")
    '返回:医保卡号||医保证号||个人记录号||姓名||身份证号码||单位名称||性别||出生日期
    
    With g病人身份_南充阆中
        .医保卡号 = strArr(0)
        .医保证号 = strArr(1)
        .记录号 = strArr(2)
        .姓名 = strArr(3)
        .身份证号码 = strArr(4)
        .单位名称 = strArr(5)
        .性别 = strArr(6)
        .出生日期 = zlCommFun.AddDate(strArr(7))
        .年龄 = Get年龄(.出生日期)
        .社保中心 = cbo社保.ItemData(cbo社保.ListIndex)
        
        .个人ID = ""
        .参加工作年月 = ""
        
        .退休年月 = ""
        .职务级别 = ""
        .职称级别 = ""
        .人员分类 = ""
        .异地居住标志 = ""
        .单位ID = ""
        .年月 = ""
        .住院性质 = ""
        .基本医疗标志 = ""
        .补充医疗标志 = ""
        .公务员标志 = ""
        .基本待遇状态 = ""
        .补充待遇状态 = ""
        .公务员待遇状态 = ""
        .年内住院次数 = ""
        .年内已报销金额 = ""
        .缴费年限 = ""
        .提取时间 = ""
        .住院记录号 = ""
        
        .str住院信息 = ""
    End With
    
    '获取帐户余额
    '    YBJGBH  PCHAR   保险机构编号
    '    CPASSWORD   PCHAR   持卡人卡密码
    '有问题，根据机构编号怎么获取.
    StrInput = g病人身份_南充阆中.机构编码
    StrInput = StrInput & vbTab & g病人身份_南充阆中.密码
    If 业务请求_南充阆中(获取帐户余额_旺苍, StrInput, strOutput) = False Then Exit Function
    g病人身份_南充阆中.帐户余额 = Val(strOutput)
    
    获取参保人员信息 = True
    Exit Function
errHand:
        If ErrCenter = 1 Then
            Resume
        End If
End Function
Private Function Get年龄(ByVal strDate As String) As Integer
    Dim rsTemp As New ADODB.Recordset
    Err = 0
    On Error GoTo errHand:
    gstrSQL = "Select (sysdate-to_date('" & strDate & "','yyyy-mm-dd'))/365 as 年龄 from dual "
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    If Not rsTemp.EOF Then
        Get年龄 = Int(Nvl(rsTemp!年龄, 0))
        Exit Function
    End If
    Exit Function
errHand:
End Function
Private Sub ClearData()
    Dim i As Long
    '清除相关信息
    With g病人身份_南充阆中
        .医保卡号 = ""
        .医保证号 = ""
        .记录号 = ""
        .姓名 = ""
        .身份证号码 = ""
        .单位名称 = ""
        .性别 = ""
        .出生日期 = ""
        .年龄 = 0
   
        .退休年月 = ""
        .职务级别 = ""
        .职称级别 = ""
        .人员分类 = ""
        .异地居住标志 = ""
        .单位ID = ""
        .年月 = ""
        .住院性质 = ""
        .基本医疗标志 = ""
        .补充医疗标志 = ""
        .公务员标志 = ""
        .基本待遇状态 = ""
        .补充待遇状态 = ""
        .公务员待遇状态 = ""
        .年内住院次数 = ""
        .年内已报销金额 = ""
        .缴费年限 = ""
        .提取时间 = ""
        .住院记录号 = ""
        
        .str住院信息 = ""
    End With
    For i = 0 To lblEdit.UBound
        lblEdit(i).Caption = ""
    Next
End Sub
Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If 获取参保人员信息_住院 = False Then
        cmd确定.Enabled = False
        Call ClearData
        Exit Sub
    End If
    Call LoadCtrlData
    cmd确定.Enabled = True
    zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtEdit, KeyAscii, m文本式
End Sub
Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If 获取参保人员信息 = False Then
        cmd确定.Enabled = False
        Call ClearData
        Exit Sub
    End If
    Call LoadCtrlData
    cmd确定.Enabled = True
    zlCommFun.PressKey vbKeyTab
End Sub

