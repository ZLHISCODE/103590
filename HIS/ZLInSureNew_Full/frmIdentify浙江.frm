VERSION 5.00
Begin VB.Form frmIdentify浙江 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "身份验证"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5595
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -473
      TabIndex        =   21
      Top             =   1965
      Width           =   6450
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   8
      Left            =   3671
      TabIndex        =   20
      Top             =   1485
      Width           =   1755
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   7
      Left            =   956
      TabIndex        =   18
      Top             =   1485
      Width           =   1755
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   6
      Left            =   3671
      TabIndex        =   16
      Top             =   1035
      Width           =   1755
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   5
      Left            =   956
      TabIndex        =   14
      Top             =   1035
      Width           =   1755
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   4
      Left            =   3671
      TabIndex        =   12
      Top             =   585
      Width           =   1755
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   3
      Left            =   956
      TabIndex        =   10
      Top             =   585
      Width           =   1755
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   2
      Left            =   3671
      TabIndex        =   8
      Top             =   150
      Width           =   1755
   End
   Begin VB.TextBox txtInfo 
      Enabled         =   0   'False
      Height          =   300
      Index           =   1
      Left            =   956
      TabIndex        =   0
      Top             =   150
      Width           =   1755
   End
   Begin VB.TextBox txtInfo 
      Height          =   300
      Index           =   0
      Left            =   956
      MaxLength       =   10
      TabIndex        =   4
      Top             =   150
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   400
      Left            =   4326
      TabIndex        =   3
      Top             =   2145
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Enabled         =   0   'False
      Height          =   400
      Left            =   3191
      TabIndex        =   2
      Top             =   2145
      Width           =   1100
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "读卡(&R)"
      Height          =   400
      Left            =   176
      TabIndex        =   1
      Top             =   2145
      Width           =   1100
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "出生日期"
      Height          =   180
      Index           =   8
      Left            =   2880
      TabIndex        =   19
      Top             =   1560
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "民族"
      Height          =   180
      Index           =   7
      Left            =   525
      TabIndex        =   17
      Top             =   1560
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "性别"
      Height          =   180
      Index           =   6
      Left            =   3240
      TabIndex        =   15
      Top             =   1110
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "人员状态"
      Height          =   180
      Index           =   5
      Left            =   165
      TabIndex        =   13
      Top             =   1110
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "身份证号"
      Height          =   180
      Index           =   4
      Left            =   2880
      TabIndex        =   11
      Top             =   660
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "姓名"
      Height          =   180
      Index           =   3
      Left            =   525
      TabIndex        =   9
      Top             =   660
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "医保号"
      Height          =   180
      Index           =   2
      Left            =   3060
      TabIndex        =   7
      Top             =   225
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "卡号"
      Height          =   180
      Index           =   1
      Left            =   525
      TabIndex        =   6
      Top             =   225
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "密码"
      Height          =   180
      Index           =   0
      Left            =   529
      TabIndex        =   5
      Top             =   225
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "frmIdentify浙江"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytType As Byte
Public mstrPatient As String, mstrOther As String

Public Function GetPatient(bytType As Byte, str卡号 As String) As String
    '参数：bytType-识别类型，0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐
    mbytType = bytType
    Me.Show vbModal
    str卡号 = txtInfo(1).Text
    GetPatient = mstrPatient & mstrOther
End Function

Private Sub cmdCancel_Click()
    mstrPatient = "": mstrOther = ""
    Me.Hide
    '取消
End Sub

Private Sub cmdOK_Click()
    '确定
    mstrOther = "": mstrPatient = ""
reQuery1:
    gstrInfo = Space(1024)
    glngReturn = QUERY_HANDLE("13|" & txtInfo(2).Text & "|DF0431|", gstrInfo)  '读取病人信息
    WriteInfo Trim(gstrInfo)
    If CheckReturn浙江() = False Then
        If gstrInfo = "" Then
            GoTo reQuery1
        Else
            Exit Sub
        End If
    End If
    txtInfo(0).Tag = "(" & Split(gstrInfo, "|")(8) & ")"
    
reQuery2:
    gstrInfo = Space(1024)
    glngReturn = QUERY_HANDLE("13|" & txtInfo(2).Text & "|DF0432|", gstrInfo)  '读取病人信息
    WriteInfo Trim(gstrInfo)
    If CheckReturn浙江() = False Then
        If gstrInfo = "" Then
            GoTo reQuery2
        Else
            Exit Sub
        End If
    End If
    txtInfo(2).Tag = Split(gstrInfo, "|")(0)
    
    mstrPatient = txtInfo(1).Text & ";"                                 '0 卡号
    mstrPatient = mstrPatient & txtInfo(2).Text & ";"                   '1 医保帐号
    mstrPatient = mstrPatient & txtInfo(0).Text & ";"                   '2 密码
    mstrPatient = mstrPatient & txtInfo(3).Text & ";"                   '3 姓名
    mstrPatient = mstrPatient & txtInfo(6).Text & ";"                   '4 性别
    mstrPatient = mstrPatient & txtInfo(8).Text & ";"                   '5 出生日期
    mstrPatient = mstrPatient & txtInfo(4).Text & ";"                   '6 身份证
    mstrPatient = mstrPatient & txtInfo(0).Tag & ";"                    '7 单位名称/编码
    
    mstrOther = mstrOther & txtInfo(1).Tag & ";"                        '8 医保机构编码(中心)
    mstrOther = mstrOther & ";"                                         '9 顺序号
    mstrOther = mstrOther & ";"                                         '10 身份
    mstrOther = mstrOther & txtInfo(2).Tag & ";"                        '11 余额
    mstrOther = mstrOther & ";"                                         '12 当前状态
    mstrOther = mstrOther & ";"                                         '13 病种ID
    mstrOther = mstrOther & ";"                                         '14 在职状态
    mstrOther = mstrOther & ";"                                         '15 退休证号
    mstrOther = mstrOther & ";"                                         '16 年龄段
    mstrOther = mstrOther & ";"                                         '17 灰度级
    mstrOther = mstrOther & ";"                                         '18 帐户增加累计
    mstrOther = mstrOther & ";"                                         '19 帐户支出累计
    mstrOther = mstrOther & ";"                                         '20 进入统筹累计
    mstrOther = mstrOther & ";"                                         '21 统筹报销累计
    mstrOther = mstrOther & ";"                                         '22 住院次数累计
    mstrOther = mstrOther & ";"                                         '23 就诊类别
    mstrOther = mstrOther & ";"                                         '24 本次起付线
    mstrOther = mstrOther & ";"                                         '25 起付线累计
    mstrOther = mstrOther & ";"                                         '26 基本统筹限额
    
    Me.Hide
End Sub

Private Sub cmdRead_Click()
    '读卡
    Dim strTemp As String
    Dim strfksj As String
    Dim datCurr As Date
    
    datCurr = zlDatabase.Currentdate
    
    WriteInfo "开始读卡"
    cmdOK.Enabled = False
    If txtInfo(1).Text = "" Then
        WriteInfo "从IC卡获取数据"
        glngReturn = readCardID(strTemp)
        WriteInfo "读卡返回：" & strTemp
        If glngReturn < 0 Then Exit Sub
        txtInfo(1).Text = Trim(strTemp)
    Else
        WriteInfo "手工输入卡号"
    End If
reQuery01:
      '*******06年06月13日补充，加入持卡人资格审查功能********
      '取发卡时间
      gstrInfo = Space(1024)
      glngReturn = QUERY_HANDLE("13|" & txtInfo(1).Text & "|MF11|", gstrInfo)
      If CheckReturn浙江() = False Then
        If gstrInfo = "" Then
            GoTo reQuery01
        Else
            Exit Sub
        End If
     End If
      WriteInfo Trim(gstrInfo)
     strfksj = Trim(Split(gstrInfo, "|")(4))
     '取资格信息
      gstrInfo = Space(1024)
      glngReturn = QUERY_HANDLE("04|" & txtInfo(1).Text & "|" & Format(datCurr, "yyyymmdd") & "|" & strfksj & "|", gstrInfo)
      If CheckReturn浙江() = False Then
        If gstrInfo = "" Then
            GoTo reQuery01
        Else
            Exit Sub
        End If
     End If
      WriteInfo Trim(gstrInfo)
           
      If Trim(Split(gstrInfo, "|")(0)) = "0" Then '无封锁
      Else
         If Trim(Split(gstrInfo, "|")(0)) = "1" Then '个人封锁
             MsgBox "该卡被个人封锁，日期范围：" & Trim(Split(gstrInfo, "|")(1)) = "1" & "---" & Trim(Split(gstrInfo, "|")(2)) & " 封锁原因：" & Trim(Split(gstrInfo, "|")(3))
         Else '单位封锁
             MsgBox "该卡被单位封锁，日期范围：" & Trim(Split(gstrInfo, "|")(1)) = "1" & "---" & Trim(Split(gstrInfo, "|")(2)) & " 封锁原因：" & Trim(Split(gstrInfo, "|")(3))
         End If
         Exit Sub
      End If
                 
      '*********结束*********
    
    gstrInfo = Space(1024)
     
    glngReturn = QUERY_HANDLE("13|" & txtInfo(1).Text & "|MF12|", gstrInfo)
    WriteInfo Trim(gstrInfo)
    If CheckReturn浙江() = False Then
        If gstrInfo = "" Then
            GoTo reQuery01
        Else
            Exit Sub
        End If
    End If
    cmdOK.Enabled = True
    txtInfo(1).Text = Trim(Split(gstrInfo, "|")(0))
    txtInfo(2).Text = Trim(Split(gstrInfo, "|")(1))
    txtInfo(3).Text = Trim(Split(gstrInfo, "|")(3))
    txtInfo(4).Text = Trim(Split(gstrInfo, "|")(2))
'    txtInfo(5).Text = IIf(Trim(Split(gstrInfo, "|")(4)) = "1", "公务员", "非公务员")
    txtInfo(6).Text = IIf(Trim(Split(gstrInfo, "|")(5)) = "1", "男", "女")
    txtInfo(7).Text = Trim(Split(gstrInfo, "|")(6))
    strTemp = Trim(Split(gstrInfo, "|")(7))
    txtInfo(8).Text = Left(strTemp, 4) & "-" & Mid(strTemp, 5, 2) & "-" & Mid(strTemp, 7)
    On Error Resume Next
    cmdOK.SetFocus
    cmdRead.Enabled = False
reQuery1:
    gstrInfo = Space(1024)
    glngReturn = QUERY_HANDLE("13|" & txtInfo(2).Text & "|DF0431|", gstrInfo)  '读取病人信息
    WriteInfo Trim(gstrInfo)
    If CheckReturn浙江() = False Then
        If gstrInfo = "" Then
            GoTo reQuery1
        Else
            Exit Sub
        End If
    End If
    txtInfo(5).Text = IIf(Trim(Split(gstrInfo, "|")(9)) = "1", "公务员", "非公务员")
End Sub

Private Sub Form_Load()
    cmdOK.Enabled = False
End Sub

Private Sub txtInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 1 And KeyAscii = vbKeyReturn Then
        cmdRead_Click
    End If
End Sub
