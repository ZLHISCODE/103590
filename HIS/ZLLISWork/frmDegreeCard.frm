VERSION 5.00
Begin VB.Form frmDegreeCard 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病人信息"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8895
   Icon            =   "frmDegreeCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fra在院信息 
      ForeColor       =   &H00C00000&
      Height          =   1335
      Left            =   75
      TabIndex        =   80
      Top             =   30
      Width           =   8730
      Begin VB.TextBox txt护理 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3090
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   885
         Width           =   1170
      End
      Begin VB.TextBox txt床位等级 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   885
         Width           =   1170
      End
      Begin VB.TextBox txt出院时间 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   7335
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   885
         Width           =   1110
      End
      Begin VB.TextBox txt入院时间 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5250
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   885
         Width           =   1065
      End
      Begin VB.TextBox txt科室 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   7335
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   225
         Width           =   1110
      End
      Begin VB.TextBox txt床号 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5235
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   225
         Width           =   1065
      End
      Begin VB.TextBox txt住院号 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3090
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   225
         Width           =   1170
      End
      Begin VB.TextBox txt费别 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   7335
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   555
         Width           =   1110
      End
      Begin VB.TextBox txt性别 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3090
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   555
         Width           =   1170
      End
      Begin VB.TextBox txt病人ID 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   225
         Width           =   1170
      End
      Begin VB.TextBox txt姓名 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   555
         Width           =   1170
      End
      Begin VB.TextBox txt年龄 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5250
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   555
         Width           =   1065
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "护理"
         Height          =   180
         Left            =   2700
         TabIndex        =   18
         Top             =   945
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "床位等级"
         Height          =   180
         Left            =   345
         TabIndex        =   16
         Top             =   945
         Width           =   720
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出院时间"
         Height          =   180
         Left            =   6585
         TabIndex        =   22
         Top             =   960
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "入院时间"
         Height          =   180
         Left            =   4470
         TabIndex        =   20
         Top             =   945
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "科室"
         Height          =   180
         Left            =   6930
         TabIndex        =   6
         Top             =   285
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "床号"
         Height          =   180
         Left            =   4845
         TabIndex        =   4
         Top             =   285
         Width           =   360
      End
      Begin VB.Label lbl住院号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院号"
         Height          =   180
         Left            =   2505
         TabIndex        =   2
         Top             =   285
         Width           =   540
      End
      Begin VB.Label lbl姓名 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名"
         Height          =   180
         Left            =   705
         TabIndex        =   8
         Top             =   615
         Width           =   360
      End
      Begin VB.Label lbl性别 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别"
         Height          =   180
         Left            =   2700
         TabIndex        =   10
         Top             =   615
         Width           =   360
      End
      Begin VB.Label lbl年龄 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄"
         Height          =   180
         Left            =   4830
         TabIndex        =   12
         Top             =   615
         Width           =   360
      End
      Begin VB.Label lbl费别 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "费别"
         Height          =   180
         Left            =   6930
         TabIndex        =   14
         Top             =   630
         Width           =   360
      End
      Begin VB.Label lbl病人ID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人ID"
         ForeColor       =   &H80000007&
         Height          =   180
         Left            =   510
         TabIndex        =   0
         Top             =   285
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   780
      TabIndex        =   77
      Top             =   5460
      Width           =   1100
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "退出(&X)"
      Height          =   350
      Left            =   7050
      TabIndex        =   76
      Top             =   5460
      Width           =   1100
   End
   Begin VB.Frame fra基本信息 
      ForeColor       =   &H00C00000&
      Height          =   3345
      Left            =   75
      TabIndex        =   78
      Top             =   1380
      Width           =   8745
      Begin VB.TextBox txt医疗付款 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   240
         Width           =   1170
      End
      Begin VB.TextBox txt出生日期 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   7335
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   570
         Width           =   1140
      End
      Begin VB.TextBox txt联系人关系 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   1890
         Width           =   2000
      End
      Begin VB.TextBox txt身份 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   570
         Width           =   1170
      End
      Begin VB.TextBox txt职业 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3105
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   570
         Width           =   1170
      End
      Begin VB.TextBox txt婚姻状况 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   570
         Width           =   1170
      End
      Begin VB.TextBox txt国籍 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3105
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   240
         Width           =   1170
      End
      Begin VB.TextBox txt学历 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   7335
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   240
         Width           =   1140
      End
      Begin VB.TextBox txt民族 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   240
         Width           =   1170
      End
      Begin VB.TextBox txt出生地点 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   900
         Width           =   3225
      End
      Begin VB.TextBox txt家庭地址 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   1230
         Width           =   2790
      End
      Begin VB.TextBox txt户口邮编 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   47
         Top             =   1230
         Width           =   1170
      End
      Begin VB.TextBox txt联系人姓名 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   1560
         Width           =   1170
      End
      Begin VB.TextBox txt联系人地址 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   55
         Top             =   1890
         Width           =   3225
      End
      Begin VB.TextBox txt联系人电话 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   57
         Top             =   2220
         Width           =   2000
      End
      Begin VB.TextBox txt工作单位 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   59
         Top             =   2220
         Width           =   3225
      End
      Begin VB.TextBox txt单位电话 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   61
         Top             =   2550
         Width           =   2000
      End
      Begin VB.TextBox txt单位邮编 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   63
         Top             =   2550
         Width           =   1170
      End
      Begin VB.TextBox txt单位开户行 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   65
         Top             =   2880
         Width           =   3135
      End
      Begin VB.TextBox txt单位帐号 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   67
         Top             =   2880
         Width           =   3225
      End
      Begin VB.TextBox txt家庭电话 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   49
         Top             =   1560
         Width           =   2000
      End
      Begin VB.TextBox txt身份证号 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   41
         Top             =   900
         Width           =   3150
      End
      Begin VB.Label lbl医疗付款 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "医疗付款"
         Height          =   180
         Left            =   345
         TabIndex        =   24
         Top             =   300
         Width           =   720
      End
      Begin VB.Label lbl出生日期 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出生日期"
         Height          =   180
         Left            =   6570
         TabIndex        =   38
         Top             =   630
         Width           =   720
      End
      Begin VB.Label lbl出生地点 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出生地点"
         Height          =   180
         Left            =   4470
         TabIndex        =   42
         Top             =   960
         Width           =   720
      End
      Begin VB.Label lbl身份证号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "身份证号"
         Height          =   180
         Left            =   345
         TabIndex        =   40
         Top             =   960
         Width           =   720
      End
      Begin VB.Label lbl身份 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "身份"
         Height          =   180
         Left            =   4830
         TabIndex        =   36
         Top             =   630
         Width           =   360
      End
      Begin VB.Label lbl职业 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "职业"
         Height          =   180
         Left            =   2685
         TabIndex        =   34
         Top             =   630
         Width           =   360
      End
      Begin VB.Label lbl民族 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "民族"
         Height          =   180
         Left            =   4830
         TabIndex        =   28
         Top             =   300
         Width           =   360
      End
      Begin VB.Label lbl国籍 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "国籍"
         Height          =   180
         Left            =   2685
         TabIndex        =   26
         Top             =   300
         Width           =   360
      End
      Begin VB.Label lbl学历 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "学历"
         Height          =   180
         Left            =   6930
         TabIndex        =   30
         Top             =   300
         Width           =   360
      End
      Begin VB.Label lvl婚姻状况 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "婚姻状况"
         Height          =   180
         Left            =   345
         TabIndex        =   32
         Top             =   630
         Width           =   720
      End
      Begin VB.Label lbl家庭地址 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "家庭地址"
         Height          =   180
         Left            =   345
         TabIndex        =   44
         Top             =   1290
         Width           =   720
      End
      Begin VB.Label lbl家庭电话 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "家庭电话"
         Height          =   180
         Left            =   345
         TabIndex        =   48
         Top             =   1620
         Width           =   720
      End
      Begin VB.Label lbl户口邮编 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "家庭地址邮编"
         Height          =   180
         Left            =   4110
         TabIndex        =   46
         Top             =   1290
         Width           =   1080
      End
      Begin VB.Label lbl联系人姓名 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "联系人姓名"
         Height          =   180
         Left            =   4290
         TabIndex        =   50
         Top             =   1635
         Width           =   900
      End
      Begin VB.Label lbl联系人关系 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "联系人关系"
         Height          =   180
         Left            =   165
         TabIndex        =   52
         Top             =   1950
         Width           =   900
      End
      Begin VB.Label lbl联系人地址 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "联系人地址"
         Height          =   180
         Left            =   4290
         TabIndex        =   54
         Top             =   1950
         Width           =   900
      End
      Begin VB.Label lbl联系人电话 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "联系人电话"
         Height          =   180
         Left            =   165
         TabIndex        =   56
         Top             =   2280
         Width           =   900
      End
      Begin VB.Label lbl工作单位 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "工作单位"
         Height          =   180
         Left            =   4470
         TabIndex        =   58
         Top             =   2280
         Width           =   720
      End
      Begin VB.Label lbl单位电话 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位电话"
         Height          =   180
         Left            =   345
         TabIndex        =   60
         Top             =   2610
         Width           =   720
      End
      Begin VB.Label lbl单位邮编 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位邮编"
         Height          =   180
         Left            =   4470
         TabIndex        =   62
         Top             =   2610
         Width           =   720
      End
      Begin VB.Label lbl单位开户行 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位开户行"
         Height          =   180
         Left            =   165
         TabIndex        =   64
         Top             =   2940
         Width           =   900
      End
      Begin VB.Label lbl单位帐号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位帐号"
         Height          =   180
         Left            =   4470
         TabIndex        =   66
         Top             =   2940
         Width           =   720
      End
   End
   Begin VB.Frame fra费用信息 
      ForeColor       =   &H00C00000&
      Height          =   705
      Left            =   75
      TabIndex        =   79
      Top             =   4530
      Width           =   8745
      Begin VB.TextBox txt费用余额 
         Alignment       =   1  'Right Justify
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   71
         Top             =   240
         Width           =   1155
      End
      Begin VB.TextBox txt预交余额 
         Alignment       =   1  'Right Justify
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   69
         Top             =   240
         Width           =   1155
      End
      Begin VB.TextBox txt担保额 
         Alignment       =   1  'Right Justify
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   7380
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   75
         Top             =   240
         Width           =   1080
      End
      Begin VB.TextBox txt担保人 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   73
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "未结费用"
         Height          =   180
         Left            =   2370
         TabIndex        =   70
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "预交余额"
         Height          =   180
         Left            =   375
         TabIndex        =   68
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "担保额"
         Height          =   180
         Left            =   6765
         TabIndex        =   74
         Top             =   300
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "担保人"
         Height          =   180
         Left            =   4695
         TabIndex        =   72
         Top             =   300
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmDegreeCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit '要求变量声明

Public Function ShowInfo(ByVal frmMain As Form, ByVal lngKey As Long) As Boolean
    
    If ReadCard(lngKey) = False Then
        MsgBox "不能正确读取病人信息,请与系统管理员联系！", vbInformation, gstrSysName
        Exit Function
    End If
    
    Me.Show 1, frmMain
            
    ShowInfo = True
    
End Function

Private Function ReadCard(ByVal lng病人ID As Long) As Boolean

    '功能：读取指定病人信息,并显示在界面上
    
    Dim rsTmp As New ADODB.Recordset
    Dim strsql As String
    
    On Error GoTo ErrHand
        
    strsql = "Select A.*," & _
            "DECODE(A.当前科室id,NULL,就诊诊室,(SELECT 名称 FROM 部门表 WHERE ID=A.当前科室id)) AS 科室 " & _
            "From 病人信息 A Where A.病人ID=[1] "
    
    Set rsTmp = zldatabase.OpenSQLRecord(strsql, Me.Caption, lng病人ID)
    
    If rsTmp.EOF Then Exit Function
    If rsTmp.RecordCount <> 1 Then Exit Function
    
    '住院信息
    txt病人ID.Text = lng病人ID
    txt姓名.Text = rsTmp!姓名
    
    '基本信息
    txt性别.Text = zlCommFun.Nvl(rsTmp("性别"))
    txt年龄.Text = zlCommFun.Nvl(rsTmp("年龄"))
    txt医疗付款.Text = zlCommFun.Nvl(rsTmp("医疗付款方式"))
    txt国籍.Text = zlCommFun.Nvl(rsTmp("国籍"))
    txt民族.Text = zlCommFun.Nvl(rsTmp("民族"))
    txt学历.Text = zlCommFun.Nvl(rsTmp("学历"))
    txt婚姻状况.Text = zlCommFun.Nvl(rsTmp("婚姻状况"))
    txt职业.Text = zlCommFun.Nvl(rsTmp("职业"))
    txt身份.Text = zlCommFun.Nvl(rsTmp("身份"))
    txt出生日期.Text = Format(zlCommFun.Nvl(rsTmp("出生日期")), "yyyy-mm-dd")
    txt身份证号.Text = zlCommFun.Nvl(rsTmp("身份证号"))
    txt出生地点.Text = zlCommFun.Nvl(rsTmp("出生地点"))
    txt家庭地址.Text = zlCommFun.Nvl(rsTmp("家庭地址"))
    txt家庭电话.Text = zlCommFun.Nvl(rsTmp("家庭电话"))
    txt户口邮编.Text = zlCommFun.Nvl(rsTmp("家庭地址邮编"))
    txt联系人姓名.Text = zlCommFun.Nvl(rsTmp("联系人姓名"))
    txt联系人关系.Text = zlCommFun.Nvl(rsTmp("联系人关系"))
    txt联系人地址.Text = zlCommFun.Nvl(rsTmp("联系人地址"))
    txt联系人电话.Text = zlCommFun.Nvl(rsTmp("联系人电话"))
    txt工作单位.Text = zlCommFun.Nvl(rsTmp("工作单位"))
    txt单位电话.Text = zlCommFun.Nvl(rsTmp("单位电话"))
    txt单位邮编.Text = zlCommFun.Nvl(rsTmp("单位邮编"))
    txt单位开户行.Text = zlCommFun.Nvl(rsTmp("单位开户行"))
    txt单位帐号.Text = zlCommFun.Nvl(rsTmp("单位帐号"))
    txt科室.Text = zlCommFun.Nvl(rsTmp("科室"))
    
    '费用信息
    txt担保人.Text = zlCommFun.Nvl(rsTmp("担保人"))
    txt担保额.Text = zlCommFun.Nvl(rsTmp("担保额"))

    If zlCommFun.Nvl(rsTmp("当前科室id"), 0) > 0 Then
        txt住院号.Text = zlCommFun.Nvl(rsTmp("住院号"))
        txt床号.Text = zlCommFun.Nvl(rsTmp("当前床号"))
        txt入院时间.Text = Format(zlCommFun.Nvl(rsTmp("入院时间")), "yyyy-MM-dd HH:mm")
        txt出院时间.Text = Format(zlCommFun.Nvl(rsTmp("出院时间")), "yyyy-MM-dd HH:mm")
    
        strsql = "Select B.费别 AS 住院费别 From 病人信息 A,病案主页 B " & _
                "Where A.病人ID=B.病人ID And A.主页id=B.主页ID And A.病人ID=[1]"
        Set rsTmp = zldatabase.OpenSQLRecord(strsql, Me.Caption, lng病人ID)
        If rsTmp.BOF = False Then txt费别.Text = zlCommFun.Nvl(rsTmp("住院费别"))
        
    Else
        lbl住院号.Caption = "门诊号"
        txt住院号.Text = zlCommFun.Nvl(rsTmp("门诊号"))
    End If
        
        
    strsql = "Select " & gConst_病人余额_列名 & "  From 病人余额 a Where 性质=1 And 病人ID= [1] "
    Set rsTmp = zldatabase.OpenSQLRecord(strsql, Me.Caption, lng病人ID)
    
    If Not rsTmp.EOF Then
        txt费用余额.Text = Format(zlCommFun.Nvl(rsTmp("费用余额")), "0.00")
        txt预交余额.Text = Format(zlCommFun.Nvl(rsTmp("预交余额")), "0.00")
    End If
    
    ReadCard = True
    
    Exit Function
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub CmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

