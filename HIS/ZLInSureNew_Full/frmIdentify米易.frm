VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmIdentify米易 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "身份验证"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6825
   Icon            =   "frmIdentify米易.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txt报销限额 
      Enabled         =   0   'False
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
      Left            =   1800
      TabIndex        =   32
      Top             =   5100
      Width           =   2865
   End
   Begin VB.TextBox txt顺序号 
      Enabled         =   0   'False
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
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   1050
      Visible         =   0   'False
      Width           =   2865
   End
   Begin VB.TextBox txt报销比例 
      Enabled         =   0   'False
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
      Left            =   1815
      TabIndex        =   30
      Top             =   4650
      Width           =   2865
   End
   Begin VB.TextBox txt起付线 
      Enabled         =   0   'False
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
      Left            =   1800
      TabIndex        =   28
      Top             =   4200
      Width           =   2865
   End
   Begin VB.TextBox txt原密码 
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
      IMEMode         =   3  'DISABLE
      Left            =   1800
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   150
      Width           =   2865
   End
   Begin VB.CommandButton cmd病种 
      Caption         =   "…"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4290
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt病种 
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
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   2505
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "修改密码(M)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5160
      TabIndex        =   36
      Top             =   5010
      Width           =   1470
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5190
      TabIndex        =   35
      Top             =   900
      Width           =   1380
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5190
      TabIndex        =   34
      Top             =   330
      Width           =   1380
   End
   Begin VB.Frame Frame1 
      Height          =   8115
      Left            =   4920
      TabIndex        =   33
      Top             =   -450
      Width           =   30
   End
   Begin VB.TextBox txt帐户余额 
      Enabled         =   0   'False
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
      Left            =   1800
      TabIndex        =   26
      Top             =   3750
      Width           =   2865
   End
   Begin VB.TextBox txt单位名称 
      Enabled         =   0   'False
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
      Left            =   1800
      TabIndex        =   24
      Top             =   3300
      Width           =   2865
   End
   Begin VB.TextBox txt人员类别 
      Enabled         =   0   'False
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
      Left            =   1800
      TabIndex        =   22
      Top             =   2850
      Width           =   2865
   End
   Begin VB.TextBox txt年龄 
      Enabled         =   0   'False
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
      Left            =   3870
      TabIndex        =   16
      Top             =   1500
      Width           =   795
   End
   Begin VB.TextBox txt身份证号 
      Enabled         =   0   'False
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
      Left            =   1800
      TabIndex        =   20
      Top             =   2400
      Width           =   2865
   End
   Begin VB.ComboBox cbo性别 
      Enabled         =   0   'False
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
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   1500
      Width           =   885
   End
   Begin VB.TextBox txt姓名 
      Enabled         =   0   'False
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
      Left            =   1800
      TabIndex        =   12
      Top             =   1050
      Width           =   2865
   End
   Begin VB.TextBox txt就诊编号 
      Enabled         =   0   'False
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
      Left            =   1800
      TabIndex        =   10
      Top             =   600
      Width           =   2865
   End
   Begin VB.TextBox txt卡号 
      Enabled         =   0   'False
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
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   2865
   End
   Begin MSMask.MaskEdBox txt出生日期 
      Height          =   345
      Left            =   1800
      TabIndex        =   18
      Top             =   1950
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   609
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "####-##-##"
      PromptChar      =   "_"
   End
   Begin VB.Label lbl报销限额 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "报销限额(&F)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   330
      TabIndex        =   31
      Top             =   5160
      Width           =   1425
   End
   Begin VB.Label lbl顺序号 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "个人编号(&X)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   330
      TabIndex        =   7
      Top             =   1110
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label lbl报销比例 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "报销比例(&L)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   345
      TabIndex        =   29
      Top             =   4710
      Width           =   1425
   End
   Begin VB.Label lbl起付线 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "起付线(&Q)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   585
      TabIndex        =   27
      Top             =   4260
      Width           =   1170
   End
   Begin VB.Label lbl原密码 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "密码(&O)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   825
      TabIndex        =   0
      Top             =   210
      Width           =   915
   End
   Begin VB.Label lbl病种 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "病种(I)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   840
      TabIndex        =   2
      Top             =   660
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label lbl帐户余额 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "帐户余额(&D)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   330
      TabIndex        =   25
      Top             =   3810
      Width           =   1425
   End
   Begin VB.Label lbl单位名称 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "单位名称(&W)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   330
      TabIndex        =   23
      Top             =   3360
      Width           =   1425
   End
   Begin VB.Label lbl人员类别 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "人员类别(&T)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   330
      TabIndex        =   21
      Top             =   2910
      Width           =   1425
   End
   Begin VB.Label lbl年龄 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "年龄(&G)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2910
      TabIndex        =   15
      Top             =   1560
      Width           =   915
   End
   Begin VB.Label lbl身份证号 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "身份证号(&K)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   330
      TabIndex        =   19
      Top             =   2460
      Width           =   1425
   End
   Begin VB.Label lbl出生日期 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "出生日期(&B)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   330
      TabIndex        =   17
      Top             =   1995
      Width           =   1425
   End
   Begin VB.Label lbl性别 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "性别(&S)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   840
      TabIndex        =   13
      Top             =   1560
      Width           =   915
   End
   Begin VB.Label lbl姓名 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "姓名(&N)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   840
      TabIndex        =   11
      Top             =   1110
      Width           =   915
   End
   Begin VB.Label lbl就诊编号 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "就诊编号(&R)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   330
      TabIndex        =   9
      Top             =   660
      Width           =   1425
   End
   Begin VB.Label lbl卡号 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "卡号(&A)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   840
      TabIndex        =   5
      Top             =   660
      Visible         =   0   'False
      Width           =   915
   End
End
Attribute VB_Name = "frmIdentify米易"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng病人ID As Long
Private mbytType As Long
Private mstrReturn As String
Private rsTemp As New ADODB.Recordset

Public mstr新密码 As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdModify_Click()
    mstr新密码 = frm修改密码.ChangePassword(txt原密码)
End Sub

Private Sub cmdOK_Click()
    Dim strIdentify As String, strAddition As String
    Dim lngResult As Long
    
    If Trim(txt顺序号) = "" Or Trim(txt卡号) = "" Then
        MsgBox "请插入卡并输入密码后按回车！", vbInformation, gstrSysName
        txt原密码.SetFocus
        Exit Sub
    End If
    
    If Trim(txt就诊编号.Text) = "" Then
        MsgBox "未得到就诊编号，无法继续！", vbInformation, gstrSysName
        txt原密码.SetFocus
        Exit Sub
    End If
    
    '检查病人状态
    gstrSQL = "select nvl(当前状态,0) as 状态 from 保险帐户 where 险类=[1] and 医保号=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, type_米易, CStr(txt就诊编号.Text))
    
    If rsTemp.RecordCount > 0 Then
        If rsTemp("状态") > 0 Then
            MsgBox "该病人已经在院，不能通过身份验证。", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    '构成字符串
    '建立病人档案信息，传入格式：
    '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);
    '8.中心代码;9.顺序号;10人员身份;11帐户余额;12当前状态;13病种ID;14在职(0,1);15退休证号;16年龄段;17灰度级
    '18帐户增加累计,19帐户支出累计,20进入统筹累计,21统筹报销累计,22住院次数累计
    strIdentify = gComInfo_米易.卡号                         '0卡号
    strIdentify = strIdentify & ";" & gComInfo_米易.个人编号     '1医保号（个人编号）
    strIdentify = strIdentify & ";" & gComInfo_米易.密码                              '2密码
    strIdentify = strIdentify & ";" & txt姓名.Text              '3姓名
    strIdentify = strIdentify & ";" & cbo性别.Text              '4性别
    strIdentify = strIdentify & ";" & txt出生日期.Text          '5出生日期
    strIdentify = strIdentify & ";" & txt身份证号.Text          '6身份证
    strIdentify = strIdentify & ";" & txt单位名称.Text          '7.单位名称(编码)
    strAddition = ";0"                                          '8.中心代码
    strAddition = strAddition & ";" & txt就诊编号.Text                              '9.顺序号
    strAddition = strAddition & ";" & txt人员类别.Tag           '10人员身份
    strAddition = strAddition & ";" & Val(txt帐户余额.Text)     '11帐户余额
    strAddition = strAddition & ";0"                            '12当前状态
    strAddition = strAddition & ";" & Val(txt病种.Tag)           '13病种ID
    strAddition = strAddition & ";1"                            '14在职(1,2,3)
    strAddition = strAddition & ";"                             '15退休证号
    strAddition = strAddition & ";" & Val(txt年龄.Text)         '16年龄段
    strAddition = strAddition & ";"                             '17灰度级
    strAddition = strAddition & ";" & Val(txt帐户余额.Text)     '18帐户增加累计
    strAddition = strAddition & ";0"                            '19帐户支出累计
    strAddition = strAddition & ";0"                            '20上年工资总额
    strAddition = strAddition & ";0"                            '21住院次数累计
    
    mlng病人ID = BuildPatiInfo(0, strIdentify & strAddition, mlng病人ID, type_米易)
    '返回格式:中间插入病人ID
    If mlng病人ID > 0 Then
        mstrReturn = strIdentify & ";" & mlng病人ID & strAddition
    End If
    
    If mstr新密码 <> "" Then
        '调用修改密码接口，不成功给出提示并继续
        gComInfo_米易.新密码 = mstr新密码
        gstrPara_米易 = "<code>" & gComInfo_米易.卡号 & "</code>" & GetParaCode(密码, gComInfo_米易.密码) & _
            GetParaCode(新密码, gComInfo_米易.新密码)
        If 调用接口_米易("modifypassword") Then
            gComInfo_米易.密码 = mstr新密码
            '更新个人帐户中的信息
            gstrSQL = "zl_保险帐户_更新信息(" & mlng病人ID & "," & type_米易 & ",'密码','''" & mstr新密码 & "''')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "更新密码")
            
            '写卡:0 as 修改正确，4 as  写卡错误, 3 as 转换数据错误,2 as  读卡错误,1 as 输入原来口令不正确
            '先调用的米易接口，如果正确执行，则说明原密码没错
            lngResult = Card_ChangePsd(gintComPort, txt原密码.Text, mstr新密码)
            If lngResult <> 0 Then
                gComInfo_米易.密码 = txt原密码.Text
                MsgBox "密码修改失败，仍可继续操作！", vbInformation, gstrSysName
            Else
                gComInfo_米易.密码 = mstr新密码
            End If
        End If
    Else
        gComInfo_米易.密码 = txt原密码.Text
    End If
    
    Unload Me
End Sub

Private Sub cmd病种_Click()
    Dim rs病种 As ADODB.Recordset
    
    gstrSQL = " Select A.ID,A.编码,A.名称,A.简码,decode(A.类别,1,'慢性病',2,'特种病','普通病') as 类别 " & _
            " From 保险病种 A where A.险类=[1]"
    Set rs病种 = New ADODB.Recordset
    Set rs病种 = zlDatabase.OpenSQLRecord(gstrSQL, "身份验证", type_米易)
    If rs病种.RecordCount > 0 Then
        If frmListSel.ShowSelect(type_米易, rs病种, "ID", "医保病种选择", "请选择医保病种：") = True Then
            txt病种.Text = rs病种("名称")
            txt病种.Tag = rs病种("ID")
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    With cbo性别
        .AddItem "男"
        .AddItem "女"
        .ListIndex = 0
    End With
    
    If mbytType = 0 Then
        '门诊
        Me.Height = 4545
        Me.cmdModify.Top = 3350
    End If
    
    If mlng病人ID <> 0 Then Call ReadPatient
End Sub

Private Sub ClearCons()
    '清除垃圾数据
    txt卡号.Text = ""
    txt顺序号.Text = ""
    txt就诊编号 = ""
    txt姓名 = ""
    cbo性别.ListIndex = 0
    txt年龄 = ""
    txt出生日期 = "____-__-__"
    txt身份证号 = ""
    txt人员类别 = ""
    txt单位名称 = ""
    txt帐户余额 = ""
    txt起付线 = ""
    txt报销比例 = ""
    txt报销限额 = ""
End Sub

Private Sub ReadPatient()
    Dim intField As Integer
    '读取该病人的详细信息
    If mlng病人ID = 0 Then Exit Sub
    gstrSQL = " Select A.卡号,A.医保号,A.密码,B.姓名,Decode(B.性别,'女',1,0) 性别,nvl(B.年龄,0) 年龄,Nvl(A.病种ID,0) 病种ID,C.名称 病种," & _
              " A.顺序号 就诊编号,to_Char(B.出生日期,'yyyy-MM-dd') 出生日期,B.身份证号,B.工作单位 单位名称,Nvl(A.帐户余额,0) 帐户余额,A.在职 人员类别,A.备注" & _
              " From 保险帐户 A,病人信息 B,保险病种 C" & _
              " Where A.病人ID=B.病人ID And A.病种ID=C.ID(+) And A.病人ID=[1] And A.险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取帐户信息", mlng病人ID, type_米易)
    
    With rsTemp
        If .EOF Then Exit Sub
        '只读出基本信息，这样操作员必须有卡才行。无法作假
'        txt卡号 = !卡号
'        txt顺序号 = !医保号
'        txt病种.Tag = !病种ID
'        txt病种 = !病种
'        txt就诊编号 = NVL(!顺序号, "")
        txt姓名 = Nvl(!姓名, "")
        cbo性别.ListIndex = !性别
        txt年龄 = Format(!年龄, "#####0;#####0; ;")
        txt出生日期 = !出生日期
        txt身份证号 = !身份证号
        txt人员类别 = 转换人员类别(!人员类别)
        txt单位名称 = Nvl(!单位名称, "")
'        txt帐户余额 = Format(!帐户余额, "#####0;#####0; ;")
'        If Not IsNull(!备注) Then
'            '起付线;报销比例;报销限额
'            txt起付线 = Format(Split(!备注, ";")(0), "#####0.00;-#####0.00; ;")
'            txt报销比例 = Format(Split(!备注, ";")(1), "#####0.00;-#####0.00; ;")
'            txt报销限额 = Format(Split(!备注, ";")(2), "#####0.00;-#####0.00; ;")
'        End If
        
'        gComInfo_米易.卡号 = txt卡号
'        gComInfo_米易.个人编号 = txt顺序号
'        gComInfo_米易.就诊编号 = txt就诊编号
        gComInfo_米易.姓名 = txt姓名
        gComInfo_米易.性别 = Me.cbo性别.ListIndex + 1
        gComInfo_米易.年龄 = txt年龄
        gComInfo_米易.出生日期 = txt出生日期
        gComInfo_米易.身份证号 = txt身份证号
        gComInfo_米易.单位名称 = txt单位名称
'        gComInfo_米易.帐户余额 = Val(txt帐户余额)
'        gComInfo_米易.本次起付线 = Val(txt起付线)
'        gComInfo_米易.本次统筹报销比例 = Val(txt报销比例)
'        gComInfo_米易.本次统筹限额 = Val(txt报销限额)
        gComInfo_米易.人员类别 = !人员类别
'        Call 获取病种编码(Val(txt病种.Tag))
    End With
End Sub

Public Function GetPatient(ByVal bytType As Byte, Optional ByVal lng病人ID As Long = 0) As String
    mstrReturn = ""
    mlng病人ID = lng病人ID
    mbytType = bytType
    Me.Show 1
    
    GetPatient = mstrReturn
End Function

Private Sub txt病种_KeyPress(KeyAscii As Integer)
    Dim rsTemp As New ADODB.Recordset, strText As String
    Dim blnReturn As Boolean, lst As ListItem
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If txt病种.Text = "" Or txt病种.Tag <> "" Then
        Exit Sub
    End If
    
    On Error GoTo errHandle
    
    strText = txt病种.Text
    gstrSQL = "Select A.ID,A.编码,A.名称,A.简码,decode(A.类别,1,'慢性病',2,'特殊病','普通病') 类别 " & _
             "   FROM 保险病种 A WHERE A.险类=[1] And (A.编码 like [2] ||'%' or A.名称 like [2] ||'%' or A.简码 like [2] ||'%')"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, type_米易, strText)
    
    If rsTemp.RecordCount > 0 Then
        '出现选择器
        If rsTemp.RecordCount > 1 Then
            '对于字段大于3的，即使只有一条记录把该对话框显示出来，以便让用户得到更多的信息
            blnReturn = frmListSel.ShowSelect(type_米易, rsTemp, "ID", "医保病种选择", "请选择特定的医保病种：")
        Else
            blnReturn = True
        End If
    End If
    
    If blnReturn = False Then
        '记录集中没有可选择的数据
        zlControl.TxtSelAll txt病种
        Exit Sub
    Else
        '肯定是有记录集的
        txt病种.Text = rsTemp("名称")
        txt病种.Tag = rsTemp("ID")
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub WriteFace()
    '填充界面
    txt卡号.Text = gComInfo_米易.卡号
    txt顺序号.Text = gComInfo_米易.个人编号
    txt就诊编号.Text = gComInfo_米易.就诊编号
    txt姓名.Text = gComInfo_米易.姓名
    cbo性别.ListIndex = gComInfo_米易.性别 - 1
    txt年龄 = gComInfo_米易.年龄
    txt出生日期 = gComInfo_米易.出生日期
    txt身份证号 = gComInfo_米易.身份证号
    txt人员类别 = 转换人员类别(gComInfo_米易.人员类别)
    txt单位名称 = gComInfo_米易.单位名称
    txt帐户余额 = Format(gComInfo_米易.帐户余额, "#####0.00;-#####0.00; ;")
    txt起付线 = gComInfo_米易.本次起付线
    txt报销比例 = gComInfo_米易.本次统筹报销比例
    txt报销限额 = gComInfo_米易.本次统筹限额
End Sub

Private Function 卡解析() As Boolean
    Dim recode As Integer, strResult As String
    strResult = Card_userinfo(gintComPort, txt原密码.Text, recode)
    If recode <> 0 Then
        MsgBox strResult & "获取参保病人信息错误，可能是密码错误！", vbInformation, gstrSysName
        Exit Function
    End If

    gComInfo_米易.个人编号 = Mid(strResult, 11, 15)
    gComInfo_米易.卡号 = Mid(strResult, 1, 10)
    gComInfo_米易.帐户余额 = Val(Mid(strResult, 26, 8)) / 100                   '将分为单位记录的金额转换为以元为单位的金额
    gComInfo_米易.密码 = txt原密码.Text
    卡解析 = True
End Function

Private Function 转换人员类别(ByVal str人员类别 As String) As String
    Select Case str人员类别
    Case 11, 12
        转换人员类别 = IIf(str人员类别 = "11", "在职", "在职长期驻外")
    Case 21, 22
        转换人员类别 = IIf(str人员类别 = "21", "退休", "退休异地安置")
    Case Else
        转换人员类别 = "离休"
    End Select
End Function

Private Sub txt原密码_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(txt原密码.Text) = "" Then
        MsgBox "请输入密码！", vbInformation, gstrSysName
        txt原密码.SetFocus
        Exit Sub
    End If
    
    '清除界面并解析
    Call ClearCons
    If Not 卡解析 Then Exit Sub
    If Trim(gComInfo_米易.个人编号) <> "" And Trim(gComInfo_米易.卡号) <> "" Then
        '调用身份验证接口
        If mbytType = 0 Then     '门诊调用身份验证接口
            gstrPara_米易 = GetParaCode(个人编号, gComInfo_米易.个人编号) & GetParaCode(卡号, gComInfo_米易.卡号) & _
                GetParaCode(服务机构编号, gComInfo_米易.服务机构编号)
            If Not 调用接口_米易("identifyinfogetting") Then Exit Sub
        Else                    '住院调用待遇审批接口
            If Not 调用接口_米易("getsysdate") Then Exit Sub
            
            gComInfo_米易.支付类别 = "0301"    '普通住院
            Call 获取病种编码(Val(txt病种.Tag))
            
            gstrPara_米易 = GetParaCode(个人编号, gComInfo_米易.个人编号) & GetParaCode(卡号, gComInfo_米易.卡号) & _
                GetParaCode(密码, gComInfo_米易.密码) & GetParaCode(服务机构编号, gComInfo_米易.服务机构编号) & _
                GetParaCode(支付类别, gComInfo_米易.支付类别) & GetParaCode(病种编码, gComInfo_米易.病种编码) & _
                GetParaCode(待遇获取时间, gComInfo_米易.系统时间)
            If Not 调用接口_米易("audittreatment") Then Exit Sub
        End If
        Call WriteFace
    End If
End Sub
