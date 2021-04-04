VERSION 5.00
Begin VB.Form frmDegreeCard 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病人信息卡片"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8850
   Icon            =   "frmDegreeCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   8850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picPati 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1260
      Left            =   75
      ScaleHeight     =   1260
      ScaleWidth      =   8730
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   75
      Width           =   8730
      Begin VB.TextBox txt年龄 
         Height          =   300
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   480
         Width           =   1170
      End
      Begin VB.TextBox txt姓名 
         Height          =   300
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   480
         Width           =   1170
      End
      Begin VB.TextBox txt病人ID 
         Height          =   300
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   90
         Width           =   1170
      End
      Begin VB.TextBox txt性别 
         Height          =   300
         Left            =   2985
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   480
         Width           =   1170
      End
      Begin VB.TextBox txt费别 
         Height          =   300
         Left            =   7350
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   480
         Width           =   1170
      End
      Begin VB.TextBox txt门诊号 
         Height          =   300
         Left            =   2985
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   90
         Width           =   1170
      End
      Begin VB.TextBox txt住院号 
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   90
         Width           =   1170
      End
      Begin VB.TextBox txt床号 
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   7350
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   90
         Width           =   1170
      End
      Begin VB.TextBox txt病区 
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   870
         Width           =   1170
      End
      Begin VB.TextBox txt科室 
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   2985
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   870
         Width           =   1170
      End
      Begin VB.TextBox txt入院时间 
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   870
         Width           =   1170
      End
      Begin VB.TextBox txt出院时间 
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   7350
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   870
         Width           =   1170
      End
      Begin VB.Label lbl病人ID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人ID"
         ForeColor       =   &H80000007&
         Height          =   180
         Left            =   525
         TabIndex        =   73
         Top             =   150
         Width           =   540
      End
      Begin VB.Label lbl费别 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "费别"
         Height          =   180
         Left            =   6945
         TabIndex        =   72
         Top             =   555
         Width           =   360
      End
      Begin VB.Label lbl年龄 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄"
         Height          =   180
         Left            =   4845
         TabIndex        =   71
         Top             =   540
         Width           =   360
      End
      Begin VB.Label lbl性别 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别"
         Height          =   180
         Left            =   2595
         TabIndex        =   70
         Top             =   540
         Width           =   360
      End
      Begin VB.Label lbl姓名 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名"
         Height          =   180
         Left            =   720
         TabIndex        =   69
         Top             =   540
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "门诊号"
         Height          =   180
         Left            =   2415
         TabIndex        =   68
         Top             =   150
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院号"
         Height          =   180
         Left            =   4665
         TabIndex        =   67
         Top             =   150
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "床号"
         Height          =   180
         Left            =   6945
         TabIndex        =   66
         Top             =   150
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "当前病区"
         Height          =   180
         Left            =   360
         TabIndex        =   65
         Top             =   930
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "科室"
         Height          =   180
         Left            =   2595
         TabIndex        =   64
         Top             =   930
         Width           =   360
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "入院时间"
         Height          =   180
         Left            =   4485
         TabIndex        =   63
         Top             =   930
         Width           =   720
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出院时间"
         Height          =   180
         Left            =   6600
         TabIndex        =   62
         Top             =   945
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "退出(&X)"
      Height          =   350
      Left            =   6960
      TabIndex        =   36
      Top             =   5880
      Width           =   1100
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4275
      Left            =   75
      ScaleHeight     =   4275
      ScaleWidth      =   8730
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   1380
      Width           =   8730
      Begin VB.TextBox txt户口地址 
         Height          =   300
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   76
         Top             =   3960
         Width           =   3015
      End
      Begin VB.TextBox txt户口邮编 
         Height          =   300
         Left            =   5280
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   75
         Top             =   3960
         Width           =   1170
      End
      Begin VB.TextBox txt医疗付款 
         Height          =   300
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   105
         Width           =   1170
      End
      Begin VB.TextBox txt身份证号 
         Height          =   300
         Left            =   1140
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   20
         Top             =   885
         Width           =   3015
      End
      Begin VB.TextBox txt家庭电话 
         Height          =   300
         Left            =   1140
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   24
         Top             =   1665
         Width           =   2000
      End
      Begin VB.TextBox txt单位帐号 
         Height          =   300
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   3225
         Width           =   3255
      End
      Begin VB.TextBox txt单位开户行 
         Height          =   300
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   3225
         Width           =   3015
      End
      Begin VB.TextBox txt单位邮编 
         Height          =   300
         Left            =   5280
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   31
         Top             =   2835
         Width           =   1170
      End
      Begin VB.TextBox txt单位电话 
         Height          =   300
         Left            =   1140
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   30
         Top             =   2835
         Width           =   2000
      End
      Begin VB.TextBox txt工作单位 
         Height          =   300
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   2445
         Width           =   3255
      End
      Begin VB.TextBox txt联系人电话 
         Height          =   300
         Left            =   1140
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   28
         Top             =   2445
         Width           =   2000
      End
      Begin VB.TextBox txt联系人地址 
         Height          =   300
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   2055
         Width           =   3255
      End
      Begin VB.TextBox txt联系人姓名 
         Height          =   300
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   1665
         Width           =   1170
      End
      Begin VB.TextBox txt家庭邮编 
         Height          =   300
         Left            =   5280
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   23
         Top             =   1275
         Width           =   1170
      End
      Begin VB.TextBox txt家庭地址 
         Height          =   300
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   1275
         Width           =   3015
      End
      Begin VB.TextBox txt出生地点 
         Height          =   300
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   885
         Width           =   3255
      End
      Begin VB.TextBox txt民族 
         Height          =   300
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   105
         Width           =   1170
      End
      Begin VB.TextBox txt学历 
         Height          =   300
         Left            =   7350
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   105
         Width           =   1170
      End
      Begin VB.TextBox txt国籍 
         Height          =   300
         Left            =   2985
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   105
         Width           =   1170
      End
      Begin VB.TextBox txt婚姻状况 
         Height          =   300
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   495
         Width           =   1170
      End
      Begin VB.TextBox txt职业 
         Height          =   300
         Left            =   2985
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   495
         Width           =   1170
      End
      Begin VB.TextBox txt身份 
         Height          =   300
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   495
         Width           =   1170
      End
      Begin VB.TextBox txt联系人关系 
         Height          =   300
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   2055
         Width           =   2000
      End
      Begin VB.TextBox txt出生日期 
         Height          =   300
         Left            =   7350
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   495
         Width           =   1170
      End
      Begin VB.TextBox txt担保人 
         Height          =   300
         Left            =   1140
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   34
         Top             =   3615
         Width           =   1980
      End
      Begin VB.TextBox txt担保额 
         Height          =   300
         Left            =   5280
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   35
         Top             =   3615
         Width           =   1170
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "户口地址"
         Height          =   180
         Left            =   360
         TabIndex        =   78
         Top             =   4020
         Width           =   720
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "户口邮编"
         Height          =   180
         Left            =   4440
         TabIndex        =   77
         Top             =   4005
         Width           =   720
      End
      Begin VB.Label lbl医疗付款 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "医疗付款"
         Height          =   180
         Left            =   360
         TabIndex        =   74
         Top             =   165
         Width           =   720
      End
      Begin VB.Label lbl单位帐号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位帐号"
         Height          =   180
         Left            =   4485
         TabIndex        =   60
         Top             =   3285
         Width           =   720
      End
      Begin VB.Label lbl单位开户行 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位开户行"
         Height          =   180
         Left            =   180
         TabIndex        =   59
         Top             =   3285
         Width           =   900
      End
      Begin VB.Label lbl单位邮编 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位邮编"
         Height          =   180
         Left            =   4485
         TabIndex        =   58
         Top             =   2895
         Width           =   720
      End
      Begin VB.Label lbl单位电话 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位电话"
         Height          =   180
         Left            =   360
         TabIndex        =   57
         Top             =   2895
         Width           =   720
      End
      Begin VB.Label lbl工作单位 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "工作单位"
         Height          =   180
         Left            =   4485
         TabIndex        =   56
         Top             =   2505
         Width           =   720
      End
      Begin VB.Label lbl联系人电话 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "联系人电话"
         Height          =   180
         Left            =   180
         TabIndex        =   55
         Top             =   2505
         Width           =   900
      End
      Begin VB.Label lbl联系人地址 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "联系人地址"
         Height          =   180
         Left            =   4305
         TabIndex        =   54
         Top             =   2115
         Width           =   900
      End
      Begin VB.Label lbl联系人关系 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "联系人关系"
         Height          =   180
         Left            =   180
         TabIndex        =   53
         Top             =   2115
         Width           =   900
      End
      Begin VB.Label lbl联系人姓名 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "联系人姓名"
         Height          =   180
         Left            =   4305
         TabIndex        =   52
         Top             =   1725
         Width           =   900
      End
      Begin VB.Label lbl户口邮编 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "家庭邮编"
         Height          =   180
         Left            =   4485
         TabIndex        =   51
         Top             =   1335
         Width           =   720
      End
      Begin VB.Label lbl家庭电话 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "家庭电话"
         Height          =   180
         Left            =   360
         TabIndex        =   50
         Top             =   1725
         Width           =   720
      End
      Begin VB.Label lbl家庭地址 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "家庭地址"
         Height          =   180
         Left            =   360
         TabIndex        =   49
         Top             =   1335
         Width           =   720
      End
      Begin VB.Label lvl婚姻状况 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "婚姻状况"
         Height          =   180
         Left            =   360
         TabIndex        =   48
         Top             =   555
         Width           =   720
      End
      Begin VB.Label lbl学历 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "学历"
         Height          =   180
         Left            =   6945
         TabIndex        =   47
         Top             =   165
         Width           =   360
      End
      Begin VB.Label lbl国籍 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "国籍"
         Height          =   180
         Left            =   2595
         TabIndex        =   46
         Top             =   165
         Width           =   360
      End
      Begin VB.Label lbl民族 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "民族"
         Height          =   180
         Left            =   4845
         TabIndex        =   45
         Top             =   165
         Width           =   360
      End
      Begin VB.Label lbl职业 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "职业"
         Height          =   180
         Left            =   2595
         TabIndex        =   44
         Top             =   555
         Width           =   360
      End
      Begin VB.Label lbl身份 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "身份"
         Height          =   180
         Left            =   4845
         TabIndex        =   43
         Top             =   555
         Width           =   360
      End
      Begin VB.Label lbl身份证号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "身份证号"
         Height          =   180
         Left            =   360
         TabIndex        =   42
         Top             =   945
         Width           =   720
      End
      Begin VB.Label lbl出生地点 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出生地点"
         Height          =   180
         Left            =   4485
         TabIndex        =   41
         Top             =   945
         Width           =   720
      End
      Begin VB.Label lbl出生日期 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出生日期"
         Height          =   180
         Left            =   6585
         TabIndex        =   40
         Top             =   555
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "担保人"
         Height          =   180
         Left            =   540
         TabIndex        =   39
         Top             =   3675
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "担保额"
         Height          =   180
         Left            =   4665
         TabIndex        =   38
         Top             =   3675
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
Public mlng病人ID As Long '要修改或查看的病人ID
Private mblnUnload As Boolean

Private Function ReadCard(lngID As Long) As Boolean
'功能：读取指定病人信息,并显示在界面上
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
        
    strSQL = "Select A.*,B.名称 as 当前病区,C.名称 as 当前科室" & _
        " From 病人信息 A,部门表 B,部门表 C" & _
        " Where A.当前病区ID=B.ID(+) And A.当前科室ID=C.ID(+) And A.病人ID=[1]"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, lngID)
    
    If rsTmp.EOF Then Exit Function
    If rsTmp.RecordCount <> 1 Then Exit Function
    
    txt病人ID.Text = lngID
    txt姓名.Text = rsTmp!姓名
    
    txt门诊号.Text = IIf(IsNull(rsTmp!门诊号), "", rsTmp!门诊号)
    txt住院号.Text = IIf(IsNull(rsTmp!住院号), "", rsTmp!住院号)
    txt床号.Text = IIf(IsNull(rsTmp!当前床号), "", rsTmp!当前床号)
    txt病区.Text = IIf(IsNull(rsTmp!当前病区), "", rsTmp!当前病区)
    txt科室.Text = IIf(IsNull(rsTmp!当前科室), "", rsTmp!当前科室)
    txt入院时间.Text = Format(IIf(IsNull(rsTmp!入院时间), "", rsTmp!入院时间), "yyyy-MM-dd")
    txt出院时间.Text = Format(IIf(IsNull(rsTmp!出院时间), "", rsTmp!出院时间), "yyyy-MM-dd")
    
    txt性别.Text = IIf(IsNull(rsTmp!性别), "", rsTmp!性别)
    txt年龄.Text = IIf(IsNull(rsTmp!年龄), "", rsTmp!年龄)
    txt费别.Text = IIf(IsNull(rsTmp!费别), "", rsTmp!费别)
    txt医疗付款.Text = IIf(IsNull(rsTmp!医疗付款方式), "", rsTmp!医疗付款方式)
    txt国籍.Text = IIf(IsNull(rsTmp!国籍), "", rsTmp!国籍)
    txt民族.Text = IIf(IsNull(rsTmp!民族), "", rsTmp!民族)
    txt学历.Text = IIf(IsNull(rsTmp!学历), "", rsTmp!学历)
    txt婚姻状况.Text = IIf(IsNull(rsTmp!婚姻状况), "", rsTmp!婚姻状况)
    txt职业.Text = IIf(IsNull(rsTmp!职业), "", rsTmp!职业)
    txt身份.Text = IIf(IsNull(rsTmp!身份), "", rsTmp!身份)
    txt出生日期.Text = Format(IIf(IsNull(rsTmp!出生日期), "", rsTmp!出生日期), "yyyy-MM-dd")
    txt身份证号.Text = IIf(IsNull(rsTmp!身份证号), "", rsTmp!身份证号)
    txt出生地点.Text = IIf(IsNull(rsTmp!出生地点), "", rsTmp!出生地点)
    txt家庭地址.Text = IIf(IsNull(rsTmp!家庭地址), "", rsTmp!家庭地址)
    txt家庭电话.Text = IIf(IsNull(rsTmp!家庭电话), "", rsTmp!家庭电话)
    txt户口邮编.Text = IIf(IsNull(rsTmp!家庭地址邮编), "", rsTmp!家庭地址邮编)
    txt联系人姓名.Text = IIf(IsNull(rsTmp!联系人姓名), "", rsTmp!联系人姓名)
    txt联系人关系.Text = IIf(IsNull(rsTmp!联系人关系), "", rsTmp!联系人关系)
    txt联系人地址.Text = IIf(IsNull(rsTmp!联系人地址), "", rsTmp!联系人地址)
    txt联系人电话.Text = IIf(IsNull(rsTmp!联系人电话), "", rsTmp!联系人电话)
    txt工作单位.Text = IIf(IsNull(rsTmp!工作单位), "", rsTmp!工作单位)
    txt单位电话.Text = IIf(IsNull(rsTmp!单位电话), "", rsTmp!单位电话)
    txt单位邮编.Text = IIf(IsNull(rsTmp!单位邮编), "", rsTmp!单位邮编)
    txt单位开户行.Text = IIf(IsNull(rsTmp!单位开户行), "", rsTmp!单位开户行)
    txt单位帐号.Text = IIf(IsNull(rsTmp!单位帐号), "", rsTmp!单位帐号)
    txt担保人.Text = IIf(IsNull(rsTmp!担保人), "", rsTmp!担保人)
    txt担保额.Text = IIf(IsNull(rsTmp!担保额), "", rsTmp!担保额)
    txt户口地址.Text = IIf(IsNull(rsTmp!户口地址), "", rsTmp!户口地址)
    txt家庭邮编.Text = IIf(IsNull(rsTmp!家庭地址邮编), "", rsTmp!家庭地址邮编)
    '74428：李南春，2014-7-8，病人姓名显示颜色处理
    Call SetPatiColor(txt姓名, Nvl(rsTmp!病人类型), IIf(IsNull(rsTmp!险类), Me.ForeColor, vbRed))
    
    ReadCard = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmdExit_Click()
    Unload Me
End Sub



Private Sub Form_Activate()
    If mblnUnload Then Unload Me: Exit Sub
    cmdExit.SetFocus
End Sub

Private Sub Form_Load()
    mblnUnload = False
    If Not ReadCard(mlng病人ID) Then
        MsgBox "不能正确读取病人信息,请与系统管理员联系！", vbExclamation, gstrSysName
        mblnUnload = True
    End If
    
    zlcontrol.PicShowFlat picInfo, -1, , taCenterAlign
    zlcontrol.PicShowFlat picPati, -1, , taCenterAlign
End Sub
