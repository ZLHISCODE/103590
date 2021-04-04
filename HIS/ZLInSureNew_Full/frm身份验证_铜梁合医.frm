VERSION 5.00
Begin VB.Form frm身份验证_铜梁合医 
   Caption         =   "铜梁合医身份验证"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8010
   Icon            =   "frm身份验证_铜梁合医.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   331
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   534
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Cmd_ok 
      Caption         =   "确定(&O)"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   25
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton Cmd_cancle 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   24
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "就诊信息"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   16
      Top             =   3120
      Width           =   7815
      Begin VB.CheckBox Chk慢性病 
         Caption         =   "是否门诊慢性病"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   248
         Width           =   1695
      End
      Begin VB.CheckBox Chk组织体验 
         Caption         =   "是否农合办组织体验"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   608
         Width           =   1935
      End
      Begin VB.TextBox Txt病种名称 
         Height          =   270
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox Txt病种代码 
         Height          =   270
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   240
         Width           =   2295
      End
      Begin VB.CheckBox Chk预防接种 
         Caption         =   "是否预防接种"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2640
         TabIndex        =   33
         Top             =   608
         Width           =   1575
      End
      Begin VB.ComboBox Cob状态 
         Height          =   300
         Left            =   3120
         TabIndex        =   32
         Text            =   "一般"
         Top             =   225
         Width           =   855
      End
      Begin VB.TextBox Txt诊断结果 
         BackColor       =   &H80000009&
         Height          =   270
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   16
         TabIndex        =   26
         Top             =   960
         Width           =   4935
      End
      Begin VB.Label Label16 
         Caption         =   "病种名称"
         Height          =   270
         Left            =   4560
         TabIndex        =   40
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label15 
         Caption         =   "病种代码"
         Height          =   270
         Left            =   4560
         TabIndex        =   39
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label17 
         Caption         =   "状态:"
         Height          =   255
         Left            =   2640
         TabIndex        =   38
         Top             =   248
         Width           =   495
      End
      Begin VB.Label Label14 
         Caption         =   "诊断结果"
         Height          =   270
         Left            =   120
         TabIndex        =   27
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "病人信息"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      Begin VB.TextBox Txt机构代码 
         BackColor       =   &H80000009&
         Height          =   270
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox Txt合医编码 
         Height          =   270
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox Txt住址 
         BackColor       =   &H80000009&
         Height          =   270
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   2640
         Width           =   4575
      End
      Begin VB.TextBox Txt姓名 
         BackColor       =   &H80000009&
         Height          =   270
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox Txt性别 
         BackColor       =   &H80000009&
         Height          =   270
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox Txt出生日期 
         BackColor       =   &H80000009&
         Height          =   270
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   2280
         Width           =   2055
      End
      Begin VB.TextBox Txt户主姓名 
         BackColor       =   &H80000009&
         Height          =   270
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox Txt户主性别 
         BackColor       =   &H80000009&
         Height          =   270
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1875
         Width           =   2055
      End
      Begin VB.TextBox Txt帐户余额 
         BackColor       =   &H80000009&
         Height          =   270
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox Txt身份证号 
         BackColor       =   &H80000009&
         Height          =   270
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox Txt家庭帐号 
         BackColor       =   &H80000009&
         Height          =   270
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox Txt户主身份证 
         BackColor       =   &H80000009&
         Height          =   270
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox Txt户主关系 
         BackColor       =   &H80000009&
         Height          =   270
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   1875
         Width           =   1815
      End
      Begin VB.Label Label13 
         Caption         =   "医疗机构代码"
         Height          =   255
         Left            =   3480
         TabIndex        =   31
         Top             =   375
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "合医编码"
         Height          =   270
         Left            =   240
         TabIndex        =   30
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "住    址"
         Height          =   270
         Left            =   240
         TabIndex        =   23
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "户主性别"
         Height          =   270
         Left            =   240
         TabIndex        =   21
         Top             =   1875
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "户主姓名"
         Height          =   270
         Left            =   240
         TabIndex        =   20
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "出生日期"
         Height          =   270
         Left            =   240
         TabIndex        =   19
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "性    别"
         Height          =   270
         Left            =   240
         TabIndex        =   18
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "姓    名"
         Height          =   270
         Left            =   240
         TabIndex        =   17
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "身份证号"
         Height          =   255
         Left            =   3480
         TabIndex        =   15
         Top             =   735
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "家庭帐号"
         Height          =   255
         Left            =   3480
         TabIndex        =   14
         Top             =   1095
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "户主身份证号"
         Height          =   255
         Left            =   3480
         TabIndex        =   13
         Top             =   1455
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "户主关系"
         Height          =   255
         Left            =   3480
         TabIndex        =   12
         Top             =   1875
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "帐户余额"
         Height          =   255
         Left            =   3480
         TabIndex        =   11
         Top             =   2295
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frm身份验证_铜梁合医"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private 返回串1 As String, 返回串2 As String, 流水号 As String, C确定 As Boolean, b就医类别 As Byte
Private R姓名 As String, R性别 As String, R身份证号 As String, R家庭帐号 As String, var机构代码 As String, var机构名称 As String
Private R出生日期 As String, R户主姓名 As String, R户主性别 As String, R合医编码 As String
Private R户主身份证 As String, R住址 As String, R户主关系 As String, R帐户余额 As Double, Str合医信息 As String
Private R病种代码 As String, R病种名称 As String

Public Function GetIdentify(Optional bytType As Byte, Optional lng病人ID As Long = 0, Optional ByRef intinsure As Integer = 0) As String
C确定 = False
b就医类别 = bytType
    Me.Show 1
    If C确定 Then
        lng病人ID = BuildPatiInfo(bytype, 返回串1 & 返回串2, lng病人ID, type_铜梁合医)
        gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & intinsure & ",'合医信息','''" & Str合医信息 & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存门诊诊断等")
        gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & intinsure & ",'帐户余额','" & R帐户余额 & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "更新余额")
        gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & intinsure & ",'单位编码','" & Trim(MidUni(R家庭帐号, 1, 20)) & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存家庭帐号在保险帐户的单位编码中")
        GetIdentify = 返回串1 & ";" & lng病人ID & 返回串2
        gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & intinsure & ",'病种代码','''" & Trim(MidUni(R病种代码, 1, 20)) & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存病种代码在保险帐户中")
        gstrSQL = "zl_保险帐户_更新信息(" & lng病人ID & "," & intinsure & ",'病种名称','''" & Trim(MidUni(R病种名称, 1, 200)) & "''')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存病种名称在保险帐户中")
    Else
        GetIdentify = ""
    End If
End Function

Private Sub Cmd_cancle_Click()
Unload Me
End Sub
Private Sub Cmd_ok_Click()
    If Me.Txt合医编码.Text = "" Then
        MsgBox "合医编号不能为空", vbInformation, gstrSysName
        Me.Txt合医编码.SetFocus
        Exit Sub
    End If
    
    '构造返回串
    If b就医类别 = 0 Then '门诊
        If Me.Chk慢性病.Value = 1 And (Trim(Me.txt病种代码.Text) = "" Or Trim(Me.txt病种名称.Text) = "") Then
            MsgBox "慢性病必须选择病种代码", vbInformation, gstrSysName
            Me.txt病种代码.SetFocus
            Exit Sub
        End If
    ElseIf b就医类别 = 1 Then '住院
        If Trim(Me.txt病种代码.Text) = "" Or Trim(Me.txt病种名称.Text) = "" Then
            MsgBox "住院必须获取病种代码和病种名称", vbInformation, gstrSysName
            Me.txt病种代码.SetFocus
            Exit Sub
        End If
    End If
    返回串1 = Me.Txt合医编码.Text & _
            ";" & Me.Txt合医编码.Text & _
            ";;" & Me.txt姓名.Text & _
            ";" & Me.txt性别.Text & _
            ";" & Mid(Me.txt出生日期.Text, 1, 4) & "-" & Mid(Me.txt出生日期.Text, 6, 2) & "-" & Mid(Me.txt出生日期.Text, 9, 2) & _
            ";" & Trim(Me.txt身份证号.Text) & _
            ";" & Me.Txt住址.Text
    返回串2 = ";" & gintInsure & _
            ";" & 流水号 & _
            ";" & MidUni(Me.Txt户主关系.Text, 1, 8) & _
            ";" & Me.txt帐户余额.Text & _
            ";0;;1;" & Me.Txt户主身份证 & _
            ";;;;;;;;;"
    Str合医信息 = Trim(Me.Txt机构代码.Text) & "|" & IIf(Me.Chk组织体验 = 0, 0, 1) & "|" & IIf(Me.Chk慢性病.Value = 0, 0, 1) & "|" & IIf(Me.Chk预防接种.Value = 0, 0, 1) & "|" & Trim(Me.Cob状态.Text) & "|" & MidUni(Trim(Me.Txt诊断结果.Text), 1, 16)
    C确定 = True
    Unload Me
End Sub


Private Sub Form_Load()
    Me.Cob状态.AddItem ("一般")
    Me.Cob状态.AddItem ("危")
    Me.Cob状态.AddItem ("急")
    Me.Cob状态.AddItem ("其它")
End Sub

Private Sub Txt病种代码_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
    R病种代码 = Space(30)
    R病种名称 = Space(200)
        If GetBZDM(R病种代码, R病种名称, 150, 150) <> 1 Then
            MsgBox "错误信息" & GetMyLastError(), vbInformation, "合医返回信息"
        Else
            Me.txt病种代码 = Trim(MidUni(R病种代码, 1, 30))
            Me.txt病种名称 = Trim(MidUni(R病种名称, 1, 200))
            Me.Cmd_ok.SetFocus
        End If
    End If
End Sub

Private Sub Txt合医编码_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        var机构代码 = Space(20)
        var机构名称 = Space(100)
        R合医编码 = Space(20)
        Dim rs医院编码 As New ADODB.Recordset
       
        gstrSQL = "select  医院编码 from 保险类别 where 序号=[1]"
        Set rs医院编码 = zlDatabase.OpenSQLRecord(gstrSQL, "取医院编码", gintInsure)
        var机构代码 = rs医院编码!医院编码
        rs医院编码.Close
        Set rs医院编码 = Nothing
        
        If Trim(Me.Txt合医编码) = "" Then
            If GetCHRYBM(R合医编码, 150, 150) <> 1 Then
                MsgBox "错误信息:" & GetMyLastError(), vbInformation, "合医返回信息"
                Exit Sub
            End If
        Call OutInfo
        
         Else
            R合医编码 = Trim(Me.Txt合医编码.Text)
            Call OutInfo
            Txt诊断结果.SetFocus
        End If
    End If
End Sub
Private Sub OutInfo()
R姓名 = Space(10)
R性别 = Space(4)
R身份证号 = Space(20)
R家庭帐号 = Space(20)
R出生日期 = Space(12)
R户主姓名 = Space(10)
R户主性别 = Space(4)
R户主身份证 = Space(20)
R住址 = Space(50)
R户主关系 = Space(10)

    If GetRyInfo(R合医编码, R姓名, R性别, R身份证号, R家庭帐号, R出生日期, R户主姓名, R户主性别, R户主身份证, R住址, R户主关系, R帐户余额) <> 1 Then
            MsgBox "错误信息" & GetMyLastError, vbInformation, "合医返回信息"
            Me.Cmd_ok.Enabled = False
            Me.Txt合医编码.SetFocus
            Exit Sub
    Else
        '输出信息到界面
        With Me
            .Txt合医编码.Text = Trim(MidUni(R合医编码, 1, 20))
            .Txt机构代码 = Trim(MidUni(var机构代码, 1, 10))
            .txt姓名.Text = Trim(MidUni(R姓名, 1, 10))
            .txt性别.Text = IIf(Trim(MidUni(R性别, 1, 4)) <> "男" And Trim(MidUni(R性别, 1, 4)) <> "女", "未知", Trim(MidUni(R性别, 1, 4)))
            .txt身份证号.Text = Trim(MidUni(R身份证号, 1, 20))
            .Txt家庭帐号.Text = Trim(MidUni(R家庭帐号, 1, 20))
            .txt出生日期.Text = Trim(MidUni(R出生日期, 1, 12))
            .Txt户主姓名.Text = Trim(MidUni(R户主姓名, 1, 10))
            .Txt户主性别.Text = Trim(MidUni(R户主性别, 1, 4))
            .Txt户主身份证.Text = Trim(MidUni(R户主身份证, 1, 20))
            .Txt住址.Text = Trim(MidUni(R住址, 1, 50))
            .Txt户主关系.Text = Trim(MidUni(R户主关系, 1, 10))
            .txt帐户余额.Text = Trim(MidUni(CStr(R帐户余额), 1, 10))
            .Txt诊断结果.Locked = False
            .Txt诊断结果.BackColor = &HFFFFFF
            .txt帐户余额.ForeColor = &HFF0000
            .Chk慢性病.Enabled = True
            .Chk组织体验.Enabled = True
            .Chk预防接种.Enabled = True
        End With
        If b就医类别 = 1 Then
            Me.txt病种代码.SetFocus
        Else
            Me.Txt诊断结果.SetFocus
        End If
        Me.Cmd_ok.Enabled = True
    End If
End Sub


Private Sub Txt诊断结果_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
           Cmd_ok.SetFocus
    End If
End Sub

