VERSION 5.00
Begin VB.Form frmIdentify福建巨龙 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "身份验证"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8010
   Icon            =   "frmIdentify福建巨龙.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   8010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fra基本 
      Caption         =   "病人基本信息"
      Height          =   2925
      Index           =   0
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      Begin VB.TextBox txt姓名 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   2
         Top             =   285
         Width           =   2265
      End
      Begin VB.TextBox TxtIC卡状态 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1200
         MaxLength       =   19
         TabIndex        =   12
         Top             =   1980
         Width           =   2265
      End
      Begin VB.TextBox Txt所属地区 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4830
         MaxLength       =   19
         TabIndex        =   18
         Top             =   720
         Width           =   2625
      End
      Begin VB.TextBox Txt所属分中心 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4830
         MaxLength       =   19
         TabIndex        =   20
         Top             =   1140
         Width           =   2625
      End
      Begin VB.TextBox Txt单位名称 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4830
         MaxLength       =   19
         TabIndex        =   16
         Top             =   285
         Width           =   2625
      End
      Begin VB.TextBox Txt年度累计 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4830
         MaxLength       =   19
         TabIndex        =   26
         Top             =   1980
         Width           =   2625
      End
      Begin VB.TextBox Txt帐户余额 
         Enabled         =   0   'False
         Height          =   300
         Left            =   6030
         MaxLength       =   19
         TabIndex        =   24
         Top             =   1560
         Width           =   1425
      End
      Begin VB.TextBox Txt住院次数 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4830
         MaxLength       =   19
         TabIndex        =   22
         Top             =   1560
         Width           =   405
      End
      Begin VB.TextBox Txt工作状态 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1200
         MaxLength       =   19
         TabIndex        =   14
         Top             =   2400
         Width           =   2265
      End
      Begin VB.TextBox txt年龄 
         Enabled         =   0   'False
         Height          =   300
         Left            =   3030
         MaxLength       =   20
         TabIndex        =   6
         Top             =   720
         Width           =   435
      End
      Begin VB.TextBox txt医保号 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1200
         MaxLength       =   19
         TabIndex        =   10
         Top             =   1560
         Width           =   2265
      End
      Begin VB.CommandButton cmd病种 
         Caption         =   "…"
         Height          =   240
         Left            =   7170
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   2430
         Width           =   255
      End
      Begin VB.ComboBox cob性别 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txt病种 
         Height          =   300
         Left            =   4830
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   2400
         Width           =   2625
      End
      Begin VB.TextBox txt卡号 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1200
         MaxLength       =   12
         TabIndex        =   8
         Top             =   1140
         Width           =   2265
      End
      Begin VB.Label lbl姓名 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "姓名(&N)"
         Enabled         =   0   'False
         Height          =   180
         Left            =   480
         TabIndex        =   1
         Top             =   345
         Width           =   630
      End
      Begin VB.Label LblIC卡状态 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "IC卡状态(&S)"
         Enabled         =   0   'False
         Height          =   180
         Left            =   120
         TabIndex        =   11
         Top             =   2040
         Width           =   990
      End
      Begin VB.Label Lbl所属地区 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "所属地区(&A)"
         Enabled         =   0   'False
         Height          =   180
         Left            =   3750
         TabIndex        =   17
         Top             =   780
         Width           =   990
      End
      Begin VB.Label Lbl所属分中心 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "所属分中心(&F)"
         Enabled         =   0   'False
         Height          =   180
         Left            =   3570
         TabIndex        =   19
         Top             =   1200
         Width           =   1170
      End
      Begin VB.Label Lbl单位名称 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "单位名称(&W)"
         Enabled         =   0   'False
         Height          =   180
         Left            =   3750
         TabIndex        =   15
         Top             =   345
         Width           =   990
      End
      Begin VB.Label Lbl年度累计 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "年度累计(&L)"
         Enabled         =   0   'False
         Height          =   180
         Left            =   3750
         TabIndex        =   25
         Top             =   2040
         Width           =   990
      End
      Begin VB.Label Lbl帐户余额 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "余额(&Q)"
         Enabled         =   0   'False
         Height          =   180
         Left            =   5310
         TabIndex        =   23
         Top             =   1620
         Width           =   630
      End
      Begin VB.Label Lbl住院次数 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "住院次数(&Z)"
         Enabled         =   0   'False
         Height          =   180
         Left            =   3750
         TabIndex        =   21
         Top             =   1620
         Width           =   990
      End
      Begin VB.Label lbl工作状态 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "工作状态(&T)"
         Enabled         =   0   'False
         Height          =   180
         Left            =   120
         TabIndex        =   13
         Top             =   2460
         Width           =   990
      End
      Begin VB.Label lbl病种 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "病种(&F)"
         Height          =   180
         Left            =   4080
         TabIndex        =   27
         Top             =   2460
         Width           =   630
      End
      Begin VB.Label lbl年龄 
         AutoSize        =   -1  'True
         Caption         =   "年龄(&A)"
         Enabled         =   0   'False
         Height          =   180
         Left            =   2310
         TabIndex        =   5
         Top             =   780
         Width           =   630
      End
      Begin VB.Label lbl卡号 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "卡号(&D)"
         Enabled         =   0   'False
         Height          =   180
         Left            =   480
         TabIndex        =   7
         Top             =   1200
         Width           =   630
      End
      Begin VB.Label lbl医保号 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "医保号(&Y)"
         Enabled         =   0   'False
         Height          =   180
         Left            =   300
         TabIndex        =   9
         Top             =   1620
         Width           =   810
      End
      Begin VB.Label lbl性别 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "性别(&X)"
         Enabled         =   0   'False
         Height          =   180
         Left            =   480
         TabIndex        =   3
         Top             =   780
         Width           =   630
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6600
      TabIndex        =   31
      Top             =   3150
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5340
      TabIndex        =   30
      Top             =   3150
      Width           =   1100
   End
End
Attribute VB_Name = "frmIdentify福建巨龙"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private blnOK As Boolean
Dim mlng病人ID As Long
Dim mint险类 As Integer

Public Function ShowCard(Optional lng病人ID As Long, Optional ByVal int险类 As Integer) As Boolean
'功能：返回医保病人的身份信息
'参数：bytType-识别类型，0-门诊，1-住院
'返回：
    '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);
    '8.中心代码;9.顺序号;10人员身份;11帐户余额;12当前状态;13病种ID;14在职(0,1);15退休证号;16年龄段;17灰度级
    '18帐户增加累计,19帐户支出累计,20进入统筹累计,21统筹报销累计,22住院次数累计
    Dim rsTemp As New ADODB.Recordset
    mlng病人ID = lng病人ID
    mint险类 = int险类
    
    cob性别.Clear
    gstrSQL = "select 编码,名称 from 性别 order by 编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    Do Until rsTemp.EOF
        cob性别.AddItem rsTemp("编码") & "." & rsTemp("名称")
        rsTemp.MoveNext
    Loop
    cob性别.ListIndex = 0
    rsTemp.Close
    Call Get帐户情况

    frmIdentify福建巨龙.Show vbModal
    ShowCard = blnOK
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    '更新病种
    
    gstrSQL = "zl_保险帐户_更新信息(" & mlng病人ID & "," & mint险类 & ",'病种ID','" & IIf(txt病种.Tag = "", "NULL", txt病种.Tag) & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "更新病种")
    blnOK = True
    Unload Me
End Sub

Private Sub Get帐户情况()
'从已经存在的记录中读出帐户信息
    Dim strValue As String
    Dim rs帐户 As ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim lngIndex As Long
    
    gstrSQL = " Select A.病人ID as ID,A.卡号,A.医保号,B.姓名,B.性别,B.年龄, " & _
              " A.病种ID,substr(D.名称,instr(D.名称,'@@')+2) as 病种,单位编码 As 工作状态,Nvl(人员身份,0) As 住院次数,退休证号 As 年度医保费用累计,帐户余额" & _
              " From 保险帐户 A,病人信息 B,保险中心目录 C,保险病种 D" & _
              " Where A.病人ID=B.病人ID and A.险类=" & mint险类 & _
              " And A.险类=C.险类 and A.中心=C.序号 and A.病种ID=D.ID(+) And A.病人ID=" & mlng病人ID
    Set rs帐户 = frmPubSel.ShowSelect(Me, gstrSQL, 0, "保险帐户", , "", "", False, True)
    If Not rs帐户 Is Nothing Then
    
        '其它可用的数据
        mlng病人ID = rs帐户!ID
        txt姓名.Text = IIf(IsNull(rs帐户("姓名")), "", rs帐户("姓名"))
        txt年龄.Text = IIf(IsNull(rs帐户("年龄")), "", rs帐户("年龄"))
        Call SetComboByText(cob性别, IIf(IsNull(rs帐户("性别")), "", rs帐户("性别")), True)
        txt卡号.Text = IIf(IsNull(rs帐户("卡号")), "", rs帐户("卡号"))
        txt医保号.Text = IIf(IsNull(rs帐户("医保号")), "", rs帐户("医保号"))
        txt住院次数.Text = IIf(IsNull(rs帐户("住院次数")), "", rs帐户("住院次数"))
        Txt工作状态.Text = IIf(IsNull(rs帐户("工作状态")), "", rs帐户("工作状态"))
        txt帐户余额.Text = Format(rs帐户("帐户余额"), "#####0.00;-#####0.00; ;")
        Txt年度累计.Text = Format(rs帐户("年度医保费用累计"), "#####0.00;-#####0.00; ;")
'        txt病种.Text = IIf(IsNull(rs帐户("病种")), "", rs帐户("病种"))
'        txt病种.Tag = IIf(IsNull(rs帐户("病种ID")), "", rs帐户("病种ID"))
    End If
    
    '填写附加信息
    Call Record_Locate(mrsIniItems, "名称,Dwmc00")
    txt单位名称.Text = Nvl(mrsIniItems!值, "")
    Call Record_Locate(mrsIniItems, "名称,Icztmc")
    TxtIC卡状态.Text = Nvl(mrsIniItems!值, "")
    Call Record_Locate(mrsIniItems, "名称,Dqmc00")
    Txt所属地区.Text = Nvl(mrsIniItems!值, "")
    Call Record_Locate(mrsIniItems, "名称,Fzxmc0")
    Txt所属分中心.Text = Nvl(mrsIniItems!值, "")
End Sub

Private Sub cmd病种_Click()
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = " Select A.ID,substr(A.名称,1,instr(A.名称,'@@')-1) 编码,substr(A.名称,instr(A.名称,'@@')+2) 名称,A.简码,decode(A.类别,1,'慢性病',2,'特种病','普通病') as 类别 " & _
            " From 保险病种 A where A.险类=" & mint险类
    
    Set rsTemp = frmPubSel.ShowSelect(Me, gstrSQL, 0, "医保病种", , txt病种.Text)
    If rsTemp.State = 0 Then Exit Sub
    If Not rsTemp Is Nothing Then
        txt病种.Text = rsTemp("名称")
        txt病种.Tag = rsTemp("ID")
        zlControl.TxtSelAll txt病种
    End If
    txt病种.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    blnOK = False
End Sub

Private Sub txt病种_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        txt病种.Text = ""
        txt病种.Tag = ""
    End If
End Sub

Private Function Record_Locate(ByRef rsObj As ADODB.Recordset, ByVal strPrimary As String, Optional ByVal blnDelete As Boolean = False) As Boolean
    Dim arrTmp
    '定位到指定记录
    'strPrimary:主健,值
    'blnDelete=True,则该记录集存在"删除"字段
    Record_Locate = False
    
    arrTmp = Split(strPrimary, ",")
    With rsObj
        If .RecordCount = 0 Then Exit Function
        .MoveFirst
        .Find arrTmp(0) & "='" & arrTmp(1) & "'"
        If .EOF Then Exit Function
        If blnDelete Then
            Do While Not .EOF
                If !删除 = 0 Then Record_Locate = True: Exit Do
                .MoveNext
            Loop
        Else
            Record_Locate = True
        End If
    End With
End Function
