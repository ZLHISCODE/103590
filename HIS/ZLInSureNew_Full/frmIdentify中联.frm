VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmIdentify中联 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "医保病人身份识别"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7125
   Icon            =   "frmIdentify中联.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3930
      TabIndex        =   39
      Top             =   5070
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5340
      TabIndex        =   40
      Top             =   5070
      Width           =   1100
   End
   Begin VB.Frame fra基本 
      Caption         =   "病人帐户情况"
      Height          =   1305
      Index           =   1
      Left            =   150
      TabIndex        =   30
      Top             =   3570
      Width           =   6795
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   10
         Left            =   5130
         MaxLength       =   14
         TabIndex        =   38
         Top             =   720
         Width           =   1395
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   9
         Left            =   1590
         MaxLength       =   14
         TabIndex        =   36
         Top             =   720
         Width           =   1395
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   8
         Left            =   5130
         MaxLength       =   14
         TabIndex        =   34
         Top             =   330
         Width           =   1395
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   7
         Left            =   1590
         MaxLength       =   14
         TabIndex        =   32
         Top             =   330
         Width           =   1395
      End
      Begin VB.Label lbl提示 
         AutoSize        =   -1  'True
         Caption         =   "进入统筹累计(&G)"
         Height          =   180
         Index           =   9
         Left            =   180
         TabIndex        =   35
         Top             =   780
         Width           =   1350
      End
      Begin VB.Label lbl提示 
         AutoSize        =   -1  'True
         Caption         =   "帐户支出累计(&W)"
         Height          =   180
         Index           =   8
         Left            =   3690
         TabIndex        =   33
         Top             =   390
         Width           =   1350
      End
      Begin VB.Label lbl提示 
         AutoSize        =   -1  'True
         Caption         =   "帐户增加累计(&A)"
         Height          =   180
         Index           =   7
         Left            =   180
         TabIndex        =   31
         Top             =   390
         Width           =   1350
      End
      Begin VB.Label lbl提示 
         AutoSize        =   -1  'True
         Caption         =   "统筹报销累计(&P)"
         Height          =   180
         Index           =   10
         Left            =   3690
         TabIndex        =   37
         Top             =   780
         Width           =   1350
      End
   End
   Begin VB.Frame fra基本 
      Caption         =   "病人基本信息"
      Height          =   3195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   6795
      Begin VB.ComboBox Cbo当前状态 
         Height          =   300
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   12
         Left            =   4440
         MaxLength       =   18
         TabIndex        =   17
         Top             =   1515
         Width           =   2085
      End
      Begin VB.ComboBox cmb性别 
         Height          =   300
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1125
         Width           =   2085
      End
      Begin VB.ComboBox cmb中心 
         Height          =   300
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   330
         Width           =   2085
      End
      Begin VB.TextBox txtEdit 
         Enabled         =   0   'False
         Height          =   300
         Index           =   2
         Left            =   4440
         MaxLength       =   26
         TabIndex        =   26
         Top             =   2310
         Width           =   2085
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "…"
         Height          =   240
         Index           =   2
         Left            =   6240
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1935
         Width           =   255
      End
      Begin MSComCtl2.DTPicker dtp生日 
         Height          =   300
         Left            =   1320
         TabIndex        =   15
         Top             =   1515
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   87031811
         CurrentDate     =   36526
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   11
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   11
         Top             =   1125
         Width           =   1455
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "…"
         Height          =   240
         Index           =   1
         Left            =   6240
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   2730
         Width           =   255
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "…"
         Height          =   240
         Index           =   0
         Left            =   2490
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   6
         Left            =   4440
         MaxLength       =   2
         TabIndex        =   9
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   3
         Left            =   1320
         MaxLength       =   8
         TabIndex        =   19
         Top             =   1905
         Width           =   1455
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   1
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   7
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   2
         Top             =   330
         Width           =   1455
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   4
         Left            =   4440
         MaxLength       =   8
         TabIndex        =   21
         Top             =   1905
         Width           =   2085
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   5
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   2700
         Width           =   2085
      End
      Begin VB.Label lbl提示 
         AutoSize        =   -1  'True
         Caption         =   "当前状态(&K)"
         Height          =   180
         Index           =   16
         Left            =   240
         TabIndex        =   23
         Top             =   2340
         Width           =   990
      End
      Begin VB.Label lbl提示 
         AutoSize        =   -1  'True
         Caption         =   "出生日期(&B)"
         Height          =   180
         Index           =   15
         Left            =   240
         TabIndex        =   14
         Top             =   1575
         Width           =   990
      End
      Begin VB.Label lbl提示 
         AutoSize        =   -1  'True
         Caption         =   "性别(&X)"
         Height          =   180
         Index           =   14
         Left            =   3720
         TabIndex        =   12
         Top             =   1185
         Width           =   630
      End
      Begin VB.Label lbl提示 
         AutoSize        =   -1  'True
         Caption         =   "医保中心(&R)"
         Height          =   180
         Index           =   13
         Left            =   3360
         TabIndex        =   4
         Top             =   390
         Width           =   990
      End
      Begin VB.Label lbl提示 
         AutoSize        =   -1  'True
         Caption         =   "身份证号(&I)"
         Height          =   180
         Index           =   12
         Left            =   3360
         TabIndex        =   16
         Top             =   1575
         Width           =   990
      End
      Begin VB.Label lbl提示 
         AutoSize        =   -1  'True
         Caption         =   "姓名(&N)"
         Height          =   180
         Index           =   11
         Left            =   600
         TabIndex        =   10
         Top             =   1185
         Width           =   630
      End
      Begin VB.Label lbl提示 
         AutoSize        =   -1  'True
         Caption         =   "住院次数(&S)"
         Height          =   180
         Index           =   6
         Left            =   3360
         TabIndex        =   8
         Top             =   780
         Width           =   990
      End
      Begin VB.Label lbl提示 
         AutoSize        =   -1  'True
         Caption         =   "病种(&F)"
         Height          =   180
         Index           =   5
         Left            =   3720
         TabIndex        =   27
         Top             =   2760
         Width           =   630
      End
      Begin VB.Label lbl提示 
         AutoSize        =   -1  'True
         Caption         =   "单位编码(&U)"
         Height          =   180
         Index           =   4
         Left            =   3360
         TabIndex        =   20
         Top             =   1965
         Width           =   990
      End
      Begin VB.Label lbl提示 
         AutoSize        =   -1  'True
         Caption         =   "人员身份(&E)"
         Height          =   180
         Index           =   3
         Left            =   240
         TabIndex        =   18
         Top             =   1965
         Width           =   990
      End
      Begin VB.Label lbl提示 
         AutoSize        =   -1  'True
         Caption         =   "退休证号(&Z)"
         Enabled         =   0   'False
         Height          =   180
         Index           =   2
         Left            =   3360
         TabIndex        =   25
         Top             =   2370
         Width           =   990
      End
      Begin VB.Label lbl提示 
         AutoSize        =   -1  'True
         Caption         =   "医保号(&Y)"
         Height          =   180
         Index           =   1
         Left            =   420
         TabIndex        =   6
         Top             =   780
         Width           =   810
      End
      Begin VB.Label lbl提示 
         AutoSize        =   -1  'True
         Caption         =   "卡号(&D)"
         Height          =   180
         Index           =   0
         Left            =   600
         TabIndex        =   1
         Top             =   390
         Width           =   630
      End
   End
End
Attribute VB_Name = "frmIdentify中联"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum 文本Enum
    Text卡号 = 0
    Text医保号 = 1
    Text退休证号 = 2
    Text人员身份 = 3
    Text病人单位 = 4
    Text病种 = 5
    Text住院次数 = 6
    Text帐户增加累计 = 7
    Text帐户支出累计 = 8
    Text进入统筹累计 = 9
    Text统筹报销累计 = 10
    Text姓名 = 11
    Text身份证号 = 12
End Enum

Private Enum 选择Enum
    Select卡号 = 0
    Select病种 = 1
    Select单位 = 2
End Enum

Dim mstrIdentify As String
Dim mbytType As Byte
Dim mlng病人ID As Long
Dim mint险类 As Integer

Public Function ShowCard(Optional bytType As Byte, Optional lng病人ID As Long, Optional ByVal int险类 As Integer) As String
'功能：返回医保病人的身份信息
'参数：bytType-识别类型，0-门诊，1-住院
'返回：
    '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);
    '8.中心代码;9.顺序号;10人员身份;11帐户余额;12当前状态;13病种ID;14在职(0,1);15退休证号;16年龄段;17灰度级
    '18帐户增加累计,19帐户支出累计,20进入统筹累计,21统筹报销累计,22住院次数累计,23就诊类型
    Dim rsTemp As New ADODB.Recordset
    mbytType = bytType
    mlng病人ID = lng病人ID
    mint险类 = int险类
    mstrIdentify = ""
    
    cmb性别.Clear
    gstrSQL = "select 编码,名称 from 性别 order by 编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    Do Until rsTemp.EOF
        cmb性别.AddItem rsTemp("编码") & "." & rsTemp("名称")
        rsTemp.MoveNext
    Loop
    
    cmb中心.Clear
    gstrSQL = "select A.具有中心,B.序号,B.编码,B.名称 from 保险类别 A,保险中心目录 B where A.序号=[1] and A.序号=b.险类"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mint险类)
    
    If rsTemp("具有中心") = 0 Then
        lbl提示(13).Visible = False
        cmb中心.Visible = False
        cmb中心.AddItem "1.中心" '单中心
    End If
    Do Until rsTemp.EOF
        cmb中心.AddItem rsTemp("编码") & "." & rsTemp("名称")
        cmb中心.ItemData(cmb中心.NewIndex) = rsTemp("序号")
        rsTemp.MoveNext
    Loop
    cmb中心.ListIndex = 0
    
    '1-在职;2-退休;3-离休
    Cbo当前状态.Clear
    Cbo当前状态.AddItem "在职"
    Cbo当前状态.ItemData(Cbo当前状态.NewIndex) = 1
    Cbo当前状态.AddItem "退休"
    Cbo当前状态.ItemData(Cbo当前状态.NewIndex) = 2
    Cbo当前状态.AddItem "离休"
    Cbo当前状态.ItemData(Cbo当前状态.NewIndex) = 3
    Cbo当前状态.ListIndex = 0
        
    dtp生日.MaxDate = zlDatabase.Currentdate
    frmIdentify中联.Show vbModal
    ShowCard = mstrIdentify
End Function

Private Sub Cbo当前状态_Click()
    TxtEdit(Text退休证号).Enabled = (Cbo当前状态.ListIndex <> 0)
    lbl提示(Text退休证号).Enabled = (Cbo当前状态.ListIndex <> 0)
End Sub

Private Sub Cbo当前状态_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub cmb中心_Click()
    Dim lng卡号长度 As Long, lng退休证长度 As Long
    Dim rsTemp As New ADODB.Recordset
    
    '缺省值
    lng卡号长度 = 20
    lng退休证长度 = 26
    
    gstrSQL = "select 参数名,参数值 from 保险参数 where 险类=[1] and (中心 is null or 中心=[2])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mint险类, CInt(cmb中心.ItemData(cmb中心.ListIndex)))
    
    Do Until rsTemp.EOF
        Select Case rsTemp("参数名")
            Case "卡号长度"
                If IsNull(rsTemp("参数值")) = False Then lng卡号长度 = Val(rsTemp("参数值"))
            Case "退休证长度"
                If IsNull(rsTemp("参数值")) = False Then lng退休证长度 = Val(rsTemp("参数值"))
        End Select
        rsTemp.MoveNext
    Loop
    
    TxtEdit(Text卡号).MaxLength = lng卡号长度
    TxtEdit(Text退休证号).MaxLength = lng退休证长度
End Sub

Private Sub cmdCancel_Click()
    mstrIdentify = ""
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim rsTemp As New ADODB.Recordset
    Dim strIdentify As String, strAddition As String
    Dim lng病人ID As Long, lng中心 As Long
    
    '首先验数据的正确性
    If IsValid() = False Then
        Exit Sub
    End If
    
    '得到中心序号
    If cmb中心.Visible = False Then
        lng中心 = 0
    Else
        If cmb中心.ListIndex < 0 Then
            MsgBox "请选择病人所属医保中心。", vbInformation, gstrSysName
            cmb中心.SetFocus
            Exit Sub
        End If
        lng中心 = cmb中心.ItemData(cmb中心.ListIndex)
    End If
    
    '检查病人状态
    gstrSQL = "select nvl(当前状态,0) as 状态 from 保险帐户 where 险类=[1] and 中心=[2] and 医保号=[3]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mint险类, lng中心, CStr(Trim(TxtEdit(Text医保号).Text)))
    
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
    '18帐户增加累计,19帐户支出累计,20进入统筹累计,21统筹报销累计,22住院次数累计,23就诊类型
    strIdentify = Trim(TxtEdit(Text卡号).Text)                         '0卡号
    strIdentify = strIdentify & ";" & Trim(TxtEdit(Text医保号).Text)   '1医保号
    strIdentify = strIdentify & ";"                                    '2密码
    strIdentify = strIdentify & ";" & Trim(TxtEdit(Text姓名).Text)     '3姓名
    strIdentify = strIdentify & ";" & Replace(GetTextFromCombo(cmb性别, True), "'", "") '4性别
    strIdentify = strIdentify & ";" & Format(dtp生日.Value, "yyyy-MM-dd") '5出生日期
    strIdentify = strIdentify & ";" & Trim(TxtEdit(Text身份证号).Text)    '6身份证
    strIdentify = strIdentify & ";" & Trim(TxtEdit(Text病人单位).Text) & "(" & Trim(TxtEdit(Text病人单位).Text) & ")"  '7.单位名称(编码)
    strAddition = ";" & lng中心                                 '8.中心代码
    strAddition = strAddition & ";"                             '9.顺序号
    strAddition = strAddition & ";" & Trim(TxtEdit(Text人员身份).Text)       '10人员身份
    strAddition = strAddition & ";" & Val(TxtEdit(Text帐户增加累计).Text) - Val(TxtEdit(Text帐户支出累计).Text)  '11帐户余额
    strAddition = strAddition & ";0"                            '12当前状态
    strAddition = strAddition & ";" & TxtEdit(Text病种).Tag     '13病种ID
    strAddition = strAddition & ";" & Cbo当前状态.ItemData(Cbo当前状态.ListIndex) '14在职(1,2,3)
    strAddition = strAddition & ";" & Trim(TxtEdit(Text退休证号).Text) '15退休证号
    strAddition = strAddition & ";" & DateDiff("yyyy", dtp生日.Value, dtp生日.MaxDate) '16年龄段
    strAddition = strAddition & ";"                             '17灰度级
    strAddition = strAddition & ";" & Val(TxtEdit(Text帐户增加累计).Text)       '18帐户增加累计
    strAddition = strAddition & ";" & Val(TxtEdit(Text帐户支出累计).Text)       '19帐户支出累计
    strAddition = strAddition & ";" & Val(TxtEdit(Text进入统筹累计).Text)       '20进入统筹累计
    strAddition = strAddition & ";" & Val(TxtEdit(Text统筹报销累计).Text)       '21统筹报销累计
    strAddition = strAddition & ";" & Int(Val(TxtEdit(Text住院次数).Text))      '22住院次数累计
    strAddition = strAddition & ";"                                             '23就诊类型 (1、急诊门诊)
    
    lng病人ID = BuildPatiInfo(mbytType, strIdentify & strAddition, mlng病人ID, mint险类)
    '返回格式:中间插入病人ID
    If lng病人ID > 0 Then
        mstrIdentify = strIdentify & ";" & lng病人ID & strAddition
    End If
    Unload Me
End Sub

Private Function IsValid() As Boolean
'功能：检查数据的正确性
    Dim lngIndex As Long
    
    For lngIndex = TxtEdit.LBound To TxtEdit.UBound
        If TxtEdit(lngIndex).Enabled = True Then
            If zlCommFun.StrIsValid(TxtEdit(lngIndex), TxtEdit(lngIndex).MaxLength) = False Then
                zlControl.TxtSelAll TxtEdit(lngIndex)
                TxtEdit(lngIndex).SetFocus
                Exit Function
            End If
        End If
        
        If lngIndex >= Text住院次数 And lngIndex <= Text统筹报销累计 Then
            If IsNumeric(TxtEdit(lngIndex).Text) = False Then
                MsgBox "请输入合法的数值。", vbInformation, gstrSysName
                zlControl.TxtSelAll TxtEdit(lngIndex)
                TxtEdit(lngIndex).SetFocus
                Exit Function
            End If
            
            
            If lngIndex = Text住院次数 Then
                If Val(TxtEdit(lngIndex).Text) < 0 Or Val(TxtEdit(lngIndex).Text) > 100 Then
                    MsgBox "住院次数不能小于0，且不能超过100。", vbInformation, gstrSysName
                    zlControl.TxtSelAll TxtEdit(Text住院次数)
                    TxtEdit(Text住院次数).SetFocus
                    Exit Function
                End If
            Else
                If Val(TxtEdit(lngIndex).Text) < 0 Or Val(TxtEdit(lngIndex).Text) > 1000000 Then
                    MsgBox "金额不能小于0，且不能超过100万。", vbInformation, gstrSysName
                    zlControl.TxtSelAll TxtEdit(lngIndex)
                    TxtEdit(lngIndex).SetFocus
                    Exit Function
                End If
            End If
        End If
        If (lngIndex = Text卡号 Or lngIndex = Text医保号 Or lngIndex = Text姓名) And Trim(TxtEdit(lngIndex).Text) = "" Then
            MsgBox "卡号、医保号、姓名都不能为空。", vbInformation, gstrSysName
            zlControl.TxtSelAll TxtEdit(lngIndex)
            TxtEdit(lngIndex).SetFocus
            Exit Function
        End If
    Next
    
    
    If Val(TxtEdit(Text帐户增加累计).Text) < Val(TxtEdit(Text帐户支出累计).Text) Then
        MsgBox "帐户累计支出不能超过帐户累计增加。", vbInformation, gstrSysName
        zlControl.TxtSelAll TxtEdit(Text帐户支出累计)
        TxtEdit(Text帐户支出累计).SetFocus
        Exit Function
    End If
    
    If Val(TxtEdit(Text进入统筹累计).Text) < Val(TxtEdit(Text统筹报销累计).Text) Then
        MsgBox "统筹报销累计不能超过进入统筹累计。", vbInformation, gstrSysName
        zlControl.TxtSelAll TxtEdit(Text统筹报销累计)
        TxtEdit(Text统筹报销累计).SetFocus
        Exit Function
    End If
    
    IsValid = True
End Function

Private Sub cmdSelect_Click(Index As Integer)
    Dim rsTemp As ADODB.Recordset
    
    Select Case Index
        Case Select卡号
            gstrSQL = " Select A.病人ID as ID,A.卡号,A.医保号,B.姓名,B.性别,B.出生日期,B.身份证号,C.序号 as 中心ID " & _
                    " ,A.人员身份,A.单位编码,A.病种ID,D.名称 as 病种,A.在职 as 在职ID,A.退休证号" & _
                    " From 保险帐户 A,病人信息 B,保险中心目录 C,保险病种 D" & _
                    "  where A.病人ID=B.病人ID and A.险类=" & mint险类 & _
                    "  and A.险类=C.险类 and A.中心=C.序号 and A.病种ID=D.ID(+)"
            
            Call Get帐户情况
            zlControl.TxtSelAll TxtEdit(Text卡号)
            TxtEdit(Text卡号).SetFocus
        Case Select单位
            Set rsTemp = frmPubSel.ShowSelect(Me, _
                    " Select ID,上级ID,末级,编码,名称,地址,电话,开户银行,帐号,联系人 From 合约单位" & _
                    " Start With 上级ID is NULL Connect by Prior ID=上级ID", _
                    2, "工作单位", , TxtEdit(Text病人单位).Text)
            If Not rsTemp Is Nothing Then
                TxtEdit(Text病人单位).Text = rsTemp("编码")
                zlControl.TxtSelAll TxtEdit(Text病人单位)
            End If
            TxtEdit(Text病人单位).SetFocus
        
        Case Select病种
            gstrSQL = " Select A.ID,A.编码,A.名称,A.简码,decode(A.类别,1,'慢性病',2,'特种病','普通病') as 类别 " & _
                    " From 保险病种 A where A.险类=" & mint险类
            
            Set rsTemp = frmPubSel.ShowSelect(Me, gstrSQL, 0, "医保病种", , TxtEdit(Text病种).Text)
            If Not rsTemp Is Nothing Then
                TxtEdit(Text病种).Text = rsTemp("名称")
                TxtEdit(Text病种).Tag = rsTemp("ID")
                zlControl.TxtSelAll TxtEdit(Text病种)
            End If
            TxtEdit(Text病种).SetFocus
    End Select
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll TxtEdit(Index)
    Select Case Index
        Case Text姓名
            zlCommFun.OpenIme True
        Case Else
            zlCommFun.OpenIme False
    End Select
End Sub

Private Sub dtp生日_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = Text病种 Then
        If KeyCode = vbKeyDelete Then
            TxtEdit(Text病种).Text = ""
            TxtEdit(Text病种).Tag = ""
        End If
    End If
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strCode As String
    Dim str条件 As String
    Dim rsTemp As New ADODB.Recordset
    
    If Index = Text卡号 Then
        If Len(TxtEdit(Text卡号).Text) = TxtEdit(Text卡号).MaxLength Or KeyAscii = vbKeyReturn Then
            strCode = Replace(Trim(TxtEdit(Text卡号).Text), "'", "")
            
            If IsNumeric(strCode) And IsNumeric(Left(strCode, 1)) Then '刷卡
                str条件 = " and A.卡号='" & strCode & "' and A.中心=" & cmb中心.ItemData(cmb中心.ListIndex)
            ElseIf (Left(strCode, 1) = "A" Or Left(strCode, 1) = "-") And IsNumeric(Mid(strCode, 2)) Then '病人ID
                str条件 = " and A.病人ID=" & Mid(strCode, 2)
            ElseIf (Left(strCode, 1) = "B" Or Left(strCode, 1) = "+") And IsNumeric(Mid(strCode, 2)) Then '住院号(对住(过)院的病人)
                str条件 = " and B.住院号=" & Mid(strCode, 2)
            ElseIf (Left(strCode, 1) = "D" Or Left(strCode, 1) = "*") And IsNumeric(Mid(strCode, 2)) Then '门诊号(仅对门诊病人)
                str条件 = " and B.门诊号=" & Mid(strCode, 2)
            Else '当作姓名
                str条件 = " and B.姓名='" & strCode & "'"
            End If
        
            gstrSQL = " Select A.病人ID as ID,A.卡号,A.医保号,B.姓名,B.性别,B.出生日期,B.身份证号,C.序号 as 中心ID " & _
                    " ,A.人员身份,A.单位编码,A.病种ID,D.名称 as 病种,A.在职 as 在职ID,A.退休证号" & _
                    " From 保险帐户 A,病人信息 B,保险中心目录 C,保险病种 D" & _
                    "  where A.病人ID=B.病人ID and A.险类=" & mint险类 & _
                    "  and A.险类=C.险类 and A.中心=C.序号 and A.病种ID=D.ID(+)" & str条件
            
            Call Get帐户情况
        End If
    End If
    
    If KeyAscii = vbKeyReturn Then
        
        KeyAscii = 0  '消除响声
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub cmb中心_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub cmb性别_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    If Index >= Text住院次数 And Index <= Text统筹报销累计 Then
        If Index = Text住院次数 Then
            TxtEdit(Index).Text = Format(Val(TxtEdit(Index).Text), "#0;0;0;0")
        Else
            TxtEdit(Index).Text = Format(Val(TxtEdit(Index).Text), "######0.00;0.00;0.00;0.00")
        End If
    End If

End Sub

Private Sub Get帐户情况()
'从已经存在的记录中读出帐户信息
    Dim rs帐户 As ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim lngIndex As Long
    
    
    Set rs帐户 = frmPubSel.ShowSelect(Me, gstrSQL, 0, "保险帐户", , TxtEdit(Text卡号).Text, "", False, True)
    If Not rs帐户 Is Nothing Then
    
        TxtEdit(Text卡号).Text = rs帐户("卡号")
        '其它可用的数据
        TxtEdit(Text医保号).Text = IIf(IsNull(rs帐户("医保号")), "", rs帐户("医保号"))
        TxtEdit(Text姓名).Text = IIf(IsNull(rs帐户("姓名")), "", rs帐户("姓名"))
        TxtEdit(Text身份证号).Text = IIf(IsNull(rs帐户("身份证号")), "", rs帐户("身份证号"))
        TxtEdit(Text人员身份).Text = IIf(IsNull(rs帐户("人员身份")), "", rs帐户("人员身份"))
        TxtEdit(Text病人单位).Text = IIf(IsNull(rs帐户("单位编码")), "", rs帐户("单位编码"))
        TxtEdit(Text病种).Text = IIf(IsNull(rs帐户("病种")), "", rs帐户("病种"))
        TxtEdit(Text病种).Tag = IIf(IsNull(rs帐户("病种ID")), "", rs帐户("病种ID"))
        
        Call SetComboByText(cmb性别, IIf(IsNull(rs帐户("性别")), "", rs帐户("性别")), True)
        Cbo当前状态.ListIndex = rs帐户("在职ID") - 1
        TxtEdit(Text退休证号).Text = ""
        If Cbo当前状态.ListIndex <> 0 Then
            TxtEdit(Text退休证号).Text = IIf(IsNull(rs帐户("退休证号")), "", rs帐户("退休证号"))
        End If
        If IsNull(rs帐户("出生日期")) = False Then
            dtp生日.Value = rs帐户("出生日期")
        End If
        
        For lngIndex = 0 To cmb中心.ListCount - 1
            If cmb中心.ItemData(lngIndex) = rs帐户("中心ID") Then
                cmb中心.ListIndex = lngIndex
                Exit For
            End If
        Next
        
        '再读出帐户年度信息
        gstrSQL = "select * from 帐户年度信息 where 险类=[1]" & _
            " and 病人ID=[2] and 年度=[3]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mint险类, CLng(rs帐户("ID")), Year(dtp生日.MaxDate))
        
        If rsTemp.EOF = False Then
            '设置帐户情况
            TxtEdit(Text住院次数).Text = Format(rsTemp("住院次数累计"), "#0;0;0;0")
            TxtEdit(Text帐户增加累计).Text = Format(rsTemp("帐户增加累计"), "######0.00;0.00;0.00;0.00")
            TxtEdit(Text帐户支出累计).Text = Format(rsTemp("帐户支出累计"), "######0.00;0.00;0.00;0.00")
            TxtEdit(Text进入统筹累计).Text = Format(rsTemp("进入统筹累计"), "######0.00;0.00;0.00;0.00")
            TxtEdit(Text统筹报销累计).Text = Format(rsTemp("统筹报销累计"), "######0.00;0.00;0.00;0.00")
        Else
            TxtEdit(Text住院次数).Text = "0"
            TxtEdit(Text帐户增加累计).Text = "0.00"
            TxtEdit(Text帐户支出累计).Text = "0.00"
            TxtEdit(Text进入统筹累计).Text = "0.00"
            TxtEdit(Text统筹报销累计).Text = "0.00"
        End If
        
    End If
End Sub

