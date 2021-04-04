VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmIdentify中梁山 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "医保病人身份识别"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5235
   Icon            =   "frmIdentify中梁山.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4020
      TabIndex        =   21
      Top             =   210
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4020
      TabIndex        =   22
      Top             =   690
      Width           =   1100
   End
   Begin VB.Frame fra基本 
      Caption         =   "病人基本信息"
      Height          =   4695
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   3705
      Begin VB.TextBox txtEdit 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   420
         Index           =   5
         Left            =   240
         MaxLength       =   14
         TabIndex        =   20
         Top             =   3990
         Width           =   3195
      End
      Begin VB.ComboBox cbo类别 
         Height          =   300
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   2670
         Width           =   2085
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   3
         Left            =   1320
         MaxLength       =   18
         TabIndex        =   7
         Top             =   1110
         Width           =   2085
      End
      Begin VB.ComboBox cbo性别 
         Height          =   300
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1500
         Width           =   2085
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   2
         Left            =   1320
         MaxLength       =   26
         TabIndex        =   13
         Top             =   2280
         Width           =   2085
      End
      Begin MSComCtl2.DTPicker dtp生日 
         Height          =   300
         Left            =   1320
         TabIndex        =   11
         Top             =   1890
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   23855107
         CurrentDate     =   36526
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   1
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   5
         Top             =   720
         Width           =   2085
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "…"
         Height          =   240
         Index           =   1
         Left            =   3120
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   3090
         Width           =   255
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "…"
         Height          =   240
         Index           =   0
         Left            =   3120
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   2
         Top             =   330
         Width           =   2085
      End
      Begin VB.TextBox txtEdit 
         ForeColor       =   &H00000080&
         Height          =   300
         Index           =   4
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   3060
         Width           =   2085
      End
      Begin VB.Label lbl提示 
         AutoSize        =   -1  'True
         Caption         =   "统筹报销累计(&P)"
         Height          =   180
         Index           =   5
         Left            =   240
         TabIndex        =   19
         Top             =   3720
         Width           =   1350
      End
      Begin VB.Label lbl提示 
         AutoSize        =   -1  'True
         Caption         =   "人员类别(&K)"
         Height          =   180
         Index           =   16
         Left            =   240
         TabIndex        =   14
         Top             =   2730
         Width           =   990
      End
      Begin VB.Label lbl提示 
         AutoSize        =   -1  'True
         Caption         =   "出生日期(&B)"
         Height          =   180
         Index           =   15
         Left            =   240
         TabIndex        =   10
         Top             =   1950
         Width           =   990
      End
      Begin VB.Label lbl提示 
         AutoSize        =   -1  'True
         Caption         =   "性别(&X)"
         Height          =   180
         Index           =   14
         Left            =   600
         TabIndex        =   8
         Top             =   1560
         Width           =   630
      End
      Begin VB.Label lbl提示 
         AutoSize        =   -1  'True
         Caption         =   "身份证号(&I)"
         Height          =   180
         Index           =   3
         Left            =   240
         TabIndex        =   6
         Top             =   1170
         Width           =   990
      End
      Begin VB.Label lbl提示 
         AutoSize        =   -1  'True
         Caption         =   "姓名(&N)"
         Height          =   180
         Index           =   1
         Left            =   600
         TabIndex        =   4
         Top             =   780
         Width           =   630
      End
      Begin VB.Label lbl提示 
         AutoSize        =   -1  'True
         Caption         =   "病种(&F)"
         Height          =   180
         Index           =   4
         Left            =   600
         TabIndex        =   16
         Top             =   3120
         Width           =   630
      End
      Begin VB.Label lbl提示 
         AutoSize        =   -1  'True
         Caption         =   "退休证号(&Z)"
         Height          =   180
         Index           =   2
         Left            =   240
         TabIndex        =   12
         Top             =   2340
         Width           =   990
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
Attribute VB_Name = "frmIdentify中梁山"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum 文本Enum
    Text卡号 = 0
    Text姓名 = 1
    Text退休证号 = 2
    Text身份证号 = 3
    Text病种 = 4
    Text统筹报销累计 = 5
End Enum

Private Enum 选择Enum
    Select卡号 = 0
    Select病种 = 1
End Enum

Dim mstrIdentify As String
Dim mbytType As Byte
Dim mlng病人ID As Long

Public Function ShowCard(Optional bytType As Byte, Optional lng病人ID As Long) As String
'功能：返回医保病人的身份信息
'参数：0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐
'返回：
    '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);
    '8.中心代码;9.顺序号;10人员身份;11帐户余额;12当前状态;13病种ID;14在职(0,1);15退休证号;16年龄段;17灰度级
    '18帐户增加累计,19帐户支出累计,20进入统筹累计,21统筹报销累计,22住院次数累计,23就诊类型
    Dim rsTemp As New ADODB.Recordset
    Dim lng卡号长度 As Long, lng退休证长度 As Long
    
    If bytType <> 1 Then
        MsgBox "本医保只支持入院登记。", vbInformation, gstrSysName
        Exit Function
    End If
    
    mbytType = bytType
    mlng病人ID = lng病人ID
    mstrIdentify = ""
    
    cbo性别.Clear
    gstrSQL = "select 编码,名称 from 性别 order by 编码"
    Call OpenRecordset(rsTemp, Me.Caption)
    
    Do Until rsTemp.EOF
        cbo性别.AddItem rsTemp("编码") & "." & rsTemp("名称")
        rsTemp.MoveNext
    Loop
    
    cbo类别.Clear
    gstrSQL = "select A.序号,A.名称 from 保险人群 A where A.险类=" & TYPE_重庆中梁山
    Call OpenRecordset(rsTemp, Me.Caption)
    Do Until rsTemp.EOF
        cbo类别.AddItem rsTemp("序号") & "." & rsTemp("名称")
        cbo类别.ItemData(cbo类别.NewIndex) = rsTemp("序号")
        rsTemp.MoveNext
    Loop
    cbo类别.ListIndex = 0
    
    '缺省值
    lng卡号长度 = 20
    lng退休证长度 = 26
    
    gstrSQL = "select 参数名,参数值 from 保险参数 where 险类=" & TYPE_重庆中梁山 & " and 中心=0"
    Call OpenRecordset(rsTemp, Me.Caption)
    
    Do Until rsTemp.EOF
        Select Case rsTemp("参数名")
            Case "卡号长度"
                If IsNull(rsTemp("参数值")) = False Then lng卡号长度 = Val(rsTemp("参数值"))
            Case "退休证长度"
                If IsNull(rsTemp("参数值")) = False Then lng退休证长度 = Val(rsTemp("参数值"))
        End Select
        rsTemp.MoveNext
    Loop
    
    txtEdit(Text卡号).MaxLength = lng卡号长度
    txtEdit(Text退休证号).MaxLength = lng退休证长度
    
    dtp生日.MaxDate = zlDatabase.Currentdate
    frmIdentify中梁山.Show vbModal
    ShowCard = mstrIdentify
End Function

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
    If cbo类别.ListIndex < 0 Then
        MsgBox "请选择病人类别。", vbInformation, gstrSysName
        cbo类别.SetFocus
        Exit Sub
    End If
    lng中心 = 0
    
    '检查病人状态
    gstrSQL = "select nvl(当前状态,0) as 状态 from 保险帐户 where 险类=" & TYPE_重庆中梁山 & " and 中心=" & lng中心 & " and 医保号='" & Trim(txtEdit(Text卡号).Text) & "'"
    Call OpenRecordset(rsTemp, Me.Caption)
    
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
    strIdentify = Trim(txtEdit(Text卡号).Text)                         '0卡号
    strIdentify = strIdentify & ";" & Trim(txtEdit(Text卡号).Text)     '1医保号 使用相同号码
    strIdentify = strIdentify & ";"                                    '2密码
    strIdentify = strIdentify & ";" & Trim(txtEdit(Text姓名).Text)     '3姓名
    strIdentify = strIdentify & ";" & Replace(GetTextFromCombo(cbo性别, True), "'", "") '4性别
    strIdentify = strIdentify & ";" & Format(dtp生日.Value, "yyyy-MM-dd") '5出生日期
    strIdentify = strIdentify & ";" & Trim(txtEdit(Text身份证号).Text)    '6身份证
    strIdentify = strIdentify & ";" & "()"                                '7.单位名称(编码)
    strAddition = ";" & lng中心                                           '8.中心代码
    strAddition = strAddition & ";"                                       '9.顺序号
    strAddition = strAddition & ";"                                       '10人员身份
    strAddition = strAddition & ";0"                                      '11帐户余额
    strAddition = strAddition & ";0"                                      '12当前状态
    strAddition = strAddition & ";" & txtEdit(Text病种).Tag               '13病种ID
    strAddition = strAddition & ";" & cbo类别.ItemData(cbo类别.ListIndex) '14在职(1,2,3)
    strAddition = strAddition & ";" & Trim(txtEdit(Text退休证号).Text)    '15退休证号
    strAddition = strAddition & ";" & DateDiff("yyyy", dtp生日.Value, dtp生日.MaxDate) '16年龄段
    strAddition = strAddition & ";"                             '17灰度级
    strAddition = strAddition & ";0"                                      '18帐户增加累计
    strAddition = strAddition & ";0"                                      '19帐户支出累计
    strAddition = strAddition & ";0"                                      '20进入统筹累计
    strAddition = strAddition & ";" & Val(txtEdit(Text统筹报销累计).Text) '21统筹报销累计
    strAddition = strAddition & ";0"                                      '22住院次数累计
    strAddition = strAddition & ";"                                       '23就诊类型 (1、急诊门诊)
    
    lng病人ID = BuildPatiInfo(mbytType, strIdentify & strAddition, mlng病人ID, TYPE_重庆中梁山)
    '返回格式:中间插入病人ID
    If lng病人ID > 0 Then
        mstrIdentify = strIdentify & ";" & lng病人ID & strAddition
    End If
    Unload Me
End Sub

Private Function IsValid() As Boolean
'功能：检查数据的正确性
    Dim lngIndex As Long
    
    For lngIndex = txtEdit.LBound To txtEdit.UBound
        If txtEdit(lngIndex).Enabled = True Then
            If zlCommFun.StrIsValid(txtEdit(lngIndex), txtEdit(lngIndex).MaxLength) = False Then
                zlControl.TxtSelAll txtEdit(lngIndex)
                txtEdit(lngIndex).SetFocus
                Exit Function
            End If
        End If
    Next
    If Len(txtEdit(Text卡号).Text) <> txtEdit(Text卡号).MaxLength Then
        MsgBox "卡号长度不足" & txtEdit(Text卡号).MaxLength & "位。", vbInformation, gstrSysName
        zlControl.TxtSelAll txtEdit(Text卡号)
        txtEdit(Text卡号).SetFocus
        Exit Function
    End If
    If Trim(txtEdit(Text姓名).Text) = "" Then
        MsgBox "姓名不能为空。", vbInformation, gstrSysName
        zlControl.TxtSelAll txtEdit(Text姓名)
        txtEdit(Text姓名).SetFocus
        Exit Function
    End If
    
    If IsNumeric(txtEdit(Text统筹报销累计).Text) = False Then
        MsgBox "统筹报销累计输入合法的数值。", vbInformation, gstrSysName
        zlControl.TxtSelAll txtEdit(Text统筹报销累计)
        txtEdit(Text统筹报销累计).SetFocus
        Exit Function
    End If
    
    If Val(txtEdit(Text统筹报销累计).Text) < 0 Or Val(txtEdit(Text统筹报销累计).Text) > 1000000 Then
        MsgBox "金额不能小于0，且不能超过100万。", vbInformation, gstrSysName
        zlControl.TxtSelAll txtEdit(Text统筹报销累计)
        txtEdit(Text统筹报销累计).SetFocus
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
                    "  where A.病人ID=B.病人ID and A.险类=" & TYPE_重庆中梁山 & _
                    "  and A.险类=C.险类 and A.中心=C.序号 and A.病种ID=D.ID(+)"
            
            Call Get帐户情况
            zlControl.TxtSelAll txtEdit(Text卡号)
            txtEdit(Text卡号).SetFocus
        Case Select病种
            gstrSQL = " Select A.ID,A.编码,A.名称,A.简码,decode(A.类别,1,'慢性病',2,'特种病','普通病') as 类别 " & _
                    " From 保险病种 A where A.险类=" & TYPE_重庆中梁山
            
            Set rsTemp = frmPubSel.ShowSelect(Me, gstrSQL, 0, "医保病种", , txtEdit(Text病种).Text)
            If Not rsTemp Is Nothing Then
                txtEdit(Text病种).Text = rsTemp("名称")
                txtEdit(Text病种).Tag = rsTemp("ID")
                zlControl.TxtSelAll txtEdit(Text病种)
            End If
            txtEdit(Text病种).SetFocus
    End Select
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
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
            txtEdit(Text病种).Text = ""
            txtEdit(Text病种).Tag = ""
        End If
    End If
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strCode As String
    Dim str条件 As String
    Dim rsTemp As New ADODB.Recordset
    
    If Index = Text卡号 Then
        If Len(txtEdit(Text卡号).Text) = txtEdit(Text卡号).MaxLength Or KeyAscii = vbKeyReturn Then
            strCode = Replace(Trim(txtEdit(Text卡号).Text), "'", "")
            
            If IsNumeric(strCode) And IsNumeric(Left(strCode, 1)) Then '刷卡
                str条件 = " and A.卡号='" & strCode & "'"
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
                    "  where A.病人ID=B.病人ID and A.险类=" & TYPE_重庆中梁山 & _
                    "  and A.险类=C.险类 and A.中心=C.序号 and A.病种ID=D.ID(+)" & str条件
            
            Call Get帐户情况
        End If
    ElseIf KeyAscii = asc("*") Then
        Call cmdSelect_Click(Select病种)
    End If
    
    If KeyAscii = vbKeyReturn Then
        If Index = Text身份证号 Then
            strCode = Get出生日期(txtEdit(Text身份证号).Text, 0)
            If IsDate(strCode) = True Then
                dtp生日.Value = CDate(strCode)
            End If
        End If
        KeyAscii = 0  '消除响声
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub cbo类别_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub cbo性别_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    If Index = Text统筹报销累计 Then
        txtEdit(Index).Text = Format(Val(txtEdit(Index).Text), "######0.00;0.00;0.00;0.00")
    End If

End Sub

Private Sub Get帐户情况()
'从已经存在的记录中读出帐户信息
    Dim rs帐户 As ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim lngIndex As Long
    
    
    Set rs帐户 = frmPubSel.ShowSelect(Me, gstrSQL, 0, "保险帐户", , txtEdit(Text卡号).Text, "", False, True)
    If Not rs帐户 Is Nothing Then
    
        txtEdit(Text卡号).Text = rs帐户("卡号")
        '其它可用的数据
        txtEdit(Text姓名).Text = IIf(IsNull(rs帐户("姓名")), "", rs帐户("姓名"))
        txtEdit(Text身份证号).Text = IIf(IsNull(rs帐户("身份证号")), "", rs帐户("身份证号"))
        txtEdit(Text病种).Text = IIf(IsNull(rs帐户("病种")), "", rs帐户("病种"))
        txtEdit(Text病种).Tag = IIf(IsNull(rs帐户("病种ID")), "", rs帐户("病种ID"))
        
        Call SetComboByText(cbo性别, IIf(IsNull(rs帐户("性别")), "", rs帐户("性别")), True)
        txtEdit(Text退休证号).Text = ""
        If IsNull(rs帐户("出生日期")) = False Then
            dtp生日.Value = rs帐户("出生日期")
        End If
        
        For lngIndex = 0 To cbo类别.ListCount - 1
            If cbo类别.ItemData(lngIndex) = rs帐户("在职ID") Then
                cbo类别.ListIndex = lngIndex
                Exit For
            End If
        Next
        
        '再读出帐户年度信息
        gstrSQL = "select * from 帐户年度信息 where 险类=" & TYPE_重庆中梁山 & _
            " and 病人ID=" & rs帐户("ID") & " and 年度=" & Year(dtp生日.MaxDate)
        Call OpenRecordset(rsTemp, Me.Caption)
        
        If rsTemp.EOF = False Then
            '设置帐户情况
            txtEdit(Text统筹报销累计).Text = Format(rsTemp("统筹报销累计"), "######0.00;0.00;0.00;0.00")
        Else
            txtEdit(Text统筹报销累计).Text = "0.00"
        End If
        
    End If
End Sub

