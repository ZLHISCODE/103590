VERSION 5.00
Begin VB.Form frmIdentify成都南充 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "身份验证"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4530
   Icon            =   "frmIdentify成都南充.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton CmdSelect 
      Caption         =   "…"
      Height          =   285
      Index           =   0
      Left            =   2580
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   330
      Width           =   285
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3330
      TabIndex        =   8
      Top             =   870
      Width           =   1100
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3330
      TabIndex        =   7
      Top             =   420
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   3120
      TabIndex        =   9
      Top             =   -180
      Width           =   30
   End
   Begin VB.ComboBox Cbo性别 
      Height          =   300
      Left            =   1140
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   1
      Left            =   1140
      MaxLength       =   3
      TabIndex        =   6
      Top             =   1110
      Width           =   1725
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   0
      Left            =   1140
      MaxLength       =   20
      TabIndex        =   1
      Top             =   330
      Width           =   1725
   End
   Begin VB.Label lbl年龄 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "年龄(&A)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   420
      TabIndex        =   5
      Top             =   1170
      Width           =   630
   End
   Begin VB.Label Lbl性别 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "性别(&S)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   420
      TabIndex        =   3
      Top             =   780
      Width           =   630
   End
   Begin VB.Label Lbl姓名 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "姓名(&N)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   420
      TabIndex        =   0
      Top             =   390
      Width           =   630
   End
End
Attribute VB_Name = "frmIdentify成都南充"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum 文本Enum
    Text姓名 = 0
    Text年龄 = 1
End Enum

Private Enum 选择Enum
    Select姓名 = 0
End Enum

Dim mstrIdentify As String
Dim mbytType As Byte
Dim mlng病人ID As Long

Public Function ShowCard(Optional bytType As Byte, Optional lng病人ID As Long) As String
'功能：返回医保病人的身份信息
'参数：bytType-识别类型，0-门诊，1-住院
'返回：
    '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);
    '8.中心代码;9.顺序号;10人员身份;11帐户余额;12当前状态;13病种ID;14在职(0,1);15退休证号;16年龄段;17灰度级
    '18帐户增加累计,19帐户支出累计,20进入统筹累计,21统筹报销累计,22住院次数累计
    Dim rsTemp As New ADODB.Recordset
    mbytType = bytType
    mlng病人ID = lng病人ID
    
    cbo性别.Clear
    gstrSQL = "select 编码,名称 from 性别 order by 编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    Do Until rsTemp.EOF
        cbo性别.AddItem rsTemp("编码") & "." & rsTemp("名称")
        rsTemp.MoveNext
    Loop
    cbo性别.ListIndex = 0
    rsTemp.Close
    
    frmIdentify成都南充.Show vbModal
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
    lng中心 = 0
    
    '检查病人状态
    gstrSQL = "select nvl(当前状态,0) as 状态 from 保险帐户 where 险类=[1] and 中心=[2] and 病人ID=[3]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_成都南充, lng中心, mlng病人ID)
    
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
    strIdentify = ""                         '0卡号
    strIdentify = strIdentify & ";"          '1医保号
    strIdentify = strIdentify & ";"                                    '2密码
    strIdentify = strIdentify & ";" & Trim(txtEdit(Text姓名).Text)     '3姓名
    strIdentify = strIdentify & ";" & Replace(GetTextFromCombo(cbo性别, True), "'", "") '4性别
    strIdentify = strIdentify & ";" & "" '5出生日期
    strIdentify = strIdentify & ";" & ""    '6身份证
    strIdentify = strIdentify & ";" & ""  '7.单位名称(编码)
    strAddition = ";" & lng中心                                 '8.中心代码
    strAddition = strAddition & ";"                             '9.顺序号
    strAddition = strAddition & ";" & ""       '10人员身份
    strAddition = strAddition & ";" & ""  '11帐户余额
    strAddition = strAddition & ";0"                            '12当前状态
    strAddition = strAddition & ";" & ""     '13病种ID
    strAddition = strAddition & ";" & "1" '14在职(1,2,3)
    strAddition = strAddition & ";" & "" '15退休证号
    strAddition = strAddition & ";" & Val(txtEdit(Text年龄)) '16年龄段
    strAddition = strAddition & ";"                             '17灰度级
    strAddition = strAddition & ";" & 0       '18帐户增加累计
    strAddition = strAddition & ";" & 0       '19帐户支出累计
    strAddition = strAddition & ";" & 0        '20上年工资总额
    strAddition = strAddition & ";0;0"      '21住院次数累计
    
    lng病人ID = BuildPatiInfo(mbytType, strIdentify & strAddition, mlng病人ID, TYPE_成都南充)
    '返回格式:中间插入病人ID
    If lng病人ID > 0 Then
        mstrIdentify = strIdentify & ";" & lng病人ID & strAddition
        '强制把登记顺序号、及新的医保号填入
        gstrSQL = "ZL_保险帐户_修改医保号(" & lng病人ID & "," & TYPE_成都南充 & _
                    ",NULL,'" & lng病人ID & "',NULL)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
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
        
        If lngIndex = Text年龄 Then
            If IsNumeric(txtEdit(lngIndex).Text) = False Then
                MsgBox "请输入合法的数值。", vbInformation, gstrSysName
                zlControl.TxtSelAll txtEdit(lngIndex)
                txtEdit(lngIndex).SetFocus
                Exit Function
            End If
            If Val(txtEdit(lngIndex).Text) < 0 Or Val(txtEdit(lngIndex).Text) > 200 Then
                MsgBox "年龄不能小于0，且不能超过200。", vbInformation, gstrSysName
                zlControl.TxtSelAll txtEdit(lngIndex)
                txtEdit(lngIndex).SetFocus
                Exit Function
            End If
        End If
        If (lngIndex = Text姓名 Or lngIndex = Text年龄) And Trim(txtEdit(lngIndex).Text) = "" Then
            MsgBox "姓名、年龄都不能为空。", vbInformation, gstrSysName
            zlControl.TxtSelAll txtEdit(lngIndex)
            txtEdit(lngIndex).SetFocus
            Exit Function
        End If
    Next
    
    IsValid = True
End Function

Private Sub cmdSelect_Click(Index As Integer)
    Dim rsTemp As ADODB.Recordset
    
    Select Case Index
        Case Select姓名
            gstrSQL = " Select A.病人ID as ID,A.卡号,A.医保号,B.姓名,B.性别,B.年龄,B.出生日期,B.身份证号,C.序号 as 中心ID " & _
                    " ,A.人员身份,A.单位编码,A.病种ID,D.名称 as 病种,A.在职 as 在职ID,A.退休证号" & _
                    " From 保险帐户 A,病人信息 B,保险中心目录 C,保险病种 D" & _
                    "  where A.病人ID=B.病人ID and A.险类=" & TYPE_成都南充 & _
                    "  and A.险类=C.险类 and A.中心=C.序号 and A.病种ID=D.ID(+)"
            
            Call Get帐户情况
            zlControl.TxtSelAll txtEdit(Text姓名)
            txtEdit(Text姓名).SetFocus
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

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strCode As String
    Dim str条件 As String
    Dim rsTemp As New ADODB.Recordset
    
    If KeyAscii = vbKeyReturn Then
        
        KeyAscii = 0  '消除响声
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub cbo性别_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub Get帐户情况()
'从已经存在的记录中读出帐户信息
    Dim rs帐户 As ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim lngIndex As Long
    
    
    Set rs帐户 = frmPubSel.ShowSelect(Me, gstrSQL, 0, "保险帐户", , txtEdit(Text姓名).Text, "", False, True)
    If Not rs帐户 Is Nothing Then
    
        '其它可用的数据
        mlng病人ID = rs帐户!ID
        txtEdit(Text姓名).Text = IIf(IsNull(rs帐户("姓名")), "", rs帐户("姓名"))
        txtEdit(Text年龄).Text = IIf(IsNull(rs帐户("年龄")), "", rs帐户("年龄"))
        
        Call SetComboByText(cbo性别, IIf(IsNull(rs帐户("性别")), "", rs帐户("性别")), True)
    End If
End Sub



