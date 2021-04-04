VERSION 5.00
Begin VB.Form frm诊疗项目编码选择 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "诊疗项目编码选择"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5715
   Icon            =   "frm诊疗项目编码选择.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmd病种 
      Caption         =   "…"
      Height          =   285
      Left            =   4860
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1395
      Width           =   255
   End
   Begin VB.Frame fra 
      Height          =   30
      Index           =   1
      Left            =   0
      TabIndex        =   6
      Top             =   915
      Width           =   5715
   End
   Begin VB.Frame fra 
      Height          =   30
      Index           =   0
      Left            =   15
      TabIndex        =   5
      Top             =   2355
      Width           =   5715
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3090
      TabIndex        =   3
      Top             =   2550
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4335
      TabIndex        =   2
      Top             =   2535
      Width           =   1100
   End
   Begin VB.TextBox txt项目 
      Height          =   300
      Left            =   1230
      TabIndex        =   1
      Top             =   1395
      Width           =   3870
   End
   Begin VB.Label lbl 
      Caption         =   "选择挂号中的诊疗费与中心的医保项目进行对码。"
      Height          =   225
      Index           =   0
      Left            =   825
      TabIndex        =   7
      Top             =   510
      Width           =   4965
   End
   Begin VB.Image img 
      Height          =   555
      Left            =   45
      Picture         =   "frm诊疗项目编码选择.frx":000C
      Stretch         =   -1  'True
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "项目编码"
      Height          =   180
      Index           =   1
      Left            =   450
      TabIndex        =   4
      Top             =   1470
      Width           =   720
   End
End
Attribute VB_Name = "frm诊疗项目编码选择"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnOK As Boolean
Dim mcnTest As New ADODB.Connection
Private mblnChange As Boolean
Dim mblnFirst As Boolean
Dim mstrCode As String
 
Public Function ShowCard(strCode As String) As Boolean
    mblnChange = False
    
    Me.Show vbModal
    ShowCard = mblnOK
    strCode = mstrCode
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cmdCancel_Click()
    mstrCode = ""
    Unload Me
End Sub


Private Sub cmd病种_Click()
        '刘兴宏:20040706
        Dim strCode As String
        Dim STRNAME As String
        
        On Error Resume Next
        If frm保险项目选择重庆渝北.GetCode(Me, strCode, STRNAME, True) = True Then
            Me.txt项目.Text = strCode & "-" & STRNAME
            Me.txt项目.Tag = strCode
            If cmdOK.Enabled Then cmdOK.SetFocus
        End If
End Sub

Private Sub Form_Activate()
    Dim rsTemp As New ADODB.Recordset
    
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    gstrSQL = "Select * From 保险参数 where 险类=" & TYPE_重庆渝北 & " and 参数名='诊疗项目编码'"
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    With rsTemp
        Do While Not .EOF
            Select Case Nvl(!参数名)
            Case "诊疗项目编码"
                  txt项目.Text = Nvl(!参数值)
                  txt项目.Tag = txt项目.Text
            End Select
            .MoveNext
        Loop
    End With
 End Sub

Private Sub Form_Load()
    mblnFirst = True
End Sub


Private Sub cmdOK_Click()
    
    If IsValid = False Then Exit Sub
    mstrCode = txt项目.Tag
    mblnOK = True
    mblnChange = False
    Unload Me
End Sub

Private Function IsValid() As Boolean
    Dim lngCount As Long
    Dim strTitle As String
    Dim rsTemp As New ADODB.Recordset
    If txt项目.Tag = "" Then
        ShowMsgbox "诊疗项目所对应的医保项目未选择!"
        txt项目.SetFocus
        Exit Function
    End If
        
    IsValid = True
End Function


Private Sub txt项目_Change()
    txt项目.Tag = ""
End Sub

Private Sub txt项目_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    Dim rsTemp As New ADODB.Recordset
    Dim strLeft As String
    Dim strTemp As String
    Dim blnReturn As Boolean
    If txt项目.Text = "" Then Exit Sub
    strLeft = IIf(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = 0, "%", "")
    strTemp = "'" & strLeft & txt项目.Text & "%'"
    
    gstrSQL = " select  商品代码 as 医保编码,  医院大类编码, 药品通用中文名, 药品通用英文名,商品名, 商品曾用名, 服务项目结算方式, 报销标识, 医保标识, 是否处方用药, 药品适应症, 限制医生, 审批权限, 别名, 包装规格, " & _
             "         最小包装单位, 最小计量单位, 每日最大用量, 指导价格, 招标价格, 基金支付限价1, 基金支付限价2, 基金支付限价3, 实际执行价格, 自付比例1, 自付比例2, 自付比例3, 自付比例4, 自付比例5, 自付比例6, 自付比例7, 自付比例8,  " & _
             "         自付比例9, 自付比例10, 自付比例11, 自付比例12, 医院使用状态, 中心使用状态, 标准编号,  " & _
             "         五笔助记码1, 五笔助记码2, 五笔助记码3, 拼音助记码1, 拼音助记码2, 拼音助记码3, 备注, 医保经办机构,机构标准编号, 医疗机构编号, " & _
             "          修改时间, 目录分类  " & _
             "  from 医保服务项目目录" & _
             "  where 医院大类编码='61' and ( 商品代码 like " & strTemp & " Or 商品名 like " & strTemp & " Or " & _
             "        五笔助记码1 like " & UCase(strTemp) & " Or " & _
             "        拼音助记码1 like " & UCase(strTemp) & ")"
    If gcnOracle_CQYB.State = adStateOpen Then
        rsTemp.Open gstrSQL, gcnOracle_CQYB, adOpenStatic, adLockReadOnly
    Else
        '强制使记录集为打开状态
        gstrSQL = "Select 编码  医保编码,名称,简码 FROM 保险项目 Where Rownum<1"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    End If
                   
    If rsTemp.RecordCount > 0 Then
        '出现选择器
        If rsTemp.RecordCount > 1 Then
            '对于字段大于3的，即使只有一条记录把该对话框显示出来，以便让用户得到更多的信息
            blnReturn = frmListSel.ShowSelect(TYPE_重庆渝北, rsTemp, "医保编码", "医保项目选择", "请选择对应的医保项目：")
        Else
            blnReturn = True
        End If
    Else
        MsgBox "无此项目!"
        Exit Sub
    End If
    
    If blnReturn = False Then Exit Sub
        '肯定是有记录集的
    txt项目.Text = rsTemp("医保编码") & "-" & Nvl(rsTemp!商品名)
    txt项目.Tag = rsTemp("医保编码")
    If cmdOK.Enabled Then cmdOK.SetFocus
End Sub

Private Sub txt项目_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txt项目, KeyAscii, m文本式
End Sub


