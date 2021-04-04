VERSION 5.00
Begin VB.Form frm病种选择_沈阳 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病种选择"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   Icon            =   "frm病种选择_沈阳.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txt并发症 
      Height          =   300
      Left            =   1350
      TabIndex        =   9
      Top             =   2010
      Width           =   3675
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2610
      TabIndex        =   11
      Top             =   2550
      Width           =   1100
   End
   Begin VB.CommandButton cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3870
      TabIndex        =   12
      Top             =   2550
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   30
      TabIndex        =   10
      Top             =   2430
      Width           =   6075
   End
   Begin VB.TextBox txt疾病信息 
      Height          =   300
      Index           =   1
      Left            =   1350
      TabIndex        =   5
      Top             =   1170
      Width           =   3375
   End
   Begin VB.CommandButton cmd疾病信息 
      Caption         =   "…"
      Height          =   300
      Index           =   1
      Left            =   4740
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1170
      Width           =   285
   End
   Begin VB.TextBox txt疾病信息 
      Height          =   300
      Index           =   0
      Left            =   1350
      TabIndex        =   2
      Top             =   780
      Width           =   3375
   End
   Begin VB.CommandButton cmd疾病信息 
      Caption         =   "…"
      Height          =   300
      Index           =   0
      Left            =   4740
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   780
      Width           =   285
   End
   Begin VB.Label lbl并发症 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "并发症(&M)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   450
      TabIndex        =   8
      Top             =   2070
      Width           =   810
   End
   Begin VB.Label lblDemo 
      Caption         =   "    还没有为该病人设置出院病种，所以出院病种缺省为入院病种，点确定将保存出院病种"
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   1560
      Width           =   3915
   End
   Begin VB.Label lblPatient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "姓名：沈阳医保    卡号：01234567    "
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   405
      Left            =   330
      TabIndex        =   0
      Top             =   180
      Width           =   4785
   End
   Begin VB.Label lbl疾病信息 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "出院病种(&U)"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   1
      Left            =   300
      TabIndex        =   4
      Top             =   1230
      Width           =   990
   End
   Begin VB.Label lbl疾病信息 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "入院病种(&R)"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   0
      Left            =   300
      TabIndex        =   1
      Top             =   840
      Width           =   990
   End
End
Attribute VB_Name = "frm病种选择_沈阳"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mblnStart As Boolean
Private mint险类 As Integer
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mstr入院病种 As String
Private mstr出院病种 As String
Private mstr并发症 As String
Private Enum 病种
    入院病种
    出院病种
End Enum

Private Sub cmdOK_Click()
    On Error GoTo errHand
    '必须选择病种信息
    If txt疾病信息(入院病种).Tag = "" Then
        MsgBox "请为该参保病人选择疾病编码信息！", vbInformation, gstrSysName
        txt疾病信息(入院病种).SetFocus
        Exit Sub
    End If
    If txt疾病信息(出院病种).Tag = "" Then
        MsgBox "请为该参保病人选择疾病编码信息！", vbInformation, gstrSysName
        txt疾病信息(出院病种).SetFocus
        Exit Sub
    End If
    
    gstrSQL = "ZL_保险帐户_更新信息(" & mlng病人ID & "," & mint险类 & ",'病种ID','" & Split(txt疾病信息(入院病种).Tag, "|")(1) & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "更新入院病种")
    gstrSQL = "ZL_保险帐户_更新信息(" & mlng病人ID & "," & mint险类 & ",'出院病种ID','" & Split(txt疾病信息(出院病种).Tag, "|")(1) & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "更新出院病种")
    gstrSQL = "ZL_保险帐户_更新信息(" & mlng病人ID & "," & mint险类 & ",'并发症','''" & txt并发症.Text & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "更新并发症")
    
    mblnOK = True
    mstr入院病种 = Split(txt疾病信息(入院病种).Tag, "|")(0)
    mstr出院病种 = Split(txt疾病信息(出院病种).Tag, "|")(0)
    mstr并发症 = txt并发症.Text
    
    Unload Me
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cmd疾病信息_Click(Index As Integer)
    Dim rs病种 As New ADODB.Recordset
    gstrSQL = " Select A.ID,A.编码,A.名称,A.简码 " & _
            " From 保险病种 A where A.险类=[1] Order by A.简码"
    Set rs病种 = New ADODB.Recordset
    Set rs病种 = zlDatabase.OpenSQLRecord(gstrSQL, "身份验证", TYPE_沈阳市)
    If rs病种.RecordCount > 0 Then
        If frmListSel.ShowSelect(TYPE_沈阳市, rs病种, "ID", "医保病种选择", "请选择" & IIf(Index = 0, "入院", "出院") & "病种：") = True Then
            txt疾病信息(Index).Tag = rs病种!编码 & "|" & rs病种!ID
            txt疾病信息(Index).Text = "(" & rs病种!编码 & ")" & rs病种!名称
            lbl疾病信息(Index).Tag = txt疾病信息(Index).Text '用于恢复显示
            
            If txt疾病信息(出院病种).Tag = "" Then
                txt疾病信息(出院病种).Text = "(" & rs病种!编码 & ")" & rs病种!名称
                txt疾病信息(出院病种).Tag = rs病种!编码 & "|" & rs病种!ID
                lbl疾病信息(出院病种).Tag = txt疾病信息(入院病种).Text
            End If
        End If
    End If
End Sub

Private Sub cmd取消_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If Not mblnStart Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim blnSet As Boolean               '说明是否设置出院病种
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '读取该病人的基本信息
    gstrSQL = " Select B.姓名,A.卡号,A.医保号,C.ID 入院病种ID,C.编码 入院病种编码,C.名称 入院病种名称," & _
              " D.ID 出院病种ID,D.编码 出院病种编码,D.名称 出院病种名称,A.并发症" & _
              " From 保险帐户 A,病人信息 B," & _
              " (Select * From 保险病种 Where 险类=" & mint险类 & ") C," & _
              " (Select * From 保险病种 Where 险类=" & mint险类 & ") D" & _
              " Where A.病人ID=B.病人ID And A.病人ID=[1] And A.险类=[2]" & _
              " And C.ID(+)=A.病种ID And D.ID(+)=A.出院病种ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取该病人的基本信息", mlng病人ID, mint险类)
    
    lblPatient.Caption = "姓名：" & Nvl(rsTemp!姓名) & Space(4) & "卡号：" & Nvl(rsTemp!卡号) & Space(4) & "个人编号：" & Nvl(rsTemp!医保号)
    txt并发症.Text = Nvl(rsTemp!并发症)
    If Not IsNull(rsTemp!入院病种编码) Then
        txt疾病信息(入院病种).Text = "(" & rsTemp!入院病种编码 & ")" & rsTemp!入院病种名称
        txt疾病信息(入院病种).Tag = rsTemp!入院病种编码 & "|" & rsTemp!入院病种ID
        lbl疾病信息(入院病种).Tag = txt疾病信息(入院病种).Text
    End If
    If Not IsNull(rsTemp!出院病种编码) Then
        blnSet = True
        txt疾病信息(出院病种).Text = "(" & rsTemp!出院病种编码 & ")" & rsTemp!出院病种名称
        txt疾病信息(出院病种).Tag = rsTemp!出院病种编码 & "|" & rsTemp!出院病种ID
        lbl疾病信息(出院病种).Tag = txt疾病信息(出院病种).Text
    Else
        blnSet = False
        If Not IsNull(rsTemp!入院病种编码) Then
            txt疾病信息(出院病种).Text = "(" & rsTemp!入院病种编码 & ")" & rsTemp!入院病种名称
            txt疾病信息(出院病种).Tag = rsTemp!入院病种编码 & "|" & rsTemp!入院病种ID
            lbl疾病信息(出院病种).Tag = txt疾病信息(入院病种).Text
        End If
    End If
    
    '如果未设置出院病种，调整窗体大小
    If blnSet Then
        Me.lblDemo.Visible = False
        Me.lbl并发症.Top = Me.lbl疾病信息(出院病种).Top - Me.lbl疾病信息(入院病种).Top + Me.lbl疾病信息(出院病种).Top
        Me.txt并发症.Top = Me.txt疾病信息(出院病种).Top - Me.txt疾病信息(入院病种).Top + Me.txt疾病信息(出院病种).Top
    End If
    mblnStart = True
    Exit Sub
errHand:
    MsgBox "请确认保险帐户表的结构是最新的！", vbInformation, gstrSysName
End Sub

Public Function ShowSelect(ByVal int险类 As Integer, ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
    ByRef str入院病种 As String, ByRef str出院病种 As String, ByRef str并发症 As String) As Boolean
    '选择病人的入院病种及出院病种，同时将病人本次住院的相关信息显示出来
    '更新保险帐户的病种ID（入院病种）及出院病种，并将入院病种及出院病种编码返回给调用模块
    mblnOK = False
    mint险类 = int险类
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    Me.Show 1
    str入院病种 = mstr入院病种
    str出院病种 = mstr出院病种
    str并发症 = mstr并发症
    ShowSelect = mblnOK
End Function

Private Sub txt并发症_GotFocus()
    Call zlControl.TxtSelAll(txt并发症)
End Sub

Private Sub txt并发症_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdOK.SetFocus
End Sub

Private Sub txt疾病信息_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txt疾病信息(Index))
End Sub

Private Sub txt疾病信息_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim rsTemp As New ADODB.Recordset, strText As String
    Dim blnReturn As Boolean
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If txt疾病信息(Index).Text = "" And txt疾病信息(Index).Tag <> "" Then Exit Sub
    
    On Error GoTo errHandle
    strText = txt疾病信息(Index).Text
    If InStr(1, strText, "(") <> 0 Then
        If InStr(1, strText, ")") <> 0 Then
            strText = Mid(strText, 2, InStr(1, strText, ")") - 2)
        End If
    End If
    
    gstrSQL = "Select A.ID,A.编码,A.名称,A.简码" & _
             "   FROM 保险病种 A WHERE A.险类=[1] And (" & _
             " A.编码 like [2] || '%'  or  A.名称 like [2] || '%'   or  A.简码 like [2] || '%')" & _
             " Order by A.简码"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mint险类, strText)
    If rsTemp.RecordCount = 0 Then
        MsgBox "不存在该病种，请重新输入！", vbInformation, gstrSysName
        txt疾病信息(Index).Text = lbl疾病信息(Index).Tag
        zlControl.TxtSelAll txt疾病信息(Index)
        Exit Sub
    Else
        '出现选择器
        If rsTemp.RecordCount > 1 Then
            '对于字段大于3的，即使只有一条记录把该对话框显示出来，以便让用户得到更多的信息
            blnReturn = frmListSel.ShowSelect(TYPE_沈阳市, rsTemp, "ID", "医保病种选择", "请选择医保病种：")
        Else
            blnReturn = True
        End If
    End If
    
    If blnReturn = False Then
        '记录集中没有可选择的数据
        txt疾病信息(Index).Text = lbl疾病信息(Index).Tag
        zlControl.TxtSelAll txt疾病信息(Index)
    Else
        '肯定是有记录集的
        txt疾病信息(Index).Tag = rsTemp!编码 & "|" & rsTemp!ID
        txt疾病信息(Index).Text = "(" & rsTemp!编码 & ")" & rsTemp!名称
        lbl疾病信息(Index).Tag = txt疾病信息(Index).Text '用于恢复显示
            
        If txt疾病信息(出院病种).Tag = "" Then
            txt疾病信息(出院病种).Text = "(" & rsTemp!编码 & ")" & rsTemp!名称
            txt疾病信息(出院病种).Tag = rsTemp!编码 & "|" & rsTemp!ID
            lbl疾病信息(出院病种).Tag = txt疾病信息(入院病种).Text
        End If
        
        If Index = 0 Then
            txt疾病信息(1).SetFocus
        Else
            txt并发症.SetFocus
        End If
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
