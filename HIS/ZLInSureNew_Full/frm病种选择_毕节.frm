VERSION 5.00
Begin VB.Form frm病种选择_毕节 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病种选择"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4590
   Icon            =   "frm病种选择_毕节.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -90
      TabIndex        =   3
      Top             =   1290
      Width           =   4965
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2910
      TabIndex        =   5
      Top             =   1500
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   1620
      TabIndex        =   4
      Top             =   1500
      Width           =   1100
   End
   Begin VB.CommandButton cmd病种 
      Appearance      =   0  'Flat
      Caption         =   "…"
      Height          =   270
      Left            =   3750
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   810
      Width           =   285
   End
   Begin VB.TextBox txt病种 
      Height          =   300
      Left            =   1230
      TabIndex        =   1
      Top             =   810
      Width           =   2835
   End
   Begin VB.Label lblNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   195
      Index           =   1
      Left            =   540
      TabIndex        =   7
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label lblNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   195
      Index           =   0
      Left            =   540
      TabIndex        =   6
      Top             =   210
      Width           =   3495
   End
   Begin VB.Label lbl病种 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "病种(&D)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   510
      TabIndex        =   0
      Top             =   870
      Width           =   630
   End
End
Attribute VB_Name = "frm病种选择_毕节"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mlng病人ID As Long
Private mlng病种ID As Long
Private mstr病种名称 As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    mlng病种ID = Val(txt病种.Tag)
    If mlng病种ID = 0 Then
        MsgBox "必须要选择一个病种！", vbInformation, gstrSysName
        txt病种.SetFocus
        Exit Sub
    End If
    
    If InStr(1, txt病种.Text, ")") <> 0 Then
        mstr病种名称 = Mid(txt病种.Text, InStr(1, txt病种.Text, ")") + 1)
    End If
    
    mblnOK = True
    Unload Me
End Sub

Public Function 病种选择(ByVal lng病人ID As String, lng病种ID As Long, str病种名称 As String) As Boolean
    mlng病人ID = lng病人ID
    mlng病种ID = 0
    mstr病种名称 = ""
    mblnOK = False
    Me.Show 1
    If mblnOK Then
        lng病种ID = mlng病种ID
        str病种名称 = mstr病种名称
    End If
    病种选择 = mblnOK
End Function

Private Sub Form_Load()
    Dim lng病种ID As Long
    Dim str病种名称 As String
    Dim STR姓名 As String
    Dim str社会保障号 As String
    Dim rsTemp As New ADODB.Recordset
    '将以前选择的病种显示出来
    gstrSQL = "Select 病种ID,医保号 From 保险帐户 Where 险类=[1] And 病人ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取当前选择的病种ID", TYPE_毕节, mlng病人ID)
    lng病种ID = Nvl(rsTemp!病种ID, 0)
    str社会保障号 = Nvl(rsTemp!医保号)
    
    '读取该病种的信息
    gstrSQL = "Select 病种代码,病种名称 From 病种目录表 Where ID=" & lng病种ID
    With rsTemp
        If .State = 1 Then .Close
        .Open gstrSQL, gcnGYBJYB
        If .RecordCount <> 0 Then
            Me.txt病种.Text = "(" & rsTemp!病种代码 & ")" & rsTemp!病种名称
            Me.txt病种.Tag = lng病种ID
        End If
    End With
    
    '取病人信息
    gstrSQL = "Select 姓名 From 病人信息 Where 病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取病人信息", mlng病人ID)
    STR姓名 = rsTemp!姓名
    
    Me.lblNote(0).Caption = "病人姓名：" & STR姓名
    Me.lblNote(1).Caption = "社会保障号：" & str社会保障号
End Sub

Private Sub txt病种_GotFocus()
    Call zlControl.TxtSelAll(txt病种)
End Sub

Private Sub txt病种_KeyPress(KeyAscii As Integer)
    Dim rsTemp As New ADODB.Recordset, strText As String
    Dim blnReturn As Boolean, lst As ListItem
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If txt病种.Text = "" And txt病种.Tag <> "" Then Exit Sub
    
    On Error GoTo errHandle
    
    strText = UCase(txt病种.Text)
    If InStr(1, strText, "(") <> 0 Then
        If InStr(1, strText, ")") <> 0 Then
            strText = Mid(strText, 2, InStr(1, strText, ")") - 2)
        End If
    End If
    gstrSQL = " Select ID,病种代码 As 编码,病种名称,中医名称,病种类别,个人自付比例,个人起付金额 " & _
              " From 病种目录表 A" & _
              " Where (" & zlCommFun.GetLike("A", "病种代码", strText) & " or " & zlCommFun.GetLike("A", "病种名称", strText) & " or zlspellcode(病种名称) Like '" & strText & "%')"
    With rsTemp
        If .State = 1 Then .Close
        .Open gstrSQL, gcnGYBJYB
    End With
    
    If rsTemp.RecordCount = 0 Then
        MsgBox "不存在该病种，请重新输入！", vbInformation, gstrSysName
        txt病种.Text = lbl病种.Tag
        zlControl.TxtSelAll txt病种
        Exit Sub
    Else
        '出现选择器
        If rsTemp.RecordCount > 1 Then
            '对于字段大于3的，即使只有一条记录把该对话框显示出来，以便让用户得到更多的信息
            blnReturn = frmListSel.ShowSelect(TYPE_毕节, rsTemp, "ID", "医保病种选择", "请选择医保病种：")
        Else
            blnReturn = True
        End If
    End If
    
    If blnReturn = False Then
        '记录集中没有可选择的数据
        txt病种.Text = lbl病种.Tag
        zlControl.TxtSelAll txt病种
        Exit Sub
    Else
        '肯定是有记录集的
        txt病种.Tag = rsTemp!ID
        txt病种.Text = "(" & rsTemp!编码 & ")" & rsTemp!病种名称
        lbl病种.Tag = txt病种.Text '用于恢复显示
    End If
    
    Call zlCommFun.PressKey(vbKeyTab)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmd病种_Click()
    Dim blnReturn As Boolean
    Dim rsTemp As New ADODB.Recordset
        
    gstrSQL = " Select ID,病种代码 As 编码,病种名称,中医名称,病种类别,个人自付比例,个人起付金额 " & _
              " From 病种目录表"
    With rsTemp
        If .State = 1 Then .Close
        .Open gstrSQL, gcnGYBJYB
    End With
    
    blnReturn = frmListSel.ShowSelect(TYPE_毕节, rsTemp, "ID", "医保病种选择", "请选择医保病种：")
    If blnReturn = False Then
        '记录集中没有可选择的数据
        txt病种.Text = lbl病种.Tag
        zlControl.TxtSelAll txt病种
        Exit Sub
    Else
        '肯定是有记录集的
        txt病种.Tag = rsTemp!ID
        txt病种.Text = "(" & rsTemp!编码 & ")" & rsTemp!病种名称
        lbl病种.Tag = txt病种.Text '用于恢复显示
    End If
End Sub
