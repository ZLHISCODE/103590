VERSION 5.00
Begin VB.Form frmStoreSpaceCard 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "货位编辑"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4515
   Icon            =   "frmStoreSpaceCard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cboRoom 
      Height          =   300
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   120
      Width           =   2640
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Left            =   0
      TabIndex        =   10
      Top             =   600
      Width           =   9810
   End
   Begin VB.TextBox txt编码 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1200
      TabIndex        =   1
      Top             =   720
      Width           =   2625
   End
   Begin VB.TextBox txt备注 
      Appearance      =   0  'Flat
      Height          =   1020
      Left            =   1200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2160
      Width           =   2895
   End
   Begin VB.TextBox txt简码 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1200
      TabIndex        =   3
      Top             =   1680
      Width           =   2625
   End
   Begin VB.TextBox txt名称 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1200
      TabIndex        =   2
      Top             =   1200
      Width           =   2625
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存(&S)"
      Height          =   350
      Left            =   1680
      TabIndex        =   5
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "关闭(&C)"
      Height          =   350
      Left            =   2880
      TabIndex        =   0
      Top             =   3480
      Width           =   1100
   End
   Begin VB.Label lblComment 
      AutoSize        =   -1  'True
      Caption         =   "已增加数量：0"
      Height          =   180
      Left            =   120
      TabIndex        =   13
      Top             =   3570
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "编码"
      Height          =   180
      Left            =   480
      TabIndex        =   11
      Top             =   780
      Width           =   360
   End
   Begin VB.Label lbl备注 
      AutoSize        =   -1  'True
      Caption         =   "备注"
      Height          =   180
      Left            =   480
      TabIndex        =   9
      Top             =   2220
      Width           =   360
   End
   Begin VB.Label lbl简码 
      AutoSize        =   -1  'True
      Caption         =   "简码"
      Height          =   180
      Left            =   480
      TabIndex        =   8
      Top             =   1740
      Width           =   360
   End
   Begin VB.Label lblSpace 
      AutoSize        =   -1  'True
      Caption         =   "名称"
      Height          =   180
      Left            =   480
      TabIndex        =   7
      Top             =   1260
      Width           =   360
   End
   Begin VB.Label lblStore 
      AutoSize        =   -1  'True
      Caption         =   "部门"
      Height          =   180
      Left            =   480
      TabIndex        =   6
      Top             =   180
      Width           =   360
   End
End
Attribute VB_Name = "frmStoreSpaceCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mint编辑状态 As Integer '1-新增 2-修改
Private mlng库房ID As Long
Private mlng货位id As Long
Private mblnRefresh As Boolean
Private mintAddCount As Integer '新增数量


Private Function GetNextCode() As String
    Dim rsTemp As ADODB.Recordset
    
    '取下一个编码
    gstrSql = "Select Max(编码) as 编码 From 药品库房货位 "
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "库房货位最大编码")
    
    If NVL(rsTemp!编码) = "" Then
        GetNextCode = "00001"
    Else
        GetNextCode = zlCommFun.IncStr(rsTemp!编码)
    End If
End Function

Public Function ShowMe(ByVal int编辑状态 As Integer, ByVal lng库房ID As Long, ByVal lng货位id As Long, ByVal fraPar As Form) As Boolean
    
    mint编辑状态 = int编辑状态
    mlng库房ID = lng库房ID
    mlng货位id = lng货位id
    
    Me.Show vbModal, fraPar
    
    ShowMe = mblnRefresh
End Function

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim colData As New Collection
    Dim i As Integer
    Dim rsData As ADODB.Recordset
    
    On Error GoTo errH
    
    If Trim(txt编码.Text) = "" Then
        MsgBox "编码不能为空！", vbInformation, gstrSysName
        txt编码.SetFocus
        Exit Sub
    End If
    
    If LenB(StrConv(txt编码.Text, vbFromUnicode)) > txt编码.MaxLength Then
        MsgBox "名称长度超过" & txt编码.MaxLength & "个字符！", vbInformation, gstrSysName
        txt编码.SetFocus
        Exit Sub
    End If
    
    If Trim(txt名称.Text) = "" Then
        MsgBox "货位名称不能为空！", vbInformation, gstrSysName
        txt名称.SetFocus
        Exit Sub
    End If
    
    If LenB(StrConv(txt名称.Text, vbFromUnicode)) > txt名称.MaxLength Then
        MsgBox "名称长度超过" & txt名称.MaxLength & "个字符或者 " & Int(txt名称.MaxLength / 2) & "个汉字！", vbInformation, gstrSysName
        txt名称.SetFocus
        Exit Sub
    End If
    
    If Trim(txt简码.Text) = "" Then
        MsgBox "简码不能为空！", vbInformation, gstrSysName
        txt简码.SetFocus
        Exit Sub
    End If
    
    If LenB(StrConv(txt简码.Text, vbFromUnicode)) > txt简码.MaxLength Then
        MsgBox "简码长度超过" & txt简码.MaxLength & "个字符或者" & Int(txt简码.MaxLength / 2) & "个汉字！", vbInformation, gstrSysName
        txt简码.SetFocus
        Exit Sub
    End If
    
    If LenB(StrConv(txt备注.Text, vbFromUnicode)) > txt备注.MaxLength Then
        MsgBox "备注长度超过" & txt备注.MaxLength & "个字符或者" & Int(txt备注.MaxLength / 2) & "个汉字！", vbInformation, gstrSysName
        txt备注.SetFocus
        Exit Sub
    End If
    
    '检查编码重复
    If mint编辑状态 = 1 Then
        '新增时全表检查
        gstrSql = "Select 1 From 药品库房货位 Where 编码 = [1]"
    Else
        '修改时，排除自身
        gstrSql = "Select 1 From 药品库房货位 Where 编码 = [1] And ID <> [2] "
    End If
    Set rsData = zldatabase.OpenSQLRecord(gstrSql, "检查编码重复", txt编码.Text, mlng货位id)
    
    If Not rsData.EOF Then
        MsgBox "编码重复，请重新录入！", vbInformation, gstrSysName
        txt编码.SetFocus
        Exit Sub
    End If
    
    '检查名称重复
    If mint编辑状态 = 1 Then
        '新增时全表检查
        gstrSql = "Select 1 From 药品库房货位 Where 库房id = [2] And 名称 = [1]"
    Else
        '修改时，排除自身
        gstrSql = "Select 1 From 药品库房货位 Where 库房id = [2] And 名称 = [1] And ID <> [3] "
    End If
    Set rsData = zldatabase.OpenSQLRecord(gstrSql, "检查名称重复", txt名称.Text, Val(cboRoom.ItemData(cboRoom.ListIndex)), mlng货位id)
    
    If Not rsData.EOF Then
        MsgBox "名称重复，请重新录入！", vbInformation, gstrSysName
        txt名称.SetFocus
        Exit Sub
    End If
    
    
    If mint编辑状态 = 1 Then
        '新增
        gstrSql = "Zl_药品库房货位_Insert("
        '编码
        gstrSql = gstrSql & "'" & txt编码.Text & "'"
        '名称_In   In 药品库房货位.名称%Type,
        gstrSql = gstrSql & ",'" & txt名称.Text & "'"
        '简码_In   In 药品库房货位.简码%Type,
        gstrSql = gstrSql & ",'" & txt简码.Text & "'"
        '库房id_In In 药品库房货位.库房id%Type
        gstrSql = gstrSql & "," & Val(cboRoom.ItemData(cboRoom.ListIndex))
        '备注_In In 药品库房货位.备注%Type
        gstrSql = gstrSql & "," & IIf(txt备注.Text = "", "null", "'" & txt备注.Text & "'")
        gstrSql = gstrSql & ")"
        
        colData.Add gstrSql, "k_1"
    Else
        '修改
        gstrSql = "Zl_药品库房货位_Update("
        'ID
        gstrSql = gstrSql & mlng货位id
        '编码
        gstrSql = gstrSql & ",'" & txt编码.Text & "'"
        '名称_In   In 药品库房货位.名称%Type,
        gstrSql = gstrSql & ",'" & txt名称.Text & "'"
        '简码_In   In 药品库房货位.简码%Type,
        gstrSql = gstrSql & ",'" & txt简码.Text & "'"
        '库房id_In In 药品库房货位.库房id%Type
        gstrSql = gstrSql & "," & Val(cboRoom.ItemData(cboRoom.ListIndex))
        '备注_In In 药品库房货位.备注%Type
        gstrSql = gstrSql & "," & IIf(txt备注.Text = "", "null", "'" & txt备注.Text & "'")
        gstrSql = gstrSql & ")"
        
        colData.Add gstrSql, "k_2"
    End If
    
    gcnOracle.BeginTrans
    For i = 1 To colData.Count
        Call zldatabase.ExecuteProcedure(colData(i), Me.Caption)
    Next
    gcnOracle.CommitTrans
    
    mblnRefresh = colData.Count > 0
    
    If mint编辑状态 = 1 Then
        '清空界面数据，可以连续增加
        txt编码.Text = GetNextCode
        txt名称.Text = ""
        txt简码.Text = ""
        txt备注.Text = ""
        txt名称.SetFocus
        
        mintAddCount = mintAddCount + 1
        If lblComment.Visible = False Then lblComment.Visible = True
        lblComment.Caption = "已新增数量：" & mintAddCount
    Else
        Unload Me
    End If
    
    Exit Sub
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    txt名称.SetFocus
    
    lblComment.Visible = (mint编辑状态 = 1)
    lblComment.Caption = "已增加数量：0"
End Sub

Private Sub SetBorder(ByVal objControl As Variant, Optional ByVal blnIsFocuse As Boolean = True)
    '功能：设置文本框背景色
    
    If blnIsFocuse Then
        objControl.BackColor = &HDCDBC5
    Else
        objControl.BackColor = &H80000005
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim rsTemp As ADODB.Recordset
    
    mintAddCount = 0
    
    With cboRoom
        .Clear
        For i = 0 To frmStoreSpace.cboRoom.ListCount - 1
            .AddItem frmStoreSpace.cboRoom.List(i)
            .ItemData(.NewIndex) = frmStoreSpace.cboRoom.ItemData(i)
        Next
        
        If .ListIndex <> 0 Then .ListIndex = frmStoreSpace.cboRoom.ListIndex
        
        .Enabled = False
    End With
    
    gstrSql = "Select ID, 编码, 名称, 简码, 库房id, 备注 From 药品库房货位 where id = [1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "库房货位", mlng货位id)
                
    txt编码.MaxLength = rsTemp.Fields("编码").DefinedSize
    txt名称.MaxLength = rsTemp.Fields("名称").DefinedSize
    txt简码.MaxLength = rsTemp.Fields("简码").DefinedSize
    txt备注.MaxLength = rsTemp.Fields("备注").DefinedSize
        
    If Not rsTemp.EOF Then
        txt编码.Text = rsTemp!编码
        txt名称.Text = rsTemp!名称
        txt简码.Text = NVL(rsTemp!简码)
        txt备注.Text = NVL(rsTemp!备注)
    End If
    
    If mint编辑状态 = 1 Then
        txt编码.Text = GetNextCode
    End If
End Sub


Private Sub txt备注_GotFocus()
    zlControl.TxtSelAll txt备注
    SetBorder txt备注
End Sub

Private Sub txt备注_LostFocus()
    SetBorder txt备注, False
End Sub


Private Sub txt编码_GotFocus()
    zlControl.TxtSelAll txt编码
    SetBorder txt编码
End Sub

Private Sub txt编码_LostFocus()
    SetBorder txt编码, False
End Sub
Private Sub txt编码_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub txt编码_Validate(Cancel As Boolean)
    txt编码.Text = String(txt编码.MaxLength - Len(txt编码.Text), "0") & txt编码.Text
End Sub

Private Sub txt简码_LostFocus()
    SetBorder txt简码, False
End Sub

Private Sub txt名称_Change()
    Dim strTmp As String
    
    strTmp = MoveSpecialChar(txt名称.Text)
    If txt名称.Text <> strTmp Then
        txt名称.Text = strTmp
    End If
    Me.txt简码.Text = zlStr.GetCodeByORCL(strTmp, False, 10)
End Sub

Private Sub txt名称_GotFocus()
    zlControl.TxtSelAll txt名称
    SetBorder txt名称
End Sub

Private Sub txt名称_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If Not (KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete) Then
        If LenB(StrConv(txt名称.Text, vbFromUnicode)) >= 50 Then KeyAscii = 0
    End If
End Sub

Private Sub txt备注_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If Not (KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete) Then
        If LenB(StrConv(txt备注.Text, vbFromUnicode)) >= 100 Then KeyAscii = 0
    End If
End Sub

Private Sub txt简码_GotFocus()
    zlControl.TxtSelAll txt简码
    SetBorder txt简码
End Sub

Private Sub txt简码_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If Not (KeyAscii = vbKeyDelete Or KeyAscii = vbKeyBack) Then
        If LenB(StrConv(txt简码.Text, vbFromUnicode)) >= 10 Then KeyAscii = 0
    End If
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
End Sub

Private Sub txt名称_LostFocus()
    Me.txt简码.Text = zlStr.GetCodeByORCL(txt名称.Text, False, 10)
    SetBorder txt名称, False
End Sub


