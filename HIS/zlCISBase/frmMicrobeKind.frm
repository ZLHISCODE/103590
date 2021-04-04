VERSION 5.00
Begin VB.Form frmMicrobeKind 
   BorderStyle     =   0  'None
   Caption         =   "细菌类型"
   ClientHeight    =   1755
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5055
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1755
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txt英文 
      Height          =   300
      Left            =   975
      MaxLength       =   60
      TabIndex        =   6
      Top             =   915
      Width           =   3870
   End
   Begin VB.TextBox txt中文 
      Height          =   300
      Left            =   975
      MaxLength       =   60
      TabIndex        =   2
      Top             =   510
      Width           =   3870
   End
   Begin VB.TextBox txt编码 
      Height          =   300
      Left            =   975
      MaxLength       =   13
      TabIndex        =   1
      Top             =   120
      Width           =   1185
   End
   Begin VB.TextBox txt缩写 
      Height          =   300
      Left            =   975
      MaxLength       =   10
      TabIndex        =   0
      Top             =   1305
      Width           =   1185
   End
   Begin VB.Label lbl英文 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "英文名称"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   180
      TabIndex        =   7
      Top             =   975
      Width           =   720
   End
   Begin VB.Label lbl中文 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "中文名称"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   180
      TabIndex        =   5
      Top             =   585
      Width           =   720
   End
   Begin VB.Label lbl编码 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "类型编码"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   180
      TabIndex        =   4
      Top             =   195
      Width           =   720
   End
   Begin VB.Label lbl缩写 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "英文缩写"
      Height          =   180
      Left            =   180
      TabIndex        =   3
      Top             =   1365
      Width           =   720
   End
End
Attribute VB_Name = "frmMicrobeKind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngKindId As Long          '当前显示的类型id

Dim lngCount As Long

'--------------------------------------------
'以下为窗体公共方法
'--------------------------------------------
Public Function zlRefresh(lngKindId As Long) As Boolean
    '功能：根据项目id刷新当前显示内容
    Dim rsTemp As New ADODB.Recordset
    mlngKindId = lngKindId
    
    '清除此前项目的显示
    Me.txt编码.Text = "": Me.txt中文.Text = "": Me.txt英文.Text = "": Me.txt缩写.Text = ""
    If lngKindId = 0 Then zlRefresh = True: Exit Function
    
    '获取指定项目的信息
    Err = 0: On Error GoTo ErrHand
    gstrSql = "Select 编码, 中文名称, 英文名称, 简码 From 检验细菌类型 Where ID = [1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngKindId)
    With rsTemp
        Me.txt编码.MaxLength = .Fields("编码").DefinedSize
        Me.txt中文.MaxLength = .Fields("中文名称").DefinedSize
        Me.txt英文.MaxLength = .Fields("英文名称").DefinedSize
        Me.txt缩写.MaxLength = .Fields("简码").DefinedSize
        If .RecordCount > 0 Then
            Me.txt编码.Text = "" & !编码
            Me.txt中文.Text = "" & !中文名称
            Me.txt英文.Text = "" & !英文名称
            Me.txt缩写.Text = "" & !简码
        End If
    End With
    zlRefresh = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefresh = False: Exit Function
End Function

Public Function zlEditStart(blnAdd As Boolean, lngKindId As Long) As Boolean
    '功能：开始项目编辑
    '参数： blnAdd-是否增加，否则为修改
    '       lngKindId-增加的参照项目，或者指定编辑的项目
    Dim rsTemp As New ADODB.Recordset
        
    
    If blnAdd Then
        Err = 0: On Error GoTo ErrHand
        gstrSql = "Select Nvl(Max(编码), 0) As 编码, Nvl(Max(Length(编码)), 0) As 长度 From 检验细菌类型"
        
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "zlEditStart")
        With rsTemp
'            If .State = adStateOpen Then .Close
'            Call SQLTest(App.ProductName, Me.Caption, gstrSql)
            
'            Call SQLTest
            If !长度 <> 0 And !长度 <= Me.txt编码.MaxLength Then
                Me.txt编码.Text = Format(Val(!编码) + 1, String(!长度, "0"))
            Else
                Me.txt编码.Text = Format(Val(!编码) + 1, String(Me.txt编码.MaxLength, "0"))
            End If
        End With
        
        '清除并设置默认值
        Me.txt中文.Text = "": Me.txt英文.Text = "": Me.txt缩写.Text = ""
    End If

    mlngKindId = lngKindId
    Me.Enabled = True: Me.Tag = IIf(blnAdd, "增加", "修改")
    Me.BackColor = RGB(250, 250, 250)
    Me.txt编码.SetFocus
    zlEditStart = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditStart = False: Exit Function
End Function

Public Sub zlEditCancel()
    '功能：放弃正在进行的编辑
    Me.Enabled = False: Me.Tag = ""
    Me.BackColor = &H8000000F
    Call Me.zlRefresh(mlngKindId)
End Sub

Public Function zlEditSave() As Long
    '功能：保存正在进行的编辑,并返回正在编辑项目id,保存失败返回0
    Dim lngNewId As Long
    
    '一般特性检查
    If Trim(Me.txt编码.Text) = "" Then
        MsgBox "请输入编码！", vbInformation, gstrSysName
        Me.txt编码.SetFocus: zlEditSave = 0: Exit Function
    End If
    If Val(Me.txt编码.Text) > Val(String(Me.txt编码.MaxLength, "9")) Then
        MsgBox "编码太大！", vbInformation, gstrSysName
        Me.txt编码.SetFocus: zlEditSave = 0: Exit Function
    End If
    If Trim(Me.txt中文.Text) = "" Then
        MsgBox "请输入中文名称！", vbInformation, gstrSysName
        Me.txt中文.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt中文.Text), vbFromUnicode)) > Me.txt中文.MaxLength Then
        MsgBox "中文名称超长（最多" & Me.txt中文.MaxLength & "个字符或等长汉字）！", vbInformation, gstrSysName
        Me.txt中文.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt英文.Text), vbFromUnicode)) > Me.txt英文.MaxLength Then
        MsgBox "英文名称超长（最多" & Me.txt英文.MaxLength & "个字符）！", vbInformation, gstrSysName
        Me.txt英文.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt缩写.Text), vbFromUnicode)) > Me.txt缩写.MaxLength Then
        MsgBox "缩写超长（最多" & Me.txt缩写.MaxLength & "个字符）！", vbInformation, gstrSysName
        Me.txt缩写.SetFocus: zlEditSave = 0: Exit Function
    End If
    
    '数据保存语句组织
    If Me.Tag = "增加" Then
        lngNewId = zldatabase.GetNextId("检验细菌类型")
    Else
        lngNewId = mlngKindId
    End If

    gstrSql = "'" & Trim(Me.txt编码.Text) & "','" & Trim(Me.txt中文.Text) & "','" & Trim(Me.txt英文.Text) & "','" & Trim(Me.txt缩写.Text) & "'"
    
    If Me.Tag = "增加" Then
        gstrSql = "Zl_检验细菌类型_Insert(" & lngNewId & "," & gstrSql & ")"
    Else
        gstrSql = "Zl_检验细菌类型_Update(" & lngNewId & "," & gstrSql & ")"
    End If
    
    Err = 0: On Error GoTo ErrHand
    Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
    
    If Me.Tag = "增加" Then mlngKindId = lngNewId
    Me.Enabled = False: Me.Tag = ""
    Me.BackColor = &H8000000F
    zlEditSave = mlngKindId: Exit Function
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditSave = 0: Exit Function
End Function

'--------------------------------------------
'以下为窗体控件响应事件
'--------------------------------------------
 
Private Sub Form_Load()
    mlngKindId = 0
End Sub

Private Sub txt编码_GotFocus()
    Me.txt编码.SelStart = 0: Me.txt编码.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt编码_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt缩写_GotFocus()
    Me.txt缩写.SelStart = 0: Me.txt缩写.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt缩写_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(Trim(GCST_INVALIDCHAR), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt英文_GotFocus()
    Me.txt英文.SelStart = 0: Me.txt英文.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt英文_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(Trim(GCST_INVALIDCHAR), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt中文_GotFocus()
    Me.txt中文.SelStart = 0: Me.txt中文.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt中文_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub
