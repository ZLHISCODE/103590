VERSION 5.00
Begin VB.Form frmMassRuleEdit 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5190
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txt参数 
      Height          =   300
      Index           =   5
      Left            =   4305
      MaxLength       =   5
      TabIndex        =   20
      Top             =   1590
      Width           =   690
   End
   Begin VB.TextBox txt参数 
      Height          =   300
      Index           =   4
      Left            =   4305
      MaxLength       =   5
      TabIndex        =   18
      Top             =   1410
      Width           =   690
   End
   Begin VB.TextBox txt参数 
      Height          =   300
      Index           =   3
      Left            =   4305
      MaxLength       =   5
      TabIndex        =   16
      Top             =   1245
      Width           =   690
   End
   Begin VB.TextBox txt种类 
      Enabled         =   0   'False
      Height          =   300
      Left            =   630
      MaxLength       =   60
      TabIndex        =   1
      Text            =   "常用控制规则"
      Top             =   165
      Width           =   2235
   End
   Begin VB.Frame fraLine 
      Height          =   15
      Left            =   -30
      TabIndex        =   24
      Top             =   2850
      Width           =   5385
   End
   Begin VB.TextBox txt说明 
      Height          =   720
      Left            =   630
      MaxLength       =   60
      MultiLine       =   -1  'True
      TabIndex        =   22
      Top             =   1935
      Width           =   4365
   End
   Begin VB.TextBox txt参数 
      Height          =   300
      Index           =   2
      Left            =   4305
      MaxLength       =   5
      TabIndex        =   14
      Top             =   1095
      Width           =   690
   End
   Begin VB.TextBox txt参数 
      Height          =   300
      Index           =   1
      Left            =   4305
      MaxLength       =   5
      TabIndex        =   12
      Top             =   930
      Width           =   690
   End
   Begin VB.TextBox txt参数 
      Height          =   300
      Index           =   0
      Left            =   4305
      MaxLength       =   5
      TabIndex        =   10
      Top             =   607
      Width           =   690
   End
   Begin VB.ComboBox cbo形式 
      Height          =   300
      Left            =   630
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1491
      Width           =   2235
   End
   Begin VB.TextBox txt编码 
      Height          =   300
      Left            =   630
      MaxLength       =   13
      TabIndex        =   3
      Top             =   607
      Width           =   780
   End
   Begin VB.TextBox txt名称 
      Height          =   300
      Left            =   630
      MaxLength       =   60
      TabIndex        =   5
      Top             =   1049
      Width           =   2235
   End
   Begin VB.Label lbl参数 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "h倍标准差"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   5
      Left            =   3420
      TabIndex        =   19
      Top             =   1650
      Width           =   810
   End
   Begin VB.Label lbl参数 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "k倍标准差"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   4
      Left            =   3420
      TabIndex        =   17
      Top             =   1470
      Width           =   810
   End
   Begin VB.Label lbl参数 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "P假失控率"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   3
      Left            =   3420
      TabIndex        =   15
      Top             =   1305
      Width           =   810
   End
   Begin VB.Label lbl变量 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "规则相关参量："
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3405
      TabIndex        =   8
      Top             =   285
      Width           =   1260
   End
   Begin VB.Label lbl种类 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "种类"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   165
      TabIndex        =   0
      Top             =   225
      Width           =   360
   End
   Begin VB.Label lbl参数 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "M个测定值"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   2
      Left            =   3420
      TabIndex        =   13
      Top             =   1155
      Width           =   810
   End
   Begin VB.Label lbl参数 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "X倍标准差"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   1
      Left            =   3420
      TabIndex        =   11
      Top             =   990
      Width           =   810
   End
   Begin VB.Label lbl说明 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "说明"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   180
      TabIndex        =   21
      Top             =   1980
      Width           =   360
   End
   Begin VB.Label lbl参数 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "N个测定值"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   0
      Left            =   3420
      TabIndex        =   9
      Top             =   660
      Width           =   810
   End
   Begin VB.Image imgNote 
      Height          =   240
      Left            =   195
      Picture         =   "frmMassRuleEdit.frx":0000
      Top             =   2925
      Width           =   240
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMassRuleEdit.frx":058A
      ForeColor       =   &H00008000&
      Height          =   4680
      Left            =   465
      TabIndex        =   23
      Top             =   2985
      Width           =   4500
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl形式 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "形式"
      Height          =   180
      Left            =   180
      TabIndex        =   6
      Top             =   1551
      Width           =   360
   End
   Begin VB.Label lbl编码 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "编码"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   180
      TabIndex        =   2
      Top             =   667
      Width           =   360
   End
   Begin VB.Label lbl名称 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "名称"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   180
      TabIndex        =   4
      Top             =   1109
      Width           =   360
   End
End
Attribute VB_Name = "frmMassRuleEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mPar       '参数枚举
    n = 0: x: m: p: k: h
End Enum

Private mlngItemID As Long          '当前显示的项目id

Dim objItem As ListItem
Dim lngCount As Long

'--------------------------------------------
'以下为窗体公共方法
'--------------------------------------------

Public Function zlRefresh(lngItemID As Long) As Boolean
    '功能：根据项目id刷新当前显示内容
    Dim rsTemp As New ADODB.Recordset
    
    mlngItemID = lngItemID
    
    '清除此前项目的显示
    
    Me.txt编码.Text = "": Me.txt名称.Text = "": Me.txt说明.Text = ""
    Me.txt种类.Text = "": Me.cbo形式.Clear
    For lngCount = 0 To Me.txt参数.UBound
        Me.lbl参数(lngCount).Visible = False
        Me.txt参数(lngCount).Visible = False: Me.txt参数(lngCount).Text = ""
    Next
    If lngItemID = 0 Then zlRefresh = True: Exit Function


    '获取指定项目的信息
    Err = 0: On Error GoTo ErrHand

    gstrSql = "Select 种类, 编码, 名称, 说明, 形式, N, X, M, P, K, H From 检验质控规则 Where Id = [1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemID)
    With rsTemp
        Me.txt编码.MaxLength = .Fields("编码").DefinedSize
        Me.txt名称.MaxLength = .Fields("名称").DefinedSize
        Me.txt说明.MaxLength = .Fields("说明").DefinedSize
        If .RecordCount > 0 Then
            Me.txt种类.Tag = Val("" & !种类)
            Select Case Val("" & !种类)
            Case 1
                Me.txt种类.Text = "常用控制规则"
                Me.cbo形式.AddItem "N-Xs": Me.cbo形式.AddItem "R-Xs": Me.cbo形式.AddItem "N-T": Me.cbo形式.AddItem "N-X": Me.cbo形式.AddItem "(M of N)Xs"
            Case 2
                Me.txt种类.Text = "计算控制界限规则"
                Me.cbo形式.AddItem "N-P": Me.cbo形式.AddItem "X-P": Me.cbo形式.AddItem "R-P"
            Case 3
                Me.txt种类.Text = "累积和规则"
                Me.cbo形式.AddItem "CS(k:h)"
            End Select
            Me.txt编码.Text = "" & !编码
            Me.txt名称.Text = "" & !名称
            Me.txt说明.Text = "" & !说明
            Me.cbo形式.ListIndex = Val("" & !形式)
            Me.txt参数(mPar.n).Text = Val("" & !n)
            Me.txt参数(mPar.x).Text = Replace(Replace(" 0" & !x, " 0.", "0."), " 0", "")
            Me.txt参数(mPar.m).Text = Val("" & !m)
            Me.txt参数(mPar.p).Text = Replace(Replace(" 0" & !p, " 0.", "0."), " 0", "")
            Me.txt参数(mPar.k).Text = Replace(Replace(" 0" & !k, " 0.", "0."), " 0", "")
            Me.txt参数(mPar.h).Text = Replace(Replace(" 0" & !h, " 0.", "0."), " 0", "")
        End If
    End With
        
    zlRefresh = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefresh = False: Exit Function
End Function

Public Function zlEditStart(blnAdd As Boolean, lngItemID As Long) As Boolean
    '功能：开始项目编辑
    '参数： blnAdd-是否增加，否则为修改
    '       lngItemId-增加的参照项目，或者指定编辑的项目
    Dim rsTemp As New ADODB.Recordset
    
    If blnAdd Then
        Err = 0: On Error GoTo ErrHand
        gstrSql = "Select Nvl(Max(To_Number(编码)), 0) As 编码, Nvl(Max(Length(编码)), 0) As 长度 From 检验质控规则 Where 种类 = 1"
        
'            If .State = adStateOpen Then .Close
'            Call SQLTest(App.ProductName, Me.Caption, gstrSql)
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "zlEditStart")
'            Call SQLTest
        With rsTemp
            If !长度 <> 0 And !长度 <= Me.txt编码.MaxLength Then
                Me.txt编码.Text = Format(Val(!编码) + 1, String(!长度, "0"))
            Else
                Me.txt编码.Text = Format(Val(!编码) + 1, String(Me.txt编码.MaxLength, "0"))
            End If
        End With
        
        Me.txt名称.Text = "": Me.txt说明.Text = ""
        For lngCount = 0 To Me.txt参数.UBound
            Me.txt参数(lngCount).Text = ""
        Next
        If Val(Me.txt种类.Tag) <> 1 Then
            Me.txt种类.Tag = 1
            Me.txt种类.Text = "常用控制规则"
            With Me.cbo形式
                .Clear
                .AddItem "N-Xs": .AddItem "R-Xs": .AddItem "N-T": .AddItem "N-X": .AddItem "(M of N)Xs"
                .ListIndex = 0
            End With
        End If
    End If

    Me.Tag = IIf(blnAdd, "增加", "修改")
    Me.Enabled = True: Me.BackColor = RGB(250, 250, 250)
    
    Me.txt编码.SetFocus
    zlEditStart = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditStart = False: Exit Function
End Function

Public Sub zlEditCancel()
    '功能：放弃正在进行的编辑
    Me.Tag = ""
    Me.Enabled = False: Me.BackColor = &H8000000F
    Call Me.zlRefresh(mlngItemID)
End Sub

Public Function zlEditSave() As Long
    '功能：保存正在进行的编辑,并返回正在编辑项目id,保存失败返回0
    Dim lngNewId As Long, strLists As String
    
    '一般特性检查
    If Me.cbo形式.ListIndex = -1 Then
        MsgBox "规则形式未设置！", vbInformation, gstrSysName
        Me.cbo形式.SetFocus: zlEditSave = 0: Exit Function
    End If
    If Trim(Me.txt编码.Text) = "" Then
        MsgBox "请输入编码！", vbInformation, gstrSysName
        Me.txt编码.SetFocus: zlEditSave = 0: Exit Function
    End If
    If Val(Me.txt编码.Text) > Val(String(Me.txt编码.MaxLength, "9")) Then
        MsgBox "编码太大！", vbInformation, gstrSysName
        Me.txt编码.SetFocus: zlEditSave = 0: Exit Function
    End If
    If Trim(Me.txt名称.Text) = "" Then
        MsgBox "请输入名称！", vbInformation, gstrSysName
        Me.txt名称.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt名称.Text), vbFromUnicode)) > Me.txt名称.MaxLength Then
        MsgBox "名称超长（最多" & Me.txt名称.MaxLength & "个字符或等长汉字）！", vbInformation, gstrSysName
        Me.txt名称.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt说明.Text), vbFromUnicode)) > Me.txt说明.MaxLength Then
        MsgBox "说明超长（最多" & Me.txt说明.MaxLength & "个字符或等长汉字）！", vbInformation, gstrSysName
        Me.txt说明.SetFocus: zlEditSave = 0: Exit Function
    End If
    For lngCount = 0 To Me.txt参数.UBound
        If Me.txt参数(lngCount).Visible Then
            If Val(Trim(Me.txt参数(lngCount).Text)) = 0 Then
                MsgBox "规则约定" & Me.lbl参数(lngCount).Caption & "参数必须指定！", vbInformation, gstrSysName
                Me.txt参数(lngCount).SetFocus: zlEditSave = 0: Exit Function
            End If
            Me.txt参数(lngCount).Text = Val(Trim(Me.txt参数(lngCount).Text))
        Else
            Me.txt参数(lngCount).Text = 0
        End If
    Next
    If Me.txt参数(mPar.x).Visible Then
        If Val(Val(Me.txt参数(mPar.x).Text) * 10) <> Int(Val(Val(Me.txt参数(mPar.x).Text) * 10)) Then
            MsgBox "X参数精度太高！", vbInformation, gstrSysName
            Me.txt参数(mPar.x).SetFocus: zlEditSave = 0: Exit Function
        End If
    End If
    If Me.txt参数(mPar.n).Visible And Me.txt参数(mPar.m).Visible Then
        If Val(Me.txt参数(mPar.n).Text) <= Val(Me.txt参数(mPar.m).Text) Then
            MsgBox "N参数必须大于M参数！", vbInformation, gstrSysName
            Me.txt参数(mPar.n).SetFocus: zlEditSave = 0: Exit Function
        End If
    End If
    
    '数据保存语句组织
    gstrSql = "'" & Trim(Me.txt编码.Text) & "','" & Trim(Me.txt名称.Text) & "'," & Me.cbo形式.ListIndex
    gstrSql = gstrSql & "," & IIf(Me.cbo形式.ListIndex = 1, 2, Val(Me.txt参数(mPar.n).Text)) & "," & Val(Me.txt参数(mPar.x).Text) & "," & Val(Me.txt参数(mPar.m).Text)
    gstrSql = gstrSql & ",'" & Trim(Me.txt说明.Text) & "'"
    lngNewId = mlngItemID
    If Me.Tag = "增加" Then
        lngNewId = zldatabase.GetNextId("检验质控规则")
        gstrSql = "Zl_检验质控规则_Edit(1," & lngNewId & "," & gstrSql & ")"
    Else
        gstrSql = "Zl_检验质控规则_Edit(2," & lngNewId & "," & gstrSql & ")"
    End If
    
    Err = 0: On Error GoTo ErrHand
    Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
    
    If Me.Tag = "增加" Then mlngItemID = lngNewId
    
    Me.Tag = ""
    Me.Enabled = False: Me.BackColor = &H8000000F
    
    zlEditSave = mlngItemID: Exit Function
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditSave = 0: Exit Function
End Function

'--------------------------------------------
'以下为窗体控件响应事件
'--------------------------------------------

Private Sub cbo形式_Click()
    Dim intCount As Integer, intVisible As Integer
    
    For intCount = Me.txt参数.LBound To Me.txt参数.UBound
        Me.lbl参数(intCount).Visible = False
        Me.txt参数(intCount).Visible = False
    Next
    
    Select Case Val(Me.txt种类.Tag)
    Case 1
        Select Case Me.cbo形式.ListIndex
        Case 0 '"N-Xs"
            Me.lbl参数(mPar.n).Visible = True: Me.txt参数(mPar.n).Visible = True
            Me.lbl参数(mPar.x).Visible = True: Me.txt参数(mPar.x).Visible = True
        Case 1 '"R-Xs"
            Me.lbl参数(mPar.x).Visible = True: Me.txt参数(mPar.x).Visible = True
        Case 2 '"N-T"
            Me.lbl参数(mPar.n).Visible = True: Me.txt参数(mPar.n).Visible = True
        Case 3 '"N-X"
            Me.lbl参数(mPar.n).Visible = True: Me.txt参数(mPar.n).Visible = True
        Case 4 '"(M of N)Xs"
            Me.lbl参数(mPar.n).Visible = True: Me.txt参数(mPar.n).Visible = True
            Me.lbl参数(mPar.x).Visible = True: Me.txt参数(mPar.x).Visible = True
            Me.lbl参数(mPar.m).Visible = True: Me.txt参数(mPar.m).Visible = True
        End Select
    Case 2
        Select Case Me.cbo形式.ListIndex
        Case 0 '"N-P"
            Me.lbl参数(mPar.n).Visible = True: Me.txt参数(mPar.n).Visible = True
            Me.lbl参数(mPar.p).Visible = True: Me.txt参数(mPar.p).Visible = True
        Case 1 '"X-P"
            Me.lbl参数(mPar.p).Visible = True: Me.txt参数(mPar.p).Visible = True
        Case 2 '"R-P"
            Me.lbl参数(mPar.p).Visible = True: Me.txt参数(mPar.p).Visible = True
        End Select
    Case 3
        Me.lbl参数(mPar.k).Visible = True: Me.txt参数(mPar.k).Visible = True
        Me.lbl参数(mPar.h).Visible = True: Me.txt参数(mPar.h).Visible = True
    End Select
    
    intVisible = 0
    For intCount = Me.txt参数.LBound To Me.txt参数.UBound
        If Me.txt参数(intCount).Visible Then
            Me.lbl参数(intCount).Top = Me.lbl编码.Top + (Me.lbl名称.Top - Me.lbl编码.Top) * intVisible
            Me.txt参数(intCount).Top = Me.txt编码.Top + (Me.txt名称.Top - Me.txt编码.Top) * intVisible
            intVisible = intVisible + 1
        End If
    Next
End Sub

Private Sub cbo形式_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_Load()
    mlngItemID = 0
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

Private Sub txt参数_GotFocus(Index As Integer)
    Me.txt参数(Index).SelStart = 0: Me.txt参数(Index).SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt参数_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        Select Case Index
        Case mPar.x, mPar.p, mPar.k, mPar.h
            If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then Exit Sub
        Case Else
            If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
        End Select
    End Select
    KeyAscii = 0
End Sub

Private Sub txt名称_GotFocus()
    Me.txt名称.SelStart = 0: Me.txt名称.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt名称_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt说明_GotFocus()
    Me.txt说明.SelStart = 0: Me.txt说明.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt说明_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub
