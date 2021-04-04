VERSION 5.00
Begin VB.Form frmBaseInfoEdit 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2445
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14985
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   14985
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fraEdit 
      Caption         =   "基础信息编辑"
      Height          =   2055
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   12165
      Begin VB.ComboBox cbo分类1 
         Height          =   300
         Left            =   11880
         TabIndex        =   5
         Text            =   "cbo分类1"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txt名称 
         Height          =   300
         Left            =   3570
         MaxLength       =   60
         TabIndex        =   1
         Top             =   360
         Width           =   2235
      End
      Begin VB.TextBox txt编码 
         Height          =   300
         Left            =   1080
         MaxLength       =   12
         TabIndex        =   0
         Top             =   360
         Width           =   1380
      End
      Begin VB.ComboBox cbo适用性别 
         Height          =   300
         Left            =   7035
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   2235
      End
      Begin VB.TextBox txt说明 
         Height          =   720
         Left            =   1080
         MaxLength       =   60
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   840
         Width           =   3285
      End
      Begin VB.TextBox txt简码 
         Height          =   300
         Left            =   7035
         MaxLength       =   10
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txt管码 
         Height          =   300
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   7
         Top             =   840
         Width           =   615
      End
      Begin VB.CheckBox chk缺省标志 
         Caption         =   "缺省标志(注意这个标志具有排他性)"
         Height          =   255
         Left            =   4980
         TabIndex        =   8
         Top             =   840
         Width           =   3255
      End
      Begin VB.ComboBox cbo分类 
         Height          =   300
         Left            =   10320
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   1395
      End
      Begin VB.Label lbl名称 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "名称"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3000
         TabIndex        =   14
         Top             =   420
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
         Left            =   480
         TabIndex        =   13
         Top             =   420
         Width           =   195
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "适用性别"
         Height          =   180
         Left            =   6375
         TabIndex        =   12
         Top             =   420
         Width           =   360
      End
      Begin VB.Label lbl说明 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "说明"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   480
         TabIndex        =   11
         Top             =   840
         Width           =   360
      End
      Begin VB.Label lbl分类 
         Caption         =   "分类"
         Height          =   180
         Left            =   9720
         TabIndex        =   10
         Top             =   420
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmBaseInfoEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstr编码OLD As String          '当前显示的项目id
Private mstr编码New As String

Dim objItem As ListItem
Dim lngCount As Long

'--------------------------------------------
'以下为窗体公共方法
'--------------------------------------------

Public Function zlRefresh(strItemName As String, ByVal str编码 As String) As Boolean
    '功能：根据项目id刷新当前显示内容
    Dim i As Integer
    Dim rsTemp As New ADODB.Recordset, rsRec As New ADODB.Recordset
    Dim strTmp As String
    
    mstr编码OLD = str编码
    
    '清除控件文本
    Me.txt编码.Text = "": Me.txt名称.Text = "": Me.txt说明.Text = ""
    Me.txt简码.Text = "": Me.txt管码.Text = "": Me.cbo适用性别.Clear
    Me.chk缺省标志.Value = 0: Me.cbo分类.Clear

    If str编码 = "" Then zlRefresh = True: Exit Function

    '获取指定项目的信息
    Err = 0: On Error GoTo ErrHand

    gstrSql = "Select * From " & strItemName & " Where 编码 = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, str编码)
    With rsTemp
        If Not .EOF Then
            Me.txt编码.Text = Nvl(!编码)
            Me.txt名称.Text = Nvl(!名称)

            Select Case Trim(strItemName)
            Case "诊疗检验标本"
                Me.txt简码.Text = Nvl(!简码)
                For i = 0 To cbo适用性别.ListCount - 1
                    If cbo适用性别.List(i) = Nvl(!适用性别) Then
                        cbo适用性别.ListIndex = i
                        Exit For
                    End If
                Next
            Case "诊疗检验类型"
                Me.txt简码.Text = Nvl(!简码)
                Me.txt管码.Text = Nvl(!管码)
                Me.chk缺省标志.Value = Val(Nvl(!缺省标志))
            Case "检验备注文字", "检验评语文字"
                Me.txt简码.Text = Nvl(!简码)
                Me.txt说明.Text = Nvl(!说明)
        
                cbo分类.Clear
                gstrSql = "select distinct 名称 from 诊疗检验类型"
                Set rsRec = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
                
                With rsRec
                    Do While Not .EOF
                        cbo分类.AddItem Nvl(!名称)
                        .MoveNext
                    Loop
                End With
                For i = 0 To cbo分类.ListCount - 1
                    If cbo分类.List(i) = !分类 Then
                        cbo分类.ListIndex = i
                        Exit For
                    End If
                Next
            Case "检验培养文字"
                Me.txt简码.Text = Nvl(!简码)
                Me.txt说明.Text = Nvl(!说明)
            Case "检验拒收理由"
                Me.txt说明.Text = Nvl(!名称)
            Case "检验标本形态"
                Me.txt说明.Text = Nvl(!说明) '
            Case "检验审核类别", "检验细菌类别", "革兰染色分类"
                Me.txt简码.Text = Nvl(!简码)
                Me.chk缺省标志.Value = Val(Nvl(!缺省标志))
            Case "检验细菌菌属", "质控检验方法", "细菌检测方法"
                Me.txt简码.Text = Nvl(!简码)
            Case "质控报告词句"
                Me.txt简码.Text = Nvl(!简码)
                'cbo适用性别.Clear
                cbo分类1.Clear
                gstrSql = "Select Distinct 分组 From 质控报告词句"
                Set rsRec = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
                
                With rsRec
                    Do While Not .EOF
                        'cbo适用性别.AddItem Nvl(!分组)
                        cbo分类1.AddItem Nvl(!分组)
                        .MoveNext
                    Loop
                End With
                For i = 0 To cbo分类1.ListCount - 1
                    If cbo分类1.List(i) = !分组 Then
                        cbo分类1.ListIndex = i
                        Exit For
                    End If
                Next
            Case "质控试剂来源"
                Me.txt简码.Text = Nvl(!简码)
                Me.txt管码.Text = Nvl(!QC编码)
            Case "检验结果描述"
                Me.txt简码.Text = Nvl(!简码)
                cbo分类1.Clear
                gstrSql = "select distinct 分类 from 检验结果描述"
                Set rsRec = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
                With rsRec
                    Do While Not .EOF
                        cbo分类1.AddItem Nvl(!分类)
                        .MoveNext
                    Loop
                End With
                    
                For i = 0 To cbo分类1.ListCount - 1
                    If cbo分类1.List(i) = !分类 Then
                        cbo分类1.ListIndex = i
                        Exit For
                    End If
                Next
               
            End Select
        End If
    End With
        
    zlRefresh = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function zlEditStart(blnAdd As Boolean, strItemName As String, str编码 As String) As Boolean
    '功能：开始项目编辑
    '参数： blnAdd-是否增加，否则为修改
    '       lngItemId-增加的参照项目，或者指定编辑的项目
    Dim i As Integer
    Dim strTmp As String
    Dim rsTemp As New ADODB.Recordset, rsLength As New ADODB.Recordset
    
    frmBaseInfoList.sbType.Enabled = False
    Err = 0: On Error GoTo ErrHand
    If blnAdd Then
        gstrSql = "Select Nvl(Max(To_Number(编码)), 0) As 编码, Nvl(Max(Length(编码)), 0) As 长度 From " & strItemName
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "获取上一个编码的长度")
        If rsTemp!长度 <> 0 And rsTemp!长度 <= Me.txt编码.MaxLength Then
            'Me.txt编码.MaxLength = rsTemp!长度
            Me.txt编码.Text = Format(Val(rsTemp!编码) + 1, String(rsTemp!长度, "0"))
        Else
            gstrSql = " select data_length as ColLength from user_tab_columns where table_name=[1] and column_Name=[2]"
            Set rsLength = zlDatabase.OpenSQLRecord(gstrSql, "编码长度", strItemName, "编码")
            Me.txt编码.MaxLength = rsLength!Collength
            Me.txt编码.Text = Format(Val(rsTemp!编码) + 1, String(rsLength!Collength, "0"))
        End If
        
        Me.txt名称.Text = "": Me.txt管码.Text = "": Me.txt简码.Text = ""
        Me.txt说明.Text = "": Me.chk缺省标志.Value = 0
    End If
    
    Select Case Trim(strItemName)
        Case "诊疗检验标本"
            strTmp = Me.cbo适用性别.Text
            Me.cbo适用性别.Clear
            Me.cbo适用性别.AddItem (""): Me.cbo适用性别.AddItem ("男"): Me.cbo适用性别.AddItem ("女")
            For i = 0 To cbo适用性别.ListCount - 1
                If cbo适用性别.List(i) = strTmp Then
                    cbo适用性别.ListIndex = i
                    Exit For
                End If
            Next
        Case "检验备注文字", "检验评语文字"
            strTmp = cbo分类.Text
            cbo分类.Clear
            gstrSql = "select distinct 名称 from 诊疗检验类型"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
            
            With rsTemp
                Do While Not .EOF
                    cbo分类.AddItem Nvl(!名称)
                    .MoveNext
                Loop
            End With
            For i = 0 To cbo分类.ListCount - 1
                If cbo分类.List(i) = strTmp Then
                    cbo分类.ListIndex = i
                    Exit For
                End If
            Next
        Case "检验结果描述"
            strTmp = cbo分类1.Text
            cbo分类1.Clear
            gstrSql = "select distinct 分类 from 检验结果描述"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
            With rsTemp
                Do While Not .EOF
                    cbo分类1.AddItem Nvl(!分类)
                    .MoveNext
                Loop
            End With
            For i = 0 To cbo分类1.ListCount - 1
                If cbo分类1.List(i) = strTmp Then
                    cbo分类1.ListIndex = i
                    Exit For
                End If
            Next
        
        Case "质控报告词句"
            strTmp = cbo分类1.Text
            cbo分类1.Clear
            gstrSql = "Select Distinct 分组 From 质控报告词句"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
            
            With rsTemp
                Do While Not .EOF
                    cbo分类1.AddItem Nvl(!分组)
                    .MoveNext
                Loop
            End With
            For i = 0 To cbo分类1.ListCount - 1
                If cbo分类1.List(i) = strTmp Then
                    cbo分类1.ListIndex = i
                    Exit For
                End If
            Next
    End Select
    
    Me.Tag = IIf(blnAdd, "增加", "修改")
    Me.Enabled = True: Me.BackColor = RGB(250, 250, 250)
    
    Me.txt编码.Enabled = True
    Me.txt编码.SetFocus
    
    zlEditStart = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub zlEditCancel()
    '功能：放弃正在进行的编辑
    Me.Tag = ""
    Me.Enabled = False: Me.BackColor = &H8000000F
    fraEdit.BackColor = &H8000000F
    frmBaseInfoList.sbType.Enabled = True
    Call Me.zlRefresh(gstrItemName, mstr编码OLD)
End Sub

Public Function zlEditSave(strItemName As String) As String
    '功能：保存正在进行的编辑,并返回正在编辑项目id,保存失败返回0
    Dim lngNewId As Long, strLists As String
    
    frmBaseInfoList.sbType.Enabled = True
    '一般特性检查
    If Trim(Me.txt编码.Text) = "" Then
        MsgBox "请输入编码！", vbInformation, gstrSysName
        Me.txt编码.SetFocus: zlEditSave = "": Exit Function
    End If
    mstr编码New = Trim(Me.txt编码.Text)
    
    '检验编码是否重复
    If zlCodeRepeat(mstr编码New, strItemName) Then
        txt编码.SetFocus: zlEditSave = "": Exit Function
    End If
    
    If strItemName = "检验拒收理由" Then
        If Trim(Me.txt说明.Text) = "" Then
            MsgBox "请输入名称！", vbInformation, gstrSysName
            Me.txt说明.SetFocus: zlEditSave = "": Exit Function
        End If
    Else
        If Trim(Me.txt名称.Text) = "" Then
            MsgBox "请输入名称！", vbInformation, gstrSysName
            Me.txt名称.SetFocus: zlEditSave = "": Exit Function
        End If
    End If

    '字符串合法性验证
    If zlCommFun.StrIsValid(Trim(txt编码.Text), , , "编码") = False Then
        txt编码.SetFocus: zlEditSave = "": Exit Function
    End If
    
    If zlCommFun.StrIsValid(Trim(IIf(strItemName = "检验拒收理由", txt说明.Text, txt名称.Text)), IIf(strItemName = "检验拒收理由", _
                txt说明.MaxLength, txt名称.MaxLength), IIf(strItemName = "检验拒收理由", txt说明.hWnd, txt名称.hWnd), "名称") = False Then
        If strItemName = "检验拒收理由" Then
            Me.txt说明.SetFocus
        Else
            Me.txt名称.SetFocus
        End If
        zlEditSave = ""
        Exit Function
    End If
    
    If (strItemName = "诊疗检验标本" Or strItemName = "诊疗检验类型" Or strItemName = "检验备注文字" Or strItemName = "检验培养文字" _
            Or strItemName = "检验评语文字" Or strItemName = "检验审核类别" Or strItemName = "检验细菌菌属" Or strItemName = "检验细菌类别" Or strItemName = "革兰染色分类" Or _
            strItemName = "质控报告词句" Or strItemName = "质控检验方法" Or strItemName = "质控试剂来源" Or strItemName = "细菌检测方法" Or _
            strItemName = "检验结果描述" Or strItemName = "细菌耐药机制") Then
            
        If zlCommFun.StrIsValid(Trim(txt简码.Text), txt简码.MaxLength, txt简码.hWnd, "简码") = False Then Me.txt简码.SetFocus: zlEditSave = "": Exit Function
    End If
    
    If (strItemName = "检验备注文字" Or strItemName = "检验标本形态" Or strItemName = "检验培养文字" Or strItemName = "检验评语文字") Then
        If zlCommFun.StrIsValid(Trim(txt说明.Text), txt说明.MaxLength, Me.txt说明.hWnd, "说明") = False Then Me.txt说明.SetFocus: zlEditSave = "": Exit Function
    End If
    
    If strItemName = "诊疗检验类型" Then
        If zlCommFun.StrIsValid(Trim(txt管码.Text), txt管码.MaxLength, txt管码.hWnd, "管码") = False Then Me.txt管码.SetFocus: zlEditSave = "": Exit Function
    End If

    If strItemName = "质控报告词句" Then
        If zlCommFun.StrIsValid(Trim(cbo分类1.Text), 4, cbo分类1.hWnd, "分组") = False Then Me.cbo分类1.SetFocus: zlEditSave = "": Exit Function
    End If
    
    If strItemName = "质控试剂来源" Then
        If zlCommFun.StrIsValid(Trim(txt管码.Text), txt管码.MaxLength, txt管码.hWnd, "QC编码") = False Then Me.txt管码.SetFocus: zlEditSave = "": Exit Function
    End If
    '数据保存语句组织
    Select Case Trim(strItemName)
        Case "诊疗检验标本"
            gstrSql = "'" & mstr编码New & "','" & mstr编码OLD & "','" & Trim(txt名称.Text) & "','" & Trim(txt简码.Text) & "','" & Trim(cbo适用性别.Text) & "'"
        Case "诊疗检验类型"
            
            If zlCommFun.StrIsValid(Trim(txt简码.Text), txt简码.MaxLength, txt简码.hWnd, "简码") = False Then
                Me.txt简码.SetFocus: zlEditSave = "": Exit Function
            End If
            gstrSql = "'" & mstr编码New & "','" & mstr编码OLD & "','" & Trim(txt名称.Text) & "','" & Trim(txt简码.Text) & "','" & _
                            chk缺省标志.Value & "','" & Trim(txt管码.Text) & "'"
        Case "检验备注文字", "检验评语文字"
            gstrSql = "'" & mstr编码New & "','" & mstr编码OLD & "','" & Trim(txt名称.Text) & "','" & Trim(txt简码.Text) & "','" & _
                            Trim(txt说明.Text) & "','" & Trim(cbo分类.Text) & "'"
        Case "检验培养文字"
            gstrSql = "'" & mstr编码New & "','" & mstr编码OLD & "','" & Trim(txt名称.Text) & "','" & Trim(txt简码.Text) & "','" & Trim(txt说明.Text) & "'"
        Case "检验标本形态"
            gstrSql = "'" & mstr编码New & "','" & mstr编码OLD & "','" & Trim(txt名称.Text) & "','" & Trim(txt说明.Text) & "'"
        Case "检验拒收理由"
            gstrSql = "'" & mstr编码New & "','" & mstr编码OLD & "','" & Trim(Me.txt说明.Text) & "'"
        Case "检验分析用途"
            gstrSql = "'" & mstr编码New & "','" & mstr编码OLD & "','" & Trim(txt名称.Text) & "'"
        Case "检验审核类别", "检验细菌类别", "革兰染色分类"
            gstrSql = "'" & mstr编码New & "','" & mstr编码OLD & "','" & Trim(txt名称.Text) & " ','" & Trim(txt简码.Text) & "','" & _
            chk缺省标志.Value & "'"
        Case "检验细菌菌属", "质控检验方法", "细菌检测方法", "细菌耐药机制"
            gstrSql = "'" & mstr编码New & "','" & mstr编码OLD & "','" & Trim(txt名称.Text) & "','" & Trim(txt简码.Text) & "'"
        Case "质控报告词句"
            If Len(Trim(cbo分类1.Text)) > 4 Then MsgBox "请确保分组名称长度不超过4位！", vbInformation, gstrSysName: cbo分类1.SetFocus: zlEditSave = "": Exit Function
            gstrSql = "'" & mstr编码New & "','" & mstr编码OLD & "','" & Trim(txt名称.Text) & "','" & Trim(txt简码.Text) & "','" & _
                    Trim(cbo分类1.Text) & "'"
        Case "质控试剂来源"
            gstrSql = "'" & mstr编码New & "','" & mstr编码OLD & "','" & Trim(txt名称.Text) & "','" & Trim(txt简码.Text) & "','" & _
                    Trim(txt管码.Text) & "'"
        Case "检验结果描述"
            gstrSql = "'" & mstr编码New & "','" & mstr编码OLD & "','" & Trim(txt名称.Text) & "','" & Trim(txt简码.Text) & "','" & _
                    Trim(cbo分类1.Text) & "'"
    End Select
    
    Err = 0: On Error GoTo ErrHand

    If Me.Tag = "增加" Then
        If zlDatabase.OpenSQLRecord("select 名称 from " & strItemName & " where 名称 ='" & Trim(txt名称.Text) & "'", Me.Caption).RecordCount > 0 Then
            MsgBox strItemName & "名称出现重复！", vbInformation, gstrSysName
            txt名称.SetFocus: zlEditSave = "": Exit Function
        End If
        gstrSql = "zl_" & strItemName & "_Edit(1," & gstrSql & ")"
    Else
        gstrSql = "zl_" & strItemName & "_Edit(2," & gstrSql & ")"
    End If
    Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    
    Me.Tag = ""
    Me.Enabled = False: Me.BackColor = &H8000000F
    fraEdit.BackColor = &H8000000F
    
    zlEditSave = mstr编码New
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function zlCodeRepeat(strInputCode As String, strItemName As String) As Boolean
    '----------------------------------
    '功能：检查编码的是否与现有编码重复，重复则给出提示
    '入参：strInputCode-输入的编码
    '出参：重复返回True；否则反馈Flase
    '----------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    Err = 0: On Error GoTo ErrHand
    'strSQL = "select 编码,名称 from (select 编码,名称 from " & strItemName & " where 编码<>[1]) where 编码=[1]"
    strSql = "select 编码,名称 from " & strItemName & " where 编码=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "判断编码是否重复", strInputCode)
        
    With rsTmp
        If .RecordCount <> 0 And mstr编码OLD <> mstr编码New Then
            MsgBox "该项目与【" & Nvl(!编码) & "-" & Nvl(!名称) & "】编码重复！", vbExclamation, gstrSysName
            zlCodeRepeat = True
        Else
            zlCodeRepeat = False
        End If
    End With
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlCodeRepeat = True
End Function

Private Sub cbo分类1_KeyPress(KeyAscii As Integer)
    If Len(Trim(cbo分类1.Text)) = 4 And KeyAscii <> 8 Then KeyAscii = 0: Exit Sub
End Sub

'
Private Sub cbo适用性别_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo分类_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk缺省标志_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_Load()
    mstr编码OLD = ""
End Sub

Public Sub txt编码_GotFocus()
    zlControl.TxtSelAll txt编码
    'Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt编码_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
            KeyAscii = 0
            Exit Sub
        End If
    End Select
End Sub

Private Sub txt管码_GotFocus()
    zlControl.TxtSelAll txt管码
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt管码_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
            KeyAscii = 0
            Exit Sub
        End If
    End Select
End Sub

Private Sub txt简码_GotFocus()
    Me.txt名称.SelStart = 0: Me.txt名称.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt简码_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt名称_Change()
    txt简码.Text = zlCommFun.SpellCode(txt名称.Text)
End Sub

'
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


