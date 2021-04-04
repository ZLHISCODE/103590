VERSION 5.00
Begin VB.Form frmParEdit 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   Icon            =   "frmParEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtName 
      BackColor       =   &H8000000F&
      Height          =   300
      Index           =   0
      Left            =   150
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   450
      Width           =   1215
   End
   Begin VB.TextBox txtType 
      BackColor       =   &H8000000F&
      Height          =   300
      Index           =   0
      Left            =   2925
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   450
      Width           =   465
   End
   Begin VB.ComboBox cboGroup 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   0
      Left            =   5550
      TabIndex        =   2
      Top             =   450
      Width           =   1575
   End
   Begin VB.ComboBox cboValue 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   0
      Left            =   3435
      TabIndex        =   1
      Top             =   450
      Width           =   1905
   End
   Begin VB.TextBox txtAlias 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   0
      Left            =   1410
      MaxLength       =   20
      TabIndex        =   0
      Top             =   450
      Width           =   1470
   End
   Begin VB.PictureBox picCmd 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   7230
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   825
      Width           =   7230
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   5580
         TabIndex        =   4
         Top             =   195
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   4185
         TabIndex        =   3
         Top             =   195
         Width           =   1100
      End
      Begin VB.Frame Frame1 
         Height          =   75
         Left            =   -105
         TabIndex        =   8
         Top             =   30
         Width           =   7875
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "参数名"
      Height          =   180
      Left            =   487
      TabIndex        =   13
      Top             =   105
      Width           =   540
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "所属组"
      Height          =   180
      Left            =   6067
      TabIndex        =   12
      Top             =   90
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "缺省值"
      Height          =   180
      Left            =   4117
      TabIndex        =   11
      Top             =   105
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "类型"
      Height          =   180
      Left            =   2977
      TabIndex        =   10
      Top             =   120
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "别名"
      Height          =   180
      Left            =   1875
      TabIndex        =   9
      Top             =   120
      Width           =   360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      X1              =   15
      X2              =   8000
      Y1              =   360
      Y2              =   360
   End
End
Attribute VB_Name = "frmParEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mobjPars As FuncPars '入/出
Public mstrPars As String '入：参数串
Public mlngSys As Long '入：系统
Public mstrOwner As String '入：所有者
Private arrCustom() As CustomPar

Private Sub cboGroup_GotFocus(Index As Integer)
    '重新获取所有组名
    Dim strGroup As String, arrGroup() As String
    Dim i As Integer, strText As String
    
    '填充组名
    strGroup = ""
    For i = 0 To cboGroup.UBound
        If InStr(strGroup & ",", "," & cboGroup(i).Text & ",") = 0 And cboGroup(i).Text <> "" Then
            strGroup = strGroup & "," & cboGroup(i).Text
        End If
    Next
    If strGroup <> "" Then
        strGroup = Mid(strGroup, 2)
        arrGroup = Split(strGroup, ",")
        
        '为各组填充组名
        strText = cboGroup(Index).Text
        cboGroup(Index).Clear
        For i = 0 To UBound(arrGroup)
            cboGroup(Index).AddItem arrGroup(i)
        Next
        
        cboGroup(Index).Text = strText
'        If cboGroup(Index).Text = "" Then
'            If Index > 0 Then
'                If cboGroup(Index - 1).Text <> "" Then cboGroup(Index).Text = cboGroup(Index - 1).Text
'            ElseIf Index < cboGroup.UBound Then
'                If cboGroup(Index + 1).Text <> "" Then cboGroup(Index).Text = cboGroup(Index + 1).Text
'            End If
'        End If
        
        SelAll cboGroup(Index)
    End If
End Sub

Private Sub cboGroup_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr("~`@#$%^&*()=+][}{'"";/?.>,<\|", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub cboValue_Click(Index As Integer)
    Dim tmpPar As FuncPar, blnDo As Boolean, blnOK As Boolean
    
    blnOK = gblnOK
    
    If cboValue(Index).Text Like "*…" Then
        cboValue(Index).ToolTipText = "按 F2 进入" & cboValue(Index).Text
    Else
        cboValue(Index).ToolTipText = ""
    End If

    If Visible Then
        If cboValue(Index).Text = "固定值列表…" Then
            frmFixValue.mbytDataType = txtType(Index).Tag
            frmFixValue.mstrParName = IIf(txtAlias(Index).Text = "", txtName(Index).Text, txtAlias(Index).Text)
            frmFixValue.mbytSelType = arrCustom(Index).格式
            
            '可能从选择器定义切换过来
            If InStr(arrCustom(Index).值列表, "√") > 0 And InStr(arrCustom(Index).值列表, ",") > 0 Then
                frmFixValue.mstrValue = arrCustom(Index).值列表
            Else
                frmFixValue.mstrValue = ""
            End If
            On Error Resume Next
            frmFixValue.Show 1, Me
            On Error GoTo 0
            If gblnOK Then
                arrCustom(Index).值列表 = frmFixValue.mstrValue
                arrCustom(Index).格式 = frmFixValue.mbytSelType
                Unload frmFixValue
            ElseIf arrCustom(Index).值列表 = "" Then
                 cboValue(Index).Text = ""
            End If
        ElseIf cboValue(Index).Text = "选择器定义…" Then
            frmSelValue.mstrSQLList = arrCustom(Index).明细SQL
            frmSelValue.mstrSQLTree = arrCustom(Index).分类SQL
            frmSelValue.mstrFLDList = arrCustom(Index).明细字段
            frmSelValue.mstrFLDTree = arrCustom(Index).分类字段
            frmSelValue.mstrObj = arrCustom(Index).对象
            '可能从固定值切换过来
            frmSelValue.mstrDef = IIf(InStr(arrCustom(Index).值列表, "√") > 0, "", arrCustom(Index).值列表)

            frmSelValue.mbytDataType = txtType(Index).Tag
            frmSelValue.mstrParName = IIf(txtAlias(Index).Text = "", txtName(Index).Text, txtAlias(Index).Text)
            frmSelValue.mlngSys = mlngSys
            frmSelValue.mstrOwner = mstrOwner

            frmSelValue.Show 1, Me
            If gblnOK Then
                arrCustom(Index).明细SQL = frmSelValue.mstrSQLList
                arrCustom(Index).分类SQL = frmSelValue.mstrSQLTree
                arrCustom(Index).明细字段 = frmSelValue.mstrFLDList
                arrCustom(Index).分类字段 = frmSelValue.mstrFLDTree
                arrCustom(Index).对象 = frmSelValue.mstrObj
                arrCustom(Index).值列表 = frmSelValue.mstrDef
                Unload frmSelValue
            ElseIf arrCustom(Index).明细SQL = "" Then
                cboValue(Index).Text = ""
            End If
        End If
    End If
    
    gblnOK = blnOK
End Sub

Private Sub cboValue_GotFocus(Index As Integer)
    SelAll cboValue(Index)
End Sub

Private Sub cboValue_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 And cboValue(Index).Text Like "*…" Then Call cboValue_Click(Index)
End Sub

Private Sub cboValue_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr("&~`!@#$^""…" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer, j As Integer, strPar As String
    Dim tmpPar As FuncPar, curPar As FuncPar

    '检查参数合法性
    For i = 0 To txtName.UBound
        If txtAlias(i).Text = "" Then
            MsgBox "参数""" & txtName(i).Text & """没有输入参数的别名！", vbInformation, App.Title
            txtName(i).SetFocus: Exit Sub
        End If
        If TLen(txtAlias(i).Text) > 40 Then
            MsgBox "参数""" & txtName(i).Text & """别名长度不能超过40个字符！", vbInformation, App.Title
            txtName(i).SetFocus: Exit Sub
        End If

        For j = 0 To txtName.UBound
            If j <> i And UCase(txtAlias(i).Text) = UCase(txtAlias(j).Text) Then
                MsgBox "参数""" & txtName(i).Text & """别名参数""" & txtName(j).Text & """别名重复！", vbInformation, App.Title
                txtAlias(j).SetFocus: Exit Sub
            End If
        Next

        If TLen(cboValue(i).Text) > 255 Then
            MsgBox "参数""" & txtName(i).Text & """缺省值长度不能超过250个字符！", vbInformation, App.Title
            cboValue(i).SetFocus: Exit Sub
        End If
        If TLen(cboGroup(i).Text) > 30 Then
            MsgBox "参数""" & txtName(i).Text & """的组名长度不能超过30个字符！", vbInformation, App.Title
            cboGroup(i).SetFocus: Exit Sub
        End If

        If cboValue(i).Text <> "" And Not cboValue(i).Text Like "*…" Then
            If Val(txtType(i).Tag) = 1 Then
                If Not IsNumeric(cboValue(i).Text) Then
                    MsgBox "参数""" & txtName(i).Text & """缺省值类型应该为数字型！", vbInformation, App.Title
                    cboValue(i).SetFocus: Exit Sub
                End If
            ElseIf Val(txtType(i).Tag) = 2 Then
                If Not IsDate(cboValue(i).Text) And cboValue(i).ListIndex = -1 Then
                    MsgBox "参数""" & txtName(i).Text & """缺省值类型应该为日期/时间型！", vbInformation, App.Title
                    cboValue(i).SetFocus: Exit Sub
                End If
            End If
        End If

        '自定义内容检查
        If cboValue(i).Text = "固定值列表…" Then
            If arrCustom(i).值列表 = "" Then
                MsgBox "参数""" & txtName(i).Text & """还没有定义可选择的固定值列表！", vbInformation, App.Title
                cboValue(i).SetFocus: Exit Sub
            End If
            '检查类型
            Select Case Val(txtType(i).Tag)
                Case 1 '数字
                    For j = 0 To UBound(Split(arrCustom(i).值列表, "|"))
                        If Not IsNumeric(Split(Split(arrCustom(i).值列表, "|")(j), ",")(1)) Then
                            MsgBox "参数""" & txtName(i).Text & """的固定值列表中存在非数字型绑定值！", vbInformation, App.Title
                            cboValue(i).SetFocus: Exit Sub
                        End If
                    Next
                Case 2 '日期
                    For j = 0 To UBound(Split(arrCustom(i).值列表, "|"))
                        If Not IsDate(Split(Split(arrCustom(i).值列表, "|")(j), ",")(1)) Then
                            MsgBox "参数""" & txtName(i).Text & """的固定值列表中存在非日期型绑定值！", vbInformation, App.Title
                            cboValue(i).SetFocus: Exit Sub
                        End If
                    Next
            End Select
        End If

        If cboValue(i).Text = "选择器定义…" Then
            If arrCustom(i).明细SQL = "" Then
                MsgBox "参数""" & txtName(i).Text & """还没有定义选择器的内容！", vbInformation, App.Title
                cboValue(i).SetFocus: Exit Sub
            End If
            '检查类型(之所以要再判断一次因为用户可能改了类型)
            For j = 0 To UBound(Split(arrCustom(i).明细字段, "|"))
                If InStr(Split(Split(arrCustom(i).明细字段, "|")(j), ",")(2), "&B") > 0 Then
                    If Val(txtType(i).Tag) = 1 Then
                        Select Case CLng(Split(Split(arrCustom(i).明细字段, "|")(j), ",")(1))
                            Case adNumeric, adVarNumeric  '正常数字型
                            Case Else '非数字型
                                If MsgBox("参数""" & txtName(i).Text & """的选择器绑定字段 [" & Split(Split(arrCustom(i).明细字段, "|")(j), ",")(0) & "] 非数字类型,要继续吗？", vbQuestion + vbYesNo, App.Title) = vbNo Then
                                    cboValue(i).SetFocus: Exit Sub
                                End If
                        End Select
                    ElseIf Val(txtType(i).Tag) = 2 Then
                        Select Case CLng(Split(Split(arrCustom(i).明细字段, "|")(j), ",")(1))
                            Case adDBTimeStamp '正常日期型
                            Case Else '非日期型
                                If MsgBox("参数""" & txtName(i).Text & """的选择器绑定字段 [" & Split(Split(arrCustom(i).明细字段, "|")(j), ",")(0) & "] 非日期类型,要继续吗？", vbQuestion + vbYesNo, App.Title) = vbNo Then
                                    cboValue(i).SetFocus: Exit Sub
                                End If
                        End Select
                    End If
                End If
            Next
        End If
    Next

    '相同组名至少含有两个参数(且只能是相邻的参数才能属于同一组)单选框不能设置所属组
    j = 0: strPar = ""
    For i = 0 To UBound(arrCustom)
        If arrCustom(i).格式 = 1 And cboGroup(i).Text <> "" Then
            MsgBox "参数""" & txtName(i).Text & """不能属于任何参数组，因为该参数是单选框形式！", vbInformation, App.Title
            cboGroup(i).SetFocus: Exit Sub
        End If
    Next
    For i = 0 To UBound(arrCustom)
        If strPar <> cboGroup(i).Text Then
            If cboGroup(i).Text = "" Then
                If Not (j = 0 Or j > 1) Then
                    MsgBox "每个参数组至少要有两个参数（" & txtName(i).Text & "）！", vbInformation, App.Title
                    cboGroup(i).SetFocus: Exit Sub
                End If
                strPar = cboGroup(i).Text
                j = 1
            Else
                If j = 0 Or j > 1 Or strPar = "" Then
                    strPar = cboGroup(i).Text
                    j = 1
                Else
                    MsgBox "每个参数组至少要有两个参数（" & txtName(i).Text & "）！", vbInformation, App.Title
                    cboGroup(i).SetFocus: Exit Sub
                End If
            End If
        Else
            j = j + 1
        End If
    Next
    
    If Not (j = 0 Or j > 1 Or strPar = "") Then
        MsgBox "每个参数组至少要有两个参数（" & txtName(i - 1).Text & "）！", vbInformation, App.Title
        cboGroup(i - 1).SetFocus
        Exit Sub
    End If

    '确定输入内容
    Set mobjPars = New FuncPars
    For i = 0 To txtName.UBound
        '如果以前自定义了内容而现在不使用，同样存入数据库，以便以后使用。
        Set curPar = Nothing
        With arrCustom(i)
            Set curPar = mobjPars.Add(cboGroup(i).Text, CByte(i), txtName(i).Text, txtAlias(i).Text, Val(txtType(i).Tag), _
                cboValue(i).Text, .格式, .值列表, .分类SQL, .明细SQL, .分类字段, .明细字段, .对象, "_" & txtName(i).Text)
        End With
    Next

    gblnOK = True
    Hide
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: SendKeys "{Tab}"
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Dim intCount As Integer, i As Integer, j As Integer
    Dim strName As String, strType As String
    
    gblnOK = False
    
    intCount = UBound(Split(mstrPars, ";")) + 1
    ReDim arrCustom(intCount - 1) As CustomPar
    
    For i = 0 To intCount - 1
        If i <> 0 Then
            Load txtName(i): txtName(i).Left = txtName(0).Left: txtName(i).Top = txtName(0).Top + 450 * i: txtName(i).Visible = True
            Load txtAlias(i): txtAlias(i).Left = txtAlias(0).Left: txtAlias(i).Top = txtAlias(0).Top + 450 * i: txtAlias(i).TabIndex = txtAlias(0).TabIndex + 4 * i: txtAlias(i).Visible = True
            Load txtType(i): txtType(i).Left = txtType(0).Left: txtType(i).Top = txtType(0).Top + 450 * i:  txtType(i).Visible = True
            Load cboValue(i): cboValue(i).Left = cboValue(0).Left: cboValue(i).Top = cboValue(0).Top + 450 * i: cboValue(i).TabIndex = cboValue(0).TabIndex + 4 * i: cboValue(i).Visible = True
            Load cboGroup(i): cboGroup(i).Left = cboGroup(0).Left: cboGroup(i).Top = cboGroup(0).Top + 450 * i: cboGroup(i).TabIndex = cboGroup(0).TabIndex + 4 * i: cboGroup(i).Visible = True
        End If

        '固定可填写的值
        strName = Split(Split(mstrPars, ";")(i), ",")(0)
        txtName(i).Text = strName
        strType = Split(Split(mstrPars, ";")(i), ",")(1)
        If UCase(strType) Like "*NUMBER*" Then
            txtType(i).Text = "数值"
            txtType(i).Tag = 1
        ElseIf UCase(strType) Like "*CHAR*" Then
            txtType(i).Text = "字符"
            txtType(i).Tag = 0
        ElseIf UCase(strType) Like "*DATE*" Then
            txtType(i).Text = "日期"
            txtType(i).Tag = 2
        End If
        
        '缺省的值
        txtAlias(i).Text = ""
        cboValue(i).Text = ""
        cboValue(i).AddItem "固定值列表…"
        cboValue(i).AddItem "选择器定义…"
        cboGroup(i).Text = ""
                        
        '值设置
        If UCase(strName) = "ZLBEGINTIME" Or UCase(strName) = "ZLENDTIME" Then
            '动态时间参数固定设置
            txtAlias(i).Locked = True
            txtAlias(i).TabStop = False
            txtAlias(i).BackColor = txtName(i).BackColor
            cboValue(i).Locked = True
            cboValue(i).TabStop = False
            cboValue(i).BackColor = txtName(i).BackColor
            
            If UCase(strName) = "ZLBEGINTIME" Then
                txtAlias(i).Text = "开始时间"
            ElseIf UCase(strName) = "ZLENDTIME" Then
                txtAlias(i).Text = "结束时间"
            End If
        Else
            txtAlias(i).Locked = False
            txtAlias(i).TabStop = True
            txtAlias(i).BackColor = &HFFFFFF
            cboValue(i).Locked = False
            cboValue(i).TabStop = True
            cboValue(i).BackColor = &HFFFFFF
        End If
        
        '尽量保持原有的值
        For j = 1 To mobjPars.Count
            If UCase(mobjPars(j).名称) = UCase(strName) Then
                If Not txtAlias(i).Locked Then
                    arrCustom(i).格式 = mobjPars(j).格式
                    arrCustom(i).值列表 = mobjPars(j).值列表
                    arrCustom(i).分类SQL = mobjPars(j).分类SQL
                    arrCustom(i).明细SQL = mobjPars(j).明细SQL
                    arrCustom(i).分类字段 = mobjPars(j).分类字段
                    arrCustom(i).明细字段 = mobjPars(j).明细字段
                    arrCustom(i).对象 = mobjPars(j).对象
                    
                    txtAlias(i).Text = mobjPars(j).中文名
                    If mobjPars(j).缺省值 Like "*…" Then
                        cboValue(i).ListIndex = GetCboIndex(cboValue(i), mobjPars(j).缺省值)
                    Else
                        cboValue(i).Text = mobjPars(j).缺省值
                    End If
                End If
                cboGroup(i).Text = mobjPars(j).组名
            End If
        Next
    Next
    
    cmdOK.TabIndex = cboGroup(cboGroup.UBound).TabIndex + 1
    cmdCancel.TabIndex = cmdOK.TabIndex + 1
    Height = txtName(txtName.UBound).Top + 1365
End Sub

Private Sub txtAlias_GotFocus(Index As Integer)
    SelAll txtAlias(Index)
End Sub

Private Sub txtAlias_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Trim(txtAlias(Index).Text) = "" Then txtAlias(Index).Text = txtName(Index).Text
    End If
End Sub

Private Sub txtName_GotFocus(Index As Integer)
    SelAll txtName(Index)
End Sub

Private Sub txtName_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr("~`@#$%^&*()=+][}{'"";/?.>,<\|", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtType_GotFocus(Index As Integer)
    SelAll txtType(Index)
End Sub
