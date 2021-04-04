VERSION 5.00
Begin VB.Form frmParEdit 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8640
   Icon            =   "frmParEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   8640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chkLock 
      Height          =   225
      Index           =   0
      Left            =   7920
      TabIndex        =   5
      Top             =   480
      Width           =   210
   End
   Begin VB.ComboBox cboGroup 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   0
      Left            =   5550
      TabIndex        =   4
      Top             =   450
      Width           =   1905
   End
   Begin VB.ComboBox cboValue 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   0
      Left            =   3540
      TabIndex        =   3
      Top             =   450
      Width           =   1905
   End
   Begin VB.ComboBox cboType 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   0
      Left            =   2385
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   450
      Width           =   1005
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   0
      Left            =   825
      MaxLength       =   20
      TabIndex        =   1
      Top             =   450
      Width           =   1395
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
      ScaleWidth      =   8640
      TabIndex        =   8
      Top             =   825
      Width           =   8640
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   7320
         TabIndex        =   7
         Top             =   195
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   5925
         TabIndex        =   6
         Top             =   195
         Width           =   1100
      End
      Begin VB.Frame Frame1 
         Height          =   75
         Left            =   -105
         TabIndex        =   9
         Top             =   30
         Width           =   10000
      End
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "执行时锁定"
      Height          =   180
      Left            =   7560
      TabIndex        =   15
      Top             =   90
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "所属组"
      Height          =   180
      Left            =   6000
      TabIndex        =   14
      Top             =   90
      Width           =   540
   End
   Begin VB.Label lblNO 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   180
      Index           =   0
      Left            =   420
      TabIndex        =   0
      Top             =   510
      Width           =   90
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "缺省值"
      Height          =   180
      Left            =   4222
      TabIndex        =   13
      Top             =   105
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "类型"
      Height          =   180
      Left            =   2707
      TabIndex        =   12
      Top             =   120
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "名称"
      Height          =   180
      Left            =   1342
      TabIndex        =   11
      Top             =   120
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "序号"
      Height          =   180
      Left            =   360
      TabIndex        =   10
      Top             =   120
      Width           =   360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      X1              =   15
      X2              =   9000
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
Private mobjPars As RPTPars '入/出
Private arrCustom() As CustomPar
Private intPreIdx As Integer

Private mstrSQL As String
Private mlngSys As Long
Private mobjData As RPTData
Private mobjDatas As RPTDatas
Private mblnOK As Boolean
Private mlngReportID As Long

Public Function ShowMe(objParent As Object, ByVal lngSys As Long, objData As RPTData, objDatas As RPTDatas, _
    ByRef objPars As RPTPars, ByRef strSQL As String, ByVal lngReportID As Long) As Boolean
    
    Set mobjPars = objPars
    mlngSys = lngSys
    mstrSQL = strSQL
    Set mobjData = objData
    Set mobjDatas = objDatas
    mlngReportID = lngReportID
    
    Me.Show 1, objParent
    Set objData = mobjData
    Set objDatas = mobjDatas
    Set objPars = mobjPars
    strSQL = mstrSQL
    ShowMe = mblnOK
End Function

Private Sub cboGroup_Change(Index As Integer)
    Dim IntSelStart As Integer
    IntSelStart = cboGroup(Index).SelStart
    cboGroup(Index).Text = UCase(cboGroup(Index).Text)
    arrCustom(Index).组名 = cboGroup(Index).Text
    cboGroup(Index).SelStart = IntSelStart
End Sub

Private Sub cboGroup_Click(Index As Integer)
    arrCustom(Index).组名 = cboGroup(Index).Text
End Sub

Private Sub cboGroup_GotFocus(Index As Integer)
    '重新获取所有组名
    Dim strGroup As String, arrGroup
    Dim IntGroup As Integer, IntAdd As Integer
    '填充组名
    strGroup = ""
    For IntGroup = 0 To UBound(arrCustom)
        If InStr(1, strGroup, "^" & arrCustom(IntGroup).组名 & ",") = 0 And arrCustom(IntGroup).组名 <> "" Then strGroup = strGroup & "^" & arrCustom(IntGroup).组名 & ","
    Next
    If strGroup <> "" Then
        strGroup = Mid(strGroup, 1, Len(strGroup) - 1)
        arrGroup = Split(strGroup, ",")
        
        '为各组填充组名
        cboGroup(Index).Clear
        For IntAdd = 0 To UBound(arrGroup)
            cboGroup(Index).AddItem Mid(arrGroup(IntAdd), 2)
        Next
        If cboGroup(Index).ListCount <> 0 Then cboGroup(Index).Text = arrCustom(Index).组名
        
        cboGroup(Index).SelStart = 0
        cboGroup(Index).SelLength = 1000
    End If
End Sub

Private Sub cboGroup_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr("~`@#$%^&*()=+][}{'"";/?.>,<\|", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub cboType_Click(Index As Integer)
    '只有无类型参数才可以使用多选选择器
    If cboType(Index).ListIndex <> 3 And cboValue(Index).Text = "选择器定义…" Then
        arrCustom(Index).格式 = 0
    End If
End Sub

Private Sub cboValue_Click(Index As Integer)
    Dim tmpPar As RPTPar, tmpData As RPTData
    Dim blnDo As Boolean
    
    If cboValue(Index).Text Like "*…" Then
        cboValue(Index).ToolTipText = "按 F2 进入" & cboValue(Index).Text
    Else
        cboValue(Index).ToolTipText = ""
    End If
    
    If Visible Then
        If cboValue(Index).Text = "固定值列表…" Then
            '将其它数据源中相同参数的值列表复制过来
            If arrCustom(Index).值列表 = "" Then
                For Each tmpData In mobjDatas
                    If tmpData.名称 <> mobjData.名称 Then
                        For Each tmpPar In tmpData.Pars
                            If tmpPar.名称 = txtName(Index).Text And tmpPar.缺省值 = cboValue(Index).Text Then
                                arrCustom(Index).值列表 = tmpPar.值列表
                                arrCustom(Index).格式 = tmpPar.格式
                                blnDo = True: Exit For
                            End If
                        Next
                    End If
                    If blnDo Then Exit For
                Next
            End If
            
            frmFixValue.bytType = cboType(Index).ListIndex
            frmFixValue.strName = txtName(Index).Text
            frmFixValue.IntSelType = arrCustom(Index).格式
            '可能从选择器定义切换过来
            '不好的分隔符
            If InStr(arrCustom(Index).值列表, "√") > 0 And InStr(arrCustom(Index).值列表, ",") > 0 Then
                frmFixValue.strValues = arrCustom(Index).值列表
            End If
            frmFixValue.Show 1, Me
            If gblnOK Then
                arrCustom(Index).值列表 = frmFixValue.strValues
                arrCustom(Index).格式 = frmFixValue.IntSelType
                Unload frmFixValue
                
                '同时更改其它数据源中相同参数的值列表
                For Each tmpData In mobjDatas
                    If tmpData.名称 <> mobjData.名称 Then
                        For Each tmpPar In tmpData.Pars
                            If tmpPar.名称 = txtName(Index).Text And tmpPar.缺省值 = cboValue(Index).Text Then
                                tmpPar.值列表 = arrCustom(Index).值列表
                                tmpPar.格式 = arrCustom(Index).格式
                            End If
                        Next
                    End If
                Next
            End If
        ElseIf cboValue(Index).Text = "选择器定义…" Then
            '将其它数据源中相同参数的值复制过来
            If arrCustom(Index).明细SQL = "" Then
                For Each tmpData In mobjDatas
                    If tmpData.名称 <> mobjData.名称 Then
                        For Each tmpPar In tmpData.Pars
                            If tmpPar.名称 = txtName(Index).Text And tmpPar.缺省值 = cboValue(Index).Text Then
                                arrCustom(Index).明细SQL = tmpPar.明细SQL
                                arrCustom(Index).明细字段 = tmpPar.明细字段
                                arrCustom(Index).分类SQL = tmpPar.分类SQL
                                arrCustom(Index).分类字段 = tmpPar.分类字段
                                arrCustom(Index).对象 = tmpPar.对象
                                arrCustom(Index).格式 = tmpPar.格式
                                arrCustom(Index).值列表 = tmpPar.值列表
                                blnDo = True: Exit For
                            End If
                        Next
                    End If
                    If blnDo Then Exit For
                Next
            End If
            '只有无类型参数才可以使用多选选择器
            If cboType(Index).ListIndex <> 3 Then arrCustom(Index).格式 = 0
            
            frmSelValue.mstrSQLList = arrCustom(Index).明细SQL
            frmSelValue.mstrSQLTree = arrCustom(Index).分类SQL
            frmSelValue.mstrFLDList = arrCustom(Index).明细字段
            frmSelValue.mstrFLDTree = arrCustom(Index).分类字段
            frmSelValue.mstrObj = arrCustom(Index).对象
            '可能从固定值切换过来
            frmSelValue.mstrDef = IIF(InStr(arrCustom(Index).值列表, "√") > 0, "", arrCustom(Index).值列表)
            
            frmSelValue.mbytType = cboType(Index).ListIndex
            frmSelValue.mstrName = txtName(Index).Text
            frmSelValue.mblnMulti = arrCustom(Index).格式 = 1
            frmSelValue.mlngSys = mlngSys
            Set frmSelValue.mobjDatas = mobjDatas
            Set frmSelValue.mobjData = mobjData
            
            frmSelValue.Show 1, Me
            If gblnOK Then
                arrCustom(Index).明细SQL = frmSelValue.mstrSQLList
                arrCustom(Index).分类SQL = frmSelValue.mstrSQLTree
                arrCustom(Index).明细字段 = frmSelValue.mstrFLDList
                arrCustom(Index).分类字段 = frmSelValue.mstrFLDTree
                arrCustom(Index).对象 = frmSelValue.mstrObj
                arrCustom(Index).值列表 = frmSelValue.mstrDef
                arrCustom(Index).格式 = IIF(frmSelValue.mblnMulti, 1, 0)
                Unload frmSelValue
                
                '同时更改其它数据源中相同参数的值
                For Each tmpData In mobjDatas
                    If tmpData.名称 <> mobjData.名称 Then
                        For Each tmpPar In tmpData.Pars
                            If tmpPar.名称 = txtName(Index).Text And tmpPar.缺省值 = cboValue(Index).Text Then
                                tmpPar.明细SQL = arrCustom(Index).明细SQL
                                tmpPar.分类SQL = arrCustom(Index).分类SQL
                                tmpPar.明细字段 = arrCustom(Index).明细字段
                                tmpPar.分类字段 = arrCustom(Index).分类字段
                                tmpPar.对象 = arrCustom(Index).对象
                                tmpPar.值列表 = arrCustom(Index).值列表
                                tmpPar.格式 = arrCustom(Index).格式
                            End If
                        Next
                    End If
                Next
            End If
        End If
    End If
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

Private Sub chkLock_GotFocus(Index As Integer)
    chkLock(Index).Tag = "" & chkLock(Index).BackColor
    chkLock(Index).BackColor = &HC0C0C0
End Sub

Private Sub chkLock_LostFocus(Index As Integer)
    chkLock(Index).BackColor = Val(chkLock(Index).Tag)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer, j As Integer
    Dim tmpPar As RPTPar, tmpData As RPTData
    Dim curPar As RPTPar
    Dim strAddDelBefor As String, strAddDelAfter As String, strInfo As String '判断是否删除和新增了参数
    Dim strSQL As String, rsTmp As Recordset
    
    If Not CheckFormInput(Me, True) Then Exit Sub
    
    '检查参数合法性
    For i = 0 To lblNO.UBound
        If txtName(i).Text = "" Then
            MsgBox "第 " & i & " 的个参数没有输入参数名称！", vbInformation, App.Title
            txtName(i).SetFocus: Exit Sub
        End If
        If TLen(txtName(i).Text) > 20 Then
            MsgBox "第 " & i & " 的个参数名称长度不能超过20个字符！", vbInformation, App.Title
            txtName(i).SetFocus: Exit Sub
        End If
        
        For j = 0 To lblNO.UBound
            If j <> i And txtName(i).Text = txtName(j).Text Then
                MsgBox "第 " & j & " 的个参数名称与第 " & i & " 的个参数名称重复！", vbInformation, App.Title
                txtName(j).SetFocus: Exit Sub
            End If
        Next
        
        If TLen(cboValue(i).Text) > 255 Then
            MsgBox "第 " & i & " 的个参数 [" & txtName(i).Text & "] 缺省值长度不能超过255个字符！", vbInformation, App.Title
            cboValue(i).SetFocus: Exit Sub
        End If
        If TLen(cboGroup(i).Text) > 30 Then
            MsgBox "第 " & i & " 的个参数 [" & txtName(i).Text & "] 的组名长度不能超过30个字符！", vbInformation, App.Title
            cboGroup(i).SetFocus: Exit Sub
        End If
        
        If cboValue(i).Text <> "" And Not cboValue(i).Text Like "*…" Then
            If cboType(i).ListIndex = 1 Then
                If Not IsNumeric(cboValue(i).Text) Then
                    MsgBox "第 " & i & " 的个参数 [" & txtName(i).Text & "] 缺省值类型应该为数字型！", vbInformation, App.Title
                    cboValue(i).SetFocus: Exit Sub
                End If
            ElseIf cboType(i).ListIndex = 2 Then
                If Not IsDate(cboValue(i).Text) And cboValue(i).ListIndex = -1 Then
                    MsgBox "第 " & i & " 的个参数 [" & txtName(i).Text & "] 缺省值类型应该为日期/时间型！", vbInformation, App.Title
                    cboValue(i).SetFocus: Exit Sub
                End If
            End If
        End If
            
        '报表中参数名称如果相同则类型和向缺省值应该相同
'        For Each tmpData In frmSQLEdit.objDatas
'            If tmpData.名称 <> frmSQLEdit.objData.名称 Then
'                For Each tmpPar In tmpData.Pars
'                    If tmpPar.名称 = txtName(i).Text And (tmpPar.类型 <> cboType(i).ListIndex Or tmpPar.缺省值 <> cboValue(i).Text) Then
'                        MsgBox "在报表其它数据源中发现有相同名称的参数""" & txtName(i).Text & """,但它们的类型或缺省值不相同！", vbInformation, App.Title: Exit Sub
'                    End If
'                Next
'            End If
'        Next

'        '检查当前参数是否与其他数据源的参数同名
'        For Each tmpData In mobjDatas
'            If tmpData.名称 <> mobjData.名称 Then
'                For Each tmpPar In tmpData.Pars
'                    If UCase(Trim(tmpPar.名称)) = UCase(Trim(txtName(i).Text)) Then
'                        MsgBox "参数名“" & Trim(txtName(i).Text) & "”与其他数据源的参数名重名，请检查！", vbInformation, App.Title
'                        Exit Sub
'                    End If
'                Next
'            End If
'        Next
        
        '自定义内容检查
        If cboValue(i).Text = "固定值列表…" Then
            If arrCustom(i).值列表 = "" Then
                MsgBox "第 " & i & " 的个参数 [" & txtName(i).Text & "] 还没有定义可选择的固定值列表！", vbInformation, App.Title
                cboValue(i).SetFocus: Exit Sub
            End If
            '检查类型
            Select Case cboType(i).ListIndex
                Case 1 '数字
                    '不好的分隔符
                    For j = 0 To UBound(Split(arrCustom(i).值列表, "|"))
                        If Not IsNumeric(Split(Split(arrCustom(i).值列表, "|")(j), ",")(1)) Then
                            MsgBox "第 " & i & " 的个参数 [" & txtName(i).Text & "] 的固定值列表中存在非数字型绑定值！", vbInformation, App.Title
                            cboValue(i).SetFocus: Exit Sub
                        End If
                    Next
                Case 2 '日期
                    '不好的分隔符
                    For j = 0 To UBound(Split(arrCustom(i).值列表, "|"))
                        If Not IsDate(Split(Split(arrCustom(i).值列表, "|")(j), ",")(1)) Then
                            MsgBox "第 " & i & " 的个参数 [" & txtName(i).Text & "] 的固定值列表中存在非日期型绑定值！", vbInformation, App.Title
                            cboValue(i).SetFocus: Exit Sub
                        End If
                    Next
            End Select
        End If
        
        If cboValue(i).Text = "选择器定义…" Then
            If arrCustom(i).明细SQL = "" Then
                MsgBox "第 " & i & " 的个参数 [" & txtName(i).Text & "] 还没有定义选择器的内容！", vbInformation, App.Title
                cboValue(i).SetFocus: Exit Sub
            End If
            '检查类型(之所以要再判断一次因为用户可能改了类型)
            For j = 0 To UBound(Split(arrCustom(i).明细字段, "|"))
                If InStr(Split(Split(arrCustom(i).明细字段, "|")(j), ",")(2), "&B") > 0 Then
                    If cboType(i).ListIndex = 1 Then
                        Select Case CLng(Split(Split(arrCustom(i).明细字段, "|")(j), ",")(1))
                            Case adNumeric, adVarNumeric  '正常数字型
                            Case Else '非数字型
                                If MsgBox("第 " & i & " 的个参数 [" & txtName(i).Text & "] 的选择器绑定字段 [" & Split(Split(arrCustom(i).明细字段, "|")(j), ",")(0) & "] 非数字类型,要继续吗？", vbQuestion + vbYesNo, App.Title) = vbNo Then
                                    cboValue(i).SetFocus: Exit Sub
                                End If
                        End Select
                    ElseIf cboType(i).ListIndex = 2 Then
                        Select Case CLng(Split(Split(arrCustom(i).明细字段, "|")(j), ",")(1))
                            Case adDBTimeStamp '正常日期型
                            Case Else '非日期型
                                If MsgBox("第 " & i & " 的个参数 [" & txtName(i).Text & "] 的选择器绑定字段 [" & Split(Split(arrCustom(i).明细字段, "|")(j), ",")(0) & "] 非日期类型,要继续吗？", vbQuestion + vbYesNo, App.Title) = vbNo Then
                                    cboValue(i).SetFocus: Exit Sub
                                End If
                        End Select
                    End If
                End If
            Next
        End If
    Next
    
    '第种组名至少含有两个参数(且只能是相邻的参数才能属于同一组)单选框不能设置所属组
    Dim IntPar As Integer, StrPar As String, IntUseCount As Integer
    IntUseCount = 0: StrPar = ""
    For IntPar = 0 To UBound(arrCustom)
        If arrCustom(IntPar).格式 = 1 And arrCustom(IntPar).组名 <> "" Then
            MsgBox "参数" & IntPar & "不能设置所属组（因该参数的选择模式是单选框）！", vbInformation, App.Title
            cboGroup(IntPar).SetFocus
            Exit Sub
        End If
    Next
    For IntPar = 0 To UBound(arrCustom)
        If StrPar <> arrCustom(IntPar).组名 Then
            If arrCustom(IntPar).组名 = "" Then
                If Not (IntUseCount = 0 Or IntUseCount > 1) Then
                    MsgBox "每组最少须有两个参数（参数" & IntPar & "）！", vbInformation, App.Title
                    cboGroup(IntPar).SetFocus
                    Exit Sub
                End If
                StrPar = arrCustom(IntPar).组名
                IntUseCount = 1
            Else
                If IntUseCount = 0 Or IntUseCount > 1 Or StrPar = "" Then
                    StrPar = arrCustom(IntPar).组名
                    IntUseCount = 1
                Else
                    MsgBox "每组最少须有两个参数（参数" & IntPar & "）！", vbInformation, App.Title
                    cboGroup(IntPar).SetFocus
                    Exit Sub
                End If
            End If
        Else
            IntUseCount = IntUseCount + 1
        End If
    Next
    If Not (IntUseCount = 0 Or IntUseCount > 1 Or StrPar = "") Then
        MsgBox "每组最少须有两个参数（参数" & IntPar - 1 & "）！", vbInformation, App.Title
        cboGroup(IntPar - 1).SetFocus
        Exit Sub
    End If
    
    '获取之前的参数
    For i = 1 To mobjPars.count
        strAddDelBefor = strAddDelBefor & "," & mobjPars.Item(i).名称
    Next
    '确定输入内容
    Set mobjPars = New RPTPars
    For i = 0 To lblNO.UBound
        '清除不使用数据源的参数相关内容,否则可能影响授权
        If cboValue(i).Text <> "选择器定义…" Then
            arrCustom(i).明细SQL = ""
            arrCustom(i).明细字段 = ""
            arrCustom(i).分类SQL = ""
            arrCustom(i).分类字段 = ""
            arrCustom(i).对象 = ""
        End If
        Set curPar = Nothing
        Set curPar = mobjPars.Add(arrCustom(i).组名, CByte(i), txtName(i).Text, cboType(i).ListIndex, cboValue(i).Text _
                        , arrCustom(i).格式, arrCustom(i).值列表, arrCustom(i).分类SQL, arrCustom(i).明细SQL _
                        , arrCustom(i).分类字段, arrCustom(i).明细字段, arrCustom(i).对象, "_" & i, , chkLock(i).Value)
                        
        '同时自动替换其它数据源中相同名称参数的内容
        For Each tmpData In mobjDatas
            If tmpData.名称 <> mobjData.名称 Then
                For Each tmpPar In tmpData.Pars
                    If tmpPar.名称 = curPar.名称 Then
                        tmpPar.格式 = curPar.格式
                        tmpPar.类型 = curPar.类型
                        tmpPar.缺省值 = curPar.缺省值
                        tmpPar.值列表 = curPar.值列表
                        tmpPar.明细SQL = curPar.明细SQL
                        tmpPar.明细字段 = curPar.明细字段
                        tmpPar.分类SQL = curPar.分类SQL
                        tmpPar.分类字段 = curPar.分类字段
                        tmpPar.对象 = curPar.对象
                        tmpPar.是否锁定 = curPar.是否锁定
                    End If
                Next
            End If
        Next
    Next
    '获取之前的参数
    For i = 1 To mobjPars.count
        strAddDelAfter = strAddDelAfter & "," & mobjPars.Item(i).名称
    Next
    If strAddDelAfter <> strAddDelBefor Then
        '提示关联此报表的报表
        strSQL = "Select Distinct b.编号, b.名称 From Zlrptrelation A, zlReports B Where a.报表id = b.Id And a.关联报表id = [1] "
        On Error GoTo errH
        Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, mlngReportID)
        Do While Not rsTmp.EOF
            strInfo = strInfo & vbCrLf & rsTmp!名称 & "(" & rsTmp!编号 & ")"
            rsTmp.MoveNext
        Loop
        strInfo = Mid(strInfo, 2)
        If strInfo <> "" Then
            MsgBox "以下报表会关联查询本报表，本报表调整了参数后，可能需要调整以下报表的关联信息，请检查：" & strInfo, vbInformation, Me.Caption
        End If
    End If
    
    mblnOK = True
    Hide
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: SendKeys "{Tab}"
End Sub

Private Sub Form_Load()
    Dim intCount As Integer
    Dim i As Integer
    
    mblnOK = False
    intPreIdx = -1
    
    intCount = GetParCount(mstrSQL)
    
    ReDim arrCustom(intCount - 1) As CustomPar
    
    For i = 0 To intCount - 1
        If i <> 0 Then
            Load lblNO(i): lblNO(i).Left = lblNO(0).Left: lblNO(i).Top = lblNO(0).Top + 450 * i: lblNO(i).Visible = True
            Load txtName(i): txtName(i).Left = txtName(0).Left: txtName(i).Top = txtName(0).Top + 450 * i: txtName(i).TabIndex = txtName(0).TabIndex + 5 * i: txtName(i).Visible = True
            Load cboType(i): cboType(i).Left = cboType(0).Left: cboType(i).Top = cboType(0).Top + 450 * i: cboType(i).TabIndex = cboType(0).TabIndex + 5 * i: cboType(i).Visible = True
            Load cboValue(i): cboValue(i).Left = cboValue(0).Left: cboValue(i).Top = cboValue(0).Top + 450 * i: cboValue(i).TabIndex = cboValue(0).TabIndex + 5 * i: cboValue(i).Visible = True
            Load cboGroup(i): cboGroup(i).Left = cboGroup(0).Left: cboGroup(i).Top = cboGroup(0).Top + 450 * i: cboGroup(i).TabIndex = cboGroup(0).TabIndex + 5 * i: cboGroup(i).Visible = True
            Load chkLock(i): chkLock(i).Left = chkLock(0).Left: chkLock(i).Top = chkLock(0).Top + 450 * i: chkLock(i).TabIndex = chkLock(0).TabIndex + 5 * i: chkLock(i).Visible = True: chkLock(i).Value = 0
        End If
        lblNO(i).Caption = i
        cboType(i).AddItem "字符": cboType(i).AddItem "数字": cboType(i).AddItem "日期": cboType(i).AddItem "无类型"
        cboValue(i).AddItem "&当前日期" '宏表达式
        cboValue(i).AddItem "&当前日期时间"
        
        cboValue(i).AddItem "&当天开始时间"
        cboValue(i).AddItem "&当天结束时间"
        cboValue(i).AddItem "&前一天开始时间"
        cboValue(i).AddItem "&前一天结束时间"
        cboValue(i).AddItem "&前一天同时间"
        cboValue(i).AddItem "&后一天同时间"
        cboValue(i).AddItem "&后一天结束时间"
        cboValue(i).AddItem "&后一天日期"
        
        cboValue(i).AddItem "&前一周日期"
        cboValue(i).AddItem "&前一月日期"
        cboValue(i).AddItem "&前一季日期"
        cboValue(i).AddItem "&前一年日期"
        
        cboValue(i).AddItem "&下一周日期"
        cboValue(i).AddItem "&下一月日期"
        cboValue(i).AddItem "&下一季日期"
        cboValue(i).AddItem "&下一年日期"
        
        cboValue(i).AddItem "&本月初时间"
        cboValue(i).AddItem "&本月末时间"
        cboValue(i).AddItem "&上月初时间"
        cboValue(i).AddItem "&上月末时间"
        cboValue(i).AddItem "&本年初时间"
        cboValue(i).AddItem "&本年末时间"
        cboValue(i).AddItem "&上年初时间"
        cboValue(i).AddItem "&上年末时间"
        
        '新增自定义内容
        cboValue(i).AddItem "固定值列表…"
        cboValue(i).AddItem "选择器定义…"
        
        If mobjPars.count >= i + 1 Then '尽量保持原有设置
            txtName(i).Text = mobjPars("_" & i).名称
            cboType(i).ListIndex = mobjPars("_" & i).类型
            If Left(mobjPars("_" & i).缺省值, 1) = "&" Or mobjPars("_" & i).缺省值 Like "*…" Then
                cboValue(i).ListIndex = GetCboIndex(cboValue(i), mobjPars("_" & i).缺省值)
            Else
                cboValue(i).Text = mobjPars("_" & i).缺省值
            End If
            chkLock(i).Value = IIF(mobjPars("_" & i).是否锁定, 1, 0)
            
            '自定义内容
            arrCustom(i).值列表 = mobjPars("_" & i).值列表
            arrCustom(i).分类SQL = mobjPars("_" & i).分类SQL
            arrCustom(i).明细SQL = mobjPars("_" & i).明细SQL
            arrCustom(i).分类字段 = mobjPars("_" & i).分类字段
            arrCustom(i).明细字段 = mobjPars("_" & i).明细字段
            arrCustom(i).对象 = mobjPars("_" & i).对象
            arrCustom(i).格式 = mobjPars("_" & i).格式
            arrCustom(i).组名 = mobjPars("_" & i).组名
        Else
            txtName(i).Text = ""
            cboType(i).ListIndex = 0
            cboValue(i).Text = ""
        End If
    Next
    Call LoadGroup
    
    Height = txtName(txtName.UBound).Top + 1365
End Sub

Private Sub txtName_GotFocus(Index As Integer)
    SelAll txtName(Index)
End Sub

Private Sub txtName_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim tmpData As RPTData, tmpPar As RPTPar
    
    If KeyCode = 13 And txtName(Index).Text <> "" Then
        For Each tmpData In mobjDatas
            If tmpData.名称 <> mobjData.名称 Then
                For Each tmpPar In tmpData.Pars
                    If tmpPar.名称 = txtName(Index).Text Then
                        cboType(Index).ListIndex = tmpPar.类型
                        cboValue(Index).ListIndex = GetCboIndex(cboValue(Index), tmpPar.缺省值)
                        If cboValue(Index).ListIndex = -1 Then cboValue(Index).Text = tmpPar.缺省值
                    End If
                Next
            End If
        Next
    End If
End Sub

Private Sub txtName_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr("~`@#$%^&*()=+][}{'"";/?.>,<\|", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub LoadGroup()
    Dim ItemPar As RPTPar, strGroup As String, arrGroup
    Dim IntGroup As Integer, IntAdd As Integer
    '填充组名
    strGroup = ""
    For Each ItemPar In mobjPars
        If InStr(1, strGroup, "^" & ItemPar.组名 & ",") = 0 And ItemPar.组名 <> "" Then strGroup = strGroup & "^" & ItemPar.组名 & ","
    Next
    If strGroup <> "" Then
        strGroup = Mid(strGroup, 1, Len(strGroup) - 1)
        arrGroup = Split(strGroup, ",")
        
        '为各组填充组名
        For IntGroup = 0 To cboGroup.UBound
            cboGroup(IntGroup).Clear
            For IntAdd = 0 To UBound(arrGroup)
                cboGroup(IntGroup).AddItem Mid(arrGroup(IntAdd), 2)
            Next
            If cboGroup(IntGroup).ListCount <> 0 Then cboGroup(IntGroup).Text = arrCustom(IntGroup).组名
        Next
    End If
End Sub
