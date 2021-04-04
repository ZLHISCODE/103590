VERSION 5.00
Begin VB.Form frmEPRFileEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病历文件命名"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   Icon            =   "frmEPRFileEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraEditType 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   765
      TabIndex        =   21
      Top             =   3195
      Width           =   4905
      Begin VB.CheckBox chk门诊快捷病历 
         Caption         =   "门诊快捷病历"
         Height          =   180
         Left            =   1780
         TabIndex        =   24
         Top             =   112
         Visible         =   0   'False
         Width           =   1400
      End
      Begin VB.OptionButton optEditType 
         Caption         =   "表格式病历编辑器"
         Height          =   225
         Index           =   1
         Left            =   0
         TabIndex        =   23
         Top             =   360
         Width           =   1845
      End
      Begin VB.OptionButton optEditType 
         Caption         =   "全文病历编辑器"
         Height          =   225
         Index           =   0
         Left            =   15
         TabIndex        =   22
         Top             =   90
         Value           =   -1  'True
         Width           =   1665
      End
   End
   Begin VB.ComboBox cbo等级 
      Height          =   300
      Left            =   1455
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   2025
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.CheckBox chkCopy 
      Caption         =   "复制(&V)"
      Height          =   195
      Left            =   3210
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   473
      Width           =   2670
   End
   Begin VB.ComboBox cboKind 
      Height          =   300
      Left            =   1455
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   420
      Width           =   1425
   End
   Begin VB.OptionButton optPage 
      Caption         =   "使用共用页面(&2)"
      Height          =   180
      Index           =   1
      Left            =   780
      TabIndex        =   12
      Top             =   2910
      Width           =   1725
   End
   Begin VB.TextBox txtPageName 
      Height          =   300
      Left            =   3210
      TabIndex        =   14
      Top             =   2460
      Width           =   2490
   End
   Begin VB.TextBox txtPageNo 
      Enabled         =   0   'False
      Height          =   300
      Left            =   2535
      TabIndex        =   13
      Top             =   2460
      Width           =   645
   End
   Begin VB.OptionButton optPage 
      Caption         =   "使用新建页面(&1)"
      Height          =   180
      Index           =   0
      Left            =   780
      TabIndex        =   11
      Top             =   2505
      Value           =   -1  'True
      Width           =   1725
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4600
      TabIndex        =   17
      Top             =   4155
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3480
      TabIndex        =   16
      Top             =   4155
      Width           =   1100
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   -45
      TabIndex        =   20
      Top             =   4035
      Width           =   6390
   End
   Begin VB.ComboBox cboPage 
      Enabled         =   0   'False
      Height          =   300
      Left            =   2535
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   2850
      Width           =   3165
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   0
      Left            =   765
      TabIndex        =   18
      Top             =   840
      Width           =   5325
   End
   Begin VB.TextBox txt编号 
      Height          =   300
      Left            =   1455
      TabIndex        =   4
      Top             =   975
      Width           =   645
   End
   Begin VB.TextBox txt名称 
      Height          =   300
      Left            =   3210
      TabIndex        =   6
      Top             =   975
      Width           =   2490
   End
   Begin VB.TextBox txt说明 
      Height          =   540
      Left            =   1455
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   1380
      Width           =   4245
   End
   Begin VB.Label lbl等级 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "适用(&A)                 及以上护理等级的病人。"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   780
      TabIndex        =   9
      Top             =   2085
      Visible         =   0   'False
      Width           =   4140
   End
   Begin VB.Label lblKind 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "种类(&K)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   795
      TabIndex        =   0
      Top             =   480
      Width           =   630
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "对病历文件进行命名，并指定其正式打印输出的页面。"
      Height          =   180
      Left            =   780
      TabIndex        =   19
      Top             =   120
      Width           =   4335
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   180
      Picture         =   "frmEPRFileEdit.frx":058A
      Top             =   150
      Width           =   480
   End
   Begin VB.Label lbl编号 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "编号(&D)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   780
      TabIndex        =   3
      Top             =   1035
      Width           =   630
   End
   Begin VB.Label lbl名称 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "名称(&N)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2505
      TabIndex        =   5
      Top             =   1035
      Width           =   630
   End
   Begin VB.Label lbl说明 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "说明(&M)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   780
      TabIndex        =   7
      Top             =   1440
      Width           =   630
   End
End
Attribute VB_Name = "frmEPRFileEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------
'说明：
'   1、上级程序通过本窗体ShowMe函数，将父窗体、编辑文件ID,编辑状态等信息传递进入本程序
'   2、编辑状态：由Me.tag存放，分别为"增加"、"修改"，由上级程序通过ShowMe传入
'---------------------------------------------------
Private mlngFileID As Long          '被编辑或用于复制增加的文件ID，修改、查阅时由上级程序通过ShowMe传递进入,增加时为0.
Private mblnSpecial As Boolean      '传入文件是否特殊病历
Private mblnSpecialWave As Boolean  '传入的文件是否是专科体温单
Private mblnOK As Boolean

Public Function ShowMe(ByVal frmParent As Object, ByVal strKinds As String, ByVal blnAdd As Boolean, Optional ByVal lngFileID As Long) As Long
    '---------------------------------------------------
    '功能：上级程序调用本窗体的，传递参数，并显示窗体
    '---------------------------------------------------
Dim rsTemp As New ADODB.Recordset
Dim lngCount As Long
    With Me.cbo等级
        .Clear
        .AddItem "0-特级护理"
        .AddItem "1-一级护理"
        .AddItem "2-二级护理"
        .AddItem "3-三级护理"
        .ListIndex = .ListCount - 1
    End With
    
    If InStr(1, "," & strKinds, ",1") > 0 Then Me.cboKind.AddItem "1-门诊病历"
    If InStr(1, "," & strKinds, ",2") > 0 Then Me.cboKind.AddItem "2-住院病历"
    If InStr(1, "," & strKinds, ",3") > 0 Then Me.cboKind.AddItem "3-护理记录"
    If InStr(1, "," & strKinds, ",4") > 0 Then Me.cboKind.AddItem "4-护理病历"
    If InStr(1, "," & strKinds, ",5") > 0 Then Me.cboKind.AddItem "5-疾病证明报告"
    If InStr(1, "," & strKinds, ",6") > 0 Then Me.cboKind.AddItem "6-知情文件"
    If Me.cboKind.ListCount <= 1 Then Me.cboKind.Enabled = False
    
    If blnAdd Then
        Me.Tag = "增加": mlngFileID = 0
    Else
        Me.Tag = "修改": mlngFileID = lngFileID
    End If
    
    mblnSpecialWave = False
    '当前数据读取
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select l.种类, l.编号, l.名称, l.说明, l.保留,l.子类, f.编号 As 页面号, f.名称 As 页面名, Nvl(f.报表, 0) As 等级" & _
            " From 病历文件列表 l, 病历页面格式 f" & _
            " Where l.种类 = f.种类(+) And l.页面 = f.编号(+) And l.Id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngFileID)
    With rsTemp
        Me.txt编号.MaxLength = .Fields("编号").DefinedSize
        Me.txt名称.MaxLength = .Fields("名称").DefinedSize
        Me.txt说明.MaxLength = .Fields("说明").DefinedSize
        Me.txtPageNo.MaxLength = .Fields("页面号").DefinedSize
        Me.txtPageName.MaxLength = .Fields("页面名").DefinedSize
        If .RecordCount > 0 Then
            mblnSpecial = (NVL(!保留, 0) < 0 Or NVL(!保留, 0) = 2)
            mblnSpecialWave = NVL(!保留, 0) < 0 And NVL(!种类, 0) = 3 And NVL(!子类) = "1"
            Me.cbo等级.ListIndex = !等级
            If Me.Tag = "增加" Then
                Me.txtPageNo.Text = !页面号: Me.txtPageName.Text = !页面名
                '特殊病历不能用于复制
                Me.chkCopy.Tag = lngFileID: Me.chkCopy.Caption = "复制(&V)" & !名称
                If mblnSpecial Then Me.chkCopy.Value = vbUnchecked: Me.chkCopy.Visible = False
                If mblnSpecialWave Then Me.chkCopy.Value = vbChecked: Me.chkCopy.Visible = True: Me.chkCopy.Enabled = False
            Else
                Me.txt编号.Text = !编号: Me.txt名称.Text = !名称: Me.txt说明.Text = "" & !说明
                Me.txtPageNo.Text = !页面号: Me.txtPageName.Text = !页面名
                Me.chkCopy.Value = vbUnchecked: Me.chkCopy.Visible = False
                Me.cboKind.Enabled = False: optEditType(0).Enabled = False: optEditType(1).Enabled = False
                If NVL(!保留, 0) < 0 Then optEditType(0).Value = False: optEditType(1).Value = False
                If NVL(!保留, 0) = 0 Or NVL(!保留, 0) = 1 Then optEditType(0).Value = True: optEditType(1).Value = False
                
            End If
            Me.cboKind.Tag = !种类
            For lngCount = 0 To Me.cboKind.ListCount - 1
                If Val(Left(Me.cboKind.List(lngCount), 1)) = !种类 Then
                    Me.cboKind.ListIndex = lngCount
                    Exit For
                End If
            Next
            If !编号 = "" & !页面号 Or Val(Me.cboKind.Tag) = 3 Then
                Me.optPage(0).Value = True
            Else
                Me.optPage(1).Value = True
                For lngCount = 0 To Me.cboPage.ListCount - 1
                    If Val(Me.cboPage.List(lngCount)) = Val("" & !页面号) Then
                        Me.cboPage.ListIndex = lngCount
                        Exit For
                    End If
                Next
            End If
            If Me.Tag = "修改" Then
                If NVL(!保留, 0) = 2 Then
                    optEditType(0).Value = False: optEditType(1).Value = True: optPage(1).Enabled = False: cboPage.Enabled = False
                ElseIf NVL(!保留, 0) = 3 Then
                    chk门诊快捷病历.Value = 1
                End If
            End If
        Else
            If Me.Tag = "增加" Then
                Me.cboKind.ListIndex = 0
            Else
                MsgBox "指定文件丢失！(可能被其他用户删除)", vbInformation, gstrSysName
                ShowMe = 0: Unload Me: Exit Function
            End If
        End If
    End With
    
    '显示窗体
    Me.Show vbModal, frmParent
    If mblnOK = False Then ShowMe = 0: Unload Me: Exit Function
    ShowMe = mlngFileID
    Unload Me: Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cboKind_Click()
Dim intKind As Integer
Dim rsTemp As New ADODB.Recordset
    intKind = Left(Me.cboKind.Text, 1)
    
    If Me.Tag = "增加" Then
        gstrSQL = "Select nvl(max(编号),'" & String(Me.txt编号.MaxLength, "0") & "') as 编号 From 病历文件列表 Where 种类 = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, intKind)
        Me.txt编号.Text = Format(Val(rsTemp!编号) + 1, String(Me.txt编号.MaxLength, "0"))
        
        If Val(Me.cboKind.Tag) = intKind Then
            Me.chkCopy.Enabled = Not mblnSpecialWave
        Else
            Me.chkCopy.Value = vbUnchecked: Me.chkCopy.Enabled = False
        End If
    End If
    
    If intKind = 3 Then
        Me.lbl等级.Visible = True: Me.cbo等级.Visible = True
    Else
        Me.lbl等级.Visible = False: Me.cbo等级.Visible = False
    End If
    chk门诊快捷病历.Visible = intKind = 1
    chk门诊快捷病历.Enabled = optEditType(0).Value = True And Me.Tag = "增加"
    
    Me.cboPage.Clear
    Select Case intKind
    Case 2, 4   '2-住院病历;4-护理病历
        gstrSQL = "Select f.编号, f.名称, Count(l.ID) As 使用" & _
                " From 病历页面格式 f, 病历文件列表 l" & _
                " Where f.种类 = l.种类 And f.编号 = l.页面 And l.保留 Between 0 And 1 And f.种类 = [1]" & _
                " Group By f.编号, f.名称" & _
                " Order By 编号"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, intKind)
        With rsTemp
            If Me.Tag = "增加" Then
                Do While Not .EOF
                    Me.cboPage.AddItem !编号 & "-" & !名称
                    .MoveNext
                Loop
                If Me.cboPage.ListCount = 0 Then
                    Me.optPage(0).Value = True: Me.optPage(0).Enabled = False
                    Me.optPage(1).Value = False: Me.optPage(1).Enabled = False
                Else
                    Me.optPage(0).Enabled = True: Me.optPage(1).Enabled = True
                    Me.cboPage.ListIndex = 0
                End If
            Else
                Do While Not .EOF
                    If !编号 <> Trim(Me.txt编号.Text) Then
                        Me.cboPage.AddItem !编号 & "-" & !名称
                    Else
                        Me.txtPageNo.Text = !编号: Me.txtPageName.Text = !名称
                        If !使用 > 1 Then
                            Me.cboPage.AddItem !编号 & "-" & !名称
                            Me.cboPage.ListIndex = Me.cboPage.NewIndex
                        Else
                            Me.optPage(0).Value = True
                        End If
                    End If
                    .MoveNext
                Loop
                If mblnSpecial Then
                    Me.optPage(0).Value = True: Me.optPage(0).Enabled = False
                    Me.optPage(1).Value = False: Me.optPage(1).Enabled = False
                    Me.txtPageNo.Text = Me.txt编号.Text: Me.txtPageNo.Enabled = False
                    Me.txtPageName.Text = Me.txt名称.Text: Me.txtPageName.Enabled = False
                ElseIf Me.cboPage.ListCount = 0 Or mblnSpecial Then
                    Me.optPage(0).Value = True: Me.optPage(0).Enabled = False
                    Me.optPage(1).Value = False: Me.optPage(1).Enabled = False
                Else
                    Me.optPage(0).Enabled = True: Me.optPage(1).Enabled = True
                End If
            End If
        End With
    Case Else
        Me.optPage(0).Value = True: Me.optPage(0).Enabled = False
        Me.optPage(1).Value = False: Me.optPage(1).Enabled = False
        Me.txtPageNo.Text = Me.txt编号.Text: Me.txtPageNo.Enabled = False
        Me.txtPageName.Text = Me.txt名称.Text: Me.txtPageName.Enabled = False
    End Select
    
    If Me.Tag = "增加" Then '新增时对护理记录限制
        If intKind = 3 Then '护理记录
            optEditType(0).Enabled = False: optEditType(1).Enabled = False
        ElseIf intKind = 5 Then  '疾病申报卡
            optEditType(0).Enabled = True: optEditType(1).Enabled = True
        Else
            optEditType(0).Enabled = True: optEditType(1).Enabled = True
        End If
    End If
    
    optEditType(0).Value = True
End Sub

Private Sub cboKind_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cboPage_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo等级_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chkCopy_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False: Me.Hide
End Sub

Private Sub cmdOK_Click()
    Dim intType As Integer
    
    If Trim(Me.txt编号.Text) = "" Then MsgBox "请输入编号！", vbInformation, gstrSysName: Me.txt编号.SetFocus: Exit Sub
    If Len(Me.txt编号.Text) < Me.txt编号.MaxLength Then MsgBox "编号长度不足！", vbInformation, gstrSysName: Me.txt编号.SetFocus: Exit Sub
    If Trim(Me.txt名称.Text) = "" Then MsgBox "请输入名称！", vbInformation, gstrSysName: Me.txt名称.SetFocus: Exit Sub
    If LenB(StrConv(Trim(Me.txt名称.Text), vbFromUnicode)) > Me.txt名称.MaxLength Then
        MsgBox "名称超长（最多" & Me.txt名称.MaxLength & "个字符或等长的汉字）！", vbInformation, gstrSysName: Me.txt名称.SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txt说明.Text), vbFromUnicode)) > Me.txt说明.MaxLength Then
        MsgBox "说明超长（最多" & Me.txt说明.MaxLength & "个字符或等长的汉字）！", vbInformation, gstrSysName: Me.txt说明.SetFocus: Exit Sub
    End If
    If Me.optPage(0).Value Then
        If Trim(Me.txtPageName.Text) = "" Then MsgBox "请输入页面名称！", vbInformation, gstrSysName: Me.txtPageName.SetFocus: Exit Sub
        If LenB(StrConv(Trim(Me.txtPageName.Text), vbFromUnicode)) > Me.txtPageName.MaxLength Then
            MsgBox "页面名称超长（最多" & Me.txt名称.MaxLength & "个字符或等长的汉字）！", vbInformation, gstrSysName: Me.txtPageName.SetFocus: Exit Sub
        End If
    Else
        If Me.cboPage.ListIndex = -1 Then MsgBox "请选择共用页面！", vbInformation, gstrSysName: Me.cboPage.SetFocus: Exit Sub
    End If
    
    '数据保存
    If Me.Tag = "增加" Then
        mlngFileID = zlDatabase.GetNextId("病历文件列表")
        gstrSQL = mlngFileID & "," & Val(Left(Me.cboKind.Text, 1))
    Else
        gstrSQL = mlngFileID
    End If
    gstrSQL = gstrSQL & ",'" & Trim(Me.txt编号.Text) & "','" & Trim(Me.txt名称.Text) & "','" & Replace(Me.txt说明, Chr(vbKeyReturn), "") & "'"
    If Me.optPage(0).Value Then
        gstrSQL = gstrSQL & ",'" & Trim(Me.txtPageNo.Text) & "','" & Trim(Me.txtPageName.Text) & "'"
    Else
        gstrSQL = gstrSQL & ",'" & Left(Me.cboPage.Text, Me.txt编号.MaxLength) & "','" & Trim(Mid(Me.cboPage.Text, Me.txt编号.MaxLength + 2)) & "'"
    End If
    If Val(Left(Me.cboKind.Text, 1)) <> 3 Then
        gstrSQL = gstrSQL & ",0"
    Else
        gstrSQL = gstrSQL & "," & Me.cbo等级.ListIndex
    End If
    
    If Me.Tag = "增加" Then
        If mblnSpecialWave = False Then '不是专科体温单
            If optEditType(1).Value Then
                intType = 2
            Else
                intType = IIf(chk门诊快捷病历.Visible And chk门诊快捷病历.Value = 1, 3, 0)
            End If
            gstrSQL = "Zl_病历文件列表_Insert(" & gstrSQL & "," & IIf(Me.chkCopy.Value = vbChecked, Val(Me.chkCopy.Tag), 0) & "," & intType & ")"
        Else
            intType = -1
            gstrSQL = "Zl_病历文件列表_Insert(" & gstrSQL & "," & IIf(Me.chkCopy.Value = vbChecked, Val(Me.chkCopy.Tag), 0) & "," & intType & ",'1')"
        End If
    Else
        gstrSQL = "Zl_病历文件列表_Modify(" & gstrSQL & ")"
    End If
    
    Err = 0: On Error GoTo errHand
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    mblnOK = True: Me.Hide: Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Me.txt编号.SetFocus
End Sub

Private Sub optEditType_Click(Index As Integer)
    If Index = 1 Then
        optPage(0).Value = True: optPage(1).Value = False
        optPage(1).Enabled = False: cboPage.Enabled = False: chkCopy.Enabled = False
        chk门诊快捷病历.Enabled = False
    Else
        If Val(cboKind.Text) = 2 Or Val(cboKind.Text) = 4 Then
            optPage(1).Enabled = True: cboPage.Enabled = True
        End If
        chkCopy.Enabled = chkCopy.Tag <> "" And Val(cboKind.Tag) = Val(cboKind.Text) And Not mblnSpecialWave
        chk门诊快捷病历.Enabled = True
    End If
End Sub

Private Sub optPage_Click(Index As Integer)
    If Me.optPage(0).Value Then
        Me.txtPageName.Enabled = True: Me.cboPage.Enabled = False
        Me.txtPageNo.Text = Me.txt编号.Text: Me.txtPageName.Text = Me.txt名称.Text
        If Me.txtPageName.Visible Then Me.txtPageName.SetFocus
    Else
        Me.txtPageName.Enabled = False: Me.cboPage.Enabled = True
        If Me.cboPage.Visible Then Me.cboPage.SetFocus
    End If
End Sub

Private Sub optPage_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtPageName_Change()
    ValidControlText txtPageName
End Sub

Private Sub txtPageName_GotFocus()
    Me.txtPageName.SelStart = 0: Me.txtPageName.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtPageName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt编号_Change()
    ValidControlText txt编号
    Me.txtPageNo.Text = Me.txt编号.Text
End Sub

Private Sub txt编号_GotFocus()
    Me.txt编号.SelStart = 0: Me.txt编号.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt编号_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt名称_Change()
    ValidControlText txt名称
    Me.txtPageName.Text = Me.txt名称.Text
End Sub

Private Sub txt名称_GotFocus()
    Me.txt名称.SelStart = 0: Me.txt名称.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt名称_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt说明_Change()
    ValidControlText txt说明
End Sub

Private Sub txt说明_GotFocus()
    Me.txt说明.SelStart = 0: Me.txt说明.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt说明_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt说明_LostFocus()
    Me.txt说明.Text = Replace(Me.txt说明, Chr(vbKeyReturn), "")
    Call zlCommFun.OpenIme(False)
End Sub
