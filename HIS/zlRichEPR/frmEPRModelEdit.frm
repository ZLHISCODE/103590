VERSION 5.00
Begin VB.Form frmEPRModelEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "示范编辑"
   ClientHeight    =   5055
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5100
   Icon            =   "frmEPRModelEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cbo 
      Height          =   300
      Index           =   0
      Left            =   1155
      TabIndex        =   25
      Top             =   2655
      Width           =   3660
   End
   Begin VB.TextBox txt简码 
      Height          =   300
      Left            =   1155
      MaxLength       =   10
      TabIndex        =   9
      Top             =   2295
      Width           =   3660
   End
   Begin VB.TextBox txt说明 
      Height          =   660
      Left            =   1155
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   3030
      Width           =   3660
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   -435
      TabIndex        =   21
      Top             =   4560
      Width           =   5760
   End
   Begin VB.Frame fraLine 
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   0  'None
      Height          =   1515
      Index           =   0
      Left            =   -15
      TabIndex        =   0
      Top             =   0
      Width           =   5910
      Begin VB.OptionButton opt性质 
         BackColor       =   &H00FDD6C6&
         Caption         =   "表格式范文(&T): 此种范文用表格式病历编辑器编辑"
         Height          =   225
         Index           =   2
         Left            =   435
         TabIndex        =   26
         Top             =   1200
         Width           =   4455
      End
      Begin VB.OptionButton opt性质 
         BackColor       =   &H00FDD6C6&
         Caption         =   "片段(&S): "
         Height          =   180
         Index           =   1
         Left            =   435
         TabIndex        =   3
         Top             =   765
         Width           =   1020
      End
      Begin VB.OptionButton opt性质 
         BackColor       =   &H00FDD6C6&
         Caption         =   "范文(&M):"
         Height          =   180
         Index           =   0
         Left            =   435
         TabIndex        =   2
         Top             =   345
         Value           =   -1  'True
         Width           =   1020
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "示范性质"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   480
         TabIndex        =   1
         Top             =   30
         Width           =   720
      End
      Begin VB.Image imgNote 
         Height          =   240
         Left            =   150
         Picture         =   "frmEPRModelEdit.frx":058A
         Top             =   15
         Width           =   240
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "包含文件部分提纲内容的示范, 病历编辑时可叠加选择多个片段."
         Height          =   360
         Index           =   1
         Left            =   1485
         TabIndex        =   23
         Top             =   750
         Width           =   3420
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "包含完整文件格式和内容的示范文档, 病历编辑时选用一个范文并将覆盖此前内容;"
         Height          =   360
         Index           =   0
         Left            =   1485
         TabIndex        =   22
         Top             =   330
         Width           =   3420
         WordWrap        =   -1  'True
      End
   End
   Begin VB.ComboBox cbo科室 
      Height          =   300
      Left            =   1155
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   4110
      Width           =   2370
   End
   Begin VB.OptionButton opt范围 
      Caption         =   "&3)个人使用"
      Height          =   180
      Index           =   2
      Left            =   3675
      TabIndex        =   15
      Top             =   3780
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton opt范围 
      Caption         =   "&2)科内通用"
      Height          =   180
      Index           =   1
      Left            =   2385
      TabIndex        =   14
      Top             =   3780
      Width           =   1215
   End
   Begin VB.OptionButton opt范围 
      Caption         =   "&1)全院通用"
      Height          =   180
      Index           =   0
      Left            =   1140
      TabIndex        =   13
      Top             =   3780
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2265
      TabIndex        =   19
      Top             =   4665
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3645
      TabIndex        =   20
      Top             =   4665
      Width           =   1215
   End
   Begin VB.TextBox txt名称 
      Height          =   300
      Left            =   1155
      TabIndex        =   7
      Top             =   1935
      Width           =   3660
   End
   Begin VB.TextBox txt编号 
      Height          =   300
      Left            =   1155
      TabIndex        =   5
      Top             =   1575
      Width           =   3660
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "分类(&F)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   435
      TabIndex        =   24
      Top             =   2700
      Width           =   630
   End
   Begin VB.Label lbl简码 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "简码(&D)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   435
      TabIndex        =   8
      Top             =   2355
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
      Left            =   435
      TabIndex        =   10
      Top             =   3075
      Width           =   630
   End
   Begin VB.Label lbl科室 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "制作(&R)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   435
      TabIndex        =   16
      Top             =   4170
      Width           =   630
   End
   Begin VB.Label lbl人员 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   3600
      TabIndex        =   18
      Top             =   4110
      Width           =   1230
   End
   Begin VB.Label lbl范围 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "使用(&U)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   435
      TabIndex        =   12
      Top             =   3780
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
      Left            =   435
      TabIndex        =   6
      Top             =   1980
      Width           =   630
   End
   Begin VB.Label lbl编号 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "编号(&B)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   435
      TabIndex        =   4
      Top             =   1635
      Width           =   630
   End
End
Attribute VB_Name = "frmEPRModelEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------
'说明：
'   1、上级程序通过本窗体ShowMe函数，将父窗体、编辑范文ID,编辑状态等信息传递进入本程序
'   2、编辑状态：由Me.tag存放，分别为"新增"、"修改"，由上级程序通过ShowMe传入
'---------------------------------------------------
Private mlngFileId As Long       '提纲ID
Private mlngRecID As Long       '记录ID
Private mblnOK As Boolean        '是否完成编辑退出
 
Public Function ShowMe(ByVal frmParent As Object, ByVal blnAdd As Boolean, ByVal bytPower As Byte, ByVal lngFileId As Long _
                    , Optional ByVal lngRecId As Long, Optional ByVal EditType As Byte) As Long
    '---------------------------------------------------
    '功能：上级程序调用本窗体的，传递参数，并显示窗体
    '参数：bytPower-管理权限（=0，全院；=1，科室；=2，个人）；lngFileId-提纲ID；lngRecID-记录ID;EditType=0自定义病历 EditType＝1系统自带 EditType=2 表格式病历
    '返回：确定返回新增或修改的ID；取消返回0
    '---------------------------------------------------
Dim rsTemp As New ADODB.Recordset
Dim lngCount As Long
    mlngFileId = lngFileId: mlngRecID = lngRecId
    If blnAdd Then
        Me.Tag = "新增": mlngRecID = 0
    Else
        Me.Tag = "修改"
    End If
    
    '基本数据信息
    '------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select Distinct D.ID, D.编码, D.名称, R.缺省, R.人员id, P.姓名" & vbNewLine & _
            "From 部门表 D, 部门人员 R, 人员表 P, 上机人员表 U, 部门性质说明 C," & vbNewLine & _
            "     (Select 种类, 通用 From 病历文件列表 Where ID = [1]) L" & vbNewLine & _
            "Where D.ID = R.部门id And R.人员id = P.ID And P.ID = U.人员id And U.用户名 = User And D.ID = C.部门id And" & vbNewLine & _
            "      C.工作性质 In ('临床', '检查', '检验', '手术', '治疗', '护理', '营养', '体检') And" & vbNewLine & _
            "      (Nvl(L.通用, 0) <> 2 Or L.种类 = 7 Or" & vbNewLine & _
            "      L.种类 <> 7 And L.通用 = 2 And D.ID In (Select 科室id From 病历应用科室 Where 文件id = [1]))" & vbNewLine & _
            "Order By R.缺省 Desc, D.编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngFileId)
    With rsTemp
        Do While Not .EOF
            Me.cbo科室.AddItem !编码 & "-" & !名称
            Me.cbo科室.ItemData(Me.cbo科室.NewIndex) = !ID
            If !缺省 = 1 Then Me.cbo科室.ListIndex = Me.cbo科室.NewIndex
            Me.lbl人员.Tag = !人员ID: Me.lbl人员.Caption = !姓名
            .MoveNext
        Loop
        If Me.cbo科室.ListCount = 0 Then
            MsgBox "你目前不属于该病历应用科室范围，不能管理范文！", vbExclamation, gstrSysName
            ShowMe = 0: Unload Me: Exit Function
        ElseIf Me.cbo科室.ListIndex = -1 Then
            Me.cbo科室.ListIndex = 0
        End If
    End With
    
    cbo(0).Clear
    cbo(0).AddItem ""
    gstrSQL = "Select Distinct a.分类 From 病历范文目录 a Where a.文件id =[1] And a.分类 Is Not Null"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngFileId)
    If rsTemp.BOF = False Then
        Do While Not rsTemp.EOF
            cbo(0).AddItem rsTemp("分类").Value
            rsTemp.MoveNext
        Loop
    End If
    cbo(0).ListIndex = 0
    
    If blnAdd Then
        If EditType = 2 Then
            opt性质(2).Value = True: fraLine(0).Enabled = False: opt性质(0).Enabled = False: opt性质(1).Enabled = False
        Else
            opt性质(2).Enabled = False
        End If
    End If
    '内容数据提取
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select l.编号, l.名称, l.简码, l.分类, l.性质, l.说明, l.通用级, l.科室id, d.编码, d.名称 As 部门, l.人员id, p.姓名 As 人员" & _
            " From 病历范文目录 l, 部门表 d, 人员表 p" & _
            " Where l.科室id = d.Id And l.人员id = p.Id And l.id =[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngRecID)
    With rsTemp
        If .RecordCount > 0 Then
            opt性质(NVL(!性质, 0)).Value = True
            Me.fraLine(0).Enabled = False
            Me.txt编号.Text = !编号
            Me.txt名称.Text = "" & !名称
            Me.txt简码.Text = "" & !简码
            Me.txt说明.Text = "" & !说明
            Me.opt范围(IIf(IsNull(!通用级), 0, !通用级)).Value = True
            If !人员ID <> Me.lbl人员.Tag Then
                Me.lbl人员.Tag = !人员ID: Me.lbl人员.Caption = !人员
                Me.cbo科室.Clear
                Me.cbo科室.AddItem !编码 & "-" & !部门
                Me.cbo科室.ItemData(Me.cbo科室.NewIndex) = !科室ID
                Me.cbo科室.ListIndex = Me.cbo科室.NewIndex
                Me.cbo科室.Enabled = False
            Else
                For lngCount = 0 To Me.cbo科室.ListCount - 1
                    If Me.cbo科室.ItemData(lngCount) = IIf(IsNull(!科室ID), 0, !科室ID) Then
                        Me.cbo科室.ListIndex = lngCount: Exit For
                    End If
                Next
            End If
            cbo(0).Text = zlCommFun.NVL(!分类)
        End If
        Me.txt编号.MaxLength = .Fields("编号").DefinedSize
        Me.txt名称.MaxLength = .Fields("名称").DefinedSize
        Me.txt简码.MaxLength = .Fields("简码").DefinedSize
        Me.txt说明.MaxLength = .Fields("说明").DefinedSize
    End With
    Select Case bytPower
    Case 2: Me.opt范围(0).Enabled = False: Me.opt范围(1).Enabled = False
    Case 1: Me.opt范围(0).Enabled = False
    End Select
    
    If Me.Tag = "新增" Then
        Me.txt编号.Text = GetMax("病历范文目录", "编号", 5, " Where 文件id=" & mlngFileId)
    End If
    
    '显示窗体
    Me.Show vbModal, frmParent
    If mblnOK Then
        ShowMe = mlngRecID
    Else
        ShowMe = 0
    End If
    Unload Me
    Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    ShowMe = 0
End Function

Private Sub cbo_GotFocus(Index As Integer)
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub cbo_LostFocus(Index As Integer)
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbo_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(cbo(Index).Text, 50)
End Sub

Private Sub cbo科室_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False: Me.Hide
End Sub

Private Sub cmdOK_Click()

    If Trim(Me.txt编号.Text) = "" Then MsgBox "请输入编号！", vbInformation, gstrSysName: Me.txt编号.SetFocus: Exit Sub
    
    If Trim(Me.txt名称.Text) = "" Then MsgBox "请输入名称！", vbInformation, gstrSysName: Me.txt名称.SetFocus: Exit Sub
    
    If LenB(StrConv(Trim(Me.txt名称.Text), vbFromUnicode)) > Me.txt名称.MaxLength Then
        MsgBox "名称超长（最多" & Me.txt名称.MaxLength & "个字符或等长的汉字）！", vbInformation, gstrSysName: Me.txt名称.SetFocus: Exit Sub
    End If
    
    If LenB(StrConv(Trim(Me.txt简码.Text), vbFromUnicode)) > Me.txt简码.MaxLength Then
        MsgBox "简码超长（最多" & Me.txt简码.MaxLength & "个字符或等长的汉字）！", vbInformation, gstrSysName: Me.txt简码.SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txt说明.Text), vbFromUnicode)) > Me.txt说明.MaxLength Then
        MsgBox "说明超长（最多" & Me.txt说明.MaxLength & "个字符或等长的汉字）！", vbInformation, gstrSysName: Me.txt说明.SetFocus: Exit Sub
    End If
    
    If Me.cbo科室.ListIndex = -1 Then MsgBox "请输入科室！", vbInformation, gstrSysName: Me.cbo科室.SetFocus: Exit Sub
    
    '数据保存
    If Me.Tag = "新增" Then
        mlngRecID = zlDatabase.GetNextId("病历范文目录")
        gstrSQL = mlngRecID & "," & mlngFileId & ",'" & Trim(Me.txt编号.Text) & "','" & Trim(Me.txt名称.Text) & "','" & Trim(Me.txt简码.Text) & "'"
        gstrSQL = gstrSQL & "," & IIf(Me.opt性质(0).Value, 0, IIf(opt性质(1).Value, 1, 2)) & ",'" & Replace(Trim(Me.txt说明.Text), Chr(vbKeyReturn), "") & "'"
        If Me.opt范围(0).Value Then
            gstrSQL = gstrSQL & ",0"
        ElseIf Me.opt范围(1).Value Then
            gstrSQL = gstrSQL & ",1"
        Else
            gstrSQL = gstrSQL & ",2"
        End If
        gstrSQL = gstrSQL & "," & Me.cbo科室.ItemData(Me.cbo科室.ListIndex) & "," & Me.lbl人员.Tag & ",'" & cbo(0).Text & "'"
        gstrSQL = "Zl_病历范文目录_Insert(" & gstrSQL & ")"
    Else
        gstrSQL = mlngRecID & ",'" & Trim(Me.txt编号.Text) & "','" & Trim(Me.txt名称.Text) & "','" & Trim(Me.txt简码.Text) & "'"
        gstrSQL = gstrSQL & ",'" & Replace(Trim(Me.txt说明.Text), Chr(vbKeyReturn), "") & "'"
        If Me.opt范围(0).Value Then
            gstrSQL = gstrSQL & ",0"
        ElseIf Me.opt范围(1).Value Then
            gstrSQL = gstrSQL & ",1"
        Else
            gstrSQL = gstrSQL & ",2"
        End If
        gstrSQL = gstrSQL & "," & Me.cbo科室.ItemData(Me.cbo科室.ListIndex) & ",'" & cbo(0).Text & "'"
        gstrSQL = "Zl_病历范文目录_Update(" & gstrSQL & ")"
    End If
    
    Err = 0: On Error GoTo errHand
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    mblnOK = True: Me.Hide
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub optEditType_Click(Index As Integer)
    
End Sub

Private Sub opt范围_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub opt性质_GotFocus(Index As Integer)
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub opt性质_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txt编号_Change()
'    txt编号 = Val(txt编号)
End Sub

Private Sub txt编号_GotFocus()
    Me.txt编号.SelStart = 0: Me.txt编号.SelLength = 100
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
        If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z")) Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt编号_Validate(Cancel As Boolean)
    Cancel = Not StrIsValid(txt编号.Text, txt编号.MaxLength)
End Sub

Private Sub txt简码_GotFocus()
    Me.txt简码.SelStart = 0: Me.txt简码.SelLength = 4000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt简码_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt简码_Validate(Cancel As Boolean)
    Cancel = Not StrIsValid(txt简码.Text, txt简码.MaxLength)
End Sub

Private Sub txt名称_Change()
    ValidControlText txt名称
    Me.txt简码.Text = Left(zlCommFun.SpellCode(Me.txt名称.Text), 10)
End Sub

Private Sub txt名称_GotFocus()
    Me.txt名称.SelStart = 0: Me.txt名称.SelLength = 4000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt名称_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()[]{}_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt名称_Validate(Cancel As Boolean)
    Cancel = Not StrIsValid(txt名称.Text, txt名称.MaxLength)
End Sub

Private Sub txt说明_Change()
    ValidControlText txt说明
End Sub

Private Sub txt说明_GotFocus()
    Me.txt说明.SelStart = 0: Me.txt说明.SelLength = 4000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt说明_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0:  Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr("'%", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt说明_LostFocus()
    Me.txt说明.Text = Replace(Me.txt说明, Chr(vbKeyReturn), "")
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt说明_Validate(Cancel As Boolean)
    Cancel = Not StrIsValid(txt说明.Text, txt说明.MaxLength)
End Sub
