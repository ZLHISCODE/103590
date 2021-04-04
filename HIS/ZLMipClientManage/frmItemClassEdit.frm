VERSION 5.00
Begin VB.Form frmItemClassEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "#"
   ClientHeight    =   2250
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5775
   Icon            =   "frmItemClassEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fra 
      Height          =   1605
      Left            =   30
      TabIndex        =   3
      Top             =   0
      Width           =   4380
      Begin VB.CommandButton cmd 
         Height          =   300
         Index           =   0
         Left            =   3930
         Picture         =   "frmItemClassEdit.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1035
         Width           =   300
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   960
         TabIndex        =   6
         Top             =   645
         Width           =   3255
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   0
         Left            =   1200
         TabIndex        =   5
         Top             =   300
         Width           =   2100
      End
      Begin VB.TextBox txt 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1035
         Width           =   2940
      End
      Begin VB.TextBox txtParentCode 
         Enabled         =   0   'False
         Height          =   300
         Left            =   960
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "上级(&S)"
         Height          =   180
         Index           =   2
         Left            =   285
         TabIndex        =   11
         Top             =   1110
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "名称(&N)"
         Height          =   180
         Index           =   1
         Left            =   285
         TabIndex        =   10
         Top             =   720
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "编码(&B)"
         Height          =   180
         Index           =   0
         Left            =   285
         TabIndex        =   9
         Top             =   330
         Width           =   630
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4530
      TabIndex        =   2
      Top             =   600
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4530
      TabIndex        =   1
      Top             =   120
      Width           =   1100
   End
   Begin VB.CheckBox chk 
      Caption         =   "允许更改编码长度，并按此调整各同级编码(&L)"
      Height          =   285
      Index           =   0
      Left            =   150
      TabIndex        =   0
      Top             =   1770
      Width           =   4095
   End
End
Attribute VB_Name = "frmItemClassEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'Option Explicit
'
''######################################################################################################################
'Private mstrDataKey As String
'Private mstrUpDataKey As String
'Private mblnAllType As Boolean
'Private mblnOK As Boolean
'Private mfrmMain As Object
'Private mbytMode As Byte
'Private mblnDataChanged As Boolean
'Private mlngSvrMaxLen As Long
'Private mblnContiune As Boolean
'Private mrsPara As ADODB.Recordset
'Private mstrDataCode As String
'
'Public Event AfterNewData(ByVal DataKey As String)
'Public Event AfterModifyData(ByVal DataKey As String)
'Public Event AfterDeleteData(ByVal DataKey As String)
'
'
''######################################################################################################################
'Public Function InitDialog(ByVal frmParent As Object) As Boolean
'    '******************************************************************************************************************
'    '功能：
'    '参数：
'    '返回：
'    '******************************************************************************************************************
'    Set mfrmMain = frmParent
'
'    InitDialog = True
'
'End Function
'
'Public Sub NewData(ByVal strDataCode As String, ByVal strUpDataKey As String)
'    '******************************************************************************************************************
'    '功能：
'    '参数：
'    '返回：
'    '******************************************************************************************************************
'    mbytMode = 1
'    mstrDataKey = ""
'    mstrDataCode = strDataCode
'    mstrUpDataKey = strUpDataKey
'
'    Me.Caption = "新增消息分类"
'
'    Call ExecuteCommand("初始数据")
'    Call ExecuteCommand("缺省编码")
'
'
'    Call AdjustCodePostion(Me, txtParentCode, txt(0))
'    mblnDataChanged = False
'
'    Me.Show 1, mfrmMain
'
'End Sub
'
'Public Sub ModifyData(ByVal strDataCode As String, ByVal strDataKey As String)
'    '******************************************************************************************************************
'    '功能：
'    '参数：
'    '返回：
'    '******************************************************************************************************************
'
'    mbytMode = 2
'    mstrDataKey = strDataKey
'    mstrDataCode = strDataCode
'    Me.Caption = "修改消息分类息"
'
'    Call ExecuteCommand("初始数据")
'    Call ExecuteCommand("读取分类")
'    Call AdjustCodePostion(Me, txtParentCode, txt(0))
'    mblnDataChanged = False
'
'    Me.Show 1, mfrmMain
'
'End Sub
'
'Public Sub DeleteData(ByVal strDataCode As String, ByVal strDataKey As String)
'    '******************************************************************************************************************
'    '功能：
'    '参数：
'    '返回：
'    '******************************************************************************************************************
'    mbytMode = 3
'    mstrDataCode = strDataCode
'    If strDataKey = "" Then Exit Sub
'    mstrDataKey = strDataKey
'
'    Set mrsPara = zlCommFun.CreateParameter
'    Call zlCommFun.SetParameter(mrsPara, "ID", mstrDataKey)
'
'    If gclsBusiness.ItemClassEdit("Delete", mrsPara) Then
'        RaiseEvent AfterDeleteData(mstrDataKey)
'    End If
'End Sub
'
''######################################################################################################################
'Private Function NewDefaultCode(ByVal strUpKey As String, ByRef objTxtParent As Object, ByRef objTxt As Object, ByRef objChk As Object) As Boolean
'    '******************************************************************************************************************
'    '功能：产生缺省编码
'    '参数：
'    '返回：
'    '******************************************************************************************************************
'    Dim rs As New ADODB.Recordset
'    Dim intMaxLength As Integer
'    Dim str最大编码 As String
'    Dim str上级编码 As String
'    Dim blnMsg As Boolean '是否提示
'    Dim rsCondition As ADODB.Recordset
'
'    Set rsCondition = zlCommFun.CreateCondition
'
'    '读取上级编码
'
'    If strUpKey = "" Then
'        str上级编码 = ""
'
'        Call zlCommFun.SetCondition(rsCondition, "id", "-1")
'        Set rs = gclsBusiness.ItemClassRead("ID", rsCondition)
'
'    Else
'        Call zlCommFun.SetCondition(rsCondition, "id", strUpKey)
'        Set rs = gclsBusiness.ItemClassRead("ID", rsCondition)
'
'        If rs.BOF = False Then
'            str上级编码 = zlCommFun.NVL(rs("编码").Value)
'        End If
'    End If
'
'    intMaxLength = rs.Fields("编码").DefinedSize
'
'    Set rs = gclsBusiness.ItemClassMaxCode(strUpKey)
'    If rs.BOF = False Then
'        str最大编码 = Trim(zlCommFun.NVL(rs("最大编码").Value))
'    End If
'
'    If mblnAllType = False Then
'        blnMsg = False
'        Set rs = gclsBusiness.GetClassDefaultCode(str上级编码, str最大编码, intMaxLength, blnMsg)
'        If blnMsg = False Then
'            If rs.BOF = False Then
'                objTxtParent.Text = zlCommFun.NVL(rs("上级编码").Value)
'                objChk.Value = zlCommFun.NVL(rs("调整编码").Value, 0)
'                objTxt.Text = zlCommFun.NVL(rs("缺省编码").Value)
'                objTxt.MaxLength = zlCommFun.NVL(rs("允许编码长度").Value, 0)
'                objTxt.Tag = zlCommFun.NVL(rs("允许编码最大长度").Value)
'                objChk.Enabled = (zlCommFun.NVL(rs("允许调整").Value, 0) = 1)
'            End If
'        Else
'            objTxtParent.Text = ""
'            objChk.Value = 0
'            objTxt.Text = ""
'            objTxt.MaxLength = Len(str最大编码)
'            objTxt.Tag = intMaxLength
'            objChk.Enabled = True
'        End If
'    Else
'        objTxtParent.Text = ""
'        objChk.Value = 0
'        objTxt.Text = ""
'        objTxt.MaxLength = Len(str最大编码)
'        objTxt.Tag = intMaxLength
'        objChk.Enabled = True
'    End If
'
'    NewDefaultCode = True
'End Function
'
'Private Function AnalyzeCode(ByVal strKey As String, ByRef objTxtParent As Object, ByRef objTxt As Object) As Boolean
'    '******************************************************************************************************************
'    '功能：分解编码
'    '参数：
'    '返回：
'    '******************************************************************************************************************
'    Dim rs As New ADODB.Recordset
'    Dim rsCondition As ADODB.Recordset
'
'    Set rsCondition = zlCommFun.CreateCondition
'
'    Call zlCommFun.SetCondition(rsCondition, "id", strKey)
'    Set rs = gclsBusiness.ItemClassRead("ID", rsCondition)
'
'    If rs.BOF Then Exit Function
'
'    objTxtParent.Text = zlCommFun.NVL(rs("上级编码").Value)
'    objTxt.Text = zlCommFun.NVL(rs("编码").Value)
'
'    If Len(objTxt.Text) >= Len(objTxtParent.Text) Then objTxt.Text = Mid(objTxt.Text, Len(objTxtParent.Text) + 1)
'
'    objTxt.MaxLength = Len(objTxt.Text)
'    objTxt.Tag = rs.Fields("编码").DefinedSize - Len(objTxtParent.Text)
'
'    AnalyzeCode = True
'End Function
'
'Private Function ExecuteCommand(ByVal strCommand As String, ParamArray varPara() As Variant) As Boolean
'    '--------------------------------------------------------------------------------------------------------------
'    '功能：
'    '参数：
'    '返回：
'    '--------------------------------------------------------------------------------------------------------------
'    Dim intLoop As Integer
'    Dim rs As New ADODB.Recordset
'    Dim rsSQL As New ADODB.Recordset
'    Dim rsCondition As ADODB.Recordset
'
'    On Error GoTo errHand
'
'    Call SQLRecord(rsSQL)
'
'    Set rsCondition = zlCommFun.CreateCondition
'
'
'    Select Case strCommand
'    '--------------------------------------------------------------------------------------------------------------
'    Case "初始数据"
'
'        '设置最大输入长度
'        Set rs = gclsBusiness.ItemClassStruct()
'        If Not (rs Is Nothing) Then
'            txtParentCode.MaxLength = rs("folder_code").DefinedSize
'            txt(1).MaxLength = rs("folder_name").Precision
'        End If
'
'        '读取上级分类
'        If mstrUpDataKey <> "" Then
'            Call zlCommFun.SetCondition(rsCondition, "id", mstrUpDataKey)
'
'            Set rs = gclsBusiness.ItemClassRead("ID", rsCondition)
'            If rs.BOF = False Then
'
'                txt(2).Text = AppendCode(zlCommFun.NVL(rs("名称").Value), zlCommFun.NVL(rs("编码").Value))
'                cmd(0).Tag = zlCommFun.NVL(rs("ID").Value, 0)
'                mstrUpDataKey = Trim(cmd(0).Tag)
'
'            End If
'        End If
'
'    '--------------------------------------------------------------------------------------------------------------
'    Case "清空数据"
'
'        Call ExecuteCommand("缺省编码")
'
'        txt(1).Text = ""
'        txt(0).SetFocus
'
'        mblnDataChanged = False
'
'    '--------------------------------------------------------------------------------------------------------------
'    Case "缺省编码"
'
'        Call NewDefaultCode(mstrUpDataKey, txtParentCode, txt(0), chk(0))
'
'    '--------------------------------------------------------------------------------------------------------------
'    Case "读取分类"
'
'        Call zlCommFun.SetCondition(rsCondition, "id", mstrDataKey)
'
'        Set rs = gclsBusiness.ItemClassRead("ID", rsCondition)
'        If rs.BOF = False Then
'
'            txt(1).Text = zlCommFun.NVL(rs("名称").Value)
'
'
'            txt(2).Text = AppendCode(zlCommFun.NVL(rs("上级名称").Value), zlCommFun.NVL(rs("上级编码").Value))
'
'            cmd(0).Tag = zlCommFun.NVL(rs("上级id").Value, 0)
'
'            Call AnalyzeCode(mstrDataKey, txtParentCode, txt(0))
'
'        End If
'    End Select
'
'    ExecuteCommand = True
'
'    Exit Function
'errHand:
'
'    If ErrCenter = 1 Then
'        Resume
'    End If
'    Call SaveErrLog
'
'End Function
'
'Private Function ValidData() As Boolean
'    '******************************************************************************************************************
'    '功能：校验编辑数据的有效性
'    '参数：
'    '返回：
'    '******************************************************************************************************************
'
'    If txt(0).MaxLength = 0 Then
'        ShowSimpleMsg "上级编码已经达到最大长度，不能设置下级！"
'        cmdCancel.SetFocus
'        Exit Function
'    End If
'
'    If chk(0).Value = 0 And Len(Trim(txt(0).Text)) <> txt(0).MaxLength Then
'        ShowSimpleMsg "编码长度必须为" & txt(0).MaxLength & "位，除非你选择更改长度选项"
'        LocationObj txt(0)
'        Exit Function
'    End If
'
'    If Trim(txt(0).Text) = "" Then
'        ShowSimpleMsg "编码不能为空值，必须输入！"
'        LocationObj txt(0)
'        Exit Function
'    End If
'
'    '检查编码是否为数字字符
'    If CheckStrType(Trim(txt(0).Text), 99, "0123456789") = False Then
'        ShowSimpleMsg "编码必须为数字字符！"
'        LocationObj txt(0)
'        Exit Function
'    End If
'
'    If Trim(txt(1).Text) = "" Then
'        ShowSimpleMsg "名称不能为空值，必须输入！"
'        LocationObj txt(1)
'        Exit Function
'    End If
'
'
'    ValidData = True
'
'End Function
'
'
'Private Function SaveData(ByRef strDataKey As String) As Boolean
'    '******************************************************************************************************************
'    '功能：保存编辑数据到数据库
'    '参数：
'    '返回：
'    '******************************************************************************************************************
'    Dim rsPara As New ADODB.Recordset
'
'    On Error GoTo errHand
'
'    Set rsPara = zlCommFun.CreateParameter
'
'    Call zlCommFun.SetParameter(rsPara, "id", strDataKey)
'    Call zlCommFun.SetParameter(rsPara, "data_code", mstrDataCode)
'    Call zlCommFun.SetParameter(rsPara, "parent_id", Trim(cmd(0).Tag))
'    Call zlCommFun.SetParameter(rsPara, "folder_code", Trim(txtParentCode.Text & txt(0).Text))
'    Call zlCommFun.SetParameter(rsPara, "folder_name", Trim(txt(1).Text))
'    Call zlCommFun.SetParameter(rsPara, "adjustcodelength", chk(0).Value)
'
'    Select Case mbytMode
'    '------------------------------------------------------------------------------------------------------------------
'    Case 1          '新增
'        strDataKey = zlCommFun.GetGUID
'        Call zlCommFun.SetParameter(rsPara, "id", strDataKey)
'
'        SaveData = gclsBusiness.ItemClassEdit("INSERT", rsPara)
'    '------------------------------------------------------------------------------------------------------------------
'    Case 2          '修改
'        SaveData = gclsBusiness.ItemClassEdit("UPDATE", rsPara)
'    End Select
'
'    Exit Function
'
'errHand:
'
'    If ErrCenter = 1 Then
'        Resume
'    End If
'End Function
'
'Private Sub chk_Click(Index As Integer)
'    If chk(Index).Value = 1 Then
'        mlngSvrMaxLen = txt(0).MaxLength
'        txt(0).MaxLength = Val(txt(0).Tag)
'    Else
'        txt(0).MaxLength = mlngSvrMaxLen
'        txt(0).Text = Mid(txt(0).Text, 1, txt(0).MaxLength)
'    End If
'
'    mblnDataChanged = True
'
'    On Error Resume Next
'    txt(0).SetFocus
'End Sub
'
'Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'        KeyAscii = 0
'        zlCommFun.PressKey vbKeyTab
'    Else
'        If Chr(KeyAscii) = "'" Then KeyAscii = 0
'    End If
'End Sub
'
'Private Sub cmd_Click(Index As Integer)
''    Dim rsData As New ADODB.Recordset
''    Dim rs As New ADODB.Recordset
''
''    Select Case mstrTemplate
''    Case "消息项目分类"
''        Set rsData = gclsPackage.Get_Elementclasstreesel(mlngKey)
''    End Select
''
''    If gclsBase.ShowPubSelect(Me, txt(2), 1, "", Me.Name & "\" & mstrTemplate & "选择", "请从下面选择一个" & mstrTemplate, rsData, rs, cmd(0).Left + cmd(0).Width - txt(2).Left, 3900, , mlngKey, , False) = 1 Then
''        If Val(cmd(0).Tag) <> zlCommFun.NVL(rs("ID")) Then
''            If zlCommFun.NVL(rs("ID")) = -1 Then
''                txt(2).Text = ""
''                cmd(0).Tag = ""
''                mblnAllType = True
''            Else
''                txt(2).Text = zlCommFun.NVL(rs("名称"))
''                cmd(0).Tag = zlCommFun.NVL(rs("ID"))
''                mblnAllType = False
''            End If
''
''            mlngUpKey = Val(cmd(0).Tag)
''
''            Call ExecuteCommand("缺省编码")
''            DataChanged = True
''            mblnAllType = False
''            Call AdjustCodePostion(Me, txtParentCode, txt(0))
''        End If
''    End If
'End Sub
'
'Private Sub cmdCancel_Click()
'    Unload Me
'End Sub
'
'Private Sub cmdOK_Click()
'
'    Dim strOldDataKey As String
'    Dim rsTmp As ADODB.Recordset
'
'    If mblnDataChanged = True Then
'        If ValidData = True Then
'
'            If SaveData(mstrDataKey) = True Then
'
'                Select Case mbytMode
'                Case 1
'                    RaiseEvent AfterNewData(mstrDataKey)
'                Case 2
'                    RaiseEvent AfterModifyData(mstrDataKey)
'                End Select
'
'                If mblnContiune = False Then
'                    mblnDataChanged = False
'                    Unload Me
'                Else
'                    '重置环境，进入下一次新增状态
'                    If mbytMode = 1 Then
'                        mstrDataKey = ""
'                        Call ExecuteCommand("缺省编码")
'                        txt(1).Text = ""
'                    End If
'                    Call LocationObj(txt(0))
'                    mblnDataChanged = False
'                End If
'            End If
'        End If
'    Else
'        Unload Me
'    End If
'
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'    If mblnDataChanged Then
'        Cancel = (MsgBox("新增或修改的数据必须保存后才生效，是否不保存就退出？", vbYesNo + vbQuestion + vbDefaultButton2, ParamInfo.系统名称) = vbNo)
'    End If
'End Sub
'
'Private Sub txt_Change(Index As Integer)
'    mblnDataChanged = True
'End Sub
'
'Private Sub txt_GotFocus(Index As Integer)
'
'    zlControl.TxtSelAll txt(Index)
'
'    Select Case Index
'    Case 1, 3
'        zlCommFun.OpenIme True
'    End Select
'
'End Sub
'
'Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyDelete Then
'        txt(2).Text = ""
'        cmd(0).Tag = ""
'
'        mstrUpDataKey = ""
'        Call ExecuteCommand("缺省编码")
'        mblnDataChanged = True
'
'        Call AdjustCodePostion(Me, txtParentCode, txt(0))
'    End If
'End Sub
'
'Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then
'        KeyAscii = 0
'        zlCommFun.PressKey vbKeyTab
'    Else
'        If Chr(KeyAscii) = "'" Then KeyAscii = 0
'        If Index = 0 Then
'            If FilterKeyAscii(KeyAscii, 99, "0123456789") = 0 Then KeyAscii = 0
'        End If
'    End If
'End Sub
'
'Private Sub txt_LostFocus(Index As Integer)
'    Select Case Index
'    Case 1, 3
'        zlCommFun.OpenIme False
'    End Select
'End Sub
'
'Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Button = 2 And txt(Index).Locked Then
'        glngTXTProc = GetWindowLong(txt(Index).hWnd, GWL_WNDPROC)
'        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
'    End If
'End Sub
'
'Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Button = 2 And txt(Index).Locked Then
'        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, glngTXTProc)
'    End If
'End Sub
'
'Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
'    Cancel = Not zlCommFun.StrIsValid(txt(Index).Text, txt(Index).MaxLength)
'End Sub
'
'Private Function AdjustCodePostion(ByVal frmMain As Object, ByRef objTxtParent As Object, ByRef objTxt As Object) As Boolean
'    '******************************************************************************************************************
'    '功能：
'    '参数：
'    '返回：
'    '******************************************************************************************************************
'    objTxt.Top = objTxtParent.Top + 45
'    objTxt.Left = objTxtParent.Left + frmMain.TextWidth(objTxtParent.Text) + 60
'    objTxt.Width = objTxtParent.Width - frmMain.TextWidth(objTxtParent.Text) - 120
'    objTxt.BackColor = objTxtParent.BackColor
'
'    AdjustCodePostion = True
'
'End Function
'
