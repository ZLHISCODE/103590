VERSION 5.00
Begin VB.Form frmTemplateClass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "#"
   ClientHeight    =   2250
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5775
   Icon            =   "frmPeisTemplateClass.frx":0000
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
         Picture         =   "frmPeisTemplateClass.frx":000C
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
Attribute VB_Name = "frmTemplateClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'######################################################################################################################
Private mlngKey As Long
Private mlngUpKey As Long
Private mblnAllType As Boolean
Private mblnOK As Boolean
Private mfrmMain As Object
Private mbytMode As Byte
Private mblnChanged As Boolean
Private mlngSvrMaxLen As Long
Private mstrTemplate As String

Public Event AfterSaved(ByVal SaveKey As Long)

'######################################################################################################################

Public Function ShowEdit(ByVal frmMain As Object, ByVal strTemplate As String, ByVal lngKey As Long, ByVal lngUpKey As Long, Optional ByVal bytMode As Byte = 1) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mstrTemplate = strTemplate
    
    mblnOK = False
    mlngKey = lngKey
    mlngUpKey = lngUpKey
    mbytMode = bytMode
    
    Set mfrmMain = frmMain
    Me.Caption = mstrTemplate
    
    If ExecuteCommand("初始数据") = False Then Exit Function

    If mbytMode = 1 Then
        Call ExecuteCommand("缺省编码")
    Else
        If ExecuteCommand("读取分类") = False Then Exit Function
    End If
    
    Call AdjustCodePostion(Me, txtParentCode, txt(0))
    
    DataChanged = False
    
    Me.Show 1, frmMain

    ShowEdit = mblnOK
    
End Function

Private Function NewDefaultCode(ByVal lngUpKey As Long, ByRef objTxtParent As Object, ByRef objTxt As Object, ByRef objChk As Object) As Boolean
    '******************************************************************************************************************
    '功能：产生缺省编码
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim intMaxLength As Integer
    Dim str最大编码 As String
    Dim str上级编码 As String
    Dim blnMsg As Boolean '是否提示
    
    '读取上级编码
        
    If lngUpKey <= 0 Then
        str上级编码 = ""
        
        Select Case mstrTemplate
        Case "体检指标分类"
            Set rs = gclsPackage.Get_Elementclass(-1)
        Case "体检项目分类"
            Set rs = gclsPackage.Get_Medicalclass(-1)
        Case "危害因素分类"
            Set rs = gclsPackage.Get_Virusclass(-1)
        Case "体检套餐分类"
            Set rs = gclsPackage.Get_Packageclass(-1)
        Case "体检诊断分类"
            Set rs = gclsPackage.Get_Diagnoseclass(-1)
        End Select
        
    Else
    
        Select Case mstrTemplate
        Case "体检指标分类"
            Set rs = gclsPackage.Get_Elementclass(lngUpKey)
        Case "体检项目分类"
            Set rs = gclsPackage.Get_Medicalclass(lngUpKey)
        Case "危害因素分类"
            Set rs = gclsPackage.Get_Virusclass(lngUpKey)
        Case "体检套餐分类"
            Set rs = gclsPackage.Get_Packageclass(lngUpKey)
        Case "体检诊断分类"
            Set rs = gclsPackage.Get_Diagnoseclass(lngUpKey)
        End Select
        
        If rs.BOF = False Then
            str上级编码 = zlCommFun.NVL(rs("编码").Value)
        End If
    End If
    
    intMaxLength = rs.Fields("编码").DefinedSize
    
    Select Case mstrTemplate
    Case "体检指标分类"
        Set rs = gclsPackage.Get_Elementmaxcode(IIf(lngUpKey <= 0, 0, lngUpKey), 1)
    Case "体检项目分类"
        Set rs = gclsPackage.Get_Medicalmaxcode(IIf(lngUpKey <= 0, 0, lngUpKey), 1)
    Case "危害因素分类"
        Set rs = gclsPackage.Get_Virusmaxcode(IIf(lngUpKey <= 0, 0, lngUpKey), 1)
    Case "体检套餐分类"
        Set rs = gclsPackage.Get_Packagemaxcode(IIf(lngUpKey <= 0, 0, lngUpKey), 1)
    Case "体检诊断分类"
        Set rs = gclsPackage.Get_Diagnosemaxcode(IIf(lngUpKey <= 0, 0, lngUpKey), 1)
    End Select
    
    If rs.BOF = False Then
        str最大编码 = Trim(zlCommFun.NVL(rs("最大编码").Value))
    End If
    
    If mblnAllType = False Then
        blnMsg = False
        Set rs = gclsPackage.Get_Classdefaultcode(str上级编码, str最大编码, intMaxLength, blnMsg)
        If blnMsg = False Then
            If rs.BOF = False Then
                objTxtParent.Text = zlCommFun.NVL(rs("上级编码").Value)
                objChk.Value = zlCommFun.NVL(rs("调整编码").Value, 0)
                objTxt.Text = zlCommFun.NVL(rs("缺省编码").Value)
                objTxt.MaxLength = zlCommFun.NVL(rs("允许编码长度").Value, 0)
                objTxt.Tag = zlCommFun.NVL(rs("允许编码最大长度").Value)
                objChk.Enabled = (zlCommFun.NVL(rs("允许调整").Value, 0) = 1)
            End If
        Else
            objTxtParent.Text = ""
            objChk.Value = 0
            objTxt.Text = ""
            objTxt.MaxLength = Len(str最大编码)
            objTxt.Tag = intMaxLength
            objChk.Enabled = True
        End If
    Else
        objTxtParent.Text = ""
        objChk.Value = 0
        objTxt.Text = ""
        objTxt.MaxLength = Len(str最大编码)
        objTxt.Tag = intMaxLength
        objChk.Enabled = True
    End If
    
    NewDefaultCode = True
End Function

Private Function AnalyzeCode(ByVal lngKey As Long, ByRef objTxtParent As Object, ByRef objTxt As Object) As Boolean
    '******************************************************************************************************************
    '功能：分解编码
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Dim rs As New ADODB.Recordset
    
    Select Case mstrTemplate
    Case "体检指标分类"
        Set rs = gclsPackage.Get_Elementclass(lngKey)
    Case "体检项目分类"
        Set rs = gclsPackage.Get_Medicalclass(lngKey)
    Case "危害因素分类"
        Set rs = gclsPackage.Get_Virusclass(lngKey)
    Case "体检套餐分类"
        Set rs = gclsPackage.Get_Packageclass(lngKey)
    Case "体检诊断分类"
        Set rs = gclsPackage.Get_Diagnoseclass(lngKey)
    End Select
    
    If rs.BOF Then Exit Function
    
    objTxtParent.Text = zlCommFun.NVL(rs("上级编码").Value)
    objTxt.Text = zlCommFun.NVL(rs("编码").Value)
    
    If Len(objTxt.Text) >= Len(objTxtParent.Text) Then objTxt.Text = Mid(objTxt.Text, Len(objTxtParent.Text) + 1)
    
    objTxt.MaxLength = Len(objTxt.Text)
    objTxt.Tag = rs.Fields("编码").DefinedSize - Len(objTxtParent.Text)
    
    AnalyzeCode = True
End Function

Private Property Let DataChanged(ByVal blnData As Boolean)
    mblnChanged = blnData
End Property

Private Property Get DataChanged() As Boolean
    DataChanged = mblnChanged
End Property

Private Function ExecuteCommand(ParamArray varCmd() As Variant) As Boolean
    '--------------------------------------------------------------------------------------------------------------
    '功能：
    '参数：
    '返回：
    '--------------------------------------------------------------------------------------------------------------
    Dim intLoop As Integer
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim strSQL As String

    On Error GoTo errHand

    Call SQLRecord(rsSQL)

    For intLoop = 0 To UBound(varCmd)
        Select Case varCmd(intLoop)
        '--------------------------------------------------------------------------------------------------------------
        Case "初始数据"


            '设置最大输入长度
            txtParentCode.MaxLength = gclsPackage.Get_Maxlength(mstrTemplate, "编码")
            txt(1).MaxLength = gclsPackage.Get_Maxlength(mstrTemplate, "名称")
            
            '读取上级分类
            If mlngUpKey > 0 Then
                                
                Select Case mstrTemplate
                Case "体检指标分类"
                    Set rs = gclsPackage.Get_Elementclass(mlngUpKey)
                Case "体检项目分类"
                    Set rs = gclsPackage.Get_Medicalclass(mlngUpKey)
                Case "危害因素分类"
                    Set rs = gclsPackage.Get_Virusclass(mlngUpKey)
                Case "体检套餐分类"
                    Set rs = gclsPackage.Get_Packageclass(mlngUpKey)
                Case "体检诊断分类"
                    Set rs = gclsPackage.Get_Diagnoseclass(mlngUpKey)
                End Select
                
                If rs.BOF = False Then

                    txt(2).Text = AppendCode(zlCommFun.NVL(rs("名称").Value), zlCommFun.NVL(rs("编码").Value))
                    cmd(0).Tag = zlCommFun.NVL(rs("ID").Value, 0)
                    mlngUpKey = Val(cmd(0).Tag)
                    
                End If
            End If
            
        '--------------------------------------------------------------------------------------------------------------
        Case "清空数据"
        
            Call ExecuteCommand("缺省编码")
            
            txt(1).Text = ""
            txt(0).SetFocus
            
            DataChanged = False
                    
        '--------------------------------------------------------------------------------------------------------------
        Case "缺省编码"
            
            Call NewDefaultCode(mlngUpKey, txtParentCode, txt(0), chk(0))
                    
        '--------------------------------------------------------------------------------------------------------------
        Case "读取分类"
        
            Select Case mstrTemplate
            Case "体检指标分类"
                Set rs = gclsPackage.Get_Elementclass(mlngKey)
            Case "体检项目分类"
                Set rs = gclsPackage.Get_Medicalclass(mlngKey)
            Case "危害因素分类"
                Set rs = gclsPackage.Get_Virusclass(mlngKey)
            Case "体检套餐分类"
                Set rs = gclsPackage.Get_Packageclass(mlngKey)
            Case "体检诊断分类"
                Set rs = gclsPackage.Get_Diagnoseclass(mlngKey)
            End Select
            
            If rs.BOF = False Then
                
                txt(1).Text = zlCommFun.NVL(rs("名称").Value)

                
                txt(2).Text = AppendCode(zlCommFun.NVL(rs("上级名称").Value), zlCommFun.NVL(rs("上级编码").Value))
                
                cmd(0).Tag = zlCommFun.NVL(rs("上级id").Value, 0)
                    
                Call AnalyzeCode(mlngKey, txtParentCode, txt(0))

            End If
        '--------------------------------------------------------------------------------------------------------------
        Case "校验数据"
            ExecuteCommand = ValidEdit
            Exit Function
        '--------------------------------------------------------------------------------------------------------------
        Case "保存数据"
            ExecuteCommand = SaveEdit
            Exit Function
        End Select
    Next

    ExecuteCommand = True

    Exit Function
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function

Private Function ValidEdit() As Boolean
    '******************************************************************************************************************
    '功能：校验编辑数据的有效性
    '参数：
    '返回：
    '******************************************************************************************************************
    
    If txt(0).MaxLength = 0 Then
        ShowSimpleMsg "上级编码已经达到最大长度，不能设置下级！"
        cmdCancel.SetFocus
        Exit Function
    End If
    
    If chk(0).Value = 0 And Len(Trim(txt(0).Text)) <> txt(0).MaxLength Then
        ShowSimpleMsg "编码长度必须为" & txt(0).MaxLength & "位，除非你选择更改长度选项"
        LocationObj txt(0)
        Exit Function
    End If
    
    If Trim(txt(0).Text) = "" Then
        ShowSimpleMsg "编码不能为空值，必须输入！"
        LocationObj txt(0)
        Exit Function
    End If
    
    '检查编码是否为数字字符
    If CheckStrType(Trim(txt(0).Text), 99, "0123456789") = False Then
        ShowSimpleMsg "编码必须为数字字符！"
        LocationObj txt(0)
        Exit Function
    End If
    
    If Trim(txt(1).Text) = "" Then
        ShowSimpleMsg "名称不能为空值，必须输入！"
        LocationObj txt(1)
        Exit Function
    End If
    
    
    ValidEdit = True
    
End Function

Private Function SaveEdit() As Boolean
    '******************************************************************************************************************
    '功能：保存编辑数据到数据库
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngKey As Long
    Dim rsSQL As New ADODB.Recordset
    
    On Error GoTo errHand
        
    Call SQLRecord(rsSQL)
    
    If mlngKey = 0 Then
        '新增类型
        
        lngKey = zlDataBase.GetNextId(mstrTemplate)
        
        gstrSQL = "zl_" & mstrTemplate & "_Insert(" & lngKey & ",'" & Trim(txtParentCode.Text & txt(0).Text) & "','" & txt(1).Text & "'," & Val(cmd(0).Tag) & "," & chk(0).Value & ")"
        Call SQLRecordAdd(rsSQL, gstrSQL)
    Else
        '修改类型
        lngKey = mlngKey
        gstrSQL = "zl_" & mstrTemplate & "_Update(" & lngKey & ",'" & Trim(txtParentCode.Text & txt(0).Text) & "','" & txt(1).Text & "'," & Val(cmd(0).Tag) & "," & chk(0).Value & ")"
        Call SQLRecordAdd(rsSQL, gstrSQL)
    End If
    
    If SQLRecordExecute(rsSQL, Me.Caption) = False Then
        Exit Function
    End If

    mlngKey = lngKey
    SaveEdit = True
    
    Exit Function
    
errHand:
    
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub chk_Click(Index As Integer)
    If chk(Index).Value = 1 Then
        mlngSvrMaxLen = txt(0).MaxLength
        txt(0).MaxLength = Val(txt(0).Tag)
    Else
        txt(0).MaxLength = mlngSvrMaxLen
        txt(0).Text = Mid(txt(0).Text, 1, txt(0).MaxLength)
    End If
    
    DataChanged = True
    
    On Error Resume Next
    txt(0).SetFocus
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim rsData As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
'    Dim objPoint As POINTAPI
    
    Select Case mstrTemplate
    Case "体检指标分类"
        Set rsData = gclsPackage.Get_Elementclasstreesel(mlngKey)
    Case "体检项目分类"
        Set rsData = gclsPackage.Get_Medicalclasstreesel(mlngKey)
    Case "危害因素分类"
        Set rsData = gclsPackage.Get_Virusclasstreesel(mlngKey)
    Case "体检套餐分类"
        Set rsData = gclsPackage.Get_Packageclasstreesel(mlngKey)
    Case "体检诊断分类"
        Set rsData = gclsPackage.Get_Diagnoseclasstreesel(mlngKey)
    End Select
    
'    Call ClientToScreen(txt(2).hWnd, objPoint)
    
    If gclsBase.ShowPubSelect(Me, txt(2), 1, "", Me.Name & "\" & mstrTemplate & "选择", "请从下面选择一个" & mstrTemplate, rsData, rs, cmd(0).Left + cmd(0).Width - txt(2).Left, 3900, , mlngKey, , False) = 1 Then
'    If frmPubSelDialog.ShowDialog(Me, 1, rs, "", "请从下面选择一个" & mstrTemplate, objPoint.X * 15 - 30, objPoint.Y * 15 + txt(2).Height - 30, cmd(0).Left + cmd(0).Width - txt(2).Left, 3900, txt(2).Height, mlngKey, Me.Name & "\" & mstrTemplate & "选择", , False) Then
    
        If Val(cmd(0).Tag) <> zlCommFun.NVL(rs("ID")) Then
            If zlCommFun.NVL(rs("ID")) = -1 Then
                txt(2).Text = ""
                cmd(0).Tag = ""
                mblnAllType = True
            Else
                txt(2).Text = zlCommFun.NVL(rs("名称"))
                cmd(0).Tag = zlCommFun.NVL(rs("ID"))
                mblnAllType = False
            End If
            
            mlngUpKey = Val(cmd(0).Tag)
            
            Call ExecuteCommand("缺省编码")
            DataChanged = True
            mblnAllType = False
            Call AdjustCodePostion(Me, txtParentCode, txt(0))
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    
    If DataChanged Then
        If ExecuteCommand("校验数据") = False Then Exit Sub
        If ExecuteCommand("保存数据") Then
            
            RaiseEvent AfterSaved(mlngKey)
            
            On Error Resume Next
            Call mfrmMain.RefreshClass(mlngKey)
            On Error GoTo 0
            
            DataChanged = False
            
            mblnOK = True
        Else
            
        End If
    End If
    
    If mbytMode <> 1 Then
        Unload Me
    Else
        '清空相关的数据，以便进行下一次增加录入
        mlngKey = 0
        Call ExecuteCommand("清空数据", "缺省编码")
        
        DataChanged = False
        Call LocationObj(txt(0))
        
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If DataChanged Then
        Cancel = (MsgBox("新增或修改的数据必须保存后才生效，是否不保存就退出？", vbYesNo + vbQuestion + vbDefaultButton2, ParamInfo.系统名称) = vbNo)
    End If
End Sub

Private Sub txt_Change(Index As Integer)
    DataChanged = True
End Sub

Private Sub txt_GotFocus(Index As Integer)
    
    zlControl.TxtSelAll txt(Index)
    
    Select Case Index
    Case 1, 3
        zlCommFun.OpenIme True
    End Select
        
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        txt(2).Text = ""
        cmd(0).Tag = ""
        
        mlngUpKey = 0
        Call ExecuteCommand("缺省编码")
        DataChanged = True
        
        Call AdjustCodePostion(Me, txtParentCode, txt(0))
    End If
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
        If Index = 0 Then
            If FilterKeyAscii(KeyAscii, 99, "0123456789") = 0 Then KeyAscii = 0
        End If
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
    Select Case Index
    Case 1, 3
        zlCommFun.OpenIme False
    End Select
End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        glngTXTProc = GetWindowLong(txt(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
End Sub



