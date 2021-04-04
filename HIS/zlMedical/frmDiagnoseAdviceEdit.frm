VERSION 5.00
Begin VB.Form frmDiagnoseAdviceEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "体检诊断编辑"
   ClientHeight    =   5550
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7590
   Icon            =   "frmDiagnoseAdviceEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin zl9Medical.VsfGrid vsf 
      Height          =   2175
      Left            =   60
      TabIndex        =   21
      Top             =   3300
      Width           =   6090
      _ExtentX        =   10742
      _ExtentY        =   3836
   End
   Begin VB.Frame fra 
      Height          =   3285
      Left            =   60
      TabIndex        =   20
      Top             =   -30
      Width           =   6090
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   5
         Left            =   3465
         TabIndex        =   13
         Top             =   2865
         Width           =   330
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   3
         Left            =   5415
         TabIndex        =   16
         Top             =   2865
         Width           =   345
      End
      Begin VB.CheckBox chk 
         Caption         =   "随访期限(&3)"
         Height          =   240
         Index           =   2
         Left            =   4140
         TabIndex        =   15
         Top             =   2895
         Width           =   1320
      End
      Begin VB.CheckBox chk 
         Caption         =   "复查间隔(&2)"
         Height          =   240
         Index           =   1
         Left            =   2160
         TabIndex        =   12
         Top             =   2910
         Width           =   1320
      End
      Begin VB.CheckBox chk 
         Caption         =   "疾病(&1)"
         Height          =   240
         Index           =   0
         Left            =   1140
         TabIndex        =   11
         Top             =   2910
         Width           =   975
      End
      Begin VB.TextBox txt 
         Height          =   1110
         Index           =   6
         Left            =   1140
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   1320
         Width           =   4785
      End
      Begin VB.CommandButton cmd 
         Caption         =   "…"
         Height          =   255
         Index           =   0
         Left            =   5640
         TabIndex        =   10
         Top             =   2520
         Width           =   255
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   2
         Left            =   1140
         TabIndex        =   5
         Top             =   960
         Width           =   4785
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   1140
         TabIndex        =   3
         Top             =   600
         Width           =   4785
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   0
         Left            =   1140
         TabIndex        =   1
         Top             =   240
         Width           =   4785
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   4
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   2490
         Width           =   4785
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "月"
         Height          =   180
         Index           =   6
         Left            =   3810
         TabIndex        =   14
         Top             =   2925
         Width           =   180
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "月"
         Height          =   180
         Index           =   3
         Left            =   5775
         TabIndex        =   17
         Top             =   2895
         Width           =   180
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "诊断建议(&A)"
         Height          =   180
         Index           =   5
         Left            =   90
         TabIndex        =   6
         Top             =   1335
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "所属分类(&U)"
         Height          =   180
         Index           =   4
         Left            =   90
         TabIndex        =   8
         Top             =   2550
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "诊断简码(&S)"
         Height          =   180
         Index           =   2
         Left            =   90
         TabIndex        =   4
         Top             =   1050
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "诊断名称(&N)"
         Height          =   180
         Index           =   1
         Left            =   90
         TabIndex        =   2
         Top             =   675
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "诊断编码(&B)"
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   0
         Top             =   330
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6375
      TabIndex        =   19
      Top             =   510
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6375
      TabIndex        =   18
      Top             =   75
      Width           =   1100
   End
End
Attribute VB_Name = "frmDiagnoseAdviceEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'（１）窗体级变量定义**************************************************************************************************
Private mblnStartUp As Boolean                          '窗体启动标志
Private mblnOK As Boolean
Private mfrmMain As Object
Private mlngKey As Long
Private mlngUpKey As Long


Private Enum mCol
    名称 = 1
    编码
    类别
    
End Enum

'（２）自定义过程或函数************************************************************************************************

Private Function CheckHave(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  检查是否有重复的项目
    '参数:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    
    For lngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(lngLoop)) = lngKey And vsf.Row <> lngLoop Then
            CheckHave = True
            Exit Function
        End If
    Next
End Function

Public Function ShowEdit(ByVal frmMain As Object, ByVal lngKey As Long, ByVal lngUpKey As Long) As Boolean
    
    mblnStartUp = True
    mblnOK = False
    
    mlngKey = lngKey
    mlngUpKey = lngUpKey
        
    Set mfrmMain = frmMain
    
    If InitData = False Then Exit Function
    
    If mlngKey > 0 Then
        
        '修改存在的项目
        If ReadData(mlngKey) = False Then Exit Function
    Else
        
        '新增加类型,产生缺省的编码
        
        txt(0).Text = NewDefaultCode(mlngUpKey)
        
    End If
    
    cmdOK.Tag = ""
    
    Me.Show 1, frmMain
    
    ShowEdit = mblnOK
    
End Function

Private Function NewDefaultCode(ByVal lngUpKey As Long) As String
    
    '------------------------------------------------------------------------------------------------------------------
    '功能:生成缺省编码
    '------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------
    '功能:产生缺省编码
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim intMaxLength As Integer
    Dim str最大编码 As String
    Dim str上级编码 As String
    
    '读取上级编码
    strSQL = "SELECT B.编码 AS 上级编码 FROM 体检诊断建议 B WHERE B.序号=[1]"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngUpKey)
    If rs.BOF Then Exit Function
    
    intMaxLength = rs.Fields(0).DefinedSize
    str上级编码 = zlCommFun.NVL(rs("上级编码").Value)
            
    If intMaxLength = Len(str上级编码) Then
        MsgBox "最大编码和编码长度已经达到最大限制，无法递增编码", vbExclamation, gstrSysName
        Exit Function
    End If
        
    '读取同级最大编码+1
    If lngUpKey = 0 Then
        strSQL = "SELECT MAX(B.编码) AS 最大编码 FROM 体检诊断建议 B WHERE B.末级=1 AND B.上级序号 IS NULL "
        Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Else
        strSQL = "SELECT MAX(B.编码) AS 最大编码 FROM 体检诊断建议 B WHERE B.末级=1 AND B.上级序号=[1]"
        Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngUpKey)
    End If
    
    If rs.BOF Then Exit Function
    
    str最大编码 = Trim(zlCommFun.NVL(rs("最大编码").Value, ""))
  
    If str最大编码 = "" Then
        str最大编码 = str上级编码 & "001"
    Else
        str最大编码 = Format(Val(str最大编码) + 1, String(Len(str最大编码), "0"))
    End If
    
    NewDefaultCode = str最大编码
End Function

Private Function ReadData(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
        
    On Error GoTo errHand
    
    gstrSQL = "SELECT A.*,Decode(C.名称,Null,'','【'||C.编码||'】'||C.名称) As 上级名称 " & _
            "FROM 体检诊断建议 A,体检诊断建议 C WHERE A.上级序号=C.序号(+) And A.序号=[1]"
    
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
    If rs.BOF = False Then
        txt(0).Text = zlCommFun.NVL(rs("编码").Value)
        txt(1).Text = zlCommFun.NVL(rs("名称").Value)
        txt(2).Text = zlCommFun.NVL(rs("简码").Value)
        
        txt(4).Text = zlCommFun.NVL(rs("上级名称").Value)
        cmd(0).Tag = zlCommFun.NVL(rs("上级序号").Value)
        
        chk(0).Value = zlCommFun.NVL(rs("是否疾病").Value, 0)
        
        txt(6).Text = zlCommFun.NVL(rs("诊断建议").Value)
        
        txt(5).Text = zlCommFun.NVL(rs("复查间隔").Value)
        txt(3).Text = zlCommFun.NVL(rs("随访期限").Value)
        
        chk(1).Value = IIf(Val(txt(5).Text) > 0, 1, 0)
        chk(2).Value = IIf(Val(txt(3).Text) > 0, 1, 0)
        
        txt(5).Visible = (chk(1).Value = 1)
        lbl(6).Visible = (chk(1).Value = 1)
        
        txt(3).Visible = (chk(2).Value = 1)
        lbl(3).Visible = (chk(2).Value = 1)
        
    End If
    
    gstrSQL = "SELECT B.ID,B.名称,b.编码,c.名称 As 类别 FROM 体检诊断依据 A,诊疗项目目录 B,诊疗项目类别 c WHERE A.诊疗项目id=B.ID And A.诊断序号=[1] And c.编码=b.类别"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
    If rs.BOF = False Then
        Call FillGrid(vsf, rs)
    End If
    
    ReadData = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
    
End Function

Private Function InitData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
        
    txt(0).MaxLength = GetMaxLength("体检诊断建议", "编码")
    txt(1).MaxLength = GetMaxLength("体检诊断建议", "名称")
    txt(2).MaxLength = GetMaxLength("体检诊断建议", "简码")
        
    txt(6).MaxLength = GetMaxLength("体检诊断建议", "诊断建议")
        
    gstrSQL = "SELECT '['||编码||']'||名称 AS 名称 FROM 体检诊断建议 WHERE 序号=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngUpKey)
    If rs.BOF = False Then
        
        txt(4).Text = zlCommFun.NVL(rs("名称"))
        cmd(0).Tag = mlngUpKey
        
    End If

    With vsf
        .Cols = 0
        .NewColumn "", 255
        .NewColumn "名称", 2400, 1, "...", 1
        .NewColumn "编码", 1500, 1
        .NewColumn "类别", 900, 1
        .FixedCols = 1
    End With
    
    
    InitData = True
    
    Exit Function
    
errHand:
    
    If ErrCenter = 1 Then Resume
    
End Function

Private Function ValidEdit() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:校验编辑数据的有效性
    '------------------------------------------------------------------------------------------------------------------
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

Private Function SaveEdit(ByRef lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '
    '------------------------------------------------------------------------------------------------------------------
    Dim blnTran As Boolean
    Dim lngLoop As Long
    Dim strSQL() As String
    
    On Error GoTo errHand
    
    ReDim Preserve strSQL(1 To 1)
    
    If mlngKey = 0 Then
        '新增类型
        
        lngKey = GetMaxNo
        strSQL(ReDimArray(strSQL)) = "ZL_体检诊断建议_INSERT(" & lngKey & ",'" & Trim(txt(0).Text) & "','" & txt(1).Text & "','" & txt(2).Text & "'," & chk(0).Value & ",'" & txt(6).Text & "'," & IIf(chk(1).Value = 0, "NULL", Val(txt(5).Text)) & "," & IIf(chk(2).Value = 0, "NULL", Val(txt(3).Text)) & "," & Val(cmd(0).Tag) & ",1)"
    Else
        '修改类型
        lngKey = mlngKey
        strSQL(ReDimArray(strSQL)) = "ZL_体检诊断建议_UPDATE(" & lngKey & ",'" & Trim(txt(0).Text) & "','" & txt(1).Text & "','" & txt(2).Text & "'," & chk(0).Value & ",'" & txt(6).Text & "'," & IIf(chk(1).Value = 0, "NULL", Val(txt(5).Text)) & "," & IIf(chk(2).Value = 0, "NULL", Val(txt(3).Text)) & "," & Val(cmd(0).Tag) & ")"
    End If
    
    strSQL(ReDimArray(strSQL)) = "ZL_体检诊断依据_Delete(" & lngKey & ")"
    For lngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(lngLoop)) > 0 Then
            strSQL(ReDimArray(strSQL)) = "ZL_体检诊断依据_Insert(" & lngKey & "," & Val(vsf.RowData(lngLoop)) & ")"
        End If
    Next
    
    blnTran = True
    gcnOracle.BeginTrans
    For lngLoop = 1 To UBound(strSQL)
        If strSQL(lngLoop) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(lngLoop), Me.Caption)
    Next
    gcnOracle.CommitTrans
    blnTran = False
    
    SaveEdit = True
    
    Exit Function
    
errHand:
    
    If ErrCenter = 1 Then
        Resume
    End If
    
    If blnTran Then gcnOracle.RollbackTrans
    
End Function

Private Function GetMaxNo() As Long
    '------------------------------------------------------------------------------------------------------------------
    '
    '------------------------------------------------------------------------------------------------------------------
    
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    gstrSQL = "SELECT NVL(MAX(序号),0)+1 AS 序号 FROM 体检诊断建议"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If rs.BOF = False Then GetMaxNo = rs("序号").Value
        
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub chk_Click(Index As Integer)
    cmdOK.Tag = "Changed"
    
    txt(5).Visible = (chk(1).Value = 1)
    lbl(6).Visible = txt(5).Visible
    
    txt(3).Visible = (chk(2).Value = 1)
    lbl(3).Visible = txt(3).Visible
    
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim rs As New ADODB.Recordset
    Dim objPoint As POINTAPI
    
    
    If Index = 0 Then
        gstrSQL = "SELECT 0 As 末级,-1 AS ID,0 AS 上级id,'所有分类' AS 名称,'' AS 编码 FROM DUAL " & _
                    "UNION ALL " & _
                    "SELECT 0 As 末级,序号 AS ID,DECODE(上级序号,NULL,-1,上级序号) AS 上级id,'【'||编码||'】'||名称 AS 名称,编码 FROM 体检诊断建议 WHERE 末级=0  START WITH 上级序号 IS NULL AND 序号<>" & mlngKey & " CONNECT BY PRIOR 序号=上级序号 "
    
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        
        Call ClientToScreen(txt(4).hWnd, objPoint)
        
        If frmSelectDialog.ShowSelect(Me, 1, rs, "", "请从下面选择一个分类", objPoint.X * 15 - 30, objPoint.Y * 15 + txt(4).Height - 30, txt(4).Width, 3900, txt(4).Height, mlngKey, Me.Name & "\体检类型分类选择", , False) Then
            If Val(cmd(0).Tag) <> zlCommFun.NVL(rs("ID")) Then
                If zlCommFun.NVL(rs("ID")) = -1 Then
                    txt(4).Text = ""
                    cmd(0).Tag = ""
                Else
                    txt(4).Text = zlCommFun.NVL(rs("名称"))
                    cmd(0).Tag = zlCommFun.NVL(rs("ID"))
                End If
                           
                cmdOK.Tag = "Changed"
            End If
        End If
    End If
    
End Sub

'（３）窗体及其控件的事件处理******************************************************************************************
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim lngKey As Long
        
    If cmdOK.Tag <> "" Then
            
        If ValidEdit() = False Then Exit Sub
        If SaveEdit(lngKey) = False Then Exit Sub
        mblnOK = True
        
        '更新调用窗体的数据显示
        
        Call mfrmMain.EditRefresh("体检诊断目录", lngKey)
        
        If mlngKey = 0 Then
            
            txt(0).Text = NewDefaultCode(Val(cmd(0).Tag))
            txt(1).Text = ""
            txt(2).Text = ""
            
            txt(6).Text = ""
                    
            txt(0).SetFocus
            
            cmdOK.Tag = ""
            Exit Sub
        End If
        
    End If
    
    cmdOK.Tag = ""
    
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If cmdOK.Tag <> "" Then
        Cancel = (MsgBox("新增或修改的数据必须保存后才生效，是否不保存就退出？", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo)
    End If
End Sub

Private Sub txt_Change(Index As Integer)
    cmdOK.Tag = "Changed"
            
End Sub

Private Sub txt_GotFocus(Index As Integer)
    
    zlControl.TxtSelAll txt(Index)
    
    Select Case Index
    Case 1, 6
        zlCommFun.OpenIme True
    End Select
        
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)

    Dim rs As New ADODB.Recordset
    

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        zlCommFun.PressKey vbKeyTab
        
        If Index = 4 Then zlCommFun.PressKey vbKeyTab
                
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
        
        If Index = 0 Then If FilterKeyAscii(KeyAscii, 99, "0123456789") = 0 Then KeyAscii = 0
        If Index = 2 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
                
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
    Select Case Index
    Case 1, 6
        zlCommFun.OpenIme False
        If Index = 1 Then
            If InStr(txt(Index).Text, "'") = 0 Then txt(2).Text = zlGetSymbol(txt(Index).Text)
        End If
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
    If Cancel Then Exit Sub
    
End Sub

Private Sub vsf_AfterDeleteRow(ByVal Row As Long, ByVal Col As Long)
    cmdOK.Tag = "Changed"
End Sub

Private Sub vsf_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Cancel = (Val(vsf.RowData(Row)) <= 0)
End Sub

Private Sub vsf_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsData As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    
    Select Case Col
        Case mCol.名称
            
            gstrSQL = GetPublicSQL(SQL.体检项目选择)
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, 1, 2)
            If ShowGrdSelect(Me, vsf, "编码,1200,0,1;名称,2700,0,0;单位,900,0,0;类别,900,0,0", Me.Name & "\体检项目选择", "请从列表中选择一个体检项目。", rsData, rs, 8790, 5100) Then
                
                If CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                    ShowSimpleMsg "选择的项目“" & zlCommFun.NVL(rs("名称").Value) & "”已被选择！"
                    Exit Sub
                End If
                
                vsf.EditText = zlCommFun.NVL(rs("名称").Value)
                vsf.TextMatrix(Row, mCol.类别) = zlCommFun.NVL(rs("类别").Value)
                vsf.TextMatrix(Row, mCol.名称) = zlCommFun.NVL(rs("名称").Value)
                vsf.TextMatrix(Row, mCol.编码) = zlCommFun.NVL(rs("编码").Value)
                vsf.RowData(Row) = zlCommFun.NVL(rs("ID").Value)
                
                cmdOK.Tag = "Changed"
                
            End If

    End Select

End Sub

Private Sub vsf_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, ByVal ComboList As String, KeyCode As Integer, ByVal Shift As Integer, Cancel As Boolean)
    Dim strTmp As String
    Dim strText As String
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    
    
    If KeyCode = vbKeyReturn Then
        If ComboList = "..." Then
            
            If InStr(vsf.EditText, "'") > 0 Then
                KeyCode = 0
                vsf.EditText = ""
                Cancel = True
                Exit Sub
            End If
    
            Select Case Col
                Case mCol.名称
                    
                    strText = UCase(vsf.EditText)
                    gstrSQL = GetPublicSQL(SQL.体检项目过滤选择, strText)
                    
                    If ParamInfo.项目输入匹配方式 = 1 Then
                        strTmp = strText & "%"
                    Else
                        strTmp = "%" & strText & "%"
                    End If

                    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "C", "D", strText & "%", strTmp, 1, 2)
                    
                    If ShowGrdFilter(Me, vsf, "编码,1200,0,1;名称,2700,0,0;单位,900,0,0;类别,900,0,0", Me.Name & "\体检项目过滤选择", "请从列表中选择一个体检项目。", rsData, rs, 8790, 5100) Then

                        If CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                            ShowSimpleMsg "选择的项目“" & zlCommFun.NVL(rs("名称").Value) & "”已被选择！"
                            Exit Sub
                        End If

                        vsf.EditText = zlCommFun.NVL(rs("名称").Value)
                        vsf.TextMatrix(Row, mCol.类别) = zlCommFun.NVL(rs("类别").Value)
                        vsf.TextMatrix(Row, mCol.名称) = zlCommFun.NVL(rs("名称").Value)
                        vsf.TextMatrix(Row, mCol.编码) = zlCommFun.NVL(rs("编码").Value)
                        vsf.Cell(flexcpData, Row, Col) = vsf.TextMatrix(Row, Col)
                        vsf.RowData(Row) = zlCommFun.NVL(rs("ID").Value)
                        
                        
                        cmdOK.Tag = "Changed"
                    Else
                        KeyCode = 0
                        Cancel = True
                        
                        vsf.Cell(flexcpData, Row, Col) = vsf.Cell(flexcpData, Row, Col)
                        vsf.EditText = vsf.Cell(flexcpData, Row, Col)
                        vsf.TextMatrix(Row, Col) = vsf.Cell(flexcpData, Row, Col)
                        
                    End If
            End Select
        End If
    Else
        cmdOK.Tag = "Changed"
    End If
End Sub
