VERSION 5.00
Begin VB.Form frmKindEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "体检类型"
   ClientHeight    =   3270
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6240
   Icon            =   "frmKindEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fra 
      Height          =   3165
      Left            =   105
      TabIndex        =   15
      Top             =   0
      Width           =   4545
      Begin VB.ComboBox cbo 
         Height          =   300
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2355
         Width           =   3255
      End
      Begin VB.CommandButton cmd 
         Caption         =   "&P"
         Height          =   255
         Index           =   0
         Left            =   4050
         TabIndex        =   12
         Top             =   2760
         Width           =   255
      End
      Begin VB.TextBox txt 
         Height          =   825
         Index           =   3
         Left            =   1080
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   1440
         Width           =   3255
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   2
         Left            =   1080
         TabIndex        =   5
         Top             =   1035
         Width           =   3255
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   1080
         TabIndex        =   3
         Top             =   645
         Width           =   3255
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   0
         Left            =   1080
         TabIndex        =   1
         Top             =   240
         Width           =   3255
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   4
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   2730
         Width           =   3255
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "适用(&R)"
         Height          =   180
         Index           =   5
         Left            =   405
         TabIndex        =   8
         Top             =   2430
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "上级(&U)"
         Height          =   180
         Index           =   4
         Left            =   405
         TabIndex        =   10
         Top             =   2790
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "说明(&T)"
         Height          =   180
         Index           =   3
         Left            =   405
         TabIndex        =   6
         Top             =   1470
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "简码(&S)"
         Height          =   180
         Index           =   2
         Left            =   405
         TabIndex        =   4
         Top             =   1125
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "名称(&N)"
         Height          =   180
         Index           =   1
         Left            =   405
         TabIndex        =   2
         Top             =   720
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "编码(&B)"
         Height          =   180
         Index           =   0
         Left            =   405
         TabIndex        =   0
         Top             =   330
         Width           =   630
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4995
      TabIndex        =   14
      Top             =   765
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4995
      TabIndex        =   13
      Top             =   285
      Width           =   1100
   End
End
Attribute VB_Name = "frmKindEdit"
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

'（２）自定义过程或函数************************************************************************************************

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
    strSQL = "SELECT B.编码 AS 上级编码 FROM 体检类型 B WHERE B.序号=[1]"
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
        strSQL = "SELECT MAX(B.编码) AS 最大编码 FROM 体检类型 B WHERE B.末级=1 AND B.上级序号 IS NULL "
        Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Else
        strSQL = "SELECT MAX(B.编码) AS 最大编码 FROM 体检类型 B WHERE B.末级=1 AND B.上级序号=[1]"
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
    
    gstrSQL = "SELECT * FROM 体检类型 WHERE 序号=[1]"
    
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
    If rs.BOF = False Then
        txt(0).Text = zlCommFun.NVL(rs("编码").Value)
        txt(1).Text = zlCommFun.NVL(rs("名称").Value)
        txt(2).Text = zlCommFun.NVL(rs("简码").Value)
        txt(3).Text = zlCommFun.NVL(rs("说明").Value)
        
        zlControl.CboLocate cbo, zlCommFun.NVL(rs("适用范围").Value, 0), True
        
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
        
    txt(0).MaxLength = GetMaxLength("体检类型", "编码")
    txt(1).MaxLength = GetMaxLength("体检类型", "名称")
    txt(2).MaxLength = GetMaxLength("体检类型", "简码")
    txt(3).MaxLength = GetMaxLength("体检类型", "说明")
        
    gstrSQL = "SELECT '['||编码||']'||名称 AS 名称 FROM 体检类型 WHERE 序号=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngUpKey)
    If rs.BOF = False Then
        
        txt(4).Text = zlCommFun.NVL(rs("名称"))
        cmd(0).Tag = mlngUpKey
        
    End If
    
    cbo.Clear
    
    cbo.AddItem "0-所有"
    cbo.ItemData(cbo.NewIndex) = 0
    
    cbo.AddItem "1-个人"
    cbo.ItemData(cbo.NewIndex) = 1
    
    cbo.AddItem "2-团体"
    cbo.ItemData(cbo.NewIndex) = 2
    
    cbo.ListIndex = 0
    
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
        strSQL(ReDimArray(strSQL)) = "ZL_体检类型_INSERT(" & lngKey & ",'" & Trim(txt(0).Text) & "','" & txt(1).Text & "','" & txt(2).Text & "','" & txt(3).Text & "'," & cbo.ItemData(cbo.ListIndex) & "," & Val(cmd(0).Tag) & ",1)"
    Else
        '修改类型
        lngKey = mlngKey
        strSQL(ReDimArray(strSQL)) = "ZL_体检类型_UPDATE(" & lngKey & ",'" & Trim(txt(0).Text) & "','" & txt(1).Text & "','" & txt(2).Text & "','" & txt(3).Text & "'," & cbo.ItemData(cbo.ListIndex) & "," & Val(cmd(0).Tag) & ")"
    End If
    
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
    
    If ErrCenter = 1 Then Resume
    If blnTran Then gcnOracle.RollbackTrans
    
End Function

Private Function GetMaxNo() As Long
    '------------------------------------------------------------------------------------------------------------------
    '
    '------------------------------------------------------------------------------------------------------------------
    
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    gstrSQL = "SELECT NVL(MAX(序号),0)+1 AS 序号 FROM 体检类型"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If rs.BOF = False Then GetMaxNo = rs("序号").Value
        
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub chk_Click()
    cmdOK.Tag = "Changed"
End Sub

Private Sub chk_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbo_Click()
    cmdOK.Tag = "Changed"
End Sub

Private Sub cbo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim rs As New ADODB.Recordset
    Dim objPoint As POINTAPI
    
    gstrSQL = "SELECT 0 As 末级,-1 AS ID,0 AS 上级id,'所有分类' AS 名称,'' AS 编码 FROM DUAL " & _
                "UNION ALL " & _
                "SELECT 0 As 末级,序号 AS ID,DECODE(上级序号,NULL,-1,上级序号) AS 上级id,'['||编码||']'||名称 AS 名称,编码 FROM 体检类型 WHERE 末级=0 START WITH 上级序号 IS NULL AND 序号<>" & mlngKey & " CONNECT BY PRIOR 序号=上级序号 "

    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    Call ClientToScreen(txt(4).hWnd, objPoint)
    
    If frmSelectDialog.ShowSelect(Me, 1, rs, "", "请从下面选择一个分类", objPoint.X * 15 - 30, objPoint.Y * 15 + txt(4).Height - 30, txt(4).Width, 3900, txt(4).Height, mlngKey, Me.Name & "\体检类型分类选择", , False) Then
    
        If Val(cmd(0).Tag) <> zlCommFun.NVL(rs("ID")) Then
            If zlCommFun.NVL(rs("ID")) = -1 Then
'                mstr上级编码 = ""
                txt(4).Text = ""
                cmd(0).Tag = ""
            Else
'                mstr上级编码 = zlCommFun.NVL(rs("编码"))
                txt(4).Text = zlCommFun.NVL(rs("名称"))
                cmd(0).Tag = zlCommFun.NVL(rs("ID"))
                
            End If
                       
            cmdOK.Tag = "Changed"
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
        
        Call mfrmMain.EditRefresh("体检类型", lngKey)
        
        If mlngKey = 0 Then
            
            txt(0).Text = NewDefaultCode(Val(cmd(0).Tag))
            txt(1).Text = ""
            txt(2).Text = ""
            txt(3).Text = ""
                    
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
    Case 1, 3
        zlCommFun.OpenIme True
    End Select
        
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        zlCommFun.PressKey vbKeyTab
        If Index = 4 Then zlCommFun.PressKey vbKeyTab
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
        If Index = 2 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If Index = 0 Then
            If FilterKeyAscii(KeyAscii, 99, "0123456789") = 0 Then KeyAscii = 0
        End If
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
    Select Case Index
    Case 1, 3
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
End Sub
