VERSION 5.00
Begin VB.Form frmGroupEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "团体信息"
   ClientHeight    =   6150
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   10170
   Icon            =   "frmGroupEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   10170
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   165
      TabIndex        =   21
      Top             =   5640
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   8895
      TabIndex        =   20
      Top             =   5625
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   7695
      TabIndex        =   19
      Top             =   5625
      Width           =   1100
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5415
      Left            =   -15
      ScaleHeight     =   5415
      ScaleWidth      =   10125
      TabIndex        =   22
      Top             =   0
      Width           =   10125
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   9
         Left            =   5580
         TabIndex        =   15
         Top             =   1830
         Width           =   4410
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   8
         Left            =   1230
         TabIndex        =   13
         Top             =   1860
         Width           =   2940
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   10
         Left            =   1230
         TabIndex        =   17
         Top             =   2250
         Width           =   8775
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   2
         Left            =   3585
         TabIndex        =   5
         Top             =   480
         Width           =   6435
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   7
         Left            =   1230
         TabIndex        =   9
         Top             =   1485
         Width           =   2940
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   0
         Left            =   1230
         TabIndex        =   1
         Top             =   120
         Width           =   8790
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   3
         Left            =   1230
         TabIndex        =   7
         Top             =   1125
         Width           =   2940
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   1230
         MaxLength       =   18
         TabIndex        =   3
         Top             =   480
         Width           =   1590
      End
      Begin VB.TextBox txt 
         Height          =   2745
         Index           =   4
         Left            =   225
         TabIndex        =   18
         Top             =   2610
         Width           =   9780
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   5
         Left            =   5580
         TabIndex        =   11
         Top             =   1470
         Width           =   4410
      End
      Begin VB.Label lbl身份证号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "编码(&U)"
         Height          =   180
         Left            =   195
         TabIndex        =   2
         Top             =   540
         Width           =   630
      End
      Begin VB.Label lbl身份 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "银行帐号(&Z)"
         Height          =   180
         Left            =   4455
         TabIndex        =   14
         Top             =   1875
         Width           =   990
      End
      Begin VB.Label lbl职业 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "开户银行(&B)"
         Height          =   180
         Left            =   195
         TabIndex        =   12
         Top             =   1935
         Width           =   990
      End
      Begin VB.Label lbl民族 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "名称(&N)"
         Height          =   180
         Left            =   195
         TabIndex        =   0
         Top             =   195
         Width           =   630
      End
      Begin VB.Label lbl国籍 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "简码(&I)"
         Height          =   180
         Left            =   2895
         TabIndex        =   4
         Top             =   555
         Width           =   630
      End
      Begin VB.Label lbl学历 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "联系电话(&T)"
         Height          =   180
         Left            =   195
         TabIndex        =   8
         Top             =   1575
         Width           =   990
      End
      Begin VB.Label lvl婚姻状况 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "联系地址(&A)"
         Height          =   180
         Left            =   195
         TabIndex        =   16
         Top             =   2295
         Width           =   990
      End
      Begin VB.Label lbl家庭地址 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "联系人(&L)"
         Height          =   180
         Left            =   195
         TabIndex        =   6
         Top             =   1200
         Width           =   810
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "电子邮件(&E)"
         Height          =   180
         Index           =   7
         Left            =   4455
         TabIndex        =   10
         Top             =   1560
         Width           =   990
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         X1              =   150
         X2              =   10320
         Y1              =   930
         Y2              =   930
      End
   End
End
Attribute VB_Name = "frmGroupEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'（１）窗体级变量定义**************************************************************************************************
Private mblnStartUp As Boolean                          '窗体启动标志
Private mblnOK As Boolean
Private mfrmMain As Object
Private mvarParam As Variant
Private mblnDataChange As Boolean
Private mrsGroup As New ADODB.Recordset

'（２）自定义过程或函数************************************************************************************************
Private Property Let EditChanged(ByVal vData As Boolean)
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '值域:
    '------------------------------------------------------------------------------------------------------------------
    mblnDataChange = vData

End Property

Private Function ClearData(Optional ByVal strMenuItem As String = "") As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '参数:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long

    On Error Resume Next

    For lngLoop = 0 To txt.UBound
        txt(lngLoop).Text = ""
        txt(lngLoop).Tag = ""
    Next

    On Error GoTo 0

    Call InitData

    EditChanged = True


End Function

Public Function ShowEdit(ByVal frmMain As Object, ByRef rsGroup As ADODB.Recordset, Optional ByVal blnModify As Boolean = True) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  显示编辑窗体，是与调用窗体的接口函数
    '参数:  frmMain         调用窗体对象
    '       lngKey          预约登记id
    '返回:  True
    '       False
    '------------------------------------------------------------------------------------------------------------------
    mblnStartUp = True
    mblnOK = False
        
    'mvarParam = Split(strParam, "'")
    Set mrsGroup = rsGroup

    Set mfrmMain = frmMain

    If InitData = False Then Exit Function
    If ReadData = False Then Exit Function
    
    If Trim(txt(1).Text) = "" Then txt(1).Text = GetNextCode("合约单位", "编码", "上级id IS NULL")
        
'    If blnModify Then
'        txt(0).Text = mvarParam(1)
'        txt(3).Text = mvarParam(2)
'        txt(7).Text = mvarParam(3)
'    End If
        
    EditChanged = False

    Me.Show 1, frmMain
        
    ShowEdit = mblnOK
    If mblnOK Then Set rsGroup = mrsGroup
    
End Function

Private Function ReadData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  读取数据
    '参数:  lngKey      体检类型序号
    '返回:  True        读取成功
    '       False       读取失败
    '------------------------------------------------------------------------------------------------------------------

    On Error GoTo errHand

    If mrsGroup.BOF = False Then
        txt(0).Text = zlCommFun.NVL(mrsGroup("名称").Value)
        txt(1).Text = zlCommFun.NVL(mrsGroup("编码").Value)
        txt(2).Text = zlCommFun.NVL(mrsGroup("简码").Value)
        txt(3).Text = zlCommFun.NVL(mrsGroup("联系人").Value)
        txt(7).Text = zlCommFun.NVL(mrsGroup("电话").Value)
        txt(5).Text = zlCommFun.NVL(mrsGroup("电子邮件").Value)
        txt(8).Text = zlCommFun.NVL(mrsGroup("开户银行").Value)
        txt(9).Text = zlCommFun.NVL(mrsGroup("帐号").Value)
        txt(10).Text = zlCommFun.NVL(mrsGroup("地址").Value)
        txt(4).Text = zlCommFun.NVL(mrsGroup("说明").Value)
    End If

    ReadData = True

    Exit Function

errHand:
    If ErrCenter = 1 Then Resume

End Function

Private Function InitData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  初始化设置
    '返回:  True        初始化成功
    '       False       初始化失败
    '------------------------------------------------------------------------------------------------------------------

    On Error GoTo errHand

    '设置最大输入长度

    txt(0).MaxLength = GetMaxLength("合约单位", "名称")
    txt(1).MaxLength = GetMaxLength("合约单位", "编码")
    txt(2).MaxLength = GetMaxLength("合约单位", "简码")
    txt(3).MaxLength = GetMaxLength("合约单位", "联系人")
    txt(7).MaxLength = GetMaxLength("合约单位", "电话")
    txt(5).MaxLength = GetMaxLength("合约单位", "电子邮件")
    txt(8).MaxLength = GetMaxLength("合约单位", "开户银行")
    txt(9).MaxLength = GetMaxLength("合约单位", "帐号")
    txt(10).MaxLength = GetMaxLength("合约单位", "地址")
    txt(4).MaxLength = GetMaxLength("合约单位", "说明")
    
'    If mblnModify = False Then
'        txt(0).Locked = True
'        txt(1).Locked = True
'        txt(2).Locked = True
'        txt(3).Locked = True
'        txt(4).Locked = True
'        txt(5).Locked = True
'        txt(7).Locked = True
'        txt(8).Locked = True
'        txt(9).Locked = True
'        txt(10).Locked = True
'
'        mnuFileSave.Visible = False
'        mnuFileRestore.Visible = False
'
'        mnuFile_1.Visible = False
'
'        tbrThis.Buttons("保存").Visible = False
'        tbrThis.Buttons("重填").Visible = False
'        tbrThis.Buttons("Split_1").Visible = False
'        tbrThis.Buttons("Split_2").Visible = False
'    End If
    
    InitData = True

    Exit Function

errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function ValidEdit() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  校验数据的有效性
    '返回:  True        数据有效
    '       False       数据无效
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long


    ValidEdit = True

End Function

'Private Function SaveEdit(ByRef lngKey As Long) As Boolean
'    '------------------------------------------------------------------------------------------------------------------
'    '功能:  保存数据
'    '返回:  True        保存成功
'    '       False       保存失败
'    '------------------------------------------------------------------------------------------------------------------
'    Dim blnTran As Boolean
'    Dim lngLoop As Long
'    Dim strSQL() As String
'    Dim rsPati As New ADODB.Recordset
'
'    On Error GoTo errHand
'
'    ReDim Preserve strSQL(1 To 1)
'
'    gstrSQL = "SELECT * FROM 合约单位 WHERE ID=" & lngKey
'    Call OpenRecord(rsPati, gstrSQL, Me.Caption)
'    If rsPati.BOF = False Then
'        '存在
''        ID_IN IN 合约单位.ID%TYPE,
''        上级ID_IN IN 合约单位.上级ID%TYPE,
''        编码_IN IN 合约单位.编码%TYPE,
''        名称_IN IN 合约单位.名称%TYPE,
''        简码_IN IN 合约单位.简码%TYPE,
''        地址_IN IN 合约单位.地址%TYPE := NULL,
''        电话_IN IN 合约单位.电话%TYPE := NULL,
''        开户银行_IN IN 合约单位.开户银行%TYPE := NULL,
''        帐号_IN IN 合约单位.帐号%TYPE := NULL,
''        联系人_IN IN 合约单位.联系人%TYPE := NULL,
''        原长度_IN IN PLS_INTEGER,
''        电子邮件_IN IN 合约单位.电子邮件%TYPE := NULL,
''        说明_IN IN 合约单位.说明%TYPE := NULL
'
'        gstrSQL = "zl_合约单位_Update(" & lngKey & "," & _
'                                        IIf(IsNull(rsPati("上级ID").Value), "NULL", rsPati("上级ID").Value) & ",'" & _
'                                        txt(1).Text & "','" & _
'                                        txt(0).Text & "','" & _
'                                        txt(2).Text & "','" & _
'                                        txt(10).Text & "','" & _
'                                        txt(7).Text & "','" & _
'                                        txt(8).Text & "','" & _
'                                        txt(9).Text & "','" & _
'                                        txt(3).Text & "',0,'" & txt(5).Text & "','" & txt(4).Text & "')"
'        strSQL(ReDimArray(strSQL)) = gstrSQL
'    Else
'        '不存在
'    '    ID_IN IN 合约单位.ID%TYPE,
'    '    上级ID_IN IN 合约单位.上级ID%TYPE,
'    '    编码_IN IN 合约单位.编码%TYPE,
'    '    名称_IN IN 合约单位.名称%TYPE,
'    '    简码_IN IN 合约单位.简码%TYPE := NULL,
'    '    地址_IN IN 合约单位.地址%TYPE := NULL,
'    '    电话_IN IN 合约单位.电话%TYPE := NULL,
'    '    开户银行_IN IN 合约单位.开户银行%TYPE := NULL,
'    '    帐号_IN IN 合约单位.帐号%TYPE := NULL,
'    '    联系人_IN IN 合约单位.联系人%TYPE := NULL,
'    '    末级_IN IN 合约单位.末级%TYPE := 1,
''        电子邮件_IN IN 合约单位.电子邮件%TYPE := NULL,
''        说明_IN IN 合约单位.说明%TYPE := NULL
'        lngKey = zlDatabase.GetNextId("合约单位")
'        gstrSQL = "zl_合约单位_Insert(" & lngKey & "," & _
'                                        "NULL,'" & _
'                                        txt(1).Text & "','" & _
'                                        txt(0).Text & "','" & _
'                                        txt(2).Text & "','" & _
'                                        txt(10).Text & "','" & _
'                                        txt(7).Text & "','" & _
'                                        txt(8).Text & "','" & _
'                                        txt(9).Text & "','" & _
'                                        txt(3).Text & "',1,'" & txt(5).Text & "','" & txt(4).Text & "')"
'        strSQL(ReDimArray(strSQL)) = gstrSQL
'    End If
'
'    blnTran = True
'    gcnOracle.BeginTrans
'    For lngLoop = 1 To UBound(strSQL)
'        If strSQL(lngLoop) <> "" Then Call ExecuteProc(strSQL(lngLoop), Me.Caption)
'    Next
'    gcnOracle.CommitTrans
'    blnTran = False
'
'    SaveEdit = True
'
'    Exit Function
'
'errHand:
'
'    If ErrCenter = 1 Then Resume
'    If blnTran Then gcnOracle.RollbackTrans
'
'End Function


'（３）窗体及其控件的事件处理******************************************************************************************

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    
    If ValidEdit = False Then Exit Sub

    mblnOK = True
    
    mrsGroup("名称").Value = txt(0).Text
    mrsGroup("编码").Value = txt(1).Text
    mrsGroup("简码").Value = txt(2).Text
    mrsGroup("联系人").Value = txt(3).Text
    mrsGroup("电话").Value = txt(7).Text
    mrsGroup("电子邮件").Value = txt(5).Text
    mrsGroup("开户银行").Value = txt(8).Text
    mrsGroup("帐号").Value = txt(9).Text
    mrsGroup("地址").Value = txt(10).Text
    mrsGroup("说明").Value = txt(4).Text
    
    EditChanged = False
    Unload Me

    
End Sub


Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If mblnDataChange Then
        Cancel = (MsgBox("数据必须保存后才生效，是否不保存就退出？", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo)
        If Cancel Then Exit Sub
    End If
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub cmdHelp_Click()
   Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
End Sub

Private Sub txt_Change(Index As Integer)
    EditChanged = True
End Sub

Private Sub txt_GotFocus(Index As Integer)
    zlControl.TxtSelAll txt(Index)
    Select Case Index
    Case 0, 3, 4, 10
        zlCommFun.OpenIme True
    End Select
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
        
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
        If Index = 2 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
    Select Case Index
    Case 0, 3, 4, 10
        zlCommFun.OpenIme False
    End Select
    
    If Index = 0 Then
        If InStr(txt(Index).Text, "'") = 0 Then txt(2).Text = zlGetSymbol(txt(Index).Text)
    End If
End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And txt(Index).Locked Then
        glngTXTProc = GetWindowLong(txt(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And txt(Index).Locked Then
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
End Sub


