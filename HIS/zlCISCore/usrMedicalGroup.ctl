VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl usrMedicalGroup 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   4005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7200
   ScaleHeight     =   4005
   ScaleWidth      =   7200
   Begin zl9CISCore.VsfGrid vsf 
      Height          =   1695
      Left            =   90
      TabIndex        =   4
      Top             =   465
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   2990
   End
   Begin VB.Frame fra 
      BackColor       =   &H80000005&
      Height          =   2325
      Index           =   0
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   7110
      Begin VB.PictureBox pic 
         BorderStyle     =   0  'None
         Height          =   330
         Index           =   0
         Left            =   30
         ScaleHeight     =   330
         ScaleWidth      =   5550
         TabIndex        =   1
         Top             =   120
         Width           =   5550
         Begin VB.CheckBox chk 
            Caption         =   "所有项目评估"
            Height          =   210
            Left            =   2670
            TabIndex        =   9
            Top             =   90
            Width           =   1830
         End
         Begin MSComctlLib.Toolbar cbr 
            Height          =   330
            Index           =   0
            Left            =   780
            TabIndex        =   2
            Top             =   0
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   582
            ButtonWidth     =   1349
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Wrappable       =   0   'False
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "ils16"
            HotImageList    =   "ils16"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   2
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "评估"
                  Key             =   "评估"
                  Object.ToolTipText     =   "按评估规则评估结论"
                  Object.Tag             =   "评估"
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "清空"
                  Key             =   "清空"
                  Object.ToolTipText     =   "清空下面的所有结论"
                  Object.Tag             =   "清空"
                  ImageIndex      =   1
               EndProperty
            EndProperty
         End
         Begin VB.Label picTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "结论描述"
            Height          =   180
            Index           =   0
            Left            =   60
            TabIndex        =   3
            Top             =   75
            Width           =   720
         End
      End
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrMedicalGroup.ctx":0000
            Key             =   "cls"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrMedicalGroup.ctx":0296
            Key             =   "search"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrMedicalGroup.ctx":6AF8
            Key             =   "new"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrMedicalGroup.ctx":D35A
            Key             =   "newadvice"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrMedicalGroup.ctx":13BBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrMedicalGroup.ctx":1A41E
            Key             =   "SelAll"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrMedicalGroup.ctx":20C80
            Key             =   "SelDel"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fra 
      BackColor       =   &H80000009&
      Height          =   1290
      Index           =   1
      Left            =   105
      TabIndex        =   5
      Top             =   2385
      Width           =   6960
      Begin VB.TextBox rtb 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   765
         Left            =   195
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   10
         Top             =   465
         Width           =   1170
      End
      Begin VB.PictureBox pic 
         BorderStyle     =   0  'None
         Height          =   330
         Index           =   1
         Left            =   15
         ScaleHeight     =   330
         ScaleWidth      =   4695
         TabIndex        =   6
         Top             =   105
         Width           =   4695
         Begin MSComctlLib.Toolbar cbr 
            Height          =   330
            Index           =   1
            Left            =   570
            TabIndex        =   7
            Top             =   15
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   582
            ButtonWidth     =   1349
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Wrappable       =   0   'False
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "ils16"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   2
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "生成"
                  Key             =   "生成"
                  Object.ToolTipText     =   "按上面的结论生成缺省建议"
                  Object.Tag             =   "生成"
                  ImageKey        =   "newadvice"
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "清空"
                  Key             =   "清空"
                  Object.ToolTipText     =   "清空下面的建议内容"
                  Object.Tag             =   "清空"
                  ImageKey        =   "cls"
               EndProperty
            EndProperty
         End
         Begin VB.Label picTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "建议"
            Height          =   180
            Index           =   1
            Left            =   75
            TabIndex        =   8
            Top             =   75
            Width           =   360
         End
      End
   End
End
Attribute VB_Name = "usrMedicalGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mstr挂号单 As String                    '外界传入
Private mlng病历id As Long                      '外界传入
Private mlng医嘱id As Long                      '外界传入

Private mblnMode As Boolean '为真是表示是用户进行的编辑，这时才赋值
Private mDispMode As Boolean
Private mReturnErrnumber As Long
Private mReturnErrDescription As String

Private mobjParentObject As Object

Private Enum mCol
    结论描述 = 1
    异常结果
    疾病
    诊断建议
End Enum

Private Function InitControl() As Boolean
    
    With vsf
    
        .Cols = 0
        .NewColumn "", 255
        .NewColumn "结论描述", 2400, 1, "...", 1
        .NewColumn "异常结果", 3000, 1, , 1
        .NewColumn "疾病", 600, 1, , 1
        .NewColumn "诊断建议", 900, 1, , 1
        .FixedCols = 1
        
        .ColDataType(mCol.疾病) = flexDTBoolean
        
        .TextMatrix(1, mCol.结论描述) = "未见异常"
        
        .Body.Appearance = flexXPThemes
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
                
    End With

End Function

Public Property Set ParentObject(vData As Object)
    Set mobjParentObject = vData
End Property

Public Property Get ParentObject() As Object
    Set ParentObject = mobjParentObject
End Property

Private Property Let Modified(vData As Boolean)
    
    On Error Resume Next
    
    If mobjParentObject Is Nothing Then Exit Property
    
    mobjParentObject.Modified = vData
    
End Property

Private Property Get Modified() As Boolean
    
    On Error Resume Next
    
    If mobjParentObject Is Nothing Then Exit Property
    
    Modified = mobjParentObject.Modified
    
End Property

'公共方法、属性
Public Sub SetgcnOracle()
    '------------------------------------------------------------------------------------------------------------------
    '接口过程
    '------------------------------------------------------------------------------------------------------------------
    Call InitCommon(gcnOracle)
    
End Sub

Public Property Get DispMode() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '接口过程:是否为显示模式
    '------------------------------------------------------------------------------------------------------------------
    DispMode = mDispMode
End Property

Public Property Let DispMode(ByVal New_DispMode As Boolean)
    '------------------------------------------------------------------------------------------------------------------
    '
    '------------------------------------------------------------------------------------------------------------------
    mDispMode = New_DispMode
    
    ShowUsrControl mlng医嘱id, Not mDispMode
    PropertyChanged "DispMode"
    
    If mDispMode Then
        vsf.Body.Editable = flexEDNone
        
        rtb.Locked = True
        
        cbr(0).Buttons("评估").Enabled = False
        cbr(0).Buttons("清空").Enabled = False
        cbr(0).Buttons("全选").Enabled = False
        cbr(0).Buttons("全清").Enabled = False
        
        cbr(1).Buttons("生成").Enabled = False
        cbr(1).Buttons("清空").Enabled = False
        
        cbr(0).Visible = False
        cbr(1).Visible = False
    Else
        cbr(0).Visible = True
        cbr(1).Visible = True
    End If
    
End Property

Public Property Let 挂号单(ByVal New_挂号单 As String)
    '------------------------------------------------------------------------------------------------------------------
    '设置挂号单
    '------------------------------------------------------------------------------------------------------------------
    
    mstr挂号单 = New_挂号单
    
End Property

Public Property Get ID病人病历() As Long
    '------------------------------------------------------------------------------------------------------------------
    '返回病人病历ID
    '------------------------------------------------------------------------------------------------------------------
    
    ID病人病历 = mlng病历id
    
End Property

Public Property Let ID病人病历(ByVal New_ID病人病历 As Long)
    '------------------------------------------------------------------------------------------------------------------
    '设置病人病历ID,并检查该病历是不是存在
    '------------------------------------------------------------------------------------------------------------------
    
    mlng病历id = New_ID病人病历
    ShowUsrControl mlng医嘱id, Not mDispMode
    
End Property

Public Sub SetDiagItem(ByVal New_医嘱ID As Long, ByVal New_发送号)
    '------------------------------------------------------------------------------------------------------------------
    '
    '------------------------------------------------------------------------------------------------------------------
    
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    mlng医嘱id = New_医嘱ID
    
End Sub

Public Property Get Get医嘱id() As Long
    '------------------------------------------------------------------------------------------------------------------
    '
    '------------------------------------------------------------------------------------------------------------------
    
    Get医嘱id = mlng医嘱id
        
End Property

Public Property Get Text() As String
    '------------------------------------------------------------------------------------------------------------------
    '为每一个控件加上文本转储属性
    '------------------------------------------------------------------------------------------------------------------
    
    Dim lngLoop As Long
    Dim strTmp As String
    Dim intCount As Integer
    
    On Error GoTo errHand
    
    '转储结论记录
    strTmp = strTmp & "一、结论：" & vbCrLf
    intCount = 0
    For lngLoop = 1 To vsf.Rows - 1

        If vsf.TextMatrix(lngLoop, mCol.结论描述) <> "" Then
            intCount = intCount + 1
            strTmp = strTmp & intCount & "、" & vsf.TextMatrix(lngLoop, mCol.结论描述) & vbCrLf
        End If
        
    Next
    strTmp = strTmp & vbCrLf
    
    '转储建议内容
    strTmp = strTmp & "二、建议：" & vbCrLf
    strTmp = strTmp & rtb.Text
    
    Text = strTmp
    
    Exit Property
    
errHand:
    
End Property

Public Sub ClearData()
    '------------------------------------------------------------------------------------------------------------------
    '功能:接口
    '参数:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    vsf.Rows = 2
    vsf.RowData(1) = 0
    vsf.Cell(flexcpText, 1, 1, vsf.Rows - 1, vsf.Cols - 1) = ""
    
    rtb.Text = ""
End Sub

Public Function SaveData(lng病人ID As Long, lng主页ID As Long, lng病历ID As Long, strReturnSQL As String, strError As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:接口
    '参数:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim strTmp As String
    Dim strSQL() As String
    Dim intCount As Integer
    Dim rs As New ADODB.Recordset
    Dim blnFlag As Boolean
    
    On Error GoTo errHand
    
    blnFlag = False
    gstrSql = "Select 参数值 From 系统参数表 Where 参数号=[1]"
    Set rs = OpenSQLRecord(gstrSql, "体检小结", 131)
    If rs.BOF = False Then
        blnFlag = (Val(zlCommFun.NVL(rs("参数值").Value, "0")) = 1)
    End If
    
    If blnFlag Then
        If MsgBox("请检查小结或体检小结是否填写！", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then
'            strError = "请检查小结或体检小结是否填写！"
            Exit Function
        End If
    End If
    
    For lngLoop = 1 To vsf.Rows - 1
        If StrIsValid(vsf.TextMatrix(lngLoop, 1), 100) = False Then
            vsf.Row = lngLoop
            vsf.Col = 1
            vsf.ShowCell vsf.Row, vsf.Col
            Exit Function
        End If
    Next
    
    If StrIsValid(rtb.Text, 4000) = False Then
        rtb.SetFocus
        Exit Function
    End If
    
    ReDim Preserve strSQL(0 To vsf.Rows + 1)
    
    strSQL(0) = "ZL_体检人员结论_DELETE(" & lng病历ID & ")"
    intCount = 0
    For lngLoop = 1 To vsf.Rows - 1
        If Trim(vsf.TextMatrix(lngLoop, mCol.结论描述)) <> "" Then
            
            intCount = intCount + 1
            
            strSQL(lngLoop) = "ZL_体检人员结论_INSERT(" & lng病人ID & "," & _
                                                        lng主页ID & "," & _
                                                        lng病历ID & "," & _
                                                        "0," & _
                                                        intCount & ",'" & _
                                                        vsf.TextMatrix(lngLoop, mCol.结论描述) & "','" & _
                                                        vsf.TextMatrix(lngLoop, mCol.异常结果) & "'," & _
                                                        "NULL," & _
                                                        Val(vsf.RowData(lngLoop)) & ",NULL,NULL," & _
                                                        Abs(Val(vsf.TextMatrix(lngLoop, mCol.疾病))) & "," & _
                                                        "'" & vsf.TextMatrix(lngLoop, mCol.诊断建议) & "')"
        End If
    Next
    
    strSQL(lngLoop + 1) = "ZL_体检人员结论_INSERT(" & lng病人ID & "," & _
                                                        lng主页ID & "," & _
                                                        lng病历ID & "," & _
                                                        "1," & _
                                                        "1," & _
                                                        "NULL,NULL,'" & _
                                                        rtb.Text & "'," & _
                                                        "NULL," & _
                                                        "NULL," & _
                                                        "NULL," & _
                                                        "0," & _
                                                        "NULL)"
        
    strTmp = ""
    For lngLoop = 0 To UBound(strSQL)
        If strSQL(lngLoop) <> "" Then
            
            strSQL(lngLoop) = Replace(strSQL(lngLoop), Chr(9), Chr(32))
            
            If strTmp = "" Then
                strTmp = strSQL(lngLoop)
            Else
                strTmp = strTmp & Chr(9) & strSQL(lngLoop)
            End If
        End If
    Next
    
    '返回SQL语句
    strReturnSQL = strTmp
    
    SaveData = True
    
    Exit Function
    
errHand:

    strError = "体检专用纸保存失败！"
    
End Function

Private Function ShowOpenList(Optional strText As String, Optional ByVal bytMode As Byte = 1) As Byte
    '------------------------------------------------------------------------------------------------------------------
    '功能:打开列表结构的诊疗检验标本数据
    '返回:出错返回2;成功返回1;取消返回0
    '------------------------------------------------------------------------------------------------------------------
    Dim strLvw As String
    Dim strTitle As String
    Dim strDescrible As String
    
    Dim sglX As Single
    Dim sglY As Single
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHand
    
    ShowOpenList = 2
    
    strText = "'%" & UCase(strText) & "%'"
    
    If bytMode = 1 Then
        
        strLvw = "编码,900,0,1;名称,1800,0,0;诊断建议,2700,0,0"
        strTitle = "体检结论过滤"
        strDescrible = "请从下表中选择一个体检结论"
        
        strSQL = _
                    "SELECT A.序号 AS ID, " & _
                            "A.编码, " & _
                            "A.名称, " & _
                            "A.是否疾病,A.诊断建议 " & _
                    "FROM 体检诊断建议 A " & _
                    "WHERE NVL(末级,0)=1 "
        strSQL = strSQL & " AND (A.编码 Like " & strText & " OR A.名称 Like " & strText & " OR A.简码 Like " & UCase(strText) & ")"
    End If
    
    Call OpenRecord(rs, strSQL, "体检结论")
    If rs.BOF Then
        ShowOpenList = 0
        Exit Function
    End If
    
    If rs.RecordCount = 1 Then GoTo Over
    Call CalcPosition(sglX, sglY, vsf)
    If frmSelectList.ShowSelect(Me, rs, strLvw, sglX + 60, sglY, 9000, 5100, strTitle, strDescrible) Then
        GoTo Over
    End If
    
    Exit Function
    
Over:
    If bytMode = 1 Then
        vsf.RowData(vsf.Row) = zlCommFun.NVL(rs("ID").Value)
        vsf.EditText = zlCommFun.NVL(rs("名称").Value)
        vsf.Cell(flexcpData, vsf.Row, vsf.Col) = zlCommFun.NVL(rs("名称").Value)
        vsf.TextMatrix(vsf.Row, vsf.Col) = zlCommFun.NVL(rs("名称").Value)
        vsf.TextMatrix(vsf.Row, mCol.疾病) = zlCommFun.NVL(rs("是否疾病").Value, 0)
        vsf.TextMatrix(vsf.Row, mCol.诊断建议) = zlCommFun.NVL(rs("诊断建议").Value)
    End If
    
    Modified = True
    
    ShowOpenList = 1
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function


Private Function ShowOpenTree(Optional ByVal bytMode As Byte = 1) As Byte
    '------------------------------------------------------------------------------------------------------------------
    '功能:打开列表结构的诊疗检验标本数据
    '返回:出错返回2;成功返回1;取消返回0
    '------------------------------------------------------------------------------------------------------------------
    Dim strLvw As String
    Dim strTitle As String
    Dim strDescrible As String
    
    Dim sglX As Single
    Dim sglY As Single
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHand
    
    ShowOpenTree = 2
    
    If bytMode = 1 Then
        strLvw = "编码,900,0,1;名称,1800,0,0;诊断建议,2700,0,0"
        strTitle = "体检结论选择"
        strDescrible = "请从下表中选择一个体检诊断"
        
        strSQL = "SELECT -1 AS ID," & _
                            "0 AS 上级ID," & _
                            "0 AS 末级," & _
                            "'' AS 编码," & _
                            "'所有分类' AS 名称, " & _
                            "Null+0 AS 是否疾病,'' As 诊断建议 " & _
                    "FROM dual "
                    
        strSQL = strSQL & _
                " UNION ALL " & _
                "SELECT 序号 AS ID," & _
                            "DECODE(上级序号,NULL,-1,上级序号) AS 上级ID," & _
                            "0 AS 末级," & _
                            "编码," & _
                            "名称, " & _
                            "Null+0 AS 是否疾病,'' As 诊断建议 " & _
                    "FROM 体检诊断建议 " & _
                    "WHERE NVL(末级,0)=0 " & _
                    "START WITH 上级序号 is NULL CONNECT BY PRIOR 序号 = 上级序号 "
        
        strSQL = strSQL & _
                    "UNION ALL " & _
                    "SELECT A.序号 AS ID, " & _
                            "DECODE(上级序号,NULL,-1,上级序号) AS 上级ID, " & _
                            "1 AS 末级, " & _
                            "A.编码, " & _
                            "A.名称, " & _
                            "A.是否疾病,A.诊断建议 " & _
                    "FROM 体检诊断建议 A " & _
                    "WHERE NVL(A.末级,0)=1"
    End If
    
    Call OpenRecord(rs, strSQL, "体检结论")
    
    If rs.BOF Then
        ShowOpenTree = 0
        Exit Function
    End If
        
    Call CalcPosition(sglX, sglY, vsf)
    
    If frmSelectTree.ShowSelect(Screen, rs, sglX, sglY, 9000, 5100, vsf.CellHeight, strTitle, strLvw, strDescrible) Then
        GoTo Over
    End If
    
    Exit Function
    
Over:
    If bytMode = 1 Then
    
        vsf.RowData(vsf.Row) = zlCommFun.NVL(rs("ID").Value)
        vsf.EditText = zlCommFun.NVL(rs("名称").Value)
        vsf.Cell(flexcpData, vsf.Row, vsf.Col, vsf.Row, vsf.Col) = zlCommFun.NVL(rs("名称").Value)
        vsf.TextMatrix(vsf.Row, vsf.Col) = zlCommFun.NVL(rs("名称").Value)
        
        vsf.TextMatrix(vsf.Row, mCol.疾病) = zlCommFun.NVL(rs("是否疾病").Value, 0)
        vsf.TextMatrix(vsf.Row, mCol.诊断建议) = zlCommFun.NVL(rs("诊断建议").Value)
        
    End If
    
    Modified = True
    
    ShowOpenTree = 1
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub CalcPosition(ByRef x As Single, ByRef y As Single, ByVal objBill As Object)
    '------------------------------------------------------------------------------------------------------------------
    '功能： 计算X,Y的实际坐标，并考虑屏幕超界的问题
    '参数： X---返回横坐标参数
    '       Y---返回纵坐标参数
    '------------------------------------------------------------------------------------------------------------------
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(objBill.hWnd, objPoint)
    
    x = objPoint.x * 15 + objBill.CellLeft - 45
    y = objPoint.y * 15 + objBill.CellTop + objBill.CellHeight - 30
End Sub

Private Function GetAdvice() As String
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '------------------------------------------------------------------------------------------------------------------
    
    Dim rs As New ADODB.Recordset
    Dim lngLoop As Long
    Dim strSQL As String
        
    On Error GoTo errHand
    
    GetAdvice = ""
    
    For lngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(lngLoop)) > 0 Then
            
            strSQL = "SELECT 参考建议 FROM 体检诊断建议 WHERE 序号 = " & Val(vsf.RowData(lngLoop))
            Call OpenRecord(rs, strSQL, "体检结论")
            If rs.BOF = False Then
                
                If zlCommFun.NVL(rs("参考建议").Value) <> "" Then
                    If vsf.TextMatrix(lngLoop, mCol.异常结果) <> "" Then
                        GetAdvice = GetAdvice & lngLoop & "." & vsf.TextMatrix(lngLoop, mCol.结论描述) & " {" & vsf.TextMatrix(lngLoop, mCol.异常结果) & "}：" & vbCrLf
                    Else
                        GetAdvice = GetAdvice & lngLoop & "." & vsf.TextMatrix(lngLoop, mCol.结论描述) & "：" & vbCrLf
                    End If

                    GetAdvice = GetAdvice & zlCommFun.NVL(rs("参考建议").Value) & vbCrLf & vbCrLf
                End If
                
            End If
            
        End If
    Next
    
    Exit Function
    
errHand:
        
End Function

Private Sub SetErr(lngErrNum As Long, strErr As String)
    '------------------------------------------------------------------------------------------------------------------
    '设置错误描述及错误号
    '如果lngErrNum=-1 表示 控件自己定义的错误
    '------------------------------------------------------------------------------------------------------------------
    
    mReturnErrnumber = lngErrNum
    mReturnErrDescription = strErr
End Sub

Private Function InDesign() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：判断当前运行程序是否在VB的工程环境中
    '------------------------------------------------------------------------------------------------------------------
    
    On Error Resume Next
    
    Debug.Print 1 / 0
    If Err.Number <> 0 Then Err.Clear: InDesign = True
    
End Function

Private Sub ShowUsrControl(lngKey As Long, Optional ByVal blnEditMode As Boolean = False)
    '------------------------------------------------------------------------------------------------------------------
    '功能：外部调用显示
    '------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    Dim intRow As Integer
    Dim blnSave As Boolean
    
    
    On Error GoTo errHand
        
    mDispMode = Not blnEditMode
    
    'Begin  <初始化处理>
    blnSave = Modified
    
    If gcnOracle Is Nothing Then SetErr -1, "连接对象没有初始化": Exit Sub
    If gcnOracle.State <> adStateOpen Then SetErr -1, "连接对象没有连接": Exit Sub
    
    'End    <初始化处理>


    'Begin  <读取数据>
    
    Call InitControl
    vsf.ExtendLastCol = True
    
    intRow = 0
    
    strSQL = "SELECT A.* FROM 体检人员结论 A WHERE A.病历id=" & mlng病历id & " ORDER BY A.记录性质,A.记录序号"
    Call OpenRecord(rs, strSQL, "体检结论")
    If rs.BOF = False Then
        Do While Not rs.EOF
            
            If zlCommFun.NVL(rs("记录性质").Value) = 0 Then
                
                intRow = intRow + 1
                vsf.Rows = intRow + 1
                
                vsf.RowData(intRow) = zlCommFun.NVL(rs("结论id").Value)
                
                vsf.TextMatrix(intRow, mCol.结论描述) = zlCommFun.NVL(rs("结论描述").Value)
                vsf.TextMatrix(intRow, mCol.异常结果) = zlCommFun.NVL(rs("异常结果").Value)
                vsf.TextMatrix(intRow, mCol.诊断建议) = zlCommFun.NVL(rs("诊断建议").Value)
                vsf.TextMatrix(intRow, mCol.疾病) = zlCommFun.NVL(rs("是否疾病").Value)
                
            Else
                rtb.Text = zlCommFun.NVL(rs("参考建议").Value)
            End If
            
            rs.MoveNext
        Loop
    End If
    
    'End    <读取数据>

    Call UserControl_Resize
    
    Modified = blnSave
    
    Exit Sub
    
errHand:

    If Ambient.UserMode = False Or InDesign = False Then
        SetErr Err.Number, Err.Description
        Exit Sub
    End If
    
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function EditRefresh(ByVal objVsf As Object) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim LngCount As Long
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    If MsgBox("是否要替换原来的体检结论？", vbYesNo + vbQuestion + vbDefaultButton1, gstrSysName) = vbYes Then
        vsf.Rows = 2
        vsf.RowData(1) = 0
        vsf.Cell(flexcpText, 1, 1, vsf.Rows - 1, vsf.Cols - 1) = ""
        
    End If
    
    For lngLoop = 1 To objVsf.Rows - 1
        If Val(objVsf.RowData(lngLoop)) > 0 Then
            If Abs(Val(objVsf.TextMatrix(lngLoop, 0))) = 1 Then
                
                '检查Val(objVsf.RowData(lngLoop))是否已经存在
                For LngCount = 0 To vsf.Rows - 1
                    If Trim(vsf.TextMatrix(LngCount, 1)) = Trim(objVsf.TextMatrix(lngLoop, 1)) Then
                        GoTo NextLoop
                    End If
                Next
                
                If Val(vsf.RowData(vsf.Rows - 1)) > 0 Then vsf.Rows = vsf.Rows + 1
                
                vsf.RowData(vsf.Rows - 1) = Val(objVsf.RowData(lngLoop))
                vsf.TextMatrix(vsf.Rows - 1, 0) = vsf.Rows - 1 & "、"
                vsf.TextMatrix(vsf.Rows - 1, mCol.结论描述) = objVsf.TextMatrix(lngLoop, 1)
                vsf.TextMatrix(vsf.Rows - 1, mCol.异常结果) = objVsf.TextMatrix(lngLoop, 2)
                vsf.TextMatrix(vsf.Rows - 1, mCol.疾病) = Abs(Val(objVsf.Cell(flexcpData, lngLoop, 1, lngLoop, 1)))
                                
                gstrSql = "Select A.诊断建议 From 体检诊断建议 A where A.序号=" & Val(objVsf.RowData(lngLoop))

                Call OpenRecord(rs, gstrSql, "体检结论")
                If rs.BOF = False Then
                    vsf.TextMatrix(vsf.Rows - 1, mCol.诊断建议) = zlCommFun.NVL(rs("诊断建议"))
                End If
                
                
            End If
        End If
        
NextLoop:
        
    Next
    
    EditRefresh = True
    
    Exit Function
    
errHand:
    
End Function


Private Sub rtb_Change()
    
    Modified = True
    
End Sub

Private Sub cbr_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
    
    Dim lng病人ID As Long
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    If mDispMode Then Exit Sub
    
    Select Case Index
    Case 0
        Select Case Button.Key
            
        Case "评估"
            
            strSQL = "SELECT 病人id FROM 病人医嘱记录 WHERE ID=" & mlng医嘱id
            
            Call OpenRecord(rs, strSQL, "体检结论")
            If rs.BOF = False Then
                If chk.Value = 0 Then
                    Call frmMedicalResult.ShowEdit(Me, zlCommFun.NVL(rs("病人id"), 0) & "'" & mlng医嘱id & "'" & mstr挂号单)
                Else
                    Call frmMedicalResult.ShowEdit(Me, zlCommFun.NVL(rs("病人id"), 0) & "'0'" & mstr挂号单)
                End If
            End If
            
        Case "清空"
            
            vsf.Rows = 2
            vsf.RowData(1) = 0
            vsf.Cell(flexcpText, 1, 1, vsf.Rows - 1, vsf.Cols - 1) = ""
            
            Modified = True
            
        End Select
    Case 1
        Select Case Button.Key
        Case "生成"
            rtb.Text = GetAdvice
        Case "清空"
            rtb.Text = ""
        End Select
        
        Modified = True
    End Select
End Sub

Private Sub rtb_GotFocus()
    zlCommFun.OpenIme True
End Sub

Private Sub rtb_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub UserControl_Initialize()
    
    Call InitControl
    
End Sub

Private Sub UserControl_Resize()
    
    On Error Resume Next
        
    With fra(0)
        .Left = 0
        .Top = -90
        .Width = UserControl.Width
    End With
    
    With fra(1)
        .Left = 0
        .Top = fra(0).Top + fra(0).Height - 90
        .Width = fra(0).Width
        .Height = UserControl.Height + 90 - fra(0).Height - 90
    End With
            
    With pic(0)
        .Left = 30
        .Top = 120
        .Width = fra(0).Width - .Left - 45
    End With

    
    With vsf
        .Left = 15
        .Top = pic(0).Top + pic(0).Height - 90
        .Width = fra(0).Width - .Left - 30
    End With
       
    With pic(1)
        .Left = 30
        .Top = 120
        .Width = fra(1).Width - .Left - 45
    End With

                    
    With rtb
        .Left = 15
        .Top = pic(1).Top + pic(1).Height
        .Width = fra(1).Width - .Left - 30
        .Height = fra(1).Height - .Top - 30
    End With
        
End Sub

Private Sub UserControl_Terminate()
    On Error Resume Next
    
    Set mobjParentObject = Nothing
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Modified = True
End Sub

Private Sub vsf_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    
    If mDispMode Then Exit Sub
    
    Select Case Col
    Case mCol.结论描述
        
        Call ShowOpenTree(1)
        
    End Select
    
End Sub

Private Sub vsf_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, ByVal ComboList As String, KeyCode As Integer, ByVal Shift As Integer, Cancel As Boolean)
    Dim strSvrText As String

    If mDispMode Then Exit Sub

    If KeyCode = vbKeyReturn Then
        '对于2-文字型的情况

        If InStr(vsf.EditText, "'") > 0 Then
            KeyCode = 0
            Exit Sub
        End If
        
        If Col = mCol.结论描述 Then
            strSvrText = vsf.EditText
            Select Case ShowOpenList(vsf.EditText)
            Case 0
            
                '没有匹配的项目
                vsf.Cell(flexcpData, Row, Col) = strSvrText
    
            Case 1
                '选取了一个项目
    '            mblnChangeEdit = True
    '            Call AdjustEnableState
            Case 2
                '取消了本次选择
                KeyCode = 0
    
                vsf.Cell(flexcpData, Row, Col) = vsf.Cell(flexcpData, Row, Col)
                vsf.TextMatrix(Row, Col) = vsf.Cell(flexcpData, Row, Col)
    
            End Select
        End If
        
    Else
'        mblnChangeEdit = True
'        Call AdjustEnableState
    End If
End Sub
