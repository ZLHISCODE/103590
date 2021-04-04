VERSION 5.00
Begin VB.Form frmMedicalItemsEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "检查组合"
   ClientHeight    =   5685
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8865
   Icon            =   "frmMedicalItemsEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   8865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   7635
      TabIndex        =   6
      Top             =   1380
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7635
      TabIndex        =   5
      Top             =   510
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   7635
      TabIndex        =   4
      Top             =   45
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   555
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   7515
      Begin VB.TextBox txt 
         Height          =   300
         Left            =   735
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   165
         Width           =   6405
      End
      Begin VB.CommandButton cmd 
         Caption         =   "&P"
         Height          =   270
         Left            =   7170
         TabIndex        =   1
         Top             =   180
         Width           =   285
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "项目(&D)"
         Height          =   180
         Index           =   12
         Left            =   75
         TabIndex        =   3
         Top             =   225
         Width           =   630
      End
   End
   Begin VB.Frame Frame4 
      Height          =   5160
      Left            =   15
      TabIndex        =   7
      Top             =   465
      Width           =   7515
      Begin VB.CommandButton cmdMove 
         Height          =   345
         Index           =   1
         Left            =   6675
         Picture         =   "frmMedicalItemsEdit.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "向后移动"
         Top             =   150
         Width           =   345
      End
      Begin VB.CommandButton cmdMove 
         Height          =   345
         Index           =   0
         Left            =   7080
         Picture         =   "frmMedicalItemsEdit.frx":0159
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "向前移动"
         Top             =   150
         Width           =   345
      End
      Begin zl9Medical.VsfGrid vsf 
         Height          =   4545
         Left            =   75
         TabIndex        =   8
         Top             =   525
         Width           =   7350
         _ExtentX        =   12965
         _ExtentY        =   8017
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "组成项目(&M)"
         Height          =   180
         Index           =   14
         Left            =   420
         TabIndex        =   11
         Top             =   225
         Width           =   990
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   0
         Left            =   75
         Picture         =   "frmMedicalItemsEdit.frx":02A6
         Top             =   195
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmMedicalItemsEdit"
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
Private mblnChanged As Boolean
Private mstrName As String
Private Enum mCol
    中文名 = 1
    编码
    英文名
    类型
    长度
    小数
    单位
    
End Enum

Private Property Let EditChanged(ByVal vData As Boolean)
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '值域:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngSvrKey As Long
            
    txt.Locked = False
    cmd.Enabled = True
    
    If vData = False Then
        cmdOK.Tag = ""
    Else
        cmdOK.Tag = "Changed"
        txt.Locked = True
        cmd.Enabled = False
    End If
End Property

Private Property Get EditChanged() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '值域:
    '------------------------------------------------------------------------------------------------------------------
            
    EditChanged = (cmdOK.Tag = "Changed")
    
End Property


Public Function ShowEdit(ByVal frmMain As Object, ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  显示编辑窗体，是与调用窗体的接口函数
    '参数:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    mblnStartUp = True
    mblnOK = False
            
    mlngKey = lngKey
    
    Set mfrmMain = frmMain
    
    If InitData = False Then Exit Function
    
    If mlngKey > 0 Then
        Call ReadData(mlngKey)
    End If
            
    EditChanged = False
    
    Me.Show 1, frmMain
    
    ShowEdit = mblnOK
    
End Function

Private Function ReadData(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  读取数据
    '参数:  lngKey      体检类型序号
    '返回:  True        读取成功
    '       False       读取失败
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
        
    On Error GoTo errHand
    
    gstrSQL = "SELECT 名称 FROM 诊疗项目目录 WHERE ID=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
    If rs.BOF = False Then txt.Text = rs("名称").Value
    mstrName = txt.Text
    
    
    gstrSQL = "Select a.ID,a.编码,a.中文名,a.英文名,A.类型,A.长度,A.小数,A.单位 " & _
                    "From 诊治所见项目 a,病历元素目录 b,病历所见单 c where b.类型=-1 and  b.id=c.元素id and c.所见项id=a.id and c.行=[1] Order By c.控件号 "
        
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
    If rs.BOF = False Then
        Do While Not rs.EOF
            
            If Val(vsf.RowData(vsf.Rows - 1)) > 0 Then
                vsf.Rows = vsf.Rows + 1
            End If
            
            vsf.RowData(vsf.Rows - 1) = zlCommFun.NVL(rs("ID"), 0)
            vsf.TextMatrix(vsf.Rows - 1, mCol.编码) = zlCommFun.NVL(rs("编码"))
            vsf.TextMatrix(vsf.Rows - 1, mCol.中文名) = zlCommFun.NVL(rs("中文名"))
            vsf.TextMatrix(vsf.Rows - 1, mCol.英文名) = zlCommFun.NVL(rs("英文名"))
            vsf.TextMatrix(vsf.Rows - 1, mCol.类型) = zlCommFun.NVL(rs("类型"))
            vsf.TextMatrix(vsf.Rows - 1, mCol.长度) = zlCommFun.NVL(rs("长度"))
            vsf.TextMatrix(vsf.Rows - 1, mCol.小数) = zlCommFun.NVL(rs("小数"))
            vsf.TextMatrix(vsf.Rows - 1, mCol.单位) = zlCommFun.NVL(rs("单位"))
                        
            rs.MoveNext
        Loop
    End If
    
    ReadData = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Function

Private Function InitData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  初始化设置
    '返回:  True        初始化成功
    '       False       初始化失败
    '------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHand

    
    With vsf
        .Cols = 0
        .NewColumn "", 255, 4
        .NewColumn "中文名", 2700, 1, "...", 1
        .NewColumn "编码", 1200, 1
        .NewColumn "英文名", 900, 1
        .NewColumn "类型", 900, 1
        .NewColumn "长度", 600, 7
        .NewColumn "小数", 600, 7
        .NewColumn "单位", 900, 1
        
        .Body.ColHidden(mCol.类型) = True
        
        .FixedCols = 1
        
        .SelectMode = True
    End With
        
    InitData = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function


Private Function SaveEdit() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  保存数据
    '返回:  True        保存成功
    '       False       保存失败
    '------------------------------------------------------------------------------------------------------------------
    Dim blnTran As Boolean
    Dim lngLoop As Long
    Dim strSQL() As String
    Dim lngElementID As Long
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    ReDim Preserve strSQL(1 To 1)
    
    strSQL(ReDimArray(strSQL)) = "ZL_体检检查组合_DELETE(" & mlngKey & ")"
    
    gstrSQL = "Select ID From 病历元素目录 Where 类型=-1"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If rs.BOF = False Then
        lngElementID = rs("ID").Value
    Else
        lngElementID = zlDatabase.GetNextId("病历元素目录")
        strSQL(ReDimArray(strSQL)) = "ZL_病历元素_INSERT(-1," & lngElementID & ",'000000','体检检查对应','','宋体,9',1,Null,'00001')"
        
    End If
    
    For lngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(lngLoop)) > 0 Then
            
            gstrSQL = "ZL_所见单_SAVE("
            gstrSQL = gstrSQL & lngElementID & ","
            gstrSQL = gstrSQL & lngLoop & ","
            gstrSQL = gstrSQL & "'2',"
            gstrSQL = gstrSQL & "NULL,"
            gstrSQL = gstrSQL & mlngKey & ","
            gstrSQL = gstrSQL & "NULL,"
            gstrSQL = gstrSQL & "NULL,"
            gstrSQL = gstrSQL & "NULL,"
            gstrSQL = gstrSQL & "NULL,"
            gstrSQL = gstrSQL & "NULL,"
            gstrSQL = gstrSQL & Val(vsf.RowData(lngLoop)) & ","
            gstrSQL = gstrSQL & "NULL,"
            gstrSQL = gstrSQL & "NULL,"
            gstrSQL = gstrSQL & "'" & vsf.TextMatrix(lngLoop, mCol.单位) & "',"                   '单位
            gstrSQL = gstrSQL & "NULL)"
            
            strSQL(ReDimArray(strSQL)) = gstrSQL
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


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdMove_Click(Index As Integer)
    If Index = 0 Then
        If vsf.Row > 1 Then
            
            Call MoveItem(vsf.Row, -1)
            vsf.Row = vsf.Row - 1
            cmdOK.Tag = "Changed"
            
        End If
    ElseIf vsf.Row < vsf.Rows - 1 Then
        
        Call MoveItem(vsf.Row, 1)
        vsf.Row = vsf.Row + 1
        cmdOK.Tag = "Changed"
        
    End If
    
    vsf.ShowCell vsf.Row, vsf.Col
    vsf.SetFocus
End Sub

Private Function MoveItem(ByVal intCurRow As Integer, Optional ByVal intMove As Integer = 1) As Boolean
    '-----------------------------------------------------------------------------------------
    '功能:
    '参数:
    '-----------------------------------------------------------------------------------------
    Dim strTmp As String
    Dim intCol As Integer
    
    On Error GoTo errHand
    
    strTmp = CStr(vsf.RowData(intCurRow))
            
    vsf.RowData(intCurRow) = vsf.RowData(intCurRow + intMove)
    vsf.RowData(intCurRow + intMove) = Val(strTmp)
    
    For intCol = 0 To vsf.Cols - 1
        
        strTmp = vsf.TextMatrix(intCurRow, intCol)
        
        vsf.TextMatrix(vsf.Row, intCol) = vsf.TextMatrix(intCurRow + intMove, intCol)
        
        vsf.TextMatrix(intCurRow + intMove, intCol) = strTmp
        
    Next
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmdOK_Click()
       
    If EditChanged Then
    
        If SaveEdit Then
            mblnOK = True
            
            EditChanged = False
        End If
        
    End If
    
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If cmdOK.Tag <> "" Then
        Cancel = (MsgBox("新增或修改的数据必须保存后才生效，是否不保存就退出？", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo)
    End If
End Sub

Private Sub vsf_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Select Case Col
        Case mCol.中文名
            
            
            gstrSQL = "Select ID,上级id,0 As 末级,编码 ,名称 ,'' As 英文名,0 as 类型,0 As 长度,0 As 小数,'' As 单位 from 诊治所见分类 where 性质=4 Start With 上级id is null connect by prior id =上级id "
            
            gstrSQL = gstrSQL & " Union All Select A.ID,A.分类id As 上级id,1 As 末级,A.编码,A.中文名 As 名称,a.英文名,A.类型,A.长度,A.小数,A.单位 " & _
                    "From 诊治所见项目 A  "
                    
            gstrSQL = "Select * From (" & gstrSQL & ") A ORDER BY A.末级, A.编码"
            
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
            If ShowGrdSelect(Me, vsf, "编码,1200,0,1;名称,2700,0,0;单位,900,0,0;英文名,900,0,0;类型,900,0,0", Me.Name & "\检查项目选择", "请从列表中选择一个检查项目。", rsData, rs, 8790, 5100) Then
                
                If CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                    ShowSimpleMsg "选择的项目“" & zlCommFun.NVL(rs("名称").Value) & "”已被选择！"
                    Exit Sub
                End If
                
                vsf.EditText = zlCommFun.NVL(rs("名称").Value)
                vsf.TextMatrix(Row, mCol.编码) = zlCommFun.NVL(rs("编码").Value)
                vsf.TextMatrix(Row, mCol.类型) = zlCommFun.NVL(rs("类型").Value)
                vsf.TextMatrix(Row, mCol.单位) = zlCommFun.NVL(rs("单位").Value)
                vsf.TextMatrix(Row, mCol.长度) = zlCommFun.NVL(rs("长度").Value)
                vsf.TextMatrix(Row, mCol.小数) = zlCommFun.NVL(rs("小数").Value)
                vsf.TextMatrix(Row, mCol.中文名) = zlCommFun.NVL(rs("名称").Value)
                vsf.TextMatrix(Row, mCol.英文名) = zlCommFun.NVL(rs("英文名").Value)
                vsf.RowData(Row) = zlCommFun.NVL(rs("ID").Value)
                
                EditChanged = True
                
            End If

    End Select
End Sub

Private Sub vsf_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, ByVal ComboList As String, KeyCode As Integer, ByVal Shift As Integer, Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim strText As String
    Dim strTmp As String
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
                Case mCol.中文名
                    
                    gstrSQL = "Select A.ID,1 As 末级,A.编码,A.中文名 As 名称,a.英文名,A.类型,A.长度,A.小数,A.单位 " & _
                    "From 诊治所见项目 A Where 分类id In (Select ID from 诊治所见分类 where 性质=4) And (编码 Like [1] Or Upper(中文名) Like [2] Or Upper(英文名) Like [2])"
                               
                    strText = UCase(vsf.EditText) & "%"
                    If ParamInfo.项目输入匹配方式 = 0 Then strTmp = "%" & strText
                                
                    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strText, strTmp)
                    If ShowGrdFilter(Me, vsf, "编码,1200,0,1;名称,2700,0,0;单位,900,0,0;英文名,900,0,0;类型,900,0,0", Me.Name & "\检查项目过滤", "请从列表中选择一个检查项目。", rsData, rs, 8790, 5100) Then
                        
                        
                        If CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                            ShowSimpleMsg "选择的项目“" & zlCommFun.NVL(rs("名称").Value) & "”已被选择！"
                            Exit Sub
                        End If
                           
                        vsf.EditText = zlCommFun.NVL(rs("名称").Value)
                        vsf.TextMatrix(Row, mCol.编码) = zlCommFun.NVL(rs("编码").Value)
                        vsf.TextMatrix(Row, mCol.类型) = zlCommFun.NVL(rs("类型").Value)
                        vsf.TextMatrix(Row, mCol.单位) = zlCommFun.NVL(rs("单位").Value)
                        vsf.TextMatrix(Row, mCol.长度) = zlCommFun.NVL(rs("长度").Value)
                        vsf.TextMatrix(Row, mCol.小数) = zlCommFun.NVL(rs("小数").Value)
                        vsf.TextMatrix(Row, mCol.中文名) = zlCommFun.NVL(rs("名称").Value)
                        vsf.TextMatrix(Row, mCol.英文名) = zlCommFun.NVL(rs("英文名").Value)
                        vsf.RowData(Row) = zlCommFun.NVL(rs("ID").Value)
                        
                        EditChanged = True
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
        EditChanged = True
    End If
End Sub
