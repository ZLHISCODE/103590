VERSION 5.00
Begin VB.Form frmMedicalItemsArrange 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "体检项目排列"
   ClientHeight    =   6135
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8910
   Icon            =   "frmMedicalItemsArrange.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame4 
      Height          =   6135
      Left            =   45
      TabIndex        =   3
      Top             =   -45
      Width           =   7515
      Begin VB.CommandButton cmdMove 
         Height          =   345
         Index           =   0
         Left            =   7095
         Picture         =   "frmMedicalItemsArrange.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "向前移动"
         Top             =   180
         Width           =   345
      End
      Begin VB.CommandButton cmdMove 
         Height          =   345
         Index           =   1
         Left            =   7095
         Picture         =   "frmMedicalItemsArrange.frx":0159
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "向后移动"
         Top             =   570
         Width           =   345
      End
      Begin zl9Medical.VsfGrid vsf 
         Height          =   5895
         Left            =   75
         TabIndex        =   4
         Top             =   150
         Width           =   6945
         _ExtentX        =   12250
         _ExtentY        =   10398
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   7650
      TabIndex        =   2
      Top             =   45
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7650
      TabIndex        =   1
      Top             =   510
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   7650
      TabIndex        =   0
      Top             =   1380
      Width           =   1100
   End
End
Attribute VB_Name = "frmMedicalItemsArrange"
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

Private Enum mCol
    名称 = 1
    编码
    单位
    类别
End Enum

Private Property Let EditChanged(ByVal vData As Boolean)
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '值域:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngSvrKey As Long
    
    If vData = False Then
        cmdOK.Tag = ""
    Else
        cmdOK.Tag = "Changed"
    
    End If
End Property

Private Property Get EditChanged() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '值域:
    '------------------------------------------------------------------------------------------------------------------
            
    EditChanged = (cmdOK.Tag = "Changed")
    
End Property


Public Function ShowEdit(ByVal frmMain As Object) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  显示编辑窗体，是与调用窗体的接口函数
    '参数:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    mblnStartUp = True
    mblnOK = False
                
    Set mfrmMain = frmMain
    
    If InitData = False Then Exit Function
    
    
    Call ReadData
    
            
    EditChanged = False
    
    Me.Show 1, frmMain
    
    ShowEdit = mblnOK
    
End Function

Private Function ReadData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  读取数据
    '参数:  lngKey      体检类型序号
    '返回:  True        读取成功
    '       False       读取失败
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
        
    On Error GoTo errHand
        
    
    gstrSQL = "Select a.ID,a.编码,a.名称,A.计算单位 As 单位,DECODE(A.类别,'C','检验','检查') AS 类别 " & _
                    "From 诊疗项目目录 a,体检项目排列 b where b.诊疗项目id=a.ID And b.排列性质=1  Order By b.排列顺序 "
        
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If rs.BOF = False Then
        Do While Not rs.EOF
            
            If Val(vsf.RowData(vsf.Rows - 1)) > 0 Then
                vsf.Rows = vsf.Rows + 1
            End If
            
            vsf.RowData(vsf.Rows - 1) = zlCommFun.NVL(rs("ID"), 0)
            vsf.TextMatrix(vsf.Rows - 1, mCol.编码) = zlCommFun.NVL(rs("编码"))
            vsf.TextMatrix(vsf.Rows - 1, mCol.名称) = zlCommFun.NVL(rs("名称"))
            vsf.TextMatrix(vsf.Rows - 1, mCol.单位) = zlCommFun.NVL(rs("单位"))
            vsf.TextMatrix(vsf.Rows - 1, mCol.类别) = zlCommFun.NVL(rs("类别"))
                        
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
        .NewColumn "名称", 2700, 1, "...", 1
        .NewColumn "编码", 1200, 1
        .NewColumn "单位", 900, 1
        .NewColumn "类别", 900, 1
        
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
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    ReDim Preserve strSQL(1 To 1)
    
    strSQL(ReDimArray(strSQL)) = "ZL_体检项目排列_DELETE(1)"
    
    For lngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(lngLoop)) > 0 Then
            
            gstrSQL = "ZL_体检项目排列_INSERT("
            gstrSQL = gstrSQL & Val(vsf.RowData(lngLoop)) & ","
            gstrSQL = gstrSQL & lngLoop & ",1)"
            
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
        Case mCol.名称
        
            gstrSQL = GetPublicSQL(SQL.体检项目选择)
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, 1, 2)
            If ShowGrdSelect(Me, vsf, "编码,1200,0,1;名称,2700,0,0;单位,900,0,0;类别,900,0,0", Me.Name & "\体检项目选择", "请从列表中选择一个体检项目。", rsData, rs, 8790, 4500) Then
                '选取了一个项目
                If CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                    ShowSimpleMsg "选择的项目“" & zlCommFun.NVL(rs("名称").Value) & "”已被选择！"
                    Exit Sub
                End If
                
                vsf.Cell(flexcpText, Row, mCol.名称 + 1, Row, vsf.Cols - 1) = ""
                
                vsf.EditText = zlCommFun.NVL(rs("名称").Value)
                vsf.TextMatrix(Row, mCol.名称) = zlCommFun.NVL(rs("名称").Value)
                vsf.TextMatrix(Row, mCol.编码) = zlCommFun.NVL(rs("编码").Value)
                vsf.TextMatrix(Row, mCol.单位) = zlCommFun.NVL(rs("单位").Value)
                vsf.TextMatrix(Row, mCol.类别) = zlCommFun.NVL(rs("类别").Value)
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
                Case mCol.名称
                    
                    strText = UCase(vsf.EditText)
                    
                    gstrSQL = GetPublicSQL(SQL.体检项目过滤选择, strText)
                    
                    strText = strText & "%"
                    If ParamInfo.项目输入匹配方式 = 0 Then strTmp = "%" & strText
                                
                    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "C", "D", strText, strTmp, 1, 2)
                    If ShowGrdFilter(Me, vsf, "编码,1200,0,1;名称,2700,0,0;单位,900,0,0;类别,900,0,0", Me.Name & "\体检项目过滤", "请从列表中选择一个体检项目。", rsData, rs, 8790, 5100) Then
                                                
                        If CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                            ShowSimpleMsg "选择的项目“" & zlCommFun.NVL(rs("名称").Value) & "”已被选择！"
                            Exit Sub
                        End If
                        
                        vsf.EditText = zlCommFun.NVL(rs("名称").Value)
                        vsf.TextMatrix(Row, mCol.编码) = zlCommFun.NVL(rs("编码").Value)
                        vsf.TextMatrix(Row, mCol.单位) = zlCommFun.NVL(rs("单位").Value)
                        vsf.TextMatrix(Row, mCol.名称) = zlCommFun.NVL(rs("名称").Value)
                        vsf.TextMatrix(Row, mCol.类别) = zlCommFun.NVL(rs("类别").Value)
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


