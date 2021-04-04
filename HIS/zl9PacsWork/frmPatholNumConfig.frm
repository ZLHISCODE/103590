VERSION 5.00
Begin VB.Form frmPatholNumConfig 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "病理号别配置"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   11955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdDel 
      Caption         =   "删除"
      Height          =   400
      Left            =   7680
      TabIndex        =   7
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "新增(&A)"
      Height          =   400
      Left            =   6240
      TabIndex        =   6
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保 存(&S)"
      Height          =   400
      Left            =   9120
      TabIndex        =   4
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取 消(&C)"
      Height          =   400
      Left            =   10560
      TabIndex        =   3
      Top             =   5640
      Width           =   1095
   End
   Begin VB.OptionButton optYear 
      Caption         =   "4位"
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   5730
      Value           =   -1  'True
      Width           =   615
   End
   Begin VB.OptionButton optMonth 
      Caption         =   "2位"
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   5730
      Width           =   615
   End
   Begin zl9PACSWork.ucFlexGrid ufgData 
      Height          =   4935
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   8705
      DefaultCols     =   ""
      GridRows        =   11
      HeadCheckValue  =   1
      IsCopyAdoMode   =   0   'False
      IsEjectConfig   =   -1  'True
      HeadFontCharset =   134
      HeadFontWeight  =   400
      HeadColor       =   0
      DataFontCharset =   134
      DataFontWeight  =   400
      DataColor       =   -2147483640
      GridLineColor   =   14737632
      ExtendLastCol   =   -1  'True
   End
   Begin VB.Label Label4 
      Caption         =   "规则中使用              年份"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   5760
      Width           =   2535
   End
End
Attribute VB_Name = "frmPatholNumConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const C_str提示 = "提示"

Private Function reInitData() As Boolean
    Dim strSql As String
    On Error GoTo errH
    reInitData = False
    gcnOracle.BeginTrans
    
    strSql = "ZL_病理号码规则_Insert(1,0,'CG',1,1,1,3,4,1,'常规')"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    strSql = "ZL_病理号码规则_Insert(2,1,'BD',1,1,1,3,4,1,'冰冻')"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    strSql = "ZL_病理号码规则_Insert(3,2,'XB',1,1,1,3,4,1,'细胞')"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    strSql = "ZL_病理号码规则_Insert(4,3,'HZ',1,1,1,3,4,1,'会诊')"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    strSql = "ZL_病理号码规则_Insert(5,4,'SJ',1,1,1,3,4,1,'尸检')"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    strSql = "ZL_病理号码规则_Insert(6,5,'KSSL',1,1,1,3,4,1,'快速石蜡')"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    gcnOracle.CommitTrans
    reInitData = True
    Exit Function
errH:
    Call gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
End Function


Private Sub LordPatholNumRules()
'载入病理号规则，显示在列表中
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim strWhere As String
    Dim curDate As Date
    Dim i As Integer
    Dim intID As Integer
    
    On Error GoTo errH
    
    strSql = "select ID,类型,前缀,年,月,日,年份位数,序号位数,起始数,名称  from 病理号码规则  order by ID"
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    Set ufgData.AdoData = rsData
    
    If rsData.RecordCount > 5 Then '号码规则数量正常
    
    ElseIf rsData.RecordCount = 0 Then '号码规则数量为0，需要初始化
        If reInitData() = False Then
            Call MsgBoxD(Me, "病理号码规则数据初始化失败，请联系软件维护人员解决", vbOKOnly, "病理号码配置")
            Exit Sub
        End If
    Else '
        Call MsgBoxD(Me, "病理号码规则数据异常，请联系软件维护人员解决", vbOKOnly, "病理号码配置")
        Exit Sub
    End If

    Call ufgData.RefreshData
    
    curDate = zlDatabase.Currentdate
    
    For i = 1 To ufgData.GridRows - 1
        If ufgData.Text(i, gstrPatholNumSet_ID) <> "" Then

            strSql = "select 当前序号 as 起始数 from 病理号码记录 where 号码规则ID=[1]"
            strWhere = ""
            
            If InStr(ufgData.Text(i, gstrPatholNumSet_年), "是") > 0 Then
                strWhere = strWhere & " and 年=[2]"
            End If

            If InStr(ufgData.Text(i, gstrPatholNumSet_月), "是") > 0 Then
                strWhere = strWhere & " and 月=[3]"
            End If

            If InStr(ufgData.Text(i, gstrPatholNumSet_日), "是") > 0 Then
                strWhere = strWhere & " and 日=[4]"
            End If
 
            Set rsData = zlDatabase.OpenSQLRecord(strSql & strWhere, Me.Caption, ufgData.Text(i, gstrPatholNumSet_ID), Val(Format(curDate, "yyyy")), Val(Format(curDate, "mm")), Val(Format(curDate, "dd")))
            
            If rsData.RecordCount = 0 Then
                intID = Val(ufgData.Text(i, gstrPatholNumSet_ID)) - 1
                strSql = "select 当前序号 as 起始数 from 病理号码记录 where 类型=[1]"
                Set rsData = zlDatabase.OpenSQLRecord(strSql & strWhere, Me.Caption, intID, Val(Format(curDate, "yyyy")), Val(Format(curDate, "mm")), Val(Format(curDate, "dd")))
                If rsData.RecordCount > 0 Then
                    ufgData.Text(i, gstrPatholNumSet_起始数) = Nvl(rsData!起始数, "0")
                    ufgData.Text(i, gstrPatholNumSet_起始数) = Val(ufgData.Text(i, gstrPatholNumSet_起始数)) + 1
                Else
                    ufgData.Text(i, gstrPatholNumSet_起始数) = 1
                End If
            
            Else
                If rsData.RecordCount > 0 Then
                    ufgData.Text(i, gstrPatholNumSet_起始数) = Nvl(rsData!起始数, "0")
                    ufgData.Text(i, gstrPatholNumSet_起始数) = Val(ufgData.Text(i, gstrPatholNumSet_起始数)) + 1
                Else
                    ufgData.Text(i, gstrPatholNumSet_起始数) = 1
                End If
            End If

        End If
    Next

    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub initUfgList()
'初始化列表
    
    On Error GoTo errH
    
    ufgData.IsKeepRows = False
    ufgData.RowHeightMin = glngStandardRowHeight
    
    ufgData.DefaultColNames = gstrPatholNumSetCols
    ufgData.ColNames = gstrPatholNumSetCols

    '禁止右键弹出列表配置窗口
    ufgData.IsEjectConfig = False
    '禁止弹出鼠标右键菜单
    ufgData.IsShowPopupMenu = False
    
    ufgData.ColConvertFormat = gstrPatholNumSetConvertFormat
    
    Exit Sub
errH:
    Call err.Raise(0, , "初始化列表失败")
End Sub

Private Sub cmdAdd_Click()
    Dim lngNewRow As Long
    
    On Error GoTo errH

    lngNewRow = ufgData.NewRow
    
    ufgData.Text(lngNewRow, gstrPatholNumSet_ID) = ""
    ufgData.Text(lngNewRow, gstrPatholNumSet_类型) = ""
    ufgData.Text(lngNewRow, gstrPatholNumSet_前缀) = ""
    ufgData.Text(lngNewRow, gstrPatholNumSet_年) = "1-是"
    ufgData.Text(lngNewRow, gstrPatholNumSet_月) = "1-是"
    ufgData.Text(lngNewRow, gstrPatholNumSet_日) = "1-是"
    ufgData.Text(lngNewRow, gstrPatholNumSet_年份位数) = IIf(optYear.value = True, "4", "2")
    ufgData.Text(lngNewRow, gstrPatholNumSet_序号位数) = 4
    ufgData.Text(lngNewRow, gstrPatholNumSet_名称) = ""
    ufgData.Text(lngNewRow, gstrPatholNumSet_起始数) = "1"
    
    Call ufgData_OnAfterEdit(lngNewRow, ufgData.DataGrid.Col)
    Exit Sub
errH:
    Call MsgBoxD(Me, err.Description, vbOKOnly, C_str提示)
End Sub

Private Function CheckDelable(ByVal lngID As Long) As Boolean
'检查是否有关联表使用本数据，用于判断是否可以删除
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    CheckDelable = False
    
    strSql = "select 号码规则ID from  病理检查信息 where 号码规则ID=[1] and rownum <2 "
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "查询号码规则是否使用", lngID)
    
    If rsData.RecordCount = 0 Then CheckDelable = True
        
End Function

Private Sub DelRow()
'删除行
    Dim lngRow As Long
    Dim strSql As String
    Dim intTMP As Integer
    
    On Error GoTo errH
    
    lngRow = ufgData.SelectionRow
    
    '若有ID 判断能否删除，没有ID说明是刚新建的，提示确认后删除
    If ufgData.Text(lngRow, gstrPatholNumSet_ID) <> "" Then
    
        intTMP = Val(ufgData.Text(lngRow, gstrPatholNumSet_ID))
        
        If intTMP >= 1 And intTMP <= 6 Then
            Call MsgBoxD(Me, "该号码规则属于基本规则,不可删除", vbOKOnly, C_str提示)
            Exit Sub
        Else
            If CheckDelable(Val(ufgData.Text(lngRow, gstrPatholNumSet_ID))) = True Then
            
                If MsgBoxD(Me, "请认真考虑是否删除该号码规则?", vbYesNo + vbDefaultButton2 + vbCritical, C_str提示) = vbNo Then Exit Sub
                
                strSql = "ZL_病理号码规则_Delete(" & ufgData.Text(lngRow, gstrPatholNumSet_ID) & ")"
                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
                
                Call ufgData.DelRow(lngRow, False)
                  
            Else
                Call MsgBoxD(Me, "该号码规则正在使用,不可删除", vbOKOnly, C_str提示)
            End If
        End If
    Else
        If MsgBoxD(Me, "是否确认删除该号码规则", vbYesNo + vbDefaultButton2, C_str提示) = vbNo Then Exit Sub
        
        Call ufgData.DelRow(lngRow, False)
        Call ufgData.RefreshData
    End If
    
    
    Exit Sub
errH:
    Call err.Raise(0, , "删除数据失败" & err.Description)
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdDel_Click()
    On Error GoTo errH
    If Not ufgData.IsSelectionRow Then
        Call MsgBoxD(Me, "请选则需要删除的项目,", vbOKOnly, C_str提示)
        Exit Sub
    End If

    Call DelRow
    Exit Sub
errH:
    Call MsgBoxD(Me, err.Description, vbOKOnly, C_str提示)
End Sub

Private Function CheckHaveErrCell() As Boolean
'判断是否有数据错误（根据颜色）
    Dim i As Integer
    Dim j As Integer
    Dim iCol As Integer
    
    On Error GoTo errH
    
    CheckHaveErrCell = True

    For i = 1 To ufgData.GridRows - 1
        If Not ufgData.RowHidden(i) Then
            For j = 1 To ufgData.GridCols - 1
                If ufgData.CellColor(i, j) = ufgData.ErrCellColor Then
                    CheckHaveErrCell = False
                    Exit Function
                End If
            Next
        End If
    Next
    Exit Function
errH:
    Call err.Raise(0, , "判断是否有无效数据失败" & err.Description)
End Function

Private Sub cmdSave_Click()
    Dim i As Integer
    Dim iCol As Integer
    Dim intMax As Integer
    Dim intID As Integer
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    On Error GoTo errH
    
    If CheckHaveErrCell = False Then
        Call MsgBoxD(Me, "存在无效数据，终止保存，请注意修改红色部分", vbOKOnly, C_str提示)
        Exit Sub
    End If

    Call gcnOracle.BeginTrans
    For i = 1 To ufgData.GridRows - 1
    
        If ufgData.RowState(i) <> TDataRowState.Del Then '删除状态的不保存，否则会时删除无效
        
            strSql = " select max(ID) as 最大ID from 病理号码规则 "
            Set rsData = zlDatabase.OpenSQLRecord(strSql, "查询最大号码规则ID")
            
            intMax = rsData!最大ID + 1
            
            If intMax > 90 And intMax < 100 Then
                Call MsgBoxD(Me, "病理号码规则数量即将超过限制，请尽快通知数据库管理人员处理", vbOKOnly, C_str提示)
            ElseIf intMax > 99 Then
                Call MsgBoxD(Me, "病理号码规则数量已经超过限制，无法进行新增操作，请尽快通知数据库管理人员处理", vbOKOnly, C_str提示)
                Exit Sub
            End If
            
    
            If ufgData.Text(i, gstrPatholNumSet_ID) = "" Then
                intID = intMax
            Else
                intID = Val(ufgData.Text(i, gstrPatholNumSet_ID))
            End If

                                                            
            strSql = "ZL_病理号码规则_Insert(" & intID & "," & _
                                                            Val(ufgData.Text(i, gstrPatholNumSet_类型)) & ",'" & _
                                                            ufgData.Text(i, gstrPatholNumSet_前缀) & "'," & _
                                                            Val(ufgData.Text(i, gstrPatholNumSet_年)) & "," & _
                                                            Val(ufgData.Text(i, gstrPatholNumSet_月)) & "," & _
                                                            Val(ufgData.Text(i, gstrPatholNumSet_日)) & "," & _
                                                            Val(ufgData.Text(i, gstrPatholNumSet_序号位数)) & "," & _
                                                            IIf(optYear.value = True, 4, 2) & "," & _
                                                            Val(ufgData.Text(i, gstrPatholNumSet_起始数)) & ",'" & _
                                                            ufgData.Text(i, gstrPatholNumSet_名称) & "')"
                                                            
                                                            
            Call zlDatabase.ExecuteProcedure(strSql, "病理号码规则_新建")
            ufgData.Text(i, gstrPatholNumSet_ID) = intID
            
        End If
    Next
    
    Call gcnOracle.CommitTrans
    Me.Hide

    Exit Sub
errH:
    Call gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Load()
    Call initUfgList
    Call LordPatholNumRules
    
    Call ufgData_OnAfterEdit(ufgData.DataGrid.Row, ufgData.DataGrid.Col)
End Sub

Private Sub ufgData_OnAfterEdit(ByVal Row As Long, ByVal Col As Long)
    On Error GoTo errH
    
    Dim i As Integer
    Dim iCol As Integer
    Dim blEx As Boolean  '用于判断前缀
    Dim intTMP As Integer
    Dim lngSelectRow As Long '选中的行数
    Dim strTMPEx As String
    Dim strTMPName As String
    
    lngSelectRow = ufgData.SelectionRow
    
    If ufgData.GridRows < 2 Then Exit Sub
    
    blEx = False
    
    With ufgData
       '先将颜色恢复为正常
       For i = 1 To .GridRows - 1
           
           .CellColor(lngSelectRow, .GetColIndex(gstrPatholNumSet_名称)) = .BackColor
           .CellColor(lngSelectRow, .GetColIndex(gstrPatholNumSet_前缀)) = .BackColor
           .CellColor(lngSelectRow, .GetColIndex(gstrPatholNumSet_类型)) = .BackColor
           .CellColor(lngSelectRow, .GetColIndex(gstrPatholNumSet_起始数)) = .BackColor
       Next
       
       '不可编辑区颜色DisCellColor
       For i = 1 To .GridRows - 1
           intTMP = Val(.Text(i, gstrPatholNumSet_ID))
           If intTMP >= 1 And intTMP <= 6 Then
               .CellColor(i, .GetColIndex(gstrPatholNumSet_名称)) = .DisCellColor
           End If
       Next
       
      ' 判断前缀、名称、类型是否为空
       For i = 1 To .GridRows - 1
           
           iCol = .GetColIndex(gstrPatholNumSet_名称)
           If Trim(.Text(i, gstrPatholNumSet_名称)) = "" Then
               .CellColor(i, iCol) = .ErrCellColor
           End If
           
           iCol = .GetColIndex(gstrPatholNumSet_类型)
           If Trim(.Text(i, gstrPatholNumSet_类型)) = "" Then
               .CellColor(i, iCol) = .ErrCellColor
           End If
       Next
       
       '判断前缀和名称是否重复
       For i = 1 To .GridRows - 1
           strTMPName = .Text(lngSelectRow, gstrPatholNumSet_名称)
           
           If (i <> lngSelectRow) And (.Text(i, gstrPatholNumSet_名称) = strTMPName And Trim(strTMPName) <> "") Then
               iCol = .GetColIndex(gstrPatholNumSet_名称)
               .CellColor(lngSelectRow, iCol) = .ErrCellColor
           End If
       Next
       
       '起始数
       iCol = .GetColIndex(gstrPatholNumSet_起始数)
       
       For i = 1 To .GridRows - 1
           If Val(.Text(i, gstrPatholNumSet_起始数)) < 0 Then
               .CellColor(i, iCol) = .ErrCellColor
               Call MsgBoxD(Me, "起始数只能为不小于0的整数,请检查", vbOKOnly, C_str提示)
               Exit Sub
           End If
       Next
    
       
       '判断前缀长度超过5
       iCol = .GetColIndex(gstrPatholNumSet_前缀)
       
       If .CellColor(Row, iCol) <> .ErrCellColor Then
       
           For i = 1 To Len(.Text(Row, gstrPatholNumSet_前缀))
              
               intTMP = Asc(Mid(.Text(Row, gstrPatholNumSet_前缀), i, 1))
               
               '这一段若不是字母或数字显示为红色
               If intTMP <= 47 Or (intTMP >= 58 And intTMP <= 64) Or (intTMP >= 91 And intTMP <= 96) Or intTMP >= 123 Then
                   .CellColor(Row, iCol) = .ErrCellColor
                   blEx = True
                   Exit For
               End If
           Next
           
           If Len(.Text(Row, gstrPatholNumSet_前缀)) > 5 Then
               .CellColor(Row, iCol) = .ErrCellColor
               blEx = True
           End If
       End If
       
    End With
    
    If blEx = True Then Call MsgBoxD(Me, "前缀字符数最大为5,且只能用数字或字母,请检查", vbOKOnly, C_str提示)
    
    Exit Sub
errH:
    Call MsgBoxD(Me, err.Description, vbOKOnly, C_str提示)
End Sub

Private Sub ufgData_OnKeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
' 禁止输入一些比如",:'"之类的符号
    If (KeyAscii >= 37 And KeyAscii <= 43) Or (KeyAscii >= 58 And KeyAscii <= 63) Or KeyAscii = 44 Or KeyAscii = 46 Or KeyAscii = 32 Then
        KeyAscii = 0
    End If
End Sub

Private Sub ufgData_OnStartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    On Error GoTo errH
    If 1 <= Val(ufgData.KeyValue(ufgData.SelectionRow)) And Val(ufgData.KeyValue(ufgData.SelectionRow)) <= 6 Then
        '试图编辑的如果是名称，则禁止编辑
        If Col = ufgData.GetColIndex(gstrPatholNumSet_名称) Then
            Cancel = True
            Call MsgBoxD(Me, "基础病理类型的名称不可修改,", vbOKOnly, C_str提示)
            Exit Sub
        End If
    End If
    
    Exit Sub
errH:
    Call MsgBoxD(Me, err.Description, vbOKOnly, C_str提示)
End Sub

