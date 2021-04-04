VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmCompoundMedicine 
   BorderStyle     =   0  'None
   Caption         =   "输液配药记录"
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   12825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picExec 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3000
      Left            =   11865
      ScaleHeight     =   3000
      ScaleWidth      =   960
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   960
      Begin XtremeCommandBars.CommandBars cbsExec 
         Left            =   120
         Top             =   30
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
      End
   End
   Begin VB.Frame fraExecUD 
      BorderStyle     =   0  'None
      Height          =   2445
      Left            =   1680
      MousePointer    =   9  'Size W E
      TabIndex        =   2
      Top             =   120
      Width           =   45
   End
   Begin VSFlex8Ctl.VSFlexGrid vsExec 
      Bindings        =   "frmCompoundMedicine.frx":0000
      Height          =   2955
      Left            =   1800
      TabIndex        =   0
      Top             =   0
      Width           =   10035
      _cx             =   17701
      _cy             =   5212
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   6
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   320
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmCompoundMedicine.frx":0028
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   2
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VSFlex8Ctl.VSFlexGrid vsSend 
      Align           =   3  'Align Left
      Bindings        =   "frmCompoundMedicine.frx":0177
      Height          =   3000
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1695
      _cx             =   2990
      _cy             =   5292
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   1
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   320
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmCompoundMedicine.frx":018B
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
End
Attribute VB_Name = "frmCompoundMedicine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event SetEditState(ByVal blnEditState As Boolean)      '当编辑状态时设置禁止其转移焦点的其他操作
Public Event StatusTextUpdate(ByVal Text As String) '要求更新主窗体状态栏文字

Private mlng病区ID As Long      '当前界面选择的病区
Private mlngAdviceID As Long    '给药途径的医嘱ID
Private mlng病人ID  As Long
Private mlng主页ID As Long
Private mlng病人性质  As Long
Private mlng科室ID  As Long
Private mstr姓名 As String
Private mstr住院号 As String
Private mstr床号 As String
Private mlng医嘱期效 As Long
Private mCol As Collection
Private Const Col发送时间 = 0
Private mrsCompoundGroup As ADODB.Recordset '配药批次
Private mblnEdit As Boolean '是否进入了修改模式，重读数据后，自动结束此模式
Private mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private mbln销帐申请 As Boolean '非打包的输液单在配液之后是否可以进行销帐申请
Private mbln打包修改 As Boolean
Private mbln摆药后不能改状态 As Boolean
Private mfrmParent As Object '父窗体对象

Private Const conMenu_Adjust = 100
Private Const conMenu_Save = 101
Private Const conMenu_Undo = 102
Private Const conMenu_AdjustCancle = 103


Public Sub RefreshData(ByVal lngAdviceID As Long, ByVal lng病区ID As Long, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lng病人性质 As Long, ByVal lng医嘱期效 As Long, _
        Optional ByRef objMip As Object, Optional frmParent As Object)
'功能：根据医嘱记录（给药途径的医嘱ID），刷新数据
    mlngAdviceID = lngAdviceID
    mlng病区ID = lng病区ID
    mlng病人ID = lng病人ID
    mlng病人性质 = lng病人性质
    mlng主页ID = lng主页ID
    mlng医嘱期效 = lng医嘱期效
    Set mfrmParent = frmParent
    If Not (objMip Is Nothing) Then Set mclsMipModule = objMip
    Call LoadSendList
End Sub


Private Sub Form_Load()
    Dim i As Long
    
    If GetInsidePrivs(p住院医嘱发送) = "" Then Exit Sub
    mbln销帐申请 = Val(zlDatabase.GetPara("配液输液单配药后允许销帐申请", glngSys, 1345, 0)) = 1
    mbln打包修改 = Val(zlDatabase.GetPara("打包设置", glngSys, 1345, 0)) = 1
    mbln摆药后不能改状态 = Val(zlDatabase.GetPara("输液单摆药后临床不允许改变打包状态", glngSys, 1345, 0)) = 1
    vsSend.Rows = vsSend.FixedRows
    vsExec.Rows = vsExec.FixedRows
    
    Set mCol = New Collection
    For i = 0 To vsExec.Cols - 1
        mCol.Add i, vsExec.TextMatrix(0, i)
    Next
    
    Set mrsCompoundGroup = GetCompoundGroup
    vsExec.ColDataType(mCol("打包")) = flexDTBoolean
    vsExec.ColHidden(mCol("销帐原因")) = Not gbln医嘱终止原因
    Call InitExecBar
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If WindowState = 1 Then Exit Sub
    
    vsExec.Top = Me.ScaleTop
    vsExec.Left = Me.ScaleLeft + vsSend.Width + 60
    vsExec.Width = Me.ScaleWidth - vsSend.Width - 60 - picExec.Width
    vsExec.Height = Me.ScaleHeight
    
    fraExecUD.Top = vsExec.Top
    fraExecUD.Left = vsExec.Left - fraExecUD.Width
    fraExecUD.Height = vsExec.Height
      
End Sub

Private Sub fraExecUD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If vsSend.Width + X < 50 Or vsExec.Width - X < 100 Then Exit Sub
        fraExecUD.Left = fraExecUD.Left + X
                
        vsSend.Width = vsSend.Width + X
        vsExec.Width = vsExec.Width - X
        vsExec.Left = vsExec.Left + X
    End If
End Sub


Private Function SaveData() As Boolean
    Dim colSQL As New Collection, blnTrans As Boolean
    Dim strSQL As String, i As Long, strCurDate As String
    Dim bytTmp As Byte
    Dim lngTmp As Long
    Dim strIDs As String, rsTmp As Recordset
    Dim colMsg As New Collection
    
    strCurDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    
    With vsExec
        For i = .FixedRows To .Rows - 1
            If Val(.Cell(flexcpData, i, mCol("状态"))) = 1 Then
                lngTmp = .Cell(flexcpChecked, i, mCol("打包"))
                If lngTmp <> 1 Then lngTmp = 0
                strIDs = strIDs & "," & .RowData(i)
                
                strSQL = "Zl_输液配药记录_Update(" & .RowData(i) & "," & lngTmp & "," & _
                    IIF(.Cell(flexcpData, i, mCol("配药批次")) = "", "Null", Val(.Cell(flexcpData, i, mCol("配药批次")))) & ",'" & UserInfo.姓名 & "'," & strCurDate & ")"
                colSQL.Add strSQL, "C" & colSQL.Count + 1
                
                colMsg.Add .RowData(i) & "," & Val(.Cell(flexcpData, i, mCol("配药批次"))), "K" & i
            End If
        Next
    End With
    If colSQL.Count = 0 Then
        RaiseEvent StatusTextUpdate("没有调整任何批次！")
    Else
        On Error GoTo errH
        strSQL = "select Count(1) as 是否锁定 from 输液配药记录 where 是否锁定=1 And ID in(Select Column_Value From Table(f_Num2list([1])))"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strIDs, 2))
        If rsTmp!是否锁定 > 0 Then
            MsgBox "当前调整的配药记录已经被输液配药中心锁定，暂时不允许进行调整。", vbInformation, "输液配液记录"
            Exit Function
        End If
        gcnOracle.BeginTrans: blnTrans = True
            For i = 1 To colSQL.Count
                Call zlDatabase.ExecuteProcedure(colSQL("C" & i), Me.Caption)
            Next
        gcnOracle.CommitTrans: blnTrans = False
        
        If Not (mclsMipModule Is Nothing) Then
            If mclsMipModule.IsConnect Then
                For i = 1 To colMsg.Count
                    Call ZLHIS_CIS_008(mclsMipModule, mlng病人ID, mstr姓名, mstr住院号, , mlng主页ID, mlng病区ID, , mlng科室ID, "", , mstr床号, mlngAdviceID, mlng医嘱期效, _
                    Split(colMsg(i), ",")(0), Split(colMsg(i), ",")(1))
                Next
            End If
        End If
        
        RaiseEvent StatusTextUpdate("数据保存成功！")
        For i = vsExec.FixedRows To vsExec.Rows - 1
            vsExec.Cell(flexcpData, i, mCol("状态")) = 0
        Next
    End If
        
    SaveData = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub LoadExecList(ByVal lngSendNO As Long)
'功能：读取并显示配药批次记录
'参数：lngSendNO=医嘱发送号
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim i As Long, blnDo As Boolean
    Dim rsState As Recordset  '输液配药状态
    Dim strIDs As String, strTmp As String
    Dim lng配液科室ID As Long
 
    strSQL = "Select ID,部门id as 配液科室ID,To_Char(执行时间, 'YYYY-MM-DD HH24:MI') 执行时间, Nvl(是否打包,0) 是否打包, 配药批次,瓶签号," & vbNewLine & _
            "       Decode(操作状态,1, '待摆药',2, '待配药', 3,'待配药', 4,'已配药',5,'已发送',6,'已签收',7,'已拒绝签收',8,'已确认拒收',9,'已销帐申请',10,'已销帐审核','已发送') As 状态," & _
            "'' AS 销帐申请人,'' as 销帐申请时间,'' as 销帐审核时间,姓名,住院号,床号,病人科室id" & vbNewLine & _
            "From 输液配药记录" & vbNewLine & _
            "Where 医嘱id = [1] And 发送号 = [2] And 操作状态 <> 8" & vbNewLine & _
            "Order By 执行时间"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngAdviceID, lngSendNO)
    Do While Not rsTmp.EOF
        strIDs = strIDs & "," & rsTmp!ID
        rsTmp.MoveNext
    Loop
    strIDs = Mid(strIDs, 2)
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        mstr姓名 = rsTmp!姓名 & ""
        mstr住院号 = rsTmp!住院号 & ""
        mstr床号 = rsTmp!床号 & ""
        mlng科室ID = Val(rsTmp!病人科室id & "")
        lng配液科室ID = Val(rsTmp!配液科室ID & "")
        mrsCompoundGroup.Filter = "配置中心id=" & lng配液科室ID
        For i = 1 To mrsCompoundGroup.RecordCount
            strTmp = strTmp & "|" & "#" & mrsCompoundGroup!批次 & ";第" & mrsCompoundGroup!批次 & "批:" & mrsCompoundGroup!配药时间
            mrsCompoundGroup.MoveNext
        Next
        strTmp = Mid(strTmp, 2)
        vsExec.ColComboList(mCol("配药批次")) = strTmp
    End If
    If strIDs <> "" Then
        strSQL = "Select 配药ID,操作类型,操作人员,To_Char(操作时间,'YYYY-MM-DD HH24:MI') as 操作时间,操作说明,操作时间 as 排序 from 输液配药状态 Where 配药ID in(select Column_Value From Table(Cast(f_num2list([1]) As ZLTOOLS.t_numlist)))"
        Set rsState = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strIDs)
    End If
    
    With vsExec
        .Redraw = False
        .Rows = .FixedRows
        .Rows = .FixedRows + rsTmp.RecordCount
                        
        For i = .FixedRows To .Rows - 1
            .RowData(i) = Val(rsTmp!ID)
            .TextMatrix(i, mCol("执行时间")) = rsTmp!执行时间
            .TextMatrix(i, mCol("打包")) = rsTmp!是否打包
            .Cell(flexcpData, i, mCol("打包")) = Val(rsTmp!是否打包)
            
            If IsNull(rsTmp!配药批次) Then
                .TextMatrix(i, mCol("配药批次")) = ""
                .Cell(flexcpData, i, mCol("配药批次")) = ""
            Else
                .TextMatrix(i, mCol("配药批次")) = "第" & rsTmp!配药批次 & "批"
                .Cell(flexcpData, i, mCol("配药批次")) = Val(rsTmp!配药批次)
            End If
            
            .TextMatrix(i, mCol("瓶签号")) = "" & rsTmp!瓶签号
            .TextMatrix(i, mCol("状态")) = rsTmp!状态
            .Cell(flexcpData, i, mCol("状态")) = 0
            If blnDo = False Then
                If rsTmp!状态 = "待摆药" Then blnDo = True
            End If
            If rsTmp!状态 = "已销帐申请" Or rsTmp!状态 = "已销帐审核" Then
                rsState.Filter = "配药ID=" & rsTmp!ID & " And 操作类型=9"
                rsState.Sort = "排序 desc"
                If rsState.RecordCount > 0 Then
                    rsState.MoveFirst
                    .TextMatrix(i, mCol("销帐申请人")) = "" & rsState!操作人员
                    .TextMatrix(i, mCol("销帐申请时间")) = "" & rsState!操作时间
                    .TextMatrix(i, mCol("销帐原因")) = "" & rsState!操作说明
                End If
            End If
            If rsTmp!状态 = "已销帐审核" Then
                rsState.Filter = "配药ID=" & rsTmp!ID & " And 操作类型=10"
                If rsState.RecordCount > 0 Then
                    rsState.MoveFirst
                    .TextMatrix(i, mCol("销帐审核时间")) = "" & rsState!操作时间
                End If
            End If
            
            If IsNull(rsTmp!配药批次) Then
                .TextMatrix(i, mCol("配药工作时间")) = ""
            Else
                mrsCompoundGroup.Filter = "批次=" & rsTmp!配药批次 & " and 配置中心id=" & lng配液科室ID
                If mrsCompoundGroup.RecordCount > 0 Then
                    .TextMatrix(i, mCol("配药工作时间")) = mrsCompoundGroup!配药时间
                End If
            End If
            
            .Cell(flexcpBackColor, i, mCol("打包")) = COLEditBackColor   '浅绿
            .Cell(flexcpBackColor, i, mCol("配药批次")) = COLEditBackColor
            rsTmp.MoveNext
        Next
        
        .Redraw = True
        If .Rows > .FixedRows Then
            .Row = .Rows - 1
            .TopRow = .Row
        End If
    End With
    
    If blnDo = False Then
        vsExec.Tag = "false"
    Else
        vsExec.Tag = ""
    End If
    mblnEdit = False
    vsExec.Editable = flexEDNone
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadSendList()
'功能：读取并显示医嘱发送记录
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim i As Long
 
    strSQL = "Select To_Char(发送时间, 'YYYY-MM-DD HH24:MI') 发送时间, 发送号 From 病人医嘱发送 Where 医嘱id = [1] Order by 发送时间 Desc"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngAdviceID)

    With vsSend
        .Redraw = False
        .Rows = .FixedRows
        .Rows = .FixedRows + rsTmp.RecordCount
        For i = 1 To rsTmp.RecordCount
            .TextMatrix(i, 0) = rsTmp!发送时间
            .Cell(flexcpData, i, Col发送时间) = Val(rsTmp!发送号)
            rsTmp.MoveNext
        Next
        .Redraw = True
        If .Rows > 1 Then .Row = 1
    End With

    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetCompoundGroup() As ADODB.Recordset
'功能：读取配药工作批次
    Dim strSQL As String
    
    strSQL = "Select 配置中心id,批次, 配药时间 From 配药工作批次 Order By 批次"
    On Error GoTo errH
    Set GetCompoundGroup = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set mCol = Nothing
    Set mrsCompoundGroup = Nothing
    Set mclsMipModule = Nothing
End Sub

Private Sub vsExec_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = mCol("配药批次") Then
        With vsExec
            '未选择时离开焦点
            If .ComboIndex = -1 Or .Cell(flexcpData, Row, Col) = .ComboData Then
                If Val(.Cell(flexcpData, Row, Col)) <> 0 Then .TextMatrix(Row, Col) = "第" & .Cell(flexcpData, Row, Col) & "批"
                Exit Sub
            End If
            
            
            .Cell(flexcpData, Row, Col) = CStr(.ComboData)
            .TextMatrix(Row, Col) = "第" & .ComboData & "批"
            .TextMatrix(Row, mCol("配药工作时间")) = Mid(.ComboItem, InStr(.ComboItem, ":") + 1)
            
            .Cell(flexcpData, Row, mCol("状态")) = 1 '表示修改过的记录
        End With
    ElseIf Col = mCol("打包") Then
        With vsExec
            If Val(.Cell(flexcpData, Row, mCol("打包"))) <> Val(.TextMatrix(Row, mCol("打包"))) Then
                .Cell(flexcpData, Row, mCol("状态")) = 1 '表示修改过的记录
            End If
        End With
    End If
End Sub

Private Sub vsExec_ComboCloseUp(ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)
    Call vsExec_KeyPressEdit(Row, Col, 13)  '将自动调用vsExec_AfterEdit
End Sub

Private Sub vsExec_DblClick()
    Dim objControl As CommandBarControl
    
    If vsExec.Editable = flexEDNone Then
        If (vsExec.TextMatrix(vsExec.Row, mCol("状态")) = "待摆药" Or vsExec.MouseCol = mCol("打包") And vsExec.TextMatrix(vsExec.Row, mCol("状态")) = "待配药") And vsExec.TextMatrix(vsExec.Row, mCol("销帐申请人")) = "" Then
            If MsgBox("你确定要调整" & IIF(mbln打包修改, "配药批次或打包", "配药批次") & "吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                                
                Set objControl = cbsExec.FindControl(, conMenu_Adjust)
                If Not objControl Is Nothing Then Call cbsExec_Execute(objControl)
                
            End If
        End If
    End If
End Sub

Private Sub vsExec_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        vsExec.Col = Col + 1
    End If
End Sub

Private Sub vsExec_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'说明：修改打包只是能 待摆药或待配药 两个状态，受参数 打包设置 控制。
'                       待配药 状态修改打包时，受参数 输液单摆药后临床不允许改变打包状态 控制。
    If (Col = mCol("打包") And mbln打包修改 Or Col = mCol("配药批次")) Then
        Cancel = (vsExec.TextMatrix(Row, mCol("状态")) <> "待摆药")
        If Col = mCol("打包") Then
            If vsExec.TextMatrix(Row, mCol("状态")) = "待配药" Then
                If mbln摆药后不能改状态 Then
                    Cancel = True
                    MsgBox "不允许调整已经摆药的输液单的打包状态。", vbInformation, Me.Caption
                    Exit Sub
                Else
                    Cancel = False
                End If
            End If
        End If
        If Cancel = False Then
            If vsExec.TextMatrix(Row, mCol("销帐申请人")) <> "" Then
                Cancel = True
                MsgBox "已经申请销帐的记录不允许修改。", vbInformation, Me.Caption
            End If
            '判断权限
            If Not Cancel Then
                If Col = mCol("打包") Then
                    If InStr(GetInsidePrivs(p住院医嘱发送), ";修改配液打包状态;") = 0 Then
                        Cancel = True
                        MsgBox "您没有修改打包状态权限，不能进调整。", vbInformation, Me.Caption
                    End If
                End If
                
                If Col = mCol("配药批次") Then
                    If InStr(GetInsidePrivs(p住院医嘱发送), ";修改配液批次;") = 0 Then
                        Cancel = True
                        MsgBox "您没有修改配药批次权限，不能进调整。", vbInformation, Me.Caption
                    End If
                End If
            End If
        Else
            If Col = mCol("打包") Then
                MsgBox "只有待摆药或待配药的记录可打包。", vbInformation, Me.Caption
            Else
                MsgBox "只有待摆药的记录可调整批次。", vbInformation, Me.Caption
            End If
        End If
    Else
        Cancel = True
    End If
End Sub

Private Sub vsSend_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow <> NewRow And Me.Visible = True Then
        If Val(vsSend.Cell(flexcpData, NewRow, Col发送时间)) <> 0 Then
            Call LoadExecList(Val(vsSend.Cell(flexcpData, NewRow, Col发送时间)))
        End If
    End If
End Sub


Private Sub InitExecBar()
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    Dim strPrivs As String

    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsExec.VisualTheme = xtpThemeOfficeXP  '使用2003风格时，按钮有突出效果，只有一个按钮时不好看
    With Me.cbsExec.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .UseFadedIcons = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
    End With
    Set cbsExec.Icons = zlCommFun.GetPubIcons
    cbsExec.EnableCustomization False
    cbsExec.ActiveMenuBar.Visible = False
        
    Set objBar = cbsExec.Add("工具栏", xtpBarTop)
    objBar.EnableDocking xtpFlagStretched   '宽度不够时自动换行
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = True
    objBar.SetIconSize 24, 24
    
    strPrivs = GetInsidePrivs(p住院记帐操作)
            
    With objBar.Controls
        If InStr(strPrivs, ";药品销帐申请;") > 0 Then
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ChargeDelApply, "销帐申请")
            objControl.IconId = 3821
        End If
        Set objControl = .Add(xtpControlButton, conMenu_AdjustCancle, "取消申请")
        objControl.BeginGroup = True
        objControl.ToolTipText = "取消销帐申请"
        objControl.IconId = conMenu_Edit_Untread
        Set objControl = .Add(xtpControlButton, conMenu_Adjust, "调整批次")
        objControl.BeginGroup = True
        objControl.IconId = 3564
        Set objControl = .Add(xtpControlButton, conMenu_Save, "保存")
        objControl.Visible = False
        objControl.IconId = 3503
        Set objControl = .Add(xtpControlButton, conMenu_Undo, "放弃")
        objControl.Visible = False
        objControl.IconId = 3014
        
        objControl.BeginGroup = True
    End With
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    picExec.BackColor = cbsExec.GetSpecialColor(STDCOLOR_BTNFACE)
End Sub

Private Sub cbsExec_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
            
    Select Case Control.ID
        Case conMenu_Adjust    '调整批次
            mblnEdit = True
            
            vsExec.Editable = flexEDKbdMouse
            vsSend.Enabled = False
            vsExec.SetFocus
            RaiseEvent SetEditState(True)
            
        Case conMenu_Save '保存
            
            If SaveData Then
                mblnEdit = False
                
                vsExec.Editable = flexEDNone
                vsSend.Enabled = True
                RaiseEvent SetEditState(False)
                Call LoadExecList(Val(vsSend.Cell(flexcpData, vsSend.Row, Col发送时间)))
            End If
            
        Case conMenu_Undo  '放弃
            mblnEdit = False
            
            vsExec.Editable = flexEDNone
            vsSend.Enabled = True
            RaiseEvent SetEditState(False)
        
            Call LoadExecList(Val(vsSend.Cell(flexcpData, vsSend.Row, Col发送时间)))
        Case conMenu_Edit_ChargeDelApply    '销帐申请
            Call ExecChargeDelApply(Control.Caption = "销帐")
        Case conMenu_AdjustCancle    '取消申请销帐
            Call ExecCancleChargeDelApply
    End Select
    
End Sub

Private Sub cbsExec_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim str状态 As String
    Dim blnVisible As Boolean
     
    Select Case Control.ID
        Case conMenu_Adjust    '调整批次
            Control.Enabled = Not mblnEdit And vsExec.Tag = "" And vsExec.TextMatrix(vsExec.Row, mCol("销帐申请人")) = ""
            Control.Visible = Not mblnEdit
        Case conMenu_Save, conMenu_Undo '保存
            Control.Visible = mblnEdit
        Case conMenu_Edit_ChargeDelApply
            '先设置成false，最后设了名称才设置成true，因为commandbar的bug，设置了名字鼠标移上去才会更新名称
            Control.Visible = False
            If vsExec.Row >= vsExec.FixedRows Then
                str状态 = vsExec.TextMatrix(vsExec.Row, mCol("状态"))
                Control.Enabled = False
                If Not mblnEdit Then
                    blnVisible = True
                    If vsExec.TextMatrix(vsExec.Row, mCol("销帐申请人")) = "" Then
                        If str状态 = "待配药" Or str状态 = "待摆药" Then
                            Control.Enabled = True
                        ElseIf Not (str状态 = "待配药" Or str状态 = "待摆药") Then
                            If vsExec.Cell(flexcpChecked, vsExec.Row, mCol("打包")) = 1 Then
                                Control.Enabled = True
                            Else
                                If mbln销帐申请 Then Control.Enabled = True
                            End If
                        End If
                    ElseIf vsExec.TextMatrix(vsExec.Row, mCol("销帐申请人")) <> "" And vsExec.TextMatrix(vsExec.Row, mCol("销帐审核时间")) = "" Then
                        blnVisible = False
                    End If
                Else
                    blnVisible = False
                End If
                If str状态 = "待摆药" And Control.Enabled Then
                    Control.Caption = "销帐"
                    Control.ToolTipText = "销帐"
                Else
                    Control.Caption = "销帐申请"
                    Control.ToolTipText = "销帐申请"
                End If
            Else
                Control.Enabled = False
                Control.Caption = "销帐申请"
                Control.ToolTipText = "销帐申请"
            End If
            Control.Visible = blnVisible
        Case conMenu_AdjustCancle
            If vsExec.Row >= vsExec.FixedRows Then
                If Not mblnEdit And vsExec.TextMatrix(vsExec.Row, mCol("销帐申请人")) <> "" And vsExec.TextMatrix(vsExec.Row, mCol("销帐审核时间")) = "" Then
                    Control.Visible = True
                Else
                    Control.Visible = False
                End If
            Else
                Control.Visible = False
            End If
    End Select
End Sub


Private Sub ExecChargeDelApply(Optional ByVal blnAutoAduit As Boolean)
'功能：执行销帐申请
'参数：blnAutoAduit=true自动审核销帐申请
    Dim lng配药记录 As Long, strDate As String, strNote As String
    Dim rsTmp As ADODB.Recordset, strSQL As String, i As Long
    Dim colSQL As New Collection, blnTrans As Boolean
    Dim strProg As String
    Dim strTmp As String
    Dim lng审核部门ID As Long
    Dim str销帐原因 As String
    Dim strTab As String
    
    If InStr(GetInsidePrivs(p住院记帐操作), ";药品销帐申请;") = 0 Then
        MsgBox "您没有住院记账操作模块中的药品销帐申请权限，不能进行申请销帐。"
        Exit Sub
        
    End If
    If blnAutoAduit Then
        strProg = "销帐"
        If InStr(GetInsidePrivs(p住院记帐操作), ";销帐审核;") = 0 Then
            MsgBox "您没有住院记账操作模块中的药品销帐审核权限，不能进行销帐。"
            Exit Sub
        End If
    Else
        strProg = "销帐申请"
    End If
    
    With vsExec
        If .TextMatrix(.Row, mCol("配药批次")) = "" Then
            MsgBox "当前选择了第" & .Row & "行,配药批次为空，不能进行" & strProg & "。", vbInformation, gstrSysName
            Exit Sub
        End If
        On Error GoTo errH
        strSQL = "select 是否锁定,操作状态,是否打包,部门ID from 输液配药记录 where ID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.RowData(.Row)))
        lng审核部门ID = Val(rsTmp!部门ID & "")
        If Val(rsTmp!是否锁定 & "") = 1 Then
            MsgBox "当前销帐的配药记录已经被输液配药中心锁定，暂时不允许进行销帐。", vbInformation, "输液配液记录"
            Exit Sub
        End If
        If Val(rsTmp!操作状态 & "") = 9 Or Val(rsTmp!操作状态 & "") = 10 Then
            MsgBox "当前配药记录已经被" & strProg & "，请检查。", vbInformation, "输液配液记录"
            Call LoadExecList(Val(vsSend.Cell(flexcpData, vsSend.Row, Col发送时间)))
            Exit Sub
        End If
        If Val(rsTmp!操作状态 & "") >= 4 Then
            If Val(rsTmp!是否打包 & "") = 0 And mbln销帐申请 = False Then
                MsgBox "当前未打包的记录已经配药，不允许再销帐申请。", vbInformation, "输液配液记录"
                Call LoadExecList(Val(vsSend.Cell(flexcpData, vsSend.Row, Col发送时间)))
                Exit Sub
            End If
        End If
	If mlng医嘱期效 = 0 And mlng病人性质 <> 1 Then
            strTab = "住院费用记录"
        Else
            If GetAdviceFeeKind(mlngAdviceID) = 2 Then    '住院医生站的临嘱可发送到门诊
                strTab = "住院费用记录"
            Else
                strTab = "门诊费用记录"
            End If
        End If
        '77686,李南春,2014/9/18,单据类别限制
        strSQL = "Select b.费用id, b.药品id As 收费细目id, Sum(a.数量) As 数量, c.住院包装, c.住院单位, d.名称, b.No, e.序号,E.记录状态" & vbNewLine & _
            "From 输液配药内容 A, 药品收发记录 B, " & strTab & " E, 药品规格 C, 收费项目目录 D" & vbNewLine & _
            "Where a.记录id = [1] And a.收发id = b.Id And b.费用id = e.Id And b.药品id = c.药品id And c.药品id = d.Id" & vbNewLine & _
            IIF(Not blnAutoAduit, " And b.审核人 is Not null", "") & vbNewLine & _
            " And instr( ',8,9,10,21,24,25,26,',','||B.单据||',')>0 " & _
            "Group By b.费用id, b.药品id, c.住院包装, c.住院单位, d.名称, b.No, e.序号,E.记录状态" & vbNewLine & _
            "Order By e.序号"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.RowData(.Row)))
        If rsTmp.RecordCount = 0 Then
            MsgBox "当前选择了第" & .Row & "行，没有找到相应的配药内容，不能进行" & strProg & "。", vbInformation, gstrSysName
            Exit Sub
        Else
            '已配药的都是已发药的，如果不是自动审核按已发药类型申请销帐
            '一个配药批次，一般至少两个药品
            For i = 1 To rsTmp.RecordCount
                If Val(rsTmp!记录状态 & "") = 0 Then
                    MsgBox "当前配药记录是划价单，无需进行销帐。", vbInformation, gstrSysName
                    Exit Sub
                End If
                strNote = strNote & vbCrLf & rsTmp!名称 & "：" & FormatEx(rsTmp!数量 / rsTmp!住院包装, 5) & rsTmp!住院单位
                rsTmp.MoveNext
            Next
            If gbln医嘱终止原因 Then
                Call frmAdviceStopTime.ShowMe(mfrmParent, mlngAdviceID, mlng科室ID, 2, , str销帐原因)
                If str销帐原因 = "" Then Exit Sub
            End If
            If MsgBox("当前选择了第" & .Row & "行配药批次：" & strNote & vbCrLf & "你确定要对这些药品" & strProg & "吗？", vbYesNo + vbQuestion, gstrSysName) = vbNo Then
                Exit Sub
            End If
            
            rsTmp.MoveFirst
            
            strTmp = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
            strDate = "To_Date('" & strTmp & "','YYYY-MM-DD HH24:MI:SS')"
            For i = 1 To rsTmp.RecordCount
                strSQL = "Zl_病人费用销帐_Insert(" & rsTmp!费用ID & "," & rsTmp!收费细目ID & "," & _
                      mlng病区ID & "," & rsTmp!数量 & ",'" & UserInfo.姓名 & "'," & strDate & "," & IIF(blnAutoAduit, "0", "1") & ",1," & Val(.RowData(.Row)) & ",'" & str销帐原因 & "')"
                colSQL.Add strSQL, "C" & colSQL.Count + 1
                '待摆药状态的自动审核
                If blnAutoAduit Then
                    strSQL = "Zl_病人费用销帐_Audit(" & rsTmp!费用ID & "," & strDate & ",'" & _
                          UserInfo.姓名 & "'," & strDate & ",1,1,0" & ")"
                    colSQL.Add strSQL, "C" & colSQL.Count + 1
                    If strTab = "门诊费用记录" Then
                        strSQL = "Zl_门诊记帐记录_Delete('" & rsTmp!NO & "', '" & rsTmp!序号 & ":" & rsTmp!数量 & ":" & Val(.RowData(.Row)) & "', '" & UserInfo.编号 & "', '" & UserInfo.姓名 & "', 0)"
                    Else
                        strSQL = "Zl_住院记帐记录_Delete('" & rsTmp!NO & "', '" & rsTmp!序号 & ":" & rsTmp!数量 & ":" & Val(.RowData(.Row)) & "', '" & UserInfo.编号 & "', '" & UserInfo.姓名 & "', 2)"
                    End If
                    colSQL.Add strSQL, "C" & colSQL.Count + 1
                End If
                rsTmp.MoveNext
            Next
            
            On Error GoTo errH
            gcnOracle.BeginTrans: blnTrans = True
                For i = 1 To colSQL.Count
                    Call zlDatabase.ExecuteProcedure(colSQL("C" & i), Me.Caption)
                Next
            gcnOracle.CommitTrans: blnTrans = False
            'ZLHIS_CIS_013-住院患者输液销帐申请
            If Not (mclsMipModule Is Nothing) Then
                If mclsMipModule.IsConnect Then
                    Call ZLHIS_CIS_013(mclsMipModule, mlng病人ID, mstr姓名, mstr住院号, mlng主页ID, mlng病区ID, , mlng科室ID, , mlngAdviceID, Val(vsExec.RowData(vsExec.Row)), strTmp, UserInfo.姓名, mlng病区ID, , lng审核部门ID)
                End If
            End If
            
            MsgBox strProg & "操作成功！", vbInformation, gstrSysName
            i = .Row
            Call vsSend_AfterRowColChange(0, 0, vsSend.Row, vsSend.Col)
            If .Rows > i Then
                .Row = i
            End If
        End If
    End With
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub ExecCancleChargeDelApply()
'功能：取消销帐申请
    Dim lng配药记录 As Long, strDate As String, strNote As String
    Dim rsTmp As ADODB.Recordset, strSQL As String, i As Long
    Dim colSQL As New Collection, blnTrans As Boolean
    Dim strProg As String
    Dim strTmp As String
    
    If InStr(GetInsidePrivs(p住院记帐操作), ";药品销帐申请;") = 0 Then
        MsgBox "您没有住院记账操作模块中的药品销帐申请权限，不能进行取消申请销帐。"
        Exit Sub
    End If
    strProg = "取消销帐申请"
    With vsExec
        If .TextMatrix(.Row, mCol("配药批次")) = "" Then
            MsgBox "当前选择了第" & .Row & "行,配药批次为空，不能进行" & strProg & "。", vbInformation, gstrSysName
            Exit Sub
        End If
        On Error GoTo errH
        strSQL = "select Count(1) as 是否锁定 from 输液配药记录 where 是否锁定=1 And ID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.RowData(.Row)))
        If rsTmp!是否锁定 > 0 Then
            MsgBox "当前销帐的配药记录已经被输液配药中心锁定，暂时不允许进行取消销帐申请。", vbInformation, "输液配液记录"
            Exit Sub
        End If
        '77686,李南春,2014/9/18,单据类别限制
        strSQL = "Select distinct c.费用id" & vbNewLine & _
                "From 输液配药内容 A, 药品收发记录 B, 病人费用销帐 C" & vbNewLine & _
                "Where a.记录id = [1] And a.收发id = b.Id And b.费用id = c.费用id And c.审核人 Is Null " & _
                "And instr( ',8,9,10,21,24,25,26,',','||B.单据||',')>0"

        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.RowData(.Row)))
        If rsTmp.RecordCount = 0 Then
            MsgBox "当前选择了第" & .Row & "行，没有找到相应的销帐申请记录，不能进行" & strProg & "。", vbInformation, gstrSysName
            Exit Sub
        Else

            For i = 1 To rsTmp.RecordCount
                strTmp = strTmp & "," & rsTmp!费用ID
                rsTmp.MoveNext
            Next
            strTmp = Mid(strTmp, 2)
            
            strSql = "Zl_病人费用销帐_Delete('" & strTmp & "'," & Val(.RowData(.Row)) & ")"
            colSQL.Add strSQL, "C" & colSQL.Count + 1

            If MsgBox("当前选择了第" & .Row & "行配药记录," & vbCrLf & "你确定要对这些药品" & strProg & "吗？", vbYesNo + vbQuestion, gstrSysName) = vbNo Then
                Exit Sub
            End If
            
            On Error GoTo errH
            gcnOracle.BeginTrans: blnTrans = True
                For i = 1 To colSQL.Count
                    Call zlDatabase.ExecuteProcedure(colSQL("C" & i), Me.Caption)
                Next
            gcnOracle.CommitTrans: blnTrans = False
            
            MsgBox strProg & "操作成功！", vbInformation, gstrSysName
            i = .Row
            Call vsSend_AfterRowColChange(0, 0, vsSend.Row, vsSend.Col)
            If .Rows > i Then
                .Row = i
            End If
        End If
    End With
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

