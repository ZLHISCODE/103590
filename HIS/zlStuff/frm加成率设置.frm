VERSION 5.00
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Begin VB.Form frm加成率设置 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "加成率设置"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8520
   Icon            =   "frm加成率设置.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   8520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txt限价 
      Height          =   300
      Left            =   6204
      MaxLength       =   16
      TabIndex        =   4
      Text            =   "800.00"
      Top             =   912
      Width           =   2196
   End
   Begin VB.ComboBox cbo计算方法 
      Height          =   276
      ItemData        =   "frm加成率设置.frx":030A
      Left            =   1128
      List            =   "frm加成率设置.frx":0314
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   924
      Width           =   2184
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6045
      TabIndex        =   6
      Top             =   5325
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7335
      TabIndex        =   7
      Top             =   5325
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   75
      Left            =   -500
      TabIndex        =   9
      Top             =   5070
      Width           =   10000
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   0
      TabIndex        =   8
      Top             =   645
      Width           =   10275
   End
   Begin ZL9BillEdit.BillEdit mshBill 
      Height          =   3804
      Left            =   60
      TabIndex        =   5
      Top             =   1260
      Width           =   8352
      _ExtentX        =   14737
      _ExtentY        =   6720
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
      Cols            =   2
      RowHeight0      =   315
      RowHeightMin    =   315
      ColWidth0       =   1005
      BackColor       =   -2147483643
      BackColorBkg    =   -2147483643
      BackColorSel    =   10249818
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      ForeColorSel    =   -2147483634
      GridColor       =   -2147483630
      ColAlignment0   =   9
      ListIndex       =   -1
      CellBackColor   =   -2147483643
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "最高限价(&X)"
      Height          =   180
      Index           =   1
      Left            =   5136
      TabIndex        =   3
      Top             =   972
      Width           =   1008
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "计算方法(&J)"
      Height          =   180
      Index           =   0
      Left            =   108
      TabIndex        =   1
      Top             =   972
      Width           =   1008
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   1
      Left            =   165
      Picture         =   "frm加成率设置.frx":0330
      Top             =   135
      Width           =   480
   End
   Begin VB.Label lblInfor 
      Caption         =   $"frm加成率设置.frx":09B1
      Height          =   480
      Left            =   780
      TabIndex        =   0
      Top             =   240
      Width           =   7668
   End
End
Attribute VB_Name = "frm加成率设置"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnChange As Boolean

Dim mstrSql As String
Dim mblnReturn As Boolean
Dim mblnFirst As Boolean

Dim mstrPriv As String           '权限串

Private mintPreCol As Integer               '前一次单据头的排序列
Private mintsort As Integer                 '前一次单据头的排序
Private Enum marBillCol
    序号 = 0
    最低价
    最高价
    加成率
    分段最高限价
    说明
End Enum

'----------------------------------------------------------------------------------------------------------
'刘兴宏:增加小数位数的格式串
'修改:2007/03/06
Private mFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------
Private Sub cbo计算方法_Click()
 mblnChange = True
 SetCtlEnable
End Sub

Private Sub cbo计算方法_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub mshBill_AfterDeleteRow()
    '调整工资值
    Call ReFormal
    mblnChange = True
    '设置控件属性
    SetCtlEnable
End Sub
Private Sub mshbill_EditChange(curText As String)
    mblnChange = True
    SetCtlEnable
End Sub
Private Sub mshbill_EnterCell(Row As Long, Col As Long)
    With mshBill
        SetInputFormat .Row
        Select Case .Col
            Case marBillCol.说明
                ImeLanguage True
                .TxtCheck = False
                .MaxLength = 50
                .TxtSetFocus
            Case marBillCol.最低价, marBillCol.最高价, marBillCol.加成率
                .TxtCheck = True
                .MaxLength = 11
                .TextMask = ".1234567890"
        End Select
        SetCtlEnable
    End With
End Sub

Private Sub SetInputFormat(ByVal intRow As Integer)
    With mshBill
        If intRow <> 1 Then
            .ColData(marBillCol.最低价) = 5               '禁止
        Else
            .ColData(marBillCol.最低价) = 4               '纯文本输入
        End If
    End With
End Sub

Private Sub mshbill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    Dim strSQL As String
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With mshBill
        .Text = UCase(Trim(.Text))
        strKey = UCase(Trim(.Text))
        Select Case .Col
            Case marBillCol.说明
                OS.OpenIme False
                If strKey = "" Then
                    .Text = " "
                    .TextMatrix(.Row, marBillCol.说明) = " "
                End If
            Case marBillCol.最低价
                
               If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "最低价必须为数字型,请重新输入！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" Then
                    If Val(strKey) < 0 Then
                        MsgBox "最低价必须大于零,请重新输入！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Val(strKey) >= 10 ^ 11 - 1 Then
                        MsgBox "最低价必须小于" & (10 ^ 11 - 1), vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    .Text = Format(Val(strKey), mFMT.FM_成本价)
                End If
                ReFormal

            Case marBillCol.最高价
                
                If .Row - 1 > 1 Then
                    .TextMatrix(.Row, marBillCol.最低价) = Format(Val(.TextMatrix(.Row - 1, marBillCol.最高价)), mFMT.FM_成本价)
                End If
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "最高价必须为数字型,请重新输入！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" Then
                    If Val(strKey) < 0 Then
                        MsgBox "最高价必须大于零,请重新输入！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Val(strKey) >= 10 ^ 14 - 1 Then
                        MsgBox "最高价必须小于" & (10 ^ 14 - 1), vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    If Val(strKey) < Val(.TextMatrix(.Row, marBillCol.最低价)) And Val(strKey) <> 0 Then
                        MsgBox "最高价必须大于最低价", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    .Text = Format(Val(strKey), mFMT.FM_成本价)
                    .TextMatrix(.Row, .Col) = .Text
                ElseIf Val(.TextMatrix(.Row, marBillCol.最高价)) = 0 Then
                    .TextMatrix(.Row, .Col) = " "
                    .Text = " "
                End If
                ReFormal
            Case marBillCol.加成率
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "加成率必须为数字型,请重新输入！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" Then
                    If Val(strKey) < 0 Then
                        MsgBox "加成率必须大于零,请重新输入！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    If Val(strKey) > 100 Then
                        MsgBox "加成率必须小于100%", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    strKey = Format(strKey, GFM_VBJCL)
                    .Text = strKey
                    .TextMatrix(.Row, marBillCol.加成率) = strKey
                End If
            Case marBillCol.分段最高限价
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "分段最高限价必须为数字型,请重新输入！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If strKey <> "" Then
                    If Val(strKey) < 0 Then
                        MsgBox "分段最高限价必须大于零,请重新输入！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                End If
        End Select
    End With
End Sub

Private Sub mshBill_LeaveCell(Row As Long, Col As Long)
    OS.OpenIme False
End Sub

Private Sub mshBill_LostFocus()
    OS.OpenIme False
End Sub


'------------------------------------------------------------------
'------------------------------------------------------------------
'-1：表示该列可以选择，是布尔型［"√"，" "］
' 0：表示该列可以选择，但不能修改
' 1：表示该列可以输入，外部显示为按钮选择
' 2：表示该列是日期列，外部显示为按钮选择，弹出是日期选择框
' 3：表示该列是选择列，外部显示为下拉框选择
'4:  表示该列为单纯的文本框供用户输入
'5:  表示该列不允许选择
'-----------------------------------------------------------------
'-----------------------------------------------------------------

Private Function ValidData() As Boolean
    Dim intLop As Integer
    Dim dbl上次最高价 As Double
    
    Dim blnStock As Boolean
    
    ValidData = False
    blnStock = False
    
    dbl上次最高价 = 0
    If cbo计算方法.ListIndex < 0 Then
        ShowMsgBox "计算方法必需选择!"
        If cbo计算方法.Enabled Then cbo计算方法.SetFocus
        Exit Function
    End If
    If txt限价.Text = "" Or Val(txt限价.Text) = 0 Then
        ShowMsgBox "必需输入最高限价!"
        If txt限价.Enabled Then txt限价.SetFocus
        Exit Function
    End If
    If Abs(Val(txt限价)) > 10 ^ 11 - 1 Then
        ShowMsgBox "最高限价必需在(-" & 10 ^ 11 - 1 & " ～ " & 10 ^ 11 - 1 & ")!"
        If txt限价.Enabled Then txt限价.SetFocus
        Exit Function
    End If
    With mshBill
            For intLop = 1 To .Rows - 1
                If .TextMatrix(intLop, marBillCol.最低价) <> "" Or .TextMatrix(intLop, marBillCol.最高价) <> "" Then           '先判有否数据
                    If intLop = 1 Then
                        dbl上次最高价 = Val(.TextMatrix(intLop, marBillCol.最高价))
                    Else
                        If Val(.TextMatrix(intLop, marBillCol.最低价)) <> dbl上次最高价 Then
                            ShowMsgBox "在第" & intLop & "行的最低价不等于" & intLop - 1 & "行的最高价!"
                            Exit Function
                        End If
                        dbl上次最高价 = Val(.TextMatrix(intLop, marBillCol.最高价))
                        .SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = marBillCol.最低价
                    End If
                    
                    
                    If Val(.TextMatrix(intLop, marBillCol.最低价)) > Val(.TextMatrix(intLop, marBillCol.最高价)) And Val(.TextMatrix(intLop, marBillCol.最高价)) <> 0 Then
                        ShowMsgBox "在第" & intLop & "行的最低价大于了最高价!"
                        .SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = marBillCol.最低价
                            Exit Function
                    End If
                                        
                    
                    If Trim(Trim(.TextMatrix(intLop, marBillCol.加成率))) = "" Then
                        ShowMsgBox "第" & intLop & "行加成率为空了，请检查！"
                        .SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = marBillCol.加成率
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, marBillCol.最低价)) > 9999999999# Then
                        ShowMsgBox "  第" & intLop & "行的最低价大于了数据库能够保存的" & vbCrLf & "最大范围9999999999，请检查！"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = marBillCol.最低价
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, marBillCol.最高价)) > 9999999999# Then
                        ShowMsgBox "  第" & intLop & "行的最高价大于了数据库能够保存的" & vbCrLf & "最大范围9999999999，请检查！"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = marBillCol.最高价
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, marBillCol.加成率)) > 100 Then
                        ShowMsgBox "  第" & intLop & "行的加成率大于了100%，请检查！"
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = marBillCol.最高价
                        Exit Function
                    End If
                End If
                
                If LenB(StrConv(.TextMatrix(intLop, marBillCol.说明), vbFromUnicode)) > 50 Then
                    MsgBox "第" & intLop & "行说明列长度大于50个字符了，请重新输入！", vbInformation, gstrSysName
                    .SetFocus
                    .Row = intLop
                    .Col = marBillCol.说明
                    Exit Function
                End If
            Next
    End With
    ValidData = True
End Function

Private Function SaveCard() As Boolean
    Dim str说明 As String
    Dim dbl最低价 As Double
    Dim dbl最高价 As Double
    Dim dbl加成率 As Double
    Dim byt计算方法 As Byte
    Dim dbl限价 As Double
    
    Dim strSQL As String
    Dim intRow As Integer
    
    SaveCard = False
    With mshBill
        On Error GoTo ErrHandle
        dbl限价 = Val(txt限价.Text)
        byt计算方法 = cbo计算方法.ItemData(cbo计算方法.ListIndex)
        gcnOracle.BeginTrans
        
        '清除原来的工资方案
        strSQL = "ZL_材料加成方案_DELETE()"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        '增加固定的总限价值
        If Trim(Me.txt限价.Text) <> "" Then
            strSQL = "ZL_材料加成方案_INSERT(0,null,null,null," _
                   & byt计算方法 & "," _
                   & dbl限价 & ",'总限价')"
            zlDatabase.ExecuteProcedure strSQL, Me.Caption
        End If
        
        For intRow = 1 To .Rows - 1
            If Val(.TextMatrix(intRow, marBillCol.最低价)) <> 0 Or Val(.TextMatrix(intRow, marBillCol.最高价)) <> 0 Then           '先判有否数据
                
                str说明 = Trim(.TextMatrix(intRow, marBillCol.说明))
                str说明 = IIf(str说明 = "", "Null", "'" & str说明 & "'")
                dbl最低价 = Val(.TextMatrix(intRow, marBillCol.最低价))
                dbl最高价 = Val(.TextMatrix(intRow, marBillCol.最高价))
                dbl加成率 = Val(.TextMatrix(intRow, marBillCol.加成率))
                dbl限价 = Val(.TextMatrix(intRow, marBillCol.分段最高限价))
                
                '存储过程的参数如下:
                'ZL_材料加成方案_INSERT(
                '  序号_In     In 材料加成方案.序号%Type,
                '  最低价_In   In 材料加成方案.最低价%Type,
                '  最高价_In   In 材料加成方案.最高价%Type,
                '  加成率_In   In 材料加成方案.加成率%Type,
                '  计算方法_In In 材料加成方案.计算方法%Type,
                '  限价_In     In 材料加成方案.限价%Type,
                '  说明_In     In 材料加成方案.说明%Type
                
                strSQL = "ZL_材料加成方案_INSERT(" & _
                    intRow & "," & _
                    IIf(dbl最低价 = 0, "Null", dbl最低价) & "," & _
                    IIf(dbl最高价 = 0, "Null", dbl最高价) & "," & _
                    IIf(dbl加成率 = 0, "Null", dbl加成率) & "," & _
                    byt计算方法 & "," & _
                    dbl限价 & "," & _
                    str说明 & ")"
                    
                zlDatabase.ExecuteProcedure strSQL, Me.Caption
            End If
        Next
        gcnOracle.CommitTrans
        mblnReturn = True
        mblnChange = False
    End With
    SaveCard = True
    Exit Function
ErrHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub initGrid()
    With mshBill
        .Active = True
        .Cols = 6
        .MsfObj.FixedCols = 1
        .TextMatrix(0, marBillCol.序号) = "序号"
        .TextMatrix(0, marBillCol.最低价) = "最低价"
        .TextMatrix(0, marBillCol.最高价) = "最高价"
        .TextMatrix(0, marBillCol.加成率) = "加成率"
        .TextMatrix(0, marBillCol.分段最高限价) = "分段最高限价"
        .TextMatrix(0, marBillCol.说明) = "说明"
        
        .ColWidth(marBillCol.序号) = 600
        .ColWidth(marBillCol.最低价) = 1400
        .ColWidth(marBillCol.最高价) = 1400
        .ColWidth(marBillCol.加成率) = 1000
        .ColWidth(marBillCol.分段最高限价) = 1400
        .ColWidth(marBillCol.说明) = 2000
        
        '-1：表示该列可以选择，是布尔型［"√"，" "］
        ' 0：表示该列可以选择，但不能修改
        ' 1：表示该列可以输入，外部显示为按钮选择
        ' 2：表示该列是日期列，外部显示为按钮选择，弹出是日期选择框
        ' 3：表示该列是选择列，外部显示为下拉框选择
        '4:  表示该列为单纯的文本框供用户输入
        '5:  表示该列不允许选择
            
        .ColData(marBillCol.序号) = 5
 
        .ColData(marBillCol.最低价) = 4
        .ColData(marBillCol.最高价) = 4
        .ColData(marBillCol.加成率) = 4
        .ColData(marBillCol.分段最高限价) = 4
        .ColData(marBillCol.说明) = 4
        
        .ColAlignment(marBillCol.最低价) = flexAlignRightCenter
        .ColAlignment(marBillCol.最高价) = flexAlignRightCenter
        .ColAlignment(marBillCol.加成率) = flexAlignRightCenter
        .ColAlignment(marBillCol.分段最高限价) = flexAlignRightCenter
        .ColAlignment(marBillCol.说明) = flexAlignLeftCenter
        
        .PrimaryCol = marBillCol.最高价
        .LocateCol = marBillCol.最高价
    End With
End Sub
Private Sub CmdHelp_Click()
        ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub

    mblnFirst = False
    err = 0
    On Error Resume Next
    '加载卡片信息
    If LoadCardInfor = False Then
        Unload Me
        Exit Sub
    End If
   mblnChange = False
    SetCtlEnable
End Sub
Private Sub Form_Load()
    mblnFirst = True
    Call initGrid
    With cbo计算方法
        .Clear
        .AddItem "0-整体计算"
        .ItemData(.NewIndex) = 0
        .ListIndex = .NewIndex
        .AddItem "1-分段计算"
        .ItemData(.NewIndex) = 1
    End With
     
    '刘兴宏:增加小数格式化串
    With mFMT
        .FM_成本价 = GetFmtString(0, g_成本价)
        .FM_金额 = GetFmtString(0, g_金额)
        .FM_零售价 = GetFmtString(0, g_售价)
        .FM_数量 = GetFmtString(0, g_数量)
    End With
    
        
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim intTmp As Integer

    '验证输入的值的合法性
    If ValidData() = False Then Exit Sub
    
    '保存办所输入的值
    If SaveCard() = False Then Exit Sub
    mblnReturn = True
    Unload Me
End Sub

Private Function LoadCardInfor() As Boolean
    '---------------------------------------------------------------------------------------------------------
    '功能:加载要修改的卡片信息
    '返回:加载成功,返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------
    '只有对修改才起作用
    Dim intRow As Integer
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    LoadCardInfor = False
    mblnChange = False
    
    err = 0
    On Error GoTo ErrHand:
    
    strSQL = "select 序号,最低价,最高价,加成率,计算方法,限价,说明 from 材料加成方案 order by 序号"
    
    zlDatabase.OpenRecordset rsTmp, strSQL, Me.Caption
    Call initGrid
    If Not rsTmp.EOF Then
        For intRow = 0 To cbo计算方法.ListCount - 1
            If cbo计算方法.ItemData(intRow) = Val(zlStr.Nvl(rsTmp!计算方法)) Then
                cbo计算方法.ListIndex = intRow
                Exit For
            End If
        Next
        txt限价.Text = Format(Val(zlStr.Nvl(rsTmp!限价)), mFMT.FM_成本价)
    End If
    
    If cbo计算方法.ListIndex < 0 Then cbo计算方法.ListIndex = 0
    With mshBill
        .ClearBill
        .Rows = 2
        intRow = 1
        If Not rsTmp.EOF Then
            If rsTmp!序号 = 0 Then rsTmp.MoveNext
            Do While Not rsTmp.EOF
                .TextMatrix(intRow, marBillCol.序号) = zlStr.Nvl(rsTmp!序号, 0)
                .TextMatrix(intRow, marBillCol.最低价) = Format(Val(zlStr.Nvl(rsTmp!最低价)), mFMT.FM_成本价)
                .TextMatrix(intRow, marBillCol.最高价) = Format(Val(zlStr.Nvl(rsTmp!最高价)), mFMT.FM_成本价)
                .TextMatrix(intRow, marBillCol.加成率) = Format(Val(zlStr.Nvl(rsTmp!加成率)), GFM_VBJCL)
                .TextMatrix(intRow, marBillCol.分段最高限价) = Format(Val(zlStr.Nvl(rsTmp!限价)), mFMT.FM_成本价)
                .TextMatrix(intRow, marBillCol.说明) = zlStr.Nvl(rsTmp!说明)
                .Rows = .Rows + 1
                intRow = intRow + 1
                rsTmp.MoveNext
            Loop
        End If
    End With
    LoadCardInfor = True
    mblnChange = False
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub EditCard(ByVal frmMain As Object, ByVal strPriv As String, ByRef blnreturn As Boolean)

    '------------------------------------------------------------------------------------------------------
    '--功能:编辑卡片,用来与调用窗口进行通讯的程序
    '--入参数:str工资类型: 要编辑的表的主关键字
    '         strPriv:权限串
    '--出参数:BlnReturn,返回值,true代表增加或修改成功.false代表未新增或修改
    '--返回:
    '------------------------------------------------------------------------------------------------------
    mstrPriv = strPriv
    mblnChange = False
    mblnReturn = False
    SetCtlEnable
    frm加成率设置.Show 1, frmMain
    blnreturn = mblnReturn
End Sub
Private Sub SetCtlEnable()
    '------------------------------------------------------------------------------------------------------------------------
    '功能:设置控件的Enable属性
    '------------------------------------------------------------------------------------------------------------------------
    Dim blnSave As Boolean
    blnSave = Trim(mshBill.TextMatrix(1, marBillCol.最低价)) <> "" Or Trim(mshBill.TextMatrix(1, marBillCol.最高价)) <> ""
    Me.cmdOk.Enabled = blnSave And mblnChange = True
End Sub

Private Sub ReFormal()
    '------------------------------------------------------------------------------------------------------
    '--功能:重算开始工资和最高价
    '--参数:
    '--返回:
    '------------------------------------------------------------------------------------------------------
    Dim intRow As Integer
    Dim dbl间隔 As Double
    Dim dbl最低价 As Double
    Dim dbl最高价 As Double
    Dim dbl上行最高价 As Double
    
    With mshBill
        dbl间隔 = 0
        dbl上行最高价 = 0
        For intRow = 1 To .Rows - 1
            dbl最高价 = Val(.TextMatrix(intRow, marBillCol.最高价))
            dbl最低价 = Val(.TextMatrix(intRow, marBillCol.最低价))
            If dbl最高价 <> 0 Or dbl最低价 <> 0 Then
                .TextMatrix(intRow, marBillCol.序号) = intRow
                If intRow <> 1 Then
                    dbl间隔 = dbl最高价 - dbl最低价
                    If dbl上行最高价 <> dbl最低价 Then
                        '不等于最低价,重新算当前工资
                        .TextMatrix(intRow, marBillCol.最低价) = Format(dbl上行最高价, mFMT.FM_成本价)
                        If dbl最高价 <> 0 Then
                            .TextMatrix(intRow, marBillCol.最高价) = Format(dbl上行最高价 + dbl间隔, mFMT.FM_成本价)
                        End If
                    End If
                    dbl上行最高价 = Val(.TextMatrix(intRow, marBillCol.最高价))
                Else
                    dbl上行最高价 = dbl最高价
                End If
            End If
        Next
    End With
End Sub
Private Sub txt限价_Change()
 mblnChange = True
 SetCtlEnable
End Sub

Private Sub txt限价_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub txt限价_KeyPress(KeyAscii As Integer)
    Call zlControl.TxtCheckKeyPress(txt限价, KeyAscii, m金额式)
End Sub

Private Sub txt限价_LostFocus()
    txt限价.Text = Format(Val(txt限价.Text), mFMT.FM_成本价)
End Sub
