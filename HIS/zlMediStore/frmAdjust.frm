VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmAdjust 
   Caption         =   "药品调价单"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10110
   Icon            =   "frmAdjust.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6135
   ScaleWidth      =   10110
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox Chk定价 
      Caption         =   "时价药品改为定价销售(&D)"
      Enabled         =   0   'False
      Height          =   210
      Left            =   2505
      TabIndex        =   6
      Top             =   3450
      Width           =   2370
   End
   Begin MSComctlLib.ListView lvwItem 
      Height          =   2865
      Left            =   5790
      TabIndex        =   15
      Top             =   30
      Visible         =   0   'False
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   5054
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.TreeView tvwItem 
      Height          =   2880
      Left            =   5775
      TabIndex        =   17
      Top             =   15
      Visible         =   0   'False
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   5080
      _Version        =   393217
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin ZL9BillEdit.BillEdit bfgPrice 
      Height          =   2955
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   5212
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
   Begin VB.CommandButton cmdCpt 
      Caption         =   "按当前库存计算(&T)…"
      Height          =   350
      Left            =   6075
      Picture         =   "frmAdjust.frx":0442
      TabIndex        =   12
      Top             =   3915
      Width           =   1965
   End
   Begin VB.CommandButton cmdPstor 
      Caption         =   "打印库存变动表(&S)…"
      Height          =   350
      Left            =   8085
      Picture         =   "frmAdjust.frx":058C
      TabIndex        =   13
      Top             =   3915
      Width           =   1965
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "打印(&P)…"
      Height          =   350
      Left            =   8880
      Picture         =   "frmAdjust.frx":06D6
      TabIndex        =   11
      Top             =   900
      Width           =   1100
   End
   Begin VB.CommandButton cmdCanc 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   8880
      Picture         =   "frmAdjust.frx":0820
      TabIndex        =   10
      Top             =   480
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   8880
      Picture         =   "frmAdjust.frx":096A
      TabIndex        =   9
      Top             =   45
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -435
      TabIndex        =   16
      Top             =   3810
      Width           =   16815
   End
   Begin VB.TextBox txtRegistrar 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   300
      Left            =   6285
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   3015
      Width           =   2445
   End
   Begin MSComCtl2.DTPicker dtpRunDate 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "yyyy-MM-dd"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   0
      EndProperty
      Height          =   300
      Left            =   6285
      TabIndex        =   8
      Top             =   3390
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy年MM月dd日 HH:mm:ss"
      Format          =   127270915
      CurrentDate     =   36846.5833333333
   End
   Begin VB.TextBox txtSummary 
      Height          =   300
      Left            =   825
      TabIndex        =   2
      Top             =   3015
      Width           =   4485
   End
   Begin VB.CheckBox chkImmediately 
      Caption         =   "所有价格立即生效(&I)"
      Height          =   210
      Left            =   75
      TabIndex        =   5
      Top             =   3450
      Width           =   2040
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hgdStore 
      Height          =   1815
      Left            =   30
      TabIndex        =   14
      Top             =   4305
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   3201
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   14737632
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "库存变动表："
      Height          =   180
      Left            =   60
      TabIndex        =   18
      Top             =   4050
      Width           =   1080
   End
   Begin VB.Label lblRegistrar 
      AutoSize        =   -1  'True
      Caption         =   "调价人"
      Height          =   180
      Left            =   5655
      TabIndex        =   3
      Top             =   3075
      Width           =   540
   End
   Begin VB.Label lblRunDate 
      AutoSize        =   -1  'True
      Caption         =   "执行日期"
      Height          =   180
      Left            =   5475
      TabIndex        =   7
      Top             =   3450
      Width           =   720
   End
   Begin VB.Label lblSummary 
      AutoSize        =   -1  'True
      Caption         =   "调价说明"
      Height          =   180
      Left            =   30
      TabIndex        =   1
      Top             =   3075
      Width           =   720
   End
End
Attribute VB_Name = "frmAdjust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public blnImmediately As Boolean
'是否需要立即生效（时价药品调价需要）
Public lngBillId As Long
'功能类型:0-调价处理;其他-显示lngBillId确定的历史调价单
Public lngMediId As Long
'进入类型:0-未指定调价药品;其他-进入时直接显示lngMediId的原价格情况
Public intUnit As Integer   '0-售价单位;1-门诊单位;2-药库单位;3-住院单位
'进入类型:0-未指定调价药品;其他-进入时直接显示lngMediId的原价格情况
Private BlnModify As Boolean

'--------调价单列常量--------------
Const conCol药品id As Integer = 0
Const conCol品名 As Integer = 1
Const conCol规格 As Integer = 2
Const conCol产地 As Integer = 3
Const conCol单位 As Integer = 4
Const conCol原价 As Integer = 5
Const conCol现价 As Integer = 6
Const conCol收入ID As Integer = 7
Const conColOld收入ID As Integer = 8
Const conCol收入名称 As Integer = 9

'---------------------------------
Dim rsTemp As New ADODB.Recordset
Private StrFindStyle As String
Dim intCount As Integer
Dim objItem As ListItem
Dim objNode As Node
Dim dtToday As Date

Private Const mconintPriceBit As Integer = 7            '单价小数位数
Private Sub bfgPrice_CommandClick()
    Dim RecCheck As New ADODB.Recordset
    Dim LngmediIDThis As Long
    Dim RecReturn As New ADODB.Recordset
    
    On Error GoTo errHandle
    If Me.bfgPrice.Col = conCol品名 Then
'        With Me.bfgPrice
'            Set RecReturn = Frm药品选择器.ShowMe(Me, 1)
'            If RecReturn.EOF Then Exit Sub
'            LngmediIDThis = RecReturn!药品ID
'            If LngmediIDThis = 0 Then Exit Sub
'
'            '是变价药品则退出
'            gstrSQL = " Select Nvl(是否变价,0) 变价 From 药品规格 Where ID=[1]"
'            Set RecCheck = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, LngmediIDThis)
'            With RecCheck
'                '如果是时价药品，调价记录必须立即生效
'                If !变价 = 1 And blnImmediately = False Then
'                    blnImmediately = True
'                    chkImmediately.Value = 1
'                    chkImmediately.Enabled = False
'                End If
'                If !变价 = 1 Then Chk定价.Enabled = True
'            End With
'
'            If chkImmediately.Value = 1 Then
'
'                    '判断是否有未执行的历史价格
'                gstrSQL = " Select Count(*) Records From 收费价目 Where 变动原因=0 And 收费细目ID=[1]"
'                Set RecCheck = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, LngmediIDThis)
'
'                With RecCheck
'                    If Not .EOF Then
'                        If Not IsNull(!Records) Then
'                            If !Records <> 0 And chkImmediately.Value = 1 Then
'                                MsgBox "该药品存在未执行价格，不能设置为立即执行！", vbInformation, gstrSysName
'                                If chkImmediately.Enabled Then
'                                    chkImmediately.Value = 0
'                                Else
'                                    Exit Sub
'                                End If
'                            End If
'                        End If
'                    End If
'                End With
'            End If
'
'            .TextMatrix(.Row, conCol药品id) = RecReturn!药品ID
'            .TextMatrix(.Row, conCol品名) = "[" & RecReturn!药品编码 & "]" & RecReturn!商品名
'            .TextMatrix(.Row, conCol规格) = IIf(IsNull(RecReturn!规格), "", RecReturn!规格)
'            .TextMatrix(.Row, conCol产地) = IIf(IsNull(RecReturn!产地), "", RecReturn!产地)
'            .TextMatrix(.Row, conCol单位) = IIf(IsNull(RecReturn!售价单位), "", RecReturn!售价单位)
'            Call getMediPrice(RecReturn!药品ID)
'            .CmdVisible = False
'            .Col = conCol现价
'            BlnModify = True
'        End With
    Else
        With Me.tvwItem
            .Left = bfgPrice.Left + bfgPrice.MsfObj.CellLeft
            .Top = Me.bfgPrice.Top + Me.bfgPrice.CellTop + Me.bfgPrice.RowHeight(1)
            If Me.ScaleHeight - .Top < 3000 Then
                .Height = Me.ScaleHeight - .Top
            Else
                .Height = 3000
            End If
            .Visible = True
            For intCount = 1 To .Nodes.count
                If InStr(1, .Nodes(intCount).Text & "-", Me.bfgPrice.TextMatrix(Me.bfgPrice.Row, Me.bfgPrice.Col) & "-") > 0 Then
                    .Nodes(intCount).Selected = True
                    Exit For
                End If
            Next
            .SetFocus
        End With
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub bfgPrice_EnterCell(Row As Long, Col As Long)
    Select Case Col
    Case conCol收入名称
        Me.bfgPrice.TextMatrix(Row, conCol现价) = zlStr.FormatEx(Me.bfgPrice.TextMatrix(Row, conCol现价), 7)
    End Select
    
End Sub

Private Sub bfgPrice_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strInput As String
    If KeyCode <> 13 Then Exit Sub
    On Error GoTo errHandle
    Select Case Me.bfgPrice.Col
    Case conCol品名
        If Trim(Me.bfgPrice.Text) = "" Then Exit Sub
        strInput = UCase(Me.bfgPrice.Text)
        
        Me.lvwItem.Tag = conCol品名
        With Me.lvwItem.ColumnHeaders
            .Clear
            .Add , "编码", "编码", 900
            .Add , "通用名称", "通用名称", 2000
            .Add , "规格", "规格", 1200
            .Add , "产地", "产地", 1100
            .Add , "单位", "单位", 450
        End With
        
        gstrSQL = " Select Distinct D.药品ID,C.编码,NVL(A.名称,C.名称) 通用名称,C.规格,C.产地,C.计算单位 单位,Nvl(C.是否变价,0) 变价" & _
                      " From 收费项目别名 A,药品规格 D," & _
                      "     (Select B.* From 收费项目别名 A,收费项目目录 B" & _
                      "     Where A.收费细目ID=B.ID And B.类别 In ('5','6','7') " & _
                      "           ANd (A.简码 Like [1] Or A.名称 Like [1] Or B.编码 Like [1])) C " & _
                      " WHERE D.药品ID=C.ID And (C.撤档时间 Is Null Or C.撤档时间=To_Date('3000-01-01','yyyy-MM-dd'))" & _
                      " And D.药品ID=A.收费细目ID(+) and A.性质(+)=3 and A.码类(+)=1"
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption, StrFindStyle & strInput & "%")
                       
        With rsTemp
            If .EOF Then
                MsgBox "未找到相关药品，请重新输入！", vbInformation, gstrSysName
                Cancel = True
                bfgPrice.TxtSetFocus
                Exit Sub
            End If
            
            '如果是时价药品，调价记录必须立即生效
            If !变价 = 1 And blnImmediately = False Then
                blnImmediately = True
                chkImmediately.Value = 1
                chkImmediately.Enabled = False
            End If
            If !变价 = 1 Then Chk定价.Enabled = True
            
            If chkImmediately.Value = 1 Then
                Dim RecCheck As New ADODB.Recordset
                
                gstrSQL = " Select Count(*) Records From 收费价目 Where 变动原因=0 And 收费细目ID=[1] " & _
                        GetPriceClassString("")
                
                Set RecCheck = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption, CLng(rsTemp!药品id))
                
                With RecCheck
                    If Not .EOF Then
                        If Not IsNull(!Records) Then
                            If !Records <> 0 And chkImmediately.Value = 1 Then
                                MsgBox "该药品存在未执行价格，不能设置为立即执行！", vbInformation, gstrSysName
                                If chkImmediately.Enabled Then
                                    chkImmediately.Value = 0
                                Else
                                    Cancel = True
                                    bfgPrice.TxtSetFocus
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                End With
            End If
            
            Me.lvwItem.ListItems.Clear
            Do While Not .EOF
                Set objItem = Me.lvwItem.ListItems.Add(, "_" & !药品id, !编码)
                objItem.SubItems(1) = !通用名称
                objItem.SubItems(2) = IIf(IsNull(!规格), "", !规格)
                objItem.SubItems(3) = IIf(IsNull(!产地), "", !产地)
                objItem.SubItems(4) = IIf(IsNull(!单位), "", !单位)
                If Me.lvwItem.SelectedItem Is Nothing Then
                    objItem.Selected = True
                End If
                .MoveNext
            Loop
            If Me.lvwItem.ListItems.count = 0 Then
                MsgBox "该药品不存在", vbExclamation, gstrSysName
                Cancel = True
                Exit Sub
            ElseIf Me.lvwItem.ListItems.count = 1 Then
                lvwItem_DblClick
                Cancel = True
                Exit Sub
            End If
        End With
        With Me.lvwItem
            .Left = Me.bfgPrice.Left
            .Top = Me.bfgPrice.Top + Me.bfgPrice.CellTop + Me.bfgPrice.RowHeight(1)
            If Me.ScaleHeight - .Top < 3000 Then
                .Height = Me.ScaleHeight - .Top
            Else
                .Height = 3000
            End If
            .Visible = True
            .SetFocus
        End With
        Cancel = True
    Case conCol收入名称
        If Trim(Me.bfgPrice.Text) = "" Then Exit Sub
        strInput = UCase(Me.bfgPrice.Text)
        
        Me.lvwItem.Tag = conCol收入名称
        With Me.lvwItem.ColumnHeaders
            .Clear
            .Add , "编码", "编码", 600
            .Add , "名称", "名称", 1000
        End With
        
        gstrSQL = "select id,编码,名称 from 收入项目 U" & _
                    " where rownum<100 and nvl(撤档时间,to_date('3000-01-01','YYYY-MM-DD'))>trunc(sysdate)+1" & _
                    "      and NOT exists(select 1 from 收入项目 D where D.上级id=U.id)" & _
                    "      and (编码 like [1] or 简码 like [1] or 名称 like [1])"
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption, strInput & "%")
        
        With rsTemp
            Me.lvwItem.ListItems.Clear
            Do While Not .EOF
                Set objItem = Me.lvwItem.ListItems.Add(, "_" & !Id, !编码)
                objItem.SubItems(1) = !名称
                If Me.lvwItem.SelectedItem Is Nothing Then
                    objItem.Selected = True
                End If
                .MoveNext
            Loop
            If Me.lvwItem.ListItems.count = 0 Then
                MsgBox "该项目不存在", vbExclamation, gstrSysName
                Cancel = True
                Exit Sub
            ElseIf Me.lvwItem.ListItems.count = 1 Then
                lvwItem_DblClick
                Cancel = True
                Exit Sub
            End If
        End With
        With Me.lvwItem
            .Left = bfgPrice.Left + bfgPrice.MsfObj.CellLeft
            .Top = Me.bfgPrice.Top + Me.bfgPrice.CellTop + Me.bfgPrice.RowHeight(1)
            If Me.ScaleHeight - .Top < 2000 Then
                .Height = Me.ScaleHeight - .Top
            Else
                .Height = 2000
            End If
            .Visible = True
            .SetFocus
        End With
        Cancel = True
    Case conCol现价
        Dim lng药品ID As Long
        With bfgPrice
            lng药品ID = Val(bfgPrice.TextMatrix(bfgPrice.Row, conCol药品id))
            If lng药品ID = 0 Then Exit Sub
            
            '现价不能大于指导零售价
            gstrSQL = " Select Nvl(指导零售价,0) 指导零售价 From 药品规格 Where 药品ID=[1]"
            Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption, lng药品ID)
            
            If Val(.Text) > rsTemp!指导零售价 Then
                MsgBox "现价不能大于指导零售价！（" & Format(rsTemp!指导零售价, "#####0.0000000;-#####0.0000000; ;") & "）"
                Cancel = True
                .TxtSetFocus
            End If
            BlnModify = True
        End With
    End Select
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub chkImmediately_Click()
    Dim RecCheck As New ADODB.Recordset
    Dim LngmediIDThis As Long, IntCheck As Integer
    
    On Error GoTo errHandle
    If chkImmediately.Value = 1 Then
        '循环判断所有药品
        For IntCheck = 1 To bfgPrice.rows - 1
            LngmediIDThis = Val(bfgPrice.TextMatrix(IntCheck, conCol药品id))
            If LngmediIDThis <> 0 Then
                '判断是否有未执行的历史价格
                
                 gstrSQL = " Select Count(*) Records From 收费价目 Where 变动原因=0 And 收费细目ID=[1]" & _
                 GetPriceClassString("")
                 
                 Set RecCheck = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption, LngmediIDThis)
                 
                 With RecCheck
                    If Not .EOF Then
                        If Not IsNull(!Records) Then
                            If !Records <> 0 Then
                                MsgBox "药品" & bfgPrice.TextMatrix(IntCheck, conCol品名) & "存在未执行价格，不能设置为立即执行！", vbInformation, gstrSysName
                                chkImmediately.Value = 0
                                Exit Sub
                            End If
                        End If
                    End If
                End With
            End If
        Next
    End If
    
    If Me.chkImmediately.Value Then
        Me.dtpRunDate.Enabled = False
    Else
        Me.dtpRunDate.Enabled = True
    End If
    
    On Error Resume Next
    Me.bfgPrice.SetFocus
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdCanc_Click()
    lngBillId = 0
    lngMediId = 0
    Unload Me
End Sub



Private Sub cmdOK_Click()
'    Dim strID As String, LngCurID As Long
'    Dim ArrayID
'    Dim lngAdjId As Long
'    Dim strOldId As String
'    Dim strNewId As String
'
'    '检测相关输入合法性
'    If CheckPrice = False Then Exit Sub
'    '如果即时执行，则调用过程zl_药品收发记录_Adjust
'
'    dtToday = Sys.Currentdate()
'    Err = 0
'    On Error GoTo ErrHand
'    With rsTemp
'        gstrSQL = "select 收费价目_ID.nextval from dual"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.Title, Me.Caption, gstrSQL)
'        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "cmdOK_Click")
'        Call SQLTest
'
'        lngAdjId = .Fields(0).Value
'    End With
'
'    gcnOracle.BeginTrans
'    With Me.bfgPrice
'        strOldId = ""
'        strNewId = ""
'        strID = ""
'        For intCount = 1 To IIf(Trim(.TextMatrix(.rows - 1, 0)) = "", .rows - 2, .rows - 1)
'            LngCurID = Sys.NextId("收费价目")
'            strID = strID & IIf(strID = "", "", ",") & LngCurID
'            If CLng(.RowData(intCount)) <> 0 Then
'                If .RowData(intCount) <> -1 And InStr(1, strOldId & ",", "," & .RowData(intCount) & ",") > 0 Then
'                    gcnOracle.RollbackTrans: .SetFocus: Exit Sub
'                    MsgBox "在一次调价中不能对相同品种(" & .TextMatrix(intCount, conCol品名) & ")重复调价", vbExclamation, gstrSysName
'                End If
'                If .RowData(intCount) = -1 And InStr(1, strNewId & ",", "," & .TextMatrix(intCount, conCol药品id) & ",") > 0 Then
'                    gcnOracle.RollbackTrans: .SetFocus: Exit Sub
'                    MsgBox "不能对相同品种(" & .TextMatrix(intCount, conCol品名) & ")重复设置价格", vbExclamation, gstrSysName
'                End If
''                If .TextMatrix(intCount, conCol原价) = .TextMatrix(intCount, conCol现价) Then
''                    MsgBox .TextMatrix(intCount, conCol品名) & " 现价未调整，请检查", vbExclamation, gstrSysName
''                    gcnOracle.RollbackTrans:.SetFocus:Exit Sub
''                End If
'                If .RowData(intCount) <> -1 Then
'                    strOldId = strOldId & "," & .RowData(intCount)
'                Else
'                    strNewId = strNewId & "," & .TextMatrix(intCount, conCol药品id)
'                End If
'
'                If Val(.TextMatrix(intCount, conCol现价)) <> 0 Then
'                    '设置上一次的价格记录终止执行
'                    gstrSQL = "zl_收费价目_stop(" & .TextMatrix(intCount, conCol药品id) & ","
'                    If Me.chkImmediately.Value Then
'                        gstrSQL = gstrSQL & "to_date('" & Format(DateAdd("s", -1, dtToday), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
'                    Else
'                        gstrSQL = gstrSQL & "to_date('" & Format(DateAdd("s", -1, Me.dtpRunDate.Value), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
'                    End If
'                    gstrSQL = gstrSQL & ")"
'                    Call zlDataBase.ExecuteProcedure(gstrSQL, Me.Caption)
'
'                    '产生价格记录
'                    gstrSQL = "zl_收费价目_Insert(" & LngCurID & "," & IIf(.RowData(intCount) = -1, "NUll", .RowData(intCount)) & _
'                              "," & .TextMatrix(intCount, conCol药品id) & "," & Val(.TextMatrix(intCount, conColOld收入ID)) & "," & Val(.TextMatrix(intCount, conCol原价)) & "," & Val(.TextMatrix(intCount, conCol现价)) & _
'                              ",NULL,NULL,'" & Me.txtSummary.Text & "'," & lngAdjId & ",'" & Trim(Me.txtRegistrar.Text) & "',"
'                    If Me.chkImmediately.Value Then
'                        gstrSQL = gstrSQL & "to_date('" & Format(dtToday, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
'                    Else
'                        gstrSQL = gstrSQL & "to_date('" & Format(Me.dtpRunDate.Value, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
'                    End If
'                    gstrSQL = gstrSQL & ",0)"
'                    Call zlDataBase.ExecuteProcedure(gstrSQL, Me.Caption)
'                End If
'
'            End If
'        Next
'    End With
'
'    '循环执行过程
'    ArrayID = Split(strID, ",")
'    For intCount = 0 To UBound(ArrayID)
'        If Me.chkImmediately.Value Or bfgPrice.RowData(intCount + 1) = -1 Then
'            gstrSQL = "zl_药品收发记录_Adjust(" & ArrayID(intCount) & "," & Chk定价.Value & ")"
'            Call zlDataBase.ExecuteProcedure(gstrSQL, Me.Caption & "-产生价格调整记录")
'        End If
'    Next
'
'    gcnOracle.CommitTrans
'    lngBillId = 0
'    lngMediId = 0
'
'    BlnModify = False
'    Unload Me
    Exit Sub
    
ErrHand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    Me.bfgPrice.SetFocus
End Sub

Private Sub cmdCpt_Click()
    Dim lngMediId As Long
    Dim dblOldPrice As Double
    Dim dblNewPrice As Double
    
    Dim intRows As Integer
    
    Me.hgdStore.Redraw = False
    Me.hgdStore.rows = 1
    
    On Error GoTo errHandle
    With Me.bfgPrice
        For intCount = 1 To .rows - 1
            lngMediId = Val(.TextMatrix(intCount, conCol药品id))
            dblOldPrice = Val(.TextMatrix(intCount, conCol原价))
            dblNewPrice = Val(.TextMatrix(intCount, conCol现价))
            If lngMediId <> 0 And dblOldPrice <> dblNewPrice Then
                gstrSQL = "SELECT DISTINCT D.名称 AS 库房,'['||C.编码||']'||NVL(A.名称,C.名称) AS 药品," & _
                    "      C.规格,C.产地,C.计算单位 AS 单位,S.批号,S.数量,S.批次" & _
                    " FROM " & _
                    "      (SELECT S.库房ID,S.药品ID,S.上次批号 批号,S.实际数量 AS 数量,S.批次" & _
                    "      FROM 药品库存 S" & _
                    "      WHERE S.实际数量<>0 AND S.药品ID=[1] And S.性质=1) S, " & _
                    "      部门表 D,药品规格 M,收费项目别名 A,收费项目目录 C" & _
                    " WHERE D.ID=S.库房ID AND M.药品ID=C.ID AND S.药品ID=M.药品ID     " & _
                    " AND M.药品ID=A.收费细目ID(+) AND A.性质(+)=3 AND A.码类(+)=1" & _
                    " ORDER BY '['||C.编码||']'||NVL(A.名称,C.名称),S.批号"
                Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption, lngMediId)
                
                With rsTemp
                    intRows = Me.hgdStore.rows
                    Me.hgdStore.rows = Me.hgdStore.rows + .RecordCount
                    Do While Not .EOF
                        Me.hgdStore.TextMatrix(intRows + .AbsolutePosition - 1, 0) = !库房
                        Me.hgdStore.TextMatrix(intRows + .AbsolutePosition - 1, 1) = !药品
                        Me.hgdStore.TextMatrix(intRows + .AbsolutePosition - 1, 2) = IIf(IsNull(!规格), "", !规格) & IIf(IsNull(!产地), "", "|" & !产地)
                        Me.hgdStore.TextMatrix(intRows + .AbsolutePosition - 1, 3) = IIf(IsNull(!单位), "", !单位)
                        Me.hgdStore.TextMatrix(intRows + .AbsolutePosition - 1, 4) = IIf(IsNull(!批号), "", !批号)
                        Me.hgdStore.TextMatrix(intRows + .AbsolutePosition - 1, 5) = Format(!数量, "0.00000")
                        Me.hgdStore.TextMatrix(intRows + .AbsolutePosition - 1, 6) = zlStr.FormatEx(dblOldPrice, mconintPriceBit)
                        Me.hgdStore.TextMatrix(intRows + .AbsolutePosition - 1, 7) = zlStr.FormatEx(dblNewPrice, mconintPriceBit)
                        Me.hgdStore.TextMatrix(intRows + .AbsolutePosition - 1, 8) = Format(!数量 * (dblNewPrice - dblOldPrice), "0.00")
                        .MoveNext
                    Loop
                End With
            End If
        Next
    End With
    
    If Me.hgdStore.rows < 2 Then
        Me.hgdStore.rows = 2
    End If
    Me.hgdStore.FixedRows = 1
    Me.hgdStore.Redraw = True
    If Me.bfgPrice.Active Then Me.bfgPrice.SetFocus
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdPrint_Click()
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    
    If Trim(Me.bfgPrice.TextMatrix(1, 0)) = "" Then Exit Sub
    objPrint.Title.Text = "药品调价通知单"
    
    Set objRow = New zlTabAppRow
    objRow.Add "调价说明:" & Me.txtSummary.Text
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "执行时间:" & Format(IIf(Me.chkImmediately.Value, Sys.Currentdate, Me.dtpRunDate.Value), "yyyy年MM月DD日 HH:mm:ss")
    objRow.Add "调价人:" & Me.txtRegistrar.Text
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "打印人:" & gstrUserName
    objRow.Add "打印时间:" & Format(Sys.Currentdate, "yyyy年MM月DD日 HH:mm:ss")
    objPrint.BelowAppRows.Add objRow
    
    Set objPrint.Body = Me.bfgPrice.MsfObj
    objPrint.PageFooter = 2
     
    Select Case zlPrintAsk(objPrint)
    Case 1
         zlPrintOrView1Grd objPrint, 1
    Case 2
        zlPrintOrView1Grd objPrint, 2
    Case 3
        zlPrintOrView1Grd objPrint, 3
    End Select
    Set objPrint = Nothing

End Sub

Private Sub cmdPstor_Click()
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    
    If Me.cmdCpt.Enabled Then
        Call cmdCpt_Click
    End If
    If Trim(Me.hgdStore.TextMatrix(1, 0)) = "" Then Exit Sub
    
    objPrint.Title.Text = "调价库存变动表"
    
    Set objRow = New zlTabAppRow
    objRow.Add "调价说明:" & Me.txtSummary.Text
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "执行时间:" & Format(IIf(Me.chkImmediately.Value, Sys.Currentdate, Me.dtpRunDate.Value), "yyyy年MM月DD日 HH:mm:ss")
    objRow.Add "调价人:" & Me.txtRegistrar.Text
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "打印人:" & gstrUserName
    objRow.Add "打印时间:" & Format(Sys.Currentdate, "yyyy年MM月DD日 HH:mm:ss")
    objPrint.BelowAppRows.Add objRow
    
    Set objPrint.Body = Me.hgdStore
    objPrint.PageFooter = 2
     
    Select Case zlPrintAsk(objPrint)
    Case 1
         zlPrintOrView1Grd objPrint, 1
    Case 2
        zlPrintOrView1Grd objPrint, 2
    Case 3
        zlPrintOrView1Grd objPrint, 3
    End Select
    Set objPrint = Nothing

End Sub

Private Sub dtpRunDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Me.cmdOk.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyEscape Then
        If Me.ActiveControl.Name = "tvwItem" Then
            tvwItem.Visible = False
            bfgPrice.SetFocus
        ElseIf Me.ActiveControl.Name = "lvwItem" Then
            lvwItem.Visible = False
            bfgPrice.SetFocus
        Else
            cmdCanc_Click
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    RestoreWinState Me
    BlnModify = False
    StrFindStyle = IIf(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = "0", "%", "")
    
    On Error GoTo errHandle
    With Me.hgdStore
        .rows = 2
        .Cols = 9
        .Redraw = False
        .TextMatrix(0, 0) = "库房"
        .TextMatrix(0, 1) = "药品"
        .TextMatrix(0, 2) = "规格|产地"
        .TextMatrix(0, 3) = "单位"
        .TextMatrix(0, 4) = "批号"
        .TextMatrix(0, 5) = "数量"
        .TextMatrix(0, 6) = "原价"
        .TextMatrix(0, 7) = "现价"
        .TextMatrix(0, 8) = "调整金额"
        
        .ColWidth(0) = 800
        .ColWidth(1) = 2800
        .ColWidth(2) = 1350
        .ColWidth(3) = 400
        .ColWidth(4) = 800
        .ColWidth(5) = 1000
        .ColWidth(6) = 900
        .ColWidth(7) = 900
        .ColWidth(8) = 1050
    
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 4
        .ColAlignment(4) = 1
        .ColAlignment(5) = 7
        .ColAlignment(6) = 7
        .ColAlignment(7) = 7
        .ColAlignment(8) = 7
    
        .ColAlignmentFixed(0) = 4
        .ColAlignmentFixed(1) = 4
        .ColAlignmentFixed(2) = 4
        .ColAlignmentFixed(3) = 4
        .ColAlignmentFixed(4) = 4
        .ColAlignmentFixed(5) = 4
        .ColAlignmentFixed(6) = 4
        .ColAlignmentFixed(7) = 4
        .ColAlignmentFixed(8) = 4
        .Redraw = True
    End With
    
    With Me.bfgPrice
        .Cols = 10
        .MsfObj.FixedCols = 0
        .TextMatrix(0, conCol药品id) = "药品id"
        .TextMatrix(0, conCol品名) = "品名"
        .TextMatrix(0, conCol规格) = "规格"
        .TextMatrix(0, conCol产地) = "产地"
        .TextMatrix(0, conCol单位) = "单位"
        .TextMatrix(0, conCol原价) = "原价"
        .TextMatrix(0, conCol现价) = "现价"
        .TextMatrix(0, conCol收入ID) = "收入id"
        .TextMatrix(0, conColOld收入ID) = "原收入id"
        .TextMatrix(0, conCol收入名称) = "收入项目"
        
        .ColWidth(conCol药品id) = 0
        .ColWidth(conCol品名) = 2800
        .ColWidth(conCol规格) = 1200
        .ColWidth(conCol产地) = 1000
        .ColWidth(conCol单位) = 400
        .ColWidth(conCol原价) = 975
        .ColWidth(conCol现价) = 1000
        .ColWidth(conCol收入ID) = 0
        .ColWidth(conColOld收入ID) = 0
        .ColWidth(conCol收入名称) = 1000
        
        .ColData(conCol药品id) = 5
        .ColData(conCol品名) = 1
        .ColData(conCol规格) = 5
        .ColData(conCol产地) = 5
        .ColData(conCol单位) = 5
        .ColData(conCol原价) = 5
        .ColData(conCol现价) = 4
        .ColData(conCol收入ID) = 5
        .ColData(conColOld收入ID) = 5
        .ColData(conCol收入名称) = 1

        .ColAlignment(conCol药品id) = 1
        .ColAlignment(conCol品名) = 1
        .ColAlignment(conCol规格) = 1
        .ColAlignment(conCol产地) = 1
        .ColAlignment(conCol单位) = 4
        .ColAlignment(conCol原价) = 7
        .ColAlignment(conCol现价) = 7
        .ColAlignment(conCol收入ID) = 1
        .ColAlignment(conColOld收入ID) = 1
        .ColAlignment(conCol收入名称) = 1

        .PrimaryCol = conCol品名
        .LocateCol = conCol品名
    End With
    
    Dim StrToday As String
    StrToday = Format(Sys.Currentdate(), "yyyy-MM-dd hh:mm:ss")
    If lngBillId = 0 Then
        '进入调价编辑状态
        Me.bfgPrice.Active = True
        If blnImmediately Then
            Me.chkImmediately.Value = 1
            Me.chkImmediately.Enabled = False
        End If
        
        Me.lblTitle.Caption = "库存变动表：(由于调价未保存，反映的库存可能不准确)"
        Me.dtpRunDate.MinDate = DateAdd("s", 1, Format(Sys.Currentdate, "yyyy-MM-dd"))
        Me.dtpRunDate.Value = DateAdd("d", 1, Format(Sys.Currentdate, "YYYY-MM-DD"))
        Me.txtRegistrar.Text = gstrUserName
        With rsTemp
            
            gstrSQL = "select id,上级id,编码,名称 from 收入项目 start with 上级id is null connect by  prior id=上级id order by level"
            If .State = adStateOpen Then .Close
            
            Call SQLTest(App.Title, Me.Caption, gstrSQL)
            Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "Form_Load")
            Call SQLTest
            
            Me.tvwItem.Nodes.Clear
            Do While Not .EOF
                If IsNull(!上级ID) Then
                    Me.tvwItem.Nodes.Add , , "_" & !Id, !编码 & "_" & !名称
                Else
                    Me.tvwItem.Nodes.Add "_" & !上级ID, 4, "_" & !Id, !编码 & "_" & !名称
                End If
                .MoveNext
            Loop
        End With
        
        If lngMediId = 0 Then Exit Sub
        '如果指定首先调价的药品，则直接将该药品调入
            
         gstrSQL = " Select Distinct D.药品ID,'['||C.编码||']'||NVL(A.名称,C.名称) 品名,C.规格,C.产地,C.计算单位 单位" & _
                  " From 收费项目别名 A,药品规格 D,收费项目目录 C " & _
                  " Where D.药品ID =[1] And D.药品ID=C.ID" & _
                  " And D.药品ID=A.收费细目ID(+) And A.性质(+)=3 And A.码类(+)=1"
         Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption, lngMediId)
         
         With rsTemp
            If .RecordCount = 0 Then Exit Sub
            Me.bfgPrice.TextMatrix(.AbsolutePosition, conCol药品id) = !药品id
            Me.bfgPrice.TextMatrix(.AbsolutePosition, conCol品名) = !品名
            Me.bfgPrice.TextMatrix(.AbsolutePosition, conCol规格) = IIf(IsNull(!规格), "", !规格)
            Me.bfgPrice.TextMatrix(.AbsolutePosition, conCol产地) = IIf(IsNull(!产地), "", !产地)
            Me.bfgPrice.TextMatrix(.AbsolutePosition, conCol单位) = IIf(IsNull(!单位), "", !单位)
            Call getMediPrice(lngMediId)
            Me.bfgPrice.Col = conCol现价
        End With
    
    Else
        '进入调价显示状态
        Me.bfgPrice.Active = False
        Me.cmdOk.Visible = False
        Me.cmdCanc.Caption = "返回(&C)"
        Me.cmdCanc.Top = Me.cmdOk.Top
        Me.txtSummary.Enabled = False
        Me.chkImmediately.Value = 0
        Me.chkImmediately.Enabled = False
        Me.dtpRunDate.Enabled = False
        
        Dim strBills As String
        Dim strUnit As String
        
        strBills = ""
        
        Select Case intUnit
            Case 1
                strUnit = ",C.计算单位 AS 单位,P.原价,P.现价"
            Case 2
                strUnit = ",M.门诊单位 AS 单位,P.原价*M.门诊包装 As 原价,P.现价*M.门诊包装 As 现价"
            Case 3
                strUnit = ",M.药库单位 AS 单位,P.原价*M.药库包装 As 原价,P.现价*M.药库包装 As 现价"
            Case 4
                strUnit = ",M.住院单位 AS 单位,P.原价*M.住院包装 As 原价,P.现价*M.住院包装 As 现价"
        End Select
        
        gstrSQL = "SELECT DISTINCT P.ID,M.药品ID,'['||C.编码||'] '||NVL(A.名称,C.名称) AS 品名,C.规格,C.产地 " & strUnit & _
            "      ,P.收入项目ID,I.名称 AS 收入名称,TO_CHAR(P.执行日期,'YYYY-MM-DD HH24:MI:SS') 执行日期,P.变动原因,P.调价说明,P.调价人" & _
            " FROM 收费价目 P,药品规格 M,收入项目 I,收费项目别名 A,收费项目目录 C" & _
            " WHERE P.收费细目ID=M.药品ID AND P.收入项目ID=I.ID AND M.药品ID=C.ID " & _
            " AND M.药品ID=A.收费细目ID(+) AND A.性质(+)=3 AND A.码类(+)=1" & _
            " AND P.ID=[1] " & _
            GetPriceClassString("P") & _
            " ORDER BY P.ID"                            '因调价ID取的是价格记录ID的上一个ID
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption, lngBillId)
        
        Me.bfgPrice.rows = rsTemp.RecordCount + 1
        Do While Not rsTemp.EOF
            strBills = strBills & "," & rsTemp!Id
            Me.bfgPrice.TextMatrix(rsTemp.AbsolutePosition, conCol药品id) = rsTemp!药品id
            Me.bfgPrice.TextMatrix(rsTemp.AbsolutePosition, conCol品名) = rsTemp!品名
            Me.bfgPrice.TextMatrix(rsTemp.AbsolutePosition, conCol规格) = IIf(IsNull(rsTemp!规格), "", rsTemp!规格)
            Me.bfgPrice.TextMatrix(rsTemp.AbsolutePosition, conCol产地) = IIf(IsNull(rsTemp!产地), "", rsTemp!产地)
            Me.bfgPrice.TextMatrix(rsTemp.AbsolutePosition, conCol单位) = IIf(IsNull(rsTemp!单位), "", rsTemp!单位)
            Me.bfgPrice.TextMatrix(rsTemp.AbsolutePosition, conCol原价) = zlStr.FormatEx(rsTemp!原价, mconintPriceBit)
            Me.bfgPrice.TextMatrix(rsTemp.AbsolutePosition, conCol现价) = zlStr.FormatEx(rsTemp!现价, mconintPriceBit)
            Me.bfgPrice.TextMatrix(rsTemp.AbsolutePosition, conCol收入ID) = rsTemp!收入项目ID
            Me.bfgPrice.TextMatrix(rsTemp.AbsolutePosition, conCol收入名称) = rsTemp!收入名称
            Me.txtSummary = IIf(IsNull(rsTemp!调价说明), "", rsTemp!调价说明)
            Me.txtRegistrar.Text = IIf(IsNull(rsTemp!调价人), "", rsTemp!调价人)
            Me.dtpRunDate.Value = rsTemp!执行日期
            
            If rsTemp!执行日期 <= StrToday And rsTemp!变动原因 = 0 Then        '未进行调价计算,则执行计算
                gstrSQL = "zl_药品收发记录_Adjust(" & Val(rsTemp!药品id) & ")"
                Call zlDataBase.ExecuteProcedure(gstrSQL, Me.Caption & "-产生价格调整记录")
            End If
            rsTemp.MoveNext
        Loop
        If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
        
        If rsTemp!执行日期 > StrToday Then
            '如果执行时间未到，则只能模拟显示库存变动
            Me.lblTitle.Caption = "库存变动表：(由于执行时间未到，反映的库存可能不准确)"
            Call cmdCpt_Click
        Else
            '执行时间已到，肯定也进行了调价计算，直接从收发记录提取调价变动情况
            Me.cmdCpt.Enabled = False
            Me.lblTitle.Caption = "库存变动表："

            Select Case intUnit
                Case 1
                    strUnit = ",C.计算单位 AS 单位,S.原价,S.现价,S.数量"
                Case 2
                    strUnit = ",M.门诊单位 AS 单位,S.原价*M.门诊包装 As 原价,S.现价*M.门诊包装 As 现价,S.数量/M.门诊包装 As 数量"
                Case 3
                    strUnit = ",M.药库单位 AS 单位,S.原价*M.药库包装 As 原价,S.现价*M.药库包装 As 现价,S.数量/M.药库包装 As 数量"
                Case 4
                    strUnit = ",M.住院单位 AS 单位,S.原价*M.住院包装 As 原价,S.现价*M.住院包装 As 现价,S.数量/M.住院包装 As 数量"
            End Select

            gstrSQL = "SELECT DISTINCT S.ID,D.名称 AS 库房,'['||C.编码||']'||NVL(A.名称,C.名称) AS 药品,C.规格,C.产地,S.批号,S.调整金额" & strUnit & _
                " FROM (SELECT ID,库房ID,药品ID,批号,填写数量 AS 数量,成本价 AS 原价,零售价 AS 现价,零售金额 AS 调整金额" & _
                "       FROM " & _
                "       (SELECT P.ID,N.库房ID,N.药品ID,N.批号,N.填写数量,N.成本价,N.零售价,N.零售金额" & _
                "       FROM 药品收发记录 N," & _
                "           (SELECT ID,收费细目ID,执行日期,终止日期 FROM 收费价目" & _
                "           WHERE ID=[1]" & GetPriceClassString("") & ") P" & _
                "       WHERE N.药品ID=P.收费细目ID AND 单据=13" & _
                "       AND N.审核日期 BETWEEN P.执行日期 AND NVL(终止日期,SYSDATE))) S," & _
                "       部门表 D,药品规格 M,收费项目别名 A,收费项目目录 C" & _
                " WHERE S.库房ID=D.ID AND S.药品ID=M.药品ID AND M.药品ID=C.ID " & _
                " AND M.药品ID=A.收费细目ID(+) AND A.性质(+)=3 AND A.码类(+)=1" & _
                " ORDER BY '['||C.编码||']'||NVL(A.名称,C.名称),S.批号"
            Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption, Mid(strBills, 2))
            
            If rsTemp.RecordCount > 0 Then Me.hgdStore.rows = rsTemp.RecordCount + 1
            Do While Not rsTemp.EOF
                Me.hgdStore.TextMatrix(rsTemp.AbsolutePosition, 0) = rsTemp!库房
                Me.hgdStore.TextMatrix(rsTemp.AbsolutePosition, 1) = rsTemp!药品
                Me.hgdStore.TextMatrix(rsTemp.AbsolutePosition, 2) = IIf(IsNull(rsTemp!规格), "", rsTemp!规格) & IIf(IsNull(rsTemp!产地), "", "|" & rsTemp!产地)
                Me.hgdStore.TextMatrix(rsTemp.AbsolutePosition, 3) = IIf(IsNull(rsTemp!单位), "", rsTemp!单位)
                Me.hgdStore.TextMatrix(rsTemp.AbsolutePosition, 4) = IIf(IsNull(rsTemp!批号), "", rsTemp!批号)
                Me.hgdStore.TextMatrix(rsTemp.AbsolutePosition, 5) = Format(rsTemp!数量, "0.00000")
                Me.hgdStore.TextMatrix(rsTemp.AbsolutePosition, 6) = zlStr.FormatEx(rsTemp!原价, mconintPriceBit)
                Me.hgdStore.TextMatrix(rsTemp.AbsolutePosition, 7) = zlStr.FormatEx(rsTemp!现价, mconintPriceBit)
                Me.hgdStore.TextMatrix(rsTemp.AbsolutePosition, 8) = Format(rsTemp!调整金额, "0.00")
                rsTemp.MoveNext
            Loop
           
        End If
            
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    If Me.Height < 5000 Then
        Me.Height = 5000
    End If
    If Me.Width < 9720 Then
        Me.Width = 9720
    End If
    
    Me.cmdOk.Left = Me.ScaleWidth - Me.cmdOk.Width - 150
    Me.cmdCanc.Left = Me.cmdOk.Left
    Me.cmdPrint.Left = Me.cmdOk.Left
    
    
    Me.bfgPrice.Width = Me.cmdOk.Left - 150
    Me.txtRegistrar.Left = Me.bfgPrice.Left + Me.bfgPrice.Width - Me.txtRegistrar.Width
    Me.lblRegistrar.Left = txtRegistrar.Left - lblRegistrar.Width - 50
    Me.txtSummary.Width = lblRegistrar.Left - txtSummary.Left - 300
    
    Me.dtpRunDate.Left = Me.bfgPrice.Left + Me.bfgPrice.Width - Me.dtpRunDate.Width
    Me.lblRunDate.Left = dtpRunDate.Left - lblRunDate.Width - 50
    
    Me.cmdPstor.Left = Me.cmdOk.Left + Me.cmdOk.Width - Me.cmdPstor.Width
    Me.cmdCpt.Left = Me.cmdPstor.Left - 45 - Me.cmdCpt.Width
    
    Me.hgdStore.Width = Me.ScaleWidth - 30
    Me.hgdStore.Height = Me.ScaleHeight - 30 - Me.hgdStore.Top

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If BlnModify Then If MsgBox("你确定要退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Cancel = 1: Exit Sub
    SaveWinState Me
End Sub

Private Sub lvwItem_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lvwItem
        .Sorted = False
        .SortKey = ColumnHeader.Index - 2
        .SortOrder = IIf(.SortOrder = lvwDescending, lvwAscending, lvwDescending)
        .Sorted = True
    End With
End Sub

Private Sub lvwItem_DblClick()
    Dim LngmediIDThis As Long
    Dim RecCheck As New ADODB.Recordset
    
    On Error GoTo errHandle
    If Me.lvwItem.SelectedItem Is Nothing Then Exit Sub
    Set objItem = Me.lvwItem.SelectedItem
    LngmediIDThis = Mid(objItem.Key, 2)
    If LngmediIDThis = 0 Then Exit Sub
    If Me.lvwItem.Tag = conCol品名 Then
        
        '是变价药品则退出

       gstrSQL = " Select Nvl(是否变价,0) 变价 From 收费项目目录 Where 药品ID=[1]"
       Set RecCheck = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption, LngmediIDThis)
       
       '如果是时价药品，调价记录必须立即生效
       If RecCheck!变价 = 1 And blnImmediately = False Then
           blnImmediately = True
           chkImmediately.Value = 1
           chkImmediately.Enabled = False
       End If
        
        If chkImmediately.Value = 1 Then
            '判断是否有未执行的历史价格
            
            gstrSQL = " Select Count(*) Records From 收费价目 Where 变动原因=0 And 收费细目ID=[1] " & _
                    GetPriceClassString("")
            
            Set RecCheck = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption, LngmediIDThis)
            
            If Not RecCheck.EOF Then
                If Not IsNull(RecCheck!Records) Then
                    If RecCheck!Records <> 0 And chkImmediately.Value = 1 Then
                        MsgBox "该药品存在未执行价格，不能设置为立即执行！", vbInformation, gstrSysName
                        If chkImmediately.Enabled Then
                            chkImmediately.Value = 0
                        Else
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If
        
        With Me.bfgPrice
            .TextMatrix(.Row, conCol药品id) = Mid(objItem.Key, 2)
            .TextMatrix(.Row, conCol品名) = "[" & objItem.Text & "] " & objItem.SubItems(1)
            .TextMatrix(.Row, conCol规格) = objItem.SubItems(2)
            .TextMatrix(.Row, conCol产地) = objItem.SubItems(3)
            .TextMatrix(.Row, conCol单位) = objItem.SubItems(4)
            Call getMediPrice(.TextMatrix(.Row, conCol药品id))
            .CmdVisible = False
            .Col = conCol现价
        End With
    Else
        With Me.bfgPrice
            .TextMatrix(.Row, conCol收入ID) = Mid(objItem.Key, 2)
            .TextMatrix(.Row, conCol收入名称) = objItem.SubItems(1)
            .CmdVisible = False
            .Col = conCol收入名称
        End With
    End If
    Me.lvwItem.Visible = False
    bfgPrice.SetFocus
    BlnModify = True
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lvwItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> 13 Then Exit Sub
    lvwItem_DblClick
End Sub

Private Sub lvwItem_LostFocus()
    Me.lvwItem.Visible = False
End Sub

Private Sub tvwItem_DblClick()
    If Me.tvwItem.SelectedItem Is Nothing Then Exit Sub
    If Me.tvwItem.SelectedItem.Selected = False Then Exit Sub
    If Me.tvwItem.SelectedItem.Children = 0 Then
        tvwItem_NodeClick Me.tvwItem.SelectedItem
        Me.tvwItem.Visible = False
    End If
    bfgPrice.SetFocus
End Sub

Private Sub tvwItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> 13 Then Exit Sub
    tvwItem_DblClick
End Sub

Private Sub tvwItem_LostFocus()
    Me.tvwItem.Visible = False
End Sub

Private Sub tvwItem_NodeClick(ByVal Node As MSComctlLib.Node)
    If Node.Children = 0 Then
        BlnModify = True
        Me.bfgPrice.TextMatrix(Me.bfgPrice.Row, conCol收入ID) = Mid(Node.Key, 2)
        Me.bfgPrice.TextMatrix(Me.bfgPrice.Row, conCol收入名称) = Mid(Node.Text, InStr(1, Node.Text, "_") + 1)
    End If
End Sub

Private Function getMediPrice(lngMediId As Long)
    Dim bln时价 As Boolean
    Dim rs时价 As New ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = " Select Nvl(是否变价,0) 变价 From 收费项目目录 Where ID=[1]"
    Set rs时价 = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption, lngMediId)

    bln时价 = (rs时价!变价 = 1)
    If bln时价 Then Chk定价.Enabled = True

'    If bln时价 Then
'        '表示时价药品调价，取库存金额/库存数量做为其价格
'        gstrSQL = "" & _
'            " SELECT P.ID,DECODE(K.库存数量,0,P.现价,K.库存金额/NVL(K.库存数量,1)) 现价,P.执行日期,P.收入项目ID,I.名称 AS 收入名称" & _
'            " FROM 收费价目 P,收入项目 I," & _
'            "   (SELECT 药品ID,SUM(实际金额) 库存金额,SUM(实际数量) 库存数量" & _
'            "    FROM 药品库存 WHERE 性质=1 And 药品ID=" & lngMediId & _
'            "    GROUP BY 药品ID ) K" & _
'            " WHERE P.收入项目ID=I.ID AND P.收费细目ID=K.药品ID(+) AND P.收费细目ID=" & lngMediId & _
'            "       AND NVL(P.终止日期,TO_DATE('3000-01-01','YYYY-MM-DD'))=TO_DATE('3000-01-01','YYYY-MM-DD')"
'    Else
'        '非时价药品调价，取其价格记录中的价格
'        gstrSQL = "" & _
'            " SELECT P.ID,P.现价,P.执行日期,P.收入项目ID,I.名称 AS 收入名称" & _
'            " FROM 收费价目 P,收入项目 I" & _
'            " WHERE P.收入项目ID=I.ID AND P.收费细目ID=" & lngMediId & _
'            "       AND NVL(P.终止日期,TO_DATE('3000-01-01','YYYY-MM-DD'))=TO_DATE('3000-01-01','YYYY-MM-DD')"
'    End If
'    If .State = adStateOpen Then .Close
'
    
    If bln时价 Then
        '表示时价药品调价，取库存金额/库存数量做为其价格
        gstrSQL = "" & _
            " SELECT P.ID,DECODE(K.库存数量,0,P.现价,K.库存金额/NVL(K.库存数量,1)) 现价,P.执行日期,P.收入项目ID,I.名称 AS 收入名称" & _
            " FROM 收费价目 P,收入项目 I," & _
            "   (SELECT 药品ID,SUM(实际金额) 库存金额,SUM(实际数量) 库存数量" & _
            "    FROM 药品库存 WHERE 性质=1 And 药品ID=[1] " & _
            "    GROUP BY 药品ID ) K" & _
            " WHERE P.收入项目ID=I.ID AND P.收费细目ID=K.药品ID(+) AND P.收费细目ID=[1] " & _
            GetPriceClassString("P") & _
            "       AND NVL(P.终止日期,TO_DATE('3000-01-01','YYYY-MM-DD'))=TO_DATE('3000-01-01','YYYY-MM-DD')"
    Else
        '非时价药品调价，取其价格记录中的价格
        gstrSQL = "" & _
            " SELECT P.ID,P.现价,P.执行日期,P.收入项目ID,I.名称 AS 收入名称" & _
            " FROM 收费价目 P,收入项目 I" & _
            " WHERE P.收入项目ID=I.ID AND P.收费细目ID=[1] " & _
            "       AND NVL(P.终止日期,TO_DATE('3000-01-01','YYYY-MM-DD'))=TO_DATE('3000-01-01','YYYY-MM-DD')" & _
            GetPriceClassString("P")
    End If
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption, lngMediId)
    
    With rsTemp
        If .RecordCount > 0 Then
            Me.bfgPrice.RowData(Me.bfgPrice.Row) = !Id
            Me.bfgPrice.TextMatrix(Me.bfgPrice.Row, conCol原价) = zlStr.FormatEx(!现价, mconintPriceBit)
            Me.bfgPrice.TextMatrix(Me.bfgPrice.Row, conCol现价) = zlStr.FormatEx(!现价, mconintPriceBit)
            Me.bfgPrice.TextMatrix(Me.bfgPrice.Row, conCol收入ID) = !收入项目ID
            Me.bfgPrice.TextMatrix(Me.bfgPrice.Row, conColOld收入ID) = !收入项目ID
            Me.bfgPrice.TextMatrix(Me.bfgPrice.Row, conCol收入名称) = !收入名称
            If Me.dtpRunDate.MinDate <= DateAdd("d", 1, !执行日期) Then
                Me.dtpRunDate.MinDate = DateAdd("d", 1, CDate(Format(!执行日期, "yyyy-MM-dd 00:00:00")))
            End If
            If Me.dtpRunDate.Value <= Me.dtpRunDate.MinDate Then
                Me.dtpRunDate.Value = Me.dtpRunDate.MinDate
            End If
        Else
            Me.bfgPrice.RowData(Me.bfgPrice.Row) = -1
            Me.bfgPrice.TextMatrix(Me.bfgPrice.Row, conCol原价) = Format(0, "0.0000000")
            Me.bfgPrice.TextMatrix(Me.bfgPrice.Row, conCol现价) = Format(0, "0.0000000")
            If Me.bfgPrice.Row > 1 Then
                Me.bfgPrice.TextMatrix(Me.bfgPrice.Row, conCol收入ID) = Me.bfgPrice.TextMatrix(Me.bfgPrice.Row - 1, conCol收入ID)
                Me.bfgPrice.TextMatrix(Me.bfgPrice.Row, conColOld收入ID) = Me.bfgPrice.TextMatrix(Me.bfgPrice.Row - 1, conCol收入ID)
                Me.bfgPrice.TextMatrix(Me.bfgPrice.Row, conCol收入名称) = Me.bfgPrice.TextMatrix(Me.bfgPrice.Row - 1, conCol收入名称)
            Else
                For Each objNode In Me.tvwItem.Nodes
                    If objNode.Children = 0 Then
                        Me.bfgPrice.TextMatrix(Me.bfgPrice.Row, conCol收入ID) = Mid(objNode.Key, 2)
                        Me.bfgPrice.TextMatrix(Me.bfgPrice.Row, conColOld收入ID) = Mid(objNode.Key, 2)
                        Me.bfgPrice.TextMatrix(Me.bfgPrice.Row, conCol收入名称) = objNode.Text
                    End If
                Next
            End If
        End If
    End With
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub txtSummary_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Me.dtpRunDate.Enabled Then Me.dtpRunDate.SetFocus
End Sub

Private Function CheckPrice() As Boolean
    Dim IntCheck As Integer
    Dim RecCheck As New ADODB.Recordset
    '检测各执行价格是否正确
    '以及收入项目相同的情况下现价是否与原价相同
    CheckPrice = False
    With bfgPrice
        For IntCheck = 1 To .rows - 1
            If Val(.TextMatrix(IntCheck, conCol药品id)) <> 0 Then
                If Not IsNumeric(Trim(.TextMatrix(IntCheck, conCol现价))) Then
                    MsgBox "第" & IntCheck & "行的药品现价中含有非法字符！", vbInformation, gstrSysName
                    Exit Function
                End If
                If Val(.TextMatrix(IntCheck, conCol现价)) = 0 Then
                    MsgBox "第" & IntCheck & "行的药品现价不能为空！", vbInformation, gstrSysName
                    Exit Function
                End If
                If Val(.TextMatrix(IntCheck, conColOld收入ID)) = Val(.TextMatrix(IntCheck, conCol收入ID)) Then
                    If Val(.TextMatrix(IntCheck, conCol现价)) = Val(.TextMatrix(IntCheck, conCol原价)) Then
                        MsgBox "第" & IntCheck & "行的药品现价与原价相同，不能执行调价！"
                        Exit Function
                    End If
                End If
            End If
            
        Next
    End With
    CheckPrice = True
End Function
