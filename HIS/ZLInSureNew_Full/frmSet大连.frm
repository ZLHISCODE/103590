VERSION 5.00
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.3#0"; "ZL9BillEdit.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSet大连 
   AutoRedraw      =   -1  'True
   Caption         =   " "
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7650
   Icon            =   "frmSet大连.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5460
   ScaleWidth      =   7650
   StartUpPosition =   1  '所有者中心
   Begin TabDlg.SSTab tabSel 
      Height          =   4155
      Left            =   60
      TabIndex        =   3
      Top             =   735
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   7329
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmSet大连.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "mshBill"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "可选参数"
      TabPicture(1)   =   "frmSet大连.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "chkQ"
      Tab(1).Control(1)=   "chkBig"
      Tab(1).Control(2)=   "chk实时(0)"
      Tab(1).Control(3)=   "chk实时(1)"
      Tab(1).Control(4)=   "chkSingleD"
      Tab(1).Control(5)=   "chkSingleK"
      Tab(1).Control(6)=   "chk开发区"
      Tab(1).Control(7)=   "txtEdit"
      Tab(1).Control(8)=   "lblEdit(3)"
      Tab(1).ControlCount=   9
      Begin VB.CheckBox chkQ 
         Caption         =   "门诊企离使用住院比例(&4)"
         Height          =   195
         Left            =   -72150
         TabIndex        =   13
         Top             =   1965
         Width           =   2415
      End
      Begin VB.CheckBox chkBig 
         Caption         =   "门诊大病使用住院比例(&3)"
         Height          =   195
         Left            =   -74640
         TabIndex        =   12
         Top             =   1965
         Width           =   2415
      End
      Begin VB.CheckBox chk实时 
         Caption         =   "门诊明细时实上传(&M)"
         Height          =   285
         Index           =   0
         Left            =   -74640
         TabIndex        =   11
         Top             =   1050
         Width           =   2085
      End
      Begin VB.CheckBox chk实时 
         Caption         =   "住院明细时实上传(&Z)"
         Height          =   285
         Index           =   1
         Left            =   -72150
         TabIndex        =   10
         Top             =   1050
         Width           =   2085
      End
      Begin VB.CheckBox chkSingleD 
         Caption         =   "大连市出院单病种提示(&1)"
         Height          =   240
         Left            =   -74640
         TabIndex        =   9
         Top             =   1530
         Width           =   2475
      End
      Begin VB.CheckBox chkSingleK 
         Caption         =   "开发区出院单病种提示(&2)"
         Height          =   240
         Left            =   -72150
         TabIndex        =   8
         Top             =   1530
         Width           =   2475
      End
      Begin VB.CheckBox chk开发区 
         Caption         =   "开发区(&K)"
         Height          =   255
         Left            =   -72150
         TabIndex        =   6
         Top             =   645
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   -73590
         MaxLength       =   2
         TabIndex        =   5
         Text            =   "1"
         Top             =   600
         Width           =   360
      End
      Begin ZL9BillEdit.BillEdit mshBill 
         Height          =   3645
         Left            =   60
         TabIndex        =   4
         Top             =   405
         Width           =   7425
         _ExtentX        =   13097
         _ExtentY        =   6429
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
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "当前串口(&D)     号串口"
         Height          =   180
         Index           =   3
         Left            =   -74640
         TabIndex        =   7
         Top             =   660
         Width           =   1980
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6435
      TabIndex        =   2
      Top             =   5025
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5175
      TabIndex        =   1
      Top             =   5025
      Width           =   1100
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   60
      Picture         =   "frmSet大连.frx":0044
      Top             =   150
      Width           =   480
   End
   Begin VB.Label lbl 
      Caption         =   "设置设备的串口号及设置窗口是否默认为开发区,并将其收费类型与相关的医保项目相对应"
      Height          =   315
      Left            =   540
      TabIndex        =   0
      Top             =   285
      Width           =   7125
   End
End
Attribute VB_Name = "frmSet大连"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnReturn As Boolean
Private mlng医保中心 As Long
Private mlng险类 As Long
Private Enum mColHead
    收费类别 = 0
    保费项目
    分类项目
End Enum
Private Sub chk开发区_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbTab
    End If
End Sub


Private Sub chk实时_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdCancel_Click()
    mblnReturn = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim lngRow As Long
    
    If Trim(txtEdit) = "" Then Exit Sub
    SaveRegInFor g公共模块, "操作", "端口号", Me.txtEdit
    SaveRegInFor g公共模块, "操作", "开发区", Me.chk开发区.Value
    If Val(txtEdit) = 0 Then
        gintComPort_大连 = 1
    Else
        gintComPort_大连 = Val(txtEdit)
    End If
    gblnKFQCom_大连 = IIf(chk开发区.Value = 1, True, False)
    gintComPort = txtEdit.Text
        
    '删除已经数据
    gcnOracle.BeginTrans
    On Error GoTo ErrHand
    
    gstrSQL = "zl_保险参数_Delete(" & mlng险类 & ",NUll)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    With mshBill
        For lngRow = 1 To .Rows - 1
            If .TextMatrix(lngRow, mColHead.收费类别) <> "" Then
                '新增参数数据
                gstrSQL = "zl_保险参数_Insert(" & mlng险类 & ",null,'" & .TextMatrix(lngRow, mColHead.收费类别) & "' ,'" & .TextMatrix(lngRow, mColHead.保费项目) & ";" & .TextMatrix(lngRow, mColHead.分类项目) & "'," & lngRow + 2 & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            End If
        Next
    End With
    '保存
    
    gstrSQL = "zl_保险参数_Update(" & mlng险类 & ",NULL,'门诊明细时实上传' ,'" & IIf(chk实时(0).Value = 1, "1", "0") & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)

    gstrSQL = "zl_保险参数_Update(" & mlng险类 & ",NULL,'住院明细时实上传' ,'" & IIf(chk实时(1).Value = 1, "1", "0") & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)

    If mlng险类 = 82 Then
        gstrSQL = "zl_保险参数_Update(" & mlng险类 & ",NULL,'单病种出院提示' ,'" & IIf(chkSingleD.Value = 1, "1", "0") & "')"
    Else
        gstrSQL = "zl_保险参数_Update(" & mlng险类 & ",NULL,'单病种出院提示' ,'" & IIf(chkSingleK.Value = 1, "1", "0") & "')"
    End If
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)

    gstrSQL = "zl_保险参数_Update(" & mlng险类 & ",NULL,'门诊大病使用住院比例' ,'" & IIf(chkBig.Value = 1, "1", "0") & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    If mlng险类 = 82 Then
        gstrSQL = "zl_保险参数_Update(" & mlng险类 & ",NULL,'门诊企离使用住院比例' ,'" & IIf(chkQ.Value = 1, "1", "0") & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    End If
    
    gcnOracle.CommitTrans
    mblnReturn = True
    Unload Me
    Exit Sub
ErrHand:
    gcnOracle.RollbackTrans
    Resume
    Call ErrCenter
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    Dim strReg As String
    mblnReturn = False
    
    Call GetRegInFor(g公共模块, "操作", "端口号", strReg)
    If Val(strReg) = 0 Then
        txtEdit.Text = 1
    Else
        txtEdit.Text = Val(strReg)
    End If
    
    Call GetRegInFor(g公共模块, "操作", "开发区", strReg)
    If Val(strReg) = 1 Then
        Me.chk开发区.Value = 1
    Else
        Me.chk开发区.Value = 0
    End If
    
    Call GetRegInFor(g公共模块, "操作", "大连市单病种", strReg)
    If Val(strReg) = 1 Then
        Me.chkSingleD.Value = 1
    Else
        Me.chkSingleD.Value = 0
    End If
    
    Call GetRegInFor(g公共模块, "操作", "开发区单病种", strReg)
    If Val(strReg) = 1 Then
        Me.chkSingleK.Value = 1
    Else
        Me.chkSingleK.Value = 0
    End If
    
    Call GetRegInFor(g公共模块, "操作", "门诊大病使用住院比例", strReg)
    If Val(strReg) = 1 Then
        Me.chkBig.Value = 1
    Else
        Me.chkBig.Value = 0
    End If
    
    If mlng险类 = 82 Then
        Call GetRegInFor(g公共模块, "操作", "门诊企离使用住院比例", strReg)
        If Val(strReg) = 1 Then
            Me.chkQ.Value = 1
        Else
            Me.chkQ.Value = 0
        End If
    End If
    RestoreWinState Me, App.ProductName
    
    '初始数据
    Call iniData
End Sub

Public Function ShowMe(ByVal lng险类 As Long, ByVal lng医保中心 As Long) As Boolean
    mlng医保中心 = lng医保中心
    mlng险类 = lng险类
    
    Me.Show 1
    ShowMe = mblnReturn
End Function

Private Sub Form_Resize()
    On Error Resume Next
    
    With cmdCancel
        .Top = ScaleHeight - .Height - 100
        .Left = ScaleWidth - .Width - 50
    End With
    With cmdOK
        .Top = cmdCancel.Top
        .Left = cmdCancel.Left - 50 - .Width
    End With
    
    With tabSel
        .Width = ScaleWidth - 50
        .Height = cmdOK.Top - .Top - 100
    End With
    
    With mshBill
        .Top = tabSel.Top - 300
        .Left = tabSel.Left + 50
        .Height = cmdOK.Top - 1400
        .Width = tabSel.Width - 200
    End With
    mshBill.ZOrder
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub mshBill_EnterCell(Row As Long, Col As Long)
    With mshBill
        Select Case Col
            Case mColHead.保费项目
                mshBill.Clear
                mshBill.AddItem "诊察费"
                mshBill.AddItem "草药费"
                mshBill.AddItem "成药费"
                mshBill.AddItem "西药费"
                mshBill.AddItem "检查费"
                mshBill.AddItem "大检费"
                mshBill.AddItem "治疗费"
                mshBill.AddItem "特殊治疗费"
                mshBill.AddItem "血费"
                
            Case mColHead.分类项目
                mshBill.Clear
                mshBill.AddItem "A中草药费"
                mshBill.AddItem "B中成药费"
                mshBill.AddItem "C西药费"
                mshBill.AddItem "D检查费"
                mshBill.AddItem "E输氧费"
                mshBill.AddItem "F放射费"
                mshBill.AddItem "G手术费"
                mshBill.AddItem "H化验费"
                mshBill.AddItem "I诊疗费"
                mshBill.AddItem "J麻醉费"
                mshBill.AddItem "K床位费"
                mshBill.AddItem "L护理费"
                mshBill.AddItem "X输血费"
                mshBill.AddItem "M其它费用"
        End Select
    End With
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtEdit, KeyAscii, m数字式
End Sub

Private Function iniData() As Boolean
    '初始数据
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim lngRow As Long
    Dim strTmp As String
    
    '设置页头
    Err = 0
    On Error Resume Next
    strSQL = "Select * from 保险中心目录 where 险类=" & mlng险类
    zlDatabase.OpenRecordset rsTmp, strSQL, Me.Caption
    If rsTmp.EOF Then
        tabSel.Caption = "无"
    Else
        tabSel.Caption = Nvl(rsTmp!名称)
    End If
    rsTmp.Close
  
    If mlng险类 = type_大连开发区 Then
        Me.chk开发区.Value = 1
    Else
        Me.chk开发区.Value = 0
    End If
    
    '设置表列头
    Call initGrid
    strSQL = "" & _
        "   Select A.类别,b.参数值 From 收费类别 a,(Select * From 保险参数 where 险类=" & mlng险类 & ") b " & _
        "   Where A.类别=b.参数名(+) " & _
        "   order by A.编码 "
    zlDatabase.OpenRecordset rsTmp, strSQL, Me.Caption
    With mshBill
        .ClearBill
        If rsTmp.RecordCount = 0 Then
            .Rows = 2
        Else
            .Rows = rsTmp.RecordCount + 1
        End If
        lngRow = 1
        Do While Not rsTmp.EOF
            .TextMatrix(lngRow, mColHead.收费类别) = Nvl(rsTmp!类别)
            strTmp = Nvl(rsTmp!参数值)
            If InStr(1, strTmp, ";") <> 0 Then
                .TextMatrix(lngRow, mColHead.保费项目) = Split(strTmp, ";")(0)
                .TextMatrix(lngRow, mColHead.分类项目) = Split(strTmp, ";")(1)
            End If
            lngRow = lngRow + 1
            rsTmp.MoveNext
        Loop
        
        strSQL = "Select 参数名,参数值 From 保险参数 " & _
                " Where 参数名 in('门诊明细时实上传','住院明细时实上传','医嘱明细时实上传'," & _
                "'单病种出院提示','门诊大病使用住院比例','门诊企离使用住院比例') and 险类=" & mlng险类
        zlDatabase.OpenRecordset rsTmp, strSQL, Me.Caption
        chk实时(0).Value = 1
        chk实时(1).Value = 1
        chkSingleD.Value = 1
        chkSingleK.Value = 1
        chkBig.Value = 1
        chkQ.Value = 0
        chkQ.Visible = False
        Do While Not rsTmp.EOF
            Select Case Nvl(rsTmp!参数名)
            Case "门诊明细时实上传"
                chk实时(0).Value = IIf(Val(Nvl(rsTmp!参数值)) = 1, 1, 0)
            Case "住院明细时实上传"
                chk实时(1).Value = IIf(Val(Nvl(rsTmp!参数值)) = 1, 1, 0)
            Case "单病种出院提示"
                If mlng险类 = 82 Then
                    chkSingleD.Value = IIf(Val(Nvl(rsTmp!参数值)) = 1, 1, 0)
                Else
                    chkSingleK.Value = IIf(Val(Nvl(rsTmp!参数值)) = 1, 1, 0)
                End If
            Case "门诊大病使用住院比例"
                chkBig.Value = IIf(Val(Nvl(rsTmp!参数值)) = 1, 1, 0)
            Case "门诊企离使用住院比例"
                If mlng险类 = 82 Then
                    chkQ.Visible = True
                    chkQ.Value = IIf(Val(Nvl(rsTmp!参数值)) = 1, 1, 0)
                Else
                    chkQ.Visible = False
                End If
            End Select
            rsTmp.MoveNext
        Loop
    End With
End Function
Private Sub initGrid()
    With mshBill
        .Active = True
        .Cols = 3
        
        .msfObj.FixedCols = 1
        .AllowAddRow = False
        
        .TextMatrix(0, mColHead.收费类别) = "收费类别"
        .TextMatrix(0, mColHead.保费项目) = "保费项目"
        .TextMatrix(0, mColHead.分类项目) = "分类项目"
        
        
        .ColWidth(mColHead.收费类别) = 1500
        .ColWidth(mColHead.保费项目) = 2000
        .ColWidth(mColHead.分类项目) = 2000
        
        '-1：表示该列可以选择，是布尔型［"√"，" "］
        ' 0：表示该列可以选择，但不能修改
        ' 1：表示该列可以输入，外部显示为按钮选择
        ' 2：表示该列是日期列，外部显示为按钮选择，弹出是日期选择框
        ' 3：表示该列是选择列，外部显示为下拉框选择
        '4:  表示该列为单纯的文本框供用户输入
        '5:  表示该列不允许选择

        .ColData(mColHead.收费类别) = 5
        .ColData(mColHead.保费项目) = 3
        .ColData(mColHead.分类项目) = 3
        
        .ColAlignment(mColHead.收费类别) = flexAlignLeftCenter
        .ColAlignment(mColHead.保费项目) = flexAlignLeftCenter
        .ColAlignment(mColHead.分类项目) = flexAlignLeftCenter
        .PrimaryCol = mColHead.保费项目
        .LocateCol = mColHead.保费项目
    End With
End Sub



