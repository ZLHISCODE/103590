VERSION 5.00
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form Frm按单号退料 
   Caption         =   "按单据进行退料"
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9690
   Icon            =   "Frm按单号退料.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   9690
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdAllCls 
      Caption         =   "全清(&L)"
      Height          =   350
      Left            =   1320
      TabIndex        =   17
      Top             =   4980
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CommandButton cmdAllSel 
      Caption         =   "全退(&A)"
      Height          =   350
      Left            =   12
      TabIndex        =   16
      Top             =   4980
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   8472
      TabIndex        =   13
      Top             =   5088
      Width           =   1100
   End
   Begin VB.Frame fraTop 
      Height          =   672
      Left            =   -12
      TabIndex        =   14
      Top             =   24
      Width           =   9660
      Begin VB.ComboBox cbo收费单据 
         Height          =   300
         Left            =   852
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   288
         Width           =   1116
      End
      Begin VB.TextBox TxtNo 
         Height          =   300
         Left            =   2700
         TabIndex        =   3
         Top             =   276
         Width           =   1668
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "单据类型"
         Height          =   180
         Index           =   3
         Left            =   96
         TabIndex        =   0
         Top             =   336
         Width           =   720
      End
      Begin VB.Label LblNote 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "未输入任何处方"
         ForeColor       =   &H80000002&
         Height          =   180
         Left            =   5412
         TabIndex        =   4
         Top             =   324
         Width           =   4116
      End
      Begin VB.Label LblNo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "单据号"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2136
         TabIndex        =   2
         Top             =   336
         Width           =   540
      End
   End
   Begin ZL9BillEdit.BillEdit Bill 
      Height          =   3540
      Left            =   -12
      TabIndex        =   5
      Top             =   744
      Width           =   9660
      _ExtentX        =   17039
      _ExtentY        =   6244
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      TxtCheck        =   -1  'True
      TxtCheck        =   -1  'True
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
      Active          =   -1  'True
      Cols            =   2
      RowHeight0      =   360
      RowHeightMin    =   360
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
   Begin VB.Frame fraBillDown 
      Height          =   684
      Left            =   0
      TabIndex        =   15
      Top             =   4236
      Width           =   9624
      Begin VB.TextBox txtDown 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   300
         Index           =   2
         Left            =   7620
         TabIndex        =   11
         Text            =   "2009-09-09 23:59:59"
         Top             =   204
         Width           =   1944
      End
      Begin VB.TextBox txtDown 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   5124
         TabIndex        =   9
         Text            =   "测试人员"
         Top             =   204
         Width           =   1308
      End
      Begin VB.TextBox txtDown 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   888
         TabIndex        =   7
         Text            =   "[01]一病区"
         Top             =   204
         Width           =   3336
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "时间"
         Height          =   180
         Index           =   2
         Left            =   7140
         TabIndex        =   10
         Top             =   264
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "开单"
         Height          =   180
         Index           =   1
         Left            =   4668
         TabIndex        =   8
         Top             =   264
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "开单科室"
         Height          =   180
         Index           =   0
         Left            =   132
         TabIndex        =   6
         Top             =   264
         Width           =   720
      End
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "退料(&O)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   7188
      TabIndex        =   12
      Top             =   5088
      Width           =   1100
   End
End
Attribute VB_Name = "Frm按单号退料"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String
Private mblnFirst As Boolean
Private mlng发料部门 As Long
Private mfrmMain As Form
Private mblnSucces As Boolean
Private mintUnit As Integer
Private Const mstrCaption As String = "按单据进行退料"
'----------------------------------------------------------------------------------------------------------
'刘兴宏:增加小数位数的格式串
'修改:2007/03/06
Private mFMT As g_FmtString
Private mOraFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------

Private mblnChange As Boolean
Private Enum mCol
    c_类别 = 0
    c_执行科室
    c_姓名
    c_年龄
    c_床号
    c_项目
    c_规格
    c_批号
    c_批次
    c_换算系数
    c_单位
    c_付数
    c_数次
    c_原始数量
    c_已退数
    c_准退数
    c_退料数
    c_单价
    c_金额
End Enum
Private mintPreBillType  As Integer
Private Const mCols = 19
Private Const mlngModule = 1723

Private mobjPlugIn As Object             '外挂接口对象

Public Property Get In_PlugIn() As Object
    Set In_PlugIn = mobjPlugIn
End Property
Public Property Set In_PlugIn(ByVal objVal As Object)
    Set mobjPlugIn = objVal
End Property
Public Function ShowCard(ByVal frmMain As Form, ByVal lng发料部门ID As Long, ByVal strPrivs As String) As Boolean
    '------------------------------------------------------------------------------------------------------
    '功能:对指定的处方进行部分退料的入口
    '入参:mfrmMain-主窗口
    '     lng发料部门ID-发料部门ID
    '     strPrivs-权限串
    '出参:
    '返回:退料成功,返回true,否则返回false
    '修改人:刘兴宏
    '修改时间:2007/3/1
    '------------------------------------------------------------------------------------------------------
    Set mfrmMain = frmMain
    mstrPrivs = strPrivs
    mlng发料部门 = lng发料部门ID
    Me.Show 1, frmMain
    ShowCard = mblnSucces
End Function

 

Private Sub Bill_BeforeDeleteRow(Row As Long, Cancel As Boolean)
        Cancel = True
End Sub

Private Sub Bill_EditChange(curText As String)
    mblnChange = True
End Sub

Private Sub Bill_EnterCell(Row As Long, Col As Long)
    With Bill
        Select Case .Col
            Case mCol.c_退料数
                .TxtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"
        End Select
    End With
End Sub

Private Sub Bill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    If KeyCode <> vbKeyReturn Then Exit Sub
    With Bill
        .Text = UCase(Trim(.Text))
        strKey = UCase(Trim(.Text))
        Select Case .Col
            
            Case mCol.c_退料数
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "退料数量必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" Then
                    If Val(strKey) = 0 Then
                        MsgBox "退料数量必须大于零,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Val(Format(Val(strKey), mFMT.FM_数量)) > Val(Format(Val(.TextMatrix(.Row, c_准退数)), mFMT.FM_数量)) Then
                        MsgBox "退料数量不能大于准退数量,请重输！", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If

                    If Abs(Val(strKey)) >= 10 ^ 11 - 1 Then
                        MsgBox "退料数量必须在(-" & (10 ^ 11 - 1) & " 至 " & (10 ^ 11 - 1) & ") 之间", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    strKey = Format(strKey, mFMT.FM_数量)
                    .Text = strKey
                    .TextMatrix(.Row, .Col) = .Text
                End If
        End Select
    End With
End Sub

Private Sub cbo收费单据_Click()
    If mintPreBillType = cbo收费单据.ItemData(cbo收费单据.ListIndex) Then Exit Sub
    mintPreBillType = cbo收费单据.ItemData(cbo收费单据.ListIndex)
    With Bill
        .Rows = 2
        .ClearBill
        txtDown(0).Text = ""
        txtDown(1).Text = ""
        txtDown(2).Text = ""
    End With
End Sub

Private Sub cbo收费单据_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub cmdAllCls_Click()
    Dim intRow As Integer
    With Bill
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, mCol.c_项目) <> "" Then
                .TextMatrix(intRow, mCol.c_退料数) = ""
            End If
        Next
    End With
End Sub

Private Sub cmdAllSel_Click()
    Dim intRow As Integer
    With Bill
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, mCol.c_项目) <> "" Then
                .TextMatrix(intRow, mCol.c_退料数) = .TextMatrix(intRow, mCol.c_准退数)
            End If
        Next
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    '退料
    Dim strDate As String
    
    Dim bln退料单 As Boolean
    If ISValied = False Then Exit Sub
    strDate = Format(Sys.Currentdate, "yyyy-mm-dd HH:MM:SS")
    If Save退料(strDate) = False Then Exit Sub
    bln退料单 = InStr(1, mstrPrivs, "退料通知单") <> 0
    
    If bln退料单 Then
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1723_2", Me, "退料时间=" & strDate, "单位=" & mintUnit + 1, 2)
    End If
    mblnSucces = True
    mblnChange = False
    MsgBox "退料成功", vbInformation + vbDefaultButton1, gstrSysName
    Bill.ClearBill
End Sub
Private Function Save退料(ByVal strDate As String) As Boolean
    '检查退料的相关条件
    Dim strNo As String
    Dim lngId As Long
    Dim lngRow As Long
    Dim 材料ID As Long
    Dim int自动销帐 As Integer
    Dim strReg As String
    Dim cllTemp As New Collection
    Dim dbl数量 As Double
    Dim strReturnInfo As String
    Dim strReserve As String
    Dim rsTemp As New ADODB.Recordset
    Dim bln备货卫材 As Boolean
    Dim int自动销帐_原始值 As Integer
    
    int自动销帐_原始值 = IIf(Val(zlDatabase.GetPara("自动销帐", glngSys, mlngModule)) = 1, 1, 0)
    
    Save退料 = False
    err = 0
    
    With Bill
        For lngRow = 1 To .Rows - 1
                
                 
                If Trim(.TextMatrix(lngRow, mCol.c_项目)) <> "" And Val(.TextMatrix(lngRow, mCol.c_退料数)) <> 0 Then
                    int自动销帐 = int自动销帐_原始值
                    
                    If int自动销帐 <> 1 Then
                        '判断是否备货卫材
                        gstrSQL = " Select 1 From 药品收发记录 Where 单据 = 21 And 审核日期 Is Not Null And 费用id = (select 费用id from 药品收发记录 where id=[1]) And Rownum < 2 "
                        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "判断是否备货卫材", Val(.RowData(lngRow)))
                        bln备货卫材 = Not rsTemp.EOF
                    
                        '如果是高值卫材也进行自动销帐
                        If bln备货卫材 Then int自动销帐 = 1
                    End If
                 
                 '   过程参数:ID_IN,审核人_IN,审核日期_IN,批号_IN,效期_IN,产地_IN,退料数量_IN,自动销帐_IN(1-自动销帐,0-不自动销帐)
                   gstrSQL = "zl_材料收发记录_部门退料("
                   gstrSQL = gstrSQL & .RowData(lngRow) & ","
                   gstrSQL = gstrSQL & "'" & gstrUserName & "',"
                   gstrSQL = gstrSQL & "to_date('" & strDate & "','yyyy-mm-dd HH24:mi:ss'),"
                   gstrSQL = gstrSQL & "'" & Replace(.TextMatrix(lngRow, mCol.c_批号), "(" & .TextMatrix(lngRow, mCol.c_批次) & ")", "") & "',"
                   gstrSQL = gstrSQL & "NULL" & ","
                   gstrSQL = gstrSQL & "NULL" & ","
                   dbl数量 = 0
                   If Val(.TextMatrix(lngRow, mCol.c_准退数)) = Val(.TextMatrix(lngRow, mCol.c_退料数)) Then
                            dbl数量 = Val(.TextMatrix(lngRow, mCol.c_原始数量))
                   Else
                    If mintUnit = 0 Then
                         dbl数量 = Val(.TextMatrix(lngRow, mCol.c_退料数))
                    Else
                         dbl数量 = Round(Val(.TextMatrix(lngRow, mCol.c_退料数)) * Val(.TextMatrix(lngRow, mCol.c_换算系数)), g_小数位数.obj_散装小数.数量小数)
                    End If
                   End If
                   gstrSQL = gstrSQL & dbl数量
                   gstrSQL = gstrSQL & "," & int自动销帐 & ")"
                   Call AddArray(cllTemp, gstrSQL)
                   
                   strReturnInfo = IIf(strReturnInfo = "", "", strReturnInfo & "|") & NVL(.RowData(lngRow)) & "," & dbl数量
                End If
        Next
    End With
    
    On Error GoTo ErrHand:
    
    Call ExecuteProcedureArrAy(cllTemp, mstrCaption)
    
    '调用退药后的外挂接口
    If Not mobjPlugIn Is Nothing And strReturnInfo <> "" Then
        mobjPlugIn.DrugReturnByID mlng发料部门, strReturnInfo, CDate(strDate), strReserve
    End If
    
    Save退料 = True
    Exit Function
ErrHand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ISValied() As Boolean
    '------------------------------------------------------------------------------------------------------
    '功能:检查数据的合法性
    '入参:
    '出参:
    '返回:数据合法,返回ture,否则返回False
    '修改人:刘兴宏
    '修改时间:2007/3/2
    '------------------------------------------------------------------------------------------------------
    Dim intRow As Integer
    Dim blnHave As Boolean
    Dim dblTemp As Double
    blnHave = False
    With Bill
        For intRow = 1 To .Rows - 1
            If Trim(.TextMatrix(intRow, mCol.c_项目)) <> "" Then
                dblTemp = Val(.TextMatrix(intRow, mCol.c_退料数))
                If dblTemp > Val(.TextMatrix(intRow, mCol.c_准退数)) Then
                    ShowMsgBox "退料数量(￥:" & Format(dblTemp, mFMT.FM_数量) & ") 大于了准退数量(￥:" & Format(Val(.TextMatrix(intRow, mCol.c_准退数)), mFMT.FM_数量) & "),请检查!"
                    .Row = intRow
                    .Col = c_退料数
                    .SetFocus
                    Exit Function
                End If
                If Abs(dblTemp) >= 10 ^ 11 - 1 Then
                    MsgBox "退料数量必须在(-" & (10 ^ 11 - 1) & " 至 " & (10 ^ 11 - 1) & ") 之间", vbInformation + vbOKOnly, gstrSysName
                    .Row = intRow
                    .Col = c_退料数
                    .TxtSetFocus
                    Exit Function
                End If
                If dblTemp <> 0 Then
                    blnHave = True
                End If
            End If
        Next
    End With
    If blnHave = False Then
        ShowMsgBox "你还未输入退料数量,请检查!"
        Bill.Row = 1
        Bill.Col = c_退料数
        Bill.SetFocus
        Exit Function
    End If
    ISValied = True
End Function
Private Sub Form_Load()
    Dim strReg As String
    mblnFirst = True
    
    strReg = Val(zlDatabase.GetPara("卫材单位", glngSys, mlngModule, "0"))
    mintUnit = Val(strReg)
  
    '刘兴宏:增加小数格式化串
    With mFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价)
        .FM_金额 = GetFmtString(mintUnit, g_金额)
        .FM_零售价 = GetFmtString(mintUnit, g_售价)
        .FM_数量 = GetFmtString(mintUnit, g_数量)
    End With
  
    With mOraFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价, True)
        .FM_金额 = GetFmtString(mintUnit, g_金额, True)
        .FM_零售价 = GetFmtString(mintUnit, g_售价, True)
        .FM_数量 = GetFmtString(mintUnit, g_数量, True)
    End With
    Call initGrid

    RestoreWinState Me, App.ProductName, mstrCaption
    
End Sub
Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    mintPreBillType = 0
    
    '初始
    Call InitData
    cmdOK.Enabled = True
    mblnChange = False
    cmdAllCls.Visible = True
    cmdAllSel.Visible = True
    txtDown(0).Text = ""
    txtDown(1).Text = ""
    txtDown(2).Text = ""
    
End Sub
Private Sub InitData()
    '------------------------------------------------------------------------------------------------------
    '功能:初始一些必要条件数据
    '入参:
    '出参:
    '返回:
    '修改人:刘兴宏
    '修改时间:2007/3/21
    '------------------------------------------------------------------------------------------------------
    Dim strReg As String
    
    strReg = Trim(zlDatabase.GetPara("最后退料单据类型", glngSys, mlngModule, "25", Array(lbl(3), cbo收费单据), zlStr.IsHavePrivs(mstrPrivs, "参数设置")))
    If Val(strReg) = 0 Then strReg = "25"
    With cbo收费单据
        .Clear
        .AddItem "0-收费"
        .ItemData(.NewIndex) = 24
        If Val(strReg) = 24 Then .ListIndex = .NewIndex
        .AddItem "1-记帐"
        .ItemData(.NewIndex) = 25
        If Val(strReg) = 25 Then .ListIndex = .NewIndex
        .AddItem "2-记帐表"
        .ItemData(.NewIndex) = 26
        If Val(strReg) = 26 Then .ListIndex = .NewIndex
        If .ListIndex < 0 Then .ListIndex = 0
    End With

End Sub

Private Sub initGrid()
    '------------------------------------------------------------------------------------------------------
    '功能:初始网格控件
    '入参:
    '出参:
    '返回:
    '修改人:刘兴宏
    '修改时间:2007/3/1
    '------------------------------------------------------------------------------------------------------
     With Bill
        .Cols = mCols
      '  .MsfObj.FixedCols = 1
        .AllowAddRow = False
        .TextMatrix(0, c_类别) = "类别"
        .TextMatrix(0, c_项目) = "项目"
        .TextMatrix(0, c_规格) = "规格"
        .TextMatrix(0, c_批号) = "批号"
        .TextMatrix(0, c_批次) = "批次"
        .TextMatrix(0, c_换算系数) = "换算系数"
        .TextMatrix(0, c_单位) = "单位"
        .TextMatrix(0, c_付数) = "付数"
        
        .TextMatrix(0, c_数次) = "数次"
        .TextMatrix(0, c_原始数量) = "原始数量"
        .TextMatrix(0, c_已退数) = "已退数"
        .TextMatrix(0, c_准退数) = "准退数"
        .TextMatrix(0, c_退料数) = "退料数"
        .TextMatrix(0, c_单价) = "单价"
        .TextMatrix(0, c_金额) = "金额"
        .TextMatrix(0, c_执行科室) = "执行科室"
        .TextMatrix(0, c_姓名) = "姓名"
        .TextMatrix(0, c_年龄) = "年龄"
        .TextMatrix(0, c_床号) = "床号"
 
        .ColWidth(c_类别) = 600
        .ColWidth(c_项目) = 2000
        .ColWidth(c_规格) = 1000
        .ColWidth(c_批次) = 0
        .ColWidth(c_换算系数) = 0
        .ColWidth(c_原始数量) = 0
        .ColWidth(c_批号) = 1000
        
        .ColWidth(c_单位) = 1000
        .ColWidth(c_付数) = 1000
        .ColWidth(c_数次) = 1000
        .ColWidth(c_已退数) = 1000
        .ColWidth(c_准退数) = 1000
        
        .ColWidth(c_退料数) = 1000
        .ColWidth(c_单价) = 1000
        .ColWidth(c_金额) = 1000
        .ColWidth(c_执行科室) = 0
        .ColWidth(c_姓名) = 1000
        .ColWidth(c_年龄) = 1000
        .ColWidth(c_床号) = 1000
        
        '-1：表示该列可以选择，是布尔型［"√"，" "］
        ' 0：表示该列可以选择，但不能修改
        ' 1：表示该列可以输入，外部显示为按钮选择
        ' 2：表示该列是日期列，外部显示为按钮选择，弹出是日期选择框
        ' 3：表示该列是选择列，外部显示为下拉框选择
        '4:  表示该列为单纯的文本框供用户输入
        '5:  表示该列不允许选择
        .ColData(c_类别) = 5
        .ColData(c_项目) = 5
        .ColData(c_规格) = 5
        .ColData(c_批号) = 5
        .ColData(c_批次) = 5
        .ColData(c_换算系数) = 5
        
        .ColData(c_单位) = 5
        .ColData(c_付数) = 5
        .ColData(c_数次) = 5
        .ColData(c_原始数量) = 5
        .ColData(c_已退数) = 5
        .ColData(c_准退数) = 5
        .ColData(c_退料数) = 4
        .ColData(c_单价) = 5
        .ColData(c_金额) = 5
        .ColData(c_执行科室) = 5
        .ColData(c_姓名) = 5
        .ColData(c_年龄) = 5
        .ColData(c_床号) = 5
            
        .PrimaryCol = c_项目
        .LocateCol = c_退料数
        .Active = True
    End With
End Sub
Private Function InitBill(ByVal strNo As String, ByVal IntBill As Integer) As Boolean
    '------------------------------------------------------------------------------------------------------
    '功能:初始单据内容
    '入参:strNO-处方号
    '    intBill-单据类型(24-收费处方发料；25-记帐单处方发料；26-记帐表处方发料)
    '出参:
    '返回:初始内容成功,返回true,否则返回false
    '修改人:刘兴宏
    '修改时间:2007/1/25
    '------------------------------------------------------------------------------------------------------
    Dim rsBill As New ADODB.Recordset
    Dim strFields As String
    Dim lngRow As Long
    Dim str门诊 As String

    err = 0: On Error GoTo ErrHand:
    '24-收费处方发料；25-记帐单处方发料；26-记帐表处方发料；
    Select Case mintUnit
    Case 0  '散装单位
         strFields = "S.计算单位 单位,D.换算系数,ltrim(to_char(S.付数,'9999999999')) 付数,S.已发数量 as 原始数量,ltrim(to_char(S.实际数量," & mOraFMT.FM_数量 & ")) 数量,ltrim(to_char(S.已退数量," & mOraFMT.FM_数量 & ")) as 已退数,ltrim(to_char(S.已发数量," & mOraFMT.FM_数量 & ")) as 准退数,'' 退料数,ltrim(to_char(S.零售价," & mOraFMT.FM_零售价 & ")) 单价,''  库存数, "
    Case Else
         strFields = "D.包装单位 单位,D.换算系数,ltrim(to_char(S.付数,'9999999999')) 付数,S.已发数量 as 原始数量,ltrim(to_char(S.实际数量/D.换算系数," & mOraFMT.FM_数量 & ")) 数量,ltrim(to_char(S.已退数量/D.换算系数," & mOraFMT.FM_数量 & ")) as 已退数,ltrim(to_char(S.已发数量/D.换算系数," & mOraFMT.FM_数量 & ")) as 准退数,'' 退料数,ltrim(to_char(S.零售价*D.换算系数," & mOraFMT.FM_零售价 & ")) 单价,'' 库存数, "
    End Select
    
    
    gstrSQL = "" & _
        " SELECT DISTINCT S.id,S.记录状态 ,S.费用ID,s.开单医生,decode(S.单据,24,'收费',25,'记帐单',26,'记帐表' ) as 类型,s.住院号,s.操作员,s.记帐员, S.ID,S.单据,S.药品ID as 材料id,S.NO,S.扣率,P.名称 科室,s.门诊标志,to_char(s.发生时间,'yyyy-mm-dd hh24:mi:ss') as 发生时间,S.年龄,s.床号,s.姓名," & _
        "            '['||X.编码||']'||X.名称  卫材名称,NVL(D.在用分批,0) 在用分批,DECODE(x.规格,NULL,x.产地,DECODE(x.产地,NULL,x.规格,x.规格||'|'||x.产地)) 规格," & strFields & _
        "  DECODE(S.批号,NULL,'',S.批号)||DECODE(S.批次,NULL,'',0,'','('||S.批次||')') 批号,NVL(S.批次,0) 批次,S.效期," & _
        "  S.零售金额 金额,S.摘要 说明,S.审核人,TO_CHAR(S.审核日期,'YYYY-MM-DD HH24:MI:SS') 发料时间,s.可操作" & _
        " FROM (    SELECT A.ID,A.NO,A.单据,A.药品ID,A.费用ID,A.批次,A.批号,A.效期,NVL(A.扣率,0) 扣率," & _
        "                   NVL(A.付数,1) 付数,A.实际数量 实际数量,NVL(A.付数,1)*A.实际数量-B.已发数量 已退数量,B.已发数量,A.记录状态," & _
        "                   A.零售价,A.零售金额,A.单量,A.频次,A.用法,A.摘要,A.审核人,A.审核日期,A.对方部门ID,A.库房ID,A.开单医生,A.计算单位,A.住院号,A.操作员,A.记帐员,A.门诊标志,A.发生时间 ,A.年龄,A.床号,A.姓名,A.可操作" & _
        "           FROM(SELECT A.ID,A.NO,A.药品id,A.序号,A.单据,A.费用ID,A.批次,A.批号,A.效期,nvl(A.扣率,0) 扣率,nvl(A.付数,0) 付数,A.实际数量,A.记录状态," & _
        "                       A.零售价,A.零售金额,A.单量,A.频次,A.用法,A.摘要,A.审核人,A.审核日期,A.对方部门id,A.库房ID," & _
        "                       m.开单人 as 开单医生,M.计算单位,m.标识号 as 住院号,m.操作员姓名 as 操作员,m.开单人 记帐员,m.门诊标志,m.发生时间,m.年龄,'' 床号,m.姓名,1 可操作 " & _
        "                FROM 药品收发记录 A,门诊费用记录 M" & _
        "                WHERE  A.审核人 IS NOT NULL and A.费用id=M.ID  and nvl(a.发药方式,0)<>-1 AND (A.记录状态=1 OR MOD(A.记录状态,3)=0)" & _
        "                       AND A.库房ID+0=[1] and a.单据=[2] and a.NO=[3] ) A," & _
        "               (SELECT A.NO,A.单据,A.药品ID,A.序号,SUM(NVL(A.付数,1)*A.实际数量) 已发数量" & _
        "                FROM 药品收发记录 A" & _
        "                WHERE A.审核人 IS NOT NULL and nvl(a.发药方式,0)<>-1 AND A.库房ID+0=[1]" & _
        "                        and A.NO=[3] and A.单据=[2]" & _
        "               GROUP BY A.NO,A.单据,A.药品ID,A.序号 ) B" & _
        "           WHERE A.NO = B.NO AND A.单据 = B.单据 AND A.药品ID+0 = B.药品ID AND A.序号 = B.序号 AND B.已发数量<>0" & _
        "       ) S,部门表 P,材料特性 D,收费项目目录 X" & _
        " WHERE S.药品ID=D.材料ID AND S.对方部门ID+0=P.ID  AND d.材料ID=X.ID" & _
        "       AND (S.记录状态=1 OR MOD(S.记录状态,3)=0) AND S.实际数量*S.付数>S.已退数量 " & _
        "       AND S.审核人 IS NOT NULL AND S.库房ID+0=[1] "
    
    If IntBill = 25 Then
        str门诊 = gstrSQL
        gstrSQL = Replace(gstrSQL, "'' 床号", "M.床号")
        gstrSQL = Replace(gstrSQL, "m.年龄", "nvl(R.年龄,m.年龄) 年龄")
        gstrSQL = Replace(gstrSQL, "m.姓名", "nvl(R.姓名,m.姓名) 姓名")
        gstrSQL = Replace(gstrSQL, "门诊费用记录 M", "住院费用记录 M,病案主页 R")
        gstrSQL = Replace(gstrSQL, "A.费用id=M.ID", "A.费用id=M.ID And M.病人id=R.病人id And M.主页id=R.主页id ")
        gstrSQL = str门诊 & " Union All " & gstrSQL
    ElseIf IntBill = 26 Then
        gstrSQL = Replace(gstrSQL, "'' 床号", "M.床号")
        gstrSQL = Replace(gstrSQL, "m.年龄", "nvl(R.年龄,m.年龄) 年龄")
        gstrSQL = Replace(gstrSQL, "m.姓名", "nvl(R.姓名,m.姓名) 姓名")
        gstrSQL = Replace(gstrSQL, "A.费用id=M.ID", "A.费用id=M.ID And M.病人id=R.病人id And M.主页id=R.主页id ")
        gstrSQL = Replace(gstrSQL, "门诊费用记录 M", "住院费用记录 M,病案主页 R")
    End If
    
    gstrSQL = gstrSQL & " Order By No,单据"
    
    Set rsBill = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, mlng发料部门, IntBill, strNo)
    With rsBill
        '清除单据内容
        Bill.ClearBill
        If .RecordCount = 0 Then
            Bill.Rows = 2
            Bill.ClearBill
            txtDown(0).Text = ""
            txtDown(1).Text = ""
            txtDown(2).Text = ""
        
            lblNote.Caption = "未输入任何处方"
            MsgBox "你所查找的处方单不存在,请检查!", vbInformation + vbDefaultButton1
            Exit Function
        Else
            Bill.Rows = .RecordCount + 1
            lblNote.Caption = "找到处方"
        End If
        txtDown(0).Text = NVL(!科室)
        txtDown(1).Text = NVL(!开单医生)
        txtDown(2).Text = NVL(!发生时间)
           
        lngRow = 1
        Do While Not .EOF
            Bill.TextMatrix(lngRow, c_类别) = NVL(!类型)
            Bill.TextMatrix(lngRow, c_项目) = NVL(!卫材名称)
            Bill.TextMatrix(lngRow, c_规格) = NVL(!规格)
            Bill.TextMatrix(lngRow, c_批号) = NVL(!批号)
            Bill.TextMatrix(lngRow, c_批次) = NVL(!批次)
            
            Bill.TextMatrix(lngRow, c_单位) = NVL(!单位)
            Bill.TextMatrix(lngRow, c_付数) = NVL(!付数)
            Bill.TextMatrix(lngRow, c_换算系数) = NVL(!换算系数)
            Bill.TextMatrix(lngRow, c_数次) = Format(Val(NVL(!数量)), mFMT.FM_数量)
            Bill.TextMatrix(lngRow, c_原始数量) = Val(NVL(!原始数量))
            Bill.TextMatrix(lngRow, c_已退数) = Format(Val(NVL(!已退数)), mFMT.FM_数量)
            Bill.TextMatrix(lngRow, c_准退数) = Format(Val(NVL(!准退数)), mFMT.FM_数量)
            Bill.TextMatrix(lngRow, c_退料数) = Format(Val(NVL(!准退数)), mFMT.FM_数量)
            Bill.TextMatrix(lngRow, c_单价) = Format(NVL(!单价), mFMT.FM_零售价)
            Bill.TextMatrix(lngRow, c_金额) = Format(NVL(!金额), mFMT.FM_金额)
            Bill.TextMatrix(lngRow, c_执行科室) = ""
            Bill.TextMatrix(lngRow, c_姓名) = NVL(!姓名)
            Bill.TextMatrix(lngRow, c_年龄) = NVL(!年龄)
            Bill.TextMatrix(lngRow, c_床号) = NVL(!床号)
            Bill.RowData(lngRow) = Val(NVL(!Id))
            lngRow = lngRow + 1
            .MoveNext
        Loop
    End With
    
    InitBill = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Resize()
    err = 0: On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.Width < 7730 Then Me.Width = 7730
    If Me.Height < 6000 Then Me.Height = 6000
     
    With fraTop
       .Top = ScaleTop + 50
       .Left = ScaleLeft + 50
       .Width = ScaleWidth - .Left
    End With
    
    With CmdCancel
        .Top = Me.ScaleHeight - .Height - 50
        .Left = Me.ScaleWidth - .Width - 100
    End With
    With cmdOK
        .Top = CmdCancel.Top
        .Left = CmdCancel.Left - 50 - .Width
    End With
    With cmdAllSel
        .Top = CmdCancel.Top
        .Left = fraTop.Left
    End With
    With cmdAllCls
        .Top = CmdCancel.Top
        .Left = cmdAllSel.Left + cmdAllSel.Width + 50
    End With
    With fraBillDown
        .Top = CmdCancel.Top - .Height - 50
        .Left = fraTop.Left
        .Width = fraTop.Width
    End With
    With Bill
        .Top = fraTop.Top + fraTop.Height
        .Left = fraTop.Left
        .Width = fraTop.Width
        .Height = fraBillDown.Top - .Top
    End With
    With lblNote
        .Left = fraTop.Width - .Width - 10
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If mblnChange = True Then
        If MsgBox("退料数据可能已改变，但还未退料，真要退出吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1
            Exit Sub
        End If
    End If
    err = 0: On Error Resume Next
    SaveWinState Me, App.ProductName, mstrCaption
    Call zlDatabase.SetPara("最后退料单据类型", Me.cbo收费单据.ItemData(Me.cbo收费单据.ListIndex), glngSys, mlngModule)
 
End Sub

Private Sub TxtNo_GotFocus()
    zlControl.TxtSelAll txtNO
End Sub

Private Sub TxtNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strNo As String, IntBill As Integer
    
    err = 0: On Error GoTo ErrHand:
    
    If cbo收费单据.ListIndex < 0 Then Exit Sub
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(txtNO) = "" Then Exit Sub
    Me.txtNO = UCase(LTrim(Me.txtNO))
    Me.txtNO.Text = zlCommFun.GetFullNo(Me.txtNO.Text, 13)
    strNo = txtNO.Text
    IntBill = cbo收费单据.ItemData(cbo收费单据.ListIndex)
    If InitBill(Me.txtNO, IntBill) = False Then
       If txtNO.Enabled Then txtNO.SetFocus
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


