VERSION 5.00
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm特殊项目审批 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "特殊项目审批"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10740
   Icon            =   "frm特殊项目审批.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   10740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmd确定 
      Caption         =   "提交数据(&O)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   7620
      TabIndex        =   15
      Top             =   6750
      Width           =   1545
   End
   Begin VB.CommandButton cmd全部不处理 
      Caption         =   "全部不处理(&N)"
      Height          =   465
      Left            =   5760
      TabIndex        =   14
      Top             =   6750
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.CheckBox chk显示所有项目 
      Caption         =   "显示已审核项目(&A)"
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   7290
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1020
      Width           =   1875
   End
   Begin VB.CommandButton cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   405
      Left            =   9450
      TabIndex        =   16
      Top             =   1410
      Width           =   1155
   End
   Begin VB.CommandButton cmd未审核项目查询 
      Caption         =   "未审核项目查询(&R)"
      Height          =   465
      Left            =   360
      TabIndex        =   11
      Top             =   6750
      Width           =   1695
   End
   Begin VB.CommandButton cmd全部转为自费 
      Caption         =   "全部转为自费(&F)"
      Height          =   465
      Left            =   4080
      TabIndex        =   13
      Top             =   6750
      Width           =   1545
   End
   Begin VB.CommandButton cmd全部审核通过 
      Caption         =   "全部审核通过(&V)"
      Height          =   465
      Left            =   2400
      TabIndex        =   12
      Top             =   6750
      Width           =   1545
   End
   Begin TabDlg.SSTab TabList 
      Height          =   6375
      Left            =   120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   990
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   11245
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "高收费项目(&1)"
      TabPicture(0)   =   "frm特殊项目审批.frx":0E42
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Bill(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "血液白蛋白(&2)"
      TabPicture(1)   =   "frm特殊项目审批.frx":0E5E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Bill(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin ZL9BillEdit.BillEdit Bill 
         Height          =   5265
         Index           =   0
         Left            =   60
         TabIndex        =   9
         Top             =   360
         Width           =   9045
         _ExtentX        =   15954
         _ExtentY        =   9287
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
      Begin ZL9BillEdit.BillEdit Bill 
         Height          =   5265
         Index           =   1
         Left            =   -74940
         TabIndex        =   10
         Top             =   360
         Width           =   9045
         _ExtentX        =   15954
         _ExtentY        =   9287
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
   End
   Begin VB.Frame Frame1 
      Caption         =   "病人基本信息"
      Height          =   735
      Left            =   210
      TabIndex        =   0
      Top             =   120
      Width           =   9135
      Begin VB.TextBox txt医保号 
         Enabled         =   0   'False
         Height          =   300
         Left            =   6780
         TabIndex        =   6
         Top             =   270
         Width           =   2085
      End
      Begin VB.TextBox txt住院号 
         Height          =   300
         Left            =   960
         TabIndex        =   2
         Top             =   270
         Width           =   1635
      End
      Begin VB.TextBox txt病人姓名 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4080
         TabIndex        =   4
         Top             =   270
         Width           =   1515
      End
      Begin VB.Label lbl医保号 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "医保号"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6150
         TabIndex        =   5
         Top             =   330
         Width           =   540
      End
      Begin VB.Label lbl住院号 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "住院号"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   330
         TabIndex        =   1
         Top             =   330
         Width           =   540
      End
      Begin VB.Label lbl病人姓名 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "病人姓名"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3270
         TabIndex        =   3
         Top             =   330
         Width           =   720
      End
   End
End
Attribute VB_Name = "frm特殊项目审批"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'程序修改清单：zl9Insure.vbp，frm特殊项目审批，frm医保帐户，mdl重庆
Private mintInsure As Integer
Private mlng病人ID As Long, mlng主页ID As Long
Private Enum ColDefine
    Col_费用ID
    Col_项目信息
    Col_规格
    Col_数量
    Col_金额
    Col_处方流水号
    Col_审核标志
    Col_Count
End Enum
Private Enum BillDefine '等同于tabList.tab
    高收费项目
    血液白蛋白
End Enum
Private Enum Marker
    不处理
    审核通过
    转为自费
End Enum

Private Const gstr费用审批 As String = "21"
'入参:住院号|处方明细流水号|审批标记
'返回:－1，审批失败；1，审批成功。
'审批标记：0，表示该处方明细流水号对应的待审批项目未通过审批，按自费项目处理；1，表示该待审批项目通过审批，费用按目录中规定的比例纳入本次结算医保费。
Private Const gstr未审核项目查询 As String = "22"
'入参:住院号|审批类型|日期
'返回:未审批项目数量|处方明细流水号1|处方明细流水号2|…|处方明细流水号n，最多返回60条，但数量是实际的数量
'审批类型：1，高收费审批；2血液白蛋白审批
'日期：格式为YYYYMMDD

Public Sub ShowMe(ByVal intInsure As Integer)
    mintInsure = intInsure
    Me.Show 1
End Sub

Private Sub Bill_DblClick(Index As Integer, Cancel As Boolean)
    With Bill(Index)
        If Trim(.TextMatrix(.Row, Col_处方流水号)) = "" Then Exit Sub
        
        If .TextMatrix(.Row, Col_审核标志) = "" Then
            .TextMatrix(.Row, Col_审核标志) = "√"
        ElseIf .TextMatrix(.Row, Col_审核标志) = "√" Then
            .TextMatrix(.Row, Col_审核标志) = "Х"
        Else
            .TextMatrix(.Row, Col_审核标志) = "√"      '只能在这两种状态中切换，不要给用户以能取消的假像
        End If
    End With
End Sub

Private Sub chk显示所有项目_Click()
    '提取所有项目并按处方流水号顺序显示
    
    If txt医保号.Text = "" Then
        MsgBox "请先确定病人！", vbInformation, gstrSysName
        txt住院号.SetFocus
        Exit Sub
    End If
    
    Call ShowData
End Sub

Private Sub cmd取消_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub cmd全部不处理_Click()
    Call BillMarker(不处理)
End Sub

Private Sub cmd全部审核通过_Click()
    Call BillMarker(审核通过)
End Sub

Private Sub cmd全部转为自费_Click()
    Call BillMarker(转为自费)
End Sub

Private Sub cmd确定_Click()
    Dim objTarget As BillEdit
    Dim str处方流水号 As String
    Dim intVerify As Integer, strInput As String, OutputData
    Dim intTab As Integer, lngRow As Long, lngRows As Long
    On Error GoTo errHand
    
    If MsgBox("你确定已完成审核向中心上传数据吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    '更新HIS数据库，同步更新医保
    For intTab = 高收费项目 To 血液白蛋白
        Set objTarget = Bill(intTab)
        lngRows = objTarget.Rows - 1
        For lngRow = 1 To lngRows
            str处方流水号 = objTarget.TextMatrix(lngRow, Col_处方流水号)
            If objTarget.TextMatrix(lngRow, Col_审核标志) = "√" Then
                intVerify = 1
            ElseIf objTarget.TextMatrix(lngRow, Col_审核标志) = "Х" Then
                intVerify = 0
            Else
                intVerify = -1
            End If
            
            '更新数据
            If intVerify <> -1 And str处方流水号 <> "" Then
                If mintInsure = TYPE_重庆市 Then
                    strInput = "21|" & GetIdentify(mlng病人ID, mlng主页ID) & "|" & str处方流水号 & "|" & intVerify
                    If HandleBusiness(strInput, OutputData) Then
                        gstrSQL = "zlYB_审核项目表_UPDATE(" & intTab + 1 & "," & mlng病人ID & "," & mlng主页ID & ",'" & str处方流水号 & "'," & intVerify + 1 & ")"
                        Call zlDatabase.ExecuteProcedure(gstrSQL, "插入审批项目表")
                    End If
                Else
                    gstrSQL = "Update 中间库_处方明细 Set 高收费审批编号='" & IIf(intVerify = 1, "11", "00") & "' Where 处方流水号='" & str处方流水号 & "'"
                    gcn重庆银海版.Execute gstrSQL
                    
                    gstrSQL = "zlYB_审核项目表_UPDATE(" & intTab + 1 & "," & mlng病人ID & "," & mlng主页ID & ",'" & str处方流水号 & "'," & intVerify + 1 & ")"
                    gcn重庆银海版.Execute gstrSQL, , adCmdStoredProc
                End If
            End If
        Next
    Next
    
    Call ShowData
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmd未审核项目查询_Click()
    Dim strInput As String
    Dim lngCount As Long, lngMin As Long, LNGMAX As Long
    Dim OutputData
    
    If Me.txt住院号.Text = "" Then
        MsgBox "请先确定病人身份!", vbInformation, gstrSysName
        Me.txt住院号.SetFocus
        Exit Sub
    End If
    
    strInput = gstr未审核项目查询 & "|" & GetIdentify(mlng病人ID, mlng主页ID) & "|" & TabList.Tab + 1 & "|" & Format(zlDatabase.Currentdate(), "yyyyMMdd")
    If HandleBusiness(strInput, OutputData) Then
        '依次插入数据库
        lngCount = Val(OutputData(1))
        LNGMAX = IIf(lngCount > 60, 60, lngCount)
        For lngMin = 1 To LNGMAX
            gstrSQL = "zlYB_审核项目表_UPDATE(" & TabList.Tab + 1 & "," & mlng病人ID & "," & mlng主页ID & ",'" & OutputData(lngMin + 1) & "',0)"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "插入审批项目表")
        Next
        
        Call ShowData
        If lngCount > 60 Then
            MsgBox "本次共有" & lngCount & "条明细需要审核,由于东软接口一次最多返回60条明细,请审批完当前数据后重新获取余下的待审核明细进行审批！", vbInformation, gstrSysName
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call InitBill(Bill(高收费项目))
    Call InitBill(Bill(血液白蛋白))
    
    cmd未审核项目查询.Visible = (mintInsure = TYPE_重庆市)
    If mintInsure = TYPE_重庆银海版 Then Me.Caption = "特殊项目审批(银海用户注意，已审批的明细上传后不允许修改审批状态)"
    Call gclsInsure.InitInsure(gcnOracle, mintInsure)
End Sub

Private Sub BillMarker(ByVal intState As Integer)
    Dim strState As String
    Dim objTarget As BillEdit
    Dim lngRow As Long, lngRows As Long
    
    strState = IIf(intState = 1, "√", IIf(intState = 2, "Х", ""))
    Set objTarget = Bill(TabList.Tab)
    lngRows = objTarget.Rows - 1
    
    For lngRow = 1 To lngRows
        objTarget.TextMatrix(lngRow, Col_审核标志) = strState
    Next
End Sub

Private Sub InitBill(ByVal objTarget As BillEdit)
    With objTarget
        .ClearBill
        .Rows = 2
        .Cols = Col_Count
        
        .TextMatrix(0, Col_费用ID) = "ID"
        .TextMatrix(0, Col_项目信息) = "项目信息"
        .TextMatrix(0, Col_规格) = "规格"
        .TextMatrix(0, Col_数量) = "数量"
        .TextMatrix(0, Col_金额) = "金额"
        .TextMatrix(0, Col_处方流水号) = "处方流水号"
        .TextMatrix(0, Col_审核标志) = "审核标志"
        .ColWidth(Col_费用ID) = 0
        .ColWidth(Col_项目信息) = 2200
        .ColWidth(Col_规格) = 1500
        .ColWidth(Col_数量) = 1200
        .ColWidth(Col_金额) = 1200
        .ColWidth(Col_处方流水号) = 1300
        .ColWidth(Col_审核标志) = 800
        .ColData(Col_费用ID) = 5
        .ColData(Col_项目信息) = 5
        .ColData(Col_规格) = 5
        .ColData(Col_数量) = 5
        .ColData(Col_金额) = 5
        .ColData(Col_处方流水号) = 5
        .ColData(Col_审核标志) = 0
        
        .PrimaryCol = Col_费用ID
        .LocateCol = Col_审核标志
        .AllowAddRow = False
        .Active = True
    End With
End Sub

Private Sub ShowData()
    Dim objTarget As BillEdit
    Dim rsTemp As New ADODB.Recordset
    
    Set objTarget = Bill(TabList.Tab)
    Call InitBill(objTarget)
    
    '提取所有项目
    If mintInsure = TYPE_重庆市 Then
        gstrSQL = " Select A.ID,'['||C.编码||']'||C.名称 AS 项目信息,C.规格,A.付数*A.数次 AS 数量,A.实收金额 AS 金额,B.处方流水号,B.审核标志" & _
                  " From 住院费用记录 A,审核项目表 B,收费细目 C" & _
                  " Where Substr(A.摘要||'|',1,Instr(A.摘要||'|','|',1,1)-1)=B.处方流水号 And A.病人ID=B.病人ID And A.主页ID=B.主页ID " & _
                  " And A.收费细目ID=C.ID And Nvl(A.实收金额,0)<>0" & _
                  " And B.病人ID=" & mlng病人ID & " And B.主页ID=" & mlng主页ID & " And B.类型=" & TabList.Tab + 1 & _
                  IIf(chk显示所有项目.Value = 1, "", " And Nvl(B.审核标志,0)=0") & _
                  " Order by B.处方流水号"
        Call OpenRecordset(rsTemp, "提取所有项目")
    Else
        gstrSQL = " Select 0 AS ID,'['||C.编码||']'||C.名称 AS 项目信息,C.规格,A.数量,A.金额,A.处方流水号,B.审核标志" & _
                  " From 中间库_处方明细 A,审核项目表 B,ZLHIS.收费细目 C" & _
                  " Where A.处方流水号=B.处方流水号 And A.项目编码=C.编码" & _
                  " And B.病人ID=" & mlng病人ID & " And B.主页ID=" & mlng主页ID & " And B.类型=" & TabList.Tab + 1 & _
                  IIf(chk显示所有项目.Value = 1, "", " And Nvl(B.审核标志,0)=0") & _
                  " Order by B.处方流水号"
        Call OpenRecordset(rsTemp, "提取所有项目", gstrSQL, gcn重庆银海版)
    End If
    
    With rsTemp
        Do While Not .EOF
            objTarget.TextMatrix(.AbsolutePosition, Col_费用ID) = !ID
            objTarget.TextMatrix(.AbsolutePosition, Col_项目信息) = !项目信息
            objTarget.TextMatrix(.AbsolutePosition, Col_规格) = Nvl(!规格)
            objTarget.TextMatrix(.AbsolutePosition, Col_数量) = !数量
            objTarget.TextMatrix(.AbsolutePosition, Col_金额) = !金额
            objTarget.TextMatrix(.AbsolutePosition, Col_处方流水号) = !处方流水号
            objTarget.TextMatrix(.AbsolutePosition, Col_审核标志) = IIf(!审核标志 = 1, "Х", IIf(!审核标志 = 2, "√", ""))
            
            .MoveNext
            objTarget.Rows = objTarget.Rows + 1
        Loop
    End With
End Sub

Private Sub TabList_Click(PreviousTab As Integer)
    Call ShowData
End Sub

Private Sub txt住院号_KeyDown(KeyCode As Integer, Shift As Integer)
    '提取该病人的基本信息，非医保病人直接退出
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    '提取病人基本信息
    gstrSQL = " Select A.病人ID,A.住院次数 AS 主页ID,A.姓名,A.住院号,B.医保号 " & _
              " From 病人信息 A,保险帐户 B" & _
              " Where A.病人ID=B.病人ID And B.险类=" & mintInsure & " And A.住院号='" & txt住院号.Text & "'"
    Call OpenRecordset(rsTemp, "提取病人基本信息")
    If rsTemp.RecordCount = 0 Then
        txt病人姓名.Text = ""
        txt医保号.Text = ""
        MsgBox "该病人不属于医保病人，或者录入的住院号不存在！", vbInformation, gstrSysName
        txt住院号.SetFocus
        Exit Sub
    End If
    
    '显示病人的基本信息
    Me.txt病人姓名.Text = rsTemp!姓名
    Me.txt医保号.Text = rsTemp!医保号
    mlng病人ID = rsTemp!病人ID
    mlng主页ID = rsTemp!主页ID
    
    Call ShowData
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
