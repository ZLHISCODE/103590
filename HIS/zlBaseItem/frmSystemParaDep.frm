VERSION 5.00
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmSystemParaDep 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "科室编号"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5565
   Icon            =   "frmSystemParaDep.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton CmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   300
      Left            =   180
      TabIndex        =   3
      Top             =   4140
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "取消(&C)"
      Height          =   300
      Left            =   4200
      TabIndex        =   2
      Top             =   4140
      Width           =   1100
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "确定(&O)"
      Height          =   300
      Left            =   2910
      TabIndex        =   1
      Top             =   4140
      Width           =   1100
   End
   Begin ZL9BillEdit.BillEdit Bill药品科室 
      Height          =   3885
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   5355
      _ExtentX        =   9446
      _ExtentY        =   6853
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
Attribute VB_Name = "frmSystemParaDep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mIntItem As Integer                 '保存的记录数
Dim mIntSequence As Integer             '号码序号
Dim mCboIndex As Integer                '列表控件Index
Private Const mstrChar As String = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"

Private Sub Bill药品科室_cboClick(ListIndex As Long)
    mCboIndex = ListIndex
End Sub

Private Sub Bill药品科室_cboKeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        With Me.Bill药品科室
            .RowData(.Row) = .ItemData(mCboIndex)
        End With
    End If
End Sub

Private Sub Bill药品科室_EditChange(curText As String)
    If Len(curText) > 1 Then
        Bill药品科室.Text = Mid(curText, 1, 1)
    End If
    Bill药品科室.Text = UCase(Bill药品科室.Text)
    Bill药品科室.SelStart = 1
    Bill药品科室.SelLength = Len(Bill药品科室.Text)
End Sub

Private Sub Bill药品科室_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    If KeyCode = 13 Then
        Me.CmdOK.SetFocus
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    '检查数据是否合法
    If IsValid = True Then Exit Sub
    Call Save药品科室
    Unload Me
End Sub

Private Sub Form_Load()
    InitAll
End Sub
Public Function ShowMe(objfrm As Object, IntSequence As Integer, DepStr As String) As Boolean
    '''''''''''''''''''''''''''''''''''''''''
    '功能               提供给上级窗体调用
    '参数
    'IntSequence        序号ID
    'DepStr             科室和科室编号字串
    '返回               科室和科室编号字串
    '''''''''''''''''''''''''''''''''''''''''
    mIntSequence = IntSequence
    mDepStr = DepStr
    Me.Show vbModal, objfrm
End Function
Sub InitAll()
    Call InitBill
    Call load药品科室
End Sub

Sub InitBill()
    With Bill药品科室
        .Cols = 2 '多了一列隐藏列
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .TextMatrix(0, 0) = "部门"
        .TextMatrix(0, 1) = "编号"
        .ColWidth(0) = 2000
        .ColWidth(1) = 1000
        .ColData(0) = 3
        .ColData(1) = 4
        .PrimaryCol = 0
        .Active = True
        .TxtCheck = True
        .TextMask = mstrChar
        .PrimaryCol = 0
    End With
End Sub
Sub load药品科室()
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer
    gstrSQL = "select distinct A.ID,A.名称,A.编码 " & _
                   " from  部门性质说明 b,部门表 a " & _
                   " where B.工作性质 in ('中药库','西药库','成药库','制剂室','中药房','西药房','成药房') " & _
                   " and  b.部门ID=a.ID  order by 编码"
    On Error GoTo errH
    zldatabase.OpenRecordset rsTmp, gstrSQL, Me.Caption
    i = 0
    With Bill药品科室
        Do Until rsTmp.EOF
            .AddItem rsTmp("编码") & "-" & rsTmp("名称")
            .ItemData(i) = rsTmp("id")
            rsTmp.MoveNext
            i = i + 1
        Loop

        gstrSQL = "select A.编号, B.id, B.编码, B.名称 from 科室号码表 a ,部门表 b where a.科室id = b.id and 项目序号 = [1] "
        Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, mIntSequence)
        
        Do Until rsTmp.EOF
            .TextMatrix(.Rows - 1, 0) = rsTmp("编码") & "-" & rsTmp("名称")
            .TextMatrix(.Rows - 1, 1) = rsTmp("编号")
            .RowData(.Rows - 1) = rsTmp("id")
            .Rows = .Rows + 1
            rsTmp.MoveNext
        Loop
        If .Rows > 2 Then
            .Rows = .Rows - 1
        End If
    End With
    Exit Sub
errH:
    If ERRCENTER() = 1 Then Resume
End Sub
Private Function IsValid() As Boolean
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能           检查科室编号输入是否正确
    '返回           =True表示有问题 =False表示可以通过去时
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim strDept As String
    Dim strNumber As String
    With Me.Bill药品科室
        For i = 1 To .Rows - 1
            If Len(Trim(.TextMatrix(i, 0))) > 0 And Len(Trim(.TextMatrix(i, 1))) > 0 Then
                If InStr(1, strDept & ",", "," & .TextMatrix(i, 0) & ",") > 0 Then
                    MsgBox "第" & i & "行出现科室重复!", vbInformation, gstrSysName
                    .Row = i
                    .Col = 0
                    .TxtSetFocus
                    IsValid = True
                    Exit Function
                End If
                strDept = strDept & "," & .TextMatrix(i, 0)
                
                If InStr(1, strDept & ",", "," & .TextMatrix(i, 1) & ",") > 0 Then
                    If InStr(1, strDept & ",", "," & .TextMatrix(i, 1) & ",") > 0 Then
                        If MsgBox("第" & i & "行出现科室编号重复!", vbYesNo + vbInformation, gstrSysName) = vbNo Then
                            .Row = i
                            .Col = 1
                            .TxtSetFocus
                            IsValid = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        Next
    End With
End Function
Sub Save药品科室()
    
    On Error GoTo errH
    '先删除以前的再保存
    gstrSQL = "ZL_科室号码表_DELETE(" & mIntSequence & ")"
    zldatabase.ExecuteProcedure gstrSQL, Me.Caption
    With Me.Bill药品科室
        For i = 1 To .Rows - 1
            If Len(Trim(.TextMatrix(i, 0))) > 0 And Len(Trim(.TextMatrix(i, 1))) > 0 Then
                gstrSQL = "ZL_科室号码表_INSERT(" & mIntSequence & "," & .RowData(i) & ",'" & .TextMatrix(i, 1) & "')"
                zldatabase.ExecuteProcedure gstrSQL, Me.Caption
            End If
        Next
    End With
    Exit Sub
errH:
    If ERRCENTER() = 1 Then Resume
End Sub

