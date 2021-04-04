VERSION 5.00
Begin VB.Form frmPriorityCause 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "优先原因"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   Icon            =   "frmPriorityCause.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdDel 
      Height          =   300
      Left            =   4250
      Picture         =   "frmPriorityCause.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "删除当前选择的常用原因"
      Top             =   2750
      Width           =   300
   End
   Begin VB.CommandButton cmdAdd 
      Height          =   300
      Left            =   3950
      Picture         =   "frmPriorityCause.frx":699C
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "将当前原因设为常用原因"
      Top             =   2750
      Width           =   300
   End
   Begin VB.ComboBox cboPriorityCause 
      Height          =   300
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   3855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   3240
      Width           =   1075
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   3240
      Width           =   1075
   End
   Begin VB.ListBox lstJQueueList 
      Height          =   2040
      Left            =   75
      TabIndex        =   0
      Top             =   315
      Width           =   4480
   End
   Begin VB.Label lblCause 
      Caption         =   "调整原因"
      Height          =   255
      Left            =   90
      TabIndex        =   4
      Top             =   2430
      Width           =   1455
   End
   Begin VB.Label lblTitle 
      Caption         =   "号码      患者姓名        状态"
      Height          =   255
      Left            =   90
      TabIndex        =   3
      Top             =   90
      Width           =   3405
   End
End
Attribute VB_Name = "frmPriorityCause"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrQueueName As String        '队列名称
Private mstrCruuentWorkID As String      '业务ID
Private mfrmParent As Form             '父窗体
Private mstrTempQueueName As String    '当前选择数据的队列名称
Private mstrSelectedName As String     '当前选择的患者姓名
Private mintCurNextIndex As String     '待调整数据的下一条数据的Index
Private mstrMaxCode As String          '获取优先原因的最大编码
Private mstrArrQueueNum() As String

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const CB_LIMITTEXT = &H141

Private Enum mCol
    队列名称 = 0: Id: 病人ID: 排队标记: 排队号码:  排队序号: 患者姓名: 优先: 回诊序号: 回诊排序号: 科室ID: 诊室: 医生姓名: 排队状态: 排队时间: 呼叫医生: 业务类型: 业务ID: 呼叫时间
End Enum




Private Sub cmdDel_Click()
'功能: 删除优先原因
    Dim strSql As String
    Dim strDelCode As String
    Dim strMaxCode As String
    Dim i As Integer
    
    On Error GoTo errHandle
    
    If cboPriorityCause.ListIndex = -1 Then
        cboPriorityCause.Text = ""
        cboPriorityCause.SetFocus
        Exit Sub
    End If
    
    strDelCode = cboPriorityCause.ItemData(cboPriorityCause.ListIndex)
    strSql = "zl_排队优先原因_delete('" & Format(strDelCode, "00000") & "')"
    Call zlDatabase.ExecuteProcedure(strSql, "优先原因")
    
    Call cboPriorityCause.RemoveItem(cboPriorityCause.ListIndex)
    
    strMaxCode = "0"
    
    If CLng(strDelCode) = CLng(mstrMaxCode) Then '获取删除最大号之后的最大code
        If cboPriorityCause.ListCount > 0 Then
            For i = 0 To cboPriorityCause.ListCount - 1
                If cboPriorityCause.ItemData(i) > strMaxCode Then
                    strMaxCode = cboPriorityCause.ItemData(i)
                End If
            Next
        End If
        
        mstrMaxCode = strMaxCode
    End If
    
    If cboPriorityCause.ListCount <= 0 Then mstrMaxCode = ""
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cmdAdd_Click()
'功能: 新增优先原因
    Dim i As Integer
    Dim strSql As String
    
    On Error GoTo errHandle
    
    If Trim(cboPriorityCause.Text) = "" Then
        MsgBox "请输入需要新增的内容！", vbOKOnly Or vbInformation, Me.Caption
        cboPriorityCause.SetFocus
        Exit Sub
    End If
    
    For i = 0 To cboPriorityCause.ListCount - 1
        If UCase(Trim(cboPriorityCause.List(i))) = UCase(Trim(cboPriorityCause.Text)) Then
            MsgBox "该内容已经在优先原因中！", vbOKOnly Or vbInformation, Me.Caption
            cboPriorityCause.SetFocus
            Exit Sub
        End If
    Next
    
    strSql = "zl_排队优先原因_insert('" & Trim(cboPriorityCause.Text) & "','" & zlCommFun.zlGetSymbol(cboPriorityCause.Text) & "')"
    Call zlDatabase.ExecuteProcedure(strSql, "优先原因")
    
    Call cboPriorityCause.AddItem(Trim(cboPriorityCause.Text))
    cboPriorityCause.ItemData(cboPriorityCause.ListCount - 1) = CLng(IIf(mstrMaxCode = "", 0, mstrMaxCode) + 1)
    mstrMaxCode = IIf(mstrMaxCode = "", 0, mstrMaxCode) + 1
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub LoadPriorityCause()
'功能: 加载优先原因
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    cboPriorityCause.Clear
    
    strSql = "select 编码,名称,使用频率 from 排队优先原因 order by 使用频率 desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "优先原因")
    
    If rsTemp.RecordCount <= 0 Then Exit Sub

    rsTemp.MoveFirst
    mstrMaxCode = Nvl(rsTemp!编码)
    
    Do While Not rsTemp.EOF
        cboPriorityCause.AddItem Nvl(rsTemp!名称)
        cboPriorityCause.ItemData(cboPriorityCause.ListCount - 1) = Nvl(rsTemp!编码)
        If CLng(Nvl(rsTemp!编码)) > CLng(mstrMaxCode) Then mstrMaxCode = CLng(Nvl(rsTemp!编码))
        rsTemp.MoveNext
    Loop
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cmdCancel_Click()
On Error GoTo errHandle
    
    '隐藏窗体
    Me.Hide

    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub ShowPriorityCause(frmParent As Form, ByVal strCurrentCaption As String, ByVal strCurrentWorkID As String, _
                            ByVal strTempQueueName As String, ByVal strSelectedName As String)
    
    '保存传入参数
    mstrQueueName = strCurrentCaption
    mstrCruuentWorkID = strCurrentWorkID
    mstrTempQueueName = strTempQueueName
    mstrSelectedName = strSelectedName
    
    Set mfrmParent = frmParent
    
    '加载List控件数据
    Call LoadListData
    
    '打开窗体
    Me.Show 1, frmParent
    
End Sub

Private Sub cmdOK_Click()
On Error GoTo errHandle
    Dim i As Integer
    Dim strSql As String
    Dim intJQueueID As Long        '需插队ID
    Dim intNeedJQueueID As Long    '被插队ID
    
    '判断是否选择数据列
    If lstJQueueList.SelCount < 1 Then
        MsgBox "请选择被插队的病人。", vbOKOnly Or vbInformation, Me.Caption
        Exit Sub
    End If
    
    '判断是否选择了相同的数据
    If mstrArrQueueNum(lstJQueueList.ListIndex) = mstrSelectedName Then
        MsgBox "不能插自己的队。", vbOKOnly Or vbInformation, Me.Caption
        Exit Sub
    End If
    
    '判断是否选择了当前数据的下一条数据，因为插队到下一条数据其实队列位置没变。
    If lstJQueueList.ListIndex = mintCurNextIndex Then
        MsgBox "插队到该病人前，队列位置不变，请另外选择。", vbOKOnly Or vbInformation, Me.Caption
        Exit Sub
    End If
    
    '判断录入原因是否为空
    If Trim(cboPriorityCause.Text) = "" Then
         MsgBox "优先原因为空。", vbOKOnly Or vbInformation, Me.Caption
         Exit Sub
    End If
    
    '得到被插队 和 需插队的 排队ID
    intJQueueID = Val(Mid(mstrArrQueueNum(lstJQueueList.ListIndex), InStr(mstrArrQueueNum(lstJQueueList.ListIndex), ",") + 1, 100))
    intNeedJQueueID = Val(Mid(mstrSelectedName, InStr(mstrSelectedName, ",") + 1, 100))
    
    '执行插队 队列修改以及原因写入
    strSql = "ZL_排队叫号队列_优先('" & mstrQueueName & "'," & mstrCruuentWorkID & ",'" & Trim(cboPriorityCause.Text) & "'," & intNeedJQueueID & "," & intJQueueID & ")"
    zlDatabase.ExecuteProcedure strSql, "优先原因"
    
    For i = 0 To cboPriorityCause.ListCount - 1
        If UCase(Trim(cboPriorityCause.List(i))) = UCase(Trim(cboPriorityCause.Text)) Then
            '更新使用频率
            strSql = "zl_排队优先原因_Update('" & Format(cboPriorityCause.ItemData(i), "00000") & "')"
            Call zlDatabase.ExecuteProcedure(strSql, "优先原因")
            
            Me.Hide
            Exit Sub
        End If
    Next
    '将优先原因写入数据库
    strSql = "zl_排队优先原因_insert('" & Trim(cboPriorityCause.Text) & "','" & zlCommFun.zlGetSymbol(cboPriorityCause.Text) & "',1)"
    Call zlDatabase.ExecuteProcedure(strSql, "优先原因")
    
    Call cboPriorityCause.AddItem(Trim(cboPriorityCause.Text))
    cboPriorityCause.ItemData(cboPriorityCause.ListCount - 1) = CLng(IIf(mstrMaxCode = "", 0, mstrMaxCode) + 1)
    mstrMaxCode = IIf(mstrMaxCode = "", 0, mstrMaxCode) + 1
    '完成后隐藏窗体
    Me.Hide
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub LoadListData()
'加载ListBox数据
    Dim i As Integer
    Dim j As Integer
    
    j = 0
    With mfrmParent.rptQueueList
        ReDim mstrArrQueueNum(.Rows.Count)
        
        For i = 0 To .Rows.Count - 1
            If .Rows(i).GroupRow <> True Then
                '通过对比判断只加载控件中当前选中队列的数据
                If .Rows(i).Record(mCol.队列名称).value = mstrTempQueueName Then
                
                    '将对应的数据存入数组
                    mstrArrQueueNum(j) = .Rows(i).Record(mCol.排队号码).value & .Rows(i).Record(mCol.患者姓名).value & "," & .Rows(i).Record(mCol.Id).value
                
                    '给ListBox 赋值
                    lstJQueueList.List(j) = "  " & .Rows(i).Record(mCol.排队号码).value & "号   " & .Rows(i).Record(mCol.患者姓名).value & IIf(mstrArrQueueNum(j) = mstrSelectedName, "   待调整", "")
                    
                    If mstrArrQueueNum(j) = mstrSelectedName Then mintCurNextIndex = j + 1
                    
                    j = j + 1
                End If
            End If
        Next i
    End With
    
    '默认选中第一项
    If lstJQueueList.ListCount > 0 Then lstJQueueList.ListIndex = 0
    
End Sub

Private Sub Form_Load()
    '加载优先原因
    Call LoadPriorityCause
    '优先原因的最大长度限制为64位
    SendMessage cboPriorityCause.hwnd, CB_LIMITTEXT, 64, 0&
End Sub
