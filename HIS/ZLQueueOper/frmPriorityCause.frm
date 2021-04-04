VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Begin VB.Form frmPriorityCause 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "插队"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9765
   Icon            =   "frmPriorityCause.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7758.509
   ScaleMode       =   0  'User
   ScaleWidth      =   9765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin XtremeReportControl.ReportControl rptQueueList 
      Height          =   6450
      Left            =   90
      TabIndex        =   0
      Tag             =   "0"
      Top             =   90
      Width           =   9585
      _Version        =   589884
      _ExtentX        =   16907
      _ExtentY        =   11377
      _StockProps     =   0
      BorderStyle     =   3
      AllowColumnSort =   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin VB.ComboBox cbxMemo 
      Height          =   300
      Left            =   960
      TabIndex        =   4
      Top             =   6720
      Width           =   3870
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   420
      Left            =   8280
      TabIndex        =   2
      Top             =   6675
      Width           =   1380
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   420
      Left            =   6720
      TabIndex        =   1
      Top             =   6675
      Width           =   1380
   End
   Begin VB.Label labMemo 
      Alignment       =   1  'Right Justify
      Caption         =   "备注："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   60
      TabIndex        =   3
      Top             =   6735
      Width           =   870
   End
End
Attribute VB_Name = "frmPriorityCause"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfrmParent As Object            '用户控件UcQueueStation对象
Private mobjQueueManage As clsQueueOperation
Attribute mobjQueueManage.VB_VarHelpID = -1

Private mlngInsertQueueRow As Long      '待插队的行索引
Private mlngQueueId As Long
Private mlngWorkType As Long            '业务类型
Private mlngSelRow As Long
Private mblnOk As Boolean
Private mstrReason As String            '插队原因



Private Sub cmdCancel_Click()
On Error GoTo errHandle
    
    mblnOk = False
    
    Call Me.Hide

    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Public Function ShowPriorityCause(frmParent As UcQueue, objQueueReportControl As Object, _
    ByVal lngInsertQueueRow As Long, ByVal lngWorkType As Long, Optional ByVal strReason As String = "") As Boolean
    
    '保存传入参数
    mblnOk = False
    mstrReason = strReason
    mlngSelRow = -1
    mlngInsertQueueRow = lngInsertQueueRow
    mlngQueueId = Val(objQueueReportControl.Rows(lngInsertQueueRow).Record(GetColIndex("ID", objQueueReportControl)).value) '排队ID
    mlngWorkType = lngWorkType
        
    Set mfrmParent = frmParent
    Set mobjQueueManage = frmParent.QueueOper

    Call ConfigFont(frmParent.Font)
    
    Call CopyDataToReportControl(objQueueReportControl, rptQueueList, lngInsertQueueRow)
    
    '打开窗体
    Me.Show 1, frmParent

    ShowPriorityCause = mblnOk
End Function

Private Sub ConfigFont(ft As StdFont)
    Set cmdOK.Font = ft
    Set cmdCancel.Font = ft
    Set labMemo.Font = ft
    Set cbxMemo.Font = ft
End Sub


Private Sub CopyDataToReportControl(objSourceRC As ReportControl, objTargetRC As ReportControl, ByVal lngInsertQueueRow As Long)
    Dim i As Long, j As Long
    Dim Column As ReportColumn
    Dim rptRecord As ReportRecord
    Dim strQueueName As String
    Dim lngQueueNameColIndex As Long
    Dim lngCurInsertRow As Long
    Dim lngColIndex As Long
    
    Set objTargetRC.PaintManager.CaptionFont = objSourceRC.PaintManager.CaptionFont
    Set objTargetRC.PaintManager.TextFont = objSourceRC.PaintManager.TextFont
    
    '配置显示列
    objTargetRC.Columns.DeleteAll
    
    For i = 0 To objSourceRC.Columns.Count - 1
        Set Column = objTargetRC.Columns.Add(i, objSourceRC.Columns(i).Caption, objSourceRC.Columns(i).Width, True)
        Column.Groupable = objSourceRC.Columns(i).Groupable
        Column.Visible = objSourceRC.Columns(i).Visible
    Next i
    
    
    lngQueueNameColIndex = GetColIndex("队列名称", objSourceRC)
    strQueueName = objSourceRC.Rows(lngInsertQueueRow).Record(lngQueueNameColIndex).value
    
    lngCurInsertRow = -1
    objTargetRC.Records.DeleteAll
    
    '配置数据
    For i = 0 To objSourceRC.Rows.Count - 1
        If objSourceRC.Rows(i).GroupRow <> True Then
            If objSourceRC.Rows(i).Record(lngQueueNameColIndex).value = strQueueName Then
                Set rptRecord = objTargetRC.Records.Add
                lngCurInsertRow = lngCurInsertRow + 1
                
                For j = 0 To objSourceRC.Columns.Count - 1
                    lngColIndex = objSourceRC.Columns(j).ItemIndex
                    
                    If i = lngInsertQueueRow Then
                        rptRecord.AddItem objSourceRC.Rows(i).Record(lngColIndex).value '"当前排队位置"
                        If objSourceRC.Columns(j).Caption = "排队号码" Then rptRecord.Item(j).Icon = 3558
                        
                        rptRecord.Item(j).BackColor = &H80FF80
                        
                        '从新更新位置行信息
                        mlngInsertQueueRow = lngCurInsertRow
                    Else
                        rptRecord.AddItem objSourceRC.Rows(i).Record(lngColIndex).value
                        If objSourceRC.Columns(j).Caption = "排队号码" Then
                            rptRecord.Item(j).Icon = objSourceRC.Rows(i).Record(lngColIndex).Icon ' 807    ' IIf(i = lngInsertQueueRow + 1, 8216, 807)
                        End If
                        
                        rptRecord.Item(j).BackColor = IIf(i = lngInsertQueueRow + 1, &HC0C0C0, vbWhite)
                    End If
                
                Next j

            End If
        End If
    Next i
    
    Set rptRecord = objTargetRC.Records.Add
    For j = 0 To objSourceRC.Columns.Count - 1
        rptRecord.AddItem ""
    Next j
    
    objTargetRC.SortOrder.DeleteAll
    objTargetRC.AllowColumnSort = False
    
    objTargetRC.Populate
    
    If mlngInsertQueueRow <> 0 Then
        objTargetRC.Rows(0).Selected = True
    End If
End Sub

Private Sub InitInsertQueueList()
    '初始化排队队列显示字段
    Call rptQueueList.Columns.DeleteAll
    
    Set rptQueueList.Icons = zlCommFun.GetPubIcons

    '初始化列表相关属性
    rptQueueList.AllowColumnRemove = False
    rptQueueList.ShowItemsInGroups = False
    rptQueueList.SkipGroupsFocus = True
    rptQueueList.MultipleSelection = False

    With rptQueueList.PaintManager
        .ColumnStyle = xtpColumnShaded
        .GridLineColor = RGB(225, 225, 225)
        .NoGroupByText = "将列标题拖动到此,可按该列分组..."
        .NoItemsText = "没有可显示的项目..."
        .VerticalGridStyle = xtpGridSolid
    End With
    
    rptQueueList.AllowColumnSort = False
End Sub

Private Function AllowInsert() As String
'判断所选位置是否允许插队
    AllowInsert = ""
    
    If rptQueueList.SelectedRows.Count <= 0 Then
        AllowInsert = "请选择要插入到的队列位置。"
        Exit Function
    End If
    
    If rptQueueList.SelectedRows(0).Index = mlngInsertQueueRow _
        Or rptQueueList.SelectedRows(0).Index = mlngInsertQueueRow + 1 Then
        AllowInsert = "不能在当前位置进行插队操作。"
        Exit Function
    End If
    
    
End Function



Private Function GetInsertOrder(ByRef lngRowIndex As Long) As String
'获取插队的位置，返回序号1和序号2，插入位置在序号1和序号2之间
    Dim strSql As String
    Dim strQueueName As String

    Dim strLastQueueOrder As String
    Dim strCurrQueueOrder As String
    
    Dim lngLastQueueID As Long
    Dim lngCurrQueueID As Long
    
    Dim strCustomOrderWhere As String
    
    Dim rsData As ADODB.Recordset

    GetInsertOrder = ""
    lngRowIndex = -1
    
    strCustomOrderWhere = mfrmParent.CustomOrderField
    If Trim(strCustomOrderWhere) = "" Then
        strCustomOrderWhere = mobjQueueManage.CustomOrder
    End If
    
    strQueueName = rptQueueList.SelectedRows(0).Record(GetColIndex("队列名称", rptQueueList)).value
    lngCurrQueueID = Val(rptQueueList.SelectedRows(0).Record(GetColIndex("ID", rptQueueList)).value)
    
    strCurrQueueOrder = mobjQueueManage.GetOrder(lngCurrQueueID)
    
    lngRowIndex = rptQueueList.SelectedRows(0).Index
    
    '插队前，需要判断排队队列直接是否可能存在暂停弃号等队列数据
    If rptQueueList.SelectedRows(0).Index = rptQueueList.Rows.Count - 1 Then
        '表示插入到最后面
        GetInsertOrder = mobjQueueManage.GetInsertOrder(mlngQueueId, strLastQueueOrder, -1)
        Exit Function
        
    ElseIf rptQueueList.SelectedRows(0).Index > 0 Then
        lngLastQueueID = Val(rptQueueList.Rows(rptQueueList.SelectedRows(0).Index - 1).Record(GetColIndex("ID", rptQueueList)).value)
        strLastQueueOrder = mobjQueueManage.GetOrder(lngLastQueueID)
        
        strSql = "select Id,排队序号 from 排队叫号队列 where 业务类型=[1] and 队列名称=[2] " & _
                IIf(Trim(strCustomOrderWhere) <> "", " order by " & strCustomOrderWhere, "")
                
        Set rsData = zlDatabase.OpenSQLRecord(strSql, "查询未排队队列", mlngWorkType, strQueueName)
        If rsData.RecordCount > 0 Then
            Do While Not rsData.EOF
                If Nvl(rsData!排队序号) = strCurrQueueOrder Then
                    Call rsData.MovePrevious
                    
                    strLastQueueOrder = Nvl(rsData!排队序号)
                    Exit Do
                End If
                
                Call rsData.MoveNext
            Loop
        End If
        
    Else
        strSql = "select Id,排队序号 from 排队叫号队列 where  业务类型=[1] and 队列名称=[2] " & _
                IIf(Trim(strCustomOrderWhere) <> "", " order by " & strCustomOrderWhere, "")
                
        Set rsData = zlDatabase.OpenSQLRecord(strSql, "查询未排队队列", mlngWorkType, strQueueName)
        If rsData.RecordCount > 0 Then
            Do While Not rsData.EOF
                If Nvl(rsData!排队序号) = strCurrQueueOrder Then
                    If rsData.AbsolutePosition <> 1 Then
                        Call rsData.MovePrevious
                        strLastQueueOrder = Nvl(rsData!排队序号)
                    Else
                        strLastQueueOrder = 0
                    End If
                                        
                    Exit Do
                End If
                
                Call rsData.MoveNext
            Loop
        Else
            strLastQueueOrder = 0
        End If
        
    End If
    
    '获取新的插队序号
    GetInsertOrder = mobjQueueManage.GetInsertOrder(mlngQueueId, strLastQueueOrder, strCurrQueueOrder)
End Function

Private Sub cmdOK_Click()
On Error GoTo errHandle
    Dim strNewOrder As String
    Dim strResult As String

    strResult = AllowInsert
    If strResult <> "" Then
        MsgBox strResult, vbOKOnly + vbInformation, Me.Caption
        Exit Sub
    End If
    
    '获取插队序号
    strNewOrder = GetInsertOrder(mlngSelRow)
    
    '调用插队方法
    If mobjQueueManage.ChangeOrder(mlngQueueId, strNewOrder, cbxMemo) Then
        mblnOk = True
    
        '完成后卸载窗体
        Unload Me
    End If

    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
'    Resume
End Sub

Private Sub Form_Load()
    Call InitInsertQueueList
    Call LoadInsertReason
End Sub


Private Sub LoadInsertReason()
'载入插队原因
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim aryReason() As String
    Dim i As Long
    
    cbxMemo.Clear
    
    If mstrReason = "" Then
        strSql = "Select 名称 from 排队优先原因 order by 名称"
        Set rsData = zlDatabase.OpenSQLRecord(strSql, "查询排队优先原因")
        
        If rsData.RecordCount <= 0 Then Exit Sub
        
        While Not rsData.EOF
            cbxMemo.AddItem Nvl(rsData!名称)
            rsData.MoveNext
        Wend
    Else
        aryReason = Split(mstrReason & ",", ",")
        
        For i = 0 To UBound(aryReason)
            If Trim(aryReason(i)) <> "" Then
                cbxMemo.AddItem aryReason(i)
            End If
        Next i
    End If
End Sub
