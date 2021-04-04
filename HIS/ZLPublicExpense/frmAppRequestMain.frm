VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Begin VB.Form frmAppRequestMain 
   BorderStyle     =   0  'None
   Caption         =   "frmAppRequestMain"
   ClientHeight    =   5550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   7020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin XtremeReportControl.ReportControl rptMain 
      Height          =   3765
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   5535
      _Version        =   589884
      _ExtentX        =   9763
      _ExtentY        =   6641
      _StockProps     =   0
      BorderStyle     =   1
      PreviewMode     =   -1  'True
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
   End
End
Attribute VB_Name = "frmAppRequestMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mstr登记人 As String
Public mdat开始时间 As Date
Public mdat结束时间 As Date
Public mdat处理开始 As Date
Public mdat处理结束 As Date
Public mstr处理人 As String
Public mbln显示处理 As Boolean
Public mbyt复诊方式 As Byte
Public mbln登记时间 As Boolean
Public mbln处理时间 As Boolean

Private Sub Form_Load()
    Dim objCol As ReportColumn
    Dim objRecord As ReportRecord
    Dim ObjItem As ReportRecordItem
    With rptMain
        .AutoColumnSizing = False '不使用自动列宽
        .AllowColumnRemove = False '不允许拖动删除列
        .ShowGroupBox = True '显示分组框
        .ShowItemsInGroups = False '不显示已分组的列
        .MultipleSelection = False '不允许多行选择
        .PreviewMode = False
        .AllowColumnReorder = False

        With .PaintManager
            .HighlightBackColor = 16772055
            .HighlightForeColor = RGB(0, 0, 0)
            .MaxPreviewLines = 1
            .ColumnStyle = xtpColumnFlat
            .GridLineColor = RGB(180, 180, 180)
            .VerticalGridStyle = xtpGridSolid
            .HorizontalGridStyle = xtpGridSolid '横向表格线格式
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .GroupBoxBackColor = RGB(180, 180, 180)
            .NoItemsText = ""
        End With
        
        With .Columns
            Set objCol = .Add(0, "号码", 40, True)
            objCol.Alignment = xtpAlignmentCenter
            Set objCol = .Add(1, "号类", 60, True)
            objCol.Alignment = xtpAlignmentCenter
            Set objCol = .Add(2, "科室", 100, True)
            objCol.Alignment = xtpAlignmentCenter
            Set objCol = .Add(3, "项目", 200, True)
            objCol.Alignment = xtpAlignmentLeft
            Set objCol = .Add(4, "复诊方式", 100, True)
            objCol.Alignment = xtpAlignmentLeft
            rptMain.GroupsOrder.Add objCol
            objCol.Visible = False
            Set objCol = .Add(5, "复诊原因", 200, True)
            objCol.Alignment = xtpAlignmentLeft
            Set objCol = .Add(6, "开始日期", 150, True)
            objCol.Alignment = xtpAlignmentLeft
            Set objCol = .Add(7, "终止日期", 150, True)
            objCol.Alignment = xtpAlignmentLeft
            Set objCol = .Add(8, "病人姓名", 100, True)
            objCol.Alignment = xtpAlignmentLeft
            Set objCol = .Add(9, "登记人", 100, True)
            objCol.Alignment = xtpAlignmentLeft
            Set objCol = .Add(10, "登记时间", 200, True)
            objCol.Alignment = xtpAlignmentLeft
            Set objCol = .Add(11, "处理人", 100, True)
            objCol.Alignment = xtpAlignmentLeft
            Set objCol = .Add(12, "处理时间", 200, True)
            objCol.Alignment = xtpAlignmentLeft
        End With
    End With
    Call LoadRecord(True)
End Sub

Private Sub Form_Resize()
    With rptMain
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight
    End With
End Sub

Private Sub LoadRecord(Optional ByVal blnFirst As Boolean)
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim i As Integer
    Dim objCol As ReportColumn
    Dim objRecord As ReportRecord
    Dim ObjItem As ReportRecordItem
    strSQL = "Select a.id As 消息id, a.号码, d.号类, c.名称 As 项目, e.名称 As 科室, a.医生姓名 As 挂号医生, a.开始时间, a.终止时间, b.姓名 As 病人姓名, b.门诊号, b.性别, b.年龄, b.家庭电话 As 联系电话, a.登记人," & vbNewLine & _
            "       a.登记时间, a.通知原因 As 登记原因, a.处理时间, a.处理人, a.处理说明, a.复诊方式 , a.数量 " & vbNewLine & _
            "From 病人服务信息记录 A, 病人信息 B, 收费项目目录 C, 临床出诊号源 D, 部门表 E " & vbNewLine & _
            "Where a.项目id = c.Id And a.病人id = b.病人id And a.通知类型 = 3 And a.号源id = d.id And d.科室id = e.id "
    If blnFirst Then
        strSQL = strSQL & "And a.登记人 = [1] And a.登记时间 Between Trunc(Sysdate) And Trunc(Sysdate+1)-1/24/60/60 And a.处理时间 Is Null"
        mstr登记人 = UserInfo.姓名
    Else
        If mstr登记人 <> "" Then strSQL = strSQL & " And a.登记人=[1]"
        If mbln登记时间 Then
            strSQL = strSQL & " And a.登记时间 Between [2] And [3]"
        End If
        If mbln显示处理 = False Then
            strSQL = strSQL & " And a.处理时间 Is Null"
        Else
            If mstr处理人 <> "" Then strSQL = strSQL & " And (a.处理人=[4] Or a.处理人 Is Null)"
            If mbln处理时间 Then
                strSQL = strSQL & " And (a.处理时间 Between [5] And [6] Or a.处理时间 Is Null)"
            End If
        End If
        If mbyt复诊方式 <> 0 Then strSQL = strSQL & " And a.复诊方式 = [7]"
    End If
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr登记人, mdat开始时间, mdat结束时间, mstr处理人, mdat处理开始, mdat处理结束, mbyt复诊方式)
    With rptMain
        .Records.DeleteAll
        Do While Not rsTemp.EOF
            Set objRecord = .Records.Add
            objRecord.Tag = Val(Nvl(rsTemp!消息ID))
            Set ObjItem = objRecord.AddItem(Nvl(rsTemp!号码))
            Set ObjItem = objRecord.AddItem(Nvl(rsTemp!号类))
            Set ObjItem = objRecord.AddItem(Nvl(rsTemp!科室))
            Set ObjItem = objRecord.AddItem(Nvl(rsTemp!项目))
            Select Case Val(rsTemp!复诊方式)
            Case 1
                Set ObjItem = objRecord.AddItem(rsTemp!数量 & "个疗程后复诊")
            Case 2
                Set ObjItem = objRecord.AddItem(rsTemp!数量 & "个月后复诊")
            Case 3
                Set ObjItem = objRecord.AddItem(rsTemp!数量 & "周后复诊")
            Case 4
                Set ObjItem = objRecord.AddItem(rsTemp!数量 & "天后复诊")
            End Select
            Set ObjItem = objRecord.AddItem(Nvl(rsTemp!登记原因))
            Set ObjItem = objRecord.AddItem(Format(rsTemp!开始时间, "yyyy-mm-dd"))
            Set ObjItem = objRecord.AddItem(Format(rsTemp!终止时间, "yyyy-mm-dd"))
            Set ObjItem = objRecord.AddItem(Nvl(rsTemp!病人姓名))
            Set ObjItem = objRecord.AddItem(Nvl(rsTemp!登记人))
            Set ObjItem = objRecord.AddItem(Format(rsTemp!登记时间, "yyyy-MM-dd hh:mm:ss"))
            Set ObjItem = objRecord.AddItem(Nvl(rsTemp!处理人))
            Set ObjItem = objRecord.AddItem(Nvl(rsTemp!处理时间))
            
            If Not IsNull(rsTemp!处理时间) Then
                For i = 0 To 12
                    objRecord.Item(i).ForeColor = vbBlue
                Next i
            End If
            rsTemp.MoveNext
        Loop
        .Populate
    End With
End Sub

Public Sub RefreshData()
    Call LoadRecord(False)
End Sub

Private Sub rptMain_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If rptMain.SelectedRows.Count = 0 Then Exit Sub
    If rptMain.SelectedRows.Row(0).Record Is Nothing Then Exit Sub
    Call frmAppRequestEdit.ReadBill(Me, Val(rptMain.SelectedRows.Row(0).Record.Tag))
End Sub
