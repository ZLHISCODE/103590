VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Begin VB.Form frmServiceMessage 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   8550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin XtremeReportControl.ReportControl rptMain 
      Height          =   5475
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      _Version        =   589884
      _ExtentX        =   11880
      _ExtentY        =   9657
      _StockProps     =   0
      BorderStyle     =   1
      PreviewMode     =   -1  'True
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   2190
      Top             =   5670
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceMessage.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceMessage.frx":059A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceMessage.frx":0B34
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceMessage.frx":10CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceMessage.frx":1668
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceMessage.frx":1C02
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmServiceMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmMain As Object
Public mdatBegin As Date, mdatEnd As Date
Public mstr登记人 As String, mstr消息类型 As String
Public mblnShowRead As Boolean, mblnFilter As Boolean

Public Sub ShowMe(frmMain As Object)
    Set mfrmMain = frmMain
End Sub

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
        .SetImageList img16
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
            Set objCol = .Add(0, "", 20, False)
            objCol.Groupable = False
            objCol.Alignment = xtpAlignmentCenter
            Set objCol = .Add(1, "消息类型", 60, True)
            objCol.Alignment = xtpAlignmentCenter
            rptMain.GroupsOrder.Add objCol
            objCol.Visible = False
            Set objCol = .Add(2, "主题", 200, True)
            objCol.Alignment = xtpAlignmentLeft
            Set objCol = .Add(3, "通知原因", 200, True)
            objCol.Alignment = xtpAlignmentLeft
            Set objCol = .Add(4, "登记人", 60, True)
            objCol.Alignment = xtpAlignmentCenter
            Set objCol = .Add(5, "登记时间", 100, True)
            objCol.Alignment = xtpAlignmentCenter
            Set objCol = .Add(6, "ID", 0, False)
            objCol.Alignment = xtpAlignmentCenter
            objCol.Visible = False
        End With
        
    End With
    Call LoadMessage(True)
End Sub

Public Sub LoadMessage(blnFirst As Boolean)
    Dim objCol As ReportColumn
    Dim objRecord As ReportRecord
    Dim ObjItem As ReportRecordItem
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim datNow As Date, strTemp As String
    datNow = zlDatabase.Currentdate
'    If blnFirst Then
'        strTemp = " And [1] Between a.开始时间 And a.终止时间 "
'    Else
'        If mblnShowRead = False Then
'            strTemp = " And [1] Between a.开始时间 And a.终止时间 "
'        End If
'    End If
    With rptMain
        .Records.DeleteAll
        strSQL = "Select ID, 通知类型, 号码, 部门, 医生姓名, 项目名称, 通知原因, 登记人, 登记时间, 病人姓名, 开始时间, 终止时间, 替诊医生姓名, 停诊开始时间, 处理时间" & vbNewLine & _
                "From (Select a.Id, a.通知类型, b.号码, c.名称 As 部门, Nvl(d.姓名, b.医生姓名) As 医生姓名, e.名称 As 项目名称, a.通知原因, a.登记人, a.登记时间, f.姓名 As 病人姓名," & vbNewLine & _
                "              a.开始时间, a.终止时间, Null As 替诊医生姓名, Null As 停诊开始时间, a.处理时间" & vbNewLine & _
                "       From 病人服务信息记录 A, 临床出诊号源 B, 部门表 C, 人员表 D, 收费项目目录 E, 病人信息 F" & vbNewLine & _
                "       Where a.号源id = b.Id And b.科室id = c.Id And b.医生id = d.Id(+) And" & vbNewLine & _
                "             b.项目id = e.Id And a.病人id = f.病人id And a.通知类型 = 3  " & IIf(blnFirst, " And Sysdate Between a.开始时间 And a.终止时间 ", " And (a.开始时间 Between [3] And [4] Or a.终止时间 Between [3] And [4]) ") & vbNewLine & _
                "       Union" & vbNewLine & _
                "       Select a.Id, a.通知类型, g.号码, c.名称 As 部门, Nvl(d.姓名, b.医生姓名) As 医生姓名, e.名称 As 项目名称, a.通知原因, a.登记人, a.登记时间, Null As 病人姓名," & vbNewLine & _
                "              a.开始时间, a.终止时间, Nvl(Ds.姓名, b.替诊医生姓名) As 替诊医生姓名, b.停诊开始时间," & vbNewLine & _
                "              Case" & vbNewLine & _
                "                When a.处理时间 < Sysdate - 9998 Then" & vbNewLine & _
                "                 Null" & vbNewLine & _
                "                Else" & vbNewLine & _
                "                 a.处理时间" & vbNewLine & _
                "              End As 处理时间" & vbNewLine & _
                "       From (Select Min(ID) As ID, 通知类型, 记录id, Min(登记人) As 登记人, Min(登记时间) As 登记时间, Min(病人id) As 病人id, Min(通知原因) As 通知原因," & vbNewLine & _
                "                     Min(Nvl(处理时间, Sysdate - 9999)) As 处理时间, 开始时间, 终止时间" & vbNewLine & _
                "              From 病人服务信息记录" & vbNewLine & _
                "              Where 通知类型 In (1, 2)" & IIf(blnFirst, "", " And 登记时间 Between [3] And [4] ") & vbNewLine & _
                "              Group By 通知类型, 记录id, 开始时间, 终止时间) A, 临床出诊记录 B, 部门表 C, 人员表 D, 人员表 Ds, 收费项目目录 E, 临床出诊号源 G" & vbNewLine & _
                "       Where a.记录id = b.Id And g.科室id = c.Id And b.号源id = g.Id And b.医生id = d.Id(+) And b.替诊医生id = Ds.Id(+) And" & vbNewLine & _
                "             b.项目id = e.Id And a.通知类型 In (1, 2))" & vbNewLine & _
                "Where 1 = 1 "
        
        If blnFirst Then
            strSQL = strSQL & " And 处理时间 Is Null"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, datNow)
        Else
            If Not mblnShowRead Then strSQL = strSQL & " And 处理时间 Is Null"
            If mstr消息类型 <> "" Then
                strSQL = strSQL & " And 通知类型 In (Select Column_Value From Table(f_Str2list([2])))"
            End If
            If mstr登记人 <> "" Then strSQL = strSQL & " And 登记人 = [5]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, datNow, mstr消息类型, mdatBegin, mdatEnd, mstr登记人)
        End If
        Do While Not rsTemp.EOF
            Select Case Val(rsTemp!通知类型)
                Case 1
                    Set objRecord = .Records.Add()
                    Set ObjItem = objRecord.AddItem("")
                    If Not IsNull(rsTemp!处理时间) Then
                        ObjItem.Icon = 1
                    Else
                        ObjItem.Icon = 0
                    End If
                    Set ObjItem = objRecord.AddItem("医生停诊")
                    objRecord.AddItem Nvl(rsTemp!部门) & "," & Nvl(rsTemp!医生姓名) & _
                                "(" & Nvl(rsTemp!项目名称) & ")" & "将于" & Nvl(rsTemp!停诊开始时间) & "停诊"
                    
                    objRecord.AddItem Nvl(rsTemp!通知原因)
                    objRecord.AddItem Nvl(rsTemp!登记人)
                    objRecord.AddItem Nvl(rsTemp!登记时间)
                    objRecord.AddItem Val(Nvl(rsTemp!ID))
                    
                    If Not IsNull(rsTemp!处理时间) Then
                        objRecord.Item(1).ForeColor = vbBlue
                        objRecord.Item(2).ForeColor = vbBlue
                        objRecord.Item(3).ForeColor = vbBlue
                        objRecord.Item(4).ForeColor = vbBlue
                        objRecord.Item(5).ForeColor = vbBlue
                    End If
                Case 2
                    Set objRecord = .Records.Add()
                    Set ObjItem = objRecord.AddItem("")
                    If Not IsNull(rsTemp!处理时间) Then
                        ObjItem.Icon = 3
                    Else
                        ObjItem.Icon = 2
                    End If
                    Set ObjItem = objRecord.AddItem("医生替诊")
                    objRecord.AddItem Nvl(rsTemp!部门) & "," & Nvl(rsTemp!医生姓名) & _
                                "(" & Nvl(rsTemp!项目名称) & ")" & "将于" & Nvl(rsTemp!停诊开始时间) & "停诊" & _
                                ",由" & Nvl(rsTemp!替诊医生姓名) & "替诊"
                    objRecord.AddItem Nvl(rsTemp!通知原因)
                    objRecord.AddItem Nvl(rsTemp!登记人)
                    objRecord.AddItem Nvl(rsTemp!登记时间)
                    objRecord.AddItem Val(Nvl(rsTemp!ID))
                    If Not IsNull(rsTemp!处理时间) Then
                        objRecord.Item(1).ForeColor = vbBlue
                        objRecord.Item(2).ForeColor = vbBlue
                        objRecord.Item(3).ForeColor = vbBlue
                        objRecord.Item(4).ForeColor = vbBlue
                        objRecord.Item(5).ForeColor = vbBlue
                    End If
                Case 3
                    Set objRecord = .Records.Add()
                    Set ObjItem = objRecord.AddItem("")
                    If Not IsNull(rsTemp!处理时间) Then
                        ObjItem.Icon = 5
                    Else
                        ObjItem.Icon = 4
                    End If
                    Set ObjItem = objRecord.AddItem("预约登记")
                    objRecord.AddItem Nvl(rsTemp!病人姓名) & "病人需在" & Nvl(rsTemp!开始时间) & "至" & Nvl(rsTemp!终止时间) & "间复查"
                    objRecord.AddItem Nvl(rsTemp!通知原因)
                    objRecord.AddItem Nvl(rsTemp!登记人)
                    objRecord.AddItem Nvl(rsTemp!登记时间)
                    objRecord.AddItem Val(Nvl(rsTemp!ID))
                    If Not IsNull(rsTemp!处理时间) Then
                        objRecord.Item(1).ForeColor = vbBlue
                        objRecord.Item(2).ForeColor = vbBlue
                        objRecord.Item(3).ForeColor = vbBlue
                        objRecord.Item(4).ForeColor = vbBlue
                        objRecord.Item(5).ForeColor = vbBlue
                    End If
            End Select
            rsTemp.MoveNext
        Loop
        .Populate
        If .Rows.Count <> 0 Then
            .Rows.Row(0).Selected = True
            Call rptMain_SelectionChanged
        Else
            If Not mfrmMain Is Nothing Then Call mfrmMain.NoData
        End If
    End With
End Sub

Private Sub Form_Resize()
    rptMain.Width = Me.ScaleWidth
    rptMain.Height = Me.ScaleHeight
End Sub

Private Sub rptMain_SelectionChanged()
    With rptMain
        If .SelectedRows.Count = 0 Then Exit Sub
        If mfrmMain Is Nothing Then Exit Sub
        If .SelectedRows.Row(0).GroupRow = True Then Call mfrmMain.NoData: Exit Sub
        Call mfrmMain.LoadData(.SelectedRows.Row(0).Record.Item(6).Value)
    End With
End Sub
