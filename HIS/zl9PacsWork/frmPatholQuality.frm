VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmPatholQuality 
   Caption         =   "病理质量管理"
   ClientHeight    =   5895
   ClientLeft      =   75
   ClientTop       =   405
   ClientWidth     =   10110
   Icon            =   "frmPatholQuality.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5895
   ScaleWidth      =   10110
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame framZhengti 
      Height          =   1575
      Left            =   120
      TabIndex        =   5
      Top             =   3480
      Width           =   9855
      Begin RichTextLib.RichTextBox rtfAdvice 
         Height          =   855
         Left            =   1080
         TabIndex        =   9
         Top             =   600
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   1508
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         ScrollBars      =   2
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmPatholQuality.frx":179A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.ComboBox cbxQuality 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmPatholQuality.frx":1837
         Left            =   1080
         List            =   "frmPatholQuality.frx":1839
         TabIndex        =   6
         Top             =   200
         Width           =   2055
      End
      Begin VB.Label labAdvice 
         Caption         =   "评价意见："
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   975
      End
      Begin VB.Label labFuhe 
         Caption         =   "符合情况："
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.PictureBox picControl 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   9855
      TabIndex        =   4
      Top             =   5160
      Width           =   9855
      Begin VB.CommandButton cmdSave 
         Caption         =   "数据保存(&S)"
         Height          =   400
         Left            =   8520
         TabIndex        =   2
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "删除记录(&D)"
         Height          =   400
         Left            =   7200
         TabIndex        =   3
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.Frame framRecord 
      Caption         =   "评价记录"
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9855
      Begin zl9PACSWork.ucFlexGrid ufgData 
         Height          =   2775
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   4895
         GridRows        =   21
         IsCopyAdoMode   =   0   'False
         IsEjectConfig   =   -1  'True
         HeadFontCharset =   134
         HeadFontWeight  =   400
         DataFontCharset =   134
         DataFontWeight  =   400
      End
   End
End
Attribute VB_Name = "frmPatholQuality"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private mlngCurAdviceId As Long
Private mstrPrivs As String
Private mblnMoved As Boolean
Private mlngCurDepartmentId As Long


Private mrecStudyInf As TStudyStateInf


Public blnIsOk As Boolean
Public strQuality As String


Public Sub zlRefresh(lngAdviceID As Long, ByVal blnReadOnly As Boolean, _
    strPrivs As String, ByVal blnMoved As Boolean, _
    ByVal lngCurDepartmentId As Long, Optional owner As Form = Nothing)
    
    If lngAdviceID <= 0 Then
        Call ConfigQualityFace(False, "医嘱ID无效请检查。")
        Exit Sub
    End If
    
'    If mlngCurAdviceId = lngAdviceId Then Exit Sub

    mlngCurAdviceId = lngAdviceID
    mstrPrivs = strPrivs
    mblnMoved = blnMoved
    mlngCurDepartmentId = lngCurDepartmentId
    
    blnIsOk = False
    strQuality = ""


    Call GetPatholStudyState(lngAdviceID, mrecStudyInf)

    If Trim(mrecStudyInf.strPatholNumber) = "" Then
        Call ConfigQualityFace(False, "该检查尚未生成有效的病理号，请确认该检查是否已被核收。")

        If Not (owner Is Nothing) Then
            Call MsgBoxD(Me, "该检查尚未生成有效的病理号，请确认该检查是否已被核收。", vbOKOnly, Me.Caption)
        End If

        Exit Sub
    Else
        Call ConfigQualityFace(True)
    End If

    Call LoadQualityData(mrecStudyInf.lngPatholAdviceId)
    Call LoadEnsembleQuality(mrecStudyInf.lngPatholAdviceId)

    Call ConfigPopedom(blnReadOnly)
    
    If Not (owner Is Nothing) Then
        Call Me.Show(1, owner)
    End If
End Sub



Private Sub ConfigPopedom(ByVal blnIsReadOnly As Boolean)
'配置权限
    Dim blnIsAllowDelay As Boolean
    
    blnIsAllowDelay = CheckPopedom(mstrPrivs, "质量管理")
    
    cmdDelete.Enabled = blnIsAllowDelay And Not blnIsReadOnly
    cmdSave.Enabled = blnIsAllowDelay And Not blnIsReadOnly
    
    cbxQuality.Enabled = Not blnIsReadOnly
    rtfAdvice.Enabled = Not blnIsReadOnly
    
    ufgData.ReadOnly = blnIsReadOnly
End Sub



Private Sub ConfigQualityFace(ByVal blnIsValid As Boolean, Optional ByVal strHintInf As String = "")
'配置质量管理界面
    cmdSave.Enabled = blnIsValid
    cmdDelete.Enabled = blnIsValid
    
    cbxQuality.Enabled = blnIsValid
    rtfAdvice.Enabled = blnIsValid
    
    labFuhe.Enabled = blnIsValid
    labAdvice.Enabled = blnIsValid
    
    If blnIsValid Then
        Call ufgData.CloseHintInf
    Else
        Call ufgData.ShowHintInf(strHintInf)
    End If
End Sub




Private Sub InitQualityDataList()
'初始化报告延迟显示列表
    Dim strTemp As String
    

     
    strTemp = zlDatabase.GetPara("检查质量列表配置", glngSys, G_LNG_PATHOLSYS_NUM, "")
    ufgData.DefaultColNames = gstrPatholQualityCols
     
    If strTemp = "" Then
        ufgData.ColNames = gstrPatholQualityCols
    Else
        ufgData.ColNames = strTemp
    End If
    '设置行数
    ufgData.GridRows = glngStandardRowCount
    '设置行高
    ufgData.RowHeightMin = glngStandardRowHeight
    ufgData.ColConvertFormat = gstrPatholQualityConvertFormat
End Sub

Private Sub ufgData_OnColFormartChange()
    '关闭窗口时保存列表配置
     zlDatabase.SetPara "检查质量列表配置", ufgData.GetColsString(ufgData), glngSys, G_LNG_PATHOLSYS_NUM
End Sub


Private Sub ConfigCbxQuality()
On Error Resume Next
    Dim strQuality As String
    Dim aryQuality() As String
    Dim i As Long
    
'    strQuality = ufgData.DataGrid.ColComboList(ufgData.vfgHelper.GetColumnIndex(gstrPatholQuality_评价结果))
'
'    aryQuality = Split(strQuality, "|")
'
'    Call cbxQuality.Clear
'    Call cbxQuality.AddItem("")
'
'    For i = LBound(aryQuality) To UBound(aryQuality)
'        Call cbxQuality.AddItem(aryQuality(i))
'    Next i

    Call cbxQuality.Clear
    
    cbxQuality.AddItem "符合"
    cbxQuality.AddItem "基本符合"
    cbxQuality.AddItem "不符合"
    
    cbxQuality.Text = "符合"
End Sub


Private Sub LoadQualityData(ByVal lngPatholAdviceId As Long)
'载入质量数据
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    strSql = "select ID,评价项目,评价结果,评价意见,改进方法,备注,评价人,评价时间 from 病理质量信息 where 病理医嘱ID=[1] order by 评价时间"
'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    Set ufgData.AdoData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPatholAdviceId)
    
    Call ufgData.RefreshData
End Sub


Private Sub LoadEnsembleQuality(ByVal lngPatholAdviceId As Long)
'载入病理检查综合质量
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    strSql = "select 综合质量,综合意见　from 病理检查信息 where 病理医嘱ID=[1]"
'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPatholAdviceId)

    If rsData.RecordCount < 0 Then Exit Sub
    
    cbxQuality.Text = IIf(Nvl(rsData!综合质量) = "", "符合", Nvl(rsData!综合质量))
    rtfAdvice.Text = Nvl(rsData!综合意见)
End Sub



Private Sub AdjustFace()
'调整界面布局
    framRecord.Left = 120
    framRecord.Top = 120
    framRecord.Width = Me.Width - 360
    framRecord.Height = Me.Height - picControl.Height - framZhengti.Height - 900
    
    ufgData.Left = 120
    ufgData.Top = 240
    ufgData.Width = framRecord.Width - 240
    ufgData.Height = framRecord.Height - 360
    
    
    framZhengti.Left = 120
    framZhengti.Top = framRecord.Top + framRecord.Height + 120
    framZhengti.Width = Me.Width - 360
    
    labFuhe.Left = 120
    labAdvice.Left = 120
    cbxQuality.Left = labFuhe.Left + labFuhe.Width + 120
    rtfAdvice.Left = labAdvice.Left + labAdvice.Width + 120
    
    rtfAdvice.Width = framZhengti.Width - labAdvice.Width - 360
    
    
    
    picControl.Left = 120
    picControl.Top = framZhengti.Top + framZhengti.Height + 120
    picControl.Width = Me.Width - 360
    

    
    cmdSave.Left = picControl.Width - cmdSave.Width
    cmdSave.Top = 0
    
    cmdDelete.Left = cmdSave.Left - cmdDelete.Width - 120
    cmdDelete.Top = 0
    
End Sub


Private Sub cmdDelete_Click()
'删除质量评价
On Error GoTo errHandle
    If ufgData.ShowingRowCount <= 0 Then Exit Sub
    
    If Not ufgData.IsSelectionRow Then
        Call MsgBoxD(Me, "请选择需要进行删除的病理质量记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If MsgBoxD(Me, "确认要删除该病理质量数据吗？", vbYesNo, Me.Caption) = vbNo Then Exit Sub
    
    Call ufgData.DelCurRow
    
    '保存删除的数据
    Call SaveQualityData(True)
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdSave_Click()
On Error GoTo errHandle
    Dim blnValid As Boolean
    
'    If ufgData.ShowDataRows <= 0 Then
'        Call MsgBoxD(Me, "请录入需要保存的质量评价信息。", vbOKOnly, Me.Caption)
'        Exit Sub
'    End If
    
    blnValid = Not ufgData.IsErrColorWithList
    If Not blnValid Then
        Call MsgBoxD(Me, "检测到病理质量列表中存在无效数据，请确认相关数据是否正确完整的录入，“红色”标记的单元格为必录数据。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    '保存套餐信息
    Call SaveQualityData
    
    
    blnIsOk = True
    strQuality = cbxQuality.Text
    
    Call Me.Hide
    'Call MsgBoxD(Me, "数据已保存成功。", vbOKOnly, Me.Caption)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub SaveQualityData(Optional ByVal blnIsSaveOnlyDel As Boolean = False)
'blnIsSaveOnlyDel:是否保存仅删除的数据

'保存报告延迟数据
    Dim i As Long
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim dtServicesTime As Date
    
    For i = 1 To ufgData.GridRows - 1
        Select Case ufgData.RowState(i)
            Case IIf(blnIsSaveOnlyDel, -1, TDataRowState.Add)
                dtServicesTime = zlDatabase.Currentdate
                
                strSql = "select Zl_病理质量_新增([1],[2],[3],[4],[5],[6],[7],[8]) as 返回值 from dual"
                Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, _
                                                        mrecStudyInf.lngPatholAdviceId, _
                                                        ufgData.Text(i, gstrPatholQuality_评价项目), _
                                                        ufgData.Text(i, gstrPatholQuality_评价结果), _
                                                        ufgData.Text(i, gstrPatholQuality_评价意见), _
                                                        ufgData.Text(i, gstrPatholQuality_改进方法), _
                                                        ufgData.Text(i, gstrPatholQuality_备注), _
                                                        UserInfo.姓名, _
                                                        CDate(dtServicesTime))
                                                        
                If rsData.RecordCount <= 0 Then
                    Call err.Raise(0, "SaveReportDelayData", "未成功获取新增后的病理质量ID,处理失败。")
                    Exit Sub
                End If
                
                
                ufgData.Text(i, gstrPatholQuality_ID) = rsData!返回值
                ufgData.Text(i, gstrPatholQuality_评价人) = UserInfo.姓名
                ufgData.Text(i, gstrpatholQuality_评价时间) = dtServicesTime
                                                        
            Case TDataRowState.Del
                strSql = "Zl_病理质量_删除(" & Val(ufgData.KeyValue(i)) & ")"
                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
            Case IIf(blnIsSaveOnlyDel, -1, TDataRowState.Update)
                strSql = "Zl_病理质量_更新(" & Val(ufgData.KeyValue(i)) & ",'" & _
                                                ufgData.Text(i, gstrPatholQuality_评价项目) & "','" & _
                                                ufgData.Text(i, gstrPatholQuality_评价结果) & "','" & _
                                                ufgData.Text(i, gstrPatholQuality_评价意见) & "','" & _
                                                ufgData.Text(i, gstrPatholQuality_改进方法) & "','" & _
                                                ufgData.Text(i, gstrPatholQuality_备注) & "')"
                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)

        End Select
        
        
        '更新行状态
        ufgData.RowState(i) = TDataRowState.Normal
    Next i
    
    '保存综合质量
    strSql = "Zl_病理质量_综合('" & mrecStudyInf.lngPatholAdviceId & "','" & cbxQuality.Text & "','" & rtfAdvice.Text & "')"
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
End Sub


Private Sub Form_Load()
On Error GoTo errHandle
    Call RestoreWinState(Me, App.ProductName)
    
    Call InitQualityDataList
    
    Call ConfigCbxQuality
    
    blnIsOk = False
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Form_Resize()
On Error Resume Next
    Call AdjustFace
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
End Sub


Private Sub ufgData_OnAfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim iCol As Long
    Dim i As Long
    
    If ufgData.IsNullRow(Row) Then
        ufgData.RowState(Row) = TDataRowState.Normal
        Call ufgData.SetRowColor(Row, ufgData.BackColor)
        
        Exit Sub
    End If
        
    
    '如果未录入评价项目，则显示淡红色
    iCol = ufgData.GetColIndex(gstrPatholQuality_评价项目)
    
    ufgData.CellColor(Row, iCol) = IIf(ufgData.Text(Row, gstrPatholQuality_评价项目) = "", ufgData.ErrCellColor, ufgData.BackColor)
           
    
    
    '如果未录入评价结果，则显示淡红色
    iCol = ufgData.GetColIndex(gstrPatholQuality_评价结果)
    
    ufgData.CellColor(Row, iCol) = IIf(ufgData.Text(Row, gstrPatholQuality_评价结果) = "", ufgData.ErrCellColor, ufgData.BackColor)
End Sub


Private Sub ufgData_OnColsNameReSet()
On Error GoTo errHandle
    Call LoadQualityData(mrecStudyInf.lngPatholAdviceId)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub
