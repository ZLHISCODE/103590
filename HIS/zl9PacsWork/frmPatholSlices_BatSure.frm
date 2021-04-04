VERSION 5.00
Begin VB.Form frmPatholSlices_BatSure 
   Caption         =   "批量确认"
   ClientHeight    =   7008
   ClientLeft      =   72
   ClientTop       =   408
   ClientWidth     =   11484
   Icon            =   "frmPatholSlices_BatSure.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7008
   ScaleWidth      =   11484
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame framSureRecord 
      Caption         =   "已确认记录："
      Height          =   4215
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Width           =   9855
      Begin zl9PACSWork.ucFlexGrid ufgData 
         Height          =   3615
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   9615
         _ExtentX        =   16955
         _ExtentY        =   6371
         DefaultCols     =   ""
         IsKeepRows      =   0   'False
         IsCopyAdoMode   =   0   'False
         IsEjectConfig   =   -1  'True
         HeadFontCharset =   134
         HeadFontWeight  =   400
         DataFontCharset =   134
         DataFontWeight  =   400
      End
   End
   Begin VB.PictureBox picControl 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   120
      ScaleHeight     =   612
      ScaleWidth      =   9612
      TabIndex        =   9
      Top             =   5640
      Width           =   9615
      Begin VB.CommandButton cmdBatSure 
         Caption         =   "开始确认(&S)"
         Height          =   400
         Left            =   7080
         TabIndex        =   4
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "退 出(&B)"
         Height          =   400
         Left            =   8400
         TabIndex        =   5
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label labRecordInf 
         Caption         =   "需制片总数：0    已确认总数：0"
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   120
         Width           =   3495
      End
   End
   Begin VB.Frame framFilter 
      Height          =   975
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   9855
      Begin VB.OptionButton optUserCodeBar 
         Caption         =   "使用条码号"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   0
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdSure 
         Caption         =   "确 定(&S)"
         Height          =   400
         Left            =   5760
         TabIndex        =   2
         Top             =   330
         Width           =   1215
      End
      Begin VB.TextBox txtSureCount 
         Height          =   375
         Left            =   4560
         TabIndex        =   1
         Text            =   "1"
         ToolTipText     =   "这里输入需求进行确认的玻片数量。"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtSureNum 
         Height          =   375
         Left            =   1080
         TabIndex        =   0
         ToolTipText     =   "当未启用条码时，可在这里直接输入“病理号”查找。"
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "玻片数量："
         Height          =   255
         Left            =   3600
         TabIndex        =   8
         Top             =   435
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "确认号码："
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   435
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmPatholSlices_BatSure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mufgParentGrid As ucFlexGrid
Private mstrCurPatholNum As String

Public blnIsOk As Boolean


Public Sub ShowSlicesSureWindow(ufgParentGrid As ucFlexGrid, ByVal strPatholNum As String, owner As Form)
'显示制片确认窗口
    Set mufgParentGrid = ufgParentGrid
    
    mstrCurPatholNum = strPatholNum
    blnIsOk = False
    
    Call Me.Show(1, owner)
End Sub



Private Sub RefreshSureCount()
'刷新确认的数量信息
    Dim i As Long
    Dim iNeedCount As Long
    Dim iSureCount As Long
    
    iNeedCount = 0
    iSureCount = 0
    
    For i = 1 To ufgData.GridRows - 1
        iNeedCount = iNeedCount + Val(ufgData.Text(i, gstrSlicesSure_需制片数))
        iSureCount = iSureCount + Val(ufgData.Text(i, gstrSlicesSure_已确认数))
    Next i
    
    labRecordInf.Caption = "需制片总数：" & iNeedCount & "    已确认总数：" & iSureCount
End Sub


Private Sub DecodeSureNum(ByVal strSureNum As String, ByRef strPatholNum As String, ByRef strSlicesId As String)
'分解确认号码
    Dim lngFindSplitChar As Long
    
    If optUserCodeBar.value Then
        strPatholNum = ""
        strSlicesId = Trim(strSureNum)
    Else
        strPatholNum = Trim(strSureNum)
        strSlicesId = ""
    End If
    
'    lngFindSplitChar = InStr(1, strSureNum, "-")
'
'    If lngFindSplitChar > 0 Then
'        strPatholNum = Mid(strSureNum, 1, lngFindSplitChar - 1)
'        strSlicesId = Mid(strSureNum, lngFindSplitChar + 1, 20)
'    Else
'        strPatholNum = strSureNum
'        strSlicesId = ""
'    End If
End Sub



Private Sub SureSlices(ByVal strSureNum As String)
'确认制片处理过程
'strSureNum号码格式为“病理号-制片号”
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim lngFindRow As Integer
    Dim strPatholNum As String
    Dim strSlicesId As String
    
    strPatholNum = ""
    strSlicesId = ""
    
    Call DecodeSureNum(strSureNum, strPatholNum, strSlicesId)
    
    lngFindRow = ufgData.FindRowIndex(strSlicesId, gstrSlicesSure_ID)
    If lngFindRow > 0 Then GoTo errFindSlices
    
    
    If Trim(strPatholNum) = "" Then
        strSql = "select 病理号 from 病理检查信息 where 病理医嘱ID = (select 病理医嘱ID from 病理制片信息 where ID=[1] and rownum = 1)"
        'If mblnMoved Then strSql = GetMovedDataSql(strSql)
        
        Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(strSlicesId))
        
        If rsData.RecordCount > 0 Then
            strPatholNum = Val(Nvl(rsData!病理号))
        End If
    End If
    
    
    If Trim(strPatholNum) = "" Then
        Call MsgBoxD(Me, "输入的号码无效，不能根据此号码找到对应的病理号。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    
    
    '如果从列表中找不到当前录入的病理号，则从数据库中读取相关信息并加载到列表中
    lngFindRow = ufgData.FindRowIndex(strPatholNum, ufgData.KeyName)
    If lngFindRow <= 0 Then

        strSql = "Select a.id, c.病理号, c.病理医嘱ID, e.姓名, 0 as 已确认数,c.检查类型, a.制片类型,a.制片方式,a.材块ID,b.序号, d.标本名称, " & _
                        " a.当前状态, case a.当前状态 when 2 then 0 else a.制片数 end as 需制片数" & _
                        " From 病理制片信息 A, 病理取材信息 B, 病理检查信息 C, 病理标本信息 D, 病人医嘱记录 E " & _
                        " Where a.材块id = b.材块id And b.病理医嘱ID = c.病理医嘱ID And c.医嘱id = e.Id And b.标本id = d.标本id and c.病理号=upper([1]) and a.当前状态<> 2" & _
                        " order by b.序号,a.id"
        'If mblnMoved Then strSql = GetMovedDataSql(strSql)
        
        '查询数据
        Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strPatholNum)
        
        If rsData.RecordCount > 0 Then
            Set ufgData.AdoData = rsData
            Call ufgData.RefreshData(False)
        End If
    End If
    
errFindSlices:
    '查找需要确认的制片记录
    lngFindRow = ufgData.FindRowIndex(strSlicesId, gstrSlicesSure_ID)
    If lngFindRow > 0 Then
        ufgData.Text(lngFindRow, gstrSlicesSure_已确认数) = ufgData.Text(lngFindRow, gstrSlicesSure_已确认数) + Val(txtSureCount.Text)
        
        If ufgData.Text(lngFindRow, gstrSlicesSure_已确认数) = ufgData.Text(lngFindRow, gstrSlicesSure_需制片数) Then
            ufgData.CellColor(lngFindRow, ufgData.GetColIndex(gstrSlicesSure_已确认数)) = &HC0FFC0
        End If
        
        Call ufgData.LocateRow(lngFindRow)
    End If
    
End Sub





Private Sub InitSureList()

    '设置行数
    ufgData.GridRows = glngStandardRowCount
    '设置行高
    ufgData.RowHeightMin = glngStandardRowHeight
    
    '初始化确认数据显示列表
    ufgData.ColConvertFormat = gstrSlicesSureConvertFormat
    ufgData.DefaultColNames = gstrSlicesSureColsWithMaterialNum
    ufgData.ColNames = gstrSlicesSureColsWithMaterialNum

End Sub




Private Sub AdjustFace()
'调整界面布局
    framFilter.Left = 120
    framFilter.Top = 120
'    framFilter.Width = Me.Width - 360
    
    
    
    framSureRecord.Left = 120
    framSureRecord.Top = framFilter.Top + framFilter.Height + 30
    framSureRecord.Width = Me.Width - 360
    framSureRecord.Height = Me.Height - framFilter.Height - picControl.Height - 680
    
    ufgData.Left = 120
    ufgData.Top = 240
    ufgData.Width = framSureRecord.Width - 240
    ufgData.Height = framSureRecord.Height - 360
    
    
    
    
    picControl.Left = 120
    picControl.Top = framSureRecord.Top + framSureRecord.Height + 120
    picControl.Width = Me.Width - 360
    
    
    cmdExit.Left = picControl.Width - cmdExit.Width
    cmdExit.Top = 0
    
    
    cmdBatSure.Left = cmdExit.Left - cmdBatSure.Width - 120
    cmdBatSure.Top = 0
    
    
    labRecordInf.Left = 0
    labRecordInf.Top = cmdBatSure.Top + 60
    
End Sub


Private Sub StartBatSure()
'开始批量确认
    Dim i As Long
    Dim strSql As String
    Dim dtServicesTime As Date
    Dim strSurePatholNum As String
    Dim lngRowCheck As Long
    Dim blnUpdateParentGrid As Boolean
    
    
    dtServicesTime = zlDatabase.Currentdate
    
    blnUpdateParentGrid = False
    
    lngRowCheck = ufgData.GetColIndexWithRowCheck()
    For i = 1 To ufgData.GridRows - 1
        If ufgData.GetCellCheckState(i, lngRowCheck) Then
            
            If strSurePatholNum <> ufgData.KeyValue(i) Then
                strSurePatholNum = ufgData.KeyValue(i)
                
                strSql = "Zl_病理制片_确认('" & strSurePatholNum & "'," & zlStr.To_Date(dtServicesTime) & ")"

                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
            End If
            
            If strSurePatholNum = mstrCurPatholNum Then
                blnUpdateParentGrid = True
            End If
            
            ufgData.Text(i, gstrSlicesSure_当前状态) = "已完成"
            ufgData.Text(i, gstrSlicesSure_确认状态) = "已确认"
        End If
    Next i
    
    '更新调用界面列表中的确认状态
    If blnUpdateParentGrid And Not (mufgParentGrid Is Nothing) Then
        For i = 1 To mufgParentGrid.GridRows - 1
            If mufgParentGrid.Text(i, gstrSlicesSure_当前状态) = "已接受" Then
                 mufgParentGrid.Text(i, gstrSlices_当前状态) = "已完成"
                 mufgParentGrid.Text(i, gstrSlices_制片时间) = dtServicesTime
            End If
        Next i
    End If
End Sub



Private Function IsAllowSure() As Boolean
'判断确认列表中，是否允许进行制片确认
    Dim i As Long
    Dim lngRowCheck As Long
    
    IsAllowSure = True
    
    lngRowCheck = ufgData.GetColIndexWithRowCheck()
    For i = 1 To ufgData.GridRows - 1
        If ufgData.GetCellCheckState(i, lngRowCheck) Then
            If ufgData.Text(i, gstrSlicesSure_需制片数) <> ufgData.Text(i, gstrSlicesSure_已确认数) Or _
                Val(ufgData.Text(i, gstrSlicesSure_需制片数)) <= 0 Or _
                Val(ufgData.Text(i, gstrSlicesSure_已确认数)) <= 0 Then
                
                Call ufgData.LocateRow(i)
                IsAllowSure = False
                
                Exit Function
            End If
        End If
    Next i

End Function


Private Sub cmdBatSure_Click()
'执行批量确认
On Error GoTo errHandle
    
    If Not ufgData.IsCheckedRow Then
        Call MsgBoxD(Me, "请选择需要确认的制片信息。", vbOKOnly, Me.Caption)
        Exit Sub
    End If

    If Not IsAllowSure() Then
        Call MsgBoxD(Me, "检测到要确认的制片数与实际的需制片数不一致或者数量有误，不能进行确认。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Call StartBatSure
    
    Call RefreshSureCount
    
    blnIsOk = True
    
    Call MsgBoxD(Me, "批量确认已处理完成。", vbOKOnly, Me.Caption)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdExit_Click()
    blnIsOk = False
    Call Me.Hide
End Sub

Private Sub cmdSure_Click()
'制片确认
On Error GoTo errHandle
    Call SureSlices(txtSureNum.Text)
    
    Call RefreshSureCount
    
Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub Form_Load()
On Error GoTo errHandle
    Dim strValue As String
    
    Call RestoreWinState(Me, App.ProductName)

    Call InitSureList
    
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



Private Sub txtSureCount_GotFocus()
On Error Resume Next
    
    txtSureCount.SelStart = 0
    txtSureCount.SelLength = Len(txtSureCount.Text)
End Sub

Private Sub txtSureCount_KeyPress(KeyAscii As Integer)
'如果按下回车，则进行确认
    If KeyAscii = 13 Then
        Call cmdSure_Click
    End If
End Sub


Private Sub txtSureNum_GotFocus()
On Error Resume Next
    
    txtSureNum.SelStart = 0
    txtSureNum.SelLength = Len(txtSureNum.Text)
End Sub

Private Sub txtSureNum_KeyPress(KeyAscii As Integer)
On Error GoTo errHandle
'    '临时进行磁卡读取调整
'
'    Dim blnCard As Boolean
'    Dim rsData As ADODB.Recordset
'
'     If KeyAscii = 13 Then
'
'        txtSureCount.SetFocus
'
'        Exit Sub
'    End If
'
'
'    If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
'
'    '判断是否为刷卡
'    blnCard = zlCommFun.InputIsCard(txtSureNum, KeyAscii, glngSys)
'    If blnCard And Len(txtSureNum.Text) = mbyt磁卡 - 1 And KeyAscii <> 8 Then
'
'        txtSureNum.Text = txtSureNum.Text & Chr(KeyAscii)
'        txtSureNum.SelStart = Len(txtSureNum.Text)
'
'        KeyAscii = 0
'
'        txtSureNum.SelStart = 0
'        txtSureNum.SelLength = Len(txtSureNum.Text)
'
'        Call SureSlices(txtSureNum.Text)
'        Call RefreshSureCount
'    End If

    '条码扫描确认
    If KeyAscii = 13 Then
    
        Call SureSlices(txtSureNum.Text)
        Call RefreshSureCount
       
        txtSureNum.SelStart = 0
        txtSureNum.SelLength = Len(txtSureNum.Text)
    
       Exit Sub
    End If
    
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub ufgData_OnAfterEdit(ByVal Row As Long, ByVal Col As Long)
    If ufgData.Text(Row, gstrSlicesSure_已确认数) = ufgData.Text(Row, gstrSlicesSure_需制片数) Then
        ufgData.CellColor(Row, ufgData.GetColIndex(gstrSlicesSure_已确认数)) = &HC0FFC0
    End If
    
    Call RefreshSureCount
End Sub

