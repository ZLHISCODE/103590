VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Begin VB.Form frmFind 
   Caption         =   "查找"
   ClientHeight    =   5415
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   8730
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   8730
   Begin XtremeReportControl.ReportControl rptQueueList 
      Height          =   3615
      Left            =   90
      TabIndex        =   3
      Tag             =   "0"
      Top             =   765
      Width           =   8310
      _Version        =   589884
      _ExtentX        =   14658
      _ExtentY        =   6376
      _StockProps     =   0
      BorderStyle     =   3
      AllowColumnSort =   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin VB.Timer timerCard 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   1125
      Top             =   4455
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   0
      ScaleHeight     =   630
      ScaleWidth      =   8730
      TabIndex        =   8
      Top             =   4785
      Width           =   8730
      Begin VB.CommandButton cmdExit 
         Caption         =   "退出(&Q)"
         Height          =   400
         Index           =   0
         Left            =   7020
         TabIndex        =   6
         Top             =   90
         Width           =   1380
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "自定义(&C)"
         Height          =   400
         Index           =   1
         Left            =   3930
         TabIndex        =   4
         Top             =   90
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "恢复(&R)"
         Height          =   400
         Index           =   2
         Left            =   5475
         TabIndex        =   5
         Top             =   90
         Width           =   1380
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   8730
      TabIndex        =   7
      Top             =   0
      Width           =   8730
      Begin VB.ComboBox cboFindWay 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   300
         ItemData        =   "frmFind.frx":000C
         Left            =   105
         List            =   "frmFind.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   165
         Width           =   2115
      End
      Begin VB.TextBox txtFindData 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2235
         TabIndex        =   0
         Top             =   165
         Width           =   2655
      End
      Begin VB.CommandButton cmdStartFind 
         Caption         =   "开始查找(&F)"
         Height          =   405
         Left            =   4965
         TabIndex        =   1
         Top             =   105
         Width           =   1770
      End
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnOk As Boolean


Private mobjOwner As UcQueue
Private mlngReadCount As Long
Private mlngStartTime As Long
Private mlngAvgTime As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long


Public Event OnCustomFindButton(ByVal lngQueueId As Long)

Public Event OnFind(ByVal strFindWay As String, ByVal strFindValue As String, rsData As ADODB.Recordset, _
    txtFind As TextBox, ByVal blnUseCustom As Boolean)
    
Public Event OnReadBefore(rsData As ADODB.Recordset, ByVal lngListType As Long, blnCancel As Boolean)
Public Event OnReadAfter(rsData As ADODB.Recordset, ByVal lngListType As Long, objReportRecord As Object)

Public Function ShowFind(owner As UcQueue) As Boolean
    
    Set mobjOwner = owner
    
    Call SetFont(mobjOwner.Font)
    
    Call ConfigFindWay(owner.FindWay)
    
    Call InitQueueList(rptQueueList, mobjOwner.GroupName, mobjOwner.CustomOrderField, mobjOwner.DisplayQueueFields, UCase(mobjOwner.QueueKernel.DefQueryCols))
'    Call CopyCols(objShowReportCols, rptQueueList)

    
    Me.Show 1, owner
End Function

'Private Sub CopyCols(objSourceRC As ReportColumns, objTargetRC As ReportControl)
'    Dim i As Long
'    Dim Column As ReportColumn
'
'    '配置显示列
'    objTargetRC.Columns.DeleteAll
'
'    For i = 0 To objSourceRC.Count - 1
'        Set Column = objTargetRC.Columns.Add(i, objSourceRC(i).Caption, objSourceRC(i).Width, True)
'        Column.Groupable = objSourceRC(i).Groupable
'        Column.Visible = objSourceRC(i).Visible
'    Next i
'
'
'
'    objTargetRC.Populate
'
'End Sub

'Private Sub InitInsertQueueList()
'    '初始化排队队列显示字段
'    Call rptQueueList.Columns.DeleteAll
'
'    Set rptQueueList.Icons = zlCommFun.GetPubIcons
'
'    '初始化列表相关属性
'    rptQueueList.AllowColumnRemove = False
'    rptQueueList.ShowItemsInGroups = False
'    rptQueueList.SkipGroupsFocus = True
'    rptQueueList.MultipleSelection = False
'
'    With rptQueueList.PaintManager
'        .ColumnStyle = xtpColumnShaded
'        .GridLineColor = RGB(225, 225, 225)
'        .NoGroupByText = "将列标题拖动到此,可按该列分组..."
'        .NoItemsText = "没有可显示的项目..."
'        .VerticalGridStyle = xtpGridSolid
'    End With
'
'    rptQueueList.AllowColumnSort = False
'End Sub


Private Sub ConfigCustomButton(ByVal strCustomButtonCaption As String)
    If strCustomButtonCaption = "" Then
        Me.cmdExit(1).Visible = False
    Else
        Me.cmdExit(1).Caption = strCustomButtonCaption
        Me.cmdExit(1).Visible = True
    End If
End Sub

Private Sub ConfigFindWay(ByVal strFindWay As String)
    Dim aryFindWays() As String
    Dim i As Long
    Dim lngSelectIndex As Long
    
    lngSelectIndex = cboFindWay.ListIndex
    
    cboFindWay.Clear
    
    cboFindWay.AddItem "排队号"
    cboFindWay.AddItem "姓名"
        
    If strFindWay <> "" Then
        aryFindWays = Split(strFindWay, ",")
        For i = LBound(aryFindWays) To UBound(aryFindWays)
            If aryFindWays(i) <> "" And aryFindWays(i) <> "排队号" And aryFindWays(i) <> "姓名" Then
                cboFindWay.AddItem aryFindWays(i)
            End If
        Next i
    End If
    
    If lngSelectIndex < cboFindWay.ListCount Then cboFindWay.ListIndex = lngSelectIndex
End Sub


Private Sub SetFont(ft As StdFont)
    Dim dbTextHeight As Single
    
    Set Me.Font = ft
    Set cboFindWay.Font = ft
    Set cmdStartFind.Font = ft
    Set txtFindData.Font = ft
    
    Set rptQueueList.PaintManager.CaptionFont = ft
    Set rptQueueList.PaintManager.TextFont = ft
    
    Set cmdExit(0).Font = ft
    Set cmdExit(1).Font = ft
    Set cmdExit(2).Font = ft
    
    dbTextHeight = TextHeight("高")
    
    txtFindData.Height = dbTextHeight
    cmdStartFind.Height = dbTextHeight * 2
    
    cmdExit(0).Height = dbTextHeight * 2
    cmdExit(1).Height = dbTextHeight * 2
    cmdExit(2).Height = dbTextHeight * 2
    
    Picture1.Height = cmdStartFind.Height + cmdStartFind.Top * 2
    Picture2.Height = cmdExit(0).Height + cmdExit(0).Top * 2
End Sub


Private Sub cmdExit_Click(Index As Integer)
On Error GoTo errHandle
    Dim strQueueId As String
    
    Select Case Index
        Case 0
            Call Me.Hide
        Case 1, 2
            strQueueId = GetSelectId()
          
            If Trim(strQueueId) = "" Then
                MsgBox "尚未选择一条需要进行复诊操作的数据。", vbInformation, "排队叫号系统"
                Exit Sub
            End If
            
            If Index = 1 Then
                Call Execute_自定义(Val(strQueueId))
            ElseIf Index = 2 Then
                Call Execute_恢复(Val(strQueueId))
            End If
            
            '刷新数据
            Call cmdStartFind_Click
            
            'MsgBox "操作执行完成。", vbInformation, "排队叫号系统"
    End Select
    
Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Function GetSelectId() As String
'***************************************
'
'取得当前选中的数据
'
'***************************************
    On Error GoTo errHandle
        
'        If lvwQueueData.SelectedItem Is Nothing Then
'          GetSelectId = ""
'          Exit Function
'        End If
'
'        GetSelectId = lvwQueueData.SelectedItem.Tag
        
    Exit Function
errHandle:
      GetSelectId = ""
      If ErrCenter = 1 Then Resume
End Function


Private Sub cmdStartFind_Click()
    Dim rsData As ADODB.Recordset
    Dim strFindType As String
    Dim strFindValue As String
    Dim blnUseCustom As Boolean
    
    On Error GoTo errHandle
    strFindValue = txtFindData.Text

    If Trim(strFindValue) = "" Then
        MsgBox "请录入需要查找的数据值。", vbOKOnly, Me.Caption

        Call txtFindData.SetFocus
        Exit Sub
    End If

    '取得检索类型
    strFindType = cboFindWay.Text
    
    blnUseCustom = False
    RaiseEvent OnFind(strFindType, strFindValue, rsData, txtFindData, blnUseCustom)
    
    If Not blnUseCustom Then
        '使用默认的查询
        Set rsData = FindQueueData(strFindType, strFindValue)
    End If

    Call rptQueueList.Records.DeleteAll
    Call rptQueueList.Populate

    If rsData Is Nothing Then
'        MsgBox "没有检索到所需数据。", vbInformation, Me.Caption
        Exit Sub
    End If

    If rsData.RecordCount <= 0 Then
'        MsgBox "没有检索到所需数据。", vbInformation, Me.Caption
        Exit Sub
    End If

    Call LoadQueueData(rptQueueList, rsData)
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub



Private Function FindQueueData(ByVal findType As String, ByVal findData As String) As ADODB.Recordset
    Dim strSql As String, strFilter As String
    Dim str排队号 As String, str姓名 As String
    Dim strQueueNames As String
    
    On Error GoTo errHandle
    
    strFilter = ""
    
    Select Case findType  ' '0-排队号;1-姓名;
    Case "排队号"
        str排队号 = Val(findData)
        strFilter = " and 排队号码 = [2]"
    Case "姓名"
        str姓名 = findData & "%"
        strFilter = " and 患者姓名 Like [3]"
    End Select
    
    strQueueNames = mobjOwner.QueryQueueNames
    
    If strQueueNames <> "" Then
        strQueueNames = Replace(strQueueNames, ",", "','")
        strFilter = strFilter & " and 队列名称 in ('" & strQueueNames & "') "
    End If
    
    strSql = "select * from 排队叫号队列　where  业务类型=[1] " & strFilter & " order by 排队号码 "

    Set FindQueueData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mobjOwner.WorkType, str排队号, str姓名)
     
    Exit Function
errHandle:
    Set FindQueueData = Nothing
    If ErrCenter = 1 Then Resume

End Function

Private Sub LoadQueueData(objQueueList As ReportControl, rsData As ADODB.Recordset)
'载入队列数据

On Error GoTo errHandle
    Dim rptRecord As ReportRecord
    Dim blnCancel As Boolean
    Dim i As Long

'加载数据到列表

    Call objQueueList.Records.DeleteAll
    Call objQueueList.Populate
    
    If rsData.RecordCount <= 0 Then Exit Sub

    While Not rsData.EOF

        blnCancel = False
        RaiseEvent OnReadBefore(rsData, TQueueFromType.qftFindQueue, blnCancel)
        
        If Not blnCancel Then
            Set rptRecord = objQueueList.Records.Add
            
            For i = 0 To objQueueList.Columns.Count - 1
                rptRecord.AddItem ""
            Next
    
            Call SetReportRecordItem(rptRecord, objQueueList, rsData)
            
            RaiseEvent OnReadAfter(rsData, TQueueFromType.qftFindQueue, rptRecord)
        End If

        rsData.MoveNext
    Wend

    objQueueList.Populate

    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub Execute_恢复(ByVal Id As Long)
    On Error GoTo errHandle
        
        Dim strSql As String
        
        strSql = "ZL_排队叫号队列_恢复(" & Id & ")"
                
        Call zlDatabase.ExecuteProcedure(strSql, "复诊")
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub Execute_自定义(ByVal Id As Long)
On Error GoTo errHandle
    
    RaiseEvent OnCustomFindButton(Id)
Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub Form_Load()
    '恢复窗口状态
    Call RestoreWinState(Me, App.ProductName)
        
    cboFindWay.ListIndex = 1
    
'    Call InitInsertQueueList
End Sub


Private Sub Form_Resize()
On Error Resume Next
    rptQueueList.Left = 100
    rptQueueList.Top = Picture1.Height + 100
    rptQueueList.Width = Me.ScaleWidth - 200
    rptQueueList.Height = Picture2.Top - Picture1.Height - 200
err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub


Private Sub Picture2_Resize()
On Error Resume Next
    
    cmdExit(0).Left = Me.ScaleWidth - cmdExit(0).Width - 200
    cmdExit(2).Left = cmdExit(0).Left - cmdExit(2).Width - 50
    cmdExit(1).Left = cmdExit(2).Left - cmdExit(1).Width - 50
    
err.Clear
End Sub




Private Sub timerCard_Timer()
On Error GoTo errHandle
    If GetTickCount - mlngStartTime > 200 Then
        '大于200毫秒时，自动认为刷卡结束
        timerCard.Enabled = False
        
        mlngStartTime = 0
        mlngAvgTime = 0
        mlngReadCount = 0
        
        Call zlControl.TxtSelAll(txtFindData)
        
        Call cmdStartFind_Click
    End If
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub txtFindData_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Call cmdStartFind_Click
        Exit Sub
    End If
    
    If KeyAscii = 8 Then Exit Sub
    
    mlngReadCount = mlngReadCount + 1
    If mlngStartTime <> 0 Then
        If GetTickCount - mlngStartTime > 200 Then
            mlngReadCount = 1
            mlngAvgTime = 0
        Else
            mlngAvgTime = mlngAvgTime + (GetTickCount() - mlngStartTime)
        End If
    End If
    
    mlngStartTime = GetTickCount
    
    '取三次平均录入时间
    If mlngReadCount = 3 Then
        mlngAvgTime = Fix(mlngAvgTime / 3)
        
        If mlngAvgTime <= 30 Then timerCard.Enabled = True
    End If

End Sub
