VERSION 5.00
Begin VB.Form frmPatholDecalcification 
   Caption         =   "脱钙任务管理"
   ClientHeight    =   6495
   ClientLeft      =   75
   ClientTop       =   405
   ClientWidth     =   9240
   Icon            =   "frmPatholDecalcification.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   9240
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer timeDate 
      Interval        =   1000
      Left            =   4920
      Top             =   5880
   End
   Begin VB.Timer timeDecalin 
      Interval        =   30000
      Left            =   3960
      Top             =   5880
   End
   Begin VB.CommandButton cmdSucceed 
      Caption         =   "完 成(&F)"
      Height          =   400
      Left            =   7920
      TabIndex        =   2
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "换 缸(&H)"
      Height          =   400
      Left            =   6600
      TabIndex        =   1
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Frame framDecalin 
      Caption         =   "脱钙记录"
      Height          =   5655
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   9015
      Begin zl9PACSWork.ucFlexGrid ufgData 
         Height          =   5295
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   9340
         DefaultCols     =   ""
         GridRows        =   21
         IsKeepRows      =   0   'False
         BackColor       =   12648447
         IsCopyAdoMode   =   0   'False
         IsEjectConfig   =   -1  'True
         Editable        =   0
         HeadFontCharset =   134
         HeadFontWeight  =   400
         DataFontCharset =   134
         DataFontWeight  =   400
      End
   End
   Begin VB.Label labTime 
      Caption         =   "当前时间: 2011-11-11 11:11:11"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   6000
      Width           =   3015
   End
End
Attribute VB_Name = "frmPatholDecalcification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrPrivs As String
Private mblnMoved As Boolean
Private mlngModul As Long

Private mblnIsSoundHint As Boolean
Private mlngHintTime As Long

Private mfrmParent As Form

Private mblnPlaySound As Boolean


Public Sub ShowDecalinTaskWind(ByVal strPrivs As String, ByVal blnMoved As Boolean, ByVal lngModul As Long, owner As Form)
'显示脱钙任务窗口
    SetWindowPos Me.hWnd, -1, Me.CurrentX, Me.CurrentY, Me.ScaleWidth, Me.ScaleHeight, 3 '将窗口置顶
    
    mstrPrivs = strPrivs
    mblnMoved = blnMoved
    mlngModul = lngModul
    
    Set mfrmParent = owner
    
    
    '初始化参数
    Call InitParameter
    
    
    If Not owner.Visible Then Exit Sub
    
    Call Me.Show(0, owner)
End Sub

Private Sub InitDecalinList()
'初始化脱钙任务列表
    Dim strTemp As String
    


     '判断数据库参数表是否有数据 有则读取数据库参数  没有则加载默认
    strTemp = zlDatabase.GetPara("脱钙任务列表配置", glngSys, G_LNG_PATHOLSYS_NUM, "")
    ufgData.DefaultColNames = gstrDecalinTaskCols
	
    If strTemp = "" Then
        ufgData.ColNames = gstrDecalinTaskCols
    Else
        ufgData.ColNames = strTemp
    End If
        '设置行数
    ufgData.GridRows = glngStandardRowCount
    '设置行高
    ufgData.RowHeightMin = glngStandardRowHeight
    ufgData.ColConvertFormat = gstrDecalinConvertFormat
End Sub


Private Sub ufgData_OnColFormartChange()
  '保存列表参数
    zlDatabase.SetPara "脱钙任务列表配置", ufgData.GetColsString(ufgData), glngSys, G_LNG_PATHOLSYS_NUM
End Sub


Private Sub LoadDecalinData()
'载入脱钙信息
    Dim strSQL As String
    
    strSQL = "select a.ID,a.标本ID,c.病理号, b.标本名称,a.开始时间,case when a.所需时长 / 60 < 1 then '0' else '' end || to_char(a.所需时长 / 60) as 所需时长, (case when a.所需时长 - ((sysdate - a.开始时间) * 24 * 60 ) < 0 then 0 else trunc(a.所需时长 - ((sysdate - a.开始时间) * 24 * 60 )) end) as 剩余时长, (a.开始时间 + a.所需时长/60/24) as 结束时间, a.当前缸次,a.完成状态,a.操作员" & _
                " from 病理脱钙信息 a, 病理标本信息 b, 病理检查信息 c" & _
                " where a.标本id = b.标本id and b.医嘱ID = c.医嘱ID and a.操作员=[1] and a.完成状态<>1 and a.开始时间>sysdate - 30 order by 完成状态,剩余时长,开始时间,ID"
    
'    If mblnMoved Then strSql = GetMovedDataSql(strSql)
    
    Set ufgData.AdoData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.姓名)
    
    Call ufgData.RefreshData
End Sub


Private Sub Decalin_Change(ByVal dtStart As Date, ByVal lngTimeLen As Double)
'脱钙换缸
    Dim strSQL As String
    Dim lngDecalinId As Long
    
    lngDecalinId = ufgData.KeyValue(ufgData.SelectionRow)
    
    strSQL = "Zl_病理脱钙_换缸(" & lngDecalinId & "," & zlStr.To_Date(dtStart) & "," & Fix(lngTimeLen * 60) & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    '更新脱钙显示列表
    ufgData.Text(ufgData.SelectionRow, gstrDecalin_当前缸次) = Val(ufgData.Text(ufgData.SelectionRow, gstrDecalin_当前缸次)) + 1
    ufgData.Text(ufgData.SelectionRow, gstrDecalin_开始时间) = dtStart
    ufgData.Text(ufgData.SelectionRow, gstrDecalin_所需时长) = Format$(lngTimeLen, "0.0")
    ufgData.Text(ufgData.SelectionRow, gstrDecalin_剩余时长) = Fix(lngTimeLen * 60)
    ufgData.Text(ufgData.SelectionRow, gstrDecalin_结束时间) = DateAdd("n", lngTimeLen * 60, dtStart)
End Sub


Private Sub cmdChange_Click()
On Error GoTo ErrHandle
    Dim frmChangeInput As frmPatholMaterials_Change

    If Not ufgData.IsSelectionRow Then
        Call MsgBoxD(Me, "请选择需要换缸的记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgData.IsNullRow(ufgData.SelectionRow) Then
        Call MsgBoxD(Me, "请选择需要换缸的记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    '判断当前记录是否已经开始脱钙
    If ufgData.IsEmptyKey(ufgData.SelectionRow) Then
        Call MsgBoxD(Me, "该标本尚未开始脱钙，不能执行换缸操作，请先执行脱钙。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgData.Text(ufgData.SelectionRow, gstrDecalin_当前状态) = "已完成" Then
        Call MsgBoxD(Me, "脱钙任务已完成，不能进行换缸操作。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Set frmChangeInput = New frmPatholMaterials_Change
    On Error GoTo errFree
    
        Call frmChangeInput.ShowChangeWindow(Me)
            
        If Not frmChangeInput.IsSure Then Exit Sub
        
        '换缸
        Call Decalin_Change(frmChangeInput.StartTime, frmChangeInput.TimeLen)
errFree:
    Unload frmChangeInput
    Set frmChangeInput = Nothing
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Decalin_Succed()
'完成脱钙
    Dim strSQL As String
    Dim lngDecalinId As Long
    
    lngDecalinId = ufgData.KeyValue(ufgData.SelectionRow)
    
    strSQL = "Zl_病理脱钙_完成(" & lngDecalinId & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    '更新脱钙显示列表
    ufgData.Text(ufgData.SelectionRow, gstrDecalin_当前状态) = "已完成"
End Sub


Private Sub cmdSucceed_Click()
On Error GoTo ErrHandle


    If Not ufgData.IsSelectionRow Then
        Call MsgBoxD(Me, "请选择需要完成脱钙的记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    If ufgData.IsNullRow(ufgData.SelectionRow) Then
        Call MsgBoxD(Me, "请选择需要完成脱钙的记录。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    '判断当前记录是否已经开始脱钙
    If ufgData.IsEmptyKey(ufgData.SelectionRow) Then
        Call MsgBoxD(Me, "该标本尚未开始脱钙，不能执行该操作，请先执行脱钙。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    Call Decalin_Succed
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle
    Call RestoreWinState(Me, App.ProductName)
    
    Call InitDecalinList
    
    Call LoadDecalinData
    
    Call CheckListState
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub InitParameter()
On Error Resume Next
    mblnIsSoundHint = Val(zlDatabase.GetPara("脱钙声音提醒", glngSys, mlngModul, 1))
    mlngHintTime = Val(zlDatabase.GetPara("提醒间隔时长", glngSys, mlngModul, "30"))
    
    timeDecalin.Interval = mlngHintTime * 1000
End Sub


Private Sub AdjustFace()
    framDecalin.Left = 120
    framDecalin.Top = 120
    framDecalin.Width = Me.Width - 360
    framDecalin.Height = Me.Height - cmdSucceed.Height - 900
    
    ufgData.Left = 120
    ufgData.Top = 240
    ufgData.Width = framDecalin.Width - 240
    ufgData.Height = framDecalin.Height - 360
    
    cmdSucceed.Left = Me.Width - cmdSucceed.Width - 240
    cmdSucceed.Top = Me.Height - cmdSucceed.Height - 620
    
    cmdChange.Left = cmdSucceed.Left - cmdChange.Width - 120
    cmdChange.Top = cmdSucceed.Top
    
    labTime.Left = 120
    labTime.Top = cmdSucceed.Top + 60
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call Me.Hide
    If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Call AdjustFace
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
End Sub


Private Sub timeDate_Timer()
On Error Resume Next
    labTime.Caption = "当前时间：" & Format(zlDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")
    
    '其他提示
    If mblnPlaySound Then
        If Not (mfrmParent Is Nothing) Then
            If InStr(mfrmParent.Caption, "        ") <= 0 Then mfrmParent.Caption = mfrmParent.Caption & "        "
            
            If mfrmParent.Caption Like "*脱钙任务已完成*" Then
                mfrmParent.Caption = Replace(mfrmParent.Caption, "        脱钙任务已完成！！！", "        ")
            Else
                mfrmParent.Caption = Replace(mfrmParent.Caption, "        ", "        脱钙任务已完成！！！")
            End If
        End If
    End If
End Sub

Private Sub timeDecalin_Timer()
On Error Resume Next
    Call LoadDecalinData
    
    Call CheckListState
End Sub


Private Sub PalyHintSound()
'播放提示声音
    Call Beep(2000, 100)
    Call Beep(1000, 100)
    Call Beep(2000, 100)
    Call Beep(1000, 100)
    Call Beep(2000, 100)
    Call Beep(1000, 100)
End Sub


Private Sub CheckListState()
'检查脱钙任务列表状态
    Dim i As Long
    
    
    mblnPlaySound = False
    For i = 1 To ufgData.GridRows - 1
        If Val(ufgData.Text(i, gstrDecalin_剩余时长)) = 0 Then
            Call ufgData.SetRowColor(i, &H80FF80)
            
            mblnPlaySound = True
            
        ElseIf Val(ufgData.Text(i, gstrDecalin_剩余时长)) < 5 Then
            Call ufgData.SetRowColor(i, &H80FFFF)
        Else
            Call ufgData.SetRowColor(i, ufgData.BackColor)
        End If
    Next i
    
    
    '声音提示
    If mblnPlaySound And mblnIsSoundHint Then Call PalyHintSound
End Sub




Private Sub ufgData_OnColsNameReSet()
On Error GoTo ErrHandle

    Call LoadDecalinData

Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub
