VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFilm 
   Caption         =   "胶片打印"
   ClientHeight    =   8895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14400
   Icon            =   "frmFilm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   8895
   ScaleWidth      =   14400
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdFull 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Height          =   735
      Left            =   13320
      Picture         =   "frmFilm.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox picControl 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   10920
      ScaleHeight     =   1455
      ScaleWidth      =   3375
      TabIndex        =   1
      Top             =   7440
      Width           =   3375
      Begin VB.ComboBox cboPrint 
         Height          =   300
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "删 除(&D)"
         Height          =   375
         Left            =   1200
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "打 印(&P)"
         Height          =   375
         Left            =   2280
         TabIndex        =   2
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label labPrint 
         Caption         =   "打印机"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   300
         Width           =   735
      End
   End
   Begin MSComctlLib.ListView lvwFilm 
      Height          =   1455
      Left            =   0
      TabIndex        =   6
      Top             =   7440
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   2566
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.PictureBox picFilmPreview 
      BackColor       =   &H00000000&
      Height          =   7455
      Left            =   0
      ScaleHeight     =   7395
      ScaleWidth      =   14205
      TabIndex        =   0
      Top             =   0
      Width           =   14265
      Begin VB.Timer timerRefresh 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   5160
         Top             =   6120
      End
   End
End
Attribute VB_Name = "frmFilm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrOrderNo As String   '医嘱号码
Private mlngFilmHand As Long    '嵌入的胶片预览窗口句柄
Private mlngFilmProcessId As Long
Private mlngPrnCenterDevId As Long
Private mstrPrnCenterPath As String
Private mlngPrintType As Long   '-1-无数据,0-黑白相机,1-彩色相机

Private mblnIsFullScreen As Boolean
Private mlngSubFormHwnd As Long

Public Function ShowFilmPrintWnd(ByVal strOrderNo As String, objOwner As Object) As Boolean
'显示胶片打印窗口

    mstrOrderNo = strOrderNo
    mlngPrintType = -1
    mlngPrnCenterDevId = 0
    mblnIsFullScreen = False
    mlngSubFormHwnd = 0
    
    '初始化胶片列表
    Call InitFilmList
    
'    '加载胶片存放目录
'    Call LoadDicomCenterPath
    
    '加载胶片数据
    Call LoadFilmData(strOrderNo)
    
    '加载相机数据
    If mlngPrintType >= 0 Then
        Call LoadDicomPrint(mlngPrintType)
    End If

    Call Me.Show(1, objOwner)
End Function

Private Sub LoadDicomCenterPath()
'载入按需打印中心胶片存放路径
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    strSQL = "Select F_PARAM_VALUE From Ris.T_R_System_Param Where F_PARAM_ID=1380"
    Set rsData = gcnXWDBServer.Execute(strSQL)
    
    If rsData.RecordCount <= 0 Then Exit Sub
    
    mstrPrnCenterPath = Nvl(rsData!F_PARAM_VALUE)
    
End Sub


Private Sub LoadDicomPrint(ByVal lngPrintType As Long)
'载入dicom打印机
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim strPrintType As String
    
    strPrintType = IIf(lngPrintType = 0, "12", 50)
    
    strSQL = " select * from t_device l , t_dev_param m " & _
             " Where l.f_dev_id = m.f_dev_id " & _
                    " and l.f_dev_id in(select a.f_dev_id " & _
                                        " from t_device a, t_dev_param b " & _
                                        " where a.f_dev_id=b.f_dev_id and a.f_type_id=13 and upper(b.f_param_name) = upper('scu') and f_param_value='" & strPrintType & "') " & _
                    " and  upper(m.f_param_name)=upper('AE Title') " & _
                    " and  m.f_param_value not in(select substr(i.f_param_value, 0, instr(i.f_param_value, ',')-1) " & _
                                                    " from t_device h, t_dev_param i where h.f_type_id=90 and upper(i.f_param_name)=upper('DPC LocalAETitle') ) "
    Set rsData = gcnXWDBServer.Execute(strSQL)
    
    cboPrint.Clear
    
    If rsData.RecordCount <= 0 Then Exit Sub
    
    While Not rsData.EOF
        Call cboPrint.AddItem(Nvl(rsData!F_DEV_NAME))
        rsData.MoveNext
    Wend
    
    cboPrint.ListIndex = 0
End Sub


Private Sub LoadFilmData(ByVal strOrderNo As String)
'查询胶片数据
    Dim strSQL As String
    Dim rsFilm As ADODB.Recordset
        
    strSQL = "Select F_FILM_ID as 胶片ID, F_FILM_SIZE as 胶片大小,F_PRN_CENTER_DEV_ID as 打印中心ID, F_FILM_ORIEN as 胶片方位, F_FILM_FORMAT as 胶片格式,F_FILM_FILE as 胶片文件, " & _
                    " F_MODALITY As 设备类型, F_FILM_PRN_STATUS as 打印状态, F_TIME_RECV As 生成日期, F_TIME_PRINT As 打印日期, F_FILM_TYPE As 胶片类型 " & _
                    " From ris.v_p_film " & _
                    " Where F_PAT_NO = '" & strOrderNo & "' order by F_FILM_ID "
    
    Set rsFilm = gcnXWDBServer.Execute(strSQL)
    
    Call FillFilmData(rsFilm)
    
    If lvwFilm.ListItems.Count <= 0 Then Exit Sub
    
    '加载胶片显示样式
    lvwFilm.ListItems(1).Selected = True
    Call lvwFilm_Click
End Sub


Private Sub InitFilmList()
'初始化序列列表
    Dim tmpItem As ListItem
    
    With lvwFilm
        .ListItems.Clear
        
        '如果未初始化列，则进行初始化
        If .ColumnHeaders.Count <= 0 Then
            With .ColumnHeaders
                .Clear
                .Add , , "胶片ID", 750
                .Add , , "设备类型", 1000
                .Add , , "胶片格式", 1200
                .Add , , "胶片大小", 1200
                .Add , , "胶片方向", 1200
                .Add , , "胶片类型", 1200
                .Add , , "生成日期", 1400
                .Add , , "打印状态", 1200
                .Add , , "打印日期", 1400
                .Add , , "胶片文件", 0
            End With
        End If
    End With
End Sub


Private Sub FillFilmData(rsFilm As ADODB.Recordset)
'填充检查数据
    Dim tmpItem As ListItem
    
    lvwFilm.ListItems.Clear
    
    If Not rsFilm.EOF Then
        '判断当前胶片的相机类型
        mlngPrintType = IIf(UCase(Nvl(rsFilm!胶片类型)) = "GRAYSCALE", 0, 1)
        mlngPrnCenterDevId = Val(Nvl(rsFilm!打印中心ID))
        
        Do While Not rsFilm.EOF
            Set tmpItem = lvwFilm.ListItems.Add(, "_" & rsFilm("胶片ID"), Nvl(rsFilm("胶片ID")))
            With tmpItem
                .SubItems(1) = Nvl(rsFilm("设备类型"))
                .SubItems(2) = Nvl(rsFilm("胶片格式"))
                .SubItems(3) = Nvl(rsFilm("胶片大小"))
                .SubItems(4) = Nvl(rsFilm("胶片方位"))
                .SubItems(5) = Nvl(rsFilm("胶片类型"))
                .SubItems(6) = Nvl(rsFilm("生成日期"))
                .SubItems(7) = IIf(Nvl(rsFilm("打印日期")) = "", "", _
                                Decode(Val(Nvl(rsFilm("打印状态"))), 2101, "开始打印", 2102, "打印完成", 2103, "打印错误", 2104, "直接打印", ""))
                .SubItems(8) = Nvl(rsFilm("打印日期"))
                .SubItems(9) = Nvl(rsFilm("胶片文件"))
                '.Checked = True
            End With
            rsFilm.MoveNext
        Loop
    End If
End Sub

Private Sub DelFilm()
    Dim i As Long
    Dim strFilmId As String
    
    For i = lvwFilm.ListItems.Count To 1 Step -1
        If lvwFilm.ListItems(i).Selected Then
            strFilmId = Mid(lvwFilm.ListItems(i).Key, 2)
            
            If XWFilmDelete(strFilmId) = False Then
                MsgBoxD Me, "删除胶片ID为[" & strFilmId & "] 时失败。", vbOKOnly, gstrSysName
                Exit Sub
            End If
            
            Call lvwFilm.ListItems.Remove(i)
        End If
    Next i
    
    If lvwFilm.ListItems.Count > 0 Then
        lvwFilm.ListItems(lvwFilm.ListItems.Count).Selected = True
        Call lvwFilm_Click
    End If
End Sub

 

Private Sub cmdDel_Click()
'删除当前选中的胶片
On Error GoTo ErrHandle
    If lvwFilm.SelectedItem Is Nothing Then
        MsgBoxD Me, "请选择需要删除的胶片记录。", vbOKOnly, gstrSysName
        Exit Sub
    End If
    
    '胶片删除操作确认提示
    If MsgBoxD(Me, "是否确认删除所选的胶片？删除后将不能恢复。", vbYesNo, gstrSysName) = vbNo Then
        Exit Sub
    End If
    
    Call DelFilm
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub PrintFilm()
'打印胶片
    Dim lngPrintResult As Long
    Dim i As Long
    
    lngPrintResult = XWFilmPrint(mstrOrderNo, mlngPrintType, cboPrint.Text)
    If lngPrintResult <> 0 Then
        Call MsgBoxD(Me, "胶片打印发送失败，错误代码：" & lngPrintResult, vbOKOnly, gstrSysName)
        Exit Sub
    End If
                
    For i = 1 To lvwFilm.ListItems.Count
        lvwFilm.ListItems(i).SubItems(7) = "已发送"
    Next i
            
End Sub

Private Sub cmdFull_Click()
On Error GoTo ErrHandle
    Dim lngScrollBarHwnd As Long
    
    mlngSubFormHwnd = FindWindowEx(mlngFilmHand, 0, "AfxFrameOrView42", vbNullString)
    If mlngSubFormHwnd > 0 Then

        SetParent mlngSubFormHwnd, Me.hWnd
        
        Call MoveWindow(mlngSubFormHwnd, 0, 0, Me.ScaleX(Me.ScaleWidth, vbTwips, vbPixels), _
                                            Me.ScaleY(Me.ScaleHeight, vbTwips, vbPixels), 1)
                                            
        mblnIsFullScreen = True
    End If
Exit Sub
ErrHandle:
    MsgBox err.Description
End Sub

Private Sub cmdPrint_Click()
'打印当前选中的胶片
On Error GoTo ErrHandle
    If lvwFilm.ListItems.Count <= 0 Then
        MsgBoxD Me, "没有可供打印的胶片数据。", vbOKOnly, gstrSysName
        Exit Sub
    End If
    
    If cboPrint.Text = "" Then
        MsgBoxD Me, "没有选择打印机，不能执行打印操作。", vbOKOnly, gstrSysName
        Exit Sub
    End If
    
    Call PrintFilm
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdSure_Click()
'确认当前选中的胶片
On Error GoTo ErrHandle
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyEscape Then
        If mblnIsFullScreen = True Then
            Call QuitFullScreen
            mblnIsFullScreen = False
        End If
    End If
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandle
    Call RestoreWinState(Me, App.ProductName)
Exit Sub
ErrHandle:
End Sub

Private Sub QuitFullScreen()
'退出全屏
    SetParent mlngSubFormHwnd, mlngFilmHand
    
    Call picFilmPreview_Resize
    Call MoveWindow(mlngSubFormHwnd, 0, 55, picFilmPreview.ScaleX(picFilmPreview.ScaleWidth, vbTwips, vbPixels), _
                                        picFilmPreview.ScaleY(picFilmPreview.ScaleHeight, vbTwips, vbPixels) - 55, 1)
                                        
    mlngSubFormHwnd = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mblnIsFullScreen = True Then
        Cancel = True

        Call QuitFullScreen
        
        mblnIsFullScreen = False
    End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
    picFilmPreview.Left = 60
    picFilmPreview.Top = 60
    picFilmPreview.Width = Me.ScaleWidth - 120
    picFilmPreview.Height = Me.ScaleHeight - lvwFilm.Height - 180
    
    lvwFilm.Left = 60
    lvwFilm.Top = picFilmPreview.Top + picFilmPreview.Height + 60
    lvwFilm.Width = Me.ScaleWidth - picControl.Width - 120
    
    picControl.Left = lvwFilm.Left + lvwFilm.Width
    picControl.Top = lvwFilm.Top
    picControl.Height = lvwFilm.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    If mlngFilmHand <> 0 Then
        SendMessage mlngFilmHand, WM_CLOSE, 0, 0
        mlngFilmHand = 0
    End If
    
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub lvwFilm_Click()
'加载选中的胶片样式
On Error GoTo ErrHandle
    Dim strFilmId As String
    Dim lngResult As Long
    Dim lngProcessId As Long
    
    If lvwFilm.SelectedItem Is Nothing Then Exit Sub
    
    strFilmId = Mid(lvwFilm.SelectedItem.Key, 2)
    
    If strFilmId <= 0 Then Exit Sub
    
    If mlngFilmHand <> 0 Then
        '隐藏已经嵌入的窗口
        ShowWindow mlngFilmHand, SW_HIDE
        SetParent mlngFilmHand, 0
    End If

    lngResult = XWFilmPreviewEx(strFilmId)
    
    If lngResult <> 0 Then
        If mlngFilmHand <> 0 Then

            SendMessage mlngFilmHand, WM_CLOSE, 0, 0
            mlngFilmHand = 0
        End If
        
        MsgBoxD Me, "错误代码：" & lngResult, vbOKOnly, gstrSysName
        
        Exit Sub
    End If
    
    cmdFull.Visible = IIf(lngResult = 0, True, False)
    timerRefresh.Enabled = True
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Function SetWindowStyle(ByVal lngHandle As Long) As Long
'设置窗口样式（取消窗口边框）
    Dim lngWindowStyle As Long
    
    lngWindowStyle = GetWindowLong(lngHandle, GWL_STYLE)
    
    lngWindowStyle = lngWindowStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX Or WS_THICKFRAME) 'Or WS_THICKFRAME

    SetWindowStyle = SetWindowLong(lngHandle, GWL_STYLE, lngWindowStyle)    'Or WS_CHILD
End Function


Private Sub picFilmPreview_Resize()
On Error Resume Next
    If mlngFilmHand > 0 Then
        Call MoveWindow(mlngFilmHand, 0, 0, picFilmPreview.ScaleX(picFilmPreview.ScaleWidth, vbTwips, vbPixels), _
                                            picFilmPreview.ScaleY(picFilmPreview.ScaleHeight, vbTwips, vbPixels), 1)
                                            
        If mblnIsFullScreen = True And mlngSubFormHwnd > 0 Then
            Call MoveWindow(mlngSubFormHwnd, 0, 0, Me.ScaleX(Me.ScaleWidth, vbTwips, vbPixels), _
                                                Me.ScaleY(Me.ScaleHeight, vbTwips, vbPixels), 1)
        End If
    End If
    
    cmdFull.Left = picFilmPreview.ScaleWidth - cmdFull.Width - 365
    cmdFull.Top = picFilmPreview.ScaleHeight - cmdFull.Height - 40
End Sub

Private Sub timerRefresh_Timer()
On Error GoTo ErrHandle
    Dim lngToolbarHwnd As Long
    
    timerRefresh.Enabled = False
    
    '设置为嵌入式窗口
    If mlngFilmHand = 0 Then
        mlngFilmHand = FindWindow(vbNullString, "FilmPreview")
        
        lngToolbarHwnd = FindWindowEx(mlngFilmHand, 0, "ToolbarWindow32", vbNullString)
        
        If lngToolbarHwnd <> 0 Then
            '隐藏toolbar中的部分功能按钮
            SendMessage lngToolbarHwnd, WM_USER + 22, 0, 1
            SendMessage lngToolbarHwnd, WM_USER + 22, 0, 2
            SendMessage lngToolbarHwnd, WM_USER + 22, 0, 3
        End If
        
    End If
    
    If mlngFilmHand <> 0 Then
        Call SetWindowStyle(mlngFilmHand)
        
        Call SetParent(mlngFilmHand, picFilmPreview.hWnd)
        
        Call MoveWindow(mlngFilmHand, 0, 0, picFilmPreview.ScaleX(picFilmPreview.ScaleWidth, vbTwips, vbPixels), _
                                            picFilmPreview.ScaleY(picFilmPreview.ScaleHeight, vbTwips, vbPixels), 1)
        
        Call ShowWindow(mlngFilmHand, SW_SHOWMAXIMIZED)
    End If
    
Exit Sub
ErrHandle:
    timerRefresh.Enabled = False
End Sub
