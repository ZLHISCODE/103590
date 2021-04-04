VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmAlert 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "提醒消息"
   ClientHeight    =   6330
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7680
   Icon            =   "frmAlert.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Visible         =   0   'False
   Begin MSComctlLib.ImageList ils16 
      Left            =   6135
      Top             =   -30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   10
      ImageHeight     =   10
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":01CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":0408
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.UpDown udn 
      Height          =   300
      Left            =   3690
      TabIndex        =   6
      Top             =   5850
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txt"
      BuddyDispid     =   196611
      OrigLeft        =   3810
      OrigTop         =   5175
      OrigRight       =   4050
      OrigBottom      =   5520
      Max             =   20
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Height          =   300
      Left            =   2970
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "1"
      Top             =   5850
      Width           =   705
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "退出(&X)"
      Height          =   350
      Left            =   6525
      TabIndex        =   4
      Top             =   5835
      Width           =   1100
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "刷新(&R)"
      Height          =   350
      Left            =   5385
      TabIndex        =   3
      Top             =   5835
      Width           =   1100
   End
   Begin MSWinsockLib.Winsock winSock 
      Left            =   7830
      Top             =   1305
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrIcon 
      Left            =   7785
      Top             =   1935
   End
   Begin VB.PictureBox picNotify 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   4680
      ScaleHeight     =   345
      ScaleWidth      =   225
      TabIndex        =   7
      Top             =   5835
      Visible         =   0   'False
      Width           =   225
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   2355
      Left            =   60
      TabIndex        =   0
      Top             =   285
      Width           =   7575
      _cx             =   13361
      _cy             =   4154
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16761024
      ForeColorSel    =   -2147483641
      BackColorBkg    =   -2147483634
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483639
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   270
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   1
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   8
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
      Begin VB.Line lnY 
         Index           =   0
         Visible         =   0   'False
         X1              =   2940
         X2              =   2940
         Y1              =   1080
         Y2              =   1905
      End
      Begin VB.Line lnX 
         Index           =   0
         Visible         =   0   'False
         X1              =   960
         X2              =   2220
         Y1              =   1830
         Y2              =   1830
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3090
      Left            =   60
      TabIndex        =   9
      Top             =   2580
      Width           =   7575
      Begin VB.TextBox txtDetail 
         Height          =   2340
         Left            =   90
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   2
         Top             =   675
         Width           =   7395
      End
      Begin VB.TextBox txtReport 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1110
         Locked          =   -1  'True
         MouseIcon       =   "frmAlert.frx":069E
         MousePointer    =   99  'Custom
         TabIndex        =   1
         Top             =   450
         Width           =   6315
      End
      Begin VB.Label lblCheck 
         AutoSize        =   -1  'True
         Caption         =   "检查时间:"
         Height          =   180
         Left            =   105
         TabIndex        =   11
         Top             =   195
         Width           =   810
      End
      Begin VB.Label lblReport 
         AutoSize        =   -1  'True
         Caption         =   "提醒报表:"
         Height          =   180
         Left            =   105
         TabIndex        =   10
         Top             =   450
         Width           =   810
      End
   End
   Begin MSWinsockLib.Winsock wskTest 
      Left            =   4980
      Top             =   5820
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "弹出窗口每条提醒信息停留时间(&T)            秒"
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   5925
      Width           =   4050
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "你目前总共有 2 条提醒信息。"
      Height          =   180
      Index           =   0
      Left            =   75
      TabIndex        =   8
      Top             =   45
      Width           =   2430
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Index           =   0
      Left            =   7935
      Picture         =   "frmAlert.frx":09A8
      Top             =   3225
      Width           =   240
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Index           =   1
      Left            =   8040
      Top             =   2790
      Width           =   240
   End
End
Attribute VB_Name = "frmAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mintIcon As Integer
Private mblnIcon As Boolean
Private mblnStartUp As Boolean

Private mlngModual As Long
Private mlngSys As Long

Private mlngBind As Long
Private mblnRefreshing As Boolean
Public mblnUnLoad As Boolean

Private Function Ping(ByVal strServer As String) As Boolean
'    Dim lngReturn As Long
'    Dim lngProcess As Long
'
'    If strServer = "" Then Exit Function
'
'    lngReturn = Shell("ping " & strServer, vbHide)
'    lngProcess = OpenProcess(Process_Query_Information, False, lngReturn)
'    Do
'        Sleep 100
'        GetExitCodeProcess lngProcess, lngReturn
'        DoEvents
'    Loop While lngReturn = Still_Active
'    CloseHandle lngProcess
'
'    Ping = (lngReturn = 0)
    
'    Ping = pingnet(strServer)
End Function

Public Sub InitAlert()
    '-----------------------------------------
    Dim rs As New ADODB.Recordset
        
    On Error GoTo errHand

    If ConnectServer = False Then
        Exit Sub
    End If
            
    '动态配置工作站的端口号
'    DoEvents
    
    mlngBind = SearchBind
    
    If mlngBind > 1024 Then
        
        On Error Resume Next
                    
        winSock.SendData ""
        
        '将客户端的用户情况通知给服务器
        winSock.SendData "[SYS-COMPUTER]" & winSock.LocalHostName & ";" & mlngBind & ";" & gstrUserName & ";" & gstrDbUser & ";" & glngSys & ";" & winSock.RemoteHost
                           
        '启动检查请求
        winSock.SendData "[SYS-STARTUP]" & winSock.LocalHostName & ";" & mlngBind & ";" & gstrUserName & ";" & gstrDbUser & ";" & glngSys & ";" & winSock.RemoteHost
        
    End If
    
errHand:
    
End Sub

Public Sub InitData()
    
    txt.Text = Val(zlDatabase.GetPara("自动消息停留时间"))
    If Val(txt.Text) < udn.Min Or Val(txt.Text) > udn.Max Then txt.Text = "3"
        
    With vsf
        .Rows = 2
        .Cols = 6
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = ""
        
        .Cell(flexcpPictureAlignment, 0, 0) = flexPicAlignCenterCenter
        .Cell(flexcpPictureAlignment, 0, 1) = flexPicAlignCenterCenter
        Set .Cell(flexcpPicture, 1, 0) = Nothing
        Set .Cell(flexcpPicture, 1, 1) = Nothing
        
        .TextMatrix(0, 2) = "系统;模块"
        .TextMatrix(0, 3) = "报表名称"
        .TextMatrix(0, 4) = "检查时间"
        
        .TextMatrix(0, 5) = "内容"
        
        .ColHidden(2) = True
        .ColHidden(3) = True
        .ColHidden(4) = True
        Set .Cell(flexcpPicture, 0, 0) = ils16.ListImages(2).Picture
        Set .Cell(flexcpPicture, 0, 1) = ils16.ListImages(1).Picture
        .ExtendLastCol = True
    End With
        
    vsf.ColAlignment(0) = 4
    vsf.ColWidth(0) = 300
    vsf.ColWidth(1) = 180
    vsf.ColWidth(2) = 0
    
    Call AppendSapceRows(vsf, lnX, lnY)
    
    DoEvents
    Me.SetFocus

    Call cmdRefresh_Click
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdRefresh_Click()
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim lngSvrKey As Long
    Dim lngLoop As Long
    
    lngSvrKey = Val(vsf.RowData(vsf.Row))
    mblnRefreshing = True
    
    On Error GoTo errHand

    vsf.Rows = 2
    vsf.RowData(1) = ""
    vsf.Cell(flexcpText, 1, 0, 1, vsf.Cols - 1) = ""
    txtReport.Text = ""
    txtDetail.Text = ""
    lblCheck.Caption = "检查时间:"

    gstrSQL = "select A.序号,A.系统, C.程序ID AS 模块,C.系统 As 报表系统, B.提醒内容 AS 结果内容, C.名称 AS 提醒报表, A.提醒声音,B.检查时间,B.已读标志 " & _
             "from zlNotices A, " & _
                  "zlNoticeRec B, " & _
                  "(SELECT * FROM zlReports WHERE 发布时间 IS NOT NULL) C " & _
            "where B.用户名 = [1] and B.提醒标志 >0 AND C.编号(+) = A.提醒报表 AND " & _
                  "A.序号 = B.提醒序号 AND B.提醒内容 IS NOT NULL"

    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, gstrUserName)
    If rs.BOF = False Then
        Do While Not rs.EOF
        
            vsf.Cell(flexcpPictureAlignment, vsf.Rows - 1, 0) = flexPicAlignCenterCenter
            vsf.Cell(flexcpPictureAlignment, vsf.Rows - 1, 1) = flexPicAlignCenterCenter
            
            Set vsf.Cell(flexcpPicture, vsf.Rows - 1, 0) = ils16.ListImages(3).Picture
            
            vsf.RowData(vsf.Rows - 1) = NVL(rs("序号").Value, 0)
            
            vsf.TextMatrix(vsf.Rows - 1, 2) = CStr(NVL(rs("报表系统").Value, 0) & ";" & NVL(rs("模块").Value, 0))
            vsf.TextMatrix(vsf.Rows - 1, 3) = NVL(rs("提醒报表").Value)
            vsf.TextMatrix(vsf.Rows - 1, 4) = Format(NVL(rs("检查时间").Value), "yyyy年MM月dd日 HH时mm分")
            vsf.TextMatrix(vsf.Rows - 1, 5) = NVL(rs("结果内容").Value)
            
            If NVL(rs("已读标志").Value, 0) = 0 Then
                vsf.Cell(flexcpFontBold, vsf.Rows - 1, 0, vsf.Rows - 1, vsf.Cols - 1) = True
            End If
            
            If NVL(rs("提醒报表").Value, "") <> "" Then
                Set vsf.Cell(flexcpPicture, vsf.Rows - 1, 1) = ils16.ListImages(1).Picture
            End If

            vsf.Rows = vsf.Rows + 1

            rs.MoveNext
        Loop

        If vsf.Rows > 1 Then vsf.Rows = vsf.Rows - 1
    End If

    If rs.RecordCount > 0 Then
        lbl(0).Caption = "你目前总共有 " & rs.RecordCount & " 条提醒信息。"
    Else
        lbl(0).Caption = "你目前没有任何提醒信息。"
    End If
    
    vsf.Row = 0
    For lngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(lngLoop)) = lngSvrKey Then
            vsf.Row = lngLoop
            Exit For
        End If
    Next
    
    If lngLoop = vsf.Rows Then
        vsf.Row = 1
    End If
    
    mblnRefreshing = False
    Call vsf_AfterRowColChange(0, 0, vsf.Row, vsf.Col)
    
    Call AppendSapceRows(vsf, lnX, lnY)
    
    Exit Sub

errHand:

    If ErrCenter = 1 Then Resume
End Sub

Private Sub Form_Activate()
    
    If mblnStartUp = False Then Exit Sub
    mblnStartUp = False
                
End Sub

Private Sub Form_Load()

    mblnStartUp = True
    mblnUnLoad = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If mblnUnLoad = False Then
        Cancel = True
        Me.Hide
        Exit Sub
    End If
    
    On Error GoTo errHand
    
    If ConnectServer Then
    
        '将客户端的用户情况通知给服务器
        
        winSock.SendData "[SYS-DISCONNECT]" & winSock.LocalHostName & ";" & mlngBind & ";" & gstrUserName & ";" & gstrDbUser & ";" & glngSys & ";" & winSock.RemoteHost
        
    End If
    
errHand:

End Sub

Private Sub tmrIcon_Timer()

    If mblnIcon = False Then Exit Sub
    
    On Error Resume Next

    mintIcon = mintIcon + 1
    mintIcon = mintIcon Mod (imgIcon.UBound + 1)

    Call ModifyIcon(picNotify.hWnd, imgIcon(mintIcon).Picture)

End Sub

Private Sub txt_Change()
    Call zlDatabase.SetPara("自动消息停留时间", Val(txt.Text))
End Sub

Private Sub txt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt.Locked Then
        glngTXTProc = GetWindowLong(txt.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt.Locked Then
        Call SetWindowLong(txt.hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txtDetail_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txtDetail.Locked Then
        glngTXTProc = GetWindowLong(txtDetail.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtDetail.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtDetail_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txtDetail.Locked Then
        Call SetWindowLong(txtDetail.hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txtReport_Click()
    Call txtReport_GotFocus
End Sub

Private Sub txtReport_DblClick()
    If txtReport.Tag <> "" Then

        mlngSys = Val(Split(txtReport.Tag, ";")(0))
        mlngModual = Val(Split(txtReport.Tag, ";")(1))
        
        If mlngModual > 0 Then
            Call gfrmMain.RunModual(mlngSys, mlngModual)
        End If
        
    End If
End Sub

Private Sub txtReport_GotFocus()
    
    txtReport.SelStart = 0
    txtReport.SelLength = Len(txtReport.Text)
    
End Sub

Private Sub txtReport_KeyPress(KeyAscii As Integer)
    Call txtReport_GotFocus
End Sub

Private Sub txtReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txtReport.Locked Then
        glngTXTProc = GetWindowLong(txtReport.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtReport.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtReport_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txtReport.Locked Then
        Call SetWindowLong(txtReport.hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub udn_Change()
    Call zlDatabase.SetPara("自动消息停留时间", Val(txt.Text))
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim lngLoop As Long
    
    If mblnRefreshing Then Exit Sub
    
    If NewRow <> OldRow Then
        lblCheck.Caption = "检查时间:  " & vsf.TextMatrix(NewRow, 4)
        txtReport.Text = vsf.TextMatrix(NewRow, 3)
        txtReport.Tag = vsf.TextMatrix(NewRow, 2)
        txtDetail.Text = vsf.TextMatrix(NewRow, 5)
        
        If vsf.Cell(flexcpFontBold, NewRow, 0, NewRow, vsf.Cols - 1) = True Then
            
            vsf.Cell(flexcpFontBold, NewRow, 0, NewRow, vsf.Cols - 1) = False
            
            If ConnectServer Then
                '回置已读标志
                winSock.SendData "[SYS-READED]" & vsf.RowData(NewRow) & ";" & gstrUserName
            End If
            
        End If
                
        '检查是否还有新消息
        For lngLoop = 1 To vsf.Rows - 1
            If vsf.Cell(flexcpFontBold, lngLoop, 5) = True Then
                '有新消息
                Exit For
            End If
        Next
        
        If lngLoop = vsf.Rows Then
        
            Call RemoveIcon(picNotify.hWnd)

        End If
    End If
    
End Sub

Private Sub vsf_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Call AppendSapceRows(vsf, lnX, lnY)
End Sub

Private Sub vsf_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call AppendSapceRows(vsf, lnX, lnY)
End Sub

Private Sub vsf_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then Cancel = True
End Sub

Private Sub winSock_DataArrival(ByVal bytesTotal As Long)
    
    Dim strData As String
    
    On Error Resume Next
    
    winSock.GetData strData
    
    Select Case strData
    Case "[SYS-COMPUTER]"
                        
        If ConnectServer Then
            '服务器请求客户端资料
            winSock.SendData "[SYS-COMPUTER]" & winSock.LocalHostName & ";" & mlngBind & ";" & gstrUserName & ";" & gstrDbUser & ";" & glngSys & ";" & winSock.RemoteHost
        End If
        
        DoEvents
    
    Case "[SYS-TEST]"           '检查工作站是否存在
    
        If ConnectServer Then
            '服务器请求客户端资料
            winSock.SendData "[SYS-TEST]"
        End If
        
        DoEvents
    Case Else
                
        If Trim(strData) <> "" Then
            mblnIcon = True
            tmrIcon.Enabled = True
            
            Call AddIcon(picNotify.hWnd, imgIcon(0).Picture)
            DoEvents
            
            Call frmAlertMessage.ShowAlert(strData, gfrmMain)
    
        Else
            mblnIcon = False
            tmrIcon.Enabled = False
            
            Call RemoveIcon(picNotify.hWnd)
                        
        End If
        
    End Select
    
End Sub

Private Sub picNotify_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '--------------------------------------------------------------------------------------------------
    '功能:  处理picNotify的各种处理事件,主要用于自动提醒相关功能(陈渝编写)
    '--------------------------------------------------------------------------------------------------
    Dim frm As New frmAlert
    
    Select Case Hex(X) '
        Case "1E3C"     'Right-Button-Down
        Case "1E4B"     'Right-Button-Up

        Case "1830"     'Right-Button-Down LARGE FONTS '
        Case "1E1E"     'Left-Button-up
        Case "1E0F"     'Left-Button-Down '
        Case "1E2D"     'Left-Button-Double-Click '

            On Error Resume Next

            frmAlert.Show , gfrmMain
            Call frmAlert.InitData
            
        Case "1824"     'Left-Button-Double-Click LARGE FONTS
        Case "1E5A"     'Right-Button-Double-Click '
    End Select '

End Sub

Private Function SearchBind() As Long
    
    Dim lngLoop As Long
    
    SearchBind = 0
    
    DoEvents
    
    For lngLoop = 1 To 100
            
        On Error Resume Next
        
        Err = 0
        wskTest.Close
        wskTest.Protocol = sckUDPProtocol
        wskTest.Bind 1024 + lngLoop
        If Err = 0 Then
            SearchBind = 1024 + lngLoop
            wskTest.Close
            
            winSock.LocalPort = SearchBind
            Exit Function
        End If
        
        On Error GoTo 0
    Next
    
End Function

Private Function ConnectServer() As Boolean
    '--------------------------------------------------------------------------------------------------------
    '功能:
    '--------------------------------------------------------------------------------------------------------
    
    Dim rs As New ADODB.Recordset
    Dim varParam As Variant
    Dim strSQL As String
    On Error GoTo ErrH
    
    winSock.Close
    winSock.Protocol = sckUDPProtocol
    
    '格式:服务器;端口号;状态
    strSQL = "SELECT 参数值 FROM zloptions WHERE 参数号=7"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
'    rs.Open "SELECT 参数值 FROM zloptions WHERE 参数号=7", gcnOracle
    
    If rs.BOF Then Exit Function
    varParam = Split(NVL(rs("参数值").Value, ""), ";")
    If UBound(varParam) < 2 Then Exit Function
    '没有启动
    If Val(varParam(2)) <> 1 Then Exit Function
        
    Call SQLTest(App.EXEName, Me.Caption, "连接消息服务器")
    
    winSock.RemoteHost = varParam(0)            '取服务器IP地址
    winSock.RemotePort = varParam(1)            '取服务器配置的端口号
    
    Call SQLTest
    
    ConnectServer = True
    
    Exit Function
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    
End Function
