VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "Msinet.ocx"
Begin VB.Form frmPACSDevEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "影像设备属性"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   Icon            =   "frmPACSDevEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin InetCtlsObjects.Inet Inet 
      Left            =   2640
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   2
      RemotePort      =   21
      URL             =   "ftp://"
      RequestTimeout  =   5
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4795
      TabIndex        =   24
      Top             =   3270
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3540
      TabIndex        =   23
      Top             =   3270
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   150
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   3255
      Width           =   1100
   End
   Begin VB.Frame Frame2 
      Height          =   2985
      Left            =   -100
      TabIndex        =   26
      Top             =   90
      Width           =   6250
      Begin VB.CommandButton cmdPath 
         Caption         =   "…"
         Height          =   240
         Left            =   5520
         TabIndex        =   29
         TabStop         =   0   'False
         ToolTipText     =   "选择项目(*)"
         Top             =   2040
         Width           =   300
      End
      Begin VB.TextBox txtItem 
         Height          =   300
         Index           =   9
         Left            =   3780
         MaxLength       =   20
         TabIndex        =   19
         Top             =   1618
         Width           =   2055
      End
      Begin VB.TextBox txtItem 
         Height          =   300
         Index           =   8
         Left            =   1275
         MaxLength       =   20
         TabIndex        =   17
         Top             =   1618
         Width           =   1455
      End
      Begin VB.CheckBox chkOffLine 
         Caption         =   "离线设备"
         Height          =   285
         Left            =   1230
         TabIndex        =   28
         Top             =   2430
         Width           =   1335
      End
      Begin VB.TextBox txtItem 
         Height          =   300
         Index           =   7
         Left            =   1260
         MaxLength       =   100
         TabIndex        =   21
         ToolTipText     =   "Ftp目录在服务器上的本地路径"
         Top             =   2010
         Width           =   4575
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "连接测试(&T)"
         Height          =   350
         Left            =   4530
         TabIndex        =   22
         Top             =   2400
         Width           =   1215
      End
      Begin VB.ComboBox cboRoom 
         Height          =   300
         IMEMode         =   3  'DISABLE
         ItemData        =   "frmPACSDevEdit.frx":000C
         Left            =   1275
         List            =   "frmPACSDevEdit.frx":001F
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   472
         Width           =   1500
      End
      Begin VB.TextBox txtItem 
         Height          =   300
         Index           =   0
         Left            =   1275
         MaxLength       =   3
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   90
         Width           =   1455
      End
      Begin VB.TextBox txtItem 
         Height          =   300
         Index           =   1
         Left            =   3795
         MaxLength       =   20
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   90
         Width           =   2055
      End
      Begin VB.TextBox txtItem 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   6
         Left            =   3795
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   15
         Top             =   1260
         Width           =   2055
      End
      Begin VB.TextBox txtItem 
         Height          =   300
         Index           =   3
         Left            =   1275
         MaxLength       =   4
         TabIndex        =   9
         Top             =   854
         Width           =   1455
      End
      Begin VB.TextBox txtItem 
         Height          =   300
         Index           =   4
         Left            =   3795
         MaxLength       =   100
         TabIndex        =   11
         Top             =   870
         Width           =   2055
      End
      Begin VB.TextBox txtItem 
         Height          =   300
         Index           =   2
         Left            =   3795
         MaxLength       =   15
         TabIndex        =   7
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox txtItem 
         Height          =   300
         Index           =   5
         Left            =   1275
         MaxLength       =   20
         TabIndex        =   13
         Top             =   1236
         Width           =   1455
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "设备AE(&D)"
         Height          =   180
         Index           =   5
         Left            =   2920
         TabIndex        =   18
         Top             =   1645
         Width           =   810
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "本地AE(&L)"
         Height          =   180
         Index           =   2
         Left            =   420
         TabIndex        =   16
         Top             =   1645
         Width           =   810
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "本地路径(&L)"
         Height          =   180
         Index           =   0
         Left            =   210
         TabIndex        =   20
         Top             =   2070
         Width           =   990
      End
      Begin VB.Label lblRoom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "类型(&S)"
         Height          =   180
         Left            =   600
         TabIndex        =   4
         Top             =   540
         Width           =   630
      End
      Begin VB.Label Label6 
         Caption         =   "设备号(&H)"
         Height          =   255
         Left            =   420
         TabIndex        =   0
         Top             =   150
         Width           =   975
      End
      Begin VB.Label Label3 
         Height          =   135
         Left            =   0
         TabIndex        =   27
         Top             =   -20
         Width           =   6255
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "设备名(&M)"
         Height          =   180
         Index           =   8
         Left            =   2940
         TabIndex        =   2
         Top             =   150
         Width           =   810
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "密码(&C)"
         Height          =   180
         Index           =   6
         Left            =   3100
         TabIndex        =   14
         Top             =   1290
         Width           =   630
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "端口(&P)"
         Height          =   180
         Index           =   1
         Left            =   600
         TabIndex        =   8
         Top             =   935
         Width           =   630
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "Ftp目录(&F)"
         Height          =   180
         Index           =   4
         Left            =   2840
         TabIndex        =   10
         Top             =   935
         Width           =   900
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "IP地址(&A)"
         Height          =   180
         Index           =   3
         Left            =   2940
         TabIndex        =   6
         Top             =   540
         Width           =   810
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "用户名(&N)"
         Height          =   180
         Index           =   7
         Left            =   420
         TabIndex        =   12
         Top             =   1290
         Width           =   810
      End
   End
End
Attribute VB_Name = "frmPACSDevEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************API调用*****************************************
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const MAX_PATH = 260
Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As String
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
'***********************************************************************************
Private strDeviceNO As String
Private ifOK As Boolean

Public Function ShowMe(objParent As Object, ByVal DeviceNO As String) As Boolean
    strDeviceNO = DeviceNO
    
    Me.Show vbModal, objParent
    ShowMe = ifOK
End Function

Private Sub cboRoom_Click()
    Dim blnEnabled As Boolean
    
    blnEnabled = True
    If cboRoom.ListIndex > 0 Then blnEnabled = False
    Me.txtItem(3).Enabled = Not blnEnabled
    Me.txtItem(4).Enabled = blnEnabled
    Me.txtItem(5).Enabled = blnEnabled
    Me.txtItem(6).Enabled = blnEnabled
    Me.txtItem(8).Enabled = Not blnEnabled
    Me.txtItem(9).Enabled = Not blnEnabled
    Me.txtItem(7).Enabled = cboRoom.ListIndex = 0
    Me.lblItem(0).Caption = "本地路径(&L)"
    If cboRoom.ListIndex <> 0 Then txtItem(7) = ""
    
    If cboRoom.ListIndex = cboRoom.ListCount - 1 Then
        
        Me.txtItem(2).Enabled = False
        Me.txtItem(3).Enabled = False
        Me.txtItem(4).Enabled = False
        Me.txtItem(5).Enabled = True
        Me.txtItem(6).Enabled = True
        Me.txtItem(8).Enabled = False
        Me.txtItem(9).Enabled = False
        Me.txtItem(7).Enabled = True
        Me.lblItem(0).Caption = "远程路径(&L)"
        Me.cmdPath.Visible = True
    Else
        Me.cmdPath.Visible = False
        Me.txtItem(2).Enabled = True
    End If
'    Me.lblItem(4).Caption = IIf(cboRoom.ListIndex = 2, "网络路径(&F)", "Ftp目录(&F)")
    
    
End Sub

Private Sub cboRoom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmd_Click()

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    Dim strSQL As String
    
    On Error GoTo DBError
    If Len(Trim(txtItem(0))) = 0 Then
        MsgBox "请输入设备号！", vbInformation, gstrSysName
        txtItem(0).SetFocus: Exit Sub
    End If
    If Len(Trim(txtItem(1))) = 0 Then
        MsgBox "请输入设备名！", vbInformation, gstrSysName
        txtItem(1).SetFocus: Exit Sub
    End If
    If Len(Trim(txtItem(2))) = 0 And Me.cboRoom.ListIndex <> cboRoom.ListCount - 1 Then
        MsgBox "请输入IP地址！", vbInformation, gstrSysName
        txtItem(2).SetFocus: Exit Sub
    End If
    If cboRoom.ListIndex = 1 And (Len(Trim(txtItem(3))) = 0 Or Not IsNumeric(txtItem(3))) Then
        MsgBox "请输入正确的端口号！", vbInformation, gstrSysName
        txtItem(3).SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(txtItem(1).Text), vbFromUnicode)) > txtItem(1).MaxLength Then
        MsgBox "设备名超长（最多" & txtItem(1).MaxLength & "个字符或" & CInt(txtItem(1).MaxLength / 2) & "个汉字）！", vbInformation, gstrSysName
        txtItem(1).SetFocus: Exit Sub
    End If

    strSQL = "zl_影像设备目录_Update('" & txtItem(0) & "','" & Trim(txtItem(1)) & "'," & cboRoom.ListIndex + 1 & _
        ",'" & Trim(txtItem(2)) & "','" & Trim(txtItem(4)) & "','" & Trim(txtItem(3)) & "','" & Trim(txtItem(5)) & "','" & Trim(txtItem(6)) & "','" & _
        Trim(txtItem(8)) & "','" & Trim(txtItem(9)) & "','" & Trim(txtItem(7)) & "'," & chkOffLine.Value & ")"
    ExecuteProc strSQL, Me.Caption
    
    ifOK = True
    Unload Me
    Exit Sub
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdPath_Click()
    Dim strTmp As String
    '得到路径
    strTmp = BrowPath(Me.hwnd, "请选定保存的文件目录：")
    '当用新的路径时才保存
    If strTmp <> "" And strTmp <> Me.txtItem(7).Text Then
        Me.txtItem(7).Text = strTmp
    End If
End Sub

Private Sub cmdTest_Click()
    Dim objGlobal As New DicomGlobal
    Dim vTmpData As Variant
    Dim strUse As String
    Dim strPas As String
    Dim strPath As String
    Dim duTime As Double
    On Error GoTo TestError
    
    If cboRoom.ListIndex <> 0 Then '测试影像网关
        If cboRoom.ListIndex = cboRoom.ListCount - 1 Then
            If Len(Trim(Me.txtItem(7).Text)) <= 0 Then
                MsgBox "请选择或输入要访问的远程主机", vbInformation, gstrSysName
                Me.txtItem(7).SetFocus
                Exit Sub
            End If
            cmdTest.Enabled = False
            If Len(Trim(Me.txtItem(6).Text)) > 0 Then
                strPas = Trim(Me.txtItem(6).Text)
                strUse = Trim(Me.txtItem(5).Text)
                strPath = Trim(Me.txtItem(7).Text)
                Shell "net use " & strPath & " " & strPas & " /user:" & strUse, vbHide
            End If
            duTime = Timer
            Do Until CLng(Timer - duTime) >= 20
                Shell "net use " & strPath & " " & strPas & " /user:" & strUse, vbHide
                If WriteTest(False) = True Then
                    Exit Do
                End If
                DoEvents
            Loop
            If WriteTest(False) = False Then
                MsgBox "无法连接到指定的接收主机！", vbInformation, gstrSysName
            Else
                MsgBox "连接测试成功！", vbInformation, gstrSysName
            End If
            Shell "net use " & strPath & " /delete "
            Me.MousePointer = vbDefault: cmdTest.Enabled = True
        Else
            If Len(Trim(txtItem(2))) = 0 Then
                MsgBox "请输入IP地址！", vbInformation, gstrSysName
                txtItem(2).SetFocus: Exit Sub
            End If
            If Len(Trim(txtItem(3))) = 0 Or Not IsNumeric(txtItem(3)) Then
                MsgBox "请输入正确的端口号！", vbInformation, gstrSysName
                txtItem(3).SetFocus: Exit Sub
            End If
            With objGlobal
                Me.MousePointer = vbHourglass: cmdTest.Enabled = False
                If .Echo(txtItem(2), CLng(txtItem(3)), "", "") <> 0 Then
                    MsgBox "无法连接到指定的接收主机！", vbInformation, gstrSysName
                    txtItem(2).SetFocus
                Else
                    MsgBox "连接测试成功！", vbInformation, gstrSysName
                End If
                Me.MousePointer = vbDefault: cmdTest.Enabled = True
            End With
        End If
    Else
        With Inet
'            .AccessType = 0: .URL = "ftp://" & IIf(Len(Trim(txtItem(5))) = 0, "", _
'                Trim(txtItem(5)) & IIf(Len(Trim(txtItem(6))) = 0, "", ":" & Trim(txtItem(6))) & "@") & txtItem(2)
'
'            Me.MousePointer = vbHourglass: cmdTest.Enabled = False
'            .Execute , "MKDIR /" & IIf(Len(Trim(txtItem(4))) = 0, "", Trim(txtItem(4)) & "/") & "Tmp"
'            Do While .StillExecuting
'                DoEvents
'            Loop
'            .Execute , "SIZE /" & IIf(Len(Trim(txtItem(4))) = 0, "", Trim(txtItem(4)) & "/")
'            Do While .StillExecuting
'                DoEvents
'            Loop
'            vTmpData = .GetChunk(1024)
'            If Len(vTmpData) > 0 Then
'                MsgBox "连接测试成功！", vbInformation, gstrSysName
'            Else
'                MsgBox "无法连接到指定的设备！", vbInformation, gstrSysName
'                txtItem(2).SetFocus
'            End If
'            .Execute , "RMDIR /" & IIf(Len(Trim(txtItem(4))) = 0, "", Trim(txtItem(4)) & "/") & "Tmp"
'            Do While .StillExecuting
'                DoEvents
'            Loop
'            .Execute , "CLOSE"
'
'            .Cancel
            
            Me.MousePointer = vbHourglass: cmdTest.Enabled = False
            .URL = "ftp://" & txtItem(2).Text & IIf(Len(Trim(txtItem(4))) > 0, "/" & txtItem(4), "")
            .UserName = Me.txtItem(5).Text
            .Password = Me.txtItem(6).Text
            .Execute
            MsgBox "连接测试成功！", vbInformation, gstrSysName
            
            Me.MousePointer = vbDefault: cmdTest.Enabled = True
        End With
    End If
    Exit Sub
TestError:
    Inet.Cancel
    Me.MousePointer = vbDefault: cmdTest.Enabled = True
    
    MsgBox "无法连接到指定的设备！", vbInformation, gstrSysName
    txtItem(2).SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyEscape Then Exit Sub
    Call cmdCancel_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    ifOK = False
    
    On Error GoTo DBError
    If Len(Trim(strDeviceNO)) > 0 Then
        strSQL = "Select 设备号,设备名,Nvl(类型,1) As 类型,IP地址,端口号,Ftp目录,用户名,密码,本机目录,Nvl(状态,0),本地AE,设备AE" & _
            " From 影像设备目录 Where 设备号='" & strDeviceNO & "'"
        OpenRecord rsTmp, strSQL, Me.Caption
        Me.txtItem(0) = rsTmp(0)
        Me.txtItem(1) = rsTmp(1)
        Me.cboRoom.ListIndex = rsTmp(2) - 1
        Me.txtItem(2) = Nvl(rsTmp(3))
        Me.txtItem(3) = Nvl(rsTmp(4))
        Me.txtItem(4) = Nvl(rsTmp(5))
        Me.txtItem(5) = Nvl(rsTmp(6))
        Me.txtItem(6) = Nvl(rsTmp(7))
        Me.txtItem(7) = Nvl(rsTmp(8))
        Me.chkOffLine.Value = rsTmp(9)
        Me.txtItem(8) = Nvl(rsTmp(10))
        Me.txtItem(9) = Nvl(rsTmp(11))
        
        Me.txtItem(0).Enabled = False: cboRoom.Enabled = False
    Else
        Me.txtItem(0) = GetNewNo
        Me.cboRoom.ListIndex = 0
        Me.chkOffLine.Value = 0
    End If
    
    Exit Sub
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Inet.Cancel
End Sub

Private Sub txtItem_GotFocus(Index As Integer)
    With Me.txtItem(Index)
        .SelStart = 0: .SelLength = .MaxLength
    End With
    Select Case Index
        Case 1
            Call zlCommFun.OpenIme(True)
        Case Else
            Call zlCommFun.OpenIme(False)
    End Select
End Sub

Private Sub txtItem_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtItem_KeyPress(Index As Integer, KeyAscii As Integer)
    If ifEditKey(KeyAscii, False) Then Exit Sub
    
    If LenB(StrConv(Trim(txtItem(Index).Text), vbFromUnicode)) >= txtItem(Index).MaxLength Then
        KeyAscii = 0
        Exit Sub
    End If
    Select Case Index
        Case 0, 3
            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then KeyAscii = 0
    End Select
End Sub

Private Sub txtItem_LostFocus(Index As Integer)
    Dim objFileSystem As New Scripting.FileSystemObject
    Select Case Index
        Case 1
            Call zlCommFun.OpenIme(False)
        Case 7 '本机目录
            If Len(Trim(txtItem(7))) > 0 Then
                If Me.cboRoom.ListIndex <> Me.cboRoom.ListCount - 1 Then
                    txtItem(7) = objFileSystem.GetAbsolutePathName(txtItem(7))
                End If
            End If
    End Select
End Sub

'判断是否为编辑键
Private Function ifEditKey(ByVal KeyAscii As Integer, Optional ByVal AllowSubtract As Boolean = True) As Boolean
    If KeyAscii = vbKeyBack Or (KeyAscii = vbKeyInsert And AllowSubtract) Or KeyAscii = vbKeyDelete Or _
      KeyAscii = vbKeyHome Or KeyAscii = vbKeyEnd Or KeyAscii = vbKeyLeft Or KeyAscii = vbKeyRight Or _
      KeyAscii = vbKeyEscape Or KeyAscii = vbKeyReturn Then
        ifEditKey = True
    Else
        ifEditKey = False
    End If
End Function

Private Function GetNewNo() As String
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo DBError
    strSQL = "Select Nvl(Max(To_Char(设备号,'000')),1) From 影像设备目录"
    OpenRecord rsTmp, strSQL, Me.Caption
    If rsTmp.EOF Then
        GetNewNo = "001"
    Else
        GetNewNo = Format(Val(rsTmp(0)) + 1, "000")
    End If
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'显示保存目录
Private Function BrowPath(lWindowHwnd As Long, Optional ByVal sTitle As String = "") As String
    Dim iNull As Integer, lpIDList As Long
    Dim sPath As String, udtBI As BrowseInfo
    On Error GoTo OpenFileError
    With udtBI
        '设置浏览窗口
        .hWndOwner = lWindowHwnd
        '返回选中的目录
        .ulFlags = BIF_RETURNONLYFSDIRS
        If sTitle = "" Then
            .lpszTitle = "请选定开始搜索的文件夹："
        Else
            .lpszTitle = sTitle
        End If
    End With
    '调出浏览窗口
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        '获取路径
        SHGetPathFromIDList lpIDList, sPath
        '释放内存
        CoTaskMemFree lpIDList
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    End If
    BrowPath = sPath
    Exit Function
OpenFileError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function WriteTest(ShowErrMsg As Boolean) As Boolean
    Dim strTmpPath As String
    On Error GoTo CopyError
    strTmpPath = IIf(Len(App.Path) > 3, App.Path & "\", App.Path) & "temp.txt"
    Open strTmpPath For Output As #1
    Close #1
    FileCopy strTmpPath, IIf(Len(Me.txtItem(7).Text) > 3, Me.txtItem(7).Text & "\", Me.txtItem(7).Text) & "temp.txt"
    Kill IIf(Len(Me.txtItem(7).Text) > 3, Me.txtItem(7).Text & "\", Me.txtItem(7).Text) & "temp.txt"
    Kill strTmpPath
    WriteTest = True
    Exit Function
CopyError:
    If ShowErrMsg = False Then Exit Function
    If Err.Number = 75 Then
        MsgBox "写入测试失败!请查看[" & Me.txtItem(7).Text & "]是否有写入权限!", vbInformation, App.EXEName
    Else
        MsgBox "发生其他错误！", vbQuestion, App.EXEName
    End If
End Function
