VERSION 5.00
Object = "{4CB25E99-6F8F-48B9-9C58-7CD01371DDA3}#380.0#0"; "ztsi.ocx"
Begin VB.Form frmConn昭通 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5325
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Visible         =   0   'False
   Begin VB.Timer TimeAvi 
      Interval        =   50
      Left            =   30
      Top             =   1590
   End
   Begin ztsi.AppClient AppClient 
      Height          =   585
      Left            =   30
      TabIndex        =   3
      Top             =   1470
      Visible         =   0   'False
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   1032
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3990
      TabIndex        =   1
      Top             =   1455
      Width           =   1100
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   75
      Left            =   0
      Picture         =   "frmConn昭通.frx":0000
      ScaleHeight     =   5
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   355
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1080
      Width           =   5325
   End
   Begin VB.Label LblNote 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "正在与中心交换数据，请稍候......"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1215
      TabIndex        =   2
      Top             =   495
      Width           =   2880
   End
End
Attribute VB_Name = "frmConn昭通"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strReturnInfo As String, mlngRows As Long
Private strReturnData As String, bytData() As Byte, blnIsConnect As Boolean, blnFlag As Boolean, blnIsGet As Boolean

Public Function ReadCard(ByVal IntPort As Integer) As String
    On Error GoTo errHand
    Call ShowWindow(Me.hwnd, 9)
    ReadCard = AppClient.ReadCard(IntPort)
    Call ShowWindow(Me.hwnd, 0)
    Exit Function

errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call ShowWindow(Me.hwnd, 0)
End Function

Public Function readPassword(ByVal IntPort As Integer) As String
    On Error GoTo errHand
    Call ShowWindow(Me.hwnd, 9)
    readPassword = AppClient.readPassword(IntPort)
    Call ShowWindow(Me.hwnd, 0)
    Exit Function

errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call ShowWindow(Me.hwnd, 0)
End Function

Public Function ConnCenter(ByVal strServerIP As String, ByVal lngServerPort As Integer, ByVal strCN As String, Optional ByVal HisUserID As Long = 0) As Boolean
    Dim rsUser As New ADODB.Recordset
    Dim strUser As String, strPass As String
    Dim strDataSoure As String
    On Error GoTo errHand
    
    Call ShowWindow(Me.hwnd, 9)
    
    If gcn昭通.State = 0 Then
        strDataSoure = Mid(gcnOracle.ConnectionString, InStr(UCase(gcnOracle.ConnectionString), "SERVER=") + 7)
        strDataSoure = Left(strDataSoure, InStr(strDataSoure, """;") - 1)
        
        gcn昭通.ConnectionString = "Provider=MSDAORA.1;Password=his;User ID=ybuser;Data Source=" & strDataSoure & ";Persist Security Info=True"
        gcn昭通.CursorLocation = adUseClient
        gcn昭通.Open
    End If
    
    '表“tab_czry”存放登录医保使用的用户名和密码
    If HisUserID = 0 Then
        strUser = "00"
        strPass = "123456"
        gstrSQL = "Select * From tab_czry Where OPER=0"
        Set rsUser = gcn昭通.Execute(gstrSQL)
        If rsUser.RecordCount <> 0 Then
            strPass = rsUser!Password
        End If
    Else
        gstrSQL = "Select * From tab_czry Where HISID=" & HisUserID
        Set rsUser = gcn昭通.Execute(gstrSQL)
        If rsUser.EOF Then
            MsgBox "用户没有使用医保的权限", vbInformation, "权限错误"
            Call ShowWindow(Me.hwnd, 0)
            Exit Function
        End If
        strUser = rsUser!oper
        strPass = rsUser!Password
    End If
    ConnCenter = AppClient.Login(strServerIP, lngServerPort, strCN, strUser, strPass)
    
    If ConnCenter = False Then
        MsgBox AppClient.getMessages, vbInformation, gstrSysName
    End If
    Call ShowWindow(Me.hwnd, 0)
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call ShowWindow(Me.hwnd, 0)
End Function

Public Function Query(lngRowNum As Long, lngRows As Long, Optional strMessage As String = "") As Boolean
    Dim arrData() As String
    Dim intCol As Integer, intCols As Integer
    On Error GoTo errHand
    
    Call ShowWindow(Me.hwnd, 9)
    
    arrData = AppClient.getResultSetARow(lngRowNum)
    '组织成以前的格式（格式：;是行分隔符；,是列分隔符）
    strReturnInfo = ""
    intCols = UBound(arrData)
    For intCol = 0 To intCols
        strReturnInfo = strReturnInfo & IIf(strReturnInfo = "", "", vbTab) & arrData(intCol)
    Next
    
    Query = True
    Call ShowWindow(Me.hwnd, 0)
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call ShowWindow(Me.hwnd, 0)
End Function

Public Function Execute(str交易码 As String, lng子码 As Long, str参数 As String, str提示信息 As String) As Boolean
    Dim intReturn As Integer
    Dim intRow As Integer
    On Error GoTo errHand
    
    Call ShowWindow(Me.hwnd, 9)
    
    intReturn = AppClient.executeTrade(str交易码, CInt(lng子码), str参数)
    Execute = (intReturn = 0)
    
    If Execute = False Then
        MsgBox "错误信息：" & AppClient.getMessages & vbCrLf & "入参：" & str参数, vbInformation, gstrSysName
    Else
        '获取总行数
        mlngRows = AppClient.GetRows
        '获取的返回值（格式：chr(10)是行分隔符；vbtab是列分隔符）
        strReturnInfo = AppClient.getResultSet(0, 1)
    
    End If
    Call ShowWindow(Me.hwnd, 0)
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call ShowWindow(Me.hwnd, 0)
End Function

Public Function ConnClose() As Boolean
    On Error GoTo errHand
    
    Call ShowWindow(Me.hwnd, 9)
    Call AppClient.logout
    
    ConnClose = True
    Unload Me
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Unload Me
End Function

Private Sub cmdCancel_Click()
    If MsgBox("交易未成功执行，是否取消？", vbQuestion + vbYesNo, "取消交易") = vbYes Then
        WriteInfo "取消医保交易"
        
        On Error Resume Next
        Call ShowWindow(Me.hwnd, 0)
        Exit Sub
    End If
End Sub

Private Sub TimeAvi_Timer()
    Static i As Long
    i = i + 20
    If i >= Picture1.ScaleWidth Then i = 1
    
    Picture1.PaintPicture Picture1.Picture, i, 0, Picture1.ScaleWidth - i, Picture1.ScaleHeight, 0, 0, Picture1.ScaleWidth - i, Picture1.ScaleHeight
    Picture1.PaintPicture Picture1.Picture, 0, 0, i, Picture1.ScaleHeight, Picture1.ScaleWidth - i, 0, i, Picture1.ScaleHeight
End Sub
'
'Private Function Byt2Long(bytInData() As Byte, lngStart As Long, lngLen As Long) As Long
'    Dim lngLoop As Long, strTemp As String
'    strTemp = ""
'    For lngLoop = lngStart To lngStart + lngLen - 1
'        strTemp = strTemp & Right("00" & hex(bytInData(lngLoop)), 2)
'    Next
'    Byt2Long = CLng("&H" & strTemp)
'End Function
'
'Private Function L2S(lngInData As Long, intStart As Integer) As String
'    Dim strTemp As String, bytTemp(3) As Byte
'    strTemp = Right("00000000" & hex(lngInData), 8)
'    bytData(intStart) = CLng("&H" & Mid(strTemp, 1, 2))
'    bytData(intStart + 1) = CLng("&H" & Mid(strTemp, 3, 2))
'    bytData(intStart + 2) = CLng("&H" & Mid(strTemp, 5, 2))
'    bytData(intStart + 3) = CLng("&H" & Mid(strTemp, 7, 2))
'End Function
'
'Private Sub wsckConn_Connect()
'    blnIsConnect = True
'End Sub
'
'Private Sub wsckConn_DataArrival(ByVal bytesTotal As Long)
'    Dim strFlag As String, lngCode As Long, lngRows As Long, lngType As Long, lngInfoLen As Long, _
'        strInfo As String, bytReturnData() As Byte, lngTheLen As Long, bytTemp As Byte
'    wsckConn.GetData strReturnData, vbString, 1
'    If strReturnData <> "R" And strReturnData <> "S" Then
'        MsgBox "发生错误：返回信息格式错误", vbInformation, "错误"
'        WriteInfo "医保交易错误:返回信息格式错误[0x" & hex(asc(strReturnData)) & "]"
'        blnFlag = False
'        Exit Sub
'    End If
'    If strReturnData = "R" Then
'        wsckConn.GetData bytReturnData(), vbArray + vbByte, 13
'        lngCode = Byt2Long(bytReturnData, 0, 4)
'        lngRows = Byt2Long(bytReturnData, 4, 4)
'        lngType = bytReturnData(8)
'        lngInfoLen = Byt2Long(bytReturnData, 9, 4)
'
'        wsckConn.GetData strInfo, vbString, lngInfoLen
'
'        mlngRows = lngRows
'        If lngCode <> 0 And Len(strInfo) <> 0 Then
'            MsgBox "发生错误：" & vbCrLf & "    " & strInfo & ";错误编号:" & lngCode, vbInformation, "错误"
'            WriteInfo "医保交易错误:" & strInfo
'            WriteInfo "错误编号:" & lngCode
'            blnFlag = False
'        ElseIf lngCode <> 0 Then
'            MsgBox "发生错误，错误信息未返回", vbInformation, "错误"
'            WriteInfo "医保交易错误,未返回错误信息"
'            blnFlag = False
'        Else
'            strReturnInfo = strInfo
'            WriteInfo "医保交易结果:" & strInfo
'            WriteInfo "结果集行数：" & lngRows
'            blnFlag = True
'        End If
'        On Error Resume Next
'        Call ShowWindow(Me.hwnd, 0)
'        unload me
'    Else
'        wsckConn.GetData bytReturnData(), vbArray + vbByte, 4
'        lngInfoLen = Byt2Long(bytReturnData, 0, 4)
'        wsckConn.GetData strInfo, vbString, lngInfoLen
'        strReturnInfo = strInfo
'        WriteInfo "医保交易结果:" & strInfo
'        blnFlag = True
'    End If
'    blnIsGet = True
'End Sub
'
'Private Sub wsckConn_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'    MsgBox "发生错误：" & Description, vbInformation, "连接"
'    WriteInfo "网络连接错误:" & Description
'    wsckConn.Close
'    blnIsConnect = False
'    blnFlag = False
''    SetPos Me.hwnd, True
'    On Error Resume Next
'    Call ShowWindow(Me.hwnd, 0)
'    unload me
'End Sub
Private Sub Form_Load()

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.Top = Screen.Height / 2 - Me.Height / 2
    Me.Left = Screen.Width / 2 - Me.Width / 2
End Sub
