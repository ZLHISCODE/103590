VERSION 5.00
Begin VB.Form frmSet神木大兴 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtPath 
      Height          =   300
      Left            =   1695
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   2955
      Width           =   3675
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "…"
      Height          =   300
      Left            =   5400
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2955
      Width           =   285
   End
   Begin VB.Frame fra医保服务器 
      Height          =   1875
      Index           =   0
      Left            =   270
      TabIndex        =   5
      Top             =   870
      Width           =   5595
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   1200
         MaxLength       =   40
         TabIndex        =   9
         Top             =   555
         Width           =   3075
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1200
         MaxLength       =   40
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   945
         Width           =   3075
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1200
         MaxLength       =   40
         TabIndex        =   7
         Top             =   1335
         Width           =   3075
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "测试(&T)"
         Height          =   1095
         Left            =   4515
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   555
         Width           =   1005
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "用户名(&U)"
         Height          =   180
         Index           =   0
         Left            =   330
         TabIndex        =   12
         Top             =   615
         Width           =   810
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "密码(&P)"
         Height          =   180
         Index           =   1
         Left            =   510
         TabIndex        =   11
         Top             =   1005
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "服务器(&S)"
         Height          =   180
         Index           =   2
         Left            =   330
         TabIndex        =   10
         Top             =   1395
         Width           =   810
      End
   End
   Begin VB.Frame fra 
      Height          =   45
      Index           =   1
      Left            =   -30
      TabIndex        =   3
      Top             =   3435
      Width           =   7665
   End
   Begin VB.Frame fra 
      Height          =   45
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   660
      Width           =   7665
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3540
      TabIndex        =   0
      Top             =   3615
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4725
      TabIndex        =   1
      Top             =   3615
      Width           =   1100
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "病人信息所在目录"
      Height          =   180
      Index           =   0
      Left            =   195
      TabIndex        =   15
      Top             =   3030
      Width           =   1440
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "设置配置中间库的的相关信息，以便能随利进行数据交换."
      Height          =   180
      Left            =   645
      TabIndex        =   4
      Top             =   390
      Width           =   4590
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   150
      Picture         =   "frmSet神木大兴.frx":0000
      Top             =   150
      Width           =   480
   End
End
Attribute VB_Name = "frmSet神木大兴"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mcnTest As New ADODB.Connection
Private mblnChange As Boolean
Private mblnFirst As Boolean
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private Enum enum文本
    text医保用户 = 0
    Text医保密码 = 1
    Text医保服务器 = 2
End Enum

Public Function 参数设置() As Boolean
    mblnChange = False
    frmSet神木大兴.Show vbModal, frm医保类别
    参数设置 = mblnOK
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdSel_Click()
    Dim strPath As String
    strPath = OpenDire(Me, "请指定目录：")
    If strPath = "" Then Exit Sub
    txtPath.Text = strPath
End Sub

Private Sub cmdTest_Click()
    Dim rsTemp As New ADODB.Recordset
    If mcnTest.State = adStateOpen Then mcnTest.Close
    
    If OraDataOpen(mcnTest, txtEdit(Text医保服务器).Text, txtEdit(text医保用户).Text, txtEdit(Text医保密码).Tag) = False Then
        Exit Sub
    End If
    
    MsgBox "连接成功！", vbInformation, gstrSysName
End Sub


Private Sub Form_Activate()
    Dim rsTemp As New ADODB.Recordset
    If mblnFirst = False Then Exit Sub
    
    mblnFirst = False
    gstrSQL = "Select * From 保险参数 where 险类=" & TYPE_陕西大兴
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    With rsTemp
        Do While Not .EOF
            Select Case Nvl(!参数名)
            Case "医保用户名"
                  txtEdit(text医保用户).Text = Nvl(!参数值)
            Case "医保用户密码"
                  txtEdit(Text医保密码).Text = Nvl(!参数值)
            Case "医保服务器"
                  txtEdit(Text医保服务器).Text = Nvl(!参数值)
            End Select
            .MoveNext
        Loop
    End With
   txtPath.Text = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("病人目录"), App.Path)
 End Sub
Private Function OpenDire(odtvOwner As Form, Optional odtvTitle As String) As String
   Dim lpIDList As Long
   Dim sBuffer As String
   Dim szTitle As String
   Dim tBrowseInfo As BrowseInfo
   szTitle = odtvTitle
   With tBrowseInfo
      .hwndOwner = odtvOwner.hwnd
      .lpszTitle = lstrcat(szTitle, "")
      .ulFlags = BIF_RETURNONLYFSDIRS ' + BIF_DONTGOBELOWDOMAIN
   End With
   lpIDList = SHBrowseForFolder(tBrowseInfo)
   If (lpIDList) Then
      sBuffer = Space(MAX_PATH)
      SHGetPathFromIDList lpIDList, sBuffer
      sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
      OpenDire = sBuffer
   End If
End Function

Private Sub Form_Load()
    mblnFirst = True
End Sub

Private Sub lblODBC_Click(Index As Integer)
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index = Text医保密码 Then
        txtEdit(Index).Tag = txtEdit(Index).Text
    End If
    If Index = Text医保服务器 Or Index = Text医保密码 Or Index = text医保用户 Then
        '关闭对医保服务器的连接，因为在参数设置完成时需要重新打开
        If mcnTest.State = adStateOpen Then mcnTest.Close
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Private Sub cmdOK_Click()
    If IsValid = False Then Exit Sub
    If SaveData = False Then Exit Sub
    
    mblnOK = True
    mblnChange = False
    Unload Me
End Sub

Private Function IsValid() As Boolean
    Dim lngCount As Long
    Dim strTitle As String
    Dim rsTemp As New ADODB.Recordset
    
    
    For lngCount = txtEdit.LBound To txtEdit.UBound
        If zlCommFun.StrIsValid(txtEdit(lngCount).Text, txtEdit(lngCount).MaxLength) = False Then
            zlControl.TxtSelAll txtEdit(lngCount)
            txtEdit(lngCount).SetFocus
            Exit Function
        End If
    Next
    
    If mcnTest.State = adStateClosed Then
        If OraDataOpen(mcnTest, txtEdit(Text医保服务器).Text, txtEdit(text医保用户).Text, txtEdit(Text医保密码).Tag, False) = False Then
            If MsgBox("医保服务器不能正常连接，是否继续？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End If
        
    IsValid = True
End Function

Private Function SaveData() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim lst As ListItem
    
    gcnOracle.BeginTrans
    On Error GoTo errHandle
    
    
    '删除已经数据
    gstrSQL = "zl_保险参数_Delete(" & TYPE_陕西大兴 & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '新增参数数据
    gstrSQL = "zl_保险参数_Insert(" & TYPE_陕西大兴 & ",null,'医保用户名','" & txtEdit(text医保用户).Text & "',1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_陕西大兴 & ",null,'医保用户密码','" & txtEdit(Text医保密码).Tag & "',2)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_陕西大兴 & ",null,'医保服务器','" & txtEdit(Text医保服务器).Text & "',3)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gcnOracle.CommitTrans
    
    SaveSetting "ZLSOFT", "公共模块\zl9Insure", UCase("病人目录"), Trim(txtPath.Text)
    SaveData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    gcnOracle.RollbackTrans
End Function
Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub
