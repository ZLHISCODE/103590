VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUsers 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "用户验证"
   ClientHeight    =   6780
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   6570
   Icon            =   "frmUsers.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5160
      TabIndex        =   5
      Top             =   6000
      Width           =   1100
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   6405
      Width           =   6570
      _ExtentX        =   11589
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmUsers.frx":578A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8017
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1111
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "17:27"
            Key             =   "STANUM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraMain 
      Height          =   5985
      Left            =   0
      TabIndex        =   8
      Top             =   -60
      Width           =   6555
      Begin VB.TextBox txtDBAPwd 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   4470
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1060
         Width           =   1725
      End
      Begin VB.TextBox txtDBAUser 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1080
         TabIndex        =   7
         Text            =   "System"
         Top             =   1060
         Width           =   1725
      End
      Begin VSFlex8Ctl.VSFlexGrid vsHis 
         Height          =   4035
         Left            =   480
         TabIndex        =   4
         Top             =   1800
         Width           =   5700
         _cx             =   10054
         _cy             =   7117
         Appearance      =   2
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
         BackColor       =   16777215
         ForeColor       =   0
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   0
         BackColorSel    =   14737632
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         BackColorAlternate=   16777215
         GridColor       =   14737632
         GridColorFixed  =   12632256
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   3
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmUsers.frx":601C
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
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
         Editable        =   2
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   0   'False
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Frame fraTop 
         Height          =   120
         Left            =   15
         TabIndex        =   9
         Top             =   570
         Width           =   11280
      End
      Begin VB.Label lblHistory 
         AutoSize        =   -1  'True
         Caption         =   "历史库用户"
         Height          =   180
         Left            =   240
         TabIndex        =   3
         Top             =   1460
         Width           =   900
      End
      Begin VB.Label lblDBA 
         AutoSize        =   -1  'True
         Caption         =   "DBA用户"
         Height          =   180
         Left            =   240
         TabIndex        =   0
         Top             =   780
         Width           =   630
      End
      Begin VB.Label lblDBAPwd 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "密  码"
         Height          =   180
         Left            =   3840
         TabIndex        =   11
         Top             =   1125
         Width           =   540
      End
      Begin VB.Label lblDBAUser 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "用户名"
         Height          =   180
         Left            =   480
         TabIndex        =   1
         Top             =   1125
         Width           =   540
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "请验证如下历史库用户或DBA用户。"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   180
         TabIndex        =   10
         Top             =   225
         Width           =   3510
      End
   End
End
Attribute VB_Name = "frmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'====================================================================
'==变量
'====================================================================
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const EM_SETPASSWORDCHAR = &HCC
Private Enum HisCols
    HC_ID = 0
    HC_HisDB = 1
    HC_Server = 2
    HC_PWD = 3
End Enum

Private mrsSource               As ADODB.Recordset '初始界面所需要数据源，并记录界面项目的选择状态
Private mblnOK                  As Boolean
Private mblnCheckDBA            As Boolean

'====================================================================
'==公共接口
'====================================================================
Public Function ShowMe(Optional ByRef rsSource As ADODB.Recordset, Optional ByVal blnCheckDBA As Boolean) As Boolean
'功能：展示选择界面
'           rsSource=初始界面所需要数据源
'返回：rsSource=界面选择状态
'         ShowMe=是否退出，暂时未使用
    If Not rsSource Is Nothing Then rsSource.Filter = ""
    Set mrsSource = rsSource
    mblnCheckDBA = blnCheckDBA
    Me.Show 1
    Set rsSource = mrsSource
    ShowMe = mblnOK
End Function
'====================================================================
'==控件事件
'====================================================================

Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Me.ActiveControl Is txtDBAPwd Then
            If vsHis.Rows > vsHis.FixedRows Then
                vsHis.SetFocus
            Else
                cmdOk.SetFocus
            End If
        ElseIf Me.ActiveControl Is vsHis Then
            cmdOk.SetFocus
        Else
            cmdOk.SetFocus
        End If
'        PressKey vbKeyTab
    End If
End Sub

Private Sub Form_Load()
    Me.Caption = "用户验证-" & Server
    txtDBAPwd.Text = ""
    Call LoadData
End Sub

Private Sub Form_Resize()
    Me.Height = 7200
    Me.Width = 6660
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strTmp  As String
    
    mblnOK = True
    If Not mrsSource Is Nothing Then
        mrsSource.Filter = "验证=0"
        If mrsSource.RecordCount <> 0 Then
            Do While Not mrsSource.EOF
                strTmp = strTmp & vbNewLine & "【" & mrsSource!系统名称 & "】的表空间-" & mrsSource!名称
                mrsSource.MoveNext
            Loop

        End If
    End If
    If mblnCheckDBA And Not IsDBAOK Then
        strTmp = strTmp & vbNewLine & "DBA用户"
    End If
    If strTmp <> "" Then
        If MsgBox("以下用户未验证成功：" & strTmp & "！退出可能会使部分延迟修正无法执行，是否退出。", vbInformation + vbYesNo, gstrSysName) = vbNo Then
            Cancel = 1
        Else
            mblnOK = False
        End If
    End If
End Sub

Private Sub txtDBAPwd_GotFocus()
    Call SelAll(txtDBAPwd)
End Sub

Private Sub txtDBAPwd_Validate(Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim cnTmp As New ADODB.Connection
    Dim strErr As String
    
    On Error Resume Next
    If txtDBAPwd.Text <> "" And txtDBAUser.Text <> "" Then
        If UCase(txtDBAUser.Text) = UCase(DBAUser) And DBAPWD <> txtDBAPwd.Text And IsDBAOK Then
            MsgBox "DBA用户密码错误！", vbInformation, gstrSysName
            txtDBAPwd.Text = ""
            Cancel = True: Exit Sub
        End If
        If UCase(txtDBAUser.Text) = UCase(DBAUser) And DBAPWD = txtDBAPwd.Text And IsDBAOK Then
        
        Else
            Set cnTmp = gobjRegister.GetConnection(gstrServer, txtDBAUser.Text, txtDBAPwd.Text, False, MSODBC, strErr, False)
            If cnTmp.State = adStateClosed Then
                MsgBox strErr, vbInformation, "验证失败"
                txtDBAPwd.Text = ""
                Cancel = True: Exit Sub
            End If
            
            '检查是否DBA
            If CheckIsDBA(cnTmp) = False Then
                MsgBox "该用户不具有DBA身份！", vbExclamation, gstrSysName
                txtDBAPwd.Text = ""
                txtDBAUser.Text = ""
                txtDBAUser.SetFocus: Exit Sub
            End If
            DBAUser = txtDBAUser.Text
            DBAPWD = txtDBAPwd.Text
            IsDBAOK = True
            Call LoadData '重新加载数据
        End If
    End If
End Sub

Private Sub txtDBAUser_GotFocus()
    Call SelAll(txtDBAUser)
End Sub

Private Sub txtDBAUser_Validate(Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim cnTmp As New ADODB.Connection
    
    If UCase(txtDBAUser.Text) = UCase(DBAUser) And DBAUser <> "" Then
        txtDBAPwd.Text = DBAPWD
    Else
        txtDBAPwd.Text = ""
    End If
    If txtDBAPwd.Text <> "" And txtDBAUser.Text <> "" Then
        '因为可能大小写敏感，因此去掉大写转换
        If UCase(txtDBAUser.Text) = UCase(DBAUser) And DBAPWD <> txtDBAPwd.Text Then
            MsgBox "DBA用户密码错误！", vbInformation, gstrSysName
             Cancel = True: Exit Sub
        End If
        If UCase(txtDBAUser.Text) = UCase(DBAUser) And DBAPWD = txtDBAPwd.Text And IsDBAOK Then
            '用户没有发生变化
        Else
            Set cnTmp = gobjRegister.GetConnection(gstrServer, txtDBAUser.Text, txtDBAPwd.Text, False, MSODBC, "", False)
            If cnTmp.State = adStateClosed Then
                Cancel = True: Exit Sub
            End If
            On Error GoTo 0
            '检查是否DBA
            If CheckIsDBA(cnTmp) = False Then
                MsgBox "该用户不具有DBA身份！", vbExclamation, gstrSysName
                txtDBAUser.SetFocus: Exit Sub
            End If
            
            DBAUser = txtDBAUser.Text
            DBAPWD = txtDBAPwd.Text
            IsDBAOK = True
            Call LoadData '重新加载数据
        End If
    End If
End Sub


Private Sub vsHis_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        Select Case Col
        Case HC_PWD
            vsHis.Cell(flexcpData, Row, Col) = IIf(InStr(1, vsHis.TextMatrix(Row, Col), "*") <> 0, vsHis.Cell(flexcpData, Row, Col), vsHis.TextMatrix(Row, Col))
            vsHis.TextMatrix(Row, Col) = String(Len(vsHis.TextMatrix(Row, Col)), "*")
        End Select
        Call RefreshColor(Row)
End Sub

Private Sub vsHis_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col = HC_PWD Then
        If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
           If KeyAscii < Asc("a") Or KeyAscii > Asc("z") Then
              If KeyAscii < Asc("A") Or KeyAscii > Asc("Z") Then
                  If InStr(1, Chr(KeyAscii), "_") = 0 Then
                      If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyBack Then
                      Else
                          KeyAscii = 0
                      End If
                      Exit Sub
                  End If
              End If
           End If
        End If
        vsHis.Cell(flexcpData, Row, Col) = vsHis.Cell(flexcpData, Row, Col) & Chr(KeyAscii)
    End If
End Sub

Private Sub vsHis_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    '刘兴宏:加入
    '设置编辑的密码
    If Col = HC_PWD Then
        SendMessage vsHis.EditWindow, EM_SETPASSWORDCHAR, Asc("*"), 0
    End If
End Sub

Private Sub vsHis_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = Col <> HC_PWD
End Sub

Private Sub vsHis_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strUsername As String, strBakName As String, strPassword As String, strServer As String
    Dim cnTmp As ADODB.Connection, rsTmp As ADODB.Recordset
    Dim strFilter As String, strMaxVer As String
    Dim strDbLink As String
    
    Select Case Col
        Case HC_PWD
            If InStr(1, vsHis.EditText, "*") > 0 Then
                strPassword = vsHis.Cell(flexcpData, Row, Col)
            Else
                strPassword = vsHis.EditText
            End If
            strServer = vsHis.TextMatrix(Row, HC_Server)
            strPassword = Trim(strPassword)
            strServer = UCase(Trim(strServer))
            strBakName = UCase(Trim(vsHis.TextMatrix(Row, HC_HisDB)))
            strUsername = UCase(Trim(vsHis.Cell(flexcpData, Row, HC_HisDB)))
            strDbLink = UCase(Trim(vsHis.Cell(flexcpData, Row, HC_Server)))

            If strPassword <> "" And strUsername <> "" And strServer <> "" Then
                Set cnTmp = gobjRegister.GetConnection(strServer, strUsername, strPassword, False, OraOLEDB, "", False)
                If cnTmp.State = adStateOpen Then
                    Call RecUpdate(mrsSource, "ID=" & Val(vsHis.TextMatrix(Row, HC_ID)) & " And 验证=0", "验证", 1, "密码", strPassword)
                Else
                    Cancel = True
                    Exit Sub
                End If
            End If
    End Select
    Call LoadData '重新加载数据
    Call RefreshColor(Row)
End Sub

Private Sub LoadData()
    Dim vsTmp As VSFlexGrid
    
    txtDBAUser.Text = DBAUser
    If Not mblnCheckDBA Then
        txtDBAUser.Enabled = False
        txtDBAPwd.Enabled = False
        txtDBAPwd.Text = DBAPWD
    Else
        txtDBAUser.Enabled = True
        txtDBAPwd.Enabled = True
    End If
    If Not mrsSource Is Nothing Then mrsSource.Filter = ""
    Set vsTmp = vsHis
    If Not mrsSource Is Nothing Then mrsSource.Sort = "系统编号,当前,编号,ID"
    With vsTmp
        .Rows = .FixedRows
        If Not mrsSource Is Nothing Then
            Do While Not mrsSource.EOF
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, HC_ID) = mrsSource!Id
                .TextMatrix(.Rows - 1, HC_HisDB) = mrsSource!名称 & ""
                .Cell(flexcpData, .Rows - 1, HC_HisDB) = mrsSource!所有者 & ""
                .Cell(flexcpData, .Rows - 1, HC_PWD) = mrsSource!密码 & ""
                If mrsSource!密码 & "" <> "" Then
                    .TextMatrix(.Rows - 1, HC_PWD) = String(Len(mrsSource!密码 & ""), "*")
                End If
                .TextMatrix(.Rows - 1, HC_Server) = mrsSource!服务器 & ""
                .Cell(flexcpData, .Rows - 1, HC_Server) = mrsSource!DB连接 & ""
                .RowData(.Rows - 1) = 0
                mrsSource.MoveNext
            Loop
        End If
    End With
    
    Call RefreshColor
End Sub

Private Sub RefreshColor(Optional ByVal lngRow As Long)
    Dim i As Long
    With vsHis
        If lngRow < .FixedRows Then
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpData, i, HC_PWD) = "" Then
                    .Cell(flexcpForeColor, i, HC_HisDB, i, .Cols - 1) = &H2222B2 '火砖红
                Else
                    .Cell(flexcpForeColor, i, HC_HisDB, i, .Cols - 1) = .ForeColor
                End If
            Next
        Else
            If .Cell(flexcpData, lngRow, HC_PWD) = "" Then
                .Cell(flexcpForeColor, lngRow, HC_HisDB, lngRow, .Cols - 1) = &H2222B2 '火砖红
            Else
                .Cell(flexcpForeColor, lngRow, HC_HisDB, lngRow, .Cols - 1) = .ForeColor
            End If
        End If
    End With
End Sub

