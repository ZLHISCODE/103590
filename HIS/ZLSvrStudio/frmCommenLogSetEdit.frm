VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCommenLogSetEdit 
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "日志分类"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5520
   Icon            =   "frmCommenLogSetEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chkAppAll 
      BackColor       =   &H80000005&
      Caption         =   "应用至所有分类"
      Height          =   180
      Index           =   3
      Left            =   3600
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   3480
      Width           =   1620
   End
   Begin VB.CheckBox chkAppAll 
      BackColor       =   &H80000005&
      Caption         =   "应用至所有分类"
      Height          =   180
      Index           =   2
      Left            =   3600
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   3000
      Width           =   1620
   End
   Begin VB.CheckBox chkAppAll 
      BackColor       =   &H80000005&
      Caption         =   "应用至所有分类"
      Height          =   180
      Index           =   1
      Left            =   3600
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   2540
      Width           =   1620
   End
   Begin VB.CheckBox chkAppAll 
      BackColor       =   &H80000005&
      Caption         =   "应用至所有分类"
      Height          =   180
      Index           =   0
      Left            =   3600
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1860
      Width           =   1620
   End
   Begin VB.TextBox txtDescription 
      Height          =   900
      Left            =   1080
      MaxLength       =   250
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   3915
      Width           =   4095
   End
   Begin VB.ComboBox cboLogLevel 
      Height          =   300
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3435
      Width           =   2175
   End
   Begin VB.ComboBox cboLogMode 
      Height          =   300
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2970
      Width           =   2175
   End
   Begin VB.TextBox txtKeepDays 
      Height          =   320
      Left            =   1080
      MaxLength       =   5
      TabIndex        =   3
      Top             =   2510
      Width           =   1095
   End
   Begin VB.TextBox txtName 
      Height          =   320
      Left            =   1080
      MaxLength       =   50
      TabIndex        =   0
      Top             =   1115
      Width           =   4095
   End
   Begin VB.Frame fraEnd 
      Height          =   45
      Index           =   1
      Left            =   -90
      TabIndex        =   13
      Top             =   900
      Width           =   6195
   End
   Begin VB.Frame fraEnd 
      Height          =   45
      Index           =   0
      Left            =   0
      TabIndex        =   9
      Top             =   4920
      Width           =   6075
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4080
      TabIndex        =   8
      Top             =   5085
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "保存(&S)"
      Height          =   350
      Left            =   2865
      TabIndex        =   7
      Top             =   5085
      Width           =   1100
   End
   Begin MSComCtl2.DTPicker dtpBeginTime 
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Top             =   1560
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   121372675
      CurrentDate     =   43077.4366782407
   End
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   5520
      TabIndex        =   10
      Top             =   0
      Width           =   5520
      Begin VB.Label lblEXP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1.日志级别是包含关系，高级别包含低级别日志。"
         Height          =   180
         Index           =   1
         Left            =   405
         TabIndex        =   12
         Top             =   135
         Width           =   4500
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblEXP 
         BackStyle       =   0  'Transparent
         Caption         =   "2.日志保留天数作用于记录方式含数据库的日志，所有开启日志的最大保留天数作用于本地日志。"
         Height          =   450
         Index           =   0
         Left            =   405
         TabIndex        =   11
         Top             =   345
         Width           =   4710
      End
   End
   Begin MSComCtl2.DTPicker dtpEndTime 
      Height          =   315
      Left            =   1080
      TabIndex        =   2
      Top             =   2025
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   121372675
      CurrentDate     =   43077.4366782407
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "说明"
      Height          =   180
      Left            =   480
      TabIndex        =   20
      Top             =   3960
      Width           =   360
   End
   Begin VB.Label lblLogLevel 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "日志级别"
      Height          =   180
      Left            =   240
      TabIndex        =   19
      Top             =   3495
      Width           =   720
   End
   Begin VB.Label lblLogMode 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "记录方式"
      Height          =   180
      Left            =   240
      TabIndex        =   18
      Top             =   3030
      Width           =   720
   End
   Begin VB.Label lblKeepDays 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "保留天数"
      Height          =   180
      Left            =   240
      TabIndex        =   17
      Top             =   2565
      Width           =   720
   End
   Begin VB.Label lblEndTime 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "结束记录"
      Height          =   180
      Left            =   240
      TabIndex        =   16
      Top             =   2100
      Width           =   720
   End
   Begin VB.Label lblBeginTime 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "开始记录"
      Height          =   180
      Left            =   240
      TabIndex        =   15
      Top             =   1635
      Width           =   720
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "名称"
      Height          =   180
      Left            =   600
      TabIndex        =   14
      Top             =   1170
      Width           =   360
   End
End
Attribute VB_Name = "frmCommenLogSetEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DATETIME_VB6 As String = "yyyy-mm-dd Hh:Nn:Ss"
Private Const DATETIME_ORA As String = "YYYY-MM-DD HH24:MI:SS"

Private mblnOK          As Boolean
Private mlngId          As Long
Private vsfCategory     As VSFlexGrid
Private mcnOracle       As ADODB.Connection

Public Function ShowMe(ByRef cnOracle As ADODB.Connection, Optional ByVal objCategory As Object, _
    Optional ByVal lngID As Long = 0) As Boolean
    mblnOK = False
    mlngId = lngID
    Set vsfCategory = objCategory
    Set mcnOracle = cnOracle
    Me.Show vbModal, frmMDIMain
    ShowMe = mblnOK
End Function

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strSQL As String
    Dim arrSQL() As Variant
    Dim i As Integer
    Dim blnDo As Boolean
    
    On Error GoTo errH
    
    If txtName.Text = "" Then
        MsgBox "请输入日志分类名称！", vbInformation, gstrSysName
        txtName.SetFocus
        Exit Sub
    End If
    If dtpEndTime.value < dtpBeginTime.value Then
        MsgBox "结束记录时间必须大于开始记录时间！", vbInformation, gstrSysName
        dtpEndTime.SetFocus
        Exit Sub
    End If
    If val(txtKeepDays.Text) <= 0 Then
        MsgBox "请输入保留天数！", vbInformation, gstrSysName
        txtKeepDays.SetFocus
        Exit Sub
    End If
    
    '提醒
    If cboLogMode.Text <> "不记录" Then
        If Not ExistsLogDetail(mlngId) Then
            MsgBox "未发现该日志分类有日志规则的设置，因此所有产品终端都将生效，这样可能会带来性能影响，建议立即设置日志规则！" _
                , vbExclamation, gstrSysName
        End If
    End If
    
    strSQL = "Zllogcategory_Edit(" & IIf(mlngId <= 0, 0, 1) & _
        ", " & mlngId & _
        ", '" & txtName.Text & "'" & _
        ", " & SQLAdjust(txtDescription.Text) & _
        ", to_date('" & Format(dtpBeginTime.value, DATETIME_VB6) & "', '" & DATETIME_ORA & "')" & _
        ", to_date('" & Format(dtpEndTime.value, DATETIME_VB6) & "', '" & DATETIME_ORA & "')" & _
        ", "
    strSQL = strSQL & val(txtKeepDays.Text) & ", " & cboLogMode.ListIndex & ", " & cboLogLevel.ListIndex & ")"
    arrSQL = Array()
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = strSQL

    blnDo = False
    For i = chkAppAll.LBound To chkAppAll.UBound
        blnDo = blnDo Or chkAppAll(i).Visible And chkAppAll(i).value = 1
    Next
    If blnDo Then
        '需要应用至所有分类
        With vsfCategory
            For i = .FixedRows To .Rows - 1
                If val(.TextMatrix(i, .ColIndex("ID"))) <> mlngId Then
                    strSQL = "Zllogcategory_Edit(1" & _
                        ", " & val(.TextMatrix(i, .ColIndex("ID"))) & _
                        ", '" & .TextMatrix(i, .ColIndex("名称")) & "'" & _
                        ", " & SQLAdjust(.TextMatrix(i, .ColIndex("说明")))
                    '时间
                    If chkAppAll(val("0-时间")).value = 1 Then
                        strSQL = strSQL & _
                            ", to_date('" & Format(dtpBeginTime.value, DATETIME_VB6) & "', '" & DATETIME_ORA & "')" & _
                            ", to_date('" & Format(dtpEndTime.value, DATETIME_VB6) & "', '" & DATETIME_ORA & "')"
                    Else
                        strSQL = strSQL & _
                            ", to_date('" & .TextMatrix(i, .ColIndex("启用时间")) & "', '" & DATETIME_ORA & "')" & _
                            ", to_date('" & .TextMatrix(i, .ColIndex("停止时间")) & "', '" & DATETIME_ORA & "')"
                    End If
                    '保留天数
                    If chkAppAll(val("1-保留天数")).value = 1 Then
                        strSQL = strSQL & ", " & val(txtKeepDays.Text)
                    Else
                        strSQL = strSQL & ", " & val(.TextMatrix(i, .ColIndex("保留天数")))
                    End If
                    '记录方式
                    If chkAppAll(val("2-记录方式")).value = 1 Then
                        strSQL = strSQL & ", " & cboLogMode.ListIndex
                    Else
                        strSQL = strSQL & ", " & val(.TextMatrix(i, .ColIndex("日志方式")))
                    End If
                    '日志级别
                    If chkAppAll(val("3-日志级别")).value = 1 _
                        And InStr(";公共日志;服务公共日志;", ";" & Trim(.TextMatrix(i, .ColIndex("名称"))) & ";") <= 0 Then
                        strSQL = strSQL & ", " & cboLogLevel.ListIndex
                    Else
                        strSQL = strSQL & ", " & val(.TextMatrix(i, .ColIndex("日志级别")))
                    End If
                    strSQL = strSQL & ")"
                    
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = strSQL
                End If
            Next
        End With
    End If
    
    Call gclsBase.ExecuteProcedureBeach(mcnOracle, arrSQL, Me.Caption)
    mblnOK = True
    Unload Me
    Exit Sub
    
errH:
    MsgBox "保存日志分类出错：" & err.Description, vbInformation, gstrSysName
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim dtCurrent As Date
    Dim i As Integer
    
    On Error GoTo errH
    
    If mlngId = 0 Then
        Me.Caption = Me.Caption & "(新增)"
    Else
        Me.Caption = Me.Caption & "(修改)"
    End If
    
    cboLogMode.addItem "不记录"
    cboLogMode.addItem "本地记录"
    cboLogMode.addItem "数据库记录"
    cboLogMode.addItem "本地和数据库记录"
    cboLogMode.ListIndex = 0
    
    cboLogLevel.addItem "0-关闭"
    cboLogLevel.addItem "1-错误"
    cboLogLevel.addItem "2-警告"
    cboLogLevel.addItem "3-重要"
    cboLogLevel.addItem "4-跟踪"
    cboLogLevel.addItem "5-全开"
    cboLogLevel.ListIndex = 0
    
    If mlngId > 0 Then
        '修改
        strSQL = "Select ID, Name, Description, Builtin, Begin_Time, End_Time, Log_Keep_Days, Log_Mode, Log_Level " & vbNewLine & _
                 "From Zllogcategory " & vbNewLine & _
                 "Where ID = [1]"
        Set rsTmp = gclsBase.OpenSQLRecord(mcnOracle, strSQL, Me.Caption, mlngId)
        With rsTmp
            txtName.Text = !name & ""
            dtpBeginTime.value = Nvl(!Begin_Time, 0)
            dtpEndTime.value = Nvl(!End_Time, 0)
            txtKeepDays.Text = !Log_Keep_Days & ""
            cboLogMode.ListIndex = val(!Log_Mode & "")
            cboLogLevel.ListIndex = val(!Log_Level & "")
            txtDescription.Text = !Description & ""
            
            txtName.Locked = Not Nvl(!BuiltIn, 0) = 0
            If txtName.Locked Then txtName.BackColor = vbMenuBar
            .Close
        End With
        txtDescription.Locked = txtName.Locked
        txtDescription.BackColor = txtName.BackColor
    Else
        '新增
        dtCurrent = CurrentDate(mcnOracle)
        dtpBeginTime.value = dtCurrent
        dtpEndTime.value = CDate("3000/1/1")
    End If
    
    '应用至所有分类
    For i = chkAppAll.LBound To chkAppAll.UBound
        chkAppAll(i).Visible = mlngId > 0
    Next
    
    '固定日志级别的项目不允许修改
    If InStr(";公共日志;服务公共日志;", ";" & Trim(txtName.Text) & ";") > 0 Then
        cboLogLevel.Enabled = False
        chkAppAll(3).Enabled = False
    End If
    
    Exit Sub
    
errH:
    MsgBox err.Description, vbInformation, gstrSysName
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mcnOracle = Nothing
End Sub

Private Sub txtDescription_KeyPress(KeyAscii As Integer)
    If InStr("`~!@#$%^&*()=+|\{}[];':"",./<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtKeepDays_KeyPress(KeyAscii As Integer)
    If Not Chr(KeyAscii) Like "#" Then
        If Chr(KeyAscii) <> vbBack Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If InStr("`~!@#$%^&*()=+|\{}[];':"",./<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Function ExistsLogDetail(ByVal lngCategoryID As Long) As Boolean
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo hErr
    strSQL = "Select Count(1) Rec From ZllogSet Where Category_Id = [1] "
    Set rsTemp = gclsBase.OpenSQLRecord(mcnOracle, strSQL, "获取日志分类的日志规则", lngCategoryID)
    ExistsLogDetail = rsTemp!Rec >= 1
    rsTemp.Close
    Exit Function
    
hErr:
    MsgBox err.Description, vbInformation, gstrSysName
End Function
