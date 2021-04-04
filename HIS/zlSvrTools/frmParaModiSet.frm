VERSION 5.00
Begin VB.Form frmParaModiSet 
   Caption         =   "调整参数值与新增参数配置"
   ClientHeight    =   3090
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4845
   Icon            =   "frmParaModiSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4845
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fraSplit 
      BackColor       =   &H80000012&
      Height          =   30
      Index           =   0
      Left            =   0
      TabIndex        =   11
      Top             =   2400
      Width           =   6700
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   4845
      TabIndex        =   10
      Top             =   2475
      Width           =   4845
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   2400
         TabIndex        =   8
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   3555
         TabIndex        =   9
         Top             =   120
         Width           =   1100
      End
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   0
      ScaleHeight     =   2295
      ScaleWidth      =   4995
      TabIndex        =   12
      Top             =   0
      Width           =   5000
      Begin VB.TextBox txtOld 
         BackColor       =   &H8000000F&
         Height          =   300
         Left            =   1365
         MaxLength       =   4000
         TabIndex        =   16
         Top             =   960
         Visible         =   0   'False
         Width           =   2970
      End
      Begin VB.Frame fraSplit 
         BackColor       =   &H80000012&
         Height          =   30
         Index           =   1
         Left            =   0
         TabIndex        =   14
         Top             =   600
         Width           =   6700
      End
      Begin VB.CommandButton cmdPC 
         Caption         =   "…"
         Height          =   300
         Left            =   4065
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1755
         Width           =   270
      End
      Begin VB.CommandButton cmdUser 
         Caption         =   "…"
         Height          =   300
         Left            =   4065
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1440
         Width           =   270
      End
      Begin VB.TextBox txtPC 
         BackColor       =   &H8000000F&
         Height          =   300
         Left            =   1365
         Locked          =   -1  'True
         MaxLength       =   4000
         TabIndex        =   4
         Top             =   1740
         Width           =   2730
      End
      Begin VB.TextBox txtUser 
         BackColor       =   &H8000000F&
         Height          =   300
         Left            =   1365
         Locked          =   -1  'True
         MaxLength       =   4000
         TabIndex        =   1
         Top             =   1440
         Width           =   2730
      End
      Begin VB.TextBox txtValue 
         Height          =   300
         Left            =   1365
         MaxLength       =   4000
         TabIndex        =   7
         Top             =   2040
         Width           =   2970
      End
      Begin VB.Label lblOld 
         AutoSize        =   -1  'True
         Caption         =   "原参数值 "
         Height          =   180
         Left            =   480
         TabIndex        =   17
         Top             =   1020
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label lblChangeInfo 
         AutoSize        =   -1  'True
         Caption         =   "当前选择了三条参数信息"
         Height          =   180
         Left            =   480
         TabIndex        =   15
         Top             =   720
         Visible         =   0   'False
         Width           =   1980
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         Caption         =   "请确认参数值范围与参数值组成规则，然后使用该功能。"
         Height          =   180
         Left            =   120
         TabIndex        =   13
         Top             =   180
         Width           =   4500
      End
      Begin VB.Label lblUser 
         AutoSize        =   -1  'True
         Caption         =   "用户名(&U)"
         Height          =   180
         Left            =   480
         TabIndex        =   0
         Top             =   1500
         Width           =   810
      End
      Begin VB.Label lblValue 
         AutoSize        =   -1  'True
         Caption         =   "参数值(&V)"
         Height          =   180
         Left            =   480
         TabIndex        =   6
         Top             =   2100
         Width           =   810
      End
      Begin VB.Label lblPC 
         AutoSize        =   -1  'True
         Caption         =   "机器名(&M)"
         Height          =   180
         Left            =   480
         TabIndex        =   3
         Top             =   1800
         Width           =   810
      End
   End
End
Attribute VB_Name = "frmParaModiSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintFunType As Integer '0-修改参数值，1-新增参数设置
'通用相关变量
Private mstrValue As String
Private mblnOk As Boolean
'新增参数设置相关变量
Private mint私有 As Integer '是否是私有参数
Private mint本机 As Integer '是否是本机参数
Private mstrUsers As String '用户名组成的字符串
Private mstrPCs As String '机器名组成的字符串
Private mstrSysOwner As String '当前系统所有者
Private mstrNote As String '参数值修改时的提示
Private mlngParaID As Long '参数ID

Public Function ShowMe(ByVal frmParent As Object, ByVal intFunType As Integer, ByVal strParInfo As String, ByVal strNote As String, ByVal strSysOwner As String, ByVal lngParaID As Long, ByRef strValue As String, Optional ByRef strUsers As String, Optional ByRef strPCs As String) As Boolean
'功能：该窗体的入口
'参数：frmParent=父窗体
'          intFunType=功能：0-修改参数值，1-新增参数设置
'          strParInfo=格式：是否本机,是否私有。本机私有参数可以用1,1标识，本级公共可用1,0等
'          strSysOwner=系统所有者
'          strNote=参数值修改时的提示
'返回=True:确认操作，False-取消操作
'          strValue=新的参数值，仅供返回
'          strUsers=用户名组成的字符串，用逗号分割，仅当新增私有类型参数设置时返回
'          strPCs=机器名组成的字符串，用逗号分割，仅当新增本机类型参数设置时返回
    Dim arrTmp As Variant
    
    arrTmp = Split(strParInfo & ",", ",")
    mint本机 = Val(arrTmp(0))
    mint私有 = Val(arrTmp(1))
    mintFunType = intFunType
    mstrSysOwner = strSysOwner
    mlngParaID = lngParaID
    mstrNote = strNote
    mstrValue = strValue
    If strSysOwner = "" And intFunType = 1 Then
        MsgBox "当前服务器没有安装部门人员数据，不能新增参数设置！", vbInformation, gstrSysName
        Exit Function
    End If
    mstrPCs = ""
    mstrUsers = ""
    mblnOk = False
    
    Me.Show vbModal
    
    ShowMe = mblnOk
    strUsers = mstrUsers
    strPCs = mstrPCs
    strValue = mstrValue
End Function

Private Sub cmdCancel_Click()
    mblnOk = False
    mstrValue = ""
    mstrPCs = ""
    mstrUsers = ""
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strTmp As String
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strParas As String, strMsg As String
    
    If InStr(txtUser.Text, "^") > 0 Then
        MsgBox "用户名含有非法字符""^""，请检查！", vbInformation, gstrSysName
        txtUser.SetFocus
        Exit Sub
    ElseIf InStr(txtUser.Text, "#") > 0 Then
        MsgBox "用户名含有非法字符""#""，请检查！", vbInformation, gstrSysName
        txtUser.SetFocus
        Exit Sub
    ElseIf InStr(txtUser.Text, "'") > 0 Then
        MsgBox "用户名含有非法字符""'""，请检查！", vbInformation, gstrSysName
        txtUser.SetFocus
        Exit Sub
    End If
    
    If InStr(txtPC.Text, "^") > 0 Then
        MsgBox "机器名含有非法字符""^""，请检查！", vbInformation, gstrSysName
        txtPC.SetFocus
        Exit Sub
    ElseIf InStr(txtPC.Text, "#") > 0 Then
        MsgBox "机器名含有非法字符""#""，请检查！", vbInformation, gstrSysName
        txtPC.SetFocus
        Exit Sub
    ElseIf InStr(txtPC.Text, "'") > 0 Then
        MsgBox "机器名含有非法字符""'""，请检查！", vbInformation, gstrSysName
        txtPC.SetFocus
        Exit Sub
    End If
    
    If InStr(txtValue.Text, "^") > 0 Then
        MsgBox "参数值含有非法字符""^""，请检查！", vbInformation, gstrSysName
        txtValue.SetFocus
        Exit Sub
    ElseIf InStr(txtValue.Text, "#") > 0 Then
        MsgBox "参数值含有非法字符""#""，请检查！", vbInformation, gstrSysName
        txtValue.SetFocus
        Exit Sub
    ElseIf InStr(txtValue.Text, "'") > 0 Then
        MsgBox "参数值含有非法字符""'""，请检查！", vbInformation, gstrSysName
        txtValue.SetFocus
        Exit Sub
    End If

    If txtValue.Text = "" Then
        If MsgBox("参数值为空，是否继续？", vbInformation + vbYesNo, gstrSysName) = vbNo Then
            txtValue.SetFocus
            Exit Sub
        End If
    End If
    If txtUser.Visible And txtUser.Text = "" Then
        MsgBox "请输入用户名！", vbInformation, gstrSysName
        txtUser.SetFocus
        Exit Sub
    End If
    If txtPC.Visible And txtPC.Text = "" Then
        MsgBox "请输入机器名！", vbInformation, gstrSysName
        txtPC.SetFocus
        Exit Sub
    End If
    '检测已经存在的参数设置
    strSQL = "Select 参数id, c.用户名, c.机器名" & vbNewLine & _
                "From (Select a.用户名, b.机器名" & vbNewLine & _
                "       From (Select Distinct Column_Value 用户名 From Table(f_Str2list(Nvl([2], ',')))) a," & vbNewLine & _
                "            (Select Distinct Column_Value 机器名 From Table(f_Str2list(Nvl([3], ',')))) b) c, Zluserparas d" & vbNewLine & _
                "Where d.参数id = [1] And Nvl(d.用户名, '空空') = Nvl(c.用户名, '空空') And Nvl(d.机器名, '空空') = Nvl(c.机器名, '空空')"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption, mlngParaID, txtUser.Tag, txtPC.Text)
    If rsTmp.RecordCount <> 0 Then
        If rsTmp.RecordCount = 1 Then
            If mint本机 = 1 And mint私有 = 1 Then
                 strMsg = rsTmp!用户名 & "在" & rsTmp!机器名 & "上的参数设置已经存在"
            Else
                strMsg = IIf(mint私有 = 1, rsTmp!用户名 & "", rsTmp!机器名 & "") & "的参数设置已经存在"
            End If
        Else
            strMsg = "共有" & rsTmp.RecordCount & "条参数设置已经存在"
        End If
        '进行询问是否覆盖，若覆盖，则删除原有参数配置
        If MsgBox(strMsg & "，是否覆盖原有设置？", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Do While Not rsTmp.EOF
                strTmp = rsTmp!用户名 & "^" & rsTmp!机器名
                If ActualLen(strParas & "#" & strTmp) >= 2000 Then
                    Call ExecuteProcedure("Zlparameters_Del_Details(" & mlngParaID & ",'" & strParas & "')", "删除参数设置")
                    strParas = strTmp
                Else
                    strParas = IIf(strParas = "", strTmp, strParas & "#" & strTmp)
                End If
                rsTmp.MoveNext
            Loop
            If strParas <> "" Then
                Call ExecuteProcedure("Zlparameters_Del_Details(" & mlngParaID & ",'" & strParas & "')", "删除参数设置")
            End If
        End If
    End If
    
    mblnOk = True
    mstrValue = txtValue.Text
    mstrPCs = txtPC.Text
    mstrUsers = txtUser.Tag
    Unload Me
End Sub

Private Sub cmdPC_Click()
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean
    Dim strTmp As String
    Dim i As Long

    strSQL = "Select a.Id, a.上级id, a.编码, a.名称, 0 末级" & vbNewLine & _
                    "From " & mstrSysOwner & ".部门表 a" & vbNewLine & _
                    "Where a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000/1/1', 'yyyy-mm-dd')" & vbNewLine & _
                    "Start With 上级id Is Null" & vbNewLine & _
                    "Connect By Prior Id = 上级id" & vbNewLine & _
                    "Union All" & vbNewLine & _
                    "Select RowNum ID, a.Id, a.编码, b.工作站 名称, 1 末级" & vbNewLine & _
                    "From " & mstrSysOwner & ".部门表 a, Zlclients b" & vbNewLine & _
                    "Where a.名称 = b.部门 And (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000/1/1', 'yyyy-mm-dd'))"

    Set rsTmp = gclsBase.ShowSQLSelectEx(gcnOracle, Me, txtPC, strSQL, 2, "工作站选择", False, "", "", True, True, False, blnCancel, True, True, True, "NotShowNon=1")
    If Not blnCancel And Not rsTmp Is Nothing Then
        For i = 1 To rsTmp.RecordCount
            strTmp = strTmp & "," & rsTmp!名称
            rsTmp.MoveNext
        Next
        If strTmp <> "" Then strTmp = Mid(strTmp, 2)
        txtPC.Text = strTmp
    ElseIf Not blnCancel Then
        txtPC.Text = ""
    End If
End Sub

Private Sub cmdUser_Click()
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean
    Dim strTmp As String, strTmp1 As String
    Dim i As Long
    
    strSQL = "Select a.Id, a.上级id, a.编码, a.名称 姓名, ' ' 用户名, 0 末级" & vbNewLine & _
                    "From " & mstrSysOwner & ".部门表 a" & vbNewLine & _
                    "Where a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000/1/1', 'yyyy-mm-dd')" & vbNewLine & _
                    "Start With 上级id Is Null" & vbNewLine & _
                    "Connect By Prior Id = 上级id" & vbNewLine & _
                    "Union All" & vbNewLine & _
                    "Select c.Id, a.Id, c.编号, c.姓名, d.用户名, 1 末级" & vbNewLine & _
                    "From " & mstrSysOwner & ".部门表 a, " & mstrSysOwner & ".部门人员 b, " & mstrSysOwner & ".人员表 c, " & mstrSysOwner & ".上机人员表 d" & vbNewLine & _
                    "Where a.Id = b.部门id And b.人员id = c.Id And c.Id = d.人员id And" & vbNewLine & _
                    "      (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000/1/1', 'yyyy-mm-dd')) And b.缺省=1"

    Set rsTmp = gclsBase.ShowSQLSelectEx(gcnOracle, Me, txtUser, strSQL, 2, "人员选择器", False, "", "", False, True, False, blnCancel, True, True, True, "NotShowNon=1")
    If Not blnCancel And Not rsTmp Is Nothing Then
        For i = 1 To rsTmp.RecordCount
            If InStr("," & strTmp & ",", "," & rsTmp!用户名 & ",") = 0 Then
                strTmp = strTmp & "," & rsTmp!用户名
            End If
            If InStr("," & strTmp1 & ",", "," & rsTmp!姓名 & ",") = 0 Then
                strTmp1 = strTmp1 & "," & rsTmp!姓名
            End If
            rsTmp.MoveNext
        Next
        If strTmp <> "" Then strTmp = Mid(strTmp, 2)
        If strTmp1 <> "" Then strTmp1 = Mid(strTmp1, 2)
        txtUser.Text = strTmp1
        txtUser.Tag = strTmp
    ElseIf Not blnCancel Then
        txtUser.Text = ""
        txtUser.Tag = ""
    End If
End Sub

Private Sub Form_Activate()
    If mintFunType = 0 Then
        txtValue.SetFocus
    Else
        If mint私有 = 0 Then
            txtPC.SetFocus
        Else
            txtUser.SetFocus
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("^") Then
        KeyAscii = 0
    ElseIf KeyAscii = Asc("'") Then
        KeyAscii = 0
    ElseIf KeyAscii = Asc("#") Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    lblPC.Visible = mintFunType = 1 And mint本机 = 1
    txtPC.Visible = lblPC.Visible
    cmdPC.Visible = lblPC.Visible
    lblUser.Visible = mintFunType = 1 And mint私有 = 1
    txtUser.Visible = lblUser.Visible
    cmdUser.Visible = lblUser.Visible
    If mintFunType = 0 Then
        lblChangeInfo.Caption = mstrNote
        lblChangeInfo.Visible = True
        lblOld.Visible = True: txtOld.Visible = True
        txtOld.Text = mstrValue
        txtValue.Text = mstrValue
        Call SetCtrlSameDistance(True, 0, 2, fraSplit(1), lblChangeInfo, lblOld, lblValue, fraSplit(0))
        Call SetCtrlPosOnLine(False, 0, lblOld, 60, txtOld)
        Call SetCtrlPosOnLine(False, 0, lblValue, 60, txtValue)
        Me.Caption = "调整参数值"
    Else
        Me.Caption = "新增参数设置"
        If mint私有 = 0 Then
            Call SetCtrlSameDistance(True, 0, 2, fraSplit(1), lblPC, lblValue, fraSplit(0))
            Call SetCtrlPosOnLine(False, 0, lblPC, 60, txtPC, 0, cmdPC)
            Call SetCtrlPosOnLine(False, 0, lblValue, 60, txtValue)
        ElseIf mint本机 = 0 Then
            Call SetCtrlSameDistance(True, 0, 2, fraSplit(1), lblUser, lblValue, fraSplit(0))
            Call SetCtrlPosOnLine(False, 0, lblUser, 60, txtUser, 0, cmdUser)
            Call SetCtrlPosOnLine(False, 0, lblValue, 60, txtValue)
        Else
            Call SetCtrlSameDistance(True, 0, 2, fraSplit(1), lblUser, lblPC, lblValue, fraSplit(0))
            Call SetCtrlPosOnLine(False, 0, lblUser, 60, txtUser, 0, cmdUser)
            Call SetCtrlPosOnLine(False, 0, lblPC, 60, txtPC, 0, cmdPC)
            Call SetCtrlPosOnLine(False, 0, lblValue, 60, txtValue)
        End If
    End If
End Sub

Private Sub Form_Resize()
    Me.Height = 3660
    Me.Width = 5085
End Sub

Private Sub txtPC_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyDelete Then
        txtPC.Text = ""
        txtPC.Tag = ""
     End If
End Sub

Private Sub txtUser_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyDelete Then
        txtUser.Text = ""
        txtUser.Tag = ""
     End If
End Sub

Private Sub txtValue_GotFocus()
    Call SelAll(txtValue)
End Sub
