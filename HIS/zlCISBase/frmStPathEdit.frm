VERSION 5.00
Begin VB.Form frmStPathEdit 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "标准路径修改"
   ClientHeight    =   6105
   ClientLeft      =   8400
   ClientTop       =   4605
   ClientWidth     =   5415
   Icon            =   "frmStPathEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cboType 
      Enabled         =   0   'False
      Height          =   300
      ItemData        =   "frmStPathEdit.frx":076A
      Left            =   3600
      List            =   "frmStPathEdit.frx":0774
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   975
      Width           =   1695
   End
   Begin VB.TextBox txtPathName 
      Height          =   300
      Left            =   1200
      MaxLength       =   80
      TabIndex        =   1
      Top             =   120
      Width           =   4095
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "选择手术(&M)"
      Height          =   350
      Index           =   1
      Left            =   3945
      TabIndex        =   11
      Top             =   3240
      Width           =   1350
   End
   Begin VB.CommandButton cmdSel 
      Caption         =   "选择疾病(&D)"
      Height          =   350
      Index           =   0
      Left            =   3945
      TabIndex        =   8
      Top             =   1320
      Width           =   1350
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      BackColor       =   &H00F0F4E4&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   5415
      TabIndex        =   17
      Top             =   5490
      Width           =   5415
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   2760
         TabIndex        =   14
         Top             =   160
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   4200
         TabIndex        =   15
         Top             =   160
         Width           =   1100
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   0
         X2              =   10000
         Y1              =   45
         Y2              =   45
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   0
         X2              =   10000
         Y1              =   30
         Y2              =   30
      End
   End
   Begin VB.TextBox txtSuitCode 
      Height          =   1335
      Index           =   1
      Left            =   120
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   3720
      Width           =   5175
   End
   Begin VB.TextBox txtSuitCode 
      Height          =   1335
      Index           =   0
      Left            =   120
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   1800
      Width           =   5175
   End
   Begin VB.ComboBox cboVersion 
      Height          =   300
      Left            =   3600
      TabIndex        =   5
      Top             =   520
      Width           =   1695
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1200
      MaxLength       =   8
      TabIndex        =   7
      Top             =   980
      Width           =   1695
   End
   Begin VB.ComboBox cboDept 
      Height          =   300
      ItemData        =   "frmStPathEdit.frx":0788
      Left            =   1200
      List            =   "frmStPathEdit.frx":078A
      TabIndex        =   3
      Top             =   520
      Width           =   1695
   End
   Begin VB.Label lblType 
      AutoSize        =   -1  'True
      Caption         =   "类型(&T)"
      Height          =   180
      Left            =   3000
      TabIndex        =   19
      Top             =   1035
      Width           =   630
   End
   Begin VB.Label lblPathName 
      AutoSize        =   -1  'True
      Caption         =   "路径名称(&N)"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   990
   End
   Begin VB.Label lblAttention 
      Caption         =   "提示：多个项目以逗号分割"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   5160
      Width           =   5175
   End
   Begin VB.Label lblOprCode 
      Caption         =   "适用手术(ICD-9-CM3手术编码)(&G)"
      Height          =   180
      Left            =   120
      TabIndex        =   12
      Top             =   3360
      Width           =   2895
   End
   Begin VB.Label lblDiseaseCode 
      Caption         =   "适用疾病(ICD-10疾病编码)(&F)"
      Height          =   180
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   3495
   End
   Begin VB.Label lblCode 
      Caption         =   "编    码(&S)"
      Height          =   180
      Left            =   120
      TabIndex        =   6
      Top             =   1040
      Width           =   990
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      Caption         =   "版本(&V)"
      Height          =   180
      Left            =   3000
      TabIndex        =   4
      Top             =   585
      Width           =   630
   End
   Begin VB.Label lblDept 
      AutoSize        =   -1  'True
      Caption         =   "科室名称(&K)"
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   580
      Width           =   990
   End
End
Attribute VB_Name = "frmStPathEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintMode As Integer  '新增或修改，0-新增，1-修改
Private mlngStPathID As Long '要修改的标准路径ID
Private mblnOK As Boolean
Private mstr科室名称 As String
Private mstr版本     As String
Private mstr路径名称 As String
Private mstr编码     As String
Private mstr疾病编码s As String
Private mstr手术编码s As String
Private mbytType As Byte

Private Enum LenMax '字段在数据库中的长度
    LM_编码 = 8
    LM_科室名称 = 100
    LM_路径名称 = 80
    LM_版本说明 = 20
    LM_疾病编码 = 100
    LM_手术编码 = 100
End Enum
Public Function ShowMe(FrmParent As Object, ByVal intMode As Integer, Optional ByRef lngStPathID As Long, Optional ByVal str路径名称 As String, _
            Optional ByVal str编码 As String, Optional ByVal str科室名称 As String, Optional ByVal str版本 As String, _
                Optional ByVal str疾病编码s As String, Optional ByVal str手术编码s As String, Optional ByVal bytType As Byte) As Boolean
'说明：路径维护功能中的更新标准路径，添加标准路径时调用
'   intMode:'0-新增标准路径
'            1-修改更新标准路径与标准路径对应疾病编码
'   lngStPathID,更新标准路径时传入,在新增标准路径时传出新增的标准路径ID
'   str路径名称,str编码,str科室名称,str版本,str疾病编码s,str手术编码s:更新标准路径时传入
'   bytType：0-西医,1-中医
    
    mintMode = intMode
    mlngStPathID = lngStPathID
    mstr路径名称 = str路径名称
    mstr编码 = str编码
    mstr科室名称 = str科室名称
    mstr版本 = str版本
    mstr疾病编码s = str疾病编码s
    mstr手术编码s = str手术编码s
    mbytType = bytType
    
    Me.Show 1, FrmParent
    ShowMe = mblnOK
    lngStPathID = mlngStPathID
End Function


Private Sub cboVersion_Change()
    Call CheckInput(False)
End Sub

Private Sub cboDept_Change()
    Call CheckInput(False)
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub LoadCboData()
'功能：加载下拉列表数据
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    strSql = "Select  科室名称,Rownum ID  From (Select Distinct 科室名称 From 标准路径目录 where NVl(类别,0)=[1] Order By 科室名称)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mbytType)
    Call zlControl.CboAddData(cboDept, rsTmp, True)
    strSql = "Select 版本说明,Rownum ID From (Select Distinct 版本说明 From 标准路径目录 where NVl(类别,0)=[1] Order By 版本说明)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mbytType)
    Call zlControl.CboAddData(cboVersion, rsTmp, True)
    '类型
    Call zlControl.CboLocate(cboType, mbytType, True)
    
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Load编码()
'功能：加载手术编码与疾病编码
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    strSql = "Select 疾病编码, 手术编码 From 标准路径病种 Where 标准路径id =[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngStPathID)
    If rsTmp.RecordCount > 0 Then
        txtSuitCode(0).Text = Nvl(rsTmp!疾病编码)
        txtSuitCode(1).Text = Nvl(rsTmp!手术编码)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdOK_Click()
    If Not SaveData Then Exit Sub
    mblnOK = True
    Unload Me
End Sub

Private Sub cmdSel_Click(Index As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim str编码s As String
    
    On Error GoTo errH
    'D:ICD-10疾病编码 S:ICD-9-CM3手术编码 B:中医疾病编码
    Set rsTmp = zlDatabase.ShowILLSelect(Me, IIf(Index = 0, IIf(mbytType = 0, "D", "B,D"), "S"), 0, , True, , IIf(Trim(txtSuitCode(Index).Text) = "", "", "," & txtSuitCode(Index).Text & ","))
    
    If rsTmp Is Nothing Then Exit Sub
    
    If rsTmp.RecordCount <> 0 Then
            str编码s = ""
            rsTmp.MoveFirst
            Do While Not rsTmp.EOF
                str编码s = str编码s & "," & rsTmp!编码
                rsTmp.MoveNext
            Loop
            str编码s = Mid(str编码s, 2)
            txtSuitCode(Index).Text = str编码s
    End If
    Call CheckInput(False)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
'功能：初始化界面数据
    Call LoadCboData
    
    If mintMode = 0 Then
        Me.Caption = "新增标准路径"
        cboDept.Text = mstr科室名称
    Else
        Call Load编码
        Me.Caption = "修改标准路径"
        cboVersion.Text = mstr版本
        cboDept.Text = mstr科室名称
        txtPathName.Text = mstr路径名称
        txtCode.Text = mstr编码
        txtSuitCode(0).Text = mstr疾病编码s
        txtSuitCode(1).Text = mstr疾病编码s
    End If
    
    If mbytType = 0 Then
        lblDiseaseCode.Caption = "适用疾病(ICD-10疾病编码)(&F)"
    Else
        lblDiseaseCode.Caption = "适用疾病(TCD编码或ICD-10疾病编码)(&F)"
    End If

End Sub



Private Function SaveData() As Boolean
'功能：数据合理性检查并保存
    Dim rsTmp As ADODB.Recordset
    Dim str疾病编码s As String, str手术编码s As String
    Dim strSql As String
    
    If CheckInput(True) = False Then
        SaveData = False: Exit Function
    End If
    
    str疾病编码s = Replace(Trim(txtSuitCode(0).Text), "，", ",")
    str手术编码s = Replace(Trim(txtSuitCode(1).Text), "，", ",")
    
    On Error GoTo errH
    '新增
    If mintMode = 0 Then
        '获取新增路径的ID
        strSql = "Select 标准路径目录_Id.Nextval ID From Dual"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
        mlngStPathID = rsTmp!ID
        
        strSql = "Zl_标准路径目录_Insert(" & mlngStPathID & ",'" & Trim(cboDept.Text) & "','" & Trim(txtCode.Text) & "','" & Trim(txtPathName.Text) & "','" & Trim(cboVersion.Text) & "','" & _
                str疾病编码s & "','" & str手术编码s & "'," & mbytType & ")"
        Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    Else '修改
        strSql = "Zl_标准路径目录_Update(" & mlngStPathID & ",'" & Trim(cboDept.Text) & "','" & Trim(txtCode.Text) & "','" & Trim(txtPathName.Text) & "','" & Trim(cboVersion.Text) & "','" & _
                str疾病编码s & "','" & str手术编码s & "')"
        Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    End If
    SaveData = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckInput(ByVal blnCheckNull As Boolean) As Boolean
'功能：进行输入合法性检查
'参数：blnCheckNull，是否进行空值检验

    Dim strMsg As String
    '空值检查
    If blnCheckNull Then
        If Trim(txtPathName.Text) = "" Then
            MsgBox "你尚未输入标准路径名称", vbInformation, gstrSysName
            txtPathName.SetFocus: Exit Function
        End If
        If Trim(txtCode.Text) = "" Then
            MsgBox "你尚未输入标准路径编码", vbInformation, gstrSysName
            txtCode.SetFocus: Exit Function
        End If
        If Trim(cboVersion.Text) = "" Then
            MsgBox "你尚未输入标准路径名称", vbInformation, gstrSysName
            cboVersion.SetFocus: Exit Function
        End If
        If Trim(cboDept.Text) = "" Then
            MsgBox "你尚未输入标准路径编码", vbInformation, gstrSysName
            cboDept.SetFocus: Exit Function
        End If
    End If
    '长度检查
    If LenB(StrConv(txtPathName.Text, vbFromUnicode)) > LM_路径名称 Then
        strMsg = "你输入的路径名称超过了允许的最大长度" & LM_路径名称 & "(" & LM_路径名称 \ 2 & "个中文的长度或" & LM_路径名称 & "个字母或数字的长度)"
        MsgBox strMsg, vbInformation, gstrSysName
        txtPathName.SetFocus
        CheckInput = False
        Exit Function
    End If
    
    If LenB(StrConv(cboDept.Text, vbFromUnicode)) > LM_科室名称 Then
        strMsg = "你输入的科室名称超过了允许的最大长度" & LM_科室名称 & "(" & LM_科室名称 \ 2 & "个中文的长度或" & LM_科室名称 & "个字母或数字的长度)"
        MsgBox strMsg, vbInformation, gstrSysName
        cboDept.SetFocus
        CheckInput = False
        Exit Function
    End If
    
    If LenB(StrConv(cboVersion.Text, vbFromUnicode)) > LM_版本说明 Then
        strMsg = "你输入的版本超过了允许的最大长度" & LM_版本说明 & "(" & LM_版本说明 \ 2 & "个中文的长度或" & LM_版本说明 & "个字母或数字的长度)"
        MsgBox strMsg, vbInformation, gstrSysName
        cboVersion.SetFocus
        CheckInput = False
        Exit Function
    End If
    
    If LenB(StrConv(txtCode.Text, vbFromUnicode)) > LM_编码 Then
        strMsg = "你输入的编码超过了允许的最大长度" & LM_编码 & "(" & LM_编码 \ 2 & "个中文的长度或" & LM_编码 & "个字母或数字的长度)"
        MsgBox strMsg, vbInformation, gstrSysName
        txtCode.SetFocus
        CheckInput = False
        Exit Function
    End If
    
    If LenB(StrConv(txtSuitCode(0).Text, vbFromUnicode)) > LM_疾病编码 Then
        strMsg = "你输入的疾病编码超过了允许的最大长度" & LM_疾病编码 & "(" & LM_疾病编码 \ 2 & "个中文的长度或" & LM_疾病编码 & "个字母或数字的长度)"
        MsgBox strMsg, vbInformation, gstrSysName
        txtSuitCode(0).SetFocus
        CheckInput = False
        Exit Function
    End If
    
    If LenB(StrConv(txtSuitCode(1).Text, vbFromUnicode)) > LM_手术编码 Then
        strMsg = "你输入的手术编码超过了允许的最大长度" & LM_手术编码 & "(" & LM_手术编码 \ 2 & "个中文的长度或" & LM_手术编码 & "个字母或数字的长度)"
        MsgBox strMsg, vbInformation, gstrSysName
        txtSuitCode(1).SetFocus
        CheckInput = False
        Exit Function
    End If
    
    CheckInput = True
    
End Function

Private Sub Form_Resize()
    
    If Me.WindowState = vbMaximized Or Me.WindowState = vbMinimized Then Exit Sub
    '不允许改变窗体大小
    If Me.Width < 5500 Then Me.Width = 5500
    If Me.Height < 6500 Then Me.Height = 6500
    
End Sub

Private Sub txtCode_Change()
    Call CheckInput(False)
End Sub

Private Sub txtcode_GotFocus()
'功能：获得焦点是选中文本
    txtCode.SelStart = 0
    txtCode.SelLength = Len(txtCode.Text)
End Sub

Private Sub txtPathName_Change()
    Call CheckInput(False)
End Sub

Private Sub txtPathName_GotFocus()
'功能：获得焦点是选中文本
    txtPathName.SelStart = 0
    txtPathName.SelLength = Len(txtPathName.Text)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'功能：回车定位下一个控件
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub
