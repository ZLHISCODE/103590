VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmPatientGroupEdit 
   Caption         =   "人员划分"
   ClientHeight    =   5580
   ClientLeft      =   2775
   ClientTop       =   4050
   ClientWidth     =   9645
   Icon            =   "frmPatientGroupEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   9645
   Begin zl9Medical.VsfGrid vsf 
      Height          =   2475
      Left            =   2910
      TabIndex        =   12
      Top             =   1785
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4366
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   16
      Top             =   5220
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPatientGroupEdit.frx":076A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11959
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
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
   Begin MSComctlLib.ImageList ilsMenu 
      Left            =   8115
      Top             =   780
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientGroupEdit.frx":0FFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientGroupEdit.frx":1218
            Key             =   "Start"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientGroupEdit.frx":1438
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientGroupEdit.frx":1652
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientGroupEdit.frx":1872
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsHotMenu 
      Left            =   7515
      Top             =   780
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientGroupEdit.frx":1A92
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientGroupEdit.frx":1CAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientGroupEdit.frx":1EC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientGroupEdit.frx":2218
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientGroupEdit.frx":2438
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   705
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   1244
      BandCount       =   1
      _CBWidth        =   9645
      _CBHeight       =   705
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinHeight1      =   645
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   645
         Left            =   30
         TabIndex        =   18
         Top             =   30
         Width           =   9525
         _ExtentX        =   16801
         _ExtentY        =   1138
         ButtonWidth     =   1296
         ButtonHeight    =   1138
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ilsMenu"
         HotImageList    =   "ilsHotMenu"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&S.保存"
               Key             =   "保存"
               Object.ToolTipText     =   "保存(Alt+S)"
               Object.Tag             =   "&S.保存"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&R.恢复"
               Key             =   "恢复"
               Object.ToolTipText     =   "恢复(Alt+R)"
               Object.Tag             =   "&R.恢复"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&H.帮助"
               Key             =   "帮助"
               Object.ToolTipText     =   "帮助(Alt+H)"
               Object.Tag             =   "&H.帮助"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&X.退出"
               Key             =   "退出"
               Object.ToolTipText     =   "退出(Alt+X)"
               Object.Tag             =   "&X.退出"
               ImageIndex      =   5
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   10155
      Top             =   4665
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientGroupEdit.frx":2658
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fra 
      Height          =   4425
      Left            =   30
      TabIndex        =   0
      Top             =   720
      Width           =   2835
      Begin VB.CommandButton cmdClear 
         Caption         =   "不选择(&C)"
         Height          =   350
         Left            =   120
         TabIndex        =   21
         Top             =   3105
         Width           =   1470
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   2310
         Width           =   2580
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "选择(&S)"
         Height          =   350
         Left            =   120
         TabIndex        =   9
         Top             =   2685
         Width           =   1470
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   0
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   435
         Width           =   2580
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   1
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1665
         Width           =   2580
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   1050
         Width           =   2580
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&4.姓名"
         Height          =   180
         Index           =   5
         Left            =   105
         TabIndex        =   7
         Top             =   2055
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&2.年龄"
         Height          =   180
         Index           =   0
         Left            =   105
         TabIndex        =   3
         Top             =   795
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&1.性别"
         Height          =   180
         Index           =   4
         Left            =   90
         TabIndex        =   1
         Top             =   195
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&3.婚姻状况"
         Height          =   180
         Index           =   1
         Left            =   105
         TabIndex        =   5
         Top             =   1425
         Width           =   900
      End
   End
   Begin VB.Frame fra1 
      Height          =   585
      Left            =   3225
      TabIndex        =   19
      Top             =   4695
      Width           =   6585
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   2
         Left            =   1275
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   180
         Width           =   2580
      End
      Begin VB.CommandButton cmdAdjust 
         Caption         =   "调整(&J)"
         Height          =   350
         Left            =   3975
         TabIndex        =   15
         Top             =   150
         Width           =   1470
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&6.调整为组别"
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   225
         Width           =   1080
      End
   End
   Begin VB.Frame fra2 
      Height          =   540
      Left            =   2850
      TabIndex        =   20
      Top             =   1185
      Width           =   6525
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   3
         Left            =   705
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   165
         Width           =   2580
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&5.组别"
         Height          =   180
         Index           =   3
         Left            =   75
         TabIndex        =   10
         Top             =   225
         Width           =   540
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFileSave 
         Caption         =   "保存(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileRestore 
         Caption         =   "恢复(&R)"
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "查看(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "工具栏(&T)"
         Begin VB.Menu mnuViewToolButton 
            Caption         =   "标准按钮(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolText 
            Caption         =   "文本标签(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "状态栏(&S)"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpTopic 
         Caption         =   "帮助主题(&T)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "&Web上的中联"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "中联主页(&H)"
         End
         Begin VB.Menu mnuHelpWebForum
            Caption         =   "中联论坛(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "发送反馈(&K)..."
         End
      End
      Begin VB.Menu mnuHelp_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
   End
End
Attribute VB_Name = "frmPatientGroupEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'（１）窗体级变量定义**************************************************************************************************
Private mblnStartUp As Boolean                          '窗体启动标志
Private mblnOK As Boolean
Private mfrmMain As Object
Private mlngKey As Long
Private mblnChanged As Boolean
Private mstrKey As String
Private mrsMember As New ADODB.Recordset
Private mlngLoop As Long

'（２）自定义过程或函数************************************************************************************************
Private Property Let EditChanged(ByVal vData As Boolean)
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '值域:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngSvrKey As Long
    
    mnuFileSave.Enabled = True
    mnuFileRestore.Enabled = True
        
    If vData = False Then
        mnuFileSave.Enabled = False
        mnuFileRestore.Enabled = False
    End If
    
    tbrThis.Buttons("保存").Enabled = mnuFileSave.Enabled
    tbrThis.Buttons("恢复").Enabled = mnuFileRestore.Enabled
    
End Property

Private Function ClearData(Optional ByVal strMenuItem As String = "") As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '参数:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    
    Call ResetVsf(vsf)
    vsf.AppendRow = True
    
    Call InitData
    
    EditChanged = True
        
End Function

Public Function ShowEdit(ByVal frmMain As Object, ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  显示编辑窗体，是与调用窗体的接口函数
    '参数:  frmMain         调用窗体对象
    '       lngKey          预约登记id
    '返回:  True
    '       False
    '------------------------------------------------------------------------------------------------------------------
    mblnStartUp = True
    mblnOK = False
    
    Call ClearData
                    
    mlngKey = lngKey
    Set mfrmMain = frmMain
    
    If InitData = False Then Exit Function
    If ReadData() = False Then Exit Function
    
    stbThis.Panels(2).Text = "将团体体检人员划分为不同的组，进行不同的体检。"
    
    EditChanged = False
    
    mblnStartUp = False
    
    Call cbo_Click(3)
    
    Me.Show 1, frmMain
    
    ShowEdit = mblnOK
    
End Function

Private Function ReadData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  读取数据
    '参数:  lngKey      体检类型序号
    '返回:  True        读取成功
    '       False       读取失败
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
        
    On Error GoTo errHand
    
               
'    将所有人员保存在窗体变量中
    gstrSQL = "select A.病人id AS ID,0 AS 选择,A.姓名,A.门诊号,A.性别,A.年龄,TO_CHAR(A.出生日期,'yyyy-mm-dd') AS 出生日期,A.婚姻状况,NVL(C.组别名称,'') AS 体检组别,0 As 前景色 " & _
                "from 病人信息 A,体检人员档案 B,(SELECT * FROM 体检组别 WHERE 登记id=" & mlngKey & ") C  " & _
                "WHERE A.病人ID=B.病人ID AND B.组别名称=C.组别名称(+) AND B.登记id=" & mlngKey & " Order By C.组别名称,A.门诊号"
    
    Set mrsMember = New ADODB.Recordset
    mrsMember.Open gstrSQL, gcnOracle, adOpenStatic, adLockBatchOptimistic
        
    ReadData = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
    
End Function

Private Function InitData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  初始化设置
    '返回:  True        初始化成功
    '       False       初始化失败
    '------------------------------------------------------------------------------------------------------------------
    Dim strVsf As String
    Dim rs As New ADODB.Recordset
    
    mstrKey = ""
    
    On Error GoTo errHand
    
    With vsf
        .Cols = 0
        .NewColumn "", 255, 4
        .NewColumn "选择", 450, 1, , 1, , flexDTBoolean
        .NewColumn "姓名", 1200, 1
        .NewColumn "门诊号", 900, 7
        .NewColumn "性别", 600, 1
        .NewColumn "年龄", 900, 1
        .NewColumn "婚姻状况", 900, 1
        .NewColumn "出生日期", 1080, 1
        .NewColumn "体检组别", 1500, 1, GetCombList("SELECT 组别名称 FROM 体检组别 Where 登记id=" & mlngKey), 1
        .NewColumn "", 15, 1
        .ExtendLastCol = True
        .FixedCols = 1
        .Body.GridColor = &HC1C1C1
        .AppendRow = True
    End With
    
    cbo(2).Clear
    cbo(3).Clear
    
    cbo(3).AddItem "<所有>"
    '读取组别信息
    gstrSQL = "SELECT 组别名称 AS 名称,ROWNUM AS ID FROM 体检组别 WHERE 登记id=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey)
    If rs.BOF = False Then
        Call AddComboData(cbo(2), rs)
        Call AddComboData(cbo(3), rs, False)
    End If
    If cbo(2).ListCount > 0 Then cbo(2).ListIndex = 0
    If cbo(3).ListCount > 0 Then cbo(3).ListIndex = 0
    
    cbo(0).Clear
    cbo(0).AddItem "<所有>"
    gstrSQL = "SELECT 编码||'-'||名称 AS 名称,0 AS ID FROM 性别 ORDER BY 编码"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If rs.BOF = False Then Call AddComboData(cbo(0), rs, False)
    If cbo(0).ListCount > 0 Then cbo(0).ListIndex = 0
    
    cbo(1).Clear
    cbo(1).AddItem "<所有>"
    gstrSQL = "SELECT 编码||'-'||名称 AS 名称,0 AS ID FROM 婚姻状况 ORDER BY 编码"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If rs.BOF = False Then Call AddComboData(cbo(1), rs, False)
    If cbo(1).ListCount > 0 Then cbo(1).ListIndex = 0
    
    InitData = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function ValidEdit() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  校验数据的有效性
    '返回:  True        数据有效
    '       False       数据无效
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
        
    ValidEdit = True
    
End Function

Private Function SaveEdit() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  保存数据
    '返回:  True        保存成功
    '       False       保存失败
    '------------------------------------------------------------------------------------------------------------------
    Dim blnTran As Boolean
    Dim strSQL As String
    Dim lngLoop As Long
    Dim rsPati As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    
    On Error GoTo errHand
    
    Call SQLRecord(rsSQL)
    
    mrsMember.Filter = ""
    If mrsMember.RecordCount > 0 Then
        mrsMember.MoveFirst
        Do While Not mrsMember.EOF
            
            strSQL = "ZL_体检人员档案_CLASS(" & mlngKey & "," & mrsMember("ID").Value & ",'" & mrsMember("体检组别").Value & "')"
            Call SQLRecordAdd(rsSQL, strSQL)
            
            mrsMember.MoveNext
        Loop
    End If
    
    blnTran = True
    gcnOracle.BeginTrans
    
    If rsSQL.RecordCount > 0 Then rsSQL.MoveFirst
    For lngLoop = 1 To rsSQL.RecordCount
        Call zlDatabase.ExecuteProcedure(CStr(rsSQL("SQL").Value), Me.Caption)
        rsSQL.MoveNext
    Next
    
    gcnOracle.CommitTrans
    blnTran = False

    SaveEdit = True

    Exit Function

errHand:

    If ErrCenter = 1 Then Resume
    If blnTran Then gcnOracle.RollbackTrans

End Function


Private Sub cbo_Click(Index As Integer)
    
    If mblnStartUp = True Then Exit Sub
    
    Select Case Index
    Case 2
        '
        
    Case 3
        
        Call ResetVsf(vsf)
        
        mrsMember.Filter = ""
        If cbo(Index).Text <> "<所有>" Then
            mrsMember.Filter = "体检组别='" & cbo(Index).Text & "'"
        End If
        
        mrsMember.Sort = "体检组别,门诊号"
           
        If mrsMember.RecordCount > 0 Then
            mrsMember.MoveFirst
            Call FillGrid(vsf, mrsMember, Array("", "", "", "", "", "", "yyyy-MM-dd"))
        End If
        vsf.AppendRow = True
        
    End Select
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdAdjust_Click()
    Dim lngLoop As Long
    Dim blnFlag As Boolean
    
    For lngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(lngLoop)) > 0 Then
            If Abs(Val(vsf.TextMatrix(lngLoop, 1))) = 1 Then
                vsf.TextMatrix(lngLoop, 8) = cbo(2).Text
                Call vsf_AfterEdit(lngLoop, 8)
                blnFlag = True
            End If
        End If
    Next
    
    If blnFlag Then Call cbo_Click(3)
    
End Sub

Private Sub cmdClear_Click()
    Dim lngLoop As Long
    Dim strFilter As String
    Dim varTmp2 As Variant
    Dim strTmp As String
    Dim strStart As String
    Dim strEnd As String
    
    strFilter = ""
    
    If cbo(3).Text <> "<所有>" Then strFilter = " AND 体检组别='" & cbo(3).Text & "'"
    If cbo(0).Text <> "<所有>" Then strFilter = strFilter & " AND 性别='" & zlCommFun.GetNeedName(cbo(0).Text) & "'"
    If cbo(1).Text <> "<所有>" Then strFilter = strFilter & " AND 婚姻状况='" & zlCommFun.GetNeedName(cbo(1).Text) & "'"
    
    If Trim(txt(1).Text) <> "" Then strFilter = " AND 姓名='" & txt(1).Text & "'"
    
    varTmp2 = Split(Trim(txt(0).Text), ",")
    strTmp = ""
    For lngLoop = 0 To UBound(varTmp2)
        If InStr(varTmp2(lngLoop), "-") = 0 Then
            strTmp = strTmp & "  OR 年龄='" & varTmp2(lngLoop) & "'"
        Else
            strTmp = strTmp & "  OR (年龄>='" & Mid(varTmp2(lngLoop), 1, InStr(varTmp2(lngLoop), "-") - 1) & "' AND 年龄<='" & Mid(varTmp2(lngLoop), InStr(varTmp2(lngLoop), "-") + 1) & "')"
        End If
    Next
    If strTmp <> "" Then strFilter = strFilter & " AND (" & Mid(strTmp, 5) & ")"
        
    mrsMember.Filter = ""
    If strFilter <> "" Then mrsMember.Filter = Mid(strFilter, 6)
                                
    If mrsMember.RecordCount > 0 Then
        mrsMember.MoveFirst
        Do While Not mrsMember.EOF
            mrsMember.Update "选择", 0
            mrsMember.Update "前景色", 0
            
            mrsMember.MoveNext
        Loop
    End If
    
    Call cbo_Click(3)
    
'    If mrsMember.RecordCount > 0 Then
'        mrsMember.MoveFirst
'        Do While Not mrsMember.EOF
'            mrsMember.Update "选择", 0
'            mrsMember.MoveNext
'        Loop
'    End If
End Sub

Private Sub cmdSelect_Click()
    
    Dim lngLoop As Long
    Dim strFilter As String
    Dim varTmp2 As Variant
    Dim strTmp As String
    Dim strStart As String
    Dim strEnd As String
    
    strFilter = ""
    
    If cbo(3).Text <> "<所有>" Then strFilter = " AND 体检组别='" & cbo(3).Text & "'"
    If cbo(0).Text <> "<所有>" Then strFilter = strFilter & " AND 性别='" & zlCommFun.GetNeedName(cbo(0).Text) & "'"
    If cbo(1).Text <> "<所有>" Then strFilter = strFilter & " AND 婚姻状况='" & zlCommFun.GetNeedName(cbo(1).Text) & "'"
    
    If Trim(txt(1).Text) <> "" Then strFilter = " AND 姓名='" & txt(1).Text & "'"
    
    varTmp2 = Split(Trim(txt(0).Text), ",")
    strTmp = ""
    
    For lngLoop = 0 To UBound(varTmp2)
        If InStr(varTmp2(lngLoop), "-") = 0 Then
            strTmp = strTmp & "  OR 年龄='" & varTmp2(lngLoop) & "'"
        Else
            strTmp = strTmp & "  OR (年龄>='" & Mid(varTmp2(lngLoop), 1, InStr(varTmp2(lngLoop), "-") - 1) & "' AND 年龄<='" & Mid(varTmp2(lngLoop), InStr(varTmp2(lngLoop), "-") + 1) & "')"
        End If
    Next
    
'    For mlngLoop = 0 To UBound(varTmp2)
'
'        If InStr(varTmp2(mlngLoop), "-") = 0 Then
'
'            'Call GetBirth(Val(varTmp2(mlngLoop)), strStart, strEnd)
'            strTmp = strTmp & " OR (年龄>='" & strStart & "' AND 出生日期<='" & strEnd & "')"
'        Else
'
'            'Call GetBirth(Val(Mid(varTmp2(mlngLoop), 1, InStr(varTmp2(mlngLoop), "-") - 1)), strStart, strEnd)
'            strTmp = strTmp & " OR (出生日期<='" & strEnd & "'"
'
'            'Call GetBirth(Val(Mid(varTmp2(mlngLoop), InStr(varTmp2(mlngLoop), "-") + 1)), strStart, strEnd)
'            strTmp = strTmp & " AND 出生日期>='" & strStart & "')"
'
'        End If
'    Next
    If strTmp <> "" Then strFilter = strFilter & " AND (" & Mid(strTmp, 5) & ")"
        
    mrsMember.Filter = ""
    If strFilter <> "" Then mrsMember.Filter = Mid(strFilter, 6)
                                
    If mrsMember.RecordCount > 0 Then
        mrsMember.MoveFirst
        Do While Not mrsMember.EOF
            mrsMember.Update "选择", 1
            mrsMember.Update "前景色", 16711680
            mrsMember.MoveNext
        Loop
        
    End If
    
    Call cbo_Click(3)
    
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 4 Then
        Select Case KeyCode
        Case vbKeyS
            If tbrThis.Buttons("保存").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("保存"))
        Case vbKeyR
            If tbrThis.Buttons("恢复").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("恢复"))
        Case vbKeyH
            If tbrThis.Buttons("帮助").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("帮助"))
        Case vbKeyX
            If tbrThis.Buttons("退出").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("退出"))
        End Select
    ElseIf Shift = 0 Then
        If KeyCode = vbKeyEscape Then
            If tbrThis.Buttons("退出").Enabled Then Call tbrThis_ButtonClick(tbrThis.Buttons("退出"))
        End If
    End If
End Sub

'（３）窗体及其控件的事件处理******************************************************************************************
Private Sub Form_Load()
    glngFormW = 9765
    glngFormH = 6270
    If Not InDesign Then
        glngOld = GetWindowLong(Me.hWnd, GWL_WNDPROC)
        Call SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf Custom_WndMessage)
    End If
    
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
        
    With fra
        .Left = 0
        .Top = IIf(cbrThis.Visible, cbrThis.Height, 0) - 90
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
    End With
    
    With fra2
        .Left = fra.Left + fra.Width
        .Top = IIf(cbrThis.Visible, cbrThis.Height, 0) - 90
        .Width = Me.ScaleWidth - .Left
    End With
    
    With vsf
        .Left = fra2.Left
        .Top = fra2.Top + fra2.Height
        .Width = fra2.Width
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0) - fra1.Height + 90
    End With
    
    With fra1
        .Left = vsf.Left
        .Top = vsf.Top + vsf.Height - 90
        .Width = vsf.Width
    End With
    
    vsf.AppendRow = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mnuFileSave.Enabled Then
        Cancel = (MsgBox("数据必须保存后才生效，是否不保存就退出？", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo)
        If Cancel Then Exit Sub
    End If
    Call SaveWinState(Me, App.ProductName)
    
    If Not InDesign Then
        Call SetWindowLong(Me.hWnd, GWL_WNDPROC, glngOld)
    End If
    
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileRestore_Click()
    
    If MsgBox("确实要恢复以前所选项目吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    Call ClearData
    
    Call ReadData
        
    EditChanged = False
    
End Sub

Private Sub mnuFileSave_Click()
    
    If SaveEdit() Then
                
        On Error Resume Next
        
        EditChanged = False
        mblnOK = True
        
        Unload Me
        
    End If
    
End Sub

Private Sub mnuHelpAbout_Click()
    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
End Sub

Private Sub mnuHelpTopic_Click()
   Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hWnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hWnd)
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    cbrThis.Visible = mnuViewToolButton.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim intLoop As Integer
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For intLoop = 1 To tbrThis.Buttons.Count
        tbrThis.Buttons(intLoop).Caption = IIf(mnuViewToolText.Checked, tbrThis.Buttons(intLoop).Tag, "")
    Next
    cbrThis.Bands(1).MinHeight = tbrThis.Height
    Call Form_Resize
    
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "保存"
        Call mnuFileSave_Click
    Case "恢复"
        Call mnuFileRestore_Click
    Case "帮助"
        Call mnuHelpTopic_Click
    Case "退出"
        Call mnuFileExit_Click
    End Select
End Sub

Private Sub txt_GotFocus(Index As Integer)
    zlControl.TxtSelAll txt(Index)
    Select Case Index
    Case 1
        zlCommFun.OpenIme True
    End Select
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
        Select Case Index
        Case 0      '
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            KeyAscii = FilterKeyAscii(KeyAscii, 99, "0123456789-,")
        End Select
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
    Select Case Index
    Case 1
        zlCommFun.OpenIme False
    End Select
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = 8 Then
        mrsMember.Filter = ""
        mrsMember.Filter = "ID=" & Val(vsf.RowData(Row))
        
        If mrsMember.RecordCount > 0 Then
            mrsMember("体检组别").Value = vsf.TextMatrix(Row, Col)
            EditChanged = True
        End If
    End If
End Sub

Private Sub vsf_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub vsf_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Cancel = True
End Sub



Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

