VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPersonEdit 
   Caption         =   "受检人员"
   ClientHeight    =   5865
   ClientLeft      =   2775
   ClientTop       =   3765
   ClientWidth     =   10095
   Icon            =   "frmPersonEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   10095
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chk 
      Caption         =   "保存的同时进行报到处理(&3)"
      Height          =   195
      Left            =   3600
      TabIndex        =   0
      Top             =   105
      Width           =   3150
   End
   Begin VB.ComboBox cbo 
      Height          =   300
      Left            =   675
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   30
      Width           =   2745
   End
   Begin VB.PictureBox picButton 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   75
      ScaleHeight     =   615
      ScaleWidth      =   10650
      TabIndex        =   9
      Top             =   4740
      Width           =   10650
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   8115
         TabIndex        =   12
         Top             =   135
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   9315
         TabIndex        =   11
         Top             =   135
         Width           =   1100
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "帮助(&H)"
         Height          =   350
         Left            =   90
         TabIndex        =   10
         Top             =   135
         Width           =   1100
      End
   End
   Begin VB.Frame fra2 
      Height          =   3615
      Left            =   645
      TabIndex        =   7
      Top             =   315
      Width           =   6210
      Begin zl9Medical.VsfGrid vsfPerson 
         Height          =   3045
         Left            =   45
         TabIndex        =   8
         Top             =   150
         Width           =   4755
         _extentx        =   8387
         _extenty        =   5371
      End
   End
   Begin VB.CommandButton cmd 
      Height          =   345
      Index           =   11
      Left            =   8700
      Picture         =   "frmPersonEdit.frx":076A
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "更多资料信息"
      Top             =   0
      Width           =   345
   End
   Begin VB.CommandButton cmd 
      Height          =   345
      Index           =   14
      Left            =   8265
      Picture         =   "frmPersonEdit.frx":15AC
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "将信息写回IC卡"
      Top             =   0
      Width           =   345
   End
   Begin VB.CommandButton cmd 
      Height          =   345
      Index           =   15
      Left            =   7875
      Picture         =   "frmPersonEdit.frx":7DFE
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "从IC卡读信息"
      Top             =   0
      Width           =   345
   End
   Begin VB.CommandButton cmd 
      Height          =   345
      Index           =   16
      Left            =   8250
      Picture         =   "frmPersonEdit.frx":E650
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "单位人员选择"
      Top             =   525
      Width           =   345
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   5505
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPersonEdit.frx":14EA2
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12726
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
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
   Begin MSComctlLib.ImageList ils16 
      Left            =   4020
      Top             =   4110
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonEdit.frx":15736
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonEdit.frx":1A7A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonEdit.frx":1AA9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonEdit.frx":1B034
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonEdit.frx":1B5CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPersonEdit.frx":1B728
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "&1.组别"
      Height          =   180
      Index           =   4
      Left            =   60
      TabIndex        =   14
      Top             =   75
      Width           =   540
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "&2.项目"
      Height          =   180
      Index           =   0
      Left            =   60
      TabIndex        =   13
      Top             =   405
      Width           =   540
   End
End
Attribute VB_Name = "frmPersonEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'（１）窗体级变量定义**************************************************************************************************
Private mblnStartUp As Boolean                          '窗体启动标志
Private mblnOK As Boolean
Private mfrmMain As Object
Private mlngLoop As Long
Private mblnDataChange As Boolean
Private mrsPersons As New ADODB.Recordset                 '用于暂时体检人员
Private mblnChanged As Boolean
Private mblnGroup As Boolean
Private mlngGroup As Long
Private mstrGroup As String
Private mbytMode As Byte
Private mlngKey As Long
Private mblnNo As Boolean
Private mblnRegister As Boolean

Private Enum mPersonCol
    姓名 = 1
    门诊号
    健康号
    性别
    年龄
    婚姻状况
    出生日期
    身份证
    民族
    国籍
    学历
    职业
    身份
    联系人姓名
    联系人电话
    电子邮件
    联系人地址
    工作单位
    病人id
    IC卡号
    就诊卡号
    前景色
    新加
End Enum

'（２）自定义过程或函数************************************************************************************************
'设置人员缺省值
Private Function SetDefault(ByVal intRow As Integer) As Boolean
    
    '先按上行读取
    With vsfPerson
        If intRow > 1 Then
            .TextMatrix(intRow, mPersonCol.性别) = .TextMatrix(intRow - 1, mPersonCol.性别)
            .TextMatrix(intRow, mPersonCol.婚姻状况) = .TextMatrix(intRow - 1, mPersonCol.婚姻状况)
        End If
        
    End With
    
End Function

Private Function CountGroup() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:按组别统计项目数量、人数（男、女）
    '参数:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim strTmp As String
    Dim lngCount1 As Long
    Dim lngCount2 As Long
    
    If mblnGroup Then
        strTmp = """" & cbo.Text & """组别下"
    End If
    
    If mblnGroup Then
        lngCount1 = 0
        lngCount2 = 0
        
        For lngLoop = 1 To vsfPerson.Rows - 1
            If Trim(vsfPerson.TextMatrix(lngLoop, mPersonCol.姓名)) <> "" Then
                If InStr(vsfPerson.TextMatrix(lngLoop, mPersonCol.性别), "男") > 0 Then
                    lngCount1 = lngCount1 + 1
                Else
                    lngCount2 = lngCount2 + 1
                End If
            End If
        Next
        
        strTmp = strTmp & "共有人员" & lngCount1 + lngCount2 & "个(男性:" & lngCount1 & "个,女性:" & lngCount2 & "个)"
    End If
    
    stbThis.Panels(2).Text = strTmp
    
End Function

Private Function CheckHavePerson(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  检查是否有重复的项目
    '参数:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    
    For lngLoop = 1 To vsfPerson.Rows - 1
        If Val(vsfPerson.TextMatrix(lngLoop, mPersonCol.病人id)) = lngKey And vsfPerson.Row <> lngLoop And Val(vsfPerson.TextMatrix(lngLoop, mPersonCol.病人id)) > 0 Then
            CheckHavePerson = True
            Exit Function
        End If
    Next
End Function

Private Property Let DataChange(ByVal vData As Boolean)
        mblnDataChange = vData
End Property

Private Property Get DataChange() As Boolean
        DataChange = mblnDataChange
End Property

Private Function GetPatientInfo(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：
    '参数：
    '返回：
    '------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    
    strSQL = "SELECT A.* FROM 病人信息 A WHERE A.病人id=[1]"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngKey)
    If rs.BOF = False Then
        

        If mlngGroup <> Val(zlCommFun.NVL(rs("合同单位id"))) And Val(zlCommFun.NVL(rs("合同单位id"))) > 0 And mlngGroup > 0 Then

            If MsgBox("不是当前团体的人员，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function

        End If
        
        vsfPerson.EditText = zlCommFun.NVL(rs("姓名"))
        vsfPerson.Cell(flexcpData, vsfPerson.Row, vsfPerson.Col) = zlCommFun.NVL(rs("姓名").Value)
        vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.姓名) = zlCommFun.NVL(rs("姓名"))
        
        Call SetDefault(vsfPerson.Row)
        
        vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.门诊号) = zlCommFun.NVL(rs("门诊号"))
        
        vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.年龄) = zlCommFun.NVL(rs("年龄"))
        vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.身份证) = zlCommFun.NVL(rs("身份证号"))
        vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.出生日期) = Format(zlCommFun.NVL(rs("出生日期")), "yyyy-MM-dd")
        vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.性别) = zlCommFun.NVL(rs("性别").Value)
        vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.婚姻状况) = zlCommFun.NVL(rs("婚姻状况").Value)
        vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.病人id) = zlCommFun.NVL(rs("病人id"))
        
        DataChange = True
        
    End If
    
    GetPatientInfo = True
    
End Function


Public Function ShowEdit(ByVal frmMain As Object, _
                        ByVal lngKey As Long, _
                        ByRef rsPersons As ADODB.Recordset, _
                        Optional blnGroup As Boolean = False, _
                        Optional ByVal bytMode As Byte = 1, _
                        Optional lngGroup As Long, _
                        Optional ByRef blnRegister As Boolean = True) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  显示编辑窗体，是与调用窗体的接口函数
    '参数:  frmMain         调用窗体对象
    '       lngKey          预约登记id
    '返回:  True
    '       False
    '------------------------------------------------------------------------------------------------------------------
    Dim varGroup As Variant

    mblnNo = True
    mblnStartUp = True
    mblnOK = False

    Set mfrmMain = frmMain
    
    mstrGroup = ""
    
    mlngGroup = lngGroup
    mlngKey = lngKey
    Call CopyRecord(rsPersons, mrsPersons)
    mblnGroup = blnGroup
    mbytMode = bytMode

    Call ClearData
    If InitData = False Then Exit Function
    If ReadData() = False Then Exit Function

    DataChange = False
    
    mblnNo = False
    
    Call cbo_Click
    
    Call vsfPerson_AfterRowColChange(0, 0, vsfPerson.Row, vsfPerson.Col)
    
    Me.Show 1, frmMain

    rsPersons.Filter = ""
    If mblnOK Then Call CopyRecord(mrsPersons, rsPersons)
    blnRegister = mblnRegister
    
    ShowEdit = mblnOK

End Function

Private Function ClearData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '参数:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long


    cbo.Clear
    Call ResetVsf(vsfPerson)

'    vsfPerson.AppendRow = True

    DataChange = False


End Function

Private Function InitData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  初始化设置
    '返回:  True        初始化成功
    '       False       初始化失败
    '------------------------------------------------------------------------------------------------------------------

    On Error GoTo errHand
    
    chk.Visible = (mbytMode = 2)
    chk.Value = Val(GetSetting("ZLSOFT", "私有模块\" & Me.Name, "报到处理", "1"))
    
    cbo.AddItem "缺省"

    With vsfPerson
        .Cols = 0
        .NewColumn "", 255, 4
        .NewColumn "姓名", 1080, 1, "...", 1, GetMaxLength("病人信息", "姓名")
        .NewColumn "门诊号", 810, 1
        .NewColumn "健康号", 900, 1, , 1
        .NewColumn "性别", 750, 1, GetCombList("SELECT 名称 FROM 性别"), 1, GetMaxLength("病人信息", "性别")
        .NewColumn "年龄", 600, 1, , 1, GetMaxLength("病人信息", "年龄")
        .NewColumn "婚姻状况", 900, 1, GetCombList("SELECT 名称 FROM 婚姻状况"), 1, GetMaxLength("病人信息", "婚姻状况")
        .NewColumn "出生日期", 990, 1, , 1
        .NewColumn "身份证", 1800, 1, , 1, GetMaxLength("病人信息", "身份证号")
        .NewColumn "民族", 0, 1, , , GetMaxLength("病人信息", "民族")
        .NewColumn "国籍", 0, 1, , , GetMaxLength("病人信息", "国籍")
        .NewColumn "学历", 0, 1, , , GetMaxLength("病人信息", "学历")
        .NewColumn "职业", 0, 1, , , GetMaxLength("病人信息", "职业")
        .NewColumn "身份", 0, 1, , , GetMaxLength("病人信息", "身份")
        .NewColumn "联系人姓名", 0, 1, , , GetMaxLength("病人信息", "联系人姓名")
        .NewColumn "联系人电话", 0, 1, , , GetMaxLength("病人信息", "联系人电话")
        .NewColumn "电子邮件", 0, 1, , , GetMaxLength("病人信息", "电子邮件")
        .NewColumn "联系人地址", 0, 1, , , GetMaxLength("病人信息", "联系人地址")
        .NewColumn "工作单位", 0, 1, , , GetMaxLength("病人信息", "工作单位")
        .NewColumn "病人id", 0, 1
        .NewColumn "IC卡号", 0, 1
        .NewColumn "就诊卡号", 0, 1
        
        .NewColumn "前景色", 0, 1
        .NewColumn "新加", 0, 1
'        .NewColumn "", 15, 1
'        .ExtendLastCol = True
        .FixedCols = 1
        .Body.GridColor = &HC1C1C1
        .Body.GridColorFixed = &HC1C1C1
'        .AppendRow = True
        
        .Body.ColEditMask(mPersonCol.出生日期) = "0000-00-00"
        
    End With

    If mblnGroup = False Then
        cbo.Visible = False
        lbl(4).Visible = False
    End If
    
    InitData = True

    Exit Function

errHand:
    If ErrCenter = 1 Then
        Resume
    End If
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


    '读取体检组别及体检项目
    
    mblnNo = True
    
    cbo.Clear

    gstrSQL = "SELECT A.组别名称 AS 组别, rownum AS ID FROM 体检组别 A WHERE A.登记id=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey)
    If rs.BOF = False Then
        Do While Not rs.EOF
            cbo.AddItem rs("组别").Value
            rs.MoveNext
        Loop
    Else
        cbo.AddItem "缺省"
    End If

    '读取体检项目
    
    If cbo.ListCount > 0 Then cbo.ListIndex = 0
    
    mblnNo = False
    
    Call cbo_Click

    ReadData = True

    Exit Function

errHand:

    If ErrCenter = 1 Then Resume

End Function

Private Function CheckHave(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  检查是否有重复的项目
    '参数:
    '返回:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long

    For lngLoop = 1 To vsfPerson.Rows - 1
        If Val(vsfPerson.RowData(lngLoop)) = lngKey And vsfPerson.Row <> lngLoop Then
            CheckHave = True
            Exit Function
        End If
    Next
End Function


Private Function SaveItems(ByVal strGroup As String) As Boolean

    Dim lngLoop As Long

    On Error GoTo errHand

    '保存所选择的检验项目
    mrsPersons.Filter = ""
    mrsPersons.Filter = "组别='" & strGroup & "' AND 删除<>'1'"

    Call DeleteRecord(mrsPersons)

    For lngLoop = 1 To vsfPerson.Rows - 1

        If vsfPerson.TextMatrix(lngLoop, mPersonCol.姓名) <> "" Then
            mrsPersons.AddNew

            mrsPersons("组别").Value = strGroup
            mrsPersons("病人id").Value = vsfPerson.TextMatrix(lngLoop, mPersonCol.病人id)
            mrsPersons("IC卡号").Value = vsfPerson.TextMatrix(lngLoop, mPersonCol.IC卡号)
            mrsPersons("健康号").Value = vsfPerson.TextMatrix(lngLoop, mPersonCol.健康号)
            mrsPersons("姓名").Value = vsfPerson.TextMatrix(lngLoop, mPersonCol.姓名)
            mrsPersons("门诊号").Value = vsfPerson.TextMatrix(lngLoop, mPersonCol.门诊号)
            mrsPersons("身份证").Value = vsfPerson.TextMatrix(lngLoop, mPersonCol.身份证)
            mrsPersons("性别").Value = vsfPerson.TextMatrix(lngLoop, mPersonCol.性别)
            mrsPersons("出生日期").Value = vsfPerson.TextMatrix(lngLoop, mPersonCol.出生日期)
            mrsPersons("婚姻状况").Value = vsfPerson.TextMatrix(lngLoop, mPersonCol.婚姻状况)
            mrsPersons("年龄").Value = vsfPerson.TextMatrix(lngLoop, mPersonCol.年龄)
            mrsPersons("民族").Value = vsfPerson.TextMatrix(lngLoop, mPersonCol.民族)
            mrsPersons("国籍").Value = vsfPerson.TextMatrix(lngLoop, mPersonCol.国籍)
            mrsPersons("学历").Value = vsfPerson.TextMatrix(lngLoop, mPersonCol.学历)
            mrsPersons("职业").Value = vsfPerson.TextMatrix(lngLoop, mPersonCol.职业)
            mrsPersons("身份").Value = vsfPerson.TextMatrix(lngLoop, mPersonCol.身份)
            mrsPersons("联系人姓名").Value = vsfPerson.TextMatrix(lngLoop, mPersonCol.联系人姓名)
            mrsPersons("联系人电话").Value = vsfPerson.TextMatrix(lngLoop, mPersonCol.联系人电话)
            mrsPersons("电子邮件").Value = vsfPerson.TextMatrix(lngLoop, mPersonCol.电子邮件)
            mrsPersons("联系人地址").Value = vsfPerson.TextMatrix(lngLoop, mPersonCol.联系人地址)
            mrsPersons("工作单位").Value = vsfPerson.TextMatrix(lngLoop, mPersonCol.工作单位)
'            mrsPersons("就诊卡号").Value = vsfPerson.TextMatrix(lngLoop, mPersonCol.就诊卡号)
            mrsPersons("新加").Value = vsfPerson.TextMatrix(lngLoop, mPersonCol.新加)
            mrsPersons("前景色").Value = vsfPerson.TextMatrix(lngLoop, mPersonCol.前景色)

        End If
    Next

    SaveItems = True

errHand:

End Function

Private Function ValidEdit() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:  校验数据的有效性
    '返回:  True        数据有效
    '       False       数据无效
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long

    For lngLoop = 1 To vsfPerson.Rows - 1
        
        If vsfPerson.EditMode(mPersonCol.姓名) = 1 Then
            If StrIsValid(vsfPerson.TextMatrix(lngLoop, mPersonCol.姓名), GetMaxLength("病人信息", "姓名")) = False Then
                vsfPerson.Row = lngLoop
                vsfPerson.Col = mPersonCol.姓名
                vsfPerson.ShowCell vsfPerson.Row, vsfPerson.Col
                vsfPerson.SetFocus
                
                Exit Function
            End If
            
            If StrIsValid(vsfPerson.TextMatrix(lngLoop, mPersonCol.身份证), GetMaxLength("病人信息", "身份证号")) = False Then
                vsfPerson.Row = lngLoop
                vsfPerson.Col = mPersonCol.身份证
                vsfPerson.ShowCell vsfPerson.Row, vsfPerson.Col
                vsfPerson.SetFocus
                
                Exit Function
            End If
                        
            If StrIsValid(vsfPerson.TextMatrix(lngLoop, mPersonCol.婚姻状况), GetMaxLength("病人信息", "婚姻状况")) = False Then
                vsfPerson.Row = lngLoop
                vsfPerson.Col = mPersonCol.婚姻状况
                vsfPerson.ShowCell vsfPerson.Row, vsfPerson.Col
                vsfPerson.SetFocus
                
                Exit Function
            End If
            
            If StrIsValid(vsfPerson.TextMatrix(lngLoop, mPersonCol.电子邮件), GetMaxLength("体检人员档案", "电子邮件")) = False Then
                vsfPerson.Row = lngLoop
                vsfPerson.Col = mPersonCol.电子邮件
                vsfPerson.ShowCell vsfPerson.Row, vsfPerson.Col
                vsfPerson.SetFocus
                
                Exit Function
            End If
            
            If StrIsValid(vsfPerson.TextMatrix(lngLoop, mPersonCol.性别), GetMaxLength("病人信息", "性别")) = False Then
                vsfPerson.Row = lngLoop
                vsfPerson.Col = mPersonCol.性别
                vsfPerson.ShowCell vsfPerson.Row, vsfPerson.Col
                vsfPerson.SetFocus
                
                Exit Function
            End If
            
            If Trim(vsfPerson.TextMatrix(lngLoop, mPersonCol.出生日期)) <> "" Then
                
                If CheckStrValid(vsfPerson.TextMatrix(lngLoop, mPersonCol.出生日期), CHECKFORMAT.日期) = False Then
                    vsfPerson.Row = lngLoop
                    vsfPerson.Col = mPersonCol.出生日期
                    vsfPerson.ShowCell vsfPerson.Row, vsfPerson.Col
                    vsfPerson.SetFocus
                    
                    Exit Function
                End If
            End If
            
                        
            If CheckStrValid(vsfPerson.TextMatrix(lngLoop, mPersonCol.电子邮件), CHECKFORMAT.电子邮件) = False Then
                    vsfPerson.Row = lngLoop
                    vsfPerson.Col = mPersonCol.电子邮件
                    vsfPerson.ShowCell vsfPerson.Row, vsfPerson.Col
                    vsfPerson.SetFocus
                    
                    Exit Function
            End If
                
            If CheckStrValid(vsfPerson.TextMatrix(lngLoop, mPersonCol.身份证), CHECKFORMAT.身份证号) = False Then
                    vsfPerson.Row = lngLoop
                    vsfPerson.Col = mPersonCol.身份证
                    vsfPerson.ShowCell vsfPerson.Row, vsfPerson.Col
                    vsfPerson.SetFocus
                    
                    Exit Function
            End If
        End If
    Next
    
    ValidEdit = True

End Function

Private Function ReadItems(ByVal strGroup As String) As Boolean

    mrsPersons.Filter = ""
    mrsPersons.Filter = "组别='" & strGroup & "' AND 删除<>'1'"
    If mrsPersons.RecordCount > 0 Then
        mrsPersons.MoveFirst
        Call FillGrid(vsfPerson, mrsPersons)
    End If

    ReadItems = True

End Function

Private Sub cbo_Click()
    If mblnNo Then Exit Sub
    
    If mstrGroup <> cbo.Text Then
        Call SaveItems(mstrGroup)
        
        mstrGroup = cbo.Text
        
        Call ResetVsf(vsfPerson)
        Call ReadItems(mstrGroup)
    End If
    
End Sub

Private Sub cbo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then

        zlCommFun.PressKey vbKeyTab

    End If
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim clsCard As Object
    Dim strInfo() As String
    Dim lngLoop As Long
    Dim strParam As String
    Dim varParam As Variant
    Dim strItem As String
    Dim strValue As String
    Dim strCardNo1 As String
    Dim strCardNo2 As String
    
    On Error GoTo errHand
    Select Case Index
    Case 11
        
        strParam = vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.病人id) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.姓名) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.身份证) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.性别) & "'"
        strParam = strParam & Format(vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.出生日期), "yyyy-MM-dd") & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.婚姻状况) & "'"
        
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.民族) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.国籍) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.学历) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.职业) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.身份) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.联系人姓名) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.联系人电话) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.电子邮件) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.联系人地址) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.工作单位) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.年龄) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.健康号)
                
        If frmPatientEdit.ShowEdit(Me, strParam) Then
            varParam = Split(strParam, "'")
            
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.姓名) = varParam(1)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.身份证) = varParam(2)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.性别) = varParam(3)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.出生日期) = varParam(4)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.婚姻状况) = varParam(5)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.病人id) = Val(varParam(0))
            
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.民族) = varParam(6)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.国籍) = varParam(7)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.学历) = varParam(8)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.职业) = varParam(9)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.身份) = varParam(10)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.联系人姓名) = varParam(11)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.联系人电话) = varParam(12)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.电子邮件) = varParam(13)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.联系人地址) = varParam(14)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.工作单位) = varParam(15)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.年龄) = varParam(16)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.健康号) = varParam(17)
            
            
        End If
        
        If vsfPerson.Visible Then vsfPerson.SetFocus
        
    Case 14    '写卡
                    
        Set clsCard = CreateObject("zl9ICCard.clsICCard")
        If Not (clsCard Is Nothing) Then
            
            ReDim strInfo(1 To 16)
            
            strCardNo1 = clsCard.GetCardNo
            strCardNo2 = vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.IC卡号)
            If strCardNo2 <> "" Then
                '病人有卡，但和当前的卡不是同一张卡
                If strCardNo1 <> strCardNo2 Then
                    ShowSimpleMsg "此卡不是当前病人的卡！"
                    Exit Sub
                End If
            Else
                '病人没有卡
                
                If strCardNo1 = "" Then
                
                    '新卡，自动开卡
                    strCardNo1 = "11111111"
                    strCardNo2 = strCardNo1
                    
                    '写卡号
                    If clsCard.SetCardNo(strCardNo1) = False Then Exit Sub
                    vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.IC卡号) = strCardNo2
                    
                Else
                
                    '不是新卡
                    ShowSimpleMsg "此卡不是新卡，不能进行写入操作！"
                    Exit Sub
                End If
            End If
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.IC卡号) = strCardNo2
            
            strInfo(1) = "姓名=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.姓名)
            strInfo(2) = "身份证号=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.身份证)
            strInfo(3) = "性别=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.性别)
            strInfo(4) = "出生日期=" & Format(vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.出生日期), "yyyy-MM-dd")
            strInfo(5) = "婚姻状况=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.婚姻状况)
            strInfo(6) = "民族=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.民族)
            strInfo(7) = "国籍=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.国籍)
            strInfo(8) = "学历=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.学历)
            strInfo(9) = "职业=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.职业)
            strInfo(10) = "身份=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.身份)
            strInfo(11) = "联系人姓名=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.联系人姓名)
            strInfo(12) = "联系人电话=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.联系人电话)
            strInfo(13) = "联系人地址=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.联系人地址)
            strInfo(14) = "电子邮件=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.电子邮件)
            strInfo(15) = "工作单位=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.工作单位)
            strInfo(16) = "年龄=" & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.年龄)
                                    
            If clsCard.SetPatient(strInfo) Then
                ShowSimpleMsg "更新当前病人信息成功！"
            End If
        End If
        If vsfPerson.Visible Then vsfPerson.SetFocus
    Case 15    '读卡
        
        Set clsCard = CreateObject("zl9ICCard.clsICCard")
        If Not (clsCard Is Nothing) Then
            
            strCardNo1 = clsCard.GetCardNo
            strCardNo2 = vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.IC卡号)
            
            If strCardNo2 <> "" Then
                '记录的病人有卡，但和当前的卡不是同一张卡
                If strCardNo1 <> strCardNo2 Then
                    ShowSimpleMsg "此卡不是当前病人的卡！"
                    Exit Sub
                End If
            Else
            
                '病人没有卡，则将当前的卡号付给病人
                strCardNo2 = strCardNo1
                                
            End If
            
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.IC卡号) = strCardNo2
            
            If GetPatientID(strCardNo2) > 0 Then
                
                '在系统中找到了病人
                
                vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.病人id) = GetPatientID(strCardNo2)
                Call GetPatientInfo(Val(vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.病人id)))
                
            ElseIf clsCard.GetPatient(strInfo) Then
                For lngLoop = LBound(strInfo) To UBound(strInfo)
                    If InStr(strInfo(lngLoop), "=") > 0 Then
                        strItem = Mid(strInfo(lngLoop), 1, InStr(strInfo(lngLoop), "=") - 1)
                        strValue = Mid(strInfo(lngLoop), InStr(strInfo(lngLoop), "=") + 1)
                        
                        Select Case strItem
                        Case "姓名"
                        
                            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.姓名) = strValue
                            
                        Case "身份证号"
                            
                            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.身份证) = strValue
                                                        
                        Case "性别"
                            
                            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.性别) = strValue
                            
                        Case "出生日期"
                            
                            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.出生日期) = strValue
                            
                        Case "婚姻状况"
                            
                            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.婚姻状况) = strValue
                            
                        Case "民族"
                            
                            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.民族) = strValue
                            
                        Case "国籍"
                            
                            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.国籍) = strValue
                            
                        Case "学历"
                            
                            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.学历) = strValue
                            
                        Case "职业"
                        
                            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.职业) = strValue
                            
                        Case "身份"
                            
                            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.身份) = strValue
                            
                        Case "联系人姓名"
                            
                            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.联系人姓名) = strValue
                            
                        Case "联系人电话"
                            
                            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.联系人电话) = strValue
                            
                        Case "联系人地址"
                            
                            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.联系人地址) = strValue
                            
                        Case "电子邮件"
                            
                            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.电子邮件) = strValue
                            
                        Case "工作单位"
                            
                            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.工作单位) = strValue
                            
                        Case "年龄"
                            
                            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.年龄) = strValue
                            
                        End Select
                    End If
                Next
                
                
            End If
        End If
        If vsfPerson.Visible Then vsfPerson.SetFocus
    '------------------------------------------------------------------------------------------------------------------
    Case 16
        '选择单位人员
        Dim rsData As New ADODB.Recordset
        Dim rs As New ADODB.Recordset
        
        If frmSelectGroupPerson.ShowFilter(Me, mlngGroup, rs) Then
            rs.Filter = 0
            rs.Filter = "选择=1"
            If rs.RecordCount > 0 Then

                If Val(vsfPerson.RowData(1)) > 0 Then
                    If MsgBox("是否要清除已选择的受检人员？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                        Call ResetVsf(vsfPerson)
                    End If
                End If

                rs.MoveFirst

                Do While Not rs.EOF
                    
                    If CheckHavePerson(rs("ID").Value) = False Then
                        With vsfPerson
                        
                            .Row = .Rows - 1
                            If Val(.RowData(.Row)) > 0 Then
                                .Rows = .Rows + 1
                                .Row = .Rows - 1
                            End If
            
                            .TextMatrix(.Row, mPersonCol.姓名) = zlCommFun.NVL(rs("姓名").Value)
                            .TextMatrix(.Row, mPersonCol.门诊号) = zlCommFun.NVL(rs("门诊号").Value)
                            .TextMatrix(.Row, mPersonCol.健康号) = zlCommFun.NVL(rs("健康号").Value)
                            .TextMatrix(.Row, mPersonCol.性别) = zlCommFun.NVL(rs("性别").Value)
                            .TextMatrix(.Row, mPersonCol.年龄) = zlCommFun.NVL(rs("年龄").Value)
                            .TextMatrix(.Row, mPersonCol.婚姻状况) = zlCommFun.NVL(rs("婚姻状况").Value)
                            .TextMatrix(.Row, mPersonCol.出生日期) = zlCommFun.NVL(rs("出生日期").Value)
                            .TextMatrix(.Row, mPersonCol.身份证) = zlCommFun.NVL(rs("身份证号").Value)
                            .TextMatrix(.Row, mPersonCol.民族) = zlCommFun.NVL(rs("民族").Value)
                            .TextMatrix(.Row, mPersonCol.国籍) = zlCommFun.NVL(rs("国籍").Value)
                            .TextMatrix(.Row, mPersonCol.学历) = zlCommFun.NVL(rs("学历").Value)
                            .TextMatrix(.Row, mPersonCol.职业) = zlCommFun.NVL(rs("职业").Value)
                            .TextMatrix(.Row, mPersonCol.身份) = zlCommFun.NVL(rs("身份").Value)
                            .TextMatrix(.Row, mPersonCol.联系人姓名) = zlCommFun.NVL(rs("联系人姓名").Value)
                            .TextMatrix(.Row, mPersonCol.联系人电话) = zlCommFun.NVL(rs("联系人电话").Value)
'                            .TextMatrix(.Row, mPersonCol.电子邮件) = zlCommFun.NVL(rs("电子邮件").Value)
                            .TextMatrix(.Row, mPersonCol.联系人地址) = zlCommFun.NVL(rs("联系人地址").Value)
                            .TextMatrix(.Row, mPersonCol.工作单位) = zlCommFun.NVL(rs("工作单位").Value)
                            .TextMatrix(.Row, mPersonCol.病人id) = zlCommFun.NVL(rs("ID").Value, 0)
                            .TextMatrix(.Row, mPersonCol.IC卡号) = zlCommFun.NVL(rs("IC卡号").Value)
                            .TextMatrix(.Row, mPersonCol.就诊卡号) = zlCommFun.NVL(rs("就诊卡号").Value)
        
                            .RowData(.Row) = zlCommFun.NVL(rs("ID").Value)
                        End With
                    End If

                    rs.MoveNext

                Loop

                DataChange = True
            End If

        End If

        Call EnterFocus(vsfPerson)
    End Select
    
    Exit Sub
    
errHand:
    If ErrCenter = -1 Then
        Resume
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
   Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
End Sub

Private Sub cmdOK_Click()

    Dim lngKey As Long

    If Trim(cbo.Text) <> "" Then Call SaveItems(Trim(cbo.Text))

    If ValidEdit = False Then Exit Sub

    mrsPersons.Filter = ""

    mblnOK = True
    DataChange = False
    mblnRegister = (chk.Value = 1)
    
    Unload Me

End Sub


Private Sub Form_Load()

    glngFormW = 10770
    glngFormH = 6780
    If Not InDesign Then
        glngOld = GetWindowLong(Me.hWnd, GWL_WNDPROC)
        Call SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf Custom_WndMessage)
    End If
    
    Call RestoreWinState(Me, App.ProductName)

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
       
    With fra2
        .Left = 0
        .Top = -90 + IIf(cbo.Visible, cbo.Height + cbo.Top + 30, 0)
        .Width = Me.ScaleWidth - .Left
        .Height = Me.ScaleHeight - .Top - picButton.Height + 90 - stbThis.Height
    End With
       
    With picButton
        .Left = fra2.Left
        .Top = fra2.Top + fra2.Height
        .Width = fra2.Width
    End With
    
    
    With vsfPerson
        .Left = 45
        .Top = 120
        .Width = fra2.Width - .Left - 45
        .Height = fra2.Height - .Top - 45
    End With
    
    With cmd(11)
        .Left = Me.ScaleWidth - .Width - 45
    End With
    
    With cmd(16)
        .Left = cmd(11).Left - .Width - 45
        .Top = cmd(11).Top
    End With
    
    With cmd(14)
        .Left = cmd(16).Left - .Width - 45
    End With
    
    With cmd(15)
        .Left = cmd(14).Left - .Width - 45
    End With
      
    cmdCancel.Left = picButton.Width - cmdCancel.Width - 60
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 120

End Sub

Private Sub Form_Unload(Cancel As Integer)

    If DataChange Then
        Cancel = (MsgBox("数据必须保存后才生效，是否不保存就退出？", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo)
        If Cancel Then Exit Sub
    End If

    Call SaveWinState(Me, App.ProductName)
    SaveSetting "ZLSOFT", "私有模块\" & Me.Name, "报到处理", chk.Value
    
    If Not InDesign Then
        Call SetWindowLong(Me.hWnd, GWL_WNDPROC, glngOld)
    End If
End Sub

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)

    Dim strParam As String
    Dim varParam As Variant
        
    On Error GoTo errHand

    Select Case Button.Key
    Case "详细资料"
        
        strParam = vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.病人id) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.姓名) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.身份证) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.性别) & "'"
        strParam = strParam & Format(vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.出生日期), "yyyy-MM-dd") & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.婚姻状况) & "'"
        
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.民族) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.国籍) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.学历) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.职业) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.身份) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.联系人姓名) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.联系人电话) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.电子邮件) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.联系人地址) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.工作单位) & "'"
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.年龄) & "'"
        
        strParam = strParam & vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.健康号)
                        
        If frmPatientEdit.ShowEdit(Me, strParam) Then
            varParam = Split(strParam, "'")
            
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.姓名) = varParam(1)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.身份证) = varParam(2)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.性别) = varParam(3)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.出生日期) = varParam(4)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.婚姻状况) = varParam(5)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.病人id) = Val(varParam(0))
            
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.民族) = varParam(6)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.国籍) = varParam(7)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.学历) = varParam(8)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.职业) = varParam(9)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.身份) = varParam(10)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.联系人姓名) = varParam(11)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.联系人电话) = varParam(12)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.电子邮件) = varParam(13)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.联系人地址) = varParam(14)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.工作单位) = varParam(15)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.年龄) = varParam(16)
            vsfPerson.TextMatrix(vsfPerson.Row, mPersonCol.健康号) = varParam(17)
            
            
        End If

    End Select

    Exit Sub

errHand:
        If ErrCenter = 1 Then Resume
End Sub


Private Sub vsfPerson_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lngPos As Long
    
    If Col = mPersonCol.出生日期 Then
        If Trim(vsfPerson.TextMatrix(Row, Col)) <> "" Then
        
            vsfPerson.TextMatrix(Row, Col) = zlCommFun.AddDate(vsfPerson.TextMatrix(Row, Col))
            
            If IsDate(vsfPerson.TextMatrix(Row, Col)) = False Then
                vsfPerson.TextMatrix(Row, Col) = ""
            End If
        End If
    End If
    
    
    If Col = mPersonCol.电子邮件 Then
        If Trim(vsfPerson.TextMatrix(Row, Col)) <> "" Then
            
            lngPos = InStr(vsfPerson.TextMatrix(Row, Col), "@")
            
            If lngPos = 0 Then
                vsfPerson.TextMatrix(Row, Col) = ""
            Else
                If Trim(Mid(vsfPerson.TextMatrix(Row, Col), 1, lngPos - 1)) = "" Then
                    vsfPerson.TextMatrix(Row, Col) = ""
                ElseIf Trim(Mid(vsfPerson.TextMatrix(Row, Col), lngPos + 1)) = "" Then
                    vsfPerson.TextMatrix(Row, Col) = ""
                End If
            End If
        End If
    End If
    
    If Col = mPersonCol.性别 Then
        Call CountGroup
    End If
    
End Sub

Private Sub vsfPerson_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    
    '设置编辑状态
    If Val(vsfPerson.TextMatrix(NewRow, mPersonCol.新加)) = 1 Then
        
        If vsfPerson.EditMode(mPersonCol.姓名) <> 0 Then
            vsfPerson.EditMode(mPersonCol.姓名) = 0
            vsfPerson.EditMode(mPersonCol.身份证) = 0
            vsfPerson.EditMode(mPersonCol.出生日期) = 0
            vsfPerson.EditMode(mPersonCol.性别) = 0
            vsfPerson.EditMode(mPersonCol.年龄) = 0
            vsfPerson.EditMode(mPersonCol.婚姻状况) = 0
            
            vsfPerson.ComboList(mPersonCol.性别) = ""
            vsfPerson.ComboList(mPersonCol.婚姻状况) = ""
            vsfPerson.ComboList(mPersonCol.姓名) = ""
            
            cmd(11).Enabled = False
            cmd(14).Enabled = False
            cmd(15).Enabled = False
        End If
        
    Else
        If vsfPerson.EditMode(mPersonCol.姓名) <> 1 Then
            vsfPerson.EditMode(mPersonCol.姓名) = 1
            vsfPerson.EditMode(mPersonCol.身份证) = 1
            vsfPerson.EditMode(mPersonCol.年龄) = 1
            vsfPerson.EditMode(mPersonCol.出生日期) = 1
            vsfPerson.EditMode(mPersonCol.性别) = 1
            vsfPerson.EditMode(mPersonCol.婚姻状况) = 1
            
            vsfPerson.ComboList(mPersonCol.姓名) = "..."
            vsfPerson.ComboList(mPersonCol.性别) = GetCombList("SELECT 名称 FROM 性别")
            vsfPerson.ComboList(mPersonCol.婚姻状况) = GetCombList("SELECT 名称 FROM 婚姻状况")
            cmd(11).Enabled = True
            cmd(14).Enabled = True
            cmd(15).Enabled = True
        End If
      
    End If
End Sub

Private Sub vsfPerson_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    If Val(vsfPerson.TextMatrix(Row, mPersonCol.新加)) = 1 And mbytMode = 2 Then
        
        Cancel = True
        Exit Sub
        
    End If
    
End Sub

Private Sub vsfPerson_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    
    On Error Resume Next
    
    If Val(vsfPerson.TextMatrix(NewRow, mPersonCol.门诊号)) > 0 Then
        vsfPerson.EditMode(mPersonCol.门诊号) = 0
    Else
        vsfPerson.EditMode(mPersonCol.门诊号) = 1
    End If
End Sub

Private Sub vsfPerson_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim lngKey As Long
    Dim rs As New ADODB.Recordset

    If frmPatientFind.ShowFind(Me, lngKey) Then
        If lngKey > 0 Then

            gstrSQL = "SELECT A.* FROM 病人信息 A WHERE A.病人id=[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
            If rs.BOF = False Then

                If mlngGroup <> Val(zlCommFun.NVL(rs("合同单位id"))) And Val(zlCommFun.NVL(rs("合同单位id"))) > 0 And mlngGroup > 0 Then

                    If MsgBox("不是当前团体的人员，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub

                End If
                
                vsfPerson.EditText = zlCommFun.NVL(rs("姓名"))
                vsfPerson.Cell(flexcpData, Row, vsfPerson.Col) = zlCommFun.NVL(rs("姓名").Value)
                vsfPerson.TextMatrix(Row, mPersonCol.姓名) = zlCommFun.NVL(rs("姓名"))
                
                Call SetDefault(Row)
                
                vsfPerson.TextMatrix(Row, mPersonCol.门诊号) = zlCommFun.NVL(rs("门诊号"))
                vsfPerson.TextMatrix(Row, mPersonCol.健康号) = zlCommFun.NVL(rs("健康号"))
                vsfPerson.TextMatrix(Row, mPersonCol.年龄) = zlCommFun.NVL(rs("年龄"))
                vsfPerson.TextMatrix(Row, mPersonCol.身份证) = zlCommFun.NVL(rs("身份证号"))
                vsfPerson.TextMatrix(Row, mPersonCol.出生日期) = Format(zlCommFun.NVL(rs("出生日期")), "yyyy-MM-dd")
                vsfPerson.TextMatrix(Row, mPersonCol.性别) = zlCommFun.NVL(rs("性别").Value)
                vsfPerson.TextMatrix(Row, mPersonCol.婚姻状况) = zlCommFun.NVL(rs("婚姻状况").Value)
                vsfPerson.TextMatrix(Row, mPersonCol.病人id) = zlCommFun.NVL(rs("病人id"))
                
                DataChange = True

            End If

        End If
    End If

End Sub

Private Sub vsfPerson_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, ByVal ComboList As String, KeyCode As Integer, ByVal Shift As Integer, Cancel As Boolean)
    
    Dim strText As String
    Dim strInput As String
    Dim rs As New ADODB.Recordset
    Dim strSvrText As String
    Dim rsData As New ADODB.Recordset
    Dim blnCard As Boolean
    
    If Chr(KeyCode) = "'" Then KeyCode = 0
    
    If Col = mPersonCol.姓名 Then
        
        strText = vsfPerson.EditText
        If KeyCode <> 8 And KeyCode <> 13 Then
            strText = strText & Chr(KeyCode)
        End If
        
        If InStr(vsfPerson.EditText, "'") > 0 Then
            KeyCode = 0
            ShowSimpleMsg "在个人姓名中有非法字符 ' ！"
            vsfPerson.EditText = ""
            vsfPerson.TextMatrix(Row, Col) = vsfPerson.Cell(flexcpData, Row, Col)
            Cancel = True
            Exit Sub
        End If
        
        blnCard = InputIsCard(vsfPerson.EditText, KeyCode)
        
        If blnCard And Len(vsfPerson.EditText) = ParamInfo.就诊卡号码长度 - 1 And KeyCode <> 8 And KeyCode <> vbKeyReturn Then
            vsfPerson.Body.EditSelStart = Len(vsfPerson.EditText)
            strInput = strInput & " AND C.就诊卡号=[1] "
        End If
        
        If KeyCode = vbKeyReturn Then
            If blnCard Then
                '是就诊卡
                strInput = strInput & " AND C.就诊卡号=[1] "
            Else
                '非就诊卡
                blnCard = False

                strText = vsfPerson.EditText
    
                Select Case UCase(Left(strText, 1))
                Case "-", "A"                 '病人id,就诊卡号
                    strInput = strInput & " AND C.病人id=[1]"
                Case "+", "B"                 '住院号
                    strInput = " AND C.住院号=[1]"
                Case "*", "D"                 '门诊号
                    strInput = strInput & " AND C.门诊号=[1]"
                Case "/", "C"                 '当前床号
                    strInput = strInput & " AND C.当前床号=[1]"
                Case Else
                    strSvrText = vsfPerson.Cell(flexcpData, Row, Col)
                    vsfPerson.Cell(flexcpData, Row, Col) = strText
                End Select
            End If
        End If
            
        If strInput <> "" Then
            gstrSQL = GetPublicSQL(SQL.人员过滤选择, strInput)
            
            If blnCard Then
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, UCase(strText))
            ElseIf UCase(Left(strText, 1)) = "/" Or UCase(Left(strText, 1)) = "C" Then
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Trim(Mid(strText, 2)))
            Else
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Mid(strText, 2)))
            End If
            
            If ShowGrdFilter(Me, vsfPerson, "姓名,1200,0,0;性别,810,0,0;出生日期,1200,0,0;婚姻状况,900,0,0;身份证号,1500,0,0", Me.Name & "\人员过滤选择Grid", "请从下面选择一个人员", rsData, rs, , , , False) Then
                                                                        
                vsfPerson.EditText = zlCommFun.NVL(rs("姓名"))
                strText = vsfPerson.EditText
                If mlngGroup <> Val(zlCommFun.NVL(rs("合同单位id"))) And Val(zlCommFun.NVL(rs("合同单位id"))) > 0 And mlngGroup > 0 Then

                    If MsgBox("不是当前团体的人员，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        KeyCode = 0
                        vsfPerson.EditText = ""
                        vsfPerson.TextMatrix(Row, Col) = strSvrText
                        Cancel = True
                        Exit Sub
                    End If
                End If
    
                If CheckHavePerson(Val(zlCommFun.NVL(rs("ID")))) Then
                    ShowSimpleMsg "病人“" & zlCommFun.NVL(rs("姓名").Value) & "”已经存在！"
                    KeyCode = 0
                    vsfPerson.EditText = ""
                    vsfPerson.TextMatrix(Row, Col) = strSvrText
                    Cancel = True
                    Exit Sub
                End If
    
                vsfPerson.Cell(flexcpData, Row, vsfPerson.Col) = zlCommFun.NVL(rs("姓名").Value)
                vsfPerson.TextMatrix(Row, mPersonCol.姓名) = zlCommFun.NVL(rs("姓名"))
                Call SetDefault(Row)
                vsfPerson.TextMatrix(Row, mPersonCol.身份证) = zlCommFun.NVL(rs("身份证号"))
                vsfPerson.TextMatrix(Row, mPersonCol.出生日期) = Format(zlCommFun.NVL(rs("出生日期")), "yyyy-MM-dd")
                vsfPerson.TextMatrix(Row, mPersonCol.性别) = zlCommFun.NVL(rs("性别").Value)
                vsfPerson.TextMatrix(Row, mPersonCol.婚姻状况) = zlCommFun.NVL(rs("婚姻状况").Value)
                vsfPerson.TextMatrix(Row, mPersonCol.病人id) = zlCommFun.NVL(rs("ID"))
                vsfPerson.TextMatrix(Row, mPersonCol.年龄) = zlCommFun.NVL(rs("年龄"))
                vsfPerson.TextMatrix(Row, mPersonCol.门诊号) = zlCommFun.NVL(rs("门诊号"))
                vsfPerson.TextMatrix(Row, mPersonCol.健康号) = zlCommFun.NVL(rs("健康号"))
                                
                vsfPerson.Cell(flexcpForeColor, Row, 0, Row, vsfPerson.Cols - 1) = COLOR.黑色
                
                vsfPerson.EditMode(mPersonCol.门诊号) = 0
    
                If blnCard Then
                    vsfPerson.Cell(flexcpData, Row, Col) = strText
                    vsfPerson.TextMatrix(Row, Col) = vsfPerson.Cell(flexcpData, Row, Col)
                    KeyCode = 13
                End If
                
                DataChange = True
            Else
                '取消了本次选择
                    
                vsfPerson.EditMode(mPersonCol.门诊号) = 1
                vsfPerson.Cell(flexcpForeColor, Row, 0, Row, vsfPerson.Cols - 1) = COLOR.兰色
                
                vsfPerson.Cell(flexcpData, Row, Col) = vsfPerson.EditText
                vsfPerson.EditText = vsfPerson.Cell(flexcpData, Row, Col)
                vsfPerson.TextMatrix(Row, Col) = vsfPerson.Cell(flexcpData, Row, Col)
                vsfPerson.TextMatrix(Row, mPersonCol.门诊号) = ""
                vsfPerson.TextMatrix(Row, mPersonCol.身份证) = ""
                vsfPerson.TextMatrix(Row, mPersonCol.病人id) = ""
                vsfPerson.TextMatrix(Row, mPersonCol.出生日期) = ""
                vsfPerson.TextMatrix(Row, mPersonCol.年龄) = ""
                Call SetDefault(Row)
                
            End If
        ElseIf KeyCode = vbKeyReturn Then
            '新病人，允许输入门诊号
            
            vsfPerson.EditMode(mPersonCol.门诊号) = 1
            vsfPerson.Cell(flexcpForeColor, Row, 0, Row, vsfPerson.Cols - 1) = COLOR.兰色
            vsfPerson.TextMatrix(Row, mPersonCol.病人id) = ""
            vsfPerson.TextMatrix(Row, mPersonCol.门诊号) = ""
            vsfPerson.TextMatrix(Row, mPersonCol.身份证) = ""
            vsfPerson.TextMatrix(Row, mPersonCol.出生日期) = ""
            vsfPerson.TextMatrix(Row, mPersonCol.年龄) = ""
            
            Call SetDefault(Row)
        End If
    End If
End Sub

Private Sub vsfPerson_KeyPress(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer, Cancel As Boolean)

    If KeyAscii = vbKeyReturn Then

        If Col = 1 Then
            If Trim(vsfPerson.TextMatrix(Row, Col)) = "" Then
                
                KeyAscii = 0
                
                cmdOK.SetFocus
                Cancel = True

            End If
        End If
    End If

End Sub

Private Sub vsfPerson_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    
    Select Case Col
    Case mPersonCol.门诊号
        '检查门诊号是否存在
        If Val(vsfPerson.EditText) > 0 Then
            gstrSQL = "Select 1 From 病人信息 Where 门诊号=[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(vsfPerson.EditText))
            If rs.BOF = False Then
                '存在
                Cancel = True
                
                vsfPerson.TextMatrix(Row, Col) = vsfPerson.EditText
                
                ShowSimpleMsg "当前门诊号：" & Val(vsfPerson.EditText) & "已经存在，不允许重复！"
                vsfPerson.EditText = ""
                vsfPerson.TextMatrix(Row, Col) = ""
                
            End If
        End If
    End Select
End Sub


