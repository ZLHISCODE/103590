VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmNurseFileEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "文件编辑"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4545
   Icon            =   "frmNurseFileEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCanCel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2640
      TabIndex        =   8
      Top             =   1920
      Width           =   1155
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   1410
      TabIndex        =   7
      Top             =   1920
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -120
      TabIndex        =   6
      Top             =   1710
      Width           =   4545
   End
   Begin MSMask.MaskEdBox mskEdit 
      Height          =   285
      Left            =   1230
      TabIndex        =   5
      Top             =   1110
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   19
      Format          =   "yyyy-MM-dd HH:mm:ss"
      Mask            =   "####-##-## ##:##:##"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txt文件名称 
      Height          =   285
      Left            =   1230
      MaxLength       =   50
      TabIndex        =   3
      Top             =   720
      Width           =   2895
   End
   Begin VB.ComboBox cbo格式来源 
      Height          =   300
      Left            =   1230
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   330
      Width           =   2895
   End
   Begin VB.Label lbl开始时间 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "开始时间"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   450
      TabIndex        =   4
      Top             =   1155
      Width           =   720
   End
   Begin VB.Label lbl文件名称 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "文件名称"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   450
      TabIndex        =   2
      Top             =   765
      Width           =   720
   End
   Begin VB.Label lbl格式来源 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "格式来源"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   450
      TabIndex        =   0
      Top             =   390
      Width           =   720
   End
End
Attribute VB_Name = "frmNurseFileEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintShowTime As Integer '体温单文件的缺省开始时间:1-入科时间;0-入院时间
Private mstr入院时间 As String
Private mstr入科时间 As String
Private mstr出院时间 As String

Private mlngFile As Long        '文件ID,传入0表示新增,否则表示修改(修改时不允许修改文件来源)
Private mlngFormat As Long
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mint婴儿 As Long
Private mlng科室ID As Long
Private mstrDept As String      '当前科室
Private mblnOK As Boolean       '是否保存成功
Private mblnExist体温单 As Boolean
Private mblnExist记录单 As Boolean
Private mblnExist产程图 As Boolean
Private mIntPartogramID As Boolean
Private mstrCurForamt As String '病人现有体温单格式ID组合(按文件开始时间排序),格式:30,40
Private mblnOnly As Boolean     '住院病人同一时点只记录一份护理文件

Public Function ShowEditor(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal int婴儿 As Integer, ByVal lng科室ID As Long, _
    ByVal str科室 As String, Optional lngFile As Long = 0, Optional lngFormat As Long = 0) As Boolean
    mblnOK = False
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    mint婴儿 = int婴儿
    mlng科室ID = lng科室ID
    mlngFile = lngFile
    mlngFormat = lngFormat
    mstrDept = str科室
    mIntPartogramID = -1
    Me.Show 1
    lngFile = mlngFile
    lngFormat = mlngFormat            '返回格式ID,用于定位到相同格式的文件
    ShowEditor = mblnOK
End Function

Private Sub cbo格式来源_Click()
    Dim bln体温单 As Boolean
    Dim bln产程图 As Boolean
    Dim strDate As String
    
    txt文件名称.Text = Split(cbo格式来源.Text, "-")(1)
    If InStr(1, "," & cbo格式来源.Tag & ",", "," & cbo格式来源.ListIndex & ",") = 0 Or cbo格式来源.Tag = "" Then
        If mIntPartogramID <> Me.cbo格式来源.ItemData(Me.cbo格式来源.ListIndex) Then
            txt文件名称.Text = "[" & mstrDept & "]" & txt文件名称.Text '记录单时如下处理
        Else
            '产程图
            bln产程图 = True
        End If
    Else
        '体温单:目前允许添加多份体温单
        txt文件名称.Text = "[" & mstrDept & "]" & txt文件名称.Text
    End If
    
    '新增:如果不存护理文件则缺省时间为入院时间,否则为当前时间
    '修改:护理文件的开始时间不能小于入院时间,不能大于数据发生时间,如无数据则不能大于当前时间
    bln体温单 = (cbo格式来源.Tag <> "" And InStr(1, "," & cbo格式来源.Tag & ",", "," & cbo格式来源.ListIndex & ",") > 0)
    If mlngFile = 0 Then
        If (Not mblnExist记录单 And Not bln体温单) Or (Not mblnExist体温单 And bln体温单) Or (Not mblnExist产程图 And bln产程图) Then
            '如果已入科则显示入科时间,否则显示入院时间
            mskEdit.Text = Format(IIf(mstr入科时间 = "", mstr入院时间, mstr入科时间), "YYYY-MM-DD HH:mm:ss")
        Else
            If bln体温单 Then
                strDate = GetCreateWaveDate
            Else
                strDate = Format(Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm") & ":00", "YYYY-MM-DD HH:mm:ss")
            End If
            If mstr出院时间 <> "" And Format(strDate, "YYYY-MM-DD HH:mm:ss") > Format(mstr出院时间, "YYYY-MM-DD HH:mm:ss") Then
                strDate = Format(mstr出院时间, "YYYY-MM-DD") & " 00:00:00"
            End If
            mskEdit.Text = strDate
        End If
        mskEdit.Tag = mskEdit.Text
    End If
    '如果选中的是第一份体温单的话
    If IsFirstCurve Then
        mskEdit.Text = Format(IIf(mintShowTime = 1 And mstr入科时间 <> "", mstr入科时间, mstr入院时间), "YYYY-MM-DD HH:mm:ss")
        '56627:放开婴儿体温单不能修改的限制，只要时间不小于婴儿出生时间即可。
        'mskEdit.Enabled = mint婴儿 = 0
    Else
        mskEdit.Enabled = True
    End If
End Sub

Private Function GetCreateWaveDate() As String
'-----------------------------------------------------------
'功能：新建体温单时获取文件开始时间
'如果转入科室时间要大于病人之前体温单文件的最大数据发生时间或开始时间，
'则病人转科后新建体温单时间为转入科室时间,否则为当前系统时间
'返回：创建体温单的时间
'-----------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strDate As String
    
    On Error GoTo ErrHand
    gstrSQL = _
        " SELECT E.开始时间" & vbNewLine & _
        " FROM 病人变动记录 e" & vbNewLine & _
        " WHERE e.病人id = [1] AND e.主页id = [2] AND Nvl(e.附加床位, 0) = 0 AND e.开始时间 IS NOT NULL AND e.开始原因 IN (3, 15) AND" & vbNewLine & _
        "      e.终止时间 IS NULL AND e.开始时间 > (SELECT Nvl(MAX(b.发生时间), MAX(a.开始时间))" & vbNewLine & _
        "                                       FROM 病历文件列表 c, 病人护理文件 a, 病人护理数据 b" & vbNewLine & _
        "                                       WHERE a.病人id = [1] AND a.主页id = [2] AND Nvl(婴儿, 0) = [3] AND a.Id = b.文件id(+) AND" & vbNewLine & _
        "                                             c.Id = a.格式id AND c.种类 = 3 AND c.保留 = -1)"
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取创建文件时间", mlng病人ID, mlng主页ID, mint婴儿)
    If rsTemp.RecordCount > 0 Then
        strDate = Format(rsTemp!开始时间, "YYYY-MM-DD HH:mm:ss")
    Else
        strDate = Format(Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm") & ":00", "YYYY-MM-DD HH:mm:ss")
    End If
    GetCreateWaveDate = strDate
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function IsFirstCurve() As Boolean
'功能:检查目前选中的体温单是否是第一份体温单
    Dim arrCode() As String
    Dim blnIsSelectCurve As Boolean
    
    blnIsSelectCurve = (cbo格式来源.Tag <> "" And InStr(1, "," & cbo格式来源.Tag & ",", "," & cbo格式来源.ListIndex & ",") > 0)
    If blnIsSelectCurve = False Then IsFirstCurve = False: Exit Function
    If mblnExist体温单 = False Then IsFirstCurve = True: Exit Function
    arrCode = Split(mstrCurForamt, ",")
    If Val(arrCode(0)) = mlngFile And mlngFile <> 0 Then
        IsFirstCurve = True
    Else
        IsFirstCurve = False
    End If
End Function

Private Sub cmdCanCel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim lngID As Long
    Dim strDate As String, strTime As String
    Dim lngUpFileID As Long
    Dim strCurDate As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHand
    If txt文件名称.Text = "" Then
        MsgBox "请输入文件名称！", vbInformation, gstrSysName
        If txt文件名称.Enabled And txt文件名称.Visible Then txt文件名称.SetFocus
        Exit Sub
    End If
    If LenB(StrConv(txt文件名称.Text, vbFromUnicode)) > 50 Then
        MsgBox "文件名称超长！（最多50个字符或25个汉字）", vbInformation, gstrSysName
        If txt文件名称.Enabled And txt文件名称.Visible Then txt文件名称.SetFocus
        Exit Sub
    End If
    
    strCurDate = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:mm:ss")
    strDate = Format(mskEdit.Text, "YYYY-MM-DD HH:mm:ss")
    If Not IsDate(strDate) Then
        MsgBox "文件开始时间格式不对（如：2011-4-13 23:59:00），请重新输入！", vbInformation, gstrSysName
        If IsDate(Format(mskEdit.Tag, "YYYY-MM-DD HH:mm:ss")) Then mskEdit.Text = mskEdit.Tag
        If mskEdit.Enabled And mskEdit.Visible Then mskEdit.SetFocus
        Exit Sub
    End If
    '56627
    '建立文件日期不能大于当前日期或出院时间
    If mstr出院时间 = "" Then
        If strDate > strCurDate Then
            MsgBox "文件开始时间不能大于当前时间[" & strCurDate & "]，请重新输入！", vbInformation, gstrSysName
            If IsDate(Format(mskEdit.Tag, "YYYY-MM-DD HH:mm:ss")) Then mskEdit.Text = mskEdit.Tag
            If mskEdit.Enabled And mskEdit.Visible Then mskEdit.SetFocus
            Exit Sub
        End If
    Else
        If strDate > Format(mstr出院时间, "YYYY-MM-DD HH:mm:ss") Then
            MsgBox "文件开始时间不能大于出院时间[" & mstr出院时间 & "]，请重新输入！", vbInformation, gstrSysName
            If IsDate(Format(mskEdit.Tag, "YYYY-MM-DD HH:mm:ss")) Then mskEdit.Text = mskEdit.Tag
            If mskEdit.Enabled And mskEdit.Visible Then mskEdit.SetFocus
            Exit Sub
        End If
    End If
    '56627:对于婴儿第一份体温单，体温单开始时间不能小于婴儿出生时间
    If IsFirstCurve And mint婴儿 <> 0 Then
        If Format(strDate, "YYYY-MM-DD HH:mm:ss") < Format(mstr入科时间, "YYYY-MM-DD HH:mm:ss") Then
            MsgBox "婴儿文件的开始时间不能小于出生时间[" & Format(mstr入科时间, "YYYY-MM-DD HH:mm:ss") & "],请重新输入！", vbInformation, gstrSysName
            If IsDate(Format(mskEdit.Tag, "YYYY-MM-DD HH:mm:ss")) Then mskEdit.Text = mskEdit.Tag
            If mskEdit.Enabled And mskEdit.Visible Then mskEdit.SetFocus
            Exit Sub
        End If
    End If
    '对于体温单
    If cbo格式来源.Tag <> "" And InStr(1, "," & cbo格式来源.Tag & ",", "," & cbo格式来源.ListIndex & ",") > 0 Then
        gstrSQL = _
            " SELECT A.ID,MAX(A.文件名称) 文件名称,MAX(A.开始时间) 开始时间,NVL(MAX(C.发生时间),MAX(A.开始时间)) 发生时间" & vbNewLine & _
            " FROM 病人护理文件 A, 病历文件列表 B, 病人护理数据 C" & vbNewLine & _
            " WHERE A.格式ID = B.ID AND A.ID = C.文件ID(+) AND A.病人ID = [1] AND A.主页ID = [2] AND A.婴儿 = [3] AND B.保留 = -1" & vbNewLine & _
            " GROUP BY A.ID"
        Call SQLDIY(gstrSQL)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "判断是否已定义体温单", mlng病人ID, mlng主页ID, mint婴儿)
        '如果是新建文件(文件的开始时间一定要大于上一文件开始时间或数据发生时间)
        '如果是修改文件,修改文件的开始时间要大于上一文件文件开始时间或数据发生时间,要小于下一文件的开始时间
        If mlngFile = 0 Or (mlngFile <> 0 And Not IsDate(Format(mskEdit.Tag, "YYYY-MM-DD HH:mm:ss"))) Then
            rsTemp.Filter = ""
            strTime = strDate
        Else
            strTime = Format(mskEdit.Tag, "YYYY-MM-DD HH:mm:ss")
            rsTemp.Filter = "开始时间< '" & strTime & "'"
        End If
        rsTemp.Sort = "开始时间 DESC"
        If rsTemp.RecordCount > 0 Then
            lngUpFileID = rsTemp!ID
            If CDate(strDate) <= CDate(rsTemp!发生时间) Then
                MsgBox "文件开始时间要大于上一文件【" & NVL(rsTemp!文件名称) & "】的开始或数据发生时间【" & Format(rsTemp!发生时间, "YYYY-MM-DD HH:mm:ss") & "】，请重新输入！", vbInformation, gstrSysName
                If IsDate(Format(mskEdit.Tag, "YYYY-MM-DD HH:mm:ss")) Then mskEdit.Text = mskEdit.Tag
                If mskEdit.Enabled And mskEdit.Visible Then mskEdit.SetFocus
                Exit Sub
            End If
        End If
        strTime = Format(mskEdit.Tag, "YYYY-MM-DD HH:mm:ss")
        If mlngFile <> 0 And IsDate(strTime) Then
            rsTemp.Filter = "开始时间> '" & strTime & "'"
            rsTemp.Sort = "开始时间 ASC"
            If rsTemp.RecordCount > 0 Then
                If CDate(strDate) >= CDate(rsTemp!开始时间) Then
                    MsgBox "文件开始时间要小于下一文件【" & NVL(rsTemp!文件名称) & "】的开始时间【" & Format(rsTemp!开始时间, "YYYY-MM-DD HH:mm:ss") & "】，请重新输入！", vbInformation, gstrSysName
                    If IsDate(Format(mskEdit.Tag, "YYYY-MM-DD HH:mm:ss")) Then mskEdit.Text = mskEdit.Tag
                    If mskEdit.Enabled And mskEdit.Visible Then mskEdit.SetFocus
                    Exit Sub
                End If
            End If
        End If
    End If
    
    If mlngFile = 0 Then
        lngID = zlDatabase.GetNextId("病人护理文件")
    Else
        lngID = mlngFile
    End If
    
    gstrSQL = "ZL_病人护理文件_UPDATE(" & lngID & "," & mlng科室ID & "," & mlng病人ID & "," & mlng主页ID & "," & mint婴儿 & "," & _
              cbo格式来源.ItemData(cbo格式来源.ListIndex) & ",'" & txt文件名称.Text & "',to_date('" & strDate & "','yyyy-MM-dd hh24:mi:ss')," & IIf(mlngFile = 0 And mblnOnly, 1, 0) & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "更新数据")
    mlngFile = lngID
    mlngFormat = mlngFile 'cbo格式来源.ItemData(cbo格式来源.ListIndex)
    
    '如果当前创建的是记录单且不存在体温单则自动创建一份体温单
    If Not mblnExist体温单 And Not (cbo格式来源.Tag <> "" And InStr(1, "," & cbo格式来源.Tag & ",", "," & cbo格式来源.ListIndex & ",") > 0) Then
        lngID = zlDatabase.GetNextId("病人护理文件")
        gstrSQL = "ZL_病人护理文件_UPDATE(" & lngID & "," & mlng科室ID & "," & mlng病人ID & "," & mlng主页ID & "," & mint婴儿 & "," & _
                  cbo格式来源.ItemData(Val(cbo格式来源.Tag)) & ",'病人体温表',to_date('" & IIf(mintShowTime = 1 And mstr入科时间 <> "", mstr入科时间, mstr入院时间) & "','yyyy-MM-dd hh24:mi:ss'),0)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "更新数据")
    ElseIf mblnExist体温单 And (cbo格式来源.Tag <> "" And InStr(1, "," & cbo格式来源.Tag & ",", "," & cbo格式来源.ListIndex & ",") > 0) Then
        '将上一体温单文件的结束时间更新为本文件的开始时间-1S
        If lngUpFileID > 0 Then
            strDate = CDate(strDate) - (1 / 24 / 60 / 60)
            gstrSQL = "ZL_病人护理文件_STATE(" & lngUpFileID & ",1,To_Date('" & strDate & "','yyyy-MM-dd hh24:mi:ss'))"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "标记文件结束")
        End If
    End If
    
    mblnOK = True
    Unload Me
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim blnSeek As Boolean
    Dim lng格式 As Long, lng保留 As Long, str文件名称 As String, str开始时间 As String, str格式 As String
    Dim rsTemp As New ADODB.Recordset
    Dim intIndex As Integer

    On Error GoTo ErrHand
    
    mskEdit.Enabled = True
    mintShowTime = zlDatabase.GetPara("体温单文件开始时间", glngSys, 1255, 1)
    
    '提取当前病人入院时间
    gstrSQL = " Select A.入院日期,A.出院日期,B.名称 From 病案主页 A,部门表 B" & _
        " Where A.病人ID=[1] And A.主页ID=[2] And A.出院科室ID=B.ID(+)"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取当前病人入院时间", mlng病人ID, mlng主页ID)
    mstr入院时间 = Format(rsTemp!入院日期, "yyyy-MM-dd HH:mm:ss")
    mstr出院时间 = Format(NVL(rsTemp!出院日期), "yyyy-MM-dd HH:mm:ss")
    If mstrDept = "" Then mstrDept = NVL(rsTemp!名称)
    
    mstr入科时间 = ""
    gstrSQL = " Select 开始时间 From 病人变动记录 Where 病人ID=[1] And 主页ID=[2] And 开始原因=2 Order by 开始时间"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取当前病人入科时间", mlng病人ID, mlng主页ID)
    If rsTemp.RecordCount <> 0 Then
        mstr入科时间 = Format(rsTemp!开始时间, "yyyy-MM-dd HH:mm:ss")
    End If
    
    '如果是婴儿这提取婴儿登记时间为文件开始时间
    If mint婴儿 <> 0 Then
        gstrSQL = "select 出生时间 from 病人新生儿记录 where 病人ID=[1] And 主页ID=[2] And 序号=[3] "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取婴儿登记时间", mlng病人ID, mlng主页ID, mint婴儿)
        If rsTemp.RecordCount <> 0 Then
            mstr入科时间 = Format(NVL(rsTemp!出生时间, mstr入科时间), "yyyy-MM-dd HH:mm:ss")
            mstr入院时间 = mstr入科时间
        End If
        '提取婴儿出院日期
        gstrSQL = "Select b.病人id, b.主页id, b.婴儿, 开始执行时间" & vbNewLine & _
            " From 病人医嘱记录 b, 诊疗项目目录 c" & vbNewLine & _
            " Where b.诊疗项目id + 0 = c.Id And b.医嘱状态 = 8 And Nvl(b.婴儿, 0) <> 0 And c.类别 = 'Z' And c.操作类型 In ('3', '5', '11') And" & vbNewLine & _
            "      b.病人id = [1] And b.主页id = [2] And b.婴儿 = [3]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取婴儿登记时间", mlng病人ID, mlng主页ID, mint婴儿)
        If rsTemp.RecordCount <> 0 Then
            mstr出院时间 = Format(NVL(rsTemp!开始执行时间, mstr出院时间), "yyyy-MM-dd HH:mm:ss")
        End If
    End If
    '检查是否已设定体温单,如已存在则不允许再次添加体温单
    gstrSQL = " Select B.保留,编号,A.ID,A.开始时间" & _
              " From 病人护理文件 A,病历文件列表 B" & _
              " Where A.格式ID=B.ID And A.病人ID=[1] And A.主页ID=[2] And A.婴儿=[3] Order by B.保留,A.开始时间"
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "判断是否已定义体温单", mlng病人ID, mlng主页ID, mint婴儿)
    rsTemp.Filter = "保留=-1"
    rsTemp.Sort = "保留,开始时间"
    mstrCurForamt = ""
    mblnExist体温单 = rsTemp.RecordCount
    With rsTemp
        Do While Not .EOF
            mstrCurForamt = IIf(mstrCurForamt = "", "", mstrCurForamt & ",") & Val(!ID)
            .MoveNext
        Loop
    End With
    
    rsTemp.Filter = "保留=1"
    mblnExist产程图 = rsTemp.RecordCount
    rsTemp.Filter = "保留<>-1"
    mblnExist记录单 = rsTemp.RecordCount
    rsTemp.Filter = 0
    
    If mint婴儿 <> 0 Then mblnExist产程图 = True
    
    '读取文件设置
    gstrSQL = "SELECT A.科室ID, B.名称 AS 科室, A.格式ID, A.文件名称, A.开始时间, C.保留,C.编号 || '-' || C.名称 As 格式" & vbNewLine & _
            "  FROM 病人护理文件 A, 病历文件列表 C, 部门表 B" & vbNewLine & _
            "  WHERE A.格式ID = C.ID AND A.科室ID = B.ID AND A.ID = [1]"
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取文件设置", mlngFile)
    If rsTemp.RecordCount <> 0 Then
        mlng科室ID = rsTemp!科室ID
        mstrDept = rsTemp!科室
        lng格式 = rsTemp!格式ID
        str文件名称 = rsTemp!文件名称
        str开始时间 = Format(rsTemp!开始时间, "yyyy-MM-dd HH:mm:ss")
        lng保留 = Val(NVL(rsTemp!保留, 0))
        str格式 = NVL(rsTemp!格式, "-")
    End If
    
    '读取护理相关的病历文件并加载
'    gstrSQL = " Select ID,保留,编号,编号||'-'||名称 AS 格式 From 病历文件列表 " & _
'              " Where 种类=3 And (通用 =1 OR (通用=2 And ID IN " & _
'              "     (Select 文件ID FROM 病历应用科室 Where 科室ID = [1]))) " & _
'              " Order by 保留,编号"
    gstrSQL = "Select ID, 保留, 编号, 格式" & vbNewLine & _
        "From (Select ID, 保留, 编号, 编号 || '-' || 名称 As 格式" & vbNewLine & _
        "       From 病历文件列表" & vbNewLine & _
        "       Where 种类 = 3 And 保留 <> 1 And (通用 = 1 Or (通用 = 2 And ID In (Select 文件id From 病历应用科室 Where 科室id = [1])))" & vbNewLine & _
        "       Union All" & vbNewLine & _
        "       Select ID, 保留, 编号, 编号 || '-' || 名称 As 格式" & vbNewLine & _
        "       From 病历文件列表" & vbNewLine & _
        "       Where 种类 = 3 And 保留 = 1 And" & vbNewLine & _
        "             1 = (Select 1" & vbNewLine & _
        "                  From 部门性质说明 A, 部门表 B" & vbNewLine & _
        "                  Where a.工作性质 = '产科' And a.部门id = b.Id And b.Id = [1] And Rownum < 2))" & vbNewLine & _
        "Order By 保留, 编号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取护理相关的病历文件", glng病区ID)
    With rsTemp
        Me.cbo格式来源.Clear
        Me.cbo格式来源.Tag = ""
        intIndex = 0
        Do While Not .EOF
            'If (((!保留 = -1 And NOT mblnExist体温单) Or (!保留 = 1 And Not mblnExist产程图) Or InStr(1, ",-1,1,", "," & !保留 & ",") = 0) And mlngFile = 0) Or mlngFile <> 0 Then
            If (((!保留 = 1 And Not mblnExist产程图) Or !保留 <> 1) And mlngFile = 0) Or mlngFile <> 0 Then
                Me.cbo格式来源.AddItem !格式
                Me.cbo格式来源.ItemData(Me.cbo格式来源.NewIndex) = !ID
                If !保留 = -1 Then Me.cbo格式来源.Tag = IIf(Me.cbo格式来源.Tag = "", "", Me.cbo格式来源.Tag & ",") & intIndex
                If !保留 = 1 Then mIntPartogramID = !ID
                If !ID = lng格式 Then
                    Me.cbo格式来源.ListIndex = intIndex
                    blnSeek = True
                End If
                intIndex = intIndex + 1
            End If
            .MoveNext
        Loop
        If .RecordCount <> 0 And Not blnSeek Then Me.cbo格式来源.ListIndex = 0
    End With
    
    If mlngFile <> 0 Then
        Me.txt文件名称.Text = str文件名称
        '病人换病区(A->B)修改文件时，如果此文件适用于A病区，上面语句是无法提取到此文件的信息，为了保证文件正确性，需要进行特殊处理
        If Not (Me.cbo格式来源.ItemData(Me.cbo格式来源.ListIndex) = lng格式) Then
            Me.cbo格式来源.AddItem str格式
            Me.cbo格式来源.ItemData(Me.cbo格式来源.NewIndex) = lng格式
            If lng保留 = -1 Then Me.cbo格式来源.Tag = IIf(Me.cbo格式来源.Tag = "", "", Me.cbo格式来源.Tag & ",") & Me.cbo格式来源.NewIndex
            If lng保留 = 1 Then mIntPartogramID = lng格式
            Me.cbo格式来源.ListIndex = Me.cbo格式来源.NewIndex
        End If
        If str开始时间 <> "" Then Me.mskEdit.Text = Format(str开始时间, "YYYY-MM-DD HH:mm:ss"): mskEdit.Tag = mskEdit.Text
    Else
        mblnOnly = (Val(zlDatabase.GetPara("对应多份护理文件", glngSys, 1255, 0)) = 0)
    End If
    Me.cbo格式来源.Enabled = (mlngFile = 0)
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub mskEdit_GotFocus()
    Call zlControl.TxtSelAll(mskEdit)
End Sub

Private Sub txt文件名称_GotFocus()
    Call zlControl.TxtSelAll(txt文件名称)
End Sub
