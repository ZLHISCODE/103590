VERSION 5.00
Begin VB.Form frm设置清算方式 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "设置清算方式"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6165
   Icon            =   "frm设置清算方式.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txt大额补助分担比例 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1980
      TabIndex        =   12
      Top             =   2880
      Width           =   3885
   End
   Begin VB.TextBox txt大额补助清算标准 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1980
      TabIndex        =   10
      Top             =   2490
      Width           =   3885
   End
   Begin VB.TextBox txt基本统筹分担比例 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1980
      TabIndex        =   8
      Top             =   2100
      Width           =   3885
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   30
      TabIndex        =   16
      Top             =   3390
      Width           =   6165
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "下载(&W)"
      Height          =   350
      Left            =   180
      TabIndex        =   15
      Top             =   3570
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3390
      TabIndex        =   13
      Top             =   3570
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4650
      TabIndex        =   14
      Top             =   3570
      Width           =   1100
   End
   Begin VB.TextBox txt基本统筹清算标准 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1980
      TabIndex        =   6
      Top             =   1710
      Width           =   3885
   End
   Begin VB.TextBox txt疾病信息 
      Height          =   300
      Left            =   1980
      TabIndex        =   3
      Top             =   1320
      Width           =   3615
   End
   Begin VB.CommandButton cmd疾病信息 
      Caption         =   "…"
      Height          =   300
      Left            =   5580
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1320
      Width           =   285
   End
   Begin VB.Label lbl大额补助分担比例 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "大额补助分担比例"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   450
      TabIndex        =   11
      Top             =   2940
      Width           =   1440
   End
   Begin VB.Label lbl大额补助清算标准 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "大额补助清算标准"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   450
      TabIndex        =   9
      Top             =   2550
      Width           =   1440
   End
   Begin VB.Label lbl基本统筹分担比例 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "基本统筹分担比例"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   450
      TabIndex        =   7
      Top             =   2160
      Width           =   1440
   End
   Begin VB.Label lblNote 
      Caption         =   "    请选择一个单病种，本次住院将按该病种对应的清算方式对费用进行结算"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   405
      Index           =   1
      Left            =   1260
      TabIndex        =   1
      Top             =   750
      Width           =   4845
   End
   Begin VB.Label lbl基本统筹清算标准 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "基本统筹清算标准"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   450
      TabIndex        =   5
      Top             =   1770
      Width           =   1440
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   360
      Picture         =   "frm设置清算方式.frx":000C
      Top             =   360
      Width           =   480
   End
   Begin VB.Label lblNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    如果是第一次使用或医院的单病种数据发生变化，请使用下载功能，将单病种清算数据下载到本地。"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   555
      Index           =   0
      Left            =   1260
      TabIndex        =   0
      Top             =   150
      Width           =   4845
   End
   Begin VB.Label lbl疾病信息 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "单病种(&J)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1080
      TabIndex        =   2
      Top             =   1380
      Width           =   810
   End
End
Attribute VB_Name = "frm设置清算方式"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const madLongVarCharDefault As Integer = 10          '字符型字段缺省长度
Private Const madDoubleDefault As Integer = 18               '数字型字段缺省长度
Private Const madDbDateDefault As Integer = 20               '日期型字段缺省长度

Private mblnOK As Boolean
Private mint险类 As Integer
Private mint保险类别 As Integer
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mstr卡号 As String
Private mstr医保号 As String
Private mstr分中心编号 As String
Private mstr密码 As String
Private mbln居民 As Boolean
Private mrs病种 As New ADODB.Recordset

Private Sub Record_Add(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String)
    Dim arrFields, arrValues, intField As Integer
    '添加记录
    'strFields:字段名|字段名
    'strValues:值|值
    
    '例子：
    'Dim strFields As String, strValues As String
    'strFields = "RecordID|科目ID|摘要"
    'strValues = "5188|6666|科目名称"
    'Call Record_Update(rsVoucher, strFields, strValues)

    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField = 0 Then Exit Sub

    With rsObj
        .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Private Sub Record_Init(ByRef rsObj As ADODB.Recordset, ByVal strFields As String)
    Dim arrFields, intField As Integer
    Dim strFieldName As String, intType As Integer, lngLength As Long
    '初始化映射记录集
    'strFields:字段名,类型,长度|字段名,类型,长度    如果长度为零,则取默认长度
    '字符型:adLongVarChar;数字型:adDouble;日期型:adDBDate
    
    '例子：
    'Dim rsVoucher As New ADODB.Recordset, strFields As String
    'strFields = "RecordID," & adDouble & ",18|科目ID," & adDouble & ",18|摘要, " & adLongVarChar & ",50|" & _
    '"删除," & adDouble & ",1"
    'Call Record_Init(rsVoucher, strFields)

    arrFields = Split(strFields, "|")
    Set rsObj = New ADODB.Recordset

    With rsObj
        If .State = 1 Then .Close
        For intField = 0 To UBound(arrFields)
            strFieldName = Split(arrFields(intField), ",")(0)
            intType = Split(arrFields(intField), ",")(1)
            lngLength = Split(arrFields(intField), ",")(2)

            '获取字段缺省长度
            If lngLength = 0 Then
                Select Case intType
                Case adDouble
                    lngLength = madDoubleDefault
                Case adVarChar
                    lngLength = madLongVarCharDefault
                Case adLongVarChar
                    lngLength = madLongVarCharDefault
                Case Else
                    lngLength = madDbDateDefault
                End Select
            End If
            .Fields.Append strFieldName, intType, lngLength, adFldIsNullable
        Next
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDown_Click()
    If InitXML = False Then Exit Sub
    Call InsertChild(mdomInput.documentElement, "CENTERCODE", mstr医保中心编码_贵阳)
    If Not CommServer("GETHOSPSINGLEILLNESS") Then Exit Sub
    MsgBox "下载成功！", vbInformation, gstrSysName
End Sub

Private Sub cmdOK_Click()
    Dim rsTemp As New ADODB.Recordset
    If txt疾病信息.Tag = "" And txt疾病信息 <> "()控制线清算" Then
        MsgBox "请选择一个单病种！", vbInformation, gstrSysName
        txt疾病信息.SetFocus
        Exit Sub
    End If
    
    On Error Resume Next
    If mlng主页ID <> 0 Then
        gstrSQL = " Select NVL(特殊结算方式,'00') AS 特殊结算方式 From 医保病人住院信息 Where 病人ID=[1] And 主页ID=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "判断是否允许修改清算方式", mlng病人ID, mlng主页ID)
        If rsTemp.RecordCount <> 0 Then
            If Err = 0 Then
                If Mid(rsTemp!特殊结算方式, 2, 1) <> "0" Then
                    MsgBox "医保规则限制：不允许修改该病人的清算方式！", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        End If
    End If
    
    '将选择的清算方式上传到医保中心
    '对XML DomDocument对象进行初始化
    If InitXML = False Then Exit Sub
    Call InsertChild(mdomInput.documentElement, "PERSONCODE", mstr医保号)
    Call InsertChild(mdomInput.documentElement, "RECKONINGTYPE", txt基本统筹清算标准.Tag)
    Call InsertChild(mdomInput.documentElement, "SINGLEILLNESSCODE", txt疾病信息.Tag)
    Call InsertChild(mdomInput.documentElement, "OPERATOR", gstrUserName) ' 操作员
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(zlDatabase.Currentdate(), "yyyy-MM-dd HH:mm:ss")) ' 办理日期
    If CommServer("SETRECKONINGTYPE") = False Then Exit Sub
    
    gstrSQL = "zl_保险帐户_更新信息(" & mlng病人ID & "," & mint险类 & ",'单病种','''" & txt疾病信息.Tag & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存单病种编码")
    gstrSQL = "zl_保险帐户_更新信息(" & mlng病人ID & "," & mint险类 & ",'清算方式','''" & txt基本统筹清算标准.Tag & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存单病种编码")
    
    If mlng主页ID <> 0 Then
        '保存设置的清算方式
        gstrSQL = "ZL_医保病人住院信息_INSERT(" & _
                  mlng病人ID & "," & mlng主页ID & ",'" & gstrUserName & "',2," & mint保险类别 & ",'" & Split(txt疾病信息.Text, ")")(1) & "',NULL,NULL," & _
                  "NULL,NULL,NULL,NULL,NULL,'" & txt疾病信息.Tag & "','" & txt基本统筹清算标准.Text & "','" & txt基本统筹分担比例.Text & "','" & _
                  txt大额补助清算标准.Text & "','" & txt大额补助分担比例.Text & "','" & txt基本统筹清算标准.Tag & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "保存单病种编码")
        '程池富 2011-05-09 记录设置清算方式的日志，备查
        gstrSQL = "zl_清算信息日志_INSERT(" & mlng病人ID & "," & mlng主页ID & ",'" & txt疾病信息.Text & "','" & txt基本统筹清算标准.Tag & "','" & UserInfo.姓名 & "',sysdate)"
        gcnGYYB.Execute gstrSQL
    End If
    MsgBox "清算方式设置成功！", vbInformation, gstrSysName
    
    mblnOK = True
    Unload Me
End Sub

Private Sub cmd疾病信息_Click()
    Dim blnReturn As Boolean
    blnReturn = frmListSel.ShowSelect(mint险类, mrs病种, "ID", "单病种选择", "请选择单病种：")
    If Not blnReturn Then mrs病种.Filter = 0: Exit Sub
    
    txt疾病信息.Text = "(" & mrs病种!编码 & ")" & mrs病种!名称
    txt疾病信息.Tag = mrs病种!编码
    txt基本统筹清算标准.Tag = mrs病种!清算方式
    txt基本统筹清算标准.Text = Nvl(mrs病种!基本统筹清算标准)
    txt基本统筹分担比例.Text = Nvl(mrs病种!基本统筹分担比例)
    txt大额补助分担比例.Text = Nvl(mrs病种!大额补助分担比例)
    txt大额补助清算标准.Text = Nvl(mrs病种!大额补助清算标准)
    mrs病种.Filter = 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    On Error Resume Next
    '读取该病人的医保信息
    gstrSQL = "Select 保险类别,单病种 From 保险帐户 Where 病人ID=[1] And 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取该病人的医保信息", mlng病人ID, mint险类)
    txt疾病信息.Text = Nvl(rsTemp!单病种)
    mint保险类别 = rsTemp!保险类别
    mbln居民 = (rsTemp!保险类别 = "6")
    
    Call Get验证_贵阳(1, mstr卡号, mstr医保号, mstr分中心编号, mstr密码, mlng病人ID)
    
    Call 获取单病种
    Call 显示病种信息
End Sub

Public Function ShowSelect(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal int险类 As Integer, ByVal frmParent As Object) As Boolean
    mblnOK = False
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    mint险类 = int险类
    Me.Show 1, frmParent
    ShowSelect = mblnOK
End Function

Private Function 获取单病种() As Boolean
    Dim strFields As String, strValues As String
    Dim str编码 As String, str名称 As String, str简码 As String, str清算方式 As String
    Dim str基本统筹清单标准 As String, str基本统筹分担比例 As String
    Dim str大额补助清算标准 As String, str大额补助分担比例 As String
    
    Dim str当前日期 As String, str开始日期 As String, str结束日期 As String
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
    
    Set mrs病种 = New ADODB.Recordset
    strFields = "ID," & adVarChar & ",30|" & _
                "编码," & adLongVarChar & ",30|" & _
                "名称," & adLongVarChar & ",200|" & _
                "简码," & adLongVarChar & ",30|" & _
                "清算方式," & adLongVarChar & ",10|" & _
                "基本统筹清算标准," & adLongVarChar & ",500|" & _
                "基本统筹分担比例," & adLongVarChar & ",500|" & _
                "大额补助清算标准," & adLongVarChar & ",500|" & _
                "大额补助分担比例," & adLongVarChar & ",500"
    Call Record_Init(mrs病种, strFields)
    
    '对XML DomDocument对象进行初始化
    If InitXML = False Then Exit Function
    Call InsertChild(mdomInput.documentElement, "MITYPE", IIf(mbln居民, "2", "1"))
    If CommServer("QUERYHOSPSINGLEILLNESS") = False Then Exit Function
    
    Set nodRowset = mdomOutput.documentElement.selectSingleNode("ROWSET")
    If nodRowset Is Nothing Then Exit Function
    '根据编码得到险种名称
    str当前日期 = Format(zlDatabase.Currentdate(), "yyyy-MM-dd")
    strFields = "ID|编码|名称|简码|清算方式|基本统筹清算标准|基本统筹分担比例|大额补助清算标准|大额补助分担比例"
    
    '固定增加控制线结算
    str编码 = ""
    str名称 = "控制线清算"
    str清算方式 = 1
    str基本统筹清单标准 = 0
    str基本统筹分担比例 = 0
    str大额补助清算标准 = 0
    str大额补助分担比例 = 0
    str开始日期 = ""
    str结束日期 = ""
    str简码 = zlCommFun.SpellCode(str名称)
    strValues = str编码 & "|" & str编码 & "|" & str名称 & "|" & str简码 & "|" & str清算方式 & "|" & str基本统筹清单标准 & "|" & str基本统筹分担比例 & "|" & str大额补助清算标准 & "|" & str大额补助分担比例
    Call Record_Add(mrs病种, strFields, strValues)

    For Each nodRow In nodRowset.childNodes
        str编码 = GetAttributeValue(nodRow, "SINGLEILLNESSCODE")
        str名称 = GetAttributeValue(nodRow, "SINGLEILLNESSNAME")
        str清算方式 = GetAttributeValue(nodRow, "RECKONINGTYPE")
        str基本统筹清单标准 = GetAttributeValue(nodRow, "PAYSTD")
        str基本统筹分担比例 = GetAttributeValue(nodRow, "PAYRATE")
        str大额补助清算标准 = GetAttributeValue(nodRow, "PAY2STD")
        str大额补助分担比例 = GetAttributeValue(nodRow, "PAY2RATE")
        str开始日期 = Mid(GetAttributeValue(nodRow, "STARTDATE"), 1, 10)
        str结束日期 = Mid(GetAttributeValue(nodRow, "ENDDATE"), 1, 10)
        str简码 = zlCommFun.SpellCode(str名称)
        If str编码 <> "" And str当前日期 >= str开始日期 And str当前日期 <= str结束日期 Then
            strValues = str编码 & "|" & str编码 & "|" & str名称 & "|" & str简码 & "|" & str清算方式 & "|" & _
                str基本统筹清单标准 & "|" & str基本统筹分担比例 & "|" & str大额补助清算标准 & "|" & str大额补助分担比例
            Call Record_Add(mrs病种, strFields, strValues)
        End If
    Next
    获取单病种 = True
End Function

Private Function 显示病种信息(Optional ByVal bln任意匹配 As Boolean = False) As Boolean
    Dim blnReturn As Boolean
    Dim StrInput As String, strFilter As String
    
    If Trim(txt疾病信息.Text) = "" Then Exit Function
    If InStr(1, txt疾病信息.Text, "(") <> 0 Then
        If InStr(1, txt疾病信息.Text, ")") <> 0 Then
            StrInput = Mid(txt疾病信息.Text, 2, InStr(1, txt疾病信息.Text, ")") - 2)
        Else
            StrInput = Mid(txt疾病信息.Text, 2, Len(txt疾病信息.Text) - 1)
        End If
    Else
        StrInput = txt疾病信息.Text
    End If
    'bln任意匹配:如果不是任意匹配，表明是从数据库里读上次已选择的病种，因此采取从左匹配，怕有编码存在相似的，而操作通过输入来查病种时需要任意匹配
    If bln任意匹配 Then
        StrInput = UCase("'" & StrInput & "*'")
        strFilter = "编码 Like " & StrInput & " Or 名称 Like " & StrInput & " Or 简码 Like " & StrInput
    Else
        StrInput = UCase("'" & StrInput & "'")
        strFilter = "编码=" & StrInput
    End If
    
    With mrs病种
        .Filter = strFilter
        If .RecordCount = 0 Then
            If bln任意匹配 Then
                MsgBox "没有找到指定的单病种！[病种编码为:" & UCase(txt疾病信息.Text) & "]", vbInformation, gstrSysName
            End If
            Call zlControl.TxtSelAll(txt疾病信息)
            txt疾病信息.Text = ""
            txt疾病信息.Tag = ""
            txt基本统筹清算标准.Text = ""
            txt基本统筹清算标准.Tag = 1
            txt基本统筹分担比例.Text = ""
            txt大额补助分担比例.Text = ""
            txt大额补助清算标准.Text = ""
            .Filter = 0
            Exit Function
        Else
            If mrs病种.RecordCount > 1 Then
                blnReturn = frmListSel.ShowSelect(mint险类, mrs病种, "ID", "单病种选择", "请选择单病种：")
            Else
                blnReturn = True
            End If
            If blnReturn = False Then
                txt疾病信息.Text = ""
                txt疾病信息.Tag = ""
                txt基本统筹清算标准.Text = ""
                txt基本统筹分担比例.Text = ""
                txt大额补助分担比例.Text = ""
                txt大额补助清算标准.Text = ""
                txt基本统筹清算标准.Tag = 1
                Call zlControl.TxtSelAll(txt疾病信息)
            Else
                txt疾病信息.Text = "(" & mrs病种!编码 & ")" & mrs病种!名称
                txt疾病信息.Tag = mrs病种!编码
                txt基本统筹清算标准.Tag = mrs病种!清算方式
                txt基本统筹清算标准.Text = Nvl(mrs病种!基本统筹清算标准)
                txt基本统筹分担比例.Text = Nvl(mrs病种!基本统筹分担比例)
                txt大额补助分担比例.Text = Nvl(mrs病种!大额补助分担比例)
                txt大额补助清算标准.Text = Nvl(mrs病种!大额补助清算标准)
                显示病种信息 = True
            End If
        End If
        .Filter = 0
    End With
End Function

Private Sub txt疾病信息_GotFocus()
    Call zlControl.TxtSelAll(txt疾病信息)
End Sub

Private Sub txt疾病信息_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(txt疾病信息.Text) = "" Then Exit Sub
    
    If Not 显示病种信息(True) Then Exit Sub
    Call zlCommFun.PressKey(vbKeyTab)
End Sub
