VERSION 5.00
Begin VB.Form frmSet重庆银海版 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "运行参数设置"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5880
   Icon            =   "frmSet重庆银海版.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fra医院等级 
      Caption         =   "医院等级"
      Height          =   1365
      Left            =   150
      TabIndex        =   8
      Top             =   1980
      Width           =   4155
      Begin VB.CommandButton cmd医院等级 
         Caption         =   "…"
         Height          =   300
         Left            =   3585
         TabIndex        =   12
         Top             =   900
         Width           =   285
      End
      Begin VB.TextBox txt医院等级 
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1350
         MaxLength       =   40
         TabIndex        =   11
         Top             =   900
         Width           =   2235
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "医院等级(&L)"
         Height          =   180
         Index           =   3
         Left            =   300
         TabIndex        =   10
         Top             =   960
         Width           =   990
      End
      Begin VB.Label lbl说明 
         Caption         =   "    该等级用于计算部分按医院等级进行限价的诊疗项目的实际价格。"
         Height          =   480
         Left            =   390
         TabIndex        =   9
         Top             =   330
         Width           =   3450
      End
   End
   Begin VB.Frame fra医保服务器 
      Caption         =   "医院前置医保服务器"
      Height          =   1605
      Left            =   150
      TabIndex        =   0
      Top             =   210
      Width           =   4155
      Begin VB.CommandButton cmdTest 
         Caption         =   "测试(&T)"
         Height          =   1095
         Left            =   3000
         TabIndex        =   7
         Top             =   330
         Width           =   1005
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   6
         Top             =   1110
         Width           =   1635
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1260
         MaxLength       =   40
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   720
         Width           =   1635
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   2
         Top             =   330
         Width           =   1635
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "服务器(&S)"
         Height          =   180
         Index           =   2
         Left            =   390
         TabIndex        =   5
         Top             =   1170
         Width           =   810
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "密码(&P)"
         Height          =   180
         Index           =   1
         Left            =   570
         TabIndex        =   3
         Top             =   780
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "用户名(&U)"
         Height          =   180
         Index           =   0
         Left            =   390
         TabIndex        =   1
         Top             =   390
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4560
      TabIndex        =   14
      Top             =   780
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4560
      TabIndex        =   13
      Top             =   300
      Width           =   1100
   End
End
Attribute VB_Name = "frmSet重庆银海版"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum enum文本
    text医保用户 = 0
    Text医保密码 = 1
    Text医保服务器 = 2
End Enum

Private mblnOK As Boolean
Private mblnChange As Boolean
Private Const madLongVarCharDefault As Integer = 10          '字符型字段缺省长度
Private Const madDoubleDefault As Integer = 18               '数字型字段缺省长度
Private Const madDbDateDefault As Integer = 20               '日期型字段缺省长度

Dim mcnTest As New ADODB.Connection

Private Sub cmdTest_Click()
    Dim rsTemp As New ADODB.Recordset
    If mcnTest.State = adStateOpen Then mcnTest.Close
    
    If OraDataOpen(mcnTest, txtEdit(Text医保服务器).Text, txtEdit(text医保用户).Text, txtEdit(Text医保密码).Tag) = False Then
        Exit Sub
    End If
    
    MsgBox "连接成功！", vbInformation, gstrSysName
End Sub

Private Sub cmd医院等级_Click()
    Dim strFields As String
    Dim rsHos_Info As New ADODB.Recordset
    On Error GoTo errHand
    If Not 医保初始化_重庆银海版(True) Then Exit Sub
    
    '初始化内部记录集
    strFields = "医疗机构编号," & adVarChar & "," & 20 & "|医疗机构名称," & adVarChar & "," & 50 & _
                "|医疗机构等级," & adVarChar & "," & 5 & "|最高起付标准," & adVarChar & "," & 20
    Call Record_Init(rsHos_Info, strFields)
    
    '调用接口获取医疗机构信息
    Call 调用接口_准备_重庆银海版("05", "C:\CQYB_YH\Hos_info.txt")
    If Not 调用接口_重庆银海版 Then Exit Sub
    If Not AnalyFile_HosInfo(rsHos_Info) Then Exit Sub
    
    '让操作员选择医院等级
    If frmListSel.ShowSelect(TYPE_重庆银海版, rsHos_Info, "医疗机构编号", "请选择医院等级！", "以下是中心认可的医疗机构信息:") = True Then
        '01-一级;05-二级;08-三级
        txt医院等级.Tag = rsHos_Info!医疗机构编号
        txt医院等级.Text = IIf(rsHos_Info!医疗机构等级 = "01", "三级", IIf(rsHos_Info!医疗机构等级 = "05", "二级", "一级"))
        MsgBox "成功获取本院的医院等级和医院代码！", vbInformation, gstrSysName
    End If
    
    rsHos_Info.Filter = 0
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    rsHos_Info.Filter = 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
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
    gstrSQL = "zl_保险参数_Delete(" & TYPE_重庆银海版 & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '新增参数数据
    gstrSQL = "zl_保险参数_Insert(" & TYPE_重庆银海版 & ",null,'医保用户名','" & txtEdit(text医保用户).Text & "',1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_重庆银海版 & ",null,'医保用户密码','" & txtEdit(Text医保密码).Tag & "',2)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_重庆银海版 & ",null,'医保服务器','" & txtEdit(Text医保服务器).Text & "',3)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_重庆银海版 & ",null,'医院等级','" & txt医院等级.Text & "',4)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '更新医院编号
    gstrSQL = "Select 名称,说明,是否禁止 From 保险类别 Where 序号= [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_重庆银海版)
    '调试重庆医保银海版 204-04-07
    gstrSQL = "zl_保险类别_Update(" & TYPE_重庆银海版 & ",'" & rsTemp!名称 & "','" & IIf(IsNull(rsTemp!说明), "", rsTemp!说明) & "','" & Me.txt医院等级.Tag & "'," & IIf(IsNull(rsTemp!是否禁止), 0, rsTemp!是否禁止) & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    gcnOracle.CommitTrans
    SaveData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    gcnOracle.RollbackTrans
End Function

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

Public Function 参数设置() As Boolean
'功能：设置与东大阿尔派的医保接口
    Dim rsTemp As New ADODB.Recordset
    Dim str参数值 As String
    
    mblnOK = False
    
    On Error GoTo errHandle
    
    '取保险参数
    gstrSQL = "select 参数名,参数值 from 保险参数 " & _
              " where 险类=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_重庆银海版)
    Do Until rsTemp.EOF
        str参数值 = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
        Select Case rsTemp("参数名")
            Case "医保用户名"
                txtEdit(text医保用户) = str参数值
            Case "医保服务器"
                txtEdit(Text医保服务器) = str参数值
            Case "医保用户密码"
                txtEdit(Text医保密码).Text = "        "    '假密码
                txtEdit(Text医保密码).Tag = str参数值
            Case "医院等级"
                txt医院等级.Text = str参数值
        End Select
        
        rsTemp.MoveNext
    Loop
    '取医院编号,因为保存时要更新
    gstrSQL = "Select 医院编码 From 保险类别 Where 序号= [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_重庆银海版)
    txt医院等级.Tag = Nvl(rsTemp!医院编码)
    
    mblnChange = False
    frmSet重庆银海版.Show vbModal, frm医保类别
    
    参数设置 = mblnOK
    Exit Function

errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

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

Private Function AnalyFile_HosInfo(ByVal rsHos_Info As ADODB.Recordset) As Boolean
    '分析接口返回的待遇文件，并保存到中间库（预结算返回的结果80%不准确，因此建议不保存）
    Dim lngCol As Long, lngCols As Long
    Dim strData As String, strHosinfo As String, strBuffer As String, strFields As String
    Dim arrCol
    Dim objStream As TextStream, objFileSystem As New FileSystemObject
    
    On Error GoTo errHand
    
    If Not objFileSystem.FileExists("C:\CQYB_YH\Hos_info.txt") Then Exit Function
    Set objStream = objFileSystem.OpenTextFile("C:\CQYB_YH\Hos_info.txt", ForReading, False, TristateMixed)
    
    strFields = ""
    For lngCol = 0 To rsHos_Info.Fields.Count - 1
        strFields = strFields & "|" & rsHos_Info.Fields(lngCol).Name
    Next
    strFields = Mid(strFields, 2)
    
    Do While Not objStream.AtEndOfStream
        strBuffer = objStream.ReadLine
        strHosinfo = ""
        arrCol = Split(strBuffer, vbTab)
        lngCols = UBound(arrCol)
        For lngCol = 0 To lngCols
            strHosinfo = strHosinfo & "|" & arrCol(lngCol)
        Next
        strHosinfo = Mid(strHosinfo, 2)
        Call Record_Add(rsHos_Info, strFields, strHosinfo)
    Loop
    objStream.Close
    
    AnalyFile_HosInfo = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
