VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCustomQueryCall 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "自定义查询"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5805
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCustomQueryCall.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSScriptControlCtl.ScriptControl sctExecute 
      Left            =   855
      Top             =   3210
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin MSComCtl2.DTPicker dtpObj 
      Height          =   330
      Index           =   0
      Left            =   1965
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   100663297
      CurrentDate     =   41297
   End
   Begin VB.TextBox txtObj 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   1950
      TabIndex        =   5
      Top             =   1485
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.ComboBox cbxObj 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   1965
      TabIndex        =   4
      Text            =   "cbxObj"
      Top             =   1920
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.ListBox lstObj 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      Index           =   0
      ItemData        =   "frmCustomQueryCall.frx":000C
      Left            =   1980
      List            =   "frmCustomQueryCall.frx":000E
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   2310
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Frame framButton 
      Height          =   780
      Left            =   -45
      TabIndex        =   0
      Top             =   4620
      Width           =   5895
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取 消(&C)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4410
         TabIndex        =   2
         Top             =   240
         Width           =   1185
      End
      Begin VB.CommandButton cmdSure 
         Caption         =   "确 定(&S)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3045
         TabIndex        =   1
         Top             =   240
         Width           =   1185
      End
   End
   Begin VB.Image imgQuery 
      Height          =   720
      Left            =   105
      Picture         =   "frmCustomQueryCall.frx":0010
      Stretch         =   -1  'True
      Top             =   135
      Width           =   720
   End
   Begin VB.Label labMemo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Height          =   630
      Left            =   870
      TabIndex        =   9
      Top             =   165
      Width           =   4650
   End
   Begin VB.Shape shpBack 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00CEFFFA&
      FillStyle       =   0  'Solid
      Height          =   840
      Left            =   75
      Shape           =   4  'Rounded Rectangle
      Top             =   60
      Width           =   5670
   End
   Begin VB.Label labError 
      Alignment       =   2  'Center
      Caption         =   "没有需要录入的项目。"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1425
      TabIndex        =   8
      Top             =   2325
      Visible         =   0   'False
      Width           =   3435
   End
   Begin VB.Label labObj 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "标题占位:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   615
      TabIndex        =   7
      Top             =   1140
      Visible         =   0   'False
      Width           =   885
   End
End
Attribute VB_Name = "frmCustomQueryCall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mlngSchemeId As Long            '查询方案ID
Public mlngDepartId As Long            '当前科室ID
Public mlngModule As Long              '当前模块编号
Public mstrReturnQuery As String       '“确定”按钮按下后，返回的查询sql
Public mstrPars As Variant

Private mobjLastControl As Object       '创建录入界面时，保存的上一次创建的录入组件
Private mstrInputProTotal As String     '保存总的需要录入的项目
Private mrsCfgData As ADODB.Recordset   '查询方案配置的数据
Private mobjReg As Scripting.Dictionary  '保存录入组件值改变后，需关联触发的控件


Private mstrSchemeQuerySql As String
Public mintEnabledRules As Integer      '是否启用了规则，1-启用了规则,0-没有启用规则查询


Public Function ShowCustomQuery(ByVal lngSchemeId As String, ByVal lngDepartId As Long, _
    ByVal lngModule As Long, ByRef strPars As Variant, objOwner As Object) As String
    
    ShowCustomQuery = ""
        
    Me.mlngSchemeId = lngSchemeId
    Me.mlngDepartId = lngDepartId
    Me.mlngModule = lngModule
    
    Call Me.Show(1, objOwner)
    
    strPars = Me.mstrPars
    ShowCustomQuery = Me.mstrReturnQuery
End Function


Private Sub cbxObj_Change(Index As Integer)
'下拉框数据值被用户改变后，需要处理的数据加载
On Error GoTo errHandle
    Call ControlValChange(cbxObj, Index)
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdCancel_Click()
On Error GoTo errHandle
    mstrReturnQuery = ""
    
    Unload Me
    
    Exit Sub
errHandle:
End Sub

Private Sub cmdSure_Click()
On Error GoTo errHandle
    Dim strSql As String
    Dim varValues As Variant
    
    If Trim(mstrSchemeQuerySql) = "" Then
        MsgBoxD Me, "无效的查询方案，不能继续执行。", vbOKOnly, Me.Caption
        Exit Sub
    End If

    strSql = mstrSchemeQuerySql
    Call InitQuerySqlAndPars(strSql, varValues)
    
    mstrReturnQuery = strSql
    mstrPars = varValues
    
    Unload Me
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Public Sub GetQuerySqlAndPars(ByVal lngSchemeId As Long, ByRef strQuerySql As String, ByRef varParameters As Variant)
    Dim varValues As Variant
    Dim strReturn As String
    Dim rsData As ADODB.Recordset
    
    mintEnabledRules = 0
    
    Set rsData = GetSchemeData(lngSchemeId)
    
    If rsData Is Nothing Then Exit Sub
    
    If rsData.RecordCount <= 0 Then
        Call MsgBoxD(Me, "未找到配置信息，不能完成对应查询。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    strReturn = Nvl(rsData!查询语句)
    mintEnabledRules = Val(Nvl(rsData!是否启用规则))
    
    Call InitQuerySqlAndPars(strReturn, varValues)
    
    strQuerySql = strReturn
    varParameters = varValues
End Sub

Private Function GetSchemeData(ByVal lngSchemeId As Long) As ADODB.Recordset
On Error GoTo errHandle
    Dim strSql As String
    
    strSql = "select 查询语句,方案说明,是否启用规则 from 影像查询方案 where id=[1]"
    Set GetSchemeData = zlDatabase.OpenSQLRecord(strSql, "查询方案信息", lngSchemeId)
Exit Function
errHandle:
    Set GetSchemeData = Nothing
End Function

Private Sub InitQuerySqlAndPars(ByRef strQuerySql As String, ByRef varParameters As Variant)
    Dim strSql As String
    Dim strPars(20) As String
    Dim varValues(20) As Variant
    Dim i As Long
    Dim strCurPar As String
    Dim strField As String

    strSql = strQuerySql
    
    If Not GetParameterNames(strSql, strPars()) Then
        strQuerySql = strSql
        varParameters = strPars()
        
        Exit Sub
    End If
    
    '格式化sql查询语句，将参数和查询值分离
    For i = 1 To 20
        strCurPar = strPars(i)
        varValues(i) = GetParameterValue(strCurPar)
        
        If strCurPar <> "" Then
            If Trim(varValues(i)) = "" Then    '如果录入的数据为空，则不使用该条件
                strField = GetParameterField(strCurPar, strSql)
                If strField <> "" Then
                    strSql = Replace(strSql, strCurPar, 1)
                    strSql = Replace(strSql, strField, 1)
                Else
                    strSql = Replace(strSql, strCurPar, "[" & i & "]")
                End If
            Else
                strSql = Replace(strSql, strCurPar, "[" & i & "]")
            End If
            
        End If
    Next i
    
    strQuerySql = strSql
    varParameters = varValues()
End Sub

Private Function GetParameterField(ByVal strParameter As String, ByVal strSql As String) As String
'获取参数在查询语句中对应使用的字段名称
    Dim strTemp As String
    Dim lngParIndex As Long
    Dim lngFieldIndex As Long
    Dim strPrefix As String
    
    GetParameterField = ""
    
    '例strSql = "select id from tab a where a.test = [par1] "
    lngParIndex = InStr(strSql, strParameter)
    If lngParIndex <= 0 Then Exit Function
    
    
    '执行后取得的strsql部分将为“select id from tab a where a.test = ”
    strTemp = Mid(strSql, 1, lngParIndex - 1)
    
    
    lngFieldIndex = InStrRev(strTemp, ".")
    
    '判断字段与参数之间是否存在相关函数，如下：
    'select id from tab a where a.pid is null and insert([par1], a.Field) > 0
    If InStr(lngFieldIndex, UCase(strTemp), "INSTR(") > 0 Then
        '获取insert部分的字段......
        Exit Function
    ElseIf InStr(lngFieldIndex, UCase(strTemp), "DECODE(") > 0 Then
        '获取decode部分的字段......
        Exit Function
    Else
    
        strPrefix = strTemp
    
        '执行后取得的strsql部分将为“test = ”
        strTemp = Trim(Mid(strTemp, lngFieldIndex + 1, 100))
        
                
        strPrefix = Mid(strPrefix, 1, lngFieldIndex - 1)
        strPrefix = Trim(Mid(strPrefix, InStrRev(strPrefix, " "), 100))
        
        
        '以下语句执行后，将取得字段名
        
        lngFieldIndex = InStr(strTemp, " ")
        
        If lngFieldIndex > 0 Then
            strTemp = Mid(strTemp, 1, lngFieldIndex - 1)
        Else
            strTemp = Mid(strTemp, 1, Len(strTemp) - 1)
        End If
    End If
    
    GetParameterField = strPrefix & "." & strTemp
End Function

Private Sub Form_Unload(Cancel As Integer)
    Dim objFree As Object
    
    For Each objFree In txtObj
        If Not objFree Is Nothing Then
            If objFree.Index <> 0 Then Unload objFree
        End If
    Next
    
    For Each objFree In cbxObj
        If Not objFree Is Nothing Then
            If objFree.Index <> 0 Then Unload objFree
        End If
    Next
    
    For Each objFree In dtpObj
        If Not objFree Is Nothing Then
            If objFree.Index <> 0 Then Unload objFree
        End If
    Next
    
    For Each objFree In lstObj
        If Not objFree Is Nothing Then
            If objFree.Index <> 0 Then Unload objFree
        End If
    Next
    
    Set mobjReg = Nothing
    
    Call sctExecute.Reset
End Sub

Private Sub cbxObj_Click(Index As Integer)
'下拉框数据值被用户改变后，需要处理的数据加载
On Error GoTo errHandle
    Call ControlValChange(cbxObj, Index)
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub dtpObj_Change(Index As Integer)
'日期框数据值被用户改变后，需要处理的数据加载
On Error GoTo errHandle
    Call ControlValChange(dtpObj, Index)
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub lstObj_Click(Index As Integer)
'多选框数据值被用户改变后，需要处理的数据加载
On Error GoTo errHandle
    Call ControlValChange(lstObj, Index)
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub txtObj_Change(Index As Integer)
'文本框数据值被用户改变后，需要处理的数据加载
On Error GoTo errHandle
    Dim i As Long
    Dim strCurControlName As String
    Dim strKey As Variant
    
    Call ControlValChange(txtObj, Index)
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ControlValChange(objControl As Object, ByVal intIndex As Integer)
'控件数据值被用户改变后，需要处理的数据加载
On Error GoTo errHandle
    Dim strCurControlName As String
    Dim strKey As Variant
    
    If Not objControl(intIndex).Visible Then Exit Sub
    
    strCurControlName = objControl(intIndex).Tag
    
    For Each strKey In mobjReg.Keys
        If InStr(strKey, strCurControlName) > 0 Then
            Call UpdateInputData(mobjReg(strKey).Tag)
        End If
    Next
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Form_Load()
On Error GoTo errHandle
    Dim rsData As ADODB.Recordset
    Dim strSql As String
    
    mstrReturnQuery = ""
    mstrInputProTotal = ""
    mstrSchemeQuerySql = ""
    mintEnabledRules = 0
    
    Set mobjLastControl = Nothing
    Set mrsCfgData = Nothing
    
    Set rsData = GetSchemeData(mlngSchemeId)
    
    If rsData Is Nothing Then Exit Sub
    
    If rsData.RecordCount <= 0 Then
        Call MsgBoxD(Me, "未找到配置信息，不能完成界面加载。", vbOKOnly, Me.Caption)
        Exit Sub
    End If
    
    mstrSchemeQuerySql = Nvl(rsData!查询语句)
    mintEnabledRules = Val(Nvl(rsData!是否启用规则))
    
    labMemo.Caption = "说明：" & Nvl(rsData!方案说明) 'IIf(Trim(Nvl(rsData!方案说明)) <> "", "说明：" & Nvl(rsData!方案说明), "")
    
    
    
    Set mobjReg = New Scripting.Dictionary
    
        
    Call ConfigFaceInput
    
    Call sctExecute.AddObject("Me", Me, True)
    
    Call LoadInputData
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub LoadInputData()
'加载可选录入数据
    Dim i As Long
    Dim strDataFrom As String
    Dim strParameters(20) As String
    Dim blnHasInputPar As Boolean
    Dim blnHasPar As Boolean
    Dim rsTemp As ADODB.Recordset
    Dim lngObjOrder As Long
    Dim lngInputType As Long
    Dim strDefaultValue As String
    Dim objInputControl As Object
    Dim strParValues(20) As Variant
    Dim strParField As String
    
    
    If mrsCfgData Is Nothing Then Exit Sub
    If mrsCfgData.RecordCount <= 0 Then Exit Sub
    
    mrsCfgData.MoveFirst
    
    While Not mrsCfgData.EOF
        strDataFrom = Nvl(mrsCfgData!数据来源)
        lngObjOrder = Val(Nvl(mrsCfgData!录入顺序))
        strDefaultValue = GetParameterValue(Nvl(mrsCfgData!默认值))
        lngInputType = Val(Nvl(mrsCfgData!录入类型))
        
        If strDataFrom <> "" Then
            '获取该sql源包含的所有参数名称
            blnHasPar = GetParameterNames(strDataFrom, strParameters)
            
            For i = 1 To 20
                strParValues(i) = GetParameterValue(strParameters(i))
                If strParameters(i) <> "" Then
                    If Trim(strParValues(i)) = "" Then    '如果录入的数据为空，则不使用该条件
                        strParField = GetParameterField(strParameters(i), strDataFrom)
                        If strParField <> "" Then
'                            strDataFrom = Replace(strDataFrom, strParameters(i), strParField)
                            strDataFrom = Replace(strDataFrom, strParameters(i), 1)
                            strDataFrom = Replace(strDataFrom, strParField, 1)
                        Else
                            strDataFrom = Replace(strDataFrom, strParameters(i), "[" & i & "]")
                        End If
                    Else
                        strDataFrom = Replace(strDataFrom, strParameters(i), "[" & i & "]")
                    End If
                End If
            Next i
            
            Set rsTemp = zlDatabase.OpenSQLRecord(strDataFrom, "配置可录数据", strParValues(1), strParValues(2), strParValues(3), _
                                    strParValues(4), strParValues(5), strParValues(6), strParValues(7), strParValues(8), _
                                    strParValues(9), strParValues(10), strParValues(11), strParValues(12), strParValues(13), _
                                    strParValues(14), strParValues(15), strParValues(16), strParValues(17), strParValues(18), _
                                    strParValues(19), strParValues(20))
            
            If rsTemp.RecordCount > 0 Then
                Select Case lngInputType
                
                    Case 0
                        '读取文本框显示的数据
                        Call SetControlValue(lngInputType, lngObjOrder, rsTemp(0).value)
                        
                        If strDefaultValue <> "" Then
                            Call SetControlValue(lngInputType, lngObjOrder, strDefaultValue)
                        End If
                    Case 1, 2, 3
                        '读取日期框显示的数据
                        Call SetControlValue(lngInputType, lngObjOrder, rsTemp(0).value)
                        
                        If strDefaultValue <> "" Then
                            Call SetControlValue(lngInputType, lngObjOrder, strDefaultValue)
                        End If
                    Case 4
                        '读取下拉框显示的数据
                        cbxObj(lngObjOrder).Clear
                        
                        While Not rsTemp.EOF
                            cbxObj(lngObjOrder).AddItem rsTemp(0).value
                            rsTemp.MoveNext
                        Wend
                        
                        If strDefaultValue <> "" Then
                            Call SetControlValue(lngInputType, lngObjOrder, strDefaultValue)
                        Else
                            cbxObj(lngObjOrder).ListIndex = 0
                        End If
                    Case 5
                        '读取可多选列表框显示的数据
                        lstObj(lngObjOrder).Clear
                        
                        While Not rsTemp.EOF
                            lstObj(lngObjOrder).AddItem rsTemp(0).value
                            
                            If InStr(strDefaultValue, rsTemp(0).value) > 0 Then
                                lstObj(lngObjOrder).Selected(lstObj(lngObjOrder).ListCount - 1) = True
                            End If
                            
                            rsTemp.MoveNext
                        Wend
                End Select
            End If
            
            '注册需要关联改变的录入控件
            For i = 1 To 20
                If InStr(mstrInputProTotal, strParameters(i)) > 0 And strParameters(i) <> "" Then
                    '配置录入值改变后需要对应改变的控件
                    Select Case lngInputType

                        Case 0
                            Set objInputControl = txtObj(lngObjOrder)
                        Case 1, 2, 3
                            Set objInputControl = dtpObj(lngObjOrder)
                        Case 4
                            Set objInputControl = cbxObj(lngObjOrder)
                        Case 5
                            Set objInputControl = lstObj(lngObjOrder)
                    End Select

                    Call mobjReg.Add(strParameters(i) & lngObjOrder, objInputControl)
                End If
            Next i
        Else
            '判断是否设置了默认值，如果设置了默认值，则需要加载默认值
            If strDefaultValue <> "" Then
                Call SetControlValue(lngInputType, lngObjOrder, strDefaultValue)
            End If
        End If
        
        mrsCfgData.MoveNext
    Wend
End Sub


Private Sub UpdateInputData(ByVal strInputName As String)
'更新录入数据显示
    Dim rsTemp As ADODB.Recordset
    Dim strTemp As String
    Dim lngInputType As Long
    Dim lngInputOrder As Long
    Dim strSqlFrom As String
    Dim strParameters(20) As String
    Dim i As Long
    Dim strParValue(20) As Variant
    Dim strField As String
    
    
    If mrsCfgData Is Nothing Then Exit Sub
    If mrsCfgData.RecordCount <= 0 Then Exit Sub
    
    strTemp = Replace(strInputName, "[", "")
    strTemp = Replace(strTemp, "]", "")
    
    '过滤出该录入项目对应的配置信息
    mrsCfgData.Filter = "录入项目='" & strTemp & "'"
    
    If mrsCfgData.RecordCount <= 0 Then Exit Sub
    
    lngInputType = Val(Nvl(mrsCfgData!录入类型))
    lngInputOrder = Val(Nvl(mrsCfgData!录入顺序))
    strSqlFrom = Nvl(mrsCfgData!数据来源)
    
    If strSqlFrom = "" Then Exit Sub
    
    '获取该sql源包含的所有参数名称
    Call GetParameterNames(strSqlFrom, strParameters)
    
    For i = 1 To 20
        strParValue(i) = GetParameterValue(strParameters(i))
        If strParameters(i) <> "" Then
            If Trim(strParValue(i)) = "" Then '如果录入的数据为空，则不使用该条件
                strField = GetParameterField(strParameters(i), strSqlFrom)
                If strField <> "" Then
'                    strSqlFrom = Replace(strSqlFrom, strParameters(i), strField)
                    strSqlFrom = Replace(strSqlFrom, strParameters(i), 1)
                    strSqlFrom = Replace(strSqlFrom, strField, 1)
                Else
                    strSqlFrom = Replace(strSqlFrom, strParameters(i), "[" & i & "]")
                End If
            Else
                strSqlFrom = Replace(strSqlFrom, strParameters(i), "[" & i & "]")
            End If
        End If
    Next i
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSqlFrom, "配置可录数据", strParValue(1), strParValue(2), strParValue(3), _
                            strParValue(4), strParValue(5), strParValue(6), strParValue(7), strParValue(8), _
                            strParValue(9), strParValue(10), strParValue(11), strParValue(12), strParValue(13), _
                            strParValue(14), strParValue(15), strParValue(16), strParValue(17), strParValue(18), _
                            strParValue(19), strParValue(20))
                
    
    Select Case lngInputType
    
        Case 0
            If rsTemp.RecordCount <= 0 Then Exit Sub
            Call SetControlValue(lngInputType, lngInputOrder, rsTemp(0).value)
            
        Case 1, 2, 3
            If rsTemp.RecordCount <= 0 Then Exit Sub
            Call SetControlValue(lngInputType, lngInputOrder, rsTemp(0).value)
            
        Case 4
            cbxObj(lngInputOrder).Clear
            If rsTemp.RecordCount <= 0 Then Exit Sub
            
            While Not rsTemp.EOF
                cbxObj(lngInputOrder).AddItem rsTemp(0).value
                rsTemp.MoveNext
            Wend
            
            If cbxObj(lngInputOrder).ListCount > 0 Then cbxObj(lngInputOrder).ListIndex = 0
        Case 5
            lstObj(lngInputOrder).Clear
            If rsTemp.RecordCount <= 0 Then Exit Sub
            
            While Not rsTemp.EOF
                lstObj(lngInputOrder).AddItem rsTemp(0).value
                rsTemp.MoveNext
            Wend
    End Select
End Sub

Private Sub SetControlValue(ByVal lngInputType As Long, ByVal lngObjOrder As Long, ByVal strValue As String)
'对控件的文本或者value属性赋值
On Error Resume Next
    Dim i As Long
    
    Select Case lngInputType
        Case 0
            txtObj(lngObjOrder).Text = strValue
        Case 1, 2, 3
            dtpObj(lngObjOrder).value = strValue
        Case 4
            cbxObj(lngObjOrder).Text = strValue
        Case 5
            For i = 0 To lstObj(lngObjOrder).ListCount - 1
                If lstObj(lngObjOrder).list(i) = strValue Then
                    lstObj(lngObjOrder).Selected(i) = True
                End If
            Next i
    End Select
End Sub


Private Function GetParameterNames(ByVal strSqlFrom As String, ByRef strParameters() As String) As Boolean
'判断数据源sql语句是否包含参数
    Dim strTemp As String
    Dim lngStart As Long, lngEnd As Long
    Dim lngParCount As Long
    
    strTemp = strSqlFrom
    lngStart = InStr(strTemp, "[")
    lngEnd = InStr(strTemp, "]")
    
    GetParameterNames = False
'    blnHasInputPar = False
    
    If lngStart <= 0 Or lngEnd <= 0 Then Exit Function
    
    lngParCount = 0
    
    '循环获取所有的参数名称
    While lngStart > 0 And lngEnd > 0
        
        lngParCount = lngParCount + 1
        
        strTemp = Mid(strTemp, lngStart, 1024)
        lngEnd = InStr(strTemp, "]")
        
        strParameters(lngParCount) = Mid(strTemp, 1, lngEnd)
        
'        If InStr(mstrInputProTotal, strParameters(lngParCount)) > 0 Then
'            blnHasInputPar = True
'        End If
        
        strTemp = Mid(strTemp, lngEnd + 1, 1024)
        
        lngStart = InStr(strTemp, "[")
        lngEnd = InStr(strTemp, "]")
    Wend
       
    GetParameterNames = IIf(lngParCount > 0, True, False)
End Function


Private Sub ConfigFaceInput()
'配置界面录入
    
    Dim strSql As String
    
    strSql = "select 录入项目,录入类型,录入顺序,默认值,数据来源 from 影像查询配置 where 方案Id=[1] order by 录入顺序"
    
    Set mrsCfgData = zlDatabase.OpenSQLRecord(strSql, "获取查询配置", mlngSchemeId)
    If mrsCfgData.RecordCount <= 0 Then Exit Sub
    
    While Not mrsCfgData.EOF
        If mstrInputProTotal <> "" Then mstrInputProTotal = mstrInputProTotal & ","
        mstrInputProTotal = mstrInputProTotal & "[" & Nvl(mrsCfgData!录入项目) & "]"
        
'        Call CreateInputControl(Nvl(mrsCfgData!录入项目), Val(Nvl(mrsCfgData!录入类型)), GetParameterValue(Nvl(mrsCfgData!默认值)), Val(Nvl(mrsCfgData!录入顺序)))
        Call CreateInputControl(Nvl(mrsCfgData!录入项目), Val(Nvl(mrsCfgData!录入类型)), "", Val(Nvl(mrsCfgData!录入顺序)))
        mrsCfgData.MoveNext
    Wend
    
    
    If Not mobjLastControl Is Nothing Then
        framButton.Left = -45
        framButton.Width = Me.ScaleWidth + 90
    
        framButton.Top = mobjLastControl.Top + mobjLastControl.Height + 120 + 45
        Me.Height = framButton.Top + framButton.Height + 400 - 45
        
        labError.Visible = False
    Else
        labError.Visible = True
    End If
End Sub


Private Function GetParameterValue(ByVal strParameterName As String) As Variant
    Dim j As Long
    Dim objInputCon As Object

    GetParameterValue = ""
        
    If strParameterName = "" Then Exit Function
    If Not IsParameterFormat(strParameterName) Then
        '如果不是参数格式，则可能是直接由默认值配置传入的数据值，比如默认值配置的是“2012-03-05”，并没有采用“[当前时间]”方式
        GetParameterValue = strParameterName
        Exit Function
    End If
    
    Select Case strParameterName
        Case "[当前日期]"
            GetParameterValue = date
            Exit Function
        Case "[当前时间]"
            GetParameterValue = Now
            Exit Function
        Case "[当前用户ID]"
            GetParameterValue = UserInfo.ID
            Exit Function
        Case "[当前科室ID]"
            If mlngDepartId <= 0 Then
                GetParameterValue = ""
            Else
                GetParameterValue = mlngDepartId
            End If
            Exit Function
        Case "[当前系统编号]"
            GetParameterValue = glngSys
            Exit Function
        Case "[当前模块编号]"
            GetParameterValue = mlngModule
            Exit Function
        Case Else
            '获取文本框对应的值
           For Each objInputCon In txtObj
               If Not objInputCon Is Nothing Then
                   If objInputCon.Tag = strParameterName And objInputCon.Index <> 0 Then
                       GetParameterValue = objInputCon.Text
                       Exit Function
                   End If
               End If
           Next
           
           '获取日期框对应的值
           For Each objInputCon In dtpObj
               If Not objInputCon Is Nothing Then
                   If objInputCon.Tag = strParameterName And objInputCon.Index <> 0 Then
                       GetParameterValue = objInputCon.value
                       Exit Function
                   End If
               End If
           Next
           
           
           '获取下拉框的值
           For Each objInputCon In cbxObj
               If Not objInputCon Is Nothing Then
                   If objInputCon.Tag = strParameterName And objInputCon.Index <> 0 Then
                       GetParameterValue = objInputCon.Text
                       Exit Function
                   End If
               End If
           Next
           
           '获取可多选列表框的值
           For Each objInputCon In lstObj
               If Not objInputCon Is Nothing Then
                   If objInputCon.Tag = strParameterName And objInputCon.Index <> 0 Then
                       For j = 0 To objInputCon.ListCount - 1
                           If objInputCon.Selected(j) Then
                               If GetParameterValue <> "" Then GetParameterValue = GetParameterValue & ","
                               GetParameterValue = GetParameterValue & objInputCon.list(j)
                           End If
                       Next j
                       
                       Exit Function
                   End If
               End If
           Next
    End Select
    
    '在前面的代码中，如果找到对应的参数，就会直接将值覆盖函数并返回，如果执行到这里，说明没有找到对应参数，即可能是自定义脚本如“[now-1]”
    
    '执行脚本代码
    GetParameterValue = RunScripting(strParameterName)
End Function

Private Function RunScripting(ByVal strScript As String) As String
'执行vbs脚本
    Dim strFormatScript As String

    strFormatScript = Replace(Replace(strScript, "[", ""), "]", "")

On Error GoTo errHandle
    RunScripting = sctExecute.Eval(strFormatScript)
    Exit Function
errHandle:
    strFormatScript = "function return()" & vbCrLf & strFormatScript & " end function"
    Call sctExecute.AddCode(strFormatScript)
    
    RunScripting = sctExecute.Run("return")
End Function


Private Function IsParameterFormat(ByVal strData As String) As Boolean
'判断数据是否为参数数据
    IsParameterFormat = False
    
    If strData = "" Then Exit Function
    If Left(strData, 1) <> "[" Or Right(strData, 1) <> "]" Then Exit Function
    
    IsParameterFormat = True
End Function

Private Sub CreateInputControl(ByVal strName As String, ByVal lngInputType As Long, _
    ByVal strDefault As String, ByVal lngOrder As Long)
'创建录入组件
    
    Select Case lngInputType
        Case 0
            '创建文本框组件
            Load txtObj(lngOrder)
            
'            Call SetControlValue(lngInputType, lngOrder, strDefault)
            txtObj(lngOrder).Tag = "[" & strName & "]"
            
            txtObj(lngOrder).Left = 1950
            
            If mobjLastControl Is Nothing Then
                txtObj(lngOrder).Top = 1080 '315
            Else
                txtObj(lngOrder).Top = mobjLastControl.Top + mobjLastControl.Height + 120
            End If
            
            Set mobjLastControl = txtObj(lngOrder)
            
        Case 1, 2, 3
            '创建日期框组件
            Load dtpObj(lngOrder)
                        
            dtpObj(lngOrder).Format = dtpCustom
            dtpObj(lngOrder).CustomFormat = IIf(lngInputType = 1, "yyyy-MM-dd", IIf(lngInputType = 2, "HH:mm", "yyyy-MM-dd HH:mm"))
            
'            Call SetControlValue(lngInputType, lngOrder, strDefault)
            dtpObj(lngOrder).Tag = "[" & strName & "]"
            
            dtpObj(lngOrder).Left = 1950
            
            If mobjLastControl Is Nothing Then
                dtpObj(lngOrder).Top = 1080 '315
            Else
                dtpObj(lngOrder).Top = mobjLastControl.Top + mobjLastControl.Height + 120
            End If
            
            Set mobjLastControl = dtpObj(lngOrder)
            
        Case 4
            '创建下拉框
            Load cbxObj(lngOrder)
            
'            Call SetControlValue(lngInputType, lngOrder, strDefault)
            cbxObj(lngOrder).Tag = "[" & strName & "]"
            
            cbxObj(lngOrder).Left = 1950
            
            If mobjLastControl Is Nothing Then
                cbxObj(lngOrder).Top = 1080 '315
            Else
                cbxObj(lngOrder).Top = mobjLastControl.Top + mobjLastControl.Height + 120
            End If
            
            Set mobjLastControl = cbxObj(lngOrder)
        Case 5
            '创建可多选的列表框
            Load lstObj(lngOrder)
            
            lstObj(lngOrder).Tag = "[" & strName & "]"
            
            lstObj(lngOrder).Left = 1950
            
            If mobjLastControl Is Nothing Then
                lstObj(lngOrder).Top = 1080 '315
            Else
                lstObj(lngOrder).Top = mobjLastControl.Top + mobjLastControl.Height + 120
            End If
            
            Set mobjLastControl = lstObj(lngOrder)
    End Select
    
    mobjLastControl.Visible = True
    
    '创建Label数据
    Load labObj(lngOrder)
    
    labObj(lngOrder).Caption = strName
    labObj(lngOrder).Left = mobjLastControl.Left - labObj(lngOrder).Width - 120
    labObj(lngOrder).Top = mobjLastControl.Top + 60
    labObj(lngOrder).Visible = True
End Sub


