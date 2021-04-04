VERSION 5.00
Begin VB.Form frmExternalAllocationData 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "数据源提取编辑"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmd验证 
      Caption         =   "验证(&V)"
      Height          =   350
      Left            =   120
      TabIndex        =   6
      Top             =   3960
      Width           =   1100
   End
   Begin VB.CommandButton cmd取消 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4080
      TabIndex        =   5
      Top             =   3960
      Width           =   1100
   End
   Begin VB.CommandButton cmd确定 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2760
      TabIndex        =   4
      Top             =   3960
      Width           =   1100
   End
   Begin VB.TextBox txtEdit 
      Appearance      =   0  'Flat
      Height          =   2895
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   960
      Width           =   5055
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "3.[部门ID]传值为病区或科室列表的值。"
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   705
      Width           =   3240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "2.[就诊ID]门诊病人传值为就诊ID,住院传值为主页ID。"
      Height          =   180
      Left            =   120
      TabIndex        =   1
      Top             =   500
      Width           =   4410
   End
   Begin VB.Label lblTip 
      Caption         =   "1.SQL中的参数格式为固定的[参数名]，参数名按程序固定预制的[病人ID],[就诊ID],[部门ID],[医嘱ID]。"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "frmExternalAllocationData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrReturn As String

Public Function ShowMe(ByVal frmMain As Form, ByVal strSQLText As String) As String
    mstrReturn = strSQLText
    Me.Show 1, frmMain
    
    ShowMe = mstrReturn
End Function

Private Sub cmd取消_Click()
    Unload Me
End Sub

Private Sub cmd确定_Click()
    '功能：返回结果到主界面

    mstrReturn = txtEdit.Text
    
    Unload Me
End Sub

Private Function TrueObject(ByVal strObject As String) As String
    '功能：SQLObject函数的子函数,用于去除对象名中的无用字符
    Dim i As Integer
    '寻找第一个正常字符位置
    For i = 1 To Len(strObject)
        If InStr(Chr(32) & Chr(13) & Chr(10) & Chr(9), Mid(strObject, i, 1)) = 0 Then Exit For
    Next
    strObject = Mid(strObject, i)
    '寻找后面第一个非正常字符
    For i = 1 To Len(strObject)
        If InStr(Chr(32) & Chr(13) & Chr(10) & Chr(9), Mid(strObject, i, 1)) > 0 Then Exit For
    Next
    If i <= Len(strObject) Then strObject = Left(strObject, i - 1)
    TrueObject = strObject
End Function

Private Function TransSpecialChar(ByRef strSql As String, Optional ByVal blnRestore As Boolean = False) As Boolean
    '功能：转换SQL中的特殊字符；如：[]字符，避免与参数的符号冲突
    '返回：True成功；False失败

    Const STR_ORIGINAL As String = "[|]|(|)"
    Const STR_TRANS As String = "<左中括号>|<右中括号>|<左括号>|<右括号>"

    Dim strResult As String, strTmp As String
    Dim arrOriginal As Variant, arrTrans As Variant
    Dim i As Long, j As Long, lngBegin As Long
    Dim intLen As Integer
    
    If Trim(strSql = "") Then Exit Function
    
    On Error GoTo hErr
    
    strResult = strSql
    If blnRestore Then
        '还原
        arrOriginal = Split(STR_TRANS, "|")
        arrTrans = Split(STR_ORIGINAL, "|")
    Else
        '转换
        arrOriginal = Split(STR_ORIGINAL, "|")
        arrTrans = Split(STR_TRANS, "|")
    End If
    
    '检查SQL字符里是否存在[]字符
    i = 1
    lngBegin = 0
    Do While Mid(strResult, i) Like "*'*"
        If Mid(strResult, i, 1) = "'" Then
            If lngBegin <= 0 Then
                '开始
                lngBegin = i
            Else
                '结束
                lngBegin = 0
            End If
        Else
            If lngBegin > 0 Then
                '查找''字符内参数的特殊字符，即：SQL语句的字符串
                strTmp = Mid(strResult, lngBegin + 1)
                If InStr(strTmp, "'") > 0 Then
                    strTmp = Left(strTmp, InStr(strTmp, "'") - 1)
                    strTmp = Replace(strTmp, arrTrans(0), arrOriginal(0))
                Else
                    strTmp = ""
                End If
                
                If Not (strTmp Like "*[[][0-9][]]*" Or strTmp Like "*[[][0-9][0-9][]]*") Then
                    For j = LBound(arrOriginal) To UBound(arrOriginal)
                        intLen = Len(arrOriginal(j))
                        If Mid(strResult, i, intLen) = arrOriginal(j) Then
                            strResult = Left(strResult, i - 1) & arrTrans(j) & Mid(strResult, i + intLen)
                        End If
                    Next
                End If
            End If
        End If
        
        i = i + 1
    Loop
    
    strSql = strResult
    TransSpecialChar = True
    Exit Function
    
hErr:
End Function

Private Function GetWithAsTables(ByVal strSql As String) As String
    '功能：获取With as 之间的表名串，以逗号分隔
    Dim lngL As Long, lngR As Long, lngS As Long, strTabs As String
    Dim strTmp As String, blnFirst As Boolean
        
    strSql = Replace(strSql, vbCrLf, " ")
    strSql = Replace(strSql, vbTab, " ")
    strSql = Replace(strSql, "  ", " ")
    strSql = Replace(strSql, "  ", " ")
    strSql = Replace(strSql, "AS (", "AS(")
    
    lngL = InStr(1, strSql, "WITH")
    If lngL = 0 Then
        Exit Function
    Else
        lngL = lngL + 4
        blnFirst = True
    End If
        
    Do
        lngR = InStr(lngL, strSql, " AS(")
        If lngR = 0 Then
            Exit Do
        Else
            If Not blnFirst Then
                lngL = InStrRev(strSql, ",", lngR) + 1
            End If
            
            strTmp = Trim(Mid(strSql, lngL, lngR - lngL))
            '11G R2支持，例如：with T（column alias 1,column alias 2,......）
            lngS = InStr(strTmp, "(")
            If lngS > 1 Then
                strTmp = Mid(strTmp, 1, strTmp - 1)
            End If
            
            strTabs = strTabs & "," & strTmp
        End If
        
        blnFirst = False
        lngL = lngR + Len(" AS(")
    Loop
    GetWithAsTables = Mid(strTabs, 2)
End Function

Private Function TrimChar(Str As String) As String
    '功能:去除字符串中连续的空格和回车(含两头的空格,回车),不去除TAB字符,哪怕是连续的
    Dim strTmp As String
    
    If Trim(Str) = "" Then TrimChar = "": Exit Function
    
    strTmp = Trim(Str)
    
    strTmp = Replace(strTmp, "  ", " ")
    strTmp = Replace(strTmp, "  ", " ")

    strTmp = Replace(strTmp, vbCrLf & vbCrLf, vbCrLf)
    strTmp = Replace(strTmp, vbCrLf & vbCrLf, vbCrLf)
    
    If Left(strTmp, 2) = vbCrLf Then strTmp = Mid(strTmp, 3)
    If Right(strTmp, 2) = vbCrLf Then strTmp = Mid(strTmp, 1, Len(strTmp) - 2)
    TrimChar = strTmp
End Function

Private Function SQLObject(ByVal strSql As String, Optional ByVal strWithas As String) As String
'功能：分析SQL语句所用到的对象名
'参数：strSQL=要分析的原始SQL语句
'返回：SQL语句所访问到的对象名,如"部门表,病人费用记录,ZLHIS.人员表"
'说明：1.与Oracle SELECT语句兼容
'      2.如果SQL语句中的对象名前加有所有者前缀,则该前缀不会被截取
'      3.需要函数TrimChar;TrueObject的支持
    Dim intB As Long, intE As Long, intL As Long, intR As Long
    Dim strAnal As String, strSub As String, strObject As String
    Dim arrFrom() As String, strCur As String, strMulti As String, strTrue As String
    Dim i As Long, j As Long
    Dim lngTmp As Long
    Dim strTmp As String, strObjectSub As String
    
    On Error GoTo errH
    
    '大写化及去除多余的字符
    strAnal = UCase(TrimChar(strSql))
    If strWithas = "" Then
        strWithas = GetWithAsTables(strAnal)
    End If
    
    If InStr(strAnal, "SELECT") = 0 Or InStr(strAnal, "FROM") = 0 Then Exit Function
    If TransSpecialChar(strAnal) = False Then Exit Function
    
    '先分解处理嵌套子查询
    Do While InStr(strAnal, "(") > 0
        intB = InStr(strAnal, "("): intE = intB '匹配的左右括号位置
        intL = 1: intR = 0
        For i = intB + 1 To Len(strAnal)
            If Mid(strAnal, i, 1) = "(" Then
                intL = intL + 1
            ElseIf Mid(strAnal, i, 1) = ")" Then
                intR = intR + 1
            End If
            If intL = intR Then
                intE = i
                strTmp = Mid(strAnal, 1, intB - 1)
                lngTmp = 0
                If InStrRev(strTmp, " TABLE") > 0 Or InStrRev(strTmp, " TABLE ") > 0 Then
                    lngTmp = IIf(InStrRev(strTmp, " TABLE ") > 0, InStrRev(strTmp, " TABLE "), InStrRev(strTmp, " TABLE"))
                    strTmp = Mid(strTmp, lngTmp + 6)
                    strTmp = Trim(strTmp)
                End If
                If intE - intB - 1 <= 0 Then
                    '对于非子查询,将括号换成其它符号,以使循环继续
                    strAnal = Left(strAnal, intB - 1) & "@" & Mid(strAnal, intB + 1)
                    strAnal = Left(strAnal, intE - 1) & "@" & Mid(strAnal, intE + 1)
                ElseIf InStr(Mid(strAnal, intB + 1, intE - intB - 1), "SELECT") > 0 _
                    And InStr(Mid(strAnal, intB + 1, intE - intB - 1), "FROM") > 0 Then
                    '子查询语句
                    strSub = Mid(strAnal, intB + 1, intE - intB - 1)
                    '将该子查询部份作为为特殊对象名
                    strAnal = Replace(strAnal, Mid(strAnal, intB, intE - intB + 1), "嵌套查询")
                    '递归分析
                    strObjectSub = SQLObject(strSub, strWithas)
                    If InStr(strObject & "," & strWithas & ",", "," & strObjectSub & ",") = 0 Then
                        strObject = strObject & "," & strObjectSub
                    End If
                ElseIf strTmp = "" And lngTmp <> 0 Then
                    '去除Table动态内存表
                    strAnal = Replace(strAnal, Mid(strAnal, lngTmp + 1, intE - lngTmp + 1 + 1), "动态内存表")
                Else
                    strAnal = Left(strAnal, intB - 1) & "@" & Mid(strAnal, intB + 1)
                    strAnal = Left(strAnal, intE - 1) & "@" & Mid(strAnal, intE + 1)
                End If
                Exit For
            End If
        Next
        '无匹配右括号
        If intE = intB Then strAnal = Left(strAnal, intB - 1) & "@" & Mid(strAnal, intB + 1)
    Loop
    
    '分解分析(此时strAnal为简单查询,可能带Union等连接)
    arrFrom = Split(strAnal, "FROM")
    For i = 1 To UBound(arrFrom) '从第一个From后面部份开始
        strCur = arrFrom(i)
        If InStr(strCur, "WHERE") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "WHERE") - 1)
        ElseIf InStr(strCur, "START WITH") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "START WITH") - 1)
        ElseIf InStr(strCur, "CONNECT BY") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "CONNECT BY") - 1)
        ElseIf InStr(strCur, "GROUP") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "GROUP") - 1)
        ElseIf InStr(strCur, "HAVING") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "HAVING") - 1)
        ElseIf InStr(strCur, "ORDER") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "ORDER") - 1)
        ElseIf InStr(strCur, "UNION") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "UNION") - 1)
        ElseIf InStr(strCur, "MINUS") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "MINUS") - 1)
        ElseIf InStr(strCur, "INTERSECT") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "INTERSECT") - 1)
        Else
            strMulti = strCur
        End If
        For j = 0 To UBound(Split(strMulti, ","))
            strTrue = TrueObject(Split(strMulti, ",")(j))
            If InStr(strObject & "," & strWithas & ",", "," & strTrue & ",") = 0 And strTrue <> "嵌套查询" And strTrue <> "动态内存表" Then
                If InStr(strTrue, "'") = 0 And InStr(strTrue, "@") = 0 Then
                    strObject = strObject & "," & strTrue
                End If
            End If
        Next
    Next
    '完成
    SQLObject = Mid(strObject, 2)
    SQLObject = Replace(SQLObject, ",,", ",")
    Exit Function
errH:
    err.Clear
End Function

Private Sub cmd验证_Click()
    '功能：检验SQL的正确性
    Dim strSql As String
    Dim strObject As String
    Dim rsData As ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    strSql = txtEdit.Text
    
    'SQL检查
    '-----------------------------------------------
    strObject = SQLObject(strSql)
    If strObject = "" And InStr(UCase(strSql), "TABLE") = 0 And InStr(UCase(strSql), "@") = 0 Then
        MsgBox "不能分析SQL语句所查询的数据对象,请检查是否正确书写！", vbInformation, App.Title
        Exit Sub
    End If
    '-----------------------------------------------
    
    'SQL执行
    '-----------------------------------------------
    strSql = UCase(strSql)
    
    strSql = Replace(strSql, "[病人ID]", "0")
    strSql = Replace(strSql, "[就诊ID]", "0")
    strSql = Replace(strSql, "[部门ID]", "0")
    strSql = Replace(strSql, "[医嘱ID]", "0")
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    '-----------------------------------------------
    
    cmd确定.Enabled = True
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    txtEdit.Text = mstrReturn
    cmd确定.Enabled = False
End Sub

Private Sub txtEdit_Change()
    cmd确定.Enabled = False
End Sub
