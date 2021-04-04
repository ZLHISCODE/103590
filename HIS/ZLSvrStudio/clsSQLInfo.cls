VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSQLInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'保持属性值的局部变量
Private mvarSQL As String '局部复制
Private mvarTip As String '局部复制
Private mvarBlock As Boolean '局部复制
Private mvarBlockType As String '局部复制
Private mvarBlockName As String '局部复制
Private mvarFile As String '局部复制
Private mvarFileLine As Long '局部复制
Private mvarFileLines As Long '局部复制
Private mvarBlockWrap As Boolean '局部复制
Private mvarLobDDL As Boolean '局部复制
'-----------------------------------------------------------------------
'以下属性用以应对超长的块类型SQL(此种SQL在字符串替换等处理时会产生性能问题)
'------------------------------------------------------------------------
'判断SQLInfo中是否有SQL
Private mvarHaveSQL   As Boolean '局部复制
'有效SQL的前面五行，一般可以标识SQL特征
Private mvarPartSQL As String '局部复制
Private mvarForeign As Boolean '是否是外键
Private mvarLackOnline As Boolean '是否可以增加Online关键字
Private mblnInit As Boolean

Public Sub ClearSQL()
    mblnInit = True
    Me.SQL = ""
    Me.Tip = ""
    Me.Block = False
    Me.BlockWrap = False
    Me.BlockType = ""
    Me.BlockName = ""
    Me.FileLine = 0
    Me.HaveSQL = False
    Me.PartSQL = ""
    Me.Foreign = False
    Me.LackOnline = False
    Me.LobDDL = False
    mblnInit = False
End Sub

Public Function CopyMe() As clsSQLInfo
    Dim objMe As New clsSQLInfo
    mblnInit = True
    objMe.SQL = Me.SQL
    objMe.Tip = Me.Tip
    objMe.Block = Me.Block
    objMe.BlockWrap = Me.BlockWrap
    objMe.BlockType = Me.BlockType
    objMe.BlockName = Me.BlockName
    objMe.File = Me.File
    objMe.FileLine = Me.FileLine
    objMe.FileLines = Me.FileLines
    objMe.HaveSQL = Me.HaveSQL
    objMe.PartSQL = Me.PartSQL
    objMe.Foreign = Me.Foreign
    objMe.LackOnline = Me.LackOnline
    objMe.LobDDL = mvarLobDDL
    mblnInit = False
    Set CopyMe = objMe
End Function

Public Sub CopySQL(ByVal objSQL As clsSQLInfo)
    Me.SQL = objSQL.SQL
    Me.Tip = objSQL.Tip
    Me.Block = objSQL.Block
    Me.BlockWrap = objSQL.BlockWrap
    Me.BlockType = objSQL.BlockType
    Me.BlockName = objSQL.BlockName
    Me.File = objSQL.File
    Me.FileLine = objSQL.FileLine
    Me.FileLines = objSQL.FileLines
    Me.HaveSQL = objSQL.HaveSQL
    Me.PartSQL = objSQL.PartSQL
    Me.Foreign = objSQL.Foreign
    Me.LackOnline = objSQL.LackOnline
End Sub

Public Function IsSameTo(ByVal objSQL As clsSQLInfo) As Boolean
'功能：当前SQL对象与另一个SQL对比，判断两者是否是同一个对象，返回对比结果
    If objSQL Is Nothing Then Exit Function
    If objSQL.SQL <> Me.SQL Then Exit Function
    If objSQL.Tip <> Me.Tip Then Exit Function
    If objSQL.Block <> Me.Block Then Exit Function
    If objSQL.BlockWrap <> Me.BlockWrap Then Exit Function
    If objSQL.BlockType <> Me.BlockType Then Exit Function
    If objSQL.BlockName <> Me.BlockName Then Exit Function
    If objSQL.File <> Me.File Then Exit Function
    If objSQL.FileLine <> Me.FileLine Then Exit Function
    If objSQL.FileLines <> Me.FileLines Then Exit Function
    If objSQL.HaveSQL <> Me.HaveSQL Then Exit Function
    If objSQL.PartSQL <> Me.PartSQL Then Exit Function
    If objSQL.Foreign <> Me.Foreign Then Exit Function
    If objSQL.LackOnline <> Me.LackOnline Then Exit Function
    IsSameTo = True
End Function

Public Property Let BlockWrap(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.BlockWrap = 5
    mvarBlockWrap = vData
End Property

Public Property Get BlockWrap() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.BlockWrap
    BlockWrap = mvarBlockWrap
End Property

Public Property Let FileLines(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.FileLines = 5
    mvarFileLines = vData
End Property

Public Property Get FileLines() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.FileLines
    FileLines = mvarFileLines
End Property

Public Property Let FileLine(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.FileLine = 5
    mvarFileLine = vData
End Property

Public Property Get FileLine() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.FileLine
    FileLine = mvarFileLine
End Property

Public Property Let File(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.File = 5
    mvarFile = vData
End Property

Public Property Get File() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.File
    File = mvarFile
End Property

Public Property Let BlockName(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.BlockName = 5
    mvarBlockName = vData
End Property

Public Property Get BlockName() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.BlockName
    BlockName = mvarBlockName
End Property

Public Property Let BlockType(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.BlockType = 5
    mvarBlockType = vData
End Property

Public Property Get BlockType() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.BlockType
    BlockType = mvarBlockType
End Property

Public Property Let Block(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Block = 5
    mvarBlock = vData
End Property

Public Property Get Block() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Block
    Block = mvarBlock
End Property

Public Property Let Tip(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Tip = 5
    mvarTip = vData
End Property

Public Property Get Tip() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Tip
    Tip = mvarTip
End Property

Public Property Let SQL(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.SQL = 5
    mvarSQL = vData
End Property

Public Property Get SQL() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.SQL
    SQL = mvarSQL
End Property

Public Property Let HaveSQL(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.SQL = 5
    mvarHaveSQL = vData
End Property

Public Property Get HaveSQL() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.SQL
    HaveSQL = mvarHaveSQL
End Property

Public Property Let LobDDL(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.SQL = 5
    mvarLobDDL = vData
End Property

Public Property Get LobDDL() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.SQL
    LobDDL = mvarLobDDL
End Property

Public Property Let PartSQL(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.SQL = 5
    mvarPartSQL = vData
    If mblnInit Then Exit Property
    If Not Me.Block Then
        If Me.PartSQL Like "ALTER *" Then
            If Me.PartSQL Like "ALTER TABLE * ADD CONSTRAINT * FOREIGN KEY*" Then
                Me.Foreign = True
            Else
                If Not Me.PartSQL Like "* ONLINE*" Then '没有OnLIne关键字
                    If Me.PartSQL Like "ALTER TABLE * ADD CONSTRAINT * PRIMARY KEY* USING INDEX*" Then
                        Me.LackOnline = True
                    ElseIf Me.PartSQL Like "ALTER TABLE * ADD CONSTRAINT * UNIQUE* USING INDEX*" Then
                        Me.LackOnline = True
                    ElseIf Me.PartSQL Like "ALTER INDEX * REBUILD*" Then
                        Me.LackOnline = True
                    ElseIf Me.PartSQL Like "ALTER TABLE * MODIFY[( ]* [CB]LOB*" Then
                        Me.LobDDL = True
                    End If
                ElseIf Me.PartSQL Like "ALTER TABLE * MODIFY[( ]* [CB]LOB*" Then
                    Me.LobDDL = True
                End If
            End If
        ElseIf Me.PartSQL Like "CREATE INDEX *" Then
            If Not Me.PartSQL Like "* ONLINE*" Then '没有OnLIne关键字
                Me.LackOnline = True
            End If
        End If
    End If
End Property

Public Property Get PartSQL() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.SQL
    PartSQL = mvarPartSQL
End Property

Public Property Let Foreign(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.SQL = 5
    mvarForeign = vData
End Property

Public Property Get Foreign() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.SQL
    Foreign = mvarForeign
End Property

Public Property Let LackOnline(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.SQL = 5
    mvarLackOnline = vData
End Property

Public Property Get LackOnline() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.SQL
    LackOnline = mvarLackOnline
End Property

Public Function AnsySQL() As ADODB.Recordset
'功能：对存储过程、函数进行参数解析
    If Me.Block Then
        If Me.BlockType Like "*PROCEDURE*" Or Me.BlockType Like "*FUNCTION*" Then
            Set AnsySQL = AnsyProcedurePara(True)
        End If
    End If
End Function


Private Function AnsyProcedurePara(Optional ByVal blnTruncOwner As Boolean) As ADODB.Recordset
  '******************************************************************************************************************
    '功能：检查
    '参数：blnTruncOwner=是否去掉所有者
    '返回：成功返回true，否则返回false
    '******************************************************************************************************************
    Dim strFMTSQL As String, strType As String, strName As String, strProcPara As String, strReturn As String
    Dim strParName As String, strParDirection  As String, strParType As String, strParTypeMask As String, strParDefault As String
    Dim cllStrs As Collection
    Dim arrTmp As Variant, arrTmp1 As Variant
    Dim arrLine As Variant, rsReturn As ADODB.Recordset, intIndex As Integer, i As Long, j As Long
    '获取标准头
    strFMTSQL = GetFMTSQLStr(TrimCommentLossless(Me.SQL), cllStrs)
    If strFMTSQL = "" Then Exit Function
    '获取过程函数头
    strFMTSQL = Split(strFMTSQL, " PIPELINED AS ")(0)
    If strFMTSQL = "" Then Exit Function
    strFMTSQL = Split(strFMTSQL, " PIPELINED;")(0)
    If strFMTSQL = "" Then Exit Function
    strFMTSQL = Split(strFMTSQL, " AS ")(0)
    If strFMTSQL = "" Then Exit Function
    strFMTSQL = Split(strFMTSQL, " IS ")(0)
    If strFMTSQL = "" Then Exit Function
    arrTmp = Split(Replace(strFMTSQL, "(", ")"), ")")
    If arrTmp(0) Like "* FUNCTION *" Then
        strType = "FUNCTION"
    ElseIf arrTmp(0) Like "* PROCEDURE *" Then
        strType = "PROCEDURE"
    End If
    If strType = "" Then Exit Function
    arrTmp1 = Split(arrTmp(0), " " & strType & " ")
    If UBound(arrTmp1) < 1 Then Exit Function
    arrTmp1(1) = Trim(TrimEx(arrTmp1(1), True))
    strName = Split(arrTmp1(1), " ")(0)
    If UBound(arrTmp) = 0 Then '不存在参数
        If strType = "FUNCTION" Then
            strReturn = Trim(Split(arrTmp(0), " RETURN ")(1))
        End If
    Else
        If strType = "FUNCTION" And UBound(arrTmp) = 2 Then
            strReturn = Trim(Split(arrTmp(2), "RETURN ")(1))
        End If
        strProcPara = arrTmp(1)
    End If
    
    If blnTruncOwner Then
        If InStr(strName, ".") > 0 Then '带所有者前缀
            strName = Split(strName, ".")(1)
        End If
    End If
    '针对 IN \ OUT  表名 . 字段名 % Type : = 'A' 进行处理
    '针对 IN  OUT   IN \ OUT   默认 进行处理
    '优先将字符串常量解析出来，并用占位符占位,防止字符串存在语句分界符
    On Error GoTo errH
    Set rsReturn = CopyNewRec(Nothing, True, , _
                        Array("位置", adInteger, 3, 0, "名称", adVarChar, 50, Empty, _
                              "方向", adVarChar, 20, Empty, "类型描述", adVarChar, 100, Empty, _
                              "类型", adVarChar, 20, Empty, "默认值", adVarChar, 100, Empty))

    rsReturn.AddNew Array("位置", "名称", "方向", "类型描述", "类型", "默认值"), _
                    Array(-1, strName, "", "", strType, "")
    If strReturn <> "" Then
        rsReturn.AddNew Array("位置", "名称", "方向", "类型描述", "类型", "默认值"), _
                        Array(0, "返回值", "", "", strReturn, "")
    End If
    arrLine = Split(strProcPara, ","): intIndex = 0
    For i = LBound(arrLine) To UBound(arrLine)
        arrLine(i) = Trim(arrLine(i))
        If arrLine(i) <> "" Then
            '解析参数内容
            arrTmp = Split(Trim(arrLine(i)), " ")
            strParName = arrTmp(0): strParDirection = "": strParType = arrTmp(UBound(arrTmp)): strParDefault = "": strParTypeMask = ""
            For j = LBound(arrTmp) + 1 To UBound(arrTmp) - 1
                strParDirection = strParDirection & " " & arrTmp(j)
            Next
            strParDirection = Trim(strParDirection)
            If strParDirection = "IN OUT" Then
                strParDirection = "IN\OUT"
            ElseIf strParDirection = "" Then
                strParDirection = "IN"
            End If
            '处理类型:=A
            arrTmp = Split(strParType, ":=")
            strParType = arrTmp(0)
            If UBound(arrTmp) > 0 Then
                strParDefault = arrTmp(1)
                If strParDefault Like "[[]*[]]" Then
                    strParDefault = "'" & cllStrs(Mid(strParDefault, 2, Len(strParDefault) - 2)) & "'"
                End If
                If strParDefault = "''" Or strParDefault = "" Then strParDefault = "NULL"
            End If
            '处理表名.字段名%Type
            If strParType Like "*%TYPE" Then
                strParTypeMask = Mid(strParType, 1, Len(strParType) - 5)
                strParType = ""
            End If
            intIndex = intIndex + 1
            rsReturn.AddNew Array("位置", "名称", "方向", "类型描述", "类型", "默认值"), _
                            Array(intIndex, strParName, strParDirection, strParTypeMask, strParType, strParDefault)
        End If
    Next
    Set AnsyProcedurePara = rsReturn
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description
End Function

Public Function LoadSQL(ByVal strSQL As String, Optional ByVal strLineSplit As String = vbNewLine) As Boolean
'加载SQL,生成SQL对象
'包、包体、函数、存储过程、Type、匿名块的创建SQL以单行"/"做结束符。其余类型SQL以";"(分号)做语句结束符
    Dim objScript As New clsRunScript
    If objScript.AnalysisSQLString(strSQL, , strLineSplit) Then
        If Not objScript.EOF Then
            Call Me.CopySQL(objScript.SQLInfo)
            LoadSQL = True
        End If
    End If
End Function


