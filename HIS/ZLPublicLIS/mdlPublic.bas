Attribute VB_Name = "mdlpublic"
Option Explicit

Public Function GetLisSample() As ADODB.Recordset
    '获取检验标本
    Dim strSql  As String
    If gblnNewLis Then
        strSql = "select A.名称, B.顺序 from 检验标本类型 A, 标本顺序 B where A.名称 = B.名称(+) order by B.顺序"
    Else
        strSql = "select A.名称, B.顺序 from 诊疗检验标本 A, 标本顺序 B where A.名称 = B.名称(+) order by B.顺序"
    End If
    Set GetLisSample = gobjDatabase.OpenSQLRecord(strSql, gstrSysName)
End Function

Public Function GetLisType() As ADODB.Recordset
    '获取检验类别
    Dim strSql  As String
    If gblnNewLis Then
        strSql = "select A.分类 As 名称, B.顺序 from (select distinct 分类 from 检验组合项目) A, 类别顺序 B where A.分类 = B.名称(+) order by B.顺序"
    Else
        strSql = "select A.名称, B.顺序 from 诊疗检验类型 A, 类别顺序 B where A.名称 = B.名称(+) order by B.顺序"
    End If
    Set GetLisType = gobjDatabase.OpenSQLRecord(strSql, gstrSysName)
End Function

Public Function GetLisName() As ADODB.Recordset
    '获取检验名称
    Dim strSql  As String
    If gblnNewLis Then
        strSql = "select A.名称, B.顺序" & vbNewLine & _
            "  from 检验组合项目 A, 项目顺序 B" & vbNewLine & _
            " where A.名称 = B.名称(+)" & vbNewLine & _
            "   And (A.停用日期 is null or A.停用日期 > sysdate)" & vbNewLine & _
            " order by B.顺序"
            Set GetLisName = gobjPublicHisCommLis.openSqlOtherDB(1, strSql, gstrSysName)
    Else
        strSql = "select Decode(Instr(A.名称, '('),0,A.名称,substr(A.名称, 1, Instr(A.名称, '(') - 1)) As 名称, B.顺序" & vbNewLine & _
            "  from 诊疗项目目录 A, 项目顺序 B" & vbNewLine & _
            " where A.名称 = B.名称(+)" & vbNewLine & _
            "   And A.类别 = 'C'" & vbNewLine & _
            "   And A.单独应用 = 1" & vbNewLine & _
            "   And (A.撤档时间 is null or A.撤档时间 > sysdate)" & vbNewLine & _
            " order by B.顺序"
         Set GetLisName = gobjDatabase.OpenSQLRecord(strSql, gstrSysName)
    End If
   
    
End Function

Public Function JustType() As Integer
    Dim strSql  As String
    Dim rsType  As ADODB.Recordset
    Dim rsSamp  As ADODB.Recordset
    Dim rsName  As ADODB.Recordset

    strSql = "select 顺序,count(*) As 数量 from 类别顺序 group by 顺序  Order By 顺序"
    Set rsType = gobjDatabase.OpenSQLRecord(strSql, gstrSysName)
    strSql = "select 顺序,count(*) As 数量 from 标本顺序 group by 顺序  Order By 顺序 "
    Set rsSamp = gobjDatabase.OpenSQLRecord(strSql, gstrSysName)
    strSql = "select 顺序,count(*) As 数量 from 项目顺序 group by 顺序  Order By 顺序"
    Set rsName = gobjDatabase.OpenSQLRecord(strSql, gstrSysName)
    If rsType.RecordCount > 1 Then
        JustType = 0
    ElseIf rsSamp.RecordCount > 1 Then
        JustType = 1
    ElseIf rsType.RecordCount > 1 And rsSamp.RecordCount > 1 Then
        JustType = 2
    ElseIf rsName.RecordCount > 1 Then
        JustType = 3
    Else
        JustType = 0
    End If

'    If JustType < 0 Then JustType = 0
End Function

Public Function CopyRecordStruct(ByVal rsFrom As ADODB.Recordset, Optional ByVal blnRowID As Boolean = False, Optional ByVal blnNotOpen As Boolean = False) As ADODB.Recordset
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************

    Dim lngLoop As Long
    Dim rs As ADODB.Recordset

    On Error GoTo errHand

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.LockType = adLockBatchOptimistic
    rs.CursorType = adOpenStatic

    For lngLoop = 0 To rsFrom.Fields.count - 1

        Select Case rsFrom.Fields(lngLoop).type
        Case 135            'Oracle的Date型
            rs.Fields.Append rsFrom.Fields(lngLoop).Name, adVarChar, 100, rsFrom.Fields(lngLoop).Attributes
        Case Else
            rs.Fields.Append rsFrom.Fields(lngLoop).Name, adVarChar, rsFrom.Fields(lngLoop).DefinedSize + 100
        End Select

    Next
    If blnRowID Then
        rs.Fields.Append "行号", adVarChar, 30
    End If

    If blnNotOpen = False Then rs.Open

    Set CopyRecordStruct = rs

    Exit Function
errHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function CopyRecordData(ByVal rsFrom As ADODB.Recordset, ByRef rsTo As ADODB.Recordset, Optional blnAll As Boolean = True) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strTmp As String
    Dim lngLoop As Long

    On Error GoTo errHand

    If blnAll Then
        If rsFrom.RecordCount > 0 Then rsFrom.MoveFirst
    End If

    Do While Not rsFrom.EOF
        rsTo.AddNew
        For lngLoop = 0 To rsFrom.Fields.count - 1

            On Error Resume Next
            strTmp = ""
            strTmp = rsTo.Fields(rsFrom.Fields(lngLoop).Name).Name
            On Error GoTo errHand

            If UCase(strTmp) = UCase(rsFrom.Fields(lngLoop).Name) Then
                rsTo.Fields(strTmp).Value = Trim(Nvl(rsFrom.Fields(lngLoop).Value))
            End If

        Next
        If blnAll = False Then Exit Do
        rsFrom.MoveNext
    Loop

    CopyRecordData = True

    Exit Function

errHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function



