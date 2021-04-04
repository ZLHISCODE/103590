VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFunction"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Sub About(Optional frmParent As Object)
'功能：显示关于窗口
    Call ShowAbout(frmParent)
End Sub

Public Function GetFunScript(ByVal cnLink As ADODB.Connection, ByVal strExp As String) As String
'功能：获取指定函数的描述
'参数：strFunc="ZLHIS.ZL_FUN_XXXX(X,X,X,X,)"
'返回："函数名:XXX~中文名:XXX~说明:XXX||参数名:XXX~中文名:XXX~值:XXX|参数名:XXX~....."
    Dim strSQL As String, strScript As String
    Dim rsTmp As New ADODB.Recordset

    Dim lngSys As Long, strOwner As String, strFunc As String
    Dim strVal As String, strVals As String
    Dim objPars As FuncPars, tmpPar As FuncPar
    Dim i As Integer, j As Integer, strTmp As String
    
    If cnLink Is Nothing Then Exit Function
    If cnLink.State = adStateClosed Then Exit Function
    Set gcnOracle = cnLink
    gstrDBUser = GetDBUser
    
    Call SplitFunc(strExp, strOwner, strFunc, strVals)
    
    On Error GoTo errH
    
    strSQL = "Select * From zlFunctions Where Upper(函数名)='" & UCase(strFunc) & "' And 系统 IN(Select 编号 From zlSystems Where Upper(所有者)='" & UCase(strOwner) & "')"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "获取指定函数的描述")
    If rsTmp.EOF Then Exit Function
    lngSys = rsTmp!系统
    
    Set objPars = ReadFuncPars(lngSys, rsTmp!函数号)
    For i = 1 To objPars.Count
        objPars(i).分类SQL = Replace(objPars(i).分类SQL, "[系统]", lngSys)
        objPars(i).明细SQL = Replace(objPars(i).明细SQL, "[系统]", lngSys)
    Next
        
    i = 0: strScript = ""
    For Each tmpPar In objPars
        '当前值
        strVal = ""
        If UBound(Split(strVals, "|")) >= i Then
            strVal = Split(strVals, "|")(i)
            If Left(strVal, 1) = "'" And Right(strVal, 1) = "'" Then
                strVal = Mid(strVal, 2, Len(strVal) - 2)
            End If
        Else
            Exit For
        End If
        
        i = i + 1
        
        strScript = strScript & "|参数名:" & tmpPar.名称 & "~中文名:" & tmpPar.中文名 & "~值:"
        
        If UCase(strVal) = "[ZLBEGINTIME]" Or UCase(strVal) = "[ZLENDTIME]" Then
            strScript = strScript & "动态确定"
        Else
            If tmpPar.缺省值 = "固定值列表…" Then
                For j = 0 To UBound(Split(tmpPar.值列表, "|"))
                    '绑定值相同则赋值
                    If strVal = Split(Split(tmpPar.值列表, "|")(j), ",")(1) And strVal <> "" Then
                        strScript = strScript & Split(Split(tmpPar.值列表, "|")(j), ",")(0)
                        Exit For
                    End If
                Next
            ElseIf tmpPar.缺省值 = "选择器定义…" Then
                '根据绑定值赋值
                If strVal <> "" Then
                    strTmp = ""
                    If tmpPar.值列表 Like "*|*" Then
                        strTmp = Split(tmpPar.值列表, "|")(1)
                        If (strVal = strTmp) Or (UCase(strVal) = "NULL" And Trim(strTmp) = "") Then
                            strScript = strScript & Split(tmpPar.值列表, "|")(0)
                            strTmp = "OK"
                        Else
                            strTmp = ""
                        End If
                    End If
                    
                    If strTmp = "" Then
                        If InStr(tmpPar.对象, "|") > 0 Then strTmp = Split(tmpPar.对象, "|")(0)
                        strTmp = SQLOwner(Replace(RemoveNote(tmpPar.明细SQL), "[*]", ""), strTmp)
                        strTmp = GetBalndValue(strTmp, tmpPar.明细字段, strVal)
                        If strTmp <> "" Then strScript = strScript & Split(strTmp, "|")(0)
                    End If
                End If
            Else
                If tmpPar.类型 = 2 Then
                    '非宏日期才赋值
                    If Left(tmpPar.缺省值, 1) <> "&" And strVal <> "" Then
                        If UCase(strVal) Like "TO_DATE('*','*')" Then
                            strScript = strScript & Format(GetDate(strVal), "yyyy-MM-dd")
                        End If
                    End If
                Else
                    '赋值
                    If strVal <> "" Then strScript = strScript & strVal
                End If
            End If
        End If
    Next
    
    strScript = "函数名:" & rsTmp!函数名 & "~中文名:" & rsTmp!中文名 & "~说明:" & IIf(IsNull(rsTmp!说明), "", rsTmp!说明) & "||" & Mid(strScript, 2)
    GetFunScript = strScript
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function FuncGuide(ByVal cnMain As ADODB.Connection, frmParent As Object, Optional strExp As String) As String
'功能：选择函数并获取公式
'参数：
'     cnMain:主程序的数据库链接
'     frmParent:主窗体
'返回：选择函数的调用公式
    If cnMain Is Nothing Then Exit Function
    If cnMain.State = adStateClosed Then Exit Function
    Set gcnOracle = cnMain
    '49074
    InitCommon gcnOracle
    gstrDBUser = GetDBUser
    Set gfrmMain = frmParent
    
    gblnDBA = HaveDBA()
    
    If grsObject Is Nothing Then Set grsObject = UserObject
    If grsObject Is Nothing Then Exit Function
    If grsObject.State = adStateClosed Then Set grsObject = UserObject
       
    frmParInput.mstrExp = strExp
    Load frmParInput
    
    Err.Clear
    On Error Resume Next
    frmParInput.Show 1, frmParent
    If Err.Number <> 0 Then
        Err.Clear
        frmParInput.Show 1
    End If
    If gblnOK Then
        FuncGuide = frmParInput.mstrExp
        frmParInput.mstrExp = ""
    End If
End Function

Public Function FuncManage(ByVal cnMain As ADODB.Connection, Optional frmParent As Object) As Boolean
'功能：函数管理
'参数：
'     cnMain:主程序的数据库链接
'     frmParent:主窗体
'返回：调用是否成功
    If cnMain Is Nothing Then Exit Function
    If cnMain.State = adStateClosed Then Exit Function
    Set gcnOracle = cnMain
    '49074
    InitCommon gcnOracle
    gstrDBUser = GetDBUser
    Set gfrmMain = frmParent
    '49074
    InitCommon gcnOracle
    gblnDBA = HaveDBA()
    
    If grsObject Is Nothing Then Set grsObject = UserObject
    If grsObject Is Nothing Then Exit Function
    If grsObject.State = adStateClosed Then Set grsObject = UserObject
    
    If Not frmParent Is Nothing Then
        Load frmMain
        Err.Clear
        On Error Resume Next
        frmMain.Show , frmParent
        If Err.Number <> 0 Then
            Err.Clear
            frmMain.Show 1
        End If
    Else
        frmMain.Show
    End If
    
    FuncManage = True
End Function

Private Sub Class_Initialize()
    Call RegFuncFile
End Sub

Private Sub Class_Terminate()
    Set gcnOracle = Nothing
    Set grsObject = Nothing
End Sub

Public Function CloseWindows() As Boolean
'功能:关闭所有子窗口
    Dim frmThis As Form
    
    On Error Resume Next
    
    For Each frmThis In Forms
        Unload frmThis
    Next
    CloseWindows = (Forms.Count = 0)
End Function

