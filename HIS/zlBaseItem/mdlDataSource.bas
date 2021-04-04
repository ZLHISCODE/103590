Attribute VB_Name = "mdlDataSource"
Option Explicit

Private Const SQL_HANDLE_ENV As Integer = 1
Private Const SQL_HANDLE_DBC As Integer = 2
Private Const SQL_ATTR_ODBC_VERSION = 200
Private Const SQL_OV_ODBC3 = 3&
Private Const SQL_OV_ODBC2 = 2&

Private Const SQL_DRIVER_NOPROMPT As Long = 0
Private Const SQL_DRIVER_COMPLETE As Long = 1
Private Const SQL_DRIVER_PROMPT As Long = 2
Private Const SQL_DRIVER_COMPLETE_REQUIRED As Long = 3

' Options for SQLFetchScroll
Private Const SQL_FETCH_NEXT = 1
Private Const SQL_FETCH_FIRST = 2
Private Const SQL_FETCH_LAST = 3
Private Const SQL_FETCH_PRIOR = 4
Private Const SQL_FETCH_ABSOLUTE = 5
Private Const SQL_FETCH_RELATIVE = 6
Private Const SQL_FETCH_FIRST_USER = 31
Private Const SQL_FETCH_FIRST_SYSTEM = 32

'  RETCODEs
Private Const SQL_SUCCESS As Long = 0
Private Const SQL_SUCCESS_WITH_INFO As Long = 1
Private Const SQL_ERROR As Long = -1
Private Const SQL_INVALID_HANDLE As Long = -2
Private Const SQL_NO_DATA  As Long = 100

'CreateDataSource
Private Const ODBC_ADD_DSN As Long = 1
Private Const ODBC_CONFIG_DSN As Long = 2
Private Const ODBC_REMOVE_DSN As Long = 3
Private Const ODBC_ADD_SYS_DSN As Long = 4
Private Const ODBC_CONFIG_SYS_DSN As Long = 5
Private Const ODBC_REMOVE_SYS_DSN As Long = 6

Private Declare Function SQLConfigDataSource Lib "odbccp32" (ByVal hwnd As Long, ByVal lngRequest As Long, ByVal strDriver As String, ByVal strAttributes As String) As Boolean
Private Declare Function SQLCreateDataSource Lib "odbccp32" (ByVal hwnd As Long, ByVal strDSN As String) As Boolean
Private Declare Function SQLDataSources Lib "odbc32.dll" (ByVal hEnv As Long, ByVal lngDirect As Long, ByVal SourceName As String, ByVal lngNameBuff As Long, lngNameReturn As Long, ByVal Description As String, ByVal lngDescBuff As Long, lngDescReturn As Long) As Long

Private Declare Function SQLAllocHandle Lib "odbc32.dll" (ByVal iHandleType As Integer, ByVal lInputHandle As Long, lOutputHandlePtr As Long) As Integer
Private Declare Function SQLFreeHandle Lib "odbc32.dll" (ByVal iHandleType As Integer, ByVal lInputHandle As Long) As Integer
Private Declare Function SQLSetEnvAttr Lib "odbc32.dll" (ByVal hEnv As Long, ByVal lAttribute As Long, ByVal sValuePtr As String, ByVal lStringLength As Long) As Integer
Private Declare Function SQLSetEnvAttrLong Lib "odbc32.dll" Alias "SQLSetEnvAttr" (ByVal hEnv As Long, ByVal lAttribute As Long, ByVal lValue As Long, ByVal lStringLength As Long) As Integer
Private Declare Function SQLDriverConnect Lib "odbc32.dll" (ByVal hConnection As Long, ByVal hwnd As Long, ByVal sInConnectionString As String, ByVal iStringLength1 As Integer, ByVal sOutConnectionString As String, ByVal iBufferLength As Integer, iStringLength2Ptr As Integer, ByVal iDriverCompletion As Integer) As Integer
Private Declare Function SQLDisconnect Lib "odbc32.dll" (ByVal hConnection As Long) As Integer
Private Declare Function SQLGetInstalledDrivers Lib "odbccp32" (ByVal Drviers As String, ByVal BuffLen As Long, ReturnLen As Long) As Boolean

Public Function GetODBCDrivers() As String
    Dim lngReturn As Long
    
    On Error GoTo APIError
    GetODBCDrivers = Space(1000)
    Call SQLGetInstalledDrivers(GetODBCDrivers, LenB(GetODBCDrivers), lngReturn)
    If lngReturn > 2 Then
        GetODBCDrivers = Mid(GetODBCDrivers, 1, lngReturn - 2)
    Else
        GetODBCDrivers = ""
    End If
    Exit Function
APIError:
    If ErrCenter() = 1 Then Resume
    GetODBCDrivers = ""
    Call SaveErrLog
End Function

Public Function GetODBCSources() As String
    Dim lngReturn As Long
    Dim strSourceName As String, lenNameReturn As Long
    Dim strDescription As String, lenDescReturn As Long
    Dim hEnv As Long
    
    On Error GoTo APIError
    GetODBCSources = "": strSourceName = Space(200): strDescription = Space(500)
    
    Call SQLAllocHandle(SQL_HANDLE_ENV, 0&, hEnv)
    Call SQLSetEnvAttrLong(hEnv, SQL_ATTR_ODBC_VERSION, SQL_OV_ODBC2, 0)
    
    lngReturn = SQLDataSources(hEnv, SQL_FETCH_FIRST_USER, strSourceName, LenB(strSourceName), lenNameReturn, _
        strDescription, LenB(strDescription), lenDescReturn)
    If lngReturn <> SQL_NO_DATA Then
        GetODBCSources = Replace(Mid(strSourceName, 1, lenNameReturn), Chr(0), "") + Chr(0) + Replace(Mid(strDescription, 1, lenDescReturn), Chr(0), "")
        
        Do While True
            lngReturn = SQLDataSources(hEnv, SQL_FETCH_NEXT, strSourceName, LenB(strSourceName), lenNameReturn, _
                strDescription, LenB(strDescription), lenDescReturn)
            If lngReturn = SQL_NO_DATA Then Exit Do
        
            GetODBCSources = GetODBCSources + Chr(0) + Chr(0) + Replace(Mid(strSourceName, 1, lenNameReturn), Chr(0), "") + Chr(0) + Replace(Mid(strDescription, 1, lenDescReturn), Chr(0), "")
        Loop
    End If
    Exit Function
APIError:
    If ErrCenter() = 1 Then Resume
    GetODBCSources = ""
    Call SaveErrLog
End Function

Public Function SetConnect(ByVal ParentHwnd As Long, ByVal DriverName As String, Optional ByVal ConnectString As String = "") As String
    Dim rsTmp As ADODB.Recordset
    Dim hEnv As Long, hConn As Long
    Dim iLenReturn As Integer
    Dim strInConn As String
    
    SetConnect = Space(1000)
    strInConn = IIF(Len(Trim(DriverName)) = 0, ConnectString, "DRIVER={" & DriverName & "};" & ConnectString)
    
    On Error GoTo APIError
    Call SQLAllocHandle(SQL_HANDLE_ENV, 0&, hEnv)
    Call SQLSetEnvAttrLong(hEnv, SQL_ATTR_ODBC_VERSION, SQL_OV_ODBC2, 0)
    Call SQLAllocHandle(SQL_HANDLE_DBC, hEnv, hConn)
    Call SQLDriverConnect(hConn, ParentHwnd, strInConn, LenB(strInConn), SetConnect, LenB(SetConnect), iLenReturn, SQL_DRIVER_COMPLETE)
    If hConn > 0 Then
        Call SQLDisconnect(hConn)
        Call SQLFreeHandle(SQL_HANDLE_DBC, hConn)
    End If
    If hEnv > 0 Then Call SQLFreeHandle(SQL_HANDLE_ENV, hEnv)
    
    If iLenReturn > 0 Then
        SetConnect = Mid(SetConnect, 1, iLenReturn)
    Else
        SetConnect = ""
    End If
    Exit Function

APIError:
    If ErrCenter() = 1 Then Resume
    SetConnect = ""
    Call SaveErrLog
End Function

Public Function CreateDataSource(ByVal ParentHwnd As Long, Optional ByVal DataSource As String = "") As Boolean
    CreateDataSource = SQLCreateDataSource(ParentHwnd, DataSource)
End Function

Public Function ConfigDataSource(ByVal ParentHwnd As Long, ByVal DataSource As String, ByVal DriverName As String) As Boolean
    ConfigDataSource = SQLConfigDataSource(ParentHwnd, ODBC_CONFIG_DSN, DriverName, "DSN=" & DataSource)
'    If Not ConfigDataSource Then ConfigDataSource = SQLConfigDataSource(ParentHwnd, ODBC_CONFIG_SYS_DSN, DriverName, "DSN=" & DataSource)
End Function

Public Function RemoveDataSource(ByVal ParentHwnd As Long, ByVal DataSource As String, ByVal DriverName As String) As Boolean
    RemoveDataSource = SQLConfigDataSource(ParentHwnd, ODBC_REMOVE_DSN, DriverName, "DSN=" & DataSource)
'    If Not RemoveDataSource Then RemoveDataSource = SQLConfigDataSource(ParentHwnd, ODBC_REMOVE_SYS_DSN, DriverName, "DSN=" & DataSource)
End Function

Public Function getTables(ByVal ConnectString As String) As Variant()
    Dim DataConn As New ADODB.Connection, rsTables As ADODB.Recordset
    Dim i As Long, tmpTables() As Variant
    Dim DataEngine As New DAO.DBEngine, DBWork As DAO.Workspace
    Dim DBase As DAO.Database
    
    On Error GoTo DBError
    tmpTables = Array(): getTables = tmpTables
    
    Set DBWork = DataEngine.CreateWorkspace("JetWork", "Admin", "", dbUseJet)
    Set DBase = getDatabase(DBWork, ConnectString)
    
    If DBase.TableDefs.Count < 1 Then Exit Function

    ReDim tmpTables(DBase.TableDefs.Count - 1)
    For i = 0 To DBase.TableDefs.Count - 1
        tmpTables(i) = DBase.TableDefs(i).Name
    Next

'   Process with ADO
'    With DataConn
'        Call .Open(ConnectString)
'        Set rsTables = .OpenSchema(adSchemaTables)
'    End With
'
'    If rsTables.RecordCount < 1 Then Exit Function
'
'    ReDim tmpTables(rsTables.RecordCount - 1)
'    For i = 0 To rsTables.RecordCount - 1
'        tmpTables(i) = IIf(IsNull(rsTables("Table_Schema")) Or Len(Trim(rsTables("Table_Schema"))) = 0, "", rsTables("Table_Schema") & ".") & rsTables("Table_Name")
'
'        rsTables.MoveNext
'    Next
    getTables = tmpTables
    Exit Function
    
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function getDatabase(DBWork As DAO.Workspace, ByVal ConnectString As String) As DAO.Database
    Dim strDBName As String, strConn As String
    
    On Error GoTo DBError
    
    Set getDatabase = Nothing
    Select Case True
        Case UCase(ConnectString) Like "*DBASE*" 'DBase文件
            strDBName = GetConnectAttr(ConnectString, "DefaultDir=")
            If Len(strDBName) = 0 Then
                strDBName = GetConnectAttr(ConnectString, "DBQ=")
                If Len(strDBName) = 0 Then strDBName = App.Path
            End If
            
            strConn = GetConnectAttr(ConnectString, "FIL=")
            If Len(strConn) = 0 Then strConn = "dBase III"
            Set getDatabase = DBWork.OpenDatabase(strDBName, dbDriverComplete, True, strConn)
        Case UCase(ConnectString) Like "*MDB*" Or UCase(ConnectString) Like "*MS ACCESS*" 'Mdb文件
            strDBName = GetConnectAttr(ConnectString, "DBQ=")
            Set getDatabase = DBWork.OpenDatabase(strDBName, dbDriverComplete, True, "")
        Case UCase(ConnectString) Like "*TXT*" Or UCase(ConnectString) Like "*TEXT*" 'Text文件
            strDBName = GetConnectAttr(ConnectString, "DefaultDir=")
            strConn = GetConnectAttr(ConnectString, "FIL=")
            If Len(strConn) = 0 Then strConn = "Text"
            Set getDatabase = DBWork.OpenDatabase(strDBName, dbDriverComplete, True, strConn)
        Case Else
            Set getDatabase = DBWork.OpenDatabase("", dbDriverComplete, True, ConnectString)
    End Select
    Exit Function
    
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetFieldTypeName(ByVal iTypeConst As Integer) As String
    Select Case iTypeConst
        Case dbBigInt
            GetFieldTypeName = "数字" 'Big Integer
        Case dbBinary
            GetFieldTypeName = "二进制" 'Binary
        Case dbBoolean
            GetFieldTypeName = "数字" 'Boolean
        Case dbByte
            GetFieldTypeName = "数字" 'Byte
        Case dbChar
            GetFieldTypeName = "字符" 'Char
        Case dbCurrency
            GetFieldTypeName = "数字" 'Currency
        Case dbDate
            GetFieldTypeName = "日期" 'Date / Time
        Case dbDecimal
            GetFieldTypeName = "数字" 'Decimal
        Case dbDouble
            GetFieldTypeName = "数字" 'Double
        Case dbFloat
            GetFieldTypeName = "数字" 'Float
        Case dbGUID
            GetFieldTypeName = "数字" 'Guid
        Case dbInteger
            GetFieldTypeName = "数字" 'Integer
        Case dbLong
            GetFieldTypeName = "数字" 'Long
        Case dbLongBinary
            GetFieldTypeName = "二进制" 'Long Binary (OLE Object)
        Case dbMemo
            GetFieldTypeName = "字符串" 'Memo
        Case dbNumeric
            GetFieldTypeName = "数字" 'Numeric
        Case dbSingle
            GetFieldTypeName = "数字" 'Single
        Case dbText
            GetFieldTypeName = "字符串" 'Text
        Case dbTime
            GetFieldTypeName = "日期" 'Time
        Case dbTimeStamp
            GetFieldTypeName = "日期" 'Time Stamp
        Case dbVarBinary
            GetFieldTypeName = "二进制" 'VarBinaryEnd
        Case Else
            GetFieldTypeName = "其他"
    End Select
End Function

Private Function GetConnectAttr(ByVal ConnectString As String, ByVal AttrName As String) As String
    Dim iPos1 As Long, iPos2 As Long

    iPos1 = InStr(ConnectString, AttrName)
    If iPos1 = 0 Then
        GetConnectAttr = ""
    Else
        iPos1 = iPos1 + Len(AttrName)
    End If
    If iPos1 > 0 Then
        iPos2 = InStr(iPos1, ConnectString, ";")
        If iPos2 = 0 Then
            GetConnectAttr = Mid(ConnectString, iPos1)
        Else
            GetConnectAttr = Mid(ConnectString, iPos1, iPos2 - iPos1)
        End If
    End If
End Function

